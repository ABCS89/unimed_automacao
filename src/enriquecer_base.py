"""
enriquecer_base.py - Automação de enriquecimento cadastral do mês corrente
Recupera histórico dos meses anteriores em teste.ods para preencher automaticamente
endereço, CPF, data de nascimento (formatada dd/mm/aaaa), CEP (formatado),
e-mail e condição na nova competência.
"""
import re
import shutil
from pathlib import Path
import pandas as pd

from .config import ARQUIVO_BASE
from .utils import normalizar_nome, limpa

MESES_ORDEM = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

COLUNAS_PADRAO = [
    "Nro Funcional",
    "Funcionário",
    "cpf",
    "data_nascimento",
    "Mensalidade",
    "Coparticipação",
    "Total",
    "condição",
    "mail",
    "CEP",
    "endereço",
    "bairro",
    "complemento",
    "cidade",
    "uf"
]


def _limpar_funcional(val):
    if pd.isna(val) or val is None:
        return ""
    try:
        return str(int(float(val)))
    except (ValueError, TypeError):
        return str(val).strip()


def _formatar_data_br(val):
    """Garante que a data seja gravada como texto legível no formato DD/MM/AAAA."""
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ("nat", "nan", "none"):
        return ""
    try:
        dt = pd.to_datetime(val, dayfirst=True)
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return val_str


def _formatar_cep(val):
    """Garante a formatação padrão do CEP como XXXXX-XXX."""
    if pd.isna(val) or val is None:
        return ""
    c = str(val).strip()
    if not c or c.lower() in ("nan", "none"):
        return ""
    digitos = re.sub(r"\D", "", c)
    if len(digitos) == 8:
        return f"{digitos[:5]}-{digitos[5:]}"
    return c


def encontrar_ultimo_mes_preenchido(excel_file):
    """
    Identifica qual é o último mês com registros brutos ou necessitando de enriquecimento.
    Prioriza o mês mais recente que possua dados.
    """
    for mes in reversed(MESES_ORDEM):
        if mes in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=mes)
            if not df.empty and len(df) > 0:
                return mes
    return None


def coletar_historico_cadastral(excel_file, mes_alvo):
    """
    Varre os meses anteriores ao mes_alvo (do mais recente para o mais antigo)
    e monta um dicionário de cadastros consolidados por funcional e por nome normalizado.
    """
    try:
        idx_alvo = MESES_ORDEM.index(mes_alvo)
        meses_anteriores = [MESES_ORDEM[i] for i in range(idx_alvo - 1, -1, -1)]
    except ValueError:
        meses_anteriores = [m for m in reversed(MESES_ORDEM) if m != mes_alvo]

    historico_por_func = {}
    historico_por_nome = {}
    origem_dados = {}

    for mes in meses_anteriores:
        if mes not in excel_file.sheet_names:
            continue
        df_mes = pd.read_excel(excel_file, sheet_name=mes)
        if df_mes.empty:
            continue

        # Mapeamento semântico insensível a maiúsculas/minúsculas
        mapa_cols = {}
        for col in df_mes.columns:
            c_low = col.strip().lower()
            if "funcional" in c_low:
                mapa_cols[col] = "nro_funcional"
            elif "nome" in c_low or "funcion" in c_low:
                mapa_cols[col] = "funcionário"
            elif "cpf" in c_low:
                mapa_cols[col] = "cpf"
            elif "nasc" in c_low:
                mapa_cols[col] = "data_nascimento"
            elif "condiç" in c_low or "condic" in c_low:
                mapa_cols[col] = "condição"
            elif "mail" in c_low or "email" in c_low or "e-mail" in c_low:
                mapa_cols[col] = "mail"
            elif "cep" in c_low:
                mapa_cols[col] = "cep"
            elif "endereço" in c_low or "endereco" in c_low:
                mapa_cols[col] = "endereço"
            elif "bairro" in c_low:
                mapa_cols[col] = "bairro"
            elif "complemento" in c_low:
                mapa_cols[col] = "complemento"
            elif "cidade" in c_low:
                mapa_cols[col] = "cidade"
            elif "uf" in c_low or "estado" in c_low:
                mapa_cols[col] = "uf"

        if "nro_funcional" not in mapa_cols.values() and "funcionário" not in mapa_cols.values():
            continue

        df_mes_norm = df_mes.rename(columns=mapa_cols)

        for _, row in df_mes_norm.iterrows():
            func = _limpar_funcional(row.get("nro_funcional"))
            nome = limpa(row.get("funcionário"))
            nome_norm = normalizar_nome(nome)

            cad = {
                "cpf": limpa(row.get("cpf")),
                "data_nascimento": _formatar_data_br(row.get("data_nascimento")),
                "condição": limpa(row.get("condição")),
                "mail": limpa(row.get("mail")),
                "CEP": _formatar_cep(row.get("cep")),
                "endereço": limpa(row.get("endereço")),
                "bairro": limpa(row.get("bairro")),
                "complemento": limpa(row.get("complemento")),
                "cidade": limpa(row.get("cidade")),
                "uf": limpa(row.get("uf")),
            }

            # Considera válido se possui dados essenciais (endereço, CPF, CEP ou e-mail)
            tem_dados = any([cad["endereço"], cad["cpf"], cad["CEP"], cad["mail"]])
            if tem_dados:
                if func and func not in historico_por_func:
                    historico_por_func[func] = cad
                    origem_dados[func] = mes
                if nome_norm and nome_norm not in historico_por_nome:
                    historico_por_nome[nome_norm] = cad
                    if nome_norm not in origem_dados:
                        origem_dados[nome_norm] = mes

    return historico_por_func, historico_por_nome, origem_dados


def enriquecer_base_mes(caminho_base=None, mes_alvo=None, fazer_backup=True):
    """
    Enriquece os registros brutos do mês alvo com base no histórico dos meses anteriores.
    Atualiza tanto a aba do mês respectivo quanto a 'Planilha1'.
    """
    if caminho_base is None:
        caminho_base = ARQUIVO_BASE
    caminho_base = Path(caminho_base)

    if not caminho_base.exists():
        print(f"[ERRO] Arquivo não encontrado: {caminho_base}")
        return False

    print("\n" + "=" * 60)
    print(">>> ENRIQUECIMENTO AUTOMATICO DE REGISTROS (TESTE.ODS)")
    print("=" * 60)

    with pd.ExcelFile(caminho_base, engine="odf") as excel:
        todas_abas = {s: pd.read_excel(excel, sheet_name=s) for s in excel.sheet_names}

        if mes_alvo is None:
            mes_alvo = encontrar_ultimo_mes_preenchido(excel)

        if not mes_alvo or mes_alvo not in todas_abas:
            print(f"[ERRO] Mês alvo '{mes_alvo}' não encontrado nas abas do arquivo!")
            return False

        print(f"[INFO] Mês alvo para enriquecimento: {mes_alvo}")
        df_alvo = todas_abas[mes_alvo]

        # Se a aba do mês estiver vazia mas Planilha1 tiver dados, usar Planilha1
        if df_alvo.empty and "Planilha1" in todas_abas and not todas_abas["Planilha1"].empty:
            print(f"[INFO] A aba {mes_alvo} estava vazia. Usando os registros brutos de Planilha1.")
            df_alvo = todas_abas["Planilha1"]

        if df_alvo.empty:
            print(f"[ERRO] Nenhum registro encontrado para enriquecer no mês {mes_alvo}!")
            return False

        print(f"[INFO] Registros brutos encontrados no mês {mes_alvo}: {len(df_alvo)}")

        # Coletar histórico dos meses anteriores
        hist_func, hist_nome, origens = coletar_historico_cadastral(excel, mes_alvo)

    print(f"[INFO] Base histórica acumulada: {len(hist_func)} servidores mapeados dos meses anteriores.")

    # Cruzar dados
    linhas_enriquecidas = []
    recuperados = 0
    novos = []

    for _, row in df_alvo.iterrows():
        func_orig = row.get("Nro Funcional")
        func_str = _limpar_funcional(func_orig)
        nome = limpa(row.get("Funcionário") or row.get("funcionário"))
        nome_norm = normalizar_nome(nome)

        cad = hist_func.get(func_str) or hist_nome.get(nome_norm) or {}

        if cad:
            recuperados += 1
            mes_origem = origens.get(func_str) or origens.get(nome_norm) or "histórico"
        else:
            novos.append((func_str, nome))
            mes_origem = None

        # Montar linha nas 15 colunas padrão
        nova_linha = {
            "Nro Funcional": int(func_orig) if pd.notna(func_orig) and str(func_orig).isdigit() else func_orig,
            "Funcionário": nome,
            "cpf": cad.get("cpf", ""),
            "data_nascimento": cad.get("data_nascimento", ""),
            "Mensalidade": float(row.get("Mensalidade", 0.0) or 0.0),
            "Coparticipação": float(row.get("Coparticipação", 0.0) or 0.0),
            "Total": float(row.get("Total", 0.0) or 0.0),
            "condição": cad.get("condição", ""),
            "mail": cad.get("mail", ""),
            "CEP": cad.get("CEP", ""),
            "endereço": cad.get("endereço", ""),
            "bairro": cad.get("bairro", ""),
            "complemento": cad.get("complemento", ""),
            "cidade": cad.get("cidade", ""),
            "uf": cad.get("uf", ""),
        }
        linhas_enriquecidas.append(nova_linha)

    df_enriquecido = pd.DataFrame(linhas_enriquecidas, columns=COLUNAS_PADRAO)

    # Formatar também datas e CEPs nas abas históricas para garantir consistência visual no LibreOffice
    for nome_aba, df_aba in todas_abas.items():
        if df_aba is not None and not df_aba.empty:
            for c in df_aba.columns:
                c_low = c.strip().lower()
                if "nasc" in c_low:
                    df_aba[c] = df_aba[c].apply(_formatar_data_br)
                elif "cep" in c_low:
                    df_aba[c] = df_aba[c].apply(_formatar_cep)

    # Backup de segurança antes de gravar
    if fazer_backup:
        caminho_backup = caminho_base.parent / f"{caminho_base.stem}_backup{caminho_base.suffix}"
        shutil.copyfile(caminho_base, caminho_backup)
        print(f"[BACKUP] Cópia de segurança salva em: {caminho_backup.name}")

    # Atualiza as abas no dicionário
    todas_abas[mes_alvo] = df_enriquecido
    todas_abas["Planilha1"] = df_enriquecido

    # Gravar de volta no arquivo ODS preservando todas as abas
    with pd.ExcelWriter(caminho_base, engine="odf") as writer:
        for nome_aba, df_aba in todas_abas.items():
            if df_aba.empty and len(df_aba.columns) == 0:
                pd.DataFrame().to_excel(writer, sheet_name=nome_aba, index=False)
            else:
                df_aba.to_excel(writer, sheet_name=nome_aba, index=False)

    print("\n" + "=" * 60)
    print(f"[SUCESSO] Base do mês {mes_alvo} enriquecida com sucesso!")
    print(f"  * Total de registros no mês: {len(df_enriquecido)}")
    print(f"  * Registros recuperados do histórico: {recuperados} ({round((recuperados/len(df_enriquecido))*100, 1)}%)")
    print(f"  * Registros novos (sem histórico prévio): {len(novos)}")
    print(f"  * Abas atualizadas no ODS: '{mes_alvo}' e 'Planilha1'")

    if novos:
        print("\n[ATENÇÃO] Os seguintes servidores são novos e precisam de cadastro cadastral manual:")
        for func_n, nome_n in novos:
            print(f"   - Funcional: {func_n:<8} | Nome: {nome_n}")
    print("=" * 60 + "\n")

    return True
