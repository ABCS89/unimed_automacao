"""
cruzamento.py - Cruzamento automático inteligente entre bases de dados:
- Base mensal de servidores (teste.ods)
- Devedores e inadimplência (devedores.xlsx - Inadimplentes e Cancelados)
- Regra de Inadimplência:
    * 1 registro de boleto em aberto: Normal (continua ativo e cobrança regular)
    * 2 registros de boleto em aberto: Aviso de cancelamento (notificação com débitos)
    * 3 ou mais registros de boleto em aberto: Cancelamento (notificação de rescisão com débitos)
- Cancelados formalizados (planilha avulsa ou aba Cancelados de devedores.xlsx)
- Desligados do mês (planilha avulsa desligados.xlsx ou aba Desligados de devedores.xlsx)
"""
from pathlib import Path
import re
import pandas as pd

from .config import ENTRADA_DIR, ARQUIVO_DEVEDORES, ARQUIVO_DESLIGADOS
from .utils import limpa, normalizar_nome


def _padronizar_funcional(val):
    """Converte valores numéricos ou textuais em string limpa de matrícula."""
    if pd.isna(val) or val is None:
        return ""
    try:
        return str(int(float(val)))
    except (ValueError, TypeError):
        return str(val).strip()


def _extrair_funcionais_e_nomes(df):
    """Detecta dinamicamente colunas de funcional e nome em um DataFrame."""
    col_func = None
    col_nome = None
    for col in df.columns:
        c_low = str(col).strip().lower()
        if any(k in c_low for k in ["funcional", "matricula", "matrícula", "registro", "nro", "func"]):
            col_func = col
        elif any(k in c_low for k in ["nome", "funcion", "servidor"]):
            col_nome = col

    funcs = set()
    nomes = set()
    if col_func:
        for val in df[col_func].dropna():
            f_str = _padronizar_funcional(val)
            if f_str:
                funcs.add(f_str)
    if col_nome:
        for val in df[col_nome].dropna():
            n_norm = normalizar_nome(val)
            if n_norm:
                nomes.add(n_norm)
    return funcs, nomes


def carregar_lista_desligados(pasta_entrada=ENTRADA_DIR, arquivo_devedores=ARQUIVO_DEVEDORES):
    """
    Carrega servidores desligados procurando primeiro por arquivo avulso em entrada/
    (ex: desligados.xlsx, desligados.ods, desligados.csv) e, como fallback,
    na aba 'Desligados' de devedores.xlsx.
    """
    extensoes = [".xlsx", ".ods", ".xls", ".csv"]
    arquivo_encontrado = None

    for ext in extensoes:
        candidato = pasta_entrada / f"desligados{ext}"
        if candidato.exists():
            arquivo_encontrado = candidato
            break

    if not arquivo_encontrado and pasta_entrada.exists():
        for f in pasta_entrada.iterdir():
            if f.is_file() and "desligad" in f.stem.lower() and f.suffix.lower() in extensoes:
                arquivo_encontrado = f
                break

    if arquivo_encontrado:
        try:
            if arquivo_encontrado.suffix.lower() == ".csv":
                df_desl = pd.read_csv(arquivo_encontrado)
            elif arquivo_encontrado.suffix.lower() == ".ods":
                df_desl = pd.read_excel(arquivo_encontrado, engine="odf")
            else:
                df_desl = pd.read_excel(arquivo_encontrado)

            funcs, nomes = _extrair_funcionais_e_nomes(df_desl)
            qtd = len(funcs or nomes)
            origem = f"planilha avulsa '{arquivo_encontrado.name}' ({qtd} registros)"
            print(f"[INFO] Servidores desligados carregados via {origem}")
            return funcs, nomes, origem
        except Exception as e:
            print(f"  ⚠️ Aviso ao ler '{arquivo_encontrado.name}': {e}. Tentando alternativa...")

    if arquivo_devedores.exists():
        try:
            excel_dev = pd.ExcelFile(arquivo_devedores)
            if "Desligados" in excel_dev.sheet_names:
                df_desl = pd.read_excel(arquivo_devedores, sheet_name="Desligados")
                funcs, nomes = _extrair_funcionais_e_nomes(df_desl)
                qtd = len(funcs or nomes)
                origem = f"aba 'Desligados' de {arquivo_devedores.name} ({qtd} registros)"
                print(f"[INFO] Servidores desligados carregados via {origem}")
                return funcs, nomes, origem
        except Exception as e:
            print(f"  ⚠️ Aviso ao carregar aba Desligados de {arquivo_devedores.name}: {e}")

    origem = "Nenhuma base de desligados encontrada"
    return set(), set(), origem


def carregar_lista_cancelados(pasta_entrada=ENTRADA_DIR, arquivo_devedores=ARQUIVO_DEVEDORES):
    """
    Carrega cancelados formalizados (próprio pedido / ofício) da planilha avulsa de cancelamentos
    e/ou da aba 'Cancelados' de devedores.xlsx.
    """
    extensoes = [".xlsx", ".ods", ".xls", ".csv"]
    funcs_canc = set()
    nomes_canc = set()
    df_canc_total = pd.DataFrame()

    # 1. Procurar por arquivo avulso de cancelados em entrada/ (ex: cancelados.xlsx, cancelamentos.xlsx)
    if pasta_entrada.exists():
        for f in pasta_entrada.iterdir():
            if f.is_file() and any(k in f.stem.lower() for k in ["cancelad", "cancelament"]) and f.suffix.lower() in extensoes:
                try:
                    df = pd.read_csv(f) if f.suffix.lower() == ".csv" else pd.read_excel(f)
                    fn, nm = _extrair_funcionais_e_nomes(df)
                    funcs_canc.update(fn)
                    nomes_canc.update(nm)
                    df_canc_total = pd.concat([df_canc_total, df], ignore_index=True)
                    print(f"[INFO] Cancelamentos mensais carregados via planilha avulsa '{f.name}' ({len(fn or nm)} registros)")
                except Exception as e:
                    print(f"  ⚠️ Aviso ao ler '{f.name}': {e}")

    # 2. Ler aba 'Cancelados' de devedores.xlsx
    if arquivo_devedores.exists():
        try:
            excel_dev = pd.ExcelFile(arquivo_devedores)
            if "Cancelados" in excel_dev.sheet_names:
                df_dev_canc = pd.read_excel(arquivo_devedores, sheet_name="Cancelados")
                if "Funcional" in df_dev_canc.columns:
                    df_dev_canc["Funcional"] = df_dev_canc["Funcional"].apply(_padronizar_funcional)
                fn, nm = _extrair_funcionais_e_nomes(df_dev_canc)
                funcs_canc.update(fn)
                nomes_canc.update(nm)
                df_canc_total = pd.concat([df_canc_total, df_dev_canc], ignore_index=True)
                print(f"[INFO] Devedores com plano já cancelado: {len(fn)} matrículas (aba Cancelados)")
        except Exception as e:
            print(f"  ⚠️ Aviso ao ler aba Cancelados de {arquivo_devedores.name}: {e}")

    return funcs_canc, nomes_canc, df_canc_total


def carregar_bases_cruzamento(arquivo_base, arquivo_devedores, pasta_entrada=ENTRADA_DIR):
    """
    Carrega todas as bases e calcula as contagens de boletos por devedor:
    - Desligados
    - Cancelados prévios (ofício ou pedido)
    - Inadimplentes divididos por quantidade de boletos:
        * 1 boleto: Normal
        * 2 boletos: Aviso de cancelamento
        * 3+ boletos: Cancelamento por inadimplência
    """
    funcs_desligados, nomes_desligados, origem_desl = carregar_lista_desligados(pasta_entrada, arquivo_devedores)
    funcs_cancelados_previos, nomes_cancelados_previos, df_cancelados = carregar_lista_cancelados(pasta_entrada, arquivo_devedores)

    df_dividas = pd.DataFrame()
    contagem_boletos_func = {}
    contagem_boletos_nome = {}

    if arquivo_devedores.exists():
        try:
            excel_dev = pd.ExcelFile(arquivo_devedores)
            if "Inadimplentes" in excel_dev.sheet_names:
                df_dividas = pd.read_excel(arquivo_devedores, sheet_name="Inadimplentes")
                df_dividas["Funcional"] = df_dividas["Funcional"].apply(_padronizar_funcional)

                # Mapeia contagem de boletos por funcional
                for f_val in df_dividas["Funcional"].dropna():
                    if f_val:
                        contagem_boletos_func[f_val] = contagem_boletos_func.get(f_val, 0) + 1

                # Mapeia contagem de boletos por nome normalizado
                if "Nome" in df_dividas.columns:
                    for n_val in df_dividas["Nome"].dropna():
                        n_norm = normalizar_nome(n_val)
                        if n_norm:
                            contagem_boletos_nome[n_norm] = contagem_boletos_nome.get(n_norm, 0) + 1

                qtd_1 = sum(1 for v in contagem_boletos_func.values() if v == 1)
                qtd_2 = sum(1 for v in contagem_boletos_func.values() if v == 2)
                qtd_3p = sum(1 for v in contagem_boletos_func.values() if v >= 3)
                print(f"[INFO] Inadimplentes mapeados: {len(contagem_boletos_func)} servidores na aba Inadimplentes")
                print(f"       * {qtd_1} com 1 boleto -> Cobrança regular (Normal)")
                print(f"       * {qtd_2} com 2 boletos -> Aviso de Cancelamento")
                print(f"       * {qtd_3p} com 3+ boletos -> Cancelamento do Benefício")
        except Exception as e:
            print(f"  ⚠️ Erro ao carregar aba Inadimplentes de {arquivo_devedores}: {e}")

    return {
        "funcs_desligados": funcs_desligados,
        "nomes_desligados": nomes_desligados,
        "origem_desligados": origem_desl,
        "funcs_cancelados": funcs_cancelados_previos,
        "nomes_cancelados": nomes_cancelados_previos,
        "df_cancelados": df_cancelados,
        "df_dividas": df_dividas,
        "contagem_boletos_func": contagem_boletos_func,
        "contagem_boletos_nome": contagem_boletos_nome,
    }


def classificar_servidor(linha, contexto_cruzamento):
    """
    Classifica o servidor conforme a regra solicitada:
    1. 'não enviar' na coluna condição de teste.ods -> ignora
    2. Constar nos Desligados -> 'desligado'
    3. Constar nos Cancelados prévios (aba Cancelados ou planilha) -> 'cancelado'
    4. Constar em Inadimplentes:
       - 3 ou mais boletos -> 'cancelado' (cancelamento por inadimplência)
       - 2 boletos -> 'aviso' (aviso de cancelamento)
       - 1 boleto -> 'normal' (cobrança mensal regular)
    5. Sobrescrita manual em teste.ods ('desligado', 'aviso', 'cancelado') se indicada
    6. Caso contrário -> 'normal'
    """
    condicao_manual = limpa(linha.get("condição")).lower()

    if "não enviar" in condicao_manual or "nao enviar" in condicao_manual:
        return "ignorado"

    matricula = _padronizar_funcional(linha.get("Nro Funcional"))
    nome_norm = normalizar_nome(linha.get("Funcionário", ""))

    # 1. Servidor Desligado
    if (matricula and matricula in contexto_cruzamento["funcs_desligados"]) or (
        nome_norm and nome_norm in contexto_cruzamento["nomes_desligados"]
    ):
        return "desligado"

    # 2. Servidor Cancelado previamente (a pedido ou por ofício)
    if (matricula and matricula in contexto_cruzamento["funcs_cancelados"]) or (
        nome_norm and nome_norm in contexto_cruzamento["nomes_cancelados"]
    ):
        return "cancelado"

    # 3. Regra de Boletos em Aberto na aba Inadimplentes
    qtd_boletos = (
        contexto_cruzamento["contagem_boletos_func"].get(matricula)
        or contexto_cruzamento["contagem_boletos_nome"].get(nome_norm)
        or 0
    )

    if qtd_boletos >= 3:
        return "cancelado"
    elif qtd_boletos == 2:
        return "aviso"
    elif qtd_boletos == 1:
        return "normal"

    # 4. Sobrescrita manual caso tenha sido informada explicitamente na base
    if condicao_manual in ("desligado", "aviso", "cancelado"):
        return condicao_manual

    return "normal"


def obter_debitos_servidor(linha, contexto_cruzamento, tipo):
    """
    Recupera as parcelas em atraso do servidor:
    - Se for 'aviso': busca na aba Inadimplentes.
    - Se for 'cancelado': busca na aba Inadimplentes (caso seja cancelamento por 3+ boletos)
      ou na base de Cancelados.
    """
    matricula = _padronizar_funcional(linha.get("Nro Funcional"))
    nome_norm = normalizar_nome(linha.get("Funcionário", ""))

    def _filtrar(df):
        if df is None or df.empty:
            return pd.DataFrame()
        cond_f = df["Funcional"].astype(str) == matricula if matricula and "Funcional" in df.columns else False
        cond_n = (
            df["Nome"].apply(normalizar_nome) == nome_norm
            if "Nome" in df.columns and nome_norm
            else False
        )
        return df[cond_f | cond_n]

    df_inad = contexto_cruzamento.get("df_dividas", pd.DataFrame())
    df_canc = contexto_cruzamento.get("df_cancelados", pd.DataFrame())

    if tipo == "aviso":
        return _filtrar(df_inad)
    elif tipo == "cancelado":
        res_inad = _filtrar(df_inad)
        if not res_inad.empty:
            return res_inad
        return _filtrar(df_canc)

    return pd.DataFrame()
