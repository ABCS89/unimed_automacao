"""
refis.py - Geração de termos de negociação Refis para servidores
"""
from datetime import datetime
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate

from .config import (
    ARQUIVO_BASE,
    ARQUIVO_DEVEDORES,
    TEMPLATE_REFIS,
    REFIS_DIR,
    MESES_PT,
)
from .utils import (
    limpa,
    capitalizar_nome,
    formatar_valor_br,
    valor_por_extenso,
    limpar_nome_arquivo,
    normalizar_nome,
)


def gerar_todos_refis():
    """Gera os termos do Refis em saida/refis/ com cruzamento de dados cadastrais e valores reais."""
    print("\n" + "=" * 60)
    print(">>> INICIANDO GERACAO DE TERMOS DO REFIS")
    print("=" * 60)

    REFIS_DIR.mkdir(parents=True, exist_ok=True)

    if not ARQUIVO_DEVEDORES.exists() or not TEMPLATE_REFIS.exists():
        print("[ERRO] Arquivo devedores.xlsx ou template_refis.docx ausente.")
        return

    # Base cadastral para buscar endereço, CEP e cidade
    df_base = pd.read_excel(ARQUIVO_BASE, engine="odf") if ARQUIVO_BASE.exists() else pd.DataFrame()
    if not df_base.empty:
        df_base["Nro Funcional"] = pd.to_numeric(df_base["Nro Funcional"], errors="coerce").astype("Int64").astype(str)

    # Verificar aba de devedores
    excel_file = pd.ExcelFile(ARQUIVO_DEVEDORES)
    aba_alvo = "Desligados" if "Desligados" in excel_file.sheet_names else excel_file.sheet_names[0]

    df_desligados = pd.read_excel(ARQUIVO_DEVEDORES, sheet_name=aba_alvo)
    df_desligados["Funcional"] = pd.to_numeric(df_desligados["Funcional"], errors="coerce").astype("Int64").astype(str)

    hoje = datetime.now()
    gerados = 0

    for _, row in df_desligados.iterrows():
        nome = limpa(row.get("Nome") or row.get("Funcionário"))
        if not nome:
            continue

        matricula = limpa(row.get("Funcional"))
        nome_cap = capitalizar_nome(nome)
        nome_upper = nome.upper()
        nome_arquivo = limpar_nome_arquivo(nome_cap)

        # Buscar dados de endereço em teste.ods
        linha_cadastral = None
        if not df_base.empty:
            match_func = df_base[df_base["Nro Funcional"] == matricula]
            if not match_func.empty:
                linha_cadastral = match_func.iloc[0]
            else:
                nome_norm = normalizar_nome(nome)
                for _, b_row in df_base.iterrows():
                    if normalizar_nome(b_row.get("Funcionário", "")) == nome_norm:
                        linha_cadastral = b_row
                        break

        if linha_cadastral is not None:
            endereco_rua = limpa(linha_cadastral.get("endereço"))
            complemento = limpa(linha_cadastral.get("complemento"))
            bairro = limpa(linha_cadastral.get("bairro"))
            if complemento:
                endereco_rua += f" – {complemento}"
            if bairro:
                endereco_rua += f" – {bairro}"
            cep = limpa(linha_cadastral.get("CEP")) or "13400-000"
            cidade = limpa(linha_cadastral.get("cidade")) or "PIRACICABA"
            endereco_rua = capitalizar_nome(endereco_rua)
        else:
            endereco_rua = ""
            cep = "13400-000"
            cidade = "PIRACICABA"

        valor_num = float(row.get("Saldo (Atualizado)", 0) or 0)
        valor_formatado = formatar_valor_br(valor_num)
        valor_ext = valor_por_extenso(valor_num)
        referencia = limpa(row.get("Mês/Ano", ""))

        data_limite_val = row.get("Data de Vencimento", "")
        if pd.notna(data_limite_val):
            try:
                data_limite = pd.to_datetime(data_limite_val).strftime("%d/%m/%Y")
            except Exception:
                data_limite = str(data_limite_val)
        else:
            data_limite = ""

        contexto = {
            "nome": nome_cap,
            "nome_cap": nome_cap,
            "nome_upper": nome_upper,
            "valor": valor_formatado,
            "valor_extenso": valor_ext,
            "referencia": referencia,
            "data_limite": data_limite,
            "endereco": endereco_rua,
            "linha_endereco": endereco_rua,
            "CEP": cep,
            "cidade": cidade,
            "dia": str(hoje.day),
            "mes": MESES_PT[hoje.month],
            "ano": str(hoje.year),
        }

        doc = DocxTemplate(TEMPLATE_REFIS)
        doc.render(contexto)

        # Tratar tags residuais em colchetes caso presentes no template
        substituicoes_colchetes = {
            "[CEP do servidor]": cep,
            "[cidade]": cidade,
            "[dia atual]": str(hoje.day),
            "[mês atual]": MESES_PT[hoje.month],
            "[ano atual]": str(hoje.year),
            "[valor numérico]": valor_formatado,
            "[valor por extenso]": valor_ext,
        }
        for p in doc.paragraphs:
            for k, v in substituicoes_colchetes.items():
                if k in p.text:
                    for r in p.runs:
                        if k in r.text:
                            r.text = r.text.replace(k, str(v))
                    if k in p.text:
                        p.text = p.text.replace(k, str(v))

        caminho_saida = REFIS_DIR / f"{nome_arquivo}.docx"
        doc.save(caminho_saida)
        gerados += 1

    print(f"[OK] Total de termos do Refis gerados: {gerados} (em {REFIS_DIR})")
