"""
cartas.py - Geração de cartas mensais (Base, Desligados, Avisos, Cancelados) e Multas
"""
from datetime import datetime
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate

from .config import (
    ARQUIVO_BASE,
    ARQUIVO_DEVEDORES,
    PASTA_PDFS_ENVIO,
    TEMPLATE_BASE,
    TEMPLATE_DESLIGADO,
    TEMPLATE_AVISO,
    TEMPLATE_CANCELADO,
    TEMPLATE_MULTA,
    CARTAS_BASE_DIR,
    CARTAS_CANCELADOS_DIR,
    CARTAS_MULTA_DIR,
    MESES_PT,
)
from .utils import (
    limpa,
    capitalizar_nome,
    formatar_valor_br,
    valor_por_extenso,
    ultimo_dia_util_do_mes,
    formata_competencia,
    extrair_texto_pdf,
    extrair_data_email_do_pdf,
    mapear_pdfs_por_funcional,
    mes_anterior,
    limpar_nome_arquivo,
)


def _preparar_contextos_data(hoje=None):
    """Monta contextos de data atual e vencimento compatíveis com todos os templates."""
    if hoje is None:
        hoje = datetime.today()

    data_venc = ultimo_dia_util_do_mes(hoje.year, hoje.month)

    return {
        "dia": str(hoje.day),
        "mes": MESES_PT[hoje.month],
        "ano": str(hoje.year),
        "dia_atual": str(hoje.day),
        "mes_atual": MESES_PT[hoje.month],
        "ano_atual": str(hoje.year),
        "ultimo_dia_util": str(data_venc.day),
        "ultimo_dia_do_mes": str(data_venc.day),
        "dia_limite": str(data_venc.day),
        "dia_vencimento": str(data_venc.day),
        "mes_vencimento": MESES_PT[data_venc.month],
        "ano_vencimento": str(data_venc.year),
    }


def _substituir_colchetes_remanescentes(doc, mapa_substituicoes):
    """Garante a substituição de tags em colchetes caso o template ainda as possua."""
    for p in doc.paragraphs:
        for k, v in mapa_substituicoes.items():
            if k in p.text:
                for r in p.runs:
                    if k in r.text:
                        r.text = r.text.replace(k, str(v))
                if k in p.text:
                    p.text = p.text.replace(k, str(v))


def gerar_carta_base_ou_desligado(linha, template_path, contexto_datas, pdf_map, pasta_saida):
    nro_funcional = linha["Nro Funcional"]
    funcionario_raw = linha["Funcionário"]

    if nro_funcional in pdf_map:
        caminho_pdf = PASTA_PDFS_ENVIO / pdf_map[nro_funcional]
        texto_pdf = extrair_texto_pdf(caminho_pdf)
        contexto_email = extrair_data_email_do_pdf(texto_pdf)
    else:
        contexto_email = {"dia_email": "dia", "mes_email": "mês", "ano_email": "ano"}

    endereco_completo = limpa(linha.get("endereço"))
    complemento = limpa(linha.get("complemento"))
    bairro = limpa(linha.get("bairro"))
    if complemento:
        endereco_completo += f" – {complemento}"
    if bairro:
        endereco_completo += f" – {bairro}"

    valor_total = float(linha.get("Total", 0) or 0)
    valor_formatado = formatar_valor_br(valor_total)
    valor_ext = valor_por_extenso(valor_total)

    contexto = {
        "nome_cap": capitalizar_nome(funcionario_raw),
        "nome_upper": str(funcionario_raw).upper(),
        "linha_endereco": endereco_completo,
        "CEP": limpa(linha.get("CEP")),
        "cidade": limpa(linha.get("cidade")),
        "valor": valor_formatado,
        "valor_extenso": valor_ext,
        "email": limpa(linha.get("mail")) or "mail",
        **contexto_datas,
        **contexto_email,
    }

    doc = DocxTemplate(template_path)
    doc.render(contexto)

    # Suporte retrocompatível para templates com colchetes [campo]
    mapa_colchetes = {
        "[dia atual]": contexto["dia"],
        "[mês atual]": contexto["mes"],
        "[ano atual]": contexto["ano"],
        "[ultimo dia do mês atual]": contexto["ultimo_dia_util"],
        "[nome do servidor cap]": contexto["nome_cap"],
        "[nome do servidor upper]": contexto["nome_upper"],
        "[endereço do servidor]": contexto["linha_endereco"],
        "[CEP do servidor]": contexto["CEP"],
        "[cidade]": contexto["cidade"],
        "[valor numérico]": contexto["valor"],
        "[valor por extenso]": contexto["valor_extenso"],
        "[r-mail]": contexto["email"],
    }
    _substituir_colchetes_remanescentes(doc, mapa_colchetes)

    nome_seguro = limpar_nome_arquivo(funcionario_raw)
    doc.save(pasta_saida / f"{nome_seguro}.docx")


def gerar_carta_aviso_ou_cancelado(linha, condicao, template_path, df_dividas, contexto_datas, pasta_saida):
    nome = limpa(linha.get("Funcionário"))
    matricula = limpa(linha.get("Nro Funcional"))

    df_pessoa = df_dividas[df_dividas["Funcional"] == matricula]
    if df_pessoa.empty:
        print(f"  [AVISO] Sem dívidas cadastradas para: {nome} ({condicao}) - carta não gerada")
        return False

    linha_endereco = " – ".join(filter(None, [
        limpa(linha.get("endereço")),
        limpa(linha.get("bairro")),
        limpa(linha.get("complemento")),
    ]))

    tabela = []
    for _, divida in df_pessoa.iterrows():
        principal = float(divida.get("Principal (Saldo)", 0) or 0)
        total = float(divida.get("Saldo (Atualizado)", 0) or 0)
        encargos = total - principal

        data_vencimento = pd.to_datetime(
            divida.get("Data de Vencimento"), dayfirst=True, errors="coerce"
        )

        tabela.append({
            "competencia": formata_competencia(divida.get("Mês/Ano")),
            "vencimento": data_vencimento.strftime("%d/%m/%Y") if pd.notna(data_vencimento) else "",
            "principal": f"R$ {formatar_valor_br(principal)}",
            "encargos": f"R$ {formatar_valor_br(encargos)}",
            "total": f"R$ {formatar_valor_br(total)}",
        })

    contexto = {
        "nome_cap": capitalizar_nome(nome),
        "nome_upper": nome.upper(),
        "linha_endereco": linha_endereco,
        "CEP": limpa(linha.get("CEP")),
        "cidade": limpa(linha.get("cidade")),
        "uf": limpa(linha.get("uf")) or "SP",
        "tabela": tabela,
        **contexto_datas,
    }

    doc = DocxTemplate(template_path)
    doc.render(contexto)

    sufixo = "aviso" if condicao == "aviso" else "cancelado"
    nome_seguro = limpar_nome_arquivo(f"{nome} - {sufixo}")
    doc.save(pasta_saida / f"{nome_seguro}.docx")
    return True


def gerar_todas_cartas_mensais():
    """Gera todas as cartas mensais a partir de teste.ods e devedores.xlsx."""
    print("\n" + "=" * 60)
    print(">>> INICIANDO GERACAO DE CARTAS MENSAIS")
    print("=" * 60)

    CARTAS_BASE_DIR.mkdir(parents=True, exist_ok=True)
    CARTAS_CANCELADOS_DIR.mkdir(parents=True, exist_ok=True)

    if not ARQUIVO_BASE.exists():
        print(f"[ERRO] Arquivo base não encontrado: {ARQUIVO_BASE}")
        return

    if not ARQUIVO_DEVEDORES.exists():
        print(f"[ERRO] Arquivo de devedores não encontrado: {ARQUIVO_DEVEDORES}")
        return

    df_base = pd.read_excel(ARQUIVO_BASE, engine="odf")
    df_dividas = pd.read_excel(ARQUIVO_DEVEDORES, sheet_name="Inadimplentes")
    df_cancelados = pd.read_excel(ARQUIVO_DEVEDORES, sheet_name="Cancelados")

    df_base["Nro Funcional"] = pd.to_numeric(df_base["Nro Funcional"], errors="coerce").astype("Int64").astype(str)
    df_dividas["Funcional"] = pd.to_numeric(df_dividas["Funcional"], errors="coerce").astype("Int64").astype(str)
    df_cancelados["Funcional"] = pd.to_numeric(df_cancelados["Funcional"], errors="coerce").astype("Int64").astype(str)

    hoje = datetime.today()
    contexto_datas = _preparar_contextos_data(hoje)

    print(f"[INFO] Data atual: {contexto_datas['dia']}/{contexto_datas['mes']}/{contexto_datas['ano']}")
    print(f"[INFO] Vencimento considerado: dia {contexto_datas['ultimo_dia_util']}")

    pdf_map = mapear_pdfs_por_funcional(PASTA_PDFS_ENVIO, df_base)
    print(f"[INFO] Comprovantes de e-mail (PDFs) vinculados: {len(pdf_map)}")

    contadores = {"base": 0, "desligado": 0, "aviso": 0, "cancelado": 0, "ignorado": 0}

    for _, linha in df_base.iterrows():
        condicao = limpa(linha.get("condição")).lower()

        if "não enviar" in condicao or "nao enviar" in condicao:
            contadores["ignorado"] += 1
            continue

        if condicao == "desligado":
            gerar_carta_base_ou_desligado(
                linha, TEMPLATE_DESLIGADO, contexto_datas, pdf_map, CARTAS_BASE_DIR
            )
            contadores["desligado"] += 1

        elif condicao == "aviso":
            ok = gerar_carta_aviso_ou_cancelado(
                linha, condicao, TEMPLATE_AVISO, df_dividas, contexto_datas, CARTAS_CANCELADOS_DIR
            )
            if ok:
                contadores["aviso"] += 1

        elif condicao == "cancelado":
            ok = gerar_carta_aviso_ou_cancelado(
                linha, condicao, TEMPLATE_CANCELADO, df_cancelados, contexto_datas, CARTAS_CANCELADOS_DIR
            )
            if ok:
                contadores["cancelado"] += 1

        elif condicao in ("", "nan"):
            gerar_carta_base_ou_desligado(
                linha, TEMPLATE_BASE, contexto_datas, pdf_map, CARTAS_BASE_DIR
            )
            contadores["base"] += 1

    print("\n[OK] Resumo da geração de cartas mensais:")
    print(f"   * Cartas Base: {contadores['base']} (em {CARTAS_BASE_DIR.relative_to(CARTAS_BASE_DIR.parent.parent.parent)})")
    print(f"   * Cartas Desligado: {contadores['desligado']} (em {CARTAS_BASE_DIR.relative_to(CARTAS_BASE_DIR.parent.parent.parent)})")
    print(f"   * Cartas Aviso: {contadores['aviso']} (em {CARTAS_CANCELADOS_DIR.relative_to(CARTAS_CANCELADOS_DIR.parent.parent.parent)})")
    print(f"   * Cartas Cancelado: {contadores['cancelado']} (em {CARTAS_CANCELADOS_DIR.relative_to(CARTAS_CANCELADOS_DIR.parent.parent.parent)})")
    print(f"   * Ignorados ('não enviar'): {contadores['ignorado']}")


def gerar_todas_cartas_multa():
    """Gera cartas de multa para servidores elegíveis."""
    print("\n" + "=" * 60)
    print(">>> INICIANDO GERACAO DE CARTAS DE MULTA")
    print("=" * 60)

    CARTAS_MULTA_DIR.mkdir(parents=True, exist_ok=True)
    # Limpar arquivos anteriores na pasta de multa para evitar arquivos obsoletos
    for arquivo_antigo in CARTAS_MULTA_DIR.glob("*.docx"):
        try:
            arquivo_antigo.unlink()
        except Exception:
            pass

    if not ARQUIVO_BASE.exists() or not ARQUIVO_DEVEDORES.exists():
        print("[ERRO] Arquivos de entrada ausentes.")
        return

    df_base = pd.read_excel(ARQUIVO_BASE, engine="odf")
    df_dividas = pd.read_excel(ARQUIVO_DEVEDORES, sheet_name="Inadimplentes")

    df_base["Nro Funcional"] = pd.to_numeric(df_base["Nro Funcional"], errors="coerce").astype("Int64").astype(str)
    df_dividas["Funcional"] = pd.to_numeric(df_dividas["Funcional"], errors="coerce").astype("Int64").astype(str)

    hoje = datetime.today()
    contexto_datas = _preparar_contextos_data(hoje)

    # Nova regra: identifica parcelas onde o Saldo (Atualizado) é positivo e MENOR que o Principal (Saldo).
    # Isso indica que o servidor pagou o boleto com atraso e faltou a cobrança da diferença/juros.
    df_dividas["data_venc_dt"] = pd.to_datetime(df_dividas["Data de Vencimento"], dayfirst=True, errors="coerce")
    df_dividas["principal_num"] = pd.to_numeric(df_dividas["Principal (Saldo)"], errors="coerce").fillna(0)
    df_dividas["saldo_num"] = pd.to_numeric(df_dividas["Saldo (Atualizado)"], errors="coerce").fillna(0)

    multas_elegiveis = df_dividas[
        (df_dividas["saldo_num"] > 0) &
        (df_dividas["saldo_num"] < df_dividas["principal_num"])
    ]

    pdf_map = mapear_pdfs_por_funcional(PASTA_PDFS_ENVIO, df_base)
    total_geradas = 0

    for _, divida in multas_elegiveis.iterrows():
        matricula = divida["Funcional"]
        pessoa_base = df_base[df_base["Nro Funcional"] == matricula]
        if pessoa_base.empty:
            continue

        linha_base = pessoa_base.iloc[0]
        condicao = limpa(linha_base.get("condição")).lower()
        if condicao not in ("", "nan"):
            continue  # apenas para condição base

        funcionario_raw = linha_base["Funcionário"]
        if matricula in pdf_map:
            caminho_pdf = PASTA_PDFS_ENVIO / pdf_map[matricula]
            texto_pdf = extrair_texto_pdf(caminho_pdf)
            contexto_email = extrair_data_email_do_pdf(texto_pdf)
        else:
            contexto_email = {"dia_email": "dia", "mes_email": "mês", "ano_email": "ano"}

        endereco_completo = limpa(linha_base.get("endereço"))
        complemento = limpa(linha_base.get("complemento"))
        bairro = limpa(linha_base.get("bairro"))
        if complemento:
            endereco_completo += f" – {complemento}"
        if bairro:
            endereco_completo += f" – {bairro}"

        saldo_multa = divida["saldo_num"]
        valor_mensal = float(linha_base.get("Total", 0) or 0)
        data_venc_original = divida["data_venc_dt"].strftime("%d/%m/%Y") if pd.notna(divida["data_venc_dt"]) else ""

        contexto = {
            "nome_cap": capitalizar_nome(funcionario_raw),
            "nome_upper": str(funcionario_raw).upper(),
            "linha_endereco": endereco_completo,
            "CEP": limpa(linha_base.get("CEP")),
            "cidade": limpa(linha_base.get("cidade")),
            "email": limpa(linha_base.get("mail")) or "mail",
            "valor": formatar_valor_br(valor_mensal),
            "valor_extenso": valor_por_extenso(valor_mensal),
            "valor_multa": formatar_valor_br(saldo_multa),
            "valor_multa_extenso": valor_por_extenso(saldo_multa),
            "multa_competencia": formata_competencia(divida.get("Mês/Ano")),
            "multa_vencimento": data_venc_original,
            **contexto_datas,
            **contexto_email,
        }

        doc = DocxTemplate(TEMPLATE_MULTA)
        doc.render(contexto)

        nome_seguro = limpar_nome_arquivo(f"{funcionario_raw} - multa")
        doc.save(CARTAS_MULTA_DIR / f"{nome_seguro}.docx")
        total_geradas += 1

    print(f"[OK] Total de cartas de multa geradas: {total_geradas} (em {CARTAS_MULTA_DIR})")
