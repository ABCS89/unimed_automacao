"""
retroativos.py - Geração de documentos de cobrança de retroativos para servidores em folha
Gera exclusivamente a Notificação formal de desconto de ofício (saida/retroativos/oficio/).
"""
import copy
from datetime import datetime
import io
from pathlib import Path
import re
import subprocess
import unicodedata
import urllib.parse
import xml.etree.ElementTree as ET
import zipfile
import docx
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import pandas as pd
from docxtpl import DocxTemplate

from .config import (
    ARQUIVO_BASE,
    ARQUIVO_DEVEDORES,
    TEMPLATE_RETROATIVO_OFICIO,
    TEMPLATE_ETIQUETA,
    TEMPLATE_LISTA,
    RETROATIVOS_DIR,
    RETROATIVOS_OFICIO_DIR,
    LISTAS_DIR,
    EMAILS_DIR,
    EMAILS_RETROATIVOS_MD,
    MESES_PT,
)
from .utils import (
    limpa,
    normalizar_nome,
    capitalizar_nome,
    formatar_valor_br,
    limpar_nome_arquivo,
)
from .converter_pdf import localizar_executavel_soffice


def _formatar_funcional_display(val):
    """Formata matrícula com máscara padrão de Piracicaba (ex: 232882 -> 23.288-2)."""
    v = str(val).strip()
    digitos = re.sub(r"\D", "", v)
    if len(digitos) == 6:
        return f"{digitos[:2]}.{digitos[2:5]}-{digitos[5]}"
    return v


def _proxima_folha(hoje=None):
    """Calcula o mês e ano da próxima folha de pagamento para desconto."""
    if hoje is None:
        hoje = datetime.today()
    if hoje.month == 12:
        return 1, hoje.year + 1
    return hoje.month + 1, hoje.year


def _obter_mes_ano_sequencial(mes_inicial, ano_inicial, offset):
    """
    Retorna (mes_nome, ano) dado o mês inicial (1..12), ano inicial e um offset (0, 1, 2...).
    Ex: mes_inicial=10, ano_inicial=2026, offset=0 -> ('Outubro', 2026)
        offset=3 -> ('Janeiro', 2027)
    """
    mes_zero = (mes_inicial - 1) + offset
    mes_num = (mes_zero % 12) + 1
    ano = ano_inicial + (mes_zero // 12)
    return MESES_PT[mes_num].capitalize(), ano


def _obter_mes_ano_extenso(val, venc=None):
    """
    Converte datas ou competências para formato por extenso:
    Exemplos:
      '2024-11-01' -> 'Novembro de 2024'
      '11/2024'    -> 'Novembro de 2024'
      'Ago/26'     -> 'Agosto de 2026'
    """
    if pd.isna(val) or not str(val).strip():
        val = venc
    if pd.isna(val) or not str(val).strip():
        return "Competência não informada"

    if isinstance(val, (datetime, pd.Timestamp)):
        mes_nome = MESES_PT.get(val.month, "").capitalize()
        return f"{mes_nome} de {val.year}"

    s = str(val).strip()
    dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.notna(dt):
        mes_nome = MESES_PT.get(dt.month, "").capitalize()
        return f"{mes_nome} de {dt.year}"

    # Siglas em português (ex: Ago/26, Nov/24)
    abrevs = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
    for idx, abrev in enumerate(abrevs, 1):
        if abrev in s.lower():
            mes_nome = MESES_PT.get(idx, "").capitalize()
            anos = re.findall(r"\d{2,4}", s)
            ano = int(anos[0]) if anos else datetime.today().year
            if ano < 100:
                ano += 2000
            return f"{mes_nome} de {ano}"

    return s


def _extrair_mes_normalizado(val, venc=None):
    """Extrai o nome do mês em formato normalizado (ex: 'fevereiro', 'abril')."""
    if pd.isna(val) or not str(val).strip():
        val = venc
    if pd.isna(val) or not str(val).strip():
        return ""
    if isinstance(val, (datetime, pd.Timestamp)):
        m_num = val.month
        nome = MESES_PT.get(m_num, "")
        return unicodedata.normalize("NFKD", nome).encode("ASCII", "ignore").decode("utf-8").lower().strip()
    s = str(val).strip()
    dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.notna(dt):
        nome = MESES_PT.get(dt.month, "")
        return unicodedata.normalize("NFKD", nome).encode("ASCII", "ignore").decode("utf-8").lower().strip()
    abrevs = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
    for idx, abrev in enumerate(abrevs, 1):
        if abrev in s.lower():
            nome = MESES_PT.get(idx, "")
            return unicodedata.normalize("NFKD", nome).encode("ASCII", "ignore").decode("utf-8").lower().strip()
    return ""


def _carregar_dados_teste_ods():
    """Carrega dados de Mensalidade e Coparticipação de todas as abas mensais de teste.ods."""
    mapa = {}
    if not ARQUIVO_BASE.exists():
        return mapa
    try:
        f_ods = pd.ExcelFile(ARQUIVO_BASE, engine="odf")
        for sheet in f_ods.sheet_names:
            sheet_norm = unicodedata.normalize("NFKD", str(sheet)).encode("ASCII", "ignore").decode("utf-8").lower().strip()
            df_teste = pd.read_excel(f_ods, sheet_name=sheet)
            if df_teste.empty:
                continue

            func_col = next((c for c in df_teste.columns if "func" in unicodedata.normalize("NFKD", str(c)).lower()), None)
            mensal_col = next((c for c in df_teste.columns if "mensal" in unicodedata.normalize("NFKD", str(c)).lower()), None)
            copart_col = next((c for c in df_teste.columns if "copart" in unicodedata.normalize("NFKD", str(c)).lower()), None)
            total_col = next((c for c in df_teste.columns if "total" in unicodedata.normalize("NFKD", str(c)).lower()), None)

            if not func_col:
                continue

            for _, row in df_teste.iterrows():
                func_raw = limpa(row.get(func_col))
                func_digitos = re.sub(r"\D", "", func_raw)
                if not func_digitos:
                    continue

                val_m = pd.to_numeric(row.get(mensal_col), errors="coerce") if mensal_col else None
                val_c = pd.to_numeric(row.get(copart_col), errors="coerce") if copart_col else None
                val_t = pd.to_numeric(row.get(total_col), errors="coerce") if total_col else None

                entry = {
                    "mensalidade": float(val_m) if pd.notna(val_m) else None,
                    "coparticipacao": float(val_c) if pd.notna(val_c) else None,
                    "total": float(val_t) if pd.notna(val_t) else None,
                }
                # Indexa por (matrícula, mês da aba)
                mapa[(func_digitos, sheet_norm)] = entry
                # Fallback genérico por matrícula
                mapa[func_digitos] = entry
    except Exception as e:
        print(f"[AVISO] Não foi possível ler teste.ods para detalhamento de valores: {e}")
    return mapa


def _extrair_servidores_folha(arquivo_devedores=ARQUIVO_DEVEDORES):
    """
    Varre as abas Inadimplentes e Cancelados,
    filtrando registros onde a Situação do Vínculo contém 'folha'.
    """
    if not arquivo_devedores.exists():
        print(f"[ERRO] Arquivo não encontrado: {arquivo_devedores}")
        return pd.DataFrame()

    excel_dev = pd.ExcelFile(arquivo_devedores)
    linhas_encontradas = []

    # Abas principais filtrando Sit. Vinculo com 'folha'
    for sheet in ["Inadimplentes", "Cancelados"]:
        if sheet in excel_dev.sheet_names:
            df = pd.read_excel(excel_dev, sheet_name=sheet)
            if "Sit. Vinculo" in df.columns:
                m = df["Sit. Vinculo"].astype(str).str.lower().str.contains("folha")
                df_folha = df[m].copy()
                df_folha["_origem_aba"] = sheet
                linhas_encontradas.append(df_folha)

    if not linhas_encontradas:
        return pd.DataFrame()

    df_total = pd.concat(linhas_encontradas, ignore_index=True)
    # Remove eventuais duplicatas exatas de funcional + número da guia
    if "#" in df_total.columns:
        df_total = df_total.drop_duplicates(subset=["Funcional", "#"], keep="first")
    return df_total


def _inserir_tabela_valores_em_aberto(doc, target_p, itens_debito):
    """
    Substitui o parágrafo target_p por uma tabela de 3 colunas compacta e elegante:
    [Referência] | [Valor] | [Previsão de Desconto]
    com linhas finas, altura reduzida e texto em linha única para máxima economia de espaço vertical.
    """
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._tbl.tblPr

    # 1. Bordas sutis e finas
    tblBorders = parse_xml(
        '<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '  <w:top w:val="single" w:sz="4" w:space="0" w:color="D0D0D0"/>'
        '  <w:left w:val="single" w:sz="4" w:space="0" w:color="D0D0D0"/>'
        '  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D0D0"/>'
        '  <w:right w:val="single" w:sz="4" w:space="0" w:color="D0D0D0"/>'
        '  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        '  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        '</w:tblBorders>'
    )
    tblPr.append(tblBorders)

    # 2. Padding interno compacto (altura mínima sem desperdício de espaço)
    tblCellMar = parse_xml(
        '<w:tblCellMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '  <w:top w:w="25" w:type="dxa"/>'
        '  <w:left w:w="70" w:type="dxa"/>'
        '  <w:bottom w:w="25" w:type="dxa"/>'
        '  <w:right w:w="70" w:type="dxa"/>'
        '</w:tblCellMar>'
    )
    tblPr.append(tblCellMar)

    # Distribuição que garante texto de referência completo em 1 única linha
    larguras = [Pt(225), Pt(75), Pt(150)]

    # 3. Cabeçalho
    hdr_row = table.rows[0]
    hdr_trPr = hdr_row._tr.get_or_add_trPr()
    hdr_trPr.append(parse_xml('<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
    hdr_trPr.append(parse_xml('<w:tblHeader xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))

    titulos = ["Referência", "Valor", "Previsão de Desconto"]
    aligns_hdr = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT]

    for c, tit, align, w in zip(hdr_row.cells, titulos, aligns_hdr, larguras):
        c.width = w
        p = c.paragraphs[0]
        p.text = tit
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        p.runs[0].bold = True
        p.runs[0].font.size = Pt(8.5)
        c._tc.get_or_add_tcPr().append(
            parse_xml('<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="F4F4F4"/>')
        )

    # 4. Linhas de dados compactas
    for item in itens_debito:
        row = table.add_row()
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml('<w:cantSplit xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
        row_cells = row.cells

        row_cells[0].width = larguras[0]
        p0 = row_cells[0].paragraphs[0]
        p0.text = item["referencia"]
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        p0.paragraph_format.line_spacing = 1.0
        p0.runs[0].font.size = Pt(8.0)

        row_cells[1].width = larguras[1]
        p1 = row_cells[1].paragraphs[0]
        p1.text = item["valor"]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        p1.paragraph_format.line_spacing = 1.0
        p1.runs[0].font.size = Pt(8.0)

        row_cells[2].width = larguras[2]
        p2 = row_cells[2].paragraphs[0]
        p2.text = item["previsao"]
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        p2.paragraph_format.line_spacing = 1.0
        p2.runs[0].font.size = Pt(8.0)

    target_p._p.addnext(table._tbl)
    target_p._p.getparent().remove(target_p._p)


def _carregar_mapa_cadastros():
    """
    Varre todas as abas de teste.ods e constrói um índice unificado
    com nome, endereço completo, CEP, cidade, UF e e-mail de cada servidor.
    """
    mapa = {}
    if not ARQUIVO_BASE.exists():
        return mapa
    try:
        f_ods = pd.ExcelFile(ARQUIVO_BASE, engine="odf")
        for sheet in f_ods.sheet_names:
            df = pd.read_excel(f_ods, sheet_name=sheet)
            if df.empty:
                continue
            col_nome = next((c for c in df.columns if "func" in str(c).lower() and "nro" not in str(c).lower()), None)
            col_matr = next((c for c in df.columns if "nro" in str(c).lower() or "funcional" in str(c).lower()), None)
            col_end = next((c for c in df.columns if "endere" in str(c).lower()), None)
            col_bairro = next((c for c in df.columns if "bairro" in str(c).lower()), None)
            col_comp = next((c for c in df.columns if "complem" in str(c).lower()), None)
            col_cep = next((c for c in df.columns if "cep" in str(c).lower()), None)
            col_cid = next((c for c in df.columns if "cidade" in str(c).lower()), None)
            col_uf = next((c for c in df.columns if "uf" in str(c).lower()), None)
            col_mail = next((c for c in df.columns if "mail" in str(c).lower() or "email" in str(c).lower()), None)

            if not col_nome:
                continue

            for _, r in df.iterrows():
                nome_raw = limpa(r.get(col_nome))
                if not nome_raw:
                    continue
                n_norm = normalizar_nome(nome_raw)
                matr_raw = re.sub(r"\D", "", limpa(r.get(col_matr))) if col_matr else ""
                end = limpa(r.get(col_end)) if col_end else ""
                bairro = limpa(r.get(col_bairro)) if col_bairro else ""
                comp = limpa(r.get(col_comp)) if col_comp else ""
                cep = limpa(r.get(col_cep)) if col_cep else ""
                cid = limpa(r.get(col_cid)) if col_cid else "PIRACICABA"
                uf = limpa(r.get(col_uf)) if col_uf else "SP"
                mail = limpa(r.get(col_mail)) if col_mail else ""
                if mail and ("@" not in mail or mail.lower() == "nan"):
                    mail = ""

                cep_dig = re.sub(r"\D", "", cep)
                if len(cep_dig) == 8:
                    cep = f"{cep_dig[:5]}-{cep_dig[5:]}"

                entry = mapa.get(n_norm, {
                    "nome": nome_raw,
                    "matricula": matr_raw,
                    "endereco": "",
                    "bairro": "",
                    "complemento": "",
                    "cep": "",
                    "cidade": "PIRACICABA",
                    "uf": "SP",
                    "email": "",
                })

                if end:
                    entry["endereco"] = end
                    entry["bairro"] = bairro
                    entry["complemento"] = comp
                    entry["cep"] = cep
                    entry["cidade"] = cid or "PIRACICABA"
                    entry["uf"] = uf or "SP"
                if mail:
                    entry["email"] = mail

                mapa[n_norm] = entry
                if matr_raw:
                    mapa[matr_raw] = entry
    except Exception as e:
        print(f"[AVISO] Falha ao extrair histórico cadastral de teste.ods: {e}")
    return mapa


def _gerar_etiquetas_retroativos(servidores_etiquetas):
    """
    Gera o arquivo de etiquetas para correspondência física em:
    saida/retroativos/etiquetas_retroativos.odt (e .pdf).
    Utiliza o modelo entrada/template_etiqueta.odt preservando
    dimensões (10x6cm) e os 2 blocos de etiqueta por folha.
    """
    if not TEMPLATE_ETIQUETA.exists():
        print(f"[AVISO] Template de etiqueta não encontrado: {TEMPLATE_ETIQUETA}")
        return None, None

    NS = {
        "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
        "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
        "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
        "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
        "svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    }
    for prefix, uri in NS.items():
        ET.register_namespace(prefix, uri)

    with zipfile.ZipFile(TEMPLATE_ETIQUETA, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    root = ET.fromstring(files["content.xml"])
    text_elem = root.find(".//office:text", NS)
    orig_frame = text_elem.find(".//draw:frame", NS)
    if orig_frame is None:
        print("[AVISO] Quadro de etiquetas não localizado no template ODT.")
        return None, None

    def substituir_display(elem, dados):
        tag = elem.tag.split("}")[-1]
        if tag == "database-display":
            col = elem.get(f"{{{NS['text']}}}column-name", "")
            txt = dados.get(col, "")
            elem.tag = f"{{{NS['text']}}}span"
            elem.attrib.clear()
            elem.text = txt
        for child in list(elem):
            substituir_display(child, dados)

    seq = text_elem.find(".//text:sequence-decls", NS)
    text_elem.clear()
    if seq is not None:
        text_elem.append(seq)

    auto_styles = root.find(".//office:automatic-styles", NS)
    if auto_styles is not None:
        style_break = ET.SubElement(auto_styles, f"{{{NS['style']}}}style", {
            f"{{{NS['style']}}}name": "P_PAGE_BREAK",
            f"{{{NS['style']}}}family": "paragraph",
            f"{{{NS['style']}}}parent-style-name": "Standard",
        })
        ET.SubElement(style_break, f"{{{NS['style']}}}paragraph-properties", {
            f"{{{NS['fo']}}}break-before": "page",
        })

    for idx, serv in enumerate(servidores_etiquetas, 1):
        f = copy.deepcopy(orig_frame)
        f.set(f"{{{NS['draw']}}}name", f"Quadro{idx}")
        f.set(f"{{{NS['text']}}}anchor-type", "paragraph")
        if f"{{{NS['text']}}}anchor-page-number" in f.attrib:
            del f.attrib[f"{{{NS['text']}}}anchor-page-number"]

        dados_tpl = {
            "Funcionário": serv.get("nome", "").upper(),
            "Endereço": serv.get("endereco", "").upper(),
            "Bairro": serv.get("bairro", "").upper(),
            "Complemento": serv.get("complemento", "").upper(),
            "CEP": serv.get("cep", ""),
            "cidade": serv.get("cidade", "PIRACICABA").upper(),
            "uf": serv.get("uf", "SP").upper(),
        }
        substituir_display(f, dados_tpl)

        p = ET.Element(f"{{{NS['text']}}}p")
        if idx > 1:
            p.set(f"{{{NS['text']}}}style-name", "P_PAGE_BREAK")
        else:
            p.set(f"{{{NS['text']}}}style-name", "Standard")
        p.append(f)
        text_elem.append(p)

    files["content.xml"] = ET.tostring(root, encoding="utf-8")

    RETROATIVOS_DIR.mkdir(parents=True, exist_ok=True)
    caminho_odt = RETROATIVOS_DIR / "etiquetas_retroativos.odt"
    with zipfile.ZipFile(caminho_odt, "w") as zout:
        for name, content in files.items():
            zout.writestr(name, content)

    caminho_pdf = None
    soffice = localizar_executavel_soffice()
    if soffice:
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", str(caminho_odt), "--outdir", str(RETROATIVOS_DIR)],
                check=True,
                capture_output=True,
            )
            caminho_pdf = RETROATIVOS_DIR / "etiquetas_retroativos.pdf"
        except Exception as e:
            print(f"[AVISO] Não foi possível converter etiquetas para PDF: {e}")

    return caminho_odt, caminho_pdf


def _gerar_tabela_markdown_retroativos(itens_debito, total_formatado):
    """Gera tabela visual profissional formatada em HTML para o corpo do e-mail de retroativos."""
    linhas = [
        '<table width="100%" border="1" cellspacing="0" cellpadding="8" bordercolor="#b0b0b0" style="border-collapse: collapse; width: 100%; border: 1px solid #b0b0b0; font-family: Calibri, Arial, sans-serif; font-size: 13px; margin: 15px 0;">',
        '  <thead>',
        '    <tr bgcolor="#e6ecf5" style="background-color: #e6ecf5; font-weight: bold; text-align: center;">',
        '      <th width="45%" bgcolor="#e6ecf5" align="left" style="border: 1px solid #b0b0b0; padding: 8px; text-align: left;">Referência</th>',
        '      <th width="25%" bgcolor="#e6ecf5" align="center" style="border: 1px solid #b0b0b0; padding: 8px; text-align: center;">Valor</th>',
        '      <th width="30%" bgcolor="#e6ecf5" align="left" style="border: 1px solid #b0b0b0; padding: 8px; text-align: left;">Previsão de Desconto</th>',
        '    </tr>',
        '  </thead>',
        '  <tbody>',
    ]
    for item in itens_debito:
        linhas.append('    <tr>')
        linhas.append(f'      <td width="45%" align="left" style="border: 1px solid #b0b0b0; padding: 8px; text-align: left;">{item["referencia"]}</td>')
        linhas.append(f'      <td width="25%" align="center" style="border: 1px solid #b0b0b0; padding: 8px; text-align: center; font-weight: bold;">{item["valor"]}</td>')
        linhas.append(f'      <td width="30%" align="left" style="border: 1px solid #b0b0b0; padding: 8px; text-align: left;">{item["previsao"]}</td>')
        linhas.append('    </tr>')
    linhas.append('    <tr bgcolor="#f9f9f9" style="background-color: #f9f9f9; font-weight: bold;">')
    linhas.append('      <td width="45%" align="left" style="border: 1px solid #b0b0b0; padding: 8px; text-align: left;"><strong>Total do Débito</strong></td>')
    linhas.append(f'      <td width="25%" align="center" style="border: 1px solid #b0b0b0; padding: 8px; text-align: center; font-weight: bold; color: #b30000;">R$ {total_formatado}</td>')
    linhas.append('      <td width="30%" align="left" style="border: 1px solid #b0b0b0; padding: 8px; text-align: left;">Quitação em folha</td>')
    linhas.append('    </tr>')
    linhas.append('  </tbody>')
    linhas.append('</table>')
    return "\n".join(linhas)


def _gerar_emails_retroativos(lista_dados_emails, hoje=None):
    """
    Salva a lista de e-mails de cobrança de retroativos em saida/emails/emails_retroativos.md.
    """
    if hoje is None:
        hoje = datetime.today()

    EMAILS_DIR.mkdir(parents=True, exist_ok=True)
    mes_nome = MESES_PT[hoje.month]
    ano_str = str(hoje.year)

    conteudo_emails = []
    for d in lista_dados_emails:
        nome_cap = d["nome_cap"]
        mail = d["email"].strip()
        mes_upper = d["mes_upper"]
        total_formatado = d["total_formatado"]
        tabela_html = _gerar_tabela_markdown_retroativos(d["itens_tabela"], total_formatado)

        corpo = (
            f"A/C {nome_cap}\n\n"
            f"Comunicamos que, tendo em vista a notificação prévia a V. Sa., referente aos valores pendentes "
            f"do Plano de Saúde Unimed de competências anteriores que deixaram de ser descontados em folha "
            f"no período oportuno, e considerando a ausência de quitação voluntária dos referidos débitos; "
            f"Informamos que, com o objetivo de evitar a inscrição do montante em Dívida Ativa Municipal "
            f"(conforme previsto no art. 5º da Lei Municipal nº 9.988, de 14 de novembro de 2023), será "
            f"realizado o desconto em folha de pagamento, a partir da folha do mês de {mes_upper}, de acordo "
            f"com o cronograma discriminado abaixo.\n\n"
            f"As guias em aberto serão canceladas conforme quitação em folha com a desconsideração dos encargos "
            f"financeiros, consolidando-se a quitação pelo valor principal devido (Total de R$ {total_formatado}).\n\n"
            f"{tabela_html}\n\n"
            f"* DETERMINAÇÃO PARA INÍCIO DE DESCONTO DE OFÍCIO NA FOLHA DE {mes_upper}.\n\n"
            f"Em caso de discordância ou eventual contraproposta, pedimos para que entre em contato com o DRH "
            f"da Prefeitura do Município de Piracicaba até o dia 20 de {mes_nome} de {ano_str}.\n\n"
            f"Em caso de dúvidas, entre em contato com o Setor de Controle de Frequência e Elaboração da "
            f"Folha de Pagamento / DRH da Prefeitura de Piracicaba pelo telefone 3403-1006.\n\n"
            f"Permanecemos à disposição para esclarecimentos."
        )

        assunto = f"Comunicado de Desconto em Folha – Plano de Saúde Unimed – {mes_upper}/{ano_str}"

        conteudo_emails.append({
            "nome": nome_cap,
            "email": mail,
            "assunto": assunto,
            "mensagem": corpo,
        })

    with open(EMAILS_RETROATIVOS_MD, "w", encoding="utf-8") as f:
        f.write("# 📧 Emails de Cobrança de Retroativos (Desconto em Folha)\n\n")
        f.write(f"> Total de destinatários: **{len(conteudo_emails)}**\n\n")

        for email in conteudo_emails:
            f.write("---\n\n")
            f.write(f"## 👤 {email['nome']}\n\n")

            mail_limpo = email["email"].strip()
            if mail_limpo and mail_limpo.lower() != "nan" and "@" in mail_limpo:
                assunto_enc = urllib.parse.quote(email["assunto"])
                mailto_link = f"mailto:{mail_limpo}?subject={assunto_enc}"
                f.write(f"**Para:** [`{mail_limpo}`]({mailto_link})  \n")
            else:
                f.write(f"**Para:** `E-mail não cadastrado`  \n")

            f.write(f"**Assunto:** `{email['assunto']}`  \n\n")
            f.write("### ✉️ Mensagem:\n\n")
            f.write(f"{email['mensagem']}\n\n")

    return EMAILS_RETROATIVOS_MD


def _gerar_lista_retroativos(servidores_dados, hoje=None):
    """
    Gera a lista formal de protocolo/assinatura/recebimento baseada em template_lista.docx
    salvando em saida/retroativos/lista_retroativos.docx (e .pdf) e cópia em saida/listas/.
    """
    if not TEMPLATE_LISTA.exists():
        print(f"[AVISO] Template de lista não encontrado: {TEMPLATE_LISTA}")
        return None, None

    if hoje is None:
        hoje = datetime.today()

    data_str = f"Piracicaba, {hoje.day} de {MESES_PT[hoje.month]} de {hoje.year}."

    nomes_unicos = sorted(list({s.get("nome", "").upper() for s in servidores_dados if s.get("nome")}))
    pessoas = [{"nome": n} for n in nomes_unicos]

    doc = DocxTemplate(TEMPLATE_LISTA)
    doc.render({"pessoas": pessoas})

    # Atualiza data e assunto
    for p in doc.paragraphs:
        if "Piracicaba," in p.text:
            p.text = data_str
        if "Assunto:" in p.text:
            p.text = "Assunto: Regularização de Valores Pendentes do Plano de Saúde Unimed (Retroativos)"

    # Remove linha residual vazia se houver
    if doc.tables:
        t = doc.tables[0]
        for row in list(t.rows):
            if not any(c.text.strip() for c in row.cells):
                t._tbl.remove(row._tr)

    RETROATIVOS_DIR.mkdir(parents=True, exist_ok=True)
    LISTAS_DIR.mkdir(parents=True, exist_ok=True)

    caminho_docx = RETROATIVOS_DIR / "lista_retroativos.docx"
    doc.save(caminho_docx)
    doc.save(LISTAS_DIR / "lista_retroativos.docx")

    caminho_pdf = None
    soffice = localizar_executavel_soffice()
    if soffice:
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", str(caminho_docx), "--outdir", str(RETROATIVOS_DIR)],
                check=True,
                capture_output=True,
            )
            caminho_pdf = RETROATIVOS_DIR / "lista_retroativos.pdf"
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", str(LISTAS_DIR / "lista_retroativos.docx"), "--outdir", str(LISTAS_DIR)],
                check=True,
                capture_output=True,
            )
        except Exception as e:
            print(f"[AVISO] Não foi possível converter lista de retroativos para PDF: {e}")

    return caminho_docx, caminho_pdf


def gerar_documentos_retroativos():
    """
    Gera todo o ecossistema de cobrança de retroativos:
    1. Notificações formais de desconto de ofício em saida/retroativos/oficio/
    2. Etiquetas de correspondência física em saida/retroativos/etiquetas_retroativos.odt (.pdf)
    3. Modelos de e-mails em saida/emails/emails_retroativos.md
    4. Lista de protocolo/assinatura em saida/retroativos/lista_retroativos.docx (.pdf)
    """
    print("\n" + "=" * 60)
    print(">>> INICIANDO GERACAO DE DOCUMENTOS DE COBRANCA DE RETROATIVOS")
    print("=" * 60)

    RETROATIVOS_OFICIO_DIR.mkdir(parents=True, exist_ok=True)

    # Limpar docx anteriores para evitar sobras
    for f in RETROATIVOS_OFICIO_DIR.glob("*.docx"):
        try:
            f.unlink()
        except Exception:
            pass

    if not TEMPLATE_RETROATIVO_OFICIO.exists():
        print(f"[ERRO] Template de ofício não encontrado: {TEMPLATE_RETROATIVO_OFICIO}")
        return

    df_folha = _extrair_servidores_folha()
    if df_folha.empty:
        print("[AVISO] Nenhum servidor com vínculo em folha encontrado para cobrança de retroativos.")
        return

    # Base cadastral completa (endereços, bairros, CEPs e e-mails)
    mapa_cadastros = _carregar_mapa_cadastros()

    hoje = datetime.today()
    mes_desc_num = hoje.month
    ano_desc = hoje.year
    mes_desc_upper = MESES_PT[hoje.month].upper()

    total_servidores = 0
    grupos = df_folha.groupby("Funcional")

    dados_etiquetas = []
    dados_emails = []

    for func, grupo in grupos:
        nome_raw = limpa(grupo.iloc[0].get("Nome"))
        if not nome_raw:
            continue

        func_digitos = re.sub(r"\D", "", str(func))
        matricula_display = _formatar_funcional_display(func)
        nome_cap = capitalizar_nome(nome_raw)
        n_norm = normalizar_nome(nome_raw)

        # Dados cadastrais obtidos de teste.ods
        cad = mapa_cadastros.get(n_norm) or mapa_cadastros.get(func_digitos, {})
        email_servidor = cad.get("email", "")

        dados_etiquetas.append({
            "nome": cad.get("nome") or nome_raw,
            "endereco": cad.get("endereco", ""),
            "bairro": cad.get("bairro", ""),
            "complemento": cad.get("complemento", ""),
            "cep": cad.get("cep", ""),
            "cidade": cad.get("cidade", "PIRACICABA"),
            "uf": cad.get("uf", "SP"),
        })

        # Soma dos valores de principal
        total_principal = pd.to_numeric(grupo["Principal (Saldo)"], errors="coerce").fillna(0).sum()
        total_formatado = formatar_valor_br(total_principal)

        # Salário fixo
        salario_val = None
        for col_sal in grupo.columns:
            col_norm = unicodedata.normalize("NFKD", str(col_sal)).encode("ASCII", "ignore").decode("utf-8").lower()
            if any(k in col_norm for k in ["salario", "vencimento", "remun"]):
                for raw_val in grupo[col_sal].dropna():
                    if isinstance(raw_val, (int, float)) and raw_val > 0:
                        salario_val = float(raw_val)
                        break
                    v_str = str(raw_val).replace("R$", "").replace("r$", "").replace("\xa0", " ").strip()
                    if "," in v_str:
                        v_str = v_str.replace(".", "").replace(",", ".")
                    try:
                        v_float = float(v_str)
                        if v_float > 0:
                            salario_val = v_float
                            break
                    except (ValueError, TypeError):
                        pass
                if salario_val:
                    break

        salario_str = f"R$ {formatar_valor_br(salario_val)}" if salario_val else "R$ ___________ (A preencher)"

        # 1. Ordenação cronológica para encontrar o boleto mais antigo
        grupo_ord = grupo.copy()
        dt_venc = pd.to_datetime(grupo_ord["Data de Vencimento"], dayfirst=True, errors="coerce")
        dt_comp = pd.to_datetime(grupo_ord["Mês/Ano"], dayfirst=True, errors="coerce")
        grupo_ord["_dt_sort"] = dt_venc.fillna(dt_comp).fillna(pd.Timestamp("2099-01-01"))
        grupo_ord = grupo_ord.sort_values(by="_dt_sort", ascending=True)

        linha_mais_antiga = grupo_ord.iloc[0]
        val_mais_antigo = float(linha_mais_antiga.get("Principal (Saldo)", 0) or 0)
        qtd_boletos = len(grupo_ord)

        # 2. Definição da programação de desconto baseada no boleto mais antigo
        if qtd_boletos == 1 or abs(total_principal - val_mais_antigo) < 0.01:
            simulacao_texto = f"1 parcela de R$ {formatar_valor_br(val_mais_antigo)}"
        else:
            simulacao_texto = f"Desconto mensal de R$ {formatar_valor_br(val_mais_antigo)} até a quitação"

        itens_tabela = []
        for idx_bol, (_, row) in enumerate(grupo_ord.iterrows()):
            comp_extenso = _obter_mes_ano_extenso(row.get("Mês/Ano"), row.get("Data de Vencimento"))
            mes_desconto_str, ano_desconto_num = _obter_mes_ano_sequencial(mes_desc_num, ano_desc, idx_bol)
            val_linha = float(row.get("Principal (Saldo)", 0) or 0)
            itens_tabela.append({
                "referencia": f"Valores referentes a {comp_extenso}",
                "valor": f"R$ {formatar_valor_br(val_linha)}",
                "previsao": f"Será descontado em {mes_desconto_str} {ano_desconto_num}",
            })

        dados_emails.append({
            "nome_cap": nome_cap,
            "email": email_servidor,
            "mes_upper": mes_desc_upper,
            "total_formatado": total_formatado,
            "itens_tabela": itens_tabela,
        })

        contexto = {
            "dia": str(hoje.day),
            "mes": MESES_PT[hoje.month],
            "ano": str(hoje.year),
            "nome_cap": nome_cap,
            "matricula": matricula_display,
            "total_debito": total_formatado,
            "salario_fixo": salario_str,
            "discriminacao_debito": "__TABELA_VALORES_ABERTO__",
            "simulacao_parcelas": simulacao_texto,
            "mes_upper": mes_desc_upper,
            "mes_desconto_upper": mes_desc_upper,
            "ano_desconto": str(ano_desc),
        }

        # Gerar documento de Ofício e injetar tabela formal
        doc_ofc = DocxTemplate(TEMPLATE_RETROATIVO_OFICIO)
        doc_ofc.render(contexto)

        buf = io.BytesIO()
        doc_ofc.save(buf)
        buf.seek(0)

        doc_final = docx.Document(buf)
        for p in doc_final.paragraphs:
            if "__TABELA_VALORES_ABERTO__" in p.text:
                _inserir_tabela_valores_em_aberto(doc_final, p, itens_tabela)
                break

        nome_arq_ofc = limpar_nome_arquivo(f"{nome_raw}.docx")
        doc_final.save(RETROATIVOS_OFICIO_DIR / nome_arq_ofc)

        total_servidores += 1

    # 3. Geração das Etiquetas de correspondência física
    caminho_odt, caminho_pdf = _gerar_etiquetas_retroativos(dados_etiquetas)

    # 4. Geração da Lista formal de Protocolo/Assinatura (template_lista.docx)
    caminho_lista_docx, caminho_lista_pdf = _gerar_lista_retroativos(dados_etiquetas, hoje)

    # 5. Geração dos E-mails em Markdown
    caminho_md = _gerar_emails_retroativos(dados_emails, hoje)

    print("\n[OK] Resumo completo da geração de cobrança de retroativos:")
    print(f"   * Servidores processados com vínculo em folha: {total_servidores}")
    print(f"   * Documentos de Ofício salvos em: {RETROATIVOS_OFICIO_DIR}")
    if caminho_odt:
        msg_etq = f"   * Etiquetas de correspondência salvas em: {caminho_odt.name}"
        if caminho_pdf:
            msg_etq += f" e {caminho_pdf.name}"
        print(msg_etq)
    if caminho_lista_docx:
        msg_lista = f"   * Lista de assinatura/protocolo salva em: {caminho_lista_docx.name}"
        if caminho_lista_pdf:
            msg_lista += f" e {caminho_lista_pdf.name}"
        print(msg_lista)
    if caminho_md:
        print(f"   * Modelos de e-mail salvos em: {caminho_md}")
