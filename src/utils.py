"""
utils.py - Funções utilitárias compartilhadas em unimed_automacao
"""
import calendar
from datetime import datetime, timedelta
import os
from pathlib import Path
import re

import pandas as pd
import PyPDF2
from num2words import num2words
try:
    import holidays
    FERIADOS_SP = holidays.Brazil(subdiv="SP")
except ImportError:
    FERIADOS_SP = None

from .config import MESES_PT, MESES_ABREV


def limpa(valor):
    """Converte NaN em string vazia; senão retorna str().strip()."""
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def normalizar_nome(nome):
    """Remove acentos/pontuação e deixa em minúsculo, para comparar nomes."""
    nome = str(nome).lower()
    substituicoes = {
        "[áàãâä]": "a", "[éèêë]": "e", "[íìîï]": "i",
        "[óòõôö]": "o", "[úùûü]": "u", "[ç]": "c",
    }
    for padrao, letra in substituicoes.items():
        nome = re.sub(padrao, letra, nome)
    return re.sub(r"[^a-z0-9]", "", nome)


def capitalizar_nome(nome):
    """Capitaliza cada palavra de um nome (Ex: 'JOAO DA SILVA' -> 'Joao Da Silva')."""
    return " ".join(palavra.capitalize() for palavra in str(nome).lower().split())


def limpar_nome_arquivo(nome):
    """Remove caracteres inválidos para sistema de arquivos."""
    nome = str(nome).strip()
    return re.sub(r'[\\/*?:"<>|]', "", nome)


def formatar_valor_br(valor):
    """Converte 1234.5 em '1.234,50' (sem o prefixo R$)."""
    try:
        val = float(valor)
        return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return str(valor)


def formatar_moeda(valor):
    """Converte 1234.5 em 'R$ 1.234,50'."""
    val_br = formatar_valor_br(valor)
    return f"R$ {val_br}"


def valor_por_extenso(valor):
    """Gera valor por extenso em pt_BR (ex: 123.45 -> 'cento e vinte e três reais e quarenta e cinco centavos')."""
    try:
        valor_float = float(valor)
    except (ValueError, TypeError):
        return ""

    reais = int(valor_float)
    centavos = int(round((valor_float - reais) * 100))

    texto = f"{num2words(reais, lang='pt_BR')} reais"
    if centavos > 0:
        texto += f" e {num2words(centavos, lang='pt_BR')} centavos"
    return texto


def ultimo_dia_util_do_mes(ano, mes, considerar_feriados=True):
    """
    Retorna um objeto datetime correspondente ao último dia útil do mês,
    recuando enquanto cair em sábado, domingo ou feriados (SP).
    """
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    data = datetime(ano, mes, ultimo_dia)

    while True:
        if data.weekday() >= 5:
            data -= timedelta(days=1)
            continue
        if considerar_feriados and FERIADOS_SP and data.date() in FERIADOS_SP:
            data -= timedelta(days=1)
            continue
        break

    return data


def data_por_extenso(data):
    """Retorna '1 de agosto de 2026'."""
    return f"{data.day} de {MESES_PT[data.month]} de {data.year}"


def mes_anterior(hoje=None):
    """Devolve (ano, mes) do mês anterior."""
    if hoje is None:
        hoje = datetime.today()
    if hoje.month == 1:
        return hoje.year - 1, 12
    return hoje.year, hoje.month - 1


def mes_referencia_texto(hoje=None):
    """Retorna 'agosto/2026' referente ao mês anterior."""
    ano, mes = mes_anterior(hoje)
    return f"{MESES_PT[mes]}/{ano}"


def formata_competencia(valor):
    """Converte data/competência em formato abreviado (Ex: 'Jun/26')."""
    if pd.isna(valor):
        return ""
    if isinstance(valor, str):
        val_str = valor.strip()
        if re.match(r"^[A-Za-z]{3}/\d{2,4}$", val_str):
            return val_str
    data = pd.to_datetime(valor, errors="coerce")
    if pd.isna(data):
        return str(valor)
    return f"{MESES_ABREV[data.month - 1]}/{str(data.year)[-2:]}"


def extrair_texto_pdf(caminho_pdf):
    """Extrai texto concatenado de todas as páginas de um PDF via PyPDF2."""
    texto = ""
    try:
        with open(caminho_pdf, "rb") as arquivo:
            leitor = PyPDF2.PdfReader(arquivo)
            for pagina in leitor.pages:
                texto += pagina.extract_text() or ""
    except Exception as erro:
        print(f"  ⚠️ Erro ao ler PDF {caminho_pdf}: {erro}")
    return texto


def extrair_data_email_do_pdf(texto_pdf):
    """
    Procura padrão 'Data DD/MM/AAAA' no texto do PDF e retorna dicionário com dia, mês por extenso e ano.
    Se não encontrar, retorna marcadores para revisão manual.
    """
    match = re.search(r"Data\s+(\d{2}/\d{2}/\d{4})", texto_pdf)
    if not match:
        return {"dia_email": "dia", "mes_email": "mês", "ano_email": "ano"}

    data_obj = datetime.strptime(match.group(1), "%d/%m/%Y")
    return {
        "dia_email": data_obj.day,
        "mes_email": MESES_PT[data_obj.month],
        "ano_email": data_obj.year,
    }


def mapear_pdfs_por_funcional(pasta_pdfs, df_base):
    """
    Varre a pasta de PDFs de envio e vincula cada PDF ao Nro Funcional
    correspondente da planilha base.
    """
    mapa = {}
    pasta_path = Path(pasta_pdfs)
    if not pasta_path.exists():
        return mapa

    arquivos_pdf = [f for f in os.listdir(pasta_path) if f.lower().endswith(".pdf")]

    for nome_arquivo in arquivos_pdf:
        nome_extraido = os.path.splitext(nome_arquivo)[0].split(" - ")[0]
        nome_normalizado = normalizar_nome(nome_extraido)

        for _, linha in df_base.iterrows():
            nome_planilha = normalizar_nome(linha.get("Funcionário", ""))
            if nome_normalizado and (nome_normalizado in nome_planilha or nome_planilha in nome_normalizado):
                mapa[linha["Nro Funcional"]] = nome_arquivo
                break

    return mapa
