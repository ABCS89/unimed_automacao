"""
config.py - Configurações centrais e caminhos para unimed_automacao
"""
from pathlib import Path

# Diretório base do projeto unimed_automacao
BASE_DIR = Path(__file__).resolve().parent.parent

# Pastas de entrada
ENTRADA_DIR = BASE_DIR / "entrada"
ARQUIVO_BASE = ENTRADA_DIR / "teste.ods"
ARQUIVO_DEVEDORES = ENTRADA_DIR / "devedores.xlsx"
PASTA_PDFS_ENVIO = ENTRADA_DIR / "pdfs_envio"

# Pastas de templates
TEMPLATES_DIR = BASE_DIR / "templates"
EMAILS_TEMPLATES_DIR = TEMPLATES_DIR / "emails"

TEMPLATE_BASE = TEMPLATES_DIR / "template_base.docx"
TEMPLATE_DESLIGADO = TEMPLATES_DIR / "template_desligado.docx"
TEMPLATE_AVISO = TEMPLATES_DIR / "template_aviso.docx"
TEMPLATE_CANCELADO = TEMPLATES_DIR / "template_cancelado.docx"
TEMPLATE_MULTA = TEMPLATES_DIR / "template_base_multa.docx"
TEMPLATE_MEMORANDO = TEMPLATES_DIR / "template_memorando.docx"
TEMPLATE_REFIS = TEMPLATES_DIR / "template_refis.docx"
TEMPLATE_LISTA = TEMPLATES_DIR / "template_lista.docx"

# Pastas de saída
SAIDA_DIR = BASE_DIR / "saida"
CARTAS_DIR = SAIDA_DIR / "cartas"
CARTAS_BASE_DIR = CARTAS_DIR / "base"
CARTAS_CANCELADOS_DIR = CARTAS_DIR / "cancelados_aviso"
CARTAS_MULTA_DIR = CARTAS_DIR / "multa"
EMAILS_DIR = SAIDA_DIR / "emails"
MEMORANDO_DIR = SAIDA_DIR / "memorando"
REFIS_DIR = SAIDA_DIR / "refis"
LISTAS_DIR = SAIDA_DIR / "listas"

# Constantes de regras de negócio
MESES_PT = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}

MESES_ABREV = [
    "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
    "Jul", "Ago", "Set", "Out", "Nov", "Dez"
]
