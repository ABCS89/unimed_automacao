"""
lista.py - Geração de listagem consolidada de servidores em Word
"""
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate

from .config import (
    ARQUIVO_BASE,
    TEMPLATE_LISTA,
    LISTAS_DIR,
)
from .utils import limpa, capitalizar_nome


def gerar_lista():
    """Gera o documento com a lista de servidores em saida/listas/."""
    print("\n" + "=" * 60)
    print(">>> INICIANDO GERACAO DE LISTA CONSOLIDADA")
    print("=" * 60)

    LISTAS_DIR.mkdir(parents=True, exist_ok=True)

    if not ARQUIVO_BASE.exists() or not TEMPLATE_LISTA.exists():
        print("[ERRO] Arquivo teste.ods ou template_lista.docx ausente.")
        return

    df = pd.read_excel(ARQUIVO_BASE, engine="odf")
    coluna_nome = "Funcionário"

    pessoas = []
    for _, row in df.iterrows():
        nome = limpa(row.get(coluna_nome))
        if nome:
            pessoas.append({
                "nome": nome.upper(),
                "nome_cap": capitalizar_nome(nome),
            })

    doc = DocxTemplate(TEMPLATE_LISTA)
    doc.render({"pessoas": pessoas})

    caminho_saida = LISTAS_DIR / "lista_servidores.docx"
    doc.save(caminho_saida)

    print(f"\n[OK] Lista consolidada gerada com sucesso: {caminho_saida.name}")
    print(f"   * Total de servidores: {len(pessoas)}")

