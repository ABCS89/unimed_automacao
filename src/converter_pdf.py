"""
converter_pdf.py - Conversão de documentos DOCX em PDF na pasta saida/
Etapa separada executada após a validação manual dos documentos Word.
"""
import os
from pathlib import Path
import shutil
import subprocess

from .config import (
    SAIDA_DIR,
    CARTAS_DIR,
    MEMORANDO_DIR,
    REFIS_DIR,
    LISTAS_DIR,
)


def localizar_executavel_soffice():
    """Localiza o executável do LibreOffice (soffice) no sistema."""
    # 1. Checar se está no PATH
    caminho = shutil.which("soffice") or shutil.which("libreoffice")
    if caminho:
        return caminho

    # 2. Caminhos padrão do Windows
    candidatos = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files\OpenOffice 4\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\OpenOffice 4\program\soffice.exe"),
    ]

    for p in candidatos:
        if p.exists():
            return str(p)

    return None


def converter_lote_libreoffice(arquivos_docx, pasta_destino, executavel_soffice):
    """Converte um lote de arquivos .docx para .pdf usando o LibreOffice em modo headless."""
    if not arquivos_docx:
        return 0

    pasta_destino = Path(pasta_destino)
    pasta_destino.mkdir(parents=True, exist_ok=True)

    # Filtra arquivos que não sejam temporários do Word (~$...)
    arquivos_validos = [str(f) for f in arquivos_docx if not Path(f).name.startswith("~$")]
    if not arquivos_validos:
        return 0

    cmd = [
        executavel_soffice,
        "--headless",
        "--convert-to", "pdf",
        *arquivos_validos,
        "--outdir", str(pasta_destino),
    ]

    try:
        resultado = subprocess.run(cmd, capture_output=True, text=True)
        if resultado.returncode == 0:
            return len(arquivos_validos)
        else:
            print(f"  [AVISO] Ocorreu uma mensagem durante a conversão: {resultado.stderr.strip()}")
            return len(arquivos_validos)
    except Exception as e:
        print(f"  [ERRO] Falha ao executar conversão: {e}")
        return 0


def converter_pasta(pasta_origem, pasta_destino=None, recursivo=True):
    """Varre uma pasta em busca de arquivos .docx e converte para .pdf."""
    pasta_origem = Path(pasta_origem)
    if not pasta_origem.exists():
        print(f"[AVISO] Pasta não encontrada: {pasta_origem}")
        return 0

    soffice = localizar_executavel_soffice()
    if not soffice:
        print("[ERRO] LibreOffice (soffice.exe) não foi encontrado no sistema.")
        print("       Instale o LibreOffice ou adicione soffice.exe ao PATH do Windows.")
        return 0

    # Localizar pastas que contenham arquivos .docx
    total_convertidos = 0
    if recursivo:
        subpastas = {f.parent for f in pasta_origem.glob("**/*.docx") if not f.name.startswith("~$")}
    else:
        subpastas = [pasta_origem] if list(pasta_origem.glob("*.docx")) else []

    if not subpastas:
        print(f"[INFO] Nenhum arquivo .docx encontrado em {pasta_origem.name}")
        return 0

    for pasta in sorted(subpastas):
        arquivos = [f for f in pasta.glob("*.docx") if not f.name.startswith("~$")]
        if not arquivos:
            continue

        destino = pasta_destino if pasta_destino else pasta
        print(f"[INFO] Convertendo {len(arquivos)} arquivo(s) em: {pasta.relative_to(SAIDA_DIR.parent)}")

        qtd = converter_lote_libreoffice(arquivos, destino, soffice)
        total_convertidos += qtd

    return total_convertidos


def converter_todas_cartas():
    print("\n" + "=" * 60)
    print(">>> CONVERTENDO CARTAS PARA PDF")
    print("=" * 60)
    qtd = converter_pasta(CARTAS_DIR, recursivo=True)
    print(f"[OK] Total de cartas convertidas para PDF: {qtd}\n")
    return qtd


def converter_memorando():
    print("\n" + "=" * 60)
    print(">>> CONVERTENDO MEMORANDO PARA PDF")
    print("=" * 60)
    qtd = converter_pasta(MEMORANDO_DIR, recursivo=False)
    print(f"[OK] Memorando convertido para PDF: {qtd}\n")
    return qtd


def converter_refis():
    print("\n" + "=" * 60)
    print(">>> CONVERTENDO TERMOS DO REFIS PARA PDF")
    print("=" * 60)
    qtd = converter_pasta(REFIS_DIR, recursivo=False)
    print(f"[OK] Termos do Refis convertidos para PDF: {qtd}\n")
    return qtd


def converter_lista():
    print("\n" + "=" * 60)
    print(">>> CONVERTENDO LISTA DE SERVIDORES PARA PDF")
    print("=" * 60)
    qtd = converter_pasta(LISTAS_DIR, recursivo=False)
    print(f"[OK] Lista de servidores convertida para PDF: {qtd}\n")
    return qtd


def converter_toda_saida():
    print("\n" + "=" * 60)
    print(">>> CONVERTENDO TODOS OS DOCUMENTOS DE SAIDA/ PARA PDF")
    print("=" * 60)
    qtd = converter_pasta(SAIDA_DIR, recursivo=True)
    print("\n" + "#" * 60)
    print(f"[SUCESSO] CONVERSAO CONCLUIDA! Total de PDFs gerados: {qtd}")
    print("#" * 60 + "\n")
    return qtd


def menu_conversao_pdf():
    """Menu exclusivo para conversão manual/separada para PDF."""
    while True:
        print("\n" + "=" * 60)
        print("    CONVERSAO DE DOCUMENTOS PARA PDF (ETAPA SEPARADA)")
        print("=" * 60)
        print("  1. [CARTAS] Converter apenas Cartas (Base, Aviso, Multa)")
        print("  2. [MEMORANDO] Converter apenas Memorando")
        print("  3. [REFIS] Converter apenas Termos do Refis")
        print("  4. [LISTA] Converter apenas Lista de Servidores")
        print("  5. [TUDO] Converter TODOS os arquivos .docx da pasta saida/")
        print("  0. [VOLTAR] Voltar ao Menu Principal")
        print("=" * 60)

        opcao = input("Digite a opcao desejada [0-5]: ").strip()

        if opcao == "1":
            converter_todas_cartas()
        elif opcao == "2":
            converter_memorando()
        elif opcao == "3":
            converter_refis()
        elif opcao == "4":
            converter_lista()
        elif opcao == "5":
            converter_toda_saida()
        elif opcao == "0":
            break
        else:
            print("\n[AVISO] Opção inválida! Digite um número entre 0 e 5.")


if __name__ == "__main__":
    converter_toda_saida()
