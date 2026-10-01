"""
main.py - Ponto de entrada centralizado para o ecossistema unimed_automacao
Permite rodar as rotinas via menu interativo ou diretamente por linha de comando.
"""
import argparse
import io
import sys
from pathlib import Path

# Garantir suporte a UTF-8 no terminal Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Adiciona o diretório atual ao sys.path para garantir importações relativas e diretas
sys.path.insert(0, str(Path(__file__).resolve().parent))

from src.cartas import gerar_todas_cartas_mensais, gerar_todas_cartas_multa
from src.emails import gerar_todos_emails
from src.memorando import gerar_memorando
from src.refis import gerar_todos_refis
from src.lista import gerar_lista
from src.converter_pdf import menu_conversao_pdf, converter_toda_saida


def executar_tudo():
    print("\n" + "#" * 60)
    print(">>> EXECUTANDO TODAS AS ROTINAS DO MES (GERACAO DOCX)")
    print("#" * 60)
    gerar_todas_cartas_mensais()
    gerar_todas_cartas_multa()
    gerar_todos_emails()
    gerar_memorando()
    gerar_todos_refis()
    gerar_lista()
    print("\n" + "#" * 60)
    print("[SUCESSO] TODAS AS ROTINAS DE GERACAO FORAM CONCLUIDAS!")
    print("Nota: Para converter os arquivos gerados em PDF apos a sua")
    print("validacao, utilize a Opcao 8 do menu ou '--acao converter_pdf'.")
    print("#" * 60 + "\n")


def menu_interativo():
    while True:
        print("\n" + "=" * 60)
        print("    AUTOMACAO UNIMED - DEPARTAMENTO DE RECURSOS HUMANOS")
        print("=" * 60)
        print("  1. [CARTAS] Gerar Cartas Mensais (Base, Desligados, Avisos, Cancelados)")
        print("  2. [MULTA] Gerar Cartas de Multa por Atraso")
        print("  3. [EMAILS] Gerar Modelos de E-mails com Tabelas de Debitos")
        print("  4. [MEMORANDO] Gerar Memorando Geral em Word")
        print("  5. [REFIS] Gerar Termos de Acordo do Refis")
        print("  6. [LISTA] Gerar Lista Consolidada de Servidores")
        print("  7. [TUDO] Executar TUDO (Rotina Mensal de Geracao DOCX)")
        print("  8. [CONVERTER PDF] Converter DOCX para PDF (Pos-Validacao)")
        print("  0. [SAIR] Sair")
        print("=" * 60)

        opcao = input("Digite a opcao desejada [0-8]: ").strip()

        if opcao == "1":
            gerar_todas_cartas_mensais()
        elif opcao == "2":
            gerar_todas_cartas_multa()
        elif opcao == "3":
            gerar_todos_emails()
        elif opcao == "4":
            gerar_memorando()
        elif opcao == "5":
            gerar_todos_refis()
        elif opcao == "6":
            gerar_lista()
        elif opcao == "7":
            executar_tudo()
        elif opcao == "8":
            menu_conversao_pdf()
        elif opcao == "0":
            print("\nEncerrando o programa. Até logo!\n")
            break
        else:
            print("\n[AVISO] Opcao invalida! Digite um numero entre 0 e 8.")


def main():
    parser = argparse.ArgumentParser(description="Automação do fluxo Unimed (DRH)")
    parser.add_argument(
        "--acao",
        choices=["cartas", "multa", "emails", "memorando", "refis", "lista", "tudo", "converter_pdf"],
        help="Executa uma ação diretamente sem abrir o menu interativo"
    )

    args = parser.parse_args()

    if args.acao == "cartas":
        gerar_todas_cartas_mensais()
    elif args.acao == "multa":
        gerar_todas_cartas_multa()
    elif args.acao == "emails":
        gerar_todos_emails()
    elif args.acao == "memorando":
        gerar_memorando()
    elif args.acao == "refis":
        gerar_todos_refis()
    elif args.acao == "lista":
        gerar_lista()
    elif args.acao == "tudo":
        executar_tudo()
    elif args.acao == "converter_pdf":
        converter_toda_saida()
    else:
        menu_interativo()


if __name__ == "__main__":
    main()

