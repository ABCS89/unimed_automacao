"""
memorando.py - Geração do memorando consolidado do Plano de Saúde
"""
from datetime import datetime
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate

from .config import (
    ARQUIVO_BASE,
    TEMPLATE_MEMORANDO,
    MEMORANDO_DIR,
)
from .utils import (
    limpa,
    capitalizar_nome,
    formatar_valor_br,
    data_por_extenso,
    ultimo_dia_util_do_mes,
)


def gerar_memorando():
    """Gera o arquivo de memorando consolidado em docx."""
    print("\n" + "=" * 60)
    print(">>> INICIANDO GERACAO DE MEMORANDO")
    print("=" * 60)

    MEMORANDO_DIR.mkdir(parents=True, exist_ok=True)

    if not ARQUIVO_BASE.exists() or not TEMPLATE_MEMORANDO.exists():
        print("[ERRO] Arquivo base (teste.ods) ou template de memorando ausente.")
        return

    df = pd.read_excel(ARQUIVO_BASE, engine="odf")

    hoje = datetime.today()
    data_venc = ultimo_dia_util_do_mes(hoje.year, hoje.month)

    pessoas = []
    ignorados = 0
    guia_num = 1

    for _, row in df.iterrows():
        condicao = limpa(row.get("condição")).lower()
        if "não enviar" in condicao or "nao enviar" in condicao:
            ignorados += 1
            continue

        nome = limpa(row.get("Funcionário"))
        if not nome:
            continue

        endereco_rua = limpa(row.get("endereço"))
        complemento = limpa(row.get("complemento"))
        if complemento:
            endereco_rua += f" – {complemento}"

        bairro = limpa(row.get("bairro"))
        cidade = limpa(row.get("cidade")) or "PIRACICABA"
        uf = limpa(row.get("uf")) or "SP"

        cpf = limpa(row.get("cpf"))

        dt_nasc = row.get("data_nascimento")
        if pd.notna(dt_nasc):
            try:
                data_nasc_str = pd.to_datetime(dt_nasc).strftime("%d/%m/%Y")
            except Exception:
                data_nasc_str = str(dt_nasc)
        else:
            data_nasc_str = ""

        # Monta discriminação de valores (Mensalidade e Coparticipação)
        valores = []
        try:
            val_mensalidade = float(row.get("Mensalidade", 0) or 0)
            if val_mensalidade > 0:
                valores.append(f"Mensalidade: R$ {formatar_valor_br(val_mensalidade)}")
        except (ValueError, TypeError):
            pass

        try:
            val_copart = float(row.get("Coparticipação", 0) or 0)
            if val_copart > 0:
                valores.append(f"Coparticipação: R$ {formatar_valor_br(val_copart)}")
        except (ValueError, TypeError):
            pass

        if not valores:
            valores.append(f"Mensalidade: R$ {formatar_valor_br(row.get('Total', 0))}")

        val_total = float(row.get("Total", 0) or 0)

        pessoas.append({
            "guia": guia_num,
            "nome": capitalizar_nome(nome),
            "cpf": cpf,
            "data_nascimento": data_nasc_str,
            "endereco": capitalizar_nome(endereco_rua),
            "bairro": capitalizar_nome(bairro),
            "cidade": capitalizar_nome(cidade),
            "uf": uf.upper(),
            "valores": "\n".join(valores),
            "total": f"R$ {formatar_valor_br(val_total)}",
        })
        guia_num += 1

    contexto = {
        "data_atual": data_por_extenso(hoje),
        "data_vencimento": data_venc.strftime("%d/%m/%Y"),
        "pessoas": pessoas,
    }

    doc = DocxTemplate(TEMPLATE_MEMORANDO)
    doc.render(contexto)

    nome_saida = f"memorando_unimed_{hoje.strftime('%m-%Y')}.docx"
    caminho_saida = MEMORANDO_DIR / nome_saida
    doc.save(caminho_saida)

    print(f"\n[OK] Memorando gerado com sucesso: {caminho_saida.name}")
    print(f"   * Total de servidores listados: {len(pessoas)}")
    if ignorados:
        print(f"   * Ignorados ('não enviar'): {ignorados}")
