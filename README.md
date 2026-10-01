# Automação Unimed — DRH

Módulo unificado e profissional para gestão e processamento mensal do convênio Unimed no Departamento de Recursos Humanos.

Este projeto consolida e substitui com melhorias todos os scripts avulsos anteriores, centralizando em um único lugar a geração de cartas de notificação, cartas de cobrança de multa, modelos de e-mails com tabelas de débitos, memorandos oficiais, termos de Refis e listagens de servidores.

---

## 📁 Estrutura de Pastas

```text
unimed_automacao/
├── README.md                     # Manual completo de uso
├── requirements.txt              # Dependências Python unificadas
├── main.py                       # Ponto de entrada (Menu interativo e CLI)
│
├── entrada/                      # Coloque aqui os arquivos do mês
│   ├── teste.ods                 # Planilha base dos servidores e mensalidades
│   ├── devedores.xlsx            # Planilha de débitos (abas Inadimplentes e Cancelados)
│   └── pdfs_envio/               # Comprovantes de envio de e-mail (PDFs)
│
├── templates/                    # Modelos oficiais de documentos e e-mails
│   ├── template_base.docx        # Carta mensal para ativos
│   ├── template_desligado.docx   # Carta mensal para desligados
│   ├── template_aviso.docx       # Carta com tabela de débitos (aviso de cancelamento)
│   ├── template_cancelado.docx   # Carta com tabela de débitos (já cancelados)
│   ├── template_base_multa.docx  # Carta especial de cobrança de multa por atraso
│   ├── template_memorando.docx   # Memorando oficial consolidado
│   ├── template_refis.docx       # Termo de acordo e confissão de dívida Refis
│   ├── template_lista.docx       # Listagem consolidada de servidores
│   └── emails/                   # Templates de texto para e-mails
│       ├── email_normal.txt
│       ├── email_desligados.txt
│       ├── email_aviso.txt
│       └── email_cancelado.txt
│
├── saida/                        # Gerado automaticamente a cada execução
│   ├── cartas/
│   │   ├── base/                 # Cartas base e de desligados (.docx)
│   │   ├── cancelados_aviso/     # Cartas de aviso e cancelados com tabelas (.docx)
│   │   └── multa/                # Cartas de cobrança de multa (.docx)
│   ├── emails/                   # Arquivos Markdown com modelos prontos para envio
│   │   ├── emails_normais.md
│   │   ├── emails_desligados.md
│   │   ├── emails_aviso.md
│   │   └── emails_cancelados.md
│   ├── memorando/                # Memorando mensal gerado (.docx)
│   ├── refis/                    # Termos de acordo do Refis (.docx)
│   └── listas/                   # Lista consolidada de servidores (.docx)
│
└── src/                          # Código-fonte modularizado
    ├── __init__.py
    ├── config.py                 # Caminhos absolutos e configurações
    ├── utils.py                  # Funções de data, moeda, num2words e texto
    ├── cartas.py                 # Rotinas de geração de cartas mensais e multas
    ├── emails.py                 # Rotinas de geração dos modelos de e-mail
    ├── memorando.py              # Rotina de geração do memorando
    ├── refis.py                  # Rotina de geração de acordos Refis
    ├── lista.py                  # Rotina de listagem de servidores
    └── converter_pdf.py          # Conversão em lote de DOCX para PDF (Pós-validação)
```

---

## ⚙️ Configuração Inicial (Apenas na 1ª vez)

No terminal, acesse a pasta `unimed_automacao/`:

```bash
# 1. Criar o ambiente virtual
python -m venv venv

# 2. Ativar o ambiente virtual
# No Windows:
venv\Scripts\activate
# No Linux/Mac:
source venv/bin/activate

# 3. Instalar as dependências
pip install -r requirements.txt
```

---

## 🚀 Como Usar no Dia a Dia

O fluxo de trabalho foi desenhado para respeitar o seu processo de **validação manual**:
1. Você executa a geração dos arquivos `.docx`.
2. Abre a pasta `saida/`, confere e ajusta o que for necessário nos documentos Word.
3. Executa a etapa de **Conversão para PDF** para gerar todos os PDFs automaticamente em segundos.

### Modo 1: Menu Interativo no Terminal (Recomendado)

Basta rodar o comando principal:

```bash
python main.py
```

Será exibido o menu interativo:

```text
============================================================
    AUTOMACAO UNIMED - DEPARTAMENTO DE RECURSOS HUMANOS
============================================================
  1. [CARTAS] Gerar Cartas Mensais (Base, Desligados, Avisos, Cancelados)
  2. [MULTA] Gerar Cartas de Multa por Atraso
  3. [EMAILS] Gerar Modelos de E-mails com Tabelas de Debitos
  4. [MEMORANDO] Gerar Memorando Geral em Word
  5. [REFIS] Gerar Termos de Acordo do Refis
  6. [LISTA] Gerar Lista Consolidada de Servidores
  7. [TUDO] Executar TUDO (Rotina Mensal de Geracao DOCX)
  8. [CONVERTER PDF] Converter DOCX para PDF (Pos-Validacao)
  0. [SAIR] Sair
============================================================
Digite a opcao desejada [0-8]:
```

Ao escolher a opção **`8`**, um submenu permite converter pastas específicas (só cartas, só memorando, etc.) ou converter toda a pasta `saida/` de uma única vez.

### Modo 2: Linha de Comando Direta (Para automações ou atalhos)

Você pode passar a ação desejada diretamente via parâmetro `--acao`:

```bash
# Para gerar todos os documentos Word:
python main.py --acao tudo

# Para converter todos os documentos gerados para PDF após sua validação:
python main.py --acao converter_pdf

# Ou executar tarefas específicas:
python main.py --acao cartas
python main.py --acao multa
python main.py --acao emails
python main.py --acao memorando
python main.py --acao refis
python main.py --acao lista
```

---

## 📄 Conversão para PDF (Como Funciona)

- A conversão utiliza o motor nativo do **LibreOffice** (`soffice.exe`) em segundo plano (*headless*), que é ultra-rápido (converte dezenas de arquivos em poucos segundos) e mantém 100% da fidelidade visual dos cabeçalhos, rodapés e tabelas.
- Os PDFs são salvos diretamente ao lado dos seus respectivos arquivos `.docx` na pasta `saida/`.
- Arquivos temporários do Word (iniciados por `~$`) são ignorados automaticamente para evitar falhas caso você esteja com algum documento aberto no momento da conversão.

---

## 🔍 Regras de Negócio e Comportamento

1. **Cálculo de Datas:**
   - O vencimento padrão é calculado automaticamente como o **último dia útil do mês atual**, considerando fins de semana e feriados estaduais de SP (usando a biblioteca `holidays`).
   - A competência de referência padrão assume o mês anterior ao atual.
2. **Vinculação de E-mails (PDFs de envio):**
   - Ao gerar as cartas mensais, o script busca em `entrada/pdfs_envio/` se existe um PDF com o nome do servidor para extrair a data real em que o e-mail foi disparado. Caso não encontre, insere placeholders no documento Word para fácil identificação e preenchimento manual.
3. **Cartas de Multa:**
   - Analisa a aba `Inadimplentes` da planilha `devedores.xlsx`.
   - Identifica parcelas onde o **`Saldo (Atualizado)` é positivo e menor do que a coluna `Principal (Saldo)`**. Isso ocorre quando o servidor pagou o boleto com atraso e restou pendente apenas a cobrança da diferença/juros, que é inserida para quitação conjunta com a guia do mês.
4. **Modelos de E-mail:**
   - Classifica os servidores em quatro categorias: *Normais*, *Desligados*, *Avisos* e *Cancelados*.
   - Para as categorias com dívidas (*Avisos* e *Cancelados*), gera uma tabela formatada em Markdown detalhando competência, vencimento, valor principal, encargos e total.
