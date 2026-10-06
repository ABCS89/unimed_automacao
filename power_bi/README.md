# Guia Passo a Passo: Dashboard de Devedores Unimed no Power BI

Este guia ensina como conectar o Power BI Desktop à planilha `devedores.xlsx` e montar o painel com a listagem completa dos devedores, quantidades de boletos e valores devidos.

---

## 🚀 Passo 1: Importar os Dados com a Consulta Pronta (Menos de 1 minuto)

1. Abra o **Power BI Desktop**.
2. Na barra superior (guia **Página Inicial**), clique no botão **Transformar dados** (ícone de tabela com engrenagem).
   - Isso abrirá a janela do **Editor do Power Query**.
3. No menu à esquerda (em cima da lista de consultas), clique com o botão direito e escolha:
   **Nova Consulta** ➔ **Consulta em Branco**.
4. Na barra de ferramentas superior, clique no botão **Editor Avançado**.
5. Apague tudo o que estiver lá dentro e **cole** o conteúdo do arquivo:
   `unimed_automacao/power_bi/consulta_power_query.m`
6. Clique em **Concluído**.
7. No painel à esquerda, renomeie essa consulta para: `Base_Devedores_Unimed`.
8. Clique no botão **Fechar e Aplicar** (no canto superior esquerdo).

Pronto! Todas as abas (`Inadimplentes`, `Cancelados` e `Desligados`) foram unificadas, tratadas e carregadas automaticamente no modelo!

---

## 📊 Passo 2: Criar as Medidas DAX Principais

No Power BI Desktop, com a tabela `Base_Devedores_Unimed` selecionada à direita:
1. Clique em **Nova Medida** na barra superior e cole cada uma das fórmulas do arquivo `medidas_dax.txt`:
   * **Total Atualizado:** `Total Atualizado = SUM('Base_Devedores_Unimed'[Saldo (Atualizado)])`
   * **Nº de Boletos:** `Nº de Boletos = DISTINCTCOUNT('Base_Devedores_Unimed'[#])`
   * **Total de Devedores:** `Total de Devedores = DISTINCTCOUNT('Base_Devedores_Unimed'[Nome])`
2. Formate as medidas de valor como Moeda (R$) na guia superior de formatação.

---

## 🎨 Passo 3: Montar a Listagem com Visual Profissional

### 1. A Tabela de Inadimplentes (Listagem Completa)
1. No painel **Visualizações**, clique no ícone de **Tabela** ou **Matriz**.
2. Arraste para a tabela os seguintes campos:
   * `Nome`
   * `Status do Servidor` *(Ativo ou Desligado)*
   * `Status do Plano` *(Inadimplente (Aviso) ou Cancelado)*
   * `Sit. Vínculo` *(Auxílio Doença, Em Folha, Afastamento sem Vencimentos, etc.)*
   * Medida `[Nº de Boletos]`
   * Medida `[Total Atualizado]`
3. **Ativar Barras de Dados Coloridas:**
   * Com a tabela selecionada, vá no painel de **Formatação Visual** (ícone de pincel).
   * Abra a seção **Elementos da Célula**.
   * Selecione a série `Total Atualizado` e marque **Barra de dados**.
   * Agora, cada linha terá uma barra visual elegante proporcional ao tamanho da dívida!
4. **Ordenação:**
   * Clique no cabeçalho `Total Atualizado` para ordenar automaticamente do maior para o menor.

---

### 2. Cartões de Indicadores no Topo (KPIs)
Adicione 3 visuais de **Cartão** no topo do relatório:
* **Cartão 1:** Medida `[Total Atualizado]` *(exibe o valor total da dívida acumulada)*
  * *Dica de formatação*: selecione o cartão > vá no pincel de Formatação > **Valor do balão** > em **Exibir unidades**, mude de "Automático" para **Nenhum** (ou deixe 2 casas decimais) para exibir o valor monetário formatado.
* **Cartão 2:** Medida `[Total de Devedores]` *(exibe a contagem distinta de servidores)*
* **Cartão 3:** Medida `[Nº de Boletos]` *(exibe a quantidade de boletos em aberto)*

---

### 3. Segmentação de Dados (Filtros Clicáveis)
Adicione 3 visuais de **Segmentação de Dados (Filtro)**:
* **Filtro 1:** Campo `Status do Servidor` (*Ativo* ou *Desligado*).
* **Filtro 2:** Campo `Status do Plano` (*Inadimplente (Aviso)* ou *Cancelado*).
* **Filtro 3:** Campo `Sit. Vínculo` (*Auxílio Doença*, *Em Folha*, *Afastamento*, etc.).

---

## 🔄 Como Atualizar no Próximo Mês

Sempre que você editar, colar novas parcelas ou substituir a planilha [devedores.xlsx](file:///c:/Users/abusilva/Desktop/Github/unimed_automacao/entrada/devedores.xlsx):
1. Basta abrir o relatório no Power BI.
2. Clicar no botão **Atualizar** (na guia Página Inicial).
3. Todo o painel, a listagem, as quantidades de boletos e os totais se recalculam sozinhos em 2 segundos!
