let
    // 1. Caminho dinâmico para a planilha devedores.xlsx
    CaminhoArquivo = "C:\Users\abusilva\Desktop\Github\unimed_automacao\entrada\devedores.xlsx",
    Fonte = Excel.Workbook(File.Contents(CaminhoArquivo), null, true),

    // 2. Extrair aba Inadimplentes
    Inad_Sheet = Fonte{[Item="Inadimplentes", Kind="Sheet"]}[Data],
    Inad_Headers = Table.PromoteHeaders(Inad_Sheet, [PromoteAllScalars=true]),
    Inad_Status = Table.AddColumn(Inad_Headers, "Status do Plano", each "Inadimplente (Aviso)", type text),

    // 3. Extrair aba Cancelados
    Canc_Sheet = Fonte{[Item="Cancelados", Kind="Sheet"]}[Data],
    Canc_Headers = Table.PromoteHeaders(Canc_Sheet, [PromoteAllScalars=true]),
    Canc_Status = Table.AddColumn(Canc_Headers, "Status do Plano", each "Cancelado", type text),

    // 4. Extrair aba Desligados
    Desl_Sheet = Fonte{[Item="Desligados", Kind="Sheet"]}[Data],
    Desl_Headers = Table.PromoteHeaders(Desl_Sheet, [PromoteAllScalars=true]),
    Desl_Status = Table.AddColumn(Desl_Headers, "Status do Plano", each "Desligado", type text),

    // 5. Unificar as 3 bases em uma tabela consolidada
    Tabelas_Unidas = Table.Combine({Inad_Status, Canc_Status, Desl_Status}),

    // 6. Filtrar linhas onde o Nome é nulo ou cabeçalho residual
    Filtrar_Validos = Table.SelectRows(Tabelas_Unidas, each ([Nome] <> null and [Nome] <> "" and [Nome] <> "Nome")),

    // 7. Selecionar e ordenar as colunas oficiais
    Colunas_Finais = Table.SelectColumns(Filtrar_Validos, {
        "Status do Plano",
        "Funcional",
        "Nome",
        "CPF",
        "#",
        "Mês/Ano",
        "Data de Vencimento",
        "Principal (Saldo)",
        "Saldo (Atualizado)",
        "Sit. Dívida",
        "Sit. Vinculo"
    }),

    // 8. Definir tipagem precisa de cada campo
    Tipagem = Table.TransformColumnTypes(Colunas_Finais, {
        {"Status do Plano", type text},
        {"Funcional", type text},
        {"Nome", type text},
        {"CPF", type text},
        {"#", type text},
        {"Mês/Ano", type date},
        {"Data de Vencimento", type date},
        {"Principal (Saldo)", Currency.Type},
        {"Saldo (Atualizado)", Currency.Type},
        {"Sit. Dívida", type text},
        {"Sit. Vinculo", type text}
    })
in
    Tipagem
