let
    // 1. Caminho dinâmico para a planilha devedores.xlsx
    CaminhoArquivo = "C:\Users\abusilva\Desktop\Github\unimed_automacao\entrada\devedores.xlsx",
    Fonte = Excel.Workbook(File.Contents(CaminhoArquivo), null, true),

    // 2. Extrair aba Inadimplentes (Servidor Ativo / Plano com Aviso de Inadimplência)
    Inad_Sheet = Fonte{[Item="Inadimplentes", Kind="Sheet"]}[Data],
    Inad_Headers = Table.PromoteHeaders(Inad_Sheet, [PromoteAllScalars=true]),
    Inad_Cols = Table.AddColumn(Inad_Headers, "Status do Plano", each "Inadimplente (Aviso)", type text),
    Inad_Final = Table.AddColumn(Inad_Cols, "Status do Servidor", each "Ativo", type text),

    // 3. Extrair aba Cancelados (Servidor Ativo / Plano Cancelado por Inadimplência)
    Canc_Sheet = Fonte{[Item="Cancelados", Kind="Sheet"]}[Data],
    Canc_Headers = Table.PromoteHeaders(Canc_Sheet, [PromoteAllScalars=true]),
    Canc_Cols = Table.AddColumn(Canc_Headers, "Status do Plano", each "Cancelado", type text),
    Canc_Final = Table.AddColumn(Canc_Cols, "Status do Servidor", each "Ativo", type text),

    // 4. Extrair aba Desligados (Servidor Desligado / Plano Cancelado)
    Desl_Sheet = Fonte{[Item="Desligados", Kind="Sheet"]}[Data],
    Desl_Headers = Table.PromoteHeaders(Desl_Sheet, [PromoteAllScalars=true]),
    Desl_Cols = Table.AddColumn(Desl_Headers, "Status do Plano", each "Cancelado", type text),
    Desl_Final = Table.AddColumn(Desl_Cols, "Status do Servidor", each "Desligado", type text),

    // 5. Unificar as 3 bases
    Tabelas_Unidas = Table.Combine({Inad_Final, Canc_Final, Desl_Final}),

    // 6. Filtrar linhas vazias ou de cabeçalho residual
    Filtrar_Validos = Table.SelectRows(Tabelas_Unidas, each ([Nome] <> null and [Nome] <> "" and [Nome] <> "Nome")),

    // 7. Padronização e Limpeza da Situação do Vínculo
    Tratar_Vinculo = Table.AddColumn(Filtrar_Validos, "Situação Vínculo Padronizada", each 
        let
            v = Text.Trim(Text.Upper(Text.From([Sit. Vinculo] ?? "")))
        in
            if v = "" or v = "NULL" then "Não Informado"
            else if Text.Contains(v, "AFAST") then "Afastamento sem Vencimentos"
            else if Text.Contains(v, "DOEN") then "Auxílio Doença"
            else if Text.Contains(v, "FOLHA") then "Em Folha"
            else if Text.Contains(v, "PER") or Text.Contains(v, "RETORNO") then "Retorno / Perícia"
            else if Text.Contains(v, "DESLIG") then "Desligado"
            else if Text.Contains(v, "SA") then "Tratamento de Saúde"
            else if Text.Contains(v, "COMISS") then "Comissão"
            else if Text.Contains(v, "CEDID") then "Cedido"
            else if Text.Contains(v, "FALTA") then "Falta"
            else if Text.Contains(v, "FUNCIONAL") then "Outro Funcional"
            else Text.Proper(Text.From([Sit. Vinculo])),
        type text
    ),

    // 8. Selecionar colunas finais organizadas
    Colunas_Finais = Table.SelectColumns(Tratar_Vinculo, {
        "Status do Servidor",
        "Status do Plano",
        "Situação Vínculo Padronizada",
        "Funcional",
        "Nome",
        "CPF",
        "#",
        "Mês/Ano",
        "Data de Vencimento",
        "Principal (Saldo)",
        "Saldo (Atualizado)",
        "Sit. Dívida"
    }),

    // 9. Renomear para exibição mais amigável
    Renomear = Table.RenameColumns(Colunas_Finais, {
        {"Situação Vínculo Padronizada", "Sit. Vínculo"}
    }),

    // 10. Tipagem correta dos dados
    Tipagem = Table.TransformColumnTypes(Renomear, {
        {"Status do Servidor", type text},
        {"Status do Plano", type text},
        {"Sit. Vínculo", type text},
        {"Funcional", type text},
        {"Nome", type text},
        {"CPF", type text},
        {"#", type text},
        {"Mês/Ano", type date},
        {"Data de Vencimento", type date},
        {"Principal (Saldo)", Currency.Type},
        {"Saldo (Atualizado)", Currency.Type},
        {"Sit. Dívida", type text}
    })
in
    Tipagem
