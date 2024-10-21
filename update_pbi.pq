let
    // Definindo a data de hoje e ontem
    hoje = DateTime.LocalNow(),
    ontem = Date.AddDays(hoje, -1),
    
    // Formatação da data de ontem e hoje
    data_ontem = Text.From(Date.Day(ontem)) & "-" & Text.PadStart(Text.From(Date.Month(ontem)), 2, "0") & "-" & Text.From(Date.Year(ontem)),
    data_hoje = Text.From(Date.Day(hoje)) & "-" & Text.PadStart(Text.From(Date.Month(hoje)), 2, "0") & "-" & Text.From(Date.Year(hoje)),

    // Construção da URL para ontem
    url_ontem = Text.Format("http://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DOTACOES/CONSOLIDADO/comparativo_dotacao_despesa_consolidado_2018-2024_siafe_gerado_em_#[data].csv", [data = data_ontem]),

    // Tentativa de carregamento com a data de ontem
    Fonte_ontem = try Csv.Document(Web.Contents(url_ontem), [Delimiter=";", Columns=54, Encoding=1252, QuoteStyle=QuoteStyle.None]),

    // Caso falhe, tenta com a data de hoje
    Fonte_final = if Fonte_ontem[HasError] then
        let
            // Construção da URL para hoje
            url_hoje = Text.Format("http://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DOTACOES/CONSOLIDADO/comparativo_dotacao_despesa_consolidado_2018-2024_siafe_gerado_em_#[data].csv", [data = data_hoje]),
            
            // Tentativa de carregamento com a data de hoje
            Fonte_hoje = Csv.Document(Web.Contents(url_hoje), [Delimiter=";", Columns=54, Encoding=1252, QuoteStyle=QuoteStyle.None])
        in
            Fonte_hoje
    else
        Fonte_ontem[Value]
in
    Fonte_final