let
    // Definindo a data de hoje e ontem
    hoje = DateTime.LocalNow(),
    ontem = Date.AddDays(hoje, -1),
    
    // Formatação da data de ontem e hoje
    data_ontem = Text.From(Date.Day(ontem)) & "-" & Text.PadStart(Text.From(Date.Month(ontem)), 2, "0") & "-" & Text.From(Date.Year(ontem)),

    // Construção da URL correta
    url_ontem = Text.Format("http://extrator.sefaz.al.gov.br/DESPESAS/COMPARATIVO-DOTACOES/CONSOLIDADO/comparativo_dotacao_despesa_consolidado_2018-2025_siafe_gerado_em_#[data].csv", [data = data_ontem]),

    // Tentativa de carregamento com a data de ontem e fallback para hoje
    Fonte_final =  Csv.Document(Web.Contents(url_ontem), [Delimiter=";", Columns=54, Encoding=1252, QuoteStyle=QuoteStyle.None])


in
    Fonte_final
