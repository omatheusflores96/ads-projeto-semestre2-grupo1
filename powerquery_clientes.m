// Power BI: Obter dados > Consulta em branco > Editor Avançado > colar este código.
// Troque o caminho do arquivo pelo local real do clientes.xlsx.
let
    Fonte = Excel.Workbook(File.Contents("C:\BolosDaLyz\clientes.xlsx"), null, true),
    Aba = Fonte{[Item = "clientes", Kind = "Sheet"]}[Data],
    Cabecalho = Table.PromoteHeaders(Aba, [PromoteAllScalars = true]),
    Tipos = Table.TransformColumnTypes(
        Cabecalho,
        {{"nome", type text}, {"telefone", type text}, {"dia", Int64.Type}, {"mes", Int64.Type}}
    ),
    Hoje = Date.From(DateTime.LocalNow()),
    ProxAniv = Table.AddColumn(
        Tipos,
        "proximo_aniversario",
        each
            let
                f = (ano as number) as date =>
                    if _[mes] = 2 and _[dia] = 29 and not Date.IsLeapYear(ano)
                    then #date(ano, 2, 28)
                    else #date(ano, _[mes], _[dia]),
                esteAno = f(Date.Year(Hoje))
            in
                if esteAno >= Hoje then esteAno else f(Date.Year(Hoje) + 1),
        type date
    ),
    Dias = Table.AddColumn(
        ProxAniv, "dias_restantes", each Duration.Days([proximo_aniversario] - Hoje), Int64.Type
    ),
    Status = Table.AddColumn(
        Dias,
        "status",
        each
            if [dias_restantes] <= 30 then "URGENTE (30 dias)"
            else if [dias_restantes] <= 60 then "ATENÇÃO (60 dias)"
            else "Fora da janela",
        type text
    )
in
    Status
