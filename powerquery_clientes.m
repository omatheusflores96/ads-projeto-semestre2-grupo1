// Power BI: Obter dados > Consulta em branco > Editor Avançado > colar este código.
// Troque o caminho do arquivo pelo local real do clientes.xlsx.
let
    // Abre o arquivo Excel (TROQUE o caminho pelo local real do seu clientes.xlsx)
    Fonte = Excel.Workbook(File.Contents("C:\BolosDaLyz\clientes.xlsx"), null, true),
    // Seleciona a aba chamada "clientes" (mesmo nome usado pelo Python)
    Aba = Fonte{[Item = "clientes", Kind = "Sheet"]}[Data],
    // Transforma a primeira linha em cabeçalho (nome, telefone, dia, mes)
    Cabecalho = Table.PromoteHeaders(Aba, [PromoteAllScalars = true]),
    // Define o tipo de cada coluna: texto para nome/telefone, número inteiro para dia/mes
    Tipos = Table.TransformColumnTypes(
        Cabecalho,
        {{"nome", type text}, {"telefone", type text}, {"dia", Int64.Type}, {"mes", Int64.Type}}
    ),
    // Data de hoje, recalculada a cada atualização do Power BI
    Hoje = Date.From(DateTime.LocalNow()),
    // Nova coluna: data do próximo aniversário de cada cliente
    ProxAniv = Table.AddColumn(
        Tipos,
        "proximo_aniversario",
        each
            let
                // Função que monta o aniversário em um ano; 29/02 em ano não bissexto vira 28/02
                f = (ano as number) as date =>
                    if _[mes] = 2 and _[dia] = 29 and not Date.IsLeapYear(ano)
                    then #date(ano, 2, 28)
                    else #date(ano, _[mes], _[dia]),
                // Aniversário no ano atual
                esteAno = f(Date.Year(Hoje))
            in
                // Se ainda não passou (ou é hoje) usa este ano; senão, o ano seguinte
                if esteAno >= Hoje then esteAno else f(Date.Year(Hoje) + 1),
        type date
    ),
    // Nova coluna: quantos dias faltam até o próximo aniversário
    Dias = Table.AddColumn(
        ProxAniv, "dias_restantes", each Duration.Days([proximo_aniversario] - Hoje), Int64.Type
    ),
    // Nova coluna: nível do alerta (usada na formatação condicional de cores)
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
    // Resultado final: a tabela completa, com as 3 colunas calculadas
    Status