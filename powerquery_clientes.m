// CRIE ESTA CONSULTA DEPOIS da consulta "Contatos" (powerquery_contatos.m) e dê a ela o nome: Clientes
// Power BI: Obter dados > Consulta em branco > Editor Avançado > colar este código.
// Troque o caminho do arquivo pelo local real do clientes.xlsx.
let
    // Abre o arquivo Excel (TROQUE o caminho pelo local real do seu clientes.xlsx)
    Fonte = Excel.Workbook(File.Contents("C:\BolosDaLyz\clientes.xlsx"), null, true),
    // Seleciona a aba chamada "clientes"
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
                // Monta o aniversário em um ano; 29/02 em ano não bissexto vira 28/02
                f = (ano as number) as date =>
                    if _[mes] = 2 and _[dia] = 29 and not Date.IsLeapYear(ano)
                    then #date(ano, 2, 28)
                    else #date(ano, _[mes], _[dia]),
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
    ),
    // Nova coluna: situação do contato NESTE aniversário (procura na consulta "Contatos")
    Contato = Table.AddColumn(
        Status,
        "contato",
        each
            let
                ano = Date.Year([proximo_aniversario]),
                // Contatos deste cliente feitos para o aniversário deste ano
                achou = Table.SelectRows(Contatos, (c) => c[telefone] = _[telefone] and c[ano_aniversario] = ano)
            in
                if Table.IsEmpty(achou) then "Não contatado"
                else if achou{0}[pedido_confirmado] = "Sim" then "Pedido confirmado"
                else if achou{0}[pedido_confirmado] = "Não" then "Sem pedido"
                else "Aguardando resposta",
        type text
    )
in
    Contato