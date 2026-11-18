// CRIE ESTA CONSULTA PRIMEIRO e dê a ela o nome: Contatos
// Power BI: Obter dados > Consulta em branco > Editor Avançado > colar este código.
// Troque o caminho do arquivo pelo local real do clientes.xlsx.
let
    // Abre o arquivo Excel (TROQUE o caminho pelo local real do seu clientes.xlsx)
    Fonte = Excel.Workbook(File.Contents("C:\BolosDaLyz\clientes.xlsx"), null, true),
    // Seleciona a aba "contatos" (histórico de contatos e pedidos)
    Aba = Fonte{[Item = "contatos", Kind = "Sheet"]}[Data],
    // Transforma a primeira linha em cabeçalho
    Cabecalho = Table.PromoteHeaders(Aba, [PromoteAllScalars = true]),
    // Define o tipo de cada coluna
    Tipos = Table.TransformColumnTypes(
        Cabecalho,
        {
            {"telefone", type text},
            {"nome", type text},
            {"data_contato", type text},
            {"ano_aniversario", Int64.Type},
            {"pedido_confirmado", type text}
        }
    ),
    // A data é guardada como texto AAAA-MM-DD; aqui vira data de verdade
    Datas = Table.TransformColumns(Tipos, {{"data_contato", each Date.FromText(_), type date}})
in
    Datas
