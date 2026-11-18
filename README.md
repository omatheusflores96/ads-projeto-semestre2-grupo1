# [ADS2] Projeto Integrador - Big Data com Python em Empresa Parceira

## Resumo do Projeto
Sistema de apoio à decisão para a **Bolos da Lyz**, confeitaria artesanal (MEI). O sistema substitui o caderno de anotações por um cadastro de clientes em Python e emite alertas de aniversário **30 e 60 dias antes da data**, para a confeiteira contatar o cliente a tempo e oferecer bolos e doces (fidelização e recompra).

## Integrantes e Responsabilidades
| Nome Completo | Matrícula | Usuário GitHub | Função no Projeto |
| :--- | :--- | :--- | :--- |
| Matheus de Almeida Flores | 202502541831 | @omatheusflores96 | Desenvolvedor |
| Igor Esteferson Vacario | 202503944202 | @IgorVacario | Desenvolvedor |
| Klaydiana da Silva Medeiros | 202502264811 | @daianamedeirossilva86-svg | Gerente de Projeto |
| Maria Eloisa de Oliveira Freitas | 202408452969 | @Maria20241413 | Analista de Negócio |

## Funcionalidades
* Cadastro, edição e remoção de clientes (nome, telefone e dia/mês do aniversário), com validação de telefone, de datas e de duplicidade.
* Alertas por janela de tempo: **30 dias** (urgente) e **60 dias** (atenção).
* Mensagem pronta e botão de WhatsApp (link `wa.me`) para cada cliente em alerta.
* Registro de contatos e pedidos e cálculo da **taxa de conversão**.
* Exportação da lista de clientes com os alertas em Excel.
* Painel no Power BI com os indicadores e a tabela de alertas.

## Stack Tecnológica
* **Linguagem:** Python 3.10+
* **Interface:** Streamlit
* **Tratamento de dados:** Pandas e openpyxl
* **Armazenamento:** Excel (`clientes.xlsx`, com as abas `clientes` e `contatos`)
* **Dashboard:** Power BI Desktop
* **Gestão e Versão:** Git, GitHub, Trello e Teams

## Arquivos do repositório
| Arquivo | Função |
| :--- | :--- |
| `app.py` | Interface web (cadastro, alertas e lista de clientes) |
| `logica.py` | Regras de negócio: validações, Excel, cálculo de alertas e conversão |
| `teste_logica.py` | Testes automáticos das regras |
| `powerquery_clientes.m`, `powerquery_contatos.m` | Consultas do Power BI que leem o Excel |
| `medidas_dax.txt` | Medidas do Power BI (taxa de conversão etc.) |
| `requirements.txt` | Bibliotecas necessárias |

## Como executar localmente
1. Clone o repositório:
   `git clone https://github.com/omatheusflores96/anuncios-projeto-semestre2-grupo1.git`
2. Acesse a branch de desenvolvimento:
   `git checkout develop`
3. Instale as dependências:
   `pip install -r requirements.txt`
4. Rode o sistema (abre no navegador em `http://localhost:8501`):
   `streamlit run app.py`
5. (Opcional) Rode os testes:
   `python teste_logica.py`

O arquivo `clientes.xlsx` é criado automaticamente no primeiro uso.

## Como conectar o Power BI
1. Abra o Power BI Desktop e crie as consultas com os arquivos `powerquery_contatos.m` (nome: **Contatos**) e `powerquery_clientes.m` (nome: **Clientes**).
2. Troque o caminho do arquivo dentro dos `.m` pelo local do seu `clientes.xlsx`.
3. Crie as medidas do `medidas_dax.txt`.
4. Sempre que houver dados novos, clique em **Atualizar** (o Power BI Desktop não atualiza sozinho).

## Observações
* O `clientes.xlsx` contém dados pessoais (nomes e telefones). Não o envie para o GitHub: inclua `clientes.xlsx` no `.gitignore`.
* Se o Excel estiver com o arquivo aberto, o app avisa para fechá-lo antes de salvar.
* Evoluções previstas: migração para banco de dados (SQLite), atualização automática do Power BI e exportação em CSV.