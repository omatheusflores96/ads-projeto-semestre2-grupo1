"""Regras de negócio do sistema de aniversários - Bolos da Lyz."""
# Bibliotecas padrão do Python
import re                  # expressões regulares (usada para limpar o telefone)
import urllib.parse        # codifica o texto da mensagem para virar parte de um link
from datetime import date  # trabalhar com datas
from pathlib import Path   # representar o caminho do arquivo de forma segura

# Biblioteca externa: Pandas, usada para ler/gravar o Excel e manipular a tabela
import pandas as pd

# ---------- Configurações do arquivo ----------
ARQUIVO = Path("clientes.xlsx")           # arquivo Excel onde ficam os clientes (na mesma pasta do app)
ABA = "clientes"                          # nome da aba dentro do Excel (o Power BI procura por este nome)
COLUNAS = ["nome", "telefone", "dia", "mes"]  # colunas da tabela de clientes


# ---------- Validações (RNF04) ----------
def limpar_telefone(telefone: str) -> str:
    """Deixa só os dígitos e remove o código do país (55), se houver."""
    # Remove tudo que não é número (espaços, parênteses, traços, "+")
    # O "or ''" evita erro caso o telefone venha vazio (None)
    digitos = re.sub(r"\D", "", telefone or "")
    # Se tem 12 ou 13 dígitos e começa com 55, é número com código do país: tira o 55
    if len(digitos) in (12, 13) and digitos.startswith("55"):
        digitos = digitos[2:]
    return digitos


def telefone_valido(telefone: str) -> bool:
    """DDD + número: 10 dígitos (fixo) ou 11 (celular)."""
    return len(telefone) in (10, 11)


def data_valida(dia: int, mes: int) -> bool:
    """Usa um ano bissexto para aceitar 29/02 e bloquear 31/02, 31/04 etc."""
    try:
        # 2024 é bissexto, então 29/02 é aceito; datas impossíveis geram ValueError
        date(2024, int(mes), int(dia))
        return True
    except (ValueError, TypeError):
        # ValueError: data inexistente | TypeError: valor vazio ou de tipo errado
        return False


# ---------- Leitura e gravação (RF02, RF04) ----------
def carregar() -> pd.DataFrame:
    # Se o arquivo ainda não existe (primeiro uso), devolve uma tabela vazia com as colunas certas
    if not ARQUIVO.exists():
        return pd.DataFrame(columns=COLUNAS)
    # dtype={"telefone": str} evita que o Excel apague zeros à esquerda do telefone
    df = pd.read_excel(ARQUIVO, sheet_name=ABA, dtype={"telefone": str})
    # Garante a ordem e o conjunto de colunas esperado
    return df[COLUNAS]


def salvar(df: pd.DataFrame) -> None:
    # Grava a tabela inteira no Excel, sem a coluna de índice (0, 1, 2...)
    df.to_excel(ARQUIVO, sheet_name=ABA, index=False)


def adicionar_cliente(nome: str, telefone: str, dia: int, mes: int):
    """Retorna (sucesso, mensagem)."""
    # Limpa o nome (tira espaços extras) e o telefone (só números)
    nome = (nome or "").strip()
    tel = limpar_telefone(telefone)

    # Validações: cada uma interrompe o cadastro e devolve a mensagem de erro
    if not nome:
        return False, "Informe o nome do cliente."
    if not telefone_valido(tel):
        return False, "Telefone inválido. Use DDD + número (ex.: 32 91234-5678)."
    if not data_valida(dia, mes):
        return False, "Data de aniversário inválida."

    # Carrega os clientes já cadastrados para checar duplicidade (o telefone é a chave)
    df = carregar()
    if tel in df["telefone"].values:
        return False, "Já existe um cliente com esse telefone."

    # Monta uma tabela de 1 linha com o novo cliente
    novo = pd.DataFrame([{"nome": nome, "telefone": tel, "dia": int(dia), "mes": int(mes)}])
    # Junta com os existentes e salva no Excel
    salvar(pd.concat([df, novo], ignore_index=True))
    return True, f"{nome} cadastrado(a) com sucesso!"


def remover_cliente(telefone: str) -> None:
    df = carregar()
    # Mantém todos os clientes, exceto o que tem o telefone informado
    salvar(df[df["telefone"] != telefone])


# ---------- Alertas (RF03, RF05, RF07) ----------
def proximo_aniversario(dia: int, mes: int, hoje: date | None = None) -> date:
    """Próxima data de aniversário a partir de hoje (trata 29/02 em ano não bissexto)."""
    # Se não passar a data de hoje, usa a data atual do computador (o parâmetro serve para testes)
    hoje = hoje or date.today()

    def data_no_ano(ano: int) -> date:
        # Monta o aniversário dentro de um ano específico
        try:
            return date(ano, mes, dia)
        except ValueError:  # 29/02 em ano não bissexto
            # Nesse caso, comemora-se em 28/02
            return date(ano, 2, 28)

    # Primeiro tenta o aniversário deste ano
    candidata = data_no_ano(hoje.year)
    # Se ainda não passou (ou é hoje), é ela; senão, o próximo é no ano seguinte
    return candidata if candidata >= hoje else data_no_ano(hoje.year + 1)


def classificar(dias: int) -> str:
    # Define o nível do alerta conforme quantos dias faltam
    if dias <= 30:
        return "URGENTE (30 dias)"
    if dias <= 60:
        return "ATENÇÃO (60 dias)"
    return "Fora da janela"


def tabela_alertas(df: pd.DataFrame, hoje: date | None = None) -> pd.DataFrame:
    hoje = hoje or date.today()
    df = df.copy()  # trabalha numa cópia para não alterar a tabela original
    # Sem clientes: devolve a tabela vazia, já com as colunas novas
    if df.empty:
        return df.assign(proximo_aniversario=[], dias_restantes=[], status=[])
    # Calcula a data do próximo aniversário de cada cliente
    df["proximo_aniversario"] = [
        proximo_aniversario(int(d), int(m), hoje) for d, m in zip(df["dia"], df["mes"])
    ]
    # Quantos dias faltam entre hoje e esse aniversário
    df["dias_restantes"] = [(p - hoje).days for p in df["proximo_aniversario"]]
    # Classifica cada cliente em URGENTE / ATENÇÃO / Fora da janela
    df["status"] = df["dias_restantes"].apply(classificar)
    # Ordena do aniversário mais próximo para o mais distante
    return df.sort_values("dias_restantes").reset_index(drop=True)


# ---------- Mensagem pronta (RF09) ----------
def link_whatsapp(nome: str, telefone: str, dia: int, mes: int) -> str:
    # Usa só o primeiro nome para a mensagem ficar mais natural
    primeiro_nome = nome.split()[0]
    # Texto padrão da mensagem, com nome e data do cliente
    texto = (
       f"Olá, {primeiro_nome}! Tudo bem por aí? \n"
       f"Aqui é da Bolos da Lyz! Vi que o seu aniversário está a aproximar-se no dia {int(dia):02d}/{int(mes):02d} e já estamos a preparar muita doçura por aqui. \n"
       f"Quer garantir o bolo e os docinhos para comemorar essa data especial com a gente? É só me avisar! "
    )
    # Link do WhatsApp: 55 = Brasil; quote() converte espaços e acentos para o formato de link
    return f"https://wa.me/55{telefone}?text={urllib.parse.quote(texto)}"