"""Regras de negócio do sistema de aniversários - Bolos da Lyz (versão 3)."""
# Bibliotecas padrão do Python
import calendar            # saber se um ano é bissexto
import io                  # cria um "arquivo na memória" (usado na exportação para Excel)
import logging             # registra os erros no terminal para facilitar a manutenção
import os                  # variáveis de ambiente e troca segura de arquivos
import re                  # expressões regulares (usada para limpar o telefone)
import urllib.parse        # codifica o texto da mensagem para virar parte de um link
from datetime import date  # trabalhar com datas
from pathlib import Path   # representar o caminho do arquivo de forma segura

# Biblioteca externa: Pandas, usada para ler/gravar o Excel e manipular as tabelas
import pandas as pd

logger = logging.getLogger("bolos_da_lyz")

# ---------- Configurações do arquivo ----------
# Caminho do Excel: pode ser trocado pela variável de ambiente CLIENTES_XLSX
ARQUIVO = Path(os.environ.get("CLIENTES_XLSX", "clientes.xlsx"))
ABA = "clientes"                  # aba com a lista de clientes
ABA_CONTATOS = "contatos"         # aba com o histórico de contatos e pedidos (RF08)
COLUNAS = ["nome", "telefone", "dia", "mes"]
# Cada linha da aba "contatos" é um contato feito pela confeiteira em um ano de aniversário:
# - ano_aniversario: ano do aniversário que motivou o contato (separa um ano do outro)
# - pedido_confirmado: "Aguardando", "Sim" ou "Não"
COLUNAS_CONTATOS = ["telefone", "nome", "data_contato", "ano_aniversario", "pedido_confirmado"]

# Texto mostrado no app/Excel para cada situação do contato
TEXTO_CONTATO = {
    "Aguardando": "Aguardando resposta",
    "Sim": "Pedido confirmado",
    "Não": "Sem pedido",
}


class ErroDeDados(Exception):
    """Problema com o arquivo de dados (corrompido, aba ou coluna ausente).

    A mensagem já é amigável: o app mostra o texto direto para a confeiteira.
    """


# ---------- Validações (RNF04) ----------
def limpar_telefone(telefone: str) -> str:
    """Deixa só os dígitos e remove o código do país (55), se houver."""
    # Remove tudo que não é número (espaços, parênteses, traços, "+")
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
        date(2024, int(mes), int(dia))
        return True
    except (ValueError, TypeError):
        return False


def _validar_dados(nome: str, tel: str, dia: int, mes: int):
    """Devolve a mensagem de erro, ou None se estiver tudo certo (usada no cadastro e na edição)."""
    if not nome:
        return "Informe o nome do cliente."
    if not telefone_valido(tel):
        return "Telefone inválido. Use DDD + número (ex.: 32 91234-5678)."
    if not data_valida(dia, mes):
        return "Data de aniversário inválida."
    return None


def separar_validas(df: pd.DataFrame):
    """Separa as linhas válidas das inválidas da lista de clientes.

    Serve para o caso de dados digitados direto na planilha, sem passar pelas validações
    do formulário. Retorna (validas, invalidas):
    - validas: linhas já normalizadas (nome sem espaços extras, telefone só com dígitos, dia/mês inteiros);
    - invalidas: as linhas originais, com a coluna "motivo".
    """
    nome = df["nome"].fillna("").astype(str).str.strip()
    # O "[.]0$" remove o ".0" que aparece quando o Excel guarda o telefone como número decimal
    tel = df["telefone"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).map(limpar_telefone)
    dia = pd.to_numeric(df["dia"], errors="coerce")
    mes = pd.to_numeric(df["mes"], errors="coerce")

    motivos = []
    vistos = set()  # telefones já vistos, para achar duplicados
    for i in range(len(df)):
        if not nome.iloc[i]:
            motivo = "nome vazio"
        elif not telefone_valido(tel.iloc[i]):
            motivo = "telefone inválido"
        elif pd.isna(dia.iloc[i]) or pd.isna(mes.iloc[i]) or not data_valida(dia.iloc[i], mes.iloc[i]):
            motivo = "data de aniversário inválida"
        elif tel.iloc[i] in vistos:
            motivo = "telefone duplicado"
        else:
            motivo = ""
            vistos.add(tel.iloc[i])
        motivos.append(motivo)

    motivos = pd.Series(motivos, index=df.index, dtype=object)
    ok = motivos == ""

    validas = pd.DataFrame({
        "nome": nome[ok],
        "telefone": tel[ok],
        "dia": dia[ok].astype(int),
        "mes": mes[ok].astype(int),
    }).reset_index(drop=True)
    invalidas = df[~ok].assign(motivo=motivos[~ok]).reset_index(drop=True)
    return validas, invalidas


# ---------- Leitura e gravação (RF02, RF04) ----------
def _abas_existentes() -> list:
    """Nomes das abas do arquivo (lista vazia se o arquivo ainda não existe)."""
    if not ARQUIVO.exists():
        return []
    try:
        with pd.ExcelFile(ARQUIVO) as xls:
            return list(xls.sheet_names)
    except PermissionError:
        raise  # o app avisa para fechar o Excel
    except Exception as erro:  # arquivo corrompido, vazio ou que não é um Excel de verdade
        logger.exception("Falha ao abrir %s", ARQUIVO)
        raise ErroDeDados(
            f"Não foi possível abrir o arquivo {ARQUIVO.name}. "
            "Ele pode estar corrompido ou não ser um Excel válido. Confira se abre no Excel."
        ) from erro


def _ler_aba(aba: str, colunas: list) -> pd.DataFrame:
    """Lê uma aba do Excel. Se o arquivo ou a aba não existir, devolve uma tabela vazia."""
    if aba not in _abas_existentes():
        return pd.DataFrame(columns=colunas)
    try:
        # dtype={"telefone": str} evita que o Excel apague zeros à esquerda do telefone
        df = pd.read_excel(ARQUIVO, sheet_name=aba, dtype={"telefone": str})
    except PermissionError:
        raise
    except Exception as erro:
        logger.exception("Falha ao ler a aba %s", aba)
        raise ErroDeDados(f"Não foi possível ler a aba '{aba}' do arquivo {ARQUIVO.name}.") from erro

    faltando = [c for c in colunas if c not in df.columns]
    if faltando:
        raise ErroDeDados(
            f"A aba '{aba}' do arquivo {ARQUIVO.name} não tem a(s) coluna(s): {', '.join(faltando)}."
        )
    return df[colunas]


def carregar() -> pd.DataFrame:
    """Lista de clientes (como está na planilha, sem corrigir nada)."""
    return _ler_aba(ABA, COLUNAS)


def carregar_contatos() -> pd.DataFrame:
    """Histórico de contatos e pedidos."""
    df = _ler_aba(ABA_CONTATOS, COLUNAS_CONTATOS)
    # Número inválido vira 0 em vez de derrubar o sistema
    df["ano_aniversario"] = pd.to_numeric(df["ano_aniversario"], errors="coerce").fillna(0).astype(int)
    # Garante que a data fique como texto AAAA-MM-DD (evita o Excel trocar o formato)
    df["data_contato"] = df["data_contato"].astype(str).str[:10]
    return df


def salvar_tudo(clientes: pd.DataFrame, contatos: pd.DataFrame) -> None:
    """Grava as duas abas no mesmo arquivo (precisa gravar as duas para não perder nenhuma).

    A gravação é feita em um arquivo temporário e só depois troca o original. Assim, se algo
    falhar no meio, o clientes.xlsx não fica corrompido.
    """
    temporario = ARQUIVO.with_name(ARQUIVO.name + ".tmp")
    try:
        with pd.ExcelWriter(temporario, engine="openpyxl") as escritor:
            clientes.to_excel(escritor, sheet_name=ABA, index=False)
            contatos.to_excel(escritor, sheet_name=ABA_CONTATOS, index=False)
        os.replace(temporario, ARQUIVO)  # se o Excel estiver aberto no Windows, gera PermissionError
    finally:
        temporario.unlink(missing_ok=True)  # apaga o temporário, tenha dado certo ou não


def salvar(clientes: pd.DataFrame) -> None:
    """Grava a lista de clientes mantendo o histórico de contatos como está."""
    salvar_tudo(clientes, carregar_contatos())


def garantir_arquivo() -> None:
    """Cria o clientes.xlsx (ou a aba que faltar) para o Power BI sempre encontrar as duas abas."""
    abas = _abas_existentes()
    if ABA not in abas or ABA_CONTATOS not in abas:
        salvar_tudo(carregar(), carregar_contatos())


def adicionar_cliente(nome: str, telefone: str, dia: int, mes: int):
    """Retorna (sucesso, mensagem)."""
    nome = (nome or "").strip()
    tel = limpar_telefone(telefone)

    erro = _validar_dados(nome, tel, dia, mes)
    if erro:
        return False, erro

    # O telefone é a chave do cliente: não pode repetir
    df = carregar()
    if tel in df["telefone"].values:
        return False, "Já existe um cliente com esse telefone."

    novo = pd.DataFrame([{"nome": nome, "telefone": tel, "dia": int(dia), "mes": int(mes)}])
    salvar(pd.concat([df, novo], ignore_index=True))
    return True, f"{nome} cadastrado(a) com sucesso!"


def editar_cliente(telefone_atual: str, nome: str, telefone: str, dia: int, mes: int):
    """Altera nome, telefone e data de um cliente (RF02). Retorna (sucesso, mensagem)."""
    nome = (nome or "").strip()
    tel = limpar_telefone(telefone)

    erro = _validar_dados(nome, tel, dia, mes)
    if erro:
        return False, erro

    df = carregar()
    if telefone_atual not in df["telefone"].values:
        return False, "Cliente não encontrado."
    # O novo telefone não pode ser de OUTRO cliente (pode ser o mesmo telefone dele)
    outros = df[df["telefone"] != telefone_atual]
    if tel in outros["telefone"].values:
        return False, "Já existe outro cliente com esse telefone."

    # Atualiza a linha do cliente
    posicao = df.index[df["telefone"] == telefone_atual][0]
    df.loc[posicao, COLUNAS] = [nome, tel, int(dia), int(mes)]

    # Se o telefone/nome mudou, acompanha também o histórico de contatos dele
    contatos = carregar_contatos()
    do_cliente = contatos["telefone"] == telefone_atual
    contatos.loc[do_cliente, ["telefone", "nome"]] = [tel, nome]

    salvar_tudo(df, contatos)
    return True, "Cliente atualizado com sucesso!"


def remover_cliente(telefone: str) -> None:
    df = carregar()
    # Mantém todos os clientes, exceto o que tem o telefone informado
    # (o histórico de contatos é mantido, para não perder as estatísticas)
    salvar(df[df["telefone"] != telefone])


# ---------- Alertas (RF03, RF05, RF07) ----------
def proximo_aniversario(dia: int, mes: int, hoje: date | None = None) -> date:
    """Próxima data de aniversário a partir de hoje (trata 29/02 em ano não bissexto).

    Data inexistente (ex.: 31/04) gera ValueError, em vez de ser corrigida em silêncio.
    """
    hoje = hoje or date.today()

    def data_no_ano(ano: int) -> date:
        # 29/02 em ano não bissexto: comemora-se em 28/02
        if (mes, dia) == (2, 29) and not calendar.isleap(ano):
            return date(ano, 2, 28)
        return date(ano, mes, dia)

    candidata = data_no_ano(hoje.year)
    return candidata if candidata >= hoje else data_no_ano(hoje.year + 1)


def classificar(dias: int) -> str:
    if dias <= 30:
        return "URGENTE (30 dias)"
    if dias <= 60:
        return "ATENÇÃO (60 dias)"
    return "Fora da janela"


def tabela_alertas(df: pd.DataFrame, hoje: date | None = None) -> pd.DataFrame:
    """Calcula dias restantes e status. Linhas inválidas da planilha são ignoradas (não derrubam o app)."""
    hoje = hoje or date.today()
    df, _ = separar_validas(df)
    if df.empty:
        return df.assign(proximo_aniversario=[], dias_restantes=[], status=[])
    df["proximo_aniversario"] = [
        proximo_aniversario(int(d), int(m), hoje) for d, m in zip(df["dia"], df["mes"])
    ]
    df["dias_restantes"] = [(p - hoje).days for p in df["proximo_aniversario"]]
    df["status"] = df["dias_restantes"].apply(classificar)
    return df.sort_values("dias_restantes").reset_index(drop=True)


# ---------- Contato e pedido (RF08) ----------
def adicionar_contato(tabela: pd.DataFrame, contatos: pd.DataFrame) -> pd.DataFrame:
    """Acrescenta a coluna "contato" na tabela de alertas, olhando o histórico de contatos.

    Só vale o contato feito para o aniversário DESTE ciclo (mesmo ano do próximo aniversário).
    Assim, quem foi contatado ano passado volta a aparecer como "Não contatado" neste ano.
    """
    tabela = tabela.copy()
    if tabela.empty:
        return tabela.assign(contato=[])

    def situacao(linha):
        achou = contatos[
            (contatos["telefone"] == linha["telefone"])
            & (contatos["ano_aniversario"] == linha["proximo_aniversario"].year)
        ]
        if achou.empty:
            return "Não contatado"
        return TEXTO_CONTATO.get(achou.iloc[0]["pedido_confirmado"], "Aguardando resposta")

    tabela["contato"] = tabela.apply(situacao, axis=1)
    return tabela


def registrar_contato(telefone: str, nome: str, ano_aniversario: int, hoje: date | None = None):
    """Marca que a confeiteira contatou o cliente neste ciclo. Retorna (sucesso, mensagem)."""
    hoje = hoje or date.today()
    contatos = carregar_contatos()
    ja_existe = (contatos["telefone"] == telefone) & (contatos["ano_aniversario"] == ano_aniversario)
    if ja_existe.any():
        return False, "Este cliente já foi marcado como contatado neste ano."

    novo = pd.DataFrame([{
        "telefone": telefone,
        "nome": nome,
        "data_contato": hoje.isoformat(),  # AAAA-MM-DD
        "ano_aniversario": int(ano_aniversario),
        "pedido_confirmado": "Aguardando",
    }])
    salvar_tudo(carregar(), pd.concat([contatos, novo], ignore_index=True))
    return True, "Contato registrado."


def atualizar_pedido(telefone: str, ano_aniversario: int, resultado: str):
    """Informa se o contato virou pedido ("Sim") ou não ("Não"). Retorna (sucesso, mensagem)."""
    if resultado not in ("Sim", "Não"):
        return False, "Resultado inválido."
    contatos = carregar_contatos()
    linha = (contatos["telefone"] == telefone) & (contatos["ano_aniversario"] == ano_aniversario)
    if not linha.any():
        return False, "Contato não encontrado."
    contatos.loc[linha, "pedido_confirmado"] = resultado
    salvar_tudo(carregar(), contatos)
    return True, "Resultado registrado."


def taxa_conversao(contatos: pd.DataFrame):
    """Pedidos confirmados ÷ contatos feitos. Retorna (contatos, confirmados, percentual)."""
    feitos = len(contatos)
    confirmados = int((contatos["pedido_confirmado"] == "Sim").sum())
    percentual = round(confirmados / feitos * 100, 1) if feitos else 0.0
    return feitos, confirmados, percentual


# ---------- Exportação (RF10) ----------
def exportar_excel(tabela: pd.DataFrame) -> bytes:
    """Gera, na memória, um Excel com a lista de clientes e seus alertas (para o botão de download)."""
    saida = pd.DataFrame({
        "Nome": tabela["nome"],
        "Telefone": tabela["telefone"],
        "Aniversário": [f"{int(d):02d}/{int(m):02d}" for d, m in zip(tabela["dia"], tabela["mes"])],
        "Próximo aniversário": [p.strftime("%d/%m/%Y") for p in tabela["proximo_aniversario"]],
        "Dias restantes": tabela["dias_restantes"],
        "Status": tabela["status"],
        "Contato": tabela["contato"],
    })
    memoria = io.BytesIO()
    saida.to_excel(memoria, sheet_name="alertas", index=False)
    return memoria.getvalue()


# ---------- Mensagem pronta (RF09) ----------
def link_whatsapp(nome: str, telefone: str, dia: int, mes: int) -> str:
    primeiro_nome = nome.split()[0]
    texto = (
        f"Oi, {primeiro_nome}! Tudo bem? Aqui é da Bolos da Lyz 🎂 "
        f"Vi que seu aniversário é dia {int(dia):02d}/{int(mes):02d} e já estou "
        f"abrindo as reservas de bolos e doces para a data. Quer garantir o seu?"
    )
    # 55 = Brasil; quote() converte espaços e acentos para o formato de link
    return f"https://wa.me/55{telefone}?text={urllib.parse.quote(texto)}"