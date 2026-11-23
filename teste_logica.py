# Testes automáticos das regras do sistema (versão 3).
# Execute com:  python teste_logica.py   (se nenhum erro aparecer, está tudo certo)
import io
import logging
import os
import subprocess
import sys
import tempfile
import time
from datetime import date
from pathlib import Path

import pandas as pd

import logica

PASTA = Path(__file__).resolve().parent


def novo_arquivo(nome="clientes.xlsx"):
    """Aponta o sistema para um arquivo temporário, sem mexer no clientes.xlsx de verdade."""
    logica.ARQUIVO = Path(tempfile.mkdtemp()) / nome


def deve_falhar(excecao, funcao, *argumentos):
    """Confere se a função gera a exceção esperada e devolve a exceção."""
    try:
        funcao(*argumentos)
    except excecao as erro:
        return erro
    raise AssertionError(f"Esperava {excecao.__name__}, mas nada aconteceu")


novo_arquivo()

# ---- Validações ----
assert logica.limpar_telefone("+55 (32) 91234-5678") == "32912345678"
assert logica.telefone_valido("32912345678") and not logica.telefone_valido("1234")
assert logica.data_valida(29, 2) and not logica.data_valida(31, 2)

# ---- Datas e alertas ----
assert logica.proximo_aniversario(10, 11, date(2026, 11, 5)) == date(2026, 11, 10)
assert logica.proximo_aniversario(1, 11, date(2026, 11, 5)) == date(2027, 11, 1)
assert logica.proximo_aniversario(29, 2, date(2026, 11, 5)) == date(2027, 2, 28)
assert logica.proximo_aniversario(29, 2, date(2027, 11, 5)) == date(2028, 2, 29)
deve_falhar(ValueError, logica.proximo_aniversario, 31, 4, date(2026, 11, 5))  # data inexistente
assert logica.classificar(30) == "URGENTE (30 dias)"
assert logica.classificar(31) == "ATENÇÃO (60 dias)"
assert logica.classificar(61) == "Fora da janela"

# ---- Arquivo novo: as duas abas são criadas ----
logica.garantir_arquivo()
assert logica.carregar().empty and logica.carregar_contatos().empty

# ---- Cadastro e duplicidade ----
assert logica.adicionar_cliente("Maria Silva", "(32) 91234-5678", 10, 12)[0]
assert not logica.adicionar_cliente("Outra Maria", "32912345678", 1, 1)[0]  # telefone repetido
assert logica.adicionar_cliente("João Souza", "32 98888-7777", 5, 3)[0]
assert not logica.adicionar_cliente("Ana", "32 97777-6666", 31, 2)[0]       # 31/02

# ---- Edição (RF02) ----
assert logica.editar_cliente("32912345678", "Maria S. Lima", "32 95555-4444", 11, 12)[0]
clientes = logica.carregar()
assert "32955554444" in clientes["telefone"].values and "32912345678" not in clientes["telefone"].values
assert not logica.editar_cliente("32955554444", "Maria", "32988887777", 11, 12)[0]  # telefone de outro
assert not logica.editar_cliente("00000000000", "X", "32911112222", 1, 1)[0]        # não existe

# ---- Contato e pedido (RF08) ----
hoje = date(2026, 11, 17)
alertas = logica.adicionar_contato(logica.tabela_alertas(logica.carregar(), hoje), logica.carregar_contatos())
assert set(alertas["contato"]) == {"Não contatado"}

assert logica.registrar_contato("32955554444", "Maria S. Lima", 2026, hoje)[0]
assert not logica.registrar_contato("32955554444", "Maria S. Lima", 2026, hoje)[0]  # já contatada
alertas = logica.adicionar_contato(logica.tabela_alertas(logica.carregar(), hoje), logica.carregar_contatos())
assert alertas.loc[alertas["telefone"] == "32955554444", "contato"].iloc[0] == "Aguardando resposta"

assert logica.atualizar_pedido("32955554444", 2026, "Sim")[0]
assert logica.registrar_contato("32988887777", "João Souza", 2027, hoje)[0]
assert logica.atualizar_pedido("32988887777", 2027, "Não")[0]
assert not logica.atualizar_pedido("32988887777", 2027, "Talvez")[0]
assert logica.taxa_conversao(logica.carregar_contatos()) == (2, 1, 50.0)

# O contato de um ano não vale para o ano seguinte
alertas_2027 = logica.adicionar_contato(
    logica.tabela_alertas(logica.carregar(), date(2027, 1, 1)), logica.carregar_contatos()
)
maria = alertas_2027[alertas_2027["telefone"] == "32955554444"].iloc[0]
assert maria["contato"] == "Não contatado"

# Editar o telefone acompanha o histórico de contatos
assert logica.editar_cliente("32955554444", "Maria S. Lima", "32 94444-3333", 11, 12)[0]
assert "32944443333" in logica.carregar_contatos()["telefone"].values

# ---- Exportação (RF10) ----
dados = logica.exportar_excel(logica.adicionar_contato(
    logica.tabela_alertas(logica.carregar(), hoje), logica.carregar_contatos()))
exportado = pd.read_excel(io.BytesIO(dados), sheet_name="alertas")
assert list(exportado.columns) == [
    "Nome", "Telefone", "Aniversário", "Próximo aniversário", "Dias restantes", "Status", "Contato"]
assert len(exportado) == 2

# ---- Remoção mantém o histórico de contatos ----
logica.remover_cliente("32988887777")
assert len(logica.carregar()) == 1 and len(logica.carregar_contatos()) == 2

# ---- Gravação segura: não sobra arquivo temporário ----
assert not logica.ARQUIVO.with_name(logica.ARQUIVO.name + ".tmp").exists()

# ---- Linhas inválidas (dados digitados direto na planilha) não derrubam o sistema ----
sujo = pd.DataFrame({
    "nome": ["Ok Normal", "Tel Curto", "Data Ruim", "  ", "Duplicada", "Ok Duplicada", "Sem Dia", "Formatada", "Decimal"],
    "telefone": ["32911110001", "1234", "32911110003", "32911110004", "32911110005",
                 "32911110005", "32911110007", "(32) 91111-0008", "32911110009.0"],
    "dia": [10, 10, 31, 10, 10, 10, None, 10, 10],
    "mes": [12, 12, 4, 12, 12, 12, 12, 12, 12],
})
validas, invalidas = logica.separar_validas(sujo)
assert list(validas["nome"]) == ["Ok Normal", "Duplicada", "Formatada", "Decimal"]
assert list(validas["telefone"]) == ["32911110001", "32911110005", "32911110008", "32911110009"]
assert dict(zip(invalidas["nome"], invalidas["motivo"])) == {
    "Tel Curto": "telefone inválido",
    "Data Ruim": "data de aniversário inválida",
    "  ": "nome vazio",
    "Ok Duplicada": "telefone duplicado",
    "Sem Dia": "data de aniversário inválida",
}
alertas_sujo = logica.tabela_alertas(sujo, hoje)  # não pode dar erro
assert len(alertas_sujo) == 4
# Planilha vazia continua funcionando
assert logica.tabela_alertas(pd.DataFrame(columns=logica.COLUNAS), hoje).empty

# ---- Arquivo corrompido: mensagem amigável, sem derrubar com erro técnico ----
logging.disable(logging.CRITICAL)  # o sistema registra o erro no terminal; aqui o aviso é esperado e fica escondido
novo_arquivo()
logica.ARQUIVO.write_bytes(b"isto nao e um arquivo excel")
erro = deve_falhar(logica.ErroDeDados, logica.carregar)
assert "corrompido" in str(erro)
deve_falhar(logica.ErroDeDados, logica.garantir_arquivo)

# ---- Coluna ausente: a mensagem diz qual ----
novo_arquivo()
pd.DataFrame({"nome": ["Ana"], "telefone": ["32911112222"], "dia": [1]}).to_excel(
    logica.ARQUIVO, sheet_name="clientes", index=False)
erro = deve_falhar(logica.ErroDeDados, logica.carregar)
assert "mes" in str(erro)

# ---- Arquivo antigo (só com a aba "clientes") continua funcionando ----
novo_arquivo("antigo.xlsx")
pd.DataFrame([{"nome": "Ana", "telefone": "32911112222", "dia": 1, "mes": 1}]).to_excel(
    logica.ARQUIVO, sheet_name="clientes", index=False)
assert logica.carregar_contatos().empty
logica.garantir_arquivo()
assert logica.carregar_contatos().empty and len(logica.carregar()) == 1

# ---- Variável de ambiente CLIENTES_XLSX ----
alvo = Path(tempfile.mkdtemp()) / "outro.xlsx"
comando = [sys.executable, "-c", "import logica; print(logica.ARQUIVO)"]
com_variavel = subprocess.run(comando, cwd=PASTA, capture_output=True, text=True,
                              env={**os.environ, "CLIENTES_XLSX": str(alvo)})
assert com_variavel.stdout.strip() == str(alvo), com_variavel.stderr
sem_variavel = subprocess.run(comando, cwd=PASTA, capture_output=True, text=True,
                              env={k: v for k, v in os.environ.items() if k != "CLIENTES_XLSX"})
assert sem_variavel.stdout.strip() == "clientes.xlsx", sem_variavel.stderr

# ---- Desempenho (RNF01): 1.000 registros em até 5 segundos ----
novo_arquivo("volume.xlsx")
base = pd.DataFrame({
    "nome": [f"Cliente {i}" for i in range(1000)],
    "telefone": [f"3291{i:07d}" for i in range(1000)],
    "dia": [(i % 28) + 1 for i in range(1000)],
    "mes": [(i % 12) + 1 for i in range(1000)],
})
inicio = time.perf_counter()
logica.salvar(base)
lido = logica.carregar()
logica.tabela_alertas(lido)
duracao = time.perf_counter() - inicio
assert len(lido) == 1000 and duracao < 5, duracao

print(f"Todos os testes passaram. (1.000 registros: {duracao:.2f} s)")