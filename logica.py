"""Regras de negócio do sistema de aniversários - Bolos da Lyz."""
import re
import urllib.parse
from datetime import date
from pathlib import Path

import pandas as pd

ARQUIVO = Path("clientes.xlsx")
ABA = "clientes"
COLUNAS = ["nome", "telefone", "dia", "mes"]


# ---------- Validações (RNF04) ----------
def limpar_telefone(telefone: str) -> str:
    """Deixa só os dígitos e remove o código do país (55), se houver."""
    digitos = re.sub(r"\D", "", telefone or "")
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


# ---------- Leitura e gravação (RF02, RF04) ----------
def carregar() -> pd.DataFrame:
    if not ARQUIVO.exists():
        return pd.DataFrame(columns=COLUNAS)
    df = pd.read_excel(ARQUIVO, sheet_name=ABA, dtype={"telefone": str})
    return df[COLUNAS]


def salvar(df: pd.DataFrame) -> None:
    df.to_excel(ARQUIVO, sheet_name=ABA, index=False)


def adicionar_cliente(nome: str, telefone: str, dia: int, mes: int):
    """Retorna (sucesso, mensagem)."""
    nome = (nome or "").strip()
    tel = limpar_telefone(telefone)

    if not nome:
        return False, "Informe o nome do cliente."
    if not telefone_valido(tel):
        return False, "Telefone inválido. Use DDD + número (ex.: 32 91234-5678)."
    if not data_valida(dia, mes):
        return False, "Data de aniversário inválida."

    df = carregar()
    if tel in df["telefone"].values:
        return False, "Já existe um cliente com esse telefone."

    novo = pd.DataFrame([{"nome": nome, "telefone": tel, "dia": int(dia), "mes": int(mes)}])
    salvar(pd.concat([df, novo], ignore_index=True))
    return True, f"{nome} cadastrado(a) com sucesso!"


def remover_cliente(telefone: str) -> None:
    df = carregar()
    salvar(df[df["telefone"] != telefone])


# ---------- Alertas (RF03, RF05, RF07) ----------
def proximo_aniversario(dia: int, mes: int, hoje: date | None = None) -> date:
    """Próxima data de aniversário a partir de hoje (trata 29/02 em ano não bissexto)."""
    hoje = hoje or date.today()

    def data_no_ano(ano: int) -> date:
        try:
            return date(ano, mes, dia)
        except ValueError:  # 29/02 em ano não bissexto
            return date(ano, 2, 28)

    candidata = data_no_ano(hoje.year)
    return candidata if candidata >= hoje else data_no_ano(hoje.year + 1)


def classificar(dias: int) -> str:
    if dias <= 30:
        return "URGENTE (30 dias)"
    if dias <= 60:
        return "ATENÇÃO (60 dias)"
    return "Fora da janela"


def tabela_alertas(df: pd.DataFrame, hoje: date | None = None) -> pd.DataFrame:
    hoje = hoje or date.today()
    df = df.copy()
    if df.empty:
        return df.assign(proximo_aniversario=[], dias_restantes=[], status=[])
    df["proximo_aniversario"] = [
        proximo_aniversario(int(d), int(m), hoje) for d, m in zip(df["dia"], df["mes"])
    ]
    df["dias_restantes"] = [(p - hoje).days for p in df["proximo_aniversario"]]
    df["status"] = df["dias_restantes"].apply(classificar)
    return df.sort_values("dias_restantes").reset_index(drop=True)


# ---------- Mensagem pronta (RF09) ----------
def link_whatsapp(nome: str, telefone: str, dia: int, mes: int) -> str:
    primeiro_nome = nome.split()[0]
    texto = (
        f"Oi, {primeiro_nome}! Tudo bem? Aqui é da Bolos da Lyz 🎂 "
        f"Vi que seu aniversário é dia {int(dia):02d}/{int(mes):02d} e já estou "
        f"abrindo as reservas de bolos e doces para a data. Quer garantir o seu?"
    )
    return f"https://wa.me/55{telefone}?text={urllib.parse.quote(texto)}"
