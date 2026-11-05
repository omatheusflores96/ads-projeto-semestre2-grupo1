"""Interface de cadastro e alertas - Bolos da Lyz.
Execute com:  streamlit run app.py
"""
import streamlit as st

import logica

st.set_page_config(page_title="Bolos da Lyz", page_icon="🎂", layout="centered")
st.title("🎂 Bolos da Lyz - Aniversários")

aba_cadastro, aba_alertas, aba_clientes = st.tabs(
    ["➕ Cadastrar", "🔔 Alertas", "📋 Clientes"]
)

# ---------------- Cadastro ----------------
with aba_cadastro:
    with st.form("form_cliente", clear_on_submit=True):
        nome = st.text_input("Nome do cliente")
        telefone = st.text_input("Telefone (com DDD)", placeholder="32 91234-5678")
        col1, col2 = st.columns(2)
        dia = col1.number_input("Dia", min_value=1, max_value=31, value=1, step=1)
        mes = col2.number_input("Mês", min_value=1, max_value=12, value=1, step=1)
        enviado = st.form_submit_button("Salvar cliente")

    if enviado:
        try:
            ok, msg = logica.adicionar_cliente(nome, telefone, dia, mes)
            (st.success if ok else st.error)(msg)
        except PermissionError:
            st.error("Feche o arquivo clientes.xlsx no Excel e tente novamente.")

# ---------------- Alertas ----------------
with aba_alertas:
    janela = st.radio(
        "Janela de tempo", ["Todas", "30 dias", "60 dias"], horizontal=True
    )
    tabela = logica.tabela_alertas(logica.carregar())
    tabela = tabela[tabela["status"] != "Fora da janela"] if not tabela.empty else tabela

    if janela == "30 dias":
        tabela = tabela[tabela["dias_restantes"] <= 30]
    elif janela == "60 dias":
        tabela = tabela[(tabela["dias_restantes"] > 30) & (tabela["dias_restantes"] <= 60)]

    urgentes = int((tabela["dias_restantes"] <= 30).sum()) if not tabela.empty else 0
    atencao = len(tabela) - urgentes
    c1, c2 = st.columns(2)
    c1.metric("🔴 Urgentes (até 30 dias)", urgentes)
    c2.metric("🟡 Atenção (31 a 60 dias)", atencao)

    if tabela.empty:
        st.info("Nenhum aniversário nessa janela.")
    for _, linha in tabela.iterrows():
        urgente = linha["dias_restantes"] <= 30
        msg = (
            f"**{linha['nome']}** - {int(linha['dia']):02d}/{int(linha['mes']):02d} "
            f"(faltam {linha['dias_restantes']} dias)"
        )
        (st.error if urgente else st.warning)(msg)
        st.link_button(
            "Enviar mensagem no WhatsApp",
            logica.link_whatsapp(
                linha["nome"], linha["telefone"], linha["dia"], linha["mes"]
            ),
            key=f"wa_{linha['telefone']}",
        )

# ---------------- Lista de clientes ----------------
with aba_clientes:
    df = logica.carregar()
    st.dataframe(df, use_container_width=True, hide_index=True)
    if not df.empty:
        escolhido = st.selectbox(
            "Remover cliente",
            df["telefone"],
            format_func=lambda t: f"{df.loc[df['telefone'] == t, 'nome'].iloc[0]} ({t})",
        )
        if st.button("Remover"):
            try:
                logica.remover_cliente(escolhido)
                st.rerun()
            except PermissionError:
                st.error("Feche o arquivo clientes.xlsx no Excel e tente novamente.")
