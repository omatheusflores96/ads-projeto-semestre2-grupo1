"""Interface de cadastro e alertas - Bolos da Lyz.
Execute com:  streamlit run app.py
"""
import streamlit as st  # biblioteca que cria a interface web

import logica  # nosso arquivo com as regras (validação, Excel, alertas)

# Configurações da página (título da aba do navegador, ícone e largura)
st.set_page_config(page_title="Bolos da Lyz", page_icon="🎂", layout="centered")
st.title("🎂 Bolos da Lyz - Aniversários")

# Cria 3 abas na tela e guarda cada uma numa variável
aba_cadastro, aba_alertas, aba_clientes = st.tabs(
    ["➕ Cadastrar", "🔔 Alertas", "📋 Clientes"]
)

# ---------------- Cadastro ----------------
with aba_cadastro:
    # st.form agrupa os campos: nada é processado até clicar no botão de salvar
    # clear_on_submit=True limpa os campos depois de enviar
    with st.form("form_cliente", clear_on_submit=True):
        nome = st.text_input("Nome do cliente")
        telefone = st.text_input("Telefone (com DDD)", placeholder="32 91234-5678")
        # Dois campos lado a lado: dia (1 a 31) e mês (1 a 12)
        col1, col2 = st.columns(2)
        dia = col1.number_input("Dia", min_value=1, max_value=31, value=1, step=1)
        mes = col2.number_input("Mês", min_value=1, max_value=12, value=1, step=1)
        enviado = st.form_submit_button("Salvar cliente")

    # Só executa quando o botão foi clicado
    if enviado:
        try:
            # Tenta cadastrar; a função devolve (deu certo?, mensagem)
            ok, msg = logica.adicionar_cliente(nome, telefone, dia, mes)
            # Mostra a mensagem em verde (sucesso) ou vermelho (erro)
            (st.success if ok else st.error)(msg)
        except PermissionError:
            # Acontece quando o clientes.xlsx está aberto no Excel e não pode ser gravado
            st.error("Feche o arquivo clientes.xlsx no Excel e tente novamente.")

# ---------------- Alertas ----------------
with aba_alertas:
    # Filtro de janela de tempo (RF07)
    janela = st.radio(
        "Janela de tempo", ["Todas", "30 dias", "60 dias"], horizontal=True
    )
    # Calcula dias restantes e status de todos os clientes
    tabela = logica.tabela_alertas(logica.carregar())
    # Mantém só quem está dentro das janelas de 30 ou 60 dias
    tabela = tabela[tabela["status"] != "Fora da janela"] if not tabela.empty else tabela

    # Aplica o filtro escolhido pelo usuário
    if janela == "30 dias":
        tabela = tabela[tabela["dias_restantes"] <= 30]
    elif janela == "60 dias":
        tabela = tabela[(tabela["dias_restantes"] > 30) & (tabela["dias_restantes"] <= 60)]

    # Contadores no topo: quantos urgentes (até 30) e quantos em atenção (31 a 60)
    urgentes = int((tabela["dias_restantes"] <= 30).sum()) if not tabela.empty else 0
    atencao = len(tabela) - urgentes
    c1, c2 = st.columns(2)
    c1.metric("🔴 Urgentes (até 30 dias)", urgentes)
    c2.metric("🟡 Atenção (31 a 60 dias)", atencao)

    # Mensagem amigável quando não há ninguém na janela
    if tabela.empty:
        st.info("Nenhum aniversário nessa janela.")
    # Cria um cartão para cada cliente
    for _, linha in tabela.iterrows():
        urgente = linha["dias_restantes"] <= 30
        # Texto do cartão: nome, data (dd/mm) e dias que faltam
        msg = (
            f"**{linha['nome']}** - {int(linha['dia']):02d}/{int(linha['mes']):02d} "
            f"(faltam {linha['dias_restantes']} dias)"
        )
        # Vermelho (error) para urgente, amarelo (warning) para atenção
        (st.error if urgente else st.warning)(msg)
        # Botão que abre o WhatsApp com a mensagem pronta (key evita conflito entre botões)
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
    # Mostra a tabela completa de clientes cadastrados
    st.dataframe(df, use_container_width=True, hide_index=True)
    # Só mostra a opção de remover se houver clientes
    if not df.empty:
        # Lista suspensa: o valor é o telefone, mas exibe "Nome (telefone)"
        escolhido = st.selectbox(
            "Remover cliente",
            df["telefone"],
            format_func=lambda t: f"{df.loc[df['telefone'] == t, 'nome'].iloc[0]} ({t})",
        )
        if st.button("Remover"):
            try:
                logica.remover_cliente(escolhido)
                st.rerun()  # recarrega a página para atualizar a lista
            except PermissionError:
                st.error("Feche o arquivo clientes.xlsx no Excel e tente novamente.")