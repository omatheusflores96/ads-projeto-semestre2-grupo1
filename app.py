"""Interface de cadastro e alertas - Bolos da Lyz (versão 3).
Execute com:  streamlit run app.py
"""
import streamlit as st  # biblioteca que cria a interface web

import logica  # nosso arquivo com as regras (validação, Excel, alertas)

st.set_page_config(page_title="Bolos da Lyz", page_icon="🎂", layout="centered")
st.title("🎂 Bolos da Lyz - Aniversários")


def executar_seguro(funcao, *argumentos):
    """Executa uma leitura/gravação e mostra uma mensagem amigável se algo der errado."""
    try:
        return funcao(*argumentos)
    except PermissionError:
        # O Windows trava o arquivo quando ele está aberto no Excel
        st.error("Feche o arquivo clientes.xlsx no Excel e tente novamente.")
        st.stop()
    except logica.ErroDeDados as erro:
        # Arquivo corrompido, aba ou coluna ausente: a mensagem já vem explicada
        st.error(str(erro))
        st.stop()


def inteiro_ou(valor, padrao, minimo, maximo):
    """Converte para inteiro dentro dos limites; se o valor for inválido, usa o padrão."""
    try:
        return min(max(int(valor), minimo), maximo)
    except (TypeError, ValueError):
        return padrao


def texto_ou_vazio(valor):
    """Evita mostrar 'nan' quando a célula da planilha está vazia."""
    return "" if valor is None or valor != valor else str(valor)


# Garante que o clientes.xlsx exista, com as abas "clientes" e "contatos"
executar_seguro(logica.garantir_arquivo)

# Carrega os dados uma vez por execução da tela
clientes = executar_seguro(logica.carregar)
contatos = executar_seguro(logica.carregar_contatos)
# Linhas com dados inválidos (ex.: digitadas direto na planilha) ficam de fora dos alertas
_, invalidas = logica.separar_validas(clientes)

aba_cadastro, aba_alertas, aba_clientes = st.tabs(
    ["➕ Cadastrar", "🔔 Alertas", "📋 Clientes"]
)

# ---------------- Cadastro ----------------
with aba_cadastro:
    # clear_on_submit=True limpa os campos depois de enviar
    with st.form("form_cliente", clear_on_submit=True):
        nome = st.text_input("Nome do cliente")
        telefone = st.text_input("Telefone (com DDD)", placeholder="32 91234-5678")
        col1, col2 = st.columns(2)
        dia = col1.number_input("Dia", min_value=1, max_value=31, value=1, step=1)
        mes = col2.number_input("Mês", min_value=1, max_value=12, value=1, step=1)
        enviado = st.form_submit_button("Salvar cliente")

    if enviado:
        ok, msg = executar_seguro(logica.adicionar_cliente, nome, telefone, dia, mes)
        (st.success if ok else st.error)(msg)

# ---------------- Alertas ----------------
with aba_alertas:
    if not invalidas.empty:
        st.warning(
            f"{len(invalidas)} registro(s) com dados inválidos foram ignorados nos alertas. "
            "Veja o motivo na aba Clientes."
        )

    # Filtro de janela de tempo (RF07)
    janela = st.radio(
        "Janela de tempo", ["Todas", "30 dias", "60 dias"], horizontal=True
    )

    # Calcula dias restantes e status e acrescenta a situação do contato (RF08)
    tabela = logica.adicionar_contato(logica.tabela_alertas(clientes), contatos)
    # Mantém só quem está dentro das janelas de 30 ou 60 dias
    tabela = tabela[tabela["status"] != "Fora da janela"]

    if janela == "30 dias":
        tabela = tabela[tabela["dias_restantes"] <= 30]
    elif janela == "60 dias":
        tabela = tabela[(tabela["dias_restantes"] > 30) & (tabela["dias_restantes"] <= 60)]

    # Os contadores mostram só quem AINDA falta contatar
    a_contatar = tabela[tabela["contato"] == "Não contatado"]
    urgentes = int((a_contatar["dias_restantes"] <= 30).sum())
    atencao = len(a_contatar) - urgentes
    feitos, confirmados, taxa = logica.taxa_conversao(contatos)

    c1, c2, c3 = st.columns(3)
    c1.metric("🔴 Urgentes a contatar", urgentes)
    c2.metric("🟡 Atenção a contatar", atencao)
    c3.metric("📈 Conversão", f"{taxa:.0f}%", help=f"{confirmados} pedidos em {feitos} contatos")

    if tabela.empty:
        st.info("Nenhum aniversário nessa janela.")

    # Cria um cartão para cada cliente
    for _, linha in tabela.iterrows():
        tel = linha["telefone"]
        ano = linha["proximo_aniversario"].year  # ano do aniversário deste ciclo
        situacao = linha["contato"]
        msg = (
            f"**{linha['nome']}** - {int(linha['dia']):02d}/{int(linha['mes']):02d} "
            f"(faltam {linha['dias_restantes']} dias)"
        )

        # A cor do cartão depende da situação: verde (pedido), azul (sem pedido) ou pelo prazo
        if situacao == "Pedido confirmado":
            st.success(msg + " - ✅ pedido confirmado")
        elif situacao == "Sem pedido":
            st.info(msg + " - sem pedido desta vez")
        elif linha["dias_restantes"] <= 30:
            st.error(msg)
        else:
            st.warning(msg)

        # Linha de botões do cartão
        botoes = st.columns(3)
        botoes[0].link_button(
            "WhatsApp",
            logica.link_whatsapp(linha["nome"], tel, linha["dia"], linha["mes"]),
            key=f"wa_{tel}",
        )

        if situacao == "Não contatado":
            if botoes[1].button("📞 Contatei", key=f"contato_{tel}_{ano}"):
                executar_seguro(logica.registrar_contato, tel, linha["nome"], ano)
                st.rerun()  # recarrega a tela para mostrar o novo estado
        elif situacao == "Aguardando resposta":
            if botoes[1].button("✅ Pedido confirmado", key=f"sim_{tel}_{ano}"):
                executar_seguro(logica.atualizar_pedido, tel, ano, "Sim")
                st.rerun()
            if botoes[2].button("❌ Não fechou", key=f"nao_{tel}_{ano}"):
                executar_seguro(logica.atualizar_pedido, tel, ano, "Não")
                st.rerun()

# ---------------- Lista de clientes ----------------
with aba_clientes:
    # Mensagem guardada antes de recarregar a tela (senão ela sumiria)
    if "aviso" in st.session_state:
        st.success(st.session_state.pop("aviso"))

    st.dataframe(clientes, hide_index=True)

    # Registros com problema, com o motivo, para a confeiteira corrigir
    if not invalidas.empty:
        with st.expander(f"⚠️ {len(invalidas)} registro(s) com problema"):
            st.dataframe(invalidas, hide_index=True)

    # Só dá para editar ou remover quem tem telefone (é a chave do cliente)
    com_telefone = clientes[clientes["telefone"].fillna("").astype(str).str.strip() != ""]

    if not com_telefone.empty:
        # Exportação para Excel (RF10): lista de clientes com os alertas
        exportacao = logica.adicionar_contato(logica.tabela_alertas(clientes), contatos)
        if not exportacao.empty:
            st.download_button(
                "⬇️ Baixar lista com alertas (Excel)",
                data=logica.exportar_excel(exportacao),
                file_name="clientes_alertas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.divider()
        # Lista suspensa: o valor é o telefone, mas exibe "Nome (telefone)"
        escolhido = st.selectbox(
            "Selecione um cliente",
            com_telefone["telefone"],
            format_func=lambda t: (
                f"{texto_ou_vazio(com_telefone.loc[com_telefone['telefone'] == t, 'nome'].iloc[0])} ({t})"
            ),
        )
        atual = com_telefone[com_telefone["telefone"] == escolhido].iloc[0]

        # Edição (RF02): o formulário já vem preenchido com os dados atuais
        st.subheader("Editar cliente")
        with st.form(f"form_editar_{escolhido}"):
            novo_nome = st.text_input("Nome", value=texto_ou_vazio(atual["nome"]))
            novo_tel = st.text_input("Telefone (com DDD)", value=texto_ou_vazio(atual["telefone"]))
            e1, e2 = st.columns(2)
            novo_dia = e1.number_input(
                "Dia", min_value=1, max_value=31, value=inteiro_ou(atual["dia"], 1, 1, 31), step=1
            )
            novo_mes = e2.number_input(
                "Mês", min_value=1, max_value=12, value=inteiro_ou(atual["mes"], 1, 1, 12), step=1
            )
            salvou = st.form_submit_button("Salvar alterações")

        if salvou:
            ok, msg = executar_seguro(
                logica.editar_cliente, escolhido, novo_nome, novo_tel, novo_dia, novo_mes
            )
            if ok:
                st.session_state["aviso"] = msg
                st.rerun()
            else:
                st.error(msg)

        # Remoção
        if st.button("🗑️ Remover este cliente"):
            executar_seguro(logica.remover_cliente, escolhido)
            st.rerun()