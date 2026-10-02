import streamlit as st
import pandas as pd
from pathlib import Path

ARQUIVO = Path("EDP_MARKET.xlsx")

st.set_page_config(
    page_title="EDP Market",
    page_icon="🌱",
    layout="wide"
)

@st.cache_data
def carregar_dados():
    if not ARQUIVO.exists():
        return None, None

    alunos = pd.read_excel(ARQUIVO, sheet_name="Alunos")
    edps = pd.read_excel(ARQUIVO, sheet_name="EDPs")

    alunos["EDP atribuído"] = (
        alunos["EDP atribuído"].fillna("").astype(str).str.strip().str.upper()
    )
    edps["Código"] = (
        edps["Código"].fillna("").astype(str).str.strip().str.upper()
    )

    return alunos, edps


alunos, edps = carregar_dados()

if alunos is None or edps is None:
    st.error("Não foi possível encontrar a planilha EDP_MARKET.xlsx.")
    st.stop()

edp_lookup = (
    edps.drop_duplicates("Código")
    .set_index("Código")[["EDP", "Etapa"]]
    .to_dict("index")
)

alunos["EDP"] = alunos["EDP atribuído"].map(
    lambda codigo: edp_lookup.get(codigo, {}).get("EDP", "") if codigo else ""
)
alunos["Etapa automática"] = alunos["EDP atribuído"].map(
    lambda codigo: edp_lookup.get(codigo, {}).get("Etapa", "") if codigo else ""
)

menu = st.sidebar.radio(
    "Navegação",
    ["🎮 Início", "👥 Alunos", "🌱 EDPs", "🛒 Mercado"]
)

if menu == "🎮 Início":

    st.title("🌱 EDP MARKET")
    st.subheader("Mercado de Decisões em Ecodesign")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("EDPs", len(edps))
    with col2:
        st.metric("Alunos", len(alunos))
    with col3:
        st.metric("Etapas", 5)
    with col4:
        st.metric("Orçamento inicial", "1000 EC")

    st.divider()

    atribuidos = int((alunos["EDP atribuído"] != "").sum())
    validos = int(
        ((alunos["EDP atribuído"] != "") &
         (alunos["EDP"].fillna("") != "")).sum()
    )

    c1, c2 = st.columns(2)

    with c1:
        st.metric("Alunos com EDP atribuído", f"{atribuidos}/{len(alunos)}")
    with c2:
        st.metric("Atribuições válidas", f"{validos}/{len(alunos)}")

    if validos == len(alunos):
        st.success("✅ Todos os alunos estão associados a um EDP válido.")
    else:
        st.warning(
            "⚠️ Existem alunos sem EDP ou com código de EDP não encontrado."
        )

    st.info(
        "O código digitado na coluna 'EDP atribuído' é usado para "
        "buscar automaticamente o nome e a etapa do EDP."
    )

elif menu == "👥 Alunos":

    st.title("👥 Alunos")
    st.write(f"**Total cadastrado:** {len(alunos)} alunos")

    colunas = [
        "ID", "Nome", "EDP atribuído", "EDP",
        "Etapa automática", "Saldo", "Pontos"
    ]

    st.dataframe(
        alunos[colunas],
        use_container_width=True,
        hide_index=True
    )

    st.divider()
    st.subheader("🔎 Consultar aluno")

    nomes_alunos = alunos["Nome"].dropna().tolist()

    if nomes_alunos:

        nome = st.selectbox("Selecione um aluno", nomes_alunos)
        aluno = alunos[alunos["Nome"] == nome].iloc[0]

        c1, c2, c3 = st.columns(3)

        with c1:
            st.metric("ID", int(aluno["ID"]))
        with c2:
            st.metric("Saldo", f'{float(aluno["Saldo"]):.0f} EC')
        with c3:
            st.metric("Pontos", f'{float(aluno["Pontos"]):.0f}')

        st.divider()

        codigo = aluno["EDP atribuído"]

        if codigo and codigo in edp_lookup:

            info = edp_lookup[codigo]

            st.success("✅ EDP identificado automaticamente")
            st.markdown(f"### {codigo}")

            d1, d2 = st.columns(2)

            with d1:
                st.write("**EDP:**")
                st.write(info.get("EDP", "Não informado"))

            with d2:
                st.write("**Etapa do ciclo de vida:**")
                st.write(info.get("Etapa", "Não informada"))

        elif codigo:
            st.error(
                f"❌ O código **{codigo}** não foi encontrado na aba EDPs."
            )
        else:
            st.warning("⚠️ Este aluno ainda não possui um EDP atribuído.")


elif menu == "🛒 Mercado":

    st.title("🛒 EDP MARKET")
    st.subheader("Mercado de Decisões em Ecodesign")

    st.info(
        "Nesta etapa, os alunos utilizam seus 1.000 EcoCoins para "
        "comprar os EDPs apresentados pelos colegas."
    )

    # --------------------------------------------------------
    # Seleção do comprador
    # --------------------------------------------------------

    nomes_alunos = alunos["Nome"].dropna().tolist()

    if not nomes_alunos:
        st.warning("Nenhum aluno cadastrado.")
        st.stop()

    comprador_nome = st.selectbox(
        "👤 Quem está comprando?",
        nomes_alunos,
        key="comprador_mercado"
    )

    comprador = alunos[alunos["Nome"] == comprador_nome].iloc[0]
    comprador_id = int(comprador["ID"])
    saldo = float(comprador["Saldo"])

    c1, c2, c3 = st.columns(3)

    with c1:
        st.metric("Comprador", comprador_nome)

    with c2:
        st.metric("Saldo disponível", f"{saldo:.0f} EC")

    with c3:
        st.metric("Orçamento inicial", "1000 EC")

    st.divider()

    # --------------------------------------------------------
    # EDPs disponíveis
    # --------------------------------------------------------

    st.subheader("🌱 EDPs disponíveis para compra")

    disponiveis = edps[
        edps["Status"].fillna("").astype(str).str.strip().str.lower()
        == "à venda"
    ].copy()

    if len(disponiveis) == 0:
        st.warning(
            "Nenhum EDP está à venda neste momento. "
            "Após as apresentações, o professor deverá cadastrar "
            "o preço e alterar o status do EDP para 'À venda'."
        )
    else:

        for _, edp in disponiveis.iterrows():

            codigo = str(edp["Código"])
            nome_edp = str(edp["EDP"])
            etapa = str(edp["Etapa"])
            vendedor = str(edp["Vendedor"])
            preco = float(edp["Preço"])

            # O próprio vendedor não pode comprar seu EDP.
            if vendedor == comprador_nome:
                continue

            with st.container(border=True):

                col1, col2, col3 = st.columns([4, 2, 1])

                with col1:
                    st.markdown(f"### {codigo}")
                    st.write(nome_edp)
                    st.write(f"**Etapa:** {etapa}")
                    st.write(f"**Vendedor:** {vendedor}")

                with col2:
                    st.metric("Preço", f"{preco:.0f} EC")

                    if preco <= saldo:
                        st.caption("💰 Compra possível")
                    else:
                        st.caption("❌ Saldo insuficiente")

                with col3:

                    pode_comprar = (
                        preco <= saldo
                        and vendedor != comprador_nome
                    )

                    if st.button(
                        "🛒 COMPRAR",
                        key=f"comprar_{codigo}_{comprador_id}",
                        disabled=not pode_comprar
                    ):

                        # A compra será registrada na sessão nesta primeira versão.
                        if "compras" not in st.session_state:
                            st.session_state.compras = []

                        st.session_state.compras.append({
                            "Comprador ID": comprador_id,
                            "Comprador": comprador_nome,
                            "Vendedor": vendedor,
                            "Código": codigo,
                            "EDP": nome_edp,
                            "Etapa": etapa,
                            "Preço": preco
                        })

                        st.success(
                            f"Compra registrada: {codigo} por {preco:.0f} EC."
                        )

                        st.rerun()

    # --------------------------------------------------------
    # Compras realizadas nesta sessão
    # --------------------------------------------------------

    if "compras" in st.session_state:

        compras = pd.DataFrame(st.session_state.compras)

        compras_aluno = compras[
            compras["Comprador"] == comprador_nome
        ]

        if len(compras_aluno) > 0:

            st.divider()
            st.subheader("📦 Suas compras nesta sessão")

            st.dataframe(
                compras_aluno[
                    [
                        "Código",
                        "EDP",
                        "Etapa",
                        "Vendedor",
                        "Preço"
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )

            total = compras_aluno["Preço"].sum()

            st.metric(
                "Total comprometido",
                f"{total:.0f} EC"
            )

            st.warning(
                "⚠️ Nesta primeira versão, as compras ficam registradas "
                "na sessão do Streamlit. Ainda não alteramos o Excel. "
                "Vamos fazer a gravação permanente no próximo passo."
            )

elif menu == "🌱 EDPs":

    st.title("🌱 Eco-design Principles")
    st.write(f"**Total de EDPs:** {len(edps)}")

    st.dataframe(
        edps,
        use_container_width=True,
        hide_index=True
    )

    st.divider()
    st.subheader("🔎 Consultar EDP")

    codigos = edps["Código"].dropna().tolist()

    if codigos:

        codigo = st.selectbox("Selecione um EDP", codigos)
        edp = edps[edps["Código"] == codigo].iloc[0]

        st.markdown(f"### {edp['Código']}")

        c1, c2, c3 = st.columns(3)

        with c1:
            st.write("**EDP:**")
            st.write(edp["EDP"] if pd.notna(edp["EDP"]) else "Não cadastrado")

        with c2:
            st.write("**Etapa:**")
            st.write(edp["Etapa"] if pd.notna(edp["Etapa"]) else "Não definida")

        with c3:
            st.write("**Status:**")
            st.write(edp["Status"] if pd.notna(edp["Status"]) else "Não anunciado")

        vendedores = alunos[alunos["EDP atribuído"] == codigo]

        if len(vendedores) > 0:
            st.divider()
            st.write("**Aluno(s) responsável(is):**")
            st.dataframe(
                vendedores[["ID", "Nome"]],
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("Nenhum aluno foi associado a este EDP.")
