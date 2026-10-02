import streamlit as st

# CONTROLE DE ORÇAMENTO — compras podem ultrapassar 1000 EC

import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from datetime import datetime

ARQUIVO = Path("EDP_MARKET.xlsx")

st.set_page_config(
    page_title="EDP Market",
    page_icon="🌱",
    layout="wide"
)


# ============================================================
# FUNÇÕES DE DADOS
# ============================================================

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

    # Garante as colunas necessárias, caso a planilha ainda esteja na versão inicial.
    if "Saldo" not in alunos.columns:
        alunos["Saldo"] = alunos.get("Orçamento inicial", 1000)
    if "Pontos" not in alunos.columns:
        alunos["Pontos"] = 0

    if "Vendedor" not in edps.columns:
        edps["Vendedor"] = ""
    if "Preço" not in edps.columns:
        edps["Preço"] = None
    if "Status" not in edps.columns:
        edps["Status"] = "Não anunciado"

    return alunos, edps


def carregar_compras_excel():
    """Carrega as compras permanentes da aba Compras, se ela existir."""
    if not ARQUIVO.exists():
        return []

    try:
        wb = load_workbook(ARQUIVO, read_only=True, data_only=True)
        if "Compras" not in wb.sheetnames:
            wb.close()
            return []

        ws = wb["Compras"]
        valores = list(ws.iter_rows(values_only=True))
        wb.close()

        if not valores:
            return []

        cabecalho = [str(v).strip() if v is not None else "" for v in valores[0]]
        registros = []
        for linha in valores[1:]:
            if not any(v is not None and str(v).strip() != "" for v in linha):
                continue
            registro = dict(zip(cabecalho, linha))
            if registro.get("Comprador ID") is None or registro.get("Código") is None:
                continue
            try:
                registro["Comprador ID"] = int(registro["Comprador ID"])
            except (TypeError, ValueError):
                pass
            if "Preço" in registro and registro["Preço"] is not None:
                try:
                    registro["Preço"] = float(registro["Preço"])
                except (TypeError, ValueError):
                    pass
            registros.append(registro)

        return registros

    except Exception as e:
        st.error(f"❌ Erro ao carregar a aba Compras: {e}")
        return []


def registrar_compras_excel(compras_novas):
    """Grava um lote de compras na aba Compras, criando a aba se necessário."""
    if not compras_novas:
        return True

    colunas = [
        "ID compra",
        "Comprador ID",
        "Comprador",
        "Vendedor",
        "Código",
        "EDP",
        "Etapa",
        "Preço",
        "Data/hora",
    ]

    try:
        wb = load_workbook(ARQUIVO)

        if "Compras" not in wb.sheetnames:
            ws = wb.create_sheet("Compras")
            for col_idx, coluna in enumerate(colunas, start=1):
                ws.cell(row=1, column=col_idx, value=coluna)
        else:
            ws = wb["Compras"]
            # Garante um cabeçalho compatível caso a aba já exista.
            cabecalho_existente = [
                ws.cell(row=1, column=c).value
                for c in range(1, ws.max_column + 1)
            ]
            cabecalho_existente = [
                str(v).strip() if v is not None else ""
                for v in cabecalho_existente
            ]
            if cabecalho_existente != colunas:
                # Mantém dados existentes, mas adiciona/organiza as colunas esperadas.
                dados_existentes = list(ws.iter_rows(min_row=2, values_only=True))
                wb.remove(ws)
                ws = wb.create_sheet("Compras")
                for col_idx, coluna in enumerate(colunas, start=1):
                    ws.cell(row=1, column=col_idx, value=coluna)
                for linha in dados_existentes:
                    registro = dict(zip(cabecalho_existente, linha))
                    nova_linha = [registro.get(c) for c in colunas]
                    ws.append(nova_linha)

        # Define o próximo ID de compra.
        ids = []
        for valor in ws.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
            v = valor[0]
            if isinstance(v, (int, float)):
                ids.append(int(v))
        proximo_id = max(ids, default=0) + 1

        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for compra in compras_novas:
            ws.append([
                proximo_id,
                compra["Comprador ID"],
                compra["Comprador"],
                compra["Vendedor"],
                compra["Código"],
                compra["EDP"],
                compra["Etapa"],
                compra["Preço"],
                agora,
            ])
            proximo_id += 1

        wb.save(ARQUIVO)
        return True

    except PermissionError:
        st.error(
            "❌ Não foi possível registrar as compras no Excel. "
            "Feche o arquivo EDP_MARKET.xlsx caso ele esteja aberto no Excel."
        )
        return False
    except Exception as e:
        st.error(f"❌ Erro ao registrar as compras no Excel: {e}")
        return False


def salvar_aba(df, nome_aba):
    """Atualiza apenas uma aba do Excel, preservando as demais."""
    if not ARQUIVO.exists():
        st.error("Arquivo EDP_MARKET.xlsx não encontrado.")
        return False

    try:
        # openpyxl permite alterar uma aba sem apagar as outras.
        wb = load_workbook(ARQUIVO)

        if nome_aba in wb.sheetnames:
            ws = wb[nome_aba]
            wb.remove(ws)

        ws = wb.create_sheet(nome_aba)

        # Cabeçalho
        for col_idx, coluna in enumerate(df.columns, start=1):
            ws.cell(row=1, column=col_idx, value=coluna)

        # Dados
        for row_idx, (_, linha) in enumerate(df.iterrows(), start=2):
            for col_idx, valor in enumerate(linha, start=1):
                # NaN/NaT não devem ir para o Excel.
                if pd.isna(valor):
                    valor = None
                ws.cell(row=row_idx, column=col_idx, value=valor)

        wb.save(ARQUIVO)
        return True

    except PermissionError:
        st.error(
            "❌ Não foi possível salvar o Excel. "
            "Verifique se o arquivo EDP_MARKET.xlsx está aberto em outro programa."
        )
        return False
    except Exception as e:
        st.error(f"❌ Erro ao salvar o Excel: {e}")
        return False


def atualizar_edp_no_excel(codigo, vendedor, preco):
    """Coloca um EDP à venda diretamente pelo Streamlit."""
    try:
        wb = load_workbook(ARQUIVO)
        ws = wb["EDPs"]

        # Localiza as colunas pelo nome.
        cabecalho = {
            str(ws.cell(row=1, column=c).value).strip(): c
            for c in range(1, ws.max_column + 1)
        }

        col_codigo = cabecalho["Código"]
        col_vendedor = cabecalho["Vendedor"]
        col_preco = cabecalho["Preço"]
        col_status = cabecalho["Status"]

        for linha in range(2, ws.max_row + 1):
            valor_codigo = str(ws.cell(row=linha, column=col_codigo).value or "").strip().upper()

            if valor_codigo == codigo:
                ws.cell(row=linha, column=col_vendedor, value=vendedor)
                ws.cell(row=linha, column=col_preco, value=float(preco))
                ws.cell(row=linha, column=col_status, value="À venda")
                wb.save(ARQUIVO)
                return True

        return False

    except PermissionError:
        st.error(
            "❌ Não foi possível salvar o preço. "
            "Feche o EDP_MARKET.xlsx caso ele esteja aberto no Excel."
        )
        return False
    except Exception as e:
        st.error(f"❌ Erro ao atualizar o EDP: {e}")
        return False


# ============================================================
# CARREGAMENTO
# ============================================================

alunos, edps = carregar_dados()

if alunos is None or edps is None:
    st.error("Não foi possível encontrar a planilha EDP_MARKET.xlsx.")
    st.stop()

# Status de entrega dos EDPs durante a simulação.
# Não altera a compra nem o orçamento; será usado futuramente na pontuação.
if "entregas_edp" not in st.session_state:
    st.session_state.entregas_edp = {}

# Carrega as compras permanentes do Excel uma única vez por sessão.
# Assim, as compras continuam disponíveis após reruns/reinício do Streamlit.
if "compras_carregadas" not in st.session_state:
    st.session_state.compras = carregar_compras_excel()
    st.session_state.compras_carregadas = True

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
    ["🎮 Início", "👥 Alunos", "🌱 EDPs", "🛒 Mercado", "📊 Portfólio"]
)


# ============================================================
# INÍCIO
# ============================================================

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
        "Durante a aula, o professor poderá definir o preço de cada EDP "
        "diretamente no Streamlit. O Excel será usado como banco de dados, "
        "sem necessidade de ficar alternando de tela."
    )


# ============================================================
# ALUNOS
# ============================================================

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


# ============================================================
# EDPs
# ============================================================

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


# ============================================================
# MERCADO
# ============================================================

elif menu == "🛒 Mercado":

    st.title("🛒 EDP MARKET")
    st.subheader("Mercado de Decisões em Ecodesign")

    st.info(
        "🎤 Nesta tela você controla a apresentação, define o preço e coloca "
        "o EDP à venda. Depois, no mesmo local, os colegas podem comprar. "
        "Não é necessário abrir o Excel durante a aula."
    )

    # --------------------------------------------------------
    # BLOCO 1 — APRESENTAÇÃO E DEFINIÇÃO DO PREÇO
    # --------------------------------------------------------

    st.markdown("## 🎤 Apresentação e venda")

    nomes_alunos = alunos["Nome"].dropna().tolist()

    if not nomes_alunos:
        st.warning("Nenhum aluno cadastrado.")
        st.stop()

    vendedor_nome = st.selectbox(
        "👤 Quem está apresentando?",
        nomes_alunos,
        key="vendedor_apresentacao"
    )

    vendedor = alunos[alunos["Nome"] == vendedor_nome].iloc[0]
    codigo_vendedor = str(vendedor["EDP atribuído"]).strip().upper()

    if not codigo_vendedor or codigo_vendedor not in edp_lookup:
        st.error("❌ Este aluno não possui um EDP válido atribuído.")
    else:
        info_vendedor = edp_lookup[codigo_vendedor]

        c1, c2 = st.columns([3, 1])

        with c1:
            st.markdown(f"### {codigo_vendedor}")
            st.write(f"**EDP:** {info_vendedor['EDP']}")
            
        edp_atual = edps[edps["Código"] == codigo_vendedor].iloc[0]
        status_atual = str(edp_atual["Status"] or "Não anunciado").strip()

        with c2:
            st.metric("Status", status_atual)

        if status_atual.lower() == "à venda":
            preco_atual = edp_atual["Preço"]
            if pd.notna(preco_atual):
                st.success(
                    f"🟢 {codigo_vendedor} já está à venda por "
                    f"**{float(preco_atual):.0f} EcoCoins**."
                )
            st.caption("Este EDP já foi anunciado e não precisa ser anunciado novamente.")

        else:
            st.write("**Defina o preço após a apresentação:**")

            preco = st.number_input(
                "Preço do EDP (EcoCoins)",
                min_value=1,
                max_value=1000,
                value=250,
                step=10,
                key=f"preco_{codigo_vendedor}"
            )

            if st.button(
                "🟢 COLOCAR EDP À VENDA",
                type="primary",
                key=f"vender_{codigo_vendedor}"
            ):
                sucesso = atualizar_edp_no_excel(
                    codigo_vendedor,
                    vendedor_nome,
                    preco
                )

                if sucesso:
                    st.success(
                        f"✅ {codigo_vendedor} colocado à venda por "
                        f"**{preco:.0f} EcoCoins**."
                    )
                    st.caption(
                        f"Salvo no Excel às {datetime.now().strftime('%H:%M:%S')}."
                    )
                    st.cache_data.clear()
                    st.rerun()

    st.divider()

    # --------------------------------------------------------
    # BLOCO 2 — MERCADO / COMPRAS
    # --------------------------------------------------------

    st.markdown("## 🛒 Registro rápido das compras")

    st.info(
        "💡 Em vez de registrar uma compra por vez, selecione o EDP vendido "
        "e marque todos os alunos que compraram. Depois clique uma única vez "
        "em **REGISTRAR COMPRADORES**."
    )

    disponiveis = edps[
        edps["Status"].fillna("").astype(str).str.strip().str.lower()
        == "à venda"
    ].copy()

    if len(disponiveis) == 0:
        st.warning("Nenhum EDP está à venda neste momento.")
    else:
        opcoes_edp = {}
        for _, linha in disponiveis.iterrows():
            codigo = str(linha["Código"])
            vendedor_edp = str(linha["Vendedor"])
            preco_edp = float(linha["Preço"])
            opcoes_edp[codigo] = (
                f"{codigo} — {linha['EDP']} | "
                f"{preco_edp:.0f} EC | {vendedor_edp}"
            )

        codigo_compra = st.selectbox(
            "🌱 Qual EDP está sendo vendido agora?",
            list(opcoes_edp.keys()),
            format_func=lambda codigo: opcoes_edp[codigo],
            key="edp_registro_venda"
        )

        edp_venda = disponiveis[disponiveis["Código"] == codigo_compra].iloc[0]
        vendedor_edp = str(edp_venda["Vendedor"])
        preco_edp = float(edp_venda["Preço"])

        st.markdown(
            f"### {codigo_compra} — {edp_venda['EDP']}"
        )
        c1, c2 = st.columns(2)
        with c1:
            st.write(f"**Vendedor:** {vendedor_edp}")
        with c2:
            st.write(f"**Preço:** {preco_edp:.0f} EcoCoins")

        nomes_compradores = [
            nome for nome in nomes_alunos if nome != vendedor_edp
        ]

        compradores_selecionados = st.multiselect(
            "👥 Marque todos os alunos que compraram este EDP:",
            nomes_compradores,
            placeholder="Clique para selecionar os compradores...",
            key=f"compradores_{codigo_compra}"
        )

        if compradores_selecionados:
            st.success(
                f"{len(compradores_selecionados)} comprador(es) selecionado(s)."
            )

        if st.button(
            "🛒 REGISTRAR COMPRADORES",
            type="primary",
            disabled=len(compradores_selecionados) == 0,
            key=f"registrar_{codigo_compra}"
        ):
            if "compras" not in st.session_state:
                st.session_state.compras = []

            compradores_ja_registrados = {
                (c["Comprador ID"], c["Código"])
                for c in st.session_state.compras
            }

            compras_novas = []
            ignorados = 0

            for nome_comprador in compradores_selecionados:
                linha_comprador = alunos[
                    alunos["Nome"] == nome_comprador
                ].iloc[0]
                comprador_id = int(linha_comprador["ID"])
                chave = (comprador_id, codigo_compra)

                if chave in compradores_ja_registrados:
                    ignorados += 1
                    continue

                # O orçamento NÃO bloqueia a compra.
                # O sistema apenas acumula o valor gasto para o controle
                # do professor. O saldo pode ficar negativo.
                compras_novas.append({
                    "Comprador ID": comprador_id,
                    "Comprador": nome_comprador,
                    "Vendedor": vendedor_edp,
                    "Código": codigo_compra,
                    "EDP": str(edp_venda["EDP"]),
                    "Etapa": str(edp_venda["Etapa"]),
                    "Preço": preco_edp
                })

            registrados = len(compras_novas)

            # Primeiro grava no Excel. Só adiciona à sessão se a gravação
            # permanente tiver sido concluída com sucesso.
            if registrados > 0:
                sucesso_compras = registrar_compras_excel(compras_novas)
                if sucesso_compras:
                    st.session_state.compras.extend(compras_novas)
                else:
                    registrados = 0
            if ignorados > 0:
                st.warning(
                    f"⚠️ {ignorados} seleção(ões) não foram registradas "
                    "(a compra já havia sido registrada)."
                )

            st.rerun()

    # --------------------------------------------------------
    # CONTROLE DE ORÇAMENTO — PROFESSOR
    # --------------------------------------------------------

    if "compras" in st.session_state and len(st.session_state.compras) > 0:
        compras_orcamento = pd.DataFrame(st.session_state.compras)

        gastos_por_aluno = (
            compras_orcamento.groupby("Comprador ID")["Preço"]
            .sum()
            .to_dict()
        )
    else:
        gastos_por_aluno = {}

    controle = alunos[["ID", "Nome", "Orçamento inicial"]].copy()
    controle["Gasto acumulado"] = controle["ID"].apply(
        lambda x: float(gastos_por_aluno.get(int(x), 0))
    )
    controle["Saldo disponível"] = (
        controle["Orçamento inicial"].astype(float)
        - controle["Gasto acumulado"]
    )
    controle["% gasto"] = (
        controle["Gasto acumulado"]
        / controle["Orçamento inicial"].replace(0, pd.NA)
        * 100
    ).fillna(0)

    st.divider()
    st.subheader("💰 Controle de orçamento — professor")
    st.caption(
        "Controle interno da simulação. Os valores abaixo são calculados "
        "a partir das compras registradas na aba Compras e ainda não alteram "
        "o saldo-base gravado na aba Alunos."
    )

    total_gasto = controle["Gasto acumulado"].sum()
    total_orcamento = controle["Orçamento inicial"].sum()
    total_saldo = controle["Saldo disponível"].sum()

    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric("Orçamento total", f"{total_orcamento:.0f} EC")
    with m2:
        st.metric("Já gasto", f"{total_gasto:.0f} EC")
    with m3:
        st.metric("Saldo restante", f"{total_saldo:.0f} EC")

    st.dataframe(
        controle[["ID", "Nome", "Orçamento inicial", "Gasto acumulado", "Saldo disponível", "% gasto"]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Orçamento inicial": st.column_config.NumberColumn("Orçamento", format="%.0f EC"),
            "Gasto acumulado": st.column_config.NumberColumn("Gasto acumulado", format="%.0f EC"),
            "Saldo disponível": st.column_config.NumberColumn("Saldo disponível", format="%.0f EC"),
            "% gasto": st.column_config.NumberColumn("% gasto", format="%.1f%%"),
        }
    )

    aluno_controle = st.selectbox(
        "🔎 Ver orçamento de um aluno",
        controle["Nome"].tolist(),
        key="aluno_controle_orcamento"
    )
    detalhe = controle[controle["Nome"] == aluno_controle].iloc[0]
    d1, d2, d3 = st.columns(3)
    with d1:
        st.metric("Orçamento", f"{float(detalhe['Orçamento inicial']):.0f} EC")
    with d2:
        st.metric("Gasto", f"{float(detalhe['Gasto acumulado']):.0f} EC")
    with d3:
        st.metric("Saldo", f"{float(detalhe['Saldo disponível']):.0f} EC")

    # --------------------------------------------------------
    # COMPRAS POR ALUNO — PROFESSOR
    # --------------------------------------------------------
    st.divider()
    st.subheader("👤 Compras de cada aluno")
    st.caption(
        "Consulta das compras registradas. "
        "A seleção abaixo permite acompanhar o portfólio de cada aluno "
        "sem alterar nenhuma compra."
    )

    if "compras" not in st.session_state or len(st.session_state.compras) == 0:
        st.info("Ainda não há compras registradas.")
    else:
        compras_alunos = pd.DataFrame(st.session_state.compras)

        nomes_compras = (
            alunos["Nome"]
            .dropna()
            .tolist()
        )

        aluno_consulta = st.selectbox(
            "🔎 Selecione o aluno",
            nomes_compras,
            key="consulta_compras_aluno"
        )

        linha_aluno = alunos[
            alunos["Nome"] == aluno_consulta
        ].iloc[0]
        id_aluno_consulta = int(linha_aluno["ID"])

        compras_aluno = compras_alunos[
            compras_alunos["Comprador ID"] == id_aluno_consulta
        ].copy()

        if len(compras_aluno) == 0:
            st.warning("Este aluno ainda não realizou nenhuma compra.")
        else:
            total_aluno = compras_aluno["Preço"].sum()

            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("Compras", len(compras_aluno))
            with c2:
                st.metric("Total gasto", f"{total_aluno:.0f} EC")
            with c3:
                st.metric(
                    "Saldo",
                    f"{1000 - total_aluno:.0f} EC"
                )

            st.dataframe(
                compras_aluno[
                    [
                        "Código",
                        "EDP",
                        "Vendedor",
                        "Preço"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Preço": st.column_config.NumberColumn(
                        "Preço",
                        format="%.0f EC"
                    )
                }
            )



    # --------------------------------------------------------
    # CONTROLE DE ENTREGA DOS EDPs — PROFESSOR
    # --------------------------------------------------------
    st.divider()
    st.subheader("📦 Controle de entrega dos EDPs")
    st.caption(
        "Controle interno do professor. Registre se cada EDP foi entregue "
        "corretamente após a resolução do desafio de implementação. "
        "Este registro não altera as compras nem o orçamento e será usado "
        "posteriormente na ponderação dos pontos."
    )

    edps_vendidos = edps[
        edps["Status"].fillna("").astype(str).str.strip().str.lower() == "à venda"
    ].copy()

    if len(edps_vendidos) == 0:
        st.info("Nenhum EDP foi colocado à venda ainda.")
    else:
        for _, linha_entrega in edps_vendidos.iterrows():
            codigo_entrega = str(linha_entrega["Código"])
            status_entrega = st.session_state.entregas_edp.get(
                codigo_entrega, "Ainda não avaliado"
            )

            c1, c2, c3 = st.columns([1, 4, 2])

            with c1:
                st.markdown(f"### {codigo_entrega}")

            with c2:
                st.write(
                    f"**{linha_entrega['EDP']}**  "
                    f"Vendedor: {linha_entrega['Vendedor']} | "
                    f"Preço: {float(linha_entrega['Preço']):.0f} EC"
                )

            with c3:
                novo_status = st.selectbox(
                    "Entrega",
                    [
                        "Ainda não avaliado",
                        "Entregue corretamente",
                        "Não entregue corretamente"
                    ],
                    index=[
                        "Ainda não avaliado",
                        "Entregue corretamente",
                        "Não entregue corretamente"
                    ].index(status_entrega),
                    key=f"entrega_{codigo_entrega}"
                )
                st.session_state.entregas_edp[codigo_entrega] = novo_status

    # --------------------------------------------------------
    # COMPRAS REGISTRADAS
    # --------------------------------------------------------

    if "compras" in st.session_state and len(st.session_state.compras) > 0:

        compras = pd.DataFrame(st.session_state.compras)

        st.divider()
        st.subheader("📦 Compras registradas")

        st.dataframe(
            compras[
                ["Comprador", "Código", "EDP", "Vendedor", "Preço"]
            ],
            use_container_width=True,
            hide_index=True
        )

        total = compras["Preço"].sum()

        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Compras registradas", len(compras))
        with c2:
            st.metric("Total comprometido", f"{total:.0f} EC")
        with c3:
            entregues = sum(
                1 for status in st.session_state.entregas_edp.values()
                if status == "Entregue corretamente"
            )
            st.metric("EDPs entregues corretamente", entregues)

        st.success(
            "💾 As compras desta turma estão gravadas permanentemente na aba "
            "**Compras** do arquivo EDP_MARKET.xlsx. O registro também é "
            "recarregado automaticamente quando o aplicativo é reiniciado."
        )


elif menu == "📊 Portfólio":
    st.title("📊 Portfólio dos alunos")
    st.caption(
        "Visualização das compras de cada aluno organizadas pelas cinco etapas "
        "do ciclo de vida. É permitido ter vários EDPs em uma mesma etapa."
    )

    ETAPAS_PORTFOLIO = [
        "Extração / materiais",
        "Produção",
        "Distribuição",
        "Uso",
        "Fim de vida",
    ]

    mapa_etapas = {
        "Extração/materiais": "Extração / materiais",
        "Extração / materiais": "Extração / materiais",
        "Produção": "Produção",
        "Distribuição": "Distribuição",
        "Uso": "Uso",
        "Fim de vida": "Fim de vida",
    }

    if "compras" not in st.session_state or len(st.session_state.compras) == 0:
        st.info(
            "Ainda não há compras registradas. "
            "O portfólio será preenchido automaticamente após as compras."
        )
    else:
        compras_portfolio = pd.DataFrame(st.session_state.compras)
        alunos_portfolio = alunos[["ID", "Nome"]].drop_duplicates().copy()

        aluno_portfolio = st.selectbox(
            "👤 Selecione o aluno",
            alunos_portfolio["Nome"].tolist(),
            key="portfolio_aluno"
        )

        id_portfolio = int(
            alunos_portfolio.loc[
                alunos_portfolio["Nome"] == aluno_portfolio, "ID"
            ].iloc[0]
        )

        compras_aluno = compras_portfolio[
            compras_portfolio["Comprador ID"] == id_portfolio
        ].copy()

        if len(compras_aluno) > 0:
            compras_aluno["Etapa portfolio"] = (
                compras_aluno["Etapa"]
                .map(mapa_etapas)
                .fillna(compras_aluno["Etapa"])
            )
        else:
            compras_aluno["Etapa portfolio"] = pd.Series(
                dtype="object"
            )

        etapas_cobertas = set(
            compras_aluno["Etapa portfolio"]
        ).intersection(ETAPAS_PORTFOLIO)
        cobertura = len(etapas_cobertas)

        # Resumo do aluno
        c1, c2, c3 = st.columns(3)

        with c1:
            st.metric("Etapas cobertas", f"{cobertura}/5")

        with c2:
            st.metric("EDPs comprados", len(compras_aluno))

        with c3:
            gasto = (
                compras_aluno["Preço"].sum()
                if len(compras_aluno) > 0
                else 0
            )
            st.metric("Investimento", f"{gasto:.0f} EC")

        st.divider()
        st.subheader("🎯 Cobertura do ciclo de vida")

        # Cinco cartões, um para cada etapa
        cols = st.columns(5)

        for i, etapa in enumerate(ETAPAS_PORTFOLIO):
            quantidade = int(
                (compras_aluno["Etapa portfolio"] == etapa).sum()
            )

            with cols[i]:
                if quantidade > 0:
                    st.success(f"✅ {etapa}")
                else:
                    st.error(f"❌ {etapa}")

                st.metric("EDPs", quantidade)

        if cobertura == 5:
            st.success(
                "🏆 MISSÃO DE COBERTURA COMPLETA — o aluno possui pelo menos "
                "um EDP em cada uma das cinco etapas."
            )
        else:
            faltantes = [
                etapa for etapa in ETAPAS_PORTFOLIO
                if etapa not in etapas_cobertas
            ]
            st.warning(
                f"Faltam {5 - cobertura} etapa(s) para completar a cobertura."
            )
            st.write(
                "**Etapas ainda sem EDP:** " + ", ".join(faltantes)
            )

        st.divider()
        st.subheader("🧩 EDPs comprados")

        if len(compras_aluno) == 0:
            st.info("Este aluno ainda não comprou nenhum EDP.")
        else:
            for etapa in ETAPAS_PORTFOLIO:
                dados_etapa = compras_aluno[
                    compras_aluno["Etapa portfolio"] == etapa
                ].copy()

                if len(dados_etapa) == 0:
                    continue

                st.markdown(f"### {etapa}")

                st.dataframe(
                    dados_etapa[
                        ["Código", "EDP", "Vendedor", "Preço"]
                    ],
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Preço": st.column_config.NumberColumn(
                            "Preço",
                            format="%.0f EC"
                        )
                    }
                )

        st.divider()
        st.subheader("📋 Visão geral da turma")

        resumo_turma = []

        for _, aluno_row in alunos_portfolio.iterrows():
            id_aluno = int(aluno_row["ID"])

            compras = compras_portfolio[
                compras_portfolio["Comprador ID"] == id_aluno
            ].copy()

            if len(compras) > 0:
                etapas = set(
                    compras["Etapa"]
                    .map(mapa_etapas)
                    .dropna()
                ).intersection(ETAPAS_PORTFOLIO)
                investimento = compras["Preço"].sum()
            else:
                etapas = set()
                investimento = 0

            resumo_turma.append({
                "ID": id_aluno,
                "Aluno": aluno_row["Nome"],
                "Cobertura": f"{len(etapas)}/5",
                "EDPs comprados": len(compras),
                "Investimento": investimento,
            })

        st.dataframe(
            pd.DataFrame(resumo_turma),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Investimento": st.column_config.NumberColumn(
                    "Investimento",
                    format="%.0f EC"
                )
            }
        )
