import streamlit as st
import re

# CONTROLE DE ORÇAMENTO — compras podem ultrapassar 1000 EC

import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from datetime import datetime

ARQUIVO = Path("EDP_MARKET.xlsx")
ARQUIVO_FORMULARIO = Path("Respostas.xlsx")
ARQUIVO_SELECOES = Path("RegistroEDP.xlsx")
LIMITE_ERRO = 200  # acerto quando |resposta - gabarito| < 200

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
# LEITURA DAS ESCOLHAS DE EDPs DOS ALUNOS
# ============================================================

@st.cache_data
def carregar_selecoes_edps():
    """
    Lê o arquivo exportado pelo formulário de compras dos EDPs.

    Pela estrutura atual:
        Coluna C = nome do aluno
        Coluna D = EDP apresentado pelo aluno
        Coluna E = EDPs/princípios escolhidos pelo aluno

    A coluna E contém códigos Pxx. A função extrai os códigos,
    consulta o preço de cada EDP na aba "EDPs" do EDP_MARKET.xlsx
    e calcula a pontuação de cada princípio.
    """
    if not ARQUIVO_SELECOES.exists():
        return pd.DataFrame(), pd.DataFrame(), "Arquivo de seleções não encontrado."

    try:
        df = pd.read_excel(ARQUIVO_SELECOES, sheet_name=0)

        if df.shape[1] < 5:
            return pd.DataFrame(), pd.DataFrame(), (
                "O arquivo de seleções não possui as colunas C, D e E esperadas."
            )

        selecoes = df.iloc[:, [2, 3, 4]].copy()
        selecoes.columns = ["Aluno", "EDP apresentado", "Princípios escolhidos"]

        def normalizar_codigo(valor):
            if pd.isna(valor) or str(valor).strip() == "":
                return ""
            try:
                return f"P{int(float(valor)):02d}"
            except (TypeError, ValueError):
                texto = str(valor).strip().upper()
                if texto.startswith("P"):
                    try:
                        return f"P{int(texto[1:]):02d}"
                    except (TypeError, ValueError):
                        return texto
                return texto

        selecoes["EDP apresentado"] = selecoes["EDP apresentado"].apply(
            normalizar_codigo
        )

        def extrair_principios(valor):
            if pd.isna(valor):
                return []
            return re.findall(r"\bP\d{2}\b", str(valor).upper())

        selecoes["Princípios escolhidos"] = selecoes[
            "Princípios escolhidos"
        ].apply(extrair_principios)

        selecoes["Quantidade"] = selecoes["Princípios escolhidos"].apply(len)

        # Busca os preços diretamente na aba EDPs do EDP_MARKET.xlsx.
        _, edps = carregar_dados()

        precos = {}
        if edps is not None and not edps.empty:
            for _, linha in edps.iterrows():
                codigo = normalizar_codigo(linha.get("Código", ""))
                if not codigo:
                    continue
                try:
                    preco = float(linha.get("Preço"))
                    if pd.isna(preco):
                        preco = 0.0
                except (TypeError, ValueError):
                    preco = 0.0
                precos[codigo] = preco

        # Regra de pontuação solicitada.
        principios_1_ponto = {"P24", "P37"}
        principios_5_pontos = {"P22", "P39", "P41", "P60", "P30", "P02", "P46"}

        def pontos_principio(codigo):
            if codigo in principios_1_ponto:
                return 1
            if codigo in principios_5_pontos:
                return 5
            return 2

        registros = []
        for _, linha in selecoes.iterrows():
            for principio in linha["Princípios escolhidos"]:
                registros.append({
                    "Aluno": linha["Aluno"],
                    "EDP apresentado": linha["EDP apresentado"],
                    "Princípio": principio,
                    "Valor": precos.get(principio, 0.0),
                    "Pontuação": pontos_principio(principio),
                })

        detalhado = pd.DataFrame(
            registros,
            columns=[
                "Aluno",
                "EDP apresentado",
                "Princípio",
                "Valor",
                "Pontuação",
            ]
        )

        if len(detalhado) > 0:
            totais = (
                detalhado.groupby("Aluno", dropna=False)
                .agg(
                    **{
                        "Valor gasto": ("Valor", "sum"),
                        "Pontuação": ("Pontuação", "sum"),
                    }
                )
                .reset_index()
            )
        else:
            totais = pd.DataFrame(
                columns=["Aluno", "Valor gasto", "Pontuação"]
            )

        selecoes = selecoes.merge(totais, on="Aluno", how="left")
        selecoes["Valor gasto"] = selecoes["Valor gasto"].fillna(0.0)
        selecoes["Pontuação"] = selecoes["Pontuação"].fillna(0).astype(int)

        return selecoes, detalhado, ""

    except Exception as e:
        return pd.DataFrame(), pd.DataFrame(), f"Erro ao ler as seleções: {e}"


# ============================================================
# VALIDAÇÃO DAS RESPOSTAS DO FORMULÁRIO
# ============================================================

def validar_respostas_formulario():
    """
    Cruza as respostas do formulário com a aba Gabarito pelo EDP.

    Formulário:
        D = EDP
        E:G = resultados respondidos

    Gabarito:
        C = EDP
        E:G = valores corretos do gabarito (NÃO são alterados)
        H = Respostas acertadas

    O valor inicial de H é -1, significando "ainda não validado".
    Somente depois que a validação for executada H pode ser alterado
    para 1, 2 ou 3, conforme a quantidade de resultados corretos.

    Um resultado é considerado correto quando:
        abs(resposta - gabarito) < LIMITE_ERRO
    """
    if not ARQUIVO_FORMULARIO.exists():
        return None, "Arquivo de respostas do formulário não encontrado."

    try:
        respostas = pd.read_excel(
            ARQUIVO_FORMULARIO,
            sheet_name=0
        )

        colunas_form = list(respostas.columns)
        if len(colunas_form) < 7:
            return None, "O formulário não possui as colunas esperadas D:G."

        # Usa exatamente D:G pela posição, independentemente do texto do cabeçalho.
        respostas = respostas.iloc[:, [3, 4, 5, 6]].copy()
        respostas.columns = ["EDP", "Preprodu", "Distribui", "Descarte"]

        def normalizar_edp(valor):
            if pd.isna(valor):
                return ""
            try:
                return f"P{int(float(valor)):02d}"
            except (TypeError, ValueError):
                return str(valor).strip().upper()

        respostas["EDP"] = respostas["EDP"].apply(normalizar_edp)

        for col in ["Preprodu", "Distribui", "Descarte"]:
            respostas[col] = pd.to_numeric(respostas[col], errors="coerce")

        # Se houver mais de uma resposta para o mesmo EDP,
        # usa a resposta mais recente do formulário.
        respostas = respostas.drop_duplicates("EDP", keep="last")

        # Workbook com fórmulas preservadas para gravação.
        wb = load_workbook(ARQUIVO)

        if "Gabarito" not in wb.sheetnames:
            wb.close()
            return None, "A aba Gabarito não foi encontrada no EDP_MARKET.xlsx."

        ws = wb["Gabarito"]

        def calcular_gabarito(formula, linha):
            """
            Calcula o valor de uma fórmula simples do Excel quando o
            openpyxl não possui o resultado calculado em cache.

            O Gabarito atual usa fórmulas aritméticas baseadas na coluna D
            (CPF), por exemplo:
                =D2*400*10
                =D2*500*0.001*10*140
                =D2*10*4.1

            Mantemos a fórmula na planilha e calculamos apenas seu valor
            para fazer a validação.
            """
            if formula is None:
                return None

            if isinstance(formula, (int, float)):
                return float(formula)

            formula = str(formula).strip()

            if not formula.startswith("="):
                try:
                    return float(formula)
                except (TypeError, ValueError):
                    return None

            expressao = formula[1:].strip()

            # Substitui referências simples de células pelos respectivos
            # valores da mesma linha/planilha.
            import re

            def substituir_celula(match):
                referencia = match.group(0)
                valor = ws[referencia].value
                if valor is None:
                    raise ValueError(f"Célula {referencia} sem valor.")
                return str(float(valor))

            try:
                expressao = re.sub(
                    r"\$?[A-Z]{1,3}\$?\d+",
                    substituir_celula,
                    expressao
                )

                # Aceita somente números, operadores e parênteses.
                if not re.fullmatch(r"[0-9eE+\-*/(). ]+", expressao):
                    return None

                return float(eval(expressao, {"__builtins__": {}}, {}))
            except (TypeError, ValueError, SyntaxError, ZeroDivisionError):
                return None

        cabecalho = {
            str(ws.cell(row=1, column=c).value).strip(): c
            for c in range(1, ws.max_column + 1)
            if ws.cell(row=1, column=c).value is not None
        }

        col_codigo = cabecalho.get("EDP atribuído", 3)  # coluna C
        col_acertos = cabecalho.get("Respostas acertadas", 8)  # coluna H

        # Garante que o cabeçalho exista, sem alterar as colunas E:G.
        ws.cell(row=1, column=col_acertos, value="Respostas acertadas")

        mapa_respostas = respostas.set_index("EDP").to_dict("index")
        detalhes = []

        for linha in range(2, ws.max_row + 1):
            edp = normalizar_edp(ws.cell(linha, col_codigo).value)
            resposta = mapa_respostas.get(edp)

            # O valor do gabarito em E:G é lido SOMENTE de ws_valores.
            # Nunca escrevemos nas colunas E:G.
            acertos = 0
            respostas_encontradas = 0

            if resposta is not None:
                for col_idx, nome_col in zip(
                    [5, 6, 7],
                    ["Preprodu", "Distribui", "Descarte"]
                ):
                    valor_resposta = resposta[nome_col]

                    if pd.isna(valor_resposta):
                        continue

                    respostas_encontradas += 1

                    try:
                        formula_gabarito = ws.cell(linha, col_idx).value
                        gab = calcular_gabarito(formula_gabarito, linha)

                        if gab is not None and abs(float(valor_resposta) - gab) < LIMITE_ERRO:
                            acertos += 1
                    except (TypeError, ValueError):
                        pass

            # H começa em -1 e só muda quando esta rotina de validação
            # efetivamente analisa o EDP.
            #
            # Conforme solicitado:
            #   - 1, 2 ou 3 = quantidade de acertos;
            #   - -1 = nenhum acerto / ainda sem resultado positivo.
            #
            # As colunas E:G permanecem intactas.
            if resposta is not None and respostas_encontradas > 0 and acertos > 0:
                ws.cell(linha, col_acertos, acertos)
            else:
                ws.cell(linha, col_acertos, -1)

            detalhes.append({
                "EDP": edp,
                "Respostas enviadas": respostas_encontradas,
                "Respostas acertadas": (
                    acertos if (resposta is not None and respostas_encontradas > 0)
                    else -1
                ),
                "Situação": (
                    "3/3" if acertos == 3 else
                    "2/3" if acertos == 2 else
                    "1/3" if acertos == 1 else
                    "Não validado / nenhum acerto"
                )
            })

        wb.save(ARQUIVO)
        wb.close()

        return pd.DataFrame(detalhes), None

    except PermissionError:
        return None, (
            "Não foi possível atualizar o Excel. "
            "Feche o EDP_MARKET.xlsx caso ele esteja aberto."
        )
    except Exception as e:
        return None, f"Erro ao validar respostas: {e}"


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
    ["🎮 Início", "👥 Alunos", "🌱 EDPs", "🛒 Mercado", "📊 Portfólio", "🧩 Princípios escolhidos", "✅ Validação"]
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



elif menu == "🧩 Princípios escolhidos":
    st.title("🧩 Princípios escolhidos pelos alunos")
    st.caption(
        "Os valores são buscados na coluna 'Preço' da aba 'EDPs' do "
        "EDP_MARKET.xlsx. A pontuação segue a regra definida para cada EDP."
    )

    selecoes, detalhado, erro_selecoes = carregar_selecoes_edps()

    if erro_selecoes:
        st.error(f"❌ {erro_selecoes}")
    elif isinstance(selecoes, pd.DataFrame) and len(selecoes) > 0:
        total_alunos = len(selecoes)
        total_escolhas = int(selecoes["Quantidade"].sum())
        principios_unicos = (
            detalhado["Princípio"].nunique() if len(detalhado) > 0 else 0
        )
        gasto_total = float(selecoes["Valor gasto"].sum())
        pontuacao_total = int(selecoes["Pontuação"].sum())

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Alunos com respostas", total_alunos)
        c2.metric("Escolhas registradas", total_escolhas)
        c3.metric("Princípios diferentes", principios_unicos)
        c4.metric("Valor total gasto", f"{gasto_total:.0f} EC")
        c5.metric("Pontuação total", pontuacao_total)

        st.divider()

        st.subheader("👤 Consulta por aluno")
        nomes = selecoes["Aluno"].dropna().astype(str).tolist()

        aluno_escolhido = st.selectbox(
            "Selecione o aluno",
            nomes,
            key="consulta_principios_aluno"
        )

        registro = selecoes[
            selecoes["Aluno"].astype(str) == str(aluno_escolhido)
        ].iloc[0]

        principios_aluno = registro["Princípios escolhidos"]

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.write("**Aluno:**")
            st.write(registro["Aluno"])
        with c2:
            st.write("**EDP apresentado:**")
            st.write(registro["EDP apresentado"] or "Não informado")
        with c3:
            st.metric("Valor gasto", f"{float(registro['Valor gasto']):.0f} EC")
        with c4:
            st.metric("Pontuação", int(registro["Pontuação"]))

        st.metric("Quantidade de princípios escolhidos", len(principios_aluno))

        if principios_aluno:
            st.success(
                "Princípios escolhidos: " + ", ".join(principios_aluno)
            )

            detalhes_aluno = detalhado[
                detalhado["Aluno"].astype(str) == str(aluno_escolhido)
            ].copy()

            if len(detalhes_aluno) > 0:
                detalhes_aluno["Valor"] = detalhes_aluno["Valor"].map(
                    lambda x: f"{float(x):.0f} EC"
                )

            st.dataframe(
                detalhes_aluno[
                    ["Princípio", "Valor", "Pontuação"]
                ],
                use_container_width=True,
                hide_index=True
            )
        else:
            st.warning("Este aluno não informou princípios no formulário.")

        st.divider()

        st.subheader("📋 Todas as escolhas da turma")
        tabela_turma = selecoes[
            [
                "Aluno",
                "EDP apresentado",
                "Princípios escolhidos",
                "Quantidade",
                "Valor gasto",
                "Pontuação",
            ]
        ].copy()

        tabela_turma["Princípios escolhidos"] = tabela_turma[
            "Princípios escolhidos"
        ].apply(lambda x: ", ".join(x))

        tabela_turma["Valor gasto"] = tabela_turma["Valor gasto"].map(
            lambda x: f"{float(x):.0f} EC"
        )

        st.dataframe(
            tabela_turma,
            use_container_width=True,
            hide_index=True
        )

        st.divider()

        st.subheader("📊 Frequência de escolha dos princípios")
        if len(detalhado) > 0:
            frequencia = (
                detalhado.groupby("Princípio")
                .agg(
                    **{
                        "Número de escolhas": ("Princípio", "size"),
                        "Valor unitário": ("Valor", "first"),
                        "Pontuação unitária": ("Pontuação", "first"),
                    }
                )
                .reset_index()
                .sort_values(
                    ["Número de escolhas", "Princípio"],
                    ascending=[False, True]
                )
            )

            frequencia["Valor unitário"] = frequencia["Valor unitário"].map(
                lambda x: f"{float(x):.0f} EC"
            )

            st.dataframe(
                frequencia,
                use_container_width=True,
                hide_index=True
            )

            st.caption(
                "Pontuação: 1 ponto para P24 e P37; 5 pontos para "
                "P22, P39, P41, P60, P30, P02 e P46; 2 pontos para os demais EPDs."
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

# ============================================================
# VALIDAÇÃO
# ============================================================

elif menu == "✅ Validação":
    st.title("✅ Validação das respostas")
    st.caption(
        "Compara as respostas do arquivo Respostas.xlsx com os valores "
        "calculados do Gabarito. As colunas E:G do Gabarito não são alteradas."
    )

    st.info(
        "📌 Critério: uma resposta é considerada correta quando "
        f"|resposta − gabarito| < {LIMITE_ERRO}. "
        "A coluna 'Respostas acertadas' recebe 1, 2 ou 3 quando houver "
        "acertos; permanece -1 quando não houver nenhum acerto."
    )

    if not ARQUIVO_FORMULARIO.exists():
        st.error(
            f"❌ O arquivo **{ARQUIVO_FORMULARIO.name}** não foi encontrado "
            "na mesma pasta do aplicativo."
        )
    elif not ARQUIVO.exists():
        st.error(
            f"❌ O arquivo **{ARQUIVO.name}** não foi encontrado "
            "na mesma pasta do aplicativo."
        )
    else:
        if st.button("🔎 VALIDAR RESPOSTAS", type="primary"):
            with st.spinner("Validando respostas..."):
                resultado, erro = validar_respostas_formulario()

            if erro:
                st.error(f"❌ {erro}")
            else:
                st.success(
                    "✅ Validação concluída. A coluna 'Respostas acertadas' "
                    "foi atualizada no Gabarito."
                )
                st.cache_data.clear()

                if resultado is not None and not resultado.empty:
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("EDPs no gabarito", len(resultado))
                    c2.metric(
                        "1 acerto",
                        int((resultado["Respostas acertadas"] == 1).sum())
                    )
                    c3.metric(
                        "2 acertos",
                        int((resultado["Respostas acertadas"] == 2).sum())
                    )
                    c4.metric(
                        "3 acertos",
                        int((resultado["Respostas acertadas"] == 3).sum())
                    )

                    st.divider()
                    st.subheader("📋 Resultado da validação")

                    tabela = resultado.copy()
                    tabela["Situação"] = tabela["Situação"].replace(
                        "Não validado / nenhum acerto",
                        "❌ Nenhum acerto / -1"
                    )

                    st.dataframe(
                        tabela,
                        use_container_width=True,
                        hide_index=True
                    )

                    st.caption(
                        "Os valores de E, F e G do Gabarito permanecem "
                        "inalterados. Somente a coluna H é atualizada."
                    )

