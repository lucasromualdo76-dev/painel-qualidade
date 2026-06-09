import streamlit as st
import pandas as pd
import plotly.express as px
import streamlit.components.v1 as components
import base64
from pathlib import Path
from datetime import datetime, date
import os
import io
import re

from openai import AzureOpenAI
from openpyxl import load_workbook

import pdfplumber


# ======================================================
# CONFIGURAÇÃO DO AZURE OPENAI
# ======================================================
client = AzureOpenAI(
    api_key=os.getenv("AZURE_OPENAI_API_KEY"),
    api_version="2024-02-15-preview",
    azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT")
)
DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT")

# ======================================================
# CONFIGURAÇÃO DA PÁGINA
# ======================================================
st.set_page_config(page_title="Painel de Qualidade", layout="wide")

# ======================================================
# CSS GLOBAL
# ======================================================
st.markdown("""
<style>
/* Input compacto */
input[type="text"] {
    height: 32px !important;
    padding: 4px 8px !important;
    font-size: 13px !important;
}

/* Botão compacto */
button[data-testid="baseButton-secondary"],
button[kind="secondary"] {
    padding: 2px 10px !important;
    font-size: 13px !important;
    height: 32px !important;
}

/* Reduz espaçamento entre linhas */
div[data-testid="stHorizontalBlock"] {
    gap: 8px;
    margin-bottom: 4px;
}

/* Badge do contato (embaixo do badge da data) */
.contato-suporte {
    position: fixed;
    top: 110px; 
    right: 15px;
    z-index: 9999;
    font-size: 13px;
    color: #ffffff;
    background: rgba(0,0,0,0.45);
    padding: 6px 12px;
    border-radius: 6px;
    font-family: Arial;
}

/* Remove barra do header */
header[data-testid="stHeader"] { background: transparent !important; }
header[data-testid="stHeader"]::after { display: none; }
</style>

<div class="contato-suporte">
📩 <strong>Dúvidas:</strong> Lucas.silva9@volkswagen.com.br
</div>
""", unsafe_allow_html=True)

# ======================================================
# USUÁRIOS (exemplo simples - ideal depois migrar p/ SSO)
# ======================================================
USUARIOS = {
    "aannutb": "12345",
    "ufcmart": "12345",
    "vyplfbt": "12345",
    "gibvvr7": "12345",
    "admin": "admin"
}

# ======================================================
# SESSÃO
# ======================================================
if "logado" not in st.session_state:
    st.session_state.logado = False
if "usuario" not in st.session_state:
    st.session_state.usuario = ""

# ======================================================
# DATA + KW
# ======================================================
def aplicar_background_login():
    img = Path("login_bg.png")
    if img.exists():
        st.markdown(
            f"""
            <style>
            .stApp {{
                background-image: url("data:image/png;base64,{base64.b64encode(img.read_bytes()).decode()}");
                background-size: cover;
                background-position: center;
            }}
            section[data-testid="stSidebar"] {{
                display: none;
            }}
            </style>
            """,
            unsafe_allow_html=True
        )

def data_kw_atual():
    hoje = datetime.now()
    return f"{hoje.strftime('%d/%m/%Y')} | KW {hoje.isocalendar().week}"

def mostrar_data_kw():
    st.markdown(f"""
    <div style="
        position: fixed;
        top: 70px;
        right: 15px;
        z-index: 9999;
        font-size: 13px;
        color: white;
        background: rgba(0,0,0,0.45);
        padding: 6px 12px;
        border-radius: 6px;
        font-family: Arial;">
        📅 <strong>{data_kw_atual()}</strong>
    </div>
    """, unsafe_allow_html=True)

# ======================================================
# CARRINHO
# ======================================================
def mostrar_carrinho_animado_painel():
    st.markdown("""
    <style>
    .faixa-painel { height: 40px; overflow: hidden; }
    .carro {
        position: relative;
        font-size: 24px;
        animation: mover 12s linear infinite alternate;
    }
    @keyframes mover {
        from { left: 0; }
        to { left: calc(100% - 40px); }
    }
    </style>
    <div class="faixa-painel">
        <div class="carro">🚗</div>
    </div>
    """, unsafe_allow_html=True)

# ======================================================
# HELPERS EXCEL
# ======================================================
def _to_float_ptbr(value):
    """Converte valores numéricos pt-BR e ignora textos."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip()
    if not s:
        return None

    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return None

def extrair_total_coluna_j_openpyxl(uploaded_file, sheet_name):
    """
    Extrai o TOTAL da Coluna J (coluna 10 no Excel) varrendo de baixo pra cima
    e pegando o primeiro valor numérico encontrado.
    """
    wb = load_workbook(io.BytesIO(uploaded_file.getvalue()), data_only=True)
    ws = wb[sheet_name]

    col_j = 10  # J = 10
    for r in range(ws.max_row, 0, -1):
        v = ws.cell(row=r, column=col_j).value
        num = _to_float_ptbr(v)
        if num is not None:
            return num
    return None

def formatar_moeda_br(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ======================================================
# HELPERS PDF (MIS12 / MIS36)
# ======================================================
def _to_float_ptbr_num(s):
    """Converte número pt-BR do PDF (ex.: '108,4' -> 108.4)."""
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)

    txt = str(s).strip()
    if not txt:
        return None

    txt = txt.replace(" ", "")
    txt = re.sub(r"[^0-9,\.\-]", "", txt)
    if not txt:
        return None

    txt = txt.replace(".", "").replace(",", ".")
    try:
        return float(txt)
    except:
        return None

def extrair_anos_pdf(file_bytes):
    """Retorna anos HJ encontrados no PDF (1ª página)."""
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            if not pdf.pages:
                return []
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return []

    t = re.sub(r"\s+", " ", text.replace("\n", " ")).strip()
    years = sorted({int(y) for y in re.findall(r"\b20\d{2}\b", t)})
    return years

def extrair_titulo_pdf(file_bytes):
    """
    Extrai o título/código do relatório (ex.: 604-VW216-IND-CY21-24).
    Funciona para os relatórios 604 (como os PDFs enviados).
    """
    # Ex.: 604-VW216-IND-CY21-24 ou 604-VW216-EUR-CY21-24
    pattern = re.compile(r"\b\d{3}-VW\d{3}-[A-Z]{3}-CY\d{2}-\d{2}\b")

    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            if not pdf.pages:
                return None
            page = pdf.pages[0]
            text = page.extract_text() or ""
            text = re.sub(r"\s+", " ", text.replace("\n", " ")).strip()

            m = pattern.search(text)
            if m:
                return m.group(0)

            # fallback: procura "Nome do ficheiro 604-VW216-..."
            m2 = re.search(r"\bNome do ficheiro\b\s+(\d{3}-VW\d{3}-[A-Z]{3}-CY\d{2}-\d{2})\b", text)
            if m2:
                return m2.group(1)

    except Exception:
        return None

    return None

def extrair_mis12_mis36_por_ano_pdf(file_bytes, ano_alvo):
    """
    Extrai MIS12 e MIS36 por ANO usando o TEXTO da 1ª página do PDF (estável p/ relatórios 604).
    Retorna: {"ano": int, "MIS12": float|None, "MIS36": float|None, "anos_disponiveis": [..]}
    """
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            if not pdf.pages:
                return {"ano": int(ano_alvo), "MIS12": None, "MIS36": None, "anos_disponiveis": []}
            text = pdf.pages[0].extract_text() or ""
    except Exception:
        return {"ano": int(ano_alvo), "MIS12": None, "MIS36": None, "anos_disponiveis": []}

    t = re.sub(r"\s+", " ", text.replace("\n", " ")).strip()

    # pega colunas MIS (ordem do cabeçalho)
    header_match = re.search(r"\bHJ\b(.*?\bMIS36\b)", t)
    if header_match:
        header_part = header_match.group(1)
        mis_cols = re.findall(r"\bMIS\d+\b", header_part)
    else:
        mis_cols = re.findall(r"\bMIS\d+\b", t)

    # remove duplicados preservando ordem
    seen = set()
    mis_cols = [m for m in mis_cols if not (m in seen or seen.add(m))]

    # anos disponíveis
    years = sorted({int(y) for y in re.findall(r"\b20\d{2}\b", t)})

    ano = int(ano_alvo)
    if ano not in years:
        return {"ano": ano, "MIS12": None, "MIS36": None, "anos_disponiveis": years}

    # bloco do ano até próximo ano ou "Difere" ou "HJ Troca"
    pattern = rf"\b{ano}\b(.*?)(?=\b20\d{{2}}\b|\bDifere\b|\bHJ\b\s*Troca\b)"
    m = re.search(pattern, t)
    if not m:
        return {"ano": ano, "MIS12": None, "MIS36": None, "anos_disponiveis": years}

    block = m.group(1)

    nums = re.findall(r"-?\d+(?:\.\d{3})*,\d+|-?\d+,\d+", block)
    vals = [_to_float_ptbr_num(n) for n in nums]

    # mapeia por posição: MIS0->vals[0], MIS1->vals[1], ...
    mapping = {mis_cols[i]: (vals[i] if i < len(vals) else None) for i in range(len(mis_cols))}

    return {
        "ano": ano,
        "MIS12": mapping.get("MIS12"),
        "MIS36": mapping.get("MIS36"),
        "anos_disponiveis": years
    }


# ======================================================
# AGENDAMENTO
# ======================================================
def aba_agendamento_veiculos():
    st.subheader("🚗 Agendamento de Veículos")

    dia_selecionado = st.date_input(
        "📅 Selecione o dia do agendamento",
        value=date.today(),
        format="DD/MM/YYYY"
    )

    chave_dia = dia_selecionado.strftime("%Y-%m-%d")

    horarios = [
        "08:00 - 08:30", "08:30 - 09:00",
        "09:00 - 09:30", "09:30 - 10:00",
        "10:00 - 10:30", "10:30 - 11:00",
        "11:00 - 11:30", "11:30 - 12:00",
        "12:00 - 12:30", "12:30 - 13:00",
        "13:00 - 13:30", "13:30 - 14:00",
        "14:00 - 14:30", "14:30 - 15:00",
        "15:00 - 15:30", "15:30 - 16:00",
        "16:00 - 16:30", "16:30 - 17:00"
    ]

    if "agenda_veiculos" not in st.session_state:
        st.session_state.agenda_veiculos = {}

    if chave_dia not in st.session_state.agenda_veiculos:
        st.session_state.agenda_veiculos[chave_dia] = {
            h: {"usuario": "", "descricao": "", "salvo": False}
            for h in horarios
        }

    c_h, c_u, c_d, c_s = st.columns([2, 3, 5, 2])
    c_h.markdown("**Horário**")
    c_u.markdown("**Usuário**")
    c_d.markdown("**Descrição**")
    c_s.markdown("**Salvar**")

    st.divider()

    for h in horarios:
        dados = st.session_state.agenda_veiculos[chave_dia][h]

        col1, col2, col3, col4 = st.columns([2, 3, 5, 2])
        col1.write(h)

        col2.write(dados["usuario"] if dados["salvo"] else "")

        dados["descricao"] = col3.text_input(
            "",
            placeholder="Descreva o motivo da utilização do veículo",
            value=dados["descricao"],
            disabled=dados["salvo"],
            key=f"desc_{chave_dia}_{h}"
        )

        if dados["salvo"]:
            col4.markdown("✅")
        else:
            if col4.button("Salvar", key=f"save_{chave_dia}_{h}"):
                if dados["descricao"].strip():
                    dados["usuario"] = st.session_state.usuario
                    dados["salvo"] = True
                    st.rerun()
                else:
                    st.warning("Preencha a descrição antes de salvar.")

    st.caption("ℹ️ O usuário só é exibido após o salvamento do agendamento.")


# ======================================================
# COPILOTO IA
# ======================================================
def responder_dashboard(pergunta, historico=None):
    if not DEPLOYMENT:
        return "⚠️ DEPLOYMENT não configurado. Defina AZURE_OPENAI_DEPLOYMENT."

    mensagens = [{"role": "system", "content": "Especialista em Qualidade Automotiva VW. Seja objetivo e claro."}]
    if historico:
        mensagens.extend(historico)

    mensagens.append({"role": "user", "content": pergunta})

    resp = client.chat.completions.create(
        model=DEPLOYMENT,
        messages=mensagens,
        temperature=0.3,
        max_tokens=300
    )
    return resp.choices[0].message.content

def pagina_input_budget_gmp21():

    st.subheader("🔎 Consulta de Milestone (GMP21)")

    # ==================================
    # PROJETO
    # ==================================
    modelo = st.selectbox(
        "Projeto / Plataforma",
        [
            "Plataformas - Milestone"
        ]
    )

    # ==================================
    # ANO E MÊS
    # ==================================
    col1, col2 = st.columns(2)

    with col1:
        ano = st.selectbox("Ano", [2027, 2028, 2029, 2030, 2031])

    with col2:
        mes = st.selectbox("Mês", list(range(1, 13)))

    # ==================================
    # TIPO MILESTONE
    # ==================================
    tipo_milestone = st.selectbox(
        "Tipo de Milestone",
        ["PLATAFORMA", "HUT", "MOTOR"]
    )

    mapa_milestones = {
        "PLATAFORMA": ["PM/PP", "PD/ZV", "PF", "KF", "PLF", "BF", "LF", "VFF", "PVS", "O-S", "SOP", "ME"],
        "HUT": ["PS", "PM/PP", "PD/ZV", "PF", "KF", "PLF", "BF", "LF", "VFF", "PVS", "O-S", "SOP", "ME"],
        "MOTOR": ["KB-A", "PF-A", "AE", "TF", "BF-A", "HSF", "VFF-A", "PVS-A", "O-S A", "SOP-A"]
    }

    milestone = st.selectbox("Milestone", mapa_milestones[tipo_milestone])

    # ==================================
    # 🔥 MAPA TEMPO
    # ==================================
    mapa_tempo = {

        "PLATAFORMA": {
            (2027, 3): "PM/PP",
            (2027, 6): "PD/ZV",
            (2027, 9): "PF",
            (2027, 12): "KF",
            (2028, 7): "PLF",
            (2028, 12): "BF",
            (2029, 6): "LF",
            (2030, 1): "VFF",
            (2030, 3): "PVS",
            (2030, 8): "O-S",
            (2031, 1): "SOP",
            (2031, 4): "ME",
        },

        "HUT": {
            (2027, 1): "PS",
            (2027, 6): "PM/PP",
            (2027, 9): "PD/ZV",
            (2027, 11): "PF",
            (2028, 3): "KF",
            (2028, 10): "PLF",
            (2029, 3): "BF",
            (2029, 9): "LF",
            (2030, 1): "VFF",
            (2030, 3): "PVS",
            (2030, 8): "O-S",
            (2031, 1): "SOP",
            (2031, 4): "ME",
        },

        "MOTOR": {
            (2027, 1): "KB-A",
            (2027, 5): "PF-A",
            (2027, 11): "AE",
            (2028, 7): "TF",
            (2029, 2): "BF-A",
            (2029, 9): "HSF",
            (2029, 12): "VFF-A",
            (2030, 2): "PVS-A",
            (2030, 7): "O-S A",
            (2030, 12): "SOP-A",
        }
    }

    # ==================================
    # ✅ VALIDAÇÃO
    # ==================================
    milestone_esperado = mapa_tempo.get(tipo_milestone, {}).get((ano, mes))

    if milestone_esperado:

        if milestone_esperado != milestone:
            st.warning(
                f"⚠️ Para {mes}/{ano} o correto é {milestone_esperado} ({tipo_milestone})"
            )
        else:
            st.success("✅ Milestone correto")

    else:
        st.info("ℹ️ Esse mês não possui milestone definido")



def pagina_copiloto_ia():
    st.subheader("🤖 Copiloto IA")

    with st.expander("🔎 Diagnóstico Azure OpenAI (clique para abrir)", expanded=False):
        st.write("AZURE_OPENAI_ENDPOINT:", "✅ OK" if os.getenv("AZURE_OPENAI_ENDPOINT") else "❌ VAZIO")
        st.write("AZURE_OPENAI_API_KEY:", "✅ OK" if os.getenv("AZURE_OPENAI_API_KEY") else "❌ VAZIO")
        st.write("AZURE_OPENAI_DEPLOYMENT:", os.getenv("AZURE_OPENAI_DEPLOYMENT") or "❌ VAZIO")

        if st.button("🧪 Testar conexão com Azure OpenAI"):
            try:
                test = client.chat.completions.create(
                    model=DEPLOYMENT,
                    messages=[{"role": "user", "content": "Responda apenas: OK"}],
                    max_tokens=10,
                    temperature=0
                )
                st.success("✅ Conectou! Resposta: " + test.choices[0].message.content)
            except Exception as e:
                st.error("❌ Falhou ao chamar Azure OpenAI.")
                st.exception(e)

    if "chat_messages" not in st.session_state:
        st.session_state.chat_messages = [
            {"role": "assistant", "content": "Olá! Posso ajudar com KPIs, processos e dúvidas do time de Qualidade."}
        ]

    for m in st.session_state.chat_messages:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])

    prompt = st.chat_input("Digite sua pergunta…")
    if prompt:
        st.session_state.chat_messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        try:
            with st.chat_message("assistant"):
                with st.spinner("Pensando..."):
                    historico = [x for x in st.session_state.chat_messages if x["role"] in ("user", "assistant")]
                    answer = responder_dashboard(prompt, historico=historico)
                    st.markdown(answer)

            st.session_state.chat_messages.append({"role": "assistant", "content": answer})

        except Exception as e:
            st.error("❌ Erro ao chamar o Azure OpenAI (veja detalhes abaixo).")
            st.exception(e)


# ======================================================
# COMPARATIVO EXCEL
# ======================================================
def Comparativo_Custo_Reparo_Prognose():
    st.subheader("🧮 Custo Médio de Reparo")

    st.markdown(
        """
        **Como funciona:** envie o Excel (com as abas dos projetos) e o sistema irá capturar
        o **último valor numérico** encontrado na **Coluna J** de cada aba (TOTAL) e comparar.
        
        ⚠️ Se o total em J for fórmula, **salve a planilha** antes de enviar.
        """
    )

    arquivo = st.file_uploader("📄 Upload da planilha Excel", type=["xlsx", "xls"])

    if not arquivo:
        st.info("Envie a planilha para gerar o comparativo automático.")
        return

    try:
        wb = load_workbook(io.BytesIO(arquivo.getvalue()), data_only=True)
        abas = wb.sheetnames
    except Exception as e:
        st.error("Não foi possível abrir a planilha. Verifique se o arquivo não está corrompido.")
        st.exception(e)
        return

    sugestao = [s for s in abas if str(s).startswith("203-")]
    default_sel = sugestao[:3] if len(sugestao) >= 3 else abas[:3]

    abas_sel = st.multiselect(
        "Selecione as abas (projetos) para comparar",
        options=abas,
        default=default_sel
    )

    if len(abas_sel) < 2:
        st.warning("Selecione pelo menos 2 abas para comparar.")
        return

    st.divider()

    resultados = []
    for aba in abas_sel:
        total_j = extrair_total_coluna_j_openpyxl(arquivo, aba)
        resultados.append({
            "Projeto/Aba": aba,
            "Total Coluna J": total_j
        })

    df = pd.DataFrame(resultados)

    if df["Total Coluna J"].notna().any():
        maxv = df["Total Coluna J"].max()
        minv = df["Total Coluna J"].min()
        df["Diferença p/ Máx"] = maxv - df["Total Coluna J"]
        df["Diferença p/ Mín"] = df["Total Coluna J"] - minv

    col1, col2, col3 = st.columns(3)
    col1.metric("Maior Total (J)", formatar_moeda_br(df["Total Coluna J"].max() if df["Total Coluna J"].notna().any() else None))
    col2.metric("Menor Total (J)", formatar_moeda_br(df["Total Coluna J"].min() if df["Total Coluna J"].notna().any() else None))
    col3.metric("Delta (Máx - Mín)", formatar_moeda_br(
        (df["Total Coluna J"].max() - df["Total Coluna J"].min())
        if df["Total Coluna J"].notna().any() else None
    ))

    st.markdown("### ✅ Comparativo Custo de Reparo")
    df_view = df.copy()
    df_view["Total Coluna J"] = df_view["Total Coluna J"].apply(formatar_moeda_br)
    if "Diferença p/ Máx" in df_view.columns:
        df_view["Diferença p/ Máx"] = df["Diferença p/ Máx"].apply(formatar_moeda_br)
        df_view["Diferença p/ Mín"] = df["Diferença p/ Mín"].apply(formatar_moeda_br)

    st.dataframe(df_view, use_container_width=True)

    st.divider()
    st.markdown("### 📊 Visual (Totais da Coluna J)")
    df_plot = df.dropna(subset=["Total Coluna J"]).copy()
    if len(df_plot):
        fig = px.bar(df_plot, x="Projeto/Aba", y="Total Coluna J", text="Total Coluna J")
        fig.update_traces(texttemplate="%{text:.2f}", textposition="outside")
        fig.update_layout(yaxis_title="Total Coluna J", xaxis_title="Projeto/Aba")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Não foi possível localizar valores numéricos na Coluna J das abas selecionadas.")

    st.divider()
    st.download_button(
        "Baixar comparativo (CSV)",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name="comparativo_total_coluna_J.csv",
        mime="text/csv"
    )


# ======================================================
# COMPARATIVO PDF (MIS12 / MIS36)
# ======================================================
def Comparativo_MIS_PDF():
    st.subheader("📄 Comparativo MIS12 e MIS36 (PDF)")

    st.markdown("""
**Como funciona:** envie **2 PDFs** (A e B). O sistema identifica o **título/código** do relatório
e extrai os valores **MIS12** e **MIS36** do **ano selecionado (HJ)**.

✅ Este modo funciona muito bem para os relatórios 604 (texto selecionável).  
""")

    colA, colB = st.columns(2)
    with colA:
        pdf_a = st.file_uploader("Upload PDF A", type=["pdf"], key="pdf_a")
    with colB:
        pdf_b = st.file_uploader("Upload PDF B", type=["pdf"], key="pdf_b")

    if not pdf_a or not pdf_b:
        st.info("Envie os dois PDFs para iniciar o comparativo.")
        return

    bytes_a = pdf_a.getvalue()
    bytes_b = pdf_b.getvalue()

    # Identificação
    with st.spinner("Lendo títulos/códigos dos PDFs..."):
        titulo_a = extrair_titulo_pdf(bytes_a) or "Não identificado"
        titulo_b = extrair_titulo_pdf(bytes_b) or "Não identificado"

    st.markdown("### 🏷️ Identificação dos PDFs (para diferenciar)")
    c1, c2 = st.columns(2)
    c1.info(f"**PDF A:** {titulo_a}")
    c2.info(f"**PDF B:** {titulo_b}")

    # anos disponíveis (união dos 2 PDFs)
    anos_disp = sorted(set(extrair_anos_pdf(bytes_a)) | set(extrair_anos_pdf(bytes_b)))
    if not anos_disp:
        st.error("Não foi possível identificar anos no PDF. Verifique se o PDF tem texto extraível.")
        return

    default_ano = max(anos_disp)
    ano_sel = st.selectbox("Ano para comparação (HJ)", options=anos_disp, index=anos_disp.index(default_ano))

    with st.spinner("Extraindo MIS12/MIS36 do PDF A..."):
        res_a = extrair_mis12_mis36_por_ano_pdf(bytes_a, int(ano_sel))
    with st.spinner("Extraindo MIS12/MIS36 do PDF B..."):
        res_b = extrair_mis12_mis36_por_ano_pdf(bytes_b, int(ano_sel))

    def fmt_num(x):
        if x is None:
            return "—"
        return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def delta(a, b):
        if a is None or b is None:
            return None
        return b - a

    def delta_pct(a, b):
        if a is None or b is None or a == 0:
            return None
        return (b - a) / a * 100.0

    d_mis12 = delta(res_a["MIS12"], res_b["MIS12"])
    p_mis12 = delta_pct(res_a["MIS12"], res_b["MIS12"])
    d_mis36 = delta(res_a["MIS36"], res_b["MIS36"])
    p_mis36 = delta_pct(res_a["MIS36"], res_b["MIS36"])

    if res_a["MIS12"] is None or res_a["MIS36"] is None or res_b["MIS12"] is None or res_b["MIS36"] is None:
        st.warning("Algum valor não foi encontrado para este ano em um dos PDFs (pode estar ausente no relatório).")
        st.write("Anos disponíveis PDF A:", res_a.get("anos_disponiveis", []))
        st.write("Anos disponíveis PDF B:", res_b.get("anos_disponiveis", []))

    st.divider()
    st.markdown("### ✅ Métricas (Ano selecionado)")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("MIS12 - PDF A", fmt_num(res_a["MIS12"]))
    c2.metric("MIS12 - PDF B", fmt_num(res_b["MIS12"]))
    c3.metric("Δ MIS12 (B - A)", fmt_num(d_mis12))
    c4.metric("Δ% MIS12", "—" if p_mis12 is None else f"{p_mis12:.1f}%")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("MIS36 - PDF A", fmt_num(res_a["MIS36"]))
    c2.metric("MIS36 - PDF B", fmt_num(res_b["MIS36"]))
    c3.metric("Δ MIS36 (B - A)", fmt_num(d_mis36))
    c4.metric("Δ% MIS36", "—" if p_mis36 is None else f"{p_mis36:.1f}%")

    st.divider()
    st.markdown("### 📋 Tabela")

    df = pd.DataFrame([
        {"PDF": "A", "Título": titulo_a, "Ano": int(ano_sel), "MIS12": res_a["MIS12"], "MIS36": res_a["MIS36"]},
        {"PDF": "B", "Título": titulo_b, "Ano": int(ano_sel), "MIS12": res_b["MIS12"], "MIS36": res_b["MIS36"]},
    ])

    df_view = df.copy()
    df_view["MIS12"] = df_view["MIS12"].apply(fmt_num)
    df_view["MIS36"] = df_view["MIS36"].apply(fmt_num)
    st.dataframe(df_view, use_container_width=True)

    st.markdown("### 📊 Visual")
    df_plot = df.dropna(subset=["MIS12", "MIS36"], how="all").copy()
    if len(df_plot):
        df_melt = df_plot.melt(id_vars=["PDF", "Título", "Ano"], value_vars=["MIS12", "MIS36"],
                               var_name="MIS", value_name="Valor")
        fig = px.bar(df_melt, x="MIS", y="Valor", color="PDF", barmode="group", text="Valor", hover_data=["Título"])
        fig.update_traces(texttemplate="%{text:.2f}", textposition="outside")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Sem valores suficientes para plotar.")

    st.divider()
    st.download_button(
        "Baixar comparativo (CSV)",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name=f"comparativo_MIS12_MIS36_{ano_sel}.csv",
        mime="text/csv"
    )


# ======================================================
# LOGIN
# ======================================================
def tela_login():

    # CSS
    st.markdown("""
    <style>
    .login-box {
        width: 320px;
        margin: auto;
        margin-top: 60px;
        padding: 20px;
        border-radius: 12px;

        background: rgba(255, 255, 255, 0.15);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);

        box-shadow: 0px 4px 15px rgba(0,0,0,0.25);
        text-align: center;
        border: 1px solid rgba(255,255,255,0.2);
    }

    .titulo {
        font-size: 22px;
        font-weight: bold;
        color: white;
        margin-top: 8px;
    }

    .subtitulo {
        font-size: 12px;
        color: rgba(255,255,255,0.8);
        margin-bottom: 15px;
    }

    .logo img {
        width: 45px;
    }
    </style>
    """, unsafe_allow_html=True)

    # ✅ CAIXA VISUAL (faltava isso)
    
    st.markdown("""
    <div class="login-box">
    <div class="logo">
        <img src="https://upload.wikimedia.org/wikipedia/commons/6/6d/Volkswagen_logo_2019.svg">
    </div>
    <div class="titulo">Design for Quality</div>
    <div class="subtitulo">Sistema de Qualidade</div>
    </div>
    """, unsafe_allow_html=True)


    # ✅ FORM LOGIN (mantém funcional)
    col1, col2, col3 = st.columns([3,4,3])

    with col2:
        with st.form("login"):
            user = st.text_input("Usuário REDE VW")
            pwd = st.text_input("Senha", type="password")

            if st.form_submit_button("Entrar"):
                if user.lower() in USUARIOS and USUARIOS[user.lower()] == pwd:
                    st.session_state.logado = True
                    st.session_state.usuario = user.lower()
                    st.rerun()
                else:
                    st.error("Usuário ou senha inválidos")


def pagina_links_ferramentas():
    st.subheader("🔗 Links e Ferramentas do Dia a Dia")

    recursos = [
        {
            "nome": "Power BI - KPMs",
            "url": "https://app.powerbi.com/reportEmbed?reportId=e373352c-48e7-4ea0-936e-db63d70c84b1",
            "desc": "Dashboard principal de KPIs",
            "tag": "BI"
        },
        {

            "nome": "SharePoint - Qualidade",
            "url": "https://volkswagengroup.sharepoint.com/:f:/r/sites/QAProttipos/Shared%20Documents/General?csf=1&web=1&e=eNV37D",
            "desc": "Documentos e procedimentos do time",
            "tag": "Docs"

        },
        {
            "nome": "Teams - Squad QA",
            "url": "https://SEU_LINK_AQUI",
            "desc": "Canal do time / comunicação",
            "tag": "Teams"
        },
        {
            "nome": "Pasta de Trabalho - Rede",
            "url": r"G:\ANCBQD01\S1004_B-QP_ Plan_Central Novos_Projetos\S2043_B-QP_VSC_QA_&_Eng_Prototipo\DESIGN FOR QUALITY",
            "desc": "Atalho para pasta da rede (copiar caminho)",
            "tag": "Files"
        },
    ]

    colA, colB, colC = st.columns(3)
    cols = [colA, colB, colC]

    for i, r in enumerate(recursos):
        with cols[i % 3]:
            url = r["url"]
            is_http = url.lower().startswith("http")

            link_html = (
                f'<a href="{url}" target="_blank" style="text-decoration:none;font-weight:600;">Abrir ↗</a>'
                if is_http else
                f'<span style="font-size:12px;opacity:.9;">{url}</span>'
            )

            st.markdown(
                f"""
                <div style="border-radius:16px;padding:14px;background:#1118270f;border:1px solid #e5e7eb;">
                  <div style="font-size:14px;font-weight:700;margin-bottom:6px;">{r['nome']}</div>
                  <div style="font-size:12px;opacity:.85;margin-bottom:10px;">{r['desc']}</div>
                  <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;">
                    <span style="font-size:11px;background:#11182715;padding:3px 8px;border-radius:999px;white-space:nowrap;">{r['tag']}</span>
                    <div style="text-align:right;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:260px;">
                      {link_html}
                    </div>
                  </div>
                </div>
                """,
                unsafe_allow_html=True
            )

            if not is_http:
                st.text_input("Copiar caminho:", value=url, label_visibility="collapsed", key=f"path_{i}")

# ======================================================
# PÁGINA: TEMPLATES
# ======================================================
def pagina_templates():
    st.subheader("📄 Templates e Arquivos")

    pasta = Path("templates")
    pasta.mkdir(exist_ok=True)

    arquivos = sorted(pasta.glob("*"))
    if not arquivos:
        st.info("Nenhum arquivo em /templates ainda. Coloque aqui os modelos (xlsx, pptx, pdf).")
        return

    for arq in arquivos:
        col1, col2 = st.columns([7, 2])
        col1.write(f"📌 {arq.name}")
        with col2:
            st.download_button(
                "Baixar",
                data=arq.read_bytes(),
                file_name=arq.name,
                mime="application/octet-stream",
                key=f"dl_{arq.name}"
            )

# ======================================================
# PAINEL
# ======================================================
# =========================================
# HELPERS DE NAVEGAÇÃO
# =========================================
def ir_para(pagina):
    st.session_state.pagina_atual = pagina
    st.rerun()

def botao_voltar():
    if st.button("⬅️ Voltar"):
        ir_para("HOME")


# =========================================
# PAINEL PRINCIPAL
# =========================================
def painel():

    mostrar_data_kw()
    st.title("Design for Quality")
    mostrar_carrinho_animado_painel()

    if "pagina_atual" not in st.session_state:
        st.session_state.pagina_atual = "HOME"

    if "subpagina" not in st.session_state:
        st.session_state.subpagina = None

    pagina = st.session_state.pagina_atual

    # ======================
    # CSS DOS CARDS
    # ======================
    st.markdown("""
    <style>
    .card {height: 130px; border-radius: 12px; padding: 12px; position: relative;
           box-shadow: 0px 6px 15px rgba(0,0,0,0.2);}
    .card-locked {background: #d1d5db; color: #555;}
    .card-red {background: linear-gradient(135deg, #e53935, #8e0000); color:white;}
    .card-blue {background: linear-gradient(135deg, #64b5f6, #1565c0); color:white;}
    .card-black {background: linear-gradient(135deg, #444, #000); color:white;}
    .titulo {font-size: 13px; font-weight: bold;}
    .status {position:absolute; bottom:10px; left:12px; font-size:12px;}
    .letra {position:absolute; right:10px; bottom:0px; font-size:80px;
            color:rgba(255,255,255,0.15); font-weight:bold;}
    </style>
    """, unsafe_allow_html=True)

    # ======================
    # HOME
    # ======================
    if pagina == "HOME":

        st.markdown("## 🏠 Módulos do Sistema")
        col1, col2, col3, col4, col5 = st.columns(5)

        with col1:
            st.markdown("<div class='card card-locked'>Módulo X</div>", unsafe_allow_html=True)

        with col2:
            st.markdown("<div class='card card-black'>KPM</div>", unsafe_allow_html=True)
            if st.button("", key="kpm"):
                st.session_state.pagina_atual = "KPM"
                st.rerun()

        with col3:
            st.markdown("<div class='card card-red'>GMP21</div>", unsafe_allow_html=True)
            if st.button("", key="gmp21"):
                st.session_state.pagina_atual = "GMP21"
                st.rerun()

        with col4:
            st.markdown("<div class='card card-blue'>STATUS</div>", unsafe_allow_html=True)
            if st.button("", key="status"):
                st.session_state.pagina_atual = "STATUS"
                st.rerun()

        with col5:
            st.markdown("<div class='card card-black'>ENTREGA</div>", unsafe_allow_html=True)
            if st.button("", key="entrega veiculos QA"):
                st.session_state.pagina_atual = "ENTREGA VEICULOS QA"
                st.rerun()

    # ======================
    # ENTREGA DFQ
    # ======================
    elif pagina == "ENTREGA VEICULOS QA":

        botao_voltar()

        st.subheader("Status Liberações ZP8/Rodagem 2026")

        import plotly.graph_objects as go

        # ✅ LER EXCEL
        df = pd.read_csv("dados_rodagem.csv")

        # ✅ LIMPAR COLUNAS
        df.columns = df.columns.str.strip()
        df.columns = ["Mes", "Prevista", "Liberados"]

        # ✅ LISTAS
        meses = df["Mes"].tolist()
        prevista = df["Prevista"].tolist()
        liberados = df["Liberados"].tolist()

        # ✅ TRATAR VAZIOS
        
        
        prevista = [int(v) if pd.notna(v) else None for v in prevista]
        liberados = [int(v) if pd.notna(v) else None for v in liberados]



        # ✅ GRÁFICO
        fig = go.Figure()

        fig.add_trace(go.Bar(
            name="Rodagem Prevista",
            x=meses,
            y=prevista,
            text=[v if v else "" for v in prevista],
            textposition="outside"
        ))

        fig.add_trace(go.Bar(
            name="Veículos Liberados",
            x=meses,
            y=liberados,
            text=[v if v else "" for v in liberados],
            textposition="outside"
        ))

        fig.update_layout(
            barmode='group',
            title="Performance 2026 | Total de veículos liberados",
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.2,
                xanchor="center",
                x=0.5
            )
        )

        col1, col2 = st.columns([3,1])

        with col1:
            st.plotly_chart(fig, use_container_width=True)

        with col2:

            total_prevista = sum(v for v in prevista if v is not None)
            total_liberados = sum(v for v in liberados if v is not None)

            st.markdown("### 📊 Totais")

            st.markdown(f"""
            <div style="margin-bottom:20px;">
                <div style="color:#90CAF9;">Rodagem Prevista</div>
                <div style="color:#90CAF9; font-size:28px; font-weight:bold;">
                    {total_prevista}
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.markdown(f"""
            <div style="margin-bottom:20px;">
                <div style="color:#1E88E5;">Veículos Liberados</div>
                <div style="color:#1E88E5; font-size:28px; font-weight:bold;">
                    {total_liberados}
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("---")

            st.markdown("### Programas Avaliados")
            st.markdown("""
            - VW247 Udara PLAT AGT  
            - VW247 Udara HUT AGT  
            - VW246 SSA South Africa Entry  
            - PL8 STEP III – TCROSS  
            - AQ300 GEN2 MQB27 Export  
            - M0B37W SAGA – AGT PHASE  
            - NIVUS GTS ARGENTINA  
            """)

    # ======================
    # OUTROS
    # ======================
    elif pagina == "GMP21":
        botao_voltar()
        st.subheader("GMP21 Budget")

    elif pagina == "KPM":
        botao_voltar()
        st.subheader("Dashboard KPM")

    elif pagina == "STATUS":
        botao_voltar()
        Comparativo_Custo_Reparo_Prognose()


# =========================================
# LOGOUT
# =========================================
def logout():
    st.session_state.logado = False
    st.session_state.pagina_atual = "HOME"
    st.rerun()


# =========================================
# FLUXO PRINCIPAL DO APP
# =========================================
if st.session_state.logado:
    painel()
else:
    aplicar_background_login()
    mostrar_data_kw()
    tela_login()
