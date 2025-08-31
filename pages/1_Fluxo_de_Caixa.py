import streamlit as st
import pandas as pd
import altair as alt
from datetime import date, datetime
from pathlib import Path

# =========================
# CONFIG & HELPERS
# =========================
DATA_PATH = Path("data/transacoes.csv")
CATEG_PADRAO = [
    "Vendas", "Serviços", "Salário", "Investimentos",
    "Aluguel", "Marketing", "Folha", "Infra", "Impostos", "Outros"
]

@st.cache_data
def load_transacoes() -> pd.DataFrame:
    if DATA_PATH.exists():
        df = pd.read_csv(DATA_PATH, parse_dates=["data"], dayfirst=True)
    else:
        df = pd.DataFrame(columns=[
            "data","tipo","categoria","descricao","valor","conta","pago"
        ])
    # dtypes
    if not df.empty:
        df["pago"] = df["pago"].astype(bool)
        df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0.0)
        df["data"]  = pd.to_datetime(df["data"])
    return df.sort_values("data")

def save_transacoes(df: pd.DataFrame):
    DATA_PATH.parent.mkdir(parents=True, exist_ok=True)
    out = df.copy()
    out["data"] = pd.to_datetime(out["data"]).dt.strftime("%d/%m/%Y")
    out.to_csv(DATA_PATH, index=False)
    load_transacoes.clear()  # limpa cache

def _fmt_currency(x): 
    try: 
        return f"R$ {float(x):,.2f}".replace(",", ".")
    except: 
        return x

# =========================
# UI PRINCIPAL
# =========================
def fluxo_de_caixa_ui():
    st.header("📅 Fluxo de Caixa — Lançamentos por Data")

    # ------- NOVO LANÇAMENTO
    with st.expander("➕ Novo lançamento", expanded=True):
        c1, c2, c3, c4 = st.columns([1.2, 1, 1, 1.2])
        data_i     = c1.date_input("Data", value=date.today())
        tipo_i     = c2.selectbox("Tipo", ["Receita","Despesa"])
        categoria_i= c3.selectbox("Categoria", sorted(set(CATEG_PADRAO)))
        valor_i    = c4.number_input("Valor (R$)", min_value=0.0, step=10.0, format="%.2f")

        c5, c6 = st.columns([2,1])
        descricao_i= c5.text_input("Descrição", placeholder="ex.: Venda pedido #1234")
        conta_i    = c6.text_input("Conta", value="Geral")

        c7, c8, c9 = st.columns([1,1,2])
        pago_i        = c7.checkbox("Pago?", value=True)
        recorrente    = c8.checkbox("Recorrente mensal?")
        meses_qtd     = c9.number_input("Meses (se recorrente)", value=1, min_value=1, max_value=60)

        if st.button("Adicionar"):
            df = load_transacoes()
            linhas = []
            if recorrente:
                d = pd.to_datetime(data_i)
                for k in range(meses_qtd):
                    linhas.append({
                        "data": (d + pd.DateOffset(months=k)).date(),
                        "tipo": tipo_i, "categoria": categoria_i,
                        "descricao": descricao_i if k == 0 else f"{descricao_i} (M{k+1}/{meses_qtd})",
                        "valor": valor_i, "conta": conta_i, "pago": pago_i
                    })
            else:
                linhas.append({
                    "data": pd.to_datetime(data_i).date(),
                    "tipo": tipo_i, "categoria": categoria_i,
                    "descricao": descricao_i, "valor": valor_i,
                    "conta": conta_i, "pago": pago_i
                })
            novo = pd.DataFrame(linhas)
            novo["data"] = pd.to_datetime(novo["data"])
            save_transacoes(pd.concat([df, novo], ignore_index=True))
            st.success(f"{len(linhas)} lançamento(s) adicionado(s).")

    # ------- PLANILHA EDITÁVEL
    st.subheader("🧾 Planilha de lançamentos (clique para editar)")
    df = load_transacoes()
    editor = st.data_editor(
        df, num_rows="dynamic", height=360,
        column_config={
            "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
            "tipo": st.column_config.SelectboxColumn("Tipo", options=["Receita","Despesa"]),
            "categoria": st.column_config.SelectboxColumn("Categoria", options=sorted(set(CATEG_PADRAO + df["categoria"].dropna().astype(str).tolist()))),
            "descricao": st.column_config.TextColumn("Descrição"),
            "valor": st.column_config.NumberColumn("Valor (R$)", step=10.0, format="%.2f"),
            "conta": st.column_config.TextColumn("Conta"),
            "pago": st.column_config.CheckboxColumn("Pago?")
        }
    )
    cSave, cDel = st.columns([1,1])
    if cSave.button("💾 Salvar alterações"):
        save_transacoes(editor)
        st.success("Dados salvos.")
    if cDel.button("🗑️ Apagar todos os lançamentos (cuidado!)"):
        save_transacoes(pd.DataFrame(columns=editor.columns))
        st.warning("Base zerada.")

    # ------- FILTROS
    st.subheader("🔎 Filtros")
    df = load_transacoes()
    if df.empty:
        st.info("Sem lançamentos ainda. Adicione acima.")
        return

    colf1, colf2, colf3, colf4 = st.columns(4)
    ini = colf1.date_input("De", value=df["data"].min().date())
    fim = colf2.date_input("Até", value=df["data"].max().date())
    tipo_sel = colf3.multiselect("Tipo", ["Receita","Despesa"], default=["Receita","Despesa"])
    contas   = ["(todas)"] + sorted(df["conta"].dropna().unique().tolist())
    conta_sel= colf4.selectbox("Conta", contas, index=0)

    cat_opts = ["(todas)"] + sorted(df["categoria"].dropna().unique().tolist())
    cat_sel = st.selectbox("Categoria", cat_opts, index=0)

    f = df.copy()
    f = f[(f["data"].dt.date >= ini) & (f["data"].dt.date <= fim)]
    f = f[f["tipo"].isin(tipo_sel)]
    if conta_sel != "(todas)":
        f = f[f["conta"] == conta_sel]
    if cat_sel != "(todas)":
        f = f[f["categoria"] == cat_sel]

    # ------- RESUMO MENSAL + DESTAQUES
    f["ano_mes"] = f["data"].dt.to_period("M").astype(str)
    entradas = f.loc[f["tipo"]=="Receita"].groupby("ano_mes")["valor"].sum()
    saidas   = f.loc[f["tipo"]=="Despesa"].groupby("ano_mes")["valor"].sum()
    resumo   = pd.concat([entradas.rename("Entradas"), saidas.rename("Saídas")], axis=1).fillna(0)
    resumo["Fluxo Líquido"] = resumo["Entradas"] - resumo["Saídas"]

    c1,c2,c3,c4 = st.columns(4)
    if not resumo.empty:
        mes_mais_conta   = resumo["Saídas"].idxmax()
        mes_menos_conta  = resumo["Saídas"].idxmin()
        mes_mais_ganho   = resumo["Entradas"].idxmax()
        mes_menos_ganho  = resumo["Entradas"].idxmin()

        c1.metric("🟥 Mês com MAIS contas",   mes_mais_conta,  _fmt_currency(resumo.loc[mes_mais_conta,"Saídas"]))
        c2.metric("🟦 Mês com MENOS contas",  mes_menos_conta, _fmt_currency(resumo.loc[mes_menos_conta,"Saídas"]))
        c3.metric("🟩 Mês com MAIS ganhos",   mes_mais_ganho,  _fmt_currency(resumo.loc[mes_mais_ganho,"Entradas"]))
        c4.metric("🟨 Mês com MENOS ganhos",  mes_menos_ganho, _fmt_currency(resumo.loc[mes_menos_ganho,"Entradas"]))

    # ------- GRÁFICOS (Altair)
    st.markdown("### 📊 Entradas vs Saídas por mês")
    df_bar = resumo.reset_index().rename(columns={"index":"Mês"})
    df_bar = df_bar.melt("ano_mes", value_vars=["Entradas","Saídas"], var_name="Tipo", value_name="Valor")
    chart_bar = alt.Chart(df_bar).mark_bar().encode(
        x=alt.X("ano_mes:N", title="Mês"),
        y=alt.Y("Valor:Q", title="R$"),
        color=alt.Color("Tipo:N", scale=alt.Scale(scheme="tableau10"))
    ).properties(height=300)
    st.altair_chart(chart_bar, use_container_width=True)

    st.markdown("### 📈 Fluxo Líquido (Entradas - Saídas)")
    df_line = resumo.reset_index()
    chart_line = alt.Chart(df_line).mark_line(point=True).encode(
        x=alt.X("ano_mes:N", title="Mês"),
        y=alt.Y("Fluxo Líquido:Q", title="R$"),
        tooltip=["ano_mes","Fluxo Líquido"]
    ).properties(height=280)
    st.altair_chart(chart_line, use_container_width=True)

    st.markdown("### 🍩 Despesas por categoria")
    despesas_cat = f[f["tipo"]=="Despesa"].groupby("categoria")["valor"].sum().reset_index()
    if not despesas_cat.empty:
        chart_donut = alt.Chart(despesas_cat).mark_arc(innerRadius=60).encode(
            theta="valor:Q",
            color=alt.Color("categoria:N", legend=alt.Legend(title="Categoria"), scale=alt.Scale(scheme="tableau10")),
            tooltip=["categoria","valor"]
        ).properties(height=300)
        st.altair_chart(chart_donut, use_container_width=True)
    else:
        st.info("Sem despesas no período filtrado.")

    # ------- TABELA RESUMO & EXPORT
    st.markdown("### 📄 Lançamentos filtrados")
    st.dataframe(
        f.sort_values("data")
         .assign(data=lambda d: d["data"].dt.strftime("%d/%m/%Y"))
         .rename(columns={"data":"Data","tipo":"Tipo","categoria":"Categoria","descricao":"Descrição",
                          "valor":"Valor (R$)","conta":"Conta","pago":"Pago"})
    , use_container_width=True, height=260)

    csv = f.sort_values("data").assign(data=lambda d: d["data"].dt.strftime("%d/%m/%Y"))
    st.download_button("⬇️ Baixar lançamentos do período (CSV)",
                       csv.to_csv(index=False).encode("utf-8"),
                       "lancamentos_periodo.csv", "text/csv")

# Se estiver usando como página separada no Streamlit multipage:
if __name__ == "__main__":
    fluxo_de_caixa_ui()
