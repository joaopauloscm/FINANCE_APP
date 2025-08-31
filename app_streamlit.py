# app_streamlit.py — v3 (DRE + Fluxo + KPIs + Gráficos + Alertas + Excel formatado)
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date, datetime
import numpy as np
import matplotlib.pyplot as plt

st.set_page_config(page_title="Finance v3 — DRE / Fluxo / KPIs", layout="wide")

# ----------------- Helpers -----------------
def brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
    except:
        return "R$ 0,00"

def pct(v):
    try:
        return f"{float(v):.1f}%"
    except:
        return "0.0%"

def ensure_history_df(df: pd.DataFrame):
    """Normaliza DataFrame de histórico para colunas esperadas e tipos."""
    needed = [
        "mes","receita_liq","cpv_csv","despesas",
        "entradas","saidas",
        "orcado_receita","orcado_despesas"
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=needed)
    df = df.copy()
    # normaliza nomes
    df.columns = [c.strip().lower() for c in df.columns]
    rename_map = {
        "receita": "receita_liq",
        "receita_liquida": "receita_liq",
        "cpv": "cpv_csv",
        "csv": "cpv_csv",
        "orcado_receitas": "orcado_receita",
        "orcado_despesa": "orcado_despesas",
    }
    df = df.rename(columns=rename_map)
    for c in needed:
        if c not in df.columns:
            df[c] = 0.0
    # mes: aceita YYYY-MM, YYYY/MM, ou 1º dia do mês
    df["mes"] = pd.to_datetime(df["mes"].astype(str).str.replace("/", "-")+"-01", errors="coerce")
    for c in [c for c in needed if c!="mes"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df = df[needed].dropna(subset=["mes"])
    df = df.sort_values("mes").reset_index(drop=True)
    return df

def dre_from_inputs(
    vendas_prod, prest_serv, outras_rec,
    devolucoes, descontos, imp_vendas,
    cpv_prod, csv_serv, mao_obra,
    comissoes, marketing, vendas,
    sal_admin, aluguel, utilidades, mat_escr, contab, seguros,
    deprec, provisoes,
    jur_pagos, iof, tarifas, jur_receb, rend_apl,
    ir_val, csll_val
):
    receita_bruta = vendas_prod + prest_serv + outras_rec
    deducoes = devolucoes + descontos + imp_vendas
    receita_liq = max(receita_bruta - deducoes, 0.0)
    cpv_csv = cpv_prod + csv_serv + mao_obra
    lucro_bruto = receita_liq - cpv_csv

    desp_comerciais = comissoes + marketing + vendas
    desp_adm = sal_admin + aluguel + utilidades + mat_escr + contab + seguros
    outras_oper = deprec + provisoes
    despesas_oper = desp_comerciais + desp_adm + outras_oper

    ebit = lucro_bruto - despesas_oper

    despesas_fin = jur_pagos + iof + tarifas
    receitas_fin = jur_receb + rend_apl
    lair = ebit - despesas_fin + receitas_fin

    lucro_liq = lair - ir_val - csll_val

    return {
        "receita_bruta": receita_bruta,
        "deducoes": deducoes,
        "receita_liq": receita_liq,
        "cpv_csv": cpv_csv,
        "lucro_bruto": lucro_bruto,
        "desp_comerciais": desp_comerciais,
        "desp_adm": desp_adm,
        "outras_oper": outras_oper,
        "despesas_oper": despesas_oper,
        "ebit": ebit,
        "despesas_fin": despesas_fin,
        "receitas_fin": receitas_fin,
        "lair": lair,
        "lucro_liq": lucro_liq
    }

def make_pizza_series_current_month(**kwargs):
    """Retorna dict com composição de despesas operacionais por bloco para pizza."""
    return {
        "Comerciais": kwargs.get("desp_comerciais", 0.0),
        "Administrativas": kwargs.get("desp_adm", 0.0),
        "Outras operacionais": kwargs.get("outras_oper", 0.0),
    }

# ----------------- UI -----------------
st.title("📘 DRE + 💰 Fluxo + 📈 KPIs — v3")

with st.expander("🔁 (Opcional) Carregar histórico de 12 meses para gráficos", expanded=False):
    st.caption("Formato CSV com colunas: mes,receita_liq,cpv_csv,despesas,entradas,saidas,orcado_receita,orcado_despesas")
    file = st.file_uploader("Envie seu CSV", type=["csv"])
    hist_df = None
    if file:
        try:
            hist_df = pd.read_csv(file)
        except Exception:
            try:
                hist_df = pd.read_csv(file, sep=";")
            except Exception as e:
                st.error(f"Não foi possível ler o CSV: {e}")
                hist_df = None
    hist_df = ensure_history_df(hist_df)
    if not hist_df.empty:
        st.success("Histórico carregado.")
        st.dataframe(hist_df.tail(12), use_container_width=True, hide_index=True)

with st.form("form_principal"):
    c0, c1 = st.columns(2)
    cliente = c0.text_input("Cliente", "Cliente Exemplo")
    mes_ref = c1.date_input("Mês de referência", value=date.today().replace(day=1))

    st.markdown("### 1) Receita Bruta e Deduções")
    a1, a2, a3 = st.columns(3)
    vendas_prod = a1.number_input("Vendas de produtos (R$)", min_value=0.0, step=100.0)
    prest_serv  = a2.number_input("Prestação de serviços (R$)", min_value=0.0, step=100.0)
    outras_rec  = a3.number_input("Outras receitas (R$)", min_value=0.0, step=100.0)

    b1, b2, b3 = st.columns(3)
    devolucoes = b1.number_input("Devoluções (R$)", min_value=0.0, step=50.0)
    descontos  = b2.number_input("Descontos concedidos (R$)", min_value=0.0, step=50.0)
    imp_vendas = b3.number_input("Impostos sobre vendas (R$)", min_value=0.0, step=50.0)

    st.markdown("### 2) CPV / CSV")
    c1, c2, c3 = st.columns(3)
    cpv_prod   = c1.number_input("Custo dos produtos vendidos (R$)", min_value=0.0, step=100.0)
    csv_serv   = c2.number_input("Custo dos serviços prestados (R$)", min_value=0.0, step=100.0)
    mao_obra   = c3.number_input("Mão de obra direta (R$)", min_value=0.0, step=100.0)

    st.markdown("### 3) Despesas Operacionais")
    d1, d2, d3 = st.columns(3)
    comissoes  = d1.number_input("Comissões (R$)", min_value=0.0, step=50.0)
    marketing  = d2.number_input("Marketing (R$)", min_value=0.0, step=50.0)
    vendas     = d3.number_input("Outras de Vendas (R$)", min_value=0.0, step=50.0)

    e1, e2, e3 = st.columns(3)
    sal_admin  = e1.number_input("Salários administrativos (R$)", min_value=0.0, step=100.0)
    aluguel    = e2.number_input("Aluguel (R$)", min_value=0.0, step=100.0)
    utilidades = e3.number_input("Energia/Água/Telefone (R$)", min_value=0.0, step=50.0)

    e4, e5, e6 = st.columns(3)
    mat_escr   = e4.number_input("Material de escritório (R$)", min_value=0.0, step=50.0)
    contab     = e5.number_input("Contabilidade (R$)", min_value=0.0, step=50.0)
    seguros    = e6.number_input("Seguros (R$)", min_value=0.0, step=50.0)

    f1, f2 = st.columns(2)
    deprec    = f1.number_input("Depreciação (R$)", min_value=0.0, step=50.0)
    provisoes = f2.number_input("Provisões (R$)", min_value=0.0, step=50.0)

    st.markdown("### 4) Resultado Financeiro")
    g1, g2, g3 = st.columns(3)
    jur_pagos  = g1.number_input("Juros pagos (R$)", min_value=0.0, step=50.0)
    iof        = g2.number_input("IOF (R$)", min_value=0.0, step=10.0)
    tarifas    = g3.number_input("Tarifas bancárias (R$)", min_value=0.0, step=10.0)
    h1, h2 = st.columns(2)
    jur_receb  = h1.number_input("Juros recebidos (R$)", min_value=0.0, step=10.0)
    rend_apl   = h2.number_input("Rendimentos de aplicações (R$)", min_value=0.0, step=10.0)

    st.markdown("### 5) Tributos sobre o Lucro (IR/CSLL)")
    i1, i2 = st.columns(2)
    ir_val   = i1.number_input("Imposto de Renda (R$)", min_value=0.0, step=100.0)
    csll_val = i2.number_input("Contribuição Social (R$)", min_value=0.0, step=100.0)

    st.markdown("### 6) Fluxo de Caixa — mês atual")
    j1, j2, j3 = st.columns(3)
    saldo_inicial = j1.number_input("Saldo inicial (R$)", min_value=0.0, step=100.0)
    recebimentos  = j2.number_input("Recebimentos de vendas (R$)", min_value=0.0, step=100.0)
    outras_ent    = j3.number_input("Outras entradas (R$)", min_value=0.0, step=100.0)

    k1, k2, k3 = st.columns(3)
    saidas_forn  = k1.number_input("Pagamento a fornecedores (R$)", min_value=0.0, step=100.0)
    salarios     = k2.number_input("Salários e encargos (R$)", min_value=0.0, step=100.0)
    impostos     = k3.number_input("Impostos (R$)", min_value=0.0, step=100.0)

    k4, k5 = st.columns(2)
    outras_saidas = k4.number_input("Outras saídas (R$)", min_value=0.0, step=100.0)
    emprest_par   = k5.number_input("Empréstimos (parcelas) (R$)", min_value=0.0, step=100.0)

    st.markdown("### 7) Orçado do mês")
    o1, o2 = st.columns(2)
    orcado_receita_mes = o1.number_input("Orçado de RECEITA (R$)", min_value=0.0, step=100.0)
    orcado_despesas_mes = o2.number_input("Orçado de DESPESAS (R$)", min_value=0.0, step=100.0)

    st.markdown("### 8) Alertas — parâmetros")
    p1, p2 = st.columns(2)
    limiar_margem = p1.slider("Margem líquida mínima (%)", min_value=0, max_value=100, value=10)
    considerar_hist = p2.checkbox("Usar histórico para gráficos/projeção (se enviado)", value=True)

    ok = st.form_submit_button("Calcular, Gerar Gráficos e Exportar Excel")

# ----------------- CÁLCULOS -----------------
if ok:
    # DRE mês atual
    dre = dre_from_inputs(
        vendas_prod, prest_serv, outras_rec,
        devolucoes, descontos, imp_vendas,
        cpv_prod, csv_serv, mao_obra,
        comissoes, marketing, vendas,
        sal_admin, aluguel, utilidades, mat_escr, contab, seguros,
        deprec, provisoes,
        jur_pagos, iof, tarifas, jur_receb, rend_apl,
        ir_val, csll_val
    )

    # Fluxo mês atual
    entradas = recebimentos + outras_ent
    # Saídas do caixa (operacionais + empréstimos + etc.)
    saidas = saidas_forn + salarios + impostos + outras_saidas + emprest_par
    saldo_final = saldo_inicial + entradas - saidas

    # KPIs/margens
    receita_liq = dre["receita_liq"]
    margem_bruta = (dre["lucro_bruto"]/receita_liq*100) if receita_liq else 0.0
    margem_oper  = (dre["ebit"]/receita_liq*100) if receita_liq else 0.0
    margem_liq   = (dre["lucro_liq"]/receita_liq*100) if receita_liq else 0.0

    # Base de histórico (para gráficos)
    db = pd.DataFrame(columns=[
        "mes","receita_liq","cpv_csv","despesas","lucro_liq",
        "entradas","saidas","delta_caixa","acumulado",
        "orcado_receita","orcado_despesas","margem_liq_pct"
    ])

    # Se veio histórico e optou por considerar:
    if considerar_hist and not (hist_df is None or hist_df.empty):
        db = ensure_history_df(hist_df)
        # calcula lucro_liq aproximado se não existir (aqui usamos receita_liq - cpv - despesas - IR/CSLL ~ 0)
        if "lucro_liq" not in db.columns:
            db["lucro_liq"] = db["receita_liq"] - db["cpv_csv"] - db["despesas"]
        # entradas/saidas faltando -> aproxima
        db["entradas"] = np.where(db["entradas"]==0, db["receita_liq"], db["entradas"])
        db["saidas"] = np.where(db["saidas"]==0, db["cpv_csv"] + db["despesas"], db["saidas"])
        db["delta_caixa"] = db["entradas"] - db["saidas"]
        db["acumulado"] = db["delta_caixa"].cumsum()
        db["margem_liq_pct"] = np.where(db["receita_liq"]>0, db["lucro_liq"]/db["receita_liq"]*100, 0.0)

    # Registra/atualiza mês atual na base
    linha_atual = pd.DataFrame([{
        "mes": pd.to_datetime(mes_ref),
        "receita_liq": receita_liq,
        "cpv_csv": dre["cpv_csv"],
        "despesas": dre["despesas_oper"],
        "lucro_liq": dre["lucro_liq"],
        "entradas": entradas,
        "saidas": saidas,
        "delta_caixa": entradas - saidas,
        "orcado_receita": orcado_receita_mes,
        "orcado_despesas": orcado_despesas_mes,
    }])
    # Atualiza/insere
    if db.empty:
        db = linha_atual
    else:
        db = db[db["mes"] != pd.to_datetime(mes_ref)]
        db = pd.concat([db, linha_atual], ignore_index=True)
    db = db.sort_values("mes").reset_index(drop=True)
    db["acumulado"] = db["delta_caixa"].cumsum()
    db["margem_liq_pct"] = np.where(db["receita_liq"]>0, db["lucro_liq"]/db["receita_liq"]*100, 0.0)

    # --------- ALERTAS ----------
    alertas = []
    if margem_liq < limiar_margem:
        alertas.append(f"Margem líquida ({margem_liq:.1f}%) abaixo de {limiar_margem}%")
    if orcado_despesas_mes > 0 and dre["despesas_oper"] > orcado_despesas_mes:
        alertas.append("Despesas do mês ACIMA do orçado.")
    if saldo_final < 0:
        alertas.append("Fluxo de caixa NEGATIVO (saldo final < 0).")

    if alertas:
        for a in alertas: st.error("⚠️ " + a)
    else:
        st.success("✅ Sem alertas no mês.")

    # --------- CARDS (tela) ----------
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Receita Líquida", brl(receita_liq))
    c2.metric("Lucro Líquido", brl(dre["lucro_liq"]))
    c3.metric("Margem Líquida", pct(margem_liq))
    c4.metric("Caixa (saldo final)", brl(saldo_final))

    # --------- GRÁFICOS (tela) ----------
    # Evolução mensal do faturamento (12m) & Margem de lucro por mês
    if len(db) >= 1:
        show = db.tail(12).copy()
        show["mes_label"] = show["mes"].dt.strftime("%b/%y")
        # 1) Receita (linha)
        fig1, ax1 = plt.subplots()
        ax1.plot(show["mes_label"], show["receita_liq"], marker="o")
        ax1.set_title("Evolução mensal do faturamento (12 meses)")
        ax1.set_xlabel("Mês"); ax1.set_ylabel("Receita líquida (R$)")
        st.pyplot(fig1)

        # 2) Margem líquida por mês (linha)
        fig2, ax2 = plt.subplots()
        ax2.plot(show["mes_label"], show["margem_liq_pct"], marker="o")
        ax2.set_title("Margem de lucro por mês (%)")
        ax2.set_xlabel("Mês"); ax2.set_ylabel("%")
        st.pyplot(fig2)

        # 3) Fluxo de caixa acumulado (linha)
        fig3, ax3 = plt.subplots()
        ax3.plot(show["mes_label"], show["acumulado"], marker="o")
        ax3.set_title("Fluxo de caixa acumulado")
        ax3.set_xlabel("Mês"); ax3.set_ylabel("R$")
        st.pyplot(fig3)

        # 4) Orçado vs realizado (colunas) — receita
        fig4, ax4 = plt.subplots()
        width = 0.35
        x = np.arange(len(show))
        ax4.bar(x - width/2, show["orcado_receita"], width, label="Orçado")
        ax4.bar(x + width/2, show["receita_liq"], width, label="Realizado")
        ax4.set_xticks(x, show["mes_label"])
        ax4.set_title("Comparativo orçado vs realizado (Receita)")
        ax4.set_xlabel("Mês"); ax4.set_ylabel("R$")
        ax4.legend()
        st.pyplot(fig4)

    # 5) Pizza — composição de despesas (blocos) do mês atual
    pizza = make_pizza_series_current_month(**dre)
    fig5, ax5 = plt.subplots()
    labels = list(pizza.keys())
    valores = [pizza[k] for k in labels]
    if sum(valores) == 0:
        valores = [1,1,1]
    ax5.pie(valores, labels=labels, autopct="%1.1f%%")
    ax5.set_title("Composição de despesas operacionais (mês)")
    st.pyplot(fig5)

    # --------- PROJEÇÃO DE DESPESA (MM3) ----------
    proj_despesa = 0.0
    if len(db) >= 2:
        base = db.copy()
        # usar últimos 3 meses (sem o próximo) — inclui mês atual na média se houver <3 anteriores
        proj_despesa = base["despesas"].tail(3).mean()
    st.info(f"🧮 Projeção de despesa (próximo mês, MM3): {brl(proj_despesa)}")

    # --------- TABELAS (tela) ----------
    st.markdown("#### DRE do mês (resumo)")
    dre_rows = [
        ("Receita Bruta", dre["receita_bruta"]),
        ("(-) Deduções (dev/desc/impostos)", -dre["deducoes"]),
        ("= Receita Líquida", receita_liq),
        ("(-) CPV/CSV", -dre["cpv_csv"]),
        ("= Lucro Bruto", dre["lucro_bruto"]),
        ("(-) Desp. Comerciais", -dre["desp_comerciais"]),
        ("(-) Desp. Administrativas", -dre["desp_adm"]),
        ("(-) Outras Desp. Operacionais", -dre["outras_oper"]),
        ("= EBIT (Operacional)", dre["ebit"]),
        ("(-) Despesas Financeiras", -dre["despesas_fin"]),
        ("(+) Receitas Financeiras", dre["receitas_fin"]),
        ("= LAIR", dre["lair"]),
        ("(-) IR", -ir_val),
        ("(-) CSLL", -csll_val),
        ("= Lucro Líquido", dre["lucro_liq"]),
    ]
    df_dre_table = pd.DataFrame(dre_rows, columns=["Conta","Valor (R$)"])
    st.dataframe(df_dre_table, use_container_width=True, hide_index=True)

    st.markdown("#### Banco de Dados (mensal)")
    db_show = db.copy()
    db_show["mes"] = db_show["mes"].dt.strftime("%Y-%m")
    st.dataframe(db_show.tail(12), use_container_width=True, hide_index=True)

    # ----------------- EXPORTAR EXCEL FORMATADO -----------------
    def export_excel_formatted():
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
            wb = writer.book

            # Formats
            fmt_title = wb.add_format({"bold": True, "font_size": 16})
            fmt_sub = wb.add_format({"bold": True, "font_size": 12})
            fmt_card = wb.add_format({"border":1, "align":"left", "valign":"vcenter"})
            fmt_money = wb.add_format({"num_format": "R$ #,##0.00"})
            fmt_pct = wb.add_format({"num_format": "0.0%"})
            fmt_head = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border":1})
            fmt_cell = wb.add_format({"border":1})

            # ---- Banco de Dados ----
            db_to_write = db.copy()
            ws_db = wb.add_worksheet("Banco de Dados")
            cols = ["mes","receita_liq","cpv_csv","despesas","lucro_liq","entradas","saidas","delta_caixa","acumulado","orcado_receita","orcado_despesas","margem_liq_pct"]
            ws_db.write_row(0,0,["Mês","Receita Líq.","CPV/CSV","Despesas","Lucro Líq.","Entradas","Saídas","Delta Caixa","Acumulado","Orçado Receita","Orçado Despesas","Margem Líq %"], fmt_head)
            for i, row in enumerate(db_to_write[cols].itertuples(index=False), start=1):
                ws_db.write_datetime(i, 0, row[0].to_pydatetime(), fmt_cell)
                for j, val in enumerate(row[1:], start=1):
                    if j in [1,2,3,4,5,6,7,8,9]:  # money cols
                        ws_db.write_number(i, j, float(val), fmt_money)
                    elif j == 10:  # margem %
                        ws_db.write_number(i, j, float(val)/100.0, fmt_pct)
                    else:
                        ws_db.write(i, j, val, fmt_cell)
            ws_db.set_column(0, 0, 12)
            ws_db.set_column(1, 10, 18)

            # ---- DRE ----
            ws_dre = wb.add_worksheet("DRE")
            ws_dre.write("A1", f"DRE — {cliente} ({mes_ref.strftime('%m/%Y')})", fmt_title)
            ws_dre.write_row(2,0,["Conta","Valor (R$)"], fmt_head)
            for i,(conta, valor) in enumerate(dre_rows, start=3):
                ws_dre.write(i,0,conta, fmt_cell)
                ws_dre.write_number(i,1, float(valor), fmt_money)
            ws_dre.set_column(0,0,36)
            ws_dre.set_column(1,1,18)

            # ---- Fluxo ----
            ws_fx = wb.add_worksheet("Fluxo")
            ws_fx.write("A1", f"Fluxo de Caixa — {cliente} ({mes_ref.strftime('%m/%Y')})", fmt_title)
            ws_fx.write_row(2,0,["Saldo inicial","Entradas","Saídas","Saldo final"], fmt_head)
            ws_fx.write_number(3,0, float(saldo_inicial), fmt_money)
            ws_fx.write_number(3,1, float(entradas), fmt_money)
            ws_fx.write_number(3,2, float(saidas), fmt_money)
            ws_fx.write_number(3,3, float(saldo_final), fmt_money)

            # ---- KPIs ----
            ws_k = wb.add_worksheet("KPIs")
            ws_k.write("A1", "KPIs", fmt_title)
            ws_k.write_row(2,0,["Indicador","Valor"], fmt_head)
            kpis = [
                ("Margem Bruta", margem_bruta/100.0, "pct"),
                ("Margem Operacional", margem_oper/100.0, "pct"),
                ("Margem Líquida", margem_liq/100.0, "pct"),
                ("Projeção de Despesa (MM3)", proj_despesa, "money"),
            ]
            for i,(nome,val,t) in enumerate(kpis, start=3):
                ws_k.write(i,0,nome, fmt_cell)
                if t=="pct":
                    ws_k.write_number(i,1, float(val), fmt_pct)
                else:
                    ws_k.write_number(i,1, float(val), fmt_money)
            ws_k.set_column(0,0,36); ws_k.set_column(1,1,22)

            # ---- Orçado ----
            ws_orc = wb.add_worksheet("Orçado")
            ws_orc.write("A1", "Orçado (mensal)", fmt_title)
            ws_orc.write_row(2,0,["Mês","Orçado Receita","Orçado Despesas"], fmt_head)
            for i,row in enumerate(db_to_write[["mes","orcado_receita","orcado_despesas"]].itertuples(index=False), start=3):
                ws_orc.write_datetime(i,0,row[0].to_pydatetime(), fmt_cell)
                ws_orc.write_number(i,1,float(row[1]), fmt_money)
                ws_orc.write_number(i,2,float(row[2]), fmt_money)
            ws_orc.set_column(0,0,12); ws_orc.set_column(1,2,20)

            # ---- Início (cards + alertas) ----
            ws_ini = wb.add_worksheet("Início")
            ws_ini.write("A1", f"Relatório — {cliente} ({mes_ref.strftime('%m/%Y')})", fmt_title)
            ws_ini.write("A3", "Resumo do mês", fmt_sub)
            cards = [
                ("Receita Líquida", receita_liq),
                ("Lucro Líquido", dre["lucro_liq"]),
                ("Margem Líquida", margem_liq/100.0, "pct"),
                ("Caixa (Saldo final)", saldo_final),
            ]
            row = 4
            for nome, val, *t in cards:
                ws_ini.write(row,0,nome, fmt_card)
                if t and t[0]=="pct":
                    ws_ini.write_number(row,1,float(val), fmt_pct)
                else:
                    ws_ini.write_number(row,1,float(val), fmt_money)
                row += 1
            ws_ini.write("A9","Alertas", fmt_sub)
            if alertas:
                for i,a in enumerate(alertas, start=10):
                    ws_ini.write(i,0,"⚠️ "+a)
            else:
                ws_ini.write(10,0,"Nenhum alerta no mês.")

            ws_ini.write("A13","Projeção de Despesa (MM3):", fmt_sub)
            ws_ini.write_number(13,1,float(proj_despesa), fmt_money)
            ws_ini.set_column(0,0,36); ws_ini.set_column(1,1,24)

            # ---- Dash (gráficos em Excel) ----
            ws_dash = wb.add_worksheet("Dash")
            ws_dash.write("A1", "Dashboard", fmt_title)

            # Preparar ranges
            n = len(db_to_write)
            # Gráfico 1: evolução faturamento
            chart1 = wb.add_chart({"type":"line"})
            chart1.add_series({
                "name": "Receita Líquida",
                "categories": ["Banco de Dados", 1, 0, n, 0],
                "values": ["Banco de Dados", 1, 1, n, 1],
            })
            chart1.set_title({"name":"Evolução mensal do faturamento"})
            chart1.set_x_axis({"num_format":"mmm/yy"})
            chart1.set_y_axis({"num_format":"R$ #,##0"})
            ws_dash.insert_chart("A3", chart1, {"x_scale":1.2, "y_scale":1.2})

            # Gráfico 2: margem líquida por mês
            chart2 = wb.add_chart({"type":"line"})
            chart2.add_series({
                "name": "Margem Líq. (%)",
                "categories": ["Banco de Dados", 1, 0, n, 0],
                "values": ["Banco de Dados", 1, 11, n, 11],  # margem_liq_pct
                "y2_axis": False,
            })
            chart2.set_title({"name":"Margem de lucro por mês"})
            chart2.set_x_axis({"num_format":"mmm/yy"})
            chart2.set_y_axis({"num_format":"0.0%"})
            ws_dash.insert_chart("I3", chart2, {"x_scale":1.2, "y_scale":1.2})

            # Gráfico 3: fluxo de caixa acumulado
            chart3 = wb.add_chart({"type":"line"})
            chart3.add_series({
                "name":"Acumulado",
                "categories": ["Banco de Dados", 1, 0, n, 0],
                "values": ["Banco de Dados", 1, 8, n, 8],  # acumulado
            })
            chart3.set_title({"name":"Fluxo de caixa acumulado"})
            chart3.set_x_axis({"num_format":"mmm/yy"})
            chart3.set_y_axis({"num_format":"R$ #,##0"})
            ws_dash.insert_chart("A20", chart3, {"x_scale":1.2, "y_scale":1.2})

            # Gráfico 4: orçado vs realizado (Receita)
            chart4 = wb.add_chart({"type":"column"})
            chart4.add_series({
                "name":"Orçado",
                "categories":["Banco de Dados", 1,0, n,0],
                "values":["Banco de Dados", 1,9, n,9],  # orcado_receita
            })
            chart4.add_series({
                "name":"Realizado",
                "categories":["Banco de Dados", 1,0, n,0],
                "values":["Banco de Dados", 1,1, n,1],  # receita_liq
            })
            chart4.set_title({"name":"Comparativo orçado vs realizado (Receita)"})
            chart4.set_x_axis({"num_format":"mmm/yy"})
            chart4.set_y_axis({"num_format":"R$ #,##0"})
            ws_dash.insert_chart("I20", chart4, {"x_scale":1.2, "y_scale":1.2})

            # Gráfico 5: pizza composição de despesas operacionais (mês atual)
            # escreve mini-tabela suporte
            ws_dash.write_row("A38", ["Bloco", "Valor"], fmt_head)
            pie_blocks = list(make_pizza_series_current_month(**dre).items())
            for i,(nome,val) in enumerate(pie_blocks, start=39):
                ws_dash.write(i,0,nome, fmt_cell)
                ws_dash.write_number(i,1,float(val), fmt_money)
            chart5 = wb.add_chart({"type":"pie"})
            chart5.add_series({
                "name": "Composição Despesas Operacionais",
                "categories": ["Dash", 38, 0, 38+len(pie_blocks), 0],
                "values": ["Dash", 38, 1, 38+len(pie_blocks), 1],
            })
            chart5.set_title({"name":"Composição de despesas (mês)"})
            ws_dash.insert_chart("G38", chart5, {"x_scale":1.2, "y_scale":1.2})

        output.seek(0)
        return output

    # Botão de download do Excel
    excel_bytes = export_excel_formatted()
    st.download_button(
        "⬇️ Baixar Excel (Início, Dash, Banco de Dados, DRE, Fluxo, KPIs, Orçado)",
        data=excel_bytes,
        file_name=f"Finance_v3_{cliente}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Relatório HTML simples (mesmo conteúdo dos cards + DRE e alertas)
    html = f"""
    <html><head><meta charset="utf-8"><style>
    body{{font-family:Arial;margin:24px}}
    .card{{display:inline-block;margin:6px;padding:10px;border:1px solid #eee;border-radius:10px}}
    table{{border-collapse:collapse;width:100%}} td,th{{border:1px solid #eee;padding:8px}}
    </style></head><body>
    <h1>Relatório — {cliente} ({mes_ref.strftime('%m/%Y')})</h1>
    <div class="card"><b>Receita Líquida:</b> {brl(receita_liq)}</div>
    <div class="card"><b>Lucro Líquido:</b> {brl(dre['lucro_liq'])}</div>
    <div class="card"><b>Margem Líquida:</b> {margem_liq:.1f}%</div>
    <div class="card"><b>Caixa (saldo final):</b> {brl(saldo_final)}</div>
    <h2>Alertas</h2>
    <ul>{"".join([f"<li>{a}</li>" for a in alertas]) if alertas else "<li>Sem alertas.</li>"}</ul>
    <h2>DRE</h2>
    {pd.DataFrame(dre_rows, columns=['Conta','Valor (R$)']).to_html(index=False)}
    <p style="color:#666">*Gerado automaticamente.</p>
    </body></html>
    """.encode("utf-8")

    st.download_button(
        "⬇️ Baixar Relatório (HTML)",
        data=html,
        file_name=f"Relatorio_{cliente}_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
        mime="text/html"
    )
# ----------------- FIM -----------------