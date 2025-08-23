import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date, datetime
import base64
import matplotlib.pyplot as plt

st.set_page_config(page_title="Finance MVP", layout="wide")

def brl(v): 
    return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X",".")

st.title("📊 Finance MVP — Formulário → Excel + HTML")

with st.form("form"):
    c1,c2 = st.columns(2)
    cliente = c1.text_input("Cliente", "Cliente Exemplo")
    mes_ref = c2.date_input("Mês de referência", value=date.today().replace(day=1))
    c3,c4 = st.columns(2)
    receita = c3.number_input("Receita do mês (R$)", min_value=0.0, step=100.0)
    cpv     = c4.number_input("CPV/CSV do mês (R$)", min_value=0.0, step=100.0)
    c5,c6 = st.columns(2)
    despesas = c5.number_input("Despesas operacionais (R$)", min_value=0.0, step=100.0)
    saldo_ini = c6.number_input("Saldo inicial de caixa (R$)", min_value=0.0, step=100.0)
    ok = st.form_submit_button("Gerar")
    
if ok:
    # cálculos
    lucro_bruto = receita - cpv
    ebit = lucro_bruto - despesas
    ir = ebit * 0.10 if ebit > 0 else 0.0
    lucro_liq = ebit - ir
    margem_bruta = (lucro_bruto/receita*100) if receita else 0.0
    margem_liq = (lucro_liq/receita*100) if receita else 0.0
    entradas = receita
    saidas = cpv + despesas
    saldo_final = saldo_ini + entradas - saidas

    # preview
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Receita", brl(receita))
    c2.metric("Lucro líquido", brl(lucro_liq))
    c3.metric("Margem líquida", f"{margem_liq:.1f}%")
    c4.metric("Saldo final de caixa", brl(saldo_final))

    # Excel (3 abas simples)
    df_dre = pd.DataFrame([{
        "mes": mes_ref, "receita": receita, "cpv": cpv, "lucro_bruto": lucro_bruto,
        "despesas": despesas, "ebit": ebit, "ir(10%)": ir, "lucro_liquido": lucro_liq,
        "margem_bruta_%": margem_bruta, "margem_liquida_%": margem_liq
    }])
    df_fluxo = pd.DataFrame([{
        "mes": mes_ref, "saldo_inicial": saldo_ini, "entradas": entradas,
        "saidas": saidas, "saldo_final": saldo_final
    }])
    df_kpis = pd.DataFrame([{
        "margem_bruta_%": margem_bruta, "margem_liquida_%": margem_liq
    }])

    xls = BytesIO()
    with pd.ExcelWriter(xls, engine="xlsxwriter") as w:
        df_dre.to_excel(w, index=False, sheet_name="DRE")
        df_fluxo.to_excel(w, index=False, sheet_name="Fluxo")
        df_kpis.to_excel(w, index=False, sheet_name="KPIs")
    xls.seek(0)

    # gráfico pizza pro HTML
    fig, ax = plt.subplots()
    partes = [max(cpv,0), max(despesas,0), max(lucro_liq,0)]
    labels = ["CPV/CSV", "Despesas", "Lucro Líquido"]
    ax.pie(partes, labels=labels, autopct="%1.1f%%")
    ax.set_title("Composição do Resultado")
    buf = BytesIO(); plt.tight_layout(); fig.savefig(buf, format="png", dpi=150); buf.seek(0)
    img64 = base64.b64encode(buf.read()).decode("utf-8"); plt.close(fig)

    # HTML cliente
    html = f"""
    <html><head><meta charset="utf-8">
    <style>body{{font-family:Arial;margin:24px}}
      .card{{display:inline-block;padding:12px 16px;margin:6px;border:1px solid #eee;border-radius:10px}}
    </style></head><body>
    <h1>Relatório — {cliente} ({mes_ref.strftime('%m/%Y')})</h1>
    <div class="card"><b>Receita:</b> {brl(receita)}</div>
    <div class="card"><b>Lucro Líquido:</b> {brl(lucro_liq)}</div>
    <div class="card"><b>Margem Líquida:</b> {margem_liq:.1f}%</div>
    <div class="card"><b>Caixa (saldo final):</b> {brl(saldo_final)}</div>
    <h2>Gráfico</h2>
    <img style="max-width:100%;" src="data:image/png;base64,{img64}">
    <p style="color:#666">*Gerado automaticamente pelo Finance MVP.</p>
    </body></html>
    """.encode("utf-8")

    # botões de download
    a,b = st.columns(2)
    a.download_button("⬇️ Baixar Excel (DRE/Fluxo/KPIs)", data=xls,
                      file_name=f"Finance_{cliente}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    b.download_button("⬇️ Baixar Relatório (HTML)", data=html,
                      file_name=f"Relatorio_{cliente}_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                      mime="text/html")
