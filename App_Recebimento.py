import os, re, unicodedata
from pathlib import Path
from urllib.parse import quote
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Controle de Recebimento de Materiais", page_icon="📦", layout="wide")
APP_DIR = Path(__file__).resolve().parent
BASE_XLSX = APP_DIR / "Controle de Recebimento de Materiais.xlsx"

def norm(v):
    s="" if pd.isna(v) else str(v)
    return unicodedata.normalize("NFKD",s).encode("ascii","ignore").decode().upper().strip()

@st.cache_data(show_spinner="Carregando base de recebimentos...")
def load_data():
    if not BASE_XLSX.exists():
        st.error(f"Base não encontrada: {BASE_XLSX}"); st.stop()
    d=pd.read_excel(str(BASE_XLSX),sheet_name="Recebimento Manual",engine="openpyxl")
    d.columns=[str(c).strip() for c in d.columns]
    for c in ["DATA_BASE","DATA RECEBIMENTO","DATA LANÇAMENTO SAP"]:
        d[c]=pd.to_datetime(d[c],errors="coerce")
    d["SLA_RECEBIMENTO"]=pd.to_numeric(d["SLA_RECEBIMENTO"],errors="coerce")
    d["MES_REF"]=d["DATA RECEBIMENTO"].dt.to_period("M").dt.to_timestamp()
    tipo=d["Tipo_recebimento"].map(norm); processo=d["PROCESSO"].map(norm)
    d=d[(tipo.str.contains("TRANSFER")|tipo.str.contains("REVERS")) & ~processo.str.contains("FORNECEDOR",na=False)].copy()
    d["FAIXA_SLA"]=pd.cut(d["SLA_RECEBIMENTO"],[-float("inf"),0,1,2,3,4,float("inf")],labels=["D+0","D+1","D+2","D+3","D+4","Acima de D+4"])
    return d

def fmt_int(v): return f"{int(v):,}".replace(",",".")
def fmt_pct(v): return f"{v:.1%}".replace(".",",")
def cards(data):
    total = len(data)
    cols = st.columns(6)
    faixas = [("D+0", 0), ("D+1", 1), ("D+2", 2), ("D+3", 3), ("D+4", 4), ("Acima de D+4", None)]

    for col, (label, lim) in zip(cols, faixas):
        if lim is None:
            qtd = int((data["SLA_RECEBIMENTO"] > 4).sum())
            percentual = qtd / total if total else 0
            sub = f"↑ {fmt_pct(percentual)} do total"
        else:
            qtd = int((data["SLA_RECEBIMENTO"] == lim).sum())
            acumulado = int((data["SLA_RECEBIMENTO"] <= lim).sum())
            percentual = acumulado / total if total else 0
            sub = f"↑ Acumulado até D+{lim}: {fmt_pct(percentual)}"

        # Card HTML próprio para garantir a centralização exata em qualquer versão do Streamlit.
        col.markdown(
            f"""
            <div class="sla-card">
                <div class="sla-card-title">SLA {label}</div>
                <div class="sla-card-value">{fmt_int(qtd)}</div>
                <div class="sla-card-detail">{sub}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

def resumo(data):
    if data.empty:return "Sem dados para os filtros selecionados."
    mx=data.MES_REF.max(); x=data[data.MES_REF>=mx-pd.DateOffset(months=2)]
    linhas=["RESUMO EXECUTIVO | CONTROLE DE RECEBIMENTO DE MATERIAIS",""]
    for mes,g in x.groupby("MES_REF"):
        n=len(g); a1=(g.SLA_RECEBIMENTO<=1).sum(); a4=(g.SLA_RECEBIMENTO<=4).sum(); f=(g.SLA_RECEBIMENTO>4).sum()
        linhas.append(f"{mes:%m/%Y}: {fmt_int(n)} recebimentos | Até D+1: {fmt_pct(a1/n)} | Até D+4: {fmt_pct(a4/n)} | Acima de D+4: {fmt_int(f)} ({fmt_pct(f/n)})")
    return "\n".join(linhas)

st.markdown('''<style>
.block-container{padding-top:1.4rem}.kpi-note{background:#fff7f7;border-left:7px solid #e30613;padding:17px 22px;border-radius:10px;margin:8px 0 20px}
div[data-testid="stMetric"]{text-align:center} div[data-testid="stMetric"]>div{justify-content:center}
div[data-testid="stMetricLabel"],div[data-testid="stMetricValue"],div[data-testid="stMetricDelta"]{justify-content:center;width:100%;text-align:center}
div[data-testid="stMetricLabel"]>div{width:100%;display:flex;justify-content:center;text-align:center}
div[data-testid="stMetricLabel"] p{width:100%;margin:0;text-align:center!important}
div[data-testid="stMetricValue"]>div{width:100%;text-align:center}
div[data-testid="stMetricDelta"]>div{width:100%;display:flex;justify-content:center;text-align:center}
.mes-referencia-wrap{width:100%;display:flex;justify-content:center;margin:4px 0 14px}
.mes-referencia-badge{display:inline-flex;align-items:center;justify-content:center;gap:8px;background:#f3f4f7;border:1px solid #d9dce3;border-radius:18px;padding:7px 18px;color:#202536;font-size:15px;line-height:1.2}
.mes-referencia-badge strong{font-size:16px;color:#e30613}
.sla-card{width:100%;display:flex;flex-direction:column;align-items:center;justify-content:flex-start;text-align:center;padding:7px 0 12px;box-sizing:border-box}
.sla-card-title{width:100%;text-align:center!important;font-size:14px;line-height:1.35;margin:0 0 12px;color:#202536;white-space:nowrap}
.sla-card-value{width:100%;text-align:center!important;font-size:38px;line-height:1.15;font-weight:400;color:#202536;margin:0 0 10px}
.sla-card-detail{width:100%;min-height:28px;display:flex;align-items:center;justify-content:center;text-align:center!important;background:#e5f7ec;color:#008a3b;border-radius:18px;padding:4px 8px;font-size:14px;line-height:1.25;box-sizing:border-box;white-space:nowrap}
</style>''',unsafe_allow_html=True)
st.title("📦 Controle de Recebimento de Materiais")
st.markdown('<div class="kpi-note"><b>Transferência e Reversa</b><br>Monitoramento do prazo entre o recebimento físico e o lançamento no SAP, excluindo fornecedor de materiais novos.</div>',unsafe_allow_html=True)
df=load_data()
with st.sidebar:
    st.header("Visualização"); vis=st.radio("Selecione",["📅 Visão diária","📊 Evolução mensal"],label_visibility="collapsed")
    st.header("Filtros"); meses=sorted(df.MES_REF.dropna().unique(),reverse=True); mes=st.selectbox("Mês de referência",meses,format_func=lambda x:pd.Timestamp(x).strftime("%m/%Y"))
    def multi(label,col): return st.multiselect(label,sorted(df[col].dropna().astype(str).str.strip().unique()))
    fs=[(multi("Tipo de recebimento","Tipo_recebimento"),"Tipo_recebimento"),(multi("CD corrigido","CD_CORRIGIDO"),"CD_CORRIGIDO"),(multi("Empresa","EMPRESA"),"EMPRESA"),(multi("Processo","PROCESSO"),"PROCESSO"),(multi("Origem","ORIGEM"),"ORIGEM")]
    nf=st.text_input("Buscar Nota Fiscal")
base=df.copy()
for vals,col in fs:
    if vals:base=base[base[col].astype(str).str.strip().isin(vals)]
if nf:base=base[base.NOTAFISCAL.astype(str).str.contains(re.escape(nf),case=False,na=False)]

if vis.startswith("📅"):
    mes_formatado = pd.Timestamp(mes).strftime("%m/%Y")
    st.markdown(
        f"""
        <div class="mes-referencia-wrap">
            <div class="mes-referencia-badge">
                📅 Mês de referência selecionado: <strong>{mes_formatado}</strong>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    data=base[base.MES_REF==pd.Timestamp(mes)].copy(); cards(data); c1,c2=st.columns([1.55,1])
    with c1:
        daily=data.groupby(data["DATA RECEBIMENTO"].dt.date,as_index=False).agg(Recebimentos=("NOTAFISCAL","size"))
        daily["Dia"]=pd.to_datetime(daily["DATA RECEBIMENTO"]).dt.day.astype(str); ordem=daily.Dia.tolist()
        fig=px.bar(daily,x="Dia",y="Recebimentos",text="Recebimentos",title=f"Recebimentos por dia | {pd.Timestamp(mes):%m/%Y}",category_orders={"Dia":ordem})
        fig.update_traces(textposition="inside",textfont_color="white"); fig.update_xaxes(title_text="Dia do mês",type="category",tickmode="array",tickvals=ordem,ticktext=ordem,tickangle=0); fig.update_layout(bargap=.18)
        st.plotly_chart(fig,use_container_width=True)
    with c2:
        dist=data.FAIXA_SLA.value_counts(sort=False).rename_axis("Faixa").reset_index(name="Quantidade")
        st.plotly_chart(px.bar(dist,x="Faixa",y="Quantidade",color="Faixa",text_auto=True,title="Distribuição por faixa de SLA"),use_container_width=True)
    st.subheader("Detalhamento dos recebimentos"); cols=["DATA RECEBIMENTO","DATA LANÇAMENTO SAP","Tipo_recebimento","SLA_RECEBIMENTO","STATUS RECEBIMENTO","CD_CORRIGIDO","EMPRESA","PROCESSO","ORIGEM","NOTAFISCAL"]
    st.dataframe(data[cols].sort_values("DATA RECEBIMENTO",ascending=False),use_container_width=True,hide_index=True)
else:
    n=st.segmented_control("Período para comparação",[3,6,9,12],default=3,format_func=lambda x:f"Últimos {x} meses"); mx=base.MES_REF.max(); evo=base[base.MES_REF>=mx-pd.DateOffset(months=int(n)-1)]
    mensal=evo.groupby("MES_REF",as_index=False).agg(Recebimentos=("NOTAFISCAL","size"),SLA_D0=("SLA_RECEBIMENTO",lambda s:(s<=0).mean()),SLA_D1=("SLA_RECEBIMENTO",lambda s:(s<=1).mean()),SLA_D2=("SLA_RECEBIMENTO",lambda s:(s<=2).mean()),SLA_D3=("SLA_RECEBIMENTO",lambda s:(s<=3).mean()),SLA_D4=("SLA_RECEBIMENTO",lambda s:(s<=4).mean()))
    cs=["SLA_D0","SLA_D1","SLA_D2","SLA_D3","SLA_D4"]; longa=mensal.melt(id_vars=["MES_REF","Recebimentos"],value_vars=cs,var_name="SLA",value_name="Percentual"); longa["Mês"]=longa.MES_REF.dt.strftime("%m/%Y")
    fig=px.line(longa,x="Mês",y="Percentual",color="SLA",markers=True,text=longa.Percentual.map(fmt_pct),title="Evolução mensal do SLA acumulado"); fig.update_yaxes(tickformat=".0%",range=[0,1.08]); fig.update_traces(textposition="top center"); st.plotly_chart(fig,use_container_width=True)
    st.plotly_chart(px.bar(mensal.assign(Mês=mensal.MES_REF.dt.strftime("%m/%Y")),x="Mês",y="Recebimentos",text_auto=True,title="Volume mensal de recebimentos"),use_container_width=True)
    tabela=mensal.assign(Mês=mensal.MES_REF.dt.strftime("%m/%Y"))[["Mês","Recebimentos"]+cs].copy(); tabela[cs]=tabela[cs].mul(100)
    st.dataframe(tabela,use_container_width=True,hide_index=True,column_config={c:st.column_config.ProgressColumn(c,format="%.2f%%",min_value=0,max_value=100) for c in cs})

st.divider(); st.subheader("✉️ Envio por e-mail | resumo executivo dos últimos 3 meses"); texto=resumo(base)
with st.expander("Visualizar resumo"):st.text(texto)
a,b=st.columns([2,1]); dest=a.text_input("Destinatários",placeholder="nome@empresa.com; outro@empresa.com"); assunto=b.text_input("Assunto",value="Controle de Recebimento de Materiais | Resumo dos últimos 3 meses")
url="https://outlook.office.com/mail/deeplink/compose?to="+quote(dest.replace(";",","))+"&subject="+quote(assunto)+"&body="+quote(texto); st.link_button("Abrir e-mail no Outlook Web",url,use_container_width=True)
