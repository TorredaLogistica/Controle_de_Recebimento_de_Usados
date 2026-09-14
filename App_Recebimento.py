import os, re, unicodedata
from urllib.parse import quote
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Controle de Recebimento de Materiais", page_icon="📦", layout="wide")
BASE_XLSX = "Controle de Recebimento de Materiais.xlsx"

def norm(v):
    s = "" if pd.isna(v) else str(v)
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().upper().strip()

@st.cache_data(show_spinner="Carregando base de recebimentos...")
def load_data():
    if os.path.exists(BASE_XLSX):
        d = pd.read_excel(
            BASE_XLSX,
            sheet_name="Recebimento Manual",
            engine="openpyxl",
            usecols=[
                "ANO/MÊS_RECEBIMENTO", "Tipo_recebimento", "STATUS RECEBIMENTO",
                "SLA_RECEBIMENTO", "SLA_ENVIO", "CD_CORRIGIDO", "DATA_BASE",
                "CD", "EMPRESA", "PROCESSO", "ORIGEM", "NOTAFISCAL",
                "DATA RECEBIMENTO", "DATA LANÇAMENTO SAP"
            ]
        )
    else:
        st.error(f"Base não encontrada. Inclua o arquivo {BASE_XLSX} na mesma pasta do app.")
        st.stop()
    d.columns = [str(c).strip() for c in d.columns]
    for c in ["DATA_BASE", "DATA RECEBIMENTO", "DATA LANÇAMENTO SAP"]:
        d[c] = pd.to_datetime(d[c], errors="coerce")
    d["SLA_RECEBIMENTO"] = pd.to_numeric(d["SLA_RECEBIMENTO"], errors="coerce")
    # Mês de referência: coluna G da base (DATA_BASE).
    d["MES_REF"] = d["DATA_BASE"].dt.to_period("M").dt.to_timestamp()

    # Mantém todos os tipos de recebimento da base, inclusive Obsoletos.
    # O filtro lateral permite selecionar Transferência, Reversa Massivo ou Obsoletos.
    d["FAIXA_SLA"] = pd.cut(d["SLA_RECEBIMENTO"], bins=[-float("inf"),0,1,2,3,4,float("inf")], labels=["D+0","D+1","D+2","D+3","D+4","Acima de D+4"])
    return d

def fmt_int(v): return f"{int(v):,}".replace(",", ".")
def fmt_pct(v): return f"{v:.1%}".replace(".", ",")

def metric_cards(data):
    total = len(data)
    buckets = [("D+0", 0), ("D+1", 1), ("D+2", 2), ("D+3", 3), ("D+4", 4), ("Acima de D+4", None)]
    cols = st.columns(6)
    for col, (label, lim) in zip(cols, buckets):
        if lim is None:
            qtd = int((data["SLA_RECEBIMENTO"] > 4).sum())
            pct = qtd / total if total else 0
            detalhe = f"{fmt_pct(pct)} do total"
        else:
            qtd = int((data["SLA_RECEBIMENTO"] == lim).sum())
            acumulado = int((data["SLA_RECEBIMENTO"] <= lim).sum())
            pct = acumulado / total if total else 0
            detalhe = f"Acumulado até D+{lim}: {fmt_pct(pct)}"
        col.markdown(f"""
        <div class="sla-card">
          <div class="sla-card-title">SLA {label}</div>
          <div class="sla-card-value">{fmt_int(qtd)}</div>
          <div class="sla-card-detail">{detalhe}</div>
        </div>""", unsafe_allow_html=True)

def executive_summary(data, months=3):
    if data.empty: return "Sem dados para os filtros selecionados."
    mx=data["MES_REF"].max(); start=mx-pd.DateOffset(months=months-1)
    x=data[data["MES_REF"].between(start,mx)].copy()
    lines=["RESUMO EXECUTIVO | CONTROLE DE RECEBIMENTO DE MATERIAIS", ""]
    for mes,g in x.groupby("MES_REF", sort=True):
        total=len(g); ate1=(g["SLA_RECEBIMENTO"]<=1).sum(); ate4=(g["SLA_RECEBIMENTO"]<=4).sum(); fora=(g["SLA_RECEBIMENTO"]>4).sum()
        lines.append(f"{mes.strftime('%m/%Y')}: {fmt_int(total)} recebimentos | Até D+1: {fmt_pct(ate1/total if total else 0)} | Até D+4: {fmt_pct(ate4/total if total else 0)} | Acima de D+4: {fmt_int(fora)} ({fmt_pct(fora/total if total else 0)})")
    lines += ["", "Escopo: todos os tipos de recebimento disponíveis na base, conforme os filtros selecionados."]
    return "\n".join(lines)

st.markdown("""<style>
.block-container{padding-top:1.4rem}.kpi-note{background:#fff7f7;border-left:7px solid #e30613;padding:17px 22px;border-radius:10px;margin:8px 0 20px}.small{color:#65707d;font-size:.89rem}
.sla-card{width:100%;min-height:170px;display:flex;flex-direction:column;align-items:center;text-align:center;padding:5px 3px 12px;box-sizing:border-box}
.sla-card-title{width:100%;min-height:38px;display:flex;align-items:center;justify-content:center;text-align:center;font-size:14px;line-height:1.3}
.sla-card-value{width:100%;text-align:center;font-size:39px;line-height:1.15;color:#202536;margin:3px 0 10px}
.sla-card-detail{width:100%;min-height:62px;display:flex;align-items:center;justify-content:center;text-align:center;white-space:normal!important;overflow:visible!important;background:#e5f7ec;color:#008a3b;border-radius:14px;padding:8px;font-size:14px;font-weight:600;line-height:1.35;box-sizing:border-box}

</style>""", unsafe_allow_html=True)
st.title("📦 Controle de Recebimento de Materiais")
st.markdown('<div class="kpi-note"><b>Controle de Recebimentos</b><br>Monitoramento do prazo entre o recebimento físico e o lançamento no SAP, considerando os tipos selecionados nos filtros.</div>', unsafe_allow_html=True)

df=load_data()
with st.sidebar:
    st.header("Visualização")
    vis=st.radio("Selecione", ["📅 Visão diária", "📊 Evolução mensal"], label_visibility="collapsed")
    st.header("Filtros")
    meses=sorted(df["MES_REF"].dropna().unique(), reverse=True)
    mes=st.selectbox("Mês de referência", meses, format_func=lambda x: pd.Timestamp(x).strftime("%m/%Y"))
    def multi(label,col):
        vals=sorted(df[col].dropna().astype(str).str.strip().unique())
        return st.multiselect(label, vals)
    f_tipo=multi("Tipo de recebimento","Tipo_recebimento")
    f_cd=multi("CD corrigido","CD_CORRIGIDO")
    f_empresa=multi("Empresa","EMPRESA")
    f_processo=multi("Processo","PROCESSO")
    f_origem=multi("Origem","ORIGEM")
    nf=st.text_input("Buscar Nota Fiscal")

def apply(d):
    x=d.copy()
    for vals,col in [(f_tipo,"Tipo_recebimento"),(f_cd,"CD_CORRIGIDO"),(f_empresa,"EMPRESA"),(f_processo,"PROCESSO"),(f_origem,"ORIGEM")]:
        if vals: x=x[x[col].astype(str).str.strip().isin(vals)]
    if nf: x=x[x["NOTAFISCAL"].astype(str).str.contains(re.escape(nf),case=False,na=False)]
    return x
base=apply(df)

if vis.startswith("📅"):
    data=base[base["MES_REF"]==pd.Timestamp(mes)].copy()
    metric_cards(data)
    ordem_faixas = ["D+0", "D+1", "D+2", "D+3", "D+4", "Acima de D+4"]
    cores_sla = {
        "D+0": "#0B70C9", "D+1": "#72B9EB", "D+2": "#FF1F26",
        "D+3": "#F7A6AA", "D+4": "#2AA894", "Acima de D+4": "#6DDE98",
    }

    # Gráfico diário em largura total.
    daily = data.dropna(subset=["DATA RECEBIMENTO"]).copy()
    daily["Dia_num"] = daily["DATA RECEBIMENTO"].dt.day
    daily = (daily.groupby("Dia_num", as_index=False)
             .agg(Recebimentos=("NOTAFISCAL", "size"))
             .sort_values("Dia_num"))
    daily["Dia"] = daily["Dia_num"].astype(str)
    ordem_dias = daily["Dia"].tolist()
    if daily.empty:
        st.info("Não há recebimentos com data válida para o mês e os filtros selecionados.")
    else:
        fig_dia = px.bar(daily, x="Dia", y="Recebimentos", text="Recebimentos",
                         title=f"Recebimentos por dia | {pd.Timestamp(mes).strftime('%m/%Y')}",
                         category_orders={"Dia": ordem_dias})
        fig_dia.update_traces(textposition="inside", textfont_color="white", textfont_size=14)
        fig_dia.update_xaxes(title_text="Dia do mês", type="category", tickmode="array",
                             tickvals=ordem_dias, ticktext=ordem_dias, tickangle=0)
        fig_dia.update_layout(height=470, bargap=0.18)
        st.plotly_chart(fig_dia, use_container_width=True)

    # Distribuição geral em largura total para os percentuais ficarem legíveis.
    dist = (data["FAIXA_SLA"].value_counts(sort=False)
            .reindex(ordem_faixas, fill_value=0)
            .rename_axis("Faixa").reset_index(name="Quantidade"))
    total_dist = int(dist["Quantidade"].sum())
    dist["Percentual"] = dist["Quantidade"] / total_dist if total_dist else 0
    dist["Percentual_acumulado"] = dist["Percentual"].cumsum()
    dist["Pct"] = dist["Percentual"].map(fmt_pct)
    dist["Pct_acum"] = dist["Percentual_acumulado"].map(fmt_pct)
    dist["Rotulo"] = dist.apply(
        lambda r: f"<b>{fmt_int(r['Quantidade'])}</b><br><b>{r['Pct']}</b><br>Acum. {r['Pct_acum']}", axis=1)
    fig = px.bar(dist, x="Faixa", y="Quantidade", color="Faixa", text="Rotulo",
                 custom_data=["Pct", "Pct_acum"], title="Distribuição por faixa de SLA",
                 category_orders={"Faixa": ordem_faixas}, color_discrete_map=cores_sla)
    fig.update_traces(textposition="outside", cliponaxis=False, textfont_size=15,
                      hovertemplate="Faixa: %{x}<br>Quantidade: %{y}<br>Percentual: %{customdata[0]}<br>Percentual acumulado: %{customdata[1]}<extra></extra>")
    maior = int(dist["Quantidade"].max()) if not dist.empty else 0
    fig.update_yaxes(range=[0, max(maior * 1.38, 1)])
    fig.update_layout(height=610, margin=dict(t=75, r=30, b=90, l=65), showlegend=True)
    st.plotly_chart(fig, use_container_width=True)

    # Comparativo por CD no mesmo padrão visual da distribuição geral.
    st.subheader("Distribuição do SLA por CD")
    dados_cd = data.dropna(subset=["CD_CORRIGIDO", "FAIXA_SLA"]).copy()
    dados_cd["CD_CORRIGIDO"] = dados_cd["CD_CORRIGIDO"].astype(str).str.strip()
    dados_cd = dados_cd[dados_cd["CD_CORRIGIDO"] != ""]

    if dados_cd.empty:
        st.info("Não há dados de CD para o mês e os filtros selecionados.")
    else:
        sla_cd = (
            dados_cd.groupby(["CD_CORRIGIDO", "FAIXA_SLA"], observed=True)
            .size()
            .reset_index(name="Quantidade")
        )
        totais_cd = (
            sla_cd.groupby("CD_CORRIGIDO", as_index=False)["Quantidade"]
            .sum()
            .rename(columns={"Quantidade": "Total_CD"})
        )
        sla_cd = sla_cd.merge(totais_cd, on="CD_CORRIGIDO", how="left")
        sla_cd["Percentual_CD"] = sla_cd["Quantidade"] / sla_cd["Total_CD"]
        sla_cd["Percentual_formatado"] = sla_cd["Percentual_CD"].map(fmt_pct)

        ordem_cds = (
            totais_cd.sort_values("Total_CD", ascending=False)["CD_CORRIGIDO"]
            .tolist()
        )

        fig_cd = px.bar(
            sla_cd,
            x="CD_CORRIGIDO",
            y="Quantidade",
            color="FAIXA_SLA",
            text="Quantidade",
            barmode="group",
            title=f"SLA por CD | {pd.Timestamp(mes).strftime('%m/%Y')}",
            category_orders={
                "CD_CORRIGIDO": ordem_cds,
                "FAIXA_SLA": ordem_faixas,
            },
            color_discrete_map=cores_sla,
            custom_data=["Percentual_formatado", "Total_CD"],
            labels={
                "CD_CORRIGIDO": "CD",
                "FAIXA_SLA": "Faixa",
                "Quantidade": "Quantidade",
            },
        )
        fig_cd.update_traces(
            textposition="outside",
            cliponaxis=False,
            hovertemplate=(
                "CD: %{x}<br>"
                "Faixa: %{fullData.name}<br>"
                "Quantidade: %{y}<br>"
                "Percentual no CD: %{customdata[0]}<br>"
                "Total do CD: %{customdata[1]}<extra></extra>"
            ),
        )
        fig_cd.update_layout(
            xaxis_title="CD",
            yaxis_title="Quantidade",
            legend_title_text="Faixa",
            bargap=0.18,
            bargroupgap=0.06,
            height=max(480, min(760, 440 + len(ordem_cds) * 12)),
        )
        fig_cd.update_xaxes(tickangle=-35)
        st.plotly_chart(fig_cd, use_container_width=True)

    st.subheader("Detalhamento dos recebimentos")
    detail_cols=["DATA RECEBIMENTO","DATA LANÇAMENTO SAP","Tipo_recebimento","SLA_RECEBIMENTO","STATUS RECEBIMENTO","CD_CORRIGIDO","EMPRESA","PROCESSO","ORIGEM","NOTAFISCAL"]
    st.dataframe(data[detail_cols].sort_values("DATA RECEBIMENTO",ascending=False),use_container_width=True,hide_index=True)
else:
    qtd_meses=st.segmented_control("Período para comparação",[3,6,9,12],default=3,format_func=lambda n:f"Últimos {n} meses")
    maxm=base["MES_REF"].max()
    evo=base[base["MES_REF"]>=maxm-pd.DateOffset(months=int(qtd_meses)-1)].copy() if pd.notna(maxm) else base
    mensal=evo.groupby("MES_REF",as_index=False).agg(Recebimentos=("NOTAFISCAL","size"),SLA_D0=("SLA_RECEBIMENTO",lambda s:(s<=0).mean()),SLA_D1=("SLA_RECEBIMENTO",lambda s:(s<=1).mean()),SLA_D2=("SLA_RECEBIMENTO",lambda s:(s<=2).mean()),SLA_D3=("SLA_RECEBIMENTO",lambda s:(s<=3).mean()),SLA_D4=("SLA_RECEBIMENTO",lambda s:(s<=4).mean()))
    longa=mensal.melt(id_vars=["MES_REF","Recebimentos"],value_vars=["SLA_D0","SLA_D1","SLA_D2","SLA_D3","SLA_D4"],var_name="SLA",value_name="Percentual")
    longa["Mês"]=longa["MES_REF"].dt.strftime("%m/%Y")
    fig=px.line(longa,x="Mês",y="Percentual",color="SLA",markers=True,text=longa["Percentual"].map(lambda x:fmt_pct(x)),title="Evolução mensal do SLA acumulado")
    fig.update_yaxes(tickformat=".0%",range=[0,1.08]); fig.update_traces(textposition="top center")
    st.plotly_chart(fig,use_container_width=True)

    # Comparativo mensal em linhas, com destaque visual para Até D+2 e Até D+4.
    nomes_sla = {
        "SLA_D0": "Até D+0",
        "SLA_D1": "Até D+1",
        "SLA_D2": "Até D+2",
        "SLA_D3": "Até D+3",
        "SLA_D4": "Até D+4",
    }
    cores_acumuladas = {
        "Até D+0": "#AEB8C4",
        "Até D+1": "#8FC5EA",
        "Até D+2": "#FF1F26",
        "Até D+3": "#F4B0B4",
        "Até D+4": "#1A9D88",
    }
    barras_mensais = mensal.melt(
        id_vars=["MES_REF"],
        value_vars=list(nomes_sla),
        var_name="SLA",
        value_name="Percentual",
    )
    barras_mensais["Mês"] = barras_mensais["MES_REF"].dt.strftime("%m/%Y")
    barras_mensais["Faixa acumulada"] = barras_mensais["SLA"].map(nomes_sla)
    barras_mensais["Rótulo"] = barras_mensais["Percentual"].map(fmt_pct)
    ordem_meses = mensal["MES_REF"].dt.strftime("%m/%Y").tolist()

    fig_barras = px.line(
        barras_mensais,
        x="Mês",
        y="Percentual",
        color="Faixa acumulada",
        markers=True,
        text="Rótulo",
        title="SLA acumulado por mês | Destaque Até D+2 e Até D+4",
        category_orders={
            "Mês": ordem_meses,
            "Faixa acumulada": ["Até D+0", "Até D+1", "Até D+2", "Até D+3", "Até D+4"],
        },
        color_discrete_map=cores_acumuladas,
    )

    # Linhas secundárias permanecem mais discretas para priorizar D+2 e D+4.
    estilos_linhas = {
        "Até D+0": {"width": 1.5, "dash": "dot"},
        "Até D+1": {"width": 1.5, "dash": "dot"},
        "Até D+2": {"width": 4.5, "dash": "solid"},
        "Até D+3": {"width": 1.5, "dash": "dot"},
        "Até D+4": {"width": 4.5, "dash": "solid"},
    }
    for trace in fig_barras.data:
        destaque = trace.name in ["Até D+2", "Até D+4"]

        # Exibe rótulos somente nos SLAs destacados para evitar sobreposição.
        # Até D+4 fica acima da linha e Até D+2 abaixo da linha.
        if trace.name == "Até D+4":
            modo_trace = "lines+markers+text"
            posicao_texto = "top center"
        elif trace.name == "Até D+2":
            modo_trace = "lines+markers+text"
            posicao_texto = "bottom center"
        else:
            modo_trace = "lines+markers"
            posicao_texto = "top center"

        trace.update(
            mode=modo_trace,
            line=estilos_linhas[trace.name],
            marker=dict(
                size=11 if destaque else 7,
                line=dict(width=2 if destaque else 0, color="white"),
            ),
            textfont=dict(
                size=14 if destaque else 11,
                color=cores_acumuladas[trace.name],
            ),
            textposition=posicao_texto,
            text=None if not destaque else trace.text,
            opacity=1.0 if destaque else 0.42,
            hovertemplate="Mês: %{x}<br>%{fullData.name}: %{customdata}<extra></extra>",
            customdata=trace.text,
        )

    fig_barras.update_yaxes(
        title_text="Percentual",
        tickformat=".0%",
        range=[0, 1.13],
    )
    fig_barras.update_layout(
        height=620,
        legend_title_text="SLA acumulado",
        hovermode="x unified",
        margin=dict(t=80, r=30, b=70, l=65),
    )
    st.plotly_chart(fig_barras, use_container_width=True)

    fig2=px.bar(mensal.assign(Mês=mensal["MES_REF"].dt.strftime("%m/%Y")),x="Mês",y="Recebimentos",text_auto=True,title="Volume mensal de recebimentos")
    st.plotly_chart(fig2,use_container_width=True)
    # Na tabela, converte a proporção decimal para percentual real.
    # Exemplo: 0.2708 passa a ser exibido como 27,08%, e não 0,3%.
    colunas_sla = ["SLA_D0", "SLA_D1", "SLA_D2", "SLA_D3", "SLA_D4"]
    tabela_mensal = mensal.assign(
        **{"Mês": mensal["MES_REF"].dt.strftime("%m/%Y")}
    )[["Mês", "Recebimentos"] + colunas_sla].copy()

    tabela_mensal[colunas_sla] = tabela_mensal[colunas_sla].mul(100)

    st.dataframe(
        tabela_mensal,
        use_container_width=True,
        hide_index=True,
        column_config={
            c: st.column_config.ProgressColumn(
                c,
                format="%.2f%%",
                min_value=0,
                max_value=100,
            )
            for c in colunas_sla
        },
    )

st.divider(); st.subheader("✉️ Envio por e-mail | resumo executivo dos últimos 3 meses")
resumo=executive_summary(base,3)
with st.expander("Visualizar resumo",expanded=False): st.text(resumo)
c1,c2=st.columns([2,1])
with c1:
    destinatarios=st.text_input("Destinatários",placeholder="nome@empresa.com; outro@empresa.com")
with c2:
    assunto=st.text_input("Assunto",value="Controle de Recebimento de Materiais | Resumo dos últimos 3 meses")
url="https://outlook.office.com/mail/deeplink/compose?to="+quote(destinatarios.replace(";",","))+"&subject="+quote(assunto)+"&body="+quote(resumo)
st.link_button("Abrir e-mail no Outlook Web",url,use_container_width=True)
st.download_button("Baixar resumo em texto",resumo,file_name="Resumo_Executivo_Recebimentos.txt",mime="text/plain")
st.caption("O envio permanece manual e seguro: o botão abre a mensagem pronta no Outlook Web para revisão antes do envio.")
