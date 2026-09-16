from pathlib import Path
import pandas as pd
import plotly.express as px
import streamlit as st
from parser import procesar_archivo

st.set_page_config(page_title='KAIRO · Control de Gestión', page_icon='📊', layout='wide')
MI_EMPRESA='KAIRO'
REGIONES=['EE.UU.','Asia','Europa']
TECS=['Tec 1','Tec 2','Tec 3','Tec 4']

st.markdown('''<style>
.block-container{padding-top:1.4rem;max-width:1500px}.kpi{border:1px solid rgba(128,128,128,.25);border-radius:14px;padding:14px 12px;min-height:112px;overflow:hidden}.kpi-label{font-size:.78rem;opacity:.72;white-space:normal}.kpi-value{font-size:clamp(1.15rem,2.1vw,1.9rem);font-weight:750;line-height:1.15;margin-top:7px;overflow-wrap:anywhere}.kpi-note{font-size:.72rem;opacity:.62;margin-top:6px}.section-note{opacity:.72;font-size:.9rem}div[data-testid="stMetric"]{overflow:hidden}
</style>''', unsafe_allow_html=True)

def kpi(col,label,value,note=''):
    col.markdown(f'<div class="kpi"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div><div class="kpi-note">{note}</div></div>',unsafe_allow_html=True)
def n(x,d=1):
    if x is None or pd.isna(x): return '—'
    return f'{float(x):,.{d}f}'.replace(',','X').replace('.',',').replace('X','.')
def money(x): return '—' if x is None or pd.isna(x) else f'US$ {n(x,1)}k'
def pct(x): return '—' if x is None or pd.isna(x) else f'{n(x,2)}%'
def bar_rank(df,x,title,label):
    if df.empty:return None
    z=df.sort_values(x,ascending=True).copy(); z['KAIRO']=z.equipo.eq(MI_EMPRESA)
    fig=px.bar(z,x=x,y='equipo',orientation='h',text=x,title=title,labels={x:label,'equipo':''},color='KAIRO',color_discrete_map={True:'#5B6CFF',False:'#A9B1C7'})
    fig.update_layout(showlegend=False,height=max(330,45*len(z)))
    fig.update_traces(texttemplate='%{text:,.2f}',textposition='outside')
    return fig

def load_all():
    root=Path(__file__).parent; folder=root/'resultados'; files=[]
    if folder.exists(): files=sorted(list(folder.glob('*.xlsx'))+list(folder.glob('*.xls')))
    # Local fallback makes the package easy to test.
    if not files: files=sorted(list(root.glob('results-r*.xlsx'))+list(root.glob('results-r*.xls')))
    allsets={k:[] for k in ['mercado','market_share','finanzas','demanda_insatisfecha','produccion','stock','logistica_detalle','rrhh','promocion','finanzas_control','logistica_financiera']}
    rounds=[]
    for f in files:
        try:
            r=procesar_archivo(f.name,f.read_bytes()); rounds.append(r['ronda'])
            for k in allsets:
                if k in r and isinstance(r[k],pd.DataFrame) and not r[k].empty: allsets[k].append(r[k])
        except Exception as e: st.warning(f'No pude leer {f.name}: {e}')
    return sorted(set(rounds)),{k:(pd.concat(v,ignore_index=True) if v else pd.DataFrame()) for k,v in allsets.items()}

rounds,D=load_all()
logo=Path(__file__).parent/'assets'/'kairo_logo.png'
head1,head2=st.columns([1,5])
if logo.exists(): head1.image(str(logo),width=125)
head2.title('KAIRO · Control de Gestión')
head2.caption('CESIM · análisis por ronda, región, tecnología y competidores')
if not rounds: st.error('No encontré archivos results-rXX.xlsx en /resultados.'); st.stop()
ronda=st.sidebar.selectbox('Ronda',rounds,index=len(rounds)-1)
st.sidebar.caption('Los datos se leen automáticamente desde la carpeta resultados.')

def R(key):
    d=D[key]
    return d[d.ronda==ronda].copy() if not d.empty and 'ronda' in d else pd.DataFrame()
merc,share,fin,unmet,prod,stock,logd,rrhh,promo,fctl,lfin=[R(k) for k in ['mercado','market_share','finanzas','demanda_insatisfecha','produccion','stock','logistica_detalle','rrhh','promocion','finanzas_control','logistica_financiera']]

def finval(k):
    x=fin[(fin.equipo==MI_EMPRESA)&(fin.kpi==k)] if not fin.empty else pd.DataFrame()
    return None if x.empty else float(x.valor.iloc[0])
def shareval():
    x=share[(share.equipo==MI_EMPRESA)&(share.region=='Global')&(share.tecnologia=='Total')] if not share.empty else pd.DataFrame()
    return None if x.empty else float(x.market_share_pct.iloc[0])

# Compact top cards: no overflow.
cols=st.columns(5)
kpi(cols[0],'Ingresos',money(finval('Ingresos por ventas')))
kpi(cols[1],'EBITDA',money(finval('EBITDA')))
kpi(cols[2],'Beneficio',money(finval('Beneficio de la ronda')))
kpi(cols[3],'Market share global',pct(shareval()))
kpi(cols[4],'ROE',pct(finval('ROE, %')))
st.write('')

tabs=st.tabs(['📊 Demanda y Market Share','🏭 Producción','👥 Recursos Humanos','📣 Marketing y Promoción','💵 Finanzas, Préstamos y Deuda','🚚 Logística y Precios de Transferencia'])

# 1 DEMANDA + SHARE
with tabs[0]:
    st.subheader('Demanda, ventas y cuota de mercado')
    c1,c2=st.columns(2); reg=c1.selectbox('Región',REGIONES,key='dreg'); tech=c2.selectbox('Tecnología',['Total']+TECS,key='dtech')
    s=share[(share.region==reg)&(share.tecnologia==tech)] if not share.empty else pd.DataFrame()
    if not s.empty:
        fig=bar_rank(s.rename(columns={'market_share_pct':'valor'}),'valor',f'Market share · {reg} · {tech}','%'); st.plotly_chart(fig,width='stretch')
    if tech!='Total' and not merc.empty:
        m=merc[(merc.region==reg)&(merc.tecnologia==tech)].copy()
        if not m.empty:
            m['desvio_demanda']=m.ventas-m.demanda
            st.dataframe(m[['equipo','demanda','ventas','desvio_demanda','precio','caracteristicas']].sort_values('ventas',ascending=False),hide_index=True,width='stretch')
            g=m.melt(id_vars='equipo',value_vars=['demanda','ventas'],var_name='Serie',value_name='Miles unidades')
            st.plotly_chart(px.bar(g,x='equipo',y='Miles unidades',color='Serie',barmode='group',title='Demanda vs ventas'),width='stretch')

# 2 PRODUCTION
with tabs[1]:
    st.subheader('Producción, disponibilidad e inventario final')
    st.caption('El disponible y el inventario final se toman del detalle logístico de CESIM; no se estiman sólo como producción menos ventas.')
    if logd.empty: st.info('No encontré detalle logístico.')
    else:
        lk=logd[logd.equipo==MI_EMPRESA].copy()
        # Avoid double-counting production: it is recorded at production locations only.
        pint=lk.produccion_interna.fillna(0).sum(); pcon=lk.produccion_contratada.fillna(0).sum(); disp=lk.disponible.fillna(0).sum(); sf=lk.stock_final.fillna(0).sum()
        cc=st.columns(4); kpi(cc[0],'Producción interna',f'{n(pint,1)} mil u.'); kpi(cc[1],'Producción contratada',f'{n(pcon,1)} mil u.'); kpi(cc[2],'Disponible total',f'{n(disp,1)} mil u.','incluye inventario/importaciones por ubicación'); kpi(cc[3],'Inventario final',f'{n(sf,1)} mil u.')
        st.markdown('### KAIRO por tecnología y ubicación')
        show=lk[['region','tecnologia','stock_inicial','produccion_interna','produccion_contratada','disponible','ventas','stock_final']].fillna(0)
        st.dataframe(show,hide_index=True,width='stretch')
        st.markdown('### Producción interna y contratada · KAIRO vs competidores')
        pcomp=logd.groupby('equipo',as_index=False)[['produccion_interna','produccion_contratada']].sum()
        pm=pcomp.melt(id_vars='equipo',var_name='Tipo',value_name='Miles unidades'); pm['Tipo']=pm.Tipo.replace({'produccion_interna':'Interna','produccion_contratada':'Contratada'})
        st.plotly_chart(px.bar(pm,x='equipo',y='Miles unidades',color='Tipo',barmode='stack',title='Producción por empresa'),width='stretch')
        st.markdown('### Inventario final · KAIRO vs competidores')
        sr=logd.groupby('equipo',as_index=False).stock_final.sum(); st.plotly_chart(bar_rank(sr,'stock_final','Ranking de inventario final','Miles de unidades'),width='stretch')
        with st.expander('Ver stock por empresa, región y tecnología'):
            st.dataframe(logd[['equipo','region','tecnologia','disponible','ventas','stock_final']].sort_values(['equipo','region','tecnologia']),hide_index=True,width='stretch')
        st.markdown('### Demanda insatisfecha')
        ur=logd.groupby('equipo',as_index=False).demanda_insatisfecha.sum(); st.plotly_chart(bar_rank(ur,'demanda_insatisfecha','Demanda insatisfecha por empresa','Miles de unidades'),width='stretch')

# 3 HR
with tabs[2]:
    st.subheader('Recursos Humanos')
    labels={'capacitacion_mensual_usd':'Capacitación mensual (USD)','multiplicador_eficiencia':'Multiplicador de eficiencia','salario_mensual_usd':'Salario mensual (USD)','dotacion':'Dotación I+D','rotacion_pct':'Rotación (%)'}
    if rrhh.empty: st.info('No encontré el Informe de RRHH.')
    else:
        piv=rrhh.pivot_table(index='equipo',columns='indicador',values='valor',aggfunc='first').reset_index()
        rk=piv[piv.equipo==MI_EMPRESA].iloc[0] if not piv[piv.equipo==MI_EMPRESA].empty else None
        if rk is not None:
            cc=st.columns(5)
            for col,key in zip(cc,labels): kpi(col,labels[key],n(rk.get(key),2 if key in ['multiplicador_eficiencia','rotacion_pct'] else 0))
        indicador=st.selectbox('Comparar con competidores',list(labels),format_func=lambda x:labels[x])
        sub=rrhh[rrhh.indicador==indicador].rename(columns={'valor':'metrica'})
        st.plotly_chart(bar_rank(sub,'metrica',labels[indicador],labels[indicador]),width='stretch')
        st.dataframe(piv.rename(columns=labels),hide_index=True,width='stretch')

# 4 MARKETING
with tabs[3]:
    st.subheader('Marketing: precios, características y promoción')
    a,b=st.columns(2); reg=a.selectbox('Región',REGIONES,key='mreg'); tech=b.selectbox('Tecnología',TECS,key='mtech')
    m=merc[(merc.region==reg)&(merc.tecnologia==tech)].copy() if not merc.empty else pd.DataFrame()
    if not m.empty:
        st.markdown('### Precios y propuesta competitiva')
        st.dataframe(m[['equipo','precio','caracteristicas','ventas','marketing']].sort_values('precio',ascending=False),hide_index=True,width='stretch')
        st.plotly_chart(px.bar(m.sort_values('precio'),x='equipo',y='precio',text='precio',title=f'Precio · {reg} · {tech}',labels={'precio':f'Precio ({m.moneda.iloc[0]})'}),width='stretch')
        st.plotly_chart(px.scatter(m,x='precio',y='caracteristicas',size='ventas',text='equipo',title='Precio vs características (tamaño = ventas)'),width='stretch')
    st.markdown('### Inversión en promoción')
    pr=promo[promo.region==reg].copy() if not promo.empty else pd.DataFrame()
    if not pr.empty:
        st.plotly_chart(bar_rank(pr.rename(columns={'promocion_miles_usd':'valor'}),'valor',f'Promoción · {reg}','Miles USD'),width='stretch')
        st.dataframe(pr.sort_values('promocion_miles_usd',ascending=False),hide_index=True,width='stretch')

# 5 FINANCE
with tabs[4]:
    st.subheader('Finanzas, préstamos y deuda')
    if fctl.empty: st.info('No encontré los datos financieros de control.')
    else:
        fk=fctl[fctl.equipo==MI_EMPRESA]
        def fv(ind,region='Global'):
            x=fk[(fk.indicador==ind)&(fk.region==region)]; return None if x.empty else float(x.valor.iloc[0])
        cc=st.columns(4); kpi(cc[0],'Deuda corto plazo',money(fv('deuda_corto_plazo'))); kpi(cc[1],'Deuda largo plazo',money(fv('deuda_largo_plazo'))); kpi(cc[2],'Préstamos internos',money(fv('prestamos_internos'))); kpi(cc[3],'ROE',pct(finval('ROE, %')))
        st.markdown('### Inversión / desinversión de fábricas')
        inv=fctl[fctl.indicador=='inversion_desinversion_fabricas'].copy()
        if not inv.empty:
            st.caption('CESIM: inversión aparece con signo negativo y desinversión con signo positivo.')
            st.dataframe(inv.sort_values(['region','equipo']),hide_index=True,width='stretch')
            ik=inv[inv.equipo==MI_EMPRESA]
            st.plotly_chart(px.bar(ik,x='region',y='valor',text='valor',title='KAIRO · inversión/desinversión por región',labels={'valor':'Miles USD','region':''}),width='stretch')
        st.markdown('### Deuda · comparación')
        ind=st.selectbox('Indicador financiero',['deuda_corto_plazo','deuda_largo_plazo','prestamos_internos'],format_func=lambda x:{'deuda_corto_plazo':'Deuda corto plazo','deuda_largo_plazo':'Deuda largo plazo','prestamos_internos':'Préstamos internos'}[x])
        q=fctl[(fctl.indicador==ind)&(fctl.region=='Global')].rename(columns={'valor':'metrica'})
        st.plotly_chart(bar_rank(q,'metrica','Comparación financiera','Miles USD'),width='stretch')

# 6 LOGISTICS
with tabs[5]:
    st.subheader('Logística y precios de transferencia')
    st.markdown('### Costos de transporte y aranceles')
    tr=lfin[lfin.indicador=='costos_transporte_aranceles'].copy() if not lfin.empty else pd.DataFrame()
    reg=st.selectbox('Región logística',['Global']+REGIONES,key='lreg')
    t=tr[tr.region==reg].rename(columns={'valor':'metrica'}) if not tr.empty else pd.DataFrame()
    if not t.empty:
        st.plotly_chart(bar_rank(t,'metrica',f'Costos de transporte y aranceles · {reg}','Miles USD'),width='stretch')
        st.dataframe(t[['equipo','metrica']].sort_values('metrica',ascending=False),hide_index=True,width='stretch')
    st.markdown('### Transferencias internas')
    trans=lfin[(lfin.indicador=='transferencias_internas')].copy() if not lfin.empty else pd.DataFrame()
    if not trans.empty:
        st.dataframe(trans.sort_values(['region','equipo']),hide_index=True,width='stretch')
    st.info('El archivo de resultados de CESIM no expone el multiplicador del precio de transferencia como un campo separado. Por eso el tablero muestra el efecto reportado de las transferencias internas y no inventa un multiplicador.')

st.divider(); st.caption('KAIRO · Dashboard de control de gestión · datos leídos directamente de los Excel de resultados CESIM')
