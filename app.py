import io, json
from datetime import datetime
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

st.set_page_config(page_title='Evaluador Ambiental - Método Conesa', page_icon='🌱', layout='wide')

ATTRS = {
    'IN': ('Intensidad','Grado de destrucción o beneficio',{1:'Baja',2:'Media',4:'Alta',8:'Muy alta',12:'Total'}),
    'EX': ('Extensión','Área de influencia',{1:'Puntual',2:'Parcial',4:'Extensa',8:'Total'}),
    'MO': ('Momento','Plazo de manifestación',{1:'Largo plazo',2:'Medio plazo',4:'Inmediato'}),
    'PE': ('Persistencia','Permanencia del efecto',{1:'Fugaz',2:'Temporal',4:'Permanente'}),
    'RV': ('Reversibilidad','Reconstrucción natural',{1:'Corto plazo',2:'Medio plazo',4:'Irreversible'}),
    'SI': ('Sinergia','Potenciación entre efectos',{1:'Simple',2:'Sinérgico',4:'Muy sinérgico'}),
    'AC': ('Acumulación','Incremento progresivo',{1:'Simple',4:'Acumulativo'}),
    'EF': ('Efecto','Relación causa-efecto',{1:'Indirecto',4:'Directo'}),
    'PR': ('Periodicidad','Regularidad del efecto',{1:'Irregular',2:'Periódico',4:'Continuo'}),
    'MC': ('Recuperabilidad','Reconstrucción mediante intervención humana',{1:'Inmediata',2:'Medio plazo',4:'Mitigable',8:'Irrecuperable'}),
}

DEFAULT_ACTIONS = [
    'Campamento y almacén','Patio de máquinas','Desvío provisional del tránsito','Desbroce ribereño',
    'Cantera y extracción de agregados','Depósito de material excedente (DME)','Movimiento de tierras y excavación',
    'Obras de concreto (estribos, tablero)','Montaje de vigas metálicas','Defensas ribereñas',
    'Tránsito de maquinaria y transporte','Cierre y revegetación','Tránsito vehicular sobre el puente'
]

DEFAULT_FACTORS = [
    {'factor':'Suelo y geomorfología','medio':'Físico','uip':90},
    {'factor':'Agua superficial','medio':'Físico','uip':140},
    {'factor':'Dinámica de cauce y socavación','medio':'Físico','uip':120},
    {'factor':'Calidad del aire','medio':'Físico','uip':55},
    {'factor':'Ruido y vibraciones','medio':'Físico','uip':45},
    {'factor':'Paisaje','medio':'Físico','uip':50},
    {'factor':'Flora ribereña','medio':'Biológico','uip':70},
    {'factor':'Fauna terrestre','medio':'Biológico','uip':45},
    {'factor':'Hidrobiología','medio':'Biológico','uip':95},
    {'factor':'Economía y uso agrícola','medio':'Socioeconómico','uip':70},
    {'factor':'Accesibilidad','medio':'Socioeconómico','uip':110},
    {'factor':'Seguridad y salud','medio':'Socioeconómico','uip':55},
    {'factor':'Patrimonio cultural','medio':'Cultural','uip':30},
    {'factor':'Empleo e ingresos','medio':'Socioeconómico','uip':45},
    {'factor':'Calidad de vida y aceptación','medio':'Socioeconómico','uip':30},
]

def calc_importance(signo, vals):
    base = 3*vals['IN'] + 2*vals['EX'] + vals['MO'] + vals['PE'] + vals['RV'] + vals['SI'] + vals['AC'] + vals['EF'] + vals['PR'] + vals['MC']
    return int(signo*base)

def classify(v):
    a=abs(v)
    if a<25: return 'Compatible'
    if a<50: return 'Moderado'
    if a<75: return 'Severo'
    return 'Crítico'

def init_state():
    if 'project' not in st.session_state:
        st.session_state.project={'nombre':'Construcción del Puente Chugsén–Matara','ubicacion':'Matara, Cajamarca','responsable':'','etapa':'Construcción','descripcion':'Evaluación ambiental matricial del proyecto.'}
    if 'actions' not in st.session_state: st.session_state.actions=DEFAULT_ACTIONS.copy()
    if 'factors' not in st.session_state: st.session_state.factors=pd.DataFrame(DEFAULT_FACTORS)
    if 'interactions' not in st.session_state: st.session_state.interactions=[]

def interactions_df():
    out=[]
    for i,r in enumerate(st.session_state.interactions,1):
        out.append({'N°':i,'Acción':r['accion'],'Factor':r['factor'],'Naturaleza':'Positivo' if r['signo']>0 else 'Negativo','I sin medidas':r['i_antes'],'Clase sin medidas':r['clase_antes'],'I residual':r['i_despues'],'Clase residual':r['clase_despues'],'Reducción |I|':abs(r['i_antes'])-abs(r['i_despues']),'Medida':r.get('medida','')})
    return pd.DataFrame(out)

def export_excel():
    bio=io.BytesIO()
    with pd.ExcelWriter(bio,engine='openpyxl') as w:
        pd.DataFrame([st.session_state.project]).to_excel(w,sheet_name='Proyecto',index=False)
        pd.DataFrame({'Acciones':st.session_state.actions}).to_excel(w,sheet_name='Acciones',index=False)
        st.session_state.factors.to_excel(w,sheet_name='Factores',index=False)
        interactions_df().to_excel(w,sheet_name='Matriz_Impactos',index=False)
        detail=[]
        for i,r in enumerate(st.session_state.interactions,1):
            for esc,attrs in [('Sin medidas',r['attrs_antes']),('Con medidas',r['attrs_despues'])]:
                row={'N°':i,'Acción':r['accion'],'Factor':r['factor'],'Escenario':esc,**attrs}
                detail.append(row)
        pd.DataFrame(detail).to_excel(w,sheet_name='Atributos',index=False)
    return bio.getvalue()

def export_pdf():
    bio=io.BytesIO(); styles=getSampleStyleSheet(); styles.add(ParagraphStyle(name='Small2',parent=styles['BodyText'],fontSize=7,leading=8))
    doc=SimpleDocTemplate(bio,pagesize=landscape(A4),rightMargin=1*cm,leftMargin=1*cm,topMargin=1*cm,bottomMargin=1*cm)
    story=[Paragraph('REPORTE DE EVALUACIÓN DE IMPACTO AMBIENTAL',styles['Title']),Paragraph('Metodología matricial de Conesa',styles['Heading2']),Spacer(1,.2*cm)]
    p=st.session_state.project
    pdata=[['Proyecto',p.get('nombre','')],['Ubicación',p.get('ubicacion','')],['Responsable',p.get('responsable','')],['Etapa',p.get('etapa','')],['Descripción',p.get('descripcion','')]]
    t=Table(pdata,colWidths=[4*cm,21*cm]); t.setStyle(TableStyle([('GRID',(0,0),(-1,-1),.4,colors.grey),('BACKGROUND',(0,0),(0,-1),colors.lightgrey),('VALIGN',(0,0),(-1,-1),'TOP'),('FONTNAME',(0,0),(0,-1),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),8)])); story += [t,Spacer(1,.3*cm),Paragraph('Matriz de impactos',styles['Heading2'])]
    df=interactions_df()
    if df.empty: story.append(Paragraph('No existen interacciones registradas.',styles['BodyText']))
    else:
        cols=['N°','Acción','Factor','Naturaleza','I sin medidas','Clase sin medidas','I residual','Clase residual','Medida']
        data=[[Paragraph(str(c),styles['Small2']) for c in cols]]+[[Paragraph(str(row[c]),styles['Small2']) for c in cols] for _,row in df[cols].iterrows()]
        tab=Table(data,colWidths=[.7*cm,4.1*cm,3.9*cm,1.8*cm,1.6*cm,2.2*cm,1.4*cm,2*cm,7.4*cm],repeatRows=1)
        tab.setStyle(TableStyle([('GRID',(0,0),(-1,-1),.35,colors.grey),('BACKGROUND',(0,0),(-1,0),colors.lightgrey),('VALIGN',(0,0),(-1,-1),'TOP')]))
        story.append(tab)
    doc.build(story); return bio.getvalue()

init_state()
st.title('🌱 Evaluador de Impacto Ambiental')
st.caption('Aplicación web editable basada en la metodología matricial de Conesa. Permite identificar, valorar y comparar impactos antes y después de medidas de manejo.')

tabs=st.tabs(['1. Proyecto','2. Acciones y factores','3. Matriz de identificación','4. Valoración','5. Resultados','6. Exportar / Importar'])

with tabs[0]:
    st.subheader('Datos generales del proyecto')
    c1,c2=st.columns(2)
    with c1:
        st.session_state.project['nombre']=st.text_input('Nombre del proyecto',st.session_state.project['nombre'])
        st.session_state.project['ubicacion']=st.text_input('Ubicación',st.session_state.project['ubicacion'])
        st.session_state.project['responsable']=st.text_input('Responsable / evaluador',st.session_state.project['responsable'])
    with c2:
        etapas=['Preliminar','Construcción','Cierre','Operación']; cur=st.session_state.project['etapa']; idx=etapas.index(cur) if cur in etapas else 1
        st.session_state.project['etapa']=st.selectbox('Etapa principal',etapas,index=idx)
        st.session_state.project['descripcion']=st.text_area('Descripción breve',st.session_state.project['descripcion'],height=125)
    st.info('Puedes reemplazar completamente los datos del Puente Chugsén–Matara y usar la aplicación para otro proyecto.')

with tabs[1]:
    st.subheader('Inventario de acciones y factores')
    ca,cf=st.columns(2)
    with ca:
        st.markdown('#### Acciones del proyecto')
        new_action=st.text_input('Nueva acción',placeholder='Ej. Excavación de estribos')
        if st.button('Agregar acción',use_container_width=True):
            if new_action.strip() and new_action.strip() not in st.session_state.actions:
                st.session_state.actions.append(new_action.strip()); st.rerun()
        for i,a in enumerate(st.session_state.actions):
            x,y=st.columns([8,1]); x.write(f'{i+1}. {a}')
            if y.button('✕',key=f'del_action_{i}'):
                st.session_state.actions.pop(i); st.rerun()
    with cf:
        st.markdown('#### Factores ambientales y ponderación UIP')
        edited=st.data_editor(st.session_state.factors,num_rows='dynamic',use_container_width=True,column_config={'factor':st.column_config.TextColumn('Factor'),'medio':st.column_config.SelectboxColumn('Medio',options=['Físico','Biológico','Socioeconómico','Cultural']),'uip':st.column_config.NumberColumn('UIP',min_value=0,step=1)},key='factor_editor')
        st.session_state.factors=edited.reset_index(drop=True)
        total=int(pd.to_numeric(st.session_state.factors['uip'],errors='coerce').fillna(0).sum())
        if total==1000: st.success('La ponderación suma 1000 UIP.')
        else: st.warning(f'La ponderación actual suma {total} UIP. Ajusta los pesos si deseas mantener 1000 unidades.')

with tabs[2]:
    st.subheader('Matriz acción × factor')
    st.write('Selecciona una acción y marca los factores con los que existe una relación causa–efecto razonable.')
    factor_names=[str(x) for x in st.session_state.factors.get('factor',pd.Series(dtype=str)).dropna().tolist() if str(x).strip()]
    if not st.session_state.actions or not factor_names: st.warning('Agrega al menos una acción y un factor.')
    else:
        action=st.selectbox('Acción',st.session_state.actions,key='identify_action')
        selected=st.multiselect('Factores afectados',factor_names,key='identify_factors')
        naturaleza=st.radio('Naturaleza preliminar',['Negativo','Positivo'],horizontal=True)
        if st.button('Agregar relaciones a la matriz',type='primary'):
            added=0
            for factor in selected:
                if not any(r['accion']==action and r['factor']==factor for r in st.session_state.interactions):
                    signo=-1 if naturaleza=='Negativo' else 1
                    attrs={k:1 for k in ATTRS}
                    i0=calc_importance(signo,attrs)
                    st.session_state.interactions.append({'accion':action,'factor':factor,'signo':signo,'attrs_antes':attrs.copy(),'attrs_despues':attrs.copy(),'i_antes':i0,'i_despues':i0,'clase_antes':classify(i0),'clase_despues':classify(i0),'medida':''}); added+=1
            st.success(f'Se agregaron {added} relaciones nuevas.')
    dfid=interactions_df()
    if not dfid.empty: st.dataframe(dfid[['N°','Acción','Factor','Naturaleza']],use_container_width=True,hide_index=True)

with tabs[3]:
    st.subheader('Valoración cualitativa de importancia')
    if not st.session_state.interactions: st.warning('Primero agrega interacciones en la pestaña 3.')
    else:
        labels=[f"{i+1}. {r['accion']} → {r['factor']}" for i,r in enumerate(st.session_state.interactions)]
        idx=st.selectbox('Interacción a valorar',range(len(labels)),format_func=lambda i: labels[i])
        r=st.session_state.interactions[idx]
        naturaleza=st.radio('Naturaleza del impacto',['Negativo','Positivo'],index=0 if r['signo']<0 else 1,horizontal=True,key=f'sign_{idx}')
        r['signo']=-1 if naturaleza=='Negativo' else 1
        st.latex(r'I=\pm(3IN+2EX+MO+PE+RV+SI+AC+EF+PR+MC)')
        cb,ca=st.columns(2); vb={}; va={}
        with cb:
            st.markdown('### Sin medidas')
            for code,(name,desc,optsmap) in ATTRS.items():
                opts=list(optsmap); cur=r['attrs_antes'].get(code,opts[0]); pos=opts.index(cur) if cur in opts else 0
                vb[code]=st.selectbox(f'{code} · {name}',opts,index=pos,format_func=lambda x,m=optsmap:f'{x} · {m[x]}',key=f'b_{idx}_{code}',help=desc)
        with ca:
            st.markdown('### Con medidas')
            for code,(name,desc,optsmap) in ATTRS.items():
                opts=list(optsmap); cur=r['attrs_despues'].get(code,opts[0]); pos=opts.index(cur) if cur in opts else 0
                va[code]=st.selectbox(f'{code} · {name}',opts,index=pos,format_func=lambda x,m=optsmap:f'{x} · {m[x]}',key=f'a_{idx}_{code}',help=desc)
        ib=calc_importance(r['signo'],vb); ia=calc_importance(r['signo'],va)
        m1,m2,m3=st.columns(3); m1.metric('Importancia sin medidas',ib); m2.metric('Importancia residual',ia); m3.metric('Reducción de |I|',abs(ib)-abs(ia))
        st.write(f'**Clasificación sin medidas:** {classify(ib)}')
        st.write(f'**Clasificación residual:** {classify(ia)}')
        medida=st.text_area('Medida preventiva / correctora',r.get('medida',''),key=f'measure_{idx}')
        if st.button('Guardar valoración',type='primary',use_container_width=True):
            r.update({'attrs_antes':vb,'attrs_despues':va,'i_antes':ib,'i_despues':ia,'clase_antes':classify(ib),'clase_despues':classify(ia),'medida':medida})
            st.session_state.interactions[idx]=r; st.success('Valoración guardada.')

with tabs[4]:
    st.subheader('Resultados e interpretación')
    df=interactions_df()
    if df.empty: st.info('Aún no existen resultados.')
    else:
        total=len(df); neg=int((df['I residual']<0).sum()); pos=int((df['I residual']>0).sum()); high=int(df['Clase residual'].isin(['Severo','Crítico']).sum())
        a,b,c,d=st.columns(4); a.metric('Interacciones',total); b.metric('Negativas',neg); c.metric('Positivas',pos); d.metric('Severas / críticas',high)
        st.dataframe(df,use_container_width=True,hide_index=True)
        chart=df[['Acción','Factor','I residual']].copy(); chart['Interacción']=chart['Acción'].str.slice(0,24)+' → '+chart['Factor'].str.slice(0,24)
        st.bar_chart(chart.set_index('Interacción')['I residual'])
        if high: st.warning(f'Se registran {high} impactos residuales severos o críticos. Deben revisarse prioritariamente las medidas propuestas.')
        else: st.success('No se registran impactos residuales severos o críticos en las valoraciones guardadas. La conclusión debe validarse con línea base y criterio técnico.')

with tabs[5]:
    st.subheader('Guardar, trasladar y exportar el proyecto')
    payload={'project':st.session_state.project,'actions':st.session_state.actions,'factors':st.session_state.factors.to_dict(orient='records'),'interactions':st.session_state.interactions,'exported_at':datetime.now().isoformat(timespec='seconds')}
    c1,c2,c3=st.columns(3)
    c1.download_button('⬇️ Descargar proyecto JSON',json.dumps(payload,ensure_ascii=False,indent=2).encode('utf-8'),'evaluacion_ambiental_conesa.json','application/json',use_container_width=True)
    c2.download_button('⬇️ Descargar Excel',export_excel(),'matriz_evaluacion_ambiental.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
    c3.download_button('⬇️ Descargar PDF',export_pdf(),'reporte_evaluacion_ambiental.pdf','application/pdf',use_container_width=True)
    uploaded=st.file_uploader('Importar proyecto JSON',type=['json'])
    if uploaded is not None:
        try:
            d=json.loads(uploaded.getvalue().decode('utf-8'))
            if st.button('Cargar proyecto importado',type='primary'):
                st.session_state.project=d.get('project',st.session_state.project)
                st.session_state.actions=d.get('actions',st.session_state.actions)
                st.session_state.factors=pd.DataFrame(d.get('factors',DEFAULT_FACTORS))
                st.session_state.interactions=d.get('interactions',[])
                st.success('Proyecto importado correctamente.'); st.rerun()
        except Exception as e: st.error(f'No se pudo leer el archivo: {e}')
    st.caption('La herramienta automatiza cálculos y organización. La selección de interacciones, atributos, ponderaciones y medidas debe sustentarse técnicamente para cada proyecto.')
