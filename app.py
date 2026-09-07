
import io, json
from datetime import datetime
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

st.set_page_config(
    page_title="Evaluador de Impacto Ambiental – Método Conesa | Perú",
    page_icon="🌱", layout="wide"
)

st.markdown("""
<style>
h1 {font-weight:800}
[data-testid="stMetricValue"] {font-size:1.8rem}
.hero{padding:16px;border:1px solid #ddd;border-radius:12px;margin-bottom:14px}
.note{font-size:.9rem;color:#666}
</style>
""", unsafe_allow_html=True)

ATTRS = {
"IN":("Intensidad",{1:"Baja",2:"Media",4:"Alta",8:"Muy alta",12:"Total"}),
"EX":("Extensión",{1:"Puntual",2:"Parcial",4:"Extensa",8:"Total"}),
"MO":("Momento",{1:"Largo plazo",2:"Medio plazo",4:"Inmediato"}),
"PE":("Persistencia",{1:"Fugaz",2:"Temporal",4:"Permanente"}),
"RV":("Reversibilidad",{1:"Corto plazo",2:"Medio plazo",4:"Irreversible"}),
"SI":("Sinergia",{1:"Simple",2:"Sinérgico",4:"Muy sinérgico"}),
"AC":("Acumulación",{1:"Simple",4:"Acumulativo"}),
"EF":("Efecto",{1:"Indirecto",4:"Directo"}),
"PR":("Periodicidad",{1:"Irregular",2:"Periódico",4:"Continuo"}),
"MC":("Recuperabilidad",{1:"Inmediata",2:"Medio plazo",4:"Mitigable",8:"Irrecuperable"})
}

ACTIONS = [
"Campamento y almacén","Patio de máquinas","Desvío provisional del tránsito",
"Desbroce ribereño","Cantera y extracción de agregados",
"Depósito de material excedente (DME)","Movimiento de tierras y excavación",
"Obras de concreto (estribos, tablero)","Montaje de vigas metálicas",
"Defensas ribereñas","Tránsito de maquinaria y transporte",
"Cierre y revegetación","Tránsito vehicular sobre el puente"
]

FACTORS = [
["Suelo y geomorfología","Físico",90],["Agua superficial","Físico",140],
["Dinámica de cauce y socavación","Físico",120],["Calidad del aire","Físico",55],
["Ruido y vibraciones","Físico",45],["Paisaje","Físico",50],
["Flora ribereña","Biológico",70],["Fauna terrestre","Biológico",45],
["Hidrobiología","Biológico",95],["Economía y uso agrícola","Socioeconómico",70],
["Accesibilidad","Socioeconómico",110],["Seguridad y salud","Socioeconómico",55],
["Patrimonio cultural","Cultural",30],["Empleo e ingresos","Socioeconómico",45],
["Calidad de vida y aceptación","Socioeconómico",30]
]

def importance(sign, v):
    return int(sign*(3*v["IN"]+2*v["EX"]+v["MO"]+v["PE"]+v["RV"]+v["SI"]+v["AC"]+v["EF"]+v["PR"]+v["MC"]))

def cls(i):
    a=abs(i)
    return "Compatible" if a<25 else "Moderado" if a<50 else "Severo" if a<75 else "Crítico"

def init():
    st.session_state.setdefault("project",{
        "nombre":"Construcción del Puente Chugsén–Matara",
        "ubicacion":"Matara, Cajamarca, Perú","responsable":"",
        "etapa":"Construcción",
        "descripcion":"Evaluación ambiental matricial del proyecto de infraestructura."
    })
    st.session_state.setdefault("actions",ACTIONS.copy())
    st.session_state.setdefault("factors",pd.DataFrame(FACTORS,columns=["factor","medio","uip"]))
    st.session_state.setdefault("interactions",[])
init()

def df_results():
    rows=[]
    for n,r in enumerate(st.session_state.interactions,1):
        rows.append({
            "N°":n,"Acción":r["accion"],"Factor":r["factor"],
            "Naturaleza":"Positivo" if r["signo"]>0 else "Negativo",
            "I sin medidas":r["antes"],"Clase inicial":r["clase_antes"],
            "I residual":r["despues"],"Clase residual":r["clase_despues"],
            "Reducción |I|":abs(r["antes"])-abs(r["despues"]),
            "Medida":r.get("medida","")
        })
    return pd.DataFrame(rows)

def excel_bytes():
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine="openpyxl") as w:
        pd.DataFrame([st.session_state.project]).to_excel(w,"Proyecto",index=False)
        pd.DataFrame({"Acciones":st.session_state.actions}).to_excel(w,"Acciones",index=False)
        st.session_state.factors.to_excel(w,"Factores",index=False)
        df_results().to_excel(w,"Matriz_Impactos",index=False)
    b.seek(0); return b.getvalue()

def pdf_bytes():
    b=io.BytesIO()
    doc=SimpleDocTemplate(b,pagesize=landscape(A4),leftMargin=cm,rightMargin=cm,topMargin=cm,bottomMargin=cm)
    s=getSampleStyleSheet()
    story=[Paragraph("EVALUADOR DE IMPACTO AMBIENTAL – MÉTODO CONESA",s["Title"]),Spacer(1,.3*cm)]
    p=st.session_state.project
    data=[["Proyecto",p["nombre"]],["Ubicación",p["ubicacion"]],["Responsable",p["responsable"]],["Etapa",p["etapa"]]]
    t=Table(data,colWidths=[4*cm,21*cm]); t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.5,colors.grey),("BACKGROUND",(0,0),(0,-1),colors.lightgrey)]))
    story += [t,Spacer(1,.4*cm)]
    df=df_results()
    if df.empty: story.append(Paragraph("No existen interacciones registradas.",s["BodyText"]))
    else:
        cols=["N°","Acción","Factor","I sin medidas","Clase inicial","I residual","Clase residual","Medida"]
        d=[cols]+df[cols].astype(str).values.tolist()
        tt=Table(d,repeatRows=1,colWidths=[.7*cm,4.4*cm,4.1*cm,1.8*cm,2.2*cm,1.6*cm,2.1*cm,8*cm])
        tt.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.35,colors.grey),("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("FONTSIZE",(0,0),(-1,-1),7)]))
        story.append(tt)
    doc.build(story); b.seek(0); return b.getvalue()

st.title("🌱 Evaluador de Impacto Ambiental – Método Conesa | Perú")
st.caption("Herramienta web para identificar, valorar y comparar impactos ambientales en proyectos de ingeniería mediante una matriz acción–factor.")

st.markdown("""<div class="hero"><b>Evaluación ambiental editable</b><br>
Registra acciones y factores, calcula la importancia del impacto, compara escenarios sin medidas y con medidas, analiza el impacto residual y exporta los resultados.</div>""",unsafe_allow_html=True)

with st.sidebar:
    st.header("Guía rápida")
    st.write("1. Proyecto")
    st.write("2. Acciones y factores")
    st.write("3. Identificación")
    st.write("4. Valoración")
    st.write("5. Resultados")
    st.write("6. Exportación")
    st.divider()
    st.latex(r"I=\pm(3IN+2EX+MO+PE+RV+SI+AC+EF+PR+MC)")
    st.caption("Herramienta académica de apoyo. La valoración final debe sustentarse técnicamente.")

tabs=st.tabs(["Inicio","1. Proyecto","2. Acciones y factores","3. Matriz de identificación","4. Valoración","5. Resultados","6. Exportar / Importar"])

with tabs[0]:
    st.subheader("Metodología de trabajo")
    a,b,c=st.columns(3)
    a.info("**Identificación**\n\nRelaciona acciones del proyecto con factores ambientales.")
    b.info("**Valoración**\n\nAsigna los atributos del método de Conesa.")
    c.info("**Mitigación**\n\nCompara el impacto inicial con el impacto residual.")
    st.markdown("#### Clasificación por importancia")
    st.dataframe(pd.DataFrame([
        ["|I| < 25","Compatible"],["25 ≤ |I| < 50","Moderado"],
        ["50 ≤ |I| < 75","Severo"],["|I| ≥ 75","Crítico"]
    ],columns=["Rango","Clasificación"]),hide_index=True,use_container_width=True)

with tabs[1]:
    st.subheader("Datos generales del proyecto")
    p=st.session_state.project
    c1,c2=st.columns(2)
    p["nombre"]=c1.text_input("Nombre del proyecto",p["nombre"])
    p["ubicacion"]=c1.text_input("Ubicación",p["ubicacion"])
    p["responsable"]=c1.text_input("Responsable / evaluador",p["responsable"])
    etapas=["Preliminar","Construcción","Cierre","Operación"]
    p["etapa"]=c2.selectbox("Etapa principal",etapas,index=etapas.index(p["etapa"]) if p["etapa"] in etapas else 1)
    p["descripcion"]=c2.text_area("Descripción breve",p["descripcion"],height=125)
    st.success("Puedes reemplazar todos los datos y utilizar la aplicación para cualquier proyecto.")

with tabs[2]:
    st.subheader("Inventario editable")
    ca,cf=st.columns(2)
    with ca:
        st.markdown("#### Acciones")
        new=st.text_input("Nueva acción")
        if st.button("Agregar acción") and new.strip() and new.strip() not in st.session_state.actions:
            st.session_state.actions.append(new.strip()); st.rerun()
        for i,x in enumerate(st.session_state.actions):
            x1,x2=st.columns([8,1]); x1.write(f"{i+1}. {x}")
            if x2.button("✕",key=f"da{i}"): st.session_state.actions.pop(i); st.rerun()
    with cf:
        st.markdown("#### Factores ambientales y UIP")
        st.session_state.factors=st.data_editor(
            st.session_state.factors,num_rows="dynamic",use_container_width=True,
            column_config={
                "medio":st.column_config.SelectboxColumn("Medio",options=["Físico","Biológico","Socioeconómico","Cultural"]),
                "uip":st.column_config.NumberColumn("UIP",min_value=0,step=1)
            },key="feditor"
        ).reset_index(drop=True)
        total=int(pd.to_numeric(st.session_state.factors["uip"],errors="coerce").fillna(0).sum())
        (st.success if total==1000 else st.warning)(f"Ponderación total: {total} UIP.")

with tabs[3]:
    st.subheader("Matriz de identificación acción × factor")
    factors=[str(x) for x in st.session_state.factors["factor"].dropna() if str(x).strip()]
    if not st.session_state.actions or not factors:
        st.warning("Agrega acciones y factores antes de continuar.")
    else:
        action=st.selectbox("Acción",st.session_state.actions)
        chosen=st.multiselect("Factores afectados",factors)
        nat=st.radio("Naturaleza preliminar",["Negativo","Positivo"],horizontal=True)
        if st.button("Agregar relaciones",type="primary"):
            n=0
            for f in chosen:
                if not any(r["accion"]==action and r["factor"]==f for r in st.session_state.interactions):
                    sign=-1 if nat=="Negativo" else 1
                    attrs={k:1 for k in ATTRS}
                    I=importance(sign,attrs)
                    st.session_state.interactions.append({"accion":action,"factor":f,"signo":sign,"attrs_antes":attrs.copy(),"attrs_despues":attrs.copy(),"antes":I,"despues":I,"clase_antes":cls(I),"clase_despues":cls(I),"medida":""}); n+=1
            st.success(f"Relaciones nuevas agregadas: {n}")
    d=df_results()
    if not d.empty: st.dataframe(d[["N°","Acción","Factor","Naturaleza"]],hide_index=True,use_container_width=True)

with tabs[4]:
    st.subheader("Valoración cualitativa")
    if not st.session_state.interactions:
        st.warning("Primero agrega interacciones.")
    else:
        labels=[f"{i+1}. {r['accion']} → {r['factor']}" for i,r in enumerate(st.session_state.interactions)]
        idx=st.selectbox("Interacción",range(len(labels)),format_func=lambda i:labels[i])
        r=st.session_state.interactions[idx]
        nat=st.radio("Naturaleza",["Negativo","Positivo"],index=0 if r["signo"]<0 else 1,horizontal=True,key=f"nat{idx}")
        r["signo"]=-1 if nat=="Negativo" else 1
        c1,c2=st.columns(2); va={}; vd={}
        with c1:
            st.markdown("### Sin medidas")
            for k,(name,optsmap) in ATTRS.items():
                opts=list(optsmap)
                cur=r["attrs_antes"].get(k,opts[0])
                va[k]=st.selectbox(f"{k} · {name}",opts,index=opts.index(cur) if cur in opts else 0,format_func=lambda x,m=optsmap:f"{x} · {m[x]}",key=f"a{idx}{k}")
        with c2:
            st.markdown("### Con medidas")
            for k,(name,optsmap) in ATTRS.items():
                opts=list(optsmap)
                cur=r["attrs_despues"].get(k,opts[0])
                vd[k]=st.selectbox(f"{k} · {name}",opts,index=opts.index(cur) if cur in opts else 0,format_func=lambda x,m=optsmap:f"{x} · {m[x]}",key=f"d{idx}{k}")
        ia,idd=importance(r["signo"],va),importance(r["signo"],vd)
        m1,m2,m3=st.columns(3); m1.metric("I sin medidas",ia); m2.metric("I residual",idd); m3.metric("Reducción |I|",abs(ia)-abs(idd))
        st.write(f"**Clasificación inicial:** {cls(ia)} | **Clasificación residual:** {cls(idd)}")
        if abs(idd)>abs(ia): st.warning("El impacto residual es mayor que el inicial. Revisa los atributos con medidas.")
        medida=st.text_area("Medida preventiva / correctora",r.get("medida",""),key=f"med{idx}")
        if st.button("Guardar valoración",type="primary"):
            r.update({"attrs_antes":va,"attrs_despues":vd,"antes":ia,"despues":idd,"clase_antes":cls(ia),"clase_despues":cls(idd),"medida":medida})
            st.session_state.interactions[idx]=r; st.success("Valoración guardada.")

with tabs[5]:
    st.subheader("Resultados")
    d=df_results()
    if d.empty: st.info("Aún no existen resultados.")
    else:
        t=len(d); neg=int((d["I residual"]<0).sum()); pos=int((d["I residual"]>0).sum()); high=int(d["Clase residual"].isin(["Severo","Crítico"]).sum())
        a,b,c,e=st.columns(4); a.metric("Interacciones",t); b.metric("Negativas",neg); c.metric("Positivas",pos); e.metric("Severas / críticas",high)
        st.dataframe(d,hide_index=True,use_container_width=True)
        chart=d.copy(); chart["Interacción"]=chart["Acción"].str[:24]+" → "+chart["Factor"].str[:24]
        st.bar_chart(chart.set_index("Interacción")["I residual"])
        worst=d.loc[d["I residual"].abs().idxmax()]
        st.info(f"Interacción residual de mayor magnitud: {worst['Acción']} → {worst['Factor']} (|I|={abs(int(worst['I residual']))}, {worst['Clase residual']}).")

with tabs[6]:
    st.subheader("Exportar / importar")
    payload={"project":st.session_state.project,"actions":st.session_state.actions,"factors":st.session_state.factors.to_dict("records"),"interactions":st.session_state.interactions,"exported_at":datetime.now().isoformat(timespec="seconds")}
    j=json.dumps(payload,ensure_ascii=False,indent=2).encode("utf-8")
    c1,c2,c3=st.columns(3)
    c1.download_button("⬇️ JSON",j,"evaluacion_ambiental_conesa.json","application/json",use_container_width=True)
    c2.download_button("⬇️ Excel",excel_bytes(),"matriz_evaluacion_ambiental.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    c3.download_button("⬇️ PDF",pdf_bytes(),"reporte_evaluacion_ambiental.pdf","application/pdf",use_container_width=True)
    up=st.file_uploader("Importar proyecto JSON",type=["json"])
    if up:
        try:
            data=json.loads(up.getvalue().decode("utf-8"))
            if st.button("Cargar proyecto"):
                st.session_state.project=data.get("project",st.session_state.project)
                st.session_state.actions=data.get("actions",st.session_state.actions)
                st.session_state.factors=pd.DataFrame(data.get("factors",FACTORS))
                st.session_state.interactions=data.get("interactions",[])
                st.rerun()
        except Exception as e: st.error(f"No se pudo leer el archivo: {e}")

st.divider()
st.markdown('<div class="note">Evaluador de Impacto Ambiental – Método Conesa | Perú · Herramienta académica de apoyo para proyectos de ingeniería.</div>',unsafe_allow_html=True)
