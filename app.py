
import io, json
from datetime import datetime
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

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
    a = abs(i)
    if i > 0:
        if a < 25:
            return "Positivo bajo"
        if a < 50:
            return "Positivo moderado"
        if a < 75:
            return "Positivo alto"
        return "Positivo muy alto"

    if a < 25:
        return "Compatible"
    if a < 50:
        return "Moderado"
    if a < 75:
        return "Severo"
    return "Crítico"

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
    columns = [
        "N°","Acción","Factor","Naturaleza","I sin medidas","Clase inicial",
        "I residual","Clase residual","Reducción |I|","Medida"
    ]
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
    return pd.DataFrame(rows, columns=columns)

def excel_bytes():
    b = io.BytesIO()

    proyecto_df = pd.DataFrame([st.session_state.project])
    acciones_df = pd.DataFrame({"Acciones": list(st.session_state.actions)})
    factores_df = st.session_state.factors.copy()
    resultados_df = df_results()

    # Asegurar estructura válida aun cuando todavía no existan interacciones.
    if factores_df.empty:
        factores_df = pd.DataFrame(columns=["factor", "medio", "uip"])

    with pd.ExcelWriter(b, engine="openpyxl", mode="w") as writer:
        proyecto_df.to_excel(writer, sheet_name="Proyecto", index=False)
        acciones_df.to_excel(writer, sheet_name="Acciones", index=False)
        factores_df.to_excel(writer, sheet_name="Factores", index=False)
        resultados_df.to_excel(writer, sheet_name="Matriz_Impactos", index=False)

    b.seek(0)
    return b.getvalue()

def pdf_bytes():
    b = io.BytesIO()
    doc = SimpleDocTemplate(
        b,
        pagesize=A4,
        leftMargin=1.8*cm,
        rightMargin=1.8*cm,
        topMargin=1.6*cm,
        bottomMargin=1.6*cm
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="TituloInforme",
        parent=styles["Title"],
        fontSize=16,
        leading=20,
        alignment=1,
        spaceAfter=10
    ))
    styles.add(ParagraphStyle(
        name="SubtituloInforme",
        parent=styles["Heading2"],
        fontSize=11,
        leading=14,
        alignment=1,
        spaceAfter=12
    ))
    styles.add(ParagraphStyle(
        name="SeccionInforme",
        parent=styles["Heading2"],
        fontSize=12,
        leading=15,
        spaceBefore=8,
        spaceAfter=6
    ))
    styles.add(ParagraphStyle(
        name="TextoInforme",
        parent=styles["BodyText"],
        fontSize=9,
        leading=13,
        alignment=4,
        spaceAfter=7
    ))
    styles.add(ParagraphStyle(
        name="TextoPequeno",
        parent=styles["BodyText"],
        fontSize=7.3,
        leading=9
    ))

    story = []
    p = st.session_state.project
    d = df_results()

    # ---------------- Portada / cabecera ----------------
    story.append(Paragraph(
        "INFORME DE EVALUACIÓN DE IMPACTO AMBIENTAL",
        styles["TituloInforme"]
    ))
    story.append(Paragraph(
        "Aplicación del método matricial de Conesa",
        styles["SubtituloInforme"]
    ))
    story.append(Spacer(1, 0.15*cm))

    datos = [
        ["Proyecto", p.get("nombre", "")],
        ["Ubicación", p.get("ubicacion", "")],
        ["Responsable / evaluador", p.get("responsable", "")],
        ["Etapa principal", p.get("etapa", "")],
        ["Descripción", p.get("descripcion", "")]
    ]
    tabla_datos = Table(datos, colWidths=[4.0*cm, 12.5*cm])
    tabla_datos.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.4,colors.grey),
        ("BACKGROUND",(0,0),(0,-1),colors.lightgrey),
        ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),8),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",(0,0),(-1,-1),5),
        ("RIGHTPADDING",(0,0),(-1,-1),5),
        ("TOPPADDING",(0,0),(-1,-1),4),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),
    ]))
    story.append(tabla_datos)
    story.append(Spacer(1, 0.35*cm))

    # ---------------- Introducción ----------------
    story.append(Paragraph("1. Introducción", styles["SeccionInforme"]))
    story.append(Paragraph(
        "El presente informe resume la evaluación de impactos ambientales realizada "
        "mediante una matriz acción–factor basada en el método de Conesa. La herramienta "
        "permite identificar interacciones entre las actividades del proyecto y los factores "
        "ambientales, asignar atributos de importancia, comparar el escenario sin medidas "
        "con el escenario posterior a la aplicación de medidas preventivas o correctoras, "
        "y visualizar el impacto residual.",
        styles["TextoInforme"]
    ))
    story.append(Paragraph(
        "Los resultados deben interpretarse como un apoyo técnico y académico. La selección "
        "de interacciones, los valores asignados a cada atributo y las medidas de manejo "
        "deben sustentarse con la línea base ambiental, el expediente técnico, mediciones "
        "de campo, normativa aplicable y criterio profesional.",
        styles["TextoInforme"]
    ))

    # ---------------- Metodología ----------------
    story.append(Paragraph("2. Metodología de evaluación", styles["SeccionInforme"]))
    story.append(Paragraph(
        "La importancia de cada interacción se obtiene mediante la expresión: "
        "<b>I = ±(3IN + 2EX + MO + PE + RV + SI + AC + EF + PR + MC)</b>, "
        "donde IN corresponde a intensidad, EX a extensión, MO a momento, PE a persistencia, "
        "RV a reversibilidad, SI a sinergia, AC a acumulación, EF a efecto, PR a periodicidad "
        "y MC a recuperabilidad. El signo indica si la interacción es negativa o positiva.",
        styles["TextoInforme"]
    ))
    story.append(Paragraph(
        "Para impactos negativos, la magnitud se interpreta como Compatible cuando |I| < 25, "
        "Moderado cuando 25 ≤ |I| < 50, Severo cuando 50 ≤ |I| < 75 y Crítico cuando |I| ≥ 75. "
        "Para impactos positivos, la aplicación expresa el resultado como beneficio positivo "
        "bajo, moderado, alto o muy alto, evitando utilizar categorías de severidad negativa.",
        styles["TextoInforme"]
    ))

    # ---------------- Resumen ejecutivo ----------------
    story.append(Paragraph("3. Resumen de resultados", styles["SeccionInforme"]))

    if d.empty:
        story.append(Paragraph(
            "No se registraron interacciones ambientales en la evaluación.",
            styles["TextoInforme"]
        ))
    else:
        total = len(d)
        neg = int((d["I residual"] < 0).sum())
        pos = int((d["I residual"] > 0).sum())
        high_neg = int(((d["I residual"] < 0) & d["Clase residual"].isin(["Severo","Crítico"])).sum())
        improved = int((d["Reducción |I|"] > 0).sum())

        story.append(Paragraph(
            f"Se evaluaron <b>{total}</b> interacciones ambientales. De ellas, "
            f"<b>{neg}</b> presentan naturaleza negativa y <b>{pos}</b> naturaleza positiva. "
            f"Después de incorporar medidas de manejo, <b>{improved}</b> interacciones reducen "
            f"su magnitud absoluta y se registran <b>{high_neg}</b> impactos negativos residuales "
            f"en las categorías Severo o Crítico.",
            styles["TextoInforme"]
        ))

        # ---------------- Tabla principal ----------------
        story.append(Paragraph("4. Matriz de resultados", styles["SeccionInforme"]))
        cols = [
            "N°","Acción","Factor","Naturaleza","I sin medidas",
            "Clase inicial","I residual","Clase residual","Reducción |I|"
        ]
        table_data = [[Paragraph(str(c), styles["TextoPequeno"]) for c in cols]]
        for _, row in d[cols].iterrows():
            table_data.append([
                Paragraph(str(row[c]), styles["TextoPequeno"]) for c in cols
            ])

        tabla = Table(
            table_data,
            repeatRows=1,
            colWidths=[0.6*cm,3.2*cm,2.7*cm,1.5*cm,1.45*cm,1.8*cm,1.35*cm,1.8*cm,1.5*cm]
        )
        tabla.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.3,colors.grey),
            ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("FONTSIZE",(0,0),(-1,-1),7),
            ("LEFTPADDING",(0,0),(-1,-1),3),
            ("RIGHTPADDING",(0,0),(-1,-1),3),
            ("TOPPADDING",(0,0),(-1,-1),3),
            ("BOTTOMPADDING",(0,0),(-1,-1),3),
        ]))
        story.append(tabla)
        story.append(Spacer(1, 0.3*cm))

        # ---------------- Interpretación por interacción ----------------
        story.append(Paragraph("5. Interpretación de las interacciones", styles["SeccionInforme"]))
        for _, row in d.iterrows():
            naturaleza = row["Naturaleza"]
            accion = row["Acción"]
            factor = row["Factor"]
            i0 = int(row["I sin medidas"])
            ir = int(row["I residual"])
            c0 = row["Clase inicial"]
            cr = row["Clase residual"]
            reduccion = int(row["Reducción |I|"])
            medida = str(row.get("Medida","")).strip()

            if naturaleza == "Negativo":
                if reduccion > 0:
                    interpretacion = (
                        f"La interacción <b>{accion} → {factor}</b> presenta inicialmente "
                        f"una importancia de <b>{i0}</b> ({c0}). Luego de aplicar medidas de manejo, "
                        f"el impacto residual disminuye a <b>{ir}</b> ({cr}), con una reducción "
                        f"de <b>{reduccion}</b> unidades de importancia."
                    )
                else:
                    interpretacion = (
                        f"La interacción <b>{accion} → {factor}</b> presenta una importancia "
                        f"inicial de <b>{i0}</b> ({c0}) y un impacto residual de <b>{ir}</b> ({cr}). "
                        "No se observa reducción de la magnitud, por lo que conviene revisar "
                        "la medida propuesta o los atributos asignados."
                    )
            else:
                interpretacion = (
                    f"La interacción <b>{accion} → {factor}</b> es positiva. Su importancia "
                    f"es de <b>{i0}</b> y el resultado residual es <b>{ir}</b>, clasificado como "
                    f"<b>{cr}</b>. Este valor representa un beneficio asociado al proyecto y no "
                    "una severidad ambiental negativa."
                )

            story.append(Paragraph(interpretacion, styles["TextoInforme"]))
            if medida:
                story.append(Paragraph(
                    f"<b>Medida registrada:</b> {medida}",
                    styles["TextoInforme"]
                ))

        # ---------------- Medidas ----------------
        story.append(Paragraph("6. Medidas de manejo ambiental", styles["SeccionInforme"]))
        medidas = d[d["Medida"].astype(str).str.strip() != ""][["Acción","Factor","Medida"]]
        if medidas.empty:
            story.append(Paragraph(
                "No se registraron medidas preventivas o correctoras.",
                styles["TextoInforme"]
            ))
        else:
            med_data = [[
                Paragraph("Acción", styles["TextoPequeno"]),
                Paragraph("Factor", styles["TextoPequeno"]),
                Paragraph("Medida propuesta", styles["TextoPequeno"])
            ]]
            for _, row in medidas.iterrows():
                med_data.append([
                    Paragraph(str(row["Acción"]), styles["TextoPequeno"]),
                    Paragraph(str(row["Factor"]), styles["TextoPequeno"]),
                    Paragraph(str(row["Medida"]), styles["TextoPequeno"])
                ])
            mt = Table(med_data, repeatRows=1, colWidths=[4.0*cm,3.5*cm,9.2*cm])
            mt.setStyle(TableStyle([
                ("GRID",(0,0),(-1,-1),0.3,colors.grey),
                ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                ("VALIGN",(0,0),(-1,-1),"TOP"),
                ("FONTSIZE",(0,0),(-1,-1),7.2),
                ("LEFTPADDING",(0,0),(-1,-1),4),
                ("RIGHTPADDING",(0,0),(-1,-1),4),
                ("TOPPADDING",(0,0),(-1,-1),4),
                ("BOTTOMPADDING",(0,0),(-1,-1),4),
            ]))
            story.append(mt)

        # ---------------- Conclusiones ----------------
        story.append(Paragraph("7. Conclusiones", styles["SeccionInforme"]))

        negativos = d[d["I residual"] < 0]
        positivos = d[d["I residual"] > 0]

        if not negativos.empty:
            peor_neg = negativos.loc[negativos["I residual"].abs().idxmax()]
            story.append(Paragraph(
                f"• El impacto negativo residual de mayor magnitud corresponde a "
                f"<b>{peor_neg['Acción']} → {peor_neg['Factor']}</b>, con "
                f"I = <b>{int(peor_neg['I residual'])}</b> y clasificación "
                f"<b>{peor_neg['Clase residual']}</b>.",
                styles["TextoInforme"]
            ))

        if not positivos.empty:
            mejor_pos = positivos.loc[positivos["I residual"].abs().idxmax()]
            story.append(Paragraph(
                f"• El beneficio positivo de mayor magnitud corresponde a "
                f"<b>{mejor_pos['Acción']} → {mejor_pos['Factor']}</b>, con "
                f"I = <b>{int(mejor_pos['I residual'])}</b> y clasificación "
                f"<b>{mejor_pos['Clase residual']}</b>.",
                styles["TextoInforme"]
            ))

        story.append(Paragraph(
            "• La comparación entre el escenario inicial y el escenario con medidas permite "
            "visualizar el efecto de la mitigación y priorizar las interacciones que requieren "
            "mayor control o seguimiento.",
            styles["TextoInforme"]
        ))

        # ---------------- Recomendaciones ----------------
        story.append(Paragraph("8. Recomendaciones", styles["SeccionInforme"]))
        story.append(Paragraph(
            "• Verificar que las medidas propuestas sean aplicables al diseño definitivo del proyecto "
            "y asignar indicadores de seguimiento para comprobar su eficacia durante la ejecución.",
            styles["TextoInforme"]
        ))
        story.append(Paragraph(
            "• Revisar especialmente los impactos residuales que permanezcan en categoría Moderado, "
            "Severo o Crítico y, de ser necesario, reforzar las medidas preventivas y correctoras.",
            styles["TextoInforme"]
        ))
        story.append(Paragraph(
            "• Complementar la evaluación con información de línea base, campañas de monitoreo, "
            "expediente técnico y normativa ambiental aplicable antes de utilizar los resultados "
            "como sustento de una evaluación ambiental formal.",
            styles["TextoInforme"]
        ))

    story.append(Spacer(1, 0.25*cm))
    story.append(Paragraph(
        "Reporte generado automáticamente por el Evaluador de Impacto Ambiental – Método Conesa | Perú.",
        styles["TextoPequeno"]
    ))

    doc.build(story)
    b.seek(0)
    return b.getvalue()

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
        ["Negativo", "|I| < 25", "Compatible"],
        ["Negativo", "25 ≤ |I| < 50", "Moderado"],
        ["Negativo", "50 ≤ |I| < 75", "Severo"],
        ["Negativo", "|I| ≥ 75", "Crítico"],
        ["Positivo", "|I| < 25", "Positivo bajo"],
        ["Positivo", "25 ≤ |I| < 50", "Positivo moderado"],
        ["Positivo", "50 ≤ |I| < 75", "Positivo alto"],
        ["Positivo", "|I| ≥ 75", "Positivo muy alto"]
    ],columns=["Naturaleza","Rango","Clasificación"]),hide_index=True,use_container_width=True)

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
        if r["signo"] > 0:
            st.caption("En impactos positivos, la magnitud se expresa como nivel de beneficio y no como severidad ambiental negativa.")
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
        t=len(d); neg=int((d["I residual"]<0).sum()); pos=int((d["I residual"]>0).sum()); high=int(((d["I residual"] < 0) & d["Clase residual"].isin(["Severo","Crítico"])).sum())
        a,b,c,e=st.columns(4); a.metric("Interacciones",t); b.metric("Negativas",neg); c.metric("Positivas",pos); e.metric("Negativas severas / críticas",high)
        st.dataframe(d,hide_index=True,use_container_width=True)
        chart=d.copy(); chart["Interacción"]=chart["Acción"].str[:24]+" → "+chart["Factor"].str[:24]
        st.bar_chart(chart.set_index("Interacción")["I residual"])
        worst=d.loc[d["I residual"].abs().idxmax()]
        tipo = "beneficio" if int(worst["I residual"]) > 0 else "impacto negativo"
        st.info(
            f"Interacción residual de mayor magnitud: {worst['Acción']} → {worst['Factor']} "
            f"(|I|={abs(int(worst['I residual']))}, {worst['Clase residual']}; {tipo})."
        )

with tabs[6]:
    st.subheader("Exportar / importar")
    payload={"project":st.session_state.project,"actions":st.session_state.actions,"factors":st.session_state.factors.to_dict("records"),"interactions":st.session_state.interactions,"exported_at":datetime.now().isoformat(timespec="seconds")}
    j=json.dumps(payload,ensure_ascii=False,indent=2).encode("utf-8")
    c1,c2,c3=st.columns(3)
    c1.download_button("⬇️ JSON",j,"evaluacion_ambiental_conesa.json","application/json",use_container_width=True)

    try:
        excel_data = excel_bytes()
        c2.download_button(
            "⬇️ Excel", excel_data, "matriz_evaluacion_ambiental.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    except Exception as exc:
        c2.error(f"No se pudo generar Excel: {exc}")

    try:
        pdf_data = pdf_bytes()
        c3.download_button(
            "⬇️ PDF", pdf_data, "reporte_evaluacion_ambiental.pdf",
            "application/pdf", use_container_width=True
        )
    except Exception as exc:
        c3.error(f"No se pudo generar PDF: {exc}")
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
