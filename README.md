# Evaluador Ambiental - Método Conesa

Aplicación web en **Python + Streamlit** para evaluación matricial de impactos ambientales.

## Incluye
- Datos generales del proyecto.
- Acciones editables.
- Factores ambientales editables y ponderación UIP.
- Matriz acción × factor.
- Valoración con IN, EX, MO, PE, RV, SI, AC, EF, PR y MC.
- Fórmula: `I = ±(3IN + 2EX + MO + PE + RV + SI + AC + EF + PR + MC)`.
- Clasificación automática: compatible, moderado, severo y crítico.
- Comparación sin medidas / con medidas.
- Impacto residual.
- Tabla y gráfico de resultados.
- Exportación a JSON, Excel y PDF.
- Importación de proyectos JSON.

## Ejecutar en tu computadora
1. Instala Python 3.11 o superior.
2. Abre una terminal dentro de esta carpeta.
3. Instala dependencias:
   `pip install -r requirements.txt`
4. Ejecuta:
   `streamlit run app.py`
5. Se abrirá automáticamente en tu navegador.

## Publicarlo en Internet
La carpeta está preparada para subirse a GitHub y desplegarse en Streamlit Community Cloud. Una vez probado, se puede publicar con una URL pública y luego trabajar su indexación en Google.
