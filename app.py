import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import io

st.set_page_config(layout="wide")

# Botón para reiniciar la app desde el sidebar
st.sidebar.button("🔄 Reiniciar", on_click=lambda: st.session_state.clear())

st.title("Pegado de datos automatizado 📂")
st.write("Sube tus archivos para pegar datos de la base a la plantilla.")

# --- LÓGICA DE PROCESAMIENTO ---
def procesar_archivos_optimizados(base_file, template_file, base_sheet, template_sheet, headers_row, start_row):
    try:
        df_base = pd.read_excel(base_file, sheet_name=base_sheet, engine="openpyxl")
        wb = load_workbook(template_file)
        ws = wb[template_sheet]

        headers_plantilla = {
            str(cell.value).strip(): cell.column
            for cell in ws[int(headers_row)] if cell.value
        }
        
        columnas_comunes = [col for col in df_base.columns if col in headers_plantilla]
        
        if not columnas_comunes:
            raise ValueError("No se encontraron columnas coincidentes entre la base y la plantilla.")

        df_filtrado = df_base[columnas_comunes]

        rows_to_paste = dataframe_to_rows(df_filtrado, index=False, header=False)
        
        for r_idx, row in enumerate(rows_to_paste, start=int(start_row)):
            for col_name, value in zip(df_filtrado.columns, row):
                col_idx = headers_plantilla[col_name]
                ws.cell(row=r_idx, column=col_idx, value=value)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return output, len(df_filtrado)

    except KeyError as e:
        st.error(f"❌ Error: La hoja '{e.args[0]}' no existe en el archivo. Por favor, verifica el nombre.")
        return None, 0
    except Exception as e:
        st.error(f"❌ Ocurrió un error inesperado durante el procesamiento: {e}")
        return None, 0

# --- INTERFAZ DE USUARIO DENTRO DE UN FORMULARIO ---
with st.form(key='my_form'):
    # Inicializar estado si no existe (esto va fuera del formulario)
    if "procesado" not in st.session_state:
        st.session_state.procesado = False
        st.session_state.output_file = None
        st.session_state.longitud_max = 0

    uploaded_file_base = st.file_uploader("Sube tu archivo base (con los datos a copiar)", type=["xlsx"])
    uploaded_file_template = st.file_uploader("Sube tu archivo de plantilla (donde se pegarán los datos)", type=["xlsx"])

    base_sheet_name = st.text_input("Nombre de la hoja de la base (ej. 'Sheet1')", value="Sheet1")
    template_sheet_name = st.text_input("Nombre de la hoja de la plantilla (ej. 'Workbook Consolidado')", value="Workbook Consolidado")
    headers_row = st.number_input("Ingresa la fila de encabezados de la plantilla", min_value=1, value=1)
    start_row = st.number_input("Ingresa la fila de inicio para el pegado", min_value=1, value=3426)
    
    # El botón de procesamiento debe estar DENTRO del formulario
    submit_button = st.form_submit_button("🚀 Procesar y Pegar Datos")

# --- LÓGICA DE EJECUCIÓN (AHORA SE ACTIVA CON EL BOTÓN DEL FORMULARIO) ---
if submit_button:
    if uploaded_file_base and uploaded_file_template and base_sheet_name and template_sheet_name:
        with st.spinner("Procesando... por favor, espera."):
            output_file, longitud_max = procesar_archivos_optimizados(
                uploaded_file_base,
                uploaded_file_template,
                base_sheet_name,
                template_sheet_name,
                headers_row,
                start_row
            )
        
        if output_file:
            st.session_state.procesado = True
            st.session_state.output_file = output_file
            st.session_state.longitud_max = longitud_max
    else:
        st.error("❌ Por favor, sube ambos archivos y llena todos los campos de las hojas.")

# El botón de descarga se muestra solo si la variable de estado es verdadera
if st.session_state.procesado and st.session_state.output_file:
    st.success(f"✅ ¡Pegado de {st.session_state.longitud_max} filas completado desde la fila {start_row}!")
    st.download_button(
        label="📥 Descargar archivo modificado",
        data=st.session_state.output_file,
        file_name="wb_modificado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )