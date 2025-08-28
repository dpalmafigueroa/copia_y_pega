import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io

st.set_page_config(layout="wide")

# 🔄 Botón para reiniciar la app desde el sidebar
st.sidebar.button("🔄 Reiniciar app", on_click=lambda: st.session_state.clear())

st.title("Pegado de datos automatizado 📂")
st.write("Sube tus archivos para pegar datos de la base a la plantilla.")

# Inicializar estado si no existe
if "procesado" not in st.session_state:
    st.session_state.procesado = False
    st.session_state.output_file = None
    st.session_state.longitud_max = 0

# Widgets para subir archivos
uploaded_file_base = st.file_uploader("Sube tu archivo base (base_limpia.xlsx)", type=["xlsx"])
uploaded_file_template = st.file_uploader("Sube tu archivo plantilla (wb.xlsx)", type=["xlsx"])

# Campos para los nombres de las hojas y la fila de encabezados
base_sheet_name = st.text_input("Nombre de la hoja de la base (donde se tomarán los datos)")
template_sheet_name = st.text_input("Nombre de la hoja de la plantilla (donde se vaciarán los datos)")
headers_row = st.number_input("Ingresa la fila de encabezados de la plantilla", min_value=1, value=1)

# Campo para definir la fila de inicio
start_row = st.number_input("Ingresa la fila de inicio para el pegado", min_value=1, value=3426)

# Botón de procesamiento
if st.button("Procesar y Pegar Datos"):
    if uploaded_file_base and uploaded_file_template and base_sheet_name and template_sheet_name:
        try:
            # Leer archivos
            df_base = pd.read_excel(uploaded_file_base, sheet_name=base_sheet_name, engine="openpyxl")
            wb = load_workbook(uploaded_file_template)
            ws = wb[template_sheet_name]

            # Obtener encabezados de la plantilla
            headers_plantilla = {
                str(cell.value).strip(): cell.column
                for cell in ws[int(headers_row)] if cell.value
            }

            # Mapear columnas coincidentes
            mapeo = {col: col for col in df_base.columns if col in headers_plantilla}

            if not mapeo:
                st.error("❌ No hay coincidencia entre columnas de la base y la plantilla.")
            else:
                longitud_max = len(df_base)
                for col_base, col_plantilla in mapeo.items():
                    col_idx = headers_plantilla[col_plantilla]
                    datos = df_base[col_base].tolist()
                    for r, valor in enumerate(datos, start=int(start_row)):
                        valor_final = "" if pd.isna(valor) else valor
                        ws.cell(row=r, column=col_idx, value=valor_final)

                # Guardar a memoria
                output = io.BytesIO()
                wb.save(output)

                # Actualizar estado
                st.session_state.output_file = output
                st.session_state.procesado = True
                st.session_state.longitud_max = longitud_max

        except KeyError as e:
            st.error(f"❌ Error: Verifica el nombre de la hoja. {e}")
        except Exception as e:
            st.error(f"❌ Ocurrió un error: {e}")
    else:
        st.error("❌ Por favor, sube ambos archivos y llena todos los campos.")

# Placeholder para el botón de descarga
descarga_placeholder = st.empty()

# Mostrar botón de descarga si se procesó correctamente
if st.session_state.procesado and st.session_state.output_file:
    with descarga_placeholder:
        st.success(f"✅ ¡Pegado de {st.session_state.longitud_max} filas completado desde la fila {start_row}!")
        st.download_button(
            label="📥 Descargar archivo modificado",
            data=st.session_state.output_file.getvalue(),
            file_name="wb_modificado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
