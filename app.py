import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io

st.set_page_config(layout="wide")

st.title("Pegado de datos automatizado 📂")
st.write("Sube tus archivos para pegar datos de la base a la plantilla.")

# Widgets para subir archivos
uploaded_file_base = st.file_uploader("Sube tu archivo base (base_limpia.xlsx)", type=["xlsx"])
uploaded_file_template = st.file_uploader("Sube tu archivo plantilla (wb.xlsx)", type=["xlsx"])

# Campo para definir la fila de inicio
start_row = st.number_input("Ingresa la fila de inicio para el pegado", min_value=1, value=3426)

if st.button("Procesar y Pegar Datos"):
    if uploaded_file_base and uploaded_file_template:
        try:
            # 1. Leer archivos desde la memoria
            df_base = pd.read_excel(uploaded_file_base, engine="openpyxl")
            wb = load_workbook(uploaded_file_template)
            ws = wb["Workbook Consolidado"]
    
            # 2. Encabezados de la plantilla (fila 1)
            headers_plantilla = {str(cell.value).strip(): cell.column for cell in ws[1] if cell.value}
    
            # 3. Mapear columnas que coinciden
            mapeo = {col: col for col in df_base.columns if col in headers_plantilla}
            if not mapeo:
                st.error("Error: No hay coincidencia entre columnas de la base y la plantilla.")
            else:
                # 4. Pegar datos desde la fila indicada
                longitud_max = len(df_base)
                for col_base, col_plantilla in mapeo.items():
                    col_idx = headers_plantilla[col_plantilla]
                    datos = df_base[col_base].tolist()
                    for r, valor in enumerate(datos, start=int(start_row)):
                        ws.cell(row=r, column=col_idx, value=valor)
    
                # 5. Guardar cambios en la memoria
                output = io.BytesIO()
                wb.save(output)
    
                st.success(f"✅ ¡Pegado de {longitud_max} filas completado desde la fila {start_row}!")
    
                # Botón de descarga
                st.download_button(
                    label="Descargar archivo modificado",
                    data=output.getvalue(),
                    file_name="wb_modificado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
        except KeyError as e:
            st.error(f"Error: La hoja 'Workbook Consolidado' no se encontró en la plantilla. {e}")
        except Exception as e:
            st.error(f"Ocurrió un error en el proceso: {e}")
    else:
        st.error("Por favor, sube ambos archivos para continuar.")