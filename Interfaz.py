import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO


# Backend
def procesar_excel(archivo_excel):
    data = pd.read_excel(archivo_excel)

    posicion = -1
    leer_nombres = False
    caracteristicas = {}

    for celda in data['Unnamed: 0']:
        posicion += 1 
        if celda == "Nombre del Trabajador":
            leer_nombres = True
            continue

        if leer_nombres and pd.notna(celda):
            if data.iloc[posicion, 0] == "Párametros":
                break
            nombre = data.iloc[posicion, 0]
            identificador = data.iloc[posicion, 2]
            cargo = data.iloc[posicion, 7]
            caracteristicas[nombre, identificador, cargo] = posicion

    data_procesada = pd.DataFrame(columns=['Nombre', 'ID', 'Cargo', 'Código de pagos', 'Horas'])
    caracteristicas_items = list(caracteristicas.items())

    for i in range(len(caracteristicas_items) - 1):
        llave, fila = caracteristicas_items[i]
        siguiente_llave, siguiente_fila = caracteristicas_items[i + 1]

        for pos_fila in range(fila, siguiente_fila):
            if pd.notna(data.iloc[pos_fila, 17]):
                nuevo_registro = {
                    'Nombre': str(llave[0]),
                    'ID': int(llave[1]),
                    'Cargo': str(llave[2]),
                    'Código de pagos': str(data.iloc[pos_fila, 17]),
                    'Horas': float(data.iloc[pos_fila, 23]) if pd.notna(data.iloc[pos_fila, 23]) else np.nan
                }
                data_procesada = pd.concat([data_procesada, pd.DataFrame([nuevo_registro])], ignore_index=True)

    return data_procesada


# Interfaz Streamlit (Frontend)
st.title("Automatización Payroll Transfer")

archivo = st.file_uploader("Sube el archivo Excel", type=["xls", "xlsx"])

if archivo is not None:
    try:
        df_procesado = procesar_excel(archivo)
        st.success("Procesamiento completado. Vista previa:")

        st.dataframe(df_procesado.head())

        # Convertir a Excel para descargar
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_procesado.to_excel(writer, index=False, sheet_name='Procesado')
            workbook = writer.book
            worksheet = writer.sheets['Procesado']

            # Formato para encabezados
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D9E1F2',  # Color claro (azul pálido)
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            # Aplicar formato a los encabezados
            for col_num, value in enumerate(df_procesado.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # Formato para las celdas del cuerpo (con bordes)
            cell_format = workbook.add_format({'border': 1})

            # Aplicar formato con bordes a todas las celdas del DataFrame
            for row in range(1, len(df_procesado) + 1):
                for col in range(len(df_procesado.columns)):
                    worksheet.write(row, col, df_procesado.iloc[row - 1, col], cell_format)

        output.seek(0)

        st.download_button(
            label="📥 Descargar Excel procesado",
            data=output,
            file_name="Payroll Transfer Procesado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Ocurrió un error: {e}")
