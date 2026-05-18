import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# Configuración de la página web
st.set_page_config(page_title="Procesador de Visitas", page_icon="📊", layout="centered")

st.title("📊 Control de Visitas Asistenciales")
st.markdown("Subí el archivo consolidado mensual para filtrar las visitas **Liberadas** y excluir al personal médico.")

# Componente para subir el archivo Excel o CSV
uploaded_file = st.file_uploader("Seleccioná el archivo de Visitas (Excel o CSV)", type=["xlsx", "csv"])

if uploaded_file is not None:
    try:
        # 1. Leer el archivo dinámicamente según su extensión
        if uploaded_file.name.endswith('.csv'):
            # Detectamos si viene con la fila vacía inicial saltando 1 línea si es necesario
            df = pd.read_csv(uploaded_file, skiprows=1)
        else:
            df = pd.read_excel(uploaded_file, skiprows=1)
            
        # Pequeña validación de columnas críticas
        columnas_requeridas = ['EstadoCoordinacion', 'TipoProfesional', 'Profesional', 'Paciente']
        if not all(col in df.columns for col in columnas_requeridas):
            st.error("El archivo no tiene el formato esperado. Asegurate de que incluya las columnas: " + ", ".join(columnas_requeridas))
        else:
            st.success("¡Archivo cargado con éxito! Procesando datos...")

            # 2. Aplicar filtros solicitados
            df_filtered = df[(df['EstadoCoordinacion'] == 'Liberada') & (df['TipoProfesional'] != 'Medico')]

            # 3. Agrupar y sumarizar por Profesional y Paciente
            summary = df_filtered.groupby(['Profesional', 'TipoProfesional', 'Paciente']).size().reset_index(name='Visitas Liberadas')
            summary = summary.sort_values(by=['Profesional', 'Paciente']).reset_index(drop=True)

            # Mostrar una vista previa en la web
            st.markdown("### Vista previa del resultado:")
            st.dataframe(summary, use_container_width=True)

            # 4. Generar el archivo Excel formateado en memoria
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Resumen Auxiliares"
            ws.views.sheetView[0].showGridLines = True

            # Estilos estéticos
            header_fill = PatternFill(start_color="1F4E5B", end_color="1F4E5B", fill_type="solid")
            header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
            zebra_fill = PatternFill(start_color="F4F8F9", end_color="F4F8F9", fill_type="solid")
            white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            title_font = Font(name="Arial", size=14, bold=True, color="1F4E5B")
            regular_font = Font(name="Arial", size=10)
            bold_font = Font(name="Arial", size=10, bold=True)

            thin_border_side = Side(style='thin', color='D3D3D3')
            thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

            # Encabezados de reporte internos
            ws['A1'] = "Control de Visitas (Excluyendo Médicos)"
            ws['A1'].font = title_font
            ws['A2'] = "Filtrado exclusivo: Visitas LIBERADAS"
            ws['A2'].font = Font(name="Arial", size=10, italic=True, color="555555")

            headers = ["Profesional", "Tipo Profesional", "Paciente", "Visitas Liberadas"]
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=4, column=col_idx, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center" if col_idx == 4 else "left", vertical="center")
                cell.border = thin_border
            ws.row_dimensions[4].height = 25

            # Volcar datos al Excel
            start_row = 5
            for idx, row in summary.iterrows():
                r_idx = start_row + idx
                c1 = ws.cell(row=r_idx, column=1, value=row['Profesional'])
                c2 = ws.cell(row=r_idx, column=2, value=row['TipoProfesional'])
                c3 = ws.cell(row=r_idx, column=3, value=row['Paciente'])
                c4 = ws.cell(row=r_idx, column=4, value=int(row['Visitas Liberadas']))
                
                c1.alignment = Alignment(horizontal="left", vertical="center")
                c2.alignment = Alignment(horizontal="left", vertical="center")
                c3.alignment = Alignment(horizontal="left", vertical="center")
                c4.alignment = Alignment(horizontal="right", vertical="center")
                c4.number_format = '#,##0'
                
                current_fill = zebra_fill if idx % 2 == 0 else white_fill
                for cell in [c1, c2, c3, c4]:
                    cell.font = regular_font
                    cell.fill = current_fill
                    cell.border = thin_border
                ws.row_dimensions[r_idx].height = 18

            # Fila de Total General
            tot_row = start_row + len(summary)
            ws.cell(row=tot_row, column=1, value="Total General").font = bold_font
            ws.cell(row=tot_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
            ws.cell(row=tot_row, column=1).border = thin_border
            
            for c_idx in [2, 3]:
                ws.cell(row=tot_row, column=c_idx, value="").border = thin_border

            total_cell = ws.cell(row=tot_row, column=4, value=f"=SUM(D5:D{tot_row-1})")
            total_cell.font = bold_font
            total_cell.alignment = Alignment(horizontal="right", vertical="center")
            total_cell.number_format = '#,##0'
            total_cell.border = thin_border

            ws.freeze_panes = "A5"

            # Auto-ajustar ancho de columnas
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.row < 4: continue
                    if cell.value: max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

            wb.save(output)
            processed_data = output.getvalue()

            # Botón de descarga para el usuario
            st.markdown("---")
            st.download_button(
                label="📥 Descargar Excel Formateado",
                data=processed_data,
                file_name="Resumen_Visitas_Procesado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
