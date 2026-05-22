import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import holidays

# Configuración de la página web
st.set_page_config(page_title="Auditoría de Prestaciones", page_icon="📊", layout="wide")

st.title("📊 Control de Prestaciones y Visitas por Paciente")
st.markdown("Subí el archivo mensual para desglosar las prestaciones por paciente, profesionales intervinientes y alertas de **días feriados** (Excluye Médicos y visitas No Liberadas).")

# Componente para subir el archivo
uploaded_file = st.file_uploader("Seleccioná el archivo de Visitas (Excel o CSV)", type=["xlsx", "csv"])

if uploaded_file is not None:
    try:
        # 1. Leer el archivo saltando la fila vacía inicial si corresponde
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, skiprows=1)
        else:
            df = pd.read_excel(uploaded_file, skiprows=1)
            
        # Validar columnas necesarias incluyendo ahora 'PlanCuidado' y 'FechaInicioProF'
        columnas_requeridas = ['EstadoCoordinacion', 'TipoProfesional', 'Profesional', 'Paciente', 'PlanCuidado', 'FechaInicioProF']
        if not all(col in df.columns for col in columnas_requeridas):
            st.error("El archivo no tiene el formato esperado. Asegurate de que incluya las columnas: " + ", ".join(columnas_requeridas))
        else:
            st.success("¡Archivo cargado con éxito!")

            # 2. Filtros solicitados (Visitas Liberadas y No Médicos)
            df_filtered = df[(df['EstadoCoordinacion'] == 'Liberada') & (df['TipoProfesional'] != 'Medico')].copy()

            # 3. Procesar fechas para detectar feriados de Argentina
            # Convertimos la columna a formato fecha real por si viene como texto
            df_filtered['Fecha_Date'] = pd.to_datetime(df_filtered['FechaInicioProF'], errors='coerce')
            
            # Configuramos los feriados oficiales de Argentina
            ar_holidays = holidays.Argentina()
            
            # Creamos la columna que identifica si es feriado o no
            df_filtered['¿Es Feriado?'] = df_filtered['Fecha_Date'].apply(lambda x: "SÍ" if x in ar_holidays else "NO")

            # 4. Agrupar sumando las visitas por cada combinación y ver si fue feriado
            # Agrupamos incluyendo '¿Es Feriado?' para separar las visitas comunes de las de días feriados
            summary = df_filtered.groupby(
                ['Paciente', 'PlanCuidado', 'Profesional', 'TipoProfesional', '¿Es Feriado?']
            ).size().reset_index(name='Cantidad Visitas')

            # Ordenar por Paciente para que quede todo agrupado visualmente por la persona que recibe la atención
            summary = summary.sort_values(by=['Paciente', 'PlanCuidado', 'Profesional']).reset_index(drop=True)

            # Mostrar vista previa en la web
            st.markdown("### Vista previa de la auditoría mensual:")
            st.dataframe(summary, use_container_width=True)

            # 5. Generar el archivo Excel profesional
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Auditoría Prestaciones"
            ws.views.sheetView[0].showGridLines = True

            # Estilos de diseño
            header_fill = PatternFill(start_color="1F4E5B", end_color="1F4E5B", fill_type="solid")
            header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
            zebra_fill = PatternFill(start_color="F4F8F9", end_color="F4F8F9", fill_type="solid")
            white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            feriado_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Amarillo tenue para resaltar feriados
            title_font = Font(name="Arial", size=14, bold=True, color="1F4E5B")
            regular_font = Font(name="Arial", size=10)
            bold_font = Font(name="Arial", size=10, bold=True)

            thin_border_side = Side(style='thin', color='D3D3D3')
            thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

            # Encabezados del reporte
            ws['A1'] = "Reporte de Auditoría: Prestaciones y Visitas en Feriados"
            ws['A1'].font = title_font
            ws['A2'] = "Filtro activo: Solo visitas LIBERADAS de personal auxiliar y técnico"
            ws['A2'].font = Font(name="Arial", size=10, italic=True, color="555555")

            headers = ["Paciente", "Prestación (Plan de Cuidado)", "Profesional", "Tipo Profesional", "¿Es Feriado?", "Cantidad Visitas"]
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=4, column=col_idx, value=header)
                cell.font = header_font
                cell.fill = header_fill
                if col_idx in [5, 6]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.border = thin_border
            ws.row_dimensions[4].height = 25

            # Cargar los datos al Excel renglón por renglón
            start_row = 5
            for idx, row in summary.iterrows():
                r_idx = start_row + idx
                c1 = ws.cell(row=r_idx, column=1, value=row['Paciente'])
                c2 = ws.cell(row=r_idx, column=2, value=row['PlanCuidado'])
                c3 = ws.cell(row=r_idx, column=3, value=row['Profesional'])
                c4 = ws.cell(row=r_idx, column=4, value=row['TipoProfesional'])
                c5 = ws.cell(row=r_idx, column=5, value=row['¿Es Feriado?'])
                c6 = ws.cell(row=r_idx, column=6, value=int(row['Cantidad Visitas']))
                
                # Alineaciones
                c1.alignment = Alignment(horizontal="left", vertical="center")
                c2.alignment = Alignment(horizontal="left", vertical="center")
                c3.alignment = Alignment(horizontal="left", vertical="center")
                c4.alignment = Alignment(horizontal="left", vertical="center")
                c5.alignment = Alignment(horizontal="center", vertical="center")
                c6.alignment = Alignment(horizontal="right", vertical="center")
                c6.number_format = '#,##0'
                
                # Color de fondo: si es feriado va con fondo amarillo claro para llamar la atención, si no, va con el intercalado normal (cebra)
                if row['¿Es Feriado?'] == "SÍ":
                    current_fill = feriado_fill
                else:
                    current_fill = zebra_fill if idx % 2 == 0 else white_fill
                
                for cell in [c1, c2, c3, c4, c5, c6]:
                    cell.font = regular_font
                    cell.fill = current_fill
                    cell.border = thin_border
                
                ws.row_dimensions[r_idx].height = 18

            # Fila de Cierre con el Total General de visitas
            tot_row = start_row + len(summary)
            ws.cell(row=tot_row, column=1, value="Total General de Visitas").font = bold_font
            ws.cell(row=tot_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
            ws.cell(row=tot_row, column=1).border = thin_border
            
            for c_idx in range(2, 6):
                ws.cell(row=tot_row, column=c_idx, value="").border = thin_border

            total_cell = ws.cell(row=tot_row, column=6, value=f"=SUM(F5:F{tot_row-1})")
            total_cell.font = bold_font
            total_cell.alignment = Alignment(horizontal="right", vertical="center")
            total_cell.number_format = '#,##0'
            total_cell.border = thin_border

            ws.freeze_panes = "A5"

            # Ajuste inteligente del ancho de las columnas
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.row < 4: continue
                    if cell.value: max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

            wb.save(output)
            processed_data = output.getvalue()

            st.markdown("---")
            st.download_button(
                label="📥 Descargar Excel de Auditoría Completa",
                data=processed_data,
                file_name="Auditoria_Prestaciones_Feriados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
