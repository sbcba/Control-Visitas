import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import datetime
import holidays

# Configuración de la página web
st.set_page_config(page_title="Auditoría Avanzada de Visitas", page_icon="⚙️", layout="wide")

st.title("⚙️ Sistema Avanzado de Auditoría y Liquidación Asistencial")
st.markdown("Cargá el reporte mensual completo. Seleccioná los filtros en la barra lateral y presioná **Aplicar Filtros** para ver los resultados.")

# Inicializar el estado del botón si no existe (arranca en Falso para no mostrar nada)
if 'filtros_aplicados' not in st.session_state:
    st.session_state.filtros_aplicados = False

# Componente para subir el archivo
uploaded_file = st.file_uploader("Subí el archivo de Visitas completo (Excel o CSV)", type=["xlsx", "csv"])

if uploaded_file is not None:
    try:
        # 1. Leer el archivo saltando la fila vacía inicial
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, skiprows=1)
        else:
            df = pd.read_excel(uploaded_file, skiprows=1)
            
        # Validar columnas requeridas del Excel completo
        columnas_requeridas = ['EstadoCoordinacion', 'TipoModulo', 'TipoProfesional', 'Profesional', 'Paciente', 'PlanCuidado', 'FechaInicioProF', 'FechaFinProf', 'DuracionMinutosProf']
        if not all(col in df.columns for col in columnas_requeridas):
            st.error("El archivo no tiene todas las columnas requeridas. Verificá el formato.")
        else:
            # Filtro base: Solo visitas Liberadas
            df_base = df[df['EstadoCoordinacion'] == 'Liberada'].copy()
            
            # Convertir fechas a objetos reales de Python
            df_base['FechaInicioProF'] = pd.to_datetime(df_base['FechaInicioProF'], errors='coerce')
            df_base['FechaFinProf'] = pd.to_datetime(df_base['FechaFinProf'], errors='coerce')
            
            # Cargar calendario oficial de feriados de Argentina
            ar_holidays = holidays.Argentina()
            
            # 2. SECCIÓN DE FILTROS DENTRO DE UN FORMULARIO (En la barra lateral)
            st.sidebar.header("🎯 Panel de Filtros")
            
            # Listas de opciones disponibles basadas en el archivo subido
            modulos_disponibles = sorted(list(df_base['TipoModulo'].dropna().unique()))
            profesionales_disponibles = sorted(list(df_base['TipoProfesional'].dropna().unique()))
            default_prof = [p for p in profesionales_disponibles if p != 'Medico']

            # Creamos el formulario en la barra lateral
            with st.sidebar.form(key='formulario_filtros'):
                st.markdown("### Seleccioná tus criterios:")
                
                # Desplegables de selección múltiple (tildar/destildar)
                modulos_seleccionados = st.multiselect("Tipos de Módulo:", modulos_disponibles, default=modulos_disponibles)
                profesionales_seleccionados = st.multiselect("Tipos de Profesional:", profesionales_disponibles, default=default_prof)
                
                st.markdown("---")
                # Al presionar el botón, cambiamos el estado de la aplicación a True
                boton_aceptar = st.form_submit_button(label="✅ Aplicar Filtros")
                if boton_aceptar:
                    st.session_state.filtros_aplicados = True
            
            # 3. CONTROL DE RENDERIZADO: Solo procesar y mostrar si "filtros_aplicados" es True
            if not st.session_state.filtros_aplicados:
                st.info("💡 Por favor, seleccioná los criterios en la barra lateral de la izquierda y hacé clic en el botón **'Aplicar Filtros'** para procesar la información mensual.")
            else:
                # Aplicar filtros dinámicos seleccionados por el usuario
                df_filtered = df_base[
                    df_base['TipoModulo'].isin(modulos_seleccionados) & 
                    df_base['TipoProfesional'].isin(profesionales_seleccionados)
                ].copy()
                
                Rows_Auditoria = []
                
                for idx, row in df_filtered.iterrows():
                    paciente = row['Paciente']
                    modulo = row['TipoModulo']
                    prestacion = row['PlanCuidado']
                    profesional = row['Profesional']
                    tipo_prof = row['TipoProfesional']
                    minutos = row['DuracionMinutosProf']
                    start = row['FechaInicioProF']
                    end = row['FechaFinProf']
                    
                    # Valores por defecto
                    horas_totales = 0.0
                    horas_feriado = 0.0
                    es_feriado_visita = "NO"
                    visitas_comunes = 0
                    visitas_feriado = 0
                    
                    if pd.isna(start):
                        continue
                    if pd.isna(end):
                        end = start
                    
                    # Verificar si el rango toca algún feriado argentino
                    rango_dias = pd.date_range(start=start.date(), end=end.date()).date
                    toca_feriado = any(d in ar_holidays for d in rango_dias)
                    if toca_feriado:
                        es_feriado_visita = "SÍ"
                    
                    # Caso A: ENFERMERO GUARDIA (Se liquida por Horas)
                    if tipo_prof == "Enfermero Guardia":
                        if not pd.isna(minutos) and minutos > 0:
                            horas_totales = round((minutos / 60.0) * 2) / 2
                            
                            if toca_feriado:
                                if start.date() == end.date():
                                    if start.date() in ar_holidays:
                                        horas_feriado = horas_totales
                                else:
                                    medianoche = datetime.datetime.combine(start.date() + datetime.timedelta(days=1), datetime.time.min)
                                    min_dia1 = (medianoche - start).total_seconds() / 60.0
                                    min_dia2 = (end - medianoche).total_seconds() / 60.0
                                    
                                    min_feriado = 0.0
                                    if start.date() in ar_holidays:
                                        min_feriado += min_dia1
                                    if end.date() in ar_holidays:
                                        min_feriado += min_dia2
                                        
                                    horas_feriado = round((min_feriado / 60.0) * 2) / 2
                    
                    # Caso B: OTROS PROFESIONALES (Se liquida por Visita unidad)
                    else:
                        if es_feriado_visita == "SÍ":
                            visitas_feriado = 1
                        else:
                            visitas_comunes = 1
                    
                    Rows_Auditoria.append({
                        'Paciente': paciente,
                        'TipoModulo': modulo,
                        'PlanCuidado': prestacion,
                        'Profesional': profesional,
                        'TipoProfesional': tipo_prof,
                        'Visitas Comunes': visitas_comunes,
                        'Visitas Feriado': visitas_feriado,
                        'Horas Totales Guardia': horas_totales,
                        'Horas Feriado Guardia': horas_feriado
                    })
                    
                # 4. AGRUPACIÓN Y CONSOLIDACIÓN DE RESULTADOS
                if len(Rows_Auditoria) == 0:
                    st.warning("No hay registros que coincidan con los filtros seleccionados.")
                else:
                    df_audit = pd.DataFrame(Rows_Auditoria)
                    
                    summary = df_audit.groupby(
                        ['Paciente', 'TipoModulo', 'PlanCuidado', 'Profesional', 'TipoProfesional']
                    ).agg({
                        'Visitas Comunes': 'sum',
                        'Visitas Feriado': 'sum',
                        'Horas Totales Guardia': 'sum',
                        'Horas Feriado Guardia': 'sum'
                    }).reset_index()
                    
                    def format_horas_texto(val):
                        if val == 0: return "-"
                        if val % 1 == 0: return f"{int(val)} hs"
                        return f"{int(val)} hs y media"
                    
                    summary_view = summary.copy()
                    summary_view['Horas Totales Guardia'] = summary_view['Horas Totales Guardia'].apply(format_horas_texto)
                    summary_view['Horas Feriado Guardia'] = summary_view['Horas Feriado Guardia'].apply(format_horas_texto)
                    
                    st.markdown(f"### 📋 Resultados de la Auditoría ({len(summary)} combinaciones encontradas)")
                    st.dataframe(summary_view, use_container_width=True)
                    
                    # 5. GENERAR EL EXCEL FORMATEADO EN MEMORIA
                    output = io.BytesIO()
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Auditoría de Liquidación"
                    ws.views.sheetView[0].showGridLines = True
                    
                    header_fill = PatternFill(start_color="1F4E5B", end_color="1F4E5B", fill_type="solid")
                    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
                    zebra_fill = PatternFill(start_color="F4F8F9", end_color="F4F8F9", fill_type="solid")
                    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                    feriado_alerta_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                    
                    thin_side = Side(style='thin', color='D3D3D3')
                    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
                    
                    ws['A1'] = "Consolidado de Auditoría, Prestaciones y Liquidación Horas/Visitas"
                    ws['A1'].font = Font(name="Arial", size=14, bold=True, color="1F4E5B")
                    ws['A2'] = f"Generado el: {datetime.datetime.now().strftime('%d/%m/%Y')} - Filtros aplicados bajo confirmación"
                    ws['A2'].font = Font(name="Arial", size=9, italic=True, color="666666")
                    
                    headers = [
                        "Paciente", "Módulo", "Plan de Cuidado / Prestación", 
                        "Profesional", "Tipo Profesional", "Visitas Comunes", 
                        "Visitas Feriado", "Horas Totales Guardia", "Horas Feriado Guardia"
                    ]
                    
                    for col_idx, h in enumerate(headers, 1):
                        cell = ws.cell(row=4, column=col_idx, value=h)
                        cell.font = header_font
                        cell.fill = header_fill
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        cell.border = thin_border
                    ws.row_dimensions[4].height = 28
                    
                    start_row = 5
                    for idx, row in summary.iterrows():
                        r_idx = start_row + idx
                        ws.cell(row=r_idx, column=1, value=row['Paciente']).alignment = Alignment(horizontal="left")
                        ws.cell(row=r_idx, column=2, value=row['TipoModulo']).alignment = Alignment(horizontal="center")
                        ws.cell(row=r_idx, column=3, value=row['PlanCuidado']).alignment = Alignment(horizontal="left")
                        ws.cell(row=r_idx, column=4, value=row['Profesional']).alignment = Alignment(horizontal="left")
                        ws.cell(row=r_idx, column=5, value=row['TipoProfesional']).alignment = Alignment(horizontal="left")
                        
                        v_com = ws.cell(row=r_idx, column=6, value=int(row['Visitas Comunes']))
                        v_fer = ws.cell(row=r_idx, column=7, value=int(row['Visitas Feriado']))
                        v_com.number_format = '#,##0'
                        v_fer.number_format = '#,##0'
                        
                        h_tot = ws.cell(row=r_idx, column=8, value=float(row['Horas Totales Guardia']))
                        h_fer = ws.cell(row=r_idx, column=9, value=float(row['Horas Feriado Guardia']))
                        h_tot.number_format = '#,##0.0'
                        h_fer.number_format = '#,##0.0'
                        
                        for c_num in [6, 7, 8, 9]:
                            ws.cell(row=r_idx, column=c_num).alignment = Alignment(horizontal="right")
                        
                        if row['Visitas Feriado'] > 0 or row['Horas Feriado Guardia'] > 0:
                            row_fill = feriado_alerta_fill
                        else:
                            row_fill = zebra_fill if idx % 2 == 0 else white_fill
                            
                        for c_idx in range(1, 10):
                            cell = ws.cell(row=r_idx, column=c_idx)
                            cell.font = Font(name="Arial", size=9)
                            cell.fill = row_fill
                            cell.border = thin_border
                            
                        ws.row_dimensions[r_idx].height = 18
                    
                    tot_row = start_row + len(summary)
                    ws.cell(row=tot_row, column=1, value="Total General").font = Font(name="Arial", size=9, bold=True)
                    ws.cell(row=tot_row, column=1).border = thin_border
                    for c_empty in range(2, 6):
                        ws.cell(row=tot_row, column=c_empty).border = thin_border
                        
                    col_letters = ['F', 'G', 'H', 'I']
                    for col_let in col_letters:
                        c_idx_f = headers.index(headers[5 + col_letters.index(col_let)]) + 1
                        res_cell = ws.cell(row=tot_row, column=c_idx_f, value=f"=SUM({col_let}5:{col_let}{tot_row-1})")
                        res_cell.font = Font(name="Arial", size=9, bold=True)
                        res_cell.border = thin_border
                        if col_let in ['H', 'I']:
                            res_cell.number_format = '#,##0.0'
                        else:
                            res_cell.number_format = '#,##0'
                    
                    ws.freeze_panes = "A5"
                    
                    for col in ws.columns:
                        max_len = 0
                        col_letter = get_column_letter(col[0].column)
                        for cell in col:
                            if cell.row < 4: continue
                            if cell.value: max_len = max(max_len, len(str(cell.value)))
                        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)
                        
                    wb.save(output)
                    processed_data = output.getvalue()
                    
                    st.markdown("---")
                    st.download_button(
                        label="📥 Descargar Excel de Auditoría Potenciado",
                        data=processed_data,
                        file_name="Auditoria_Liquidacion_Avanzada.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
    except Exception as e:
        st.error(f"Error general en el procesamiento de columnas: {e}")
