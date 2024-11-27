import pandas as pd
from io import BytesIO
import xlsxwriter

image_path = 'static/images/lgb.jpg' 

def procesar_excel_mes(file):
    try:
        # Leer el archivo Excel y seleccionar las columnas necesarias
        df = pd.read_excel(file, usecols=[
            'Fecha de Asignación', 'Servicio', 'Fecha de cierre', 'ESTATUS ETA', 'tecnico'
        ])
        df = df.dropna(how='all')  # Eliminar filas donde todas las columnas son NaN
        df = df.dropna(axis=1, how='all')  # Eliminar columnas donde todas las filas son NaN
        df = df.where(pd.notnull(df), None)  # Reemplazar NaN con None
        
        # Convertir las columnas de fecha a formato datetime
        df['Fecha de Asignación'] = pd.to_datetime(df['Fecha de Asignación'], errors='coerce')
        df['Fecha de cierre'] = pd.to_datetime(df['Fecha de cierre'], errors='coerce')
        
        # Filtrar filas con fechas válidas
        df = df.dropna(subset=['Fecha de Asignación'])

        return df
    except Exception as e:
        raise ValueError(f"Error al procesar el archivo Excel: {str(e)}")

def indicadores_por_mes_y_servicio(df):
    try:
        # Asegúrate de que la columna de fecha sea del tipo datetime
        df['Dia'] = df['Fecha de Asignación'].dt.strftime('%Y-%m-%d')

        # Crear un DataFrame vacío para los resultados
        resultados = []

        # Agrupar por 'Servicio'
        for servicio in df['Servicio'].unique():
            # Contar todas las asignaciones para el servicio
            asignaciones = df[df['Servicio'] == servicio]
            conteo_asignaciones = asignaciones.groupby('Dia').size().reindex(df['Dia'].unique(), fill_value=0)
            for dia, conteo in conteo_asignaciones.items():
                resultados.append((servicio, 'ASIGNACION', dia, conteo))

            # Contar cerradas donde 'ESTATUS ETA' sea 'Completada'
            cerradas = df[(df['Servicio'] == servicio) & (df['ESTATUS ETA'] == 'Completada')]
            conteo_cerradas = cerradas.groupby('Dia').size().reindex(df['Dia'].unique(), fill_value=0)
            for dia, conteo in conteo_cerradas.items():
                resultados.append((servicio, 'CERRADAS', dia, conteo))

            # Contar técnicos
            tecnicos = df[(df['Servicio'] == servicio) & (df['tecnico'].notna())]
            conteo_tecnicos = tecnicos.groupby('Dia').size().reindex(df['Dia'].unique(), fill_value=0)
            for dia, conteo in conteo_tecnicos.items():
                resultados.append((servicio, 'TECNICOS', dia, conteo))

        # Crear un DataFrame a partir de los resultados
        resultados_df = pd.DataFrame(resultados, columns=['Servicio', 'Estatus', 'Dia', 'Conteo'])

        # Pivotar el DataFrame para tener las fechas como columnas
        resultados_df = resultados_df.pivot_table(index=['Servicio', 'Estatus'], columns='Dia', values='Conteo', fill_value=0).reset_index()

        return resultados_df
    except Exception as e:
        raise ValueError(f"Error al generar indicadores: {str(e)}")

def generar_excel_mes(indicadores):
    output = BytesIO()
    image_path = 'static/images/lgb.jpg'
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Crear la hoja principal
        crear_hoja_indicadores(writer, workbook, 'Indicadores por Mes', indicadores, image_path)

        # Crear otras hojas con diferentes indicadores
        #crear_hoja_resumen_semanal(writer, workbook, 'RESUMEN SEMANAL')
        #crear_hoja_suspendidas(writer, workbook, 'SUSPENDIDAS')
        #crear_hoja_estatus_atencion(writer, workbook, 'ESTATUS DE ATENCION')
        #crear_hoja_reporte_tiempo_instalacion(writer, workbook, 'REPORTE TIEMPO DE INSTALACION')
        #crear_hoja_promedio_instalacion(writer, workbook, 'PROMEDIO DE INSTALACION')

    output.seek(0)
    return output.getvalue()

def crear_hoja_indicadores(writer, workbook, sheet_name, indicadores, image_path):
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet

    # Títulos personalizados
    title = f"{sheet_name.upper()}"
    worksheet.write('A1', title, workbook.add_format({'bold': True, 'font_size': 14}))

    # Insertar imagen
    worksheet.insert_image('A3', image_path, {'x_scale': 1, 'y_scale': 1.5})

    # Definir formatos
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'align': 'center',
        'bg_color': '#F0F0F0',
        'border': 1
    })

    cell_format = workbook.add_format({
        'align': 'center',
        'border': 1,
        'text_wrap': True  # Ajuste de texto
    })

    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd', 'align': 'center', 'border': 1})

    # Aplicar formato a los encabezados
    for col_num, value in enumerate(indicadores.columns.values):
        worksheet.write(10, col_num, value, header_format)

    # Aplicar formato a las celdas del cuerpo y fusionar
    current_service = None
    start_row = 11  # Inicio de la fila de datos
    merge_start_row = start_row  # Fila donde comenzará la fusión

    for row in range(len(indicadores)):
        service = indicadores.iloc[row, 0]
        status = indicadores.iloc[row, 1]

        # Verificar si el servicio ha cambiado
        if service != current_service:
            # Si ha cambiado y no es la primera fila, fusionar las celdas del servicio anterior
            if current_service is not None:
                worksheet.merge_range(merge_start_row, 0, row + start_row - 1, 0, current_service, cell_format)

            # Actualizar el servicio actual
            current_service = service
            merge_start_row = row + start_row  # Actualizar la fila de inicio para la fusión

            # Escribir el nuevo servicio con borde
            worksheet.write(row + start_row, 0, service, cell_format)
            worksheet.write(row + start_row, 1, status, cell_format)
        else:
            # Solo escribir el estatus si el servicio es el mismo
            worksheet.write(row + start_row, 1, status, cell_format)

        # Aplicar formato a las celdas de datos
        for col_num in range(2, len(indicadores.columns)):
            worksheet.write(row + start_row, col_num, indicadores.iloc[row, col_num], cell_format)

    # Fusionar la última serie de celdas del servicio
    if current_service is not None:
        worksheet.merge_range(merge_start_row, 0, len(indicadores) + start_row - 1, 0, current_service, cell_format)

    # Agregar sección de "QUEJAS" con valores en 0
    quejas_start_row = len(indicadores) + start_row
    worksheet.merge_range(quejas_start_row, 0, quejas_start_row + 2, 0, 'QUEJAS', cell_format)
    worksheet.write(quejas_start_row, 1, 'ASIGNACION', cell_format)
    worksheet.write(quejas_start_row + 1, 1, 'CERRADAS', cell_format)
    worksheet.write(quejas_start_row + 2, 1, 'TECNICOS', cell_format)
    for col_num in range(2, len(indicadores.columns)):
        worksheet.write(quejas_start_row, col_num, 0, cell_format)
        worksheet.write(quejas_start_row + 1, col_num, 0, cell_format)
        worksheet.write(quejas_start_row + 2, col_num, 0, cell_format)

    # Ajustar el ancho de las columnas
    worksheet.set_column(0, len(indicadores.columns) - 1, 20)

    # Aplicar bordes a todo el rango de datos
    last_row = quejas_start_row + 2  # Fila donde termina el contenido
    last_col = len(indicadores.columns) - 1  # Última columna (0-indexed)
    worksheet.conditional_format(start_row, 0, last_row, last_col, {
        'type': 'no_blanks',
        'format': workbook.add_format({'border': 1}),
    })

    # Calcular Resumenes
    total_cerradas = indicadores[indicadores['Estatus'] == 'CERRADAS'].iloc[:, 2:].sum().sum()
    total_tecnicos = indicadores[indicadores['Estatus'] == 'TECNICOS'].iloc[:, 2:].sum().sum()
    total_asignaciones = indicadores[indicadores['Estatus'] == 'ASIGNACION'].iloc[:, 2:].sum().sum()
    promedio_tecnicos = total_tecnicos / len(indicadores.columns[2:])
    promedio_ordenes_diarias_por_tecnico = total_asignaciones / total_tecnicos

    # Calcular la última columna utilizada por la tabla principal
    resumen_start_col = last_col + 3  # Dejar un espacio de 2 columnas

    # Escribir la primera tabla de resumen
    resumen_start_row = 10
    worksheet.write(resumen_start_row, resumen_start_col, 'RESUMEN', header_format)
    worksheet.write(resumen_start_row + 1, resumen_start_col, 'ORDENES CERRADAS', cell_format)
    worksheet.write(resumen_start_row + 1, resumen_start_col + 1, total_cerradas, cell_format)
    worksheet.write(resumen_start_row + 2, resumen_start_col, 'TECNICOS PROMEDIO', cell_format)
    worksheet.write(resumen_start_row + 2, resumen_start_col + 1, promedio_tecnicos, cell_format)
    worksheet.write(resumen_start_row + 3, resumen_start_col, 'PROMEDIO DE ORDENES DIARIAS POR TECNICO', cell_format)
    worksheet.write(resumen_start_row + 3, resumen_start_col + 1, promedio_ordenes_diarias_por_tecnico, cell_format)

    # Escribir la segunda tabla de resumen
    resumen2_start_row = resumen_start_row + 6
    worksheet.write(resumen2_start_row, resumen_start_col, 'RESUMEN', header_format)
    worksheet.write(resumen2_start_row, resumen_start_col + 1, 'TOTAL', header_format)
    worksheet.write(resumen2_start_row, resumen_start_col + 2, 'PROMEDIO', header_format)
    worksheet.write(resumen2_start_row + 1, resumen_start_col, 'ASIGNACION', cell_format)
    worksheet.write(resumen2_start_row + 1, resumen_start_col + 1, total_asignaciones, cell_format)
    worksheet.write(resumen2_start_row + 1, resumen_start_col + 2, total_asignaciones / len(indicadores.columns[2:]), cell_format)
    worksheet.write(resumen2_start_row + 2, resumen_start_col, 'CIERRE', cell_format)
    worksheet.write(resumen2_start_row + 2, resumen_start_col + 1, total_cerradas, cell_format)
    worksheet.write(resumen2_start_row + 2, resumen_start_col + 2, total_cerradas / len(indicadores.columns[2:]), cell_format)
    worksheet.write(resumen2_start_row + 3, resumen_start_col, 'TECNICOS', cell_format)
    worksheet.write(resumen2_start_row + 3, resumen_start_col + 1, total_tecnicos, cell_format)
    worksheet.write(resumen2_start_row + 3, resumen_start_col + 2, promedio_tecnicos, cell_format)

    # Ajustar el ancho de las columnas para las tablas de resumen
    worksheet.set_column(resumen_start_col, resumen_start_col + 2, 30)

    # Calcular totales por día para todos los servicios
    total_asignaciones_por_dia = indicadores[indicadores['Estatus'] == 'ASIGNACION'].iloc[:, 2:].sum()
    total_cerradas_por_dia = indicadores[indicadores['Estatus'] == 'CERRADAS'].iloc[:, 2:].sum()
    total_tecnicos_por_dia = indicadores[indicadores['Estatus'] == 'TECNICOS'].iloc[:, 2:].sum()

    # Convertir las fechas a formato 'YYYY-MM-DD'
    total_asignaciones_por_dia.index = pd.to_datetime(total_asignaciones_por_dia.index).strftime('%Y-%m-%d')
    total_cerradas_por_dia.index = pd.to_datetime(total_cerradas_por_dia.index).strftime('%Y-%m-%d')
    total_tecnicos_por_dia.index = pd.to_datetime(total_tecnicos_por_dia.index).strftime('%Y-%m-%d')

    # Escribir la tabla de totales por día
    totales_start_row = last_row + 3  # Dejar un espacio de 3 filas
    worksheet.write(totales_start_row, 0, 'SERVICIO', header_format)
    worksheet.write(totales_start_row, 1, 'ESTATUS', header_format)
    for col_num, dia in enumerate(total_asignaciones_por_dia.index, start=2):
        worksheet.write(totales_start_row, col_num, dia, header_format)

    worksheet.write(totales_start_row + 1, 0, 'TOTAL', cell_format)
    worksheet.write(totales_start_row + 1, 1, 'ASIGNACION', cell_format)
    for col_num, total in enumerate(total_asignaciones_por_dia, start=2):
        worksheet.write(totales_start_row + 1, col_num, total, cell_format)

    worksheet.write(totales_start_row + 2, 1, 'CERRADAS', cell_format)
    for col_num, total in enumerate(total_cerradas_por_dia, start=2):
        worksheet.write(totales_start_row + 2, col_num, total, cell_format)

    worksheet.write(totales_start_row + 3, 1, 'TECNICOS', cell_format)
    for col_num, total in enumerate(total_tecnicos_por_dia, start=2):
        worksheet.write(totales_start_row + 3, col_num, total, cell_format)

    # Ajustar el ancho de las columnas para la tabla de totales
    worksheet.set_column(0, len(total_asignaciones_por_dia) + 1, 20)

    # Crear nuevas tablas con valores en 0 y fechas reconocidas
    nuevas_tablas_start_row = totales_start_row + 7
    worksheet.write(nuevas_tablas_start_row, 0, '08:00 A.M.', header_format)
    worksheet.write(nuevas_tablas_start_row, 1, 'TECNICOS', header_format)
    for col_num, dia in enumerate(total_asignaciones_por_dia.index, start=2):
        worksheet.write(nuevas_tablas_start_row, col_num, dia, header_format)
    worksheet.write(nuevas_tablas_start_row + 1, 0, 'INICIO DE PRIMER ORDEN', cell_format)
    worksheet.write(nuevas_tablas_start_row + 1, 1, 'INICIO', cell_format)
    for col_num in range(2, len(total_asignaciones_por_dia) + 2):
        worksheet.write(nuevas_tablas_start_row + 1, col_num, 0, cell_format)
    worksheet.write(nuevas_tablas_start_row + 2, 1, 'PENDIENTE', cell_format)
    for col_num in range(2, len(total_asignaciones_por_dia) + 2):
        worksheet.write(nuevas_tablas_start_row + 2, col_num, 0, cell_format)
    worksheet.write(nuevas_tablas_start_row + 3, 1, '%', cell_format)
    for col_num in range(2, len(total_asignaciones_por_dia) + 2):
        worksheet.write(nuevas_tablas_start_row + 3, col_num, '0%', cell_format)

    worksheet.write(nuevas_tablas_start_row + 5, 0, '12:00 P.M.', header_format)
    worksheet.write(nuevas_tablas_start_row + 5, 1, 'TECNICOS', header_format)
    for col_num, dia in enumerate(total_asignaciones_por_dia.index, start=2):
        worksheet.write(nuevas_tablas_start_row + 5, col_num, dia, header_format)
    worksheet.write(nuevas_tablas_start_row + 6, 0, 'CIERRE DE PRIMER ORDEN', cell_format)
    worksheet.write(nuevas_tablas_start_row + 6, 1, 'CERRADA', cell_format)
    for col_num in range(2, len(total_asignaciones_por_dia) + 2):
        worksheet.write(nuevas_tablas_start_row + 6, col_num, 0, cell_format)
    worksheet.write(nuevas_tablas_start_row + 7, 1, 'PENDIENTE', cell_format)
    for col_num in range(2, len(total_asignaciones_por_dia) + 2):
        worksheet.write(nuevas_tablas_start_row + 7, col_num, 0, cell_format)
    worksheet.write(nuevas_tablas_start_row + 8, 1, '%', cell_format)
    for col_num in range(2, len(total_asignaciones_por_dia) + 2):
        worksheet.write(nuevas_tablas_start_row + 8, col_num, '0%', cell_format)

    # Mover las gráficas debajo de las nuevas tablas
    chart1 = workbook.add_chart({'type': 'column'})
    chart1.add_series({
        'name': 'Asignación',
        'categories': [sheet_name, totales_start_row, 2, totales_start_row, len(total_asignaciones_por_dia) + 1],
        'values': [sheet_name, totales_start_row + 1, 2, totales_start_row + 1, len(total_asignaciones_por_dia) + 1],
    })
    chart1.add_series({
        'name': 'Cerradas',
        'categories': [sheet_name, totales_start_row, 2, totales_start_row, len(total_cerradas_por_dia) + 1],
        'values': [sheet_name, totales_start_row + 2, 2, totales_start_row + 2, len(total_cerradas_por_dia) + 1],
    })
    chart1.set_title({'name': 'Asignación vs Cerradas'})
    chart1.set_x_axis({'name': 'Día'})
    chart1.set_y_axis({'name': 'Cantidad'})
    worksheet.insert_chart(nuevas_tablas_start_row + 10, 0, chart1, {'x_scale': 1.5, 'y_scale': 1.5})

    chart2 = workbook.add_chart({'type': 'column'})
    chart2.add_series({
        'name': 'Cerradas',
        'categories': [sheet_name, totales_start_row, 2, totales_start_row, len(total_cerradas_por_dia) + 1],
        'values': [sheet_name, totales_start_row + 2, 2, totales_start_row + 2, len(total_cerradas_por_dia) + 1],
    })
    chart2.set_title({'name': 'Cerradas'})
    chart2.set_x_axis({'name': 'Día'})
    chart2.set_y_axis({'name': 'Cantidad'})
    worksheet.insert_chart(nuevas_tablas_start_row + 10, 10, chart2, {'x_scale': 1.5, 'y_scale': 1.5})

def crear_hoja_resumen_semanal(writer, workbook, sheet_name):
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet

    # Títulos personalizados
    title = f"{sheet_name.upper()}"
    worksheet.write('A1', title, workbook.add_format({'bold': True, 'font_size': 14}))
    worksheet.insert_image('A3', image_path, {'x_scale': 1, 'y_scale': 1.5})
    

def crear_hoja_suspendidas(writer, workbook, sheet_name):
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet

    # Títulos personalizados
    title = f"{sheet_name.upper()}"
    worksheet.write('A1', title, workbook.add_format({'bold': True, 'font_size': 14}))
    worksheet.insert_image('A3', image_path, {'x_scale': 1, 'y_scale': 1.5})
    

def crear_hoja_estatus_atencion(writer, workbook, sheet_name):
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet

    # Títulos personalizados
    title = f"{sheet_name.upper()}"
    worksheet.write('A1', title, workbook.add_format({'bold': True, 'font_size': 14}))
    worksheet.insert_image('A3', image_path, {'x_scale': 1, 'y_scale': 1.5})
    

def crear_hoja_reporte_tiempo_instalacion(writer, workbook, sheet_name):
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet

    # Títulos personalizados
    title = f"{sheet_name.upper()}"
    worksheet.write('A1', title, workbook.add_format({'bold': True, 'font_size': 14}))
    worksheet.insert_image('A3', image_path, {'x_scale': 1, 'y_scale': 1.5})
    

def crear_hoja_promedio_instalacion(writer, workbook, sheet_name):
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet

    # Títulos personalizados
    title = f"{sheet_name.upper()}"
    worksheet.write('A1', title, workbook.add_format({'bold': True, 'font_size': 14}))
    # Insertar imagen
    worksheet.insert_image('A3', image_path, {'x_scale': 1, 'y_scale': 1.5})