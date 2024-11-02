import pandas as pd
from io import BytesIO
import xlsxwriter

image_path = 'static/images/lgb.jpg'

def procesar_excel_reportes(file):
    try:
        df = pd.read_excel(file, usecols=[
            'Fecha de Asignación', 'Orden', 'Tipo de Orden', 'Dilacion', 'Servicio', 'Comentario de Criterio',
            'Fecha de cierre', 'Criterio', 'tecnico', 'AREA', 'ESTATUS ETA', 'Dilaciòn', 'Dilacion 2',
            'Ejecutable', 'COMENTARIOS SOBRE ESTATUS', 'ZONA', 'SEMANA', 'MES'
        ])
        df = df.dropna(how='all')
        df = df.dropna(axis=1, how='all')
        df = df.where(pd.notnull(df), None)
        df['Fecha de Asignación'] = pd.to_datetime(df['Fecha de Asignación'], errors='coerce')
        df['Fecha de cierre'] = pd.to_datetime(df['Fecha de cierre'], errors='coerce')
        df = df.dropna(subset=['Fecha de Asignación'])
        return df
    except Exception as e:
        raise ValueError(f"Error al procesar el archivo Excel: {str(e)}")

def indicadores_estatus_atencion(df):
    try:
        estatus_counts = df['ESTATUS ETA'].value_counts()
        total = estatus_counts.sum()
        estatus_df = pd.DataFrame({
            'ESTATUS': estatus_counts.index,
            'TOTAL': estatus_counts.values,
            'PORCENTAJE': (estatus_counts.values / total) * 100
        })
        estatus_df.loc[len(estatus_df)] = ['Total general', total, 100.0]
        
        # Ordenar los estatus
        order = ['Completada', 'Pendiente', 'Suspendida', 'Total general']
        estatus_df['ESTATUS'] = pd.Categorical(estatus_df['ESTATUS'], categories=order, ordered=True)
        estatus_df = estatus_df.sort_values('ESTATUS')
        
        return estatus_df
    except Exception as e:
        raise ValueError(f"Error al generar indicadores: {str(e)}")

def indicadores_por_zona(df):
    try:
        zona_counts = df.pivot_table(index='ESTATUS ETA', columns='ZONA', aggfunc='size', fill_value=0)
        zona_counts['Total general'] = zona_counts.sum(axis=1)
        zona_counts.loc['Total general'] = zona_counts.sum()
        
        # Ordenar los estatus
        order = ['Completada', 'Pendiente', 'Suspendida', 'Total general']
        zona_counts = zona_counts.reindex(order)
        
        return zona_counts.reset_index()
    except Exception as e:
        raise ValueError(f"Error al generar indicadores por zona: {str(e)}")

def generar_excel_reportes(indicadores, indicadores_zona):
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Crear la hoja de ESTATUS ATENCION
        crear_hoja_estatus_atencion(writer, workbook, 'ESTATUS ATENCION', indicadores, indicadores_zona, image_path)

    output.seek(0)
    return output.getvalue()

def crear_hoja_estatus_atencion(writer, workbook, sheet_name, indicadores, indicadores_zona, image_path):
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet

    # Títulos personalizados
    title = f"{sheet_name.upper()}"
    worksheet.write('A1', title, workbook.add_format({'bold': True, 'font_size': 14}))

    # Insertar imagen
    worksheet.insert_image('A3', image_path, {'x_scale': 1, 'y_scale': 1.5})

    # Escribir el DataFrame de estatus a la hoja de Excel
    indicadores.to_excel(writer, index=False, sheet_name=sheet_name, startrow=10)

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
        'text_wrap': True
    })

    # Aplicar formato a los encabezados
    for col_num, value in enumerate(indicadores.columns.values):
        worksheet.write(10, col_num, value, header_format)

    # Aplicar formato a las celdas del cuerpo
    for row in range(len(indicadores)):
        for col in range(len(indicadores.columns)):
            worksheet.write(row + 11, col, indicadores.iloc[row, col], cell_format)

    # Ajustar el ancho de las columnas
    worksheet.set_column(0, len(indicadores.columns) - 1, 20)

    # Escribir el DataFrame de zona a la hoja de Excel
    start_row = len(indicadores) + 13
    indicadores_zona.to_excel(writer, index=False, sheet_name=sheet_name, startrow=start_row)

    # Aplicar formato a los encabezados de la segunda tabla
    for col_num, value in enumerate(indicadores_zona.columns.values):
        worksheet.write(start_row, col_num, value, header_format)

    # Aplicar formato a las celdas del cuerpo de la segunda tabla
    for row in range(len(indicadores_zona)):
        for col in range(len(indicadores_zona.columns)):
            worksheet.write(row + start_row + 1, col, indicadores_zona.iloc[row, col], cell_format)

    # Ajustar el ancho de las columnas de la segunda tabla
    worksheet.set_column(0, len(indicadores_zona.columns) - 1, 20)

    # Calcular la posición para la gráfica
    chart_start_row = start_row + len(indicadores_zona) + 3

    # Crear la gráfica circular 3D
    chart = workbook.add_chart({'type': 'pie', 'subtype': '3d'})
    chart.add_series({
        'name': 'Distribución de Estatus',
        'categories': [sheet_name, 11, 0, 13, 0],  # Rango de categorías (ESTATUS)
        'values': [sheet_name, 11, 1, 13, 1],      # Rango de valores (TOTAL)
        'data_labels': {'value': True},            # Mostrar valores en lugar de porcentajes
    })
    chart.set_title({'name': 'Distribución de Estatus'})
    chart.set_style(10)
    worksheet.insert_chart(chart_start_row, 0, chart, {'x_scale': 1.5, 'y_scale': 1.5})