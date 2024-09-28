import pandas as pd
from io import BytesIO
import xlsxwriter

def procesar_excel(file):
    try:
        # Leer solo las columnas necesarias
        df = pd.read_excel(file, usecols=['Fecha de Asignación', 'Area', 'SEMANA'])
        print(df.head())  # Para verificar el contenido del DataFrame
        
        # Limpiar el DataFrame
        df = df.dropna(how='all')  # Eliminar filas completamente vacías
        df = df.dropna(axis=1, how='all')  # Eliminar columnas completamente vacías
        df = df.where(pd.notnull(df), None)  # Reemplazar NaN con None
        
        # Asegurarse de que las fechas estén en el formato correcto
        df['Fecha de Asignación'] = pd.to_datetime(df['Fecha de Asignación'], format='%d/%m/%Y', errors='coerce')
        
        # Filtrar filas con fechas inválidas
        df = df.dropna(subset=['Fecha de Asignación'])
        
        return df
    except Exception as e:
        raise ValueError(f"Error al procesar el archivo Excel: {str(e)}")

def indicadores_por_dia(df):
    # Agrupamos por 'Area' y 'Fecha de Asignación', contando las órdenes
    indicadores = df.groupby(['Area', 'Fecha de Asignación']).size().reset_index(name='Total Órdenes')

    # Aseguramos que las fechas estén en el formato correcto
    indicadores['Fecha'] = indicadores['Fecha de Asignación'].dt.strftime('%d/%m/%Y')

    # Crear la tabla pivote
    pivot_df = indicadores.pivot_table(
        index='Area', 
        columns='Fecha', 
        values='Total Órdenes', 
        fill_value=0
    )

    # Ordenar las columnas por fecha
    pivot_df = pivot_df.reindex(sorted(pivot_df.columns, key=lambda x: pd.to_datetime(x, format='%d/%m/%Y')), axis=1)

    # Añadir columna de totales por fila
    pivot_df['Total general'] = pivot_df.sum(axis=1)

    # Añadir fila de totales por columna
    pivot_df.loc['Total general'] = pivot_df.sum()

    # Resetear el índice para convertirlo en DataFrame
    return pivot_df.reset_index()

def indicadores_por_semana(df):
    # Agrupamos por 'Area' y 'SEMANA', contando las órdenes
    indicadores = df.groupby(['Area', 'SEMANA']).size().reset_index(name='Total Órdenes')

    # Crear la tabla pivote
    pivot_df = indicadores.pivot_table(
        index='Area', 
        columns='SEMANA', 
        values='Total Órdenes', 
        fill_value=0
    )

    # Añadir columna de totales por fila
    pivot_df['Total general'] = pivot_df.sum(axis=1)

    # Añadir fila de totales por columna
    pivot_df.loc['Total general'] = pivot_df.sum()

    # Resetear el índice para convertirlo en DataFrame
    return pivot_df.reset_index()

def indicadores_por_mes(df):
    # Agrupamos por 'Area' y 'Mes', contando las órdenes
    df['Mes'] = df['Fecha de Asignación'].dt.to_period('M')
    indicadores = df.groupby(['Area', 'Mes']).size().reset_index(name='Total Órdenes')

    # Convertir la columna 'Mes' de Period a string
    indicadores['Mes'] = indicadores['Mes'].astype(str)  # Convertir Period a string

    # Crear la tabla pivote
    pivot_df = indicadores.pivot_table(
        index='Area', 
        columns='Mes', 
        values='Total Órdenes', 
        fill_value=0
    )

    # Añadir columna de totales por fila
    pivot_df['Total general'] = pivot_df.sum(axis=1)

    # Añadir fila de totales por columna
    pivot_df.loc['Total general'] = pivot_df.sum()

    # Resetear el índice para convertirlo en DataFrame
    return pivot_df.reset_index()

def indicadores_por_anio(df):
    # Agrupamos por 'Area' y 'Año', contando las órdenes
    df['Año'] = df['Fecha de Asignación'].dt.year
    indicadores = df.groupby(['Area', 'Año']).size().reset_index(name='Total Órdenes')

    # Crear la tabla pivote
    pivot_df = indicadores.pivot_table(
        index='Area', 
        columns='Año', 
        values='Total Órdenes', 
        fill_value=0
    )

    # Añadir fila de totales
    pivot_df.loc['Total'] = pivot_df.sum()

    # Resetear el índice para convertirlo en DataFrame
    return pivot_df.reset_index()

def generar_excel(indicadores):
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        for sheet_name, df in indicadores.items():
            # Escribir el DataFrame a la hoja de Excel
            df.to_excel(writer, index=False, sheet_name=sheet_name)

            worksheet = writer.sheets[sheet_name]

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
                'border': 1
            })

            # Aplicar formato a los encabezados
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # Aplicar formato a las celdas del cuerpo
            for row in range(1, len(df) + 1):
                for col in range(len(df.columns)):
                    worksheet.write(row, col, df.iloc[row - 1, col], cell_format)

            # Ajustar el ancho de las columnas
            worksheet.set_column(0, len(df.columns) - 1, 20)

    output.seek(0)
    return output.getvalue()