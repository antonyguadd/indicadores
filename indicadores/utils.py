import pandas as pd
from io import BytesIO
import xlsxwriter

def procesar_excel(file):
    try:
        df = pd.read_excel(file, usecols=['Fecha de Asignación', 'Area', 'SEMANA'])
        df = df.dropna(how='all')
        df = df.dropna(axis=1, how='all')
        df = df.where(pd.notnull(df), None)
        df['Fecha de Asignación'] = pd.to_datetime(df['Fecha de Asignación'], format='%d/%m/%Y', errors='coerce')
        df = df.dropna(subset=['Fecha de Asignación'])
        return df
    except Exception as e:
        raise ValueError(f"Error al procesar el archivo Excel: {str(e)}")

def indicadores_por_dia(df):
    indicadores = df.groupby(['Area', 'Fecha de Asignación']).size().reset_index(name='Total Órdenes')
    indicadores['Fecha'] = indicadores['Fecha de Asignación'].dt.strftime('%d/%m/%Y')
    pivot_df = indicadores.pivot_table(index='Area', columns='Fecha', values='Total Órdenes', fill_value=0)
    pivot_df = pivot_df.reindex(sorted(pivot_df.columns, key=lambda x: pd.to_datetime(x, format='%d/%m/%Y')), axis=1)
    pivot_df['Total general'] = pivot_df.sum(axis=1)
    pivot_df.loc['Total general'] = pivot_df.sum()
    return pivot_df.reset_index()

def indicadores_por_semana(df):
    indicadores = df.groupby(['Area', 'SEMANA']).size().reset_index(name='Total Órdenes')
    pivot_df = indicadores.pivot_table(index='Area', columns='SEMANA', values='Total Órdenes', fill_value=0)
    pivot_df['Total general'] = pivot_df.sum(axis=1)
    pivot_df.loc['Total general'] = pivot_df.sum()
    return pivot_df.reset_index()

def indicadores_por_mes(df):
    df['Mes'] = df['Fecha de Asignación'].dt.to_period('M')
    indicadores = df.groupby(['Area', 'Mes']).size().reset_index(name='Total Órdenes')
    indicadores['Mes'] = indicadores['Mes'].astype(str)
    pivot_df = indicadores.pivot_table(index='Area', columns='Mes', values='Total Órdenes', fill_value=0)
    pivot_df['Total general'] = pivot_df.sum(axis=1)
    pivot_df.loc['Total general'] = pivot_df.sum()
    return pivot_df.reset_index()

def indicadores_por_anio(df):
    df['Año'] = df['Fecha de Asignación'].dt.year
    indicadores = df.groupby(['Area', 'Año']).size().reset_index(name='Total Órdenes')
    pivot_df = indicadores.pivot_table(index='Area', columns='Año', values='Total Órdenes', fill_value=0)
    pivot_df.loc['Total'] = pivot_df.sum()
    return pivot_df.reset_index()

def generar_excel(indicadores):
    output = BytesIO()
    image_path = 'static/images/lgb.jpg'
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        for sheet_name, df in indicadores.items():
            worksheet = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = worksheet
            
            # Títulos personalizados
            title = f"ORDENES RECIBIDAS POR TECNOLOGIA por {sheet_name.lower()}"
            worksheet.write('A1', title, workbook.add_format({'bold': True, 'font_size': 14}))
            
            # Insertar imagen
            worksheet.insert_image('A3', image_path, {'x_scale': 1, 'y_scale': 1.5})  # Ajustar escala según sea necesario
            
            # Escribir el DataFrame a la hoja de Excel
            df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=10)
            
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
                worksheet.write(10, col_num, value, header_format)

            # Aplicar formato a las celdas del cuerpo
            for row in range(11, len(df) + 11):
                for col in range(len(df.columns)):
                    worksheet.write(row, col, df.iloc[row - 11, col], cell_format)

            # Ajustar el ancho de las columnas
            worksheet.set_column(0, len(df.columns) - 1, 20)

    output.seek(0)
    return output.getvalue()