import pandas as pd
from io import BytesIO
import xlsxwriter
from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
from rest_framework.decorators import api_view
import logging
import os
from indicadores.utils.utils_reportes import procesar_excel_reportes, indicadores_estatus_atencion, indicadores_por_zona, generar_excel_reportes
logger = logging.getLogger(__name__)

# Directorio temporal para guardar archivos
TEMP_DIR = 'temp_files'
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

def procesar_excel(file):
    required_columns = [
        'Fecha de Asignación', 'Orden', 'Tipo de Orden', 'Dilacion', 'Servicio', 'Comentario de Criterio',
        'Fecha de cierre', 'Criterio', 'tecnico', 'AREA', 'ESTATUS ETA', 'Dilaciòn', 'Dilacion 2', 'Ejecutable',
        'COMENTARIOS SOBRE ESTATUS', 'ZONA', 'SEMANA', 'MES'
    ]
    
    try:
        df = pd.read_excel(file)
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Columnas faltantes en el archivo Excel: {', '.join(missing_columns)}")
        
        df = df[required_columns]
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
        estatus_percentages = (estatus_counts / total * 100).round(2)
        
        indicadores = pd.DataFrame({
            'ESTATUS': estatus_counts.index,
            'TOTAL': estatus_counts.values,
            'PORCENTAJE': estatus_percentages.values
        })
        
        total_row = pd.DataFrame({
            'ESTATUS': ['Total general'],
            'TOTAL': [total],
            'PORCENTAJE': [100.00]
        })
        
        indicadores = pd.concat([indicadores, total_row], ignore_index=True)
        return indicadores
    except Exception as e:
        raise ValueError(f"Error al calcular indicadores: {str(e)}")

def generar_excel(indicadores):
    output = BytesIO()
    image_path = 'static/images/lgb.jpg'
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        crear_hoja_estatus_atencion(writer, workbook, 'ESTATUS ATENCION', indicadores, image_path)
        
    output.seek(0)
    return output.getvalue()

def crear_hoja_estatus_atencion(writer, workbook, sheet_name, indicadores, image_path):
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet

    # Títulos personalizados
    title = f"{sheet_name.upper()}"
    worksheet.write('A1', title, workbook.add_format({'bold': True, 'font_size': 14}))

    # Insertar imagen
    worksheet.insert_image('A3', image_path, {'x_scale': 1, 'y_scale': 1.5})

    # Escribir el DataFrame a la hoja de Excel
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
        'border': 1
    })

    # Aplicar formato a los encabezados
    for col_num, value in enumerate(indicadores.columns.values):
        worksheet.write(10, col_num, value, header_format)

    # Aplicar formato a las celdas del cuerpo
    for row in range(11, len(indicadores) + 11):
        for col in range(len(indicadores.columns)):
            worksheet.write(row, col, indicadores.iloc[row - 11, col], cell_format)

    # Ajustar el ancho de las columnas
    worksheet.set_column(0, len(indicadores.columns) - 1, 20)

@csrf_exempt
@api_view(['POST'])
def upload_excel_report(request):
    if 'file' not in request.FILES:
        return JsonResponse({'error': 'No file provided'}, status=400)
    
    file = request.FILES['file']
    file_path = os.path.join(TEMP_DIR, file.name)
    
    try:
        # Guardar el archivo en el directorio temporal
        with open(file_path, 'wb') as f:
            for chunk in file.chunks():
                f.write(chunk)
        
        # Contar las líneas del archivo Excel
        df = pd.read_excel(file_path)
        line_count = len(df)
        
        return JsonResponse({'message': 'File uploaded successfully', 'file_path': file_path, 'file_name': file.name, 'line_count': line_count})
    except Exception as e:
        logger.error(f"Error uploading file: {str(e)}")
        return JsonResponse({'error': f"Error uploading file: {str(e)}"}, status=500)

@csrf_exempt
@api_view(['POST'])
def upload_excel_report(request):
    if 'file' not in request.FILES:
        return JsonResponse({'error': 'No file provided'}, status=400)
    
    file = request.FILES['file']
    file_path = os.path.join(TEMP_DIR, file.name)
    
    try:
        # Guardar el archivo en el directorio temporal
        with open(file_path, 'wb') as f:
            for chunk in file.chunks():
                f.write(chunk)
        
        # Contar las líneas del archivo Excel
        df = pd.read_excel(file_path)
        line_count = len(df)
        
        return JsonResponse({'message': 'File uploaded successfully', 'file_path': file_path, 'file_name': file.name, 'line_count': line_count})
    except Exception as e:
        logger.error(f"Error uploading file: {str(e)}")
        return JsonResponse({'error': f"Error uploading file: {str(e)}"}, status=500)

@csrf_exempt
@api_view(['POST'])
def upload_and_generate_report_indicators(request):
    file_path = request.POST.get('file_path')
    
    if not file_path or not os.path.exists(file_path):
        return JsonResponse({'error': 'File not found'}, status=400)
    
    try:
        # Procesar el archivo Excel y generar indicadores
        df = procesar_excel_reportes(file_path)
        logger.info(f"Datos procesados: {df.head()}")  # Para verificar el contenido del DataFrame
        
        # Generar los indicadores de estatus de atención
        indicadores = indicadores_estatus_atencion(df)
        
        # Generar los indicadores por zona
        indicadores_zona = indicadores_por_zona(df)
        
        # Generar el archivo Excel con los indicadores
        output_file = generar_excel_reportes(indicadores, indicadores_zona)

        # Guardar el archivo generado temporalmente en el servidor
        generated_file_path = os.path.join(TEMP_DIR, 'indicadores_estatus_atencion.xlsx')
        with open(generated_file_path, 'wb') as f:
            f.write(output_file)

        # URL de descarga del archivo
        download_url = request.build_absolute_uri('/api/download_report/')
        return JsonResponse({'message': 'File processed successfully', 'download_url': download_url})

    except ValueError as e:
        logger.error(f"Error processing file: {str(e)}")
        return JsonResponse({"error": f"Error processing file: {str(e)}"}, status=400)
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        return JsonResponse({"error": f"Unexpected error: {str(e)}"}, status=500)

@csrf_exempt
def download_file_report(request):
    file_path = os.path.join(TEMP_DIR, 'indicadores_estatus_atencion.xlsx')

    try:
        with open(file_path, 'rb') as f:
            response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = 'attachment; filename="indicadores_estatus_atencion.xlsx"'
            return response
    except FileNotFoundError:
        return JsonResponse({'error': 'File not found'}, status=404)