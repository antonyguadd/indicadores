from django.shortcuts import render, redirect
from rest_framework.response import Response
from rest_framework.decorators import api_view
import logging
from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
import pandas as pd
import os
from ..utils.utils_dia import procesar_excel, indicadores_por_dia, indicadores_por_semana, indicadores_por_mes, indicadores_por_anio, generar_excel
from django.contrib.auth.models import User 
from rest_framework import generics
from ..serializers import UserSerializer

logger = logging.getLogger(__name__)

# Directorio temporal para guardar archivos
TEMP_DIR = 'temp_files'
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

# Función para procesar el archivo Excel
def procesar_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip().str.upper()  # Convertir todas las columnas a mayúsculas
        df = df.rename(columns={'AREA': 'Area', 'SEMANA': 'SEMANA', 'FECHA DE ASIGNACIÓN': 'Fecha de Asignación'})  # Renombrar columnas si es necesario
        required_columns = ['SEMANA', 'Area', 'Fecha de Asignación']  # Agregar todas las columnas necesarias
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Columnas faltantes en el archivo Excel: {', '.join(missing_columns)}")
        df = df[required_columns]  # Seleccionar las columnas esperadas
        return df
    except ValueError as e:
        raise ValueError(f"Error al procesar el archivo Excel: {str(e)}")

# Vista para subir un archivo y almacenarlo temporalmente
@csrf_exempt
@api_view(['POST'])
def upload_excel(request):
    if 'file' not in request.FILES:
        return JsonResponse({'error': 'No file provided'}, status=400)
    
    file = request.FILES['file']
    file_path = os.path.join(TEMP_DIR, file.name)
    
    try:
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

# Vista para procesar el archivo almacenado y generar indicadores
@csrf_exempt
@api_view(['POST'])
def upload_and_generate_indicators(request):
    file_path = request.POST.get('file_path')
    
    if not file_path or not os.path.exists(file_path):
        return JsonResponse({'error': 'File not found'}, status=400)
    
    try:
        # Procesar el archivo Excel y generar indicadores
        df = procesar_excel(file_path)
        print(df.head())  # Verificar el contenido del DataFrame
        all_indicators = {
            'Dia': indicadores_por_dia(df),
            'Semana': indicadores_por_semana(df),
            'Mes': indicadores_por_mes(df),
            'Anio': indicadores_por_anio(df),
        }
        
        # Generar el archivo Excel con los indicadores
        output_file = generar_excel(all_indicators)

        # Guardar temporalmente el archivo en el servidor (directorio temporal)
        generated_file_path = os.path.join(TEMP_DIR, 'indicadores_generados.xlsx')
        with open(generated_file_path, 'wb') as f:
            f.write(output_file)

        # Renderizar la plantilla con la URL de descarga
        download_url = request.build_absolute_uri('/api/download/')
        return JsonResponse({'message': 'File processed successfully', 'download_url': download_url})

    except ValueError as e:
        logger.error(f"Error processing file: {str(e)}")
        return JsonResponse({"error": f"Error processing file: {str(e)}"}, status=400)
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        return JsonResponse({"error": f"Unexpected error: {str(e)}"}, status=500)

# Vista para mostrar el botón de descarga en 'download_file.html'
@csrf_exempt
def download_file(request):
    file_path = os.path.join(TEMP_DIR, 'indicadores_generados.xlsx')  # Ruta donde se guarda el archivo

    try:
        with open(file_path, 'rb') as f:
            response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = 'attachment; filename="indicadores_generados.xlsx"'
            return response
    except FileNotFoundError:
        return JsonResponse({'error': 'File not found'}, status=404)


# Vista para crear un nuevo usuario
class UserCreate(generics.CreateAPIView):
    queryset = User.objects.all()
    serializer_class = UserSerializer