from django.shortcuts import render, redirect
from rest_framework.response import Response
from rest_framework.decorators import api_view
import logging
from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
import pandas as pd
import os
from .utils import procesar_excel, indicadores_por_dia, indicadores_por_semana, indicadores_por_mes, indicadores_por_anio, generar_excel

logger = logging.getLogger(__name__)

# Directorio temporal para guardar archivos
TEMP_DIR = 'temp_files'
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

# Vista para subir un archivo y generar indicadores
@csrf_exempt
@api_view(['POST'])
def upload_and_generate_indicators(request):
    logger.info("Received file upload request.")
    
    if 'file' not in request.FILES:
        return JsonResponse({'error': 'No file provided'}, status=400)
    
    file = request.FILES['file']
    
    try:
        # Procesar el archivo Excel y generar indicadores
        df = procesar_excel(file)
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
        file_path = os.path.join(TEMP_DIR, 'indicadores_generados.xlsx')
        with open(file_path, 'wb') as f:
            f.write(output_file)

        # Renderizar la plantilla con la URL de descarga
        download_url = request.build_absolute_uri('/api/download/')
        return render(request, 'download_file.html', {'download_url': download_url})

    except Exception as e:
        logger.error(f"Error processing file: {str(e)}")
        return JsonResponse({"error": f"Error processing file: {str(e)}"}, status=500)

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

# Vista para cargar el formulario de subida
def upload_form(request):
    if request.method == 'POST':
        if 'file' not in request.FILES:
            return HttpResponse("No file uploaded")
        excel_file = request.FILES['file']
        try:
            df = pd.read_excel(excel_file)
            data_sample = df.head().to_dict()
            return HttpResponse(f"Excel file processed successfully. Data: {data_sample}")
        except Exception as e:
            return HttpResponse(f"Error processing file: {str(e)}")
    return render(request, 'upload_form.html')

# API para subir el archivo Excel
@api_view(['POST'])
@csrf_exempt
def upload_excel(request):
    if 'file' not in request.FILES:
        return Response({"error": "No file uploaded"}, status=400)

    excel_file = request.FILES['file']

    try:
        df = pd.read_excel(excel_file)
        df = df.where(pd.notnull(df), None)  # Reemplazar NaN con None
        total_rows = df.shape[0]
        return Response({
            "message": "Excel file processed successfully.",
            "total_rows": total_rows
        })
    except Exception as e:
        return Response({"error": str(e)}, status=500)

# API para generar indicadores
@csrf_exempt
@api_view(['POST'])
def generar_indicadores(request):
    if request.method == 'POST':
        if 'file' not in request.FILES:
            return JsonResponse({"error": "No file uploaded"}, status=400)

        file = request.FILES['file']
        
        try:
            df = procesar_excel(file)
        except Exception as e:
            return JsonResponse({"error": f"Error processing Excel file: {str(e)}"}, status=500)

        # Obtener periodo y tipo_reporte del cuerpo de la solicitud
        periodo = request.POST.get('periodo')
        tipo_reporte = request.POST.get('tipo_reporte')

        if not periodo or not tipo_reporte:
            return JsonResponse({'error': 'Periodo y tipo_reporte son requeridos.'}, status=400)

        # Generar indicadores según el tipo de reporte
        try:
            if tipo_reporte == 'diario':
                indicadores = indicadores_por_dia(df)
            elif tipo_reporte == 'semanal':
                indicadores = indicadores_por_semana(df)
            elif tipo_reporte == 'mensual':
                indicadores = indicadores_por_mes(df)
            elif tipo_reporte == 'anual':
                indicadores = indicadores_por_anio(df)
            else:
                return JsonResponse({'error': 'Tipo de reporte inválido'}, status=400)
        except Exception as e:
            return JsonResponse({'error': f"Error generating indicators: {str(e)}"}, status=500)

        # Generar el archivo Excel con todos los indicadores
        try:
            output_file = generar_excel({
                'Diario': indicadores_por_dia(df),
                'Semanal': indicadores_por_semana(df),
                'Mensual': indicadores_por_mes(df),
                'Anual': indicadores_por_anio(df)
            })
        except Exception as e:
            return JsonResponse({'error': f"Error generating Excel file: {str(e)}"}, status=500)
        
        # Crear respuesta para descargar el archivo
        response = HttpResponse(output_file, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename="indicadores_generados.xlsx"'
        return response

    return JsonResponse({'error': 'Método no permitido. Usa POST.'}, status=405)