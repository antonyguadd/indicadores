from rest_framework.response import Response
from rest_framework.decorators import api_view
from django.shortcuts import render
from django.http import HttpResponse
import pandas as pd

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

@api_view(['POST'])
def upload_excel(request):
    if 'file' not in request.FILES:
        return Response({"error": "No file uploaded"}, status=400)

    excel_file = request.FILES['file']

    try:
        # Procesar el archivo Excel directamente desde la memoria
        df = pd.read_excel(excel_file)
        
        # Reemplazar valores NaN con None
        df = df.where(pd.notnull(df), None)
        
        # Mostrar el número total de filas procesadas
        total_rows = df.shape[0]
        
        return Response({
            "message": "Excel file processed successfully.",
            "total_rows": total_rows
        })
    except Exception as e:
        return Response({"error": str(e)}, status=500)