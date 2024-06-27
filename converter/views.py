import os
from django.http import HttpResponse
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.parsers import MultiPartParser, FormParser
from .serializers import FileUploadSerializer
from .utils import pdf_to_excel
from tempfile import NamedTemporaryFile

class FileUploadView(APIView):
    parser_classes = (MultiPartParser, FormParser)

    def post(self, request, *args, **kwargs):
        file_serializer = FileUploadSerializer(data=request.data)
        if file_serializer.is_valid():
            pdf_file = request.FILES['file']

            # Save the uploaded PDF file temporarily
            with NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
                temp_pdf_path = temp_pdf.name
                for chunk in pdf_file.chunks():
                    temp_pdf.write(chunk)

            # Convert PDF to Excel
            excel_file = pdf_to_excel(temp_pdf_path)

            # Prepare response with the Excel file
            with open(excel_file, 'rb') as f:
                response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                response['Content-Disposition'] = f'attachment; filename={os.path.basename(excel_file)}'

            # Clean up temporary files
            os.remove(temp_pdf_path)
            os.remove(excel_file)

            return response

        else:
            return Response(file_serializer.errors, status=400)
