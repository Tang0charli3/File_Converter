import os
from django.http import HttpResponse
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.parsers import MultiPartParser, FormParser
from .serializers import FileUploadSerializer
from .utils import pdf_to_excel, docx_to_excel
from tempfile import NamedTemporaryFile

class FileUploadView(APIView):
    parser_classes = (MultiPartParser, FormParser)

    def post(self, request, *args, **kwargs):
        file_serializer = FileUploadSerializer(data=request.data)
        if file_serializer.is_valid():
            uploaded_file = request.FILES['file']
            file_type = request.data.get('file_type')

            # Save the uploaded file temporarily
            with NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as temp_file:
                temp_file_path = temp_file.name
                for chunk in uploaded_file.chunks():
                    temp_file.write(chunk)

            if file_type == 'pdf':
                # Convert PDF to Excel
                excel_file = pdf_to_excel(temp_file_path)
            elif file_type == 'docx':
                # Convert DOCX to Excel
                excel_file = docx_to_excel(temp_file_path)
            elif file_type == 'docs':
                # Convert DOCS to Excel
                excel_file = docx_to_excel(temp_file_path)
            else:
                return Response({"error": "Unsupported file type."}, status=400)

            # Prepare response with the Excel file
            with open(excel_file, 'rb') as f:
                response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                response['Content-Disposition'] = f'attachment; filename={os.path.basename(excel_file)}'

            # Clean up temporary files
            os.remove(temp_file_path)
            os.remove(excel_file)

            return response

        else:
            return Response(file_serializer.errors, status=400)
