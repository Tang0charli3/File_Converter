import os
import tempfile
from django.http import HttpResponse
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.parsers import MultiPartParser, FormParser
from .serializers import FileUploadSerializer
from .utils import pdf_to_excel, docx_to_excel, ppt_to_excel, excel_to_pdf

class FileUploadView(APIView):
    parser_classes = (MultiPartParser, FormParser)

    def post(self, request, *args, **kwargs):
        file_serializer = FileUploadSerializer(data=request.data)
        if file_serializer.is_valid():
            uploaded_file = request.FILES['file']
            file_type = request.data.get('file_type')

            # Save the uploaded file in a temporary directory with its original name
            temp_dir = tempfile.gettempdir()
            save_path = os.path.join(temp_dir, uploaded_file.name)

            with open(save_path, 'wb') as f:
                for chunk in uploaded_file.chunks():
                    f.write(chunk)

            # Use switch-case for file type handling
            match file_type:
                case 'pdf to excel':
                    output_file = pdf_to_excel(save_path)
                    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                case 'docx to excel':
                    output_file = docx_to_excel(save_path)
                    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                case 'docs to excel':
                    output_file = docx_to_excel(save_path)
                    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                case 'ppt to excel':
                    output_file = ppt_to_excel(save_path)
                    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                case 'excel to pdf':
                    output_file = excel_to_pdf(save_path, os.path.splitext(save_path)[0] + '.pdf')
                    content_type = 'application/pdf'
                case _:
                    return Response({"error": "Unsupported file type."}, status=400)

            # Prepare response with the output file
            with open(output_file, 'rb') as f:
                response = HttpResponse(f.read(), content_type=content_type)
                response['Content-Disposition'] = f'attachment; filename={os.path.basename(output_file)}'

            # Clean up the uploaded file and converted file
            os.remove(save_path)
            os.remove(output_file)

            return response
        else:
            return Response(file_serializer.errors, status=400)
