from django.http import HttpResponseRedirect, HttpResponse
from django.shortcuts import render
from .models import PythonScript, ExcelFile, Result
import subprocess
import json
from django.http import FileResponse
import os
import subprocess
from django.shortcuts import render
from django.conf import settings
import os
import subprocess
from django.http import FileResponse



def display_python_files(request):
    python_scripts = []
    folder_path = 'back/macros/plugin_dir'

    for file_name in os.listdir(folder_path):
        if file_name.endswith('.py'):
            python_scripts.append(file_name)

    return render(request, 'home.html', {'python_scripts': python_scripts})


def home(request):
    if request.method == 'POST':
        if 'python_script' in request.FILES:
            python_script = request.FILES['python_script']
            PythonScript.objects.create(file=python_script)
        if 'excel_file' in request.FILES:
            excel_file = request.FILES['excel_file']
            ExcelFile.objects.create(file=excel_file)

    python_scripts = PythonScript.objects.all()
    excel_files = ExcelFile.objects.all()
    return render(request, 'home.html', {'plugin_dir': python_scripts, 'excel_files': excel_files})


def run_task(request):
    return HttpResponseRedirect('/')



def upload_file(request):
    files_list_python = []
    files_list_excel = []

    if request.method == 'POST' and request.FILES['file']:
        uploaded_file = request.FILES['file']
        action = request.POST.get('action')

        if action == 'upload_file':
            upload_folder = os.path.join(settings.MEDIA_ROOT, 'uploaded_files')
        elif action == 'upload_file2':
            upload_folder = os.path.join(settings.MEDIA_ROOT2, 'uploaded_files2')
        else:
            return render(request, 'upload_file.html', {'error': 'Invalid action'})

        if not os.path.exists(upload_folder):
            os.makedirs(upload_folder)

        file_path = os.path.join(upload_folder, uploaded_file.name)

        with open(file_path, 'wb') as destination:
            for chunk in uploaded_file.chunks():
                destination.write(chunk)

        # Проверяем, является ли загруженный файл Excel
        if uploaded_file.name.endswith('.xlsx'):
            # Выполняем скрипт Python через subprocess
            subprocess.run(['python', 'script.py', file_path])

            result_file_path = os.path.join(os.getcwd(), 'result.txt')

            # Проверяем, существует ли файл 'result.txt'
            if os.path.exists(result_file_path):
                with open(result_file_path, 'rb') as result_file:
                    response = FileResponse(result_file)
                    response['Content-Disposition'] = f'attachment; filename=result.txt'
                    return response
            else:
                return render(request, 'upload_file.html', {'error': 'Result file not found'})

    files_list_python.clear()
    files_list_python.extend(os.listdir(os.path.join(settings.MEDIA_ROOT, 'uploaded_files')))

    files_list_excel.clear()
    files_list_excel.extend(os.listdir(os.path.join(settings.MEDIA_ROOT2, 'uploaded_files2')))

    return render(request, 'upload_file.html',
    {'files_list_python': files_list_python, 'files_list_excel': files_list_excel})









