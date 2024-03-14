from django.http import HttpResponseRedirect, HttpResponse, FileResponse
import os
import subprocess
from django.conf import settings
from django.contrib.auth import authenticate, login
from django.contrib.auth import logout, login
from django.http import HttpResponse, HttpResponseNotFound, Http404
from django.shortcuts import render, redirect, get_object_or_404
from .forms import *
from django.contrib.auth.decorators import login_required


def register_user(request):
    if request.method == 'POST':
        form = RegisterUserForm(request.POST)
        if form.is_valid():
            form.save()
            username = form.cleaned_data.get('username')
            raw_password = form.cleaned_data.get('password1')
            user = authenticate(username=username, password=raw_password)
            login(request, user)
            return redirect('login')
    else:
        form = RegisterUserForm()
    return render(request, 'register.html', {'form': form})


def login_user(request):
    if request.method == 'POST':
        print("9320")
        form = LoginUserForm(data=request.POST)
        if form.is_valid():
            username = form.cleaned_data.get('username')
            password = form.cleaned_data.get('password')
            user = authenticate(username=username, password=password)
            if user is not None:
                login(request, user)
                return redirect('index')
    else:
        form = LoginUserForm()
    return render(request, 'login.html', {'form': form})



def upload_file(request):
    if request.method == 'POST' and request.FILES.get('file'):
        uploaded_file = request.FILES['file']
        action = request.POST.get('action')
        if action == 'upload_file':
            upload_folder = os.path.join(settings.MEDIA_ROOT, 'uploaded_files')
            result_folder = os.path.join(settings.MEDIA_ROOT, 'back', 'macros', 'result')
        elif action == 'upload_file2':
            upload_folder = os.path.join(settings.MEDIA_ROOT2, 'uploaded_files2')
            result_folder = os.path.join(settings.MEDIA_ROOT2, 'back', 'macros', 'result')
        elif action == 'upload_file3':
            upload_folder = os.path.join(settings.MEDIA_ROOT3, 'uploaded_files3')
            result_folder = os.path.join(settings.MEDIA_ROOT3, 'back', 'macros', 'result')
            if not os.path.exists(upload_folder):
                os.makedirs(upload_folder)  # Создание папки, если она не существует
        else:
            return render(request, 'upload_file.html', {'error': 'Invalid action'})

        if not os.path.exists(upload_folder):
            os.makedirs(upload_folder)

        file_path = os.path.join(upload_folder, uploaded_file.name)

        with open(file_path, 'wb') as destination:
            for chunk in uploaded_file.chunks():
                destination.write(chunk)

        if uploaded_file.name.endswith('.xlsx'):
            subprocess.run(['python', 'sc.py', file_path])
            result_file_path = os.path.join(result_folder, 'result.txt')

            if os.path.exists(result_file_path):
                with open(result_file_path, 'rb') as result_file:
                    response = FileResponse(result_file)
                    response['Content-Disposition'] = f'attachment; filename=result.txt'
                    return response
            else:
                return render(request, 'upload_file.html', {'error': 'Result file not found'})

    files_list_python = os.listdir(os.path.join(settings.MEDIA_ROOT, 'uploaded_files'))
    files_list_excel = os.listdir(os.path.join(settings.MEDIA_ROOT2, 'uploaded_files2'))
    files_list_result = os.listdir(os.path.join(settings.MEDIA_ROOT3, 'uploaded_files3'))

    return render(request, 'upload_file.html',
                  {'files_list_python': files_list_python, 'files_list_excel': files_list_excel, 'files_list_result': files_list_result})



def index(request):
    print("1")
    if request.method == 'POST':
        plugin_file = request.FILES['plugin_file']
        with open(os.path.join('macros/plugins', plugin_file.name), 'wb+') as destination:
            destination.write(plugin_file.read())
        print('2')
        inParameter = 'macros/user1/input'
        outParameter = 'macros/user1/output'

        if plugin_file.name.endswith(".py"):
            plugin_name = plugin_file.name[:-3]
        else:
            plugin_name = plugin_file.name
        print("3")
        result = subprocess.run(['python3', 'macros/main.py', '-p', plugin_name, '-i', inParameter, '-o', outParameter],
                                stdout=subprocess.PIPE)

        print("4")
        if result.returncode == 0:
            return HttpResponse("Выполнено успешно")
        else:
            return HttpResponse("Не выполнено")

    return render(request, 'index.html')


