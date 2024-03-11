from django.urls import path
from . import views


urlpatterns = [
    path('', views.home, name='home'),
    path('run_task/', views.run_task, name='run_task'),
    path('upload/', views.upload_file, name='upload_file'),

    # Добавьте дополнительные маршруты здесь, если необходимо
]


