from django.urls import path
from . import views
from .views import *


urlpatterns = [

    path('k/', views.index, name='index'),
    path('', login_user, name='login'),
    path('register/', register_user, name='register'),





    # Добавьте дополнительные маршруты здесь, если необходимо
]


