from django.db import models
from django.contrib.auth.models import AbstractUser
from django.contrib.auth.models import Group, Permission
from django.contrib.auth.models import AbstractUser
from django.db import models


class PythonScript(models.Model):
    script = models.FileField(upload_to='plugin_dir/')

class ExcelFile(models.Model):
    file = models.FileField(upload_to='excel_files/')

class Result(models.Model):
    result = models.FileField(upload_to='results/')


