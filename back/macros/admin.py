from django.contrib import admin
from .models import PythonScript, ExcelFile, Result

@admin.register(PythonScript)
class PythonScriptAdmin(admin.ModelAdmin):
    pass

@admin.register(ExcelFile)
class ExcelFileAdmin(admin.ModelAdmin):
    pass

@admin.register(Result)
class ResultAdmin(admin.ModelAdmin):
    pass
