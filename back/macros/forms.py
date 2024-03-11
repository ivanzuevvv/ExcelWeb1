from django import forms
from .models import PythonScript, ExcelFile

class PythonScriptForm(forms.ModelForm):
    class Meta:
        model = PythonScript
        fields = ['file']

class ExcelFileForm(forms.ModelForm):
    class Meta:
        model = ExcelFile
        fields = ['file']
