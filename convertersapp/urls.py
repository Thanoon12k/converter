
from django.urls import path
from .views import *

urlpatterns = [
    path('', Index),
   path('img_to_pdf', ImageToPDF,name='image-to-pdf'),
   path('word_to_pdf', WordToPdf,name='word-to-pdf'),
   path('excel_to_pdf', ExcelToPDF,name='excel-to-pdf'),
   path('powerpoint_to_pdf', PowerPointToPDF,name='powerpoint-to-pdf'),


]
