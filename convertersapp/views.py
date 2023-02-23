from django.shortcuts import render
import convertapi
from .actions import *
def Index(request):
    return render (request,'index.html')
def ImageToPDF(request):
    if request.method=='GET':
        return render (request,'image_to_pdf.html')
    elif request.method=='POST':
        return image_to_pdf(request)
def WordToPdf(request):
    if request.method=='GET':
        return render (request,'word_to_pdf.html')
    elif request.method=='POST':
        return word_to_PDF(request)
def PowerPointToPDF(request):
    if request.method=='GET':
        return render (request,'powerpoint_to_pdf.html')
    elif request.method=='POST':
        return ppt_to_pdf(request)
def ExcelToPDF(request):
    if request.method=='GET':
        return render (request,'excel_to_pdf.html')
    elif request.method=='POST':
        return excel_to_pdf(request)