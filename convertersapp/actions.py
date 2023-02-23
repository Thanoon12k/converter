import comtypes.client
from django.core.files.storage import FileSystemStorage
import os
from PIL import Image
from django.http import HttpResponse, Http404,FileResponse
from django.shortcuts import render

                                       
from docx2pdf import convert  #word to pdf
import img2pdf                #image to pdf
import comtypes   #ppt_to_pdf
import win32com.client as win32

def word_to_PDF(request):
    try:
        file = request.FILES['image']
    except:
        return render (request,'word_to_pdf.html',{'errors':'no file uploaded'})


    fs = FileSystemStorage()
    filename = fs.save('temp/word_to_pdf/'+file.name, file)
    word_path = os.path.join(fs.location, filename)
    q=convert(word_path)
    output_path=word_path[:-4] + 'pdf'
    print(word_path+'\n'+output_path+'\n'+str(os.path.exists(word_path)))
    if os.path.exists(output_path):
        with open(output_path, 'rb') as pdf_file:
            response = HttpResponse(pdf_file.read(), content_type='application/pdf')
            response['Content-Disposition'] = 'attachment; filename="output.pdf"'
            return response
    
   
def image_to_pdf(request):
    try:
        file = request.FILES['image']
    except:
        return render (request,'image_to_pdf.html',{'errors':'no file uploaded'})



    fs = FileSystemStorage()
    filename = fs.save('temp/image_to_pdf/'+file.name, file)
    image_path = os.path.join(fs.location, filename)  
    pdf_path = image_path[:-3] + 'pdf'

    image = Image.open(image_path)
    pdf_bytes = img2pdf.convert(image.filename)
    file = open(pdf_path, "wb")
    file.write(pdf_bytes)
    image.close()
    file.close()
    with open(pdf_path, 'rb') as pdf_file:
        response = HttpResponse(pdf_file.read(), content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="output.pdf"'
        return response     
def ppt_to_pdf(request):
    comtypes.CoInitialize()
    fs = FileSystemStorage()

    try:
        file = request.FILES['PowerPoint']
    except:
        return render (request,'powerpoint_to_pdf.html',{'errors':'no file uploaded'})
    
   
    filename = fs.save('temp/ppt_to_pdf/'+file.name, file)
    ppt_path = os.path.join(fs.location, filename)
    pdf_path=ppt_path[:-4] + 'pdf'

    powerpoint = win32.Dispatch('Powerpoint.Application')
    powerpoint.Visible = True
    ppt = powerpoint.Presentations.Open(ppt_path)
    ppt.SaveAs(pdf_path, 32)
    ppt.Close()
    powerpoint.Quit()

    with open(pdf_path, 'rb') as pdf_file:
        response = HttpResponse(pdf_file.read(), content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="output.pdf"'
        return response
    
def excel_to_pdf(request):
    comtypes.CoInitialize()
    fs = FileSystemStorage()

    try:
        file = request.FILES['excel']
    except:
        return render (request,'excel_to_pdf.html',{'errors':'no file uploaded'})
    
   
    filename = fs.save('temp/excel_to_pdf/'+file.name, file)
    excel_path = os.path.join(fs.location, filename)
    pdf_path=excel_path[:-4] + 'pdf'

    excel = win32.Dispatch('Excel.Application')
    workbook = excel.Workbooks.Open(excel_path)
    workbook.ExportAsFixedFormat(0, pdf_path)
    workbook.Close()
    excel.Quit()

    if os.path.exists(pdf_path):
        with open(pdf_path, 'rb') as pdf_file:
            response = HttpResponse(pdf_file.read(), content_type='application/pdf')
            response['Content-Disposition'] = 'attachment; filename="output.pdf"'
            return response
    


