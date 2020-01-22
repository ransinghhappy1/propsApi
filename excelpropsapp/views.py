from django.shortcuts import render
from geopy import Nominatim
from django.http import HttpResponse,Http404
import xlrd
import os
import xlsxwriter
def addressLongitude(request):
    if request.method =='POST':
        getexcel=request.FILES['file']


        excel_data=xlrd.open_workbook(file_contents=getexcel.read())
        excel_sheet=excel_data.sheet_names()
        required_data=[]
        for sheetname in excel_sheet:
            sh=excel_data.sheet_by_name(sheetname)
            for row in range(sh.nrows):
                row_valaues=sh.row_values(row)
                required_data.append((row_valaues[0]))
        excel_data=xlsxwriter.Workbook('E:\\excelfolder\\address1.xlsx')
        worksheet=excel_data.add_worksheet()
        bold=excel_data.add_format({'bold':1})
        worksheet.write('A1','Address',bold)
        worksheet.write('B1','Latitude',bold)
        worksheet.write('C1','Longitude',bold)
        a=[]
        for address in required_data[1::]:
            locator=Nominatim(user_agent='myGeocoder')
            location=locator.geocode(address)
            a.append([address,location.latitude,location.longitude])
        data=tuple(a)

        row=1
        col=0
        for addr,latitude,longitude in (data):
            worksheet.write_string(row,col,addr)
            worksheet.write_number(row,col+1,latitude)
            worksheet.write_number(row,col+2,longitude)
            row +=1
        excel_data.close()
    return render(request,'excel.html')


import os
from django.http import HttpResponse, Http404

def downloadFile(request):
    file_path = 'E:\excelfolder\ address1.xlsx'
    if os.path.exists(file_path):
        with open(file_path, 'rb') as fh:
            response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
            response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
            return response
    raise Http404

