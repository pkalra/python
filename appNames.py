'''
Created on Jun 16, 2014

@author: pkalra
'''

#!/usr/bin/env python

import urllib3
import xlrd as xd
import xlsxwriter as xw
import lxml.html

input_file = xd.open_workbook("Appname_Mapping.xlsx")
input_sheet = input_file.sheet_by_name("Appname Map")

http = urllib3.PoolManager()
package_name = list()
app_name = list()
app_category = list()
app_update_date = list()
app_version = list()

for app in input_sheet.col(0)[1:]:  
    r = http.request('GET', "https://play.google.com/store/apps/details?id=" + app.value)
    root = lxml.html.fromstring(r.data)
    title = root.cssselect("div.document-title")
    name = ""
    if len(title) > 0: name = title[0].text_content().strip()
    app_name.append(name)
    
    subtitle = root.cssselect("a.document-subtitle")
    category = ""
    if len(subtitle) > 0: category = subtitle[1].text_content().strip()
    app_category.append(category)
    
    content = root.cssselect("div.content")
    update_date = ""
    if len(content) > 0: update_date = content[0].text_content().strip()
    app_update_date.append(update_date)
    
    content = root.cssselect("div.content")
    version = ""
    if len(content) > 0: version = content[3].text_content().strip()
    app_version.append(version)
    
    package_name.append(app.value)

header = list()
header.append("Package Name")
header.append("Application Name")
header.append("Application category")
header.append("Version Update Date")
header.append("Latest Version") 

output_file = xw.Workbook("App Names & Categories.xlsx")
output_sheet1 = output_file.add_worksheet("Appname Map")  
output_sheet1.write_row("A1", header)  
output_sheet1.write_column('A2', package_name)
output_sheet1.write_column('B2', app_name)
output_sheet1.write_column('C2', app_category)
output_sheet1.write_column('D2', app_update_date)  
output_sheet1.write_column('E2', app_version)    

output_file.close()   

