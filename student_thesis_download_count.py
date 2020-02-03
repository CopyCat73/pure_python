import xlwt
import requests
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element, SubElement
import re

def construct_xml(input_size,input_offset):
    downloadsQuery = Element('downloadsQuery')
    family = SubElement(downloadsQuery, 'family')
    family.text = 'StudentThesis'
    accessTimeBeforeDate = SubElement(downloadsQuery, 'accessTimeBeforeDate')
    accessTimeBeforeDate.text = '2020-01-01T00:00:00.001Z'
    accessTimeAfterDate = SubElement(downloadsQuery, 'accessTimeAfterDate')
    accessTimeAfterDate.text = '2019-12-31T00:00:00.001Z'
    size = SubElement(downloadsQuery, 'size')
    size.text = str(input_size)
    offset = SubElement(downloadsQuery, 'offset')
    offset.text = str(input_offset)
    navigationLink = SubElement(downloadsQuery, 'navigationLink')
    navigationLink.text = 'true'
    return ET.tostring(downloadsQuery, method='xml')

api_key = 'your key here'
api_url = 'https://pureserver/ws/api/516/downloads'
headers = {'Content-Type': 'application/xml', 'api-key': api_key }
namespaces = {'xsi': 'http://www.w3.org/2001/XMLSchema-instance'}

wb = xlwt.Workbook()
ws = wb.add_sheet('Student thesis count')
row = 0

number_found = 1
size = 100
step = 0
while number_found:

    step +=1
    count = step * size
    offset = count - size
    xml = construct_xml(size,offset)
    response = requests.post(api_url, data=xml, headers=headers)
    tree = ET.fromstring(response.content)
    downloads = tree.findall('.//items/download',namespaces)

    for download in downloads:
        pureid = download.findall('.//pureId',namespaces)[0].text
        title = download.findall('.//name/text',namespaces)[0].text
        download_count = download.findall('.//downloadCount',namespaces)[0].text
        ws.write(row,0,pureid)
        ws.write(row,1,title)
        ws.write(row,2,download_count)
        row += 1

    number_found = len(downloads)
    print("step "+str(step)+" number found "+str(number_found))

wb.save('student-thesis-count.xls')
