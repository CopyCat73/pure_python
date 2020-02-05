import xlwt
import requests
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element, SubElement
import re

def construct_xml(input_size,input_offset, before_date, after_date):
    downloadsQuery = Element('downloadsQuery')
    family = SubElement(downloadsQuery, 'family')
    family.text = 'StudentThesis'
    accessTimeBeforeDate = SubElement(downloadsQuery, 'accessTimeBeforeDate')
    accessTimeBeforeDate.text = before_date
    accessTimeAfterDate = SubElement(downloadsQuery, 'accessTimeAfterDate')
    accessTimeAfterDate.text = after_date
    size = SubElement(downloadsQuery, 'size')
    size.text = str(input_size)
    offset = SubElement(downloadsQuery, 'offset')
    offset.text = str(input_offset)
    navigationLink = SubElement(downloadsQuery, 'navigationLink')
    navigationLink.text = 'true'
    return ET.tostring(downloadsQuery, method='xml')

api_key = 'Your api key here'
base_api_url = 'https://pureserver/ws/api/516'
download_url = base_api_url + '/downloads'
detail_url = base_api_url + '/student-theses'
get_details = True  # True: more details from theses record, but slow. False = id, title and count only. Significally faster.
before_date = '2020-02-05T00:00:00.001Z'
after_date = '2015-11-01T00:00:00.001Z'
headers = {'Content-Type': 'application/xml', 'api-key': api_key }
namespaces = {'xsi': 'http://www.w3.org/2001/XMLSchema-instance'}

wb = xlwt.Workbook()
ws = wb.add_sheet('Student thesis download count')
row = 0

number_found = 1
size = 100
step = 0
total_download_count = 0
while number_found:

    step +=1
    count = step * size
    offset = count - size
    xml = construct_xml(size,offset,before_date,after_date)
    response = requests.post(download_url, data=xml, headers=headers)
    tree = ET.fromstring(response.content)
    total_count = tree.findall('.//count',namespaces)[0].text
    downloads = tree.findall('.//items/download',namespaces)

    for download in downloads:
        pureid = download.findall('.//pureId',namespaces)[0].text
        title = download.findall('.//name/text',namespaces)[0].text
        download_count = download.findall('.//downloadCount',namespaces)[0].text
        uuid = download.findall('.//contentRef',namespaces)[0].attrib['uuid']

        ws.write(row,0,pureid)
        ws.write(row,1,title)
        ws.write(row,2,download_count)

        if(get_details):
            detail_response = requests.get(detail_url + "/" + uuid, headers=headers)
            student_thesis = ET.fromstring(detail_response.content)
            faculty = student_thesis.findall('.//managingOrganisationalUnit',namespaces)[0].attrib['externalId']
            award_date_year = student_thesis.findall('.//awardDate/year',namespaces)[0].text
            ws.write(row,3,award_date_year)
            ws.write(row,4,faculty)
        row += 1
        if row % 1000 == 0:
            ws.flush_row_data()

    number_found = len(downloads)
    total_download_count += number_found
    print("fetched "+str(total_download_count)+"/"+total_count)

wb.save('student-thesis-download-count.xls')
