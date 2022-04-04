from opencage.geocoder import OpenCageGeocode
import time, openpyxl, os
from pprint import pprint

app = OpenCageGeocode(os.environ("API_KEY"))


def seek_location(address_string):
    location = app.geocode(address_string)
    return location

workbook = openpyxl.load_workbook('FLCUSTOM.xlsx')
sheet = workbook.active
sheet['Q1'] = 'Latitude'
sheet['R1'] = 'Longitude'
rows = iter(sheet.iter_rows())
next(rows)
for row in rows:
    street_num = row[0].value
    street_name = row[1].value
    full_address = str(street_num)+' '+ str(street_name)+', Lawrence, KS'
    data = seek_location(full_address)
    pprint(data)
    result = data[0]
    specific = result['geometry']
    lat = specific['lat']
    lon = specific['lng']
    row[16].value = lat
    row[17].value = lon
    pprint(str(lat)+" "+ str(lon))
    time.sleep(1.5)

workbook.save('output.xlsx')
workbook.close()