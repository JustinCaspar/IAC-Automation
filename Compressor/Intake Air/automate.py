"""
This script is used to generate the IAC recommendation for Air Intake Compressors
"""

import json5, sys, os
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace, docx_blocks
sys.path.append(os.path.join('..', '..')) 
from Shared.IAC import *
import numpy as np
import pgeocode
from datetime import datetime
from meteostat import Point, Monthly, units

# Load utility cost
jsonDict = json5.load(open(os.path.join('..', '..', 'Utility.json5')))
# Load database
jsonDict.update(json5.load(open('database.json5')))
# Convert to easydict
iac = EasyDict(jsonDict)

## Retreiving AVerage Outside Temperature
# Use US zipcodes
nomi = pgeocode.Nominatim('us')
query = nomi.query_postal_code(iac.ZIP)
# Return lat and lon information
lat = query['latitude']
lon = query['longitude']

start = datetime(2018, 10, 1)
end = datetime(2022, 5, 31)
# Use coordinates for weather info
point = Point(lat, lon)
# Get monthly data
data = Monthly(point, start, end)
data.convert(units.imperial)
df = data.fetch()
# Extract average temperature
df = df['tavg']
df = df[df.index.month.isin([10,11,12,1,2,3,4,5])]
iac.TO = round(df.mean())

## Calculations
# Compressor Work Reduction
iac.CWR = round((iac.TI - iac.TO)/(iac.TI + 460) * 100, 2)
# Operating hours
iac.OH = iac.HR * iac.DY * iac.WK
# Power Reduction
iac.PR = round((iac.HP * 0.746 * (iac.FR/100) * (iac.CWR/100)) / (iac.ETA/100), 1)

## Savings
# Electrcity Savings
iac.ES = round(iac.PR * iac.OH)
# Demand Savings
iac.DS = round(iac.PR * (iac.CF/100) * iac.C2)
# Electrcicity Cost Savings
iac.ECS = round(iac.ES * iac.EC) 
# Demand Cost Savings
iac.DCS = round(iac.DS * iac.DC)
# Annual Cost Savigns
iac.ACS = round(iac.ECS + iac.DCS)

## Rebate
iac.PB = payback(iac.ACS, iac.IC)

## Format strings
# set to 2 digits accuracy
iac = dollar(['EC','DC'],iac,2)
# set the rest to integer
varList = ['ACS', 'ECS', 'DCS', 'IC']
iac = dollar(varList,iac,0)
# Format all numbers to string with thousand separator
iac = grouping_num(iac)

# Import docx template
doc = Document('template.docx')

# Replacing keys
docx_replace(doc, **iac)

savefile(doc, iac.REC)

# Caveats
caveat("Please change implementation cost references if necessary.")