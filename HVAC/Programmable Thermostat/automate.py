"""
This script is used to generate the IAC recommendation for Install programmable thermostats
"""

import json5, sys, os
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace, docx_blocks
sys.path.append(os.path.join('..', '..'))
from Shared.IAC import *

# Load utility cost
jsonDict = json5.load(open(os.path.join('..', '..', 'Utility.json5')))
# Load database
jsonDict.update(json5.load(open('database.json5')))
# Convert to easydict
iac = EasyDict(jsonDict)

# Constants
C1 = 12000.0 # Conversion constant; 12,000 BTU/hr/ton
C2 = 1000.0 # Conversion constant; kW/W
# Calculations
iac.PD = round(iac.TON * C1 * (iac.LF/100) / (iac.EER * C2))
iac.OHE = iac.CHR * iac.CDY * iac.CWK
iac.OHP = iac.PHR * iac.PDY * iac.PWK
iac.ES = round(iac.PD * iac.OHE * (1 - iac.MCDH / iac.CDH))
iac.NGS = round(iac.NGU * (1 - iac.MHDH / iac.HDH))
if iac.COOL == True:
    iac.ECS = round(iac.ES * iac.EC)
else:
    iac.ECS = 0
if iac.HEAT == True:
    iac.NGCS = round(iac.NGS * iac.NGC)
else:
    iac.NGCS = 0
iac.ACS = iac.ECS + iac.NGCS
# Implementation
iac.MC = iac.PT * iac.NT
iac.LB = round(iac.NT * iac.IT * iac.LR)
iac.IC = iac.MC + iac.LB
iac.PB = payback(iac.ACS, iac.IC)
## Format strings
# set electricity cost to 3 digits accuracy
iac = dollar(['EC'],iac,3)
# set the natural gas and demand to 2 digits accuracy
iac = dollar(['NGC', 'DC'],iac,2)
# set the rest to integer
varList = ['LR', 'PT', 'LB', 'MC', 'IC', 'ECS', 'NGCS', 'ACS']
iac = dollar(varList,iac,0)
# Format all numbers to string with thousand separator
iac = grouping_num(iac)

# Import docx template
doc = Document('template.docx')

# Replacing keys
docx_replace(doc, **iac)

# If both false
if iac.COOL == False and iac.HEAT == False:
    raise Exception("You need to have at least one section enabled")
# If both true
if iac.COOL and iac.HEAT:
    iac.DOUBLE = True
else:
    iac.DOUBLE = False
# Remove table row
summary = doc.tables[0]
if iac.COOL == False:
    summary._tbl.remove(summary.rows[3]._tr)
if iac.HEAT == False:
    summary._tbl.remove(summary.rows[4]._tr)

docx_blocks(doc, COOL = iac.COOL)
docx_blocks(doc, HEAT = iac.HEAT)
docx_blocks(doc, DOUBLE = iac.DOUBLE)

savefile(doc, iac.REC)

# Caveats
caveat("Please change implementation cost references if necessary.")