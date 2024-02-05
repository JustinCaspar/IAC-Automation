"""
This script is used to generate the IAC recommendation for Exhuast Heat Compressors
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

## Calculations
# Operating hours
iac.OH = iac.HR * iac.DY * iac.WK

## Savings
# Natural Gas Savings
iac.NGS = round(iac.HP * iac.FR/100 * iac.EC/100 * 0.002544 * iac.EHR/100 * iac.OH)
# Annual Cost Savigns
iac.ACS = round(iac.NGS * iac.NGC)
## Rebate
iac.PB = payback(iac.ACS, iac.IC)

## Format strings
# set to 2 digits accuracy
iac = dollar(['NGC'],iac,2)
# set the rest to integer
varList = ['ACS', 'IC']
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