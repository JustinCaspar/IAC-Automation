"""
This script is used to generate the IAC recommendation for Reduce Compressor Set Pressure.
"""

import json5, sys, os, math
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace
sys.path.append(os.path.join('..', '..'))
from Shared.IAC import *

# Load utility cost
jsonDict = json5.load(open(os.path.join('..', '..', 'Utility.json5')))
# Load database
jsonDict.update(json5.load(open('database.json5')))
# Convert to easydict
iac = EasyDict(jsonDict)

# Constants
AP = 14.7
k = 1.4

# Calculating Power reduction
iac.POW = 1-(((iac.RCP+AP)/AP)**((k-1)/(k*iac.N))-1)/(((iac.CCP+AP)/AP)**((k-1)/(k*iac.N))-1)
# 1 decimal point percent
iac.POW = round(iac.POW*100,1)

# Power Draw Reduction
iac.PDR = iac.HP * 0.746 * (iac.LF/100) * (iac.POW/100) / (iac.ETA/100)
# round to 2 significant digits
iac.PDR = round(iac.PDR, -int(math.floor(math.log10(abs(iac.PDR))))+2)

# Opearting Hours
iac.OH = iac.HR * iac.DY * iac.WK

# Energy Savings
iac.ES = round(iac.PDR * iac.OH)

# Demand Savings
iac.DS = round(iac.PDR * (iac.CF/100) * 12)

# Cost savings
iac.ECS = round(iac.ES * iac.EC)
iac.DCS = round(iac.DS * iac.DC)
iac.ACS = iac.ECS + iac.DCS

iac.PB  = payback(iac.ACS, iac.IC)

## Format strings
# set electricity cost to 3 digits accuracy
iac = dollar(['EC'],iac,3)
# set demand to 2 digits accuracy
iac = dollar(['DC'],iac,2)
# set the rest to integer
varList = ['IC', 'ACS', 'ECS', 'DCS']
iac = dollar(varList,iac,0)
# Format all numbers to string with thousand separator
iac = grouping_num(iac)

# Import docx template
doc = Document('template.docx')

# Add equations
# Requires double backslash / curly bracket for LaTeX characters
POWEqn = '\\frac{{ \\left( \\frac{{ {0} + 14.7}}{{14.7}} \\right)^{{\\left( \\frac{{1.4-1}}{{1.4\\times {1} }} \\right)}}-1}}{{ \\left( \\frac{{ {2} + 14.7}}{{14.7}} \\right)^{{\\left( \\frac{{1.4-1}}{{1.4\\times {3} }} \\right)}}-1}}' \
    .format(iac.RCP, iac.N, iac.CCP, iac.N)
add_eqn(doc, iac, '${POWEqn}', POWEqn)

# Replacing keys
docx_replace(doc, **iac)

savefile(doc, iac.REC)

# Caveats
caveat("Please manually change the font size of equations to 16.")
caveat("Please change implementation cost references if necessary.")