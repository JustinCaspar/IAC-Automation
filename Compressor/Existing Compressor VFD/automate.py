"""
This script is used to generate the IAC recommendation for Installing VFD on Air Compressor
"""

import json5, sys, os
from docx import Document
from easydict import EasyDict
from python_docx_replace import docx_replace, docx_blocks
sys.path.append(os.path.join('..', '..')) 
from Shared.IAC import *
import numpy as np

# Load utility cost
jsonDict = json5.load(open(os.path.join('..', '..', 'Utility.json5')))
# Load database
jsonDict.update(json5.load(open('database.json5')))
# Convert to easydict
iac = EasyDict(jsonDict)

## VFD table
Load = np.linspace(20, 100, num=17)
VFD = np.array([25, 28, 33, 38, 42, 47, 52, 57, 61, 65, 70, 75, 80, 85, 90, 95, 105])

## Calculations
# Operating hours
iac.OH = iac.HR * iac.DY * iac.WK

# Power Fraction without VFD
# Blow Off
if iac.CT == 1:
   iac.FPC = 100
   iac.CT = "blow off"
# Modulation
elif iac.CT == 2:
   iac.FPC = round(0.3 * iac.LF + 70)
   iac.CT = "modulation"
# Load/Unload
elif iac.CT == 3:
   iac.FPC = round(0.5 * iac.LF + 50)
   iac.CT = "load/unload"
else:
   raise Exception("Wrong control type!")
# Power Fraction with VFD
iac.FPV = round(np.interp(iac.LF, Load, VFD).item())
# Current Power Draw
iac.CPD = round((iac.HP * 0.746 * (iac.FPC/100)) / (iac.ETAE/100))
# Proposed Power Draw
iac.PPD = round((iac.HP * 0.746 * (iac.FPV/100)) / (iac.ETAP/100))

## Savings
# Annual Energy Savings
iac.ES = (iac.CPD - iac.PPD) * iac.OH
# Annual Demand Savings
iac.DS = (iac.CPD - iac.PPD) * (iac.CF/100) * 12
# Estimated Cost Savings
iac.ECS = round(iac.ES * iac.EC)
# Demand Cost Savings
iac.DCS = round(iac.DS * iac.DC)
# Total Cost Savings
iac.ACS = iac.ECS + iac.DCS
# Total Installation Cost
if (iac.TANK == True):
  iac.IC = iac.VFD + iac.AIC + iac.ATP
else:
  iac.IC = iac.VFD + iac.AIC

## Rebate
iac.RB = round(iac.RR * iac.ES)
iac.MRB = min(iac.RB, iac.IC/2)
iac.MIC = iac.IC - iac.MRB
iac.PB = payback(iac.ACS, iac.MIC)

## Format strings
# set electricity cost / rebate to 3 digits accuracy
iac = dollar(['EC', 'RR'],iac,3)
# set demand to 2 digits accuracy
iac = dollar(['DC'],iac,2)
# set the rest to integer
varList = ['ACS', 'ECS', 'DCS', 'VFD', 'AIC', 'IC', 'RB', 'MRB', 'MIC', 'ATP']
iac = dollar(varList,iac,0)
# Format all numbers to string with thousand separator
iac = grouping_num(iac)

# Import docx template
doc = Document('template.docx')

# Replacing keys
docx_replace(doc, **iac)

docx_blocks(doc, TANK=iac.TANK)

savefile(doc, iac.REC)

# Caveats
caveat("Please change implementation cost references if necessary.")