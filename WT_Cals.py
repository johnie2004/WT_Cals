
import openpyxl
import os

from Bursting import Min_WT_Presssure_Containment
from Collapse import Min_WT_Collaps
from Propgation_Buckling import Min_WT_Propgation_Buckling
from Reeling import Min_WT_Reeling_No_Clad


os.chdir(r'C:\Users\ss7n0564\Documents\My Work\A Pipeline Study and Tender Tool\Python\WT_Cals')

wb = openpyxl.load_workbook('Inputs.xlsx',data_only = True)

sheet = wb['WT_Cals_Inputs']

Min_WT = round(max(Min_WT_Presssure_Containment,Min_WT_Collaps,Min_WT_Propgation_Buckling, Min_WT_Reeling_No_Clad),3)


sheet['K1'].value = 'Min. Required WT for Pressure Containment'

sheet['O1'].value = str(round(Min_WT_Presssure_Containment,3)) + 'mm'

sheet['K2'].value = 'Min. Required WT for Collapse'

sheet['O2'].value = str(round(Min_WT_Collaps,3)) + 'mm'

sheet['K3'].value = 'Min. Required WT for Propgation Buckling'

sheet['O3'].value = str(round(Min_WT_Propgation_Buckling,3)) + 'mm'

sheet['K4'].value = 'Min. Required WT for Reeling'

sheet['O4'].value = str(round(Min_WT_Reeling_No_Clad,3)) + 'mm'

sheet['K5'].value = 'Min. Required WT'

sheet['O5'].value = str(round(Min_WT,3)) + 'mm'


my_red = openpyxl.styles.colors.Color(rgb='00FF0000')
my_fill = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=my_red)

sheet['O5'].fill = my_fill

wb.save('WT_Cals_Results.xlsx')

