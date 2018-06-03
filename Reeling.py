import openpyxl
import os

os.chdir(r'C:\Users\ss7n0564\Documents\My Work\A Pipeline Study and Tender Tool\Python\WT_Cals')

wb = openpyxl.load_workbook('Inputs.xlsx',data_only = True)

sheet = wb['WT_Cals_Inputs']


# Read the inputs sheets and assign them to the variables

#Geometry Inputs
Steel_Pipe_Diameter_Type = sheet['C2'].value
Steel_Pipe_Diameter = sheet['C3'].value

#Material Inputs
Fabrication_Method = sheet['F2'].value
Pipe_Material = sheet['F3'].value

#Vessel Information
Vessel = sheet['B15'].value

#Lined/clad Layer Inputs
Any_Inner_Metal_Layer = sheet['F8'].value
Cladded_Or_Lined = sheet['F9'].value
Metal_Layer_Type = sheet['F10'].value

if sheet['F38'].value == None:
    if Fabrication_Method == 'HFW':
        Max_YT_Ratio = 0.95
    elif Fabrication_Method == 'SMLS':
        Max_YT_Ratio = 0.89
    else:
        Max_YT_Ratio = 0.90
else:
    Max_YT_Ratio = sheet['F38'].value

if Vessel == '7Oceans':
    Reel_Diameter = 18000 #Unit mm
elif Vessel == '7Navica':
    Reel_Diameter = 15000 #Unit mm
else: 
    print ('Erro, no installation vessel is found')

Resistance_Strain_Factor = 2.0   #Low for installation, Table 7-5 of OS-F101

Load_Effect_Factor = 1.2 # Table 4-4 of OS-F101

if sheet['B37'].value == None:
    Ext_Coating_Thickness = 3  # Unit mm, Default is 3mm 3LPP
else:
    Ext_Coating_Thickness = sheet['B37'].value/1000


#Calcuations for Min Required WT for Reeling

if Fabrication_Method == 'SMLS':
    Condition_Effect_Factor = 0.77
elif Fabrication_Method == 'HFW':
    Condition_Effect_Factor = (0.59+(Max_YT_Ratio-0.9)/0.3)
else:
    Condition_Effect_Factor = 0.82

Max_Nominal_Reeling_Strain = (Steel_Pipe_Diameter/2)/((Reel_Diameter+Steel_Pipe_Diameter)/2 + Ext_Coating_Thickness)

Design_Comp_Strain = Max_Nominal_Reeling_Strain*Load_Effect_Factor*Condition_Effect_Factor

from scipy.optimize import fsolve

x0 = Steel_Pipe_Diameter/20

def f_reeling(t): 
    if Steel_Pipe_Diameter/t <= 20:
        Girth_Weld_Factor = 1.0
    elif Steel_Pipe_Diameter/t <= 45 and Steel_Pipe_Diameter/t > 20:
        Girth_Weld_Factor = 1 - (Steel_Pipe_Diameter/t - 20)/100
    else:
        print('Girth Weld Factor is not defined due to the D/t is larger then 45')
    return Design_Comp_Strain - (0.78*(t/Steel_Pipe_Diameter-0.01)*Max_YT_Ratio**(-1.5)*Girth_Weld_Factor)/Resistance_Strain_Factor


Min_WT_Reeling = round(float(fsolve(f_reeling,x0)),3)

if sheet['F39'].value == None:
    Clad_Layer_Thickness = 3
else:
    Clad_Layer_Thickness = sheet['F39'].value



if Any_Inner_Metal_Layer == 'Yes':
    if Cladded_Or_Lined == 'Cladded':
        Min_WT_Reeling_No_Clad  = Min_WT_Reeling - Clad_Layer_Thickness
    else: 
        Min_WT_Reeling_No_Clad = Min_WT_Reeling
else:
    Min_WT_Reeling_No_Clad = Min_WT_Reeling

print(Min_WT_Reeling_No_Clad)



