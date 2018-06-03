import openpyxl
import os

os.chdir(r'C:\Users\ss7n0564\Documents\My Work\A Pipeline Study and Tender Tool\Python\WT_Cals')

wb = openpyxl.load_workbook('Inputs.xlsx',data_only = True)

sheet = wb['WT_Cals_Inputs']

#Constants
Gravity = 9.807
Sys_Test_Cont_Density = 1025 #kg/m3

# Read the inputs sheets and assign them to the variables

#Geometry Inputs
Steel_Pipe_Diameter_Type = sheet['C2'].value
Steel_Pipe_Diameter = sheet['C3'].value
Pipe_Corrosion_Allowance = sheet['C4'].value

#Safety Class Inputs
Contents_Type = sheet['B7'].value
Operation_Zone = sheet['B8'].value

#Material Inputs
Fabrication_Method = sheet['F2'].value
Pipe_Material = sheet['F3'].value
Max_Design_Temp = sheet['F4'].value
Supp_D = sheet['F5'].value
Supp_U = sheet['F6'].value

#Lined/clad Layer Inputs
Any_Inner_Metal_Layer = sheet['F8'].value
Cladded_Or_Lined = sheet['F9'].value
Metal_Layer_Type = sheet['F10'].value

#Load Inputs
Design_Pressure = sheet['I2'].value
Ref_Level = sheet['I3'].value
Max_Contents_Desity = sheet['I4'].value
Water_Depth_For_Collapse_Prop = sheet['I6'].value
Sea_Water_Density = sheet['I7'].value

#Design Checks
Collapse_Check = sheet['B12'].value

# Calcuated the other parameters

if Supp_D =='Yes':
    if Fabrication_Method == "SMLS":
        Tol = 0.1 # 10%
    elif Fabrication_Method == "HFW": # for 15mm < WT < 20mm
        Tol = 0.8 # mm
    elif Fabrication_Method == "MWP": # for 15mm < WT < 20mm
        Tol = 0.8 # mm
else:
    if Fabrication_Method == "SMLS":
        Tol = 0.125 # 12.5% for 10mm < WT < 25mm
    elif Fabrication_Method == "HFW": # for  WT > 15mm
        Tol = 1.0 # mm
    elif Fabrication_Method == "MWP": # for  WT > 15mm
        Tol = 1.0 # mm

if Pipe_Material == 'DNV 450': # Table 7-5
    SMYS = 450 #unit MPa
    SMTS = 535
elif Pipe_Material == 'DNV 415':
    SMYS = 415 #unit MPa
    SMTS = 520
elif Pipe_Material == 'DNV 390':
    SMYS = 390 #unit MPa
    SMTS = 490
elif Pipe_Material == 'DNV 360':
    SMYS = 360 #unit MPa
    SMTS = 460
elif Pipe_Material == 'DNV 22Cr':   # Table 7-11
    SMYS = 450 #unit MPa
    SMTS = 620
elif Pipe_Material == 'DNV 25Cr':
    SMYS = 550 #unit MPa
    SMTS = 750
else:
    print('Error, Material must be defined')

if Pipe_Material == 'DNV 450' or 'DNV 415' or 'DNV 390' or 'DNV 360':
    Derate_YS = 0
    Derate_TS = 0
    if Max_Design_Temp >= 0 and Max_Design_Temp < 50:
        Derate_YS = 0
        Derate_TS = 0
    elif Max_Design_Temp >= 50 and Max_Design_Temp < 100:
        Derate_YS = 0.6 * Max_Design_Temp - 30
        Derate_TS = 0.6 * Max_Design_Temp - 30
    elif Max_Design_Temp >= 100 and Max_Design_Temp < 200:
        Derate_YS = 0.4 * Max_Design_Temp - 10
        Derate_TS = 0.4 * Max_Design_Temp - 10
    else:
        print("No Material Derating is defined")
elif Pipe_Material == 'DNV 22Cr' or 'DNV 25Cr':
    Derate_YS = 0
    Derate_TS = 0
    if Max_Design_Temp >= 0 and Max_Design_Temp < 20:
        Derate_YS = 0
        Derate_TS = 0
    elif Max_Design_Temp >= 20 and Max_Design_Temp < 50:
        Derate_YS = 4/3 * Max_Design_Temp - 80/3
        Derate_TS = 4/3 * Max_Design_Temp - 80/3
    elif Max_Design_Temp >= 50 and Max_Design_Temp < 100:
        Derate_YS =  Max_Design_Temp - 10
        Derate_TS =  Max_Design_Temp - 10
    elif Max_Design_Temp >= 100 and Max_Design_Temp < 150:
        Derate_YS = 0.6 * Max_Design_Temp + 30
        Derate_YS = 0.6 * Max_Design_Temp + 30
    elif Max_Design_Temp >= 150 and Max_Design_Temp < 200:
        Derate_YS = 0.4 * Max_Design_Temp + 60
        Derate_YS = 0.4 * Max_Design_Temp + 60
    else:
        print("No Material Derating is defined")
else:
    print("No Material Derating is defined")



if Supp_U == 'Yes':
    Material_Strength_Factor = 1.0
else:
    Material_Strength_Factor = 0.96

Char_Yield_Stress = (SMYS - Derate_YS)* Material_Strength_Factor

Char_Tensile_Strength = (SMTS - Derate_TS)* Material_Strength_Factor

if sheet['I18'].value == None:
    Incidental_Design_Factor_Ope = 1.1
else:
    Incidental_Design_Factor_Ope = sheet['I18'].value

if sheet['I19'].value == None:
    Incidental_Sys_Test_Factor_Sys_Test = 1.0
else:
    Incidental_Sys_Test_Factor_Sys_Test = sheet['I19'].value

Material_Factor = 1.15 #For SLS/USL/ALS

# Define Safety Class

if Contents_Type == 'Flammable': # Table 2-1 of OS-F101
    if Operation_Zone == 'Zone 2':
        Safety_Class_Ope = 'High'
    elif Operation_Zone == 'Zone 1':
        Safety_Class_Ope = 'Medium'
    else:
        print('No Safety Class is defined')


if Contents_Type == 'Non-flammable':
    if Operation_Zone == 'Zone 2':
        Safety_Class_Ope = 'Medium'
    elif Operation_Zone == 'Zone 1':
        Safety_Class_Ope = 'Low'
    else:
        print('No Safety Class for Operation is defined')

Safety_Class_Sys_Test = 'Low'

if Safety_Class_Ope == 'High':  #Table 5-3 of OS-F101
    Safety_Class_Resistance_Ope = 1.308

if Safety_Class_Ope == 'Medium':
    Safety_Class_Resistance_Ope = 1.138

if Safety_Class_Ope == 'Low':
    Safety_Class_Resistance_Ope = 1.046

if Safety_Class_Sys_Test == 'Low':
    Safety_Class_Resistance_Sys_Test = 1.046
else:
    print('Safety Class Resistance of System Test is not Normal')


if Collapse_Check == 'Yes':
    if sheet['B27'].value == None:
        Max_Pipe_Ovality = 3.0/100
    else:
        Max_Pipe_Ovality = sheet['B27'].value/100
        
    if sheet['B28'].value == None:
        Min_Internal_Pressure = 0
    else:
        Min_Internal_Pressure = sheet['B28'].value

    
    if sheet['F27'].value == None:
        Young_Modulus = 207000 #Unit MPa
    else:
        Young_Modulus = sheet['F27'].value
    
    if sheet['F28'].value == None:
        Poisson_Ratio = 0.3
    else:
        Poisson_Ratio = sheet['F28'].value
    
    if Fabrication_Method == 'SMLS': #Table 5-5 of OS-F101
        Fabrication_Factor = 1.0
    elif Fabrication_Method == 'HFW':
        Fabrication_Factor = 0.93
    elif Fabrication_Method == 'MWP':
        Fabrication_Factor = 0.93
    else:
        print("No factication method is defined")
    
    if Safety_Class_Ope == 'Low':   #Table 5-3 of OS-F101, Safety Classes for Pressure Containment and others are different
        Safety_Class_Resistance_Factor_Others = 1.04
    elif Safety_Class_Ope == 'Medium':
        Safety_Class_Resistance_Factor_Others = 1.14
    elif Safety_Class_Ope == 'High':
        Safety_Class_Resistance_Factor_Others = 1.26
    else:
        print("No Safety Class for Others is defined")
    
    Maximum_External_Pressure = Sea_Water_Density * Gravity * Water_Depth_For_Collapse_Prop * 1.0e-6

from scipy.optimize import fsolve

     
def f_collaps(x): #define the equation for collpase, see eqn 5.10-5.14 of OS-F101  
    pc = (Maximum_External_Pressure - Min_Internal_Pressure)* Material_Factor * Safety_Class_Resistance_Factor_Others
    def pel(x):
        return (2*Young_Modulus*(x/Steel_Pipe_Diameter)**3)/(1-Poisson_Ratio**2)
    def pp(x):
        return Char_Yield_Stress*Fabrication_Factor*2*x/Steel_Pipe_Diameter
    return (pc-pel(x))*(pc**2-pp(x)**2)-pc*pel(x)*pp(x)*Max_Pipe_Ovality*Steel_Pipe_Diameter/x
            

def pel_test(x):
    return (2*Young_Modulus*(x/Steel_Pipe_Diameter)**3)/(1-Poisson_Ratio**2)

x0 = Steel_Pipe_Diameter/20  
#x0 = Steel_Pipe_Diameter/20.0 # assume the guess value is pipe OD/20


Min_WT_Collaps_WO_Tol = round(float(fsolve(f_collaps,x0)),3)

if Fabrication_Method == 'SMLS':
    Min_WT_Collaps = (Min_WT_Collaps_WO_Tol + Pipe_Corrosion_Allowance)/(1.0-Tol)
if Fabrication_Method == 'HFW':
    Min_WT_Collaps = Min_WT_Collaps_WO_Tol + Tol + Pipe_Corrosion_Allowance
if Fabrication_Method == 'MWP':
    Min_WT_Collaps = Min_WT_Collaps_WO_Tol + Tol + Pipe_Corrosion_Allowance



print('Min_WT_Collaps =', Min_WT_Collaps)



        
    