import openpyxl,os

def BuscarValor(tablename,value,column):
    """ Esta Función se utiliza para simular el Vlookup en Excel
    """
    book=openpyxl.load_workbook(os.getcwd()+"\\Tables.xlsx",read_only=True)
    sheet=book[tablename]
    Filerows=sheet.rows
    for row in Filerows:
        if row[0].value==value:
            return row[column-1].value

#Valores Permitidos
UnitsFlow=["Kg/h","m3/h","Nm3/h","l/h","lb/h (Spec.)","Sft3/h (Spec.)","SCFH (Spec.)","MMSCFD (Spec.)"]
UnitsPressure=["Bar","psi","kPa","kg/cm2","mmH2O (Spec.)","mmHg (Spec.)"]
UnitsPressure_2=["(a)","(g)"]
UnitsTemperature=["°C","°F","°K (Spec.)"]
UnitsViscosity=["Pa*s","cP","cSt","/  K factor"]
UnitsFlowCoeficient=["Kv","Cv"]
Fuids=["Acetylene","AIR","AMMONIA","ARGON","BENZENE","BUTANE","CARBON DIOXIDE","CARBON MONOXIDE","CHLORINE","DOWTHERM-A","ETHANE","ETHYLENE","FLUORINE","GLYCOL","HELIUM","HYDROGEN","HYDROGEN CHLORIDE","Hydrogen Sulphide","ISOBUTANE","ISOBUTYLENE","METHANE","METHANOL","NATURAL GAS","Neon , Krypton","NITROGEN","Nitrogen (Nitric) Oxide","NITROUS OXIDE","OXYGEN","PHOSGENE","PROPANE","PROPYLENE","STEAM Saturated","STEAM Superheated","Sulphur Dioxide","WATER","Other :" ,"2-Phased Flow :"]
State=["Liquid","Steam Saturated","Steam Superheated","Gas","Vapor","2-Phased Liquid/Gas","2-Phased Liquid/Vapor","2-Phased Gas/Vapor"]

#---------------------------------------------------
#Objeto Fluid
Fluid={
    "Fluid":"PROPYLENE",
    "State":"Steam Saturated"
}
Fluid["Crit Press"]=int(BuscarValor("FluidData",Fluid["Fluid"],5))
#---------------------------------------------------
#Objeto FlowRate
FlowRate={
    "Units":["Sft3/h (Spec.)"],
    "Max Flow":7000,
    "Norm Flow":5000,
    "Min Flow":10,
    "Shut-Off":10
}
#---------------------------------------------------
#Objeto PhaseFlowRate
PhaseFlowRate={
    "Units":["Sft3/h (Spec.)"],
    "Max Flow":7000,
    "Norm Flow":5000,
    "Min Flow":10,
    "Vap/Liq":None
}
#---------------------------------------------------
#Objeto InletPressure
InletPressure={
    "Units":["Bar"],
    "Units_2":["(a)"],
    "Max Flow":7000,
    "Norm Flow":5000,
    "Min Flow":10
}
#---------------------------------------------------
#Objeto OutletPressure
OutletPressure={
    "Units":["Bar"],
    "Units_2":["(a)"],
    "Max Flow":7000,
    "Norm Flow":5000,
    "Min Flow":10  
}
#---------------------------------------------------
#Inlet Temperature
InleTemperature={
    "Units":"°K (Spec.)",
    "Max Flow":7000,
    "Norm Flow":5000,
    "Min Flow":10
}
#---------------------------------------------------
#Inlet Temperature STEAM
InleTemperatureSTEAM={
    "Units":"°K (Spec.)",
    "Max Flow":7000,
    "Norm Flow":5000,
    "Min Flow":10
}
#---------------------------------------------------
#Objeto SpecWt/Mol Wt.
SpectMol={
    "Units":"Kg/m3|Kg/kmol",
    "Max Flow":7000,
    "Min Flow":"Spec Wt (2-ph Vapor)"
}
SpectMol["Norm Flow"]=BuscarValor("FluidData",Fluid["Fluid"],7)
#---------------------------------------------------
##Objeto Viscosity
ViscositySpec={
    "Units":"cP"
}
ViscositySpec["Norm Flow"]=BuscarValor("FluidData",Fluid["Fluid"],9)
#---------------------------------------------------
#Objeto Vapor Presure
VaporPresure={
    "Units":InletPressure["Units"],
    "Units_2":InletPressure["Units_2"],
    "Max Flow":7000
}
#---------------------------------------------------
inputData={
    "AbsoluteVaporPress":0.056,
    "PressureRecoveryFactorF1":0.9,
    "ReynoldsNumberFactor":1,
    "Densityp1":SpectMol["Max Flow"],
    "MAXAbsoluteInlet_P":InletPressure["Max Flow"],
    "MAXAbsoluteOutlet_P":OutletPressure["Max Flow"],
    "NORAbsoluteInlet_P":InletPressure["Norm Flow"],
    "NORAbsoluteOutlet_P":OutletPressure["Norm Flow"],
    "MINAbsoluteIntel_P":InletPressure["Min Flow"],
    "MINAbsluteOutlet_P":OutletPressure["Min Flow"],
    "MAXTemperature_t1":InleTemperature["Max Flow"],
    "NORTemperature_t1":InleTemperature["Norm Flow"],
    "MINTemperature_t1":InleTemperature["Min Flow"]
}
#InputData[AbsluteThermodyn]
if Fluid["Crit Press"]>0:
    inputData["AbsoluteThermodyn"]=Fluid["Crit Press"]
else:
    inputData["AbsoluteThermodyn"]=0

#InputData[MAXFlowRate]
if Fluid["Units"]=="m3/h":
    inputData["MAXFlowRate"]=FlowRate["Max Flow"]
else:
    inputData["MAXFlowRate"]=FlowRate["Max Flow"]/SpectMol["Max Flow"]

#NOR FlowRate
if Fluid["Units"]=="m3/h":
    inputData["NORFlowRate"]=FlowRate["Norm Flow"]
else:
    inputData["NORFlowRate"]=FlowRate["Norm Flow"]/SpectMol["Max Flow"]

#MIN FlowRate
if Fluid["Units"]=="m3/h":
    inputData["MINFlowRate"]=FlowRate["Min Flow"]
else:
    inputData["MINFlowRate"]=FlowRate["Min Flow"]/SpectMol["Max Flow"]
#---------------------------------------------------
#Flow Coefficient
FlowCoefficient={
    "Units":"Kv",
    "Max Flow":"d",
    "Norm Flow":44,
    "Min Flow":55
}
#---------------------------------------------------