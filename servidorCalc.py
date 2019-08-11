import xlwings
from flask import Flask,render_template,request
def Calculations(FluidState,Maxflow,NormFlow,MinFlow):
    book=xlwings.Book("OriginalCalculations.xls")
    sheet=book.sheets["SELECTION"]

    #setting Values
    sheet.range("T10").value=FluidState #Fluid State
    sheet.range("R12").value=Maxflow #Max Flow
    sheet.range("X12").value=NormFlow #Norm Flow
    sheet.range("AD12").value=MinFlow #Norm Flow

    #Results
    MaxFlowResult=sheet.range("R21").value
    NormFlowResult=sheet.range("X21").value
    MinFlowResult=sheet.range("AD21").value

    return{
        "MaxFlow":MaxFlowResult,
        "NormFlow":NormFlowResult,
        "MinFlow":MinFlowResult
    }


State=["Liquid","Steam Saturated","Steam Superheated","Gas","Vapor","2-Phased Liquid/Gas","2-Phased Liquid/Vapor","2-Phased Gas/Vapor"]

#Server
app = Flask(__name__)

@app.route('/',methods=['POST','GET'])
def Inicio():
    if request.method=='GET':
        return render_template('Inicio.html',states=State)
    elif request.method=='POST':
        result=Calculations(request.form["state"],request.form['MaxFluid'],request.form["NormFluid"],request.form['MinFluid'])
        return render_template('Inicio.html',maxfluid=result['MaxFlow'],states=State)