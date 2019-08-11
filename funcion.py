import xlwings

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

print(Calculations("Steam Saturated",7000,5000,300))