Attribute VB_Name = "rsSCRECHW3"
'---------------------------------------------------------
Option Explicit
Type typeSCRECHW3
    SCREC3ETB       As Long
    SCREC3AGE       As Long
    SCREC3SER       As String * 2
    SCREC3SSE       As String * 2
    SCREC3NAT       As String * 3
    SCREC3DEV       As String * 3
    
    SCREC3DOS       As Long
    SCREC3PRE       As Long
    SCREC3ECH       As Long
    SCREC3TYP       As String * 2
    SCREC3NCL       As String * 7
    SCREC3MTR       As Currency
    SCREC3MON       As Currency
    SCREC3CAP       As Currency
    SCREC3TAF       As Double
    SCREC3MAR       As Double
    SCREC3NBJ       As Long
    
    SCREC3KMY       As Currency
    SCREC3CFC       As Currency
    SCREC3MFC       As Currency
    SCREC3MDC       As Currency

End Type
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsSCRECHW3_GetBuffer(rsSab As ADODB.Recordset, rsSCRECHW3 As typeSCRECHW3)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsSCRECHW3_GetBuffer = Null

rsSCRECHW3.SCREC3ETB = rsSab("SCREC3ETB")
rsSCRECHW3.SCREC3AGE = rsSab("SCREC3AGE")
rsSCRECHW3.SCREC3SER = rsSab("SCREC3SER")
rsSCRECHW3.SCREC3SSE = rsSab("SCREC3SSE")
rsSCRECHW3.SCREC3NAT = rsSab("SCREC3NAT")
rsSCRECHW3.SCREC3DEV = rsSab("SCREC3DEV")

rsSCRECHW3.SCREC3DOS = rsSab("SCREC3DOS")
rsSCRECHW3.SCREC3PRE = rsSab("SCREC3PRE")
rsSCRECHW3.SCREC3ECH = rsSab("SCREC3ECH")
rsSCRECHW3.SCREC3TYP = rsSab("SCREC3TYP")
rsSCRECHW3.SCREC3NCL = rsSab("SCREC3NCL")
rsSCRECHW3.SCREC3MTR = rsSab("SCREC3MTR")
rsSCRECHW3.SCREC3MON = rsSab("SCREC3MON")
rsSCRECHW3.SCREC3CAP = rsSab("SCREC3CAP")
rsSCRECHW3.SCREC3TAF = rsSab("SCREC3TAF")
rsSCRECHW3.SCREC3MAR = rsSab("SCREC3MAR")
rsSCRECHW3.SCREC3NBJ = rsSab("SCREC3NBJ")

rsSCRECHW3.SCREC3KMY = rsSab("SCREC3KMY")
rsSCRECHW3.SCREC3CFC = rsSab("SCREC3CFC")
rsSCRECHW3.SCREC3MFC = rsSab("SCREC3MFC")
rsSCRECHW3.SCREC3MDC = rsSab("SCREC3MDC")


Exit Function

Error_Handler:

rsSCRECHW3_GetBuffer = Error

End Function



'











