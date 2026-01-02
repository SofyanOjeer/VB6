Attribute VB_Name = "rsSCRECHW4"
'---------------------------------------------------------
Option Explicit
Type typeSCRECHW4
    SCREC4ETB       As Long
    SCREC4AGE       As Long
    SCREC4SER       As String * 2
    SCREC4SSE       As String * 2
    SCREC4NAT       As String * 3
    SCREC4DEV       As String * 3
    SCREC4KMY       As Currency
    SCREC4CFC       As Currency
    SCREC4MFC       As Currency
    SCREC4MDC       As Currency

End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsSCRECHW4_GetBuffer(rsSab As ADODB.Recordset, rsSCRECHW4 As typeSCRECHW4)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsSCRECHW4_GetBuffer = Null

rsSCRECHW4.SCREC4ETB = rsSab("SCREC4ETB")
rsSCRECHW4.SCREC4AGE = rsSab("SCREC4AGE")
rsSCRECHW4.SCREC4SER = rsSab("SCREC4SER")
rsSCRECHW4.SCREC4SSE = rsSab("SCREC4SSE")
rsSCRECHW4.SCREC4NAT = rsSab("SCREC4NAT")
rsSCRECHW4.SCREC4DEV = rsSab("SCREC4DEV")
rsSCRECHW4.SCREC4KMY = rsSab("SCREC4KMY")
rsSCRECHW4.SCREC4CFC = rsSab("SCREC4CFC")
rsSCRECHW4.SCREC4MFC = rsSab("SCREC4MFC")
rsSCRECHW4.SCREC4MDC = rsSab("SCREC4MDC")


Exit Function

Error_Handler:

rsSCRECHW4_GetBuffer = Error

End Function



'










