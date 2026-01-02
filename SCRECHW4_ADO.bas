Attribute VB_Name = "adoSCRECHW4"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsSCRECHW4_PutBuffer(rsado As ADODB.Recordset, rsSCRECHW4 As typeSCRECHW4)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsSCRECHW4_PutBuffer = Null
rsado("SCREC4ETB") = rsSCRECHW4.SCREC4ETB
rsado("SCREC4AGE") = rsSCRECHW4.SCREC4AGE
rsado("SCREC4SER") = rsSCRECHW4.SCREC4SER
rsado("SCREC4SSE") = rsSCRECHW4.SCREC4SSE
rsado("SCREC4NAT") = rsSCRECHW4.SCREC4NAT
rsado("SCREC4DEV") = rsSCRECHW4.SCREC4DEV
rsado("SCREC4KMY") = rsSCRECHW4.SCREC4KMY
rsado("SCREC4CFC") = rsSCRECHW4.SCREC4CFC
rsado("SCREC4MFC") = rsSCRECHW4.SCREC4MFC
rsado("SCREC4MDC") = rsSCRECHW4.SCREC4MDC
    
Exit Function

Error_Handler:

rsSCRECHW4_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoSCRECHW4_AddNew(rsado As ADODB.Recordset, rsSCRECHW4 As typeSCRECHW4)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoSCRECHW4_AddNew = Null
rsado.AddNew
adoSCRECHW4_AddNew = rsSCRECHW4_PutBuffer(rsado, rsSCRECHW4)
rsado.Update

Exit Function

Error_Handler:

adoSCRECHW4_AddNew = Error

End Function


