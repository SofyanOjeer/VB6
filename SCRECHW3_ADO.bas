Attribute VB_Name = "adoSCRECHW3"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsSCRECHW3_PutBuffer(rsado As ADODB.Recordset, rsSCRECHW3 As typeSCRECHW3)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsSCRECHW3_PutBuffer = Null
rsado("SCREC3ETB") = rsSCRECHW3.SCREC3ETB
rsado("SCREC3AGE") = rsSCRECHW3.SCREC3AGE
rsado("SCREC3SER") = rsSCRECHW3.SCREC3SER
rsado("SCREC3SSE") = rsSCRECHW3.SCREC3SSE
rsado("SCREC3NAT") = rsSCRECHW3.SCREC3NAT
rsado("SCREC3DEV") = rsSCRECHW3.SCREC3DEV

rsado("SCREC3DOS") = rsSCRECHW3.SCREC3DOS
rsado("SCREC3PRE") = rsSCRECHW3.SCREC3PRE
rsado("SCREC3ECH") = rsSCRECHW3.SCREC3ECH
rsado("SCREC3TYP") = rsSCRECHW3.SCREC3TYP
rsado("SCREC3NCL") = rsSCRECHW3.SCREC3NCL
rsado("SCREC3MTR") = rsSCRECHW3.SCREC3MTR
rsado("SCREC3MON") = rsSCRECHW3.SCREC3MON
rsado("SCREC3CAP") = rsSCRECHW3.SCREC3CAP
rsado("SCREC3TAF") = rsSCRECHW3.SCREC3TAF
rsado("SCREC3MAR") = rsSCRECHW3.SCREC3MAR
rsado("SCREC3NBJ") = rsSCRECHW3.SCREC3NBJ

rsado("SCREC3KMY") = rsSCRECHW3.SCREC3KMY
rsado("SCREC3CFC") = rsSCRECHW3.SCREC3CFC
rsado("SCREC3MFC") = rsSCRECHW3.SCREC3MFC
rsado("SCREC3MDC") = rsSCRECHW3.SCREC3MDC
    
Exit Function

Error_Handler:

rsSCRECHW3_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoSCRECHW3_AddNew(rsado As ADODB.Recordset, rsSCRECHW3 As typeSCRECHW3)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoSCRECHW3_AddNew = Null
rsado.AddNew
adoSCRECHW3_AddNew = rsSCRECHW3_PutBuffer(rsado, rsSCRECHW3)
rsado.Update

Exit Function

Error_Handler:

adoSCRECHW3_AddNew = Error

End Function



