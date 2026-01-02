Attribute VB_Name = "adoYBIATAB0"
Option Explicit

'---------------------------------------------------------
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYBIATAB0_PutBuffer(rsado As ADODB.Recordset, rsYBIATAB0 As typeYBIATAB0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYBIATAB0_PutBuffer = Null

rsado("BIATABID") = rsYBIATAB0.BIATABID
rsado("BIATABK1") = rsYBIATAB0.BIATABK1
rsado("BIATABK2") = rsYBIATAB0.BIATABK2
rsado("BIATABTXT") = rsYBIATAB0.BIATABTXT
   
Exit Function

Error_Handler:

rsYBIATAB0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoYBIATAB0_AddNew(rsado As ADODB.Recordset, rsYBIATAB0 As typeYBIATAB0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoYBIATAB0_AddNew = Null
rsado.AddNew
adoYBIATAB0_AddNew = rsYBIATAB0_PutBuffer(rsado, rsYBIATAB0)
rsado.Update

Exit Function

Error_Handler:

adoYBIATAB0_AddNew = Error

End Function


