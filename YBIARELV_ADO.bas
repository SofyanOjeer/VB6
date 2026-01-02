Attribute VB_Name = "adoYBIARELV"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYBIARELV_PutBuffer(rsado As ADODB.Recordset, rsYBIARELV As typeYBIARELV)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYBIARELV_PutBuffer = Null

rsado("BIARELCOM") = rsYBIARELV.BIARELCOM
rsado("BIARELREL") = rsYBIARELV.BIARELREL
rsado("BIARELID") = rsYBIARELV.BIARELID
rsado("BIARELNUM") = rsYBIARELV.BIARELNUM
rsado("BIARELSD0") = rsYBIARELV.BIARELSD0
rsado("BIARELD0") = rsYBIARELV.BIARELD0
rsado("BIARELSD1") = rsYBIARELV.BIARELSD1
rsado("BIARELD1") = rsYBIARELV.BIARELD1
rsado("BIAOLDCOM") = rsYBIARELV.BIAOLDCOM
rsado("BIAOLDDEV") = rsYBIARELV.BIAOLDDEV

    
Exit Function

Error_Handler:

rsYBIARELV_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoYBIARELV_AddNew(rsado As ADODB.Recordset, rsYBIARELV As typeYBIARELV)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoYBIARELV_AddNew = Null
rsado.AddNew
adoYBIARELV_AddNew = rsYBIARELV_PutBuffer(rsado, rsYBIARELV)
rsado.Update

Exit Function

Error_Handler:

adoYBIARELV_AddNew = Error

End Function

