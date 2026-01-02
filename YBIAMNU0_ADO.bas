Attribute VB_Name = "adoYBIAMNU0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYBIAMNU0_PutBuffer(rsADO As ADODB.Recordset, rsYBIAMNU0 As typeYBIAMNU0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYBIAMNU0_PutBuffer = Null

rsADO("Src") = rsYBIAMNU0.Src
rsADO("ID") = rsYBIAMNU0.ID
rsADO("Memo") = rsYBIAMNU0.Memo

Exit Function

Error_Handler:

rsYBIAMNU0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoYBIAMNU0_AddNew(rsADO As ADODB.Recordset, rsYBIAMNU0 As typeYBIAMNU0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoYBIAMNU0_AddNew = Null
rsADO.AddNew
adoYBIAMNU0_AddNew = rsYBIAMNU0_PutBuffer(rsADO, rsYBIAMNU0)
rsADO.Update

Exit Function

Error_Handler:

adoYBIAMNU0_AddNew = Error

End Function
