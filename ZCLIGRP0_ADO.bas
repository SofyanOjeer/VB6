Attribute VB_Name = "adoZXXXXXX0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZXXXXXX0_PutBuffer(rsADO As ADODB.Recordset, rsZXXXXXX0 As typeZXXXXXX0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZXXXXXX0_PutBuffer = Null

    
Exit Function

Error_Handler:

rsZXXXXXX0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZXXXXXX0_AddNew(rsADO As ADODB.Recordset, rsZXXXXXX0 As typeZXXXXXX0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZXXXXXX0_AddNew = Null
rsADO.AddNew
adoZXXXXXX0_AddNew = rsZXXXXXX0_PutBuffer(rsADO, rsZXXXXXX0)
rsADO.Update

Exit Function

Error_Handler:

adoZXXXXXX0_AddNew = Error

End Function
