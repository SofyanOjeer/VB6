Attribute VB_Name = "adoZCOMREF0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCOMREF0_PutBuffer(rsADO As ADODB.Recordset, rsZCOMREF0 As typeZCOMREF0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCOMREF0_PutBuffer = Null

rsADO("COMREFETA") = rsZCOMREF0.COMREFETA
rsADO("COMREFPLA") = rsZCOMREF0.COMREFPLA
rsADO("COMREFCOM") = rsZCOMREF0.COMREFCOM
rsADO("COMREFCOR") = rsZCOMREF0.COMREFCOR
rsADO("COMREFREF") = rsZCOMREF0.COMREFREF
Exit Function

Error_Handler:

rsZCOMREF0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZCOMREF0_AddNew(rsADO As ADODB.Recordset, rsZCOMREF0 As typeZCOMREF0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZCOMREF0_AddNew = Null
rsADO.AddNew
adoZCOMREF0_AddNew = rsZCOMREF0_PutBuffer(rsADO, rsZCOMREF0)
rsADO.Update

Exit Function

Error_Handler:

adoZCOMREF0_AddNew = Error

End Function
