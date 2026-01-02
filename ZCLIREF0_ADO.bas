Attribute VB_Name = "adoZCLIREF0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCLIREF0_PutBuffer(rsADO As ADODB.Recordset, rsZCLIREF0 As typeZCLIREF0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCLIREF0_PutBuffer = Null
rsADO("CLIREFETA") = rsZCLIREF0.CLIREFETA
rsADO("CLIREFCLI") = rsZCLIREF0.CLIREFCLI
rsADO("CLIREFCOR") = rsZCLIREF0.CLIREFCOR
rsADO("CLIREFREF") = rsZCLIREF0.CLIREFREF
Exit Function

Error_Handler:

rsZCLIREF0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoZCLIREF0_AddNew(rsADO As ADODB.Recordset, rsZCLIREF0 As typeZCLIREF0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoZCLIREF0_AddNew = Null
rsADO.AddNew
adoZCLIREF0_AddNew = rsZCLIREF0_PutBuffer(rsADO, rsZCLIREF0)
rsADO.Update

Exit Function

Error_Handler:

adoZCLIREF0_AddNew = Error

End Function

