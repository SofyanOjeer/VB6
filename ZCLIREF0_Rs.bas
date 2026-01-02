Attribute VB_Name = "rsZCLIREF0"
'---------------------------------------------------------
Option Explicit
Type typeZCLIREF0
    CLIREFETA       As Integer                        ' ETABLISSEMENT
    CLIREFCLI       As String * 7                     ' NUMERO CLIENT
    CLIREFCOR       As String * 2                     ' CODE REFERENCE
    CLIREFREF       As String * 15                    ' REFERENCE CLIENT
    
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCLIREF0_GetBuffer(rsADO As ADODB.Recordset, rsZCLIREF0 As typeZCLIREF0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCLIREF0_GetBuffer = Null

rsZCLIREF0.CLIREFETA = rsADO("CLIREFETA")
rsZCLIREF0.CLIREFCLI = rsADO("CLIREFCLI")
rsZCLIREF0.CLIREFCOR = rsADO("CLIREFCOR")
rsZCLIREF0.CLIREFREF = rsADO("CLIREFREF")
Exit Function

Error_Handler:

rsZCLIREF0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsZCLIREF0_Init(rsZCLIREF0 As typeZCLIREF0)
'---------------------------------------------------------

End Sub


'







