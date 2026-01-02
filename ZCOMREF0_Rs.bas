Attribute VB_Name = "rsZCOMREF0"
'---------------------------------------------------------
Option Explicit
Type typeZCOMREF0
    COMREFETA       As Integer                        ' ETABLISSEMENT
    COMREFPLA       As Long                           ' NUMERO PLAN
    COMREFCOM       As String * 20                    ' NUMERO COMPTE
    COMREFCOR       As String * 2                     ' CODE REFERENCE
    COMREFREF       As String * 15                    ' REFERENCE COMPTE
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCOMREF0_GetBuffer(rsAdo As ADODB.Recordset, rsZCOMREF0 As typeZCOMREF0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCOMREF0_GetBuffer = Null

rsZCOMREF0.COMREFETA = rsAdo("COMREFETA")
rsZCOMREF0.COMREFPLA = rsAdo("COMREFPLA")
rsZCOMREF0.COMREFCOM = rsAdo("COMREFCOM")
rsZCOMREF0.COMREFCOR = rsAdo("COMREFCOR")
rsZCOMREF0.COMREFREF = rsAdo("COMREFREF")

Exit Function

Error_Handler:

rsZCOMREF0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsZCOMREF0_Init(rsZCOMREF0 As typeZCOMREF0)
'---------------------------------------------------------

End Sub


'







