Attribute VB_Name = "srvYCOMREF0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const constYCOMREF0 = "YCOMREF0"
Type typeYCOMREF0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    COMREFETA       As Integer                        ' ETABLISSEMENT
    COMREFPLA       As Long                           ' NUMERO PLAN
    COMREFCOM       As String * 20                    ' NUMERO COMPTE
    COMREFCOR       As String * 2                     ' CODE REFERENCE
    COMREFREF       As String * 15                    ' REFERENCE COMPTE

End Type
Public Sub srvYCOMREF0_Init(recYCOMREF0 As typeYCOMREF0)
recYCOMREF0.Obj = "YCOMREF0"
recYCOMREF0.Method = ""
recYCOMREF0.Err = ""
recYCOMREF0.COMREFETA = 0
recYCOMREF0.COMREFPLA = 0
recYCOMREF0.COMREFCOM = ""
recYCOMREF0.COMREFCOR = ""
recYCOMREF0.COMREFREF = ""
End Sub
Public Function srvYCOMREF0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCOMREF0 As typeYCOMREF0)
On Error GoTo Error_Handler
srvYCOMREF0_GetBuffer_ODBC = Null
recYCOMREF0.COMREFETA = rsADO("COMREFETA")
recYCOMREF0.COMREFPLA = rsADO("COMREFPLA")
recYCOMREF0.COMREFCOM = rsADO("COMREFCOM")
recYCOMREF0.COMREFCOR = rsADO("COMREFCOR")
recYCOMREF0.COMREFREF = rsADO("COMREFREF")
Exit Function
Error_Handler:
srvYCOMREF0_GetBuffer_ODBC = Error
End Function
Public Sub srvYCOMREF0_ElpDisplay(recYCOMREF0 As typeYCOMREF0)
frmElpDisplay.fgData.Rows = 6
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMREFETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMREF0.COMREFETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMREFPLA    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMREF0.COMREFPLA
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMREFCOM   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMREF0.COMREFCOM
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMREFCOR    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE REFERENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMREF0.COMREFCOR
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMREFREF   15A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REFERENCE COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCOMREF0.COMREFREF
frmElpDisplay.Show vbModal
End Sub
