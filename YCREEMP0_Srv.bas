Attribute VB_Name = "srvYCREEMP0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const constYCREEMP0 = "YCREEMP0"
Type typeYCREEMP0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    CREEMPETA       As Integer                        ' ETABLISSEMENT
    CREEMPAGE       As Integer                        ' AGENCE
    CREEMPSER       As String * 2                     ' SERVICE
    CREEMPSSE       As String * 2                     ' SOUS-SERVICE
    CREEMPDOS       As Long                           ' NUMERO DOSSIER
    CREEMPSEQ       As Long                           ' NUMERO SEQUENCE
    CREEMPNCL       As String * 7                     ' N° CLIENT

End Type
Public Sub srvYCREEMP0_Init(recYCREEMP0 As typeYCREEMP0)
recYCREEMP0.Obj = "YCREEMP0"
recYCREEMP0.Method = ""
recYCREEMP0.Err = ""
recYCREEMP0.CREEMPETA = 0
recYCREEMP0.CREEMPAGE = 0
recYCREEMP0.CREEMPSER = ""
recYCREEMP0.CREEMPSSE = ""
recYCREEMP0.CREEMPDOS = 0
recYCREEMP0.CREEMPSEQ = 0
recYCREEMP0.CREEMPNCL = ""
End Sub
Public Function srvYCREEMP0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCREEMP0 As typeYCREEMP0)
On Error GoTo Error_Handler
srvYCREEMP0_GetBuffer_ODBC = Null
recYCREEMP0.CREEMPETA = rsADO("CREEMPETA")
recYCREEMP0.CREEMPAGE = rsADO("CREEMPAGE")
recYCREEMP0.CREEMPSER = rsADO("CREEMPSER")
recYCREEMP0.CREEMPSSE = rsADO("CREEMPSSE")
recYCREEMP0.CREEMPDOS = rsADO("CREEMPDOS")
recYCREEMP0.CREEMPSEQ = rsADO("CREEMPSEQ")
recYCREEMP0.CREEMPNCL = rsADO("CREEMPNCL")
Exit Function
Error_Handler:
srvYCREEMP0_GetBuffer_ODBC = Error
End Function
Public Sub srvYCREEMP0_ElpDisplay(recYCREEMP0 As typeYCREEMP0)
frmElpDisplay.fgData.Rows = 8
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEMPETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEMP0.CREEMPETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEMPAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEMP0.CREEMPAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEMPSER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEMP0.CREEMPSER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEMPSSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEMP0.CREEMPSSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEMPDOS    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEMP0.CREEMPDOS
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEMPSEQ    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO SEQUENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEMP0.CREEMPSEQ
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEMPNCL    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° CLIENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEMP0.CREEMPNCL
frmElpDisplay.Show vbModal
End Sub
