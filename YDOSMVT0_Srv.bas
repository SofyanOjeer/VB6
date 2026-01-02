Attribute VB_Name = "srvYDOSMVT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYDOSMVT0
    DOSMVTOPE   As String
    DOSMVTNUM   As Long
    DOSMVTDEV   As String
    DOSMVTPCI  As String
    DOSMVTCLI   As String
    DOSMVTMTD   As Currency
    DOSMVTEVE   As String
    DOSMVTDTR   As Long
    DOSMVTPIE   As Long
    DOSMVTECR   As Long
    DOSMVTANN   As String
    DOSMVTKDC   As String
End Type
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYDOSMVT0_GetBuffer(rsAdo As ADODB.Recordset, rsYDOSMVT0 As typeYDOSMVT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYDOSMVT0_GetBuffer = Null

rsYDOSMVT0.DOSMVTOPE = rsAdo("DOSMVTOPE")
rsYDOSMVT0.DOSMVTNUM = rsAdo("DOSMVTNUM")
rsYDOSMVT0.DOSMVTDEV = rsAdo("DOSMVTDEV")
rsYDOSMVT0.DOSMVTPCI = rsAdo("DOSMVTPCI")
rsYDOSMVT0.DOSMVTCLI = rsAdo("DOSMVTCLI")

rsYDOSMVT0.DOSMVTMTD = rsAdo("DOSMVTMTD")
rsYDOSMVT0.DOSMVTEVE = rsAdo("DOSMVTEVE")
rsYDOSMVT0.DOSMVTDTR = rsAdo("DOSMVTDTR")
rsYDOSMVT0.DOSMVTPIE = rsAdo("DOSMVTPIE")
rsYDOSMVT0.DOSMVTECR = rsAdo("DOSMVTECR")
rsYDOSMVT0.DOSMVTANN = rsAdo("DOSMVTANN")
rsYDOSMVT0.DOSMVTKDC = rsAdo("DOSMVTKDC")

Exit Function

Error_Handler:

rsYDOSMVT0_GetBuffer = Error

End Function









Public Sub rsYDOSMVT0_Init(lYDOSMVT0 As typeYDOSMVT0)
lYDOSMVT0.DOSMVTPCI = ""
lYDOSMVT0.DOSMVTDEV = ""
lYDOSMVT0.DOSMVTCLI = ""
lYDOSMVT0.DOSMVTMTD = 0
lYDOSMVT0.DOSMVTOPE = ""
lYDOSMVT0.DOSMVTEVE = ""
lYDOSMVT0.DOSMVTDTR = 0
lYDOSMVT0.DOSMVTNUM = 0
lYDOSMVT0.DOSMVTPIE = 0
lYDOSMVT0.DOSMVTECR = 0
lYDOSMVT0.DOSMVTANN = ""
lYDOSMVT0.DOSMVTKDC = ""

End Sub



