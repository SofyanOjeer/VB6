Attribute VB_Name = "srvYSWIDOS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYSWIDOS0
    SWIDOSSABK   As Long
    SWIDOSSER    As String
    SWIDOSSSE    As String
    SWIDOSOPEC   As String
    SWIDOSOPEN   As Long
    SWIDOSOPEK   As Long
    SWIDOSMTK    As String
    SWIDOSMON    As Currency
    SWIDOSDEV    As String
    SWIDOSDENV   As Long
    SWIDOSRCV    As String
    SWIDOS20     As String
    SWIDOS21     As String
    SWIDOS50PI   As String
    SWIDOS52A    As String
    SWIDOS59PI    As String
    SWIDOS57A    As String
    SWIDOSROUT   As String

End Type
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYSWIDOS0_GetBuffer(rsAdo As ADODB.Recordset, rsYSWIDOS0 As typeYSWIDOS0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYSWIDOS0_GetBuffer = Null

rsYSWIDOS0.SWIDOSSABK = rsAdo("SWIDOSSABK")
rsYSWIDOS0.SWIDOSSER = rsAdo("SWIDOSSER")
rsYSWIDOS0.SWIDOSSSE = rsAdo("SWIDOSSSE")
rsYSWIDOS0.SWIDOSOPEC = rsAdo("SWIDOSOPEC")
rsYSWIDOS0.SWIDOSOPEN = rsAdo("SWIDOSOPEN")
rsYSWIDOS0.SWIDOSOPEK = rsAdo("SWIDOSOPEK")
rsYSWIDOS0.SWIDOSMTK = rsAdo("SWIDOSMTK")
rsYSWIDOS0.SWIDOSMON = rsAdo("SWIDOSMON")
rsYSWIDOS0.SWIDOSDEV = rsAdo("SWIDOSDEV")
rsYSWIDOS0.SWIDOSDENV = rsAdo("SWIDOSDENV")
rsYSWIDOS0.SWIDOSRCV = rsAdo("SWIDOSRCV")
rsYSWIDOS0.SWIDOS20 = rsAdo("SWIDOS20")
rsYSWIDOS0.SWIDOS21 = rsAdo("SWIDOS21")
rsYSWIDOS0.SWIDOS50PI = rsAdo("SWIDOS50PI")
rsYSWIDOS0.SWIDOS52A = rsAdo("SWIDOS52A")
rsYSWIDOS0.SWIDOS59PI = rsAdo("SWIDOS59PI")
rsYSWIDOS0.SWIDOS57A = rsAdo("SWIDOS57A")
rsYSWIDOS0.SWIDOSROUT = rsAdo("SWIDOSROUT")

Exit Function

Error_Handler:

rsYSWIDOS0_GetBuffer = Error

End Function









Public Sub rsYSWIDOS0_Init(lYSWIDOS0 As typeYSWIDOS0)
lYSWIDOS0.SWIDOSSABK = 0
lYSWIDOS0.SWIDOSSER = ""
lYSWIDOS0.SWIDOSSSE = ""
lYSWIDOS0.SWIDOSOPEC = ""
lYSWIDOS0.SWIDOSOPEN = 0
lYSWIDOS0.SWIDOSOPEK = 0
lYSWIDOS0.SWIDOSMTK = ""
lYSWIDOS0.SWIDOSMON = 0
lYSWIDOS0.SWIDOSDEV = ""
lYSWIDOS0.SWIDOSDENV = ""
lYSWIDOS0.SWIDOSRCV = ""
lYSWIDOS0.SWIDOS20 = ""
lYSWIDOS0.SWIDOS21 = ""
lYSWIDOS0.SWIDOS50PI = ""
lYSWIDOS0.SWIDOS52A = ""
lYSWIDOS0.SWIDOS59PI = ""
lYSWIDOS0.SWIDOS57A = ""
lYSWIDOS0.SWIDOSROUT = ""
End Sub


