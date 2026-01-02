Attribute VB_Name = "srvYDOSSLD0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYDOSSLD0
    DOSSLDOPE   As String
    DOSSLDNUM   As Long
    DOSSLDDEV   As String
    DOSSLDPCI  As String
    DOSSLDCLI   As String
    DOSSLDMDB   As Currency
    DOSSLDMCR   As Currency
    DOSSLDMSD   As Currency
    DOSSLDGDB   As Currency
    DOSSLDGCR   As Currency
    DOSSLDGSD   As Currency
    DOSSLDSTA   As String
    DOSSLDSVC   As String
End Type
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYDOSSLD0_GetBuffer(rsAdo As ADODB.Recordset, rsYDOSSLD0 As typeYDOSSLD0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYDOSSLD0_GetBuffer = Null

rsYDOSSLD0.DOSSLDOPE = rsAdo("DOSSLDOPE")
rsYDOSSLD0.DOSSLDNUM = rsAdo("DOSSLDNUM")
rsYDOSSLD0.DOSSLDDEV = rsAdo("DOSSLDDEV")
rsYDOSSLD0.DOSSLDPCI = rsAdo("DOSSLDPCI")
rsYDOSSLD0.DOSSLDCLI = rsAdo("DOSSLDCLI")

rsYDOSSLD0.DOSSLDMDB = rsAdo("DOSSLDMDB")
rsYDOSSLD0.DOSSLDMCR = rsAdo("DOSSLDMCR")
rsYDOSSLD0.DOSSLDMSD = rsAdo("DOSSLDMSD")
rsYDOSSLD0.DOSSLDGDB = rsAdo("DOSSLDGDB")
rsYDOSSLD0.DOSSLDGCR = rsAdo("DOSSLDGCR")
rsYDOSSLD0.DOSSLDGSD = rsAdo("DOSSLDGSD")
rsYDOSSLD0.DOSSLDSTA = rsAdo("DOSSLDSTA")
rsYDOSSLD0.DOSSLDSVC = rsAdo("DOSSLDSVC")

Exit Function

Error_Handler:

rsYDOSSLD0_GetBuffer = Error

End Function









Public Sub rsYDOSSLD0_Init(lYDOSSLD0 As typeYDOSSLD0)
lYDOSSLD0.DOSSLDPCI = ""
lYDOSSLD0.DOSSLDDEV = ""
lYDOSSLD0.DOSSLDCLI = ""
lYDOSSLD0.DOSSLDMDB = 0
lYDOSSLD0.DOSSLDOPE = ""
lYDOSSLD0.DOSSLDGCR = 0
lYDOSSLD0.DOSSLDMCR = 0
lYDOSSLD0.DOSSLDMSD = 0
lYDOSSLD0.DOSSLDGSD = 0
lYDOSSLD0.DOSSLDGDB = 0
lYDOSSLD0.DOSSLDNUM = 0
lYDOSSLD0.DOSSLDSTA = ""
lYDOSSLD0.DOSSLDSVC = ""

End Sub


