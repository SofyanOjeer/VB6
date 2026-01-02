Attribute VB_Name = "srvYDOSSLD1"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYDOSSLD1
    DOSSLDDEV   As String
    DOSSLDPCI  As String
    DOSSLDCLI   As String
    DOSSLDMDB   As Currency
    DOSSLDMCR   As Currency
    DOSSLDMSD   As Currency
    DOSSLDGDB   As Currency
    DOSSLDGCR   As Currency
    DOSSLDGSD   As Currency
End Type
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYDOSSLD1_GetBuffer(rsAdo As ADODB.Recordset, rsYDOSSLD1 As typeYDOSSLD1)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYDOSSLD1_GetBuffer = Null

rsYDOSSLD1.DOSSLDDEV = rsAdo("DOSSLDDEV")
rsYDOSSLD1.DOSSLDPCI = rsAdo("DOSSLDPCI")
rsYDOSSLD1.DOSSLDCLI = rsAdo("DOSSLDCLI")

rsYDOSSLD1.DOSSLDMDB = rsAdo("DOSSLDMDB")
rsYDOSSLD1.DOSSLDMCR = rsAdo("DOSSLDMCR")
rsYDOSSLD1.DOSSLDMSD = rsAdo("DOSSLDMSD")
rsYDOSSLD1.DOSSLDGDB = rsAdo("DOSSLDGDB")
rsYDOSSLD1.DOSSLDGCR = rsAdo("DOSSLDGCR")
rsYDOSSLD1.DOSSLDGSD = rsAdo("DOSSLDGSD")

Exit Function

Error_Handler:

rsYDOSSLD1_GetBuffer = Error

End Function









Public Sub rsYDOSSLD1_Init(lYDOSSLD1 As typeYDOSSLD1)
lYDOSSLD1.DOSSLDPCI = ""
lYDOSSLD1.DOSSLDDEV = ""
lYDOSSLD1.DOSSLDCLI = ""
lYDOSSLD1.DOSSLDMDB = 0
lYDOSSLD1.DOSSLDGCR = 0
lYDOSSLD1.DOSSLDMCR = 0
lYDOSSLD1.DOSSLDGDB = 0
lYDOSSLD1.DOSSLDMSD = 0
lYDOSSLD1.DOSSLDGSD = 0

End Sub



