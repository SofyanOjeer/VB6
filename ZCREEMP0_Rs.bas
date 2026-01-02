Attribute VB_Name = "rsZCREEMP0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCREEMP0
    CREEMPETA       As Integer                        ' ETABLISSEMENT
    CREEMPAGE       As Integer                        ' AGENCE
    CREEMPSER       As String * 2                     ' SERVICE
    CREEMPSSE       As String * 2                     ' SOUS-SERVICE
    CREEMPDOS       As Long                           ' NUMERO DOSSIER
    CREEMPSEQ       As Long                           ' NUMERO SEQUENCE
    CREEMPNCL       As String * 7                     ' N° CLIENT

End Type
Public Sub rsZCREEMP0_Init(rsYCREEMP0 As typeZCREEMP0)
rsYCREEMP0.CREEMPETA = 0
rsYCREEMP0.CREEMPAGE = 0
rsYCREEMP0.CREEMPSER = ""
rsYCREEMP0.CREEMPSSE = ""
rsYCREEMP0.CREEMPDOS = 0
rsYCREEMP0.CREEMPSEQ = 0
rsYCREEMP0.CREEMPNCL = ""
End Sub
Public Function rsZCREEMP0_GetBuffer(rsAdo As ADODB.Recordset, rsZCREEMP0 As typeZCREEMP0)
On Error GoTo Error_Handler
rsZCREEMP0_GetBuffer = Null
rsZCREEMP0.CREEMPETA = rsAdo("CREEMPETA")
rsZCREEMP0.CREEMPAGE = rsAdo("CREEMPAGE")
rsZCREEMP0.CREEMPSER = rsAdo("CREEMPSER")
rsZCREEMP0.CREEMPSSE = rsAdo("CREEMPSSE")
rsZCREEMP0.CREEMPDOS = rsAdo("CREEMPDOS")
rsZCREEMP0.CREEMPSEQ = rsAdo("CREEMPSEQ")
rsZCREEMP0.CREEMPNCL = rsAdo("CREEMPNCL")
Exit Function
Error_Handler:
rsZCREEMP0_GetBuffer = Error
End Function

