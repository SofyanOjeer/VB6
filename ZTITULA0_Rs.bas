Attribute VB_Name = "rsZTITULA0"
'---------------------------------------------------------
Option Explicit
Type typeZTITULA0

    TITULAETA       As Integer                        ' ETABLISSEMENT
    TITULAPLA       As Long                           ' NUMERO PLAN
    TITULACOM       As String * 20                    ' NUMERO COMPTE
    TITULACLI       As String * 7                     ' NUMERO CLIENT
    TITULAPRI       As String * 1                     ' 0:PRINCIPAL, 1:AUTRE
    TITULATPR       As String * 1                     ' 0:PRINCIPAL, 1:AUTRE

End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZTITULA0_GetBuffer(rsAdo As ADODB.Recordset, rsZTITULA0 As typeZTITULA0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZTITULA0_GetBuffer = Null

    rsZTITULA0.TITULAETA = rsAdo("TITULAETA")    '
    rsZTITULA0.TITULAPLA = rsAdo("TITULAPLA")
    rsZTITULA0.TITULACOM = rsAdo("TITULACOM")
    rsZTITULA0.TITULACLI = rsAdo("TITULACLI")
    rsZTITULA0.TITULAPRI = rsAdo("TITULAPRI")
    rsZTITULA0.TITULATPR = rsAdo("TITULATPR")

Exit Function

Error_Handler:

rsZTITULA0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsZTITULA0_Init(rsZTITULA0 As typeZTITULA0)
'---------------------------------------------------------

End Sub


'








