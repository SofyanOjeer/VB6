Attribute VB_Name = "rsZCDOIRR0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCDOIRR0
    CDOIRRETB       As Integer                        ' CODE ETABLISSEMENT
    CDOIRRAGE       As Integer                        ' AGENCE
    CDOIRRSER       As String * 2                     ' SERVICE
    CDOIRRSSE       As String * 2                     ' SOUS-SERVICE
    CDOIRRCOP       As String * 3                     ' CODE OPERATION
    CDOIRRDOS       As Long                           ' NUMERO DOSSIER
    CDOIRRNUR       As Long                           ' N° RENOUVELLEMENT
    CDOIRRUTI       As Long                           ' N° UTILISATION
    CDOIRRSEQ       As Long                           ' N° SEQUENCE
    CDOIRRTEX       As String * 75                    ' TEXTE

End Type
Public Sub rsZCDOIRR0_Init(rsYCDOIRR0 As typeZCDOIRR0)
rsYCDOIRR0.CDOIRRETB = 0
rsYCDOIRR0.CDOIRRAGE = 0
rsYCDOIRR0.CDOIRRSER = ""
rsYCDOIRR0.CDOIRRSSE = ""
rsYCDOIRR0.CDOIRRCOP = ""
rsYCDOIRR0.CDOIRRDOS = 0
rsYCDOIRR0.CDOIRRNUR = 0
rsYCDOIRR0.CDOIRRUTI = 0
rsYCDOIRR0.CDOIRRSEQ = 0
rsYCDOIRR0.CDOIRRTEX = ""
End Sub
Public Function rsZCDOIRR0_GetBuffer(rsAdo As ADODB.Recordset, rsZCDOIRR0 As typeZCDOIRR0)
On Error GoTo Error_Handler
rsZCDOIRR0_GetBuffer = Null
rsZCDOIRR0.CDOIRRETB = rsAdo("CDOIRRETB")
rsZCDOIRR0.CDOIRRAGE = rsAdo("CDOIRRAGE")
rsZCDOIRR0.CDOIRRSER = rsAdo("CDOIRRSER")
rsZCDOIRR0.CDOIRRSSE = rsAdo("CDOIRRSSE")
rsZCDOIRR0.CDOIRRCOP = rsAdo("CDOIRRCOP")
rsZCDOIRR0.CDOIRRDOS = rsAdo("CDOIRRDOS")
rsZCDOIRR0.CDOIRRNUR = rsAdo("CDOIRRNUR")
rsZCDOIRR0.CDOIRRUTI = rsAdo("CDOIRRUTI")
rsZCDOIRR0.CDOIRRSEQ = rsAdo("CDOIRRSEQ")
rsZCDOIRR0.CDOIRRTEX = rsAdo("CDOIRRTEX")
Exit Function
Error_Handler:
rsZCDOIRR0_GetBuffer = Error
End Function

