Attribute VB_Name = "rsZCDODES0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCDODES0
    CDODESETB       As Integer                        ' CODE ETABLISSEMENT
    CDODESAGE       As Integer                        ' AGENCE
    CDODESSER       As String * 2                     ' SERVICE
    CDODESSSE       As String * 2                     ' SOUS-SERVICE
    CDODESCOP       As String * 3                     ' CODE OPERATION
    CDODESDOS       As Long                           ' NUMERO DOSSIER
    CDODESNUR       As Long                           ' N° RENOUVELLEMENT
    CDODESUTI       As Long                           ' N° UTILISATION
    CDODESSEQ       As Long                           ' N° SEQUENCE
    CDODESTEX       As String * 65                    ' TEXTE

End Type
Public Sub rsZCDODES0_Init(rsYCDODES0 As typeZCDODES0)
rsYCDODES0.CDODESETB = 0
rsYCDODES0.CDODESAGE = 0
rsYCDODES0.CDODESSER = ""
rsYCDODES0.CDODESSSE = ""
rsYCDODES0.CDODESCOP = ""
rsYCDODES0.CDODESDOS = 0
rsYCDODES0.CDODESNUR = 0
rsYCDODES0.CDODESUTI = 0
rsYCDODES0.CDODESSEQ = 0
rsYCDODES0.CDODESTEX = ""
End Sub
Public Function rsZCDODES0_GetBuffer(rsAdo As ADODB.Recordset, rsZCDODES0 As typeZCDODES0)
On Error GoTo Error_Handler
rsZCDODES0_GetBuffer = Null
rsZCDODES0.CDODESETB = rsAdo("CDODESETB")
rsZCDODES0.CDODESAGE = rsAdo("CDODESAGE")
rsZCDODES0.CDODESSER = rsAdo("CDODESSER")
rsZCDODES0.CDODESSSE = rsAdo("CDODESSSE")
rsZCDODES0.CDODESCOP = rsAdo("CDODESCOP")
rsZCDODES0.CDODESDOS = rsAdo("CDODESDOS")
rsZCDODES0.CDODESNUR = rsAdo("CDODESNUR")
rsZCDODES0.CDODESUTI = rsAdo("CDODESUTI")
rsZCDODES0.CDODESSEQ = rsAdo("CDODESSEQ")
rsZCDODES0.CDODESTEX = rsAdo("CDODESTEX")
Exit Function
Error_Handler:
rsZCDODES0_GetBuffer = Error
End Function

