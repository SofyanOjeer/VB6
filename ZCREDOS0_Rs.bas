Attribute VB_Name = "rsZCREDOS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCREDOS0
    CREDOSETA       As Integer                        ' ETABLISSEMENT
    CREDOSAGE       As Integer                        ' AGENCE
    CREDOSSER       As String * 2                     ' SERVICE
    CREDOSSSE       As String * 2                     ' SOUS-SERVICE
    CREDOSDOS       As Long                           ' NUMERO DOSSIER
    CREDOSNCR       As String * 3                     ' NATURE CREDIT
    CREDOSMNT       As Currency                       ' MONTANT
    CREDOSDEV       As String * 3                     ' DEVISE
    CREDOSDDE       As Long                           ' AUTORISATION
    CREDOSDFI       As Long                           ' AUTORISATION
    CREDOSREF       As String * 50                    ' REFERENCES
    CREDOSUTI       As Integer                        ' UTILISATEUR
    CREDOSDMO       As Long                           ' DATE MODIFICATION
    CREDOSOFI       As String * 6                     ' OBJET FINANCEMENT
    CREDOSCET       As Long                           ' CODE ETAT
    CREDOSDCE       As Long                           ' DATE CODE ETAT
    CREDOSDOD       As Long                           ' DU DOSSIER
    CREDOSDVA       As Long                           ' DATE VALIDATION
    CREDOSDGE       As Long                           ' CREDIT ENGAGE
    CREDOSTYP       As String * 1                     ' TYPE DE CREDIT
    CREDOSCOP       As Long                           ' CO-PARTICIPATION

End Type
Public Sub rsZCREDOS0_Init(rsYCREDOS0 As typeZCREDOS0)
rsYCREDOS0.CREDOSETA = 0
rsYCREDOS0.CREDOSAGE = 0
rsYCREDOS0.CREDOSSER = ""
rsYCREDOS0.CREDOSSSE = ""
rsYCREDOS0.CREDOSDOS = 0
rsYCREDOS0.CREDOSNCR = ""
rsYCREDOS0.CREDOSMNT = 0
rsYCREDOS0.CREDOSDEV = ""
rsYCREDOS0.CREDOSDDE = 0
rsYCREDOS0.CREDOSDFI = 0
rsYCREDOS0.CREDOSREF = ""
rsYCREDOS0.CREDOSUTI = 0
rsYCREDOS0.CREDOSDMO = 0
rsYCREDOS0.CREDOSOFI = ""
rsYCREDOS0.CREDOSCET = 0
rsYCREDOS0.CREDOSDCE = 0
rsYCREDOS0.CREDOSDOD = 0
rsYCREDOS0.CREDOSDVA = 0
rsYCREDOS0.CREDOSDGE = 0
rsYCREDOS0.CREDOSTYP = ""
rsYCREDOS0.CREDOSCOP = 0
End Sub
Public Function rsZCREDOS0_GetBuffer(rsAdo As ADODB.Recordset, rsZCREDOS0 As typeZCREDOS0)
On Error GoTo Error_Handler
rsZCREDOS0_GetBuffer = Null
rsZCREDOS0.CREDOSETA = rsAdo("CREDOSETA")
rsZCREDOS0.CREDOSAGE = rsAdo("CREDOSAGE")
rsZCREDOS0.CREDOSSER = rsAdo("CREDOSSER")
rsZCREDOS0.CREDOSSSE = rsAdo("CREDOSSSE")
rsZCREDOS0.CREDOSDOS = rsAdo("CREDOSDOS")
rsZCREDOS0.CREDOSNCR = rsAdo("CREDOSNCR")
rsZCREDOS0.CREDOSMNT = rsAdo("CREDOSMNT")
rsZCREDOS0.CREDOSDEV = rsAdo("CREDOSDEV")
rsZCREDOS0.CREDOSDDE = rsAdo("CREDOSDDE")
rsZCREDOS0.CREDOSDFI = rsAdo("CREDOSDFI")
rsZCREDOS0.CREDOSREF = rsAdo("CREDOSREF")
rsZCREDOS0.CREDOSUTI = rsAdo("CREDOSUTI")
rsZCREDOS0.CREDOSDMO = rsAdo("CREDOSDMO")
rsZCREDOS0.CREDOSOFI = rsAdo("CREDOSOFI")
rsZCREDOS0.CREDOSCET = rsAdo("CREDOSCET")
rsZCREDOS0.CREDOSDCE = rsAdo("CREDOSDCE")
rsZCREDOS0.CREDOSDOD = rsAdo("CREDOSDOD")
rsZCREDOS0.CREDOSDVA = rsAdo("CREDOSDVA")
rsZCREDOS0.CREDOSDGE = rsAdo("CREDOSDGE")
rsZCREDOS0.CREDOSTYP = rsAdo("CREDOSTYP")
rsZCREDOS0.CREDOSCOP = rsAdo("CREDOSCOP")
Exit Function
Error_Handler:
rsZCREDOS0_GetBuffer = Error
End Function

