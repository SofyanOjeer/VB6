Attribute VB_Name = "rsZCDOCOM0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCDOCOM0
    CDOCOMETB       As Integer                        ' CODE ETABLISSEMENT
    CDOCOMAGE       As Integer                        ' AGENCE
    CDOCOMSER       As String * 2                     ' SERVICE
    CDOCOMSSE       As String * 2                     ' SOUS-SERVICE
    CDOCOMCOP       As String * 3                     ' CODE OPERATION
    CDOCOMDOS       As Long                           ' NUMERO DOSSIER
    CDOCOMNUR       As Long                           ' N° RENOUVELLEMENT
    CDOCOMUTI       As Long                           ' N° UTILILSAT°./MODIF
    CDOCOMEVE       As String * 2                     ' EVENEMENT
    CDOCOMSEQ       As Long                           ' N° SEQUENCE
    CDOCOMCOM       As String * 6                     ' CODE COMMISSION
    CDOCOMDEM       As Long                           ' DT DEMANDE
    CDOCOMREG       As Long                           ' DT REGLEMENT
    CDOCOMCPT       As String * 20                    ' NUMERO DU COMPTE
    CDOCOMDEV       As String * 3                     ' DEVISE COMMISSION
    CDOCOMVAL       As Long                           ' DATE VALEUR
    CDOCOMCOU       As Double                         ' COURS DEVCOM/DEVCPT
    CDOCOMMRE       As String * 3                     ' MODE DE REGLEMENT
    CDOCOMBEN       As String * 1                     ' BENEFICIAIRE O/N
    CDOCOMMON       As Currency                       ' MONTANT COMMISSION
    CDOCOMMTV       As Currency                       ' MONTANT TVA
    CDOCOMAVI       As String * 1                     ' 1 NON,2 A EDIT,3 EDI
    CDOCOMPRO       As String * 1                     ' A PROVISIONNER (O/N)
    CDOCOMUTR       As Long                           ' UTILISATION DU REGLE
    CDOCOMNRE       As Long                           ' N° REGLEMENT
    CDOCOMETA       As String * 2                     ' ETAT
    CDOCOMSPE       As Long                           ' N°SEQUENCE PERIODIQ
    CDOCOMDBP       As Long                           ' DATE DEBUT PERIODE
    CDOCOMFNP       As Long                           ' DATE FIN PERIODE
    CDOCOMCUT       As Integer                        ' UTILISATEUR SAISIE
    CDOCOMCER       As String * 1                     ' COTAT°(O=CERTAIN/N)

End Type
Public Sub rsZCDOCOM0_Init(rsYCDOCOM0 As typeZCDOCOM0)
rsYCDOCOM0.CDOCOMETB = 0
rsYCDOCOM0.CDOCOMAGE = 0
rsYCDOCOM0.CDOCOMSER = ""
rsYCDOCOM0.CDOCOMSSE = ""
rsYCDOCOM0.CDOCOMCOP = ""
rsYCDOCOM0.CDOCOMDOS = 0
rsYCDOCOM0.CDOCOMNUR = 0
rsYCDOCOM0.CDOCOMUTI = 0
rsYCDOCOM0.CDOCOMEVE = ""
rsYCDOCOM0.CDOCOMSEQ = 0
rsYCDOCOM0.CDOCOMCOM = ""
rsYCDOCOM0.CDOCOMDEM = 0
rsYCDOCOM0.CDOCOMREG = 0
rsYCDOCOM0.CDOCOMCPT = ""
rsYCDOCOM0.CDOCOMDEV = ""
rsYCDOCOM0.CDOCOMVAL = 0
rsYCDOCOM0.CDOCOMCOU = 0
rsYCDOCOM0.CDOCOMMRE = ""
rsYCDOCOM0.CDOCOMBEN = ""
rsYCDOCOM0.CDOCOMMON = 0
rsYCDOCOM0.CDOCOMMTV = 0
rsYCDOCOM0.CDOCOMAVI = ""
rsYCDOCOM0.CDOCOMPRO = ""
rsYCDOCOM0.CDOCOMUTR = 0
rsYCDOCOM0.CDOCOMNRE = 0
rsYCDOCOM0.CDOCOMETA = ""
rsYCDOCOM0.CDOCOMSPE = 0
rsYCDOCOM0.CDOCOMDBP = 0
rsYCDOCOM0.CDOCOMFNP = 0
rsYCDOCOM0.CDOCOMCUT = 0
rsYCDOCOM0.CDOCOMCER = ""
End Sub
Public Function rsZCDOCOM0_GetBuffer(rsAdo As ADODB.Recordset, rsZCDOCOM0 As typeZCDOCOM0)
On Error GoTo Error_Handler
rsZCDOCOM0_GetBuffer = Null
rsZCDOCOM0.CDOCOMETB = rsAdo("CDOCOMETB")
rsZCDOCOM0.CDOCOMAGE = rsAdo("CDOCOMAGE")
rsZCDOCOM0.CDOCOMSER = rsAdo("CDOCOMSER")
rsZCDOCOM0.CDOCOMSSE = rsAdo("CDOCOMSSE")
rsZCDOCOM0.CDOCOMCOP = rsAdo("CDOCOMCOP")
rsZCDOCOM0.CDOCOMDOS = rsAdo("CDOCOMDOS")
rsZCDOCOM0.CDOCOMNUR = rsAdo("CDOCOMNUR")
rsZCDOCOM0.CDOCOMUTI = rsAdo("CDOCOMUTI")
rsZCDOCOM0.CDOCOMEVE = rsAdo("CDOCOMEVE")
rsZCDOCOM0.CDOCOMSEQ = rsAdo("CDOCOMSEQ")
rsZCDOCOM0.CDOCOMCOM = rsAdo("CDOCOMCOM")
rsZCDOCOM0.CDOCOMDEM = rsAdo("CDOCOMDEM")
rsZCDOCOM0.CDOCOMREG = rsAdo("CDOCOMREG")
rsZCDOCOM0.CDOCOMCPT = rsAdo("CDOCOMCPT")
rsZCDOCOM0.CDOCOMDEV = rsAdo("CDOCOMDEV")
rsZCDOCOM0.CDOCOMVAL = rsAdo("CDOCOMVAL")
rsZCDOCOM0.CDOCOMCOU = rsAdo("CDOCOMCOU")
rsZCDOCOM0.CDOCOMMRE = rsAdo("CDOCOMMRE")
rsZCDOCOM0.CDOCOMBEN = rsAdo("CDOCOMBEN")
rsZCDOCOM0.CDOCOMMON = rsAdo("CDOCOMMON")
rsZCDOCOM0.CDOCOMMTV = rsAdo("CDOCOMMTV")
rsZCDOCOM0.CDOCOMAVI = rsAdo("CDOCOMAVI")
rsZCDOCOM0.CDOCOMPRO = rsAdo("CDOCOMPRO")
rsZCDOCOM0.CDOCOMUTR = rsAdo("CDOCOMUTR")
rsZCDOCOM0.CDOCOMNRE = rsAdo("CDOCOMNRE")
rsZCDOCOM0.CDOCOMETA = rsAdo("CDOCOMETA")
rsZCDOCOM0.CDOCOMSPE = rsAdo("CDOCOMSPE")
rsZCDOCOM0.CDOCOMDBP = rsAdo("CDOCOMDBP")
rsZCDOCOM0.CDOCOMFNP = rsAdo("CDOCOMFNP")
rsZCDOCOM0.CDOCOMCUT = rsAdo("CDOCOMCUT")
rsZCDOCOM0.CDOCOMCER = rsAdo("CDOCOMCER")
Exit Function
Error_Handler:
rsZCDOCOM0_GetBuffer = Error
End Function

