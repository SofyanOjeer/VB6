Attribute VB_Name = "rsZCDOREG0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCDOREG0
    CDOREGETB       As Integer                        ' CODE ETABLISSEMENT
    CDOREGAGE       As Integer                        ' AGENCE
    CDOREGSER       As String * 2                     ' SERVICE
    CDOREGSSE       As String * 2                     ' SOUS-SERVICE
    CDOREGCOP       As String * 3                     ' CODE OPERATION
    CDOREGDOS       As Long                           ' NUMERO DOSSIER
    CDOREGNUR       As Long                           ' N° RENOUVELLEMENT
    CDOREGUTI       As Long                           ' N° UTILISATION
    CDOREGPAI       As Long                           ' N° PAIEMENT
    CDOREGREG       As Long                           ' N° REGLEMENT/ENCAIS
    CDOREGCRD       As String * 1                     ' CREDIT /DEBIT C/D
    CDOREGMON       As Currency                       ' MONTANT DEV. UTILIS
    CDOREGMOR       As Currency                       ' MONTANT REGLE/ENCAI
    CDOREGDEV       As String * 3                     ' DEVISE REGLEM/ENCAI
    CDOREGDRE       As Long                           ' DATE REGLEM/ENCAIS.
    CDOREGDEM       As Long                           ' DATE EMISSION
    CDOREGDCR       As Long                           ' DATE COMPTA REG/ENC
    CDOREGDUT       As Long                           ' DATE COMPTA UTILISA
    CDOREGRES       As Long                           ' REFERENCE ESCOMPTE
    CDOREGRRE       As Long                           ' REFERENCE REFINANC.
    CDOREGDEC       As String * 1                     ' DESTINAT. CLI/TIERS
    CDOREGDES       As String * 7                     ' DESTINATAIRE
    CDOREGMOD       As String * 3                     ' MODE REGLEMENT/ENCA
    CDOREGINT       As String * 1                     ' CPT NOSTRO  (O/N)
    CDOREGCOM       As String * 20                    ' COMPTE
    CDOREGINC       As String * 1                     ' INTERMED. CLI/TIERS
    CDOREGINS       As String * 7                     ' INTERMEDIAIRE
    CDOREGPAC       As String * 1                     ' BANQ DEST CLI/TIERS
    CDOREGPAS       As String * 7                     ' BANQ. DEST.-PAYEUR
    CDOREGENV       As Long                           ' DATE ENVOI COURRIER
    CDOREGCOU       As Double                         ' COURS DEVREG/DEVDOS
    CDOREGDEN       As Long                           ' DATE ENGAGEMENT
    CDOREGDRP       As Long                           ' DATE RECEP.PREVUE
    CDOREGDRR       As Long                           ' DATE RECEP.REELLE
    CDOREGDAE       As Long                           ' DATE ECHEANCE
    CDOREGDVA       As Long                           ' DATE VALEUR
    CDOREGDIC       As Long                           ' DATE INIT CHANGE
    CDOREGBDF       As String * 3                     ' CODE BDF
    CDOREGPAY       As String * 3                     ' CODE PAYS
    CDOREGSIR       As String * 9                     ' N°SIREN
    CDOREGTRN       As String * 16                    ' TRN SAGITTAIRE
    CDOREGTCR       As String * 1                     ' TYPE CRP
    CDOREGCBA       As Long                           ' CODE BANQUE
    CDOREGCGU       As Long                           ' CODE GUICHET
    CDOREGATG       As String * 1                     ' ATTENTE GEST.
    CDOREGVA1       As Integer                        ' 1ER VALIDEUR
    CDOREGVA2       As Integer                        ' 2EME VALIDEUR
    CDOREGEVE       As String * 2                     ' EVENEMENT
    CDOREGATT       As String * 2                     ' ATTENTE
    CDOREGETA       As String * 2                     ' ETAT
    CDOREGNUA       As Long                           ' NUM OPE ATT
    CDOREGCAA       As String * 12                    ' CODE AUTOR AVAL
    CDOREGCER       As String * 1                     ' COTAT°(O=CERTAIN/N)

End Type
Public Sub rsZCDOREG0_Init(rsYCDOREG0 As typeZCDOREG0)
rsYCDOREG0.CDOREGETB = 0
rsYCDOREG0.CDOREGAGE = 0
rsYCDOREG0.CDOREGSER = ""
rsYCDOREG0.CDOREGSSE = ""
rsYCDOREG0.CDOREGCOP = ""
rsYCDOREG0.CDOREGDOS = 0
rsYCDOREG0.CDOREGNUR = 0
rsYCDOREG0.CDOREGUTI = 0
rsYCDOREG0.CDOREGPAI = 0
rsYCDOREG0.CDOREGREG = 0
rsYCDOREG0.CDOREGCRD = ""
rsYCDOREG0.CDOREGMON = 0
rsYCDOREG0.CDOREGMOR = 0
rsYCDOREG0.CDOREGDEV = ""
rsYCDOREG0.CDOREGDRE = 0
rsYCDOREG0.CDOREGDEM = 0
rsYCDOREG0.CDOREGDCR = 0
rsYCDOREG0.CDOREGDUT = 0
rsYCDOREG0.CDOREGRES = 0
rsYCDOREG0.CDOREGRRE = 0
rsYCDOREG0.CDOREGDEC = ""
rsYCDOREG0.CDOREGDES = ""
rsYCDOREG0.CDOREGMOD = ""
rsYCDOREG0.CDOREGINT = ""
rsYCDOREG0.CDOREGCOM = ""
rsYCDOREG0.CDOREGINC = ""
rsYCDOREG0.CDOREGINS = ""
rsYCDOREG0.CDOREGPAC = ""
rsYCDOREG0.CDOREGPAS = ""
rsYCDOREG0.CDOREGENV = 0
rsYCDOREG0.CDOREGCOU = 0
rsYCDOREG0.CDOREGDEN = 0
rsYCDOREG0.CDOREGDRP = 0
rsYCDOREG0.CDOREGDRR = 0
rsYCDOREG0.CDOREGDAE = 0
rsYCDOREG0.CDOREGDVA = 0
rsYCDOREG0.CDOREGDIC = 0
rsYCDOREG0.CDOREGBDF = ""
rsYCDOREG0.CDOREGPAY = ""
rsYCDOREG0.CDOREGSIR = ""
rsYCDOREG0.CDOREGTRN = ""
rsYCDOREG0.CDOREGTCR = ""
rsYCDOREG0.CDOREGCBA = 0
rsYCDOREG0.CDOREGCGU = 0
rsYCDOREG0.CDOREGATG = ""
rsYCDOREG0.CDOREGVA1 = 0
rsYCDOREG0.CDOREGVA2 = 0
rsYCDOREG0.CDOREGEVE = ""
rsYCDOREG0.CDOREGATT = ""
rsYCDOREG0.CDOREGETA = ""
rsYCDOREG0.CDOREGNUA = 0
rsYCDOREG0.CDOREGCAA = ""
rsYCDOREG0.CDOREGCER = ""
End Sub
Public Function rsZCDOREG0_GetBuffer(rsAdo As ADODB.Recordset, rsZCDOREG0 As typeZCDOREG0)
On Error GoTo Error_Handler
rsZCDOREG0_GetBuffer = Null
rsZCDOREG0.CDOREGETB = rsAdo("CDOREGETB")
rsZCDOREG0.CDOREGAGE = rsAdo("CDOREGAGE")
rsZCDOREG0.CDOREGSER = rsAdo("CDOREGSER")
rsZCDOREG0.CDOREGSSE = rsAdo("CDOREGSSE")
rsZCDOREG0.CDOREGCOP = rsAdo("CDOREGCOP")
rsZCDOREG0.CDOREGDOS = rsAdo("CDOREGDOS")
rsZCDOREG0.CDOREGNUR = rsAdo("CDOREGNUR")
rsZCDOREG0.CDOREGUTI = rsAdo("CDOREGUTI")
rsZCDOREG0.CDOREGPAI = rsAdo("CDOREGPAI")
rsZCDOREG0.CDOREGREG = rsAdo("CDOREGREG")
rsZCDOREG0.CDOREGCRD = rsAdo("CDOREGCRD")
rsZCDOREG0.CDOREGMON = rsAdo("CDOREGMON")
rsZCDOREG0.CDOREGMOR = rsAdo("CDOREGMOR")
rsZCDOREG0.CDOREGDEV = rsAdo("CDOREGDEV")
rsZCDOREG0.CDOREGDRE = rsAdo("CDOREGDRE")
rsZCDOREG0.CDOREGDEM = rsAdo("CDOREGDEM")
rsZCDOREG0.CDOREGDCR = rsAdo("CDOREGDCR")
rsZCDOREG0.CDOREGDUT = rsAdo("CDOREGDUT")
rsZCDOREG0.CDOREGRES = rsAdo("CDOREGRES")
rsZCDOREG0.CDOREGRRE = rsAdo("CDOREGRRE")
rsZCDOREG0.CDOREGDEC = rsAdo("CDOREGDEC")
rsZCDOREG0.CDOREGDES = rsAdo("CDOREGDES")
rsZCDOREG0.CDOREGMOD = rsAdo("CDOREGMOD")
rsZCDOREG0.CDOREGINT = rsAdo("CDOREGINT")
rsZCDOREG0.CDOREGCOM = rsAdo("CDOREGCOM")
rsZCDOREG0.CDOREGINC = rsAdo("CDOREGINC")
rsZCDOREG0.CDOREGINS = rsAdo("CDOREGINS")
rsZCDOREG0.CDOREGPAC = rsAdo("CDOREGPAC")
rsZCDOREG0.CDOREGPAS = rsAdo("CDOREGPAS")
rsZCDOREG0.CDOREGENV = rsAdo("CDOREGENV")
rsZCDOREG0.CDOREGCOU = rsAdo("CDOREGCOU")
rsZCDOREG0.CDOREGDEN = rsAdo("CDOREGDEN")
rsZCDOREG0.CDOREGDRP = rsAdo("CDOREGDRP")
rsZCDOREG0.CDOREGDRR = rsAdo("CDOREGDRR")
rsZCDOREG0.CDOREGDAE = rsAdo("CDOREGDAE")
rsZCDOREG0.CDOREGDVA = rsAdo("CDOREGDVA")
rsZCDOREG0.CDOREGDIC = rsAdo("CDOREGDIC")
rsZCDOREG0.CDOREGBDF = rsAdo("CDOREGBDF")
rsZCDOREG0.CDOREGPAY = rsAdo("CDOREGPAY")
rsZCDOREG0.CDOREGSIR = rsAdo("CDOREGSIR")
rsZCDOREG0.CDOREGTRN = rsAdo("CDOREGTRN")
rsZCDOREG0.CDOREGTCR = rsAdo("CDOREGTCR")
rsZCDOREG0.CDOREGCBA = rsAdo("CDOREGCBA")
rsZCDOREG0.CDOREGCGU = rsAdo("CDOREGCGU")
rsZCDOREG0.CDOREGATG = rsAdo("CDOREGATG")
rsZCDOREG0.CDOREGVA1 = rsAdo("CDOREGVA1")
rsZCDOREG0.CDOREGVA2 = rsAdo("CDOREGVA2")
rsZCDOREG0.CDOREGEVE = rsAdo("CDOREGEVE")
rsZCDOREG0.CDOREGATT = rsAdo("CDOREGATT")
rsZCDOREG0.CDOREGETA = rsAdo("CDOREGETA")
rsZCDOREG0.CDOREGNUA = rsAdo("CDOREGNUA")
rsZCDOREG0.CDOREGCAA = rsAdo("CDOREGCAA")
rsZCDOREG0.CDOREGCER = rsAdo("CDOREGCER")
Exit Function
Error_Handler:
rsZCDOREG0_GetBuffer = Error
End Function

