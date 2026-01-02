Attribute VB_Name = "rsZCDODOS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCDODOS0
    CDODOSETB       As Integer                        ' CODE ETABLISSEMENT
    CDODOSAGE       As Integer                        ' AGENCE
    CDODOSSER       As String * 2                     ' SERVICE
    CDODOSSSE       As String * 2                     ' SOUS-SERVICE
    CDODOSCOP       As String * 3                     ' CODE OPERATION
    CDODOSDOS       As Long                           ' NUMERO DOSSIER
    CDODOSNUR       As Long                           ' N° RENOUVELLEMENT
    CDODOSNAT       As String * 3                     ' NATURE
    CDODOSEXT       As String * 16                    ' REFERENCE EXTERNE
    CDODOSMON       As Currency                       ' MONTANT DOSSIER
    CDODOSDEV       As String * 3                     ' DEVISE
    CDODOSMOA       As Currency                       ' MONTANT ADDITIONNEL
    CDODOSMOT       As Currency                       ' MONTANT TOTAL
    CDODOSMOC       As Currency                       ' MONTANT CONFIRME
    CDODOSMOD       As Currency                       ' MONTANT DUCROIRE
    CDODOSCON       As String * 1                     ' CONFIRM NOTIFI PARTI
    CDODOSIRR       As String * 1                     ' IRREVOCABLE (O/N)
    CDODOSFRA       As String * 1                     ' FRACTIONNABLE (O/N)
    CDODOSREN       As String * 1                     ' RENOUVELABLE (O/N)
    CDODOSCUM       As String * 1                     ' CUMULATIF (O/N)
    CDODOSTRS       As String * 1                     ' TRANSFERABLE
    CDODOSTOL       As Currency                       ' TOLERANCE +
    CDODOSTO2       As Currency                       ' TOLERANCE -
    CDODOSDOR       As String * 1                     ' DONN. ORDRE CLI/TIE
    CDODOSDON       As String * 7                     ' DONNEUR ORDRE IMPORT
    CDODOSDOE       As String * 64                    ' DONNEUR ORDRE EXPORT
    CDODOSBER       As String * 1                     ' BENEFICIAIR CLI/TIE
    CDODOSBEN       As String * 7                     ' BENEFICIAIRE EXPORT
    CDODOSBEI       As String * 64                    ' BENEFICIAIRE IMPORT
    CDODOSBAR       As String * 1                     ' BANQU.BENEF.CLI/TIE
    CDODOSBAB       As String * 7                     ' BANQUE BENEF
    CDODOSNOR       As String * 1                     ' NOTIF/CONFI OU EMETT
    CDODOSNOT       As String * 7                     ' NOTIF/CONFI OU EMETT
    CDODOSBIC       As String * 12                    ' BIC SUPPLEMEN. EMETT
    CDODOSCOT       As String * 1                     ' CORRESPOND. CLI/TIE
    CDODOSCOR       As String * 7                     ' CORRESPONDANT
    CDODOSPRT       As String * 1                     ' LIEU PRES CLI/TIE
    CDODOSPRR       As String * 7                     ' LIEU PRESENTATION
    CDODOSUTV       As String * 32                    ' LIEU PRESENTATION
    CDODOSPAT       As String * 1                     ' LIEU PAIE CLI/TIE
    CDODOSPAR       As String * 7                     ' LIEU PAIEMENT
    CDODOSPAV       As String * 32                    ' LIEU PAIEMENT
    CDODOSOUV       As Long                           ' DATE OUVERTURE
    CDODOSEMI       As Long                           ' DATE EMISSION
    CDODOSVAL       As Long                           ' DATE VALIDITE
    CDODOSDEP       As Long                           ' DATE EXTREME PAYMT
    CDODOSDTR       As Long                           ' DATE DE TRANSFERT
    CDODOSVCP       As Long                           ' DATE VALID. COMPTA
    CDODOSCLO       As Long                           ' DATE CLOTURE
    CDODOSREJ       As String * 3                     ' MOTIF REJET (CLOTUR)
    CDODOSOBJ       As String * 6                     ' OBJET CREDIT
    CDODOSAVU       As Long                           ' % PAIEM. A VUE
    CDODOSMOV       As Currency                       ' MONTANT A VUE
    CDODOSCAC       As Long                           ' % PAIEM. CTR ACCEPT.
    CDODOSMCA       As Currency                       ' MONTANT CTR ACCEPT.
    CDODOSDIF       As Long                           ' % PAIEM. DIFFERE
    CDODOSMDI       As Currency                       ' MONTANT. DIFFERE
    CDODOSPMO       As Currency                       ' MONTANT PROVISIONNE
    CDODOSPCD       As String * 20                    ' PROV. DEBIT  COMPTE
    CDODOSPCC       As String * 20                    ' PROV. CREDIT COMPTE
    CDODOSPDE       As Currency                       ' PROVISION DEVISE DOS
    CDODOSPPO       As Long                           ' PROVISION POURCEN
    CDODOSAUT       As String * 12                    ' CODE AUTORISATION
    CDODOSREG       As Currency                       ' MONTANT PAYE
    CDODOSENC       As Currency                       ' MONTANT ENCAISSE
    CDODOSDAN       As Long                           ' DATE ANNULATION
    CDODOSANN       As Currency                       ' MONTANT ANNULE
    CDODOSPCO       As Double                         ' COURS DEVPRO/DEVDOS
    CDODOSLEM       As String * 30                    ' LIEU EMBARQUEMENT
    CDODOSLDE       As String * 30                    ' LIEU DESTINATION
    CDODOSDLE       As Long                           ' DATE LIMITE EMBARQU.
    CDODOSEPA       As String * 1                     ' EXPED.PARTIE.AUTORI
    CDODOSTRA       As String * 1                     ' TRANBORDEMENT AUTORI
    CDODOSFCD       As String * 1                     ' FRAI CHARGE D.O. BEN
    CDODOSCUS       As Integer                        ' UTILI. DE SAISIE
    CDODOSCUV       As Integer                        ' 1ER VALIDEUR
    CDODOSCU2       As Integer                        ' 2EME VALIDEUR
    CDODOSOPE       As String * 1                     ' OPERATIVITE DU CRED.
    CDODOSPOO       As String * 1                     ' EXISTENCE POOL
    CDODOSPBE       As Currency                       ' PART.BANQUE EXPORT
    CDODOSGAG       As String * 1                     ' GAGE MARCHANDISE
    CDODOSSTB       As String * 1                     ' STAND BY
    CDODOSMRE       As String * 3                     ' MODE DE REALISAT°
    CDODOSNPD       As Long                           ' NBJ PRES. DOCUMENT
    CDODOSTJD       As String * 1                     ' TY JOUR DOCS
    CDODOSPDO       As String * 60                    ' PER.PRE.DOCS.
    CDODOSGAR       As String * 64                    ' LIBELLE GARANTIE
    CDODOSOBM       As String * 64                    ' OBJET DE MODIF.
    CDODOSTBR       As String * 1                     ' TIERS BQ REMBOURS
    CDODOSBRE       As String * 7                     ' BQ REMBOURSEMENT
    CDODOSBEC       As String * 1                     ' BENEF PAY.COMMIS°
    CDODOSRNO       As String * 16                    ' REF.NOTIFICATEUR
    CDODOSDPA       As String * 3                     ' DESTINATION PAYS
    CDODOSDVI       As String * 32                    ' DESTINATION VILLE
    CDODOSEPY       As String * 3                     ' EMBARQUEMENT PAYS
    CDODOSEVI       As String * 32                    ' EMBARQUEMENT VILLE
    CDODOSVPA       As String * 3                     ' VALIDITE PAYS
    CDODOSVVI       As String * 32                    ' VALIDIT VILLE
    CDODOSNDE       As Long                           ' DOSSIER EXPORT
    CDODOSNAE       As String * 3                     ' NATURE EXPORT
    CDODOSEVE       As String * 2                     ' EVENEMENT
    CDODOSETA       As String * 2                     ' ETAT DOSSIER
    CDODOSDP2       As String * 32                    ' DESTIN.PAYS LIBELLE
    CDODOSEP2       As String * 32                    ' EMBARQ.PAYS LIBELLE
    CDODOSPD2       As String * 80                    ' PER.PRES.DOC.SUITE
    CDODOSAUN       As String * 12                    ' CODE AUT. NOTIFIE
    CDODOSCER       As String * 1                     ' COTAT°(O=CERTAIN/N)
    CDODOSCRE       As String * 12                    ' CODE REGLE
    CDODOSREM       As String * 35                    ' LIBELLE EMISSION
    CDODOSRGR       As String * 1                     ' REGLE RBT
    CDODOSLED       As String * 65                    ' LIEU EMBARQMT
    CDODOSLDA       As String * 65                    ' LIEU DEBARQMT

End Type


Type typeWCDOCOM0
    
    WCDOCOMCOM       As String * 6                     ' CODE COMMISSION
    WCDOCOMMON       As Currency                       ' MONTANT COMMISSION
    WCDOCOMDEV       As String * 3                     ' DEVISE COMMISSION
    WCDOCOMMTV       As Currency                       ' MONTANT TVA

    WCDOCO2TX1       As Double                         ' Taux tranche 1
    WCDOCO2PER       As String * 1                     ' Périodicité
    WCDOCO2MIN       As Currency                       ' MONTANT minimum
    
    WCDOTC2DEV       As String * 3                     ' Devise
    WCDOTC2MTF       As Currency                       ' Montant fixe

End Type

Public Sub rsZCDODOS0_Init(rsZCDODOS0 As typeZCDODOS0)
rsZCDODOS0.CDODOSETB = 0
rsZCDODOS0.CDODOSAGE = 0
rsZCDODOS0.CDODOSSER = ""
rsZCDODOS0.CDODOSSSE = ""
rsZCDODOS0.CDODOSCOP = ""
rsZCDODOS0.CDODOSDOS = 0
rsZCDODOS0.CDODOSNUR = 0
rsZCDODOS0.CDODOSNAT = ""
rsZCDODOS0.CDODOSEXT = ""
rsZCDODOS0.CDODOSMON = 0
rsZCDODOS0.CDODOSDEV = ""
rsZCDODOS0.CDODOSMOA = 0
rsZCDODOS0.CDODOSMOT = 0
rsZCDODOS0.CDODOSMOC = 0
rsZCDODOS0.CDODOSMOD = 0
rsZCDODOS0.CDODOSCON = ""
rsZCDODOS0.CDODOSIRR = ""
rsZCDODOS0.CDODOSFRA = ""
rsZCDODOS0.CDODOSREN = ""
rsZCDODOS0.CDODOSCUM = ""
rsZCDODOS0.CDODOSTRS = ""
rsZCDODOS0.CDODOSTOL = 0
rsZCDODOS0.CDODOSTO2 = 0
rsZCDODOS0.CDODOSDOR = ""
rsZCDODOS0.CDODOSDON = ""
rsZCDODOS0.CDODOSDOE = ""
rsZCDODOS0.CDODOSBER = ""
rsZCDODOS0.CDODOSBEN = ""
rsZCDODOS0.CDODOSBEI = ""
rsZCDODOS0.CDODOSBAR = ""
rsZCDODOS0.CDODOSBAB = ""
rsZCDODOS0.CDODOSNOR = ""
rsZCDODOS0.CDODOSNOT = ""
rsZCDODOS0.CDODOSBIC = ""
rsZCDODOS0.CDODOSCOT = ""
rsZCDODOS0.CDODOSCOR = ""
rsZCDODOS0.CDODOSPRT = ""
rsZCDODOS0.CDODOSPRR = ""
rsZCDODOS0.CDODOSUTV = ""
rsZCDODOS0.CDODOSPAT = ""
rsZCDODOS0.CDODOSPAR = ""
rsZCDODOS0.CDODOSPAV = ""
rsZCDODOS0.CDODOSOUV = 0
rsZCDODOS0.CDODOSEMI = 0
rsZCDODOS0.CDODOSVAL = 0
rsZCDODOS0.CDODOSDEP = 0
rsZCDODOS0.CDODOSDTR = 0
rsZCDODOS0.CDODOSVCP = 0
rsZCDODOS0.CDODOSCLO = 0
rsZCDODOS0.CDODOSREJ = ""
rsZCDODOS0.CDODOSOBJ = ""
rsZCDODOS0.CDODOSAVU = 0
rsZCDODOS0.CDODOSMOV = 0
rsZCDODOS0.CDODOSCAC = 0
rsZCDODOS0.CDODOSMCA = 0
rsZCDODOS0.CDODOSDIF = 0
rsZCDODOS0.CDODOSMDI = 0
rsZCDODOS0.CDODOSPMO = 0
rsZCDODOS0.CDODOSPCD = ""
rsZCDODOS0.CDODOSPCC = ""
rsZCDODOS0.CDODOSPDE = 0
rsZCDODOS0.CDODOSPPO = 0
rsZCDODOS0.CDODOSAUT = ""
rsZCDODOS0.CDODOSREG = 0
rsZCDODOS0.CDODOSENC = 0
rsZCDODOS0.CDODOSDAN = 0
rsZCDODOS0.CDODOSANN = 0
rsZCDODOS0.CDODOSPCO = 0
rsZCDODOS0.CDODOSLEM = ""
rsZCDODOS0.CDODOSLDE = ""
rsZCDODOS0.CDODOSDLE = 0
rsZCDODOS0.CDODOSEPA = ""
rsZCDODOS0.CDODOSTRA = ""
rsZCDODOS0.CDODOSFCD = ""
rsZCDODOS0.CDODOSCUS = 0
rsZCDODOS0.CDODOSCUV = 0
rsZCDODOS0.CDODOSCU2 = 0
rsZCDODOS0.CDODOSOPE = ""
rsZCDODOS0.CDODOSPOO = ""
rsZCDODOS0.CDODOSPBE = 0
rsZCDODOS0.CDODOSGAG = ""
rsZCDODOS0.CDODOSSTB = ""
rsZCDODOS0.CDODOSMRE = ""
rsZCDODOS0.CDODOSNPD = 0
rsZCDODOS0.CDODOSTJD = ""
rsZCDODOS0.CDODOSPDO = ""
rsZCDODOS0.CDODOSGAR = ""
rsZCDODOS0.CDODOSOBM = ""
rsZCDODOS0.CDODOSTBR = ""
rsZCDODOS0.CDODOSBRE = ""
rsZCDODOS0.CDODOSBEC = ""
rsZCDODOS0.CDODOSRNO = ""
rsZCDODOS0.CDODOSDPA = ""
rsZCDODOS0.CDODOSDVI = ""
rsZCDODOS0.CDODOSEPY = ""
rsZCDODOS0.CDODOSEVI = ""
rsZCDODOS0.CDODOSVPA = ""
rsZCDODOS0.CDODOSVVI = ""
rsZCDODOS0.CDODOSNDE = 0
rsZCDODOS0.CDODOSNAE = ""
rsZCDODOS0.CDODOSEVE = ""
rsZCDODOS0.CDODOSETA = ""
rsZCDODOS0.CDODOSDP2 = ""
rsZCDODOS0.CDODOSEP2 = ""
rsZCDODOS0.CDODOSPD2 = ""
rsZCDODOS0.CDODOSAUN = ""
rsZCDODOS0.CDODOSCER = ""
rsZCDODOS0.CDODOSCRE = ""
rsZCDODOS0.CDODOSREM = ""
rsZCDODOS0.CDODOSRGR = ""
rsZCDODOS0.CDODOSLED = ""
rsZCDODOS0.CDODOSLDA = ""
End Sub

Public Function rsZCDODOS0_GetBuffer(rsAdo As ADODB.Recordset, rsZCDODOS0 As typeZCDODOS0)
On Error GoTo Error_Handler
rsZCDODOS0_GetBuffer = Null
rsZCDODOS0.CDODOSETB = rsAdo("CDODOSETB")
rsZCDODOS0.CDODOSAGE = rsAdo("CDODOSAGE")
rsZCDODOS0.CDODOSSER = rsAdo("CDODOSSER")
rsZCDODOS0.CDODOSSSE = rsAdo("CDODOSSSE")
rsZCDODOS0.CDODOSCOP = rsAdo("CDODOSCOP")
rsZCDODOS0.CDODOSDOS = rsAdo("CDODOSDOS")
rsZCDODOS0.CDODOSNUR = rsAdo("CDODOSNUR")
rsZCDODOS0.CDODOSNAT = rsAdo("CDODOSNAT")
rsZCDODOS0.CDODOSEXT = rsAdo("CDODOSEXT")
rsZCDODOS0.CDODOSMON = rsAdo("CDODOSMON")
rsZCDODOS0.CDODOSDEV = rsAdo("CDODOSDEV")
rsZCDODOS0.CDODOSMOA = rsAdo("CDODOSMOA")
rsZCDODOS0.CDODOSMOT = rsAdo("CDODOSMOT")
rsZCDODOS0.CDODOSMOC = rsAdo("CDODOSMOC")
rsZCDODOS0.CDODOSMOD = rsAdo("CDODOSMOD")
rsZCDODOS0.CDODOSCON = rsAdo("CDODOSCON")
rsZCDODOS0.CDODOSIRR = rsAdo("CDODOSIRR")
rsZCDODOS0.CDODOSFRA = rsAdo("CDODOSFRA")
rsZCDODOS0.CDODOSREN = rsAdo("CDODOSREN")
rsZCDODOS0.CDODOSCUM = rsAdo("CDODOSCUM")
rsZCDODOS0.CDODOSTRS = rsAdo("CDODOSTRS")
rsZCDODOS0.CDODOSTOL = rsAdo("CDODOSTOL")
rsZCDODOS0.CDODOSTO2 = rsAdo("CDODOSTO2")
rsZCDODOS0.CDODOSDOR = rsAdo("CDODOSDOR")
rsZCDODOS0.CDODOSDON = rsAdo("CDODOSDON")
rsZCDODOS0.CDODOSDOE = rsAdo("CDODOSDOE")
rsZCDODOS0.CDODOSBER = rsAdo("CDODOSBER")
rsZCDODOS0.CDODOSBEN = rsAdo("CDODOSBEN")
rsZCDODOS0.CDODOSBEI = rsAdo("CDODOSBEI")
rsZCDODOS0.CDODOSBAR = rsAdo("CDODOSBAR")
rsZCDODOS0.CDODOSBAB = rsAdo("CDODOSBAB")
rsZCDODOS0.CDODOSNOR = rsAdo("CDODOSNOR")
rsZCDODOS0.CDODOSNOT = rsAdo("CDODOSNOT")
rsZCDODOS0.CDODOSBIC = rsAdo("CDODOSBIC")
rsZCDODOS0.CDODOSCOT = rsAdo("CDODOSCOT")
rsZCDODOS0.CDODOSCOR = rsAdo("CDODOSCOR")
rsZCDODOS0.CDODOSPRT = rsAdo("CDODOSPRT")
rsZCDODOS0.CDODOSPRR = rsAdo("CDODOSPRR")
rsZCDODOS0.CDODOSUTV = rsAdo("CDODOSUTV")
rsZCDODOS0.CDODOSPAT = rsAdo("CDODOSPAT")
rsZCDODOS0.CDODOSPAR = rsAdo("CDODOSPAR")
rsZCDODOS0.CDODOSPAV = rsAdo("CDODOSPAV")
rsZCDODOS0.CDODOSOUV = rsAdo("CDODOSOUV")
rsZCDODOS0.CDODOSEMI = rsAdo("CDODOSEMI")
rsZCDODOS0.CDODOSVAL = rsAdo("CDODOSVAL")
rsZCDODOS0.CDODOSDEP = rsAdo("CDODOSDEP")
rsZCDODOS0.CDODOSDTR = rsAdo("CDODOSDTR")
rsZCDODOS0.CDODOSVCP = rsAdo("CDODOSVCP")
rsZCDODOS0.CDODOSCLO = rsAdo("CDODOSCLO")
rsZCDODOS0.CDODOSREJ = rsAdo("CDODOSREJ")
rsZCDODOS0.CDODOSOBJ = rsAdo("CDODOSOBJ")
rsZCDODOS0.CDODOSAVU = rsAdo("CDODOSAVU")
rsZCDODOS0.CDODOSMOV = rsAdo("CDODOSMOV")
rsZCDODOS0.CDODOSCAC = rsAdo("CDODOSCAC")
rsZCDODOS0.CDODOSMCA = rsAdo("CDODOSMCA")
rsZCDODOS0.CDODOSDIF = rsAdo("CDODOSDIF")
rsZCDODOS0.CDODOSMDI = rsAdo("CDODOSMDI")
rsZCDODOS0.CDODOSPMO = rsAdo("CDODOSPMO")
rsZCDODOS0.CDODOSPCD = rsAdo("CDODOSPCD")
rsZCDODOS0.CDODOSPCC = rsAdo("CDODOSPCC")
rsZCDODOS0.CDODOSPDE = rsAdo("CDODOSPDE")
rsZCDODOS0.CDODOSPPO = rsAdo("CDODOSPPO")
rsZCDODOS0.CDODOSAUT = rsAdo("CDODOSAUT")
rsZCDODOS0.CDODOSREG = rsAdo("CDODOSREG")
rsZCDODOS0.CDODOSENC = rsAdo("CDODOSENC")
rsZCDODOS0.CDODOSDAN = rsAdo("CDODOSDAN")
rsZCDODOS0.CDODOSANN = rsAdo("CDODOSANN")
rsZCDODOS0.CDODOSPCO = rsAdo("CDODOSPCO")
rsZCDODOS0.CDODOSLEM = rsAdo("CDODOSLEM")
rsZCDODOS0.CDODOSLDE = rsAdo("CDODOSLDE")
rsZCDODOS0.CDODOSDLE = rsAdo("CDODOSDLE")
rsZCDODOS0.CDODOSEPA = rsAdo("CDODOSEPA")
rsZCDODOS0.CDODOSTRA = rsAdo("CDODOSTRA")
rsZCDODOS0.CDODOSFCD = rsAdo("CDODOSFCD")
rsZCDODOS0.CDODOSCUS = rsAdo("CDODOSCUS")
rsZCDODOS0.CDODOSCUV = rsAdo("CDODOSCUV")
rsZCDODOS0.CDODOSCU2 = rsAdo("CDODOSCU2")
rsZCDODOS0.CDODOSOPE = rsAdo("CDODOSOPE")
rsZCDODOS0.CDODOSPOO = rsAdo("CDODOSPOO")
rsZCDODOS0.CDODOSPBE = rsAdo("CDODOSPBE")
rsZCDODOS0.CDODOSGAG = rsAdo("CDODOSGAG")
rsZCDODOS0.CDODOSSTB = rsAdo("CDODOSSTB")
rsZCDODOS0.CDODOSMRE = rsAdo("CDODOSMRE")
rsZCDODOS0.CDODOSNPD = rsAdo("CDODOSNPD")
rsZCDODOS0.CDODOSTJD = rsAdo("CDODOSTJD")
rsZCDODOS0.CDODOSPDO = rsAdo("CDODOSPDO")
rsZCDODOS0.CDODOSGAR = rsAdo("CDODOSGAR")
rsZCDODOS0.CDODOSOBM = rsAdo("CDODOSOBM")
rsZCDODOS0.CDODOSTBR = rsAdo("CDODOSTBR")
rsZCDODOS0.CDODOSBRE = rsAdo("CDODOSBRE")
rsZCDODOS0.CDODOSBEC = rsAdo("CDODOSBEC")
rsZCDODOS0.CDODOSRNO = rsAdo("CDODOSRNO")
rsZCDODOS0.CDODOSDPA = rsAdo("CDODOSDPA")
rsZCDODOS0.CDODOSDVI = rsAdo("CDODOSDVI")
rsZCDODOS0.CDODOSEPY = rsAdo("CDODOSEPY")
rsZCDODOS0.CDODOSEVI = rsAdo("CDODOSEVI")
rsZCDODOS0.CDODOSVPA = rsAdo("CDODOSVPA")
rsZCDODOS0.CDODOSVVI = rsAdo("CDODOSVVI")
rsZCDODOS0.CDODOSNDE = rsAdo("CDODOSNDE")
rsZCDODOS0.CDODOSNAE = rsAdo("CDODOSNAE")
rsZCDODOS0.CDODOSEVE = rsAdo("CDODOSEVE")
rsZCDODOS0.CDODOSETA = rsAdo("CDODOSETA")
rsZCDODOS0.CDODOSDP2 = rsAdo("CDODOSDP2")
rsZCDODOS0.CDODOSEP2 = rsAdo("CDODOSEP2")
rsZCDODOS0.CDODOSPD2 = rsAdo("CDODOSPD2")
rsZCDODOS0.CDODOSAUN = rsAdo("CDODOSAUN")
rsZCDODOS0.CDODOSCER = rsAdo("CDODOSCER")
rsZCDODOS0.CDODOSCRE = rsAdo("CDODOSCRE")
rsZCDODOS0.CDODOSREM = rsAdo("CDODOSREM")
rsZCDODOS0.CDODOSRGR = rsAdo("CDODOSRGR")
rsZCDODOS0.CDODOSLED = rsAdo("CDODOSLED")
rsZCDODOS0.CDODOSLDA = rsAdo("CDODOSLDA")
Exit Function
Error_Handler:
rsZCDODOS0_GetBuffer = Error
End Function


Public Sub rsWCDOCOM0_Init(rsWCDOCOM0 As typeWCDOCOM0)
rsWCDOCOM0.WCDOCOMCOM = ""
rsWCDOCOM0.WCDOCOMMON = 0
rsWCDOCOM0.WCDOCOMDEV = ""
rsWCDOCOM0.WCDOCOMMTV = 0

rsWCDOCOM0.WCDOCO2TX1 = 0
rsWCDOCOM0.WCDOCO2PER = ""
rsWCDOCOM0.WCDOCO2MIN = 0

rsWCDOCOM0.WCDOTC2DEV = ""
rsWCDOCOM0.WCDOTC2MTF = 0

End Sub
