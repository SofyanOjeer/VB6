Attribute VB_Name = "rsZCDOMOD0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCDOMOD0
    CDOMODETB       As Integer                        ' CODE ETABLISSEMENT
    CDOMODAGE       As Integer                        ' AGENCE
    CDOMODSER       As String * 2                     ' SERVICE
    CDOMODSSE       As String * 2                     ' SOUS-SERVICE
    CDOMODCOP       As String * 3                     ' CODE OPERATION
    CDOMODDOS       As Long                           ' NUMERO DOSSIER
    CDOMODNMO       As Long                           ' N° MODIFICATION
    CDOMODNUR       As Long                           ' N° RENOUVELLEMENT
    CDOMODNAT       As String * 3                     ' NATURE
    CDOMODEXT       As String * 16                    ' REFERENCE EXTERNE
    CDOMODMON       As Currency                       ' MONTANT DOSSIER
    CDOMODDEV       As String * 3                     ' DEVISE
    CDOMODMOA       As Currency                       ' MONTANT ADDITIONNEL
    CDOMODMOT       As Currency                       ' MONTANT TOTAL
    CDOMODMOC       As Currency                       ' MONTANT CONFIRME
    CDOMODMOD       As Currency                       ' MONTANT DUCROIRE
    CDOMODCON       As String * 1                     ' CONFIRM NOTIFI PARTI
    CDOMODIRR       As String * 1                     ' IRREVOCABLE (O/N)
    CDOMODFRA       As String * 1                     ' FRACTIONNABLE (O/N)
    CDOMODREN       As String * 1                     ' RENOUVELABLE (O/N)
    CDOMODCUM       As String * 1                     ' CUMULATIF (O/N)
    CDOMODTRS       As String * 1                     ' TRANSFERABLE
    CDOMODTOL       As Currency                       ' TOLERANCE +
    CDOMODTO2       As Currency                       ' TOLERANCE -
    CDOMODDOR       As String * 1                     ' DONN. ORDRE CLI/TIE
    CDOMODDON       As String * 7                     ' DONNEUR ORDRE IMPORT
    CDOMODDOE       As String * 64                    ' DONNEUR ORDRE EXPORT
    CDOMODBER       As String * 1                     ' BENEFICIAIR CLI/TIE
    CDOMODBEN       As String * 7                     ' BENEFICIAIRE EXPORT
    CDOMODBEI       As String * 64                    ' BENEFICIAIRE IMPORT
    CDOMODBAR       As String * 1                     ' BANQU.BENEF.CLI/TIE
    CDOMODBAB       As String * 7                     ' BANQUE BENEF
    CDOMODNOR       As String * 1                     ' NOTIF/CONFI OU EMETT
    CDOMODNOT       As String * 7                     ' NOTIF/CONFI OU EMETT
    CDOMODBIC       As String * 12                    ' BIC SUPPLEMEN. EMETT
    CDOMODCOT       As String * 1                     ' CORRESPOND. CLI/TIE
    CDOMODCOR       As String * 7                     ' CORRESPONDANT
    CDOMODPRT       As String * 1                     ' LIEU PRES CLI/TIE
    CDOMODPRR       As String * 7                     ' LIEU PRESENTATION
    CDOMODUTV       As String * 32                    ' LIEU PRESENTATION
    CDOMODPAT       As String * 1                     ' LIEU PAIE CLI/TIE
    CDOMODPAR       As String * 7                     ' LIEU PAIEMENT
    CDOMODPAV       As String * 32                    ' LIEU PAIEMENT
    CDOMODOUV       As Long                           ' DATE OUVERTURE
    CDOMODEMI       As Long                           ' DATE EMISSION
    CDOMODVAL       As Long                           ' DATE VALIDITE
    CDOMODDEP       As Long                           ' DATE EXTREME PAYMT
    CDOMODDTR       As Long                           ' DATE DE TRANSFERT
    CDOMODVCP       As Long                           ' DATE VALID. COMPTA
    CDOMODCLO       As Long                           ' DATE CLOTURE
    CDOMODREJ       As String * 3                     ' MOTIF REJET (CLOTUR)
    CDOMODOBJ       As String * 6                     ' OBJET CREDIT
    CDOMODAVU       As Long                           ' % PAIEM. A VUE
    CDOMODMOV       As Currency                       ' MONTANT A VUE
    CDOMODCAC       As Long                           ' % PAIEM. CTR ACCEPT.
    CDOMODMCA       As Currency                       ' MONTANT CTR ACCEPT.
    CDOMODDIF       As Long                           ' % PAIEM. DIFFERE
    CDOMODMDI       As Currency                       ' MONTANT. DIFFERE
    CDOMODPMO       As Currency                       ' MONTANT PROVISIONNE
    CDOMODPCD       As String * 20                    ' PROV. DEBIT  COMPTE
    CDOMODPCC       As String * 20                    ' PROV. CREDIT COMPTE
    CDOMODPDE       As Currency                       ' PROVISION DEVISE DOS
    CDOMODPPO       As Long                           ' PROVISION POURCEN
    CDOMODAUT       As String * 12                    ' CODE AUTORISATION
    CDOMODREG       As Currency                       ' MONTANT PAYE
    CDOMODENC       As Currency                       ' MONTANT ENCAISSE
    CDOMODDAN       As Long                           ' DATE ANNULATION
    CDOMODANN       As Currency                       ' MONTANT ANNULE
    CDOMODPCO       As Double                         ' COURS DEVPRO/DEVDOS
    CDOMODLEM       As String * 30                    ' LIEU EMBARQUEMENT
    CDOMODLDE       As String * 30                    ' LIEU DESTINATION
    CDOMODDLE       As Long                           ' DATE LIMITE EMBARQU.
    CDOMODEPA       As String * 1                     ' EXPED.PARTIE.AUTORI
    CDOMODTRA       As String * 1                     ' TRANBORDEMENT AUTORI
    CDOMODFCD       As String * 1                     ' FRAI CHARGE D.O. BEN
    CDOMODCUS       As Integer                        ' UTILI. DE SAISIE
    CDOMODCUV       As Integer                        ' 1ER VALIDEUR
    CDOMODCU2       As Integer                        ' 2EME VALIDEUR
    CDOMODOPE       As String * 1                     ' OPERATIVITE DU CRED.
    CDOMODPOO       As String * 1                     ' EXISTENCE POOL
    CDOMODPBE       As Currency                       ' PART.BANQUE EXPORT
    CDOMODGAG       As String * 1                     ' GAGE MARCHANDISE
    CDOMODSTB       As String * 1                     ' STAND BY
    CDOMODMRE       As String * 3                     ' MODE DE REALISAT°
    CDOMODNPD       As Long                           ' NBJ PRES. DOCUMENT
    CDOMODTJD       As String * 1                     ' TY JOUR DOCS
    CDOMODPDO       As String * 60                    ' PER.PRE.DOCS.
    CDOMODGAR       As String * 64                    ' LIBELLE GARANTIE
    CDOMODOBM       As String * 64                    ' OBJET DE MODIF.
    CDOMODTBR       As String * 1                     ' TIERS BQ REMBOURS
    CDOMODBRE       As String * 7                     ' BQ REMBOURSEMENT
    CDOMODBEC       As String * 1                     ' BENEF.PAY.COMMIS°
    CDOMODRNO       As String * 16                    ' REF.NOTIFICATEUR
    CDOMODDPA       As String * 3                     ' DESTINATION PAYS
    CDOMODDVI       As String * 32                    ' DESTINATION VILLE
    CDOMODEPY       As String * 3                     ' EMBARQUEMENT PAYS
    CDOMODEVI       As String * 32                    ' EMBARQUEMENT VILLE
    CDOMODVPA       As String * 3                     ' VALIDITE PAYS
    CDOMODVVI       As String * 32                    ' VALIDIT VILLE
    CDOMODNDE       As Long                           ' DOSSIER EXPORT
    CDOMODNAE       As String * 3                     ' NATURE EXPORT
    CDOMODEVE       As String * 2                     ' EVENEMENT
    CDOMODETA       As String * 2                     ' ETAT DE LA MODIF
    CDOMODDMO       As Long                           ' DATE MODIFICATION
    CDOMODDRM       As Long                           ' DATE RECEPTION MOD
    CDOMODDP2       As String * 32                    ' DESTIN.PAYS LIBELLE
    CDOMODEP2       As String * 32                    ' EMBARQ.PAYS LIBELLE
    CDOMODPD2       As String * 80                    ' PER.PRES.DOC.SUITE
    CDOMODAUN       As String * 12                    ' CODE AUT. NOTIFIE
    CDOMODCER       As String * 1                     ' COTAT°(O=CERTAIN/N)
    CDOMODCRE       As String * 12                    ' CODE REGLE
    CDOMODREM       As String * 35                    ' LIBELLE EMISSION
    CDOMODRGR       As String * 1                     ' REGLE RBT
    CDOMODLED       As String * 65                    ' LIEU EMBARQMT
    CDOMODLDA       As String * 65                    ' LIEU DEBARQMT

End Type
Public Sub rsZCDOMOD0_Init(rsYCDOMOD0 As typeZCDOMOD0)
rsYCDOMOD0.CDOMODETB = 0
rsYCDOMOD0.CDOMODAGE = 0
rsYCDOMOD0.CDOMODSER = ""
rsYCDOMOD0.CDOMODSSE = ""
rsYCDOMOD0.CDOMODCOP = ""
rsYCDOMOD0.CDOMODDOS = 0
rsYCDOMOD0.CDOMODNMO = 0
rsYCDOMOD0.CDOMODNUR = 0
rsYCDOMOD0.CDOMODNAT = ""
rsYCDOMOD0.CDOMODEXT = ""
rsYCDOMOD0.CDOMODMON = 0
rsYCDOMOD0.CDOMODDEV = ""
rsYCDOMOD0.CDOMODMOA = 0
rsYCDOMOD0.CDOMODMOT = 0
rsYCDOMOD0.CDOMODMOC = 0
rsYCDOMOD0.CDOMODMOD = 0
rsYCDOMOD0.CDOMODCON = ""
rsYCDOMOD0.CDOMODIRR = ""
rsYCDOMOD0.CDOMODFRA = ""
rsYCDOMOD0.CDOMODREN = ""
rsYCDOMOD0.CDOMODCUM = ""
rsYCDOMOD0.CDOMODTRS = ""
rsYCDOMOD0.CDOMODTOL = 0
rsYCDOMOD0.CDOMODTO2 = 0
rsYCDOMOD0.CDOMODDOR = ""
rsYCDOMOD0.CDOMODDON = ""
rsYCDOMOD0.CDOMODDOE = ""
rsYCDOMOD0.CDOMODBER = ""
rsYCDOMOD0.CDOMODBEN = ""
rsYCDOMOD0.CDOMODBEI = ""
rsYCDOMOD0.CDOMODBAR = ""
rsYCDOMOD0.CDOMODBAB = ""
rsYCDOMOD0.CDOMODNOR = ""
rsYCDOMOD0.CDOMODNOT = ""
rsYCDOMOD0.CDOMODBIC = ""
rsYCDOMOD0.CDOMODCOT = ""
rsYCDOMOD0.CDOMODCOR = ""
rsYCDOMOD0.CDOMODPRT = ""
rsYCDOMOD0.CDOMODPRR = ""
rsYCDOMOD0.CDOMODUTV = ""
rsYCDOMOD0.CDOMODPAT = ""
rsYCDOMOD0.CDOMODPAR = ""
rsYCDOMOD0.CDOMODPAV = ""
rsYCDOMOD0.CDOMODOUV = 0
rsYCDOMOD0.CDOMODEMI = 0
rsYCDOMOD0.CDOMODVAL = 0
rsYCDOMOD0.CDOMODDEP = 0
rsYCDOMOD0.CDOMODDTR = 0
rsYCDOMOD0.CDOMODVCP = 0
rsYCDOMOD0.CDOMODCLO = 0
rsYCDOMOD0.CDOMODREJ = ""
rsYCDOMOD0.CDOMODOBJ = ""
rsYCDOMOD0.CDOMODAVU = 0
rsYCDOMOD0.CDOMODMOV = 0
rsYCDOMOD0.CDOMODCAC = 0
rsYCDOMOD0.CDOMODMCA = 0
rsYCDOMOD0.CDOMODDIF = 0
rsYCDOMOD0.CDOMODMDI = 0
rsYCDOMOD0.CDOMODPMO = 0
rsYCDOMOD0.CDOMODPCD = ""
rsYCDOMOD0.CDOMODPCC = ""
rsYCDOMOD0.CDOMODPDE = 0
rsYCDOMOD0.CDOMODPPO = 0
rsYCDOMOD0.CDOMODAUT = ""
rsYCDOMOD0.CDOMODREG = 0
rsYCDOMOD0.CDOMODENC = 0
rsYCDOMOD0.CDOMODDAN = 0
rsYCDOMOD0.CDOMODANN = 0
rsYCDOMOD0.CDOMODPCO = 0
rsYCDOMOD0.CDOMODLEM = ""
rsYCDOMOD0.CDOMODLDE = ""
rsYCDOMOD0.CDOMODDLE = 0
rsYCDOMOD0.CDOMODEPA = ""
rsYCDOMOD0.CDOMODTRA = ""
rsYCDOMOD0.CDOMODFCD = ""
rsYCDOMOD0.CDOMODCUS = 0
rsYCDOMOD0.CDOMODCUV = 0
rsYCDOMOD0.CDOMODCU2 = 0
rsYCDOMOD0.CDOMODOPE = ""
rsYCDOMOD0.CDOMODPOO = ""
rsYCDOMOD0.CDOMODPBE = 0
rsYCDOMOD0.CDOMODGAG = ""
rsYCDOMOD0.CDOMODSTB = ""
rsYCDOMOD0.CDOMODMRE = ""
rsYCDOMOD0.CDOMODNPD = 0
rsYCDOMOD0.CDOMODTJD = ""
rsYCDOMOD0.CDOMODPDO = ""
rsYCDOMOD0.CDOMODGAR = ""
rsYCDOMOD0.CDOMODOBM = ""
rsYCDOMOD0.CDOMODTBR = ""
rsYCDOMOD0.CDOMODBRE = ""
rsYCDOMOD0.CDOMODBEC = ""
rsYCDOMOD0.CDOMODRNO = ""
rsYCDOMOD0.CDOMODDPA = ""
rsYCDOMOD0.CDOMODDVI = ""
rsYCDOMOD0.CDOMODEPY = ""
rsYCDOMOD0.CDOMODEVI = ""
rsYCDOMOD0.CDOMODVPA = ""
rsYCDOMOD0.CDOMODVVI = ""
rsYCDOMOD0.CDOMODNDE = 0
rsYCDOMOD0.CDOMODNAE = ""
rsYCDOMOD0.CDOMODEVE = ""
rsYCDOMOD0.CDOMODETA = ""
rsYCDOMOD0.CDOMODDMO = 0
rsYCDOMOD0.CDOMODDRM = 0
rsYCDOMOD0.CDOMODDP2 = ""
rsYCDOMOD0.CDOMODEP2 = ""
rsYCDOMOD0.CDOMODPD2 = ""
rsYCDOMOD0.CDOMODAUN = ""
rsYCDOMOD0.CDOMODCER = ""
rsYCDOMOD0.CDOMODCRE = ""
rsYCDOMOD0.CDOMODREM = ""
rsYCDOMOD0.CDOMODRGR = ""
rsYCDOMOD0.CDOMODLED = ""
rsYCDOMOD0.CDOMODLDA = ""
End Sub
Public Function rsZCDOMOD0_GetBuffer(rsAdo As ADODB.Recordset, rsZCDOMOD0 As typeZCDOMOD0)
On Error GoTo Error_Handler
rsZCDOMOD0_GetBuffer = Null
rsZCDOMOD0.CDOMODETB = rsAdo("CDOMODETB")
rsZCDOMOD0.CDOMODAGE = rsAdo("CDOMODAGE")
rsZCDOMOD0.CDOMODSER = rsAdo("CDOMODSER")
rsZCDOMOD0.CDOMODSSE = rsAdo("CDOMODSSE")
rsZCDOMOD0.CDOMODCOP = rsAdo("CDOMODCOP")
rsZCDOMOD0.CDOMODDOS = rsAdo("CDOMODDOS")
rsZCDOMOD0.CDOMODNMO = rsAdo("CDOMODNMO")
rsZCDOMOD0.CDOMODNUR = rsAdo("CDOMODNUR")
rsZCDOMOD0.CDOMODNAT = rsAdo("CDOMODNAT")
rsZCDOMOD0.CDOMODEXT = rsAdo("CDOMODEXT")
rsZCDOMOD0.CDOMODMON = rsAdo("CDOMODMON")
rsZCDOMOD0.CDOMODDEV = rsAdo("CDOMODDEV")
rsZCDOMOD0.CDOMODMOA = rsAdo("CDOMODMOA")
rsZCDOMOD0.CDOMODMOT = rsAdo("CDOMODMOT")
rsZCDOMOD0.CDOMODMOC = rsAdo("CDOMODMOC")
rsZCDOMOD0.CDOMODMOD = rsAdo("CDOMODMOD")
rsZCDOMOD0.CDOMODCON = rsAdo("CDOMODCON")
rsZCDOMOD0.CDOMODIRR = rsAdo("CDOMODIRR")
rsZCDOMOD0.CDOMODFRA = rsAdo("CDOMODFRA")
rsZCDOMOD0.CDOMODREN = rsAdo("CDOMODREN")
rsZCDOMOD0.CDOMODCUM = rsAdo("CDOMODCUM")
rsZCDOMOD0.CDOMODTRS = rsAdo("CDOMODTRS")
rsZCDOMOD0.CDOMODTOL = rsAdo("CDOMODTOL")
rsZCDOMOD0.CDOMODTO2 = rsAdo("CDOMODTO2")
rsZCDOMOD0.CDOMODDOR = rsAdo("CDOMODDOR")
rsZCDOMOD0.CDOMODDON = rsAdo("CDOMODDON")
rsZCDOMOD0.CDOMODDOE = rsAdo("CDOMODDOE")
rsZCDOMOD0.CDOMODBER = rsAdo("CDOMODBER")
rsZCDOMOD0.CDOMODBEN = rsAdo("CDOMODBEN")
rsZCDOMOD0.CDOMODBEI = rsAdo("CDOMODBEI")
rsZCDOMOD0.CDOMODBAR = rsAdo("CDOMODBAR")
rsZCDOMOD0.CDOMODBAB = rsAdo("CDOMODBAB")
rsZCDOMOD0.CDOMODNOR = rsAdo("CDOMODNOR")
rsZCDOMOD0.CDOMODNOT = rsAdo("CDOMODNOT")
rsZCDOMOD0.CDOMODBIC = rsAdo("CDOMODBIC")
rsZCDOMOD0.CDOMODCOT = rsAdo("CDOMODCOT")
rsZCDOMOD0.CDOMODCOR = rsAdo("CDOMODCOR")
rsZCDOMOD0.CDOMODPRT = rsAdo("CDOMODPRT")
rsZCDOMOD0.CDOMODPRR = rsAdo("CDOMODPRR")
rsZCDOMOD0.CDOMODUTV = rsAdo("CDOMODUTV")
rsZCDOMOD0.CDOMODPAT = rsAdo("CDOMODPAT")
rsZCDOMOD0.CDOMODPAR = rsAdo("CDOMODPAR")
rsZCDOMOD0.CDOMODPAV = rsAdo("CDOMODPAV")
rsZCDOMOD0.CDOMODOUV = rsAdo("CDOMODOUV")
rsZCDOMOD0.CDOMODEMI = rsAdo("CDOMODEMI")
rsZCDOMOD0.CDOMODVAL = rsAdo("CDOMODVAL")
rsZCDOMOD0.CDOMODDEP = rsAdo("CDOMODDEP")
rsZCDOMOD0.CDOMODDTR = rsAdo("CDOMODDTR")
rsZCDOMOD0.CDOMODVCP = rsAdo("CDOMODVCP")
rsZCDOMOD0.CDOMODCLO = rsAdo("CDOMODCLO")
rsZCDOMOD0.CDOMODREJ = rsAdo("CDOMODREJ")
rsZCDOMOD0.CDOMODOBJ = rsAdo("CDOMODOBJ")
rsZCDOMOD0.CDOMODAVU = rsAdo("CDOMODAVU")
rsZCDOMOD0.CDOMODMOV = rsAdo("CDOMODMOV")
rsZCDOMOD0.CDOMODCAC = rsAdo("CDOMODCAC")
rsZCDOMOD0.CDOMODMCA = rsAdo("CDOMODMCA")
rsZCDOMOD0.CDOMODDIF = rsAdo("CDOMODDIF")
rsZCDOMOD0.CDOMODMDI = rsAdo("CDOMODMDI")
rsZCDOMOD0.CDOMODPMO = rsAdo("CDOMODPMO")
rsZCDOMOD0.CDOMODPCD = rsAdo("CDOMODPCD")
rsZCDOMOD0.CDOMODPCC = rsAdo("CDOMODPCC")
rsZCDOMOD0.CDOMODPDE = rsAdo("CDOMODPDE")
rsZCDOMOD0.CDOMODPPO = rsAdo("CDOMODPPO")
rsZCDOMOD0.CDOMODAUT = rsAdo("CDOMODAUT")
rsZCDOMOD0.CDOMODREG = rsAdo("CDOMODREG")
rsZCDOMOD0.CDOMODENC = rsAdo("CDOMODENC")
rsZCDOMOD0.CDOMODDAN = rsAdo("CDOMODDAN")
rsZCDOMOD0.CDOMODANN = rsAdo("CDOMODANN")
rsZCDOMOD0.CDOMODPCO = rsAdo("CDOMODPCO")
rsZCDOMOD0.CDOMODLEM = rsAdo("CDOMODLEM")
rsZCDOMOD0.CDOMODLDE = rsAdo("CDOMODLDE")
rsZCDOMOD0.CDOMODDLE = rsAdo("CDOMODDLE")
rsZCDOMOD0.CDOMODEPA = rsAdo("CDOMODEPA")
rsZCDOMOD0.CDOMODTRA = rsAdo("CDOMODTRA")
rsZCDOMOD0.CDOMODFCD = rsAdo("CDOMODFCD")
rsZCDOMOD0.CDOMODCUS = rsAdo("CDOMODCUS")
rsZCDOMOD0.CDOMODCUV = rsAdo("CDOMODCUV")
rsZCDOMOD0.CDOMODCU2 = rsAdo("CDOMODCU2")
rsZCDOMOD0.CDOMODOPE = rsAdo("CDOMODOPE")
rsZCDOMOD0.CDOMODPOO = rsAdo("CDOMODPOO")
rsZCDOMOD0.CDOMODPBE = rsAdo("CDOMODPBE")
rsZCDOMOD0.CDOMODGAG = rsAdo("CDOMODGAG")
rsZCDOMOD0.CDOMODSTB = rsAdo("CDOMODSTB")
rsZCDOMOD0.CDOMODMRE = rsAdo("CDOMODMRE")
rsZCDOMOD0.CDOMODNPD = rsAdo("CDOMODNPD")
rsZCDOMOD0.CDOMODTJD = rsAdo("CDOMODTJD")
rsZCDOMOD0.CDOMODPDO = rsAdo("CDOMODPDO")
rsZCDOMOD0.CDOMODGAR = rsAdo("CDOMODGAR")
rsZCDOMOD0.CDOMODOBM = rsAdo("CDOMODOBM")
rsZCDOMOD0.CDOMODTBR = rsAdo("CDOMODTBR")
rsZCDOMOD0.CDOMODBRE = rsAdo("CDOMODBRE")
rsZCDOMOD0.CDOMODBEC = rsAdo("CDOMODBEC")
rsZCDOMOD0.CDOMODRNO = rsAdo("CDOMODRNO")
rsZCDOMOD0.CDOMODDPA = rsAdo("CDOMODDPA")
rsZCDOMOD0.CDOMODDVI = rsAdo("CDOMODDVI")
rsZCDOMOD0.CDOMODEPY = rsAdo("CDOMODEPY")
rsZCDOMOD0.CDOMODEVI = rsAdo("CDOMODEVI")
rsZCDOMOD0.CDOMODVPA = rsAdo("CDOMODVPA")
rsZCDOMOD0.CDOMODVVI = rsAdo("CDOMODVVI")
rsZCDOMOD0.CDOMODNDE = rsAdo("CDOMODNDE")
rsZCDOMOD0.CDOMODNAE = rsAdo("CDOMODNAE")
rsZCDOMOD0.CDOMODEVE = rsAdo("CDOMODEVE")
rsZCDOMOD0.CDOMODETA = rsAdo("CDOMODETA")
rsZCDOMOD0.CDOMODDMO = rsAdo("CDOMODDMO")
rsZCDOMOD0.CDOMODDRM = rsAdo("CDOMODDRM")
rsZCDOMOD0.CDOMODDP2 = rsAdo("CDOMODDP2")
rsZCDOMOD0.CDOMODEP2 = rsAdo("CDOMODEP2")
rsZCDOMOD0.CDOMODPD2 = rsAdo("CDOMODPD2")
rsZCDOMOD0.CDOMODAUN = rsAdo("CDOMODAUN")
rsZCDOMOD0.CDOMODCER = rsAdo("CDOMODCER")
rsZCDOMOD0.CDOMODCRE = rsAdo("CDOMODCRE")
rsZCDOMOD0.CDOMODREM = rsAdo("CDOMODREM")
rsZCDOMOD0.CDOMODRGR = rsAdo("CDOMODRGR")
rsZCDOMOD0.CDOMODLED = rsAdo("CDOMODLED")
rsZCDOMOD0.CDOMODLDA = rsAdo("CDOMODLDA")
Exit Function
Error_Handler:
rsZCDOMOD0_GetBuffer = Error
End Function

