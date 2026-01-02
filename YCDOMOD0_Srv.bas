Attribute VB_Name = "srvYCDOMOD0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCDOMOD0Len = 1135 ' 34 + 1101
Public Const recYCDOMOD0_Block = 20
Public Const constYCDOMOD0 = "YCDOMOD0"
Dim meYbase As typeYBase
Dim paramYCDOMOD0_Import As String

Type typeYCDOMOD0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

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
End Type
    
'---------------------------------------------------------
Public Function srvYCDOMOD0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCDOMOD0 As typeYCDOMOD0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCDOMOD0_GetBuffer_ODBC = Null

    recYCDOMOD0.CDOMODETB = rsADO("CDOMODETB")   'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOMOD0.CDOMODAGE = rsADO("CDOMODAGE")   'CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOMOD0.CDOMODSER = rsADO("CDOMODSER")   'mId$(MsgTxt, K + 11, 2)
    recYCDOMOD0.CDOMODSSE = rsADO("CDOMODSSE")   'mId$(MsgTxt, K + 13, 2)
    recYCDOMOD0.CDOMODCOP = rsADO("CDOMODCOP")   'mId$(MsgTxt, K + 15, 3)
    recYCDOMOD0.CDOMODDOS = rsADO("CDOMODDOS")   'CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOMOD0.CDOMODNMO = rsADO("CDOMODNMO")   'CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOMOD0.CDOMODNUR = rsADO("CDOMODNUR")   'CLng(Val(mId$(MsgTxt, K + 32, 4)))
    recYCDOMOD0.CDOMODNAT = rsADO("CDOMODNAT")   'mId$(MsgTxt, K + 36, 3)
    recYCDOMOD0.CDOMODEXT = rsADO("CDOMODEXT")   'mId$(MsgTxt, K + 39, 16)
    recYCDOMOD0.CDOMODMON = rsADO("CDOMODMON")   'CCur(Val(mId$(MsgTxt, K + 55, 16))) / 100
    recYCDOMOD0.CDOMODDEV = rsADO("CDOMODDEV")   'mId$(MsgTxt, K + 71, 3)
    recYCDOMOD0.CDOMODMOA = rsADO("CDOMODMOA")   'CCur(Val(mId$(MsgTxt, K + 74, 16))) / 100
    recYCDOMOD0.CDOMODMOT = rsADO("CDOMODMOT")   'CCur(Val(mId$(MsgTxt, K + 90, 16))) / 100
    recYCDOMOD0.CDOMODMOC = rsADO("CDOMODMOC")   'CCur(Val(mId$(MsgTxt, K + 106, 16))) / 100
    recYCDOMOD0.CDOMODMOD = rsADO("CDOMODMOD")   'CCur(Val(mId$(MsgTxt, K + 122, 16))) / 100
    recYCDOMOD0.CDOMODCON = rsADO("CDOMODCON")   'mId$(MsgTxt, K + 138, 1)
    recYCDOMOD0.CDOMODIRR = rsADO("CDOMODIRR")   'mId$(MsgTxt, K + 139, 1)
    recYCDOMOD0.CDOMODFRA = rsADO("CDOMODFRA")   'mId$(MsgTxt, K + 140, 1)
    recYCDOMOD0.CDOMODREN = rsADO("CDOMODREN")   'mId$(MsgTxt, K + 141, 1)
    recYCDOMOD0.CDOMODCUM = rsADO("CDOMODCUM")   'mId$(MsgTxt, K + 142, 1)
    recYCDOMOD0.CDOMODTRS = rsADO("CDOMODTRS")   'mId$(MsgTxt, K + 143, 1)
    recYCDOMOD0.CDOMODTOL = rsADO("CDOMODTOL")   'CCur(Val(mId$(MsgTxt, K + 144, 4))) / 100
    recYCDOMOD0.CDOMODTO2 = rsADO("CDOMODTO2")   'CCur(Val(mId$(MsgTxt, K + 148, 4))) / 100
    recYCDOMOD0.CDOMODDOR = rsADO("CDOMODDOR")   'mId$(MsgTxt, K + 152, 1)
    recYCDOMOD0.CDOMODDON = rsADO("CDOMODDON")   'mId$(MsgTxt, K + 153, 7)
    recYCDOMOD0.CDOMODDOE = rsADO("CDOMODDOE")   'mId$(MsgTxt, K + 160, 64)
    recYCDOMOD0.CDOMODBER = rsADO("CDOMODBER")   'mId$(MsgTxt, K + 224, 1)
    recYCDOMOD0.CDOMODBEN = rsADO("CDOMODBEN")   'mId$(MsgTxt, K + 225, 7)
    recYCDOMOD0.CDOMODBEI = rsADO("CDOMODBEI")   'mId$(MsgTxt, K + 232, 64)
    recYCDOMOD0.CDOMODBAR = rsADO("CDOMODBAR")   'mId$(MsgTxt, K + 296, 1)
    recYCDOMOD0.CDOMODBAB = rsADO("CDOMODBAB")   'mId$(MsgTxt, K + 297, 7)
    recYCDOMOD0.CDOMODNOR = rsADO("CDOMODNOR")   'mId$(MsgTxt, K + 304, 1)
    recYCDOMOD0.CDOMODNOT = rsADO("CDOMODNOT")   'mId$(MsgTxt, K + 305, 7)
    recYCDOMOD0.CDOMODBIC = rsADO("CDOMODBIC")   'mId$(MsgTxt, K + 312, 12)
    recYCDOMOD0.CDOMODCOT = rsADO("CDOMODCOT")   'mId$(MsgTxt, K + 324, 1)
    recYCDOMOD0.CDOMODCOR = rsADO("CDOMODCOR")   'mId$(MsgTxt, K + 325, 7)
    recYCDOMOD0.CDOMODPRT = rsADO("CDOMODPRT")   'mId$(MsgTxt, K + 332, 1)
    recYCDOMOD0.CDOMODPRR = rsADO("CDOMODPRR")   'mId$(MsgTxt, K + 333, 7)
    recYCDOMOD0.CDOMODUTV = rsADO("CDOMODUTV")   'mId$(MsgTxt, K + 340, 32)
    recYCDOMOD0.CDOMODPAT = rsADO("CDOMODPAT")   'mId$(MsgTxt, K + 372, 1)
    recYCDOMOD0.CDOMODPAR = rsADO("CDOMODPAR")   'mId$(MsgTxt, K + 373, 7)
    recYCDOMOD0.CDOMODPAV = rsADO("CDOMODPAV")   'mId$(MsgTxt, K + 380, 32)
    recYCDOMOD0.CDOMODOUV = rsADO("CDOMODOUV")   'CLng(Val(mId$(MsgTxt, K + 412, 8)))
    recYCDOMOD0.CDOMODEMI = rsADO("CDOMODEMI")   'CLng(Val(mId$(MsgTxt, K + 420, 8)))
    recYCDOMOD0.CDOMODVAL = rsADO("CDOMODVAL")   'CLng(Val(mId$(MsgTxt, K + 428, 8)))
    recYCDOMOD0.CDOMODDEP = rsADO("CDOMODDEP")   'CLng(Val(mId$(MsgTxt, K + 436, 8)))
    recYCDOMOD0.CDOMODDTR = rsADO("CDOMODDTR")   'CLng(Val(mId$(MsgTxt, K + 444, 8)))
    recYCDOMOD0.CDOMODVCP = rsADO("CDOMODVCP")   'CLng(Val(mId$(MsgTxt, K + 452, 8)))
    recYCDOMOD0.CDOMODCLO = rsADO("CDOMODCLO")   'CLng(Val(mId$(MsgTxt, K + 460, 8)))
    recYCDOMOD0.CDOMODREJ = rsADO("CDOMODREJ")   'mId$(MsgTxt, K + 468, 3)
    recYCDOMOD0.CDOMODOBJ = rsADO("CDOMODOBJ")   'mId$(MsgTxt, K + 471, 6)
    recYCDOMOD0.CDOMODAVU = rsADO("CDOMODAVU")   'CLng(Val(mId$(MsgTxt, K + 477, 4)))
    recYCDOMOD0.CDOMODMOV = rsADO("CDOMODMOV")   'CCur(Val(mId$(MsgTxt, K + 481, 16))) / 100
    recYCDOMOD0.CDOMODCAC = rsADO("CDOMODCAC")   'CLng(Val(mId$(MsgTxt, K + 497, 4)))
    recYCDOMOD0.CDOMODMCA = rsADO("CDOMODMCA")   'CCur(Val(mId$(MsgTxt, K + 501, 16))) / 100
    recYCDOMOD0.CDOMODDIF = rsADO("CDOMODDIF")   'CLng(Val(mId$(MsgTxt, K + 517, 4)))
    recYCDOMOD0.CDOMODMDI = rsADO("CDOMODMDI")   'CCur(Val(mId$(MsgTxt, K + 521, 16))) / 100
    recYCDOMOD0.CDOMODPMO = rsADO("CDOMODPMO")   'CCur(Val(mId$(MsgTxt, K + 537, 16))) / 100
    recYCDOMOD0.CDOMODPCD = rsADO("CDOMODPCD")   'mId$(MsgTxt, K + 553, 20)
    recYCDOMOD0.CDOMODPCC = rsADO("CDOMODPCC")   'mId$(MsgTxt, K + 573, 20)
    recYCDOMOD0.CDOMODPDE = rsADO("CDOMODPDE")   'CCur(Val(mId$(MsgTxt, K + 593, 16))) / 100
    recYCDOMOD0.CDOMODPPO = rsADO("CDOMODPPO")   'CLng(Val(mId$(MsgTxt, K + 609, 4)))
    recYCDOMOD0.CDOMODAUT = rsADO("CDOMODAUT")   'mId$(MsgTxt, K + 613, 12)
    recYCDOMOD0.CDOMODREG = rsADO("CDOMODREG")   'CCur(Val(mId$(MsgTxt, K + 625, 16))) / 100
    recYCDOMOD0.CDOMODENC = rsADO("CDOMODENC")   'CCur(Val(mId$(MsgTxt, K + 641, 16))) / 100
    recYCDOMOD0.CDOMODDAN = rsADO("CDOMODDAN")   'CLng(Val(mId$(MsgTxt, K + 657, 8)))
    recYCDOMOD0.CDOMODANN = rsADO("CDOMODANN")   'CCur(Val(mId$(MsgTxt, K + 665, 16))) / 100
    recYCDOMOD0.CDOMODPCO = rsADO("CDOMODPCO")   'CDbl(Val(mId$(MsgTxt, K + 681, 15))) / 1000000000
    recYCDOMOD0.CDOMODLEM = rsADO("CDOMODLEM")   'mId$(MsgTxt, K + 696, 30)
    recYCDOMOD0.CDOMODLDE = rsADO("CDOMODLDE")   'mId$(MsgTxt, K + 726, 30)
    recYCDOMOD0.CDOMODDLE = rsADO("CDOMODDLE")   'CLng(Val(mId$(MsgTxt, K + 756, 8)))
    recYCDOMOD0.CDOMODEPA = rsADO("CDOMODEPA")   'mId$(MsgTxt, K + 764, 1)
    recYCDOMOD0.CDOMODTRA = rsADO("CDOMODTRA")   'mId$(MsgTxt, K + 765, 1)
    recYCDOMOD0.CDOMODFCD = rsADO("CDOMODFCD")   'mId$(MsgTxt, K + 766, 1)
    recYCDOMOD0.CDOMODCUS = rsADO("CDOMODCUS")   'CInt(Val(mId$(MsgTxt, K + 767, 5)))
    recYCDOMOD0.CDOMODCUV = rsADO("CDOMODCUV")   'CInt(Val(mId$(MsgTxt, K + 772, 5)))
    recYCDOMOD0.CDOMODCU2 = rsADO("CDOMODCU2")   'CInt(Val(mId$(MsgTxt, K + 777, 5)))
    recYCDOMOD0.CDOMODOPE = rsADO("CDOMODOPE")   'mId$(MsgTxt, K + 782, 1)
    recYCDOMOD0.CDOMODPOO = rsADO("CDOMODPOO")   'mId$(MsgTxt, K + 783, 1)
    recYCDOMOD0.CDOMODPBE = rsADO("CDOMODPBE")   'CCur(Val(mId$(MsgTxt, K + 784, 16))) / 100
    recYCDOMOD0.CDOMODGAG = rsADO("CDOMODGAG")   'mId$(MsgTxt, K + 800, 1)
    recYCDOMOD0.CDOMODSTB = rsADO("CDOMODSTB")   'mId$(MsgTxt, K + 801, 1)
    recYCDOMOD0.CDOMODMRE = rsADO("CDOMODMRE")   'mId$(MsgTxt, K + 802, 3)
    recYCDOMOD0.CDOMODNPD = rsADO("CDOMODNPD")   'CLng(Val(mId$(MsgTxt, K + 805, 4)))
    recYCDOMOD0.CDOMODTJD = rsADO("CDOMODTJD")   'mId$(MsgTxt, K + 809, 1)
    recYCDOMOD0.CDOMODPDO = rsADO("CDOMODPDO")   'mId$(MsgTxt, K + 810, 60)
    recYCDOMOD0.CDOMODGAR = rsADO("CDOMODGAR")   'mId$(MsgTxt, K + 870, 64)
    recYCDOMOD0.CDOMODOBM = rsADO("CDOMODOBM")   'mId$(MsgTxt, K + 934, 64)
    recYCDOMOD0.CDOMODTBR = rsADO("CDOMODTBR")   'mId$(MsgTxt, K + 998, 1)
    recYCDOMOD0.CDOMODBRE = rsADO("CDOMODBRE")   'mId$(MsgTxt, K + 999, 7)
    recYCDOMOD0.CDOMODBEC = rsADO("CDOMODBEC")   'mId$(MsgTxt, K + 1006, 1)
    recYCDOMOD0.CDOMODRNO = rsADO("CDOMODRNO")   'mId$(MsgTxt, K + 1007, 16)
    recYCDOMOD0.CDOMODDPA = rsADO("CDOMODDPA")   'mId$(MsgTxt, K + 1023, 3)
    recYCDOMOD0.CDOMODDVI = rsADO("CDOMODDVI")   'mId$(MsgTxt, K + 1026, 32)
    recYCDOMOD0.CDOMODEPY = rsADO("CDOMODEPY")   'mId$(MsgTxt, K + 1058, 3)
    recYCDOMOD0.CDOMODEVI = rsADO("CDOMODEVI")   'mId$(MsgTxt, K + 1061, 32)
    recYCDOMOD0.CDOMODVPA = rsADO("CDOMODVPA")   'mId$(MsgTxt, K + 1093, 3)
    recYCDOMOD0.CDOMODVVI = rsADO("CDOMODVVI")   'mId$(MsgTxt, K + 1096, 32)
    recYCDOMOD0.CDOMODNDE = rsADO("CDOMODNDE")   'CLng(Val(mId$(MsgTxt, K + 1128, 10)))
    recYCDOMOD0.CDOMODNAE = rsADO("CDOMODNAE")   'mId$(MsgTxt, K + 1138, 3)
    recYCDOMOD0.CDOMODEVE = rsADO("CDOMODEVE")   'mId$(MsgTxt, K + 1141, 2)
    recYCDOMOD0.CDOMODETA = rsADO("CDOMODETA")   'mId$(MsgTxt, K + 1143, 2)
    recYCDOMOD0.CDOMODDMO = rsADO("CDOMODDMO")   'CLng(Val(mId$(MsgTxt, K + 1145, 8)))
    recYCDOMOD0.CDOMODDRM = rsADO("CDOMODDRM")   'CLng(Val(mId$(MsgTxt, K + 1153, 8)))
    recYCDOMOD0.CDOMODDP2 = rsADO("CDOMODDP2")   'mId$(MsgTxt, K + 1161, 32)
    recYCDOMOD0.CDOMODEP2 = rsADO("CDOMODEP2")   'mId$(MsgTxt, K + 1193, 32)
    recYCDOMOD0.CDOMODPD2 = rsADO("CDOMODPD2")   'mId$(MsgTxt, K + 1225, 80)
    recYCDOMOD0.CDOMODAUN = rsADO("CDOMODAUN")   'mId$(MsgTxt, K + 1305, 12)
    recYCDOMOD0.CDOMODCER = rsADO("CDOMODCER")   'mId$(MsgTxt, K + 1317, 1)


Exit Function

Error_Handler:
srvYCDOMOD0_GetBuffer_ODBC = Error

End Function

Public Function srvYCDOMOD0_Import(lX As String)
Dim xIn As String, x As String, Nb As Long
On Error GoTo Error_Handle

recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOMOD0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    srvYCDOMOD0_Import = Null
    lX = CStr(meYbase.Text)
    Exit Function
End If


srvYCDOMOD0_Import = "?"

paramYCDOMOD0_Import = paramYBase_DataF & Trim(constYCDOMOD0) & paramYBase_Data_ExtensionP

Open Trim(paramYCDOMOD0_Import) For Input As #1

Nb = 0
x = "delete * from YBase where Id = " & Chr$(34) & Trim(constYCDOMOD0) & Chr$(34)
MDB.Execute x

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYCDOMOD0
            meYbase.K1 = mId$(xIn, 15, 17) 'recYCDOMOD0.CDODOSCOP & recYCDOMOD0.CDODOSDOS & recYCDOMOD0.CDODOSNMO
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop


Close
srvYCDOMOD0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = YBIATAB0_DATE_CPT_J & "_" & constYCDOMOD0
lX = DSys & "_" & time_Hms & "_" & Nb
meYbase.Text = lX
dbYBase_Update meYbase

Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOMOD0_Import" & xIn, vbCritical, Error
Close

srvYCDOMOD0_Import = Error
End Function

Public Function srvYCDOMOD0_Import_Read(lId As String, lYCDOMOD0 As typeYCDOMOD0)

Dim xIn As String, x As String

On Error GoTo Error_Handle

srvYCDOMOD0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYCDOMOD0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYCDOMOD0_GetBuffer lYCDOMOD0
    srvYCDOMOD0_Import_Read = Null
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYCDOMOD0_Import_Read" & xIn, vbCritical, Error
srvYCDOMOD0_Import_Read = Error
End Function





Public Sub srvYCDOMOD0_Load(lYCDOMOD0() As typeYCDOMOD0, lYCDOMOD0_Nb As Integer)
Dim mMethod As String, blnYCDOMOD0_Suite
Dim wNbMax As Integer
Dim wYCDOMOD0 As typeYCDOMOD0

mMethod = Trim(lYCDOMOD0(0).Method) & "+"
blnYCDOMOD0_Suite = True: lYCDOMOD0_Nb = 0
wNbMax = recYCDOMOD0_Block + 2: ReDim Preserve lYCDOMOD0(wNbMax)

wYCDOMOD0 = lYCDOMOD0(1)
Do Until Not blnYCDOMOD0_Suite
    MsgTxtLen = 0
    Call srvYCDOMOD0_PutBuffer(wYCDOMOD0)
    Call srvYCDOMOD0_PutBuffer(lYCDOMOD0(0))
    If IsNull(SndRcv()) Then
        MsgTxtIndex = 0
        Do While MsgTxtIndex < MsgTxtLen
            If IsNull(srvYCDOMOD0_GetBuffer(wYCDOMOD0)) Then
            
                lYCDOMOD0_Nb = lYCDOMOD0_Nb + 1
                If lYCDOMOD0_Nb > wNbMax Then
                    wNbMax = wNbMax + recYCDOMOD0_Block
                    ReDim Preserve lYCDOMOD0(wNbMax)
                End If
            
                lYCDOMOD0(lYCDOMOD0_Nb) = wYCDOMOD0
                blnYCDOMOD0_Suite = True
            Else
                blnYCDOMOD0_Suite = False
                Exit Do
            End If
        Loop
    End If

    lYCDOMOD0(0).Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Public Function srvYCDOMOD0_Monitor(recYCDOMOD0 As typeYCDOMOD0)
'-----------------------------------------------------

Select Case mId$(Trim(recYCDOMOD0.Method), 1, 4)
    Case "Seek"
                srvYCDOMOD0_Monitor = srvYCDOMOD0_Seek(recYCDOMOD0)
    Case Else
                recYCDOMOD0.Err = recYCDOMOD0.Method
                Call srvYCDOMOD0_Error(recYCDOMOD0)
                srvYCDOMOD0_Monitor = recYCDOMOD0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYCDOMOD0_Error(recYCDOMOD0 As typeYCDOMOD0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YCDOMOD0" & Chr$(10) & Chr$(13)

Select Case mId$(recYCDOMOD0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYCDOMOD0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YCDOMOD0s.bas  ( " _
                & Trim(recYCDOMOD0.obj) & " : " & Trim(recYCDOMOD0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYCDOMOD0_GetBuffer(recYCDOMOD0 As typeYCDOMOD0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYCDOMOD0_GetBuffer = Null
recYCDOMOD0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYCDOMOD0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYCDOMOD0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYCDOMOD0.Err = Space$(10) Then
    recYCDOMOD0.CDOMODETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCDOMOD0.CDOMODAGE = CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCDOMOD0.CDOMODSER = mId$(MsgTxt, K + 11, 2)
    recYCDOMOD0.CDOMODSSE = mId$(MsgTxt, K + 13, 2)
    recYCDOMOD0.CDOMODCOP = mId$(MsgTxt, K + 15, 3)
    recYCDOMOD0.CDOMODDOS = CLng(Val(mId$(MsgTxt, K + 18, 10)))
    recYCDOMOD0.CDOMODNMO = CLng(Val(mId$(MsgTxt, K + 28, 4)))
    recYCDOMOD0.CDOMODNUR = CLng(Val(mId$(MsgTxt, K + 32, 4)))
    recYCDOMOD0.CDOMODNAT = mId$(MsgTxt, K + 36, 3)
    recYCDOMOD0.CDOMODEXT = mId$(MsgTxt, K + 39, 16)
    recYCDOMOD0.CDOMODMON = CCur(Val(mId$(MsgTxt, K + 55, 16))) / 100
    recYCDOMOD0.CDOMODDEV = mId$(MsgTxt, K + 71, 3)
    recYCDOMOD0.CDOMODMOA = CCur(Val(mId$(MsgTxt, K + 74, 16))) / 100
    recYCDOMOD0.CDOMODMOT = CCur(Val(mId$(MsgTxt, K + 90, 16))) / 100
    recYCDOMOD0.CDOMODMOC = CCur(Val(mId$(MsgTxt, K + 106, 16))) / 100
    recYCDOMOD0.CDOMODMOD = CCur(Val(mId$(MsgTxt, K + 122, 16))) / 100
    recYCDOMOD0.CDOMODCON = mId$(MsgTxt, K + 138, 1)
    recYCDOMOD0.CDOMODIRR = mId$(MsgTxt, K + 139, 1)
    recYCDOMOD0.CDOMODFRA = mId$(MsgTxt, K + 140, 1)
    recYCDOMOD0.CDOMODREN = mId$(MsgTxt, K + 141, 1)
    recYCDOMOD0.CDOMODCUM = mId$(MsgTxt, K + 142, 1)
    recYCDOMOD0.CDOMODTRS = mId$(MsgTxt, K + 143, 1)
    recYCDOMOD0.CDOMODTOL = CCur(Val(mId$(MsgTxt, K + 144, 4))) / 100
    recYCDOMOD0.CDOMODTO2 = CCur(Val(mId$(MsgTxt, K + 148, 4))) / 100
    recYCDOMOD0.CDOMODDOR = mId$(MsgTxt, K + 152, 1)
    recYCDOMOD0.CDOMODDON = mId$(MsgTxt, K + 153, 7)
    recYCDOMOD0.CDOMODDOE = mId$(MsgTxt, K + 160, 64)
    recYCDOMOD0.CDOMODBER = mId$(MsgTxt, K + 224, 1)
    recYCDOMOD0.CDOMODBEN = mId$(MsgTxt, K + 225, 7)
    recYCDOMOD0.CDOMODBEI = mId$(MsgTxt, K + 232, 64)
    recYCDOMOD0.CDOMODBAR = mId$(MsgTxt, K + 296, 1)
    recYCDOMOD0.CDOMODBAB = mId$(MsgTxt, K + 297, 7)
    recYCDOMOD0.CDOMODNOR = mId$(MsgTxt, K + 304, 1)
    recYCDOMOD0.CDOMODNOT = mId$(MsgTxt, K + 305, 7)
    recYCDOMOD0.CDOMODBIC = mId$(MsgTxt, K + 312, 12)
    recYCDOMOD0.CDOMODCOT = mId$(MsgTxt, K + 324, 1)
    recYCDOMOD0.CDOMODCOR = mId$(MsgTxt, K + 325, 7)
    recYCDOMOD0.CDOMODPRT = mId$(MsgTxt, K + 332, 1)
    recYCDOMOD0.CDOMODPRR = mId$(MsgTxt, K + 333, 7)
    recYCDOMOD0.CDOMODUTV = mId$(MsgTxt, K + 340, 32)
    recYCDOMOD0.CDOMODPAT = mId$(MsgTxt, K + 372, 1)
    recYCDOMOD0.CDOMODPAR = mId$(MsgTxt, K + 373, 7)
    recYCDOMOD0.CDOMODPAV = mId$(MsgTxt, K + 380, 32)
    recYCDOMOD0.CDOMODOUV = CLng(Val(mId$(MsgTxt, K + 412, 8)))
    recYCDOMOD0.CDOMODEMI = CLng(Val(mId$(MsgTxt, K + 420, 8)))
    recYCDOMOD0.CDOMODVAL = CLng(Val(mId$(MsgTxt, K + 428, 8)))
    recYCDOMOD0.CDOMODDEP = CLng(Val(mId$(MsgTxt, K + 436, 8)))
    recYCDOMOD0.CDOMODDTR = CLng(Val(mId$(MsgTxt, K + 444, 8)))
    recYCDOMOD0.CDOMODVCP = CLng(Val(mId$(MsgTxt, K + 452, 8)))
    recYCDOMOD0.CDOMODCLO = CLng(Val(mId$(MsgTxt, K + 460, 8)))
    recYCDOMOD0.CDOMODREJ = mId$(MsgTxt, K + 468, 3)
    recYCDOMOD0.CDOMODOBJ = mId$(MsgTxt, K + 471, 6)
    recYCDOMOD0.CDOMODAVU = CLng(Val(mId$(MsgTxt, K + 477, 4)))
    recYCDOMOD0.CDOMODMOV = CCur(Val(mId$(MsgTxt, K + 481, 16))) / 100
    recYCDOMOD0.CDOMODCAC = CLng(Val(mId$(MsgTxt, K + 497, 4)))
    recYCDOMOD0.CDOMODMCA = CCur(Val(mId$(MsgTxt, K + 501, 16))) / 100
    recYCDOMOD0.CDOMODDIF = CLng(Val(mId$(MsgTxt, K + 517, 4)))
    recYCDOMOD0.CDOMODMDI = CCur(Val(mId$(MsgTxt, K + 521, 16))) / 100
    recYCDOMOD0.CDOMODPMO = CCur(Val(mId$(MsgTxt, K + 537, 16))) / 100
    recYCDOMOD0.CDOMODPCD = mId$(MsgTxt, K + 553, 20)
    recYCDOMOD0.CDOMODPCC = mId$(MsgTxt, K + 573, 20)
    recYCDOMOD0.CDOMODPDE = CCur(Val(mId$(MsgTxt, K + 593, 16))) / 100
    recYCDOMOD0.CDOMODPPO = CLng(Val(mId$(MsgTxt, K + 609, 4)))
    recYCDOMOD0.CDOMODAUT = mId$(MsgTxt, K + 613, 12)
    recYCDOMOD0.CDOMODREG = CCur(Val(mId$(MsgTxt, K + 625, 16))) / 100
    recYCDOMOD0.CDOMODENC = CCur(Val(mId$(MsgTxt, K + 641, 16))) / 100
    recYCDOMOD0.CDOMODDAN = CLng(Val(mId$(MsgTxt, K + 657, 8)))
    recYCDOMOD0.CDOMODANN = CCur(Val(mId$(MsgTxt, K + 665, 16))) / 100
    recYCDOMOD0.CDOMODPCO = CDbl(Val(mId$(MsgTxt, K + 681, 15))) / 1000000000
    recYCDOMOD0.CDOMODLEM = mId$(MsgTxt, K + 696, 30)
    recYCDOMOD0.CDOMODLDE = mId$(MsgTxt, K + 726, 30)
    recYCDOMOD0.CDOMODDLE = CLng(Val(mId$(MsgTxt, K + 756, 8)))
    recYCDOMOD0.CDOMODEPA = mId$(MsgTxt, K + 764, 1)
    recYCDOMOD0.CDOMODTRA = mId$(MsgTxt, K + 765, 1)
    recYCDOMOD0.CDOMODFCD = mId$(MsgTxt, K + 766, 1)
    recYCDOMOD0.CDOMODCUS = CInt(Val(mId$(MsgTxt, K + 767, 5)))
    recYCDOMOD0.CDOMODCUV = CInt(Val(mId$(MsgTxt, K + 772, 5)))
    recYCDOMOD0.CDOMODCU2 = CInt(Val(mId$(MsgTxt, K + 777, 5)))
    recYCDOMOD0.CDOMODOPE = mId$(MsgTxt, K + 782, 1)
    recYCDOMOD0.CDOMODPOO = mId$(MsgTxt, K + 783, 1)
    recYCDOMOD0.CDOMODPBE = CCur(Val(mId$(MsgTxt, K + 784, 16))) / 100
    recYCDOMOD0.CDOMODGAG = mId$(MsgTxt, K + 800, 1)
    recYCDOMOD0.CDOMODSTB = mId$(MsgTxt, K + 801, 1)
    recYCDOMOD0.CDOMODMRE = mId$(MsgTxt, K + 802, 3)
    recYCDOMOD0.CDOMODNPD = CLng(Val(mId$(MsgTxt, K + 805, 4)))
    recYCDOMOD0.CDOMODTJD = mId$(MsgTxt, K + 809, 1)
    recYCDOMOD0.CDOMODPDO = mId$(MsgTxt, K + 810, 60)
    recYCDOMOD0.CDOMODGAR = mId$(MsgTxt, K + 870, 64)
    recYCDOMOD0.CDOMODOBM = mId$(MsgTxt, K + 934, 64)
    recYCDOMOD0.CDOMODTBR = mId$(MsgTxt, K + 998, 1)
    recYCDOMOD0.CDOMODBRE = mId$(MsgTxt, K + 999, 7)
    recYCDOMOD0.CDOMODBEC = mId$(MsgTxt, K + 1006, 1)
    recYCDOMOD0.CDOMODRNO = mId$(MsgTxt, K + 1007, 16)
    recYCDOMOD0.CDOMODDPA = mId$(MsgTxt, K + 1023, 3)
    recYCDOMOD0.CDOMODDVI = mId$(MsgTxt, K + 1026, 32)
    recYCDOMOD0.CDOMODEPY = mId$(MsgTxt, K + 1058, 3)
    recYCDOMOD0.CDOMODEVI = mId$(MsgTxt, K + 1061, 32)
    recYCDOMOD0.CDOMODVPA = mId$(MsgTxt, K + 1093, 3)
    recYCDOMOD0.CDOMODVVI = mId$(MsgTxt, K + 1096, 32)
    recYCDOMOD0.CDOMODNDE = CLng(Val(mId$(MsgTxt, K + 1128, 10)))
    recYCDOMOD0.CDOMODNAE = mId$(MsgTxt, K + 1138, 3)
    recYCDOMOD0.CDOMODEVE = mId$(MsgTxt, K + 1141, 2)
    recYCDOMOD0.CDOMODETA = mId$(MsgTxt, K + 1143, 2)
    recYCDOMOD0.CDOMODDMO = CLng(Val(mId$(MsgTxt, K + 1145, 8)))
    recYCDOMOD0.CDOMODDRM = CLng(Val(mId$(MsgTxt, K + 1153, 8)))
    recYCDOMOD0.CDOMODDP2 = mId$(MsgTxt, K + 1161, 32)
    recYCDOMOD0.CDOMODEP2 = mId$(MsgTxt, K + 1193, 32)
    recYCDOMOD0.CDOMODPD2 = mId$(MsgTxt, K + 1225, 80)
    recYCDOMOD0.CDOMODAUN = mId$(MsgTxt, K + 1305, 12)
    recYCDOMOD0.CDOMODCER = mId$(MsgTxt, K + 1317, 1)


Else
    srvYCDOMOD0_GetBuffer = recYCDOMOD0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYCDOMOD0Len

End Function

'---------------------------------------------------------
Private Sub srvYCDOMOD0_PutBuffer(recYCDOMOD0 As typeYCDOMOD0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYCDOMOD0Len) = Space$(recYCDOMOD0Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYCDOMOD0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYCDOMOD0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYCDOMOD0.CDOMODETB, "0000 ")
    Mid$(MsgTxt, K + 6, 5) = Format$(recYCDOMOD0.CDOMODAGE, "0000 ")
    Mid$(MsgTxt, K + 11, 2) = recYCDOMOD0.CDOMODSER
    Mid$(MsgTxt, K + 13, 2) = recYCDOMOD0.CDOMODSSE
    Mid$(MsgTxt, K + 15, 3) = recYCDOMOD0.CDOMODCOP
    Mid$(MsgTxt, K + 18, 10) = Format$(recYCDOMOD0.CDOMODDOS, "000000000 ")
    Mid$(MsgTxt, K + 28, 4) = Format$(recYCDOMOD0.CDOMODNMO, "000 ")
    Mid$(MsgTxt, K + 32, 4) = Format$(recYCDOMOD0.CDOMODNUR, "000 ")
    Mid$(MsgTxt, K + 36, 3) = recYCDOMOD0.CDOMODNAT
    Mid$(MsgTxt, K + 39, 16) = recYCDOMOD0.CDOMODEXT
    Mid$(MsgTxt, K + 55, 16) = Format$(recYCDOMOD0.CDOMODMON * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 71, 3) = recYCDOMOD0.CDOMODDEV
    Mid$(MsgTxt, K + 74, 16) = Format$(recYCDOMOD0.CDOMODMOA * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 90, 16) = Format$(recYCDOMOD0.CDOMODMOT * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 106, 16) = Format$(recYCDOMOD0.CDOMODMOC * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 122, 16) = Format$(recYCDOMOD0.CDOMODMOD * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 138, 1) = recYCDOMOD0.CDOMODCON
    Mid$(MsgTxt, K + 139, 1) = recYCDOMOD0.CDOMODIRR
    Mid$(MsgTxt, K + 140, 1) = recYCDOMOD0.CDOMODFRA
    Mid$(MsgTxt, K + 141, 1) = recYCDOMOD0.CDOMODREN
    Mid$(MsgTxt, K + 142, 1) = recYCDOMOD0.CDOMODCUM
    Mid$(MsgTxt, K + 143, 1) = recYCDOMOD0.CDOMODTRS
    Mid$(MsgTxt, K + 144, 4) = Format$(recYCDOMOD0.CDOMODTOL * 100, "000 ")
    Mid$(MsgTxt, K + 148, 4) = Format$(recYCDOMOD0.CDOMODTO2 * 100, "000 ")
    Mid$(MsgTxt, K + 152, 1) = recYCDOMOD0.CDOMODDOR
    Mid$(MsgTxt, K + 153, 7) = recYCDOMOD0.CDOMODDON
    Mid$(MsgTxt, K + 160, 64) = recYCDOMOD0.CDOMODDOE
    Mid$(MsgTxt, K + 224, 1) = recYCDOMOD0.CDOMODBER
    Mid$(MsgTxt, K + 225, 7) = recYCDOMOD0.CDOMODBEN
    Mid$(MsgTxt, K + 232, 64) = recYCDOMOD0.CDOMODBEI
    Mid$(MsgTxt, K + 296, 1) = recYCDOMOD0.CDOMODBAR
    Mid$(MsgTxt, K + 297, 7) = recYCDOMOD0.CDOMODBAB
    Mid$(MsgTxt, K + 304, 1) = recYCDOMOD0.CDOMODNOR
    Mid$(MsgTxt, K + 305, 7) = recYCDOMOD0.CDOMODNOT
    Mid$(MsgTxt, K + 312, 12) = recYCDOMOD0.CDOMODBIC
    Mid$(MsgTxt, K + 324, 1) = recYCDOMOD0.CDOMODCOT
    Mid$(MsgTxt, K + 325, 7) = recYCDOMOD0.CDOMODCOR
    Mid$(MsgTxt, K + 332, 1) = recYCDOMOD0.CDOMODPRT
    Mid$(MsgTxt, K + 333, 7) = recYCDOMOD0.CDOMODPRR
    Mid$(MsgTxt, K + 340, 32) = recYCDOMOD0.CDOMODUTV
    Mid$(MsgTxt, K + 372, 1) = recYCDOMOD0.CDOMODPAT
    Mid$(MsgTxt, K + 373, 7) = recYCDOMOD0.CDOMODPAR
    Mid$(MsgTxt, K + 380, 32) = recYCDOMOD0.CDOMODPAV
    Mid$(MsgTxt, K + 412, 8) = Format$(recYCDOMOD0.CDOMODOUV, "0000000 ")
    Mid$(MsgTxt, K + 420, 8) = Format$(recYCDOMOD0.CDOMODEMI, "0000000 ")
    Mid$(MsgTxt, K + 428, 8) = Format$(recYCDOMOD0.CDOMODVAL, "0000000 ")
    Mid$(MsgTxt, K + 436, 8) = Format$(recYCDOMOD0.CDOMODDEP, "0000000 ")
    Mid$(MsgTxt, K + 444, 8) = Format$(recYCDOMOD0.CDOMODDTR, "0000000 ")
    Mid$(MsgTxt, K + 452, 8) = Format$(recYCDOMOD0.CDOMODVCP, "0000000 ")
    Mid$(MsgTxt, K + 460, 8) = Format$(recYCDOMOD0.CDOMODCLO, "0000000 ")
    Mid$(MsgTxt, K + 468, 3) = recYCDOMOD0.CDOMODREJ
    Mid$(MsgTxt, K + 471, 6) = recYCDOMOD0.CDOMODOBJ
    Mid$(MsgTxt, K + 477, 4) = Format$(recYCDOMOD0.CDOMODAVU, "000 ")
    Mid$(MsgTxt, K + 481, 16) = Format$(recYCDOMOD0.CDOMODMOV * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 497, 4) = Format$(recYCDOMOD0.CDOMODCAC, "000 ")
    Mid$(MsgTxt, K + 501, 16) = Format$(recYCDOMOD0.CDOMODMCA * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 517, 4) = Format$(recYCDOMOD0.CDOMODDIF, "000 ")
    Mid$(MsgTxt, K + 521, 16) = Format$(recYCDOMOD0.CDOMODMDI * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 537, 16) = Format$(recYCDOMOD0.CDOMODPMO * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 553, 20) = recYCDOMOD0.CDOMODPCD
    Mid$(MsgTxt, K + 573, 20) = recYCDOMOD0.CDOMODPCC
    Mid$(MsgTxt, K + 593, 16) = Format$(recYCDOMOD0.CDOMODPDE * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 609, 4) = Format$(recYCDOMOD0.CDOMODPPO, "000 ")
    Mid$(MsgTxt, K + 613, 12) = recYCDOMOD0.CDOMODAUT
    Mid$(MsgTxt, K + 625, 16) = Format$(recYCDOMOD0.CDOMODREG * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 641, 16) = Format$(recYCDOMOD0.CDOMODENC * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 657, 8) = Format$(recYCDOMOD0.CDOMODDAN, "0000000 ")
    Mid$(MsgTxt, K + 665, 16) = Format$(recYCDOMOD0.CDOMODANN * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 681, 15) = Format$(recYCDOMOD0.CDOMODPCO * 1000000000, "00000000000000 ")
    Mid$(MsgTxt, K + 696, 30) = recYCDOMOD0.CDOMODLEM
    Mid$(MsgTxt, K + 726, 30) = recYCDOMOD0.CDOMODLDE
    Mid$(MsgTxt, K + 756, 8) = Format$(recYCDOMOD0.CDOMODDLE, "0000000 ")
    Mid$(MsgTxt, K + 764, 1) = recYCDOMOD0.CDOMODEPA
    Mid$(MsgTxt, K + 765, 1) = recYCDOMOD0.CDOMODTRA
    Mid$(MsgTxt, K + 766, 1) = recYCDOMOD0.CDOMODFCD
    Mid$(MsgTxt, K + 767, 5) = Format$(recYCDOMOD0.CDOMODCUS, "0000 ")
    Mid$(MsgTxt, K + 772, 5) = Format$(recYCDOMOD0.CDOMODCUV, "0000 ")
    Mid$(MsgTxt, K + 777, 5) = Format$(recYCDOMOD0.CDOMODCU2, "0000 ")
    Mid$(MsgTxt, K + 782, 1) = recYCDOMOD0.CDOMODOPE
    Mid$(MsgTxt, K + 783, 1) = recYCDOMOD0.CDOMODPOO
    Mid$(MsgTxt, K + 784, 16) = Format$(recYCDOMOD0.CDOMODPBE * 100, "000000000000000 ")
    Mid$(MsgTxt, K + 800, 1) = recYCDOMOD0.CDOMODGAG
    Mid$(MsgTxt, K + 801, 1) = recYCDOMOD0.CDOMODSTB
    Mid$(MsgTxt, K + 802, 3) = recYCDOMOD0.CDOMODMRE
    Mid$(MsgTxt, K + 805, 4) = Format$(recYCDOMOD0.CDOMODNPD, "000 ")
    Mid$(MsgTxt, K + 809, 1) = recYCDOMOD0.CDOMODTJD
    Mid$(MsgTxt, K + 810, 60) = recYCDOMOD0.CDOMODPDO
    Mid$(MsgTxt, K + 870, 64) = recYCDOMOD0.CDOMODGAR
    Mid$(MsgTxt, K + 934, 64) = recYCDOMOD0.CDOMODOBM
    Mid$(MsgTxt, K + 998, 1) = recYCDOMOD0.CDOMODTBR
    Mid$(MsgTxt, K + 999, 7) = recYCDOMOD0.CDOMODBRE
    Mid$(MsgTxt, K + 1006, 1) = recYCDOMOD0.CDOMODBEC
    Mid$(MsgTxt, K + 1007, 16) = recYCDOMOD0.CDOMODRNO
    Mid$(MsgTxt, K + 1023, 3) = recYCDOMOD0.CDOMODDPA
    Mid$(MsgTxt, K + 1026, 32) = recYCDOMOD0.CDOMODDVI
    Mid$(MsgTxt, K + 1058, 3) = recYCDOMOD0.CDOMODEPY
    Mid$(MsgTxt, K + 1061, 32) = recYCDOMOD0.CDOMODEVI
    Mid$(MsgTxt, K + 1093, 3) = recYCDOMOD0.CDOMODVPA
    Mid$(MsgTxt, K + 1096, 32) = recYCDOMOD0.CDOMODVVI
    Mid$(MsgTxt, K + 1128, 10) = Format$(recYCDOMOD0.CDOMODNDE, "000000000 ")
    Mid$(MsgTxt, K + 1138, 3) = recYCDOMOD0.CDOMODNAE
    Mid$(MsgTxt, K + 1141, 2) = recYCDOMOD0.CDOMODEVE
    Mid$(MsgTxt, K + 1143, 2) = recYCDOMOD0.CDOMODETA
    Mid$(MsgTxt, K + 1145, 8) = Format$(recYCDOMOD0.CDOMODDMO, "0000000 ")
    Mid$(MsgTxt, K + 1153, 8) = Format$(recYCDOMOD0.CDOMODDRM, "0000000 ")
    Mid$(MsgTxt, K + 1161, 32) = recYCDOMOD0.CDOMODDP2
    Mid$(MsgTxt, K + 1193, 32) = recYCDOMOD0.CDOMODEP2
    Mid$(MsgTxt, K + 1225, 80) = recYCDOMOD0.CDOMODPD2
    Mid$(MsgTxt, K + 1305, 12) = recYCDOMOD0.CDOMODAUN
    Mid$(MsgTxt, K + 1317, 1) = recYCDOMOD0.CDOMODCER


MsgTxtLen = MsgTxtLen + recYCDOMOD0Len
End Sub



Public Sub srvYCDOMOD0_ElpDisplay(recYCDOMOD0 As typeYCDOMOD0)
frmElpDisplay.fgData.Rows = 111
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODSER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODSER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODSSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODSSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODCOP    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPERATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODCOP
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDOS    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDOS
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODNMO    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° MODIFICATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODNMO
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODNUR    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° RENOUVELLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODNUR
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODNAT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NATURE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODNAT
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODEXT   16A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REFERENCE EXTERNE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODEXT
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODMON 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODMON
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDEV
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODMOA 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT ADDITIONNEL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODMOA
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODMOT 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT TOTAL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODMOT
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODMOC 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT CONFIRME"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODMOC
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODMOD 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DUCROIRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODMOD
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODCON    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CONFIRM NOTIFI PARTI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODCON
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODIRR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "IRREVOCABLE (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODIRR
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODFRA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "FRACTIONNABLE (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODFRA
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODREN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RENOUVELABLE (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODREN
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODCUM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CUMULATIF (O/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODCUM
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODTRS    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TRANSFERABLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODTRS
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODTOL  3.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TOLERANCE +"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODTOL
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODTO2  3.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TOLERANCE -"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODTO2
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDOR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DONN. ORDRE CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDOR
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDON    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DONNEUR ORDRE IMPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDON
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDOE   64A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DONNEUR ORDRE EXPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDOE
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODBER    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BENEFICIAIR CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODBER
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODBEN    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BENEFICIAIRE EXPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODBEN
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODBEI   64A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BENEFICIAIRE IMPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODBEI
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODBAR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BANQU.BENEF.CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODBAR
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODBAB    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BANQUE BENEF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODBAB
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODNOR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NOTIF/CONFI OU EMETT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODNOR
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODNOT    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NOTIF/CONFI OU EMETT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODNOT
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODBIC   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BIC SUPPLEMEN. EMETT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODBIC
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODCOT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CORRESPOND. CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODCOT
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODCOR    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CORRESPONDANT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODCOR
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPRT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU PRES CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPRT
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPRR    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU PRESENTATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPRR
frmElpDisplay.fgData.Row = 40
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODUTV   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU PRESENTATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODUTV
frmElpDisplay.fgData.Row = 41
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPAT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU PAIE CLI/TIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPAT
frmElpDisplay.fgData.Row = 42
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPAR    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU PAIEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPAR
frmElpDisplay.fgData.Row = 43
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPAV   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU PAIEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPAV
frmElpDisplay.fgData.Row = 44
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODOUV    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE OUVERTURE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODOUV
frmElpDisplay.fgData.Row = 45
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODEMI    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE EMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODEMI
frmElpDisplay.fgData.Row = 46
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODVAL    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE VALIDITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODVAL
frmElpDisplay.fgData.Row = 47
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDEP    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE EXTREME PAYMT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDEP
frmElpDisplay.fgData.Row = 48
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDTR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DE TRANSFERT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDTR
frmElpDisplay.fgData.Row = 49
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODVCP    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE VALID. COMPTA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODVCP
frmElpDisplay.fgData.Row = 50
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODCLO    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE CLOTURE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODCLO
frmElpDisplay.fgData.Row = 51
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODREJ    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MOTIF REJET (CLOTUR)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODREJ
frmElpDisplay.fgData.Row = 52
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODOBJ    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "OBJET CREDIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODOBJ
frmElpDisplay.fgData.Row = 53
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODAVU    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "% PAIEM. A VUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODAVU
frmElpDisplay.fgData.Row = 54
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODMOV 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT A VUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODMOV
frmElpDisplay.fgData.Row = 55
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODCAC    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "% PAIEM. CTR ACCEPT."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODCAC
frmElpDisplay.fgData.Row = 56
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODMCA 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT CTR ACCEPT."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODMCA
frmElpDisplay.fgData.Row = 57
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDIF    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "% PAIEM. DIFFERE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDIF
frmElpDisplay.fgData.Row = 58
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODMDI 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT. DIFFERE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODMDI
frmElpDisplay.fgData.Row = 59
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPMO 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT PROVISIONNE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPMO
frmElpDisplay.fgData.Row = 60
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPCD   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PROV. DEBIT  COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPCD
frmElpDisplay.fgData.Row = 61
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPCC   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PROV. CREDIT COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPCC
frmElpDisplay.fgData.Row = 62
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPDE 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PROVISION DEVISE DOS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPDE
frmElpDisplay.fgData.Row = 63
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPPO    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PROVISION POURCEN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPPO
frmElpDisplay.fgData.Row = 64
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODAUT   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE AUTORISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODAUT
frmElpDisplay.fgData.Row = 65
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODREG 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT PAYE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODREG
frmElpDisplay.fgData.Row = 66
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODENC 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT ENCAISSE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODENC
frmElpDisplay.fgData.Row = 67
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDAN    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ANNULATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDAN
frmElpDisplay.fgData.Row = 68
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODANN 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT ANNULE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODANN
frmElpDisplay.fgData.Row = 69
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPCO 14.9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COURS DEVPRO/DEVDOS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPCO
frmElpDisplay.fgData.Row = 70
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODLEM   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU EMBARQUEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODLEM
frmElpDisplay.fgData.Row = 71
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODLDE   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIEU DESTINATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODLDE
frmElpDisplay.fgData.Row = 72
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDLE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE LIMITE EMBARQU."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDLE
frmElpDisplay.fgData.Row = 73
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODEPA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EXPED.PARTIE.AUTORI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODEPA
frmElpDisplay.fgData.Row = 74
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODTRA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TRANBORDEMENT AUTORI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODTRA
frmElpDisplay.fgData.Row = 75
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODFCD    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "FRAI CHARGE D.O. BEN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODFCD
frmElpDisplay.fgData.Row = 76
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODCUS    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILI. DE SAISIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODCUS
frmElpDisplay.fgData.Row = 77
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODCUV    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "1ER VALIDEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODCUV
frmElpDisplay.fgData.Row = 78
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODCU2    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "2EME VALIDEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODCU2
frmElpDisplay.fgData.Row = 79
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODOPE    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "OPERATIVITE DU CRED."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODOPE
frmElpDisplay.fgData.Row = 80
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPOO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EXISTENCE POOL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPOO
frmElpDisplay.fgData.Row = 81
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPBE 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PART.BANQUE EXPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPBE
frmElpDisplay.fgData.Row = 82
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODGAG    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "GAGE MARCHANDISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODGAG
frmElpDisplay.fgData.Row = 83
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODSTB    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "STAND BY"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODSTB
frmElpDisplay.fgData.Row = 84
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODMRE    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MODE DE REALISAT°"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODMRE
frmElpDisplay.fgData.Row = 85
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODNPD    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NBJ PRES. DOCUMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODNPD
frmElpDisplay.fgData.Row = 86
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODTJD    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TY JOUR DOCS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODTJD
frmElpDisplay.fgData.Row = 87
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPDO   60A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PER.PRE.DOCS."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPDO
frmElpDisplay.fgData.Row = 88
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODGAR   64A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBELLE GARANTIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODGAR
frmElpDisplay.fgData.Row = 89
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODOBM   64A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "OBJET DE MODIF."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODOBM
frmElpDisplay.fgData.Row = 90
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODTBR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TIERS BQ REMBOURS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODTBR
frmElpDisplay.fgData.Row = 91
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODBRE    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BQ REMBOURSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODBRE
frmElpDisplay.fgData.Row = 92
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODBEC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BENEF.PAY.COMMIS°"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODBEC
frmElpDisplay.fgData.Row = 93
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODRNO   16A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REF.NOTIFICATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODRNO
frmElpDisplay.fgData.Row = 94
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDPA    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DESTINATION PAYS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDPA
frmElpDisplay.fgData.Row = 95
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDVI   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DESTINATION VILLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDVI
frmElpDisplay.fgData.Row = 96
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODEPY    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EMBARQUEMENT PAYS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODEPY
frmElpDisplay.fgData.Row = 97
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODEVI   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EMBARQUEMENT VILLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODEVI
frmElpDisplay.fgData.Row = 98
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODVPA    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "VALIDITE PAYS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODVPA
frmElpDisplay.fgData.Row = 99
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODVVI   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "VALIDIT VILLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODVVI
frmElpDisplay.fgData.Row = 100
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODNDE    9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DOSSIER EXPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODNDE
frmElpDisplay.fgData.Row = 101
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODNAE    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NATURE EXPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODNAE
frmElpDisplay.fgData.Row = 102
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODEVE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EVENEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODEVE
frmElpDisplay.fgData.Row = 103
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODETA    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETAT DE LA MODIF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODETA
frmElpDisplay.fgData.Row = 104
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDMO    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE MODIFICATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDMO
frmElpDisplay.fgData.Row = 105
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDRM    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE RECEPTION MOD"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDRM
frmElpDisplay.fgData.Row = 106
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODDP2   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DESTIN.PAYS LIBELLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODDP2
frmElpDisplay.fgData.Row = 107
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODEP2   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EMBARQ.PAYS LIBELLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODEP2
frmElpDisplay.fgData.Row = 108
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODPD2   80A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PER.PRES.DOC.SUITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODPD2
frmElpDisplay.fgData.Row = 109
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODAUN   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE AUT. NOTIFIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODAUN
frmElpDisplay.fgData.Row = 110
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CDOMODCER    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTAT°(O=CERTAIN/N)"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCDOMOD0.CDOMODCER
frmElpDisplay.Show vbModal
End Sub

'---------------------------------------------------------
Private Function srvYCDOMOD0_Seek(recYCDOMOD0 As typeYCDOMOD0)
'---------------------------------------------------------

srvYCDOMOD0_Seek = "?"
MsgTxtLen = 0
Call srvYCDOMOD0_PutBuffer(recYCDOMOD0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYCDOMOD0_GetBuffer(recYCDOMOD0)) Then
            srvYCDOMOD0_Seek = Null
        Else
            Call srvYCDOMOD0_Error(recYCDOMOD0)
        End If
    End If
End If

End Function

'-----------------------------------------------------
Function srvYCDOMOD0_Update(recYCDOMOD0 As typeYCDOMOD0)
'-----------------------------------------------------

srvYCDOMOD0_Update = "?"

MsgTxtLen = 0
Call srvYCDOMOD0_PutBuffer(recYCDOMOD0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYCDOMOD0_GetBuffer(recYCDOMOD0)) Then
        Call srvYCDOMOD0_Error(recYCDOMOD0)
        srvYCDOMOD0_Update = recYCDOMOD0.Err
        Exit Function
    Else
        srvYCDOMOD0_Update = Null
    End If
Else
    recYCDOMOD0.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYCDOMOD0_Init(recYCDOMOD0 As typeYCDOMOD0)
'---------------------------------------------------------
MsgTxt = Space$(recYCDOMOD0Len)
MsgTxtIndex = 0
Call srvYCDOMOD0_GetBuffer(recYCDOMOD0)
recYCDOMOD0.obj = "ZCDOMOD0_S"
recYCDOMOD0.CDOMODETB = 1 '      As Integer                        ' CODE ETABLISSEMENT
recYCDOMOD0.CDOMODAGE = 1 '      As Integer                        ' AGENCE
recYCDOMOD0.CDOMODSER = "00" '=""     'As String * 2                     ' SERVICE
recYCDOMOD0.CDOMODSSE = "00" '      As String * 2                     ' SOUS-SERVICE
recYCDOMOD0.CDOMODCOP = ""   'As String * 3                     ' CODE OPERATION
recYCDOMOD0.CDOMODDOS = 0   'As Long                           ' NUMERO DOSSIER
recYCDOMOD0.CDOMODNMO = 0   'As Long                           ' N° MODIFICATION
recYCDOMOD0.CDOMODNUR = 0   'As Long                           ' N° RENOUVELLEMENT
recYCDOMOD0.CDOMODNAT = ""   'As String * 3                     ' NATURE
recYCDOMOD0.CDOMODEXT = ""   'As String * 16                    ' REFERENCE EXTERNE
recYCDOMOD0.CDOMODMON = 0   'As Currency                       ' MONTANT DOSSIER
recYCDOMOD0.CDOMODDEV = ""   'As String * 3                     ' DEVISE
recYCDOMOD0.CDOMODMOA = 0   'As Currency                       ' MONTANT ADDITIONNEL
recYCDOMOD0.CDOMODMOT = 0   'As Currency                       ' MONTANT TOTAL
recYCDOMOD0.CDOMODMOC = 0   'As Currency                       ' MONTANT CONFIRME
recYCDOMOD0.CDOMODMOD = 0   'As Currency                       ' MONTANT DUCROIRE
recYCDOMOD0.CDOMODCON = ""   'As String * 1                     ' CONFIRM NOTIFI PARTI
recYCDOMOD0.CDOMODIRR = ""   'As String * 1                     ' IRREVOCABLE (O/N)
recYCDOMOD0.CDOMODFRA = ""   'As String * 1                     ' FRACTIONNABLE (O/N)
recYCDOMOD0.CDOMODREN = ""   'As String * 1                     ' RENOUVELABLE (O/N)
recYCDOMOD0.CDOMODCUM = ""   'As String * 1                     ' CUMULATIF (O/N)
recYCDOMOD0.CDOMODTRS = ""   'As String * 1                     ' TRANSFERABLE
recYCDOMOD0.CDOMODTOL = 0   'As Currency                       ' TOLERANCE +
recYCDOMOD0.CDOMODTO2 = 0   'As Currency                       ' TOLERANCE -
recYCDOMOD0.CDOMODDOR = ""   'As String * 1                     ' DONN. ORDRE CLI/TIE
recYCDOMOD0.CDOMODDON = ""   'As String * 7                     ' DONNEUR ORDRE IMPORT
recYCDOMOD0.CDOMODDOE = ""   'As String * 64                    ' DONNEUR ORDRE EXPORT
recYCDOMOD0.CDOMODBER = ""   'As String * 1                     ' BENEFICIAIR CLI/TIE
recYCDOMOD0.CDOMODBEN = ""   'As String * 7                     ' BENEFICIAIRE EXPORT
recYCDOMOD0.CDOMODBEI = ""   'As String * 64                    ' BENEFICIAIRE IMPORT
recYCDOMOD0.CDOMODBAR = ""   'As String * 1                     ' BANQU.BENEF.CLI/TIE
recYCDOMOD0.CDOMODBAB = ""   'As String * 7                     ' BANQUE BENEF
recYCDOMOD0.CDOMODNOR = ""   'As String * 1                     ' NOTIF/CONFI OU EMETT
recYCDOMOD0.CDOMODNOT = ""   'As String * 7                     ' NOTIF/CONFI OU EMETT
recYCDOMOD0.CDOMODBIC = ""   'As String * 12                    ' BIC SUPPLEMEN. EMETT
recYCDOMOD0.CDOMODCOT = ""   'As String * 1                     ' CORRESPOND. CLI/TIE
recYCDOMOD0.CDOMODCOR = ""   'As String * 7                     ' CORRESPONDANT
recYCDOMOD0.CDOMODPRT = ""   'As String * 1                     ' LIEU PRES CLI/TIE
recYCDOMOD0.CDOMODPRR = ""   'As String * 7                     ' LIEU PRESENTATION
recYCDOMOD0.CDOMODUTV = ""   'As String * 32                    ' LIEU PRESENTATION
recYCDOMOD0.CDOMODPAT = ""   'As String * 1                     ' LIEU PAIE CLI/TIE
recYCDOMOD0.CDOMODPAR = ""   'As String * 7                     ' LIEU PAIEMENT
recYCDOMOD0.CDOMODPAV = ""   'As String * 32                    ' LIEU PAIEMENT
recYCDOMOD0.CDOMODOUV = 0   'As Long                           ' DATE OUVERTURE
recYCDOMOD0.CDOMODEMI = 0   'As Long                           ' DATE EMISSION
recYCDOMOD0.CDOMODVAL = 0   'As Long                           ' DATE VALIDITE
recYCDOMOD0.CDOMODDEP = 0   'As Long                           ' DATE EXTREME PAYMT
recYCDOMOD0.CDOMODDTR = 0   'As Long                           ' DATE DE TRANSFERT
recYCDOMOD0.CDOMODVCP = 0   'As Long                           ' DATE VALID. COMPTA
recYCDOMOD0.CDOMODCLO = 0   'As Long                           ' DATE CLOTURE
recYCDOMOD0.CDOMODREJ = ""   'As String * 3                     ' MOTIF REJET (CLOTUR)
recYCDOMOD0.CDOMODOBJ = ""   'As String * 6                     ' OBJET CREDIT
recYCDOMOD0.CDOMODAVU = 0   'As Long                           ' % PAIEM. A VUE
recYCDOMOD0.CDOMODMOV = 0   'As Currency                       ' MONTANT A VUE
recYCDOMOD0.CDOMODCAC = 0   'As Long                           ' % PAIEM. CTR ACCEPT.
recYCDOMOD0.CDOMODMCA = 0   'As Currency                       ' MONTANT CTR ACCEPT.
recYCDOMOD0.CDOMODDIF = 0   'As Long                           ' % PAIEM. DIFFERE
recYCDOMOD0.CDOMODMDI = 0   'As Currency                       ' MONTANT. DIFFERE
recYCDOMOD0.CDOMODPMO = 0   'As Currency                       ' MONTANT PROVISIONNE
recYCDOMOD0.CDOMODPCD = ""   'As String * 20                    ' PROV. DEBIT  COMPTE
recYCDOMOD0.CDOMODPCC = ""   'As String * 20                    ' PROV. CREDIT COMPTE
recYCDOMOD0.CDOMODPDE = 0   'As Currency                       ' PROVISION DEVISE DOS
recYCDOMOD0.CDOMODPPO = 0   'As Long                           ' PROVISION POURCEN
recYCDOMOD0.CDOMODAUT = ""   'As String * 12                    ' CODE AUTORISATION
recYCDOMOD0.CDOMODREG = 0   'As Currency                       ' MONTANT PAYE
recYCDOMOD0.CDOMODENC = 0   'As Currency                       ' MONTANT ENCAISSE
recYCDOMOD0.CDOMODDAN = 0   'As Long                           ' DATE ANNULATION
recYCDOMOD0.CDOMODANN = 0   'As Currency                       ' MONTANT ANNULE
recYCDOMOD0.CDOMODPCO = 0   'As Double                         ' COURS DEVPRO/DEVDOS
recYCDOMOD0.CDOMODLEM = ""   'As String * 30                    ' LIEU EMBARQUEMENT
recYCDOMOD0.CDOMODLDE = ""   'As String * 30                    ' LIEU DESTINATION
recYCDOMOD0.CDOMODDLE = 0   'As Long                           ' DATE LIMITE EMBARQU.
recYCDOMOD0.CDOMODEPA = ""   'As String * 1                     ' EXPED.PARTIE.AUTORI
recYCDOMOD0.CDOMODTRA = ""   'As String * 1                     ' TRANBORDEMENT AUTORI
recYCDOMOD0.CDOMODFCD = ""   'As String * 1                     ' FRAI CHARGE D.O. BEN
recYCDOMOD0.CDOMODCUS = 0   'As Integer                        ' UTILI. DE SAISIE
recYCDOMOD0.CDOMODCUV = 0   'As Integer                        ' 1ER VALIDEUR
recYCDOMOD0.CDOMODCU2 = 0   'As Integer                        ' 2EME VALIDEUR
recYCDOMOD0.CDOMODOPE = ""   'As String * 1                     ' OPERATIVITE DU CRED.
recYCDOMOD0.CDOMODPOO = ""   'As String * 1                     ' EXISTENCE POOL
recYCDOMOD0.CDOMODPBE = 0   'As Currency                       ' PART.BANQUE EXPORT
recYCDOMOD0.CDOMODGAG = ""   'As String * 1                     ' GAGE MARCHANDISE
recYCDOMOD0.CDOMODSTB = ""   'As String * 1                     ' STAND BY
recYCDOMOD0.CDOMODMRE = ""   'As String * 3                     ' MODE DE REALISAT°
recYCDOMOD0.CDOMODNPD = 0   'As Long                           ' NBJ PRES. DOCUMENT
recYCDOMOD0.CDOMODTJD = ""   'As String * 1                     ' TY JOUR DOCS
recYCDOMOD0.CDOMODPDO = ""   'As String * 60                    ' PER.PRE.DOCS.
recYCDOMOD0.CDOMODGAR = ""   'As String * 64                    ' LIBELLE GARANTIE
recYCDOMOD0.CDOMODOBM = ""   'As String * 64                    ' OBJET DE MODIF.
recYCDOMOD0.CDOMODTBR = ""   'As String * 1                     ' TIERS BQ REMBOURS
recYCDOMOD0.CDOMODBRE = ""   'As String * 7                     ' BQ REMBOURSEMENT
recYCDOMOD0.CDOMODBEC = ""   'As String * 1                     ' BENEF.PAY.COMMIS°
recYCDOMOD0.CDOMODRNO = ""   'As String * 16                    ' REF.NOTIFICATEUR
recYCDOMOD0.CDOMODDPA = ""   'As String * 3                     ' DESTINATION PAYS
recYCDOMOD0.CDOMODDVI = ""   'As String * 32                    ' DESTINATION VILLE
recYCDOMOD0.CDOMODEPY = ""   'As String * 3                     ' EMBARQUEMENT PAYS
recYCDOMOD0.CDOMODEVI = ""   'As String * 32                    ' EMBARQUEMENT VILLE
recYCDOMOD0.CDOMODVPA = ""   'As String * 3                     ' VALIDITE PAYS
recYCDOMOD0.CDOMODVVI = ""   'As String * 32                    ' VALIDIT VILLE
recYCDOMOD0.CDOMODNDE = 0   'As Long                           ' DOSSIER EXPORT
recYCDOMOD0.CDOMODNAE = ""   'As String * 3                     ' NATURE EXPORT
recYCDOMOD0.CDOMODEVE = ""   'As String * 2                     ' EVENEMENT
recYCDOMOD0.CDOMODETA = ""   'As String * 2                     ' ETAT DE LA MODIF
recYCDOMOD0.CDOMODDMO = 0   'As Long                           ' DATE MODIFICATION
recYCDOMOD0.CDOMODDRM = 0   'As Long                           ' DATE RECEPTION MOD
recYCDOMOD0.CDOMODDP2 = ""   'As String * 32                    ' DESTIN.PAYS LIBELLE
recYCDOMOD0.CDOMODEP2 = ""   'As String * 32                    ' EMBARQ.PAYS LIBELLE
recYCDOMOD0.CDOMODPD2 = ""   'As String * 80                    ' PER.PRES.DOC.SUITE
recYCDOMOD0.CDOMODAUN = ""   'As String * 12                    ' CODE AUT. NOTIFIE
recYCDOMOD0.CDOMODCER = ""   'As String * 1                     ' COTAT°(O=CERTAIN/N)

End Sub






