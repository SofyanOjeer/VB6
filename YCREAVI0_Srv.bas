Attribute VB_Name = "srvYCREAVI0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCREAVI0Len = 500 ' 34 + ??????
Public Const recYCREAVI0_Block = 100 '????
Public Const constYCREAVI0 = "YCREAVI0"
Dim paramYCREAVI0_Import As String

Type typeYCREAVI0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CREAVIETA       As Long                           ' ETABLISSEMENT
    CREAVIAGE       As Long                           ' AGENCE
    CREAVISER       As String * 2                     ' SERVICE
    CREAVISSE       As String * 2                     ' SOUS-SERVICE
    CREAVIDOS       As Long                           ' N° DE DOSSIER
    CREAVIPRE       As Long                           ' N° DE PRET
    CREAVITYP       As String * 2                     ' TYPE EVNT
    CREAVINAC       As String * 3                     ' NATURE DU CREDIT
    CREAVILNC       As String * 30                    ' LIB NATUR CREDIT
    CREAVINAT       As String * 3                     ' NATURE DU PRET
    CREAVILNA       As String * 30                    ' LIB NATUR PRET
    CREAVICLI       As String * 7                     ' PAYEUR
    CREAVICET       As String * 4                     ' CODE ETAT
    CREAVILET       As String * 30                    ' LIBELLE CODE
    CREAVIRA1       As String * 32                    ' RAIS1
    CREAVIRA2       As String * 32                    ' RAIS2
    CREAVIAD1       As String * 32                    ' ADRESSE 1
    CREAVIAD2       As String * 32                    ' ADRESSE 2
    CREAVIAD3       As String * 32                    ' ADRESSE 3
    CREAVICOP       As String * 6                     ' CODE POSTAL
    CREAVIVIL       As String * 25                    ' VILLE
    CREAVIPAY       As String * 25                    ' PAYS
    CREAVIMOD       As String * 3                     ' MODE REGLEMENT
    CREAVILMO       As String * 12                    ' LIB MODE REGLT
    CREAVIPLA       As Long                           ' N° PLAN
    CREAVICOM       As String * 30                    ' COMPTE OU RIB
    CREAVIDEV       As String * 3                     ' DEVISE
    CREAVILDE       As String * 12                    ' LIBELLE DEVISE
    CREAVIPER       As String * 1                     ' PERIODICITE M S T A
    CREAVIREF       As String * 50                    ' REFERENCE EXTERNE
    CREAVIMDO       As Currency                       ' MT DU DOSSIER
    CREAVIDED       As String * 3                     ' DEVISE
    CREAVILDD       As String * 12                    ' LIBELLE DEVISE
    CREAVIMPR       As Currency                       ' MT DU PRET
    CREAVIDEP       As String * 3                     ' DEVISE
    CREAVILDP       As String * 12                    ' LIBELLE DEVISE
    CREAVIMON       As Currency                       ' MT:MAD,AMORT,ASS,COM
    CREAVIMIN       As Currency                       ' INTERETS
    CREAVITVA       As Currency                       ' MONTANT DE LA TVA
    CREAVITAU       As Double                         ' TAUX DU PRET
    CREAVICOT       As String * 6                     ' CODE TAUX
    CREAVIMAR       As Double                         ' MARGE
    CREAVITTV       As String * 6                     ' TX DE LA TVA
    CREAVIVTT       As Double                         ' VALEUR TX TVA
    CREAVIRGL       As Long                           ' DATE REGLEMENT
    CREAVIECH       As Long                           ' DATE ECHEANCE
    CREAVIDEB       As Long                           ' DATE DEBUT CALCUL
    CREAVIFIN       As Long                           ' DATE FIN CALCUL
    CREAVICC1       As String * 6                     ' CODE   COMMISSION
    CREAVICL1       As String * 30                    ' LIBEL. COMMISSION
    CREAVICM1       As Currency                       ' MT COMMISSION
    CREAVICS1       As String * 1                     ' A RECEVOIR
    CREAVICB1       As String * 2                     ' BASE COMMISSION
    CREAVIC11       As Double                         ' TAUX COMMISSION
    CREAVIC21       As String * 6                     ' CODE TAUX TVA
    CREAVIC31       As Double                         ' VALEUR TX TVA
    CREAVIC41       As Currency                       ' MT TVA
    CREAVICA1       As Currency                       ' ASSIETTE
    CREAVICC2       As String * 6                     ' CODE   COMMISSION
    CREAVICL2       As String * 30                    ' LIBEL. COMMISSION
    CREAVICM2       As Currency                       ' MT COMMISSION
    CREAVICS2       As String * 1                     ' A RECEVOIR
    CREAVICB2       As String * 2                     ' BASE COMMISSION
    CREAVIC12       As Double                         ' TAUX COMMISSION
    CREAVIC22       As String * 6                     ' CODE TAUX TVA
    CREAVIC32       As Double                         ' VALEUR TX TVA
    CREAVIC42       As Currency                       ' MT TVA
    CREAVICA2       As Currency                       ' ASSIETTE
    CREAVIAC1       As String * 6                     ' CODE   ASSURANCE
    CREAVIAL1       As String * 30                    ' LIBEL. ASSURANCE
    CREAVIAM1       As Currency                       ' MT ASSURANCE
    CREAVIAB1       As String * 2                     ' BASE ASSURANCE
    CREAVIA11       As Double                         ' TAUX ASSURANCE
    CREAVIA21       As String * 6                     ' CODE TAUX TVA
    CREAVIA31       As Double                         ' VALEUR TX TVA
    CREAVIA41       As Currency                       ' MT TVA
    CREAVIAA1       As Currency                       ' ASSIETTE
    CREAVIAC2       As String * 6                     ' CODE   ASSURANCE
    CREAVIAL2       As String * 30                    ' LIBEL. ASSURANCE
    CREAVIAM2       As Currency                       ' MT ASSURANCE
    CREAVIAB2       As String * 2                     ' BASE ASSURANCE
    CREAVIA12       As Double                         ' TAUX ASSURANCE
    CREAVIA22       As String * 6                     ' CODE TAUX TVA
    CREAVIA32       As Double                         ' VALEUR TX TVA
    CREAVIA42       As Currency                       ' MT TVA
    CREAVIAA2       As Currency                       ' ASSIETTE
    CREAVIAC3       As String * 6                     ' CODE   ASSURANCE
    CREAVIAL3       As String * 30                    ' LIBEL. ASSURANCE
    CREAVIAM3       As Currency                       ' MT ASSURANCE
    CREAVIAB3       As String * 2                     ' BASE ASSURANCE
    CREAVIA13       As Double                         ' TAUX ASSURANCE
    CREAVIA23       As String * 6                     ' CODE TAUX TVA
    CREAVIA33       As Double                         ' VALEUR TX TVA
    CREAVIA43       As Currency                       ' MT TVA
    CREAVIAA3       As Currency                       ' ASSIETTE
    CREAVIAC4       As String * 6                     ' CODE   ASSURANCE
    CREAVIAL4       As String * 30                    ' LIBEL. ASSURANCE
    CREAVIAM4       As Currency                       ' MT ASSURANCE
    CREAVIAB4       As String * 2                     ' BASE ASSURANCE
    CREAVIA14       As Double                         ' TAUX ASSURANCE
    CREAVIA24       As String * 6                     ' CODE TAUX TVA
    CREAVIA34       As Double                         ' VALEUR TX TVA
    CREAVIA44       As Currency                       ' MT TVA
    CREAVIAA4       As Currency                       ' ASSIETTE
    CREAVIAC5       As String * 6                     ' CODE   ASSURANCE
    CREAVIAL5       As String * 30                    ' LIBEL. ASSURANCE
    CREAVIAM5       As Currency                       ' MT ASSURANCE
    CREAVIAB5       As String * 2                     ' BASE ASSURANCE
    CREAVIA15       As Double                         ' TAUX ASSURANCE
    CREAVIA25       As String * 6                     ' CODE TAUX TVA
    CREAVIA35       As Double                         ' VALEUR TX TVA
    CREAVIA45       As Currency                       ' MT TVA
    CREAVIAA5       As Currency                       ' ASSIETTE
    CREAVINET       As Currency                       ' MT REGLE
    CREAVIDAT       As Long                           ' DATE AVIS = 0
    CREAVICRD       As Currency                       ' CRD AVT ECHEANCE
    CREAVINUM       As Integer                        ' NUMERO ECHEANCE
    CREAVITYC       As String * 1                     ' TYPE DE CREDIT
    CREAVIMDR       As Currency                       ' MT REGLE EN DER
    CREAVIPRC       As String * 1                     ' PRECOMPTE O/N
    CREAVICOU       As Double                         ' COURS
    CREAVITEL       As String * 20                    ' N° TEL
    CREAVIFAX       As String * 20                    ' N° FAX
    CREAVICRP       As Currency                       ' CRD APR ECHEANCE
    CREAVIAUT       As String * 12                    ' CODE AUTO
    CREAVINPL       As Long                           ' NUMERO PLAN
    CREAVIPAL       As Long                           ' NUMERO PALIER
    CREAVINEC       As Long                           ' NUMERO ECHEANCE
    CREAVIITC       As Currency                       ' INT REPORTES PAYES
    CREAVISE1       As Long                           ' SEQUENCE 1
    CREAVISE2       As Long                           ' SEQUENCE 2
    CREAVIDTC       As Long                           ' DATE CREATION AVIS
    
    
End Type
    
'---------------------------------------------------------
Public Function srvYCREAVI0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCREAVI0 As typeYCREAVI0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCREAVI0_GetBuffer_ODBC = Null

 '   recYCREAVI0.CREAVIETA = rsADO("CREAVIxxx")    'rsADO("CREAVIxxx") 'CInt(Val(mId$(MsgTxt, K + 1, 5)))

    recYCREAVI0.CREAVIETA = rsADO("CREAVIETA")    'CLng(Val(mId$(MsgTxt, K + 1, 5)))
    recYCREAVI0.CREAVIAGE = rsADO("CREAVIAGE")    'CLng(Val(mId$(MsgTxt, K + 6, 5)))
    recYCREAVI0.CREAVISER = rsADO("CREAVISER")    'mId$(MsgTxt, K + 11, 2)
    recYCREAVI0.CREAVISSE = rsADO("CREAVISSE")    'mId$(MsgTxt, K + 13, 2)
    recYCREAVI0.CREAVIDOS = rsADO("CREAVIDOS")    'CLng(Val(mId$(MsgTxt, K + 15, 8)))
    recYCREAVI0.CREAVIPRE = rsADO("CREAVIPRE")    'CLng(Val(mId$(MsgTxt, K + 23, 4)))
    recYCREAVI0.CREAVITYP = rsADO("CREAVITYP")    'mId$(MsgTxt, K + 27, 2)
    recYCREAVI0.CREAVINAC = rsADO("CREAVINAC")    'mId$(MsgTxt, K + 29, 3)
    recYCREAVI0.CREAVILNC = rsADO("CREAVILNC")    'mId$(MsgTxt, K + 32, 30)
    recYCREAVI0.CREAVINAT = rsADO("CREAVINAT")    'mId$(MsgTxt, K + 62, 3)
    recYCREAVI0.CREAVILNA = rsADO("CREAVILNA")    'mId$(MsgTxt, K + 65, 30)
    recYCREAVI0.CREAVICLI = rsADO("CREAVICLI")    'mId$(MsgTxt, K + 95, 7)
    recYCREAVI0.CREAVICET = rsADO("CREAVICET")    'mId$(MsgTxt, K + 102, 4)
    recYCREAVI0.CREAVILET = rsADO("CREAVILET")    'mId$(MsgTxt, K + 106, 30)
    recYCREAVI0.CREAVIRA1 = rsADO("CREAVIRA1")    'mId$(MsgTxt, K + 136, 32)
    recYCREAVI0.CREAVIRA2 = rsADO("CREAVIRA2")    'mId$(MsgTxt, K + 168, 32)
    recYCREAVI0.CREAVIAD1 = rsADO("CREAVIAD1")    'mId$(MsgTxt, K + 200, 32)
    recYCREAVI0.CREAVIAD2 = rsADO("CREAVIAD2")    'mId$(MsgTxt, K + 232, 32)
    recYCREAVI0.CREAVIAD3 = rsADO("CREAVIAD3")    'mId$(MsgTxt, K + 264, 32)
    recYCREAVI0.CREAVICOP = rsADO("CREAVICOP")    'mId$(MsgTxt, K + 296, 6)
    recYCREAVI0.CREAVIVIL = rsADO("CREAVIVIL")    'mId$(MsgTxt, K + 302, 25)
    recYCREAVI0.CREAVIPAY = rsADO("CREAVIPAY")    'mId$(MsgTxt, K + 327, 25)
    recYCREAVI0.CREAVIMOD = rsADO("CREAVIMOD")    'mId$(MsgTxt, K + 352, 3)
    recYCREAVI0.CREAVILMO = rsADO("CREAVILMO")    'mId$(MsgTxt, K + 355, 12)
    recYCREAVI0.CREAVIPLA = rsADO("CREAVIPLA")    'CLng(Val(mId$(MsgTxt, K + 367, 4)))
    recYCREAVI0.CREAVICOM = rsADO("CREAVICOM")    'mId$(MsgTxt, K + 371, 30)
    recYCREAVI0.CREAVIDEV = rsADO("CREAVIDEV")    'mId$(MsgTxt, K + 401, 3)
    recYCREAVI0.CREAVILDE = rsADO("CREAVILDE")    'mId$(MsgTxt, K + 404, 12)
    recYCREAVI0.CREAVIPER = rsADO("CREAVIPER")    'mId$(MsgTxt, K + 416, 1)
    recYCREAVI0.CREAVIREF = rsADO("CREAVIREF")    'mId$(MsgTxt, K + 417, 50)
    recYCREAVI0.CREAVIMDO = rsADO("CREAVIMDO")    'CCur(Val(mId$(MsgTxt, K + 467, 16))) / 100
    recYCREAVI0.CREAVIDED = rsADO("CREAVIDED")    'mId$(MsgTxt, K + 483, 3)
    recYCREAVI0.CREAVILDD = rsADO("CREAVILDD")    'mId$(MsgTxt, K + 486, 12)
    recYCREAVI0.CREAVIMPR = rsADO("CREAVIMPR")    'CCur(Val(mId$(MsgTxt, K + 498, 16))) / 100
    recYCREAVI0.CREAVIDEP = rsADO("CREAVIDEP")    'mId$(MsgTxt, K + 514, 3)
    recYCREAVI0.CREAVILDP = rsADO("CREAVILDP")    'mId$(MsgTxt, K + 517, 12)
    recYCREAVI0.CREAVIMON = rsADO("CREAVIMON")    'CCur(Val(mId$(MsgTxt, K + 529, 16))) / 100
    recYCREAVI0.CREAVIMIN = rsADO("CREAVIMIN")    'CCur(Val(mId$(MsgTxt, K + 545, 16))) / 100
    recYCREAVI0.CREAVITVA = rsADO("CREAVITVA")    'CCur(Val(mId$(MsgTxt, K + 561, 16))) / 100
    recYCREAVI0.CREAVITAU = rsADO("CREAVITAU")    'CDbl(Val(mId$(MsgTxt, K + 577, 13))) / 1000000000
    recYCREAVI0.CREAVICOT = rsADO("CREAVICOT")    'mId$(MsgTxt, K + 590, 6)
    recYCREAVI0.CREAVIMAR = rsADO("CREAVIMAR")    'CDbl(Val(mId$(MsgTxt, K + 596, 13))) / 1000000000
    recYCREAVI0.CREAVITTV = rsADO("CREAVITTV")    'mId$(MsgTxt, K + 609, 6)
    recYCREAVI0.CREAVIVTT = rsADO("CREAVIVTT")    'CDbl(Val(mId$(MsgTxt, K + 615, 13))) / 1000000000
    recYCREAVI0.CREAVIRGL = rsADO("CREAVIRGL")    'CLng(Val(mId$(MsgTxt, K + 628, 8)))
    recYCREAVI0.CREAVIECH = rsADO("CREAVIECH")    'CLng(Val(mId$(MsgTxt, K + 636, 8)))
    recYCREAVI0.CREAVIDEB = rsADO("CREAVIDEB")    'CLng(Val(mId$(MsgTxt, K + 644, 8)))
    recYCREAVI0.CREAVIFIN = rsADO("CREAVIFIN")    'CLng(Val(mId$(MsgTxt, K + 652, 8)))
    recYCREAVI0.CREAVICC1 = rsADO("CREAVICC1")    'mId$(MsgTxt, K + 660, 6)
    recYCREAVI0.CREAVICL1 = rsADO("CREAVICL1")    'mId$(MsgTxt, K + 666, 30)
    recYCREAVI0.CREAVICM1 = rsADO("CREAVICM1")    'CCur(Val(mId$(MsgTxt, K + 696, 16))) / 100
    recYCREAVI0.CREAVICS1 = rsADO("CREAVICS1")    'mId$(MsgTxt, K + 712, 1)
    recYCREAVI0.CREAVICB1 = rsADO("CREAVICB1")    'mId$(MsgTxt, K + 713, 2)
    recYCREAVI0.CREAVIC11 = rsADO("CREAVIC11")    'CDbl(Val(mId$(MsgTxt, K + 715, 8))) / 100000
    recYCREAVI0.CREAVIC21 = rsADO("CREAVIC21")    'mId$(MsgTxt, K + 723, 6)
    recYCREAVI0.CREAVIC31 = rsADO("CREAVIC31")    'CDbl(Val(mId$(MsgTxt, K + 729, 13))) / 1000000000
    recYCREAVI0.CREAVIC41 = rsADO("CREAVIC41")    'CCur(Val(mId$(MsgTxt, K + 742, 16))) / 100
    recYCREAVI0.CREAVICA1 = rsADO("CREAVICA1")    'CCur(Val(mId$(MsgTxt, K + 758, 16))) / 100
    recYCREAVI0.CREAVICC2 = rsADO("CREAVICC2")    'mId$(MsgTxt, K + 774, 6)
    recYCREAVI0.CREAVICL2 = rsADO("CREAVICL2")    'mId$(MsgTxt, K + 780, 30)
    recYCREAVI0.CREAVICM2 = rsADO("CREAVICM2")    'CCur(Val(mId$(MsgTxt, K + 810, 16))) / 100
    recYCREAVI0.CREAVICS2 = rsADO("CREAVICS2")    'mId$(MsgTxt, K + 826, 1)
    recYCREAVI0.CREAVICB2 = rsADO("CREAVICB2")    'mId$(MsgTxt, K + 827, 2)
    recYCREAVI0.CREAVIC12 = rsADO("CREAVIC12")    'CDbl(Val(mId$(MsgTxt, K + 829, 8))) / 100000
    recYCREAVI0.CREAVIC22 = rsADO("CREAVIC22")    'mId$(MsgTxt, K + 837, 6)
    recYCREAVI0.CREAVIC32 = rsADO("CREAVIC32")    'CDbl(Val(mId$(MsgTxt, K + 843, 13))) / 1000000000
    recYCREAVI0.CREAVIC42 = rsADO("CREAVIC42")    'CCur(Val(mId$(MsgTxt, K + 856, 16))) / 100
    recYCREAVI0.CREAVICA2 = rsADO("CREAVICA2")    'CCur(Val(mId$(MsgTxt, K + 872, 16))) / 100
    recYCREAVI0.CREAVIAC1 = rsADO("CREAVIAC1")    'mId$(MsgTxt, K + 888, 6)
    recYCREAVI0.CREAVIAL1 = rsADO("CREAVIAL1")    'mId$(MsgTxt, K + 894, 30)
    recYCREAVI0.CREAVIAM1 = rsADO("CREAVIAM1")    'CCur(Val(mId$(MsgTxt, K + 924, 16))) / 100
    recYCREAVI0.CREAVIAB1 = rsADO("CREAVIAB1")    'mId$(MsgTxt, K + 940, 2)
    recYCREAVI0.CREAVIA11 = rsADO("CREAVIA11")    'CDbl(Val(mId$(MsgTxt, K + 942, 8))) / 100000
    recYCREAVI0.CREAVIA21 = rsADO("CREAVIA21")    'mId$(MsgTxt, K + 950, 6)
    recYCREAVI0.CREAVIA31 = rsADO("CREAVIA31")    'CDbl(Val(mId$(MsgTxt, K + 956, 13))) / 1000000000
    recYCREAVI0.CREAVIA41 = rsADO("CREAVIA41")    'CCur(Val(mId$(MsgTxt, K + 969, 16))) / 100
    recYCREAVI0.CREAVIAA1 = rsADO("CREAVIAA1")    'CCur(Val(mId$(MsgTxt, K + 985, 16))) / 100
    recYCREAVI0.CREAVIAC2 = rsADO("CREAVIAC2")    'mId$(MsgTxt, K + 1001, 6)
    recYCREAVI0.CREAVIAL2 = rsADO("CREAVIAL2")    'mId$(MsgTxt, K + 1007, 30)
    recYCREAVI0.CREAVIAM2 = rsADO("CREAVIAM2")    'CCur(Val(mId$(MsgTxt, K + 1037, 16))) / 100
    recYCREAVI0.CREAVIAB2 = rsADO("CREAVIAB2")    'mId$(MsgTxt, K + 1053, 2)
    recYCREAVI0.CREAVIA12 = rsADO("CREAVIA12")    'CDbl(Val(mId$(MsgTxt, K + 1055, 8))) / 100000
    recYCREAVI0.CREAVIA22 = rsADO("CREAVIA22")    'mId$(MsgTxt, K + 1063, 6)
    recYCREAVI0.CREAVIA32 = rsADO("CREAVIA32")    'CDbl(Val(mId$(MsgTxt, K + 1069, 13))) / 1000000000
    recYCREAVI0.CREAVIA42 = rsADO("CREAVIA42")    'CCur(Val(mId$(MsgTxt, K + 1082, 16))) / 100
    recYCREAVI0.CREAVIAA2 = rsADO("CREAVIAA2")    'CCur(Val(mId$(MsgTxt, K + 1098, 16))) / 100
    recYCREAVI0.CREAVIAC3 = rsADO("CREAVIAC3")    'mId$(MsgTxt, K + 1114, 6)
    recYCREAVI0.CREAVIAL3 = rsADO("CREAVIAL3")    'mId$(MsgTxt, K + 1120, 30)
    recYCREAVI0.CREAVIAM3 = rsADO("CREAVIAM3")    'CCur(Val(mId$(MsgTxt, K + 1150, 16))) / 100
    recYCREAVI0.CREAVIAB3 = rsADO("CREAVIAB3")    'mId$(MsgTxt, K + 1166, 2)
    recYCREAVI0.CREAVIA13 = rsADO("CREAVIA13")    'CDbl(Val(mId$(MsgTxt, K + 1168, 8))) / 100000
    recYCREAVI0.CREAVIA23 = rsADO("CREAVIA23")    'mId$(MsgTxt, K + 1176, 6)
    recYCREAVI0.CREAVIA33 = rsADO("CREAVIA33")    'CDbl(Val(mId$(MsgTxt, K + 1182, 13))) / 1000000000
    recYCREAVI0.CREAVIA43 = rsADO("CREAVIA43")    'CCur(Val(mId$(MsgTxt, K + 1195, 16))) / 100
    recYCREAVI0.CREAVIAA3 = rsADO("CREAVIAA3")    'CCur(Val(mId$(MsgTxt, K + 1211, 16))) / 100
    recYCREAVI0.CREAVIAC4 = rsADO("CREAVIAC4")    'mId$(MsgTxt, K + 1227, 6)
    recYCREAVI0.CREAVIAL4 = rsADO("CREAVIAL4")    'mId$(MsgTxt, K + 1233, 30)
    recYCREAVI0.CREAVIAM4 = rsADO("CREAVIAM4")    'CCur(Val(mId$(MsgTxt, K + 1263, 16))) / 100
    recYCREAVI0.CREAVIAB4 = rsADO("CREAVIAB4")    'mId$(MsgTxt, K + 1279, 2)
    recYCREAVI0.CREAVIA14 = rsADO("CREAVIA14")    'CDbl(Val(mId$(MsgTxt, K + 1281, 8))) / 100000
    recYCREAVI0.CREAVIA24 = rsADO("CREAVIA24")    'mId$(MsgTxt, K + 1289, 6)
    recYCREAVI0.CREAVIA34 = rsADO("CREAVIA34")    'CDbl(Val(mId$(MsgTxt, K + 1295, 13))) / 1000000000
    recYCREAVI0.CREAVIA44 = rsADO("CREAVIA44")    'CCur(Val(mId$(MsgTxt, K + 1308, 16))) / 100
    recYCREAVI0.CREAVIAA4 = rsADO("CREAVIAA4")    'CCur(Val(mId$(MsgTxt, K + 1324, 16))) / 100
    recYCREAVI0.CREAVIAC5 = rsADO("CREAVIAC5")    'mId$(MsgTxt, K + 1340, 6)
    recYCREAVI0.CREAVIAL5 = rsADO("CREAVIAL5")    'mId$(MsgTxt, K + 1346, 30)
    recYCREAVI0.CREAVIAM5 = rsADO("CREAVIAM5")    'CCur(Val(mId$(MsgTxt, K + 1376, 16))) / 100
    recYCREAVI0.CREAVIAB5 = rsADO("CREAVIAB5")    'mId$(MsgTxt, K + 1392, 2)
    recYCREAVI0.CREAVIA15 = rsADO("CREAVIA15")    'CDbl(Val(mId$(MsgTxt, K + 1394, 8))) / 100000
    recYCREAVI0.CREAVIA25 = rsADO("CREAVIA25")    'mId$(MsgTxt, K + 1402, 6)
    recYCREAVI0.CREAVIA35 = rsADO("CREAVIA35")    'CDbl(Val(mId$(MsgTxt, K + 1408, 13))) / 1000000000
    recYCREAVI0.CREAVIA45 = rsADO("CREAVIA45")    'CCur(Val(mId$(MsgTxt, K + 1421, 16))) / 100
    recYCREAVI0.CREAVIAA5 = rsADO("CREAVIAA5")    'CCur(Val(mId$(MsgTxt, K + 1437, 16))) / 100
    recYCREAVI0.CREAVINET = rsADO("CREAVINET")    'CCur(Val(mId$(MsgTxt, K + 1453, 16))) / 100
    recYCREAVI0.CREAVIDAT = rsADO("CREAVIDAT")    'CLng(Val(mId$(MsgTxt, K + 1469, 8)))
    recYCREAVI0.CREAVICRD = rsADO("CREAVICRD")    'CCur(Val(mId$(MsgTxt, K + 1477, 16))) / 100
    recYCREAVI0.CREAVINUM = rsADO("CREAVINUM")    'CInt(Val(mId$(MsgTxt, K + 1493, 5)))
    recYCREAVI0.CREAVITYC = rsADO("CREAVITYC")    'mId$(MsgTxt, K + 1498, 1)
    recYCREAVI0.CREAVIMDR = rsADO("CREAVIMDR")    'CCur(Val(mId$(MsgTxt, K + 1499, 16))) / 100
    recYCREAVI0.CREAVIPRC = rsADO("CREAVIPRC")    'mId$(MsgTxt, K + 1515, 1)
    recYCREAVI0.CREAVICOU = rsADO("CREAVICOU")    'CDbl(Val(mId$(MsgTxt, K + 1516, 16))) / 10000000000#
    recYCREAVI0.CREAVITEL = rsADO("CREAVITEL")    'mId$(MsgTxt, K + 1532, 20)
    recYCREAVI0.CREAVIFAX = rsADO("CREAVIFAX")    'mId$(MsgTxt, K + 1552, 20)
    recYCREAVI0.CREAVICRP = rsADO("CREAVICRP")    'CCur(Val(mId$(MsgTxt, K + 1572, 16))) / 100
    recYCREAVI0.CREAVIAUT = rsADO("CREAVIAUT")    'mId$(MsgTxt, K + 1588, 12)
    recYCREAVI0.CREAVINPL = rsADO("CREAVINPL")    'CLng(Val(mId$(MsgTxt, K + 1600, 4)))
    recYCREAVI0.CREAVIPAL = rsADO("CREAVIPAL")    'CLng(Val(mId$(MsgTxt, K + 1604, 4)))
    recYCREAVI0.CREAVINEC = rsADO("CREAVINEC")    'CLng(Val(mId$(MsgTxt, K + 1608, 4)))
    recYCREAVI0.CREAVIITC = rsADO("CREAVIITC")    'CCur(Val(mId$(MsgTxt, K + 1612, 16))) / 100
    recYCREAVI0.CREAVISE1 = rsADO("CREAVISE1")    'CLng(Val(mId$(MsgTxt, K + 1628, 4)))
    recYCREAVI0.CREAVISE2 = rsADO("CREAVISE2")    'CLng(Val(mId$(MsgTxt, K + 1632, 4)))
    recYCREAVI0.CREAVIDTC = rsADO("CREAVIDTC")    'CLng(Val(mId$(MsgTxt, K + 1636, 8)))


Exit Function

Error_Handler:
srvYCREAVI0_GetBuffer_ODBC = Error

End Function


'---------------------------------------------------------
Public Sub recYCREAVI0_Init(recYCREAVI0 As typeYCREAVI0)
'---------------------------------------------------------
'MsgTxt = Space$(recYCREAVI0Len)
'MsgTxtIndex = 0
'Call srvYCREAVI0_GetBuffer(recYCREAVI0)
recYCREAVI0.obj = "ZCREAVI0_S"

recYCREAVI0.CREAVIETA = 0  'As Long                           ' ETABLISSEMENT
recYCREAVI0.CREAVIAGE = 0  'As Long                           ' AGENCE
recYCREAVI0.CREAVISER = ""  'As string * 2                     ' SERVICE
recYCREAVI0.CREAVISSE = ""  'As string * 2                     ' SOUS-SERVICE
recYCREAVI0.CREAVIDOS = 0  'As Long                           ' N° DE DOSSIER
recYCREAVI0.CREAVIPRE = 0  'As Long                           ' N° DE PRET
recYCREAVI0.CREAVITYP = ""  'As string * 2                     ' TYPE EVNT
recYCREAVI0.CREAVINAC = ""  'As string * 3                     ' NATURE DU CREDIT
recYCREAVI0.CREAVILNC = ""  'As string * 30                    ' LIB NATUR CREDIT
recYCREAVI0.CREAVINAT = ""  'As string * 3                     ' NATURE DU PRET
recYCREAVI0.CREAVILNA = ""  'As string * 30                    ' LIB NATUR PRET
recYCREAVI0.CREAVICLI = ""  'As string * 7                     ' PAYEUR
recYCREAVI0.CREAVICET = ""  'As string * 4                     ' CODE ETAT
recYCREAVI0.CREAVILET = ""  'As string * 30                    ' LIBELLE CODE
recYCREAVI0.CREAVIRA1 = ""  'As string * 32                    ' RAIS1
recYCREAVI0.CREAVIRA2 = ""  'As string * 32                    ' RAIS2
recYCREAVI0.CREAVIAD1 = ""  'As string * 32                    ' ADRESSE 1
recYCREAVI0.CREAVIAD2 = ""  'As string * 32                    ' ADRESSE 2
recYCREAVI0.CREAVIAD3 = ""  'As string * 32                    ' ADRESSE 3
recYCREAVI0.CREAVICOP = ""  'As string * 6                     ' CODE POSTAL
recYCREAVI0.CREAVIVIL = ""  'As string * 25                    ' VILLE
recYCREAVI0.CREAVIPAY = ""  'As string * 25                    ' PAYS
recYCREAVI0.CREAVIMOD = ""  'As string * 3                     ' MODE REGLEMENT
recYCREAVI0.CREAVILMO = ""  'As string * 12                    ' LIB MODE REGLT
recYCREAVI0.CREAVIPLA = 0  'As Long                           ' N° PLAN
recYCREAVI0.CREAVICOM = ""  'As string * 30                    ' COMPTE OU RIB
recYCREAVI0.CREAVIDEV = ""  'As string * 3                     ' DEVISE
recYCREAVI0.CREAVILDE = ""  'As string * 12                    ' LIBELLE DEVISE
recYCREAVI0.CREAVIPER = ""  'As string * 1                     ' PERIODICITE M S T A
recYCREAVI0.CREAVIREF = ""  'As string * 50                    ' REFERENCE EXTERNE
recYCREAVI0.CREAVIMDO = 0  'As Currency                       ' MT DU DOSSIER
recYCREAVI0.CREAVIDED = ""  'As string * 3                     ' DEVISE
recYCREAVI0.CREAVILDD = ""  'As string * 12                    ' LIBELLE DEVISE
recYCREAVI0.CREAVIMPR = 0  'As Currency                       ' MT DU PRET
recYCREAVI0.CREAVIDEP = ""  'As string * 3                     ' DEVISE
recYCREAVI0.CREAVILDP = ""  'As string * 12                    ' LIBELLE DEVISE
recYCREAVI0.CREAVIMON = 0  'As Currency                       ' MT:MAD,AMORT,ASS,COM
recYCREAVI0.CREAVIMIN = 0  'As Currency                       ' INTERETS
recYCREAVI0.CREAVITVA = 0  'As Currency                       ' MONTANT DE LA TVA
recYCREAVI0.CREAVITAU = 0  'As Double                         ' TAUX DU PRET
recYCREAVI0.CREAVICOT = ""  'As string * 6                     ' CODE TAUX
recYCREAVI0.CREAVIMAR = 0  'As Double                         ' MARGE
recYCREAVI0.CREAVITTV = ""  'As string * 6                     ' TX DE LA TVA
recYCREAVI0.CREAVIVTT = 0  'As Double                         ' VALEUR TX TVA
recYCREAVI0.CREAVIRGL = 0  'As Long                           ' DATE REGLEMENT
recYCREAVI0.CREAVIECH = 0  'As Long                           ' DATE ECHEANCE
recYCREAVI0.CREAVIDEB = 0  'As Long                           ' DATE DEBUT CALCUL
recYCREAVI0.CREAVIFIN = 0  'As Long                           ' DATE FIN CALCUL
recYCREAVI0.CREAVICC1 = ""  'As string * 6                     ' CODE   COMMISSION
recYCREAVI0.CREAVICL1 = ""  'As string * 30                    ' LIBEL. COMMISSION
recYCREAVI0.CREAVICM1 = 0  'As Currency                       ' MT COMMISSION
recYCREAVI0.CREAVICS1 = ""  'As string * 1                     ' A RECEVOIR
recYCREAVI0.CREAVICB1 = ""  'As string * 2                     ' BASE COMMISSION
recYCREAVI0.CREAVIC11 = 0  'As Double                         ' TAUX COMMISSION
recYCREAVI0.CREAVIC21 = ""  'As string * 6                     ' CODE TAUX TVA
recYCREAVI0.CREAVIC31 = 0  'As Double                         ' VALEUR TX TVA
recYCREAVI0.CREAVIC41 = 0  'As Currency                       ' MT TVA
recYCREAVI0.CREAVICA1 = 0  'As Currency                       ' ASSIETTE
recYCREAVI0.CREAVICC2 = ""  'As string * 6                     ' CODE   COMMISSION
recYCREAVI0.CREAVICL2 = ""  'As string * 30                    ' LIBEL. COMMISSION
recYCREAVI0.CREAVICM2 = 0  'As Currency                       ' MT COMMISSION
recYCREAVI0.CREAVICS2 = ""  'As string * 1                     ' A RECEVOIR
recYCREAVI0.CREAVICB2 = ""  'As string * 2                     ' BASE COMMISSION
recYCREAVI0.CREAVIC12 = 0  'As Double                         ' TAUX COMMISSION
recYCREAVI0.CREAVIC22 = ""  'As string * 6                     ' CODE TAUX TVA
recYCREAVI0.CREAVIC32 = 0  'As Double                         ' VALEUR TX TVA
recYCREAVI0.CREAVIC42 = 0  'As Currency                       ' MT TVA
recYCREAVI0.CREAVICA2 = 0  'As Currency                       ' ASSIETTE
recYCREAVI0.CREAVIAC1 = ""  'As string * 6                     ' CODE   ASSURANCE
recYCREAVI0.CREAVIAL1 = ""  'As string * 30                    ' LIBEL. ASSURANCE
recYCREAVI0.CREAVIAM1 = 0  'As Currency                       ' MT ASSURANCE
recYCREAVI0.CREAVIAB1 = ""  'As string * 2                     ' BASE ASSURANCE
recYCREAVI0.CREAVIA11 = 0  'As Double                         ' TAUX ASSURANCE
recYCREAVI0.CREAVIA21 = ""  'As string * 6                     ' CODE TAUX TVA
recYCREAVI0.CREAVIA31 = 0  'As Double                         ' VALEUR TX TVA
recYCREAVI0.CREAVIA41 = 0  'As Currency                       ' MT TVA
recYCREAVI0.CREAVIAA1 = 0  'As Currency                       ' ASSIETTE
recYCREAVI0.CREAVIAC2 = ""  'As string * 6                     ' CODE   ASSURANCE
recYCREAVI0.CREAVIAL2 = ""  'As string * 30                    ' LIBEL. ASSURANCE
recYCREAVI0.CREAVIAM2 = 0  'As Currency                       ' MT ASSURANCE
recYCREAVI0.CREAVIAB2 = ""  'As string * 2                     ' BASE ASSURANCE
recYCREAVI0.CREAVIA12 = 0  'As Double                         ' TAUX ASSURANCE
recYCREAVI0.CREAVIA22 = ""  'As string * 6                     ' CODE TAUX TVA
recYCREAVI0.CREAVIA32 = 0  'As Double                         ' VALEUR TX TVA
recYCREAVI0.CREAVIA42 = 0  'As Currency                       ' MT TVA
recYCREAVI0.CREAVIAA2 = 0  'As Currency                       ' ASSIETTE
recYCREAVI0.CREAVIAC3 = ""  'As string * 6                     ' CODE   ASSURANCE
recYCREAVI0.CREAVIAL3 = ""  'As string * 30                    ' LIBEL. ASSURANCE
recYCREAVI0.CREAVIAM3 = 0  'As Currency                       ' MT ASSURANCE
recYCREAVI0.CREAVIAB3 = ""  'As string * 2                     ' BASE ASSURANCE
recYCREAVI0.CREAVIA13 = 0  'As Double                         ' TAUX ASSURANCE
recYCREAVI0.CREAVIA23 = ""  'As string * 6                     ' CODE TAUX TVA
recYCREAVI0.CREAVIA33 = 0  'As Double                         ' VALEUR TX TVA
recYCREAVI0.CREAVIA43 = 0  'As Currency                       ' MT TVA
recYCREAVI0.CREAVIAA3 = 0  'As Currency                       ' ASSIETTE
recYCREAVI0.CREAVIAC4 = ""  'As string * 6                     ' CODE   ASSURANCE
recYCREAVI0.CREAVIAL4 = ""  'As string * 30                    ' LIBEL. ASSURANCE
recYCREAVI0.CREAVIAM4 = 0  'As Currency                       ' MT ASSURANCE
recYCREAVI0.CREAVIAB4 = ""  'As string * 2                     ' BASE ASSURANCE
recYCREAVI0.CREAVIA14 = 0  'As Double                         ' TAUX ASSURANCE
recYCREAVI0.CREAVIA24 = ""  'As string * 6                     ' CODE TAUX TVA
recYCREAVI0.CREAVIA34 = 0  'As Double                         ' VALEUR TX TVA
recYCREAVI0.CREAVIA44 = 0  'As Currency                       ' MT TVA
recYCREAVI0.CREAVIAA4 = 0  'As Currency                       ' ASSIETTE
recYCREAVI0.CREAVIAC5 = ""  'As string * 6                     ' CODE   ASSURANCE
recYCREAVI0.CREAVIAL5 = ""  'As string * 30                    ' LIBEL. ASSURANCE
recYCREAVI0.CREAVIAM5 = 0  'As Currency                       ' MT ASSURANCE
recYCREAVI0.CREAVIAB5 = ""  'As string * 2                     ' BASE ASSURANCE
recYCREAVI0.CREAVIA15 = 0  'As Double                         ' TAUX ASSURANCE
recYCREAVI0.CREAVIA25 = ""  'As string * 6                     ' CODE TAUX TVA
recYCREAVI0.CREAVIA35 = 0  'As Double                         ' VALEUR TX TVA
recYCREAVI0.CREAVIA45 = 0  'As Currency                       ' MT TVA
recYCREAVI0.CREAVIAA5 = 0  'As Currency                       ' ASSIETTE
recYCREAVI0.CREAVINET = 0  'As Currency                       ' MT REGLE
recYCREAVI0.CREAVIDAT = 0  'As Long                           ' DATE AVIS = 0
recYCREAVI0.CREAVICRD = 0  'As Currency                       ' CRD AVT ECHEANCE
recYCREAVI0.CREAVINUM = 0  'As Integer                        ' NUMERO ECHEANCE
recYCREAVI0.CREAVITYC = ""  'As string * 1                     ' TYPE DE CREDIT
recYCREAVI0.CREAVIMDR = 0  'As Currency                       ' MT REGLE EN DER
recYCREAVI0.CREAVIPRC = ""  'As string * 1                     ' PRECOMPTE O/N
recYCREAVI0.CREAVICOU = 0  'As Double                         ' COURS
recYCREAVI0.CREAVITEL = ""  'As string * 20                    ' N° TEL
recYCREAVI0.CREAVIFAX = ""  'As string * 20                    ' N° FAX
recYCREAVI0.CREAVICRP = 0  'As Currency                       ' CRD APR ECHEANCE
recYCREAVI0.CREAVIAUT = ""  'As string * 12                    ' CODE AUTO
recYCREAVI0.CREAVINPL = 0  'As Long                           ' NUMERO PLAN
recYCREAVI0.CREAVIPAL = 0  'As Long                           ' NUMERO PALIER
recYCREAVI0.CREAVINEC = 0  'As Long                           ' NUMERO ECHEANCE
recYCREAVI0.CREAVIITC = 0  'As Currency                       ' INT REPORTES PAYES
recYCREAVI0.CREAVISE1 = 0  'As Long                           ' SEQUENCE 1
recYCREAVI0.CREAVISE2 = 0  'As Long                           ' SEQUENCE 2
recYCREAVI0.CREAVIDTC = 0  'As Long                           ' DATE CREATION AVIS



End Sub


Public Sub srvYCREAVI0_ElpDisplay(recYCREAVI0 As typeYCREAVI0)
frmElpDisplay.fgData.Rows = 133
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIETA    4S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAGE    4S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVISER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVISER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVISSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVISSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIDOS    7S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° DE DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIDOS
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIPRE    3S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° DE PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIPRE
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVITYP    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE EVNT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVITYP
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVINAC    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NATURE DU CREDIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVINAC
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVILNC   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIB NATUR CREDIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVILNC
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVINAT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NATURE DU PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVINAT
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVILNA   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIB NATUR PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVILNA
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICLI    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PAYEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICLI
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICET    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETAT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICET
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVILET   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBELLE CODE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVILET
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIRA1   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RAIS1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIRA1
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIRA2   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RAIS2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIRA2
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAD1   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ADRESSE 1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAD1
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAD2   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ADRESSE 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAD2
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAD3   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ADRESSE 3"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAD3
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICOP    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE POSTAL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICOP
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIVIL   25A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "VILLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIVIL
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIPAY   25A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PAYS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIPAY
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIMOD    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MODE REGLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIMOD
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVILMO   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIB MODE REGLT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVILMO
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIPLA    3S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIPLA
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICOM   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPTE OU RIB"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICOM
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIDEV
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVILDE   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBELLE DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVILDE
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIPER    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PERIODICITE M S T A"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIPER
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIREF   50A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REFERENCE EXTERNE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIREF
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIMDO 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT DU DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIMDO
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIDED    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIDED
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVILDD   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBELLE DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVILDD
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIMPR 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT DU PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIMPR
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIDEP    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIDEP
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVILDP   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBELLE DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVILDP
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIMON 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT:MAD,AMORT,ASS,COM"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIMON
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIMIN 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERETS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIMIN
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVITVA 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DE LA TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVITVA
frmElpDisplay.fgData.Row = 40
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVITAU 12.9S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX DU PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVITAU
frmElpDisplay.fgData.Row = 41
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICOT    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE TAUX"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICOT
frmElpDisplay.fgData.Row = 42
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIMAR 12.9S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MARGE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIMAR
frmElpDisplay.fgData.Row = 43
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVITTV    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TX DE LA TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVITTV
frmElpDisplay.fgData.Row = 44
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIVTT 12.9S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "VALEUR TX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIVTT
frmElpDisplay.fgData.Row = 45
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIRGL    7S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE REGLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIRGL
frmElpDisplay.fgData.Row = 46
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIECH    7S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ECHEANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIECH
frmElpDisplay.fgData.Row = 47
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIDEB    7S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DEBUT CALCUL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIDEB
frmElpDisplay.fgData.Row = 48
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIFIN    7S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE FIN CALCUL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIFIN
frmElpDisplay.fgData.Row = 49
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICC1    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE   COMMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICC1
frmElpDisplay.fgData.Row = 50
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICL1   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBEL. COMMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICL1
frmElpDisplay.fgData.Row = 51
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICM1 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT COMMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICM1
frmElpDisplay.fgData.Row = 52
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICS1    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "A RECEVOIR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICS1
frmElpDisplay.fgData.Row = 53
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICB1    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BASE COMMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICB1
frmElpDisplay.fgData.Row = 54
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIC11  7.5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX COMMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIC11
frmElpDisplay.fgData.Row = 55
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIC21    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE TAUX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIC21
frmElpDisplay.fgData.Row = 56
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIC31 12.9S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "VALEUR TX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIC31
frmElpDisplay.fgData.Row = 57
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIC41 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIC41
frmElpDisplay.fgData.Row = 58
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICA1 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ASSIETTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICA1
frmElpDisplay.fgData.Row = 59
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICC2    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE   COMMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICC2
frmElpDisplay.fgData.Row = 60
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICL2   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBEL. COMMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICL2
frmElpDisplay.fgData.Row = 61
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICM2 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT COMMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICM2
frmElpDisplay.fgData.Row = 62
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICS2    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "A RECEVOIR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICS2
frmElpDisplay.fgData.Row = 63
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICB2    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BASE COMMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICB2
frmElpDisplay.fgData.Row = 64
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIC12  7.5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX COMMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIC12
frmElpDisplay.fgData.Row = 65
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIC22    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE TAUX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIC22
frmElpDisplay.fgData.Row = 66
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIC32 12.9S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "VALEUR TX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIC32
frmElpDisplay.fgData.Row = 67
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIC42 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIC42
frmElpDisplay.fgData.Row = 68
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICA2 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ASSIETTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICA2
frmElpDisplay.fgData.Row = 69
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAC1    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE   ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAC1
frmElpDisplay.fgData.Row = 70
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAL1   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBEL. ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAL1
frmElpDisplay.fgData.Row = 71
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAM1 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAM1
frmElpDisplay.fgData.Row = 72
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAB1    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BASE ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAB1
frmElpDisplay.fgData.Row = 73
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA11  7.5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA11
frmElpDisplay.fgData.Row = 74
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA21    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE TAUX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA21
frmElpDisplay.fgData.Row = 75
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA31 12.9S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "VALEUR TX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA31
frmElpDisplay.fgData.Row = 76
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA41 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA41
frmElpDisplay.fgData.Row = 77
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAA1 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ASSIETTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAA1
frmElpDisplay.fgData.Row = 78
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAC2    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE   ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAC2
frmElpDisplay.fgData.Row = 79
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAL2   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBEL. ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAL2
frmElpDisplay.fgData.Row = 80
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAM2 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAM2
frmElpDisplay.fgData.Row = 81
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAB2    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BASE ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAB2
frmElpDisplay.fgData.Row = 82
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA12  7.5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA12
frmElpDisplay.fgData.Row = 83
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA22    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE TAUX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA22
frmElpDisplay.fgData.Row = 84
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA32 12.9S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "VALEUR TX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA32
frmElpDisplay.fgData.Row = 85
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA42 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA42
frmElpDisplay.fgData.Row = 86
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAA2 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ASSIETTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAA2
frmElpDisplay.fgData.Row = 87
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAC3    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE   ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAC3
frmElpDisplay.fgData.Row = 88
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAL3   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBEL. ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAL3
frmElpDisplay.fgData.Row = 89
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAM3 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAM3
frmElpDisplay.fgData.Row = 90
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAB3    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BASE ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAB3
frmElpDisplay.fgData.Row = 91
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA13  7.5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA13
frmElpDisplay.fgData.Row = 92
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA23    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE TAUX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA23
frmElpDisplay.fgData.Row = 93
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA33 12.9S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "VALEUR TX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA33
frmElpDisplay.fgData.Row = 94
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA43 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA43
frmElpDisplay.fgData.Row = 95
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAA3 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ASSIETTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAA3
frmElpDisplay.fgData.Row = 96
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAC4    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE   ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAC4
frmElpDisplay.fgData.Row = 97
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAL4   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBEL. ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAL4
frmElpDisplay.fgData.Row = 98
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAM4 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAM4
frmElpDisplay.fgData.Row = 99
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAB4    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BASE ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAB4
frmElpDisplay.fgData.Row = 100
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA14  7.5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA14
frmElpDisplay.fgData.Row = 101
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA24    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE TAUX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA24
frmElpDisplay.fgData.Row = 102
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA34 12.9S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "VALEUR TX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA34
frmElpDisplay.fgData.Row = 103
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA44 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA44
frmElpDisplay.fgData.Row = 104
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAA4 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ASSIETTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAA4
frmElpDisplay.fgData.Row = 105
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAC5    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE   ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAC5
frmElpDisplay.fgData.Row = 106
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAL5   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBEL. ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAL5
frmElpDisplay.fgData.Row = 107
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAM5 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAM5
frmElpDisplay.fgData.Row = 108
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAB5    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BASE ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAB5
frmElpDisplay.fgData.Row = 109
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA15  7.5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX ASSURANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA15
frmElpDisplay.fgData.Row = 110
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA25    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE TAUX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA25
frmElpDisplay.fgData.Row = 111
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA35 12.9S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "VALEUR TX TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA35
frmElpDisplay.fgData.Row = 112
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIA45 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIA45
frmElpDisplay.fgData.Row = 113
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAA5 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ASSIETTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAA5
frmElpDisplay.fgData.Row = 114
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVINET 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT REGLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVINET
frmElpDisplay.fgData.Row = 115
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIDAT    7S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE AVIS = 0"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIDAT
frmElpDisplay.fgData.Row = 116
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICRD 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CRD AVT ECHEANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICRD
frmElpDisplay.fgData.Row = 117
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVINUM    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO ECHEANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVINUM
frmElpDisplay.fgData.Row = 118
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVITYC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE DE CREDIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVITYC
frmElpDisplay.fgData.Row = 119
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIMDR 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT REGLE EN DER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIMDR
frmElpDisplay.fgData.Row = 120
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIPRC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PRECOMPTE O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIPRC
frmElpDisplay.fgData.Row = 121
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICOU15.10S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COURS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICOU
frmElpDisplay.fgData.Row = 122
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVITEL   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° TEL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVITEL
frmElpDisplay.fgData.Row = 123
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIFAX   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° FAX"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIFAX
frmElpDisplay.fgData.Row = 124
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVICRP 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CRD APR ECHEANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVICRP
frmElpDisplay.fgData.Row = 125
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIAUT   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE AUTO"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIAUT
frmElpDisplay.fgData.Row = 126
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVINPL    3S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVINPL
frmElpDisplay.fgData.Row = 127
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIPAL    3S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PALIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIPAL
frmElpDisplay.fgData.Row = 128
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVINEC    3S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO ECHEANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVINEC
frmElpDisplay.fgData.Row = 129
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIITC 15.2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INT REPORTES PAYES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIITC
frmElpDisplay.fgData.Row = 130
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVISE1    3S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SEQUENCE 1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVISE1
frmElpDisplay.fgData.Row = 131
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVISE2    3S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SEQUENCE 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVISE2
frmElpDisplay.fgData.Row = 132
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREAVIDTC    7S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE CREATION AVIS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREAVI0.CREAVIDTC
frmElpDisplay.Show vbModal
End Sub


