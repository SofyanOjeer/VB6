Attribute VB_Name = "rsZCREAVI0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCREAVI0
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
Public Sub rsZCREAVI0_Init(rsYCREAVI0 As typeZCREAVI0)
rsYCREAVI0.CREAVIETA = 0
rsYCREAVI0.CREAVIAGE = 0
rsYCREAVI0.CREAVISER = ""
rsYCREAVI0.CREAVISSE = ""
rsYCREAVI0.CREAVIDOS = 0
rsYCREAVI0.CREAVIPRE = 0
rsYCREAVI0.CREAVITYP = ""
rsYCREAVI0.CREAVINAC = ""
rsYCREAVI0.CREAVILNC = ""
rsYCREAVI0.CREAVINAT = ""
rsYCREAVI0.CREAVILNA = ""
rsYCREAVI0.CREAVICLI = ""
rsYCREAVI0.CREAVICET = ""
rsYCREAVI0.CREAVILET = ""
rsYCREAVI0.CREAVIRA1 = ""
rsYCREAVI0.CREAVIRA2 = ""
rsYCREAVI0.CREAVIAD1 = ""
rsYCREAVI0.CREAVIAD2 = ""
rsYCREAVI0.CREAVIAD3 = ""
rsYCREAVI0.CREAVICOP = ""
rsYCREAVI0.CREAVIVIL = ""
rsYCREAVI0.CREAVIPAY = ""
rsYCREAVI0.CREAVIMOD = ""
rsYCREAVI0.CREAVILMO = ""
rsYCREAVI0.CREAVIPLA = 0
rsYCREAVI0.CREAVICOM = ""
rsYCREAVI0.CREAVIDEV = ""
rsYCREAVI0.CREAVILDE = ""
rsYCREAVI0.CREAVIPER = ""
rsYCREAVI0.CREAVIREF = ""
rsYCREAVI0.CREAVIMDO = 0
rsYCREAVI0.CREAVIDED = ""
rsYCREAVI0.CREAVILDD = ""
rsYCREAVI0.CREAVIMPR = 0
rsYCREAVI0.CREAVIDEP = ""
rsYCREAVI0.CREAVILDP = ""
rsYCREAVI0.CREAVIMON = 0
rsYCREAVI0.CREAVIMIN = 0
rsYCREAVI0.CREAVITVA = 0
rsYCREAVI0.CREAVITAU = 0
rsYCREAVI0.CREAVICOT = ""
rsYCREAVI0.CREAVIMAR = 0
rsYCREAVI0.CREAVITTV = ""
rsYCREAVI0.CREAVIVTT = 0
rsYCREAVI0.CREAVIRGL = 0
rsYCREAVI0.CREAVIECH = 0
rsYCREAVI0.CREAVIDEB = 0
rsYCREAVI0.CREAVIFIN = 0
rsYCREAVI0.CREAVICC1 = ""
rsYCREAVI0.CREAVICL1 = ""
rsYCREAVI0.CREAVICM1 = 0
rsYCREAVI0.CREAVICS1 = ""
rsYCREAVI0.CREAVICB1 = ""
rsYCREAVI0.CREAVIC11 = 0
rsYCREAVI0.CREAVIC21 = ""
rsYCREAVI0.CREAVIC31 = 0
rsYCREAVI0.CREAVIC41 = 0
rsYCREAVI0.CREAVICA1 = 0
rsYCREAVI0.CREAVICC2 = ""
rsYCREAVI0.CREAVICL2 = ""
rsYCREAVI0.CREAVICM2 = 0
rsYCREAVI0.CREAVICS2 = ""
rsYCREAVI0.CREAVICB2 = ""
rsYCREAVI0.CREAVIC12 = 0
rsYCREAVI0.CREAVIC22 = ""
rsYCREAVI0.CREAVIC32 = 0
rsYCREAVI0.CREAVIC42 = 0
rsYCREAVI0.CREAVICA2 = 0
rsYCREAVI0.CREAVIAC1 = ""
rsYCREAVI0.CREAVIAL1 = ""
rsYCREAVI0.CREAVIAM1 = 0
rsYCREAVI0.CREAVIAB1 = ""
rsYCREAVI0.CREAVIA11 = 0
rsYCREAVI0.CREAVIA21 = ""
rsYCREAVI0.CREAVIA31 = 0
rsYCREAVI0.CREAVIA41 = 0
rsYCREAVI0.CREAVIAA1 = 0
rsYCREAVI0.CREAVIAC2 = ""
rsYCREAVI0.CREAVIAL2 = ""
rsYCREAVI0.CREAVIAM2 = 0
rsYCREAVI0.CREAVIAB2 = ""
rsYCREAVI0.CREAVIA12 = 0
rsYCREAVI0.CREAVIA22 = ""
rsYCREAVI0.CREAVIA32 = 0
rsYCREAVI0.CREAVIA42 = 0
rsYCREAVI0.CREAVIAA2 = 0
rsYCREAVI0.CREAVIAC3 = ""
rsYCREAVI0.CREAVIAL3 = ""
rsYCREAVI0.CREAVIAM3 = 0
rsYCREAVI0.CREAVIAB3 = ""
rsYCREAVI0.CREAVIA13 = 0
rsYCREAVI0.CREAVIA23 = ""
rsYCREAVI0.CREAVIA33 = 0
rsYCREAVI0.CREAVIA43 = 0
rsYCREAVI0.CREAVIAA3 = 0
rsYCREAVI0.CREAVIAC4 = ""
rsYCREAVI0.CREAVIAL4 = ""
rsYCREAVI0.CREAVIAM4 = 0
rsYCREAVI0.CREAVIAB4 = ""
rsYCREAVI0.CREAVIA14 = 0
rsYCREAVI0.CREAVIA24 = ""
rsYCREAVI0.CREAVIA34 = 0
rsYCREAVI0.CREAVIA44 = 0
rsYCREAVI0.CREAVIAA4 = 0
rsYCREAVI0.CREAVIAC5 = ""
rsYCREAVI0.CREAVIAL5 = ""
rsYCREAVI0.CREAVIAM5 = 0
rsYCREAVI0.CREAVIAB5 = ""
rsYCREAVI0.CREAVIA15 = 0
rsYCREAVI0.CREAVIA25 = ""
rsYCREAVI0.CREAVIA35 = 0
rsYCREAVI0.CREAVIA45 = 0
rsYCREAVI0.CREAVIAA5 = 0
rsYCREAVI0.CREAVINET = 0
rsYCREAVI0.CREAVIDAT = 0
rsYCREAVI0.CREAVICRD = 0
rsYCREAVI0.CREAVINUM = 0
rsYCREAVI0.CREAVITYC = ""
rsYCREAVI0.CREAVIMDR = 0
rsYCREAVI0.CREAVIPRC = ""
rsYCREAVI0.CREAVICOU = 0
rsYCREAVI0.CREAVITEL = ""
rsYCREAVI0.CREAVIFAX = ""
rsYCREAVI0.CREAVICRP = 0
rsYCREAVI0.CREAVIAUT = ""
rsYCREAVI0.CREAVINPL = 0
rsYCREAVI0.CREAVIPAL = 0
rsYCREAVI0.CREAVINEC = 0
rsYCREAVI0.CREAVIITC = 0
rsYCREAVI0.CREAVISE1 = 0
rsYCREAVI0.CREAVISE2 = 0
rsYCREAVI0.CREAVIDTC = 0
End Sub
Public Function rsZCREAVI0_GetBuffer(rsAdo As ADODB.Recordset, rsZCREAVI0 As typeZCREAVI0)
On Error GoTo Error_Handler
rsZCREAVI0_GetBuffer = Null
rsZCREAVI0.CREAVIETA = rsAdo("CREAVIETA")
rsZCREAVI0.CREAVIAGE = rsAdo("CREAVIAGE")
rsZCREAVI0.CREAVISER = rsAdo("CREAVISER")
rsZCREAVI0.CREAVISSE = rsAdo("CREAVISSE")
rsZCREAVI0.CREAVIDOS = rsAdo("CREAVIDOS")
rsZCREAVI0.CREAVIPRE = rsAdo("CREAVIPRE")
rsZCREAVI0.CREAVITYP = rsAdo("CREAVITYP")
rsZCREAVI0.CREAVINAC = rsAdo("CREAVINAC")
rsZCREAVI0.CREAVILNC = rsAdo("CREAVILNC")
rsZCREAVI0.CREAVINAT = rsAdo("CREAVINAT")
rsZCREAVI0.CREAVILNA = rsAdo("CREAVILNA")
rsZCREAVI0.CREAVICLI = rsAdo("CREAVICLI")
rsZCREAVI0.CREAVICET = rsAdo("CREAVICET")
rsZCREAVI0.CREAVILET = rsAdo("CREAVILET")
rsZCREAVI0.CREAVIRA1 = rsAdo("CREAVIRA1")
rsZCREAVI0.CREAVIRA2 = rsAdo("CREAVIRA2")
rsZCREAVI0.CREAVIAD1 = rsAdo("CREAVIAD1")
rsZCREAVI0.CREAVIAD2 = rsAdo("CREAVIAD2")
rsZCREAVI0.CREAVIAD3 = rsAdo("CREAVIAD3")
rsZCREAVI0.CREAVICOP = rsAdo("CREAVICOP")
rsZCREAVI0.CREAVIVIL = rsAdo("CREAVIVIL")
rsZCREAVI0.CREAVIPAY = rsAdo("CREAVIPAY")
rsZCREAVI0.CREAVIMOD = rsAdo("CREAVIMOD")
rsZCREAVI0.CREAVILMO = rsAdo("CREAVILMO")
rsZCREAVI0.CREAVIPLA = rsAdo("CREAVIPLA")
rsZCREAVI0.CREAVICOM = rsAdo("CREAVICOM")
rsZCREAVI0.CREAVIDEV = rsAdo("CREAVIDEV")
rsZCREAVI0.CREAVILDE = rsAdo("CREAVILDE")
rsZCREAVI0.CREAVIPER = rsAdo("CREAVIPER")
rsZCREAVI0.CREAVIREF = rsAdo("CREAVIREF")
rsZCREAVI0.CREAVIMDO = rsAdo("CREAVIMDO")
rsZCREAVI0.CREAVIDED = rsAdo("CREAVIDED")
rsZCREAVI0.CREAVILDD = rsAdo("CREAVILDD")
rsZCREAVI0.CREAVIMPR = rsAdo("CREAVIMPR")
rsZCREAVI0.CREAVIDEP = rsAdo("CREAVIDEP")
rsZCREAVI0.CREAVILDP = rsAdo("CREAVILDP")
rsZCREAVI0.CREAVIMON = rsAdo("CREAVIMON")
rsZCREAVI0.CREAVIMIN = rsAdo("CREAVIMIN")
rsZCREAVI0.CREAVITVA = rsAdo("CREAVITVA")
rsZCREAVI0.CREAVITAU = rsAdo("CREAVITAU")
rsZCREAVI0.CREAVICOT = rsAdo("CREAVICOT")
rsZCREAVI0.CREAVIMAR = rsAdo("CREAVIMAR")
rsZCREAVI0.CREAVITTV = rsAdo("CREAVITTV")
rsZCREAVI0.CREAVIVTT = rsAdo("CREAVIVTT")
rsZCREAVI0.CREAVIRGL = rsAdo("CREAVIRGL")
rsZCREAVI0.CREAVIECH = rsAdo("CREAVIECH")
rsZCREAVI0.CREAVIDEB = rsAdo("CREAVIDEB")
rsZCREAVI0.CREAVIFIN = rsAdo("CREAVIFIN")
rsZCREAVI0.CREAVICC1 = rsAdo("CREAVICC1")
rsZCREAVI0.CREAVICL1 = rsAdo("CREAVICL1")
rsZCREAVI0.CREAVICM1 = rsAdo("CREAVICM1")
rsZCREAVI0.CREAVICS1 = rsAdo("CREAVICS1")
rsZCREAVI0.CREAVICB1 = rsAdo("CREAVICB1")
rsZCREAVI0.CREAVIC11 = rsAdo("CREAVIC11")
rsZCREAVI0.CREAVIC21 = rsAdo("CREAVIC21")
rsZCREAVI0.CREAVIC31 = rsAdo("CREAVIC31")
rsZCREAVI0.CREAVIC41 = rsAdo("CREAVIC41")
rsZCREAVI0.CREAVICA1 = rsAdo("CREAVICA1")
rsZCREAVI0.CREAVICC2 = rsAdo("CREAVICC2")
rsZCREAVI0.CREAVICL2 = rsAdo("CREAVICL2")
rsZCREAVI0.CREAVICM2 = rsAdo("CREAVICM2")
rsZCREAVI0.CREAVICS2 = rsAdo("CREAVICS2")
rsZCREAVI0.CREAVICB2 = rsAdo("CREAVICB2")
rsZCREAVI0.CREAVIC12 = rsAdo("CREAVIC12")
rsZCREAVI0.CREAVIC22 = rsAdo("CREAVIC22")
rsZCREAVI0.CREAVIC32 = rsAdo("CREAVIC32")
rsZCREAVI0.CREAVIC42 = rsAdo("CREAVIC42")
rsZCREAVI0.CREAVICA2 = rsAdo("CREAVICA2")
rsZCREAVI0.CREAVIAC1 = rsAdo("CREAVIAC1")
rsZCREAVI0.CREAVIAL1 = rsAdo("CREAVIAL1")
rsZCREAVI0.CREAVIAM1 = rsAdo("CREAVIAM1")
rsZCREAVI0.CREAVIAB1 = rsAdo("CREAVIAB1")
rsZCREAVI0.CREAVIA11 = rsAdo("CREAVIA11")
rsZCREAVI0.CREAVIA21 = rsAdo("CREAVIA21")
rsZCREAVI0.CREAVIA31 = rsAdo("CREAVIA31")
rsZCREAVI0.CREAVIA41 = rsAdo("CREAVIA41")
rsZCREAVI0.CREAVIAA1 = rsAdo("CREAVIAA1")
rsZCREAVI0.CREAVIAC2 = rsAdo("CREAVIAC2")
rsZCREAVI0.CREAVIAL2 = rsAdo("CREAVIAL2")
rsZCREAVI0.CREAVIAM2 = rsAdo("CREAVIAM2")
rsZCREAVI0.CREAVIAB2 = rsAdo("CREAVIAB2")
rsZCREAVI0.CREAVIA12 = rsAdo("CREAVIA12")
rsZCREAVI0.CREAVIA22 = rsAdo("CREAVIA22")
rsZCREAVI0.CREAVIA32 = rsAdo("CREAVIA32")
rsZCREAVI0.CREAVIA42 = rsAdo("CREAVIA42")
rsZCREAVI0.CREAVIAA2 = rsAdo("CREAVIAA2")
rsZCREAVI0.CREAVIAC3 = rsAdo("CREAVIAC3")
rsZCREAVI0.CREAVIAL3 = rsAdo("CREAVIAL3")
rsZCREAVI0.CREAVIAM3 = rsAdo("CREAVIAM3")
rsZCREAVI0.CREAVIAB3 = rsAdo("CREAVIAB3")
rsZCREAVI0.CREAVIA13 = rsAdo("CREAVIA13")
rsZCREAVI0.CREAVIA23 = rsAdo("CREAVIA23")
rsZCREAVI0.CREAVIA33 = rsAdo("CREAVIA33")
rsZCREAVI0.CREAVIA43 = rsAdo("CREAVIA43")
rsZCREAVI0.CREAVIAA3 = rsAdo("CREAVIAA3")
rsZCREAVI0.CREAVIAC4 = rsAdo("CREAVIAC4")
rsZCREAVI0.CREAVIAL4 = rsAdo("CREAVIAL4")
rsZCREAVI0.CREAVIAM4 = rsAdo("CREAVIAM4")
rsZCREAVI0.CREAVIAB4 = rsAdo("CREAVIAB4")
rsZCREAVI0.CREAVIA14 = rsAdo("CREAVIA14")
rsZCREAVI0.CREAVIA24 = rsAdo("CREAVIA24")
rsZCREAVI0.CREAVIA34 = rsAdo("CREAVIA34")
rsZCREAVI0.CREAVIA44 = rsAdo("CREAVIA44")
rsZCREAVI0.CREAVIAA4 = rsAdo("CREAVIAA4")
rsZCREAVI0.CREAVIAC5 = rsAdo("CREAVIAC5")
rsZCREAVI0.CREAVIAL5 = rsAdo("CREAVIAL5")
rsZCREAVI0.CREAVIAM5 = rsAdo("CREAVIAM5")
rsZCREAVI0.CREAVIAB5 = rsAdo("CREAVIAB5")
rsZCREAVI0.CREAVIA15 = rsAdo("CREAVIA15")
rsZCREAVI0.CREAVIA25 = rsAdo("CREAVIA25")
rsZCREAVI0.CREAVIA35 = rsAdo("CREAVIA35")
rsZCREAVI0.CREAVIA45 = rsAdo("CREAVIA45")
rsZCREAVI0.CREAVIAA5 = rsAdo("CREAVIAA5")
rsZCREAVI0.CREAVINET = rsAdo("CREAVINET")
rsZCREAVI0.CREAVIDAT = rsAdo("CREAVIDAT")
rsZCREAVI0.CREAVICRD = rsAdo("CREAVICRD")
rsZCREAVI0.CREAVINUM = rsAdo("CREAVINUM")
rsZCREAVI0.CREAVITYC = rsAdo("CREAVITYC")
rsZCREAVI0.CREAVIMDR = rsAdo("CREAVIMDR")
rsZCREAVI0.CREAVIPRC = rsAdo("CREAVIPRC")
rsZCREAVI0.CREAVICOU = rsAdo("CREAVICOU")
rsZCREAVI0.CREAVITEL = rsAdo("CREAVITEL")
rsZCREAVI0.CREAVIFAX = rsAdo("CREAVIFAX")
rsZCREAVI0.CREAVICRP = rsAdo("CREAVICRP")
rsZCREAVI0.CREAVIAUT = rsAdo("CREAVIAUT")
rsZCREAVI0.CREAVINPL = rsAdo("CREAVINPL")
rsZCREAVI0.CREAVIPAL = rsAdo("CREAVIPAL")
rsZCREAVI0.CREAVINEC = rsAdo("CREAVINEC")
rsZCREAVI0.CREAVIITC = rsAdo("CREAVIITC")
rsZCREAVI0.CREAVISE1 = rsAdo("CREAVISE1")
rsZCREAVI0.CREAVISE2 = rsAdo("CREAVISE2")
rsZCREAVI0.CREAVIDTC = rsAdo("CREAVIDTC")
Exit Function
Error_Handler:
rsZCREAVI0_GetBuffer = Error
End Function

