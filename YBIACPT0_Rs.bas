Attribute VB_Name = "rsYBIACPT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const constYBIACPT0 = "YBIACPT0"

Type typeYBIACPT0
   
    COMPTEETA       As Integer                        ' ETABLISSEMENT
    COMPTEPLA       As Long                           ' NUMERO PLAN
    COMPTECOM       As String * 20                    ' NUMERO COMPTE
    COMPTEOBL       As String * 10                    ' COMPTE OBLIGATOIRE
    COMPTEINT       As String * 32                    ' INTITULE
    COMPTEAGE       As Integer                        ' AGENCE
    COMPTEDEV       As String * 3                     ' TABLES BASE 013
    COMPTEOUV       As Long                           ' DATE OUVERTURE
    COMPTECLO       As Long                           ' DATE CLOTURE
    COMPTELOR       As String * 1                     ' Lori/Nostri/AUTRE
    COMPTESUC       As String * 1                     ' O/N
    COMPTECLA       As Long                           ' CLASSE SECURITE
    COMPTEFON       As String * 1                     ' TABLES BASE 015
    COMPTEBLO       As Long                           ' DATE LIMITE BLOCAGE
    COMPTEMOT       As String * 32                    ' MOTIF BLOCAGE
    COMPTESEN       As String * 1                     ' CODE SENS SOLDE D/C
    COMPTEMOD       As Long                           ' DATE MODIFICATION
    
    
    CLIENAETB       As Integer                        ' CODE ETABLISSEMENT
    CLIENACLI       As String * 7                     ' NUMERO CLIENT
    CLIENAAGE       As Integer                        ' CODE AGENCE
    CLIENAETA       As String * 4                     ' CODE ETAT
    CLIENARA1       As String * 32                    ' NOM OU DESIGNATION
    CLIENARA2       As String * 32                    ' PRENOM/DESIGNATION
    CLIENASIG       As String * 12                    ' SIGLE USUEL
    CLIENASRN       As String * 9                     ' NUMERO SIREN
    CLIENASRT       As Long                           ' NUMERO SIRET
    CLIENADNA       As Long                           ' DATE DE NAISSANCE
    CLIENAREG       As String * 6                     ' SECT ACTIVITE REGLEM
    CLIENANAT       As String * 3                     ' CDE PAYS NATIONALITE
    CLIENARSD       As String * 3                     ' CDE PAYS DE RESIDENC
    CLIENARES       As String * 3                     ' RESPONS/EXPLOITATION
    CLIENAECO       As String * 3                     ' QUALITE/AG ECONOMIQU
    CLIENAACT       As String * 1                     ' COTE ACTIVITE
    CLIENAPAI       As String * 1                     ' COTE PAIEMENT
    CLIENACRD       As String * 1                     ' COTE CREDIT
    CLIENAADM       As String * 1                     ' COTE ADMISSION
    CLIENAATR       As Long                           ' DAT ATRIB/COTAT BDF
    CLIENABIL       As Long                           ' AN DERN BIL COMM BDF
    CLIENACAT       As String * 3                     ' CATEGORIE CLIENT
    CLIENACOT       As String * 3                     ' COTATION INTERNE
    CLIENACHQ       As String * 1                     ' INTERDICTION CHEQUIE
    CLIENADAT       As Long                           ' INTERDIT CHEQUIER
    CLIENASAC       As String * 6                     ' SECTEUR D ACTIVITE
    CLIENAGEO       As String * 3                     ' SECTEUR GEOGRAPHIQUE
    CLIENAENT       As String * 3                     ' ENTREPRISE LIEE
    CLIENAMES       As String * 1                     ' LANGUE MESSAGERIE
    CLIENAPAY       As Long                           ' DATE ENTREE AU PAYS
    CLIENAFIL       As String * 32                    ' NOM DE JEUNE FILLE
    CLIENABIM       As Long                           ' BILAN DE MOIS
    CLIENADOU       As String * 1                     ' CLIENT DOUTEUX O/N
    CLIENALI1       As String * 3                     ' ZONE LIBRE DE 3 CAR.
    CLIENALI2       As String * 2                     ' ZONE LIBRE DE 2 CAR.
    CLIENAEXT       As String * 32                    ' EXTENTION DU NOM
    CLIENACOL       As String * 1                     ' 0=CLI/COLL=1/AUTRE=2
    CLIENATIE       As String * 7                     ' TIERS DE REFERENCE
    CLIENASEL       As String * 3                     ' CODE SELECTION
    CLIENAPCS       As String * 4                     ' CODE PCS
    CLIENACRE       As Long                           ' DATE CREATION
    
    PLANCOPRO       As String * 3                     ' TABLES BASE 014
    
    SOLDEDMO        As Long                           ' DATE DERNIER MVT
    SOLDECEN        As Currency                       ' SOLDE ENCOURS
    COMREFREF       As String * 15                    ' EX référence

    TITULACLI       As String * 7                     ' NUMERO CLIENT
    TITULAPRI       As String * 1                     ' 0:PRINCIPAL, 1:AUTRE
    TITULATPR       As String * 1                     ' 0:PRINCIPAL, 1:AUTRE
    
    COMREFCOR       As String * 2                      ' service gestionnaire

End Type
    
'---------------------------------------------------------
Public Function rsYBIACPT0_GetBuffer(rsAdo As ADODB.Recordset, rsYBIACPT0 As typeYBIACPT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYBIACPT0_GetBuffer = Null
    rsYBIACPT0.COMPTEETA = rsAdo("COMPTEETA")
    rsYBIACPT0.COMPTEPLA = rsAdo("COMPTEPLA")
    rsYBIACPT0.COMPTECOM = rsAdo("COMPTECOM")
    rsYBIACPT0.COMPTEOBL = rsAdo("COMPTEOBL")
    rsYBIACPT0.COMPTEINT = rsAdo("COMPTEINT")
    rsYBIACPT0.COMPTEAGE = rsAdo("COMPTEAGE")
    rsYBIACPT0.COMPTEDEV = rsAdo("COMPTEDEV")
    rsYBIACPT0.COMPTEOUV = rsAdo("COMPTEOUV")
    rsYBIACPT0.COMPTECLO = rsAdo("COMPTECLO")
    rsYBIACPT0.COMPTELOR = rsAdo("COMPTELOR")
    rsYBIACPT0.COMPTESUC = rsAdo("COMPTESUC")
    rsYBIACPT0.COMPTECLA = rsAdo("COMPTECLA")
    rsYBIACPT0.COMPTEFON = rsAdo("COMPTEFON")
    rsYBIACPT0.COMPTEBLO = rsAdo("COMPTEBLO")
    rsYBIACPT0.COMPTEMOT = rsAdo("COMPTEMOT")
    rsYBIACPT0.COMPTESEN = rsAdo("COMPTESEN")
    rsYBIACPT0.COMPTEMOD = rsAdo("COMPTEMOD")
    
    rsYBIACPT0.CLIENAETB = Val(rsAdo("CLIENAETB"))
    rsYBIACPT0.CLIENACLI = rsAdo("CLIENACLI")
    rsYBIACPT0.CLIENAAGE = Val(rsAdo("CLIENAAGE"))
    rsYBIACPT0.CLIENAETA = rsAdo("CLIENAETA")
    rsYBIACPT0.CLIENARA1 = rsAdo("CLIENARA1")
    rsYBIACPT0.CLIENARA2 = rsAdo("CLIENARA2")
    rsYBIACPT0.CLIENASIG = rsAdo("CLIENASIG")
    rsYBIACPT0.CLIENASRN = rsAdo("CLIENASRN")
    rsYBIACPT0.CLIENASRT = Val(rsAdo("CLIENASRT"))
    rsYBIACPT0.CLIENADNA = Val(rsAdo("CLIENADNA"))
    rsYBIACPT0.CLIENAREG = rsAdo("CLIENAREG")
    rsYBIACPT0.CLIENANAT = rsAdo("CLIENANAT")
    rsYBIACPT0.CLIENARSD = rsAdo("CLIENARSD")
    rsYBIACPT0.CLIENARES = rsAdo("CLIENARES")
    rsYBIACPT0.CLIENAECO = rsAdo("CLIENAECO")
    rsYBIACPT0.CLIENAACT = rsAdo("CLIENAACT")
    rsYBIACPT0.CLIENAPAI = rsAdo("CLIENAPAI")
    rsYBIACPT0.CLIENACRD = rsAdo("CLIENACRD")
    rsYBIACPT0.CLIENAADM = rsAdo("CLIENAADM")
    rsYBIACPT0.CLIENAATR = Val(rsAdo("CLIENAATR"))
    rsYBIACPT0.CLIENABIL = Val(rsAdo("CLIENABIL"))
    rsYBIACPT0.CLIENACAT = rsAdo("CLIENACAT")
    rsYBIACPT0.CLIENACOT = rsAdo("CLIENACOT")
    rsYBIACPT0.CLIENACHQ = rsAdo("CLIENACHQ")
    rsYBIACPT0.CLIENADAT = Val(rsAdo("CLIENADAT"))
    rsYBIACPT0.CLIENASAC = rsAdo("CLIENASAC")
    rsYBIACPT0.CLIENAGEO = rsAdo("CLIENAGEO")
    rsYBIACPT0.CLIENAENT = rsAdo("CLIENAENT")
    rsYBIACPT0.CLIENAMES = rsAdo("CLIENAMES")
    rsYBIACPT0.CLIENAPAY = Val(rsAdo("CLIENAPAY"))
    rsYBIACPT0.CLIENAFIL = rsAdo("CLIENAFIL")
    rsYBIACPT0.CLIENABIM = Val(rsAdo("CLIENABIM"))
    rsYBIACPT0.CLIENADOU = rsAdo("CLIENADOU")
    rsYBIACPT0.CLIENALI1 = rsAdo("CLIENALI1")
    rsYBIACPT0.CLIENALI2 = rsAdo("CLIENALI2")
    rsYBIACPT0.CLIENAEXT = rsAdo("CLIENAEXT")
    rsYBIACPT0.CLIENACOL = rsAdo("CLIENACOL")
    rsYBIACPT0.CLIENATIE = rsAdo("CLIENATIE")
    rsYBIACPT0.CLIENASEL = rsAdo("CLIENASEL")
    rsYBIACPT0.CLIENAPCS = rsAdo("CLIENAPCS")
    rsYBIACPT0.CLIENACRE = Val(rsAdo("CLIENACRE"))
    
    rsYBIACPT0.PLANCOPRO = rsAdo("PLANCOPRO")
    rsYBIACPT0.SOLDEDMO = Val(rsAdo("SOLDEDMO"))
    rsYBIACPT0.SOLDECEN = rsAdo("SOLDECEN") / 1000
    If IsNull(rsAdo("COMREFREF")) Then
        rsYBIACPT0.COMREFREF = ""
    Else
        rsYBIACPT0.COMREFREF = rsAdo("COMREFREF")
    End If
    
    rsYBIACPT0.TITULACLI = rsAdo("TITULACLI")
    rsYBIACPT0.TITULAPRI = rsAdo("TITULAPRI")
    rsYBIACPT0.TITULATPR = rsAdo("TITULATPR")

Exit Function

Error_Handler:

rsYBIACPT0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsYBIACPT0_Init(rsYBIACPT0 As typeYBIACPT0)
'---------------------------------------------------------
rsYBIACPT0.COMPTEETA = 1

rsYBIACPT0.COMPTEPLA = 0 '       As Long                           ' NUMERO PLAN
rsYBIACPT0.COMPTECOM = "" '       As String * 20                    ' NUMERO COMPTE
rsYBIACPT0.COMPTEOBL = "" '       As String * 10                    ' COMPTE OBLIGATOIRE
rsYBIACPT0.COMPTEINT = "" '       As String * 32                    ' INTITULE
rsYBIACPT0.COMPTEAGE = 0 '       As Integer                        ' AGENCE
rsYBIACPT0.COMPTEDEV = "" '       As String * 3                     ' TABLES BASE 013
rsYBIACPT0.COMPTEOUV = 0 '       As Long                           ' DATE OUVERTURE
rsYBIACPT0.COMPTECLO = 0 '       As Long                           ' DATE CLOTURE
rsYBIACPT0.COMPTELOR = "" '       As String * 1                     ' Lori/Nostri/AUTRE
rsYBIACPT0.COMPTESUC = "" '       As String * 1                     ' O/N
rsYBIACPT0.COMPTECLA = 0 '       As Long                           ' CLASSE SECURITE
rsYBIACPT0.COMPTEFON = "" '       As String * 1                     ' TABLES BASE 015
rsYBIACPT0.COMPTEBLO = 0 '       As Long                           ' DATE LIMITE BLOCAGE
rsYBIACPT0.COMPTEMOT = "" '       As String * 32                    ' MOTIF BLOCAGE
rsYBIACPT0.COMPTESEN = "" '       As String * 1                     ' CODE SENS SOLDE D/C
rsYBIACPT0.COMPTEMOD = 0 '       As Long                           ' DATE MODIFICATION
    
    
rsYBIACPT0.CLIENAETB = 0 '       As Integer                        ' CODE ETABLISSEMENT
rsYBIACPT0.CLIENACLI = "" '       As String * 7                     ' NUMERO CLIENT
rsYBIACPT0.CLIENAAGE = 0 '       As Integer                        ' CODE AGENCE
rsYBIACPT0.CLIENAETA = "" '       As String * 4                     ' CODE ETAT
rsYBIACPT0.CLIENARA1 = "" '       As String * 32                    ' NOM OU DESIGNATION
rsYBIACPT0.CLIENARA2 = "" '       As String * 32                    ' PRENOM/DESIGNATION
rsYBIACPT0.CLIENASIG = "" '       As String * 12                    ' SIGLE USUEL
rsYBIACPT0.CLIENASRN = "" '       As String * 9                     ' NUMERO SIREN
rsYBIACPT0.CLIENASRT = 0 '       As Long                           ' NUMERO SIRET
rsYBIACPT0.CLIENADNA = 0 '       As Long                           ' DATE DE NAISSANCE
rsYBIACPT0.CLIENAREG = "" '       As String * 6                     ' SECT ACTIVITE REGLEM
rsYBIACPT0.CLIENANAT = "" '       As String * 3                     ' CDE PAYS NATIONALITE
rsYBIACPT0.CLIENARSD = "" '       As String * 3                     ' CDE PAYS DE RESIDENC
rsYBIACPT0.CLIENARES = "" '       As String * 3                     ' RESPONS/EXPLOITATION
rsYBIACPT0.CLIENAECO = "" '       As String * 3                     ' QUALITE/AG ECONOMIQU
rsYBIACPT0.CLIENAACT = "" '       As String * 1                     ' COTE ACTIVITE
rsYBIACPT0.CLIENAPAI = "" '       As String * 1                     ' COTE PAIEMENT
rsYBIACPT0.CLIENACRD = "" '       As String * 1                     ' COTE CREDIT
rsYBIACPT0.CLIENAADM = "" '       As String * 1                     ' COTE ADMISSION
rsYBIACPT0.CLIENAATR = 0 '       As Long                           ' DAT ATRIB/COTAT BDF
rsYBIACPT0.CLIENABIL = 0 '       As Long                           ' AN DERN BIL COMM BDF
rsYBIACPT0.CLIENACAT = "" '       As String * 3                     ' CATEGORIE CLIENT
rsYBIACPT0.CLIENACOT = "" '       As String * 3                     ' COTATION INTERNE
rsYBIACPT0.CLIENACHQ = "" '       As String * 1                     ' INTERDICTION CHEQUIE
rsYBIACPT0.CLIENADAT = 0 '       As Long                           ' INTERDIT CHEQUIER
rsYBIACPT0.CLIENASAC = "" '       As String * 6                     ' SECTEUR D ACTIVITE
rsYBIACPT0.CLIENAGEO = "" '       As String * 3                     ' SECTEUR GEOGRAPHIQUE
rsYBIACPT0.CLIENAENT = "" '       As String * 3                     ' ENTREPRISE LIEE
rsYBIACPT0.CLIENAMES = "" '       As String * 1                     ' LANGUE MESSAGERIE
rsYBIACPT0.CLIENAPAY = 0 '       As Long                           ' DATE ENTREE AU PAYS
rsYBIACPT0.CLIENAFIL = "" '       As String * 32                    ' NOM DE JEUNE FILLE
rsYBIACPT0.CLIENABIM = 0 '       As Long                           ' BILAN DE MOIS
rsYBIACPT0.CLIENADOU = "" '       As String * 1                     ' CLIENT DOUTEUX O/N
rsYBIACPT0.CLIENALI1 = "" '       As String * 3                     ' ZONE LIBRE DE 3 CAR.
rsYBIACPT0.CLIENALI2 = "" '       As String * 2                     ' ZONE LIBRE DE 2 CAR.
rsYBIACPT0.CLIENAEXT = "" '       As String * 32                    ' EXTENTION DU NOM
rsYBIACPT0.CLIENACOL = "" '       As String * 1                     ' 0=CLI/COLL=1/AUTRE=2
rsYBIACPT0.CLIENATIE = "" '       As String * 7                     ' TIERS DE REFERENCE
rsYBIACPT0.CLIENASEL = "" '       As String * 3                     ' CODE SELECTION
rsYBIACPT0.CLIENAPCS = "" '       As String * 4                     ' CODE PCS
rsYBIACPT0.CLIENACRE = 0 '       As Long                           ' DATE CREATION
    
rsYBIACPT0.PLANCOPRO = "" '       As String * 3                     ' TABLES BASE 014
    
rsYBIACPT0.SOLDEDMO = 0 '       As Long                           ' DATE DERNIER MVT
rsYBIACPT0.SOLDECEN = 0 '       As Currency                       ' SOLDE ENCOURS
rsYBIACPT0.COMREFREF = "" '       As String * 15                    ' EX référence

rsYBIACPT0.TITULACLI = "" '       As String * 7                     ' NUMERO CLIENT
rsYBIACPT0.TITULAPRI = "" '       As String * 1                     ' 0:PRINCIPAL, 1:AUTRE
rsYBIACPT0.TITULATPR = "" '       As String * 1                     ' 0:PRINCIPAL, 1:AUTRE

End Sub







