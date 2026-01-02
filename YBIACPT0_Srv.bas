Attribute VB_Name = "srvYBIACPT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYBIACPT0Len = 524 + 50 ' 34 + 490
Public Const recYBIACPT0_Block = 100
Public Const memoYBIACPT0Len = 490
Public Const constYBIACPT0 = "YBIACPT0"
Public paramYBIACPT0_Import As String
Dim meYbase As typeYBase
Public paramYBIACPT0_Nb As Long

Type typeYBIACPT0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
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

End Type
    
Public Sub srvYBIACPT0_ElpDisplay(recYBIACPT0 As typeYBIACPT0)
frmElpDisplay.fgData.Rows = 65
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTEETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEPLA    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTEPLA
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTECOM   20A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO COMPTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTECOM
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEOBL   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPTE OBLIGATOIRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTEOBL
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEINT   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTITULE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTEINT
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTEAGE
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TABLES BASE 013"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTEDEV
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEOUV    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE OUVERTURE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTEOUV
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTECLO    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE CLOTURE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTECLO
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTELOR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "Lori/Nostri/AUTRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTELOR
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTESUC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTESUC
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTECLA    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLASSE SECURITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTECLA
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEFON    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TABLES BASE 015"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTEFON
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEBLO    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE LIMITE BLOCAGE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTEBLO
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEMOT   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MOTIF BLOCAGE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTEMOT
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTESEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE SENS SOLDE D/C"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTESEN
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMPTEMOD    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE MODIFICATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMPTEMOD

frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAETB
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENACLI    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO CLIENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENACLI
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAAGE
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAETA    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETAT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAETA
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENARA1   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NOM OU DESIGNATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENARA1
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENARA2   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PRENOM/DESIGNATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENARA2
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENASIG   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SIGLE USUEL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENASIG
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENASRN    9A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO SIREN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENASRN
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENASRT    5S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO SIRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENASRT
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENADNA    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DE NAISSANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENADNA
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAREG    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SECT ACTIVITE REGLEM"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAREG
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENANAT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CDE PAYS NATIONALITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENANAT
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENARSD    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CDE PAYS DE RESIDENC"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENARSD
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENARES    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "RESPONS/EXPLOITATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENARES
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAECO    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "QUALITE/AG ECONOMIQU"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAECO
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAACT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTE ACTIVITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAACT
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAPAI    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTE PAIEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAPAI
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENACRD    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTE CREDIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENACRD
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAADM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTE ADMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAADM
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAATR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DAT ATRIB/COTAT BDF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAATR
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENABIL    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AN DERN BIL COMM BDF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENABIL
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENACAT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CATEGORIE CLIENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENACAT
frmElpDisplay.fgData.Row = 40
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENACOT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COTATION INTERNE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENACOT
frmElpDisplay.fgData.Row = 41
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENACHQ    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERDICTION CHEQUIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENACHQ
frmElpDisplay.fgData.Row = 42
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENADAT    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERDIT CHEQUIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENADAT
frmElpDisplay.fgData.Row = 43
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENASAC    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SECTEUR D ACTIVITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENASAC
frmElpDisplay.fgData.Row = 44
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAGEO    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SECTEUR GEOGRAPHIQUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAGEO
frmElpDisplay.fgData.Row = 45
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAENT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ENTREPRISE LIEE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAENT
frmElpDisplay.fgData.Row = 46
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAMES    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LANGUE MESSAGERIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAMES
frmElpDisplay.fgData.Row = 47
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAPAY    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ENTREE AU PAYS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAPAY
frmElpDisplay.fgData.Row = 48
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAFIL   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NOM DE JEUNE FILLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAFIL
frmElpDisplay.fgData.Row = 49
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENABIM    2S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BILAN DE MOIS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENABIM
frmElpDisplay.fgData.Row = 50
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENADOU    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLIENT DOUTEUX O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENADOU
frmElpDisplay.fgData.Row = 51
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENALI1    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ZONE LIBRE DE 3 CAR."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENALI1
frmElpDisplay.fgData.Row = 52
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENALI2    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ZONE LIBRE DE 2 CAR."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENALI2
frmElpDisplay.fgData.Row = 53
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAEXT   32A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EXTENTION DU NOM"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAEXT
frmElpDisplay.fgData.Row = 54
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENACOL    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "0=CLI/COLL=1/AUTRE=2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENACOL
frmElpDisplay.fgData.Row = 55
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENATIE    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TIERS DE REFERENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENATIE
frmElpDisplay.fgData.Row = 56
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENASEL    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE SELECTION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENASEL
frmElpDisplay.fgData.Row = 57
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENAPCS    4A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE PCS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENAPCS
frmElpDisplay.fgData.Row = 58
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CLIENACRE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE CREATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.CLIENACRE

 frmElpDisplay.fgData.Row = 58
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PLANCOPRO    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PLANCOPRO"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.PLANCOPRO
   
frmElpDisplay.fgData.Row = 59
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDEDMO     XX"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDEDMO "
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.SOLDEDMO
frmElpDisplay.fgData.Row = 60
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SOLDECEN    XX"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOLDECEN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.SOLDECEN
frmElpDisplay.fgData.Row = 61
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMREFREF    XX"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMREFREF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.COMREFREF

frmElpDisplay.fgData.Row = 62
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TITULACLI    7XX"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TITULACLI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.TITULACLI
frmElpDisplay.fgData.Row = 63
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TITULAPRI    XX"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TITULAPRI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.TITULAPRI
frmElpDisplay.fgData.Row = 64
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TITULATPR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TITULATPR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYBIACPT0.TITULATPR

frmElpDisplay.Show vbModal
End Sub
    
Public Sub srvYBIACPT0_Export_CSV(lIdFile_Source As Integer, lIdFile_Destination As Integer, loptSelect_CSV_Header As Boolean, lnb As Long)
Dim xIn As String, K As Integer
Dim V
Dim meCV1 As typeCV, meCV2 As typeCV

If loptSelect_CSV_Header Then
    Print #lIdFile_Destination, "?"
    Print #lIdFile_Destination, "?"
    Print #lIdFile_Destination, "?"
End If
Do Until EOF(lIdFile_Source)
      Line Input #lIdFile_Source, xIn
      lnb = lnb + 1
      K = 0
      
      meCV1.DeviseIso = mId$(xIn, 77, 3)

    If meCV1.DeviseIso <> "EUR" Then
        meCV1.DeviseN = 0
        meCV1.Montant = CCur(mId$(xIn, 436 + 12, 19)) / 1000
        meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
        meCV2.OpéAmj = YBIATAB0_DATE_CPT_J
           
        Call CV_Calc("J  ", meCV1, meCV2)
    Else
        meCV2.Montant = CCur(mId$(xIn, 436 + 12, 19)) / 1000
    End If

    Print #lIdFile_Destination, mId$(xIn, 1, 5) & ";" _
    & mId$(xIn, 6, 4) & ";" & mId$(xIn, 10, 20) & ";" & mId$(xIn, 30, 10) & ";" & mId$(xIn, 40, 32) & ";" _
    & mId$(xIn, 72, 5) & ";" & mId$(xIn, 77, 3) & ";" _
    & mId$(xIn, 80, 8) & ";" & mId$(xIn, 88, 8) & ";" & mId$(xIn, 96, 1) & ";" & mId$(xIn, 97, 1) & ";" _
    & mId$(xIn, 98, 3) & ";" & mId$(xIn, 101, 1) & ";" & mId$(xIn, 102, 8) & ";" & mId$(xIn, 110, 32) & ";" _
    & mId$(xIn, 142, 1) & ";" & mId$(xIn, 143, 8) & ";" & mId$(xIn, 150 + 1, 5) & ";" & mId$(xIn, 150 + 6, 7) & ";" _
    & mId$(xIn, 150 + 13, 5) & ";" & mId$(xIn, 150 + 18, 4) & ";" & mId$(xIn, 150 + 22, 32) & ";" & mId$(xIn, 150 + 54, 32) & ";" _
    & mId$(xIn, 150 + 86, 12) & ";" & mId$(xIn, 150 + 98, 9) & ";" & mId$(xIn, 150 + 107, 6) & ";" & mId$(xIn, 150 + 113, 8) & ";" _
    & mId$(xIn, 150 + 121, 6) & ";" & mId$(xIn, 150 + 127, 3) & ";" & mId$(xIn, 150 + 130, 3) & ";" & mId$(xIn, 150 + 133, 3) & ";" _
    & mId$(xIn, 150 + 136, 3) & ";" & mId$(xIn, 150 + 139, 1) & ";" & mId$(xIn, 150 + 140, 1) & ";" & mId$(xIn, 150 + 141, 1) & ";" _
    & mId$(xIn, 150 + 142, 1) & ";" & mId$(xIn, 150 + 143, 8) & ";" & mId$(xIn, 150 + 151, 4) & ";" _
    & mId$(xIn, 150 + 155, 3) & ";" & mId$(xIn, 150 + 158, 3) & ";" _
    & mId$(xIn, 150 + 161, 1) & ";" & mId$(xIn, 150 + 162, 8) & ";" _
    & mId$(xIn, 150 + 170, 6) & ";" & mId$(xIn, 150 + 176, 3) & ";" & mId$(xIn, 150 + 179, 3) & ";" & mId$(xIn, 150 + 182, 1) & ";" _
    & mId$(xIn, 150 + 183, 8) & ";" & mId$(xIn, 150 + 191, 32) & ";" & mId$(xIn, 150 + 223, 3) & ";" & mId$(xIn, 150 + 226, 1) & ";" _
    & mId$(xIn, 150 + 227, 3) & ";" & mId$(xIn, 150 + 230, 2) & ";" & mId$(xIn, 150 + 232, 32) & ";" & mId$(xIn, 150 + 264, 1) & ";" _
    & mId$(xIn, 150 + 265, 7) & ";" & mId$(xIn, 150 + 272, 3) & ";" & mId$(xIn, 150 + 275, 4) & ";" & mId$(xIn, 150 + 279, 8) & ";" _
    & mId$(xIn, 436 + 1, 3) & ";" & mId$(xIn, 436 + 4, 8) & ";" _
    & cur_19V(CCur(mId$(xIn, 436 + 12, 19)) / 1000) & ";" _
    & mId$(xIn, 436 + 31, 15) & ";" & mId$(xIn, 436 + 46, 7) & ";" _
    & mId$(xIn, 436 + 53, 1) & ";" & mId$(xIn, 436 + 54, 1) & ";" _
    & meCV2.Montant
Loop


End Sub



Public Function srvYBIACPT0_Import_Array(lnb As Long, marrYBIACPT0() As typeYBIACPT0)
Dim xIn As String, X As String
Dim intReturn As Integer
On Error GoTo Error_Handle

srvYBIACPT0_Import_Array = "?"
lnb = 0
recYBIACPT0_Init marrYBIACPT0(0)

meYbase.ID = constYBIACPT0
meYbase.K1 = ""
meYbase.Method = "Seek>"
intReturn = tableYBase_Read(meYbase)
'meYBase.Method = "MoveNext"
Do
    If Trim(meYbase.ID) <> constYBIACPT0 Then intReturn = -1
    If intReturn = 0 Then
        lnb = lnb + 1
        MsgTxt = Space$(34) & meYbase.Text
        MsgTxtIndex = 0
        srvYBIACPT0_GetBuffer marrYBIACPT0(lnb)
    End If
    intReturn = tableYBase_Read(meYbase)
   '  If blnJPL And lnb > 500 Then intReturn = -1
Loop Until intReturn <> 0
srvYBIACPT0_Import_Array = Null
Exit Function

Error_Handle:
    MsgBox "erreur : srvYBIACPT0_Import" & xIn, vbCritical, Error
    srvYBIACPT0_Import_Array = Error
End Function

Public Function srvYBIACPT0_Import(lX As String)
Dim xIn As String, X As String, Nb As Long
On Error GoTo Error_Handle


recYBase_Init meYbase
meYbase.Method = "Seek="

meYbase.ID = constYBase
meYbase.K1 = constYBIACPT0
If tableYBase_Read(meYbase) = 0 Then                            'Fichier du jour déjà lu ?
    lX = meYbase.Text
    If mId$(lX, 1, 8) >= YBIATAB0_DATE_CPT_J Then
        srvYBIACPT0_Import = Null
        paramYBIACPT0_Nb = CLng(mId$(lX, 26, 9))
        Exit Function
    Else
        meYbase.Method = constDelete
        Call tableYBase_Update(meYbase)
    End If
End If




srvYBIACPT0_Import = "?"

paramYBIACPT0_Import = paramYBase_DataF & Trim(constYBIACPT0) & paramYBase_Data_ExtensionP

Open Trim(paramYBIACPT0_Import) For Input As #1

Nb = 0
X = "delete * from YBase where Id = " & Chr$(34) & Trim(constYBIACPT0) & Chr$(34)
MDB.Execute X

meYbase.Method = constAddNew


Do Until EOF(1)
    Nb = Nb + 1
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then

            meYbase.ID = constYBIACPT0
            meYbase.K1 = mId$(xIn, 10, 20)  ' .BIACPTCOM .
            meYbase.Text = xIn
            dbYBase_Update meYbase
            
    End If
Loop
Close
srvYBIACPT0_Import = Null
meYbase.ID = constYBase
meYbase.K1 = constYBIACPT0
meYbase.Text = YBIATAB0_DATE_CPT_J & "_" & DSys & "_" & time_Hms & "_" & Format$(Nb, "000000000")
lX = meYbase.Text
dbYBase_Update meYbase
paramYBIACPT0_Nb = Nb

Exit Function

Error_Handle:
 MsgBox "erreur : srvYBIACPT0_Import" & xIn, vbCritical, Error
Close

srvYBIACPT0_Import = Error
End Function



Public Function srvYBIACPT0_Import_Read(lId As String, lYBIACPT0 As typeYBIACPT0)

Dim xIn As String, X As String

On Error GoTo Error_Handle

srvYBIACPT0_Import_Read = "?"

meYbase.Method = "Seek="
meYbase.ID = constYBIACPT0
meYbase.K1 = lId

If tableYBase_Read(meYbase) = 0 Then
    MsgTxt = Space$(34) & meYbase.Text
    MsgTxtIndex = 0
    srvYBIACPT0_GetBuffer lYBIACPT0
    srvYBIACPT0_Import_Read = Null
Else
    recYBIACPT0_Init lYBIACPT0
    lYBIACPT0.COMPTECOM = lId
    lYBIACPT0.COMPTEINT = lId
End If
        
Exit Function

Error_Handle:
 MsgBox "erreur : srvYBIACPT0_Import_Read" & xIn, vbCritical, Error
srvYBIACPT0_Import_Read = Error
End Function






'---------------------------------------------------------
Public Function srvYBIACPT0_GetBuffer(recYBIACPT0 As typeYBIACPT0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYBIACPT0_GetBuffer = Null
recYBIACPT0.Obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYBIACPT0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYBIACPT0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYBIACPT0.Err = Space$(10) Then
    recYBIACPT0.COMPTEETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYBIACPT0.COMPTEPLA = CLng(Val(mId$(MsgTxt, K + 6, 4)))
    recYBIACPT0.COMPTECOM = mId$(MsgTxt, K + 10, 20)
    recYBIACPT0.COMPTEOBL = mId$(MsgTxt, K + 30, 10)
    recYBIACPT0.COMPTEINT = mId$(MsgTxt, K + 40, 32)
    recYBIACPT0.COMPTEAGE = CInt(Val(mId$(MsgTxt, K + 72, 5)))
    recYBIACPT0.COMPTEDEV = mId$(MsgTxt, K + 77, 3)
    recYBIACPT0.COMPTEOUV = CLng(Val(mId$(MsgTxt, K + 80, 8)))
    recYBIACPT0.COMPTECLO = CLng(Val(mId$(MsgTxt, K + 88, 8)))
    recYBIACPT0.COMPTELOR = mId$(MsgTxt, K + 96, 1)
    recYBIACPT0.COMPTESUC = mId$(MsgTxt, K + 97, 1)
    recYBIACPT0.COMPTECLA = CLng(Val(mId$(MsgTxt, K + 98, 3)))
    recYBIACPT0.COMPTEFON = mId$(MsgTxt, K + 101, 1)
    recYBIACPT0.COMPTEBLO = CLng(Val(mId$(MsgTxt, K + 102, 8)))
    recYBIACPT0.COMPTEMOT = mId$(MsgTxt, K + 110, 32)
    recYBIACPT0.COMPTESEN = mId$(MsgTxt, K + 142, 1)
    recYBIACPT0.COMPTEMOD = CLng(Val(mId$(MsgTxt, K + 143, 8)))
    
    K = MsgTxtIndex + 34 + 150
    
    recYBIACPT0.CLIENAETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYBIACPT0.CLIENACLI = mId$(MsgTxt, K + 6, 7)
    recYBIACPT0.CLIENAAGE = CInt(Val(mId$(MsgTxt, K + 13, 5)))
    recYBIACPT0.CLIENAETA = mId$(MsgTxt, K + 18, 4)
    recYBIACPT0.CLIENARA1 = mId$(MsgTxt, K + 22, 32)
    recYBIACPT0.CLIENARA2 = mId$(MsgTxt, K + 54, 32)
    recYBIACPT0.CLIENASIG = mId$(MsgTxt, K + 86, 12)
    recYBIACPT0.CLIENASRN = mId$(MsgTxt, K + 98, 9)
    recYBIACPT0.CLIENASRT = CLng(Val(mId$(MsgTxt, K + 107, 6)))
    recYBIACPT0.CLIENADNA = CLng(Val(mId$(MsgTxt, K + 113, 8)))
    recYBIACPT0.CLIENAREG = mId$(MsgTxt, K + 121, 6)
    recYBIACPT0.CLIENANAT = mId$(MsgTxt, K + 127, 3)
    recYBIACPT0.CLIENARSD = mId$(MsgTxt, K + 130, 3)
    recYBIACPT0.CLIENARES = mId$(MsgTxt, K + 133, 3)
    recYBIACPT0.CLIENAECO = mId$(MsgTxt, K + 136, 3)
    recYBIACPT0.CLIENAACT = mId$(MsgTxt, K + 139, 1)
    recYBIACPT0.CLIENAPAI = mId$(MsgTxt, K + 140, 1)
    recYBIACPT0.CLIENACRD = mId$(MsgTxt, K + 141, 1)
    recYBIACPT0.CLIENAADM = mId$(MsgTxt, K + 142, 1)
    recYBIACPT0.CLIENAATR = CLng(Val(mId$(MsgTxt, K + 143, 8)))
    recYBIACPT0.CLIENABIL = CLng(Val(mId$(MsgTxt, K + 151, 4)))
    recYBIACPT0.CLIENACAT = mId$(MsgTxt, K + 155, 3)
    recYBIACPT0.CLIENACOT = mId$(MsgTxt, K + 158, 3)
    recYBIACPT0.CLIENACHQ = mId$(MsgTxt, K + 161, 1)
    recYBIACPT0.CLIENADAT = CLng(Val(mId$(MsgTxt, K + 162, 8)))
    recYBIACPT0.CLIENASAC = mId$(MsgTxt, K + 170, 6)
    recYBIACPT0.CLIENAGEO = mId$(MsgTxt, K + 176, 3)
    recYBIACPT0.CLIENAENT = mId$(MsgTxt, K + 179, 3)
    recYBIACPT0.CLIENAMES = mId$(MsgTxt, K + 182, 1)
    recYBIACPT0.CLIENAPAY = CLng(Val(mId$(MsgTxt, K + 183, 8)))
    recYBIACPT0.CLIENAFIL = mId$(MsgTxt, K + 191, 32)
    recYBIACPT0.CLIENABIM = CLng(Val(mId$(MsgTxt, K + 223, 3)))
    recYBIACPT0.CLIENADOU = mId$(MsgTxt, K + 226, 1)
    recYBIACPT0.CLIENALI1 = mId$(MsgTxt, K + 227, 3)
    recYBIACPT0.CLIENALI2 = mId$(MsgTxt, K + 230, 2)
    recYBIACPT0.CLIENAEXT = mId$(MsgTxt, K + 232, 32)
    recYBIACPT0.CLIENACOL = mId$(MsgTxt, K + 264, 1)
    recYBIACPT0.CLIENATIE = mId$(MsgTxt, K + 265, 7)
    recYBIACPT0.CLIENASEL = mId$(MsgTxt, K + 272, 3)
    recYBIACPT0.CLIENAPCS = mId$(MsgTxt, K + 275, 4)
    recYBIACPT0.CLIENACRE = CLng(Val(mId$(MsgTxt, K + 279, 8)))
    
    K = MsgTxtIndex + 34 + 436
    
    recYBIACPT0.PLANCOPRO = mId$(MsgTxt, K + 1, 3)
    recYBIACPT0.SOLDEDMO = CLng(Val(mId$(MsgTxt, K + 4, 8)))
    recYBIACPT0.SOLDECEN = CCur(mId$(MsgTxt, K + 12, 19)) / 1000
    recYBIACPT0.COMREFREF = mId$(MsgTxt, K + 31, 15)
    recYBIACPT0.TITULACLI = mId$(MsgTxt, K + 46, 7)
    recYBIACPT0.TITULAPRI = mId$(MsgTxt, K + 53, 1)
    recYBIACPT0.TITULATPR = mId$(MsgTxt, K + 54, 1)

Else
    srvYBIACPT0_GetBuffer = recYBIACPT0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYBIACPT0Len

End Function

'---------------------------------------------------------
Public Function srvYBIACPT0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYBIACPT0 As typeYBIACPT0)
'---------------------------------------------------------
On Error Resume Next 'GoTo Error_Handler
srvYBIACPT0_GetBuffer_ODBC = Null


    recYBIACPT0.COMPTEETA = rsADO("COMPTEETA") 'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYBIACPT0.COMPTEPLA = rsADO("COMPTEPLA") 'CLng(Val(mId$(MsgTxt, K + 6, 4)))
    recYBIACPT0.COMPTECOM = rsADO("COMPTECOM") 'mId$(MsgTxt, K + 10, 20)
    recYBIACPT0.COMPTEOBL = rsADO("COMPTEOBL") 'mId$(MsgTxt, K + 30, 10)
    recYBIACPT0.COMPTEINT = rsADO("COMPTEINT") 'mId$(MsgTxt, K + 40, 32)
    recYBIACPT0.COMPTEAGE = rsADO("COMPTEAGE") 'CInt(Val(mId$(MsgTxt, K + 72, 5)))
    recYBIACPT0.COMPTEDEV = rsADO("COMPTEDEV") 'mId$(MsgTxt, K + 77, 3)
    recYBIACPT0.COMPTEOUV = rsADO("COMPTEOUV") 'CLng(Val(mId$(MsgTxt, K + 80, 8)))
    recYBIACPT0.COMPTECLO = rsADO("COMPTECLO") 'CLng(Val(mId$(MsgTxt, K + 88, 8)))
    recYBIACPT0.COMPTELOR = rsADO("COMPTELOR") 'mId$(MsgTxt, K + 96, 1)
    recYBIACPT0.COMPTESUC = rsADO("COMPTESUC") 'mId$(MsgTxt, K + 97, 1)
    recYBIACPT0.COMPTECLA = rsADO("COMPTECLA") 'CLng(Val(mId$(MsgTxt, K + 98, 3)))
    recYBIACPT0.COMPTEFON = rsADO("COMPTEFON") 'mId$(MsgTxt, K + 101, 1)
    recYBIACPT0.COMPTEBLO = rsADO("COMPTEBLO") 'CLng(Val(mId$(MsgTxt, K + 102, 8)))
    recYBIACPT0.COMPTEMOT = rsADO("COMPTEMOT") 'mId$(MsgTxt, K + 110, 32)
    recYBIACPT0.COMPTESEN = rsADO("COMPTESEN") 'mId$(MsgTxt, K + 142, 1)
    recYBIACPT0.COMPTEMOD = rsADO("COMPTEMOD") 'CLng(Val(mId$(MsgTxt, K + 143, 8)))
    
    recYBIACPT0.CLIENAETB = rsADO("CLIENAETB") 'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYBIACPT0.CLIENACLI = rsADO("CLIENACLI") 'mId$(MsgTxt, K + 6, 7)
    recYBIACPT0.CLIENAAGE = rsADO("CLIENAAGE") 'CInt(Val(mId$(MsgTxt, K + 13, 5)))
    recYBIACPT0.CLIENAETA = rsADO("CLIENAETA") 'mId$(MsgTxt, K + 18, 4)
    recYBIACPT0.CLIENARA1 = rsADO("CLIENARA1") 'mId$(MsgTxt, K + 22, 32)
    recYBIACPT0.CLIENARA2 = rsADO("CLIENARA2") 'mId$(MsgTxt, K + 54, 32)
    recYBIACPT0.CLIENASIG = rsADO("CLIENASIG") 'mId$(MsgTxt, K + 86, 12)
    recYBIACPT0.CLIENASRN = rsADO("CLIENASRN") 'mId$(MsgTxt, K + 98, 9)
    recYBIACPT0.CLIENASRT = rsADO("CLIENASRT") 'CLng(Val(mId$(MsgTxt, K + 107, 6)))
    recYBIACPT0.CLIENADNA = rsADO("CLIENADNA") 'CLng(Val(mId$(MsgTxt, K + 113, 8)))
    recYBIACPT0.CLIENAREG = rsADO("CLIENAREG") 'mId$(MsgTxt, K + 121, 6)
    recYBIACPT0.CLIENANAT = rsADO("CLIENANAT") 'mId$(MsgTxt, K + 127, 3)
    recYBIACPT0.CLIENARSD = rsADO("CLIENARSD") 'mId$(MsgTxt, K + 130, 3)
    recYBIACPT0.CLIENARES = rsADO("CLIENARES") 'mId$(MsgTxt, K + 133, 3)
    recYBIACPT0.CLIENAECO = rsADO("CLIENAECO") 'mId$(MsgTxt, K + 136, 3)
    recYBIACPT0.CLIENAACT = rsADO("CLIENAACT") 'mId$(MsgTxt, K + 139, 1)
    recYBIACPT0.CLIENAPAI = rsADO("CLIENAPAI") 'mId$(MsgTxt, K + 140, 1)
    recYBIACPT0.CLIENACRD = rsADO("CLIENACRD") 'mId$(MsgTxt, K + 141, 1)
    recYBIACPT0.CLIENAADM = rsADO("CLIENAADM") 'mId$(MsgTxt, K + 142, 1)
    recYBIACPT0.CLIENAATR = rsADO("CLIENAATR") 'CLng(Val(mId$(MsgTxt, K + 143, 8)))
    recYBIACPT0.CLIENABIL = rsADO("CLIENABIL") 'CLng(Val(mId$(MsgTxt, K + 151, 4)))
    recYBIACPT0.CLIENACAT = rsADO("CLIENACAT") 'mId$(MsgTxt, K + 155, 3)
    recYBIACPT0.CLIENACOT = rsADO("CLIENACOT") 'mId$(MsgTxt, K + 158, 3)
    recYBIACPT0.CLIENACHQ = rsADO("CLIENACHQ") 'mId$(MsgTxt, K + 161, 1)
    recYBIACPT0.CLIENADAT = rsADO("CLIENADAT") 'CLng(Val(mId$(MsgTxt, K + 162, 8)))
    recYBIACPT0.CLIENASAC = rsADO("CLIENASAC") 'mId$(MsgTxt, K + 170, 6)
    recYBIACPT0.CLIENAGEO = rsADO("CLIENAGEO") 'mId$(MsgTxt, K + 176, 3)
    recYBIACPT0.CLIENAENT = rsADO("CLIENAENT") 'mId$(MsgTxt, K + 179, 3)
    recYBIACPT0.CLIENAMES = rsADO("CLIENAMES") 'mId$(MsgTxt, K + 182, 1)
    recYBIACPT0.CLIENAPAY = rsADO("CLIENAPAY") 'CLng(Val(mId$(MsgTxt, K + 183, 8)))
    recYBIACPT0.CLIENAFIL = rsADO("CLIENAFIL") 'mId$(MsgTxt, K + 191, 32)
    recYBIACPT0.CLIENABIM = rsADO("CLIENABIM") 'CLng(Val(mId$(MsgTxt, K + 223, 3)))
    recYBIACPT0.CLIENADOU = rsADO("CLIENADOU") 'mId$(MsgTxt, K + 226, 1)
    recYBIACPT0.CLIENALI1 = rsADO("CLIENALI1") 'mId$(MsgTxt, K + 227, 3)
    recYBIACPT0.CLIENALI2 = rsADO("CLIENALI2") 'mId$(MsgTxt, K + 230, 2)
    recYBIACPT0.CLIENAEXT = rsADO("CLIENAEXT") 'mId$(MsgTxt, K + 232, 32)
    recYBIACPT0.CLIENACOL = rsADO("CLIENACOL") 'mId$(MsgTxt, K + 264, 1)
    recYBIACPT0.CLIENATIE = rsADO("CLIENATIE") 'mId$(MsgTxt, K + 265, 7)
    recYBIACPT0.CLIENASEL = rsADO("CLIENASEL") 'mId$(MsgTxt, K + 272, 3)
    recYBIACPT0.CLIENAPCS = rsADO("CLIENAPCS") 'mId$(MsgTxt, K + 275, 4)
    recYBIACPT0.CLIENACRE = rsADO("CLIENACRE") 'CLng(Val(mId$(MsgTxt, K + 279, 8)))
    
    recYBIACPT0.PLANCOPRO = rsADO("PLANCOPRO") 'mId$(MsgTxt, K + 1, 3)
    recYBIACPT0.SOLDEDMO = rsADO("SOLDEDMO") 'CLng(Val(mId$(MsgTxt, K + 4, 8)))
    recYBIACPT0.SOLDECEN = rsADO("SOLDECEN") / 1000 'CCur(mId$(MsgTxt, K + 12, 19)) / 1000
    recYBIACPT0.COMREFREF = rsADO("COMREFREF") 'mId$(MsgTxt, K + 31, 15)
    recYBIACPT0.TITULACLI = rsADO("TITULACLI") 'mId$(MsgTxt, K + 46, 7)
    recYBIACPT0.TITULAPRI = rsADO("TITULAPRI") 'mId$(MsgTxt, K + 53, 1)
    recYBIACPT0.TITULATPR = rsADO("TITULATPR") 'mId$(MsgTxt, K + 54, 1)

Exit Function

Error_Handler:
srvYBIACPT0_GetBuffer_ODBC = Error

End Function

'---------------------------------------------------------
Public Sub srvYBIACPT0_PutBuffer(recYBIACPT0 As typeYBIACPT0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYBIACPT0.Obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYBIACPT0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYBIACPT0.COMPTEETA, "0000 ")
    Mid$(MsgTxt, K + 6, 4) = Format$(recYBIACPT0.COMPTEPLA, "000 ")
    Mid$(MsgTxt, K + 10, 20) = recYBIACPT0.COMPTECOM
    Mid$(MsgTxt, K + 30, 10) = recYBIACPT0.COMPTEOBL
    Mid$(MsgTxt, K + 40, 32) = recYBIACPT0.COMPTEINT
    Mid$(MsgTxt, K + 72, 5) = Format$(recYBIACPT0.COMPTEAGE, "0000 ")
    Mid$(MsgTxt, K + 77, 3) = recYBIACPT0.COMPTEDEV
    Mid$(MsgTxt, K + 80, 8) = Format$(recYBIACPT0.COMPTEOUV, "0000000 ")
    Mid$(MsgTxt, K + 88, 8) = Format$(recYBIACPT0.COMPTECLO, "0000000 ")
    Mid$(MsgTxt, K + 96, 1) = recYBIACPT0.COMPTELOR
    Mid$(MsgTxt, K + 97, 1) = recYBIACPT0.COMPTESUC
    Mid$(MsgTxt, K + 98, 3) = Format$(recYBIACPT0.COMPTECLA, "00 ")
    Mid$(MsgTxt, K + 101, 1) = recYBIACPT0.COMPTEFON
    Mid$(MsgTxt, K + 102, 8) = Format$(recYBIACPT0.COMPTEBLO, "0000000 ")
    Mid$(MsgTxt, K + 110, 32) = recYBIACPT0.COMPTEMOT
    Mid$(MsgTxt, K + 142, 1) = recYBIACPT0.COMPTESEN
    Mid$(MsgTxt, K + 143, 8) = Format$(recYBIACPT0.COMPTEMOD, "0000000 ")

MsgTxtLen = MsgTxtLen + recYBIACPT0Len
End Sub



'---------------------------------------------------------
Public Sub recYBIACPT0_Init(recYBIACPT0 As typeYBIACPT0)
'---------------------------------------------------------
recYBIACPT0.Obj = "ZBIACPT0_S"
recYBIACPT0.Method = ""
recYBIACPT0.Err = ""
recYBIACPT0.COMPTEETA = 1

recYBIACPT0.COMPTEPLA = 0 '       As Long                           ' NUMERO PLAN
recYBIACPT0.COMPTECOM = "" '       As String * 20                    ' NUMERO COMPTE
recYBIACPT0.COMPTEOBL = "" '       As String * 10                    ' COMPTE OBLIGATOIRE
recYBIACPT0.COMPTEINT = "" '       As String * 32                    ' INTITULE
recYBIACPT0.COMPTEAGE = 0 '       As Integer                        ' AGENCE
recYBIACPT0.COMPTEDEV = "" '       As String * 3                     ' TABLES BASE 013
recYBIACPT0.COMPTEOUV = 0 '       As Long                           ' DATE OUVERTURE
recYBIACPT0.COMPTECLO = 0 '       As Long                           ' DATE CLOTURE
recYBIACPT0.COMPTELOR = "" '       As String * 1                     ' Lori/Nostri/AUTRE
recYBIACPT0.COMPTESUC = "" '       As String * 1                     ' O/N
recYBIACPT0.COMPTECLA = 0 '       As Long                           ' CLASSE SECURITE
recYBIACPT0.COMPTEFON = "" '       As String * 1                     ' TABLES BASE 015
recYBIACPT0.COMPTEBLO = 0 '       As Long                           ' DATE LIMITE BLOCAGE
recYBIACPT0.COMPTEMOT = "" '       As String * 32                    ' MOTIF BLOCAGE
recYBIACPT0.COMPTESEN = "" '       As String * 1                     ' CODE SENS SOLDE D/C
recYBIACPT0.COMPTEMOD = 0 '       As Long                           ' DATE MODIFICATION
    
    
recYBIACPT0.CLIENAETB = 0 '       As Integer                        ' CODE ETABLISSEMENT
recYBIACPT0.CLIENACLI = "" '       As String * 7                     ' NUMERO CLIENT
recYBIACPT0.CLIENAAGE = 0 '       As Integer                        ' CODE AGENCE
recYBIACPT0.CLIENAETA = "" '       As String * 4                     ' CODE ETAT
recYBIACPT0.CLIENARA1 = "" '       As String * 32                    ' NOM OU DESIGNATION
recYBIACPT0.CLIENARA2 = "" '       As String * 32                    ' PRENOM/DESIGNATION
recYBIACPT0.CLIENASIG = "" '       As String * 12                    ' SIGLE USUEL
recYBIACPT0.CLIENASRN = "" '       As String * 9                     ' NUMERO SIREN
recYBIACPT0.CLIENASRT = 0 '       As Long                           ' NUMERO SIRET
recYBIACPT0.CLIENADNA = 0 '       As Long                           ' DATE DE NAISSANCE
recYBIACPT0.CLIENAREG = "" '       As String * 6                     ' SECT ACTIVITE REGLEM
recYBIACPT0.CLIENANAT = "" '       As String * 3                     ' CDE PAYS NATIONALITE
recYBIACPT0.CLIENARSD = "" '       As String * 3                     ' CDE PAYS DE RESIDENC
recYBIACPT0.CLIENARES = "" '       As String * 3                     ' RESPONS/EXPLOITATION
recYBIACPT0.CLIENAECO = "" '       As String * 3                     ' QUALITE/AG ECONOMIQU
recYBIACPT0.CLIENAACT = "" '       As String * 1                     ' COTE ACTIVITE
recYBIACPT0.CLIENAPAI = "" '       As String * 1                     ' COTE PAIEMENT
recYBIACPT0.CLIENACRD = "" '       As String * 1                     ' COTE CREDIT
recYBIACPT0.CLIENAADM = "" '       As String * 1                     ' COTE ADMISSION
recYBIACPT0.CLIENAATR = 0 '       As Long                           ' DAT ATRIB/COTAT BDF
recYBIACPT0.CLIENABIL = 0 '       As Long                           ' AN DERN BIL COMM BDF
recYBIACPT0.CLIENACAT = "" '       As String * 3                     ' CATEGORIE CLIENT
recYBIACPT0.CLIENACOT = "" '       As String * 3                     ' COTATION INTERNE
recYBIACPT0.CLIENACHQ = "" '       As String * 1                     ' INTERDICTION CHEQUIE
recYBIACPT0.CLIENADAT = 0 '       As Long                           ' INTERDIT CHEQUIER
recYBIACPT0.CLIENASAC = "" '       As String * 6                     ' SECTEUR D ACTIVITE
recYBIACPT0.CLIENAGEO = "" '       As String * 3                     ' SECTEUR GEOGRAPHIQUE
recYBIACPT0.CLIENAENT = "" '       As String * 3                     ' ENTREPRISE LIEE
recYBIACPT0.CLIENAMES = "" '       As String * 1                     ' LANGUE MESSAGERIE
recYBIACPT0.CLIENAPAY = 0 '       As Long                           ' DATE ENTREE AU PAYS
recYBIACPT0.CLIENAFIL = "" '       As String * 32                    ' NOM DE JEUNE FILLE
recYBIACPT0.CLIENABIM = 0 '       As Long                           ' BILAN DE MOIS
recYBIACPT0.CLIENADOU = "" '       As String * 1                     ' CLIENT DOUTEUX O/N
recYBIACPT0.CLIENALI1 = "" '       As String * 3                     ' ZONE LIBRE DE 3 CAR.
recYBIACPT0.CLIENALI2 = "" '       As String * 2                     ' ZONE LIBRE DE 2 CAR.
recYBIACPT0.CLIENAEXT = "" '       As String * 32                    ' EXTENTION DU NOM
recYBIACPT0.CLIENACOL = "" '       As String * 1                     ' 0=CLI/COLL=1/AUTRE=2
recYBIACPT0.CLIENATIE = "" '       As String * 7                     ' TIERS DE REFERENCE
recYBIACPT0.CLIENASEL = "" '       As String * 3                     ' CODE SELECTION
recYBIACPT0.CLIENAPCS = "" '       As String * 4                     ' CODE PCS
recYBIACPT0.CLIENACRE = 0 '       As Long                           ' DATE CREATION
    
recYBIACPT0.PLANCOPRO = "" '       As String * 3                     ' TABLES BASE 014
    
recYBIACPT0.SOLDEDMO = 0 '       As Long                           ' DATE DERNIER MVT
recYBIACPT0.SOLDECEN = 0 '       As Currency                       ' SOLDE ENCOURS
recYBIACPT0.COMREFREF = "" '       As String * 15                    ' EX référence

recYBIACPT0.TITULACLI = "" '       As String * 7                     ' NUMERO CLIENT
recYBIACPT0.TITULAPRI = "" '       As String * 1                     ' 0:PRINCIPAL, 1:AUTRE
recYBIACPT0.TITULATPR = "" '       As String * 1                     ' 0:PRINCIPAL, 1:AUTRE

End Sub







