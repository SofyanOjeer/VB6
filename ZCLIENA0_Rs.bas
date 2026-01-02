Attribute VB_Name = "rsZCLIENA0"
'---------------------------------------------------------
Option Explicit
Type typeZCLIENA0

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

End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCLIENA0_GetBuffer(rsAdo As ADODB.Recordset, rszCLIENA0 As typeZCLIENA0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCLIENA0_GetBuffer = Null

rszCLIENA0.CLIENAETB = rsAdo("CLIENAETB")
rszCLIENA0.CLIENACLI = rsAdo("CLIENACLI")
rszCLIENA0.CLIENAAGE = rsAdo("CLIENAAGE")
rszCLIENA0.CLIENAETA = rsAdo("CLIENAETA")
rszCLIENA0.CLIENARA1 = rsAdo("CLIENARA1")
rszCLIENA0.CLIENARA2 = rsAdo("CLIENARA2")
rszCLIENA0.CLIENASIG = rsAdo("CLIENASIG")
rszCLIENA0.CLIENASRN = rsAdo("CLIENASRN")
rszCLIENA0.CLIENASRT = rsAdo("CLIENASRT")
rszCLIENA0.CLIENADNA = rsAdo("CLIENADNA")
rszCLIENA0.CLIENAREG = rsAdo("CLIENAREG")
rszCLIENA0.CLIENANAT = rsAdo("CLIENANAT")
rszCLIENA0.CLIENARSD = rsAdo("CLIENARSD")
rszCLIENA0.CLIENARES = rsAdo("CLIENARES")
rszCLIENA0.CLIENAECO = rsAdo("CLIENAECO")
rszCLIENA0.CLIENAACT = rsAdo("CLIENAACT")
rszCLIENA0.CLIENAPAI = rsAdo("CLIENAPAI")
rszCLIENA0.CLIENACRD = rsAdo("CLIENACRD")
rszCLIENA0.CLIENAADM = rsAdo("CLIENAADM")
rszCLIENA0.CLIENAATR = rsAdo("CLIENAATR")
rszCLIENA0.CLIENABIL = rsAdo("CLIENABIL")
rszCLIENA0.CLIENACAT = rsAdo("CLIENACAT")
rszCLIENA0.CLIENACOT = rsAdo("CLIENACOT")
rszCLIENA0.CLIENACHQ = rsAdo("CLIENACHQ")
rszCLIENA0.CLIENADAT = rsAdo("CLIENADAT")
rszCLIENA0.CLIENASAC = rsAdo("CLIENASAC")
rszCLIENA0.CLIENAGEO = rsAdo("CLIENAGEO")
rszCLIENA0.CLIENAENT = rsAdo("CLIENAENT")
rszCLIENA0.CLIENAMES = rsAdo("CLIENAMES")
rszCLIENA0.CLIENAPAY = rsAdo("CLIENAPAY")
rszCLIENA0.CLIENAFIL = rsAdo("CLIENAFIL")
rszCLIENA0.CLIENABIM = rsAdo("CLIENABIM")
rszCLIENA0.CLIENADOU = rsAdo("CLIENADOU")
rszCLIENA0.CLIENALI1 = rsAdo("CLIENALI1")
rszCLIENA0.CLIENALI2 = rsAdo("CLIENALI2")
rszCLIENA0.CLIENAEXT = rsAdo("CLIENAEXT")
rszCLIENA0.CLIENACOL = rsAdo("CLIENACOL")
rszCLIENA0.CLIENATIE = rsAdo("CLIENATIE")
rszCLIENA0.CLIENASEL = rsAdo("CLIENASEL")
rszCLIENA0.CLIENAPCS = rsAdo("CLIENAPCS")
rszCLIENA0.CLIENACRE = rsAdo("CLIENACRE")
Exit Function

Error_Handler:

rsZCLIENA0_GetBuffer = Error

End Function

Public Sub rsZCLIENA0_Init(rszCLIENA0 As typeZCLIENA0)
rszCLIENA0.CLIENAETB = 0
rszCLIENA0.CLIENACLI = ""
rszCLIENA0.CLIENAAGE = 0
rszCLIENA0.CLIENAETA = ""
rszCLIENA0.CLIENARA1 = ""
rszCLIENA0.CLIENARA2 = ""
rszCLIENA0.CLIENASIG = ""
rszCLIENA0.CLIENASRN = ""
rszCLIENA0.CLIENASRT = 0
rszCLIENA0.CLIENADNA = 0
rszCLIENA0.CLIENAREG = ""
rszCLIENA0.CLIENANAT = ""
rszCLIENA0.CLIENARSD = ""
rszCLIENA0.CLIENARES = ""
rszCLIENA0.CLIENAECO = ""
rszCLIENA0.CLIENAACT = ""
rszCLIENA0.CLIENAPAI = ""
rszCLIENA0.CLIENACRD = ""
rszCLIENA0.CLIENAADM = ""
rszCLIENA0.CLIENAATR = 0
rszCLIENA0.CLIENABIL = 0
rszCLIENA0.CLIENACAT = ""
rszCLIENA0.CLIENACOT = ""
rszCLIENA0.CLIENACHQ = ""
rszCLIENA0.CLIENADAT = 0
rszCLIENA0.CLIENASAC = ""
rszCLIENA0.CLIENAGEO = ""
rszCLIENA0.CLIENAENT = ""
rszCLIENA0.CLIENAMES = ""
rszCLIENA0.CLIENAPAY = 0
rszCLIENA0.CLIENAFIL = ""
rszCLIENA0.CLIENABIM = 0
rszCLIENA0.CLIENADOU = ""
rszCLIENA0.CLIENALI1 = ""
rszCLIENA0.CLIENALI2 = ""
rszCLIENA0.CLIENAEXT = ""
rszCLIENA0.CLIENACOL = ""
rszCLIENA0.CLIENATIE = ""
rszCLIENA0.CLIENASEL = ""
rszCLIENA0.CLIENAPCS = ""
rszCLIENA0.CLIENACRE = 0
End Sub


