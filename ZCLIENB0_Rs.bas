Attribute VB_Name = "rsZCLIENB0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCLIENB0
    CLIENBETB       As Integer                        ' CODE ETABLISSEMENT
    CLIENBCLI       As String * 7                     ' NUMERO CLIENT
    CLIENBCRT       As Long                           ' DATE CREATION ENTREP
    CLIENBBIL       As Long                           ' DAT COMM PROCH BILAN
    CLIENBEFC       As Long                           ' EFFECTIF
    CLIENBCH1       As Long                           ' CHIFR AFF EN MILLIER
    CLIENBCH2       As Long                           ' CHIFR AFF EN MILLIER
    CLIENBCH3       As Long                           ' CHIFR AFF EN MILLIER
    CLIENBAF1       As String * 2                     ' ANNEE CHIFRE AFFAIRE
    CLIENBAF2       As String * 2                     ' ANNEE CHIFRE AFFAIRE
    CLIENBAF3       As String * 2                     ' ANNEE CHIFRE AFFAIRE
    CLIENBCP1       As Long                           ' CAPITAUX EN UNITES
    CLIENBMD1       As Long                           ' DATE DE MODIFICATION
    CLIENBNAS       As String * 3                     ' CODE PAYS NAISSANCE
    CLIENBINS       As String * 32                    ' DEP NAISSANCE INSEE
    CLIENBCOM       As String * 32                    ' COMMUNE DE NAISSANCE
    CLIENBLIE       As String * 1                     ' CODE LIEU NAISSANCE
    CLIENBTER       As String * 1                     ' CODE TERRITORIALITE
    CLIENBPER       As String * 32                    ' AUTRES PRENOMS
    CLIENBMAR       As String * 32                    ' NOM / PRENOM DE MARI
    CLIENBJUR       As String * 2                     ' CAPACITE JURIDIQUE
    CLIENBCAP       As String * 3                     ' DEVISE DU CAPITAL
    CLIENBBAN       As String * 2                     ' ORIGINE INTERDIT BAN
    CLIENBLIB       As String * 3                     ' ZONE LIBRE
    CLIENBDED       As String * 1                     ' DEDOM
    CLIENBSER       As String * 3                     ' SEGMENT DE RÉSULTAT
    CLIENBSEP       As String * 3                     ' SEGMENT POTENTIEL
    CLIENBCTL       As String * 1                     ' CODE CONTRÔLE
    CLIENBMUT       As Long                           ' DATE DE MUTATION
    CLIENBDEC       As Long                           ' DATE DE DÉCÈS
    CLIENBCIN       As String * 5                     ' CODE INSEE
    CLIENBTOP       As String * 1                     ' IND PGM DE RÉVISION
'    FILLER          As String * 37                    '

End Type
Public Sub rsZCLIENB0_Init(rsZCLIENB0 As typeZCLIENB0)
rsZCLIENB0.CLIENBETB = 0
rsZCLIENB0.CLIENBCLI = ""
rsZCLIENB0.CLIENBCRT = 0
rsZCLIENB0.CLIENBBIL = 0
rsZCLIENB0.CLIENBEFC = 0
rsZCLIENB0.CLIENBCH1 = 0
rsZCLIENB0.CLIENBCH2 = 0
rsZCLIENB0.CLIENBCH3 = 0
rsZCLIENB0.CLIENBAF1 = "00"
rsZCLIENB0.CLIENBAF2 = "00"
rsZCLIENB0.CLIENBAF3 = "00"
rsZCLIENB0.CLIENBCP1 = 0
rsZCLIENB0.CLIENBMD1 = 0
rsZCLIENB0.CLIENBNAS = ""
rsZCLIENB0.CLIENBINS = ""
rsZCLIENB0.CLIENBCOM = ""
rsZCLIENB0.CLIENBLIE = ""
rsZCLIENB0.CLIENBTER = ""
rsZCLIENB0.CLIENBPER = ""
rsZCLIENB0.CLIENBMAR = ""
rsZCLIENB0.CLIENBJUR = ""
rsZCLIENB0.CLIENBCAP = ""
rsZCLIENB0.CLIENBBAN = ""
rsZCLIENB0.CLIENBLIB = ""
rsZCLIENB0.CLIENBDED = ""
rsZCLIENB0.CLIENBSER = ""
rsZCLIENB0.CLIENBSEP = ""
rsZCLIENB0.CLIENBCTL = ""
rsZCLIENB0.CLIENBMUT = 0
rsZCLIENB0.CLIENBDEC = 0
rsZCLIENB0.CLIENBCIN = ""
rsZCLIENB0.CLIENBTOP = ""
'rsZCLIENB0.FILLER = ""
End Sub
Public Function rsZCLIENB0_GetBuffer(rsAdo As ADODB.Recordset, rsZCLIENB0 As typeZCLIENB0)
On Error GoTo Error_Handler
rsZCLIENB0_GetBuffer = Null
rsZCLIENB0.CLIENBETB = rsAdo("CLIENBETB")
rsZCLIENB0.CLIENBCLI = rsAdo("CLIENBCLI")
rsZCLIENB0.CLIENBCRT = rsAdo("CLIENBCRT")
rsZCLIENB0.CLIENBBIL = rsAdo("CLIENBBIL")
rsZCLIENB0.CLIENBEFC = rsAdo("CLIENBEFC")
rsZCLIENB0.CLIENBCH1 = rsAdo("CLIENBCH1")
rsZCLIENB0.CLIENBCH2 = rsAdo("CLIENBCH2")
rsZCLIENB0.CLIENBCH3 = rsAdo("CLIENBCH3")
rsZCLIENB0.CLIENBAF1 = rsAdo("CLIENBAF1")
rsZCLIENB0.CLIENBAF2 = rsAdo("CLIENBAF2")
rsZCLIENB0.CLIENBAF3 = rsAdo("CLIENBAF3")
rsZCLIENB0.CLIENBCP1 = rsAdo("CLIENBCP1")
rsZCLIENB0.CLIENBMD1 = rsAdo("CLIENBMD1")
rsZCLIENB0.CLIENBNAS = rsAdo("CLIENBNAS")
rsZCLIENB0.CLIENBINS = rsAdo("CLIENBINS")
rsZCLIENB0.CLIENBCOM = rsAdo("CLIENBCOM")
rsZCLIENB0.CLIENBLIE = rsAdo("CLIENBLIE")
rsZCLIENB0.CLIENBTER = rsAdo("CLIENBTER")
rsZCLIENB0.CLIENBPER = rsAdo("CLIENBPER")
rsZCLIENB0.CLIENBMAR = rsAdo("CLIENBMAR")
rsZCLIENB0.CLIENBJUR = rsAdo("CLIENBJUR")
rsZCLIENB0.CLIENBCAP = rsAdo("CLIENBCAP")
rsZCLIENB0.CLIENBBAN = rsAdo("CLIENBBAN")
rsZCLIENB0.CLIENBLIB = rsAdo("CLIENBLIB")
rsZCLIENB0.CLIENBDED = rsAdo("CLIENBDED")
rsZCLIENB0.CLIENBSER = rsAdo("CLIENBSER")
rsZCLIENB0.CLIENBSEP = rsAdo("CLIENBSEP")
rsZCLIENB0.CLIENBCTL = rsAdo("CLIENBCTL")
rsZCLIENB0.CLIENBMUT = rsAdo("CLIENBMUT")
rsZCLIENB0.CLIENBDEC = rsAdo("CLIENBDEC")
rsZCLIENB0.CLIENBCIN = rsAdo("CLIENBCIN")
rsZCLIENB0.CLIENBTOP = rsAdo("CLIENBTOP")
'rsZCLIENB0.FILLER = rsADO("FILLER")
Exit Function
Error_Handler:
rsZCLIENB0_GetBuffer = Error
End Function
