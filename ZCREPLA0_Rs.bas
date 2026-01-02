Attribute VB_Name = "rsZCREPLA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCREPLA0
    CREPLAETA       As Integer                        ' ETABLISSEMENT
    CREPLAAGE       As Integer                        ' AGENCE
    CREPLASER       As String * 2                     ' SERVICE
    CREPLASSE       As String * 2                     ' SOUS-SERVICE
    CREPLADOS       As Long                           ' NUMERO DOSSIER
    CREPLAPRE       As Long                           ' N° PRET
    CREPLAPLA       As Long                           ' N° PLAN
    CREPLAMAM       As Currency                       ' MONTANT AMORTI
    CREPLAMIN       As Currency                       ' MONTANT INTERETS
    CREPLAMOA       As String * 1                     ' TYPE REMBOURSEMENT
    CREPLANPC       As Long                           ' NB PERIODES CAPITAL
    CREPLAPCA       As String * 1                     ' PERIODICITE CAPITAL
    CREPLADEC       As Long                           ' DATE 1° CAPITAL
    CREPLADRE       As String * 2                     ' DATE REF. CAPITAL
    CREPLAJEC       As Long                           ' JOUR ECH. CAPITAL
    CREPLADTO       As String * 1                     ' DIFFERE TOTAL
    CREPLADAM       As String * 1                     ' DIFFERE D AMORTIS
    CREPLANPE       As Long                           ' NB PERIO. DIFFERE
    CREPLAPIN       As String * 1                     ' INTERETS SEPARES
    CREPLAPEI       As String * 1                     ' PERIODICITE INT.
    CREPLADE1       As Long                           ' DATE 1° INTERET
    CREPLADIN       As String * 2                     ' DATE REF. INTERET
    CREPLAJE1       As Long                           ' JOUR ECH. INTERET
    CREPLAINC       As String * 1                     ' INTERET CAPITALISE
    CREPLATAF       As Double                         ' TAUX DU PRET
    CREPLARTA       As String * 6                     ' REFERENCE DU TAUX
    CREPLAMAR       As Double                         ' MARGE
    CREPLATMI       As Double                         ' TAUX MINI
    CREPLATMA       As Double                         ' TAUX MAXI
    CREPLACTR       As String * 6                     ' CODE TAUX REVISION
    CREPLAAPL       As String * 1                     ' COD APPLICATION TAUX
    CREPLADPR       As Long                           ' DATE REVISION
    CREPLATVA       As String * 6                     ' CODE TVA
    CREPLATXT       As Double                         ' TAUX DE TVA
    CREPLATYR       As String * 1                     ' TYPE DE REPORT
    CREPLABAS       As Long                           ' NB DE JOURS ANNEE
    CREPLAREA       As String * 1                     ' JOUR REEL
    CREPLADUM       As Long                           ' DUREE MAXI PLAN
    CREPLATDU       As String * 1                     ' TYPE PERIO DUREE MAX
    CREPLACDR       As String * 6                     ' CODE REV ECH INT>ECH
    CREPLARES       As Currency                       ' MONTANT RESIDUEL
    CREPLADEJ       As String * 1                     ' PLAN DEJA CALCUL O/N
    CREPLANBJ       As Long                           ' DELAI D"USANCE
    CREPLASIG       As String * 1                     ' SENS DU DELAI
    CREPLATYJ       As String * 1                     ' TYPE DE JOURS
    CREPLAARR       As String * 1                     ' NBRE DE DECIMALES
    CREPLATYA       As String * 1                     ' TYPE D"ARRONDI
    CREPLACOT       As String * 3                     ' DEVISE DE COTATION

End Type
Public Sub rsZCREPLA0_Init(rsYCREPLA0 As typeZCREPLA0)
rsYCREPLA0.CREPLAETA = 0
rsYCREPLA0.CREPLAAGE = 0
rsYCREPLA0.CREPLASER = ""
rsYCREPLA0.CREPLASSE = ""
rsYCREPLA0.CREPLADOS = 0
rsYCREPLA0.CREPLAPRE = 0
rsYCREPLA0.CREPLAPLA = 0
rsYCREPLA0.CREPLAMAM = 0
rsYCREPLA0.CREPLAMIN = 0
rsYCREPLA0.CREPLAMOA = ""
rsYCREPLA0.CREPLANPC = 0
rsYCREPLA0.CREPLAPCA = ""
rsYCREPLA0.CREPLADEC = 0
rsYCREPLA0.CREPLADRE = ""
rsYCREPLA0.CREPLAJEC = 0
rsYCREPLA0.CREPLADTO = ""
rsYCREPLA0.CREPLADAM = ""
rsYCREPLA0.CREPLANPE = 0
rsYCREPLA0.CREPLAPIN = ""
rsYCREPLA0.CREPLAPEI = ""
rsYCREPLA0.CREPLADE1 = 0
rsYCREPLA0.CREPLADIN = ""
rsYCREPLA0.CREPLAJE1 = 0
rsYCREPLA0.CREPLAINC = ""
rsYCREPLA0.CREPLATAF = 0
rsYCREPLA0.CREPLARTA = ""
rsYCREPLA0.CREPLAMAR = 0
rsYCREPLA0.CREPLATMI = 0
rsYCREPLA0.CREPLATMA = 0
rsYCREPLA0.CREPLACTR = ""
rsYCREPLA0.CREPLAAPL = ""
rsYCREPLA0.CREPLADPR = 0
rsYCREPLA0.CREPLATVA = ""
rsYCREPLA0.CREPLATXT = 0
rsYCREPLA0.CREPLATYR = ""
rsYCREPLA0.CREPLABAS = 0
rsYCREPLA0.CREPLAREA = ""
rsYCREPLA0.CREPLADUM = 0
rsYCREPLA0.CREPLATDU = ""
rsYCREPLA0.CREPLACDR = ""
rsYCREPLA0.CREPLARES = 0
rsYCREPLA0.CREPLADEJ = ""
rsYCREPLA0.CREPLANBJ = 0
rsYCREPLA0.CREPLASIG = ""
rsYCREPLA0.CREPLATYJ = ""
rsYCREPLA0.CREPLAARR = ""
rsYCREPLA0.CREPLATYA = ""
rsYCREPLA0.CREPLACOT = ""
End Sub
Public Function rsZCREPLA0_GetBuffer(rsAdo As ADODB.Recordset, rsZCREPLA0 As typeZCREPLA0)
On Error GoTo Error_Handler
rsZCREPLA0_GetBuffer = Null
rsZCREPLA0.CREPLAETA = rsAdo("CREPLAETA")
rsZCREPLA0.CREPLAAGE = rsAdo("CREPLAAGE")
rsZCREPLA0.CREPLASER = rsAdo("CREPLASER")
rsZCREPLA0.CREPLASSE = rsAdo("CREPLASSE")
rsZCREPLA0.CREPLADOS = rsAdo("CREPLADOS")
rsZCREPLA0.CREPLAPRE = rsAdo("CREPLAPRE")
rsZCREPLA0.CREPLAPLA = rsAdo("CREPLAPLA")
rsZCREPLA0.CREPLAMAM = rsAdo("CREPLAMAM")
rsZCREPLA0.CREPLAMIN = rsAdo("CREPLAMIN")
rsZCREPLA0.CREPLAMOA = rsAdo("CREPLAMOA")
rsZCREPLA0.CREPLANPC = rsAdo("CREPLANPC")
rsZCREPLA0.CREPLAPCA = rsAdo("CREPLAPCA")
rsZCREPLA0.CREPLADEC = rsAdo("CREPLADEC")
rsZCREPLA0.CREPLADRE = rsAdo("CREPLADRE")
rsZCREPLA0.CREPLAJEC = rsAdo("CREPLAJEC")
rsZCREPLA0.CREPLADTO = rsAdo("CREPLADTO")
rsZCREPLA0.CREPLADAM = rsAdo("CREPLADAM")
rsZCREPLA0.CREPLANPE = rsAdo("CREPLANPE")
rsZCREPLA0.CREPLAPIN = rsAdo("CREPLAPIN")
rsZCREPLA0.CREPLAPEI = rsAdo("CREPLAPEI")
rsZCREPLA0.CREPLADE1 = rsAdo("CREPLADE1")
rsZCREPLA0.CREPLADIN = rsAdo("CREPLADIN")
rsZCREPLA0.CREPLAJE1 = rsAdo("CREPLAJE1")
rsZCREPLA0.CREPLAINC = rsAdo("CREPLAINC")
rsZCREPLA0.CREPLATAF = rsAdo("CREPLATAF")
rsZCREPLA0.CREPLARTA = rsAdo("CREPLARTA")
rsZCREPLA0.CREPLAMAR = rsAdo("CREPLAMAR")
rsZCREPLA0.CREPLATMI = rsAdo("CREPLATMI")
rsZCREPLA0.CREPLATMA = rsAdo("CREPLATMA")
rsZCREPLA0.CREPLACTR = rsAdo("CREPLACTR")
rsZCREPLA0.CREPLAAPL = rsAdo("CREPLAAPL")
rsZCREPLA0.CREPLADPR = rsAdo("CREPLADPR")
rsZCREPLA0.CREPLATVA = rsAdo("CREPLATVA")
rsZCREPLA0.CREPLATXT = rsAdo("CREPLATXT")
rsZCREPLA0.CREPLATYR = rsAdo("CREPLATYR")
rsZCREPLA0.CREPLABAS = rsAdo("CREPLABAS")
rsZCREPLA0.CREPLAREA = rsAdo("CREPLAREA")
rsZCREPLA0.CREPLADUM = rsAdo("CREPLADUM")
rsZCREPLA0.CREPLATDU = rsAdo("CREPLATDU")
rsZCREPLA0.CREPLACDR = rsAdo("CREPLACDR")
rsZCREPLA0.CREPLARES = rsAdo("CREPLARES")
rsZCREPLA0.CREPLADEJ = rsAdo("CREPLADEJ")
rsZCREPLA0.CREPLANBJ = rsAdo("CREPLANBJ")
rsZCREPLA0.CREPLASIG = rsAdo("CREPLASIG")
rsZCREPLA0.CREPLATYJ = rsAdo("CREPLATYJ")
rsZCREPLA0.CREPLAARR = rsAdo("CREPLAARR")
rsZCREPLA0.CREPLATYA = rsAdo("CREPLATYA")
rsZCREPLA0.CREPLACOT = rsAdo("CREPLACOT")
Exit Function
Error_Handler:
rsZCREPLA0_GetBuffer = Error
End Function

