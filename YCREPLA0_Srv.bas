Attribute VB_Name = "srvYCREPLA0"

'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCREPLA0Len = 500 ' 34 + ??????
Public Const recYCREPLA0_Block = 100 '????
Public Const constYCREPLA0 = "YCREPLA0"
Dim meYbase As typeYBase
Dim paramYCREPLA0_Import As String

Type typeYCREPLA0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
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
    
'---------------------------------------------------------
Public Function srvYCREPLA0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCREPLA0 As typeYCREPLA0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCREPLA0_GetBuffer_ODBC = Null

    
    recYCREPLA0.CREPLAETA = rsADO("CREPLAETA") 'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCREPLA0.CREPLAAGE = rsADO("CREPLAAGE") 'CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCREPLA0.CREPLASER = rsADO("CREPLASER") 'mId$(MsgTxt, K + 11, 2)
    recYCREPLA0.CREPLASSE = rsADO("CREPLASSE") 'mId$(MsgTxt, K + 13, 2)
    recYCREPLA0.CREPLADOS = rsADO("CREPLADOS") 'CLng(Val(mId$(MsgTxt, K + 15, 8)))
    recYCREPLA0.CREPLAPRE = rsADO("CREPLAPRE") 'CLng(Val(mId$(MsgTxt, K + 23, 4)))
    recYCREPLA0.CREPLAPLA = rsADO("CREPLAPLA") 'CLng(Val(mId$(MsgTxt, K + 27, 4)))
    recYCREPLA0.CREPLAMAM = rsADO("CREPLAMAM") 'CCur(Val(mId$(MsgTxt, K + 31, 16))) / 100
    recYCREPLA0.CREPLAMIN = rsADO("CREPLAMIN") 'CCur(Val(mId$(MsgTxt, K + 47, 16))) / 100
    recYCREPLA0.CREPLAMOA = rsADO("CREPLAMOA") 'mId$(MsgTxt, K + 63, 1)
    recYCREPLA0.CREPLANPC = rsADO("CREPLANPC") 'CLng(Val(mId$(MsgTxt, K + 64, 4)))
    recYCREPLA0.CREPLAPCA = rsADO("CREPLAPCA") 'mId$(MsgTxt, K + 68, 1)
    recYCREPLA0.CREPLADEC = rsADO("CREPLADEC") 'CLng(Val(mId$(MsgTxt, K + 69, 8)))
    recYCREPLA0.CREPLADRE = rsADO("CREPLADRE") 'mId$(MsgTxt, K + 77, 2)
    recYCREPLA0.CREPLAJEC = rsADO("CREPLAJEC") 'CLng(Val(mId$(MsgTxt, K + 79, 4)))
    recYCREPLA0.CREPLADTO = rsADO("CREPLADTO") 'mId$(MsgTxt, K + 83, 1)
    recYCREPLA0.CREPLADAM = rsADO("CREPLADAM") 'mId$(MsgTxt, K + 84, 1)
    recYCREPLA0.CREPLANPE = rsADO("CREPLANPE") 'CLng(Val(mId$(MsgTxt, K + 85, 4)))
    recYCREPLA0.CREPLAPIN = rsADO("CREPLAPIN") 'mId$(MsgTxt, K + 89, 1)
    recYCREPLA0.CREPLAPEI = rsADO("CREPLAPEI") 'mId$(MsgTxt, K + 90, 1)
    recYCREPLA0.CREPLADE1 = rsADO("CREPLADE1") 'CLng(Val(mId$(MsgTxt, K + 91, 8)))
    recYCREPLA0.CREPLADIN = rsADO("CREPLADIN") 'mId$(MsgTxt, K + 99, 2)
    recYCREPLA0.CREPLAJE1 = rsADO("CREPLAJE1") 'CLng(Val(mId$(MsgTxt, K + 101, 4)))
    recYCREPLA0.CREPLAINC = rsADO("CREPLAINC") 'mId$(MsgTxt, K + 105, 1)
    recYCREPLA0.CREPLATAF = rsADO("CREPLATAF") 'CDbl(Val(mId$(MsgTxt, K + 106, 13))) / 1000000000
    recYCREPLA0.CREPLARTA = rsADO("CREPLARTA") 'mId$(MsgTxt, K + 119, 6)
    recYCREPLA0.CREPLAMAR = rsADO("CREPLAMAR") 'CDbl(Val(mId$(MsgTxt, K + 125, 13))) / 1000000000
    recYCREPLA0.CREPLATMI = rsADO("CREPLATMI") 'CDbl(Val(mId$(MsgTxt, K + 138, 13))) / 1000000000
    recYCREPLA0.CREPLATMA = rsADO("CREPLATMA") 'CDbl(Val(mId$(MsgTxt, K + 151, 13))) / 1000000000
    recYCREPLA0.CREPLACTR = rsADO("CREPLACTR") 'mId$(MsgTxt, K + 164, 6)
    recYCREPLA0.CREPLAAPL = rsADO("CREPLAAPL") 'mId$(MsgTxt, K + 170, 1)
    recYCREPLA0.CREPLADPR = rsADO("CREPLADPR") 'CLng(Val(mId$(MsgTxt, K + 171, 8)))
    recYCREPLA0.CREPLATVA = rsADO("CREPLATVA") 'mId$(MsgTxt, K + 179, 6)
    recYCREPLA0.CREPLATXT = rsADO("CREPLATXT") 'CDbl(Val(mId$(MsgTxt, K + 185, 9))) / 10000000
    recYCREPLA0.CREPLATYR = rsADO("CREPLATYR") 'mId$(MsgTxt, K + 194, 1)
    recYCREPLA0.CREPLABAS = rsADO("CREPLABAS") 'CLng(Val(mId$(MsgTxt, K + 195, 2)))
    recYCREPLA0.CREPLAREA = rsADO("CREPLAREA") 'mId$(MsgTxt, K + 197, 1)
    recYCREPLA0.CREPLADUM = rsADO("CREPLADUM") 'CLng(Val(mId$(MsgTxt, K + 198, 4)))
    recYCREPLA0.CREPLATDU = rsADO("CREPLATDU") 'mId$(MsgTxt, K + 202, 1)
    recYCREPLA0.CREPLACDR = rsADO("CREPLACDR") 'mId$(MsgTxt, K + 203, 6)
    recYCREPLA0.CREPLARES = rsADO("CREPLARES") 'CCur(Val(mId$(MsgTxt, K + 209, 16))) / 100
    recYCREPLA0.CREPLADEJ = rsADO("CREPLADEJ") 'mId$(MsgTxt, K + 225, 1)
    recYCREPLA0.CREPLANBJ = rsADO("CREPLANBJ") 'CLng(Val(mId$(MsgTxt, K + 226, 4)))
    recYCREPLA0.CREPLASIG = rsADO("CREPLASIG") 'mId$(MsgTxt, K + 230, 1)
    recYCREPLA0.CREPLATYJ = rsADO("CREPLATYJ") 'mId$(MsgTxt, K + 231, 1)
    recYCREPLA0.CREPLAARR = rsADO("CREPLAARR") 'mId$(MsgTxt, K + 232, 1)
    recYCREPLA0.CREPLATYA = rsADO("CREPLATYA") 'mId$(MsgTxt, K + 233, 1)
    recYCREPLA0.CREPLACOT = rsADO("CREPLACOT") 'mId$(MsgTxt, K + 234, 3)
    
Exit Function

Error_Handler:
srvYCREPLA0_GetBuffer_ODBC = Error

End Function


Public Sub srvYCREPLA0_ElpDisplay(recYCREPLA0 As typeYCREPLA0)
frmElpDisplay.fgData.Rows = 49
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLASER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLASER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLASSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLASSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLADOS    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLADOS
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAPRE    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAPRE
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAPLA    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAPLA
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAMAM 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT AMORTI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAMAM
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAMIN 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT INTERETS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAMIN
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAMOA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE REMBOURSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAMOA
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLANPC    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NB PERIODES CAPITAL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLANPC
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAPCA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PERIODICITE CAPITAL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAPCA
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLADEC    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE 1° CAPITAL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLADEC
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLADRE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE REF. CAPITAL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLADRE
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAJEC    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "JOUR ECH. CAPITAL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAJEC
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLADTO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DIFFERE TOTAL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLADTO
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLADAM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DIFFERE D AMORTIS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLADAM
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLANPE    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NB PERIO. DIFFERE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLANPE
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAPIN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERETS SEPARES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAPIN
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAPEI    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PERIODICITE INT."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAPEI
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLADE1    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE 1° INTERET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLADE1
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLADIN    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE REF. INTERET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLADIN
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAJE1    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "JOUR ECH. INTERET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAJE1
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAINC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERET CAPITALISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAINC
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLATAF 12.9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX DU PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLATAF
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLARTA    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REFERENCE DU TAUX"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLARTA
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAMAR 12.9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MARGE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAMAR
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLATMI 12.9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX MINI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLATMI
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLATMA 12.9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX MAXI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLATMA
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLACTR    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE TAUX REVISION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLACTR
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAAPL    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COD APPLICATION TAUX"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAAPL
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLADPR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE REVISION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLADPR
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLATVA    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLATVA
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLATXT  8.7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX DE TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLATXT
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLATYR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE DE REPORT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLATYR
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLABAS    1S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NB DE JOURS ANNEE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLABAS
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAREA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "JOUR REEL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAREA
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLADUM    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DUREE MAXI PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLADUM
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLATDU    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE PERIO DUREE MAX"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLATDU
frmElpDisplay.fgData.Row = 40
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLACDR    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE REV ECH INT>ECH"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLACDR
frmElpDisplay.fgData.Row = 41
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLARES 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT RESIDUEL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLARES
frmElpDisplay.fgData.Row = 42
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLADEJ    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PLAN DEJA CALCUL O/N"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLADEJ
frmElpDisplay.fgData.Row = 43
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLANBJ    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DELAI D USANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLANBJ
frmElpDisplay.fgData.Row = 44
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLASIG    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SENS DU DELAI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLASIG
frmElpDisplay.fgData.Row = 45
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLATYJ    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE DE JOURS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLATYJ
frmElpDisplay.fgData.Row = 46
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLAARR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NBRE DE DECIMALES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLAARR
frmElpDisplay.fgData.Row = 47
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLATYA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE D ARRONDI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLATYA
frmElpDisplay.fgData.Row = 48
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPLACOT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE DE COTATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPLA0.CREPLACOT
frmElpDisplay.Show vbModal
End Sub
'---------------------------------------------------------
Public Sub recYCREPLA0_Init(recYCREPLA0 As typeYCREPLA0)
'---------------------------------------------------------

recYCREPLA0.obj = "ZCREPLA0_S"
recYCREPLA0.CREPLAETA = 0   '       As Integer                        ' ETABLISSEMENT
recYCREPLA0.CREPLAAGE = 0   '       As Integer                        ' AGENCE
recYCREPLA0.CREPLASER = ""  '       As String * 2                     ' SERVICE
recYCREPLA0.CREPLASSE = ""  '       As String * 2                     ' SOUS-SERVICE
recYCREPLA0.CREPLADOS = 0   '       As Long                           ' NUMERO DOSSIER
recYCREPLA0.CREPLAPRE = 0   '       As Long                           ' N° PRET
recYCREPLA0.CREPLAPLA = 0   '       As Long                           ' N° PLAN
recYCREPLA0.CREPLAMAM = 0   '       As Currency                       ' MONTANT AMORTI
recYCREPLA0.CREPLAMIN = 0   '       As Currency                       ' MONTANT INTERETS
recYCREPLA0.CREPLAMOA = ""  '       As String * 1                     ' TYPE REMBOURSEMENT
recYCREPLA0.CREPLANPC = 0   '       As Long                           ' NB PERIODES CAPITAL
recYCREPLA0.CREPLAPCA = ""  '       As String * 1                     ' PERIODICITE CAPITAL
recYCREPLA0.CREPLADEC = 0   '       As Long                           ' DATE 1° CAPITAL
recYCREPLA0.CREPLADRE = ""  '       As String * 2                     ' DATE REF. CAPITAL
recYCREPLA0.CREPLAJEC = 0   '       As Long                           ' JOUR ECH. CAPITAL
recYCREPLA0.CREPLADTO = ""  '       As String * 1                     ' DIFFERE TOTAL
recYCREPLA0.CREPLADAM = ""  '       As String * 1                     ' DIFFERE D AMORTIS
recYCREPLA0.CREPLANPE = 0   '       As Long                           ' NB PERIO. DIFFERE
recYCREPLA0.CREPLAPIN = ""  '       As String * 1                     ' INTERETS SEPARES
recYCREPLA0.CREPLAPEI = ""  '       As String * 1                     ' PERIODICITE INT.
recYCREPLA0.CREPLADE1 = 0   '       As Long                           ' DATE 1° INTERET
recYCREPLA0.CREPLADIN = ""  '       As String * 2                     ' DATE REF. INTERET
recYCREPLA0.CREPLAJE1 = 0   '       As Long                           ' JOUR ECH. INTERET
recYCREPLA0.CREPLAINC = ""  '       As String * 1                     ' INTERET CAPITALISE
recYCREPLA0.CREPLATAF = 0   '       As Double                         ' TAUX DU PRET
recYCREPLA0.CREPLARTA = ""  '       As String * 6                     ' REFERENCE DU TAUX
recYCREPLA0.CREPLAMAR = 0   '       As Double                         ' MARGE
recYCREPLA0.CREPLATMI = 0   '       As Double                         ' TAUX MINI
recYCREPLA0.CREPLATMA = 0   '       As Double                         ' TAUX MAXI
recYCREPLA0.CREPLACTR = ""  '       As String * 6                     ' CODE TAUX REVISION
recYCREPLA0.CREPLAAPL = ""  '       As String * 1                     ' COD APPLICATION TAUX
recYCREPLA0.CREPLADPR = 0   '       As Long                           ' DATE REVISION
recYCREPLA0.CREPLATVA = ""  '       As String * 6                     ' CODE TVA
recYCREPLA0.CREPLATXT = 0   '       As Double                         ' TAUX DE TVA
recYCREPLA0.CREPLATYR = ""  '       As String * 1                     ' TYPE DE REPORT
recYCREPLA0.CREPLABAS = 0   '       As Long                           ' NB DE JOURS ANNEE
recYCREPLA0.CREPLAREA = ""  '       As String * 1                     ' JOUR REEL
recYCREPLA0.CREPLADUM = 0   '       As Long                           ' DUREE MAXI PLAN
recYCREPLA0.CREPLATDU = ""  '       As String * 1                     ' TYPE PERIO DUREE MAX
recYCREPLA0.CREPLACDR = ""  '       As String * 6                     ' CODE REV ECH INT>ECH
recYCREPLA0.CREPLARES = 0   '       As Currency                       ' MONTANT RESIDUEL
recYCREPLA0.CREPLADEJ = ""  '       As String * 1                     ' PLAN DEJA CALCUL O/N
recYCREPLA0.CREPLANBJ = 0   '       As Long                           ' DELAI D"USANCE
recYCREPLA0.CREPLASIG = ""  '       As String * 1                     ' SENS DU DELAI
recYCREPLA0.CREPLATYJ = ""  '       As String * 1                     ' TYPE DE JOURS
recYCREPLA0.CREPLAARR = ""  '       As String * 1                     ' NBRE DE DECIMALES
recYCREPLA0.CREPLATYA = ""  '       As String * 1                     ' TYPE D"ARRONDI
recYCREPLA0.CREPLACOT = ""  '       As String * 3                     ' DEVISE DE COTATION


End Sub





