Attribute VB_Name = "srvYCREEVE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCREEVE0Len = 500 ' 34 + ??????
Public Const recYCREEVE0_Block = 100 '????
Public Const constYCREEVE0 = "YCREEVE0"
Dim meYbase As typeYBase
Dim paramYCREEVE0_Import As String

Type typeYCREEVE0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    CREEVEETA       As Integer                        ' ETABLISSEMENT
    CREEVEAGE       As Integer                        ' AGENCE
    CREEVESER       As String * 2                     ' SERVICE
    CREEVESSE       As String * 2                     ' SOUS-SERVICE
    CREEVEDOS       As Long                           ' N° DE DOSSIER
    CREEVEPRE       As Long                           ' N° DE PRET
    CREEVETYP       As String * 2                     ' TYPE EVENEMENT
    CREEVEPAY       As String * 7                     ' PAYEUR
    CREEVEMOD       As String * 3                     ' MODE REGLEMENT
    CREEVEPLA       As Long                           ' N° PLAN COMPTAB
    CREEVECOM       As String * 30                    ' COMPTE OU RIB
    CREEVEEMI       As Long                           ' EMISSION PREVUE
    CREEVEREG       As Long                           ' DATE EMISSION
    CREEVEDTR       As Long                           ' DATE DU CALCUL
    CREEVECPT       As Long                           ' COMPTABILISATION
    CREEVEAVI       As Long                           ' EDITION AVIS
    CREEVEDEB       As Long                           ' DEBUT DE PERIODE
    CREEVEFIN       As Long                           ' FIN DE PERIODE
    CREEVEMAM       As Currency                       ' AMORTISSEMENT
    CREEVEMIN       As Currency                       ' INTERETS
    CREEVEITC       As Currency                       ' REPORTES +ITC
    CREEVEREP       As Currency                       ' REPORTES N PAYES
    CREEVESEC       As Long                           ' SEQ COM OU ASSUR
    CREEVECAS       As String * 6                     ' COMMI. OU ASSUR.
    CREEVECOP       As Long                           ' SEQUENCE COPART
    CREEVETAU       As Double                         ' TAUX
    CREEVECOU       As Double                         ' COURS
    CREEVEBAS       As String * 1                     ' BASE / RECEVOIR
    CREEVENUM       As Integer                        ' NUMERO ECHEANCE
    CREEVEMTT       As Currency                       ' MONTANT DE TVA
    CREEVEDRE       As String * 3                     ' DEVISE REGLEMENT
    CREEVEMRE       As Currency                       ' MONTANT REGLEMENT
    CREEVECOC       As Currency                       ' MT COM CUMULABLE
    CREEVEASC       As Currency                       ' MT ASS CUMULABLE
    CREEVENPL       As Long                           ' NUMERO PLAN
    CREEVEPAL       As Long                           ' NUMERO PALIER
    CREEVEECH       As Long                           ' NUMERO ECHEANCE

End Type
    
'---------------------------------------------------------
Public Function srvYCREEVE0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCREEVE0 As typeYCREEVE0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCREEVE0_GetBuffer_ODBC = Null

    recYCREEVE0.CREEVEETA = rsADO("CREEVEETA") 'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCREEVE0.CREEVEAGE = rsADO("CREEVEAGE") 'CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCREEVE0.CREEVESER = rsADO("CREEVESER") 'mId$(MsgTxt, K + 11, 2)
    recYCREEVE0.CREEVESSE = rsADO("CREEVESSE") 'mId$(MsgTxt, K + 13, 2)
    recYCREEVE0.CREEVEDOS = rsADO("CREEVEDOS") 'CLng(Val(mId$(MsgTxt, K + 15, 8)))
    recYCREEVE0.CREEVEPRE = rsADO("CREEVEPRE") 'CLng(Val(mId$(MsgTxt, K + 23, 4)))
    recYCREEVE0.CREEVETYP = rsADO("CREEVETYP") 'mId$(MsgTxt, K + 27, 2)
    recYCREEVE0.CREEVEPAY = rsADO("CREEVEPAY") 'mId$(MsgTxt, K + 29, 7)
    recYCREEVE0.CREEVEMOD = rsADO("CREEVEMOD") 'mId$(MsgTxt, K + 36, 3)
    recYCREEVE0.CREEVEPLA = rsADO("CREEVEPLA") 'CLng(Val(mId$(MsgTxt, K + 39, 2)))
    recYCREEVE0.CREEVECOM = rsADO("CREEVECOM") 'mId$(MsgTxt, K + 41, 30)
    recYCREEVE0.CREEVEEMI = rsADO("CREEVEEMI") 'CLng(Val(mId$(MsgTxt, K + 71, 8)))
    recYCREEVE0.CREEVEREG = rsADO("CREEVEREG") 'CLng(Val(mId$(MsgTxt, K + 79, 8)))
    recYCREEVE0.CREEVEDTR = rsADO("CREEVEDTR") 'CLng(Val(mId$(MsgTxt, K + 87, 8)))
    recYCREEVE0.CREEVECPT = rsADO("CREEVECPT") 'CLng(Val(mId$(MsgTxt, K + 95, 8)))
    recYCREEVE0.CREEVEAVI = rsADO("CREEVEAVI") 'CLng(Val(mId$(MsgTxt, K + 103, 8)))
    recYCREEVE0.CREEVEDEB = rsADO("CREEVEDEB") 'CLng(Val(mId$(MsgTxt, K + 111, 8)))
    recYCREEVE0.CREEVEFIN = rsADO("CREEVEFIN") 'CLng(Val(mId$(MsgTxt, K + 119, 8)))
    recYCREEVE0.CREEVEMAM = rsADO("CREEVEMAM") 'CCur(Val(mId$(MsgTxt, K + 127, 16))) / 100
    recYCREEVE0.CREEVEMIN = rsADO("CREEVEMIN") 'CCur(Val(mId$(MsgTxt, K + 143, 16))) / 100
    recYCREEVE0.CREEVEITC = rsADO("CREEVEITC") 'CCur(Val(mId$(MsgTxt, K + 159, 16))) / 100
    recYCREEVE0.CREEVEREP = rsADO("CREEVEREP") 'CCur(Val(mId$(MsgTxt, K + 175, 16))) / 100
    recYCREEVE0.CREEVESEC = rsADO("CREEVESEC") 'CLng(Val(mId$(MsgTxt, K + 191, 4)))
    recYCREEVE0.CREEVECAS = rsADO("CREEVECAS") 'mId$(MsgTxt, K + 195, 6)
    recYCREEVE0.CREEVECOP = rsADO("CREEVECOP") 'CLng(Val(mId$(MsgTxt, K + 201, 4)))
    recYCREEVE0.CREEVETAU = rsADO("CREEVETAU") 'CDbl(Val(mId$(MsgTxt, K + 205, 13))) / 1000000000
    recYCREEVE0.CREEVECOU = rsADO("CREEVECOU") 'CDbl(Val(mId$(MsgTxt, K + 218, 16))) / 10000000000#
    recYCREEVE0.CREEVEBAS = rsADO("CREEVEBAS") 'mId$(MsgTxt, K + 234, 1)
    recYCREEVE0.CREEVENUM = rsADO("CREEVENUM") 'CInt(Val(mId$(MsgTxt, K + 235, 5)))
    recYCREEVE0.CREEVEMTT = rsADO("CREEVEMTT") 'CCur(Val(mId$(MsgTxt, K + 240, 16))) / 100
    recYCREEVE0.CREEVEDRE = rsADO("CREEVEDRE") 'mId$(MsgTxt, K + 256, 3)
    recYCREEVE0.CREEVEMRE = rsADO("CREEVEMRE") 'CCur(Val(mId$(MsgTxt, K + 259, 16))) / 100
    recYCREEVE0.CREEVECOC = rsADO("CREEVECOC") 'CCur(Val(mId$(MsgTxt, K + 275, 16))) / 100
    recYCREEVE0.CREEVEASC = rsADO("CREEVEASC") 'CCur(Val(mId$(MsgTxt, K + 291, 16))) / 100
    recYCREEVE0.CREEVENPL = rsADO("CREEVENPL") 'CLng(Val(mId$(MsgTxt, K + 307, 4)))
    recYCREEVE0.CREEVEPAL = rsADO("CREEVEPAL") 'CLng(Val(mId$(MsgTxt, K + 311, 4)))
    recYCREEVE0.CREEVEECH = rsADO("CREEVEECH") 'CLng(Val(mId$(MsgTxt, K + 315, 4)))
    
Exit Function

Error_Handler:
srvYCREEVE0_GetBuffer_ODBC = Error

End Function


Public Sub srvYCREEVE0_ElpDisplay(recYCREEVE0 As typeYCREEVE0)
frmElpDisplay.fgData.Rows = 38
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVESER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVESER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVESSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVESSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEDOS    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° DE DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEDOS
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEPRE    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° DE PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEPRE
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVETYP    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE EVENEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVETYP
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEPAY    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PAYEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEPAY
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEMOD    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MODE REGLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEMOD
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEPLA    1P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° PLAN COMPTAB"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEPLA
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVECOM   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPTE OU RIB"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVECOM
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEEMI    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EMISSION PREVUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEEMI
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEREG    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE EMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEREG
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEDTR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DU CALCUL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEDTR
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVECPT    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPTABILISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVECPT
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEAVI    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EDITION AVIS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEAVI
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEDEB    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEBUT DE PERIODE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEDEB
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEFIN    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "FIN DE PERIODE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEFIN
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEMAM 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AMORTISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEMAM
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEMIN 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERETS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEMIN
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEITC 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REPORTES +ITC"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEITC
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEREP 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REPORTES N PAYES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEREP
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVESEC    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SEQ COM OU ASSUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVESEC
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVECAS    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMMI. OU ASSUR."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVECAS
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVECOP    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SEQUENCE COPART"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVECOP
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVETAU 12.9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVETAU
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVECOU15.10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COURS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVECOU
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEBAS    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BASE / RECEVOIR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEBAS
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVENUM    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO ECHEANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVENUM
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEMTT 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DE TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEMTT
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEDRE    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE REGLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEDRE
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEMRE 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT REGLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEMRE
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVECOC 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT COM CUMULABLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVECOC
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEASC 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT ASS CUMULABLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEASC
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVENPL    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVENPL
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEPAL    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PALIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEPAL
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREEVEECH    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO ECHEANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREEVE0.CREEVEECH
frmElpDisplay.Show vbModal
End Sub
'---------------------------------------------------------
Public Sub recYCREEVE0_Init(recYCREEVE0 As typeYCREEVE0)
'---------------------------------------------------------
recYCREEVE0.obj = "ZCREEVE0_S"

recYCREEVE0.CREEVEETA = 0   '       As Integer                        ' ETABLISSEMENT
recYCREEVE0.CREEVEAGE = 0   '       As Integer                        ' AGENCE
recYCREEVE0.CREEVESER = ""  '       As String * 2                     ' SERVICE
recYCREEVE0.CREEVESSE = ""  '       As String * 2                     ' SOUS-SERVICE
recYCREEVE0.CREEVEDOS = 0   '       As Long                           ' N° DE DOSSIER
recYCREEVE0.CREEVEPRE = 0   '       As Long                           ' N° DE PRET
recYCREEVE0.CREEVETYP = ""  '       As String * 2                     ' TYPE EVENEMENT
recYCREEVE0.CREEVEPAY = ""  '       As String * 7                     ' PAYEUR
recYCREEVE0.CREEVEMOD = ""  '       As String * 3                     ' MODE REGLEMENT
recYCREEVE0.CREEVEPLA = 0   '       As Long                           ' N° PLAN COMPTAB
recYCREEVE0.CREEVECOM = ""  '       As String * 30                    ' COMPTE OU RIB
recYCREEVE0.CREEVEEMI = 0   '       As Long                           ' EMISSION PREVUE
recYCREEVE0.CREEVEREG = 0   '       As Long                           ' DATE EMISSION
recYCREEVE0.CREEVEDTR = 0   '       As Long                           ' DATE DU CALCUL
recYCREEVE0.CREEVECPT = 0   '       As Long                           ' COMPTABILISATION
recYCREEVE0.CREEVEAVI = 0   '       As Long                           ' EDITION AVIS
recYCREEVE0.CREEVEDEB = 0   '       As Long                           ' DEBUT DE PERIODE
recYCREEVE0.CREEVEFIN = 0   '       As Long                           ' FIN DE PERIODE
recYCREEVE0.CREEVEMAM = 0   '       As Currency                       ' AMORTISSEMENT
recYCREEVE0.CREEVEMIN = 0   '       As Currency                       ' INTERETS
recYCREEVE0.CREEVEITC = 0   '       As Currency                       ' REPORTES +ITC
recYCREEVE0.CREEVEREP = 0   '       As Currency                       ' REPORTES N PAYES
recYCREEVE0.CREEVESEC = 0   '       As Long                           ' SEQ COM OU ASSUR
recYCREEVE0.CREEVECAS = ""  '       As String * 6                     ' COMMI. OU ASSUR.
recYCREEVE0.CREEVECOP = 0   '       As Long                           ' SEQUENCE COPART
recYCREEVE0.CREEVETAU = 0   '       As Double                         ' TAUX
recYCREEVE0.CREEVECOU = 0   '       As Double                         ' COURS
recYCREEVE0.CREEVEBAS = ""  '       As String * 1                     ' BASE / RECEVOIR
recYCREEVE0.CREEVENUM = 0   '       As Integer                        ' NUMERO ECHEANCE
recYCREEVE0.CREEVEMTT = 0   '       As Currency                       ' MONTANT DE TVA
recYCREEVE0.CREEVEDRE = ""  '       As String * 3                     ' DEVISE REGLEMENT
recYCREEVE0.CREEVEMRE = 0   '       As Currency                       ' MONTANT REGLEMENT
recYCREEVE0.CREEVECOC = 0   '       As Currency                       ' MT COM CUMULABLE
recYCREEVE0.CREEVEASC = 0   '       As Currency                       ' MT ASS CUMULABLE
recYCREEVE0.CREEVENPL = 0   '       As Long                           ' NUMERO PLAN
recYCREEVE0.CREEVEPAL = 0   '       As Long                           ' NUMERO PALIER
recYCREEVE0.CREEVEECH = 0   '       As Long                           ' NUMERO ECHEANCE

End Sub







