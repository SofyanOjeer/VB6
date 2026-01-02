Attribute VB_Name = "srvYCREBIS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const constYCREBIS0 = "YCREBIS0"
Type typeYCREBIS0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    CREBISETA       As Integer                        ' ETABLISSEMENT
    CREBISAGE       As Integer                        ' AGENCE
    CREBISSER       As String * 2                     ' SERVICE
    CREBISSSE       As String * 2                     ' SOUS-SERVICE
    CREBISDOS       As Long                           ' N° DE DOSSIER
    CREBISPRE       As Long                           ' N° DE PRET
    CREBISTYP       As String * 2                     ' TYPE EVENEMENT
    CREBISPAY       As String * 7                     ' PAYEUR
    CREBISMOD       As String * 3                     ' MODE REGLEMENT
    CREBISPLA       As Long                           ' N° PLAN COMPTAB
    CREBISCOM       As String * 30                    ' COMPTE OU RIB
    CREBISEMI       As Long                           ' EMISSION PREVUE
    CREBISREG       As Long                           ' DATE EMISSION
    CREBISDTR       As Long                           ' DATE DU CALCUL
    CREBISCPT       As Long                           ' COMPTABILISATION
    CREBISAVI       As Long                           ' EDITION AVIS
    CREBISDEB       As Long                           ' DEBUT DE PERIODE
    CREBISFIN       As Long                           ' FIN DE PERIODE
    CREBISMAM       As Currency                       ' AMORTISSEMENT
    CREBISMIN       As Currency                       ' INTERETS
    CREBISITC       As Currency                       ' REPORTES +ITC
    CREBISREP       As Currency                       ' REPORTES N PAYES
    CREBISSEC       As Long                           ' SEQ COM OU ASSUR
    CREBISCAS       As String * 6                     ' COMMI. OU ASSUR.
    CREBISCOP       As Long                           ' SEQUENCE COPART
    CREBISTAU       As Double                         ' TAUX
    CREBISCOU       As Double                         ' COURS
    CREBISBAS       As String * 1                     ' BASE / RECEVOIR
    CREBISNUM       As Integer                        ' NUMERO ECHEANCE
    CREBISMTT       As Currency                       ' MONTANT DE TVA
    CREBISDRE       As String * 3                     ' DEVISE REGLEMENT
    CREBISMRE       As Currency                       ' MONTANT REGLEMENT
    CREBISCOC       As Currency                       ' MT COM CUMULABLE
    CREBISASC       As Currency                       ' MT ASS CUMULABLE
    CREBISNPL       As Long                           ' NUMERO PLAN
    CREBISPAL       As Long                           ' NUMERO PALIER
    CREBISECH       As Long                           ' NUMERO ECHEANCE

End Type
Public Sub srvYCREBIS0_Init(recYCREBIS0 As typeYCREBIS0)
recYCREBIS0.Obj = "YCREBIS0"
recYCREBIS0.Method = ""
recYCREBIS0.Err = ""
recYCREBIS0.CREBISETA = 0
recYCREBIS0.CREBISAGE = 0
recYCREBIS0.CREBISSER = ""
recYCREBIS0.CREBISSSE = ""
recYCREBIS0.CREBISDOS = 0
recYCREBIS0.CREBISPRE = 0
recYCREBIS0.CREBISTYP = ""
recYCREBIS0.CREBISPAY = ""
recYCREBIS0.CREBISMOD = ""
recYCREBIS0.CREBISPLA = 0
recYCREBIS0.CREBISCOM = ""
recYCREBIS0.CREBISEMI = 0
recYCREBIS0.CREBISREG = 0
recYCREBIS0.CREBISDTR = 0
recYCREBIS0.CREBISCPT = 0
recYCREBIS0.CREBISAVI = 0
recYCREBIS0.CREBISDEB = 0
recYCREBIS0.CREBISFIN = 0
recYCREBIS0.CREBISMAM = 0
recYCREBIS0.CREBISMIN = 0
recYCREBIS0.CREBISITC = 0
recYCREBIS0.CREBISREP = 0
recYCREBIS0.CREBISSEC = 0
recYCREBIS0.CREBISCAS = ""
recYCREBIS0.CREBISCOP = 0
recYCREBIS0.CREBISTAU = 0
recYCREBIS0.CREBISCOU = 0
recYCREBIS0.CREBISBAS = ""
recYCREBIS0.CREBISNUM = 0
recYCREBIS0.CREBISMTT = 0
recYCREBIS0.CREBISDRE = ""
recYCREBIS0.CREBISMRE = 0
recYCREBIS0.CREBISCOC = 0
recYCREBIS0.CREBISASC = 0
recYCREBIS0.CREBISNPL = 0
recYCREBIS0.CREBISPAL = 0
recYCREBIS0.CREBISECH = 0
End Sub
Public Function srvYCREBIS0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCREBIS0 As typeYCREBIS0)
On Error GoTo Error_Handler
srvYCREBIS0_GetBuffer_ODBC = Null
recYCREBIS0.CREBISETA = rsADO("CREBISETA")
recYCREBIS0.CREBISAGE = rsADO("CREBISAGE")
recYCREBIS0.CREBISSER = rsADO("CREBISSER")
recYCREBIS0.CREBISSSE = rsADO("CREBISSSE")
recYCREBIS0.CREBISDOS = rsADO("CREBISDOS")
recYCREBIS0.CREBISPRE = rsADO("CREBISPRE")
recYCREBIS0.CREBISTYP = rsADO("CREBISTYP")
recYCREBIS0.CREBISPAY = rsADO("CREBISPAY")
recYCREBIS0.CREBISMOD = rsADO("CREBISMOD")
recYCREBIS0.CREBISPLA = rsADO("CREBISPLA")
recYCREBIS0.CREBISCOM = rsADO("CREBISCOM")
recYCREBIS0.CREBISEMI = rsADO("CREBISEMI")
recYCREBIS0.CREBISREG = rsADO("CREBISREG")
recYCREBIS0.CREBISDTR = rsADO("CREBISDTR")
recYCREBIS0.CREBISCPT = rsADO("CREBISCPT")
recYCREBIS0.CREBISAVI = rsADO("CREBISAVI")
recYCREBIS0.CREBISDEB = rsADO("CREBISDEB")
recYCREBIS0.CREBISFIN = rsADO("CREBISFIN")
recYCREBIS0.CREBISMAM = rsADO("CREBISMAM")
recYCREBIS0.CREBISMIN = rsADO("CREBISMIN")
recYCREBIS0.CREBISITC = rsADO("CREBISITC")
recYCREBIS0.CREBISREP = rsADO("CREBISREP")
recYCREBIS0.CREBISSEC = rsADO("CREBISSEC")
recYCREBIS0.CREBISCAS = rsADO("CREBISCAS")
recYCREBIS0.CREBISCOP = rsADO("CREBISCOP")
recYCREBIS0.CREBISTAU = rsADO("CREBISTAU")
recYCREBIS0.CREBISCOU = rsADO("CREBISCOU")
recYCREBIS0.CREBISBAS = rsADO("CREBISBAS")
recYCREBIS0.CREBISNUM = rsADO("CREBISNUM")
recYCREBIS0.CREBISMTT = rsADO("CREBISMTT")
recYCREBIS0.CREBISDRE = rsADO("CREBISDRE")
recYCREBIS0.CREBISMRE = rsADO("CREBISMRE")
recYCREBIS0.CREBISCOC = rsADO("CREBISCOC")
recYCREBIS0.CREBISASC = rsADO("CREBISASC")
recYCREBIS0.CREBISNPL = rsADO("CREBISNPL")
recYCREBIS0.CREBISPAL = rsADO("CREBISPAL")
recYCREBIS0.CREBISECH = rsADO("CREBISECH")
Exit Function
Error_Handler:
srvYCREBIS0_GetBuffer_ODBC = Error
End Function
Public Sub srvYCREBIS0_ElpDisplay(recYCREBIS0 As typeYCREBIS0)
frmElpDisplay.fgData.Rows = 38
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISSER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISSER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISSSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISSSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISDOS    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° DE DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISDOS
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISPRE    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° DE PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISPRE
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISTYP    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE EVENEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISTYP
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISPAY    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PAYEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISPAY
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISMOD    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MODE REGLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISMOD
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISPLA    1P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "N° PLAN COMPTAB"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISPLA
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISCOM   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPTE OU RIB"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISCOM
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISEMI    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EMISSION PREVUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISEMI
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISREG    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE EMISSION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISREG
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISDTR    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DU CALCUL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISDTR
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISCPT    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPTABILISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISCPT
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISAVI    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EDITION AVIS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISAVI
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISDEB    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEBUT DE PERIODE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISDEB
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISFIN    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "FIN DE PERIODE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISFIN
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISMAM 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AMORTISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISMAM
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISMIN 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERETS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISMIN
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISITC 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REPORTES +ITC"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISITC
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISREP 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REPORTES N PAYES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISREP
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISSEC    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SEQ COM OU ASSUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISSEC
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISCAS    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMMI. OU ASSUR."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISCAS
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISCOP    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SEQUENCE COPART"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISCOP
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISTAU 12.9P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISTAU
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISCOU15.10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COURS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISCOU
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISBAS    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BASE / RECEVOIR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISBAS
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISNUM    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO ECHEANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISNUM
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISMTT 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT DE TVA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISMTT
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISDRE    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE REGLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISDRE
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISMRE 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT REGLEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISMRE
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISCOC 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT COM CUMULABLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISCOC
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISASC 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MT ASS CUMULABLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISASC
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISNPL    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISNPL
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISPAL    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PALIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISPAL
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREBISECH    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO ECHEANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREBIS0.CREBISECH
frmElpDisplay.Show vbModal
End Sub
