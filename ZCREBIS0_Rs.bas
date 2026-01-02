Attribute VB_Name = "rsZCREBIS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCREBIS0
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
Public Sub rsZCREBIS0_Init(rsYCREBIS0 As typeZCREBIS0)
rsYCREBIS0.CREBISETA = 0
rsYCREBIS0.CREBISAGE = 0
rsYCREBIS0.CREBISSER = ""
rsYCREBIS0.CREBISSSE = ""
rsYCREBIS0.CREBISDOS = 0
rsYCREBIS0.CREBISPRE = 0
rsYCREBIS0.CREBISTYP = ""
rsYCREBIS0.CREBISPAY = ""
rsYCREBIS0.CREBISMOD = ""
rsYCREBIS0.CREBISPLA = 0
rsYCREBIS0.CREBISCOM = ""
rsYCREBIS0.CREBISEMI = 0
rsYCREBIS0.CREBISREG = 0
rsYCREBIS0.CREBISDTR = 0
rsYCREBIS0.CREBISCPT = 0
rsYCREBIS0.CREBISAVI = 0
rsYCREBIS0.CREBISDEB = 0
rsYCREBIS0.CREBISFIN = 0
rsYCREBIS0.CREBISMAM = 0
rsYCREBIS0.CREBISMIN = 0
rsYCREBIS0.CREBISITC = 0
rsYCREBIS0.CREBISREP = 0
rsYCREBIS0.CREBISSEC = 0
rsYCREBIS0.CREBISCAS = ""
rsYCREBIS0.CREBISCOP = 0
rsYCREBIS0.CREBISTAU = 0
rsYCREBIS0.CREBISCOU = 0
rsYCREBIS0.CREBISBAS = ""
rsYCREBIS0.CREBISNUM = 0
rsYCREBIS0.CREBISMTT = 0
rsYCREBIS0.CREBISDRE = ""
rsYCREBIS0.CREBISMRE = 0
rsYCREBIS0.CREBISCOC = 0
rsYCREBIS0.CREBISASC = 0
rsYCREBIS0.CREBISNPL = 0
rsYCREBIS0.CREBISPAL = 0
rsYCREBIS0.CREBISECH = 0
End Sub
Public Function rsZCREBIS0_GetBuffer(rsAdo As ADODB.Recordset, rsZCREBIS0 As typeZCREBIS0)
On Error GoTo Error_Handler
rsZCREBIS0_GetBuffer = Null
rsZCREBIS0.CREBISETA = rsAdo("CREBISETA")
rsZCREBIS0.CREBISAGE = rsAdo("CREBISAGE")
rsZCREBIS0.CREBISSER = rsAdo("CREBISSER")
rsZCREBIS0.CREBISSSE = rsAdo("CREBISSSE")
rsZCREBIS0.CREBISDOS = rsAdo("CREBISDOS")
rsZCREBIS0.CREBISPRE = rsAdo("CREBISPRE")
rsZCREBIS0.CREBISTYP = rsAdo("CREBISTYP")
rsZCREBIS0.CREBISPAY = rsAdo("CREBISPAY")
rsZCREBIS0.CREBISMOD = rsAdo("CREBISMOD")
rsZCREBIS0.CREBISPLA = rsAdo("CREBISPLA")
rsZCREBIS0.CREBISCOM = rsAdo("CREBISCOM")
rsZCREBIS0.CREBISEMI = rsAdo("CREBISEMI")
rsZCREBIS0.CREBISREG = rsAdo("CREBISREG")
rsZCREBIS0.CREBISDTR = rsAdo("CREBISDTR")
rsZCREBIS0.CREBISCPT = rsAdo("CREBISCPT")
rsZCREBIS0.CREBISAVI = rsAdo("CREBISAVI")
rsZCREBIS0.CREBISDEB = rsAdo("CREBISDEB")
rsZCREBIS0.CREBISFIN = rsAdo("CREBISFIN")
rsZCREBIS0.CREBISMAM = rsAdo("CREBISMAM")
rsZCREBIS0.CREBISMIN = rsAdo("CREBISMIN")
rsZCREBIS0.CREBISITC = rsAdo("CREBISITC")
rsZCREBIS0.CREBISREP = rsAdo("CREBISREP")
rsZCREBIS0.CREBISSEC = rsAdo("CREBISSEC")
rsZCREBIS0.CREBISCAS = rsAdo("CREBISCAS")
rsZCREBIS0.CREBISCOP = rsAdo("CREBISCOP")
rsZCREBIS0.CREBISTAU = rsAdo("CREBISTAU")
rsZCREBIS0.CREBISCOU = rsAdo("CREBISCOU")
rsZCREBIS0.CREBISBAS = rsAdo("CREBISBAS")
rsZCREBIS0.CREBISNUM = rsAdo("CREBISNUM")
rsZCREBIS0.CREBISMTT = rsAdo("CREBISMTT")
rsZCREBIS0.CREBISDRE = rsAdo("CREBISDRE")
rsZCREBIS0.CREBISMRE = rsAdo("CREBISMRE")
rsZCREBIS0.CREBISCOC = rsAdo("CREBISCOC")
rsZCREBIS0.CREBISASC = rsAdo("CREBISASC")
rsZCREBIS0.CREBISNPL = rsAdo("CREBISNPL")
rsZCREBIS0.CREBISPAL = rsAdo("CREBISPAL")
rsZCREBIS0.CREBISECH = rsAdo("CREBISECH")
Exit Function
Error_Handler:
rsZCREBIS0_GetBuffer = Error
End Function

