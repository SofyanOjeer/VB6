Attribute VB_Name = "rsZCREEVE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCREEVE0
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
Public Sub rsZCREEVE0_Init(rsYCREEVE0 As typeZCREEVE0)
rsYCREEVE0.CREEVEETA = 0
rsYCREEVE0.CREEVEAGE = 0
rsYCREEVE0.CREEVESER = ""
rsYCREEVE0.CREEVESSE = ""
rsYCREEVE0.CREEVEDOS = 0
rsYCREEVE0.CREEVEPRE = 0
rsYCREEVE0.CREEVETYP = ""
rsYCREEVE0.CREEVEPAY = ""
rsYCREEVE0.CREEVEMOD = ""
rsYCREEVE0.CREEVEPLA = 0
rsYCREEVE0.CREEVECOM = ""
rsYCREEVE0.CREEVEEMI = 0
rsYCREEVE0.CREEVEREG = 0
rsYCREEVE0.CREEVEDTR = 0
rsYCREEVE0.CREEVECPT = 0
rsYCREEVE0.CREEVEAVI = 0
rsYCREEVE0.CREEVEDEB = 0
rsYCREEVE0.CREEVEFIN = 0
rsYCREEVE0.CREEVEMAM = 0
rsYCREEVE0.CREEVEMIN = 0
rsYCREEVE0.CREEVEITC = 0
rsYCREEVE0.CREEVEREP = 0
rsYCREEVE0.CREEVESEC = 0
rsYCREEVE0.CREEVECAS = ""
rsYCREEVE0.CREEVECOP = 0
rsYCREEVE0.CREEVETAU = 0
rsYCREEVE0.CREEVECOU = 0
rsYCREEVE0.CREEVEBAS = ""
rsYCREEVE0.CREEVENUM = 0
rsYCREEVE0.CREEVEMTT = 0
rsYCREEVE0.CREEVEDRE = ""
rsYCREEVE0.CREEVEMRE = 0
rsYCREEVE0.CREEVECOC = 0
rsYCREEVE0.CREEVEASC = 0
rsYCREEVE0.CREEVENPL = 0
rsYCREEVE0.CREEVEPAL = 0
rsYCREEVE0.CREEVEECH = 0
End Sub
Public Function rsZCREEVE0_GetBuffer(rsAdo As ADODB.Recordset, rsZCREEVE0 As typeZCREEVE0)
On Error GoTo Error_Handler
rsZCREEVE0_GetBuffer = Null
rsZCREEVE0.CREEVEETA = rsAdo("CREEVEETA")
rsZCREEVE0.CREEVEAGE = rsAdo("CREEVEAGE")
rsZCREEVE0.CREEVESER = rsAdo("CREEVESER")
rsZCREEVE0.CREEVESSE = rsAdo("CREEVESSE")
rsZCREEVE0.CREEVEDOS = rsAdo("CREEVEDOS")
rsZCREEVE0.CREEVEPRE = rsAdo("CREEVEPRE")
rsZCREEVE0.CREEVETYP = rsAdo("CREEVETYP")
rsZCREEVE0.CREEVEPAY = rsAdo("CREEVEPAY")
rsZCREEVE0.CREEVEMOD = rsAdo("CREEVEMOD")
rsZCREEVE0.CREEVEPLA = rsAdo("CREEVEPLA")
rsZCREEVE0.CREEVECOM = rsAdo("CREEVECOM")
rsZCREEVE0.CREEVEEMI = rsAdo("CREEVEEMI")
rsZCREEVE0.CREEVEREG = rsAdo("CREEVEREG")
rsZCREEVE0.CREEVEDTR = rsAdo("CREEVEDTR")
rsZCREEVE0.CREEVECPT = rsAdo("CREEVECPT")
rsZCREEVE0.CREEVEAVI = rsAdo("CREEVEAVI")
rsZCREEVE0.CREEVEDEB = rsAdo("CREEVEDEB")
rsZCREEVE0.CREEVEFIN = rsAdo("CREEVEFIN")
rsZCREEVE0.CREEVEMAM = rsAdo("CREEVEMAM")
rsZCREEVE0.CREEVEMIN = rsAdo("CREEVEMIN")
rsZCREEVE0.CREEVEITC = rsAdo("CREEVEITC")
rsZCREEVE0.CREEVEREP = rsAdo("CREEVEREP")
rsZCREEVE0.CREEVESEC = rsAdo("CREEVESEC")
rsZCREEVE0.CREEVECAS = rsAdo("CREEVECAS")
rsZCREEVE0.CREEVECOP = rsAdo("CREEVECOP")
rsZCREEVE0.CREEVETAU = rsAdo("CREEVETAU")
rsZCREEVE0.CREEVECOU = rsAdo("CREEVECOU")
rsZCREEVE0.CREEVEBAS = rsAdo("CREEVEBAS")
rsZCREEVE0.CREEVENUM = rsAdo("CREEVENUM")
rsZCREEVE0.CREEVEMTT = rsAdo("CREEVEMTT")
rsZCREEVE0.CREEVEDRE = rsAdo("CREEVEDRE")
rsZCREEVE0.CREEVEMRE = rsAdo("CREEVEMRE")
rsZCREEVE0.CREEVECOC = rsAdo("CREEVECOC")
rsZCREEVE0.CREEVEASC = rsAdo("CREEVEASC")
rsZCREEVE0.CREEVENPL = rsAdo("CREEVENPL")
rsZCREEVE0.CREEVEPAL = rsAdo("CREEVEPAL")
rsZCREEVE0.CREEVEECH = rsAdo("CREEVEECH")
Exit Function
Error_Handler:
rsZCREEVE0_GetBuffer = Error
End Function

