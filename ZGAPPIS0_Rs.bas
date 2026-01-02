Attribute VB_Name = "rsZGAPPIS0"
'---------------------------------------------------------
Option Explicit
Type typeZGAPPIS0
    GAPPISTAB As Integer     '  CODE ETAT
    GAPPISECH As Integer     ' NUMÉRO ÉCHÉANCIER
    GAPPISCLA As Integer     '  N° CLASSE ECHEANCIER
    GAPPISETA As Integer     '  CODE ÉTABLISSEMENT
    GAPPISAGE As Integer     '  CODE AGENCE
    GAPPISSER As String * 2  '  CODE SERVICE
    GAPPISSSE As String * 2  ' CODE SOUS-SERVICE
    GAPPISOPE As String * 3  '  CODE OPÉRATION
    GAPPISNAT As String * 3  ' CODE NATURE
    GAPPISNUO As Long        '  NUMÉRO OPÉRATION
    GAPPISDEV As String * 3  ' DEVISE
    GAPPISSEN As String * 1  ' SENS
    GAPPISDEC As Long        '  DATE ÉCHÉANCE
    GAPPISRUB As String * 10 ' RUBRIQUE COMPTABLE
    GAPPISTPR As String * 9  ' TYPE PRODUIT
    GAPPISCLI As String * 7  ' NUMÉRO CLIENT
    GAPPISMON As Currency    '  MONTANT DU FLUX
    GAPPISTTI As String * 1  ' TYPE DE TAUX INTERNE
    GAPPISTTE As String * 1  'TYPE DE TAUX EXTERNE
    GAPPISRTV As String * 6  'CODE TAUX
    GAPPISTAU As Double      ' VALEUR DU TAUX
    GAPPISSOL As Currency    ' SOLDE RUBRI COMPTABL
    GAPPISPOU As Double      ' POURCENTAGE
    GAPPISSIG As String * 13 ' SIGLE DU CLIENT
    GAPPISVIL As String * 12 ' VILLE
    
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZGAPPIS0_GetBuffer(rsAdo As ADODB.Recordset, rsZGAPPIS0 As typeZGAPPIS0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZGAPPIS0_GetBuffer = Null
rsZGAPPIS0.GAPPISTAB = rsAdo("GAPPISTAB")
rsZGAPPIS0.GAPPISECH = rsAdo("GAPPISECH")
rsZGAPPIS0.GAPPISCLA = rsAdo("GAPPISCLA")
rsZGAPPIS0.GAPPISETA = rsAdo("GAPPISETA")
rsZGAPPIS0.GAPPISAGE = rsAdo("GAPPISAGE")
rsZGAPPIS0.GAPPISSER = rsAdo("GAPPISSER")
rsZGAPPIS0.GAPPISSSE = rsAdo("GAPPISSSE")
rsZGAPPIS0.GAPPISOPE = rsAdo("GAPPISOPE")
rsZGAPPIS0.GAPPISNAT = rsAdo("GAPPISNAT")
rsZGAPPIS0.GAPPISNUO = rsAdo("GAPPISNUO")
rsZGAPPIS0.GAPPISDEV = rsAdo("GAPPISDEV")
rsZGAPPIS0.GAPPISSEN = rsAdo("GAPPISSEN")
rsZGAPPIS0.GAPPISDEC = rsAdo("GAPPISDEC")
rsZGAPPIS0.GAPPISRUB = rsAdo("GAPPISRUB")
rsZGAPPIS0.GAPPISTPR = rsAdo("GAPPISTPR")
rsZGAPPIS0.GAPPISCLI = rsAdo("GAPPISCLI")
rsZGAPPIS0.GAPPISMON = rsAdo("GAPPISMON")
rsZGAPPIS0.GAPPISTTI = rsAdo("GAPPISTTI")
rsZGAPPIS0.GAPPISTTE = rsAdo("GAPPISTTE")
rsZGAPPIS0.GAPPISRTV = rsAdo("GAPPISRTV")
rsZGAPPIS0.GAPPISTAU = rsAdo("GAPPISTAU")
rsZGAPPIS0.GAPPISSOL = rsAdo("GAPPISSOL")
rsZGAPPIS0.GAPPISPOU = rsAdo("GAPPISPOU")
rsZGAPPIS0.GAPPISSIG = rsAdo("GAPPISSIG")
rsZGAPPIS0.GAPPISVIL = rsAdo("GAPPISVIL")

Exit Function

Error_Handler:

rsZGAPPIS0_GetBuffer = Error

End Function


'







