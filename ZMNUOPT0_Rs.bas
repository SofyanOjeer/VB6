Attribute VB_Name = "rsZMNUOPT0"
'---------------------------------------------------------
Option Explicit
Type typeZMNUOPT0
    MNUOPTCOD       As Long                           ' CODE OPTION
    MNUOPTCLI       As String * 7                     ' CLIENT
    MNUOPTLIB       As String * 35                    ' LIBELLE
    MNUOPTENS       As String * 8                     ' ENSEMBLE
    MNUOPTENT       As String * 8                     ' POINT ENTREE
    MNUOPTSTR       As String * 1                     ' OPTION STRAB
    MNUOPTARE       As String * 1                     ' ARRET LOGICIEL
    MNUOPTBAT       As String * 1                     ' OPTION BATCH
    MNUOPTVAL       As String * 1                     ' VALID. BATCH
    MNUOPTSUP       As String * 1                     ' A SUPPRIMER
    MNUOPTOIA       As String * 1                     ' INTER-AGENCE
    MNUOPTGES       As String * 1                     ' SECUR.INTER-ETAB.
    
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZMNUOPT0_GetBuffer(rsAdo As ADODB.Recordset, rsZMNUOPT0 As typeZMNUOPT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZMNUOPT0_GetBuffer = Null

rsZMNUOPT0.MNUOPTCOD = rsAdo("MNUOPTCOD")
rsZMNUOPT0.MNUOPTCLI = rsAdo("MNUOPTCLI")
rsZMNUOPT0.MNUOPTLIB = rsAdo("MNUOPTLIB")
rsZMNUOPT0.MNUOPTENS = rsAdo("MNUOPTENS")
rsZMNUOPT0.MNUOPTENT = rsAdo("MNUOPTENT")
rsZMNUOPT0.MNUOPTSTR = rsAdo("MNUOPTSTR")
rsZMNUOPT0.MNUOPTARE = rsAdo("MNUOPTARE")
rsZMNUOPT0.MNUOPTBAT = rsAdo("MNUOPTBAT")
rsZMNUOPT0.MNUOPTVAL = rsAdo("MNUOPTVAL")
rsZMNUOPT0.MNUOPTSUP = rsAdo("MNUOPTSUP")
rsZMNUOPT0.MNUOPTOIA = rsAdo("MNUOPTOIA")
rsZMNUOPT0.MNUOPTGES = rsAdo("MNUOPTGES")

Exit Function

Error_Handler:

rsZMNUOPT0_GetBuffer = Error

End Function


'






