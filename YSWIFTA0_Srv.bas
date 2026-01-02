Attribute VB_Name = "srvYSWIFTA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const constYSWIFTA0 = "YSWIFTA0"
Type typeYSWIFTA0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    SWIFTAETA       As Integer                        ' ETABLISSEMENT
    SWIFTAREF       As String * 16                    ' REFERNECE
    SWIFTANEN       As String * 1                     ' NUMERO DE RENVOI
    SWIFTAPRI       As String * 2                     ' CODE PROIRITE
    SWIFTAMES       As String * 3                     ' TYPE MESSAGE
    SWIFTADOR       As String * 12                    ' DONNEUR ORDRE
    SWIFTADES       As String * 12                    ' DESTINATAIRE
    SWIFTADVA       As Long                           ' DATE VALEUR
    SWIFTADE1       As String * 3                     ' DEVISE 1
    SWIFTAMON       As Currency                       ' DATE VALEUR
    SWIFTADE2       As String * 3                     ' DEVISE 2
    SWIFTADEN       As Long                           ' DATE ENVOI
    SWIFTAHEN       As Long                           ' HEURE ENVOI
    SWIFTACOM       As String * 1                     ' COMPLET
    SWIFTATES       As String * 1                     ' TEST OU REEL
    SWIFTASUP       As String * 1                     ' SUPPRIME
    SWIFTAVAL       As String * 1                     ' TOP VALIDATION
    SWIFTAAGE       As Integer                        ' AGENCE
    SWIFTASER       As String * 2                     ' SERVICE
    SWIFTASSE       As String * 2                     ' SOUS SERVICE
    SWIFTAUTI       As String * 10                    ' UTILISATEUR
    SWIFTANUM       As Long                           ' NUMERO INTERNE
    SWIFTAUT1       As String * 10                    ' UTILISA SAISIE
    SWIFTAPVA       As String * 1                     ' 1ERE VALIDATION
    SWIFTAUT2       As String * 10                    ' UTILISA 1ER VALID

End Type
Public Sub srvYSWIFTA0_Init(recYSWIFTA0 As typeYSWIFTA0)
recYSWIFTA0.Obj = "YSWIFTA0"
recYSWIFTA0.Method = ""
recYSWIFTA0.Err = ""
recYSWIFTA0.SWIFTAETA = 0
recYSWIFTA0.SWIFTAREF = ""
recYSWIFTA0.SWIFTANEN = ""
recYSWIFTA0.SWIFTAPRI = ""
recYSWIFTA0.SWIFTAMES = ""
recYSWIFTA0.SWIFTADOR = ""
recYSWIFTA0.SWIFTADES = ""
recYSWIFTA0.SWIFTADVA = 0
recYSWIFTA0.SWIFTADE1 = ""
recYSWIFTA0.SWIFTAMON = 0
recYSWIFTA0.SWIFTADE2 = ""
recYSWIFTA0.SWIFTADEN = 0
recYSWIFTA0.SWIFTAHEN = 0
recYSWIFTA0.SWIFTACOM = ""
recYSWIFTA0.SWIFTATES = ""
recYSWIFTA0.SWIFTASUP = ""
recYSWIFTA0.SWIFTAVAL = ""
recYSWIFTA0.SWIFTAAGE = 0
recYSWIFTA0.SWIFTASER = ""
recYSWIFTA0.SWIFTASSE = ""
recYSWIFTA0.SWIFTAUTI = ""
recYSWIFTA0.SWIFTANUM = 0
recYSWIFTA0.SWIFTAUT1 = ""
recYSWIFTA0.SWIFTAPVA = ""
recYSWIFTA0.SWIFTAUT2 = ""
End Sub
Public Function srvYSWIFTA0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYSWIFTA0 As typeYSWIFTA0)
On Error GoTo Error_Handler
srvYSWIFTA0_GetBuffer_ODBC = Null
recYSWIFTA0.SWIFTAETA = rsADO("SWIFTAETA")
recYSWIFTA0.SWIFTAREF = rsADO("SWIFTAREF")
recYSWIFTA0.SWIFTANEN = rsADO("SWIFTANEN")
recYSWIFTA0.SWIFTAPRI = rsADO("SWIFTAPRI")
recYSWIFTA0.SWIFTAMES = rsADO("SWIFTAMES")
recYSWIFTA0.SWIFTADOR = rsADO("SWIFTADOR")
recYSWIFTA0.SWIFTADES = rsADO("SWIFTADES")
recYSWIFTA0.SWIFTADVA = rsADO("SWIFTADVA")
recYSWIFTA0.SWIFTADE1 = rsADO("SWIFTADE1")
recYSWIFTA0.SWIFTAMON = rsADO("SWIFTAMON")
recYSWIFTA0.SWIFTADE2 = rsADO("SWIFTADE2")
recYSWIFTA0.SWIFTADEN = rsADO("SWIFTADEN")
recYSWIFTA0.SWIFTAHEN = rsADO("SWIFTAHEN")
recYSWIFTA0.SWIFTACOM = rsADO("SWIFTACOM")
recYSWIFTA0.SWIFTATES = rsADO("SWIFTATES")
recYSWIFTA0.SWIFTASUP = rsADO("SWIFTASUP")
recYSWIFTA0.SWIFTAVAL = rsADO("SWIFTAVAL")
recYSWIFTA0.SWIFTAAGE = rsADO("SWIFTAAGE")
recYSWIFTA0.SWIFTASER = rsADO("SWIFTASER")
recYSWIFTA0.SWIFTASSE = rsADO("SWIFTASSE")
recYSWIFTA0.SWIFTAUTI = rsADO("SWIFTAUTI")
recYSWIFTA0.SWIFTANUM = rsADO("SWIFTANUM")
recYSWIFTA0.SWIFTAUT1 = rsADO("SWIFTAUT1")
recYSWIFTA0.SWIFTAPVA = rsADO("SWIFTAPVA")
recYSWIFTA0.SWIFTAUT2 = rsADO("SWIFTAUT2")
Exit Function
Error_Handler:
srvYSWIFTA0_GetBuffer_ODBC = Error
End Function
Public Sub srvYSWIFTA0_ElpDisplay(recYSWIFTA0 As typeYSWIFTA0)
frmElpDisplay.fgData.Rows = 26
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTAETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTAETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTAREF   16A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REFERNECE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTAREF
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTANEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DE RENVOI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTANEN
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTAPRI    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE PROIRITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTAPRI
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTAMES    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE MESSAGE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTAMES
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTADOR   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DONNEUR ORDRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTADOR
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTADES   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DESTINATAIRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTADES
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTADVA    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE VALEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTADVA
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTADE1    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE 1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTADE1
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTAMON 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE VALEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTAMON
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTADE2    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTADE2
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTADEN    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ENVOI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTADEN
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTAHEN    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "HEURE ENVOI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTAHEN
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTACOM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPLET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTACOM
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTATES    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TEST OU REEL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTATES
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTASUP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SUPPRIME"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTASUP
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTAVAL    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TOP VALIDATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTAVAL
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTAAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTAAGE
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTASER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTASER
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTASSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTASSE
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTAUTI   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTAUTI
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTANUM    8P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO INTERNE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTANUM
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTAUT1   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISA SAISIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTAUT1
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTAPVA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "1ERE VALIDATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTAPVA
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIFTAUT2   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISA 1ER VALID"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIFTA0.SWIFTAUT2
frmElpDisplay.Show vbModal
End Sub

