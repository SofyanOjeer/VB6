Attribute VB_Name = "srvYSWIHIA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const constYSWIHIA0 = "YSWIHIA0"
Type typeYSWIHIA0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    SWIHIAETA       As Integer                        ' ETABLISSEMENT
    SWIHIAREF       As String * 16                    ' REFERNECE
    SWIHIANEN       As String * 1                     ' NUMERO DE RENVOI
    SWIHIAPRI       As String * 2                     ' CODE PROIRITE
    SWIHIAMES       As String * 3                     ' TYPE MESSAGE
    SWIHIADOR       As String * 12                    ' DONNEUR ORDRE
    SWIHIADES       As String * 12                    ' DESTINATAIRE
    SWIHIADVA       As Long                           ' DATE VALEUR
    SWIHIADE1       As String * 3                     ' DEVISE 1
    SWIHIAMON       As Currency                       ' MONTANT
    SWIHIADE2       As String * 3                     ' DEVISE 2
    SWIHIADEN       As Long                           ' DATE ENVOI
    SWIHIAHEN       As Long                           ' HEURE ENVOI
    SWIHIACOM       As String * 1                     ' COMPLET
    SWIHIATES       As String * 1                     ' TEST OU REEL
    SWIHIASUP       As String * 1                     ' SUPPRIME
    SWIHIAVAL       As String * 1                     ' TOP VALIDATION
    SWIHIAAGE       As Integer                        ' AGENCE
    SWIHIASER       As String * 2                     ' SERVICE
    SWIHIASSE       As String * 2                     ' SOUS SERVICE
    SWIHIAUTI       As String * 10                    ' UTILISATEUR
    SWIHIANUM       As Long                           ' NUMERO INTERNE
    SWIHIAUT1       As String * 10                    ' UTILISA SAISIE
    SWIHIAPVA       As String * 1                     ' 1ERE VALIDATION
    SWIHIAUT2       As String * 10                    ' UTILISA 1ER VALID

End Type
Public Sub srvYSWIHIA0_Init(recYSWIHIA0 As typeYSWIHIA0)
recYSWIHIA0.Obj = "YSWIHIA0"
recYSWIHIA0.Method = ""
recYSWIHIA0.Err = ""
recYSWIHIA0.SWIHIAETA = 0
recYSWIHIA0.SWIHIAREF = ""
recYSWIHIA0.SWIHIANEN = ""
recYSWIHIA0.SWIHIAPRI = ""
recYSWIHIA0.SWIHIAMES = ""
recYSWIHIA0.SWIHIADOR = ""
recYSWIHIA0.SWIHIADES = ""
recYSWIHIA0.SWIHIADVA = 0
recYSWIHIA0.SWIHIADE1 = ""
recYSWIHIA0.SWIHIAMON = 0
recYSWIHIA0.SWIHIADE2 = ""
recYSWIHIA0.SWIHIADEN = 0
recYSWIHIA0.SWIHIAHEN = 0
recYSWIHIA0.SWIHIACOM = ""
recYSWIHIA0.SWIHIATES = ""
recYSWIHIA0.SWIHIASUP = ""
recYSWIHIA0.SWIHIAVAL = ""
recYSWIHIA0.SWIHIAAGE = 0
recYSWIHIA0.SWIHIASER = ""
recYSWIHIA0.SWIHIASSE = ""
recYSWIHIA0.SWIHIAUTI = ""
recYSWIHIA0.SWIHIANUM = 0
recYSWIHIA0.SWIHIAUT1 = ""
recYSWIHIA0.SWIHIAPVA = ""
recYSWIHIA0.SWIHIAUT2 = ""
End Sub
Public Function srvYSWIHIA0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYSWIHIA0 As typeYSWIHIA0)
On Error GoTo Error_Handler
srvYSWIHIA0_GetBuffer_ODBC = Null
recYSWIHIA0.SWIHIAETA = rsADO("SWIHIAETA")
recYSWIHIA0.SWIHIAREF = rsADO("SWIHIAREF")
recYSWIHIA0.SWIHIANEN = rsADO("SWIHIANEN")
recYSWIHIA0.SWIHIAPRI = rsADO("SWIHIAPRI")
recYSWIHIA0.SWIHIAMES = rsADO("SWIHIAMES")
recYSWIHIA0.SWIHIADOR = rsADO("SWIHIADOR")
recYSWIHIA0.SWIHIADES = rsADO("SWIHIADES")
recYSWIHIA0.SWIHIADVA = rsADO("SWIHIADVA")
recYSWIHIA0.SWIHIADE1 = rsADO("SWIHIADE1")
recYSWIHIA0.SWIHIAMON = rsADO("SWIHIAMON")
recYSWIHIA0.SWIHIADE2 = rsADO("SWIHIADE2")
recYSWIHIA0.SWIHIADEN = rsADO("SWIHIADEN")
recYSWIHIA0.SWIHIAHEN = rsADO("SWIHIAHEN")
recYSWIHIA0.SWIHIACOM = rsADO("SWIHIACOM")
recYSWIHIA0.SWIHIATES = rsADO("SWIHIATES")
recYSWIHIA0.SWIHIASUP = rsADO("SWIHIASUP")
recYSWIHIA0.SWIHIAVAL = rsADO("SWIHIAVAL")
recYSWIHIA0.SWIHIAAGE = rsADO("SWIHIAAGE")
recYSWIHIA0.SWIHIASER = rsADO("SWIHIASER")
recYSWIHIA0.SWIHIASSE = rsADO("SWIHIASSE")
recYSWIHIA0.SWIHIAUTI = rsADO("SWIHIAUTI")
recYSWIHIA0.SWIHIANUM = rsADO("SWIHIANUM")
recYSWIHIA0.SWIHIAUT1 = rsADO("SWIHIAUT1")
recYSWIHIA0.SWIHIAPVA = rsADO("SWIHIAPVA")
recYSWIHIA0.SWIHIAUT2 = rsADO("SWIHIAUT2")
Exit Function
Error_Handler:
srvYSWIHIA0_GetBuffer_ODBC = Error
End Function
Public Sub srvYSWIHIA0_ElpDisplay(recYSWIHIA0 As typeYSWIHIA0)
frmElpDisplay.fgData.Rows = 26
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIAETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIAETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIAREF   16A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REFERNECE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIAREF
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIANEN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DE RENVOI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIANEN
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIAPRI    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE PROIRITE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIAPRI
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIAMES    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE MESSAGE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIAMES
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIADOR   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DONNEUR ORDRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIADOR
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIADES   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DESTINATAIRE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIADES
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIADVA    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE VALEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIADVA
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIADE1    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE 1"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIADE1
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIAMON 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIAMON
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIADE2    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE 2"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIADE2
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIADEN    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE ENVOI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIADEN
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIAHEN    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "HEURE ENVOI"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIAHEN
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIACOM    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPLET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIACOM
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIATES    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TEST OU REEL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIATES
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIASUP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SUPPRIME"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIASUP
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIAVAL    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TOP VALIDATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIAVAL
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIAAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIAAGE
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIASER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIASER
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIASSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIASSE
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIAUTI   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIAUTI
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIANUM    8P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO INTERNE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIANUM
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIAUT1   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISA SAISIE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIAUT1
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIAPVA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "1ERE VALIDATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIAPVA
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "SWIHIAUT2   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISA 1ER VALID"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYSWIHIA0.SWIHIAUT2
frmElpDisplay.Show vbModal
End Sub

