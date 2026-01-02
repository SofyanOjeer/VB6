Attribute VB_Name = "srvYCREDOS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCREDOS0Len = 500 ' 34 + ??????
Public Const recYCREDOS0_Block = 100 '????
Public Const constYCREDOS0 = "YCREDOS0"
Dim meYbase As typeYBase
Dim paramYCREDOS0_Import As String

Type typeYCREDOS0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CREDOSETA       As Integer                        ' ETABLISSEMENT
    CREDOSAGE       As Integer                        ' AGENCE
    CREDOSSER       As String * 2                     ' SERVICE
    CREDOSSSE       As String * 2                     ' SOUS-SERVICE
    CREDOSDOS       As Long                           ' NUMERO DOSSIER
    CREDOSNCR       As String * 3                     ' NATURE CREDIT
    CREDOSMNT       As Currency                       ' MONTANT
    CREDOSDEV       As String * 3                     ' DEVISE
    CREDOSDDE       As Long                           ' AUTORISATION
    CREDOSDFI       As Long                           ' AUTORISATION
    CREDOSREF       As String * 50                    ' REFERENCES
    CREDOSUTI       As Integer                        ' UTILISATEUR
    CREDOSDMO       As Long                           ' DATE MODIFICATION
    CREDOSOFI       As String * 6                     ' OBJET FINANCEMENT
    CREDOSCET       As Long                           ' CODE ETAT
    CREDOSDCE       As Long                           ' DATE CODE ETAT
    CREDOSDOD       As Long                           ' DU DOSSIER
    CREDOSDVA       As Long                           ' DATE VALIDATION
    CREDOSDGE       As Long                           ' CREDIT ENGAGE
    CREDOSTYP       As String * 1                     ' TYPE DE CREDIT
    CREDOSCOP       As Long                           ' CO-PARTICIPATION
End Type
    
'---------------------------------------------------------
Public Function srvYCREDOS0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCREDOS0 As typeYCREDOS0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCREDOS0_GetBuffer_ODBC = Null

    recYCREDOS0.CREDOSETA = rsADO("CREDOSETA") 'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCREDOS0.CREDOSAGE = rsADO("CREDOSAGE") 'CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCREDOS0.CREDOSSER = rsADO("CREDOSSER") 'mId$(MsgTxt, K + 11, 2)
    recYCREDOS0.CREDOSSSE = rsADO("CREDOSSSE")    'mId$(MsgTxt, K + 13, 2)
    recYCREDOS0.CREDOSDOS = rsADO("CREDOSDOS")    'CLng(Val(mId$(MsgTxt, K + 15, 8)))
    recYCREDOS0.CREDOSNCR = rsADO("CREDOSNCR")    'mId$(MsgTxt, K + 23, 3)
    recYCREDOS0.CREDOSMNT = rsADO("CREDOSMNT")    'CCur(Val(mId$(MsgTxt, K + 26, 16))) / 100
    recYCREDOS0.CREDOSDEV = rsADO("CREDOSDEV")    'mId$(MsgTxt, K + 42, 3)
    recYCREDOS0.CREDOSDDE = rsADO("CREDOSDDE")    'CLng(Val(mId$(MsgTxt, K + 45, 8)))
    recYCREDOS0.CREDOSDFI = rsADO("CREDOSDFI")    'CLng(Val(mId$(MsgTxt, K + 53, 8)))
    recYCREDOS0.CREDOSREF = rsADO("CREDOSREF")    'mId$(MsgTxt, K + 61, 50)
    recYCREDOS0.CREDOSUTI = rsADO("CREDOSUTI")    'CInt(Val(mId$(MsgTxt, K + 111, 5)))
    recYCREDOS0.CREDOSDMO = rsADO("CREDOSDMO")    'CLng(Val(mId$(MsgTxt, K + 116, 8)))
    recYCREDOS0.CREDOSOFI = rsADO("CREDOSOFI")    'mId$(MsgTxt, K + 124, 6)
    recYCREDOS0.CREDOSCET = rsADO("CREDOSCET")    'CLng(Val(mId$(MsgTxt, K + 130, 4)))
    recYCREDOS0.CREDOSDCE = rsADO("CREDOSDCE")    'CLng(Val(mId$(MsgTxt, K + 134, 8)))
    recYCREDOS0.CREDOSDOD = rsADO("CREDOSDOD")    'CLng(Val(mId$(MsgTxt, K + 142, 8)))
    recYCREDOS0.CREDOSDVA = rsADO("CREDOSDVA")    'CLng(Val(mId$(MsgTxt, K + 150, 8)))
    recYCREDOS0.CREDOSDGE = rsADO("CREDOSDGE")    'CLng(Val(mId$(MsgTxt, K + 158, 8)))
    recYCREDOS0.CREDOSTYP = rsADO("CREDOSTYP")    'mId$(MsgTxt, K + 166, 1)
    recYCREDOS0.CREDOSCOP = rsADO("CREDOSCOP")    'CLng(Val(mId$(MsgTxt, K + 167, 4)))
Exit Function

Error_Handler:
srvYCREDOS0_GetBuffer_ODBC = Error

End Function


Public Sub srvYCREDOS0_ElpDisplay(recYCREDOS0 As typeYCREDOS0)
frmElpDisplay.fgData.Rows = 22
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSSER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSSER
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSSSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSSSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSDOS    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSDOS
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSNCR    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NATURE CREDIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSNCR
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSMNT 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSMNT
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSDEV
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSDDE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AUTORISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSDDE
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSDFI    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AUTORISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSDFI
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSREF   50A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REFERENCES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSREF
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSUTI    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSUTI
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSDMO    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE MODIFICATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSDMO
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSOFI    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "OBJET FINANCEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSOFI
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSCET    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETAT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSCET
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSDCE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE CODE ETAT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSDCE
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSDOD    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DU DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSDOD
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSDVA    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE VALIDATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSDVA
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSDGE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CREDIT ENGAGE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSDGE
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSTYP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE DE CREDIT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSTYP
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREDOSCOP    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CO-PARTICIPATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREDOS0.CREDOSCOP
frmElpDisplay.Show vbModal
End Sub
'---------------------------------------------------------
Public Sub recYCREDOS0_Init(recYCREDOS0 As typeYCREDOS0)
'---------------------------------------------------------
'MsgTxt = Space$(recYCREDOS0Len)
'MsgTxtIndex = 0
'Call srvYCREDOS0_GetBuffer(recYCREDOS0)
recYCREDOS0.obj = "ZCREDOS0_S"

recYCREDOS0.CREDOSETA = 0 '       As Integer                        ' ETABLISSEMENT
recYCREDOS0.CREDOSAGE = 0 '       As Integer                        ' AGENCE
recYCREDOS0.CREDOSSER = "" '       As String * 2                     ' SERVICE
recYCREDOS0.CREDOSSSE = "" '       As String * 2                     ' SOUS-SERVICE
recYCREDOS0.CREDOSDOS = 0 '       As Long                           ' NUMERO DOSSIER
recYCREDOS0.CREDOSNCR = "" '       As String * 3                     ' NATURE CREDIT
recYCREDOS0.CREDOSMNT = 0 '       As Currency                       ' MONTANT
recYCREDOS0.CREDOSDEV = "" '       As String * 3                     ' DEVISE
recYCREDOS0.CREDOSDDE = 0 '       As Long                           ' AUTORISATION
recYCREDOS0.CREDOSDFI = 0 '       As Long                           ' AUTORISATION
recYCREDOS0.CREDOSREF = "" '       As String * 50                    ' REFERENCES
recYCREDOS0.CREDOSUTI = 0 '       As Integer                        ' UTILISATEUR
recYCREDOS0.CREDOSDMO = 0 '       As Long                           ' DATE MODIFICATION
recYCREDOS0.CREDOSOFI = "" '       As String * 6                     ' OBJET FINANCEMENT
recYCREDOS0.CREDOSCET = 0 '       As Long                           ' CODE ETAT
recYCREDOS0.CREDOSDCE = 0 '       As Long                           ' DATE CODE ETAT
recYCREDOS0.CREDOSDOD = 0 '       As Long                           ' DU DOSSIER
recYCREDOS0.CREDOSDVA = 0 '       As Long                           ' DATE VALIDATION
recYCREDOS0.CREDOSDGE = 0 '       As Long                           ' CREDIT ENGAGE
recYCREDOS0.CREDOSTYP = "" '       As String * 1                     ' TYPE DE CREDIT
recYCREDOS0.CREDOSCOP = 0 '       As Long                           ' CO-PARTICIPATION
End Sub



