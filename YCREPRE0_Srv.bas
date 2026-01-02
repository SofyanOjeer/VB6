Attribute VB_Name = "srvYCREPRE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYCREPRE0Len = 500 ' 34 + ??????
Public Const recYCREPRE0_Block = 100 '????
Public Const constYCREPRE0 = "YCREPRE0"
Dim meYbase As typeYBase
Dim paramYCREPRE0_Import As String

Type typeYCREPRE0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CREPREETA       As Integer                        ' ETABLISSEMENT
    CREPREAGE       As Integer                        ' AGENCE
    CREPRESCE       As String * 2                     ' SERVICE
    CREPRESSE       As String * 2                     ' SOUS SERVICE
    CREPREDOS       As Long                           ' NUMERO DOSSIER
    CREPREPRE       As Long                           ' NUMERO PRET
    CREPRENAT       As String * 3                     ' NATURE PRET
    CREPREDEV       As String * 3                     ' DEVISE PRET
    CREPREDAE       As Long                           ' DERN ECH K TRAITEE
    CREPREDPE       As Long                           ' PROC ECH K A TRAITER
    CREPRECAL       As Long                           ' DTE DERN CALCUL INT
    CREPREDAI       As Long                           ' DERN ECH I TRAITEE
    CREPREDPI       As Long                           ' PROC ECH I A TRAITER
    CREPREDET       As Long                           ' DATE CODE ETAT
    CREPREOUV       As Long                           ' DATE OUVER PRET
    CREPREDER       As Long                           ' DATE DER. MODIF
    CREPREDIC       As Long                           ' DTE DERN CALCUL IC
    CREPREMON       As Currency                       ' MONTANT PRET
    CREPRECAP       As Currency                       ' CAPITAL RESTANT DU
    CREPREINT       As Currency                       ' INTERETS REPORTES
    CREPREICO       As Currency                       ' INTERETS COURUS
    CREPREICV       As Currency                       ' INTERETS COURUS CTVL
    CREPRECRS       As Double                         ' COURS
    CREPRECTA       As Long                           ' CODE ETAT
    CREPREPLA       As Long                           ' NUMERO PLAN
    CREPREPAL       As Long                           ' NUMERO PALIER
    CREPREECH       As Long                           ' NUMERO ECHEANCE
    CREPREAVI       As String * 1                     ' AVIS ECHEANCE
    CREPRETYR       As String * 1                     ' ANC.TYPE REPORT TYR
    CREPREINR       As String * 1                     ' ANC.INTEGRATION INR
    CREPREBAS       As Long                           ' ANC.BASE NBJ/AN BAS
    CREPREREA       As String * 1                     ' ANC.JOURS REELS REA
    CREPREPRC       As String * 1                     ' INTERETS PRECOMPTES
    CREPRESUP       As Long                           ' JOURS SUPPLEMENTS
    CREPRECOM       As Long                           ' ITC-COMMI-ASSUR
    CREPREAUT       As String * 12                    ' AUTORISATION
    CREPREUTI       As Integer                        ' UTILISATEUR
    CREPREOBJ       As String * 6                     ' OBJET PRET
    CREPREBAR       As String * 6                     ' BAREME
    CREPREREM       As String * 6                     ' REMBT ANTICIPE
    CREPREIMP       As String * 6                     ' GESTION IMPAYE
    CREPREFNC       As Double                         ' TAUX REFINANCEMENT
    CREPREINC       As Integer                        ' INCREMENT ECHEANCE
    CREPRETDO       As String * 1                     ' TYPE DE RECOUVREMENT
    CREPRESUS       As String * 1                     ' SUSPENSION DE RECOUV
    CREPREEXI       As String * 1                     ' EXISTENCE DS IMPAYES
    CREPREAGI       As String * 1                     ' AGIOS COMPENSES
    CREPRERGL       As String * 1                     ' COMPTE CREDITEUR
    CREPRECOD       As Integer                        ' UTILISAT. DERN MODIF
    CREPREOPT       As Long                           ' OPTION DERNIER MODIF

End Type
    
'---------------------------------------------------------
Public Function srvYCREPRE0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYCREPRE0 As typeYCREPRE0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYCREPRE0_GetBuffer_ODBC = Null

    recYCREPRE0.CREPREETA = rsADO("CREPREETA") 'CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYCREPRE0.CREPREAGE = rsADO("CREPREAGE") 'CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYCREPRE0.CREPRESCE = rsADO("CREPRESCE") 'mId$(MsgTxt, K + 11, 2)
    recYCREPRE0.CREPRESSE = rsADO("CREPRESSE") 'mId$(MsgTxt, K + 13, 2)
    recYCREPRE0.CREPREDOS = rsADO("CREPREDOS") 'CLng(Val(mId$(MsgTxt, K + 15, 8)))
    recYCREPRE0.CREPREPRE = rsADO("CREPREPRE") 'CLng(Val(mId$(MsgTxt, K + 23, 4)))
    recYCREPRE0.CREPRENAT = rsADO("CREPRENAT") 'mId$(MsgTxt, K + 27, 3)
    recYCREPRE0.CREPREDEV = rsADO("CREPREDEV") 'mId$(MsgTxt, K + 30, 3)
    recYCREPRE0.CREPREDAE = rsADO("CREPREDAE") 'CLng(Val(mId$(MsgTxt, K + 33, 8)))
    recYCREPRE0.CREPREDPE = rsADO("CREPREDPE") 'CLng(Val(mId$(MsgTxt, K + 41, 8)))
    recYCREPRE0.CREPRECAL = rsADO("CREPRECAL") 'CLng(Val(mId$(MsgTxt, K + 49, 8)))
    recYCREPRE0.CREPREDAI = rsADO("CREPREDAI") 'CLng(Val(mId$(MsgTxt, K + 57, 8)))
    recYCREPRE0.CREPREDPI = rsADO("CREPREDPI") 'CLng(Val(mId$(MsgTxt, K + 65, 8)))
    recYCREPRE0.CREPREDET = rsADO("CREPREDET") 'CLng(Val(mId$(MsgTxt, K + 73, 8)))
    recYCREPRE0.CREPREOUV = rsADO("CREPREOUV") 'CLng(Val(mId$(MsgTxt, K + 81, 8)))
    recYCREPRE0.CREPREDER = rsADO("CREPREDER") 'CLng(Val(mId$(MsgTxt, K + 89, 8)))
    recYCREPRE0.CREPREDIC = rsADO("CREPREDIC") 'CLng(Val(mId$(MsgTxt, K + 97, 8)))
    recYCREPRE0.CREPREMON = rsADO("CREPREMON") 'CCur(Val(mId$(MsgTxt, K + 105, 16))) / 100
    recYCREPRE0.CREPRECAP = rsADO("CREPRECAP") 'CCur(Val(mId$(MsgTxt, K + 121, 16))) / 100
    recYCREPRE0.CREPREINT = rsADO("CREPREINT") 'CCur(Val(mId$(MsgTxt, K + 137, 16))) / 100
    recYCREPRE0.CREPREICO = rsADO("CREPREICO") 'CCur(Val(mId$(MsgTxt, K + 153, 16))) / 100
    recYCREPRE0.CREPREICV = rsADO("CREPREICV") 'CCur(Val(mId$(MsgTxt, K + 169, 16))) / 100
    recYCREPRE0.CREPRECRS = rsADO("CREPRECRS") 'CDbl(Val(mId$(MsgTxt, K + 185, 16))) / 10000000000#
    recYCREPRE0.CREPRECTA = rsADO("CREPRECTA") 'CLng(Val(mId$(MsgTxt, K + 201, 4)))
    recYCREPRE0.CREPREPLA = rsADO("CREPREPLA") 'CLng(Val(mId$(MsgTxt, K + 205, 4)))
    recYCREPRE0.CREPREPAL = rsADO("CREPREPAL") 'CLng(Val(mId$(MsgTxt, K + 209, 4)))
    recYCREPRE0.CREPREECH = rsADO("CREPREECH") 'CLng(Val(mId$(MsgTxt, K + 213, 4)))
    recYCREPRE0.CREPREAVI = rsADO("CREPREAVI") 'mId$(MsgTxt, K + 217, 1)
    recYCREPRE0.CREPRETYR = rsADO("CREPRETYR") 'mId$(MsgTxt, K + 218, 1)
    recYCREPRE0.CREPREINR = rsADO("CREPREINR") 'mId$(MsgTxt, K + 219, 1)
    recYCREPRE0.CREPREBAS = rsADO("CREPREBAS") 'CLng(Val(mId$(MsgTxt, K + 220, 2)))
    recYCREPRE0.CREPREREA = rsADO("CREPREREA") 'mId$(MsgTxt, K + 222, 1)
    recYCREPRE0.CREPREPRC = rsADO("CREPREPRC") 'mId$(MsgTxt, K + 223, 1)
    recYCREPRE0.CREPRESUP = rsADO("CREPRESUP") 'CLng(Val(mId$(MsgTxt, K + 224, 4)))
    recYCREPRE0.CREPRECOM = rsADO("CREPRECOM") 'CLng(Val(mId$(MsgTxt, K + 228, 4)))
    recYCREPRE0.CREPREAUT = rsADO("CREPREAUT") 'mId$(MsgTxt, K + 232, 12)
    recYCREPRE0.CREPREUTI = rsADO("CREPREUTI") 'CInt(Val(mId$(MsgTxt, K + 244, 5)))
    recYCREPRE0.CREPREOBJ = rsADO("CREPREOBJ") 'mId$(MsgTxt, K + 249, 6)
    recYCREPRE0.CREPREBAR = rsADO("CREPREBAR") 'mId$(MsgTxt, K + 255, 6)
    recYCREPRE0.CREPREREM = rsADO("CREPREREM") 'mId$(MsgTxt, K + 261, 6)
    recYCREPRE0.CREPREIMP = rsADO("CREPREIMP") 'mId$(MsgTxt, K + 267, 6)
    recYCREPRE0.CREPREFNC = rsADO("CREPREFNC") 'CDbl(Val(mId$(MsgTxt, K + 273, 8))) / 100000
    recYCREPRE0.CREPREINC = rsADO("CREPREINC") 'CInt(Val(mId$(MsgTxt, K + 281, 5)))
    recYCREPRE0.CREPRETDO = rsADO("CREPRETDO") 'mId$(MsgTxt, K + 286, 1)
    recYCREPRE0.CREPRESUS = rsADO("CREPRESUS") 'mId$(MsgTxt, K + 287, 1)
    recYCREPRE0.CREPREEXI = rsADO("CREPREEXI") 'mId$(MsgTxt, K + 288, 1)
    recYCREPRE0.CREPREAGI = rsADO("CREPREAGI") 'mId$(MsgTxt, K + 289, 1)
    recYCREPRE0.CREPRERGL = rsADO("CREPRERGL") 'mId$(MsgTxt, K + 290, 1)
    recYCREPRE0.CREPRECOD = rsADO("CREPRECOD") 'CInt(Val(mId$(MsgTxt, K + 291, 5)))
    recYCREPRE0.CREPREOPT = rsADO("CREPREOPT") 'CLng(Val(mId$(MsgTxt, K + 296, 8)))
Exit Function

Error_Handler:
srvYCREPRE0_GetBuffer_ODBC = Error

End Function


Public Sub srvYCREPRE0_ElpDisplay(recYCREPRE0 As typeYCREPRE0)
frmElpDisplay.fgData.Rows = 51
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREETA    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREETA
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREAGE
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPRESCE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPRESCE
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPRESSE    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPRESSE
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREDOS    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO DOSSIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREDOS
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREPRE    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREPRE
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPRENAT    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NATURE PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPRENAT
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREDEV    3A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DEVISE PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREDEV
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREDAE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DERN ECH K TRAITEE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREDAE
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREDPE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PROC ECH K A TRAITER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREDPE
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPRECAL    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DTE DERN CALCUL INT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPRECAL
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREDAI    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DERN ECH I TRAITEE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREDAI
frmElpDisplay.fgData.Row = 13
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREDPI    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "PROC ECH I A TRAITER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREDPI
frmElpDisplay.fgData.Row = 14
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREDET    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE CODE ETAT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREDET
frmElpDisplay.fgData.Row = 15
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREOUV    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE OUVER PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREOUV
frmElpDisplay.fgData.Row = 16
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREDER    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DATE DER. MODIF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREDER
frmElpDisplay.fgData.Row = 17
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREDIC    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DTE DERN CALCUL IC"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREDIC
frmElpDisplay.fgData.Row = 18
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREMON 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MONTANT PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREMON
frmElpDisplay.fgData.Row = 19
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPRECAP 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CAPITAL RESTANT DU"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPRECAP
frmElpDisplay.fgData.Row = 20
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREINT 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERETS REPORTES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREINT
frmElpDisplay.fgData.Row = 21
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREICO 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERETS COURUS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREICO
frmElpDisplay.fgData.Row = 22
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREICV 15.2P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERETS COURUS CTVL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREICV
frmElpDisplay.fgData.Row = 23
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPRECRS15.10P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COURS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPRECRS
frmElpDisplay.fgData.Row = 24
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPRECTA    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE ETAT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPRECTA
frmElpDisplay.fgData.Row = 25
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREPLA    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PLAN"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREPLA
frmElpDisplay.fgData.Row = 26
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREPAL    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO PALIER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREPAL
frmElpDisplay.fgData.Row = 27
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREECH    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NUMERO ECHEANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREECH
frmElpDisplay.fgData.Row = 28
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREAVI    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AVIS ECHEANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREAVI
frmElpDisplay.fgData.Row = 29
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPRETYR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ANC.TYPE REPORT TYR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPRETYR
frmElpDisplay.fgData.Row = 30
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREINR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ANC.INTEGRATION INR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREINR
frmElpDisplay.fgData.Row = 31
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREBAS    1S"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ANC.BASE NBJ/AN BAS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREBAS
frmElpDisplay.fgData.Row = 32
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREREA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ANC.JOURS REELS REA"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREREA
frmElpDisplay.fgData.Row = 33
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREPRC    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTERETS PRECOMPTES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREPRC
frmElpDisplay.fgData.Row = 34
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPRESUP    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "JOURS SUPPLEMENTS"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPRESUP
frmElpDisplay.fgData.Row = 35
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPRECOM    3P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ITC-COMMI-ASSUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPRECOM
frmElpDisplay.fgData.Row = 36
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREAUT   12A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AUTORISATION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREAUT
frmElpDisplay.fgData.Row = 37
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREUTI    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREUTI
frmElpDisplay.fgData.Row = 38
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREOBJ    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "OBJET PRET"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREOBJ
frmElpDisplay.fgData.Row = 39
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREBAR    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "BAREME"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREBAR
frmElpDisplay.fgData.Row = 40
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREREM    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "REMBT ANTICIPE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREREM
frmElpDisplay.fgData.Row = 41
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREIMP    6A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "GESTION IMPAYE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREIMP
frmElpDisplay.fgData.Row = 42
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREFNC  7.5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TAUX REFINANCEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREFNC
frmElpDisplay.fgData.Row = 43
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREINC    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INCREMENT ECHEANCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREINC
frmElpDisplay.fgData.Row = 44
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPRETDO    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "TYPE DE RECOUVREMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPRETDO
frmElpDisplay.fgData.Row = 45
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPRESUS    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SUSPENSION DE RECOUV"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPRESUS
frmElpDisplay.fgData.Row = 46
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREEXI    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "EXISTENCE DS IMPAYES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREEXI
frmElpDisplay.fgData.Row = 47
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREAGI    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGIOS COMPENSES"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREAGI
frmElpDisplay.fgData.Row = 48
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPRERGL    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "COMPTE CREDITEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPRERGL
frmElpDisplay.fgData.Row = 49
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPRECOD    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISAT. DERN MODIF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPRECOD
frmElpDisplay.fgData.Row = 50
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CREPREOPT    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "OPTION DERNIER MODIF"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYCREPRE0.CREPREOPT
frmElpDisplay.Show vbModal
End Sub
'---------------------------------------------------------
Public Sub recYCREPRE0_Init(recYCREPRE0 As typeYCREPRE0)
'---------------------------------------------------------
recYCREPRE0.obj = "ZCREPRE0_S"

recYCREPRE0.CREPREETA = 0   '       As Integer                        ' ETABLISSEMENT
recYCREPRE0.CREPREAGE = 0   '       As Integer                        ' AGENCE
recYCREPRE0.CREPRESCE = ""  '       As String * 2                     ' SERVICE
recYCREPRE0.CREPRESSE = ""  '       As String * 2                     ' SOUS SERVICE
recYCREPRE0.CREPREDOS = 0   '       As Long                           ' NUMERO DOSSIER
recYCREPRE0.CREPREPRE = 0   '       As Long                           ' NUMERO PRET
recYCREPRE0.CREPRENAT = ""  '       As String * 3                     ' NATURE PRET
recYCREPRE0.CREPREDEV = ""  '       As String * 3                     ' DEVISE PRET
recYCREPRE0.CREPREDAE = 0   '       As Long                           ' DERN ECH K TRAITEE
recYCREPRE0.CREPREDPE = 0   '       As Long                           ' PROC ECH K A TRAITER
recYCREPRE0.CREPRECAL = 0   '       As Long                           ' DTE DERN CALCUL INT
recYCREPRE0.CREPREDAI = 0   '       As Long                           ' DERN ECH I TRAITEE
recYCREPRE0.CREPREDPI = 0   '       As Long                           ' PROC ECH I A TRAITER
recYCREPRE0.CREPREDET = 0   '       As Long                           ' DATE CODE ETAT
recYCREPRE0.CREPREOUV = 0   '       As Long                           ' DATE OUVER PRET
recYCREPRE0.CREPREDER = 0   '       As Long                           ' DATE DER. MODIF
recYCREPRE0.CREPREDIC = 0   '       As Long                           ' DTE DERN CALCUL IC
recYCREPRE0.CREPREMON = 0   '       As Currency                       ' MONTANT PRET
recYCREPRE0.CREPRECAP = 0   '       As Currency                       ' CAPITAL RESTANT DU
recYCREPRE0.CREPREINT = 0   '       As Currency                       ' INTERETS REPORTES
recYCREPRE0.CREPREICO = 0   '       As Currency                       ' INTERETS COURUS
recYCREPRE0.CREPREICV = 0   '       As Currency                       ' INTERETS COURUS CTVL
recYCREPRE0.CREPRECRS = 0   '       As Double                         ' COURS
recYCREPRE0.CREPRECTA = 0   '       As Long                           ' CODE ETAT
recYCREPRE0.CREPREPLA = 0   '       As Long                           ' NUMERO PLAN
recYCREPRE0.CREPREPAL = 0   '       As Long                           ' NUMERO PALIER
recYCREPRE0.CREPREECH = 0   '       As Long                           ' NUMERO ECHEANCE
recYCREPRE0.CREPREAVI = ""  '       As String * 1                     ' AVIS ECHEANCE
recYCREPRE0.CREPRETYR = ""  '       As String * 1                     ' ANC.TYPE REPORT TYR
recYCREPRE0.CREPREINR = ""  '       As String * 1                     ' ANC.INTEGRATION INR
recYCREPRE0.CREPREBAS = 0   '       As Long                           ' ANC.BASE NBJ/AN BAS
recYCREPRE0.CREPREREA = ""  '       As String * 1                     ' ANC.JOURS REELS REA
recYCREPRE0.CREPREPRC = ""  '       As String * 1                     ' INTERETS PRECOMPTES
recYCREPRE0.CREPRESUP = 0   '       As Long                           ' JOURS SUPPLEMENTS
recYCREPRE0.CREPRECOM = 0   '       As Long                           ' ITC-COMMI-ASSUR
recYCREPRE0.CREPREAUT = ""  '       As String * 12                    ' AUTORISATION
recYCREPRE0.CREPREUTI = 0   '       As Integer                        ' UTILISATEUR
recYCREPRE0.CREPREOBJ = ""  '       As String * 6                     ' OBJET PRET
recYCREPRE0.CREPREBAR = ""  '       As String * 6                     ' BAREME
recYCREPRE0.CREPREREM = ""  '       As String * 6                     ' REMBT ANTICIPE
recYCREPRE0.CREPREIMP = ""  '       As String * 6                     ' GESTION IMPAYE
recYCREPRE0.CREPREFNC = 0   '       As Double                         ' TAUX REFINANCEMENT
recYCREPRE0.CREPREINC = 0   '       As Integer                        ' INCREMENT ECHEANCE
recYCREPRE0.CREPRETDO = ""  '       As String * 1                     ' TYPE DE RECOUVREMENT
recYCREPRE0.CREPRESUS = ""  '       As String * 1                     ' SUSPENSION DE RECOUV
recYCREPRE0.CREPREEXI = ""  '       As String * 1                     ' EXISTENCE DS IMPAYES
recYCREPRE0.CREPREAGI = ""  '       As String * 1                     ' AGIOS COMPENSES
recYCREPRE0.CREPRERGL = ""  '       As String * 1                     ' COMPTE CREDITEUR
recYCREPRE0.CREPRECOD = 0   '       As Integer                        ' UTILISAT. DERN MODIF
recYCREPRE0.CREPREOPT = 0   '       As Long                           ' OPTION DERNIER MODIF

End Sub




