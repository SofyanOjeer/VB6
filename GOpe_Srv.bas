Attribute VB_Name = "srvGOpe"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recGOpeLen = 442 ' 34 + 408
Public Const recGOpe_Block = 15

Type typeGOpe
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    IdRéférence             As Long
    Application             As String * 5
    Nature                  As String * 5
    
    Devise1                 As String * 3
    Montant1                As Currency
    TauxRéférence1          As String * 10
    TauxMarge1              As Double
    TauxActuariel1          As Double
    TEG1                    As Double
    
    Devise2                 As String * 3
    Montant2                As Currency
    TauxRéférence2          As String * 10
    TauxMarge2              As Double
    
    AmjEngagement           As String * 8
    AmjDébut                As String * 8
    AmjFin                  As String * 8
    AmjEchéance1            As String * 8
    AmjEchéanceS            As String * 1
    PréavisNbj              As Integer
    Périodicité             As String * 1
    PériodeNb               As Integer
    IPA                     As String * 1
    NbjBase                 As String * 1
    Devise3                 As String * 3
   
    Mensualité              As Currency
    Frais1                  As Currency
    Frais2                  As Currency
    Frais3                  As Currency
    
    EngagementCompte        As String * 11
    EngagementCorrCompte    As String * 11
    EngagementCorrSwiftN    As String * 11
    EngagementCorrSwiftL    As String * 11
    EchéanceCompte          As String * 11
    EchéanceCorrCompte      As String * 11
    EchéanceCorrSwiftN      As String * 11
    EchéanceCorrSwiftL      As String * 11
    
    RéférenceInterne        As String * 16
    RéférenceExterne        As String * 16
    IdRéférenceLiée         As Long
    optReprise              As String * 1
    TauxRéférenceInterne    As String * 10
    TauxMargeInterne        As Double

    Statut                  As String * 1
    StatutPlus              As String * 2
    Flag1                   As String * 1
    Flag2                   As String * 1
    Flag3                   As String * 1
   
    ElpId                   As Long
    ElpUpdate               As Integer
    ElpControl              As String * 10
    
End Type
    
Public arrGOpe() As typeGOpe
Public arrGOpe_NB As Integer
Public arrGOpe_NBMax As Integer
Public arrGOpe_Index As Integer
Public arrGOpe_Suite As Boolean
Public Sub srvGOpe_ElpDisplay(recGOpe As typeGOpe)
frmElpDisplay.fgData.Rows = 54
frmElpDisplay.fgData.Row = 1
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "obj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.obj
frmElpDisplay.fgData.Row = 2
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Method"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Method
frmElpDisplay.fgData.Row = 3
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Err"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Err
frmElpDisplay.fgData.Row = 4
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "IdRéférence"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.IdRéférence
frmElpDisplay.fgData.Row = 5
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Application"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Application
frmElpDisplay.fgData.Row = 6
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Nature"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Nature
    
frmElpDisplay.fgData.Row = 7
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Devise1"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Devise1
frmElpDisplay.fgData.Row = 8
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Montant1"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Montant1
frmElpDisplay.fgData.Row = 9
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TauxRéférence1"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.TauxRéférence1
frmElpDisplay.fgData.Row = 10
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TauxMarge1"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.TauxMarge1
frmElpDisplay.fgData.Row = 11
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TauxActuariel1"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.TauxActuariel1
frmElpDisplay.fgData.Row = 12
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TEG1"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.TEG1
    
    
frmElpDisplay.fgData.Row = 13
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Devise2"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Devise2
frmElpDisplay.fgData.Row = 14
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Montant2"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Montant2
frmElpDisplay.fgData.Row = 15
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TauxRéférence2"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.TauxRéférence2
frmElpDisplay.fgData.Row = 16
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TauxMarge2"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.TauxMarge2

frmElpDisplay.fgData.Row = 17
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AmjEngagement"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.AmjEngagement
frmElpDisplay.fgData.Row = 18
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AmjDébut"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.AmjDébut
frmElpDisplay.fgData.Row = 19
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AmjFin"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.AmjFin
frmElpDisplay.fgData.Row = 20
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AmjEchéance1"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.AmjEchéance1
frmElpDisplay.fgData.Row = 21
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AmjEchéanceS"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.AmjEchéanceS

frmElpDisplay.fgData.Row = 22
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PréavisNbj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.PréavisNbj
frmElpDisplay.fgData.Row = 23
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Périodicité"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Périodicité
frmElpDisplay.fgData.Row = 24
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "PériodeNb"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.PériodeNb
frmElpDisplay.fgData.Row = 25
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "IPA"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.IPA
frmElpDisplay.fgData.Row = 26
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "NbjBase"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.NbjBase
frmElpDisplay.fgData.Row = 27
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Devise3"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Devise3
frmElpDisplay.fgData.Row = 28
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Mensualité"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Mensualité
frmElpDisplay.fgData.Row = 29
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Frais1"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Frais1
frmElpDisplay.fgData.Row = 30
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Frais2"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Frais2
frmElpDisplay.fgData.Row = 31
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Frais3"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Frais3
    
frmElpDisplay.fgData.Row = 32
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EngagementCompte"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.EngagementCompte
frmElpDisplay.fgData.Row = 33
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EngagementCorrCompte"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.EngagementCorrCompte
frmElpDisplay.fgData.Row = 34
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EngagementCorrSwiftN"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.EngagementCorrSwiftN
frmElpDisplay.fgData.Row = 35
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EngagementCorrSwiftL"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.EngagementCorrSwiftL
frmElpDisplay.fgData.Row = 36
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EchéanceCompte"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.EchéanceCompte
frmElpDisplay.fgData.Row = 37
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EchéanceCorrCompte"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.EchéanceCorrCompte
frmElpDisplay.fgData.Row = 38
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EchéanceCorrSwiftN"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.EchéanceCorrSwiftN
frmElpDisplay.fgData.Row = 39
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EchéanceCorrSwiftL"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.EchéanceCorrSwiftL
frmElpDisplay.fgData.Row = 40
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "RéférenceInterne"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.RéférenceInterne
frmElpDisplay.fgData.Row = 41
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "RéférenceExterne"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.RéférenceExterne
frmElpDisplay.fgData.Row = 42
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "IdRéférenceLiée"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.IdRéférenceLiée
frmElpDisplay.fgData.Row = 43
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "optReprise"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.optReprise
frmElpDisplay.fgData.Row = 44
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TauxRéférenceInterne"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.TauxRéférenceInterne
frmElpDisplay.fgData.Row = 45
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TauxMargeInterne"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.TauxMargeInterne

frmElpDisplay.fgData.Row = 46
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Statut"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Statut
frmElpDisplay.fgData.Row = 47
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "StatutPlus"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.StatutPlus
frmElpDisplay.fgData.Row = 48
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Flag1"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Flag1
frmElpDisplay.fgData.Row = 49
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Flag2"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Flag2
frmElpDisplay.fgData.Row = 50
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Flag3"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.Flag3
frmElpDisplay.fgData.Row = 51
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ElpId"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.ElpId
frmElpDisplay.fgData.Row = 52
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ElpUpdate"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.ElpUpdate
frmElpDisplay.fgData.Row = 53
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ElpControl"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGOpe.ElpControl

frmElpDisplay.Show vbModal

End Sub

Public Sub srvGOpe_Load(recGOpeMin As typeGOpe, recGOpeMax As typeGOpe)
Dim mMethod As String

mMethod = Trim(recGOpeMin.Method) & "+"
arrGOpe_NBMax = 0
arrGOpe_Suite = True: arrGOpe_NB = 0
arrGOpe_NBMax = recGOpe_Block: ReDim arrGOpe(arrGOpe_NBMax)

arrGOpe(0) = recGOpeMax
arrGOpe_Suite = True
Do Until Not arrGOpe_Suite
    srvGOpe_Monitor recGOpeMin
    recGOpeMin = arrGOpe(arrGOpe_NB)
    recGOpeMin.Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Public Function srvGOpe_Find(lGope As typeGOpe)
'-----------------------------------------------------
Dim I As Integer, X As Variant

X = "?"

For I = 1 To arrGOpe_NB
    If lGope.IdRéférence = arrGOpe(I).IdRéférence Then
        lGope = arrGOpe(I)
        X = Null
        Exit For
'!!!!!!!========
    End If
Next I

If Not IsNull(X) Then
    lGope.Method = "SeekP0"
    X = srvGOpe_Seek(lGope)
End If

srvGOpe_Find = X

End Function




'-----------------------------------------------------
Public Function srvGOpe_Monitor(recGOpe As typeGOpe)
'-----------------------------------------------------
blnFR_Convert = False
arrGOpe_Suite = False
Select Case mId$(Trim(recGOpe.Method), 1, 4)
    Case "Seek"
                srvGOpe_Monitor = srvGOpe_Seek(recGOpe)
    Case "Snap"
              srvGOpe_Monitor = srvGOpe_Snap(recGOpe)
    Case Else
                recGOpe.Err = recGOpe.Method
                Call srvGOpe_Error(recGOpe)
                srvGOpe_Monitor = recGOpe.Err
End Select

End Function

'-----------------------------------------------------
Sub srvGOpe_Error(recGOpe As typeGOpe)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "GOpe" & Chr$(10) & Chr$(13)

Select Case mId$(recGOpe.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recGOpe.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : GOpes.bas  ( " _
                & Trim(recGOpe.obj) & " : " & Trim(recGOpe.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvGOpe_GetBuffer(recGOpe As typeGOpe)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvGOpe_GetBuffer = Null
recGOpe.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recGOpe.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recGOpe.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recGOpe.Err = Space$(10) Then
    recGOpe.IdRéférence = CLng(Val(mId$(MsgTxt, K + 1, 12)))
    recGOpe.Application = mId$(MsgTxt, K + 13, 5)
    recGOpe.Nature = mId$(MsgTxt, K + 18, 5)
    recGOpe.Devise1 = mId$(MsgTxt, K + 23, 3)
    recGOpe.Montant1 = CCur(Val(mId$(MsgTxt, K + 26, 17))) / 100
    
    recGOpe.TauxRéférence1 = mId$(MsgTxt, K + 43, 10)
    If Trim(recGOpe.TauxRéférence1) = "Montant" Then
        recGOpe.TauxMarge1 = CDbl(Val(mId$(MsgTxt, K + 53, 9))) / 100
    Else
        recGOpe.TauxMarge1 = CDbl(Val(mId$(MsgTxt, K + 53, 9))) / 1000000
    End If
    recGOpe.TauxActuariel1 = CDbl(Val(mId$(MsgTxt, K + 62, 9))) / 1000000
    recGOpe.TEG1 = CDbl(Val(mId$(MsgTxt, K + 71, 9))) / 1000000
    
    recGOpe.Devise2 = mId$(MsgTxt, K + 80, 3)
    recGOpe.Montant2 = CCur(Val(mId$(MsgTxt, K + 83, 17))) / 100
    recGOpe.TauxRéférence2 = mId$(MsgTxt, K + 100, 10)
    If Trim(recGOpe.TauxRéférence2) = "Montant" Then
        recGOpe.TauxMarge2 = CDbl(Val(mId$(MsgTxt, K + 110, 9))) / 100
    Else
        recGOpe.TauxMarge2 = CDbl(Val(mId$(MsgTxt, K + 110, 9))) / 1000000
    End If
    
    recGOpe.AmjEngagement = mId$(MsgTxt, K + 119, 8)
    recGOpe.AmjDébut = mId$(MsgTxt, K + 127, 8)
    recGOpe.AmjFin = mId$(MsgTxt, K + 135, 8)
    recGOpe.AmjEchéance1 = mId$(MsgTxt, K + 143, 8)
    recGOpe.AmjEchéanceS = mId$(MsgTxt, K + 151, 1)
    recGOpe.PréavisNbj = CInt(Val(mId$(MsgTxt, K + 152, 3)))
    recGOpe.Périodicité = mId$(MsgTxt, K + 155, 1)
    recGOpe.PériodeNb = CInt(Val(mId$(MsgTxt, K + 156, 3)))
    recGOpe.IPA = mId$(MsgTxt, K + 159, 1)
    recGOpe.NbjBase = mId$(MsgTxt, K + 160, 1)
    
    recGOpe.Devise3 = mId$(MsgTxt, K + 161, 3)
    recGOpe.Mensualité = CCur(Val(mId$(MsgTxt, K + 164, 17))) / 100
    recGOpe.Frais1 = CCur(Val(mId$(MsgTxt, K + 181, 15))) / 100
    recGOpe.Frais2 = CCur(Val(mId$(MsgTxt, K + 196, 15))) / 100
    recGOpe.Frais3 = CCur(Val(mId$(MsgTxt, K + 211, 15))) / 100
    
    recGOpe.EngagementCompte = mId$(MsgTxt, K + 226, 11)
    recGOpe.EngagementCorrCompte = mId$(MsgTxt, K + 237, 11)
    recGOpe.EngagementCorrSwiftN = mId$(MsgTxt, K + 248, 11)
    recGOpe.EngagementCorrSwiftL = mId$(MsgTxt, K + 259, 11)
    
    recGOpe.EchéanceCompte = mId$(MsgTxt, K + 270, 11)
    recGOpe.EchéanceCorrCompte = mId$(MsgTxt, K + 281, 11)
    recGOpe.EchéanceCorrSwiftN = mId$(MsgTxt, K + 292, 11)
    recGOpe.EchéanceCorrSwiftL = mId$(MsgTxt, K + 303, 11)
    
    recGOpe.RéférenceInterne = mId$(MsgTxt, K + 314, 16)
    recGOpe.RéférenceExterne = mId$(MsgTxt, K + 330, 16)
    recGOpe.IdRéférenceLiée = CLng(Val(mId$(MsgTxt, K + 346, 12)))
    recGOpe.optReprise = mId$(MsgTxt, K + 358, 1)
    
    recGOpe.TauxRéférenceInterne = mId$(MsgTxt, K + 359, 10)
    If Trim(recGOpe.TauxRéférenceInterne) = "Montant" Then
        recGOpe.TauxMargeInterne = CDbl(Val(mId$(MsgTxt, K + 369, 9))) / 100
    Else
        recGOpe.TauxMargeInterne = CDbl(Val(mId$(MsgTxt, K + 369, 9))) / 1000000
    End If

    recGOpe.Statut = mId$(MsgTxt, K + 378, 1)
    recGOpe.StatutPlus = mId$(MsgTxt, K + 379, 2)
    recGOpe.Flag1 = mId$(MsgTxt, K + 381, 1)
    recGOpe.Flag2 = mId$(MsgTxt, K + 382, 1)
    recGOpe.Flag3 = mId$(MsgTxt, K + 383, 1)
    recGOpe.ElpId = CLng(Val(mId$(MsgTxt, K + 384, 12)))
    recGOpe.ElpUpdate = CInt(Val(mId$(MsgTxt, K + 396, 3)))
    recGOpe.ElpControl = mId$(MsgTxt, K + 399, 10)

Else
    srvGOpe_GetBuffer = recGOpe.Err
End If

MsgTxtIndex = MsgTxtIndex + recGOpeLen

End Function

'---------------------------------------------------------
Public Sub srvGOpe_PutBuffer(recGOpe As typeGOpe)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recGOpe.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recGOpe.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Mid$(MsgTxt, K + 1, 12) = Format$(recGOpe.IdRéférence, "000000000000")
Mid$(MsgTxt, K + 13, 5) = recGOpe.Application
Mid$(MsgTxt, K + 18, 5) = recGOpe.Nature

Mid$(MsgTxt, K + 23, 3) = recGOpe.Devise1
Mid$(MsgTxt, K + 26, 17) = Format$(recGOpe.Montant1 * 100, "00000000000000000")
Mid$(MsgTxt, K + 43, 10) = recGOpe.TauxRéférence1
If Trim(recGOpe.TauxRéférence1) = "Montant" Then
    Mid$(MsgTxt, K + 53, 9) = Format$(recGOpe.TauxMarge1 * 100, "000000000")
Else
    Mid$(MsgTxt, K + 53, 9) = Format$(recGOpe.TauxMarge1 * 1000000, "000000000")
End If
Mid$(MsgTxt, K + 62, 9) = Format$(recGOpe.TauxActuariel1 * 1000000, "000000000")
Mid$(MsgTxt, K + 71, 9) = Format$(recGOpe.TEG1 * 1000000, "000000000")
    
    
Mid$(MsgTxt, K + 80, 3) = recGOpe.Devise2
Mid$(MsgTxt, K + 83, 17) = Format$(recGOpe.Montant2 * 100, "00000000000000000")
Mid$(MsgTxt, K + 100, 10) = recGOpe.TauxRéférence2
If Trim(recGOpe.TauxRéférence1) = "Montant" Then
    Mid$(MsgTxt, K + 110, 9) = Format$(recGOpe.TauxMarge2 * 100, "000000000")
Else
    Mid$(MsgTxt, K + 110, 9) = Format$(recGOpe.TauxMarge2 * 1000000, "000000000")
End If
    
Mid$(MsgTxt, K + 119, 8) = recGOpe.AmjEngagement
Mid$(MsgTxt, K + 127, 8) = recGOpe.AmjDébut
Mid$(MsgTxt, K + 135, 8) = recGOpe.AmjFin
Mid$(MsgTxt, K + 143, 8) = recGOpe.AmjEchéance1
Mid$(MsgTxt, K + 151, 1) = recGOpe.AmjEchéanceS
Mid$(MsgTxt, K + 152, 3) = Format$(recGOpe.PréavisNbj, "000")
Mid$(MsgTxt, K + 155, 1) = recGOpe.Périodicité
Mid$(MsgTxt, K + 156, 3) = Format$(recGOpe.PériodeNb, "000")
Mid$(MsgTxt, K + 159, 1) = recGOpe.IPA
Mid$(MsgTxt, K + 160, 1) = recGOpe.NbjBase
    
Mid$(MsgTxt, K + 161, 3) = recGOpe.Devise3
Mid$(MsgTxt, K + 164, 17) = Format$(recGOpe.Mensualité * 100, "00000000000000000")
Mid$(MsgTxt, K + 181, 15) = Format$(recGOpe.Frais1 * 100, "000000000000000")
Mid$(MsgTxt, K + 196, 15) = Format$(recGOpe.Frais2 * 100, "000000000000000")
Mid$(MsgTxt, K + 211, 15) = Format$(recGOpe.Frais3 * 100, "000000000000000")


Mid$(MsgTxt, K + 226, 11) = recGOpe.EngagementCompte
Mid$(MsgTxt, K + 237, 11) = recGOpe.EngagementCorrCompte
Mid$(MsgTxt, K + 248, 11) = recGOpe.EngagementCorrSwiftN
Mid$(MsgTxt, K + 259, 11) = recGOpe.EngagementCorrSwiftL
    
Mid$(MsgTxt, K + 270, 11) = recGOpe.EchéanceCompte
Mid$(MsgTxt, K + 281, 11) = recGOpe.EchéanceCorrCompte
Mid$(MsgTxt, K + 292, 11) = recGOpe.EchéanceCorrSwiftN
Mid$(MsgTxt, K + 303, 11) = recGOpe.EchéanceCorrSwiftL
    
Mid$(MsgTxt, K + 314, 16) = recGOpe.RéférenceInterne
Mid$(MsgTxt, K + 330, 16) = recGOpe.RéférenceExterne
Mid$(MsgTxt, K + 346, 12) = Format$(recGOpe.IdRéférenceLiée, "000000000000")
Mid$(MsgTxt, K + 358, 1) = recGOpe.optReprise
    
Mid$(MsgTxt, K + 359, 10) = recGOpe.TauxRéférenceInterne
If Trim(recGOpe.TauxRéférenceInterne) = "Montant" Then
    Mid$(MsgTxt, K + 369, 9) = Format$(recGOpe.TauxMargeInterne * 100, "000000000")
Else
    Mid$(MsgTxt, K + 369, 9) = Format$(recGOpe.TauxMargeInterne * 1000000, "000000000")
End If
    
    

Mid$(MsgTxt, K + 378, 1) = recGOpe.Statut
Mid$(MsgTxt, K + 379, 2) = recGOpe.StatutPlus
Mid$(MsgTxt, K + 381, 1) = recGOpe.Flag1
Mid$(MsgTxt, K + 382, 1) = recGOpe.Flag2
Mid$(MsgTxt, K + 383, 1) = recGOpe.Flag3
Mid$(MsgTxt, K + 384, 12) = Format$(recGOpe.ElpId, "000000000000")
Mid$(MsgTxt, K + 396, 3) = Format$(recGOpe.ElpUpdate, "000")
Mid$(MsgTxt, K + 399, 10) = recGOpe.ElpControl


MsgTxtLen = MsgTxtLen + recGOpeLen
End Sub



'---------------------------------------------------------
Private Function srvGOpe_Seek(recGOpe As typeGOpe)
'---------------------------------------------------------

srvGOpe_Seek = "?"
MsgTxtLen = 0
Call srvGOpe_PutBuffer(recGOpe)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvGOpe_GetBuffer(recGOpe)) Then
            srvGOpe_Seek = Null
        Else
            Call srvGOpe_Error(recGOpe)
        End If
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvGOpe_Snap(recGOpe As typeGOpe)
'---------------------------------------------------------
srvGOpe_Snap = "?"
MsgTxtLen = 0
Call srvGOpe_PutBuffer(recGOpe)
Call srvGOpe_PutBuffer(arrGOpe(0))
If IsNull(SndRcv()) Then
    srvGOpe_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvGOpe_GetBuffer(recGOpe)) Then
            Call arrGOpe_AddItem(recGOpe)
            arrGOpe_Suite = True
        Else
            arrGOpe_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'-----------------------------------------------------
Function srvGOpe_Update(recGOpe As typeGOpe)
'-----------------------------------------------------

If blnMsgTxt_Concat_Transaction Then
    Call srvGOpe_PutBuffer(recGOpe)
    srvGOpe_Update = Null
    Exit Function
End If
    
srvGOpe_Update = "?"

MsgTxtLen = 0
Call srvGOpe_PutBuffer(recGOpe)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvGOpe_GetBuffer(recGOpe)) Then
        Call srvGOpe_Error(recGOpe)
        srvGOpe_Update = recGOpe.Err
        Exit Function
    Else
        srvGOpe_Update = Null
    End If
Else
    recGOpe.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recGOpe_Init(recGOpe As typeGOpe)
'---------------------------------------------------------
MsgTxt = Space$(recGOpeLen)
MsgTxtIndex = 0
Call srvGOpe_GetBuffer(recGOpe)
recGOpe.obj = "SRVGOPE    "
End Sub

'---------------------------------------------------------
Public Sub arrGOpe_AddItem(recGOpe As typeGOpe)
'---------------------------------------------------------
          
arrGOpe_NB = arrGOpe_NB + 1
    
If arrGOpe_NB > arrGOpe_NBMax Then
    arrGOpe_NBMax = arrGOpe_NBMax + 10
    ReDim Preserve arrGOpe(arrGOpe_NBMax)
End If
            
arrGOpe(arrGOpe_NB) = recGOpe
End Sub



Public Function fctGOpe_PériodeSuivante(mGOpe As typeGOpe, xAmjDébut As String, xAmjFin As String)
Dim wAmj As String

fctGOpe_PériodeSuivante = Null

wAmj = xAmjFin
xAmjDébut = dateElp("Jour", 1, wAmj)

Select Case mGOpe.Périodicité
    Case "M": xAmjFin = dateElp("MoisAdd", 1, wAmj)
    Case "T": xAmjFin = dateElp("MoisAdd", 3, wAmj)
    Case "S": xAmjFin = dateElp("MoisAdd", 6, wAmj)
    Case "A": xAmjFin = dateElp("MoisAdd", 12, wAmj)
    Case Else: fctGOpe_PériodeSuivante = "? Périodicité": Exit Function
End Select

If mGOpe.AmjEchéanceS = "M" Then xAmjFin = dateFinDeMois(xAmjFin)

End Function
Public Function fctGOpe_AmjFinPrécédente(mGOpe As typeGOpe, mAmjFin As String)

fctGOpe_AmjFinPrécédente = "? Erreur"

Select Case mGOpe.Périodicité
    Case "M": mAmjFin = dateElp("MoisAdd", -1, mAmjFin)
    Case "T": mAmjFin = dateElp("MoisAdd", -3, mAmjFin)
    Case "S": mAmjFin = dateElp("MoisAdd", -6, mAmjFin)
    Case "A": mAmjFin = dateElp("MoisAdd", -12, mAmjFin)
    Case Else: fctGOpe_AmjFinPrécédente = "? Périodicité : " & mGOpe.Périodicité: Exit Function
End Select

If mGOpe.AmjEchéanceS = "M" Then mAmjFin = dateFinDeMois(mAmjFin)

fctGOpe_AmjFinPrécédente = Null

End Function

Public Function fctGOpe_Mensualité(mGOpe As typeGOpe, mCV As typeCV, xMensualité As Currency, xTaux As Double, xTauxActuariel As Double, xTEG As Double)
Dim curX As Currency, Nb As Integer
fctGOpe_Mensualité = "? Erreur"

xMensualité = 0: xTaux = 0
If Trim(mGOpe.TauxRéférence1) <> "" Then fctGOpe_Mensualité = "? Taux indéxé non pgm": Exit Function
If mGOpe.AmjDébut > mGOpe.AmjFin Then fctGOpe_Mensualité = "? Date Début > fin": Exit Function

Select Case mGOpe.Périodicité
    Case "M": xTaux = mGOpe.TauxMarge1 / 1200: Nb = 12
    Case "T": xTaux = mGOpe.TauxMarge1 / 400: Nb = 4
    Case "S": xTaux = mGOpe.TauxMarge1 / 200: Nb = 2
    Case "A": xTaux = mGOpe.TauxMarge1 / 100: Nb = 1
    Case Else: fctGOpe_Mensualité = "? Périodicité": Exit Function
End Select
If xTaux = 0 Then fctGOpe_Mensualité = "? Taux nul": Exit Function

curX = (mGOpe.Montant1 * xTaux) / (1 - (1 + xTaux) ^ (-mGOpe.PériodeNb))
xMensualité = Round(curX, mCV.maxD)

xTauxActuariel = (1 + xTaux) ^ Nb - 1
xTauxActuariel = Round(xTauxActuariel * 100, 5)

xTEG = xTaux * Nb
xTEG = Round(xTEG * 100, 5)

fctGOpe_Mensualité = Null

End Function


Public Function fctGOpe_Compare(recGOpe As typeGOpe, mGOpe As typeGOpe)
fctGOpe_Compare = Null
If recGOpe.IdRéférence <> mGOpe.IdRéférence Then fctGOpe_Compare = "IdRéférence": Exit Function
If recGOpe.Application <> mGOpe.Application Then fctGOpe_Compare = "Service": Exit Function
If recGOpe.Nature <> mGOpe.Nature Then fctGOpe_Compare = "Nature": Exit Function

If recGOpe.Devise1 <> mGOpe.Devise1 Then fctGOpe_Compare = "Devise1": Exit Function
If recGOpe.Montant1 <> mGOpe.Montant1 Then fctGOpe_Compare = "Montant1": Exit Function
If recGOpe.TauxRéférence1 <> mGOpe.TauxRéférence1 Then fctGOpe_Compare = "TauxRéférence1": Exit Function
If recGOpe.TauxMarge1 <> mGOpe.TauxMarge1 Then fctGOpe_Compare = "TauxMarge1": Exit Function
If recGOpe.TauxActuariel1 <> mGOpe.TauxActuariel1 Then fctGOpe_Compare = "TauxActuariel1": Exit Function
If recGOpe.TEG1 <> mGOpe.TEG1 Then fctGOpe_Compare = "TEG1": Exit Function

If recGOpe.Devise2 <> mGOpe.Devise2 Then fctGOpe_Compare = "Devise2": Exit Function
If recGOpe.Montant2 <> mGOpe.Montant2 Then fctGOpe_Compare = "Montant2": Exit Function
If recGOpe.TauxRéférence2 <> mGOpe.TauxRéférence2 Then fctGOpe_Compare = "TauxRéférence2": Exit Function
If recGOpe.TauxMarge2 <> mGOpe.TauxMarge2 Then fctGOpe_Compare = "TauxMarge2": Exit Function

If recGOpe.AmjDébut <> mGOpe.AmjDébut Then fctGOpe_Compare = "AmjDébut": Exit Function
If recGOpe.AmjFin <> mGOpe.AmjFin Then fctGOpe_Compare = "AmjFin": Exit Function
If recGOpe.AmjEchéance1 <> mGOpe.AmjEchéance1 Then fctGOpe_Compare = "AmjEchéance1": Exit Function
If recGOpe.AmjEchéanceS <> mGOpe.AmjEchéanceS Then fctGOpe_Compare = "AmjEchéanceS": Exit Function
If recGOpe.PréavisNbj <> mGOpe.PréavisNbj Then fctGOpe_Compare = "PréavisNbj": Exit Function
If recGOpe.Périodicité <> mGOpe.Périodicité Then fctGOpe_Compare = "Périodicité": Exit Function
If recGOpe.PériodeNb <> mGOpe.PériodeNb Then fctGOpe_Compare = "PériodeNb": Exit Function
If recGOpe.IPA <> mGOpe.IPA Then fctGOpe_Compare = "IPA": Exit Function
If recGOpe.NbjBase <> mGOpe.NbjBase Then fctGOpe_Compare = "NbjBase": Exit Function

If recGOpe.Devise3 <> mGOpe.Devise3 Then fctGOpe_Compare = "Devise3": Exit Function
If recGOpe.Mensualité <> mGOpe.Mensualité Then fctGOpe_Compare = "Mensualité": Exit Function
If recGOpe.Frais1 <> mGOpe.Frais1 Then fctGOpe_Compare = "Frais1": Exit Function
If recGOpe.Frais2 <> mGOpe.Frais2 Then fctGOpe_Compare = "Frais2": Exit Function
If recGOpe.Frais3 <> mGOpe.Frais3 Then fctGOpe_Compare = "Frais3": Exit Function

If recGOpe.EngagementCompte <> mGOpe.EngagementCompte Then fctGOpe_Compare = "EngagementCompte": Exit Function
If recGOpe.EngagementCorrCompte <> mGOpe.EngagementCorrCompte Then fctGOpe_Compare = "EngagementCorrCompte": Exit Function
If recGOpe.EngagementCorrSwiftN <> mGOpe.EngagementCorrSwiftN Then fctGOpe_Compare = "EngagementCorrSwiftN": Exit Function
If recGOpe.EngagementCorrSwiftL <> mGOpe.EngagementCorrSwiftL Then fctGOpe_Compare = "EngagementCorrSwiftL": Exit Function

If recGOpe.EchéanceCompte <> mGOpe.EchéanceCompte Then fctGOpe_Compare = "EchéanceCompte": Exit Function
If recGOpe.EchéanceCorrCompte <> mGOpe.EchéanceCorrCompte Then fctGOpe_Compare = "EchéanceCorrCompte": Exit Function
If recGOpe.EchéanceCorrSwiftN <> mGOpe.EchéanceCorrSwiftN Then fctGOpe_Compare = "EchéanceCorrSwiftN": Exit Function
If recGOpe.EchéanceCorrSwiftL <> mGOpe.EchéanceCorrSwiftL Then fctGOpe_Compare = "EchéanceCorrSwiftN": Exit Function

If recGOpe.RéférenceInterne <> mGOpe.RéférenceInterne Then fctGOpe_Compare = "RéférenceInterne": Exit Function
If recGOpe.RéférenceExterne <> mGOpe.RéférenceExterne Then fctGOpe_Compare = "RéférenceExterne ": Exit Function
If recGOpe.IdRéférenceLiée <> mGOpe.IdRéférenceLiée Then fctGOpe_Compare = "IdRéférenceLiée": Exit Function
If recGOpe.optReprise <> mGOpe.optReprise Then fctGOpe_Compare = "optReprise": Exit Function
If recGOpe.TauxRéférenceInterne <> mGOpe.TauxRéférenceInterne Then fctGOpe_Compare = "TauxRéférenceInterne": Exit Function
If recGOpe.TauxMargeInterne <> mGOpe.TauxMargeInterne Then fctGOpe_Compare = "TauxMargeInterne": Exit Function

If recGOpe.Statut <> mGOpe.Statut Then fctGOpe_Compare = "Statut": Exit Function
If recGOpe.StatutPlus <> mGOpe.StatutPlus Then fctGOpe_Compare = "StatutPlus": Exit Function
If recGOpe.Flag1 <> mGOpe.Flag1 Then fctGOpe_Compare = "Flag1": Exit Function
If recGOpe.Flag2 <> mGOpe.Flag2 Then fctGOpe_Compare = "Flag2": Exit Function
If recGOpe.Flag3 <> mGOpe.Flag3 Then fctGOpe_Compare = "Flag3": Exit Function

If recGOpe.ElpId <> mGOpe.ElpId Then fctGOpe_Compare = "ElpId": Exit Function
If recGOpe.ElpUpdate <> mGOpe.ElpUpdate Then fctGOpe_Compare = "ElpUpdate": Exit Function
If recGOpe.ElpControl <> mGOpe.ElpControl Then fctGOpe_Compare = "ElpControl": Exit Function
End Function
