Attribute VB_Name = "srvTOpe"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recTOpeLen = 358 ' 34 + 324

Type typeTOpe
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    IdRéférence             As Long
    Application             As String * 3
    Nature                  As String * 5
    Devise                  As String * 3
    Capital                 As Currency
    IPA                     As String * 1
    NbjBase                 As String * 1
    TauxRéférence           As String * 10
    TauxMarge               As Double
    TauxActuariel           As Double
    TEG                     As Double
      
    AmjDébut                As String * 8
    AmjFin                  As String * 8
    PréavisNbj              As Integer
    Périodicité             As String * 1
    PériodeNb               As Integer
    Mensualité              As Currency
    AmjEchéance1            As String * 8
    AmjEchéanceS            As String * 1
    Frais                   As Currency
    
    EngagementCompte        As String * 11
    EngagementCorrCompte    As String * 11
    EngagementCorrSwift     As String * 11
    EchéanceCompte          As String * 11
    EchéanceCorrCompte      As String * 11
    EchéanceCorrSwift       As String * 11
    RéférenceInterne        As String * 16
    RéférenceExterne        As String * 16
    IdRéférenceLiée         As Long
    optReprise              As String * 1

    MajUsr                  As String * 10
    MajAMJ                  As String * 8
    MajHMS                  As String * 6
    ValUsr                  As String * 10
    valAMJ                  As String * 8
    ValHMS                  As String * 6

    Statut                  As String * 1
    StatutPlus              As String * 2
   
    ElpId                   As Long
    ElpUpdate               As Integer
    ElpControl              As String * 10
    
End Type
    
Public arrTOpe() As typeTOpe
Public arrTOpe_NB As Integer
Public arrTOpe_NBMax As Integer
Public arrTOpe_Index As Integer
Public arrTOpe_Suite As Boolean
'-----------------------------------------------------
Public Function srvTope_Find(lTOpe As typeTOpe)
'-----------------------------------------------------
Dim I As Integer, x As Variant

x = "?"

For I = 1 To arrTOpe_NB
    If lTOpe.IdRéférence = arrTOpe(I).IdRéférence Then
        lTOpe = arrTOpe(I)
        x = Null
        Exit For
'!!!!!!!========
    End If
Next I

If Not IsNull(x) Then
    lTOpe.Method = "SeekP0"
    x = srvTOpe_Seek(lTOpe)
End If

srvTope_Find = x

End Function




'-----------------------------------------------------
Public Function srvTOpe_Monitor(recTope As typeTOpe)
'-----------------------------------------------------

arrTOpe_Suite = False
Select Case mId$(Trim(recTope.Method), 1, 4)
    Case "Seek"
                srvTOpe_Monitor = srvTOpe_Seek(recTope)
    Case "Snap"
              srvTOpe_Monitor = srvTOpe_Snap(recTope)
    Case Else
                recTope.Err = recTope.Method
                Call srvTOpe_Error(recTope)
                srvTOpe_Monitor = recTope.Err
End Select

End Function

'-----------------------------------------------------
Sub srvTOpe_Error(recTope As typeTOpe)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "TOpe" & Chr$(10) & Chr$(13)

Select Case mId$(recTope.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recTope.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : TOpes.bas  ( " _
                & Trim(recTope.obj) & " : " & Trim(recTope.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvTOpe_GetBuffer(recTope As typeTOpe)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvTOpe_GetBuffer = Null
recTope.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recTope.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recTope.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recTope.Err = Space$(10) Then
    recTope.IdRéférence = CLng(Val(mId$(MsgTxt, K + 1, 10)))
    recTope.Application = mId$(MsgTxt, K + 11, 3)
    recTope.Nature = mId$(MsgTxt, K + 14, 5)
    recTope.Devise = mId$(MsgTxt, K + 19, 3)
    recTope.Capital = CCur(Val(mId$(MsgTxt, K + 22, 17))) / 100
    recTope.IPA = mId$(MsgTxt, K + 39, 1)
    recTope.NbjBase = mId$(MsgTxt, K + 40, 1)
    recTope.TauxRéférence = mId$(MsgTxt, K + 41, 10)
    If Trim(recTope.TauxRéférence) = "Montant" Then
        recTope.TauxMarge = CDbl(Val(mId$(MsgTxt, K + 51, 9))) / 100
    Else
        recTope.TauxMarge = CDbl(Val(mId$(MsgTxt, K + 51, 9))) / 1000000
    End If
    recTope.TauxActuariel = CDbl(Val(mId$(MsgTxt, K + 60, 9))) / 1000000
    recTope.TEG = CDbl(Val(mId$(MsgTxt, K + 69, 9))) / 1000000
    
    recTope.AmjDébut = mId$(MsgTxt, K + 78, 8)
    recTope.AmjFin = mId$(MsgTxt, K + 86, 8)
    recTope.PréavisNbj = CInt(Val(mId$(MsgTxt, K + 94, 3)))
    recTope.Périodicité = mId$(MsgTxt, K + 97, 1)
    recTope.PériodeNb = CInt(Val(mId$(MsgTxt, K + 98, 3)))
    recTope.Mensualité = CCur(Val(mId$(MsgTxt, K + 101, 15))) / 100
    recTope.AmjEchéance1 = mId$(MsgTxt, K + 116, 8)
    recTope.AmjEchéanceS = mId$(MsgTxt, K + 124, 1)
    recTope.Frais = CCur(Val(mId$(MsgTxt, K + 125, 15))) / 100
    recTope.EngagementCompte = mId$(MsgTxt, K + 140, 11)
    recTope.EngagementCorrCompte = mId$(MsgTxt, K + 151, 11)
    recTope.EngagementCorrSwift = mId$(MsgTxt, K + 162, 11)
    recTope.EchéanceCompte = mId$(MsgTxt, K + 173, 11)
    recTope.EchéanceCorrCompte = mId$(MsgTxt, K + 184, 11)
    recTope.EchéanceCorrSwift = mId$(MsgTxt, K + 195, 11)
    recTope.RéférenceInterne = mId$(MsgTxt, K + 206, 16)
    recTope.RéférenceExterne = mId$(MsgTxt, K + 222, 16)
    recTope.IdRéférenceLiée = CLng(Val(mId$(MsgTxt, K + 238, 10)))
    recTope.optReprise = mId$(MsgTxt, K + 248, 1)
    
    recTope.MajUsr = mId$(MsgTxt, K + 249, 10)
    recTope.MajAMJ = Format$(Val(mId$(MsgTxt, K + 259, 8)), "00000000")
    recTope.MajHMS = Format$(Val(mId$(MsgTxt, K + 267, 6)), "000000")
    recTope.ValUsr = mId$(MsgTxt, K + 273, 10)
    recTope.valAMJ = Format$(Val(mId$(MsgTxt, K + 283, 8)), "00000000")
    recTope.ValHMS = Format$(Val(mId$(MsgTxt, K + 291, 6)), "000000")
   
    recTope.Statut = mId$(MsgTxt, K + 297, 1): 'jpl 2001.04.04 FR_ConvertEtoA_X recTope.Statut
    recTope.StatutPlus = mId$(MsgTxt, K + 298, 2)
    recTope.ElpId = CLng(Val(mId$(MsgTxt, K + 300, 12)))
    recTope.ElpUpdate = CInt(Val(mId$(MsgTxt, K + 312, 3)))
    recTope.ElpControl = mId$(MsgTxt, K + 315, 10)

Else
    srvTOpe_GetBuffer = recTope.Err
End If

MsgTxtIndex = MsgTxtIndex + recTOpeLen

End Function

'---------------------------------------------------------
Private Sub srvTOpe_PutBuffer(recTope As typeTOpe)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recTope.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recTope.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 10) = Format$(recTope.IdRéférence, "0000000000")
Mid$(MsgTxt, K + 11, 3) = recTope.Application
Mid$(MsgTxt, K + 14, 5) = recTope.Nature
Mid$(MsgTxt, K + 19, 3) = recTope.Devise
Mid$(MsgTxt, K + 22, 17) = Format$(recTope.Capital * 100, "00000000000000000")
Mid$(MsgTxt, K + 39, 1) = recTope.IPA
Mid$(MsgTxt, K + 40, 1) = recTope.NbjBase
Mid$(MsgTxt, K + 41, 10) = recTope.TauxRéférence
If Trim(recTope.TauxRéférence) = "Montant" Then
    Mid$(MsgTxt, K + 51, 9) = Format$(recTope.TauxMarge * 100, "000000000")
Else
    Mid$(MsgTxt, K + 51, 9) = Format$(recTope.TauxMarge * 1000000, "000000000")
End If
Mid$(MsgTxt, K + 60, 9) = Format$(recTope.TauxActuariel * 1000000, "000000000")
Mid$(MsgTxt, K + 69, 9) = Format$(recTope.TEG * 1000000, "000000000")
 
Mid$(MsgTxt, K + 78, 8) = recTope.AmjDébut
Mid$(MsgTxt, K + 86, 8) = recTope.AmjFin
Mid$(MsgTxt, K + 94, 3) = Format$(recTope.PréavisNbj, "000")
Mid$(MsgTxt, K + 97, 1) = recTope.Périodicité
Mid$(MsgTxt, K + 98, 3) = Format$(recTope.PériodeNb, "000")
Mid$(MsgTxt, K + 101, 15) = Format$(recTope.Mensualité * 100, "000000000000000")
Mid$(MsgTxt, K + 116, 8) = recTope.AmjEchéance1
Mid$(MsgTxt, K + 124, 1) = recTope.AmjEchéanceS
Mid$(MsgTxt, K + 125, 15) = Format$(recTope.Frais * 100, "000000000000000")
Mid$(MsgTxt, K + 140, 11) = recTope.EngagementCompte
Mid$(MsgTxt, K + 151, 11) = recTope.EngagementCorrCompte
Mid$(MsgTxt, K + 162, 11) = recTope.EngagementCorrSwift
Mid$(MsgTxt, K + 173, 11) = recTope.EchéanceCompte
Mid$(MsgTxt, K + 184, 11) = recTope.EchéanceCorrCompte
Mid$(MsgTxt, K + 195, 11) = recTope.EchéanceCorrSwift
Mid$(MsgTxt, K + 206, 16) = recTope.RéférenceInterne
Mid$(MsgTxt, K + 222, 16) = recTope.RéférenceExterne
Mid$(MsgTxt, K + 238, 10) = Format$(recTope.IdRéférenceLiée, "0000000000")
Mid$(MsgTxt, K + 248, 1) = recTope.optReprise
 
Mid$(MsgTxt, K + 249, 10) = recTope.MajUsr
Mid$(MsgTxt, K + 259, 8) = Format$(recTope.MajAMJ, "00000000")
Mid$(MsgTxt, K + 267, 6) = Format$(recTope.MajHMS, "000000")
Mid$(MsgTxt, K + 273, 10) = recTope.ValUsr
Mid$(MsgTxt, K + 283, 8) = Format$(recTope.valAMJ, "00000000")
Mid$(MsgTxt, K + 291, 6) = Format$(recTope.ValHMS, "000000")

Mid$(MsgTxt, K + 297, 1) = recTope.Statut
Mid$(MsgTxt, K + 298, 2) = recTope.StatutPlus
Mid$(MsgTxt, K + 300, 12) = Format$(recTope.ElpId, "000000000000")
Mid$(MsgTxt, K + 312, 3) = Format$(recTope.ElpUpdate, "000")
Mid$(MsgTxt, K + 315, 10) = recTope.ElpControl

MsgTxtLen = MsgTxtLen + recTOpeLen
End Sub



'---------------------------------------------------------
Private Function srvTOpe_Seek(recTope As typeTOpe)
'---------------------------------------------------------

srvTOpe_Seek = "?"
MsgTxtLen = 0
Call srvTOpe_PutBuffer(recTope)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvTOpe_GetBuffer(recTope)) Then
            srvTOpe_Seek = Null
        Else
            Call srvTOpe_Error(recTope)
        End If
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvTOpe_Snap(recTope As typeTOpe)
'---------------------------------------------------------
srvTOpe_Snap = "?"
MsgTxtLen = 0
Call srvTOpe_PutBuffer(recTope)
Call srvTOpe_PutBuffer(arrTOpe(0))
If IsNull(SndRcv()) Then
    srvTOpe_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvTOpe_GetBuffer(recTope)) Then
            Call arrTOpe_AddItem(recTope)
            arrTOpe_Suite = True
        Else
            arrTOpe_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'-----------------------------------------------------
Function srvTope_Update(recTope As typeTOpe)
'-----------------------------------------------------

srvTope_Update = "?"

MsgTxtLen = 0
Call srvTOpe_PutBuffer(recTope)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvTOpe_GetBuffer(recTope)) Then
        Call srvTOpe_Error(recTope)
        srvTope_Update = recTope.Err
        Exit Function
    Else
        srvTope_Update = Null
    End If
Else
    recTope.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recTOpe_Init(recTope As typeTOpe)
'---------------------------------------------------------
MsgTxt = Space$(recTOpeLen)
MsgTxtIndex = 0
Call srvTOpe_GetBuffer(recTope)
recTope.obj = "SRVTOPE    "
End Sub

'---------------------------------------------------------
Public Sub arrTOpe_AddItem(recTope As typeTOpe)
'---------------------------------------------------------
          
arrTOpe_NB = arrTOpe_NB + 1
    
If arrTOpe_NB > arrTOpe_NBMax Then
    arrTOpe_NBMax = arrTOpe_NBMax + 10
    ReDim Preserve arrTOpe(arrTOpe_NBMax)
End If
            
arrTOpe(arrTOpe_NB) = recTope
End Sub



Public Function fctTOpe_PériodeSuivante(mTope As typeTOpe, xAmjDébut As String, xAmjFin As String)
Dim wAmj As String

fctTOpe_PériodeSuivante = Null

wAmj = xAmjFin
xAmjDébut = dateElp("Jour", 1, wAmj)

Select Case mTope.Périodicité
    Case "M": xAmjFin = dateElp("MoisAdd", 1, wAmj)
    Case "T": xAmjFin = dateElp("MoisAdd", 3, wAmj)
    Case "S": xAmjFin = dateElp("MoisAdd", 6, wAmj)
    Case "A": xAmjFin = dateElp("MoisAdd", 12, wAmj)
    Case Else: fctTOpe_PériodeSuivante = "? Périodicité": Exit Function
End Select

If mTope.AmjEchéanceS = "M" Then xAmjFin = dateFinDeMois(xAmjFin)

End Function
Public Function fctTOpe_AmjFinPrécédente(mTope As typeTOpe, mAmjFin As String)

fctTOpe_AmjFinPrécédente = "? Erreur"

Select Case mTope.Périodicité
    Case "M": mAmjFin = dateElp("MoisAdd", -1, mAmjFin)
    Case "T": mAmjFin = dateElp("MoisAdd", -3, mAmjFin)
    Case "S": mAmjFin = dateElp("MoisAdd", -6, mAmjFin)
    Case "A": mAmjFin = dateElp("MoisAdd", -12, mAmjFin)
    Case Else: fctTOpe_AmjFinPrécédente = "? Périodicité : " & mTope.Périodicité: Exit Function
End Select

If mTope.AmjEchéanceS = "M" Then mAmjFin = dateFinDeMois(mAmjFin)

fctTOpe_AmjFinPrécédente = Null

End Function

Public Function fctTOpe_Intérêts(mTope As typeTOpe, mCV As typeCV, xIntérêts As Currency, xNbj As Long)
Dim wNbjBase As Double, curX As Currency, V1 As Variant, V2 As Variant

fctTOpe_Intérêts = "? Erreur"

xIntérêts = 0: xNbj = 0
If Trim(mTope.TauxRéférence) <> "" Then fctTOpe_Intérêts = "? Taux indéxé non pgm": Exit Function
If mTope.AmjDébut > mTope.AmjFin Then fctTOpe_Intérêts = "? Date Début > fin": Exit Function

Select Case mTope.NbjBase
    Case "0": wNbjBase = 36000
    Case "5": wNbjBase = 36500
    Case Else: fctTOpe_Intérêts = "? Base": Exit Function
End Select
V1 = CDate(dateImp_jjMoisAAAA(mTope.AmjDébut))
V2 = CDate(dateImp_jjMoisAAAA(mTope.AmjFin))
xNbj = DateDiff("d", V1, V2)

curX = mTope.Capital * mTope.TauxMarge * xNbj / wNbjBase

xIntérêts = Round(curX, mCV.maxD)
fctTOpe_Intérêts = Null

End Function

Public Function fctTOpe_Mensualité(mTope As typeTOpe, mCV As typeCV, xMensualité As Currency, xTaux As Double, xTauxActuariel As Double, xTEG As Double)
Dim curX As Currency, Nb As Integer
fctTOpe_Mensualité = "? Erreur"

xMensualité = 0: xTaux = 0
If Trim(mTope.TauxRéférence) <> "" Then fctTOpe_Mensualité = "? Taux indéxé non pgm": Exit Function
If mTope.AmjDébut > mTope.AmjFin Then fctTOpe_Mensualité = "? Date Début > fin": Exit Function

Select Case mTope.Périodicité
    Case "M": xTaux = mTope.TauxMarge / 1200: Nb = 12
    Case "T": xTaux = mTope.TauxMarge / 400: Nb = 4
    Case "S": xTaux = mTope.TauxMarge / 200: Nb = 2
    Case "A": xTaux = mTope.TauxMarge / 100: Nb = 1
    Case Else: fctTOpe_Mensualité = "? Périodicité": Exit Function
End Select
If xTaux = 0 Then fctTOpe_Mensualité = "? Taux nul": Exit Function

curX = (mTope.Capital * xTaux) / (1 - (1 + xTaux) ^ (-mTope.PériodeNb))
xMensualité = Round(curX, mCV.maxD)

xTauxActuariel = (1 + xTaux) ^ Nb - 1
xTauxActuariel = Round(xTauxActuariel * 100, 5)

xTEG = xTaux * Nb
xTEG = Round(xTEG * 100, 5)

fctTOpe_Mensualité = Null

End Function


Public Function fctTOpe_Compare(recTope As typeTOpe, mTope As typeTOpe)
fctTOpe_Compare = Null
If recTope.IdRéférence <> mTope.IdRéférence Then fctTOpe_Compare = "IdRéférence": Exit Function
If recTope.Application <> mTope.Application Then fctTOpe_Compare = "Service": Exit Function
If recTope.Nature <> mTope.Nature Then fctTOpe_Compare = "Nature": Exit Function
If recTope.Devise <> mTope.Devise Then fctTOpe_Compare = "Devise": Exit Function
If recTope.Capital <> mTope.Capital Then fctTOpe_Compare = "Capital": Exit Function
If recTope.IPA <> mTope.IPA Then fctTOpe_Compare = "IPA": Exit Function
If recTope.NbjBase <> mTope.NbjBase Then fctTOpe_Compare = "NbjBase": Exit Function
If recTope.TauxRéférence <> mTope.TauxRéférence Then fctTOpe_Compare = "TauxRéférence": Exit Function
If recTope.TauxMarge <> mTope.TauxMarge Then fctTOpe_Compare = "TauxMarge": Exit Function
If recTope.TauxActuariel <> mTope.TauxActuariel Then fctTOpe_Compare = "TauxActuariel": Exit Function
If recTope.TEG <> mTope.TEG Then fctTOpe_Compare = "TEG": Exit Function

If recTope.AmjDébut <> mTope.AmjDébut Then fctTOpe_Compare = "AmjDébut": Exit Function
If recTope.AmjFin <> mTope.AmjFin Then fctTOpe_Compare = "AmjFin": Exit Function
If recTope.PréavisNbj <> mTope.PréavisNbj Then fctTOpe_Compare = "PréavisNbj": Exit Function
If recTope.Périodicité <> mTope.Périodicité Then fctTOpe_Compare = "Périodicité": Exit Function
If recTope.PériodeNb <> mTope.PériodeNb Then fctTOpe_Compare = "PériodeNb": Exit Function
If recTope.Mensualité <> mTope.Mensualité Then fctTOpe_Compare = "Mensualité": Exit Function
If recTope.AmjEchéance1 <> mTope.AmjEchéance1 Then fctTOpe_Compare = "AmjEchéance1": Exit Function
If recTope.AmjEchéanceS <> mTope.AmjEchéanceS Then fctTOpe_Compare = "AmjEchéanceS": Exit Function
If recTope.Frais <> mTope.Frais Then fctTOpe_Compare = "Frais": Exit Function
If recTope.EngagementCompte <> mTope.EngagementCompte Then fctTOpe_Compare = "EngagementCompte": Exit Function
If recTope.EngagementCorrCompte <> mTope.EngagementCorrCompte Then fctTOpe_Compare = "EngagementCorrCompte": Exit Function
If recTope.EngagementCorrSwift <> mTope.EngagementCorrSwift Then fctTOpe_Compare = "EngagementCorrSwift": Exit Function
If recTope.EchéanceCompte <> mTope.EchéanceCompte Then fctTOpe_Compare = "EchéanceCompte": Exit Function
If recTope.EchéanceCorrCompte <> mTope.EchéanceCorrCompte Then fctTOpe_Compare = "EchéanceCorrCompte": Exit Function
If recTope.RéférenceInterne <> mTope.RéférenceInterne Then fctTOpe_Compare = "RéférenceInterne": Exit Function
If recTope.RéférenceExterne <> mTope.RéférenceExterne Then fctTOpe_Compare = "RéférenceExterne ": Exit Function
If recTope.IdRéférenceLiée <> mTope.IdRéférenceLiée Then fctTOpe_Compare = "IdRéférenceLiée": Exit Function
If recTope.optReprise <> mTope.optReprise Then fctTOpe_Compare = "optReprise": Exit Function

If recTope.MajUsr <> mTope.MajUsr Then fctTOpe_Compare = "MajUsr": Exit Function
If recTope.MajAMJ <> mTope.MajAMJ Then fctTOpe_Compare = "MajAMJ": Exit Function
If recTope.MajHMS <> mTope.MajHMS Then fctTOpe_Compare = "MajHMS": Exit Function
If recTope.ValUsr <> mTope.ValUsr Then fctTOpe_Compare = "ValUsr": Exit Function
If recTope.valAMJ <> mTope.valAMJ Then fctTOpe_Compare = "valAMJ": Exit Function
If recTope.ValHMS <> mTope.ValHMS Then fctTOpe_Compare = "ValHMS": Exit Function
If recTope.Statut <> mTope.Statut Then fctTOpe_Compare = "Statut": Exit Function
If recTope.StatutPlus <> mTope.StatutPlus Then fctTOpe_Compare = "StatutPlus": Exit Function
If recTope.ElpId <> mTope.ElpId Then fctTOpe_Compare = "ElpId": Exit Function
If recTope.ElpUpdate <> mTope.ElpUpdate Then fctTOpe_Compare = "ElpUpdate": Exit Function
If recTope.ElpControl <> mTope.ElpControl Then fctTOpe_Compare = "ElpControl": Exit Function
End Function
