Attribute VB_Name = "srvElpXXXXXX"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'Public Const recXXXXXXLen = *** ' 34 +

Type typeXXXXXX
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    
End Type
    
Public Sub srvXXXXXX_Load(lXXXXXX() As typeXXXXXX, lXXXXXX_Nb As Integer)
Dim mMethod As String, blnXXXXXX_Suite
Dim wNbMax As Integer
Dim wXXXXXX As typeXXXXXX

mMethod = Trim(lXXXXXX(0).Method) & "+"
blnXXXXXX_Suite = True: lXXXXXX_Nb = 0
wNbMax = recXXXXXX_Block + 2: ReDim Preserve arrXXXXXX(wNbMax)

wXXXXXX = lXXXXXX(1)
Do Until Not blnXXXXXX_Suite
    MsgTxtLen = 0
    Call srvXXXXXX_PutBuffer(wXXXXXX)
    Call srvXXXXXX_PutBuffer(lXXXXXX(0))
    If IsNull(SndRcv()) Then
        MsgTxtIndex = 0
        Do While MsgTxtIndex < MsgTxtLen
            If IsNull(srvXXXXXX_GetBuffer(wXXXXXX)) Then
            
                lXXXXXX_Nb = lXXXXXX_Nb + 1
                If lXXXXXX_Nb > wNbMax Then
                    wNbMax = wNbMax + recXXXXXX_Block
                    ReDim Preserve arrXXXXXX(lXXXXXX_NBMax)
                End If
            
                lXXXXXX(lXXXXXX_Nb) = wXXXXXX
                blnXXXXXX_Suite = True
            Else
                blnXXXXXX_Suite = False
                Exit Do
            End If
        Loop
    End If

    lXXXXXX(0).Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Public Function srvXXXXXX_Monitor(recXXXXXX As typeXXXXXX)
'-----------------------------------------------------

Select Case mId$(Trim(recXXXXXX.Method), 1, 4)
    Case "Seek"
                srvXXXXXX_Monitor = srvXXXXXX_Seek(recXXXXXX)
    Case Else
                recXXXXXX.Err = recXXXXXX.Method
                Call srvXXXXXX_Error(recXXXXXX)
                srvXXXXXX_Monitor = recXXXXXX.Err
End Select

End Function

'-----------------------------------------------------
Sub srvXXXXXX_Error(recXXXXXX As typeXXXXXX)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "XXXXXX" & Chr$(10) & Chr$(13)

Select Case mId$(recXXXXXX.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recXXXXXX.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : XXXXXXs.bas  ( " _
                & Trim(recXXXXXX.obj) & " : " & Trim(recXXXXXX.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvXXXXXX_GetBuffer(recXXXXXX As typeXXXXXX)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvXXXXXX_GetBuffer = Null
recXXXXXX.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recXXXXXX.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recXXXXXX.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recXXXXXX.Err = Space$(10) Then
    recXXXXXX.IdRéférence = CLng(Val(mId$(MsgTxt, K + 1, 10)))
    recXXXXXX.Application = mId$(MsgTxt, K + 11, 3)
    recXXXXXX.Nature = mId$(MsgTxt, K + 14, 5)
    recXXXXXX.Devise = mId$(MsgTxt, K + 19, 3)
    recXXXXXX.Capital = CCur(Val(mId$(MsgTxt, K + 22, 17))) / 100
    recXXXXXX.IPA = mId$(MsgTxt, K + 39, 1)
    recXXXXXX.NbjBase = mId$(MsgTxt, K + 40, 1)
    recXXXXXX.TauxRéférence = mId$(MsgTxt, K + 41, 10)
    If Trim(recXXXXXX.TauxRéférence) = "Montant" Then
        recXXXXXX.TauxMarge = CDbl(Val(mId$(MsgTxt, K + 51, 9))) / 100
    Else
        recXXXXXX.TauxMarge = CDbl(Val(mId$(MsgTxt, K + 51, 9))) / 1000000
    End If
    recXXXXXX.TauxActuariel = CDbl(Val(mId$(MsgTxt, K + 60, 9))) / 1000000
    recXXXXXX.TEG = CDbl(Val(mId$(MsgTxt, K + 69, 9))) / 1000000
    
    recXXXXXX.AmjDébut = mId$(MsgTxt, K + 78, 8)
    recXXXXXX.AmjFin = mId$(MsgTxt, K + 86, 8)
    recXXXXXX.PréavisNbj = CInt(Val(mId$(MsgTxt, K + 94, 3)))
    recXXXXXX.Périodicité = mId$(MsgTxt, K + 97, 1)
    recXXXXXX.PériodeNb = CInt(Val(mId$(MsgTxt, K + 98, 3)))
    recXXXXXX.Mensualité = CCur(Val(mId$(MsgTxt, K + 101, 15))) / 100
    recXXXXXX.AmjEchéance1 = mId$(MsgTxt, K + 116, 8)
    recXXXXXX.AmjEchéanceS = mId$(MsgTxt, K + 124, 1)
    recXXXXXX.Frais = CCur(Val(mId$(MsgTxt, K + 125, 15))) / 100
    recXXXXXX.EngagementCompte = mId$(MsgTxt, K + 140, 11)
    recXXXXXX.EngagementCorrCompte = mId$(MsgTxt, K + 151, 11)
    recXXXXXX.EngagementCorrSwift = mId$(MsgTxt, K + 162, 11)
    recXXXXXX.EchéanceCompte = mId$(MsgTxt, K + 173, 11)
    recXXXXXX.EchéanceCorrCompte = mId$(MsgTxt, K + 184, 11)
    recXXXXXX.EchéanceCorrSwift = mId$(MsgTxt, K + 195, 11)
    recXXXXXX.RéférenceInterne = mId$(MsgTxt, K + 206, 16)
    recXXXXXX.RéférenceExterne = mId$(MsgTxt, K + 222, 16)
    recXXXXXX.IdRéférenceLiée = CLng(Val(mId$(MsgTxt, K + 238, 10)))
    recXXXXXX.optReprise = mId$(MsgTxt, K + 248, 1)
    
    recXXXXXX.MajUsr = mId$(MsgTxt, K + 249, 10)
    recXXXXXX.MajAMJ = Format$(Val(mId$(MsgTxt, K + 259, 8)), "00000000")
    recXXXXXX.MajHMS = Format$(Val(mId$(MsgTxt, K + 267, 6)), "000000")
    recXXXXXX.ValUsr = mId$(MsgTxt, K + 273, 10)
    recXXXXXX.valAMJ = Format$(Val(mId$(MsgTxt, K + 283, 8)), "00000000")
    recXXXXXX.ValHMS = Format$(Val(mId$(MsgTxt, K + 291, 6)), "000000")
   
    recXXXXXX.Statut = mId$(MsgTxt, K + 297, 1): 'jpl 2001.04.04 FR_ConvertEtoA_X recXXXXXX.Statut
    recXXXXXX.StatutPlus = mId$(MsgTxt, K + 298, 2)
    recXXXXXX.ElpId = CLng(Val(mId$(MsgTxt, K + 300, 12)))
    recXXXXXX.ElpUpdate = CInt(Val(mId$(MsgTxt, K + 312, 3)))
    recXXXXXX.ElpControl = mId$(MsgTxt, K + 315, 10)

Else
    srvXXXXXX_GetBuffer = recXXXXXX.Err
End If

MsgTxtIndex = MsgTxtIndex + recXXXXXXLen

End Function

'---------------------------------------------------------
Private Sub srvXXXXXX_PutBuffer(recXXXXXX As typeXXXXXX)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recXXXXXX.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recXXXXXX.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34


MsgTxtLen = MsgTxtLen + recXXXXXXLen
End Sub



'---------------------------------------------------------
Private Function srvXXXXXX_Seek(recXXXXXX As typeXXXXXX)
'---------------------------------------------------------

srvXXXXXX_Seek = "?"
MsgTxtLen = 0
Call srvXXXXXX_PutBuffer(recXXXXXX)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvXXXXXX_GetBuffer(recXXXXXX)) Then
            srvXXXXXX_Seek = Null
        Else
            Call srvXXXXXX_Error(recXXXXXX)
        End If
    End If
End If

End Function

'-----------------------------------------------------
Function srvXXXXXX_Update(recXXXXXX As typeXXXXXX)
'-----------------------------------------------------

srvXXXXXX_Update = "?"

MsgTxtLen = 0
Call srvXXXXXX_PutBuffer(recXXXXXX)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvXXXXXX_GetBuffer(recXXXXXX)) Then
        Call srvXXXXXX_Error(recXXXXXX)
        srvXXXXXX_Update = recXXXXXX.Err
        Exit Function
    Else
        srvXXXXXX_Update = Null
    End If
Else
    recXXXXXX.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recXXXXXX_Init(recXXXXXX As typeXXXXXX)
'---------------------------------------------------------
MsgTxt = Space$(recXXXXXXLen)
MsgTxtIndex = 0
Call srvXXXXXX_GetBuffer(recXXXXXX)
recXXXXXX.obj = "SRVXXXXXX"
End Sub


Public Function fctXXXXXX_Compare(recXXXXXX As typeXXXXXX, mXXXXXX As typeXXXXXX)
fctXXXXXX_Compare = Null
'If recXXXXXX.IdRéférence <> mXXXXXX.IdRéférence Then fctXXXXXX_Compare = "IdRéférence": Exit Function
End Function


