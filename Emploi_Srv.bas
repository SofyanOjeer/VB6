Attribute VB_Name = "srvEmploi"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recEmploiLen = 193 ' 34 + 159

Type typeEmploi
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Société                 As String * 3
    Agence                  As String * 3
    Devise                  As String * 3           ' Num 3
    Compte                  As String * 11
    
    Capital                 As Currency
    Intérêts                As Currency
    Taux                    As Double
    NbjBase                 As String * 1
    Type                    As String * 1
    NbjCouru                As Long
    Si4028                  As String * 1
    
    AmjDépart               As String * 8
    AmjEchéance             As String * 8
    TagEchéance             As String * 1
    Intitulé                As String * 40
    Intitulé2               As String * 40

End Type
    
Public arrEmploi() As typeEmploi
Public arrEmploiNb As Integer
Public arrEmploiNbMax As Integer
Public arrEmploiIndex As Integer
Public arrEmploiSuite As Boolean

'-----------------------------------------------------
Public Function srvEmploi_Monitor(recEmploi As typeEmploi)
'-----------------------------------------------------

arrEmploiSuite = False
Select Case Trim(recEmploi.Method)
    Case "SeekLE"
                srvEmploi_Monitor = srvEmploi_Seek(recEmploi)
    Case "SnapLE", "SnapLE+", "PrevLE", "PrevLE+"
              srvEmploi_Monitor = srvEmploi_Snap(recEmploi)
    Case Else
                recEmploi.Err = recEmploi.Method
                Call srvEmploi_error(recEmploi)
                srvEmploi_Monitor = recEmploi.Err
End Select

End Function

'-----------------------------------------------------
Sub srvEmploi_error(recEmploi As typeEmploi)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "emploi" & Chr$(10) & Chr$(13)

Select Case mId$(recEmploi.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recEmploi.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : emplois.bas  ( " _
                & Trim(recEmploi.obj) & " : " & Trim(recEmploi.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvEmploi_GetBuffer(recEmploi As typeEmploi)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvEmploi_GetBuffer = Null
recEmploi.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recEmploi.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recEmploi.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recEmploi.Err = Space$(10) Then
    recEmploi.Société = mId$(MsgTxt, K + 1, 3)
    recEmploi.Agence = mId$(MsgTxt, K + 4, 3)
    recEmploi.Devise = mId$(MsgTxt, K + 7, 3)
    recEmploi.Compte = mId$(MsgTxt, K + 10, 11)
    recEmploi.Capital = CDbl(Val(mId$(MsgTxt, K + 21, 13)) / 100)
    recEmploi.Intérêts = CDbl(Val(mId$(MsgTxt, K + 34, 13)) / 100)
    recEmploi.Taux = CDbl(Val(mId$(MsgTxt, K + 47, 8)) / 1000000)
    recEmploi.NbjBase = mId$(MsgTxt, K + 55, 1)
    recEmploi.Type = mId$(MsgTxt, K + 56, 1)
    recEmploi.NbjCouru = CLng(Val(mId$(MsgTxt, K + 57, 5)))
    recEmploi.Si4028 = mId$(MsgTxt, K + 62, 1)
    recEmploi.AmjDépart = mId$(MsgTxt, K + 63, 8)
    recEmploi.AmjEchéance = mId$(MsgTxt, K + 71, 8)
    recEmploi.TagEchéance = mId$(MsgTxt, K + 79, 1)
    recEmploi.Intitulé = mId$(MsgTxt, K + 80, 40)
    recEmploi.Intitulé2 = mId$(MsgTxt, K + 120, 40)

Else
    srvEmploi_GetBuffer = recEmploi.Err
End If

MsgTxtIndex = MsgTxtIndex + recEmploiLen

End Function

'---------------------------------------------------------
Private Sub srvEmploi_PutBuffer(recEmploi As typeEmploi)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recEmploi.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recEmploi.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 3) = recEmploi.Société
Mid$(MsgTxt, K + 4, 3) = recEmploi.Agence
Mid$(MsgTxt, K + 7, 3) = recEmploi.Devise
Mid$(MsgTxt, K + 10, 11) = recEmploi.Compte
Mid$(MsgTxt, K + 71, 8) = recEmploi.AmjEchéance
Mid$(MsgTxt, K + 80, 40) = recEmploi.Intitulé

MsgTxtLen = MsgTxtLen + recEmploiLen
End Sub



'---------------------------------------------------------
Private Function srvEmploi_Seek(recEmploi As typeEmploi)
'---------------------------------------------------------

srvEmploi_Seek = "?"
MsgTxtLen = 0
Call srvEmploi_PutBuffer(recEmploi)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvEmploi_GetBuffer(recEmploi)) Then
        srvEmploi_Seek = Null
    Else
        Call srvEmploi_error(recEmploi)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvEmploi_Snap(recEmploi As typeEmploi)
'---------------------------------------------------------
srvEmploi_Snap = "?"
MsgTxtLen = 0
Call srvEmploi_PutBuffer(recEmploi)
Call srvEmploi_PutBuffer(arrEmploi(0))
If IsNull(SndRcv()) Then
    srvEmploi_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvEmploi_GetBuffer(recEmploi)) Then
            Call arrEmploi_AddItem(recEmploi)
            arrEmploiSuite = True
        Else
            arrEmploiSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recEmploi_Init(recEmploi As typeEmploi)
'---------------------------------------------------------
 MsgTxt = Space$(recEmploiLen)
 MsgTxtIndex = 0
 Call srvEmploi_GetBuffer(recEmploi)
 recEmploi.obj = "SRVEMPLOI   "
End Sub

'---------------------------------------------------------
Public Sub arrEmploi_AddItem(recEmploi As typeEmploi)
'---------------------------------------------------------
          
arrEmploiNb = arrEmploiNb + 1
    
If arrEmploiNb > arrEmploiNbMax Then
    arrEmploiNbMax = arrEmploiNbMax + 10
    ReDim Preserve arrEmploi(arrEmploiNbMax)
End If
            
arrEmploi(arrEmploiNb) = recEmploi
End Sub
