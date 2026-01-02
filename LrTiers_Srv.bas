Attribute VB_Name = "srvLrTiers"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recLrTiersLen = 419 ' 34 + 385

Type typeLrTiers
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Text                    As String * 385
    
  End Type
    
Public arrLrTiersSuite As Boolean
Public arrLrTiersNb As Integer
'-----------------------------------------------------
Public Function Monitor(recLrTiers As typeLrTiers)
'-----------------------------------------------------

arrLrTiersSuite = False
Select Case Mid$(Trim(recLrTiers.Method), 1, 4)
    Case "Snap"
              Monitor = Snap(recLrTiers)
    Case Else
                recLrTiers.Err = recLrTiers.Method
                Call ErrorX(recLrTiers)
                Monitor = recLrTiers.Err
End Select

End Function

'-----------------------------------------------------
Sub ErrorX(recLrTiers As typeLrTiers)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "LrTiers: "

Select Case Mid$(recLrTiers.Err, 9, 2)
    Case "22"
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recLrTiers.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvLrTiers.bas  ( " _
                & Trim(recLrTiers.obj) & " : " & Trim(recLrTiers.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function GetBuffer(recLrTiers As typeLrTiers)
'---------------------------------------------------------
Dim K As Integer, I As Integer
GetBuffer = Null
recLrTiers.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recLrTiers.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recLrTiers.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recLrTiers.Err = Space$(10) Then
    recLrTiers.Text = Mid$(MsgTxt, K + 1, 385)
Else
    GetBuffer = recLrTiers.Err
End If

MsgTxtIndex = MsgTxtIndex + recLrTiersLen

End Function

'---------------------------------------------------------
Private Sub PutBuffer(recLrTiers As typeLrTiers)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recLrTiers.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recLrTiers.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 385) = recLrTiers.Text
MsgTxtLen = MsgTxtLen + recLrTiersLen
End Sub



'---------------------------------------------------------
Private Function Snap(recLrTiers As typeLrTiers)
'---------------------------------------------------------
Dim I As Integer
Snap = "?"
MsgTxtLen = 0
Call PutBuffer(recLrTiers)
'Call PutBuffer(arrLrTiers(0))
If IsNull(SndRcv()) Then
    Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(GetBuffer(recLrTiers)) Then
            arrLrTiersNb = arrLrTiersNb + 1
            Print #1, recLrTiers.Text
            arrLrTiersSuite = True
        Else
            arrLrTiersSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub Init(recLrTiers As typeLrTiers)
'---------------------------------------------------------
MsgTxt = Space$(recLrTiersLen)
MsgTxtIndex = 0
Call GetBuffer(recLrTiers)
recLrTiers.obj = "SRVLRTIERS"
End Sub



