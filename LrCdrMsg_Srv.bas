Attribute VB_Name = "srvLrCdrMsg"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recLrCdrMsgLen = 166 ' 34 + 132

Type typeLrCdrMsg
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Text                    As String * 132
    
  End Type
    
Public arrLrCdrMsgSuite As Boolean
Public arrLrCdrMsgNb As Integer
'-----------------------------------------------------
Public Function Monitor(recLrCdrMsg As typeLrCdrMsg)
'-----------------------------------------------------

arrLrCdrMsgSuite = False
Select Case Mid$(Trim(recLrCdrMsg.Method), 1, 4)
    Case "Snap"
              Monitor = Snap(recLrCdrMsg)
    Case Else
                recLrCdrMsg.Err = recLrCdrMsg.Method
                Call ErrorX(recLrCdrMsg)
                Monitor = recLrCdrMsg.Err
End Select

End Function

'-----------------------------------------------------
Sub ErrorX(recLrCdrMsg As typeLrCdrMsg)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "LrCdrMsg: "

Select Case Mid$(recLrCdrMsg.Err, 9, 2)
    Case "22"
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recLrCdrMsg.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvLrCdrMsg.bas  ( " _
                & Trim(recLrCdrMsg.obj) & " : " & Trim(recLrCdrMsg.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function GetBuffer(recLrCdrMsg As typeLrCdrMsg)
'---------------------------------------------------------
Dim K As Integer, I As Integer
GetBuffer = Null
recLrCdrMsg.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recLrCdrMsg.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recLrCdrMsg.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recLrCdrMsg.Err = Space$(10) Then
    recLrCdrMsg.Text = Mid$(MsgTxt, K + 1, 132)
Else
    GetBuffer = recLrCdrMsg.Err
End If

MsgTxtIndex = MsgTxtIndex + recLrCdrMsgLen

End Function

'---------------------------------------------------------
Private Sub PutBuffer(recLrCdrMsg As typeLrCdrMsg)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recLrCdrMsg.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recLrCdrMsg.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 132) = recLrCdrMsg.Text
MsgTxtLen = MsgTxtLen + recLrCdrMsgLen
End Sub



'---------------------------------------------------------
Private Function Snap(recLrCdrMsg As typeLrCdrMsg)
'---------------------------------------------------------
Dim I As Integer
Snap = "?"
MsgTxtLen = 0
Call PutBuffer(recLrCdrMsg)
'Call PutBuffer(arrLrCdrMsg(0))
If IsNull(SndRcv()) Then
    Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(GetBuffer(recLrCdrMsg)) Then
            arrLrCdrMsgNb = arrLrCdrMsgNb + 1
            Print #1, recLrCdrMsg.Text
            arrLrCdrMsgSuite = True
        Else
            arrLrCdrMsgSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub Init(recLrCdrMsg As typeLrCdrMsg)
'---------------------------------------------------------
MsgTxt = Space$(recLrCdrMsgLen)
MsgTxtIndex = 0
Call GetBuffer(recLrCdrMsg)
recLrCdrMsg.obj = "SRVLrCdrM"
End Sub




