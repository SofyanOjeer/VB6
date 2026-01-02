Attribute VB_Name = "srvLrBafiMsg"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recLrBafiMsgLen = 166 ' 34 + 132

Type typeLrBafiMsg
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Text                    As String * 132
    
  End Type
    
Public arrLrBafiMsgSuite As Boolean
Public arrLrBafiMsgNb As Integer
'-----------------------------------------------------
Public Function Monitor(recLrBafiMsg As typeLrBafiMsg)
'-----------------------------------------------------

arrLrBafiMsgSuite = False
Select Case Mid$(Trim(recLrBafiMsg.Method), 1, 4)
    Case "Snap"
              Monitor = Snap(recLrBafiMsg)
    Case Else
                recLrBafiMsg.Err = recLrBafiMsg.Method
                Call ErrorX(recLrBafiMsg)
                Monitor = recLrBafiMsg.Err
End Select

End Function

'-----------------------------------------------------
Sub ErrorX(recLrBafiMsg As typeLrBafiMsg)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "LrBafiMsg: "

Select Case Mid$(recLrBafiMsg.Err, 9, 2)
    Case "22"
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recLrBafiMsg.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvLrBafiMsg.bas  ( " _
                & Trim(recLrBafiMsg.obj) & " : " & Trim(recLrBafiMsg.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function GetBuffer(recLrBafiMsg As typeLrBafiMsg)
'---------------------------------------------------------
Dim K As Integer, I As Integer
GetBuffer = Null
recLrBafiMsg.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recLrBafiMsg.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recLrBafiMsg.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recLrBafiMsg.Err = Space$(10) Then
    recLrBafiMsg.Text = Mid$(MsgTxt, K + 1, 132)
Else
    GetBuffer = recLrBafiMsg.Err
End If

MsgTxtIndex = MsgTxtIndex + recLrBafiMsgLen

End Function

'---------------------------------------------------------
Private Sub PutBuffer(recLrBafiMsg As typeLrBafiMsg)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recLrBafiMsg.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recLrBafiMsg.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 132) = recLrBafiMsg.Text
MsgTxtLen = MsgTxtLen + recLrBafiMsgLen
End Sub



'---------------------------------------------------------
Private Function Snap(recLrBafiMsg As typeLrBafiMsg)
'---------------------------------------------------------
Dim I As Integer
Snap = "?"
MsgTxtLen = 0
Call PutBuffer(recLrBafiMsg)
'Call PutBuffer(arrLrBafiMsg(0))
If IsNull(SndRcv()) Then
    Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(GetBuffer(recLrBafiMsg)) Then
            arrLrBafiMsgNb = arrLrBafiMsgNb + 1
            Print #1, recLrBafiMsg.Text
            arrLrBafiMsgSuite = True
        Else
            arrLrBafiMsgSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub Init(recLrBafiMsg As typeLrBafiMsg)
'---------------------------------------------------------
MsgTxt = Space$(recLrBafiMsgLen)
MsgTxtIndex = 0
Call GetBuffer(recLrBafiMsg)
recLrBafiMsg.obj = "SRVLRBAFIM"
End Sub




