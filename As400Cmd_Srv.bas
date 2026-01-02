Attribute VB_Name = "srvAs400Cmd"
Option Explicit

Public Const recAs400CmdLen = 512         ' 34 + 478

Type typeAs400Cmd
    obj         As String * 12
    Method     As String * 12
    Err        As String * 10
    Text       As String * 478
 
End Type
    


'-----------------------------------------------------
Sub ErrorX(recAs400Cmd As typeAs400Cmd)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "AS400 Commande " & Chr$(10) & Chr$(13)

Select Case mId$(recAs400Cmd.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recAs400Cmd.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvAs400Cmd.bas  ( " _
                & Trim(recAs400Cmd.obj) & " : " & Trim(recAs400Cmd.Method) & " )"

End Sub
'---------------------------------------------------------
Public Function GetBuffer(recAs400Cmd As typeAs400Cmd)
'---------------------------------------------------------
Dim K As Integer
GetBuffer = Null
recAs400Cmd.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recAs400Cmd.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recAs400Cmd.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recAs400Cmd.Err = Space$(10) Then
    recAs400Cmd.Text = mId$(MsgTxt, K + 1, 478)
Else
    GetBuffer = recAs400Cmd.Err
End If

MsgTxtIndex = MsgTxtIndex + recAs400CmdLen

End Function

'---------------------------------------------------------
Public Sub Init(recAs400Cmd As typeAs400Cmd)
'---------------------------------------------------------
 MsgTxt = Space$(recAs400CmdLen)
 MsgTxtIndex = 0
 Call GetBuffer(recAs400Cmd)
 recAs400Cmd.obj = "SRVCMD      "
End Sub





'---------------------------------------------------------
Private Sub PutBuffer(recAs400Cmd As typeAs400Cmd)
'---------------------------------------------------------
Dim K As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recAs400Cmd.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recAs400Cmd.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 478) = recAs400Cmd.Text

MsgTxtLen = MsgTxtLen + recAs400CmdLen
End Sub


'-----------------------------------------------------
Function Update(recAs400Cmd As typeAs400Cmd)
'-----------------------------------------------------

Update = "?"

MsgTxtLen = 0
Call PutBuffer(recAs400Cmd)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    recAs400Cmd.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
    If Trim(recAs400Cmd.Err) <> "" Then
        Call ErrorX(recAs400Cmd)
        Update = recAs400Cmd.Err
        Exit Function
    Else
        Update = Null
    End If
Else
    recAs400Cmd.Err = "srv"
End If

End Function





Public Sub FTP_Get(lAS400CL As String, lFileNameOrig As String, lFileNameFTP As String)
Dim x As String, recAs400Cmd As typeAs400Cmd

x = Dir(lFileNameFTP)
If x <> "" Then Kill lFileNameFTP

FileCopy lFileNameOrig, lFileNameFTP

srvAs400Cmd.Init recAs400Cmd
recAs400Cmd.Method = "SBMJOB"

x = "SBMJOB CMD(CALL PGM(" & lAS400CL & "))"
recAs400Cmd.Text = x & " JOB(" & lAS400CL & ") USER(" & Trim(usrId) & ") JOBQ(QINTER)"
srvAs400Cmd.Update recAs400Cmd


End Sub
Public Sub SBMJOB(lAS400CL As String)
Dim x As String, recAs400Cmd As typeAs400Cmd

srvAs400Cmd.Init recAs400Cmd
recAs400Cmd.Method = "SBMJOB"

x = "SBMJOB CMD(CALL PGM(" & lAS400CL & "))"
recAs400Cmd.Text = x & " JOB(" & lAS400CL & ") USER(" & Trim(usrId) & ") JOBQ(QINTER)"
srvAs400Cmd.Update recAs400Cmd


End Sub

