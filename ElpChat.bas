Attribute VB_Name = "ElpJpl"
Option Explicit

Public blnElpChat_Password As Boolean, blnElpChat_Auto As Boolean, blnElpChat_Receive As Boolean
Public paramElpChat_Send As String, paramElpChat_Receive As String
Public xElpTable_Chat As typeElpTable
Public paramElpChat_Folder As String
Public paramElpChat_Wait As String
Public paramElpChat_Flash As String

Public recElpChat As typeElpTimer
'---------------------------------------------------------
Public Sub frmElpChat_Show()
'---------------------------------------------------------
Dim X As String

frmElpChat.Show vbModeless
frmElpChat.WindowState = vbNormal
frmElpChat.Visible = True
X = frmElpChat.Caption
AppActivate X
End Sub

Public Sub ElpChat_Init()
Dim X As String, wPass As String

wPass = "6440"
paramElpChat_Folder = ""
paramElpChat_Send = ""
paramElpChat_Receive = ""
blnElpChat_Auto = False
paramElpChat_Wait = 300000
paramElpChat_Flash = 0 '30000

recElpChat.Function = "ElpChat"
recElpChat.HmsStart = "000000"
recElpChat.HmsStop = "240000"
recElpChat.HmsDelay = "000100"
    recElpChat.Command = ""
recElpChat.blnStop = False
recElpChat.Nb = 0
recElpChat.HmsNext = recElpChat.HmsStart
recElpChat.SssNext = Time_Hms_Sss(recElpChat.HmsStart)
recElpChat.SssDelay = Time_Hms_Sss(recElpChat.HmsDelay)
recElpChat.SssStop = Time_Hms_Sss(recElpChat.HmsStop)
recElpChat.SssNext = recElpChat.SssNext - recElpChat.SssDelay

recElpTable_Init xElpTable_Chat
'xElpTable_Chat.Id = "ElpChat"
'xElpTable_Chat.Method = "Seek="
'xElpTable_Chat.K1 = "Folder"
'xElpTable_Chat.Err = tableElpTable_Read(xElpTable_Chat)
'If xElpTable_Chat.Err <> 0 Then Exit Sub
'If IsNull(xElpTable_Chat.Memo) Then Exit Sub
'paramElpChat_Folder = Trim(xElpTable_Chat.Memo)


blnElpChat_Auto = ElpChat_Monitor(wPass)

X = "106104114098007005004002002000006007106107115098119119124117126117104115089065070066095081070108"

paramElpChat_Folder = ElpCipher_D(X, wPass)
paramElpChat_Folder = "W:\Loulergue\"
If blnElpChat_Auto Then
    xElpTable_Chat.Method = "Seek="
    xElpTable_Chat.K1 = "Key"
    xElpTable_Chat.Err = tableElpTable_Read(xElpTable_Chat)
    If xElpTable_Chat.Err = 0 Then
            paramElpChat_Send = paramElpChat_Folder & mId$(xElpTable_Chat.Memo, 61, 60)
            paramElpChat_Receive = paramElpChat_Folder & mId$(xElpTable_Chat.Memo, 121, 60)
            ElpChat_Timer "Start"
    End If
End If

End Sub

Public Sub ElpChat_Timer(Fct As String)
On Error Resume Next
Dim X As String, SssSys As Long
Static lWait As Long

Select Case Fct
    Case "Auto"
            SssSys = Time_Sys_Sss
            If Not recElpChat.blnStop Then
               If recElpChat.SssNext <= SssSys Then
               
                    X = Dir(paramElpChat_Receive)
                    If X <> "" Then
                      frmElp.lblMain.ForeColor = txtUsr.ForeColor
                      blnElpChat_Receive = True
                      ElpTimer_Next recElpChat

                    End If
                End If
              End If
 '                       If blnElpChat_Receive Then
 '                            If lWait < paramElpChat_Flash Then
 '                                Call FlashWindow(frmElp.hwnd, True)
 '                                lWait = lWait + frmElp.Timer1.Interval
 '                            End If
 '                        Else
 '                            lWait = lWait + frmElp.Timer1.Interval
 '                            If lWait > paramElpChat_Wait Then
 '                                lWait = 0
 '                                X = Dir(paramElpChat_Receive)
 '                                If X <> "" Then
 '                                    frmElp.lblMain.ForeColor = txtUsr.ForeColor
 '                                    ''''frmElp.cmdElpChat.Visible = True
 '                                    blnElpChat_Receive = True
 '                                   frmElp.Timer1.Interval = 60000 '500 ' flash                    Else
 '                                    frmElp.Timer1.Interval = 60000
 '                                End If
 '                            End If
 '                        End If
    Case "Stop"
            ''''frmElp.cmdElpChat.Visible = False
            blnElpChat_Receive = False
            Call FlashWindow(frmElp.hwnd, False)
            frmElp.Timer1.Enabled = False
            lWait = 0
    Case "Start"
            frmElp.lblMain.ForeColor = lblUsr.ForeColor
            '''frmElp.cmdElpChat.Visible = False
            blnElpChat_Receive = False
            frmElp.Timer1.Enabled = blnElpChat_Auto
            lWait = 0
End Select

End Sub
Public Function ElpChat_Monitor(wKey As String) As Boolean
Dim w30 As String * 30, x30 As String * 30, wPass As String
ElpChat_Monitor = False
wPass = Trim(wKey)
If wPass = "" Then Exit Function
xElpTable_Chat.Method = "Seek>="
xElpTable_Chat.ID = "ElpChat"
xElpTable_Chat.K1 = "Monitor"
xElpTable_Chat.K2 = ""
xElpTable_Chat.Err = 0
w30 = Trim(usrId)
x30 = ElpCipher_C(w30, wPass)
Do
    xElpTable_Chat.Err = tableElpTable_Read(xElpTable_Chat)
    If xElpTable_Chat.Err = 0 Then
        If Trim(xElpTable_Chat.K1) <> "Monitor" _
        Or Trim(xElpTable_Chat.ID) <> "ElpChat" Then
            xElpTable_Chat.Err = 9996
        Else
            If mId$(xElpTable_Chat.Memo, 1, 30) = x30 Then
                ElpChat_Monitor = True
                Exit Function
            Else
                xElpTable_Chat.Method = "Seek>"
            End If
        End If
    End If
Loop While xElpTable_Chat.Err = 0

End Function




