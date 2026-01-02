Attribute VB_Name = "ElpTimer"
Option Explicit

Public blnElpTimer_Password As Boolean, blnElpTimer_Auto As Boolean, blnElpTimer_Receive As Boolean
Public paramElpTimer_Send As String, paramElpTimer_Receive As String
Public paramElpTimer_Folder As String
Public paramElpTimer_Wait As String
Public paramElpTimer_Flash As String
Public paramElpTimer_Day As Integer
Public paramElpTimer_Id  As String


Dim vShellId, SssSys As Long

Type typeElpTimer
    Obj             As String * 12
    Method          As String * 12
    Err             As String * 10
    Function        As String * 12
    HmsStart        As String * 6
    HmsStop         As String * 6
    HmsDelay        As String * 6
    HmsPrevious     As String * 6
    HmsNext         As String * 6
    Command         As String
    blnStop         As Boolean
    SssNext         As Long
    SssDelay        As Long
    SssStop         As Long
    Nb              As Long
    FunctionX       As String * 5
    Name            As String * 40
End Type
    

Public arrElpTimer() As typeElpTimer
Public arrElpTimer_Nb As Integer
Public arrElpTimer_NbMax As Integer
Public arrElpTimer_Index As Integer

Public Sub ElpTimer_Init()
Dim K As Integer, wWeekDay As Integer
Dim X As String, xMemo As String, xK2 As String

DSYS_Init
mainSoc_AMJCPT_Load

'aujourd'hui est-il ouvré ?
'----------------------------
If DSys < YBIATAB0_DATE_CPT_JS1 Then
    End
'    MsgBox "Automate non activé (date du jour < date compta SAB) : " & DSys & " < " & YBIATAB0_DATE_CPT_JS1
'    Exit Sub
End If
paramElpTimer_Day = Day(Now)
wWeekDay = Weekday(Now)

paramElpTimer_Folder = ""
paramElpTimer_Send = ""
paramElpTimer_Receive = ""
blnElpTimer_Auto = True 'False
paramElpTimer_Wait = 60000 '300000
paramElpTimer_Flash = 5000

ReDim arrElpTimer(10): arrElpTimer_Nb = 0

X = "select * from ElpTable where SNN = 0" _
    & " and id = 'Timer'" _
    & " and K1 = '" & paramElpTimer_Id & "'"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    xMemo = rsMDB("Memo")
    xK2 = Trim(rsMDB("K2"))
    If Mid$(xMemo, 21 + wWeekDay, 1) <> "0" Then
        arrElpTimer_Nb = arrElpTimer_Nb + 1
        If arrElpTimer_Nb >= UBound(arrElpTimer) Then ReDim Preserve arrElpTimer(arrElpTimer_Nb + 10)
        arrElpTimer(arrElpTimer_Nb).Function = xK2
        arrElpTimer(arrElpTimer_Nb).Name = rsMDB("Name")
        arrElpTimer(arrElpTimer_Nb).FunctionX = UCase$(Mid$(xK2, 1, 5))
        arrElpTimer(arrElpTimer_Nb).HmsStart = Mid$(xMemo, 1, 6)
        arrElpTimer(arrElpTimer_Nb).HmsStop = Mid$(xMemo, 8, 6)
        arrElpTimer(arrElpTimer_Nb).HmsDelay = Mid$(xMemo, 15, 6)
        K = Len(xMemo) - 29
        If K > 0 Then
            arrElpTimer(arrElpTimer_Nb).Command = Chr$(34) & Trim(Mid$(xMemo, 30, K)) & Chr$(34)
        Else
            arrElpTimer(arrElpTimer_Nb).Command = ""
        End If
        arrElpTimer(arrElpTimer_Nb).blnStop = False
        arrElpTimer(arrElpTimer_Nb).Nb = 0
        arrElpTimer(arrElpTimer_Nb).HmsNext = arrElpTimer(arrElpTimer_Nb).HmsStart
        arrElpTimer(arrElpTimer_Nb).SssNext = Time_Hms_Sss(arrElpTimer(arrElpTimer_Nb).HmsStart)
        arrElpTimer(arrElpTimer_Nb).SssDelay = Time_Hms_Sss(arrElpTimer(arrElpTimer_Nb).HmsDelay)
        arrElpTimer(arrElpTimer_Nb).SssStop = Time_Hms_Sss(arrElpTimer(arrElpTimer_Nb).HmsStop)
        arrElpTimer(arrElpTimer_Nb).SssNext = arrElpTimer(arrElpTimer_Nb).SssNext - arrElpTimer(arrElpTimer_Nb).SssDelay
        If arrElpTimer(arrElpTimer_Nb).SssStop = 0 Then
            arrElpTimer(arrElpTimer_Nb).SssStop = arrElpTimer(arrElpTimer_Nb).SssNext + 1
            arrElpTimer(arrElpTimer_Nb).HmsStop = arrElpTimer(arrElpTimer_Nb).HmsStart
        End If
        ''If arrElpTimer(arrElpTimer_Nb).HmsDelay = "000000" Then arrElpTimer(arrElpTimer_Nb).HmsDelay = "000001"
        Call ElpTimer_Next(arrElpTimer(arrElpTimer_Nb))
        If UCase$(Trim(arrElpTimer(arrElpTimer_Nb).Function)) = "STOP" Then arrElpTimer(arrElpTimer_Nb).blnStop = False
    End If
    rsMDB.MoveNext
Loop
ElpTimer_Display

If Not frmElp.Timer1.Enabled Then ElpTimer_Monitor "Start"
End Sub
Public Sub ElpTimer_Monitor(Fct As String)
Dim I As Integer

On Error Resume Next
'$$$$$$$$$$$$$$$$$$$
frmElp.lblElpTimer.Caption = mCommand & " : " & Time
frmElp.Caption = frmElp.lblElpTimer.Caption

Select Case Fct
    Case "Auto"
                If paramElpTimer_Day <> Day(Now) Then ElpTimer_Init

                SssSys = Time_Sys_Sss
                For I = 1 To arrElpTimer_Nb
                    If Not arrElpTimer(I).blnStop Then
                        If arrElpTimer(I).SssNext <= SssSys Then
                            frmElp.lblElpTimer_Next.Caption = DSys & " : " & Time & " : " & arrElpTimer(I).Function
                            arrElpTimer(I).Nb = arrElpTimer(I).Nb + 1
                            ElpTimer_Function arrElpTimer(I)
                            ElpTimer_Next arrElpTimer(I)
                            ElpTimer_Display
                        End If
                    End If
                Next I
                
    Case "Stop"
            frmElp.lblElpTimer.Visible = False
            frmElp.lblElpTimer_Next.Visible = False
            blnElpTimer_Receive = False
            Call FlashWindow(frmElp.hwnd, False)
            frmElp.Timer1.Enabled = False
            Unload frmElp
    Case "Start"
            Call FlashWindow(frmElp.hwnd, True)
            frmElp.lstMain.Enabled = False
            frmElp.lblElpTimer.ForeColor = warnUsrColor
            frmElp.lblElpTimer.Visible = True: frmElp.lblElpTimer.Caption = Time
            frmElp.lblElpTimer_Next.Visible = True
            frmElp.lblElpTimer_Next.Caption = Time & " : $Auto_start"
            blnElpTimer_Receive = False
            frmElp.Timer1.Enabled = blnElpTimer_Auto
End Select

End Sub


Public Sub ElpTimer_Display()
Dim I As Integer, X As String
frmElp.lstMain.Clear
frmElp.lstMain.Visible = True
frmElp.lstMain.Height = 240
XLabel.Visible = True
XLabel.Caption = "Planing"
frmElp.lstMain.Width = 13000
frmElp.lstMain.FontSize = 9
frmElp.lstMain.Top = 600

For I = 1 To arrElpTimer_Nb
    X = IIf(arrElpTimer(I).blnStop, "   ___ ", "  <<<  ")
    If arrElpTimer(I).HmsDelay = "000000" Then
        frmElp.lstMain.AddItem timeImp(arrElpTimer(I).HmsNext) & X & Trim(arrElpTimer(I).Function) & vbTab & Trim(arrElpTimer(I).Name)
    Else
        frmElp.lstMain.AddItem timeImp(arrElpTimer(I).HmsNext) & X & Trim(arrElpTimer(I).Function) & vbTab & Trim(arrElpTimer(I).Name) & vbTab & " + " & timeImp(arrElpTimer(I).HmsDelay) & X & timeImp(arrElpTimer(I).HmsStop) & vbTab & arrElpTimer(I).Nb
    End If
Next I

Elp_ResizeControl frmElp.lstMain

End Sub

Public Sub ElpTimer_Next(lElpTimer As typeElpTimer)
Dim wL As Long, blnOk As Boolean
blnOk = False
SssSys = Time_Sys_Sss
wL = lElpTimer.SssNext
Do
    wL = wL + lElpTimer.SssDelay
    If wL >= lElpTimer.SssStop Then
         lElpTimer.blnStop = True
         blnOk = True
    Else
        If wL > SssSys Then
            lElpTimer.HmsPrevious = lElpTimer.HmsNext
            lElpTimer.HmsNext = Time_Sss_Hms(wL)
            lElpTimer.SssNext = wL
            blnOk = True
        Else
            If lElpTimer.SssDelay = 0 Then lElpTimer.blnStop = True: blnOk = True
       End If
    End If
Loop Until blnOk
   
End Sub




Public Sub ElpTimer_Function(lElpTimer As typeElpTimer)
Dim X13 As String * 13

Dim Msg As String
Select Case lElpTimer.FunctionX
    Case "STOP ": End
    Case "SHELL":
        vShellId = Shell(lElpTimer.Command, 1)
        AppActivate vShellId
        DoEvents
'jpl2001.09.21        SendKeys "%{F4}", True

    Case Else:
        X13 = lElpTimer.Function
        Msg = X13 & lElpTimer.Command & Space(100): Call Msg_Monitor(Msg)
End Select


End Sub
