Attribute VB_Name = "srvElpCom"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Private hInputData As Long
Private hInputQueue As Long
Private hOutputData As Long
Private hOutputQueue As Long
Const hError As Long = 0
Public paramCAV4_Wait As Long  '= 150 '$JPL 2002.11.05 pb durée WRKSPLF    '30

Private As400Dtaq     As String * 31744
Private As400DtaqIn   As String
Private As400DtaqOut  As String
Private As400DtaqLen As Long
Private As400SndOk As Boolean
Private As400RcvOk As Boolean

Type typeXcom
   SrvObj       As String * 12
   SrvMethod    As String * 12
   SrvErr       As String * 10
   usrId       As String * 10
   pcId        As String * 10
   SrvType     As String * 10
   SrvId       As String * 10
   SrvDtaqLib  As String * 10
   SrvDtaqIn   As String * 10
   SrvDTaqOut  As String * 10
   SrvDTaqLen  As String * 5
   jplFree     As String * 5
End Type

Private Xcom As typeXcom
Private blnXCom As Boolean
Public XComlen  As Integer
Public elpSrvXcom  As String
Public elpSrvTxtin As Boolean
Public elpSrvTxtOut As Boolean

Public MsgTxt As String * 31630 ' 31744-114   '8078
Public MsgTxtLen As Integer
Public MsgTxtIndex  As Integer
Public Const MsgTxtLenMax As Long = 31630 '8078
Private Kerr        As Integer

Public blnSnd As Boolean, blnRcv As Boolean
Public blnFR_Convert  As Boolean

Public blnMsgTxt_Concat As Boolean, blnMsgTxt_Concat_Transaction As Boolean
'---------------------------------------------------------
Function SndRcv()
'----------------------------------------------------------
srvIdle = False

Screen.MousePointer = vbHourglass
blnXCom = False
Do
    As400Dtaq = ""
    Call XCom_PutBuffer
    As400Dtaq = mId$(As400Dtaq, 1, XComlen) & MsgTxt
    As400DtaqLen = XComlen + MsgTxtLen

    Select Case elpSrvXcom
        Case "CAV4": SndRcv = CAV4_SndRcv()
        Case "": blnXCom = True: SndRcv = Null
        Case Else: MsgBox "ATTENTION : màj non faite , prévenir JPL", vbCritical, "srvElpCom_SndRcv": End
    End Select
    If Elp.usrId = Xcom.usrId _
    And Elp.pcId = Xcom.pcId Then blnXCom = True
Loop While Not blnXCom

Elp = Xcom

If elpSrvTxtOut Then
    Write #1, As400DtaqLen, As400Dtaq
End If

If As400DtaqLen < XComlen Then
    SndRcv = "As400DtaqLen"
Else
    MsgTxtLen = As400DtaqLen - XComlen
    MsgTxt = mId$(As400Dtaq, XComlen + 1, MsgTxtLen)
End If
Screen.MousePointer = vbDefault

srvIdle = True


End Function

'---------------------------------------------------------
Private Sub XCom_PutBuffer()
'---------------------------------------------------------

''''Xcom.usrId = Elp.usrId
''''Xcom.pcId = Elp.pcId

Mid$(As400Dtaq, 1, 12) = Xcom.SrvObj
Mid$(As400Dtaq, 13, 12) = Xcom.SrvMethod
Mid$(As400Dtaq, 25, 10) = Space$(10)
Mid$(As400Dtaq, 35, 10) = Xcom.usrId
Mid$(As400Dtaq, 45, 10) = Xcom.pcId
Mid$(As400Dtaq, 55, 10) = Xcom.SrvType
Mid$(As400Dtaq, 65, 10) = Xcom.SrvId
Mid$(As400Dtaq, 75, 10) = Xcom.SrvDtaqLib
Mid$(As400Dtaq, 85, 10) = Xcom.SrvDtaqIn
Mid$(As400Dtaq, 95, 10) = Xcom.SrvDTaqOut
Mid$(As400Dtaq, 105, 5) = Xcom.SrvDTaqLen
Mid$(As400Dtaq, 110, 5) = Xcom.jplFree

End Sub

'---------------------------------------------------------
Private Sub XCom_GetBuffer()
'---------------------------------------------------------

With Xcom
    .SrvObj = mId$(As400Dtaq, 1, 12)
    .SrvMethod = mId$(As400Dtaq, 13, 12)
    .SrvErr = mId$(As400Dtaq, 25, 10)
    .usrId = mId$(As400Dtaq, 35, 10)
    .pcId = mId$(As400Dtaq, 45, 10)
    .SrvType = mId$(As400Dtaq, 55, 10)
    .SrvId = mId$(As400Dtaq, 65, 10)
    .SrvDtaqLib = mId$(As400Dtaq, 75, 10)
    .SrvDtaqIn = mId$(As400Dtaq, 85, 10)
    .SrvDTaqOut = mId$(As400Dtaq, 95, 10)
    .SrvDTaqLen = mId$(As400Dtaq, 105, 5)
    .jplFree = mId$(As400Dtaq, 110, 5)
End With

'''''Elp.usrId = Xcom.usrId
'''''Elp.pcId = Xcom.pcId

End Sub

'---------------------------------------------------------
Public Sub sndMsgTxt_Init()
'---------------------------------------------------------

Mid$(MsgTxt, 1, 12) = "SRVMSGTXT   "
Mid$(MsgTxt, 13, 12) = Space$(12)
Mid$(MsgTxt, 25, 10) = Space$(10)

MsgTxtLen = 34

End Sub

'---------------------------------------------------------
Public Function sndMsgTxt_Ok()
'---------------------------------------------------------

'$$$$$ à faire : tester màj de chaque enregistrement et rollback

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    sndMsgTxt_Ok = Null
Else
    sndMsgTxt_Ok = "????"
End If


End Function




'---------------------------------------------------------
Public Function SndRcv_Init()
'---------------------------------------------------------
Dim X As String

On Error GoTo ErrorX
blnSnd = True: blnRcv = True
blnMsgTxt_Concat = True
blnMsgTxt_Concat_Transaction = blnMsgTxt_Concat

SndRcv_Init = Null
XComlen = 114
Xcom = Elp

X = "c:\" & Trim(Xcom.SrvDtaqLib) & "\" & Trim(Xcom.SrvDtaqIn) & ".TXT"
If elpSrvTxtOut Then Open X For Output As #1

Select Case elpSrvXcom
    Case Is = "CAV4": SndRcv_Init = CAV4_Init
    Case ""

End Select

Exit Function

ErrorX:
        MsgBox "Erreur :" & Err & " : " & Error$(Err), vbCritical, "Initialisation : " & X
    SndRcv_Init = Err
    Exit Function
End Function

'---------------------------------------------------------
Public Sub XCom_End()
'---------------------------------------------------------

MsgTxtLen = 0
Xcom.SrvMethod = "ELPDTAQEND  "
SndRcv

End Sub

Public Function CAV4_Init()
Dim X1 As String, X2 As String, X3 As String

paramCAV4_Wait = 150

CAV4_Init = Null
hInputData = cwbDQ_CreateData()
If hInputData = 0 Then
    CAV4_Init = "err"
    MsgBox "Erreur CreateData :", vbCritical, " Input"
    Exit Function
End If

Kerr = cwbDQ_SetConvert(hInputData, CWB_TRUE)
If Kerr <> 0 Then
    CAV4_Init = Kerr
    MsgBox "Erreur SetMode :" & CStr(Kerr), vbCritical, "Liaison Serveur AS400"
    Exit Function
End If

Kerr = cwbDQ_Open(Trim(Xcom.SrvDtaqIn), Trim(Xcom.SrvDtaqLib), Trim(Xcom.SrvId), hInputQueue, hError)
If Kerr <> 0 Then
    CAV4_Init = Kerr
    MsgBox "Erreur CAV4_Init :" & CStr(Kerr), vbCritical, "Liaison Serveur AS400"
    Exit Function
End If

hOutputData = cwbDQ_CreateData()
If hOutputData = 0 Then
    CAV4_Init = "err"
    MsgBox "Erreur CreateData :", vbCritical, " Output"
    Exit Function
End If

Kerr = cwbDQ_SetConvert(hOutputData, CWB_TRUE)
If Kerr <> 0 Then
    CAV4_Init = Kerr
    MsgBox "Erreur SetMode :" & CStr(Kerr), vbCritical, "Liaison Serveur AS400"
    Exit Function
End If

Kerr = cwbDQ_Open(Trim(Xcom.SrvDTaqOut), Trim(Xcom.SrvDtaqLib), Trim(Xcom.SrvId), hOutputQueue, hError)
If Kerr <> 0 Then
    CAV4_Init = Kerr
    MsgBox "Erreur open output :" & CStr(Kerr), vbCritical, "Liaison Serveur AS400"
    Exit Function
End If
MsgTxt = Space$(31744 - XComlen)
MsgTxtLen = 0
Xcom.SrvMethod = "PCINIT      "

If IsNull(SndRcv()) Then
    Xcom.SrvMethod = Space$(12)
    Kerr = cwbDQ_Close(hOutputQueue)
    If Kerr <> 0 Then
        CAV4_Init = Kerr
        MsgBox "Erreur close output :" & CStr(Kerr), vbCritical, "Liaison Serveur AS400"
        Exit Function
    Else
        Kerr = cwbDQ_Open(Trim(Xcom.SrvDTaqOut), Trim(Xcom.SrvDtaqLib), Trim(Xcom.SrvId), hOutputQueue, hError)
        If Kerr <> 0 Then
            CAV4_Init = Kerr
            MsgBox "Erreur open output :" & CStr(Kerr), vbCritical, "Liaison Serveur AS400"
            Exit Function
        End If
    End If
End If

End Function

Public Function CAV4_SndRcv()
Dim X As String, I As Integer, Iter As Integer
Dim blnSndCAV4 As Boolean, blnRcvCAV4 As Boolean

'jpl.2000.01.26 testDebug.Print "CAV4_SndRcv S :", countTimer; Time   'jpl.2000.01.26

CAV4_SndRcv = Null
Iter = 0
'2001.08.23 JPL cf AS400 : SRVANSIRCV SRVANSISND   If blnFR_Convert Then FR_ConvertAtoE

blnSndCAV4 = blnSnd: blnSnd = True
blnRcvCAV4 = blnRcv: blnRcv = True
'''Call MsgBox(mId$(As400Dtaq, 1, As400DtaqLen), vbInformation, "CAV4_SndRcv : " & As400DtaqLen)
Do
    As400SndOk = True
    If Not blnSndCAV4 Then
        Kerr = 0
    Else
        Kerr = cwbDQ_SetDataAddr(hInputData, As400Dtaq, As400DtaqLen)
        If Kerr = 0 Then Kerr = cwbDQ_Write(hInputQueue, hInputData, CWB_FALSE, hError)
    End If
    DoEvents
    If Kerr <> 0 Then
        I = MsgBox("Erreur SetData / Write :" & CStr(Kerr), vbRetryCancel + vbCritical, "Liaison Serveur AS400")
        If I = vbRetry Then
            As400SndOk = False
        Else
            CAV4_SndRcv = Kerr
            End 'Exit Function
        End If
    Else
      
        If blnRcvCAV4 Then
            Do
                DoEvents
                As400Dtaq = ""
                As400RcvOk = True
                As400DtaqLen = 31744
                Kerr = cwbDQ_SetData(hOutputData, As400Dtaq, As400DtaqLen)
                If Kerr <> 0 Then
                    I = MsgBox("Erreur SetData (réception) :" & CStr(Kerr), vbRetryCancel + vbCritical, "Liaison Serveur AS400")
                Else
                    
                    Kerr = cwbDQ_Read(hOutputQueue, hOutputData, ByVal paramCAV4_Wait, hError)
                    If Kerr = 0 Then Kerr = cwbDQ_GetData(hOutputData, As400Dtaq)
                    If Kerr <> 0 Then
                       'Unload frmElp
                    
                        I = MsgBox("Erreur réception :" & CStr(Kerr), vbRetryCancel + vbCritical, "Liaison Serveur AS400")
                        If I = vbRetry Then
                            As400RcvOk = False
                        Else
                            CAV4_SndRcv = Kerr
                            End 'Exit Function
                        End If
                    Else
    '                  As400DtaqLen = Len(RTrim(As400Dtaq))
    '                    If As400DtaqLen = 0 Then
    '                        As400RcvOk = False
    '                    End If
                    End If
                End If
            DoEvents
            Loop Until As400RcvOk
        End If
        
 'jpl.2000.01.26 testDebug.Print "CAV4_SndRcv R :", countTimer; Time   'jpl.2000.01.26
     
        Call XCom_GetBuffer
        As400DtaqLen = Xcom.SrvDTaqLen
        Select Case Xcom.SrvErr
            Case Space$(10): As400RcvOk = True
            Case "SRVRETRY  "
                Xcom = Elp: Call XCom_PutBuffer
                As400SndOk = False
                Iter = Iter + 1
                If Iter > 5 Then
                    I = MsgBox("AS400 indisponible ", vbRetryCancel + vbCritical, "Liaison Serveur AS400")
                    If I = vbRetry Then
                        Iter = 0
                    Else
                        As400SndOk = True
                        CAV4_SndRcv = 9999
                        End   'xit Function
                    End If
                End If
            Case Else
                CAV4_SndRcv = 9999
                MsgBox "Erreur AS400 : " & Xcom.SrvErr
            End Select
        
    End If
Loop Until As400SndOk
'jpl.2000.01.26 testDebug.Print "CAV4_SndRcv X :", countTimer; Time   'jpl.2000.01.26

'2001.08.23 JPL cf AS400 : SRVANSIRCV SRVANSISND '' If blnFR_Convert Then FR_ConvertEtoA
blnFR_Convert = True

'jpl.2000.01.26 testDebug.Print "CAV4_SndRcv E :", countTimer; Time   'jpl.2000.01.26

End Function


Public Sub CAV4_Close()
Dim I As Integer
If hInputData <> 0 Then I = cwbDQ_DeleteData(hInputData)
If hOutputData <> 0 Then I = cwbDQ_DeleteData(hOutputData)
If hInputQueue <> 0 Then I = cwbDQ_Close(hInputQueue)
If hOutputQueue <> 0 Then I = cwbDQ_Close(hOutputQueue)
End Sub




Public Sub FR_ConvertEtoA()

Exit Sub '2001.08.23 JPL cf AS400 : SRVANSIRCV SRVANSISND

Dim I As Integer
For I = XComlen + 1 To As400DtaqLen

    If Asc(mId$(As400Dtaq, I, 1)) >= 128 Then
        Select Case Asc(mId$(As400Dtaq, I, 1))
            Case Is = 133: Mid$(As400Dtaq, I, 1) = Chr$(224)
            Case Is = 130: Mid$(As400Dtaq, I, 1) = Chr$(233)
            Case Is = 138: Mid$(As400Dtaq, I, 1) = Chr$(232)
            Case Is = 136: Mid$(As400Dtaq, I, 1) = Chr$(234)
            Case Is = 137: Mid$(As400Dtaq, I, 1) = Chr$(235)
            Case Is = 135: Mid$(As400Dtaq, I, 1) = Chr$(231)
            Case Is = 151: Mid$(As400Dtaq, I, 1) = Chr$(249)
            Case Is = 129: Mid$(As400Dtaq, I, 1) = Chr$(252)
            Case Is = 139: Mid$(As400Dtaq, I, 1) = Chr$(239)
            Case Is = 140: Mid$(As400Dtaq, I, 1) = Chr$(238)
            Case Is = 147: Mid$(As400Dtaq, I, 1) = Chr$(244)
            Case Is = 148: Mid$(As400Dtaq, I, 1) = Chr$(246)
            Case Is = 216: Mid$(As400Dtaq, I, 1) = Chr$(207)
            Case Is = 245: Mid$(As400Dtaq, I, 1) = Chr$(167)
            Case Is = 153: Mid$(As400Dtaq, I, 1) = Chr$(214)
     End Select
    End If
    
Next I

'jpl.2000.03.26 Dim I As Integer, K As Integer
'jpl.2000.03.26 For I = XComlen + 1 To As400DtaqLen
'jpl.2000.03.26     K = Asc(mId$(As400Dtaq, I, 1))
'jpl.2000.03.26 modèle as400dtaq==> lX            Case Is = 133: Mid$(As400Dtaq, I, 1) = Chr$(224)
End Sub


Public Sub FR_ConvertEtoA_X(lX As String)
Exit Sub '2001.08.23 JPL cf AS400 : SRVANSIRCV SRVANSISND


'jpl.2000.03.26 appel du sous-programme uniquement pour les intitulés

Dim I As Integer, K As Integer
For I = 1 To Len(lX)

    K = Asc(mId$(lX, I, 1))
    If K >= 128 Then
        Select Case K
            Case Is = 133: Mid$(lX, I, 1) = Chr$(224)
            Case Is = 130: Mid$(lX, I, 1) = Chr$(233)
            Case Is = 138: Mid$(lX, I, 1) = Chr$(232)
            Case Is = 136: Mid$(lX, I, 1) = Chr$(234)
            Case Is = 137: Mid$(lX, I, 1) = Chr$(235)
            Case Is = 135: Mid$(lX, I, 1) = Chr$(231)
            Case Is = 151: Mid$(lX, I, 1) = Chr$(249)
            Case Is = 129: Mid$(lX, I, 1) = Chr$(252)
            Case Is = 139: Mid$(lX, I, 1) = Chr$(239)
            Case Is = 140: Mid$(lX, I, 1) = Chr$(238)
            Case Is = 147: Mid$(lX, I, 1) = Chr$(244)
            Case Is = 148: Mid$(lX, I, 1) = Chr$(246)
            Case Is = 216: Mid$(lX, I, 1) = Chr$(207)
            Case Is = 245: Mid$(lX, I, 1) = Chr$(167)
            Case Is = 153: Mid$(lX, I, 1) = Chr$(214)
     End Select
    End If
   
Next I


End Sub

Public Sub FR_ConvertAtoE()

Exit Sub '2001.08.23 JPL cf AS400 : SRVANSIRCV SRVANSISND

Dim I As Integer
For I = XComlen + 1 To As400DtaqLen
'    If Asc(Mid$(As400Dtaq, I, 1)) = 0 Then
'        Mid$(As400Dtaq, I, 1) = " "
'    End If
    If Asc(mId$(As400Dtaq, I, 1)) >= 167 Then
        Select Case Asc(mId$(As400Dtaq, I, 1))
            Case Is = 167: Mid$(As400Dtaq, I, 1) = Chr$(245)
            Case Is = 207: Mid$(As400Dtaq, I, 1) = Chr$(216)
            Case Is = 214: Mid$(As400Dtaq, I, 1) = Chr$(153)
            Case Is = 224: Mid$(As400Dtaq, I, 1) = Chr$(133)
            Case Is = 233: Mid$(As400Dtaq, I, 1) = Chr$(130)
            Case Is = 232: Mid$(As400Dtaq, I, 1) = Chr$(138)
            Case Is = 234: Mid$(As400Dtaq, I, 1) = Chr$(136)
            Case Is = 235: Mid$(As400Dtaq, I, 1) = Chr$(137)
            Case Is = 231: Mid$(As400Dtaq, I, 1) = Chr$(135)
            Case Is = 249: Mid$(As400Dtaq, I, 1) = Chr$(151)
            Case Is = 252: Mid$(As400Dtaq, I, 1) = Chr$(129)
            Case Is = 239: Mid$(As400Dtaq, I, 1) = Chr$(139)
            Case Is = 238: Mid$(As400Dtaq, I, 1) = Chr$(140)
            Case Is = 244: Mid$(As400Dtaq, I, 1) = Chr$(147)
            Case Is = 246: Mid$(As400Dtaq, I, 1) = Chr$(148)
       End Select
    End If
    
Next I

End Sub

Public Sub Xcom_UsrId(Msg As String)
Xcom.usrId = Msg
End Sub

Public Sub Shell_Exe(lFileName As String)
Dim IdShell As Variant, X As String
'X = "CMD/Q " & Chr$(34) & lFilename & Chr$(34)
IdShell = Shell(lFileName, 1)
If IdShell > 0 Then
    AppActivate IdShell
Else
    MsgBox lFileName, vbCritical, "Shell_Exe"
End If

DoEvents

End Sub
Public Sub Shell_FTP(lFTP_Nt_Filename As String, lFTP_AS400_Library As String, lFTP_AS400_File As String, blnFTP_Get As Boolean, blnFTP_AS400_Binary As Boolean)
Dim wFTP_Nt_Filename As String, wFTP_Nt_Filename_Bat As String, wFTP_Nt_Filename_Dta As String
Dim K As Integer, K1 As Integer
Dim wFTP_Nt_Filename_Log As String
Dim IdShell
Dim iWait As Long, xIn As String, blnOk As Boolean, blnLog_Quit As Boolean
Dim fsoFile As File
Dim wAMJHMS As String
On Error GoTo Error_Handle
Dim X As String
Dim xTrace As String

wFTP_Nt_Filename = Trim(lFTP_Nt_Filename)
wAMJHMS = "_" & DSys & "_" & time_Hms
xTrace = " : Init "
K1 = Len(wFTP_Nt_Filename)
K = InStr(K1 - 5, wFTP_Nt_Filename, ".")
If K > 0 Then K1 = K - 1
X = mId$(wFTP_Nt_Filename, 1, K1)
    
wFTP_Nt_Filename_Bat = X & wAMJHMS & ".bat"
wFTP_Nt_Filename_Dta = X & wAMJHMS & ".dta"
wFTP_Nt_Filename_Log = X & wAMJHMS & ".log"
Open wFTP_Nt_Filename_Dta For Output As #2
Print #2, paramIBM_BIA_Auto
Print #2, paramIBM_BIA_AUTO_Password
If blnFTP_AS400_Binary Then Print #2, "bin"
Print #2, "cd " & Trim(lFTP_AS400_Library)
If blnFTP_Get Then
    Print #2, "get " & Trim(lFTP_AS400_File) & " " & wFTP_Nt_Filename
Else
    Print #2, "put " & wFTP_Nt_Filename & " " & Trim(lFTP_AS400_File)
End If

Print #2, "quit"

Open wFTP_Nt_Filename_Bat For Output As #1
K1 = 1
Do
    K = InStr(K1, wFTP_Nt_Filename, "\")
    If K > 0 Then K1 = K + 1
Loop Until K <= 0

Print #1, "cd /d " & mId$(wFTP_Nt_Filename, 1, K1 - 2)
Print #1, "FTP -s:" & wFTP_Nt_Filename_Dta & " " & Trim("SABA") ' cboFTP_AS400_host)
Print #1, "del " & wFTP_Nt_Filename_Dta
Print #1, "del " & wFTP_Nt_Filename_Bat

Close
xTrace = " : Shell "

blnOk = False
blnLog_Quit = False

IdShell = Shell(wFTP_Nt_Filename_Bat & " > " & wFTP_Nt_Filename_Log, 1)
DoEvents

If IdShell > 0 Then
    xTrace = " : AppActivate "
                                        On Error Resume Next            '$$$$$$$$$$$
                                        AppActivate IdShell, True       '$$$$$$$$$$$
                                        On Error GoTo Error_Handle      '$$$$$$$$$$$
    
    xTrace = " : Read Log "
    Sleep 1000
    Do
        For iWait = 1 To 30
            DoEvents
            If Dir(wFTP_Nt_Filename_Log) <> "" Then
                Open wFTP_Nt_Filename_Log For Input As #1
                Do Until EOF(1)
                    Line Input #1, xIn
                 ''
                   If Trim(xIn) = "250 File transfer completed successfully." Then blnOk = True
                   If Trim(xIn) = "221 QUIT subcommand received." Then blnLog_Quit = True
                Loop
                Close #1
                If blnLog_Quit Then Exit For
                Sleep 1000
            End If
           ' Debug.Print iWait, xIn
        Next iWait
        If Not blnLog_Quit Then
            X = MsgBox("Fin du processus (221 QUIT)  non détecté, Voulez-vous tester à nouveau ?", vbYesNo + vbQuestion + vbDefaultButton1, "Shell_FTP")
            If X = vbYes Then blnLog_Quit = True
        End If
    Loop Until blnLog_Quit
End If
Close

Sleep 1000

xTrace = " : Close Log "
If blnOk Then
    xTrace = " : Kill bat "
    If Dir(wFTP_Nt_Filename_Bat) <> "" Then msFileSystem.DeleteFile wFTP_Nt_Filename_Bat
    xTrace = " : Kill log "
    If Dir(wFTP_Nt_Filename_Log) <> "" Then msFileSystem.DeleteFile wFTP_Nt_Filename_Log

Else
    Shell_MsgBox "Shell_FTP : " & wFTP_Nt_Filename_Log & " : en erreur ", vbCritical, XForm.Caption, False
End If

DoEvents

Exit Sub

Error_Handle:

Close
Shell_MsgBox wFTP_Nt_Filename & xTrace & ":" & Error, vbCritical, "Shell_FTP", False

End Sub

Public Sub Shell_MsgBox(lMsg As String, lConst As Long, lTitle As String, blnNetSend As Boolean)
Dim wNt_Filename_Bat As String
Dim IdShell
Dim intFile As Integer
Dim X As String, I As Integer
On Error Resume Next  'GoTo Error_Handle

If blnTimer_Enabled Or blnNetSend Then
    wNt_Filename_Bat = paramTemp_Folder & "\" & DSys & "_" & time_Hms & "_Net_Send.bat"
    intFile = FreeFile(0)
    Open wNt_Filename_Bat For Output As #intFile
    '''x = usrId & " : " & Trim(lTitle) & " : " & lMsg
     ''X = usrId & " : " & lMsg
     X = lMsg
   For I = 1 To Len(X)
        Select Case Asc(mId$(X, I, 1))
            
            Case Is < 32: Mid$(X, I, 1) = " "
            Case Is > 127: Mid$(X, I, 1) = " "
        End Select
    Next I
    Print #intFile, "Net Send LOULERGUE " & X
    'If Trim(usrId) <> "LOULERGUE" Then Print #intFile, "Net Send " & x
    Close intFile
    DoEvents
    
    IdShell = Shell(wNt_Filename_Bat, 1)
    
    If IdShell > 0 Then
        AppActivate IdShell, True       '$$$$$$$$$$$
    End If
    DoEvents
    Wait_SS 5
    If Dir(wNt_Filename_Bat) <> "" Then msFileSystem.DeleteFile wNt_Filename_Bat
  ''' Kill wNt_Filename_Bat

End If

DoEvents

If Not blnTimer_Enabled Then MsgBox lMsg, lConst, lTitle
Exit Sub

Error_Handle:
Close
If Not blnTimer_Enabled Then MsgBox wNt_Filename_Bat & ":" & Error, vbCritical, "Shell_MsgBox"

End Sub



