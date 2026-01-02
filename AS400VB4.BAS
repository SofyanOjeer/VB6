Attribute VB_Name = "srvAs400"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Declare Sub Sleep Lib "c:\windows\system\kernel32.dll" (ByVal dwMilliseconds As Long)

Declare Function EHNDQ_SetMode Lib "c:\pcs\EHNDQW.DLL" (ByVal hWnd As Integer, ByVal IpszQueueLocation As String, ByVal IMode As Long) As Integer

Declare Function EHNDQ_Clear Lib "c:\pcs\EHNDQW.DLL" (ByVal hWnd As Integer, ByVal IpszQueueName As String, ByVal IpszQeueLocation As String) As Integer

Declare Function EHNDQ_Send Lib "c:\pcs\EHNDQW.DLL" (ByVal hWnd As Integer, ByVal IpszQueueName As String, ByVal IpszQueueLocation As String, ByVal IpachDataBuffer As String, ByVal IDataLength As Long) As Integer

Declare Function EHNDQ_Receive Lib "c:\pcs\EHNDQW.DLL" (ByVal hWnd As Integer, ByVal IpszQueueName As String, ByVal IpszQueueLocation As String, ByVal IWaitTime As Long, ByVal nSenderID As Integer, ByVal IpachDataBuffer As String, IplLength As Long, ByVal IpachSenderIDinfo As String) As Integer


Private As400Dtaq     As String * 8192
Private As400DtaqIn   As String
Private As400DtaqOut  As String
Private As400DtaqLen As Long
Private As400SndOk As Boolean
Private As400RcvOk As Boolean

Private Kerr        As Integer

Public elpSrvXcom  As String
Public elpSrvTxtin As Boolean
Public elpSrvTxtOut As Boolean
Private XcomLen  As Integer

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
End Type

Private Xcom As typeXcom
Private tstXcom As Boolean

Public MsgTxt As String * 8088
Public MsgTxtLen As Integer
Public MsgTxtIndex  As Integer

'---------------------------------------------------------
Function SndRcv()
'----------------------------------------------------------
srvIdle = False

Screen.MousePointer = vbHourglass
tstXcom = False
Do
    As400Dtaq = ""
    Call XcomPutBuffer
    As400Dtaq = Mid$(As400Dtaq, 1, XcomLen) & MsgTxt
    As400DtaqLen = XcomLen + MsgTxtLen

    Select Case elpSrvXcom
        Case "PCS": SndRcv = SndRcvPCS()
        Case "MSCOMM1": SndRcv = SndRcvMSComm1(frmElp.MSComm1)
        Case ""
        Case Else: SndRcv = SndRcvXXX()
    End Select

    If Elp.usrId = Xcom.usrId _
    And Elp.pcId = Xcom.pcId Then tstXcom = True
Loop While Not tstXcom

Elp = Xcom

If elpSrvTxtOut Then
    Write #1, As400DtaqLen, As400Dtaq
End If

If As400DtaqLen < XcomLen Then
    SndRcv = "As400DtaqLen"
Else
    MsgTxtLen = As400DtaqLen - XcomLen
    MsgTxt = Mid$(As400Dtaq, XcomLen + 1, MsgTxtLen)
End If
Screen.MousePointer = vbDefault

srvIdle = True

End Function

'---------------------------------------------------------
Function SndRcvPCS()
'---------------------------------------------------------
Dim X As String, I As Integer, Iter As Integer
SndRcvPCS = Null
Iter = 0

Do
    I = InStr(1, Xcom.SrvDtaqLib, " ")
    X = Mid$(Xcom.SrvDtaqLib, 1, I - 1) & "/"
    As400DtaqIn = X & Xcom.SrvDtaqIn
    As400DtaqOut = X & Xcom.SrvDTaqOut

    As400SndOk = True
    Kerr = EHNDQ_Send(Screen.ActiveForm.hWnd, _
                As400DtaqIn, Xcom.SrvId, As400Dtaq, As400DtaqLen)
    If Kerr <> 0 Then
        I = MsgBox("Erreur envoi :" & CStr(Kerr), vbRetryCancel + vbCritical, "Liaison Serveur AS400")
        If I = vbRetry Then
            As400SndOk = False
        Else
            SndRcvPCS = Kerr
            Exit Function
        End If
    Else
        Do
            As400RcvOk = True
            As400DtaqLen = 8192
            Kerr = EHNDQ_Receive(Screen.ActiveForm.hWnd, _
                    As400DtaqOut, Xcom.SrvId, 15, 0, As400Dtaq, As400DtaqLen, "")
            
            
            
           If Kerr <> 0 Then
                I = MsgBox("Erreur réception :" & CStr(Kerr), vbRetryCancel + vbCritical, "Liaison Serveur AS400")
                If I = vbRetry Then
                    As400RcvOk = False
                Else
                    SndRcvPCS = Kerr
                    Exit Function
                End If
            Else
                If As400DtaqLen = 0 Then
                    As400RcvOk = False
                End If
            End If
        DoEvents
        Loop Until As400RcvOk
        
        Call XcomGetBuffer
        Select Case Xcom.SrvErr
            Case Space$(10): As400RcvOk = True
            Case "SRVRETRY  "
                Xcom = Elp: Call XcomPutBuffer
                As400SndOk = False
                Iter = Iter + 1
                If Iter > 5 Then
                    I = MsgBox("AS400 indisponible ", vbRetryCancel + vbCritical, "Liaison Serveur AS400")
                    If I = vbRetry Then
                        Iter = 0
                    Else
                        As400SndOk = True
                        SndRcvPCS = 9999
                        Exit Function
                    End If
                End If
            Case Else
                SndRcvPCS = 9999
                MsgBox "Erreur AS400 : " & Xcom.SrvErr
            End Select
        
    End If
        
    
Loop Until As400SndOk


For I = XcomLen + 1 To As400DtaqLen

    If Asc(Mid$(As400Dtaq, I, 1)) >= 128 Then
        Select Case Asc(Mid$(As400Dtaq, I, 1))
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
        End Select
    End If
    
Next I

End Function

'---------------------------------------------------------
Function SndRcvInitPCS()
'---------------------------------------------------------

SndRcvInitPCS = Null
Kerr = EHNDQ_SetMode(Screen.ActiveForm.hWnd, Xcom.SrvId, 1)
If Kerr <> 0 Then
    SndRcvInitPCS = Kerr
    MsgBox "Erreur SetMode :" & CStr(Kerr), vbCritical, "Liaison Serveur AS400"
    Exit Function
Else
    MsgTxt = Space$(8192 - XcomLen)
    MsgTxtLen = 0
    Xcom.SrvMethod = "PCINIT      "
    
    If IsNull(SndRcv()) Then
            Xcom.SrvMethod = Space$(12)
            Kerr = EHNDQ_Clear(Screen.ActiveForm.hWnd, As400DtaqOut, Xcom.SrvId)
            If Kerr <> 0 Then
                SndRcvInitPCS = Kerr
                MsgBox "Erreur Clear :" & CStr(Kerr), vbCritical, "Liaison Serveur AS400"
                Exit Function
            End If
        End If
End If

End Function


'---------------------------------------------------------
Function SndRcvXXX()
'---------------------------------------------------------
On Error GoTo SndRcvXXXError
SndRcvXXX = Null
Input #1, As400DtaqLen, As400Dtaq
Call XcomGetBuffer
Exit Function

SndRcvXXXError:
SndRcvXXX = "error"
Exit Function

End Function




'---------------------------------------------------------
Private Sub XcomPutBuffer()
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

End Sub

'---------------------------------------------------------
Private Sub XcomGetBuffer()
'---------------------------------------------------------

With Xcom
    .SrvObj = Mid$(As400Dtaq, 1, 12)
    .SrvMethod = Mid$(As400Dtaq, 13, 12)
    .SrvErr = Mid$(As400Dtaq, 25, 10)
    .usrId = Mid$(As400Dtaq, 35, 10)
    .pcId = Mid$(As400Dtaq, 45, 10)
    .SrvType = Mid$(As400Dtaq, 55, 10)
    .SrvId = Mid$(As400Dtaq, 65, 10)
    .SrvDtaqLib = Mid$(As400Dtaq, 75, 10)
    .SrvDtaqIn = Mid$(As400Dtaq, 85, 10)
    .SrvDTaqOut = Mid$(As400Dtaq, 95, 10)
End With

'''''Elp.usrId = Xcom.usrId
'''''Elp.pcId = Xcom.pcId

End Sub

'---------------------------------------------------------
Public Function SndRcvInit()
'---------------------------------------------------------
Dim X As String

On Error GoTo ErrorX

SndRcvInit = Null

XcomLen = 104

Xcom = Elp

X = "c:\" & Trim(Xcom.SrvDtaqLib) & "\" & Trim(Xcom.SrvDtaqIn) & ".TXT"

If elpSrvTxtOut Then Open X For Output As #1

Select Case elpSrvXcom
    Case Is = "PCS": SndRcvInit = SndRcvInitPCS
    Case Is = "MSCOMM1": SndRcvInit = SndRcvInitMSComm1
   Case ""
    Case Else
        elpSrvTxtin = True
        Open X For Input As #1
        SndRcvInit = SndRcv
End Select

Exit Function

ErrorX:
    MsgBox "Erreur :" & Err & " : " & Error$(Err), vbCritical, "Initialisation : " & X
    SndRcvInit = Err
    Exit Function
End Function

'---------------------------------------------------------
Public Sub XcomEnd()
'---------------------------------------------------------

MsgTxtLen = 0
Xcom.SrvMethod = "ELPDTAQEND  "
SndRcv

End Sub

'---------------------------------------------------------
Public Function SndRcvMSComm1(MSComm1 As MSComm)
'---------------------------------------------------------
Static kOpen As Boolean

SndRcvMSComm1 = Null

If Not kOpen Then: kOpen = True: setMSComm1

sndMSComm1

rcvMSComm1
Call XcomGetBuffer

'MSComm1.PortOpen = False

End Function

'---------------------------------------------------------
Public Sub RcvSndMSComm1(MSComm1 As MSComm)
'---------------------------------------------------------

Kerr = EHNDQ_SetMode(Screen.ActiveForm.hWnd, Elp.SrvId, 1)
If Kerr <> 0 Then
    MsgBox "Erreur SetMode :" & CStr(Kerr), vbCritical, "Liaison Serveur AS400"
    Exit Sub
Else
    setMSComm1
'   do While kCom
    Do
        rcvMSComm1
            
       If frmElp.chkDtaq Then frmElp.lblIN = As400Dtaq: DoEvents
        Call XcomGetBuffer
        Call SndRcvPCS
       If frmElp.chkDtaq Then frmElp.lblOUT = As400Dtaq
        
        sndMSComm1
    Loop

    MSComm1.PortOpen = False

End If

End Sub

'---------------------------------------------------------
Public Function SndRcvInitMSComm1()
'---------------------------------------------------------

SndRcvInitMSComm1 = Null
MsgTxt = Space$(8192 - XcomLen)
MsgTxtLen = 0
Xcom.SrvMethod = "PCINIT      "
    
If IsNull(SndRcv()) Then
   Xcom.SrvMethod = Space$(12)
End If

End Function

'---------------------------------------------------------
Public Sub setMSComm1()
'---------------------------------------------------------

    frmElp.MSComm1.InBufferSize = 512
    frmElp.MSComm1.OutBufferSize = 512
    frmElp.MSComm1.Handshaking = 0
    frmElp.MSComm1.Settings = "19200,N,8,1"
    frmElp.MSComm1.PortOpen = True
End Sub

'---------------------------------------------------------
Public Sub sndMSComm1()
'---------------------------------------------------------
Dim X As String
Dim K As Integer
    
frmElp.MSComm1.Output = Format(As400DtaqLen, "00000") _
               & Mid$(As400Dtaq, 1, 507)

For K = 508 To As400DtaqLen Step 512
    frmElp.MSComm1.InputLen = 0
    Do:: Loop Until frmElp.MSComm1.InBufferCount >= 1
    X = frmElp.MSComm1.Input
    
    frmElp.MSComm1.Output = Mid$(As400Dtaq, K, 512)

Next K

End Sub

'---------------------------------------------------------
Public Sub rcvMSComm1()
'---------------------------------------------------------
Dim X As String
Dim K As Integer


frmElp.MSComm1.InputLen = 0
    Do:: Loop Until frmElp.MSComm1.InBufferCount >= 512     ' : do events :
X = frmElp.MSComm1.Input
As400DtaqLen = Val(Mid$(X, 1, 5))
Mid$(As400Dtaq, 1, 507) = Mid$(X, 6, 507)
    
For K = 508 To As400DtaqLen Step 512
    frmElp.MSComm1.Output = "+"
    frmElp.MSComm1.InputLen = 0
    Do:: Loop Until frmElp.MSComm1.InBufferCount >= 512 ' : do events :
    Mid$(As400Dtaq, K, 512) = Mid$(frmElp.MSComm1.Input, 1, 512)
        
Next K

End Sub


