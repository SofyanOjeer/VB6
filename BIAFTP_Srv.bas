Attribute VB_Name = "srvBIAFTP"
Option Explicit
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Lx As Long, Vx As Variant
Public msFileSystem, msFile
Public srvIdle As Boolean
Public Elp As typeXcom
Public SrvDir As String
Public usrId As String
Public usrName As String
Public DSys As String * 8

Public paramFTP_SbmJob  As String, paramFTP_Dta  As String, paramFTP_Tmp As String

'---------------------------------------------------------
Public Sub Main()
'---------------------------------------------------------
Dim I As Integer
Dim X As String, X2 As String
Dim I1 As Integer, I2 As Integer
Dim xErr As String

On Error GoTo ErrorX

Lx = 25: X = Space(25)
Vx = GetUserName(X, Lx)
usrName = Mid$(X, 1, Lx - 1)
If UCase$(usrName) = "ADMINISTRATEUR" Then usrName = "BIA_INFO"
srvIdle = True


DSys = Year(Now)
Mid$(DSys, 5, 2) = Format$(Month(Now), "00")
Mid$(DSys, 7, 2) = Format$(Day(Now), "00")


elpSrvTxtin = False
elpSrvTxtOut = False
elpSrvXcom = ""
Elp.SrvObj = "ELPDTAQ"
Elp.pcId = "FR"
Elp.SrvType = "AS400"
Elp.SrvId = "S44H1212"
Elp.SrvDtaqLib = "BIADTAQ"
Elp.SrvDtaqIn = "PC000001"
Elp.SrvDTaqOut = "PC000000"
elpSrvXcom = "CAV4"
 
paramFTP_Dta = ""
paramFTP_Tmp = ""
paramFTP_SbmJob = ""

SrvDir = ""
X = Trim(Command)
I = Len(X)
For I1 = I To 1 Step -1
    If Mid$(X, I1, 1) = "\" Then SrvDir = Mid$(X, 1, I1): Exit For
Next I1
X = Trim(Command)
If X <> "" Then
    Open X For Input As #1
    
    Do While Not EOF(1)
        
        Line Input #1, X
        I1 = InStr(1, X, Chr$(34))
        If I1 > 0 Then
            I2 = InStr(I1 + 1, X, Chr$(34))
            X2 = Trim(UCase$(Mid$(X, I1 + 1, I2 - I1 - 1)))
            Select Case UCase$(Mid$(X, 1, I1 - 1))
                Case "FTPAS400.CL=": paramFTP_SbmJob = X2
                Case "FTPFILE.TMP=": paramFTP_Tmp = X2
                Case "FTPFILE.DTA=": paramFTP_Dta = X2
               
                Case "SRVDTAQLIB=": Elp.SrvDtaqLib = X2
                Case "SRVDTAQIN=": Elp.SrvDtaqIn = X2
                Case "SRVDTAQOUT=": Elp.SrvDTaqOut = X2
                Case "SRVTXTOUT=": If X2 = "OUI" Then elpSrvTxtOut = True
                Case "SRVXCOM=":  elpSrvXcom = X2
            End Select
        End If
    Loop
    
    Close #1
End If

If Trim(paramFTP_Dta) = "" Then xErr = "Préciser le fichier origine": GoTo ErrorX2
If Trim(paramFTP_Tmp) = "" Then xErr = "Préciser le fichier FTP temporaire": GoTo ErrorX2
If Trim(paramFTP_SbmJob) = "" Then xErr = "Préciser le nom du CL AS400": GoTo ErrorX2

Elp.SrvDTaqLen = "00000"
Elp.jplFree = "00000"
Elp.usrId = UCase$(usrName)
usrId = Elp.usrId
Set msFileSystem = CreateObject("Scripting.FileSystemObject")
If Not IsNull(SndRcv_Init) Then
    Stop: End
Else
    Call FTP_Get(paramFTP_SbmJob, paramFTP_Dta, paramFTP_Tmp)
    End
End If

Exit Sub

ErrorX:
    xErr = "Erreur :" & Err & " : " & Error$(Err)
ErrorX2:
    MsgBox xErr, vbCritical, "srvBiaFTP.Main : " & X
    Stop: End
End Sub



