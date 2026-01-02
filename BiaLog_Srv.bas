Attribute VB_Name = "srvBiaLog"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recBiaLogLen = 412 ' 34 + 378
Public Const recBiaLog_Block = 50

Type typeBiaLog
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Log_Cosoc             As String * 3
    Log_Agence            As String * 3
    
    Log_CptAmj            As String * 8
    Log_Cpteur            As Long
    Log_Servic            As String * 5
    
    Log_Progr             As String * 20
    Log_Profil            As String * 20
    Log_RefCon            As String * 16
    Log_Devise            As String * 3
   
    Log_Compte           As String * 11
    Log_CodErr           As String * 12
    Log_Texte1           As String * 128
    Log_Texte2           As String * 128
    
    Log_SysAmj           As String * 8
    Log_SysHms           As String * 6
    
End Type
    
Public arrBiaLog() As typeBiaLog
Public arrBiaLog_NB As Integer
Public arrBiaLog_NBMax As Integer
Public arrBiaLog_Index As Integer
Public arrBiaLog_Suite As Boolean

Public xBiaLog As typeBiaLog

Public Sub srvBiaLog_ElpDisplay(recBiaLog As typeBiaLog)
frmElpDisplay.fgData.Rows = 24
frmElpDisplay.fgData.Row = 1
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "obj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.obj
frmElpDisplay.fgData.Row = 2
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Method"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Method
frmElpDisplay.fgData.Row = 3
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Err"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Err
frmElpDisplay.fgData.Row = 4
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_Cosoc"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_Cosoc
frmElpDisplay.fgData.Row = 5
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_Agence"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_Agence
frmElpDisplay.fgData.Row = 6
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_CptAmj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_CptAmj
frmElpDisplay.fgData.Row = 7
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_Cpteur"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_Cpteur
frmElpDisplay.fgData.Row = 8
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_Servic"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_Servic
frmElpDisplay.fgData.Row = 9
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_Progr"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_Progr
frmElpDisplay.fgData.Row = 10
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_Profil"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_Profil
frmElpDisplay.fgData.Row = 11
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_RefCon"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_RefCon
frmElpDisplay.fgData.Row = 12
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_Devise"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_Devise
frmElpDisplay.fgData.Row = 13
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_Compte"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_Compte
frmElpDisplay.fgData.Row = 14
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_CodErr"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_CodErr
frmElpDisplay.fgData.Row = 15
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_Texte1"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_Texte1
frmElpDisplay.fgData.Row = 16
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_Texte2"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_Texte2
frmElpDisplay.fgData.Row = 17
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_SysAmj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_SysAmj
frmElpDisplay.fgData.Row = 18
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Log_SysHms"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recBiaLog.Log_SysHms

frmElpDisplay.Show vbModal

End Sub



'-----------------------------------------------------
Function srvBiaLog_Update(recBiaLog As typeBiaLog)
'-----------------------------------------------------

srvBiaLog_Update = "?"

MsgTxtLen = 0
Call srvBiaLog_PutBuffer(recBiaLog)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvBiaLog_GetBuffer(recBiaLog)) Then
        Call srvBiaLog_Error(recBiaLog)
        srvBiaLog_Update = recBiaLog.Err
        Exit Function
    Else
        srvBiaLog_Update = Null
    End If
Else
    recBiaLog.Err = "srv"
End If


'=====================================================
End Function



Public Sub srvBiaLog_Load(recBiaLogMin As typeBiaLog, recBiaLogMax As typeBiaLog)
Dim mMethod As String

mMethod = Trim(recBiaLogMin.Method) & "+"
arrBiaLog_NBMax = 0
arrBiaLog_Suite = True: arrBiaLog_NB = 0
arrBiaLog_NBMax = recBiaLog_Block: ReDim arrBiaLog(arrBiaLog_NBMax)

arrBiaLog(0) = recBiaLogMax
arrBiaLog_Suite = True
Do Until Not arrBiaLog_Suite
    srvBiaLog_Monitor recBiaLogMin
    recBiaLogMin = arrBiaLog(arrBiaLog_NB)
    recBiaLogMin.Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Function srvBiaLog_Dtaq_Put(lFct As String, recBiaLog As typeBiaLog)
'-----------------------------------------------------

srvBiaLog_Dtaq_Put = Null
Select Case lFct
    Case "Init": MsgTxtLen = 0
    Case "Add": Call srvBiaLog_PutBuffer(recBiaLog)
                If MsgTxtLen + recBiaLogLen >= recBiaLog_Block * recBiaLogLen Then
                    Call srvBiaLog_Dtaq_Snd(recBiaLog): MsgTxtLen = 0
                End If
    Case "Snd": If MsgTxtLen > 0 Then Call srvBiaLog_Dtaq_Snd(recBiaLog)
    Case Else: srvBiaLog_Dtaq_Put = lFct
End Select
'=====================================================
End Function


'-----------------------------------------------------
Function srvBiaLog_Dtaq_Snd(recBiaLog As typeBiaLog)
'-----------------------------------------------------

srvBiaLog_Dtaq_Snd = "?"

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvBiaLog_GetBuffer(recBiaLog)) Then
        Call srvBiaLog_Error(recBiaLog)
        srvBiaLog_Dtaq_Snd = recBiaLog.Err
        Exit Function
    Else
        srvBiaLog_Dtaq_Snd = Null
    End If
Else
    recBiaLog.Err = "Snd"
End If


'=====================================================
End Function



'-----------------------------------------------------
Public Function srvBiaLog_Monitor(recBiaLog As typeBiaLog)
'-----------------------------------------------------
blnFR_Convert = False

arrBiaLog_Suite = False
Select Case mId$(Trim(recBiaLog.Method), 1, 4)
    Case "Seek", "Comp"
                srvBiaLog_Monitor = srvBiaLog_Seek(recBiaLog)
    Case "Snap"
              srvBiaLog_Monitor = srvBiaLog_Snap(recBiaLog)
    Case Else
                recBiaLog.Err = recBiaLog.Method
                Call srvBiaLog_Error(recBiaLog)
                srvBiaLog_Monitor = recBiaLog.Err
End Select

End Function

'-----------------------------------------------------
Sub srvBiaLog_Error(recBiaLog As typeBiaLog)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "BiaLog" & Chr$(10) & Chr$(13)

Select Case mId$(recBiaLog.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recBiaLog.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recBiaLog.Log_CptAmj & " : " & recBiaLog.Log_Cpteur _
        , I, "module : BiaLogs.bas  ( " & Trim(recBiaLog.obj) & " : " & Trim(recBiaLog.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvBiaLog_GetBuffer(recBiaLog As typeBiaLog)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvBiaLog_GetBuffer = Null
recBiaLog.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recBiaLog.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recBiaLog.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recBiaLog.Err = Space$(10) Then
    recBiaLog.Log_Cosoc = mId$(MsgTxt, K + 1, 3)
    recBiaLog.Log_Agence = mId$(MsgTxt, K + 4, 3)
    
    recBiaLog.Log_CptAmj = mId$(MsgTxt, K + 7, 8)
    recBiaLog.Log_Cpteur = CLng(Val(mId$(MsgTxt, K + 15, 7)))
    recBiaLog.Log_Servic = mId$(MsgTxt, K + 22, 5)
    
    recBiaLog.Log_Progr = mId$(MsgTxt, K + 27, 20)
    recBiaLog.Log_Profil = mId$(MsgTxt, K + 47, 20)
    recBiaLog.Log_RefCon = mId$(MsgTxt, K + 67, 16)
    recBiaLog.Log_Devise = mId$(MsgTxt, K + 83, 3)
  
    recBiaLog.Log_Compte = mId$(MsgTxt, K + 86, 11)
    recBiaLog.Log_CodErr = mId$(MsgTxt, K + 97, 12)
    recBiaLog.Log_Texte1 = mId$(MsgTxt, K + 109, 128)
    recBiaLog.Log_Texte2 = mId$(MsgTxt, K + 237, 128)
    recBiaLog.Log_SysAmj = mId$(MsgTxt, K + 365, 8)
    recBiaLog.Log_SysHms = mId$(MsgTxt, K + 373, 6)

Else
    srvBiaLog_GetBuffer = recBiaLog.Err
End If

MsgTxtIndex = MsgTxtIndex + recBiaLogLen

End Function

'---------------------------------------------------------
Private Sub srvBiaLog_PutBuffer(recBiaLog As typeBiaLog)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recBiaLog.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recBiaLog.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 3) = recBiaLog.Log_Cosoc
    Mid$(MsgTxt, K + 4, 3) = recBiaLog.Log_Agence
    
    Mid$(MsgTxt, K + 7, 8) = Format$(recBiaLog.Log_CptAmj, "00000000")
    Mid$(MsgTxt, K + 15, 7) = Format$(recBiaLog.Log_Cpteur, "0000000")
    Mid$(MsgTxt, K + 22, 5) = recBiaLog.Log_Servic
    
    Mid$(MsgTxt, K + 27, 20) = recBiaLog.Log_Progr
    Mid$(MsgTxt, K + 47, 20) = recBiaLog.Log_Profil
    Mid$(MsgTxt, K + 67, 16) = recBiaLog.Log_RefCon
    Mid$(MsgTxt, K + 83, 3) = recBiaLog.Log_Devise
  
    Mid$(MsgTxt, K + 86, 11) = recBiaLog.Log_Compte
    Mid$(MsgTxt, K + 97, 12) = recBiaLog.Log_CodErr
    Mid$(MsgTxt, K + 109, 128) = recBiaLog.Log_Texte1
    Mid$(MsgTxt, K + 237, 128) = recBiaLog.Log_Texte2
    Mid$(MsgTxt, K + 365, 8) = recBiaLog.Log_SysAmj
    Mid$(MsgTxt, K + 373, 6) = recBiaLog.Log_SysHms


MsgTxtLen = MsgTxtLen + recBiaLogLen


  
End Sub



'---------------------------------------------------------
Private Function srvBiaLog_Seek(recBiaLog As typeBiaLog)
'---------------------------------------------------------

srvBiaLog_Seek = "?"
MsgTxtLen = 0
Call srvBiaLog_PutBuffer(recBiaLog)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvBiaLog_GetBuffer(recBiaLog)) Then
        srvBiaLog_Seek = Null
    Else
       '' Call srvBiaLog_Error(recBiaLog)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvBiaLog_Snap(recBiaLog As typeBiaLog)
'---------------------------------------------------------
srvBiaLog_Snap = "?"
MsgTxtLen = 0
Call srvBiaLog_PutBuffer(recBiaLog)
Call srvBiaLog_PutBuffer(arrBiaLog(0))
If IsNull(SndRcv()) Then
    srvBiaLog_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvBiaLog_GetBuffer(recBiaLog)) Then
            Call arrBiaLog_AddItem(recBiaLog)
            arrBiaLog_Suite = True
        Else
            arrBiaLog_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recBiaLog_Init(recBiaLog As typeBiaLog)
'---------------------------------------------------------
MsgTxt = Space$(recBiaLogLen)
MsgTxtIndex = 0
Call srvBiaLog_GetBuffer(recBiaLog)
recBiaLog.obj = "SRVBIALOG    "
recBiaLog.Log_CptAmj = "00000000"
recBiaLog.Log_SysAmj = "00000000"
recBiaLog.Log_SysHms = "000000"
End Sub

'---------------------------------------------------------
Public Sub arrBiaLog_AddItem(recBiaLog As typeBiaLog)
'---------------------------------------------------------
          
arrBiaLog_NB = arrBiaLog_NB + 1
    
If arrBiaLog_NB > arrBiaLog_NBMax Then
    arrBiaLog_NBMax = arrBiaLog_NBMax + recBiaLog_Block
    ReDim Preserve arrBiaLog(arrBiaLog_NBMax)
End If
            
arrBiaLog(arrBiaLog_NB) = recBiaLog
End Sub


