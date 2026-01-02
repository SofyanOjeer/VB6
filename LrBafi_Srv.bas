Attribute VB_Name = "srvLrBafi"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recLrBafiLen = 707 ' 34 + 673

Type typeLrBafi
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Text                    As String * 673
    
  End Type
    
Public arrLrBafiSuite As Boolean
Public arrLrBafiNb As Integer


'____________________________________________EvolanReport
Public paramErBafi_AS400Trf As String
Public paramErBafi_AS400Ext As String
Public paramErBafi_FTP_LrBafiMsg  As String
Public paramErBafi_FTP_LrBafi  As String
Public paramErBafi_FTP_LrSolde As String

Public paramErBafi_Engine_Folder As String
Public paramErBafi_Engine_Start As String
Public paramErBafi_Engine_End As String
Public paramErBafi_Estd_FileName As String
Public paramErBafi_Solde_FileName As String
Public paramErBafi_Msg_FileName As String
Public paramErBafi_PilFab_FileName As String
Public paramErBafi_Descri_FileName As String
Public paramErBafi_Archive As String
Public paramErBafi_Emission As String
Public paramErBafi_Out_Folder As String

Public Function param_Init()
Dim V
param_Init = Null

recElpTable_Init recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "EvolanReport"
Call lstErr_Clear(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, "BIA.mdb : table : " & recElpTable.Id)

recElpTable.K1 = "AS400"

recElpTable.K2 = "Ext"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_AS400Ext = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "Trf"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_AS400Trf = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "FTP_LrBafiMs"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_FTP_LrBafiMsg = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "FTP_LrBafi"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_FTP_LrBafi = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "FTP_LrSolde"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_FTP_LrSolde = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))
recElpTable.K1 = "Out"
recElpTable.K2 = "Folder"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_Out_Folder = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Archive"
recElpTable.K2 = "Folder"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_Archive = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Emission"
recElpTable.K2 = "Folder"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_Emission = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Engine"
recElpTable.K2 = "Folder"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_Engine_Folder = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))


recElpTable.K1 = "Engine"
recElpTable.K2 = "Start"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_Engine_Start = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Engine"
recElpTable.K2 = "Stop"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_Engine_End = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Estd"
recElpTable.K2 = "FileName"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_Estd_FileName = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Solde"
recElpTable.K2 = "FileName"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_Solde_FileName = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Msg"
recElpTable.K2 = "FileName"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_Msg_FileName = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "PilFab"
recElpTable.K2 = "FileName"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_PilFab_FileName = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Descri"
recElpTable.K2 = "FileName"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramErBafi_Descri_FileName = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

'2000-12-18 Public Const constLrBafi_Bia_Filename = "BalPa*.*;Cad*.*;IME*.*;Inter*.*;Sit*.*"

Call lstErr_Clear(frmLrBafi.lstErr, frmLrBafi.cmdActualiser, "BIA.mdb : table : " & recElpTable.Id & ": ok ")


Exit Function

Table_Error:
param_Init = V
Exit Function

Memo_Error:
param_Init = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "srvTI.Param_Init"
Exit Function

Num_Error:
param_Init = "Num"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "srvTI.Param_Init"
End Function


'-----------------------------------------------------
Public Function Monitor(recLrBafi As typeLrBafi)
'-----------------------------------------------------

arrLrBafiSuite = False
Select Case mId$(Trim(recLrBafi.Method), 1, 4)
    Case "Snap"
              Monitor = Snap(recLrBafi)
    Case Else
                recLrBafi.Err = recLrBafi.Method
                Call ErrorX(recLrBafi)
                Monitor = recLrBafi.Err
End Select

End Function

'-----------------------------------------------------
Sub ErrorX(recLrBafi As typeLrBafi)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "LrBafi: "

Select Case mId$(recLrBafi.Err, 9, 2)
    Case "22"
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recLrBafi.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvLrBafi.bas  ( " _
                & Trim(recLrBafi.obj) & " : " & Trim(recLrBafi.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function GetBuffer(recLrBafi As typeLrBafi)
'---------------------------------------------------------
Dim K As Integer, I As Integer
GetBuffer = Null
recLrBafi.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recLrBafi.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recLrBafi.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recLrBafi.Err = Space$(10) Then
    recLrBafi.Text = mId$(MsgTxt, K + 1, 673)
Else
    GetBuffer = recLrBafi.Err
End If

MsgTxtIndex = MsgTxtIndex + recLrBafiLen

End Function

'---------------------------------------------------------
Private Sub PutBuffer(recLrBafi As typeLrBafi)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recLrBafi.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recLrBafi.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 673) = recLrBafi.Text
MsgTxtLen = MsgTxtLen + recLrBafiLen
End Sub



'---------------------------------------------------------
Private Function Snap(recLrBafi As typeLrBafi)
'---------------------------------------------------------
Dim I As Integer, x As String
Snap = "?"
MsgTxtLen = 0
Call PutBuffer(recLrBafi)
'Call PutBuffer(arrLrBafi(0))
If IsNull(SndRcv()) Then
    Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(GetBuffer(recLrBafi)) Then
            arrLrBafiNb = arrLrBafiNb + 1
            Print #1, recLrBafi.Text
            arrLrBafiSuite = True
        Else
            arrLrBafiSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub Init(recLrBafi As typeLrBafi)
'---------------------------------------------------------
MsgTxt = Space$(recLrBafiLen)
MsgTxtIndex = 0
Call GetBuffer(recLrBafi)
recLrBafi.obj = "SRVLRBAFI"
End Sub


Public Sub PeliNT_Emission(lMsg As String)
Dim x As String, xFileName As String, Y As String, vShellId, V
If Trim(paramErBafi_Emission) = "" Then
    recElpTable_Init recElpTable
    recElpTable.Method = "Seek="
    recElpTable.Id = "EvolanReport"
    recElpTable.K1 = "Emission"
    recElpTable.K2 = "Folder"
    V = dbElpTable_ReadE(recElpTable)
    If Not IsNull(V) Then Exit Sub
    If IsNull(recElpTable.Memo) Then Exit Sub
    paramErBafi_Emission = paramServer(recElpTable.Memo)
End If

x = paramErBafi_Emission & "\*.*"
xFileName = Dir(x)
If xFileName <> "" Then
    x = mId$(lMsg, 14, Len(Trim(lMsg)) - 13) & " " & xFileName
    
    ''''Y = "net send fr81815996  envoi CB : " & X: vShellId = Shell(Y, 1)
    
    vShellId = Shell(x, 1)
    AppActivate vShellId
    DoEvents
    SendKeys "%{F4}", True

End If


End Sub
