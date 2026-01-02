Attribute VB_Name = "srvLrCdr"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recLrCdrLen = 394 ' 34 + 360

Type typeLrCdr
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Text                    As String * 360
    
  End Type
    
Public arrLrCdrSuite As Boolean
Public arrLrCdrNb As Integer


Public paramLrCdr_Archive_Proc  As String
Public paramLrCdr_AS400_Trf  As String
Public paramLrCdr_AS400_Ext  As String
Public paramLrCdr_AS400_LrEstd  As String
Public paramLrCdr_AS400_LrCliCli  As String
Public paramLrCdr_BdfSend_Filename  As String
Public paramLrCdr_BdfReceive_Filename  As String

Public paramLrCdr_LrRisque_Filename  As String
Public paramLrCdr_LrRetris_Filename  As String
Public paramLrCdr_LrSgnBnf_Filename As String
Public paramLrCdr_LrClicli_Proc  As String
Public paramLrCdr_LrEstd_FileName  As String
Public paramLrCdr_LrClicli_FileName  As String
''Public paramLrCdr_Msg_FileName  As String
Public paramLrCdr_LrBdfRetour_FileName  As String
Public paramLrCdr_LrBdfAller_FileName  As String
Public paramLrCdr_PrintSopra_220  As String
Public paramLrCdr_PrintSopra_400  As String
Public paramLrCdr_PrintSopra_470  As String
Public paramLrCdr_PrintSopra_490  As String
Public paramLrCdr_PrintSopra_870  As String
Public paramLrCdr_PrintSopra_880  As String

Public Function param_Init()
Dim V
param_Init = Null

recElpTable_Init recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "LucaRisques"
Call lstErr_Clear(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, "BIA.mdb : table : " & recElpTable.Id)


recElpTable.K1 = "Archive"
recElpTable.K2 = "Proc"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_Archive_Proc = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "AS400"

recElpTable.K2 = "Ext"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_AS400_Ext = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "Trf"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_AS400_Trf = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "LrCliCli"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_AS400_LrCliCli = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "LrEstd"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_AS400_LrEstd = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "BdfSend"
recElpTable.K2 = "Filename"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_BdfSend_Filename = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "BdfReceive"
recElpTable.K2 = "Filename"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_BdfReceive_Filename = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))


recElpTable.K1 = "LrBdfAller"
recElpTable.K2 = "Filename"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_LrBdfAller_FileName = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "LrBdfRetour"
recElpTable.K2 = "Filename"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_LrBdfRetour_FileName = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "LrClicli"
recElpTable.K2 = "Filename"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_LrClicli_FileName = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "LrClicli"
recElpTable.K2 = "Proc"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_LrClicli_Proc = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "LrEstd"
recElpTable.K2 = "Filename"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_LrEstd_FileName = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "LrRetris"
recElpTable.K2 = "Filename"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_LrRetris_Filename = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "LrRisque"
recElpTable.K2 = "Filename"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_LrRisque_Filename = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "LrSgnBnf"
recElpTable.K2 = "Filename"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_LrSgnBnf_Filename = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "PrintSopra"
recElpTable.K2 = "220"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_PrintSopra_220 = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "PrintSopra"
recElpTable.K2 = "400"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_PrintSopra_400 = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "PrintSopra"
recElpTable.K2 = "470"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_PrintSopra_470 = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "PrintSopra"
recElpTable.K2 = "490"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_PrintSopra_490 = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "PrintSopra"
recElpTable.K2 = "870"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_PrintSopra_870 = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "PrintSopra"
recElpTable.K2 = "880"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramLrCdr_PrintSopra_880 = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

Call lstErr_Clear(frmLucaRisques.lstErr, frmLucaRisques.cmdActualiser, "BIA.mdb : table : " & recElpTable.Id & ": ok ")


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
Public Function Monitor(recLrCdr As typeLrCdr)
'-----------------------------------------------------

arrLrCdrSuite = False
Select Case mId$(Trim(recLrCdr.Method), 1, 4)
    Case "Snap"
              Monitor = Snap(recLrCdr)
    Case Else
                recLrCdr.Err = recLrCdr.Method
                Call ErrorX(recLrCdr)
                Monitor = recLrCdr.Err
End Select

End Function

'-----------------------------------------------------
Sub ErrorX(recLrCdr As typeLrCdr)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "LrCdr: "

Select Case mId$(recLrCdr.Err, 9, 2)
    Case "22"
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recLrCdr.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvLrCdr.bas  ( " _
                & Trim(recLrCdr.obj) & " : " & Trim(recLrCdr.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function GetBuffer(recLrCdr As typeLrCdr)
'---------------------------------------------------------
Dim K As Integer, I As Integer
GetBuffer = Null
recLrCdr.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recLrCdr.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recLrCdr.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recLrCdr.Err = Space$(10) Then
    recLrCdr.Text = mId$(MsgTxt, K + 1, 360)
Else
    GetBuffer = recLrCdr.Err
End If

MsgTxtIndex = MsgTxtIndex + recLrCdrLen

End Function

'---------------------------------------------------------
Private Sub PutBuffer(recLrCdr As typeLrCdr)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recLrCdr.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recLrCdr.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 360) = recLrCdr.Text
MsgTxtLen = MsgTxtLen + recLrCdrLen
End Sub



'---------------------------------------------------------
Private Function Snap(recLrCdr As typeLrCdr)
'---------------------------------------------------------
Dim I As Integer, x As String
Snap = "?"
MsgTxtLen = 0
Call PutBuffer(recLrCdr)
'Call PutBuffer(arrLrCdr(0))
If IsNull(SndRcv()) Then
    Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(GetBuffer(recLrCdr)) Then
            arrLrCdrNb = arrLrCdrNb + 1
            Print #1, recLrCdr.Text
            arrLrCdrSuite = True
        Else
            arrLrCdrSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub Init(recLrCdr As typeLrCdr)
'---------------------------------------------------------
MsgTxt = Space$(recLrCdrLen)
MsgTxtIndex = 0
Call GetBuffer(recLrCdr)
recLrCdr.obj = "SRVLRCDR"
End Sub

