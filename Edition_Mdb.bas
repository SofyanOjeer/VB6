Attribute VB_Name = "mdbEdition"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableEdition As Recordset
Dim tableEditionOpen As Boolean
Public mEdition_Id As Long

Type typeEdition
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Nature                  As String * 3
    ID                      As String * 20
    Position                As Long
    Memo1                   As Variant
    Memo2                   As Variant

End Type

Type typeEdition_Memo2
    Length                 As String * 6
    FormatCode             As String * 1
    Format                 As String * 10
    textCode               As String * 1
    Text                   As Variant

End Type

Public recEdition As typeEdition

Public frmRTF_Form_K2 As String, frmRTF_blnA5 As Boolean
Public frmRTF_FileName As String
Public frmRTF_Référence As String
Public frmRTF_UsrId_Origine As String
Public frmRTF_recEdition As typeEdition
Public frmRTF_Caller As String
Public frmRTF_Buffer_Name As String

Public frmRTF_blnOK As Boolean
Public frmRTF_blnCourrier As Boolean
Public frmRTF_prtOrientation As Integer
Public frmRTF_prtPaperSize As Integer

Public Function paramEdition_Init(lstErr As ListBox, lcmdContext As CommandButton)
Dim K As Integer, K1 As Integer, X As String

Dim V
paramEdition_Init = Null

recElpTable_Init recElpTable
recElpTable.Method = "Seek="
recElpTable.ID = "Edition"
Call lstErr_Clear(lstErr, lcmdContext, "BIA820I.mdb : table : " & recElpTable.ID)

recElpTable.K1 = "Splf"
recElpTable.K2 = "Folder"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEditionSplf_Folder = paramServer(recElpTable.Memo)
Call lstErr_AddItem(lstErr, lcmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Filigrane"
recElpTable.K2 = "Folder"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEditionFiligrane_Folder = paramServer(recElpTable.Memo)
Call lstErr_AddItem(lstErr, lcmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Courrier"
recElpTable.K2 = "Folder"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEditionCourrier_Folder = paramServer(recElpTable.Memo)
Call lstErr_AddItem(lstErr, lcmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Ftp"
recElpTable.K2 = "File"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEditionFtp_File = paramServer(recElpTable.Memo)
Call lstErr_AddItem(lstErr, lcmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Archive"
recElpTable.K2 = "Folder"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEditionArchive_Folder = paramServer(recElpTable.Memo)
Call lstErr_AddItem(lstErr, lcmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "Corbeille"
recElpTable.K2 = "Folder"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramEditionCorbeille_Folder = paramServer(recElpTable.Memo)
Call lstErr_AddItem(lstErr, lcmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

If blnJPL Then
    paramEditionFtp_File = "C:\Temp\S820i_Out\SPLF\SPLFFTPW0"
    paramEditionSplf_Folder = "C:\Temp\Splf\"
    paramEditionCourrier_Folder = "C:\Temp\Splf\Courrier\"
    paramEditionFiligrane_Folder = "C:\Temp\Filigrane\"
    paramEditionArchive_Folder = "C:\Temp\Splf\Archive\"
    paramEditionCorbeille_Folder = "C:\Temp\Splf\Corbeille\"
End If


Exit Function

Table_Error:
paramEdition_Init = V
Exit Function

Memo_Error:
paramEdition_Init = "Memo"
MsgBox recElpTable.ID & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "srvTI.Param_Init"
Exit Function

Num_Error:
paramEdition_Init = "Num"
MsgBox recElpTable.ID & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "srvTI.Param_Init"
End Function





Public Sub fctRTV_Format_Standard(lEdition() As typeEdition, lEdition_Nb As Integer)
Dim xEdition As typeEdition
Dim K As Integer, X As String
Dim wEdition_Memo2 As typeEdition_Memo2, xEdition_Memo2 As typeEdition_Memo2

On Error Resume Next

recEdition_Init xEdition
xEdition.Method = "Seek>="
xEdition.Nature = "RTV"
xEdition.ID = frmRTF_Buffer_Name
Do
    intReturn = tableEdition_Read(xEdition)
    If intReturn = 0 Then
        If Trim(xEdition.ID) <> frmRTF_Buffer_Name Then intReturn = -1
        If xEdition.Nature <> "RTV" Then intReturn = -1
    End If
    If intReturn = 0 Then
    
        For K = 1 To lEdition_Nb
                
            If xEdition.Memo1 = lEdition(K).Memo1 Then
                Call fctRTV_Memo2_GetBuffer(lEdition(K).Memo2, wEdition_Memo2)
                Call fctRTV_Memo2_GetBuffer(xEdition.Memo2, xEdition_Memo2)
                If wEdition_Memo2.FormatCode <> "*" Then
                    wEdition_Memo2.FormatCode = xEdition_Memo2.FormatCode
                    wEdition_Memo2.Format = xEdition_Memo2.Format
                End If
                 If wEdition_Memo2.textCode <> "*" Then
                    wEdition_Memo2.textCode = xEdition_Memo2.textCode
                    wEdition_Memo2.Text = xEdition_Memo2.Text
                End If
                Call fctRTV_Memo2_PutBuffer(lEdition(K).Memo2, wEdition_Memo2)
                    
            End If
        Next K
    End If
    xEdition.Method = "MoveNext"
Loop Until intReturn <> 0


End Sub

Public Sub arrEdition_Load(lEdition() As typeEdition, lEdition_Nb As Integer)
Dim xEdition As typeEdition

xEdition = lEdition(0)
lEdition_Nb = 0
xEdition.Method = "Seek>="
intReturn = tableEdition_Read(xEdition)
xEdition.Method = "MoveNext"
Do
    If intReturn = 0 Then
        If xEdition.Nature <> lEdition(0).Nature Then intReturn = -1
        If xEdition.ID <> lEdition(0).ID Then intReturn = -1
    End If
    If intReturn = 0 Then
        lEdition_Nb = lEdition_Nb + 1
        If lEdition_Nb >= UBound(lEdition) Then ReDim Preserve lEdition(lEdition_Nb + 10)
        lEdition(lEdition_Nb) = xEdition
        intReturn = tableEdition_Read(xEdition)
    End If
                
Loop While intReturn = 0

End Sub

'---------------------------------------------------------
'-----------------------------------------------------
Sub tableEdition_Close()
'-----------------------------------------------------
If tableEditionOpen Then
    tableEdition.Close
    tableEditionOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableEdition_GetBuffer(recEdition As typeEdition)
'---------------------------------------------------------
recEdition.Nature = tableEdition("Nature")
recEdition.ID = tableEdition("Id")
recEdition.Position = tableEdition("Position")
recEdition.Memo1 = tableEdition("Memo1")
recEdition.Memo2 = tableEdition("Memo2")

End Sub


'---------------------------------------------------------
Public Sub fctRTV_Memo2_GetBuffer(lMemo2 As Variant, lEdition_Memo2 As typeEdition_Memo2)
'---------------------------------------------------------
lEdition_Memo2.Length = mId$(lMemo2, 1, 6)
lEdition_Memo2.FormatCode = mId$(lMemo2, 7, 1)
lEdition_Memo2.Format = mId$(lMemo2, 8, 10)
lEdition_Memo2.textCode = mId$(lMemo2, 18, 1)
lEdition_Memo2.Text = mId$(lMemo2, 19, Len(lMemo2) - 18)
End Sub

'---------------------------------------------------------
Public Sub fctRTV_Memo2_PutBuffer(lMemo2 As Variant, lEdition_Memo2 As typeEdition_Memo2)
'---------------------------------------------------------
lMemo2 = lEdition_Memo2.Length & lEdition_Memo2.FormatCode & lEdition_Memo2.Format & lEdition_Memo2.textCode & lEdition_Memo2.Text

End Sub


'-----------------------------------------------------
Sub tableEdition_Open()
'-----------------------------------------------------

If Not tableEditionOpen Then
    Set tableEdition = MDB.OpenRecordset("Edition")
    tableEdition.Index = "PrimaryKey"
    tableEditionOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableEdition_PutBuffer(recEdition As typeEdition)
'---------------------------------------------------------

tableEdition("Nature") = recEdition.Nature
tableEdition("Id") = recEdition.ID
tableEdition("Position") = recEdition.Position
tableEdition("Memo1") = recEdition.Memo1
tableEdition("Memo2") = recEdition.Memo2

End Sub


'---------------------------------------------------------
Public Function tableEdition_Read(recEdition As typeEdition) As Integer
'---------------------------------------------------------

On Error GoTo tableEdition_Read_Error
tableEdition_Read = 0


Select Case Trim(recEdition.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableEdition.Seek "=", recEdition.Nature, recEdition.ID, recEdition.Position
                        If tableEdition.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableEdition.Seek "<=", recEdition.Nature, recEdition.ID, recEdition.Position
                        If tableEdition.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableEdition.Seek ">=", recEdition.Nature, recEdition.ID, recEdition.Position
                        If tableEdition.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableEdition.Seek ">", recEdition.Nature, recEdition.ID, recEdition.Position
                        If tableEdition.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableEdition.MoveNext
                        If tableEdition.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableEdition.MovePrevious
                        If tableEdition.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableEdition.MoveFirst
                        If tableEdition.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableEdition.MoveLast
                        If tableEdition.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recEdition.Method <> "AddNew      " Then
    Call tableEdition_GetBuffer(recEdition)
End If

Exit Function

'---------------------------------------------------------
tableEdition_Read_Error:
'---------------------------------------------------------

    tableEdition_Read = Err
    Resume tableEdition_Read_End

tableEdition_Read_End:

End Function

'---------------------------------------------------------
Public Function tableEdition_Update(recEdition As typeEdition) As Integer
'---------------------------------------------------------

On Error GoTo tableEditionUpdate_Error
tableEdition_Update = 0

Select Case Trim(recEdition.Method)

    Case "AddNew"
                        tableEdition.AddNew
                        Call tableEdition_PutBuffer(recEdition)
                        tableEdition.Update
    Case "Update"
                        tableEdition.Edit
                        Call tableEdition_PutBuffer(recEdition)
                        tableEdition.Update
    Case "Delete"
                        tableEdition.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableEditionUpdate_Error:
'---------------------------------------------------------
    tableEdition_Update = Err
    Resume tableEditionUpdate_End

tableEditionUpdate_End:

End Function








'-----------------------------------------------------
Sub dbEdition_Error(recEdition As typeEdition)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recEdition.Nature & ": " & recEdition.ID & Chr$(13)

Select Case mId$(recEdition.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recEdition.Err & " : " & Error(recEdition.Err): I = vbCritical
End Select

MsgBox Msg, I, "mdbEdition.bas :  ( " & Trim(recEdition.Obj) & " : " & Trim(recEdition.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbEdition_ReadE(recEdition As typeEdition)
'-----------------------------------------------------

dbEdition_ReadE = Null

recEdition.Err = tableEdition_Read(recEdition)
If recEdition.Err > 0 Then

'    If recEdition.Err < 9990 Or recEdition.Err >= 9999 Then
        Call dbEdition_Error(recEdition)
        dbEdition_ReadE = recEdition.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbEdition_Update(recEdition As typeEdition)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbEdition_Update = Null


recEdition.Err = tableEdition_Update(recEdition)

If recEdition.Err <> 0 Then
    Call dbEdition_Error(recEdition)
    dbEdition_Update = recEdition.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recEdition_Init(recEdition As typeEdition)
recEdition.Method = ""
recEdition.Obj = "Edition"
recEdition.Err = ""
recEdition.ID = ""
recEdition.Nature = ""
recEdition.Position = 0
recEdition.Memo1 = ""
recEdition.Memo2 = ""
End Sub





Public Sub frmRFT_Show(lEdition As typeEdition)

frmRTF.Show vbModal

End Sub

Public Function fctRTF_Buffer_Name(lId As String) As String
Dim K1 As Integer
K1 = InStr(1, lId, "_")
If K1 > 0 Then
    fctRTF_Buffer_Name = mId$(lId, 1, K1 - 1) & "$"
    
Else
    fctRTF_Buffer_Name = ""
End If
End Function

Public Sub fctRTV_Save(meRtvEdition() As typeEdition, meRtvEdition_Nb As Integer, saveRtvEdition() As typeEdition, saveRtvEdition_Nb As Integer)
Dim I As Integer
saveRtvEdition_Nb = meRtvEdition_Nb
ReDim saveRtvEdition(saveRtvEdition_Nb + 1)
For I = 1 To saveRtvEdition_Nb
    saveRtvEdition(I) = meRtvEdition(I)
Next I

End Sub
