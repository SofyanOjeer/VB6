Attribute VB_Name = "mdbElpKMInfo"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableElpKMInfo As Recordset
Dim tableElpKMInfoOpen As Boolean
Public mElpKMInfo_Id As Long

Type typeElpKMInfo
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    ElpKMSrc_Id             As Long
    Id                      As String * 20
    Description             As String * 40
    Pass                    As Long
    Memo                    As Variant

End Type

Public recElpKMInfo As typeElpKMInfo


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableElpKMInfo_Close()
'-----------------------------------------------------
If tableElpKMInfoOpen Then
    tableElpKMInfo.Close
    tableElpKMInfoOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableElpKMInfo_GetBuffer(recElpKMInfo As typeElpKMInfo)
'---------------------------------------------------------
recElpKMInfo.ElpKMSrc_Id = tableElpKMInfo("ElpKMSrc_Id")
recElpKMInfo.Id = tableElpKMInfo("Id")
recElpKMInfo.Description = tableElpKMInfo("Description")
recElpKMInfo.Pass = tableElpKMInfo("Pass")
recElpKMInfo.Memo = tableElpKMInfo("Memo")

End Sub


'-----------------------------------------------------
Sub tableElpKMInfo_Open()
'-----------------------------------------------------

If Not tableElpKMInfoOpen Then
    Set tableElpKMInfo = MDB.OpenRecordset("ElpKMInfo")
    tableElpKMInfo.Index = "PrimaryKey"
    tableElpKMInfoOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableElpKMInfo_PutBuffer(recElpKMInfo As typeElpKMInfo)
'---------------------------------------------------------

tableElpKMInfo("ElpKMSrc_Id") = recElpKMInfo.ElpKMSrc_Id
tableElpKMInfo("Id") = recElpKMInfo.Id
tableElpKMInfo("Description") = recElpKMInfo.Description
tableElpKMInfo("Pass") = recElpKMInfo.Pass
tableElpKMInfo("Memo") = recElpKMInfo.Memo
End Sub


'---------------------------------------------------------
Public Function tableElpKMInfo_Read(recElpKMInfo As typeElpKMInfo) As Integer
'---------------------------------------------------------

On Error GoTo tableElpKMInfo_Read_Error
tableElpKMInfo_Read = 0


Select Case Trim(recElpKMInfo.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableElpKMInfo.Seek "=", recElpKMInfo.ElpKMSrc_Id, recElpKMInfo.Id
                        If tableElpKMInfo.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableElpKMInfo.Seek "<=", recElpKMInfo.ElpKMSrc_Id, recElpKMInfo.Id
                        If tableElpKMInfo.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableElpKMInfo.Seek ">=", recElpKMInfo.ElpKMSrc_Id, recElpKMInfo.Id
                        If tableElpKMInfo.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableElpKMInfo.Seek ">", recElpKMInfo.ElpKMSrc_Id, recElpKMInfo.Id
                        If tableElpKMInfo.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableElpKMInfo.MoveNext
                        If tableElpKMInfo.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableElpKMInfo.MovePrevious
                        If tableElpKMInfo.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableElpKMInfo.MoveFirst
                        If tableElpKMInfo.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableElpKMInfo.MoveLast
                        If tableElpKMInfo.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recElpKMInfo.Method <> "AddNew      " Then
    Call tableElpKMInfo_GetBuffer(recElpKMInfo)
End If

Exit Function

'---------------------------------------------------------
tableElpKMInfo_Read_Error:
'---------------------------------------------------------

    tableElpKMInfo_Read = Err
    Resume tableElpKMInfo_Read_End

tableElpKMInfo_Read_End:

End Function

'---------------------------------------------------------
Public Function tableElpKMInfo_Update(recElpKMInfo As typeElpKMInfo) As Integer
'---------------------------------------------------------

On Error GoTo tableElpKMInfoUpdate_Error
tableElpKMInfo_Update = 0

Select Case Trim(recElpKMInfo.Method)

    Case "AddNew"
                        tableElpKMInfo.AddNew
                        Call tableElpKMInfo_PutBuffer(recElpKMInfo)
                        tableElpKMInfo.Update
    Case "Update"
                        tableElpKMInfo.Edit
                        Call tableElpKMInfo_PutBuffer(recElpKMInfo)
                        tableElpKMInfo.Update
    Case "Delete"
                        tableElpKMInfo.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableElpKMInfoUpdate_Error:
'---------------------------------------------------------
    tableElpKMInfo_Update = Err
    Resume tableElpKMInfoUpdate_End

tableElpKMInfoUpdate_End:

End Function








'-----------------------------------------------------
Sub dbElpKMInfo_Error(recElpKMInfo As typeElpKMInfo)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recElpKMInfo.Id & ": " & recElpKMInfo.ElpKMSrc_Id & Chr$(13)

Select Case mId$(recElpKMInfo.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recElpKMInfo.Err: I = vbCritical
End Select

'MsgBox Msg, I, "mdbElpKMInfo.bas :  ( " & Trim(recElpKMInfo.obj) & " : " & Trim(recElpKMInfo.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbElpKMInfo_ReadE(recElpKMInfo As typeElpKMInfo)
'-----------------------------------------------------

dbElpKMInfo_ReadE = Null

recElpKMInfo.Err = tableElpKMInfo_Read(recElpKMInfo)
If recElpKMInfo.Err > 0 Then

'    If recElpKMInfo.Err < 9990 Or recElpKMInfo.Err >= 9999 Then
        Call dbElpKMInfo_Error(recElpKMInfo)
        dbElpKMInfo_ReadE = recElpKMInfo.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbElpKMInfo_Update(recElpKMInfo As typeElpKMInfo)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbElpKMInfo_Update = Null


recElpKMInfo.Err = tableElpKMInfo_Update(recElpKMInfo)

If recElpKMInfo.Err <> 0 Then
    Call dbElpKMInfo_Error(recElpKMInfo)
    dbElpKMInfo_Update = recElpKMInfo.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recElpKMInfo_Init(recElpKMInfo As typeElpKMInfo)
recElpKMInfo.Method = ""
recElpKMInfo.obj = "ElpKMInfo"
recElpKMInfo.Err = ""
recElpKMInfo.Id = ""
recElpKMInfo.ElpKMSrc_Id = 0
recElpKMInfo.Description = ""
recElpKMInfo.Pass = 0
recElpKMInfo.Memo = Null
End Sub

