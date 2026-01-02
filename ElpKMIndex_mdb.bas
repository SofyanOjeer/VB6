Attribute VB_Name = "mdbElpKMIndex"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableElpKMIndex As Recordset
Dim tableElpKMIndexOpen As Boolean
Public mElpKMIndex_Id As Long

Type typeElpKMIndex
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Id                      As String * 16
    Classe                  As Long
    ElpKMSrc_Id             As Long             'As String * 20
    Memo                    As Variant

End Type

Public recElpKMIndex As typeElpKMIndex


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableElpKMIndex_Close()
'-----------------------------------------------------
If tableElpKMIndexOpen Then
    tableElpKMIndex.Close
    tableElpKMIndexOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableElpKMIndex_GetBuffer(recElpKMIndex As typeElpKMIndex)
'---------------------------------------------------------
recElpKMIndex.Id = tableElpKMIndex("Id")
recElpKMIndex.Classe = tableElpKMIndex("Classe")
recElpKMIndex.ElpKMSrc_Id = tableElpKMIndex("ElpKMSrc_Id")
recElpKMIndex.Memo = tableElpKMIndex("Memo")

End Sub


'-----------------------------------------------------
Sub tableElpKMIndex_Open()
'-----------------------------------------------------

If Not tableElpKMIndexOpen Then
    Set tableElpKMIndex = MDB.OpenRecordset("ElpKMIndex")
    tableElpKMIndex.Index = "PrimaryKey"
    tableElpKMIndexOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableElpKMIndex_PutBuffer(recElpKMIndex As typeElpKMIndex)
'---------------------------------------------------------
tableElpKMIndex("Id") = recElpKMIndex.Id
tableElpKMIndex("Classe") = recElpKMIndex.Classe
tableElpKMIndex("ElpKMSrc_Id") = recElpKMIndex.ElpKMSrc_Id
tableElpKMIndex("Memo") = recElpKMIndex.Memo
End Sub


'---------------------------------------------------------
Public Function tableElpKMIndex_Read(recElpKMIndex As typeElpKMIndex) As Integer
'---------------------------------------------------------

On Error GoTo tableElpKMIndex_Read_Error
tableElpKMIndex_Read = 0


Select Case Trim(recElpKMIndex.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableElpKMIndex.Seek "=", recElpKMIndex.Id, recElpKMIndex.Classe, recElpKMIndex.ElpKMSrc_Id
                        If tableElpKMIndex.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableElpKMIndex.Seek "<=", recElpKMIndex.Id, recElpKMIndex.Classe, recElpKMIndex.ElpKMSrc_Id
                        If tableElpKMIndex.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableElpKMIndex.Seek ">=", recElpKMIndex.Id, recElpKMIndex.Classe, recElpKMIndex.ElpKMSrc_Id
                        If tableElpKMIndex.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableElpKMIndex.Seek ">", recElpKMIndex.Id, recElpKMIndex.Classe, recElpKMIndex.ElpKMSrc_Id
                        If tableElpKMIndex.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableElpKMIndex.MoveNext
                        If tableElpKMIndex.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableElpKMIndex.MovePrevious
                        If tableElpKMIndex.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableElpKMIndex.MoveFirst
                        If tableElpKMIndex.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableElpKMIndex.MoveLast
                        If tableElpKMIndex.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recElpKMIndex.Method <> "AddNew      " Then
    Call tableElpKMIndex_GetBuffer(recElpKMIndex)
End If

Exit Function

'---------------------------------------------------------
tableElpKMIndex_Read_Error:
'---------------------------------------------------------

    tableElpKMIndex_Read = Err
    Resume tableElpKMIndex_Read_End

tableElpKMIndex_Read_End:

End Function

'---------------------------------------------------------
Public Function tableElpKMIndex_Update(recElpKMIndex As typeElpKMIndex) As Integer
'---------------------------------------------------------

On Error GoTo tableElpKMIndexUpdate_Error
tableElpKMIndex_Update = 0

Select Case Trim(recElpKMIndex.Method)

    Case "AddNew"
                        tableElpKMIndex.AddNew
                        Call tableElpKMIndex_PutBuffer(recElpKMIndex)
                        tableElpKMIndex.Update
    Case "Update"
                        tableElpKMIndex.Edit
                        Call tableElpKMIndex_PutBuffer(recElpKMIndex)
                        tableElpKMIndex.Update
    Case "Delete"
                        tableElpKMIndex.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableElpKMIndexUpdate_Error:
'---------------------------------------------------------
    tableElpKMIndex_Update = Err
    Resume tableElpKMIndexUpdate_End

tableElpKMIndexUpdate_End:

End Function








'-----------------------------------------------------
Sub dbElpKMIndex_Error(recElpKMIndex As typeElpKMIndex)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recElpKMIndex.Id & ": " & recElpKMIndex.Classe & Chr$(13)

Select Case mId$(recElpKMIndex.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recElpKMIndex.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbElpKMIndex.bas :  ( " & Trim(recElpKMIndex.obj) & " : " & Trim(recElpKMIndex.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbElpKMIndex_ReadE(recElpKMIndex As typeElpKMIndex)
'-----------------------------------------------------

dbElpKMIndex_ReadE = Null

recElpKMIndex.Err = tableElpKMIndex_Read(recElpKMIndex)
If recElpKMIndex.Err > 0 Then

'    If recElpKMIndex.Err < 9990 Or recElpKMIndex.Err >= 9999 Then
        Call dbElpKMIndex_Error(recElpKMIndex)
        dbElpKMIndex_ReadE = recElpKMIndex.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbElpKMIndex_Update(recElpKMIndex As typeElpKMIndex)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbElpKMIndex_Update = Null


recElpKMIndex.Err = tableElpKMIndex_Update(recElpKMIndex)

If recElpKMIndex.Err <> 0 Then
    Call dbElpKMIndex_Error(recElpKMIndex)
    dbElpKMIndex_Update = recElpKMIndex.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recElpKMIndex_Init(recElpKMIndex As typeElpKMIndex)
recElpKMIndex.Method = ""
recElpKMIndex.obj = "ElpKMIndex"
recElpKMIndex.Err = ""
recElpKMIndex.Id = ""
recElpKMIndex.Classe = 0
recElpKMIndex.ElpKMSrc_Id = 0
recElpKMIndex.Memo = Null

End Sub


