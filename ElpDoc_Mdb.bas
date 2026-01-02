Attribute VB_Name = "mdbElpDoc"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableElpDoc As Recordset
Dim tableElpDocOpen As Boolean

Type typeElpDoc
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Id           As String * 12
    K1           As String * 12
    K2           As String * 12
    Memo         As Variant

End Type

Public recElpDoc As typeElpDoc


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableElpDoc_Close()
'-----------------------------------------------------
If tableElpDocOpen Then
    tableElpDoc.Close
    tableElpDocOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableElpDoc_GetBuffer(recElpDoc As typeElpDoc)
'---------------------------------------------------------

recElpDoc.Id = tableElpDoc("Id")
recElpDoc.K1 = tableElpDoc("K1")
recElpDoc.K2 = tableElpDoc("K2")

recElpDoc.Memo = tableElpDoc("Memo")

End Sub


'-----------------------------------------------------
Sub tableElpDoc_Open()
'-----------------------------------------------------

If Not tableElpDocOpen Then
    Set tableElpDoc = MDB.OpenRecordset("ElpDoc")
    tableElpDoc.Index = "PrimaryKey"
    tableElpDocOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableElpDoc_PutBuffer(recElpDoc As typeElpDoc)
'---------------------------------------------------------

tableElpDoc("id") = recElpDoc.Id
tableElpDoc("K1") = recElpDoc.K1
tableElpDoc("K2") = recElpDoc.K2
tableElpDoc("Memo") = recElpDoc.Memo
End Sub


'---------------------------------------------------------
Public Function tableElpDoc_Read(recElpDoc As typeElpDoc) As Integer
'---------------------------------------------------------

On Error GoTo tableElpDoc_Read_Error
tableElpDoc_Read = 0


Select Case Trim(recElpDoc.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableElpDoc.Seek "=", recElpDoc.Id, recElpDoc.K1, recElpDoc.K2
                        If tableElpDoc.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableElpDoc.Seek "<=", recElpDoc.Id, recElpDoc.K1, recElpDoc.K2
                        If tableElpDoc.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableElpDoc.Seek ">=", recElpDoc.Id, recElpDoc.K1, recElpDoc.K2
                        If tableElpDoc.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>"
                        tableElpDoc.Seek ">", recElpDoc.Id, recElpDoc.K1, recElpDoc.K2
                        If tableElpDoc.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableElpDoc.MoveNext
                        If tableElpDoc.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableElpDoc.MovePrevious
                        If tableElpDoc.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableElpDoc.MoveFirst
                        If tableElpDoc.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableElpDoc.MoveLast
                        If tableElpDoc.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recElpDoc.Method <> "AddNew      " Then
    Call tableElpDoc_GetBuffer(recElpDoc)
End If

Exit Function

'---------------------------------------------------------
tableElpDoc_Read_Error:
'---------------------------------------------------------

    tableElpDoc_Read = Err
    Resume tableElpDoc_Read_End

tableElpDoc_Read_End:

End Function

'---------------------------------------------------------
Public Function tableElpDoc_Update(recElpDoc As typeElpDoc) As Integer
'---------------------------------------------------------

On Error GoTo tableElpDocUpdate_Error
tableElpDoc_Update = 0

Select Case Trim(recElpDoc.Method)

    Case "AddNew"
                        tableElpDoc.AddNew
                        Call tableElpDoc_PutBuffer(recElpDoc)
                        tableElpDoc.Update
    Case "Update"
                        tableElpDoc.Edit
                        Call tableElpDoc_PutBuffer(recElpDoc)
                        tableElpDoc.Update
    Case "Delete"
                       tableElpDoc.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableElpDocUpdate_Error:
'---------------------------------------------------------
    tableElpDoc_Update = Err
    Resume tableElpDocUpdate_End

tableElpDocUpdate_End:

End Function








'-----------------------------------------------------
Sub dbElpDoc_Error(recElpDoc As typeElpDoc)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & Trim(recElpDoc.Id) & " : " & Trim(recElpDoc.K1) & " : " & Trim(recElpDoc.K2) & Chr$(13)

Select Case mId$(recElpDoc.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recElpDoc.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbElpDoc.bas :  ( " & Trim(recElpDoc.obj) & " : " & Trim(recElpDoc.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbElpDoc_Read(recElpDoc As typeElpDoc)
'-----------------------------------------------------

dbElpDoc_Read = Null

recElpDoc.Err = tableElpDoc_Read(recElpDoc)
If recElpDoc.Err > 0 Then

    If recElpDoc.Err < 9990 Or recElpDoc.Err >= 9999 Then
        Call dbElpDoc_Error(recElpDoc)
        dbElpDoc_Read = recElpDoc.Err
    End If
End If

End Function

'-----------------------------------------------------
Function dbElpDoc_Update(recElpDoc As typeElpDoc)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbElpDoc_Update = Null


recElpDoc.Err = tableElpDoc_Update(recElpDoc)

If recElpDoc.Err <> 0 Then
    Call dbElpDoc_Error(recElpDoc)
    dbElpDoc_Update = recElpDoc.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recElpDoc_Init(recElpDoc As typeElpDoc)
recElpDoc.Method = ""
recElpDoc.obj = "ElpDoc"
recElpDoc.Err = ""
recElpDoc.Id = ""
recElpDoc.K1 = ""
recElpDoc.K2 = ""
recElpDoc.Memo = ""
End Sub

