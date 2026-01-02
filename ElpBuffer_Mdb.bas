Attribute VB_Name = "mdbElpBuffer"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableElpBuffer As Recordset
Dim tableElpBufferOpen As Boolean
Public mElpBuffer_Id As Long

Type typeElpBuffer
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Id                  As String * 10
    Seq                 As Long
    Data                As String

End Type

Public recElpBuffer As typeElpBuffer


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableElpBuffer_Close()
'-----------------------------------------------------
If tableElpBufferOpen Then
    tableElpBuffer.Close
    tableElpBufferOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableElpBuffer_GetBuffer(recElpBuffer As typeElpBuffer)
'---------------------------------------------------------
recElpBuffer.Id = tableElpBuffer("Id")
recElpBuffer.Seq = tableElpBuffer("Seq")
recElpBuffer.Data = tableElpBuffer("Data")

End Sub


'-----------------------------------------------------
Sub tableElpBuffer_Open()
'-----------------------------------------------------

If Not tableElpBufferOpen Then
    Set tableElpBuffer = MDB.OpenRecordset("ElpBuffer")
    tableElpBuffer.Index = "PrimaryKey"
    tableElpBufferOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableElpBuffer_PutBuffer(recElpBuffer As typeElpBuffer)
'---------------------------------------------------------

tableElpBuffer("Id") = recElpBuffer.Id
tableElpBuffer("Seq") = recElpBuffer.Seq
tableElpBuffer("Data") = recElpBuffer.Data
End Sub


'---------------------------------------------------------
Public Function tableElpBuffer_Read(recElpBuffer As typeElpBuffer) As Integer
'---------------------------------------------------------

On Error GoTo tableElpBuffer_Read_Error
tableElpBuffer_Read = 0


Select Case Trim(recElpBuffer.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableElpBuffer.Seek "=", recElpBuffer.Id, recElpBuffer.Seq
                        If tableElpBuffer.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableElpBuffer.Seek "<=", recElpBuffer.Id, recElpBuffer.Seq
                        If tableElpBuffer.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableElpBuffer.Seek ">=", recElpBuffer.Id, recElpBuffer.Seq
                        If tableElpBuffer.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableElpBuffer.Seek ">", recElpBuffer.Id, recElpBuffer.Seq
                        If tableElpBuffer.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableElpBuffer.MoveNext
                        If tableElpBuffer.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableElpBuffer.MovePrevious
                        If tableElpBuffer.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableElpBuffer.MoveFirst
                        If tableElpBuffer.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableElpBuffer.MoveLast
                        If tableElpBuffer.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recElpBuffer.Method <> "AddNew      " Then
    Call tableElpBuffer_GetBuffer(recElpBuffer)
End If

Exit Function

'---------------------------------------------------------
tableElpBuffer_Read_Error:
'---------------------------------------------------------

    tableElpBuffer_Read = Err
    Resume tableElpBuffer_Read_End

tableElpBuffer_Read_End:

End Function

'---------------------------------------------------------
Public Function tableElpBuffer_Update(recElpBuffer As typeElpBuffer) As Integer
'---------------------------------------------------------

On Error GoTo tableElpBufferUpdate_Error
tableElpBuffer_Update = 0

Select Case Trim(recElpBuffer.Method)

    Case "AddNew"
                        tableElpBuffer.AddNew
                        Call tableElpBuffer_PutBuffer(recElpBuffer)
                        tableElpBuffer.Update
    Case "Update"
                        tableElpBuffer.Edit
                        Call tableElpBuffer_PutBuffer(recElpBuffer)
                        tableElpBuffer.Update
    Case "Delete"
                        tableElpBuffer.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableElpBufferUpdate_Error:
'---------------------------------------------------------
    tableElpBuffer_Update = Err
    Resume tableElpBufferUpdate_End

tableElpBufferUpdate_End:

End Function








'-----------------------------------------------------
Sub dbElpBuffer_Error(recElpBuffer As typeElpBuffer)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recElpBuffer.Id & ": " & recElpBuffer.Seq & Chr$(13)

Select Case mId$(recElpBuffer.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recElpBuffer.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbElpBuffer.bas :  ( " & Trim(recElpBuffer.obj) & " : " & Trim(recElpBuffer.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbElpBuffer_ReadE(recElpBuffer As typeElpBuffer)
'-----------------------------------------------------

dbElpBuffer_ReadE = Null

recElpBuffer.Err = tableElpBuffer_Read(recElpBuffer)
If recElpBuffer.Err > 0 Then

'    If recElpBuffer.Err < 9990 Or recElpBuffer.Err >= 9999 Then
        Call dbElpBuffer_Error(recElpBuffer)
        dbElpBuffer_ReadE = recElpBuffer.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbElpBuffer_Update(recElpBuffer As typeElpBuffer)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbElpBuffer_Update = Null


recElpBuffer.Err = tableElpBuffer_Update(recElpBuffer)

If recElpBuffer.Err <> 0 Then
    Call dbElpBuffer_Error(recElpBuffer)
    dbElpBuffer_Update = recElpBuffer.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recElpBuffer_Init(recElpBuffer As typeElpBuffer)
recElpBuffer.Method = ""
recElpBuffer.obj = "ElpBuffer"
recElpBuffer.Err = ""
recElpBuffer.Id = ""
recElpBuffer.Seq = 0
recElpBuffer.Data = ""
End Sub

Public Sub recElpBuffer_Init_id(recElpBuffer As typeElpBuffer)
recElpBuffer_Init recElpBuffer
mElpBuffer_Id = mElpBuffer_Id + 1
recElpBuffer.Id = Format$(mElpBuffer_Id, "0000000000")

End Sub
