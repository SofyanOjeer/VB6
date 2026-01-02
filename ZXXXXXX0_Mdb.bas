Attribute VB_Name = "mdbZXXXXXX0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public rsYXXXXXX0 As Recordset
Dim blnZXXXXXX0_Open As Boolean
'---------------------------------------------------------
'-----------------------------------------------------
Sub mdbZXXXXXX0_Close_Rs()
'-----------------------------------------------------
If blnZXXXXXX0_Open Then
    rsYXXXXXX0.Close
    blnZXXXXXX0_Open = False
End If

End Sub


'-----------------------------------------------------
Sub mdbZXXXXXX0_Open_Rs()
'-----------------------------------------------------

If Not blnZXXXXXX0_Open Then
    Set rsYXXXXXX0 = MDB.OpenRecordset("ZXXXXXX0")
    rsYXXXXXX0.Index = "PrimaryKey"
    blnZXXXXXX0_Open = True
End If
End Sub

'---------------------------------------------------------
Public Function mdbZXXXXXX0_Read_Rs(lMethod As String, recYXXXXXX0 As typeYXXXXXX0)
'---------------------------------------------------------

On Error GoTo Error_Handler

mdbZXXXXXX0_Read_Rs = Null


Select Case Trim(lMethod)
     Case "Seek=", "AddNew", "Update", "Delete"

                        'rsYXXXXXX0.Seek "=", recYXXXXXX0.CLIENAETB, recYXXXXXX0.CLIENACLI
                        If rsYXXXXXX0.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        'rsYXXXXXX0.Seek "<=", recYXXXXXX0.CLIENAETB, recYXXXXXX0.CLIENACLI
                        If rsYXXXXXX0.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        'rsYXXXXXX0.Seek ">=", recYXXXXXX0.CLIENAETB, recYXXXXXX0.CLIENACLI
                        If rsYXXXXXX0.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        'rsYXXXXXX0.Seek ">", recYXXXXXX0.CLIENAETB, recYXXXXXX0.CLIENACLI
                        If rsYXXXXXX0.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        rsYXXXXXX0.MoveNext
                        If rsYXXXXXX0.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        rsYXXXXXX0.MovePrevious
                        If rsYXXXXXX0.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        rsYXXXXXX0.MoveFirst
                        If rsYXXXXXX0.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        rsYXXXXXX0.MoveLast
                        If rsYXXXXXX0.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If lMethod <> "AddNew" Then
    Call mdbZXXXXXX0_GetBuffer_Rs(recYXXXXXX0)
End If

Exit Function

'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

    Resume Next
    mdbZXXXXXX0_Read_Rs = Error

End Function

'---------------------------------------------------------
Public Function mdbZXXXXXX0_Update_Rs(lMethod As String, recYXXXXXX0 As typeYXXXXXX0)
'---------------------------------------------------------
On Error GoTo Error_Handler
mdbZXXXXXX0_Update_Rs = Null

Select Case Trim(lMethod)

    Case "AddNew"
                        rsYXXXXXX0.AddNew
                        Call mdbZXXXXXX0_PutBuffer_Rs(recYXXXXXX0)
                        rsYXXXXXX0.Update
    Case "Update"
                        rsYXXXXXX0.Edit
                        Call mdbZXXXXXX0_PutBuffer_Rs(recYXXXXXX0)
                        rsYXXXXXX0.Update
    Case "Delete"
                        rsYXXXXXX0.Delete
    Case Else
                        Error 9999
End Select

Exit Function

Error_Handler:
'---------------------------------------------------------
    Resume Next
    mdbZXXXXXX0_Update_Rs = Error
End Function





Public Function mdbZXXXXXX0_GetBuffer_Rs(recYXXXXXX0 As typeYXXXXXX0)
On Error GoTo Error_Handler
mdbZXXXXXX0_GetBuffer_Rs = Null


Exit Function
Error_Handler:
mdbZXXXXXX0_GetBuffer_Rs = Error
End Function
Public Function mdbZXXXXXX0_PutBuffer_Rs(recYXXXXXX0 As typeYXXXXXX0)
On Error GoTo Error_Handler
mdbZXXXXXX0_PutBuffer_Rs = Null


Exit Function
Error_Handler:
mdbZXXXXXX0_PutBuffer_Rs = Error
End Function





