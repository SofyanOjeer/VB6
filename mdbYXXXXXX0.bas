Attribute VB_Name = "YXXXXXX0_Mdb"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public rsYXXXXXX0 As Recordset
Dim blnYXXXXXX0_Open As Boolean
'---------------------------------------------------------
'-----------------------------------------------------
Sub rsYXXXXXX0_Close()
'-----------------------------------------------------
If blnYXXXXXX0_Open Then
    rsYXXXXXX0.Close
    blnYXXXXXX0_Open = False
End If

End Sub


'-----------------------------------------------------
Sub rsYXXXXXX0_Open()
'-----------------------------------------------------

If Not blnYXXXXXX0_Open Then
    Set rsYXXXXXX0 = MDB.OpenRecordset("YXXXXXX0")
    rsYXXXXXX0.Index = "PrimaryKey"
    blnYXXXXXX0_Open = True
End If
End Sub

'---------------------------------------------------------
Public Function rsYXXXXXX0_Read(lMethod As String, recYXXXXXX0 As typeYXXXXXX0)
'---------------------------------------------------------

On Error GoTo Error_Handler

rsYXXXXXX0_Read = Null


Select Case Trim(lMethod)
     Case "Seek=", "AddNew", "Update", "Delete"

                        rsYXXXXXX0.Seek "=", recYXXXXXX0.ID, recYXXXXXX0.seq
                        If rsYXXXXXX0.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        rsYXXXXXX0.Seek "<=", recYXXXXXX0.ID, recYXXXXXX0.seq
                        If rsYXXXXXX0.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        rsYXXXXXX0.Seek ">=", recYXXXXXX0.ID, recYXXXXXX0.seq
                        If rsYXXXXXX0.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        rsYXXXXXX0.Seek ">", recYXXXXXX0.ID, recYXXXXXX0.seq
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
    Call rsYXXXXXX0_GetBuffer(recYXXXXXX0)
End If

Exit Function

'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

    Resume Next
    rsYXXXXXX0_Read = Error

End Function

'---------------------------------------------------------
Public Function rsYXXXXXX0_Update(lMethod As String, recYXXXXXX0 As typeYXXXXXX0)
'---------------------------------------------------------

On Error GoTo Error_Handler
rsYXXXXXX0_Update = Null

Select Case Trim(lMethod)

    Case "AddNew"
                        rsYXXXXXX0.AddNew
                        Call rsYXXXXXX0_PutBuffer(recYXXXXXX0)
                        rsYXXXXXX0.Update
    Case "Update"
                        rsYXXXXXX0.Edit
                        Call rsYXXXXXX0_PutBuffer(recYXXXXXX0)
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
    rsYXXXXXX0_Update = Error
End Function








