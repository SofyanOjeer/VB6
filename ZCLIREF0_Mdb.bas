Attribute VB_Name = "mdbZCLIREF0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public rsYCLIREF0 As Recordset
Dim blnZCLIREF0_Open As Boolean
'---------------------------------------------------------
'-----------------------------------------------------
Sub mdbZCLIREF0_Close_Rs()
'-----------------------------------------------------
If blnZCLIREF0_Open Then
    rsYCLIREF0.Close
    blnZCLIREF0_Open = False
End If

End Sub


'-----------------------------------------------------
Sub mdbZCLIREF0_Open_Rs()
'-----------------------------------------------------

If Not blnZCLIREF0_Open Then
    Set rsYCLIREF0 = MDB.OpenRecordset("ZCLIREF0")
    rsYCLIREF0.Index = "PrimaryKey"
    blnZCLIREF0_Open = True
End If
End Sub

'---------------------------------------------------------
Public Function mdbZCLIREF0_Read_Rs(lMethod As String, recYCLIREF0 As typeYCLIREF0)
'---------------------------------------------------------

On Error GoTo Error_Handler

mdbZCLIREF0_Read_Rs = Null


Select Case Trim(lMethod)
     Case "Seek=", "AddNew", "Update", "Delete"

                        'rsYCLIREF0.Seek "=", recYCLIREF0.CLIENAETB, recYCLIREF0.CLIENACLI
                        If rsYCLIREF0.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        'rsYCLIREF0.Seek "<=", recYCLIREF0.CLIENAETB, recYCLIREF0.CLIENACLI
                        If rsYCLIREF0.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        'rsYCLIREF0.Seek ">=", recYCLIREF0.CLIENAETB, recYCLIREF0.CLIENACLI
                        If rsYCLIREF0.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        'rsYCLIREF0.Seek ">", recYCLIREF0.CLIENAETB, recYCLIREF0.CLIENACLI
                        If rsYCLIREF0.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        rsYCLIREF0.MoveNext
                        If rsYCLIREF0.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        rsYCLIREF0.MovePrevious
                        If rsYCLIREF0.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        rsYCLIREF0.MoveFirst
                        If rsYCLIREF0.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        rsYCLIREF0.MoveLast
                        If rsYCLIREF0.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If lMethod <> "AddNew" Then
    Call mdbZCLIREF0_GetBuffer_Rs(recYCLIREF0)
End If

Exit Function

'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

    Resume Next
    mdbZCLIREF0_Read_Rs = Error

End Function

'---------------------------------------------------------
Public Function mdbZCLIREF0_Update_Rs(lMethod As String, recYCLIREF0 As typeYCLIREF0)
'---------------------------------------------------------
On Error GoTo Error_Handler
mdbZCLIREF0_Update_Rs = Null

Select Case Trim(lMethod)

    Case "AddNew"
                        rsYCLIREF0.AddNew
                        Call mdbZCLIREF0_PutBuffer_Rs(recYCLIREF0)
                        rsYCLIREF0.Update
    Case "Update"
                        rsYCLIREF0.Edit
                        Call mdbZCLIREF0_PutBuffer_Rs(recYCLIREF0)
                        rsYCLIREF0.Update
    Case "Delete"
                        rsYCLIREF0.Delete
    Case Else
                        Error 9999
End Select

Exit Function

Error_Handler:
'---------------------------------------------------------
    Resume Next
    mdbZCLIREF0_Update_Rs = Error
End Function





Public Function mdbZCLIREF0_GetBuffer_Rs(recYCLIREF0 As typeYCLIREF0)
On Error GoTo Error_Handler
mdbZCLIREF0_GetBuffer_Rs = Null
recYCLIREF0.CLIREFETA = rsYCLIREF0("CLIREFETA")
recYCLIREF0.CLIREFCLI = rsYCLIREF0("CLIREFCLI")
recYCLIREF0.CLIREFCOR = rsYCLIREF0("CLIREFCOR")
recYCLIREF0.CLIREFREF = rsYCLIREF0("CLIREFREF")

Exit Function
Error_Handler:
mdbZCLIREF0_GetBuffer_Rs = Error
End Function
Public Function mdbZCLIREF0_PutBuffer_Rs(recYCLIREF0 As typeYCLIREF0)
On Error GoTo Error_Handler
mdbZCLIREF0_PutBuffer_Rs = Null

rsYCLIREF0("CLIREFETA") = recYCLIREF0.CLIREFETA
rsYCLIREF0("CLIREFCLI") = recYCLIREF0.CLIREFCLI
rsYCLIREF0("CLIREFCOR") = recYCLIREF0.CLIREFCOR
rsYCLIREF0("CLIREFREF") = recYCLIREF0.CLIREFREF
Exit Function
Error_Handler:
mdbZCLIREF0_PutBuffer_Rs = Error
End Function





