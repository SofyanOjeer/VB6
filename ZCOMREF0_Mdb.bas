Attribute VB_Name = "mdbZCOMREF0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public rsYCOMREF0 As Recordset
Dim blnZCOMREF0_Open As Boolean
'---------------------------------------------------------
'-----------------------------------------------------
Sub mdbZCOMREF0_Close_Rs()
'-----------------------------------------------------
If blnZCOMREF0_Open Then
    rsYCOMREF0.Close
    blnZCOMREF0_Open = False
End If

End Sub


'-----------------------------------------------------
Sub mdbZCOMREF0_Open_Rs()
'-----------------------------------------------------

If Not blnZCOMREF0_Open Then
    Set rsYCOMREF0 = MDB.OpenRecordset("ZCOMREF0")
    rsYCOMREF0.Index = "PrimaryKey"
    blnZCOMREF0_Open = True
End If
End Sub

'---------------------------------------------------------
Public Function mdbZCOMREF0_Read_Rs(lMethod As String, recYCOMREF0 As typeYCOMREF0)
'---------------------------------------------------------

On Error GoTo Error_Handler

mdbZCOMREF0_Read_Rs = Null


Select Case Trim(lMethod)
     Case "Seek=", "AddNew", "Update", "Delete"

                        'rsYCOMREF0.Seek "=", recYCOMREF0.CLIENAETB, recYCOMREF0.CLIENACLI
                        If rsYCOMREF0.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        'rsYCOMREF0.Seek "<=", recYCOMREF0.CLIENAETB, recYCOMREF0.CLIENACLI
                        If rsYCOMREF0.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        'rsYCOMREF0.Seek ">=", recYCOMREF0.CLIENAETB, recYCOMREF0.CLIENACLI
                        If rsYCOMREF0.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        'rsYCOMREF0.Seek ">", recYCOMREF0.CLIENAETB, recYCOMREF0.CLIENACLI
                        If rsYCOMREF0.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        rsYCOMREF0.MoveNext
                        If rsYCOMREF0.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        rsYCOMREF0.MovePrevious
                        If rsYCOMREF0.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        rsYCOMREF0.MoveFirst
                        If rsYCOMREF0.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        rsYCOMREF0.MoveLast
                        If rsYCOMREF0.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If lMethod <> "AddNew" Then
    Call mdbZCOMREF0_GetBuffer_Rs(recYCOMREF0)
End If

Exit Function

'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

    Resume Next
    mdbZCOMREF0_Read_Rs = Error

End Function

'---------------------------------------------------------
Public Function mdbZCOMREF0_Update_Rs(lMethod As String, recYCOMREF0 As typeYCOMREF0)
'---------------------------------------------------------
On Error GoTo Error_Handler
mdbZCOMREF0_Update_Rs = Null

Select Case Trim(lMethod)

    Case "AddNew"
                        rsYCOMREF0.AddNew
                        Call mdbZCOMREF0_PutBuffer_Rs(recYCOMREF0)
                        rsYCOMREF0.Update
    Case "Update"
                        rsYCOMREF0.Edit
                        Call mdbZCOMREF0_PutBuffer_Rs(recYCOMREF0)
                        rsYCOMREF0.Update
    Case "Delete"
                        rsYCOMREF0.Delete
    Case Else
                        Error 9999
End Select

Exit Function

Error_Handler:
'---------------------------------------------------------
    Resume Next
    mdbZCOMREF0_Update_Rs = Error
End Function





Public Function mdbZCOMREF0_GetBuffer_Rs(recYCOMREF0 As typeYCOMREF0)
On Error GoTo Error_Handler
mdbZCOMREF0_GetBuffer_Rs = Null

recYCOMREF0.COMREFETA = rsYCOMREF0("COMREFETA")
recYCOMREF0.COMREFPLA = rsYCOMREF0("COMREFPLA")
recYCOMREF0.COMREFCOM = rsYCOMREF0("COMREFCOM")
recYCOMREF0.COMREFCOR = rsYCOMREF0("COMREFCOR")
recYCOMREF0.COMREFREF = rsYCOMREF0("COMREFREF")
Exit Function
Error_Handler:
mdbZCOMREF0_GetBuffer_Rs = Error
End Function
Public Function mdbZCOMREF0_PutBuffer_Rs(recYCOMREF0 As typeYCOMREF0)
On Error GoTo Error_Handler
mdbZCOMREF0_PutBuffer_Rs = Null
rsYCOMREF0("COMREFETA") = recYCOMREF0.COMREFETA
rsYCOMREF0("COMREFPLA") = recYCOMREF0.COMREFPLA
rsYCOMREF0("COMREFCOM") = recYCOMREF0.COMREFCOM
rsYCOMREF0("COMREFCOR") = recYCOMREF0.COMREFCOR
rsYCOMREF0("COMREFREF") = recYCOMREF0.COMREFREF

Exit Function
Error_Handler:
mdbZCOMREF0_PutBuffer_Rs = Error
End Function





