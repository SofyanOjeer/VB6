Attribute VB_Name = "mdbZTITULA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public rsYTITULA0 As Recordset
Dim blnZTITULA0_Open As Boolean
'---------------------------------------------------------
'-----------------------------------------------------
Sub mdbZTITULA0_Close_Rs()
'-----------------------------------------------------
If blnZTITULA0_Open Then
    rsYTITULA0.Close
    blnZTITULA0_Open = False
End If

End Sub


'-----------------------------------------------------
Sub mdbZTITULA0_Open_Rs()
'-----------------------------------------------------

If Not blnZTITULA0_Open Then
    Set rsYTITULA0 = MDB.OpenRecordset("ZTITULA0")
    rsYTITULA0.Index = "PrimaryKey"
    blnZTITULA0_Open = True
End If
End Sub

'---------------------------------------------------------
Public Function mdbZTITULA0_Read_Rs(lMethod As String, recYTITULA0 As typeYTITULA0)
'---------------------------------------------------------

On Error GoTo Error_Handler

mdbZTITULA0_Read_Rs = Null


Select Case Trim(lMethod)
     Case "Seek=", "AddNew", "Update", "Delete"

                        'rsYTITULA0.Seek "=", recYTITULA0.CLIENAETB, recYTITULA0.CLIENACLI
                        If rsYTITULA0.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        'rsYTITULA0.Seek "<=", recYTITULA0.CLIENAETB, recYTITULA0.CLIENACLI
                        If rsYTITULA0.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        'rsYTITULA0.Seek ">=", recYTITULA0.CLIENAETB, recYTITULA0.CLIENACLI
                        If rsYTITULA0.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        'rsYTITULA0.Seek ">", recYTITULA0.CLIENAETB, recYTITULA0.CLIENACLI
                        If rsYTITULA0.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        rsYTITULA0.MoveNext
                        If rsYTITULA0.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        rsYTITULA0.MovePrevious
                        If rsYTITULA0.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        rsYTITULA0.MoveFirst
                        If rsYTITULA0.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        rsYTITULA0.MoveLast
                        If rsYTITULA0.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If lMethod <> "AddNew" Then
    Call mdbZTITULA0_GetBuffer_Rs(recYTITULA0)
End If

Exit Function

'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

    Resume Next
    mdbZTITULA0_Read_Rs = Error

End Function

'---------------------------------------------------------
Public Function mdbZTITULA0_Update_Rs(lMethod As String, recYTITULA0 As typeYTITULA0)
'---------------------------------------------------------
On Error GoTo Error_Handler
mdbZTITULA0_Update_Rs = Null

Select Case Trim(lMethod)

    Case "AddNew"
                        rsYTITULA0.AddNew
                        Call mdbZTITULA0_PutBuffer_Rs(recYTITULA0)
                        rsYTITULA0.Update
    Case "Update"
                        rsYTITULA0.Edit
                        Call mdbZTITULA0_PutBuffer_Rs(recYTITULA0)
                        rsYTITULA0.Update
    Case "Delete"
                        rsYTITULA0.Delete
    Case Else
                        Error 9999
End Select

Exit Function

Error_Handler:
'---------------------------------------------------------
    Resume Next
    mdbZTITULA0_Update_Rs = Error
End Function





Public Function mdbZTITULA0_GetBuffer_Rs(recYTITULA0 As typeYTITULA0)
On Error GoTo Error_Handler
mdbZTITULA0_GetBuffer_Rs = Null

recYTITULA0.TITULAETA = rsYTITULA0("TITULAETA")
recYTITULA0.TITULAPLA = rsYTITULA0("TITULAPLA")
recYTITULA0.TITULACOM = rsYTITULA0("TITULACOM")
recYTITULA0.TITULACLI = rsYTITULA0("TITULACLI")
recYTITULA0.TITULAPRI = rsYTITULA0("TITULAPRI")
recYTITULA0.TITULATPR = rsYTITULA0("TITULATPR")
Exit Function
Error_Handler:
mdbZTITULA0_GetBuffer_Rs = Error
End Function
Public Function mdbZTITULA0_PutBuffer_Rs(recYTITULA0 As typeYTITULA0)
On Error GoTo Error_Handler
mdbZTITULA0_PutBuffer_Rs = Null
rsYTITULA0("TITULAETA") = recYTITULA0.TITULAETA
rsYTITULA0("TITULAPLA") = recYTITULA0.TITULAPLA
rsYTITULA0("TITULACOM") = recYTITULA0.TITULACOM
rsYTITULA0("TITULACLI") = recYTITULA0.TITULACLI
rsYTITULA0("TITULAPRI") = recYTITULA0.TITULAPRI
rsYTITULA0("TITULATPR") = recYTITULA0.TITULATPR

Exit Function
Error_Handler:
mdbZTITULA0_PutBuffer_Rs = Error
End Function





