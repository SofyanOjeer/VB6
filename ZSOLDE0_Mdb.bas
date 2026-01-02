Attribute VB_Name = "mdbZSOLDE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public rsYSOLDE0 As Recordset
Dim blnZSOLDE0_Open As Boolean
'---------------------------------------------------------
'-----------------------------------------------------
Sub mdbZSOLDE0_Close_Rs()
'-----------------------------------------------------
If blnZSOLDE0_Open Then
    rsYSOLDE0.Close
    blnZSOLDE0_Open = False
End If

End Sub


'-----------------------------------------------------
Sub mdbZSOLDE0_Open_Rs()
'-----------------------------------------------------

If Not blnZSOLDE0_Open Then
    Set rsYSOLDE0 = MDB.OpenRecordset("ZSOLDE0")
    rsYSOLDE0.Index = "PrimaryKey"
    blnZSOLDE0_Open = True
End If
End Sub

'---------------------------------------------------------
Public Function mdbZSOLDE0_Read_Rs(lMethod As String, recYSOLDE0 As typeYSOLDE0)
'---------------------------------------------------------

On Error GoTo Error_Handler

mdbZSOLDE0_Read_Rs = Null


Select Case Trim(lMethod)
     Case "Seek=", "AddNew", "Update", "Delete"

                        'rsYSOLDE0.Seek "=", recYSOLDE0.CLIENAETB, recYSOLDE0.CLIENACLI
                        If rsYSOLDE0.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        'rsYSOLDE0.Seek "<=", recYSOLDE0.CLIENAETB, recYSOLDE0.CLIENACLI
                        If rsYSOLDE0.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        'rsYSOLDE0.Seek ">=", recYSOLDE0.CLIENAETB, recYSOLDE0.CLIENACLI
                        If rsYSOLDE0.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        'rsYSOLDE0.Seek ">", recYSOLDE0.CLIENAETB, recYSOLDE0.CLIENACLI
                        If rsYSOLDE0.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        rsYSOLDE0.MoveNext
                        If rsYSOLDE0.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        rsYSOLDE0.MovePrevious
                        If rsYSOLDE0.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        rsYSOLDE0.MoveFirst
                        If rsYSOLDE0.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        rsYSOLDE0.MoveLast
                        If rsYSOLDE0.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If lMethod <> "AddNew" Then
    Call mdbZSOLDE0_GetBuffer_Rs(recYSOLDE0)
End If

Exit Function

'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

    Resume Next
    mdbZSOLDE0_Read_Rs = Error

End Function

'---------------------------------------------------------
Public Function mdbZSOLDE0_Update_Rs(lMethod As String, recYSOLDE0 As typeYSOLDE0)
'---------------------------------------------------------
On Error GoTo Error_Handler
mdbZSOLDE0_Update_Rs = Null

Select Case Trim(lMethod)

    Case "AddNew"
                        rsYSOLDE0.AddNew
                        Call mdbZSOLDE0_PutBuffer_Rs(recYSOLDE0)
                        rsYSOLDE0.Update
    Case "Update"
                        rsYSOLDE0.Edit
                        Call mdbZSOLDE0_PutBuffer_Rs(recYSOLDE0)
                        rsYSOLDE0.Update
    Case "Delete"
                        rsYSOLDE0.Delete
    Case Else
                        Error 9999
End Select

Exit Function

Error_Handler:
'---------------------------------------------------------
    Resume Next
    mdbZSOLDE0_Update_Rs = Error
End Function





Public Function mdbZSOLDE0_GetBuffer_Rs(recYSOLDE0 As typeYSOLDE0)
On Error GoTo Error_Handler
mdbZSOLDE0_GetBuffer_Rs = Null

recYSOLDE0.SOLDEETA = rsYSOLDE0("SOLDEETA")
recYSOLDE0.SOLDEPLA = rsYSOLDE0("SOLDEPLA")
recYSOLDE0.SOLDECOM = rsYSOLDE0("SOLDECOM")
recYSOLDE0.SOLDEDMO = rsYSOLDE0("SOLDEDMO")
recYSOLDE0.SOLDEDAN = rsYSOLDE0("SOLDEDAN")
recYSOLDE0.SOLDECEN = rsYSOLDE0("SOLDECEN")
recYSOLDE0.SOLDECAN = rsYSOLDE0("SOLDECAN")
recYSOLDE0.SOLDEC01 = rsYSOLDE0("SOLDEC01")
recYSOLDE0.SOLDEC02 = rsYSOLDE0("SOLDEC02")
recYSOLDE0.SOLDEC03 = rsYSOLDE0("SOLDEC03")
recYSOLDE0.SOLDEC04 = rsYSOLDE0("SOLDEC04")
recYSOLDE0.SOLDEC05 = rsYSOLDE0("SOLDEC05")
recYSOLDE0.SOLDEC06 = rsYSOLDE0("SOLDEC06")
recYSOLDE0.SOLDEC07 = rsYSOLDE0("SOLDEC07")
recYSOLDE0.SOLDEC08 = rsYSOLDE0("SOLDEC08")
recYSOLDE0.SOLDEC09 = rsYSOLDE0("SOLDEC09")
recYSOLDE0.SOLDEC10 = rsYSOLDE0("SOLDEC10")
recYSOLDE0.SOLDEC11 = rsYSOLDE0("SOLDEC11")
recYSOLDE0.SOLDEC12 = rsYSOLDE0("SOLDEC12")
recYSOLDE0.SOLDEVEN = rsYSOLDE0("SOLDEVEN")
recYSOLDE0.SOLDEVAN = rsYSOLDE0("SOLDEVAN")
recYSOLDE0.SOLDEV01 = rsYSOLDE0("SOLDEV01")
recYSOLDE0.SOLDEV02 = rsYSOLDE0("SOLDEV02")
recYSOLDE0.SOLDEV03 = rsYSOLDE0("SOLDEV03")
recYSOLDE0.SOLDEV04 = rsYSOLDE0("SOLDEV04")
recYSOLDE0.SOLDEV05 = rsYSOLDE0("SOLDEV05")
recYSOLDE0.SOLDEV06 = rsYSOLDE0("SOLDEV06")
recYSOLDE0.SOLDEV07 = rsYSOLDE0("SOLDEV07")
recYSOLDE0.SOLDEV08 = rsYSOLDE0("SOLDEV08")
recYSOLDE0.SOLDEV09 = rsYSOLDE0("SOLDEV09")
recYSOLDE0.SOLDEV10 = rsYSOLDE0("SOLDEV10")
recYSOLDE0.SOLDEV11 = rsYSOLDE0("SOLDEV11")
recYSOLDE0.SOLDEV12 = rsYSOLDE0("SOLDEV12")
Exit Function
Error_Handler:
mdbZSOLDE0_GetBuffer_Rs = Error
End Function
Public Function mdbZSOLDE0_PutBuffer_Rs(recYSOLDE0 As typeYSOLDE0)
On Error GoTo Error_Handler
mdbZSOLDE0_PutBuffer_Rs = Null

rsYSOLDE0("SOLDEETA") = recYSOLDE0.SOLDEETA
rsYSOLDE0("SOLDEPLA") = recYSOLDE0.SOLDEPLA
rsYSOLDE0("SOLDECOM") = recYSOLDE0.SOLDECOM
rsYSOLDE0("SOLDEDMO") = recYSOLDE0.SOLDEDMO
rsYSOLDE0("SOLDEDAN") = recYSOLDE0.SOLDEDAN
rsYSOLDE0("SOLDECEN") = recYSOLDE0.SOLDECEN
rsYSOLDE0("SOLDECAN") = recYSOLDE0.SOLDECAN
rsYSOLDE0("SOLDEC01") = recYSOLDE0.SOLDEC01
rsYSOLDE0("SOLDEC02") = recYSOLDE0.SOLDEC02
rsYSOLDE0("SOLDEC03") = recYSOLDE0.SOLDEC03
rsYSOLDE0("SOLDEC04") = recYSOLDE0.SOLDEC04
rsYSOLDE0("SOLDEC05") = recYSOLDE0.SOLDEC05
rsYSOLDE0("SOLDEC06") = recYSOLDE0.SOLDEC06
rsYSOLDE0("SOLDEC07") = recYSOLDE0.SOLDEC07
rsYSOLDE0("SOLDEC08") = recYSOLDE0.SOLDEC08
rsYSOLDE0("SOLDEC09") = recYSOLDE0.SOLDEC09
rsYSOLDE0("SOLDEC10") = recYSOLDE0.SOLDEC10
rsYSOLDE0("SOLDEC11") = recYSOLDE0.SOLDEC11
rsYSOLDE0("SOLDEC12") = recYSOLDE0.SOLDEC12
rsYSOLDE0("SOLDEVEN") = recYSOLDE0.SOLDEVEN
rsYSOLDE0("SOLDEVAN") = recYSOLDE0.SOLDEVAN
rsYSOLDE0("SOLDEV01") = recYSOLDE0.SOLDEV01
rsYSOLDE0("SOLDEV02") = recYSOLDE0.SOLDEV02
rsYSOLDE0("SOLDEV03") = recYSOLDE0.SOLDEV03
rsYSOLDE0("SOLDEV04") = recYSOLDE0.SOLDEV04
rsYSOLDE0("SOLDEV05") = recYSOLDE0.SOLDEV05
rsYSOLDE0("SOLDEV06") = recYSOLDE0.SOLDEV06
rsYSOLDE0("SOLDEV07") = recYSOLDE0.SOLDEV07
rsYSOLDE0("SOLDEV08") = recYSOLDE0.SOLDEV08
rsYSOLDE0("SOLDEV09") = recYSOLDE0.SOLDEV09
rsYSOLDE0("SOLDEV10") = recYSOLDE0.SOLDEV10
rsYSOLDE0("SOLDEV11") = recYSOLDE0.SOLDEV11
rsYSOLDE0("SOLDEV12") = recYSOLDE0.SOLDEV12
Exit Function
Error_Handler:
mdbZSOLDE0_PutBuffer_Rs = Error
End Function





