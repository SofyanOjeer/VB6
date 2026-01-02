Attribute VB_Name = "mdbZADRESS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public rsYADRESS0 As Recordset
Dim blnZADRESS0_Open As Boolean
'---------------------------------------------------------
'-----------------------------------------------------
Sub mdbZADRESS0_Close_Rs()
'-----------------------------------------------------
If blnZADRESS0_Open Then
    rsYADRESS0.Close
    blnZADRESS0_Open = False
End If

End Sub


'-----------------------------------------------------
Sub mdbZADRESS0_Open_Rs()
'-----------------------------------------------------

If Not blnZADRESS0_Open Then
    Set rsYADRESS0 = MDB.OpenRecordset("ZADRESS0")
    rsYADRESS0.Index = "PrimaryKey"
    blnZADRESS0_Open = True
End If
End Sub

'---------------------------------------------------------
Public Function mdbZADRESS0_Read_Rs(lMethod As String, recYADRESS0 As typeYADRESS0)
'---------------------------------------------------------

On Error GoTo Error_Handler

mdbZADRESS0_Read_Rs = Null


Select Case Trim(lMethod)
     Case "Seek=", "AddNew", "Update", "Delete"

                        'rsYADRESS0.Seek "=", recYADRESS0.CLIENAETB, recYADRESS0.CLIENACLI
                        If rsYADRESS0.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        'rsYADRESS0.Seek "<=", recYADRESS0.CLIENAETB, recYADRESS0.CLIENACLI
                        If rsYADRESS0.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        'rsYADRESS0.Seek ">=", recYADRESS0.CLIENAETB, recYADRESS0.CLIENACLI
                        If rsYADRESS0.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        'rsYADRESS0.Seek ">", recYADRESS0.CLIENAETB, recYADRESS0.CLIENACLI
                        If rsYADRESS0.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        rsYADRESS0.MoveNext
                        If rsYADRESS0.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        rsYADRESS0.MovePrevious
                        If rsYADRESS0.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        rsYADRESS0.MoveFirst
                        If rsYADRESS0.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        rsYADRESS0.MoveLast
                        If rsYADRESS0.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If lMethod <> "AddNew" Then
    Call mdbZADRESS0_GetBuffer_Rs(recYADRESS0)
End If

Exit Function

'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

    Resume Next
    mdbZADRESS0_Read_Rs = Error

End Function

'---------------------------------------------------------
Public Function mdbZADRESS0_Update_Rs(lMethod As String, recYADRESS0 As typeYADRESS0)
'---------------------------------------------------------
On Error GoTo Error_Handler
mdbZADRESS0_Update_Rs = Null

Select Case Trim(lMethod)

    Case "AddNew"
                        rsYADRESS0.AddNew
                        Call mdbZADRESS0_PutBuffer_Rs(recYADRESS0)
                        rsYADRESS0.Update
    Case "Update"
                        rsYADRESS0.Edit
                        Call mdbZADRESS0_PutBuffer_Rs(recYADRESS0)
                        rsYADRESS0.Update
    Case "Delete"
                        rsYADRESS0.Delete
    Case Else
                        Error 9999
End Select

Exit Function

Error_Handler:
'---------------------------------------------------------
    Resume Next
    mdbZADRESS0_Update_Rs = Error
End Function





Public Function mdbZADRESS0_GetBuffer_Rs(recYADRESS0 As typeYADRESS0)
On Error GoTo Error_Handler
mdbZADRESS0_GetBuffer_Rs = Null
recYADRESS0.ADRESSETA = rsYADRESS0("ADRESSETA")
recYADRESS0.ADRESSTYP = rsYADRESS0("ADRESSTYP")
recYADRESS0.ADRESSPLA = rsYADRESS0("ADRESSPLA")
recYADRESS0.ADRESSNUM = rsYADRESS0("ADRESSNUM")
recYADRESS0.ADRESSCOA = rsYADRESS0("ADRESSCOA")
recYADRESS0.ADRESSDLI = rsYADRESS0("ADRESSDLI")
recYADRESS0.ADRESSDDE = rsYADRESS0("ADRESSDDE")
recYADRESS0.ADRESSRA1 = rsYADRESS0("ADRESSRA1")
recYADRESS0.ADRESSRA2 = rsYADRESS0("ADRESSRA2")
recYADRESS0.ADRESSAD1 = rsYADRESS0("ADRESSAD1")
recYADRESS0.ADRESSAD2 = rsYADRESS0("ADRESSAD2")
recYADRESS0.ADRESSAD3 = rsYADRESS0("ADRESSAD3")
recYADRESS0.ADRESSCOP = rsYADRESS0("ADRESSCOP")
recYADRESS0.ADRESSVIL = rsYADRESS0("ADRESSVIL")
recYADRESS0.ADRESSPAY = rsYADRESS0("ADRESSPAY")
recYADRESS0.ADRESSTEL = rsYADRESS0("ADRESSTEL")
recYADRESS0.ADRESSFAX = rsYADRESS0("ADRESSFAX")
recYADRESS0.ADRESSTEX = rsYADRESS0("ADRESSTEX")

Exit Function
Error_Handler:
mdbZADRESS0_GetBuffer_Rs = Error
End Function
Public Function mdbZADRESS0_PutBuffer_Rs(recYADRESS0 As typeYADRESS0)
On Error GoTo Error_Handler
mdbZADRESS0_PutBuffer_Rs = Null

rsYADRESS0("ADRESSETA") = recYADRESS0.ADRESSETA
rsYADRESS0("ADRESSTYP") = recYADRESS0.ADRESSTYP
rsYADRESS0("ADRESSPLA") = recYADRESS0.ADRESSPLA
rsYADRESS0("ADRESSNUM") = recYADRESS0.ADRESSNUM
rsYADRESS0("ADRESSCOA") = recYADRESS0.ADRESSCOA
rsYADRESS0("ADRESSDLI") = recYADRESS0.ADRESSDLI
rsYADRESS0("ADRESSDDE") = recYADRESS0.ADRESSDDE
rsYADRESS0("ADRESSRA1") = recYADRESS0.ADRESSRA1
rsYADRESS0("ADRESSRA2") = recYADRESS0.ADRESSRA2
rsYADRESS0("ADRESSAD1") = recYADRESS0.ADRESSAD1
rsYADRESS0("ADRESSAD2") = recYADRESS0.ADRESSAD2
rsYADRESS0("ADRESSAD3") = recYADRESS0.ADRESSAD3
rsYADRESS0("ADRESSCOP") = recYADRESS0.ADRESSCOP
rsYADRESS0("ADRESSVIL") = recYADRESS0.ADRESSVIL
rsYADRESS0("ADRESSPAY") = recYADRESS0.ADRESSPAY
rsYADRESS0("ADRESSTEL") = recYADRESS0.ADRESSTEL
rsYADRESS0("ADRESSFAX") = recYADRESS0.ADRESSFAX
rsYADRESS0("ADRESSTEX") = recYADRESS0.ADRESSTEX
Exit Function
Error_Handler:
mdbZADRESS0_PutBuffer_Rs = Error
End Function





