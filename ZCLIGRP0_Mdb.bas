Attribute VB_Name = "mdbZCLIGRP0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public rsYCLIGRP0 As Recordset
Dim blnZCLIGRP0_Open As Boolean
'---------------------------------------------------------
'-----------------------------------------------------
Sub mdbZCLIGRP0_Close_Rs()
'-----------------------------------------------------
If blnZCLIGRP0_Open Then
    rsYCLIGRP0.Close
    blnZCLIGRP0_Open = False
End If

End Sub


'-----------------------------------------------------
Sub mdbZCLIGRP0_Open_Rs()
'-----------------------------------------------------

If Not blnZCLIGRP0_Open Then
    Set rsYCLIGRP0 = MDB.OpenRecordset("ZCLIGRP0")
    rsYCLIGRP0.Index = "PrimaryKey"
    blnZCLIGRP0_Open = True
End If
End Sub

'---------------------------------------------------------
Public Function mdbZCLIGRP0_Read_Rs(lMethod As String, recYCLIGRP0 As typeYCLIGRP0)
'---------------------------------------------------------

On Error GoTo Error_Handler

mdbZCLIGRP0_Read_Rs = Null


Select Case Trim(lMethod)
     Case "Seek=", "AddNew", "Update", "Delete"

                        'rsYCLIGRP0.Seek "=", recYCLIGRP0.CLIENAETB, recYCLIGRP0.CLIENACLI
                        If rsYCLIGRP0.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        'rsYCLIGRP0.Seek "<=", recYCLIGRP0.CLIENAETB, recYCLIGRP0.CLIENACLI
                        If rsYCLIGRP0.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        'rsYCLIGRP0.Seek ">=", recYCLIGRP0.CLIENAETB, recYCLIGRP0.CLIENACLI
                        If rsYCLIGRP0.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        'rsYCLIGRP0.Seek ">", recYCLIGRP0.CLIENAETB, recYCLIGRP0.CLIENACLI
                        If rsYCLIGRP0.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        rsYCLIGRP0.MoveNext
                        If rsYCLIGRP0.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        rsYCLIGRP0.MovePrevious
                        If rsYCLIGRP0.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        rsYCLIGRP0.MoveFirst
                        If rsYCLIGRP0.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        rsYCLIGRP0.MoveLast
                        If rsYCLIGRP0.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If lMethod <> "AddNew" Then
    Call mdbZCLIGRP0_GetBuffer_Rs(recYCLIGRP0)
End If

Exit Function

'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

    Resume Next
    mdbZCLIGRP0_Read_Rs = Error

End Function

'---------------------------------------------------------
Public Function mdbZCLIGRP0_Update_Rs(lMethod As String, recYCLIGRP0 As typeYCLIGRP0)
'---------------------------------------------------------
On Error GoTo Error_Handler
mdbZCLIGRP0_Update_Rs = Null

Select Case Trim(lMethod)

    Case "AddNew"
                        rsYCLIGRP0.AddNew
                        Call mdbZCLIGRP0_PutBuffer_Rs(recYCLIGRP0)
                        rsYCLIGRP0.Update
    Case "Update"
                        rsYCLIGRP0.Edit
                        Call mdbZCLIGRP0_PutBuffer_Rs(recYCLIGRP0)
                        rsYCLIGRP0.Update
    Case "Delete"
                        rsYCLIGRP0.Delete
    Case Else
                        Error 9999
End Select

Exit Function

Error_Handler:
'---------------------------------------------------------
    Resume Next
    mdbZCLIGRP0_Update_Rs = Error
End Function





Public Function mdbZCLIGRP0_GetBuffer_Rs(recYCLIGRP0 As typeYCLIGRP0)
On Error GoTo Error_Handler
mdbZCLIGRP0_GetBuffer_Rs = Null

recYCLIGRP0.CLIGRPETB = rsYCLIGRP0("CLIGRPETB")
recYCLIGRP0.CLIGRPCLI = rsYCLIGRP0("CLIGRPCLI")
recYCLIGRP0.CLIGRPREG = rsYCLIGRP0("CLIGRPREG")
recYCLIGRP0.CLIGRPREL = rsYCLIGRP0("CLIGRPREL")
recYCLIGRP0.CLIGRPCOM = rsYCLIGRP0("CLIGRPCOM")
recYCLIGRP0.CLIGRPAUT = rsYCLIGRP0("CLIGRPAUT")
recYCLIGRP0.CLIGRPRAT = rsYCLIGRP0("CLIGRPRAT")
recYCLIGRP0.CLIGRPTAU = rsYCLIGRP0("CLIGRPTAU")
recYCLIGRP0.CLIGRPPAR = rsYCLIGRP0("CLIGRPPAR")
Exit Function
Error_Handler:
mdbZCLIGRP0_GetBuffer_Rs = Error
End Function
Public Function mdbZCLIGRP0_PutBuffer_Rs(recYCLIGRP0 As typeYCLIGRP0)
On Error GoTo Error_Handler
mdbZCLIGRP0_PutBuffer_Rs = Null
rsYCLIGRP0("CLIGRPETB") = recYCLIGRP0.CLIGRPETB
rsYCLIGRP0("CLIGRPCLI") = recYCLIGRP0.CLIGRPCLI
rsYCLIGRP0("CLIGRPREG") = recYCLIGRP0.CLIGRPREG
rsYCLIGRP0("CLIGRPREL") = recYCLIGRP0.CLIGRPREL
rsYCLIGRP0("CLIGRPCOM") = recYCLIGRP0.CLIGRPCOM
rsYCLIGRP0("CLIGRPAUT") = recYCLIGRP0.CLIGRPAUT
rsYCLIGRP0("CLIGRPRAT") = recYCLIGRP0.CLIGRPRAT
rsYCLIGRP0("CLIGRPTAU") = recYCLIGRP0.CLIGRPTAU
rsYCLIGRP0("CLIGRPPAR") = recYCLIGRP0.CLIGRPPAR

Exit Function
Error_Handler:
mdbZCLIGRP0_PutBuffer_Rs = Error
End Function





