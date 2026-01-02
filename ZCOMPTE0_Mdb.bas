Attribute VB_Name = "mdbZCOMPTE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public rsYCOMPTE0 As Recordset
Dim blnZCOMPTE0_Open As Boolean
'---------------------------------------------------------
'-----------------------------------------------------
Sub mdbZCOMPTE0_Close_Rs()
'-----------------------------------------------------
If blnZCOMPTE0_Open Then
    rsYCOMPTE0.Close
    blnZCOMPTE0_Open = False
End If

End Sub


'-----------------------------------------------------
Sub mdbZCOMPTE0_Open_Rs()
'-----------------------------------------------------

If Not blnZCOMPTE0_Open Then
    Set rsYCOMPTE0 = MDB.OpenRecordset("ZCOMPTE0")
    rsYCOMPTE0.Index = "PrimaryKey"
    blnZCOMPTE0_Open = True
End If
End Sub

'---------------------------------------------------------
Public Function mdbZCOMPTE0_Read_Rs(lMethod As String, recYCOMPTE0 As typeYCOMPTE0)
'---------------------------------------------------------

On Error GoTo Error_Handler

mdbZCOMPTE0_Read_Rs = Null


Select Case Trim(lMethod)
     Case "Seek=", "AddNew", "Update", "Delete"

                        'rsYCOMPTE0.Seek "=", recYCOMPTE0.CLIENAETB, recYCOMPTE0.CLIENACLI
                        If rsYCOMPTE0.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        'rsYCOMPTE0.Seek "<=", recYCOMPTE0.CLIENAETB, recYCOMPTE0.CLIENACLI
                        If rsYCOMPTE0.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        'rsYCOMPTE0.Seek ">=", recYCOMPTE0.CLIENAETB, recYCOMPTE0.CLIENACLI
                        If rsYCOMPTE0.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        'rsYCOMPTE0.Seek ">", recYCOMPTE0.CLIENAETB, recYCOMPTE0.CLIENACLI
                        If rsYCOMPTE0.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        rsYCOMPTE0.MoveNext
                        If rsYCOMPTE0.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        rsYCOMPTE0.MovePrevious
                        If rsYCOMPTE0.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        rsYCOMPTE0.MoveFirst
                        If rsYCOMPTE0.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        rsYCOMPTE0.MoveLast
                        If rsYCOMPTE0.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If lMethod <> "AddNew" Then
    Call mdbZCOMPTE0_GetBuffer_Rs(recYCOMPTE0)
End If

Exit Function

'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

    Resume Next
    mdbZCOMPTE0_Read_Rs = Error

End Function

'---------------------------------------------------------
Public Function mdbZCOMPTE0_Update_Rs(lMethod As String, recYCOMPTE0 As typeYCOMPTE0)
'---------------------------------------------------------
On Error GoTo Error_Handler
mdbZCOMPTE0_Update_Rs = Null

Select Case Trim(lMethod)

    Case "AddNew"
                        rsYCOMPTE0.AddNew
                        Call mdbZCOMPTE0_PutBuffer_Rs(recYCOMPTE0)
                        rsYCOMPTE0.Update
    Case "Update"
                        rsYCOMPTE0.Edit
                        Call mdbZCOMPTE0_PutBuffer_Rs(recYCOMPTE0)
                        rsYCOMPTE0.Update
    Case "Delete"
                        rsYCOMPTE0.Delete
    Case Else
                        Error 9999
End Select

Exit Function

Error_Handler:
'---------------------------------------------------------
    Resume Next
    mdbZCOMPTE0_Update_Rs = Error
End Function





Public Function mdbZCOMPTE0_GetBuffer_Rs(recYCOMPTE0 As typeYCOMPTE0)
On Error GoTo Error_Handler
mdbZCOMPTE0_GetBuffer_Rs = Null

recYCOMPTE0.COMPTEETA = rsYCOMPTE0("COMPTEETA")
recYCOMPTE0.COMPTEPLA = rsYCOMPTE0("COMPTEPLA")
recYCOMPTE0.COMPTECOM = rsYCOMPTE0("COMPTECOM")
recYCOMPTE0.COMPTEOBL = rsYCOMPTE0("COMPTEOBL")
recYCOMPTE0.COMPTEINT = rsYCOMPTE0("COMPTEINT")
recYCOMPTE0.COMPTEAGE = rsYCOMPTE0("COMPTEAGE")
recYCOMPTE0.COMPTEDEV = rsYCOMPTE0("COMPTEDEV")
recYCOMPTE0.COMPTEOUV = rsYCOMPTE0("COMPTEOUV")
recYCOMPTE0.COMPTECLO = rsYCOMPTE0("COMPTECLO")
recYCOMPTE0.COMPTELOR = rsYCOMPTE0("COMPTELOR")
recYCOMPTE0.COMPTESUC = rsYCOMPTE0("COMPTESUC")
recYCOMPTE0.COMPTECLA = rsYCOMPTE0("COMPTECLA")
recYCOMPTE0.COMPTEFON = rsYCOMPTE0("COMPTEFON")
recYCOMPTE0.COMPTEBLO = rsYCOMPTE0("COMPTEBLO")
recYCOMPTE0.COMPTEMOT = rsYCOMPTE0("COMPTEMOT")
recYCOMPTE0.COMPTESEN = rsYCOMPTE0("COMPTESEN")
recYCOMPTE0.COMPTEMOD = rsYCOMPTE0("COMPTEMOD")
Exit Function
Error_Handler:
mdbZCOMPTE0_GetBuffer_Rs = Error
End Function
Public Function mdbZCOMPTE0_PutBuffer_Rs(recYCOMPTE0 As typeYCOMPTE0)
On Error GoTo Error_Handler
mdbZCOMPTE0_PutBuffer_Rs = Null

rsYCOMPTE0("COMPTEETA") = recYCOMPTE0.COMPTEETA
rsYCOMPTE0("COMPTEPLA") = recYCOMPTE0.COMPTEPLA
rsYCOMPTE0("COMPTECOM") = recYCOMPTE0.COMPTECOM
rsYCOMPTE0("COMPTEOBL") = recYCOMPTE0.COMPTEOBL
rsYCOMPTE0("COMPTEINT") = recYCOMPTE0.COMPTEINT
rsYCOMPTE0("COMPTEAGE") = recYCOMPTE0.COMPTEAGE
rsYCOMPTE0("COMPTEDEV") = recYCOMPTE0.COMPTEDEV
rsYCOMPTE0("COMPTEOUV") = recYCOMPTE0.COMPTEOUV
rsYCOMPTE0("COMPTECLO") = recYCOMPTE0.COMPTECLO
rsYCOMPTE0("COMPTELOR") = recYCOMPTE0.COMPTELOR
rsYCOMPTE0("COMPTESUC") = recYCOMPTE0.COMPTESUC
rsYCOMPTE0("COMPTECLA") = recYCOMPTE0.COMPTECLA
rsYCOMPTE0("COMPTEFON") = recYCOMPTE0.COMPTEFON
rsYCOMPTE0("COMPTEBLO") = recYCOMPTE0.COMPTEBLO
rsYCOMPTE0("COMPTEMOT") = recYCOMPTE0.COMPTEMOT
rsYCOMPTE0("COMPTESEN") = recYCOMPTE0.COMPTESEN
rsYCOMPTE0("COMPTEMOD") = recYCOMPTE0.COMPTEMOD
Exit Function
Error_Handler:
mdbZCOMPTE0_PutBuffer_Rs = Error
End Function





