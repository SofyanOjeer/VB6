Attribute VB_Name = "mdbZPLAN0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public rsYPLAN0 As Recordset
Dim blnZPLAN0_Open As Boolean
'---------------------------------------------------------
'-----------------------------------------------------
Sub mdbZPLAN0_Close_Rs()
'-----------------------------------------------------
If blnZPLAN0_Open Then
    rsYPLAN0.Close
    blnZPLAN0_Open = False
End If

End Sub


'-----------------------------------------------------
Sub mdbZPLAN0_Open_Rs()
'-----------------------------------------------------

If Not blnZPLAN0_Open Then
    Set rsYPLAN0 = MDB.OpenRecordset("ZPLAN0")
    rsYPLAN0.Index = "PrimaryKey"
    blnZPLAN0_Open = True
End If
End Sub

'---------------------------------------------------------
Public Function mdbZPLAN0_Read_Rs(lMethod As String, recYPLAN0 As typeYPLAN0)
'---------------------------------------------------------

On Error GoTo Error_Handler

mdbZPLAN0_Read_Rs = Null


Select Case Trim(lMethod)
     Case "Seek=", "AddNew", "Update", "Delete"

                        'rsYPLAN0.Seek "=", recYPLAN0.CLIENAETB, recYPLAN0.CLIENACLI
                        If rsYPLAN0.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        'rsYPLAN0.Seek "<=", recYPLAN0.CLIENAETB, recYPLAN0.CLIENACLI
                        If rsYPLAN0.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        'rsYPLAN0.Seek ">=", recYPLAN0.CLIENAETB, recYPLAN0.CLIENACLI
                        If rsYPLAN0.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        'rsYPLAN0.Seek ">", recYPLAN0.CLIENAETB, recYPLAN0.CLIENACLI
                        If rsYPLAN0.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        rsYPLAN0.MoveNext
                        If rsYPLAN0.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        rsYPLAN0.MovePrevious
                        If rsYPLAN0.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        rsYPLAN0.MoveFirst
                        If rsYPLAN0.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        rsYPLAN0.MoveLast
                        If rsYPLAN0.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If lMethod <> "AddNew" Then
    Call mdbZPLAN0_GetBuffer_Rs(recYPLAN0)
End If

Exit Function

'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

    Resume Next
    mdbZPLAN0_Read_Rs = Error

End Function

'---------------------------------------------------------
Public Function mdbZPLAN0_Update_Rs(lMethod As String, recYPLAN0 As typeYPLAN0)
'---------------------------------------------------------
On Error GoTo Error_Handler
mdbZPLAN0_Update_Rs = Null

Select Case Trim(lMethod)

    Case "AddNew"
                        rsYPLAN0.AddNew
                        Call mdbZPLAN0_PutBuffer_Rs(recYPLAN0)
                        rsYPLAN0.Update
    Case "Update"
                        rsYPLAN0.Edit
                        Call mdbZPLAN0_PutBuffer_Rs(recYPLAN0)
                        rsYPLAN0.Update
    Case "Delete"
                        rsYPLAN0.Delete
    Case Else
                        Error 9999
End Select

Exit Function

Error_Handler:
'---------------------------------------------------------
    Resume Next
    mdbZPLAN0_Update_Rs = Error
End Function





Public Function mdbZPLAN0_GetBuffer_Rs(recYPLAN0 As typeYPLAN0)
On Error GoTo Error_Handler
mdbZPLAN0_GetBuffer_Rs = Null

recYPLAN0.PLANETABL = rsYPLAN0("PLANETABL")
recYPLAN0.PLANPLAN = rsYPLAN0("PLANPLAN")
recYPLAN0.PLANCOOBL = rsYPLAN0("PLANCOOBL")
recYPLAN0.PLANINTIT = rsYPLAN0("PLANINTIT")
recYPLAN0.PLANCOPRO = rsYPLAN0("PLANCOPRO")
recYPLAN0.PLANCLASS = rsYPLAN0("PLANCLASS")
recYPLAN0.PLANFONCT = rsYPLAN0("PLANFONCT")
recYPLAN0.PLANSESOL = rsYPLAN0("PLANSESOL")
recYPLAN0.PLANGEDEP = rsYPLAN0("PLANGEDEP")
recYPLAN0.PLANTIERS = rsYPLAN0("PLANTIERS")
recYPLAN0.PLANFICOB = rsYPLAN0("PLANFICOB")
recYPLAN0.PLANCARAC = rsYPLAN0("PLANCARAC")
recYPLAN0.PLANPESTO = rsYPLAN0("PLANPESTO")
recYPLAN0.PLANNBPER = rsYPLAN0("PLANNBPER")
recYPLAN0.PLANNBMOU = rsYPLAN0("PLANNBMOU")
recYPLAN0.PLANINEXT = rsYPLAN0("PLANINEXT")
recYPLAN0.PLANPROGR = rsYPLAN0("PLANPROGR")
Exit Function
Error_Handler:
mdbZPLAN0_GetBuffer_Rs = Error
End Function
Public Function mdbZPLAN0_PutBuffer_Rs(recYPLAN0 As typeYPLAN0)
On Error GoTo Error_Handler
mdbZPLAN0_PutBuffer_Rs = Null

rsYPLAN0("PLANETABL") = recYPLAN0.PLANETABL
rsYPLAN0("PLANPLAN") = recYPLAN0.PLANPLAN
rsYPLAN0("PLANCOOBL") = recYPLAN0.PLANCOOBL
rsYPLAN0("PLANINTIT") = recYPLAN0.PLANINTIT
rsYPLAN0("PLANCOPRO") = recYPLAN0.PLANCOPRO
rsYPLAN0("PLANCLASS") = recYPLAN0.PLANCLASS
rsYPLAN0("PLANFONCT") = recYPLAN0.PLANFONCT
rsYPLAN0("PLANSESOL") = recYPLAN0.PLANSESOL
rsYPLAN0("PLANGEDEP") = recYPLAN0.PLANGEDEP
rsYPLAN0("PLANTIERS") = recYPLAN0.PLANTIERS
rsYPLAN0("PLANFICOB") = recYPLAN0.PLANFICOB
rsYPLAN0("PLANCARAC") = recYPLAN0.PLANCARAC
rsYPLAN0("PLANPESTO") = recYPLAN0.PLANPESTO
rsYPLAN0("PLANNBPER") = recYPLAN0.PLANNBPER
rsYPLAN0("PLANNBMOU") = recYPLAN0.PLANNBMOU
rsYPLAN0("PLANINEXT") = recYPLAN0.PLANINEXT
rsYPLAN0("PLANPROGR") = recYPLAN0.PLANPROGR
Exit Function
Error_Handler:
mdbZPLAN0_PutBuffer_Rs = Error
End Function





