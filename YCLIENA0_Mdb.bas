Attribute VB_Name = "mdbYCLIENA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public rsYCLIENA0 As Recordset
Dim blnYCLIENA0_Open As Boolean
'---------------------------------------------------------
'-----------------------------------------------------
Sub mdbYCLIENA0_Close_Rs()
'-----------------------------------------------------
If blnYCLIENA0_Open Then
    rsYCLIENA0.Close
    blnYCLIENA0_Open = False
End If

End Sub


'-----------------------------------------------------
Sub mdbYCLIENA0_Open_Rs()
'-----------------------------------------------------

If Not blnYCLIENA0_Open Then
    Set rsYCLIENA0 = MDB.OpenRecordset("ZCLIENA0")
    rsYCLIENA0.Index = "PrimaryKey"
    blnYCLIENA0_Open = True
End If
End Sub

'---------------------------------------------------------
Public Function mdbYCLIENA0_Read_Rs(lMethod As String, recYCLIENA0 As typeYCLIENA0)
'---------------------------------------------------------

On Error GoTo Error_Handler

mdbYCLIENA0_Read_Rs = Null


Select Case Trim(lMethod)
     Case "Seek=", "AddNew", "Update", "Delete"

                        rsYCLIENA0.Seek "=", recYCLIENA0.ID, recYCLIENA0.seq
                        If rsYCLIENA0.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        rsYCLIENA0.Seek "<=", recYCLIENA0.ID, recYCLIENA0.seq
                        If rsYCLIENA0.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        rsYCLIENA0.Seek ">=", recYCLIENA0.ID, recYCLIENA0.seq
                        If rsYCLIENA0.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        rsYCLIENA0.Seek ">", recYCLIENA0.ID, recYCLIENA0.seq
                        If rsYCLIENA0.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        rsYCLIENA0.MoveNext
                        If rsYCLIENA0.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        rsYCLIENA0.MovePrevious
                        If rsYCLIENA0.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        rsYCLIENA0.MoveFirst
                        If rsYCLIENA0.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        rsYCLIENA0.MoveLast
                        If rsYCLIENA0.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If lMethod <> "AddNew" Then
    Call rsYCLIENA0_GetBuffer(recYCLIENA0)
End If

Exit Function

'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

    Resume Next
    mdbYCLIENA0_Read_Rs = Error

End Function

'---------------------------------------------------------
Public Function mdbYCLIENA0_Update_Rs(lMethod As String, recYCLIENA0 As typeYCLIENA0)
'---------------------------------------------------------
On Error GoTo Error_Handler
mdbYCLIENA0_Update_Rs = Null

Select Case Trim(lMethod)

    Case "AddNew"
                        rsYCLIENA0.AddNew
                        Call mdbYCLIENA0_PutBuffer_Rs(recYCLIENA0)
                        rsYCLIENA0.Update
    Case "Update"
                        rsYCLIENA0.Edit
                        Call mdbYCLIENA0_PutBuffer_Rs(recYCLIENA0)
                        rsYCLIENA0.Update
    Case "Delete"
                        rsYCLIENA0.Delete
    Case Else
                        Error 9999
End Select

Exit Function

Error_Handler:
'---------------------------------------------------------
    Resume Next
    mdbYCLIENA0_Update_Rs = Error
End Function





Public Function mdbYCLIENA0_GetBuffer_Rs(recYCLIENA0 As typeYCLIENA0)
On Error GoTo Error_Handler
mdbYCLIENA0_GetBuffer_Rs = Null
recYCLIENA0.CLIENAETB = rsYCLIENA0("CLIENAETB")
recYCLIENA0.CLIENACLI = rsYCLIENA0("CLIENACLI")
recYCLIENA0.CLIENAAGE = rsYCLIENA0("CLIENAAGE")
recYCLIENA0.CLIENAETA = rsYCLIENA0("CLIENAETA")
recYCLIENA0.CLIENARA1 = rsYCLIENA0("CLIENARA1")
recYCLIENA0.CLIENARA2 = rsYCLIENA0("CLIENARA2")
recYCLIENA0.CLIENASIG = rsYCLIENA0("CLIENASIG")
recYCLIENA0.CLIENASRN = rsYCLIENA0("CLIENASRN")
recYCLIENA0.CLIENASRT = rsYCLIENA0("CLIENASRT")
recYCLIENA0.CLIENADNA = rsYCLIENA0("CLIENADNA")
recYCLIENA0.CLIENAREG = rsYCLIENA0("CLIENAREG")
recYCLIENA0.CLIENANAT = rsYCLIENA0("CLIENANAT")
recYCLIENA0.CLIENARSD = rsYCLIENA0("CLIENARSD")
recYCLIENA0.CLIENARES = rsYCLIENA0("CLIENARES")
recYCLIENA0.CLIENAECO = rsYCLIENA0("CLIENAECO")
recYCLIENA0.CLIENAACT = rsYCLIENA0("CLIENAACT")
recYCLIENA0.CLIENAPAI = rsYCLIENA0("CLIENAPAI")
recYCLIENA0.CLIENACRD = rsYCLIENA0("CLIENACRD")
recYCLIENA0.CLIENAADM = rsYCLIENA0("CLIENAADM")
recYCLIENA0.CLIENAATR = rsYCLIENA0("CLIENAATR")
recYCLIENA0.CLIENABIL = rsYCLIENA0("CLIENABIL")
recYCLIENA0.CLIENACAT = rsYCLIENA0("CLIENACAT")
recYCLIENA0.CLIENACOT = rsYCLIENA0("CLIENACOT")
recYCLIENA0.CLIENACHQ = rsYCLIENA0("CLIENACHQ")
recYCLIENA0.CLIENADAT = rsYCLIENA0("CLIENADAT")
recYCLIENA0.CLIENASAC = rsYCLIENA0("CLIENASAC")
recYCLIENA0.CLIENAGEO = rsYCLIENA0("CLIENAGEO")
recYCLIENA0.CLIENAENT = rsYCLIENA0("CLIENAENT")
recYCLIENA0.CLIENAMES = rsYCLIENA0("CLIENAMES")
recYCLIENA0.CLIENAPAY = rsYCLIENA0("CLIENAPAY")
recYCLIENA0.CLIENAFIL = rsYCLIENA0("CLIENAFIL")
recYCLIENA0.CLIENABIM = rsYCLIENA0("CLIENABIM")
recYCLIENA0.CLIENADOU = rsYCLIENA0("CLIENADOU")
recYCLIENA0.CLIENALI1 = rsYCLIENA0("CLIENALI1")
recYCLIENA0.CLIENALI2 = rsYCLIENA0("CLIENALI2")
recYCLIENA0.CLIENAEXT = rsYCLIENA0("CLIENAEXT")
recYCLIENA0.CLIENACOL = rsYCLIENA0("CLIENACOL")
recYCLIENA0.CLIENATIE = rsYCLIENA0("CLIENATIE")
recYCLIENA0.CLIENASEL = rsYCLIENA0("CLIENASEL")
recYCLIENA0.CLIENAPCS = rsYCLIENA0("CLIENAPCS")
recYCLIENA0.CLIENACRE = rsYCLIENA0("CLIENACRE")
Exit Function
Error_Handler:
mdbYCLIENA0_GetBuffer_Rs = Error
End Function
Public Function mdbYCLIENA0_PutBuffer_Rs(recYCLIENA0 As typeYCLIENA0)
On Error GoTo Error_Handler
mdbYCLIENA0_PutBuffer_Rs = Null
rsYCLIENA0("CLIENAETB") = recYCLIENA0.CLIENAETB
rsYCLIENA0("CLIENACLI") = recYCLIENA0.CLIENACLI
rsYCLIENA0("CLIENAAGE") = recYCLIENA0.CLIENAAGE
rsYCLIENA0("CLIENAETA") = recYCLIENA0.CLIENAETA
rsYCLIENA0("CLIENARA1") = recYCLIENA0.CLIENARA1
rsYCLIENA0("CLIENARA2") = recYCLIENA0.CLIENARA2
rsYCLIENA0("CLIENASIG") = recYCLIENA0.CLIENASIG
rsYCLIENA0("CLIENASRN") = recYCLIENA0.CLIENASRN
rsYCLIENA0("CLIENASRT") = recYCLIENA0.CLIENASRT
rsYCLIENA0("CLIENADNA") = recYCLIENA0.CLIENADNA
rsYCLIENA0("CLIENAREG") = recYCLIENA0.CLIENAREG
rsYCLIENA0("CLIENANAT") = recYCLIENA0.CLIENANAT
rsYCLIENA0("CLIENARSD") = recYCLIENA0.CLIENARSD
rsYCLIENA0("CLIENARES") = recYCLIENA0.CLIENARES
rsYCLIENA0("CLIENAECO") = recYCLIENA0.CLIENAECO
rsYCLIENA0("CLIENAACT") = recYCLIENA0.CLIENAACT
rsYCLIENA0("CLIENAPAI") = recYCLIENA0.CLIENAPAI
rsYCLIENA0("CLIENACRD") = recYCLIENA0.CLIENACRD
rsYCLIENA0("CLIENAADM") = recYCLIENA0.CLIENAADM
rsYCLIENA0("CLIENAATR") = recYCLIENA0.CLIENAATR
rsYCLIENA0("CLIENABIL") = recYCLIENA0.CLIENABIL
rsYCLIENA0("CLIENACAT") = recYCLIENA0.CLIENACAT
rsYCLIENA0("CLIENACOT") = recYCLIENA0.CLIENACOT
rsYCLIENA0("CLIENACHQ") = recYCLIENA0.CLIENACHQ
rsYCLIENA0("CLIENADAT") = recYCLIENA0.CLIENADAT
rsYCLIENA0("CLIENASAC") = recYCLIENA0.CLIENASAC
rsYCLIENA0("CLIENAGEO") = recYCLIENA0.CLIENAGEO
rsYCLIENA0("CLIENAENT") = recYCLIENA0.CLIENAENT
rsYCLIENA0("CLIENAMES") = recYCLIENA0.CLIENAMES
rsYCLIENA0("CLIENAPAY") = recYCLIENA0.CLIENAPAY
rsYCLIENA0("CLIENAFIL") = recYCLIENA0.CLIENAFIL
rsYCLIENA0("CLIENABIM") = recYCLIENA0.CLIENABIM
rsYCLIENA0("CLIENADOU") = recYCLIENA0.CLIENADOU
rsYCLIENA0("CLIENALI1") = recYCLIENA0.CLIENALI1
rsYCLIENA0("CLIENALI2") = recYCLIENA0.CLIENALI2
rsYCLIENA0("CLIENAEXT") = recYCLIENA0.CLIENAEXT
rsYCLIENA0("CLIENACOL") = recYCLIENA0.CLIENACOL
rsYCLIENA0("CLIENATIE") = recYCLIENA0.CLIENATIE
rsYCLIENA0("CLIENASEL") = recYCLIENA0.CLIENASEL
rsYCLIENA0("CLIENAPCS") = recYCLIENA0.CLIENAPCS
rsYCLIENA0("CLIENACRE") = recYCLIENA0.CLIENACRE
Exit Function
Error_Handler:
mdbYCLIENA0_PutBuffer_Rs = Error
End Function




