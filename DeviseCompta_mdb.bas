Attribute VB_Name = "mdbDeviseCompta"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableDeviseCompta As Recordset
Public tableDeviseComptaOpen As Boolean
    
Public Sub dbDeviseCompta_Replication()
Dim I As Integer, X As String, Amj As String

ReDim arrDeviseChange(10): arrDeviseChangeNbMax = 10
recDeviseChange_Init XDeviseChange
XDeviseChange.Method = "Seek<="
XDeviseChange.Id1 = "EUR"
XDeviseChange.Id2 = "USD"
XDeviseChange.Amj = DSys ' "19990100"
XDeviseChange.HHMM = "9999"
XDeviseChange.Origine = "C"
tableDeviseCompta_Read XDeviseChange
XDeviseChange.Method = "SnapP0"
XDeviseChange.Amj = XDeviseChange.Amj
XDeviseChange.Origine = "C"
XDeviseChange.Id1 = ""

arrDeviseChange(0) = XDeviseChange
arrDeviseChange(0).Amj = "99999999"
arrDeviseChangeSuite = True
Do Until Not arrDeviseChangeSuite
    arrDeviseChangeNb = 0
    srvDeviseChange_Monitor XDeviseChange
    For I = 1 To arrDeviseChangeNb
        X = Trim(arrDeviseChange(I).ValidationUsr)
        If X <> constàValider And X <> "" And arrDeviseChange(I).Origine <> "T" Then
            arrDeviseChange(I).Method = constAddNew
            tableDeviseCompta_Update arrDeviseChange(I)
        End If
    Next I
    XDeviseChange = arrDeviseChange(arrDeviseChangeNb)
    XDeviseChange.Method = "SnapP0+"
Loop
End Sub


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableDeviseCompta_Close()
'-----------------------------------------------------
If tableDeviseComptaOpen Then
    tableDeviseCompta.Close
    tableDeviseComptaOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableDeviseCompta_GetBuffer(recDeviseCompta As typeDeviseChange)
'---------------------------------------------------------

recDeviseCompta.Id1 = tableDeviseCompta("ID1")
recDeviseCompta.Id2 = tableDeviseCompta("ID2")
recDeviseCompta.Amj = tableDeviseCompta("AMJ")
recDeviseCompta.Origine = tableDeviseCompta("Origine")
recDeviseCompta.HHMM = tableDeviseCompta("HHMM")
recDeviseCompta.QD1 = tableDeviseCompta("QD1")
recDeviseCompta.QD2CoursPivot = tableDeviseCompta("QD2CoursPivot")
recDeviseCompta.QD2AchatNormal = tableDeviseCompta("QD2AchatNormal")
recDeviseCompta.QD2VenteNormal = tableDeviseCompta("QD2VenteNormal")
recDeviseCompta.QD2AchatPrivilégié = tableDeviseCompta("QD2AchatPrivilégié")
recDeviseCompta.QD2VentePrivilégié = tableDeviseCompta("QD2VentePrivilégié")
recDeviseCompta.QD2AchatEnCompte = tableDeviseCompta("QD2AchatEnCompte")
recDeviseCompta.QD2VenteEnCompte = tableDeviseCompta("QD2VenteEnCompte")

recDeviseCompta.SaisieAmj = tableDeviseCompta("SaisieAMJ")
recDeviseCompta.SaisieHMS = tableDeviseCompta("SaisieHMS")
recDeviseCompta.SaisieUsr = tableDeviseCompta("SaisieUsr")
recDeviseCompta.ValidationAMJ = tableDeviseCompta("ValidationAMJ")
recDeviseCompta.ValidationHMS = tableDeviseCompta("ValidationHMS")
recDeviseCompta.ValidationUsr = tableDeviseCompta("ValidationUsr")

End Sub


'-----------------------------------------------------
Sub tableDeviseCompta_Open()
'-----------------------------------------------------

If Not tableDeviseComptaOpen Then
    Set tableDeviseCompta = MDB.OpenRecordset("DeviseCompta")
    tableDeviseCompta.Index = "PrimaryKey"
    tableDeviseComptaOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableDeviseCompta_PutBuffer(recDeviseCompta As typeDeviseChange)
'---------------------------------------------------------

tableDeviseCompta("ID1") = recDeviseCompta.Id1
tableDeviseCompta("ID2") = recDeviseCompta.Id2
tableDeviseCompta("AMJ") = recDeviseCompta.Amj
tableDeviseCompta("Origine") = recDeviseCompta.Origine
tableDeviseCompta("HHMM") = recDeviseCompta.HHMM

tableDeviseCompta("QD1") = recDeviseCompta.QD1
tableDeviseCompta("QD2CoursPivot") = recDeviseCompta.QD2CoursPivot
tableDeviseCompta("QD2AchatNormal") = recDeviseCompta.QD2AchatNormal
tableDeviseCompta("QD2VenteNormal") = recDeviseCompta.QD2VenteNormal
tableDeviseCompta("QD2AchatPrivilégié") = recDeviseCompta.QD2AchatPrivilégié
tableDeviseCompta("QD2VentePrivilégié") = recDeviseCompta.QD2VentePrivilégié
tableDeviseCompta("QD2AchatEnCompte") = recDeviseCompta.QD2AchatEnCompte
tableDeviseCompta("QD2VenteEnCompte") = recDeviseCompta.QD2VenteEnCompte

tableDeviseCompta("SaisieAMJ") = recDeviseCompta.SaisieAmj
tableDeviseCompta("SaisieHMS") = recDeviseCompta.SaisieHMS
tableDeviseCompta("SaisieUsr") = recDeviseCompta.SaisieUsr
tableDeviseCompta("ValidationAMJ") = recDeviseCompta.ValidationAMJ
tableDeviseCompta("ValidationHMS") = recDeviseCompta.ValidationHMS
tableDeviseCompta("ValidationUsr") = recDeviseCompta.ValidationUsr

End Sub


'---------------------------------------------------------
Public Function tableDeviseCompta_Read(recDeviseCompta As typeDeviseChange) As Integer
'---------------------------------------------------------

On Error GoTo tableDeviseCompta_Read_Error
tableDeviseCompta_Read = 0


Select Case recDeviseCompta.Method
     Case "Seek=       "
                        tableDeviseCompta.Seek "=", recDeviseCompta.Id1, recDeviseCompta.Id2, recDeviseCompta.Amj, recDeviseCompta.HHMM, recDeviseCompta.Origine
                        If tableDeviseCompta.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<=      "
                        tableDeviseCompta.Seek "<=", recDeviseCompta.Id1, recDeviseCompta.Id2, recDeviseCompta.Amj, recDeviseCompta.HHMM, recDeviseCompta.Origine
                        If tableDeviseCompta.NoMatch Then
                            Error 9998
                        End If
     Case "MoveNext    "
                        tableDeviseCompta.MoveNext
                        If tableDeviseCompta.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableDeviseCompta.MovePrevious
                        If tableDeviseCompta.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst   "
                        tableDeviseCompta.MoveFirst
                        If tableDeviseCompta.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast    "
                        tableDeviseCompta.MoveLast
                        If tableDeviseCompta.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recDeviseCompta.Method <> "AddNew      " Then
    Call tableDeviseCompta_GetBuffer(recDeviseCompta)
End If

Exit Function

'---------------------------------------------------------
tableDeviseCompta_Read_Error:
'---------------------------------------------------------

    tableDeviseCompta_Read = Err
    Resume tableDeviseCompta_Read_End

tableDeviseCompta_Read_End:

End Function

'---------------------------------------------------------
Public Function tableDeviseCompta_Update(recDeviseCompta As typeDeviseChange) As Integer
'---------------------------------------------------------

On Error GoTo tableDeviseComptaUpdate_Error
tableDeviseCompta_Update = 0

Select Case recDeviseCompta.Method

    Case "AddNew      "
                        tableDeviseCompta.AddNew
                        Call tableDeviseCompta_PutBuffer(recDeviseCompta)
                        tableDeviseCompta.Update
    Case "Update      "
                        tableDeviseCompta.Edit
                        Call tableDeviseCompta_PutBuffer(recDeviseCompta)
                        tableDeviseCompta.Update
    Case "Delete      "
                        tableDeviseCompta.Delete
    Case Else
                        Error 9999
End Select


Exit Function

tableDeviseComptaUpdate_Error:
'---------------------------------------------------------
    tableDeviseCompta_Update = Err
    Resume tableDeviseComptaUpdate_End

tableDeviseComptaUpdate_End:

End Function








'-----------------------------------------------------
Sub dbDeviseCompta_Error(recDeviseCompta As typeDeviseChange)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "DeviseCompta: "

Select Case mId$(recDeviseCompta.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recDeviseCompta.Err: I = vbCritical
End Select

MsgBox Msg, I, "module : .bas  ( " _
                & Trim(recDeviseCompta.obj) & " : " & Trim(recDeviseCompta.Method) & " )"

End Sub

'-----------------------------------------------------
Function dbDeviseCompta_Read(recDeviseCompta As typeDeviseChange)
'-----------------------------------------------------

dbDeviseCompta_Read = Null

recDeviseCompta.Err = tableDeviseCompta_Read(recDeviseCompta)
If recDeviseCompta.Err > 0 Then

    If recDeviseCompta.Err < 9990 Or recDeviseCompta.Err >= 9999 Then
        Call dbDeviseCompta_Error(recDeviseCompta)
        dbDeviseCompta_Read = recDeviseCompta.Err
    End If
End If

End Function

'-----------------------------------------------------
Function dbDeviseCompta_Update(recDeviseCompta As typeDeviseChange)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
'$$$BeginTrans

dbDeviseCompta_Update = Null


recDeviseCompta.Err = tableDeviseCompta_Update(recDeviseCompta)

If recDeviseCompta.Err <> 0 Then
    Call dbDeviseCompta_Error(recDeviseCompta)
    dbDeviseCompta_Update = recDeviseCompta.Err
'$$$    Rollback
    Exit Function
End If

'$$$CommitTrans


'=====================================================
End Function


