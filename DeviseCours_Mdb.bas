Attribute VB_Name = "mdbDeviseCours"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableDeviseCours As Recordset
Public tableDeviseCoursOpen As Boolean
    
Type typeCV
    DeviseIso     As String * 3
    DeviseN       As String * 3
    DeviseLibellé As String * 20
    Cours         As Double
    CoursAmj      As String * 8
    maxD          As String * 1
    EuroIn        As Boolean
    CotationCertain As Boolean
    
    Montant       As Currency
    OpéAmj        As String * 8
    CoursAmjMin   As String * 8
    AchatVente    As String * 1
    Normal        As String * 1

End Type

Public Sub dbDeviseCours_Replication(Amj As String)
Dim I As Integer, X As String


ReDim arrDeviseCours(10): arrDeviseCoursNbMax = 10
recDeviseCours_Init XDeviseCours
XDeviseCours.Method = "SnapP0"
XDeviseCours.Amj = Amj

arrDeviseCours(0) = XDeviseCours
arrDeviseCours(0).Amj = "99999999"
arrDeviseCoursSuite = True
Do Until Not arrDeviseCoursSuite
    arrDeviseCoursNb = 0
    srvDeviseCours_Monitor XDeviseCours
    For I = 1 To arrDeviseCoursNb
        X = Trim(arrDeviseCours(I).ValidationUsr)
        If X <> constàValider And X <> "" Then
            arrDeviseCours(I).Method = "AddNew"
            tableDeviseCours_Update arrDeviseCours(I)
        End If
    Next I
    XDeviseCours = arrDeviseCours(arrDeviseCoursNb)
    XDeviseCours.Method = "SnapP0+"
Loop
End Sub


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableDeviseCours_Close()
'-----------------------------------------------------
If tableDeviseCoursOpen Then
    tableDeviseCours.Close
    tableDeviseCoursOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableDeviseCours_GetBuffer(recDeviseCours As typeDeviseCours)
'---------------------------------------------------------

recDeviseCours.Id1 = tableDeviseCours("ID1")
recDeviseCours.Id2 = tableDeviseCours("ID2")
recDeviseCours.Amj = tableDeviseCours("AMJ")

recDeviseCours.QD1 = tableDeviseCours("QD1")
recDeviseCours.QD2CoursPivot = tableDeviseCours("QD2CoursPivot")
recDeviseCours.QD2AchatNormal = tableDeviseCours("QD2AchatNormal")
recDeviseCours.QD2VenteNormal = tableDeviseCours("QD2VenteNormal")
recDeviseCours.QD2AchatPrivilégié = tableDeviseCours("QD2AchatPrivilégié")
recDeviseCours.QD2VentePrivilégié = tableDeviseCours("QD2VentePrivilégié")
recDeviseCours.QD2AchatEnCompte = tableDeviseCours("QD2AchatEnCompte")
recDeviseCours.QD2VenteEnCompte = tableDeviseCours("QD2VenteEnCompte")

recDeviseCours.SaisieAMJ = tableDeviseCours("SaisieAMJ")
recDeviseCours.SaisieHMS = tableDeviseCours("SaisieHMS")
recDeviseCours.SaisieUsr = tableDeviseCours("SaisieUsr")
recDeviseCours.ValidationAMJ = tableDeviseCours("ValidationAMJ")
recDeviseCours.ValidationHMS = tableDeviseCours("ValidationHMS")
recDeviseCours.ValidationUsr = tableDeviseCours("ValidationUsr")

End Sub


'-----------------------------------------------------
Sub tableDeviseCours_Open()
'-----------------------------------------------------

If Not tableDeviseCoursOpen Then
    Set tableDeviseCours = MDB.OpenRecordset("DeviseCours")
    tableDeviseCours.Index = "PrimaryKey"
    tableDeviseCoursOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableDeviseCours_PutBuffer(recDeviseCours As typeDeviseCours)
'---------------------------------------------------------

tableDeviseCours("ID1") = recDeviseCours.Id1
tableDeviseCours("ID2") = recDeviseCours.Id2
tableDeviseCours("AMJ") = recDeviseCours.Amj

tableDeviseCours("QD1") = recDeviseCours.QD1
tableDeviseCours("QD2CoursPivot") = recDeviseCours.QD2CoursPivot
tableDeviseCours("QD2AchatNormal") = recDeviseCours.QD2AchatNormal
tableDeviseCours("QD2VenteNormal") = recDeviseCours.QD2VenteNormal
tableDeviseCours("QD2AchatPrivilégié") = recDeviseCours.QD2AchatPrivilégié
tableDeviseCours("QD2VentePrivilégié") = recDeviseCours.QD2VentePrivilégié
tableDeviseCours("QD2AchatEnCompte") = recDeviseCours.QD2AchatEnCompte
tableDeviseCours("QD2VenteEnCompte") = recDeviseCours.QD2VenteEnCompte

tableDeviseCours("SaisieAMJ") = recDeviseCours.SaisieAMJ
tableDeviseCours("SaisieHMS") = recDeviseCours.SaisieHMS
tableDeviseCours("SaisieUsr") = recDeviseCours.SaisieUsr
tableDeviseCours("ValidationAMJ") = recDeviseCours.ValidationAMJ
tableDeviseCours("ValidationHMS") = recDeviseCours.ValidationHMS
tableDeviseCours("ValidationUsr") = recDeviseCours.ValidationUsr

End Sub


'---------------------------------------------------------
Public Function tableDeviseCours_Read(recDeviseCours As typeDeviseCours) As Integer
'---------------------------------------------------------

On Error GoTo tableDeviseCours_Read_Error
tableDeviseCours_Read = 0


Select Case recDeviseCours.Method
     Case "Seek=       "
                        tableDeviseCours.Seek "=", recDeviseCours.Id1, recDeviseCours.Id2, recDeviseCours.Amj
                        If tableDeviseCours.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<=      "
                        tableDeviseCours.Seek "<=", recDeviseCours.Id1, recDeviseCours.Id2, recDeviseCours.Amj
                        If tableDeviseCours.NoMatch Then
                            Error 9998
                        End If
     Case "MoveNext    "
                        tableDeviseCours.MoveNext
                        If tableDeviseCours.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableDeviseCours.MovePrevious
                        If tableDeviseCours.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst   "
                        tableDeviseCours.MoveFirst
                        If tableDeviseCours.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast    "
                        tableDeviseCours.MoveLast
                        If tableDeviseCours.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recDeviseCours.Method <> "AddNew      " Then
    Call tableDeviseCours_GetBuffer(recDeviseCours)
End If

Exit Function

'---------------------------------------------------------
tableDeviseCours_Read_Error:
'---------------------------------------------------------

    tableDeviseCours_Read = Err
    Resume tableDeviseCours_Read_End

tableDeviseCours_Read_End:

End Function

'---------------------------------------------------------
Public Function tableDeviseCours_Update(recDeviseCours As typeDeviseCours) As Integer
'---------------------------------------------------------

On Error GoTo tableDeviseCoursUpdate_Error
tableDeviseCours_Update = 0

Select Case recDeviseCours.Method

    Case "AddNew      "
                        tableDeviseCours.AddNew
                        Call tableDeviseCours_PutBuffer(recDeviseCours)
                        tableDeviseCours.Update
    Case "Update      "
                        tableDeviseCours.Edit
                        Call tableDeviseCours_PutBuffer(recDeviseCours)
                        tableDeviseCours.Update
    Case "Delete      "
                        tableDeviseCours.Delete
    Case Else
                        Error 9999
End Select


Exit Function

tableDeviseCoursUpdate_Error:
'---------------------------------------------------------
    tableDeviseCours_Update = Err
    Resume tableDeviseCoursUpdate_End

tableDeviseCoursUpdate_End:

End Function








'-----------------------------------------------------
Sub dbDeviseCours_Error(recDeviseCours As typeDeviseCours)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "DeviseCours: "

Select Case Mid$(recDeviseCours.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recDeviseCours.Err: I = vbCritical
End Select

MsgBox Msg, I, "module : .bas  ( " _
                & Trim(recDeviseCours.obj) & " : " & Trim(recDeviseCours.Method) & " )"

End Sub

'-----------------------------------------------------
Function dbDeviseCours_Read(recDeviseCours As typeDeviseCours)
'-----------------------------------------------------

dbDeviseCours_Read = Null

recDeviseCours.Err = tableDeviseCours_Read(recDeviseCours)
If recDeviseCours.Err > 0 Then

    If recDeviseCours.Err < 9990 Or recDeviseCours.Err >= 9999 Then
        Call dbDeviseCours_Error(recDeviseCours)
        dbDeviseCours_Read = recDeviseCours.Err
    End If
End If

End Function

'-----------------------------------------------------
Function dbDeviseCours_Update(recDeviseCours As typeDeviseCours)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
'$$$BeginTrans

dbDeviseCours_Update = Null


recDeviseCours.Err = tableDeviseCours_Update(recDeviseCours)

If recDeviseCours.Err <> 0 Then
    Call dbDeviseCours_Error(recDeviseCours)
    dbDeviseCours_Update = recDeviseCours.Err
'$$$    Rollback
    Exit Function
End If

'$$$CommitTrans


'=====================================================
End Function


Public Function CV_Attribut(CV As typeCV)

CV_Attribut = Null
CV.maxD = 2
CV.EuroIn = False
CV.CotationCertain = True

If Trim(CV.DeviseIso) <> "" Then
    DevCode CV.DeviseIso
Else
    DevX CV.DeviseN
End If

If Trim(XDevise.DevX) <> "" Then
    CV.DeviseIso = XDevise.DevX
    CV.DeviseN = Format$(XDevise.DevCode, "000")
    CV.DeviseLibellé = XDevise.DevLib
Else
    CV_Attribut = CV.DeviseLibellé: Exit Function
End If

If CV.DeviseIso = "ITL" Or CV.DeviseIso = "GRD" Or CV.DeviseIso = "PTE" Or CV.DeviseIso = "ESP" _
           Or CV.DeviseIso = "BEF" Or CV.DeviseIso = "LUF" Or CV.DeviseIso = "JPY" Then
    CV.maxD = 0
End If


If CV.DeviseIso = "FRF" Or CV.DeviseIso = "DEM" Or CV.DeviseIso = "ITL" Or CV.DeviseIso = "IEP" _
           Or CV.DeviseIso = "ESP" Or CV.DeviseIso = "PTE" Or CV.DeviseIso = "ATS" Or CV.DeviseIso = "FIM" _
           Or CV.DeviseIso = "BEF" Or CV.DeviseIso = "LUF" Or CV.DeviseIso = "NLG" Then
    CV.EuroIn = True
''' CV.CotationCertain = True
End If

End Function

Public Sub CV_Init(CV As typeCV)

CV.Montant = 0
CV.DeviseIso = ""
CV.DeviseN = ""
CV.DeviseLibellé = ""
CV.Cours = 0
CV.maxD = 2
CV.EuroIn = False
CV.CotationCertain = True

CV.AchatVente = "A"
CV.Normal = " "
CV.CoursAmj = "00000000"
CV.OpéAmj = DSys
CV.CoursAmjMin = CV_Euro.CoursAmjMin

End Sub

Public Function CV_Calc(CV1 As typeCV, CV2 As typeCV, CV3 As typeCV)
Dim dblMontant As Double
CV_Calc = Null
Select Case CV1.maxD
    Case 0: CV1.Montant = Fix(CV1.Montant)
    Case Else: CV1.Montant = Fix(CV1.Montant * 100) / 100
End Select

If CV1.Cours = 1 Then
    CV3.Montant = CV1.Montant
    dblMontant = CV1.Montant
Else
    If CV1.CotationCertain Then
        dblMontant = CV1.Montant / CV1.Cours
    Else
        dblMontant = CV1.Montant * CV1.Cours
    End If
    Select Case CV3.maxD
        Case 0: CV3.Montant = Fix(dblMontant + 0.5000001)
        Case Else: CV3.Montant = Fix((dblMontant + 0.00500001) * 100) / 100
    End Select

End If
       

If CV2.Cours = 1 Then
    CV2.Montant = CV3.Montant
Else
    If CV2.CotationCertain Then
        dblMontant = dblMontant * CV2.Cours
    Else
        dblMontant = dblMontant / CV2.Cours
    End If
End If


Select Case CV2.maxD
    Case 0: CV2.Montant = Fix(dblMontant + 0.5000001)
    Case Else: CV2.Montant = Fix((dblMontant + 0.00500001) * 100) / 100
End Select

End Function

Private Function CV_Cours(CV1 As typeCV, CV3 As typeCV)
Dim recDeviseCours As typeDeviseCours
Dim mId1 As String, mId2 As String

CV_Cours = Null
If CV3.DeviseIso = CV1.DeviseIso Then
    CV1.Cours = 1
    CV1.CoursAmj = CV1.OpéAmj
    Exit Function
End If

If CV1.CotationCertain Then
    recDeviseCours.Id1 = CV3.DeviseIso
    recDeviseCours.Id2 = CV1.DeviseIso
 Else
    recDeviseCours.Id2 = CV3.DeviseIso
    recDeviseCours.Id1 = CV1.DeviseIso
End If
recDeviseCours.Amj = CV1.OpéAmj

mId1 = recDeviseCours.Id1
mId2 = recDeviseCours.Id2

recDeviseCours.Method = "Seek<="
If tableDeviseCours_Read(recDeviseCours) <> 0 Then
    CV_Cours = mId1 & " / " & mId2 & "/" & CV1.OpéAmj & " : Seek<= ?"
    Exit Function
Else
    If mId1 <> recDeviseCours.Id1 Or mId2 <> recDeviseCours.Id2 Then
        CV_Cours = mId1 & " / " & mId2 & "/" & CV1.OpéAmj & " : non trouvé ?"
        Exit Function
    End If
End If

If recDeviseCours.Amj < CV1.CoursAmjMin Then
    CV_Cours = mId1 & " / " & mId2 & "/" & recDeviseCours.Amj & " < " & CV1.CoursAmjMin
    Exit Function
End If
CV1.CoursAmj = recDeviseCours.Amj

Select Case CV1.AchatVente
    Case "A": Select Case CV1.Normal
                Case "N": CV1.Cours = recDeviseCours.QD2AchatNormal
                Case "P": CV1.Cours = recDeviseCours.QD2AchatPrivilégié
                Case "C": CV1.Cours = recDeviseCours.QD2AchatEnCompte
                Case Else: CV1.Cours = recDeviseCours.QD2CoursPivot
             End Select
    Case "V": Select Case CV1.Normal
                Case "N": CV1.Cours = recDeviseCours.QD2VenteNormal
                Case "P": CV1.Cours = recDeviseCours.QD2VentePrivilégié
                Case "C": CV1.Cours = recDeviseCours.QD2VenteEnCompte
                Case Else: CV1.Cours = recDeviseCours.QD2CoursPivot
               End Select
    Case Else: CV1.Cours = recDeviseCours.QD2CoursPivot
End Select

If recDeviseCours.QD1 <> 1 Then CV1.Cours = CV1.Cours / recDeviseCours.QD1

End Function

Public Function CV_Transitoire(CV1 As typeCV, CV2 As typeCV, CV3 As typeCV, Conversion As String)
Dim errorV

CV2.Montant = 0: CV3.Montant = 0
Conversion = " "
errorV = CV_Attribut(CV1): If Not IsNull(errorV) Then GoTo CV_Transitoire_Error
errorV = CV_Attribut(CV2): If Not IsNull(errorV) Then GoTo CV_Transitoire_Error

CV3 = CV_Euro

errorV = CV_Cours(CV1, CV3): If Not IsNull(errorV) Then GoTo CV_Transitoire_Error
errorV = CV_Cours(CV2, CV3): If Not IsNull(errorV) Then GoTo CV_Transitoire_Error

If CV1.EuroIn Then
    If CV2.EuroIn Or CV2.DeviseIso = "EUR" Then
        Conversion = "C"
    Else
        Conversion = "B"
    End If
Else
    If CV1.DeviseIso = "EUR" Then
        If CV2.EuroIn Or CV2.DeviseIso = "EUR" Then
            Conversion = "C"
        Else
            Conversion = "A"
        End If
    Else
        If CV2.EuroIn Then
            Conversion = "B"
        Else
            Conversion = "A"
        End If
    End If
End If

CV_Transitoire = CV_Calc(CV1, CV2, CV3)

Exit Function

'---------------------------------------------------------
CV_Transitoire_Error:
'---------------------------------------------------------
'''''    MsgBox errorV, vbCritical
    CV_Transitoire = errorV


End Function

Public Function CV_AttributN(CV As typeCV)
CV.DeviseIso = ""
CV_AttributN = CV_Attribut(CV)
End Function
Public Function CV_AttributS(X As String, CV As typeCV)

If IsNumeric(X) Then
    CV.DeviseIso = ""
    CV.DeviseN = Format$(Val(X), "000")
Else
    CV.DeviseIso = X
End If

CV_AttributS = CV_Attribut(CV)

End Function

