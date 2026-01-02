Attribute VB_Name = "mdbDeviseChange"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableDeviseChange As Recordset
Public tableDeviseChangeOpen As Boolean
    
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
    CoursCompta   As String * 1

End Type

Public Sub dbDeviseChange_Replication()
Dim I As Integer, X As String, Amj As String


ReDim arrDeviseChange(10): arrDeviseChangeNbMax = 10
recDeviseChange_Init XDeviseChange
XDeviseChange.Method = "MoveLast"
XDeviseChange.Amj = "19990100"
tableDeviseChange_Read XDeviseChange
XDeviseChange.Method = "SnapP0"
XDeviseChange.Amj = XDeviseChange.Amj
XDeviseChange.Origine = "T"
XDeviseChange.Id1 = ""

arrDeviseChange(0) = XDeviseChange
arrDeviseChange(0).Amj = "99999999"
arrDeviseChangeSuite = True
Do Until Not arrDeviseChangeSuite
    arrDeviseChangeNb = 0
    srvDeviseChange_Monitor XDeviseChange
    For I = 1 To arrDeviseChangeNb
        X = Trim(arrDeviseChange(I).ValidationUsr)
        If X <> constàValider And X <> "" And arrDeviseChange(I).Origine <> "C" Then
            arrDeviseChange(I).Method = constAddNew
            tableDeviseChange_Update arrDeviseChange(I)
        End If
    Next I
    XDeviseChange = arrDeviseChange(arrDeviseChangeNb)
    XDeviseChange.Method = "SnapP0+"
Loop
End Sub


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableDeviseChange_Close()
'-----------------------------------------------------
If tableDeviseChangeOpen Then
    tableDeviseChange.Close
    tableDeviseChangeOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableDeviseChange_GetBuffer(recDeviseChange As typeDeviseChange)
'---------------------------------------------------------

recDeviseChange.Id1 = tableDeviseChange("ID1")
recDeviseChange.Id2 = tableDeviseChange("ID2")
recDeviseChange.Amj = tableDeviseChange("AMJ")
recDeviseChange.Origine = tableDeviseChange("Origine")
recDeviseChange.HHMM = tableDeviseChange("HHMM")
recDeviseChange.QD1 = tableDeviseChange("QD1")
recDeviseChange.QD2CoursPivot = tableDeviseChange("QD2CoursPivot")
recDeviseChange.QD2AchatNormal = tableDeviseChange("QD2AchatNormal")
recDeviseChange.QD2VenteNormal = tableDeviseChange("QD2VenteNormal")
recDeviseChange.QD2AchatPrivilégié = tableDeviseChange("QD2AchatPrivilégié")
recDeviseChange.QD2VentePrivilégié = tableDeviseChange("QD2VentePrivilégié")
recDeviseChange.QD2AchatEnCompte = tableDeviseChange("QD2AchatEnCompte")
recDeviseChange.QD2VenteEnCompte = tableDeviseChange("QD2VenteEnCompte")

recDeviseChange.SaisieAmj = tableDeviseChange("SaisieAMJ")
recDeviseChange.SaisieHMS = tableDeviseChange("SaisieHMS")
recDeviseChange.SaisieUsr = tableDeviseChange("SaisieUsr")
recDeviseChange.ValidationAMJ = tableDeviseChange("ValidationAMJ")
recDeviseChange.ValidationHMS = tableDeviseChange("ValidationHMS")
recDeviseChange.ValidationUsr = tableDeviseChange("ValidationUsr")

End Sub


'-----------------------------------------------------
Sub tableDeviseChange_Open()
'-----------------------------------------------------

If Not tableDeviseChangeOpen Then
    Set tableDeviseChange = MDB.OpenRecordset("DeviseChange")
    tableDeviseChange.Index = "PrimaryKey"
    tableDeviseChangeOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableDeviseChange_PutBuffer(recDeviseChange As typeDeviseChange)
'---------------------------------------------------------

tableDeviseChange("ID1") = recDeviseChange.Id1
tableDeviseChange("ID2") = recDeviseChange.Id2
tableDeviseChange("AMJ") = recDeviseChange.Amj
tableDeviseChange("Origine") = recDeviseChange.Origine
tableDeviseChange("HHMM") = recDeviseChange.HHMM

tableDeviseChange("QD1") = recDeviseChange.QD1
tableDeviseChange("QD2CoursPivot") = recDeviseChange.QD2CoursPivot
tableDeviseChange("QD2AchatNormal") = recDeviseChange.QD2AchatNormal
tableDeviseChange("QD2VenteNormal") = recDeviseChange.QD2VenteNormal
tableDeviseChange("QD2AchatPrivilégié") = recDeviseChange.QD2AchatPrivilégié
tableDeviseChange("QD2VentePrivilégié") = recDeviseChange.QD2VentePrivilégié
tableDeviseChange("QD2AchatEnCompte") = recDeviseChange.QD2AchatEnCompte
tableDeviseChange("QD2VenteEnCompte") = recDeviseChange.QD2VenteEnCompte

tableDeviseChange("SaisieAMJ") = recDeviseChange.SaisieAmj
tableDeviseChange("SaisieHMS") = recDeviseChange.SaisieHMS
tableDeviseChange("SaisieUsr") = recDeviseChange.SaisieUsr
tableDeviseChange("ValidationAMJ") = recDeviseChange.ValidationAMJ
tableDeviseChange("ValidationHMS") = recDeviseChange.ValidationHMS
tableDeviseChange("ValidationUsr") = recDeviseChange.ValidationUsr

End Sub


'---------------------------------------------------------
Public Function tableDeviseChange_Read(recDeviseChange As typeDeviseChange) As Integer
'---------------------------------------------------------

On Error GoTo tableDeviseChange_Read_Error
tableDeviseChange_Read = 0


Select Case recDeviseChange.Method
     Case "Seek=       "
                        tableDeviseChange.Seek "=", recDeviseChange.Id1, recDeviseChange.Id2, recDeviseChange.Amj, recDeviseChange.HHMM, recDeviseChange.Origine
                        If tableDeviseChange.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<=      "
                        tableDeviseChange.Seek "<=", recDeviseChange.Id1, recDeviseChange.Id2, recDeviseChange.Amj, recDeviseChange.HHMM, recDeviseChange.Origine
                        If tableDeviseChange.NoMatch Then
                            Error 9998
                        End If
     Case "MoveNext    "
                        tableDeviseChange.MoveNext
                        If tableDeviseChange.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableDeviseChange.MovePrevious
                        If tableDeviseChange.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst   "
                        tableDeviseChange.MoveFirst
                        If tableDeviseChange.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast    "
                        tableDeviseChange.MoveLast
                        If tableDeviseChange.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recDeviseChange.Method <> "AddNew      " Then
    Call tableDeviseChange_GetBuffer(recDeviseChange)
End If

Exit Function

'---------------------------------------------------------
tableDeviseChange_Read_Error:
'---------------------------------------------------------

    tableDeviseChange_Read = Err
    Resume tableDeviseChange_Read_End

tableDeviseChange_Read_End:

End Function

'---------------------------------------------------------
Public Function tableDeviseChange_Update(recDeviseChange As typeDeviseChange) As Integer
'---------------------------------------------------------

On Error GoTo tableDeviseChangeUpdate_Error
tableDeviseChange_Update = 0

Select Case recDeviseChange.Method

    Case "AddNew      "
                        tableDeviseChange.AddNew
                        Call tableDeviseChange_PutBuffer(recDeviseChange)
                        tableDeviseChange.Update
    Case "Update      "
                        tableDeviseChange.Edit
                        Call tableDeviseChange_PutBuffer(recDeviseChange)
                        tableDeviseChange.Update
    Case "Delete      "
                        tableDeviseChange.Delete
    Case Else
                        Error 9999
End Select


Exit Function

tableDeviseChangeUpdate_Error:
'---------------------------------------------------------
    tableDeviseChange_Update = Err
    Resume tableDeviseChangeUpdate_End

tableDeviseChangeUpdate_End:

End Function








'-----------------------------------------------------
Sub dbDeviseChange_Error(recDeviseChange As typeDeviseChange)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "DeviseChange: "

Select Case mId$(recDeviseChange.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recDeviseChange.Err: I = vbCritical
End Select

MsgBox Msg, I, "module : .bas  ( " _
                & Trim(recDeviseChange.obj) & " : " & Trim(recDeviseChange.Method) & " )"

End Sub

'-----------------------------------------------------
Function dbDeviseChange_Read(recDeviseChange As typeDeviseChange)
'-----------------------------------------------------

dbDeviseChange_Read = Null

recDeviseChange.Err = tableDeviseChange_Read(recDeviseChange)
If recDeviseChange.Err > 0 Then

    If recDeviseChange.Err < 9990 Or recDeviseChange.Err >= 9999 Then
        Call dbDeviseChange_Error(recDeviseChange)
        dbDeviseChange_Read = recDeviseChange.Err
    End If
End If

End Function

'-----------------------------------------------------
Function dbDeviseChange_Update(recDeviseChange As typeDeviseChange)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
'$$$BeginTrans

dbDeviseChange_Update = Null


recDeviseChange.Err = tableDeviseChange_Update(recDeviseChange)

If recDeviseChange.Err <> 0 Then
    Call dbDeviseChange_Error(recDeviseChange)
    dbDeviseChange_Update = recDeviseChange.Err
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
    CV.DeviseN = "000"
    DevCode CV.DeviseIso
Else
    CV.DeviseIso = ""
    DevX CV.DeviseN
End If

If Trim(XDevise.DevX) <> "" Then
    CV.DeviseIso = XDevise.DevX
    CV.DeviseN = Format$(XDevise.DevCode, "000")
    CV.DeviseLibellé = XDevise.DevLib
Else
    CV_Attribut = "? devise : " & CV.DeviseIso & " " & CV.DeviseN: Exit Function
End If

If CV.DeviseIso = "ITL" Or CV.DeviseIso = "GRD" Or CV.DeviseIso = "PTE" Or CV.DeviseIso = "ESP" _
           Or CV.DeviseIso = "BEF" Or CV.DeviseIso = "LUF" Or CV.DeviseIso = "JPY" _
           Or CV.DeviseIso = "CFA" Or CV.DeviseIso = "XAF" Or CV.DeviseIso = "XOF" Then
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
Dim dblMontant As Double, blnNegatif As Boolean

CV_Calc = Null
If CV1.Montant < 0 Then
    blnNegatif = True
    CV1.Montant = -CV1.Montant
Else
    blnNegatif = False
End If

Select Case CV1.maxD
    Case 0:
        CV1.Montant = Fix(CV1.Montant)
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
        If CV2.Cours = 0 Then
            dblMontant = 0
        Else
            dblMontant = dblMontant / CV2.Cours
        End If
    End If
End If


Select Case CV2.maxD
    Case 0: CV2.Montant = Fix(dblMontant + 0.5000001)
    Case Else: CV2.Montant = Fix((dblMontant + 0.00500001) * 100) / 100
End Select
If blnNegatif Then
    CV1.Montant = -CV1.Montant
    CV2.Montant = -CV2.Montant
    CV3.Montant = -CV3.Montant
End If

End Function

Private Function CV_Cours(CV1 As typeCV, CV3 As typeCV)
Dim recDeviseChange As typeDeviseChange, intReturn As Integer
Dim mId1 As String, mId2 As String

CV_Cours = Null
recDeviseChange_Init recDeviseChange
If CV3.DeviseIso = CV1.DeviseIso Then
    CV1.Cours = 1
    CV1.CoursAmj = CV1.OpéAmj
    Exit Function
End If

If CV1.CotationCertain Then
    recDeviseChange.Id1 = CV3.DeviseIso
    recDeviseChange.Id2 = CV1.DeviseIso
 Else
    recDeviseChange.Id2 = CV3.DeviseIso
    recDeviseChange.Id1 = CV1.DeviseIso
End If
recDeviseChange.Amj = CV1.OpéAmj
recDeviseChange.HHMM = 9999
recDeviseChange.Origine = "Z"

mId1 = recDeviseChange.Id1
mId2 = recDeviseChange.Id2

recDeviseChange.Method = "Seek<="
If CV1.CoursCompta = "C" Then
'    recDeviseChange.Method = "LastLC"
'    recDeviseChange.Origine = "C"
'    If Not IsNull(srvDeviseChange_Monitor(recDeviseChange)) Then Exit Function
    intReturn = tableDeviseCompta_Read(recDeviseChange)
Else
    intReturn = tableDeviseChange_Read(recDeviseChange)
End If

If intReturn <> 0 Then
    CV_Cours = mId1 & " / " & mId2 & "/" & CV1.OpéAmj & " : Seek<= ?"
    Exit Function
Else
    If mId1 <> recDeviseChange.Id1 Or mId2 <> recDeviseChange.Id2 Then
        CV_Cours = mId1 & " / " & mId2 & "/" & CV1.OpéAmj & " : non trouvé ?"
        Exit Function
    End If
End If

If recDeviseChange.Amj < CV1.CoursAmjMin Then
    CV_Cours = mId1 & " / " & mId2 & "/" & recDeviseChange.Amj & " < " & CV1.CoursAmjMin
    Exit Function
End If

CV1.CoursAmj = recDeviseChange.Amj

Select Case CV1.AchatVente
    Case "A": Select Case CV1.Normal
                Case "N": CV1.Cours = recDeviseChange.QD2AchatNormal
                Case "P": CV1.Cours = recDeviseChange.QD2AchatPrivilégié
                Case "C": CV1.Cours = recDeviseChange.QD2AchatEnCompte
                Case Else: CV1.Cours = recDeviseChange.QD2CoursPivot
             End Select
    Case "V": Select Case CV1.Normal
                Case "N": CV1.Cours = recDeviseChange.QD2VenteNormal
                Case "P": CV1.Cours = recDeviseChange.QD2VentePrivilégié
                Case "C": CV1.Cours = recDeviseChange.QD2VenteEnCompte
                Case Else: CV1.Cours = recDeviseChange.QD2CoursPivot
               End Select
    Case Else: CV1.Cours = recDeviseChange.QD2CoursPivot
End Select

If recDeviseChange.QD1 <> 1 Then CV1.Cours = CV1.Cours / recDeviseChange.QD1

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
Public Function CV_Manuel(CV1 As typeCV, CV2 As typeCV, CV3 As typeCV, Conversion As String, lCotation As String)
Dim errorV

CV2.Montant = 0: CV3.Montant = 0
Conversion = " "
errorV = CV_Attribut(CV1): If Not IsNull(errorV) Then GoTo CV_Transitoire_Error
errorV = CV_Attribut(CV2): If Not IsNull(errorV) Then GoTo CV_Transitoire_Error

CV3 = CV_Euro


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
If lCotation = "*" Then
    CV2.CotationCertain = True
Else
    CV2.CotationCertain = False
End If

CV_Manuel = CV_Calc(CV1, CV2, CV3)

Exit Function

'---------------------------------------------------------
CV_Transitoire_Error:
'---------------------------------------------------------
'''''    MsgBox errorV, vbCritical
CV_Manuel = errorV


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

