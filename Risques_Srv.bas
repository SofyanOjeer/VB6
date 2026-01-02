Attribute VB_Name = "srvRisques"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------


Type typeRisques
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Id              As String * 11
    AM              As String * 6
    Intitulé        As String * 40
    CO              As Currency
    CA              As Currency
    CC              As Currency
    CD              As Currency
    TE              As Currency
    TA              As Currency
    TD              As Currency
    IT              As Currency
    AC              As Currency
    OC              As Currency
    OD              As Currency
    BM              As Currency
    BI              As Currency
    
  End Type
    
Public arrRisques() As typeRisques
Public arrRisquesNb As Integer
Public arrRisquesNbMax As Integer
Public arrRisquesIndex As Integer
Public arrRisquesSuite As Boolean

Public recRisques   As typeRisques
Public totalRisques() As typeRisques

'-----------------------------------------------------
Sub tableRisques_Close()
'-----------------------------------------------------
If tableRisquesOpen Then
    tableRisques.Close
    tableRisquesOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableRisques_GetBuffer(recRisques As typeRisques)
'---------------------------------------------------------

recRisques.Id = tableRisques("ID")
recRisques.Civilité = tableRisques("Civilité")
recRisques.Nom = tableRisques("Nom")
recRisques.Prénoms = tableRisques("Prénoms")
recRisques.Tél1 = tableRisques("Tél1")
recRisques.Tél2 = tableRisques("Tél2")
recRisques.Tél3 = tableRisques("Tél3")
recRisques.MicroSN = tableRisques("MicroSN")
recRisques.MicroIP = tableRisques("MicroIP")
recRisques.Service = tableRisques("Service")
recRisques.Bureau = tableRisques("Bureau")

End Sub


'-----------------------------------------------------
Sub tableRisques_Open()
'-----------------------------------------------------

If Not tableRisquesOpen Then
    Set tableRisques = MDB.OpenRecordset("Risques")
    tableRisques.Index = "PrimaryKey"
    tableRisquesOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableRisques_PutBuffer(recRisques As typeRisques)
'---------------------------------------------------------
Dim X As String

tableRisques("ID") = recRisques.Id
tableRisques("Civilité") = recRisques.Civilité
tableRisques("Nom") = recRisques.Nom
tableRisques("Prénoms") = recRisques.Prénoms
tableRisques("Tél1") = recRisques.Tél1
tableRisques("Tél2") = recRisques.Tél2
tableRisques("Tél3") = recRisques.Tél3
tableRisques("MicroSN") = recRisques.MicroSN
tableRisques("MicroIP") = recRisques.MicroIP
tableRisques("Service") = recRisques.Service
tableRisques("Bureau") = recRisques.Bureau

End Sub


'---------------------------------------------------------
Public Function tableRisques_Read(recRisques As typeRisques) As Integer
'---------------------------------------------------------

On Error GoTo tableRisques_Read_Error
tableRisques_Read = 0


Select Case recRisques.Method
     Case "Seek=       "
                        tableRisques.Seek "=", recRisques.Id
                        If tableRisques.NoMatch Then
                            Error 9998
                        End If
     Case "MoveNext    "
                        tableRisques.MoveNext
                        If tableRisques.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableRisques.MovePrevious
                        If tableRisques.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst   "
                        tableRisques.MoveFirst
                        If tableRisques.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast    "
                        tableRisques.MoveLast
                        If tableRisques.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recRisques.Method <> "AddNew      " Then
    Call tableRisques_GetBuffer(recRisques)
End If

Exit Function

'---------------------------------------------------------
tableRisques_Read_Error:
'---------------------------------------------------------

    tableRisques_Read = Err
    Resume tableRisques_Read_End

tableRisques_Read_End:

End Function

'---------------------------------------------------------
Public Function tableRisques_Update(recRisques As typeRisques) As Integer
'---------------------------------------------------------

On Error GoTo tableRisquesUpdate_Error
tableRisques_Update = 0

Select Case recRisques.Method

    Case "AddNew      "
                        tableRisques.AddNew
                        Call tableRisques_PutBuffer(recRisques)
                        tableRisques.Update
    Case "Update      "
                        tableRisques.Edit
                        Call tableRisques_PutBuffer(recRisques)
                        tableRisques.Update
    Case "Delete      "
                        tableRisques.Delete
    Case Else
                        Error 9999
End Select


Exit Function

tableRisquesUpdate_Error:
'---------------------------------------------------------
    tableRisques_Update = Err
    Resume tableRisquesUpdate_End

tableRisquesUpdate_End:

End Function








'-----------------------------------------------------
Sub dbRisques_Error(recRisques As typeRisques)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Risques: "

Select Case Mid$(recRisques.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recRisques.Err: I = vbCritical
End Select

MsgBox Msg, I, "module : .bas  ( " _
                & Trim(recRisques.obj) & " : " & Trim(recRisques.Method) & " )"

End Sub

'-----------------------------------------------------
Function dbRisques_Read(recRisques As typeRisques)
'-----------------------------------------------------

dbRisques_Read = Null

recRisques.Err = tableRisques_Read(recRisques)
If recRisques.Err > 0 Then

    If recRisques.Err < 9990 Or recRisques.Err >= 9999 Then
        Call dbRisques_Error(recRisques)
        dbRisques_Read = recRisques.Err
    End If
End If

End Function

'---------------------------------------------------------
Public Sub arrRisques_Load()
'---------------------------------------------------------
Dim iRead As Integer

tableRisques_Open
arrRisquesNb = 0: arrRisquesNbMax = 0

recRisques.Method = "MoveFirst"
recRisques.Id = String$(4, Chr$(0))
recRisques.obj = "Risques"
recRisques.Err = 0

iRead = tableRisques_Read(recRisques)
Do While iRead = 0

    arrRisques_AddItem recRisques

    recRisques.Method = "MoveNext    "
    iRead = tableRisques_Read(recRisques)
Loop
tableRisques_Close

End Sub


'-----------------------------------------------------
Function dbRisques_Update(recRisques As typeRisques)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
'$$$BeginTrans

dbRisques_Update = Null


recRisques.Err = tableRisques_Update(recRisques)

If recRisques.Err <> 0 Then
    Call dbRisques_Error(recRisques)
    dbRisques_Update = recRisques.Err
'$$$    Rollback
    Exit Function
End If

'$$$CommitTrans


'=====================================================
End Function

'---------------------------------------------------------
Public Sub arrRisques_AddItem(recRisques As typeRisques)
'---------------------------------------------------------
          
arrRisquesNb = arrRisquesNb + 1
    
If arrRisquesNb > arrRisquesNbMax Then
    arrRisquesNbMax = arrRisquesNbMax + 50
    ReDim Preserve arrRisques(arrRisquesNbMax)
End If
recRisques.Method = ""
arrRisquesIndex = arrRisquesNb
arrRisques(arrRisquesIndex) = recRisques
End Sub





Public Sub arrRisques_Scan(Msg As String)
Dim I As Integer, I1 As Integer, xNom As String, xPrénoms As String, xTél1 As String

I = InStr(1, Msg, ":")
xNom = Trim(Mid$(Msg, 1, I - 1))
I1 = InStr(I + 1, Msg, ":")
xPrénoms = Trim(Mid$(Msg, I + 1, I1 - I - 1))
I = InStr(I1 + 1, Msg, ":")
xTél1 = Trim(Mid$(Msg, I1 + 1, I - I1 - 1))

For arrRisquesIndex = 1 To arrRisquesNb
    If Trim(arrRisques(arrRisquesIndex).Nom) = xNom _
    And Trim(arrRisques(arrRisquesIndex).Prénoms) = xPrénoms _
    And Trim(arrRisques(arrRisquesIndex).Tél1) = xTél1 Then
        recRisques = arrRisques(arrRisquesIndex)
        Exit For
    End If
Next arrRisquesIndex

End Sub

Public Sub recRisques_Init(recRisques As typeRisques)
recRisques.obj = "Risques"
recRisques.Method = ""
recRisques.Err = 0
    
recRisques.Id = ""
recRisques.Civilité = "1"
recRisques.Id = ""
recRisques.Nom = ""
recRisques.Prénoms = ""
recRisques.Tél1 = ""
recRisques.Tél2 = ""
recRisques.Tél3 = ""
recRisques.MicroSN = ""
recRisques.MicroIP = ""
recRisques.Service = ""
recRisques.Bureau = ""

End Sub

Public Sub arrRisques_Test()
Dim I As Integer
ReDim arrRisques(12)
ReDim totalRisques(12)

For I = 1 To 12
    recRisques.Id = "12345678901"
    recRisques.Intitulé = "TEST RISQUES"
    recRisques.AM = "1998" & Format$(I, "00")
    recRisques.CO = I * 100 + 1
    recRisques.CA = I * 100 + 2
    recRisques.CC = I * 100 + 3
    recRisques.CD = I * 100 + 4
    recRisques.TE = I * 100 + 5
    recRisques.TA = I * 100 + 6
    recRisques.TD = I * 100 + 7
    recRisques.IT = I * 100 + 8
    recRisques.AC = I * 100 + 9
    recRisques.OC = I * 100 + 10
    recRisques.OD = I * 100 + 11
    recRisques.BM = I * 100 + 12
    recRisques.BI = I * 100 + 13
    arrRisques(I) = recRisques
    
    recRisques.Id = "12345678901"
    recRisques.Intitulé = "TEST RISQUES"
    recRisques.AM = "1998" & Format$(I, "00")
    recRisques.CO = I * 200 + 1
    recRisques.CA = I * 200 + 2
    recRisques.CC = I * 200 + 3
    recRisques.CD = I * 200 + 4
    recRisques.TE = I * 200 + 5
    recRisques.TA = I * 200 + 6
    recRisques.TD = I * 200 + 7
    recRisques.IT = I * 200 + 8
    recRisques.AC = I * 200 + 9
    recRisques.OC = I * 200 + 10
    recRisques.OD = I * 200 + 11
    recRisques.BM = I * 200 + 12
    recRisques.BI = I * 200 + 13
    totalRisques(I) = recRisques
Next I

End Sub
