Attribute VB_Name = "srvAnnuaire"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableAnnuaire As Recordset
Public tableAnnuaireOpen As Boolean


Type typeAnnuaire
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Id              As String * 4
    Civilité        As String * 1
    Nom             As String * 40
    Prénoms         As String * 40
    Tél1            As String * 3
    Tél2            As String * 3
    Tél3            As String * 3
    MicroSN         As String * 16
    MicroIP         As String * 12
    Service         As String * 3
    Bureau          As String * 5
    
  End Type
    
Public arrAnnuaire() As typeAnnuaire
Public arrAnnuaireNb As Integer
Public arrAnnuaireNbMax As Integer
Public arrAnnuaireIndex As Integer
Public arrAnnuaireSuite As Boolean

Public recAnnuaire   As typeAnnuaire
'-----------------------------------------------------
Sub tableAnnuaire_Close()
'-----------------------------------------------------
If tableAnnuaireOpen Then
    tableAnnuaire.Close
    tableAnnuaireOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableAnnuaire_GetBuffer(recAnnuaire As typeAnnuaire)
'---------------------------------------------------------

recAnnuaire.Id = tableAnnuaire("ID")
recAnnuaire.Civilité = tableAnnuaire("Civilité")
recAnnuaire.Nom = tableAnnuaire("Nom")
recAnnuaire.Prénoms = tableAnnuaire("Prénoms")
recAnnuaire.Tél1 = tableAnnuaire("Tél1")
recAnnuaire.Tél2 = tableAnnuaire("Tél2")
recAnnuaire.Tél3 = tableAnnuaire("Tél3")
recAnnuaire.MicroSN = tableAnnuaire("MicroSN")
recAnnuaire.MicroIP = tableAnnuaire("MicroIP")
recAnnuaire.Service = tableAnnuaire("Service")
recAnnuaire.Bureau = tableAnnuaire("Bureau")

End Sub


'-----------------------------------------------------
Sub tableAnnuaire_Open()
'-----------------------------------------------------

If Not tableAnnuaireOpen Then
    Set tableAnnuaire = MDB.OpenRecordset("Annuaire")
    tableAnnuaire.Index = "PrimaryKey"
    tableAnnuaireOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableAnnuaire_PutBuffer(recAnnuaire As typeAnnuaire)
'---------------------------------------------------------
Dim X As String

tableAnnuaire("ID") = recAnnuaire.Id
tableAnnuaire("Civilité") = recAnnuaire.Civilité
tableAnnuaire("Nom") = recAnnuaire.Nom
tableAnnuaire("Prénoms") = recAnnuaire.Prénoms
tableAnnuaire("Tél1") = recAnnuaire.Tél1
tableAnnuaire("Tél2") = recAnnuaire.Tél2
tableAnnuaire("Tél3") = recAnnuaire.Tél3
tableAnnuaire("MicroSN") = recAnnuaire.MicroSN
tableAnnuaire("MicroIP") = recAnnuaire.MicroIP
tableAnnuaire("Service") = recAnnuaire.Service
tableAnnuaire("Bureau") = recAnnuaire.Bureau

End Sub


'---------------------------------------------------------
Public Function tableAnnuaire_Read(recAnnuaire As typeAnnuaire) As Integer
'---------------------------------------------------------

On Error GoTo tableAnnuaire_Read_Error
tableAnnuaire_Read = 0


Select Case recAnnuaire.Method
     Case "Seek=       "
                        tableAnnuaire.Seek "=", recAnnuaire.Id
                        If tableAnnuaire.NoMatch Then
                            Error 9998
                        End If
     Case "MoveNext    "
                        tableAnnuaire.MoveNext
                        If tableAnnuaire.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableAnnuaire.MovePrevious
                        If tableAnnuaire.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst   "
                        tableAnnuaire.MoveFirst
                        If tableAnnuaire.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast    "
                        tableAnnuaire.MoveLast
                        If tableAnnuaire.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recAnnuaire.Method <> "AddNew      " Then
    Call tableAnnuaire_GetBuffer(recAnnuaire)
End If

Exit Function

'---------------------------------------------------------
tableAnnuaire_Read_Error:
'---------------------------------------------------------

    tableAnnuaire_Read = Err
    Resume tableAnnuaire_Read_End

tableAnnuaire_Read_End:

End Function

'---------------------------------------------------------
Public Function tableAnnuaire_Update(recAnnuaire As typeAnnuaire) As Integer
'---------------------------------------------------------

On Error GoTo tableAnnuaireUpdate_Error
tableAnnuaire_Update = 0

Select Case recAnnuaire.Method

    Case "AddNew      "
                        tableAnnuaire.AddNew
                        Call tableAnnuaire_PutBuffer(recAnnuaire)
                        tableAnnuaire.Update
    Case "Update      "
                        tableAnnuaire.Edit
                        Call tableAnnuaire_PutBuffer(recAnnuaire)
                        tableAnnuaire.Update
    Case "Delete      "
                        tableAnnuaire.Delete
    Case Else
                        Error 9999
End Select


Exit Function

tableAnnuaireUpdate_Error:
'---------------------------------------------------------
    tableAnnuaire_Update = Err
    Resume tableAnnuaireUpdate_End

tableAnnuaireUpdate_End:

End Function








'-----------------------------------------------------
Sub dbAnnuaire_Error(recAnnuaire As typeAnnuaire)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Annuaire: "

Select Case mId$(recAnnuaire.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recAnnuaire.Err: I = vbCritical
End Select

MsgBox Msg, I, "module : .bas  ( " _
                & Trim(recAnnuaire.obj) & " : " & Trim(recAnnuaire.Method) & " )"

End Sub

'-----------------------------------------------------
Function dbAnnuaire_Read(recAnnuaire As typeAnnuaire)
'-----------------------------------------------------

dbAnnuaire_Read = Null

recAnnuaire.Err = tableAnnuaire_Read(recAnnuaire)
If recAnnuaire.Err > 0 Then

    If recAnnuaire.Err < 9990 Or recAnnuaire.Err >= 9999 Then
        Call dbAnnuaire_Error(recAnnuaire)
        dbAnnuaire_Read = recAnnuaire.Err
    End If
End If

End Function

'---------------------------------------------------------
Public Sub arrAnnuaire_Load()
'---------------------------------------------------------
Dim iRead As Integer

tableAnnuaire_Open
arrAnnuaireNb = 0: arrAnnuaireNbMax = 0

recAnnuaire.Method = "MoveFirst"
recAnnuaire.Id = String$(4, Chr$(0))
recAnnuaire.obj = "Annuaire"
recAnnuaire.Err = 0

iRead = tableAnnuaire_Read(recAnnuaire)
Do While iRead = 0

    arrAnnuaire_AddItem recAnnuaire

    recAnnuaire.Method = "MoveNext    "
    iRead = tableAnnuaire_Read(recAnnuaire)
Loop
tableAnnuaire_Close

End Sub


'-----------------------------------------------------
Function dbAnnuaire_Update(recAnnuaire As typeAnnuaire)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
'$$$BeginTrans

dbAnnuaire_Update = Null


recAnnuaire.Err = tableAnnuaire_Update(recAnnuaire)

If recAnnuaire.Err <> 0 Then
    Call dbAnnuaire_Error(recAnnuaire)
    dbAnnuaire_Update = recAnnuaire.Err
'$$$    Rollback
    Exit Function
End If

'$$$CommitTrans


'=====================================================
End Function

'---------------------------------------------------------
Public Sub arrAnnuaire_AddItem(recAnnuaire As typeAnnuaire)
'---------------------------------------------------------
          
arrAnnuaireNb = arrAnnuaireNb + 1
    
If arrAnnuaireNb > arrAnnuaireNbMax Then
    arrAnnuaireNbMax = arrAnnuaireNbMax + 50
    ReDim Preserve arrAnnuaire(arrAnnuaireNbMax)
End If
recAnnuaire.Method = ""
arrAnnuaireIndex = arrAnnuaireNb
arrAnnuaire(arrAnnuaireIndex) = recAnnuaire
End Sub





Public Sub arrAnnuaire_Scan(Msg As String)
Dim I As Integer, I1 As Integer, xNom As String, xPrénoms As String, xTél1 As String
Dim V As Variant

arrAnnuaireIndex = -1
V = InStr(1, Msg, ":")
If IsNull(V) Then Exit Sub
I = CInt(V)
If I = 0 Then Exit Sub

xNom = Trim(mId$(Msg, 1, I - 1))
I1 = InStr(I + 1, Msg, ":")
xPrénoms = Trim(mId$(Msg, I + 1, I1 - I - 1))
I = InStr(I1 + 1, Msg, ":")
xTél1 = Trim(mId$(Msg, I1 + 1, I - I1 - 1))

For arrAnnuaireIndex = 1 To arrAnnuaireNb
    If Trim(arrAnnuaire(arrAnnuaireIndex).Nom) = xNom _
    And Trim(arrAnnuaire(arrAnnuaireIndex).Prénoms) = xPrénoms _
    And Trim(arrAnnuaire(arrAnnuaireIndex).Tél1) = xTél1 Then
        recAnnuaire = arrAnnuaire(arrAnnuaireIndex)
        Exit For
    End If
Next arrAnnuaireIndex

End Sub

Public Sub recAnnuaire_Init(recAnnuaire As typeAnnuaire)
recAnnuaire.obj = "Annuaire"
recAnnuaire.Method = ""
recAnnuaire.Err = 0
    
recAnnuaire.Id = ""
recAnnuaire.Civilité = "1"
recAnnuaire.Id = ""
recAnnuaire.Nom = ""
recAnnuaire.Prénoms = ""
recAnnuaire.Tél1 = ""
recAnnuaire.Tél2 = ""
recAnnuaire.Tél3 = ""
recAnnuaire.MicroSN = ""
recAnnuaire.MicroIP = ""
recAnnuaire.Service = ""
recAnnuaire.Bureau = ""

End Sub
