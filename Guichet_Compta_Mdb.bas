Attribute VB_Name = "mdbGuichet_Compta"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableGuichet_Compta As Recordset
Dim tableGuichet_ComptaOpen As Boolean

Type typeGuichet_Compta
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Société             As String * 3
    Agence              As String * 3
    SaisieUsr           As String * 10
    Devise              As String * 3
    CodeOpération       As String * 10
    CptMvtPièce         As Long
    CptMvtLigne         As Long
    Service             As String * 4
    Compte              As String * 11
    Montant             As Currency
    Sens                As String * 1
    AmjOpération        As String * 8
    AmjValeur           As String * 8
    Libellé             As String * 50
    SaisieAmj           As String * 8
    chkCompte           As String * 1
    chkSolde            As String * 1
    chkChèque           As String * 1
    chkAmjValeur        As String * 1
    Référence           As String * 10
    Devise2             As String * 3
    Montant2            As Currency
    Devise3             As String * 3
    Montant3            As Currency
    ValidationAMJ      As String * 8
    ValidationHMS      As String * 6
    ValidationUsr      As String * 10
    ComptaAMJ           As String * 8
    ComptaHMS           As String * 6
    ComptaUsr           As String * 10
End Type

Public recGuichet_Compta As typeGuichet_Compta


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableGuichet_Compta_Close()
'-----------------------------------------------------
If tableGuichet_ComptaOpen Then
    tableGuichet_Compta.Close
    tableGuichet_ComptaOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableGuichet_Compta_GetBuffer(recGuichet_Compta As typeGuichet_Compta)
'---------------------------------------------------------
recGuichet_Compta.Société = tableGuichet_Compta("Société")
recGuichet_Compta.Agence = tableGuichet_Compta("Agence")
recGuichet_Compta.SaisieUsr = tableGuichet_Compta("SaisieUsr")
recGuichet_Compta.Devise = tableGuichet_Compta("Devise")
recGuichet_Compta.CodeOpération = tableGuichet_Compta("CodeOpération")

recGuichet_Compta.CptMvtPièce = tableGuichet_Compta("CptMvtPièce")
recGuichet_Compta.CptMvtLigne = tableGuichet_Compta("CptMvtLigne")
recGuichet_Compta.Service = tableGuichet_Compta("Service")
recGuichet_Compta.Compte = tableGuichet_Compta("Compte")
recGuichet_Compta.Montant = tableGuichet_Compta("Montant")
recGuichet_Compta.Sens = tableGuichet_Compta("Sens")
recGuichet_Compta.AmjOpération = tableGuichet_Compta("AmjOpération")
recGuichet_Compta.AmjValeur = tableGuichet_Compta("AmjValeur")
recGuichet_Compta.Libellé = tableGuichet_Compta("Libellé")
recGuichet_Compta.SaisieAmj = tableGuichet_Compta("SaisieAmj")
recGuichet_Compta.chkCompte = tableGuichet_Compta("chkCompte")
recGuichet_Compta.chkSolde = tableGuichet_Compta("chkSolde")
recGuichet_Compta.chkChèque = tableGuichet_Compta("chkChèque")
recGuichet_Compta.chkAmjValeur = tableGuichet_Compta("chkAmjvaleur")

recGuichet_Compta.Référence = tableGuichet_Compta("Référence")
recGuichet_Compta.Devise2 = tableGuichet_Compta("Devise2")
recGuichet_Compta.Montant2 = tableGuichet_Compta("Montant2")
recGuichet_Compta.Devise3 = tableGuichet_Compta("Devise3")
recGuichet_Compta.Montant3 = tableGuichet_Compta("Montant3")

recGuichet_Compta.ValidationAMJ = tableGuichet_Compta("ValidationAMJ")
recGuichet_Compta.ValidationHMS = tableGuichet_Compta("ValidationHMS")
recGuichet_Compta.ValidationUsr = tableGuichet_Compta("ValidationUsr")

recGuichet_Compta.ComptaAMJ = tableGuichet_Compta("ComptaAMJ")
recGuichet_Compta.ComptaHMS = tableGuichet_Compta("ComptaHMS")
recGuichet_Compta.ComptaUsr = tableGuichet_Compta("ComptaUsr")
End Sub


'-----------------------------------------------------
Sub tableGuichet_Compta_Open()
'-----------------------------------------------------

If Not tableGuichet_ComptaOpen Then
    Set tableGuichet_Compta = MDB.OpenRecordset("Guichet_Compta")
    tableGuichet_Compta.Index = "PrimaryKey"
    tableGuichet_ComptaOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableGuichet_Compta_PutBuffer(recGuichet_Compta As typeGuichet_Compta)
'---------------------------------------------------------

tableGuichet_Compta("Société") = recGuichet_Compta.Société
tableGuichet_Compta("Agence") = recGuichet_Compta.Agence
tableGuichet_Compta("SaisieUsr") = recGuichet_Compta.SaisieUsr
tableGuichet_Compta("Devise") = recGuichet_Compta.Devise
tableGuichet_Compta("CodeOpération") = recGuichet_Compta.CodeOpération
tableGuichet_Compta("CptMvtPièce") = recGuichet_Compta.CptMvtPièce
tableGuichet_Compta("CptMvtLigne") = recGuichet_Compta.CptMvtLigne
tableGuichet_Compta("Montant") = recGuichet_Compta.Montant
tableGuichet_Compta("Service") = recGuichet_Compta.Service
tableGuichet_Compta("Compte") = recGuichet_Compta.Compte

tableGuichet_Compta("Sens") = recGuichet_Compta.Sens
tableGuichet_Compta("AmjOpération") = recGuichet_Compta.AmjOpération
tableGuichet_Compta("AmjValeur") = recGuichet_Compta.AmjValeur
tableGuichet_Compta("Libellé") = recGuichet_Compta.Libellé

tableGuichet_Compta("SaisieAmj") = recGuichet_Compta.SaisieAmj
tableGuichet_Compta("chkCompte") = recGuichet_Compta.chkCompte
tableGuichet_Compta("chkSolde") = recGuichet_Compta.chkSolde
tableGuichet_Compta("chkChèque") = recGuichet_Compta.chkChèque
tableGuichet_Compta("chkAmjvaleur") = recGuichet_Compta.chkAmjValeur

tableGuichet_Compta("Référence") = recGuichet_Compta.Référence
tableGuichet_Compta("Devise2") = recGuichet_Compta.Devise2
tableGuichet_Compta("Montant2") = recGuichet_Compta.Montant2
tableGuichet_Compta("Devise3") = recGuichet_Compta.Devise3
tableGuichet_Compta("Montant3") = recGuichet_Compta.Montant3

tableGuichet_Compta("ValidationAMJ") = recGuichet_Compta.ValidationAMJ
tableGuichet_Compta("ValidationHMS") = recGuichet_Compta.ValidationHMS
tableGuichet_Compta("ValidationUsr") = recGuichet_Compta.ValidationUsr

tableGuichet_Compta("ComptaAMJ") = recGuichet_Compta.ComptaAMJ
tableGuichet_Compta("ComptaHMS") = recGuichet_Compta.ComptaHMS
tableGuichet_Compta("ComptaUsr") = recGuichet_Compta.ComptaUsr
End Sub


'---------------------------------------------------------
Public Function tableGuichet_Compta_Read(recGuichet_Compta As typeGuichet_Compta) As Integer
'---------------------------------------------------------

On Error GoTo tableGuichet_Compta_Read_Error
tableGuichet_Compta_Read = 0


Select Case Trim(recGuichet_Compta.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableGuichet_Compta.Seek "=", recGuichet_Compta.Société, recGuichet_Compta.Agence, _
                                                    recGuichet_Compta.SaisieUsr, recGuichet_Compta.Devise, recGuichet_Compta.CptMvtPièce, recGuichet_Compta.CptMvtLigne
                        If tableGuichet_Compta.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableGuichet_Compta.Seek "<=", recGuichet_Compta.Société, recGuichet_Compta.Agence, _
                                                    recGuichet_Compta.SaisieUsr, recGuichet_Compta.Devise, recGuichet_Compta.CptMvtPièce, recGuichet_Compta.CptMvtLigne
                        If tableGuichet_Compta.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableGuichet_Compta.Seek ">=", recGuichet_Compta.Société, recGuichet_Compta.Agence, _
                                                    recGuichet_Compta.SaisieUsr, recGuichet_Compta.Devise, recGuichet_Compta.CptMvtPièce, recGuichet_Compta.CptMvtLigne
                        If tableGuichet_Compta.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableGuichet_Compta.Seek ">", recGuichet_Compta.Société, recGuichet_Compta.Agence, _
                                                    recGuichet_Compta.SaisieUsr, recGuichet_Compta.Devise, recGuichet_Compta.CptMvtPièce, recGuichet_Compta.CptMvtLigne
                        If tableGuichet_Compta.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableGuichet_Compta.MoveNext
                        If tableGuichet_Compta.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableGuichet_Compta.MovePrevious
                        If tableGuichet_Compta.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableGuichet_Compta.MoveFirst
                        If tableGuichet_Compta.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableGuichet_Compta.MoveLast
                        If tableGuichet_Compta.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recGuichet_Compta.Method <> "AddNew      " Then
    Call tableGuichet_Compta_GetBuffer(recGuichet_Compta)
End If

Exit Function

'---------------------------------------------------------
tableGuichet_Compta_Read_Error:
'---------------------------------------------------------

    tableGuichet_Compta_Read = Err
    Resume tableGuichet_Compta_Read_End

tableGuichet_Compta_Read_End:

End Function

'---------------------------------------------------------
Public Function tableGuichet_Compta_Update(recGuichet_Compta As typeGuichet_Compta) As Integer
'---------------------------------------------------------

On Error GoTo tableGuichet_ComptaUpdate_Error
tableGuichet_Compta_Update = 0

Select Case Trim(recGuichet_Compta.Method)

    Case "AddNew"
                        tableGuichet_Compta.AddNew
                        Call tableGuichet_Compta_PutBuffer(recGuichet_Compta)
                        tableGuichet_Compta.Update
    Case "Update"
                        tableGuichet_Compta.Edit
                        Call tableGuichet_Compta_PutBuffer(recGuichet_Compta)
                        tableGuichet_Compta.Update
    Case "Delete"
                        tableGuichet_Compta.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableGuichet_ComptaUpdate_Error:
'---------------------------------------------------------
    tableGuichet_Compta_Update = Err
    Resume tableGuichet_ComptaUpdate_End

tableGuichet_ComptaUpdate_End:

End Function








'-----------------------------------------------------
Sub dbGuichet_Compta_Error(recGuichet_Compta As typeGuichet_Compta)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & Trim(recGuichet_Compta.SaisieUsr) & " : " & Trim(recGuichet_Compta.CptMvtPièce) & " : " & Trim(recGuichet_Compta.CptMvtLigne) & Chr$(13)

Select Case mId$(recGuichet_Compta.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recGuichet_Compta.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbGuichet_Compta.bas :  ( " & Trim(recGuichet_Compta.obj) & " : " & Trim(recGuichet_Compta.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbGuichet_Compta_ReadE(recGuichet_Compta As typeGuichet_Compta)
'-----------------------------------------------------

dbGuichet_Compta_ReadE = Null

recGuichet_Compta.Err = tableGuichet_Compta_Read(recGuichet_Compta)
If recGuichet_Compta.Err > 0 Then

'    If recGuichet_Compta.Err < 9990 Or recGuichet_Compta.Err >= 9999 Then
        Call dbGuichet_Compta_Error(recGuichet_Compta)
        dbGuichet_Compta_ReadE = recGuichet_Compta.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbGuichet_Compta_Update(recGuichet_Compta As typeGuichet_Compta)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbGuichet_Compta_Update = Null


recGuichet_Compta.Err = tableGuichet_Compta_Update(recGuichet_Compta)

If recGuichet_Compta.Err <> 0 Then
    Call dbGuichet_Compta_Error(recGuichet_Compta)
    dbGuichet_Compta_Update = recGuichet_Compta.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recGuichet_Compta_Init(recGuichet_Compta As typeGuichet_Compta)
recGuichet_Compta.Method = ""
recGuichet_Compta.obj = "Guichet_Compta"
recGuichet_Compta.Err = ""
recGuichet_Compta.SaisieUsr = ""
recGuichet_Compta.SaisieAmj = ""
recGuichet_Compta.Devise = ""
recGuichet_Compta.CodeOpération = ""
recGuichet_Compta.CptMvtPièce = 0
recGuichet_Compta.CptMvtLigne = 0
recGuichet_Compta.Compte = "00000000000"
recGuichet_Compta.Montant = 0
recGuichet_Compta.Sens = " "
recGuichet_Compta.AmjValeur = "00000000"
recGuichet_Compta.Libellé = ""
recGuichet_Compta.ValidationAMJ = ""
recGuichet_Compta.ValidationHMS = ""
recGuichet_Compta.ValidationUsr = ""
recGuichet_Compta.ComptaAMJ = ""
recGuichet_Compta.ComptaHMS = ""
recGuichet_Compta.ComptaUsr = ""

End Sub
