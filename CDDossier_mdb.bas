Attribute VB_Name = "mdbCDDossier"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableCDDossier As Recordset
Dim tableCDDossierOpen As Boolean

Type typeCDDossier
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Dossier                 As Long
    Compte                  As String * 11
    Devise                  As String * 3
    Montant                 As Currency
    AMJOuverture            As String * 8
    AMJValidité             As String * 8
    NbJours                 As Long
    MontantEngagement       As Currency
    MontantUtilisé          As Currency
    CommissionD             As Currency
    CommissionP             As Currency
    Confirmé                As String * 1
    AMJSituation            As String * 8
    TIMasterKey             As Long
    TIMt226                 As Currency
    TIMt651                 As Currency
    S36Engagement           As Currency
    S36Utilisé              As Currency
    S36RC                   As Currency
    S36RE                   As Currency
    S36RI                   As Currency
    S36RA                   As Currency


End Type

Public recCDDossier As typeCDDossier

'---------------------------------------------------------
'-----------------------------------------------------
Sub tableCDDossier_Close()
'-----------------------------------------------------
If tableCDDossierOpen Then
    tableCDDossier.Close
    tableCDDossierOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableCDDossier_GetBuffer(recCDDossier As typeCDDossier)
'---------------------------------------------------------
recCDDossier.Dossier = tableCDDossier("Dossier")
recCDDossier.Compte = tableCDDossier("Compte")
recCDDossier.Devise = tableCDDossier("Devise")
recCDDossier.Montant = tableCDDossier("Montant")
recCDDossier.MontantEngagement = tableCDDossier("MontantEngagement")
recCDDossier.MontantUtilisé = tableCDDossier("MontantUtilisé")
recCDDossier.AMJOuverture = tableCDDossier("AMJOuverture")
recCDDossier.AMJValidité = tableCDDossier("AMJValidité")
recCDDossier.NbJours = tableCDDossier("NbJours")
recCDDossier.CommissionD = tableCDDossier("CommissionD")
recCDDossier.CommissionP = tableCDDossier("CommissionP")
recCDDossier.Confirmé = tableCDDossier("Confirmé")
recCDDossier.AMJSituation = tableCDDossier("AMJSituation")
recCDDossier.TIMasterKey = tableCDDossier("TIMasterKey")
recCDDossier.TIMt226 = tableCDDossier("TIMt226")
recCDDossier.TIMt651 = tableCDDossier("TIMt651")
recCDDossier.S36Engagement = tableCDDossier("S36Engagement")
recCDDossier.S36Utilisé = tableCDDossier("S36Utilisé")
recCDDossier.S36RC = tableCDDossier("S36RC")
recCDDossier.S36RE = tableCDDossier("S36RE")
recCDDossier.S36RI = tableCDDossier("S36RI")
recCDDossier.S36RA = tableCDDossier("S36RA")

End Sub


'-----------------------------------------------------
Sub tableCDDossier_Open()
'-----------------------------------------------------

If Not tableCDDossierOpen Then
    Set tableCDDossier = MDB.OpenRecordset("CDDossier")
    tableCDDossier.Index = "PrimaryKey"
    tableCDDossierOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableCDDossier_PutBuffer(recCDDossier As typeCDDossier)
'---------------------------------------------------------

tableCDDossier("Dossier") = recCDDossier.Dossier
tableCDDossier("Compte") = recCDDossier.Compte
tableCDDossier("Devise") = recCDDossier.Devise
tableCDDossier("Montant") = recCDDossier.Montant
tableCDDossier("MontantEngagement") = recCDDossier.MontantEngagement
tableCDDossier("MontantUtilisé") = recCDDossier.MontantUtilisé
tableCDDossier("AMJOuverture") = recCDDossier.AMJOuverture
tableCDDossier("AMJValidité") = recCDDossier.AMJValidité
tableCDDossier("NbJours") = recCDDossier.NbJours
tableCDDossier("CommissionD") = recCDDossier.CommissionD
tableCDDossier("CommissionP") = recCDDossier.CommissionP
tableCDDossier("Confirmé") = recCDDossier.Confirmé
tableCDDossier("AMJSituation") = recCDDossier.AMJSituation
tableCDDossier("TIMasterKey") = recCDDossier.TIMasterKey
tableCDDossier("TIMt226") = recCDDossier.TIMt226
tableCDDossier("TIMt651") = recCDDossier.TIMt651
tableCDDossier("S36Engagement") = recCDDossier.S36Engagement
tableCDDossier("S36Utilisé") = recCDDossier.S36Utilisé
tableCDDossier("S36RC") = recCDDossier.S36RC
tableCDDossier("S36RE") = recCDDossier.S36RE
tableCDDossier("S36RI") = recCDDossier.S36RI
tableCDDossier("S36RA") = recCDDossier.S36RA

End Sub


'---------------------------------------------------------
Public Function tableCDDossier_Read(recCDDossier As typeCDDossier) As Integer
'---------------------------------------------------------

On Error GoTo tableCDDossier_Read_Error
tableCDDossier_Read = 0


Select Case Trim(recCDDossier.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableCDDossier.Seek "=", recCDDossier.Dossier
                        If tableCDDossier.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableCDDossier.Seek "<=", recCDDossier.Dossier
                        If tableCDDossier.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableCDDossier.Seek ">=", recCDDossier.Dossier
                        If tableCDDossier.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableCDDossier.Seek ">", recCDDossier.Dossier
                        If tableCDDossier.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableCDDossier.MoveNext
                        If tableCDDossier.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableCDDossier.MovePrevious
                        If tableCDDossier.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableCDDossier.MoveFirst
                        If tableCDDossier.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableCDDossier.MoveLast
                        If tableCDDossier.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recCDDossier.Method <> "AddNew      " Then
    Call tableCDDossier_GetBuffer(recCDDossier)
End If

Exit Function

'---------------------------------------------------------
tableCDDossier_Read_Error:
'---------------------------------------------------------

    tableCDDossier_Read = Err
    Resume tableCDDossier_Read_End

tableCDDossier_Read_End:

End Function
'---------------------------------------------------------
Public Function tableCDDossier_Update(recCDDossier As typeCDDossier) As Integer
'---------------------------------------------------------

On Error GoTo tableCDDossierUpdate_Error
tableCDDossier_Update = 0

Select Case Trim(recCDDossier.Method)

    Case "AddNew"
                        tableCDDossier.AddNew
                        Call tableCDDossier_PutBuffer(recCDDossier)
                        tableCDDossier.Update
    Case "Update"
                        tableCDDossier.Edit
                        Call tableCDDossier_PutBuffer(recCDDossier)
                        tableCDDossier.Update
    Case "Delete"
                        tableCDDossier.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableCDDossierUpdate_Error:
'---------------------------------------------------------
    tableCDDossier_Update = Err
    Resume tableCDDossierUpdate_End

tableCDDossierUpdate_End:

End Function








'-----------------------------------------------------
Sub dbCDDossier_Error(recCDDossier As typeCDDossier)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recCDDossier.Dossier & ": " & Chr$(13)

Select Case mId$(recCDDossier.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recCDDossier.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbCDDossier.bas :  ( " & Trim(recCDDossier.obj) & " : " & Trim(recCDDossier.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbCDDossier_ReadE(recCDDossier As typeCDDossier)
'-----------------------------------------------------

dbCDDossier_ReadE = Null

recCDDossier.Err = tableCDDossier_Read(recCDDossier)
If recCDDossier.Err > 0 Then

'    If recCDDossier.Err < 9990 Or recCDDossier.Err >= 9999 Then
        Call dbCDDossier_Error(recCDDossier)
        dbCDDossier_ReadE = recCDDossier.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbCDDossier_Update(recCDDossier As typeCDDossier)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbCDDossier_Update = Null


recCDDossier.Err = tableCDDossier_Update(recCDDossier)

If recCDDossier.Err <> 0 Then
    Call dbCDDossier_Error(recCDDossier)
    dbCDDossier_Update = recCDDossier.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recCDDossier_Init(recCDDossier As typeCDDossier)
recCDDossier.Method = ""
recCDDossier.obj = "CD_Dossier"
recCDDossier.Err = ""
recCDDossier.Dossier = 0
recCDDossier.Compte = ""
recCDDossier.Devise = ""
recCDDossier.Montant = 0
recCDDossier.MontantEngagement = 0
recCDDossier.MontantUtilisé = 0
recCDDossier.AMJOuverture = "00000000"
recCDDossier.AMJValidité = "00000000"
recCDDossier.NbJours = 0
recCDDossier.CommissionD = 0
recCDDossier.CommissionP = 0
recCDDossier.Confirmé = ""
recCDDossier.AMJSituation = "00000000"
recCDDossier.TIMasterKey = 0
recCDDossier.TIMt226 = 0
recCDDossier.TIMt651 = 0

recCDDossier.S36Engagement = 0
recCDDossier.S36Utilisé = 0
recCDDossier.S36RC = 0
recCDDossier.S36RE = 0
recCDDossier.S36RI = 0
recCDDossier.S36RA = 0
End Sub


