Attribute VB_Name = "mdbCDComD"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableCDComD As Recordset
Dim tableCDComDOpen As Boolean

Type typeCDComD
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Dossier                 As Long
    Type                    As String * 2
    AmjD                    As String * 8
    AmjF                    As String * 8
    Devise                  As String * 3
    MvtEngagement           As Currency
    MvtUtilisé              As Currency
    MontantBase             As Currency
    CommissionTaux          As Double
    CommissionD             As Currency
    CommissionP             As Currency
    CommissionPAmj          As String * 8
    TIChargeKey             As Long
    CoursEur                As Double
   

End Type

Public recCDComD As typeCDComD

'---------------------------------------------------------
'-----------------------------------------------------
Sub tableCDComD_Close()
'-----------------------------------------------------
If tableCDComDOpen Then
    tableCDComD.Close
    tableCDComDOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableCDComD_GetBuffer(recCDComD As typeCDComD)
'---------------------------------------------------------
recCDComD.Dossier = tableCDComD("Dossier")
recCDComD.Type = tableCDComD("Type")
recCDComD.Devise = tableCDComD("Devise")
recCDComD.MontantBase = tableCDComD("MontantBase")
recCDComD.MvtEngagement = tableCDComD("MvtEngagement")
recCDComD.MvtUtilisé = tableCDComD("MvtUtilisé")
recCDComD.AmjD = tableCDComD("AmjD")
recCDComD.AmjF = tableCDComD("AmjF")
recCDComD.CommissionD = tableCDComD("CommissionD")
recCDComD.CommissionP = tableCDComD("CommissionP")
recCDComD.CommissionPAmj = tableCDComD("CommissionPAmj")
recCDComD.CommissionTaux = tableCDComD("CommissionTaux")
recCDComD.TIChargeKey = tableCDComD("TIChargeKey")
recCDComD.CoursEur = tableCDComD("CoursEur")

End Sub


'-----------------------------------------------------
Sub tableCDComD_Open()
'-----------------------------------------------------

If Not tableCDComDOpen Then
    Set tableCDComD = MDB.OpenRecordset("CDComD")
    tableCDComD.Index = "PrimaryKey"
    tableCDComDOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableCDComD_PutBuffer(recCDComD As typeCDComD)
'---------------------------------------------------------

tableCDComD("Dossier") = recCDComD.Dossier
tableCDComD("Type") = recCDComD.Type
tableCDComD("Devise") = recCDComD.Devise
tableCDComD("MontantBase") = recCDComD.MontantBase
tableCDComD("MvtEngagement") = recCDComD.MvtEngagement
tableCDComD("MvtUtilisé") = recCDComD.MvtUtilisé
tableCDComD("AmjD") = recCDComD.AmjD
tableCDComD("AmjF") = recCDComD.AmjF
tableCDComD("CommissionD") = recCDComD.CommissionD
tableCDComD("CommissionP") = recCDComD.CommissionP
tableCDComD("CommissionPAmj") = recCDComD.CommissionPAmj
tableCDComD("CommissionTaux") = recCDComD.CommissionTaux
tableCDComD("TIChargeKey") = recCDComD.TIChargeKey
tableCDComD("CoursEur") = recCDComD.CoursEur

End Sub


'---------------------------------------------------------
Public Function tableCDComD_Read(recCDComD As typeCDComD) As Integer
'---------------------------------------------------------

On Error GoTo tableCDComD_Read_Error
tableCDComD_Read = 0


Select Case Trim(recCDComD.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableCDComD.Seek "=", recCDComD.Dossier, recCDComD.Type, recCDComD.AmjD
                        If tableCDComD.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableCDComD.Seek "<=", recCDComD.Dossier, recCDComD.Type, recCDComD.AmjD
                        If tableCDComD.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableCDComD.Seek ">=", recCDComD.Dossier, recCDComD.Type, recCDComD.AmjD
                        If tableCDComD.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableCDComD.Seek ">", recCDComD.Dossier, recCDComD.Type, recCDComD.AmjD
                        If tableCDComD.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableCDComD.MoveNext
                        If tableCDComD.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableCDComD.MovePrevious
                        If tableCDComD.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableCDComD.MoveFirst
                        If tableCDComD.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableCDComD.MoveLast
                        If tableCDComD.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recCDComD.Method <> "AddNew      " Then
    Call tableCDComD_GetBuffer(recCDComD)
End If

Exit Function

'---------------------------------------------------------
tableCDComD_Read_Error:
'---------------------------------------------------------

    tableCDComD_Read = Err
    Resume tableCDComD_Read_End

tableCDComD_Read_End:

End Function
'---------------------------------------------------------
Public Function tableCDComD_Update(recCDComD As typeCDComD) As Integer
'---------------------------------------------------------

On Error GoTo tableCDComDUpdate_Error
tableCDComD_Update = 0

Select Case Trim(recCDComD.Method)

    Case "AddNew"
                        tableCDComD.AddNew
                        Call tableCDComD_PutBuffer(recCDComD)
                        tableCDComD.Update
    Case "Update"
                        tableCDComD.Edit
                        Call tableCDComD_PutBuffer(recCDComD)
                        tableCDComD.Update
    Case "Delete"
                        tableCDComD.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableCDComDUpdate_Error:
'---------------------------------------------------------
    tableCDComD_Update = Err
    Resume tableCDComDUpdate_End

tableCDComDUpdate_End:

End Function








'-----------------------------------------------------
Sub dbCDComD_Error(recCDComD As typeCDComD)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recCDComD.Dossier & ": " & Chr$(13)

Select Case mId$(recCDComD.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recCDComD.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbCDComD.bas :  ( " & Trim(recCDComD.obj) & " : " & Trim(recCDComD.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbCDComD_ReadE(recCDComD As typeCDComD)
'-----------------------------------------------------

dbCDComD_ReadE = Null

recCDComD.Err = tableCDComD_Read(recCDComD)
If recCDComD.Err > 0 Then

'    If recCDComD.Err < 9990 Or recCDComD.Err >= 9999 Then
        Call dbCDComD_Error(recCDComD)
        dbCDComD_ReadE = recCDComD.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbCDComD_Update(recCDComD As typeCDComD)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbCDComD_Update = Null


recCDComD.Err = tableCDComD_Update(recCDComD)

If recCDComD.Err <> 0 Then
    Call dbCDComD_Error(recCDComD)
    dbCDComD_Update = recCDComD.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recCDComD_Init(recCDComD As typeCDComD)
recCDComD.Method = ""
recCDComD.obj = "CD_Dossier"
recCDComD.Err = ""
recCDComD.Dossier = 0
recCDComD.Type = ""
recCDComD.Devise = ""
recCDComD.MontantBase = 0
recCDComD.MvtEngagement = 0
recCDComD.MvtUtilisé = 0
recCDComD.AmjD = "00000000"
recCDComD.AmjF = "00000000"
recCDComD.CommissionD = 0
recCDComD.CommissionP = 0
recCDComD.CommissionPAmj = "00000000"
recCDComD.CommissionTaux = 0
recCDComD.TIChargeKey = 0
recCDComD.CoursEur = 0

End Sub


