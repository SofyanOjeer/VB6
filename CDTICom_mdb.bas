Attribute VB_Name = "mdbCDTICom"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableCDTICom As Recordset
Dim tableCDTIComOpen As Boolean

Type typeCDTICom
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Dossier                 As Long
    TIChargeKey             As Long
    TIMasterKey             As Long
    Nature                  As String * 3
    CHCA_Devise             As String * 3
    CHCA                    As Currency
    CHBA_Devise             As String * 3
    CHBA                    As Currency
    CHAP_Devise             As String * 3
    CHAP                    As Currency
    CHAM_Devise             As String * 3
    CHAM                    As Currency
    
    CoursEur                As Double
    
    ComTaux1                As Double
    ComAMJD1                As String * 8
    ComAMJF1                As String * 8
    ComTaux2                As Double
    ComAMJD2                As String * 8
    ComAMJF2                As String * 8
    ComTaux3                As Double
    ComAMJD3                As String * 8
    ComAMJF3                As String * 8
    
    MontantPosting          As Currency
    AMJPosting              As String * 8
    TIPostingKey97          As Long
   

End Type

Public recCDTICom As typeCDTICom

'---------------------------------------------------------
'-----------------------------------------------------
Sub tableCDTICom_Close()
'-----------------------------------------------------
If tableCDTIComOpen Then
    tableCDTICom.Close
    tableCDTIComOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableCDTICom_GetBuffer(recCDTICom As typeCDTICom)
'---------------------------------------------------------
recCDTICom.Dossier = tableCDTICom("Dossier")
recCDTICom.TIChargeKey = tableCDTICom("TIChargeKey")
recCDTICom.TIMasterKey = tableCDTICom("TIMasterKey")
recCDTICom.Nature = tableCDTICom("Nature")
recCDTICom.CHCA_Devise = tableCDTICom("CHCA_Devise")
recCDTICom.CHCA = tableCDTICom("CHCA")
recCDTICom.CHBA_Devise = tableCDTICom("CHBA_Devise")
recCDTICom.CHBA = tableCDTICom("CHBA")
recCDTICom.CHAP_Devise = tableCDTICom("CHAP_Devise")
recCDTICom.CHAP = tableCDTICom("CHAP")
recCDTICom.CHAM_Devise = tableCDTICom("CHAM_Devise")
recCDTICom.CHAM = tableCDTICom("CHAM")

recCDTICom.CoursEur = tableCDTICom("CoursEur")

recCDTICom.ComTaux1 = tableCDTICom("ComTaux1")
recCDTICom.ComAMJD1 = tableCDTICom("ComAMJD1")
recCDTICom.ComAMJF1 = tableCDTICom("ComAMJF1")
recCDTICom.ComTaux2 = tableCDTICom("ComTaux2")
recCDTICom.ComAMJD2 = tableCDTICom("ComAMJD2")
recCDTICom.ComAMJF2 = tableCDTICom("ComAMJF2")
recCDTICom.ComTaux3 = tableCDTICom("ComTaux3")
recCDTICom.ComAMJD3 = tableCDTICom("ComAMJD3")
recCDTICom.ComAMJF3 = tableCDTICom("ComAMJF3")

recCDTICom.MontantPosting = tableCDTICom("MontantPosting")
recCDTICom.AMJPosting = tableCDTICom("AMJPosting")
recCDTICom.TIPostingKey97 = tableCDTICom("TIPostingKey97")

End Sub


'-----------------------------------------------------
Sub tableCDTICom_Open()
'-----------------------------------------------------

If Not tableCDTIComOpen Then
    Set tableCDTICom = MDB.OpenRecordset("CDTICom")
    tableCDTICom.Index = "PrimaryKey"
    tableCDTIComOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableCDTICom_PutBuffer(recCDTICom As typeCDTICom)
'---------------------------------------------------------

tableCDTICom("Dossier") = recCDTICom.Dossier
tableCDTICom("TIChargeKey") = recCDTICom.TIChargeKey
tableCDTICom("TIMasterKey") = recCDTICom.TIMasterKey
tableCDTICom("Nature") = recCDTICom.Nature
tableCDTICom("CHCA_Devise") = recCDTICom.CHCA_Devise
tableCDTICom("CHCA") = recCDTICom.CHCA
tableCDTICom("CHBA_Devise") = recCDTICom.CHBA_Devise
tableCDTICom("CHBA") = recCDTICom.CHBA
tableCDTICom("CHAP_Devise") = recCDTICom.CHAP_Devise
tableCDTICom("CHAP") = recCDTICom.CHAP
tableCDTICom("CHAM_Devise") = recCDTICom.CHAM_Devise
tableCDTICom("CHAM") = recCDTICom.CHAM

tableCDTICom("CoursEur") = recCDTICom.CoursEur

tableCDTICom("ComTaux1") = recCDTICom.ComTaux1
tableCDTICom("ComAMJD1") = recCDTICom.ComAMJD1
tableCDTICom("ComAMJF1") = recCDTICom.ComAMJF1
tableCDTICom("ComTaux2") = recCDTICom.ComTaux2
tableCDTICom("ComAMJD2") = recCDTICom.ComAMJD2
tableCDTICom("ComAMJF2") = recCDTICom.ComAMJF2
tableCDTICom("ComTaux3") = recCDTICom.ComTaux3
tableCDTICom("ComAMJD3") = recCDTICom.ComAMJD3
tableCDTICom("ComAMJF3") = recCDTICom.ComAMJF3

tableCDTICom("MontantPosting") = recCDTICom.MontantPosting
tableCDTICom("AMJPosting") = recCDTICom.AMJPosting
tableCDTICom("TIPostingKey97") = recCDTICom.TIPostingKey97
End Sub


'---------------------------------------------------------
Public Function tableCDTICom_Read(recCDTICom As typeCDTICom) As Integer
'---------------------------------------------------------

On Error GoTo tableCDTICom_Read_Error
tableCDTICom_Read = 0


Select Case Trim(recCDTICom.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableCDTICom.Seek "=", recCDTICom.Dossier, recCDTICom.TIChargeKey
                        If tableCDTICom.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableCDTICom.Seek "<=", recCDTICom.Dossier, recCDTICom.TIChargeKey
                        If tableCDTICom.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableCDTICom.Seek ">=", recCDTICom.Dossier, recCDTICom.TIChargeKey
                        If tableCDTICom.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableCDTICom.Seek ">", recCDTICom.Dossier, recCDTICom.TIChargeKey
                        If tableCDTICom.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableCDTICom.MoveNext
                        If tableCDTICom.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableCDTICom.MovePrevious
                        If tableCDTICom.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableCDTICom.MoveFirst
                        If tableCDTICom.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableCDTICom.MoveLast
                        If tableCDTICom.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recCDTICom.Method <> "AddNew      " Then
    Call tableCDTICom_GetBuffer(recCDTICom)
End If

Exit Function

'---------------------------------------------------------
tableCDTICom_Read_Error:
'---------------------------------------------------------

    tableCDTICom_Read = Err
    Resume tableCDTICom_Read_End

tableCDTICom_Read_End:

End Function
'---------------------------------------------------------
Public Function tableCDTICom_Update(recCDTICom As typeCDTICom) As Integer
'---------------------------------------------------------

On Error GoTo tableCDTIComUpdate_Error
tableCDTICom_Update = 0

Select Case Trim(recCDTICom.Method)

    Case "AddNew"
                        tableCDTICom.AddNew
                        Call tableCDTICom_PutBuffer(recCDTICom)
                        tableCDTICom.Update
    Case "Update"
                        tableCDTICom.Edit
                        Call tableCDTICom_PutBuffer(recCDTICom)
                        tableCDTICom.Update
    Case "Delete"
                        tableCDTICom.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableCDTIComUpdate_Error:
'---------------------------------------------------------
    tableCDTICom_Update = Err
    Resume tableCDTIComUpdate_End

tableCDTIComUpdate_End:

End Function








'-----------------------------------------------------
Sub dbCDTICom_Error(recCDTICom As typeCDTICom)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recCDTICom.TIChargeKey & ": " & Chr$(13)

Select Case mId$(recCDTICom.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recCDTICom.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbCDTICom.bas :  ( " & Trim(recCDTICom.obj) & " : " & Trim(recCDTICom.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbCDTICom_ReadE(recCDTICom As typeCDTICom)
'-----------------------------------------------------

dbCDTICom_ReadE = Null

recCDTICom.Err = tableCDTICom_Read(recCDTICom)
If recCDTICom.Err > 0 Then

'    If recCDTICom.Err < 9990 Or recCDTICom.Err >= 9999 Then
        Call dbCDTICom_Error(recCDTICom)
        dbCDTICom_ReadE = recCDTICom.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbCDTICom_Update(recCDTICom As typeCDTICom)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbCDTICom_Update = Null


recCDTICom.Err = tableCDTICom_Update(recCDTICom)

If recCDTICom.Err <> 0 Then
    Call dbCDTICom_Error(recCDTICom)
    dbCDTICom_Update = recCDTICom.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recCDTICom_Init(recCDTICom As typeCDTICom)
recCDTICom.Method = ""
recCDTICom.obj = "CDTICom"
recCDTICom.Err = ""
recCDTICom.Dossier = 0
recCDTICom.TIChargeKey = 0
recCDTICom.TIMasterKey = 0
recCDTICom.Nature = ""
recCDTICom.CHCA_Devise = ""
recCDTICom.CHCA = 0
recCDTICom.CHBA_Devise = ""
recCDTICom.CHBA = 0
recCDTICom.CHAP_Devise = ""
recCDTICom.CHAP = 0
recCDTICom.CHAM_Devise = ""
recCDTICom.CHAM = 0
recCDTICom.CoursEur = 0
recCDTICom.ComTaux1 = 0
recCDTICom.ComAMJD1 = "00000000"
recCDTICom.ComAMJF1 = "00000000"
recCDTICom.ComTaux2 = 0
recCDTICom.ComAMJD2 = "00000000"
recCDTICom.ComAMJF2 = "00000000"
recCDTICom.ComTaux3 = 0
recCDTICom.ComAMJD3 = "00000000"
recCDTICom.ComAMJF3 = "00000000"
recCDTICom.MontantPosting = 0
recCDTICom.AMJPosting = "00000000"
recCDTICom.TIPostingKey97 = 0

End Sub


