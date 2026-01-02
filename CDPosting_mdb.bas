Attribute VB_Name = "mdbCDPosting"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableCDPosting As Recordset
Dim tableCDPostingOpen As Boolean

Type typeCDPosting
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Dossier                 As Long
    Seq                     As Long
    TRAN_CODE               As String * 3
    VALUEDATE               As String * 8
    AMOUNT                  As Currency
    CCY                     As String * 3
    ACC_TYPE                As String * 2
    SK_CODE                 As String * 2
    POSTED_AS               As Long
    KEY97                   As Long
    CHARGE                  As Long

End Type

Public recCDPosting As typeCDPosting

'---------------------------------------------------------
'-----------------------------------------------------
Sub tableCDPosting_Close()
'-----------------------------------------------------
If tableCDPostingOpen Then
    tableCDPosting.Close
    tableCDPostingOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableCDPosting_GetBuffer(recCDPosting As typeCDPosting)
'---------------------------------------------------------
recCDPosting.Dossier = tableCDPosting("Dossier")
recCDPosting.Seq = tableCDPosting("Seq")
recCDPosting.TRAN_CODE = tableCDPosting("TRAN_CODE")
recCDPosting.VALUEDATE = tableCDPosting("VALUEDATE")
recCDPosting.AMOUNT = tableCDPosting("AMOUNT")
recCDPosting.CCY = tableCDPosting("CCY")
recCDPosting.ACC_TYPE = tableCDPosting("ACC_TYPE")
recCDPosting.SK_CODE = tableCDPosting("SK_CODE")
recCDPosting.POSTED_AS = tableCDPosting("POSTED_AS")
recCDPosting.KEY97 = tableCDPosting("KEY97")
recCDPosting.CHARGE = tableCDPosting("CHARGE")

End Sub


'-----------------------------------------------------
Sub tableCDPosting_Open()
'-----------------------------------------------------

If Not tableCDPostingOpen Then
    Set tableCDPosting = MDB.OpenRecordset("CDPosting")
    tableCDPosting.Index = "PrimaryKey"
    tableCDPostingOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableCDPosting_PutBuffer(recCDPosting As typeCDPosting)
'---------------------------------------------------------

tableCDPosting("Dossier") = recCDPosting.Dossier
tableCDPosting("Seq") = recCDPosting.Seq
tableCDPosting("TRAN_CODE") = recCDPosting.TRAN_CODE
tableCDPosting("VALUEDATE") = recCDPosting.VALUEDATE
tableCDPosting("AMOUNT") = recCDPosting.AMOUNT
tableCDPosting("CCY") = recCDPosting.CCY
tableCDPosting("ACC_TYPE") = recCDPosting.ACC_TYPE
tableCDPosting("SK_CODE") = recCDPosting.SK_CODE
tableCDPosting("POSTED_AS") = recCDPosting.POSTED_AS
tableCDPosting("KEY97") = recCDPosting.KEY97
tableCDPosting("CHARGE") = recCDPosting.CHARGE

End Sub


'---------------------------------------------------------
Public Function tableCDPosting_Read(recCDPosting As typeCDPosting) As Integer
'---------------------------------------------------------

On Error GoTo tableCDPosting_Read_Error
tableCDPosting_Read = 0


Select Case Trim(recCDPosting.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableCDPosting.Seek "=", recCDPosting.Dossier, recCDPosting.Seq
                        If tableCDPosting.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableCDPosting.Seek "<=", recCDPosting.Dossier, recCDPosting.Seq
                        If tableCDPosting.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableCDPosting.Seek ">=", recCDPosting.Dossier, recCDPosting.Seq
                        If tableCDPosting.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableCDPosting.Seek ">", recCDPosting.Dossier, recCDPosting.Seq
                        If tableCDPosting.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableCDPosting.MoveNext
                        If tableCDPosting.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableCDPosting.MovePrevious
                        If tableCDPosting.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableCDPosting.MoveFirst
                        If tableCDPosting.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableCDPosting.MoveLast
                        If tableCDPosting.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recCDPosting.Method <> "AddNew      " Then
    Call tableCDPosting_GetBuffer(recCDPosting)
End If

Exit Function

'---------------------------------------------------------
tableCDPosting_Read_Error:
'---------------------------------------------------------

    tableCDPosting_Read = Err
    Resume tableCDPosting_Read_End

tableCDPosting_Read_End:

End Function
'---------------------------------------------------------
Public Function tableCDPosting_Update(recCDPosting As typeCDPosting) As Integer
'---------------------------------------------------------

On Error GoTo tableCDPostingUpdate_Error
tableCDPosting_Update = 0

Select Case Trim(recCDPosting.Method)

    Case "AddNew"
                        tableCDPosting.AddNew
                        Call tableCDPosting_PutBuffer(recCDPosting)
                        tableCDPosting.Update
    Case "Update"
                        tableCDPosting.Edit
                        Call tableCDPosting_PutBuffer(recCDPosting)
                        tableCDPosting.Update
    Case "Delete"
                        tableCDPosting.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableCDPostingUpdate_Error:
'---------------------------------------------------------
    tableCDPosting_Update = Err
    Resume tableCDPostingUpdate_End

tableCDPostingUpdate_End:

End Function








'-----------------------------------------------------
Sub dbCDPosting_Error(recCDPosting As typeCDPosting)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recCDPosting.Dossier & ": " & Chr$(13)

Select Case mId$(recCDPosting.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recCDPosting.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbCDPosting.bas :  ( " & Trim(recCDPosting.obj) & " : " & Trim(recCDPosting.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbCDPosting_ReadE(recCDPosting As typeCDPosting)
'-----------------------------------------------------

dbCDPosting_ReadE = Null

recCDPosting.Err = tableCDPosting_Read(recCDPosting)
If recCDPosting.Err > 0 Then

'    If recCDPosting.Err < 9990 Or recCDPosting.Err >= 9999 Then
        Call dbCDPosting_Error(recCDPosting)
        dbCDPosting_ReadE = recCDPosting.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbCDPosting_Update(recCDPosting As typeCDPosting)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbCDPosting_Update = Null


recCDPosting.Err = tableCDPosting_Update(recCDPosting)

If recCDPosting.Err <> 0 Then
    Call dbCDPosting_Error(recCDPosting)
    dbCDPosting_Update = recCDPosting.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recCDPosting_Init(recCDPosting As typeCDPosting)
recCDPosting.Method = ""
recCDPosting.obj = "CDPosting"
recCDPosting.Err = ""
recCDPosting.Dossier = 0
recCDPosting.VALUEDATE = ""
recCDPosting.ACC_TYPE = ""
recCDPosting.AMOUNT = 0
recCDPosting.CCY = ""
recCDPosting.Seq = 0
recCDPosting.TRAN_CODE = ""
recCDPosting.POSTED_AS = 0
recCDPosting.KEY97 = 0
recCDPosting.SK_CODE = ""
recCDPosting.CHARGE = 0

End Sub



