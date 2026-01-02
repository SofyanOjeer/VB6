Attribute VB_Name = "mdbCDTIMaster"
'--------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableCDTIMaster As Recordset
Dim tableCDTIMasterOpen As Boolean

Type typeCDTIMaster
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    TIMasterKey             As Long
    Dossier                 As Long
   
End Type

Public recCDTIMaster As typeCDTIMaster

'---------------------------------------------------------
'-----------------------------------------------------
Sub tableCDTIMaster_Close()
'-----------------------------------------------------
If tableCDTIMasterOpen Then
    tableCDTIMaster.Close
    tableCDTIMasterOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableCDTIMaster_GetBuffer(recCDTIMaster As typeCDTIMaster)
'---------------------------------------------------------
recCDTIMaster.TIMasterKey = tableCDTIMaster("TIMasterKey")
recCDTIMaster.Dossier = tableCDTIMaster("Dossier")

End Sub


'-----------------------------------------------------
Sub tableCDTIMaster_Open()
'-----------------------------------------------------

If Not tableCDTIMasterOpen Then
    Set tableCDTIMaster = MDB.OpenRecordset("CDTIMaster")
    tableCDTIMaster.Index = "PrimaryKey"
    tableCDTIMasterOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableCDTIMaster_PutBuffer(recCDTIMaster As typeCDTIMaster)
'---------------------------------------------------------

tableCDTIMaster("TIMasterKey") = recCDTIMaster.TIMasterKey
tableCDTIMaster("Dossier") = recCDTIMaster.Dossier
End Sub


'---------------------------------------------------------
Public Function tableCDTIMaster_Read(recCDTIMaster As typeCDTIMaster) As Integer
'---------------------------------------------------------

On Error GoTo tableCDTIMaster_Read_Error
tableCDTIMaster_Read = 0


Select Case Trim(recCDTIMaster.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableCDTIMaster.Seek "=", recCDTIMaster.TIMasterKey
                        If tableCDTIMaster.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableCDTIMaster.Seek "<=", recCDTIMaster.TIMasterKey
                        If tableCDTIMaster.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableCDTIMaster.Seek ">=", recCDTIMaster.TIMasterKey
                        If tableCDTIMaster.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableCDTIMaster.Seek ">", recCDTIMaster.TIMasterKey
                        If tableCDTIMaster.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableCDTIMaster.MoveNext
                        If tableCDTIMaster.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableCDTIMaster.MovePrevious
                        If tableCDTIMaster.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableCDTIMaster.MoveFirst
                        If tableCDTIMaster.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableCDTIMaster.MoveLast
                        If tableCDTIMaster.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recCDTIMaster.Method <> "AddNew      " Then
    Call tableCDTIMaster_GetBuffer(recCDTIMaster)
End If

Exit Function

'---------------------------------------------------------
tableCDTIMaster_Read_Error:
'---------------------------------------------------------

    tableCDTIMaster_Read = Err
    Resume tableCDTIMaster_Read_End

tableCDTIMaster_Read_End:

End Function
'---------------------------------------------------------
Public Function tableCDTIMaster_Update(recCDTIMaster As typeCDTIMaster) As Integer
'---------------------------------------------------------

On Error GoTo tableCDTIMasterUpdate_Error
tableCDTIMaster_Update = 0

Select Case Trim(recCDTIMaster.Method)

    Case "AddNew"
                        tableCDTIMaster.AddNew
                        Call tableCDTIMaster_PutBuffer(recCDTIMaster)
                        tableCDTIMaster.Update
    Case "Update"
                        tableCDTIMaster.Edit
                        Call tableCDTIMaster_PutBuffer(recCDTIMaster)
                        tableCDTIMaster.Update
    Case "Delete"
                        tableCDTIMaster.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableCDTIMasterUpdate_Error:
'---------------------------------------------------------
    tableCDTIMaster_Update = Err
    Resume tableCDTIMasterUpdate_End

tableCDTIMasterUpdate_End:

End Function








'-----------------------------------------------------
Sub dbCDTIMaster_Error(recCDTIMaster As typeCDTIMaster)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recCDTIMaster.TIMasterKey & ": " & Chr$(13)

Select Case mId$(recCDTIMaster.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recCDTIMaster.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbCDTIMaster.bas :  ( " & Trim(recCDTIMaster.obj) & " : " & Trim(recCDTIMaster.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbCDTIMaster_ReadE(recCDTIMaster As typeCDTIMaster)
'-----------------------------------------------------

dbCDTIMaster_ReadE = Null

recCDTIMaster.Err = tableCDTIMaster_Read(recCDTIMaster)
If recCDTIMaster.Err > 0 Then

'    If recCDTIMaster.Err < 9990 Or recCDTIMaster.Err >= 9999 Then
        Call dbCDTIMaster_Error(recCDTIMaster)
        dbCDTIMaster_ReadE = recCDTIMaster.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbCDTIMaster_Update(recCDTIMaster As typeCDTIMaster)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbCDTIMaster_Update = Null


recCDTIMaster.Err = tableCDTIMaster_Update(recCDTIMaster)

If recCDTIMaster.Err <> 0 Then
    Call dbCDTIMaster_Error(recCDTIMaster)
    dbCDTIMaster_Update = recCDTIMaster.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recCDTIMaster_Init(recCDTIMaster As typeCDTIMaster)
recCDTIMaster.Method = ""
recCDTIMaster.obj = "CDTIMaster"
recCDTIMaster.Err = ""
recCDTIMaster.TIMasterKey = 0
recCDTIMaster.Dossier = 0

End Sub


