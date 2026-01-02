Attribute VB_Name = "mdbCDSTAT"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableCDSTAT As Recordset
Dim tableCDSTATOpen As Boolean
Public mCDSTAT_Id As Long

Type typeCDSTAT
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Sort                As String * 100
    Id                  As String * 16
    Text                As String

End Type

Public recCDSTAT As typeCDSTAT

'---------------------------------------------------------
'-----------------------------------------------------
Sub tableCDSTAT_Close()
'-----------------------------------------------------
If tableCDSTATOpen Then
    tableCDSTAT.Close
    tableCDSTATOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableCDSTAT_GetBuffer(recCDSTAT As typeCDSTAT)
'---------------------------------------------------------
recCDSTAT.Sort = tableCDSTAT("Sort")
recCDSTAT.Id = tableCDSTAT("Id")
recCDSTAT.Text = tableCDSTAT("Text")

End Sub


'-----------------------------------------------------
Sub tableCDSTAT_Open()
'-----------------------------------------------------

If Not tableCDSTATOpen Then
    Set tableCDSTAT = MDB.OpenRecordset("CDSTAT")
    tableCDSTAT.Index = "PrimaryKey"
    tableCDSTATOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableCDSTAT_PutBuffer(recCDSTAT As typeCDSTAT)
'---------------------------------------------------------
tableCDSTAT("Sort") = recCDSTAT.Sort

tableCDSTAT("Id") = recCDSTAT.Id
tableCDSTAT("Text") = recCDSTAT.Text
End Sub


'---------------------------------------------------------
Public Function tableCDSTAT_Read(recCDSTAT As typeCDSTAT) As Integer
'---------------------------------------------------------

On Error GoTo tableCDSTAT_Read_Error
tableCDSTAT_Read = 0


Select Case Trim(recCDSTAT.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableCDSTAT.Seek "=", recCDSTAT.Sort, recCDSTAT.Id
                        If tableCDSTAT.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableCDSTAT.Seek "<=", recCDSTAT.Sort, recCDSTAT.Id
                        If tableCDSTAT.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableCDSTAT.Seek ">=", recCDSTAT.Sort, recCDSTAT.Id
                        If tableCDSTAT.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableCDSTAT.Seek ">", recCDSTAT.Sort, recCDSTAT.Id
                        If tableCDSTAT.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableCDSTAT.MoveNext
                        If tableCDSTAT.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableCDSTAT.MovePrevious
                        If tableCDSTAT.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableCDSTAT.MoveFirst
                        If tableCDSTAT.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableCDSTAT.MoveLast
                        If tableCDSTAT.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recCDSTAT.Method <> "AddNew      " Then
    Call tableCDSTAT_GetBuffer(recCDSTAT)
End If

Exit Function

'---------------------------------------------------------
tableCDSTAT_Read_Error:
'---------------------------------------------------------

    tableCDSTAT_Read = Err
    Resume tableCDSTAT_Read_End

tableCDSTAT_Read_End:

End Function

'---------------------------------------------------------
Public Function tableCDSTAT_Update(recCDSTAT As typeCDSTAT) As Integer
'---------------------------------------------------------

On Error GoTo tableCDSTATUpdate_Error
tableCDSTAT_Update = 0

Select Case Trim(recCDSTAT.Method)

    Case "AddNew"
                        tableCDSTAT.AddNew
                        Call tableCDSTAT_PutBuffer(recCDSTAT)
                        tableCDSTAT.Update
    Case "Update"
                        tableCDSTAT.Edit
                        Call tableCDSTAT_PutBuffer(recCDSTAT)
                        tableCDSTAT.Update
    Case "Delete"
                        tableCDSTAT.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableCDSTATUpdate_Error:
'---------------------------------------------------------
    tableCDSTAT_Update = Err
    Resume tableCDSTATUpdate_End

tableCDSTATUpdate_End:

End Function

'-----------------------------------------------------
Sub dbCDSTAT_Error(recCDSTAT As typeCDSTAT)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recCDSTAT.Sort & recCDSTAT.Id & ": " & Chr$(13)

Select Case mId$(recCDSTAT.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recCDSTAT.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbCDSTAT.bas :  ( " & Trim(recCDSTAT.obj) & " : " & Trim(recCDSTAT.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbCDSTAT_ReadE(recCDSTAT As typeCDSTAT)
'-----------------------------------------------------

dbCDSTAT_ReadE = Null

recCDSTAT.Err = tableCDSTAT_Read(recCDSTAT)
If recCDSTAT.Err > 0 Then

'    If recCDSTAT.Err < 9990 Or recCDSTAT.Err >= 9999 Then
        Call dbCDSTAT_Error(recCDSTAT)
        dbCDSTAT_ReadE = recCDSTAT.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbCDSTAT_Update(recCDSTAT As typeCDSTAT)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbCDSTAT_Update = Null


recCDSTAT.Err = tableCDSTAT_Update(recCDSTAT)

If recCDSTAT.Err <> 0 Then
    Call dbCDSTAT_Error(recCDSTAT)
    dbCDSTAT_Update = recCDSTAT.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recCDSTAT_Init(recCDSTAT As typeCDSTAT)
recCDSTAT.Method = ""
recCDSTAT.obj = "CDSTAT"
recCDSTAT.Err = ""
recCDSTAT.Id = ""
recCDSTAT.Text = " "
End Sub





