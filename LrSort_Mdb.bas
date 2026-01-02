Attribute VB_Name = "mdbLrSort"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableLrSort As Recordset
Dim tableLrSortOpen As Boolean

Type typeLrSort
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    RFBENF       As String * 16
    DTCENT1      As String * 6
    MTTOTAL      As Currency
    CDCPCO       As String * 1

End Type

Public recLrSort As typeLrSort


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableLrSort_Close()
'-----------------------------------------------------
If tableLrSortOpen Then
    tableLrSort.Close
    tableLrSortOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableLrSort_GetBuffer(recLrSort As typeLrSort)
'---------------------------------------------------------

recLrSort.RFBENF = tableLrSort("RFBENF")
recLrSort.DTCENT1 = tableLrSort("DTCENT1")
recLrSort.MTTOTAL = tableLrSort("MTTOTAL")
recLrSort.CDCPCO = tableLrSort("CDCPCO")

End Sub


'-----------------------------------------------------
Sub tableLrSort_Open()
'-----------------------------------------------------

If Not tableLrSortOpen Then
    Set tableLrSort = MDB.OpenRecordset("LrSort")
    tableLrSort.Index = "PrimaryKey"
    tableLrSortOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableLrSort_PutBuffer(recLrSort As typeLrSort)
'---------------------------------------------------------

tableLrSort("RFBENF") = recLrSort.RFBENF
tableLrSort("DTCENT1") = recLrSort.DTCENT1
tableLrSort("MTTOTAL") = recLrSort.MTTOTAL
tableLrSort("CDCPCO") = recLrSort.CDCPCO
End Sub


'---------------------------------------------------------
Public Function tableLrSort_Read(recLrSort As typeLrSort) As Integer
'---------------------------------------------------------

On Error GoTo tableLrSort_Read_Error
tableLrSort_Read = 0


Select Case recLrSort.Method
     Case "Seek=       "
                        tableLrSort.Seek "=", recLrSort.MTTOTAL
                        If tableLrSort.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<=      "
                        tableLrSort.Seek "<=", recLrSort.MTTOTAL
                        If tableLrSort.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>=      "
                        tableLrSort.Seek ">=", recLrSort.MTTOTAL
                        If tableLrSort.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>       "
                        tableLrSort.Seek ">", recLrSort.MTTOTAL
                        If tableLrSort.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext    "
                        tableLrSort.MoveNext
                        If tableLrSort.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableLrSort.MovePrevious
                        If tableLrSort.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst   "
                        tableLrSort.MoveFirst
                        If tableLrSort.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast    "
                        tableLrSort.MoveLast
                        If tableLrSort.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recLrSort.Method <> "AddNew      " Then
    Call tableLrSort_GetBuffer(recLrSort)
End If

Exit Function

'---------------------------------------------------------
tableLrSort_Read_Error:
'---------------------------------------------------------

    tableLrSort_Read = Err
    Resume tableLrSort_Read_End

tableLrSort_Read_End:

End Function

'---------------------------------------------------------
Public Function tableLrSort_Update(recLrSort As typeLrSort) As Integer
'---------------------------------------------------------

On Error GoTo tableLrSortUpdate_Error
tableLrSort_Update = 0

Select Case recLrSort.Method

    Case "AddNew      "
                        tableLrSort.AddNew
                        Call tableLrSort_PutBuffer(recLrSort)
                        tableLrSort.Update
    Case "Update      "
                        tableLrSort.Edit
                        Call tableLrSort_PutBuffer(recLrSort)
                        tableLrSort.Update
    Case "Delete      "
                        tableLrSort.Delete
    Case Else
                        Error 9999
End Select


Exit Function

tableLrSortUpdate_Error:
'---------------------------------------------------------
    tableLrSort_Update = Err
    Resume tableLrSortUpdate_End

tableLrSortUpdate_End:

End Function








'-----------------------------------------------------
Sub dbLrSort_Error(recLrSort As typeLrSort)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & Trim(recLrSort.RFBENF) & " : " & Trim(recLrSort.DTCENT1) & Chr$(13)

Select Case Mid$(recLrSort.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recLrSort.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbLrSort.bas :  ( " & Trim(recLrSort.obj) & " : " & Trim(recLrSort.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbLrSort_Read(recLrSort As typeLrSort)
'-----------------------------------------------------

dbLrSort_Read = Null

recLrSort.Err = tableLrSort_Read(recLrSort)
If recLrSort.Err > 0 Then

    If recLrSort.Err < 9990 Or recLrSort.Err >= 9999 Then
        Call dbLrSort_Error(recLrSort)
        dbLrSort_Read = recLrSort.Err
    End If
End If

End Function

'-----------------------------------------------------
Function dbLrSort_Update(recLrSort As typeLrSort)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
'$$$BeginTrans

dbLrSort_Update = Null


recLrSort.Err = tableLrSort_Update(recLrSort)

If recLrSort.Err <> 0 Then
    Call dbLrSort_Error(recLrSort)
    dbLrSort_Update = recLrSort.Err
'$$$    Rollback
    Exit Function
End If

'$$$CommitTrans


'=====================================================
End Function


