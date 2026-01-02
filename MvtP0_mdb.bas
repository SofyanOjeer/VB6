Attribute VB_Name = "mdbMvtP0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableMvtP0 As Recordset
Dim tableMvtP0Open As Boolean
Public mMvtP0_Id As Long

Type typeMvtP0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    ID                  As String * 40
    Text                As String

End Type

Public recMvtp0 As typeMvtP0


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableMvtP0_Close()
'-----------------------------------------------------
If tableMvtP0Open Then
    tableMvtP0.Close
    tableMvtP0Open = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableMvtP0_GetBuffer(recMvtp0 As typeMvtP0)
'---------------------------------------------------------
recMvtp0.ID = tableMvtP0("Id")
recMvtp0.Text = tableMvtP0("Text")

End Sub


'-----------------------------------------------------
Sub tableMvtP0_Open()
'-----------------------------------------------------

If Not tableMvtP0Open Then
    Set tableMvtP0 = MDB.OpenRecordset("MvtP0")
    tableMvtP0.Index = "PrimaryKey"
    tableMvtP0Open = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableMvtP0_PutBuffer(recMvtp0 As typeMvtP0)
'---------------------------------------------------------

tableMvtP0("Id") = recMvtp0.ID
tableMvtP0("Text") = recMvtp0.Text
End Sub


'---------------------------------------------------------
Public Function tableMvtP0_Read(recMvtp0 As typeMvtP0) As Integer
'---------------------------------------------------------

On Error GoTo tableMvtP0_Read_Error
tableMvtP0_Read = 0


Select Case Trim(recMvtp0.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableMvtP0.Seek "=", recMvtp0.ID
                        If tableMvtP0.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableMvtP0.Seek "<=", recMvtp0.ID
                        If tableMvtP0.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableMvtP0.Seek ">=", recMvtp0.ID
                        If tableMvtP0.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableMvtP0.Seek ">", recMvtp0.ID
                        If tableMvtP0.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableMvtP0.MoveNext
                        If tableMvtP0.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableMvtP0.MovePrevious
                        If tableMvtP0.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableMvtP0.MoveFirst
                        If tableMvtP0.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableMvtP0.MoveLast
                        If tableMvtP0.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recMvtp0.Method <> "AddNew      " Then
    Call tableMvtP0_GetBuffer(recMvtp0)
End If

Exit Function

'---------------------------------------------------------
tableMvtP0_Read_Error:
'---------------------------------------------------------

    tableMvtP0_Read = Err
    Resume tableMvtP0_Read_End

tableMvtP0_Read_End:

End Function

'---------------------------------------------------------
Public Function tableMvtP0_Update(recMvtp0 As typeMvtP0) As Integer
'---------------------------------------------------------

On Error GoTo tableMvtP0Update_Error
tableMvtP0_Update = 0

Select Case Trim(recMvtp0.Method)

    Case "AddNew"
                        tableMvtP0.AddNew
                        Call tableMvtP0_PutBuffer(recMvtp0)
                        tableMvtP0.Update
    Case "Update"
                        tableMvtP0.Edit
                        Call tableMvtP0_PutBuffer(recMvtp0)
                        tableMvtP0.Update
    Case "Delete"
                        tableMvtP0.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableMvtP0Update_Error:
'---------------------------------------------------------
    tableMvtP0_Update = Err
    Resume tableMvtP0Update_End

tableMvtP0Update_End:

End Function








'-----------------------------------------------------
Sub dbMvtP0_Error(recMvtp0 As typeMvtP0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recMvtp0.ID & ": " & Chr$(13)

Select Case mId$(recMvtp0.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recMvtp0.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbMvtP0.bas :  ( " & Trim(recMvtp0.Obj) & " : " & Trim(recMvtp0.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbMvtP0_ReadE(recMvtp0 As typeMvtP0)
'-----------------------------------------------------

dbMvtP0_ReadE = Null

recMvtp0.Err = tableMvtP0_Read(recMvtp0)
If recMvtp0.Err > 0 Then

'    If recMvtP0.Err < 9990 Or recMvtP0.Err >= 9999 Then
        Call dbMvtP0_Error(recMvtp0)
        dbMvtP0_ReadE = recMvtp0.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbMvtP0_Update(recMvtp0 As typeMvtP0)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbMvtP0_Update = Null


recMvtp0.Err = tableMvtP0_Update(recMvtp0)

If recMvtp0.Err <> 0 Then
    Call dbMvtP0_Error(recMvtp0)
    dbMvtP0_Update = recMvtp0.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recMvtP0_Init(recMvtp0 As typeMvtP0)
recMvtp0.Method = ""
recMvtp0.Obj = "MvtP0"
recMvtp0.Err = ""
recMvtp0.ID = ""
recMvtp0.Text = ""
End Sub



