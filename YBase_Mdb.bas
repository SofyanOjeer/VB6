Attribute VB_Name = "mdbYBase"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableYBase As Recordset
Dim tableYBaseOpen As Boolean
Public mYBase_Id As Long
Public Const constYBase = "YBase     "

Type typeYBase
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    ID                  As String * 12
    K1                  As String * 50
    Text                As String

End Type

Public recYBase As typeYBase


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableYBase_Close()
'-----------------------------------------------------
If tableYBaseOpen Then
    tableYBase.Close
    tableYBaseOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableYBase_GetBuffer(recYBase As typeYBase)
'---------------------------------------------------------
recYBase.ID = tableYBase("Id")
recYBase.K1 = tableYBase("K1")
recYBase.Text = tableYBase("Text")

End Sub


'-----------------------------------------------------
Sub tableYBase_Open()
'-----------------------------------------------------

If Not tableYBaseOpen Then
    Set tableYBase = MDB.OpenRecordset("YBase")
    tableYBase.Index = "PrimaryKey"
    tableYBaseOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableYBase_PutBuffer(recYBase As typeYBase)
'---------------------------------------------------------

tableYBase("Id") = recYBase.ID
tableYBase("K1") = recYBase.K1
tableYBase("Text") = recYBase.Text
End Sub


'---------------------------------------------------------
Public Function tableYBase_Read(recYBase As typeYBase) As Integer
'---------------------------------------------------------

On Error GoTo tableYBase_Read_Error
tableYBase_Read = 0


Select Case Trim(recYBase.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableYBase.Seek "=", recYBase.ID, recYBase.K1
                        If tableYBase.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableYBase.Seek "<=", recYBase.ID, recYBase.K1
                        If tableYBase.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableYBase.Seek ">=", recYBase.ID, recYBase.K1
                        If tableYBase.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableYBase.Seek ">", recYBase.ID, recYBase.K1
                        If tableYBase.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableYBase.MoveNext
                        If tableYBase.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableYBase.MovePrevious
                        If tableYBase.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableYBase.MoveFirst
                        If tableYBase.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableYBase.MoveLast
                        If tableYBase.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recYBase.Method <> "AddNew      " Then
    Call tableYBase_GetBuffer(recYBase)
End If

Exit Function

'---------------------------------------------------------
tableYBase_Read_Error:
'---------------------------------------------------------

    tableYBase_Read = Err
    Resume tableYBase_Read_End

tableYBase_Read_End:

End Function

'---------------------------------------------------------
Public Function tableYBase_Update(recYBase As typeYBase) As Integer
'---------------------------------------------------------

On Error GoTo tableYBaseUpdate_Error
tableYBase_Update = 0

Select Case Trim(recYBase.Method)

    Case "AddNew"
                        tableYBase.AddNew
                        Call tableYBase_PutBuffer(recYBase)
                        tableYBase.Update
    Case "Update"
                        tableYBase.Edit
                        Call tableYBase_PutBuffer(recYBase)
                        tableYBase.Update
    Case "Delete"
                        tableYBase.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableYBaseUpdate_Error:
'---------------------------------------------------------
    tableYBase_Update = Err
    Resume tableYBaseUpdate_End

tableYBaseUpdate_End:

End Function








'-----------------------------------------------------
Sub dbYBase_Error(recYBase As typeYBase)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recYBase.ID & recYBase.K1 & ": " & Chr$(13)

Select Case mId$(recYBase.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recYBase.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbYBase.bas :  ( " & Trim(recYBase.obj) & " : " & Trim(recYBase.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbYBase_ReadE(recYBase As typeYBase)
'-----------------------------------------------------

dbYBase_ReadE = Null

recYBase.Err = tableYBase_Read(recYBase)
If recYBase.Err > 0 Then

'    If recYBase.Err < 9990 Or recYBase.Err >= 9999 Then
        Call dbYBase_Error(recYBase)
        dbYBase_ReadE = recYBase.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbYBase_Update(recYBase As typeYBase)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbYBase_Update = Null


recYBase.Err = tableYBase_Update(recYBase)

If recYBase.Err <> 0 Then
    Call dbYBase_Error(recYBase)
    dbYBase_Update = recYBase.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recYBase_Init(recYBase As typeYBase)
recYBase.Method = ""
recYBase.obj = "YBase"
recYBase.Err = ""
recYBase.ID = ""
recYBase.K1 = ""
recYBase.Text = ""
End Sub





