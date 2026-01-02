Attribute VB_Name = "mdbElpKMLink"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableElpKMLink As Recordset
Dim tableElpKMLinkOpen As Boolean
Public mElpKMLink_Id As Long

Type typeElpKMLink
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    ElpKMSrc_Id             As Long
    ElpKMInfo_Id            As String * 20
    Id                      As String * 20
    Pass                    As Long
    Document_Extension      As String * 3
    Document_Id             As Variant
    Memo                    As Variant

End Type

Public recElpKMLink As typeElpKMLink


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableElpKMLink_Close()
'-----------------------------------------------------
If tableElpKMLinkOpen Then
    tableElpKMLink.Close
    tableElpKMLinkOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableElpKMLink_GetBuffer(recElpKMLink As typeElpKMLink)
'---------------------------------------------------------
recElpKMLink.ElpKMSrc_Id = tableElpKMLink("ElpKMSrc_Id")
recElpKMLink.ElpKMInfo_Id = tableElpKMLink("ElpKMInfo_Id")
recElpKMLink.Id = tableElpKMLink("Id")
recElpKMLink.Pass = tableElpKMLink("Pass")
recElpKMLink.Document_Extension = tableElpKMLink("Document_Extension")
recElpKMLink.Document_Id = tableElpKMLink("Document_Id")
recElpKMLink.Memo = tableElpKMLink("Memo")

End Sub


'-----------------------------------------------------
Sub tableElpKMLink_Open()
'-----------------------------------------------------

If Not tableElpKMLinkOpen Then
    Set tableElpKMLink = MDB.OpenRecordset("ElpKMLink")
    tableElpKMLink.Index = "PrimaryKey"
    tableElpKMLinkOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableElpKMLink_PutBuffer(recElpKMLink As typeElpKMLink)
'---------------------------------------------------------

tableElpKMLink("ElpKMSrc_Id") = recElpKMLink.ElpKMSrc_Id
tableElpKMLink("ElpKMInfo_Id") = recElpKMLink.ElpKMInfo_Id
tableElpKMLink("Id") = recElpKMLink.Id
tableElpKMLink("Pass") = recElpKMLink.Pass
tableElpKMLink("Document_Extension") = recElpKMLink.Document_Extension
tableElpKMLink("Document_Id") = recElpKMLink.Document_Id
tableElpKMLink("Memo") = recElpKMLink.Memo

End Sub


'---------------------------------------------------------
Public Function tableElpKMLink_Read(recElpKMLink As typeElpKMLink) As Integer
'---------------------------------------------------------

On Error GoTo tableElpKMLink_Read_Error
tableElpKMLink_Read = 0


Select Case Trim(recElpKMLink.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableElpKMLink.Seek "=", recElpKMLink.ElpKMSrc_Id, recElpKMLink.ElpKMInfo_Id, recElpKMLink.Id
                        If tableElpKMLink.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableElpKMLink.Seek "<=", recElpKMLink.ElpKMSrc_Id, recElpKMLink.ElpKMInfo_Id, recElpKMLink.Id
                        If tableElpKMLink.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableElpKMLink.Seek ">=", recElpKMLink.ElpKMSrc_Id, recElpKMLink.ElpKMInfo_Id, recElpKMLink.Id
                        If tableElpKMLink.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableElpKMLink.Seek ">", recElpKMLink.ElpKMSrc_Id, recElpKMLink.ElpKMInfo_Id, recElpKMLink.Id
                        If tableElpKMLink.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableElpKMLink.MoveNext
                        If tableElpKMLink.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableElpKMLink.MovePrevious
                        If tableElpKMLink.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableElpKMLink.MoveFirst
                        If tableElpKMLink.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableElpKMLink.MoveLast
                        If tableElpKMLink.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recElpKMLink.Method <> "AddNew      " Then
    Call tableElpKMLink_GetBuffer(recElpKMLink)
End If

Exit Function

'---------------------------------------------------------
tableElpKMLink_Read_Error:
'---------------------------------------------------------

    tableElpKMLink_Read = Err
    Resume tableElpKMLink_Read_End

tableElpKMLink_Read_End:

End Function

'---------------------------------------------------------
Public Function tableElpKMLink_Update(recElpKMLink As typeElpKMLink) As Integer
'---------------------------------------------------------

On Error GoTo tableElpKMLinkUpdate_Error
tableElpKMLink_Update = 0

Select Case Trim(recElpKMLink.Method)

    Case "AddNew"
                        tableElpKMLink.AddNew
                        Call tableElpKMLink_PutBuffer(recElpKMLink)
                        tableElpKMLink.Update
    Case "Update"
                        tableElpKMLink.Edit
                        Call tableElpKMLink_PutBuffer(recElpKMLink)
                        tableElpKMLink.Update
    Case "Delete"
                        tableElpKMLink.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableElpKMLinkUpdate_Error:
'---------------------------------------------------------
    tableElpKMLink_Update = Err
    Resume tableElpKMLinkUpdate_End

tableElpKMLinkUpdate_End:

End Function








'-----------------------------------------------------
Sub dbElpKMLink_Error(recElpKMLink As typeElpKMLink)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recElpKMLink.Id & ": " & recElpKMLink.ElpKMSrc_Id & Chr$(13)

Select Case mId$(recElpKMLink.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recElpKMLink.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbElpKMLink.bas :  ( " & Trim(recElpKMLink.obj) & " : " & Trim(recElpKMLink.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbElpKMLink_ReadE(recElpKMLink As typeElpKMLink)
'-----------------------------------------------------

dbElpKMLink_ReadE = Null

recElpKMLink.Err = tableElpKMLink_Read(recElpKMLink)
If recElpKMLink.Err > 0 Then

'    If recElpKMLink.Err < 9990 Or recElpKMLink.Err >= 9999 Then
        Call dbElpKMLink_Error(recElpKMLink)
        dbElpKMLink_ReadE = recElpKMLink.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbElpKMLink_Update(recElpKMLink As typeElpKMLink)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbElpKMLink_Update = Null


recElpKMLink.Err = tableElpKMLink_Update(recElpKMLink)

If recElpKMLink.Err <> 0 Then
    Call dbElpKMLink_Error(recElpKMLink)
    dbElpKMLink_Update = recElpKMLink.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recElpKMLink_Init(recElpKMLink As typeElpKMLink)
recElpKMLink.Method = ""
recElpKMLink.obj = "ElpKMLinkInfo"
recElpKMLink.Err = ""
recElpKMLink.Id = ""
recElpKMLink.ElpKMSrc_Id = 0
recElpKMLink.Document_Extension = ""
recElpKMLink.Pass = 0
recElpKMLink.Memo = Null
End Sub


