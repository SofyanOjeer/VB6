Attribute VB_Name = "mdbSAB_ZMNU"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableSAB_ZMNU As Recordset
Dim tableSAB_ZMNUOpen As Boolean
Public mSAB_ZMNU_Id As Long

Type typeSAB_ZMNU
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Src                     As String * 3
    Id                      As String * 20
    Memo                    As Variant

End Type

Public recSAB_ZMNU As typeSAB_ZMNU

Type typeSAB_ZMNURUT0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    MNURUTUTI               As String * 10
    MNURUTNOM               As String * 30
    MNURUTETB               As String * 5 '3
    MNURUTCUT               As String * 5 '7
    MNURUTLOG               As String * 1

    MNUUTIETB               As String * 5 '3
    MNUUTICUT               As String * 5 '7
    MNUUTICGR               As String * 5 '7
    MNUUTIDRG               As String * 1
    MNUUTIOUT               As String * 10
    MNUUTILAN               As String * 1
    MNUUTIMSE               As String * 1
    MNUUTIAGE               As String * 5 '3
    MNUUTISER               As String * 2
    MNUUTISRV               As String * 2

End Type

Public recSAB_ZMNURUT0 As typeSAB_ZMNURUT0

'---------------------------------------------------------
'-----------------------------------------------------
Sub tableSAB_ZMNU_Close()
'-----------------------------------------------------
If tableSAB_ZMNUOpen Then
    tableSAB_ZMNU.Close
    tableSAB_ZMNUOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableSAB_ZMNU_GetBuffer(recSAB_ZMNU As typeSAB_ZMNU)
'---------------------------------------------------------
recSAB_ZMNU.Src = tableSAB_ZMNU("Src")
recSAB_ZMNU.Id = tableSAB_ZMNU("Id")
recSAB_ZMNU.Memo = tableSAB_ZMNU("Memo")

End Sub


'-----------------------------------------------------
Sub tableSAB_ZMNU_Open()
'-----------------------------------------------------

If Not tableSAB_ZMNUOpen Then
    Set tableSAB_ZMNU = MDB.OpenRecordset("SAB_ZMNU")
    tableSAB_ZMNU.Index = "PrimaryKey"
    tableSAB_ZMNUOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableSAB_ZMNU_PutBuffer(recSAB_ZMNU As typeSAB_ZMNU)
'---------------------------------------------------------

tableSAB_ZMNU("Src") = recSAB_ZMNU.Src
tableSAB_ZMNU("Id") = recSAB_ZMNU.Id
tableSAB_ZMNU("Memo") = recSAB_ZMNU.Memo
End Sub


'---------------------------------------------------------
Public Function tableSAB_ZMNU_Read(recSAB_ZMNU As typeSAB_ZMNU) As Integer
'---------------------------------------------------------

On Error GoTo tableSAB_ZMNU_Read_Error
tableSAB_ZMNU_Read = 0


Select Case Trim(recSAB_ZMNU.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableSAB_ZMNU.Seek "=", recSAB_ZMNU.Src, recSAB_ZMNU.Id
                        If tableSAB_ZMNU.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableSAB_ZMNU.Seek "<=", recSAB_ZMNU.Src, recSAB_ZMNU.Id
                        If tableSAB_ZMNU.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableSAB_ZMNU.Seek ">=", recSAB_ZMNU.Src, recSAB_ZMNU.Id
                        If tableSAB_ZMNU.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableSAB_ZMNU.Seek ">", recSAB_ZMNU.Src, recSAB_ZMNU.Id
                        If tableSAB_ZMNU.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableSAB_ZMNU.MoveNext
                        If tableSAB_ZMNU.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableSAB_ZMNU.MovePrevious
                        If tableSAB_ZMNU.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableSAB_ZMNU.MoveFirst
                        If tableSAB_ZMNU.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableSAB_ZMNU.MoveLast
                        If tableSAB_ZMNU.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recSAB_ZMNU.Method <> "AddNew      " Then
    Call tableSAB_ZMNU_GetBuffer(recSAB_ZMNU)
End If

Exit Function

'---------------------------------------------------------
tableSAB_ZMNU_Read_Error:
'---------------------------------------------------------

    tableSAB_ZMNU_Read = Err
    Resume tableSAB_ZMNU_Read_End

tableSAB_ZMNU_Read_End:

End Function

'---------------------------------------------------------
Public Function tableSAB_ZMNU_Update(recSAB_ZMNU As typeSAB_ZMNU) As Integer
'---------------------------------------------------------

On Error GoTo tableSAB_ZMNUUpdate_Error
tableSAB_ZMNU_Update = 0

Select Case Trim(recSAB_ZMNU.Method)

    Case "AddNew"
                        tableSAB_ZMNU.AddNew
                        Call tableSAB_ZMNU_PutBuffer(recSAB_ZMNU)
                        tableSAB_ZMNU.Update
    Case "Update"
                        tableSAB_ZMNU.Edit
                        Call tableSAB_ZMNU_PutBuffer(recSAB_ZMNU)
                        tableSAB_ZMNU.Update
    Case "Delete"
                        tableSAB_ZMNU.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableSAB_ZMNUUpdate_Error:
'---------------------------------------------------------
    tableSAB_ZMNU_Update = Err
    Resume tableSAB_ZMNUUpdate_End

tableSAB_ZMNUUpdate_End:

End Function








'-----------------------------------------------------
Sub dbSAB_ZMNU_Error(recSAB_ZMNU As typeSAB_ZMNU)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recSAB_ZMNU.Id & ": " & recSAB_ZMNU.Src & Chr$(13)

Select Case mId$(recSAB_ZMNU.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recSAB_ZMNU.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbSAB_ZMNU.bas :  ( " & Trim(recSAB_ZMNU.obj) & " : " & Trim(recSAB_ZMNU.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbSAB_ZMNU_ReadE(recSAB_ZMNU As typeSAB_ZMNU)
'-----------------------------------------------------

dbSAB_ZMNU_ReadE = Null

recSAB_ZMNU.Err = tableSAB_ZMNU_Read(recSAB_ZMNU)
If recSAB_ZMNU.Err > 0 Then

'    If recSAB_ZMNU.Err < 9990 Or recSAB_ZMNU.Err >= 9999 Then
        Call dbSAB_ZMNU_Error(recSAB_ZMNU)
        dbSAB_ZMNU_ReadE = recSAB_ZMNU.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbSAB_ZMNU_Update(recSAB_ZMNU As typeSAB_ZMNU)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbSAB_ZMNU_Update = Null


recSAB_ZMNU.Err = tableSAB_ZMNU_Update(recSAB_ZMNU)

If recSAB_ZMNU.Err <> 0 Then
    Call dbSAB_ZMNU_Error(recSAB_ZMNU)
    dbSAB_ZMNU_Update = recSAB_ZMNU.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recSAB_ZMNU_Init(recSAB_ZMNU As typeSAB_ZMNU)
recSAB_ZMNU.Method = ""
recSAB_ZMNU.obj = "SAB_ZMNU"
recSAB_ZMNU.Err = ""
recSAB_ZMNU.Id = ""
recSAB_ZMNU.Src = ""
recSAB_ZMNU.Memo = Null
End Sub


