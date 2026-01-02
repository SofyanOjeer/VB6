Attribute VB_Name = "mdbCDXMvt"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableCDXMvt As Recordset
Dim tableCDXMvtOpen As Boolean
Public mCDXMvt_Id As Long

Type typeCDXMvt
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Id                  As String * 16
    Text                As String

End Type

Public recCDXMvt As typeCDXMvt


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableCDXMvt_Close()
'-----------------------------------------------------
If tableCDXMvtOpen Then
    tableCDXMvt.Close
    tableCDXMvtOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableCDXMvt_GetBuffer(recCDXMvt As typeCDXMvt)
'---------------------------------------------------------
recCDXMvt.Id = tableCDXMvt("Id")
recCDXMvt.Text = tableCDXMvt("Text")

End Sub


'-----------------------------------------------------
Sub tableCDXMvt_Open()
'-----------------------------------------------------

If Not tableCDXMvtOpen Then
    Set tableCDXMvt = MDB.OpenRecordset("CDXMvt")
    tableCDXMvt.Index = "PrimaryKey"
    tableCDXMvtOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableCDXMvt_PutBuffer(recCDXMvt As typeCDXMvt)
'---------------------------------------------------------

tableCDXMvt("Id") = recCDXMvt.Id
tableCDXMvt("Text") = recCDXMvt.Text
End Sub


'---------------------------------------------------------
Public Function tableCDXMvt_Read(recCDXMvt As typeCDXMvt) As Integer
'---------------------------------------------------------

On Error GoTo tableCDXMvt_Read_Error
tableCDXMvt_Read = 0


Select Case Trim(recCDXMvt.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableCDXMvt.Seek "=", recCDXMvt.Id
                        If tableCDXMvt.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableCDXMvt.Seek "<=", recCDXMvt.Id
                        If tableCDXMvt.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableCDXMvt.Seek ">=", recCDXMvt.Id
                        If tableCDXMvt.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableCDXMvt.Seek ">", recCDXMvt.Id
                        If tableCDXMvt.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableCDXMvt.MoveNext
                        If tableCDXMvt.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableCDXMvt.MovePrevious
                        If tableCDXMvt.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableCDXMvt.MoveFirst
                        If tableCDXMvt.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableCDXMvt.MoveLast
                        If tableCDXMvt.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recCDXMvt.Method <> "AddNew      " Then
    Call tableCDXMvt_GetBuffer(recCDXMvt)
End If

Exit Function

'---------------------------------------------------------
tableCDXMvt_Read_Error:
'---------------------------------------------------------

    tableCDXMvt_Read = Err
    Resume tableCDXMvt_Read_End

tableCDXMvt_Read_End:

End Function

'---------------------------------------------------------
Public Function tableCDXMvt_Update(recCDXMvt As typeCDXMvt) As Integer
'---------------------------------------------------------

On Error GoTo tableCDXMvtUpdate_Error
tableCDXMvt_Update = 0

Select Case Trim(recCDXMvt.Method)

    Case "AddNew"
                        tableCDXMvt.AddNew
                        Call tableCDXMvt_PutBuffer(recCDXMvt)
                        tableCDXMvt.Update
    Case "Update"
                        tableCDXMvt.Edit
                        Call tableCDXMvt_PutBuffer(recCDXMvt)
                        tableCDXMvt.Update
    Case "Delete"
                        tableCDXMvt.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableCDXMvtUpdate_Error:
'---------------------------------------------------------
    tableCDXMvt_Update = Err
    Resume tableCDXMvtUpdate_End

tableCDXMvtUpdate_End:

End Function

Public Function dbCDXMvt_Import(lFileName As String, lNb As Long)
Dim x As String, xInput As String
On Error GoTo Error_Handler

Dim I As Integer, blnOk As Boolean

lNb = 0: I = 0
x = Dir(lFileName)
If x = "" Then dbCDXMvt_Import = "? dbCDXMvt_Import : Le fichier des mouvments n'existe pas": Exit Function


MDB.Execute "delete * from CDXMvt"
mdbCDXMvt.tableCDXMvt_Open
recCDXMvt_Init recCDXMvt
recCDXMvt.Method = "AddNew"

Open lFileName For Input As #1

blnOk = False
Do Until EOF(1)
    Line Input #1, xInput
    
    If mId$(xInput, 1, 3) = "$$$" Then
        blnOk = True
        'SrvCptP0_Amj = mId$(xInput, 4, 8)
        I = Val(mId$(xInput, 12, 9))
        If I <> lNb Then
            x = "? dbCDXMvt_Import : nombre enregistrements lus"
            Call MsgBox(x, vbCritical, "dbCDXMvt_Import")
            dbCDXMvt_Import = x: Exit Function
            Exit Do
        End If
    End If

    lNb = lNb + 1
    recCDXMvt.Id = mId$(xInput, 24, 11)
    recCDXMvt.Method = "AddNew"
    recCDXMvt.Text = xInput
    dbCDXMvt_Update recCDXMvt
 
Loop

Close
mdbCptP0.tableCptP0_Close
mdbCDXMvt.tableCDXMvt_Close


If Not blnOk Then
    x = "? dbCDXMvt_Import : manque fin de fichier "
    Call MsgBox(x, vbCritical, "dbCDXMvt_Import")
    dbCDXMvt_Import = x: Exit Function
End If

Exit Function
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------
    x = "? dbCDXMvt_Import : " & Err & " : " & Error(Err)
    Call MsgBox(x, vbCritical, "dbCDXMvt_Import")
    dbCDXMvt_Import = x: Exit Function

End Function








'-----------------------------------------------------
Sub dbCDXMvt_Error(recCDXMvt As typeCDXMvt)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recCDXMvt.Id & ": " & Chr$(13)

Select Case mId$(recCDXMvt.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recCDXMvt.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbCDXMvt.bas :  ( " & Trim(recCDXMvt.obj) & " : " & Trim(recCDXMvt.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbCDXMvt_ReadE(recCDXMvt As typeCDXMvt)
'-----------------------------------------------------

dbCDXMvt_ReadE = Null

recCDXMvt.Err = tableCDXMvt_Read(recCDXMvt)
If recCDXMvt.Err > 0 Then

'    If recCDXMvt.Err < 9990 Or recCDXMvt.Err >= 9999 Then
        Call dbCDXMvt_Error(recCDXMvt)
        dbCDXMvt_ReadE = recCDXMvt.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbCDXMvt_Update(recCDXMvt As typeCDXMvt)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbCDXMvt_Update = Null


recCDXMvt.Err = tableCDXMvt_Update(recCDXMvt)

If recCDXMvt.Err <> 0 Then
    Call dbCDXMvt_Error(recCDXMvt)
    dbCDXMvt_Update = recCDXMvt.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recCDXMvt_Init(recCDXMvt As typeCDXMvt)
recCDXMvt.Method = ""
recCDXMvt.obj = "CDXMvt"
recCDXMvt.Err = ""
recCDXMvt.Id = ""
recCDXMvt.Text = ""
End Sub



