Attribute VB_Name = "mdbElpTable"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableElpTable As Recordset
Dim tableElpTableOpen As Boolean

Type typeElpTable
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    ID           As String * 12
    K1           As String * 12
    K2           As String * 12
    SNN          As Long
    SNP          As Long
    SN           As Long
    Chrono       As Long
    Name         As String * 36
    Dmin         As String * 8
    Dmax         As String * 8
    Memo         As Variant

End Type

Public recElpTable As typeElpTable
Public xElpTable As typeElpTable
Public recPériodicité As typeElpTable, recStatut As typeElpTable


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableElpTable_Close()
'-----------------------------------------------------
If tableElpTableOpen Then
    tableElpTable.Close
    tableElpTableOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableElpTable_GetBuffer(recElpTable As typeElpTable)
'---------------------------------------------------------
Dim X As String

recElpTable.ID = tableElpTable("Id")
recElpTable.K1 = tableElpTable("K1")
recElpTable.K2 = tableElpTable("K2")

recElpTable.SNN = tableElpTable("SNN")
recElpTable.SNP = tableElpTable("SNP")
recElpTable.SN = tableElpTable("SN")
recElpTable.Chrono = tableElpTable("Chrono")
recElpTable.Name = tableElpTable("Name")
recElpTable.Dmin = tableElpTable("DMin")
recElpTable.Dmax = tableElpTable("DMax")
recElpTable.Memo = tableElpTable("Memo")

If Trim(recElpTable.K1) = "PasswordX" Then
    If Not IsNull(recElpTable.Memo) Then
        X = Trim(recElpTable.Memo)
        recElpTable.Memo = ElpCipher_D(X, paramElpCypher)
    Else
        recElpTable.Memo = ""
    End If
End If
End Sub

Public Function recPériodicité_Libellé(mPériodicité As String) As String
If Trim(recPériodicité.K2) <> Trim(mPériodicité) Then
    recPériodicité.Method = "Seek="
    recPériodicité.K2 = mPériodicité
    recPériodicité.Err = tableElpTable_Read(recPériodicité)
    If recPériodicité.Err <> 0 Then recPériodicité.Name = mPériodicité
End If
recPériodicité_Libellé = recPériodicité.Name
End Function

Public Function recStatut_Libellé(mStatut As String) As String
If Trim(recStatut.K2) <> Trim(mStatut) Then
    recStatut.Method = "Seek="
    recStatut.K2 = mStatut
    recStatut.Err = tableElpTable_Read(recStatut)
    If recStatut.Err <> 0 Then recStatut.Name = mStatut
End If
recStatut_Libellé = recStatut.Name
End Function

'-----------------------------------------------------
Sub tableElpTable_Open()
'-----------------------------------------------------

If Not tableElpTableOpen Then
    Set tableElpTable = MDB.OpenRecordset("ElpTable")
    tableElpTable.Index = "PrimaryKey"
    tableElpTableOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableElpTable_PutBuffer(recElpTable As typeElpTable)
'---------------------------------------------------------

tableElpTable("id") = recElpTable.ID
tableElpTable("K1") = recElpTable.K1
tableElpTable("K2") = recElpTable.K2
tableElpTable("SNN") = recElpTable.SNN
tableElpTable("SNP") = recElpTable.SNP
tableElpTable("SN") = recElpTable.SN
tableElpTable("Chrono") = recElpTable.Chrono
tableElpTable("Name") = recElpTable.Name

tableElpTable("DMin") = recElpTable.Dmin
tableElpTable("DMax") = recElpTable.Dmax
tableElpTable("Memo") = recElpTable.Memo
End Sub


'---------------------------------------------------------
Public Function tableElpTable_Read(recElpTable As typeElpTable) As Integer
'---------------------------------------------------------

On Error GoTo tableElpTable_Read_Error
tableElpTable_Read = 0


Select Case Trim(recElpTable.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableElpTable.Seek "=", recElpTable.ID, recElpTable.K1, recElpTable.K2, recElpTable.SNN
                        If tableElpTable.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableElpTable.Seek "<=", recElpTable.ID, recElpTable.K1, recElpTable.K2, recElpTable.SNN
                        If tableElpTable.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableElpTable.Seek ">=", recElpTable.ID, recElpTable.K1, recElpTable.K2, recElpTable.SNN
                        If tableElpTable.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>"
                        tableElpTable.Seek ">", recElpTable.ID, recElpTable.K1, recElpTable.K2, recElpTable.SNN
                        If tableElpTable.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableElpTable.MoveNext
                        If tableElpTable.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableElpTable.MovePrevious
                        If tableElpTable.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableElpTable.MoveFirst
                        If tableElpTable.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableElpTable.MoveLast
                        If tableElpTable.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recElpTable.Method <> "AddNew      " Then
    Call tableElpTable_GetBuffer(recElpTable)
End If

Exit Function

'---------------------------------------------------------
tableElpTable_Read_Error:
'---------------------------------------------------------

    tableElpTable_Read = Err
    Resume tableElpTable_Read_End

tableElpTable_Read_End:

End Function

'---------------------------------------------------------
Public Function tableElpTable_Update(recElpTable As typeElpTable) As Integer
'---------------------------------------------------------

On Error GoTo tableElpTableUpdate_Error
tableElpTable_Update = 0

Select Case Trim(recElpTable.Method)

    Case "AddNew"
                        tableElpTable.AddNew
                        Call tableElpTable_PutBuffer(recElpTable)
                        tableElpTable.Update
    Case "Update"
                        tableElpTable.Edit
                        Call tableElpTable_PutBuffer(recElpTable)
                        tableElpTable.Update
    Case "Delete"
                        tableElpTable.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableElpTableUpdate_Error:
'---------------------------------------------------------
    tableElpTable_Update = Err
    Resume tableElpTableUpdate_End

tableElpTableUpdate_End:

End Function








'-----------------------------------------------------
Sub dbElpTable_Error(recElpTable As typeElpTable)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & Trim(recElpTable.ID) & " : " & Trim(recElpTable.K1) & " : " & Trim(recElpTable.K2) & Chr$(13)

Select Case mId$(recElpTable.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recElpTable.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbElpTable.bas :  ( " & Trim(recElpTable.obj) & " : " & Trim(recElpTable.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbElpTable_ReadE(recElpTable As typeElpTable)
'-----------------------------------------------------

dbElpTable_ReadE = Null

recElpTable.Err = tableElpTable_Read(recElpTable)
If recElpTable.Err > 0 Then

'    If recElpTable.Err < 9990 Or recElpTable.Err >= 9999 Then
        Call dbElpTable_Error(recElpTable)
        dbElpTable_ReadE = recElpTable.Err
'    End If
End If

End Function

'-----------------------------------------------------
Sub lstElpTable_Load(lstX As ListBox, recElpTable As typeElpTable, kSelect As Integer, kDisplay As Integer)
'-----------------------------------------------------
Dim mId As String, mK1 As String

lstX.Clear
recElpTable.Err = 0
recElpTable.Method = "Seek>="
mId = recElpTable.ID
mK1 = recElpTable.K1
Do
    recElpTable.Err = tableElpTable_Read(recElpTable)
    If recElpTable.Err = 0 Then
        If mId <> recElpTable.ID Then
            recElpTable.Err = 9996
        Else
            If kSelect > 0 And mK1 <> recElpTable.K1 Then
                recElpTable.Err = 9996
            Else
                 Select Case kDisplay
                    Case 1: lstX.AddItem recElpTable.K1 & " : " & Trim(recElpTable.Name)
                    Case 2: lstX.AddItem recElpTable.K2 & " : " & Trim(recElpTable.Name)
                    Case Else: lstX.AddItem recElpTable.K1 & " " & recElpTable.K2 & " : " & Trim(recElpTable.Name)
                 End Select
                
                recElpTable.Method = "Seek>"
            
            End If
        End If
    End If
    
Loop While recElpTable.Err = 0
End Sub


'-----------------------------------------------------
Function dbElpTable_Update(recElpTable As typeElpTable)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbElpTable_Update = Null
recElpTable.Err = 0

If recElpTable.Method = constAddNew Then
    recElpTable.Err = 0
Else
    tableElpTable.Seek "=", recElpTable.ID, recElpTable.K1, recElpTable.K2, recElpTable.SNN
    If tableElpTable.NoMatch Then recElpTable.Err = 9998
End If

If recElpTable.Err = 0 Then recElpTable.Err = tableElpTable_Update(recElpTable)

If recElpTable.Err <> 0 Then
    Call dbElpTable_Error(recElpTable)
    dbElpTable_Update = recElpTable.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recElpTable_Init(recElpTable As typeElpTable)
recElpTable.Method = ""
recElpTable.obj = "ElpTable"
recElpTable.Err = ""
recElpTable.ID = ""
recElpTable.K1 = ""
recElpTable.K2 = ""
recElpTable.SNN = 0
recElpTable.SNP = 0
recElpTable.SN = 0
recElpTable.Chrono = 0
recElpTable.Name = ""
recElpTable.Dmin = "00000000"
recElpTable.Dmax = "00000000"
recElpTable.Memo = ""
End Sub

Public Sub recPériodicité_Init()
recElpTable_Init recPériodicité
recPériodicité.Method = "Seek="
recPériodicité.ID = "Param"
recPériodicité.K1 = "Périodicité"

End Sub
Public Sub recStatut_Init()
recElpTable_Init recStatut
recStatut.Method = "Seek="
recStatut.ID = "Param"
recStatut.K1 = "Statut"

End Sub


