Attribute VB_Name = "mdbSwiftHisto"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableSwiftHisto As Recordset
Dim tableSwiftHistoOpen As Boolean

Type typeSwiftHisto
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    RcvSnd                  As String * 1
    MT                      As String * 3
    BIC                     As String * 11
    Id                      As String * 16
    F20                     As String * 32
    F21                     As String * 32
    F32D                    As String * 8
    F32C                    As String * 3
    F32A                    As Currency
    Unit                    As String * 10
    AMJ                     As String * 8
    HMS                     As String * 6
    Text                    As String

End Type

Public recSwiftHisto As typeSwiftHisto

'---------------------------------------------------------
'-----------------------------------------------------
Sub tableSwiftHisto_Close()
'-----------------------------------------------------
If tableSwiftHistoOpen Then
    tableSwiftHisto.Close
    tableSwiftHistoOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableSwiftHisto_GetBuffer(recSwiftHisto As typeSwiftHisto)
'---------------------------------------------------------
recSwiftHisto.RcvSnd = tableSwiftHisto("RcvSnd")
recSwiftHisto.MT = tableSwiftHisto("MT")
recSwiftHisto.BIC = tableSwiftHisto("BIC")
recSwiftHisto.Id = tableSwiftHisto("Id")
recSwiftHisto.F20 = tableSwiftHisto("F20")
recSwiftHisto.F21 = tableSwiftHisto("F21")
recSwiftHisto.F32C = tableSwiftHisto("F32C")
recSwiftHisto.F32D = tableSwiftHisto("F32D")
recSwiftHisto.F32A = tableSwiftHisto("F32A")
recSwiftHisto.Unit = tableSwiftHisto("Unit")
recSwiftHisto.AMJ = tableSwiftHisto("AMJ")
recSwiftHisto.HMS = tableSwiftHisto("HMS")
recSwiftHisto.Text = tableSwiftHisto("Text")

End Sub


'-----------------------------------------------------
Sub tableSwiftHisto_Open()
'-----------------------------------------------------

If Not tableSwiftHistoOpen Then
    Set tableSwiftHisto = MDB.OpenRecordset("SwiftHisto")
    tableSwiftHisto.Index = "PrimaryKey"
    tableSwiftHistoOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableSwiftHisto_PutBuffer(recSwiftHisto As typeSwiftHisto)
'---------------------------------------------------------

tableSwiftHisto("RcvSnd") = recSwiftHisto.RcvSnd
tableSwiftHisto("MT") = recSwiftHisto.MT
tableSwiftHisto("BIC") = recSwiftHisto.BIC
tableSwiftHisto("Id") = recSwiftHisto.Id
tableSwiftHisto("F20") = recSwiftHisto.F20
tableSwiftHisto("F21") = recSwiftHisto.F21
tableSwiftHisto("F32C") = recSwiftHisto.F32C
tableSwiftHisto("F32D") = recSwiftHisto.F32D
tableSwiftHisto("F32A") = recSwiftHisto.F32A
tableSwiftHisto("Unit") = recSwiftHisto.Unit
tableSwiftHisto("AMJ") = recSwiftHisto.AMJ
tableSwiftHisto("HMS") = recSwiftHisto.HMS
tableSwiftHisto("Text") = recSwiftHisto.Text
End Sub


'---------------------------------------------------------
Public Function tableSwiftHisto_Read(recSwiftHisto As typeSwiftHisto) As Integer
'---------------------------------------------------------

On Error GoTo tableSwiftHisto_Read_Error
tableSwiftHisto_Read = 0


Select Case Trim(recSwiftHisto.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableSwiftHisto.Seek "=", recSwiftHisto.RcvSnd, recSwiftHisto.MT, recSwiftHisto.BIC, recSwiftHisto.Id
                        If tableSwiftHisto.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableSwiftHisto.Seek "<=", recSwiftHisto.RcvSnd, recSwiftHisto.MT, recSwiftHisto.BIC, recSwiftHisto.Id
                        If tableSwiftHisto.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableSwiftHisto.Seek ">=", recSwiftHisto.RcvSnd, recSwiftHisto.MT, recSwiftHisto.BIC, recSwiftHisto.Id
                        If tableSwiftHisto.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>", recSwiftHisto.RcvSnd, recSwiftHisto.MT, recSwiftHisto.BIC, recSwiftHisto.Id, recSwiftHisto.Id
                        If tableSwiftHisto.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableSwiftHisto.MoveNext
                        If tableSwiftHisto.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableSwiftHisto.MovePrevious
                        If tableSwiftHisto.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableSwiftHisto.MoveFirst
                        If tableSwiftHisto.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableSwiftHisto.MoveLast
                        If tableSwiftHisto.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recSwiftHisto.Method <> "AddNew      " Then
    Call tableSwiftHisto_GetBuffer(recSwiftHisto)
End If

Exit Function

'---------------------------------------------------------
tableSwiftHisto_Read_Error:
'---------------------------------------------------------

    tableSwiftHisto_Read = Err
    Resume tableSwiftHisto_Read_End

tableSwiftHisto_Read_End:

End Function
'---------------------------------------------------------
Public Function tableSwiftHisto_Update(recSwiftHisto As typeSwiftHisto) As Integer
'---------------------------------------------------------

On Error GoTo tableSwiftHistoUpdate_Error
tableSwiftHisto_Update = 0

Select Case Trim(recSwiftHisto.Method)

    Case "AddNew"
                        tableSwiftHisto.AddNew
                        Call tableSwiftHisto_PutBuffer(recSwiftHisto)
                        tableSwiftHisto.Update
    Case "Update"
                        tableSwiftHisto.Edit
                        Call tableSwiftHisto_PutBuffer(recSwiftHisto)
                        tableSwiftHisto.Update
    Case "Delete"
                        tableSwiftHisto.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableSwiftHistoUpdate_Error:
'---------------------------------------------------------
    tableSwiftHisto_Update = Err
    Resume tableSwiftHistoUpdate_End

tableSwiftHistoUpdate_End:

End Function








'-----------------------------------------------------
Sub dbSwiftHisto_Error(recSwiftHisto As typeSwiftHisto)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recSwiftHisto.Id & ": " & Chr$(13)

Select Case mId$(recSwiftHisto.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recSwiftHisto.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbSwiftHisto.bas :  ( " & Trim(recSwiftHisto.obj) & " : " & Trim(recSwiftHisto.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbSwiftHisto_ReadE(recSwiftHisto As typeSwiftHisto)
'-----------------------------------------------------

dbSwiftHisto_ReadE = Null

recSwiftHisto.Err = tableSwiftHisto_Read(recSwiftHisto)
If recSwiftHisto.Err > 0 Then

'    If recSwiftHisto.Err < 9990 Or recSwiftHisto.Err >= 9999 Then
        Call dbSwiftHisto_Error(recSwiftHisto)
        dbSwiftHisto_ReadE = recSwiftHisto.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbSwiftHisto_Update(recSwiftHisto As typeSwiftHisto)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbSwiftHisto_Update = Null


recSwiftHisto.Err = tableSwiftHisto_Update(recSwiftHisto)

If recSwiftHisto.Err <> 0 Then
    Call dbSwiftHisto_Error(recSwiftHisto)
    dbSwiftHisto_Update = recSwiftHisto.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recSwiftHisto_Init(recSwiftHisto As typeSwiftHisto)
recSwiftHisto.Method = ""
recSwiftHisto.obj = "   SwiftHisto"
recSwiftHisto.Err = ""
recSwiftHisto.RcvSnd = ""
recSwiftHisto.MT = ""
recSwiftHisto.BIC = ""
recSwiftHisto.Id = ""
recSwiftHisto.F20 = ""
recSwiftHisto.F21 = ""
recSwiftHisto.F32A = 0
recSwiftHisto.F32C = ""
recSwiftHisto.F32D = ""

recSwiftHisto.Unit = ""
recSwiftHisto.AMJ = "00000000"
recSwiftHisto.HMS = "000000"
recSwiftHisto.Text = " "

End Sub


