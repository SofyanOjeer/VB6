Attribute VB_Name = "mdbFicDatP1"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableFicDatP1 As Recordset
Dim tableFicDatP1Open As Boolean

Type typeFicDatP1
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    DAAMJ               As String * 8
    DACTRJ              As Long
    DACTRA              As Long
    DALIBJ              As String * 8
    DALIBM              As String * 9
    DALBMR              As String * 4
    DAFERJ              As String * 1
    DATSTC              As String * 1

End Type

Public paramDateBIA As typeFicDatP1
'---------------------------------------------------------
'-----------------------------------------------------
Sub tableFicDatP1_Close()
'-----------------------------------------------------
If tableFicDatP1Open Then
    tableFicDatP1.Close
    tableFicDatP1Open = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableFicDatP1_GetBuffer(recFicDatP1 As typeFicDatP1)
'---------------------------------------------------------
recFicDatP1.DAAMJ = tableFicDatP1("DAAMJ")
recFicDatP1.DACTRJ = tableFicDatP1("DACTRJ")
recFicDatP1.DACTRA = tableFicDatP1("DACTRA")
recFicDatP1.DALIBJ = tableFicDatP1("DALIBJ")
recFicDatP1.DALIBM = tableFicDatP1("DALIBM")
recFicDatP1.DALBMR = tableFicDatP1("DALBMR")
recFicDatP1.DAFERJ = tableFicDatP1("DAFERJ")
recFicDatP1.DATSTC = tableFicDatP1("DATSTC")

End Sub


'-----------------------------------------------------
Sub tableFicDatP1_Open()
'-----------------------------------------------------

If Not tableFicDatP1Open Then
    Set tableFicDatP1 = MDB.OpenRecordset("FicDatP1")
    tableFicDatP1.Index = "PrimaryKey"
    tableFicDatP1Open = True
    recFicDatP1_Init paramDateBIA
End If
End Sub

'---------------------------------------------------------
Public Sub tableFicDatP1_PutBuffer(recFicDatP1 As typeFicDatP1)
'---------------------------------------------------------

tableFicDatP1("DAAMJ") = recFicDatP1.DAAMJ
tableFicDatP1("DACTRJ") = recFicDatP1.DACTRJ
tableFicDatP1("DACTRA") = recFicDatP1.DACTRA
tableFicDatP1("DALIBJ") = recFicDatP1.DALIBJ
tableFicDatP1("DALIBM") = recFicDatP1.DALIBM
tableFicDatP1("DALBMR") = recFicDatP1.DALBMR
tableFicDatP1("DAFERJ") = recFicDatP1.DAFERJ
tableFicDatP1("DATSTC") = recFicDatP1.DATSTC
End Sub


'---------------------------------------------------------
Public Function tableFicDatP1_Read(recFicDatP1 As typeFicDatP1) As Integer
'---------------------------------------------------------

On Error GoTo tableFicDatP1_Read_Error
tableFicDatP1_Read = 0


Select Case Trim(recFicDatP1.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableFicDatP1.Seek "=", recFicDatP1.DAAMJ
                        If tableFicDatP1.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableFicDatP1.Seek "<=", recFicDatP1.DAAMJ
                        If tableFicDatP1.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableFicDatP1.Seek ">=", recFicDatP1.DAAMJ
                        If tableFicDatP1.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableFicDatP1.Seek ">", recFicDatP1.DAAMJ
                        If tableFicDatP1.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableFicDatP1.MoveNext
                        If tableFicDatP1.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableFicDatP1.MovePrevious
                        If tableFicDatP1.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableFicDatP1.MoveFirst
                        If tableFicDatP1.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableFicDatP1.MoveLast
                        If tableFicDatP1.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recFicDatP1.Method <> "AddNew      " Then
    Call tableFicDatP1_GetBuffer(recFicDatP1)
End If

Exit Function

'---------------------------------------------------------
tableFicDatP1_Read_Error:
'---------------------------------------------------------

    tableFicDatP1_Read = Err
    Resume tableFicDatP1_Read_End

tableFicDatP1_Read_End:

End Function

'---------------------------------------------------------
Public Function tableFicDatP1_Update(recFicDatP1 As typeFicDatP1) As Integer
'---------------------------------------------------------

On Error GoTo tableFicDatP1Update_Error
tableFicDatP1_Update = 0

Select Case Trim(recFicDatP1.Method)

    Case "AddNew"
                        tableFicDatP1.AddNew
                        Call tableFicDatP1_PutBuffer(recFicDatP1)
                        tableFicDatP1.Update
    Case "Update"
                        tableFicDatP1.Edit
                        Call tableFicDatP1_PutBuffer(recFicDatP1)
                        tableFicDatP1.Update
    Case "Delete"
                        tableFicDatP1.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableFicDatP1Update_Error:
'---------------------------------------------------------
    tableFicDatP1_Update = Err
    Resume tableFicDatP1Update_End

tableFicDatP1Update_End:

End Function

Public Function dbFicDatP1_Import(lFileName As String, lNb As Long)
Dim X As String, xInput As String
Dim recFicDatP1 As typeFicDatP1
On Error GoTo Error_Handler

Dim I As Integer, blnOk As Boolean

lNb = 0: I = 0
X = Dir(lFileName)
If X = "" Then dbFicDatP1_Import = "? dbFicDatP1_Import : Le fichier des dates n'existe pas": Exit Function


MDB.Execute "delete * from FicDatP1"
tableFicDatP1_Open
recFicDatP1_Init recFicDatP1
recFicDatP1.Method = "AddNew"

Open lFileName For Input As #1

blnOk = False
Do Until EOF(1)
    Line Input #1, xInput
    

    recFicDatP1.DAAMJ = Format$(mId$(xInput, 5, 2), "00") & Format$(mId$(xInput, 7, 2), "00") & Format$(mId$(xInput, 3, 2), "00") & Format$(mId$(xInput, 1, 2), "00")
    If recFicDatP1.DAAMJ > "20000000" And recFicDatP1.DAAMJ < "20100000" Then
        recFicDatP1.Method = "AddNew"
        recFicDatP1.DACTRJ = CLng(Val(mId$(xInput, 9, 5)))
        recFicDatP1.DACTRA = CLng(Val(mId$(xInput, 14, 3)))
        recFicDatP1.DALIBJ = mId$(xInput, 17, 8)
        recFicDatP1.DALIBM = mId$(xInput, 25, 9)
        recFicDatP1.DALBMR = mId$(xInput, 34, 4)
        recFicDatP1.DAFERJ = mId$(xInput, 38, 1)
        recFicDatP1.DATSTC = mId$(xInput, 39, 1)
        dbFicDatP1_Update recFicDatP1
        lNb = lNb + 1
    End If
Loop

Close
tableFicDatP1_Close


Exit Function
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------
    X = "? dbFicDatP1_Import : " & Err & " : " & Error(Err)
    Call MsgBox(X, vbCritical, "dbFicDatP1_Import")
    dbFicDatP1_Import = X: Exit Function

End Function








'-----------------------------------------------------
Sub dbFicDatP1_Error(recFicDatP1 As typeFicDatP1)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & recFicDatP1.DAAMJ & ": " & Chr$(13)

Select Case mId$(recFicDatP1.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recFicDatP1.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbFicDatP1.bas :  ( " & Trim(recFicDatP1.obj) & " : " & Trim(recFicDatP1.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbFicDatP1_ReadE(recFicDatP1 As typeFicDatP1)
'-----------------------------------------------------

dbFicDatP1_ReadE = Null

recFicDatP1.Err = tableFicDatP1_Read(recFicDatP1)
If recFicDatP1.Err > 0 Then

'    If recFicDatP1.Err < 9990 Or recFicDatP1.Err >= 9999 Then
        Call dbFicDatP1_Error(recFicDatP1)
        dbFicDatP1_ReadE = recFicDatP1.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbFicDatP1_Update(recFicDatP1 As typeFicDatP1)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbFicDatP1_Update = Null


recFicDatP1.Err = tableFicDatP1_Update(recFicDatP1)

If recFicDatP1.Err <> 0 Then
    Call dbFicDatP1_Error(recFicDatP1)
    dbFicDatP1_Update = recFicDatP1.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recFicDatP1_Init(recFicDatP1 As typeFicDatP1)
recFicDatP1.Method = ""
recFicDatP1.obj = "FicDatP1"
recFicDatP1.Err = ""
recFicDatP1.DAAMJ = ""
recFicDatP1.DACTRJ = 0
recFicDatP1.DACTRA = 0
recFicDatP1.DALIBJ = ""
recFicDatP1.DALIBM = ""
recFicDatP1.DALBMR = ""
recFicDatP1.DAFERJ = ""
recFicDatP1.DATSTC = ""
End Sub





Public Function dateBIA(ByVal Fct As String, ByVal Nb As Integer, ByVal X As String) As String
Dim K As Integer, K1 As Integer, K2 As Integer
Dim V, X8 As String * 8, X8B As String * 8
Dim Fct_Mod As String, Nb_Mod As Integer
Dim iReturn As Integer, blnOk As Boolean, mMethod As String

tableFicDatP1_Open
dateBIA = X
paramDateBIA.Method = "Seek>="
paramDateBIA.DAAMJ = X
Fct_Mod = Fct: Nb_Mod = Nb

If Nb_Mod >= 0 Then
    K1 = 1: mMethod = "MoveNext"
Else
    K1 = -1: mMethod = "MovePrevious"
End If

Select Case Fct
    Case "Ouvré"
        K = 0
        
        blnOk = False
        Do
            If tableFicDatP1_Read(paramDateBIA) = 0 Then
                If paramDateBIA.DAFERJ = " " Then
                    If blnOk = True Then K = K + K1
                    blnOk = True
                End If
                paramDateBIA.Method = mMethod
            Else
                Call MsgBox("dateBIA_Ouvré", vbCritical, "mdbFicDatp1")
                dateBIA = "?"
                Exit Function
            End If
       Loop Until K = Nb_Mod And blnOk = True
        dateBIA = paramDateBIA.DAAMJ
        
    Case "Jour"
        K = -K1
        Do
            If tableFicDatP1_Read(paramDateBIA) = 0 Then
                K = K + K1
                paramDateBIA.Method = mMethod
            Else
                Call MsgBox("dateBIA_Ouvré", vbCritical, "mdbFicDatp1")
                dateBIA = "?"
                Exit Function
            End If
       Loop Until K = Nb_Mod
        dateBIA = paramDateBIA.DAAMJ
End Select

End Function
