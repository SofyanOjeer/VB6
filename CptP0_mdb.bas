Attribute VB_Name = "mdbCptP0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableCptP0 As Recordset
Dim tableCptP0Open As Boolean
Public mCptP0_Id As Long

Type typeCptP0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Id                  As String * 34
    Text                As String

End Type

Public reccptp0 As typeCptP0

'---------------------------------------------------------
'-----------------------------------------------------
Sub tableCptP0_Close()
'-----------------------------------------------------
If tableCptP0Open Then
    tableCptP0.Close
    tableCptP0Open = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableCptP0_GetBuffer(reccptp0 As typeCptP0)
'---------------------------------------------------------
reccptp0.Id = tableCptP0("Id")
reccptp0.Text = tableCptP0("Text")

End Sub


'-----------------------------------------------------
Sub tableCptP0_Open()
'-----------------------------------------------------

If Not tableCptP0Open Then
    Set tableCptP0 = MDB.OpenRecordset("CptP0")
    tableCptP0.Index = "PrimaryKey"
    tableCptP0Open = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableCptP0_PutBuffer(reccptp0 As typeCptP0)
'---------------------------------------------------------

tableCptP0("Id") = reccptp0.Id
tableCptP0("Text") = reccptp0.Text
End Sub


'---------------------------------------------------------
Public Function tableCptP0_Read(reccptp0 As typeCptP0) As Integer
'---------------------------------------------------------

On Error GoTo tableCptP0_Read_Error
tableCptP0_Read = 0


Select Case Trim(reccptp0.Method)
     Case "Seek=", "AddNew", "Update", "Delete"

                        tableCptP0.Seek "=", reccptp0.Id
                        If tableCptP0.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<="
                        tableCptP0.Seek "<=", reccptp0.Id
                        If tableCptP0.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>="
                        tableCptP0.Seek ">=", reccptp0.Id
                        If tableCptP0.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>"
                        tableCptP0.Seek ">", reccptp0.Id
                        If tableCptP0.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext"
                        tableCptP0.MoveNext
                        If tableCptP0.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableCptP0.MovePrevious
                        If tableCptP0.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst"
                        tableCptP0.MoveFirst
                        If tableCptP0.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast"
                        tableCptP0.MoveLast
                        If tableCptP0.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If reccptp0.Method <> "AddNew      " Then
    Call tableCptP0_GetBuffer(reccptp0)
End If

Exit Function

'---------------------------------------------------------
tableCptP0_Read_Error:
'---------------------------------------------------------

    tableCptP0_Read = Err
    Resume tableCptP0_Read_End

tableCptP0_Read_End:

End Function
'-----------------------------------------------------
Public Function mdbCptP0_Find(recCompte As typeCompte)
'-----------------------------------------------------
Dim r As Integer, wCV As typeCV
Dim V

tableCptP0_Open

If Not IsNumeric(recCompte.Devise) Then
    Call CV_AttributS(recCompte.Devise, wCV)
    recCompte.Devise = wCV.DeviseN
End If

reccptp0.Id = SocId$ & SocAgence$ & recCompte.Devise & recCompte.Numéro
reccptp0.Method = "Seek="
r = tableCptP0_Read(reccptp0)
If r = 0 Then
    MsgTxtIndex = 0
    MsgTxt = Space$(34) & mId$(reccptp0.Text, 1, recCompteLen - 34)
    V = srvCompteGetBuffer(recCompte)
Else
    recCompte.Société = SocId$
    recCompte.Agence = SocAgence$
    recCompte.BiaTyp = "000"
    recCompte.BiaNum = "00"
    recCompte.Method = "SeekL1"
    V = srvCompte_InitFind(recCompte)
End If

mdbCptP0_Find = V
If Not IsNull(V) Then
    recCompteInit recCompte
End If

End Function


'-----------------------------------------------------
Public Function mdbCptInfoP0_Find(recCptInfo As typeCptInfo)
'-----------------------------------------------------
Dim r As Integer
tableCptP0_Open
reccptp0.Id = recCptInfo.Société & recCptInfo.Agence & recCptInfo.Devise & recCptInfo.Numéro
reccptp0.Method = "Seek="
r = tableCptP0_Read(reccptp0)
If r = 0 Then
    MsgTxtIndex = 0
    MsgTxt = Space$(34) & mId$(reccptp0.Text, 1, recCptInfoLen - 34)
    mdbCptInfoP0_Find = srvCptInfoGetBuffer(recCptInfo)
Else
    mdbCptInfoP0_Find = srvCptInfoFind(recCptInfo)
End If


End Function


'---------------------------------------------------------
Public Function tableCptP0_Update(reccptp0 As typeCptP0) As Integer
'---------------------------------------------------------

On Error GoTo tableCptP0Update_Error
tableCptP0_Update = 0

Select Case Trim(reccptp0.Method)

    Case "AddNew"
                        tableCptP0.AddNew
                        Call tableCptP0_PutBuffer(reccptp0)
                        tableCptP0.Update
    Case "Update"
                        tableCptP0.Edit
                        Call tableCptP0_PutBuffer(reccptp0)
                        tableCptP0.Update
    Case "Delete"
                        tableCptP0.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableCptP0Update_Error:
'---------------------------------------------------------
    tableCptP0_Update = Err
    Resume tableCptP0Update_End

tableCptP0Update_End:

End Function








'-----------------------------------------------------
Sub dbCptP0_Error(reccptp0 As typeCptP0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & reccptp0.Id & ": " & Chr$(13)

Select Case mId$(reccptp0.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23", "98": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & reccptp0.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbCptP0.bas :  ( " & Trim(reccptp0.obj) & " : " & Trim(reccptp0.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbCptP0_ReadE(reccptp0 As typeCptP0)
'-----------------------------------------------------

dbCptP0_ReadE = Null

reccptp0.Err = tableCptP0_Read(reccptp0)
If reccptp0.Err > 0 Then

'    If recCptP0.Err < 9990 Or recCptP0.Err >= 9999 Then
        Call dbCptP0_Error(reccptp0)
        dbCptP0_ReadE = reccptp0.Err
'    End If
End If

End Function

'-----------------------------------------------------
Function dbCptP0_Update(reccptp0 As typeCptP0)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
BeginTrans

dbCptP0_Update = Null


reccptp0.Err = tableCptP0_Update(reccptp0)

If reccptp0.Err <> 0 Then
    Call dbCptP0_Error(reccptp0)
    dbCptP0_Update = reccptp0.Err
    Rollback
    Exit Function
End If

CommitTrans


'=====================================================
End Function



Public Sub recCptP0_Init(reccptp0 As typeCptP0)
reccptp0.Method = ""
reccptp0.obj = "CptP0"
reccptp0.Err = ""
reccptp0.Id = ""
reccptp0.Text = ""
End Sub



Public Function mdbCptP0_Sel(lCompteMin As typeCompte, lCompteMax As typeCompte, Fct As String)
Dim X11 As String * 11, X3 As String * 3, x As String
Dim recCompte As typeCompte
mdbCptP0_Sel = Null '"?"
If selCompte_On Then
    Select Case Fct
        Case "Add":
        Case "End": selCompte_On = False: Exit Function
        Case Else
            Call MsgBox("sélection de comptes en cours d'utilisation par un autre processus", vbCritical, "SrvCompte.selCompte_Load")
            Exit Function
    End Select
End If

selCompte_On = True
If Fct <> "Add" Then selCompte_Nb = 0: ReDim selCompte(20): selCompte_NbMax = 20

mdbCptP0.tableCptP0_Open
Mid$(MsgTxt, 1, 34) = Space$(34)
reccptp0.Method = "MoveFirst"

Call dbCptP0_ReadE(reccptp0)

Do While reccptp0.Err = 0
    
    X11 = mId$(reccptp0.Id, 10, 11)
    If X11 >= lCompteMin.Numéro And X11 <= lCompteMax.Numéro Then
        X3 = mId$(reccptp0.Id, 7, 3)
        If X3 >= lCompteMin.Devise And X3 <= lCompteMax.Devise Then
   
            MsgTxtIndex = 0
            Mid$(MsgTxt, 35, memoCptInfoLen) = mId$(reccptp0.Text, 1, memoCptInfoLen)
            If IsNull(srvCompteGetBuffer(recCompte)) Then
                If selCompte_Nb >= selCompte_NbMax Then
                    selCompte_NbMax = selCompte_NbMax + 12
                    ReDim Preserve selCompte(selCompte_NbMax)
                End If
                selCompte_Nb = selCompte_Nb + 1
                selCompte(selCompte_Nb) = recCompte
            End If
        End If
    End If
    reccptp0.Method = "MoveNext    "
    reccptp0.Err = tableCptP0_Read(reccptp0)
Loop


End Function
