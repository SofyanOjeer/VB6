Attribute VB_Name = "mdbDGI_2561"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableDGI_2561 As Recordset
Dim tableDGI_2561Open As Boolean

Type typeDGI_2561
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Id       As String * 5
    ZC       As String * 40
    ZD       As String * 40
    ZG       As String * 40
    ZH       As String * 40
    ZI       As String * 40
    ZJ       As String * 40
    AI      As String * 20
    AH      As String * 20
    BR      As String * 20
    AC      As String * 20
    AE       As String * 20
    AF       As String * 20
    AO       As String * 20
    CT       As String * 20
    AR       As Currency
    BN       As Currency
    BP      As Currency
End Type

Public recDGI_2561 As typeDGI_2561

Public paramDGI_2561_Filename As String
Public paramDGI_IFUTR141P1 As String

'---------------------------------------------------------
Public Function Import_DGI_2561(MsgTxt As String, recDGI_2561 As typeDGI_2561)
'---------------------------------------------------------
Dim K As Integer, I As Integer
Import_DGI_2561 = Null
recDGI_2561.obj = "DGI_2561"
recDGI_2561.Method = ""
recDGI_2561.Err = ""

recDGI_2561.Id = mId$(MsgTxt, 1, 5)
recDGI_2561.ZC = mId$(MsgTxt, 6, 40)
 recDGI_2561.ZD = mId$(MsgTxt, 46, 40)
 recDGI_2561.ZG = mId$(MsgTxt, 86, 40)
 recDGI_2561.ZH = mId$(MsgTxt, 126, 40)
 recDGI_2561.ZI = mId$(MsgTxt, 166, 40)
 recDGI_2561.ZJ = mId$(MsgTxt, 206, 40)
 recDGI_2561.AI = mId$(MsgTxt, 246, 20)
 recDGI_2561.AH = mId$(MsgTxt, 266, 20)
 recDGI_2561.BR = mId$(MsgTxt, 286, 20)
 recDGI_2561.AC = mId$(MsgTxt, 306, 20)
 recDGI_2561.AE = mId$(MsgTxt, 326, 20)
 recDGI_2561.AF = mId$(MsgTxt, 346, 20)
 recDGI_2561.AO = mId$(MsgTxt, 366, 20)
 recDGI_2561.CT = mId$(MsgTxt, 386, 20)
recDGI_2561.AR = CCur(Val(mId$(MsgTxt, 406, 16)) / 100)
recDGI_2561.BN = CCur(Val(mId$(MsgTxt, 422, 16)) / 100)
recDGI_2561.BP = CCur(Val(mId$(MsgTxt, 438, 16)) / 100)

End Function


'---------------------------------------------------------
Public Function Export_DGI_2561(MsgTxt As String, recDGI_2561 As typeDGI_2561)
'---------------------------------------------------------
Dim K As Integer, I As Integer
Export_DGI_2561 = Null
Mid$(MsgTxt, 1, 5) = recDGI_2561.Id
Mid$(MsgTxt, 6, 40) = recDGI_2561.ZC
Mid$(MsgTxt, 46, 40) = recDGI_2561.ZD
Mid$(MsgTxt, 86, 40) = recDGI_2561.ZG
Mid$(MsgTxt, 126, 40) = recDGI_2561.ZH
Mid$(MsgTxt, 166, 40) = recDGI_2561.ZI
Mid$(MsgTxt, 206, 40) = recDGI_2561.ZJ
Mid$(MsgTxt, 246, 20) = recDGI_2561.AI
Mid$(MsgTxt, 266, 20) = recDGI_2561.AH
Mid$(MsgTxt, 286, 20) = recDGI_2561.BR
Mid$(MsgTxt, 306, 20) = recDGI_2561.AC
Mid$(MsgTxt, 326, 20) = recDGI_2561.AE
Mid$(MsgTxt, 346, 20) = recDGI_2561.AF
Mid$(MsgTxt, 366, 20) = recDGI_2561.AO
Mid$(MsgTxt, 386, 20) = recDGI_2561.CT
Mid$(MsgTxt, 406, 16) = Format$(recDGI_2561.AR * 100, "0000000000000000")
Mid$(MsgTxt, 422, 16) = Format$(recDGI_2561.BN * 100, "0000000000000000")
Mid$(MsgTxt, 438, 16) = Format$(recDGI_2561.BP * 100, "0000000000000000")
End Function



'---------------------------------------------------------
'-----------------------------------------------------
Sub tableDGI_2561_Close()
'-----------------------------------------------------
If tableDGI_2561Open Then
    tableDGI_2561.Close
    tableDGI_2561Open = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableDGI_2561_GetBuffer(recDGI_2561 As typeDGI_2561)
'---------------------------------------------------------

recDGI_2561.Id = tableDGI_2561("Id")
recDGI_2561.ZC = tableDGI_2561("ZC")

recDGI_2561.ZD = tableDGI_2561("ZD")
recDGI_2561.ZG = tableDGI_2561("ZG")
recDGI_2561.ZH = tableDGI_2561("ZH")
recDGI_2561.ZI = tableDGI_2561("ZI")
recDGI_2561.ZJ = tableDGI_2561("ZJ")
recDGI_2561.AI = tableDGI_2561("AI")
recDGI_2561.AH = tableDGI_2561("AH")
recDGI_2561.BR = tableDGI_2561("BR")
recDGI_2561.AC = tableDGI_2561("AC")
recDGI_2561.AE = tableDGI_2561("AE")
recDGI_2561.AF = tableDGI_2561("AF")
recDGI_2561.AO = tableDGI_2561("AO")
recDGI_2561.CT = tableDGI_2561("CT")
recDGI_2561.AR = tableDGI_2561("AR")
recDGI_2561.BN = tableDGI_2561("BN")
recDGI_2561.BP = tableDGI_2561("BP")

End Sub


'-----------------------------------------------------
Sub tableDGI_2561_Open()
'-----------------------------------------------------

If Not tableDGI_2561Open Then
    Set tableDGI_2561 = MDB.OpenRecordset("DGI_2561")
    tableDGI_2561.Index = "PrimaryKey"
    tableDGI_2561Open = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableDGI_2561_PutBuffer(recDGI_2561 As typeDGI_2561)
'---------------------------------------------------------

tableDGI_2561("Id") = recDGI_2561.Id
tableDGI_2561("ZC") = recDGI_2561.ZC
tableDGI_2561("ZD") = recDGI_2561.ZD
tableDGI_2561("ZG") = recDGI_2561.ZG
tableDGI_2561("ZH") = recDGI_2561.ZH
tableDGI_2561("ZI") = recDGI_2561.ZI
tableDGI_2561("ZJ") = recDGI_2561.ZJ
tableDGI_2561("AI") = recDGI_2561.AI
tableDGI_2561("AH") = recDGI_2561.AH
tableDGI_2561("BR") = recDGI_2561.BR
tableDGI_2561("AC") = recDGI_2561.AC
tableDGI_2561("AE") = recDGI_2561.AE
tableDGI_2561("AF") = recDGI_2561.AF
tableDGI_2561("AO") = recDGI_2561.AO
tableDGI_2561("CT") = recDGI_2561.CT
tableDGI_2561("AR") = recDGI_2561.AR
tableDGI_2561("BN") = recDGI_2561.BN
tableDGI_2561("BP") = recDGI_2561.BP
End Sub


'---------------------------------------------------------
Public Function tableDGI_2561_Read(recDGI_2561 As typeDGI_2561) As Integer
'---------------------------------------------------------

On Error GoTo tableDGI_2561_Read_Error
tableDGI_2561_Read = 0


Select Case recDGI_2561.Method
     Case "Seek=       "
                        tableDGI_2561.Seek "=", recDGI_2561.Id
                        If tableDGI_2561.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<=      "
                        tableDGI_2561.Seek "<=", recDGI_2561.Id
                        If tableDGI_2561.NoMatch Then
                            Error 9998
                        End If
     Case "MoveNext    "
                        tableDGI_2561.MoveNext
                        If tableDGI_2561.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableDGI_2561.MovePrevious
                        If tableDGI_2561.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst   "
                        tableDGI_2561.MoveFirst
                        If tableDGI_2561.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast    "
                        tableDGI_2561.MoveLast
                        If tableDGI_2561.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recDGI_2561.Method <> "AddNew      " Then
    Call tableDGI_2561_GetBuffer(recDGI_2561)
End If

Exit Function

'---------------------------------------------------------
tableDGI_2561_Read_Error:
'---------------------------------------------------------

    tableDGI_2561_Read = Err
    Resume tableDGI_2561_Read_End

tableDGI_2561_Read_End:

End Function

'---------------------------------------------------------
Public Function tableDGI_2561_Update(recDGI_2561 As typeDGI_2561) As Integer
'---------------------------------------------------------

On Error GoTo tableDGI_2561Update_Error
tableDGI_2561_Update = 0

Select Case recDGI_2561.Method

    Case "AddNew      "
                        tableDGI_2561.AddNew
                        Call tableDGI_2561_PutBuffer(recDGI_2561)
                        tableDGI_2561.Update
    Case "Update      "
                        tableDGI_2561.Edit
                        Call tableDGI_2561_PutBuffer(recDGI_2561)
                        tableDGI_2561.Update
    Case "Delete      "
                        tableDGI_2561.Delete
    Case Else
                        Error 9999
End Select


Exit Function

tableDGI_2561Update_Error:
'---------------------------------------------------------
    tableDGI_2561_Update = Err
    Resume tableDGI_2561Update_End

tableDGI_2561Update_End:

End Function








'-----------------------------------------------------
Sub dbDGI_2561_Error(recDGI_2561 As typeDGI_2561)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & Trim(recDGI_2561.ZC) & Chr$(13)

Select Case mId$(recDGI_2561.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recDGI_2561.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbDGI_2561.bas :  ( " & Trim(recDGI_2561.obj) & " : " & Trim(recDGI_2561.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbDGI_2561_Read(recDGI_2561 As typeDGI_2561)
'-----------------------------------------------------

dbDGI_2561_Read = Null

recDGI_2561.Err = tableDGI_2561_Read(recDGI_2561)
If recDGI_2561.Err > 0 Then

    If recDGI_2561.Err < 9990 Or recDGI_2561.Err >= 9999 Then
        Call dbDGI_2561_Error(recDGI_2561)
        dbDGI_2561_Read = recDGI_2561.Err
    End If
End If

End Function

'-----------------------------------------------------
Function dbDGI_2561_ReadZ(recDGI_2561 As typeDGI_2561)
'-----------------------------------------------------

dbDGI_2561_ReadZ = Null

recDGI_2561.Err = tableDGI_2561_Read(recDGI_2561)
If recDGI_2561.Err > 0 Then
    dbDGI_2561_ReadZ = recDGI_2561.Err
    recDGI_2561.ZG = "? " & recDGI_2561.ZC
End If

End Function


'-----------------------------------------------------
Function dbDGI_2561_Update(recDGI_2561 As typeDGI_2561)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
'$$$BeginTrans

dbDGI_2561_Update = Null


recDGI_2561.Err = tableDGI_2561_Update(recDGI_2561)

If recDGI_2561.Err <> 0 Then
    Call dbDGI_2561_Error(recDGI_2561)
    dbDGI_2561_Update = recDGI_2561.Err
'$$$    Rollback
    Exit Function
End If

'$$$CommitTrans


'=====================================================
End Function




Public Sub recDGI_25611_Init(lDGI_2561 As typeDGI_2561)

lDGI_2561.Method = ""
lDGI_2561.obj = "DGI_2561"
lDGI_2561.Err = ""

lDGI_2561.Id = ""
lDGI_2561.ZC = ""
lDGI_2561.ZD = ""
lDGI_2561.ZG = ""
lDGI_2561.ZH = ""
lDGI_2561.ZI = ""
lDGI_2561.ZJ = ""
lDGI_2561.AI = ""
lDGI_2561.AH = ""
lDGI_2561.BR = ""
lDGI_2561.AC = ""
lDGI_2561.AE = ""
lDGI_2561.AF = ""
lDGI_2561.AO = ""
lDGI_2561.CT = ""
lDGI_2561.AR = 0
lDGI_2561.BN = 0
lDGI_2561.BP = 0

End Sub
