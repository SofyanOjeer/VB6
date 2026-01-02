Attribute VB_Name = "mdbLrTiers"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableLrTiers As Recordset
Dim tableLrTiersOpen As Boolean

Type typeLrTiers
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    CDDECL       As String * 5
    RFBENF       As String * 16
    NSIREN       As String * 9
    NOMBNF       As String * 60
    PRENOM       As String * 60
    CDSEXE       As String * 1
    DTNAIS       As String * 6
    CDPAYS1      As String * 3
    CDDEPT1      As String * 2
    CDCOMM1      As String * 3
    LBCOMM1      As String * 32
    NOMCJT       As String * 60
    CDACCO       As String * 5
    CTJURI       As String * 5
    CDRESI       As String * 1
    NOVOIE       As String * 32
    CDPOST       As String * 5
    LBCOMM2      As String * 27
    CDDEPT2      As String * 2
    CDPAYS2      As String * 3
    CDTRI1       As String * 16
    CDTRI2       As String * 16
    CDAGCO       As String * 5
    FILL02       As String * 11

End Type

Public recLrTiers As typeLrTiers


'---------------------------------------------------------
Public Function Import_LrTiers(MsgTxt As String, recLrTiers As typeLrTiers)
'---------------------------------------------------------
Dim K As Integer, I As Integer
Import_LrTiers = Null
recLrTiers.obj = "LrTiers"
recLrTiers.Method = ""
recLrTiers.Err = ""

recLrTiers.CDDECL = Mid$(MsgTxt, 1, 5)
recLrTiers.RFBENF = Mid$(MsgTxt, 6, 16)
recLrTiers.NSIREN = Mid$(MsgTxt, 22, 9)
recLrTiers.NOMBNF = Mid$(MsgTxt, 31, 60)
recLrTiers.PRENOM = Mid$(MsgTxt, 91, 60)
recLrTiers.CDSEXE = Mid$(MsgTxt, 151, 1)
recLrTiers.DTNAIS = Mid$(MsgTxt, 152, 6)
recLrTiers.CDPAYS1 = Mid$(MsgTxt, 158, 3)
recLrTiers.CDDEPT1 = Mid$(MsgTxt, 161, 2)
recLrTiers.CDCOMM1 = Mid$(MsgTxt, 163, 3)
recLrTiers.LBCOMM1 = Mid$(MsgTxt, 166, 32)
recLrTiers.NOMCJT = Mid$(MsgTxt, 198, 60)
recLrTiers.CDACCO = Mid$(MsgTxt, 258, 5)
recLrTiers.CTJURI = Mid$(MsgTxt, 263, 5)
recLrTiers.CDRESI = Mid$(MsgTxt, 268, 1)
recLrTiers.NOVOIE = Mid$(MsgTxt, 269, 32)
recLrTiers.CDPOST = Mid$(MsgTxt, 301, 5)
recLrTiers.LBCOMM2 = Mid$(MsgTxt, 306, 27)
recLrTiers.CDDEPT2 = Mid$(MsgTxt, 333, 2)
recLrTiers.CDPAYS2 = Mid$(MsgTxt, 335, 3)
recLrTiers.CDTRI1 = Mid$(MsgTxt, 338, 16)
recLrTiers.CDTRI2 = Mid$(MsgTxt, 354, 16)
recLrTiers.CDAGCO = Mid$(MsgTxt, 370, 5)
recLrTiers.FILL02 = Mid$(MsgTxt, 375, 11)

End Function


'---------------------------------------------------------
Public Function Export_LrTiers(MsgTxt As String, recLrTiers As typeLrTiers)
'---------------------------------------------------------
Dim K As Integer, I As Integer
Export_LrTiers = Null
Mid$(MsgTxt, 1, 5) = recLrTiers.CDDECL
Mid$(MsgTxt, 6, 16) = recLrTiers.RFBENF
Mid$(MsgTxt, 22, 9) = recLrTiers.NSIREN
Mid$(MsgTxt, 31, 60) = recLrTiers.NOMBNF
Mid$(MsgTxt, 91, 60) = recLrTiers.PRENOM
Mid$(MsgTxt, 151, 1) = recLrTiers.CDSEXE
Mid$(MsgTxt, 152, 6) = recLrTiers.DTNAIS
Mid$(MsgTxt, 158, 3) = recLrTiers.CDPAYS1
Mid$(MsgTxt, 161, 2) = recLrTiers.CDDEPT1
Mid$(MsgTxt, 163, 3) = recLrTiers.CDCOMM1
Mid$(MsgTxt, 166, 32) = recLrTiers.LBCOMM1
Mid$(MsgTxt, 198, 60) = recLrTiers.NOMCJT
Mid$(MsgTxt, 258, 5) = recLrTiers.CDACCO
Mid$(MsgTxt, 263, 5) = recLrTiers.CTJURI
Mid$(MsgTxt, 268, 1) = recLrTiers.CDRESI
Mid$(MsgTxt, 269, 32) = recLrTiers.NOVOIE
Mid$(MsgTxt, 301, 5) = recLrTiers.CDPOST
Mid$(MsgTxt, 306, 27) = recLrTiers.LBCOMM2
Mid$(MsgTxt, 333, 2) = recLrTiers.CDDEPT2
Mid$(MsgTxt, 335, 3) = recLrTiers.CDPAYS2
Mid$(MsgTxt, 338, 16) = recLrTiers.CDTRI1
Mid$(MsgTxt, 354, 16) = recLrTiers.CDTRI2
Mid$(MsgTxt, 370, 5) = recLrTiers.CDAGCO
Mid$(MsgTxt, 375, 11) = recLrTiers.FILL02

End Function



'---------------------------------------------------------
'-----------------------------------------------------
Sub tableLrTiers_Close()
'-----------------------------------------------------
If tableLrTiersOpen Then
    tableLrTiers.Close
    tableLrTiersOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableLrTiers_GetBuffer(recLrTiers As typeLrTiers)
'---------------------------------------------------------

recLrTiers.CDDECL = tableLrTiers("CDDECL")
recLrTiers.RFBENF = tableLrTiers("RFBENF")

recLrTiers.NSIREN = tableLrTiers("NSIREN")
recLrTiers.NOMBNF = tableLrTiers("NOMBNF")
recLrTiers.PRENOM = tableLrTiers("PRENOM")
recLrTiers.CDSEXE = tableLrTiers("CDSEXE")
recLrTiers.DTNAIS = tableLrTiers("DTNAIS")
recLrTiers.CDPAYS1 = tableLrTiers("CDPAYS1")
recLrTiers.CDDEPT1 = tableLrTiers("CDDEPT1")
recLrTiers.CDCOMM1 = tableLrTiers("CDCOMM1")
recLrTiers.LBCOMM1 = tableLrTiers("LBCOMM1")
recLrTiers.NOMCJT = tableLrTiers("NOMCJT")
recLrTiers.CDACCO = tableLrTiers("CDACCO")
recLrTiers.CTJURI = tableLrTiers("CTJURI")
recLrTiers.CDRESI = tableLrTiers("CDRESI")
recLrTiers.NOVOIE = tableLrTiers("NOVOIE")
recLrTiers.CDPOST = tableLrTiers("CDPOST")
recLrTiers.LBCOMM2 = tableLrTiers("LBCOMM2")
recLrTiers.CDDEPT2 = tableLrTiers("CDDEPT2")
recLrTiers.CDPAYS2 = tableLrTiers("CDPAYS2")
recLrTiers.CDTRI1 = tableLrTiers("CDTRI1")
recLrTiers.CDTRI2 = tableLrTiers("CDTRI2")
recLrTiers.CDAGCO = tableLrTiers("CDAGCO")
recLrTiers.FILL02 = tableLrTiers("FILL02")

End Sub


'-----------------------------------------------------
Sub tableLrTiers_Open()
'-----------------------------------------------------

If Not tableLrTiersOpen Then
    Set tableLrTiers = MDB.OpenRecordset("LrTiers")
    tableLrTiers.Index = "PrimaryKey"
    tableLrTiersOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableLrTiers_PutBuffer(recLrTiers As typeLrTiers)
'---------------------------------------------------------

tableLrTiers("CDDECL") = recLrTiers.CDDECL
tableLrTiers("RFBENF") = recLrTiers.RFBENF
tableLrTiers("NSIREN") = recLrTiers.NSIREN
tableLrTiers("NOMBNF") = recLrTiers.NOMBNF
tableLrTiers("PRENOM") = recLrTiers.PRENOM
tableLrTiers("CDSEXE") = recLrTiers.CDSEXE
tableLrTiers("DTNAIS") = recLrTiers.DTNAIS
tableLrTiers("CDPAYS1") = recLrTiers.CDPAYS1
tableLrTiers("CDDEPT1") = recLrTiers.CDDEPT1
tableLrTiers("CDCOMM1") = recLrTiers.CDCOMM1
tableLrTiers("LBCOMM1") = recLrTiers.LBCOMM1
tableLrTiers("NOMCJT") = recLrTiers.NOMCJT
tableLrTiers("CDACCO") = recLrTiers.CDACCO
tableLrTiers("CTJURI") = recLrTiers.CTJURI
tableLrTiers("CDRESI") = recLrTiers.CDRESI
tableLrTiers("NOVOIE") = recLrTiers.NOVOIE
tableLrTiers("CDPOST") = recLrTiers.CDPOST
tableLrTiers("LBCOMM2") = recLrTiers.LBCOMM2
tableLrTiers("CDDEPT2") = recLrTiers.CDDEPT2
tableLrTiers("CDPAYS2") = recLrTiers.CDPAYS2
tableLrTiers("CDTRI1") = recLrTiers.CDTRI1
tableLrTiers("CDTRI2") = recLrTiers.CDTRI2
tableLrTiers("CDAGCO") = recLrTiers.CDAGCO
tableLrTiers("FILL02") = recLrTiers.FILL02
End Sub


'---------------------------------------------------------
Public Function tableLrTiers_Read(recLrTiers As typeLrTiers) As Integer
'---------------------------------------------------------

On Error GoTo tableLrTiers_Read_Error
tableLrTiers_Read = 0


Select Case recLrTiers.Method
     Case "Seek=       "
                        tableLrTiers.Seek "=", recLrTiers.FILL02, recLrTiers.RFBENF
                        If tableLrTiers.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<=      "
                        tableLrTiers.Seek "<=", recLrTiers.FILL02, recLrTiers.RFBENF
                        If tableLrTiers.NoMatch Then
                            Error 9998
                        End If
     Case "MoveNext    "
                        tableLrTiers.MoveNext
                        If tableLrTiers.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableLrTiers.MovePrevious
                        If tableLrTiers.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst   "
                        tableLrTiers.MoveFirst
                        If tableLrTiers.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast    "
                        tableLrTiers.MoveLast
                        If tableLrTiers.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recLrTiers.Method <> "AddNew      " Then
    Call tableLrTiers_GetBuffer(recLrTiers)
End If

Exit Function

'---------------------------------------------------------
tableLrTiers_Read_Error:
'---------------------------------------------------------

    tableLrTiers_Read = Err
    Resume tableLrTiers_Read_End

tableLrTiers_Read_End:

End Function

'---------------------------------------------------------
Public Function tableLrTiers_Update(recLrTiers As typeLrTiers) As Integer
'---------------------------------------------------------

On Error GoTo tableLrTiersUpdate_Error
tableLrTiers_Update = 0

Select Case recLrTiers.Method

    Case "AddNew      "
                        tableLrTiers.AddNew
                        Call tableLrTiers_PutBuffer(recLrTiers)
                        tableLrTiers.Update
    Case "Update      "
                        tableLrTiers.Edit
                        Call tableLrTiers_PutBuffer(recLrTiers)
                        tableLrTiers.Update
    Case "Delete      "
                        tableLrTiers.Delete
    Case Else
                        Error 9999
End Select


Exit Function

tableLrTiersUpdate_Error:
'---------------------------------------------------------
    tableLrTiers_Update = Err
    Resume tableLrTiersUpdate_End

tableLrTiersUpdate_End:

End Function








'-----------------------------------------------------
Sub dbLrTiers_Error(recLrTiers As typeLrTiers)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & Trim(recLrTiers.RFBENF) & Chr$(13)

Select Case Mid$(recLrTiers.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recLrTiers.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbLrTiers.bas :  ( " & Trim(recLrTiers.obj) & " : " & Trim(recLrTiers.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbLrTiers_Read(recLrTiers As typeLrTiers)
'-----------------------------------------------------

dbLrTiers_Read = Null

recLrTiers.Err = tableLrTiers_Read(recLrTiers)
If recLrTiers.Err > 0 Then

    If recLrTiers.Err < 9990 Or recLrTiers.Err >= 9999 Then
        Call dbLrTiers_Error(recLrTiers)
        dbLrTiers_Read = recLrTiers.Err
    End If
End If

End Function

'-----------------------------------------------------
Function dbLrTiers_ReadZ(recLrTiers As typeLrTiers)
'-----------------------------------------------------

dbLrTiers_ReadZ = Null

recLrTiers.Err = tableLrTiers_Read(recLrTiers)
If recLrTiers.Err > 0 Then
    dbLrTiers_ReadZ = recLrTiers.Err
    recLrTiers.NOMBNF = "? " & recLrTiers.RFBENF
End If

End Function


'-----------------------------------------------------
Function dbLrTiers_Update(recLrTiers As typeLrTiers)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
'$$$BeginTrans

dbLrTiers_Update = Null


recLrTiers.Err = tableLrTiers_Update(recLrTiers)

If recLrTiers.Err <> 0 Then
    Call dbLrTiers_Error(recLrTiers)
    dbLrTiers_Update = recLrTiers.Err
'$$$    Rollback
    Exit Function
End If

'$$$CommitTrans


'=====================================================
End Function


