Attribute VB_Name = "mdbLrSgnBnf"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableLrSgnBnf As Recordset
Dim tableLrSgnBnfOpen As Boolean

Type typeLrSgnBnf
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    CDBANQ       As String * 5
    CDDECL       As String * 5
    RFBENF       As String * 16
    NSIREN       As String * 9
    NPREFI1      As String * 2
    NSIREN1      As String * 9
    NSUFFI1      As String * 2
    AMJ1         As String * 8
    NPREFI2      As String * 2
    NSIREN2      As String * 9
    NSUFFI2      As String * 2
    AMJ2         As String * 8
    NOMBNF       As String * 60
    PRENOM       As String * 60
    CDSEXE       As String * 1
    JMA3         As String * 6
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
    LBCOMM2      As String * 32
    CDDEPT2      As String * 2
    CDPAYS2      As String * 3
    CDTRI1       As String * 16
    CDTRI2       As String * 16
    CDHABI       As String * 10
    AMJ4         As String * 8
    HMSC         As String * 8
    CDAGCO       As String * 5
    CDPHMO       As String * 1
    CDCRMD       As String * 1
    INSSIR       As String * 1
    FILL01       As String * 11
    CDRESI1      As String * 1
    CDACEN1      As String * 4
    CTJURN1      As String * 4
    CDPAYN1      As String * 2
    CDSEXE1      As String * 1
    CDPAYN2      As String * 2
    CDRESI2      As String * 1
    CDACCO2      As String * 5
    FILL02       As String * 14

End Type

Public recLrSgnBnf As typeLrSgnBnf


'---------------------------------------------------------
Public Function Import_LrSgnBnf(MsgTxt As String, recLrSgnBnf As typeLrSgnBnf)
'---------------------------------------------------------
Dim K As Integer, I As Integer
Import_LrSgnBnf = Null
recLrSgnBnf.obj = "LrSgnBnf"
recLrSgnBnf.Method = ""
recLrSgnBnf.Err = ""

recLrSgnBnf.CDBANQ = Mid$(MsgTxt, 1, 5)
recLrSgnBnf.CDDECL = Mid$(MsgTxt, 6, 5)
recLrSgnBnf.RFBENF = Mid$(MsgTxt, 11, 16)
recLrSgnBnf.NSIREN = Mid$(MsgTxt, 27, 9)
recLrSgnBnf.NPREFI1 = Mid$(MsgTxt, 36, 2)
recLrSgnBnf.NSIREN1 = Mid$(MsgTxt, 38, 9)
recLrSgnBnf.NSUFFI1 = Mid$(MsgTxt, 47, 2)
recLrSgnBnf.AMJ1 = Mid$(MsgTxt, 49, 8)
recLrSgnBnf.NPREFI2 = Mid$(MsgTxt, 57, 2)
recLrSgnBnf.NSIREN2 = Mid$(MsgTxt, 59, 9)
recLrSgnBnf.NSUFFI2 = Mid$(MsgTxt, 68, 2)
recLrSgnBnf.AMJ2 = Mid$(MsgTxt, 70, 8)
recLrSgnBnf.NOMBNF = Mid$(MsgTxt, 78, 60)
recLrSgnBnf.PRENOM = Mid$(MsgTxt, 138, 60)
recLrSgnBnf.CDSEXE = Mid$(MsgTxt, 198, 1)
recLrSgnBnf.JMA3 = Mid$(MsgTxt, 199, 6)
recLrSgnBnf.CDPAYS1 = Mid$(MsgTxt, 201, 3)
recLrSgnBnf.CDDEPT1 = Mid$(MsgTxt, 208, 2)
recLrSgnBnf.CDCOMM1 = Mid$(MsgTxt, 210, 3)
recLrSgnBnf.LBCOMM1 = Mid$(MsgTxt, 213, 32)
recLrSgnBnf.NOMCJT = Mid$(MsgTxt, 245, 60)
recLrSgnBnf.CDACCO = Mid$(MsgTxt, 305, 5)
recLrSgnBnf.CTJURI = Mid$(MsgTxt, 310, 5)
recLrSgnBnf.CDRESI = Mid$(MsgTxt, 315, 1)
recLrSgnBnf.NOVOIE = Mid$(MsgTxt, 316, 32)
recLrSgnBnf.CDPOST = Mid$(MsgTxt, 348, 5)
recLrSgnBnf.LBCOMM2 = Mid$(MsgTxt, 353, 32)
recLrSgnBnf.CDDEPT2 = Mid$(MsgTxt, 385, 2)
recLrSgnBnf.CDPAYS2 = Mid$(MsgTxt, 387, 3)
recLrSgnBnf.CDTRI1 = Mid$(MsgTxt, 390, 16)
recLrSgnBnf.CDTRI2 = Mid$(MsgTxt, 406, 16)
recLrSgnBnf.CDHABI = Mid$(MsgTxt, 422, 10)
recLrSgnBnf.AMJ4 = Mid$(MsgTxt, 432, 8)
recLrSgnBnf.HMSC = Mid$(MsgTxt, 440, 8)
recLrSgnBnf.CDAGCO = Mid$(MsgTxt, 448, 5)
recLrSgnBnf.CDPHMO = Mid$(MsgTxt, 453, 1)
recLrSgnBnf.CDCRMD = Mid$(MsgTxt, 454, 1)
recLrSgnBnf.INSSIR = Mid$(MsgTxt, 455, 1)
recLrSgnBnf.FILL01 = Mid$(MsgTxt, 456, 11)
recLrSgnBnf.CDRESI1 = Mid$(MsgTxt, 467, 1)
recLrSgnBnf.CDACEN1 = Mid$(MsgTxt, 468, 4)
recLrSgnBnf.CTJURN1 = Mid$(MsgTxt, 472, 4)
recLrSgnBnf.CDPAYN1 = Mid$(MsgTxt, 476, 2)
recLrSgnBnf.CDSEXE1 = Mid$(MsgTxt, 478, 1)
recLrSgnBnf.CDPAYN2 = Mid$(MsgTxt, 479, 2)
recLrSgnBnf.CDRESI2 = Mid$(MsgTxt, 481, 1)
recLrSgnBnf.CDACCO2 = Mid$(MsgTxt, 482, 5)
recLrSgnBnf.FILL02 = Mid$(MsgTxt, 487, 14)

End Function


'---------------------------------------------------------
'-----------------------------------------------------
Sub tableLrSgnBnf_Close()
'-----------------------------------------------------
If tableLrSgnBnfOpen Then
    tableLrSgnBnf.Close
    tableLrSgnBnfOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableLrSgnBnf_GetBuffer(recLrSgnBnf As typeLrSgnBnf)
'---------------------------------------------------------

recLrSgnBnf.CDBANQ = tableLrSgnBnf("CDBANQ")
recLrSgnBnf.CDDECL = tableLrSgnBnf("CDDECL")
recLrSgnBnf.RFBENF = tableLrSgnBnf("RFBENF")

recLrSgnBnf.NSIREN = tableLrSgnBnf("NSIREN")
recLrSgnBnf.NPREFI1 = tableLrSgnBnf("NPREFI1")
recLrSgnBnf.NSUFFI1 = tableLrSgnBnf("NSUFFI1")
recLrSgnBnf.AMJ1 = tableLrSgnBnf("AMJ1")
recLrSgnBnf.NPREFI2 = tableLrSgnBnf("NPREFI2")
recLrSgnBnf.NSIREN2 = tableLrSgnBnf("NSIREN2")
recLrSgnBnf.NSUFFI2 = tableLrSgnBnf("NSUFFI2")
recLrSgnBnf.AMJ2 = tableLrSgnBnf("AMJ2")
recLrSgnBnf.NOMBNF = tableLrSgnBnf("NOMBNF")
recLrSgnBnf.PRENOM = tableLrSgnBnf("PRENOM")
recLrSgnBnf.CDSEXE = tableLrSgnBnf("CDSEXE")
recLrSgnBnf.JMA3 = tableLrSgnBnf("JMA3")
recLrSgnBnf.CDPAYS1 = tableLrSgnBnf("CDPAYS1")
recLrSgnBnf.CDDEPT1 = tableLrSgnBnf("CDDEPT1")
recLrSgnBnf.CDCOMM1 = tableLrSgnBnf("CDCOMM1")
recLrSgnBnf.LBCOMM1 = tableLrSgnBnf("LBCOMM1")
recLrSgnBnf.NOMCJT = tableLrSgnBnf("NOMCJT")
recLrSgnBnf.CDACCO = tableLrSgnBnf("CDACCO")
recLrSgnBnf.CTJURI = tableLrSgnBnf("CTJURI")
recLrSgnBnf.CDRESI = tableLrSgnBnf("CDRESI")
recLrSgnBnf.NOVOIE = tableLrSgnBnf("NOVOIE")
recLrSgnBnf.CDPOST = tableLrSgnBnf("CDPOST")
recLrSgnBnf.LBCOMM2 = tableLrSgnBnf("LBCOMM2")
recLrSgnBnf.CDDEPT2 = tableLrSgnBnf("CDDEPT2")
recLrSgnBnf.CDPAYS2 = tableLrSgnBnf("CDPAYS2")
recLrSgnBnf.CDTRI1 = tableLrSgnBnf("CDTRI1")
recLrSgnBnf.CDTRI2 = tableLrSgnBnf("CDTRI2")
recLrSgnBnf.CDHABI = tableLrSgnBnf("CDHABI")
recLrSgnBnf.AMJ4 = tableLrSgnBnf("AMJ4")
recLrSgnBnf.HMSC = tableLrSgnBnf("HMSC")
recLrSgnBnf.CDAGCO = tableLrSgnBnf("CDAGCO")
recLrSgnBnf.CDPHMO = tableLrSgnBnf("CDPHMO")
recLrSgnBnf.CDCRMD = tableLrSgnBnf("CDCRMD")
recLrSgnBnf.INSSIR = tableLrSgnBnf("INSSIR")
recLrSgnBnf.FILL01 = tableLrSgnBnf("FILL01")
recLrSgnBnf.CDRESI1 = tableLrSgnBnf("CDRESI1")
recLrSgnBnf.CDACEN1 = tableLrSgnBnf("CDACEN1")
recLrSgnBnf.CTJURN1 = tableLrSgnBnf("CTJURN1")
recLrSgnBnf.CDPAYN1 = tableLrSgnBnf("CDPAYN1")
recLrSgnBnf.CDSEXE1 = tableLrSgnBnf("CDSEXE1")
recLrSgnBnf.CDPAYN2 = tableLrSgnBnf("CDPAYN2")
recLrSgnBnf.CDRESI2 = tableLrSgnBnf("CDRESI2")
recLrSgnBnf.CDACCO2 = tableLrSgnBnf("CDACCO2")
recLrSgnBnf.FILL02 = tableLrSgnBnf("FILL02")

End Sub


'-----------------------------------------------------
Sub tableLrSgnBnf_Open()
'-----------------------------------------------------

If Not tableLrSgnBnfOpen Then
    Set tableLrSgnBnf = MDB.OpenRecordset("LrSgnBnf")
    tableLrSgnBnf.Index = "PrimaryKey"
    tableLrSgnBnfOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableLrSgnBnf_PutBuffer(recLrSgnBnf As typeLrSgnBnf)
'---------------------------------------------------------

tableLrSgnBnf("CDBANQ") = recLrSgnBnf.CDBANQ
tableLrSgnBnf("CDDECL") = recLrSgnBnf.CDDECL
tableLrSgnBnf("RFBENF") = recLrSgnBnf.RFBENF
tableLrSgnBnf("NSIREN") = recLrSgnBnf.NSIREN
tableLrSgnBnf("NPREFI1") = recLrSgnBnf.NPREFI1
tableLrSgnBnf("NSUFFI1") = recLrSgnBnf.NSUFFI1
tableLrSgnBnf("NPREFI2") = recLrSgnBnf.NPREFI2
tableLrSgnBnf("AMJ1") = recLrSgnBnf.AMJ1
tableLrSgnBnf("NSIREN2") = recLrSgnBnf.NSIREN2
tableLrSgnBnf("NSUFFI2") = recLrSgnBnf.NSUFFI2
tableLrSgnBnf("AMJ2") = recLrSgnBnf.AMJ2
tableLrSgnBnf("NOMBNF") = recLrSgnBnf.NOMBNF
tableLrSgnBnf("PRENOM") = recLrSgnBnf.PRENOM
tableLrSgnBnf("CDSEXE") = recLrSgnBnf.CDSEXE
tableLrSgnBnf("JMA3") = recLrSgnBnf.JMA3
tableLrSgnBnf("CDPAYS1") = recLrSgnBnf.CDPAYS1
tableLrSgnBnf("CDDEPT1") = recLrSgnBnf.CDDEPT1
tableLrSgnBnf("CDCOMM1") = recLrSgnBnf.CDCOMM1
tableLrSgnBnf("LBCOMM1") = recLrSgnBnf.LBCOMM1
tableLrSgnBnf("NOMCJT") = recLrSgnBnf.NOMCJT
tableLrSgnBnf("CDACCO") = recLrSgnBnf.CDACCO
tableLrSgnBnf("CTJURI") = recLrSgnBnf.CTJURI
tableLrSgnBnf("CDRESI") = recLrSgnBnf.CDRESI
tableLrSgnBnf("NOVOIE") = recLrSgnBnf.NOVOIE
tableLrSgnBnf("CDPOST") = recLrSgnBnf.CDPOST
tableLrSgnBnf("LBCOMM2") = recLrSgnBnf.LBCOMM2
tableLrSgnBnf("CDDEPT2") = recLrSgnBnf.CDDEPT2
tableLrSgnBnf("CDPAYS2") = recLrSgnBnf.CDPAYS2
tableLrSgnBnf("CDTRI1") = recLrSgnBnf.CDTRI1
tableLrSgnBnf("CDTRI2") = recLrSgnBnf.CDTRI2
tableLrSgnBnf("CDHABI") = recLrSgnBnf.CDHABI
tableLrSgnBnf("AMJ4") = recLrSgnBnf.AMJ4
tableLrSgnBnf("HMSC") = recLrSgnBnf.HMSC
tableLrSgnBnf("CDAGCO") = recLrSgnBnf.CDAGCO
tableLrSgnBnf("CDPHMO") = recLrSgnBnf.CDPHMO
tableLrSgnBnf("CDCRMD") = recLrSgnBnf.CDCRMD
tableLrSgnBnf("INSSIR") = recLrSgnBnf.INSSIR
tableLrSgnBnf("FILL01") = recLrSgnBnf.FILL01
tableLrSgnBnf("CDRESI1") = recLrSgnBnf.CDRESI1
tableLrSgnBnf("CDACEN1") = recLrSgnBnf.CDACEN1
tableLrSgnBnf("CTJURN1") = recLrSgnBnf.CTJURN1
tableLrSgnBnf("CDPAYN1") = recLrSgnBnf.CDPAYN1
tableLrSgnBnf("CDSEXE1") = recLrSgnBnf.CDSEXE1
tableLrSgnBnf("CDPAYN2") = recLrSgnBnf.CDPAYN2
tableLrSgnBnf("CDRESI2") = recLrSgnBnf.CDRESI2
tableLrSgnBnf("CDACCO2") = recLrSgnBnf.CDACCO2
tableLrSgnBnf("FILL02") = recLrSgnBnf.FILL02
End Sub


'---------------------------------------------------------
Public Function tableLrSgnBnf_Read(recLrSgnBnf As typeLrSgnBnf) As Integer
'---------------------------------------------------------

On Error GoTo tableLrSgnBnf_Read_Error
tableLrSgnBnf_Read = 0


Select Case recLrSgnBnf.Method
     Case "Seek=       "
                        tableLrSgnBnf.Seek "=", recLrSgnBnf.RFBENF
                        If tableLrSgnBnf.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<=      "
                        tableLrSgnBnf.Seek "<=", recLrSgnBnf.RFBENF
                        If tableLrSgnBnf.NoMatch Then
                            Error 9998
                        End If
     Case "MoveNext    "
                        tableLrSgnBnf.MoveNext
                        If tableLrSgnBnf.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableLrSgnBnf.MovePrevious
                        If tableLrSgnBnf.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst   "
                        tableLrSgnBnf.MoveFirst
                        If tableLrSgnBnf.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast    "
                        tableLrSgnBnf.MoveLast
                        If tableLrSgnBnf.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recLrSgnBnf.Method <> "AddNew      " Then
    Call tableLrSgnBnf_GetBuffer(recLrSgnBnf)
End If

Exit Function

'---------------------------------------------------------
tableLrSgnBnf_Read_Error:
'---------------------------------------------------------

    tableLrSgnBnf_Read = Err
    Resume tableLrSgnBnf_Read_End

tableLrSgnBnf_Read_End:

End Function

'---------------------------------------------------------
Public Function tableLrSgnBnf_Update(recLrSgnBnf As typeLrSgnBnf) As Integer
'---------------------------------------------------------

On Error GoTo tableLrSgnBnfUpdate_Error
tableLrSgnBnf_Update = 0

Select Case recLrSgnBnf.Method

    Case "AddNew      "
                        tableLrSgnBnf.AddNew
                        Call tableLrSgnBnf_PutBuffer(recLrSgnBnf)
                        tableLrSgnBnf.Update
    Case "Update      "
                        tableLrSgnBnf.Edit
                        Call tableLrSgnBnf_PutBuffer(recLrSgnBnf)
                        tableLrSgnBnf.Update
    Case "Delete      "
                        tableLrSgnBnf.Delete
    Case Else
                        Error 9999
End Select


Exit Function

tableLrSgnBnfUpdate_Error:
'---------------------------------------------------------
    tableLrSgnBnf_Update = Err
    Resume tableLrSgnBnfUpdate_End

tableLrSgnBnfUpdate_End:

End Function








'-----------------------------------------------------
Sub dbLrSgnBnf_Error(recLrSgnBnf As typeLrSgnBnf)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & Trim(recLrSgnBnf.RFBENF) & Chr$(13)

Select Case Mid$(recLrSgnBnf.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recLrSgnBnf.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbLrSgnBnf.bas :  ( " & Trim(recLrSgnBnf.obj) & " : " & Trim(recLrSgnBnf.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbLrSgnBnf_Read(recLrSgnBnf As typeLrSgnBnf)
'-----------------------------------------------------

dbLrSgnBnf_Read = Null

recLrSgnBnf.Err = tableLrSgnBnf_Read(recLrSgnBnf)
If recLrSgnBnf.Err > 0 Then

    If recLrSgnBnf.Err < 9990 Or recLrSgnBnf.Err >= 9999 Then
        Call dbLrSgnBnf_Error(recLrSgnBnf)
        dbLrSgnBnf_Read = recLrSgnBnf.Err
    End If
End If

End Function

'-----------------------------------------------------
Function dbLrSgnBnf_ReadZ(recLrSgnBnf As typeLrSgnBnf)
'-----------------------------------------------------

dbLrSgnBnf_ReadZ = Null

recLrSgnBnf.Err = tableLrSgnBnf_Read(recLrSgnBnf)
If recLrSgnBnf.Err > 0 Then
    dbLrSgnBnf_ReadZ = recLrSgnBnf.Err
    recLrSgnBnf.NOMBNF = "? " & recLrSgnBnf.RFBENF
End If

End Function


'-----------------------------------------------------
Function dbLrSgnBnf_Update(recLrSgnBnf As typeLrSgnBnf)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
'$$$BeginTrans

dbLrSgnBnf_Update = Null


recLrSgnBnf.Err = tableLrSgnBnf_Update(recLrSgnBnf)

If recLrSgnBnf.Err <> 0 Then
    Call dbLrSgnBnf_Error(recLrSgnBnf)
    dbLrSgnBnf_Update = recLrSgnBnf.Err
'$$$    Rollback
    Exit Function
End If

'$$$CommitTrans


'=====================================================
End Function


