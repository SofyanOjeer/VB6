Attribute VB_Name = "mdbLrRisque"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableLrRisque As Recordset
Dim tableLrRisqueOpen As Boolean

Type typeLrRisque
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    CDBANQ       As String * 5
    CDDECL       As String * 5
    RFBENF       As String * 16
    CDGUIC       As String * 5
    DTCENT1      As String * 6
    CDORSP       As String * 1
    CDCPCO       As String * 1
    CDCPJO       As String * 1
    CDDMAJ       As String * 1
    CDHABI       As String * 10
    AMJDN        As String * 8
    HMSCDN       As String * 8
    CDAGCO       As String * 5
    CDSWAP       As String * 1
    TYCENT       As String * 1
    CDPERI       As String * 1
    CDTRAN       As String * 1
    IDPREF       As String * 2
    NSIREN       As String * 9
    IDSUFF       As String * 2
    MT01         As Currency
    MT02         As Currency
    MT03         As Currency
    MT04         As Currency
    MT05         As Currency
    MT06         As Currency
    MT07         As Currency
    MT08         As Currency
    MT09         As Currency
    MT10         As Currency
    MT11         As Currency
    MT12         As Currency
    MT13         As Currency
    MT14         As Currency
    MT15         As Currency
    MT16         As Currency
    MT17         As Currency
    MT18         As Currency
    MT19         As Currency
    MT20         As Currency
    MTTOTAL      As Currency
    DTC          As String * 6
    FILL01       As String * 19

End Type

Public recLrRisque As typeLrRisque


'---------------------------------------------------------
Public Function Import_LrRisque(MsgTxt As String, recLrRisque As typeLrRisque, optFRF As Boolean)
'---------------------------------------------------------
Dim K As Integer, I As Integer
Import_LrRisque = Null
recLrRisque.obj = "LRRISQUE"
recLrRisque.Method = ""
recLrRisque.Err = ""

recLrRisque.CDBANQ = mId$(MsgTxt, 1, 5)
recLrRisque.CDDECL = mId$(MsgTxt, 6, 5)
recLrRisque.RFBENF = mId$(MsgTxt, 11, 16)
recLrRisque.CDGUIC = mId$(MsgTxt, 27, 5)
recLrRisque.DTCENT1 = mId$(MsgTxt, 32, 6)
recLrRisque.CDORSP = mId$(MsgTxt, 38, 1)
recLrRisque.CDCPCO = mId$(MsgTxt, 39, 1)
recLrRisque.CDCPJO = mId$(MsgTxt, 40, 1)
recLrRisque.CDDMAJ = mId$(MsgTxt, 41, 1)
recLrRisque.CDHABI = mId$(MsgTxt, 42, 10)
recLrRisque.AMJDN = mId$(MsgTxt, 52, 8)
recLrRisque.HMSCDN = mId$(MsgTxt, 60, 8)
recLrRisque.CDAGCO = mId$(MsgTxt, 68, 5)
recLrRisque.CDSWAP = mId$(MsgTxt, 73, 1)
recLrRisque.TYCENT = mId$(MsgTxt, 74, 1)
recLrRisque.CDPERI = mId$(MsgTxt, 75, 1)
recLrRisque.CDTRAN = mId$(MsgTxt, 76, 1)
recLrRisque.IDPREF = mId$(MsgTxt, 77, 2)
recLrRisque.NSIREN = mId$(MsgTxt, 79, 9)
recLrRisque.IDSUFF = mId$(MsgTxt, 88, 2)
recLrRisque.MT01 = CCur(Val(mId$(MsgTxt, 90, 16)))
recLrRisque.MT02 = CCur(Val(mId$(MsgTxt, 106, 16)))
recLrRisque.MT03 = CCur(Val(mId$(MsgTxt, 122, 16)))
recLrRisque.MT04 = CCur(Val(mId$(MsgTxt, 138, 16)))
recLrRisque.MT05 = CCur(Val(mId$(MsgTxt, 154, 16)))
recLrRisque.MT06 = CCur(Val(mId$(MsgTxt, 170, 16)))
recLrRisque.MT07 = CCur(Val(mId$(MsgTxt, 186, 16)))
recLrRisque.MT08 = CCur(Val(mId$(MsgTxt, 202, 16)))
recLrRisque.MT09 = CCur(Val(mId$(MsgTxt, 218, 16)))
recLrRisque.MT10 = CCur(Val(mId$(MsgTxt, 234, 16)))
recLrRisque.MT11 = CCur(Val(mId$(MsgTxt, 250, 16)))
recLrRisque.MT12 = CCur(Val(mId$(MsgTxt, 266, 16)))
recLrRisque.MT13 = CCur(Val(mId$(MsgTxt, 282, 16)))
recLrRisque.MT14 = CCur(Val(mId$(MsgTxt, 298, 16)))
recLrRisque.MT15 = CCur(Val(mId$(MsgTxt, 314, 16)))
recLrRisque.MT16 = CCur(Val(mId$(MsgTxt, 330, 16)))
recLrRisque.MT17 = CCur(Val(mId$(MsgTxt, 346, 16)))
recLrRisque.MT18 = CCur(Val(mId$(MsgTxt, 362, 16)))
recLrRisque.MT19 = CCur(Val(mId$(MsgTxt, 378, 16)))
recLrRisque.MT20 = CCur(Val(mId$(MsgTxt, 394, 16)))
recLrRisque.MTTOTAL = CCur(Val(mId$(MsgTxt, 410, 16)))
recLrRisque.DTC = mId$(MsgTxt, 426, 6)
recLrRisque.FILL01 = mId$(MsgTxt, 430, 19)

If optFRF Then
    If recLrRisque.DTCENT1 > "199906" Then Import_LrRisque_CV recLrRisque
Else
    If recLrRisque.DTCENT1 <= "199906" Then Import_LrRisque_CV recLrRisque
End If

End Function
'-----------------------------------------------------
Function dbLrRisque_ReadZ(recLrRisque As typeLrRisque)
'-----------------------------------------------------

dbLrRisque_ReadZ = Null

recLrRisque.Err = tableLrRisque_Read(recLrRisque)
If recLrRisque.Err > 0 Then
    dbLrRisque_ReadZ = recLrRisque.Err
    recLrRisque.MT01 = 0
    recLrRisque.MT02 = 0
    recLrRisque.MT03 = 0
    recLrRisque.MT04 = 0
    recLrRisque.MT05 = 0
    recLrRisque.MT06 = 0
    recLrRisque.MT07 = 0
    recLrRisque.MT08 = 0
    recLrRisque.MT09 = 0
    recLrRisque.MT10 = 0
    recLrRisque.MT11 = 0
    recLrRisque.MT12 = 0
    recLrRisque.MT13 = 0
    recLrRisque.MT14 = 0
    recLrRisque.MT15 = 0
    recLrRisque.MT16 = 0
    recLrRisque.MT17 = 0
    recLrRisque.MT18 = 0
    recLrRisque.MT19 = 0
    recLrRisque.MT20 = 0
    recLrRisque.MTTOTAL = 0
End If

End Function




'---------------------------------------------------------
'-----------------------------------------------------
Sub tableLrRisque_Close()
'-----------------------------------------------------
If tableLrRisqueOpen Then
    tableLrRisque.Close
    tableLrRisqueOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableLrRisque_GetBuffer(recLrRisque As typeLrRisque)
'---------------------------------------------------------

recLrRisque.CDBANQ = tableLrRisque("CDBANQ")
recLrRisque.CDDECL = tableLrRisque("CDDECL")
recLrRisque.RFBENF = tableLrRisque("RFBENF")

recLrRisque.CDGUIC = tableLrRisque("CDGUIC")
recLrRisque.DTCENT1 = tableLrRisque("DTCENT1")
recLrRisque.CDORSP = tableLrRisque("CDORSP")
recLrRisque.CDCPJO = tableLrRisque("CDCPJO")
recLrRisque.CDCPCO = tableLrRisque("CDCPCO")
recLrRisque.CDDMAJ = tableLrRisque("CDDMAJ")
recLrRisque.CDHABI = tableLrRisque("CDHABI")
recLrRisque.AMJDN = tableLrRisque("AMJDN")

recLrRisque.HMSCDN = tableLrRisque("HMSCDN")
recLrRisque.CDAGCO = tableLrRisque("CDAGCO")
recLrRisque.CDSWAP = tableLrRisque("CDSWAP")
recLrRisque.TYCENT = tableLrRisque("TYCENT")
recLrRisque.CDPERI = tableLrRisque("CDPERI")
recLrRisque.CDTRAN = tableLrRisque("CDTRAN")
recLrRisque.IDPREF = tableLrRisque("IDPREF")
recLrRisque.NSIREN = tableLrRisque("NSIREN")
recLrRisque.IDSUFF = tableLrRisque("IDSUFF")
recLrRisque.MT01 = tableLrRisque("MT01")
recLrRisque.MT02 = tableLrRisque("MT02")
recLrRisque.MT03 = tableLrRisque("MT03")
recLrRisque.MT04 = tableLrRisque("MT04")
recLrRisque.MT05 = tableLrRisque("MT05")
recLrRisque.MT06 = tableLrRisque("MT06")
recLrRisque.MT07 = tableLrRisque("MT07")
recLrRisque.MT08 = tableLrRisque("MT08")
recLrRisque.MT09 = tableLrRisque("MT09")
recLrRisque.MT10 = tableLrRisque("MT10")
recLrRisque.MT11 = tableLrRisque("MT11")
recLrRisque.MT12 = tableLrRisque("MT12")
recLrRisque.MT13 = tableLrRisque("MT13")
recLrRisque.MT14 = tableLrRisque("MT14")
recLrRisque.MT15 = tableLrRisque("MT15")
recLrRisque.MT16 = tableLrRisque("MT16")
recLrRisque.MT17 = tableLrRisque("MT17")
recLrRisque.MT18 = tableLrRisque("MT18")
recLrRisque.MT19 = tableLrRisque("MT19")
recLrRisque.MT20 = tableLrRisque("MT20")
recLrRisque.MTTOTAL = tableLrRisque("MTTOTAL")
recLrRisque.DTC = tableLrRisque("DTC")
recLrRisque.FILL01 = tableLrRisque("FILL01")

End Sub


'-----------------------------------------------------
Sub tableLrRisque_Open()
'-----------------------------------------------------

If Not tableLrRisqueOpen Then
    Set tableLrRisque = MDB.OpenRecordset("LrRisque")
    tableLrRisque.Index = "PrimaryKey"
    tableLrRisqueOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableLrRisque_PutBuffer(recLrRisque As typeLrRisque)
'---------------------------------------------------------

tableLrRisque("CDBANQ") = recLrRisque.CDBANQ
tableLrRisque("CDDECL") = recLrRisque.CDDECL
tableLrRisque("RFBENF") = recLrRisque.RFBENF
tableLrRisque("CDGUIC") = recLrRisque.CDGUIC
tableLrRisque("DTCENT1") = recLrRisque.DTCENT1
tableLrRisque("CDORSP") = recLrRisque.CDORSP
tableLrRisque("CDCPJO") = recLrRisque.CDCPJO
tableLrRisque("CDCPCO") = recLrRisque.CDCPCO

tableLrRisque("CDDMAJ") = recLrRisque.CDDMAJ
tableLrRisque("CDHABI") = recLrRisque.CDHABI
tableLrRisque("AMJDN") = recLrRisque.AMJDN
tableLrRisque("HMSCDN") = recLrRisque.HMSCDN
tableLrRisque("CDAGCO") = recLrRisque.CDAGCO
tableLrRisque("CDSWAP") = recLrRisque.CDSWAP
tableLrRisque("TYCENT") = recLrRisque.TYCENT
tableLrRisque("CDPERI") = recLrRisque.CDPERI
tableLrRisque("CDTRAN") = recLrRisque.CDTRAN

tableLrRisque("IDPREF") = recLrRisque.IDPREF
tableLrRisque("NSIREN") = recLrRisque.NSIREN
 tableLrRisque("IDSUFF") = recLrRisque.IDSUFF
 tableLrRisque("MT01") = recLrRisque.MT01
tableLrRisque("MT02") = recLrRisque.MT02
tableLrRisque("MT03") = recLrRisque.MT03
tableLrRisque("MT04") = recLrRisque.MT04
tableLrRisque("MT05") = recLrRisque.MT05
tableLrRisque("MT06") = recLrRisque.MT06
tableLrRisque("MT07") = recLrRisque.MT07
tableLrRisque("MT08") = recLrRisque.MT08
tableLrRisque("MT09") = recLrRisque.MT09
tableLrRisque("MT10") = recLrRisque.MT10
tableLrRisque("MT11") = recLrRisque.MT11
tableLrRisque("MT12") = recLrRisque.MT12
tableLrRisque("MT13") = recLrRisque.MT13
tableLrRisque("MT14") = recLrRisque.MT14
tableLrRisque("MT15") = recLrRisque.MT15
tableLrRisque("MT16") = recLrRisque.MT16
tableLrRisque("MT17") = recLrRisque.MT17
tableLrRisque("MT18") = recLrRisque.MT18
tableLrRisque("MT19") = recLrRisque.MT19
tableLrRisque("MT20") = recLrRisque.MT20
tableLrRisque("MTTOTAL") = recLrRisque.MTTOTAL
tableLrRisque("DTC") = recLrRisque.DTC
tableLrRisque("FILL01") = recLrRisque.FILL01
End Sub


'---------------------------------------------------------
Public Function tableLrRisque_Read(recLrRisque As typeLrRisque) As Integer
'---------------------------------------------------------

On Error GoTo tableLrRisque_Read_Error
tableLrRisque_Read = 0


Select Case recLrRisque.Method
     Case "Seek=       ", "AddNew      ", "Update      ", "Delete      "

                        tableLrRisque.Seek "=", recLrRisque.RFBENF, recLrRisque.CDCPCO, recLrRisque.DTCENT1
                        If tableLrRisque.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<=      "
                        tableLrRisque.Seek "<=", recLrRisque.RFBENF, recLrRisque.CDCPCO, recLrRisque.DTCENT1
                        If tableLrRisque.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>=      "
                        tableLrRisque.Seek ">=", recLrRisque.RFBENF, recLrRisque.CDCPCO, recLrRisque.DTCENT1
                        If tableLrRisque.NoMatch Then
                            Error 9998
                        End If
       Case "Seek>       "
                        tableLrRisque.Seek ">", recLrRisque.RFBENF, recLrRisque.CDCPCO, recLrRisque.DTCENT1
                        If tableLrRisque.NoMatch Then
                            Error 9998
                        End If
   Case "MoveNext    "
                        tableLrRisque.MoveNext
                        If tableLrRisque.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableLrRisque.MovePrevious
                        If tableLrRisque.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst   "
                        tableLrRisque.MoveFirst
                        If tableLrRisque.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast    "
                        tableLrRisque.MoveLast
                        If tableLrRisque.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recLrRisque.Method <> "AddNew      " Then
    Call tableLrRisque_GetBuffer(recLrRisque)
End If

Exit Function

'---------------------------------------------------------
tableLrRisque_Read_Error:
'---------------------------------------------------------

    tableLrRisque_Read = Err
    Resume tableLrRisque_Read_End

tableLrRisque_Read_End:

End Function

'---------------------------------------------------------
Public Function tableLrRisque_Update(recLrRisque As typeLrRisque) As Integer
'---------------------------------------------------------

On Error GoTo tableLrRisqueUpdate_Error
tableLrRisque_Update = 0

Select Case recLrRisque.Method

    Case "AddNew      "
                        tableLrRisque.AddNew
                        Call tableLrRisque_PutBuffer(recLrRisque)
                        tableLrRisque.Update
    Case "Update      "
                        tableLrRisque.Edit
                        Call tableLrRisque_PutBuffer(recLrRisque)
                        tableLrRisque.Update
    Case "Delete      "
                        tableLrRisque.Delete
    Case Else
                        Error 9999
End Select

Exit Function

tableLrRisqueUpdate_Error:
'---------------------------------------------------------
    tableLrRisque_Update = Err
    Resume tableLrRisqueUpdate_End

tableLrRisqueUpdate_End:

End Function








'-----------------------------------------------------
Sub dbLrRisque_Error(recLrRisque As typeLrRisque)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & Trim(recLrRisque.RFBENF) & " : " & Trim(recLrRisque.DTCENT1) & " : " & Trim(recLrRisque.CDCPCO) & Chr$(13)

Select Case mId$(recLrRisque.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recLrRisque.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbLrRisque.bas :  ( " & Trim(recLrRisque.obj) & " : " & Trim(recLrRisque.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbLrRisque_Read(recLrRisque As typeLrRisque)
'-----------------------------------------------------

dbLrRisque_Read = Null

recLrRisque.Err = tableLrRisque_Read(recLrRisque)
If recLrRisque.Err > 0 Then

    If recLrRisque.Err < 9990 Or recLrRisque.Err >= 9999 Then
        Call dbLrRisque_Error(recLrRisque)
        dbLrRisque_Read = recLrRisque.Err
    End If
End If

End Function

'-----------------------------------------------------
Function dbLrRisque_Update(recLrRisque As typeLrRisque)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
'$$$BeginTrans

dbLrRisque_Update = Null


recLrRisque.Err = tableLrRisque_Update(recLrRisque)

If recLrRisque.Err <> 0 Then
    Call dbLrRisque_Error(recLrRisque)
    dbLrRisque_Update = recLrRisque.Err
'$$$    Rollback
    Exit Function
End If

'$$$CommitTrans


'=====================================================
End Function



Public Sub Import_LrRisque_CV(recLrRisque As typeLrRisque)
LrCdr_CV1.Montant = recLrRisque.MT01
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT01 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT02
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT02 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT03
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT03 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT04
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT04 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT05
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT05 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT06
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT06 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT07
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT07 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT08
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT08 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT09
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT09 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT10
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT10 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT11
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT11 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT12
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT12 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT13
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT13 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT14
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT14 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT15
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT15 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT16
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT16 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT17
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT17 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT18
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT18 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT19
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT19 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MT20
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MT20 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRisque.MTTOTAL
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRisque.MTTOTAL = LrCdr_CV2.Montant

End Sub
