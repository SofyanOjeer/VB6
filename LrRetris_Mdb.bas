Attribute VB_Name = "mdbLrRETRIS"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public tableLrRetris As Recordset
Dim tableLrRetrisOpen As Boolean

Type typeLrRetris
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
    CDENBF       As String * 2
    CDSWAP       As String * 1
    COTBDF       As String * 4
    DTARSS       As String * 6
    MTARSS       As Currency
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
    MT21         As Currency
    MT22         As Currency
    MT23         As Currency
    MT24         As Currency
    MT25         As Currency
    MTTOTAL      As Currency
    TYCENT       As String * 1
    FILL01       As String * 21

End Type

Public recLrRetris As typeLrRetris


'---------------------------------------------------------
Public Function Import_LrRetris(MsgTxt As String, recLrRetris As typeLrRetris, optFRF As Boolean)
'---------------------------------------------------------
Dim K As Integer, I As Integer, X5 As String * 5
Dim wX As String
Import_LrRetris = Null
recLrRetris.obj = "LrRetris"
recLrRetris.Method = ""
recLrRetris.Err = ""

recLrRetris.CDBANQ = mId$(MsgTxt, 1, 5)
recLrRetris.CDDECL = mId$(MsgTxt, 6, 5)
recLrRetris.RFBENF = mId$(MsgTxt, 11, 16)
recLrRetris.CDGUIC = mId$(MsgTxt, 27, 5)
recLrRetris.DTCENT1 = mId$(MsgTxt, 32, 6)
recLrRetris.CDORSP = mId$(MsgTxt, 38, 1)
recLrRetris.CDCPCO = mId$(MsgTxt, 39, 1)
recLrRetris.CDENBF = mId$(MsgTxt, 40, 2)
recLrRetris.CDSWAP = mId$(MsgTxt, 42, 1)
recLrRetris.COTBDF = mId$(MsgTxt, 43, 4)
wX = Trim(recLrRetris.COTBDF)
If wX = "" Or wX = "000" Or wX = "00" Then recLrRetris.COTBDF = "0000"

recLrRetris.DTARSS = mId$(MsgTxt, 47, 6)
recLrRetris.MTARSS = CCur(Val(mId$(MsgTxt, 53, 10)))
recLrRetris.MT01 = CCur(Val(mId$(MsgTxt, 63, 16)))
recLrRetris.MT02 = CCur(Val(mId$(MsgTxt, 79, 16)))
recLrRetris.MT03 = CCur(Val(mId$(MsgTxt, 95, 16)))
recLrRetris.MT04 = CCur(Val(mId$(MsgTxt, 111, 16)))
recLrRetris.MT05 = CCur(Val(mId$(MsgTxt, 127, 16)))
recLrRetris.MT06 = CCur(Val(mId$(MsgTxt, 143, 16)))
recLrRetris.MT07 = CCur(Val(mId$(MsgTxt, 159, 16)))
recLrRetris.MT08 = CCur(Val(mId$(MsgTxt, 175, 16)))
recLrRetris.MT09 = CCur(Val(mId$(MsgTxt, 191, 16)))
recLrRetris.MT10 = CCur(Val(mId$(MsgTxt, 207, 16)))
recLrRetris.MT11 = CCur(Val(mId$(MsgTxt, 223, 16)))
recLrRetris.MT12 = CCur(Val(mId$(MsgTxt, 239, 16)))
recLrRetris.MT13 = CCur(Val(mId$(MsgTxt, 255, 16)))
recLrRetris.MT14 = CCur(Val(mId$(MsgTxt, 271, 16)))
recLrRetris.MT15 = CCur(Val(mId$(MsgTxt, 287, 16)))
recLrRetris.MT16 = CCur(Val(mId$(MsgTxt, 303, 16)))
recLrRetris.MT17 = CCur(Val(mId$(MsgTxt, 319, 16)))
recLrRetris.MT18 = CCur(Val(mId$(MsgTxt, 335, 16)))
recLrRetris.MT19 = CCur(Val(mId$(MsgTxt, 351, 16)))
recLrRetris.MT20 = CCur(Val(mId$(MsgTxt, 367, 16)))
recLrRetris.MT21 = CCur(Val(mId$(MsgTxt, 383, 16)))
recLrRetris.MT22 = CCur(Val(mId$(MsgTxt, 399, 16)))
recLrRetris.MT23 = CCur(Val(mId$(MsgTxt, 415, 16)))
recLrRetris.MT24 = CCur(Val(mId$(MsgTxt, 431, 16)))
recLrRetris.MT25 = CCur(Val(mId$(MsgTxt, 447, 16)))
recLrRetris.MTTOTAL = CCur(Val(mId$(MsgTxt, 463, 16)))
recLrRetris.TYCENT = mId$(MsgTxt, 479, 1)
recLrRetris.FILL01 = mId$(MsgTxt, 480, 21)
If IsNumeric(recLrRetris.DTCENT1) Then
    If recLrRetris.DTCENT1 < 199810 Then
        recLrRetris.MT01 = recLrRetris.MT01 * 10
        recLrRetris.MT02 = recLrRetris.MT02 * 10
        recLrRetris.MT03 = recLrRetris.MT03 * 10
        recLrRetris.MT04 = recLrRetris.MT04 * 10
        recLrRetris.MT05 = recLrRetris.MT05 * 10
        recLrRetris.MT06 = recLrRetris.MT06 * 10
        recLrRetris.MT07 = recLrRetris.MT07 * 10
        recLrRetris.MT08 = recLrRetris.MT08 * 10
        recLrRetris.MT09 = recLrRetris.MT09 * 10
        recLrRetris.MT10 = recLrRetris.MT10 * 10
        recLrRetris.MT11 = recLrRetris.MT11 * 10
        recLrRetris.MT12 = recLrRetris.MT12 * 10
        recLrRetris.MT13 = recLrRetris.MT13 * 10
        recLrRetris.MT14 = recLrRetris.MT14 * 10
        recLrRetris.MT15 = recLrRetris.MT15 * 10
        recLrRetris.MT16 = recLrRetris.MT16 * 10
        recLrRetris.MT17 = recLrRetris.MT17 * 10
        recLrRetris.MT18 = recLrRetris.MT18 * 10
        recLrRetris.MT19 = recLrRetris.MT19 * 10
        recLrRetris.MT10 = recLrRetris.MT10 * 10
        recLrRetris.MT21 = recLrRetris.MT21 * 10
        recLrRetris.MT22 = recLrRetris.MT22 * 10
        recLrRetris.MT23 = recLrRetris.MT23 * 10
        recLrRetris.MT24 = recLrRetris.MT24 * 10
        recLrRetris.MT25 = recLrRetris.MT25 * 10
        recLrRetris.MTTOTAL = recLrRetris.MTTOTAL * 10
    End If
End If
        
If Len(Trim(recLrRetris.RFBENF)) > 9 Then
    K = InStr(1, recLrRetris.RFBENF, "/")
    If K > 0 Then
        X5 = mId$(recLrRetris.RFBENF, K + 2, 5)
        Select Case X5
            Case "90001": X5 = "85060"
            Case "90006": X5 = "35157"
            Case "90007": X5 = "35161"
            Case "90029": X5 = "35116"
            Case "90207": X5 = "35120"
            Case "90234": X5 = "35124"
            Case "90238": X5 = "35123"
            Case "90239": X5 = "35147"
            Case "90257": X5 = "35127"
            Case "90258": X5 = "35128"
            Case "90240": X5 = "35145"
            Case "91008": X5 = "35129"
            Case "93074": X5 = "35131"
            Case "95055": X5 = "25199"
            Case "99008": X5 = "85053"
            Case "99223": X5 = "85062"
            Case "99221": X5 = "85061"
            Case "90061": X5 = "85059"
            Case "90057": X5 = "85056"
            Case "99002": X5 = "85050"
            Case "99051": X5 = "85052"
            Case "99008": X5 = "85053"
            Case "99003": X5 = "85051"
            Case "99002": X5 = "85050"
            Case "92440": X5 = "25219"
            Case "95005": X5 = "25222"
            Case "96014": X5 = "25224"
            Case "96068": X5 = "25213"
            Case "96071": X5 = "25226"
            Case "96077": X5 = "25227"
            Case "96078": X5 = "25228"
            Case "96080": X5 = "25229"
            Case "96084": X5 = "25120"
            Case "96075": X5 = "25230"
            Case "96086": X5 = "25230"
            Case "96074": X5 = "25120"
            Case "96095": X5 = "25258"
            Case "97142": X5 = "35139"
            Case "99001": X5 = "25232"
            Case "96095": X5 = "25258"
            Case "99202": X5 = "85054"
        End Select
    
        If X5 < "30000" Then
            recLrRetris.RFBENF = X5 & "/000"
        Else
            recLrRetris.RFBENF = X5 & "/001"
        End If
    End If
End If

Select Case mId$(recLrRetris.RFBENF, 1, 5)
    Case "96085": Mid$(recLrRetris.RFBENF, 1, 5) = "25230"
    Case "99005": Mid$(recLrRetris.RFBENF, 1, 5) = "85052"
End Select

If optFRF Then
    If recLrRetris.DTCENT1 > "199906" Then Import_LrRetris_CV recLrRetris
Else
    If recLrRetris.DTCENT1 <= "199906" Then Import_LrRetris_CV recLrRetris
End If


End Function

Public Sub Import_LrRetris_CV(recLrRetris As typeLrRetris)
LrCdr_CV1.Montant = recLrRetris.MT01
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT01 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT02
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT02 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT03
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT03 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT04
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT04 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT05
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT05 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT06
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT06 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT07
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT07 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT08
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT08 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT09
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT09 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT10
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT10 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT11
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT11 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT12
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT12 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT13
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT13 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT14
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT14 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT15
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT15 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT16
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT16 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT17
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT17 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT18
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT18 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT19
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT19 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT20
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT20 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT21
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT21 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT22
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT22 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT23
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT23 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT24
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT24 = LrCdr_CV2.Montant

LrCdr_CV1.Montant = recLrRetris.MT25
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MT25 = LrCdr_CV2.Montant
LrCdr_CV1.Montant = recLrRetris.MTTOTAL
Call CV_Calc(LrCdr_CV1, LrCdr_CV2, LrCdr_CV3)
recLrRetris.MTTOTAL = LrCdr_CV2.Montant

End Sub

'---------------------------------------------------------
'-----------------------------------------------------
Sub tableLrRetris_Close()
'-----------------------------------------------------
If tableLrRetrisOpen Then
    tableLrRetris.Close
    tableLrRetrisOpen = False
End If

End Sub


'---------------------------------------------------------
Public Sub tableLrRetris_GetBuffer(recLrRetris As typeLrRetris)
'---------------------------------------------------------

recLrRetris.CDBANQ = tableLrRetris("CDBANQ")
recLrRetris.CDDECL = tableLrRetris("CDDECL")
recLrRetris.RFBENF = tableLrRetris("RFBENF")

recLrRetris.CDGUIC = tableLrRetris("CDGUIC")
recLrRetris.DTCENT1 = tableLrRetris("DTCENT1")
recLrRetris.CDORSP = tableLrRetris("CDORSP")
recLrRetris.CDCPCO = tableLrRetris("CDCPCO")
recLrRetris.CDENBF = tableLrRetris("CDENBF")
recLrRetris.CDSWAP = tableLrRetris("CDSWAP")
recLrRetris.COTBDF = tableLrRetris("COTBDF")
recLrRetris.DTARSS = tableLrRetris("DTARSS")
recLrRetris.MTARSS = tableLrRetris("MTARSS")
recLrRetris.MT01 = tableLrRetris("MT01")
recLrRetris.MT02 = tableLrRetris("MT02")
recLrRetris.MT03 = tableLrRetris("MT03")
recLrRetris.MT04 = tableLrRetris("MT04")
recLrRetris.MT05 = tableLrRetris("MT05")
recLrRetris.MT06 = tableLrRetris("MT06")
recLrRetris.MT07 = tableLrRetris("MT07")
recLrRetris.MT08 = tableLrRetris("MT08")
recLrRetris.MT09 = tableLrRetris("MT09")
recLrRetris.MT10 = tableLrRetris("MT10")
recLrRetris.MT11 = tableLrRetris("MT11")
recLrRetris.MT12 = tableLrRetris("MT12")
recLrRetris.MT13 = tableLrRetris("MT13")
recLrRetris.MT14 = tableLrRetris("MT14")
recLrRetris.MT15 = tableLrRetris("MT15")
recLrRetris.MT16 = tableLrRetris("MT16")
recLrRetris.MT17 = tableLrRetris("MT17")
recLrRetris.MT18 = tableLrRetris("MT18")
recLrRetris.MT19 = tableLrRetris("MT19")
recLrRetris.MT20 = tableLrRetris("MT20")
recLrRetris.MT21 = tableLrRetris("MT21")
recLrRetris.MT22 = tableLrRetris("MT22")
recLrRetris.MT23 = tableLrRetris("MT23")
recLrRetris.MT24 = tableLrRetris("MT24")
recLrRetris.MT25 = tableLrRetris("MT25")
recLrRetris.MTTOTAL = tableLrRetris("MTTOTAL")
recLrRetris.TYCENT = tableLrRetris("TYCENT")
recLrRetris.FILL01 = tableLrRetris("FILL01")

End Sub


'-----------------------------------------------------
Sub tableLrRetris_Open()
'-----------------------------------------------------

If Not tableLrRetrisOpen Then
    Set tableLrRetris = MDB.OpenRecordset("LrRetris")
    tableLrRetris.Index = "PrimaryKey"
    tableLrRetrisOpen = True
End If
End Sub

'---------------------------------------------------------
Public Sub tableLrRetris_PutBuffer(recLrRetris As typeLrRetris)
'---------------------------------------------------------

tableLrRetris("CDBANQ") = recLrRetris.CDBANQ
tableLrRetris("CDDECL") = recLrRetris.CDDECL
tableLrRetris("RFBENF") = recLrRetris.RFBENF
tableLrRetris("CDGUIC") = recLrRetris.CDGUIC
tableLrRetris("DTCENT1") = recLrRetris.DTCENT1
tableLrRetris("CDORSP") = recLrRetris.CDORSP
tableLrRetris("CDENBF") = recLrRetris.CDENBF
tableLrRetris("CDCPCO") = recLrRetris.CDCPCO
tableLrRetris("CDSWAP") = recLrRetris.CDSWAP
tableLrRetris("COTBDF") = recLrRetris.COTBDF
tableLrRetris("DTARSS") = recLrRetris.DTARSS
tableLrRetris("MTARSS") = recLrRetris.MTARSS
tableLrRetris("MT01") = recLrRetris.MT01
tableLrRetris("MT02") = recLrRetris.MT02
tableLrRetris("MT03") = recLrRetris.MT03
tableLrRetris("MT04") = recLrRetris.MT04
tableLrRetris("MT05") = recLrRetris.MT05
tableLrRetris("MT06") = recLrRetris.MT06
tableLrRetris("MT07") = recLrRetris.MT07
tableLrRetris("MT08") = recLrRetris.MT08
tableLrRetris("MT09") = recLrRetris.MT09
tableLrRetris("MT10") = recLrRetris.MT10
tableLrRetris("MT11") = recLrRetris.MT11
tableLrRetris("MT12") = recLrRetris.MT12
tableLrRetris("MT13") = recLrRetris.MT13
tableLrRetris("MT14") = recLrRetris.MT14
tableLrRetris("MT15") = recLrRetris.MT15
tableLrRetris("MT16") = recLrRetris.MT16
tableLrRetris("MT17") = recLrRetris.MT17
tableLrRetris("MT18") = recLrRetris.MT18
tableLrRetris("MT19") = recLrRetris.MT19
tableLrRetris("MT20") = recLrRetris.MT20
tableLrRetris("MT21") = recLrRetris.MT21
tableLrRetris("MT22") = recLrRetris.MT22
tableLrRetris("MT23") = recLrRetris.MT23
tableLrRetris("MT24") = recLrRetris.MT24
tableLrRetris("MT25") = recLrRetris.MT25
tableLrRetris("MTTOTAL") = recLrRetris.MTTOTAL
tableLrRetris("TYCENT") = recLrRetris.TYCENT
tableLrRetris("FILL01") = recLrRetris.FILL01
End Sub


'---------------------------------------------------------
Public Function tableLrRetris_Read(recLrRetris As typeLrRetris) As Integer
'---------------------------------------------------------

On Error GoTo tableLrRetris_Read_Error
tableLrRetris_Read = 0


Select Case recLrRetris.Method
     Case "Seek=       ", "AddNew      ", "Update      ", "Delete      "
                        tableLrRetris.Seek "=", recLrRetris.RFBENF, recLrRetris.CDCPCO, recLrRetris.DTCENT1
                        If tableLrRetris.NoMatch Then
                            Error 9998
                        End If
     Case "Seek<=      "
                        tableLrRetris.Seek "<=", recLrRetris.RFBENF, recLrRetris.CDCPCO, recLrRetris.DTCENT1
                        If tableLrRetris.NoMatch Then
                            Error 9998
                        End If
     Case "MoveNext    "
                        tableLrRetris.MoveNext
                        If tableLrRetris.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableLrRetris.MovePrevious
                        If tableLrRetris.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst   "
                        tableLrRetris.MoveFirst
                        If tableLrRetris.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast    "
                        tableLrRetris.MoveLast
                        If tableLrRetris.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recLrRetris.Method <> "AddNew      " Then
    Call tableLrRetris_GetBuffer(recLrRetris)
End If

Exit Function

'---------------------------------------------------------
tableLrRetris_Read_Error:
'---------------------------------------------------------

    tableLrRetris_Read = Err
    Resume tableLrRetris_Read_End

tableLrRetris_Read_End:

End Function

'---------------------------------------------------------
Public Function tableLrRetris_Update(recLrRetris As typeLrRetris) As Integer
'---------------------------------------------------------

On Error GoTo tableLrRetrisUpdate_Error
tableLrRetris_Update = 0

Select Case recLrRetris.Method

    Case "AddNew      "
                        tableLrRetris.AddNew
                        Call tableLrRetris_PutBuffer(recLrRetris)
                        tableLrRetris.Update
    Case "Update      "
                        tableLrRetris.Edit
                        Call tableLrRetris_PutBuffer(recLrRetris)
                        tableLrRetris.Update
    Case "Delete      "
                        tableLrRetris.Delete
    Case Else
                        Error 9999
End Select


Exit Function

tableLrRetrisUpdate_Error:
'---------------------------------------------------------
    tableLrRetris_Update = Err
    Resume tableLrRetrisUpdate_End

tableLrRetrisUpdate_End:

End Function








'-----------------------------------------------------
Sub dbLrRetris_Error(recLrRetris As typeLrRetris)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Key : " & Trim(recLrRetris.RFBENF) & " : " & Trim(recLrRetris.DTCENT1) & " : " & Trim(recLrRetris.CDCPCO) & Chr$(13)

Select Case mId$(recLrRetris.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recLrRetris.Err: I = vbCritical
End Select

MsgBox Msg, I, "mdbLrRetris.bas :  ( " & Trim(recLrRetris.obj) & " : " & Trim(recLrRetris.Method) & " ) "

End Sub

'-----------------------------------------------------
Function dbLrRetris_Read(recLrRetris As typeLrRetris)
'-----------------------------------------------------

dbLrRetris_Read = Null

recLrRetris.Err = tableLrRetris_Read(recLrRetris)
If recLrRetris.Err > 0 Then

    If recLrRetris.Err < 9990 Or recLrRetris.Err >= 9999 Then
        Call dbLrRetris_Error(recLrRetris)
        dbLrRetris_Read = recLrRetris.Err
    End If
End If

End Function

'-----------------------------------------------------
Function dbLrRetris_ReadZ(recLrRetris As typeLrRetris)
'-----------------------------------------------------

dbLrRetris_ReadZ = Null

recLrRetris.Err = tableLrRetris_Read(recLrRetris)
If recLrRetris.Err > 0 Then
    dbLrRetris_ReadZ = recLrRetris.Err
    recLrRetris.MT01 = 0
    recLrRetris.MT02 = 0
    recLrRetris.MT03 = 0
    recLrRetris.MT04 = 0
    recLrRetris.MT05 = 0
    recLrRetris.MT06 = 0
    recLrRetris.MT07 = 0
    recLrRetris.MT08 = 0
    recLrRetris.MT09 = 0
    recLrRetris.MT10 = 0
    recLrRetris.MT11 = 0
    recLrRetris.MT12 = 0
    recLrRetris.MT13 = 0
    recLrRetris.MT14 = 0
    recLrRetris.MT15 = 0
    recLrRetris.MT16 = 0
    recLrRetris.MT17 = 0
    recLrRetris.MT18 = 0
    recLrRetris.MT19 = 0
    recLrRetris.MT20 = 0
    recLrRetris.MT21 = 0
    recLrRetris.MT22 = 0
    recLrRetris.MT23 = 0
    recLrRetris.MT24 = 0
    recLrRetris.MT25 = 0
    recLrRetris.MTTOTAL = 0
End If

End Function


'-----------------------------------------------------
Function dbLrRetris_Update(recLrRetris As typeLrRetris)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
'$$$BeginTrans

dbLrRetris_Update = Null


recLrRetris.Err = tableLrRetris_Update(recLrRetris)

If recLrRetris.Err <> 0 Then
    Call dbLrRetris_Error(recLrRetris)
    dbLrRetris_Update = recLrRetris.Err
'$$$    Rollback
    Exit Function
End If

'$$$CommitTrans


'=====================================================
End Function


