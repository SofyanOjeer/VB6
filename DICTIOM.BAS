Attribute VB_Name = "mdbDictio"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim tableDictio As Recordset

Type typeDictio
    obj        As String * 12
    Method     As String * 12
    Err        As String * 10
    DicRub     As String * 3
    DicCode    As String * 20
    DicLib     As String * 40
    DicTxt     As String * 235
    DicAmj     As String * 8
 
End Type
    
Type typeDevise
    obj        As String * 12
    Method     As String * 12
    Err        As String * 10
    DevX       As String * 3
    DevCode    As Integer
    DevLib     As String * 20
    Cours   As Double
    DevAmj     As String * 8
    maxD    As String * 1
End Type

'99-01-20 Public CV As typeDevise
Public XDevise As typeDevise
Public XDictio As typeDictio

'---------------------------------------------------------
Public Function DicLib(ByVal Vrub As Variant, ByVal Vcode As Variant) As String
'---------------------------------------------------------
Dim recdictio As typeDictio

DicLib = "-"
If Not IsNumeric(Vrub) Then
    Error = 1
Else
    recdictio.Method = "Seek=       "
    recdictio.DicRub = Vrub
    recdictio.DicCode = Vcode
    If IsNull(dbDictioRead(recdictio)) Then
        If recdictio.Err = 0 Then DicLib = recdictio.DicLib
    Else
        Error = 2
    End If
End If

Exit Function

'---------------------------------------------------------
DicLibError:
'---------------------------------------------------------
    DicLib = " ? rub: " & Vrub & " code: " & Vcode

    Resume DicLibEnd

DicLibEnd:

End Function

'-----------------------------------------------------
Function dbDictioRead(recdictio As typeDictio)
'-----------------------------------------------------

dbDictioRead = Null

recdictio.Err = tableDictioRead(recdictio)
If recdictio.Err > 0 Then

    If recdictio.Err < 9990 Or recdictio.Err >= 9999 Then
        Call dbDictioError(recdictio)
        dbDictioRead = recdictio.Err
    End If
End If

End Function

'-----------------------------------------------------
Sub dbDictioError(recdictio As typeDictio)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Select Case recdictio.Err
    Case 9922
        Msg = "Existe déjà"
        I = 48
    Case 9923
        Msg = "N'existe pas"
        I = 48
    Case Else
        Msg = Error$(recdictio.Err) & Chr$(10) & Chr$(13) & "( " & recdictio.Method & " : Error Code : " & recdictio.Err & " )"
        I = 16
End Select

Title = recdictio.obj & " : " & Format$(recdictio.DicRub, " ###") _
                        & " : " & recdictio.DicCode

MsgBox Msg, I, Title

End Sub


'-----------------------------------------------------
Function dbDictioUpdate(recdictio As typeDictio)
'-----------------------------------------------------
Dim K As Integer

'=====================================================
'$$$BeginTrans

dbDictioUpdate = Null


recdictio.Err = tableDictioUpdate(recdictio)

If recdictio.Err <> 0 Then
    Call dbDictioError(recdictio)
    dbDictioUpdate = recdictio.Err
'$$$    Rollback
    Exit Function
End If

'$$$CommitTrans


'=====================================================
End Function
'-----------------------------------------------------
Sub tableDictio_Open()
'-----------------------------------------------------

Set tableDictio = MDB.OpenRecordset("Dictiop1")
tableDictio.Index = "PrimaryKey"

End Sub
'-----------------------------------------------------
Sub tableDictio_Close()
'-----------------------------------------------------

tableDictio.Close

End Sub

'---------------------------------------------------------
Public Sub tableDictioGetBuffer(recdictio As typeDictio)
'---------------------------------------------------------

recdictio.DicRub = Format$(tableDictio("DicRub"), "000")
recdictio.DicCode = tableDictio("DicCode")
recdictio.DicTxt = tableDictio("DicTxt")
recdictio.DicLib = tableDictio("DicLib")
recdictio.DicAmj = tableDictio("DicAmj")

End Sub

'---------------------------------------------------------
Public Sub tableDictioPutBuffer(recdictio As typeDictio)
'---------------------------------------------------------

tableDictio("DicRub") = Val(recdictio.DicRub)
tableDictio("DicCode") = recdictio.DicCode
tableDictio("DicLib") = recdictio.DicLib
tableDictio("DicTxt") = recdictio.DicTxt
tableDictio("DicAmj") = recdictio.DicAmj

End Sub

'---------------------------------------------------------
Public Function tableDictioRead(recdictio As typeDictio) As Integer
'---------------------------------------------------------

On Error GoTo tableDictioRead_Error
tableDictioRead = 0


Select Case recdictio.Method
    Case "AddNew      "
                        tableDictio.Seek "=", recdictio.DicRub, recdictio.DicCode
                        If Not tableDictio.NoMatch Then
                            Error 9922
                        End If
    Case "Update      ", "Delete      "
                        tableDictio.Seek "=", recdictio.DicRub, recdictio.DicCode
                        If tableDictio.NoMatch Then
                            Error 9923
                        End If
    Case "Seek=       "
                        tableDictio.Seek "=", recdictio.DicRub, recdictio.DicCode
                        If tableDictio.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>       "
                        tableDictio.Seek ">", recdictio.DicRub, recdictio.DicCode
                        If tableDictio.NoMatch Then
                            Error 9998
                        End If
    Case "Seek>=      "
                        tableDictio.Seek ">=", recdictio.DicRub, recdictio.DicCode
                        If tableDictio.NoMatch Then
                            Error 9998
                        End If
    Case "Seek<       "
                        tableDictio.Seek "<", recdictio.DicRub, recdictio.DicCode
                        If tableDictio.NoMatch Then
                            Error 9998
                        End If
    Case "Seek<=      "
                        tableDictio.Seek "<=", recdictio.DicRub, recdictio.DicCode
                        If tableDictio.NoMatch Then
                            Error 9998
                        End If
    Case "MoveNext    "
                        tableDictio.MoveNext
                        If tableDictio.EOF Then
                            Error 9996
                        End If
    Case "MovePrevious"
                        tableDictio.MovePrevious
                        If tableDictio.BOF Then
                            Error 9997
                        End If
    Case "MoveFirst   "
                        tableDictio.MoveFirst
                        If tableDictio.NoMatch Then
                            Error 9998
                        End If
    Case "MoveLast    "
                        tableDictio.MoveLast
                        If tableDictio.NoMatch Then
                            Error 9998
                        End If
    Case Else
                        Error 9999
End Select

If recdictio.Method <> "AddNew      " Then
    Call tableDictioGetBuffer(recdictio)
End If

Exit Function

'---------------------------------------------------------
tableDictioRead_Error:
'---------------------------------------------------------

    tableDictioRead = Err
    Resume tableDictioRead_End

tableDictioRead_End:

End Function


'---------------------------------------------------------
Public Function tableDictioUpdate(recdictio As typeDictio) As Integer
'---------------------------------------------------------

On Error GoTo tableDictioUpdate_Error
tableDictioUpdate = 0

Select Case recdictio.Method

    Case "AddNew      "
                        tableDictio.AddNew
                        Call tableDictioPutBuffer(recdictio)
                        tableDictio.Update
    Case "Update      "
                        tableDictio.Edit
                        Call tableDictioPutBuffer(recdictio)
                        tableDictio.Update
    Case "Delete      "
                        tableDictio.Delete
    Case Else
                        Error 9999
End Select


Exit Function

tableDictioUpdate_Error:
'---------------------------------------------------------
    tableDictioUpdate = Err
    Resume tableDictioUpdate_End

tableDictioUpdate_End:

End Function






'-----------------------------------------------------
Public Function dbDeviseSeek(recDevise As typeDevise)
'-----------------------------------------------------

If recDevise.Method = "DevCode     " Then
    XDictio.DicRub = 888
    XDictio.DicCode = Format$(recDevise.DevCode, "000")
Else
    XDictio.DicRub = 889
    XDictio.DicCode = Format(recDevise.DevX, "@@@")
End If

recDevise.Method = "Seek=       "
dbDeviseSeek = dbDeviseRead(recDevise)

End Function

'-----------------------------------------------------
Public Function DevX(ByVal DevCode) As String
'-----------------------------------------------------

If Val(DevCode) <> XDevise.DevCode Then
    XDevise.Method = "DevCode"
    XDevise.DevCode = Val(DevCode)
    Call dbDeviseSeek(XDevise)
End If
DevX = XDevise.DevX

End Function
'-----------------------------------------------------
Public Function DevCode(ByVal DevX) As Integer
'-----------------------------------------------------


If DevX <> XDevise.DevX Then
    XDevise.Method = "DevX"
    XDevise.DevX = Trim(DevX)
    Call dbDeviseSeek(XDevise)
End If
DevCode = XDevise.DevCode

End Function


'-----------------------------------------------------
Public Function LstDictio(DicRub As Integer, lstBox As ListBox, Optional kDisplay) As Integer
'--------------------------------------------------

Dim X, mDicrub As Integer, blnDisplay As Boolean

blnDisplay = False
If Not IsMissing(kDisplay) Then
    If kDisplay = "DicLib" Then blnDisplay = True
End If
LstDictio = DicRub
lstBox.Clear


XDictio.Method = "Seek>=      "
XDictio.DicCode = String$(20, Chr$(0))
XDictio.DicRub = DicRub
mDicrub = XDictio.DicRub

X = dbDictioRead(XDictio)
XDictio.Method = "MoveNext    "
Do While XDictio.Err = 0 _
     And XDictio.DicRub = mDicrub

    If blnDisplay Then
        X = Trim(XDictio.DicLib) & " := " & Trim(XDictio.DicCode)
        If DicRub = 19 Then
                X = X & "    =: " & mId$(XDictio.DicTxt, 7, 2)
        End If
    Else
        X = Trim(XDictio.DicCode) & Chr$(9) & XDictio.DicLib
    End If
    
    lstBox.AddItem X

X = dbDictioRead(XDictio)

Loop

End Function

'-----------------------------------------------------
Public Sub cboDictio(DicRub As Integer, cbo As ComboBox, lK2_len As Integer)
'--------------------------------------------------

Dim X, mDicrub As Integer, blnDisplay As Boolean

'blnDisplay = False
'If Not IsMissing(kDisplay) Then
'    If kDisplay = "DicLib" Then blnDisplay = True
'End If
'LstDictio = DicRub
cbo.Clear


XDictio.Method = "Seek>=      "
XDictio.DicCode = String$(20, Chr$(0))
XDictio.DicRub = DicRub
mDicrub = XDictio.DicRub

X = dbDictioRead(XDictio)
XDictio.Method = "MoveNext    "
Do While XDictio.Err = 0 _
     And XDictio.DicRub = mDicrub

    cbo.AddItem mId$(XDictio.DicCode, 1, lK2_len) & " " & XDictio.DicLib

    X = dbDictioRead(XDictio)

Loop

End Sub


'-----------------------------------------------------
Public Function dbDeviseRead(recDevise As typeDevise)
'-----------------------------------------------------
Dim mDicrub As String

mDicrub = XDictio.DicRub
XDictio.Method = recDevise.Method
dbDeviseRead = dbDictioRead(XDictio)

If XDictio.DicRub = mDicrub Then
    recDevise.Err = XDictio.Err
Else
    recDevise.Err = 9995
End If

If Trim(recDevise.Err) = "0" Then
    If XDictio.DicRub = 888 Then
        recDevise.DevCode = mId$(XDictio.DicCode, 1, 3)
        recDevise.DevX = mId$(XDictio.DicTxt, 1, 3)
    Else
        recDevise.DevX = mId$(XDictio.DicCode, 1, 3)

        recDevise.DevCode = mId$(XDictio.DicTxt, 1, 3)
    End If
    recDevise.DevLib = XDictio.DicLib
    recDevise.Cours = Val(mId$(XDictio.DicTxt, 4, 13))
    recDevise.DevAmj = mId$(XDictio.DicTxt, 17, 8)
Else
    recDevise.DevCode = 0
    recDevise.DevLib = "? " & recDevise.DevCode
    recDevise.DevX = ""
    recDevise.Cours = 1
    recDevise.DevAmj = "00000000"
End If
If recDevise.DevCode = 5 Or recDevise.DevCode = 732 _
Or recDevise.DevCode = 10 Or recDevise.DevCode = 11 _
Or recDevise.DevCode = 19 Or recDevise.DevCode = 20 _
Or recDevise.DevCode = 9 Then
    recDevise.maxD = 0
Else
    recDevise.maxD = 2
End If


End Function

Public Function curCV(curD1 As Currency, Cours As Double, maxD As String) As Currency
Dim curX As Currency
curX = curD1 / Cours
Select Case maxD
    Case 0: curCV = Fix(curX)
    Case Else: curCV = Fix(curX * 100) / 100
End Select
End Function
Public Function curCV_Excès(curD1 As Currency, Cours As Double, maxD As String) As Currency
Dim curX As Currency, curX2 As Currency
curX = curD1 / Cours
Select Case maxD
    Case 0: curX2 = Fix(curX)
            If curX = curX2 Then
                curCV_Excès = curX2
            Else
                curCV_Excès = curX2 + 1
            End If
    Case Else: curX2 = Fix(curX * 100) / 100
            If curX = curX2 Then
                curCV_Excès = curX2
            Else
                curCV_Excès = curX2 + 0.01
            End If
End Select
End Function
Public Function curChange(curD1 As Currency, Cours As Double, maxD As String) As Currency
Dim curX As Currency
curX = curD1 * Cours
Select Case maxD
    Case 0: curChange = Fix(curX)
    Case Else: curChange = Fix(curX * 100) / 100
End Select
End Function

Public Function curCVInv(curD2 As Currency, Cours As Double, maxD As String) As Currency
Dim curD1 As Currency, curD2b As Currency, I As Integer
curCVInv = 0
curD1 = Fix(curD2 * Cours * 100) / 100
For I = 1 To 100
    curD2b = curCV(curD1, Cours, maxD)
    If curD2b = curD2 Then curCVInv = curD1: Exit For
    If curD2b < curD2 Then
        curD1 = curD1 + 0.01
    Else
        curD1 = curD1 - 0.01
    End If
Next I
    
End Function
