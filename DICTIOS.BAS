Attribute VB_Name = "srvDictio"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recDictioLen = 321       ' 34 + 287


Public arrDictio() As typeDictio
Public arrDictioNb As Integer
Public arrDictioNbMax As Integer
Public arrDictioIndex As Integer
Public arrDictioSuite As Boolean
'-----------------------------------------------------
Public Function srvDictioFind(recdictio As typeDictio)
'-----------------------------------------------------
Dim I As Integer, X As Variant

X = "?"

For I = 1 To arrDictioNb
    If recdictio.DicRub = arrDictio(I).DicRub _
    And recdictio.DicCode = arrDictio(I).DicCode Then
        X = Null
        Exit For
'!!!!!!!========
    End If
Next I

If Not IsNull(X) Then
    I = arrDictioNb + 1
    recdictio.Method = "SnapP0      "
    arrDictio(0) = recdictio
    X = srvDictioSnap(recdictio)
    If I > arrDictioNb Then X = "inconnu"
End If

If IsNull(X) Then
    arrDictioIndex = I
    recdictio = arrDictio(I)
End If

srvDictioFind = X

End Function



'-----------------------------------------------------
Public Function srvDictioMon(recdictio As typeDictio)
'-----------------------------------------------------

arrDictioSuite = False
Select Case Trim(recdictio.Method)
    Case "SeekP1"
                srvDictioMon = srvDictioSeek(recdictio)
    Case "SnapP1", "SnapP1+", "DevCours"
              srvDictioMon = srvDictioSnap(recdictio)
    Case Else
                recdictio.Err = recdictio.Method
                Call srvDictioerror(recdictio)
                srvDictioMon = recdictio.Err
End Select

End Function

'-----------------------------------------------------
Sub srvDictioerror(recdictio As typeDictio)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Dictio" & Chr$(10) & Chr$(13)

Select Case mId$(recdictio.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recdictio.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : Cpt.bas  ( " _
                & Trim(recdictio.obj) & " : " & Trim(recdictio.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvDictioGetBuffer(recdictio As typeDictio)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvDictioGetBuffer = Null
recdictio.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recdictio.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recdictio.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recdictio.Err = Space$(10) Then
    recdictio.DicRub = mId$(MsgTxt, K + 1, 3)
    recdictio.DicCode = mId$(MsgTxt, K + 4, 20)
    
    If recdictio.DicRub = "013" Or recdictio.DicRub = "524" Then
        I = 40
    Else
        I = 25
    End If
    recdictio.DicLib = mId$(MsgTxt, K + 24, I)
    recdictio.DicTxt = mId$(MsgTxt, K + 24 + I, 256 - I)
    recdictio.DicAmj = mId$(MsgTxt, K + 284, 4) _
                    & mId$(MsgTxt, K + 282, 2) _
                    & mId$(MsgTxt, K + 280, 2)
Else
    srvDictioGetBuffer = recdictio.Err
End If

MsgTxtIndex = MsgTxtIndex + recDictioLen

End Function

'---------------------------------------------------------
Private Sub srvDictioPutBuffer(recdictio As typeDictio)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recdictio.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recdictio.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 3) = recdictio.DicRub
Mid$(MsgTxt, K + 4, 20) = recdictio.DicCode
If recdictio.DicRub = "013" Then
    I = 40
Else
    I = 25
End If
Mid$(MsgTxt, K + 24, I) = recdictio.DicLib
Mid$(MsgTxt, K + 24 + I, 256 - I) = recdictio.DicTxt
Mid$(MsgTxt, K + 280, 8) = mId$(recdictio.DicAmj, 7, 2) _
                         & mId$(recdictio.DicAmj, 5, 2) _
                         & mId$(recdictio.DicAmj, 1, 4)
MsgTxtLen = MsgTxtLen + recDictioLen
End Sub



'---------------------------------------------------------
Private Function srvDictioSeek(recdictio As typeDictio)
'---------------------------------------------------------

srvDictioSeek = "?"
MsgTxtLen = 0
Call srvDictioPutBuffer(recdictio)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvDictioGetBuffer(recdictio)) Then
        srvDictioSeek = Null
    Else
        Call srvDictioerror(recdictio)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvDictioSnap(recdictio As typeDictio)
'---------------------------------------------------------
srvDictioSnap = "?"
MsgTxtLen = 0
Call srvDictioPutBuffer(recdictio)
Call srvDictioPutBuffer(arrDictio(0))
If IsNull(SndRcv()) Then
    srvDictioSnap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvDictioGetBuffer(recdictio)) Then
            Call arrDictioAddItem(recdictio)
            arrDictioSuite = True
        Else
            arrDictioSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recDictioInit(recdictio As typeDictio)
'---------------------------------------------------------
 MsgTxt = Space$(recDictioLen)
 MsgTxtIndex = 0
 Call srvDictioGetBuffer(recdictio)
 recdictio.obj = "SRVDICTIO   "
End Sub

'---------------------------------------------------------
Public Sub arrDictioAddItem(recdictio As typeDictio)
'---------------------------------------------------------
          
arrDictioNb = arrDictioNb + 1
    
If arrDictioNb > arrDictioNbMax Then
    arrDictioNbMax = arrDictioNbMax + 10
    ReDim Preserve arrDictio(arrDictioNbMax)
End If
            
arrDictio(arrDictioNb) = recdictio
End Sub
