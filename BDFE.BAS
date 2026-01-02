Attribute VB_Name = "srvBdfE"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recBdfeLen = 271            ' 34 + 237

Type typeBdfE
    obj      As String * 12
    Method     As String * 12
    Err        As String * 10
    BdfE1      As String * 1
    BdfE2      As String * 5
    BdfE2B     As String * 5
    BdfE3      As String * 1
    BdfE4      As String * 40
    BdfE4B     As String * 10
    BdfE5      As String * 1
    BdfE6      As String * 2
    BdfE7      As String * 8
    BdfE8      As String * 8
    BdfE9      As String * 8
    BdfE10A    As String * 32
    BdfE10B    As String * 32
    BdfE10C    As String * 32
    BdfE10CP   As String * 5
    BdfE10D    As String * 27
    BdfE11     As String * 6
    BdfE12     As String * 5
    BdfE13     As String * 1
    BdfE14     As String * 8
   
End Type
    

Public arrBdfE() As typeBdfE
Public arrBdfENb As Integer
Public arrBdfENbMax As Integer
Public arrBdfEIndex As Integer
Public arrBdfEsuite As Boolean
'-----------------------------------------------------
Public Function srvBdfEFind(recBdfE As typeBdfE)
'-----------------------------------------------------
Dim I As Integer, X As Variant

X = "?"

For I = 1 To arrBdfENb
    If recBdfE.BdfE2 = arrBdfE(I).BdfE2 Then
        X = Null
        Exit For
'!!!!!!!========
    End If
Next I

If Not IsNull(X) Then
    I = arrBdfENb + 1
    recBdfE.Method = "SnapP0      "
    arrBdfE(0) = recBdfE
    X = srvBdfESnap(recBdfE)
    If I > arrBdfENb Then X = "inconnu"
End If

If IsNull(X) Then
    arrBdfEIndex = I
    recBdfE = arrBdfE(I)
End If

srvBdfEFind = X

End Function



'-----------------------------------------------------
Public Function SrvBdfEMon(recBdfE As typeBdfE)
'-----------------------------------------------------


Select Case recBdfE.Method
    Case "SeekP0      "
                SrvBdfEMon = srvBdfESeek(recBdfE)
    Case "SnapL1      ", "SnapL1+     ", "SnapP0+     ", _
         "SnapL2      ", "SnapL2+     ", "SnapP0      "
                SrvBdfEMon = srvBdfESnap(recBdfE)
    Case Else
                recBdfE.Err = recBdfE.Method
                Call srvBdfEError(recBdfE)
                SrvBdfEMon = recBdfE.Err
End Select

End Function

'-----------------------------------------------------
Sub srvBdfEError(recBdfE As typeBdfE)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "code établissement (Banque de France) " & Chr$(10) & Chr$(13)

Select Case Mid$(recBdfE.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recBdfE.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : BdfE.bas  ( " _
                & Trim(recBdfE.obj) & " : " & Trim(recBdfE.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvBdfEGetBuffer(recBdfE As typeBdfE)
'---------------------------------------------------------
Dim K As Integer
srvBdfEGetBuffer = Null
recBdfE.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recBdfE.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recBdfE.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recBdfE.Err = Space$(10) Then
    recBdfE.BdfE1 = Mid$(MsgTxt, K + 1, 1)
    recBdfE.BdfE2 = Mid$(MsgTxt, K + 2, 5)
    recBdfE.BdfE2B = Mid$(MsgTxt, K + 7, 5)
    recBdfE.BdfE3 = Mid$(MsgTxt, K + 12, 1)
    recBdfE.BdfE4 = Mid$(MsgTxt, K + 13, 40)
    recBdfE.BdfE4B = Mid$(MsgTxt, K + 53, 10)
    recBdfE.BdfE5 = Mid$(MsgTxt, K + 63, 1)
    recBdfE.BdfE6 = Mid$(MsgTxt, K + 64, 2)
    recBdfE.BdfE7 = Mid$(MsgTxt, K + 66, 8)
    recBdfE.BdfE8 = Mid$(MsgTxt, K + 74, 8)
    recBdfE.BdfE9 = Mid$(MsgTxt, K + 82, 8)
    recBdfE.BdfE10A = Mid$(MsgTxt, K + 90, 32)
    recBdfE.BdfE10B = Mid$(MsgTxt, K + 122, 32)
    recBdfE.BdfE10C = Mid$(MsgTxt, K + 154, 32)
    recBdfE.BdfE10CP = Mid$(MsgTxt, K + 186, 5)
    recBdfE.BdfE10D = Mid$(MsgTxt, K + 191, 27)
    recBdfE.BdfE11 = Mid$(MsgTxt, K + 218, 6)
    recBdfE.BdfE12 = Mid$(MsgTxt, K + 224, 5)
    recBdfE.BdfE13 = Mid$(MsgTxt, K + 229, 1)
    recBdfE.BdfE14 = Mid$(MsgTxt, K + 230, 8)
Else
    srvBdfEGetBuffer = recBdfE.Err
End If

MsgTxtIndex = MsgTxtIndex + recBdfeLen

End Function

'---------------------------------------------------------
Private Sub srvBdfEPutBuffer(recBdfE As typeBdfE)
'---------------------------------------------------------
Dim K As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recBdfE.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recBdfE.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 1) = recBdfE.BdfE1
Mid$(MsgTxt, K + 2, 5) = recBdfE.BdfE2
Mid$(MsgTxt, K + 7, 5) = recBdfE.BdfE2B
Mid$(MsgTxt, K + 12, 1) = recBdfE.BdfE3
Mid$(MsgTxt, K + 13, 40) = recBdfE.BdfE4
Mid$(MsgTxt, K + 53, 10) = recBdfE.BdfE4B
Mid$(MsgTxt, K + 63, 1) = recBdfE.BdfE5
Mid$(MsgTxt, K + 64, 2) = recBdfE.BdfE6
Mid$(MsgTxt, K + 66, 8) = recBdfE.BdfE7
Mid$(MsgTxt, K + 74, 8) = recBdfE.BdfE8
Mid$(MsgTxt, K + 82, 8) = recBdfE.BdfE9
Mid$(MsgTxt, K + 90, 32) = recBdfE.BdfE10A
Mid$(MsgTxt, K + 122, 32) = recBdfE.BdfE10B
Mid$(MsgTxt, K + 154, 32) = recBdfE.BdfE10C
Mid$(MsgTxt, K + 186, 5) = recBdfE.BdfE10CP
Mid$(MsgTxt, K + 191, 27) = recBdfE.BdfE10D
Mid$(MsgTxt, K + 218, 6) = recBdfE.BdfE11
Mid$(MsgTxt, K + 224, 5) = recBdfE.BdfE12
Mid$(MsgTxt, K + 229, 1) = recBdfE.BdfE13
Mid$(MsgTxt, K + 230, 8) = recBdfE.BdfE14

MsgTxtLen = MsgTxtLen + recBdfeLen
End Sub



'---------------------------------------------------------
Private Function srvBdfESeek(recBdfE As typeBdfE)
'---------------------------------------------------------

srvBdfESeek = "?"
MsgTxtLen = 0
Call srvBdfEPutBuffer(recBdfE)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvBdfEGetBuffer(recBdfE)) Then
        srvBdfESeek = Null
    Else
        Call srvBdfEError(recBdfE)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvBdfESnap(recBdfE As typeBdfE)
'---------------------------------------------------------
srvBdfESnap = "?"
MsgTxtLen = 0
Call srvBdfEPutBuffer(recBdfE)
Call srvBdfEPutBuffer(arrBdfE(0))
If IsNull(SndRcv()) Then
    srvBdfESnap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvBdfEGetBuffer(recBdfE)) Then
            Call arrBdfEAddItem(recBdfE)
        Else
            arrBdfEsuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recBdfEInit(recBdfE As typeBdfE)
'---------------------------------------------------------
 MsgTxt = Space$(recBdfeLen)
 MsgTxtIndex = 0
 Call srvBdfEGetBuffer(recBdfE)
 recBdfE.obj = "SRVBDFE      "
End Sub

'---------------------------------------------------------
Public Sub arrBdfEAddItem(recBdfE As typeBdfE)
'---------------------------------------------------------
          
arrBdfENb = arrBdfENb + 1
    
If arrBdfENb > arrBdfENbMax Then
    arrBdfENbMax = arrBdfENbMax + 10
    ReDim Preserve arrBdfE(arrBdfENbMax)
End If
            
arrBdfE(arrBdfENb) = recBdfE
End Sub
