Attribute VB_Name = "srvBdfG"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recBdfGLen = 350        ' 34 + 316


Type typeBdfG
    obj      As String * 12
    Method     As String * 12
    Err        As String * 10
    BdfG1      As String * 1
    BdfG2      As String * 5
    BdfG3      As String * 5
    BdfG4      As String * 1
    BdfG5      As String * 40
    BdfG6      As String * 20
    BdfG7      As String * 5
    BdfG8      As String * 5
    BdfG9      As String * 5
    BdfG10     As String * 5
    BdfG11     As String * 5
    BdfG12     As String * 5
    BdfG13     As String * 4
    BdfG14A    As String * 32
    BdfG14B    As String * 32
    BdfG14C    As String * 32
    BdfG14CP   As String * 5
    BdfG14D    As String * 27
    BdfG15     As String * 11
    BdfG16     As String * 8
    BdfG17     As String * 8
    BdfG18     As String * 8
    BdfG19     As String * 5
    BdfG20     As String * 24
    BdfG21     As String * 1
    BdfG22     As String * 4
    BdfG23     As String * 13
  
End Type
    

Public arrBdfG() As typeBdfG
Public arrBdfGNb As Integer
Public arrBdfGNbMax As Integer
Public arrBdfGIndex As Integer
Public arrBdfGsuite As Boolean
'-----------------------------------------------------
Public Function SrvBdfGMon(recbdfg As typeBdfG)
'-----------------------------------------------------

SrvBdfGMon = "?"

Select Case recbdfg.Method
    Case "SeekP0      "
                SrvBdfGMon = srvBdfGSeek(recbdfg)
    Case "SnapL1      ", "SnapL1+     ", "SnapP0+     ", _
         "SnapL2      ", "SnapL2+     ", "SnapP0      "
                SrvBdfGMon = srvBdfGSnap(recbdfg)
    Case Else
                recbdfg.Err = recbdfg.Method
                Call srvBdfGError(recbdfg)
                SrvBdfGMon = recbdfg.Err
End Select

End Function

'-----------------------------------------------------
Sub srvBdfGError(recbdfg As typeBdfG)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "code guichet (Banque de France) " & Chr$(10) & Chr$(13)

Select Case Mid$(recbdfg.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recbdfg.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : BdfG.bas  ( " _
                & Trim(recbdfg.obj) & " : " & Trim(recbdfg.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvBdfGGetBuffer(recbdfg As typeBdfG)
'---------------------------------------------------------
Dim K As Integer
srvBdfGGetBuffer = Null
recbdfg.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recbdfg.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recbdfg.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recbdfg.Err = Space$(10) Then
    recbdfg.BdfG1 = Mid$(MsgTxt, K + 1, 1)
    recbdfg.BdfG2 = Mid$(MsgTxt, K + 2, 5)
    recbdfg.BdfG3 = Mid$(MsgTxt, K + 7, 5)
    recbdfg.BdfG4 = Mid$(MsgTxt, K + 12, 1)
    recbdfg.BdfG5 = Mid$(MsgTxt, K + 13, 40)
    recbdfg.BdfG6 = Mid$(MsgTxt, K + 53, 20)
    recbdfg.BdfG7 = Mid$(MsgTxt, K + 73, 5)
    recbdfg.BdfG8 = Mid$(MsgTxt, K + 78, 5)
    recbdfg.BdfG9 = Mid$(MsgTxt, K + 83, 5)
    recbdfg.BdfG10 = Mid$(MsgTxt, K + 88, 5)
    recbdfg.BdfG11 = Mid$(MsgTxt, K + 93, 5)
    recbdfg.BdfG12 = Mid$(MsgTxt, K + 98, 5)
    recbdfg.BdfG13 = Mid$(MsgTxt, K + 103, 4)
    recbdfg.BdfG14A = Mid$(MsgTxt, K + 107, 32)
    recbdfg.BdfG14B = Mid$(MsgTxt, K + 139, 32)
    recbdfg.BdfG14C = Mid$(MsgTxt, K + 171, 32)
    recbdfg.BdfG14CP = Mid$(MsgTxt, K + 203, 5)
    recbdfg.BdfG14D = Mid$(MsgTxt, K + 208, 27)
    recbdfg.BdfG15 = Mid$(MsgTxt, K + 235, 11)
    recbdfg.BdfG16 = Mid$(MsgTxt, K + 246, 8)
    recbdfg.BdfG17 = Mid$(MsgTxt, K + 254, 8)
    recbdfg.BdfG18 = Mid$(MsgTxt, K + 262, 8)
    recbdfg.BdfG19 = Mid$(MsgTxt, K + 270, 5)
    recbdfg.BdfG20 = Mid$(MsgTxt, K + 275, 24)
    recbdfg.BdfG21 = Mid$(MsgTxt, K + 299, 1)
    recbdfg.BdfG22 = Mid$(MsgTxt, K + 300, 4)
    recbdfg.BdfG23 = Mid$(MsgTxt, K + 304, 13)
Else
    srvBdfGGetBuffer = recbdfg.Err
End If

MsgTxtIndex = MsgTxtIndex + recBdfGLen

End Function

'---------------------------------------------------------
Private Sub srvBdfGPutBuffer(recbdfg As typeBdfG)
'---------------------------------------------------------
Dim K As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recbdfg.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recbdfg.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 1) = recbdfg.BdfG1
Mid$(MsgTxt, K + 2, 5) = recbdfg.BdfG2
Mid$(MsgTxt, K + 7, 5) = recbdfg.BdfG3
Mid$(MsgTxt, K + 12, 1) = recbdfg.BdfG4
Mid$(MsgTxt, K + 13, 40) = recbdfg.BdfG5
Mid$(MsgTxt, K + 53, 20) = recbdfg.BdfG6
Mid$(MsgTxt, K + 73, 5) = recbdfg.BdfG7
Mid$(MsgTxt, K + 78, 5) = recbdfg.BdfG8
Mid$(MsgTxt, K + 83, 5) = recbdfg.BdfG9
Mid$(MsgTxt, K + 88, 5) = recbdfg.BdfG10
Mid$(MsgTxt, K + 93, 5) = recbdfg.BdfG11
Mid$(MsgTxt, K + 98, 5) = recbdfg.BdfG12
Mid$(MsgTxt, K + 103, 4) = recbdfg.BdfG13
Mid$(MsgTxt, K + 107, 32) = recbdfg.BdfG14A
Mid$(MsgTxt, K + 139, 32) = recbdfg.BdfG14B
Mid$(MsgTxt, K + 171, 32) = recbdfg.BdfG14C
Mid$(MsgTxt, K + 203, 5) = recbdfg.BdfG14CP
Mid$(MsgTxt, K + 208, 27) = recbdfg.BdfG14D
Mid$(MsgTxt, K + 235, 11) = recbdfg.BdfG15
Mid$(MsgTxt, K + 246, 8) = recbdfg.BdfG16
Mid$(MsgTxt, K + 254, 8) = recbdfg.BdfG17
Mid$(MsgTxt, K + 262, 8) = recbdfg.BdfG18
Mid$(MsgTxt, K + 270, 5) = recbdfg.BdfG19
Mid$(MsgTxt, K + 275, 24) = recbdfg.BdfG20
Mid$(MsgTxt, K + 299, 1) = recbdfg.BdfG21
Mid$(MsgTxt, K + 300, 4) = recbdfg.BdfG22
Mid$(MsgTxt, K + 304, 13) = recbdfg.BdfG23
MsgTxtLen = MsgTxtLen + recBdfGLen
End Sub



'---------------------------------------------------------
Private Function srvBdfGSeek(recbdfg As typeBdfG)
'---------------------------------------------------------

srvBdfGSeek = "?"
MsgTxtLen = 0
Call srvBdfGPutBuffer(recbdfg)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvBdfGGetBuffer(recbdfg)) Then
        srvBdfGSeek = Null
    Else
        Call srvBdfGError(recbdfg)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvBdfGSnap(recbdfg As typeBdfG)
'---------------------------------------------------------
srvBdfGSnap = "?"
MsgTxtLen = 0
Call srvBdfGPutBuffer(recbdfg)
Call srvBdfGPutBuffer(arrBdfG(0))
If IsNull(SndRcv()) Then
    srvBdfGSnap = Null
    MsgTxtIndex = 0 'recBdfGLen
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvBdfGGetBuffer(recbdfg)) Then
            Call arrBdfGAddItem(recbdfg)
        Else
            arrBdfGsuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recBdfGInit(recbdfg As typeBdfG)
'---------------------------------------------------------
 MsgTxt = Space$(recBdfGLen)
 MsgTxtIndex = 0
 Call srvBdfGGetBuffer(recbdfg)
 recbdfg.obj = "SRVBDFG      "
End Sub

'-----------------------------------------------------
Public Function srvBdfGFind(recbdfg As typeBdfG)
'-----------------------------------------------------
Dim I As Integer, X As Variant

X = "?"

For I = 1 To arrBdfGNb
    If recbdfg.BdfG2 = arrBdfG(I).BdfG2 _
    And recbdfg.BdfG3 = arrBdfG(I).BdfG3 Then
        X = Null
        Exit For
'!!!!!!!========
    End If
Next I

If Not IsNull(X) Then
    I = arrBdfGNb + 1
    recbdfg.Method = "SnapP0      "
    arrBdfG(0) = recbdfg
    X = srvBdfGSnap(recbdfg)
    If I > arrBdfGNb Then X = "inconnu"
End If

If IsNull(X) Then
    arrBdfGIndex = I
    recbdfg = arrBdfG(I)
End If

srvBdfGFind = X

End Function



'---------------------------------------------------------
Public Sub arrBdfGAddItem(recbdfg As typeBdfG)
'---------------------------------------------------------

arrBdfGNb = arrBdfGNb + 1
If arrBdfGNb > arrBdfGNbMax Then
    arrBdfGNbMax = arrBdfGNbMax + 10
    ReDim Preserve arrBdfG(arrBdfGNbMax)
End If
            
arrBdfG(arrBdfGNb) = recbdfg

End Sub


