Attribute VB_Name = "srvAccAut"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recAccAutLen = 112         ' 34 + 78

Type typeAccAut
    obj      As String * 12
    Method     As String * 12
    Err        As String * 10
    AccAutId      As String * 10
    AccAutK1      As String * 10
    AccAutK2      As String * 10
    AccAutTxt     As String * 20
    AccAutDD      As String * 8
    AccAutHD      As String * 6
    AccAutDF      As String * 8
    AccAutHF      As String * 6
 
End Type
    

Public arrAccAut() As typeAccAut
Public arrAccAutNb As Integer
Public arrAccAutNbMax As Integer
Public arrAccAutIndex As Integer
Public arrAccAutsuite As Boolean
'-----------------------------------------------------
Function Update(recAccAut As typeAccAut)
'-----------------------------------------------------

Update = "?"

MsgTxtLen = 0
Call PutBuffer(recAccAut)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    recAccAut.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
    If Trim(recAccAut.Err) <> "" Then
        Call ErrorX(recAccAut)
        Update = recAccAut.Err
        Exit Function
    Else
        Update = Null
    End If
Else
    recAccAut.Err = "srv"
End If

End Function


'-----------------------------------------------------
Public Function Monitor(recAccAut As typeAccAut)
'-----------------------------------------------------

Monitor = "?"
arrAccAutsuite = False

Select Case recAccAut.Method
    Case "SeekP0      "
                Monitor = srvAccAut.SeekX(recAccAut)
    Case "SnapL2      ", "SnapL2+     ", "SnapP0+     ", _
         "SnapP0      "
                Monitor = srvAccAut.SnapX(recAccAut)
    Case Else
                recAccAut.Err = recAccAut.Method
                Call ErrorX(recAccAut)
                Monitor = recAccAut.Err
End Select

End Function

'-----------------------------------------------------
Sub ErrorX(recAccAut As typeAccAut)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Accès,Autorisation " & Chr$(10) & Chr$(13)

Select Case Mid$(recAccAut.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recAccAut.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : AccAut.bas  ( " _
                & Trim(recAccAut.obj) & " : " & Trim(recAccAut.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function GetBuffer(recAccAut As typeAccAut)
'---------------------------------------------------------
Dim K As Integer
GetBuffer = Null
recAccAut.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recAccAut.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recAccAut.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recAccAut.Err = Space$(10) Then
    recAccAut.AccAutId = Mid$(MsgTxt, K + 1, 10)
    recAccAut.AccAutK1 = Mid$(MsgTxt, K + 11, 10)
    recAccAut.AccAutK2 = Mid$(MsgTxt, K + 21, 10)
    recAccAut.AccAutTxt = Mid$(MsgTxt, K + 31, 20)
    recAccAut.AccAutDD = Format$(Val(Mid$(MsgTxt, K + 51, 8)), "00000000")
    recAccAut.AccAutHD = Format$(Val(Mid$(MsgTxt, K + 59, 6)), "000000")
    recAccAut.AccAutDF = Format$(Val(Mid$(MsgTxt, K + 65, 8)), "00000000")
    recAccAut.AccAutHF = Format$(Val(Mid$(MsgTxt, K + 73, 6)), "000000")
Else
    GetBuffer = recAccAut.Err
End If

MsgTxtIndex = MsgTxtIndex + recAccAutLen

End Function

'---------------------------------------------------------
Private Sub PutBuffer(recAccAut As typeAccAut)
'---------------------------------------------------------
Dim K As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recAccAut.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recAccAut.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 10) = recAccAut.AccAutId
Mid$(MsgTxt, K + 11, 10) = recAccAut.AccAutK1
Mid$(MsgTxt, K + 21, 10) = recAccAut.AccAutK2
Mid$(MsgTxt, K + 31, 20) = recAccAut.AccAutTxt
Mid$(MsgTxt, K + 51, 8) = Format$(recAccAut.AccAutDD, "00000000")
Mid$(MsgTxt, K + 59, 6) = Format$(recAccAut.AccAutHD, "000000")
Mid$(MsgTxt, K + 65, 8) = Format$(recAccAut.AccAutDF, "00000000")
Mid$(MsgTxt, K + 73, 6) = Format$(recAccAut.AccAutHF, "000000")

MsgTxtLen = MsgTxtLen + recAccAutLen
End Sub



'---------------------------------------------------------
Private Function SeekX(recAccAut As typeAccAut)
'---------------------------------------------------------

SeekX = "?"
MsgTxtLen = 0
Call srvAccAut.PutBuffer(recAccAut)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvAccAut.GetBuffer(recAccAut)) Then
        SeekX = Null
    Else
        Call ErrorX(recAccAut)
    End If
End If

End Function

'---------------------------------------------------------
Private Function SnapX(recAccAut As typeAccAut)
'---------------------------------------------------------
SnapX = "?"
MsgTxtLen = 0
Call srvAccAut.PutBuffer(recAccAut)
Call srvAccAut.PutBuffer(arrAccAut(0))
If IsNull(SndRcv()) Then
   SnapX = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvAccAut.GetBuffer(recAccAut)) Then
            Call srvAccAut.AddItem(recAccAut)
            arrAccAutsuite = True
        Else
            arrAccAutsuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub Init(recAccAut As typeAccAut)
'---------------------------------------------------------
 MsgTxt = Space$(recAccAutLen)
 MsgTxtIndex = 0
 Call srvAccAut.GetBuffer(recAccAut)
 recAccAut.obj = "SRVACCAUT   "
End Sub

'---------------------------------------------------------
Public Sub AddItem(recAccAut As typeAccAut)
'---------------------------------------------------------
            
arrAccAutNb = arrAccAutNb + 1
            
If arrAccAutNb > arrAccAutNbMax Then
    arrAccAutNbMax = arrAccAutNbMax + 10
    ReDim Preserve arrAccAut(arrAccAutNbMax)
End If
            
arrAccAut(arrAccAutNb) = recAccAut

End Sub
