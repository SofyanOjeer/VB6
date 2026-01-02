Attribute VB_Name = "srvEchellesFusion"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recEchellesFusionLen = 78 ' 34 + 44
Type typeEchellesFusion
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    DeviseOrigine          As String * 3
    CompteOrigine          As String * 11
    DeviseFusion           As String * 3
    CompteFusion           As String * 11
    AmjDébut               As String * 8
    AmjFin                 As String * 8
End Type
    
Public arrEchellesFusion() As typeEchellesFusion
Public arrEchellesFusionNb As Integer
Public arrEchellesFusionNbMax As Integer
Public arrEchellesFusionIndex As Integer
Public arrEchellesFusionSuite As Boolean

Public XEchellesFusion As typeEchellesFusion
Public Sub srvEchellesFusion_Load(Amj As String)
ReDim arrEchellesFusion(10): arrEchellesFusionNbMax = 10
recEchellesFusion_Init XEchellesFusion
XEchellesFusion.Method = "SnapP0"

arrEchellesFusionNb = 0
arrEchellesFusion(0) = XEchellesFusion
arrEchellesFusion(0).DeviseOrigine = "999"
arrEchellesFusion(0).CompteOrigine = "999999999999"
arrEchellesFusionSuite = True
Do Until Not arrEchellesFusionSuite
    srvEchellesFusion_Monitor XEchellesFusion
    XEchellesFusion = arrEchellesFusion(arrEchellesFusionNb)
    XEchellesFusion.Method = "SnapP0+"
Loop
End Sub

'-----------------------------------------------------
Function srvEchellesFusion_Update(recEchellesFusion As typeEchellesFusion)
'-----------------------------------------------------

srvEchellesFusion_Update = "?"

MsgTxtLen = 0
Call srvEchellesFusion_PutBuffer(recEchellesFusion)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    recEchellesFusion.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
    If Trim(recEchellesFusion.Err) <> "" Then
        Call srvEchellesFusion_Error(recEchellesFusion)
        srvEchellesFusion_Update = recEchellesFusion.Err
        Exit Function
    Else
        srvEchellesFusion_Update = Null
    End If
Else
    recEchellesFusion.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Public Function srvEchellesFusion_Monitor(recEchellesFusion As typeEchellesFusion)
'-----------------------------------------------------

arrEchellesFusionSuite = False
Select Case Mid$(Trim(recEchellesFusion.Method), 1, 4)
    Case "Seek"
                srvEchellesFusion_Monitor = srvEchellesFusion_Seek(recEchellesFusion)
    Case "Snap"
              srvEchellesFusion_Monitor = srvEchellesFusion_Snap(recEchellesFusion)
    Case Else
                recEchellesFusion.Err = recEchellesFusion.Method
                Call srvEchellesFusion_Error(recEchellesFusion)
                srvEchellesFusion_Monitor = recEchellesFusion.Err
End Select

End Function

'-----------------------------------------------------
Sub srvEchellesFusion_Error(recEchellesFusion As typeEchellesFusion)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Devise Cours : " ' & Chr$(10) & Chr$(13)

Select Case Mid$(recEchellesFusion.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recEchellesFusion.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvEchellesFusion  ( " _
                & Trim(recEchellesFusion.obj) & " : " & Trim(recEchellesFusion.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvEchellesFusion_GetBuffer(recEchellesFusion As typeEchellesFusion)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvEchellesFusion_GetBuffer = Null
recEchellesFusion.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recEchellesFusion.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recEchellesFusion.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recEchellesFusion.Err = Space$(10) Then
    recEchellesFusion.DeviseOrigine = Mid$(MsgTxt, K + 1, 3)
    recEchellesFusion.CompteOrigine = Mid$(MsgTxt, K + 4, 11)
    recEchellesFusion.DeviseFusion = Mid$(MsgTxt, K + 15, 3)
    recEchellesFusion.CompteFusion = Mid$(MsgTxt, K + 18, 11)
    recEchellesFusion.AmjDébut = Mid$(MsgTxt, K + 29, 8)
    recEchellesFusion.AmjFin = Mid$(MsgTxt, K + 37, 8)

Else
    srvEchellesFusion_GetBuffer = recEchellesFusion.Err
End If

MsgTxtIndex = MsgTxtIndex + recEchellesFusionLen

End Function

'---------------------------------------------------------
Private Sub srvEchellesFusion_PutBuffer(recEchellesFusion As typeEchellesFusion)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recEchellesFusion.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recEchellesFusion.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 3) = recEchellesFusion.DeviseOrigine
Mid$(MsgTxt, K + 4, 11) = recEchellesFusion.CompteOrigine
Mid$(MsgTxt, K + 15, 3) = recEchellesFusion.DeviseFusion
Mid$(MsgTxt, K + 18, 11) = recEchellesFusion.CompteFusion
Mid$(MsgTxt, K + 29, 8) = Format$(recEchellesFusion.AmjDébut, "00000000")
Mid$(MsgTxt, K + 37, 8) = Format$(recEchellesFusion.AmjFin, "00000000")
MsgTxtLen = MsgTxtLen + recEchellesFusionLen
End Sub



'---------------------------------------------------------
Private Function srvEchellesFusion_Seek(recEchellesFusion As typeEchellesFusion)
'---------------------------------------------------------

srvEchellesFusion_Seek = "?"


MsgTxtLen = 0
Call srvEchellesFusion_PutBuffer(recEchellesFusion)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvEchellesFusion_GetBuffer(recEchellesFusion)) Then
        srvEchellesFusion_Seek = Null
'    Else
'        Call srvEchellesFusion_Error(recEchellesFusion)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvEchellesFusion_Snap(recEchellesFusion As typeEchellesFusion)
'---------------------------------------------------------
Dim I As Integer
srvEchellesFusion_Snap = "?"
MsgTxtLen = 0
Call srvEchellesFusion_PutBuffer(recEchellesFusion)
Call srvEchellesFusion_PutBuffer(arrEchellesFusion(0))
If IsNull(SndRcv()) Then
    srvEchellesFusion_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvEchellesFusion_GetBuffer(recEchellesFusion)) Then
            Call arrEchellesFusion_AddItem(recEchellesFusion)
            arrEchellesFusionSuite = True
        Else
            arrEchellesFusionSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recEchellesFusion_Init(recEchellesFusion As typeEchellesFusion)
'---------------------------------------------------------
 MsgTxt = Space$(recEchellesFusionLen)
 MsgTxtIndex = 0
 Call srvEchellesFusion_GetBuffer(recEchellesFusion)
 recEchellesFusion.obj = "SRVECHCV"
 recEchellesFusion.AmjDébut = "00000000"
 recEchellesFusion.AmjFin = "00000000"
End Sub

'---------------------------------------------------------
Public Sub arrEchellesFusion_AddItem(recEchellesFusion As typeEchellesFusion)
'---------------------------------------------------------
          
arrEchellesFusionNb = arrEchellesFusionNb + 1
    
If arrEchellesFusionNb > arrEchellesFusionNbMax Then
    arrEchellesFusionNbMax = arrEchellesFusionNbMax + 10
    ReDim Preserve arrEchellesFusion(arrEchellesFusionNbMax)
End If
            
arrEchellesFusion(arrEchellesFusionNb) = recEchellesFusion
End Sub




Public Function arrEchellesFusion_ScanDeviseOrigineCompteOrigine(recEchellesFusion As typeEchellesFusion) As Integer
arrEchellesFusion_ScanDeviseOrigineCompteOrigine = -1
For arrEchellesFusionIndex = 1 To arrEchellesFusionNb
    If arrEchellesFusion(arrEchellesFusionIndex).Method <> constDelete _
    And arrEchellesFusion(arrEchellesFusionIndex).Method <> constIgnore Then
        If arrEchellesFusion(arrEchellesFusionIndex).DeviseOrigine = recEchellesFusion.DeviseOrigine _
        And arrEchellesFusion(arrEchellesFusionIndex).CompteOrigine = recEchellesFusion.CompteOrigine Then
            arrEchellesFusion_ScanDeviseOrigineCompteOrigine = arrEchellesFusionIndex
            Exit For
        End If
    End If
Next arrEchellesFusionIndex

End Function

