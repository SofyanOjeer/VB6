Attribute VB_Name = "srvOppChq"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recOppChqLen = 52 '34 + 18

Type typeOppChq
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Société                As String * 3
    Agence                 As String * 3
    Racine                 As String * 5
    Numéro                 As String * 7

End Type
    
Public arrOppChq() As typeOppChq
Public arrOppChqNb As Integer
Public arrOppChqNbMax As Integer
Public arrOppChqIndex As Integer
Public arrOppChqSuite As Boolean

Public recOppChq As typeOppChq


'-----------------------------------------------------
Public Function srvOppChq_Monitor(recOppChq As typeOppChq)
'-----------------------------------------------------

arrOppChqSuite = False
Select Case mId$(Trim(recOppChq.Method), 1, 4)
    Case "Seek"
                srvOppChq_Monitor = srvOppChq_Seek(recOppChq)
    Case "Snap", "Prev"
              srvOppChq_Monitor = srvOppChq_Snap(recOppChq)
    Case Else
                recOppChq.Err = recOppChq.Method
                Call srvOppChq_Error(recOppChq)
                srvOppChq_Monitor = recOppChq.Err
End Select

End Function

'-----------------------------------------------------
Sub srvOppChq_Error(recOppChq As typeOppChq)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Opposition sur chèque): " ' & Chr$(10) & Chr$(13)

Select Case mId$(recOppChq.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recOppChq.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvOppChq_.bas  ( " _
                & Trim(recOppChq.obj) & " : " & Trim(recOppChq.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvOppChq_GetBuffer(recOppChq As typeOppChq)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvOppChq_GetBuffer = Null
recOppChq.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recOppChq.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recOppChq.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recOppChq.Err = Space$(10) Then
    recOppChq.Société = mId$(MsgTxt, K + 1, 3)
    recOppChq.Agence = mId$(MsgTxt, K + 4, 3)
    recOppChq.Racine = mId$(MsgTxt, K + 7, 5)
    recOppChq.Numéro = mId$(MsgTxt, K + 12, 7)
Else
    srvOppChq_GetBuffer = recOppChq.Err
End If

MsgTxtIndex = MsgTxtIndex + recOppChqLen

End Function

'---------------------------------------------------------
Public Sub srvOppChq_PutBuffer(recOppChq As typeOppChq)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recOppChq.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recOppChq.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 3) = recOppChq.Société
Mid$(MsgTxt, K + 4, 3) = recOppChq.Agence
Mid$(MsgTxt, K + 7, 5) = recOppChq.Racine
Mid$(MsgTxt, K + 12, 7) = recOppChq.Numéro


MsgTxtLen = MsgTxtLen + recOppChqLen
End Sub



'---------------------------------------------------------
Private Function srvOppChq_Seek(recOppChq As typeOppChq)
'---------------------------------------------------------

srvOppChq_Seek = "?"
MsgTxtLen = 0
Call srvOppChq_PutBuffer(recOppChq)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvOppChq_GetBuffer(recOppChq)) Then
        srvOppChq_Seek = Null
    Else
        Call srvOppChq_Error(recOppChq)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvOppChq_Snap(recOppChq As typeOppChq)
'---------------------------------------------------------
Dim I As Integer
srvOppChq_Snap = "?"
MsgTxtLen = 0
Call srvOppChq_PutBuffer(recOppChq)
Call srvOppChq_PutBuffer(arrOppChq(0))
If IsNull(SndRcv()) Then
    srvOppChq_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvOppChq_GetBuffer(recOppChq)) Then
            Call arrOppChq_AddItem(recOppChq)
            arrOppChqSuite = True
        Else
            arrOppChqSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recOppChq_Init(recOppChq As typeOppChq)
'---------------------------------------------------------
recOppChq.Method = ""
recOppChq.obj = "SRVOPPCHQ"
recOppChq.Err = ""
recOppChq.Société = "000"
recOppChq.Agence = "000"
recOppChq.Racine = "00000"
recOppChq.Numéro = "0000000"
End Sub

'---------------------------------------------------------
Public Sub arrOppChq_AddItem(recOppChq As typeOppChq)
'---------------------------------------------------------
          
arrOppChqNb = arrOppChqNb + 1
    
If arrOppChqNb > arrOppChqNbMax Then
    arrOppChqNbMax = arrOppChqNbMax + 10
    ReDim Preserve arrOppChq(arrOppChqNbMax)
End If
            
arrOppChq(arrOppChqNb) = recOppChq
End Sub




