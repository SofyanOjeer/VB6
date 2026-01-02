Attribute VB_Name = "srvYSWIALL0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYSWIALL0Len = 546 ' 34 +544
Public Const recYSWIALL0_Block = 20
Public Const memoYSWIALL0Len = 512
Public Const constYSWIALL0 = "YSWIALL0  "

Type typeYSWIALL0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    SWIALLDON       As String * 512                   ' DONNE MESSAGE
End Type
    
    
Public arrYSWIALL0() As typeYSWIALL0
Public arrYSWIALL0_NB As Integer
Public arrYSWIALL0_NBMax As Integer
Public arrYSWIALL0_Index As Integer
Public arrYSWIALL0_Suite As Boolean

'-----------------------------------------------------
Function srvYSWIALL0_Update(recYSWIALL0 As typeYSWIALL0)
'-----------------------------------------------------

srvYSWIALL0_Update = "?"

MsgTxtLen = 0
Call srvYSWIALL0_PutBuffer(recYSWIALL0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYSWIALL0_GetBuffer(recYSWIALL0)) Then
        Call srvYSWIALL0_Error(recYSWIALL0)
        srvYSWIALL0_Update = recYSWIALL0.Err
        Exit Function
    Else
        srvYSWIALL0_Update = Null
    End If
Else
    recYSWIALL0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYSWIALL0_Error(recYSWIALL0 As typeYSWIALL0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YSWIALL0" & Chr$(10) & Chr$(13)

Select Case mId$(recYSWIALL0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYSWIALL0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YSWIALL0s.bas  ( " & Trim(recYSWIALL0.obj) & " : " & Trim(recYSWIALL0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYSWIALL0_Monitor(recYSWIALL0 As typeYSWIALL0)
'-----------------------------------------------------

arrYSWIALL0_Suite = False
Select Case mId$(Trim(recYSWIALL0.Method), 1, 4)
    Case "Snap"
              srvYSWIALL0_Monitor = srvYSWIALL0_Snap(recYSWIALL0)
    Case Else
            srvYSWIALL0_Monitor = srvYSWIALL0_Seek(recYSWIALL0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYSWIALL0_GetBuffer(recYSWIALL0 As typeYSWIALL0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYSWIALL0_GetBuffer = Null
recYSWIALL0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYSWIALL0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYSWIALL0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYSWIALL0.Err = Space$(10) Then
    recYSWIALL0.SWIALLDON = mId$(MsgTxt, K + 1, 512)
Else
    srvYSWIALL0_GetBuffer = recYSWIALL0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYSWIALL0Len

End Function

'---------------------------------------------------------
Public Sub srvYSWIALL0_PutBuffer(recYSWIALL0 As typeYSWIALL0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYSWIALL0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYSWIALL0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 512) = recYSWIALL0.SWIALLDON
MsgTxtLen = MsgTxtLen + recYSWIALL0Len
End Sub



'---------------------------------------------------------
Private Function srvYSWIALL0_Seek(recYSWIALL0 As typeYSWIALL0)
'---------------------------------------------------------

srvYSWIALL0_Seek = "?"
MsgTxtLen = 0
Call srvYSWIALL0_PutBuffer(recYSWIALL0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYSWIALL0_GetBuffer(recYSWIALL0)) Then
        srvYSWIALL0_Seek = Null
    Else
        ''Call srvYSWIALL0_Error(recYSWIALL0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYSWIALL0_Snap(recYSWIALL0 As typeYSWIALL0)
'---------------------------------------------------------
srvYSWIALL0_Snap = "?"
MsgTxtLen = 0
Call srvYSWIALL0_PutBuffer(recYSWIALL0)
Call srvYSWIALL0_PutBuffer(arrYSWIALL0(0))
If IsNull(SndRcv()) Then
    srvYSWIALL0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYSWIALL0_GetBuffer(recYSWIALL0)) Then
            Call arrYSWIALL0_AddItem(recYSWIALL0)
            arrYSWIALL0_Suite = True
        Else
            arrYSWIALL0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYSWIALL0_AddItem(recYSWIALL0 As typeYSWIALL0)
'---------------------------------------------------------
          
arrYSWIALL0_NB = arrYSWIALL0_NB + 1
    
If arrYSWIALL0_NB > arrYSWIALL0_NBMax Then
    arrYSWIALL0_NBMax = arrYSWIALL0_NBMax + recYSWIALL0_Block
    ReDim Preserve arrYSWIALL0(arrYSWIALL0_NBMax)
End If
            
arrYSWIALL0(arrYSWIALL0_NB) = recYSWIALL0
End Sub



'---------------------------------------------------------
Public Sub recYSWIALL0_Init(recYSWIALL0 As typeYSWIALL0)
'---------------------------------------------------------
recYSWIALL0.obj = "YSWIALL0_S"
recYSWIALL0.Method = ""
recYSWIALL0.Err = ""
recYSWIALL0.SWIALLDON = ""
End Sub











