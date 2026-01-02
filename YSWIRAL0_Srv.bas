Attribute VB_Name = "srvYSWIRAL0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYSWIRAL0Len = 554 ' 34 +520
Public Const recYSWIRAL0_Block = 50
Public Const memoYSWIRAL0Len = 520

Type typeYSWIRAL0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    SWIRALDON       As String * 512                   ' DONNE MESSAGE
    SWIRALETA       As Integer                        '
    SWIRALMES       As String * 3                     '
End Type
    
    
Public arrYSWIRAL0() As typeYSWIRAL0
Public arrYSWIRAL0_NB As Integer
Public arrYSWIRAL0_NBMax As Integer
Public arrYSWIRAL0_Index As Integer
Public arrYSWIRAL0_Suite As Boolean

'-----------------------------------------------------
Function srvYSWIRAL0_Update(recYSWIRAL0 As typeYSWIRAL0)
'-----------------------------------------------------

srvYSWIRAL0_Update = "?"

MsgTxtLen = 0
Call srvYSWIRAL0_PutBuffer(recYSWIRAL0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYSWIRAL0_GetBuffer(recYSWIRAL0)) Then
        Call srvYSWIRAL0_Error(recYSWIRAL0)
        srvYSWIRAL0_Update = recYSWIRAL0.Err
        Exit Function
    Else
        srvYSWIRAL0_Update = Null
    End If
Else
    recYSWIRAL0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYSWIRAL0_Error(recYSWIRAL0 As typeYSWIRAL0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YSWIRAL0" & Chr$(10) & Chr$(13)

Select Case mId$(recYSWIRAL0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYSWIRAL0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YSWIRAL0s.bas  ( " & Trim(recYSWIRAL0.obj) & " : " & Trim(recYSWIRAL0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYSWIRAL0_Monitor(recYSWIRAL0 As typeYSWIRAL0)
'-----------------------------------------------------

arrYSWIRAL0_Suite = False
Select Case mId$(Trim(recYSWIRAL0.Method), 1, 4)
    Case "Snap"
              srvYSWIRAL0_Monitor = srvYSWIRAL0_Snap(recYSWIRAL0)
    Case Else
            srvYSWIRAL0_Monitor = srvYSWIRAL0_Seek(recYSWIRAL0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYSWIRAL0_GetBuffer(recYSWIRAL0 As typeYSWIRAL0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYSWIRAL0_GetBuffer = Null
recYSWIRAL0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYSWIRAL0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYSWIRAL0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYSWIRAL0.Err = Space$(10) Then
    recYSWIRAL0.SWIRALDON = mId$(MsgTxt, K + 1, 512)
    recYSWIRAL0.SWIRALETA = CInt(Val(mId$(MsgTxt, K + 513, 5)))
    recYSWIRAL0.SWIRALMES = mId$(MsgTxt, K + 518, 3)

Else
    srvYSWIRAL0_GetBuffer = recYSWIRAL0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYSWIRAL0Len

End Function

'---------------------------------------------------------
Public Sub srvYSWIRAL0_PutBuffer(recYSWIRAL0 As typeYSWIRAL0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYSWIRAL0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYSWIRAL0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 512) = recYSWIRAL0.SWIRALDON
    Mid$(MsgTxt, K + 513, 5) = Format$(recYSWIRAL0.SWIRALETA, "0000 ")
    Mid$(MsgTxt, K + 518, 3) = recYSWIRAL0.SWIRALMES
    

MsgTxtLen = MsgTxtLen + recYSWIRAL0Len
End Sub



'---------------------------------------------------------
Private Function srvYSWIRAL0_Seek(recYSWIRAL0 As typeYSWIRAL0)
'---------------------------------------------------------

srvYSWIRAL0_Seek = "?"
MsgTxtLen = 0
Call srvYSWIRAL0_PutBuffer(recYSWIRAL0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYSWIRAL0_GetBuffer(recYSWIRAL0)) Then
        srvYSWIRAL0_Seek = Null
    Else
        Call srvYSWIRAL0_Error(recYSWIRAL0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYSWIRAL0_Snap(recYSWIRAL0 As typeYSWIRAL0)
'---------------------------------------------------------
srvYSWIRAL0_Snap = "?"
MsgTxtLen = 0
Call srvYSWIRAL0_PutBuffer(recYSWIRAL0)
Call srvYSWIRAL0_PutBuffer(arrYSWIRAL0(0))
If IsNull(SndRcv()) Then
    srvYSWIRAL0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYSWIRAL0_GetBuffer(recYSWIRAL0)) Then
            Call arrYSWIRAL0_AddItem(recYSWIRAL0)
            arrYSWIRAL0_Suite = True
        Else
            arrYSWIRAL0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYSWIRAL0_AddItem(recYSWIRAL0 As typeYSWIRAL0)
'---------------------------------------------------------
          
arrYSWIRAL0_NB = arrYSWIRAL0_NB + 1
    
If arrYSWIRAL0_NB > arrYSWIRAL0_NBMax Then
    arrYSWIRAL0_NBMax = arrYSWIRAL0_NBMax + recYSWIRAL0_Block
    ReDim Preserve arrYSWIRAL0(arrYSWIRAL0_NBMax)
End If
            
arrYSWIRAL0(arrYSWIRAL0_NB) = recYSWIRAL0
End Sub



'---------------------------------------------------------
Public Sub recYSWIRAL0_Init(recYSWIRAL0 As typeYSWIRAL0)
'---------------------------------------------------------
recYSWIRAL0.obj = "ZSWIRAL0_S"
recYSWIRAL0.Method = ""
recYSWIRAL0.Err = ""
recYSWIRAL0.SWIRALDON = ""
recYSWIRAL0.SWIRALETA = 1
recYSWIRAL0.SWIRALMES = ""

End Sub




