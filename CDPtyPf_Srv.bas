Attribute VB_Name = "srvCDPtyPf"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recCDPtyPfLen = 101 ' 34 + 67
Public Const MemoCDPtyPfLen = 67
Public Const recCDPtyPf_Block = 100

Type typeCDPtyPf
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    PTKEY                  As Long
    PTNOM                  As String * 35
    PTMNM                  As String * 20
    
End Type
    
Public arrCDPtyPf() As typeCDPtyPf
Public arrCDPtyPf_NB As Integer
Public arrCDPtyPf_NBMax As Integer
Public arrCDPtyPf_Index As Integer
Public arrCDPtyPf_Suite As Boolean

Public xCDPtyPf As typeCDPtyPf

'-----------------------------------------------------
Function srvCDPtyPf_Update(recCDPtyPf As typeCDPtyPf)
'-----------------------------------------------------

srvCDPtyPf_Update = "?"

MsgTxtLen = 0
Call srvCDPtyPf_PutBuffer(recCDPtyPf)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCDPtyPf_GetBuffer(recCDPtyPf)) Then
        Call srvCDPtyPf_Error(recCDPtyPf)
        srvCDPtyPf_Update = recCDPtyPf.Err
        Exit Function
    Else
        srvCDPtyPf_Update = Null
    End If
Else
    recCDPtyPf.Err = "srv"
End If


'=====================================================
End Function



Public Sub srvCDPtyPf_Load(recCDPtyPfMin As typeCDPtyPf, recCDPtyPfMax As typeCDPtyPf)
Dim mMethod As String

mMethod = Trim(recCDPtyPfMin.Method) & "+"
arrCDPtyPf_NBMax = 0
arrCDPtyPf_Suite = True: arrCDPtyPf_NB = 0
arrCDPtyPf_NBMax = recCDPtyPf_Block: ReDim arrCDPtyPf(arrCDPtyPf_NBMax)

arrCDPtyPf(0) = recCDPtyPfMax
arrCDPtyPf_Suite = True
Do Until Not arrCDPtyPf_Suite
    srvCDPtyPf_Monitor recCDPtyPfMin
    recCDPtyPfMin = arrCDPtyPf(arrCDPtyPf_NB)
    recCDPtyPfMin.Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Function srvCDPtyPf_Dtaq_Put(lFct As String, recCDPtyPf As typeCDPtyPf)
'-----------------------------------------------------

srvCDPtyPf_Dtaq_Put = Null
Select Case lFct
    Case "Init": MsgTxtLen = 0
    Case "Add": Call srvCDPtyPf_PutBuffer(recCDPtyPf)
                If MsgTxtLen + recCDPtyPfLen >= recCDPtyPf_Block * recCDPtyPfLen Then
                    Call srvCDPtyPf_Dtaq_Snd(recCDPtyPf): MsgTxtLen = 0
                End If
    Case "Snd": If MsgTxtLen > 0 Then Call srvCDPtyPf_Dtaq_Snd(recCDPtyPf)
    Case Else: srvCDPtyPf_Dtaq_Put = lFct
End Select
'=====================================================
End Function


'-----------------------------------------------------
Function srvCDPtyPf_Dtaq_Snd(recCDPtyPf As typeCDPtyPf)
'-----------------------------------------------------

srvCDPtyPf_Dtaq_Snd = "?"

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCDPtyPf_GetBuffer(recCDPtyPf)) Then
        Call srvCDPtyPf_Error(recCDPtyPf)
        srvCDPtyPf_Dtaq_Snd = recCDPtyPf.Err
        Exit Function
    Else
        srvCDPtyPf_Dtaq_Snd = Null
    End If
Else
    recCDPtyPf.Err = "Snd"
End If


'=====================================================
End Function



'-----------------------------------------------------
Public Function srvCDPtyPf_Monitor(recCDPtyPf As typeCDPtyPf)
'-----------------------------------------------------
blnFR_Convert = False

arrCDPtyPf_Suite = False
Select Case mId$(Trim(recCDPtyPf.Method), 1, 4)
    Case "Seek"
                srvCDPtyPf_Monitor = srvCDPtyPf_Seek(recCDPtyPf)
    Case "Snap"
              srvCDPtyPf_Monitor = srvCDPtyPf_Snap(recCDPtyPf)
    Case Else
                recCDPtyPf.Err = recCDPtyPf.Method
                Call srvCDPtyPf_Error(recCDPtyPf)
                srvCDPtyPf_Monitor = recCDPtyPf.Err
End Select

End Function

'-----------------------------------------------------
Sub srvCDPtyPf_Error(recCDPtyPf As typeCDPtyPf)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "CDPtyPf" & Chr$(10) & Chr$(13)

Select Case mId$(recCDPtyPf.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recCDPtyPf.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recCDPtyPf.PTKEY & " : " & recCDPtyPf.PTMNM _
        , I, "module : CDPtyPfs.bas  ( " & Trim(recCDPtyPf.obj) & " : " & Trim(recCDPtyPf.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvCDPtyPf_GetBuffer(recCDPtyPf As typeCDPtyPf)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvCDPtyPf_GetBuffer = Null
recCDPtyPf.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recCDPtyPf.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recCDPtyPf.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recCDPtyPf.Err = Space$(10) Then
    recCDPtyPf.PTKEY = CLng(Val(mId$(MsgTxt, K + 1, 12)))
    recCDPtyPf.PTNOM = mId$(MsgTxt, K + 13, 35)
    recCDPtyPf.PTMNM = mId$(MsgTxt, K + 48, 20)

Else
    srvCDPtyPf_GetBuffer = recCDPtyPf.Err
End If

MsgTxtIndex = MsgTxtIndex + recCDPtyPfLen

End Function

'---------------------------------------------------------
Public Sub srvCDPtyPf_PutBuffer(recCDPtyPf As typeCDPtyPf)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recCDPtyPf.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recCDPtyPf.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 12) = Format$(recCDPtyPf.PTKEY, "000000000000")
Mid$(MsgTxt, K + 13, 35) = recCDPtyPf.PTNOM
Mid$(MsgTxt, K + 48, 20) = recCDPtyPf.PTMNM

MsgTxtLen = MsgTxtLen + recCDPtyPfLen

  
End Sub



'---------------------------------------------------------
Private Function srvCDPtyPf_Seek(recCDPtyPf As typeCDPtyPf)
'---------------------------------------------------------

srvCDPtyPf_Seek = "?"
MsgTxtLen = 0
Call srvCDPtyPf_PutBuffer(recCDPtyPf)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvCDPtyPf_GetBuffer(recCDPtyPf)) Then
        srvCDPtyPf_Seek = Null
    Else
        Call srvCDPtyPf_Error(recCDPtyPf)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvCDPtyPf_Snap(recCDPtyPf As typeCDPtyPf)
'---------------------------------------------------------
srvCDPtyPf_Snap = "?"
MsgTxtLen = 0
Call srvCDPtyPf_PutBuffer(recCDPtyPf)
Call srvCDPtyPf_PutBuffer(arrCDPtyPf(0))
If IsNull(SndRcv()) Then
    srvCDPtyPf_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvCDPtyPf_GetBuffer(recCDPtyPf)) Then
            Call arrCDPtyPf_AddItem(recCDPtyPf)
            arrCDPtyPf_Suite = True
        Else
            arrCDPtyPf_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recCDPtyPf_Init(recCDPtyPf As typeCDPtyPf)
'---------------------------------------------------------
MsgTxt = Space$(recCDPtyPfLen)
MsgTxtIndex = 0
Call srvCDPtyPf_GetBuffer(recCDPtyPf)
recCDPtyPf.obj = "SRVCDPTY"

End Sub

'---------------------------------------------------------
Public Sub arrCDPtyPf_AddItem(recCDPtyPf As typeCDPtyPf)
'---------------------------------------------------------
          
arrCDPtyPf_NB = arrCDPtyPf_NB + 1
    
If arrCDPtyPf_NB > arrCDPtyPf_NBMax Then
    arrCDPtyPf_NBMax = arrCDPtyPf_NBMax + recCDPtyPf_Block
    ReDim Preserve arrCDPtyPf(arrCDPtyPf_NBMax)
End If
            
arrCDPtyPf(arrCDPtyPf_NB) = recCDPtyPf
End Sub


