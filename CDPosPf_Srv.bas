Attribute VB_Name = "srvCDPosPf"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recCDPosPfLen = 172 ' 34 + 138
Public Const MemoCDPosPfLen = 138
Public Const recCDPosPf_Block = 100

Type typeCDPosPf
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    POPKEY                  As Long
    POEKEY                  As Long
    POEPFX                  As String * 3
    POENUM                  As Long
    PODKEY                  As Long
    PODPFX                  As String * 3
    PODNUM                  As Long
    POBRC                   As String * 4
    PODVAL                  As String * 8
    POATIB                  As String * 4
    POATIN                  As String * 6
    POATIS                  As String * 3
    POCPT                   As String * 25
    POTRCD                  As String * 3
    PODBCR                  As String * 1
    POAMT                   As Currency
    POCCY                   As String * 3
    POACTY                  As String * 2
    POSPCD                  As String * 6
    POSKCD                  As String * 2
    
End Type
    
Public arrCDPosPf() As typeCDPosPf
Public arrCDPosPf_NB As Integer
Public arrCDPosPf_NBMax As Integer
Public arrCDPosPf_Index As Integer
Public arrCDPosPf_Suite As Boolean

Public xCDPosPf As typeCDPosPf

'-----------------------------------------------------
Function srvCDPosPf_Update(recCDPosPf As typeCDPosPf)
'-----------------------------------------------------

srvCDPosPf_Update = "?"

MsgTxtLen = 0
Call srvCDPosPf_PutBuffer(recCDPosPf)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCDPosPf_GetBuffer(recCDPosPf)) Then
        Call srvCDPosPf_Error(recCDPosPf)
        srvCDPosPf_Update = recCDPosPf.Err
        Exit Function
    Else
        srvCDPosPf_Update = Null
    End If
Else
    recCDPosPf.Err = "srv"
End If


'=====================================================
End Function



Public Sub srvCDPosPf_Load(recCDPosPfMin As typeCDPosPf, recCDPosPfMax As typeCDPosPf)
Dim mMethod As String

mMethod = Trim(recCDPosPfMin.Method) & "+"
arrCDPosPf_NBMax = 0
arrCDPosPf_Suite = True: arrCDPosPf_NB = 0
arrCDPosPf_NBMax = recCDPosPf_Block: ReDim arrCDPosPf(arrCDPosPf_NBMax)

arrCDPosPf(0) = recCDPosPfMax
arrCDPosPf_Suite = True
Do Until Not arrCDPosPf_Suite
    srvCDPosPf_Monitor recCDPosPfMin
    recCDPosPfMin = arrCDPosPf(arrCDPosPf_NB)
    recCDPosPfMin.Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Function srvCDPosPf_Dtaq_Put(lFct As String, recCDPosPf As typeCDPosPf)
'-----------------------------------------------------

srvCDPosPf_Dtaq_Put = Null
Select Case lFct
    Case "Init": MsgTxtLen = 0
    Case "Add": Call srvCDPosPf_PutBuffer(recCDPosPf)
                If MsgTxtLen + recCDPosPfLen >= recCDPosPf_Block * recCDPosPfLen Then
                    Call srvCDPosPf_Dtaq_Snd(recCDPosPf): MsgTxtLen = 0
                End If
    Case "Snd": If MsgTxtLen > 0 Then Call srvCDPosPf_Dtaq_Snd(recCDPosPf)
    Case Else: srvCDPosPf_Dtaq_Put = lFct
End Select
'=====================================================
End Function


'-----------------------------------------------------
Function srvCDPosPf_Dtaq_Snd(recCDPosPf As typeCDPosPf)
'-----------------------------------------------------

srvCDPosPf_Dtaq_Snd = "?"

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCDPosPf_GetBuffer(recCDPosPf)) Then
        Call srvCDPosPf_Error(recCDPosPf)
        srvCDPosPf_Dtaq_Snd = recCDPosPf.Err
        Exit Function
    Else
        srvCDPosPf_Dtaq_Snd = Null
    End If
Else
    recCDPosPf.Err = "Snd"
End If


'=====================================================
End Function



'-----------------------------------------------------
Public Function srvCDPosPf_Monitor(recCDPosPf As typeCDPosPf)
'-----------------------------------------------------
blnFR_Convert = False

arrCDPosPf_Suite = False
Select Case mId$(Trim(recCDPosPf.Method), 1, 4)
    Case "Seek"
                srvCDPosPf_Monitor = srvCDPosPf_Seek(recCDPosPf)
    Case "Snap"
              srvCDPosPf_Monitor = srvCDPosPf_Snap(recCDPosPf)
    Case Else
                recCDPosPf.Err = recCDPosPf.Method
                Call srvCDPosPf_Error(recCDPosPf)
                srvCDPosPf_Monitor = recCDPosPf.Err
End Select

End Function

'-----------------------------------------------------
Sub srvCDPosPf_Error(recCDPosPf As typeCDPosPf)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "CDPosPf" & Chr$(10) & Chr$(13)

Select Case mId$(recCDPosPf.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recCDPosPf.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recCDPosPf.POPKEY & " : " & recCDPosPf.POEKEY _
        , I, "module : CDPosPfs.bas  ( " & Trim(recCDPosPf.obj) & " : " & Trim(recCDPosPf.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvCDPosPf_GetBuffer(recCDPosPf As typeCDPosPf)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvCDPosPf_GetBuffer = Null
recCDPosPf.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recCDPosPf.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recCDPosPf.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recCDPosPf.Err = Space$(10) Then
    recCDPosPf.POPKEY = CLng(Val(mId$(MsgTxt, K + 1, 12)))
    recCDPosPf.POEKEY = CLng(Val(mId$(MsgTxt, K + 13, 12)))
    recCDPosPf.POEPFX = mId$(MsgTxt, K + 25, 3)
    recCDPosPf.POENUM = CLng(Val(mId$(MsgTxt, K + 28, 6)))
    recCDPosPf.PODKEY = CLng(Val(mId$(MsgTxt, K + 34, 12)))
    recCDPosPf.PODPFX = mId$(MsgTxt, K + 46, 3)
    recCDPosPf.PODNUM = CLng(Val(mId$(MsgTxt, K + 49, 6)))
    recCDPosPf.POBRC = mId$(MsgTxt, K + 55, 4)
    recCDPosPf.PODVAL = mId$(MsgTxt, K + 59, 8)
    recCDPosPf.POATIB = mId$(MsgTxt, K + 67, 4)
    recCDPosPf.POATIN = mId$(MsgTxt, K + 71, 6)
    recCDPosPf.POATIS = mId$(MsgTxt, K + 77, 3)
    recCDPosPf.POCPT = mId$(MsgTxt, K + 80, 25)
    recCDPosPf.POTRCD = mId$(MsgTxt, K + 105, 3)
    recCDPosPf.PODBCR = mId$(MsgTxt, K + 108, 1)
    recCDPosPf.POAMT = CCur(Val(mId$(MsgTxt, K + 106, 17)) / 100)
    recCDPosPf.POCCY = mId$(MsgTxt, K + 126, 3)
    recCDPosPf.POACTY = mId$(MsgTxt, K + 129, 2)
    recCDPosPf.POSPCD = mId$(MsgTxt, K + 131, 6)
    recCDPosPf.POSKCD = mId$(MsgTxt, K + 137, 2)

Else
    srvCDPosPf_GetBuffer = recCDPosPf.Err
End If

MsgTxtIndex = MsgTxtIndex + recCDPosPfLen

End Function

'---------------------------------------------------------
Public Sub srvCDPosPf_PutBuffer(recCDPosPf As typeCDPosPf)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recCDPosPf.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recCDPosPf.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 12) = Format$(recCDPosPf.POPKEY, "000000000000")
Mid$(MsgTxt, K + 13, 12) = Format$(recCDPosPf.POEKEY, "000000000000")
Mid$(MsgTxt, K + 25, 3) = recCDPosPf.POEPFX
Mid$(MsgTxt, K + 28, 6) = Format$(recCDPosPf.POENUM, "000000")
Mid$(MsgTxt, K + 34, 12) = Format$(recCDPosPf.PODKEY, "000000000000")
Mid$(MsgTxt, K + 46, 3) = recCDPosPf.PODPFX
Mid$(MsgTxt, K + 49, 6) = Format$(recCDPosPf.PODNUM, "000000")
Mid$(MsgTxt, K + 55, 4) = recCDPosPf.POBRC
Mid$(MsgTxt, K + 59, 8) = Format$(recCDPosPf.PODVAL, "00000000")
Mid$(MsgTxt, K + 67, 4) = recCDPosPf.POATIB
Mid$(MsgTxt, K + 71, 6) = recCDPosPf.POATIN
Mid$(MsgTxt, K + 77, 3) = recCDPosPf.POATIS
Mid$(MsgTxt, K + 80, 25) = recCDPosPf.POCPT
Mid$(MsgTxt, K + 105, 3) = recCDPosPf.POTRCD
Mid$(MsgTxt, K + 108, 1) = recCDPosPf.PODBCR
Mid$(MsgTxt, K + 109, 17) = Format$(recCDPosPf.POAMT * 100, "00000000000000000")
Mid$(MsgTxt, K + 126, 3) = recCDPosPf.POCCY
Mid$(MsgTxt, K + 129, 2) = recCDPosPf.POACTY
Mid$(MsgTxt, K + 131, 6) = recCDPosPf.POSPCD
Mid$(MsgTxt, K + 137, 2) = recCDPosPf.POSKCD

MsgTxtLen = MsgTxtLen + recCDPosPfLen

  
End Sub



'---------------------------------------------------------
Private Function srvCDPosPf_Seek(recCDPosPf As typeCDPosPf)
'---------------------------------------------------------

srvCDPosPf_Seek = "?"
MsgTxtLen = 0
Call srvCDPosPf_PutBuffer(recCDPosPf)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvCDPosPf_GetBuffer(recCDPosPf)) Then
        srvCDPosPf_Seek = Null
    Else
        Call srvCDPosPf_Error(recCDPosPf)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvCDPosPf_Snap(recCDPosPf As typeCDPosPf)
'---------------------------------------------------------
srvCDPosPf_Snap = "?"
MsgTxtLen = 0
Call srvCDPosPf_PutBuffer(recCDPosPf)
Call srvCDPosPf_PutBuffer(arrCDPosPf(0))
If IsNull(SndRcv()) Then
    srvCDPosPf_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvCDPosPf_GetBuffer(recCDPosPf)) Then
            Call arrCDPosPf_AddItem(recCDPosPf)
            arrCDPosPf_Suite = True
        Else
            arrCDPosPf_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recCDPosPf_Init(recCDPosPf As typeCDPosPf)
'---------------------------------------------------------
MsgTxt = Space$(recCDPosPfLen)
MsgTxtIndex = 0
Call srvCDPosPf_GetBuffer(recCDPosPf)
recCDPosPf.obj = "SRVCDPOSPF"

End Sub

'---------------------------------------------------------
Public Sub arrCDPosPf_AddItem(recCDPosPf As typeCDPosPf)
'---------------------------------------------------------
          
arrCDPosPf_NB = arrCDPosPf_NB + 1
    
If arrCDPosPf_NB > arrCDPosPf_NBMax Then
    arrCDPosPf_NBMax = arrCDPosPf_NBMax + recCDPosPf_Block
    ReDim Preserve arrCDPosPf(arrCDPosPf_NBMax)
End If
            
arrCDPosPf(arrCDPosPf_NB) = recCDPosPf
End Sub


