Attribute VB_Name = "srvCDCgbPf"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recCDCgbPfLen = 134 ' 100 + 34
Public Const recCDCgbPf_Block = 200

Type typeCDCgbPf
    obj         As String * 12
    Method      As String * 12
    Err         As String * 10
    
    CGCENR      As String * 1
    CGDPFX      As String * 3
    CGDNUM      As Long
    CGDCCY      As String * 3
    CGCODC      As String * 2
    CGCOTH      As Currency
    CGCOEN      As Currency
    CGCOTP      As Currency
    CGCOAP      As Currency
    CGCODF      As Currency
    
End Type

Public arrCDCgbPf() As typeCDCgbPf
Public arrCDCgbPf_Nb As Integer
Public arrCDCgbPf_NbMax As Integer
Public arrCDCgbPf_Index As Integer
Public arrCDCgbPf_Suite As Boolean

Public Sub srvCDCgbPf_ElpDisplay(recCDCgbPf As typeCDCgbPf)
frmElpDisplay.fgData.Rows = 18
frmElpDisplay.fgData.Row = 1
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "obj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDCgbPf.obj
frmElpDisplay.fgData.Row = 2
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Method"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDCgbPf.Method
frmElpDisplay.fgData.Row = 3
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Err"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDCgbPf.Err
    
frmElpDisplay.fgData.Row = 4
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CGCENR"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDCgbPf.CGCENR
frmElpDisplay.fgData.Row = 5
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CGDPFX"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDCgbPf.CGDPFX
frmElpDisplay.fgData.Row = 6
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CGDNUM"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDCgbPf.CGDNUM
frmElpDisplay.fgData.Row = 7
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CGDCCY"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDCgbPf.CGDCCY
frmElpDisplay.fgData.Row = 8
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CGCODC"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDCgbPf.CGCODC
frmElpDisplay.fgData.Row = 9
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CGCOTH"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDCgbPf.CGCOTH
frmElpDisplay.fgData.Row = 10
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CGCOEN"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDCgbPf.CGCOEN
frmElpDisplay.fgData.Row = 11
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CGCOTP"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDCgbPf.CGCOTP
frmElpDisplay.fgData.Row = 12
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CGCOAP"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDCgbPf.CGCOAP
frmElpDisplay.fgData.Row = 13
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CGCODF"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDCgbPf.CGCODF

frmElpDisplay.Show vbModal

End Sub

Public Sub srvCDCgbPf_Load(recCDCgbPfMin As typeCDCgbPf, recCDCgbPfMax As typeCDCgbPf)
Dim mMethod As String

mMethod = Trim(recCDCgbPfMin.Method) & "+"
arrCDCgbPf_NbMax = 0
arrCDCgbPf_Suite = True: arrCDCgbPf_Nb = 0
arrCDCgbPf_NbMax = recCDCgbPf_Block: ReDim arrCDCgbPf(arrCDCgbPf_NbMax)

arrCDCgbPf(0) = recCDCgbPfMax
arrCDCgbPf_Suite = True
Do Until Not arrCDCgbPf_Suite
    srvCDCgbPf_Monitor recCDCgbPfMin
    recCDCgbPfMin = arrCDCgbPf(arrCDCgbPf_Nb)
    recCDCgbPfMin.Method = mMethod
Loop

End Sub


'-----------------------------------------------------
Function srvCDCgbPf_Update(recCDCgbPf As typeCDCgbPf)
'-----------------------------------------------------

srvCDCgbPf_Update = "?"

MsgTxtLen = 0
Call srvCDCgbPf_PutBuffer(recCDCgbPf)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCDCgbPf_GetBuffer(recCDCgbPf)) Then
        Call srvCDCgbPf_Error(recCDCgbPf)
        srvCDCgbPf_Update = recCDCgbPf.Err
        Exit Function
    Else
        srvCDCgbPf_Update = Null
    End If
Else
    recCDCgbPf.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Public Function srvCDCgbPf_Monitor(recCDCgbPf As typeCDCgbPf)
'-----------------------------------------------------

blnFR_Convert = False

arrCDCgbPf_Suite = False
Select Case mId$(Trim(recCDCgbPf.Method), 1, 4)
    Case "Seek"
                srvCDCgbPf_Monitor = srvCDCgbPf_Seek(recCDCgbPf)
    Case "Snap"
              srvCDCgbPf_Monitor = srvCDCgbPf_Snap(recCDCgbPf)
    Case Else
                recCDCgbPf.Err = recCDCgbPf.Method
                Call srvCDCgbPf_Error(recCDCgbPf)
                srvCDCgbPf_Monitor = recCDCgbPf.Err
End Select
End Function

'-----------------------------------------------------
Sub srvCDCgbPf_Error(recCDCgbPf As typeCDCgbPf)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "CDCgbPf" & Chr$(10) & Chr$(13)

Select Case mId$(recCDCgbPf.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recCDCgbPf.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : CDCgbPf_Srv.bas  ( " _
                & Trim(recCDCgbPf.obj) & " : " & Trim(recCDCgbPf.Method) & " )"

End Sub

'---------------------------------------------------------
Public Function srvCDCgbPf_GetBuffer(recCDCgbPf As typeCDCgbPf)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvCDCgbPf_GetBuffer = Null
recCDCgbPf.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recCDCgbPf.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recCDCgbPf.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recCDCgbPf.Err = Space$(10) Then
  
    recCDCgbPf.CGCENR = mId$(MsgTxt, K + 1, 1)
    recCDCgbPf.CGDPFX = mId$(MsgTxt, K + 2, 3)
    recCDCgbPf.CGDNUM = CLng(Val(mId$(MsgTxt, K + 5, 6)))
    recCDCgbPf.CGDCCY = mId$(MsgTxt, K + 11, 3)
    recCDCgbPf.CGCODC = mId$(MsgTxt, K + 14, 2)
    recCDCgbPf.CGCOTH = CCur(Val(mId$(MsgTxt, K + 16, 17)) / 100)
    recCDCgbPf.CGCOEN = CCur(Val(mId$(MsgTxt, K + 33, 17)) / 100)
    recCDCgbPf.CGCOTP = CCur(Val(mId$(MsgTxt, K + 50, 17)) / 100)
    recCDCgbPf.CGCOAP = CCur(Val(mId$(MsgTxt, K + 67, 17)) / 100)
    recCDCgbPf.CGCODF = CCur(Val(mId$(MsgTxt, K + 84, 17)) / 100)
Else
    srvCDCgbPf_GetBuffer = recCDCgbPf.Err
End If

MsgTxtIndex = MsgTxtIndex + recCDCgbPfLen

End Function

'---------------------------------------------------------
Private Sub srvCDCgbPf_PutBuffer(recCDCgbPf As typeCDCgbPf)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recCDCgbPf.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recCDCgbPf.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 1) = recCDCgbPf.CGCENR
Mid$(MsgTxt, K + 2, 3) = recCDCgbPf.CGDPFX
Mid$(MsgTxt, K + 5, 6) = Format$(recCDCgbPf.CGDNUM, "000000")
Mid$(MsgTxt, K + 11, 3) = recCDCgbPf.CGDCCY
Mid$(MsgTxt, K + 14, 2) = recCDCgbPf.CGCODC
Mid$(MsgTxt, K + 16, 17) = Format$(recCDCgbPf.CGCOTH * 100, "00000000000000000")
Mid$(MsgTxt, K + 33, 17) = Format$(recCDCgbPf.CGCOEN * 100, "00000000000000000")
Mid$(MsgTxt, K + 50, 17) = Format$(recCDCgbPf.CGCOTP * 100, "00000000000000000")
Mid$(MsgTxt, K + 67, 17) = Format$(recCDCgbPf.CGCOAP * 100, "00000000000000000")
Mid$(MsgTxt, K + 84, 17) = Format$(recCDCgbPf.CGCODF * 100, "00000000000000000")

MsgTxtLen = MsgTxtLen + recCDCgbPfLen

End Sub



'---------------------------------------------------------
Private Function srvCDCgbPf_Seek(recCDCgbPf As typeCDCgbPf)
'---------------------------------------------------------

srvCDCgbPf_Seek = "?"
MsgTxtLen = 0
Call srvCDCgbPf_PutBuffer(recCDCgbPf)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvCDCgbPf_GetBuffer(recCDCgbPf)) Then
        srvCDCgbPf_Seek = Null
    Else
        Call srvCDCgbPf_Error(recCDCgbPf)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvCDCgbPf_Snap(recCDCgbPf As typeCDCgbPf)
'---------------------------------------------------------
srvCDCgbPf_Snap = "?"
MsgTxtLen = 0
Call srvCDCgbPf_PutBuffer(recCDCgbPf)
Call srvCDCgbPf_PutBuffer(arrCDCgbPf(0))
If IsNull(SndRcv()) Then
    srvCDCgbPf_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvCDCgbPf_GetBuffer(recCDCgbPf)) Then
            Call arrCDCgbPf_AddItem(recCDCgbPf)
            arrCDCgbPf_Suite = True
        Else
            arrCDCgbPf_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recCDCgbPf_Init(recCDCgbPf As typeCDCgbPf)
'---------------------------------------------------------
MsgTxt = Space$(recCDCgbPfLen)
MsgTxtIndex = 0
Call srvCDCgbPf_GetBuffer(recCDCgbPf)
recCDCgbPf.obj = "SRVCDCGBPF"
End Sub

'---------------------------------------------------------
Public Sub arrCDCgbPf_AddItem(recCDCgbPf As typeCDCgbPf)
'---------------------------------------------------------
          
arrCDCgbPf_Nb = arrCDCgbPf_Nb + 1
    
If arrCDCgbPf_Nb > arrCDCgbPf_NbMax Then
    arrCDCgbPf_NbMax = arrCDCgbPf_NbMax + 10
    ReDim Preserve arrCDCgbPf(arrCDCgbPf_NbMax)
End If
            
arrCDCgbPf(arrCDCgbPf_Nb) = recCDCgbPf
End Sub


