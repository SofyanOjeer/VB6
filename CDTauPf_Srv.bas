Attribute VB_Name = "srvCDTauPf"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recCDTauPfLen = 132 ' 34 + 526
Public Const recCDTauPf_Block = 200

Type typeCDTauPf
    obj         As String * 12
    Method      As String * 12
    Err         As String * 10
    
    TACENR      As String * 1
    TADPFX      As String * 3
    TADNUM      As Long
    TACODC      As String * 2
    TADEFF      As String * 8
    TAFEFF      As String * 8
    TATAUX      As Double
    TAFRQ       As String * 1
    TAMETH      As String * 2
    TACMIN      As Currency
    TACCCY      As String * 3
    TADCRT      As String * 8
    TADLUP      As String * 8
    TAUSER      As String * 20

End Type

Public arrCDTauPf() As typeCDTauPf
Public arrCDTauPf_Nb As Integer
Public arrCDTauPf_NbMax As Integer
Public arrCDTauPf_Index As Integer
Public arrCDTauPf_Suite As Boolean

Public Sub srvCDTauPf_ElpDisplay(recCDTauPf As typeCDTauPf)
frmElpDisplay.fgData.Rows = 18
frmElpDisplay.fgData.Row = 1
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "obj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.obj
frmElpDisplay.fgData.Row = 2
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Method"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.Method
frmElpDisplay.fgData.Row = 3
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Err"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.Err
frmElpDisplay.fgData.Row = 4
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TACENR"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.TACENR

frmElpDisplay.fgData.Row = 5
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TADPFX"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.TADPFX
frmElpDisplay.fgData.Row = 6
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TADNUM"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.TADNUM
frmElpDisplay.fgData.Row = 7
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TACODC "
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.TACODC
frmElpDisplay.fgData.Row = 8
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TADEFF"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.TADEFF
frmElpDisplay.fgData.Row = 9
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TAFEFF"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.TAFEFF
frmElpDisplay.fgData.Row = 10
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TATAUX"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.TATAUX
frmElpDisplay.fgData.Row = 11
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TAFRQ"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.TAFRQ
frmElpDisplay.fgData.Row = 12
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TAMETH"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.TAMETH
frmElpDisplay.fgData.Row = 13
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TACMIN "
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.TACMIN
frmElpDisplay.fgData.Row = 14
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TACCCY"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.TACCCY
frmElpDisplay.fgData.Row = 15
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TADCRT"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.TADCRT
frmElpDisplay.fgData.Row = 16
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TADLUP"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.TADLUP
frmElpDisplay.fgData.Row = 17
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TAUSER"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDTauPf.TAUSER

frmElpDisplay.Show vbModal

End Sub

Public Sub srvCDTauPf_Load(recCDTauPfMin As typeCDTauPf, recCDTauPfMax As typeCDTauPf)
Dim mMethod As String

mMethod = Trim(recCDTauPfMin.Method) & "+"
arrCDTauPf_NbMax = 0
arrCDTauPf_Suite = True: arrCDTauPf_Nb = 0
arrCDTauPf_NbMax = recCDTauPf_Block: ReDim arrCDTauPf(arrCDTauPf_NbMax)

arrCDTauPf(0) = recCDTauPfMax
arrCDTauPf_Suite = True
Do Until Not arrCDTauPf_Suite
    srvCDTauPf_Monitor recCDTauPfMin
    recCDTauPfMin = arrCDTauPf(arrCDTauPf_Nb)
    recCDTauPfMin.Method = mMethod
Loop

End Sub


'-----------------------------------------------------
Function srvCDTauPf_Update(recCDTauPf As typeCDTauPf)
'-----------------------------------------------------

srvCDTauPf_Update = "?"

MsgTxtLen = 0
Call srvCDTauPf_PutBuffer(recCDTauPf)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCDTauPf_GetBuffer(recCDTauPf)) Then
        Call srvCDTauPf_Error(recCDTauPf)
        srvCDTauPf_Update = recCDTauPf.Err
        Exit Function
    Else
        srvCDTauPf_Update = Null
    End If
Else
    recCDTauPf.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Public Function srvCDTauPf_Monitor(recCDTauPf As typeCDTauPf)
'-----------------------------------------------------

blnFR_Convert = False

arrCDTauPf_Suite = False
Select Case mId$(Trim(recCDTauPf.Method), 1, 4)
    Case "Seek", "Comp"
                srvCDTauPf_Monitor = srvCDTauPf_Seek(recCDTauPf)
    Case "Snap"
              srvCDTauPf_Monitor = srvCDTauPf_Snap(recCDTauPf)
    Case Else
                recCDTauPf.Err = recCDTauPf.Method
                Call srvCDTauPf_Error(recCDTauPf)
                srvCDTauPf_Monitor = recCDTauPf.Err
End Select
End Function

'-----------------------------------------------------
Sub srvCDTauPf_Error(recCDTauPf As typeCDTauPf)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "CDTauPf" & Chr$(10) & Chr$(13)

Select Case mId$(recCDTauPf.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recCDTauPf.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : CDTauPf_Srv.bas  ( " _
                & Trim(recCDTauPf.obj) & " : " & Trim(recCDTauPf.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvCDTauPf_GetBuffer(recCDTauPf As typeCDTauPf)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvCDTauPf_GetBuffer = Null
recCDTauPf.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recCDTauPf.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recCDTauPf.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recCDTauPf.Err = Space$(10) Then
  

    recCDTauPf.TACENR = mId$(MsgTxt, K + 1, 1)
    recCDTauPf.TADPFX = mId$(MsgTxt, K + 2, 3)
    recCDTauPf.TADNUM = CLng(Val(mId$(MsgTxt, K + 5, 6)))
    recCDTauPf.TACODC = mId$(MsgTxt, K + 11, 2)
    recCDTauPf.TADEFF = Format$(Val(mId$(MsgTxt, K + 13, 8)), "00000000")
    recCDTauPf.TAFEFF = Format$(Val(mId$(MsgTxt, K + 21, 8)), "00000000")
    recCDTauPf.TATAUX = CDbl(Val(mId$(MsgTxt, K + 29, 11)) / 10000000)
    recCDTauPf.TAFRQ = mId$(MsgTxt, K + 40, 1)
    recCDTauPf.TAMETH = mId$(MsgTxt, K + 41, 2)
    recCDTauPf.TACMIN = CCur(Val(mId$(MsgTxt, K + 43, 17)) / 100)
    recCDTauPf.TACCCY = mId$(MsgTxt, K + 60, 3)
    recCDTauPf.TADCRT = Format$(Val(mId$(MsgTxt, K + 63, 8)), "00000000")
    recCDTauPf.TADLUP = Format$(Val(mId$(MsgTxt, K + 71, 8)), "00000000")
    recCDTauPf.TAUSER = mId$(MsgTxt, K + 79, 20)
Else
    srvCDTauPf_GetBuffer = recCDTauPf.Err
End If

MsgTxtIndex = MsgTxtIndex + recCDTauPfLen

End Function

'---------------------------------------------------------
Private Sub srvCDTauPf_PutBuffer(recCDTauPf As typeCDTauPf)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recCDTauPf.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recCDTauPf.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 1) = recCDTauPf.TACENR
Mid$(MsgTxt, K + 2, 3) = recCDTauPf.TADPFX
Mid$(MsgTxt, K + 5, 6) = Format$(recCDTauPf.TADNUM, "000000")
Mid$(MsgTxt, K + 11, 2) = recCDTauPf.TACODC
Mid$(MsgTxt, K + 13, 8) = Format$(recCDTauPf.TADEFF, "00000000")
Mid$(MsgTxt, K + 21, 8) = Format$(recCDTauPf.TAFEFF, "00000000")
Mid$(MsgTxt, K + 29, 11) = Format$(recCDTauPf.TATAUX * 10000000, "00000000000")
Mid$(MsgTxt, K + 40, 1) = recCDTauPf.TAFRQ
Mid$(MsgTxt, K + 41, 2) = recCDTauPf.TAMETH
Mid$(MsgTxt, K + 43, 17) = Format$(recCDTauPf.TACMIN * 100, "00000000000000000")
Mid$(MsgTxt, K + 60, 3) = recCDTauPf.TACCCY
Mid$(MsgTxt, K + 63, 8) = Format$(recCDTauPf.TADCRT, "00000000")
Mid$(MsgTxt, K + 71, 8) = Format$(recCDTauPf.TADLUP, "00000000")
Mid$(MsgTxt, K + 79, 20) = recCDTauPf.TAUSER

MsgTxtLen = MsgTxtLen + recCDTauPfLen
End Sub



'---------------------------------------------------------
Private Function srvCDTauPf_Seek(recCDTauPf As typeCDTauPf)
'---------------------------------------------------------

srvCDTauPf_Seek = "?"
MsgTxtLen = 0
Call srvCDTauPf_PutBuffer(recCDTauPf)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvCDTauPf_GetBuffer(recCDTauPf)) Then
        srvCDTauPf_Seek = Null
    Else
        Call srvCDTauPf_Error(recCDTauPf)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvCDTauPf_Snap(recCDTauPf As typeCDTauPf)
'---------------------------------------------------------
srvCDTauPf_Snap = "?"
MsgTxtLen = 0
Call srvCDTauPf_PutBuffer(recCDTauPf)
Call srvCDTauPf_PutBuffer(arrCDTauPf(0))
If IsNull(SndRcv()) Then
    srvCDTauPf_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvCDTauPf_GetBuffer(recCDTauPf)) Then
            Call arrCDTauPf_AddItem(recCDTauPf)
            arrCDTauPf_Suite = True
        Else
            arrCDTauPf_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recCDTauPf_Init(recCDTauPf As typeCDTauPf)
'---------------------------------------------------------
MsgTxt = Space$(recCDTauPfLen)
MsgTxtIndex = 0
Call srvCDTauPf_GetBuffer(recCDTauPf)
recCDTauPf.obj = "SRVCDTAUPF  "
End Sub

'---------------------------------------------------------
Public Sub arrCDTauPf_AddItem(recCDTauPf As typeCDTauPf)
'---------------------------------------------------------
          
arrCDTauPf_Nb = arrCDTauPf_Nb + 1
    
If arrCDTauPf_Nb > arrCDTauPf_NbMax Then
    arrCDTauPf_NbMax = arrCDTauPf_NbMax + 10
    ReDim Preserve arrCDTauPf(arrCDTauPf_NbMax)
End If
            
arrCDTauPf(arrCDTauPf_Nb) = recCDTauPf
End Sub
