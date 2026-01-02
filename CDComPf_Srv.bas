Attribute VB_Name = "srvCDComPf"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recCDComPfLen = 346 ' 312 + 34
Public Const recCDComPf_Block = 400

Type typeCDComPf
    obj         As String * 12
    Method      As String * 12
    Err         As String * 10
    
    COCENR      As String * 1
    CODPFX      As String * 3
    CODNUM      As Long
    COCODC      As String * 2
    CODEFF      As String * 8
    COFEFF      As String * 8
    CONBJ       As String * 5
    COFETH      As String * 8
    CONBTH      As String * 5
    COTAUX      As Double
    COMETH      As String * 2
    COCMIN      As Currency
    COMCCY      As String * 3
    COCPNC      As String * 6
    CODCCY      As String * 3
    COMVTD      As Currency
    COMVTC      As Currency
    CODBAS      As Currency
    CODCOM      As Currency
    CODTVA      As Currency
    CODCMP      As Currency
    COCOUR      As Double
    COCCCY      As String * 3
    COCBAS      As Currency
    COCCOM      As Currency
    COCTVA      As Currency
    COCCMP      As Currency
    CODCRT      As String * 8
    CODLUP      As String * 8
    COUSER      As String * 20

End Type

Public arrCDComPf() As typeCDComPf
Public arrCDComPf_Nb As Integer
Public arrCDComPf_NbMax As Integer
Public arrCDComPf_Index As Integer
Public arrCDComPf_Suite As Boolean

Public Sub srvCDComPf_ElpDisplay(recCDComPf As typeCDComPf)
frmElpDisplay.fgData.Rows = 18
frmElpDisplay.fgData.Row = 1
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "obj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.obj
frmElpDisplay.fgData.Row = 2
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Method"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.Method
frmElpDisplay.fgData.Row = 3
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Err"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.Err
    
frmElpDisplay.fgData.Row = 4
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COCENR"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COCENR
frmElpDisplay.fgData.Row = 5
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CODPFX"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.CODPFX
frmElpDisplay.fgData.Row = 6
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CODNUM"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.CODNUM
frmElpDisplay.fgData.Row = 7
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COCODC "
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COCODC
frmElpDisplay.fgData.Row = 8
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CODEFF"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.CODEFF
frmElpDisplay.fgData.Row = 9
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COFEFF"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COFEFF
frmElpDisplay.fgData.Row = 10
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CONBJ "
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.CONBJ
frmElpDisplay.fgData.Row = 11
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COFETH"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COFETH
frmElpDisplay.fgData.Row = 12
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CONBTH"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.CONBTH
frmElpDisplay.fgData.Row = 13
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COTAUX"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COTAUX
frmElpDisplay.fgData.Row = 14
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMETH"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COMETH
frmElpDisplay.fgData.Row = 15
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COCMIN"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COCMIN
frmElpDisplay.fgData.Row = 16
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMCCY"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COMCCY
frmElpDisplay.fgData.Row = 17
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COCPNC"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COCPNC
frmElpDisplay.fgData.Row = 18
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CODCCY"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.CODCCY
frmElpDisplay.fgData.Row = 19
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMVTD"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COMVTD
frmElpDisplay.fgData.Row = 20
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COMVTC"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COMVTC
frmElpDisplay.fgData.Row = 21
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CODBAS"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.CODBAS
frmElpDisplay.fgData.Row = 22
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CODCOM"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.CODCOM
frmElpDisplay.fgData.Row = 23
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CODTVA"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.CODTVA
frmElpDisplay.fgData.Row = 24
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CODCMP"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.CODCMP
frmElpDisplay.fgData.Row = 25
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COCOUR"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COCOUR
frmElpDisplay.fgData.Row = 26
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COCCCY"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COCCCY
frmElpDisplay.fgData.Row = 27
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COCBAS"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COCBAS
frmElpDisplay.fgData.Row = 28
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COCCOM"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COCCOM
frmElpDisplay.fgData.Row = 29
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COCTVA"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COCTVA
frmElpDisplay.fgData.Row = 30
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COCCMP"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COCCMP
frmElpDisplay.fgData.Row = 31
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CODCRT"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.CODCRT
frmElpDisplay.fgData.Row = 32
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "CODLUP"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.CODLUP
frmElpDisplay.fgData.Row = 33
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "COUSER"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recCDComPf.COUSER

frmElpDisplay.Show vbModal

End Sub

Public Sub srvCDComPf_Load(recCDComPfMin As typeCDComPf, recCDComPfMax As typeCDComPf)
Dim mMethod As String

mMethod = Trim(recCDComPfMin.Method) & "+"
arrCDComPf_NbMax = 0
arrCDComPf_Suite = True: arrCDComPf_Nb = 0
arrCDComPf_NbMax = recCDComPf_Block: ReDim arrCDComPf(arrCDComPf_NbMax)

arrCDComPf(0) = recCDComPfMax
arrCDComPf_Suite = True
Do Until Not arrCDComPf_Suite
    srvCDComPf_Monitor recCDComPfMin
    recCDComPfMin = arrCDComPf(arrCDComPf_Nb)
    recCDComPfMin.Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Function srvCDComPf_Update(recCDComPf As typeCDComPf)
'-----------------------------------------------------

srvCDComPf_Update = "?"

MsgTxtLen = 0
Call srvCDComPf_PutBuffer(recCDComPf)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCDComPf_GetBuffer(recCDComPf)) Then
        Call srvCDComPf_Error(recCDComPf)
        srvCDComPf_Update = recCDComPf.Err
        Exit Function
    Else
        srvCDComPf_Update = Null
    End If
Else
    recCDComPf.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Public Function srvCDComPf_Monitor(recCDComPf As typeCDComPf)
'-----------------------------------------------------

blnFR_Convert = False

arrCDComPf_Suite = False
Select Case mId$(Trim(recCDComPf.Method), 1, 4)
    Case "Seek", "Comp"
                srvCDComPf_Monitor = srvCDComPf_Seek(recCDComPf)
    Case "Snap"
              srvCDComPf_Monitor = srvCDComPf_Snap(recCDComPf)
    Case Else
                recCDComPf.Err = recCDComPf.Method
                Call srvCDComPf_Error(recCDComPf)
                srvCDComPf_Monitor = recCDComPf.Err
End Select
End Function

'-----------------------------------------------------
Sub srvCDComPf_Error(recCDComPf As typeCDComPf)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "CDComPf" & Chr$(10) & Chr$(13)

Select Case mId$(recCDComPf.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recCDComPf.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : CDComPf_Srv.bas  ( " _
                & Trim(recCDComPf.obj) & " : " & Trim(recCDComPf.Method) & " )"

End Sub

'---------------------------------------------------------
Public Function srvCDComPf_GetBuffer(recCDComPf As typeCDComPf)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvCDComPf_GetBuffer = Null
recCDComPf.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recCDComPf.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recCDComPf.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recCDComPf.Err = Space$(10) Then
  
    recCDComPf.COCENR = mId$(MsgTxt, K + 1, 1)
    recCDComPf.CODPFX = mId$(MsgTxt, K + 2, 3)
    recCDComPf.CODNUM = CLng(Val(mId$(MsgTxt, K + 5, 6)))
    recCDComPf.COCODC = mId$(MsgTxt, K + 11, 2)
    recCDComPf.CODEFF = Format$(Val(mId$(MsgTxt, K + 13, 8)), "00000000")
    recCDComPf.COFEFF = Format$(Val(mId$(MsgTxt, K + 21, 8)), "00000000")
    recCDComPf.CONBJ = Format$(Val(mId$(MsgTxt, K + 29, 5)), "00000")
    recCDComPf.COFETH = Format$(Val(mId$(MsgTxt, K + 34, 8)), "00000000")
    recCDComPf.CONBTH = Format$(Val(mId$(MsgTxt, K + 42, 5)), "00000")
    recCDComPf.COTAUX = CDbl(Val(mId$(MsgTxt, K + 47, 11)) / 10000000)
    recCDComPf.COMETH = mId$(MsgTxt, K + 58, 2)
    recCDComPf.COCMIN = CCur(Val(mId$(MsgTxt, K + 60, 17)) / 100)
    recCDComPf.COMCCY = mId$(MsgTxt, K + 77, 3)
    recCDComPf.COCPNC = Format$(Val(mId$(MsgTxt, K + 80, 6)), "000000")
    recCDComPf.CODCCY = mId$(MsgTxt, K + 86, 3)
    recCDComPf.COMVTD = CCur(Val(mId$(MsgTxt, K + 89, 17)) / 100)
    recCDComPf.COMVTC = CCur(Val(mId$(MsgTxt, K + 106, 17)) / 100)
    recCDComPf.CODBAS = CCur(Val(mId$(MsgTxt, K + 123, 17)) / 100)
    recCDComPf.CODCOM = CCur(Val(mId$(MsgTxt, K + 140, 17)) / 100)
    recCDComPf.CODTVA = CCur(Val(mId$(MsgTxt, K + 157, 17)) / 100)
    recCDComPf.CODCMP = CCur(Val(mId$(MsgTxt, K + 174, 17)) / 100)
    recCDComPf.COCOUR = CDbl(Val(mId$(MsgTxt, K + 191, 15)) / 10000000)
    recCDComPf.COCCCY = mId$(MsgTxt, K + 206, 3)
    recCDComPf.COCBAS = CCur(Val(mId$(MsgTxt, K + 209, 17)) / 100)
    recCDComPf.COCCOM = CCur(Val(mId$(MsgTxt, K + 226, 17)) / 100)
    recCDComPf.COCTVA = CCur(Val(mId$(MsgTxt, K + 243, 17)) / 100)
    recCDComPf.COCCMP = CCur(Val(mId$(MsgTxt, K + 260, 17)) / 100)
    recCDComPf.CODCRT = Format$(Val(mId$(MsgTxt, K + 277, 8)), "00000000")
    recCDComPf.CODLUP = Format$(Val(mId$(MsgTxt, K + 285, 8)), "00000000")
    recCDComPf.COUSER = mId$(MsgTxt, K + 293, 20)
Else
    srvCDComPf_GetBuffer = recCDComPf.Err
End If

MsgTxtIndex = MsgTxtIndex + recCDComPfLen

End Function

'---------------------------------------------------------
Private Sub srvCDComPf_PutBuffer(recCDComPf As typeCDComPf)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recCDComPf.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recCDComPf.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 1) = recCDComPf.COCENR
Mid$(MsgTxt, K + 2, 3) = recCDComPf.CODPFX
Mid$(MsgTxt, K + 5, 6) = Format$(recCDComPf.CODNUM, "000000")
Mid$(MsgTxt, K + 11, 2) = recCDComPf.COCODC
Mid$(MsgTxt, K + 13, 8) = Format$(recCDComPf.CODEFF, "00000000")
Mid$(MsgTxt, K + 21, 8) = Format$(recCDComPf.COFEFF, "00000000")
Mid$(MsgTxt, K + 29, 5) = Format$(recCDComPf.CONBJ, "00000")
Mid$(MsgTxt, K + 34, 8) = Format$(recCDComPf.COFETH, "00000000")
Mid$(MsgTxt, K + 42, 5) = Format$(recCDComPf.CONBTH, "00000")
Mid$(MsgTxt, K + 47, 11) = Format$(recCDComPf.COTAUX * 10000000, "00000000000")
Mid$(MsgTxt, K + 58, 2) = recCDComPf.COMETH
Mid$(MsgTxt, K + 60, 17) = Format$(recCDComPf.COCMIN * 100, "00000000000000000")
Mid$(MsgTxt, K + 77, 3) = recCDComPf.COMCCY
Mid$(MsgTxt, K + 80, 6) = Format$(recCDComPf.CONBJ, "000000")
Mid$(MsgTxt, K + 86, 3) = recCDComPf.COMCCY
Mid$(MsgTxt, K + 89, 17) = Format$(recCDComPf.COMVTD * 100, "00000000000000000")
Mid$(MsgTxt, K + 106, 17) = Format$(recCDComPf.COMVTC * 100, "00000000000000000")
Mid$(MsgTxt, K + 123, 17) = Format$(recCDComPf.CODBAS * 100, "00000000000000000")
Mid$(MsgTxt, K + 140, 17) = Format$(recCDComPf.CODCOM * 100, "00000000000000000")
Mid$(MsgTxt, K + 157, 17) = Format$(recCDComPf.CODTVA * 100, "00000000000000000")
Mid$(MsgTxt, K + 174, 17) = Format$(recCDComPf.CODCMP * 100, "00000000000000000")
Mid$(MsgTxt, K + 191, 11) = Format$(recCDComPf.COCOUR * 10000000, "000000000000000")
Mid$(MsgTxt, K + 206, 3) = recCDComPf.COCCCY
Mid$(MsgTxt, K + 209, 17) = Format$(recCDComPf.COCBAS * 100, "00000000000000000")
Mid$(MsgTxt, K + 226, 17) = Format$(recCDComPf.COCCOM * 100, "00000000000000000")
Mid$(MsgTxt, K + 243, 17) = Format$(recCDComPf.COCTVA * 100, "00000000000000000")
Mid$(MsgTxt, K + 260, 17) = Format$(recCDComPf.COCCMP * 100, "00000000000000000")
Mid$(MsgTxt, K + 277, 8) = Format$(recCDComPf.CODCRT, "00000000")
Mid$(MsgTxt, K + 285, 8) = Format$(recCDComPf.CODLUP, "00000000")
Mid$(MsgTxt, K + 293, 20) = recCDComPf.COUSER

MsgTxtLen = MsgTxtLen + recCDComPfLen

End Sub

'---------------------------------------------------------
Private Function srvCDComPf_Seek(recCDComPf As typeCDComPf)
'---------------------------------------------------------

srvCDComPf_Seek = "?"
MsgTxtLen = 0
Call srvCDComPf_PutBuffer(recCDComPf)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvCDComPf_GetBuffer(recCDComPf)) Then
        srvCDComPf_Seek = Null
    Else
        Call srvCDComPf_Error(recCDComPf)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvCDComPf_Snap(recCDComPf As typeCDComPf)
'---------------------------------------------------------
srvCDComPf_Snap = "?"
MsgTxtLen = 0
Call srvCDComPf_PutBuffer(recCDComPf)
Call srvCDComPf_PutBuffer(arrCDComPf(0))
If IsNull(SndRcv()) Then
    srvCDComPf_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvCDComPf_GetBuffer(recCDComPf)) Then
            Call arrCDComPf_AddItem(recCDComPf)
            arrCDComPf_Suite = True
        Else
            arrCDComPf_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recCDComPf_Init(recCDComPf As typeCDComPf)
'---------------------------------------------------------
MsgTxt = Space$(recCDComPfLen)
MsgTxtIndex = 0
Call srvCDComPf_GetBuffer(recCDComPf)
recCDComPf.obj = "SRVCDCOMPF"
End Sub

'---------------------------------------------------------
Public Sub arrCDComPf_AddItem(recCDComPf As typeCDComPf)
'---------------------------------------------------------
          
arrCDComPf_Nb = arrCDComPf_Nb + 1
    
If arrCDComPf_Nb > arrCDComPf_NbMax Then
    arrCDComPf_NbMax = arrCDComPf_NbMax + 10
    ReDim Preserve arrCDComPf(arrCDComPf_NbMax)
End If
            
arrCDComPf(arrCDComPf_Nb) = recCDComPf
End Sub


