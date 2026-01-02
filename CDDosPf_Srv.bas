Attribute VB_Name = "srvCDDosPf"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recCDDosPfLen = 369 ' 34 + 335
Public Const MemoCDDosPfLen = 335
Public Const recCDDosPf_Block = 50

Type typeCDDosPf
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    DODKEY                  As Long
    DODPFX                  As String * 3
    DODNUM                  As Long
    DOREF                   As String * 20
    DODCTR                  As String * 8
    DODEXP                  As String * 8
    DOSTAT                  As String * 4
    DOREAC                  As String * 1
    DOREV                   As String * 1
    DONAT                   As String * 1
    DOUSC1                  As String * 2
    DOUSC2                  As String * 3
    DOUSC3                  As String * 3
    DONBEV                  As Long
    DOGPER                  As Double
    DOGAMT                  As Currency
    DOGCCY                  As String * 3
    DOGTIB                  As String * 4
    DOGTIN                  As String * 6
    DOGTIS                  As String * 3
    DOGCPT                  As String * 25
    DOAMT                   As Currency
    DOCCY                   As String * 3
    DOCPER                  As Double
    DOPLUS                  As Double
    DOMINS                  As Double
    DOQUA                   As String * 1
    DOOUTS                  As Currency
    DOLIAB                  As Currency
    DOLCCY                  As String * 3
    DOBNKY                  As Long
    DOBNNC                  As String * 6
    DOAPKY                  As Long
    DOAPNC                  As String * 6
    DORCKY                  As Long
    DORCNC                  As String * 6
    DOISKY                  As Long
    DOISNC                  As String * 6
    DOIBRC                  As String * 4
    DOBBRC                  As String * 4
    RBUSER                  As String * 20
    RBDCRT                  As String * 8
    RBDLUP                  As String * 8
End Type
    
Public arrCDDosPf() As typeCDDosPf
Public arrCDDosPf_NB As Integer
Public arrCDDosPf_NBMax As Integer
Public arrCDDosPf_Index As Integer
Public arrCDDosPf_Suite As Boolean

Public xCDDosPf As typeCDDosPf

'-----------------------------------------------------
Function srvCDDosPf_Update(recCDDosPf As typeCDDosPf)
'-----------------------------------------------------

srvCDDosPf_Update = "?"

MsgTxtLen = 0
Call srvCDDosPf_PutBuffer(recCDDosPf)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCDDosPf_GetBuffer(recCDDosPf)) Then
        Call srvCDDosPf_Error(recCDDosPf)
        srvCDDosPf_Update = recCDDosPf.Err
        Exit Function
    Else
        srvCDDosPf_Update = Null
    End If
Else
    recCDDosPf.Err = "srv"
End If


'=====================================================
End Function



Public Sub srvCDDosPf_Load(recCDDosPfMin As typeCDDosPf, recCDDosPfMax As typeCDDosPf)
Dim mMethod As String

mMethod = Trim(recCDDosPfMin.Method) & "+"
arrCDDosPf_NBMax = 0
arrCDDosPf_Suite = True: arrCDDosPf_NB = 0
arrCDDosPf_NBMax = recCDDosPf_Block: ReDim arrCDDosPf(arrCDDosPf_NBMax)

arrCDDosPf(0) = recCDDosPfMax
arrCDDosPf_Suite = True
Do Until Not arrCDDosPf_Suite
    srvCDDosPf_Monitor recCDDosPfMin
    recCDDosPfMin = arrCDDosPf(arrCDDosPf_NB)
    recCDDosPfMin.Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Function srvCDDosPf_Dtaq_Put(lFct As String, recCDDosPf As typeCDDosPf)
'-----------------------------------------------------

srvCDDosPf_Dtaq_Put = Null
Select Case lFct
    Case "Init": MsgTxtLen = 0
    Case "Add": Call srvCDDosPf_PutBuffer(recCDDosPf)
                If MsgTxtLen + recCDDosPfLen >= recCDDosPf_Block * recCDDosPfLen Then
                    Call srvCDDosPf_Dtaq_Snd(recCDDosPf): MsgTxtLen = 0
                End If
    Case "Snd": If MsgTxtLen > 0 Then Call srvCDDosPf_Dtaq_Snd(recCDDosPf)
    Case Else: srvCDDosPf_Dtaq_Put = lFct
End Select
'=====================================================
End Function


'-----------------------------------------------------
Function srvCDDosPf_Dtaq_Snd(recCDDosPf As typeCDDosPf)
'-----------------------------------------------------

srvCDDosPf_Dtaq_Snd = "?"

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCDDosPf_GetBuffer(recCDDosPf)) Then
        Call srvCDDosPf_Error(recCDDosPf)
        srvCDDosPf_Dtaq_Snd = recCDDosPf.Err
        Exit Function
    Else
        srvCDDosPf_Dtaq_Snd = Null
    End If
Else
    recCDDosPf.Err = "Snd"
End If


'=====================================================
End Function



'-----------------------------------------------------
Public Function srvCDDosPf_Monitor(recCDDosPf As typeCDDosPf)
'-----------------------------------------------------
blnFR_Convert = False

arrCDDosPf_Suite = False
Select Case mId$(Trim(recCDDosPf.Method), 1, 4)
    Case "Seek"
                srvCDDosPf_Monitor = srvCDDosPf_Seek(recCDDosPf)
    Case "Snap"
              srvCDDosPf_Monitor = srvCDDosPf_Snap(recCDDosPf)
    Case Else
                recCDDosPf.Err = recCDDosPf.Method
                Call srvCDDosPf_Error(recCDDosPf)
                srvCDDosPf_Monitor = recCDDosPf.Err
End Select

End Function

'-----------------------------------------------------
Sub srvCDDosPf_Error(recCDDosPf As typeCDDosPf)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "CDDosPf" & Chr$(10) & Chr$(13)

Select Case mId$(recCDDosPf.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recCDDosPf.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recCDDosPf.DODKEY _
        , I, "module : CDDosPfs.bas  ( " & Trim(recCDDosPf.obj) & " : " & Trim(recCDDosPf.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvCDDosPf_GetBuffer(recCDDosPf As typeCDDosPf)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvCDDosPf_GetBuffer = Null
recCDDosPf.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recCDDosPf.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recCDDosPf.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recCDDosPf.Err = Space$(10) Then
    recCDDosPf.DODKEY = CLng(Val(mId$(MsgTxt, K + 1, 12)))
    recCDDosPf.DODPFX = mId$(MsgTxt, K + 13, 3)
    recCDDosPf.DODNUM = CLng(Val(mId$(MsgTxt, K + 16, 6)))
    recCDDosPf.DOREF = mId$(MsgTxt, K + 22, 20)
    recCDDosPf.DODCTR = mId$(MsgTxt, K + 42, 8)
    recCDDosPf.DODEXP = mId$(MsgTxt, K + 50, 8)
    recCDDosPf.DOSTAT = mId$(MsgTxt, K + 58, 4)
    recCDDosPf.DOREAC = mId$(MsgTxt, K + 62, 1)
    recCDDosPf.DOREV = mId$(MsgTxt, K + 63, 1)
    recCDDosPf.DONAT = mId$(MsgTxt, K + 64, 1)
    recCDDosPf.DOUSC1 = mId$(MsgTxt, K + 65, 2)
    recCDDosPf.DOUSC2 = mId$(MsgTxt, K + 67, 3)
    recCDDosPf.DOUSC3 = mId$(MsgTxt, K + 70, 3)
    recCDDosPf.DONBEV = CLng(Val(mId$(MsgTxt, K + 73, 3)))
    recCDDosPf.DOGPER = CDbl(Val(mId$(MsgTxt, K + 76, 7)) / 100)
    recCDDosPf.DOGAMT = CCur(Val(mId$(MsgTxt, K + 83, 17)) / 100)
    recCDDosPf.DOGCCY = mId$(MsgTxt, K + 100, 3)
    recCDDosPf.DOGTIB = mId$(MsgTxt, K + 103, 4)
    recCDDosPf.DOGTIN = mId$(MsgTxt, K + 107, 6)
    recCDDosPf.DOGTIS = mId$(MsgTxt, K + 113, 3)
    recCDDosPf.DOGCPT = mId$(MsgTxt, K + 116, 25)
    recCDDosPf.DOAMT = CCur(Val(mId$(MsgTxt, K + 141, 17)) / 100)
    recCDDosPf.DOCCY = mId$(MsgTxt, K + 158, 3)
    recCDDosPf.DOCPER = CDbl(Val(mId$(MsgTxt, K + 161, 7)) / 100)
    recCDDosPf.DOPLUS = CDbl(Val(mId$(MsgTxt, K + 168, 7)) / 100)
    recCDDosPf.DOMINS = CDbl(Val(mId$(MsgTxt, K + 185, 7)) / 100)
    recCDDosPf.DOQUA = mId$(MsgTxt, K + 182, 1)
    recCDDosPf.DOOUTS = CCur(Val(mId$(MsgTxt, K + 183, 17)) / 100)
    recCDDosPf.DOLIAB = CCur(Val(mId$(MsgTxt, K + 200, 17)) / 100)
    recCDDosPf.DOLCCY = mId$(MsgTxt, K + 217, 3)
    recCDDosPf.DOBNKY = CLng(Val(mId$(MsgTxt, K + 220, 12)))
    recCDDosPf.DOBNNC = mId$(MsgTxt, K + 232, 6)
    recCDDosPf.DOAPKY = CLng(Val(mId$(MsgTxt, K + 238, 12)))
    recCDDosPf.DOAPNC = mId$(MsgTxt, K + 250, 6)
    recCDDosPf.DORCKY = CLng(Val(mId$(MsgTxt, K + 256, 12)))
    recCDDosPf.DORCNC = mId$(MsgTxt, K + 268, 6)
    recCDDosPf.DOISKY = CLng(Val(mId$(MsgTxt, K + 274, 12)))
    recCDDosPf.DOISNC = mId$(MsgTxt, K + 286, 6)
    recCDDosPf.DOIBRC = mId$(MsgTxt, K + 292, 4)
    recCDDosPf.DOBBRC = mId$(MsgTxt, K + 296, 4)
    recCDDosPf.RBUSER = mId$(MsgTxt, K + 300, 20)
    recCDDosPf.RBDCRT = mId$(MsgTxt, K + 320, 8)
    recCDDosPf.RBDLUP = mId$(MsgTxt, K + 328, 8)

Else
    srvCDDosPf_GetBuffer = recCDDosPf.Err
End If

MsgTxtIndex = MsgTxtIndex + recCDDosPfLen

End Function

'---------------------------------------------------------
Public Sub srvCDDosPf_PutBuffer(recCDDosPf As typeCDDosPf)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recCDDosPf.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recCDDosPf.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 12) = Format$(recCDDosPf.DODKEY, "000000000000")

Mid$(MsgTxt, K + 13, 3) = recCDDosPf.DODPFX
Mid$(MsgTxt, K + 16, 6) = Format$(recCDDosPf.DODNUM, "000000")
 Mid$(MsgTxt, K + 22, 20) = recCDDosPf.DOREF
 Mid$(MsgTxt, K + 42, 8) = recCDDosPf.DODCTR
 Mid$(MsgTxt, K + 50, 8) = recCDDosPf.DODEXP
 Mid$(MsgTxt, K + 58, 4) = recCDDosPf.DOSTAT
 Mid$(MsgTxt, K + 62, 1) = recCDDosPf.DOREAC
 Mid$(MsgTxt, K + 63, 1) = recCDDosPf.DOREV
 Mid$(MsgTxt, K + 64, 1) = recCDDosPf.DONAT
 Mid$(MsgTxt, K + 65, 2) = recCDDosPf.DOUSC1
 Mid$(MsgTxt, K + 67, 3) = recCDDosPf.DOUSC2
 Mid$(MsgTxt, K + 70, 3) = recCDDosPf.DOUSC3
 Mid$(MsgTxt, K + 73, 3) = Format$(recCDDosPf.DONBEV, "000")
 Mid$(MsgTxt, K + 76, 7) = Format$(recCDDosPf.DOGPER, "0000000")
 Mid$(MsgTxt, K + 83, 17) = Format$(recCDDosPf.DOGAMT * 100, "00000000000000000")
 Mid$(MsgTxt, K + 100, 3) = recCDDosPf.DOGCCY
 Mid$(MsgTxt, K + 103, 4) = recCDDosPf.DOGTIB
 Mid$(MsgTxt, K + 107, 6) = recCDDosPf.DOGTIN
 Mid$(MsgTxt, K + 113, 3) = recCDDosPf.DOGTIS
 Mid$(MsgTxt, K + 116, 25) = recCDDosPf.DOGCPT
 Mid$(MsgTxt, K + 141, 17) = Format$(recCDDosPf.DOAMT * 100, "00000000000000000")
 Mid$(MsgTxt, K + 158, 3) = recCDDosPf.DOCCY
 Mid$(MsgTxt, K + 161, 7) = Format$(recCDDosPf.DOCPER * 100, "0000000")
 Mid$(MsgTxt, K + 168, 7) = Format$(recCDDosPf.DOPLUS * 100, "0000000")
 Mid$(MsgTxt, K + 185, 7) = Format$(recCDDosPf.DOMINS * 100, "0000000")
 Mid$(MsgTxt, K + 182, 1) = recCDDosPf.DOQUA
 Mid$(MsgTxt, K + 183, 17) = Format$(recCDDosPf.DOOUTS * 100, "00000000000000000")
 Mid$(MsgTxt, K + 200, 17) = Format$(recCDDosPf.DOLIAB * 100, "00000000000000000")
 Mid$(MsgTxt, K + 217, 3) = recCDDosPf.DOLCCY
 Mid$(MsgTxt, K + 220, 12) = Format$(recCDDosPf.DOBNKY, "000000000000")
 Mid$(MsgTxt, K + 232, 6) = recCDDosPf.DOBNNC
 Mid$(MsgTxt, K + 238, 12) = Format$(recCDDosPf.DOAPKY, "000000000000")
 Mid$(MsgTxt, K + 250, 6) = recCDDosPf.DOAPNC
 Mid$(MsgTxt, K + 256, 12) = Format$(recCDDosPf.DORCKY, "000000000000")
 Mid$(MsgTxt, K + 268, 6) = recCDDosPf.DORCNC
 Mid$(MsgTxt, K + 274, 12) = Format$(recCDDosPf.DOISKY, "000000000000")
 Mid$(MsgTxt, K + 286, 6) = recCDDosPf.DOISNC
 Mid$(MsgTxt, K + 292, 4) = recCDDosPf.DOIBRC
 Mid$(MsgTxt, K + 296, 4) = recCDDosPf.DOBBRC
 Mid$(MsgTxt, K + 300, 20) = recCDDosPf.RBUSER
 Mid$(MsgTxt, K + 320, 8) = recCDDosPf.RBDCRT
 Mid$(MsgTxt, K + 328, 8) = recCDDosPf.RBDLUP


MsgTxtLen = MsgTxtLen + recCDDosPfLen

  
End Sub



'---------------------------------------------------------
Private Function srvCDDosPf_Seek(recCDDosPf As typeCDDosPf)
'---------------------------------------------------------

srvCDDosPf_Seek = "?"
MsgTxtLen = 0
Call srvCDDosPf_PutBuffer(recCDDosPf)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvCDDosPf_GetBuffer(recCDDosPf)) Then
        srvCDDosPf_Seek = Null
    Else
        Call srvCDDosPf_Error(recCDDosPf)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvCDDosPf_Snap(recCDDosPf As typeCDDosPf)
'---------------------------------------------------------
srvCDDosPf_Snap = "?"
MsgTxtLen = 0
Call srvCDDosPf_PutBuffer(recCDDosPf)
Call srvCDDosPf_PutBuffer(arrCDDosPf(0))
If IsNull(SndRcv()) Then
    srvCDDosPf_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvCDDosPf_GetBuffer(recCDDosPf)) Then
            Call arrCDDosPf_AddItem(recCDDosPf)
            arrCDDosPf_Suite = True
        Else
            arrCDDosPf_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recCDDosPf_Init(recCDDosPf As typeCDDosPf)
'---------------------------------------------------------
MsgTxt = Space$(recCDDosPfLen)
MsgTxtIndex = 0
Call srvCDDosPf_GetBuffer(recCDDosPf)
recCDDosPf.obj = "SRVCDDOSPF"

End Sub

'---------------------------------------------------------
Public Sub arrCDDosPf_AddItem(recCDDosPf As typeCDDosPf)
'---------------------------------------------------------
          
arrCDDosPf_NB = arrCDDosPf_NB + 1
    
If arrCDDosPf_NB > arrCDDosPf_NBMax Then
    arrCDDosPf_NBMax = arrCDDosPf_NBMax + recCDDosPf_Block
    ReDim Preserve arrCDDosPf(arrCDDosPf_NBMax)
End If
            
arrCDDosPf(arrCDDosPf_NB) = recCDDosPf
End Sub


