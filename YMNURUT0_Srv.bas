Attribute VB_Name = "srvYMNURUT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const memoYMNURUT0Len = 51
Public Const recYMNURUT0Len = 85 ' 34 + 51
Public Const recYMNURUT0_Block = 200
Type typeYMNURUT0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    MNURUTUTI       As String * 10                    ' UTILISATEUR
    MNURUTNOM       As String * 30                    ' NOM
    MNURUTETB       As Integer                        ' ETAB. PAR DEFAUT
    MNURUTCUT       As Integer                        ' CODE INTERNE
    MNURUTLOG       As String * 1                     ' ENTREE LOGICIEL
  
End Type
    
Public Sub srvYMNURUT0_Load(lYMNURUT0() As typeYMNURUT0, lYMNURUT0_Nb As Integer)
Dim mMethod As String, blnYMNURUT0_Suite
Dim wNbMax As Integer
Dim wYMNURUT0 As typeYMNURUT0

mMethod = Trim(lYMNURUT0(0).Method) & "+"
blnYMNURUT0_Suite = True: lYMNURUT0_Nb = 0
wNbMax = recYMNURUT0_Block + 2: ReDim Preserve lYMNURUT0(wNbMax)

wYMNURUT0 = lYMNURUT0(1)
Do Until Not blnYMNURUT0_Suite
    MsgTxtLen = 0
    Call srvYMNURUT0_PutBuffer(wYMNURUT0)
    Call srvYMNURUT0_PutBuffer(lYMNURUT0(0))
    If IsNull(SndRcv()) Then
        MsgTxtIndex = 0
        Do While MsgTxtIndex < MsgTxtLen
            If IsNull(srvYMNURUT0_GetBuffer(wYMNURUT0)) Then
            
                lYMNURUT0_Nb = lYMNURUT0_Nb + 1
                If lYMNURUT0_Nb > wNbMax Then
                    wNbMax = wNbMax + recYMNURUT0_Block
                    ReDim Preserve lYMNURUT0(wNbMax)
                End If
            
                lYMNURUT0(lYMNURUT0_Nb) = wYMNURUT0
                blnYMNURUT0_Suite = True
            Else
                blnYMNURUT0_Suite = False
                Exit Do
            End If
        Loop
    End If

    lYMNURUT0(0).Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Public Function srvYMNURUT0_Monitor(recYMNURUT0 As typeYMNURUT0)
'-----------------------------------------------------

Select Case mId$(Trim(recYMNURUT0.Method), 1, 4)
    Case "Seek"
                srvYMNURUT0_Monitor = srvYMNURUT0_Seek(recYMNURUT0)
    Case Else
                recYMNURUT0.Err = recYMNURUT0.Method
                Call srvYMNURUT0_Error(recYMNURUT0)
                srvYMNURUT0_Monitor = recYMNURUT0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYMNURUT0_Error(recYMNURUT0 As typeYMNURUT0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YMNURUT0" & Chr$(10) & Chr$(13)

Select Case mId$(recYMNURUT0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYMNURUT0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YMNURUT0s.bas  ( " _
                & Trim(recYMNURUT0.obj) & " : " & Trim(recYMNURUT0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYMNURUT0_GetBuffer(recYMNURUT0 As typeYMNURUT0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYMNURUT0_GetBuffer = Null
recYMNURUT0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYMNURUT0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYMNURUT0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYMNURUT0.Err = Space$(10) Then
    recYMNURUT0.MNURUTUTI = mId$(MsgTxt, K + 1, 10)
    recYMNURUT0.MNURUTNOM = mId$(MsgTxt, K + 11, 30)
    recYMNURUT0.MNURUTETB = CInt(Val(mId$(MsgTxt, K + 41, 5)))
    recYMNURUT0.MNURUTCUT = CInt(Val(mId$(MsgTxt, K + 46, 5)))
    recYMNURUT0.MNURUTLOG = mId$(MsgTxt, K + 51, 1)
Else
    srvYMNURUT0_GetBuffer = recYMNURUT0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYMNURUT0Len

End Function

'---------------------------------------------------------
Private Sub srvYMNURUT0_PutBuffer(recYMNURUT0 As typeYMNURUT0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYMNURUT0Len) = Space$(recYMNURUT0Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYMNURUT0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYMNURUT0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 10) = recYMNURUT0.MNURUTUTI
    Mid$(MsgTxt, K + 11, 30) = recYMNURUT0.MNURUTNOM
    Mid$(MsgTxt, K + 41, 5) = Format$(recYMNURUT0.MNURUTETB, "0000 ")
    Mid$(MsgTxt, K + 46, 5) = Format$(recYMNURUT0.MNURUTCUT, "0000 ")
    Mid$(MsgTxt, K + 51, 1) = recYMNURUT0.MNURUTLOG

MsgTxtLen = MsgTxtLen + recYMNURUT0Len
End Sub

Public Sub srvYMNURUT0_ElpDisplay(recYMNURUT0 As typeYMNURUT0)
frmElpDisplay.fgData.Rows = 6
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNURUTUTI   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "UTILISATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNURUT0.MNURUTUTI
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNURUTNOM   30A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "NOM"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNURUT0.MNURUTNOM
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNURUTETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETAB. PAR DEFAUT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNURUT0.MNURUTETB
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNURUTCUT    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE INTERNE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNURUT0.MNURUTCUT
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNURUTLOG    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ENTREE LOGICIEL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNURUT0.MNURUTLOG
frmElpDisplay.Show vbModal
End Sub

'---------------------------------------------------------
Private Function srvYMNURUT0_Seek(recYMNURUT0 As typeYMNURUT0)
'---------------------------------------------------------

srvYMNURUT0_Seek = "?"
MsgTxtLen = 0
Call srvYMNURUT0_PutBuffer(recYMNURUT0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYMNURUT0_GetBuffer(recYMNURUT0)) Then
            srvYMNURUT0_Seek = Null
        Else
            Call srvYMNURUT0_Error(recYMNURUT0)
        End If
    End If
End If

End Function

'-----------------------------------------------------
Function srvYMNURUT0_Update(recYMNURUT0 As typeYMNURUT0)
'-----------------------------------------------------

srvYMNURUT0_Update = "?"

MsgTxtLen = 0
Call srvYMNURUT0_PutBuffer(recYMNURUT0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYMNURUT0_GetBuffer(recYMNURUT0)) Then
        Call srvYMNURUT0_Error(recYMNURUT0)
        srvYMNURUT0_Update = recYMNURUT0.Err
        Exit Function
    Else
        srvYMNURUT0_Update = Null
    End If
Else
    recYMNURUT0.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYMNURUT0_Init(recYMNURUT0 As typeYMNURUT0)
'---------------------------------------------------------
MsgTxt = Space$(recYMNURUT0Len)
MsgTxtIndex = 0
Call srvYMNURUT0_GetBuffer(recYMNURUT0)
recYMNURUT0.obj = "ZMNURUT0_S"
recYMNURUT0.MNURUTETB = 1

End Sub





