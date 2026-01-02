Attribute VB_Name = "srvYMNUOPT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYMNUOPT0Len = 107 ' 34 +73
Public Const recYMNUOPT0_Block = 300
Type typeYMNUOPT0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    MNUOPTCOD       As Long                           ' CODE OPTION
    MNUOPTCLI       As String * 7                     ' CLIENT
    MNUOPTLIB       As String * 35                    ' LIBELLE
    MNUOPTENS       As String * 8                     ' ENSEMBLE
    MNUOPTENT       As String * 8                     ' POINT ENTREE
    MNUOPTSTR       As String * 1                     ' OPTION STRAB
    MNUOPTARE       As String * 1                     ' ARRET LOGICIEL
    MNUOPTBAT       As String * 1                     ' OPTION BATCH
    MNUOPTVAL       As String * 1                     ' VALID. BATCH
    MNUOPTSUP       As String * 1                     ' A SUPPRIMER
    MNUOPTOIA       As String * 1                     ' INTER-AGENCE
    MNUOPTGES       As String * 1                     ' SECUR.INTER-ETAB.
  
End Type
    
Public Sub srvYMNUOPT0_Load(lYMNUOPT0() As typeYMNUOPT0, lYMNUOPT0_Nb As Integer)
Dim mMethod As String, blnYMNUOPT0_Suite
Dim wNbMax As Integer
Dim wYMNUOPT0 As typeYMNUOPT0

mMethod = Trim(lYMNUOPT0(0).Method) & "+"
blnYMNUOPT0_Suite = True: lYMNUOPT0_Nb = 0
wNbMax = recYMNUOPT0_Block + 2: ReDim Preserve lYMNUOPT0(wNbMax)

wYMNUOPT0 = lYMNUOPT0(1)
Do Until Not blnYMNUOPT0_Suite
    MsgTxtLen = 0
    Call srvYMNUOPT0_PutBuffer(wYMNUOPT0)
    Call srvYMNUOPT0_PutBuffer(lYMNUOPT0(0))
    If IsNull(SndRcv()) Then
        MsgTxtIndex = 0
        Do While MsgTxtIndex < MsgTxtLen
            If IsNull(srvYMNUOPT0_GetBuffer(wYMNUOPT0)) Then
            
                lYMNUOPT0_Nb = lYMNUOPT0_Nb + 1
                If lYMNUOPT0_Nb > wNbMax Then
                    wNbMax = wNbMax + recYMNUOPT0_Block
                    ReDim Preserve lYMNUOPT0(wNbMax)
                End If
            
                lYMNUOPT0(lYMNUOPT0_Nb) = wYMNUOPT0
                blnYMNUOPT0_Suite = True
            Else
                blnYMNUOPT0_Suite = False
                Exit Do
            End If
        Loop
    End If

    lYMNUOPT0(0).Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Public Function srvYMNUOPT0_Monitor(recYMNUOPT0 As typeYMNUOPT0)
'-----------------------------------------------------

Select Case mId$(Trim(recYMNUOPT0.Method), 1, 4)
    Case "Seek"
                srvYMNUOPT0_Monitor = srvYMNUOPT0_Seek(recYMNUOPT0)
    Case Else
                recYMNUOPT0.Err = recYMNUOPT0.Method
                Call srvYMNUOPT0_Error(recYMNUOPT0)
                srvYMNUOPT0_Monitor = recYMNUOPT0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYMNUOPT0_Error(recYMNUOPT0 As typeYMNUOPT0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YMNUOPT0" & Chr$(10) & Chr$(13)

Select Case mId$(recYMNUOPT0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYMNUOPT0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YMNUOPT0s.bas  ( " _
                & Trim(recYMNUOPT0.obj) & " : " & Trim(recYMNUOPT0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYMNUOPT0_GetBuffer(recYMNUOPT0 As typeYMNUOPT0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYMNUOPT0_GetBuffer = Null
recYMNUOPT0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYMNUOPT0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYMNUOPT0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYMNUOPT0.Err = Space$(10) Then
    recYMNUOPT0.MNUOPTCOD = CLng(Val(mId$(MsgTxt, K + 1, 8)))
    recYMNUOPT0.MNUOPTCLI = mId$(MsgTxt, K + 9, 7)
    recYMNUOPT0.MNUOPTLIB = mId$(MsgTxt, K + 16, 35)
    recYMNUOPT0.MNUOPTENS = mId$(MsgTxt, K + 51, 8)
    recYMNUOPT0.MNUOPTENT = mId$(MsgTxt, K + 59, 8)
    recYMNUOPT0.MNUOPTSTR = mId$(MsgTxt, K + 67, 1)
    recYMNUOPT0.MNUOPTARE = mId$(MsgTxt, K + 68, 1)
    recYMNUOPT0.MNUOPTBAT = mId$(MsgTxt, K + 69, 1)
    recYMNUOPT0.MNUOPTVAL = mId$(MsgTxt, K + 70, 1)
    recYMNUOPT0.MNUOPTSUP = mId$(MsgTxt, K + 71, 1)
    recYMNUOPT0.MNUOPTOIA = mId$(MsgTxt, K + 72, 1)
    recYMNUOPT0.MNUOPTGES = mId$(MsgTxt, K + 73, 1)
Else
    srvYMNUOPT0_GetBuffer = recYMNUOPT0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYMNUOPT0Len

End Function

'---------------------------------------------------------
Private Sub srvYMNUOPT0_PutBuffer(recYMNUOPT0 As typeYMNUOPT0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYMNUOPT0Len) = Space$(recYMNUOPT0Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYMNUOPT0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYMNUOPT0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 8) = Format$(recYMNUOPT0.MNUOPTCOD, "00000000")
    Mid$(MsgTxt, K + 9, 7) = recYMNUOPT0.MNUOPTCLI
    Mid$(MsgTxt, K + 16, 35) = recYMNUOPT0.MNUOPTLIB
    Mid$(MsgTxt, K + 51, 8) = recYMNUOPT0.MNUOPTENS
    Mid$(MsgTxt, K + 59, 8) = recYMNUOPT0.MNUOPTENT
    Mid$(MsgTxt, K + 67, 1) = recYMNUOPT0.MNUOPTSTR
    Mid$(MsgTxt, K + 68, 1) = recYMNUOPT0.MNUOPTARE
    Mid$(MsgTxt, K + 69, 1) = recYMNUOPT0.MNUOPTBAT
    Mid$(MsgTxt, K + 70, 1) = recYMNUOPT0.MNUOPTVAL
    Mid$(MsgTxt, K + 71, 1) = recYMNUOPT0.MNUOPTSUP
    Mid$(MsgTxt, K + 72, 1) = recYMNUOPT0.MNUOPTOIA
    Mid$(MsgTxt, K + 73, 1) = recYMNUOPT0.MNUOPTGES

MsgTxtLen = MsgTxtLen + recYMNUOPT0Len
End Sub
Public Sub srvYMNUOPT0_ElpDisplay(recYMNUOPT0 As typeYMNUOPT0)
frmElpDisplay.fgData.Rows = 13
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUOPTCOD    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPTION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUOPT0.MNUOPTCOD
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUOPTCLI    7A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CLIENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUOPT0.MNUOPTCLI
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUOPTLIB   35A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LIBELLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUOPT0.MNUOPTLIB
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUOPTENS    8A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ENSEMBLE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUOPT0.MNUOPTENS
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUOPTENT    8A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "POINT ENTREE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUOPT0.MNUOPTENT
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUOPTSTR    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "OPTION STRAB"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUOPT0.MNUOPTSTR
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUOPTARE    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ARRET LOGICIEL"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUOPT0.MNUOPTARE
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUOPTBAT    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "OPTION BATCH"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUOPT0.MNUOPTBAT
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUOPTVAL    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "VALID. BATCH"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUOPT0.MNUOPTVAL
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUOPTSUP    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "A SUPPRIMER"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUOPT0.MNUOPTSUP
frmElpDisplay.fgData.Row = 11
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUOPTOIA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTER-AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUOPT0.MNUOPTOIA
frmElpDisplay.fgData.Row = 12
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUOPTGES    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SECUR.INTER-ETAB."
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUOPT0.MNUOPTGES
frmElpDisplay.Show vbModal
End Sub

'---------------------------------------------------------
Private Function srvYMNUOPT0_Seek(recYMNUOPT0 As typeYMNUOPT0)
'---------------------------------------------------------

srvYMNUOPT0_Seek = "?"
MsgTxtLen = 0
Call srvYMNUOPT0_PutBuffer(recYMNUOPT0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYMNUOPT0_GetBuffer(recYMNUOPT0)) Then
            srvYMNUOPT0_Seek = Null
        Else
            Call srvYMNUOPT0_Error(recYMNUOPT0)
        End If
    End If
End If

End Function

'-----------------------------------------------------
Function srvYMNUOPT0_Update(recYMNUOPT0 As typeYMNUOPT0)
'-----------------------------------------------------

srvYMNUOPT0_Update = "?"

MsgTxtLen = 0
Call srvYMNUOPT0_PutBuffer(recYMNUOPT0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYMNUOPT0_GetBuffer(recYMNUOPT0)) Then
        Call srvYMNUOPT0_Error(recYMNUOPT0)
        srvYMNUOPT0_Update = recYMNUOPT0.Err
        Exit Function
    Else
        srvYMNUOPT0_Update = Null
    End If
Else
    recYMNUOPT0.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYMNUOPT0_Init(recYMNUOPT0 As typeYMNUOPT0)
'---------------------------------------------------------
MsgTxt = Space$(recYMNUOPT0Len)
MsgTxtIndex = 0
Call srvYMNUOPT0_GetBuffer(recYMNUOPT0)
recYMNUOPT0.obj = "ZMNUOPT0_S"

End Sub






