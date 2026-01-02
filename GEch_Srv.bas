Attribute VB_Name = "srvGEch"

'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recGEchLen = 160 ' 34 + 126
Public Const recGEch_Block = 40

Type typeGEch
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    IdRéférence             As Long
    EchSéquence             As Long
    
    Application             As String * 5
    FluxSéquence            As Long
     
    EchFct                  As String * 10
    EchAMJ                  As String * 8
    EchHMS                  As String * 6
    EchUsr                  As String * 10
    
    ActionFct               As String * 10
    ActionAmj               As String * 8
    ActionHms               As String * 6
    ActionUsr               As String * 10
   
    Statut                  As String * 1
    StatutPlus              As String * 2
    Flag1                   As String * 1
    Flag2                   As String * 1
    Flag3                   As String * 1
    
    ElpId                   As Long
    ElpUpdate               As Integer
    ElpControl              As String * 10
    
End Type
    
Public arrGECh() As typeGEch
Public arrGECh_Nb As Integer
Public arrGECh_NbMax As Integer
Public arrGECh_Index As Integer
Public arrGEch_Suite As Boolean
Public Sub srvGEch_ElpDisplay(recGEch As typeGEch)
frmElpDisplay.fgData.Rows = 24
frmElpDisplay.fgData.Row = 1
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "obj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.obj
frmElpDisplay.fgData.Row = 2
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Method"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.Method
frmElpDisplay.fgData.Row = 3
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Err"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.Err
frmElpDisplay.fgData.Row = 4
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "IdRéférence"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.IdRéférence
frmElpDisplay.fgData.Row = 5
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EchSéquence"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.EchSéquence
frmElpDisplay.fgData.Row = 6
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Application"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.Application
frmElpDisplay.fgData.Row = 7
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "FluxSéquence"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.FluxSéquence
frmElpDisplay.fgData.Row = 8
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EchFct"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.EchFct
frmElpDisplay.fgData.Row = 9
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EchAMJ"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.EchAMJ
frmElpDisplay.fgData.Row = 10
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EchHMS"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.EchHMS
frmElpDisplay.fgData.Row = 11
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "EchUsr"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.EchUsr
frmElpDisplay.fgData.Row = 12
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ActionFct"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.ActionFct
frmElpDisplay.fgData.Row = 13
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ActionAMJ"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.ActionAmj
frmElpDisplay.fgData.Row = 14
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ActionHMS"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.ActionHms
frmElpDisplay.fgData.Row = 15
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ActionUsr"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.ActionUsr
frmElpDisplay.fgData.Row = 16
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Statut"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.Statut
frmElpDisplay.fgData.Row = 17
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "StatutPlus"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.StatutPlus
frmElpDisplay.fgData.Row = 18
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Flag1"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.Flag1
frmElpDisplay.fgData.Row = 19
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Flag2"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.Flag2
frmElpDisplay.fgData.Row = 20
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Flag3"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.Flag3
frmElpDisplay.fgData.Row = 21
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ElpId"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.ElpId
frmElpDisplay.fgData.Row = 22
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ElpUpdate"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.ElpUpdate
frmElpDisplay.fgData.Row = 23
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ElpControl"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGEch.ElpControl

frmElpDisplay.Show vbModal

End Sub


'-----------------------------------------------------
Function srvGEch_Update(recGEch As typeGEch)
'-----------------------------------------------------

If blnMsgTxt_Concat_Transaction Then
    Call srvGEch_PutBuffer(recGEch)
    srvGEch_Update = Null
    Exit Function
End If

srvGEch_Update = "?"

MsgTxtLen = 0
Call srvGEch_PutBuffer(recGEch)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvGEch_GetBuffer(recGEch)) Then
        Call srvGEch_Error(recGEch)
        srvGEch_Update = recGEch.Err
        Exit Function
    Else
        srvGEch_Update = Null
    End If
Else
    recGEch.Err = "srv"
End If


'=====================================================
End Function



Public Sub srvGEch_Load(recGEchMin As typeGEch, recGEchMax As typeGEch)
Dim mMethod As String

mMethod = Trim(recGEchMin.Method) & "+"
arrGECh_NbMax = 0
arrGEch_Suite = True: arrGECh_Nb = 0
arrGECh_NbMax = recGEch_Block: ReDim arrGECh(arrGECh_NbMax)

arrGECh(0) = recGEchMax
arrGEch_Suite = True
Do Until Not arrGEch_Suite
    srvGEch_Monitor recGEchMin
    recGEchMin = arrGECh(arrGECh_Nb)
    recGEchMin.Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Function srvGEch_Dtaq_Put(lFct As String, recGEch As typeGEch)
'-----------------------------------------------------

srvGEch_Dtaq_Put = Null
Select Case lFct
    Case "Init": MsgTxtLen = 0
    Case "Add": Call srvGEch_PutBuffer(recGEch)
                If MsgTxtLen + recGEchLen >= recGEch_Block * recGEchLen Then
                    Call srvGEch_Dtaq_Snd(recGEch): MsgTxtLen = 0
                End If
    Case "Snd": If MsgTxtLen > 0 Then Call srvGEch_Dtaq_Snd(recGEch)
    Case Else: srvGEch_Dtaq_Put = lFct
End Select
'=====================================================
End Function


'-----------------------------------------------------
Function srvGEch_Dtaq_Snd(recGEch As typeGEch)
'-----------------------------------------------------

srvGEch_Dtaq_Snd = "?"

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvGEch_GetBuffer(recGEch)) Then
        Call srvGEch_Error(recGEch)
        srvGEch_Dtaq_Snd = recGEch.Err
        Exit Function
    Else
        srvGEch_Dtaq_Snd = Null
    End If
Else
    recGEch.Err = "Snd"
End If


'=====================================================
End Function



'-----------------------------------------------------
Public Function srvGEch_Monitor(recGEch As typeGEch)
'-----------------------------------------------------
blnFR_Convert = False

arrGEch_Suite = False
Select Case mId$(Trim(recGEch.Method), 1, 4)
    Case "Seek"
                srvGEch_Monitor = srvGEch_Seek(recGEch)
    Case "Snap"
              srvGEch_Monitor = srvGEch_Snap(recGEch)
    Case Else
                recGEch.Err = recGEch.Method
                Call srvGEch_Error(recGEch)
                srvGEch_Monitor = recGEch.Err
End Select

End Function

'-----------------------------------------------------
Sub srvGEch_Error(recGEch As typeGEch)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "GEch" & Chr$(10) & Chr$(13)

Select Case mId$(recGEch.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recGEch.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recGEch.IdRéférence & " : " & recGEch.EchSéquence & " : " & recGEch.EchFct _
        , I, "module : GEchs.bas  ( " & Trim(recGEch.obj) & " : " & Trim(recGEch.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvGEch_GetBuffer(recGEch As typeGEch)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvGEch_GetBuffer = Null
recGEch.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recGEch.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recGEch.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recGEch.Err = Space$(10) Then
    recGEch.IdRéférence = CLng(Val(mId$(MsgTxt, K + 1, 12)))
    recGEch.EchSéquence = CLng(Val(mId$(MsgTxt, K + 13, 5)))
    
    recGEch.Application = mId$(MsgTxt, K + 18, 5)
    recGEch.FluxSéquence = CLng(Val(mId$(MsgTxt, K + 23, 5)))
    recGEch.EchFct = mId$(MsgTxt, K + 28, 10)
    recGEch.EchAMJ = mId$(MsgTxt, K + 38, 8)
    recGEch.EchHMS = mId$(MsgTxt, K + 46, 6)
    recGEch.EchUsr = mId$(MsgTxt, K + 52, 10)
    
    recGEch.ActionFct = mId$(MsgTxt, K + 62, 10)
    recGEch.ActionAmj = mId$(MsgTxt, K + 72, 8)
    recGEch.ActionHms = mId$(MsgTxt, K + 80, 6)
    recGEch.ActionUsr = mId$(MsgTxt, K + 86, 10)
    
    recGEch.Statut = mId$(MsgTxt, K + 96, 1)
    recGEch.StatutPlus = mId$(MsgTxt, K + 97, 2)
    recGEch.Flag1 = mId$(MsgTxt, K + 99, 1)
    recGEch.Flag2 = mId$(MsgTxt, K + 100, 1)
    recGEch.Flag3 = mId$(MsgTxt, K + 101, 1)
    recGEch.ElpId = CLng(Val(mId$(MsgTxt, K + 102, 12)))
    recGEch.ElpUpdate = CInt(Val(mId$(MsgTxt, K + 114, 3)))
    recGEch.ElpControl = mId$(MsgTxt, K + 117, 10)

Else
    srvGEch_GetBuffer = recGEch.Err
End If

MsgTxtIndex = MsgTxtIndex + recGEchLen

End Function

'---------------------------------------------------------
Private Sub srvGEch_PutBuffer(recGEch As typeGEch)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recGEch.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recGEch.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 12) = Format$(recGEch.IdRéférence, "000000000000")
Mid$(MsgTxt, K + 13, 5) = Format$(recGEch.EchSéquence, "00000")

Mid$(MsgTxt, K + 18, 5) = recGEch.Application
Mid$(MsgTxt, K + 23, 5) = Format$(recGEch.FluxSéquence, "00000")

Mid$(MsgTxt, K + 28, 10) = recGEch.EchFct
Mid$(MsgTxt, K + 38, 8) = recGEch.EchAMJ
Mid$(MsgTxt, K + 46, 6) = recGEch.EchHMS
Mid$(MsgTxt, K + 52, 10) = recGEch.EchUsr

Mid$(MsgTxt, K + 62, 10) = recGEch.ActionFct
Mid$(MsgTxt, K + 72, 8) = recGEch.ActionAmj
Mid$(MsgTxt, K + 80, 6) = recGEch.ActionHms
Mid$(MsgTxt, K + 86, 10) = recGEch.ActionUsr

Mid$(MsgTxt, K + 96, 1) = recGEch.Statut
Mid$(MsgTxt, K + 97, 2) = recGEch.StatutPlus
Mid$(MsgTxt, K + 99, 1) = recGEch.Flag1
Mid$(MsgTxt, K + 100, 1) = recGEch.Flag2
Mid$(MsgTxt, K + 101, 1) = recGEch.Flag3
Mid$(MsgTxt, K + 102, 12) = Format$(recGEch.ElpId, "000000000000")
Mid$(MsgTxt, K + 114, 3) = Format$(recGEch.ElpUpdate, "000")
Mid$(MsgTxt, K + 117, 10) = recGEch.ElpControl

MsgTxtLen = MsgTxtLen + recGEchLen
End Sub



'---------------------------------------------------------
Private Function srvGEch_Seek(recGEch As typeGEch)
'---------------------------------------------------------

srvGEch_Seek = "?"
MsgTxtLen = 0
Call srvGEch_PutBuffer(recGEch)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvGEch_GetBuffer(recGEch)) Then
        srvGEch_Seek = Null
    Else
        Call srvGEch_Error(recGEch)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvGEch_Snap(recGEch As typeGEch)
'---------------------------------------------------------
srvGEch_Snap = "?"
MsgTxtLen = 0
Call srvGEch_PutBuffer(recGEch)
Call srvGEch_PutBuffer(arrGECh(0))
If IsNull(SndRcv()) Then
    srvGEch_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvGEch_GetBuffer(recGEch)) Then
            Call arrGEch_AddItem(recGEch)
            arrGEch_Suite = True
        Else
            arrGEch_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recGEch_Init(recGEch As typeGEch)
'---------------------------------------------------------
MsgTxt = Space$(recGEchLen)
MsgTxtIndex = 0
Call srvGEch_GetBuffer(recGEch)
recGEch.obj = "SRVGECH    "
recGEch.EchAMJ = "00000000"
recGEch.EchHMS = "000000"
recGEch.ActionAmj = "00000000"
recGEch.ActionHms = "000000"

End Sub

'---------------------------------------------------------
Public Sub arrGEch_AddItem(recGEch As typeGEch)
'---------------------------------------------------------
          
arrGECh_Nb = arrGECh_Nb + 1
    
If arrGECh_Nb > arrGECh_NbMax Then
    arrGECh_NbMax = arrGECh_NbMax + recGEch_Block
    ReDim Preserve arrGECh(arrGECh_NbMax)
End If
            
arrGECh(arrGECh_Nb) = recGEch
End Sub


