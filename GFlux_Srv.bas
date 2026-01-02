Attribute VB_Name = "srvGFlux"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recGFluxLen = 187 ' 34 + 153
Public Const recGFlux_Block = 40

Type typeGFlux
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    IdRéférence             As Long
    FluxSéquence            As Long
    
    Application             As String * 5
    OpérationCode           As String * 5
    
    Devise1                 As String * 3
    Montant1                As Currency
    
    Devise2                 As String * 3
    Montant2                As Currency
    
    Taux                    As Double
    TauxProvisoire          As String * 1
    Nbj                     As Integer
    
   
    AmjEchéanceTrt          As String * 8
    AmjDébut                As String * 8
    AmjFin                  As String * 8
    AmjOpération            As String * 8
    AmjValeur               As String * 8

    Statut                  As String * 1
    StatutPlus              As String * 2
    Flag1                   As String * 1
    Flag2                   As String * 1
    Flag3                   As String * 1
   
    ElpId                   As Long
    ElpUpdate               As Integer
    ElpControl              As String * 10
    
End Type
    
Public arrGFlux() As typeGFlux
Public arrGFlux_Nb As Integer
Public arrGFlux_NbMax As Integer
Public arrGFlux_Index As Integer
Public arrGFlux_Suite As Boolean
'-----------------------------------------------------
Function srvGFlux_Update(recGFlux As typeGFlux)
'-----------------------------------------------------

If blnMsgTxt_Concat_Transaction Then
    Call srvGFlux_PutBuffer(recGFlux)
    srvGFlux_Update = Null
    Exit Function
End If

srvGFlux_Update = "?"

MsgTxtLen = 0
Call srvGFlux_PutBuffer(recGFlux)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvGFlux_GetBuffer(recGFlux)) Then
        Call srvGFlux_Error(recGFlux)
        srvGFlux_Update = recGFlux.Err
        Exit Function
    Else
        srvGFlux_Update = Null
    End If
Else
    recGFlux.Err = "srv"
End If


'=====================================================
End Function



Public Sub srvGFlux_Load(recGFluxMin As typeGFlux, recGFluxMax As typeGFlux)
Dim mMethod As String

mMethod = Trim(recGFluxMin.Method) & "+"
arrGFlux_NbMax = 0
arrGFlux_Suite = True: arrGFlux_Nb = 0
arrGFlux_NbMax = recGFlux_Block: ReDim arrGFlux(arrGFlux_NbMax)

'recGFluxMin.Application = GFlux_paramApplication
'recGFluxMax.Application = GFlux_paramApplication
arrGFlux(0) = recGFluxMax
arrGFlux_Suite = True
Do Until Not arrGFlux_Suite
    srvGFlux_Monitor recGFluxMin
    recGFluxMin = arrGFlux(arrGFlux_Nb)
    recGFluxMin.Method = mMethod
'    recGFluxMin.Application = GFlux_paramApplication
Loop

End Sub

'-----------------------------------------------------
Function srvGFlux_Dtaq_Put(lFct As String, recGFlux As typeGFlux)
'-----------------------------------------------------

srvGFlux_Dtaq_Put = Null
Select Case lFct
    Case "Init": MsgTxtLen = 0
    Case "Add": Call srvGFlux_PutBuffer(recGFlux)
                If MsgTxtLen + recGFluxLen >= recGFlux_Block * recGFluxLen Then
                    Call srvGFlux_Dtaq_Snd(recGFlux): MsgTxtLen = 0
                End If
    Case "Snd": If MsgTxtLen > 0 Then Call srvGFlux_Dtaq_Snd(recGFlux)
    Case Else: srvGFlux_Dtaq_Put = lFct
End Select
'=====================================================
End Function


'-----------------------------------------------------
Function srvGFlux_Dtaq_Snd(recGFlux As typeGFlux)
'-----------------------------------------------------

srvGFlux_Dtaq_Snd = "?"

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvGFlux_GetBuffer(recGFlux)) Then
        Call srvGFlux_Error(recGFlux)
        srvGFlux_Dtaq_Snd = recGFlux.Err
        Exit Function
    Else
        srvGFlux_Dtaq_Snd = Null
    End If
Else
    recGFlux.Err = "Snd"
End If


'=====================================================
End Function



'-----------------------------------------------------
Public Function srvGFlux_Monitor(recGFlux As typeGFlux)
'-----------------------------------------------------
blnFR_Convert = False

arrGFlux_Suite = False
Select Case mId$(Trim(recGFlux.Method), 1, 4)
    Case "Seek"
                srvGFlux_Monitor = srvGFlux_Seek(recGFlux)
    Case "Snap"
              srvGFlux_Monitor = srvGFlux_Snap(recGFlux)
    Case Else
                recGFlux.Err = recGFlux.Method
                Call srvGFlux_Error(recGFlux)
                srvGFlux_Monitor = recGFlux.Err
End Select

End Function

'-----------------------------------------------------
Sub srvGFlux_Error(recGFlux As typeGFlux)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "GFlux" & Chr$(10) & Chr$(13)

Select Case mId$(recGFlux.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recGFlux.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recGFlux.IdRéférence & " : " & recGFlux.FluxSéquence & " : " & recGFlux.OpérationCode _
        , I, "module : GFluxs.bas  ( " & Trim(recGFlux.obj) & " : " & Trim(recGFlux.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvGFlux_GetBuffer(recGFlux As typeGFlux)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvGFlux_GetBuffer = Null
recGFlux.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recGFlux.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recGFlux.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recGFlux.Err = Space$(10) Then
    recGFlux.IdRéférence = CLng(Val(mId$(MsgTxt, K + 1, 12)))
    recGFlux.FluxSéquence = CLng(Val(mId$(MsgTxt, K + 13, 5)))
    
    recGFlux.Application = mId$(MsgTxt, K + 18, 5)
    recGFlux.OpérationCode = mId$(MsgTxt, K + 23, 5)
    recGFlux.Devise1 = mId$(MsgTxt, K + 28, 3)
    recGFlux.Montant1 = CCur(Val(mId$(MsgTxt, K + 31, 17))) / 100
    recGFlux.Devise2 = mId$(MsgTxt, K + 48, 3)
    recGFlux.Montant2 = CCur(Val(mId$(MsgTxt, K + 51, 17))) / 100
    
    recGFlux.Taux = CDbl(Val(mId$(MsgTxt, K + 68, 9))) / 1000000
    recGFlux.TauxProvisoire = mId$(MsgTxt, K + 77, 1)
    recGFlux.Nbj = CInt(Val(mId$(MsgTxt, K + 78, 5)))
    recGFlux.AmjEchéanceTrt = mId$(MsgTxt, K + 83, 8)
    recGFlux.AmjDébut = mId$(MsgTxt, K + 91, 8)
    recGFlux.AmjFin = mId$(MsgTxt, K + 99, 8)
    recGFlux.AmjOpération = mId$(MsgTxt, K + 107, 8)
    recGFlux.AmjValeur = mId$(MsgTxt, K + 115, 8)
    
    recGFlux.Statut = mId$(MsgTxt, K + 123, 1)
    recGFlux.StatutPlus = mId$(MsgTxt, K + 124, 2)
    recGFlux.Flag1 = mId$(MsgTxt, K + 126, 1)
    recGFlux.Flag2 = mId$(MsgTxt, K + 127, 1)
    recGFlux.Flag3 = mId$(MsgTxt, K + 128, 1)
    recGFlux.ElpId = CLng(Val(mId$(MsgTxt, K + 129, 12)))
    recGFlux.ElpUpdate = CInt(Val(mId$(MsgTxt, K + 141, 3)))
    recGFlux.ElpControl = mId$(MsgTxt, K + 144, 10)

Else
    srvGFlux_GetBuffer = recGFlux.Err
End If

MsgTxtIndex = MsgTxtIndex + recGFluxLen

End Function

'---------------------------------------------------------
Private Sub srvGFlux_PutBuffer(recGFlux As typeGFlux)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recGFlux.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recGFlux.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 12) = Format$(recGFlux.IdRéférence, "000000000000")
Mid$(MsgTxt, K + 13, 5) = Format$(recGFlux.FluxSéquence, "00000")
Mid$(MsgTxt, K + 18, 5) = recGFlux.Application
Mid$(MsgTxt, K + 23, 5) = recGFlux.OpérationCode

Mid$(MsgTxt, K + 28, 3) = recGFlux.Devise1
Mid$(MsgTxt, K + 31, 17) = Format$(recGFlux.Montant1 * 100, "00000000000000000")
Mid$(MsgTxt, K + 48, 3) = recGFlux.Devise2
Mid$(MsgTxt, K + 51, 17) = Format$(recGFlux.Montant2 * 100, "00000000000000000")

Mid$(MsgTxt, K + 68, 9) = Format$(recGFlux.Taux * 1000000, "000000000")
Mid$(MsgTxt, K + 77, 1) = recGFlux.TauxProvisoire
Mid$(MsgTxt, K + 78, 5) = Format$(recGFlux.Nbj, "00000")
Mid$(MsgTxt, K + 83, 8) = recGFlux.AmjEchéanceTrt
Mid$(MsgTxt, K + 91, 8) = recGFlux.AmjDébut
Mid$(MsgTxt, K + 99, 8) = recGFlux.AmjFin
Mid$(MsgTxt, K + 107, 8) = recGFlux.AmjOpération
Mid$(MsgTxt, K + 115, 8) = recGFlux.AmjValeur

Mid$(MsgTxt, K + 123, 1) = recGFlux.Statut
Mid$(MsgTxt, K + 124, 2) = recGFlux.StatutPlus
Mid$(MsgTxt, K + 126, 1) = recGFlux.Flag1
Mid$(MsgTxt, K + 127, 1) = recGFlux.Flag2
Mid$(MsgTxt, K + 128, 1) = recGFlux.Flag3
Mid$(MsgTxt, K + 129, 12) = Format$(recGFlux.ElpId, "000000000000")
Mid$(MsgTxt, K + 141, 3) = Format$(recGFlux.ElpUpdate, "000")
Mid$(MsgTxt, K + 144, 10) = recGFlux.ElpControl

MsgTxtLen = MsgTxtLen + recGFluxLen
End Sub



'---------------------------------------------------------
Private Function srvGFlux_Seek(recGFlux As typeGFlux)
'---------------------------------------------------------

srvGFlux_Seek = "?"
MsgTxtLen = 0
Call srvGFlux_PutBuffer(recGFlux)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvGFlux_GetBuffer(recGFlux)) Then
        srvGFlux_Seek = Null
    Else
        Call srvGFlux_Error(recGFlux)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvGFlux_Snap(recGFlux As typeGFlux)
'---------------------------------------------------------
srvGFlux_Snap = "?"
MsgTxtLen = 0
Call srvGFlux_PutBuffer(recGFlux)
Call srvGFlux_PutBuffer(arrGFlux(0))
If IsNull(SndRcv()) Then
    srvGFlux_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvGFlux_GetBuffer(recGFlux)) Then
            Call arrGFlux_AddItem(recGFlux)
            arrGFlux_Suite = True
        Else
            arrGFlux_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recGFlux_Init(recGFlux As typeGFlux)
'---------------------------------------------------------
MsgTxt = Space$(recGFluxLen)
MsgTxtIndex = 0
Call srvGFlux_GetBuffer(recGFlux)
recGFlux.obj = "SRVGFLUX    "

End Sub

'---------------------------------------------------------
Public Sub arrGFlux_AddItem(recGFlux As typeGFlux)
'---------------------------------------------------------
          
arrGFlux_Nb = arrGFlux_Nb + 1
    
If arrGFlux_Nb > arrGFlux_NbMax Then
    arrGFlux_NbMax = arrGFlux_NbMax + 10
    ReDim Preserve arrGFlux(arrGFlux_NbMax)
End If
            
arrGFlux(arrGFlux_Nb) = recGFlux
End Sub

Public Sub srvGFlux_ElpDisplay(recGFlux As typeGFlux)
frmElpDisplay.fgData.Rows = 28
frmElpDisplay.fgData.Row = 1
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "obj"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.obj
frmElpDisplay.fgData.Row = 2
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Method"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.Method
frmElpDisplay.fgData.Row = 3
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Err"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.Err
frmElpDisplay.fgData.Row = 4
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "IdRéférence"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.IdRéférence
frmElpDisplay.fgData.Row = 5
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "FluxSéquence"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.FluxSéquence
frmElpDisplay.fgData.Row = 6
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Application"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.Application
frmElpDisplay.fgData.Row = 7
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "OpérationCode"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.OpérationCode
frmElpDisplay.fgData.Row = 8
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Devise1"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.Devise1
frmElpDisplay.fgData.Row = 9
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Montant1"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.Montant1
frmElpDisplay.fgData.Row = 10
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Devise2"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.Devise2
frmElpDisplay.fgData.Row = 11
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Montant2"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.Montant2
frmElpDisplay.fgData.Row = 12
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Taux"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.Taux
frmElpDisplay.fgData.Row = 13
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "TauxProvisoire"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.TauxProvisoire
frmElpDisplay.fgData.Row = 14
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Nbj "
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.Nbj
frmElpDisplay.fgData.Row = 15
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AmjEchéanceTrt"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.AmjEchéanceTrt
frmElpDisplay.fgData.Row = 16
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AmjDébut"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.AmjDébut
frmElpDisplay.fgData.Row = 17
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AmjFin"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.AmjFin
frmElpDisplay.fgData.Row = 18
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AmjOpération"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.AmjOpération
frmElpDisplay.fgData.Row = 19
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "AmjValeur "
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.AmjValeur
frmElpDisplay.fgData.Row = 20
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Statut"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.Statut
frmElpDisplay.fgData.Row = 21
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "StatutPlus"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.StatutPlus
frmElpDisplay.fgData.Row = 22
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Flag1"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.Flag1
frmElpDisplay.fgData.Row = 23
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Flag2"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.Flag2
frmElpDisplay.fgData.Row = 24
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Flag3"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.Flag3
frmElpDisplay.fgData.Row = 25
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ElpId"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.ElpId
frmElpDisplay.fgData.Row = 26
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ElpUpdate"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.ElpUpdate
frmElpDisplay.fgData.Row = 27
    frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "ElpControl"
    frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = recGFlux.ElpControl

frmElpDisplay.Show vbModal

End Sub
