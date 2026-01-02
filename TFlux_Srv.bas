Attribute VB_Name = "srvTFlux"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recTFluxLen = 211 ' 34 + 177
Public Const recTFlux_Block = 35

Type typeTFlux
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    IdRéférence             As Long
    IdSéquence              As Integer
    CodeOpération           As String * 4
    
    Capital                 As Currency
    Intérêts                As Currency
    Taux                    As Double
    TauxProvisoire          As String * 1
    Nbj                     As Integer
    
   
    AmjEchéanceTrt          As String * 8
    AmjDébut                As String * 8
    AmjFin                  As String * 8
    AmjOpération            As String * 8
    AmjValeur               As String * 8

    CptMvtUsr               As String * 10
    CptMvtAMJ               As String * 8
    CptMvtHMS               As String * 6
    CptMvtLot               As Long
    CptMvtPièce             As Long
    CptMvtLigne             As Long

    Statut                  As String * 1
    StatutPlus              As String * 2
   
    ElpId                   As Long
    ElpUpdate               As Integer
    ElpControl              As String * 10
    
End Type
    
Public arrTFlux() As typeTFlux
Public arrTFlux_Nb As Integer
Public arrTFlux_NbMax As Integer
Public arrTFlux_Index As Integer
Public arrTFlux_Suite As Boolean
'-----------------------------------------------------
Function srvTFlux_Update(recTFlux As typeTFlux)
'-----------------------------------------------------

srvTFlux_Update = "?"

MsgTxtLen = 0
Call srvTFlux_PutBuffer(recTFlux)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvTFlux_GetBuffer(recTFlux)) Then
        Call srvTFlux_Error(recTFlux)
        srvTFlux_Update = recTFlux.Err
        Exit Function
    Else
        srvTFlux_Update = Null
    End If
Else
    recTFlux.Err = "srv"
End If


'=====================================================
End Function



Public Sub srvTFlux_Load(recTfluxMin As typeTFlux, recTfluxMax As typeTFlux)
Dim mMethod As String

mMethod = Trim(recTfluxMin.Method) & "+"
arrTFlux_NbMax = 0
arrTFlux_Suite = True: arrTFlux_Nb = 0
arrTFlux_NbMax = 35: ReDim arrTFlux(arrTFlux_NbMax)

recTfluxMin.CptMvtUsr = paramTFlux_Service
recTfluxMax.CptMvtUsr = paramTFlux_Service
arrTFlux(0) = recTfluxMax
arrTFlux_Suite = True
Do Until Not arrTFlux_Suite
    srvTFlux_Monitor recTfluxMin
    recTfluxMin = arrTFlux(arrTFlux_Nb)
    recTfluxMin.Method = mMethod
    recTfluxMin.CptMvtUsr = paramTFlux_Service
Loop

End Sub

'-----------------------------------------------------
Function srvTFlux_Dtaq_Put(lFct As String, recTFlux As typeTFlux)
'-----------------------------------------------------

srvTFlux_Dtaq_Put = Null
Select Case lFct
    Case "Init": MsgTxtLen = 0
    Case "Add": Call srvTFlux_PutBuffer(recTFlux)
                If MsgTxtLen + recTFluxLen >= recTFlux_Block * recTFluxLen Then
                    Call srvTFlux_Dtaq_Snd(recTFlux): MsgTxtLen = 0
                End If
    Case "Snd": If MsgTxtLen > 0 Then Call srvTFlux_Dtaq_Snd(recTFlux)
    Case Else: srvTFlux_Dtaq_Put = lFct
End Select
'=====================================================
End Function


'-----------------------------------------------------
Function srvTFlux_Dtaq_Snd(recTFlux As typeTFlux)
'-----------------------------------------------------

srvTFlux_Dtaq_Snd = "?"

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvTFlux_GetBuffer(recTFlux)) Then
        Call srvTFlux_Error(recTFlux)
        srvTFlux_Dtaq_Snd = recTFlux.Err
        Exit Function
    Else
        srvTFlux_Dtaq_Snd = Null
    End If
Else
    recTFlux.Err = "Snd"
End If


'=====================================================
End Function



'-----------------------------------------------------
Public Function srvTFlux_Monitor(recTFlux As typeTFlux)
'-----------------------------------------------------

arrTFlux_Suite = False
Select Case mId$(Trim(recTFlux.Method), 1, 4)
    Case "Seek"
                srvTFlux_Monitor = srvTFlux_Seek(recTFlux)
    Case "Snap"
              srvTFlux_Monitor = srvTFlux_Snap(recTFlux)
    Case Else
                recTFlux.Err = recTFlux.Method
                Call srvTFlux_Error(recTFlux)
                srvTFlux_Monitor = recTFlux.Err
End Select

End Function

'-----------------------------------------------------
Sub srvTFlux_Error(recTFlux As typeTFlux)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "TFlux" & Chr$(10) & Chr$(13)

Select Case mId$(recTFlux.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recTFlux.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recTFlux.IdRéférence & " : " & recTFlux.IdSéquence & " : " & recTFlux.CodeOpération _
        , I, "module : TFluxs.bas  ( " & Trim(recTFlux.obj) & " : " & Trim(recTFlux.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvTFlux_GetBuffer(recTFlux As typeTFlux)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvTFlux_GetBuffer = Null
recTFlux.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recTFlux.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recTFlux.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recTFlux.Err = Space$(10) Then
    recTFlux.IdRéférence = CLng(Val(mId$(MsgTxt, K + 1, 10)))
    recTFlux.IdSéquence = CInt(Val(mId$(MsgTxt, K + 11, 3)))
    recTFlux.CodeOpération = mId$(MsgTxt, K + 14, 4)
    recTFlux.Capital = CCur(Val(mId$(MsgTxt, K + 18, 17))) / 100
    recTFlux.Intérêts = CCur(Val(mId$(MsgTxt, K + 35, 17))) / 100
    recTFlux.Taux = CDbl(Val(mId$(MsgTxt, K + 52, 9))) / 1000000
    recTFlux.TauxProvisoire = mId$(MsgTxt, K + 61, 1)
    recTFlux.Nbj = CInt(Val(mId$(MsgTxt, K + 62, 5)))
    recTFlux.AmjEchéanceTrt = mId$(MsgTxt, K + 67, 8)
    recTFlux.AmjDébut = mId$(MsgTxt, K + 75, 8)
    recTFlux.AmjFin = mId$(MsgTxt, K + 83, 8)
    recTFlux.AmjOpération = mId$(MsgTxt, K + 91, 8)
    recTFlux.AmjValeur = mId$(MsgTxt, K + 99, 8)
    recTFlux.CptMvtUsr = mId$(MsgTxt, K + 107, 10)
    recTFlux.CptMvtAMJ = Format$(Val(mId$(MsgTxt, K + 117, 8)), "00000000")
    recTFlux.CptMvtHMS = Format$(Val(mId$(MsgTxt, K + 125, 6)), "000000")
    recTFlux.CptMvtLot = CLng(Val(mId$(MsgTxt, K + 131, 7)))
    recTFlux.CptMvtPièce = CLng(Val(mId$(MsgTxt, K + 138, 7)))
    recTFlux.CptMvtLigne = CLng(Val(mId$(MsgTxt, K + 145, 5)))
    recTFlux.Statut = mId$(MsgTxt, K + 150, 1)
    recTFlux.StatutPlus = mId$(MsgTxt, K + 151, 2)
    recTFlux.ElpId = CLng(Val(mId$(MsgTxt, K + 153, 12)))
    recTFlux.ElpUpdate = CInt(Val(mId$(MsgTxt, K + 165, 3)))
    recTFlux.ElpControl = mId$(MsgTxt, K + 168, 10)

Else
    srvTFlux_GetBuffer = recTFlux.Err
End If

MsgTxtIndex = MsgTxtIndex + recTFluxLen

End Function

'---------------------------------------------------------
Private Sub srvTFlux_PutBuffer(recTFlux As typeTFlux)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recTFlux.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recTFlux.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 10) = Format$(recTFlux.IdRéférence, "0000000000")
Mid$(MsgTxt, K + 11, 3) = Format$(recTFlux.IdSéquence, "000")
Mid$(MsgTxt, K + 14, 4) = recTFlux.CodeOpération
Mid$(MsgTxt, K + 18, 17) = Format$(recTFlux.Capital * 100, "00000000000000000")
Mid$(MsgTxt, K + 35, 17) = Format$(recTFlux.Intérêts * 100, "00000000000000000")
Mid$(MsgTxt, K + 52, 9) = Format$(recTFlux.Taux * 1000000, "000000000")
Mid$(MsgTxt, K + 61, 1) = recTFlux.TauxProvisoire
Mid$(MsgTxt, K + 62, 5) = Format$(recTFlux.Nbj, "00000")
Mid$(MsgTxt, K + 67, 8) = recTFlux.AmjEchéanceTrt
Mid$(MsgTxt, K + 75, 8) = recTFlux.AmjDébut
Mid$(MsgTxt, K + 83, 8) = recTFlux.AmjFin
Mid$(MsgTxt, K + 91, 8) = recTFlux.AmjOpération
Mid$(MsgTxt, K + 99, 8) = recTFlux.AmjValeur
Mid$(MsgTxt, K + 107, 10) = recTFlux.CptMvtUsr
Mid$(MsgTxt, K + 117, 8) = Format$(recTFlux.CptMvtAMJ, "00000000")
Mid$(MsgTxt, K + 125, 6) = Format$(recTFlux.CptMvtHMS, "000000")
Mid$(MsgTxt, K + 131, 7) = Format$(recTFlux.CptMvtLot, "0000000")
Mid$(MsgTxt, K + 138, 7) = Format$(recTFlux.CptMvtPièce, "0000000")
Mid$(MsgTxt, K + 145, 5) = Format$(recTFlux.CptMvtLigne, "00000")
Mid$(MsgTxt, K + 150, 1) = recTFlux.Statut
Mid$(MsgTxt, K + 151, 2) = recTFlux.StatutPlus
Mid$(MsgTxt, K + 153, 12) = Format$(recTFlux.ElpId, "000000000000")
Mid$(MsgTxt, K + 165, 3) = Format$(recTFlux.ElpUpdate, "000")
Mid$(MsgTxt, K + 168, 10) = recTFlux.ElpControl

MsgTxtLen = MsgTxtLen + recTFluxLen
End Sub



'---------------------------------------------------------
Private Function srvTFlux_Seek(recTFlux As typeTFlux)
'---------------------------------------------------------

srvTFlux_Seek = "?"
MsgTxtLen = 0
Call srvTFlux_PutBuffer(recTFlux)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvTFlux_GetBuffer(recTFlux)) Then
        srvTFlux_Seek = Null
    Else
        Call srvTFlux_Error(recTFlux)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvTFlux_Snap(recTFlux As typeTFlux)
'---------------------------------------------------------
srvTFlux_Snap = "?"
MsgTxtLen = 0
Call srvTFlux_PutBuffer(recTFlux)
Call srvTFlux_PutBuffer(arrTFlux(0))
If IsNull(SndRcv()) Then
    srvTFlux_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvTFlux_GetBuffer(recTFlux)) Then
            Call arrTFlux_AddItem(recTFlux)
            arrTFlux_Suite = True
        Else
            arrTFlux_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recTFlux_Init(recTFlux As typeTFlux)
'---------------------------------------------------------
MsgTxt = Space$(recTFluxLen)
MsgTxtIndex = 0
Call srvTFlux_GetBuffer(recTFlux)
recTFlux.obj = "SRVTFLUX    "

End Sub

'---------------------------------------------------------
Public Sub arrTFlux_AddItem(recTFlux As typeTFlux)
'---------------------------------------------------------
          
arrTFlux_Nb = arrTFlux_Nb + 1
    
If arrTFlux_Nb > arrTFlux_NbMax Then
    arrTFlux_NbMax = arrTFlux_NbMax + 10
    ReDim Preserve arrTFlux(arrTFlux_NbMax)
End If
            
arrTFlux(arrTFlux_Nb) = recTFlux
End Sub
