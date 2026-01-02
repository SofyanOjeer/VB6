Attribute VB_Name = "srvGAdresse"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recGAdresseLen = 231 ' 34 + 197

Type typeGAdresse
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    IdRéférence             As String * 19
    L0                      As String * 38
    L2                      As String * 38
    L3                      As String * 38
    L4                      As String * 32
    CodePostal              As String * 5
    Pays                    As String * 2
     
    ElpId                   As Long
    ElpUpdate               As Integer
    ElpControl              As String * 10
    
End Type
    
Public arrGAdresse() As typeGAdresse
Public arrGAdresse_NB As Integer
Public arrGAdresse_NBMax As Integer
Public arrGAdresse_Index As Integer
Public arrGAdresse_Suite As Boolean
Public Sub srvGadresse_Load(recGadresseMin As typeGAdresse, recGadresseMax As typeGAdresse)
Dim mMethod As String

mMethod = Trim(recGadresseMin.Method) & "+"
arrGAdresse_NBMax = 0
arrGAdresse_Suite = True: arrGAdresse_NB = 0
arrGAdresse_NBMax = 10: ReDim arrGAdresse(arrGAdresse_NBMax)

arrGAdresse(0) = recGadresseMax
arrGAdresse_Suite = True
Do Until Not arrGAdresse_Suite
    srvGAdresse_Monitor recGadresseMin
    recGadresseMin = arrGAdresse(arrGAdresse_NB)
    recGadresseMin.Method = mMethod
Loop

End Sub



'-----------------------------------------------------
Public Function srvGAdresse_Monitor(recGAdresse As typeGAdresse)
'-----------------------------------------------------

arrGAdresse_Suite = False
Select Case mId$(Trim(recGAdresse.Method), 1, 4)
    Case "Seek"
                srvGAdresse_Monitor = srvGAdresse_Seek(recGAdresse)
    Case "Snap"
              srvGAdresse_Monitor = srvGAdresse_Snap(recGAdresse)
    Case Else
                recGAdresse.Err = recGAdresse.Method
                Call srvGAdresse_Error(recGAdresse)
                srvGAdresse_Monitor = recGAdresse.Err
End Select

End Function

'-----------------------------------------------------
Sub srvGAdresse_Error(recGAdresse As typeGAdresse)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "GAdresse" & Chr$(10) & Chr$(13)

Select Case mId$(recGAdresse.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recGAdresse.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : GAdresses.bas  ( " _
                & Trim(recGAdresse.obj) & " : " & Trim(recGAdresse.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvGAdresse_GetBuffer(recGAdresse As typeGAdresse)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvGAdresse_GetBuffer = Null
recGAdresse.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recGAdresse.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recGAdresse.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recGAdresse.Err = Space$(10) Then
    recGAdresse.IdRéférence = mId$(MsgTxt, K + 1, 19)
    recGAdresse.L0 = mId$(MsgTxt, K + 20, 38)
    recGAdresse.L2 = mId$(MsgTxt, K + 58, 38)
    recGAdresse.L3 = mId$(MsgTxt, K + 96, 38)
    recGAdresse.L4 = mId$(MsgTxt, K + 134, 32)
    recGAdresse.CodePostal = mId$(MsgTxt, K + 166, 5)
    recGAdresse.Pays = mId$(MsgTxt, K + 171, 2)
    
    recGAdresse.ElpId = CLng(Val(mId$(MsgTxt, K + 173, 12)))
    recGAdresse.ElpUpdate = CInt(Val(mId$(MsgTxt, K + 185, 3)))
    recGAdresse.ElpControl = mId$(MsgTxt, K + 188, 10)

Else
    srvGAdresse_GetBuffer = recGAdresse.Err
End If

MsgTxtIndex = MsgTxtIndex + recGAdresseLen

End Function

'---------------------------------------------------------
Private Sub srvGAdresse_PutBuffer(recGAdresse As typeGAdresse)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recGAdresse.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recGAdresse.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 19) = recGAdresse.IdRéférence
Mid$(MsgTxt, K + 20, 38) = recGAdresse.L0
Mid$(MsgTxt, K + 58, 38) = recGAdresse.L2
Mid$(MsgTxt, K + 96, 38) = recGAdresse.L3
Mid$(MsgTxt, K + 134, 32) = recGAdresse.L4
Mid$(MsgTxt, K + 166, 5) = recGAdresse.CodePostal
Mid$(MsgTxt, K + 171, 2) = recGAdresse.Pays

Mid$(MsgTxt, K + 173, 12) = Format$(recGAdresse.ElpId, "000000000000")
Mid$(MsgTxt, K + 185, 3) = Format$(recGAdresse.ElpUpdate, "000")
Mid$(MsgTxt, K + 188, 10) = recGAdresse.ElpControl

MsgTxtLen = MsgTxtLen + recGAdresseLen
End Sub



'---------------------------------------------------------
Private Function srvGAdresse_Seek(recGAdresse As typeGAdresse)
'---------------------------------------------------------

srvGAdresse_Seek = "?"
MsgTxtLen = 0
Call srvGAdresse_PutBuffer(recGAdresse)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvGAdresse_GetBuffer(recGAdresse)) Then
            srvGAdresse_Seek = Null
        Else
            Call srvGAdresse_Error(recGAdresse)
        End If
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvGAdresse_Snap(recGAdresse As typeGAdresse)
'---------------------------------------------------------
srvGAdresse_Snap = "?"
MsgTxtLen = 0
Call srvGAdresse_PutBuffer(recGAdresse)
Call srvGAdresse_PutBuffer(arrGAdresse(0))
If IsNull(SndRcv()) Then
    srvGAdresse_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvGAdresse_GetBuffer(recGAdresse)) Then
            Call arrGAdresse_AddItem(recGAdresse)
            arrGAdresse_Suite = True
        Else
            arrGAdresse_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'-----------------------------------------------------
Function srvGAdresse_Update(recGAdresse As typeGAdresse)
'-----------------------------------------------------

srvGAdresse_Update = "?"

MsgTxtLen = 0
Call srvGAdresse_PutBuffer(recGAdresse)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvGAdresse_GetBuffer(recGAdresse)) Then
        Call srvGAdresse_Error(recGAdresse)
        srvGAdresse_Update = recGAdresse.Err
        Exit Function
    Else
        srvGAdresse_Update = Null
    End If
Else
    recGAdresse.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recGAdresse_Init(recGAdresse As typeGAdresse)
'---------------------------------------------------------
MsgTxt = Space$(recGAdresseLen)
MsgTxtIndex = 0
Call srvGAdresse_GetBuffer(recGAdresse)
recGAdresse.obj = "SRVGADR"
End Sub

'---------------------------------------------------------
Public Sub arrGAdresse_AddItem(recGAdresse As typeGAdresse)
'---------------------------------------------------------
          
arrGAdresse_NB = arrGAdresse_NB + 1
    
If arrGAdresse_NB > arrGAdresse_NBMax Then
    arrGAdresse_NBMax = arrGAdresse_NBMax + 10
    ReDim Preserve arrGAdresse(arrGAdresse_NBMax)
End If
            
arrGAdresse(arrGAdresse_NB) = recGAdresse
End Sub



Public Function fctGAdresse_Compare(recGAdresse As typeGAdresse, mGAdresse As typeGAdresse)
fctGAdresse_Compare = Null
If recGAdresse.IdRéférence <> mGAdresse.IdRéférence Then fctGAdresse_Compare = "Service": Exit Function
If recGAdresse.L0 <> mGAdresse.L0 Then fctGAdresse_Compare = "L0": Exit Function
If recGAdresse.L2 <> mGAdresse.L2 Then fctGAdresse_Compare = "L2": Exit Function
If recGAdresse.L3 <> mGAdresse.L3 Then fctGAdresse_Compare = "L3": Exit Function
If recGAdresse.L4 <> mGAdresse.L4 Then fctGAdresse_Compare = "L4": Exit Function
If recGAdresse.CodePostal <> mGAdresse.CodePostal Then fctGAdresse_Compare = "TEG": Exit Function
If recGAdresse.Pays <> mGAdresse.Pays Then fctGAdresse_Compare = "TauxActuariel": Exit Function

End Function
