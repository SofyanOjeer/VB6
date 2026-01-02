Attribute VB_Name = "srvGEntité"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recGEntitéLen = 162 ' 34 + 128

Type typeGEntité
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Compte                  As String * 11
    Devise                  As String * 3
    Nature                  As String * 3
    Séquence                As Integer
    ClasseInfo              As String * 3
    AdresseId               As String * 19
    AdresseL0K              As String * 1
    AdresseL1               As String * 32
    AdresseRoutage          As String * 1
    AdresseadresseAmjDébut         As String * 8
    AdresseadresseAmjfin           As String * 8
     
    ElpId                   As Long
    ElpUpdate               As Integer
    ElpControl              As String * 10
    
End Type
    
Public arrGEntité() As typeGEntité
Public arrGEntité_NB As Integer
Public arrGEntité_NBMax As Integer
Public arrGEntité_Index As Integer
Public arrGEntité_Suite As Boolean
'-----------------------------------------------------
Public Function srvGEntité_Monitor(recGEntité As typeGEntité)
'-----------------------------------------------------

arrGEntité_Suite = False
Select Case mId$(Trim(recGEntité.Method), 1, 4)
    Case "Seek"
                srvGEntité_Monitor = srvGEntité_Seek(recGEntité)
    Case "Snap"
              srvGEntité_Monitor = srvGEntité_Snap(recGEntité)
    Case Else
                recGEntité.Err = recGEntité.Method
                Call srvGEntité_Error(recGEntité)
                srvGEntité_Monitor = recGEntité.Err
End Select

End Function

'-----------------------------------------------------
Sub srvGEntité_Error(recGEntité As typeGEntité)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "GEntité" & Chr$(10) & Chr$(13)

Select Case mId$(recGEntité.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recGEntité.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : GEntités.bas  ( " _
                & Trim(recGEntité.obj) & " : " & Trim(recGEntité.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvGEntité_GetBuffer(recGEntité As typeGEntité)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvGEntité_GetBuffer = Null
recGEntité.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recGEntité.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recGEntité.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recGEntité.Err = Space$(10) Then
    recGEntité.Compte = mId$(MsgTxt, K + 1, 11)
    recGEntité.Devise = mId$(MsgTxt, K + 12, 3)
    recGEntité.Nature = mId$(MsgTxt, K + 15, 3)
    recGEntité.Séquence = CInt(Val(mId$(MsgTxt, K + 18, 2)))
    recGEntité.ClasseInfo = mId$(MsgTxt, K + 20, 3)
    recGEntité.AdresseId = mId$(MsgTxt, K + 23, 19)
    recGEntité.AdresseL0K = mId$(MsgTxt, K + 42, 1)
    recGEntité.AdresseL1 = mId$(MsgTxt, K + 43, 32)
    recGEntité.AdresseRoutage = mId$(MsgTxt, K + 75, 1)
    recGEntité.AdresseAMJDébut = mId$(MsgTxt, K + 76, 8)
    recGEntité.AdresseAMJFin = mId$(MsgTxt, K + 84, 8)
'filler
    recGEntité.ElpId = CLng(Val(mId$(MsgTxt, K + 300, 12)))
    recGEntité.ElpUpdate = CInt(Val(mId$(MsgTxt, K + 312, 3)))
    recGEntité.ElpControl = mId$(MsgTxt, K + 315, 10)

Else
    srvGEntité_GetBuffer = recGEntité.Err
End If

MsgTxtIndex = MsgTxtIndex + recGEntitéLen

End Function

'---------------------------------------------------------
Private Sub srvGEntité_PutBuffer(recGEntité As typeGEntité)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recGEntité.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recGEntité.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 11) = recGEntité.Compte
Mid$(MsgTxt, K + 12, 3) = recGEntité.Devise
Mid$(MsgTxt, K + 15, 3) = recGEntité.Nature
Mid$(MsgTxt, K + 18, 2) = Format$(recGEntité.Séquence, "00")
Mid$(MsgTxt, K + 20, 3) = recGEntité.ClasseInfo
Mid$(MsgTxt, K + 23, 19) = recGEntité.AdresseId
Mid$(MsgTxt, K + 42, 1) = recGEntité.AdresseL0K
Mid$(MsgTxt, K + 43, 32) = recGEntité.AdresseL1
Mid$(MsgTxt, K + 75, 1) = FrecGEntité.AdresseRoutage
Mid$(MsgTxt, K + 76, 8) = recGEntité.AdresseAMJDébut
Mid$(MsgTxt, K + 84, 8) = recGEntité.AdresseAMJFin
'filler
Mid$(MsgTxt, K + 104, 12) = Format$(recGEntité.ElpId, "000000000000")
Mid$(MsgTxt, K + 116, 3) = Format$(recGEntité.ElpUpdate, "000")
Mid$(MsgTxt, K + 119, 10) = recGEntité.ElpControl

MsgTxtLen = MsgTxtLen + recGEntitéLen
End Sub



'---------------------------------------------------------
Private Function srvGEntité_Seek(recGEntité As typeGEntité)
'---------------------------------------------------------

srvGEntité_Seek = "?"
MsgTxtLen = 0
Call srvGEntité_PutBuffer(recGEntité)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvGEntité_GetBuffer(recGEntité)) Then
            srvGEntité_Seek = Null
        Else
            Call srvGEntité_Error(recGEntité)
        End If
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvGEntité_Snap(recGEntité As typeGEntité)
'---------------------------------------------------------
srvGEntité_Snap = "?"
MsgTxtLen = 0
Call srvGEntité_PutBuffer(recGEntité)
Call srvGEntité_PutBuffer(arrGEntité(0))
If IsNull(SndRcv()) Then
    srvGEntité_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvGEntité_GetBuffer(recGEntité)) Then
            Call arrGEntité_AddItem(recGEntité)
            arrGEntité_Suite = True
        Else
            arrGEntité_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'-----------------------------------------------------
Function srvGEntité_Update(recGEntité As typeGEntité)
'-----------------------------------------------------

srvGEntité_Update = "?"

MsgTxtLen = 0
Call srvGEntité_PutBuffer(recGEntité)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvGEntité_GetBuffer(recGEntité)) Then
        Call srvGEntité_Error(recGEntité)
        srvGEntité_Update = recGEntité.Err
        Exit Function
    Else
        srvGEntité_Update = Null
    End If
Else
    recGEntité.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recGEntité_Init(recGEntité As typeGEntité)
'---------------------------------------------------------
MsgTxt = Space$(recGEntitéLen)
MsgTxtIndex = 0
Call srvGEntité_GetBuffer(recGEntité)
recGEntité.obj = "SRVGEntité    "
End Sub

'---------------------------------------------------------
Public Sub arrGEntité_AddItem(recGEntité As typeGEntité)
'---------------------------------------------------------
          
arrGEntité_NB = arrGEntité_NB + 1
    
If arrGEntité_NB > arrGEntité_NBMax Then
    arrGEntité_NBMax = arrGEntité_NBMax + 10
    ReDim Preserve arrGEntité(arrGEntité_NBMax)
End If
            
arrGEntité(arrGEntité_NB) = recGEntité
End Sub



Public Function fctGEntité_Compare(recGEntité As typeGEntité, mGEntité As typeGEntité)
fctGEntité_Compare = Null
If recGEntité.Compte <> mGEntité.Compte Then fctGEntité_Compare = "Service": Exit Function
If recGEntité.Devise <> mGEntité.Devise Then fctGEntité_Compare = "Devise": Exit Function
If recGEntité.Nature <> mGEntité.Nature Then fctGEntité_Compare = "Nature": Exit Function
If recGEntité.Séquence <> mGEntité.Séquence Then fctGEntité_Compare = "séquence": Exit Function
If recGEntité.ClasseInfo <> mGEntité.ClasseInfo Then fctGEntité_Compare = "classeinfo": Exit Function
If recGEntité.AdresseId <> mGEntité.AdresseId Then fctGEntité_Compare = "AdresseId": Exit Function
If recGEntité.AdresseL0K <> mGEntité.AdresseL0K Then fctGEntité_Compare = "TEG": Exit Function
If recGEntité.AdresseL1 <> mGEntité.AdresseL1 Then fctGEntité_Compare = "TauxActuariel": Exit Function
If recGEntité.AdresseRoutage <> mGEntité.AdresseRoutage Then fctGEntité_Compare = "AdresseRoutage": Exit Function

If recGEntité.AdresseAMJDébut <> mGEntité.AdresseAMJDébut Then fctGEntité_Compare = "adresseAmjDébut": Exit Function
If recGEntité.AdresseAMJFin <> mGEntité.AdresseAMJFin Then fctGEntité_Compare = "adresseAmjfin": Exit Function
End Function
