Attribute VB_Name = "srvOpEch"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recOpEchLen = 98 ' 34 + 64

Type typeOpEch
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Référence               As String * 6
    EchéanceAmj             As String * 8
    EchéanceHms             As String * 6
    EchéanceNature          As String * 10
    ActionAmj               As String * 8
    ActionHms               As String * 6
    ActionNature            As String * 10
    ActionOpérateur         As String * 10
End Type
    
Public arrOpEch() As typeOpEch
Public arrOpEchNb As Integer
Public arrOpEchNbMax As Integer
Public arrOpEchIndex As Integer
Public arrOpEchSuite As Boolean
'-----------------------------------------------------
Public Function srvOpEch_Find(recOpECh As typeOpEch)
'-----------------------------------------------------
Dim I As Integer, X As Variant

X = "?"

For I = 1 To arrOpEchNb
    If recOpECh.Référence = arrOpEch(I).Référence Then
        X = Null
        Exit For
'!!!!!!!========
    End If
Next I

If Not IsNull(X) Then
    I = arrOpEchNb + 1
    recOpECh.Method = "SnapP0      "
    arrOpEch(0) = recOpECh
    X = srvOpEch_Snap(recOpECh)
    If I > arrOpEchNb Then X = "inconnu"
End If

If IsNull(X) Then
    arrOpEchIndex = I
    recOpECh = arrOpEch(I)
End If

srvOpEch_Find = X

End Function

'-----------------------------------------------------
Function srvOpEch_Update(recOpECh As typeOpEch)
'-----------------------------------------------------

srvOpEch_Update = "?"

MsgTxtLen = 0
Call srvOpEch_PutBuffer(recOpECh)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    recOpECh.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
    If Trim(recOpECh.Err) <> "" Then
        Call srvOpEch_Error(recOpECh)
        srvOpEch_Update = recOpECh.Err
        Exit Function
    Else
        srvOpEch_Update = Null
    End If
Else
    recOpECh.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Public Function srvOpEch_Monitor(recOpECh As typeOpEch)
'-----------------------------------------------------

arrOpEchSuite = False
Select Case Mid$(Trim(recOpECh.Method), 1, 4)
    Case "Seek"
                srvOpEch_Monitor = srvOpEch_Seek(recOpECh)
    Case "Snap", "Prev"
              srvOpEch_Monitor = srvOpEch_Snap(recOpECh)
    Case Else
                recOpECh.Err = recOpECh.Method
                Call srvOpEch_Error(recOpECh)
                srvOpEch_Monitor = recOpECh.Err
End Select

End Function

'-----------------------------------------------------
Sub srvOpEch_Error(recOpECh As typeOpEch)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Opération de transfert(échéancier) : " ' & Chr$(10) & Chr$(13)

Select Case Mid$(recOpECh.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recOpECh.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvOpEch_.bas  ( " _
                & Trim(recOpECh.obj) & " : " & Trim(recOpECh.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvOpEch_GetBuffer(recOpECh As typeOpEch)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvOpEch_GetBuffer = Null
recOpECh.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recOpECh.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recOpECh.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recOpECh.Err = Space$(10) Then
    recOpECh.Référence = Mid$(MsgTxt, K + 1, 6)
    recOpECh.EchéanceAmj = Format$(Mid$(MsgTxt, K + 7, 8), "00000000")
    recOpECh.EchéanceHms = Format$(Mid$(MsgTxt, K + 15, 6), "000000")
    recOpECh.EchéanceNature = Mid$(MsgTxt, K + 21, 10)
    recOpECh.ActionAmj = Format$(Mid$(MsgTxt, K + 31, 8), "00000000")
    recOpECh.ActionHms = Format$(Mid$(MsgTxt, K + 39, 6), "000000")
    recOpECh.ActionNature = Mid$(MsgTxt, K + 45, 10)
    recOpECh.ActionOpérateur = Mid$(MsgTxt, K + 55, 10)
Else
    srvOpEch_GetBuffer = recOpECh.Err
End If

MsgTxtIndex = MsgTxtIndex + recOpEchLen

End Function

'---------------------------------------------------------
Private Sub srvOpEch_PutBuffer(recOpECh As typeOpEch)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recOpECh.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recOpECh.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 6) = recOpECh.Référence
Mid$(MsgTxt, K + 7, 8) = recOpECh.EchéanceAmj
Mid$(MsgTxt, K + 15, 6) = recOpECh.EchéanceHms
Mid$(MsgTxt, K + 21, 10) = recOpECh.EchéanceNature
Mid$(MsgTxt, K + 31, 8) = recOpECh.ActionAmj
Mid$(MsgTxt, K + 39, 6) = recOpECh.ActionHms
Mid$(MsgTxt, K + 45, 10) = recOpECh.ActionNature
Mid$(MsgTxt, K + 55, 10) = recOpECh.ActionOpérateur

MsgTxtLen = MsgTxtLen + recOpEchLen
End Sub



'---------------------------------------------------------
Private Function srvOpEch_Seek(recOpECh As typeOpEch)
'---------------------------------------------------------

srvOpEch_Seek = "?"
MsgTxtLen = 0
Call srvOpEch_PutBuffer(recOpECh)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvOpEch_GetBuffer(recOpECh)) Then
        srvOpEch_Seek = Null
    Else
        Call srvOpEch_Error(recOpECh)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvOpEch_Snap(recOpECh As typeOpEch)
'---------------------------------------------------------
Dim I As Integer
srvOpEch_Snap = "?"
MsgTxtLen = 0
Call srvOpEch_PutBuffer(recOpECh)
Call srvOpEch_PutBuffer(arrOpEch(0))
If IsNull(SndRcv()) Then
    srvOpEch_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvOpEch_GetBuffer(recOpECh)) Then
            Call arrOpEch_AddItem(recOpECh)
            arrOpEchSuite = True
        Else
            arrOpEchSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recOpEch_Init(recOpECh As typeOpEch)
'---------------------------------------------------------
MsgTxt = Space$(recOpEchLen)
MsgTxtIndex = 0
Call srvOpEch_GetBuffer(recOpECh)
recOpECh.obj = "SRVOPECH"
End Sub

'---------------------------------------------------------
Public Sub arrOpEch_AddItem(recOpECh As typeOpEch)
'---------------------------------------------------------
          
arrOpEchNb = arrOpEchNb + 1
    
If arrOpEchNb > arrOpEchNbMax Then
    arrOpEchNbMax = arrOpEchNbMax + 10
    ReDim Preserve arrOpEch(arrOpEchNbMax)
End If
            
arrOpEch(arrOpEchNb) = recOpECh
End Sub


Public Sub recOpEch_Display(XpicBox As PictureBox, recOpECh As typeOpEch)
XpicBox.ForeColor = warnUsrColor
XpicBox.CurrentX = 50
XpicBox.Print recOpECh.Référence;
XpicBox.ForeColor = libUsr.ForeColor
XpicBox.CurrentX = 1000
XpicBox.Print recOpECh.EchéanceNature;
XpicBox.CurrentX = 2000
XpicBox.Print dateImp(recOpECh.EchéanceAmj);
XpicBox.CurrentX = 3200
XpicBox.Print timeImp(recOpECh.EchéanceHms);
XpicBox.ForeColor = warnUsrColor
XpicBox.CurrentX = 4500
XpicBox.Print recOpECh.ActionOpérateur;
XpicBox.ForeColor = libUsr.ForeColor
XpicBox.CurrentX = 5500
XpicBox.Print recOpECh.ActionNature;
XpicBox.CurrentX = 6800
XpicBox.Print dateImp(recOpECh.ActionAmj);
XpicBox.CurrentX = 8000
XpicBox.Print timeImp(recOpECh.ActionHms);

End Sub

