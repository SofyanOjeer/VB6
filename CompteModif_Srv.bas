Attribute VB_Name = "srvCompteModif"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recCompteModifLen = 140          ' 34 + 106

Type typeCompteModif
    obj             As String * 12
    Method          As String * 12
    Err             As String * 10
    Société         As String * 3
    Agence          As String * 3
    Devise          As String * 3
    Numéro          As String * 11
    Intitulé        As String * 40
    Intitulé2       As String * 40
    TypeGA          As String * 1
    Situation       As String * 1
    Gestionnaire    As String * 2
    Extrait         As String * 1
    Courrier        As String * 1
      
End Type
    

Public arrCompteModif() As typeCompteModif
Public arrCompteModifNb As Integer
Public arrCompteModifNbMax As Integer
Public arrCompteModifIndex As Integer
Public arrCompteModifsuite As Boolean
Public CompteModifAut As typeAuthorization

'-----------------------------------------------------
Function srvCompteModif_Update(recCompteModif As typeCompteModif)
'-----------------------------------------------------

srvCompteModif_Update = "?"

MsgTxtLen = 0
Call srvCompteModif_PutBuffer(recCompteModif)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvCompteModif_GetBuffer(recCompteModif)) Then
        Call srvCompteModif_Error(recCompteModif)
        srvCompteModif_Update = recCompteModif.Err
        Exit Function
    Else
        srvCompteModif_Update = Null
    End If
Else
    recCompteModif.Err = "srv"
End If


'=====================================================
End Function


'-----------------------------------------------------
Public Function srvCompteModif_Mon(recCompteModif As typeCompteModif)
'-----------------------------------------------------

arrCompteModifsuite = False
Select Case recCompteModif.Method
    Case "SnapL5      ", "SnapL5+     ", _
         "SnapKE      ", "SnapKE+     ", _
         "SnapLA      ", "SnapLA+     "
          srvCompteModif_Mon = srvCompteModif_Snap(recCompteModif)
    Case Else
                recCompteModif.Err = recCompteModif.Method
                Call srvCompteModif_Error(recCompteModif)
                srvCompteModif_Mon = recCompteModif.Err
End Select

End Function

'-----------------------------------------------------
Sub srvCompteModif_Error(recCompteModif As typeCompteModif)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "CompteModif: " & recCompteModif.Devise & "." & recCompteModif.Numéro & Chr$(10) & Chr$(13)

Select Case mId$(recCompteModif.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        If Trim(recCompteModif.Err) = "FIN" Then
            Msg = Msg & "Inconnu"
        Else
            Msg = Msg & "Error Code : " & recCompteModif.Err
        End If
        I = vbCritical
End Select

MsgBox Msg, I, "module : Cpt.bas  ( " _
                & Trim(recCompteModif.obj) & " : " & Trim(recCompteModif.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvCompteModif_GetBuffer(recCompteModif As typeCompteModif)
'---------------------------------------------------------
Dim K As Integer
srvCompteModif_GetBuffer = Null
recCompteModif.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recCompteModif.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recCompteModif.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recCompteModif.Err = Space$(10) Then
    recCompteModif.Société = mId$(MsgTxt, K + 1, 3)
    recCompteModif.Agence = mId$(MsgTxt, K + 4, 3)
    recCompteModif.Devise = mId$(MsgTxt, K + 7, 3)
    recCompteModif.Numéro = mId$(MsgTxt, K + 10, 11)
    recCompteModif.Intitulé = mId$(MsgTxt, K + 21, 40)
    recCompteModif.Intitulé2 = mId$(MsgTxt, K + 61, 40)
    recCompteModif.TypeGA = mId$(MsgTxt, K + 101, 1)
    recCompteModif.Situation = mId$(MsgTxt, K + 102, 1)
    recCompteModif.Gestionnaire = mId$(MsgTxt, K + 103, 2)
    recCompteModif.Extrait = mId$(MsgTxt, K + 105, 1)
    recCompteModif.Courrier = mId$(MsgTxt, K + 106, 1)
Else
    srvCompteModif_GetBuffer = recCompteModif.Err
End If

MsgTxtIndex = MsgTxtIndex + recCompteModifLen

End Function

'---------------------------------------------------------
Private Sub srvCompteModif_PutBuffer(recCompteModif As typeCompteModif)
'---------------------------------------------------------
Dim K As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recCompteModif.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recCompteModif.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 3) = recCompteModif.Société
Mid$(MsgTxt, K + 4, 3) = recCompteModif.Agence
Mid$(MsgTxt, K + 7, 3) = Format$(Val(recCompteModif.Devise), "000")
Mid$(MsgTxt, K + 10, 11) = Format$(Val(recCompteModif.Numéro), "00000000000")
Mid$(MsgTxt, K + 21, 40) = recCompteModif.Intitulé
Mid$(MsgTxt, K + 61, 40) = recCompteModif.Intitulé2
Mid$(MsgTxt, K + 101, 1) = recCompteModif.TypeGA
Mid$(MsgTxt, K + 102, 1) = recCompteModif.Situation
Mid$(MsgTxt, K + 103, 2) = recCompteModif.Gestionnaire
Mid$(MsgTxt, K + 105, 1) = recCompteModif.Extrait
Mid$(MsgTxt, K + 106, 1) = recCompteModif.Courrier

MsgTxtLen = MsgTxtLen + recCompteModifLen
End Sub



'---------------------------------------------------------
Private Function srvCompteModif_Seek(recCompteModif As typeCompteModif)
'---------------------------------------------------------

srvCompteModif_Seek = "?"
MsgTxtLen = 0
Call srvCompteModif_PutBuffer(recCompteModif)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvCompteModif_GetBuffer(recCompteModif)) Then
        srvCompteModif_Seek = Null
    Else
        Call srvCompteModif_Error(recCompteModif)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvCompteModif_Snap(recCompteModif As typeCompteModif)
'---------------------------------------------------------
srvCompteModif_Snap = "?"
MsgTxtLen = 0
Call srvCompteModif_PutBuffer(recCompteModif)
Call srvCompteModif_PutBuffer(arrCompteModif(0))
If IsNull(SndRcv()) Then
    srvCompteModif_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvCompteModif_GetBuffer(recCompteModif)) Then
            Call srvCompteModif_AddItem(recCompteModif)
            arrCompteModifsuite = True
        Else
            arrCompteModifsuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub srvCompteModif_Init(recCompteModif As typeCompteModif)
'---------------------------------------------------------
 MsgTxt = Space$(recCompteModifLen)
 MsgTxtIndex = 0
 Call srvCompteModif_GetBuffer(recCompteModif)
 recCompteModif.obj = "SRVCPTMOD   "
End Sub

'---------------------------------------------------------
Public Sub srvCompteModif_AddItem(recCompteModif As typeCompteModif)
'---------------------------------------------------------
          
arrCompteModifNb = arrCompteModifNb + 1
    
If arrCompteModifNb > arrCompteModifNbMax Then
    arrCompteModifNbMax = arrCompteModifNbMax + 10
    ReDim Preserve arrCompteModif(arrCompteModifNbMax)
End If
            
arrCompteModif(arrCompteModifNb) = recCompteModif
End Sub

