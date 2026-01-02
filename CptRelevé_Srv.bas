Attribute VB_Name = "srvCptRelevé"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recCptRelevéLen = 58        ' 34 + 24

Type typeCptRelevé
    obj        As String * 12
    Method     As String * 12
    Err        As String * 10
    Société    As String * 3
    Agence     As String * 3
    Devise     As String * 3
    Numéro     As String * 11
    Gestionnaire       As String * 2
    Courrier           As String * 1
    ExtraitPériodicité As String * 1

End Type
    

Public arrCptRelevé() As typeCptRelevé
Public arrCptRelevéNb As Integer
Public arrCptRelevéNbMax As Integer
Public arrCptRelevéIndex As Integer
Public arrCptRelevésuite As Boolean
'-----------------------------------------------------
Public Function Monitor(recCptRelevé As typeCptRelevé)
'-----------------------------------------------------

arrCptRelevésuite = False
Select Case recCptRelevé.Method
    Case "SnapKE      ", "SnapKE+     "
          Monitor = SnapX(recCptRelevé)
    Case Else
                recCptRelevé.Err = recCptRelevé.Method
                Call srvCptRelevé.ErrorX(recCptRelevé)
                Monitor = recCptRelevé.Err
End Select

End Function

'-----------------------------------------------------
Sub ErrorX(recCptRelevé As typeCptRelevé)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Compte" & Chr$(10) & Chr$(13)

Select Case Mid$(recCptRelevé.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recCptRelevé.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvCptRelevé.bas  ( " _
                & Trim(recCptRelevé.obj) & " : " & Trim(recCptRelevé.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function GetBuffer(recCptRelevé As typeCptRelevé)
'---------------------------------------------------------
Dim K As Integer
GetBuffer = Null
recCptRelevé.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recCptRelevé.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recCptRelevé.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recCptRelevé.Err = Space$(10) Then
    recCptRelevé.Société = Mid$(MsgTxt, K + 1, 3)
    recCptRelevé.Agence = Mid$(MsgTxt, K + 4, 3)
    recCptRelevé.Devise = Mid$(MsgTxt, K + 7, 3)
    recCptRelevé.Numéro = Mid$(MsgTxt, K + 10, 11)
    recCptRelevé.Gestionnaire = Mid$(MsgTxt, K + 21, 2)
    recCptRelevé.Courrier = Mid$(MsgTxt, K + 23, 1)
    recCptRelevé.ExtraitPériodicité = Mid$(MsgTxt, K + 24, 1)

  
Else
    GetBuffer = recCptRelevé.Err
End If

MsgTxtIndex = MsgTxtIndex + recCptRelevéLen

End Function

'---------------------------------------------------------
Private Sub PutBuffer(recCptRelevé As typeCptRelevé)
'---------------------------------------------------------
Dim K As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recCptRelevé.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recCptRelevé.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 3) = recCptRelevé.Société
Mid$(MsgTxt, K + 4, 3) = recCptRelevé.Agence
Mid$(MsgTxt, K + 7, 3) = recCptRelevé.Devise
Mid$(MsgTxt, K + 10, 11) = recCptRelevé.Numéro
Mid$(MsgTxt, K + 21, 2) = recCptRelevé.Gestionnaire
Mid$(MsgTxt, K + 23, 1) = recCptRelevé.Courrier
Mid$(MsgTxt, K + 24, 1) = recCptRelevé.ExtraitPériodicité

MsgTxtLen = MsgTxtLen + recCptRelevéLen
End Sub



'---------------------------------------------------------
Private Function SnapX(recCptRelevé As typeCptRelevé)
'---------------------------------------------------------
SnapX = "?"
MsgTxtLen = 0
Call srvCptRelevé.PutBuffer(recCptRelevé)
Call srvCptRelevé.PutBuffer(arrCptRelevé(0))
If IsNull(SndRcv()) Then
    SnapX = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvCptRelevé.GetBuffer(recCptRelevé)) Then
            Call srvCptRelevé.AddItem(recCptRelevé)
            arrCptRelevésuite = True
        Else
            arrCptRelevésuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub Init(recCptRelevé As typeCptRelevé)
'---------------------------------------------------------
 MsgTxt = Space$(recCptRelevéLen)
 MsgTxtIndex = 0
 Call srvCptRelevé.GetBuffer(recCptRelevé)
 recCptRelevé.obj = "SRVCOMPTE   "
End Sub

'---------------------------------------------------------
Public Sub AddItem(recCptRelevé As typeCptRelevé)
'---------------------------------------------------------
          
arrCptRelevéNb = arrCptRelevéNb + 1
    
If arrCptRelevéNb > arrCptRelevéNbMax Then
    arrCptRelevéNbMax = arrCptRelevéNbMax + 10
    ReDim Preserve arrCptRelevé(arrCptRelevéNbMax)
End If
            
arrCptRelevé(arrCptRelevéNb) = recCptRelevé
End Sub
