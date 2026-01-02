Attribute VB_Name = "srvAdresse"
Option Explicit


'---------------------------------------------------------
Public Const recAdresseLen = 280        ' 34 + 246

Type typeAdresse
    obj        As String * 12
    Method     As String * 12
    Err        As String * 10
    Numéro     As String * 11
    Séquence   As Integer
    Adresse1   As String * 40
    Adresse2   As String * 32
    Adresse3   As String * 32
    Adresse4   As String * 32
    Adresse5   As String * 32
    AdresseCP   As String * 5
    AdresseBD   As String * 27
    AdressePays As String * 32

End Type
    

Public arrAdresse() As typeAdresse
Public arrAdresseNb As Integer
Public arrAdresseNbMax As Integer
Public arrAdresseIndex As Integer
Public arrAdressesuite As Boolean

'---------------------------------------------------------
Public Sub Init(recAdresse As typeAdresse)
'---------------------------------------------------------
 MsgTxt = Space$(recAdresseLen)
 MsgTxtIndex = 0
 Call GetBuffer(recAdresse)
 recAdresse.obj = "SRVADRESSE  "
End Sub


'-----------------------------------------------------
Sub dbError(recAdresse As typeAdresse)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Adresse" & Chr$(10) & Chr$(13)

Select Case Mid$(recAdresse.Err, 9, 2)
    Case "22": Msg = Msg & "Existe déjà": I = vbExclamation
    Case "23": Msg = Msg & "N'existe pas": I = vbExclamation
    Case Else: Msg = Msg & "Error Code : " & recAdresse.Err: I = vbCritical
End Select

MsgBox Msg, I, "module : Cpt.bas  ( " _
                & Trim(recAdresse.obj) & " : " & Trim(recAdresse.Method) & " )"

End Sub


'---------------------------------------------------------
Public Function GetBuffer(recAdresse As typeAdresse)
'---------------------------------------------------------
Dim K As Integer
GetBuffer = Null
recAdresse.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recAdresse.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recAdresse.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recAdresse.Err = Space$(10) Then
    recAdresse.Numéro = Mid$(MsgTxt, K + 1, 11)
    recAdresse.Séquence = Val(Mid$(MsgTxt, K + 12, 3))
    recAdresse.Adresse1 = Mid$(MsgTxt, K + 15, 40)
    recAdresse.Adresse2 = Mid$(MsgTxt, K + 55, 32)
    recAdresse.Adresse3 = Mid$(MsgTxt, K + 87, 32)
    recAdresse.Adresse4 = Mid$(MsgTxt, K + 119, 32)
    recAdresse.Adresse5 = Mid$(MsgTxt, K + 151, 32)
    recAdresse.AdresseCP = Mid$(MsgTxt, K + 183, 5)
    recAdresse.AdresseBD = Mid$(MsgTxt, K + 188, 27)
    recAdresse.AdressePays = Mid$(MsgTxt, K + 215, 32)
Else
    GetBuffer = recAdresse.Err
End If

MsgTxtIndex = MsgTxtIndex + recAdresseLen

End Function
'---------------------------------------------------------
Private Sub PutBuffer(recAdresse As typeAdresse)
'---------------------------------------------------------
Dim K As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recAdresse.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recAdresse.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 11) = recAdresse.Numéro
Mid$(MsgTxt, K + 12, 3) = Format(recAdresse.Séquence, "000")
Mid$(MsgTxt, K + 15, 40) = recAdresse.Adresse1
Mid$(MsgTxt, K + 55, 32) = recAdresse.Adresse2
Mid$(MsgTxt, K + 87, 32) = recAdresse.Adresse3
Mid$(MsgTxt, K + 119, 32) = recAdresse.Adresse4
Mid$(MsgTxt, K + 151, 32) = recAdresse.Adresse5
Mid$(MsgTxt, K + 183, 5) = recAdresse.AdresseCP
Mid$(MsgTxt, K + 188, 27) = recAdresse.AdresseBD
Mid$(MsgTxt, K + 215, 32) = recAdresse.AdressePays

MsgTxtLen = MsgTxtLen + recAdresseLen
End Sub

'---------------------------------------------------------
Private Function dbSeek(recAdresse As typeAdresse)
'---------------------------------------------------------

dbSeek = "?"
MsgTxtLen = 0
Call PutBuffer(recAdresse)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(GetBuffer(recAdresse)) Then
        dbSeek = Null
    Else
        Call dbError(recAdresse)
    End If
End If

End Function

'---------------------------------------------------------
Private Function dbSnap(recAdresse As typeAdresse)
'---------------------------------------------------------
dbSnap = "?"
MsgTxtLen = 0
Call PutBuffer(recAdresse)
Call PutBuffer(arrAdresse(0))
If IsNull(SndRcv()) Then
    dbSnap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(GetBuffer(recAdresse)) Then
            Call AddItem(recAdresse)
            arrAdressesuite = True
        Else
            arrAdressesuite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub AddItem(recAdresse As typeAdresse)
'---------------------------------------------------------
          
arrAdresseNb = arrAdresseNb + 1
    
If arrAdresseNb > arrAdresseNbMax Then
    arrAdresseNbMax = arrAdresseNbMax + 10
    ReDim Preserve arrAdresse(arrAdresseNbMax)
End If
            
arrAdresse(arrAdresseNb) = recAdresse
End Sub

'-----------------------------------------------------
Public Function Monitor(recAdresse As typeAdresse)
'-----------------------------------------------------

arrAdressesuite = False
Select Case recAdresse.Method
    Case "SeekL0      "
            Monitor = dbSeek(recAdresse)
    Case "SnapL0      ", "SnapL0+     "
          Monitor = dbSnap(recAdresse)
    Case Else
            recAdresse.Err = recAdresse.Method
            Call dbError(recAdresse)
            Monitor = recAdresse.Err
End Select

End Function



