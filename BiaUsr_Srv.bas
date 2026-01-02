Attribute VB_Name = "srvBiaUsr"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recBiaUsrLen = 145 '34 + 111

Type typeBiaUsr
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Id                     As String * 10
    Nom                    As String * 34
    Service                As String * 3
    Coges                  As String * 2
    Groupe                 As String * 10

End Type
    
Public arrbiausr() As typeBiaUsr
Public arrBiaUsrNb As Integer
Public arrBiaUsrNbMax As Integer
Public arrBiaUsrIndex As Integer
Public arrBiaUsrSuite As Boolean
'-----------------------------------------------------
Public Function srvBiaUsr_Monitor(recBiaUsr As typeBiaUsr)
'-----------------------------------------------------

arrBiaUsrSuite = False
Select Case mId$(Trim(recBiaUsr.Method), 1, 4)
    Case "Seek"
                srvBiaUsr_Monitor = srvBiaUsr_Seek(recBiaUsr)
    Case "Snap"
              srvBiaUsr_Monitor = srvBiaUsr_Snap(recBiaUsr)
    Case Else
    
                recBiaUsr.Err = recBiaUsr.Method
                Call srvBiaUsr_Error(recBiaUsr)
                srvBiaUsr_Monitor = recBiaUsr.Err
End Select

End Function

'-----------------------------------------------------
Sub srvBiaUsr_Error(recBiaUsr As typeBiaUsr)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Utilisateurs BIA: " ' & Chr$(10) & Chr$(13)

Select Case mId$(recBiaUsr.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recBiaUsr.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvBiaUsr_.bas  ( " _
                & Trim(recBiaUsr.obj) & " : " & Trim(recBiaUsr.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvBiaUsr_GetBuffer(recBiaUsr As typeBiaUsr)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvBiaUsr_GetBuffer = Null
recBiaUsr.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recBiaUsr.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recBiaUsr.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recBiaUsr.Err = Space$(10) Then
    recBiaUsr.Id = mId$(MsgTxt, K + 1, 10)
    recBiaUsr.Nom = Trim(mId$(MsgTxt, K + 30, 4)) & " " & Trim(mId$(MsgTxt, K + 34, 15)) & " " & Trim(mId$(MsgTxt, K + 49, 15))
    recBiaUsr.Service = mId$(MsgTxt, K + 23, 3)
    recBiaUsr.Coges = mId$(MsgTxt, K + 94, 2)
    recBiaUsr.Groupe = mId$(MsgTxt, K + 11, 10)
Else
    srvBiaUsr_GetBuffer = recBiaUsr.Err
End If

MsgTxtIndex = MsgTxtIndex + recBiaUsrLen

End Function

'---------------------------------------------------------
Public Sub srvBiaUsr_PutBuffer(recBiaUsr As typeBiaUsr)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recBiaUsr.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recBiaUsr.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 10) = recBiaUsr.Id

MsgTxtLen = MsgTxtLen + recBiaUsrLen
End Sub



'---------------------------------------------------------
Private Function srvBiaUsr_Seek(recBiaUsr As typeBiaUsr)
'---------------------------------------------------------

srvBiaUsr_Seek = "?"
MsgTxtLen = 0
Call srvBiaUsr_PutBuffer(recBiaUsr)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvBiaUsr_GetBuffer(recBiaUsr)) Then
        srvBiaUsr_Seek = Null
    Else
 '       Call srvBiaUsr_Error(recBiaUsr)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvBiaUsr_Snap(recBiaUsr As typeBiaUsr)
'---------------------------------------------------------
Dim I As Integer
srvBiaUsr_Snap = "?"
MsgTxtLen = 0
Call srvBiaUsr_PutBuffer(recBiaUsr)
Call srvBiaUsr_PutBuffer(arrbiausr(0))
If IsNull(SndRcv()) Then
    srvBiaUsr_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvBiaUsr_GetBuffer(recBiaUsr)) Then
            Call arrBiaUsr_AddItem(recBiaUsr)
            arrBiaUsrSuite = True
        Else
            arrBiaUsrSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recBiaUsr_Init(recBiaUsr As typeBiaUsr)
'---------------------------------------------------------
recBiaUsr.obj = "SRVBIAUSR"
recBiaUsr.Method = ""
recBiaUsr.Err = ""
recBiaUsr.Id = ""
recBiaUsr.Nom = ""
recBiaUsr.Service = "999"
recBiaUsr.Coges = "99"
End Sub

'---------------------------------------------------------
Public Sub arrBiaUsr_AddItem(recBiaUsr As typeBiaUsr)
'---------------------------------------------------------
          
arrBiaUsrNb = arrBiaUsrNb + 1
    
If arrBiaUsrNb > arrBiaUsrNbMax Then
    arrBiaUsrNbMax = arrBiaUsrNbMax + 10
    ReDim Preserve arrbiausr(arrBiaUsrNbMax)
End If
            
arrbiausr(arrBiaUsrNb) = recBiaUsr
End Sub





Public Sub lstBiaUsr_Load(lstX As ListBox, recBiaUsr As typeBiaUsr)
ReDim arrbiausr(50): arrBiaUsrNbMax = 50
recBiaUsr_Init recBiaUsr
recBiaUsr.Method = "SnapP0"
arrbiausr(0) = recBiaUsr
arrbiausr(0).Id = "9z"
arrBiaUsrNb = 0
arrBiaUsrSuite = True

Do Until Not arrBiaUsrSuite

    Call srvBiaUsr_Monitor(recBiaUsr)
    recBiaUsr = arrbiausr(arrBiaUsrNb)
    recBiaUsr.Method = "SnapP0+"
    
Loop
lstX.Clear
        
For arrBiaUsrIndex = 1 To arrBiaUsrNb
'''    If Trim(arrbiausr(arrBiaUsrIndex).Groupe) <> "" Then lstX.AddItem arrbiausr(arrBiaUsrIndex).Id & "  : " & Trim(arrbiausr(arrBiaUsrIndex).Nom) & " ( G: " & arrbiausr(arrBiaUsrIndex).Coges & " ) "
    lstX.AddItem arrbiausr(arrBiaUsrIndex).Id & "  : " & Trim(arrbiausr(arrBiaUsrIndex).Nom) & " ( G: " & arrbiausr(arrBiaUsrIndex).Coges & " ) "
            
Next arrBiaUsrIndex
ReDim arrbiausr(1): arrBiaUsrNbMax = 0

End Sub
