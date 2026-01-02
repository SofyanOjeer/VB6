Attribute VB_Name = "srvDeviseCours"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recDeviseCoursLen = 173 ' 34 + 139
Type typeDeviseCours
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Id1                     As String * 3
    Id2                     As String * 3
    Amj                     As String * 8
    QD1                     As Long
    QD2CoursPivot           As Double
    QD2AchatNormal          As Double
    QD2VenteNormal          As Double
    QD2AchatPrivilégié      As Double
    QD2VentePrivilégié      As Double
    QD2AchatEnCompte        As Double
    QD2VenteEnCompte        As Double
    SaisieAMJ               As String * 8
    SaisieHMS               As String * 6
    SaisieUsr               As String * 10
    ValidationAMJ           As String * 8
    ValidationHMS           As String * 6
    ValidationUsr           As String * 10
End Type
    
Public arrDeviseCours() As typeDeviseCours
Public arrDeviseCoursNb As Integer
Public arrDeviseCoursNbMax As Integer
Public arrDeviseCoursIndex As Integer
Public arrDeviseCoursSuite As Boolean

Public XDeviseCours As typeDeviseCours
Public Sub srvDeviseCours_Load(Amj As String)
ReDim arrDeviseCours(10): arrDeviseCoursNbMax = 10
recDeviseCours_Init XDeviseCours
XDeviseCours.Method = "SnapP0"
XDeviseCours.Amj = Amj

arrDeviseCoursNb = 0
arrDeviseCours(0) = XDeviseCours
arrDeviseCours(0).Id1 = "9z"
arrDeviseCoursSuite = True
Do Until Not arrDeviseCoursSuite
    srvDeviseCours_Monitor XDeviseCours
    XDeviseCours = arrDeviseCours(arrDeviseCoursNb)
    XDeviseCours.Method = "SnapP0+"
Loop
End Sub

'-----------------------------------------------------
Function srvDeviseCours_Update(recDeviseCours As typeDeviseCours)
'-----------------------------------------------------

srvDeviseCours_Update = "?"

MsgTxtLen = 0
Call srvDeviseCours_PutBuffer(recDeviseCours)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    recDeviseCours.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
    If Trim(recDeviseCours.Err) <> "" Then
        Call srvDeviseCours_Error(recDeviseCours)
        srvDeviseCours_Update = recDeviseCours.Err
        Exit Function
    Else
        srvDeviseCours_Update = Null
    End If
Else
    recDeviseCours.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Public Function srvDeviseCours_Monitor(recDeviseCours As typeDeviseCours)
'-----------------------------------------------------

arrDeviseCoursSuite = False
Select Case Mid$(Trim(recDeviseCours.Method), 1, 4)
    Case "Seek"
                srvDeviseCours_Monitor = srvDeviseCours_Seek(recDeviseCours)
    Case "Snap"
              srvDeviseCours_Monitor = srvDeviseCours_Snap(recDeviseCours)
    Case Else
                recDeviseCours.Err = recDeviseCours.Method
                Call srvDeviseCours_Error(recDeviseCours)
                srvDeviseCours_Monitor = recDeviseCours.Err
End Select

End Function

'-----------------------------------------------------
Sub srvDeviseCours_Error(recDeviseCours As typeDeviseCours)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Devise Cours : " ' & Chr$(10) & Chr$(13)

Select Case Mid$(recDeviseCours.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recDeviseCours.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvDeviseCours.bas  ( " _
                & Trim(recDeviseCours.obj) & " : " & Trim(recDeviseCours.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvDeviseCours_GetBuffer(recDeviseCours As typeDeviseCours)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvDeviseCours_GetBuffer = Null
recDeviseCours.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recDeviseCours.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recDeviseCours.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recDeviseCours.Err = Space$(10) Then
    recDeviseCours.Id1 = Mid$(MsgTxt, K + 1, 3)
    recDeviseCours.Id2 = Mid$(MsgTxt, K + 4, 3)
    recDeviseCours.Amj = Mid$(MsgTxt, K + 7, 8)
    recDeviseCours.QD1 = CLng(Val(Mid$(MsgTxt, K + 15, 7)))
    recDeviseCours.QD2CoursPivot = CDbl(Val(Mid$(MsgTxt, K + 22, 10)) / 100000)
    recDeviseCours.QD2AchatNormal = CDbl(Val(Mid$(MsgTxt, K + 32, 10)) / 100000)
    recDeviseCours.QD2VenteNormal = CDbl(Val(Mid$(MsgTxt, K + 42, 10)) / 100000)
    recDeviseCours.QD2AchatPrivilégié = CDbl(Val(Mid$(MsgTxt, K + 52, 10)) / 100000)
    recDeviseCours.QD2VentePrivilégié = CDbl(Val(Mid$(MsgTxt, K + 62, 10)) / 100000)
    recDeviseCours.QD2AchatEnCompte = CDbl(Val(Mid$(MsgTxt, K + 72, 10)) / 100000)
    recDeviseCours.QD2VenteEnCompte = CDbl(Val(Mid$(MsgTxt, K + 82, 10)) / 100000)
    recDeviseCours.SaisieAMJ = Mid$(MsgTxt, K + 92, 8)
    recDeviseCours.SaisieHMS = Mid$(MsgTxt, K + 100, 6)
    recDeviseCours.SaisieUsr = Mid$(MsgTxt, K + 106, 10)
    recDeviseCours.ValidationAMJ = Mid$(MsgTxt, K + 116, 8)
    recDeviseCours.ValidationHMS = Mid$(MsgTxt, K + 124, 6)
    recDeviseCours.ValidationUsr = Mid$(MsgTxt, K + 130, 10)

Else
    srvDeviseCours_GetBuffer = recDeviseCours.Err
End If

MsgTxtIndex = MsgTxtIndex + recDeviseCoursLen

End Function

'---------------------------------------------------------
Private Sub srvDeviseCours_PutBuffer(recDeviseCours As typeDeviseCours)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recDeviseCours.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recDeviseCours.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 3) = recDeviseCours.Id1
Mid$(MsgTxt, K + 4, 3) = recDeviseCours.Id2
Mid$(MsgTxt, K + 7, 8) = Format$(recDeviseCours.Amj, "00000000")
Mid$(MsgTxt, K + 15, 7) = Format$(recDeviseCours.QD1, "0000000")
Mid$(MsgTxt, K + 22, 10) = Format$(recDeviseCours.QD2CoursPivot * 100000, "0000000000")
Mid$(MsgTxt, K + 32, 10) = Format$(recDeviseCours.QD2AchatNormal * 100000, "0000000000")
Mid$(MsgTxt, K + 42, 10) = Format$(recDeviseCours.QD2VenteNormal * 100000, "0000000000")
Mid$(MsgTxt, K + 52, 10) = Format$(recDeviseCours.QD2AchatPrivilégié * 100000, "0000000000")
Mid$(MsgTxt, K + 62, 10) = Format$(recDeviseCours.QD2VentePrivilégié * 100000, "0000000000")
Mid$(MsgTxt, K + 72, 10) = Format$(recDeviseCours.QD2AchatEnCompte * 100000, "0000000000")
Mid$(MsgTxt, K + 82, 10) = Format$(recDeviseCours.QD2VenteEnCompte * 100000, "0000000000")
Mid$(MsgTxt, K + 92, 8) = Format$(recDeviseCours.SaisieAMJ, "00000000")
Mid$(MsgTxt, K + 100, 6) = Format$(recDeviseCours.SaisieHMS, "000000")
Mid$(MsgTxt, K + 106, 10) = recDeviseCours.SaisieUsr
Mid$(MsgTxt, K + 116, 8) = Format$(recDeviseCours.ValidationAMJ, "00000000")
Mid$(MsgTxt, K + 124, 6) = Format$(recDeviseCours.ValidationHMS, "000000")
Mid$(MsgTxt, K + 130, 10) = recDeviseCours.ValidationUsr
MsgTxtLen = MsgTxtLen + recDeviseCoursLen
End Sub



'---------------------------------------------------------
Private Function srvDeviseCours_Seek(recDeviseCours As typeDeviseCours)
'---------------------------------------------------------

srvDeviseCours_Seek = "?"

If recDeviseCours.Id1 = recDeviseCours.Id2 Then
    recDeviseCours.QD1 = 1
    recDeviseCours.QD2CoursPivot = 1
    recDeviseCours.QD2AchatNormal = 1
    recDeviseCours.QD2VenteNormal = 1
    recDeviseCours.QD2VentePrivilégié = 1
    recDeviseCours.QD2AchatEnCompte = 1
    recDeviseCours.QD2VenteEnCompte = 1
    recDeviseCours.SaisieUsr = 1
    recDeviseCours.ValidationUsr = 1
    srvDeviseCours_Seek = Null
    Exit Function
End If

MsgTxtLen = 0
Call srvDeviseCours_PutBuffer(recDeviseCours)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvDeviseCours_GetBuffer(recDeviseCours)) Then
        srvDeviseCours_Seek = Null
'    Else
'        Call srvDeviseCours_Error(recDeviseCours)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvDeviseCours_Snap(recDeviseCours As typeDeviseCours)
'---------------------------------------------------------
Dim I As Integer
srvDeviseCours_Snap = "?"
MsgTxtLen = 0
Call srvDeviseCours_PutBuffer(recDeviseCours)
Call srvDeviseCours_PutBuffer(arrDeviseCours(0))
If IsNull(SndRcv()) Then
    srvDeviseCours_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvDeviseCours_GetBuffer(recDeviseCours)) Then
            Call arrDeviseCours_AddItem(recDeviseCours)
            arrDeviseCoursSuite = True
        Else
            arrDeviseCoursSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recDeviseCours_Init(recDeviseCours As typeDeviseCours)
'---------------------------------------------------------
 MsgTxt = Space$(recDeviseCoursLen)
 MsgTxtIndex = 0
 Call srvDeviseCours_GetBuffer(recDeviseCours)
 recDeviseCours.obj = "SRVDEVCRS"
End Sub

'---------------------------------------------------------
Public Sub arrDeviseCours_AddItem(recDeviseCours As typeDeviseCours)
'---------------------------------------------------------
          
arrDeviseCoursNb = arrDeviseCoursNb + 1
    
If arrDeviseCoursNb > arrDeviseCoursNbMax Then
    arrDeviseCoursNbMax = arrDeviseCoursNbMax + 10
    ReDim Preserve arrDeviseCours(arrDeviseCoursNbMax)
End If
            
arrDeviseCours(arrDeviseCoursNb) = recDeviseCours
End Sub




Public Function arrDeviseCours_ScanId1Id2(recDeviseCours As typeDeviseCours) As Integer
arrDeviseCours_ScanId1Id2 = -1
For arrDeviseCoursIndex = 1 To arrDeviseCoursNb
    If arrDeviseCours(arrDeviseCoursIndex).Method <> constDelete _
    And arrDeviseCours(arrDeviseCoursIndex).Method <> constIgnore Then
        If arrDeviseCours(arrDeviseCoursIndex).Id1 = recDeviseCours.Id1 _
        And arrDeviseCours(arrDeviseCoursIndex).Id2 = recDeviseCours.Id2 Then
            arrDeviseCours_ScanId1Id2 = arrDeviseCoursIndex
            Exit For
        End If
    End If
Next arrDeviseCoursIndex

End Function

Public Sub recDeviseCours_Inità1(recDeviseCours As typeDeviseCours)
recDeviseCours_Init recDeviseCours
recDeviseCours.QD1 = 1
recDeviseCours.QD2AchatNormal = 1
recDeviseCours.QD2VenteNormal = 1
recDeviseCours.QD2AchatPrivilégié = 1
recDeviseCours.QD2VentePrivilégié = 1
recDeviseCours.QD2AchatEnCompte = 1
recDeviseCours.QD2VenteEnCompte = 1

End Sub
