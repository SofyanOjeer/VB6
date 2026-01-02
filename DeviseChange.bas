Attribute VB_Name = "srvDeviseChange"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recDeviseChangeLen = 178 ' 34 + 144
Type typeDeviseChange
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Id1                     As String * 3
    Id2                     As String * 3
    Amj                     As String * 8
    Origine                 As String * 1
    HHMM                    As String * 4
    QD1                     As Long
    QD2CoursPivot           As Double
    QD2AchatNormal          As Double
    QD2VenteNormal          As Double
    QD2AchatPrivilégié      As Double
    QD2VentePrivilégié      As Double
    QD2AchatEnCompte        As Double
    QD2VenteEnCompte        As Double
    SaisieAmj               As String * 8
    SaisieHMS               As String * 6
    SaisieUsr               As String * 10
    ValidationAMJ           As String * 8
    ValidationHMS           As String * 6
    ValidationUsr           As String * 10
End Type
    
Public arrDeviseChange() As typeDeviseChange
Public arrDeviseChangeNb As Integer
Public arrDeviseChangeNbMax As Integer
Public arrDeviseChangeIndex As Integer
Public arrDeviseChangeSuite As Boolean

Public XDeviseChange As typeDeviseChange
Public Sub srvDeviseChange_Load(Origine As String, Amj As String)
ReDim arrDeviseChange(10): arrDeviseChangeNbMax = 10
recDeviseChange_Init XDeviseChange
XDeviseChange.Method = "SnapP0"
XDeviseChange.Origine = Origine
XDeviseChange.Amj = Amj

'XDeviseChange.Method = "SnapL0"
'XDeviseChange.Id1 = "EUR"
'XDeviseChange.Id2 = "USD"
'XDeviseChange.Amj = 19990101

arrDeviseChangeNb = 0
arrDeviseChange(0) = XDeviseChange
arrDeviseChange(0).Id1 = "99999999999"
arrDeviseChangeSuite = True
Do Until Not arrDeviseChangeSuite
    srvDeviseChange_Monitor XDeviseChange
    XDeviseChange = arrDeviseChange(arrDeviseChangeNb)
    XDeviseChange.Method = "SnapP0+"
Loop
End Sub

'-----------------------------------------------------
Function srvDeviseChange_Update(recDeviseChange As typeDeviseChange)
'-----------------------------------------------------

srvDeviseChange_Update = "?"

MsgTxtLen = 0
Call srvDeviseChange_PutBuffer(recDeviseChange)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    recDeviseChange.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
    If Trim(recDeviseChange.Err) <> "" Then
        Call srvDeviseChange_Error(recDeviseChange)
        srvDeviseChange_Update = recDeviseChange.Err
        Exit Function
    Else
        srvDeviseChange_Update = Null
    End If
Else
    recDeviseChange.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Public Function srvDeviseChange_Monitor(recDeviseChange As typeDeviseChange)
'-----------------------------------------------------

arrDeviseChangeSuite = False
Select Case mId$(Trim(recDeviseChange.Method), 1, 4)
    Case "Seek", "Last"
                srvDeviseChange_Monitor = srvDeviseChange_Seek(recDeviseChange)
    Case "Snap"
              srvDeviseChange_Monitor = srvDeviseChange_Snap(recDeviseChange)
    Case Else
                recDeviseChange.Err = recDeviseChange.Method
                Call srvDeviseChange_Error(recDeviseChange)
                srvDeviseChange_Monitor = recDeviseChange.Err
End Select

End Function

'-----------------------------------------------------
Sub srvDeviseChange_Error(recDeviseChange As typeDeviseChange)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Devise Cours : " ' & Chr$(10) & Chr$(13)

Select Case mId$(recDeviseChange.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recDeviseChange.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvDeviseChange.bas  ( " _
                & Trim(recDeviseChange.obj) & " : " & Trim(recDeviseChange.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvDeviseChange_GetBuffer(recDeviseChange As typeDeviseChange)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvDeviseChange_GetBuffer = Null
recDeviseChange.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recDeviseChange.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recDeviseChange.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recDeviseChange.Err = Space$(10) Then
    recDeviseChange.Id1 = mId$(MsgTxt, K + 1, 3)
    recDeviseChange.Id2 = mId$(MsgTxt, K + 4, 3)
    recDeviseChange.Amj = mId$(MsgTxt, K + 7, 8)
    recDeviseChange.Origine = mId$(MsgTxt, K + 15, 1)
    recDeviseChange.HHMM = Format$(Val(mId$(MsgTxt, K + 16, 4)), "0000")
    recDeviseChange.QD1 = CLng(Val(mId$(MsgTxt, K + 20, 7)))
    recDeviseChange.QD2CoursPivot = CDbl(Val(mId$(MsgTxt, K + 27, 10)) / 100000)
    recDeviseChange.QD2AchatNormal = CDbl(Val(mId$(MsgTxt, K + 37, 10)) / 100000)
    recDeviseChange.QD2VenteNormal = CDbl(Val(mId$(MsgTxt, K + 47, 10)) / 100000)
    recDeviseChange.QD2AchatPrivilégié = CDbl(Val(mId$(MsgTxt, K + 57, 10)) / 100000)
    recDeviseChange.QD2VentePrivilégié = CDbl(Val(mId$(MsgTxt, K + 67, 10)) / 100000)
    recDeviseChange.QD2AchatEnCompte = CDbl(Val(mId$(MsgTxt, K + 77, 10)) / 100000)
    recDeviseChange.QD2VenteEnCompte = CDbl(Val(mId$(MsgTxt, K + 87, 10)) / 100000)
    recDeviseChange.SaisieAmj = mId$(MsgTxt, K + 97, 8)
    recDeviseChange.SaisieHMS = mId$(MsgTxt, K + 105, 6)
    recDeviseChange.SaisieUsr = mId$(MsgTxt, K + 111, 10)
    recDeviseChange.ValidationAMJ = mId$(MsgTxt, K + 121, 8)
    recDeviseChange.ValidationHMS = mId$(MsgTxt, K + 129, 6)
    recDeviseChange.ValidationUsr = mId$(MsgTxt, K + 135, 10)

Else
    srvDeviseChange_GetBuffer = recDeviseChange.Err
End If

MsgTxtIndex = MsgTxtIndex + recDeviseChangeLen

End Function

'---------------------------------------------------------
Private Sub srvDeviseChange_PutBuffer(recDeviseChange As typeDeviseChange)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recDeviseChange.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recDeviseChange.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
Mid$(MsgTxt, K + 1, 3) = recDeviseChange.Id1
Mid$(MsgTxt, K + 4, 3) = recDeviseChange.Id2
Mid$(MsgTxt, K + 7, 8) = Format$(recDeviseChange.Amj, "00000000")
Mid$(MsgTxt, K + 15, 1) = recDeviseChange.Origine
Mid$(MsgTxt, K + 16, 4) = Format$(recDeviseChange.HHMM, "0000")
Mid$(MsgTxt, K + 20, 7) = Format$(recDeviseChange.QD1, "0000000")
Mid$(MsgTxt, K + 27, 10) = Format$(recDeviseChange.QD2CoursPivot * 100000, "0000000000")
Mid$(MsgTxt, K + 37, 10) = Format$(recDeviseChange.QD2AchatNormal * 100000, "0000000000")
Mid$(MsgTxt, K + 47, 10) = Format$(recDeviseChange.QD2VenteNormal * 100000, "0000000000")
Mid$(MsgTxt, K + 57, 10) = Format$(recDeviseChange.QD2AchatPrivilégié * 100000, "0000000000")
Mid$(MsgTxt, K + 67, 10) = Format$(recDeviseChange.QD2VentePrivilégié * 100000, "0000000000")
Mid$(MsgTxt, K + 77, 10) = Format$(recDeviseChange.QD2AchatEnCompte * 100000, "0000000000")
Mid$(MsgTxt, K + 87, 10) = Format$(recDeviseChange.QD2VenteEnCompte * 100000, "0000000000")
Mid$(MsgTxt, K + 97, 8) = Format$(recDeviseChange.SaisieAmj, "00000000")
Mid$(MsgTxt, K + 105, 6) = Format$(recDeviseChange.SaisieHMS, "000000")
Mid$(MsgTxt, K + 111, 10) = recDeviseChange.SaisieUsr
Mid$(MsgTxt, K + 121, 8) = Format$(recDeviseChange.ValidationAMJ, "00000000")
Mid$(MsgTxt, K + 129, 6) = Format$(recDeviseChange.ValidationHMS, "000000")
Mid$(MsgTxt, K + 135, 10) = recDeviseChange.ValidationUsr
MsgTxtLen = MsgTxtLen + recDeviseChangeLen
End Sub



'---------------------------------------------------------
Private Function srvDeviseChange_Seek(recDeviseChange As typeDeviseChange)
'---------------------------------------------------------

srvDeviseChange_Seek = "?"


MsgTxtLen = 0
Call srvDeviseChange_PutBuffer(recDeviseChange)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvDeviseChange_GetBuffer(recDeviseChange)) Then
        srvDeviseChange_Seek = Null
'    Else
'        Call srvDeviseChange_Error(recDeviseChange)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvDeviseChange_Find(recDeviseChange As typeDeviseChange)
'---------------------------------------------------------

srvDeviseChange_Find = "?"

If recDeviseChange.Id1 = recDeviseChange.Id2 Then
    recDeviseChange.QD1 = 1
    recDeviseChange.QD2CoursPivot = 1
    recDeviseChange.QD2AchatNormal = 1
    recDeviseChange.QD2VenteNormal = 1
    recDeviseChange.QD2VentePrivilégié = 1
    recDeviseChange.QD2AchatEnCompte = 1
    recDeviseChange.QD2VenteEnCompte = 1
    recDeviseChange.SaisieUsr = 1
    recDeviseChange.ValidationUsr = 1
    srvDeviseChange_Find = Null
    Exit Function
End If
srvDeviseChange_Find = srvDeviseChange_Seek(recDeviseChange)

End Function


'---------------------------------------------------------
Private Function srvDeviseChange_Snap(recDeviseChange As typeDeviseChange)
'---------------------------------------------------------
Dim I As Integer
srvDeviseChange_Snap = "?"
MsgTxtLen = 0
Call srvDeviseChange_PutBuffer(recDeviseChange)
Call srvDeviseChange_PutBuffer(arrDeviseChange(0))
If IsNull(SndRcv()) Then
    srvDeviseChange_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvDeviseChange_GetBuffer(recDeviseChange)) Then
            Call arrDeviseChange_AddItem(recDeviseChange)
            arrDeviseChangeSuite = True
        Else
            arrDeviseChangeSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recDeviseChange_Init(recDeviseChange As typeDeviseChange)
'---------------------------------------------------------
 MsgTxt = Space$(recDeviseChangeLen)
 MsgTxtIndex = 0
 Call srvDeviseChange_GetBuffer(recDeviseChange)
 recDeviseChange.obj = "SRVDEVCHG"
End Sub

'---------------------------------------------------------
Public Sub arrDeviseChange_AddItem(recDeviseChange As typeDeviseChange)
'---------------------------------------------------------
          
arrDeviseChangeNb = arrDeviseChangeNb + 1
    
If arrDeviseChangeNb > arrDeviseChangeNbMax Then
    arrDeviseChangeNbMax = arrDeviseChangeNbMax + 10
    ReDim Preserve arrDeviseChange(arrDeviseChangeNbMax)
End If
            
arrDeviseChange(arrDeviseChangeNb) = recDeviseChange
End Sub




Public Function arrDeviseChange_ScanId1Id2(recDeviseChange As typeDeviseChange) As Integer
arrDeviseChange_ScanId1Id2 = -1
For arrDeviseChangeIndex = 1 To arrDeviseChangeNb
    If arrDeviseChange(arrDeviseChangeIndex).Method <> constDelete _
    And arrDeviseChange(arrDeviseChangeIndex).Method <> constIgnore Then
        If arrDeviseChange(arrDeviseChangeIndex).Id1 = recDeviseChange.Id1 _
        And arrDeviseChange(arrDeviseChangeIndex).Id2 = recDeviseChange.Id2 Then
            arrDeviseChange_ScanId1Id2 = arrDeviseChangeIndex
            Exit For
        End If
    End If
Next arrDeviseChangeIndex

End Function

Public Sub recDeviseChange_Inità1(recDeviseChange As typeDeviseChange)
recDeviseChange_Init recDeviseChange
recDeviseChange.QD1 = 1
recDeviseChange.QD2AchatNormal = 1
recDeviseChange.QD2VenteNormal = 1
recDeviseChange.QD2AchatPrivilégié = 1
recDeviseChange.QD2VentePrivilégié = 1
recDeviseChange.QD2AchatEnCompte = 1
recDeviseChange.QD2VenteEnCompte = 1

End Sub
