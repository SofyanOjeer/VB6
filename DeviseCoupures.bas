Attribute VB_Name = "srvDeviseCoupures"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recDeviseCoupuresLen = 54 ' 34 + 20

Type typeDeviseCoupures
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Id                      As String * 3
    Nature                  As String * 1
    Nominal                 As Currency
    Séquence                As Integer
    Actif                   As String * 1
End Type
    
Public arrDeviseCoupures() As typeDeviseCoupures
Public arrDeviseCoupuresNb As Integer
Public arrDeviseCoupuresNbMax As Integer
Public arrDeviseCoupuresIndex As Integer
Public arrDeviseCoupuresSuite As Boolean
Public XDeviseCoupures As typeDeviseCoupures


'-----------------------------------------------------
Function srvDeviseCoupures_Update(recDeviseCoupures As typeDeviseCoupures)
'-----------------------------------------------------

srvDeviseCoupures_Update = "?"

MsgTxtLen = 0
Call srvDeviseCoupures_PutBuffer(recDeviseCoupures)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    recDeviseCoupures.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
    If Trim(recDeviseCoupures.Err) <> "" Then
        Call srvDeviseCoupures_Error(recDeviseCoupures)
        srvDeviseCoupures_Update = recDeviseCoupures.Err
        Exit Function
    Else
        srvDeviseCoupures_Update = Null
    End If
Else
    recDeviseCoupures.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Public Function srvDeviseCoupures_Monitor(recDeviseCoupures As typeDeviseCoupures)
'-----------------------------------------------------

arrDeviseCoupuresSuite = False
Select Case Mid$(Trim(recDeviseCoupures.Method), 1, 4)
    Case "Seek"
                srvDeviseCoupures_Monitor = srvDeviseCoupures_Seek(recDeviseCoupures)
    Case "Snap", "Prev"
              srvDeviseCoupures_Monitor = srvDeviseCoupures_Snap(recDeviseCoupures)
    Case Else
                recDeviseCoupures.Err = recDeviseCoupures.Method
                Call srvDeviseCoupures_Error(recDeviseCoupures)
                srvDeviseCoupures_Monitor = recDeviseCoupures.Err
End Select

End Function

'-----------------------------------------------------
Sub srvDeviseCoupures_Error(recDeviseCoupures As typeDeviseCoupures)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Coupures Devise : " ' & Chr$(10) & Chr$(13)

Select Case Mid$(recDeviseCoupures.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recDeviseCoupures.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvDeviseCoupures.bas  ( " _
                & Trim(recDeviseCoupures.obj) & " : " & Trim(recDeviseCoupures.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvDeviseCoupures_GetBuffer(recDeviseCoupures As typeDeviseCoupures)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvDeviseCoupures_GetBuffer = Null
recDeviseCoupures.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recDeviseCoupures.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recDeviseCoupures.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recDeviseCoupures.Err = Space$(10) Then
    recDeviseCoupures.Id = Mid$(MsgTxt, K + 1, 3)
    recDeviseCoupures.Nature = Mid$(MsgTxt, K + 4, 1)
    recDeviseCoupures.Nominal = CCur(Val(Mid$(MsgTxt, K + 5, 13)) / 10000)
    recDeviseCoupures.Séquence = CInt(Val(Mid$(MsgTxt, K + 18, 2)))
    recDeviseCoupures.Actif = Mid$(MsgTxt, K + 20, 1)

Else
    srvDeviseCoupures_GetBuffer = recDeviseCoupures.Err
End If

MsgTxtIndex = MsgTxtIndex + recDeviseCoupuresLen

End Function

'---------------------------------------------------------
Private Sub srvDeviseCoupures_PutBuffer(recDeviseCoupures As typeDeviseCoupures)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recDeviseCoupures.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recDeviseCoupures.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 3) = recDeviseCoupures.Id
Mid$(MsgTxt, K + 4, 1) = recDeviseCoupures.Nature
Mid$(MsgTxt, K + 5, 13) = Format$(recDeviseCoupures.Nominal * 10000, "0000000000000")
Mid$(MsgTxt, K + 18, 2) = Format$(recDeviseCoupures.Séquence, "00")
Mid$(MsgTxt, K + 20, 1) = recDeviseCoupures.Actif
MsgTxtLen = MsgTxtLen + recDeviseCoupuresLen
End Sub



'---------------------------------------------------------
Private Function srvDeviseCoupures_Seek(recDeviseCoupures As typeDeviseCoupures)
'---------------------------------------------------------

srvDeviseCoupures_Seek = "?"
MsgTxtLen = 0
Call srvDeviseCoupures_PutBuffer(recDeviseCoupures)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvDeviseCoupures_GetBuffer(recDeviseCoupures)) Then
        srvDeviseCoupures_Seek = Null
    Else
        Call srvDeviseCoupures_Error(recDeviseCoupures)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvDeviseCoupures_Snap(recDeviseCoupures As typeDeviseCoupures)
'---------------------------------------------------------
Dim I As Integer
srvDeviseCoupures_Snap = "?"
MsgTxtLen = 0
Call srvDeviseCoupures_PutBuffer(recDeviseCoupures)
Call srvDeviseCoupures_PutBuffer(arrDeviseCoupures(0))
If IsNull(SndRcv()) Then
    srvDeviseCoupures_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvDeviseCoupures_GetBuffer(recDeviseCoupures)) Then
            Call arrDeviseCoupures_AddItem(recDeviseCoupures)
            arrDeviseCoupuresSuite = True
        Else
            arrDeviseCoupuresSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recDeviseCoupures_Init(recDeviseCoupures As typeDeviseCoupures)
'---------------------------------------------------------
 MsgTxt = Space$(recDeviseCoupuresLen)
 MsgTxtIndex = 0
 Call srvDeviseCoupures_GetBuffer(recDeviseCoupures)
 recDeviseCoupures.obj = "SRVDEVCOUP"
End Sub

'---------------------------------------------------------
Public Sub arrDeviseCoupures_AddItem(recDeviseCoupures As typeDeviseCoupures)
'---------------------------------------------------------
          
arrDeviseCoupuresNb = arrDeviseCoupuresNb + 1
    
If arrDeviseCoupuresNb > arrDeviseCoupuresNbMax Then
    arrDeviseCoupuresNbMax = arrDeviseCoupuresNbMax + 10
    ReDim Preserve arrDeviseCoupures(arrDeviseCoupuresNbMax)
End If
            
arrDeviseCoupures(arrDeviseCoupuresNb) = recDeviseCoupures
End Sub




Public Function arrDeviseCoupures_ScanSéquence(recDeviseCoupures As typeDeviseCoupures) As Integer
arrDeviseCoupures_ScanSéquence = -1
For arrDeviseCoupuresIndex = 1 To arrDeviseCoupuresNb
    If arrDeviseCoupures(arrDeviseCoupuresIndex).Séquence = recDeviseCoupures.Séquence Then
        arrDeviseCoupures_ScanSéquence = arrDeviseCoupuresIndex
        Exit For
    End If
Next arrDeviseCoupuresIndex

End Function

Public Function arrDeviseCoupures_ScanNominal(recDeviseCoupures As typeDeviseCoupures) As Integer
arrDeviseCoupures_ScanNominal = -1
For arrDeviseCoupuresIndex = 1 To arrDeviseCoupuresNb
    If arrDeviseCoupures(arrDeviseCoupuresIndex).Method <> constDelete _
    And arrDeviseCoupures(arrDeviseCoupuresIndex).Method <> constIgnore Then
        If arrDeviseCoupures(arrDeviseCoupuresIndex).Nominal = recDeviseCoupures.Nominal _
        And arrDeviseCoupures(arrDeviseCoupuresIndex).Nature = recDeviseCoupures.Nature Then
            arrDeviseCoupures_ScanNominal = arrDeviseCoupuresIndex
            Exit For
        End If
    End If
Next arrDeviseCoupuresIndex

End Function

Public Sub srvDeviseCoupures_Load(strDev As String)
ReDim arrDeviseCoupures(10): arrDeviseCoupuresNbMax = 10
recDeviseCoupures_Init XDeviseCoupures
XDeviseCoupures.Method = "SnapP0"
XDeviseCoupures.Id = strDev

arrDeviseCoupuresNb = 0
arrDeviseCoupures(0) = XDeviseCoupures
arrDeviseCoupures(0).Nature = "9"
arrDeviseCoupuresSuite = True
Do Until Not arrDeviseCoupuresSuite
    srvDeviseCoupures_Monitor XDeviseCoupures
    XDeviseCoupures = arrDeviseCoupures(arrDeviseCoupuresNb)
    XDeviseCoupures.Method = "SnapP0+"
Loop
End Sub
