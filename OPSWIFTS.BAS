Attribute VB_Name = "srvOpSwift"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recOpSwiftLen = 79 ' 34 + 45

Type typeOpSwift
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Référence               As String * 6
    FieldNature             As String * 3
    FieldLine               As String * 1
    FieldText               As String * 35
End Type
Public arrOpSwift() As typeOpSwift
Public arrOpSwiftNb As Integer
Public arrOpSwiftNbMax As Integer
Public arrOpSwiftIndex As Integer
Public arrOpSwiftSuite As Boolean

Public Sub recOpSwift_Display(XpicBox As PictureBox, recOpSwift As typeOpSwift)
XpicBox.ForeColor = warnUsrColor
XpicBox.CurrentX = 50
XpicBox.Print recOpSwift.Référence;
XpicBox.ForeColor = libUsr.ForeColor
XpicBox.CurrentX = 1000
XpicBox.Print recOpSwift.FieldText;
XpicBox.CurrentX = 2000

End Sub
  
'-----------------------------------------------------
Public Function srvOpSwift_Find(recOpSwift As typeOpSwift)
'-----------------------------------------------------
Dim I As Integer, X As Variant

X = "?"

For I = 1 To arrOpSwiftNb
    If recOpSwift.Référence = arrOpSwift(I).Référence Then
        X = Null
        Exit For
'!!!!!!!========
    End If
Next I

If Not IsNull(X) Then
    I = arrOpSwiftNb + 1
    recOpSwift.Method = "SnapP0      "
    arrOpSwift(0) = recOpSwift
    X = srvOpSwift_Snap(recOpSwift)
    If I > arrOpSwiftNb Then X = "inconnu"
End If

If IsNull(X) Then
    arrOpSwiftIndex = I
    recOpSwift = arrOpSwift(I)
End If

srvOpSwift_Find = X

End Function

'-----------------------------------------------------
Function srvOpSwift_Update(recOpSwift As typeOpSwift)
'-----------------------------------------------------

srvOpSwift_Update = "?"

MsgTxtLen = 0
Call srvOpSwift_PutBuffer(recOpSwift)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    recOpSwift.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
    If Trim(recOpSwift.Err) <> "" Then
        Call srvOpSwift_Error(recOpSwift)
        srvOpSwift_Update = recOpSwift.Err
        Exit Function
    Else
        srvOpSwift_Update = Null
    End If
Else
    recOpSwift.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Public Function srvOpSwift_Monitor(recOpSwift As typeOpSwift)
'-----------------------------------------------------

arrOpSwiftSuite = False
Select Case Mid$(Trim(recOpSwift.Method), 1, 4)
    Case "Seek"
                srvOpSwift_Monitor = srvOpSwift_Seek(recOpSwift)
    Case "Snap", "Prev"
              srvOpSwift_Monitor = srvOpSwift_Snap(recOpSwift)
    Case Else
                recOpSwift.Err = recOpSwift.Method
                Call srvOpSwift_Error(recOpSwift)
                srvOpSwift_Monitor = recOpSwift.Err
End Select

End Function

'-----------------------------------------------------
Sub srvOpSwift_Error(recOpSwift As typeOpSwift)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Référence Swift : " ' & Chr$(10) & Chr$(13)

Select Case Mid$(recOpSwift.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recOpSwift.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvOpSwift.bas  ( " _
                & Trim(recOpSwift.obj) & " : " & Trim(recOpSwift.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvOpSwift_GEtBuffer(recOpSwift As typeOpSwift)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvOpSwift_GEtBuffer = Null
recOpSwift.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recOpSwift.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recOpSwift.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recOpSwift.Err = Space$(10) Then
    recOpSwift.Référence = Mid$(MsgTxt, K + 1, 6)
    recOpSwift.FieldNature = Mid$(MsgTxt, K + 7, 3)
    recOpSwift.FieldLine = Mid$(MsgTxt, K + 10, 1)
    recOpSwift.FieldText = Mid$(MsgTxt, K + 11, 35)

Else
    srvOpSwift_GEtBuffer = recOpSwift.Err
End If

MsgTxtIndex = MsgTxtIndex + recOpSwiftLen

End Function

'---------------------------------------------------------
Private Sub srvOpSwift_PutBuffer(recOpSwift As typeOpSwift)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recOpSwift.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recOpSwift.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 6) = recOpSwift.Référence
Mid$(MsgTxt, K + 7, 3) = recOpSwift.FieldNature
Mid$(MsgTxt, K + 10, 1) = recOpSwift.FieldLine
Mid$(MsgTxt, K + 11, 35) = recOpSwift.FieldText
MsgTxtLen = MsgTxtLen + recOpSwiftLen
End Sub



'---------------------------------------------------------
Private Function srvOpSwift_Seek(recOpSwift As typeOpSwift)
'---------------------------------------------------------

srvOpSwift_Seek = "?"
MsgTxtLen = 0
Call srvOpSwift_PutBuffer(recOpSwift)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvOpSwift_GEtBuffer(recOpSwift)) Then
        srvOpSwift_Seek = Null
    Else
        Call srvOpSwift_Error(recOpSwift)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvOpSwift_Snap(recOpSwift As typeOpSwift)
'---------------------------------------------------------
Dim I As Integer
srvOpSwift_Snap = "?"
MsgTxtLen = 0
Call srvOpSwift_PutBuffer(recOpSwift)
Call srvOpSwift_PutBuffer(arrOpSwift(0))
If IsNull(SndRcv()) Then
    srvOpSwift_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvOpSwift_GEtBuffer(recOpSwift)) Then
            Call arrOpSwift_AddItem(recOpSwift)
            arrOpSwiftSuite = True
        Else
            arrOpSwiftSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recOpSwift_Init(recOpSwift As typeOpSwift)
'---------------------------------------------------------
 MsgTxt = Space$(recOpSwiftLen)
 MsgTxtIndex = 0
 Call srvOpSwift_GEtBuffer(recOpSwift)
 recOpSwift.obj = "SRVOPSWIFT"
End Sub

'---------------------------------------------------------
Public Sub arrOpSwift_AddItem(recOpSwift As typeOpSwift)
'---------------------------------------------------------
          
arrOpSwiftNb = arrOpSwiftNb + 1
    
If arrOpSwiftNb > arrOpSwiftNbMax Then
    arrOpSwiftNbMax = arrOpSwiftNbMax + 10
    ReDim Preserve arrOpSwift(arrOpSwiftNbMax)
End If
            
arrOpSwift(arrOpSwiftNb) = recOpSwift
End Sub

