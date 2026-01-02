Attribute VB_Name = "srvBic"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recBicLen = 205 ' 34 + 171

Type typeBic
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    BicId                   As String * 11
    SwiftId                 As String * 11
    TélexId                 As String * 16
    Intitulé                As String * 40
    Ville                   As String * 30
    Pays                    As String * 30
    
    optAvis                 As String * 1
    optAvisLangue           As String * 1
    VilleId                 As String * 2
    PaysId                  As String * 2
    Banque                  As String * 1
    BdfEtablissement        As String * 5
    BdfGuichet              As String * 5
    AmjMin                  As String * 8
    AmjMax                  As String * 8
End Type
    
Public arrBic() As typeBic
Public arrBicNb As Integer
Public arrBicNbMax As Integer
Public arrBicIndex As Integer
Public arrBicNext As Boolean
'-----------------------------------------------------
Public Function srvBicFind(recBic As typeBic)
'-----------------------------------------------------
Dim I As Integer, X As Variant

X = "?"

For I = 1 To arrBicNb
    If recBic.BicId = arrBic(I).BicId Then
        X = Null
        Exit For
'!!!!!!!========
    End If
Next I

If Not IsNull(X) Then
    I = arrBicNb + 1
    recBic.Method = "SeekP0"
    arrBic(0) = recBic
    X = srvBicSnap(recBic)
    If I > arrBicNb Then X = "inconnu"
End If

If IsNull(X) Then
    arrBicIndex = I
    recBic = arrBic(I)
End If

srvBicFind = X

End Function

'-----------------------------------------------------
Function srvBicUpdate(recBic As typeBic)
'-----------------------------------------------------

srvBicUpdate = "?"

MsgTxtLen = 0
Call srvBicPutBuffer(recBic)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    recBic.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
    If Trim(recBic.Err) <> "" Then
        Call srvBicerror(recBic)
        srvBicUpdate = recBic.Err
        Exit Function
    Else
        srvBicUpdate = Null
    End If
Else
    recBic.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Public Function srvBicMon(recBic As typeBic)
'-----------------------------------------------------

arrBicNext = False
Select Case Mid$(Trim(recBic.Method), 1, 4)
    Case "Seek"
                srvBicMon = srvBicSeek(recBic)
    Case "Snap", "Prev"
              srvBicMon = srvBicSnap(recBic)
    Case Else
                recBic.Err = recBic.Method
                Call srvBicerror(recBic)
                srvBicMon = recBic.Err
End Select

End Function

'-----------------------------------------------------
Sub srvBicerror(recBic As typeBic)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer
If Mid$(recBic.Err, 9, 2) = "23" Then Exit Sub
If Trim(recBic.Err) = "9923" Then Exit Sub
Msg = "Correspondant : " ' & Chr$(10) & Chr$(13)

Select Case Mid$(recBic.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recBic.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvBIC.bas  ( " _
                & Trim(recBic.obj) & " : " & Trim(recBic.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvBicGetBuffer(recBic As typeBic)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvBicGetBuffer = Null
recBic.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recBic.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recBic.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recBic.Err = Space$(10) Then
    recBic.BicId = Mid$(MsgTxt, K + 1, 11)
    recBic.SwiftId = Mid$(MsgTxt, K + 12, 11)
    recBic.TélexId = Mid$(MsgTxt, K + 23, 16)
    recBic.Intitulé = Mid$(MsgTxt, K + 39, 40)
    recBic.Ville = Mid$(MsgTxt, K + 79, 30)
    recBic.Pays = Mid$(MsgTxt, K + 109, 30)
    recBic.optAvis = Mid$(MsgTxt, K + 139, 1)
    recBic.optAvisLangue = Mid$(MsgTxt, K + 140, 1)
    recBic.VilleId = Mid$(MsgTxt, K + 141, 2)
    recBic.PaysId = Mid$(MsgTxt, K + 143, 2)
    recBic.Banque = Mid$(MsgTxt, K + 145, 1)
    recBic.BdfEtablissement = Mid$(MsgTxt, K + 146, 5)
    recBic.BdfGuichet = Mid$(MsgTxt, K + 151, 5)
    recBic.AmjMin = Mid$(MsgTxt, K + 156, 8)
    recBic.AmjMax = Mid$(MsgTxt, K + 164, 8)

Else
    srvBicGetBuffer = recBic.Err
End If

MsgTxtIndex = MsgTxtIndex + recBicLen

End Function

'---------------------------------------------------------
Private Sub srvBicPutBuffer(recBic As typeBic)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recBic.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recBic.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 11) = recBic.BicId
Mid$(MsgTxt, K + 12, 11) = recBic.SwiftId
Mid$(MsgTxt, K + 23, 16) = recBic.TélexId
Mid$(MsgTxt, K + 39, 40) = recBic.Intitulé
Mid$(MsgTxt, K + 79, 30) = recBic.Ville
Mid$(MsgTxt, K + 109, 30) = recBic.Pays
Mid$(MsgTxt, K + 139, 1) = recBic.optAvis
Mid$(MsgTxt, K + 140, 1) = recBic.optAvisLangue
Mid$(MsgTxt, K + 141, 2) = recBic.VilleId
Mid$(MsgTxt, K + 143, 2) = recBic.PaysId
Mid$(MsgTxt, K + 145, 1) = recBic.Banque
Mid$(MsgTxt, K + 146, 5) = recBic.BdfEtablissement
Mid$(MsgTxt, K + 151, 5) = recBic.BdfGuichet
Mid$(MsgTxt, K + 156, 8) = Format$(recBic.AmjMin, "00000000")
Mid$(MsgTxt, K + 164, 8) = Format$(recBic.AmjMax, "00000000")

MsgTxtLen = MsgTxtLen + recBicLen
End Sub



'---------------------------------------------------------
Private Function srvBicSeek(recBic As typeBic)
'---------------------------------------------------------

srvBicSeek = "?"
MsgTxtLen = 0
Call srvBicPutBuffer(recBic)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvBicGetBuffer(recBic)) Then
        srvBicSeek = Null
    Else
        Call srvBicerror(recBic)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvBicSnap(recBic As typeBic)
'---------------------------------------------------------
Dim I As Integer
srvBicSnap = "?"
MsgTxtLen = 0
Call srvBicPutBuffer(recBic)
Call srvBicPutBuffer(arrBic(0))
If IsNull(SndRcv()) Then
    srvBicSnap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvBicGetBuffer(recBic)) Then
            Call arrBicAddItem(recBic)
            arrBicNext = True
        Else
            arrBicNext = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recBicInit(recBic As typeBic)
'---------------------------------------------------------
 MsgTxt = Space$(recBicLen)
 MsgTxtIndex = 0
 Call srvBicGetBuffer(recBic)
 recBic.obj = "SRVBIC"
End Sub

'---------------------------------------------------------
Public Sub arrBicAddItem(recBic As typeBic)
'---------------------------------------------------------
          
arrBicNb = arrBicNb + 1
    
If arrBicNb > arrBicNbMax Then
    arrBicNbMax = arrBicNbMax + 10
    ReDim Preserve arrBic(arrBicNbMax)
End If
            
arrBic(arrBicNb) = recBic
End Sub
'---------------------------------------------------------
Public Sub recBic_Display(XBic As typeBic, XpicBox As PictureBox)
'---------------------------------------------------------

XpicBox.Visible = True
XpicBox.Cls
XpicBox.FontBold = False
XpicBox.ForeColor = txtUsr.ForeColor
XpicBox.BackColor = vbWindowBackground
XpicBox.CurrentY = 0
XpicBox.CurrentX = 50: XpicBox.Print "Code";
XpicBox.CurrentX = 600: XpicBox.Print ":";
XpicBox.CurrentY = 200
XpicBox.CurrentX = 50: XpicBox.Print "Swift";
XpicBox.CurrentX = 600: XpicBox.Print ":";
XpicBox.CurrentY = 400
XpicBox.CurrentX = 50: XpicBox.Print "Télex";
XpicBox.CurrentX = 600: XpicBox.Print ":";

XpicBox.ForeColor = libUsr.ForeColor

XpicBox.CurrentY = 0
XpicBox.CurrentX = 700: XpicBox.Print XBic.BicId;
XpicBox.FontBold = True
XpicBox.CurrentX = 2500: XpicBox.Print XBic.Intitulé;
XpicBox.FontBold = False

XpicBox.CurrentY = 200
XpicBox.CurrentX = 700: XpicBox.Print XBic.SwiftId;
XpicBox.CurrentX = 2500: XpicBox.Print XBic.Ville;

XpicBox.CurrentY = 400
XpicBox.CurrentX = 700: XpicBox.Print XBic.TélexId;
XpicBox.CurrentX = 2500: XpicBox.Print XBic.Pays;

If XBic.AmjMin <> "00000000" _
Or XBic.AmjMax <> "00000000" Then
    XpicBox.ForeColor = txtUsr.ForeColor
    XpicBox.CurrentX = 6300: XpicBox.Print "du :";
    XpicBox.CurrentX = 7800: XpicBox.Print "au :";
    XpicBox.ForeColor = errUsr.ForeColor
    XpicBox.CurrentX = 6600: XpicBox.Print dateImp(XBic.AmjMin);
    XpicBox.CurrentX = 8100: XpicBox.Print dateImp(XBic.AmjMax);
End If

End Sub





Public Function srvBic_Control(recBic As typeBic, ByVal Amj As String)
recBic.Method = "SeekP0"
srvBic_Control = Null
If Not IsNull(srvBicFind(recBic)) Then
    srvBic_Control = "? Correspondant inconnu"
Else
    If Amj < recBic.AmjMin Then
        srvBic_Control = "? Corr : validité min " & recBic.AmjMin
    Else
        If Val(recBic.AmjMax) > 0 _
        And Amj > recBic.AmjMax Then
            srvBic_Control = "? Corr : validité max " & recBic.AmjMax
         End If
    End If
End If
End Function
