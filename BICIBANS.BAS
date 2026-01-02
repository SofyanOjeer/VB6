Attribute VB_Name = "srvBicIban"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recBicIbanLen = 102 ' 34 + 68

Type typeBicIban
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    BicId                   As String * 11
    IbanDevise              As String * 3
    IbanBicId               As String * 11
    IbanPays                As String * 2
    IbanClé                 As String * 2
    IbanCompte              As String * 23
    
    AmjMin                  As String * 8
    AmjMax                  As String * 8
End Type
    
Public arrBicIban() As typeBicIban
Public arrBicIbanNb As Integer
Public arrBicIbanNbMax As Integer
Public arrBicIbanIndex As Integer
Public arrBicIbanNext As Boolean
'-----------------------------------------------------
Public Function srvBicIbanFind(recBicIban As typeBicIban)
'-----------------------------------------------------
Dim I As Integer, X As Variant

X = "?"

For I = 1 To arrBicIbanNb
'    And recBicIban.IbanBicId = arrBicIban(I).IbanBicId '_
    If recBicIban.BicId = arrBicIban(I).BicId _
    And recBicIban.IbanDevise = arrBicIban(I).IbanDevise _
    And recBicIban.IbanBicId = arrBicIban(I).IbanBicId Then
        X = Null
        Exit For
'!!!!!!!========
    End If
Next I

If Not IsNull(X) Then
    I = arrBicIbanNb + 1
    recBicIban.Method = "SnapP0"
    arrBicIban(0) = recBicIban
    If Trim(recBicIban.IbanBicId) = "" Then arrBicIban(0).IbanBicId = "99999999999"
    X = srvBicIbanSnap(recBicIban)
    If I > arrBicIbanNb Then X = "inconnu"
End If

If IsNull(X) Then
    arrBicIbanIndex = I
    recBicIban = arrBicIban(I)
End If

srvBicIbanFind = X

End Function


'-----------------------------------------------------
Function srvBicIbanUpdate(recBicIban As typeBicIban)
'-----------------------------------------------------

srvBicIbanUpdate = "?"

MsgTxtLen = 0
Call srvBicIbanPutBuffer(recBicIban)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    recBicIban.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
    If Trim(recBicIban.Err) <> "" Then
        Call srvBicIbanerror(recBicIban)
        srvBicIbanUpdate = recBicIban.Err
        Exit Function
    Else
        srvBicIbanUpdate = Null
    End If
Else
    recBicIban.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Public Function srvBicIbanMon(recBicIban As typeBicIban)
'-----------------------------------------------------

arrBicIbanNext = False
Select Case mId$(Trim(recBicIban.Method), 1, 4)
    Case "Seek"
                srvBicIbanMon = srvBicIbanSeek(recBicIban)
    Case "Snap", "Prev"
              srvBicIbanMon = srvBicIbanSnap(recBicIban)
    Case Else
                recBicIban.Err = recBicIban.Method
                Call srvBicIbanerror(recBicIban)
                srvBicIbanMon = recBicIban.Err
End Select

End Function

'-----------------------------------------------------
Sub srvBicIbanerror(recBicIban As typeBicIban)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Correspondant : références bancaires" ' & Chr$(10) & Chr$(13)

Select Case mId$(recBicIban.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recBicIban.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvBicIban.bas  ( " _
                & Trim(recBicIban.obj) & " : " & Trim(recBicIban.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvBicIbanGetBuffer(recBicIban As typeBicIban)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvBicIbanGetBuffer = Null
recBicIban.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recBicIban.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recBicIban.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recBicIban.Err = Space$(10) Then
    recBicIban.BicId = mId$(MsgTxt, K + 1, 11)
    recBicIban.IbanDevise = mId$(MsgTxt, K + 12, 3)
    recBicIban.IbanBicId = mId$(MsgTxt, K + 15, 11)
    recBicIban.IbanPays = mId$(MsgTxt, K + 26, 2)
    recBicIban.IbanClé = mId$(MsgTxt, K + 28, 2)
    recBicIban.IbanCompte = mId$(MsgTxt, K + 30, 23)
    recBicIban.AmjMin = mId$(MsgTxt, K + 53, 8)
    recBicIban.AmjMax = mId$(MsgTxt, K + 61, 8)

Else
    srvBicIbanGetBuffer = recBicIban.Err
End If

MsgTxtIndex = MsgTxtIndex + recBicIbanLen

End Function

'---------------------------------------------------------
Private Sub srvBicIbanPutBuffer(recBicIban As typeBicIban)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recBicIban.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recBicIban.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 11) = recBicIban.BicId
Mid$(MsgTxt, K + 12, 3) = recBicIban.IbanDevise
Mid$(MsgTxt, K + 15, 11) = recBicIban.IbanBicId
Mid$(MsgTxt, K + 26, 2) = recBicIban.IbanPays
Mid$(MsgTxt, K + 28, 2) = recBicIban.IbanClé
Mid$(MsgTxt, K + 30, 23) = recBicIban.IbanCompte
Mid$(MsgTxt, K + 53, 8) = Format$(recBicIban.AmjMin, "00000000")
Mid$(MsgTxt, K + 61, 8) = Format$(recBicIban.AmjMax, "00000000")

MsgTxtLen = MsgTxtLen + recBicIbanLen
End Sub



'---------------------------------------------------------
Private Function srvBicIbanSeek(recBicIban As typeBicIban)
'---------------------------------------------------------

srvBicIbanSeek = "?"
MsgTxtLen = 0
Call srvBicIbanPutBuffer(recBicIban)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvBicIbanGetBuffer(recBicIban)) Then
        srvBicIbanSeek = Null
    Else
        Call srvBicIbanerror(recBicIban)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvBicIbanSnap(recBicIban As typeBicIban)
'---------------------------------------------------------
Dim I As Integer
srvBicIbanSnap = "?"
MsgTxtLen = 0
Call srvBicIbanPutBuffer(recBicIban)
Call srvBicIbanPutBuffer(arrBicIban(0))
If IsNull(SndRcv()) Then
    srvBicIbanSnap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvBicIbanGetBuffer(recBicIban)) Then
            Call arrBicIbanAddItem(recBicIban)
            arrBicIbanNext = True
        Else
            arrBicIbanNext = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recBicIbanInit(recBicIban As typeBicIban)
'---------------------------------------------------------
 MsgTxt = Space$(recBicIbanLen)
 MsgTxtIndex = 0
 Call srvBicIbanGetBuffer(recBicIban)
 recBicIban.obj = "SRVBICIBAN"
End Sub

'---------------------------------------------------------
Public Sub arrBicIbanAddItem(recBicIban As typeBicIban)
'---------------------------------------------------------
          
arrBicIbanNb = arrBicIbanNb + 1
    
If arrBicIbanNb > arrBicIbanNbMax Then
    arrBicIbanNbMax = arrBicIbanNbMax + 10
    ReDim Preserve arrBicIban(arrBicIbanNbMax)
End If
            
arrBicIban(arrBicIbanNb) = recBicIban
End Sub
'---------------------------------------------------------
Public Sub picBicIban_Print(XBic As typeBicIban, XpicBox As PictureBox)
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
XpicBox.CurrentX = 50: XpicBox.Print "auprès de";
XpicBox.CurrentX = 600: XpicBox.Print ":";
XpicBox.CurrentY = 400
XpicBox.CurrentX = 50: XpicBox.Print "Compte";
XpicBox.CurrentX = 600: XpicBox.Print ":";

XpicBox.ForeColor = libUsr.ForeColor

XpicBox.CurrentY = 0
XpicBox.CurrentX = 700: XpicBox.Print XBic.BicId;
'XpicBox.FontBold = True
'XpicBox.CurrentX = 2500: XpicBox.Print XBic.ibanpays;
'XpicBox.FontBold = False

XpicBox.CurrentY = 200
XpicBox.CurrentX = 700: XpicBox.Print XBic.IbanBicId;
'XpicBox.CurrentX = 2500: XpicBox.Print XBic.ibanclé;

XpicBox.CurrentY = 400
'XpicBox.CurrentX = 700: XpicBox.Print XBic.Devise;
XpicBox.CurrentX = 2500
XpicBox.Print XBic.IbanPays & XBic.IbanClé & " " & XBic.IbanCompte;

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






