Attribute VB_Name = "prtLucaRisques_Sopra"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim Colonne As Integer
Dim Height6_5 As Integer
Dim NbLg As Integer
Dim blnTotalCheck As Boolean, blnPage As Boolean
Dim Filename As String

'---------------------------------------------------------
 Public Sub prtLucaRisques_SopraX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String, X2 As String
On Error Resume Next

NbLg = 0
Set XPrt = Printer
frmElpPrt.Show vbModeless

prtOrientation = vbPRORLandscape
prtTitleText = "Centralisation des risques bancaires "
prtPgmName = "prtLucaRisques_Sopra"
prtTitleUsr = usrName
prtFontName = "Courier New"

prtLineNb = 1
prtlineHeight = 180
prtHeaderHeight = 0
Filename = Trim(Msg)

Select Case Filename
    Case Is = paramLrCdr_LrBdfAller_FileName
        prtTitleText = "Centralisation des risques bancaires : Bande Aller "
        frmElpPrt.prtStdInit
        prtFontSize = 6
     Case Is = paramLrCdr_PrintSopra_400
        prtTitleText = "Centralisation des risques bancaires : Fiches signalétiques erronées "
        frmElpPrt.prtStdInit
        prtFontSize = 9
   Case Is = paramLrCdr_PrintSopra_470
        prtFormType = ""
        frmElpPrt.prtInit
        prtFontSize = 9
    Case Else
        frmElpPrt.prtStdBottomInit
        prtFontSize = 9
End Select

XPrt.FontSize = prtFontSize
XPrt.CurrentY = prtMinY
blnTotalCheck = False
blnPage = False
If Filename = paramLrCdr_PrintSopra_490 _
Or Filename = paramLrCdr_PrintSopra_220 Then blnTotalCheck = True

Open Filename For Input As #1

Do Until EOF(1)
    Line Input #1, X
     If mId$(X, 1, 1) = Chr$(10) Then
        X = mId$(X, 2, Len(X))
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    End If
   If mId$(X, 1, 1) = Chr$(12) Then
        If blnPage Then frmElpPrt.prtNewPage
        X = mId$(X, 2, Len(X))
        XPrt.CurrentY = prtMinY - prtlineHeight
        XPrt.FontSize = prtFontSize
    End If
    If XPrt.CurrentY > prtMaxY Then
        frmElpPrt.prtNewPage
        XPrt.CurrentY = prtMinY
    End If
  
   XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
   
   XPrt.FontBold = False
   XPrt.CurrentX = prtMinX
   If blnTotalCheck Then
        If mId$(X, 86, 14) = "---  TOTAL ---" Then
            XPrt.Print mId$(X, 1, 102);
            XPrt.FontBold = True
            XPrt.Print mId$(X, 103, 22);
            X = mId$(X, 125, Len(X))
            XPrt.FontBold = False
        End If
        If IsNumeric(mId$(X, 2, 5)) Then
            Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight, " ")
            XPrt.CurrentX = prtMinX
            'XPrt.Print Mid$(X, 1, 64);
            'XPrt.FontBold = True
            'XPrt.Print Mid$(X, 65, 60);
            'X = Mid$(X, 125, Len(X))
            'XPrt.FontBold = False
        End If
    End If
   blnPage = True
   XPrt.Print X;
Loop
Close #1


frmElpPrt.prtEndDoc
frmElpPrt.Hide
prtFontName = prtFontNameZ
End Sub

