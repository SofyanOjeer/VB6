Attribute VB_Name = "prtEdition"
Option Explicit

Public Sub prtCourrier_Open()
Dim mCurrentX As Integer, mCurrenty As Integer
Dim K As Integer
Dim blnAVI002P1 As Boolean
Dim wPrinter_Avis As String
wPrinter_Avis = "IMP_AVIS"
'wPrinter_Avis = "JPL_AVIS": MsgBox wPrinter_Avis
On Error Resume Next 'GoTo prtError
If prtPaperSize = vbPRPSA5 Then
'    Printer_Set (wPrinter_Avis): frmElpPrt.prtColor_Check: blnAVI002P1 = True
    If InStr(1, UCase$(Trim(Printer.Devicename)), "IMP_GDMP") Then Printer_Set (wPrinter_Avis): frmElpPrt.prtColor_Check
    If InStr(1, UCase$(Trim(Printer.Devicename)), "MFP_GDMP_HP") Then Printer_Set (wPrinter_Avis): frmElpPrt.prtColor_Check
'    If InStr(1, UCase$(Trim(Printer.Devicename)), "IMP_GUICHET") Then Printer_Set (wPrinter_Avis): frmElpPrt.prtColor_Check
'    If InStr(1, UCase$(Trim(Printer.Devicename)), "IMP_CAISSE") Then Printer_Set (wPrinter_Avis): frmElpPrt.prtColor_Check
'    If InStr(1, UCase$(Trim(Printer.Devicename)), "IMP_SOBF") Then Printer_Set (wPrinter_Avis): frmElpPrt.prtColor_Check
'    If InStr(1, UCase$(Trim(Printer.Devicename)), "IMP_ORPA") Then Printer_Set (wPrinter_Avis): frmElpPrt.prtColor_Check
    If InStr(1, UCase$(Trim(Printer.Devicename)), "IMP_AVIS") Then blnAVI002P1 = True
    
End If

Set XPrt = Printer
'frmElpPrt.Show vbModeless
'$jpl 20080304 prtOrientation = vbPRORPortrait
If prtPaperSize = vbPRPSA5 Then
    If blnAVI002P1 Then
        XPrt.PaperSize = prtPaperSize
        prtOrientation = vbPRORLandscape
    Else
        XPrt.PaperSize = vbPRPSA4
        prtOrientation = vbPRORPortrait
    End If
    
End If
prtTitleText = "Courrier"
prtPgmName = "prtCourrier"
prtTitleUsr = usrName
prtFontName = prtFontName_Arial 'prtFontNameZ

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = prtMinY + 1000

prtFormType = ""
prtSocInit
prtlineHeight = prtlineHeight66

'$jpl 20040420  ' blnFiligrane = False
If blnFiligrane Then
    If Mid$(prtFiligrane_Name, 1, 1) = "\" Then
        frmElpPrt.prtFiligrane prtFiligrane_Name
    Else
        mCurrentX = XPrt.CurrentX
        mCurrenty = XPrt.CurrentY
        'XPrt.CurrentY = 1200
        'Call frmElpPrt.prtTrame(prtMinMarge, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + 350, , 235)
        'XPrt.FontSize = 16: XPrt.FontBold = True
        'frmElpPrt.prtCentré prtMedX, prtFiligrane_Name
        XPrt.CurrentY = prtMinY
        Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMinX + 3500, XPrt.CurrentY + 350, " ", 255)
        XPrt.FontSize = 16: XPrt.FontBold = True
        XPrt.CurrentX = prtMinX
        XPrt.ForeColor = prtFiligrane_Color
        XPrt.Print prtFiligrane_Name;
        XPrt.ForeColor = prtForeColor
        XPrt.CurrentX = mCurrentX
        XPrt.CurrentY = mCurrenty
        XPrt.FontBold = False
    End If
End If
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------
On Error Resume Next

If currentUser.QSYSOPR <> "1" Then Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
End Sub

Public Sub prtEdition_Open()
On Error GoTo prtError

Set XPrt = Printer
'frmElpPrt.Show vbModeless
'' à définir avant appel
'''prtOrientation = vbPRORPortrait
''''prtFontName = prtFontNameZ

prtTitleText = "Edition"
prtPgmName = "prtEdition"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 0

frmElpPrt.prtSAB_Init
prtlineHeight = prtlineHeight66

'$JPL.20040420 blnFiligrane = False  '$JPL.20030305

If blnFiligrane Then frmElpPrt.prtFiligrane prtFiligrane_Name

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------
On Error Resume Next
If currentUser.QSYSOPR <> "1" Then Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
End Sub


Public Sub prtEdition_Close()
On Error Resume Next
Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
End Sub



