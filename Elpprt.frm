VERSION 5.00
Begin VB.Form frmElpPrt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Impression "
   ClientHeight    =   1875
   ClientLeft      =   2100
   ClientTop       =   2580
   ClientWidth     =   3660
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Elpprt.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1875
   ScaleWidth      =   3660
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   360
   End
   Begin VB.CommandButton cmd_Suite 
      Caption         =   "&Suite"
      Default         =   -1  'True
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_Annuler 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "&Annuler"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1245
   End
   Begin VB.Image imgSocLogo_PiedPage 
      Height          =   1000
      Left            =   0
      Top             =   1500
      Width           =   2500
   End
   Begin VB.Image imgFiligrane 
      Height          =   495
      Left            =   1905
      Top             =   540
      Width           =   750
   End
   Begin VB.Image imgClip 
      Height          =   495
      Left            =   1875
      Top             =   -15
      Width           =   750
   End
   Begin VB.Image imgSocLogo 
      Height          =   975
      Left            =   0
      Top             =   -15
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Image imgSocLogo_G 
      Height          =   525
      Left            =   0
      Top             =   990
      Width           =   2655
   End
End
Attribute VB_Name = "frmElpPrt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit

Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim indexTimer As Integer
Dim Shell_SendKeys As String
Dim IdShell

Dim prtRTF_FontName(30) As String
Dim blnRTF_Center As Boolean
Dim prtRTF_MinX As Integer
Dim prtRTF_TrameX1 As Integer, prtRTF_TrameY1 As Integer
Dim prtRTF_TrameX2 As Integer, prtRTF_TrameY2 As Integer
Dim blnRTF_Trame As Boolean
Dim arrRTF_FontColor(100) As Long, arrRTF_FontColor_Nb As Integer


Public Sub prtColor_Check()
Dim K As Integer
'$JPL_20060829     K = InStr(1, Printer.Devicename, "COLOR")
'$JPL_20060829     If K = 0 Then K = InStr(1, Printer.Devicename, "PDF")
 '$JPL_20060829    If K = 0 Then
 


 If InStr(1, Printer.Devicename, "PDF") > 0 Then
     blnIMP_PDF = True
     frmElpPrt.prtIMP_PDF_Monitor "Clear"
 Else
     blnIMP_PDF = False
 End If
 
 If Printer.ColorMode <> 2 Then
        prtColor_Check_1
   Else
        prtColor_Check_2
    End If

End Sub



Public Sub prtIMP_PDF_NoPaper_Mail_RELEVE_FOTC(lFrom As String, lRecipient As String, lMsg As String, lFileName As String)
On Error GoTo Error_Handler
Dim wSendMail As typeSendMail, xLib As String, xMemo As String, V
    
If lFrom = "" Then
    wSendMail.From = currentSSIWINMAIL
    If lRecipient = "" Then
        wSendMail.Recipient = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire(paramEditionNoPaper_Auto_Unit)
    Else
        wSendMail.Recipient = lRecipient
    End If
    
Else
    wSendMail.Recipient = srvSendMail.Exchange_Distribution(lFrom, lRecipient)
End If

wSendMail.FromDisplayName = "NoPaper " & paramEditionNoPaper_Auto_PgmName
wSendMail.CcRecipient = ""

wSendMail.Subject = lMsg
wSendMail.Attachment = lFileName
wSendMail.Message = mHtml_Head & "<span style='font-size:10.0pt;font-family:Calibri'>" _
                 & htmlFontColor_Black & lMsg & "<BR><BR>" & paramEditionNoPaper_Auto_Lnk & "</div></body></html>"

 wSendMail.AsHTML = True
 srvSendMail.Monitor wSendMail
 
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, "prtIMP_PDF_NoPaper_Mail_RELEVE_FOTC"

End Sub

'---------------------------------------------------------
'---------------------------------------------------------
Public Sub prtTrame(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, Optional prtBox, Optional prtRGB)
'---------------------------------------------------------
Dim mCurrenty As Integer, mFillStyle As Integer, mFillColor
Dim xBox As String
mCurrenty = XPrt.CurrentY
mFillStyle = XPrt.FillStyle
mFillColor = XPrt.FillColor
XPrt.FontTransparent = True

XPrt.FillStyle = 0
XPrt.ForeColor = RGB(255, 255, 255)
'If Not IsMissing(prtBox) Then
'    If prtBox = "B" Then
'        XPrt.ForeColor = RGB(0, 0, 0)
'    End If
'End If
If prtColorMode Then
    XPrt.FillColor = prtFillColor
    If IsMissing(prtBox) Or prtBox <> "B" Then
        XPrt.Line (X1, Y1)-(X2, Y2), vbWhite, B
    Else
        XPrt.Line (X1, Y1)-(X2, Y2), prtLineColor, B
    End If
Else
    If Not IsMissing(prtRGB) And IsNumeric(prtRGB) Then
        XPrt.FillColor = RGB(prtRGB, prtRGB, prtRGB)
    Else
        XPrt.FillColor = RGB(245, 245, 245)
    End If
    If IsMissing(prtBox) Then
        XPrt.ForeColor = RGB(255, 255, 255)
        XPrt.Line (X1, Y1)-(X2, Y2), , B
    Else
        If prtBox = "B" Then
            XPrt.ForeColor = prtLineColor
            XPrt.Line (X1, Y1)-(X2, Y2), , B
        Else
            XPrt.ForeColor = RGB(255, 255, 255)
            XPrt.Line (X1, Y1)-(X2, Y2), , B
        End If
    End If
End If

XPrt.ForeColor = prtForeColor
XPrt.FillStyle = mFillStyle
XPrt.FillColor = mFillColor
XPrt.CurrentY = mCurrenty

End Sub

'---------------------------------------------------------
Public Sub prtTrame_Color(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, Optional prtBox)
'---------------------------------------------------------
Dim mCurrenty As Integer, mFillStyle As Integer, mFillColor
Dim xBox As String
mCurrenty = XPrt.CurrentY
mFillStyle = XPrt.FillStyle
mFillColor = XPrt.FillColor
XPrt.FontTransparent = True

XPrt.FillStyle = 0
    XPrt.FillColor = prtFillColor
    If IsMissing(prtBox) Or prtBox <> "B" Then
        XPrt.Line (X1, Y1)-(X2, Y2), vbWhite, B
    Else
        XPrt.Line (X1, Y1)-(X2, Y2), prtLineColor, B
    End If
XPrt.FillStyle = mFillStyle
XPrt.FillColor = mFillColor
XPrt.CurrentY = mCurrenty

End Sub

'---------------------------------------------------------
Public Sub prtNewPage()
'---------------------------------------------------------
Dim mFontSize As Integer
On Error Resume Next
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
If TypeOf XPrt Is Printer Then
    mFontSize = XPrt.FontSize
    XPrt.FontName = prtFontNameZ
    XPrt.NewPage
    prtFormType_Select
    XPrt.FontName = prtFontName
    XPrt.FontSize = mFontSize
    XPrt.FontTransparent = True
Else
   MsgBox "suite"
   XPrt.cmd_Suite.Visible = True
    XPrt.Cls
 End If
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
End Sub

'---------------------------------------------------------
Public Sub prtStdInit()
'---------------------------------------------------------
prtFormType = "STD"
frmElpPrt.prtInit
frmElpPrt.prtFormType_Select
End Sub
'---------------------------------------------------------
Public Sub prtStdBlankInit()
'---------------------------------------------------------
prtFormType = "   "
frmElpPrt.prtInit
frmElpPrt.prtFormType_Select
End Sub

Public Sub prtSAB_Init()
'---------------------------------------------------------
prtFormType = "SAB"
frmElpPrt.prtInit
frmElpPrt.prtFormType_Select
End Sub

'---------------------------------------------------------
Public Sub prtStdTopInit()
'---------------------------------------------------------
prtFormType = "TOP"
frmElpPrt.prtInit
frmElpPrt.prtFormType_Select
End Sub

'---------------------------------------------------------
Public Sub prtStdBottomInit()
'---------------------------------------------------------
prtFormType = "BOT"
frmElpPrt.prtInit
frmElpPrt.prtFormType_Select
End Sub

'---------------------------------------------------------
Public Sub prtStdBottom()
'---------------------------------------------------------
Dim W As Integer, H As Integer
Dim SW As Single, SH As Single, SX As Single
Dim L As Integer
On Error Resume Next
XPrt.FontBold = False
XPrt.ForeColor = prtForeColor_Header

'If prtSocSigle Then
'    SH = 300 / frmElpPrt.imgSocSigle.Height
'    W = SH * frmElpPrt.imgSocSigle.Width
'    H = SH * frmElpPrt.imgSocSigle.Height

'    XPrt.PaintPicture frmElpPrt.imgSocSigle.Picture _
'                    , prtMinX + (prtMaxX - prtMinX - W) / 2 _
'                    , prtMaxY + 20 _
'                    , W, H
'End If


XPrt.FontSize = 6

XPrt.CurrentY = prtMaxY + (300 - XPrt.TextHeight(prtPgmName)) / 2
XPrt.CurrentX = prtMinX
XPrt.Print prtPgmName; Space$(5); Now; Space$(5); usrId;

If TypeOf XPrt Is Printer Then
    XPrt.CurrentX = prtMaxX - 200
    XPrt.Print XPrt.Page
End If
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
XPrt.DrawWidth = 3
XPrt.Line (prtMinX, prtMaxY)-(prtMaxX, prtMaxY), prtLineColor, B

XPrt.CurrentY = prtMinY
XPrt.CurrentX = prtMinX
XPrt.FontBold = False
XPrt.DrawWidth = 1
XPrt.ForeColor = prtForeColor

End Sub


'---------------------------------------------------------
Public Sub prtSAB_Bottom()
'---------------------------------------------------------
Dim W As Integer, H As Integer
Dim SW As Single, SH As Single, SX As Single
Dim L As Integer
On Error Resume Next
XPrt.FontBold = False

XPrt.FontSize = 6

XPrt.CurrentY = prtMaxY + (300 - XPrt.TextHeight(prtPgmName)) / 2
XPrt.CurrentX = prtMinX
XPrt.Print prtPgmName; Space$(5); Now; Space$(5); frmRTF_UsrId_Origine;
XPrt.CurrentX = XPrt.CurrentX + 1000
XPrt.FontBold = True
XPrt.Print frmRTF_Référence;

If TypeOf XPrt Is Printer Then
    XPrt.CurrentX = prtMaxX - 200
    XPrt.Print XPrt.Page
End If
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
XPrt.DrawWidth = 3
XPrt.Line (prtMinX, prtMaxY)-(prtMaxX, prtMaxY), prtLineColor, B
XPrt.DrawWidth = 1

XPrt.CurrentY = prtMinY
XPrt.CurrentX = prtMinX
XPrt.FontBold = False

End Sub


'---------------------------------------------------------
Public Sub prtEndDoc(dureePose As Long)
'---------------------------------------------------------

'__________________________________________________________
On Error Resume Next 'GoTo EndDocError

If TypeOf XPrt Is Printer Then
    If Not prtKillDoc Then
        XPrt.EndDoc
        Call Sleep(dureePose)
        'On Error GoTo 0
        
        Set XPrt = Nothing
        If blnIMP_PDF Then prtIMP_PDF_Monitor "Close"

    End If
End If

Exit Sub
'__________________________________________________________

'EndDocError:
'       If Err.Number = 482 Then
'           'do nothing
'       Else
'          ' Err.Raise Err.Number
'       End If

End Sub


'---------------------------------------------------------
Public Sub prtScreen()
'---------------------------------------------------------
Dim W As Integer, H As Integer
Dim SW As Single, SH As Single, SX As Single
On Error GoTo prtError


Set XPrt = Printer
prtOrientation = vbPRORLandscape
prtPgmName = "Elp.prtScreen"
prtTitleText = " "
prtTitleUsr = usrName

frmElpPrt.prtStdInit
frmElpPrt.imgClip = Clipboard.GetData

frmElpPrt.prtImage frmElpPrt.imgClip

Call frmElpPrt.prtEndDoc(1000)
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

'---------------------------------------------------------
Public Sub prtImage(lImage As IMAGE)
'---------------------------------------------------------
Dim mCurrenty As Integer, mCurrentX As Integer, mFillStyle As Integer, mFillColor
Dim W As Integer, H As Integer
Dim SW As Single, SH As Single, SX As Single
On Error Resume Next 'GoTo prtError

mCurrentX = XPrt.CurrentX
mCurrenty = XPrt.CurrentY
mFillStyle = XPrt.FillStyle
mFillColor = XPrt.FillColor
XPrt.FontTransparent = True

XPrt.FillStyle = 0
'XPrt.ForeColor = RGB(255, 255, 255)



SW = (prtMaxX - prtMinX) / lImage.Width
SH = (prtMaxY - prtMinY) / lImage.Height
If SH < SW Then
    SX = SH
Else
    SX = SW
End If
If SX > 1 Then '1.1 Then
    SX = 1     '1.1
End If
W = SX * lImage.Width
H = SX * lImage.Height

XPrt.PaintPicture lImage.Picture _
                , prtMinX + (prtMaxX - prtMinX - W) / 2 _
                , prtMinY + (prtMaxY - prtMinY - H) / 2 _
                , W, H
'XPrt.ForeColor = RGB(0, 0, 0)
XPrt.FillStyle = mFillStyle
XPrt.FillColor = mFillColor
XPrt.CurrentX = mCurrentX
XPrt.CurrentY = mCurrenty
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "prtImgClip")
frmElpPrt.Hide

End Sub
Public Sub prtFiligrane(lFileName As String)
Dim X As String
X = Trim(lFileName)
If X <> "" Then
    If frmElpPrt.imgFiligrane.Tag <> X Then
    
        If Dir(X) = "" Then Exit Sub
        
        frmElpPrt.imgFiligrane.Tag = X
        frmElpPrt.imgFiligrane.Picture = LoadPicture(X)
    End If
    frmElpPrt.prtImage frmElpPrt.imgFiligrane
End If
End Sub



'---------------------------------------------------------
Private Sub Cmd_Annuler_Click()
'---------------------------------------------------------
prtKillDoc = True
If TypeOf XPrt Is Printer Then
    XPrt.KillDoc
End If

End Sub

'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
prtKillDoc = False
cmd_Suite.Visible = False
End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 27: Unload Me
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub


'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
KeyPreview = True
Set XForm = Me
'MeInit

End Sub



'---------------------------------------------------------
Public Sub prtInit()
'---------------------------------------------------------
Dim H1 As Integer, H2 As Integer
On Error Resume Next

XPrt.FontTransparent = False 'True
 XPrt.FontBold = False
XPrt.FontName = prtFontName
prtKillDoc = False

prtMinX = 200: prtMinY = 300 '200
If prtOrientation = vbPRORPortrait Then
'    prtMaxY = 15800: prtMaxX = 11100
    prtMaxY = 15600: prtMaxX = 11100
Else
    prtMaxY = 11000: prtMaxX = 16000
End If
If XPrt.PaperSize = vbPRPSA5 Then
'    prtMaxY = 7900: prtMaxX = 11100
    prtMaxY = 7700: prtMaxX = 11100
End If

prtMinMarge = prtMinX + 500
prtMaxMarge = prtMaxX - 500
prtWidthMarge = prtMaxMarge - prtMinMarge

prtMedX = prtMinMarge + (prtMaxMarge - prtMinMarge) / 2
prtMedX0 = prtMedX
prtMinX1 = prtMinX
prtMinX2 = prtMedX

prtMedY = prtMinY + (prtMaxY - prtMinY) / 2
blnMinX = False

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
If TypeOf XPrt Is Printer Then
'   XPrt.Zoom = prtZoom
    XPrt.Orientation = prtOrientation
Else
    prtMaxY = 7900
    XPrt.WindowState = vbMaximized
End If
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
prtlineHeight66 = (prtMaxY - prtMinY) / 66 - 5

If prtLineNb > 0 And prtlineHeight > 0 Then
    H1 = prtMaxY - prtMinY - prtHeaderHeight
    H2 = H1 \ (prtlineHeight * prtLineNb)
    prtParagraphHeight = H1 \ H2 + 20
    prtlineHeight = (prtParagraphHeight - 20) \ prtLineNb
End If

'____________________________________________________
' ces deux instructions doivent être en fin de procèdure
'___________________________________________________________
'XPrt.ColorMode = 2
XPrt.ForeColor = prtForeColor

If blnIMP_PDF Then prtIMP_PDF_Monitor "Open"
End Sub

'---------------------------------------------------------
Public Sub prtLineY()
'---------------------------------------------------------

XPrt.DrawWidth = 2
XPrt.Line (prtMinX, prtCurrentY)-(prtMaxX, prtCurrentY), prtLineColor, B

End Sub


'---------------------------------------------------------
Public Sub prtLineX(lX As Integer)
'---------------------------------------------------------

XPrt.DrawWidth = 2
XPrt.Line (lX, prtMinY + prtHeaderHeight)-(lX, prtMaxY), prtLineColor, B

End Sub




Public Sub prtCentré(intX As Integer, strX As String)
XPrt.CurrentX = intX - XPrt.TextWidth(strX) / 2
XPrt.Print strX;
End Sub

Public Sub prtTiret()
Dim I As Integer
XPrt.CurrentX = 0
For I = 1 To 4
       XPrt.Print "....................................................................................";
Next I

End Sub

Public Sub WinWord(docName As String)
Dim X As String
On Error GoTo Error_Handler
If Dir(constWinWord, vbReadOnly + vbHidden) = "" Then
    constWinWord = constWinWord_D
    If Dir(constWinWord, vbReadOnly + vbHidden) = "" Then
        MsgBox constWinWord, vbCritical, "frmElpPrt.bas : programme absent": Exit Sub
    End If
End If
X = constWinWord & " " & Chr$(34) & docName & Chr$(34)
IdShell = Shell(X, 1)
AppActivate IdShell
DoEvents
Exit Sub

Error_Handler:
'    MsgBox Error, vbCritical, "WinWord : ", X

End Sub
Public Function Windows_Display_File(lFileName As String)
On Error Resume Next
Dim wExtension As String

wExtension = UCase$(fileName_Extension(Trim(lFileName)))
    If Dir(lFileName) <> "" Then
        Select Case wExtension
         Case "DOC": Call frmElpPrt.WinWord(lFileName)
         Case "XLS": Call frmElpPrt.Excel(lFileName)
         Case "PDF": Call frmElpPrt.Acrord32(lFileName)
         Case "TXT": Call frmElpPrt.WordPad(lFileName) 'NotePad(lFileName)
         Case "RTF": Call frmElpPrt.WordPad(lFileName)
         Case Else: Call frmElpPrt.IExplore(lFileName)
        End Select
        Wait_SS 1
        DoEvents
    End If
End Function

Public Sub IExplore(docName As String)
Dim X As String
On Error GoTo Error_Handler
X = constIExplorer & " " & Chr$(34) & docName & Chr$(34)
IdShell = Shell(X, 1)
On Error Resume Next
'AppActivate IdShell
DoEvents
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "IExplore : ", X

End Sub

Public Sub WordPad(docName As String)
Dim X As String
On Error GoTo Error_Handler
If Dir(constWordPad, vbReadOnly + vbHidden) = "" Then
    MsgBox constWordPad, vbCritical, "frmElpPrt.bas : programme absent"
Else
    X = constWordPad & " " & Chr$(34) & docName & Chr$(34)
    IdShell = Shell(X, 1)
    AppActivate IdShell
    DoEvents
End If
Exit Sub
Error_Handler:
    MsgBox Error, vbCritical, "WordPad : " ', X

End Sub

Public Sub NotePad(docName As String)
Dim X As String
    X = "notepad " & Chr$(34) & docName & Chr$(34)
    IdShell = Shell(X, 1)
    AppActivate IdShell
    DoEvents


End Sub

Public Sub Shell_Print(docName As String)
Dim xIn As String
Dim intFile As Integer
On Error GoTo Error_Handle
If Dir(docName) = "" Then
    MsgBox docName, vbCritical, "frmElpPrt.bas : document absent"
Else
    Set XPrt = Printer
    frmElpPrt.Show vbModeless
        
    blnFiligrane = False
    prtOrientation = vbPRORLandscape '
    prtStdBlankInit
    XPrt.FontSize = 8
    XPrt.FontName = prtFontName_CourierNew
    intFile = FreeFile(0)
    Open docName For Input As #intFile
    Do Until EOF(1)
        DoEvents
        Line Input #intFile, xIn
        XPrt.CurrentY = XPrt.CurrentY + 200
        If XPrt.CurrentY + 300 > prtMaxY Then XPrt.NewPage
        XPrt.CurrentX = prtMinX + 100: XPrt.Print xIn;

    Loop
    Close intFile
End If

Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide

Exit Sub

Error_Handle:

Shell_MsgBox docName & " : " & Error, vbInformation, "Shell_Print", False

End Sub

Public Sub MsPaint(docName As String)
On Error GoTo Error_Handler
Dim X As String
If Dir(constMsPaint, vbReadOnly + vbHidden) = "" Then
    MsgBox constMsPaint, vbCritical, "frmElpPrt.bas : programme absent"
Else
    X = constMsPaint & " " & Chr$(34) & docName & Chr$(34)
    IdShell = Shell(X, 1)
    AppActivate IdShell
    DoEvents
End If
Exit Sub
Error_Handler:
    MsgBox Error, vbCritical, "MsPaint : ", X
End Sub

Public Sub Excel(docName As String)
On Error GoTo Error_Handler
Dim X As String
If Dir(constExcel, vbReadOnly + vbHidden) = "" Then
    constExcel = constExcel_D
    If Dir(constExcel, vbReadOnly + vbHidden) = "" Then
    MsgBox constExcel, vbCritical, "frmElpPrt.bas : programme absent"
        Exit Sub
    End If
End If
        X = constExcel & " " & Chr$(34) & docName & Chr$(34)
        Shell_SendKeys = ""
        IdShell = Shell(X, 1)
        AppActivate IdShell
        DoEvents
Exit Sub
Error_Handler:
    MsgBox Error, vbCritical, "Excel : ", X

End Sub

Public Sub Acrord32(docName As String)
Dim X As String
If Dir(constAcrord32, vbReadOnly + vbHidden) = "" Then
    constAcrord32 = constAcrord32_D
    If Dir(constAcrord32, vbReadOnly + vbHidden) = "" Then
    MsgBox constAcrord32, vbCritical, "frmElpPrt.bas : programme absent"
        Exit Sub
    End If
End If
X = constAcrord32 & " " & Chr$(34) & docName & Chr$(34)

Shell_SendKeys = ""
IdShell = Shell(X, 1)
'AppActivate IdShell
DoEvents

End Sub

Public Sub WinWord_Print(docName As String, Interval As Integer)
On Error GoTo Error_Handler
Dim X As String
If Dir(constWinWord, vbReadOnly + vbHidden) = "" Then
    constWinWord = constWinWord_D
    If Dir(constWinWord, vbReadOnly + vbHidden) = "" Then
        MsgBox constWinWord, vbCritical, "frmElpPrt.bas : programme absent"
        Exit Sub
    End If
End If
X = constWinWord & " " & Chr$(34) & docName & Chr$(34)
IdShell = Shell(X, 1)
    AppActivate IdShell
''        WinWord docName
        DoEvents
        
        Shell_SendKeys = "Print"
        indexTimer = 0
        Timer.Interval = Interval
        Timer.Enabled = True
Exit Sub
Error_Handler:
    MsgBox Error, vbCritical, "WinWord"
End Sub

Public Sub Acrord32_Print(docName As String)
On Error GoTo Error_Handler
Dim X As String
If Dir(constAcrord32, vbReadOnly + vbHidden) = "" Then
    constAcrord32 = constAcrord32_D
    If Dir(constAcrord32, vbReadOnly + vbHidden) = "" Then
    MsgBox constAcrord32, vbCritical, "frmElpPrt.bas : programme absent"
        Exit Sub
    End If
End If
X = Chr$(34) & constAcrord32 & Chr$(34) & " /t " & Chr$(34) & docName & Chr$(34) & " " & Chr$(34) & Printer.Devicename & Chr$(34)

IdShell = Shell(X, 0)
AppActivate IdShell


Sleep 2000
AppActivate IdShell
SendKeys "%{F4}", True:    DoEvents

'indexTimer = 2
'Shell_SendKeys_Print
' SendKeys "{FIN}"

DoEvents

Exit Sub
Error_Handler:
    MsgBox Error & vbCrLf & docName, vbCritical, "Acrord32_Print"
End Sub

Public Sub WordPad_Print(docName As String, Interval As Integer)
On Error GoTo Error_Handler
Dim X As String
If Dir(constWordPad, vbReadOnly + vbHidden) = "" Then
    MsgBox constWordPad, vbCritical, "frmElpPrt.bas : programme absent"
Else
    X = constWordPad & " " & Chr$(34) & docName & Chr$(34)
    IdShell = Shell(X, 1)
    AppActivate IdShell
''''    WordPad docName
    DoEvents
    
    Shell_SendKeys = "Print"
    indexTimer = 0
    Timer.Interval = Interval
    Timer.Enabled = True
End If
Exit Sub
Error_Handler:
    MsgBox Error, vbCritical, "WordPad : "
Exit Sub
End Sub


Public Sub MsPaint_Print(docName As String, Interval As Integer)
If Dir(constMsPaint, vbReadOnly + vbHidden) = "" Then
    MsgBox constMsPaint, vbCritical, "frmElpPrt.bas : programme absent"
Else
    MsPaint docName
    DoEvents
    
    Shell_SendKeys = "Print"
    indexTimer = 0
    Timer.Interval = Interval
    Timer.Enabled = True
End If
End Sub


Public Sub Excel_Print(docName As String, Interval As Integer)
On Error GoTo Error_Handler
Dim X As String
If Dir(constExcel, vbReadOnly + vbHidden) = "" Then
    constExcel = constExcel_D
    If Dir(constExcel, vbReadOnly + vbHidden) = "" Then
    MsgBox constExcel, vbCritical, "frmElpPrt.bas : programme absent"
    Exit Sub
    End If
End If
        X = constExcel & " " & Chr$(34) & docName & Chr$(34)
        Shell_SendKeys = ""
        IdShell = Shell(X, 1)
        AppActivate IdShell
''''        Excel docName
        DoEvents
        
        Shell_SendKeys = "Print"
        indexTimer = 0
        Timer.Interval = Interval
        Timer.Enabled = True
Exit Sub
Error_Handler:
    MsgBox Error, vbCritical, "Excel: "

End Sub

Public Sub Shell_SendKeys_Print()
Select Case indexTimer
    Case 1: SendKeys "^p", True:    DoEvents
            SendKeys "{ENTER}", True:   DoEvents
    Case 2: SendKeys "%{F4}", True:    DoEvents: Timer.Enabled = False
End Select
End Sub

Private Sub Timer_Timer()
indexTimer = indexTimer + 1
Select Case Shell_SendKeys
    Case "Print": Shell_SendKeys_Print
End Select
End Sub


Public Function prtHeightDelta(Size1 As Integer, Size2 As Integer)
Dim I As Integer
On Error Resume Next
 XPrt.FontSize = Size1
I = XPrt.TextHeight("X")
XPrt.FontSize = Size2
prtHeightDelta = I - XPrt.TextHeight("X")

End Function

Public Sub prtStdTop()
On Error Resume Next

XPrt.DrawWidth = 3
XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMinY), prtLineColor
XPrt.FontBold = True
XPrt.FontSize = 9
XPrt.CurrentX = prtMinX + (prtMaxX - prtMinX - XPrt.TextWidth(Trim(prtTitleText))) / 2
XPrt.CurrentY = 0 '(prtMinY - XPrt.TextHeight(socName)) / 2
XPrt.ForeColor = prtForeColor_Header
XPrt.Print Trim(prtTitleText);

XPrt.FontSize = 6
XPrt.CurrentX = prtMinX
XPrt.Print socName;
XPrt.FontBold = False
XPrt.CurrentX = prtMaxX - XPrt.TextWidth(prtTitleUsr)
XPrt.Print prtTitleUsr;
XPrt.ForeColor = prtForeColor
End Sub

Public Sub prtStd()
prtStdTop
prtStdBottom
End Sub

Public Sub prtFormType_Select()
Select Case prtFormType
        Case "SAB": prtSAB_Bottom
        Case "STD": prtStd
        Case "SOC": prtSoc
        Case "BOT": prtStdBottom
        Case "TOP": prtStdTop
End Select

End Sub


Public Sub WinWord_Dir()
Dim wName As String, wMemo As String
On Error Resume Next


Call rsElpTable_Read("Windows.exe", "Excel", "", wName, constExcel)
If Dir(constExcel, vbReadOnly + vbHidden) = "" Then
    Call rsElpTable_Read("Windows.exe", "Excel", "1", wName, constExcel)
    If Dir(constExcel, vbReadOnly + vbHidden) = "" Then
        Call rsElpTable_Read("Windows.exe", "Excel", "2", wName, constExcel)
    End If
End If

Call rsElpTable_Read("Windows.exe", "WinWord", "", wName, constWinWord)
If Dir(constWinWord, vbReadOnly + vbHidden) = "" Then
    Call rsElpTable_Read("Windows.exe", "WinWord", "1", wName, constWinWord)
    If Dir(constWinWord, vbReadOnly + vbHidden) = "" Then
        Call rsElpTable_Read("Windows.exe", "WinWord", "2", wName, constWinWord)
    End If
End If

Call rsElpTable_Read("Windows.exe", "WordPad", "", wName, constWordPad)
If Dir(constWordPad, vbReadOnly + vbHidden) = "" Then
    Call rsElpTable_Read("Windows.exe", "WordPad", "1", wName, constWordPad)
    If Dir(constWordPad, vbReadOnly + vbHidden) = "" Then
        Call rsElpTable_Read("Windows.exe", "WordPad", "2", wName, constWordPad)
    End If
End If

Call rsElpTable_Read("Windows.exe", "MsPaint", "", wName, constMsPaint)
If Dir(constMsPaint, vbReadOnly + vbHidden) = "" Then
    Call rsElpTable_Read("Windows.exe", "MsPaint", "1", wName, constMsPaint)
    If Dir(constMsPaint, vbReadOnly + vbHidden) = "" Then
        Call rsElpTable_Read("Windows.exe", "MsPaint", "2", wName, constMsPaint)
    End If
End If

Call rsElpTable_Read("Windows.exe", "IExplorer", "", wName, constIExplorer)
If Dir(constIExplorer, vbReadOnly + vbHidden) = "" Then
    Call rsElpTable_Read("Windows.exe", "IExplorer", "1", wName, constIExplorer)
    If Dir(constIExplorer, vbReadOnly + vbHidden) = "" Then
        Call rsElpTable_Read("Windows.exe", "IExplorer", "2", wName, constIExplorer)
    End If
End If

constAcrord32 = FindPDF
constWinWord = FindDOCX

'If Dir("c:\Program Files (x86)") = "" Then
'    Call rsElpTable_Read("Windows.exe", "Acrord32", "", wName, constAcrord32)
'Else
'    Call rsElpTable_Read("Windows.exe", "Acrord32", "1", wName, constAcrord32)
'End If
'If Dir(constAcrord32, vbReadOnly + vbHidden) = "" Then
'    Call rsElpTable_Read("Windows.exe", "Acrord32", "5", wName, constAcrord32)
'    If Dir(constAcrord32, vbReadOnly + vbHidden) = "" Then
'        Call rsElpTable_Read("Windows.exe", "Acrord32", "4", wName, constAcrord32)
'        If Dir(constAcrord32, vbReadOnly + vbHidden) = "" Then
'            Call rsElpTable_Read("Windows.exe", "Acrord32", "1", wName, constAcrord32)
'            If Dir(constAcrord32, vbReadOnly + vbHidden) = "" Then
'                Call rsElpTable_Read("Windows.exe", "Acrord32", "2", wName, constAcrord32)
'                If Dir(constAcrord32, vbReadOnly + vbHidden) = "" Then
'                   Call rsElpTable_Read("Windows.exe", "Acrord32", "3", wName, constAcrord32) ' Adobe pro 7.0
'                End If
'           End If
'        End If
'    End If
'End If

If constExcel = "" Then constExcel = "C:\Program Files\Microsoft Office\Office12\Excel.exe"
If constWinWord = "" Then constWinWord = "C:\Program Files\Microsoft Office\Office12\WinWord.exe"
If constWordPad = "" Then constWordPad = "c:\Program Files\Windows NT\Accessoires\WordPad.exe"
If constMsPaint = "" Then constMsPaint = "c:\WinNT\System32\MsPaint.exe"
If constIExplorer = "" Then constIExplorer = "c:\Windows\ServicePackFiles\I386\iexplore.exe "
If constAcrord32 = "" Then constAcrord32 = "C:\Program Files\Adobe\Reader 9.0\Reader\Acrord32.exe"

End Sub

Public Sub prtFreeBottom(lPageRéférence As String, LPageNo As Integer)
XPrt.FontSize = 6

XPrt.CurrentY = prtMaxY + (300 - XPrt.TextHeight(lPageRéférence)) / 2
XPrt.CurrentX = prtMinX
XPrt.Print lPageRéférence;

If TypeOf XPrt Is Printer Then
    XPrt.CurrentX = prtMaxX - 200
    XPrt.Print Format$(LPageNo, "###");
End If

End Sub


Public Sub prtRTF(lTextRTF)
Dim startTextRTF As Long, lenTextRTF As Long, K As Long, K1 As Long, X1 As String * 1, wRTFCode As String
Dim blnRTFCode As Boolean
Dim wText As String, xCaractèreSpécial As String

arrRTF_FontColor_Nb = 0
Call prtRTF_FontName_Scan(lTextRTF)
Call prtRTF_FontColor_Scan(lTextRTF)

XPrt.CurrentX = prtMinMarge    '$ 2003.08.28 jpl

K = InStr(1, lTextRTF, "\pard")
K1 = K + 5: startTextRTF = K1
K = K1 + 1
blnRTFCode = True
lenTextRTF = Len(lTextRTF)
blnRTF_Trame = False
K = startTextRTF
wText = ""

Do
    
    X1 = Mid$(lTextRTF, K, 1)
    If X1 = Chr$(10) Or X1 = Chr$(13) Then Mid$(lTextRTF, K, 1) = " ": X1 = " "
    
        If Not blnRTFCode Then
            If X1 = "\" Then
                Select Case Mid$(lTextRTF, K + 1, 1)
                    Case "\": wText = wText & "\": K = K + 1
                    Case "~": wText = wText & " ": K = K + 1
                    Case "'":  xCaractèreSpécial = Mid$(lTextRTF, K, 4): K = K + 3
                    
                            If xCaractèreSpécial = "\'87" Then
                                prtRTF_Line wText: wText = "":   ''''XPrt.CurrentY = prtMaxY + 300
                                prtNewPage
                                If blnFiligrane Then frmElpPrt.prtFiligrane prtFiligrane_Name

                            Else
                                wText = wText & prtRTF_CaractèresSpéciaux(xCaractèreSpécial)
                    
                            End If
                    Case Else
                                prtRTF_Line wText: wText = ""                        ''''mId$(lTextRTF, K1, K - K1)
                                K1 = K
                                blnRTFCode = True
                End Select
            Else
                wText = wText & X1
            End If
        Else
            If X1 = "\" Then
                wRTFCode = Mid$(lTextRTF, K1, K - K1)
                prtRTF_Code wRTFCode
                K1 = K
                blnRTFCode = True
            Else
                If X1 = " " Then
                    wRTFCode = Mid$(lTextRTF, K1, K - K1)
                    prtRTF_Code wRTFCode
                    If Mid$(lTextRTF, K1, 2) = "\'" Then
                        K1 = K
                    Else
                        K1 = K + 1
                    End If
                    blnRTFCode = False
                End If
            End If
        End If

    K = K + 1

Loop Until K >= lenTextRTF
'''Next K

End Sub
Public Sub prtRTF_Attribut_CF(lTextRTF As Variant, lK1 As Long, lPos1 As Long, lPos2 As Long)
Dim K As Long, K1 As Long

lPos1 = 0: lPos2 = 0
K = InStr(lK1, lTextRTF, "\cf1 ")
If K > 0 Then
    lPos1 = K + 5
    lK1 = InStr(lPos1, lTextRTF, "\")
    If lK1 > 0 Then lPos2 = lK1 - 1
End If

End Sub


Public Sub prtRTF_Code(lRTFCode As String)
Dim I As Integer

'$jpl 20110518  __________________________________gestion Font Color
If Mid$(lRTFCode, 1, 3) = "\cf" Then
    I = Val(Mid$(lRTFCode, 4, Len(lRTFCode) - 3))
    XPrt.ForeColor = arrRTF_FontColor(I)
    Exit Sub
End If
'$jpl 20110518  _____________________________________________

If Mid$(lRTFCode, 1, 3) = "\fs" Then
    XPrt.FontSize = Val(Mid$(lRTFCode, 4, 2)) / 2
    I = XPrt.TextHeight("x")
    If I > prtlineHeight Then prtlineHeight = I + 10
Else
    If Mid$(lRTFCode, 1, 2) = "\f" Then
        I = Val(Mid$(lRTFCode, 3, 2))
        If I < 30 Then prtFontName = prtRTF_FontName(I): XPrt.FontName = prtFontName
    Else
        If Mid$(lRTFCode, 1, 3) = "\li" Then
            prtRTF_MinX = prtMinMarge + Val(Mid$(lRTFCode, 4, Len(lRTFCode) - 3))
            XPrt.CurrentX = prtRTF_MinX
            blnRTF_Center = False
        Else
            Select Case lRTFCode
                
                Case "\par": prtRTF_NewLine
                Case "\pard": prtRTF_MinX = prtMinMarge
                Case "\plain":  XPrt.FontSize = 8: XPrt.FontBold = False
                                XPrt.FontItalic = False: XPrt.FontUnderline = False
                                prtlineHeight = prtlineHeight66
                                blnRTF_Trame = False
                Case "\b": XPrt.FontBold = True
                Case "\b0": XPrt.FontBold = False
                Case "\i": XPrt.FontItalic = True
                Case "\i0": XPrt.FontItalic = False
                Case "\ul": XPrt.FontUnderline = True
                Case "\ul0", "\ulnone": XPrt.FontUnderline = False
                Case "\tab": XPrt.Print vbTab;
                Case "\qc": blnRTF_Center = True
'$jpl 20110518  Case "\cf0": blnRTF_Trame = False
'$jpl 20110518  Case "\cf1":    blnRTF_Trame = True
'$jpl 20110518                  prtRTF_TrameX1 = XPrt.CurrentX - 20
'$jpl 20110518                  prtRTF_TrameY1 = XPrt.CurrentY - 20
                
               
            End Select
        End If
    End If
End If

End Sub

Public Sub prtRTF_Line(lX As String)
Dim X As String, K1 As Long, K As Long

If XPrt.TextWidth(lX) < prtWidthMarge Then

    If blnRTF_Center Then
        Call prtCentré(prtMedX, lX)
    Else
        prtRTF_Trame lX
    End If
Else
    K1 = 1
    For K = 2 To Len(lX)
        If XPrt.TextWidth(Mid$(lX, K1, K - K1)) >= prtWidthMarge Then
            If Mid$(lX, K + 2, 1) = " " Then
                K = K + 3
            Else
                If Mid$(lX, K + 1, 1) = " " Then
                    K = K + 2
                Else
                    If Mid$(lX, K + 1, 1) = " " Then
                        K = K + 1
                    Else
                        If Mid$(lX, K - 3, 1) = " " Then
                            K = K - 2
                        Else
                            If Mid$(lX, K - 2, 1) = " " Then K = K - 1
                        End If
                    End If
                End If
            End If
                        
            prtRTF_Trame Mid$(lX, K1, K - K1)
            prtRTF_NewLine
            K1 = K
        End If
    Next K
    If K > K1 Then prtRTF_Trame Mid$(lX, K1, K - K1)
End If

End Sub

Public Sub prtRTF_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + prtlineHeight >= prtMaxY Then
    If Not blnMinX Then
        prtNewPage
        If blnFiligrane Then frmElpPrt.prtFiligrane prtFiligrane_Name

    Else
        If blnMinX12 Then
            prtNewPage
            frmElpPrt.prtLineX prtMedX0
            blnMinX12 = False
            prtMinMarge = prtMinX1
            prtMedX = prtMinX1 + (prtMinX2 - prtMinX1) / 2
        Else
            blnMinX12 = True
            prtMinMarge = prtMinX2
            prtMedX = prtMinX2 + (prtMaxX - prtMinX2) / 2
        End If
        XPrt.CurrentY = prtMinY + prtHeaderHeight + prtlineHeight

    End If
End If
XPrt.CurrentX = prtMinMarge
If blnRTF_Trame Then
    prtRTF_TrameX1 = XPrt.CurrentX - 20
    prtRTF_TrameY1 = XPrt.CurrentY - 20
End If

End Sub

Public Sub prtNewLine()
 XPrt.CurrentX = prtMinX
 XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + prtlineHeight >= prtMaxY Then prtNewPage
End Sub

Public Function prtRTF_CaractèresSpéciaux(lRTFCode As String) As String
Select Case lRTFCode
    Case "\'e0": prtRTF_CaractèresSpéciaux = "à"
    Case "\'e2": prtRTF_CaractèresSpéciaux = "â"
    Case "\'e4": prtRTF_CaractèresSpéciaux = "ä"
    Case "\'e7": prtRTF_CaractèresSpéciaux = "ç"
    Case "\'e8": prtRTF_CaractèresSpéciaux = "è"
    Case "\'e9": prtRTF_CaractèresSpéciaux = "é"
    Case "\'ea": prtRTF_CaractèresSpéciaux = "ê"
    Case "\'eb": prtRTF_CaractèresSpéciaux = "ë"
    Case "\'ef": prtRTF_CaractèresSpéciaux = "ï"
    Case "\'f4": prtRTF_CaractèresSpéciaux = "ô"
    Case "\'f6": prtRTF_CaractèresSpéciaux = "ö"
    Case "\'f9": prtRTF_CaractèresSpéciaux = "ù"
    Case "\'fb": prtRTF_CaractèresSpéciaux = "û"
    Case "\'fc": prtRTF_CaractèresSpéciaux = "ü"

    Case "\'c0": prtRTF_CaractèresSpéciaux = Chr$(192)
    Case "\'c2": prtRTF_CaractèresSpéciaux = Chr$(194)
    Case "\'c4": prtRTF_CaractèresSpéciaux = Chr$(196)
    Case "\'c7": prtRTF_CaractèresSpéciaux = Chr$(199)
    Case "\'c8": prtRTF_CaractèresSpéciaux = Chr$(200)
    Case "\'c9": prtRTF_CaractèresSpéciaux = Chr$(201)
    Case "\'ca": prtRTF_CaractèresSpéciaux = Chr$(202)
    Case "\'cb": prtRTF_CaractèresSpéciaux = Chr$(203)
    Case "\'cf": prtRTF_CaractèresSpéciaux = Chr$(207)
    Case "\'c4": prtRTF_CaractèresSpéciaux = Chr$(212)
    Case "\'c6": prtRTF_CaractèresSpéciaux = Chr$(214)
    Case "\'c9": prtRTF_CaractèresSpéciaux = Chr$(217)
    Case "\'cb": prtRTF_CaractèresSpéciaux = Chr$(219)
    Case "\'cc": prtRTF_CaractèresSpéciaux = Chr$(220)
    
    Case "\'b0": prtRTF_CaractèresSpéciaux = Chr$(186)   '°

End Select

End Function

Public Sub prtRTF_FontName_Scan(lTextRTF)
Dim I As Integer, K As Integer, K1 As Integer, K2 As Integer
For K = 0 To 30
    prtRTF_FontName(K) = prtFontName_Arial
Next K
K = InStr(1, lTextRTF, "{\fonttbl")
If K > 0 Then
    Do
        K2 = -1
        K1 = InStr(K, lTextRTF, "{\f")
        If K1 > 0 Then
            K = K1 + 3
            K2 = InStr(K, lTextRTF, "\")
            If K2 > 0 Then
                I = Val(Mid$(lTextRTF, K, K2 - K))
                K2 = InStr(K2, lTextRTF, " ")
                If K2 > 0 Then
                    K = K2 + 1
                    K2 = InStr(K2, lTextRTF, ";}")
                    If K2 > 0 Then
                        If I < 30 Then prtRTF_FontName(I) = Mid$(lTextRTF, K, K2 - K)
                    End If
                End If
            End If
        End If
    Loop Until K2 < 0
End If

End Sub

Public Sub prtRTF_FontColor_Scan(lTextRTF)
Dim I As Integer, K As Integer, K1 As Integer, K2 As Integer, K3 As Integer
Dim C1 As Integer, C2 As Integer, C3 As Integer
arrRTF_FontColor(0) = RGB(0, 0, 0)
K = InStr(1, lTextRTF, "{\colortbl")
If K > 0 Then
    Do
        K1 = InStr(K, lTextRTF, "\red")
        If K1 <= 0 Then
            Exit Do
        Else
            K = K1 + 4
            K2 = InStr(K, lTextRTF, "\")
            If K2 > 0 Then C1 = Val(Mid$(lTextRTF, K, K2 - K))
        End If

        K1 = InStr(K, lTextRTF, "\green")
        If K1 <= 0 Then
            Exit Do
        Else
            K = K1 + 6
            K2 = InStr(K, lTextRTF, "\")
            If K2 > 0 Then C2 = Val(Mid$(lTextRTF, K, K2 - K))
        End If

        K1 = InStr(K, lTextRTF, "\blue")
        If K1 <= 0 Then
            Exit Do
        Else
            K = K1 + 5
            K2 = InStr(K, lTextRTF, ";")
            If K2 > 0 Then C3 = Val(Mid$(lTextRTF, K, K2 - K))
        End If
        arrRTF_FontColor_Nb = arrRTF_FontColor_Nb + 1
        arrRTF_FontColor(arrRTF_FontColor_Nb) = RGB(C1, C2, C3)
    Loop
End If

End Sub

Public Sub prtRTF_Trame(lX As String)
Dim mCurrentX As Integer, mCurrenty As Integer

If Not blnRTF_Trame Then
    XPrt.Print lX;
Else
    mCurrentX = XPrt.CurrentX
    mCurrenty = XPrt.CurrentY
    prtRTF_TrameX2 = XPrt.CurrentX + XPrt.TextWidth(lX) + 20
    prtRTF_TrameY2 = XPrt.CurrentY + prtlineHeight - 20
    Call frmElpPrt.prtTrame(prtRTF_TrameX1, prtRTF_TrameY1, prtRTF_TrameX2, prtRTF_TrameY2, " ", 240)
    XPrt.CurrentX = mCurrentX
    XPrt.CurrentY = mCurrenty
    XPrt.Print lX;
    prtRTF_TrameX1 = XPrt.CurrentX - 20
    prtRTF_TrameY1 = XPrt.CurrentY + 20
End If

End Sub

Public Sub prtIMP_PDF_Monitor(lFct As String)
Static blnOpen As Boolean, prtIMP_PDF_Seq As Long, arrIMP_PDF() As String, KaF As Integer
Static objFolder, objFiles_Open, objFiles_Close

Dim K As Integer, K2 As Integer
Dim blnOk As Boolean
Dim currentFileName As String, tmpFileName As String
Dim V, X8 As String, xPath As String, xMin As String
Dim fsoFile As Scripting.File, fsoFile2 As Scripting.File
On Error GoTo Error_Handler

Select Case lFct
    Case "Open":
        prtFillColor = RGB(240, 240, 240)
        If Not blnOpen Then
            blnOpen = True
            Set objFolder = msFileSystem.GetFolder(paramIMP_PDF_Path)
            Set objFiles_Open = objFolder.Files
        End If
    Case "Close"
        If blnOpen Then
            DoEvents: Wait_SS 2: DoEvents
            currentFileName = ""
            Set objFolder = msFileSystem.GetFolder(paramIMP_PDF_Path)
            Set objFiles_Close = objFolder.Files
            blnOk = False
            For Each fsoFile In objFiles_Close
                currentFileName = fsoFile.Name
                For Each fsoFile2 In objFiles_Open
                    If fsoFile.Name = fsoFile2.Name Then blnOk = True: Exit For
                Next
                If Not blnOk Then Exit For
            Next
            tmpFileName = paramIMP_PDF_Path & "\" & currentFileName
            '''Call MsgBox(tmpFileName, vbInformation, "prtIMP_PDF_Monitor")
            prtIMP_PDF_Seq = prtIMP_PDF_Seq + 1
'$2007-01-31 JPL : archivage définitif
            If Mid$(prtPgmName, 1, 2) = "\\" Then
                prtIMP_PDF_FileName = prtPgmName
            Else
                If blnEditionNoPaper_Auto Then
                    Dim xUnit As String, xDir_Save As String, xFile_Save As String
                    xDir_Save = paramEditionNoPaper_Folder & "PDF\" & paramEditionNoPaper_Auto_Dir & "_" & YBIATAB0_DATE_CPT_J
                    If Not msFileSystem.FolderExists(xDir_Save) Then MkDir xDir_Save

                    xUnit = Table_Unit_SSI("S", paramEditionNoPaper_Auto_Unit)
                    If xUnit = "" Then xUnit = "S00"
                    xFile_Save = xUnit & "." & DSys & "_" & time_Hms & "_" & paramEditionNoPaper_Auto_PgmName & "_" & prtIMP_PDF_Seq & " (" & paramEditionNoPaper_Auto_Unit & ").pdf"
                    
                    prtIMP_PDF_FileName = xDir_Save & "\_" & xFile_Save
                    
                    paramEditionNoPaper_Auto_Lnk = "<span style='font-size:9.0pt;font-family:Calibri'>""" _
                                             & "<A HREF=" & Asc34 & Replace(prtIMP_PDF_FileName, paramEditionNoPaper_Folder & "PDF\", paramEditionNoPaper_Partage) & Asc34 & ">" _
                                            & "Cliquez ici pour afficher le document : " & xFile_Save & "</A><BR><BR>"
                                            

                Else
                     prtIMP_PDF_FileName = paramIMP_PDF_Path & "\Archive\" & DSys & "_" & time_Hms & "_" & prtIMP_PDF_Seq & "_" & prtPgmName & ".pdf"
                End If
            End If
            On Error Resume Next
            K = 0
xxx:
            DoEvents
            K = K + 1
            'Wait_SS 1: DoEvents
            
            msFileSystem.MoveFile tmpFileName, prtIMP_PDF_FileName
            
            'If Trim(Dir(tmpFileName)) <> "" Then
            If msFileSystem.FileExists(tmpFileName) Then 'DR 30/08/2018
                MsgBox prtIMP_PDF_FileName, Error, "prtIMP_Pdf_Monitor"
                DoEvents
                Wait_SS 2
                DoEvents
               ' If K > 5 Then GoTo xxx
                If K <= 5 Then GoTo xxx
            End If
        End If
        blnOpen = False
    Case "Clear"
        If Not msFileSystem.FolderExists(paramIMP_PDF_Path) Then MkDir paramIMP_PDF_Path
        xPath = paramIMP_PDF_Path & "\Archive"
        If Not msFileSystem.FolderExists(xPath) Then MkDir xPath
        
        Set objFolder = msFileSystem.GetFolder(paramIMP_PDF_Path)
        Set objFiles_Close = objFolder.Files
        For Each fsoFile In objFiles_Close
            If Err = 0 Then
                Call dateJMA6_AMJ(fsoFile.DateLastModified, X8)
                If X8 < DSys Then msFileSystem.DeleteFile fsoFile.PATH, True
            End If
        Next
        
        xMin = dateElp("Jour", -8, DSys)
        Set objFolder = msFileSystem.GetFolder(xPath)
        Set objFiles_Close = objFolder.Files
        For Each fsoFile In objFiles_Close
            If Err = 0 Then
                Call dateJMA6_AMJ(fsoFile.DateLastModified, X8)
                If X8 < xMin Then msFileSystem.DeleteFile fsoFile.PATH, True
            End If
        Next

End Select
Exit Sub

Error_Handler:
    blnOpen = False
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, "prtIMP_PDF_Monitor"

End Sub
Public Sub prtIMP_PDF_NoPaper_Init(lUnit As String, lPgmName As String, lDir As String)
On Error GoTo Error_Handler
        If Not IsEmpty(Printer) Then Printer_Previous_DeviceName = Printer.Devicename
        blnEditionNoPaper_Auto = True
        paramEditionNoPaper_Auto_Unit = lUnit
        paramEditionNoPaper_Auto_PgmName = lPgmName
        paramEditionNoPaper_Auto_Dir = lDir
        If Not xlsManual Then
            Call Printer_PDF
        End If

Exit Sub

Error_Handler:
    Dim V
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, "prtIMP_PDF_NoPaper_Init"

End Sub
Public Sub prtIMP_PDF_NoPaper_CopyFile(lUnit As String, lFileName As String, lDir As String, lPgmName As String)
On Error GoTo Error_Handler

Dim xUnit As String, xDir_Save As String, xFile_Save As String
Dim K As Integer

paramEditionNoPaper_Auto_Unit = lUnit

xDir_Save = paramEditionNoPaper_Folder & "PDF\" & lDir & "_" & YBIATAB0_DATE_CPT_J
If Not msFileSystem.FolderExists(xDir_Save) Then MkDir xDir_Save

xUnit = Table_Unit_SSI("S", lUnit)
xFile_Save = xUnit & "." & DSys & "_" & time_Hms & "_" & lPgmName & "_1" & " (" & lUnit & ")." & fileName_Extension(lFileName)

prtIMP_PDF_FileName = xDir_Save & "\_" & xFile_Save

paramEditionNoPaper_Auto_Lnk = "<span style='font-size:9.0pt;font-family:Calibri'>""" _
                         & "<A HREF=" & Asc34 & Replace(prtIMP_PDF_FileName, paramEditionNoPaper_Folder & "PDF\", paramEditionNoPaper_Partage) & Asc34 & ">" _
                        & "Cliquez ici pour afficher le document : " & xFile_Save & "</A><BR><BR>"

    K = 0
xxx:
    DoEvents
    K = K + 1
    
    msFileSystem.MoveFile lFileName, prtIMP_PDF_FileName
    
'    If Trim(Dir(lFileName)) <> "" Then
    If msFileSystem.FileExists(lFileName) Then 'DR 30/08/2018
        MsgBox prtIMP_PDF_FileName, Error, "prtIMP_Pdf_Monitor"
        DoEvents
        Wait_SS 2
        DoEvents
        If K <= 5 Then GoTo xxx
    End If

Exit Sub

Error_Handler:
    Dim V
    V = Error
Error_MsgBox:
    MsgBox "Unit : " & lUnit & vbCrLf & "FileName : " & lFileName & vbCrLf & "Dir : " & lDir & vbCrLf & "PgmName : " & lPgmName _
          & "Erreur : " & V, vbCritical, "prtIMP_PDF_NoPaper_CopyFile " & Date & " " & Time

End Sub

Public Sub prtIMP_PDF_NoPaper_Mail(lFrom As String, lRecipient As String, lMsg As String)
On Error GoTo Error_Handler
Dim wSendMail As typeSendMail, xLib As String, xMemo As String, V
    
If lFrom = "" Then
    wSendMail.From = currentSSIWINMAIL
    If lRecipient = "" Then
        wSendMail.Recipient = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire(paramEditionNoPaper_Auto_Unit)
    Else
        wSendMail.Recipient = lRecipient
    End If
    
Else
    wSendMail.Recipient = srvSendMail.Exchange_Distribution(lFrom, lRecipient)
    'wSendMail.FromDisplayName = lFrom
    'wSendMail.RecipientDisplayName = lRecipient
End If

wSendMail.FromDisplayName = "NoPaper " & paramEditionNoPaper_Auto_PgmName
wSendMail.CcRecipient = ""

V = rsElpTable_Read(constEdition_Form, "SAB", paramEditionNoPaper_Auto_PgmName, xLib, xMemo)
If IsNull(V) Then
    wSendMail.Subject = Trim(xLib)
Else
    wSendMail.Subject = paramEditionNoPaper_Auto_PgmName
End If
wSendMail.Attachment = "" 'lFileName
wSendMail.Message = mHtml_Head & "<span style='font-size:10.0pt;font-family:Calibri'>" _
                 & htmlFontColor_Black & lMsg & "<BR><BR>" & paramEditionNoPaper_Auto_Lnk & "</div></body></html>"

 wSendMail.AsHTML = True
 srvSendMail.Monitor wSendMail
 
 Printer_Reset

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, "prtIMP_PDF_NoPaper_Mail"

End Sub

Public Sub prtIMP_PDF_NoPaper_Print(lUnit As String)
On Error GoTo Error_Handler

 If paramEnvironnement = constProduction Then
    Dim meUnit As typeUnit
    meUnit.Id = lUnit
    Table_Unit meUnit
    Printer_Set meUnit.Printer
    Call Acrord32_Print(prtIMP_PDF_FileName)
    Printer_Reset
Else
    Printer_Reset
   Debug.Print "prtIMP_PDF_NoPaper Print :"; Printer.Devicename
   Call Acrord32_Print(prtIMP_PDF_FileName)
End If

Exit Sub

Error_Handler:
    Dim V
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, "prtIMP_PDF_NoPaper_Print"

End Sub

Public Function prtIMP_PDF_NoPaper_Destinaire(lUnit As String) As String
Dim xDestinataire As String
xDestinataire = srvSendMail.Exchange_Distribution("NoPaper." & lUnit, "")
If xDestinataire = "" Then xDestinataire = srvSendMail.Exchange_Distribution("NoPaper.S00", "")
prtIMP_PDF_NoPaper_Destinaire = xDestinataire
End Function

Public Sub prtColor_Check_1()
On Error Resume Next
Printer.ColorMode = 1
prtColorMode = False
prtLineColor = prtLineColor_Black
prtFillColor = prtFillColor_Black
frmElpPrt.imgSocLogo.Picture = LoadPicture(paramSocLogo_G)

End Sub

Public Sub prtColor_Check_2()
On Error Resume Next
Printer.ColorMode = 2
prtColorMode = True
prtFillColor = prtFillColor_Standard
prtLineColor = prtLineColor_Standard
frmElpPrt.imgSocLogo.Picture = LoadPicture(paramSocLogo)

End Sub

Public Sub Msg_Rcv()

End Sub
