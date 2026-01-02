VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmCptMvt 
   Caption         =   "CptMvt :Interrogation des Mouvements Comptables"
   ClientHeight    =   6300
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6300
   ScaleWidth      =   9330
   Begin VB.CheckBox chkAmj 
      Caption         =   "Période"
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   0
      Width           =   855
   End
   Begin VB.Frame fraAmj 
      Height          =   400
      Left            =   4000
      TabIndex        =   10
      Top             =   -100
      Visible         =   0   'False
      Width           =   3255
      Begin MSComCtl2.DTPicker txtAmjMax 
         Height          =   300
         Left            =   2040
         TabIndex        =   11
         Top             =   100
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   12632256
         CustomFormat    =   "dd  MM yyy"
         Format          =   69206019
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txtAmjMin 
         Height          =   300
         Left            =   480
         TabIndex        =   12
         Top             =   100
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   12632256
         CustomFormat    =   "dd  MM yyy"
         Format          =   69206019
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   -328351
      End
      Begin VB.Label lblAmjMin 
         Caption         =   "du"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   100
         Width           =   315
      End
      Begin VB.Label lblAmjMax 
         Caption         =   "au"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   100
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   8880
      Picture         =   "CPTMVT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   400
   End
   Begin VB.CheckBox chkColor 
      Caption         =   "Couleur"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   6050
      Value           =   1  'Checked
      Width           =   1000
   End
   Begin VB.CheckBox chkPageParPage 
      Caption         =   "Page / Page"
      Height          =   255
      Left            =   7900
      TabIndex        =   7
      Top             =   6050
      Value           =   1  'Checked
      Width           =   1300
   End
   Begin VB.CheckBox chkEffacer 
      Caption         =   "Effacer à chaque recherche"
      Height          =   255
      Left            =   5500
      TabIndex        =   6
      Top             =   6050
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.PictureBox picCptMvt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   4800
      Left            =   0
      ScaleHeight     =   4800
      ScaleWidth      =   9405
      TabIndex        =   5
      Top             =   1080
      Width           =   9405
      Begin VB.Shape shapeCptMvt 
         BackColor       =   &H00C0C0C0&
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00C0C0C0&
         Height          =   320
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   9300
      End
   End
   Begin VB.PictureBox picCompte 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00E0E0E0&
      Height          =   850
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   9345
      TabIndex        =   4
      Top             =   310
      Width           =   9400
   End
   Begin VB.CommandButton cmdRechercher 
      Caption         =   "Rechercher"
      Default         =   -1  'True
      Height          =   300
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdPagePrior 
      BackColor       =   &H00C0FFFF&
      Caption         =   "<<....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1005
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdPageNext 
      BackColor       =   &H00C0FFFF&
      Caption         =   "... >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1000
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ErrMsg"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   7320
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu picCptMvtmnu 
      Caption         =   "picCptMvtmnu"
      Visible         =   0   'False
      Begin VB.Menu picCptMvtmnuImpressionA4 
         Caption         =   "Impression Extrait (A4)"
      End
      Begin VB.Menu picCptMvtmnuImpressionA4_Euro 
         Caption         =   "Impression Extrait (A4) + Contre-Valeur en  Euro"
      End
      Begin VB.Menu mnuX1 
         Caption         =   "-"
      End
      Begin VB.Menu picCptMvtmnuImpressionExtrait 
         Caption         =   "Impression Extrait (préimprimé)"
      End
      Begin VB.Menu picCptMvtmnuImpressionListe 
         Caption         =   "Impression Liste"
      End
      Begin VB.Menu mnuX2 
         Caption         =   "-"
      End
      Begin VB.Menu picCptMvtmnuDétail 
         Caption         =   "Détail"
      End
      Begin VB.Menu picCptMvtmnuPièce 
         Caption         =   "Pièce"
      End
   End
End
Attribute VB_Name = "frmCptMvt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer

Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor

Dim CptMvtMethod As String * 12, mCptMvt_Obj As String * 12
Dim recCptMvt As typeCptMvt
Dim recCompte As typeCompte


Private Const frmCptMvtNb = 15
Private frmCptMvtNbMin As Integer
Private frmCptMvtNbMax As Integer

Dim valAmjMin As String * 8
Dim valAmjMax As String * 8
Dim mvalAmjMin As String * 8
Dim mvalAmjMax As String * 8

Dim mDevx As String * 3
Public Amont As Boolean
Dim prtListe As Boolean, prtPréimprimé As Boolean, optCV_Euro As Boolean
Dim mIndex As Integer

Public Me_picLineHeight As Integer
Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Then
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            MouseMoveActiveControl.ForeColor = C.ForeColor
            C.ForeColor = MouseMoveUsr.ForeColor
        End If
    End If
End If

End Sub


Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
        If TypeOf xobj Is CommandButton Then
            xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            xobj.ForeColor = MouseMoveActiveControl.ForeColor
        End If
        Exit For
    End If
Next xobj

End Sub

'---------------------------------------------------------
Public Sub frmCptMvtDShow()
'---------------------------------------------------------
Dim X As String

frmCptMvtD.Show vbModeless
frmCptMvtD.WindowState = vbNormal
frmCptMvtD.Visible = True
frmCptMvtD.frmCptMvtDInit
X = frmCptMvtD.Caption
AppActivate X

End Sub


'---------------------------------------------------------
Public Sub frmCptMvtEffacer()
'---------------------------------------------------------
Dim K As Integer

arrCptMvtNb = 0: arrCptMvtIndex = 0
recCptMvtInit recCptMvt: recCptMvt.obj = mCptMvt_Obj
picCptMvt.Visible = False

frmCptMvtNbMin = 0: frmCptMvtNbMax = 0
frmCptMvtEnabled

End Sub

'---------------------------------------------------------
Private Sub frmCptMvtEnabled()
'---------------------------------------------------------
Dim I As Integer

If arrCptMvtNb > 0 Then
    cmdPrint.Enabled = True
End If
If Amont Then
    frmcptmvtEnabledAmont
Else
    cmdPageNext.Caption = Format$(frmCptMvtNbMax, "####") & "..>>"
    If frmCptMvtNbMax < arrCptMvtNb Or arrCptMvtSuite Then
        cmdPageNext.Enabled = True
        cmdPageNext.BackColor = &HC0FFC0
    Else
        cmdPageNext.Enabled = False
        cmdPageNext.BackColor = &H8000000F
    End If
    cmdPagePrior.Caption = "<<..." & Format$(frmCptMvtNbMin, "####")
    If frmCptMvtNbMin > 1 Then
        cmdPagePrior.Enabled = True
        cmdPagePrior.BackColor = &HC0FFC0
    Else
        cmdPagePrior.Enabled = False
        cmdPagePrior.BackColor = &H8000000F
    End If
End If

End Sub


'---------------------------------------------------------
Public Sub frmCptMvtRechercher()
'---------------------------------------------------------

If Not IsNull(valDate) Then Exit Sub

If chkEffacer Then frmCptMvtEffacer

frmCptMvtSql
mvalAmjMin = valAmjMin
mvalAmjMax = valAmjMax
If Amont Then frmCptMvtJourAmont

If IsNull(srvCptMvtMon(recCptMvt)) Then
    If mId$(CptMvtMethod, 1, 4) = "Seek" Then
        Call arrCptMvtAddItem(recCptMvt)
        Call frmCptMvtItem
    Else
        mIndex = arrCptMvtIndex
        If Amont Then
            If arrCptMvtSuite Then arrCptMvtNb = frmCptMvtNb * CInt((arrCptMvtNb + 1) / frmCptMvtNb)
        Else
            If Not arrCptMvtSuite Then frmCptMvtJourAval
        End If
        Call frmCptMvtItem
    End If
End If

If chkPageParPage Then
Else
   frmCptMvtSuiteEtFin
End If

End Sub





'---------------------------------------------------------
Private Sub frmCptMvtSql()
'---------------------------------------------------------
     
recCptMvt.Société = recCompte.Société
recCptMvt.Agence = recCompte.Agence
recCptMvt.Devise = recCompte.Devise
recCptMvt.Compte = recCompte.Numéro
arrCptMvt(0) = recCptMvt

If valAmjMin = "00000000" Then
    Amont = True
    CptMvtMethod = "PrevJA"
    arrCptMvt(0).Method = CptMvtMethod
    arrCptMvt(0).AmjTraitement = valAmjMin
    arrCptMvt(0).Pièce = 0
    arrCptMvt(0).Ligne = 0
    recCptMvt.Method = CptMvtMethod
    recCptMvt.AmjTraitement = valAmjMax
    recCptMvt.Pièce = "999999999"
    recCptMvt.Ligne = "9999"
Else
    Amont = False
    CptMvtMethod = "SnapJA"
    recCptMvt.Method = CptMvtMethod
    recCptMvt.AmjTraitement = valAmjMin
    recCptMvt.Pièce = 0
    recCptMvt.Ligne = 0
    arrCptMvt(0).Method = CptMvtMethod
    arrCptMvt(0).AmjTraitement = valAmjMax
    arrCptMvt(0).Pièce = "999999999"
    arrCptMvt(0).Ligne = "9999"
End If

End Sub



'---------------------------------------------------------
Private Sub frmCptMvtSuite()
'---------------------------------------------------------
Dim X As String

If Not IsNull(valDate) Then Exit Sub
lblErrMsg = "Lecture des mouvements": lblErrMsgStd
If arrCptMvtNb > 0 Then
    recCptMvt = arrCptMvt(arrCptMvtNb)
    X = "+"
Else
    If chkEffacer Then frmCptMvtEffacer
    frmCptMvtSql
    X = " "
End If
    
recCptMvt.Method = mId$(CptMvtMethod, 1, 6) & X
 
If IsNull(srvCptMvtMon(recCptMvt)) Then
    mIndex = arrCptMvtIndex
    Call frmCptMvtItem
    If Not Amont And Not arrCptMvtSuite Then frmCptMvtJourAval: frmCptMvtItem
End If
lblErrMsg.Visible = False

End Sub

'---------------------------------------------------------
Public Sub frmCptMvtSuiteEtFin()
'---------------------------------------------------------
Dim X As String, Nb As Integer
Nb = arrCptMvtNb + 150

frmCptMvt.Enabled = False

Do Until Not arrCptMvtSuite
    frmCptMvtSuite
    If arrCptMvtNb > Nb Then
        X = MsgBox(arrCptMvtNb & " mouvements lus" & Chr$(13) & "Voulez-vous continuer?", vbYesNo + vbQuestion + vbDefaultButton2, "Extraction des mouvements")
        If X = vbNo Then
            Exit Do
        Else
            Nb = arrCptMvtNb + 100
        End If
    End If
Loop

If frmCptMvtNbMin > frmCptMvtNbMax Then cmdPagePrior_Click

frmCptMvt.Enabled = True
cmdRechercher.SetFocus

End Sub

'---------------------------------------------------------
Private Sub frmCptMvtItem()
'---------------------------------------------------------
Dim X As String, XX As String, K As Integer
Dim Mt As Currency
Dim I As Integer, K1 As Integer, K2 As Integer, K3 As Integer
Dim col2 As Integer, col3 As Integer, col4 As Integer, col5 As Integer
Dim wD As Double

col2 = 1200: col3 = 2500: col4 = 8500: col5 = 8800

If Me.WindowState = vbNormal Then
    Me_picLineHeight = picLineHeight
Else
    Me_picLineHeight = picLineHeight * picCptMvt.Height / 4800
    wD = picCptMvt.Width / 9350
    col2 = col2 * wD: col3 = col3 * wD: col4 = col4 * wD: col5 = col5 * wD
End If

picCptMvt.Visible = True
picCptMvt.Cls
If chkColor Then
    picCptMvt.BackColor = frmUsr.BackColor
Else
    picCptMvt.BackColor = vbWhite
End If

If arrCptMvtIndex < 0 Then
    arrCptMvtIndex = 0
End If
If arrCptMvtNb <= 0 Then
    lblErrMsg.Caption = "Néant    ( " & Now & " )"
    lblErrMsg.Visible = True
    lblErrMsg.Top = 2500
    lblErrMsg.Left = 2500
    lblErrMsg.FontBold = True
    lblErrMsg.FontSize = 12
    frmCptMvtNbMin = 0
Else
    frmCptMvtNbMin = arrCptMvtIndex + 1
End If
    
If Amont Then
    K2 = 0
    K1 = frmCptMvtNb - 1
    K3 = -1
Else
    K1 = 0
    K2 = frmCptMvtNb - 1
    K3 = 1
End If
For K = K1 To K2 Step K3

    If arrCptMvtIndex >= arrCptMvtNb Then
'       piccptmvt.Visible = False
    Else
       arrCptMvtIndex = arrCptMvtIndex + 1
       
       
       picCptMvt.CurrentY = K * Me_picLineHeight + 20
       picCptMvt.CurrentX = 50
       picCptMvt.ForeColor = libUsr.ForeColor
       picCptMvt.FontBold = False
       picCptMvt.Print dateImp(arrCptMvt(arrCptMvtIndex).AmjTraitement);
       
       picCptMvt.CurrentX = col3
       picCptMvt.Print arrCptMvt(arrCptMvtIndex).Libellé;
        
       If arrCptMvt(arrCptMvtIndex).AmjTraitement <> arrCptMvt(arrCptMvtIndex).AmjValeur Then
        picCptMvt.ForeColor = warnUsrColor
           picCptMvt.CurrentX = col2
           picCptMvt.Print dateImp(arrCptMvt(arrCptMvtIndex).AmjValeur);
       End If
      
       picCptMvt.FontBold = True
       
       Mt = arrCptMvt(arrCptMvtIndex).Mt
        
       picCptMvt.ForeColor = IIf(Mt < 0, errUsr.ForeColor, libUsr.ForeColor)
       X = Format$(Abs(Mt), "#### ### ### ### ##0.00")
       picCptMvt.CurrentX = col4 - TextWidth(X)
        picCptMvt.Print X;
        picCptMvt.CurrentX = col5
        picCptMvt.ForeColor = warnUsrColor
        picCptMvt.FontBold = False
        picCptMvt.Print mDevx;
        
    End If
Next K

frmCptMvtNbMax = arrCptMvtIndex

Call frmCptMvtEnabled

End Sub


'---------------------------------------------------------
Public Sub lblErrMsgStd()
'---------------------------------------------------------
Beep
lblErrMsg.Visible = True
lblErrMsg.FontBold = True
lblErrMsg.FontSize = 8
lblErrMsg.Top = 0
lblErrMsg.Left = 6500

End Sub


Private Sub chkAmj_Click()
If chkAmj = "0" Then
    chkPageParPage.Value = 1
    valAmjMin = "00000000": valAmjMax = DSys
    fraAmj.Visible = False
Else
    chkPageParPage.Value = 0
    fraAmj.Visible = True
    txtAmjMin_Control
    txtAmjMax_Control
End If

End Sub

'---------------------------------------------------------
Private Sub chkColor_Click()
'---------------------------------------------------------

frmCptMvtTête

arrCptMvtIndex = frmCptMvtNbMin - 1
frmCptMvtItem

End Sub

'-------------------------------------------------'
Private Sub txtAmjMax_Control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAmjMax.Year, "0000") & Format$(txtAmjMax.Month, "00") & Format$(txtAmjMax.Day, "00")
lblAmjMax.ForeColor = lblUsr.ForeColor
If Not IsNumeric(X) Then
    lblAmjMax.ForeColor = errUsr.ForeColor
    DTPicker_Now txtAmjMax
Else
    If mId$(X, 1, 8) > DSys Then
        lblAmjMax.ForeColor = errUsr.ForeColor
        DTPicker_Now txtAmjMax
    Else
        valAmjMax = mId$(X, 1, 8)
        If valAmjMax <> mvalAmjMax Then frmCptMvtDisabled
    End If
End If

End Sub

'-------------------------------------------------'
Private Sub txtAmjMin_Control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAmjMin.Year, "0000") & Format$(txtAmjMin.Month, "00") & Format$(txtAmjMin.Day, "00")
lblAmjMin.ForeColor = lblUsr.ForeColor
If Not IsNumeric(X) Then
    lblAmjMin.ForeColor = errUsr.ForeColor
    DTPicker_Now txtAmjMin
Else
    If mId$(X, 1, 8) > DSys Then
        lblAmjMin.ForeColor = errUsr.ForeColor
        DTPicker_Now txtAmjMin
    Else
        valAmjMin = mId$(X, 1, 8)
        If valAmjMin <> mvalAmjMin Then frmCptMvtDisabled
    End If
End If

End Sub


Private Sub chkColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkColor
End Sub


Private Sub chkEffacer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkEffacer

End Sub


'---------------------------------------------------------
Private Sub chkpageparpage_Click()
'---------------------------------------------------------

End Sub

Private Sub chkPageParPage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkPageParPage

End Sub


'---------------------------------------------------------
Private Sub cmdPageNext_Click()
'---------------------------------------------------------

If Amont Then
    PagePrior
Else
    PageNext
End If


End Sub

Private Sub cmdPageNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPageNext

End Sub


'---------------------------------------------------------
Private Sub cmdPagePrior_Click()
'---------------------------------------------------------

If Amont Then
    PageNext
Else
    PagePrior
End If

End Sub


Private Sub cmdPagePrior_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPagePrior

End Sub


'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
optCV_Euro = False
Me.PopupMenu picCptMvtmnu, vbPopupMenuRightButton
'prtListe = True
'prtPréimprimé = False
'cmdPrintX

End Sub


'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub


Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

'---------------------------------------------------------
Private Sub cmdRechercher_Click()
'---------------------------------------------------------

frmCptMvtRechercher

End Sub

Private Sub cmdRechercher_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdRechercher

End Sub


'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
'txtAmjMin.SetFocus
prtListe = True

End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 27: cmdQuit_Click
    Case Is = 34: cmdPageNext_Click
    Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
    Case Is = 13: SendKeys "{TAB}"
End Select
lblErrMsg.Visible = False
shapeCptMvt.Visible = False

End Sub


'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState




Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

CptMvtMethod = "SnapJA"
arrCptMvtNbMax = 1
arrCptMvtNb = 0
ReDim arrCptMvt(1)

Call frmCptMvtEffacer

End Sub






'---------------------------------------------------------
Public Sub frmCptMvtInit(xAmjMin As String, xAmjMax As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate


valAmjMin = xAmjMin
valAmjMax = xAmjMax
Call DTPicker_Set(txtAmjMax, valAmjMax)
If valAmjMin <> "00000000" Then
    Call DTPicker_Set(txtAmjMin, valAmjMin)
    chkAmj = "1"
Else
    Call DTPicker_Set(txtAmjMax, dateElp("Ouvré", -1, DSys))
    txtAmjMax_Control
    Call DTPicker_Set(txtAmjMin, mId$(valAmjMax, 1, 6) & "01")
    chkAmj = "0"
End If
Amont = True
chkPageParPage.Value = IIf(valAmjMin = "00000000", 1, 0)
recCompte = arrCompte(arrCompteIndex)

CV_X1.DeviseN = Format$(recCompte.Devise, "000")
Call CV_AttributN(CV_X1)
picCptMvtmnuImpressionA4_Euro.Enabled = CV_X1.EuroIn
mDevx = CV_X1.DeviseIso

mCptMvt_Obj = "SRVCPTMVT   "
frmCptMvtTête
frmCptMvtRechercher

End Sub

Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub txtAmjMax_Change()
txtAmjMax_Control

End Sub

'-------------------------------------------------'
Private Sub txtAmjMax_GotFocus()
'-------------------------------------------------'
DTPicker_GotFocus txtAmjMax
End Sub

Private Sub txtAmjMax_LostFocus()
DTPicker_LostFocus txtAmjMax
txtAmjMax_Control
End Sub


Private Sub txtAmjMax_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DTPicker_GotFocus txtAmjMax

End Sub


Private Sub txtAmjMin_Change()
txtAmjMin_Control
End Sub

'-------------------------------------------------'
Private Sub txtAmjMin_GotFocus()
'-------------------------------------------------'
DTPicker_GotFocus txtAmjMin
End Sub

'-------------------------------------------------'
Private Sub txtAmjMin_LostFocus()
'-------------------------------------------------'
DTPicker_LostFocus txtAmjMin
txtAmjMin_Control
End Sub


Private Sub txtAmjMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DTPicker_GotFocus txtAmjMin

End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Private Sub picCompte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub picCptMvt_Click()

picCptMvtmnuDétail_Click

End Sub

Private Sub picCptMvt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
shapeCptMvt.Top = ((Y - 20) \ Me_picLineHeight) * Me_picLineHeight - 20
shapeCptMvt.Visible = True
If Amont Then
    arrCptMvtIndex = frmCptMvtNbMax - (Y - 20) \ Me_picLineHeight
Else
    arrCptMvtIndex = frmCptMvtNbMin + (Y - 20) \ Me_picLineHeight
End If
If Button = vbRightButton Then
    Me.PopupMenu picCptMvtmnu, vbPopupMenuRightButton
End If

End Sub

Private Sub picCptMvt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub

'---------------------------------------------------------
Private Sub picCptMvtmnuDétail_Click()
'---------------------------------------------------------
 
frmCptMvtDShow

End Sub


Private Sub picCptMvtmnuImpressionA4_Click()
'---------------------------------------------------------
prtListe = False
prtPréimprimé = False
optCV_Euro = False
cmdPrintX


End Sub

Private Sub picCptMvtmnuImpressionA4_Euro_Click()
prtListe = False
prtPréimprimé = False
optCV_Euro = True
cmdPrintX

End Sub


'---------------------------------------------------------
Private Sub picCptMvtmnuImpressionExtrait_Click()
'---------------------------------------------------------
prtListe = False
prtPréimprimé = True
cmdPrintX

End Sub


'---------------------------------------------------------
Private Sub picCptMvtmnuImpressionListe_Click()
'---------------------------------------------------------
prtListe = True
prtPréimprimé = False
cmdPrintX

End Sub


'---------------------------------------------------------
Private Sub picCptMvtmnuPièce_Click()
'---------------------------------------------------------
MsgBox "non disponible"

End Sub






'-------------------------------------------------'
Public Sub frmCptMvtTête()
'-------------------------------------------------'

If chkColor Then
    picCompte.BackColor = picUsr.BackColor
Else
    picCompte.BackColor = vbWhite
End If

Call recCompte_Display(arrCompte(arrCompteIndex), picCompte, False)
End Sub

'-------------------------------------------------'
Public Sub frmcptmvtEnabledAmont()
'-------------------------------------------------'

cmdPageNext.Caption = "..>>"
If frmCptMvtNbMin > 1 Then
    cmdPageNext.Enabled = True
    cmdPageNext.BackColor = &HC0FFC0
Else
    cmdPageNext.Enabled = False
    cmdPageNext.BackColor = &H8000000F
End If

cmdPagePrior.Caption = "<<..."
If frmCptMvtNbMax < arrCptMvtNb Or arrCptMvtSuite Then
    cmdPagePrior.Enabled = True
    cmdPagePrior.BackColor = &HC0FFC0
Else
    cmdPagePrior.Enabled = False
    cmdPagePrior.BackColor = &H8000000F
End If

End Sub

'-------------------------------------------------'
Public Sub PageNext()
'-------------------------------------------------'

If arrCptMvtIndex < arrCptMvtNb Then
    frmCptMvtItem
Else
    arrCptMvtIndex = arrCptMvtNb
    If chkPageParPage Then
        If arrCptMvtSuite Then frmCptMvtSuite
    Else
        frmCptMvtSuiteEtFin
    End If
End If

End Sub

'-------------------------------------------------'
Public Sub PagePrior()
'-------------------------------------------------'

arrCptMvtIndex = frmCptMvtNbMin - frmCptMvtNb - 1

frmCptMvtItem

End Sub

'-------------------------------------------------'
Public Function valDate()
'-------------------------------------------------'

valDate = "?"

chkAmj_Click
If valAmjMin > valAmjMax Then
    lblErrMsg = "date début > date fin"
    lblErrMsgStd
    Exit Function
'!!!!!!!!!!!
End If
If chkPageParPage Then
Else
    If valAmjMin = "00000000" Then
        lblErrMsg = "Préciser date début"
        lblErrMsgStd
        Exit Function
    '!!!!!!!!!!!
    End If
End If
valDate = Null

End Function

'-------------------------------------------------'
Public Sub frmCptMvtDisabled()
'-------------------------------------------------'

cmdPageNext.Enabled = False
cmdPagePrior.Enabled = False
cmdPrint.Enabled = False
picCptMvt.Cls

End Sub

Public Sub cmdPrintX()
'---------------------------------------------------------
Dim Msg As String, mAmj As String * 8

frmCptMvt.Enabled = False

If arrCptMvtNb > 0 Then
    If Amont Then
        mAmj = arrCptMvt(0).AmjTraitement
        arrCptMvt(0).AmjTraitement = arrCptMvt(arrCptMvtNb).AmjTraitement
        frmCptMvtSuiteEtFin
        arrCptMvt(0).AmjTraitement = mAmj
    Else
        frmCptMvtSuiteEtFin
    End If
    arrCptMvtSuite = True
    Msg = Format$(1, "000000") & Format$(arrCptMvtNb, "000000") & " E " & valAmjMin & valAmjMax
    If Amont Then Mid$(Msg, 13, 1) = "-"
    If prtListe Then Mid$(Msg, 14, 1) = "L"
    If prtPréimprimé Then Mid$(Msg, 14, 1) = "P"
    If optCV_Euro Then Mid$(Msg, 15, 1) = "E"
    prtCptMvt_Monitor Msg
End If
arrCptMvtSuite = True
frmCptMvt.Enabled = True
frmCptMvtEnabled

End Sub



Public Sub frmCptMvtJourAmont()
mCptMvt_Obj = "SRVECRITG": recCptMvt.obj = mCptMvt_Obj: arrCptMvt(0).obj = mCptMvt_Obj
recCptMvt.AmjTraitement = "99999999": arrCptMvt(0).AmjTraitement = "00000000"
If IsNull(srvCptMvtMon(recCptMvt)) Then Call frmCptMvtSuiteEtFin: frmCptMvtItem
arrCptMvtIndex = 0

mCptMvt_Obj = "SRVCPTMVT": recCptMvt.obj = mCptMvt_Obj: arrCptMvt(0).obj = mCptMvt_Obj

frmCptMvtSql

End Sub

Public Sub frmCptMvtJourAval()
If valAmjMax < DSys Then Exit Sub
mCptMvt_Obj = "SRVECRITG": recCptMvt.obj = mCptMvt_Obj: arrCptMvt(0).obj = mCptMvt_Obj
frmCptMvtSql
recCptMvt.AmjTraitement = "00000000": arrCptMvt(0).AmjTraitement = "99999999"
If IsNull(srvCptMvtMon(recCptMvt)) Then Call frmCptMvtSuiteEtFin: arrCptMvtIndex = mIndex

mCptMvt_Obj = "SRVCPTMVT": recCptMvt.obj = mCptMvt_Obj: arrCptMvt(0).obj = mCptMvt_Obj


End Sub
