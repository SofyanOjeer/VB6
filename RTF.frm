VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRTF 
   AutoRedraw      =   -1  'True
   Caption         =   "Courrier"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   Icon            =   "RTF.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
   ScaleWidth      =   9090
   Begin VB.Frame fraRTF 
      Height          =   6165
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   9045
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   8700
         Picture         =   "RTF.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   300
      End
      Begin VB.ListBox lstErr 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6525
         TabIndex        =   2
         Top             =   465
         Width           =   2490
      End
      Begin RichTextLib.RichTextBox txtRTF 
         Height          =   5370
         Left            =   75
         TabIndex        =   1
         Top             =   750
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   9472
         _Version        =   393217
         HideSelection   =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"RTF.frx":0544
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList imglstRTF 
         Left            =   60
         Top             =   660
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RTF.frx":05BB
               Key             =   "Bold"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RTF.frx":06CD
               Key             =   "Italic"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RTF.frx":07DF
               Key             =   "Left"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RTF.frx":08F1
               Key             =   "Center"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RTF.frx":0A03
               Key             =   "Right"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RTF.frx":0B15
               Key             =   "Var"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RTF.frx":0F69
               Key             =   "Tab"
            EndProperty
         EndProperty
      End
      Begin ComCtl3.CoolBar cbRTF 
         Height          =   690
         Left            =   90
         Negotiate       =   -1  'True
         TabIndex        =   4
         Top             =   105
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   1217
         BandCount       =   2
         EmbossPicture   =   -1  'True
         _CBWidth        =   8730
         _CBHeight       =   690
         _Version        =   "6.7.8988"
         MinHeight1      =   330
         Width1          =   3795
         NewRow1         =   0   'False
         Child2          =   "tbRTF_Justification"
         MinHeight2      =   630
         Width2          =   6000
         NewRow2         =   0   'False
         Begin MSComctlLib.Toolbar tbRTF_Justification 
            Height          =   630
            Left            =   3990
            TabIndex        =   7
            Top             =   30
            Width           =   4650
            _ExtentX        =   8202
            _ExtentY        =   1111
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            ImageList       =   "imglstRTF"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Left"
                  Object.ToolTipText     =   "Aligner le paragraphe à gauche"
                  ImageKey        =   "Left"
                  Style           =   2
                  Value           =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Center"
                  Object.ToolTipText     =   "Centrer le paragraphe"
                  ImageKey        =   "Center"
                  Style           =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Right"
                  Object.ToolTipText     =   "Aligner le paragraphe à droite"
                  ImageKey        =   "Right"
                  Style           =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Tab"
                  Object.ToolTipText     =   "Tabulation"
                  ImageKey        =   "Tab"
                  Style           =   2
               EndProperty
            EndProperty
            Begin VB.ComboBox txtRTV 
               Height          =   315
               Left            =   1440
               TabIndex        =   9
               Top             =   60
               Width           =   2235
            End
            Begin VB.CommandButton cmdOK 
               BackColor       =   &H00C0FFC0&
               Caption         =   "OK"
               Height          =   345
               Left            =   3795
               Style           =   1  'Graphical
               TabIndex        =   8
               Top             =   0
               Width           =   500
            End
         End
         Begin MSComctlLib.Toolbar tbRTF_Font 
            Height          =   390
            Left            =   165
            TabIndex        =   5
            Top             =   30
            Width           =   3630
            _ExtentX        =   6403
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            Wrappable       =   0   'False
            Appearance      =   1
            ImageList       =   "imglstRTF"
            HotImageList    =   "imglstRTF"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Combo"
                  Style           =   4
                  Object.Width           =   2500
                  MixedState      =   -1  'True
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   1
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     EndProperty
                  EndProperty
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Bold"
                  Object.ToolTipText     =   "Caractères gras"
                  ImageKey        =   "Bold"
                  Style           =   1
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Italic"
                  Object.ToolTipText     =   "Caractères italiques"
                  ImageKey        =   "Italic"
                  Style           =   1
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Var"
                  Object.ToolTipText     =   "variable publipostage"
                  ImageKey        =   "Var"
                  Style           =   1
               EndProperty
            EndProperty
            Begin VB.ComboBox cboRTF_Police 
               Height          =   315
               Left            =   30
               Sorted          =   -1  'True
               TabIndex        =   10
               Top             =   45
               Width           =   1950
            End
            Begin VB.TextBox txtRTF_Size 
               Height          =   315
               Left            =   2010
               TabIndex        =   6
               Top             =   30
               Width           =   495
            End
         End
      End
   End
End
Attribute VB_Name = "frmRTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim RTFAut As typeAuthorization


Dim blnForm_Load As Boolean


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cboRTF_Police_Load()
Dim I As Integer
cboRTF_Police.ZOrder
For I = 0 To Screen.FontCount - 1
    cboRTF_Police.AddItem Printer.Fonts(I)
Next

End Sub


Private Sub tbRTF_Click(ByVal Button As MSComctlLib.Button)

    ' Supprime l'éventuel état incertain
    If Button.MixedState Then Button.MixedState = False
    
    ' Action selon le bouton
    Select Case Button.key
         Case "Var"
            txtRTF.SelColor = IIf(Button.Value = tbrPressed, vbRed, vbBlack)
       Case "Bold"
            txtRTF.SelBold = IIf(Button.Value = tbrPressed, True, False)
        Case "Italic"
            txtRTF.SelItalic = IIf(Button.Value = tbrPressed, True, False)
        Case "Left"
            txtRTF.SelAlignment = rtfLeft
        Case "Center"
            txtRTF.SelAlignment = rtfCenter
        Case "Right"
            txtRTF.SelAlignment = rtfRight
        Case "Tab"
            txtRTF.SelText = Chr$(9) & txtRTF.SelText
 End Select
    
    ' Rend le focus au texte
    On Error Resume Next
    'txtRTF.SetFocus
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

'Call BiaPgmAut_Init(mId$(Msg, 1, 12), RTFAut)
''''MsgBox "Base locale & \\BIADOC\", vbExclamation

'If mId$(Msg, 1, 12) = "X_DOC$      " And RTFAut.Saisir Then
'        If DataBase_Open <> DataBase_Master Then MDB_Open DataBase_Master, paramDataBase_Password
'End If


'blnSetfocus = True

Form_Init
blnMsgBox_Quit = False
If frmRTF_Caller = "frmEdition  Modèle" Then
'    txtRTV_Load

'    frmRTF.lstErr.Clear
'    frmRTF.lstErr.AddItem "Saisir le document  :"
'    frmRTF.lstErr.AddItem frmRTF_recEdition.Memo1

 '   frmRTF.txtRTF.TextRTF = frmRTF_recEdition.Memo2
 '   cbRTF.Visible = True
 '   cbRTF.Enabled = True
 '   frmRTF.txtRTF.Locked = False
Else
    cbRTF.Visible = False
    If frmRTF_Caller = "frmEdition  Print  " Then cmdPrintX ': Unload Me
End If
DoEvents

End Sub
Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

blnControl = False


cmdReset


End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
'cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
currentAction = ""
blnMsgBox_Quit = False

'========================RTF
cboRTF_Police_Load
txtRTF_SelChange
cbRTF.Enabled = False
frmRTF.txtRTF.Locked = True
blnControl = True
End Sub



Private Sub cmdOk_Click()
If frmRTF_Caller = "frmEdition  Modèle" Then
    blnMsgBox_Quit = False
    frmRTF_recEdition.Memo2 = txtRTF.TextRTF
    frmRTF_blnOK = True
    Unload Me
End If
End Sub

Private Sub cmdPrint_Click()
cmdPrintX
End Sub

Private Sub txtRTF_Change()
If Not txtRTF.Locked Then blnMsgBox_Quit = True

End Sub

Private Sub txtRTF_SelChange()
    ' Nouvelle valeur
    Dim nVal As Integer
    ' Bouton publipotage
    Dim B As Variant 'Button
    Set B = tbRTF_Font.Buttons("Var")
    If IsNull(txtRTF.SelColor) Then
        If Not B.MixedState Then B.MixedState = True
    Else
        If B.MixedState Then B.MixedState = False
        nVal = IIf(txtRTF.SelColor, tbrPressed, tbrUnpressed)
        If B.Value <> nVal Then B.Value = nVal
    End If
    
    ' Bouton gras
    Set B = tbRTF_Font.Buttons("Bold")
    If IsNull(txtRTF.SelBold) Then
        If Not B.MixedState Then B.MixedState = True
    Else
        If B.MixedState Then B.MixedState = False
        nVal = IIf(txtRTF.SelBold, tbrPressed, tbrUnpressed)
        If B.Value <> nVal Then B.Value = nVal
    End If
    
    ' Bouton italique
    Set B = tbRTF_Font.Buttons("Italic")
    If IsNull(txtRTF.SelItalic) Then
        If Not B.MixedState Then B.MixedState = True
    Else
        If B.MixedState Then B.MixedState = False
        nVal = IIf(txtRTF.SelItalic, tbrPressed, tbrUnpressed)
        If B.Value <> nVal Then B.Value = nVal
    End If

    ' Boutons d'alignement
    Dim key As String
    Select Case txtRTF.SelAlignment
        Case rtfLeft
            key = "Left"
        Case rtfCenter
            key = "Center"
        Case rtfRight
            key = "Right"
        Case Else
            key = ""
    End Select
    
    If key = "" Then
        If tbRTF_Justification.Buttons("Left").Value <> tbrUnpressed Then tbRTF_Justification.Buttons("Left").Value = tbrUnpressed
        If tbRTF_Justification.Buttons("Right").Value <> tbrUnpressed Then tbRTF_Justification.Buttons("Right").Value = tbrUnpressed
        If tbRTF_Justification.Buttons("Center").Value <> tbrUnpressed Then tbRTF_Justification.Buttons("Center").Value = tbrUnpressed
    Else
        If tbRTF_Justification.Buttons(key).Value <> tbrPressed Then tbRTF_Justification.Buttons(key).Value = tbrPressed
    End If
    
    ' Police de caractères
    On Error Resume Next    ' Si sélection multiple
    cboRTF_Police.Text = txtRTF.SelFontName
    If Err Then cboRTF_Police.Text = ""
    txtRTF_Size = CStr(txtRTF.SelFontSize)
    If Err Then txtRTF_Size = ""

End Sub

Private Sub cboRTF_Police_Click()
txtRTF.SelFontName = cboRTF_Police.Text

txtRTF.SetFocus

End Sub


Private Sub tbRTF_Font_ButtonClick(ByVal Button As MSComctlLib.Button)
tbRTF_Click Button

End Sub


Private Sub tbRTF_Justification_ButtonClick(ByVal Button As MSComctlLib.Button)
tbRTF_Click Button

End Sub


Private Sub txtRTF_Size_GotFocus()
txt_GotFocus txtRTF_Size

End Sub

Private Sub txtRTF_Size_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtRTF_Size_Ok
End Sub

Private Sub txtRTF_Size_Ok()
    ' Vérifie la valeur
    Dim sz As Integer
    sz = Val(txtRTF_Size)
    If sz < 1 Then sz = 1
    If sz > 2000 Then sz = 2000
    txtRTF_Size = CStr(sz)
    ' Change la taille
    txtRTF.SelFontSize = sz
    On Error Resume Next
    txtRTF.SetFocus
End Sub

Private Sub txtRTF_Size_LostFocus()
txtRTF_Size_Ok
txt_LostFocus txtRTF_Size

End Sub


'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
'frm_Control
'C.ForeColor = txtUsr.ForeColor
'C.BackColor = focusUsr.BackColor
'currentActiveControl_Name = C.Name
End Sub


'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
'arrTag(Val(C.Tag)) = True
'C.ForeColor = txtUsr.ForeColor
'C.BackColor = txtUsr.BackColor
End Sub

'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select


End Sub

Public Sub cmdContext_Quit()
blnControl = False
If Not blnMsgBox_Quit Then
    Unload Me
Else
    If Not Form_QueryUnload_Msgbox Then Unload Me
End If
End Sub

Public Sub cmdContext_Return()
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
'If Not blnForm_Load Then
'    blnForm_Load = True
    mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
    Set XForm = Me
    Call MeInit(arrTagNb)
    ReDim arrTag(arrTagNb + 1)
    blnControl = False
    Msg_Rcv " "
'End If
End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = Form_QueryUnload_Msgbox

'blnForm_Load = False
End Sub

Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

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
            If TypeOf C Is ListBox Then
                Elp_ResizeControl C
            Else
                MouseMoveActiveControl.ForeColor = C.ForeColor
                C.ForeColor = MouseMoveUsr.ForeColor
            End If
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
            If TypeOf xobj Is ListBox Then
                xobj.Height = 200
            Else
                xobj.ForeColor = MouseMoveActiveControl.ForeColor
            End If
        End If
        Exit For
    End If
Next xobj
End Sub


Public Sub txt_X()
'Call txt_GotFocus(txt)
'KeyAscii = convUCase(KeyAscii)
'Call txt_LostFocus(txt)

'Call txt_GotFocus(txt)
'If XopDevise(2).maxD = 0 Then
'    Call num_KeyAscii(KeyAscii)
'Else
'    Call num_KeyAsciiD(KeyAscii, txt)
'End If
'Call txt_LostFocus(txt)

End Sub

Public Function Form_QueryUnload_Msgbox() As Boolean
Dim X As String
If blnMsgBox_Quit Then
    X = MsgBox("Voulez-vous réellement abandonner la mise à jour?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)
    Form_QueryUnload_Msgbox = IIf(X = vbNo, True, False)
End If
End Function

Private Sub txtRTV_Click()
txtRTF.SelColor = vbRed
txtRTF.SelText = Trim(txtRTV.Text)
txtRTF.SelColor = tbrUnpressed
txtRTF.SelColor = vbBlack
End Sub



Public Sub cmdPrintX()

If frmRTF_Caller = "srvDSPFFDY0_frmRTF" Then
    frmRTF.Hide
    MsgBox "à faire : srvDSPFFDY0_frmRTF_Print"
    Unload Me
End If

prtPaperSize = frmRTF_prtPaperSize

If frmRTF_blnCourrier Then
    prtOrientation = frmRTF_prtOrientation
    prtCourrier_Open
    prtMinX = prtMinX + 300
    prtMaxX = prtMaxX - 1000
    XPrt.CurrentX = prtMinX
    Call frmElpPrt.prtRTF(txtRTF.TextRTF)
Else
    prtOrientation = frmRTF_prtOrientation
    prtEdition_Open
    XPrt.CurrentX = prtMinX
    Call frmElpPrt.prtRTF(txtRTF.TextRTF)
End If

''If Trim(Me.Tag) <> "" Then
'    Call frmEdition.ecrit_LogPrintings(Me.Tag, XPrt.Devicename) 'DRDR
''End If

prtEdition_Close
frmRTF_blnA5 = False
If frmRTF_Caller = "frmEdition  Display" Then Unload Me

End Sub
