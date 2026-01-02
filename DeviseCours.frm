VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDeviseCours 
   AutoRedraw      =   -1  'True
   Caption         =   "Guichet : cours de change"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   400
      Left            =   8880
      Picture         =   "DeviseCours.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Valider"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid fgCours 
      Height          =   2850
      Left            =   3360
      TabIndex        =   17
      Top             =   480
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   5027
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedCols       =   0
      BackColor       =   14737632
      ForeColor       =   12582912
      ForeColorFixed  =   -2147483641
      BackColorSel    =   12648384
      BackColorBkg    =   14737632
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   1
      FormatString    =   ">Unité      | Dev1    | Dev2     |> Cours Pivot  |> Achat C     |>Vente C      |>        Etat"
   End
   Begin VB.Frame fraCours 
      Caption         =   "Cours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   18
      Top             =   3480
      Width           =   9375
      Begin VB.TextBox fraCours_txtCoursPivot 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Frame fraCours_fraNormal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   3960
         TabIndex        =   25
         Top             =   360
         Width           =   5175
         Begin VB.TextBox fraCours_txtAchatNormal 
            Height          =   285
            Left            =   1080
            TabIndex        =   5
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox fraCours_txtVenteNormal 
            Height          =   285
            Left            =   3120
            TabIndex        =   8
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox fraCours_txtAchatPrivilégié 
            Height          =   285
            Left            =   1080
            TabIndex        =   6
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox fraCours_txtVentePrivilégié 
            Height          =   285
            Left            =   3120
            TabIndex        =   9
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox fraCours_txtVenteEnCompte 
            Height          =   285
            Left            =   3120
            TabIndex        =   7
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox fraCours_txtAchatEnCompte 
            Height          =   285
            Left            =   1080
            TabIndex        =   4
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblBilletsPrivilégié 
            Caption         =   "Billets(P)"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label lblBillets 
            Caption         =   "Billets"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lblEnCompte 
            Caption         =   "en Compte"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblAchat 
            Caption         =   "Achat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   33
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblVente 
            Caption         =   "Vente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   32
            Top             =   240
            Width           =   615
         End
         Begin VB.Label libAchatEnCompte 
            Caption         =   "-"
            Height          =   255
            Left            =   2280
            TabIndex        =   31
            Top             =   720
            Width           =   735
         End
         Begin VB.Label libAchatNormal 
            Caption         =   "-"
            Height          =   255
            Left            =   2280
            TabIndex        =   30
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label libAchatPrivilégié 
            Caption         =   "-"
            Height          =   255
            Left            =   2280
            TabIndex        =   29
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label libVenteEnCompte 
            Caption         =   "-"
            Height          =   255
            Left            =   4320
            TabIndex        =   28
            Top             =   600
            Width           =   735
         End
         Begin VB.Label libVenteNormal 
            Caption         =   "-"
            Height          =   255
            Left            =   4320
            TabIndex        =   27
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label libVentePrivilégié 
            Caption         =   "-"
            Height          =   255
            Left            =   4320
            TabIndex        =   26
            Top             =   1560
            Width           =   735
         End
      End
      Begin VB.CommandButton fraCours_cmdQuit 
         BackColor       =   &H00C0C0FF&
         Caption         =   "X"
         Height          =   525
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1920
         Width           =   1500
      End
      Begin VB.CommandButton fraCours_cmdOk 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ok"
         Height          =   525
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1920
         Width           =   1500
      End
      Begin VB.TextBox fraCours_txtDevise2 
         Height          =   285
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox fraCours_txtDevise1 
         Height          =   285
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox fraCours_txtUnité 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblCoursPivot 
         Caption         =   "Cours Pivot"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label libValidation 
         Caption         =   "-"
         Height          =   255
         Left            =   4080
         TabIndex        =   23
         Top             =   2520
         Width           =   4095
      End
      Begin VB.Label libSaisie 
         Caption         =   "-"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label fraCours_lblDevise2 
         Caption         =   "Devise2"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label fraCours_lblDevise1 
         Caption         =   "Devise1"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.Label fraCours_lblUnité 
         Caption         =   "Unité"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.ListBox lstDevise 
      Height          =   2790
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Recherche"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Enregistrer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   15
      Top             =   0
      Width           =   2500
   End
   Begin VB.Menu mnuDevise 
      Caption         =   "Devise"
      Visible         =   0   'False
      Begin VB.Menu mnuDeviseDisplay 
         Caption         =   "Détail Devise"
      End
   End
   Begin VB.Menu mnuCours 
      Caption         =   "Cours"
      Visible         =   0   'False
      Begin VB.Menu mnuCoursAddNew 
         Caption         =   "Ajouter un Cours"
      End
      Begin VB.Menu mnuCoursUpdate 
         Caption         =   "Modifier un Cours"
      End
      Begin VB.Menu mnuCoursDelete 
         Caption         =   "Supprimer un Cours"
      End
      Begin VB.Menu mnuCoursPrint 
         Caption         =   "Impression"
      End
   End
End
Attribute VB_Name = "frmDeviseCours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean
Dim DeviseCoursAut As typeAuthorization
Dim X As String, X1 As String, I As Integer
Dim Msg As String, valX As String

Dim recDevise1 As typeDevise, recDevise2 As typeDevise
Dim recDeviseCours As typeDeviseCours
Dim currentMethod As String, currentAMJ As String
Dim fgCours_FormatString As String, fgCours_K As Integer
Dim fgCours_BackColorFixed As Long, fgCours_BackColor As Long
Dim blnAddNew As Boolean
Dim dblX As Double
Dim CV As typeCV
'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
cmdControl
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
currentActiveControl_Name = C.Name
End Sub
'---------------------------------------------------------
Public Sub lstDevise_Display()
'---------------------------------------------------------
lstDevise.Visible = True
Call LstDictio(889, lstDevise)
End Sub

'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub



Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdOk_Click()
Dim XTest As String, XOk As String
Dim I As Integer
fraCours.Visible = False
If cmdOk.Caption = constValider Then
    XTest = constàValider
    XOk = usrId
Else
    XOk = constàValider
    XTest = ""
End If

TSys = time_Hms
For I = 1 To arrDeviseCoursNb
    If Trim(arrDeviseCours(I).ValidationUsr) = XTest Then
        If Trim(arrDeviseCours(I).Method) = "" Then arrDeviseCours(I).Method = constUpdate
        arrDeviseCours(I).ValidationAMJ = DSys
        arrDeviseCours(I).ValidationHMS = TSys
        arrDeviseCours(I).ValidationUsr = XOk
        If Trim(arrDeviseCours(I).Method) <> constIgnore Then srvDeviseCours_Update arrDeviseCours(I)
    End If
Next I
blnMsgBox_Quit = False
If cmdOk.Caption = constValider Then
    cmdPrintX " : Validation"
Else
    cmdPrintX " : liste de contrôle à valider"
End If
cmdContext_Quit
End Sub

'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
cmdPrintX ""
End Sub

Private Sub fraCours_txtDevise1_Control()
If Trim(fraCours_txtDevise1) = "" Then Call lstErr_AddItem(lstErr, fraCours_txtDevise1, "? Devise 1"): Exit Sub
If DevCode(fraCours_txtDevise1) = 0 Then Call lstErr_AddItem(lstErr, fraCours_txtDevise1, "?Devise1 inconnue"): Exit Sub
recDevise1 = XDevise
recDeviseCours.Id1 = recDevise1.DevX
'fraCours_txtDevise2.Enabled = True
'fraCours_fraNormal.Enabled = True: fraCours_fraPrivilégié.Enabled = True

CV_Init CV
CV.DeviseIso = recDevise1.DevX
CV_Attribut CV
If CV.EuroIn Then Call lstErr_AddItem(lstErr, fraCours_txtDevise1, "?Devise1 IN"): Exit Sub

End Sub





Private Sub cmdSave_Click()
Dim I As Integer
fraCours.Visible = False
For I = 1 To arrDeviseCoursNb
    If Trim(arrDeviseCours(I).ValidationUsr) = constàValider Then
        arrDeviseCours(I).Method = constUpdate
        arrDeviseCours(I).ValidationAMJ = "00000000"
        arrDeviseCours(I).ValidationHMS = "000000"
        arrDeviseCours(I).ValidationUsr = Space$(10)
    End If
    
    If Trim(arrDeviseCours(I).Method) <> "" Then
        If Trim(arrDeviseCours(I).Method) <> constIgnore Then srvDeviseCours_Update arrDeviseCours(I)
    End If
Next I
blnMsgBox_Quit = False
cmdContext_Quit

End Sub

Private Sub fgCours_Click()
lstErr.Clear
fgCours_K = fgCours.Row * fgCours.Cols
If fgCours.Row > 0 Then lstDevise_Scan Mid$(Trim(fgCours.TextArray(1 + fgCours_K)), 1, 3)
'fgCours.Col = 1: fgCours.CellBackColor = focusUsr.BackColor
Me.PopupMenu mnuCours, vbPopupMenuRightButton
'fgCours.Col = 1: fgCours.CellBackColor = fgCours_BackColor
End Sub

Private Sub fgCours_GotFocus()
fgCours.BackColorFixed = focusUsr.BackColor
fgCours.BackColor = fgCours_BackColor
End Sub


Private Sub fgCours_LostFocus()
fgCours.BackColorFixed = fgCours_BackColorFixed
'fgCours.BackColor = vbWindowBackground
End Sub


'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
    Case Is = 44: frmElpPrt.prtScreen
End Select

End Sub


'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Dim Amj As String, X As String
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
Call srvUsrAut("G_Change", DeviseCoursAut)
cmdClear
recDeviseCours_Init recDeviseCours
currentAMJ = DSys: Amj = DSys
lstDevise_Display
DevX "001": recDevise1 = XDevise: recDevise2 = XDevise
fgCours_FormatString = fgCours.FormatString
fgCours_BackColorFixed = fgCours.BackColorFixed
fgCours_BackColor = fgCours.BackColor
If DeviseCoursAut.Détail Then
    blnAddNew = True
    Do
        fgCours_Load Amj, blnAddNew
        If arrDeviseCoursNb < 1 Then
            blnAddNew = False
            Amj = dateElp("Ouvré", -1, Amj)
            X = MsgBox("Voulez-vous charger les cours du " & dateImp(Amj), vbYesNo + vbQuestion + vbDefaultButton2, "Pas de cours du jour")
            If X = vbNo Then mnuCoursAddNew_Click: Exit Do
        End If
    Loop While arrDeviseCoursNb < 1
End If


End Sub



'---------------------------------------------------------
Public Sub cmdClear()
'---------------------------------------------------------
cmdReset
cmdContext.Enabled = True: cmdContext.BackColor = vbWindowBackground
cmdContext.Caption = constcmdAbandonner: cmdContext.BackColor = errUsr.BackColor
cmdSave.Visible = False: cmdOk.Visible = False
fgCours.Enabled = True: fgCours.Clear: fgCours.Rows = 1
fraCours.Visible = False
Call lstErr_Clear(lstErr, cmdContext, "choisir un cours 'click'")
lstDevise.Enabled = False: lstDevise.BackColor = vbWindowBackground
fraCours_Clear
End Sub




'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
arrTag_Set False
lstErrClear = True
blnMsgBox_Quit = False
cmdOk.Visible = False
usrColor_Set
End Sub

'---------------------------------------------------------
Public Sub cmdValidation()
'---------------------------------------------------------
cmdControl
lstErrClear = False
End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------

End Sub

Public Sub Msg_Snd(ByVal X As String)
Msg = Space$(100)

Mid$(Msg, 1, 12) = "frm         "
Mid$(Msg, 13, 12) = Me.Name
Mid$(Msg, 25, 10) = Space$(10)
Mid$(Msg, 35, 10) = X
Mid$(Msg, 45, 6) = "      "
Msg_Monitor Msg
End Sub

Public Sub cmdControl()
Dim I As Integer, xobjControl As Boolean
If lstErrClear Then lstErr.Clear

For I = 0 To arrTagNb
    If arrTag(I) Then
        xobjControl = False
        arrTag(I) = False
        X = Format(I, "000")
        For Each xobj In Me.Controls
            If X = xobj.Tag Then
                Select Case xobj.Name
                    Case "fraCours_txtDevise1": fraCours_txtDevise1_Control: xobjControl = True
                    Case "fraCours_txtDevise2": fraCours_txtDevise2_Control: xobjControl = True
                    Case "fraCours_txtCoursPivot": fraCours_txtCoursPivot_Control: xobjControl = True
                    Case "fraCours_txtUnité": fraCours_txtUnité_Control: xobjControl = True
                    Case "fraCours_txtAchatEnCompte": fraCours_txtAchatEnCompte_Control: xobjControl = True
                    Case "fraCours_txtVenteEnCompte": fraCours_txtVenteEnCompte_Control: xobjControl = True
                    Case "fraCours_txtAchatNormal": fraCours_txtAchatNormal_Control: xobjControl = True
                    Case "fraCours_txtVenteNormal": fraCours_txtVenteNormal_Control: xobjControl = True
                    Case "fraCours_txtAchatPrivilégié": fraCours_txtAchatPrivilégié_Control: xobjControl = True
                    Case "fraCours_txtVentePrivilégié": fraCours_txtVentePrivilégié_Control: xobjControl = True
                End Select
            End If
            If xobjControl Then Exit For
        Next xobj
    End If
Next I

lstErrClear = True
End Sub


Private Sub fraCours_cmdOk_Click()
fraCours_Control
If lstErr.ListCount > 0 Then Exit Sub

cmdOk.Visible = DeviseCoursAut.Valider
blnMsgBox_Quit = True
recDeviseCours.SaisieAMJ = DSys
recDeviseCours.SaisieHMS = time_Hms
recDeviseCours.SaisieUsr = usrId

Select Case currentMethod
    Case constAddNew: fgCours_AddNew
    Case constUpdate: fgCours_Update
    Case constDelete: fgCours_Delete
End Select
If lstErr.ListCount > 0 Then Exit Sub
fgCours_Sort
fraCours_Exit
End Sub

Private Sub fraCours_cmdOk_GotFocus()
cmdControl
End Sub

Private Sub fraCours_cmdQuit_Click()
blnAddNew = False
fraCours_Exit
End Sub

Private Sub fraCours_txtAchatEnCompte_GotFocus()
Call txt_GotFocus(fraCours_txtAchatEnCompte)
End Sub


Private Sub fraCours_txtAchatEnCompte_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, fraCours_txtAchatEnCompte)
End Sub


Private Sub fraCours_txtAchatEnCompte_LostFocus()
Call txt_LostFocus(fraCours_txtAchatEnCompte)
End Sub

Private Sub fraCours_txtAchatNormal_GotFocus()
Call txt_GotFocus(fraCours_txtAchatNormal)
End Sub

Private Sub fraCours_txtAchatNormal_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, fraCours_txtAchatNormal)
End Sub


Private Sub fraCours_txtAchatNormal_LostFocus()
Call txt_LostFocus(fraCours_txtAchatNormal)
End Sub

Private Sub fraCours_txtAchatPrivilégié_GotFocus()
Call txt_GotFocus(fraCours_txtAchatPrivilégié)
End Sub

Private Sub fraCours_txtAchatPrivilégié_KeyPress(KeyAscii As Integer)
    Call num_KeyAsciiD(KeyAscii, fraCours_txtAchatPrivilégié)
End Sub


Private Sub fraCours_txtAchatPrivilégié_LostFocus()
Call txt_LostFocus(fraCours_txtAchatPrivilégié)
End Sub

Private Sub fraCours_txtCoursPivot_GotFocus()
Call txt_GotFocus(fraCours_txtCoursPivot)
End Sub


Private Sub fraCours_txtCoursPivot_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, fraCours_txtCoursPivot)
End Sub


Private Sub fraCours_txtCoursPivot_LostFocus()
Call txt_LostFocus(fraCours_txtCoursPivot)
End Sub

Private Sub fraCours_txtDevise1_GotFocus()
Call txt_GotFocus(fraCours_txtDevise1)
lstDevise.Enabled = True: lstDevise.BackColor = lstUsr.BackColor
End Sub

Private Sub fraCours_txtDevise1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub fraCours_txtDevise1_LostFocus()
Call txt_LostFocus(fraCours_txtDevise1)
lstDevise.Enabled = False: lstDevise.BackColor = vbWindowBackground
End Sub

Private Sub fraCours_txtDevise2_GotFocus()
Call txt_GotFocus(fraCours_txtDevise2)
lstDevise.Enabled = True: lstDevise.BackColor = lstUsr.BackColor
End Sub

Private Sub fraCours_txtDevise2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub fraCours_txtDevise2_LostFocus()
Call txt_LostFocus(fraCours_txtDevise2)
lstDevise.Enabled = False: lstDevise.BackColor = vbWindowBackground
End Sub

Private Sub fraCours_txtUnité_GotFocus()
Call txt_GotFocus(fraCours_txtUnité)
End Sub

Private Sub fraCours_txtUnité_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub fraCours_txtUnité_LostFocus()
Call txt_LostFocus(fraCours_txtUnité)
End Sub

Private Sub fraCours_txtVenteEnCompte_GotFocus()
Call txt_GotFocus(fraCours_txtVenteEnCompte)
End Sub


Private Sub fraCours_txtVenteEnCompte_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, fraCours_txtVenteEnCompte)

End Sub


Private Sub fraCours_txtVenteEnCompte_LostFocus()
Call txt_LostFocus(fraCours_txtVenteEnCompte)
End Sub


Private Sub fraCours_txtVenteNormal_GotFocus()
Call txt_GotFocus(fraCours_txtVenteNormal)
End Sub

Private Sub fraCours_txtVenteNormal_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, fraCours_txtVenteNormal)
End Sub


Private Sub fraCours_txtVenteNormal_LostFocus()
Call txt_LostFocus(fraCours_txtVenteNormal)
End Sub

Private Sub fraCours_txtVentePrivilégié_GotFocus()
Call txt_GotFocus(fraCours_txtVentePrivilégié)
End Sub

Private Sub fraCours_txtVentePrivilégié_KeyPress(KeyAscii As Integer)
    Call num_KeyAsciiD(KeyAscii, fraCours_txtVentePrivilégié)
End Sub


Private Sub fraCours_txtVentePrivilégié_LostFocus()
Call txt_LostFocus(fraCours_txtVentePrivilégié)
End Sub

Private Sub lstDevise_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case currentActiveControl_Name
    Case "fraCours_txtDevise1": fraCours_txtDevise1 = Mid$(lstDevise.Text, 1, 3): fraCours_txtDevise1.SetFocus
    Case "fraCours_txtDevise2": fraCours_txtDevise2 = Mid$(lstDevise.Text, 1, 3): fraCours_txtDevise2.SetFocus
End Select
SendKeys "{TAB}"

End Sub

Private Sub mnuCoursAddNew_Click()
blnAddNew = True
fgCours.Enabled = False
fraCours_Clear
fraCours.Enabled = True: fraCours.Visible = True
currentMethod = constAddNew
Call lstErr_Clear(lstErr, fraCours_txtCoursPivot, "Nouveau Cours")
recDeviseCours_Init recDeviseCours
fraCours_cmdOk.Visible = DeviseCoursAut.Saisir
fraCours_txtCoursPivot_GotFocus
fraCours_txtCoursPivot.SetFocus
End Sub

Private Sub mnuCoursDelete_Click()
fgCours_Scan
If lstErr.ListCount > 0 Then Exit Sub
If recDeviseCours.ValidationUsr <> "          " Then Exit Sub
If recDeviseCours.Method = constAddNew Then
    currentMethod = constIgnore
    fgCours_Delete
Else
    fraCours.Enabled = False: fraCours.Visible = True
    currentMethod = constDelete
    Call lstErr_Clear(lstErr, fraCours_txtUnité, "Suppression ligne")
    X = MsgBox("Voulez-vous réellement supprimer cette ligne ?", vbYesNo + vbQuestion + vbDefaultButton2, "ancienne ligne")
    If X = vbYes Then fgCours_Delete
    fraCours.Visible = False
    fgCours.Enabled = True
    fgCours.SetFocus
End If
End Sub

Private Sub fgCours_Load(Amj As String, blnZ As Boolean)
Dim blnValidation As Boolean, blnSaisie As Boolean, X As String
lstDevise.Enabled = False: lstDevise.BackColor = vbWindowBackground
srvDeviseCours_Load Amj
blnValidation = False: blnSaisie = False
For arrDeviseCoursIndex = 1 To arrDeviseCoursNb
    If blnZ Then
        arrDeviseCours(arrDeviseCoursIndex).Method = ""
    Else
        arrDeviseCours(arrDeviseCoursIndex).Method = constAddNew
        arrDeviseCours(arrDeviseCoursIndex).Amj = currentAMJ
        arrDeviseCours(arrDeviseCoursIndex).SaisieAMJ = DSys
        arrDeviseCours(arrDeviseCoursIndex).SaisieHMS = time_Hms
        arrDeviseCours(arrDeviseCoursIndex).SaisieUsr = usrId
        arrDeviseCours(arrDeviseCoursIndex).ValidationAMJ = "00000000"
        arrDeviseCours(arrDeviseCoursIndex).ValidationHMS = "000000"
        arrDeviseCours(arrDeviseCoursIndex).ValidationUsr = Space$(10)
     
    End If
    Select Case Trim(arrDeviseCours(arrDeviseCoursIndex).ValidationUsr)
        Case constàValider: blnValidation = True
        Case "": blnSaisie = True
    End Select
    
Next arrDeviseCoursIndex
If blnValidation And Not blnSaisie Then
    X = "Demande de validation"
    cmdSave.Caption = constàModifier
    cmdOk.Caption = constValider
    cmdSave.Visible = DeviseCoursAut.Valider
    cmdOk.Visible = DeviseCoursAut.Valider
    mnuCoursAddNew.Enabled = False
    mnuCoursDelete.Enabled = False
Else
    X = "Saisie en cours"
    cmdSave.Caption = constcmdEnregistrer
    cmdOk.Caption = constàValider
    cmdSave.Visible = DeviseCoursAut.Saisir
    cmdOk.Visible = DeviseCoursAut.Saisir
    mnuCoursAddNew.Enabled = True
    mnuCoursDelete.Enabled = True
End If

fgCours_Display
Call lstErr_Clear(lstErr, fgCours, X)

End Sub

Private Sub mnuCoursUpdate_Click()
fgCours_Scan
If lstErr.ListCount > 0 Then Exit Sub
fgCours.Enabled = False
fraCours_txtDevise1.Enabled = False: fraCours_txtDevise2.Enabled = False
fraCours.Enabled = True: fraCours.Visible = True
currentMethod = constUpdate
Call lstErr_Clear(lstErr, fraCours_txtUnité, "Modification Cours")
lastActiveControl_Name = "fraCours_txtVentePrivilégié"
fraCours_txtCoursPivot.SetFocus
End Sub

Private Sub mnuDeviseDisplay_Click()
If lstDevise.ListIndex >= 0 Then
    Set XListBox = frmDeviseCours.lstDevise
    frmDevise.Show vbModal
End If
End Sub

Private Sub mnuDevisePrint_Click()
Dim Msg As String

Msg = Format$(1, "000000") & Format$(999, "000000")

prtDeviseX Msg

End Sub



Public Sub fgCours_Display()
fgCours.Rows = 1
fgCours.Clear
fgCours.FormatString = fgCours_FormatString
fgCours.Enabled = True
For arrDeviseCoursIndex = 1 To arrDeviseCoursNb
    If arrDeviseCours(arrDeviseCoursIndex).Method <> constDelete _
    And arrDeviseCours(arrDeviseCoursIndex).Method <> constIgnore Then
        fgCours.Rows = fgCours.Rows + 1
        fgCours.Row = fgCours.Rows - 1
        fgCours_DisplayItem
    End If
Next arrDeviseCoursIndex
If fgCours.Rows > 1 Then fgCours_Sort

End Sub

Public Sub fgCours_DisplayItem()
fgCours_K = (fgCours.Row) * fgCours.Cols
fgCours.TextArray(1 + fgCours_K) = Format$(arrDeviseCours(arrDeviseCoursIndex).Id1, "@@@") & " / "
fgCours.TextArray(0 + fgCours_K) = Format$(arrDeviseCours(arrDeviseCoursIndex).QD1, "##########")
fgCours.TextArray(2 + fgCours_K) = Format$(arrDeviseCours(arrDeviseCoursIndex).Id2, "@@@")
'x = IIf(recDevise2.maxD = 0, "#########", "######.00")
X = "######.#####"
fgCours.TextArray(5 + fgCours_K) = Format$(arrDeviseCours(arrDeviseCoursIndex).QD2VenteEnCompte, X)
fgCours.TextArray(4 + fgCours_K) = Format$(arrDeviseCours(arrDeviseCoursIndex).QD2AchatEnCompte, X)
fgCours.TextArray(3 + fgCours_K) = Format$(arrDeviseCours(arrDeviseCoursIndex).QD2CoursPivot, X)
fgCours.TextArray(6 + fgCours_K) = Trim(arrDeviseCours(arrDeviseCoursIndex).ValidationUsr)
Select Case arrDeviseCours(arrDeviseCoursIndex).Method
    Case constAddNew: fgCours.TextArray(6 + fgCours_K) = "Créé"
    Case constUpdate: fgCours.TextArray(6 + fgCours_K) = "Modifié"
End Select
End Sub

Public Sub fgCours_AddNew()
X = num_Control(fraCours_txtUnité, valX, 7, 0)
If arrDeviseCours_ScanId1Id2(recDeviseCours) > 0 Then
    Call lstErr_AddItem(lstErr, fraCours_txtUnité, "Existe déjà")
Else
    recDeviseCours.Method = constAddNew
    recDeviseCours.Amj = currentAMJ
    recDeviseCours.ValidationAMJ = "00000000"
    recDeviseCours.ValidationHMS = "000000"
    fraCours_txtDevise1 = ""
    Call arrDeviseCours_AddItem(recDeviseCours)
    arrDeviseCoursIndex = arrDeviseCoursNb
    fgCours.Rows = fgCours.Rows + 1
    fgCours.Row = fgCours.Rows - 1
    fgCours_DisplayItem
End If
End Sub

Public Sub fgCours_Update()
If recDeviseCours.Method <> constAddNew Then recDeviseCours.Method = constUpdate
arrDeviseCours(arrDeviseCoursIndex) = recDeviseCours
fgCours_DisplayItem
End Sub

Public Sub fgCours_Delete()
recDeviseCours.Method = currentMethod
arrDeviseCours(arrDeviseCoursIndex) = recDeviseCours
fgCours_Display
End Sub

Public Sub fraCours_Clear()
lstErr.Clear
usrColor_Set
libAchatEnCompte = ""
libAchatNormal = ""
libAchatPrivilégié = ""
libVenteEnCompte = ""
libVenteNormal = ""
libVentePrivilégié = ""

fraCours_txtDevise1 = ""
fraCours_txtUnité = "1"
fraCours_txtDevise1 = "EUR"
fraCours_txtDevise2 = ""
fraCours_txtCoursPivot = ""
fraCours_txtAchatEnCompte = "": fraCours_txtVenteEnCompte = ""
fraCours_txtAchatNormal = "": fraCours_txtVenteNormal = ""
fraCours_txtAchatPrivilégié = "": fraCours_txtVentePrivilégié = ""
libSaisie = "": libValidation = ""
fraCours.Enabled = True
fraCours_txtDevise1.Enabled = True: fraCours_txtDevise2.Enabled = True
lastActiveControl_Name = "fraCours_txtVentePrivilégié"
End Sub

Public Sub fraCours_Enabled(ByVal bln As Boolean)
fgCours.Enabled = bln
fraCours_cmdOk.Visible = bln
End Sub

Public Sub cmdContext_Quit()
If fraCours.Visible Then
    fraCours_cmdQuit_Click
Else
    If blnMsgBox_Quit Then
        X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
     Else
        X = vbYes
     End If
     If X = vbYes Then Unload Me
End If

End Sub

Public Sub cmdContext_Return()
If fraCours.Enabled Then
    If ActiveControl.Name = lastActiveControl_Name Then
        fraCours_cmdOk_Click
    Else
        SendKeys "{TAB}"
    End If
Else
    SendKeys "{TAB}"
End If

End Sub

Public Sub fgCours_Scan()
fgCours_K = fgCours.Row * fgCours.Cols
recDeviseCours.Id1 = Trim(fgCours.TextArray(1 + fgCours_K))
recDeviseCours.Id2 = Trim(fgCours.TextArray(2 + fgCours_K))
If arrDeviseCours_ScanId1Id2(recDeviseCours) > 0 Then
    recDeviseCours = arrDeviseCours(arrDeviseCoursIndex)
    fraCours_DisplayItem
Else
    Call lstErr_AddItem(lstErr, fgCours, "Erreur fgCours_Scan")
End If
End Sub

Public Sub fraCours_DisplayItem()
usrColor_Set
libSaisie.ForeColor = vbMagenta: libValidation.ForeColor = vbMagenta

fraCours_txtDevise1 = recDeviseCours.Id1
fraCours_txtDevise2 = recDeviseCours.Id2
fraCours_txtUnité = num_Display(recDeviseCours.QD1, 7, 0, Lx, X1, "#")
fraCours_txtCoursPivot = num_Display(recDeviseCours.QD2CoursPivot, 10, 5, Lx, X1, "#")
fraCours_txtAchatEnCompte = num_Display(recDeviseCours.QD2AchatEnCompte, 10, 5, Lx, X1, "#")
fraCours_txtVenteEnCompte = num_Display(recDeviseCours.QD2VenteEnCompte, 10, 5, Lx, X1, "#")
fraCours_txtAchatNormal = num_Display(recDeviseCours.QD2AchatNormal, 10, 5, Lx, X1, "#")
fraCours_txtVenteNormal = num_Display(recDeviseCours.QD2VenteNormal, 10, 5, Lx, X1, "#")
fraCours_txtAchatPrivilégié = num_Display(recDeviseCours.QD2AchatPrivilégié, 10, 5, Lx, X1, "#")
fraCours_txtVentePrivilégié = num_Display(recDeviseCours.QD2VentePrivilégié, 10, 5, Lx, X1, "#")
libSaisie = "S : " & recDeviseCours.SaisieUsr & " " & dateImp(recDeviseCours.SaisieAMJ) & "   " & timeImp(recDeviseCours.SaisieHMS)
libValidation = "V : " & recDeviseCours.ValidationUsr & " " & dateImp(recDeviseCours.ValidationAMJ) & "   " & timeImp(recDeviseCours.ValidationHMS)
If recDeviseCours.ValidationUsr = "          " Then
    fraCours_cmdOk.Visible = DeviseCoursAut.Saisir
Else
    fraCours_cmdOk.Visible = False
'   Call lstErr_Clear(lstErr, fgCours, "Interdit")
End If

fraCours_txtAchatEnCompte_Control
fraCours_txtAchatNormal_Control
fraCours_txtAchatPrivilégié_Control
fraCours_txtVenteEnCompte_Control
fraCours_txtVenteNormal_Control
fraCours_txtVentePrivilégié_Control
End Sub

Public Sub fraCours_Exit()
lstDevise.Enabled = False
If blnAddNew Then
    mnuCoursAddNew_Click
Else
    fraCours.Visible = False
    fgCours.Enabled = True
    Call lstErr_Clear(lstErr, fgCours, "choisir une Cours 'click'")
    fgCours.SetFocus
End If
End Sub

Public Sub fgCours_Sort()
fgCours.Row = 1
fgCours.RowSel = fgCours.Rows - 1

fgCours.Col = 1
fgCours.ColSel = 2
fgCours.Sort = 1

End Sub

Public Sub fraCours_txtUnité_Control()
X = num_Control(fraCours_txtUnité, valX, 5, 0)
If X <> "" Then
    Call lstErr_AddItem(lstErr, fraCours_txtUnité, X)
Else
    recDeviseCours.QD1 = CLng(valX)
    If recDeviseCours.QD1 = 0 Then Call lstErr_AddItem(lstErr, fraCours_txtUnité, "préciser unité")
End If
End Sub

Public Sub fraCours_txtAchatNormal_Control()
X = num_Control(fraCours_txtAchatNormal, valX, 10, 5)
If X <> "" Then
    Call lstErr_AddItem(lstErr, fraCours_txtAchatNormal, X)
Else
    recDeviseCours.QD2AchatNormal = CDbl(valX)
    If recDeviseCours.QD2AchatNormal = 0 Then Call lstErr_AddItem(lstErr, fraCours_txtAchatNormal, "préciser Achat Normal")
    If recDeviseCours.QD2CoursPivot = 0 Then
        dblX = 0
    Else
        dblX = (recDeviseCours.QD2AchatNormal - recDeviseCours.QD2CoursPivot) / recDeviseCours.QD2CoursPivot
    End If
    dblX = Fix(dblX * 10000 - 0.5) / 100
    libAchatNormal = Format$(dblX, "###0.00")
End If
End Sub

Public Sub fraCours_txtVenteNormal_Control()
X = num_Control(fraCours_txtVenteNormal, valX, 10, 5)
If X <> "" Then
    Call lstErr_AddItem(lstErr, fraCours_txtVenteNormal, X)
Else
    recDeviseCours.QD2VenteNormal = CDbl(valX)
    If recDeviseCours.QD2VenteNormal = 0 Then Call lstErr_AddItem(lstErr, fraCours_txtVenteNormal, "préciser vente Normal")
     If recDeviseCours.QD2CoursPivot = 0 Then
        dblX = 0
    Else
        dblX = (recDeviseCours.QD2VenteNormal - recDeviseCours.QD2CoursPivot) / recDeviseCours.QD2CoursPivot
    End If
    dblX = Fix(dblX * 10000 + 0.5) / 100
    libVenteNormal = Format$(dblX, "###0.00")
End If
End Sub

Public Sub fraCours_txtAchatPrivilégié_Control()
X = num_Control(fraCours_txtAchatPrivilégié, valX, 10, 5)
If X <> "" Then
    Call lstErr_AddItem(lstErr, fraCours_txtAchatPrivilégié, X)
Else
    recDeviseCours.QD2AchatPrivilégié = CDbl(valX)
    If recDeviseCours.QD2AchatPrivilégié = 0 Then Call lstErr_AddItem(lstErr, fraCours_txtAchatPrivilégié, "préciser Achat Privilégié")
    If recDeviseCours.QD2CoursPivot = 0 Then
        dblX = 0
    Else
        dblX = (recDeviseCours.QD2AchatPrivilégié - recDeviseCours.QD2CoursPivot) / recDeviseCours.QD2CoursPivot
    End If
    dblX = Fix(dblX * 10000 - 0.5) / 100
    libAchatPrivilégié = Format$(dblX, "###0.00")
End If
End Sub

Public Sub fraCours_txtVentePrivilégié_Control()
X = num_Control(fraCours_txtVentePrivilégié, valX, 10, 5)
If X <> "" Then
    Call lstErr_AddItem(lstErr, fraCours_txtVentePrivilégié, X)
Else
    recDeviseCours.QD2VentePrivilégié = CDbl(valX)
    If recDeviseCours.QD2VentePrivilégié = 0 Then Call lstErr_AddItem(lstErr, fraCours_txtVentePrivilégié, "préciser vente Privilégié")
    If recDeviseCours.QD2CoursPivot = 0 Then
        dblX = 0
    Else
        dblX = (recDeviseCours.QD2VentePrivilégié - recDeviseCours.QD2CoursPivot) / recDeviseCours.QD2CoursPivot
    End If
    dblX = Fix(dblX * 10000 + 0.5) / 100
    libVentePrivilégié = Format$(dblX, "###0.00")
End If
End Sub

Public Sub fraCours_txtDevise2_Control()
If Trim(fraCours_txtDevise2) = "" Then Call lstErr_AddItem(lstErr, fraCours_txtDevise2, "? Devise2"): Exit Sub
If DevCode(fraCours_txtDevise2) = 0 Then Call lstErr_AddItem(lstErr, fraCours_txtDevise2, "?Devise2 inconnue"): Exit Sub
recDevise2 = XDevise
recDeviseCours.Id2 = recDevise2.DevX

CV_Init CV
CV.DeviseIso = recDevise2.DevX
CV_Attribut CV
If CV.EuroIn Then Call lstErr_AddItem(lstErr, fraCours_txtDevise2, "?Devise2 IN"): Exit Sub

End Sub

Public Sub fraCours_Control()
lstErr.Clear: arrTag_Set False

fraCours_txtDevise1_Control
fraCours_txtUnité_Control
fraCours_txtDevise2_Control
fraCours_txtAchatNormal_Control
fraCours_txtVenteNormal_Control
fraCours_txtAchatPrivilégié_Control
fraCours_txtVentePrivilégié_Control
If recDevise1.DevX = recDevise2.DevX Then Call lstErr_AddItem(lstErr, fraCours_txtDevise1, "Devise1 = Devise2")

End Sub

Public Sub lstDevise_Scan(strdev As String)
Dim K As Integer
For K = 0 To lstDevise.ListCount - 1
    lstDevise.ListIndex = K
    If Mid$(lstDevise.Text, 1, 3) = strdev Then Exit For
Next K
End Sub

Public Sub fraCours_txtCoursPivot_Control()
Dim mCours As Double, dblX As Double
mCours = recDeviseCours.QD2CoursPivot
X = num_Control(fraCours_txtCoursPivot, valX, 10, 5)
If X <> "" Then
    Call lstErr_AddItem(lstErr, fraCours_txtCoursPivot, X)
Else
    recDeviseCours.QD2CoursPivot = CDbl(valX)
    If recDeviseCours.QD2CoursPivot = 0 Then Call lstErr_AddItem(lstErr, fraCours_txtCoursPivot, "préciser cours pivot")
End If
If mCours <> recDeviseCours.QD2CoursPivot Then

    dblX = Fix(recDeviseCours.QD2CoursPivot * constDeviseCours_MargeEnCompte * 10000 + 0.5) / 10000
    fraCours_txtAchatEnCompte = Format$(CDbl(recDeviseCours.QD2CoursPivot - dblX), "####.00000")
    fraCours_txtAchatEnCompte_Control
    fraCours_txtVenteEnCompte = Format$(recDeviseCours.QD2CoursPivot + dblX, "####.00000")
    fraCours_txtVenteEnCompte_Control

    dblX = Fix(recDeviseCours.QD2CoursPivot * constDeviseCours_MargeNormal * 10000 + 0.5) / 10000
    fraCours_txtAchatNormal = Format$(recDeviseCours.QD2CoursPivot - dblX, "####.00000")
    fraCours_txtAchatNormal_Control
    fraCours_txtVenteNormal = Format$(recDeviseCours.QD2CoursPivot + dblX, "####.00000")
    fraCours_txtVenteNormal_Control
    
    dblX = Fix(recDeviseCours.QD2CoursPivot * constDeviseCours_MargePrivilégié * 10000 + 0.5) / 10000
    fraCours_txtAchatPrivilégié = Format$(recDeviseCours.QD2CoursPivot - dblX, "####.00000")
    fraCours_txtAchatPrivilégié_Control
    fraCours_txtVentePrivilégié = Format$(recDeviseCours.QD2CoursPivot + dblX, "####.00000")
    fraCours_txtVentePrivilégié_Control
End If
End Sub

Public Sub fraCours_txtAchatEnCompte_Control()
X = num_Control(fraCours_txtAchatEnCompte, valX, 10, 5)
If X <> "" Then
    Call lstErr_AddItem(lstErr, fraCours_txtAchatEnCompte, X)
Else
    recDeviseCours.QD2AchatEnCompte = CDbl(valX)
    If recDeviseCours.QD2AchatEnCompte = 0 Then Call lstErr_AddItem(lstErr, fraCours_txtAchatEnCompte, "préciser Achat en compte")
    If recDeviseCours.QD2CoursPivot = 0 Then
        dblX = 0
    Else
        dblX = (recDeviseCours.QD2AchatEnCompte - recDeviseCours.QD2CoursPivot) / recDeviseCours.QD2CoursPivot
    End If
    dblX = Fix(dblX * 10000 - 0.5) / 100
    libAchatEnCompte = Format$(dblX, "###0.00")
End If

End Sub

Public Sub fraCours_txtVenteEnCompte_Control()
X = num_Control(fraCours_txtVenteEnCompte, valX, 10, 5)
If X <> "" Then
    Call lstErr_AddItem(lstErr, fraCours_txtVenteEnCompte, X)
Else
    recDeviseCours.QD2VenteEnCompte = CDbl(valX)
    If recDeviseCours.QD2VenteEnCompte = 0 Then Call lstErr_AddItem(lstErr, fraCours_txtVenteEnCompte, "préciser vente encompte")
     If recDeviseCours.QD2CoursPivot = 0 Then
        dblX = 0
    Else
        dblX = (recDeviseCours.QD2VenteEnCompte - recDeviseCours.QD2CoursPivot) / recDeviseCours.QD2CoursPivot
    End If
    dblX = Fix(dblX * 10000 + 0.5) / 100
    libVenteEnCompte = Format$(dblX, "###0.00")
End If

End Sub

Public Sub cmdPrintX(Msg As String)
X = Format$(1, "000000") & Format$(arrDeviseCoursNb, "000000")

prtDeviseCours X, Msg

End Sub
