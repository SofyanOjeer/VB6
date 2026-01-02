VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDeviseCoupures 
   AutoRedraw      =   -1  'True
   Caption         =   "Guichet : coupures "
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
      Height          =   300
      Left            =   9000
      Picture         =   "DeviseCoupures.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   0
      Width           =   400
   End
   Begin MSFlexGridLib.MSFlexGrid fgCoupure 
      Height          =   4350
      Left            =   5160
      TabIndex        =   5
      Top             =   480
      Width           =   3850
      _ExtentX        =   6800
      _ExtentY        =   7673
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      ForeColor       =   12582912
      ForeColorFixed  =   -2147483641
      BackColorBkg    =   14737632
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   1
      FormatString    =   "Nature |>Nominal    |>Accepté |>Séquence|<        Etat"
   End
   Begin VB.Frame fraCoupure 
      Caption         =   "Coupure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   5040
      Width           =   8775
      Begin VB.CommandButton fraCoupure_cmdQuit 
         BackColor       =   &H00C0C0FF&
         Caption         =   "X"
         Height          =   400
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   600
      End
      Begin VB.CommandButton fraCoupure_cmdOk 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ok"
         Height          =   400
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   600
      End
      Begin VB.Frame fraCoupure_fraActif 
         Caption         =   "Acceptée"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5520
         TabIndex        =   11
         Top             =   240
         Width           =   1575
         Begin VB.OptionButton fraCoupure_optDesabled 
            Caption         =   "non"
            Height          =   375
            Left            =   840
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton fraCoupure_optEnabled 
            Caption         =   "oui"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   270
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.TextBox fraCoupure_txtNominal 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.Frame fraCoupure_fraNature 
         Caption         =   "Nature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   1935
         Begin VB.OptionButton fraCoupure_optBillet 
            Caption         =   "Billet"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   280
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton fraCoupure_optPièce 
            Caption         =   "Pièce"
            Height          =   375
            Left            =   960
            TabIndex        =   7
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Label fraCoupure_lblNominal 
         Caption         =   "Nominal"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   520
         Width           =   735
      End
   End
   Begin VB.ListBox lstDevise 
      Height          =   4350
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   3975
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
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFC0&
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
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   2
      Top             =   0
      Width           =   2500
   End
   Begin VB.Menu mnuDevise 
      Caption         =   "Devise"
      Visible         =   0   'False
      Begin VB.Menu mnuCoupureLoad 
         Caption         =   "Rechercher Coupures"
      End
      Begin VB.Menu mnuCoupurePrintGlobal 
         Caption         =   "Impression Globale"
      End
      Begin VB.Menu mnuDevisePrint 
         Caption         =   "Impression Cours"
      End
      Begin VB.Menu mnuDeviseDisplay 
         Caption         =   "Détail Devise"
      End
   End
   Begin VB.Menu mnuCoupure 
      Caption         =   "Coupure"
      Visible         =   0   'False
      Begin VB.Menu mnuCoupureAddNew 
         Caption         =   "Ajouter une Coupure"
      End
      Begin VB.Menu mnuCoupureUpdate 
         Caption         =   "Modifier une Coupure"
      End
      Begin VB.Menu mnuCoupureDelete 
         Caption         =   "Effacer une Coupure"
      End
      Begin VB.Menu mnuCoupurePrint 
         Caption         =   "Impression"
      End
   End
End
Attribute VB_Name = "frmDeviseCoupures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim blnMsgBox_Quit As Boolean
Dim DeviseCoupuresAut As typeAuthorization
Dim X As String, X1 As String, I As Integer
Dim Msg As String, valX As String

Dim recDevise As typeDevise
Dim recDeviseCoupures As typeDeviseCoupures
Dim currentMethod As String, strdev As String
Dim fgCoupure_FormatString As String, fgCoupure_K As Integer
Dim SéquenceMax As Long, fgCoupure_BackColorFixed As Long, fgCoupure_BackColor As Long
Dim blnAddNew As Boolean
'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
cmdControl
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor

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

'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------

'X = Format$(1, "000000") & Format$(arrDeviseCoupuresNb, "000000")

'prtX X

End Sub






Private Sub cmdOk_Click()
For I = 1 To arrDeviseCoupuresNb
    If Trim(arrDeviseCoupures(I).Method) <> "" Then
        If Not IsNull(srvDeviseCoupures_Update(arrDeviseCoupures(I))) Then Exit For
    End If
Next I
cmdClear
End Sub

Private Sub fgCoupure_Click()
lstErr.Clear
'fgCoupure.Col = 1: fgCoupure.CellBackColor = focusUsr.BackColor
If DeviseCoupuresAut.Saisir Then Me.PopupMenu mnuCoupure, vbPopupMenuRightButton
'fgCoupure.Col = 1: fgCoupure.CellBackColor = fgCoupure_BackColor
End Sub

Private Sub fgCoupure_GotFocus()
fgCoupure.BackColorFixed = focusUsr.BackColor
fgCoupure.BackColor = fgCoupure_BackColor
End Sub


Private Sub fgCoupure_LostFocus()
fgCoupure.BackColorFixed = fgCoupure_BackColorFixed
fgCoupure.BackColor = vbWindowBackground
End Sub


'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub


'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
Call BiaPgmAut_Init("G_Coupures", DeviseCoupuresAut)
cmdClear
recDeviseCoupures_Init recDeviseCoupures
lstDevise_Display
fgCoupure_FormatString = fgCoupure.FormatString
fgCoupure_BackColorFixed = fgCoupure.BackColorFixed
fgCoupure_BackColor = fgCoupure.BackColor
End Sub



'---------------------------------------------------------
Public Sub cmdClear()
'---------------------------------------------------------
cmdReset
cmdContext.Enabled = True: cmdContext.BackColor = vbWindowBackground
cmdContext.Caption = constcmdAbandonner: cmdContext.BackColor = errUsr.BackColor
cmdOk.Visible = False
fgCoupure.Enabled = False: fgCoupure.Clear: fgCoupure.Rows = 1
fraCoupure.Visible = False
Call lstErr_Clear(lstErr, cmdContext, "choisir une devise 'click'")
lstDevise.Enabled = True: lstDevise.BackColor = lstUsr.BackColor
fraCoupure_Clear
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
lstErr.Clear
X = num_Control(fraCoupure_txtNominal, valX, 7, recDevise.maxD)
If X <> "" Then
    Call lstErr_AddItem(lstErr, fraCoupure_txtNominal, X)
Else
    recDeviseCoupures.Id = strdev
    recDeviseCoupures.Nominal = CCur(valX)
    recDeviseCoupures.Nature = IIf(fraCoupure_optBillet, "B", "P")
    recDeviseCoupures.Actif = IIf(fraCoupure_optEnabled, " ", "N")
    If recDeviseCoupures.Nominal = 0 Then Call lstErr_AddItem(lstErr, fraCoupure_txtNominal, "préciser montant")
End If
End Sub

Private Sub fraCoupure_cmdOk_Click()
Dim xBillet As String
lstErr.Clear: arrTag_Set False
cmdControl
If lstErr.ListCount > 0 Then Exit Sub

cmdOk.Visible = DeviseCoupuresAut.Valider
blnMsgBox_Quit = True
Select Case currentMethod
    Case constAddNew: fgCoupure_AddNew
    Case constUpdate: fgCoupure_Update
    Case constDelete: fgCoupure_Delete
End Select
If lstErr.ListCount > 0 Then Exit Sub
fgCoupure_Sort
fraCoupure_Exit
End Sub

Private Sub fraCoupure_cmdQuit_Click()
blnAddNew = False
fraCoupure_Exit
End Sub

Private Sub fraCoupure_optBillet_Click()
If fraCoupure_txtNominal.Enabled Then fraCoupure_txtNominal.SetFocus
End Sub

Private Sub fraCoupure_optDesabled_Click()
If fraCoupure_txtNominal.Enabled Then fraCoupure_txtNominal.SetFocus
End Sub

Private Sub fraCoupure_optEnabled_Click()
If fraCoupure_txtNominal.Enabled Then fraCoupure_txtNominal.SetFocus
End Sub

Private Sub fraCoupure_optPièce_Click()
If fraCoupure_txtNominal.Enabled Then fraCoupure_txtNominal.SetFocus
End Sub

Private Sub fraCoupure_txtNominal_GotFocus()
Call txt_GotFocus(fraCoupure_txtNominal)
End Sub

Private Sub fraCoupure_txtNominal_KeyPress(KeyAscii As Integer)
If recDevise.maxD = 0 Then
    Call num_KeyAscii(KeyAscii)
Else
    Call num_KeyAsciiD(KeyAscii, fraCoupure_txtNominal)
End If

End Sub


Private Sub fraCoupure_txtNominal_LostFocus()
Call txt_LostFocus(fraCoupure_txtNominal)
End Sub

Private Sub lstDevise_Click()
strdev = mId$(lstDevise, 1, 3)
End Sub

Private Sub lstDevise_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If DevCode(mId$(lstDevise.Text, 1, 3)) = 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? Devise inconnue"): Exit Sub
recDevise = XDevise
If DeviseCoupuresAut.Consulter Then Me.PopupMenu mnuDevise, vbPopupMenuRightButton
End Sub


Private Sub mnuCoupureAddNew_Click()
If SéquenceMax < 99 Then
    blnAddNew = True
    fgCoupure.Enabled = False
    fraCoupure_Clear
    fraCoupure_txtNominal.Enabled = True
    fraCoupure.Enabled = True: fraCoupure.Visible = True
    currentMethod = constAddNew
    Call lstErr_Clear(lstErr, fraCoupure_txtNominal, "Nouvelle Coupure")
    fraCoupure_txtNominal.SetFocus
Else
    Call lstErr_Clear(lstErr, fgCoupure, "maximum 99 coupures")
End If
End Sub

Private Sub mnuCoupureDelete_Click()
fgCoupure_ScanSéquence
If lstErr.ListCount Then Exit Sub
If recDeviseCoupures.Method = constAddNew Then
    currentMethod = constIgnore
    fgCoupure_Delete
Else
    fgCoupure.Enabled = False
    fraCoupure.Enabled = False: fraCoupure.Visible = True
    currentMethod = constDelete
    Call lstErr_Clear(lstErr, fraCoupure_txtNominal, "Suppression Coupure")
    X = MsgBox("Voulez-vous réellement supprimer cette coupure ?", vbYesNo + vbQuestion + vbDefaultButton2, "Coupure ancienne")
    
    If X = vbYes Then fraCoupure_cmdOk_Click 'fgCoupure_Delete
    fraCoupure.Visible = False
    fgCoupure.Enabled = True
    fgCoupure.SetFocus
End If
End Sub

Private Sub mnuCoupureLoad_Click()
lstDevise.Enabled = False: lstDevise.BackColor = vbWindowBackground

srvDeviseCoupures_Load strdev

For arrDeviseCoupuresIndex = 1 To arrDeviseCoupuresNb
    arrDeviseCoupures(arrDeviseCoupuresIndex).Method = ""
Next arrDeviseCoupuresIndex

fgCoupure_Display
If fgCoupure.Rows > 1 Then
    Call lstErr_Clear(lstErr, fgCoupure, "choisir une coupure 'click'")
Else
    mnuCoupureAddNew_Click
End If

End Sub

Private Sub mnuCoupureUpdate_Click()
fgCoupure_ScanSéquence
If lstErr.ListCount Then Exit Sub
fgCoupure.Enabled = False
fraCoupure_txtNominal.Enabled = False
fraCoupure.Enabled = True: fraCoupure.Visible = True
currentMethod = constUpdate
Call lstErr_Clear(lstErr, fraCoupure_txtNominal, "Modification Coupure")
fraCoupure_optBillet.SetFocus

End Sub

Private Sub mnuDeviseDisplay_Click()
If lstDevise.ListIndex >= 0 Then
    Set XListBox = frmDeviseCoupures.lstDevise
    frmDevise.Show vbModal
End If
End Sub

Private Sub mnuDevisePrint_Click()
Dim Msg As String

Msg = Format$(1, "000000") & Format$(9999, "000000")

prtDeviseX Msg

End Sub



Public Sub fgCoupure_Display()
fgCoupure.Rows = 1
fgCoupure.Clear
fgCoupure.FormatString = fgCoupure_FormatString
fgCoupure.Enabled = True
SéquenceMax = 0
For arrDeviseCoupuresIndex = 1 To arrDeviseCoupuresNb
    If arrDeviseCoupures(arrDeviseCoupuresIndex).Séquence > SéquenceMax Then SéquenceMax = arrDeviseCoupures(arrDeviseCoupuresIndex).Séquence
    If arrDeviseCoupures(arrDeviseCoupuresIndex).Method <> constDelete _
    And arrDeviseCoupures(arrDeviseCoupuresIndex).Method <> constIgnore Then
        fgCoupure.Rows = fgCoupure.Rows + 1
        fgCoupure.Row = fgCoupure.Rows - 1
        fgCoupure_DisplayItem
    End If
Next arrDeviseCoupuresIndex
If arrDeviseCoupuresNb > 0 Then fgCoupure_Sort

End Sub

Public Sub fgCoupure_DisplayItem()
fgCoupure_K = (fgCoupure.Row) * fgCoupure.Cols
If recDevise.maxD = 0 Then
    X = Format$(arrDeviseCoupures(arrDeviseCoupuresIndex).Nominal, "##########")
Else
    X = Format$(arrDeviseCoupures(arrDeviseCoupuresIndex).Nominal, "#######.00")
End If
fgCoupure.TextArray(1 + fgCoupure_K) = X
fgCoupure.TextArray(0 + fgCoupure_K) = IIf(arrDeviseCoupures(arrDeviseCoupuresIndex).Nature = "B", "Billet", "Pièce")
fgCoupure.TextArray(2 + fgCoupure_K) = IIf(arrDeviseCoupures(arrDeviseCoupuresIndex).Actif = " ", "oui", "non")
fgCoupure.TextArray(3 + fgCoupure_K) = Format$(arrDeviseCoupures(arrDeviseCoupuresIndex).Séquence, "##")
Select Case arrDeviseCoupures(arrDeviseCoupuresIndex).Method
    Case constAddNew: fgCoupure.TextArray(4 + fgCoupure_K) = "Créé"
    Case constUpdate: fgCoupure.TextArray(4 + fgCoupure_K) = "Modifié"
End Select
End Sub

Public Sub fgCoupure_AddNew()
X = num_Control(fraCoupure_txtNominal, valX, 5, recDevise.maxD)
If arrDeviseCoupures_ScanNominal(recDeviseCoupures) > 0 Then
    Call lstErr_AddItem(lstErr, fraCoupure_txtNominal, "Existe déjà")
Else
    recDeviseCoupures.Method = constAddNew
    SéquenceMax = SéquenceMax + 1
    recDeviseCoupures.Séquence = SéquenceMax
    Call arrDeviseCoupures_AddItem(recDeviseCoupures)
    arrDeviseCoupuresIndex = arrDeviseCoupuresNb
    fgCoupure.Rows = fgCoupure.Rows + 1
    fgCoupure.Row = fgCoupure.Rows - 1
    fgCoupure_DisplayItem
End If
End Sub

Public Sub fgCoupure_Update()
If recDeviseCoupures.Method <> constAddNew Then recDeviseCoupures.Method = constUpdate
arrDeviseCoupures(arrDeviseCoupuresIndex) = recDeviseCoupures
fgCoupure_DisplayItem
End Sub

Public Sub fgCoupure_Delete()
recDeviseCoupures.Method = currentMethod
arrDeviseCoupures(arrDeviseCoupuresIndex) = recDeviseCoupures
fgCoupure_Display
End Sub

Public Sub fraCoupure_Clear()
fraCoupure_txtNominal = ""
fraCoupure_optBillet = True
fraCoupure.Enabled = True
End Sub

Public Sub fraCoupure_Enabled(ByVal bln As Boolean)
fgCoupure.Enabled = bln
fraCoupure_cmdOk.Visible = bln
'fraCoupure_cmdQuit.Visible = bln
End Sub

Public Sub cmdContext_Quit()
If fraCoupure.Visible Then
    fraCoupure_cmdQuit_Click
Else
    If lstDevise.Enabled Then
        Unload Me
    Else
        If blnMsgBox_Quit Then
            X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
         Else
            X = vbYes
         End If
         If X = vbYes Then cmdClear
    End If
End If

End Sub

Public Sub cmdContext_Return()
If fraCoupure.Enabled Then
    fraCoupure_cmdOk_Click
Else
    SendKeys "{TAB}"
End If

End Sub

Public Sub fgCoupure_ScanSéquence()
fgCoupure_K = fgCoupure.Row * fgCoupure.Cols
recDeviseCoupures.Séquence = Val(fgCoupure.TextArray(3 + fgCoupure_K))
If arrDeviseCoupures_ScanSéquence(recDeviseCoupures) < 0 Then
    Call lstErr_AddItem(lstErr, fgCoupure, "Séquence inconnue")
Else
    recDeviseCoupures = arrDeviseCoupures(arrDeviseCoupuresIndex)
    fraCoupure_DisplayItem
End If
End Sub

Public Sub fraCoupure_DisplayItem()
fraCoupure_txtNominal = num_Display(recDeviseCoupures.Nominal, 5, recDevise.maxD, Lx, X1, "#")
fraCoupure_optBillet = IIf(recDeviseCoupures.Nature = "B", True, False)
fraCoupure_optEnabled = IIf(recDeviseCoupures.Actif = " ", True, False)
End Sub

Public Sub fraCoupure_Exit()
If blnAddNew Then
    mnuCoupureAddNew_Click
Else
    fraCoupure.Visible = False
    fgCoupure.Enabled = True
    Call lstErr_Clear(lstErr, fgCoupure, "choisir une coupure 'click'")
    fgCoupure.SetFocus
End If
End Sub

Public Sub fgCoupure_Sort()
fgCoupure.Row = 1
fgCoupure.RowSel = fgCoupure.Rows - 1

fgCoupure.Col = 0
fgCoupure.ColSel = 1
fgCoupure.Sort = 2

End Sub
