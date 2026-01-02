VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCompte 
   AutoRedraw      =   -1  'True
   Caption         =   "Compte : Interrogation "
   ClientHeight    =   6345
   ClientLeft      =   75
   ClientTop       =   345
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6345
   ScaleWidth      =   9405
   Begin VB.CheckBox chkAmj 
      Caption         =   "Séléction mouvements "
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Frame fraAmj 
      Height          =   400
      Left            =   2040
      TabIndex        =   40
      Top             =   6000
      Visible         =   0   'False
      Width           =   3255
      Begin MSComCtl2.DTPicker txtAmjMax 
         Height          =   300
         Left            =   2040
         TabIndex        =   41
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
         Format          =   60424195
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txtAmjMin 
         Height          =   300
         Left            =   480
         TabIndex        =   42
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
         Format          =   60424195
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   -328351
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
         TabIndex        =   44
         Top             =   100
         Width           =   315
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
         TabIndex        =   43
         Top             =   100
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdoption 
      Caption         =   "&Options"
      Height          =   300
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6070
      Width           =   1000
   End
   Begin VB.Frame fraOption 
      Height          =   4695
      Left            =   6000
      TabIndex        =   28
      Top             =   1200
      Width           =   3255
      Begin VB.CheckBox chkAnnul 
         Caption         =   "inclure les comptes annulés"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   1850
         Width           =   2775
      End
      Begin VB.CheckBox chkAmjSoldé 
         Caption         =   "ignorer les comptes soldés avant le"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Frame fraSolde 
         Caption         =   "Afficher le solde"
         Height          =   2295
         Left            =   120
         TabIndex        =   32
         Top             =   2280
         Width           =   3015
         Begin VB.OptionButton optFinDeMoisOpération 
            Caption         =   "Fin de mois (date d'opération)"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1440
            Width           =   2775
         End
         Begin VB.OptionButton optFinDAnnée 
            Caption         =   "Fin d'année"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1800
            Width           =   2415
         End
         Begin VB.OptionButton optFinDeMois 
            Caption         =   "Fin de mois (date de traitement)"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1080
            Width           =   2775
         End
         Begin VB.OptionButton optVeille 
            Caption         =   "Veille"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   2415
         End
         Begin VB.OptionButton optInstantané 
            Caption         =   "Instantané"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Value           =   -1  'True
            Width           =   2415
         End
      End
      Begin VB.CheckBox chkPageParPage 
         Caption         =   "Page / Page"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   900
         Width           =   1320
      End
      Begin VB.CheckBox chkEffacer 
         Caption         =   "Effacer à chaque recherche"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Value           =   1  'Checked
         Width           =   2500
      End
      Begin VB.CheckBox chkNuméroAncien 
         Caption         =   "Afficher ancien numéro"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   300
         Width           =   2025
      End
      Begin MSComCtl2.DTPicker txtAmjSoldé 
         Height          =   300
         Left            =   1200
         TabIndex        =   46
         Top             =   1500
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
         Format          =   60424195
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   -328351
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   8950
      Picture         =   "COMPTE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   0
      Width           =   400
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ok"
      Height          =   375
      Left            =   8300
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox txtDeviseCv 
      Height          =   300
      Left            =   7560
      TabIndex        =   25
      Top             =   6045
      Width           =   495
   End
   Begin VB.CheckBox chkDeviseCV 
      Caption         =   "Contre-Valeur "
      Height          =   255
      Left            =   6000
      TabIndex        =   24
      Top             =   6045
      Width           =   1455
   End
   Begin VB.ListBox lstDevise 
      Height          =   4545
      Left            =   0
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   3400
   End
   Begin VB.PictureBox picCompte 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   800
      Index           =   5
      Left            =   0
      ScaleHeight     =   760.364
      ScaleMode       =   0  'User
      ScaleWidth      =   9315
      TabIndex        =   22
      Top             =   5250
      Visible         =   0   'False
      Width           =   9350
   End
   Begin VB.PictureBox picCompte 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   800
      Index           =   4
      Left            =   0
      ScaleHeight     =   760.364
      ScaleMode       =   0  'User
      ScaleWidth      =   9315
      TabIndex        =   21
      Top             =   4450
      Visible         =   0   'False
      Width           =   9350
   End
   Begin VB.PictureBox picCompte 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   800
      Index           =   3
      Left            =   0
      ScaleHeight     =   760.364
      ScaleMode       =   0  'User
      ScaleWidth      =   9315
      TabIndex        =   20
      Top             =   3640
      Visible         =   0   'False
      Width           =   9350
   End
   Begin VB.PictureBox picCompte 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   800
      Index           =   2
      Left            =   0
      ScaleHeight     =   760.364
      ScaleMode       =   0  'User
      ScaleWidth      =   9315
      TabIndex        =   19
      Top             =   2830
      Visible         =   0   'False
      Width           =   9350
   End
   Begin VB.PictureBox picCompte 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   800
      Index           =   1
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9285
      TabIndex        =   18
      Top             =   2020
      Visible         =   0   'False
      Width           =   9350
   End
   Begin VB.PictureBox picCompte 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   800
      Index           =   0
      Left            =   0
      ScaleHeight     =   760.364
      ScaleMode       =   0  'User
      ScaleWidth      =   9315
      TabIndex        =   17
      Top             =   1210
      Visible         =   0   'False
      Width           =   9350
   End
   Begin VB.CommandButton cmdPageNext 
      BackColor       =   &H00C0FFC0&
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
      TabIndex        =   8
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdPagePrior 
      BackColor       =   &H00C0FFC0&
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
      Left            =   1000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H8000000F&
      Height          =   900
      Left            =   0
      TabIndex        =   10
      Top             =   280
      Width           =   9375
      Begin VB.TextBox txtNuméroAncien 
         Height          =   300
         Left            =   8400
         MaxLength       =   6
         TabIndex        =   6
         Top             =   480
         Width           =   850
      End
      Begin VB.TextBox txtBiaNum 
         Height          =   300
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   4
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtBiaTyp 
         Height          =   300
         Left            =   3400
         MaxLength       =   3
         TabIndex        =   3
         Top             =   480
         Width           =   400
      End
      Begin VB.TextBox txtNuméro 
         Height          =   300
         Left            =   1440
         MaxLength       =   11
         TabIndex        =   2
         Top             =   480
         Width           =   1140
      End
      Begin VB.TextBox txtDeviseX 
         Height          =   300
         Left            =   360
         MaxLength       =   3
         TabIndex        =   1
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtAlpha 
         Height          =   300
         Left            =   5160
         MaxLength       =   15
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblBiaNum 
         Caption         =   "   N° d'ordre"
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
         Left            =   3840
         TabIndex        =   16
         Top             =   120
         Width           =   1305
      End
      Begin VB.Label lblBiaTyp 
         Caption         =   "       Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2950
         TabIndex        =   15
         Top             =   150
         Width           =   1100
      End
      Begin VB.Label lblIntitulé 
         Caption         =   "                 Alpha"
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
         Left            =   5040
         TabIndex        =   14
         Top             =   150
         Width           =   2055
      End
      Begin VB.Label lblNuméroAncien 
         Caption         =   "     Ancien"
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
         Left            =   8280
         TabIndex        =   13
         Top             =   150
         Width           =   945
      End
      Begin VB.Label lblNuméro 
         Caption         =   " Racine/Compte"
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
         Left            =   1320
         TabIndex        =   12
         Top             =   150
         Width           =   1395
      End
      Begin VB.Label lblDeviseX 
         Caption         =   "  Devise"
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
         TabIndex        =   11
         Top             =   150
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdRechercher 
      BackColor       =   &H8000000A&
      Caption         =   "Rechercher"
      Height          =   300
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1000
   End
   Begin VB.Label lbloptSolde 
      Caption         =   "solde instantané"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3360
      TabIndex        =   38
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ErrMsg"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6240
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Menu txtDeviseXmnu 
      Caption         =   "txtDeviseXmnu"
      Visible         =   0   'False
      Begin VB.Menu txtDeviseXmnuListe 
         Caption         =   "Liste des Devises"
      End
   End
   Begin VB.Menu cmdRecherchermnu 
      Caption         =   "cmdRecherchermnu"
      Visible         =   0   'False
      Begin VB.Menu cmdRecherchermnuNuméro 
         Caption         =   "Racine / Compte"
      End
      Begin VB.Menu cmdRecherchermnuBiaTyp 
         Caption         =   "Type de Compte"
      End
      Begin VB.Menu cmdRecherchermnuBiaNum 
         Caption         =   "N° ordre"
      End
      Begin VB.Menu cmdRecherchermnuAlpha 
         Caption         =   "Alpha"
      End
      Begin VB.Menu cmdRecherchermnuNuméroAncien 
         Caption         =   "Ancien Compte"
      End
   End
   Begin VB.Menu txtNuméromnu 
      Caption         =   "txtNuméromnu"
      Visible         =   0   'False
      Begin VB.Menu txtNuméromnuRechercher 
         Caption         =   "Rechercher"
      End
      Begin VB.Menu txtNuméromnuàPartirDe 
         Caption         =   "à partir de"
      End
   End
   Begin VB.Menu txtAlphamnu 
      Caption         =   "txtAlphamnu"
      Visible         =   0   'False
      Begin VB.Menu txtAlphamnuRechercher 
         Caption         =   "Rechercher"
      End
   End
   Begin VB.Menu txtNuméroAncienmnu 
      Caption         =   "txtNuméroAncienmnu"
      Visible         =   0   'False
      Begin VB.Menu txtNuméroAncienmnuRechercher 
         Caption         =   "Rechercher"
      End
      Begin VB.Menu txtNuméroAncienàPartirDe 
         Caption         =   "à partir de"
      End
   End
   Begin VB.Menu picComptemnu 
      Caption         =   "picComptemnu"
      Visible         =   0   'False
      Begin VB.Menu picComptemnuExtrait 
         Caption         =   "Extrait"
      End
      Begin VB.Menu picComptemnuFiche 
         Caption         =   "Fiche"
      End
      Begin VB.Menu picComptemnuImprimerLigne 
         Caption         =   "Imprimer Fiche Compte"
      End
      Begin VB.Menu mnucmdPrintListe 
         Caption         =   "Imprimer Liste"
      End
      Begin VB.Menu picComptemnuImprimerRib 
         Caption         =   "Imprimer RIB"
      End
      Begin VB.Menu picComptemnuGroupe 
         Caption         =   "Groupe"
      End
   End
   Begin VB.Menu txtBiaTypmnu 
      Caption         =   "txtBiaTypmnu"
      Visible         =   0   'False
      Begin VB.Menu txtBiaTypmnuRechercher 
         Caption         =   "Rechercher"
      End
      Begin VB.Menu txtBiaTypmnuàPartirDE 
         Caption         =   "à partir de"
      End
   End
   Begin VB.Menu txtBiaNummnu 
      Caption         =   "txtBiaNummnu"
      Visible         =   0   'False
      Begin VB.Menu txtBiaNummnuRechercher 
         Caption         =   "Rechercher"
      End
   End
   Begin VB.Menu lstDevisemnu 
      Caption         =   "lstDevisemnu"
      Visible         =   0   'False
      Begin VB.Menu lstDevisemnuOK 
         Caption         =   "OK"
      End
      Begin VB.Menu lstDevisemnuDétail 
         Caption         =   "Détail"
      End
      Begin VB.Menu lstDevisemnuListe 
         Caption         =   "Liste des cours"
      End
   End
   Begin VB.Menu mnucmdPrint 
      Caption         =   "mnucmdPrint"
      Visible         =   0   'False
      Begin VB.Menu picComptemnuImprimerBalance 
         Caption         =   "Imprimer balance des comptes"
      End
      Begin VB.Menu picComptemnuImprimerListe 
         Caption         =   "Imprimer Liste des comptes"
      End
   End
End
Attribute VB_Name = "frmCompte"
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

Private recCompte As typeCompte
Private CptLstMethod As String * 12
Private Const frmCompteNb = 6
Private frmCompteNbMin As Integer
Private frmCompteNbMax As Integer
Private lstDeviseK     As Integer
Private DevCode      As String * 3
Private Xtextbox
Dim mtxtDevName As String
Dim kDeviseVisible As Boolean
Dim valAmjMin As String * 8
Dim valAmjMax As String * 8
Dim valAmjSoldé As String * 8
Dim picCompte_Index As Integer, picCompte_arrCompteIndex As Integer

Dim optSolde As String
Dim varX
Dim blnSetfocus As Boolean
Dim mAMJCours As String * 8
Dim mGroupe As String

'---------------------------------------------------------
Private Sub frmCptInfoShow()
'---------------------------------------------------------
Dim X As String

frmCptInfo.Show vbModeless
frmCptInfo.Visible = True
frmCptInfo.dspCompteInit
X = frmCptInfo.Caption
AppActivate X

End Sub


Private Sub chkAmj_Click()
If chkAmj = "0" Then
    valAmjMin = "00000000": valAmjMax = DSys
    fraAmj.Visible = False
Else
    fraAmj.Visible = True
    txtAmjMin_Control
    txtAmjMax_Control
End If
    

End Sub

Private Sub chkAmj_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkAmj
End Sub


Private Sub chkDeviseCV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkDeviseCV

End Sub

Private Sub chkEffacer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkEffacer

End Sub


Private Sub chkNuméroAncien_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkNuméroAncien

End Sub

Private Sub chkPageParPage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkPageParPage

End Sub


Private Sub cmdOption_Click()
If fraOption.Visible Then
    cmdOption_Ok
Else
    fraOption.Visible = True
    chkNuméroAncien.SetFocus
End If
End Sub


Private Sub cmdoption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdoption

End Sub


Private Sub cmdPageNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPageNext
End Sub

Private Sub cmdPagePrior_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPagePrior

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint
End Sub

Private Sub cmdRechercher_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdRechercher

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

MouseMoveActiveControl_Reset
End Sub

Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraOption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraSolde_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub lblDeviseX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lblDeviseX
End Sub

Private Sub lblNuméro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lblNuméro
End Sub

Private Sub mnucmdPrintListe_Click()
picComptemnuImprimerListe_Click
End Sub

Private Sub optFinDAnnée_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optFinDAnnée

End Sub


Private Sub optFinDeMois_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optFinDeMois

End Sub


Private Sub optFinDeMoisOpération_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optFinDeMoisOpération

End Sub


Private Sub optInstantané_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optInstantané

End Sub


Private Sub optVeille_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optVeille

End Sub


Private Sub picCompte_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Private Sub picComptemnuGroupe_Click()
CptLstMethod = "SnapK5"
frmCompteRechercher

End Sub

Private Sub txtAmjMax_Change()
txtAmjMax_Control

End Sub

'-------------------------------------------------'
Private Sub txtAmjMax_GotFocus()
'-------------------------------------------------'
DTPicker_GotFocus txtAmjMax
End Sub

'-------------------------------------------------'
Private Sub txtAmjMax_Control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAmjMax.Year, "0000") & Format$(txtAmjMax.Month, "00") & Format$(txtAmjMax.Day, "00")
lblAmjMax.ForeColor = lblUsr.ForeColor
lblAmjMax = "au "
If Not IsNumeric(X) Then
    lblAmjMax = lblAmjMax & X
    lblAmjMax.ForeColor = errUsr.ForeColor
    DTPicker_Now txtAmjMax
Else
    If mId$(X, 1, 8) > DSys Then
        lblAmjMax = lblAmjMax & " > aujourd'hui"
        lblAmjMax.ForeColor = errUsr.ForeColor
        DTPicker_Now txtAmjMax
    Else
        valAmjMax = mId$(X, 1, 8)
    End If
End If
End Sub

'-------------------------------------------------'
Private Sub txtAmjMin_Control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAmjMin.Year, "0000") & Format$(txtAmjMin.Month, "00") & Format$(txtAmjMin.Day, "00")
lblAmjMin.ForeColor = lblUsr.ForeColor
lblAmjMin = "du "
If Not IsNumeric(X) Then
    lblAmjMin = lblAmjMin & X
    lblAmjMin.ForeColor = errUsr.ForeColor
    DTPicker_Now txtAmjMin
Else
    If mId$(X, 1, 8) > DSys Then
        lblAmjMin = lblAmjMin & " > aujourd'hui"
        lblAmjMin.ForeColor = errUsr.ForeColor
        DTPicker_Now txtAmjMin
    Else
        valAmjMin = mId$(X, 1, 8)
    End If
End If

End Sub

'-------------------------------------------------'
Private Sub txtAmjSoldé_Control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAmjSoldé.Year, "0000") & Format$(txtAmjSoldé.Month, "00") & Format$(txtAmjSoldé.Day, "00")
chkAmjSoldé.ForeColor = lblUsr.ForeColor
If Not IsNumeric(X) Then
    chkAmjSoldé.ForeColor = errUsr.ForeColor
    DTPicker_Now txtAmjSoldé
Else
    If mId$(X, 1, 8) > DSys Then
        chkAmjSoldé.ForeColor = errUsr.ForeColor
        DTPicker_Now txtAmjSoldé
    Else
        valAmjSoldé = mId$(X, 1, 8)
    End If
End If

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

'---------------------------------------------------------
Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

cmdOk.Visible = False
MousePointer = 0
If Trim(mId$(X, 13, 12)) <> "" Then
    txtDeviseX.Text = Format$(mId$(X, 35, 3), "000")
    If IsNumeric(mId$(X, 38, 11)) Then
        txtNuméro.Text = Trim(Format$(Val(mId$(X, 38, 11)), "###########"))
    Else
        txtAlpha.Text = mId$(X, 38, 11)
    End If
    If Trim(mId$(X, 38, 11)) <> "" Then
        chkPageParPage = "1"
        cmdRechercher_Click
    End If
    cmdOk.Tag = mId$(X, 13, 12)
    If mId$(X, 49, 1) <> "V" Then
        cmdOk.Visible = True
        MousePointer = 10
    '    MouseIcon = frmElp.imgBiaMouse
    '    cmdok.Visible = IIf(Mid$(X, 13, 12) = "", False, True)
        If cmdOk.Visible Then
            txtBiaTyp = "001"
        End If
    End If
End If
End Sub



'---------------------------------------------------------
Private Sub chkNuméroAncien_Click()
'---------------------------------------------------------

arrCompteIndex = frmCompteNbMin - 1
frmCompteItem

End Sub

'---------------------------------------------------------
Private Sub chkDeviseCV_Click()
'---------------------------------------------------------

If chkDeviseCV Then
    txtDeviseCv.Visible = True
    If blnSetfocus Then txtDeviseCv.SetFocus
Else
    txtDeviseCv.Visible = False
    If arrCompteNb > 0 Then
        arrCompteIndex = frmCompteNbMin - 1
        frmCompteItem
    End If
End If
End Sub

'---------------------------------------------------------
Private Sub cmdRechercher_Click()
'---------------------------------------------------------
Dim X As String
cmdOption_Ok
X = ""
CptLstMethod = ""

If Val(txtNuméro) = 0 Then
    cmdRecherchermnuNuméro.Enabled = False
Else
    cmdRecherchermnuNuméro.Enabled = True
    X = X & "Numéro"
End If

If Val(txtBiaTyp) = 0 Then
    cmdRecherchermnuBiaTyp.Enabled = False
Else
    cmdRecherchermnuBiaTyp.Enabled = True
    X = X & "BiaTyp"
End If
If Val(txtBiaNum) = 0 Then
    cmdRecherchermnuBiaNum.Enabled = False
Else
    cmdRecherchermnuBiaNum.Enabled = True
    X = X & "BiaNum"
End If
If Trim(txtAlpha) = "" Then
    cmdRecherchermnuAlpha.Enabled = False
Else
    cmdRecherchermnuAlpha.Enabled = True
    X = X & "Alpha"
End If
If Val(txtNuméroAncien) = 0 Then
    cmdRecherchermnuNuméroAncien.Enabled = False
Else
    cmdRecherchermnuNuméroAncien.Enabled = True
    X = X & "NuméroAncien"
End If


Select Case X
    Case Is = "Numéro"
        cmdRecherchermnuNuméro_Click
    Case Is = "BiaTyp", "NuméroBiaTyp"
        cmdRecherchermnuBiaTyp_Click
    Case Is = "BiaNum"
        cmdRecherchermnuBiaNum_Click
    Case Is = "Alpha"
        cmdRecherchermnuAlpha_Click
    Case Is = "NuméroAncien"
        cmdRecherchermnuNuméroAncien_Click
    Case Else
        If X <> "" Then Me.PopupMenu cmdRecherchermnu, vbPopupMenuRightButton
End Select

'.....frmCompteRechercher

End Sub



'---------------------------------------------------------
Private Sub frmCompteSuite()
'---------------------------------------------------------
Dim X As String

If arrCompteNb > 0 Then
    recCompte = arrCompte(arrCompteNb)
    X = "+"
Else
    If chkEffacer Then frmCompteEffacer
    frmCompteSql
    X = " "
End If
    
If Trim(mId$(CptLstMethod, 5, 2)) = "" Then Mid$(CptLstMethod, 5, 2) = "L1"

recCompte.Method = "Snap" & mId$(CptLstMethod, 5, 2) & X
frmCompteRechercher_Option
 
If IsNull(srvCompteMon(recCompte)) Then
    Call frmCompteItem
End If

End Sub


'---------------------------------------------------------
Private Sub frmCompteItem()
'---------------------------------------------------------
Dim K As Integer
picCompte_Color
If arrCompteIndex < 0 Then
    arrCompteIndex = 0
End If
If arrCompteNb <= 0 Then
    lblErrMsg.Caption = "Néant    ( " & Now & " )"
    lblErrMsg.Visible = True
    lblErrMsg.Top = 2500
    lblErrMsg.Left = 2500
    lblErrMsg.FontBold = True
    lblErrMsg.FontSize = 12
    frmCompteNbMin = 0
Else
    frmCompteNbMin = arrCompteIndex + 1
End If

CV_X1.CoursCompta = "C"
CV_X2.CoursCompta = "C"

For K = 0 To frmCompteNb - 1
    
    If arrCompteIndex >= arrCompteNb Then
        picCompte(K).Visible = False
    Else
        picCompte(K).Visible = True
        arrCompteIndex = arrCompteIndex + 1
        Call recCompte_Display(arrCompte(arrCompteIndex), picCompte(K))

    End If
Next K

frmCompteNbMax = arrCompteIndex

Call frmCompteEnabled

End Sub

'---------------------------------------------------------
Private Sub frmCompteSql()
'---------------------------------------------------------
Dim Numéro As Double, Devise As Double, Alpha As String, NuméroAncien As Double
Dim NuméroMax As Double, Devisemax As Double, AlphaMax As String, NuméroAncienMax As Double
Dim BiaNum As Integer, BiaNumMax As Integer
Dim BiaTyp As Integer, BiaTypMax As Integer
    
NuméroAncien = Val(txtNuméroAncien) * 100000
NuméroAncienMax = NuméroAncien + 99999

Devise = Val(txtDeviseX)
If Devise = 0 Then
    Devisemax = 999#
Else
    Devisemax = Devise
End If

Numéro = Val(txtNuméro)
If Numéro = 0 Then
    NuméroMax = 99999999999#
Else
    If Numéro <= 99999 Then
        Numéro = Numéro * 1000000
        NuméroMax = Numéro + 999999
    Else
        Numéro = mId$(Format$(Numéro, "00000000000"), 1, 10) & "0"
        NuméroMax = Numéro + 9
    End If
End If

BiaTyp = Val(txtBiaTyp)
If BiaTyp = 0 Then
    BiaTypMax = 999#
Else
    BiaTypMax = BiaTyp
End If

BiaNum = Val(txtBiaNum)
If BiaNum = 0 Then
    BiaNumMax = 999#
Else
    BiaNumMax = BiaNum
End If

Alpha = UCase$(Trim(txtAlpha))
AlphaMax = Alpha & "99999999999"

If Trim(CptLstMethod) <> "" Then
    If mId$(CptLstMethod, 1, 4) = "XXXX" Then
        Mid$(CptLstMethod, 1, 4) = "Snap"
        NuméroMax = 99999999999#
        Devisemax = 999#
        NuméroAncienMax = 99999999999#
    End If
Else
    If Devise <> 0 Then
        CptLstMethod = "SnapL1"
        txtDeviseX.SetFocus
    Else
        If Numéro <> 0 Then
            CptLstMethod = "SnapL5"
            txtNuméro.SetFocus
        Else
            If NuméroAncien <> 0 Then
                CptLstMethod = "SnapLX"
                txtNuméroAncien.SetFocus
         Else
                CptLstMethod = "SnapLZ"
                txtAlpha.SetFocus
            End If
        End If
    End If
End If
recCompte.Method = CptLstMethod
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = Format$(Val(Devise), "000")
recCompte.Numéro = Format$(Val(Numéro), "00000000000")
recCompte.Alpha = Alpha
recCompte.NuméroAncien = Format$(Val(NuméroAncien), "00000000000")
recCompte.BiaTyp = Format$(Val(BiaTyp), "000")
recCompte.BiaNum = Format$(Val(BiaNum), "00")

frmCompteRechercher_Option

arrCompte(0) = recCompte
arrCompte(0).Devise = Format$(Val(Devisemax), "000")
arrCompte(0).Numéro = Format$(Val(NuméroMax), "00000000000")
arrCompte(0).Alpha = AlphaMax
arrCompte(0).NuméroAncien = Format$(Val(NuméroAncienMax), "00000000000")
arrCompte(0).BiaTyp = Format$(Val(BiaTypMax), "000")
arrCompte(0).BiaNum = Format$(Val(BiaNumMax), "00")

If Trim(CptLstMethod) = "SnapK5" Then
    recCompte.Groupe = mGroupe: recCompte.Numéro = "00000000000"
    arrCompte(0).Groupe = mGroupe: arrCompte(0).Numéro = "99999999999"
End If

End Sub




'---------------------------------------------------------
Private Sub frmCompteEnabled()
'---------------------------------------------------------
Dim I As Integer

'If arrCompteNb > 0 Then
 '   cmdCptLstDétail.Enabled = True
  '  cmdCptLstOk.Enabled = True
'End If

cmdPageNext.Caption = Format$(frmCompteNbMax, "####") & "..>>"
If frmCompteNbMax < arrCompteNb Or arrComptesuite Then
    cmdPageNext.Enabled = True
    cmdPageNext.BackColor = &HC0FFC0
Else
    cmdPageNext.Enabled = False
    cmdPageNext.BackColor = &H8000000F
End If

cmdPagePrior.Caption = "<<..." & Format$(frmCompteNbMin, "####")
If frmCompteNbMin > 1 Then
    cmdPagePrior.Enabled = True
    cmdPagePrior.BackColor = &HC0FFC0
Else
    cmdPagePrior.Enabled = False
    cmdPagePrior.BackColor = &H8000000F
End If

End Sub

'---------------------------------------------------------
Public Sub frmCompteEffacer()
'---------------------------------------------------------
Dim K As Integer

arrCompteNb = 0: arrCompteIndex = 0
recCompteInit recCompte
For K = 0 To frmCompteNb - 1
    picCompte(K).Visible = False
Next K
frmCompteNbMin = 0: frmCompteNbMax = 0
frmCompteEnabled
picCompte_Color

End Sub

'---------------------------------------------------------
Private Sub cmdPageNext_Click()
'---------------------------------------------------------

If arrCompteIndex < arrCompteNb Then
    frmCompteItem
Else
    arrCompteIndex = arrCompteNb
    If chkPageParPage Then
        If arrComptesuite Then frmCompteSuite
    Else
            frmCompteSuiteEtFin
    End If
End If
End Sub

'---------------------------------------------------------
Private Sub cmdPagePrior_Click()
'---------------------------------------------------------

arrCompteIndex = frmCompteNbMin - frmCompteNb - 1

frmCompteItem

End Sub

'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
Me.PopupMenu mnucmdPrint
End Sub

'---------------------------------------------------------
Private Sub cmdRecherchermnuBiaNum_Click()
'---------------------------------------------------------
txtBiaNummnuRechercher_Click

End Sub

'---------------------------------------------------------
Private Sub cmdRecherchermnuBiaTyp_Click()
'---------------------------------------------------------
txtBiaTypmnuRechercher_Click
End Sub

'---------------------------------------------------------
Private Sub cmdRecherchermnuAlpha_Click()
'---------------------------------------------------------
Call txtAlphamnuRechercher_Click
End Sub


'---------------------------------------------------------
Private Sub cmdRecherchermnuNuméroAncien_Click()
'---------------------------------------------------------
txtNuméroAncienmnuRechercher_Click

End Sub


'---------------------------------------------------------
Private Sub cmdRecherchermnuNuméro_Click()
'---------------------------------------------------------
txtNuméromnuRechercher_Click

End Sub



'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
If txtNuméro.Enabled Then txtNuméro.SetFocus

End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 27: cmdQuit_Click
    Case Is = 34: cmdPageNext_Click
    Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
    Case Is = 13: cmdReturn_Click
End Select
lblErrMsg.Visible = False
End Sub

'---------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------
'On Error Resume Next
'frmElp.SetFocus
End Sub





'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
If fraOption.Visible Then
    fraOption.Visible = False
    cmdRechercher.SetFocus
Else
    Unload Me
End If
End Sub

'---------------------------------------------------------
Private Sub cmdReturn_Click()
'---------------------------------------------------------
If fraOption.Visible Then cmdOption_Ok
    
cmdRechercher_Click
End Sub



'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
Call BiaPgmAut_Init("Compte", CompteAut)

MouseMoveActiveControl_Name = ""
lstDeviseK = 0
mtxtDevName = ""
valAmjMin = "00000000"
valAmjMax = DSys
mAMJCours = DSys
Call DTPicker_Set(txtAmjMax, dateElp("Ouvré", -1, DSys))
txtAmjMax_Control
Call DTPicker_Set(txtAmjMin, mId$(valAmjMax, 1, 6) & "01")
valAmjSoldé = dateElp("Jour", -93, DSys)
Call DTPicker_Set(txtAmjSoldé, valAmjSoldé)
blnSetfocus = False
chkDeviseCV = "1": txtDeviseCv = "EUR": valDeviseCV
Call DevX(1)
'99-01-20 CV = XDevise
blnSetfocus = True

CptLstMethod = ""
Call frmCompteEffacer
optSolde = "I": cmdOption_Display
End Sub


'---------------------------------------------------------
Public Sub frmCompteRechercher()
'---------------------------------------------------------
fraOption.Visible = False
If chkEffacer Then frmCompteEffacer

frmCompteSql

If IsNull(srvCompteMon(recCompte)) Then
    If mId$(CptLstMethod, 1, 4) = "Seek" Then
        Call arrCompteAddItem(recCompte)
        Call frmCompteItem
    Else
        Call frmCompteItem
    End If
End If

If chkPageParPage Then
Else
    frmCompteSuiteEtFin
    If arrCompteNb > frmCompteNb Then arrCompteIndex = 0: frmCompteItem
        
    
End If
End Sub





Private Sub lblIntitulé_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    Call txtAlpha_MouseDown(Button, Shift, X, Y)
End If


End Sub


'---------------------------------------------------------
Private Sub lblNuméro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
'If Button = vbRightButton Then
   If Val(txtNuméro) = 0 Then
        MsgBox "Préciser le compte ou la racine", , "interrogation de comptes"
    Else
        Me.PopupMenu txtNuméromnu, vbPopupMenuRightButton
    End If
'End If

End Sub


'---------------------------------------------------------
Private Sub lstDevise_Click()
'---------------------------------------------------------

Xtextbox.Text = mId$(lstDevise.Text, 1, 3)
If Xtextbox.Name = "txtDeviseCv" Then valDeviseCV
End Sub

'---------------------------------------------------------
Private Sub lstDevise_GotFocus()
'---------------------------------------------------------

kDeviseVisible = False
Xtextbox.BackColor = vbCyan

End Sub

'---------------------------------------------------------
Private Sub lstDevise_LostFocus()
'---------------------------------------------------------
lstDevise.Visible = kDeviseVisible
Xtextbox.BackColor = txtUsr.BackColor

End Sub

'---------------------------------------------------------
Private Sub lstDevise_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
'If Button = vbRightButton Then
    Me.PopupMenu lstDevisemnu, vbPopupMenuRightButton
'End If

End Sub

'---------------------------------------------------------
Private Sub lstDevisemnuDétail_Click()
'---------------------------------------------------------

If lstDevise.ListIndex < 0 Then lstDevise.ListIndex = 0
    Set XListBox = frmCompte.lstDevise
    kDeviseVisible = True
    frmDevise.Show vbModal
    kDeviseVisible = False
'End If
End Sub

'---------------------------------------------------------
Private Sub lstDevisemnuListe_Click()
'---------------------------------------------------------
Dim Msg As String

Msg = Format$(1, "000000") & Format$(999, "000000")

prtDeviseX Msg


End Sub

'---------------------------------------------------------
Private Sub lstDevisemnuOK_Click()
'---------------------------------------------------------
lstDevise_Click
'Xtextbox.Text = Mid$(lstDevise.Text, 1, 3)
lstDevise.Visible = False

End Sub

Private Sub picCompte_Click(Index As Integer)
If cmdOk.Visible Then
    Msg_Send
Else
    picComptemnuExtrait_Click
End If
End Sub

'---------------------------------------------------------
Private Sub picCompte_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
If picCompte_Index > -1 Then
    picCompte(picCompte_Index).BackColor = picUsr.BackColor
    Call recCompte_Display(arrCompte(picCompte_arrCompteIndex), picCompte(picCompte_Index))
End If

arrCompteIndex = frmCompteNbMin + Index
picCompte(Index).BackColor = focusUsr.BackColor
Call recCompte_Display(arrCompte(arrCompteIndex), picCompte(Index))
picCompte_Index = Index
picCompte_arrCompteIndex = arrCompteIndex

If Button = vbRightButton Then
    picComptemnuGroupe.Enabled = False
    If arrCompte(arrCompteIndex).Groupe <> 0 Then mGroupe = arrCompte(arrCompteIndex).Groupe: picComptemnuGroupe.Enabled = True
    If ctlGestionnaire_New(arrCompte(arrCompteIndex).Numéro, arrCompte(arrCompteIndex).Gestionnaire, arrCompte(arrCompteIndex).BiaTyp) Then
        Me.PopupMenu picComptemnu, vbPopupMenuRightButton + vbPopupMenuCenterAlign
    Else
        If UCase(Trim(usrId)) = "TOUNSI" Or UCase(Trim(usrId)) = "TAN" Then
            picComptemnuExtrait.Enabled = False
            Me.PopupMenu picComptemnu, vbPopupMenuRightButton + vbPopupMenuCenterAlign
        End If
    End If
End If

End Sub



'---------------------------------------------------------
Private Sub picComptemnuExtrait_Click()
'---------------------------------------------------------
Dim X As String
If ctlGestionnaire_New(arrCompte(arrCompteIndex).Numéro, arrCompte(arrCompteIndex).Gestionnaire, arrCompte(arrCompteIndex).BiaTyp) Then
    chkAmj_Click
    frmCptMvtShow
End If

End Sub

'---------------------------------------------------------
Private Sub picComptemnuFiche_Click()
'---------------------------------------------------------

If IsNull(srvCompteCptInfo) Then frmCptInfoShow

End Sub


'---------------------------------------------------------
Private Sub picComptemnuImprimerLigne_Click()
'---------------------------------------------------------
Dim Msg As String

If IsNull(srvCompteCptInfo) Then

    Msg = Format$(0, "000000") & Format$(0, "000000")
    prtCptInfoX Msg

End If

End Sub


'---------------------------------------------------------
Private Sub picComptemnuImprimerListe_Click()
'---------------------------------------------------------
Dim Msg As String


Msg = Format$(1, "000000") & Format$(arrCompteNb, "000000") & "         "
If chkNuméroAncien Then Mid$(Msg, 13, 1) = "X"
If chkDeviseCV Then Mid$(Msg, 14, 3) = Trim(txtDeviseCv)
Mid$(Msg, 18, 1) = optSolde

prtCompteX Msg
End Sub

'---------------------------------------------------------
Private Sub picComptemnuImprimerBalance_Click()
'---------------------------------------------------------
Dim Msg As String, xPrint As String

If Trim(CptLstMethod) = "SnapL1" Or Trim(CptLstMethod) = "SnapL5" Then
    xPrint = "S=LR*"
Else
    xPrint = "S=L**"
End If

Msg = Format$(1, "000000") & Format$(arrCompteNb, "000000") & "    B " & DSys & xPrint
If chkNuméroAncien Then Mid$(Msg, 13, 1) = "X"
If chkDeviseCV Then Mid$(Msg, 14, 3) = Trim(txtDeviseCv)
Mid$(Msg, 18, 1) = optSolde
Mid$(Msg, 19, 8) = mAMJCours
prtCompteSolde_Monitor Msg 'prtCompteX Msg
End Sub


'---------------------------------------------------------
Private Sub picComptemnuImprimerRib_Click()
'---------------------------------------------------------
Dim Msg As String
If arrCompte(arrCompteIndex).BiaTyp <> "001" Then
        Msg = MsgBox("Voulez-vous imprimer un RIB pour ce compte", vbInformation + vbYesNo, "Le type du compte  n'est pas '001' (ordinaire)")
        If Msg <> vbYes Then Exit Sub
End If
If IsNull(srvCompteCptInfo) Then

    Msg = Format$(0, "000000") & Format$(0, "000000")
    prtRib Msg

End If

End Sub

Private Sub txtAmjMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DTPicker_GotFocus txtAmjMin

End Sub

Private Sub txtAmjSoldé_Change()
txtAmjSoldé_Control

End Sub


Private Sub txtAmjSoldé_GotFocus()
DTPicker_GotFocus txtAmjSoldé
End Sub


Private Sub txtAmjSoldé_LostFocus()
DTPicker_LostFocus txtAmjSoldé
txtAmjSoldé_Control
End Sub


Private Sub txtAmjSoldé_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DTPicker_GotFocus txtAmjSoldé

End Sub

'-------------------------------------------------'
Private Sub txtBiaNum_GotFocus()
'-------------------------------------------------'
Call txtGetFocus(txtBiaNum)

End Sub


'-------------------------------------------------'
Private Sub txtBiaNum_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
Call num_KeyAscii(KeyAscii)
End Sub


'---------------------------------------------------------
Private Sub txtBiaNum_LostFocus()
'-------------------------------------------------'
txtBiaNum.BackColor = txtUsr.BackColor

End Sub

'---------------------------------------------------------
Private Sub lblBiaNum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
If Button = vbRightButton Then
    If Val(txtBiaNum) = 0 Then
        MsgBox "Préciser le N° d'ordre", , "interrogation de comptes"
    Else
        Me.PopupMenu txtBiaNummnu, vbPopupMenuRightButton
    End If
End If
End Sub

'---------------------------------------------------------
Private Sub txtBiaNummnuRechercher_Click()
'---------------------------------------------------------

If Val(txtNuméro) = 0 Then
    MsgBox "Préciser la racine", , "interrogation de comptes"
Else
    txtBiaTyp = "___" 'Format$(0, txtBiaTyp.Format)
    CptLstMethod = "SnapLN"
    frmCompteRechercher
End If
End Sub

'-------------------------------------------------'
Private Sub txtBiaTyp_GotFocus()
'-------------------------------------------------'
Call txtGetFocus(txtBiaTyp)

End Sub


'-------------------------------------------------'
Private Sub txtBiaTyp_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
Call num_KeyAscii(KeyAscii)
End Sub


'---------------------------------------------------------
Private Sub txtBiaTyp_LostFocus()
'-------------------------------------------------'
txtBiaTyp.BackColor = txtUsr.BackColor

End Sub

'---------------------------------------------------------
Private Sub lblBiaTyp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
If Button = vbRightButton Then
    If Val(txtBiaTyp) = 0 Then
        MsgBox "Préciser le type de compte", , "interrogation de comptes"
    Else
        Me.PopupMenu txtBiaTypmnu, vbPopupMenuRightButton
    End If
End If
End Sub

'---------------------------------------------------------
Private Sub txtBiaTypmnuàPartirDE_Click()
'---------------------------------------------------------

CptLstMethod = "XXXXLA"
frmCompteRechercher


End Sub

'---------------------------------------------------------
Private Sub txtBiaTypmnuRechercher_Click()
'---------------------------------------------------------

txtBiaNum = "__" 'Format$(0, txtBiaNum.Format)
CptLstMethod = "SnapLA"
frmCompteRechercher


End Sub

'-------------------------------------------------'
Private Sub txtAlpha_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'

KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

'---------------------------------------------------------
Private Sub txtAlphamnuRechercher_Click()
'---------------------------------------------------------
chkPageParPage = "1"
CptLstMethod = "SnapLZ"
frmCompteRechercher

End Sub

'---------------------------------------------------------
Private Sub txtNuméroAncien_GotFocus()
'-------------------------------------------------'
Call txtGetFocus(txtNuméroAncien)

End Sub

'---------------------------------------------------------
Private Sub txtNuméroAncien_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
Call num_KeyAscii(KeyAscii)
End Sub


'---------------------------------------------------------
Private Sub txtNuméroAncien_LostFocus()
'-------------------------------------------------'
txtNuméroAncien.BackColor = txtUsr.BackColor

End Sub


'---------------------------------------------------------
Private Sub lblNuméroAncien_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
If Button = vbRightButton Then
    If Val(txtNuméroAncien) = 0 Then
        MsgBox "Préciser l'ancien compte", , "interrogation de comptes"
    Else
        Me.PopupMenu txtNuméroAncienmnu, vbPopupMenuRightButton
    End If
End If
End Sub


'---------------------------------------------------------
Private Sub txtNuméroAncienàPartirDe_Click()
'---------------------------------------------------------
CptLstMethod = "XXXXLX"
frmCompteRechercher
End Sub

'---------------------------------------------------------
Private Sub txtNuméroAncienmnuRechercher_Click()
'---------------------------------------------------------

CptLstMethod = "SnapLX"
frmCompteRechercher

End Sub

'---------------------------------------------------------
Private Sub txtDeviseX_GotFocus()
'-------------------------------------------------'
Call txtGetFocus(txtDeviseX)
Set Xtextbox = txtDeviseX
End Sub


'-------------------------------------------------'
Private Sub txtDeviseX_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
Call num_KeyAscii(KeyAscii)
End Sub


'---------------------------------------------------------
Private Sub txtDeviseX_LostFocus()
'-------------------------------------------------'
txtDeviseX.BackColor = txtUsr.BackColor
If Trim(txtDeviseX) <> "___" Then
    If Trim(DevX(Val(txtDeviseX))) = "" Then
        lblErrMsg = " ? devise"
        lblErrMsgStd
'       txtDeviseX.SetFocus
    End If
End If

End Sub


'---------------------------------------------------------
Private Sub txtAlpha_GotFocus()
'-------------------------------------------------'
Call txtGetFocus(txtAlpha)

End Sub

'---------------------------------------------------------
Private Sub txtAlpha_LostFocus()
'-------------------------------------------------'
txtAlpha.BackColor = txtUsr.BackColor

End Sub


'-------------------------------------------------'
Private Sub txtAlpha_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
If Button = vbRightButton Then
    If Trim(txtAlpha) = "" Then
        MsgBox "Préciser l'intitulé", , "interrogation de comptes"
    Else
    Me.PopupMenu txtAlphamnu, vbPopupMenuRightButton
    End If
End If
End Sub


'---------------------------------------------------------
Private Sub lblDeviseX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------

'If Button = vbRightButton Then
    Set Xtextbox = txtDeviseX
'    Me.PopupMenu txtDeviseXmnu, vbPopupMenuRightButton
'End If
mtxtDevName = "txtDeviseX"
frmComptelstDevise


End Sub

'---------------------------------------------------------
Private Sub txtDeviseXmnuListe_Click()
'---------------------------------------------------------

frmComptelstDevise

End Sub


'-------------------------------------------------'
Private Sub txtNuméro_GotFocus()
'-------------------------------------------------'
Call txtGetFocus(txtNuméro)
txtNuméro.SelStart = 0
End Sub


'-------------------------------------------------'
Private Sub txtNuméro_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
Call num_KeyAscii(KeyAscii)
End Sub


'---------------------------------------------------------
Private Sub txtNuméro_LostFocus()
'-------------------------------------------------'
txtNuméro.BackColor = txtUsr.BackColor

End Sub







'---------------------------------------------------------
Private Sub txtNuméromnuàPartirDe_Click()
'---------------------------------------------------------
If Val(txtDeviseX) <> 0 Then
    CptLstMethod = "XXXXL1"
Else
    CptLstMethod = "XXXXL5"
End If
chkPageParPage = "1"
frmCompteRechercher

End Sub

'---------------------------------------------------------
Private Sub txtNuméromnuRechercher_Click()
'---------------------------------------------------------

If Val(txtDeviseX) <> 0 Then
    CptLstMethod = "SnapL1"
Else
    CptLstMethod = "SnapL5"
End If
frmCompteRechercher

End Sub



'---------------------------------------------------------
Public Sub lblErrMsgStd()
'---------------------------------------------------------
Beep
lblErrMsg.Visible = True
lblErrMsg.FontBold = True
lblErrMsg.FontSize = 8
lblErrMsg.Top = 0
lblErrMsg.Left = 6000

End Sub



'---------------------------------------------------------
Private Sub txtDeviseCV_GotFocus()
'-------------------------------------------------'
Call txtGetFocus(txtDeviseCv)
Set Xtextbox = txtDeviseCv
If mtxtDevName = "txtDeviseCV" Then
    mtxtDevName = ""
Else
    mtxtDevName = "txtDeviseCV"
    frmComptelstDevise
End If

End Sub


'-------------------------------------------------'
Private Sub txtDeviseCV_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
KeyAscii = convUCase(KeyAscii)
End Sub


'-------------------------------------------------'
Private Sub txtDeviseCV_LostFocus()
'-------------------------------------------------'

txtDeviseCv.BackColor = txtUsr.BackColor
'lstDevise.Visible = False
valDeviseCV

End Sub



'---------------------------------------------------------
Public Sub frmCptMvtShow()
'---------------------------------------------------------
Dim X As String

frmCptMvt.Show vbModeless
frmCptMvt.WindowState = vbNormal
frmCptMvt.Visible = True
frmCptMvt.frmCptMvtInit valAmjMin, valAmjMax
X = frmCptMvt.Caption
AppActivate X

End Sub
Public Sub frmDeviseChange_show()
Dim X As String
frmDeviseChange.Show vbModeless
frmDeviseChange.WindowState = vbNormal
frmDeviseChange.Visible = True
X = frmDeviseChange.Caption
AppActivate X
End Sub

'---------------------------------------------------------
Public Sub frmCompteSuiteEtFin()
'---------------------------------------------------------
Dim X As String, Nb As Integer
Nb = arrCompteNb + 100
frmCompte.Enabled = False
Do Until Not arrComptesuite
    frmCompteSuite
    If arrCompteNb > Nb Then
        X = MsgBox("Voulez-vous continuer la lecture des comptes?", vbYesNo + vbQuestion + vbDefaultButton2, arrCompteNb & " comptes lus")
        If X = vbNo Then
            Exit Do
        Else
            Nb = arrCompteNb + 100
        End If
    End If
Loop

If frmCompteNbMin > frmCompteNbMax Then cmdPagePrior_Click
frmCompte.Enabled = True

End Sub

'---------------------------------------------------------
Public Sub frmComptelstDevise()
'---------------------------------------------------------

lstDevise.Visible = True
If lstDeviseK <> 888 Then lstDeviseK = LstDictio(888, lstDevise)
'lstDevise.ListIndex = 0
lstDevise.SetFocus

End Sub

'---------------------------------------------------------
Private Sub chkDeviseCV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------

If Button = vbRightButton Then
    Set Xtextbox = txtDeviseCv
    Me.PopupMenu txtDeviseXmnu, vbPopupMenuRightButton
End If

End Sub


'---------------------------------------------------------
Public Sub valDeviseCV()
'---------------------------------------------------------

varX = CV_AttributS(Trim(txtDeviseCv.Text), CV_X2)
If Not IsNull(varX) Then
    lblErrMsg = varX
    lblErrMsgStd
Else
    If arrCompteNb > 0 Then
       arrCompteIndex = frmCompteNbMin - 1
       frmCompteItem
    End If
End If


End Sub






Public Sub Msg_Send()
Dim X As String
X = Space$(100)

Mid$(X, 1, 12) = cmdOk.Tag
Mid$(X, 13, 12) = "frmCompte   "
Mid$(X, 25, 10) = "M         "
Mid$(X, 35, 3) = arrCompte(arrCompteIndex).Devise
Mid$(X, 38, 11) = arrCompte(arrCompteIndex).Numéro

Msg_Monitor X

Unload Me


End Sub

Public Sub picCompte_Color()
Dim I As Integer
For I = 0 To frmCompteNb - 1
    picCompte(I).BackColor = picUsr.BackColor
Next I
picCompte_Index = -1


End Sub

Public Sub cmdOption_Ok()
Dim X As String
fraOption.Visible = False
X = optSolde
optSolde = " "
If optInstantané Then optSolde = "I": mAMJCours = DSys
If optVeille Then optSolde = "V": mAMJCours = DSys
If optFinDeMois Then optSolde = "M": mAMJCours = valAmjMax = dateElp("FinDeMoisP", 0, DSys)
If optFinDeMoisOpération Then optSolde = "O": mAMJCours = dateElp("FinDeMoisP", 0, DSys)
If optFinDAnnée Then optSolde = "A": mAMJCours = dateElp("FinDAnnéeP", 0, DSys)
CV_X1.OpéAmj = mAMJCours
CV_X2.OpéAmj = mAMJCours
CV_X3.OpéAmj = mAMJCours
cmdOption_Display
If X <> optSolde Then frmCompteEffacer: cmdRechercher_Click
cmdRechercher.SetFocus

End Sub

Public Sub cmdOption_Display()
lbloptSolde.ForeColor = warnUsrColor
Select Case optSolde
    Case "I": lbloptSolde = "solde instantané"
    Case "V": lbloptSolde = "solde veille"
    Case "M": lbloptSolde = "solde fin de mois"
    Case "O": lbloptSolde = "solde fin de mois:opé"
    Case "A": lbloptSolde = "solde fin d'année"
    Case Else: lbloptSolde = "solde ?"
End Select

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




Public Sub frmCompteRechercher_Option()

recCompte.MvtceJour = optSolde

If chkAmjSoldé = "1" Then
    recCompte.MvtAmj = valAmjSoldé
Else
    recCompte.MvtAmj = "00000000"
End If

If chkAnnul = "1" Then
    recCompte.chkAnnul = "1"
Else
    recCompte.chkAnnul = "0"
End If

End Sub
