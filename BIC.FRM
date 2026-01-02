VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBic 
   AutoRedraw      =   -1  'True
   Caption         =   "BIC : Correspondants"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6360
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   8910
      Picture         =   "BIC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   0
      Width           =   400
   End
   Begin VB.CommandButton cmdScan 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   3300
      Picture         =   "BIC.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   0
      Width           =   700
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5900
      TabIndex        =   27
      Top             =   0
      Width           =   2790
   End
   Begin VB.CommandButton cmdContext 
      Height          =   300
      Left            =   0
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   0
      Width           =   1100
   End
   Begin VB.PictureBox picBic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   660
      Left            =   120
      ScaleHeight     =   600
      ScaleWidth      =   9240
      TabIndex        =   19
      Top             =   360
      Width           =   9300
   End
   Begin VB.TextBox txtBicId 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4000
      MaxLength       =   11
      TabIndex        =   0
      Top             =   0
      Width           =   1900
   End
   Begin VB.CommandButton cmdSuppress 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Supprimer"
      Height          =   300
      Left            =   1100
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Enregistrer"
      Height          =   300
      Left            =   2200
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   0
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5295
      Left            =   0
      TabIndex        =   20
      Top             =   1080
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Identification"
      TabPicture(0)   =   "BIC.frx":0544
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraBic"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Références Bancaires"
      TabPicture(1)   =   "BIC.frx":0560
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraIban"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraIban 
         Height          =   4575
         Left            =   -74640
         TabIndex        =   45
         Top             =   600
         Width           =   8895
         Begin VB.ListBox lstDevise 
            Height          =   4155
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Visible         =   0   'False
            Width           =   4305
         End
         Begin VB.Frame fraRib 
            Caption         =   "R I B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   2000
            TabIndex        =   56
            Top             =   2400
            Width           =   3735
            Begin VB.TextBox txtfraRibClé 
               Height          =   285
               Left            =   2880
               MaxLength       =   2
               TabIndex        =   41
               Top             =   550
               Width           =   350
            End
            Begin VB.TextBox txtfraRibCompte 
               Height          =   285
               Left            =   1560
               MaxLength       =   11
               TabIndex        =   40
               Top             =   550
               Width           =   1200
            End
            Begin VB.TextBox txtfraRibGuichet 
               Height          =   285
               Left            =   840
               MaxLength       =   5
               TabIndex        =   39
               Top             =   550
               Width           =   615
            End
            Begin VB.TextBox txtfraRibEtablissement 
               Height          =   285
               Left            =   120
               MaxLength       =   5
               TabIndex        =   38
               Top             =   550
               Width           =   600
            End
            Begin VB.Label libfraRibCompte 
               Caption         =   "-"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   120
               TabIndex        =   62
               Top             =   1320
               Width           =   3495
            End
            Begin VB.Label lblfrRibBanque 
               Caption         =   "Banque"
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
               Left            =   120
               TabIndex        =   61
               Top             =   250
               Width           =   600
            End
            Begin VB.Label lblfraRibGuichet 
               Caption         =   "Guichet"
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
               Left            =   840
               TabIndex        =   60
               Top             =   240
               Width           =   855
            End
            Begin VB.Label lblfraRibCompte 
               Caption         =   "Compte"
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
               Left            =   1800
               TabIndex        =   59
               Top             =   250
               Width           =   735
            End
            Begin VB.Label lblfraRibClé 
               Caption         =   "Clé"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2880
               TabIndex        =   58
               Top             =   250
               Width           =   375
            End
            Begin VB.Label libfraRibDom 
               Caption         =   "-"
               Height          =   255
               Left            =   120
               TabIndex        =   57
               Top             =   960
               Width           =   3495
            End
         End
         Begin VB.CheckBox chkIban_Clear 
            Caption         =   "effacer après mise à jour"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   4200
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.Frame Frame5 
            Caption         =   "Période de Validité"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   6240
            TabIndex        =   53
            Top             =   2400
            Width           =   2415
            Begin MSMask.MaskEdBox txtIbanAmjMin 
               Height          =   300
               Left            =   900
               TabIndex        =   42
               Top             =   480
               Width           =   1250
               _ExtentX        =   2196
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   14
               Mask            =   "## - ## - ####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtIbanAmjMax 
               Height          =   300
               Left            =   900
               TabIndex        =   43
               Top             =   1200
               Width           =   1250
               _ExtentX        =   2196
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   14
               Mask            =   "## - ## - ####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Caption         =   "du"
               Height          =   300
               Left            =   120
               TabIndex        =   55
               Top             =   480
               Width           =   495
            End
            Begin VB.Label Label2 
               Caption         =   "au"
               Height          =   300
               Left            =   120
               TabIndex        =   54
               Top             =   1200
               Width           =   495
            End
         End
         Begin VB.TextBox txtIbanBicId 
            Height          =   285
            Left            =   1995
            MaxLength       =   11
            TabIndex        =   33
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtIbanDevise 
            Height          =   285
            Left            =   2000
            MaxLength       =   3
            TabIndex        =   34
            Top             =   1000
            Width           =   400
         End
         Begin VB.TextBox txtIbanPays 
            Height          =   285
            Left            =   2000
            MaxLength       =   2
            TabIndex        =   35
            Top             =   1500
            Width           =   500
         End
         Begin VB.TextBox txtIbanClé 
            Height          =   285
            Left            =   2600
            MaxLength       =   2
            TabIndex        =   36
            Top             =   1500
            Width           =   375
         End
         Begin VB.TextBox txtIbanCompte 
            Height          =   285
            Left            =   2000
            MaxLength       =   25
            TabIndex        =   37
            Top             =   2000
            Width           =   4215
         End
         Begin VB.Label lblIbanBicId 
            Caption         =   "Bic"
            Height          =   255
            Left            =   250
            TabIndex        =   52
            Top             =   500
            Width           =   975
         End
         Begin VB.Label lblIbanDevise 
            Caption         =   "Devise"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   1000
            Width           =   855
         End
         Begin VB.Label lblIbanPays 
            Caption         =   "Iban : Pays,clé"
            Height          =   255
            Left            =   250
            TabIndex        =   50
            Top             =   1500
            Width           =   1110
         End
         Begin VB.Label lblIbanCompte 
            Caption         =   "Compte"
            Height          =   255
            Left            =   250
            TabIndex        =   49
            Top             =   2000
            Width           =   975
         End
         Begin VB.Label libIbanBicId 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4080
            TabIndex        =   48
            Top             =   480
            Width           =   4695
         End
         Begin VB.Label libIbanDevise 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4100
            TabIndex        =   47
            Top             =   1000
            Width           =   4455
         End
         Begin VB.Label libIbanPays 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4100
            TabIndex        =   46
            Top             =   1500
            Width           =   3855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Période de Validité"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   6600
         TabIndex        =   28
         Top             =   3240
         Width           =   2415
         Begin MSMask.MaskEdBox txtAmjMin 
            Height          =   300
            Left            =   900
            TabIndex        =   9
            Top             =   480
            Width           =   1250
            _ExtentX        =   2196
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "## - ## - ####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtAmjMax 
            Height          =   300
            Left            =   900
            TabIndex        =   10
            Top             =   1200
            Width           =   1250
            _ExtentX        =   2196
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "## - ## - ####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblAmjMin 
            Caption         =   "du"
            Height          =   300
            Left            =   120
            TabIndex        =   32
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblAmjMax 
            Caption         =   "au"
            Height          =   300
            Left            =   120
            TabIndex        =   31
            Top             =   1200
            Width           =   495
         End
      End
      Begin VB.Frame fraBic 
         Height          =   2535
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   8955
         Begin VB.TextBox txtBdfEtablissement 
            Height          =   285
            Left            =   2200
            MaxLength       =   5
            TabIndex        =   64
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtBdfGuichet 
            Height          =   285
            Left            =   3360
            MaxLength       =   5
            TabIndex        =   63
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtPays 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   8
            Top             =   2040
            Width           =   4575
         End
         Begin VB.TextBox txtPaysId 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   12
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox txtVille 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   7
            Top             =   1680
            Width           =   4575
         End
         Begin VB.TextBox txtVilleId 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   11
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox txtIntitulé 
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   2160
            MaxLength       =   40
            TabIndex        =   6
            Top             =   1320
            Width           =   6615
         End
         Begin VB.TextBox txtTélexId 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            MaxLength       =   16
            TabIndex        =   5
            Top             =   960
            Width           =   2175
         End
         Begin VB.CheckBox chkTélex 
            Alignment       =   1  'Right Justify
            Caption         =   "Télex"
            Height          =   300
            Left            =   130
            TabIndex        =   3
            Top             =   960
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.TextBox txtSwiftId 
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2160
            MaxLength       =   11
            TabIndex        =   4
            Top             =   600
            Width           =   2175
         End
         Begin VB.CheckBox chkSwift 
            Alignment       =   1  'Right Justify
            Caption         =   "Swift"
            Height          =   300
            Left            =   130
            TabIndex        =   2
            Top             =   600
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkBanque 
            Alignment       =   1  'Right Justify
            Caption         =   "Etabl Bancaire"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.Label libBdfGuichet 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4440
            TabIndex        =   30
            Top             =   600
            Width           =   4350
         End
         Begin VB.Label libBdfEtablissement 
            Caption         =   "-"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4440
            TabIndex        =   29
            Top             =   240
            Width           =   4350
         End
         Begin VB.Label lblPays 
            Caption         =   "Pays"
            Height          =   300
            Left            =   120
            TabIndex        =   26
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblVille 
            Caption         =   "Ville"
            Height          =   255
            Left            =   130
            TabIndex        =   25
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label lblIntitulé 
            Caption         =   "Intitulé"
            Height          =   300
            Left            =   130
            TabIndex        =   24
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Langue des Avis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   3600
         TabIndex        =   22
         Top             =   3240
         Width           =   2415
         Begin VB.OptionButton optAvisFrançaisAnglais 
            Alignment       =   1  'Right Justify
            Caption         =   "(Français/Anglais)"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   300
            Value           =   -1  'True
            Width           =   1700
         End
         Begin VB.OptionButton optAvisFrançais 
            Alignment       =   1  'Right Justify
            Caption         =   "(Français)"
            Height          =   300
            Left            =   120
            TabIndex        =   14
            Top             =   750
            Width           =   1700
         End
         Begin VB.OptionButton optAvisAnglais 
            Alignment       =   1  'Right Justify
            Caption         =   "( Anglais )       "
            Height          =   300
            Left            =   120
            TabIndex        =   15
            Top             =   1200
            Width           =   1700
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Gestion Avis D'opération"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   21
         Top             =   3240
         Width           =   2655
         Begin VB.OptionButton optAvis0 
            Alignment       =   1  'Right Justify
            Caption         =   "pas d'avis"
            Height          =   255
            Left            =   200
            TabIndex        =   69
            Top             =   300
            Width           =   1800
         End
         Begin VB.OptionButton optAvis1 
            Alignment       =   1  'Right Justify
            Caption         =   "au fil de l'eau"
            Height          =   255
            Left            =   200
            TabIndex        =   68
            Top             =   720
            Value           =   -1  'True
            Width           =   1800
         End
         Begin VB.OptionButton optAvis2 
            Alignment       =   1  'Right Justify
            Caption         =   "envoi décadaire"
            Height          =   255
            Left            =   200
            TabIndex        =   67
            Top             =   1200
            Width           =   1800
         End
      End
   End
   Begin VB.Menu fraBicmnu 
      Caption         =   "fraBicmnu"
      Visible         =   0   'False
      Begin VB.Menu fraBicmnufrmBdf 
         Caption         =   "Recherche Banque de France"
      End
   End
   Begin VB.Menu fraRibmnu 
      Caption         =   "fraRibmnu"
      Visible         =   0   'False
      Begin VB.Menu fraRibmnufrmBdf 
         Caption         =   "Recherche Banque de France"
      End
      Begin VB.Menu fraRibmnufrmCompte 
         Caption         =   "Recherche compte"
      End
   End
   Begin VB.Menu mnucmdPrint 
      Caption         =   "mnucmdPrint"
      Visible         =   0   'False
      Begin VB.Menu mnucmdPrintIntitulé 
         Caption         =   "Liste triée par intitulé"
      End
      Begin VB.Menu mnucmdPrintBicid 
         Caption         =   "Liste triée par identifiant"
      End
   End
End
Attribute VB_Name = "frmBic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean, blnControl As Boolean
Dim lastActiveControl_Name, autBic As typeAuthorization
Dim X As String, I As Integer

Dim recBic As typeBic
Dim recBicIban As typeBicIban
Dim recBdfE As typeBdfE
Dim recBdfG As typeBdfG
Dim recCompte As typeCompte

Private BicMethod As String * 12
Dim valAmjMin As String * 8
Dim valAmjMax As String * 8
Dim valIbanAmjMin As String * 8
Dim valIbanAmjMax As String * 8

Private blnBic_Nostro As Boolean
Private blnBic_Addnew As Boolean
Private blnIban_Addnew As Boolean

'--------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'--------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub
'--------------------------------------------------
Private Sub chkBanque_Click()
'--------------------------------------------------
If chkBanque Then
    txtBdfEtablissement.Enabled = True
    txtBdfGuichet.Enabled = True
Else
    txtBdfEtablissement.Enabled = False
    txtBdfGuichet.Enabled = False
End If
End Sub
'--------------------------------------------------
Private Sub chkBanque_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--------------------------------------------------
Call txtBdfEtablissement_MouseDown(Button, Shift, X, Y)
End Sub

'--------------------------------------------------
Private Sub chkSwift_Click()
'--------------------------------------------------
If chkSwift Then
    txtSwiftId.Enabled = True
Else
    txtSwiftId.Enabled = False
End If
End Sub
'--------------------------------------------------
Private Sub chkTélex_Click()
'--------------------------------------------------
If chkTélex Then
    txtTélexId.Enabled = True
Else
    txtTélexId.Enabled = False
End If
End Sub
'--------------------------------------------------
Private Sub cmdContext_Click()
'--------------------------------------------------
Select Case cmdContext.Caption
    Case Is = "&Rechercher":
        If Len(Trim(txtBicId)) < 6 Then
            cmdScan_Click
        Else
            fraBic_dbRead
        End If
    Case Is = "&Abandonner": frmBic_Clear: txtBicId.SetFocus
End Select

End Sub
'--------------------------------------------------
Public Sub cmdPageNext_Click()
'--------------------------------------------------
End Sub
'--------------------------------------------------
Public Sub cmdPagePrior_Click()
'--------------------------------------------------
End Sub
'--------------------------------------------------
Private Sub cmdPrint_Click()
'--------------------------------------------------
End Sub
'--------------------------------------------------
Private Sub cmdQuit_Click()
'--------------------------------------------------
blnControl = False
arrTag_Set False
Unload Me
End Sub

Private Sub cmdPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.PopupMenu mnucmdPrint
End Sub

'--------------------------------------------------
Private Sub cmdSave_Click()
'--------------------------------------------------
Select Case SSTab.Tab
    Case Is = 0
                If blnBic_Addnew Then
                    fraBic_dbAddNew
                Else
                    fraBic_dbUpdate
                End If
    Case Is = 1
                If blnIban_Addnew Then
                    fraIban_dbAddNew
                Else
                    fraIban_dbUpdate
                End If
End Select
End Sub
'--------------------------------------------------
Private Sub cmdScan_Click()
'--------------------------------------------------
arrTag_Set False
arrBic(0) = recBic: arrBic(0).BicId = Trim(txtBicId)
arrBicIban(0) = recBicIban: arrBicIban(0).BicId = Trim(txtBicId): arrBicIban(0).IbanBicId = Trim(txtIbanBicId)
Select Case SSTab.Tab
    Case 0
        arrBic(0).Method = "Bic"
         frmBicList.Show vbModal
        If Trim(arrBic(0).Method) = "Ok" Then
            recBic = arrBic(0)
            txtBicId = recBic.BicId
            fraBic_Init
       End If
    Case 1: arrBic(0).Method = "BicIban": arrBicIban(0).Method = "BicIban"
        frmBicList.Show vbModal
        If Trim(arrBic(0).Method) = "Ok" Then
            recBic = arrBic(0)
            txtIbanBicId = recBic.BicId: txtIbanBicId_Control
       End If
        If Trim(arrBicIban(0).Method) = "Ok" Then
            recBicIban = arrBicIban(0)
            blnIban_Addnew = False: SSTab_Init
            fraIban_Init
        Else
            blnIban_Addnew = True: SSTab_Init
       End If
End Select

End Sub
'--------------------------------------------------
Private Sub cmdSuppress_Click()
'--------------------------------------------------
X = MsgBox("Voulez-vous réellement supprimer ce dossier?", vbYesNo + vbQuestion + vbDefaultButton2, "Les informations saisies seront perdues!")
If X <> vbYes Then Exit Sub

arrTag_Set False
Select Case SSTab.Tab
    Case Is = 0: fraBic_dbDelete
    Case Is = 1: fraIban_dbDelete
End Select
End Sub
'--------------------------------------------------
Private Sub Form_Activate()
'--------------------------------------------------
Set XForm = Me
End Sub
'--------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------
Select Case KeyCode
    Case Is = 13
            KeyCode = 0
            If Not SSTab.Enabled Then
                cmdContext_Click
            Else
                If Me.ActiveControl.Name = lastActiveControl_Name Then
                    cmdSave_Click
                Else
                    SendKeys "{TAB}"
                End If
            End If
    Case Is = 27
         KeyCode = 0
        If SSTab.Tab = 0 Then
            If cmdContext.Caption = "&Abandonner" Then
                cmdContext_Click
            Else
                cmdQuit_Click
            End If
        Else
            SSTab.Tab = 0 ': SSTab_Init
        End If
    Case Is = 34: cmdPageNext_Click
    Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select
End Sub

'--------------------------------------------------
Private Sub Form_Load()
'--------------------------------------------------
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
Call BiaPgmAut_Init("BIC", autBic)

SSTab.Tab = 0
recBdfEInit recBdfE
recBdfGInit recBdfG
frmBic_Clear

End Sub
'--------------------------------------------------
Public Sub fraBic_GlobalControl()
'--------------------------------------------------
    Call txt_LostFocus(txtBdfEtablissement)
    Call txt_LostFocus(txtBdfGuichet)
    Call txt_LostFocus(txtSwiftId)
    Call txt_LostFocus(txtTélexId)
    Call txt_LostFocus(txtIntitulé)
    Call txt_LostFocus(txtVille)
    Call txt_LostFocus(txtPays)
    Call txt_LostFocus(txtAmjMin)
    Call txt_LostFocus(txtAmjMax)
    frmBic_Control

    If valAmjMax > 0 _
    And valAmjMin > valAmjMax Then
        Call lstErr_AddItem(lstErr, txtAmjMin, "? date début > date fin")
    End If

End Sub
'--------------------------------------------------
Public Sub fraBic_Clear()
'--------------------------------------------------
frmBic_Reset
recBicInit recBic
chkBanque.Value = 1
chkSwift.Value = 1
chkTélex.Value = 1

optAvis1.Value = True
optAvisFrançaisAnglais = True

txtBicId = ""
txtSwiftId = ""
txtTélexId = ""
txtIntitulé = ""
txtVille = ""
txtPays = ""
txtVilleId = ""
txtPaysId = ""
txtBdfEtablissement = ""
txtBdfGuichet = ""
valAmjMin = "00000000"
valAmjMax = "00000000"
Call dateDsp("00000000", txtAmjMin)
Call dateDsp("00000000", txtAmjMax)
blnBic_Addnew = False
cmdSuppress.Visible = False
cmdSave.Visible = False
End Sub

'--------------------------------------------------
Private Sub fraBic_dbAddNew()
'--------------------------------------------------
If frmBic_GlobalControl Then
    recBic.Method = "AddNew"
    If IsNull(srvBicUpdate(recBic)) Then
        Call lstErr_Clear(lstErr, txtBicId, "Création effectuée")
        blnBic_Addnew = False
        fraIban.Enabled = True
        SSTab.Tab = 1: SSTab_Init
    Else
        frmBic_Reset
    End If
End If
End Sub
'--------------------------------------------------
Private Sub fraBic_dbDelete()
'--------------------------------------------------
'If frmBic_GlobalControl Then
    recBicInit recBic
    recBic.BicId = Trim(txtBicId)
    recBic.Method = "Delete"
    If IsNull(srvBicUpdate(recBic)) Then
        frmBic_Clear
        Call lstErr_Clear(lstErr, txtBicId, "Suppression effectuée")
        txtBicId.SetFocus
    Else
        frmBic_Reset
    End If
'End If

End Sub
'--------------------------------------------------
Private Sub fraBic_dbRead()
'--------------------------------------------------
If Trim(txtBicId) = "" Then
    Call lstErr_Clear(lstErr, txtBicId, "? code correspondant")
Else
    frmBic_Record
    recBic.Method = "SeekP0"
    srvBicMon recBic
    fraBic_Init
End If
End Sub
'--------------------------------------------------
Private Sub fraBic_dbUpdate()
'--------------------------------------------------
If frmBic_GlobalControl Then
    recBic.Method = "Update"
    If IsNull(srvBicUpdate(recBic)) Then
        Call lstErr_Clear(lstErr, txtBicId, "Modification effectuée")
        blnBic_Addnew = False
        fraIban.Enabled = True
        SSTab.Tab = 1: SSTab_Init
    Else
        frmBic_Reset
    End If
End If

End Sub

'--------------------------------------------------
Public Sub fraBic_Init()
'--------------------------------------------------
blnBic_Nostro = IIf(Trim(txtBicId) = SocBicIdNostro, True, False)
cmdContext.Caption = "&Abandonner"
txtBicId.BackColor = txtUsr.BackColor
txtBicId.Enabled = False
blnIban_Addnew = True
SSTab.Enabled = True
If Trim(recBic.Err) = "" Then
    fraIban.Enabled = True
    blnBic_Addnew = False
    frmBic_Load
    SSTab.Tab = 1
    SSTab_Init
    txtIbanBicId.SetFocus
Else
    txtVilleId = mId$(txtBicId, 7, 2)
    txtPaysId = mId$(txtBicId, 5, 2)
    txtSwiftId = txtBicId
    txtTélexId = txtBicId
    fraIban.Enabled = False
    blnBic_Addnew = True: Beep
     SSTab.Tab = 1
    SSTab_Init
   SSTab.Tab = 0
    SSTab_Init
    txtIntitulé.SetFocus
End If

'If txtPaysId <> "FR" Then
'Else
    If txtPaysId = "FR" And Trim(txtPays) = "" Then txtPays = "France"
   If chkBanque = 1 Then
        txtBdfEtablissement.Enabled = True
        txtBdfGuichet.Enabled = True
'        txtBdfEtablissement.SetFocus
'    Else
 '       txtIntitulé.SetFocus
    End If
'End If

End Sub
'--------------------------------------------------
Public Sub fraBic_Load()
'--------------------------------------------------
txtBicId = recBic.BicId
SSTab.Visible = False

chkSwift = IIf(Trim(recBic.SwiftId) = "", 0, 1)
txtSwiftId = Trim(recBic.SwiftId)
chkTélex = IIf(Trim(recBic.TélexId) = "", 0, 1)
txtTélexId = Trim(recBic.TélexId)
txtIntitulé = Trim(recBic.Intitulé)
txtVille = Trim(recBic.Ville)
txtPays = Trim(recBic.Pays)
txtVilleId = Trim(recBic.VilleId)
txtPaysId = Trim(recBic.PaysId)

valAmjMin = recBic.AmjMin
Call dateDsp(valAmjMin, txtAmjMin)
valAmjMax = recBic.AmjMax
Call dateDsp(valAmjMax, txtAmjMax)

chkBanque.Value = IIf(Trim(recBic.BdfEtablissement) = "", 0, 1)
If Trim(recBic.BdfEtablissement) <> "" Then
    txtBdfEtablissement = Format$(recBic.BdfEtablissement, "00000")
     txtBdfEtablissement_Control
End If
If Trim(recBic.BdfGuichet) <> "" Then
    txtBdfGuichet = Format$(recBic.BdfGuichet, "00000")
    txtBdfGuichet_Control
End If
Select Case recBic.optAvis
    Case Is = 0: optAvis0 = True
    Case Is = 1: optAvis1 = True
    Case Is = 2: optAvis2 = True
End Select
Select Case recBic.optAvisLangue
    Case Is = 0: optAvisFrançaisAnglais = True
    Case Is = 1: optAvisFrançais = True
    Case Is = 2: optAvisAnglais = True
End Select

Call recBic_Display(recBic, picBic)
SSTab.Visible = True
End Sub
'--------------------------------------------------
Public Sub fraBic_Record()
'--------------------------------------------------
recBicInit recBic
recBic.BicId = Trim(txtBicId)
If chkSwift Then: recBic.SwiftId = txtSwiftId
If chkTélex Then: recBic.TélexId = txtTélexId
recBic.Intitulé = txtIntitulé
recBic.Ville = txtVille
recBic.Pays = txtPays
If optAvis0 = True Then recBic.optAvis = 0
If optAvis1 = True Then recBic.optAvis = 1
If optAvis2 = True Then recBic.optAvis = 2

If optAvisFrançaisAnglais = True Then recBic.optAvisLangue = 0
If optAvisFrançais = True Then recBic.optAvisLangue = 1
If optAvisAnglais = True Then recBic.optAvisLangue = 2

recBic.VilleId = txtVilleId
recBic.PaysId = txtPaysId
If chkBanque Then
    recBic.BdfEtablissement = Format$(Val(txtBdfEtablissement), "00000")
    recBic.BdfGuichet = Format$(Val(txtBdfGuichet), "00000")
End If
recBic.AmjMin = valAmjMin
recBic.AmjMax = valAmjMax

Call recBic_Display(recBic, picBic)
End Sub
'--------------------------------------------------
Public Sub fraIban_GlobalControl()
'--------------------------------------------------
Dim X As String, Y As String, V
Call txt_LostFocus(txtIbanDevise)
Call txt_LostFocus(txtIbanBicId)
Call txt_LostFocus(txtIbanPays)
Call txt_LostFocus(txtIbanClé)
If txtIbanPays <> "FR" Then
    Call txt_LostFocus(txtIbanCompte)
Else
    Call txt_LostFocus(txtfraRibEtablissement)
    Call txt_LostFocus(txtfraRibGuichet)
    Call txt_LostFocus(txtfraRibCompte)
    Call txt_LostFocus(txtfraRibClé)
End If
Call txt_LostFocus(txtIbanAmjMin)
Call txt_LostFocus(txtIbanAmjMax)
frmBic_Control

If valIbanAmjMax > 0 _
And valIbanAmjMin > valIbanAmjMax Then
    Call lstErr_AddItem(lstErr, txtIbanAmjMin, "? date début > date fin")
End If

If Trim(txtIbanCompte) <> "" Then
    If Trim(txtIbanClé) = "" Then
        X = txtIbanPays & "00" & txtIbanCompte
            V = Iban_Calc(X, Y)
            If Not IsNull(V) Then
                Call lstErr_AddItem(lstErr, txtIbanClé, V)
            Else
                txtIbanClé = mId$(Y, 3, 2)
            End If
        End If
        
    X = txtIbanPays & txtIbanClé & txtIbanCompte
    V = Iban_Check(X)
    If Not IsNull(V) Then
        Call lstErr_AddItem(lstErr, txtIbanClé, V)
    End If
End If

End Sub
'--------------------------------------------------
Public Sub fraIban_Clear()
'--------------------------------------------------
recBicIbanInit recBicIban
txtIbanDevise = ""
txtIbanPays = ""
txtIbanBicId = ""
txtIbanClé = ""
txtIbanCompte = ""
txtfraRibEtablissement = ""
txtfraRibGuichet = ""
txtfraRibCompte = ""
txtfraRibClé = ""
valIbanAmjMin = "00000000"
valIbanAmjMax = "00000000"
Call dateDsp("00000000", txtIbanAmjMin)
Call dateDsp("00000000", txtIbanAmjMax)

libIbanDevise = ""
libIbanPays = ""
libIbanBicId = ""
libfraRibCompte = ""
libfraRibDom = ""
End Sub
'--------------------------------------------------
Private Sub fraIban_dbAddNew()
'--------------------------------------------------
If frmBic_GlobalControl Then
    recBicIban.Method = "AddNew"
    If IsNull(srvBicIbanUpdate(recBicIban)) Then
        If chkIban_Clear = 1 Then fraIban_Clear
        Call lstErr_Clear(lstErr, txtIbanDevise, "Création effectuée")
        txtIbanDevise = ""
        blnIban_Addnew = True: SSTab_Init
        txtIbanDevise.SetFocus
    Else
       fraIban_Reset
    End If
End If

End Sub
'--------------------------------------------------
Private Sub fraIban_dbDelete()
'--------------------------------------------------
'If frmBic_GlobalControl Then
    recBicIbanInit recBicIban

    recBicIban.BicId = Trim(txtBicId)
    recBicIban.IbanDevise = Trim(txtIbanDevise)
    recBicIban.IbanBicId = Trim(txtIbanBicId)
    recBicIban.Method = "Delete"
    If IsNull(srvBicIbanUpdate(recBicIban)) Then
        If chkIban_Clear = 1 Then fraIban_Clear
        Call lstErr_Clear(lstErr, txtIbanBicId, "Suppression effectuée")
        txtIbanDevise = ""
        blnIban_Addnew = True: SSTab_Init
        txtIbanBicId.SetFocus
    Else
        fraIban_Reset
    End If
'End If

End Sub
'--------------------------------------------------
Private Sub fraIban_dbUpdate()
'--------------------------------------------------
If frmBic_GlobalControl Then
    recBicIban.Method = "Update"
    If IsNull(srvBicIbanUpdate(recBicIban)) Then
        If chkIban_Clear = 1 Then fraIban_Clear
        Call lstErr_Clear(lstErr, txtIbanDevise, "Modification effectuée")
        txtIbanDevise = ""
        blnIban_Addnew = True: SSTab_Init
        txtIbanDevise.SetFocus
    Else
        fraIban_Reset
    End If
End If

End Sub
'--------------------------------------------------
Public Sub fraIban_Init()
'--------------------------------------------------
If Trim(recBicIban.Err) = "" Then
    blnIban_Addnew = False: cmdSuppress.Visible = True
    fraIban_Load
Else
    blnIban_Addnew = True: cmdSuppress.Visible = False
End If
SSTab.Tab = 1
SSTab.Enabled = True
SSTab_Init
cmdContext.Caption = "&Abandonner"

End Sub
'--------------------------------------------------
Public Sub fraIban_Load()
'--------------------------------------------------
usrColor_Set
txtIbanBicId = Trim(recBicIban.IbanBicId)
txtIbanDevise = Trim(recBicIban.IbanDevise)
txtIbanPays = Trim(recBicIban.IbanPays)
txtIbanClé = Trim(recBicIban.IbanClé)
txtIbanCompte = Trim(recBicIban.IbanCompte)
valIbanAmjMin = recBic.AmjMin
Call dateDsp(valIbanAmjMin, txtAmjMin)
valIbanAmjMax = recBic.AmjMax
Call dateDsp(valIbanAmjMax, txtAmjMax)
txtIbanBicId_Control
txtfraRib_Control
End Sub
'--------------------------------------------------
Public Sub fraIban_Record()
'--------------------------------------------------
recBicIbanInit recBicIban

recBicIban.BicId = Trim(txtBicId)
recBicIban.IbanDevise = Format$(Val(Trim(txtIbanDevise)), "000")
recBicIban.IbanBicId = Trim(txtIbanBicId)
recBicIban.IbanPays = Trim(txtIbanPays)
recBicIban.IbanClé = Trim(txtIbanClé)
recBicIban.IbanCompte = Trim(txtIbanCompte)
recBicIban.AmjMin = "00000000" 'valibanAmjMin
recBicIban.AmjMax = "00000000" 'valIbanAmjMax

End Sub
'--------------------------------------------------
Public Sub fraIban_Reset()
'--------------------------------------------------
txtIbanCompte.MaxLength = 25
End Sub
'--------------------------------------------------
Public Function frmBic_GlobalControl() As Boolean
'--------------------------------------------------
lstErr.Clear
Select Case SSTab.Tab
    Case Is = 0: fraBic_GlobalControl
    Case Is = 1: fraIban_GlobalControl
End Select

lstErrClear = False
If lstErr.ListCount = 0 Then
    frmBic_Record
    frmBic_GlobalControl = True
Else
    frmBic_GlobalControl = False
    tag_SetFocus lstErr.Tag
End If
End Function

'--------------------------------------------------
Public Sub frmBic_Clear()
'--------------------------------------------------
picBic.Cls
fraBic_Clear
fraIban_Clear
cmdScan.Visible = True
End Sub
'--------------------------------------------------
Public Sub frmBic_Control()
'--------------------------------------------------
Dim I As Integer, xobjControl As Boolean
If Not SSTab.Enabled Then Exit Sub
If lstErrClear Then lstErr.Clear

For I = 0 To arrTagNb
    If arrTag(I) Then
        xobjControl = False
        arrTag(I) = False
        X = Format(I, "000")
        For Each xobj In Me.Controls
            If X = xobj.Tag Then
                Select Case xobj.Name
                    Case "txtBdfEtablissement": txtBdfEtablissement_Control: xobjControl = True
                    Case "txtBdfGuichet": txtBdfGuichet_Control: xobjControl = True
                    Case "txtSwiftId": txtSwiftId_Control: xobjControl = True
                    Case "txtTélexId": txtTélexId_Control: xobjControl = True
                    Case "txtIntitulé": txtIntitulé_Control: xobjControl = True
                    Case "txtVille": txtVille_Control: xobjControl = True
                    Case "txtPays": txtPays_Control: xobjControl = True
                    Case "txtAmjMin": txtAmjMin_Control: xobjControl = True
                    Case "txtAmjMax": txtAmjMax_Control: xobjControl = True
                
                    Case "txtIbanBicId": txtIbanBicId_Control: xobjControl = True
                    Case "txtIbanDevise": txtIbanDevise_Control: xobjControl = True
                    Case "txtIbanPays": txtIbanPays_Control: xobjControl = True
                    Case "txtIbanClé": txtIbanClé_Control: xobjControl = True
                    Case "txtIbanCompte": txtIbanCompte_Control: xobjControl = True
                    Case "txtfraRibEtablissement": txtfraRibEtablissement_Control: xobjControl = True
                    Case "txtfraRibGuichet": txtfraRibGuichet_Control: xobjControl = True
                    Case "txtfraRibCompte": txtfraRibCompte_Control: xobjControl = True
                    Case "txtfraRibClé": txtfraRibClé_Control: xobjControl = True
                    Case "txtIbanAmjMin": txtIbanAmjMin_Control: xobjControl = True
                    Case "txtIbanAmjMax": txtibanAmjMax_Control: xobjControl = True
                End Select
            End If
            If xobjControl Then Exit For
        Next xobj
    End If
Next I

lstErrClear = True
End Sub
'--------------------------------------------------
Public Sub frmBic_Load()
'--------------------------------------------------
Select Case SSTab.Tab
    Case Is = 0: fraBic_Load
    Case Is = 1: fraIban_Load
End Select
End Sub
'--------------------------------------------------
Public Sub frmBic_Record()
'--------------------------------------------------
Select Case SSTab.Tab
    Case Is = 0: fraBic_Record
    Case Is = 1: fraIban_Record
End Select

End Sub
'--------------------------------------------------
Public Sub frmBic_Reset()
'--------------------------------------------------
cmdContext.Caption = "&Rechercher"
cmdContext.Enabled = True
cmdSave.Visible = False
cmdSuppress.Visible = False
cmdScan.Visible = True
arrTag_Set False: SSTab.Tab = 0: SSTab.Enabled = False

arrTag_Set False
lstErrClear = True
Call lstErr_Clear(lstErr, txtBicId, "code correspondant")
usrColor_Set

txtBicId.Enabled = True
End Sub
'--------------------------------------------------
Public Sub Msg_Rcv(txtMsg As String)
'--------------------------------------------------
Select Case SSTab.Tab
    Case 0
         Select Case UCase$(Trim(mId$(txtMsg, 13, 12)))
            Case Is = "FRMBDF":
                 txtBdfEtablissement = Format$(Val(mId$(txtMsg, 35, 5)), "00000")
                 txtBdfEtablissement_Control
                 txtBdfGuichet = Format$(Val(mId$(txtMsg, 40, 5)), "00000")
                 txtBdfGuichet_Control
'                 txtIntitulé.SetFocus
        End Select
    Case 1
         Select Case UCase$(Trim(mId$(txtMsg, 13, 12)))
            Case Is = "FRMBDF":
                 txtfraRibEtablissement = mId$(txtMsg, 35, 5)
                 txtfraRibGuichet = mId$(txtMsg, 40, 5)
                 txtfraRibCompte = mId$(txtMsg, 45, 11)
                 txtfraRibClé = mId$(txtMsg, 56, 2)
 '                txtfraRibGuichet_Control
            Case Is = "FRMCOMPTE":
                 txtfraRibCompte = mId$(txtMsg, 38, 11)
                 txtIbanDevise = mId$(txtMsg, 35, 3)
                 txtfraRib_IbanCompte
        End Select
 '       arrTag_Set False: txtfraRibCompte.SetFocus
End Select
End Sub

Private Sub fraRib_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtfraRibCompte.SetFocus
arrTag_Set False
fraRibmnufrmBdf.Enabled = True
fraRibmnufrmCompte.Enabled = IIf(Val(txtfraRibEtablissement) = strSocBdfE, True, False)
Me.PopupMenu fraRibmnu, vbPopupMenuRightButton
End Sub


Private Sub fraRibmnufrmBdf_Click()
X = Space$(100)
Mid$(X, 1, 12) = "frmBdf      "
Mid$(X, 13, 12) = "frmBic      "
Mid$(X, 25, 10) = Space$(10)
Mid$(X, 35, 5) = Format$(Val(txtfraRibEtablissement), "00000")
Mid$(X, 40, 5) = Format$(Val(txtfraRibGuichet), "00000")
Mid$(X, 45, 11) = txtfraRibCompte
Mid$(X, 56, 2) = Format$(Val(txtfraRibClé), "00")
Msg_Monitor X

End Sub

Private Sub fraRibmnufrmCompte_Click()
X = Space$(100)
Mid$(X, 1, 12) = "frmCompte   "
Mid$(X, 13, 12) = "frmBic      "
Mid$(X, 25, 10) = Space$(10)
Mid$(X, 38, 11) = txtfraRibCompte
Mid$(X, 35, 3) = txtIbanDevise
Msg_Monitor X

End Sub


Private Sub lblfraRibClé_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call fraRib_MouseDown(Button, Shift, X, Y)
End Sub


Private Sub lblfraRibCompte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
fraRibmnufrmCompte_Click 'Call fraRib_MouseDown(Button, Shift, X, y)
End Sub


Private Sub lblfraRibGuichet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
fraRibmnufrmBdf_Click 'Call fraRib_MouseDown(Button, Shift, X, y)
End Sub


Private Sub lblfrRibBanque_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
fraRibmnufrmBdf_Click 'Call fraRib_MouseDown(Button, Shift, X, y)
End Sub


Private Sub libfraRibCompte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call fraRib_MouseDown(Button, Shift, X, Y)
End Sub


Private Sub libfraRibDom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call fraRib_MouseDown(Button, Shift, X, Y)
End Sub


Private Sub lstDevise_Click()
txtIbanDevise = mId$(lstDevise.Text, 1, 3)
lstDevise.Visible = False
txtIbanDevise.SetFocus: SendKeys "{TAB}"
End Sub

'---------------------------------------------------------
Public Sub lstDevise_Display()
'---------------------------------------------------------

lstDevise.Visible = True
Call LstDictio(888, lstDevise)
lstDevise.SetFocus

End Sub


Private Sub lstDevise_LostFocus()
lstDevise.Visible = False
End Sub

Private Sub mnucmdPrintBicid_Click()
Dim X As String * 100
X = txtBicId
Mid$(X, 13, 12) = "frmBic      "
Mid$(X, 25, 2) = "P0"
prtBicX X

End Sub

Private Sub mnucmdPrintIntitulé_Click()
Dim X As String * 100
X = txtBicId
Mid$(X, 13, 12) = "frmBic      "
Mid$(X, 25, 2) = "L5"
prtBicX X

End Sub


'--------------------------------------------------
Private Sub SSTab_Click(PreviousTab As Integer)
'--------------------------------------------------
SSTab_Init
End Sub

'--------------------------------------------------
Public Sub SSTab_Init()
'--------------------------------------------------
If blnBic_Addnew Then SSTab.Tab = 0
arrTag_Set False
Select Case SSTab.Tab
    Case Is = 0
        If autBic.Saisir Then
            cmdSave.Visible = True
            cmdSuppress.Visible = Not blnBic_Addnew
        Else
            cmdSave.Visible = False
            cmdSuppress.Visible = False
        End If
        lstErr.Clear
        lastActiveControl_Name = "txtAmjMax"
        cmdScan.Visible = False
    Case Is = 1
        If fraIban.Enabled Then
            If autBic.Saisir Then
                cmdSave.Visible = True
                cmdSuppress.Visible = Not blnIban_Addnew
            Else
                cmdSave.Visible = False
                cmdSuppress.Visible = False
            End If
            cmdScan.Visible = True
            lastActiveControl_Name = "txtIbanAmjMax"
            If blnBic_Nostro Then
   '             txtIbanBicId = SocBicId
                txtIbanPays = "FR": txtIbanPays.Enabled = False
                txtfraRibEtablissement = strSocBdfE: txtfraRibEtablissement.Enabled = False
                txtfraRibGuichet = strSocBdfG: txtfraRibGuichet.Enabled = False
                fraRib.Enabled = True
                txtIbanCompte.Enabled = False
            Else
                txtIbanPays.Enabled = True
                txtfraRibEtablissement.Enabled = True
                txtfraRibGuichet.Enabled = True
            End If
            lstErr.Clear
'            txtIbanClé.Enabled = False
        End If
End Select

End Sub

'--------------------------------------------------
Sub txt_GotFocus(C As Control)
'--------------------------------------------------

C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub
'--------------------------------------------------
Sub txt_LostFocus(C As Control)
'--------------------------------------------------
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
If blnControl Then
    frmBic_Control
Else
    blnControl = True
End If
End Sub
'--------------------------------------------------
Public Sub txtAmjMax_Control()
'--------------------------------------------------
X = dateCtl(txtAmjMax.Text)
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, txtAmjMax, "? validité au : " & X)
Else
    valAmjMax = mId$(X, 1, 8)
    If valAmjMax <> "00000000" Then txtAmjMax.Text = dateImp(valAmjMax)
End If
End Sub

Private Sub SSTab_GotFocus()
If SSTab.Enabled Then
    Select Case SSTab.Tab
        Case 0: txtIntitulé.SetFocus
        Case 1: txtIbanBicId.SetFocus
    End Select
End If
End Sub

'--------------------------------------------------
Private Sub txtAmjMax_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtAmjMax)
End Sub
'--------------------------------------------------
Private Sub txtAmjMax_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtAmjMax)
End Sub
'--------------------------------------------------
Public Sub txtAmjMin_Control()
'--------------------------------------------------
X = dateCtl(txtAmjMin.Text)
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, txtAmjMin, "? validité du : " & X)
Else
    valAmjMin = mId$(X, 1, 8)
    If valAmjMin <> "00000000" Then txtAmjMin.Text = dateImp(valAmjMin)
End If
End Sub
'--------------------------------------------------
Private Sub txtAmjMin_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtAmjMin)
End Sub
'--------------------------------------------------
Private Sub txtAmjMin_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtAmjMin)
End Sub
'--------------------------------------------------
Private Sub txtBdfEtablissement_Change()
'--------------------------------------------------
libBdfEtablissement = ""
libBdfGuichet = ""
End Sub
'--------------------------------------------------
Public Sub txtBdfEtablissement_Control()
'--------------------------------------------------
If chkBanque _
And txtPaysId = "FR" Then
    Select Case Trim(txtBdfEtablissement)
        Case "": Call lstErr_AddItem(lstErr, txtBdfEtablissement, "? établisement")
        Case "00000"
        Case Else
            recBdfE.BdfE2 = Format$(Val(txtBdfEtablissement), "00000")
            txtBdfEtablissement = Format$(recBdfE.BdfE2, "00000")
            If Not IsNull(srvBdfEFind(recBdfE)) Then
                Call lstErr_AddItem(lstErr, txtBdfEtablissement, "? établisement inconnu")
            Else
                libBdfEtablissement = recBdfE.BdfE4
            End If
    End Select
End If
End Sub
'--------------------------------------------------
Private Sub txtBdfEtablissement_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtBdfEtablissement)
End Sub
'--------------------------------------------------
Private Sub txtBdfEtablissement_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtBdfEtablissement)
End Sub
'--------------------------------------------------
Private Sub txtBdfEtablissement_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--------------------------------------------------
If Button = vbRightButton Then
    txtBdfEtablissement.SetFocus
    arrTag_Set False
    Me.PopupMenu fraBicmnu, vbPopupMenuRightButton
End If
End Sub
'--------------------------------------------------
Private Sub txtBdfGuichet_Change()
'--------------------------------------------------
libBdfGuichet = ""
End Sub
'--------------------------------------------------
Public Sub txtBdfGuichet_Control()
'--------------------------------------------------
If chkBanque _
And txtPaysId = "FR" Then

    Select Case Trim(txtBdfGuichet)
        Case "": Call lstErr_AddItem(lstErr, txtBdfGuichet, "? guichet")
        Case "00000"
        Case Else
            recBdfG.BdfG2 = Format$(Val(txtBdfEtablissement), "00000")
            recBdfG.BdfG3 = Format$(Val(txtBdfGuichet), "00000")
            txtBdfGuichet = Format$(recBdfG.BdfG3, "00000")
            If Not IsNull(srvBdfGFind(recBdfG)) Then
                Call lstErr_AddItem(lstErr, txtBdfGuichet, "? guichet inconnu")
            Else
                libBdfGuichet = recBdfG.BdfG20
                If Trim(txtIntitulé) = "" Then txtIntitulé = Trim(recBdfE.BdfE4) & " - " & Trim(recBdfG.BdfG20)
                If Trim(txtVille) = "" Then txtVille = Trim(recBdfG.BdfG14D)
            End If
    End Select
End If

End Sub
'--------------------------------------------------
Private Sub txtBdfGuichet_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtBdfGuichet)
End Sub
'--------------------------------------------------
Private Sub txtBdfGuichet_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtBdfGuichet)
End Sub
'--------------------------------------------------
Private Sub txtBdfGuichet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'--------------------------------------------------
Call txtBdfEtablissement_MouseDown(Button, Shift, X, Y)
End Sub
'--------------------------------------------------
Private Sub fraBicmnufrmBdf_Click()
'--------------------------------------------------
X = Space$(100)
Mid$(X, 1, 12) = "frmBdf      "
Mid$(X, 13, 12) = "frmBic      "
Mid$(X, 25, 10) = Space$(10)
Mid$(X, 35, 5) = Format$(Val(txtBdfEtablissement), "00000")
Mid$(X, 40, 5) = Format$(Val(txtBdfGuichet), "00000")
Msg_Monitor X
End Sub

'--------------------------------------------------
Private Sub txtBicId_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtBicId)
End Sub
'--------------------------------------------------
Private Sub txtBicId_KeyPress(KeyAscii As Integer)
'--------------------------------------------------
KeyAscii = convUCase(KeyAscii)
End Sub
'--------------------------------------------------
Public Sub txtfraRib_Control()
'--------------------------------------------------
If txtIbanPays <> "FR" Then
    fraRib.Enabled = False
    txtIbanCompte.Enabled = True
    txtfraRibEtablissement = ""
    txtfraRibGuichet = ""
    txtfraRibCompte = ""
    txtfraRibClé = ""
Else
    txtfraRibEtablissement.ForeColor = txtUsr.ForeColor
    txtfraRibEtablissement.BackColor = txtUsr.BackColor
    txtfraRibGuichet.ForeColor = txtUsr.ForeColor
    txtfraRibGuichet.BackColor = txtUsr.BackColor
    txtfraRibCompte.ForeColor = txtUsr.ForeColor
    txtfraRibCompte.BackColor = txtUsr.BackColor
    txtfraRibClé.ForeColor = txtUsr.ForeColor
    txtfraRibClé.BackColor = txtUsr.BackColor
    fraRib.Enabled = True
    txtIbanCompte.Enabled = False
    If Trim(txtIbanCompte) <> "" Then
        If IsNumeric(mId$(txtIbanCompte, 1, 5)) _
        And IsNumeric(mId$(txtIbanCompte, 6, 5)) Then
            txtfraRibEtablissement = Trim(mId$(txtIbanCompte, 1, 5))
            txtfraRibGuichet = (mId$(txtIbanCompte, 6, 5))
            txtfraRibCompte = RTrim(mId$(txtIbanCompte, 11, 11))
            txtfraRibClé = mId$(txtIbanCompte, 22, 2)
'           If IsNumeric(Mid$(recBicIban.IbanCompte, 22, 2)) Then txtfraRibClé = Mid$(recBicIban.IbanCompte, 22, 2)
            txtfraRibEtablissement_Control
            txtfraRibGuichet_Control
            txtfraRibCompte_Control
        End If
    End If
End If
End Sub

'--------------------------------------------------
Public Sub txtfraRib_IbanCompte()
'--------------------------------------------------
txtIbanCompte = ""
txtIbanCompte = Format$(txtfraRibEtablissement, "00000") _
            & Format$(txtfraRibGuichet, "00000") _
            & Format$(txtfraRibCompte, "@@@@@@@@@@@") _
            & Format$(txtfraRibClé, "00")
End Sub
'--------------------------------------------------
Public Sub txtfraRibClé_Control()
'--------------------------------------------------
Dim IbanE As String
X = Trim(txtfraRibCompte)
If X = "" Then Exit Sub
'If Not IsNumeric(Trim(txtfraribClé)) Then
'    Call lstErr_AddItem(lstErr, txtfraribClé, "? clé rib numérique")
'Else
    I = RibClé(txtfraRibEtablissement, txtfraRibGuichet, X, IbanE)
    If I <> Val(txtfraRibClé) Then
        X = "? clé rib : " & Format$(I, "00")
        Call lstErr_AddItem(lstErr, txtfraRibClé, X)
    End If
'End If
txtfraRib_IbanCompte
End Sub
'--------------------------------------------------
Private Sub txtfraRibClé_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtfraRibClé)
End Sub
'--------------------------------------------------
Private Sub txtfraRibClé_KeyPress(KeyAscii As Integer)
'--------------------------------------------------
num_KeyAscii KeyAscii
End Sub

'--------------------------------------------------
Private Sub txtfraribClé_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtfraRibClé)
End Sub
'--------------------------------------------------
Public Sub txtfraRibCompte_Control()
'--------------------------------------------------
If Trim(txtfraRibCompte) = "" Then
    txtfraRibClé = ""
    Call lstErr_AddItem(lstErr, txtfraRibCompte, "? Compte")
Else
 txtfraRibCompte = Rib_Compte(txtfraRibCompte)
 If Val(txtfraRibEtablissement) = strSocBdfE Then
   
    If Not IsNumeric(txtfraRibCompte) Then
        Call lstErr_AddItem(lstErr, txtfraRibCompte, "? compte BIA")
    Else
            recCompteInit recCompte
            recCompte.Method = "SeekL1"
            recCompte.Société = SocId$
            recCompte.Agence = SocAgence$
            recCompte.Devise = Format$(Val(txtIbanDevise), "000")
            recCompte.Numéro = Format$(Val(txtfraRibCompte), "00000000000")
            If IsNull(srvCompteFind(recCompte)) Then
                libfraRibCompte = recCompte.Intitulé
                If recCompte.Situation = "A" Or recCompte.Situation = "E" Then
                    Call lstErr_AddItem(lstErr, txtfraRibCompte, "? compte annulé")
                Else
                    txtfraRibClé = Format$(RibClé(txtfraRibEtablissement, txtfraRibGuichet, Trim(txtfraRibCompte), IbanE), "00")
                End If
            Else
                Call lstErr_AddItem(lstErr, txtfraRibCompte, "? compte inconnu")
            End If
        End If
    End If
End If
txtfraRib_IbanCompte
End Sub

Private Sub txtfraRibClé_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then Call fraRib_MouseDown(Button, Shift, X, Y)
End Sub

'--------------------------------------------------
Private Sub txtfraRibCompte_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtfraRibCompte)
End Sub

'--------------------------------------------------
Private Sub txtfraRibCompte_KeyPress(KeyAscii As Integer)
'--------------------------------------------------
KeyAscii = convUCase(KeyAscii)
End Sub

'--------------------------------------------------
Private Sub txtfraRibCompte_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtfraRibCompte)
End Sub

'--------------------------------------------------
Public Sub txtfraRibEtablissement_Control()
'--------------------------------------------------
On Error Resume Next
If Trim(txtfraRibEtablissement) = "" Then Call lstErr_AddItem(lstErr, txtfraRibEtablissement, "? Etablissement")
txtfraRib_IbanCompte
        
If Val(txtfraRibEtablissement) = strSocBdfE Then
    txtfraRibClé.Enabled = False
Else
    txtfraRibClé.Enabled = True
End If
End Sub

Private Sub txtfraRibCompte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then Call fraRib_MouseDown(Button, Shift, X, Y)
End Sub

'--------------------------------------------------
Private Sub txtfraRibEtablissement_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtfraRibEtablissement)
End Sub

'--------------------------------------------------
Private Sub txtfraRibEtablissement_KeyPress(KeyAscii As Integer)
'--------------------------------------------------
num_KeyAscii KeyAscii
End Sub

'--------------------------------------------------
Private Sub txtfraRibEtablissement_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtfraRibEtablissement)
End Sub

'--------------------------------------------------
Public Sub txtfraRibGuichet_Control()
'--------------------------------------------------
If Trim(txtfraRibGuichet) = "" Then Call lstErr_AddItem(lstErr, txtfraRibGuichet, "? Guichet"): recBdfG.BdfG2 = Format$(Val(txtfraRibEtablissement), "00000"): Exit Sub
recBdfG.BdfG2 = Format$(Val(txtfraRibEtablissement), "00000")
recBdfG.BdfG3 = Format$(Val(txtfraRibGuichet), "00000")
If Not IsNull(srvBdfGFind(recBdfG)) Then
    Call lstErr_AddItem(lstErr, txtfraRibGuichet, "? guichet inconnu")
    libfraRibDom = "?"
Else
    libfraRibDom = recBdfG.BdfG20
End If
txtfraRib_IbanCompte
End Sub

Private Sub txtfraRibEtablissement_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then Call fraRib_MouseDown(Button, Shift, X, Y)
End Sub

'--------------------------------------------------
Private Sub txtfraRibGuichet_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtfraRibGuichet)
End Sub

'--------------------------------------------------
Private Sub txtfraRibGuichet_KeyPress(KeyAscii As Integer)
'--------------------------------------------------
num_KeyAscii KeyAscii
End Sub

'--------------------------------------------------
Private Sub txtfraRibGuichet_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtfraRibGuichet)
End Sub

'--------------------------------------------------
Public Sub txtibanAmjMax_Control()
'--------------------------------------------------
X = dateCtl(txtIbanAmjMax.Text)
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, txtIbanAmjMax, "? validité au : " & X)
Else
    valIbanAmjMax = mId$(X, 1, 8)
    If valIbanAmjMax <> "00000000" Then txtIbanAmjMax.Text = dateImp(valIbanAmjMax)
End If
End Sub

Private Sub txtfraRibGuichet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then Call fraRib_MouseDown(Button, Shift, X, Y)
End Sub

'--------------------------------------------------
Private Sub txtIbanAmjMAx_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtIbanAmjMax)
End Sub
'--------------------------------------------------
Private Sub txtIbanAmjMAx_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtIbanAmjMax)
End Sub
'--------------------------------------------------
Public Sub txtIbanAmjMin_Control()
'--------------------------------------------------
X = dateCtl(txtIbanAmjMin.Text)
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, txtIbanAmjMin, "? validité du : " & X)
Else
    valIbanAmjMin = mId$(X, 1, 8)
    If valIbanAmjMin <> "00000000" Then txtIbanAmjMin.Text = dateImp(valIbanAmjMin)
End If
End Sub
'--------------------------------------------------
Private Sub txtIbanAmjMin_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtIbanAmjMin)
End Sub
'--------------------------------------------------
Private Sub txtIbanAmjmin_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtIbanAmjMin)
End Sub

'--------------------------------------------------
Public Sub txtIbanBicId_Control()
'--------------------------------------------------
If Trim(txtIbanBicId) <> "" Then
    recBicInit recBic
    recBic.Method = "SeekP0"
    recBic.BicId = Trim(txtIbanBicId)

    If Not IsNull(srvBicFind(recBic)) Then
        Call lstErr_AddItem(lstErr, txtIbanBicId, "? correspondant inconnu")
    Else
        libIbanBicId = recBic.Intitulé
        If Trim(txtIbanPays) = "" Then txtIbanPays = recBic.PaysId
        If Not blnBic_Nostro And recBic.PaysId = "FR" Then
           txtfraRibEtablissement = recBic.BdfEtablissement
           txtfraRibGuichet = recBic.BdfGuichet
        End If
    End If
End If
End Sub

'--------------------------------------------------
Private Sub txtIbanBicId_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtIbanBicId)
End Sub
'--------------------------------------------------
Private Sub txtIbanBicId_KeyPress(KeyAscii As Integer)
'--------------------------------------------------
KeyAscii = convUCase(KeyAscii)
End Sub
'--------------------------------------------------
Private Sub txtIbanBicId_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtIbanBicId)
End Sub
'--------------------------------------------------
Public Sub txtIbanClé_Control()
'--------------------------------------------------
End Sub
'--------------------------------------------------
Private Sub txtIbanClé_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtIbanClé)
End Sub
'--------------------------------------------------
Private Sub txtibanclé_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtIbanClé)
End Sub
'--------------------------------------------------
Public Sub txtIbanCompte_Control()
'--------------------------------------------------

End Sub

'--------------------------------------------------
Private Sub txtIbanCompte_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtIbanCompte)
End Sub
'--------------------------------------------------
Private Sub txtIbanCompte_KeyPress(KeyAscii As Integer)
'--------------------------------------------------
KeyAscii = convUCase(KeyAscii)
End Sub
'--------------------------------------------------
Private Sub txtIbanCompte_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtIbanCompte)
End Sub
'--------------------------------------------------
Private Sub txtIbanDevise_Change()
'--------------------------------------------------
'blnIban_Addnew = False: cmdsuppress.visible = False
End Sub

'--------------------------------------------------
Public Sub txtIbanDevise_Control()
'--------------------------------------------------
libIbanDevise = Trim(DevX(Val(txtIbanDevise)))
If libIbanDevise = "" Then
    Call lstErr_AddItem(lstErr, txtIbanDevise, "? devise")
End If
End Sub

Private Sub txtIbanDevise_Click()
lstDevise_Display
End Sub

'--------------------------------------------------
Private Sub txtIbanDevise_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtIbanDevise)
End Sub
'--------------------------------------------------
Private Sub txtIbanDevise_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtIbanDevise)
End Sub
'--------------------------------------------------
Public Sub txtIbanPays_Control()
'--------------------------------------------------
If Trim(txtIbanPays) = "" Then Call lstErr_AddItem(lstErr, txtIbanPays, "? pays")
txtfraRib_Control
End Sub
'--------------------------------------------------
Private Sub txtIbanPays_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtIbanPays)
End Sub
'--------------------------------------------------
Private Sub txtIbanPays_KeyPress(KeyAscii As Integer)
'--------------------------------------------------
KeyAscii = convUCase(KeyAscii)
End Sub
'--------------------------------------------------
Private Sub txtIbanPays_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtIbanPays)
End Sub
'--------------------------------------------------
Public Sub txtIntitulé_Control()
'--------------------------------------------------
If Trim(txtIntitulé) = "" Then Call lstErr_AddItem(lstErr, txtIntitulé, "? Intitulé")
End Sub
'--------------------------------------------------
Private Sub txtIntitulé_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtIntitulé)
End Sub
'--------------------------------------------------
Private Sub txtIntitulé_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtIntitulé)
End Sub
'--------------------------------------------------
Public Sub txtPays_Control()
'--------------------------------------------------
If Trim(txtPays) = "" Then Call lstErr_AddItem(lstErr, txtPays, "? pays")
End Sub
'--------------------------------------------------
Private Sub txtPays_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtPays)
End Sub
'--------------------------------------------------
Private Sub txtPays_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtPays)
End Sub
'--------------------------------------------------
Private Sub txtPaysId_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtPaysId)
End Sub
'--------------------------------------------------
Private Sub txtPaysId_KeyPress(KeyAscii As Integer)
'--------------------------------------------------
KeyAscii = convUCase(KeyAscii)
End Sub
'--------------------------------------------------
Private Sub txtPaysId_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtPaysId)
End Sub
'--------------------------------------------------
Public Sub txtSwiftId_Control()
'--------------------------------------------------
If chkSwift Then
    If Trim(txtSwiftId) = "" Then
        Call lstErr_AddItem(lstErr, txtSwiftId, "? swift")
    End If
End If
End Sub
'--------------------------------------------------
Private Sub txtSwiftId_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtSwiftId)
End Sub
'--------------------------------------------------
Private Sub txtSwiftId_KeyPress(KeyAscii As Integer)
'--------------------------------------------------
KeyAscii = convUCase(KeyAscii)
End Sub
'--------------------------------------------------
Private Sub txtSwiftId_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtSwiftId)
End Sub
'--------------------------------------------------
Public Sub txtTélexId_Control()
'--------------------------------------------------
If chkTélex Then
    If Trim(txtTélexId) = "" Then
        Call lstErr_AddItem(lstErr, txtTélexId, "? télex")
    End If
End If
End Sub
'--------------------------------------------------
Private Sub txtTélexId_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtTélexId)
End Sub
'--------------------------------------------------
Private Sub txtTélexId_KeyPress(KeyAscii As Integer)
'--------------------------------------------------
KeyAscii = convUCase(KeyAscii)
End Sub
'--------------------------------------------------
Private Sub txtTélexId_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtTélexId)
End Sub
'--------------------------------------------------
Public Sub txtVille_Control()
'--------------------------------------------------
If Trim(txtVille) = "" Then Call lstErr_AddItem(lstErr, txtVille, "? ville")
End Sub
'--------------------------------------------------
Private Sub txtVille_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtVille)
End Sub
'--------------------------------------------------
Private Sub txtVille_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtVille)
End Sub
'--------------------------------------------------
Private Sub txtVilleId_GotFocus()
'--------------------------------------------------
Call txt_GotFocus(txtVilleId)
End Sub
'--------------------------------------------------
Private Sub txtVilleId_KeyPress(KeyAscii As Integer)
'--------------------------------------------------
KeyAscii = convUCase(KeyAscii)
End Sub
'--------------------------------------------------
Private Sub txtVilleId_LostFocus()
'--------------------------------------------------
Call txt_LostFocus(txtVilleId)
End Sub

