VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBdf 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   " BDF : Etablissements, Guichets, Comptes"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8970
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6570
   ScaleWidth      =   8970
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   8250
      Picture         =   "BDF.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   120
      Width           =   400
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ok"
      Height          =   375
      Left            =   7850
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   120
      Width           =   400
   End
   Begin VB.TextBox txtRibClé 
      Height          =   300
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtCompte 
      Height          =   300
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   2
      Top             =   1080
      Width           =   1700
   End
   Begin VB.TextBox txtGuichet 
      Height          =   300
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   1
      Top             =   650
      Width           =   855
   End
   Begin VB.Frame fraEtablissement 
      Height          =   2220
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   8700
      Begin VB.TextBox txtEtablissement 
         Height          =   300
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.Label libIban 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4200
         TabIndex        =   45
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label lblRibClé 
         Caption         =   "  Clé"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   42
         Top             =   1150
         Width           =   525
      End
      Begin VB.Label lblErrMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "ErrMsg"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1440
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label libCompte 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4200
         TabIndex        =   29
         Top             =   1800
         Width           =   4380
      End
      Begin VB.Label libRacine 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4200
         TabIndex        =   28
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label libGuichet 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   27
         Top             =   600
         Width           =   5910
      End
      Begin VB.Label libEtablissement 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2400
         TabIndex        =   26
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label lblCompte 
         Caption         =   "Compte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   1150
         Width           =   1005
      End
      Begin VB.Label lblGuichet 
         Caption         =   "Guichet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   750
         Width           =   1000
      End
      Begin VB.Label lblEtablissement 
         Caption         =   "Etablissement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   280
         Width           =   1000
      End
   End
   Begin TabDlg.SSTab sstabBdf 
      Height          =   4095
      Left            =   0
      TabIndex        =   19
      Top             =   2400
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Etablissement"
      TabPicture(0)   =   "BDF.frx":0102
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblBdfE2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblBdfE10CP"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblBdfE4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblsstab0"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lstBdfE"
      Tab(0).Control(5)=   "cmdBdfERechercher"
      Tab(0).Control(6)=   "cmdBdfEDétail"
      Tab(0).Control(7)=   "cmdBdfEOk"
      Tab(0).Control(8)=   "chkBdfEEffacer"
      Tab(0).Control(9)=   "chkBdfEPageParPage"
      Tab(0).Control(10)=   "cmdBdfESuite"
      Tab(0).Control(11)=   "txtBdfE2"
      Tab(0).Control(12)=   "txtBdfE10CP"
      Tab(0).Control(13)=   "txtBdfE4"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "     Guichet     "
      TabPicture(1)   =   "BDF.frx":011E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblBdfG3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblBdfG14CP"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblBdfG2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblsstab1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdBdfGOk"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdBdfGDétail"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdBdfGRechercher"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdBdfGSuite"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lstBdfG"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chkBdfGeffacer"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "chkBdfGPageParPage"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtBdfG2"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtBdfG3"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtBdfG14CP"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.TextBox txtBdfG14CP 
         Height          =   300
         Left            =   5880
         MaxLength       =   5
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtBdfG3 
         Height          =   300
         Left            =   3120
         MaxLength       =   5
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtBdfG2 
         Height          =   300
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtBdfE4 
         Height          =   300
         Left            =   -70080
         MaxLength       =   30
         TabIndex        =   6
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtBdfE10CP 
         Height          =   300
         Left            =   -71760
         MaxLength       =   5
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtBdfE2 
         Height          =   300
         Left            =   -73800
         MaxLength       =   5
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkBdfGPageParPage 
         Caption         =   "Page / Page"
         Height          =   255
         Left            =   6840
         TabIndex        =   41
         Top             =   3764
         Width           =   1680
      End
      Begin VB.CheckBox chkBdfGeffacer 
         Caption         =   "Effacer à chaque recherche"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   3764
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CommandButton cmdBdfESuite 
         Caption         =   "....>>"
         Height          =   300
         Left            =   -69360
         TabIndex        =   8
         Top             =   44
         Width           =   1000
      End
      Begin VB.CheckBox chkBdfEPageParPage 
         Caption         =   "Page / Page"
         Height          =   255
         Left            =   -68160
         TabIndex        =   39
         Top             =   3764
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin VB.CheckBox chkBdfEEffacer 
         Caption         =   "Effacer à chaque recherche"
         Height          =   255
         Left            =   -74880
         TabIndex        =   38
         Top             =   3764
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.ListBox lstBdfG 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2595
         Left            =   120
         TabIndex        =   32
         Top             =   1004
         Width           =   8415
      End
      Begin VB.CommandButton cmdBdfGSuite 
         Caption         =   "....>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5550
         TabIndex        =   15
         Top             =   44
         Width           =   1000
      End
      Begin VB.CommandButton cmdBdfGRechercher 
         Caption         =   "&Rechercher"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4550
         TabIndex        =   14
         Top             =   44
         Width           =   1000
      End
      Begin VB.CommandButton cmdBdfGDétail 
         Caption         =   "&Détail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6570
         TabIndex        =   16
         Top             =   44
         Width           =   1000
      End
      Begin VB.CommandButton cmdBdfGOk 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7570
         TabIndex        =   17
         Top             =   44
         Width           =   1000
      End
      Begin VB.CommandButton cmdBdfEOk 
         Caption         =   "&OK"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -67380
         TabIndex        =   10
         Top             =   44
         Width           =   1000
      End
      Begin VB.CommandButton cmdBdfEDétail 
         Caption         =   "&Détail"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68340
         TabIndex        =   9
         Top             =   44
         Width           =   1000
      End
      Begin VB.CommandButton cmdBdfERechercher 
         Caption         =   "&Rechercher"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -70360
         TabIndex        =   7
         Top             =   44
         Width           =   1000
      End
      Begin VB.ListBox lstBdfE 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2400
         Left            =   -74760
         TabIndex        =   20
         Top             =   1080
         Width           =   8415
      End
      Begin VB.Label lblsstab0 
         Caption         =   "Etablissement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   -74970
         TabIndex        =   37
         Top             =   50
         Width           =   1215
      End
      Begin VB.Label lblsstab1 
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   1250
         TabIndex        =   36
         Top             =   30
         Width           =   1150
      End
      Begin VB.Label lblBdfG2 
         Caption         =   "Etablissement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblBdfG14CP 
         Caption         =   "Code Postal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4800
         TabIndex        =   34
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblBdfG3 
         Caption         =   "Guichet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblBdfE4 
         Caption         =   "ou Alpha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70800
         TabIndex        =   31
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblBdfE10CP 
         Caption         =   "Code Postal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -72720
         TabIndex        =   25
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblBdfE2 
         Caption         =   "Etablissement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74880
         TabIndex        =   24
         Top             =   480
         Width           =   1000
      End
   End
   Begin VB.Menu lstBdfEmnu 
      Caption         =   "BdfE"
      Visible         =   0   'False
      Begin VB.Menu lstBDFEmnuOK 
         Caption         =   "OK"
      End
      Begin VB.Menu lstBDFEmnuDétail 
         Caption         =   "Détail"
      End
      Begin VB.Menu lstBDFEmnuRechercher 
         Caption         =   "Rechercher"
      End
      Begin VB.Menu lstBDFEmnuSuite 
         Caption         =   "Suite"
      End
      Begin VB.Menu lstBdfEmnuImprimerLigne 
         Caption         =   "Imprimer ligne"
      End
      Begin VB.Menu lstBdfEmnuImprimerListe 
         Caption         =   "Imprimer Liste"
      End
   End
   Begin VB.Menu txtBdfE4mnu 
      Caption         =   "txtBdfE4mnu"
      Visible         =   0   'False
      Begin VB.Menu txtBdfE4mnuRechercher 
         Caption         =   "Rechercher"
      End
   End
   Begin VB.Menu txtBdfE2mnu 
      Caption         =   "txtBdfE2mnu"
      Visible         =   0   'False
      Begin VB.Menu txtBdfE2mnuRechercher 
         Caption         =   "Rechercher"
      End
   End
   Begin VB.Menu txtBdfE10CPmnu 
      Caption         =   "txtBdfE10Cpmnu"
      Visible         =   0   'False
      Begin VB.Menu txtBdfE10CPmnuRechercher 
         Caption         =   "Rechercher"
      End
   End
   Begin VB.Menu txtEtablissementmnu 
      Caption         =   "txtEtablissementmnu"
      Visible         =   0   'False
      Begin VB.Menu txtEtablissementmnuDétail 
         Caption         =   "Détail"
      End
   End
   Begin VB.Menu txtGuichetmnu 
      Caption         =   "txtGuichetmnu"
      Visible         =   0   'False
      Begin VB.Menu txtGuichetmnuDétail 
         Caption         =   "Détail"
      End
   End
   Begin VB.Menu txtComptemnu 
      Caption         =   "txtComptemnu"
      Visible         =   0   'False
      Begin VB.Menu txtComptemnuDétail 
         Caption         =   "Détail"
      End
   End
   Begin VB.Menu txtRibClémnu 
      Caption         =   "txtRibClémnu"
      Visible         =   0   'False
      Begin VB.Menu txtRibClémnuCalcul 
         Caption         =   "Calcul"
      End
      Begin VB.Menu txtRibClémnuContrôle 
         Caption         =   "Contrôle"
      End
   End
   Begin VB.Menu lstBdfGmnu 
      Caption         =   "lstBdfG"
      Visible         =   0   'False
      Begin VB.Menu lstBdfGmnuOk 
         Caption         =   "Ok"
      End
      Begin VB.Menu lstBdfGmnuDétail 
         Caption         =   "Détail"
      End
      Begin VB.Menu lstBdfgmnuSuite 
         Caption         =   "Suite"
      End
      Begin VB.Menu lstBdfGmnuImprimerLigne 
         Caption         =   "Imprimer ligne"
      End
      Begin VB.Menu lstBdfGmnuImprimerListe 
         Caption         =   "Imprimer Liste"
      End
   End
   Begin VB.Menu txtBdfg2mnu 
      Caption         =   "txtBdfg2mnu"
      Visible         =   0   'False
      Begin VB.Menu txtBdfG2mnuRechercher 
         Caption         =   "Rechercher"
      End
   End
   Begin VB.Menu txtBdfG3mnu 
      Caption         =   "txtBdfG3mnu"
      Visible         =   0   'False
      Begin VB.Menu txtBdfG3mnuRechercher 
         Caption         =   "Rechercher"
      End
   End
   Begin VB.Menu txtBdfG14CPmnu 
      Caption         =   "txtBdfG14CPmnu"
      Visible         =   0   'False
      Begin VB.Menu txtBdfG14CPmnuRechercher 
         Caption         =   "txtBdfG14CpmnuRechercher"
      End
   End
End
Attribute VB_Name = "frmBdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean, blnControl As Boolean

Private recBdfE As typeBdfE
Private BdfEMethod As String * 12

Private recBdfG As typeBdfG
Private BdfGMethod As String * 12

Private recCompte As typeCompte


Private Sub cmdbdfesuite_Click()

If chkBdfEPageParPage = 0 Then
    BdfESuiteEtFin
Else
    BdfESuite
End If

End Sub

'---------------------------------------------------------
Private Sub cmdBdfGDétail_Click()
'---------------------------------------------------------
Dim X As String

arrBdfGIndex = lstBdfG.ListIndex + 1
If arrBdfGIndex > 0 And arrBdfGIndex <= arrBdfGNb Then
    frmBdfG.Show vbModeless
    frmBdfG.Visible = True
    X = frmBdfG.Caption
    AppActivate X
End If
End Sub

'---------------------------------------------------------
Public Sub cmdBdfGOk_Click()
'---------------------------------------------------------
recBdfG = arrBdfG(lstBdfG.ListIndex + 1)
txtEtablissement.Text = Format$(recBdfG.BdfG2, "00000")
errTag = valEtablissement
txtGuichet.Text = Format$(recBdfG.BdfG3, "00000")
errTag = valGuichet
If IsNull(errTag) Then
   'txtGuichet.Enabled = True
    'txtCompte.Enabled = True
    txtCompte.SetFocus
End If
End Sub



'---------------------------------------------------------
Private Sub cmdBdfGRechercher_Click()
'---------------------------------------------------------

BdfGMethod = ""
lstBdfGRechercher
If chkBdfGPageParPage = 0 Then: BdfGSuiteEtFin

End Sub



'---------------------------------------------------------
Private Sub BdfGSuite()
'---------------------------------------------------------
Dim X As String

If arrBdfGNb > 0 Then
    recBdfG = arrBdfG(arrBdfGNb)
    X = "+"
Else
   cmdBdfGEffacer
   lstBdfGSql
    X = " "
End If
If Trim(mId$(BdfGMethod, 5, 2)) = "" Then Mid$(BdfGMethod, 5, 2) = "P0"
recBdfG.Method = "Snap" & mId$(BdfGMethod, 5, 2) & X
 
If IsNull(SrvBdfGMon(recBdfG)) Then
    Call lstBdfGAddItem
End If

Call lstBdfGEnabled
End Sub


'---------------------------------------------------------
Private Sub lstBdfGAddItem()
'---------------------------------------------------------
Dim X As String, Y As String
Dim I As Integer, K As Integer

For K = lstBdfG.ListCount + 1 To arrBdfGNb
X = arrBdfG(K).BdfG2 & "   " _
   & arrBdfG(K).BdfG3 & "   " _
    & Trim(arrBdfG(K).BdfG20)
Do While TextWidth(X) < 5040
    X = X & "."
Loop

lstBdfG.AddItem X _
              & Chr$(9) & arrBdfG(K).BdfG14CP _
              & " " & arrBdfG(K).BdfG14D
Next K

lstBdfG.ListIndex = lstBdfG.ListCount - 1
cmdBdfGSuite.Enabled = arrBdfGsuite
End Sub





Private Sub cmdBdfGSuite_Click()

If chkBdfGPageParPage = 0 Then
    BdfGSuiteEtFin
Else
    BdfGSuite
End If

End Sub

Private Sub cmdOk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
blnControl = False

End Sub

'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
Dim Msg As String

Select Case sstabBdf.Tab
    Case 0
        Msg = Format$(1, "000000") & Format$(arrBdfENb, "000000")
        prtBdfE Msg
    Case 1
        Msg = Format$(1, "000000") & Format$(arrBdfGNb, "000000")
        prtBdfG Msg
End Select
End Sub

Private Sub Form_Activate()

Set XForm = Me
If txtEtablissement.Enabled Then txtEtablissement.SetFocus

End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 27: cmdQuit_Click
'    Case Is = 34: cmdPageNext_Click
'    Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
    Case Is = 13: SendKeys "{TAB}"
End Select
lblErrMsg.Visible = False

End Sub

'---------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------
arrBdfE(0) = recBdfE
arrBdfG(0) = recBdfG
'On Error Resume Next
'frmElp.SetFocus

End Sub


Private Sub cmdOk_Click()
Dim X As String
X = Space$(100)

Mid$(X, 1, 12) = cmdOk.Tag
Mid$(X, 13, 12) = "frmBdf      "
Mid$(X, 25, 10) = Space$(10)
Mid$(X, 35, 5) = Format$(Val(txtEtablissement), "00000")
Mid$(X, 40, 5) = Format$(Val(txtGuichet), "00000")
Mid$(X, 45, 11) = mId$(txtCompte, 1, 11)
Mid$(X, 56, 2) = mId$(txtRibClé, 1, 2)

Msg_Monitor X

Unload Me

End Sub

'---------------------------------------------------------
Private Sub lstBdfE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------

Me.PopupMenu lstBdfEmnu, vbPopupMenuRightButton

End Sub

'---------------------------------------------------------
Private Sub lstBDFEmnuDétail_Click()
'---------------------------------------------------------
cmdBdfEDétail_Click
End Sub

'---------------------------------------------------------
Private Sub lstBdfEmnuImprimerLigne_Click()
'---------------------------------------------------------
Dim Msg As String

Msg = Format$(lstBdfE.ListIndex + 1, "000000") & Format$(lstBdfE.ListIndex + 1, "000000")
prtBdfE Msg

End Sub

'---------------------------------------------------------
Private Sub lstBdfEmnuImprimerListe_Click()
'---------------------------------------------------------

cmdPrint_Click

End Sub


'---------------------------------------------------------
Private Sub lstBDFEmnuOK_Click()
'---------------------------------------------------------
cmdBdfEOk_Click
sstabBdf = 1

End Sub

'---------------------------------------------------------
Private Sub lstBDFEmnuRechercher_Click()
'---------------------------------------------------------
cmdBdfERechercher_Click
End Sub


'---------------------------------------------------------
Private Sub lstBDFEmnuSuite_Click()
'---------------------------------------------------------
cmdbdfesuite_Click
End Sub


'---------------------------------------------------------
Private Sub lstBdfG_DblClick()
'---------------------------------------------------------
 Call cmdBdfGDétail_Click
End Sub


'---------------------------------------------------------
Private Sub lstBdfG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
Me.PopupMenu lstBdfGmnu, vbPopupMenuRightButton
End Sub

'---------------------------------------------------------
Private Sub lstBdfGmnuDétail_Click()
'---------------------------------------------------------
cmdBdfGDétail_Click
End Sub

'---------------------------------------------------------
Private Sub lstBdfGmnuImprimerLigne_Click()
'---------------------------------------------------------
Dim Msg As String

Msg = Format$(lstBdfG.ListIndex + 1, "000000") & Format$(lstBdfG.ListIndex + 1, "000000")
prtBdfG Msg

End Sub

'---------------------------------------------------------
Private Sub lstBdfGmnuImprimerListe_Click()
'---------------------------------------------------------

cmdPrint_Click

End Sub


'---------------------------------------------------------
Private Sub lstBdfGmnuOK_Click()
'---------------------------------------------------------
cmdBdfGOk_Click
End Sub


'---------------------------------------------------------
Private Sub lstBdfGmnuSuite_Click()
'---------------------------------------------------------
cmdBdfGSuite_Click
End Sub


Private Sub sstabBdf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case sstabBdf.Tab
    Case Is = 1
        cmdBdfGRechercher.Default = True
    Case Else 'Is = 0
        cmdBdfERechercher.Default = True
End Select

End Sub


'---------------------------------------------------------
Private Sub lblBdfE10cp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
If Button = vbRightButton Then
    Me.PopupMenu txtBdfE10CPmnu, vbPopupMenuRightButton
End If

End Sub

'---------------------------------------------------------
Private Sub txtBdfE10CPmnuRechercher_Click()
'---------------------------------------------------------
BdfEMethod = "SnapL1"
lstBdfERechercher

End Sub

'---------------------------------------------------------
Private Sub lblBdfE2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
If Button = vbRightButton Then
    Me.PopupMenu txtBdfE2mnu, vbPopupMenuRightButton
End If

End Sub

'---------------------------------------------------------
Private Sub txtBdfE2mnuRechercher_Click()
'---------------------------------------------------------
BdfEMethod = "SeekP0"
lstBdfERechercher

End Sub

'---------------------------------------------------------
Private Sub lblbdfE4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------

If Button = vbRightButton Then
    Me.PopupMenu txtBdfE4mnu, vbPopupMenuRightButton
End If
End Sub

'---------------------------------------------------------
Private Sub txtbdfE4_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------
KeyAscii = convUCase(KeyAscii)

End Sub

'---------------------------------------------------------
Private Sub txtBdfE4mnuRechercher_Click()
'---------------------------------------------------------
BdfEMethod = "SnapL2"
lstBdfERechercher

End Sub

'-------------------------------------------------'
Private Sub txtBdfG14cp_GotFocus()
'-------------------------------------------------'
cmdBdfGRechercher.Default = True
Call txtGetFocus(txtBdfG14CP)

End Sub


'-------------------------------------------------'
Private Sub txtBdfG14cp_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
Call num_KeyAscii(KeyAscii)
End Sub


'-------------------------------------------------'
Private Sub txtBdfG14cp_LostFocus()
'-------------------------------------------------'
txtBdfG14CP.BackColor = txtUsr.BackColor

End Sub

'-------------------------------------------------'
Private Sub txtBdfG14CPmnuRechercher_Click()
'-------------------------------------------------'
BdfGMethod = "SnapL1"
lstBdfGRechercher

End Sub

'---------------------------------------------------------
Private Sub txtBdfG2_GotFocus()
'---------------------------------------------------------
cmdBdfGRechercher.Default = True
Call txtGetFocus(txtBdfG2)

End Sub


'---------------------------------------------------------
Private Sub txtBdfG2_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------
Call num_KeyAscii(KeyAscii)
End Sub


'---------------------------------------------------------
Private Sub txtBdfG2_LostFocus()
'---------------------------------------------------------
txtBdfG2.BackColor = txtUsr.BackColor

End Sub

'-------------------------------------------------'
Private Sub txtBdfG2mnuRechercher_Click()
'-------------------------------------------------'
BdfGMethod = "SnapP0"
lstBdfGRechercher

End Sub

'-------------------------------------------------'
Private Sub txtBdfG3_GotFocus()
'-------------------------------------------------'
cmdBdfGRechercher.Default = True
Call txtGetFocus(txtbdfG3)

End Sub


'-------------------------------------------------'
Private Sub txtBdfG3_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
Call num_KeyAscii(KeyAscii)
End Sub


'-------------------------------------------------'
Private Sub txtBdfG3_LostFocus()
'-------------------------------------------------'
txtbdfG3.BackColor = txtUsr.BackColor

End Sub




'---------------------------------------------------------
Private Sub lstBdfGSql()
'---------------------------------------------------------
Dim G2 As Long, G3 As Long, G14CP As Long
Dim G2max As Long, G3max As Long, G14CPmax As Long

    
G14CP = Val(txtBdfG14CP)
If G14CP = 0 Then
    G14CPmax = 99999
Else
    If G14CP < 100 Then
        G14CP = G14CP * 1000
    End If
    If G14CP Mod 1000 = 0 Then
        G14CPmax = G14CP + 999
    Else
        G14CPmax = G14CP
    End If
End If

G2 = Val(txtBdfG2)
If G2 = 0 Then
    G2max = 99999
Else
    G2max = G2
End If

G3 = Val(txtbdfG3)
If G3 = 0 Then
    G3max = 99999
Else
    G3max = G3
End If

If G2 = 0 Then
    BdfGMethod = "SnapL1"
    txtBdfG14CP.Enabled = True
    txtBdfG14CP.SetFocus
Else
    If G3 = 0 Then
        BdfGMethod = "SnapL2"
        txtbdfG3.Enabled = True
        txtbdfG3.SetFocus
    Else
        BdfGMethod = "SeekP0"
        txtbdfG3.Enabled = True
        txtbdfG3.SetFocus
        G3max = 99999
        G14CP = 99999
    End If
End If

recBdfG.Method = BdfGMethod
recBdfG.BdfG2 = Format$(G2, "00000")
recBdfG.BdfG3 = Format$(G3, "00000")
recBdfG.BdfG14CP = Format$(G14CP, "00000")

arrBdfG(0) = recBdfG
arrBdfG(0).BdfG2 = Format$(G2max, "00000")
arrBdfG(0).BdfG3 = Format$(G3max, "00000")
arrBdfG(0).BdfG14CP = Format$(G14CPmax, "00000")
End Sub

'---------------------------------------------------------
Private Sub lstBdfGEnabled()
'---------------------------------------------------------

If lstBdfG.ListCount > -1 Then
    cmdBdfGDétail.Enabled = True
    cmdBdfGOk.Enabled = True
End If

End Sub

'---------------------------------------------------------
Public Sub cmdBdfGEffacer()
'---------------------------------------------------------
lstBdfG.Clear
arrBdfGNb = 0
recBdfGInit recBdfG
arrBdfGsuite = True
cmdBdfGSuite.Enabled = True 'arrBdfGsuite
End Sub




'---------------------------------------------------------
Private Sub cmdBdfEDétail_Click()
'---------------------------------------------------------
Dim X As String

arrBdfEIndex = lstBdfE.ListIndex + 1
If arrBdfEIndex > 0 And arrBdfEIndex <= arrBdfENb Then
    frmBdfE.Show vbModeless
    frmBdfE.Visible = True
    X = frmBdfE.Caption
    AppActivate X
End If
End Sub

'---------------------------------------------------------
Public Sub cmdBdfEOk_Click()
'---------------------------------------------------------

recBdfE = arrBdfE(lstBdfE.ListIndex + 1)
txtEtablissement.Text = Format$(recBdfE.BdfE2, "00000")
errTag = valEtablissement
If IsNull(errTag) Then
    Call valEtablissementOk
    txtEtablissement.Enabled = True
    txtGuichet.Enabled = True
'   txtGuichet.SetFocus
End If

End Sub

'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me
End Sub

'---------------------------------------------------------
Private Sub cmdBdfERechercher_Click()
'---------------------------------------------------------

BdfEMethod = ""
lstBdfERechercher
If chkBdfEPageParPage = 0 Then: BdfESuiteEtFin

End Sub



'---------------------------------------------------------
Private Sub BdfESuite()
'---------------------------------------------------------
Dim X As String

If arrBdfENb > 0 Then
    recBdfE = arrBdfE(arrBdfENb)
    X = "+"
Else
    cmdBdfEEffacer
    lstBdfESql
    X = " "
End If
    
If Trim(mId$(BdfEMethod, 5, 2)) = "" Then Mid$(BdfEMethod, 5, 2) = "P0"

recBdfE.Method = "Snap" & mId$(BdfEMethod, 5, 2) & X
 
If IsNull(SrvBdfEMon(recBdfE)) Then
    Call lstBdfEAddItem
End If

Call lstBdfEEnabled
End Sub
'---------------------------------------------------------
Public Sub BdfESuiteEtFin()
'---------------------------------------------------------
 Dim X As String, Nb As Integer
Nb = arrBdfENb + 100
   
Do Until Not arrBdfEsuite
    BdfESuite

    If arrBdfENb > Nb Then
        X = MsgBox("Voulez-vous continuer la lecture des guichets?", vbYesNo + vbQuestion + vbDefaultButton2, arrBdfENb & " mouvements lus")
        If X = vbNo Then
            Exit Do
        Else
            Nb = arrBdfENb + 100
        End If
    End If
Loop

End Sub



'---------------------------------------------------------
Private Sub lstBdfEAddItem()
'---------------------------------------------------------
Dim X As String
Dim I As Integer, K As Integer

For K = lstBdfE.ListCount + 1 To arrBdfENb
X = arrBdfE(K).BdfE2 & "   " _
    & Trim(arrBdfE(K).BdfE4)
Do While TextWidth(X) < 5700
    X = X & "."
Loop

lstBdfE.AddItem X _
               & Chr$(9) & arrBdfE(K).BdfE10CP _
               & " " & arrBdfE(K).BdfE10D

Next K

lstBdfE.ListIndex = lstBdfE.ListCount - 1
cmdBdfESuite.Enabled = arrBdfEsuite
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

BdfEMethod = ""
Call cmdBdfEEffacer

BdfGMethod = ""
Call cmdBdfGEffacer

End Sub

'---------------------------------------------------------
Private Sub lstBdfE_DblClick()
'---------------------------------------------------------
Call cmdBdfEDétail_Click
End Sub



'-------------------------------------------------'
Private Sub txtBdfE10cp_GotFocus()
'-------------------------------------------------'
cmdBdfERechercher.Default = True
Call txtGetFocus(txtBdfE10CP)

End Sub


'-------------------------------------------------'
Private Sub txtBdfE10cp_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
Call num_KeyAscii(KeyAscii)
End Sub


'-------------------------------------------------'
Private Sub txtBdfE10cp_LostFocus()
'-------------------------------------------------'
txtBdfE10CP.BackColor = txtUsr.BackColor

End Sub

'-------------------------------------------------'
Private Sub txtBdfE2_GotFocus()
'-------------------------------------------------'
cmdBdfERechercher.Default = True
Call txtGetFocus(txtbdfE2)

End Sub


'-------------------------------------------------'
Private Sub txtBdfE2_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
Call num_KeyAscii(KeyAscii)
End Sub


'-------------------------------------------------'
Private Sub txtBdfE2_LostFocus()
'-------------------------------------------------'
txtbdfE2.BackColor = txtUsr.BackColor

End Sub

'---------------------------------------------------------
Private Sub txtbdfE4_GotFocus()
'---------------------------------------------------------
cmdBdfERechercher.Default = True
Call txtGetFocus(txtbdfE4)

End Sub


'---------------------------------------------------------
Private Sub txtbdfE4_LostFocus()
'---------------------------------------------------------
txtbdfE4.BackColor = txtUsr.BackColor

End Sub


'-------------------------------------------------'
Private Sub txtBdfG3mnuRechercher_Click()
'-------------------------------------------------'
BdfGMethod = "SnapL2"
lstBdfGRechercher

End Sub

'-------------------------------------------------'
Private Sub txtCompte_GotFocus()
'-------------------------------------------------'
Call txtGetFocus(txtCompte)
txtRibClé.Enabled = False

BdfEGDisabled

End Sub


'-------------------------------------------------'
Private Sub txtCompte_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
libCompte = "": libRacine = "" ': txtRibClé = Format$("00", txtRibClé.Mask)
'Call RibCompte(txtCompte, KeyAscii)
End Sub


'-------------------------------------------------'
Private Sub txtCompte_LostFocus()
'-------------------------------------------------'
 If Not blnControl Then Exit Sub
txtCompte = Rib_Compte(txtCompte)
txtCompte.BackColor = txtUsr.BackColor
'If Val(txtCompte) = 0 Then
'    txtEtablissement_GotFocus
'Else
    If txtCompte.Enabled Then
        errTag = valCompte
        If IsNull(errTag) Then
            txtRibClé.Enabled = True
            txtRibClé.SetFocus
        Else
            If Val(txtCompte) <> 0 Then txtCompte_GotFocus
        End If
    End If
'End If

End Sub

'---------------------------------------------------------
Private Sub lblCompte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
If Button = vbRightButton Then Me.PopupMenu txtComptemnu, vbPopupMenuRightButton
End Sub

Private Sub txtCompte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then Me.PopupMenu txtComptemnu, vbPopupMenuRightButton
End Sub

'---------------------------------------------------------
Private Sub txtComptemnuDétail_Click()
'---------------------------------------------------------
'sstabBdf.Tab = 2
'cmdCompteDétail_Click

End Sub

'-------------------------------------------------'
Private Sub txtEtablissement_GotFocus()
'-------------------------------------------------'
txtEtablissement.TabIndex = 0
txtGuichet.TabIndex = 1
txtCompte.TabIndex = 2
txtRibClé.TabIndex = 3

Call txtGetFocus(txtEtablissement)
' sstabBdf.Tab = 0
'txtGuichet.Enabled = False
'txtCompte.Enabled = False
'txtRibClé.Enabled = False

BdfEGDisabled

End Sub


'-------------------------------------------------'
Private Sub txtEtablissement_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
libEtablissement = ""
Call num_KeyAscii(KeyAscii)
End Sub


'-------------------------------------------------'
Private Sub txtEtablissement_LostFocus()
'-------------------------------------------------'
 If Not blnControl Then Exit Sub
txtEtablissement.BackColor = txtUsr.BackColor
If txtEtablissement.Enabled Then
    errTag = valEtablissement
    If IsNull(errTag) Then
        txtGuichet.Enabled = True
 '       txtGuichet.SetFocus
    End If
End If

End Sub


'---------------------------------------------------------
Private Sub lblEtablissement_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------

End Sub

Private Sub txtEtablissement_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then Me.PopupMenu txtEtablissementmnu, vbPopupMenuRightButton
End Sub

'-------------------------------------------------'
Private Sub txtEtablissementmnuDétail_Click()
'-------------------------------------------------'
sstabBdf.Tab = 0
cmdBdfEDétail_Click

End Sub

'-------------------------------------------------'
Private Sub txtGuichet_GotFocus()
'-------------------------------------------------'

Call txtGetFocus(txtGuichet)
'sstabBdf.Tab = 1
'txtCompte.Enabled = False
'txtRibClé.Enabled = False

BdfEGDisabled

End Sub


'-------------------------------------------------'
Private Sub txtGuichet_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
libGuichet = ""
Call num_KeyAscii(KeyAscii)
End Sub


'-------------------------------------------------'
Private Sub txtGuichet_LostFocus()
'-------------------------------------------------'
 If Not blnControl Then Exit Sub

txtGuichet.BackColor = txtUsr.BackColor
If txtGuichet.Enabled Then
    errTag = valGuichet
    If IsNull(errTag) Then
        'txtCompte.Enabled = True
 ' txtCompte.SetFocus
  '  Else
   '     txtEtablissement.SetFocus
    End If
End If
End Sub

'---------------------------------------------------------
Private Sub lblGuichet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
If Button = vbRightButton Then
    Me.PopupMenu txtGuichetmnu, vbPopupMenuRightButton
End If

End Sub

Private Sub txtGuichet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then Me.PopupMenu txtGuichetmnu, vbPopupMenuRightButton

End Sub

'---------------------------------------------------------
Private Sub txtGuichetmnuDétail_Click()
'---------------------------------------------------------
sstabBdf.Tab = 1
cmdBdfGDétail_Click

End Sub

'-------------------------------------------------'
Private Sub txtRibClé_GotFocus()
'-------------------------------------------------'
Call txtGetFocus(txtRibClé)
BdfEGDisabled
End Sub


'-------------------------------------------------'
Private Sub txtRibClé_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
Call num_KeyAscii(KeyAscii)
End Sub


'-------------------------------------------------'
Private Sub txtRibClé_LostFocus()
'-------------------------------------------------'
 If Not blnControl Then Exit Sub

txtRibClé.BackColor = txtUsr.BackColor
If txtRibClé.Enabled Then
    errTag = valRibClé
    If IsNull(errTag) Then
        txtEtablissement.SetFocus
    Else
        txtRibClé_GotFocus
    End If
End If


End Sub



'---------------------------------------------------------
Private Sub lstBdfESql()
'---------------------------------------------------------
Dim E2 As Long, E4 As String, E10CP As Long
Dim E2max As Long, E4Max As String, E10CPMax As Long

    
E10CP = Val(txtBdfE10CP)
If E10CP = 0 Then
    E10CPMax = 99999
Else
    If E10CP < 100 Then
        E10CP = E10CP * 1000
    End If
    If E10CP Mod 1000 = 0 Then
        E10CPMax = E10CP + 999
    Else
        E10CPMax = E10CP
    End If
End If

E2 = Val(txtbdfE2)
E2max = 99999

E4 = UCase$(Trim(txtbdfE4))
E4Max = E4 & "99999999999"

If Trim(BdfEMethod) = "" Then
    If E2 <> 0 Then
        BdfEMethod = "SeekP0"
        txtbdfE2.Enabled = True
        txtbdfE2.SetFocus
    Else
        If E10CP <> 0 Then
            BdfEMethod = "SnapL1"
            txtBdfE10CP.Enabled = True
            txtBdfE10CP.SetFocus
        Else
            BdfEMethod = "SnapL2"
            txtbdfE4.Enabled = True
            txtbdfE4.SetFocus
        End If
    End If
End If

recBdfE.Method = BdfEMethod
recBdfE.BdfE2 = Format$(E2, "00000")
recBdfE.BdfE4 = E4
recBdfE.BdfE10CP = Format$(E10CP, "00000")

arrBdfE(0) = recBdfE
arrBdfE(0).BdfE2 = Format$(E2max, "00000")
arrBdfE(0).BdfE4 = E4Max
arrBdfE(0).BdfE10CP = Format$(E10CPMax, "00000")

End Sub

'---------------------------------------------------------
Private Sub lstBdfEEnabled()
'---------------------------------------------------------

'cmdBdfESuite.Enabled = True
If lstBdfE.ListCount > -1 Then
    cmdBdfEDétail.Enabled = True
    cmdBdfEOk.Enabled = True
End If

End Sub

'---------------------------------------------------------
Public Sub cmdBdfEEffacer()
'---------------------------------------------------------
lstBdfE.Clear
arrBdfENb = 0
recBdfEInit recBdfE
arrBdfEsuite = True
cmdBdfESuite.Enabled = True 'arrBdfEsuite
End Sub

'---------------------------------------------------------
Private Function valEtablissement()
'---------------------------------------------------------
Dim X As String, V As Variant
   
valEtablissement = txtEtablissement.Tag
libGuichet = ""
X = Format$(Val(txtEtablissement), "00000")
If X = "00000" Then
    X = ""
    With recBdfE
        .BdfE4 = ""
        .BdfE2 = ""
    End With

Else
    recBdfE.BdfE2 = X
    V = srvBdfEFind(recBdfE)
    If IsNull(V) Then
       valEtablissement = Null
       lstBdfEAddItem
       If arrBdfEIndex > 1 Then lstBdfE.ListIndex = arrBdfEIndex - 1
    Else
        recBdfE.BdfE2 = ""
        With libEtablissement
            .Caption = "Inconnu"
            .ForeColor = errUsr.ForeColor
        End With
        txtEtablissement.ForeColor = errUsr.ForeColor
        Exit Function
    '!!!!!!!!!!!========
    End If
End If
'If Trim(X) <> "" Then txtEtablissement.Text = Format$(X, txtEtablissement.Mask)
With txtEtablissement
    .ForeColor = txtUsr.ForeColor
    .BackColor = txtUsr.BackColor
End With

With libEtablissement
    .Caption = recBdfE.BdfE4
    .ForeColor = libUsr.ForeColor
    .Visible = True
End With

If arrBdfENb = 0 Then
    txtbdfE2 = txtEtablissement
End If
If arrBdfGNb = 0 Then
    txtBdfG2 = txtEtablissement
End If

End Function

'---------------------------------------------------------
Private Function valGuichet()
'---------------------------------------------------------
Dim X As String
Dim V

valGuichet = txtGuichet.Tag

If Val(txtEtablissement) = 0 Then
    libEtablissement = "Préciser l'établissement"
    txtEtablissement.SetFocus
    Exit Function
'!!!========
End If

X = Format$(Val(txtGuichet), "00000")
If X = "00000" Then
    With recBdfG
        .BdfG20 = ""
        .BdfG2 = ""
        .BdfG3 = ""
    End With

Else
    recBdfG.BdfG2 = recBdfE.BdfE2
    recBdfG.BdfG3 = X
    V = srvBdfGFind(recBdfG)
    
    If IsNull(V) Then
       valGuichet = Null
        lstBdfGAddItem
       lstBdfG.ListIndex = arrBdfGIndex - 1
    Else
            recBdfG.BdfG3 = ""
            With libGuichet
                .Caption = "Inconnu"
                .ForeColor = errUsr.ForeColor
            End With
            txtGuichet.ForeColor = errUsr.ForeColor
            Exit Function
'!!!!!!!!!!!========
    End If
End If
With txtGuichet
    .Text = Format$(X, "@@@@@")
    .ForeColor = txtUsr.ForeColor
    .BackColor = txtUsr.BackColor
End With

With libGuichet
    .Caption = recBdfG.BdfG20
    .ForeColor = libUsr.ForeColor
    .Visible = True
End With


End Function

'---------------------------------------------------------
Public Function valRibClé()
'---------------------------------------------------------
Dim I As Integer
Dim X As String

valRibClé = Null

If Val(txtEtablissement) = 0 Then
    libEtablissement = "Préciser l'établissement"
    txtEtablissement.SetFocus
    Exit Function
'!!!========
End If

If Val(txtGuichet) = 0 Then
    libGuichet = "Préciser le guichet"
    txtGuichet.Enabled = True
    txtGuichet.SetFocus
    Exit Function
'!!!========
End If

If Trim(txtCompte) = "" Then
    libRacine = "Préciser le compte"
'    txtCompte.SetFocus
    Exit Function
'!!!========
End If
I = RibClé(txtEtablissement, txtGuichet, txtCompte, IbanE)
libIban = Iban_Print(IbanE)
If Trim(txtRibClé) = "" Then
    txtRibClé.ForeColor = libUsr.ForeColor
    txtRibClé = Format$(I, "00")
Else
    If Val(txtRibClé) <> I Then
        lblErrMsg = "  Clé :" & Format$(I, "00")
        Call elpErrMsg(lblRibClé, lblErrMsg)
        valRibClé = txtGuichet.Tag
    End If
End If

End Function

'---------------------------------------------------------
Private Sub lblRibClé_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------
If Button = vbRightButton Then
    Me.PopupMenu txtRibClémnu, vbPopupMenuRightButton
End If

End Sub



'---------------------------------------------------------
Public Sub valEtablissementOk()
'---------------------------------------------------------
On Error Resume Next
    txtBdfG2 = txtEtablissement.Text
    txtGuichet = "": txtbdfG3 = "": txtBdfG14CP = ""
    txtCompte = "": 'txtRibClé = Format$("00", txtRibClé.Mask)
    
'    If Trim(txtBdfG2) <> "" Then
'        sstabBdf.Tab = 1
'        cmdBdfGEffacer
''''''  cmdBdfGRechercher
'        If arrBdfGNb = 1 Then
'            Call cmdBdfGOk_Click
'        Else
'           txtGuichet.SetFocus
'        End If
'    End If
End Sub

'---------------------------------------------------------
Private Sub lstBdfERechercher()
'---------------------------------------------------------
cmdBdfEDétail.Enabled = False
cmdBdfEOk.Enabled = False

If chkBdfEEffacer Then cmdBdfEEffacer

lstBdfESql

If IsNull(SrvBdfEMon(recBdfE)) Then
    If mId$(BdfEMethod, 1, 4) = "Seek" Then
        Call arrBdfEAddItem(recBdfE)
        Call lstBdfEAddItem
    Else
        Call lstBdfEAddItem
    End If
End If

Call lstBdfEEnabled

End Sub

'---------------------------------------------------------
Private Sub txtRibClémnuCalcul_Click()
'---------------------------------------------------------
txtRibClé = "00"
valRibClé

End Sub

'---------------------------------------------------------
Private Sub txtRibClémnuContrôle_Click()
'---------------------------------------------------------
 If IsNull(valRibClé) Then
    lblErrMsg = "Clé RIB conforme"
    Call elpErrMsg(txtRibClé, lblErrMsg)
End If
 
End Sub


'---------------------------------------------------------
Private Sub lstBdfGRechercher()
'---------------------------------------------------------

cmdBdfGDétail.Enabled = False
cmdBdfGOk.Enabled = False
If chkBdfGeffacer Then cmdBdfGEffacer

lstBdfGSql

If IsNull(SrvBdfGMon(recBdfG)) Then
    If mId$(BdfGMethod, 1, 4) = "Seek" Then
        Call arrBdfGAddItem(recBdfG)
        Call lstBdfGAddItem
    Else
        Call lstBdfGAddItem
    End If
End If

Call lstBdfGEnabled


End Sub


'---------------------------------------------------------
Public Function valCompte()
'---------------------------------------------------------
 
 valCompte = Null
 'libRacine.Visible = False
 'libCompte.Visible = False
 
 If Val(txtEtablissement) = SocBdfE _
 And Val(txtGuichet) = SocBdfG Then
    
    If Not IsNumeric(txtCompte) Then
         txtCompte.ForeColor = errUsr.ForeColor
         libRacine = "le numéro de compte est bizarre"
    Else
            recCompteInit recCompte
            recCompte.Method = "FindFirstL5 "
            recCompte.Société = SocId$
            recCompte.Agence = SocAgence$
            recCompte.Devise = "000"
            recCompte.Numéro = Format$(Val(txtCompte), "00000000000")
            recCompte.BiaTyp = "000"
            recCompte.BiaNum = "00"
            If IsNull(srvCompteMon(recCompte)) Then
                With txtCompte
                .Text = Format$(recCompte.Numéro, "@@@@@@@@@@@")
                .ForeColor = txtUsr.ForeColor
                .BackColor = txtUsr.BackColor
                End With

                With libCompte
                .Caption = recCompte.Intitulé2
                .ForeColor = libUsr.ForeColor
                .Visible = True
                End With
                With libRacine
                .Caption = recCompte.Intitulé
                .ForeColor = libUsr.ForeColor
                .Visible = True
                End With

            Else
                valCompte = txtCompte.Tag
                With libRacine
                .Visible = True
                .Caption = "Inconnu"
                .ForeColor = errUsr.ForeColor
                End With
                txtCompte.ForeColor = errUsr.ForeColor
        End If
    End If
End If

End Function


'---------------------------------------------------------
Public Sub BdfEGDisabled()
'---------------------------------------------------------

cmdBdfERechercher.Default = False
'txtBdfE2.Enabled = False
'txtbdfE4.Enabled = False
'txtBdfE10cp.Enabled = False

cmdBdfGRechercher.Default = False
'txtBdfG2.Enabled = False
'txtBdfG3.Enabled = False
'txtBdfG14cp.Enabled = False

End Sub

Public Sub BdfGSuiteEtFin()
'---------------------------------------------------------
Dim X As String, Nb As Integer
Nb = arrBdfGNb + 100
   
Do Until Not arrBdfGsuite
    BdfGSuite

    If arrBdfGNb > Nb Then
        X = MsgBox("Voulez-vous continuer la lecture des guichets?", vbYesNo + vbQuestion + vbDefaultButton2, arrBdfGNb & " mouvements lus")
        If X = vbNo Then
            Exit Do
        Else
            Nb = arrBdfGNb + 100
        End If
    End If
Loop


End Sub

'---------------------------------------------------------
Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------
blnControl = True
If mId$(X, 13, 12) = "" Then
    cmdOk.Visible = False
    MousePointer = 0
Else
    cmdOk.Visible = True
    MousePointer = 10
'    MouseIcon = frmElp.imgBiaMouse
'    cmdOk.Visible = IIf(Mid$(X, 13, 12) = "", False, True)
    cmdOk.Tag = mId$(X, 13, 12)
    If Val(mId$(X, 35, 5)) <> 0 Then
        txtEtablissement.Text = Format$(mId$(X, 35, 5), "00000")
        txtbdfE2 = txtEtablissement
        txtBdfG2 = txtEtablissement
    End If
    If Val(mId$(X, 40, 5)) <> 0 Then
        txtGuichet = Format$(Val(mId$(X, 40, 5)), "00000")
    End If
    txtCompte.Text = mId$(X, 45, 11)
    If Val(mId$(X, 56, 2)) <> 0 Then txtRibClé.Text = Format$(mId$(X, 56, 2), "00")
End If
End Sub






