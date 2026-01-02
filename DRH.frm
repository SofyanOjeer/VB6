VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDRH 
   Caption         =   "DRH"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9285
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   4920
      TabIndex        =   1
      Top             =   0
      Width           =   3825
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8760
      Picture         =   "DRH.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   0
      TabIndex        =   25
      Top             =   480
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Salariés"
      TabPicture(0)   =   "DRH.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAbsences"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mouvement"
      TabPicture(1)   =   "DRH.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgSalariéMvt"
      Tab(1).Control(1)=   "fraMouvement"
      Tab(1).Control(2)=   "cmdMouvementHisto"
      Tab(1).Control(3)=   "cmdSalatiéMvtPrint"
      Tab(1).Control(4)=   "cmdMouvementSaisir"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Fiche salarié"
      TabPicture(2)   =   "DRH.frx":013A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraSalarié"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Calendrier"
      TabPicture(3)   =   "DRH.frx":0156
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraFérié"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Tickets rest"
      TabPicture(4)   =   "DRH.frx":0172
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdTRDisquette"
      Tab(4).Control(1)=   "cmdTRControl"
      Tab(4).Control(2)=   "cmdTrOk"
      Tab(4).Control(3)=   "fraTR"
      Tab(4).Control(4)=   "fgTR"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "Impression"
      TabPicture(5)   =   "DRH.frx":018E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "SSTab2"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.CommandButton cmdMouvementSaisir 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Saisir un mouvement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -69360
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   5640
         Width           =   3435
      End
      Begin VB.CommandButton cmdSalatiéMvtPrint 
         Caption         =   "Imprimer mvts"
         Height          =   375
         Left            =   -72960
         TabIndex        =   88
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CommandButton cmdMouvementHisto 
         Caption         =   "Historique"
         Height          =   375
         Left            =   -74880
         TabIndex        =   87
         Top             =   5640
         Width           =   1815
      End
      Begin VB.CommandButton cmdTRDisquette 
         Caption         =   "Copier sur disquette"
         Height          =   645
         Left            =   -68640
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   5160
         Width           =   2475
      End
      Begin VB.CommandButton cmdTRControl 
         Caption         =   "Etat de contrôle"
         Height          =   645
         Left            =   -68640
         TabIndex        =   68
         Top             =   3600
         Width           =   2475
      End
      Begin VB.CommandButton cmdTrOk 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Validation"
         Height          =   645
         Left            =   -68640
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   4320
         Width           =   2475
      End
      Begin VB.Frame fraSalarié 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   42
         Top             =   480
         Width           =   9015
         Begin VB.CommandButton cmdSalariéOK 
            BackColor       =   &H00C0FFC0&
            Caption         =   "&Ok"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   7320
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   4200
            Width           =   1395
         End
         Begin VB.Frame fraSalariéDivers 
            Height          =   1575
            Left            =   120
            TabIndex        =   58
            Top             =   3840
            Width           =   6975
            Begin VB.TextBox txtRéfInterne 
               Height          =   285
               Left            =   5040
               MaxLength       =   16
               TabIndex        =   12
               Top             =   360
               Width           =   1815
            End
            Begin VB.TextBox txtCompte 
               Height          =   285
               Left            =   1920
               TabIndex        =   11
               Top             =   240
               Width           =   1275
            End
            Begin VB.TextBox txtBureau 
               Height          =   285
               Left            =   1920
               MaxLength       =   3
               TabIndex        =   16
               Top             =   1200
               Width           =   600
            End
            Begin VB.TextBox txtTéléphone1 
               Height          =   285
               Left            =   1920
               MaxLength       =   3
               TabIndex        =   13
               Top             =   720
               Width           =   600
            End
            Begin VB.TextBox txtTéléphone2 
               Height          =   285
               Left            =   2640
               MaxLength       =   3
               TabIndex        =   14
               Top             =   720
               Width           =   600
            End
            Begin VB.TextBox txtTéléphone3 
               Height          =   285
               Left            =   3480
               MaxLength       =   3
               TabIndex        =   15
               Top             =   720
               Width           =   600
            End
            Begin VB.Label libUpd 
               Caption         =   "-"
               Height          =   255
               Left            =   4440
               TabIndex        =   63
               Top             =   1200
               Width           =   2415
            End
            Begin VB.Label lblRéfInterne 
               Caption         =   "Référence interne"
               Height          =   255
               Left            =   3600
               TabIndex        =   62
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label lblCompte 
               Caption         =   "Compte"
               Height          =   255
               Left            =   240
               TabIndex        =   61
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblBureau 
               Caption         =   "Bureau"
               Height          =   255
               Left            =   240
               TabIndex        =   60
               Top             =   1200
               Width           =   855
            End
            Begin VB.Label lblTéléphone 
               Caption         =   "Téléphone"
               Height          =   255
               Left            =   240
               TabIndex        =   59
               Top             =   720
               Width           =   975
            End
         End
         Begin VB.Frame fraSalariéId 
            Height          =   3615
            Left            =   120
            TabIndex        =   43
            Top             =   120
            Width           =   8775
            Begin VB.CheckBox chkSortieAmj 
               Alignment       =   1  'Right Justify
               Caption         =   "Date de sortie"
               Height          =   255
               Left            =   2880
               TabIndex        =   121
               Top             =   2160
               Width           =   1335
            End
            Begin VB.Frame fraNature 
               Caption         =   "Salarié"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3255
               Left            =   5880
               TabIndex        =   49
               Top             =   240
               Width           =   2775
               Begin VB.OptionButton optNatureS 
                  Caption         =   "Oui"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   52
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   615
               End
               Begin VB.OptionButton optNatureX 
                  Caption         =   "Non"
                  Height          =   255
                  Left            =   840
                  TabIndex        =   51
                  Top             =   240
                  Width           =   615
               End
               Begin VB.ListBox lstService 
                  Height          =   2595
                  Left            =   120
                  TabIndex        =   50
                  Top             =   480
                  Width           =   2535
               End
            End
            Begin VB.TextBox txtPrénom 
               Height          =   285
               Left            =   1560
               MaxLength       =   32
               TabIndex        =   7
               Top             =   1560
               Width           =   4095
            End
            Begin VB.Frame Frame5 
               Height          =   500
               Left            =   2280
               TabIndex        =   44
               Top             =   240
               Width           =   3375
               Begin VB.OptionButton optCivilitéM 
                  Caption         =   "M."
                  Height          =   255
                  Left            =   120
                  TabIndex        =   48
                  Top             =   160
                  Value           =   -1  'True
                  Width           =   495
               End
               Begin VB.OptionButton optCivilitéMme 
                  Caption         =   "Mme"
                  Height          =   255
                  Left            =   720
                  TabIndex        =   47
                  Top             =   160
                  Width           =   735
               End
               Begin VB.OptionButton optCivilitéMle 
                  Caption         =   "Mlle"
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   46
                  Top             =   160
                  Width           =   735
               End
               Begin VB.OptionButton optCivilitéAutre 
                  Caption         =   "Autre"
                  Height          =   255
                  Left            =   2520
                  TabIndex        =   45
                  Top             =   160
                  Width           =   735
               End
            End
            Begin VB.TextBox txtMatricule 
               Height          =   285
               Left            =   1560
               MaxLength       =   5
               TabIndex        =   5
               Top             =   360
               Width           =   600
            End
            Begin VB.TextBox txtNom 
               Height          =   285
               Left            =   1560
               MaxLength       =   32
               TabIndex        =   6
               Top             =   1080
               Width           =   4095
            End
            Begin VB.TextBox txtEnfantNb 
               Height          =   285
               Left            =   1560
               MaxLength       =   3
               TabIndex        =   10
               Top             =   2760
               Width           =   400
            End
            Begin MSComCtl2.DTPicker txtSortieAmj 
               Height          =   300
               Left            =   4320
               TabIndex        =   9
               Top             =   2160
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
               Format          =   28246019
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtEntréeAmj 
               Height          =   300
               Left            =   1560
               TabIndex        =   8
               Top             =   2160
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
               Format          =   28246019
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblMatricule 
               Caption         =   "Matricule"
               Height          =   375
               Left            =   240
               TabIndex        =   57
               Top             =   360
               Width           =   855
            End
            Begin VB.Label lblPrénom 
               Caption         =   "Prénom"
               Height          =   375
               Left            =   240
               TabIndex        =   56
               Top             =   1680
               Width           =   855
            End
            Begin VB.Label lblEntréeAmj 
               Caption         =   "Date d'entrée"
               Height          =   375
               Left            =   240
               TabIndex        =   55
               Top             =   2280
               Width           =   1215
            End
            Begin VB.Label lblEnfantnb 
               Caption         =   "Nombre d'enfants à charge"
               Height          =   495
               Left            =   240
               TabIndex        =   54
               Top             =   2760
               Width           =   1335
            End
            Begin VB.Label lblNom 
               Caption         =   "Nom"
               Height          =   375
               Left            =   240
               TabIndex        =   53
               Top             =   1080
               Width           =   855
            End
         End
      End
      Begin VB.Frame fraMouvement 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   35
         Top             =   960
         Width           =   8775
         Begin VB.CommandButton cmdMouvementQuit 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   3480
            Width           =   1155
         End
         Begin VB.Frame fraMouvementDétail 
            Height          =   3135
            Left            =   3360
            TabIndex        =   36
            Top             =   240
            Width           =   5295
            Begin VB.TextBox txtNbjOuvré 
               Enabled         =   0   'False
               Height          =   285
               Left            =   4440
               TabIndex        =   65
               Top             =   960
               Width           =   615
            End
            Begin VB.CheckBox chkRepriseAmjK 
               Caption         =   "Après-midi"
               Height          =   255
               Left            =   2400
               TabIndex        =   39
               Top             =   2280
               Width           =   1095
            End
            Begin VB.CheckBox chkDébutAmjK 
               Caption         =   "Après-midi"
               Height          =   255
               Left            =   2400
               TabIndex        =   38
               Top             =   1680
               Width           =   1095
            End
            Begin VB.CheckBox chkNbj 
               Caption         =   "durée de d'absence connue "
               Height          =   255
               Left            =   360
               TabIndex        =   18
               Top             =   360
               Value           =   1  'Checked
               Width           =   3015
            End
            Begin VB.TextBox txtNbj 
               Height          =   285
               Left            =   2400
               TabIndex        =   26
               Top             =   960
               Width           =   615
            End
            Begin MSComCtl2.DTPicker txtDébutAmj 
               Height          =   300
               Left            =   720
               TabIndex        =   24
               Top             =   1680
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   28246019
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtRepriseAmj 
               Height          =   300
               Left            =   720
               TabIndex        =   28
               Top             =   2280
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   28246019
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblNbj 
               Caption         =   "nb jours"
               Height          =   255
               Left            =   360
               TabIndex        =   92
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label lblNbjOuvré 
               Caption         =   "soit jours ouvrés"
               Height          =   255
               Left            =   3120
               TabIndex        =   66
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label libMvtUpd 
               Caption         =   "-"
               Height          =   255
               Left            =   720
               TabIndex        =   64
               Top             =   2760
               Width           =   4335
            End
            Begin VB.Label lblDébutAmj 
               Caption         =   "du"
               Height          =   255
               Left            =   360
               TabIndex        =   37
               Top             =   1680
               Width           =   255
            End
         End
         Begin VB.CommandButton cmdMouvementOK 
            BackColor       =   &H00C0FFC0&
            Caption         =   "&Ok"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   7080
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   3480
            Width           =   1155
         End
         Begin VB.ListBox lstMvt 
            Height          =   3765
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraAbsences 
         Height          =   5415
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   9015
         Begin VB.CheckBox chkOnglet1 
            Caption         =   "affichage automatique de l'onglet   'mouvement'"
            Height          =   255
            Left            =   4800
            TabIndex        =   91
            Top             =   5040
            Value           =   1  'Checked
            Width           =   3975
         End
         Begin MSFlexGridLib.MSFlexGrid fgSalarié 
            Height          =   5010
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   8837
            _Version        =   393216
            Rows            =   1
            Cols            =   4
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   14737632
            ForeColor       =   12582912
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   2
            AllowUserResizing=   3
            FormatString    =   "<Matricule|<Service|< Nom                                                  |"
         End
         Begin MSFlexGridLib.MSFlexGrid fgTotal 
            Height          =   4650
            Left            =   4800
            TabIndex        =   4
            Top             =   240
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   8202
            _Version        =   393216
            Rows            =   1
            Cols            =   4
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   14737632
            ForeColor       =   12582912
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   2
            AllowUserResizing=   3
            FormatString    =   "<Nature                         |>  Droits    |>Absences |>Solde   "
         End
      End
      Begin VB.Frame fraTR 
         Height          =   2895
         Left            =   -68640
         TabIndex        =   29
         Top             =   480
         Width           =   2415
         Begin VB.TextBox txtTRNbj 
            Height          =   285
            Left            =   720
            MaxLength       =   2
            TabIndex        =   22
            Top             =   600
            Width           =   735
         End
         Begin MSComCtl2.DTPicker txtTRdébutAmj 
            Height          =   300
            Left            =   480
            TabIndex        =   20
            Top             =   1440
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   28246019
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin MSComCtl2.DTPicker txtTRFinAmj 
            Height          =   300
            Left            =   480
            TabIndex        =   21
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   28246019
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblTRDébutAmj 
            Caption         =   "Absences du (inclus)"
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblTRFinAmj 
            Caption         =   "au (inclus)"
            Height          =   375
            Left            =   720
            TabIndex        =   32
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label lblTRNbj 
            Caption         =   "Nombre de tickets restaurant"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame fraFérié 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   27
         Top             =   360
         Width           =   8890
         Begin VB.Frame fraExerciceCp 
            Caption         =   "Fin d'exercice Congés Payés"
            Enabled         =   0   'False
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
            TabIndex        =   81
            Top             =   3480
            Width           =   4335
            Begin VB.CommandButton cmdExerciceCp 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Clôture / Ouverture"
               Height          =   645
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   83
               Top             =   960
               Width           =   1515
            End
            Begin VB.TextBox txtExerciceCpNbj 
               Height          =   285
               Left            =   1920
               TabIndex        =   82
               Top             =   1080
               Width           =   495
            End
            Begin MSComCtl2.DTPicker txtExerciceCpAmj 
               Height          =   300
               Left            =   2520
               TabIndex        =   84
               Top             =   360
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   28246019
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblExerciceCpNbj 
               Caption         =   "nb jours congés payés"
               Height          =   375
               Left            =   240
               TabIndex        =   86
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Label lblExerciceCp 
               Caption         =   "1er jour ouvré de l'exercice suivant"
               Height          =   495
               Left            =   240
               TabIndex        =   85
               Top             =   360
               Width           =   2055
            End
         End
         Begin VB.Frame fraExerciceCivil 
            Caption         =   "Fin d'exercice civil"
            Enabled         =   0   'False
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
            TabIndex        =   75
            Top             =   1680
            Width           =   4335
            Begin VB.TextBox txtExerciceCivilNbj 
               Height          =   285
               Left            =   1800
               TabIndex        =   78
               Top             =   1080
               Width           =   495
            End
            Begin VB.CommandButton cmdExerciceCivil 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Clôture / Ouverture"
               Height          =   645
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   76
               Top             =   960
               Width           =   1515
            End
            Begin MSComCtl2.DTPicker txtExerciceCivilAmj 
               Height          =   300
               Left            =   2520
               TabIndex        =   77
               Top             =   360
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   28246019
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblExerciceCivil 
               Caption         =   "1er jour ouvré de l'exercice suivant"
               Height          =   495
               Left            =   240
               TabIndex        =   80
               Top             =   360
               Width           =   2055
            End
            Begin VB.Label lblExerciceCivilNbj 
               Caption         =   "nb jours RTT"
               Height          =   255
               Left            =   240
               TabIndex        =   79
               Top             =   1080
               Width           =   1335
            End
         End
         Begin VB.Frame fraFériéAdd 
            Caption         =   "Jours Fériés"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   4335
            Begin VB.CheckBox chkFériéAM 
               Caption         =   "après-midi"
               Height          =   255
               Left            =   360
               TabIndex        =   74
               Top             =   840
               Width           =   1455
            End
            Begin VB.CommandButton cmdFériéOk 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Ajouter"
               Height          =   645
               Left            =   2640
               Style           =   1  'Graphical
               TabIndex        =   72
               Top             =   360
               Width           =   1515
            End
            Begin MSComCtl2.DTPicker txtFériéAmj 
               Height          =   300
               Left            =   360
               TabIndex        =   73
               Top             =   360
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   28246019
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
         End
         Begin VB.ListBox lstFérié 
            Height          =   5130
            Left            =   4560
            TabIndex        =   19
            Top             =   240
            Width           =   4215
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSalariéMvt 
         Height          =   5085
         Left            =   -74880
         TabIndex        =   41
         Top             =   480
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   8969
         _Version        =   393216
         Rows            =   1
         Cols            =   11
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   14737632
         ForeColor       =   12582912
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         TextStyleFixed  =   4
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"DRH.frx":01AA
      End
      Begin MSFlexGridLib.MSFlexGrid fgTR 
         Height          =   5370
         Left            =   -74880
         TabIndex        =   69
         Top             =   360
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   9472
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   14737632
         ForeColor       =   12582912
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         TextStyleFixed  =   4
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   "<Matricule|<Service|< Nom                                                  |>Absences | T.R.       ||"
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   93
         Top             =   600
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   9551
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Sélection"
         TabPicture(0)   =   "DRH.frx":0281
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraPrintB"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Tri / impression"
         TabPicture(1)   =   "DRH.frx":029D
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraPrintSort"
         Tab(1).Control(1)=   "fraPrintMouvement"
         Tab(1).Control(2)=   "cmdPrintSelect"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "état des absences"
         TabPicture(2)   =   "DRH.frx":02B9
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraPrint"
         Tab(2).ControlCount=   1
         Begin VB.CommandButton cmdPrintSelect 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Imprimer l'état"
            Height          =   735
            Left            =   -67920
            Style           =   1  'Graphical
            TabIndex        =   119
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Frame fraPrintMouvement 
            Caption         =   "Impression des mouvements"
            Height          =   2175
            Left            =   -71280
            TabIndex        =   115
            Top             =   720
            Width           =   2775
            Begin VB.CheckBox chkPrintMouvementDétail 
               Caption         =   "Détail des mouvements"
               Height          =   375
               Left            =   240
               TabIndex        =   118
               Top             =   240
               Value           =   1  'Checked
               Width           =   2055
            End
            Begin VB.OptionButton optPrintMouvement01 
               Caption         =   "tri par nature / date"
               Height          =   255
               Left            =   240
               TabIndex        =   117
               Top             =   960
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton optPrintMouvement02 
               Caption         =   "tri par date"
               Height          =   255
               Left            =   240
               TabIndex        =   116
               Top             =   1440
               Width           =   1815
            End
         End
         Begin VB.Frame fraPrintSort 
            Caption         =   "Tri"
            Height          =   2295
            Left            =   -74760
            TabIndex        =   110
            Top             =   600
            Width           =   3015
            Begin VB.OptionButton optPrintSort01 
               Caption         =   "Service / nom"
               Height          =   255
               Left            =   240
               TabIndex        =   114
               Top             =   360
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton optPrintSort02 
               Caption         =   "Service / matricule"
               Height          =   255
               Left            =   240
               TabIndex        =   113
               Top             =   840
               Width           =   1815
            End
            Begin VB.OptionButton optPrintSort03 
               Caption         =   "Nom"
               Height          =   255
               Left            =   240
               TabIndex        =   112
               Top             =   1320
               Width           =   1815
            End
            Begin VB.OptionButton optPrintSort04 
               Caption         =   "Matricule"
               Height          =   255
               Left            =   240
               TabIndex        =   111
               Top             =   1800
               Width           =   1815
            End
         End
         Begin VB.Frame fraPrintB 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Left            =   120
            TabIndex        =   98
            Top             =   480
            Width           =   8775
            Begin VB.CheckBox chkPrintCP 
               Caption         =   "Exclure les mvts dont la date de début est <à la date de début de période"
               Height          =   615
               Left            =   3840
               TabIndex        =   122
               Top             =   2040
               Width           =   4455
            End
            Begin VB.ComboBox cboPrintService 
               ForeColor       =   &H00FF0000&
               Height          =   315
               Left            =   5640
               Style           =   2  'Dropdown List
               TabIndex        =   120
               Top             =   3480
               Width           =   2895
            End
            Begin VB.ListBox lstPrintMouvement 
               Height          =   3885
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   109
               Top             =   720
               Width           =   3135
            End
            Begin VB.TextBox txtPrintMatricule 
               Height          =   285
               Left            =   5640
               TabIndex        =   108
               Top             =   4200
               Width           =   1215
            End
            Begin VB.CheckBox chkPrintService 
               Caption         =   "Tous les services"
               Height          =   375
               Left            =   3840
               TabIndex        =   107
               Top             =   3480
               Value           =   1  'Checked
               Width           =   1695
            End
            Begin VB.CheckBox chkPrintSalarié 
               Caption         =   "Tous les salariés"
               Height          =   375
               Left            =   3840
               TabIndex        =   106
               Top             =   4200
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.CheckBox chkPrintMouvement 
               Caption         =   "Tous les mouvements"
               Height          =   375
               Left            =   120
               TabIndex        =   105
               Top             =   240
               Value           =   1  'Checked
               Width           =   1935
            End
            Begin VB.CheckBox chkPrintPériode 
               Caption         =   "Période en cours"
               Height          =   375
               Left            =   3840
               TabIndex        =   104
               Top             =   360
               Value           =   1  'Checked
               Width           =   1575
            End
            Begin VB.Frame fraPrintPériode 
               Caption         =   "Période"
               Height          =   1335
               Left            =   5520
               TabIndex        =   99
               Top             =   360
               Width           =   2895
               Begin MSComCtl2.DTPicker txtPrintFinAMJ 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   100
                  Top             =   840
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   16777215
                  CalendarForeColor=   0
                  CalendarTitleBackColor=   8421504
                  CalendarTitleForeColor=   16777215
                  CalendarTrailingForeColor=   12632256
                  CustomFormat    =   "dd  MM yyy"
                  Format          =   28246019
                  CurrentDate     =   36299
                  MaxDate         =   401768
                  MinDate         =   -328351
               End
               Begin MSComCtl2.DTPicker txtPrintDébutAMJ 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   101
                  Top             =   360
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   16777215
                  CalendarForeColor=   0
                  CalendarTitleBackColor=   8421504
                  CalendarTitleForeColor=   16777215
                  CalendarTrailingForeColor=   12632256
                  CustomFormat    =   "dd  MM yyy"
                  Format          =   28246019
                  CurrentDate     =   36299
                  MaxDate         =   401768
                  MinDate         =   -328351
               End
               Begin VB.Label lblPrintDébutAmj 
                  Caption         =   "du (inclus)"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   103
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label lblPrintRepriseAmj 
                  Caption         =   "au (inclus)"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   102
                  Top             =   840
                  Width           =   1095
               End
            End
         End
         Begin VB.Frame fraPrint 
            Height          =   1335
            Left            =   -74640
            TabIndex        =   94
            Top             =   600
            Width           =   5415
            Begin VB.CheckBox chkPrintEnfantMalade 
               Caption         =   "Etat AM : inclure absences pour 'enfant malade'"
               Height          =   255
               Left            =   240
               TabIndex        =   95
               Top             =   840
               Width           =   4215
            End
            Begin MSComCtl2.DTPicker txtPrtAmj 
               Height          =   300
               Left            =   1440
               TabIndex        =   96
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   28246019
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblPrtAmj 
               Caption         =   "Date situation"
               Height          =   255
               Left            =   120
               TabIndex        =   97
               Top             =   240
               Width           =   1335
            End
         End
      End
   End
   Begin VB.Label libSalarié 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1200
      TabIndex        =   40
      Top             =   0
      Width           =   3735
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuDRHCréer2 
         Caption         =   "Créer un salarié"
      End
      Begin VB.Menu mnuContext_X1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintJournal 
         Caption         =   "Etat des mouvements du jour"
      End
      Begin VB.Menu mnuPrintAbsence 
         Caption         =   "Etat des absences du jour"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPrintAM 
         Caption         =   "Etat annuel des arrêts maladie"
      End
      Begin VB.Menu mnuPrintPlanning 
         Caption         =   "Planning des congés"
      End
      Begin VB.Menu mnuPrintRecap 
         Caption         =   "Situation récapitulative"
      End
      Begin VB.Menu mnuPrintSelect 
         Caption         =   "Etats à la demande"
      End
      Begin VB.Menu mnuContext_X2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuXClôtureCp 
         Caption         =   "Fermeture exercice congés payés"
      End
      Begin VB.Menu mnuXOuvertureCp 
         Caption         =   "Ouverture exercice congés payés"
      End
      Begin VB.Menu mnuXClôtureCivil 
         Caption         =   "Clôture exercice civil"
      End
      Begin VB.Menu mnuContext_X3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuDRHMvt 
      Caption         =   "mnuDRHMvt"
      Visible         =   0   'False
      Begin VB.Menu mnuDRHMvtAnnuler 
         Caption         =   "Annuler un mouvement"
      End
      Begin VB.Menu mnuDRHMvtModifier 
         Caption         =   "Modifier un mouvement"
      End
      Begin VB.Menu mnuDRHMvtEffacer 
         Caption         =   "Effacer ce mouvement"
      End
   End
   Begin VB.Menu mnuDRH 
      Caption         =   "mnuDRH"
      Visible         =   0   'False
      Begin VB.Menu mnuDRHDisplay 
         Caption         =   "Sélectionner un salarié"
      End
      Begin VB.Menu mnux1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDRHCréer 
         Caption         =   "Créer un salarié"
      End
      Begin VB.Menu mnuDRHModifier 
         Caption         =   "Modifier un salarié"
      End
      Begin VB.Menu mnuDRHEffacer 
         Caption         =   "Effacer un salarié"
      End
   End
   Begin VB.Menu mnuFérié 
      Caption         =   "mnuFérié"
      Visible         =   0   'False
      Begin VB.Menu mnuFérié_Suppress 
         Caption         =   "Supprimer ce jour férié"
      End
   End
End
Attribute VB_Name = "frmDRH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer

Dim DrhAut As typeAuthorization

Dim xTable As typeElpTable
Dim xDRH As typeDRH, mDRH As typeDRH
Dim xDRHMvt As typeDRHMvt, mDRHMvt As typeDRHMvt, wDRHMvt As typeDRHMvt, oldDRHMvt As typeDRHMvt
Dim mCP As typeDRHMvt, mCivil As typeDRHMvt, mTR As typeDRHMvt

Dim fgSalariéMvt_FormatString As String, fgSalariéMvt_K As Integer, fgSalariéMvt_TopRow As Long
Dim fgSalariéMvt_RowDisplay As Integer, fgSalariéMvt_RowClick As Integer
Dim fgSalariéMvt_ColorClick As Long, fgSalariéMvt_ColorDisplay As Long
Dim fgSalariéMvt_Sort1 As Integer, fgSalariéMvt_Sort2 As Integer

Dim fgSalarié_FormatString As String, fgSalarié_K As Integer
Dim fgSalarié_RowDisplay As Integer, fgSalarié_RowClick As Integer
Dim fgSalarié_ColorClick As Long, fgSalarié_ColorDisplay As Long
Dim fgSalarié_Sort1 As Integer, fgSalarié_Sort2 As Integer

Dim fgTR_FormatString As String, fgTR_K As Integer
Dim fgTR_RowDisplay As Integer, fgTR_RowClick As Integer
Dim fgTR_ColorClick As Long, fgTR_ColorDisplay As Long
Dim fgTR_Sort1 As Integer, fgTR_Sort2 As Integer

Dim fgtotal_FormatString As String, fgtotal_K As Integer, arrTotal_Row(99) As Integer
Dim wAmj As String * 8, wAmj1 As String * 8, wAmj2 As String * 8
Dim wAmjK As String * 1, wAmjK1 As String * 1, wAmjK2 As String * 1

Dim cmdImport_Select_Nb  As Integer, cmdImport_Nb  As Integer

Dim prtEnTête As String, prtDestinataire As String
Dim prtDocument As String * 2, prtSort As String * 1, prtDébutAmj As String * 8, prtFinAmj As String * 8
Dim prtSelectK As String * 1, prtSelect As String, prtSelectMvtK As String * 1, prtSelectMvt As String
Dim prtAmj As String * 8, prtFinAmj1 As String * 8

Dim xElpBuffer As typeElpBuffer

Dim cumNbjC As Double, maxNbjC As Double
Dim oldDRHMvt_Index As Integer
Dim wNbj As Double

Dim mprtDébutAMJ As String * 8, mprtFinAMJ As String * 8
Dim arrPrintMouvement(100) As String * 4, arrPrintMouvement_Nb As Integer

Private Sub cboPrintService_Click()
cmdPrint_Control
End Sub


Public Sub fgSalariéMvt_Load(xMethod As String)
Dim X As String, mMethod As String
arrDRHMvt_NBMax = 0
arrDRHMvt_Suite = True: arrDRHMvt_NB = 0
ReDim arrDRHMvt(1)
recDRHMvt_Init xDRHMvt

xDRHMvt.Method = xMethod
xDRHMvt.Matricule = xDRH.Matricule
xDRHMvt.IdSeq = 0
xDRHMvt.DébutAmj = "00000000"
arrDRHMvt(0) = xDRHMvt
arrDRHMvt(0).IdSeq = "99999"
arrDRHMvt(0).DébutAmj = "99999999"

mMethod = Trim(xMethod) & "+"
Do Until Not arrDRHMvt_Suite
    srvDRHMvt_Monitor xDRHMvt
    xDRHMvt = arrDRHMvt(arrDRHMvt_NB)
    xDRHMvt.Method = mMethod
Loop
fgSalariéMvt_Display
If fgSalariéMvt.Rows > 1 Then fgSalariéMvt.TopRow = fgSalariéMvt.Rows - 1: fgSalariéMvt.LeftCol = 0
fraMouvement_Reset
fgSalariéMvt_RowDisplay = 0: fgSalariéMvt_RowClick = 0

If mMethod = "SnapP0+" Then
    cmdMouvementHisto.Caption = constEnCours
Else
    cmdMouvementHisto.Caption = constHistorique
End If

cmdMouvementHisto.Visible = True

End Sub
Public Sub fgSalariéMvt_DisplayLine()
fgSalariéMvt_K = (fgSalariéMvt.Row) * fgSalariéMvt.Cols
X = dateImp(xDRHMvt.DébutAmj)
If xDRHMvt.DébutAmjK = "1" Then X = X & " midi"
fgSalariéMvt.TextArray(0 + fgSalariéMvt_K) = X
X = dateImp(xDRHMvt.RepriseAmj)
If xDRHMvt.RepriseChk = "1" Then X = X & " *"
If xDRHMvt.RepriseAmjK = "1" Then X = X & " midi"
fgSalariéMvt.TextArray(1 + fgSalariéMvt_K) = X
xTable.Method = "Seek="
xTable.Id = "DRH"
xTable.K1 = "Mvt"
xTable.K2 = xDRHMvt.MvtCode
intReturn = tableElpTable_Read(xTable)
If intReturn <> 0 Then xTable.Name = xDRHMvt.MvtCode

fgSalariéMvt.TextArray(2 + fgSalariéMvt_K) = xDRHMvt.NbjOuvré '& xDRHMvt.MvtSens
If xDRHMvt.MvtCO = "C" Then fgSalariéMvt.TextArray(3 + fgSalariéMvt_K) = xDRHMvt.Nbj '& xDRHMvt.MvtSens
If xDRHMvt.MvtSens = "-" Or xDRHMvt.MvtSens = "D" Or xDRHMvt.MvtSens = "P" Then
    fgSalariéMvt.Col = 2: fgSalariéMvt.CellForeColor = errUsr.ForeColor
    fgSalariéMvt.Col = 3: fgSalariéMvt.CellForeColor = errUsr.ForeColor
End If
fgSalariéMvt.TextArray(4 + fgSalariéMvt_K) = xTable.Name
fgSalariéMvt.TextArray(5 + fgSalariéMvt_K) = xDRHMvt.Statut
fgSalariéMvt.TextArray(6 + fgSalariéMvt_K) = xDRHMvt.RéfInterne
fgSalariéMvt.TextArray(7 + fgSalariéMvt_K) = dateImp(xDRHMvt.UpdAmj) & " " & timeImp(xDRHMvt.UpdHms)
fgSalariéMvt.TextArray(8 + fgSalariéMvt_K) = arrDRHMvt_Index
fgSalariéMvt.TextArray(9 + fgSalariéMvt_K) = xDRHMvt.DébutAmj
fgSalariéMvt.TextArray(10 + fgSalariéMvt_K) = xDRHMvt.RepriseAmj
End Sub
Private Sub fgSalariéMvt_Display()
Dim X2 As String

fraMouvement.Visible = False
cmdMouvementHisto.Visible = True

arrTotal_Init

For I = 1 To fgTotal.Rows - 1
    fgTotal.Row = I
    fgTotal.Col = 1: fgTotal.Text = ""
    fgTotal.Col = 2: fgTotal.Text = ""
Next I

fgSalariéMvt_TopRow = fgSalariéMvt.TopRow
fgSalariéMvt.Clear: fgSalariéMvt.Rows = 1

fgSalariéMvt.Rows = 1
fgSalariéMvt.FormatString = fgSalariéMvt_FormatString
fgSalariéMvt.Enabled = True
For arrDRHMvt_Index = 1 To arrDRHMvt_NB
    If arrDRHMvt(arrDRHMvt_Index).Method <> constDelete Then
        fgSalariéMvt.Rows = fgSalariéMvt.Rows + 1
        fgSalariéMvt.Row = fgSalariéMvt.Rows - 1
        xDRHMvt = arrDRHMvt(arrDRHMvt_Index)
        'I = mId$(xDRHMvt.MvtCode, 1, 2)
        If xDRHMvt.Statut = " " Then arrTotal_Add xDRHMvt
        fgSalariéMvt_DisplayLine
    End If
Next arrDRHMvt_Index

fgTotal_Display

If fgSalariéMvt.Rows > 2 Then
    fgSalariéMvt_Sort1 = 9: fgSalariéMvt_Sort2 = 9: fgSalariéMvt_Sort
    If fgSalariéMvt_TopRow < fgSalariéMvt.Rows Then fgSalariéMvt.TopRow = fgSalariéMvt_TopRow
End If

End Sub

Public Sub fgSalariéMvt_Sort()
If fgSalariéMvt.Rows > 1 Then
    fgSalariéMvt.Row = 1
    fgSalariéMvt.RowSel = fgSalariéMvt.Rows - 1
    
    fgSalariéMvt.Col = fgSalariéMvt_Sort1
    fgSalariéMvt.ColSel = fgSalariéMvt_Sort2
    fgSalariéMvt.Sort = 1
End If
End Sub

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub

Public Sub cmdContext_Quit()
blnControl = False
lstErr.Clear

If fraMouvement.Visible Then
    If lstMvt.Enabled Then
        fraMouvement.Visible = False
        cmdMouvementSaisir.Enabled = True
    Else
        fraMouvement_Reset
    End If
    Exit Sub
End If

If currentAction <> "" Then
    Select Case currentAction
        Case "Mvt_AddNew", "Mvt_Update"
            If lstMvt.Enabled Then
                fraMouvement.Visible = False
            Else
                fraMouvement_Reset
            End If
        Case Else
            cmdContext.Caption = constcmdRechercher
            cmdReset
            SSTab1.Tab = 0
    End Select
    currentAction = ""
Else
    If SSTab1.Tab <> 0 Then
        cmdContext.Caption = constcmdRechercher
        cmdReset
        SSTab1.Tab = 0
    Else
        If blnMsgBox_Quit Then
            X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
        Else
           X = vbYes
        End If
        If X = vbYes Then Unload Me
    End If
End If

End Sub

Public Sub cmdContext_Return()

SendKeys "{TAB}"

End Sub

Public Sub Form_Init()
cmdReset
paramDRHMvt_Sys0
paramDRHMvt_CP
paramDRHMvt_Civil

txtNbjOuvré.Enabled = False

libSalarié.ForeColor = vbBlue 'warnUsrColor
SSTab1.Tab = 0
fgSalariéMvt_FormatString = fgSalariéMvt.FormatString
fgSalarié_Sort1 = 0: fgSalarié_Sort2 = 0
fgSalarié_FormatString = fgSalarié.FormatString
fgSalarié_Sort1 = 2: fgSalarié_Sort2 = 2
fgtotal_FormatString = fgTotal.FormatString
fgTR_FormatString = fgTR.FormatString
fgTR_Sort1 = 2: fgTR_Sort2 = 2

constDRHFérié = "Férié"
constDRHService = "Service"
constDRHMvt = "Mvt"
fgSalarié_Load
lstService_Load
lstMvt_Load
fgTotal_Load
DTPicker_Now txtFériéAmj
recElpTable_Init mDRHCalendrier
V = Param_DRHCalendrier(mId$(DSys, 1, 6), mDRHCalendrier)
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, V)

mprtFinAMJ = dateElp("Jour", -2, mCivil.RepriseAmj)
Call DTPicker_Set(txtPrintFinAMJ, mprtFinAMJ)
mprtDébutAMJ = mCivil.DébutAmj
Call DTPicker_Set(txtPrintDébutAMJ, mprtDébutAMJ)

prtAmj = DSys
Call DTPicker_Set(txtPrtAmj, prtAmj)

prtDRHMvt_Nbj_Init DSys
fraTR_Init

recElpTable_Init recElpTable
recElpTable.Id = "DRH"
recElpTable.K1 = "Service"
Call cbo_Load(recElpTable, cboPrintService, 4)

'jpl 2000.12.26 Call MsgBox("pas de contrôle de dates !", vbInformation, "DRH")

End Sub

Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
         If TypeOf xobj Is CommandButton Or TypeOf xobj Is ListBox Then
           xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            xobj.ForeColor = MouseMoveActiveControl.ForeColor
        End If
        Exit For
    End If
Next xobj

End Sub



Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Or TypeOf C Is ListBox Then
            
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            MouseMoveActiveControl.ForeColor = C.ForeColor
            C.ForeColor = MouseMoveUsr.ForeColor
        End If
    End If
End If

End Sub


Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------
End Sub


Public Sub Msg_Snd(ByVal X As String)
End Sub

'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
currentActiveControl_Name = C.Name
    C.ForeColor = txtUsr.ForeColor
    C.BackColor = focusUsr.BackColor
End Sub


'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
arrTag(Val(C.Tag)) = True
    C.ForeColor = txtUsr.ForeColor
    C.BackColor = txtUsr.BackColor
End Sub


Private Sub chkDébutAmjK_Click()
If blnControl Then cmdMouvement_Control

End Sub

Private Sub chkFériéAM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkFériéAM
End Sub


Private Sub chkNbj_Click()
If blnControl Then cmdMouvement_Control

End Sub

Private Sub chknbj_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkNbj
End Sub


Private Sub chkPrintMouvement_Click()
cmdPrint_Control
End Sub

Private Sub chkPrintPériode_Click()
cmdPrint_Control

End Sub


Private Sub chkPrintSalarié_Click()
cmdPrint_Control

End Sub


Private Sub chkPrintService_Click()
cmdPrint_Control

End Sub


Private Sub chkRepriseAmjK_Click()
If blnControl Then cmdMouvement_Control

End Sub

Private Sub chkRepriseAmjK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkRepriseAmjK
End Sub


Private Sub chkSortieAmj_Click()
If blnControl Then cmdSalarié_Control

End Sub

Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: cmdContext_mnu
    Case Is = constcmdAbandonner: cmdReset
End Select

End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext
End Sub


Private Sub cmdPrintSelect_Click()
Call lstErr_Clear(lstErr, cmdContext, "PrintSelect: début")
arrPrintMouvement_Nb = 0
If chkPrintMouvement <> "1" Then cmdPrintSelect_Import_Mouvement_Init


Call DTPicker_Control(txtPrintDébutAMJ, mprtDébutAMJ)
Call DTPicker_Control(txtPrintFinAMJ, mprtFinAMJ)

prtAmj = mprtFinAMJ

If optPrintSort01 Then prtSort = "1"
If optPrintSort02 Then prtSort = "2"
If optPrintSort03 Then prtSort = "3"
If optPrintSort04 Then prtSort = "4"
If chkPrintMouvementDétail = "1" Then
    prtSelectMvtK = "1"
Else
    prtSelectMvtK = " "
End If
prtDocument = "10"
prtDébutAmj = mprtDébutAMJ
prtFinAmj = mprtFinAMJ
If chkPrintPériode <> "1" Then
    prtSelectK = "1"
Else
    prtSelectK = " "
End If
prtSelect = " "
 prtSelectMvt = " "
prtEnTête = "DRH " 'Congés payés, début de l 'absence entre le " & dateImp(prtDébutAmj) & " et le  " & dateImp(prtFinAmj)

prtFinAmj1 = dateElp("Jour", 1, mTR.RepriseAmj)

Call lstErr_AddItem(lstErr, cmdContext, "PrintSelect: importation des mouvements")
cmdPrintSelect_Import

'For arrDRH_Index = 1 To arrDRH_NB
'    arrDRHNbjOuvrés(arrDRH_Index, 0) = Fix(arrDRHNbjOuvrés(arrDRH_Index, 0))
'    arrDRHTR(arrDRH_Index) = mTR.NbjOuvré - arrDRHNbjOuvrés(arrDRH_Index, 0)
'Next arrDRH_Index
Call lstErr_AddItem(lstErr, cmdContext, "PrintSelect : impression ")

If cmdImport_Select_Nb = 0 Then
    Call lstErr_AddItem(lstErr, cmdPrint, "! => Aucun salarié sélectionné !")
    'GoTo cmdPrint_End
Else
    ReDim arrDRHMvt(300)
    
    cmdPrint_Monitor
End If
Me.MousePointer = 0
Me.Enabled = True
Call lstErr_AddItem(lstErr, cmdContext, "PrintSelect: fin")
ReDim arrDRHMvt(10)

End Sub

Private Sub cmdExerciceCivil_Click()

xDRHMvt = mCivil
xDRHMvt.Method = "ExerCivil"
xDRHMvt.Nbj = Val(Trim(txtExerciceCivilNbj))
xDRHMvt.NbjOuvré = xDRHMvt.Nbj
xDRHMvt.DébutAmj = mCivil.RepriseAmj
Call DTPicker_Control(txtExerciceCivilAmj, xDRHMvt.RepriseAmj)
xDRHMvt.DébutAmjK = "0"
xDRHMvt.RepriseAmjK = "0"
xDRHMvt.NbjChk = "1"

frmDRH.Enabled = False
blnControl = False
xDRHMvt.UpdAmj = DSys
xDRHMvt.UpdHms = time_Hms

V = srvDRHMvt_Update(xDRHMvt)

If IsNull(V) Then
    mCivil = xDRHMvt
    Call lstErr_Clear(lstErr, cmdContext, "Mise à jour effectuée: " & xDRH.Matricule & "_" & Trim(xDRHMvt.IdSeq))
    'cmdReset
Else
    Call lstErr_Clear(lstErr, cmdContext, V)
End If
frmDRH.Enabled = True
AppActivate frmDRH.Caption

End Sub

Private Sub cmdExerciceCp_Click()
xDRHMvt = mCP
xDRHMvt.Method = "ExerCP"
xDRHMvt.Nbj = Val(Trim(txtExerciceCpNbj))
xDRHMvt.NbjOuvré = xDRHMvt.Nbj
xDRHMvt.DébutAmj = mCP.RepriseAmj
Call DTPicker_Control(txtExerciceCpAmj, xDRHMvt.RepriseAmj)
xDRHMvt.DébutAmjK = "0"
xDRHMvt.RepriseAmjK = "0"
xDRHMvt.NbjChk = "1"

frmDRH.Enabled = False
blnControl = False
xDRHMvt.UpdAmj = DSys
xDRHMvt.UpdHms = time_Hms

V = srvDRHMvt_Update(xDRHMvt)

If IsNull(V) Then
    mCP = xDRHMvt
    Call lstErr_Clear(lstErr, cmdContext, "Mise à jour effectuée: " & xDRH.Matricule & "_" & Trim(xDRHMvt.IdSeq))
    'cmdReset
Else
    Call lstErr_Clear(lstErr, cmdContext, V)
End If
frmDRH.Enabled = True
AppActivate frmDRH.Caption

End Sub


Private Sub cmdFériéOk_Click()
V = DTPicker_Control(txtFériéAmj, wAmj)
If Not IsNull(V) Then Call lstErr_Clear(lstErr, txtFériéAmj, V): Exit Sub
recElpTable_Init xTable
xTable.Method = "AddNew"
xTable.Id = "DRH"
xTable.K1 = constDRHFérié
xTable.K2 = wAmj
V = dateImp(wAmj)
xTable.Name = Format(V, "dddd d mmmm yyyy")
If chkFériéAM = "1" Then
    xTable.Memo = "0X"
Else
    xTable.Memo = "XX"
End If

intReturn = tableElpTable_Update(xTable)
If intReturn <> 0 Then Call lstErr_Clear(lstErr, txtFériéAmj, V & " déjà enregistré (err : " & intReturn & " ) "):   Exit Sub

lstFérié_Load

Call Param_DRHCalendrier_Update(mId$(wAmj, 1, 4), mDRHCalendrier)

End Sub

Private Sub cmdFériéOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdFériéOk
End Sub


Private Sub cmdMouvementHisto_Click()
If cmdMouvementHisto.Caption = constHistorique Then
    fgSalariéMvt_Load "SnapP0"
Else
    fgSalariéMvt_Load "SnapL0" ' 2000-08-10 Pb IdSeq sur lignes annulées  "SnapL0"
End If

End Sub

Private Sub cmdMouvementHisto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdMouvementHisto
End Sub


Private Sub cmdMouvementOK_Click()
cmdMouvement_Control
If lstErr.ListCount <> 0 Then Exit Sub
cmdMouvement_Update

End Sub

Private Sub cmdMouvementOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdMouvementOK
End Sub


Private Sub cmdMouvementQuit_Click()
cmdContext_Quit 'fraMouvement_Reset
End Sub

Private Sub cmdMouvementSaisir_Click()
recDRHMvt_Init mDRHMvt
fraMouvement_Reset
fraMouvement.Visible = True
cmdMouvementSaisir.Enabled = False
End Sub

Private Sub cmdPrint_Click()
'MouseMoveActiveControl_Set cmdPrint
'prtDRH_Monitor " "

End Sub


Private Sub cmdPrintB_Click()

End Sub

Private Sub cmdSalariéOK_Click()
cmdSalarié_Control
If lstErr.ListCount <> 0 Then Exit Sub

frmDRH.Enabled = False

blnControl = False
xDRH.UpdAmj = DSys
xDRH.UpdHms = time_Hms

V = srvDRH_Update(xDRH)

If IsNull(V) Then
    Select Case xDRH.Method
        Case constUpdate: arrDRH(arrDRH_Index) = xDRH
        Case constDelete: arrDRH(arrDRH_Index) = xDRH
        Case constAddNew: Call arrDRH_AddItem(xDRH)
    End Select
    fgSalarié_Display
    Call lstErr_Clear(lstErr, cmdContext, "Mise à jour effectuée - Matricule : " & xDRH.Matricule)
    cmdContext_Quit
Else
    Call lstErr_Clear(lstErr, cmdContext, V)
End If
frmDRH.Enabled = True
AppActivate frmDRH.Caption

End Sub

Private Sub cmdSalariéOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdSalariéOK
End Sub


Private Sub cmdSalatiéMvtPrint_Click()
prtSort = "1"
prtDocument = "05"
prtDébutAmj = prtAmj
prtFinAmj = "99999999"
prtSelectK = " ": prtSelect = " "
prtSelectMvtK = " ": prtSelectMvt = " "
prtEnTête = "DRH : Situation salarié au " & dateImp(prtAmj)
cmdImport_Select_Nb = 1
cmdPrint_Monitor

End Sub

Private Sub cmdTRControl_Click()

If cmdTRControl.Caption = constAnnuler Then
    fgTR.Clear: fgTR.Rows = 1
    Kill paramTR_Filename
    fraTR.Enabled = True
    cmdTRControl.Caption = "Etat préparatoire"
    Exit Sub
End If

mTR.Matricule = "$TR"
mTR.IdSeq = CInt(mId$(DSys, 3, 4))
Call DTPicker_Control(txtTRdébutAmj, mTR.DébutAmj)
Call DTPicker_Control(txtTRFinAmj, mTR.RepriseAmj)
mTR.NbjOuvré = CDbl(Val(txtTRNbj)): mTR.Nbj = mTR.NbjOuvré
mTR.Statut = "$"
mTR.UpdAmj = DSys
mTR.UpdHms = time_Hms
Call lstErr_Clear(lstErr, cmdContext, "TR : màj $TR " & mTR.Matricule)
    
If Not IsNull(srvDRHMvt_Update(mTR)) Then
    Call MsgBox(mTR.Method & "erreur  [ $TR  ] " & mTR.IdSeq, vbCritical, "DRH : cmdTRControl")
    Exit Sub
End If

prtDRHMvt.mTR_Nbj = mTR.Nbj
prtAmj = mTR.DébutAmj
prtSort = "1"
prtDocument = "06"
prtDébutAmj = mTR.DébutAmj
prtFinAmj = mTR.RepriseAmj
prtSelectK = " ": prtSelect = " "
prtSelectMvtK = " ": prtSelectMvt = " "
prtEnTête = "DRH : Ticket restaurant, absences du " & dateImp(prtDébutAmj) & " au " & dateImp(prtFinAmj)

prtFinAmj1 = dateElp("Jour", 1, mTR.RepriseAmj)

prtDRHMvt_Nbj_Init mTR.DébutAmj
prtDRHTR_Init
Call lstErr_AddItem(lstErr, cmdContext, "TR : importation des mouvements")
cmdMouvement_Import

For arrDRH_Index = 1 To arrDRH_NB
    arrDRHNbjOuvrés(arrDRH_Index, 0) = Fix(arrDRHNbjOuvrés(arrDRH_Index, 0))
    arrDRHTR(arrDRH_Index) = mTR.NbjOuvré - arrDRHNbjOuvrés(arrDRH_Index, 0)
    If Trim(arrDRH(arrDRH_Index).SortieAmj) <> "" Then
        arrDRHTR(arrDRH_Index) = 0
    End If
Next arrDRH_Index
Call lstErr_AddItem(lstErr, cmdContext, "TR : impression de l'état des absences")

cmdPrint_Monitor
Me.Enabled = True
Call lstErr_AddItem(lstErr, cmdContext, "TR : affichage")
fgTR_Display
cmdTRDisquette.Enabled = DrhAut.Valider
cmdTrOk.Enabled = DrhAut.Valider

End Sub

Private Sub cmdTRControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdTRControl
End Sub


Private Sub cmdTRDisquette_Click()

On Error GoTo ErrMsg
Call lstErr_Clear(lstErr, cmdContext, "TR : tranfert disquette début : " & paramTR_Filename)
DoEvents
If msFileSystem.FileExists(paramTR_Disquette) Then msFileSystem.DeleteFile paramTR_Disquette, True
msFileSystem.CopyFile paramTR_Filename, paramTR_Disquette
Call lstErr_AddItem(lstErr, cmdContext, "TR : tranfert disquette fin : " & paramTR_Disquette)
Call lstErr_AddItem(lstErr, cmdContext, "TR : impression disquette fin ")
prtDRHTR_Monitor paramTR_Disquette
Exit Sub

ErrMsg:
Call MsgBox("Erreur : " & Error(Err), vbCritical, "Copie de " & paramTR_Filename & " vers " & paramTR_Disquette)
Call lstErr_AddItem(lstErr, cmdContext, "TR : tranfert disquette erreur : " & Err)

End Sub

Private Sub cmdTrOk_Click()
Dim X80 As String * 80
Call lstErr_Clear(lstErr, cmdContext, "TR : Validation début : " & paramTR_Filename)
X = Dir(paramTR_Filename)

Open paramTR_Filename For Output As #1
    
I = 0
For arrDRH_Index = 1 To arrDRH_NB
    wNbj = arrDRHTR(arrDRH_Index)  'mTR.NbjOuvré - arrDRHNbjOuvrés(arrDRH_Index, 0)
    If wNbj > 0 Then
        xDRH = arrDRH(arrDRH_Index)
        I = I + 1
        X80 = Space$(80)
        X80 = paramTR_Id
        Mid$(X80, 19, 4) = mId$(xDRH.Matricule, 2, 4)
        Mid$(X80, 30, 6) = Format$(I, "000000")
        Mid$(X80, 36, 17) = mId$(xDRH.Nom, 1, 17)
        Mid$(X80, 54, 2) = mId$(xDRH.Prénom, 1, 2)
        Mid$(X80, 56, 2) = Format$(wNbj, "00")
        Mid$(X80, 59, 4) = paramTR_Nominal
        Mid$(X80, 63, 4) = paramTR_PartPatronale
        Mid$(X80, 67, 2) = "00"
       
        Print #1, Trim(X80)
    End If
Next arrDRH_Index
    
Close #1
Call lstErr_AddItem(lstErr, cmdContext, "TR : Validation fin nb enregistrements : " & I)

End Sub

Private Sub cmdTrOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdTrOk
End Sub


Private Sub fgSalarié_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y <= fgSalarié.RowHeightMin Then
    Select Case fgSalarié.Col
        Case 0: fgSalarié_Sort1 = 0: fgSalarié_Sort2 = 0: fgSalarié_Sort
        Case 1: fgSalarié_Sort1 = 1: fgSalarié_Sort2 = 2: fgSalarié_Sort
        Case 2: fgSalarié_Sort1 = 2: fgSalarié_Sort2 = 2: fgSalarié_Sort
    End Select
Else
    mnuDRHCréer = DrhAut.Saisir
    mnuDRHModifier = False
    mnuDRHEffacer = False
    mnuDRHDisplay = False
    fgSalarié_K = fgSalarié.Row * fgSalarié.Cols
    If fgSalarié.Rows > 1 Then
        arrDRH_Index = Val(fgSalarié.TextArray(3 + fgSalarié_K))
        mDRH = arrDRH(arrDRH_Index)
        Call fgSalarié_Color(fgSalarié_RowClick, MouseMoveUsr.BackColor, fgSalarié_ColorClick)
        
            mnuDRHDisplay = DrhAut.Consulter
            mnuDRHModifier = DrhAut.Valider
            mnuDRHEffacer = DrhAut.Xspécial
          
    
    End If
     If Button = vbLeftButton Then
        mnuDRHDisplay_Click
    Else
        Me.PopupMenu mnuDRH, vbPopupMenuLeftButton
    End If
End If
End Sub

Private Sub fgSalariéMvt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y <= fgSalariéMvt.RowHeightMin Then
    Select Case fgSalariéMvt.Col
        Case 0: fgSalariéMvt_Sort1 = 9: fgSalariéMvt_Sort2 = 9: fgSalariéMvt_Sort
        Case 1: fgSalariéMvt_Sort1 = 10: fgSalariéMvt_Sort2 = 10: fgSalariéMvt_Sort
        Case 2: fgSalariéMvt_Sort1 = 2: fgSalariéMvt_Sort2 = 2: fgSalariéMvt_Sort
        Case 3: fgSalariéMvt_Sort1 = 3: fgSalariéMvt_Sort2 = 3: fgSalariéMvt_Sort
        Case 4: fgSalariéMvt_Sort1 = 4: fgSalariéMvt_Sort2 = 4: fgSalariéMvt_Sort
        Case 5: fgSalariéMvt_Sort1 = 5: fgSalariéMvt_Sort2 = 5: fgSalariéMvt_Sort
        Case 6: fgSalariéMvt_Sort1 = 6: fgSalariéMvt_Sort2 = 6: fgSalariéMvt_Sort
        Case 7: fgSalariéMvt_Sort1 = 6: fgSalariéMvt_Sort2 = 6: fgSalariéMvt_Sort
    End Select
Else
    mnuDRHMvtAnnuler = False
    mnuDRHMvtModifier = False
    mnuDRHMvtEffacer = False
    fgSalariéMvt_K = fgSalariéMvt.Row * fgSalariéMvt.Cols
    If fgSalariéMvt.Rows > 1 Then
        arrDRHMvt_Index = Val(fgSalariéMvt.TextArray(8 + fgSalariéMvt_K))
        mDRHMvt = arrDRHMvt(arrDRHMvt_Index)
        xDRHMvt = mDRHMvt
        oldDRHMvt_Index = arrDRHMvt_Index
        'Call fgSalariéMvt_Color(fgSalariéMvt_RowClick, MouseMoveUsr.BackColor, fgSalariéMvt_ColorClick)
            fraMouvement_Display
            If mDRHMvt.Statut = " " Then
                mnuDRHMvtAnnuler = DrhAut.Valider
                mnuDRHMvtModifier = DrhAut.Valider
                mnuDRHMvtEffacer = DrhAut.Xspécial
               If mDRHMvt.UpdAmj = DSys And DrhAut.Valider Then mnuDRHMvtEffacer = True
            End If
            
    '§§§§§§§§§§§§§§§§§§ spécial M DAOUD 27.02.2002
        mnuDRHMvtAnnuler = DrhAut.Valider
        mnuDRHMvtModifier = DrhAut.Valider
        mnuDRHMvtEffacer = DrhAut.Valider
    '§§§§§§§§§§§§§§§§§§ spécial M DAOUD 27.02.2002
    End If
    Me.PopupMenu mnuDRHMvt, vbPopupMenuLeftButton
End If

End Sub


Private Sub fgTR_KeyPress(KeyAscii As Integer)
Dim K As Integer, X As String

X = Trim(fgTR.Clip)
If KeyAscii = 8 Then
    K = Len(X) - 1
    If K >= 0 Then fgTR.Text = mId$(X, 1, K)
Else
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        X = X & Chr$(KeyAscii): K = CInt(Val(X))
        If K < 100 Then fgTR.Text = X:   arrDRHTR(arrDRH_Index) = K
    End If
End If

End Sub

Private Sub fgTR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y <= fgTR.RowHeightMin Then
    Select Case fgTR.Col
        Case 0: fgTR_Sort1 = 0: fgTR_Sort2 = 0: fgTR_Sort
        Case 1: fgTR_Sort1 = 1: fgTR_Sort2 = 2: fgTR_Sort
        Case 2: fgTR_Sort1 = 2: fgTR_Sort2 = 2: fgTR_Sort
        Case 3: fgTR_Sort1 = 3: fgTR_Sort2 = 3: fgTR_Sort
        Case 4: fgTR_Sort1 = 4: fgTR_Sort2 = 4: fgTR_Sort
    End Select
Else
    fgTR_K = fgTR.Row * fgTR.Cols
    If fgTR.Rows > 1 Then
        arrDRH_Index = Val(fgTR.TextArray(6 + fgTR_K))
        fgTR.Col = 4: fgTR.CellBackColor = vbCyan
        'fgTR.Text = ""
        'Call fgTR_Color(fgTR_RowClick, MouseMoveUsr.BackColor, fgTR_ColorClick)
        'arrDRH_Index = Val(fgTR.TextArray(3 + fgTR_K))
        'mDRH = arrDRH(arrDRH_Index)
    End If
End If

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


'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
Call BiaPgmAut_Init("DRH", DrhAut)


Form_Init

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Public Sub cmdReset()
cmdContext.Caption = constcmdRechercher
SSTab1.Tab = 0
lstErr.Clear: lstErr.Height = 0
currentAction = ""
fraSalarié.Enabled = False
cmdSalariéOK.Visible = False
cmdTRControl.Enabled = DrhAut.Valider 'False

fraMouvement.Enabled = False
cmdMouvementOK.Visible = False
cmdMouvementHisto.Visible = False
fraMouvement_Reset

cmdFériéOk.Visible = DrhAut.Xspécial
libSalarié = ""
fgSalariéMvt.Clear: fgSalariéMvt_TopRow = 0
fgSalarié_RowDisplay = 0: fgSalarié_RowClick = 0
fgSalariéMvt_RowDisplay = 0: fgSalariéMvt_RowClick = 0

cmdPrint_Control

End Sub

Private Sub fraAbsences_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraFérié_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraMouvement_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraMouvementDétail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraNature_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraSAlarié_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraSalariéDivers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraTR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub lstFérié_Click()
If lstFérié.ListIndex < lstFérié.ListCount Then
    Me.PopupMenu mnuFérié, vbPopupMenuLeftButton
End If
End Sub


Private Sub lstMvt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstMvt.ListIndex > -1 Then lstMvt_Select

End Sub

Private Sub mnuContextAbandonner_Click()
cmdReset
End Sub

Private Sub mnuContextQuitter_Click()
cmdContext_Quit
End Sub


Private Sub mnuDRHCréer_Click()
If DrhAut.Saisir Then
   recDRH_Init mDRH
    xDRH = mDRH
    fraSalarié_display
    mDRH.Method = constAddNew
    SSTab1.Tab = 2
    currentAction = constSaisie
    DTPicker_Now txtEntréeAmj
    txtMatricule.Enabled = True
    txtSortieAmj.Visible = False
    fraSalarié.Enabled = True
    fraSalariéId.Enabled = True
    fraSalariéDivers.Enabled = True
    cmdSalariéOK.Visible = False
    cmdSalariéOK.Caption = "Créer"
    txtMatricule.SetFocus
    cmdContext.Caption = constcmdAbandonner
    blnControl = True
End If

End Sub

Private Sub fgSalarié_Display()
SSTab1.Tab = 0

fgSalarié.Visible = True
fgSalarié.Clear: fgSalarié.Rows = 1

fgSalarié.Rows = 1
fgSalarié.FormatString = fgSalarié_FormatString
fgSalarié.Enabled = True
For arrDRH_Index = 1 To arrDRH_NB
    If arrDRH(arrDRH_Index).Method <> constDelete Then
        fgSalarié.Rows = fgSalarié.Rows + 1
        fgSalarié.Row = fgSalarié.Rows - 1
        fgSalarié_DisplayLine
    End If
Next arrDRH_Index
If fgSalarié.Rows > 1 Then fgSalarié_Sort

End Sub

Private Sub fgTR_Display()

fgTR.Visible = True
fgTR.Clear: fgTR.Rows = 1

fgTR.Rows = 1
fgTR.FormatString = fgTR_FormatString
fgTR.Enabled = True
For arrDRH_Index = 1 To arrDRH_NB
    If arrDRH(arrDRH_Index).Method <> constDelete Then
        fgTR.Rows = fgTR.Rows + 1
        fgTR.Row = fgTR.Rows - 1
        fgTR_DisplayLine
    End If
Next arrDRH_Index
If fgTR.Rows > 1 Then fgTR_Sort

End Sub


Public Sub fgSalarié_DisplayLine()
fgSalarié_K = (fgSalarié.Row) * fgSalarié.Cols
fgSalarié.TextArray(1 + fgSalarié_K) = arrDRH(arrDRH_Index).Service
fgSalarié.TextArray(0 + fgSalarié_K) = arrDRH(arrDRH_Index).Matricule
fgSalarié.TextArray(2 + fgSalarié_K) = Trim(arrDRH(arrDRH_Index).Nom) & " " & Trim(arrDRH(arrDRH_Index).Prénom)
fgSalarié.TextArray(3 + fgSalarié_K) = arrDRH_Index

End Sub

Public Sub fgTR_DisplayLine()
fgTR_K = (fgTR.Row) * fgTR.Cols
fgTR.TextArray(1 + fgTR_K) = arrDRH(arrDRH_Index).Service
fgTR.TextArray(0 + fgTR_K) = arrDRH(arrDRH_Index).Matricule
fgTR.TextArray(2 + fgTR_K) = Trim(arrDRH(arrDRH_Index).Nom) & " " & Trim(arrDRH(arrDRH_Index).Prénom)
fgTR.TextArray(3 + fgTR_K) = arrDRHNbjOuvrés(arrDRH_Index, 0)
wNbj = arrDRHTR(arrDRH_Index)
fgTR.TextArray(4 + fgTR_K) = wNbj
If wNbj < 0 Then fgTR.Col = 4: fgTR.CellForeColor = errUsr.ForeColor
fgTR.TextArray(6 + fgTR_K) = arrDRH_Index

End Sub

Public Sub fgSalarié_Load()

recDRH_Init xDRH
arrDRH_NBMax = 0: ReDim arrDRH(0)

xDRH.Method = "SnapP0"

arrDRH(0) = xDRH
arrDRH(0).Matricule = "99999"

arrDRH_Suite = True: arrDRH_NB = 0
Do Until Not arrDRH_Suite
    srvDRH_Monitor xDRH
    xDRH = arrDRH(arrDRH_NB)
    xDRH.Method = "SnapP0+"
Loop
fgSalarié_Display
End Sub
Public Sub fgSalarié_Sort()
If fgSalarié.Rows > 1 Then
    fgSalarié.Row = 1
    fgSalarié.RowSel = fgSalarié.Rows - 1
    
    fgSalarié.Col = fgSalarié_Sort1
    fgSalarié.ColSel = fgSalarié_Sort2
    fgSalarié.Sort = 1
End If
End Sub

Public Sub fgTR_Sort()
If fgTR.Rows > 1 Then
    fgTR.Row = 1
    fgTR.RowSel = fgTR.Rows - 1
    
    fgTR.Col = fgTR_Sort1
    fgTR.ColSel = fgTR_Sort2
    fgTR.Sort = 1
End If
End Sub

Private Sub mnuDRHCréer2_Click()
mnuDRHCréer_Click
End Sub

Private Sub mnuDRHDisplay_Click()
If DrhAut.Consulter Then
   
    mDRH.Method = "SnapL0" ' 2000-08-10 Pb IdSeq sur lignes annulées  "SnapL0"
    If chkOnglet1 = "1" Then SSTab1.Tab = 1
    currentAction = "Display"
    xDRH = mDRH
    fraSalarié_display
    fraSalarié.Enabled = False
    cmdSalariéOK.Visible = False
    cmdContext.Caption = constcmdAbandonner
    blnControl = True
    fgSalariéMvt_Load mDRH.Method
End If

End Sub

Private Sub mnuDRHEffacer_Click()
If DrhAut.Xspécial Then
    Call MsgBox("NE SUPPRIME PAS LES MOUVEMENTS ....", vbInformation, "DRH")
    mDRH.Method = constDelete
    SSTab1.Tab = 2
    currentAction = constDelete
    xDRH = mDRH
    fraSalarié_display
    fraSalarié.Enabled = True
    fraSalariéId.Enabled = False
    fraSalariéDivers.Enabled = False
    cmdSalariéOK.Caption = "Supprimer"
    cmdSalariéOK.Visible = True
    cmdContext.Caption = constcmdAbandonner
    blnControl = True
End If

End Sub

Private Sub mnuDRHModifier_Click()
If DrhAut.Valider Then
   
    mDRH.Method = constUpdate
    SSTab1.Tab = 2
    currentAction = constUpdate
    xDRH = mDRH
    fraSalarié_display
    txtSortieAmj.Visible = False
    txtMatricule.Enabled = False
    fraSalarié.Enabled = True
    fraSalariéId.Enabled = True
    fraSalariéDivers.Enabled = True
    cmdSalariéOK.Visible = True
    cmdSalariéOK.Caption = "Modifier"
    txtNom.SetFocus
    cmdContext.Caption = constcmdAbandonner
    blnControl = True
End If

End Sub

Private Sub mnuDRHMvtAnnuler_Click()
xDRHMvt = mDRHMvt
xDRHMvt.Method = constUpdate
xDRHMvt.Statut = "A"
cmdMouvement_Update

End Sub

Private Sub mnuDRHMvtEffacer_Click()
xDRHMvt = mDRHMvt
xDRHMvt.Method = constDelete
cmdMouvement_Update
End Sub

Private Sub mnuDRHMvtModifier_Click()
fraMouvement.Visible = True
oldDRHMvt = mDRHMvt
cmdMouvementOK.Caption = constModifier
lstMvt_Select
lstMvt.Enabled = True
End Sub

Private Sub mnuFérié_Suppress_Click()
xTable.Method = "Seek="
xTable.Id = "DRH"
xTable.K1 = constDRHFérié
xTable.K2 = ""
X = mId$(lstFérié, 1, 14)
xTable.K2 = mId$(lstFérié, 11, 4) & mId$(lstFérié, 6, 2) & mId$(lstFérié, 1, 2)
intReturn = tableElpTable_Read(xTable)
If intReturn <> 0 Then Call lstErr_Clear(lstErr, txtFériéAmj, V & " (err : " & intReturn & " ) "):   Exit Sub

xTable.Method = "Delete"
intReturn = tableElpTable_Update(xTable)
If intReturn <> 0 Then Call lstErr_Clear(lstErr, txtFériéAmj, V & " non supprimé (err : " & intReturn & " ) "):   Exit Sub

lstFérié_Load

End Sub

Private Sub mnuPrintAbsence_Click()
prtSort = "1"
prtDocument = "02"
prtDébutAmj = prtAmj
prtFinAmj = "99999999"
prtSelectK = " ": prtSelect = " "
prtSelectMvtK = " ": prtSelectMvt = " "
prtEnTête = "DRH : Etat des absences du " & dateImp(prtAmj)

cmdMouvement_Import
prtDestinataire = "": cmdPrint_Monitor
prtDestinataire = "Directeur général": cmdPrint_Monitor
prtDestinataire = "Directeur général adjoint": cmdPrint_Monitor
cmdReset
End Sub

Private Sub mnuPrintAM_Click()
prtSort = "1"
prtDocument = "03"
prtFinAmj = dateFinDeMois(prtAmj)
prtDébutAmj = dateElp("MoisAdd", -11, mId$(prtFinAmj, 1, 6) & "01")
prtSelectK = " ": prtSelect = " "
prtSelectMvtK = " ": prtSelectMvt = " "
prtEnTête = "DRH : Récapitulatif des arrêts maladie (jours ouvrés)"

prtDRHMvt_Nbj_Init prtDébutAmj
cmdMouvement_Import
fgSalarié_Sort1 = 1: fgSalarié_Sort2 = 2: fgSalarié_Sort

cmdPrint_Monitor

prtDocument = "04"
prtEnTête = "DRH : Récapitulatif des arrêts maladie (jours civils)"
cmdPrint_Monitor

End Sub


Private Sub mnuPrintJournal_Click()
prtSort = "5"
prtDocument = "01"
prtDébutAmj = prtAmj
prtFinAmj = prtAmj
prtSelectK = " ": prtSelect = " "
prtSelectMvtK = " ": prtSelectMvt = " "
prtEnTête = "DRH : Journal des mouvements du " & dateImp(prtAmj)

cmdMouvement_Import
cmdPrint_Monitor

End Sub


Private Sub optCivilitéAutre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optCivilitéAutre
End Sub


Private Sub optCivilitéM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optCivilitéM
End Sub


Private Sub optCivilitéMle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optCivilitéMle
End Sub


Private Sub optCivilitéMme_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optCivilitéMme
End Sub


Private Sub optNatureS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optNatureS
End Sub


Private Sub optNatureX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optNatureX
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
    Case 3: lstFérié_Load
End Select

End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set SSTab1

End Sub


Private Sub txtBureau_GotFocus()
txt_GotFocus txtBureau
End Sub


Private Sub txtBureau_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtBureau_LostFocus()
txt_LostFocus txtBureau
If blnControl Then cmdSalarié_Control
End Sub


Private Sub txtDébutAmj_Change()
If blnControl Then cmdMouvement_Control

End Sub

Private Sub txtDébutAmj_GotFocus()
DTPicker_GotFocus txtDébutAmj
End Sub


Private Sub txtDébutAmj_LostFocus()
DTPicker_LostFocus txtDébutAmj
If blnControl Then cmdSalarié_Control
End Sub


Private Sub txtEnfantNb_GotFocus()
txt_GotFocus txtEnfantNb
End Sub


Private Sub txtEnfantNb_LostFocus()
txt_LostFocus txtEnfantNb
If blnControl Then cmdSalarié_Control
End Sub


Private Sub txtEntréeAmj_Change()
If blnControl Then cmdSalarié_Control

End Sub

Private Sub txtEntréeAmj_GotFocus()
DTPicker_GotFocus txtEntréeAmj
End Sub


Private Sub txtEntréeAmj_LostFocus()
DTPicker_LostFocus txtEntréeAmj
End Sub


Private Sub txtExerciceCivilAmj_GotFocus()
DTPicker_GotFocus txtExerciceCivilAmj

End Sub


Private Sub txtExerciceCivilAmj_LostFocus()
DTPicker_LostFocus txtExerciceCivilAmj

End Sub


Private Sub txtExerciceCivilNbj_GotFocus()
txt_GotFocus txtExerciceCivilNbj

End Sub


Private Sub txtExerciceCivilNbj_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub


Private Sub txtExerciceCivilNbj_LostFocus()
txt_LostFocus txtExerciceCivilNbj
End Sub


Private Sub txtFériéAmj_GotFocus()
DTPicker_GotFocus txtFériéAmj
End Sub


Private Sub txtFériéAmj_LostFocus()
DTPicker_LostFocus txtFériéAmj
End Sub


Private Sub txtMatricule_GotFocus()
txt_GotFocus txtMatricule
End Sub


Private Sub txtMatricule_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtMatricule_LostFocus()
txt_LostFocus txtMatricule
If blnControl Then cmdSalarié_Control

End Sub



Private Sub txtNbj_Change()
If blnControl Then cmdMouvement_Control

End Sub

Private Sub txtnbj_GotFocus()
txt_GotFocus txtNbj
End Sub


Private Sub txtnbj_KeyPress(KeyAscii As Integer)
num_KeyAsciiD KeyAscii, txtNbj
End Sub


Private Sub txtnbj_LostFocus()
txt_LostFocus txtNbj
If blnControl Then cmdMouvement_Control
End Sub


Private Sub txtNom_GotFocus()
txt_GotFocus txtNom
End Sub


Private Sub txtNom_LostFocus()
txt_LostFocus txtNom
If blnControl Then cmdSalarié_Control
End Sub


Private Sub txtPrénom_GotFocus()
txt_GotFocus txtPrénom
End Sub


Private Sub txtPrénom_LostFocus()
txt_LostFocus txtPrénom
If blnControl Then cmdSalarié_Control

End Sub


Private Sub txtcompte_GotFocus()
txt_GotFocus txtCompte
End Sub


Private Sub txtcompte_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub


Private Sub txtcompte_LostFocus()
txt_LostFocus txtCompte
If blnControl Then cmdSalarié_Control

End Sub


Private Sub txtPrtAmj_Change()
Call DTPicker_Control(txtPrtAmj, prtAmj)
End Sub

Private Sub txtPrtAmj_GotFocus()
DTPicker_GotFocus txtFériéAmj

End Sub


Private Sub txtPrtAmj_LostFocus()
DTPicker_LostFocus txtDébutAmj

End Sub


Private Sub txtRéfInterne_GotFocus()
txt_GotFocus txtRéfInterne
End Sub


Private Sub txtRéfInterne_LostFocus()
txt_LostFocus txtRéfInterne
End Sub


Private Sub txtRepriseAmj_Change()
If blnControl Then cmdMouvement_Control

End Sub

Private Sub txtRepriseAmj_GotFocus()
DTPicker_GotFocus txtRepriseAmj
End Sub


Private Sub txtRepriseAmj_LostFocus()
DTPicker_LostFocus txtRepriseAmj

End Sub


Private Sub txtSortieAmj_GotFocus()
DTPicker_GotFocus txtSortieAmj
End Sub


Private Sub txtSortieAmj_LostFocus()
DTPicker_LostFocus txtSortieAmj
If blnControl Then cmdSalarié_Control
End Sub


Private Sub txtTéléphone1_GotFocus()
txt_GotFocus txtTéléphone1
End Sub


Private Sub txtTéléphone1_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtTéléphone1_LostFocus()
txt_LostFocus txtTéléphone1
If blnControl Then cmdSalarié_Control
End Sub


Private Sub txtTéléphone2_GotFocus()
txt_GotFocus txtTéléphone2
End Sub


Private Sub txtTéléphone2_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtTéléphone2_LostFocus()
txt_LostFocus txtTéléphone2
If blnControl Then cmdSalarié_Control
End Sub


Private Sub txtTéléphone3_GotFocus()
txt_GotFocus txtTéléphone3
End Sub


Private Sub txtTéléphone3_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtTéléphone3_LostFocus()
txt_LostFocus txtTéléphone3
If blnControl Then cmdSalarié_Control
End Sub



Public Sub cmdSalarié_Control()
Dim X As String, V As Variant

If Not frmDRH.Enabled Then Exit Sub
frmDRH.Enabled = False

cmdSalariéOK.Visible = False
blnControl = False

lstErr.Clear
lstErr.Height = 200

xDRH = mDRH

X = num_Control(txtMatricule, valX, 5, 0)
xDRH.Matricule = valX
If xDRH.Matricule = "00000" Then Call lstErr_AddItem(lstErr, cmdContext, "? Préciser le matricule"): GoTo ExitSub
If xDRH.Method = constAddNew Then
    If arrDRH_Scan(xDRH) > 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? Ce matricule existe déjà"): GoTo ExitSub
End If
X = Trim(txtNom)
If Trim(X) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? Préciser le nom"): GoTo ExitSub
xDRH.Nom = X

If optNatureS Then xDRH.Nature = "S"
If optNatureX Then xDRH.Nature = "X"
If optCivilitéM Then xDRH.Civilité = "1"
If optCivilitéMme Then xDRH.Civilité = "2"
If optCivilitéMle Then xDRH.Civilité = "3"
If optCivilitéAutre Then xDRH.Civilité = "4"
   
xDRH.Prénom = Trim(txtPrénom)
V = DTPicker_Control(txtEntréeAmj, xDRH.EntréeAmj)
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, txtEntréeAmj, V): GoTo ExitSub

If chkSortieAmj = "1" Then
    txtSortieAmj.Visible = True
    V = DTPicker_Control(txtSortieAmj, xDRH.SortieAmj)
    If Not IsNull(V) Then Call lstErr_AddItem(lstErr, txtSortieAmj, V): GoTo ExitSub
    If txtSortieAmj < txtEntréeAmj Then Call lstErr_AddItem(lstErr, txtSortieAmj, "? date sortie < date d'entrée"): GoTo ExitSub
End If

X = num_Control(txtEnfantNb, valX, 3, 0)
xDRH.EnfantNb = valX
    
X = num_Control(txtCompte, valX, 11, 0)
xDRH.Compte = valX
If lstService.ListIndex < 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? Préciser le service"): GoTo ExitSub
xDRH.Service = mId$(lstService.Text, 1, 4)
xDRH.Bureau = Trim(txtBureau)
xDRH.Téléphone1 = Trim(txtTéléphone1)
xDRH.Téléphone2 = Trim(txtTéléphone2)
xDRH.Téléphone3 = Trim(txtTéléphone3)
xDRH.RéfInterne = Trim(txtRéfInterne)

If lstErr.ListCount = 0 Then
    cmdSalariéOK.Visible = True
End If

ExitSub:

frmDRH.Enabled = True
'If cmdSalariéOK.Visible Then cmdSalariéOK.SetFocus
    
blnControl = True

End Sub
Public Sub cmdMouvement_Control()
Dim X As String, V As Variant

If Not frmDRH.Enabled Then Exit Sub
frmDRH.Enabled = False

cmdMouvementOK.Visible = False
blnControl = False

lstErr.Clear
lstErr.Height = 200

xDRHMvt = mDRHMvt
'''xDRHMvt.RéfInterne = Trim(txtMvtRéfInterne)
txtNbjOuvré = ""
xDRHMvt.RepriseChk = "0"
xDRHMvt.RepriseAmjK = "0"
txtRepriseAmj.Enabled = False
chkRepriseAmjK.Enabled = False

xDRHMvt.DébutAmjK = IIf(chkDébutAmjK = "1", "1", "0")
V = DTPicker_Control(txtDébutAmj, xDRHMvt.DébutAmj)
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, "? date début : " & V)

If chkNbj = "1" Then
    xDRHMvt.NbjChk = "1"
    txtNbj.Enabled = True
Else
    xDRHMvt.NbjChk = "0"
    txtNbj.Enabled = False
    xDRHMvt.Nbj = 0: txtNbj = ""
    xDRHMvt.RepriseAmjK = "0"
    xDRHMvt.RepriseAmj = "29991231"
    GoTo ExitSub
End If

xDRHMvt.Nbj = CDbl(Val(txtNbj)) 'valX
I = xDRHMvt.Nbj * 10 Mod 10
If I <> 0 And I <> 5 Then Call lstErr_AddItem(lstErr, cmdContext, "? par demi-journée")
If xDRHMvt.Nbj = 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? Préciser le nombre de jours")

xDRHMvt.RepriseAmjK = IIf(chkRepriseAmjK = "1", "1", "0")

V = srvDRHMvt_RepriseAmj(xDRHMvt, mDRHCalendrier)

If Not IsNull(V) Then
    Call lstErr_AddItem(lstErr, cmdContext, V)
Else
    Call DTPicker_Set(txtRepriseAmj, xDRHMvt.RepriseAmj)
    chkRepriseAmjK = xDRHMvt.RepriseAmjK
End If

txtNbjOuvré = Format$(xDRHMvt.NbjOuvré, "###.#")

If xDRHMvt.MvtSens = "-" Then
    If Not cmdMouvement_Control_Période Then Call lstErr_AddItem(lstErr, cmdContext, "? chevauchement des périodes ")
    
    If xDRHMvt.DébutAmj < xDRH.EntréeAmj Then Call lstErr_AddItem(lstErr, cmdContext, "? date début < date d'entrée BIA")
    If xDRHMvt.DébutAmj < paramDRHMvt.DébutAmj Then Call lstErr_AddItem(lstErr, cmdContext, "? date début < période en cours ")
    If xDRHMvt.DébutAmj > xDRHMvt.RepriseAmj Then
        Call lstErr_AddItem(lstErr, cmdContext, "? date début > date de reprise ")
    Else
        If xDRHMvt.DébutAmj = xDRHMvt.RepriseAmj Then
            If xDRHMvt.DébutAmjK <> "0" Or xDRHMvt.RepriseAmjK <> "1" Then
                Call lstErr_AddItem(lstErr, cmdContext, "? date début > date de reprise ")
            End If
        End If
    End If
    
    If xDRHMvt.DébutAmj < paramDRHMvt.RepriseAmj And xDRHMvt.RepriseAmj > paramDRHMvt.RepriseAmj Then Call lstErr_AddItem(lstErr, cmdContext, "? chevauchement d'exercice ")
End If

ExitSub:

If lstErr.ListCount = 0 Then
    cmdMouvementOK.Visible = DrhAut.Saisir
End If


frmDRH.Enabled = True
    
blnControl = True

End Sub


Public Sub fraSalarié_display()
Dim X As String, V As Variant
fgSalarié_RowClick = 0
Call fgSalarié_Color(fgSalarié_RowDisplay, vbCyan, fgSalarié_ColorClick) 'txtUsr.BackColor)
fgSalariéMvt.Clear
libSalarié = Trim(xDRH.Matricule) & "_" & Trim(xDRH.Nom) & " " & Trim(xDRH.Prénom)
txtMatricule = Trim(xDRH.Matricule)
Call lst_Scan(xDRH.Service, lstService)
txtNom = Trim(xDRH.Nom)
txtPrénom = Trim(xDRH.Prénom)
txtEnfantNb = Trim(xDRH.EnfantNb)
txtCompte = Trim(xDRH.Compte)
txtBureau = Trim(xDRH.Bureau)
txtTéléphone1 = Trim(xDRH.Téléphone1)
txtTéléphone2 = Trim(xDRH.Téléphone2)
txtTéléphone3 = Trim(xDRH.Téléphone3)
Call DTPicker_Set(txtEntréeAmj, xDRH.EntréeAmj)
If Trim(xDRH.SortieAmj) = "" Then
    txtSortieAmj.Visible = False
    chkSortieAmj.Value = "0"
    Call DTPicker_Set(txtSortieAmj, DSys)
Else
    fraSalarié.Enabled = False
    chkSortieAmj.Value = "1"
    txtSortieAmj.Visible = True
    Call DTPicker_Set(txtSortieAmj, xDRH.SortieAmj)
End If
txtRéfInterne = Trim(xDRH.RéfInterne)
libUpd = dateImp(xDRH.UpdAmj) & "  " & timeImp(xDRH.UpdHms)
If xDRH.Nature = "S" Then optNatureS = True
If xDRH.Nature = "X" Then optNatureX = True
If xDRH.Civilité = "1" Then optCivilitéM = True
If xDRH.Civilité = "2" Then optCivilitéMme = True
If xDRH.Civilité = "3" Then optCivilitéMle = True
If xDRH.Civilité = "4" Then optCivilitéAutre = True
   
End Sub

Private Sub txtTRNbj_GotFocus()
txt_GotFocus txtTRNbj
End Sub


Private Sub txtTRNbj_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtTRNbj_LostFocus()
'txt_LostFocus txtnbjTRNbj
End Sub



Public Sub lstFérié_Load()

lstFérié.Clear
recElpTable_Init xTable
xTable.Method = "Seek>="
xTable.Id = "DRH"
xTable.K1 = constDRHFérié

Do
    intReturn = tableElpTable_Read(xTable)
    If intReturn = 0 Then
        If constDRHFérié <> Trim(xTable.K1) Then
            intReturn = -1
        Else
            X = dateImp(xTable.K2) & "  " & Trim(xTable.Name)
            If mId$(xTable.Memo, 1, 2) = "0X" Then X = X & " ... après-midi"
            lstFérié.AddItem X
        End If
    End If
    
    xTable.Method = "MoveNext"
  
Loop While intReturn = 0

End Sub
Public Sub lstService_Load()

lstService.Clear
recElpTable_Init xTable
xTable.Method = "Seek>="
xTable.Id = "DRH"
xTable.K1 = constDRHService

Do
    intReturn = tableElpTable_Read(xTable)
    If intReturn = 0 Then
        If constDRHService <> Trim(xTable.K1) Then
            intReturn = -1
        Else
            lstService.AddItem mId$(xTable.K2, 1, 4) & Chr$(9) & xTable.Name
        End If
    End If
    
    xTable.Method = "MoveNext"
  
Loop While intReturn = 0

End Sub

Public Sub lstMvt_Load()

lstMvt.Clear: lstPrintMouvement.Clear
recElpTable_Init xTable
xTable.Method = "Seek>="
xTable.Id = "DRH"
xTable.K1 = constDRHMvt

Do
    intReturn = tableElpTable_Read(xTable)
    If intReturn = 0 Then
        If constDRHMvt <> Trim(xTable.K1) Then
            intReturn = -1
        Else
            lstMvt.AddItem mId$(xTable.K2, 1, 4) & " : " & xTable.Name
            lstPrintMouvement.AddItem mId$(xTable.K2, 1, 4) & " : " & xTable.Name
            lstPrintMouvement.Selected(lstPrintMouvement.ListCount - 1) = False
        End If
    End If
    
    xTable.Method = "MoveNext"
  
Loop While intReturn = 0

End Sub


Public Sub fgTotal_Load()

recElpTable_Init xTable
xTable.Method = "Seek="
xTable.Id = "DRH"
xTable.K1 = "Mvt$"

For I = 1 To 99 ' fgTotal.Rows - 1
    xTable.K2 = Format$(I, "00")
    intReturn = tableElpTable_Read(xTable)
    If intReturn = 0 Then
          arrDRHMvt_Libellé(I) = Trim(xTable.Name)
    Else
          arrDRHMvt_Libellé(I) = I & " ???? "
    End If
        

Next I



End Sub
Public Sub fgTotal_Display()
Dim lSolde As Double

fgTotal.Clear: fgTotal.Rows = 1
fgTotal.Rows = 1
fgTotal.FormatString = fgtotal_FormatString
fgTotal.Enabled = False

For I = 1 To 99 ' fgTotal.Rows - 1
    If arrDRHMvt_Absences_Nb(I) <> 0 Or arrDRHMvt_Droits_Nb(I) <> 0 Then
        fgTotal.Rows = fgTotal.Rows + 1
        fgTotal.Row = fgTotal.Rows - 1
        fgTotal.Col = 0
        fgTotal.Text = arrDRHMvt_Libellé(I)
        
        If arrDRHMvt_Absences_Nb(I) <> 0 Then
            fgTotal.Col = 2
            fgTotal.CellForeColor = errUsr.ForeColor
            fgTotal.Text = Format$(arrDRHMvt_Absences_Nb(I), "###.0")
        End If
          If arrDRHMvt_Droits_Nb(I) <> 0 Then
            fgTotal.Col = 1
            fgTotal.Text = Format$(arrDRHMvt_Droits_Nb(I), "###.0")
        End If
        lSolde = arrDRHMvt_Droits_Nb(I) - arrDRHMvt_Absences_Nb(I)
        fgTotal.Col = 3
        If lSolde < 0 Then
            fgTotal.CellForeColor = errUsr.ForeColor
        'Else
        '     fgTotal.CellForeColor = vbCyan
       End If
        fgTotal.Text = Format$(lSolde, "###.0")
   End If
Next I



End Sub


Public Sub fraMouvement_Reset()
cmdMouvementSaisir.Enabled = True
'fraMouvement.Visible = False
fraMouvement.Enabled = True
fraMouvementDétail.Enabled = False
cmdMouvementOK.Visible = False
cmdMouvementOK.Caption = constAjouter
cmdMouvementHisto.Visible = True
lstMvt.Enabled = True
chkNbj = "1": txtNbj.Enabled = True
'chkRepriseAmj = "0":
txtRepriseAmj.Enabled = False: chkRepriseAmjK.Enabled = False
txtNbj = "": txtNbjOuvré = ""
DTPicker_Now txtDébutAmj
DTPicker_Now txtRepriseAmj
chkDébutAmjK = "0"
chkRepriseAmjK = "0"
'txtNbj.SetFocus
'txtMvtRéfInterne.Enabled = False
End Sub

Public Sub paramDRHMvt_Init(X4 As String)
Dim X2 As String

recDRHMvt_Init paramDRHMvt

recElpTable_Init xTable
xTable.Method = "Seek="
xTable.Id = "DRH"
xTable.K1 = "Mvt"
xTable.K2 = X4
intReturn = tableElpTable_Read(xTable)
If intReturn <> 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? lecture Mvt")
paramDRHMvt.MvtCode = X4
paramDRHMvt.MvtSens = mId$(xTable.Memo, 1, 1)
paramDRHMvt.Statut = " "

X2 = mId$(xTable.Memo, 7, 2)
If Not IsNumeric(X2) Then
    paramDRHMvt_TotalK = 91
Else
    paramDRHMvt_TotalK = Val(X2)
    If paramDRHMvt_TotalK < 1 Or paramDRHMvt_TotalK > 99 Then paramDRHMvt_TotalK = 91
End If
'jpl 2000.12.21'mId$(xTable.Memo, 3, 1)

'jpl 2000.12.21 xTable.K1 = "$Mvt"
'jpl 2000.12.21 xTable.K2 = mId$(X4, 1, 2)
'jpl 2000.12.21 intReturn = tableElpTable_Read(xTable)
'jpl 2000.12.21 If intReturn <> 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? lecture $Mvt")

paramDRHMvt.MvtCO = mId$(xTable.Memo, 3, 1)

Select Case mId$(xTable.Memo, 5, 1)
    Case "1"
        paramDRHMvt.DébutAmj = mCP.DébutAmj
        paramDRHMvt.RepriseAmj = mCP.RepriseAmj
    Case Else
        paramDRHMvt.DébutAmj = mCivil.DébutAmj
        paramDRHMvt.RepriseAmj = mCivil.RepriseAmj
End Select

If paramDRHMvt.MvtCO = "C" Then
    lblNbj.Caption = "Nb jours calendaires"
Else
    lblNbj.Caption = "Nb jours ouvrés"
End If
End Sub


Public Sub fraMouvement_Display()
blnControl = False
Call fgSalariéMvt_Color(fgSalariéMvt_RowDisplay, MouseMoveUsr.BackColor, fgSalariéMvt_ColorDisplay)
fraMouvement.Enabled = True
Call lst_Scan(xDRHMvt.MvtCode, lstMvt)
txtNbj = Trim(xDRHMvt.Nbj)
Call DTPicker_Set(txtDébutAmj, xDRHMvt.DébutAmj)
Call DTPicker_Set(txtRepriseAmj, xDRHMvt.RepriseAmj)

If xDRHMvt.DébutAmjK = "1" Then chkDébutAmjK = "1"
If xDRHMvt.RepriseAmjK = "1" Then chkRepriseAmjK = "1"
chkNbj = xDRHMvt.NbjChk
'If xDRHMvt.RepriseChk = "1" Then chkRepriseAmj = "1"
'txtMvtRéfInterne = xDRHMvt.RéfInterne
libMvtUpd = dateImp(xDRHMvt.UpdAmj) & " " & timeImp(xDRHMvt.UpdHms)
blnControl = True
End Sub

Public Sub cmdMouvement_Update()
frmDRH.Enabled = False

blnControl = False
xDRHMvt.UpdAmj = DSys
xDRHMvt.UpdHms = time_Hms

V = srvDRHMvt_Update(xDRHMvt)

If IsNull(V) Then
    Select Case xDRHMvt.Method
        Case constUpdate: arrDRHMvt(oldDRHMvt_Index) = xDRHMvt
        Case constDelete: arrDRHMvt(oldDRHMvt_Index) = xDRHMvt
        Case constAddNew: Call arrDRHMvt_AddItem(xDRHMvt)
    End Select
    fgSalariéMvt_Display
    fraMouvement_Reset
    Call lstErr_Clear(lstErr, cmdContext, "Mise à jour effectuée - Matricule : " & xDRH.Matricule & "_" & Trim(xDRHMvt.RéfInterne))
    'cmdReset
Else
    Call lstErr_Clear(lstErr, cmdContext, V)
End If
frmDRH.Enabled = True
AppActivate frmDRH.Caption

End Sub

Public Sub fgSalarié_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer
mRow = fgSalarié.Row

If lRow > 0 Then
    fgSalarié.Row = lRow
    fgSalarié.Col = 0: fgSalarié.CellBackColor = lColor_Old 'fgSalarié.BackColorBkg
    fgSalarié.Col = 1: fgSalarié.CellBackColor = lColor_Old
    fgSalarié.Col = 2: fgSalarié.CellBackColor = lColor_Old
End If
lRow = 0
If mRow > 0 Then
    fgSalarié.Row = mRow
    If fgSalarié.Row > 0 Then
        lRow = fgSalarié.Row
        lColor_Old = fgSalarié.CellBackColor
        fgSalarié.Col = 0: fgSalarié.CellBackColor = lColor
        fgSalarié.Col = 1: fgSalarié.CellBackColor = lColor
        fgSalarié.Col = 2: fgSalarié.CellBackColor = lColor
    End If
End If

End Sub
Public Sub fgSalariéMvt_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, K As Integer

mRow = fgSalariéMvt.Row

If lRow > 0 Then
    If lRow < fgSalariéMvt.Rows - 1 Then fgSalariéMvt.Row = lRow
    For K = 0 To 9
        fgSalariéMvt.Col = K: fgSalariéMvt.CellBackColor = lColor_Old
    Next K
End If
lRow = 0
fgSalariéMvt.Row = mRow
If fgSalariéMvt.Row > 0 Then
    lRow = fgSalariéMvt.Row
    lColor_Old = fgSalariéMvt.CellBackColor
    For K = 0 To 9
        fgSalariéMvt.Col = K: fgSalariéMvt.CellBackColor = lColor
    Next K
End If

End Sub


Public Sub cmdContext_mnu()
mnuDRHCréer2.Enabled = DrhAut.Saisir
mnuPrintAbsence.Enabled = True   '02
mnuPrintAM.Enabled = True    '03
mnuPrintPlanning.Enabled = False
mnuPrintSelect.Enabled = False
mnuPrintRecap.Enabled = False
mnuXOuvertureCp.Enabled = False
mnuXClôtureCp.Enabled = False
mnuXClôtureCivil.Enabled = False
Me.PopupMenu mnuContext, vbPopupMenuLeftButton
End Sub

Public Sub cmdPrint_Monitor()
Dim X, Nb As Integer, curX As Currency
Dim Msg As String

If cmdImport_Select_Nb = 0 Then
    Call lstErr_AddItem(lstErr, cmdPrint, "Aucun salarié sélectionné !")
    'GoTo cmdPrint_End
End If

Msg = "000000000000" & Space$(100)
Mid$(Msg, 14, 2) = prtDocument
Mid$(Msg, 16, 1) = prtSort
Mid$(Msg, 17, 8) = prtDébutAmj
Mid$(Msg, 25, 8) = prtFinAmj
Mid$(Msg, 33, 1) = prtSelectK
Mid$(Msg, 34, 16) = prtSelect
Mid$(Msg, 50, 1) = prtSelectMvtK
Mid$(Msg, 51, 4) = prtSelectMvt

prtDRHMvt_Open Msg, prtEnTête, prtDestinataire
Call lstErr_Clear(lstErr, cmdPrint, "Impression : début")

Select Case prtDocument
    Case "01", "02", "06", "07": cmdPrint_MvtP0
    Case "03", "04": cmdPrint_03
    Case "05": cmdPrint_05
    Case "10": cmdPrint_10
End Select

prtDRHMvt_Close
Call lstErr_AddItem(lstErr, cmdPrint, "Impression terminé : " & cmdImport_Select_Nb)

cmdPrint_End:
Me.Enabled = True
AppActivate Me.Caption

End Sub
Private Sub cmdMouvement_Import()
On Error Resume Next

Dim I As Integer, X As String, iReturn As Integer, iScan As Integer
Dim blnOk As Boolean

cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0:

Me.MousePointer = vbHourglass
Me.Enabled = False

recDRH_Init xDRH
recDRHMvt_Init xDRHMvt
Select Case prtDocument
    Case "01": xDRHMvt.Method = "SnapLU"
                xDRHMvt.UpdAmj = prtAmj
                mDRHMvt = xDRHMvt
                mDRHMvt.Matricule = "99999"

    Case "02": xDRHMvt.Method = "SnapLE"
                xDRHMvt.RepriseAmj = prtAmj + 1
                mDRHMvt = xDRHMvt
                mDRHMvt.RepriseAmj = "99999999"

    Case "03": xDRHMvt.Method = "SnapLE"
                xDRHMvt.RepriseAmj = prtDébutAmj
                mDRHMvt = xDRHMvt
                mDRHMvt.RepriseAmj = "99999999"
     Case "06": xDRHMvt.Method = "SnapLE"
                xDRHMvt.RepriseAmj = prtDébutAmj
                mDRHMvt = xDRHMvt
                mDRHMvt.RepriseAmj = "99999999"
      Case "07": xDRHMvt.Method = "SnapLD"
                xDRHMvt.DébutAmj = prtDébutAmj
                mDRHMvt = xDRHMvt
                mDRHMvt.DébutAmj = prtFinAmj
                mDRHMvt.Matricule = "99999"
  
End Select



Call srvDRHMvt_ElpBuffer(xDRHMvt, mDRHMvt, xElpBuffer)
Me.MousePointer = 0
If xElpBuffer.Seq = 0 Then Exit Sub


MDB.Execute "delete * from mvtp0"
mdbMvtP0.tableMvtP0_Open

recMvtP0_Init recMvtp0
recMvtp0.Method = "AddNew"

xElpBuffer.Method = "Seek=" '
xElpBuffer.Seq = 1

Do
    iReturn = tableElpBuffer_Read(xElpBuffer)
    If iReturn = 0 Then
        blnOk = True
        cmdImport_Nb = cmdImport_Nb + 1
        MsgTxt = xElpBuffer.Data
        MsgTxtIndex = 0
        srvDRHMvt_GetBuffer xDRHMvt
        Select Case prtDocument
            Case "02":  If xDRHMvt.DébutAmj > prtDébutAmj Then blnOk = False
            Case "03":
                        If xDRHMvt.DébutAmj > prtFinAmj Then
                            blnOk = False
                        Else
                            Select Case mId$(xDRHMvt.MvtCode, 1, 2)
                                Case "03", "04"
                                Case "05": If chkPrintEnfantMalade <> "1" Or xDRHMvt.MvtCode <> "0510" Then blnOk = False
                                Case Else: blnOk = False
                            End Select
                        End If
            Case "06":
                        If xDRHMvt.DébutAmj > prtFinAmj Then blnOk = False
                        If xDRHMvt.MvtCode = "0610" Then blnOk = False  ' formation : droit à TR
            
            Case "07":  If xDRHMvt.DébutAmj > prtFinAmj Then blnOk = False
                        If xDRHMvt.MvtCode <> "0110" And xDRHMvt.MvtCode <> "0111" Then blnOk = False
         
        End Select
        
        If blnOk Then
            xDRH.Matricule = xDRHMvt.Matricule
            iScan = arrDRH_Scan(xDRH)
            recMvtp0.Id = Space$(40)
            Select Case prtSort
                Case 1: recMvtp0.Id = arrDRH(arrDRH_Index).Service & arrDRH(arrDRH_Index).Nom
            End Select
            cmdImport_Select_Nb = cmdImport_Select_Nb + 1
            Mid$(recMvtp0.Id, 35, 5) = Format$(iScan, "00000")
            Mid$(recMvtp0.Id, 30, 5) = Format$(cmdImport_Select_Nb, "00000")
            recMvtp0.Text = xElpBuffer.Data
         Select Case prtDocument
            Case "01", "02": dbMvtP0_Update recMvtp0
            Case "03": cmdMouvement_Import_Nbj_Cumul
            Case "06": cmdMouvement_Import_TR
            Case "07": dbMvtP0_Update recMvtp0
         
        End Select
      
            
            If I = 1000 Then I = 0: Call lstErr_ChangeLastItem(lstErr, cmdPrint, "Sélection : " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents
        End If
        xElpBuffer.Seq = xElpBuffer.Seq + 1
   End If
Loop Until iReturn <> 0

mdbMvtP0.tableMvtP0_Close

End Sub

Private Sub cmdPrintSelect_Import()
Dim I As Integer, wDRH As typeDRH, mMethod As String

On Error Resume Next

Dim X As String, iReturn As Integer, iScan As Integer
Dim blnOk As Boolean

cmdImport_Select_Nb = 0: cmdImport_Nb = 0: I = 0:

Me.MousePointer = vbHourglass
Me.Enabled = False

recDRHMvt_Init xDRHMvt
recDRH_Init xDRH
wDRH = xDRH
wDRH.Matricule = Format$(Trim(txtPrintMatricule), "00000")
cbo_Value wDRH.Service, cboPrintService
If cboPrintService.ListIndex = -1 Then wDRH.Service = ""

MDB.Execute "delete * from mvtp0"
mdbMvtP0.tableMvtP0_Open

recMvtP0_Init recMvtp0
recMvtp0.Method = "AddNew"
If prtSelectK = "1" Then
    mMethod = "SnapP0"
Else
    mMethod = "SnapL0"
End If


For I = 1 To arrDRH_NB


    blnOk = False
    
    If chkPrintService = "1" Then
        blnOk = True
    Else
        If arrDRH(I).Service = wDRH.Service Then blnOk = True
    End If
     
    If chkPrintSalarié <> "1" Then
        blnOk = False
        If arrDRH(I).Matricule = wDRH.Matricule Then blnOk = True
    End If
    

    If blnOk Then
        xDRHMvt.Method = mMethod
        xDRHMvt.Matricule = arrDRH(I).Matricule
        xDRHMvt.IdSeq = "00000"
        mDRHMvt = xDRHMvt
        mDRHMvt.IdSeq = "99999"
        Call lstErr_ChangeLastItem(lstErr, cmdPrint, "Sélection : " & arrDRH(I).Matricule & " : " & arrDRH(I).Nom): DoEvents
   
        cmdPrintSelect_Import_Mouvement I
    End If
Next I

mdbMvtP0.tableMvtP0_Close

End Sub





Public Sub cmdMouvement_Import_Nbj_Cumul()
Dim blnOk  As Boolean

cumNbjC = 0
If xDRHMvt.MvtCO = "C" Then
    maxNbjC = xDRHMvt.Nbj
Else
    maxNbjC = 999999999
End If

If xDRHMvt.DébutAmj < prtDébutAmj Then
    wAmj1 = prtDébutAmj
    wAmjK1 = "0"
Else
    wAmj1 = xDRHMvt.DébutAmj
    wAmjK1 = xDRHMvt.DébutAmjK
End If

If xDRHMvt.RepriseAmj > prtFinAmj Then
    wAmj2 = prtFinAmj
    wAmjK2 = "0"
Else
    wAmjK2 = xDRHMvt.RepriseAmjK
    If xDRHMvt.RepriseAmjK = "0" Then
        wAmj2 = dateElp("Jour", -1, xDRHMvt.RepriseAmj)
    Else
        wAmj2 = xDRHMvt.RepriseAmj
    End If
    
End If
blnOk = False

Do
    wAmj = dateFinDeMois(wAmj1)
    If wAmj < wAmj2 Then
        Call cmdMouvement_Import_Nbj_Cumul1(wAmj1, wAmjK1, wAmj, "0")
        wAmj1 = dateElp("Jour", 1, wAmj)
    Else
        Call cmdMouvement_Import_Nbj_Cumul1(wAmj1, wAmjK1, wAmj2, wAmjK2)
        blnOk = True
    End If
    wAmjK1 = "0"
Loop Until blnOk

End Sub

Public Sub cmdMouvement_Import_Nbj_Cumul1(lAmj1 As String, lAmjK1 As String, lAmj2 As String, lAmjK2 As String)
Dim wNbjCivils As Double, wNbjOuvrés As Double, X6 As String * 6, K As Integer, M As Integer
Dim K1 As Integer, K2 As Integer
Dim wNbj As Double

wNbjCivils = 0: wNbjOuvrés = 0
X6 = mId$(lAmj1, 1, 6)
For M = 1 To 12
    If X6 = arrDRHNbjX(M) Then Exit For
Next M

V = Param_DRHCalendrier(X6, mDRHCalendrier)
If Not IsNull(V) Then Call MsgBox(V, vbCritical, "frmDRH : cmdMouvement_Import_Nbj_Cumul1")

K1 = CInt(mId$(lAmj1, 7, 2)) * 2: If lAmjK1 = "0" Then K1 = K1 - 1
K2 = CInt(mId$(lAmj2, 7, 2)) * 2: If lAmjK2 = "1" Then K2 = K2 - 1

For K = K1 To K2
    Select Case mId$(mDRHCalendrier.Memo, K, 1)
        Case "0": wNbjCivils = wNbjCivils + 0.5: wNbjOuvrés = wNbjOuvrés + 0.5
        Case "X": wNbjCivils = wNbjCivils + 0.5
    End Select
Next K

If cumNbjC + wNbjCivils > maxNbjC Then wNbjCivils = maxNbjC - cumNbjC
cumNbjC = cumNbjC + wNbjCivils

arrDRHNbjOuvrés(arrDRH_Index, 0) = arrDRHNbjOuvrés(arrDRH_Index, 0) + wNbjOuvrés
arrDRHNbjOuvrés(arrDRH_Index, M) = arrDRHNbjOuvrés(arrDRH_Index, M) + wNbjOuvrés
arrDRHNbjCivils(arrDRH_Index, 0) = arrDRHNbjCivils(arrDRH_Index, 0) + wNbjCivils
arrDRHNbjCivils(arrDRH_Index, M) = arrDRHNbjCivils(arrDRH_Index, M) + wNbjCivils

End Sub

Public Sub cmdPrint_MvtP0()
Dim IdKey As String, mIdKey As String
mdbMvtP0.tableMvtP0_Open
Mid$(MsgTxt, 1, 40) = Space$(40)
recMvtp0.Method = "MoveFirst"

V = dbMvtP0_ReadE(recMvtp0)
IdKey = mId$(recMvtp0.Id, 1, 40): arrCompteNb = 0

Do While recMvtp0.Err = 0
    arrDRH_Index = CInt(mId$(recMvtp0.Id, 35, 5))
    
    MsgTxtIndex = 0
'    Mid$(MsgTxt, 35, memoDRHMvtLen) = mId$(recMvtp0.Text, 1, memoDRHMvtLen)
    MsgTxt = recMvtp0.Text
    If IsNull(srvDRHMvt_GetBuffer(mDRHMvt)) Then
        
            xDRH = arrDRH(arrDRH_Index)
            If prtDocument = "02" Then cmdPrint_02
            
            prtDRHMvt_Print mDRHMvt, xDRH
    End If
    recMvtp0.Method = "MoveNext    "
    recMvtp0.Err = tableMvtP0_Read(recMvtp0)
Loop

mdbMvtP0.tableMvtP0_Close

End Sub

Public Sub cmdPrint_10()
Dim marrDRH_Index As Integer

mdbMvtP0.tableMvtP0_Open
Mid$(MsgTxt, 1, 40) = Space$(40)
recMvtp0.Method = "MoveFirst"

V = dbMvtP0_ReadE(recMvtp0)
marrDRH_Index = CInt(mId$(recMvtp0.Id, 35, 5))
arrCompteNb = 0
arrDRHMvt_NB = 0: arrTotal_Init

Do While recMvtp0.Err = 0
    arrDRH_Index = CInt(mId$(recMvtp0.Id, 35, 5))
    If marrDRH_Index <> arrDRH_Index Then
        If arrDRHMvt_NB > 0 Then
            xDRH = arrDRH(marrDRH_Index)
            For I = 1 To arrDRHMvt_NB
                prtDRHMvt_Print arrDRHMvt(I), xDRH
            Next I
        End If
        marrDRH_Index = arrDRH_Index
        arrDRHMvt_NB = 0: arrTotal_Init
    End If
    
    MsgTxtIndex = 0
'    Mid$(MsgTxt, 35, memoDRHMvtLen) = mId$(recMvtp0.Text, 1, memoDRHMvtLen)
    MsgTxt = recMvtp0.Text
    If IsNull(srvDRHMvt_GetBuffer(xDRHMvt)) Then
            arrTotal_Add xDRHMvt
            arrDRHMvt_NB = arrDRHMvt_NB + 1
            arrDRHMvt(arrDRHMvt_NB) = xDRHMvt
    End If
    recMvtp0.Method = "MoveNext    "
    recMvtp0.Err = tableMvtP0_Read(recMvtp0)
Loop


If arrDRHMvt_NB > 0 Then
    xDRH = arrDRH(marrDRH_Index)
    For I = 1 To arrDRHMvt_NB
        prtDRHMvt_Print arrDRHMvt(I), xDRH
    Next I
End If

mdbMvtP0.tableMvtP0_Close

End Sub


Public Function cmdMouvement_Control_Période() As Boolean

Dim blnOk As Boolean
blnOk = True
For arrDRHMvt_Index = 1 To arrDRHMvt_NB
    If arrDRHMvt(arrDRHMvt_Index).Method <> constDelete Then
        wDRHMvt = arrDRHMvt(arrDRHMvt_Index)
        If wDRHMvt.Statut = " " And wDRHMvt.MvtSens = "-" Then
            If wDRHMvt.IdSeq <> xDRHMvt.IdSeq Then
                If xDRHMvt.DébutAmj > wDRHMvt.RepriseAmj Or xDRHMvt.RepriseAmj < wDRHMvt.DébutAmj Then
                Else
                    If xDRHMvt.DébutAmj = wDRHMvt.RepriseAmj Then
                        If xDRHMvt.DébutAmjK = "0" And wDRHMvt.RepriseAmjK = "1" Then blnOk = False: Exit Function
                    Else
                        If xDRHMvt.RepriseAmj = wDRHMvt.DébutAmj Then
                            If xDRHMvt.RepriseAmjK = "1" And wDRHMvt.DébutAmjK = "0" Then blnOk = False: Exit Function
                        Else
                            blnOk = False: Exit Function
                        End If
                    End If
                End If
            End If
            
       End If
    End If
Next arrDRHMvt_Index
cmdMouvement_Control_Période = blnOk
'                    If xDRHMvt.RepriseAmj > wDRHMvt.DébutAmj Then blnOk = False: Exit Function

End Function

Public Sub cmdPrint_03()
Dim I As Integer

For I = 1 To fgSalarié.Rows - 1
    fgSalarié.Row = I
    fgSalarié.Col = 3
    arrDRH_Index = Val(fgSalarié.Text)
    If arrDRHNbjCivils(arrDRH_Index, 0) <> 0 Then Call prtDRHMvt_Print(xDRHMvt, arrDRH(arrDRH_Index))
Next I

End Sub

Public Sub cmdPrint_05()
Dim I As Integer

For I = 1 To fgSalariéMvt.Rows - 1
    fgSalariéMvt.Row = I
    fgSalariéMvt.Col = 8
    arrDRHMvt_Index = Val(fgSalariéMvt.Text)
    prtDRHMvt_Print arrDRHMvt(arrDRHMvt_Index), mDRH
Next I
prtDRHMvt_Line05
End Sub


Public Sub paramDRHMvt_CP()
Dim X8 As String

fraExerciceCp.Enabled = DrhAut.Xspécial

recDRHMvt_Init mCP
mCP.Method = "SeekP0"
mCP.Matricule = "$Sys"
mCP.IdSeq = 1
If IsNull(srvDRHMvt_Monitor(mCP)) Then
    txtExerciceCpNbj = mCP.Nbj
    X8 = (mId$(mCP.RepriseAmj, 1, 4) + 1) & "0502"
    DTPicker_Set txtExerciceCpAmj, X8
    Exit Sub
End If

If Not DrhAut.Xspécial Then Exit Sub

X = MsgBox("Voulez-vous créer l'enregistrement ?", vbYesNo + vbQuestion + vbDefaultButton2, "paramDRHMVT_CP")
If X = vbNo Then
    Unload Me
Else
    mCP.Method = "AddNew$"
    
    mCP.MvtCode = "0101"
    mCP.DébutAmj = "1999" & "0601"
    mCP.RepriseAmj = "2000" & "0501"
    mCP.Nbj = mCP.Nbj: mCP.NbjOuvré = mCP.Nbj
    mCP.MvtSens = "C"
    mCP.MvtCO = "O"
    mCP.Statut = "$"
    mCP.UpdAmj = DSys
    mCP.UpdHms = time_Hms
    
    If Not IsNull(srvDRHMvt_Update(mCP)) Then
        Call MsgBox("erreur création [ $Sys : 1 ]", vbCritical, "paramDRHMVT_CP")
        Unload Me
    End If
End If
End Sub
Public Sub paramDRHMvt_Sys0()
recDRHMvt_Init xDRHMvt
xDRHMvt.Method = "SeekP0"
xDRHMvt.Matricule = "$Sys"
xDRHMvt.IdSeq = 0
If IsNull(srvDRHMvt_Monitor(xDRHMvt)) Then Exit Sub
If Not DrhAut.Xspécial Then Exit Sub

X = MsgBox("Voulez-vous créer l'enregistrement ?", vbYesNo + vbQuestion + vbDefaultButton2, "paramDRHMVT_Sys0")
If X = vbNo Then
    Unload Me
Else
    xDRHMvt.Method = "AddNew$"
    xDRHMvt.RéfInterne = "000000000000"
    xDRHMvt.MvtCode = "0000"
    xDRHMvt.DébutAmj = "1999" & "0101"
    xDRHMvt.RepriseAmj = "2999" & "1231"
    xDRHMvt.Nbj = 0: xDRHMvt.NbjOuvré = xDRHMvt.Nbj
    xDRHMvt.MvtSens = "C"
    xDRHMvt.MvtCO = "C"
    xDRHMvt.Statut = "$"
    xDRHMvt.UpdAmj = DSys
    xDRHMvt.UpdHms = time_Hms
    
    If Not IsNull(srvDRHMvt_Update(xDRHMvt)) Then
        Call MsgBox("erreur création [ $Sys : 0 ]", vbCritical, "paramDRHMVT_Sys0")
        Unload Me
    End If
End If
End Sub

Public Sub paramDRHMvt_Civil()
Dim X8 As String

fraExerciceCivil.Enabled = DrhAut.Xspécial

recDRHMvt_Init mCivil
mCivil.Method = "SeekP0"
mCivil.Matricule = "$Sys"
mCivil.IdSeq = 2
If IsNull(srvDRHMvt_Monitor(mCivil)) Then
    txtExerciceCivilNbj = mCivil.Nbj
    X8 = (mId$(mCivil.RepriseAmj, 1, 4) + 1) & "0102"
    DTPicker_Set txtExerciceCivilAmj, X8
    Exit Sub
End If

If Not DrhAut.Xspécial Then Exit Sub

X = MsgBox("Voulez-vous créer l'enregistrement ?", vbYesNo + vbQuestion + vbDefaultButton2, "paramDRHMVT_Civil")
If X = vbNo Then
    Unload Me
Else
    mCivil.Method = "AddNew$"
    
    mCivil.MvtCode = "0201"
    mCivil.DébutAmj = "1999" & "0101"
    mCivil.RepriseAmj = "2000" & "0101"
    mCivil.Nbj = mCivil.Nbj: mCivil.NbjOuvré = mCivil.Nbj
    mCivil.MvtSens = "C"
    mCivil.MvtCO = "O"
    mCivil.Statut = "$"
    mCivil.UpdAmj = DSys
    mCivil.UpdHms = time_Hms
    
    If Not IsNull(srvDRHMvt_Update(mCivil)) Then
        Call MsgBox("erreur création [ $Sys : 2 ]", vbCritical, "paramDRHMVT_Civil")
        Unload Me
    End If
End If

End Sub


Public Sub lstMvt_Select()
blnControl = False
lstMvt.Enabled = False
cmdMouvementHisto.Visible = False
fraMouvementDétail.Enabled = True
cmdMouvementOK.Visible = DrhAut.Valider
'chkRepriseAmj.Enabled = DrhAut.Xspécial
Call paramDRHMvt_Init(mId$(lstMvt, 1, 4))


Select Case cmdMouvementOK.Caption
    Case constAjouter
        mDRHMvt = paramDRHMvt
        mDRHMvt.Method = constAddNew
        mDRHMvt.Matricule = mDRH.Matricule
        If paramDRHMvt.MvtSens = "-" Or paramDRHMvt.MvtSens = "P" Then
            Call DTPicker_Set(txtDébutAmj, DSys)
            txtDébutAmj.Enabled = True
            'chkRepriseAmj = "0": chkRepriseAmj.Enabled = True
        Else
            Call DTPicker_Set(txtDébutAmj, paramDRHMvt.DébutAmj)
            Call DTPicker_Set(txtRepriseAmj, paramDRHMvt.RepriseAmj)
            txtDébutAmj.Enabled = False
            txtRepriseAmj.Enabled = False
            'chkRepriseAmj = "2": chkRepriseAmj.Enabled = False
        End If
        currentAction = "Mvt_AddNew"
        'jpl 2000.12.22 For I = 1 To arrDRHMvt_NB
        'jpl 2000.12.22     If mDRHMvt.IdSeq < arrDRHMvt(I).IdSeq Then mDRHMvt.IdSeq = arrDRHMvt(I).IdSeq
        'jpl 2000.12.22 Next I
        mDRHMvt.IdSeq = 99999 'jpl 2000.12.22 mDRHMvt.IdSeq + 1
    Case constModifier
        currentAction = "Mvt_Update"
        mDRHMvt.IdSeq = oldDRHMvt.IdSeq
        mDRHMvt.Method = constUpdate
        mDRHMvt.MvtCode = paramDRHMvt.MvtCode
        mDRHMvt.MvtCO = paramDRHMvt.MvtCO
        mDRHMvt.MvtSens = paramDRHMvt.MvtSens
        
        
End Select

'txtNbj.SetFocus
blnControl = True

End Sub

Public Sub fraTR_Init()
Dim X As String

recElpTable_Init xTable

recElpTable_Init xTable
xTable.Method = "Seek="
xTable.Id = "DRH"
xTable.K1 = "TR"
xTable.K2 = "Filename"
intReturn = tableElpTable_Read(xTable)
If intReturn <> 0 Then
    Call MsgBox("Paramètre manquant : DRH TR Filename", vbCritical, "paramDRH_Init")
Else
    paramTR_Filename = Trim(xTable.Memo) & mId$(DSys, 1, 6) & ".txt"
End If


xTable.K2 = "Disquette"
intReturn = tableElpTable_Read(xTable)
If intReturn <> 0 Then
    Call MsgBox("Paramètre manquant : DRH TR Disquette", vbCritical, "paramDRH_Init")
Else
    paramTR_Disquette = Trim(xTable.Memo)
End If

xTable.K2 = "Montant"
intReturn = tableElpTable_Read(xTable)
If intReturn <> 0 Then
    Call MsgBox("Paramètre manquant : DRH TR Montant", vbCritical, "paramDRH_Init")
Else
    paramTR_Nominal = mId$(xTable.Memo, 2, 4)
    paramTR_PartPatronale = mId$(xTable.Memo, 8, 4)
End If


xTable.K2 = "Id"
intReturn = tableElpTable_Read(xTable)
If intReturn <> 0 Then
    Call MsgBox("Paramètre manquant : DRH TR Id", vbCritical, "paramDRH_Init")
Else
    paramTR_Id = Trim(xTable.Memo)
End If


recDRHMvt_Init mTR
mTR.Method = "SeekP0"
mTR.Matricule = "$TR"
mTR.IdSeq = CInt(mId$(DSys, 3, 4))
If IsNull(srvDRHMvt_Monitor(mTR)) Then
    txtTRNbj = mTR.NbjOuvré
    Call DTPicker_Set(txtTRdébutAmj, mTR.DébutAmj)
    Call DTPicker_Set(txtTRFinAmj, mTR.RepriseAmj)
    mTR.Method = constUpdate
Else
    Call DTPicker_Set(txtTRFinAmj, DSys)
    wAmj = dateElp("MoisAdd", -1, DSys)
    mTR.IdSeq = CInt(mId$(wAmj, 3, 4))
    If IsNull(srvDRHMvt_Monitor(mTR)) Then
        wAmj = dateElp("Jour", 1, mTR.RepriseAmj)
    Else
        wAmj = DSys
    End If
    Call DTPicker_Set(txtTRdébutAmj, wAmj)
    mTR.Method = constAddNew
    mTR.DébutAmj = wAmj
    mTR.RepriseAmj = DSys
End If

mTR.DébutAmj = mId$(dateElp("MoisAdd", 1, DSys), 1, 6) & "01"
lblTRNbj = "Jours ouvrés en " & mId$(mTR.DébutAmj, 1, 6)
If mTR.Method = constAddNew Then
    mTR.RepriseAmj = dateElp("MoisAdd", 1, mTR.DébutAmj) ' dateFinDeMois(mTR.DébutAmj)
    mTR.MvtCO = "O"
    mTR.DébutAmjK = "0"
    mTR.RepriseAmjK = "0"

    arrDRH_Index = 0
    ReDim arrDRHNbjOuvrés(1, 12): arrDRHNbjOuvrés(0, 0) = 0
    ReDim arrDRHNbjCivils(1, 12): arrDRHNbjCivils(0, 0) = 0
    
    arrDRHNbjX(1) = mId$(mTR.DébutAmj, 1, 6)
    prtDébutAmj = "00000000"
    prtFinAmj = "99999999"
    xDRHMvt = mTR
    cmdMouvement_Import_Nbj_Cumul
    txtTRNbj = arrDRHNbjOuvrés(0, 0)
End If

X = Dir(paramTR_Filename)
If X = "" Then
    fraTR.Enabled = True
    cmdTRControl.Caption = "Etat préparatoire"
    cmdTRDisquette.Enabled = False
    cmdTrOk.Enabled = False
Else
    fraTR.Enabled = False
    cmdTRControl.Caption = constAnnuler
    cmdTRDisquette.Enabled = DrhAut.Valider
    cmdTrOk.Enabled = DrhAut.Valider
    prtDRHMvt_Nbj_Init mTR.DébutAmj
    prtDRHTR_Init
    Open paramTR_Filename For Input As #1

    Do Until EOF(1)
        Line Input #1, X
        xDRH.Matricule = "0" & mId$(X, 19, 4)
        Call arrDRH_Scan(xDRH)
        arrDRHTR(arrDRH_Index) = CDbl(mId$(X, 56, 2))
    Loop
    
    Close #1
    fgTR_Display
End If
End Sub

Public Sub cmdMouvement_Import_TR()
Dim mNbj As Double

If xDRHMvt.DébutAmj < prtDébutAmj Then xDRHMvt.DébutAmj = prtDébutAmj: xDRHMvt.DébutAmjK = "0"

If xDRHMvt.RepriseAmj > prtFinAmj1 Then xDRHMvt.RepriseAmj = prtFinAmj1: xDRHMvt.RepriseAmjK = "0"

mNbj = arrDRHNbjOuvrés(arrDRH_Index, 0)

cmdMouvement_Import_Nbj_Cumul
Mid$(recMvtp0.Text, 34 + 40, 4) = "O"
Mid$(recMvtp0.Text, 34 + 34, 4) = "0000"
Mid$(recMvtp0.Text, 34 + 53, 4) = Format$((arrDRHNbjOuvrés(arrDRH_Index, 0) - mNbj) * 10, "0000")
dbMvtP0_Update recMvtp0
End Sub

Public Sub cmdPrint_02()
Dim I As Integer, blnOk As Boolean

fgSalariéMvt_Load "SnapL1"

'Do
'blnOk = True
    For I = arrDRHMvt_NB To 1 Step -1
        If mDRHMvt.DébutAmj = arrDRHMvt(I).RepriseAmj Then
            If mDRHMvt.DébutAmj <> arrDRHMvt(I).DébutAmj Then mDRHMvt.DébutAmj = arrDRHMvt(I).DébutAmj: blnOk = False
        End If
    Next I
    For I = 1 To arrDRHMvt_NB
        If mDRHMvt.RepriseAmj = arrDRHMvt(I).DébutAmj Then
            If mDRHMvt.RepriseAmj <> arrDRHMvt(I).RepriseAmj Then mDRHMvt.RepriseAmj = arrDRHMvt(I).RepriseAmj: blnOk = False
        End If
        
    Next I
'Loop Until blnOk

End Sub

Public Sub cmdPrint_Control()
Dim X As String, V As Variant

If Not frmDRH.Enabled Then Exit Sub
frmDRH.Enabled = False

cmdPrint.Visible = False
blnControl = False

lstErr.Clear
lstErr.Height = 200

If chkPrintMouvement = "1" Then
    lstPrintMouvement.Enabled = False
Else
    lstPrintMouvement.Enabled = True
End If

If chkPrintPériode = "1" Then
    fraPrintPériode.Enabled = False
Else
    fraPrintPériode.Enabled = True
End If

If chkPrintSalarié = "1" Then
    txtPrintMatricule.Enabled = False
Else
    txtPrintMatricule.Enabled = True
End If

If chkPrintService = "1" Then
    cboPrintService.Enabled = False
Else
    cboPrintService.Enabled = True

End If

ExitSub:

If lstErr.ListCount = 0 Then
    cmdPrint.Visible = DrhAut.Saisir
End If


frmDRH.Enabled = True
    
blnControl = True


End Sub

Public Sub cmdPrintSelect_Import_Mouvement(iScan As Integer)
Dim blnOk As Boolean, iReturn As Integer

Call srvDRHMvt_ElpBuffer(xDRHMvt, mDRHMvt, xElpBuffer)
Me.MousePointer = 0
If xElpBuffer.Seq = 0 Then Exit Sub


xElpBuffer.Method = "Seek=" '
xElpBuffer.Seq = 1

Do
    iReturn = tableElpBuffer_Read(xElpBuffer)
    If iReturn = 0 Then
        blnOk = True
        cmdImport_Nb = cmdImport_Nb + 1
        MsgTxt = xElpBuffer.Data
        MsgTxtIndex = 0
        srvDRHMvt_GetBuffer xDRHMvt
        If chkPrintMouvement <> "1" Then blnOk = cmdPrintSelect_Import_Mouvement_Chk
        If blnOk And prtSelectK = "1" Then blnOk = cmdPrintSelect_Import_Mouvement_Période
        If blnOk And chkPrintCP = "1" Then blnOk = cmdPrintSelect_Import_Mouvement_Période_CP
       
        If blnOk Then
            recMvtp0.Id = Space$(40)
            Select Case prtSort
                Case 1: recMvtp0.Id = arrDRH(iScan).Service & arrDRH(iScan).Nom & arrDRH(iScan).Matricule
                Case 2: recMvtp0.Id = arrDRH(iScan).Service & arrDRH(iScan).Matricule
                Case 3: recMvtp0.Id = arrDRH(iScan).Nom & arrDRH(iScan).Matricule
                Case 4: recMvtp0.Id = arrDRH(iScan).Matricule
                Case Else: recMvtp0.Id = arrDRH(iScan).Service & arrDRH(iScan).Nom & arrDRH(iScan).Matricule
            End Select
            cmdImport_Select_Nb = cmdImport_Select_Nb + 1
            If optPrintMouvement01 Then Mid$(recMvtp0.Id, 17, 4) = xDRHMvt.MvtCode
            Mid$(recMvtp0.Id, 21, 8) = xDRHMvt.DébutAmj
            Mid$(recMvtp0.Id, 29, 1) = xDRHMvt.DébutAmjK
            Mid$(recMvtp0.Id, 30, 5) = Format$(cmdImport_Select_Nb, "00000")
            Mid$(recMvtp0.Id, 35, 5) = Format$(iScan, "00000")
            recMvtp0.Text = xElpBuffer.Data
            
            dbMvtP0_Update recMvtp0
         
      
            
            If I = 1000 Then I = 0: Call lstErr_ChangeLastItem(lstErr, cmdPrint, "Sélection : " & cmdImport_Select_Nb & " / " & cmdImport_Nb): DoEvents
        End If
        xElpBuffer.Seq = xElpBuffer.Seq + 1
   End If
Loop Until iReturn <> 0

End Sub

Public Function cmdPrintSelect_Import_Mouvement_Chk() As Boolean
Dim I As Integer
cmdPrintSelect_Import_Mouvement_Chk = False
For I = 0 To arrPrintMouvement_Nb
    If xDRHMvt.MvtCode = arrPrintMouvement(I) Then
        cmdPrintSelect_Import_Mouvement_Chk = True
        Exit For
    End If
Next I

End Function

Public Function cmdPrintSelect_Import_Mouvement_Période() As Boolean
Dim I As Integer
If xDRHMvt.DébutAmj > mprtFinAMJ Then
    cmdPrintSelect_Import_Mouvement_Période = False
Else
    If xDRHMvt.RepriseAmj <= mprtDébutAMJ Then
        cmdPrintSelect_Import_Mouvement_Période = False
    Else
        cmdPrintSelect_Import_Mouvement_Période = True
    End If
End If

End Function

Public Function cmdPrintSelect_Import_Mouvement_Période_CP() As Boolean
Dim I As Integer
If xDRHMvt.DébutAmj > mprtFinAMJ Then
    cmdPrintSelect_Import_Mouvement_Période_CP = False
Else
    If xDRHMvt.DébutAmj < mprtDébutAMJ Then
        cmdPrintSelect_Import_Mouvement_Période_CP = False
    Else
        cmdPrintSelect_Import_Mouvement_Période_CP = True
    End If
End If

End Function

Public Sub cmdPrintSelect_Import_Mouvement_Init()
Dim I As Integer

For I = 0 To lstPrintMouvement.ListCount - 1
    If lstPrintMouvement.Selected(I) Then
        arrPrintMouvement_Nb = arrPrintMouvement_Nb + 1
        lstPrintMouvement.ListIndex = I
        arrPrintMouvement(arrPrintMouvement_Nb) = mId$(lstPrintMouvement, 1, 4)
    End If
Next I

End Sub


