VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmGarantie 
   AutoRedraw      =   -1  'True
   Caption         =   "Garanties"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6765
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0FF&
      Caption         =   "en &Attente"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
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
      Height          =   500
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   6120
      TabIndex        =   21
      Top             =   0
      Width           =   2745
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Picture         =   "Garantie.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   0
      TabIndex        =   29
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   706
      TabCaption(0)   =   "Liste des dossiers"
      TabPicture(0)   =   "Garantie.frx":0102
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fgSelect"
      Tab(0).Control(1)=   "fraOption"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Garantie (1/3)"
      TabPicture(1)   =   "Garantie.frx":011E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraGarantie"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Bénéficiaire (2/3)"
      TabPicture(2)   =   "Garantie.frx":013A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraBénéficaire"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Commissions (3/3)"
      TabPicture(3)   =   "Garantie.frx":0156
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraNature"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Mouvements / Echéancier"
      TabPicture(4)   =   "Garantie.frx":0172
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fgEchéancier"
      Tab(4).ControlCount=   1
      Begin VB.Frame fraOption 
         Caption         =   "Options"
         Height          =   4455
         Left            =   -70440
         TabIndex        =   58
         Top             =   960
         Width           =   4575
         Begin VB.TextBox txtSelect 
            Height          =   285
            Left            =   2520
            TabIndex        =   59
            Top             =   480
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker txtAmjMax 
            Height          =   300
            Left            =   2520
            TabIndex        =   62
            Top             =   1440
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
            Format          =   65863683
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblSelect2 
            Caption         =   "Echéancier jusqu'au"
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
            Left            =   120
            TabIndex        =   61
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label lblSelect1 
            Caption         =   "Référence interne"
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
            Left            =   120
            TabIndex        =   60
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame fraBénéficaire 
         Caption         =   "Bénéficiaire"
         Height          =   4935
         Left            =   -74760
         TabIndex        =   33
         Top             =   840
         Width           =   8895
      End
      Begin VB.Frame fraNature 
         Caption         =   "Commissions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   -74880
         TabIndex        =   32
         Top             =   600
         Width           =   9135
         Begin VB.Frame fraCommissionPériodique 
            Caption         =   "Commission périodique"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   2400
            TabIndex        =   37
            Top             =   1560
            Width           =   6615
            Begin VB.Frame fraComPériodiqueTaux 
               Height          =   975
               Left            =   120
               TabIndex        =   63
               Top             =   240
               Width           =   4095
               Begin VB.OptionButton optComPériodiqueMontant 
                  Caption         =   "montant fixe"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   66
                  Top             =   600
                  Width           =   1935
               End
               Begin VB.OptionButton optComPériodiqueTaux 
                  Caption         =   "taux annuel"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   65
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1815
               End
               Begin VB.TextBox txtTaux 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2400
                  TabIndex        =   64
                  Top             =   360
                  Width           =   1215
               End
            End
            Begin VB.OptionButton optEchéanceFinDeMois 
               Caption         =   "Fin de mois"
               Height          =   195
               Left            =   2520
               TabIndex        =   15
               Top             =   2160
               Width           =   1455
            End
            Begin VB.OptionButton optEchéanceAnniversaire 
               Caption         =   "Anniversaire"
               Height          =   195
               Left            =   2520
               TabIndex        =   14
               Top             =   1800
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.CheckBox chkAmjEchéance 
               Caption         =   "Première échéance"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   1440
               Width           =   1695
            End
            Begin VB.Frame fraPériodicité 
               Caption         =   "Périodicité"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2100
               Left            =   4560
               TabIndex        =   38
               Top             =   240
               Width           =   1935
               Begin VB.OptionButton optAnnuel 
                  Caption         =   "Annuelle"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   27
                  Top             =   1440
                  Width           =   1500
               End
               Begin VB.OptionButton optSemestriel 
                  Caption         =   "Semestrielle"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   26
                  Top             =   1080
                  Width           =   1500
               End
               Begin VB.OptionButton optTrimestriel 
                  Caption         =   "Trimestrielle"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   25
                  Top             =   720
                  Value           =   -1  'True
                  Width           =   1500
               End
               Begin VB.OptionButton optMensuel 
                  Caption         =   "Mensuelle"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   24
                  Top             =   300
                  Width           =   1500
               End
            End
            Begin MSComCtl2.DTPicker txtAmjEchéance 
               Height          =   300
               Left            =   2520
               TabIndex        =   13
               Top             =   1440
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
               Format          =   65863683
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblEchéanceSuivantes 
               Caption         =   "échéances suivantes  à date :"
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
               Left            =   120
               TabIndex        =   39
               Top             =   1800
               Width           =   2175
            End
         End
         Begin VB.Frame fraCommissionàRéclamer 
            Caption         =   "Commission à Réclamer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   2400
            TabIndex        =   34
            Top             =   240
            Width           =   6615
            Begin VB.TextBox txtEchéanceCompte 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2280
               TabIndex        =   11
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label libEchéanceTypeDeCompte 
               Caption         =   "-"
               Height          =   375
               Left            =   3960
               TabIndex        =   40
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label libEchéanceCompte 
               Caption         =   "-"
               Height          =   255
               Left            =   2280
               TabIndex        =   36
               Top             =   720
               Width           =   4215
            End
            Begin VB.Label lblEchCompte 
               Caption         =   "Compte à débiter"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.CheckBox chkCommissionFlat 
            Caption         =   "Commission flat"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   4560
            Width           =   1455
         End
         Begin VB.TextBox txtCommissionFlat 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2400
            TabIndex        =   16
            Top             =   4560
            Width           =   1575
         End
         Begin VB.CheckBox chkCommissionàRéclamer 
            Caption         =   "Commission à réclamer"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   1935
         End
         Begin VB.CheckBox chkCommissionPériodique 
            Caption         =   "commission périodique"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Value           =   1  'Checked
            Width           =   2175
         End
      End
      Begin VB.Frame fraGarantie 
         Caption         =   "Caractéristiques de la garantie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   120
         TabIndex        =   30
         Top             =   430
         Width           =   9135
         Begin VB.TextBox txtPréavisNbj 
            Height          =   285
            Left            =   1920
            TabIndex        =   8
            Text            =   "16"
            Top             =   4080
            Width           =   495
         End
         Begin VB.CheckBox chkMainLevée 
            Alignment       =   1  'Right Justify
            Caption         =   "attendre main levée"
            Height          =   255
            Left            =   3720
            TabIndex        =   56
            Top             =   3600
            Width           =   1935
         End
         Begin VB.CheckBox chkComptaReprise 
            Alignment       =   1  'Right Justify
            Caption         =   "Reprise de l'en-cours"
            Height          =   255
            Left            =   3720
            TabIndex        =   52
            Top             =   3120
            Width           =   1935
         End
         Begin VB.ComboBox cboNature 
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   360
            Width           =   5055
         End
         Begin VB.TextBox txtDonneurDordre 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1920
            MaxLength       =   5
            TabIndex        =   4
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox txtEngagementCompte 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1920
            TabIndex        =   5
            Top             =   2400
            Width           =   1575
         End
         Begin VB.TextBox txtDevise 
            Height          =   285
            Left            =   1920
            TabIndex        =   2
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtCapital 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1920
            TabIndex        =   3
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtRéférenceInterne 
            Height          =   285
            Left            =   1920
            MaxLength       =   16
            TabIndex        =   9
            Top             =   4680
            Width           =   2655
         End
         Begin VB.TextBox txtRéférenceExterne 
            Height          =   285
            Left            =   1920
            MaxLength       =   16
            TabIndex        =   10
            Top             =   5160
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker txtAMJFin 
            Height          =   300
            Left            =   1920
            TabIndex        =   7
            Top             =   3480
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
            Format          =   65863683
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin MSComCtl2.DTPicker txtAmjEngagement 
            Height          =   300
            Left            =   1920
            TabIndex        =   6
            Top             =   3000
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
            Format          =   65863683
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin MSComCtl2.DTPicker txtAMJEffet 
            Height          =   300
            Left            =   5760
            TabIndex        =   67
            Top             =   1200
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
            Format          =   65863683
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblAMJEffet 
            Caption         =   "Date d'effet"
            Height          =   375
            Left            =   3720
            TabIndex        =   68
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblPréavisNbj 
            Caption         =   "délai courrier (jours) Echéancier ajouter                  1 jour"
            Height          =   615
            Left            =   120
            TabIndex        =   57
            Top             =   4080
            Width           =   1575
         End
         Begin VB.Label libDevise 
            Caption         =   "--"
            Height          =   255
            Left            =   3600
            TabIndex        =   55
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label libEngagementCompte 
            Caption         =   "-"
            Height          =   255
            Left            =   3720
            TabIndex        =   54
            Top             =   2400
            Width           =   4695
         End
         Begin VB.Label libDonneurDordre 
            Caption         =   "-"
            Height          =   255
            Left            =   3720
            TabIndex        =   53
            Top             =   1920
            Width           =   4695
         End
         Begin VB.Label lblNature 
            Caption         =   "Nature"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblEngCompte 
            Caption         =   "Donneur d'ordre"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblEngagementCompte 
            Caption         =   "compte de garantie"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label lblDevise 
            Caption         =   "Devise"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblCapital 
            Caption         =   "Montant"
            Height          =   315
            Left            =   120
            TabIndex        =   47
            Top             =   1200
            Width           =   1665
         End
         Begin VB.Label lblAmjEngagement 
            Caption         =   "Date d'émission"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label lblAmjFin 
            Caption         =   "Date de validité"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   3480
            Width           =   1695
         End
         Begin VB.Label lblRéférenceInterne 
            Caption         =   "Référence interne"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   4800
            Width           =   1575
         End
         Begin VB.Label lblRéférenceExterne 
            Caption         =   "Réf contrat commercial"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   5160
            Width           =   1815
         End
         Begin VB.Label libStatut 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   6120
            TabIndex        =   42
            Top             =   3600
            Width           =   2775
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   5250
         Left            =   -74880
         TabIndex        =   0
         Top             =   600
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   9260
         _Version        =   393216
         Rows            =   1
         Cols            =   17
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
         FormatString    =   $"Garantie.frx":018E
      End
      Begin MSFlexGridLib.MSFlexGrid fgEchéancier 
         Height          =   5250
         Left            =   -74880
         TabIndex        =   31
         Top             =   600
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   9260
         _Version        =   393216
         Rows            =   1
         Cols            =   12
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
         FormatString    =   $"Garantie.frx":0328
      End
   End
   Begin VB.Label libRéférenceInterne 
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
      Height          =   405
      Left            =   3600
      TabIndex        =   41
      Top             =   0
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuGarantieSaisir 
         Caption         =   "Saisir une garantie"
      End
      Begin VB.Menu mnuContextX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListàValider 
         Caption         =   "Liste des garanties à valider"
      End
      Begin VB.Menu mnuListGarantie 
         Caption         =   "Liste des garanties"
      End
      Begin VB.Menu mnuListEchéancier 
         Caption         =   "Echéancier (mvts automatiques)"
      End
      Begin VB.Menu mnuListEchéancierManuel 
         Caption         =   "Echéancier (mvts manuels)"
      End
      Begin VB.Menu mnuContextX3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComptaEchéancier 
         Caption         =   "Compta : Echéances à comptabiliser = Jour"
      End
      Begin VB.Menu mnuComptaLotsàValider 
         Caption         =   "Compta : Lots à valider"
      End
      Begin VB.Menu mnuLotComptabilisé_Annuler 
         Caption         =   "Compta : annuler un lot comptabilisé"
      End
      Begin VB.Menu mnuContextX2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuGarantie 
      Caption         =   "mnuGarantie"
      Visible         =   0   'False
      Begin VB.Menu mnuContextX4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGarantieDisplay 
         Caption         =   "Afficher cette garantie"
      End
      Begin VB.Menu mnuGarantieModifier 
         Caption         =   "Modifier cette garantie"
      End
      Begin VB.Menu mnuGarantieValider 
         Caption         =   "Valider/ Invalider cette garantie"
      End
      Begin VB.Menu mnuGarantieAnnuler 
         Caption         =   "Annuler cette garantie"
      End
      Begin VB.Menu mnuGarantieAMJFin 
         Caption         =   "Modifier la date de validité"
      End
      Begin VB.Menu mnuGarantieMainLevéePartielle 
         Caption         =   "Main levée partielle"
      End
      Begin VB.Menu mnuGarantieMainLevée 
         Caption         =   "Main levée "
      End
      Begin VB.Menu mnuGarantieAugmentation 
         Caption         =   "Augmentation"
      End
      Begin VB.Menu mnuContextX5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGarantieEffacer 
         Caption         =   "Effacer cette garantie"
      End
      Begin VB.Menu mnuGarantiePrintList 
         Caption         =   "Imprimer la liste sélectionnée"
      End
   End
   Begin VB.Menu mnucmdPrint 
      Caption         =   "mnucmdPrint"
      Visible         =   0   'False
      Begin VB.Menu mnucmdPrintGarantie 
         Caption         =   "Imprimer la garantie sélectionné"
      End
      Begin VB.Menu mnucmdPrintList_TOpe 
         Caption         =   "Imprimer la liste des garanties sélectionnées"
      End
      Begin VB.Menu mnucmdPrintList_TFlux 
         Caption         =   "Imprimer l'échéancier"
      End
   End
   Begin VB.Menu mnuLot 
      Caption         =   "mnuLot"
      Visible         =   0   'False
      Begin VB.Menu mnuLotàComptaValidation 
         Caption         =   "Lot : comptabilisation définitive"
      End
      Begin VB.Menu mnuLotàComptaAnnulation 
         Caption         =   "Lot : annuler la demande de comptabilisation"
      End
      Begin VB.Menu mnuLotàComptaPrint 
         Caption         =   "Lot : imprimer la demande de comptabilisation"
      End
   End
   Begin VB.Menu mnuEchéancier 
      Caption         =   "mnuEchéancier"
      Visible         =   0   'False
      Begin VB.Menu mnuEchéancierAvis 
         Caption         =   "Imprimer avis"
      End
      Begin VB.Menu mnuEchéancierACU 
         Caption         =   "Annuler la comptabilisation"
      End
      Begin VB.Menu mnuEchéancierEnCours 
         Caption         =   "restaurer la comptabilisation automatique"
      End
      Begin VB.Menu mnuEchéancierManuel 
         Caption         =   "gestion manuelle"
      End
   End
End
Attribute VB_Name = "frmGarantie"
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
Dim GarantieAut As typeAuthorization

Dim recTable As typeElpTable
Dim wAmjEngagement As String, wAmjEchéance As String, blnAmjEchéance As Boolean
Dim wAmjDébut  As String, wAmjFin As String
Dim paramAmjEngagementMin As String, paramAmjEngagementMax As String
Dim paramAmjEchéanceMin As String, paramAmjEchéanceMax As String
Dim wAMJEffet  As String

Dim fgEchéancier_FormatString As String, fgEchéancier_K As Integer
Dim fgEchéancier_RowDisplay As Integer, fgEchéancier_RowClick As Integer
Dim fgEchéancier_ColorClick As Long, fgEchéancier_ColorDisplay As Long
Dim fgEchéancier_Sort1 As Integer, fgEchéancier_Sort2 As Integer
Dim fgEchéancier_SortAD As Integer, fgEchéancier_Sort1_Old As Integer

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer

Dim CV1 As typeCV
Dim recTope As typeTOpe, xTOpe As typeTOpe, mTope As typeTOpe, mEchéancierTope As typeTOpe
Dim arrTFlux() As typeTFlux, recTFlux As typeTFlux, mTflux As typeTFlux
Dim arrTFlux_Nb As Integer, arrTFlux_Index As Integer, arrTFlux_NbMax As Integer
Dim saveTFlux() As typeTFlux, saveTFlux_Index As Integer, saveTFlux_Nb As Integer
Dim saveTflux_Index_GA02 As Integer

Dim totalCapital As Currency, totalIntérêts As Currency
Dim recCompte As typeCompte
Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnfgSelect_DisplayLine As Boolean, blnfgEchéance_DisplayLine As Boolean

Dim fctPériodicité As String
Dim mAmjMin As String, mAmjMax As String, mAMJReprise As String

Dim mEChéanceCompte As String, mEngagementCompte As String, mEngagementCorrCompte As String
Dim wAmjEchéanceTrt As String * 8, wAmjExtourne As String * 8, minAmjExtourne As String * 8
Dim mNature As String, mDonneurDordre As String
Dim mCboNature  As String * 5

Dim paramTFlux_CompteCommissionàRéclamer As String * 11
Dim blnControlBiatyp As Boolean, blnComptaAuto As Boolean
Dim blnEchéancier_Gen As Boolean

Public Sub fgSelect_Load()
Dim X As String, mMethod As String

recTOpe_Init xTOpe

Select Case currentAction
    Case "mnuListàValider"
            xTOpe.Method = "SnapLS"
            xTOpe.Application = paramTFlux_Service
            xTOpe.IdRéférence = "0000000000"
            xTOpe.Statut = "à"
            
            arrTOpe(0) = xTOpe
            arrTOpe(0).IdRéférence = "999999999"

    Case "mnuListGarantie"
            xTOpe.Method = "SnapLRI"
            X = Trim(txtSelect)
            xTOpe.RéférenceInterne = X
            If X = "" Then xTOpe.Method = "SnapLS"
            xTOpe.Application = paramTFlux_Service
            xTOpe.IdRéférence = 0
            xTOpe.Statut = " "
            
            arrTOpe(0) = xTOpe
            arrTOpe(0).IdRéférence = 999999999
            arrTOpe(0).RéférenceInterne = X & "9z"

End Select

mMethod = Trim(xTOpe.Method)
arrTOpe_NBMax = 0
arrTOpe_Suite = True: arrTOpe_NB = 0
Do Until Not arrTOpe_Suite
    srvTOpe_Monitor xTOpe
    xTOpe = arrTOpe(arrTOpe_NB)
    xTOpe.Method = mMethod & "+"
Loop
fgSelect_Display
End Sub
Private Sub fgSelect_Display()
Dim K2 As Integer, I As Integer
Dim curDB As Currency, curCR As Currency, curX As Currency

SSTab1.Tab = 0

fgSelect.Visible = True
fgSelect.Clear: fgSelect.Rows = 1: fgSelect_RowDisplay = 0

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Enabled = True
For arrTOpe_Index = 1 To arrTOpe_NB
    If arrTOpe(arrTOpe_Index).Method <> constIgnore And arrTOpe(arrTOpe_Index).Method <> constDelete Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine
    End If
Next arrTOpe_Index

fgSelect_SortAD = 5
If fgSelect.Rows = 1 Then Exit Sub
'fgSelect_Sort

End Sub
Public Sub fgSelect_Sort()
If fgSelect.Rows > 1 Then
    fgSelect.Row = 1
    fgSelect.RowSel = fgSelect.Rows - 1
    
    If fgSelect_Sort1_Old = fgSelect_Sort1 Then
        If fgSelect_SortAD = 5 Then
            fgSelect_SortAD = 6
        Else
            fgSelect_SortAD = 5
        End If
    Else
        fgSelect_SortAD = 5
    End If
    fgSelect_Sort1_Old = fgSelect_Sort1
    
    fgSelect.Col = fgSelect_Sort1
    fgSelect.ColSel = fgSelect_Sort2
    fgSelect.Sort = fgSelect_SortAD
End If

End Sub


Public Sub fgEchéancier_Sort()
If fgEchéancier.Rows > 1 Then
    fgEchéancier.Row = 1
    fgEchéancier.RowSel = fgEchéancier.Rows - 1
    If fgEchéancier_Sort1_Old = fgEchéancier_Sort1 Then
        If fgEchéancier_SortAD = 5 Then
            fgEchéancier_SortAD = 6
        Else
            fgEchéancier_SortAD = 5
        End If
    Else
        fgEchéancier_SortAD = 5
    End If
    fgEchéancier_Sort1_Old = fgEchéancier_Sort1
    
    fgEchéancier.Col = fgEchéancier_Sort1
    fgEchéancier.ColSel = fgEchéancier_Sort2
    fgEchéancier.Sort = fgEchéancier_SortAD
End If
    

End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
blnControl = False
If fraOption.Visible Then
    fraOption.Visible = False
Else
    If currentAction <> "" Then
        currentAction = ""
        cmdContext.Caption = constcmdRechercher
        fgSelect.Enabled = True
        fgEchéancier.Enabled = True
        fraGarantie.Enabled = False
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
If fraOption.Visible Then
    fraOption.Visible = False
Else
    If SSTab1.Tab = 0 And Trim(txtSelect) <> "" Then
        mnuListGarantie_Click
    Else
        SendKeys "{TAB}"
    End If
End If

End Sub


'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
cmdOk.Caption = constàValider: cmdOk.Visible = False
cmdSave.Caption = constEnAttente: cmdSave.Visible = False
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
libRéférenceInterne = ""
lstErr.Visible = False
blnComptaAuto = False
cmdOk.FontSize = 8: cmdOk.FontName = "MS Sans Serif"
blncmdOk_Visible = False: blncmdSave_Visible = False
blnfgSelect_DisplayLine = False: blnfgEchéance_DisplayLine = False

fraGarantie.Enabled = False
fgEchéancier.Clear: fgEchéancier.Rows = 1: fgEchéancier_RowDisplay = 0
If cboNature.ListCount > -1 Then cboNature.ListIndex = 0
CV1 = CV_Euro
CV1.DeviseIso = "FRF"
CV_Attribut CV1
txtDevise = CV1.DeviseIso
txtCapital = ""
txtTaux = ""
txtCommissionFlat = ""
optTrimestriel = True
chkCommissionPériodique = "1"
fraCommissionPériodique.Enabled = True
optComPériodiqueTaux = True

lblAMJEffet.Visible = False: txtAMJEffet.Visible = False
Call lbl_Style(lblCapital, False)
Call lbl_Style(lblAMJEffet, False)
Call lbl_Style(lblAmjFin, False)
Call lbl_Style(lblPréavisNbj, False)
Call chk_Style(chkMainLevée, False)


mAMJReprise = DSys
wAmjEngagement = DSys: Call DTPicker_Set(txtAmjEngagement, wAmjEngagement)
wAmjFin = DSys: Call DTPicker_Set(txtAmjEngagement, wAmjFin)
wAmjEchéance = dateFinDeMois(dateElp("MoisAdd", 1, DSys)): Call DTPicker_Set(txtAMJEchéance, wAmjEchéance)
txtDonneurDordre = ""
recRacineInit C_Racine
mDonneurDordre = "": mCboNature = ""
txtRéférenceInterne = ""
txtRéférenceExterne = ""
txtEngagementCompte = "": libEngagementCompte = ""
txtEchéanceCompte = "": libEchéanceCompte = ""
optEchéanceAnniversaire = True
recTOpe_Init mTope
mTope.Statut = "à"
mTope.StatutPlus = "?"
mTope.Method = constAddNew
mEChéanceCompte = Space$(11): mEngagementCompte = Space$(11): mEngagementCorrCompte = Space$(11)
Call DTPicker_Set(txtAmjEngagement, DSys)
Call DTPicker_Set(txtAMJFin, DSys)
chkComptaReprise = "0"
chkMainLevée = "0"
chkComptaReprise = "0"

fraOption.Visible = False
blnEchéancier_Gen = False
recTOpe_Init mEchéancierTope: mEchéancierTope.Application = paramTFlux_Service
saveTflux_Index_GA02 = 0
blnControl = True
End Sub



Public Sub Form_Init()

TFlux_Compta.param_Init mNature, cboNature
libRéférenceInterne.ForeColor = vbBlue

recElpTable_Init recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "TOPE" '"Param"
recElpTable.K1 = mNature
recElpTable.K2 = "ComàRéclamer"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramTFlux_CompteCommissionàRéclamer = mId$(recElpTable.Memo, 1, 11)
If Not IsNumeric(paramTFlux_Service) Then GoTo Num_Error

wAmjEchéanceTrt = dateElp("Jour", 15, DSys)
Call DTPicker_Set(txtAmjMax, wAmjEchéanceTrt)
SSTab1.Tab = 0
tableElpTable_Open
paramAmjEngagementMin = paramAmjOpérationMin   ' jpl 2000-09-01 mId$(DSys, 1, 6) & "01"
paramAmjEngagementMax = dateElp("Ouvré", 7, DSys)
ReDim arrTOpe(1)
cmdReset
mnuGarantieSaisir.Enabled = GarantieAut.Saisir
mnuListàValider.Enabled = GarantieAut.Consulter
mnuListGarantie.Enabled = GarantieAut.Consulter
mnuComptaEchéancier.Enabled = GarantieAut.Consulter
mnuListEchéancier.Enabled = GarantieAut.Consulter
mnuComptaLotsàValider.Enabled = GarantieAut.Comptabiliser
mnuLotComptabilisé_Annuler.Enabled = GarantieAut.Xspécial
blnControl = False
''txtSelect.SetFocus
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = 0
fgSelect_FormatString = fgSelect.FormatString
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0

fgEchéancier_Sort1 = 11: fgEchéancier_Sort2 = 11
fgEchéancier_Sort1_Old = 11
fgEchéancier_FormatString = fgEchéancier.FormatString
fgEchéancier_RowDisplay = 0: fgEchéancier_RowClick = 0

Exit Sub

Table_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Table", vbCritical, "frmGarantie.Form_Init"
Exit Sub

Memo_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "frmGarantie.Form_Init"
Exit Sub

Num_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "TfluxEspèces_Param_Init"
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

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate
mNature = "DAFIG_": Call BiaPgmAut_Init("DAFI_Garant", GarantieAut)
mnuGarantieSaisir.Enabled = GarantieAut.Saisir

Form_Init
If UCase$(Trim(mId$(Msg, 13, 12))) = "BIA_EXPLOIT" Then
    mnuListEchéancier_Click
    mnucmdPrintList_TFlux_Click
    Unload Me
End If

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

'-------------------------------------------------'
Private Sub txtAmjEchéance_control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAMJEchéance.Year, "0000") & Format$(txtAMJEchéance.Month, "00") & Format$(txtAMJEchéance.Day, "00")
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, cmdContext, "? erreur date")
    DTPicker_Now txtAMJEchéance
Else
    wAmjEchéance = mId$(X, 1, 8)
End If

End Sub

'-------------------------------------------------'
Private Sub txtAmjEngagement_control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAmjEngagement.Year, "0000") & Format$(txtAmjEngagement.Month, "00") & Format$(txtAmjEngagement.Day, "00")
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, cmdContext, "? erreur date")
    DTPicker_Now txtAmjEngagement
Else
    wAmjEngagement = mId$(X, 1, 8)
    
End If

End Sub

'-------------------------------------------------'
Private Sub txtAmjfin_control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAMJFin.Year, "0000") & Format$(txtAMJFin.Month, "00") & Format$(txtAMJFin.Day, "00")
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, cmdContext, "? erreur date")
    DTPicker_Now txtAMJFin
Else
    wAmjFin = mId$(X, 1, 8)
    
End If

End Sub

'-------------------------------------------------'
Private Sub txtAMJEffet_control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAMJEffet.Year, "0000") & Format$(txtAMJEffet.Month, "00") & Format$(txtAMJEffet.Day, "00")
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, cmdContext, "? erreur date")
    DTPicker_Now txtAMJEffet
Else
    wAMJEffet = mId$(X, 1, 8)
    
End If

End Sub


Private Sub cboNature_Click()
cbo_Value recTope.Nature, cboNature
If blnControl Then cmdControl

End Sub


Private Sub cboNature_GotFocus()
lblNature.ForeColor = warnUsrColor
End Sub


Private Sub cboNature_LostFocus()
lblNature.ForeColor = lblUsr.ForeColor
If blnControl Then cmdControl
End Sub


Private Sub chkAmjEchéance_Click()
If blnControl Then cmdControl

End Sub

Private Sub chkCommissionàRéclamer_Click()
txtEchéanceCompte = ""
If blnControl Then cmdControl

End Sub


Private Sub chkCommissionFlat_Click()
If blnControl Then cmdControl

End Sub


Private Sub chkCommissionPériodique_Click()
If blnControl Then cmdControl

End Sub


Private Sub chkComptaReprise_Click()
If blnControl Then cmdControl

End Sub

Private Sub chkComptaReprise_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkComptaReprise

End Sub


Private Sub chkMainLevée_Click()
If blnControl Then cmdControl
End Sub

Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub


Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext
End Sub

Private Sub cmdOk_Click()
Dim blnPrint As Boolean, wPrint_Msg As String
Dim V

blnPrint = False
wPrint_Msg = "Garantie"

If cmdOk.Caption = constàCompta Then
    cmdSave_àCompta
Else
    cmdControl
    If cmdOk.Caption = "FinAuto" Then lstErr.Clear
    If lstErr.ListCount <> 0 Then Exit Sub
    frmGarantie.Enabled = False
    Select Case cmdOk.Caption
        Case constàValider
            recTope.Statut = "à"
            recTope.StatutPlus = "V "
            recTope.MajAMJ = DSys
            recTope.MajHMS = time_Hms
            recTope.MajUsr = usrId
            wPrint_Msg = constàValider 'cmdPrint_Call constàValider
            blnPrint = True
        Case constValider
            If Not GarantieAut.Xspécial And Trim(recTope.MajUsr) = Trim(usrId) Then
                Call MsgBox("Vous ne pouvez pas valider vos propres opérations.", vbCritical, "Garantie : Validation ")
                Call lstErr_AddItem(lstErr, cmdContext, "? validation interdite")
            Else
                fgEchéancier_AddNew
                recTope.Statut = " "
                recTope.StatutPlus = "  "
                recTope.valAMJ = DSys
                recTope.ValHMS = time_Hms
                recTope.ValUsr = usrId
                blnComptaAuto = GarantieAut.Comptabiliser
            End If
        Case "OK_Validité"
            saveTflux_AddNew
            fgEchéancier_Update
            blnPrint = True
         Case "OK_MainLevéePartielle", "OK_MainLevée", "OK_Augmentation"
            saveTflux_AddNew
            fgEchéancier_Update
            fraGarantie_Load " "
            blnComptaAuto = False '$JPL20001025 GarantieAut.Comptabiliser
            blnPrint = True
         Case "FinAuto"
            recTope.Method = constUpdate
            recTope.Statut = "F"
            recTope.StatutPlus = "in"
   Case Else
            Call lstErr_AddItem(lstErr, cmdContext, "? cmdOk : " & cmdOk.Caption)
    End Select

    If lstErr.ListCount = 0 Then
        V = cmdSave_Db
        If Not IsNull(V) Then blnPrint = False
    End If
    
    frmGarantie.Enabled = True
    AppActivate frmGarantie.Caption
End If
If blnPrint Then
    fraGarantie_Load " "
    cmdPrint_Call wPrint_Msg
End If
End Sub

Private Sub cmdPrint_Click()
Me.PopupMenu mnucmdPrint, vbPopupMenuLeftButton
End Sub

Private Sub cmdSave_Click()
cmdControl
lstErr.Clear
frmGarantie.Enabled = False
Select Case cmdSave.Caption
    Case constEnAttente
        recTope.MajAMJ = DSys
        recTope.MajHMS = time_Hms
        recTope.MajUsr = usrId
    Case constàModifier
 ''''       fgEchéancier_AddNew
        recTope.Statut = "à"
        recTope.StatutPlus = "? "
        recTope.valAMJ = DSys
        recTope.ValHMS = time_Hms
        recTope.ValUsr = constàModifier
    Case constEffacer
        recTope.Method = constDelete
    Case Else
        Call lstErr_AddItem(lstErr, cmdContext, "? cmdsave : " & cmdSave.Caption)
End Select

If lstErr.ListCount = 0 Then cmdSave_Db
frmGarantie.Enabled = True
End Sub
Public Function cmdSave_Db()
If lstErr.ListCount = 0 Then
    blnControl = False
    V = srvTope_Update(recTope)
    cmdSave_Db = V
    xTOpe = recTope
    
    If IsNull(V) Then
        If blnfgSelect_DisplayLine Then
            arrTOpe(arrTOpe_Index) = recTope
            If recTope.Method = constDelete Then
                fgSelect_Display
            Else
                fgSelect_DisplayLine
            End If
        End If
        lastActiveControl_Name = ""
        cmdOk.Visible = False
        cmdSave.Visible = False
        Call lstErr_Clear(lstErr, cmdContext, "Mise à jour effectuée - Identification : " & recTope.IdRéférence)
        If blnComptaAuto Then mnuComptaDossier
        cmdContext_Quit
        SSTab1.Tab = 0
    Else
        Call lstErr_Clear(lstErr, cmdContext, V)
 ''''       cmdReset
    End If
End If
End Function

Private Sub fgEchéancier_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y <= fgEchéancier.RowHeightMin Then
    Select Case fgEchéancier.Col
        Case 0: fgEchéancier_Sort1 = 0: fgEchéancier_Sort2 = 0: fgEchéancier_Sort
        Case 1: fgEchéancier_Sort1 = 1: fgEchéancier_Sort2 = 1: fgEchéancier_Sort
        Case 2:  fgEchéancier_SortX 2
        Case 3:  fgEchéancier_SortX 3
        Case 3: fgEchéancier_Sort1 = 3: fgEchéancier_Sort2 = 3: fgEchéancier_Sort
        Case 6: fgEchéancier_Sort1 = 6: fgEchéancier_Sort2 = 6: fgEchéancier_Sort
        Case 7: fgEchéancier_Sort1 = 7: fgEchéancier_Sort2 = 7: fgEchéancier_Sort
        Case 8: fgEchéancier_Sort1 = 8: fgEchéancier_Sort2 = 8: fgEchéancier_Sort
        Case 9:  fgEchéancier_SortX 9
        Case 11: fgEchéancier_SortX 11
    End Select
Else
    fgEchéancier_K = fgEchéancier.Row * fgEchéancier.Cols
    If fgEchéancier.Rows > 1 Then
        Call fgEchéancier_Color(fgEchéancier_RowClick, MouseMoveUsr.BackColor, fgEchéancier_ColorClick)
        arrTFlux_Index = Val(fgEchéancier.TextArray(fgEchéancier.Cols - 1 + fgEchéancier_K))
        '''recTFlux.CptMvtLot = arrTFlux(arrTFlux_Index).CptMvtLot
        recTFlux = arrTFlux(arrTFlux_Index)
        
        If currentAction = constDisplay Then
            Param_CodeOpération recTFlux.CodeOpération
            mnuEchéancier_Set
            Me.PopupMenu mnuEchéancier, vbPopupMenuLeftButton
           Else
            If recTFlux.CptMvtLot > 0 Then
                mnuLotàComptaValidation = False
                mnuLotàComptaAnnulation = False
                mnuLotàComptaAnnulation = False
              
                If recTFlux.Statut = "à" And recTFlux.StatutPlus = "C " Then
                    mnuLotàComptaValidation = GarantieAut.Comptabiliser
                    mnuLotàComptaAnnulation = GarantieAut.Comptabiliser
                    mnuLotàComptaPrint = GarantieAut.Comptabiliser
                End If
        
                Me.PopupMenu mnuLot, vbPopupMenuLeftButton
            End If
        End If
    End If
End If

End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xStatut As String
If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_Sort
        Case 1:  fgSelect_SortX 1
        Case 2: fgSelect_SortX 2
        Case 3, 13: fgSelect_Sort1 = 13: fgSelect_Sort2 = 13: fgSelect_Sort
        Case 4, 14: fgSelect_Sort1 = 14: fgSelect_Sort2 = 14: fgSelect_Sort
        Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_Sort
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
        Case 9: fgSelect_Sort1 = 9: fgSelect_Sort2 = 9: fgSelect_Sort
        Case 10: fgSelect_Sort1 = 10: fgSelect_Sort2 = 12: fgSelect_Sort
        Case 11: fgSelect_Sort1 = 11: fgSelect_Sort2 = 12: fgSelect_Sort
        Case 12: fgSelect_Sort1 = 12: fgSelect_Sort2 = 12: fgSelect_Sort
        Case 16:  fgSelect_SortX 16
    End Select
Else

    fgSelect_K = fgSelect.Row * fgSelect.Cols
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
 ''''       arrTOpe_Index = Val(fgSelect.TextArray(fgSelect.Cols - 1 + fgSelect_K))
        fgSelect.Col = 16
        arrTOpe_Index = Val(fgSelect.Text)
        xTOpe = arrTOpe(arrTOpe_Index)
    
        If xTOpe.IdRéférence > 0 Then
            mnuGarantieDisplay = GarantieAut.Consulter
            mnuGarantieModifier = False
            mnuGarantieAnnuler = False
            mnuGarantieEffacer = False
            mnuGarantieValider = False
            mnuGarantieAMJFin = False
            mnuGarantieMainLevéePartielle = False
            mnuGarantieMainLevée = False
            mnuGarantieAugmentation = False
         
            xStatut = xTOpe.Statut & xTOpe.StatutPlus
            If xStatut = "à? " Then
                mnuGarantieModifier = GarantieAut.Saisir
                mnuGarantieEffacer = GarantieAut.Saisir
            End If
            If xStatut = "àV " Then
              If Not GarantieAut.Xspécial And Trim(recTope.MajUsr) = Trim(usrId) Then
                    Call lstErr_Clear(lstErr, cmdContext, "! Vous ne pouvez pas valider vos opérations")
                Else
                    mnuGarantieValider = GarantieAut.Valider
                End If
            End If
            If xStatut = "   " Then
                mnuGarantieAMJFin = GarantieAut.Saisir
                mnuGarantieMainLevéePartielle = GarantieAut.Saisir
                mnuGarantieMainLevée = GarantieAut.Saisir
                mnuGarantieAugmentation = GarantieAut.Saisir
           End If
    
            Me.PopupMenu mnuGarantie, vbPopupMenuLeftButton
        End If
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
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
fgSelect.Clear: fgSelect.Row = 0
fgEchéancier.Clear: fgEchéancier.Row = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Private Sub fraCompta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub fraNature_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraGarantie_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub mnuAbandonner_Click()
cmdContext_Quit
End Sub

Private Sub mnucmdPrintList_TFlux_Click()
Dim Msg As String
If arrTFlux_Nb > 0 Then
    prtGarantie.recTFlux = recTFlux
    prtGarantie.CV1 = CV1
    ReDim prtGarantie.P_arrTFlux(arrTFlux_Nb)
    
    For I = 1 To fgEchéancier.Rows - 1
        fgEchéancier.Row = I
        fgEchéancier.Col = 11
        arrTFlux_Index = Val(fgEchéancier.Text)
        prtGarantie.P_arrTFlux(I) = arrTFlux(arrTFlux_Index)
    Next I
Msg = Format$(1, "000000") & Format$(arrTFlux_Nb, "000000")
    prtGarantieListEchéancier_Monitor Msg
End If

End Sub

Private Sub mnucmdPrintList_TOpe_Click()
Dim Msg As String
If arrTOpe_NB > 0 Then
    prtGarantie.recTope = recTope
    prtGarantie.CV1 = CV1
    ReDim prtGarantie.P_arrTOpe(arrTOpe_NB)
    For I = 1 To fgSelect.Rows - 1
        fgSelect.Row = I
        fgSelect.Col = 16
        arrTOpe_Index = Val(fgSelect.Text)
        prtGarantie.P_arrTOpe(I) = arrTOpe(arrTOpe_Index) ' arrTOpe(I)
    Next I
    Msg = Format$(1, "000000") & Format$(arrTOpe_NB, "000000")
    prtGarantieList_Monitor Msg
End If

End Sub

Private Sub mnucmdPrintGarantie_Click()
cmdPrint_Call "Garantie"
End Sub

Private Sub mnuComptaEchéancier_Click()
wAmjEchéanceTrt = DSys
mnuComptaEchéancier_Load
End Sub

Private Sub mnuEchéancierACU_Click()
recTFlux.Statut = "A"
recTFlux.StatutPlus = "CU"
cmdSave_recTflux
End Sub

Private Sub cmdSave_recTflux()
recTFlux.CptMvtUsr = usrId
recTFlux.CptMvtAMJ = DSys
recTFlux.CptMvtHMS = time_Hms
recTFlux.Method = "Update"
V = srvTFlux_Update(recTFlux)
If IsNull(V) Then
    arrTFlux(arrTFlux_Index) = recTFlux
Else
    Call MsgBox(V, vbCritical, "frmGarantie.cmdSave_recTflux")
End If
fgEchéancier_Display "T"

End Sub

Private Sub mnuEchéancierEnCours_Click()
recTFlux.Statut = " "
recTFlux.StatutPlus = "  "
cmdSave_recTflux

End Sub

Private Sub mnuEchéancierManuel_Click()
recTFlux.Statut = "M"
recTFlux.StatutPlus = "an"
cmdSave_recTflux

End Sub

Private Sub mnuGarantieAMJFin_Click()

currentAction = "Change_Validité"
cmdReset
fraGarantie_Load "Update"
If Not IsNull(saveTflux_Init) Then Exit Sub

fraNature.Enabled = False
fraBénéficaire.Enabled = False
fgEchéancier.Enabled = False
Call lbl_Style(lblAmjFin, True)
Call lbl_Style(lblPréavisNbj, True)
Call chk_Style(chkMainLevée, True)

meEnabled_Container "fraGarantie", False
fraGarantie.Enabled = True
wAMJEffet = DSys
txtAMJFin.Enabled = True
txtPréavisNbj.Enabled = True
chkMainLevée.Enabled = True
Call lstErr_Clear(lstErr, txtAMJFin, "> modifier la date de validité")
SSTab1.Tab = 1
blnEchéancier_Gen = True
blncmdOk_Visible = True
cmdOk.Caption = "OK_Validité"
cmdOk.FontSize = 6: cmdOk.FontName = "MS Serif"

End Sub

Private Sub mnuGarantieMainLevée_Click()
currentAction = "MainLevée"
cmdReset
fraGarantie_Load "Update"
If Not IsNull(saveTflux_Init) Then Exit Sub

fraNature.Enabled = False
fraBénéficaire.Enabled = False
fgEchéancier.Enabled = False
lblAMJEffet.Visible = True: txtAMJEffet.Visible = True
lblAMJEffet = "Date de la main levée"
Call lbl_Style(lblAMJEffet, True)
wAMJEffet = DSys: Call DTPicker_Set(txtAMJEffet, wAMJEffet)
meEnabled_Container "fraGarantie", False

fraGarantie.Enabled = True
txtAMJEffet.Enabled = True
txtAmjEngagement.Enabled = True
Call lstErr_Clear(lstErr, txtAMJEffet, "> préciser la date")
SSTab1.Tab = 1
blnEchéancier_Gen = True
blncmdOk_Visible = True
cmdOk.Caption = "OK_MainLevée"
cmdOk.FontSize = 6: cmdOk.FontName = "MS Serif"

cmdOk.Visible = True
End Sub

Private Sub mnuGarantieMainLevéePartielle_Click()
currentAction = "MainLevéePartielle"
cmdReset
fraGarantie_Load "Update"
If Not IsNull(saveTflux_Init) Then Exit Sub

fraNature.Enabled = False
fraBénéficaire.Enabled = False
fgEchéancier.Enabled = False
Call lbl_Style(lblCapital, True)
lblAMJEffet.Visible = True: txtAMJEffet.Visible = True
lblAMJEffet = "Date de la main levée"
Call lbl_Style(lblAMJEffet, True)
wAMJEffet = DSys: Call DTPicker_Set(txtAMJEffet, wAMJEffet)
meEnabled_Container "fraGarantie", False

fraGarantie.Enabled = True
txtCapital.Enabled = True: txtAMJEffet.Enabled = True
txtAmjEngagement.Enabled = True
Call lstErr_Clear(lstErr, txtCapital, "> modifier le montant")
SSTab1.Tab = 1
blnEchéancier_Gen = True
blncmdOk_Visible = True
cmdOk.Caption = "OK_MainLevéePartielle"
cmdOk.FontSize = 6: cmdOk.FontName = "MS Serif"

End Sub


Private Sub mnuGarantieAugmentation_Click()
currentAction = "Augmentation"
cmdReset
fraGarantie_Load "Update"
If Not IsNull(saveTflux_Init) Then Exit Sub

fraNature.Enabled = False
fraBénéficaire.Enabled = False
fgEchéancier.Enabled = False
Call lbl_Style(lblCapital, True)
lblAMJEffet.Visible = True: txtAMJEffet.Visible = True
lblAMJEffet = "Date de l'augmentation"
Call lbl_Style(lblAMJEffet, True)
wAMJEffet = DSys: Call DTPicker_Set(txtAMJEffet, wAMJEffet)
meEnabled_Container "fraGarantie", False

fraGarantie.Enabled = True
txtCapital.Enabled = True: txtAMJEffet.Enabled = True
txtAmjEngagement.Enabled = True
Call lstErr_Clear(lstErr, txtCapital, "> modifier le montant")
SSTab1.Tab = 1
blnEchéancier_Gen = True
blncmdOk_Visible = True
cmdOk.Caption = "OK_Augmentation"
cmdOk.FontSize = 6: cmdOk.FontName = "MS Serif"

End Sub

Private Sub mnuListEchéancier_Click()
V = DTPicker_Control(txtAmjMax, wAmjEchéanceTrt)
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, V): Exit Sub
cmdReset
mnuListEchéancier_Load

End Sub


Private Sub mnuComptaLotsàValider_Click()
ReDim arrTFlux(1)

recTFlux_Init recTFlux
recTFlux.Method = "SnapLotàC"
recTFlux.Statut = "à"

arrTFlux(0) = recTFlux
arrTFlux(0).CptMvtLot = "99999999"
Call srvTFlux_Load(recTFlux, arrTFlux(0))
arrTFlux_Nb = srvTFlux.arrTFlux_Nb
ReDim arrTFlux(arrTFlux_Nb)
For I = 1 To arrTFlux_Nb
    arrTFlux(I) = srvTFlux.arrTFlux(I)
    arrTFlux(I).CodeOpération = "$Lot"
    arrTFlux(I).Capital = 0
    arrTFlux(I).Intérêts = 0
Next I
SSTab1.Tab = 4
fgEchéancier_Display " "

If arrTFlux_Nb = 0 Then
    MsgBox "frmGarantie mnuComptaLotsàValider : PAS DE LOTS à TRAITER"
Else
    currentAction = constàCompta_Valider
End If

End Sub

Private Sub mnuEchéancierAvis_Click()
Call prtGarantie_Avis(mTope, arrTFlux(arrTFlux_Index))

End Sub

Private Sub mnuListàValider_Click()
currentAction = "mnuListàValider"
fgSelect_Load
End Sub

Private Sub mnuListEchéancierManuel_Click()
V = DTPicker_Control(txtAmjMax, wAmjEchéanceTrt)
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, V): Exit Sub
cmdReset

ReDim arrTFlux(1)

recTFlux_Init recTFlux
recTFlux.Method = "SnapLS"
recTFlux.Statut = "M"
recTFlux.StatutPlus = "an"
recTFlux.AmjEchéanceTrt = "00000000"

arrTFlux(0) = recTFlux
arrTFlux(0).Statut = "M"
arrTFlux(0).StatutPlus = "an"
arrTFlux(0).AmjEchéanceTrt = wAmjEchéanceTrt
arrTFlux(0).IdRéférence = 999999999
arrTFlux(0).IdSéquence = 32000

mnuListEchéancier_Display

End Sub

Private Sub mnuListGarantie_Click()
currentAction = "mnuListGarantie"
fgSelect_Load
End Sub


Private Sub mnuLotàComptaAnnulation_Click()
Call lstErr_Clear(lstErr, cmdContext, "!! Suppression de l'échéancier")
recTFlux.Method = constàCompta_Annuler
recTFlux.Statut = "à"
Call srvTFlux_Update(recTFlux)
fgEchéancier.Clear: fgEchéancier.Rows = 1: fgEchéancier_RowDisplay = 0
End Sub

Private Sub mnuLotàComptaPrint_Click()
TFlux_Compta.LotàCompta_Demande recTFlux.CptMvtLot

End Sub

Private Sub mnuLotàComptaValidation_Click()
Dim X As String, I As Integer

frmGarantie.Enabled = False
If blnComptaAuto Then
    X = vbYes
Else
    X = MsgBox("Cette action est irréversible. Confirmez-vous votre demande ?", vbYesNo + vbQuestion + vbDefaultButton2, "Garantie : Validation définitive du lot")
End If

If X = vbYes Then
    recTFlux.Method = constàCompta_Valider
    recTFlux.CptMvtUsr = usrId
    recTFlux.CptMvtAMJ = DSys
    recTFlux.CptMvtHMS = time_Hms
    mTflux = recTFlux
    srvTFlux_Update recTFlux
    TFlux_Compta.LotàCompta_Valider mTflux
    recTOpe_Init recTope
    For I = 1 To srvTFlux.arrTFlux_Nb
        recTFlux = srvTFlux.arrTFlux(I)
        Param_CodeOpération recTFlux.CodeOpération
        If paramTFlux_CodeOpération_Avis = "A" Then
                recTope.IdRéférence = recTFlux.IdRéférence
                recTope.Method = "SeekP0"
                recTope.Application = paramTFlux_Service
                If IsNull(srvTOpe_Monitor(recTope)) Then
                    Call prtGarantie_Avis(recTope, recTFlux)
                Else
                    Call MsgBox("Erreur lecture opération : " & recTope.IdRéférence, vbCritical, "Garantie : mnuLotàComptaValidation_Click(")
                End If
        End If
        
    Next I


    cmdReset
End If

frmGarantie.Enabled = True
AppActivate frmGarantie.Caption

End Sub

Private Sub mnuLotComptabilisé_Annuler_Click()
X = InputBox("Indiquer le numéro du lot à annuler :")
recTFlux_Init recTFlux
recTFlux.Statut = "C"
recTFlux.Method = "Compta_Ann"
recTFlux.CptMvtLot = CLng(Val(X))
If recTFlux.CptMvtLot = 0 Then
    Call lstErr_Clear(lstErr, cmdContext, " ! N° lot = 0 ")
Else
    srvTFlux_Update recTFlux
    Call lstErr_Clear(lstErr, cmdContext, " annulation comptabilisation du lot " & X)
End If

End Sub

Private Sub mnuGarantieDisplay_Click()
currentAction = constDisplay
fraGarantie_Load " "

Dim I As Integer, blnTest As Boolean
If mTope.Statut = " " Then     'GarantieAut.Xspécial And
    blnTest = True
    For I = 1 To arrTFlux_Nb
            If arrTFlux(I).Statut = " " Then blnTest = False
    Next I
    If blnTest Then
        blncmdOk_Visible = True: blncmdSave_Visible = False
        cmdOk.Caption = "FinAuto"
        cmdOk.Visible = True
        cmdContext.Caption = constcmdAbandonner
    End If
End If

End Sub

Private Sub mnuGarantieEffacer_Click()
blncmdSave_Visible = True
fraGarantie_Load constEffacer

End Sub

Private Sub mnuGarantieModifier_Click()
mnuGarantieSaisir_Click
fraGarantie_Load "Update"
blncmdOk_Visible = True: blncmdSave_Visible = True
End Sub

Private Sub mnuGarantieExtourne_Click()
Dim I As Integer

'fraGarantie_Load constExtourne
'currentAction = constExtourne
'mnuGarantiesaisir_Click
'fraGarantie.Enabled = False
'fraGarantie_Load "Update"
'blncmdOk_Visible = True: blncmdSave_Visible = False
'cmdOk.Caption = constExtourne
'cmdOk.Visible = True
'cmdContext.Caption = constcmdAbandonner
'wAmjExtourne = DSys
'txtAMJ.Visible = True
'Call DTPicker_Set(txtAMJ, DSys)
'minAmjExtourne = 99991231
'For I = 1 To arrTFlux_Nb
'    If arrTFlux(I).CodeOpération = "PR02" Then
'        If arrTFlux(I).Statut <> " " Then minAmjExtourne = arrTFlux(I).AmjFin
'    End If
'Next I

End Sub

Private Sub mnuGarantieSaisir_Click()
If GarantieAut.Saisir Then
    SSTab1.Tab = 1
    currentAction = constSaisie
    fgSelect.Enabled = False
    fgEchéancier.Enabled = False
    fgEchéancier.Clear: fgEchéancier.Rows = 1: fgEchéancier_RowDisplay = 0
    cmdReset
    fraGarantie.Enabled = True
    fraNature.Enabled = True
    meEnabled_Container "fraGarantie", True
    blncmdOk_Visible = True: blncmdSave_Visible = True
    blnAmjEchéance = False
    cboNature.SetFocus
    cmdContext.Caption = constcmdAbandonner
    blnControl = True
End If

End Sub

Private Sub mnuGarantieValider_Click()
currentAction = constValider
mnuGarantieSaisir_Click
fraGarantie.Enabled = False
fraGarantie_Load "Update"
blncmdOk_Visible = True: blncmdSave_Visible = True
cmdOk.Visible = True
currentAction = constValider
cmdContext.Caption = constcmdAbandonner
End Sub

Private Sub mnuOption_Click()
SSTab1.Tab = 0
fraOption.Visible = True
End Sub

Private Sub mnuQuitter_Click()
Unload Me
End Sub

Private Sub optAnnuel_Click()
If blnControl Then cmdControl
End Sub

Private Sub optAnnuel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optAnnuel
End Sub


Private Sub optEchéanceAnniversaire_Click()
If blnControl Then cmdControl
End Sub

Private Sub optEchéanceAnniversaire_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEchéanceAnniversaire
End Sub


Private Sub optEchéanceFinDeMois_Click()
If blnControl Then cmdControl
End Sub

Private Sub optEchéanceFinDeMois_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEchéanceFinDeMois

End Sub


Private Sub optMensuel_Click()
If blnControl Then cmdControl
End Sub


Private Sub optMensuel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optMensuel

End Sub


Private Sub optSemestriel_Click()
If blnControl Then cmdControl
End Sub

Private Sub optSemestriel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optSemestriel
End Sub


Private Sub optTrimestriel_Click()
If blnControl Then cmdControl
End Sub


Private Sub optTrimestriel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optTrimestriel
End Sub


Private Sub txtAmjEchéance_Change()
blnAmjEchéance = True
txtAmjEchéance_control

End Sub

Private Sub txtAmjEchéance_GotFocus()
DTPicker_GotFocus txtAMJEchéance
End Sub

Private Sub txtAMJEchéance_LostFocus()
DTPicker_LostFocus txtAMJEchéance
txtAmjEchéance_control
If blnControl Then cmdControl

End Sub

Private Sub txtAmjEchéance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DTPicker_GotFocus txtAMJEchéance
End Sub

Private Sub txtAmjEngagement_Change()
txtAmjEngagement_control
End Sub

Private Sub txtamjfin_Change()
txtAmjfin_control
End Sub

Private Sub txtAmjEngagement_GotFocus()
DTPicker_GotFocus txtAmjEngagement

End Sub

Private Sub txtamjfin_GotFocus()
DTPicker_GotFocus txtAMJFin

End Sub

Private Sub txtAmjEngagement_LostFocus()
DTPicker_LostFocus txtAmjEngagement
txtAmjEngagement_control
If blnControl Then cmdControl

End Sub

Private Sub txtamjfin_LostFocus()
DTPicker_LostFocus txtAMJFin
txtAmjfin_control
If blnControl Then cmdControl

End Sub


Private Sub txtAmjEngagement_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DTPicker_GotFocus txtAmjEngagement

End Sub

Private Sub txtamjfin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DTPicker_GotFocus txtAMJFin

End Sub

Public Sub cmdControl()
Dim X As String, wMensualité As Currency, wAmj As String, wTaux As Double
Dim wTA As Double, wTEG As Double

If Not Me.Enabled Then Exit Sub
Me.Enabled = False

cmdOk.Visible = False
cmdSave.Visible = False
blnControl = False

lstErr.Clear
lstErr.Height = 200

recTope = mTope
If currentAction = constSaisie Then
    blnControlBiatyp = True
Else
    blnControlBiatyp = False
End If
recTope.Application = paramTFlux_Service
recTope.IPA = "E"
recTope.NbjBase = "0"


If chkCommissionPériodique = "1" Then
    fraCommissionPériodique.Enabled = True
Else
    fraCommissionPériodique.Enabled = False
End If
If chkAmjEchéance = "1" Then
    txtAMJEchéance.Enabled = True
Else
    txtAMJEchéance.Enabled = False
End If
   
X = mCboNature
Call cbo_Value(mCboNature, cboNature)
If X <> mCboNature Then
    txtEngagementCompte = ""
    txtEchéanceCompte = ""
End If
recTope.Nature = mCboNature
V = TFlux_Compta.param_Nature(recTope.Nature)
If Not IsNull(V) Then Call lstErr_AddItem(lstErr, cmdContext, V)

X = Trim(txtDevise): V = CV_AttributS(X, CV1)
If Not IsNull(V) Then
    Call lstErr_AddItem(lstErr, cmdContext, V)
Else
    txtDevise = CV1.DeviseIso
End If
recTope.Devise = CV1.DeviseIso
libDevise = CV1.DeviseLibellé

X = num_Control(txtCapital, valX, 13, CV1.maxD)
recTope.Capital = valX
If recTope.Capital <= 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le montant garanti")

recTope.EngagementCorrCompte = paramTFlux_CompteDéblocageDesFonds
Compte_Load recTope.EngagementCorrCompte
txtDonneurDordre_Control
X = Format$(C_Racine.Numéro, "00000")
If X <> mDonneurDordre Then
    mDonneurDordre = X
    txtEngagementCompte = ""
    txtEchéanceCompte = ""
End If
    
If mDonneurDordre <> "00000" Then
    txtEngagementCompte_Control
    txtEchéanceCompte_Control
End If

If chkMainLevée = "1" Then
    txtPréavisNbj.Enabled = False
    recTope.PréavisNbj = 999
Else
    txtPréavisNbj.Enabled = True
    recTope.PréavisNbj = Val(txtPréavisNbj)
End If

X = Trim(txtRéférenceInterne)
recTope.RéférenceInterne = X
If X = "" Then
    Call lstErr_AddItem(lstErr, cmdContext, "? préciser la référence interne")
Else
    cmdSave.Visible = blncmdSave_Visible
End If

X = Trim(txtRéférenceExterne)
recTope.RéférenceExterne = X
If X = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser la référence externe")

txtAmjEngagement_control: recTope.AmjDébut = wAmjEngagement
txtAmjfin_control
recTope.AmjFin = wAmjFin

If currentAction = constSaisie Or currentAction = constValider Then
    If recTope.AmjDébut < paramAmjEngagementMin And chkComptaReprise = "0" Then Call lstErr_AddItem(lstErr, cmdContext, "? Reprise : cocher la case")
    If recTope.AmjDébut > paramAmjEngagementMax Then Call lstErr_AddItem(lstErr, cmdContext, "? date du Garantie > " & dateImp(paramAmjEngagementMax))
    If recTope.AmjDébut > recTope.AmjFin Then Call lstErr_AddItem(lstErr, cmdContext, "? date d'émission > date de validité ")
End If

If chkAmjEchéance = "0" Then Call DTPicker_Set(txtAMJEchéance, recTope.AmjDébut)

If lstErr.ListCount > 0 Then GoTo ExitSub


If chkCommissionPériodique = "0" Then
    recTope.TauxMarge = "0"
    recTope.AmjEchéanceS = "A"
    recTope.Périodicité = "T"
    wAmjEchéance = recTope.AmjDébut
    recTope.AmjEchéance1 = wAmjEchéance
Else
    cmdControl_txtTaux
    
    If optMensuel Then recTope.Périodicité = "M": fctPériodicité = "MoisAdd"
    If optTrimestriel Then recTope.Périodicité = "T": fctPériodicité = "TrimestreAdd"
    If optSemestriel Then recTope.Périodicité = "S": fctPériodicité = "SemestreAdd"
    If optAnnuel Then recTope.Périodicité = "A": fctPériodicité = "AnAdd"
    
       
    txtAmjEchéance_control
    recTope.AmjEchéance1 = wAmjEchéance
    
    paramAmjEchéanceMin = recTope.AmjDébut
    paramAmjEchéanceMax = recTope.AmjFin
    If recTope.AmjEchéance1 < paramAmjEchéanceMin Then Call lstErr_AddItem(lstErr, cmdContext, "? 1 ère échéance < " & dateImp(paramAmjEchéanceMin))
    If recTope.AmjEchéance1 > paramAmjEchéanceMax Then Call lstErr_AddItem(lstErr, cmdContext, "? 1 ère échéance >" & dateImp(paramAmjEchéanceMax))
    
    If optEchéanceFinDeMois Then
        recTope.AmjEchéanceS = "M"
    Else
        recTope.AmjEchéanceS = "A"
    End If
    
    If recTope.AmjEchéanceS = "M" Then
        recTope.AmjFin = dateFinDeMois(recTope.AmjFin)
        wAmj = dateFinDeMois(recTope.AmjEchéance1)
        If recTope.AmjEchéance1 <> wAmj Then Call lstErr_AddItem(lstErr, cmdContext, recTope.AmjEchéance1 & " ? n'est pas une fin de mois")
    End If
End If


If chkCommissionFlat = "1" Then
    X = num_Control(txtCommissionFlat, valX, 13, CV1.maxD)
    recTope.Frais = valX
Else
    recTope.Frais = 0
End If

If chkComptaReprise = "1" Then
    recTope.optReprise = "R"
Else
    recTope.optReprise = " "
End If

recTope.PériodeNb = 0
wAmjFin = dateElp("Jour", -1, recTope.AmjEchéance1)
Do
    V = fctTOpe_PériodeSuivante(recTope, wAmjDébut, wAmjFin)
    If IsNull(V) Then recTope.PériodeNb = recTope.PériodeNb + 1
Loop While wAmjFin < recTope.AmjFin

Call fctTOpe_Mensualité(recTope, CV1, wMensualité, wTaux, wTA, wTEG)
Select Case currentAction
    Case constValider
            V = fctTOpe_Compare(recTope, mTope)
            If Not IsNull(V) Then
                Call MsgBox("L'enregistrement après contrôle est différent de l'enregistrement lu :" & Chr$(13) & V, vbCritical, "me : cmdControl")
                Call lstErr_AddItem(lstErr, cmdContext, "? Erreur Contrôle validation")
            End If
    Case "MainLevéePartielle"
            If recTope.Capital >= mTope.Capital Then
                Call lstErr_AddItem(lstErr, cmdContext, "? nouveau montant >= ancien")
            End If
            cmdControl_txtAMJEffet_MainLevée
    Case "MainLevée"
            cmdControl_txtAMJEffet_MainLevée
            recTope.AmjFin = recTope.AmjDébut
    Case "Augmentation"
            If recTope.Capital <= mTope.Capital Then
                Call lstErr_AddItem(lstErr, cmdContext, "? nouveau montant <= ancien")
            End If
            cmdControl_txtAMJEffet_MainLevée
End Select

If blnEchéancier_Gen Or recTope.Statut = "à" Then
    Call fgEchéancier_Gen(wTaux)
    fgEchéancier_Display "T"
End If

If lstErr.ListCount = 0 Then
    cmdOk.Visible = blncmdOk_Visible
Else
'    SSTab1.Tab = 2
End If

ExitSub:

Me.Enabled = True
If cmdOk.Visible Then cmdOk.SetFocus
    
blnControl = True


End Sub

Public Sub fgEchéancier_Display(Fct As String)
Dim I As Integer

fgEchéancier.Visible = True
fgEchéancier.Clear: fgEchéancier_RowDisplay = 0: fgEchéancier_RowClick = 0
totalCapital = 0: totalIntérêts = 0

fgEchéancier.Rows = 1
fgEchéancier.FormatString = fgEchéancier_FormatString
fgEchéancier.Enabled = True
For arrTFlux_Index = 1 To arrTFlux_Nb
    recTFlux = arrTFlux(arrTFlux_Index)
    fgEchéancier.Rows = fgEchéancier.Rows + 1
    fgEchéancier.Row = fgEchéancier.Rows - 1
    fgEchéancier_DisplayLine
Next arrTFlux_Index

fgEchéancier_K = fgEchéancier.Cols
 
fgEchéancier_SortAD = 5
If fgEchéancier.Rows > 1 Then fgEchéancier_SortX 9
End Sub

Public Sub fgEchéancier_DisplayLine()
Dim K2 As Integer

fgEchéancier_K = (fgEchéancier.Row) * fgEchéancier.Cols
If recTFlux.CodeOpération = "$Lot" Then
    mEchéancierTope.RéférenceInterne = " Lot : " & Format(recTFlux.CptMvtLot, "### ### ")
    fgEchéancier.TextArray(1 + fgEchéancier_K) = " Lot à comptabiliser"
Else

    If mEchéancierTope.IdRéférence <> recTFlux.IdRéférence Then
        mEchéancierTope.RéférenceInterne = ""
        mEchéancierTope.IdRéférence = recTFlux.IdRéférence: srvTope_Find mEchéancierTope
    End If
    fgEchéancier.TextArray(1 + fgEchéancier_K) = TFlux_Compta.Param_CodeOpération(recTFlux.CodeOpération)
End If

fgEchéancier.TextArray(0 + fgEchéancier_K) = mEchéancierTope.RéférenceInterne
fgEchéancier.TextArray(2 + fgEchéancier_K) = Format(recTFlux.Capital + recTFlux.Intérêts, "#### ### ###.00 ")
fgEchéancier.TextArray(3 + fgEchéancier_K) = dateImp(recTFlux.AmjEchéanceTrt)
fgEchéancier.TextArray(4 + fgEchéancier_K) = recStatut_Libellé(recTFlux.Statut & recTFlux.StatutPlus)
fgEchéancier.TextArray(5 + fgEchéancier_K) = Format(recTFlux.Taux, "#0.00000 ") & recTFlux.TauxProvisoire
fgEchéancier.TextArray(6 + fgEchéancier_K) = "du " & dateImp(recTFlux.AmjDébut) & " au " & dateImp(recTFlux.AmjFin) & "   (" & recTFlux.Nbj & "j)"
fgEchéancier.TextArray(7 + fgEchéancier_K) = dateImp(recTFlux.CptMvtAMJ) & " " & timeImp(recTFlux.CptMvtHMS) & " " & recTFlux.CptMvtUsr
If recTFlux.CptMvtPièce <> 0 Then
    fgEchéancier.TextArray(8 + fgEchéancier_K) = "Pièce : " & Format(recTFlux.CptMvtPièce, "### ### ") & "." & Format(recTFlux.CptMvtLigne, "### ### ") & " Lot : " & Format(recTFlux.CptMvtLot, "### ### ")
End If
fgEchéancier.TextArray(9 + fgEchéancier_K) = Format(recTFlux.IdRéférence, "### ##0 ") & "_" & Format(recTFlux.IdSéquence, "### ### ")
fgEchéancier.TextArray(10 + fgEchéancier_K) = recTFlux.AmjEchéanceTrt
fgEchéancier.TextArray(fgEchéancier.Cols - 1 + fgEchéancier_K) = arrTFlux_Index

Select Case recTFlux.CodeOpération
    Case "GA01", "GA11"
    Case Else:   fgEchéancier.Col = 2: fgEchéancier.CellForeColor = errUsr.ForeColor
End Select
If recTFlux.Statut = "A" Then fgEchéancier.Col = 4: fgEchéancier.CellForeColor = errUsr.ForeColor


End Sub

Private Sub txtCapital_GotFocus()
txt_GotFocus txtCapital

End Sub


Private Sub txtCapital_KeyPress(KeyAscii As Integer)
If CV1.maxD = 0 Then
    Call num_KeyAscii(KeyAscii)
Else
    Call num_KeyAsciiD(KeyAscii, txtCapital)
End If

End Sub


Private Sub txtCapital_LostFocus()
txt_LostFocus txtCapital
If blnControl Then cmdControl

End Sub

Private Sub txtDevise_GotFocus()
txt_GotFocus txtDevise

End Sub


Private Sub txtDevise_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtDevise_LostFocus()
txt_LostFocus txtDevise
If blnControl Then cmdControl

End Sub


Private Sub txtEchéanceCompte_GotFocus()
txt_GotFocus txtEchéanceCompte

End Sub


Private Sub txtEchéanceCompte_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtEchéanceCompte)

End Sub


Private Sub txtEchéanceCompte_LostFocus()
txt_LostFocus txtEchéanceCompte
If blnControl Then cmdControl

End Sub

Private Sub txtDonneurDordre_GotFocus()
txt_GotFocus txtDonneurDordre

End Sub


Private Sub txtDonneurDordre_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtDonneurDordre)

End Sub


Private Sub txtDonneurDordre_LostFocus()
txt_LostFocus txtDonneurDordre
If blnControl Then cmdControl

End Sub

Private Sub txtEngagementCompte_GotFocus()
txt_GotFocus txtEngagementCompte

End Sub


Private Sub txtEngagementCompte_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtEngagementCompte)

End Sub


Private Sub txtEngagementCompte_LostFocus()
txt_LostFocus txtEngagementCompte
If blnControl Then cmdControl

End Sub

Private Sub txtcommissionflat_GotFocus()
txt_GotFocus txtCommissionFlat

End Sub


Private Sub txtcommissionflat_KeyPress(KeyAscii As Integer)
If CV1.maxD = 0 Then
    Call num_KeyAscii(KeyAscii)
Else
    Call num_KeyAsciiD(KeyAscii, txtCommissionFlat)
End If

End Sub


Private Sub txtcommissionflat_LostFocus()
txt_LostFocus txtCommissionFlat
If blnControl Then cmdControl

End Sub

Private Sub txtPréavisNbj_GotFocus()
txt_GotFocus txtPréavisNbj
End Sub


Private Sub txtPréavisNbj_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtPréavisNbj)
End Sub


Private Sub txtPréavisNbj_LostFocus()
txt_LostFocus txtPréavisNbj
If blnControl Then cmdControl


End Sub

Private Sub txtRéférenceExterne_GotFocus()
txt_GotFocus txtRéférenceExterne

End Sub


Private Sub txtRéférenceExterne_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtRéférenceExterne_LostFocus()
txt_LostFocus txtRéférenceExterne
If blnControl Then cmdControl

End Sub


Private Sub txtRéférenceInterne_GotFocus()
txt_GotFocus txtRéférenceInterne

End Sub


Private Sub txtRéférenceInterne_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtRéférenceInterne_LostFocus()
txt_LostFocus txtRéférenceInterne
If blnControl Then cmdControl

End Sub


Private Sub txtTaux_GotFocus()
txt_GotFocus txtTaux

End Sub


Private Sub txtTaux_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtTaux)
End Sub


Private Sub txtTaux_LostFocus()
txt_LostFocus txtTaux
If blnControl Then cmdControl

End Sub

Public Sub fgEchéancier_Gen(mTaux As Double)
Dim V As Variant, MsgErr As String, I As Integer, wNbj As Long, wIntérêts As Currency
Dim wAmjDébut  As String, wAmjFin As String

Dim totalAmortissement As Currency
On Error GoTo Error_Handle

ReDim arrTFlux(recTope.PériodeNb + 5)
xTOpe = recTope
srvTFlux.recTFlux_Init recTFlux

fgEchéancier.Clear: fgEchéancier.Rows = 1: fgEchéancier_RowDisplay = 0

wNbj = Périodicité_Nbj(recTope.Périodicité)

With recTFlux                                   ' Engagement
    .IdSéquence = 1
    .CodeOpération = "GA01"
    .Capital = recTope.Capital
    .Taux = recTope.TauxMarge
    .AmjEchéanceTrt = recTope.AmjDébut
    .AmjDébut = recTope.AmjDébut
    .AmjFin = recTope.AmjFin
    .AmjOpération = recTope.AmjDébut
    .AmjValeur = recTope.AmjDébut
End With
arrTFlux(recTFlux.IdSéquence) = recTFlux

If recTope.Frais <> 0 Then
    If Trim(recTope.EchéanceCompte) = "" Then  ' Echéance commission
        recTFlux.CodeOpération = "GA61"
    Else
        recTFlux.CodeOpération = "GA51"
    End If
    With recTFlux                                   ' Frais
        .IdSéquence = recTFlux.IdSéquence + 1
        .Capital = 0
        .Nbj = 0
        .Intérêts = recTope.Frais
        .AmjEchéanceTrt = recTope.AmjDébut
        .AmjDébut = recTope.AmjDébut
        .AmjFin = recTope.AmjFin
        .AmjOpération = recTope.AmjDébut
        .AmjValeur = recTope.AmjDébut
    End With
    arrTFlux(recTFlux.IdSéquence) = recTFlux

End If

If recTope.TauxMarge <> 0 Then
    wAmjFin = dateElp("Jour", -1, recTope.AmjEchéance1)
    If Trim(recTope.EchéanceCompte) = "" Then  ' Echéance commission
        recTFlux.CodeOpération = "GA62"
    Else
        recTFlux.CodeOpération = "GA52"
    End If
    If Trim(recTope.TauxRéférence) = "Montant" Then
        wIntérêts = recTope.TauxMarge
    Else
        wIntérêts = Round(recTope.Capital * mTaux, CV1.maxD)
    End If
    Do
        V = fctTOpe_PériodeSuivante(xTOpe, wAmjDébut, wAmjFin)
        If Not IsNull(V) Then MsgErr = "? calcul échéance " & I: Error 9999
        
        With recTFlux
            .IdSéquence = recTFlux.IdSéquence + 1
            .Intérêts = wIntérêts
            .Capital = 0
            .Nbj = wNbj
            .AmjEchéanceTrt = wAmjDébut
            .AmjDébut = wAmjDébut
            .AmjFin = wAmjFin
            .AmjOpération = wAmjDébut
            .AmjValeur = wAmjDébut
        End With
        arrTFlux(recTFlux.IdSéquence) = recTFlux
    Loop While wAmjFin < recTope.AmjFin
End If

With recTFlux                                   ' Fin de validité
    .IdSéquence = recTFlux.IdSéquence + 1
    .Capital = recTope.Capital
    .Intérêts = 0
    .Nbj = 0
    .AmjDébut = recTope.AmjDébut
    .AmjFin = recTope.AmjFin
End With
If recTope.PréavisNbj = 999 Then
    recTFlux.CodeOpération = "GA22"                  ' attente main levée
    recTFlux.AmjEchéanceTrt = recTope.AmjFin
Else
    recTFlux.CodeOpération = "GA02"                  ' fin  de validité
    recTFlux.AmjEchéanceTrt = dateElp("Jour", recTope.PréavisNbj, recTope.AmjFin)
End If
recTFlux.AmjOpération = recTFlux.AmjEchéanceTrt
recTFlux.AmjValeur = recTFlux.AmjEchéanceTrt
arrTFlux(recTFlux.IdSéquence) = recTFlux


arrTFlux_Nb = recTFlux.IdSéquence

For I = 1 To arrTFlux_Nb
    arrTFlux(I).Method = constAddNew
    arrTFlux(I).IdRéférence = recTope.IdRéférence

   If recTope.optReprise = "R" And arrTFlux(I).AmjEchéanceTrt < mAMJReprise Then
        arrTFlux(I).Statut = "R"
        arrTFlux(I).StatutPlus = "ep"
    Else
        Param_CodeOpération arrTFlux(I).CodeOpération
        If paramTFlux_CodeOpération_Compta <> "A" Then arrTFlux(I).Statut = "M": arrTFlux(I).StatutPlus = "an"
End If

Next I
Exit Sub
'---------------------------------------------------------
Error_Handle:
'---------------------------------------------------------

Call MsgBox(MsgErr, vbCritical, "frmGarantie.fgEchéancier_Gen")


End Sub

Public Function Compte_Load(mCompteNuméro As String)
Compte_Load = Null
recCompteInit recCompte
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = CV1.DeviseN
recCompte.Numéro = mCompteNuméro
recCompte.BiaTyp = "000"
recCompte.BiaNum = "00"
recCompte.Method = "SeekL1"
If Not IsNull(srvCompteFind(recCompte)) Then Call lstErr_AddItem(lstErr, lstErr, "? compte inconnu : " & mCompteNuméro): Compte_Load = "?": Exit Function

If recCompte.Situation <> " " Then
    Select Case recCompte.Situation
        Case "E":
        Case "B":  ''''Call lstErr_AddItem(lstErr, lstErr, " ? Compte bloqué : " & mCompteNuméro): Compte_Load = "?"
        Case "A": Call lstErr_AddItem(lstErr, lstErr, " ? Compte annulé : " & mCompteNuméro): Compte_Load = "?"
        Case Else: Call lstErr_AddItem(lstErr, lstErr, " ? Situation du compte : " & mCompteNuméro): Compte_Load = "?"
    End Select
End If

End Function

Public Sub cmdPrint_Call(Fct As String)
Dim Msg As String
If arrTFlux_Nb > 0 Then
    prtGarantie.recTope = recTope
    prtGarantie.CV1 = CV1
    ReDim prtGarantie.arrTFlux(arrTFlux_Nb)
    For I = 1 To arrTFlux_Nb
        prtGarantie.arrTFlux(I) = arrTFlux(I)
    Next I
    Msg = Format$(1, "000000") & Format$(arrTFlux_Nb, "000000") & Fct
    prtGarantie_Monitor Msg
End If

End Sub

Public Sub fraGarantie_Load(Fct As String)
'2000-01-04 cmdReset
fgSelect_RowClick = 0
Call fgSelect_Color(fgSelect_RowDisplay, vbCyan, fgSelect_ColorClick) 'txtUsr.BackColor)
blnControl = False
xTOpe.Method = "SeekP0"
V = srvTOpe_Monitor(xTOpe)
If IsNull(V) Then
    libRéférenceInterne = Trim(xTOpe.RéférenceInterne) & "_" & Compte_Imp(xTOpe.EngagementCompte)
    blnAmjEchéance = True
    SSTab1.Tab = 1
    mTope = xTOpe
    mTope.Method = Fct
    mCboNature = mTope.Nature
    cbo_Scan mTope.Nature, cboNature
    mDonneurDordre = mId$(mTope.EngagementCompte, 1, 5)
    txtDonneurDordre = mDonneurDordre
    txtEngagementCompte = Compte_Display(mTope.EngagementCompte)
    txtDevise = mTope.Devise
    txtCapital = Format$(mTope.Capital, "### ### ### ##0.00")
    txtRéférenceInterne = Trim(mTope.RéférenceInterne)
    txtRéférenceExterne = Trim(mTope.RéférenceExterne)
    Call DTPicker_Set(txtAmjEngagement, mTope.AmjDébut): wAmjEngagement = mTope.AmjDébut
    Call DTPicker_Set(txtAMJFin, mTope.AmjFin): wAmjFin = mTope.AmjFin
    
    If mTope.PréavisNbj = 999 Then
        chkMainLevée = "1"
    Else
        chkMainLevée = "0"
        txtPréavisNbj = mTope.PréavisNbj
    End If
    
    If Trim(mTope.EchéanceCompte) <> "" Then 'paramTFlux_CompteCommissionàRéclamer Then
        chkCommissionàRéclamer = "0"
        txtEchéanceCompte = Compte_Display(mTope.EchéanceCompte)
    Else
        chkCommissionàRéclamer = "1"
        txtEchéanceCompte = ""
    End If
    
    If mTope.AmjEchéance1 = mTope.AmjDébut Then
        chkAmjEchéance = "0"
    Else
        chkAmjEchéance = "1"
    End If
    
    Call DTPicker_Set(txtAMJEchéance, mTope.AmjEchéance1): wAmjEchéance = mTope.AmjEchéance1
    If mTope.AmjEchéanceS = "M" Then
        optEchéanceFinDeMois = True
    Else
        optEchéanceAnniversaire = True
    End If
    If mTope.TauxMarge <> 0 Then
        chkCommissionPériodique = "1"
        If Trim(mTope.TauxRéférence) <> "Montant" Then
            txtTaux = Format$(mTope.TauxMarge, "#0.00000")
            optComPériodiqueTaux = True
        Else
            optComPériodiqueMontant = True
            txtTaux = Format$(mTope.TauxMarge, "##### ##0.00")
       End If
    Else
        chkCommissionPériodique = "0"
        txtTaux = ""
    End If
    
    Select Case mTope.Périodicité
        Case "M": optMensuel = True: fctPériodicité = "MoisAdd"
        Case "T": optTrimestriel = True: fctPériodicité = "TrimestreAdd"
        Case "S": optSemestriel = True: fctPériodicité = "SemestreAdd"
        Case "A": optAnnuel = True: fctPériodicité = "AnAdd"
        Case Else: optMensuel = True: fctPériodicité = "MoisAdd"
   End Select
      
    If mTope.Frais = 0 Then
        chkCommissionFlat = "0"
    Else
        chkCommissionFlat = "1"
        txtCommissionFlat = Format$(mTope.Frais, "### ### ### ##0.00")
    End If
    
    mAMJReprise = mTope.MajAMJ
    If mTope.optReprise = "R" Then
        chkComptaReprise = "1"
    Else
        chkComptaReprise = "0"
    End If
    libStatut = "Statut         : " & recStatut_Libellé(mTope.Statut & mTope.StatutPlus) & Chr$(13) _
                & "Référence  : " & Format$(mTope.IdRéférence, "#### ### ##0") & Chr$(13) & Chr$(13) _
                & "Saisi par  : " & mTope.MajUsr & Chr$(13) _
                & "                 : " & dateImp(mTope.MajAMJ) & " " & timeImp(mTope.MajHMS) & Chr$(13) & Chr$(13) _
                & "Validé par :" & mTope.ValUsr & Chr$(13) _
                & "                 : " & dateImp(mTope.valAMJ) & " " & timeImp(mTope.ValHMS)
   cmdControl
    
    If mTope.Statut = "à" Then
        Select Case mTope.StatutPlus
            Case Is = "V ": cmdOk.Caption = constValider
                            cmdSave.Caption = constàModifier
         
            Case Is = "? "
                        If Fct = constEffacer Then
                            cmdSave.Caption = constEffacer
                        Else
                            cmdOk.Caption = constàValider
                        End If
        End Select
    Else
        fgEchéancier_Load
    End If
    cmdSave.Visible = blncmdSave_Visible
    blnfgSelect_DisplayLine = True
End If

End Sub
Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

If lRow > 0 Then
    fgSelect.Row = lRow
    For I = 0 To 16
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 0 To 16
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.Col = 0
    End If
End If

End Sub

Public Sub fgEchéancier_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgEchéancier.Row


If lRow > 0 Then
    fgEchéancier.Row = lRow
    For I = 0 To fgEchéancier.Cols - 1
        fgEchéancier.Col = I: fgEchéancier.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgEchéancier.Row = mRow
    If fgEchéancier.Row > 0 Then
        lRow = fgEchéancier.Row
        lColor_Old = fgEchéancier.CellBackColor
        For I = 0 To fgEchéancier.Cols - 1
          fgEchéancier.Col = I: fgEchéancier.CellBackColor = lColor
        Next I
        fgEchéancier.Col = 0
    End If
End If

End Sub

Public Sub fgEchéancier_AddNew()
V = srvTFlux_Dtaq_Put("Init", recTFlux)
If Not IsNull(V) Then fgEchéancier_Delete: Exit Sub
For I = 1 To arrTFlux_Nb
'    arrTFlux(I).Method = constAddNew
'    arrTFlux(I).IdRéférence = recTope.IdRéférence
    recTFlux = arrTFlux(I)
'   If recTope.optReprise = "R" And recTFlux.AmjEchéanceTrt < DSys Then
'        recTFlux.Statut = "R"
'        recTFlux.StatutPlus = "ep"
'    Else
'        Param_CodeOpération recTFlux.CodeOpération
'        If paramTFlux_CodeOpération_Compta <> "A" Then recTFlux.Statut = "M": recTFlux.StatutPlus = "an"
'    End If
'     arrTFlux(I) = recTFlux
    V = srvTFlux_Dtaq_Put("Add", recTFlux)
    If Not IsNull(V) Then fgEchéancier_Delete: Exit Sub
Next I
V = srvTFlux_Dtaq_Put("Snd", recTFlux)
If Not IsNull(V) Then fgEchéancier_Delete: Exit Sub

End Sub

Public Sub fgEchéancier_Update()
V = srvTFlux_Dtaq_Put("Init", arrTFlux(1))
If Not IsNull(V) Then: Exit Sub
For I = 1 To arrTFlux_Nb
    If arrTFlux(I).Method = constUpdate Or arrTFlux(I).Method = constAddNew Or arrTFlux(I).Method = constDelete Then
        V = srvTFlux_Dtaq_Put("Add", arrTFlux(I))
        If Not IsNull(V) Then Exit Sub
    End If
Next I
V = srvTFlux_Dtaq_Put("Snd", arrTFlux(1))
If Not IsNull(V) Then Exit Sub

End Sub

Public Sub fgEchéancier_Delete()
Call lstErr_AddItem(lstErr, cmdContext, V)
Call lstErr_AddItem(lstErr, cmdContext, "!! Suppression de l'échéancier")
arrTFlux(1).Method = "DeleteAll"
Call srvTFlux_Update(arrTFlux(1))
End Sub

Public Sub fgEchéancier_Load()
ReDim arrTFlux(1)

recTFlux_Init recTFlux
recTFlux.Method = "SnapP0"
recTFlux.IdRéférence = mTope.IdRéférence

arrTFlux(0) = recTFlux
arrTFlux(0).IdSéquence = 999

Call srvTFlux_Load(recTFlux, arrTFlux(0))
arrTFlux_Nb = srvTFlux.arrTFlux_Nb
ReDim arrTFlux(arrTFlux_Nb)
For I = 1 To arrTFlux_Nb
    arrTFlux(I) = srvTFlux.arrTFlux(I)
Next I

fgEchéancier_Display "T"

End Sub

Public Sub cmdSave_àCompta()

blnErr = False
mTflux = arrTFlux(1)
mTflux.Method = constàCompta
V = srvTFlux_Update(mTflux)
If Not IsNull(V) Then MsgBox "frmGarantie cmdSaveàCompta : recherche numlot": Exit Sub
mTflux.CptMvtAMJ = DSys
mTflux.CptMvtHMS = time_Hms

V = srvTFlux_Dtaq_Put("Init", arrTFlux(1))

If IsNull(V) Then
    For I = 1 To arrTFlux_Nb
        arrTFlux(I).Method = "Update"
        arrTFlux(I).Statut = "à"
        arrTFlux(I).StatutPlus = "C "
        arrTFlux(I).CptMvtLot = mTflux.CptMvtLot
        arrTFlux(I).CptMvtPièce = I
        arrTFlux(I).CptMvtAMJ = mTflux.CptMvtAMJ
        arrTFlux(I).CptMvtHMS = mTflux.CptMvtHMS
        V = srvTFlux_Dtaq_Put("Add", arrTFlux(I))
        If Not IsNull(V) Then blnErr = True: Exit Sub
    Next I
    V = srvTFlux_Dtaq_Put("Snd", arrTFlux(1))
    If Not IsNull(V) Then blnErr = True: Exit Sub
End If

If Not blnErr Then
    lastActiveControl_Name = ""
    cmdOk.Visible = False
    cmdSave.Visible = False
    Call lstErr_AddItem(lstErr, cmdContext, "àCompta - N° lot : " & mTflux.CptMvtLot)
    If Not blnComptaAuto Then TFlux_Compta.LotàCompta_Demande mTflux.CptMvtLot
Else
    Call lstErr_AddItem(lstErr, cmdContext, V)
    cmdReset
    MsgBox "frmGarantie cmdSaveàCompta : à faire annuler demande de comptabilisation"
End If
fgEchéancier.Clear: fgEchéancier.Rows = 1: fgEchéancier_RowDisplay = 0
End Sub

Public Sub fgSelect_DisplayLine()
fgSelect_K = (fgSelect.Row) * fgSelect.Cols
fgSelect.TextArray(7 + fgSelect_K) = arrTOpe(arrTOpe_Index).Nature
fgSelect.TextArray(1 + fgSelect_K) = Format(arrTOpe(arrTOpe_Index).Capital, "#### ### ###.00 ")
fgSelect.TextArray(2 + fgSelect_K) = arrTOpe(arrTOpe_Index).Devise
fgSelect.TextArray(3 + fgSelect_K) = dateImp(arrTOpe(arrTOpe_Index).AmjDébut)
fgSelect.TextArray(4 + fgSelect_K) = dateImp(arrTOpe(arrTOpe_Index).AmjFin)
fgSelect.TextArray(5 + fgSelect_K) = Compte_Imp(arrTOpe(arrTOpe_Index).EngagementCompte)
Call CV_AttributS(arrTOpe(arrTOpe_Index).Devise, CV1)
recCompteInit recCompte
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = CV1.DeviseN
recCompte.Numéro = arrTOpe(arrTOpe_Index).EngagementCompte
mdbCptP0_Find recCompte
'Call Compte_Load(arrTOpe(arrTOpe_Index).EngagementCompte)
fgSelect.TextArray(6 + fgSelect_K) = recCompte.Intitulé
fgSelect.TextArray(0 + fgSelect_K) = arrTOpe(arrTOpe_Index).RéférenceInterne
fgSelect.TextArray(8 + fgSelect_K) = arrTOpe(arrTOpe_Index).RéférenceExterne
fgSelect.TextArray(9 + fgSelect_K) = ""
fgSelect.TextArray(9 + fgSelect_K) = recStatut_Libellé(arrTOpe(arrTOpe_Index).Statut & arrTOpe(arrTOpe_Index).StatutPlus)
fgSelect.TextArray(10 + fgSelect_K) = arrTOpe(arrTOpe_Index).MajUsr & " " & dateImp(arrTOpe(arrTOpe_Index).MajAMJ) & " " & timeImp(arrTOpe(arrTOpe_Index).MajHMS)
fgSelect.TextArray(11 + fgSelect_K) = arrTOpe(arrTOpe_Index).ValUsr & " " & dateImp(arrTOpe(arrTOpe_Index).valAMJ) & " " & timeImp(arrTOpe(arrTOpe_Index).ValHMS)
fgSelect.TextArray(12 + fgSelect_K) = arrTOpe(arrTOpe_Index).IdRéférence
fgSelect.TextArray(13 + fgSelect_K) = arrTOpe(arrTOpe_Index).AmjDébut
fgSelect.TextArray(14 + fgSelect_K) = arrTOpe(arrTOpe_Index).AmjFin
fgSelect.TextArray(15 + fgSelect_K) = ""
fgSelect.TextArray(16 + fgSelect_K) = arrTOpe_Index

End Sub

Public Sub txtEngagementCompte_Control()
Dim X As String, wBiatypEngagement As String * 3

If mDonneurDordre < 30000 Then
    wBiatypEngagement = "943"
Else
    wBiatypEngagement = paramTFlux_BiatypEngagement
End If

If Trim(txtEngagementCompte) = "" Then
    X = mDonneurDordre & "000010"
    Call Compte_BiaTyp(X, wBiatypEngagement)
    txtEngagementCompte = X
End If

X = num_Control(txtEngagementCompte, valX, 11, 0)
recTope.EngagementCompte = valX
V = Compte_Load(recTope.EngagementCompte)
libEngagementCompte = Trim(DicLib(13, recCompte.BiaTyp))
txtEngagementCompte = Compte_Display(recTope.EngagementCompte)
If mId$(recTope.EngagementCompte, 1, 5) <> mDonneurDordre Then
    Call lstErr_AddItem(lstErr, cmdContext, "? racine Engagement <> donneur d'ordre")
Else
    If IsNull(V) And blnControlBiatyp Then
        If recTope.EngagementCompte <> mEngagementCompte Then
            If mId$(recTope.EngagementCompte, 6, 3) <> wBiatypEngagement Then
                X = MsgBox("Le type de compte attendu est : " & wBiatypEngagement & Chr$(13) & " confirmez-vous ce compte?", vbYesNo + vbQuestion + vbDefaultButton2, "Garantie : Compte de Garantie ")
                If X = vbYes Then mEngagementCompte = recTope.EngagementCompte
            End If
        End If
    End If
End If
   
End Sub
Public Sub txtEchéanceCompte_Control()
Dim X As String

If chkCommissionàRéclamer = "1 " Then
    recTope.EchéanceCompte = "" 'paramTFlux_CompteCommissionàRéclamer
    'V = Compte_Load(recTope.EchéanceCompte)
    libEchéanceCompte = "" 'recCompte.Intitulé
Else

    If Trim(txtEchéanceCompte) = "" Then
        X = mDonneurDordre & "000010"
        Call Compte_BiaTyp(X, paramTFlux_BiatypEchéance)
        txtEchéanceCompte = X
    Else
    '    Mid$(txtEchéanceCompte, 1, 5) = mDonneurDordre
    End If
    
    X = num_Control(txtEchéanceCompte, valX, 11, 0)
    recTope.EchéanceCompte = valX
    V = Compte_Load(recTope.EchéanceCompte)
    libEchéanceCompte = recCompte.Intitulé
    libEchéanceTypeDeCompte = Trim(DicLib(13, recCompte.BiaTyp))
    If IsNull(V) And blnControlBiatyp Then
        If recTope.EchéanceCompte <> mEChéanceCompte Then
            If mId$(recTope.EchéanceCompte, 6, 3) <> paramTFlux_BiatypEchéance Then
                X = MsgBox("Le type de compte attendu est : " & paramTFlux_BiatypEchéance & Chr$(13) & " confirmez-vous ce compte?", vbYesNo + vbQuestion + vbDefaultButton2, "Garantie : Compte de Garantie ")
                If X = vbYes Then mEChéanceCompte = recTope.EchéanceCompte
            End If
        End If
    End If
End If
txtEchéanceCompte = Compte_Display(recTope.EchéanceCompte)
End Sub


Public Sub mnuComptaEchéancier_Load()
mnuListEchéancier_Load
If arrTFlux_Nb = 0 Then
    MsgBox "frmGarantie mnuComptaEchéancier : PAS D'ECHEANCE A TRAITER"
Else
    cmdOk.Caption = constàCompta: cmdOk.Visible = True
    currentAction = constàCompta
End If

End Sub

Public Sub mnuListEchéancier_Load()
ReDim arrTFlux(1)

recTFlux_Init recTFlux
recTFlux.Method = "SnapLE"
recTFlux.AmjEchéanceTrt = "00000000"

arrTFlux(0) = recTFlux
arrTFlux(0).AmjEchéanceTrt = wAmjEchéanceTrt
arrTFlux(0).IdRéférence = 999999999
arrTFlux(0).IdSéquence = 32000
mnuListEchéancier_Display

End Sub

Public Sub fgEchéancier_Extourne()
Dim curX As Currency, I As Integer, I1 As Integer, wNbj As Long

curX = 0: I1 = 0
recTFlux = arrTFlux(arrTFlux_Nb)

For I = 1 To arrTFlux_Nb
   
    If arrTFlux(I).CodeOpération = "PR02" And arrTFlux(I).Statut = " " Then
        If I1 = 0 Then I1 = I
        curX = curX + arrTFlux(I).Capital
        arrTFlux(I).Method = constUpdate
        arrTFlux(I).Statut = "A"
        arrTFlux(I).StatutPlus = "RA"
    End If
Next I

With recTFlux                                   ' Extourne
    .Method = constAddNew
    .IdSéquence = recTFlux.IdSéquence + 1
    .CodeOpération = "PR05"
    .Capital = curX
    .Intérêts = 0
    .Taux = recTope.TauxMarge
    .AmjEchéanceTrt = DSys
    .AmjDébut = wAmjExtourne
    .AmjFin = wAmjExtourne
    .AmjOpération = DSys
    .AmjValeur = wAmjExtourne
End With
ReDim Preserve arrTFlux(arrTFlux_Nb + 2)
arrTFlux_Nb = arrTFlux_Nb + 1
arrTFlux(arrTFlux_Nb) = recTFlux

If I1 > 0 Then
    xTOpe = recTope
    With xTOpe
        .Capital = curX
        .AmjDébut = arrTFlux(I1).AmjDébut
        .AmjFin = wAmjExtourne
    End With
    V = fctTOpe_Intérêts(xTOpe, CV1, curX, wNbj)
    
    If IsNull(V) Then
        If curX <> 0 Then
            With recTFlux                                   ' intérêts décalés
                .IdSéquence = recTFlux.IdSéquence + 1
                .CodeOpération = "PR04"
                .Capital = 0
                .Intérêts = curX
                .Nbj = wNbj
                .AmjDébut = arrTFlux(I1).AmjDébut
            End With
            arrTFlux_Nb = arrTFlux_Nb + 1
            arrTFlux(arrTFlux_Nb) = recTFlux
        End If
    End If
End If

End Sub

Public Sub txtDonneurDordre_Control()
If Trim(txtDonneurDordre) = "" Then
    Call lstErr_AddItem(lstErr, cmdContext, "? préciser le donneur d'ordre")
Else
    X = num_Control(txtDonneurDordre, valX, 5, 0)
    txtDonneurDordre = valX
    If valX <> C_Racine.Numéro Then
        recRacineInit C_Racine
        C_Racine.Method = "SeekL0"
        C_Racine.Numéro = valX
        If IsNull(srvRacineMon(C_Racine)) Then
            libDonneurDordre = C_Racine.Intitulé
        Else
            Call lstErr_AddItem(lstErr, cmdContext, "? racine du donneur d'ordre")
            libDonneurDordre = "? inconnu"
        End If
    End If
End If
End Sub

Public Sub mnuComptaDossier()
arrTFlux_Index = 0
For I = 1 To arrTFlux_Nb
    If arrTFlux(I).Statut = " " And arrTFlux(I).AmjEchéanceTrt <= DSys Then
        arrTFlux_Index = arrTFlux_Index + 1
        arrTFlux(arrTFlux_Index) = arrTFlux(I)
    End If
Next I

If arrTFlux_Index = 0 Then
    cmdPrint_Call constValider
    cmdReset
    Exit Sub
End If

arrTFlux_Nb = arrTFlux_Index
cmdSave_àCompta
recTFlux = arrTFlux(1)
mnuLotàComptaValidation_Click

End Sub

Public Sub mnuEchéancier_Set()
mnuEchéancierAvis.Enabled = False
mnuEchéancierACU.Enabled = False
mnuEchéancierEnCours.Enabled = False
mnuEchéancierManuel.Enabled = False

If paramTFlux_CodeOpération_Avis = "A" Then mnuEchéancierAvis.Enabled = True

If paramTFlux_CodeOpération_Compta = "A" Then
    If recTFlux.Statut = " " Then mnuEchéancierACU.Enabled = True: mnuEchéancierManuel.Enabled = True
    If recTFlux.Statut = "A" And recTFlux.StatutPlus = "CU" Then mnuEchéancierEnCours.Enabled = True
    If recTFlux.Statut = "M" And recTFlux.StatutPlus = "an" Then mnuEchéancierEnCours.Enabled = True
End If

End Sub

Public Function saveTflux_Init()
Dim I As Integer
saveTflux_Init = "?"
If arrTFlux_Nb = 0 Then Call MsgBox("? Dossier sans échéances ", vbCritical, "frmGarantie.saveTflux_Init"): Exit Function

saveTFlux_Nb = arrTFlux_Nb
ReDim saveTFlux(saveTFlux_Nb)
saveTFlux(0) = arrTFlux(1)
saveTFlux(0).IdSéquence = arrTFlux(arrTFlux_Nb).IdSéquence

For I = 1 To arrTFlux_Nb
    saveTFlux(I) = arrTFlux(I)
    If saveTFlux(I).Statut = "à" Then Call MsgBox("? échéance en cours de traitement", vbCritical, "frmGarantie.saveTflux_Init"): Exit Function
Next I
saveTflux_Init = Null

End Function
Public Function saveTflux_AddNew()
Dim I As Integer
saveTflux_AddNew = "?"

For I = 1 To saveTFlux_Nb
    saveTFlux(I).Method = ""
    Select Case saveTFlux(I).Statut
        Case "à": Call MsgBox("? échéance en cours de traitement", vbCritical, "frmGarantie.saveTflux_Init"): Exit Function
        Case " ", "M"
            Select Case saveTFlux(I).CodeOpération
                Case "GA52", "GA62"
                    If saveTFlux(I).AmjEchéanceTrt >= wAMJEffet Then saveTFlux(I).Method = constDelete
                Case "GA02", "GA22": saveTFlux_AddNew_Validité I
           End Select
    End Select
Next I

For I = 1 To arrTFlux_Nb
    If arrTFlux(I).AmjEchéanceTrt >= wAMJEffet Then
        saveTFlux_Nb = saveTFlux_Nb + 1
        ReDim Preserve saveTFlux(saveTFlux_Nb + 1)
        saveTFlux(saveTFlux_Nb) = arrTFlux(I)
        saveTFlux(saveTFlux_Nb).Method = constAddNew
        saveTFlux(0).IdSéquence = saveTFlux(0).IdSéquence + 1
        saveTFlux(saveTFlux_Nb).IdSéquence = saveTFlux(0).IdSéquence
    End If
Next I

arrTFlux_Nb = saveTFlux_Nb
ReDim arrTFlux(arrTFlux_Nb + 1)
For I = 1 To arrTFlux_Nb
    arrTFlux(I) = saveTFlux(I)
Next I

saveTflux_AddNew = Null

End Function


Public Sub mnuListEchéancier_Display()

Call srvTFlux_Load(recTFlux, arrTFlux(0))
arrTFlux_Nb = srvTFlux.arrTFlux_Nb
ReDim arrTFlux(arrTFlux_Nb)
For I = 1 To arrTFlux_Nb
    arrTFlux(I) = srvTFlux.arrTFlux(I)
Next I
SSTab1.Tab = 4
fgEchéancier_Display " "

End Sub

Public Sub cmdControl_txtTaux()
    
If optComPériodiqueTaux Then
    recTope.TauxRéférence = ""
    X = num_Control(txtTaux, valX, 9, 5)
    recTope.TauxMarge = valX
    If recTope.TauxMarge <= 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le taux")
    If recTope.TauxMarge > 25 Then Call lstErr_AddItem(lstErr, cmdContext, "? taux > 25 %")
Else
    recTope.TauxRéférence = "Montant"
    X = num_Control(txtTaux, valX, 9, 2)
    recTope.TauxMarge = valX
End If
End Sub

Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect.Cols - 1
    arrTOpe_Index = Val(fgSelect.Text)
    fgSelect.Col = 15
    Select Case lK
        Case 1: fgSelect.Text = Format$(arrTOpe(arrTOpe_Index).Capital, "000000000000000.00") & arrTOpe(arrTOpe_Index).Devise
        Case 2: fgSelect.Text = arrTOpe(arrTOpe_Index).Devise & Format$(arrTOpe(arrTOpe_Index).Capital, "000000000000000.00")
        Case 16: fgSelect.Text = Format$(arrTOpe_Index, "0000000000")
    End Select
Next I

fgSelect_Sort1 = 15: fgSelect_Sort2 = 15
fgSelect_Sort
End Sub
Public Sub fgEchéancier_SortX(lK As Integer)
Dim I As Integer
For I = 1 To fgEchéancier.Rows - 1
    fgEchéancier.Row = I
    fgEchéancier.Col = fgEchéancier.Cols - 1
    arrTFlux_Index = Val(fgEchéancier.Text)
    fgEchéancier.Col = 10
    Select Case lK
        Case 2: fgEchéancier.Text = Format$(arrTFlux(arrTFlux_Index).Capital, "000000000000000.00")
        Case 3: fgEchéancier.Text = arrTFlux(arrTFlux_Index).AmjEchéanceTrt
        Case 9: fgEchéancier.Text = Format(recTFlux.IdRéférence, "000000") & "_" & Format(recTFlux.IdSéquence, "0000000")
        Case 11: fgEchéancier.Text = Format$(arrTFlux_Index, "0000000000")
    End Select
Next I

fgEchéancier_Sort1 = 10: fgEchéancier_Sort2 = 10
fgEchéancier_Sort
End Sub


Public Sub saveTFlux_AddNew_Validité(lI As Integer)

saveTflux_Index_GA02 = lI
saveTFlux(lI).Method = constUpdate
saveTFlux(lI).CptMvtUsr = usrId
saveTFlux(lI).CptMvtAMJ = DSys
saveTFlux(lI).CptMvtHMS = time_Hms

Select Case currentAction
    Case "Change_Validité"
            saveTFlux(lI).Statut = "A"
            saveTFlux(lI).StatutPlus = "01"
    Case "MainLevéePartielle"
            saveTFlux(lI).CodeOpération = "GA04"
            saveTFlux(lI).Statut = " "
            saveTFlux(lI).StatutPlus = "  "
            saveTFlux(lI).AmjEchéanceTrt = wAMJEffet
            saveTFlux(lI).AmjFin = wAMJEffet
            saveTFlux(lI).AmjEchéanceTrt = wAMJEffet
            saveTFlux(lI).AmjOpération = wAMJEffet
            saveTFlux(lI).AmjValeur = wAMJEffet
            saveTFlux(lI).Capital = mTope.Capital - recTope.Capital
    Case "MainLevée"
            saveTFlux(lI).CodeOpération = "GA03"
            saveTFlux(lI).Statut = " "
            saveTFlux(lI).StatutPlus = "  "
            saveTFlux(lI).AmjEchéanceTrt = wAMJEffet
            saveTFlux(lI).AmjFin = wAMJEffet
            saveTFlux(lI).AmjEchéanceTrt = wAMJEffet
            saveTFlux(lI).AmjOpération = wAMJEffet
            saveTFlux(lI).AmjValeur = wAMJEffet
            saveTFlux(lI).Capital = mTope.Capital
    Case "Augmentation"
            saveTFlux(lI).CodeOpération = "GA11"
            saveTFlux(lI).Statut = " "
            saveTFlux(lI).StatutPlus = "  "
            saveTFlux(lI).AmjEchéanceTrt = wAMJEffet
            saveTFlux(lI).AmjFin = wAMJEffet
            saveTFlux(lI).AmjEchéanceTrt = wAMJEffet
            saveTFlux(lI).AmjOpération = wAMJEffet
            saveTFlux(lI).AmjValeur = wAMJEffet
            saveTFlux(lI).Capital = recTope.Capital - mTope.Capital

End Select

End Sub

Public Sub cmdControl_txtAMJEffet_MainLevée()
txtAMJEffet_control
If wAMJEffet > DSys Then
    Call lstErr_AddItem(lstErr, cmdContext, "? date main levée > " & dateImp(DSys))
End If
If wAMJEffet <= recTope.AmjDébut Then
    Call lstErr_AddItem(lstErr, cmdContext, "? date main levée <= émission " & dateImp(recTope.AmjDébut))
End If
'If wAMJEffet >= recTope.AmjFin Then
'    Call lstErr_AddItem(lstErr, cmdContext, "? date main levée >= validité " & dateImp(recTope.AmjFin))
'End If
If wAMJEffet < paramAmjEngagementMin Then
    Call lstErr_AddItem(lstErr, cmdContext, "? date main levée < COMPTA Min " & dateImp(paramAmjEngagementMin))
End If

End Sub
