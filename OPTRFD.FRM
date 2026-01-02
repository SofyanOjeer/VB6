VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmOpTrfD 
   AutoRedraw      =   -1  'True
   Caption         =   "OpTrf : saisie des opérations"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6300
   ScaleWidth      =   9285
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   8860
      Picture         =   "OPTRFD.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   0
      Width           =   400
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Valider"
      Height          =   300
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CommandButton cmdPagePrior 
      Caption         =   "<<"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   17
      Top             =   0
      Width           =   2505
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Enregistrer"
      Height          =   300
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CommandButton cmdContext 
      Height          =   300
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1100
   End
   Begin VB.PictureBox picOpTrf 
      Height          =   510
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   9240
      TabIndex        =   14
      Top             =   360
      Width           =   9300
   End
   Begin VB.CommandButton cmdPageNext 
      Caption         =   ">>"
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
      Left            =   7600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox txtOpTrf 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5160
      MaxLength       =   6
      TabIndex        =   0
      Top             =   0
      Width           =   900
   End
   Begin VB.CommandButton cmdSuppress 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Supprimer"
      Height          =   300
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5340
      Left            =   0
      TabIndex        =   16
      Top             =   915
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   9419
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   8
      TabHeight       =   520
      ForeColor       =   8421376
      TabCaption(0)   =   "Saisie"
      TabPicture(0)   =   "OPTRFD.frx":0102
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraSolde"
      Tab(0).Control(1)=   "fraSwift71A"
      Tab(0).Control(2)=   "fraSwift32A"
      Tab(0).Control(3)=   "picMT100"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Suivi"
      TabPicture(1)   =   "OPTRFD.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picHistorique"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Virement"
      TabPicture(2)   =   "OPTRFD.frx":013A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picVirement"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Swift"
      TabPicture(3)   =   "OPTRFD.frx":0156
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picSwift"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Compta"
      TabPicture(4)   =   "OPTRFD.frx":0172
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "picCompta"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame fraSolde 
         BackColor       =   &H00E0E0E0&
         Height          =   4575
         Left            =   -70440
         TabIndex        =   84
         Top             =   480
         Width           =   4455
         Begin VB.TextBox txtSoldeAutorisation 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   95
            Top             =   2400
            Width           =   1935
         End
         Begin VB.TextBox txtSoldeDépassement 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   94
            Top             =   2040
            Width           =   1935
         End
         Begin VB.TextBox txtSoldeAprès 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   93
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox txtSoldeMontant 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2280
            TabIndex        =   92
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox txtSoldeAvant 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H0080FF80&
            Height          =   285
            Left            =   2280
            TabIndex        =   91
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtSoldeNuméro 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   90
            Top             =   240
            Width           =   2055
         End
         Begin VB.Frame fraDateValeur 
            BackColor       =   &H00E0E0E0&
            Height          =   1575
            Left            =   120
            TabIndex        =   85
            Top             =   2880
            Width           =   4215
            Begin VB.TextBox txtDateValeurDébit 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2160
               TabIndex        =   87
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox txtDateValeurCrédit 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2160
               TabIndex        =   86
               Top             =   960
               Width           =   1935
            End
            Begin VB.Label lblDateValeurDébit 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Date valeur DEBIT"
               Height          =   255
               Left            =   240
               TabIndex        =   89
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Date valeur CREDIT"
               Height          =   255
               Left            =   240
               TabIndex        =   88
               Top             =   960
               Width           =   1815
            End
         End
         Begin VB.Label lblSoldeAutorisation 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Autorisation"
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   2520
            Width           =   2055
         End
         Begin VB.Label lblSoldeDépassement 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dépassement"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label lblSoldeAprès 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Solde "
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
            TabIndex        =   99
            Top             =   1800
            Width           =   2055
         End
         Begin VB.Label lblSoldeMontant 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Montant de l'opération "
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   1440
            Width           =   2055
         End
         Begin VB.Label lblSoldeAvant 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Solde instantané"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Label lblSoldeIntitulé 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Intitulé"
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   600
            Width           =   4215
         End
      End
      Begin VB.Frame fraSwift71A 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imputation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1785
         Left            =   -74640
         TabIndex        =   73
         Top             =   840
         Width           =   2415
         Begin VB.OptionButton optImputation1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "bénéficiaire"
            Height          =   195
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Value           =   -1  'True
            Width           =   1605
         End
         Begin VB.OptionButton optImputation2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "partage"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   600
            Width           =   1600
         End
         Begin VB.OptionButton optImputation4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "pas de commission"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   1200
            Width           =   1965
         End
         Begin VB.OptionButton optImputation3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "donneur d'ordre"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   900
            Width           =   1600
         End
      End
      Begin VB.Frame fraSwift32A 
         BackColor       =   &H00E0E0E0&
         Height          =   2655
         Left            =   -74400
         TabIndex        =   51
         Top             =   3120
         Width           =   5175
         Begin VB.Frame fraMT100_CompteàCréditer 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   4850
            Begin VB.TextBox txtSwift32DeviseCr 
               Height          =   300
               Left            =   1440
               MaxLength       =   3
               TabIndex        =   54
               Top             =   600
               Width           =   405
            End
            Begin VB.TextBox txtSwift32Montant 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1920
               MaxLength       =   20
               TabIndex        =   55
               Top             =   600
               Width           =   1500
            End
            Begin VB.TextBox txtSwift32CoursCr 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   3480
               TabIndex        =   56
               Top             =   600
               Width           =   1250
            End
            Begin MSMask.MaskEdBox txtSwift32AmjValeurCr 
               Height          =   300
               Left            =   120
               TabIndex        =   53
               Top             =   600
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   14
               Mask            =   "## - ## - ####"
               PromptChar      =   "_"
            End
            Begin VB.Label lblSwift32AmjValeur 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Crédit (valeur)"
               Height          =   255
               Left            =   120
               TabIndex        =   72
               Top             =   240
               Width           =   1100
            End
            Begin VB.Label lblSwift32DeviseCr 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Devise"
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
               TabIndex        =   71
               Top             =   240
               Width           =   615
            End
            Begin VB.Label lblSwift32Montant 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Montant"
               Height          =   255
               Left            =   2280
               TabIndex        =   70
               Top             =   240
               Width           =   750
            End
            Begin VB.Label lblSwift32CoursDb 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cours"
               Height          =   255
               Left            =   3720
               TabIndex        =   69
               Top             =   240
               Width           =   750
            End
            Begin VB.Label lblSwift32CompteCr 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Compte Nostro"
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
               TabIndex        =   68
               Top             =   1000
               Width           =   1245
            End
            Begin VB.Label libSwift32CompteC 
               BackColor       =   &H00E0E0E0&
               Caption         =   "-"
               Height          =   255
               Left            =   3600
               TabIndex        =   67
               Top             =   960
               Width           =   1140
            End
            Begin VB.Label libSwift32IbanBicId 
               BackColor       =   &H00E0E0E0&
               Caption         =   "-"
               Height          =   255
               Left            =   1440
               TabIndex        =   66
               Top             =   1005
               Width           =   1800
            End
         End
         Begin VB.Frame fraMT100_CompteàDébiter 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   1035
            Left            =   120
            TabIndex        =   52
            Top             =   1440
            Width           =   4850
            Begin VB.TextBox txtSwift32MontantDb 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   1920
               TabIndex        =   59
               Top             =   600
               Width           =   1500
            End
            Begin VB.TextBox txtSwift32CompteDb 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   300
               Left            =   3480
               TabIndex        =   60
               Top             =   600
               Width           =   1250
            End
            Begin VB.TextBox txtSwift32DeviseDb 
               Height          =   300
               Left            =   1440
               MaxLength       =   3
               TabIndex        =   58
               Top             =   600
               Width           =   405
            End
            Begin MSMask.MaskEdBox txtSwift32AmjValeurDb 
               Height          =   300
               Left            =   120
               TabIndex        =   57
               Top             =   600
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   14
               Mask            =   "## - ## - ####"
               PromptChar      =   "_"
            End
            Begin VB.Label lblSwift32CompteDb 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Compte"
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
               Left            =   3720
               TabIndex        =   64
               Top             =   240
               Width           =   615
            End
            Begin VB.Label lblSwift32DeviseDb 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Devise"
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
               TabIndex        =   63
               Top             =   240
               Width           =   615
            End
            Begin VB.Label lblSwift32MontantDb 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Brut à débiter"
               Height          =   255
               Left            =   2280
               TabIndex        =   62
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label lblSwift32AmjValeurDb 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Débit (valeur)"
               Height          =   255
               Left            =   120
               TabIndex        =   61
               Top             =   240
               Width           =   1095
            End
         End
      End
      Begin VB.PictureBox picSwift 
         AutoRedraw      =   -1  'True
         Height          =   4695
         Left            =   -74880
         ScaleHeight     =   4635
         ScaleWidth      =   8955
         TabIndex        =   24
         Top             =   480
         Width           =   9015
      End
      Begin VB.PictureBox picVirement 
         AutoRedraw      =   -1  'True
         Height          =   4695
         Left            =   -74880
         ScaleHeight     =   5000
         ScaleMode       =   0  'User
         ScaleWidth      =   11100
         TabIndex        =   23
         Top             =   480
         Width           =   9015
      End
      Begin VB.PictureBox picCompta 
         AutoRedraw      =   -1  'True
         Height          =   4695
         Left            =   120
         ScaleHeight     =   4635
         ScaleWidth      =   8955
         TabIndex        =   22
         Top             =   480
         Width           =   9015
         Begin VB.Frame Frame1 
            Height          =   1455
            Left            =   240
            TabIndex        =   43
            Top             =   3120
            Width           =   8175
            Begin VB.TextBox txtOpCpt1Libellé 
               Height          =   285
               Left            =   120
               MaxLength       =   50
               TabIndex        =   46
               Top             =   720
               Width           =   7815
            End
            Begin VB.TextBox txtOpCpt2Libellé 
               Height          =   285
               Left            =   120
               MaxLength       =   50
               TabIndex        =   45
               Top             =   1080
               Width           =   7815
            End
            Begin VB.CheckBox chkLibelléManuel 
               Caption         =   "Modifier les libellés"
               Height          =   255
               Left            =   120
               TabIndex        =   44
               Top             =   240
               Width           =   2055
            End
         End
      End
      Begin VB.PictureBox picHistorique 
         AutoRedraw      =   -1  'True
         Height          =   4695
         Left            =   -74880
         ScaleHeight     =   4635
         ScaleWidth      =   8955
         TabIndex        =   21
         Top             =   480
         Width           =   9015
      End
      Begin VB.PictureBox picMT100 
         AutoRedraw      =   -1  'True
         Height          =   4710
         Left            =   -74880
         ScaleHeight     =   4650
         ScaleWidth      =   8955
         TabIndex        =   19
         Top             =   480
         Width           =   9015
         Begin VB.Frame fraSwift54 
            BackColor       =   &H00C0C0C0&
            Caption         =   "54 : Correspondant"
            Height          =   735
            Left            =   840
            TabIndex        =   78
            Top             =   3480
            Visible         =   0   'False
            Width           =   3495
            Begin VB.TextBox txtSwift54 
               Height          =   285
               Left            =   1320
               MaxLength       =   11
               TabIndex        =   80
               Top             =   240
               Width           =   1935
            End
            Begin VB.CheckBox chkSwift54 
               Caption         =   "Swift Via"
               Height          =   255
               Left            =   120
               TabIndex        =   79
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame fraSwift00 
            Height          =   1440
            Left            =   120
            TabIndex        =   47
            Top             =   2160
            Width           =   1815
            Begin VB.OptionButton optMT202 
               Caption         =   "MT202"
               Height          =   375
               Left            =   240
               TabIndex        =   50
               Top             =   620
               Width           =   1095
            End
            Begin VB.OptionButton optMt100 
               Caption         =   "MT100"
               Height          =   255
               Left            =   240
               TabIndex        =   49
               Top             =   320
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton optMt210 
               Caption         =   "MT210"
               Height          =   375
               Left            =   240
               TabIndex        =   48
               Top             =   960
               Width           =   1095
            End
         End
         Begin VB.ListBox lstDevise 
            Height          =   4350
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Visible         =   0   'False
            Width           =   2685
         End
         Begin VB.Frame fraRib 
            Caption         =   "R I B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   1455
            Left            =   2880
            TabIndex        =   25
            Top             =   3120
            Width           =   6075
            Begin VB.Frame fraVirement 
               Height          =   1215
               Left            =   4560
               TabIndex        =   37
               Top             =   120
               Width           =   1455
               Begin VB.OptionButton optAutre 
                  Caption         =   "Autre"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   41
                  Top             =   870
                  Width           =   1200
               End
               Begin VB.OptionButton optCRI 
                  Caption         =   "CRI"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   40
                  Top             =   600
                  Width           =   1200
               End
               Begin VB.OptionButton optVSOT 
                  Caption         =   "VSOT"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   39
                  Top             =   370
                  Width           =   1095
               End
               Begin VB.OptionButton optAristote 
                  Caption         =   "Aristote"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   38
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   1215
               End
            End
            Begin VB.TextBox txtfraRibIban 
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               MaxLength       =   4
               TabIndex        =   33
               Text            =   "FR00"
               Top             =   400
               Width           =   500
            End
            Begin VB.CheckBox chkRib 
               Caption         =   "Clé RIB"
               Height          =   195
               Left            =   3720
               TabIndex        =   31
               Top             =   120
               Value           =   1  'Checked
               Width           =   855
            End
            Begin VB.TextBox txtfraRibEtablissement 
               Height          =   285
               Left            =   840
               MaxLength       =   5
               TabIndex        =   5
               Top             =   400
               Width           =   600
            End
            Begin VB.TextBox txtfraRibGuichet 
               Height          =   285
               Left            =   1560
               MaxLength       =   5
               TabIndex        =   6
               Top             =   400
               Width           =   615
            End
            Begin VB.TextBox txtfraRibCompte 
               Height          =   285
               Left            =   2400
               MaxLength       =   11
               TabIndex        =   7
               Top             =   400
               Width           =   1200
            End
            Begin VB.TextBox txtfraRibClé 
               Height          =   285
               Left            =   3720
               MaxLength       =   2
               TabIndex        =   8
               Top             =   400
               Width           =   350
            End
            Begin VB.Label lblfraRibIban 
               Caption         =   "IBAN"
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Top             =   200
               Width           =   400
            End
            Begin VB.Label libfraRibDom 
               Caption         =   "-"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   720
               Width           =   3975
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
               Left            =   2640
               TabIndex        =   29
               Top             =   120
               Width           =   735
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
               Left            =   1560
               TabIndex        =   28
               Top             =   120
               Width           =   855
            End
            Begin VB.Label lblfraRibBanque 
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
               Left            =   840
               TabIndex        =   27
               Top             =   120
               Width           =   720
            End
            Begin VB.Label libfraRibCompte 
               Caption         =   "-"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   1080
               Width           =   3945
            End
         End
         Begin VB.Frame fraMT100 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   3015
            Left            =   2880
            TabIndex        =   20
            Top             =   120
            Width           =   6135
            Begin VB.CommandButton cmdBicIban 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Iban..."
               Height          =   400
               Left            =   5400
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   2400
               Width           =   600
            End
            Begin VB.CommandButton cmdScan 
               BackColor       =   &H00E0E0E0&
               Height          =   400
               Left            =   5400
               Picture         =   "OPTRFD.frx":018E
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   1680
               Width           =   600
            End
            Begin VB.CommandButton fraMT100_cmdQuit 
               BackColor       =   &H00C0C0FF&
               Caption         =   "X"
               Height          =   400
               Left            =   5400
               Style           =   1  'Graphical
               TabIndex        =   1
               Top             =   480
               Width           =   600
            End
            Begin VB.CommandButton fraMt100_cmdOk 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Ok"
               Height          =   400
               Left            =   5400
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   1080
               Width           =   600
            End
            Begin VB.TextBox txtMT100f 
               Height          =   285
               Left            =   500
               MaxLength       =   32
               TabIndex        =   13
               Top             =   2500
               Width           =   4500
            End
            Begin VB.TextBox txtMT100e 
               Height          =   285
               Left            =   500
               MaxLength       =   32
               TabIndex        =   12
               Top             =   2100
               Width           =   4500
            End
            Begin VB.TextBox txtMT100d 
               Height          =   285
               Left            =   500
               MaxLength       =   32
               TabIndex        =   11
               Top             =   1700
               Width           =   4500
            End
            Begin VB.TextBox txtMT100c 
               Height          =   285
               Left            =   500
               MaxLength       =   32
               TabIndex        =   10
               Top             =   1300
               Width           =   4500
            End
            Begin VB.TextBox txtMT100a 
               Height          =   285
               Left            =   480
               MaxLength       =   32
               TabIndex        =   4
               Top             =   600
               Width           =   4500
            End
            Begin VB.TextBox txtMT100b 
               Height          =   285
               Left            =   480
               MaxLength       =   32
               TabIndex        =   9
               Top             =   960
               Width           =   4500
            End
         End
         Begin VB.Shape shapeMT100 
            BackColor       =   &H00C0C0C0&
            BorderColor     =   &H00FF0000&
            FillColor       =   &H00C0C0C0&
            Height          =   200
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   8920
         End
      End
   End
   Begin VB.Menu fraRibmnu 
      Caption         =   "fraRibmnu"
      Visible         =   0   'False
      Begin VB.Menu fraRibmnufrmBdf 
         Caption         =   "Recherche Banque de France"
      End
      Begin VB.Menu fraRibmnufrmCompte 
         Caption         =   "Recherche Compte"
      End
   End
End
Attribute VB_Name = "frmOpTrfD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim x As String, I As Integer
'Dim recBicIban As typeBicIban

Dim picMT100Index As Integer, picMT100NbMax As Integer, picMT100NbMin As Integer
Dim wNature As String * 2
Dim valRéférence As String * 6, valLibellé As String, valBen As String
Dim currentAction As String, blnAddNew As Boolean, blnGlobalControl As Boolean
Dim blntxtBox_Enabled As Boolean, blnMsgBox_Quit As Boolean, blnAuto_cmdok_Click As Boolean
Dim chkLevel As String * 1
Dim blnControl As Boolean, blnLibelléManuel As Boolean, blnMsgBox As Boolean
Dim blnfraSolde As Boolean

Dim CV1 As typeCV, CV2 As typeCV, CV3 As typeCV
Dim CV_Conversion As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor


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

Public Sub XopSwift_Update(ByVal X2 As String)
Dim blnTest As Boolean, I As Integer
blnTest = True
arrOpSwiftNb = 0
For I = 1 To XopSwiftNb
    If mId$(XopSwift(I).FieldNature, 1, 2) <> X2 Then
        arrOpSwift_AddItem XopSwift(I)
    Else
        If blnTest Then
            blnTest = False
            If X2 = "00" Then
                arrOpSwift_AddItem XopSwift(I)
                Mid$(arrOpSwift(I).FieldText, 1, 11) = mId$(Trim(txtMT100a), 1, 11)
            Else
                If X2 = "32" Then
                    arrOpSwift_AddItem XopSwift(I)
                     arrOpSwift(I).FieldText = txtSwift32AmjValeurCr & " " & XopDevise(2).DevX & " " _
                    & num_Display(XopCpt(2).Brut, 15, XopDevise(2).maxD, Lx, x, "0")
                Else
                    arrOpSwift_AddItem XopSwift(I)
                    arrOpSwift(arrOpSwiftNb).FieldText = Trim(txtMT100a)
                    If Trim(txtMT100b) <> "" Then
                        arrOpSwift_AddItem XopSwift(I)
                        arrOpSwift(arrOpSwiftNb).FieldText = Trim(txtMT100b)
                    End If
                    If Trim(txtMT100c) <> "" Then
                        arrOpSwift_AddItem XopSwift(I)
                        arrOpSwift(arrOpSwiftNb).FieldText = Trim(txtMT100c)
                    End If
                    If Trim(txtMT100d) <> "" Then
                        arrOpSwift_AddItem XopSwift(I)
                        arrOpSwift(arrOpSwiftNb).FieldText = Trim(txtMT100d)
                    End If
                    If Trim(txtMT100e) <> "" Then
                        arrOpSwift_AddItem XopSwift(I)
                        arrOpSwift(arrOpSwiftNb).FieldText = Trim(txtMT100e)
                    End If
                    If Trim(txtMT100f) <> "" Then
                        arrOpSwift_AddItem XopSwift(I)
                        arrOpSwift(arrOpSwiftNb).FieldText = Trim(txtMT100f)
                    End If
                End If
            End If
        End If
    End If
Next I

XopSwift_Load
End Sub


Private Sub chkLibelléManuel_Click()
If chkLibelléManuel = "1" Then
    txtOpCpt1Libellé.Enabled = True
    txtOpCpt2Libellé.Enabled = True
    txtOpCpt1Libellé.SetFocus
Else
    txtOpCpt1Libellé.Enabled = False
    txtOpCpt2Libellé.Enabled = False
End If

End Sub

Private Sub chkLibelléManuel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set chkLibelléManuel

End Sub


Private Sub chkRib_Click()
If chkRib = 0 Then
    txtfraRibClé.Enabled = False: txtfraRibClé = "00"
    If optAristote = "1" Then optVSOT = "1": optAristote.Enabled = False
             optCRI.Enabled = True ': optCRI = "1"
Else
    txtfraRibClé.Enabled = blntxtBox_Enabled: optAristote.Enabled = True
End If
End Sub

Private Sub chkRib_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then Call fraRib_MouseDown(Button, Shift, x, Y)
End Sub

Private Sub chkRib_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set chkRib

End Sub

Private Sub chkSwift54_Click()
chkSwift54_Control
If chkSwift54 = "1" Then: If txtSwift54.Visible Then txtSwift54.SetFocus
End Sub

Private Sub chkSwift54_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set chkSwift54

End Sub


Private Sub cmdBicIban_Click()
arrBic(0).Method = "IbanDevise"
arrBicIban(0).IbanDevise = XopCpt(2).Devise
arrBicIban(0).BicId = XopBicReceiver.BicId
arrBicIban(0).IbanBicId = ""
frmBicList.Show vbModal
If Trim(arrBicIban(0).Method) = "Ok" Then
    Mid$(XopSwift(2).FieldText, 13, 11) = arrBicIban(0).IbanBicId
    XopBicIbanReceiver = arrBicIban(0)
    Swift57_ControlIbanReceiver
End If

End Sub

Private Sub cmdBicIban_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
blnControl = False
End Sub


Private Sub cmdBicIban_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdBicIban

End Sub


'Private Sub chkSwift54_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MouseMoveActiveControl_Set chkSwift54

'End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdOk_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdOk

End Sub

Private Sub cmdPageNext_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdPageNext

End Sub


Private Sub cmdPagePrior_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdPagePrior

End Sub


Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdSave

End Sub

Private Sub cmdScan_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
blnControl = False
End Sub

Private Sub cmdScan_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdScan

End Sub

Private Sub cmdSuppress_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdSuppress

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
blnControl = False
If blnMsgBox_Quit Then
    If MsgBox("Voulez-vous vraiment quitter la saisie 'Opération de transfert' ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Cancel = True
End If

End Sub

Private Sub fraDateValeur_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraMt100_cmdOk_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set fraMt100_cmdOk
End Sub

Private Sub fraMT100_cmdQuit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
blnControl = False
End Sub

Private Sub fraMT100_cmdQuit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set fraMT100_cmdQuit

End Sub

Private Sub fraMT100_CompteàCréditer_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraMT100_CompteàDébiter_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraMT100_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraRib_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub fraSolde_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraSwift00_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraSwift32A_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraSwift54_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraSwift71A_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraVirement_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub lblfraRibBanque_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
fraRibmnufrmBdf_Click
End Sub


Private Sub lblfraRibBanque_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set lblfraRibBanque

End Sub


Private Sub lblfraRibCompte_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
fraRibmnufrmCompte_Click
'Call fraRib_MouseDown(Button, Shift, X, Y)
End Sub


Private Sub lblfraRibCompte_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set lblfraRibCompte

End Sub


Private Sub lblfraRibGuichet_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
fraRibmnufrmBdf_Click
'Call fraRib_MouseDown(Button, Shift, X, Y)
End Sub




Private Sub lblfraRibGuichet_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set lblfraRibGuichet

End Sub


Private Sub lblSwift32CompteCr_Click()
cmdScan_Click
End Sub

Private Sub lblSwift32CompteCr_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set lblSwift32CompteCr

End Sub


Private Sub lblSwift32CompteDb_Click()
currentActiveControl_Name = "txtSwift32CompteDb"
fraRibmnufrmCompte_Click
End Sub


Private Sub lblSwift32CompteDb_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set lblSwift32CompteDb

End Sub


Private Sub lblSwift32DeviseCr_Click()
currentActiveControl_Name = txtSwift32DeviseCr.Name
lstDevise_Display
End Sub

Private Sub lblSwift32DeviseCr_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set lblSwift32DeviseCr

End Sub


Private Sub lblSwift32DeviseDb_Click()
currentActiveControl_Name = txtSwift32DeviseDb.Name
lstDevise_Display
End Sub


Private Sub lblSwift32DeviseDb_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set lblSwift32DeviseDb

End Sub


Private Sub libfraRibCompte_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call fraRib_MouseDown(Button, Shift, x, Y)
End Sub

Private Sub libfraRibDom_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call fraRib_MouseDown(Button, Shift, x, Y)
End Sub

Private Sub lstDevise_Click()
Select Case currentActiveControl_Name
    Case "txtSwift32DeviseDb": txtSwift32DeviseDb = mId$(lstDevise.Text, 1, 3): Swift32DeviseDb_Control txtSwift32DeviseDb: txtSwift32DeviseDb.SetFocus
    Case "txtSwift32DeviseCr": txtSwift32DeviseCr = mId$(lstDevise.Text, 1, 3): Swift32DeviseCr_Control txtSwift32DeviseCr: txtSwift32DeviseCr.SetFocus
End Select
lstDevise.Visible = False
cmdContext_Return
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


'--------------------------------------------------
Public Sub txtfraRibClé_Control()
'--------------------------------------------------
Dim IbanE As String
x = Trim(txtfraRibCompte)
If x = "" Then Exit Sub
    I = RibClé(txtfraRibEtablissement, txtfraRibGuichet, x, IbanE)
    If I <> Val(txtfraRibClé) Then
        x = "? clé rib : " & Format$(I, "00")
        Call lstErr_AddItem(lstErr, txtfraRibClé, x)
    End If
End Sub

Private Sub picMT202_Click()

End Sub

Private Sub optAristote_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optAristote

End Sub

Private Sub optAutre_Click()
'Call fraMT100_ControlFRF(NostroBicADB, "3", "2")
End Sub

Private Sub optAristote_Click()
Call fraMT100_ControlFRF(NostroBicVir, "0", "1")
End Sub

Private Sub optAutre_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optAutre

End Sub

Private Sub optCRI_Click()
Call fraMT100_ControlFRF(NostroBicCRI, "3", "2")
End Sub

Private Sub optCRI_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optCRI

End Sub


Private Sub optImputation1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optImputation1

End Sub


Private Sub optImputation2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optImputation2

End Sub


Private Sub optImputation3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optImputation3

End Sub


Private Sub optImputation4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optImputation4

End Sub


Private Sub optMt100_Click()
If txtMT100a.Enabled Then txtMT100a.SetFocus
End Sub

Private Sub optMt100_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optMt100

End Sub


Private Sub optMT202_Click()
If txtMT100a.Enabled Then txtMT100a.SetFocus
End Sub


Private Sub optMT202_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optMT202

End Sub


Private Sub optMt210_Click()
If txtMT100a.Enabled Then txtMT100a.SetFocus
End Sub

Private Sub optMt210_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optMt210

End Sub


Private Sub optVSOT_Click()
Call fraMT100_ControlFRF(NostroBicVir, "0", "1")

End Sub

Private Sub optVSOT_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set optVSOT
End Sub


Private Sub SSTab_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set SSTab

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
Dim IbanE As String
libfraRibCompte = ""
If Trim(txtfraRibCompte) = "" Then
    If optCRI Then
        txtfraRibCompte = "": Exit Sub
    Else
        Call lstErr_AddItem(lstErr, txtfraRibCompte, "? Compte"): Exit Sub
    End If
End If

txtfraRibCompte = Rib_Compte(txtfraRibCompte)
 If txtfraRibEtablissement = strSocBdfE Then
   
    If Not IsNumeric(txtfraRibCompte) Then
        Call lstErr_AddItem(lstErr, txtfraRibCompte, "? compte BIA")
    Else
            recCompteInit recCompte
            recCompte.Method = "SeekL1"
            recCompte.Société = SocId$
            recCompte.Agence = SocAgence$
            recCompte.Devise = XopCpt(2).Devise
            recCompte.Numéro = Format$(RTrim(txtfraRibCompte), "00000000000")
            If IsNull(srvCompteMon(recCompte)) Then
                libfraRibCompte = recCompte.Intitulé
                If recCompte.Situation = "A" Or recCompte.Situation = "E" Then
                    Call lstErr_AddItem(lstErr, txtfraRibCompte, "? compte annulé")
                Else
                    txtfraRibClé = Format$(RibClé(txtfraRibEtablissement, txtfraRibGuichet, Trim(txtfraRibCompte), IbanE), "00")
                    XopCpt(2).Compte = recCompte.Numéro
                    XopCompteNostro = recCompte
 '990112            If Not blnGlobalControl And Trim(txtMT100a) = "" Then txtMT100a = recCompte.Intitulé
                    If Not blnGlobalControl Then txtMT100a = recCompte.Intitulé
                End If
            Else
                Call lstErr_AddItem(lstErr, txtfraRibCompte, "? compte inconnu")
        End If
    End If
End If
End Sub

Private Sub txtfraRibClé_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then Call fraRib_MouseDown(Button, Shift, x, Y)
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
Dim I As Integer, X11 As String * 11
x = Trim(txtfraRibCompte)
I = Len(x)
X11 = "00000000000": If I > 0 Then Mid$(X11, 12 - I, I) = x
txtfraRibCompte.Text = X11
Call txt_LostFocus(txtfraRibCompte)
End Sub

'--------------------------------------------------
Public Sub txtfraRibEtablissement_Control()
'--------------------------------------------------
If Trim(txtfraRibEtablissement) = "" Then Call lstErr_AddItem(lstErr, txtfraRibEtablissement, "? Etablissement")
fraVirement.Visible = IIf(txtfraRibEtablissement <> strSocBdfE, True, False)
End Sub

Private Sub txtfraRibCompte_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then Call fraRib_MouseDown(Button, Shift, x, Y)
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
If Trim(txtfraRibGuichet) = "" Then 'Or txtfraRibGuichet = "00000" Then
    If optCRI Then
        libfraRibDom = "non précisé"
        txtfraRibGuichet = ""
        Exit Sub
    Else
        Call lstErr_AddItem(lstErr, txtfraRibGuichet, "? Guichet"): recBdfG.BdfG2 = Format$(Val(txtfraRibEtablissement), "00000"): Exit Sub
    End If
End If
recBdfG.BdfG2 = Format$(Val(txtfraRibEtablissement), "00000")
recBdfG.BdfG3 = Format$(Val(txtfraRibGuichet), "00000")
If Not IsNull(srvBdfGFind(recBdfG)) Then
    Call lstErr_AddItem(lstErr, txtfraRibGuichet, "? guichet inconnu")
    libfraRibDom = "?"
Else
    libfraRibDom = recBdfG.BdfG20
End If

End Sub

Private Sub txtfraRibEtablissement_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then Call fraRib_MouseDown(Button, Shift, x, Y)
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

Private Sub cmdOk_Click()
If frmOpTrfD_GlobalControlàValider Then
    Me.Enabled = False: blnOpTrf_Update_on = True
    XopEch_Action currentAction
    If blnAddNew Then
        x = srvOpTrf_dbAddNew
    Else
        x = srvOpTrf_dbUpdate
    End If
    frmOpTrfD_Clear
    If Trim(x) <> "" Then
        Call lstErr_Clear(lstErr, txtOpTrf, x)
    Else
        MsgBox "Erreur pendant la mise à jour AS400", vbCritical, " Opération de transfert"
    End If
    Me.Enabled = True: blnOpTrf_Update_on = False
    AppActivate Me.Caption

End If
End Sub

Private Sub fraMT100_cmdOk_Click()
Dim blnTest As Boolean, X2 As String
lstErr.Clear: arrTag_Set False
fraMT100_Control wNature
If lstErr.ListCount > 0 Then Exit Sub

blnMsgBox_Quit = True
XopSwift_Update wNature
Call XopSwift_Display(picMT100): XopTrf_Display picOpTrf
blnTest = False
If wNature = "00" Then wNature = "02"
For XopSwiftIndex = 1 To XopSwiftNb
    X2 = mId$(XopSwift(XopSwiftIndex).FieldNature, 1, 2)
    If X2 > wNature Then
        If mId$(XopSwift(1).FieldText, 12, 1) = " " Then
            blnTest = True:    Exit For
        Else
            If X2 = "20" Or X2 = "32" Or X2 = "57" Or X2 = "59" Then blnTest = True: Exit For
        End If
    End If
Next XopSwiftIndex
picMT100Index = XopSwiftIndex - picMT100NbMin - 1
If blnTest Then
    picMT100_Line
Else
    fraMT100.Visible = False
    fraRib.Visible = False
End If
End Sub

Private Sub fraMT100_cmdQuit_Click()
blnControl = False
lstErr.Clear: lstErr.Visible = False: arrTag_Set False
fraMT100.Visible = False
fraRib.Visible = False
fraSwift54.Visible = False
fraSwift00.Visible = False
XopSwift(1) = MopSwift(1)
XopSwift(2) = MopSwift(2)
XopCpt(1) = mOpCpt(1)
XopCpt(2) = mOpCpt(2)
XopBic_Load
Call XopSwift_Display(picMT100): XopTrf_Display picOpTrf
fraSolde_Display
End Sub




Private Sub picMT100_Line()
shapeMT100.Top = (picMT100Index - 1) * lstLineHeight - 20
shapeMT100.Height = lstLineHeight
shapeMT100.Visible = True
XopSwiftIndex = picMT100Index + picMT100NbMin
If XopSwiftIndex > 1 Then XopSwiftIndex = XopSwiftIndex + 1
If XopSwiftIndex > XopSwiftNb Then XopSwiftIndex = XopSwiftNb
wNature = XopSwift(XopSwiftIndex).FieldNature
fraMT100_Init wNature
End Sub

'--------------------------------------------------
Private Sub picMT100_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'--------------------------------------------------
If fraMT100.Visible Then fraMT100_cmdQuit_Click
picMT100Index = (Y - 20) \ lstLineHeight + 1
picMT100_Line
End Sub


'---------------------------------------------------------
Private Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub
Private Sub fraRib_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtfraRibCompte.SetFocus
arrTag_Set False
fraRibmnufrmBdf.Enabled = True
fraRibmnufrmCompte.Enabled = IIf(txtfraRibEtablissement = strSocBdfE, True, False)
'Me.PopupMenu fraRibmnu, vbPopupMenuRightButton
End Sub


Private Sub fraRibmnufrmBdf_Click()
x = Space$(100)
Mid$(x, 1, 12) = "frmBdf      "
Mid$(x, 13, 12) = "frmOpTrfD   "
Mid$(x, 25, 10) = Space$(10)
Mid$(x, 35, 5) = Format$(Val(txtfraRibEtablissement), "00000")
Mid$(x, 40, 5) = Format$(Val(txtfraRibGuichet), "00000")
Mid$(x, 45, 11) = txtfraRibCompte
Mid$(x, 56, 2) = Format$(Val(txtfraRibClé), "00")
Msg_Monitor x

End Sub


Private Sub fraRibmnufrmCompte_Click()
Dim Y As String
x = Space$(100)
Mid$(x, 1, 12) = "frmCompte   "
Mid$(x, 13, 12) = "frmOpTrfD  "
Mid$(x, 25, 10) = Space$(10)
Select Case currentActiveControl_Name
    Case "txtfraRibCompte":
        If IsNumeric(txtfraRibCompte) Then
            Mid$(x, 38, 11) = Format$(Val(txtfraRibCompte), "###########")
        Else
            Mid$(x, 38, 11) = Format$(RTrim(txtfraRibCompte), "@@@@@@@@@@@")
        End If
        Mid$(x, 35, 3) = XopCpt(2).Devise
    Case "txtSwift32CompteDb"
        If IsNumeric(txtSwift32CompteDb) Then
            Mid$(x, 38, 11) = Format$(Val(txtSwift32CompteDb), "###########")
        Else
            Mid$(x, 38, 11) = Format$(RTrim(txtSwift32CompteDb), "@@@@@@@@@@@")
        End If
        Mid$(x, 35, 3) = XopCpt(1).Devise
End Select
Msg_Monitor x
End Sub



'---------------------------------------------------------
Private Sub frmOpTrfD_Clear()
'---------------------------------------------------------
picMT100NbMin = 0
XopDevise(1).DevLib = "": XopDevise(2).DevLib = ""
recBdfGInit recBdfG
recCompteInit recCompte: XopCompteNostro = recCompte: XopCompteSender = recCompte
recBicIbanInit XopBicIbanSender: XopBicIbanReceiver = XopBicIbanSender: XopBicIbanNostro = XopBicIbanSender
recBicInit XopBicSender: XopBicReceiver = XopBicSender
blnControl = True
blntxtBox_Enabled = False
blnLibelléManuel = False

blnMsgBox_Quit = False
usrColor_Set
fraMT100.BackColor = &HC0C0C0
fraSwift32A.BackColor = &HC0C0C0
'picOpTrf.BackColor = vbWindowBackground
shapeMT100.Visible = False
shapeMT100.Height = lstLineHeight

cmdContext.Caption = "&Rechercher"
cmdContext.Enabled = True
cmdSave.Visible = False
cmdSuppress.Visible = False
cmdOk.Visible = False
txtOpTrf.Enabled = True: If txtOpTrf.Visible And txtOpTrf.Enabled Then txtOpTrf.SetFocus
SSTab.Enabled = False
fraMT100.Visible = False
fraMt100_cmdOk.Visible = False
fraSolde.Visible = False

txtOpTrf = ""
valLibellé = "": valBen = ""
libSwift32CompteC = "": libSwift32IbanBicId = ""

XopSwiftNb = 0: XopEchNb = 0: XopCptNb = 0
chkLevel = "0"

picOpTrf.Cls
picMT100.Cls
picCompta.Cls
picHistorique.Cls
picSwift.Cls
picVirement.Cls
fraMT100_Clear

arrTag_Set False
lstErrClear = True
Call lstErr_Clear(lstErr, txtOpTrf, "dossier")

End Sub

'---------------------------------------------------------
Private Sub frmOpTrfD_Control()
'---------------------------------------------------------
Dim I As Integer, xobjControl As Boolean
If lstErrClear Then lstErr.Clear

For I = 0 To arrTagNb
    If arrTag(I) Then
        xobjControl = False
        arrTag(I) = False
        x = Format(I, "000")
        For Each xobj In Me.Controls
            If x = xobj.Tag Then
                Select Case xobj.Name
                    Case "txtMT100a":
                            If wNature = "57" Then Swift57_Control txtMT100a: xobjControl = True
                            If wNature = "59" Then Swift59_Control txtMT100a: xobjControl = True
                    Case "txtSwift32AmjValeurCr": Swift32AmjValeurCr_Control txtSwift32AmjValeurCr: xobjControl = True
                    Case "txtSwift32DeviseCr": Swift32DeviseCr_Control txtSwift32DeviseCr: xobjControl = True
                    Case "txtSwift32CoursCr": Swift32CoursCr_Control txtSwift32CoursCr: xobjControl = True
                    Case "txtSwift32Montant": Swift32Montant_Control txtSwift32Montant: xobjControl = True
                    Case "txtSwift32AmjValeurDb": Swift32AmjValeurdb_Control txtSwift32AmjValeurDb: xobjControl = True
                    Case "txtSwift32MontantDb": Swift32MontantDb_Control txtSwift32MontantDb: xobjControl = True
                    Case "txtSwift32CompteDb": Swift32CompteDb_Control txtSwift32CompteDb: xobjControl = True
                    Case "txtSwift32DeviseDb": Swift32DeviseDb_Control txtSwift32DeviseDb: xobjControl = True
                    Case "txtfraRibEtablissement": txtfraRibEtablissement_Control: xobjControl = True
                    Case "txtfraRibGuichet": txtfraRibGuichet_Control: xobjControl = True
                    Case "txtfraRibCompte": txtfraRibCompte_Control: xobjControl = True
                    Case "txtfraRibClé": txtfraRibClé_Control: xobjControl = True
                    Case "txtSwift54": Swift54_Control txtSwift54: xobjControl = True
                End Select
            End If
            If xobjControl Then Exit For
        Next xobj
    End If
Next I

lstErrClear = True
End Sub



'---------------------------------------------------------
Public Function frmOpTrfD_GlobalControl() As Boolean
'---------------------------------------------------------
Dim I As Integer
blnGlobalControl = True
lstErr.Clear
Swift32A_Init
For I = 1 To XopSwiftNb
    If XopSwift(I).FieldLine = 1 Then
        txtMT100a = XopSwift(I).FieldText
        Call fraMT100_Control(mId$(XopSwift(I).FieldNature, 1, 2))
    End If
Next I
If XopCpt(2).Devise = "001" Then txtfraRib_Init: txtfraRib_Control: fraRib.Visible = False

If XopBicReceiver.BicId = SocBicId Then XopCpt(2).optAvis = "1": XopCpt(2).optAvisLangue = "1"

lstErrClear = False

If lstErr.ListCount = 0 Then
    frmOpTrfD_GlobalControl = True
Else
    frmOpTrfD_GlobalControl = False
'    tag_SetFocus lstErr.Tag
End If
blnGlobalControl = False
End Function


'---------------------------------------------------------
Private Sub frmOpTrfD_Init()
'---------------------------------------------------------

blnfraSolde = False
valRéférence = Trim(txtOpTrf)
frmOpTrfD_Clear
txtOpTrf = Trim(valRéférence)

If Trim(txtOpTrf) = "" Then
    If XopAut.Saisir Then
        recOpCpt_Init arrOpCpt(0)
        arrOpCpt(0).Method = "NUMERO"
        If IsNull(srvOpCpt_Monitor(arrOpCpt(0))) Then
            If arrOpCpt(0).Err = Space$(10) Then
                valRéférence = arrOpCpt(0).Référence
                txtOpTrf = valRéférence
            Else
                Call lstErr_Clear(lstErr, txtOpTrf, "? Numéro automatique" & arrOpCpt(0).Err): Exit Sub
            End If
      Else
            Call lstErr_Clear(lstErr, txtOpTrf, "? Numéro automatique"): Exit Sub
        End If
    Else
        Call lstErr_Clear(lstErr, txtOpTrf, "? dossier"): txtOpTrf.SetFocus: Exit Sub
    End If
End If


srvOpTrf_dbRead valRéférence
XopSwift_Load
XopCpt_Load
currentAction = XopEch_Load
If XopSwiftNb = 0 Then
    If Not XopAut.Saisir Then MsgBox "Vous n'êtes pas autorisé à saisir.", vbCritical, "OpTrf : saisie": Exit Sub
    blnAddNew = True
    cmdSave.Visible = XopAut.Saisir
    cmdOk.Visible = True
    cmdOk.Caption = "à &Valider"
    fraMt100_cmdOk.Visible = XopAut.Saisir
    XopCpt_ActionSaisie valRéférence
    XopEch_ActionSaisie valRéférence
    currentAction = constSaisie
    recBicInit XopBicSender: XopBicReceiver = XopBicSender
    recBicIbanInit XopBicIbanNostro: XopBicIbanReceiver = XopBicIbanSender
    recCompteInit XopCompteSender: XopCompteNostro = XopCompteSender
Else
    XopBic_Load
    blnAddNew = False
    blnLibelléManuel = True
'    txtMT100a = Mid$(XopSwift(1).FieldText, 1, 11)
'    Call Swift00_Control(txtMT100a)
    DevX XopCpt(1).Devise: XopDevise(1) = XDevise
    DevX XopCpt(2).Devise: XopDevise(2) = XDevise
    Swift32A_Init
    txtfraRib_Init
    Call Swift32CompteDb_Control(txtSwift32CompteDb)
End If

Select Case Trim(currentAction)
    Case constSaisie
        cmdSuppress.Caption = "&Supprimer"
        cmdSave.Visible = XopAut.Saisir
        cmdSuppress.Visible = XopAut.Saisir
        cmdOk.Visible = XopAut.Saisir 'True
        cmdOk.Caption = "à &Valider"
        fraMt100_cmdOk.Visible = XopAut.Saisir
        chkLevel = "1"
        chkLibelléManuel.Enabled = True
        blnfraSolde = True
    Case constàValider
        cmdSuppress.Caption = "à &Modifier"
        cmdSuppress.Visible = XopAut.Valider
        cmdOk.Caption = "&Valider"
        cmdOk.Visible = XopAut.Valider
        chkLevel = "2"
        fraMT100.Visible = False
        chkLibelléManuel.Enabled = False
        blnfraSolde = True
End Select


cmdContext.Caption = "&Abandonner"
txtOpTrf.BackColor = txtUsr.BackColor
If XopAut.Saisir And Trim(currentAction) = constSaisie Then blntxtBox_Enabled = True
txtBox_Enabled blntxtBox_Enabled
txtOpTrf.Enabled = False
XopSwift_Init valRéférence
MopSwift(1) = XopSwift(1)
MopSwift(2) = XopSwift(2)
mOpCpt(1) = XopCpt(1)
mOpCpt(2) = XopCpt(2)
XopSwift_Display picMT100: XopTrf_Display picOpTrf
Call XopEch_Display(picHistorique)
SSTab.Tab = 0
SSTab.Enabled = True
If Trim(currentAction) = constSaisie Then
    picMT100Index = 1: picMT100_Line
Else
    fraSolde_Display
End If
End Sub

'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
'frmOpTrfD_Control
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
currentActiveControl_Name = C.Name
End Sub

Public Sub fraSolde_Display()
Dim curX As Currency, curMontant As Currency
fraSolde.Visible = True
fraSolde.Enabled = False
txtSoldeNuméro.Enabled = True
txtSoldeAvant.Enabled = True
txtSoldeMontant.Enabled = True
txtSoldeAprès.Enabled = True
txtSoldeDépassement.Enabled = True
txtSoldeAprès.Visible = blnfraSolde: lblSoldeAprès.Visible = blnfraSolde
txtSoldeDépassement.Enabled = blnfraSolde: lblSoldeDépassement.Enabled = blnfraSolde
txtSoldeAutorisation.Enabled = True
txtDateValeurDébit.Enabled = True
txtDateValeurCrédit.Enabled = True

curMontant = -XopCpt(1).Brut - XopCpt(1).ComMontantDevise
If blnfraSolde Then
    curX = XopCompteSender.SoldeInstantané + curMontant
Else
    curX = 0
End If

txtSoldeNuméro = Format$(XopCompteSender.Devise, "000") & "." & Compte_Imp(XopCompteSender.Numéro)
If XopCompteSender.Situation = " " Then
     lblSoldeIntitulé.ForeColor = libUsr.ForeColor
Else
    lblSoldeIntitulé.ForeColor = errUsr.ForeColor
End If
lblSoldeIntitulé = XopCompteSender.Intitulé

If XopCompteSender.SoldeInstantané < 0 Then
     txtSoldeAvant.ForeColor = errUsr.ForeColor
Else
     txtSoldeAvant.ForeColor = libUsr.ForeColor
End If
txtSoldeAvant = Format$(XopCompteSender.SoldeInstantané, "#### ### ### ### ##0.00")

If curMontant < 0 Then
     txtSoldeMontant.ForeColor = errUsr.ForeColor
Else
     txtSoldeMontant.ForeColor = libUsr.ForeColor
End If
txtSoldeMontant = Format$(curMontant, "#### ### ### ### ##0.00")

lblSoldeAprès = "Solde  " & DevX(XopCompteSender.Devise)
If curX < 0 Then
     txtSoldeAprès.ForeColor = errUsr.ForeColor
Else
     txtSoldeAprès.ForeColor = libUsr.ForeColor
End If
txtSoldeAprès = Format$(curX, "#### ### ### ### ##0.00")

curX = XopCompteSender.SoldeXXX + curMontant
If curX < 0 Then
    txtSoldeDépassement.ForeColor = errUsr.ForeColor
    lblSoldeDépassement.ForeColor = errUsr.ForeColor
    txtSoldeDépassement = Format$(curX, "#### ### ### ### ##0.00")
Else
    txtSoldeDépassement.ForeColor = libUsr.ForeColor
    lblSoldeDépassement.ForeColor = lblUsr.ForeColor
    txtSoldeDépassement = ""
End If

lblSoldeAutorisation.ForeColor = lblUsr.ForeColor
txtSoldeAutorisation = ""
If XopCompteSender.DécouvertMontant > 0 Then
    If Val(XopCompteSender.DécouvertAmj) < DSys Then
        lblSoldeAutorisation.ForeColor = warnUsrColor
    End If
    lblSoldeAutorisation = "Découvert > " & dateImp(XopCompteSender.DécouvertAmj)
    txtSoldeAutorisation = Format$(XopCompteSender.DécouvertMontant, "### ### ### ###")
End If

If XopCpt(1).AmjValeur = DValPrevious Then
    txtDateValeurDébit.ForeColor = libUsr.ForeColor
Else
    txtDateValeurDébit.ForeColor = errUsr.ForeColor
End If
txtDateValeurDébit = dateImp(XopCpt(1).AmjValeur)

If XopCpt(2).AmjValeur = DValNext Then
    txtDateValeurCrédit.ForeColor = libUsr.ForeColor
Else
    txtDateValeurCrédit.ForeColor = errUsr.ForeColor
End If
txtDateValeurCrédit = dateImp(XopCpt(2).AmjValeur)
End Sub

'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------

arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
If blnControl Then
    frmOpTrfD_Control
Else
    blnControl = True
End If
End Sub



Public Sub Msg_Rcv(txtMsg As String)
'---------------------------------------------------------
Select Case UCase$(Trim(mId$(txtMsg, 13, 12)))
    Case Is = "FRMOPTRF":
        Select Case Trim(mId$(txtMsg, 35, 10))
            Case constSaisie: frmOpTrfD.Caption = "OpTrf : Saisie d'une opération"
            Case constValider: frmOpTrfD.Caption = "OpTrf : Validation d'une opération"
            Case Else: frmOpTrfD.Caption = "OpTrf : Détail d'une opération"
        End Select
        frmOpTrfD_Clear
        txtOpTrf = Trim(mId$(txtMsg, 45, 6))
        'action
        If txtOpTrf <> "" Then frmOpTrfD_Init
        
    Case Is = "FRMBDF":
         txtfraRibEtablissement = Trim(mId$(txtMsg, 35, 5))
         txtfraRibGuichet = Trim(mId$(txtMsg, 40, 5))
         txtfraRibCompte = Trim(mId$(txtMsg, 45, 11))
         txtfraRibClé = Trim(mId$(txtMsg, 56, 2))
         txtfraRibGuichet_Control
    Case Is = "FRMCOMPTE":
        Select Case currentActiveControl_Name
            Case "txtfraRibCompte": txtfraRibCompte = Trim(mId$(txtMsg, 38, 11))
            Case "txtSwift32CompteDb":  txtSwift32CompteDb = Trim(mId$(txtMsg, 38, 11))
        End Select
End Select
End Sub
'---------------------------------------------------------
Private Sub cmdContext_Click()
'---------------------------------------------------------

Select Case cmdContext.Caption
    Case Is = "&Rechercher": frmOpTrfD_Init
    Case Is = "&Abandonner"
            If blnMsgBox_Quit Then
                x = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
             Else
                x = vbYes
             End If
             If x = vbYes Then frmOpTrfD_Clear '$jpl: txtOpTrf.SetFocus
End Select

End Sub


'---------------------------------------------------------
Private Sub cmdSave_Click()
'---------------------------------------------------------
Me.Enabled = False: blnOpTrf_Update_on = True
frmOpTrfD_GlobalControl
If blnAddNew Then
    x = srvOpTrf_dbAddNew
Else
    x = srvOpTrf_dbUpdate
End If
frmOpTrfD_Clear
If Trim(x) <> "" Then
    Call lstErr_Clear(lstErr, txtOpTrf, x)
Else
    MsgBox "Erreur pendant la mise à jour AS400", vbCritical, " Opération de transfert"
End If
Me.Enabled = True: blnOpTrf_Update_on = False
AppActivate Me.Caption
End Sub


'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
Dim Msg As String

Msg = ""
prtOpTrfX Msg
End Sub


'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub

'---------------------------------------------------------
Private Sub cmdSuppress_Click()
'---------------------------------------------------------
Me.Enabled = False: blnOpTrf_Update_on = True

If Trim(currentAction) = constàValider Then
    XopEch_ActionSuppress currentAction
    x = srvOpTrf_dbUpdate
Else
    x = MsgBox("Voulez-vous réellement supprimer le dossier?", vbYesNo + vbQuestion + vbDefaultButton2, "Les informations saisies seront perdues!")
    If x = vbYes Then
        x = srvOpTrf_dbDelete
   Else
        GoTo cmdSuppress_Exit
    End If
End If
    frmOpTrfD_Clear
    If Trim(x) <> "" Then
        Call lstErr_Clear(lstErr, txtOpTrf, x)
    Else
        MsgBox "Erreur pendant la mise à jour AS400", vbCritical, " Opération de transfert"
    End If
cmdSuppress_Exit:
Me.Enabled = True: blnOpTrf_Update_on = False
AppActivate Me.Caption
End Sub



'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'    Case Is = 34: cmdPageNext_Click
 '   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub


'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
If XopAut.Saisir Then
    Me.Caption = "OpTrf : Saisie d'une opération"
Else
    Me.Caption = "OpTrf : Détail d'une opération"
End If

Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
ReDim XopCpt(2): XopCptNbMax = 2
ReDim XopEch(10): XopEchNbMax = 10
ReDim XopSwift(20): XopSwiftNbMax = 20

Set fraSwift00.Container = txtMT100a.Container: fraSwift00.Left = 1600: fraSwift00.Top = 1000
Set fraSwift32A.Container = txtMT100a.Container: fraSwift32A.Left = 120: fraSwift32A.Top = 240
Set fraSwift54.Container = txtMT100a.Container: fraSwift54.Left = 1080: fraSwift54.Top = 2160
Set fraSwift71A.Container = txtMT100a.Container: fraSwift71A.Left = 1600: fraSwift71A.Top = 1000
frmOpTrfD_Clear
End Sub



Private Sub SSTab_Click(PreviousTab As Integer)
If SSTab.Tab = 4 Then picCompta.Cls: frmOpTrfD_GlobalControl
Select Case SSTab.Tab
    Case Is = 2: ssTAb_Virement
    Case Is = 4:: DoEvents
                XopCpt_Display picCompta
                txtOpCpt1Libellé = Trim(XopCpt(1).Libellé)
                txtOpCpt2Libellé = Trim(XopCpt(2).Libellé)
                chkLibelléManuel = "0": chkLibelléManuel_Click
    Case Is = 3: picSwift.Cls
        If XopCpt(2).optSwift = "1" Or XopCpt(2).optSwift = "2" Then
            picSwift.CurrentY = 0
            picSwift.ForeColor = warnUsrColor
            picSwift.Print "MT100 à envoyer à " & mId$(XopSwift(2).FieldText, 1, 11);
            Xopswift_MT100R
            Call xOpSwift_arrDisplay(picSwift, 0)
        End If
        If XopCpt(2).optSwift = "4" Then
            picSwift.CurrentY = 0
            picSwift.ForeColor = warnUsrColor
            picSwift.Print "MT100 à envoyer à " & mId$(XopSwift(1).FieldText, 13, 11);
            Xopswift_MT100R
            Call xOpSwift_arrDisplay(picSwift, 0)
        End If
        If XopCpt(2).optSwift = "2" Or XopCpt(2).optSwift = "3" Then
            picSwift.CurrentY = 0: picSwift.CurrentX = 5500
            picSwift.ForeColor = warnUsrColor
            picSwift.Print "MT202 à envoyer à " & mId$(XopSwift(1).FieldText, 13, 11);
            XopSwift_MT202
            Call xOpSwift_arrDisplay(picSwift, 5500)
        End If

End Select
End Sub


Private Sub txtfraRibGuichet_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then Call fraRib_MouseDown(Button, Shift, x, Y)
End Sub

Private Sub txtMT100a_gotfocus()
Call txt_GotFocus(txtMT100a)
End Sub


Private Sub txtMT100a_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtMT100a_LostFocus()
Call txt_LostFocus(txtMT100a)
blnAuto_cmdok_Click = True
End Sub


Private Sub txtOpCpt1Libellé_GotFocus()
Call txt_GotFocus(txtOpCpt1Libellé)
End Sub


Private Sub txtOpCpt1Libellé_LostFocus()
Call txt_LostFocus(txtOpCpt1Libellé)
If XopCpt(1).Libellé <> txtOpCpt1Libellé Then
    XopCpt(1).Libellé = txtOpCpt1Libellé
    blnLibelléManuel = True
    XopCpt_Display picCompta
End If
End Sub


Private Sub txtOpCpt2Libellé_GotFocus()
Call txt_GotFocus(txtOpCpt2Libellé)
End Sub


Private Sub txtOpCpt2Libellé_LostFocus()
Call txt_LostFocus(txtOpCpt2Libellé)
If XopCpt(2).Libellé <> txtOpCpt2Libellé Then
    XopCpt(2).Libellé = txtOpCpt2Libellé
    blnLibelléManuel = True
    XopCpt_Display picCompta
End If
End Sub


Private Sub txtSwift32AmjValeurCr_GotFocus()
Call txt_GotFocus(txtSwift32AmjValeurCr)
End Sub

Private Sub txtSwift32AmjValeurCr_LostFocus()
Call txt_LostFocus(txtSwift32AmjValeurCr)
End Sub


Private Sub txtMT100b_gotfocus()
Call txt_GotFocus(txtMT100b)
End Sub


Private Sub txtMT100b_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtMt100b_LostFocus()
Call txt_LostFocus(txtMT100b)
If blnAuto_cmdok_Click And Trim(txtMT100b) = "" Then fraMT100_cmdOk_Click
End Sub


Private Sub txtMT100c_gotfocus()
Call txt_GotFocus(txtMT100c)
End Sub


Private Sub txtMT100c_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtMT100c_LostFocus()
Call txt_LostFocus(txtMT100c)
If blnAuto_cmdok_Click And Trim(txtMT100c) = "" Then fraMT100_cmdOk_Click
End Sub


Private Sub txtSwift32AmjValeurDb_GotFocus()
Call txt_GotFocus(txtSwift32AmjValeurDb)
End Sub

Private Sub txtSwift32AmjValeurDb_LostFocus()
Call txt_LostFocus(txtSwift32AmjValeurDb)
End Sub


Private Sub txtSwift32CompteDb_GotFocus()
Call txt_GotFocus(txtSwift32CompteDb)
End Sub


Private Sub txtSwift32CompteDb_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtSwift32CompteDb_LostFocus()
Call txt_LostFocus(txtSwift32CompteDb)
End Sub


Private Sub txtMT100d_gotfocus()
Call txt_GotFocus(txtMT100d)
End Sub


Private Sub txtMT100d_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtMT100d_LostFocus()
Call txt_LostFocus(txtMT100d)
If blnAuto_cmdok_Click And Trim(txtMT100d) = "" Then fraMT100_cmdOk_Click
End Sub


Private Sub txtSwift32CoursCr_GotFocus()
Call txt_GotFocus(txtSwift32CoursCr)
End Sub

Private Sub txtSwift32CoursCr_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtSwift32CoursCr)
End Sub


Private Sub txtSwift32CoursCr_LostFocus()
Call txt_LostFocus(txtSwift32CoursCr)
End Sub


Private Sub txtSwift32MontantDb_GotFocus()
Call txt_GotFocus(txtSwift32MontantDb)
End Sub

Private Sub txtswift32montantdb_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtSwift32MontantDb)
End Sub


Private Sub txtswift32montantdb_LostFocus()
Call txt_LostFocus(txtSwift32MontantDb)
End Sub


Private Sub txtSwift32DeviseCr_GotFocus()
Call txt_GotFocus(txtSwift32DeviseCr)
End Sub

Private Sub txtSwift32DeviseCr_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtSwift32DeviseCr_LostFocus()
Call txt_LostFocus(txtSwift32DeviseCr)
End Sub


Private Sub txtMT100e_gotfocus()
Call txt_GotFocus(txtMT100e)
End Sub


Private Sub txtMT100e_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtMT100e_LostFocus()
Call txt_LostFocus(txtMT100e)
If blnAuto_cmdok_Click And Trim(txtMT100e) = "" Then fraMT100_cmdOk_Click
End Sub


Private Sub txtMT100f_gotfocus()
Call txt_GotFocus(txtMT100f)
End Sub


Private Sub txtMT100f_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtMT100f_LostFocus()
Call txt_LostFocus(txtMT100f)
End Sub

Private Sub txtSwift32DeviseDb_GotFocus()
Call txt_GotFocus(txtSwift32DeviseDb)
End Sub


Private Sub txtSwift32DeviseDb_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtSwift32DeviseDb_LostFocus()
Call txt_LostFocus(txtSwift32DeviseDb)
End Sub

Private Sub txtSwift32Montant_GotFocus()
Call txt_GotFocus(txtSwift32Montant)
End Sub

Private Sub txtSwift32Montant_KeyPress(KeyAscii As Integer)
If XopDevise(2).maxD = 0 Then
    Call num_KeyAscii(KeyAscii)
Else
    Call num_KeyAsciiD(KeyAscii, txtSwift32Montant)
End If
End Sub


Private Sub txtSwift32Montant_LostFocus()
Call txt_GotFocus(txtSwift32Montant)
Call txt_LostFocus(txtSwift32Montant)
End Sub



Private Sub txtOpTrf_GotFocus()
'-------------------------------------------------'
Call txt_GotFocus(txtOpTrf)
End Sub

Private Sub txtOpTrf_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtOpTrf_LostFocus()
'-------------------------------------------------'
Call txt_LostFocus(txtOpTrf)
End Sub
Private Sub fraMT100_Clear()
txtMT100a.Visible = True: txtMT100a = "": txtMT100a.BackColor = txtUsr.BackColor
txtMT100b.Visible = True: txtMT100b = "": txtMT100b.BackColor = txtUsr.BackColor
txtMT100c.Visible = True: txtMT100c = "": txtMT100c.BackColor = txtUsr.BackColor
txtMT100d.Visible = True: txtMT100d = "": txtMT100d.BackColor = txtUsr.BackColor
txtMT100e.Visible = True: txtMT100e = "": txtMT100e.BackColor = txtUsr.BackColor
txtMT100f.Visible = True: txtMT100f = "": txtMT100f.BackColor = txtUsr.BackColor
fraSwift32A.Visible = False
fraSwift54.Visible = False
fraSwift71A.Visible = False
fraRib.Visible = False
fraSwift54.Visible = False: chkSwift54 = "0": txtSwift54 = ""
fraSwift00.Visible = False
cmdScan.Visible = False
cmdBicIban.Visible = False

End Sub
Private Sub fraMT100_Init(X2 As String)
Dim K As Integer, Nb As Integer, XMt As String
blnControl = True
blnGlobalControl = False
blnAuto_cmdok_Click = False
MopSwift(1) = XopSwift(1)
MopSwift(2) = XopSwift(2)
mOpCpt(1) = XopCpt(1)
mOpCpt(2) = XopCpt(2)
lstErr.Visible = True
fraMT100.Visible = True: fraSolde.Visible = False
fraMT100_Clear
fraMT100.Caption = lblOpTrf("Swift", X2)
cmdScan.Visible = IIf(X2 = "00" Or X2 = "32" Or X2 = "52" Or X2 = "57", True, False)
cmdBicIban.Visible = IIf(X2 = "57", True, False)

Select Case X2
    Case "00": Nb = 1
    Case "20": Nb = 1
    Case "50": Nb = 4
    Case "32": Nb = 0
    Case "52": Nb = 5
    Case "57": Nb = 5
    Case "59": Nb = 5
    Case "70": Nb = 4
    Case "71": Nb = 0
    Case "72": Nb = 6
End Select
lastActiveControl_Name = "txtMT100f"
If Nb < 6 Then txtMT100f.Visible = False: lastActiveControl_Name = "txtMT100e"
If Nb < 5 Then txtMT100e.Visible = False: lastActiveControl_Name = "txtMT100d"
If Nb < 4 Then txtMT100d.Visible = False: lastActiveControl_Name = "txtMT100c"
If Nb < 3 Then txtMT100c.Visible = False: lastActiveControl_Name = "txtMT100b"
If Nb < 2 Then txtMT100b.Visible = False: lastActiveControl_Name = "txtMT100a"
If Nb < 1 Then txtMT100a.Visible = False


K = 0
For I = 1 To XopSwiftNb
    If mId$(XopSwift(I).FieldNature, 1, 2) > X2 Then Exit For
    If mId$(XopSwift(I).FieldNature, 1, 2) = X2 Then
        Select Case X2
            Case "00"
                txtMT100a = mId$(XopSwift(I).FieldText, 1, 11): If blntxtBox_Enabled And txtMT100a.Enabled Then txtMT100a.SetFocus
                fraSwift00.Visible = True
                fraSwift00.Enabled = blntxtBox_Enabled
                If mId$(XopSwift(1).FieldText, 12, 1) = "*" Then
                    optMT202 = "1"
                Else
                    optMt100 = "1"
                End If
            Case "32"
                lastActiveControl_Name = "txtSwift32CompteDb"
                fraSwift32A.Visible = True
                Swift32A_Init
                If blntxtBox_Enabled Then txtSwift32AmjValeurCr.SetFocus
           Case "71"
                fraSwift71A.Visible = True
                Select Case Trim(XopSwift(I).FieldText)
                    Case "BEN": optImputation1 = True: If blntxtBox_Enabled Then optImputation1.SetFocus
                    Case "SHA": optImputation2 = True: If blntxtBox_Enabled Then optImputation2.SetFocus
                    Case "OUR": optImputation3 = True: If blntxtBox_Enabled Then optImputation3.SetFocus
                    Case Else: optImputation4 = True: If blntxtBox_Enabled Then optImputation4.SetFocus
                End Select
            Case Else
                K = K + 1
                Select Case K
                    Case 1: txtMT100a = Trim(XopSwift(I).FieldText)
                            If blntxtBox_Enabled And txtMT100a.Visible Then txtMT100a.SetFocus
                    Case 2: txtMT100b = Trim(XopSwift(I).FieldText)
                    Case 3: txtMT100c = Trim(XopSwift(I).FieldText)
                    Case 4: txtMT100d = Trim(XopSwift(I).FieldText)
                    Case 5: txtMT100e = Trim(XopSwift(I).FieldText)
                    Case 6: txtMT100f = Trim(XopSwift(I).FieldText)
                End Select
        End Select
    End If
Next I
'If X2 = "52" And Trim(txtMT100a) = "" Then txtMT100a = Mid$(XopSwift(1).FieldText, 1, 11)
If mId$(XopSwift(1).FieldText, 12, 1) <> " " And X2 = "59" And Trim(txtMT100a) = "" Then txtMT100a = mId$(XopSwift(1).FieldText, 1, 11)
If X2 = "57" Then
    fraSwift54.Visible = True
        txtSwift54 = Trim(mId$(XopSwift(2).FieldText, 1, 11))
    If mId$(XopSwift(2).FieldText, 12, 1) = "C" Then
        chkSwift54 = "1"
        txtSwift54.Enabled = True
        txtSwift54 = Trim(mId$(XopSwift(2).FieldText, 1, 11))
    Else
        chkSwift54 = "0"
        txtSwift54.Enabled = False
        txtSwift54 = ""
    End If
End If
If X2 = "59" And XopBicReceiver.BicId = SocBicId Then txtfraRib_Load: If txtfraRibCompte.Visible Then txtfraRibCompte.SetFocus
If XopCpt(2).Devise = "001" And Trim(XopBicReceiver.PaysId) <> "" Then
    If X2 = "57" And XopBicReceiver.PaysId <> "FR" Then txtfraRib_Load
    If X2 = "59" And XopBicReceiver.PaysId = "FR" Then txtfraRib_Load
End If
End Sub
'--------------------------------------------------
Public Sub txtfraRib_Record()
'--------------------------------------------------
Dim x As String, Y As String
Dim strE As String, strG As String, strC As String
strE = Format$(Val(txtfraRibEtablissement), "00000")
'If strE = "00000" Then strE = "     "

strG = Format$(Val(txtfraRibGuichet), "00000")
'If strG = "00000" Then strG = "     "

strC = Format$(RTrim(txtfraRibCompte), "00000000000")
'If strC = "00000000000" Then strC = "           "

x = Format$(Trim(txtfraRibIban), "@@@@") & strE & strG & strC & Format$(txtfraRibClé, "00")

Call Iban_Calc(x, Y)
Mid$(XopSwift(3).FieldText, 1, 27) = Y
If chkRib = 0 Then Mid$(XopSwift(3).FieldText, 3, 2) = "99"
End Sub

'---------------------------------------------------------
Public Sub frmCommission_Show()
'---------------------------------------------------------
If XopCpt(2).Devise = "000" Then Call lstErr_AddItem(lstErr, cmdContext, "? 32:Devise " & x): Exit Sub
If XopCpt(2).Brut = 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? 32:montant " & x): Exit Sub
arrOpCpt(0).Method = currentAction
frmCommission.frmCommission_Init
frmCommission.Show vbModal

End Sub


Private Sub fraMT100_Control(ByVal X2 As String)
Select Case X2
    Case "00": Call Swift00_Control(txtMT100a)
    Case "20": Call Swift20_Control(txtMT100a)
    Case "32": Call Swift32A_Control
    Case "50": Call Swift50_Control(txtMT100a)
    Case "57"
        If fraRib.Visible Then
                txtfraRib_Control
            Else
                Call Swift57_Control(txtMT100a)
            End If
    Case "59"
        If fraRib.Visible Then
                txtfraRib_Control
            Else
                Call Swift59_Control(txtMT100a)
            End If

    Case "70": Call Swift70_Control(txtMT100a)
    Case "71": Call Swift71_Control(txtMT100a)
End Select

If XopCpt(2).Devise <> "001" Then
    XopCpt(2).optVirement = 0: XopCpt(2).optSwift = 2
    If mId$(XopSwift(1).FieldText, 13, 11) = mId$(XopSwift(2).FieldText, 13, 11) Then XopCpt(2).optSwift = 1
Else
    If mId$(XopSwift(3).FieldText, 5, 5) = strSocBdfE Then
        XopCpt(2).optVirement = 4 ': XopCpt(2).optSwift = 0
'    If Mid$(XopSwift(2).FieldText, 13, 11) = SocBicId Then XopCpt(2).optVirement = 4
    Else
        If optVSOT Then XopCpt(2).optVirement = 1: Call fraMT100_ControlFRF(NostroBicVir, "0", "1")
        If optAristote Then XopCpt(2).optVirement = 2: Call fraMT100_ControlFRF(NostroBicVir, "0", "1")
        If optCRI Then XopCpt(2).optVirement = 3: Call fraMT100_ControlFRF(NostroBicCRI, "4", "2")
        If optAutre Then XopCpt(2).optVirement = 5: '''''Call fraMT100_ControlFRF(NostroBicADB, "4", "2")
    End If

End If
If mId$(XopSwift(1).FieldText, 12, 1) = "*" Then XopCpt(2).optSwift = 3
If mId$(XopSwift(2).FieldText, 1, 11) = SocBicId Then XopCpt(2).optSwift = 0


If XopCompteSender.TypeGA <> "A" Then
    XopCpt(1).optAvis = "0"
Else
    If XopCompteSender.BiaTyp = "550" Or XopCompteSender.BiaTyp = "553" Then XopCpt(1).optAvis = "0"
End If
End Sub

Private Sub Swift00_Control(C As Control)
Dim mBicid As String
mBicid = XopBicSender.BicId
XopBicSender.Method = "SeekP0"
XopBicSender.BicId = Trim(C)

If Not IsNull(srvBicFind(XopBicSender)) Then
    Call lstErr_AddItem(lstErr, C, "? correspondant émetteur inconnu")
Else
    If DSys < XopBicSender.AmjMin Then
        Call lstErr_AddItem(lstErr, C, "? émetteur : validité min")
    Else
        If XopBicSender.AmjMax > 0 _
        And DSys > XopBicSender.AmjMax Then
            Call lstErr_AddItem(lstErr, C, "? émetteur : validité max")
         End If
    End If
    If mBicid <> XopBicSender.BicId Then
        XopCpt(1).Compte = Space$(11)
        XopSwift(1).FieldText = Space$(32)
        recCompteInit XopCompteSender
        If XopCpt(1).Devise <> "000" Then Swift32CompteDb_BicIban C
    End If
End If
If Trim(currentAction) = constSaisie Then
    XopCpt(1).optAvis = XopBicSender.optAvis
    
    XopCpt(1).optAvisLangue = XopBicSender.optAvisLangue
    Mid$(XopSwift(1).FieldText, 12, 1) = " "
    If optMT202 Then Mid$(XopSwift(1).FieldText, 12, 1) = "*"
End If
End Sub

Private Sub Swift32DeviseCr_Control(C As Control)
Dim X3 As String * 3
If Not IsNumeric(Trim(C)) Then Call lstErr_AddItem(lstErr, C, "? 32:Devise"): Exit Sub
X3 = Format$(Val(C), "000")
C.Text = X3
If Trim(DevX(X3)) = "" Then Call lstErr_AddItem(lstErr, C, "? 32:Devise inconnue"): txtSwift32DeviseCr.SetFocus: Exit Sub
XopDevise(2) = XDevise
If X3 <> XopCpt(2).Devise Then
    XopCpt(1).CoursDevise = 0: txtSwift32MontantDb = ""
    recOpCpt_CommissionMontantZ XopCpt(1)
'990112     XopCpt(1).Compte = 0: txtSwift32CompteDb = "": txtSwift32DeviseDb = ""
    
    XopCpt(2).CoursDevise = 0: txtSwift32CoursCr = ""
    XopCpt(2).Compte = 0
    recOpCpt_CommissionMontantZ XopCpt(2)
    Mid$(XopSwift(3).FieldText, 1, 27) = Space$(27)
    Mid$(XopSwift(1).FieldText, 13, 11) = Space$(11)
End If
XopCpt(2).Devise = X3
txtSwift32CoursCr.Enabled = IIf(XopCpt(2).Devise = "001", False, blntxtBox_Enabled)
If XopCpt(2).CoursDevise = 0 Then
    XopCpt(2).CoursDevise = XopDevise(2).Cours
    txtSwift32CoursCr = num_Display(XopCpt(2).CoursDevise, 5, 7, Lx, x, " ")
End If
'990112 If Trim(txtSwift32DeviseDb) = "" And X3 <> "000" Then XopCpt(1).Devise = X3: Swift32DeviseDb_Init

'990112 If Val(txtSwift32CompteDb) = 0 Then Swift32CompteDb_BicIban C

Swift32CompteCr_Control

End Sub
Private Sub Swift32DeviseDb_Control(C As Control)
Dim X3 As String * 3

If Trim(txtSwift32DeviseDb) = "" Then txtSwift32DeviseDb = txtSwift32DeviseCr

If Not IsNumeric(Trim(C)) Then Call lstErr_AddItem(lstErr, C, "? 32:Devise Db"): Exit Sub
X3 = Format$(Val(C), "000")
C.Text = X3
If Trim(DevX(X3)) = "" Then Call lstErr_AddItem(lstErr, C, "? 32:Devise Db inconnue"): Exit Sub
XopDevise(1) = XDevise
If X3 <> XopCpt(1).Devise Then
    recOpCpt_CommissionMontantZ XopCpt(1)
    XopCpt(1).CoursDevise = 0
    XopCpt(1).Brut = 0: txtSwift32MontantDb = ""
End If
XopCpt(1).Devise = X3
If XopCpt(1).Devise = XopCpt(2).Devise Then
    XopCpt(1).Brut = XopCpt(2).Brut
    txtSwift32MontantDb.Enabled = False
'    lastActiveControl_Name = "txtSwift32CompteDb"
Else
'    lastActiveControl_Name = "txtSwift32MontantDb"
    txtSwift32MontantDb.Enabled = blntxtBox_Enabled
End If

If XopCpt(1).Brut = 0 Then
    If IsNull(Swift32MontantDb_CV(CV_Conversion)) Then
         XopCpt(1).Brut = CV2.Montant
    End If
End If
txtSwift32MontantDb = num_Display(XopCpt(1).Brut, 13, 2, Lx, x, " ")
If Val(txtSwift32CompteDb) = 0 Then Swift32CompteDb_BicIban C
End Sub

Private Sub Swift32AmjValeurCr_Control(C As Control)
x = dateCtlDsys(txtSwift32AmjValeurCr.Text)
If Not IsNumeric(x) Or mId$(x, 1, 8) = "00000000" Then
    Call lstErr_AddItem(lstErr, C, "? date : " & x)
Else
    XopCpt(2).AmjValeur = mId$(x, 1, 8)
    Call dateDsp(XopCpt(2).AmjValeur, txtSwift32AmjValeurCr)
'    If Val(XopCpt(1).AmjValeur) = 0 Then
'        XopCpt(1).AmjValeur = Mid$(x, 1, 8)
'        Call dateDsp(XopCpt(1).AmjValeur, txtSwift32AmjValeurDb)
'    End If
End If
End Sub

Private Sub Swift32AmjValeurdb_Control(C As Control)
    x = dateCtlDsys(txtSwift32AmjValeurDb.Text)
If Not IsNumeric(x) Or mId$(x, 1, 8) = "00000000" Then
    Call lstErr_AddItem(lstErr, C, "? date : " & x)
Else
    XopCpt(1).AmjValeur = mId$(x, 1, 8)
    Call dateDsp(XopCpt(1).AmjValeur, txtSwift32AmjValeurDb)
End If
End Sub

'--------------------------------------------------
Private Sub cmdScan_Click()
'--------------------------------------------------
arrTag_Set False
If fraSwift32A.Visible Then
    cmdScan_BicIbanNostro
Else
    If wNature = "57" Then
        cmdScan_BicReceiver
    Else
        cmdScan_Bic
    End If
End If
'Select Case currentActiveControl_Name
'    Case "txtMT100a": cmdScan_Bic
'    Case Else: cmdScan_BicIbanNostro
'End Select
End Sub

Private Sub Swift32Montant_Control(C As Control)
Dim valX As String, mBrut As Currency

mBrut = XopCpt(2).Brut: XopCpt(2).Brut = 0
If Trim(C.Text) = "" Then Call lstErr_AddItem(lstErr, C, "? 32:montant " & x): Exit Sub
x = num_Control(C, valX, 13, XopDevise(2).maxD)
If x <> "" Then Call lstErr_AddItem(lstErr, C, "? 32:montant " & x): Exit Sub
If valX = 0 Then Call lstErr_AddItem(lstErr, C, "? 32:préciser montant"): Exit Sub
If valX <> mBrut Then
    recOpCpt_CommissionMontantZ XopCpt(1)
    recOpCpt_CommissionMontantZ XopCpt(2)
End If
If XopCpt(1).Brut = 0 _
And XopCpt(1).Devise = XopCpt(2).Devise Then
    XopCpt(1).Brut = valX: txtSwift32MontantDb = num_Display(XopCpt(1).Brut, 13, 2, Lx, x, " ")
End If

XopCpt(2).Brut = valX
End Sub

Private Sub Swift32CompteDb_Control(C As Control)
x = num_Control(C, XopCpt(1).Compte, 11, 0)
If x <> "" Then Call lstErr_AddItem(lstErr, C, "? 32:compte"): Exit Sub
If XopCpt(1).Compte = 0 Then Call lstErr_AddItem(lstErr, C, "? 32:compte"): Exit Sub
C.Text = Compte_Display(XopCpt(1).Compte)
recCompteInit XopCompteSender
XopCompteSender.Method = "SeekL1"
XopCompteSender.Société = SocId$
XopCompteSender.Agence = SocAgence$
XopCompteSender.Devise = XopCpt(1).Devise
XopCompteSender.Numéro = XopCpt(1).Compte
If IsNull(srvCompteMon(XopCompteSender)) Then
    If XopCompteSender.Situation = "A" Or XopCompteSender.Situation = "E" Then
        Call lstErr_AddItem(lstErr, C, "? compte DB annulé")
    End If
Else
    Call lstErr_AddItem(lstErr, C, "? compte DB inconnu")
End If
End Sub

Private Sub Swift32CoursCr_Control(C As Control)
x = num_Control(C, XopCpt(2).CoursDevise, 5, 7)
If x <> "" Then Call lstErr_AddItem(lstErr, C, "? 32:cours CR"): Exit Sub
If XopCpt(2).CoursDevise = 0 Then Call lstErr_AddItem(lstErr, C, "? 32:cours CR"): Exit Sub
C.Text = num_Display(XopCpt(2).CoursDevise, 5, 7, Lx, x, " ")
End Sub

Private Sub Swift32MontantDb_Control(C As Control)
Dim valX As String
XopCpt(1).Brut = 0
If Trim(C.Text) = "" Then Call lstErr_AddItem(lstErr, C, "? 32:montant Débit " & x): Exit Sub
x = num_Control(C, valX, 13, XopDevise(1).maxD)
If x <> "" Then Call lstErr_AddItem(lstErr, C, "? 32:montant Débit " & x): Exit Sub
If valX = 0 Then Call lstErr_AddItem(lstErr, C, "? 32:préciser montant Débit"): Exit Sub
XopCpt(1).Brut = valX
If XopCpt(2).Brut = 0 Then
    If IsNull(Swift32MontantCr_CV(CV_Conversion)) Then
         XopCpt(2).Brut = CV2.Montant
        txtSwift32Montant = num_Display(XopCpt(2).Brut, 13, 2, Lx, x, " ")
    End If
End If

End Sub


Private Sub Swift50_Control(C As Control)
If mId$(XopSwift(1).FieldText, 12, 1) = " " And Trim(C) = "" Then
    Call lstErr_AddItem(lstErr, C, "? 50:client donneur d'ordre")
End If

End Sub
Private Sub Swift57_ControlBicReceiver(C As Control)
Dim V As Variant
' If Trim(Mid$(MopSwift(2).FieldText, 1, 11)) <> Trim(C) Then

XopBicReceiver.BicId = Trim(C) ': Mid$(xopswift(3).fieldtext, 1, 27) = Space(27)
V = srvBic_Control(XopBicReceiver, XopCpt(2).AmjOpération)
If Not IsNull(V) Then
    Call lstErr_AddItem(lstErr, C, "57" & V)
Else
    XopCpt(2).optAvisLangue = XopBicReceiver.optAvisLangue
    If Trim(mId$(XopSwift(2).FieldText, 1, 11)) <> Trim(XopBicReceiver.BicId) Then XopSwift(3).FieldText = ""
    Mid$(XopSwift(2).FieldText, 1, 11) = XopBicReceiver.BicId
    If XopBicReceiver.BicId = SocBicId Then
        Mid$(XopSwift(3).FieldText, 5, 5) = strSocBdfE
        Mid$(XopSwift(3).FieldText, 10, 5) = strSocBdfG
        Mid$(XopSwift(1).FieldText, 13, 11) = Space$(11)
        Mid$(XopSwift(2).FieldText, 13, 11) = Space$(11)
    Else
        'If Mid$(XopSwift(1).FieldText, 12, 1) = " " Then
            recBicIbanInit XopBicIbanReceiver
            XopBicIbanReceiver.IbanBicId = mId$(XopSwift(2).FieldText, 13, 11)
            Swift57_ControlIbanReceiver
        'Else
         '   Mid$(XopSwift(2).FieldText, 13, 11) = Space$(11)
        'End If
    End If
End If

End Sub

Private Sub Swift59_Control(C As Control)
If Trim(C) = "" Then
    Call lstErr_AddItem(lstErr, C, "? 59:client bénéficiaire")
End If
valBen = Trim(C)
If Not blnLibelléManuel Then XopCpt(1).Libellé = valLibellé & valBen
If XopBicReceiver.BicId = SocBicId Then
    txtfraRib_Load
Else
    If XopCpt(2).Devise = "001" And XopBicReceiver.PaysId = "FR" Then
        txtfraRib_Load
    Else
        fraRib.Visible = False
    End If
End If
End Sub
Private Sub Swift71_Control(C As Control)
Dim K1 As Integer, K2 As Integer
If blnGlobalControl Then
    If XopCpt(1).optComImputation = 4 And XopCpt(2).optComImputation = 4 Then Exit Sub
    If XopCpt(1).ComMontantFRF + XopCpt(2).ComMontantFRF = 0 Then
        Call lstErr_AddItem(lstErr, txtMT100a, "? préciser commissions")
    End If
Else
    If optImputation1 = "1" Then K1 = 4: K2 = 1: txtMT100a = "BEN"
    If optImputation2 = "1" Then K1 = 2: K2 = 2: txtMT100a = "SHA"
    If optImputation3 = "1" Then K1 = 3: K2 = 4: txtMT100a = "OUR"
    If optImputation4 = "1" Then
        txtMT100a = "FREE"
        recOpCpt_CommissionReset XopCpt(1)
        XopCpt(1).optComImputation = 4
        recOpCpt_CommissionReset XopCpt(2)
        XopCpt(2).optComImputation = 4
        Exit Sub
    End If
    
    XopCpt(1).optComImputation = K1
    If XopCpt(1).optComBarème <> 4 Then
        recOpCpt_CommissionMontantZ XopCpt(1)
        If XopCpt(1).optComBarème = "0" Then XopCpt(1).optComBarème = 2
        If XopCpt(1).optComShare = "0" Then XopCpt(1).optComShare = IIf(K1 = 2, 2, 1)
    End If
    If K1 = 3 Or K1 = 2 Then
        arrOpCpt(0) = XopCpt(1)
        frmCommission_Show
        XopCpt(1) = arrOpCpt(0)
    End If
    
    XopCpt(2).optComImputation = K2
    If XopCpt(2).optComBarème <> 4 Then
        recOpCpt_CommissionMontantZ XopCpt(2)
        If XopCpt(2).optComBarème = 0 Then XopCpt(2).optComBarème = 2
        If XopCpt(2).optComShare = "0" Then XopCpt(2).optComShare = IIf(K2 = 2, 2, 1)
    End If
    If K2 = 2 Or K2 = 1 Then
        arrOpCpt(0) = XopCpt(2)
        frmCommission_Show
        XopCpt(2) = arrOpCpt(0)
    End If
End If
End Sub


Public Sub Swift20_Control(C As Control)
If Trim(C) = "" Then
    Call lstErr_AddItem(lstErr, C, "? 20:Numéro de transaction")
End If
valLibellé = Trim(C) & " (OpTrf : " & Trim(valRéférence) & ")"
If Not blnLibelléManuel Then
    XopCpt(1).Libellé = valLibellé & " " & valBen
    If XopCpt(2).Compte > "00099999999" _
    And mId$(XopCpt(2).Compte, 6, 3) = "550" Or mId$(XopCpt(2).Compte, 6, 3) = "553" Then
        XopCpt(2).Libellé = "OpTrf : " & Trim(valRéférence) & " - " & Trim(C)     'XopCpt(1).Libellé
    Else
        XopCpt(2).Libellé = XopCpt(1).Libellé
    End If
End If
'Mid$(XopCpt(2).Libellé, 1, 16) = "OpTrf : " & Trim(valRéférence)
End Sub

Public Sub txtfraRib_Init()

libfraRibDom = "": libfraRibCompte = ""
txtfraRibIban.Enabled = False

txtfraRibEtablissement.ForeColor = txtUsr.ForeColor
txtfraRibEtablissement.BackColor = txtUsr.BackColor
txtfraRibGuichet.ForeColor = txtUsr.ForeColor
txtfraRibGuichet.BackColor = txtUsr.BackColor
txtfraRibCompte.ForeColor = txtUsr.ForeColor
txtfraRibCompte.BackColor = txtUsr.BackColor
txtfraRibClé.ForeColor = txtUsr.ForeColor
txtfraRibClé.BackColor = txtUsr.BackColor
txtfraRibIban = mId$(XopSwift(3).FieldText, 1, 4)
If Trim(txtfraRibIban) = "" Then txtfraRibIban = "FR00"
txtfraRibEtablissement = Trim(mId$(XopSwift(3).FieldText, 5, 5)): If txtfraRibEtablissement = "00000" Then txtfraRibEtablissement = ""

optAristote.Enabled = True
optVSOT.Enabled = True
optCRI.Enabled = True
optAutre.Enabled = True 'False
fraVirement.Enabled = XopAut.Saisir
txtfraRibGuichet = Trim(mId$(XopSwift(3).FieldText, 10, 5)): If txtfraRibGuichet = "00000" Then txtfraRibGuichet = ""
txtfraRibCompte = Trim(mId$(XopSwift(3).FieldText, 15, 11)): If txtfraRibCompte = "00000000000" Then txtfraRibCompte = ""
txtfraRibClé = Trim(mId$(XopSwift(3).FieldText, 26, 2))
chkRib = IIf(txtfraRibClé = "00", 0, 1): chkRib_Click
If mId$(XopSwift(3).FieldText, 5, 5) = strSocBdfE Then
    fraVirement.Visible = False
Else
    fraVirement.Visible = True
    If XopCpt(2).Brut >= constMontantCRI Then
        fraVirement.Enabled = False: optCRI = "1"
    Else
        fraVirement.Enabled = True
        Select Case XopCpt(2).optVirement
            Case "1": optVSOT = "1"
            Case "2": optAristote = "1"
            Case "3": optCRI = "1"
            Case "5": optAutre = "1"
        End Select
    End If
End If
End Sub

Public Sub Swift32A_Init()
Call dateDsp(XopCpt(2).AmjValeur, txtSwift32AmjValeurCr)
'txtSwift32DeviseCr = XopCpt(2).Devise
If XopCpt(2).Devise = "000" Then
    txtSwift32DeviseCr = ""
Else
    txtSwift32DeviseCr = XopCpt(2).Devise
End If
txtSwift32Montant = num_Display(XopCpt(2).Brut, 13, 2, Lx, x, " ")
txtSwift32CoursCr = num_Display(XopCpt(2).CoursDevise, 5, 7, Lx, x, " ")

'If XopCpt(1).AmjValeur = "00000000" Then XopCpt(1).AmjValeur = DsysMinus2
Call dateDsp(XopCpt(1).AmjValeur, txtSwift32AmjValeurDb)
If XopCpt(1).Devise = "000" Then
    txtSwift32DeviseDb = ""
Else
    txtSwift32DeviseDb = XopCpt(1).Devise
End If
txtSwift32CompteDb = Compte_Display(XopCpt(1).Compte)
txtSwift32MontantDb = num_Display(XopCpt(1).Brut, 13, 2, Lx, x, " ")
libSwift32CompteC = Compte_Display(XopCpt(2).Compte)
libSwift32IbanBicId = mId$(XopSwift(1).FieldText, 13, 11)
End Sub
Public Sub Swift32A_Control()
Dim strX As String, curMt As Currency, K2 As Integer
Dim V

Call Swift32AmjValeurCr_Control(txtSwift32AmjValeurCr)
Call Swift32DeviseCr_Control(txtSwift32DeviseCr)
Call Swift32CoursCr_Control(txtSwift32CoursCr)
Call Swift32Montant_Control(txtSwift32Montant)
Call Swift32AmjValeurdb_Control(txtSwift32AmjValeurDb)
Call Swift32CompteDb_Control(txtSwift32CompteDb)
Call Swift32DeviseDb_Control(txtSwift32DeviseDb)
Call Swift32MontantDb_Control(txtSwift32MontantDb)
If chkLevel <> "0" Then
    If XopCpt(1).Brut = 0 Then
        XopCpt(1).CoursDevise = 1
    Else
        XopCpt(1).CoursDevise = XopCpt(2).Brut * XopCpt(2).CoursDevise / XopCpt(1).Brut
        If XopCpt(1).Devise <> XopCpt(2).Devise Then
            If Abs((XopCpt(1).CoursDevise - XopDevise(1).Cours) / XopCpt(1).CoursDevise) > 0.1 Then
                x = MsgBox("Cours saisi :" & num_Display(XopCpt(1).CoursDevise, 10, 5, Lx, x, " "), vbInformation, "Cours Comptable " & XopDevise(1).DevX & " / FRF :" & XopDevise(1).Cours)
            End If
        End If
    End If
End If

If XopCpt(2).AmjValeur < XopCpt(1).AmjValeur Then
    If XopCpt(1).chkAmjValeur < chkLevel Then
        blnMsgBox = True
        x = MsgBox("La date de valeur au crédit est antérieure à la date de valeur au débit.", vbCritical, Trim(XopCompteNostro.Intitulé) & " : " & dateImp(XopCpt(2).AmjValeur) & " < " & dateImp(XopCpt(1).AmjValeur))
        If x = vbYes Then
            XopCpt(1).chkAmjValeur = chkLevel
        Else
            Call lstErr_AddItem(lstErr, txtSwift32AmjValeurCr, "? date valeur débit > crédit")
        End If
    End If
End If

If XopCpt(1).chkAmjValeur <> mOpCpt(1).chkAmjValeur Then XopCpt(1).chkAmjValeur = "0"
If XopCpt(2).chkAmjValeur <> mOpCpt(2).chkAmjValeur Then XopCpt(2).chkAmjValeur = "0"


If XopCpt(1).AmjValeur < DsysValueMin Or XopCpt(1).AmjValeur > DsysValueMax Then
    If XopCpt(1).chkAmjValeur < chkLevel Then
        blnMsgBox = True
        x = MsgBox("La date de valeur au débit hors [ -7 , +7], confirmez-vous cette date ?", vbYesNo + vbQuestion + vbDefaultButton2, Trim(XopCompteSender.Intitulé) & " : " & dateImp(XopCpt(1).AmjValeur))
        If x = vbYes Then
            XopCpt(1).chkAmjValeur = chkLevel
        Else
            Call lstErr_AddItem(lstErr, txtSwift32AmjValeurDb, "? date valeur au débit")
        End If
    End If
End If

If XopCpt(2).AmjValeur < DsysValueMin Or XopCpt(2).AmjValeur > DsysValueMax Then
    If XopCpt(2).chkAmjValeur < chkLevel Then
         blnMsgBox = True
       x = MsgBox("La date de valeur au crédit hors [ -7 , +7], confirmez-vous cette date ?", vbYesNo + vbQuestion + vbDefaultButton2, Trim(XopCompteNostro.Intitulé) & " : " & dateImp(XopCpt(2).AmjValeur))
        If x = vbYes Then
            XopCpt(2).chkAmjValeur = chkLevel
        Else
            Call lstErr_AddItem(lstErr, txtSwift32AmjValeurCr, "? date valeur au crédit")
        End If
    End If
End If

V = dateImp(XopCpt(2).AmjValeur)
K2 = Weekday(V)
If K2 = 1 Or K2 = 7 Then
    blnMsgBox = True
    x = MsgBox("Le jour de valeur au crédit est férié, confirmez-vous cette date ?", vbYesNo + vbQuestion + vbDefaultButton2, Trim(XopCompteNostro.Intitulé) & " : " & dateImp(XopCpt(2).AmjValeur) & "  Week-end")
    If x <> vbYes Then
        Call lstErr_AddItem(lstErr, txtSwift32AmjValeurCr, "? date valeur au crédit week-end")
    End If
End If


V = dateImp(XopCpt(1).AmjValeur)
K2 = Weekday(V)
If K2 = 1 Or K2 = 7 Then
    blnMsgBox = True
    x = MsgBox("Le jour de valeur au débit est férié, confirmez-vous cette date ?", vbYesNo + vbQuestion + vbDefaultButton2, Trim(XopCompteSender.Intitulé) & " : " & dateImp(XopCpt(1).AmjValeur) & "  Week-end")
    If x <> vbYes Then
        Call lstErr_AddItem(lstErr, txtSwift32AmjValeurDb, "? date valeur au débit week-end")
    End If
End If
If XopCpt(1).Devise <> mOpCpt(1).Devise _
Or XopCpt(1).Compte <> mOpCpt(1).Compte _
Or XopCpt(1).Brut <> mOpCpt(1).Brut Then
    XopCpt(1).chkCompte = "0"
    XopCpt(1).chkSolde = "0"
End If
If XopCompteSender.Situation = "B" Then
    If XopCpt(1).chkCompte < chkLevel Then
        blnMsgBox = True
        x = MsgBox("Le compte à débiter est bloqué, confirmez-vous cette opération ?", vbYesNo + vbQuestion + vbDefaultButton2, Trim(XopCompteSender.Intitulé) & " : " & currentAction)
        If x = vbYes Then
            XopCpt(1).chkCompte = chkLevel
        Else
            Call lstErr_AddItem(lstErr, txtSwift32CompteDb, "Le compte à débiter est bloqué")
        End If
    End If
End If
If XopCompteSender.BiaTyp <> "550" And XopCompteSender.BiaTyp <> "553" Then
    If XopCompteSender.SoldeXXX - XopCpt(1).Brut < 0 Then
        If XopCpt(1).chkSolde < chkLevel Then
            strX = "Le compte n'a pas la provision suffisante :" & num_Display(XopCompteSender.SoldeXXX - XopCpt(1).Brut, 15, 0, Lx, x, " ") & " Débit" & Chr$(13) & "(solde provisoire incluant cette opération). Confirmez-vous ?"
            blnMsgBox = True
            x = MsgBox(strX, vbYesNo + vbQuestion + vbDefaultButton2, Trim(XopCompteSender.Intitulé))
            If x = vbYes Then
                XopCpt(1).chkSolde = chkLevel
            Else
                Call lstErr_AddItem(lstErr, txtSwift32CompteDb, "provision insuffisante")
            End If
        End If
    End If
End If


End Sub
Public Sub Swift32CompteCr_Control()
recBicIbanInit XopBicIbanNostro
XopBicIbanNostro.BicId = SocBicIdNostro
XopBicIbanNostro.IbanDevise = XopCpt(2).Devise
XopBicIbanNostro.IbanBicId = mId$(XopSwift(1).FieldText, 13, 11)

'REcherche du correspondant principal
If Trim(mId$(XopSwift(1).FieldText, 13, 11)) = "" Then
    arrBicIbanNb = 0
    Mid$(XopBicIbanNostro.BicId, 9, 3) = "COR"
    XopBicIbanNostro.Method = "SnapP0"
    arrBicIban(0) = XopBicIbanNostro
    arrBicIban(0).IbanBicId = "99999999999"
    srvBicIbanMon XopBicIbanNostro
    If arrBicIbanNb > 1 Then XopBicIbanNostro = arrBicIban(1)
End If

XopBicIbanNostro.BicId = SocBicIdNostro

'If Trim(XopBicIbanNostro.IbanBicId) = "" Then
'If XopCpt(2).Compte = 0 Then
'    Swift32CompteCr_ControlIban
'Else

    XopBicIbanNostro.Method = "SeekP0"
    If Not IsNull(srvBicIbanFind(XopBicIbanNostro)) Then
        Call lstErr_AddItem(lstErr, txtSwift32DeviseCr, "? Dev " & SocBicIdNostro)
    Else
        Swift32CompteCr_ControlOk
    End If
    
'End If
End Sub

Public Sub txtfraRib_Control()
txtfraRibGuichet_Control
txtfraRibCompte_Control
If chkRib <> 0 Then txtfraRibClé_Control
txtfraRib_Record

End Sub

Public Sub txtfraRib_Load()
txtfraRib_Init
fraRib.Visible = True
lastActiveControl_Name = "txtfraRibClé"
txtfraRibGuichet_Control
txtfraRibCompte_Control
'If blntxtBox_Enabled Then txtfraRibCompte.SetFocus
If XopBicReceiver.BicId = SocBicId Then
'    txtfraRibEtablissement = strSocBdfE
'    txtfraRibGuichet = strSocBdfG
    txtfraRibEtablissement.Enabled = False
    txtfraRibGuichet.Enabled = False
'    txtfraRibCompte.SetFocus
Else
    txtfraRibEtablissement.Enabled = blntxtBox_Enabled
    txtfraRibGuichet.Enabled = blntxtBox_Enabled
 '   txtfraRibEtablissement.SetFocus
End If

End Sub


Public Sub Swift32DeviseDb_Init()
DevX XopCpt(1).Devise
XopDevise(1) = XDevise
XopCpt(1).CoursDevise = XopDevise(1).Cours
txtSwift32DeviseDb = XopCpt(1).Devise
txtSwift32MontantDb = num_Display(XopCpt(1).Brut, 13, 2, Lx, x, " ")
txtSwift32MontantDb.Enabled = IIf(XopCpt(1).Devise = "001", False, blntxtBox_Enabled)
End Sub

Public Sub ssTAb_Virement()
picVirement.Cls
picVirement.CurrentY = 2000: picVirement.CurrentX = 3000: picVirement.FontSize = 14
Select Case XopCpt(2).optVirement
    Case Is = "0": picVirement.Print "NEANT " & Now
    Case Is = "1": picVirement.Print " VSOT via ARISTOTE "
        ' Set XPrt = picVirement: prtVirx "picVirement"
    Case Is = "2": picVirement.Print " Virement via ARISTOTE "
    Case "3", "1": picVirement.Print "virement via C R I "
    Case Is = "4": picVirement.Print "Virement interne "
    Case Is = "5": picVirement.Print "Virement autre "
End Select

End Sub

Public Sub cmdContext_Return()
If blnMsgBox Then blnMsgBox = False: Exit Sub
blnMsgBox = False
If SSTab.Enabled = False Then
    cmdContext_Click
Else
    If fraMT100.Visible Then
        If ActiveControl.Name = lastActiveControl_Name Then
           fraMT100_cmdOk_Click
        Else
            If fraSwift71A.Visible Then
                fraMT100_cmdOk_Click
            Else
                SendKeys "{TAB}"
            End If
        End If
    End If
End If

End Sub

Public Sub cmdContext_Quit()
blnControl = False
If lstDevise.Visible Then
    lstDevise.Visible = False
Else
    If fraMT100.Visible Then
        fraMT100_cmdQuit_Click
    Else
        cmdQuit_Click
    End If
End If

End Sub

Public Sub cmdScan_Bic()
arrBic(0).Method = "Bic"
arrBic(0).BicId = Trim(txtMT100a)
frmBicList.Show vbModal
If Trim(arrBic(0).Method) = "Ok" Then txtMT100a = arrBic(0).BicId
End Sub

Public Sub cmdScan_BicIbanNostro()
arrBic(0).Method = "IbanDevise"
arrBicIban(0).BicId = SocBicIdNostro
arrBicIban(0).IbanDevise = XopCpt(2).Devise
frmBicList.Show vbModal
If Trim(arrBicIban(0).Method) = "Ok" Then
    XopBicIbanNostro = arrBicIban(0)
    Swift32CompteCr_ControlOk
End If
End Sub

Public Sub Swift32CompteCr_ControlOk()

Mid$(XopSwift(1).FieldText, 13, 11) = XopBicIbanNostro.IbanBicId
XopCpt(2).Compte = mId$(XopBicIbanNostro.IbanCompte, 11, 11)
libSwift32CompteC = Compte_Display(XopCpt(2).Compte)
libSwift32IbanBicId = XopBicIbanNostro.IbanBicId
recCompteInit XopCompteNostro
XopCompteNostro.Method = "SeekL1"
XopCompteNostro.Société = SocId$
XopCompteNostro.Agence = SocAgence$
XopCompteNostro.Devise = XopCpt(2).Devise
XopCompteNostro.Numéro = XopCpt(2).Compte
 If IsNull(srvCompteMon(XopCompteNostro)) Then
'$JPL 99-09-06$    If XopCompteNostro.Situation = "A" Then
'$JPL 99-09-06$        Call lstErr_AddItem(lstErr, txtSwift32DeviseCr, "? compte à créditer annulé")
'$JPL 99-09-06$    End If
Else
    Call lstErr_AddItem(lstErr, txtSwift32DeviseCr, "? compte à créditer inconnu")
End If
If Trim(mId$(XopSwift(1).FieldText, 13, 11)) <> NostroBicVir _
And Trim(mId$(XopSwift(1).FieldText, 13, 11)) <> NostroBicCRI Then
    optAutre = "1"
End If
End Sub

Public Sub Swift32CompteDb_BicIban(C As Control)
recBicIbanInit XopBicIbanSender
XopBicIbanSender.Method = "SeekP0"
XopBicIbanSender.BicId = XopBicSender.BicId
XopBicIbanSender.IbanDevise = XopCpt(1).Devise
XopBicIbanSender.IbanBicId = SocBicId
recCompteInit XopCompteSender

If Not IsNull(srvBicIbanFind(XopBicIbanSender)) Then
    XopBicIbanSender.IbanDevise = "001"
    If Not IsNull(srvBicIbanFind(XopBicIbanSender)) Then
        Call lstErr_AddItem(lstErr, C, "? correspondant-devise-compte")
        Exit Sub
    End If
End If

XopCpt(1).Compte = mId$(XopBicIbanSender.IbanCompte, 11, 11)
txtSwift32CompteDb = Compte_Display(XopCpt(1).Compte)
End Sub


Public Sub Swift70_Control(C As Control)
If mId$(XopSwift(1).FieldText, 12, 1) = " " And Trim(C) = "" Then
    Call lstErr_AddItem(lstErr, C, "? 70: informations au bénéficiaire")
Else
    If Not blnLibelléManuel Then XopCpt(2).Libellé = Trim(C) & " (OpTrf : " & Trim(valRéférence) & ") "
End If

End Sub

Public Sub Swift57_ControlIbanReceiver()
XopBicIbanReceiver.Method = "SeekP0"
XopBicIbanReceiver.BicId = XopBicReceiver.BicId
XopBicIbanReceiver.IbanDevise = XopCpt(2).Devise
If IsNull(srvBicIbanFind(XopBicIbanReceiver)) Then
    Mid$(XopSwift(2).FieldText, 13, 11) = XopBicIbanReceiver.IbanBicId
    If Trim(mId$(XopSwift(3).FieldText, 5, 23)) = "" Then
        Mid$(XopSwift(3).FieldText, 1, 4) = XopBicIbanReceiver.IbanPays & XopBicIbanReceiver.IbanClé
        Mid$(XopSwift(3).FieldText, 5, 23) = XopBicIbanReceiver.IbanCompte
    End If
Else
    Mid$(XopSwift(2).FieldText, 13, 11) = ""
End If
If XopCpt(2).Devise <> "001" Then
    fraRib.Visible = False
'$JPL!!!!!!!!!!!!!!    If Trim(Mid$(XopSwift(2).FieldText, 13, 11)) = "" Or Trim(XopBicReceiver.SwiftId) = "" Then Call lstErr_AddItem(lstErr, txtMT100a, "? 57 non Swift")
                                'chkSwift54 = "1" '''': chkSwift54_Click
Else
    If XopBicReceiver.PaysId = "FR" Then
        If Trim(mId$(XopSwift(3).FieldText, 5, 5)) = "" Then
            Mid$(XopSwift(3).FieldText, 5, 5) = XopBicReceiver.BdfEtablissement
            Mid$(XopSwift(3).FieldText, 10, 5) = XopBicReceiver.BdfGuichet
        End If
    Else
        txtfraRib_Load
    End If
End If

End Sub

Public Sub cmdScan_BicReceiver()
XopBicIbanReceiver.BicId = XopBicReceiver.BicId
XopBicIbanReceiver.IbanDevise = XopCpt(2).Devise

cmdScan_Bic

If Trim(arrBicIban(0).Method) = "Ok" Then
    XopBicIbanReceiver = arrBicIban(0)
    Mid$(XopSwift(3).FieldText, 5, 23) = ""
    Swift57_ControlIbanReceiver
End If

End Sub

Private Sub txtSwift54_GotFocus()
Call txt_GotFocus(txtSwift54)
End Sub


Private Sub txtSwift54_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSwift54_LostFocus()
Call txt_LostFocus(txtSwift54)
End Sub



Public Sub Swift57_Control(C As Control)
If Trim(C) = "" Then Call lstErr_AddItem(lstErr, C, "? 57:compte auprès de"): Exit Sub

If Len(Trim(C.Text)) > 11 Then
    chkSwift54 = "1": chkSwift54_Control
    lastActiveControl_Name = "txtSwift54"
'    Call Swift54_Control(txtSwift54)
Else
    chkSwift54 = "0": chkSwift54_Control
    lastActiveControl_Name = "txtMT100e"
    Swift57_ControlBicReceiver C
End If
End Sub

Public Sub Swift54_Control(C As Control)
If Trim(C) = "" Then Call lstErr_AddItem(lstErr, C, "? 57:compte auprès de"): Exit Sub

Swift57_ControlBicReceiver C

End Sub

Public Sub fraMT100_ControlFRF(nostroBic As String, X1 As String, X2 As String)
If XopCpt(2).Devise = "978" Or XopCpt(2).Devise = "001" Then
    If mId$(XopSwift(1).FieldText, 13, 11) <> nostroBic Then
        Mid$(XopSwift(1).FieldText, 13, 11) = nostroBic
        Swift32CompteCr_Control
    End If
    XopCpt(2).optSwift = IIf(XopBicReceiver.PaysId = "FR", X1, X2)
End If
End Sub

Public Sub chkSwift54_Control()
Select Case chkSwift54
    Case Is = "1": txtSwift54.Enabled = True: Mid$(XopSwift(2).FieldText, 12, 1) = "C"
    Case Else: txtSwift54.Enabled = False: txtSwift54 = "": Mid$(XopSwift(2).FieldText, 12, 1) = " "
End Select

End Sub

Public Sub Swift32CompteCr_ControlIban()
arrBicIbanNb = 0
XopBicIbanNostro.Method = "SnapP0"
arrBicIban(0) = XopBicIbanNostro
arrBicIban(0).IbanBicId = "99999999999"
If Not IsNull(srvBicIbanMon(XopBicIbanNostro)) Then
    Call lstErr_AddItem(lstErr, txtSwift32DeviseCr, "? srvBicIban " & SocBicIdNostro)
Else
    If arrBicIbanNb = 1 Then
        XopBicIbanNostro = arrBicIban(1)
        Swift32CompteCr_ControlOk
    Else
        cmdScan_BicIbanNostro
    End If
End If

End Sub

Public Function frmOpTrfD_GlobalControlàValider() As Boolean
Dim X1 As String, X2 As String

frmOpTrfD_GlobalControlàValider = False

If Trim(currentAction) = constSaisie Then
    If Trim(mId$(XopSwift(1).FieldText, 13, 11)) <> "" And XopCpt(2).Devise <> "001" Then
        Swift32CompteCr_ControlIban
    End If
End If

If XopCompteNostro.Situation = "B" Then
    If XopCpt(2).chkCompte < chkLevel Then
        x = MsgBox("Le compte à créditer est bloqué, confirmez-vous cette opération ?", vbYesNo + vbQuestion + vbDefaultButton2, XopCompteNostro.Intitulé & " : " & currentAction)
        If x = vbYes Then XopCpt(2).chkCompte = chkLevel
    End If
End If

If XopCpt(1).Compte > "00099999999" Then
    If mId$(XopCpt(1).Compte, 6, 3) <> "001" And mId$(XopCpt(1).Compte, 6, 3) <> "550" And mId$(XopCpt(1).Compte, 6, 3) <> "553" Then
        blnMsgBox = True
        x = MsgBox("Le type de compte à débiter n'est ni '001' ni '550'ni '553', confirmez-vous cette opération ?", vbYesNo + vbQuestion + vbDefaultButton2, Compte_Imp(XopCpt(1).Compte))
        If x <> vbYes Then
              Call lstErr_AddItem(lstErr, cmdContext, "Le type de compte à débiter n'est ni '001' ni '550'ni '553'")
        End If
    End If
End If


If XopCpt(2).Compte > "00099999999" Then
    If mId$(XopCpt(2).Compte, 6, 3) <> "001" And mId$(XopCpt(2).Compte, 6, 3) <> "550" And mId$(XopCpt(2).Compte, 6, 3) <> "553" Then
        blnMsgBox = True
        x = MsgBox("Le type de compte à créditer n'est ni '001' ni '550', confirmez-vous cette opération ?", vbYesNo + vbQuestion + vbDefaultButton2, Compte_Imp(XopCpt(2).Compte))
        If x <> vbYes Then
              Call lstErr_AddItem(lstErr, cmdContext, "Le type de compte  à créditer n'est ni '001' ni '550'")
        End If
    End If
End If

If XopCpt(1).AmjOpération <> DSys Or XopCpt(2).AmjOpération <> DSys Then
    blnMsgBox = True
    X1 = dateImp(DSys)
    X2 = dateImp(XopCpt(2).AmjOpération)
    x = MsgBox("Voulez-vous la modifier au " & X1 & " ?", vbYesNo + vbQuestion + vbDefaultButton1, "La date d'opération est le " & X2)
    If x = vbYes Then
        XopCpt(1).AmjOpération = DSys
        XopCpt(2).AmjOpération = DSys
    End If
End If

frmOpTrfD_GlobalControlàValider = frmOpTrfD_GlobalControl

If Trim(XopCpt(1).Libellé) = "" Then
      Call lstErr_AddItem(lstErr, cmdContext, "Préciser le libellé débit")
      frmOpTrfD_GlobalControlàValider = False
End If

If Trim(XopCpt(2).Libellé) = "" Then
      Call lstErr_AddItem(lstErr, cmdContext, "Préciser le libellé crédit")
      frmOpTrfD_GlobalControlàValider = False
End If

End Function

Public Function Swift32MontantDb_CV(Conversion As String)
Dim x As String
Swift32MontantDb_CV = Null
CV_Init CV1
CV1.OpéAmj = XopCpt(1).AmjOpération
CV2 = CV1
CV3 = CV_Euro

x = Trim(XopCpt(2).Devise)

If IsNumeric(x) Then
    CV1.DeviseN = x
Else
    CV1.DeviseIso = x
End If

CV1.Montant = XopCpt(2).Brut

x = Trim(XopCpt(1).Devise)
If IsNumeric(x) Then
    CV2.DeviseN = x
Else
    CV2.DeviseIso = x
End If

CV1.AchatVente = "A"
CV2.AchatVente = "V"

CV1.Normal = "C"
CV2.Normal = "C"

Swift32MontantDb_CV = CV_Transitoire(CV1, CV2, CV3, Conversion)

End Function
Public Function Swift32MontantCr_CV(Conversion As String)
Dim x As String
Swift32MontantCr_CV = Null
CV_Init CV1
CV1.OpéAmj = XopCpt(2).AmjOpération
CV2 = CV1
CV3 = CV_Euro

x = Trim(XopCpt(1).Devise)

If IsNumeric(x) Then
    CV1.DeviseN = x
Else
    CV1.DeviseIso = x
End If

CV1.Montant = XopCpt(1).Brut

x = Trim(XopCpt(2).Devise)
If IsNumeric(x) Then
    CV2.DeviseN = x
Else
    CV2.DeviseIso = x
End If

CV1.AchatVente = "V"
CV2.AchatVente = "A"

CV1.Normal = "C"
CV2.Normal = "C"

Swift32MontantCr_CV = CV_Transitoire(CV1, CV2, CV3, Conversion)

End Function

