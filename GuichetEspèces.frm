VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "Comct232.ocx"
Begin VB.Form frmGuichetEspèces 
   AutoRedraw      =   -1  'True
   Caption         =   "Guichet Espèces"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   9420
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   6120
      TabIndex        =   115
      Top             =   0
      Width           =   2745
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Picture         =   "GuichetEspèces.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   0
      Width           =   500
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   0
      TabIndex        =   21
      Top             =   1440
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8705
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "opération  en espèces"
      TabPicture(0)   =   "GuichetEspèces.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDevise1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDevise2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Détail des coupures"
      TabPicture(1)   =   "GuichetEspèces.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCoupureAuto"
      Tab(1).Control(1)=   "fraCoupure12"
      Tab(1).Control(2)=   "txtDevise1Rendu"
      Tab(1).Control(3)=   "fraCoupure"
      Tab(1).Control(4)=   "lblDevise1Rendu"
      Tab(1).Control(5)=   "libCoupure"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Comptabilité"
      TabPicture(2)   =   "GuichetEspèces.frx":013A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picCompta"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraCoupureAuto 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   109
         Top             =   3000
         Width           =   2055
         Begin VB.OptionButton optCoupureAuto 
            Caption         =   "Automatique"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   1200
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optCoupureManuel 
            Caption         =   "Manuel"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   800
            Width           =   1695
         End
         Begin VB.OptionButton optCoupureNéant 
            Caption         =   "Néant"
            Height          =   255
            Left            =   120
            TabIndex        =   110
            Top             =   400
            Width           =   1695
         End
      End
      Begin VB.Frame fraCoupure12 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   106
         Top             =   1680
         Width           =   2055
         Begin VB.OptionButton optCoupure2 
            Caption         =   "Coupure 2"
            Height          =   375
            Left            =   120
            TabIndex        =   108
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton optCoupure1 
            Caption         =   "Coupure 1"
            Height          =   375
            Left            =   120
            TabIndex        =   107
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.PictureBox picCompta 
         AutoRedraw      =   -1  'True
         FillColor       =   &H00E0E0E0&
         Height          =   4035
         Left            =   -74880
         ScaleHeight     =   3975
         ScaleWidth      =   9000
         TabIndex        =   104
         Top             =   720
         Width           =   9060
      End
      Begin VB.TextBox txtDevise1Rendu 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74040
         MaxLength       =   20
         TabIndex        =   99
         Top             =   1200
         Width           =   1140
      End
      Begin VB.Frame fraDevise2 
         Caption         =   "Devise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   5400
         TabIndex        =   94
         Top             =   480
         Width           =   3795
         Begin VB.CheckBox chkDevise2Montant 
            Caption         =   "Montant"
            Height          =   255
            Left            =   240
            TabIndex        =   119
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtDevise2Ajustement 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   19
            Top             =   720
            Width           =   1500
         End
         Begin VB.TextBox txtDevise2Montant 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   1
            Top             =   360
            Width           =   1500
         End
         Begin VB.Frame fraDevise2Cours 
            Caption         =   "cours de change"
            Height          =   2775
            Left            =   240
            TabIndex        =   95
            Top             =   1440
            Width           =   3375
            Begin VB.OptionButton optEnCompte 
               Caption         =   "En Compte"
               Height          =   255
               Left            =   240
               TabIndex        =   15
               Top             =   1200
               Value           =   -1  'True
               Width           =   1150
            End
            Begin VB.OptionButton optPrivilégié 
               Caption         =   "BME Privilégié"
               Height          =   255
               Left            =   240
               TabIndex        =   14
               Top             =   900
               Width           =   1335
            End
            Begin VB.OptionButton optNormal 
               Caption         =   "BME Normal"
               Height          =   255
               Left            =   240
               TabIndex        =   13
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton optPivot 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Pivot"
               Height          =   255
               Left            =   240
               TabIndex        =   12
               Top             =   300
               Width           =   855
            End
            Begin VB.Label lblCours2_1 
               Caption         =   "-"
               Height          =   255
               Left            =   1500
               TabIndex        =   117
               Top             =   2400
               Width           =   1700
            End
            Begin VB.Label lblCours1_2 
               Caption         =   "-"
               Height          =   255
               Left            =   1500
               TabIndex        =   116
               Top             =   2100
               Width           =   1700
            End
            Begin VB.Label libCours2 
               Caption         =   "x"
               Height          =   255
               Left            =   240
               TabIndex        =   103
               Top             =   1800
               Width           =   2895
            End
            Begin VB.Label libCours1 
               Caption         =   "x"
               Height          =   255
               Left            =   240
               TabIndex        =   102
               Top             =   1500
               Width           =   2895
            End
         End
         Begin VB.Label lblDevise2Ajustement 
            Caption         =   "Ajustement"
            Height          =   255
            Left            =   360
            TabIndex        =   96
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame fraDevise1 
         Caption         =   "Versement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   120
         TabIndex        =   91
         Top             =   480
         Width           =   5115
         Begin VB.CheckBox chkDevise1Montant 
            Caption         =   "Montant"
            Height          =   255
            Left            =   120
            TabIndex        =   118
            Top             =   360
            Width           =   1215
         End
         Begin VB.Frame fraRetrait 
            Caption         =   "Retrait"
            Height          =   1815
            Left            =   240
            TabIndex        =   100
            Top             =   2400
            Width           =   4695
            Begin VB.ListBox lstOppChq 
               Height          =   1035
               Left            =   3120
               TabIndex        =   113
               Top             =   600
               Width           =   1455
            End
            Begin VB.OptionButton optRetraitRmboursementFrais 
               Caption         =   "BIA : bon à payer en espèces "
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   1440
               Width           =   2895
            End
            Begin VB.OptionButton optDébitEnCompte 
               Caption         =   "débit en compte"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   1140
               Width           =   1575
            End
            Begin VB.OptionButton optRetraitMiseàDisposition 
               Caption         =   "mise à disposition / faveur tiers"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   840
               Width           =   2775
            End
            Begin VB.OptionButton optRetraitOmnibus 
               Caption         =   "chèque omnibus"
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   540
               Width           =   1575
            End
            Begin VB.OptionButton optRetraitChèque 
               Caption         =   "chèque client"
               Height          =   255
               Left            =   120
               TabIndex        =   7
               Top             =   240
               Value           =   -1  'True
               Width           =   1575
            End
            Begin VB.TextBox txtChèqueNo 
               Height          =   285
               Left            =   3120
               MaxLength       =   7
               TabIndex        =   2
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label lblOppChq 
               Caption         =   "OPPOSITION >>>"
               Height          =   255
               Left            =   1800
               TabIndex        =   114
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label lblChèqueNo 
               Caption         =   "N° chèque"
               Height          =   255
               Left            =   1800
               TabIndex        =   101
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.TextBox txtIdentité 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   3
            ToolTipText     =   "identité du remettant/bénéficiaire (libellé comptable)"
            Top             =   840
            Width           =   3465
         End
         Begin VB.TextBox txtComplément1 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   4
            ToolTipText     =   "complément d'informations (impression Avis)"
            Top             =   1260
            Width           =   3465
         End
         Begin VB.TextBox txtComplément2 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   5
            ToolTipText     =   "complément d'informations (impression Avis)"
            Top             =   1680
            Width           =   3465
         End
         Begin VB.TextBox txtComplément3 
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   6
            ToolTipText     =   "complément d'informations (impression Avis)"
            Top             =   2040
            Width           =   3465
         End
         Begin VB.TextBox txtDevise1Montant 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1440
            MaxLength       =   20
            TabIndex        =   0
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblIdentité 
            Caption         =   "Identité/faveur"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   885
            Width           =   1095
         End
         Begin VB.Label lblComplément 
            Caption         =   "Complément"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   1280
            Width           =   975
         End
      End
      Begin VB.Frame fraCoupure 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -72720
         TabIndex        =   22
         Top             =   720
         Width           =   6855
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   960
            MaxLength       =   4
            TabIndex        =   44
            Top             =   500
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   960
            MaxLength       =   4
            TabIndex        =   43
            Top             =   1000
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   960
            MaxLength       =   4
            TabIndex        =   42
            Top             =   1500
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   960
            MaxLength       =   4
            TabIndex        =   41
            Top             =   2000
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   960
            MaxLength       =   4
            TabIndex        =   40
            Top             =   2500
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   960
            MaxLength       =   4
            TabIndex        =   39
            Top             =   3000
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   2500
            MaxLength       =   4
            TabIndex        =   38
            Top             =   500
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   2500
            MaxLength       =   4
            TabIndex        =   37
            Top             =   1000
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   2500
            MaxLength       =   4
            TabIndex        =   36
            Top             =   1500
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   2500
            MaxLength       =   4
            TabIndex        =   35
            Top             =   2000
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   10
            Left            =   2500
            MaxLength       =   4
            TabIndex        =   34
            Top             =   2500
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   11
            Left            =   4200
            MaxLength       =   4
            TabIndex        =   33
            Top             =   500
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   12
            Left            =   4200
            MaxLength       =   4
            TabIndex        =   32
            Top             =   1000
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   13
            Left            =   4200
            MaxLength       =   4
            TabIndex        =   31
            Top             =   1500
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   4200
            MaxLength       =   4
            TabIndex        =   30
            Top             =   2000
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   4200
            MaxLength       =   4
            TabIndex        =   29
            Top             =   2500
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   16
            Left            =   4200
            MaxLength       =   4
            TabIndex        =   28
            Top             =   3000
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   17
            Left            =   4200
            MaxLength       =   4
            TabIndex        =   27
            Top             =   3500
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   18
            Left            =   5760
            MaxLength       =   4
            TabIndex        =   26
            Top             =   500
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   19
            Left            =   5760
            MaxLength       =   4
            TabIndex        =   25
            Top             =   1000
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   20
            Left            =   5760
            MaxLength       =   4
            TabIndex        =   24
            Top             =   1500
            Width           =   500
         End
         Begin VB.TextBox txtCoupureNb 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   5760
            MaxLength       =   4
            TabIndex        =   23
            Top             =   2000
            Width           =   500
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   350
            Index           =   0
            Left            =   1560
            TabIndex        =   45
            Top             =   500
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(0)"
            BuddyDispid     =   196654
            BuddyIndex      =   0
            OrigLeft        =   1320
            OrigTop         =   720
            OrigRight       =   1560
            OrigBottom      =   975
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   350
            Index           =   1
            Left            =   1560
            TabIndex        =   46
            Top             =   1000
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(1)"
            BuddyDispid     =   196654
            BuddyIndex      =   1
            OrigLeft        =   1300
            OrigTop         =   1050
            OrigRight       =   1540
            OrigBottom      =   1300
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   350
            Index           =   2
            Left            =   1560
            TabIndex        =   47
            Top             =   1500
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(2)"
            BuddyDispid     =   196654
            BuddyIndex      =   2
            OrigLeft        =   1320
            OrigTop         =   1400
            OrigRight       =   1560
            OrigBottom      =   1650
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   350
            Index           =   3
            Left            =   1560
            TabIndex        =   48
            Top             =   2000
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   635
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(3)"
            BuddyDispid     =   196654
            BuddyIndex      =   3
            OrigLeft        =   1300
            OrigTop         =   1750
            OrigRight       =   1540
            OrigBottom      =   2000
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   350
            Index           =   4
            Left            =   1560
            TabIndex        =   49
            Top             =   2500
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(4)"
            BuddyDispid     =   196654
            BuddyIndex      =   4
            OrigLeft        =   1300
            OrigTop         =   2100
            OrigRight       =   1540
            OrigBottom      =   2350
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   350
            Index           =   5
            Left            =   1560
            TabIndex        =   50
            Top             =   3000
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(5)"
            BuddyDispid     =   196654
            BuddyIndex      =   5
            OrigLeft        =   1320
            OrigTop         =   2400
            OrigRight       =   1560
            OrigBottom      =   2655
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   285
            Index           =   6
            Left            =   3100
            TabIndex        =   51
            Top             =   500
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(6)"
            BuddyDispid     =   196654
            BuddyIndex      =   6
            OrigLeft        =   1300
            OrigTop         =   2800
            OrigRight       =   1540
            OrigBottom      =   3050
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   285
            Index           =   7
            Left            =   3100
            TabIndex        =   52
            Top             =   1000
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(7)"
            BuddyDispid     =   196654
            BuddyIndex      =   7
            OrigLeft        =   1300
            OrigTop         =   3150
            OrigRight       =   1540
            OrigBottom      =   3400
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   285
            Index           =   8
            Left            =   3100
            TabIndex        =   53
            Top             =   1500
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(8)"
            BuddyDispid     =   196654
            BuddyIndex      =   8
            OrigLeft        =   1300
            OrigTop         =   3500
            OrigRight       =   1540
            OrigBottom      =   3750
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   285
            Index           =   9
            Left            =   3100
            TabIndex        =   54
            Top             =   2000
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(9)"
            BuddyDispid     =   196654
            BuddyIndex      =   9
            OrigLeft        =   1300
            OrigTop         =   3850
            OrigRight       =   1540
            OrigBottom      =   4100
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   285
            Index           =   10
            Left            =   3100
            TabIndex        =   55
            Top             =   2500
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(10)"
            BuddyDispid     =   196654
            BuddyIndex      =   10
            OrigLeft        =   1320
            OrigTop         =   4200
            OrigRight       =   1560
            OrigBottom      =   4450
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   345
            Index           =   11
            Left            =   4920
            TabIndex        =   56
            Top             =   500
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(11)"
            BuddyDispid     =   196654
            BuddyIndex      =   11
            OrigLeft        =   3000
            OrigTop         =   720
            OrigRight       =   3240
            OrigBottom      =   1005
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   345
            Index           =   12
            Left            =   4920
            TabIndex        =   57
            Top             =   1000
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(12)"
            BuddyDispid     =   196654
            BuddyIndex      =   12
            OrigLeft        =   3000
            OrigTop         =   1050
            OrigRight       =   3240
            OrigBottom      =   1335
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   345
            Index           =   13
            Left            =   4920
            TabIndex        =   58
            Top             =   1500
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(13)"
            BuddyDispid     =   196654
            BuddyIndex      =   13
            OrigLeft        =   3000
            OrigTop         =   1400
            OrigRight       =   3240
            OrigBottom      =   1685
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   345
            Index           =   14
            Left            =   4920
            TabIndex        =   59
            Top             =   2000
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(14)"
            BuddyDispid     =   196654
            BuddyIndex      =   14
            OrigLeft        =   3000
            OrigTop         =   1750
            OrigRight       =   3240
            OrigBottom      =   2035
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   345
            Index           =   15
            Left            =   4920
            TabIndex        =   60
            Top             =   2500
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(15)"
            BuddyDispid     =   196654
            BuddyIndex      =   15
            OrigLeft        =   3000
            OrigTop         =   2100
            OrigRight       =   3240
            OrigBottom      =   2385
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   345
            Index           =   16
            Left            =   4920
            TabIndex        =   61
            Top             =   3000
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(16)"
            BuddyDispid     =   196654
            BuddyIndex      =   16
            OrigLeft        =   3000
            OrigTop         =   2450
            OrigRight       =   3240
            OrigBottom      =   2735
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   345
            Index           =   17
            Left            =   4920
            TabIndex        =   62
            Top             =   3500
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(17)"
            BuddyDispid     =   196654
            BuddyIndex      =   17
            OrigLeft        =   3000
            OrigTop         =   2800
            OrigRight       =   3240
            OrigBottom      =   3085
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   345
            Index           =   18
            Left            =   6480
            TabIndex        =   63
            Top             =   500
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(18)"
            BuddyDispid     =   196654
            BuddyIndex      =   18
            OrigLeft        =   3000
            OrigTop         =   3120
            OrigRight       =   3240
            OrigBottom      =   3405
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   345
            Index           =   19
            Left            =   6480
            TabIndex        =   64
            Top             =   1000
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(19)"
            BuddyDispid     =   196654
            BuddyIndex      =   19
            OrigLeft        =   3000
            OrigTop         =   3480
            OrigRight       =   3240
            OrigBottom      =   3765
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   345
            Index           =   20
            Left            =   6480
            TabIndex        =   65
            Top             =   1500
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(20)"
            BuddyDispid     =   196654
            BuddyIndex      =   20
            OrigLeft        =   3000
            OrigTop         =   3840
            OrigRight       =   3240
            OrigBottom      =   4125
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown1Coupure 
            Height          =   345
            Index           =   21
            Left            =   6480
            TabIndex        =   66
            Top             =   2000
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   609
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtCoupureNb(21)"
            BuddyDispid     =   196654
            BuddyIndex      =   21
            OrigLeft        =   3000
            OrigTop         =   4200
            OrigRight       =   3240
            OrigBottom      =   4485
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   "500"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   90
            Top             =   500
            Width           =   720
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   "200"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   89
            Top             =   1000
            Width           =   600
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   "100"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   88
            Top             =   1500
            Width           =   600
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   "  50"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   87
            Top             =   2000
            Width           =   600
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   "  20"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   86
            Top             =   2500
            Width           =   600
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   "10"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   85
            Top             =   3000
            Width           =   600
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   ".........."
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   6
            Left            =   1900
            TabIndex        =   84
            Top             =   500
            Width           =   500
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   ".........."
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   7
            Left            =   1900
            TabIndex        =   83
            Top             =   1000
            Width           =   500
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   ".........."
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   8
            Left            =   1900
            TabIndex        =   82
            Top             =   1500
            Width           =   500
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   ".........."
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   9
            Left            =   1900
            TabIndex        =   81
            Top             =   2000
            Width           =   500
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   ".........."
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   10
            Left            =   1900
            TabIndex        =   80
            Top             =   2500
            Width           =   500
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   "100"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   11
            Left            =   3720
            TabIndex        =   79
            Top             =   500
            Width           =   405
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   "50"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   12
            Left            =   3720
            TabIndex        =   78
            Top             =   1000
            Width           =   405
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   "20"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   13
            Left            =   3720
            TabIndex        =   77
            Top             =   1500
            Width           =   405
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   "10"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   14
            Left            =   3720
            TabIndex        =   76
            Top             =   2000
            Width           =   405
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   "5"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   15
            Left            =   3720
            TabIndex        =   75
            Top             =   2500
            Width           =   405
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   "2"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   16
            Left            =   3720
            TabIndex        =   74
            Top             =   3000
            Width           =   405
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   17
            Left            =   3720
            TabIndex        =   73
            Top             =   3500
            Width           =   405
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   ".50"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   18
            Left            =   5280
            TabIndex        =   72
            Top             =   500
            Width           =   405
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   ".20"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   19
            Left            =   5280
            TabIndex        =   71
            Top             =   1000
            Width           =   405
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   ".10"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   20
            Left            =   5280
            TabIndex        =   70
            Top             =   1500
            Width           =   405
         End
         Begin VB.Label lblCoupureNb 
            Alignment       =   1  'Right Justify
            Caption         =   ".05"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   21
            Left            =   5280
            TabIndex        =   69
            Top             =   2000
            Width           =   405
         End
         Begin VB.Label lblCoupurePièces 
            Caption         =   "Pièces"
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
            Left            =   4320
            TabIndex        =   68
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblCoupureBillets 
            Caption         =   "Billets"
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
            Left            =   960
            TabIndex        =   67
            Top             =   200
            Width           =   495
         End
      End
      Begin VB.Label lblDevise1Rendu 
         Caption         =   "rendu"
         Height          =   255
         Left            =   -74760
         TabIndex        =   98
         Top             =   1200
         Width           =   405
      End
      Begin VB.Label libCoupure 
         Caption         =   "-"
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
         Left            =   -71160
         TabIndex        =   97
         Top             =   480
         Width           =   5295
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0FF&
      Caption         =   "en &Attente"
      Default         =   -1  'True
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
      TabIndex        =   18
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
      TabIndex        =   17
      Top             =   0
      Width           =   1200
   End
   Begin VB.PictureBox picCpt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   900
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   9315
      TabIndex        =   20
      Top             =   480
      Width           =   9375
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
      TabIndex        =   16
      Top             =   0
      Width           =   1200
   End
End
Attribute VB_Name = "frmGuichetEspèces"
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
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String

Dim kCours As Integer
' 1 = devise espèces

' 2 = devise compte

Dim ClientCompte As typeCompte
Dim CoupureNominal(2, 22) As Currency, CoupureSéquence(2, 22) As Integer, CoupureNature(2, 22) As String * 1
Dim optCoupure As Integer, optCoupureAjustement As String, strcoupure As String * 88
Dim blnCoupureAuto As Boolean, blnAjustementAuto As Boolean, blnDevise2CV As Boolean
Dim curCoupure As Currency, curCoupureMini(2) As Currency, curCoupureCalc As Currency
Dim curMontant As Currency, G_curMontant2 As Currency, mMontant As Currency
Dim chkLevel As String * 1


Dim CV As typeCV
Dim wG_CV1 As typeCV, wG_CV2 As typeCV, wG_CV3 As typeCV
Dim maxDevise1D As Integer, maxDevise2D As Integer
Dim xConversion As String * 1

Dim recTable As typeElpTable
Dim valAMJ As String, valAMJ1 As String, valAMJ2 As String



Dim blnCoupureCheck As Boolean

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

Private Sub frmCompte_Show()
Exit Sub ' $$$$$$$$$$$$$$$$$$$$$$$$$$ pb vbmodal

X = Space$(100)
Mid$(X, 1, 12) = "frmCompte   "
Mid$(X, 13, 12) = "frmGuichetEs"
Mid$(X, 25, 10) = Space$(10)
Mid$(X, 35, 3) = ClientCompte.Devise
Mid$(X, 38, 11) = ClientCompte.Numéro
Msg_Monitor X

End Sub

Public Sub cmdContext_Quit()
blnControl = False
If blnMsgBox_Quit Then
    X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
 Else
    X = vbYes
 End If
 If X = vbYes Then Unload Me

End Sub


Public Sub cmdContext_Return()

If cmdOk.Visible Then 'And ActiveControl.Name = lastActiveControl_Name Then
    cmdContext.SetFocus
    X = MsgBox("Voulez-vous enregistrer cette opération?", vbYesNo + vbQuestion + vbDefaultButton1, "Confirmation de saisie")
    If X = vbYes Then
        cmdOk_Click
    Else
        If txtDevise1Montant.Enabled Then txtDevise1Montant.SetFocus
    End If
Else
    SendKeys "{TAB}"
End If

End Sub






'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
'lstErr.Clear
currentActiveControl_Name = C.Name
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub
'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
'lstErr.Clear
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub


Private Sub chkDevise1Montant_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
chkDevise1Montant.Value = "1"
chkDevise2Montant.Value = "0"
chkDevise1Montant.ForeColor = libUsr.ForeColor
chkDevise2Montant.ForeColor = lblUsr.ForeColor
MouseMoveActiveControl.ForeColor = chkDevise1Montant.ForeColor
blnDevise2CV = False: blnAjustementAuto = True
txtDevise1Montant.Enabled = True
txtDevise2Montant.Enabled = False
If txtDevise1Montant.Visible Then txtDevise1Montant.SetFocus
If blnControl Then cmdControl

End Sub

Private Sub chkDevise2Montant_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
chkDevise1Montant.Value = "0"
chkDevise2Montant.Value = "1"
chkDevise2Montant.ForeColor = libUsr.ForeColor
chkDevise1Montant.ForeColor = lblUsr.ForeColor
MouseMoveActiveControl.ForeColor = chkDevise2Montant.ForeColor
blnDevise2CV = True: blnAjustementAuto = True
txtDevise1Montant.Enabled = False
txtDevise2Montant.Enabled = True
If txtDevise2Montant.Visible Then txtDevise2Montant.SetFocus
If blnControl Then cmdControl

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
Dim I As Integer, Msg As String, X As String

prtCompta.CV1 = G_CV1
prtCompta.CV2 = G_CV2
prtCompta.CV3 = G_CV3

For I = 1 To G_arrCV030Nb
    prtCompta.arrCV030(I) = G_arrCV030(I)
Next I

Msg = Format$(1, "000000") & Format$(G_arrCV030Nb, "000000")
Me.Hide
prtCompta_Monitor Msg, "", currentAction
Me.Show vbModal

End Sub


'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub




Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext
End Sub

Private Sub cmdOk_Click()
If Not blnCoupureCheck Then
    SSTab1.Tab = 1
Else
    G_recGuichet.ValidationUsr = ""
    cmdSave_Db
End If

End Sub

Private Sub cmdOk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdOk
End Sub


Private Sub cmdSave_Click()
G_recGuichet.ValidationUsr = constEnAttente
G_recGuichet.optAvis = "2"
cmdSave_Db

End Sub


Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdSave
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
End Sub





'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
usrColor_Set

txtDevise2Ajustement.Visible = False
lblDevise2Ajustement.Visible = False
lblOppChq.Visible = False
lblOppChq.ForeColor = vbRed
lstOppChq.Visible = False
lstOppChq.ForeColor = vbRed

'libCours1.ForeColor = warnUsrColor: libCours2.ForeColor = warnUsrColor
picCpt.Cls
SSTab1.Tab = 0
cmdContext.Caption = constcmdAbandonner: blnMsgBox_Quit = False
cmdOk.Visible = False: cmdSave.Visible = False
arrTag_Set False
lstErr.Visible = False
blnAddNew = True
blnCoupureAuto = True: fraCoupure.Enabled = True: fraCoupure_Clear: blnCoupureCheck = False
G_recGuichet.MontantEspèces = 0: G_curMontant2 = 0: G_recGuichet.Montant = 0
txtDevise2Ajustement = "": txtDevise2Ajustement.Enabled = False: blnAjustementAuto = True
txtDevise1Montant = ""

chkDevise1Montant.Value = "1"

'chkDevise1Montant_MouseDown

lastActiveControl_Name = "txtComplément3"
optPivot.Enabled = False
optNormal.Enabled = True
optPrivilégié.Enabled = True
optEnCompte.Enabled = False

chkLevel = "1"
lblCours1_2.ForeColor = warnUsrColor
lblCours2_1.ForeColor = warnUsrColor
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
End Sub

Public Function Coupure_Load(strDevX As String, maxDeviseD As Integer)
Dim I As Integer, K As Integer, Kb As Integer, Kp As Integer
Coupure_Load = Null
Coupure_Clear
srvDeviseCoupures_Load strDevX
K = 0: Kb = -1: Kp = 10
For I = 1 To arrDeviseCoupuresNb
    If arrDeviseCoupures(I).Actif = " " Then
        If arrDeviseCoupures(I).Nature = "B" Then
            Kb = Kb + 1: K = Kb
        Else
            Kp = Kp + 1: K = Kp
        End If
        curCoupureMini(optCoupure) = arrDeviseCoupures(I).Nominal
        CoupureNominal(optCoupure, K) = arrDeviseCoupures(I).Nominal
        CoupureSéquence(optCoupure, K) = arrDeviseCoupures(I).Séquence
        CoupureNature(optCoupure, K) = arrDeviseCoupures(I).Nature
    End If
Next I
If K = 0 Then
    MsgBox "Pas de coupures autorisées.", vbCritical, "Devise : " & strDevX
    Coupure_Load = "? coupures"
End If

If Kp = 10 Then maxDeviseD = 0

End Function
Public Function Coupure_Display()
Dim I As Integer, K As Integer, Kb As Integer, Kp As Integer, curX As Currency
Dim X1 As String * 1, Séq As Integer

blnCoupureAuto = True

If optCoupure = 1 Then
    X1 = G_recGuichet.chkCoupureEspèces
    strcoupure = G_recGuichet.CoupureEspèces
    curCoupureCalc = G_recGuichet.MontantEspèces + G_recGuichet.MontantRendu
Else
    X1 = G_recGuichet.chkCoupureChange
    strcoupure = G_recGuichet.CoupureChange
    curCoupureCalc = G_recGuichet.Montant
End If

Select Case X1
    Case "M": optCoupureManuel = True
    Case "A": optCoupureAuto = True
    Case Else: optCoupureNéant = True
End Select

K = 0: Kb = -1: Kp = 10
For K = 0 To 21
    txtCoupureNb(K).Visible = False
    lblCoupureNb(K).Visible = False
    UpDown1Coupure(K).Visible = False
Next K

For I = 0 To 21
    If CoupureNature(optCoupure, I) <> " " Then
        If CoupureNature(optCoupure, I) = "B" Then
            Kb = Kb + 1: K = Kb
        Else
            Kp = Kp + 1: K = Kp
        End If
        txtCoupureNb(K).Visible = True
        txtCoupureNb(K) = ""
        lblCoupureNb(K).Visible = True
        curX = CoupureNominal(optCoupure, I)
        
        If curX = Fix(curX) Then
            lblCoupureNb(K).Caption = Format$(curX, "###")
        Else
             lblCoupureNb(K).Caption = Format$(curX, "##0.00")
        End If
        
        Séq = CoupureSéquence(optCoupure, I) * 4 - 3
        txtCoupureNb(K) = Format$(mId$(strcoupure, Séq, 4), "####")
       
        UpDown1Coupure(K).Visible = True
    End If
Next I
blnCoupureAuto = False
Coupure_Control
End Function


Public Sub Coupure_Clear()
Dim I As Integer
For I = 0 To 21
    txtCoupureNb(I).Visible = False
    lblCoupureNb(I).Visible = False
    UpDown1Coupure(I).Visible = False
    CoupureNominal(optCoupure, I) = 0
    CoupureSéquence(optCoupure, I) = 0
    CoupureNature(optCoupure, I) = " "
Next I
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Private Sub fraComptabilité_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraCoupure_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraDevise1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraDevise2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraDevise2Cours_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraRetrait_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub chkDevise1Montant_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkDevise1Montant
End Sub


Private Sub chkDevise2Montant_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkDevise2Montant
End Sub


Private Sub optCoupure1_Click()
optCoupure = 1
Coupure_Display
End Sub

Private Sub optCoupure2_Click()
optCoupure = 2
Coupure_Display
End Sub


Private Sub optCoupureAuto_Click()
lstErr.Clear
fraCoupure.Enabled = False
If optCoupure = 1 Then
    G_recGuichet.chkCoupureEspèces = "A"
Else
    G_recGuichet.chkCoupureChange = "A"
End If

Coupure_Calc
End Sub

Private Sub optCoupureManuel_Click()
lstErr.Clear
fraCoupure.Enabled = True
If optCoupure = 1 Then
    G_recGuichet.chkCoupureEspèces = "M"
Else
    G_recGuichet.chkCoupureChange = "M"
End If

Coupure_Calc
End Sub


Private Sub optCoupureNéant_Click()
lstErr.Clear
fraCoupure.Enabled = False
If optCoupure = 1 Then
    G_recGuichet.chkCoupureEspèces = " "
Else
    G_recGuichet.chkCoupureChange = " "
End If

Coupure_Calc
End Sub


Private Sub optEnCompte_Click()
If blnControl Then cmdControl
End Sub

Private Sub optEnCompte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optEnCompte
End Sub


Private Sub optNormal_Click()
If blnControl Then cmdControl
End Sub


Private Sub optNormal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optNormal
End Sub


Private Sub optPivot_Click()
If blnControl Then cmdControl
End Sub


Private Sub optPivot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optPivot
End Sub


Private Sub optPrivilégié_Click()
If blnControl Then cmdControl
End Sub


Private Sub optPrivilégié_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optPrivilégié
End Sub


Private Sub optRetraitChèque_Click()
txtChèqueNo.Enabled = True
G_recGuichet.chkChèque = "1"
If blnControl Then cmdControl
End Sub

Public Sub cmdControl()
Dim X As String
Dim valX As String
Dim valDevise2Montant As Currency, curX As Currency
Dim dblX As Double
Dim C As Control, blnFocus As Boolean
Dim C_name As String

If Not frmGuichetEspèces.Enabled Then Exit Sub
frmGuichetEspèces.Enabled = False

cmdOk.Visible = False
cmdSave.Visible = False
blnControl = False
blnFocus = False
C_name = ""

'For Each C In Me.Controls
'    If TypeOf C Is TextBox Then
'        If C.Name = currentActiveControl_Name Then
'            blnFocus = True
'            Exit For
'        End If
'    End If
'Next C

lstErr.Clear
lstErr.Height = 200
picCompta.Cls


G_recGuichet.Identité = Trim(txtIdentité)
G_recGuichet.Complément1 = Trim(txtComplément1)
G_recGuichet.Complément2 = Trim(txtComplément2)
G_recGuichet.Complément3 = Trim(txtComplément3)

Select Case G_recGuichet.chkChèque
    Case "0"
    Case "1", "2": If Not txtChèqueNo_Control Then C_name = "txtChèqueNo"
    Case "3", "5"
                If Trim(G_recGuichet.Identité) = "" Then C_name = "txtIdentité": Call lstErr_AddItem(lstErr, txtIdentité, "? identité")
                If Trim(G_recGuichet.Complément1) = "" Then C_name = "txtComplément1": Call lstErr_AddItem(lstErr, txtComplément1, "? complément")
End Select


If currentAction = constChange Then If Trim(G_recGuichet.Identité) = "" Then C_name = "txtIdentité": Call lstErr_AddItem(lstErr, txtIdentité, "? identité")

G_CV1.Normal = " "
If G_CV1.DeviseIso = G_CV2.DeviseIso Then
    G_CV2.Normal = " "
Else
    If optNormal Then
        G_CV1.Normal = "N"
    Else
        If optPrivilégié Then
            G_CV1.Normal = "P"
        Else
            If optEnCompte Then
                G_CV1.Normal = "C"
            End If
        End If
    End If
End If
Select Case currentAction
    Case constChange: G_CV2.Normal = G_CV1.Normal
    Case constArbitrage: G_CV2.Normal = G_CV1.Normal
End Select

mMontant = G_recGuichet.Montant
X = num_Control(txtDevise1Rendu, valX, 13, maxDevise1D)
G_recGuichet.MontantRendu = CCur(valX)
If G_recGuichet.MontantRendu <> 0 Then Coupure_Ajustement G_recGuichet.MontantRendu, curCoupureMini(1), optCoupureAjustement

If Not blnDevise2CV Then
    X = num_Control(txtDevise1Montant, valX, 13, 2)
    G_CV1.Montant = valX - G_recGuichet.MontantRendu
    Coupure_Ajustement G_CV1.Montant, curCoupureMini(1), optCoupureAjustement
    If G_CV1.Montant <= 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? montant"): GoTo ExitSub
    
    Call CV_Transitoire(G_CV1, G_CV2, G_CV3, xConversion)
Else
    X = num_Control(txtDevise2Montant, valX, 13, 2)
    G_CV2.Montant = valX
    If G_CV2.Montant = 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? montant"): GoTo ExitSub
    If currentAction = constChange Then Coupure_Ajustement G_CV2.Montant, curCoupureMini(2), "-"
    Call CV_Transitoire(G_CV2, G_CV1, G_CV3, xConversion)
End If

If Not G_CV1.EuroIn And valAMJ1 <> G_CV1.CoursAmj Then
    X = MsgBox("! cours " & G_CV1.DeviseIso & " au " & dateImp(G_CV1.CoursAmj) & "; confirmez-vous ?", vbQuestion + vbYesNo, "Contre-Valeur : contrôle date du cours ")
    If X = vbYes Then
        valAMJ1 = G_CV1.CoursAmj
    Else
        Call lstErr_AddItem(lstErr, cmdContext, "? date du cours " & G_CV1.DeviseIso)
    End If
End If


If Not G_CV2.EuroIn And valAMJ2 <> G_CV2.CoursAmj Then
    X = MsgBox("! cours " & G_CV2.DeviseIso & " au " & dateImp(G_CV2.CoursAmj) & "; confirmez-vous ?", vbQuestion + vbYesNo, "Contre-Valeur : contrôle date du cours ")
    If X = vbYes Then
        valAMJ2 = G_CV2.CoursAmj
    Else
        Call lstErr_AddItem(lstErr, cmdContext, "? date du cours " & G_CV2.DeviseIso)
    End If
End If

G_recGuichet.MontantEspèces = G_CV1.Montant
G_recGuichet.Montant = G_CV2.Montant
G_recGuichet.MontantAjustement = 0

G_recGuichet.CoursChangeEspèces = G_CV1.Cours
G_recGuichet.CoursChange = G_CV2.Cours

If blnDevise2CV Then
    wG_CV1 = G_CV1
    wG_CV2 = G_CV2
    wG_CV3 = G_CV3
    Coupure_Ajustement wG_CV1.Montant, curCoupureMini(1), optCoupureAjustement
    If G_CV1.Montant <> wG_CV1.Montant Then
        Call CV_Transitoire(wG_CV1, wG_CV2, wG_CV3, xConversion)
        G_CV3 = wG_CV3
        G_recGuichet.MontantEspèces = wG_CV1.Montant
'        G_recGuichet.MontantAjustement = Abs(wG_CV2.Montant - G_CV2.Montant)
    End If
Else
    If currentAction = constChange Then
        Coupure_Ajustement G_recGuichet.Montant, curCoupureMini(2), "-"
'        G_recGuichet.MontantAjustement = Abs(G_CV2.Montant - G_recGuichet.Montant)
   End If
End If

curMontant = G_recGuichet.MontantEspèces + G_recGuichet.MontantRendu
txtDevise1Montant = Format$(curMontant, "### ### ### ##0.00")
txtDevise2Montant = Format$(G_recGuichet.Montant, "### ### ### ##0.00")
txtDevise2Ajustement = Format$(G_recGuichet.MontantAjustement, "### ### ### ##0.00")

If currentAction = constChange Then optCoupure = 2: Coupure_Calc

optCoupure = 1: If currentAction <> constArbitrage Then Coupure_Calc

'''If G_recGuichet.MontantRendu > G_recGuichet.MontantEspèces Then Call lstErr_AddItem(lstErr, txtDevise1Rendu, "? rendu > montant")

libCours1 = dateImp(G_CV1.CoursAmj) & " : " & G_CV3.DeviseIso & " / " & G_CV1.DeviseIso & " : " & Format$(G_CV1.Cours, "## ##0.00000")
libCours2 = dateImp(G_CV2.CoursAmj) & " : " & G_CV3.DeviseIso & " / " & G_CV2.DeviseIso & " : " & Format$(G_CV2.Cours, "## ##0.00000")
lblCours1_2 = G_CV1.DeviseIso & " / " & G_CV2.DeviseIso & " : " & Format$(G_CV2.Cours / G_CV1.Cours, "## ##0.00000")
lblCours2_1 = G_CV2.DeviseIso & " / " & G_CV1.DeviseIso & " : " & Format$(G_CV1.Cours / G_CV2.Cours, "## ##0.00000")

G_recGuichet.MontantEuro = G_CV3.Montant
G_recGuichet.Conversion = xConversion


If mMontant <> G_recGuichet.Montant Then G_recGuichet.chkSolde = "0"
Compte_Control


If G_recGuichet.chkChèque = "5" Then
    If G_CV1.DeviseIso <> "FRF" Or G_CV2.DeviseIso <> "FRF" Then Call lstErr_AddItem(lstErr, cmdContext, "? FRF uniquement"): GoTo ExitSub
    If ClientCompte.TypeGA <> "G" Then Call lstErr_AddItem(lstErr, cmdContext, "? compte classe 6 uniquement"): GoTo ExitSub
    If ClientCompte.Numéro < "00060000000" Or ClientCompte.Numéro > "00069999999" Then Call lstErr_AddItem(lstErr, cmdContext, "? compte classe 6 uniquement"): GoTo ExitSub
    If Not IsNumeric(mId$(Trim(G_recGuichet.Complément1), 1, 3)) Then Call lstErr_AddItem(lstErr, cmdContext, "? code pays *** dans complément"): GoTo ExitSub
End If

If G_recGuichet.Devise = G_recGuichet.DeviseEspèces Then
    If G_recGuichet.Montant <> G_recGuichet.MontantEspèces Then
        G_recGuichet.MontantAjustement = G_recGuichet.Montant - G_recGuichet.MontantEspèces
        If G_recGuichet.Sens = "D" Then G_recGuichet.MontantAjustement = -G_recGuichet.MontantAjustement
    End If
End If

Guichet_Compta.Init
Guichet_Compta.Libellé
Guichet_Compta.Gen

Guichet_Compta.Display picCompta, lstErr

G_recGuichet.optAvis = "2"
If G_recGuichet.chkChèque = "1" Then G_recGuichet.optAvis = "0"

If lstErr.ListCount = 0 Then
    cmdOk.Visible = True: C_name = "cmdOk"
    If G_recGuichet.chkCompte <> "0" Or G_recGuichet.chkSolde <> "0" Then cmdSave.Visible = True
End If

ExitSub:

frmGuichetEspèces.Enabled = True
If cmdOk.Visible And cmdOk.Enabled Then cmdOk.Visible = False: cmdOk.Visible = True
    
'Select Case C_name
'    Case "cmdOk":
'If cmdOk.Visible And cmdOk.Enabled Then cmdOk.SetFocus
'    Case "txtChèqueNo": If txtChèqueNo.Visible And txtChèqueNo.Enabled Then txtChèqueNo.SetFocus
'    Case "txtIdentité": If txtIdentité.Visible And txtIdentité.Enabled Then txtIdentité.SetFocus
'    Case "txtComplément1": If txtComplément1.Visible And txtComplément1.Enabled Then txtComplément1.SetFocus
'End Select

blnControl = True

End Sub

Private Sub optRetraitChèque_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optRetraitChèque
End Sub

Private Sub optDébitEnCompte_Click()
txtChèqueNo.Enabled = False
G_recGuichet.chkChèque = "4"
If blnControl Then cmdControl
End Sub

Private Sub optDébitencompte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDébitEnCompte
End Sub


Private Sub optRetraitMiseàDisposition_Click()
txtChèqueNo.Enabled = False
G_recGuichet.chkChèque = "3"
If blnControl Then cmdControl
End Sub


Private Sub optRetraitMiseàDisposition_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optRetraitMiseàDisposition
End Sub


Private Sub optRetraitOmnibus_Click()
txtChèqueNo.Enabled = True
G_recGuichet.chkChèque = "2"
If blnControl Then cmdControl
End Sub


Private Sub optRetraitOmnibus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optRetraitOmnibus
End Sub


Private Sub optRetraitRmboursementFrais_Click()
txtChèqueNo.Enabled = False
G_recGuichet.chkChèque = "5"
If blnControl Then cmdControl
End Sub

Private Sub optRetraitRmboursementFrais_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optRetraitRmboursementFrais
End Sub


Private Sub picCpt_Click()
frmCompte_Show
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If PreviousTab = 0 And blnControl Then cmdControl
If SSTab1.Tab = 1 Then blnCoupureCheck = True
End Sub

Private Sub SSTab1_LostFocus()
''''If blnControl Then cmdControl

End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set SSTab1
End Sub


Private Sub txtChèqueNo_GotFocus()
Call txt_GotFocus(txtChèqueNo)
End Sub


Private Sub txtChèqueNo_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtChèqueNo_LostFocus()
Call txt_LostFocus(txtChèqueNo)
If blnControl Then cmdControl
End Sub


Private Sub txtComplément1_GotFocus()
Call txt_GotFocus(txtComplément1)
End Sub


Private Sub txtComplément1_LostFocus()
Call txt_LostFocus(txtComplément1)
If blnControl Then cmdControl
End Sub


Private Sub txtComplément2_GotFocus()
Call txt_GotFocus(txtComplément2)
End Sub


Private Sub txtComplément2_LostFocus()
Call txt_LostFocus(txtComplément2)
If blnControl Then cmdControl
End Sub


Private Sub txtComplément3_GotFocus()
Call txt_GotFocus(txtComplément3)
End Sub


Private Sub txtComplément3_LostFocus()
Call txt_LostFocus(txtComplément3)
If blnControl Then cmdControl
End Sub


Private Sub txtCoupureNb_Change(Index As Integer)
If Not blnCoupureAuto Then
    Coupure_Control
    If blnControl And lstErr.ListCount = 0 Then cmdControl
End If
End Sub

Private Sub txtCoupureNb_GotFocus(Index As Integer)
Call txt_GotFocus(txtCoupureNb(Index))
End Sub


Private Sub txtCoupureNb_LostFocus(Index As Integer)
Call txt_LostFocus(txtCoupureNb(Index))
lstErr.Clear
Coupure_Control
If blnControl And lstErr.ListCount = 0 Then cmdControl
End Sub


Private Sub txtDevise1Montant_GotFocus()
Call txt_GotFocus(txtDevise1Montant)
End Sub


Private Sub txtDevise1Montant_KeyPress(KeyAscii As Integer)
If G_CV1.maxD = 0 Then
    Call num_KeyAscii(KeyAscii)
Else
    Call num_KeyAsciiD(KeyAscii, txtDevise1Montant)
End If

End Sub


Private Sub txtDevise1Montant_LostFocus()
Call txt_LostFocus(txtDevise1Montant)
If blnControl Then cmdControl
End Sub

Private Sub txtDevise1Rendu_GotFocus()
Call txt_GotFocus(txtDevise1Rendu)
End Sub


Private Sub txtDevise1Rendu_KeyPress(KeyAscii As Integer)
If maxDevise1D = 0 Then
    Call num_KeyAscii(KeyAscii)
Else
    Call num_KeyAsciiD(KeyAscii, txtDevise1Rendu)
End If

End Sub

Private Sub txtDevise1Rendu_LostFocus()
Call txt_LostFocus(txtDevise1Rendu)
If blnControl Then cmdControl
End Sub


Private Sub txtDevise2Ajustement_GotFocus()
Call txt_GotFocus(txtDevise2Ajustement)
End Sub


Private Sub txtDevise2Ajustement_KeyPress(KeyAscii As Integer)
If G_CV2.maxD = 0 Then
    Call num_KeyAscii(KeyAscii)
Else
    Call num_KeyAsciiD(KeyAscii, txtDevise2Ajustement)
End If
End Sub

Private Sub txtDevise2Ajustement_LostFocus()
Call txt_LostFocus(txtDevise2Ajustement)
End Sub


Private Sub txtDevise2Montant_GotFocus()
Call txt_GotFocus(txtDevise2Montant)
End Sub


Private Sub txtDevise2Montant_KeyPress(KeyAscii As Integer)
If G_CV2.maxD = 0 Then
    Call num_KeyAscii(KeyAscii)
Else
    Call num_KeyAsciiD(KeyAscii, txtDevise2Montant)
End If
End Sub

Private Sub txtDevise2Montant_LostFocus()
Call txt_LostFocus(txtDevise2Montant)
If blnControl Then cmdControl

End Sub


Private Sub txtIdentité_GotFocus()
Call txt_GotFocus(txtIdentité)
End Sub


Private Sub txtIdentité_LostFocus()
Call txt_LostFocus(txtIdentité)
If blnControl Then cmdControl
End Sub



Public Sub Coupure_Control()
Dim Nb As Integer, Séq As Integer

If blnCoupureAuto Then Exit Sub

curCoupure = 0
strcoupure = String$("0", 88)
For I = 0 To 21
    Nb = Val(txtCoupureNb(I))
    Séq = CoupureSéquence(optCoupure, I) * 4 - 3
    If Séq > 0 Then Mid$(strcoupure, Séq, 4) = Format$(Nb, "0000")
    If Nb = 0 Then
        lblCoupureNb(I).ForeColor = lblUsr.ForeColor
        lblCoupureNb(I).Font.Bold = False
    Else
        lblCoupureNb(I).ForeColor = libUsr.ForeColor
        lblCoupureNb(I).Font.Bold = True
        curCoupure = curCoupure + Nb * CoupureNominal(optCoupure, I)
    End If
Next I
libCoupure_Text curCoupure
libCoupure.ForeColor = libUsr.ForeColor
If optCoupure = 1 Then
    G_recGuichet.CoupureEspèces = strcoupure
    curMontant = curCoupure - G_recGuichet.MontantRendu - G_recGuichet.MontantEspèces
Else
    G_recGuichet.CoupureChange = strcoupure
    curMontant = curCoupure - G_recGuichet.Montant
End If

If curCoupure <> 0 Then
    If curMontant <> 0 Then
        cmdOk.Visible = False
        libCoupure.ForeColor = warnUsrColor
        Call lstErr_AddItem(lstErr, fraCoupure, "Ecart coupures : " & Trim(Format$(curMontant, "##### ### ### ##0.00-")))
    End If
End If
End Sub

Public Sub txtDevise1Rendu_Control()
End Sub

Private Sub UpDown1Coupure_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lstErr.Clear
End Sub


Public Sub Form_Init(Msg As String)
tableElpTable_Open
currentAction = Trim(mId$(Msg, 25, 12))

Guichet_Compta.CV_Reset currentAction

G_CV1.DeviseIso = mId$(Msg, 38, 3)
If Not IsNull(CV_Attribut(G_CV1)) Then Call MsgBox("Devise1 inconnue: " & mId$(Msg, 38, 3), vbCritical, "frmGuichetEspèces.Form_Init")
G_CV2.DeviseIso = mId$(Msg, 41, 3)
If Not IsNull(CV_Attribut(G_CV2)) Then Call MsgBox("Devise2 inconnue: " & mId$(Msg, 41, 3), vbCritical, "frmGuichetEspèces.Form_Init")

cmdReset

G_recGuichet_Init mId$(Msg, 44, 11)

G_CV1.OpéAmj = G_recGuichet.AmjOpération
G_CV1.OpéAmj = G_recGuichet.AmjOpération
valAMJ1 = G_CV1.OpéAmj: valAMJ2 = G_CV1.OpéAmj

maxDevise1D = G_CV1.maxD
maxDevise2D = G_CV2.maxD
If G_CV1.DeviseIso = G_CV2.DeviseIso Then
    optPivot = True
    fraDevise2Cours.Visible = False
Else
    optNormal = True
    fraDevise2Cours.Enabled = True
End If

curCoupureMini(1) = 0.01: curCoupureMini(2) = 0.01
optCoupure1.Caption = G_CV1.DeviseLibellé
optCoupure = 1: Coupure_Clear

If currentAction = constArbitrage Then
    optCoupureNéant = False
    optEnCompte = True
    blnCoupureCheck = True
    lblIdentité.Visible = False
    lblComplément.Visible = False
    txtIdentité.Visible = False
    txtComplément1.Visible = False
    txtComplément2.Visible = False
    txtComplément3.Visible = False
    
Else
    If Not IsNull(Coupure_Load(G_CV1.DeviseIso, maxDevise1D)) Then Exit Sub
End If

optCoupure2.Caption = G_CV2.DeviseLibellé
If currentAction = constChange Then
    picCpt.Cls
    optCoupure1.Enabled = True
    optCoupure2.Enabled = True
    optCoupure = 2
    If Not IsNull(Coupure_Load(G_CV2.DeviseIso, maxDevise2D)) Then Exit Sub
    Compte_Load
Else
    optCoupure1.Enabled = False
    optCoupure2.Enabled = False
    Compte_Load
    recCompte_Display ClientCompte, picCpt
End If
optCoupure1_Click
cmdControl
frmGuichetEspèces.Show vbModal

End Sub


Public Sub fraCoupure_Clear()
For I = 0 To 21
    txtCoupureNb(I) = ""
    lblCoupureNb(I).ForeColor = lblUsr.ForeColor
    lblCoupureNb(I).Font.Bold = False
Next I
curCoupure = 0:
G_recGuichet.MontantEspèces = G_recGuichet.Montant
txtDevise1Rendu = ""
G_recGuichet.MontantRendu = 0
libCoupure_Text curCoupure
End Sub

Public Function Compte_Control() As String
Dim X As String
X = vbYes
If ClientCompte.Situation <> " " Then
    If G_recGuichet.chkCompte < chkLevel Then
        X = MsgBox("Le compte est bloqué, confirmez-vous cette opération ?", vbYesNo + vbQuestion + vbDefaultButton2, Trim(ClientCompte.Intitulé) & " : " & currentAction)
        If X = vbYes Then
            G_recGuichet.chkCompte = chkLevel
        Else
            Call lstErr_AddItem(lstErr, txtDevise1Montant, "Le compte est bloqué")
        End If
    End If
End If
If ClientCompte.TypeGA = "A" Then
    If currentAction = constRetrait Or currentAction = constArbitrage Then
        If ClientCompte.SoldeXXX - G_recGuichet.Montant < 0 Then
            If G_recGuichet.chkSolde < chkLevel Then
                X = MsgBox("Le compte à débiter n'a pas la provision suffisante, confirmez-vous ce montant ?" & Chr$(13) & "Dépassement :" & num_Display(ClientCompte.SoldeXXX - G_recGuichet.Montant, 15, 2, Lx, X, "0"), vbYesNo + vbQuestion + vbDefaultButton2, Trim(ClientCompte.Intitulé))
                If X = vbYes Then
                    G_recGuichet.chkSolde = chkLevel
                Else
                    Call lstErr_AddItem(lstErr, txtDevise1Montant, "provision insuffisante")
                End If
            End If
        End If
    End If
End If
Compte_Control = X

End Function

Public Sub G_recGuichet_Init(xNuméro As String)
recGuichet_Init G_recGuichet
G_recGuichet.Method = constAddNew
G_recGuichet.Séquence = 1
G_recGuichet.Société = SocId$
G_recGuichet.Agence = SocAgence$
G_recGuichet.Journal = constCaisse
G_recGuichet.Devise = G_CV2.DeviseN
G_recGuichet.Compte = xNuméro
G_recGuichet.AmjOpération = paramGuichetAMJValeur
G_recGuichet.AmjValeur = paramGuichetAMJValeur
G_recGuichet.chkCompte = "0"
G_recGuichet.chkSolde = "0"
G_recGuichet.chkAmjOpération = "0"
G_recGuichet.chkAmjValeur = "0"
G_recGuichet.optAvis = "2"
G_recGuichet.optVirement = "0"
G_recGuichet.optSwift = "0"
G_recGuichet.optAvisLangue = "1"

G_recGuichet.CptMvtPièce = 0
G_recGuichet.CptMvtLigne = 0
G_recGuichet.CptMvtService = paramGuichetService
G_recGuichet.CptMvtExonéré = "0"
G_recGuichet.chkChèque = "0"
G_recGuichet.NoChèque = ""
G_recGuichet.CoursChange = 1
G_recGuichet.CoursChangeEspèces = 1

G_recGuichet.optCours = "0"
G_recGuichet.DeviseEspèces = G_CV1.DeviseN

G_recGuichet.chkCoupureEspèces = "A"
G_recGuichet.CoupureEspèces = String$("0", 88)
G_recGuichet.chkCoupureChange = " "
G_recGuichet.CoupureChange = String$("0", 88)

G_recGuichet.SaisieAmj = DSys
G_recGuichet.SaisieHMS = time_Hms
G_recGuichet.SaisieUsr = usrId

optCoupureAjustement = "+"

Select Case currentAction
    Case constRetrait:  G_recGuichet.Sens = "D"
                        FraDevise1.Caption = "Retrait : " & Trim(G_CV1.DeviseLibellé)
                        fraDevise2.Caption = "Compte : " & Trim(G_CV2.DeviseLibellé)
                        fraRetrait.Visible = True
                        optRetraitChèque = True
                        lblIdentité.Caption = "Bénéficiaire"
                        G_recGuichet.chkChèque = "1"
                        If G_recGuichet.DeviseEspèces = G_recGuichet.Devise Then  '''' 2001.07.25 jpl Or G_recGuichet.DeviseEspèces = "001" Then
                            G_recGuichet.CodeOpération = "G002"
                            SSTab1.Caption = "Retrait en compte : " & Trim(G_CV1.DeviseIso)
                        Else
                            G_recGuichet.CodeOpération = "G005"
                            SSTab1.Caption = "Délivrance de devises : " & G_CV1.DeviseIso & " / " & G_CV2.DeviseIso
                        End If
                        optCoupureAjustement = "-"
                        txtDevise1Rendu.Visible = False: lblDevise1Rendu.Visible = False
                        Call OppChq_Load(G_recGuichet.Compte, lstOppChq)
                        If G_arrOppChq_Numéro_Nb > 0 Then lstOppChq.Visible = True: lblOppChq.Visible = True

        Case constArbitrage:  G_recGuichet.Sens = "D"
                        G_recGuichet.ContrepartieCompte = G_recGuichet.Compte
                        FraDevise1.Caption = "Compte à créditer : " & Trim(G_CV1.DeviseLibellé)
                        fraDevise2.Caption = "Compte à débiter: " & Trim(G_CV2.DeviseLibellé)
                        fraRetrait.Visible = True
                        lblIdentité.Caption = "Bénéficiaire"
                        SSTab1.Caption = "Arbitrage : " & Trim(G_CV1.DeviseIso) & " / " & Trim(G_CV2.DeviseIso)
                        G_recGuichet.CodeOpération = "G008"
                        G_recGuichet.Journal = constJournalGuichet
                        optCoupureAjustement = "-"
                        txtDevise1Rendu.Visible = False: lblDevise1Rendu.Visible = False
                        optPivot.Enabled = True
                        optNormal.Enabled = False
                        optPrivilégié.Enabled = False
                        optEnCompte.Enabled = True: optEnCompte = True
                        fraCoupure.Visible = False: fraCoupure12.Visible = False: fraCoupureAuto.Visible = False
                        fraRetrait.Visible = False
                        
        Case constVersement: G_recGuichet.Sens = "C"
                        FraDevise1.Caption = "Versement : " & Trim(G_CV1.DeviseLibellé)
                        fraDevise2.Caption = "Compte : " & Trim(G_CV2.DeviseLibellé)
                        fraRetrait.Visible = False
                       lblIdentité.Caption = "Déposant"
                        If G_recGuichet.DeviseEspèces = G_recGuichet.Devise Then
                           G_recGuichet.CodeOpération = "G001"
                            SSTab1.Caption = "Versement en compte : " & Trim(G_CV1.DeviseIso)
                        Else
                           G_recGuichet.CodeOpération = "G006"
                           SSTab1.Caption = "Versement de devises : " & Trim(G_CV1.DeviseIso) & " / " & Trim(G_CV2.DeviseIso)
                        End If
                         txtDevise1Rendu.Visible = True: lblDevise1Rendu.Visible = True
   Case constChange: G_recGuichet.CodeOpération = "G007": G_recGuichet.Sens = "C" '"D" '"C"
                        FraDevise1.Caption = "Versement : " & Trim(G_CV1.DeviseLibellé)
                        fraDevise2.Caption = "Retrait : " & Trim(G_CV2.DeviseLibellé)
                        fraRetrait.Visible = False
                        lblIdentité.Caption = "Identité"
                        SSTab1.Caption = "Change : " & Trim(G_CV1.DeviseIso) & " / " & Trim(G_CV2.DeviseIso)
                        G_recGuichet.chkCoupureChange = "A"
                        optCoupureAjustement = "-"
                        txtDevise1Rendu.Visible = False: lblDevise1Rendu.Visible = False
End Select

End Sub

Public Sub Coupure_Calc()
Dim dblNb As Double, chkCoupureX As String

If blnCoupureAuto Then Exit Sub
blnCoupureAuto = True

If optCoupure = 1 Then
    curMontant = G_recGuichet.MontantEspèces + G_recGuichet.MontantRendu
    chkCoupureX = G_recGuichet.chkCoupureEspèces
Else
    curMontant = G_recGuichet.Montant
    chkCoupureX = G_recGuichet.chkCoupureChange
End If

Select Case chkCoupureX
    Case Is = "M": fraCoupure.Enabled = True: Coupure_Display
    Case Is = "A": fraCoupure.Enabled = False
                    For I = 0 To 21: txtCoupureNb(I) = "": Next I
                    I = 0
                    Do While curMontant > 0
                        If curMontant < CoupureNominal(optCoupure, I) Or CoupureNominal(optCoupure, I) = 0 Then
                            I = I + 1
                            If I > 21 Then
                                Call lstErr_AddItem(lstErr, txtDevise1Montant, "? montant non ventilable"): Exit Do
                            End If
                        Else
                            dblNb = Fix(CDbl(curMontant / CoupureNominal(optCoupure, I)))
                            If dblNb > 9999 Then Call lstErr_AddItem(lstErr, cmdContext, "? nombre coupure > 9999"): Exit Do
                            txtCoupureNb(I) = CInt(dblNb)
                            curMontant = curMontant - CoupureNominal(optCoupure, I) * txtCoupureNb(I)
                        End If
                    Loop
    Case Else: fraCoupure.Enabled = False
                    For I = 0 To 21: txtCoupureNb(I) = "": Next I
End Select

blnCoupureAuto = False
Coupure_Control
End Sub

Public Sub libCoupure_Text(curMontant As Currency)
libCoupure = "Coupures : " & Trim(Format$(curMontant, "##### ### ### ##0.00"))
End Sub

Public Function txtChèqueNo_Control() As Boolean
txtChèqueNo_Control = False
If Trim(txtChèqueNo) = "" Then Call lstErr_AddItem(lstErr, txtChèqueNo, "? N° chèque"): Exit Function
strOppChq_Numéro = Format(Val(txtChèqueNo), "0000000")
G_recGuichet.NoChèque = strOppChq_Numéro
For I = 1 To G_arrOppChq_Numéro_Nb
    If G_arrOppChq_Numéro(I) = strOppChq_Numéro Then
        Call lstErr_AddItem(lstErr, txtChèqueNo, "? opposition sur chèque")
        Exit Function
    End If
Next I
txtChèqueNo_Control = True

End Function


Public Sub Coupure_Ajustement(curX As Currency, curCoupureMini As Currency, optCoupureAjustement)
Dim curNew As Currency
If curCoupureMini <> 0 Then
    curNew = Fix(curX / curCoupureMini + 0.1) * curCoupureMini
    If curNew <> curX Then
        Select Case optCoupureAjustement
            Case "+": curX = curNew + curCoupureMini
            Case Else: curX = curNew
        End Select
    End If
End If
End Sub

Public Sub cmdSave_Db()
blnControl = False
cmdOk.Visible = False
If IsNull(srvGuichet_Update(G_recGuichet)) Then
    lastActiveControl_Name = ""
    
    If G_recGuichet.optAvis = "1" Then
        prtguichetX "1", G_recGuichet, G_CV1, G_CV2
    Else
        If G_recGuichet.optAvis = "2" Then
            prtguichetX "1", G_recGuichet, G_CV1, G_CV2
            prtguichetX "2", G_recGuichet, G_CV1, G_CV2
        End If
    End If
    'recGuichet_Display G_recGuichet, frmGuichet.picCpt
    cmdQuit_Click
End If

End Sub


Public Sub Compte_Load()
recCompteInit ClientCompte
ClientCompte.Société = G_recGuichet.Société
ClientCompte.Agence = G_recGuichet.Agence
ClientCompte.Devise = G_recGuichet.Devise
ClientCompte.Numéro = G_recGuichet.Compte
ClientCompte.BiaTyp = "000"
ClientCompte.BiaNum = "00"
ClientCompte.Method = "SeekL1"
If Not IsNull(srvCompteMon(ClientCompte)) Then Call lstErr_AddItem(lstErr, lstErr, "? compte en " & ClientCompte.Devise): Exit Sub

End Sub

