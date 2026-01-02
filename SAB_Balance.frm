VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAB_Balance 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_Balance"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "SAB_Balance.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.ListBox lstW 
      Height          =   255
      Left            =   7560
      Sorted          =   -1  'True
      TabIndex        =   126
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstService 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   375
      TabIndex        =   114
      Top             =   8040
      Width           =   2505
   End
   Begin VB.Frame fraSelect_Options 
      Height          =   5085
      Left            =   3405
      TabIndex        =   86
      Top             =   3795
      Visible         =   0   'False
      Width           =   13275
      Begin VB.CheckBox chkSelect_COMPTECLO 
         BackColor       =   &H80000004&
         Caption         =   "Date clôture >="
         Height          =   285
         Left            =   4440
         TabIndex        =   130
         Top             =   4080
         Width           =   1545
      End
      Begin VB.ComboBox cboSelect_CLIENACAT 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   128
         Text            =   "CAT"
         Top             =   4560
         Width           =   3975
      End
      Begin VB.CheckBox chkSelect_COMPTEOUV 
         BackColor       =   &H80000004&
         Caption         =   "Date création >="
         Height          =   285
         Left            =   4440
         TabIndex        =   113
         Top             =   3600
         Width           =   1545
      End
      Begin VB.CheckBox chkSelect_Annulé 
         BackColor       =   &H80000004&
         Caption         =   "exclure comptes annulés"
         Height          =   285
         Left            =   4440
         TabIndex        =   111
         Top             =   2880
         Width           =   2235
      End
      Begin VB.CheckBox chkSelect_HB 
         BackColor       =   &H80000004&
         Caption         =   "exclure classe 9"
         Height          =   285
         Left            =   4440
         TabIndex        =   110
         Top             =   2280
         Width           =   2000
      End
      Begin VB.ComboBox cboPCEC 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   103
         Text            =   "PCEC"
         Top             =   2160
         Width           =   1300
      End
      Begin VB.ComboBox cboDevise 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   102
         Text            =   "Devise"
         Top             =   960
         Width           =   1300
      End
      Begin VB.ComboBox cboPLANCOPRO 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   101
         Text            =   "prod"
         Top             =   1560
         Width           =   1300
      End
      Begin VB.CheckBox chkSelect 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF80FF&
         Caption         =   "Tous les comptes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   100
         Top             =   360
         Width           =   2715
      End
      Begin VB.CheckBox chkSelect_Résidence 
         Caption         =   "tri RNR(CHA/PRO)"
         Height          =   300
         Left            =   8400
         TabIndex        =   99
         Top             =   2880
         Width           =   2000
      End
      Begin VB.CheckBox chkSelect_SoldeZ 
         BackColor       =   &H80000004&
         Caption         =   "exclure Solde=0"
         Height          =   285
         Left            =   4440
         TabIndex        =   98
         Top             =   480
         Value           =   1  'Checked
         Width           =   2000
      End
      Begin VB.CheckBox chkSelect_SoldeDb 
         BackColor       =   &H80000004&
         Caption         =   "exclure Débiteur"
         Height          =   255
         Left            =   4440
         TabIndex        =   97
         Top             =   1080
         Width           =   2000
      End
      Begin VB.CheckBox chkSelect_SoldeCr 
         BackColor       =   &H80000004&
         Caption         =   "exclure Créditeur"
         Height          =   285
         Left            =   4440
         TabIndex        =   96
         Top             =   1680
         Width           =   2000
      End
      Begin VB.Frame fraSelect_SOLDEDMO 
         Height          =   1335
         Left            =   11040
         TabIndex        =   92
         Top             =   480
         Width           =   2085
         Begin VB.CheckBox chkSelect_DORCPTDMV 
            BackColor       =   &H80000004&
            Caption         =   "hors Agios, Frais ...."
            Height          =   285
            Left            =   120
            TabIndex        =   127
            Top             =   960
            Width           =   1740
         End
         Begin VB.OptionButton optSelect_SOLDEDMO_Inf 
            Caption         =   "<="
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   94
            Top             =   120
            Width           =   645
         End
         Begin VB.OptionButton optSelect_SOLDEDMO_Sup 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   93
            Top             =   480
            Width           =   630
         End
         Begin MSComCtl2.DTPicker txtSelect_SOLDEDMO 
            Height          =   300
            Left            =   750
            TabIndex        =   95
            Top             =   240
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
            Format          =   97910787
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
      End
      Begin VB.CheckBox chkSelect_SOLDEDMO 
         BackColor       =   &H80000004&
         Caption         =   "date dernier mouvement"
         Height          =   285
         Left            =   8400
         TabIndex        =   91
         Top             =   1080
         Width           =   2025
      End
      Begin VB.CheckBox chkSelect_MOUVEMDCO 
         Caption         =   "select DCO"
         Height          =   345
         Left            =   8400
         TabIndex        =   90
         Top             =   3600
         Width           =   1110
      End
      Begin VB.TextBox txtSelect_CLIENARES 
         Height          =   285
         Left            =   1560
         TabIndex        =   89
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtSelect_CLIENARSD 
         Height          =   285
         Left            =   1560
         TabIndex        =   88
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox txtSelect_COMPTECLA 
         Height          =   300
         Left            =   1560
         TabIndex        =   87
         Top             =   4080
         Width           =   615
      End
      Begin MSComCtl2.DTPicker txtSelect_COMPTEOUV 
         Height          =   300
         Left            =   6000
         TabIndex        =   112
         Top             =   3600
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
         Format          =   97910787
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   -328351
      End
      Begin MSComCtl2.DTPicker txtSelect_COMPTECLO 
         Height          =   300
         Left            =   6000
         TabIndex        =   131
         Top             =   4080
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
         Format          =   97910787
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   -328351
      End
      Begin VB.Label lblSelect_CLIENACAT 
         Caption         =   "Catégorie client"
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   4560
         Width           =   1320
      End
      Begin VB.Label lblSelect_Devise 
         Caption         =   "Devise"
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   960
         Width           =   600
      End
      Begin VB.Label lblSelect_PLANCOPRO 
         Caption         =   "Produit"
         Height          =   270
         Left            =   120
         TabIndex        =   108
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label lblSelect_PCEC 
         Caption         =   "PCEC"
         Height          =   210
         Left            =   120
         TabIndex        =   107
         Top             =   2280
         Width           =   540
      End
      Begin VB.Label lblSelect_CLIENARES 
         Caption         =   "Responsable"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblSelect_CLIENARSD 
         Caption         =   "Pays Résidence"
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label lblSelect_COMPTECLA 
         Caption         =   "Classe S"
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   4080
         Width           =   615
      End
   End
   Begin VB.Frame fraBalance 
      BackColor       =   &H00F0FFFF&
      Caption         =   "Impression BALANCE "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   7200
      TabIndex        =   26
      Top             =   2640
      Width           =   6255
      Begin VB.CheckBox chkBalance_Récap_Bilan 
         BackColor       =   &H00F0FFFF&
         Caption         =   "Imprimer  Recap Bilan / HB"
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
         Left            =   3000
         TabIndex        =   118
         Top             =   760
         Width           =   3135
      End
      Begin VB.CheckBox chkBalance_Pays 
         BackColor       =   &H00F0FFFF&
         Caption         =   "Tri Pays / PCI / Compte / Dev"
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
         Left            =   3000
         TabIndex        =   117
         Top             =   1680
         Width           =   3255
      End
      Begin VB.CheckBox chkBalance_Compte_Soldé 
         BackColor       =   &H00F0FFFF&
         Caption         =   "Ignorer les comptes soldés"
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
         Left            =   3000
         TabIndex        =   116
         Top             =   1360
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.TextBox txtBalance_CSV_FileName 
         Height          =   285
         Left            =   3240
         TabIndex        =   85
         Top             =   2160
         Width           =   2895
      End
      Begin VB.CheckBox chkBalance_CSV 
         BackColor       =   &H00F0FFFF&
         Caption         =   "Exporter vers un fichier CSV"
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
         Left            =   3000
         TabIndex        =   84
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtBalance_CSV_Folder 
         Height          =   285
         Left            =   3240
         TabIndex        =   83
         Top             =   1920
         Width           =   2895
      End
      Begin VB.CheckBox chkBalance_Récap 
         BackColor       =   &H00F0FFFF&
         Caption         =   "Imprimer  Balance récapitulative"
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
         Left            =   3000
         TabIndex        =   82
         Top             =   460
         Width           =   3135
      End
      Begin VB.Frame fraBalance_Print 
         BackColor       =   &H00F0FFFF&
         Height          =   400
         Index           =   7
         Left            =   240
         TabIndex        =   69
         Top             =   2400
         Width           =   5775
         Begin VB.CheckBox chkBalance_Print 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Rupture détail"
            Height          =   255
            Index           =   7
            Left            =   20
            TabIndex        =   81
            Top             =   120
            Width           =   1695
         End
         Begin VB.CheckBox chkBalance_Print_FontBold 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Gras"
            Height          =   255
            Index           =   7
            Left            =   1920
            TabIndex        =   72
            Top             =   120
            Width           =   735
         End
         Begin VB.CheckBox chkBalance_Print_Line 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Souligné"
            Height          =   255
            Index           =   7
            Left            =   3000
            TabIndex        =   71
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtBalance_Print_Trame 
            Height          =   285
            Index           =   7
            Left            =   5100
            TabIndex        =   70
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblBalance_Print_Trame 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Trame"
            Height          =   255
            Index           =   7
            Left            =   4320
            TabIndex        =   73
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame fraBalance_Print 
         BackColor       =   &H00F0FFFF&
         Height          =   400
         Index           =   6
         Left            =   240
         TabIndex        =   64
         Top             =   5160
         Width           =   5775
         Begin VB.CheckBox chkBalance_Print 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "PCEC niveau 6"
            Height          =   255
            Index           =   6
            Left            =   20
            TabIndex        =   80
            Top             =   120
            Width           =   1695
         End
         Begin VB.CheckBox chkBalance_Print_FontBold 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Gras"
            Height          =   255
            Index           =   6
            Left            =   1920
            TabIndex        =   67
            Top             =   120
            Width           =   735
         End
         Begin VB.CheckBox chkBalance_Print_Line 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Souligné"
            Height          =   255
            Index           =   6
            Left            =   3000
            TabIndex        =   66
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtBalance_Print_Trame 
            Height          =   285
            Index           =   6
            Left            =   5100
            TabIndex        =   65
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblBalance_Print_Trame 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Trame"
            Height          =   255
            Index           =   6
            Left            =   4320
            TabIndex        =   68
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame fraBalance_Print 
         BackColor       =   &H00F0FFFF&
         Height          =   400
         Index           =   5
         Left            =   240
         TabIndex        =   59
         Top             =   4800
         Width           =   5775
         Begin VB.CheckBox chkBalance_Print 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "PCEC niveau 5"
            Height          =   255
            Index           =   5
            Left            =   20
            TabIndex        =   79
            Top             =   120
            Width           =   1695
         End
         Begin VB.CheckBox chkBalance_Print_FontBold 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Gras"
            Height          =   255
            Index           =   5
            Left            =   1920
            TabIndex        =   62
            Top             =   120
            Width           =   735
         End
         Begin VB.CheckBox chkBalance_Print_Line 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Souligné"
            Height          =   255
            Index           =   5
            Left            =   3000
            TabIndex        =   61
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtBalance_Print_Trame 
            Height          =   285
            Index           =   5
            Left            =   5100
            TabIndex        =   60
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblBalance_Print_Trame 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Trame"
            Height          =   255
            Index           =   5
            Left            =   4320
            TabIndex        =   63
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame fraBalance_Print 
         BackColor       =   &H00F0FFFF&
         Height          =   400
         Index           =   4
         Left            =   240
         TabIndex        =   54
         Top             =   4440
         Width           =   5775
         Begin VB.CheckBox chkBalance_Print 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "PCEC niveau 4"
            Height          =   255
            Index           =   4
            Left            =   20
            TabIndex        =   78
            Top             =   120
            Width           =   1695
         End
         Begin VB.CheckBox chkBalance_Print_FontBold 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Gras"
            Height          =   255
            Index           =   4
            Left            =   1920
            TabIndex        =   57
            Top             =   120
            Width           =   735
         End
         Begin VB.CheckBox chkBalance_Print_Line 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Souligné"
            Height          =   255
            Index           =   4
            Left            =   3000
            TabIndex        =   56
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtBalance_Print_Trame 
            Height          =   285
            Index           =   4
            Left            =   5100
            TabIndex        =   55
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblBalance_Print_Trame 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Trame"
            Height          =   255
            Index           =   4
            Left            =   4320
            TabIndex        =   58
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame fraBalance_Print 
         BackColor       =   &H00F0FFFF&
         Height          =   400
         Index           =   3
         Left            =   240
         TabIndex        =   49
         Top             =   4080
         Width           =   5775
         Begin VB.CheckBox chkBalance_Print 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "PCEC niveau 3"
            Height          =   255
            Index           =   3
            Left            =   20
            TabIndex        =   77
            Top             =   120
            Width           =   1695
         End
         Begin VB.CheckBox chkBalance_Print_FontBold 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Gras"
            Height          =   255
            Index           =   3
            Left            =   1920
            TabIndex        =   52
            Top             =   120
            Width           =   735
         End
         Begin VB.CheckBox chkBalance_Print_Line 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Souligné"
            Height          =   255
            Index           =   3
            Left            =   3000
            TabIndex        =   51
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtBalance_Print_Trame 
            Height          =   285
            Index           =   3
            Left            =   5100
            TabIndex        =   50
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblBalance_Print_Trame 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Trame"
            Height          =   255
            Index           =   3
            Left            =   4320
            TabIndex        =   53
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame fraBalance_Print 
         BackColor       =   &H00F0FFFF&
         Height          =   400
         Index           =   2
         Left            =   240
         TabIndex        =   44
         Top             =   3720
         Width           =   5775
         Begin VB.CheckBox chkBalance_Print 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "PCEC niveau 2"
            Height          =   255
            Index           =   2
            Left            =   20
            TabIndex        =   76
            Top             =   120
            Width           =   1695
         End
         Begin VB.CheckBox chkBalance_Print_FontBold 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Gras"
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   47
            Top             =   120
            Width           =   735
         End
         Begin VB.CheckBox chkBalance_Print_Line 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Souligné"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   46
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtBalance_Print_Trame 
            Height          =   285
            Index           =   2
            Left            =   5100
            TabIndex        =   45
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblBalance_Print_Trame 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Trame"
            Height          =   255
            Index           =   2
            Left            =   4320
            TabIndex        =   48
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame fraBalance_Print 
         BackColor       =   &H00F0FFFF&
         Height          =   400
         Index           =   1
         Left            =   240
         TabIndex        =   39
         Top             =   3360
         Width           =   5775
         Begin VB.CheckBox chkBalance_Print 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "PCEC niveau 1"
            Height          =   255
            Index           =   1
            Left            =   20
            TabIndex        =   75
            Top             =   120
            Width           =   1695
         End
         Begin VB.CheckBox chkBalance_Print_FontBold 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Gras"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   42
            Top             =   120
            Width           =   735
         End
         Begin VB.CheckBox chkBalance_Print_Line 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Souligné"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   41
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox txtBalance_Print_Trame 
            Height          =   285
            Index           =   1
            Left            =   5100
            TabIndex        =   40
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblBalance_Print_Trame 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Trame"
            Height          =   255
            Index           =   1
            Left            =   4320
            TabIndex        =   43
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.CheckBox chkBalance_Détail 
         BackColor       =   &H00F0FFFF&
         Caption         =   "Imprimer la Balance détaillée"
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
         Left            =   3000
         TabIndex        =   34
         Top             =   160
         Width           =   2775
      End
      Begin VB.Frame fraBalance_Print 
         BackColor       =   &H00F0FFFF&
         Height          =   400
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   3000
         Width           =   5775
         Begin VB.CheckBox chkBalance_Print 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "PCEC niveau 0"
            Height          =   255
            Index           =   0
            Left            =   20
            TabIndex        =   74
            Top             =   120
            Width           =   1695
         End
         Begin VB.TextBox txtBalance_Print_Trame 
            Height          =   285
            Index           =   0
            Left            =   5100
            TabIndex        =   38
            Top             =   120
            Width           =   615
         End
         Begin VB.CheckBox chkBalance_Print_Line 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Souligné"
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   36
            Top             =   120
            Width           =   1095
         End
         Begin VB.CheckBox chkBalance_Print_FontBold 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Gras"
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   35
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblBalance_Print_Trame 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Trame"
            Height          =   255
            Index           =   0
            Left            =   4320
            TabIndex        =   37
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame fraBalance_YSOLDE0 
         BackColor       =   &H00F0FFFF&
         Caption         =   "en date du "
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
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   2655
         Begin VB.OptionButton optBalance_YSOLDE0_MP2 
            BackColor       =   &H00F0FFFF&
            Caption         =   "MP2"
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
            Left            =   240
            TabIndex        =   115
            Top             =   720
            Width           =   2115
         End
         Begin VB.OptionButton optBalance_YSOLDE0_AP1 
            BackColor       =   &H00F0FFFF&
            Caption         =   "AP1"
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
            Left            =   240
            TabIndex        =   32
            Top             =   960
            Width           =   2115
         End
         Begin VB.OptionButton optBalance_YSOLDE0_MP1 
            BackColor       =   &H00F0FFFF&
            Caption         =   "MP1"
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
            Left            =   240
            TabIndex        =   31
            Top             =   480
            Width           =   2115
         End
         Begin VB.OptionButton optBalance_YSOLDE0_J 
            BackColor       =   &H00F0FFFF&
            Caption         =   "J"
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
            Left            =   240
            TabIndex        =   30
            Top             =   240
            Width           =   2000
         End
      End
      Begin VB.CommandButton cmdBalance_Quit 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Abandonner"
         Height          =   600
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdBalance_Ok 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Imprimer"
         Height          =   600
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1680
         Width           =   1215
      End
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   8640
      TabIndex        =   6
      Top             =   0
      Width           =   4560
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8970
      Left            =   -30
      TabIndex        =   4
      Top             =   465
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   15822
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Comptes"
      TabPicture(0)   =   "SAB_Balance.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mouvements"
      TabPicture(1)   =   "SAB_Balance.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraMvt"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Autorisations"
      TabPicture(2)   =   "SAB_Balance.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picYAUTE1I0"
      Tab(2).Control(1)=   "fraYAUTE1I0"
      Tab(2).Control(2)=   "fgYAUTE1I0"
      Tab(2).Control(3)=   "lblYAUTE1I0"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Stock"
      TabPicture(3)   =   "SAB_Balance.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraYBIASTO0"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraYBIASTO0 
         Height          =   8385
         Left            =   -74880
         TabIndex        =   119
         Top             =   480
         Width           =   13530
         Begin MSFlexGridLib.MSFlexGrid fgYBIASTO0 
            Height          =   7365
            Left            =   2520
            TabIndex        =   120
            Top             =   840
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   12991
            _Version        =   393216
            Rows            =   1
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   200
            BackColor       =   14737632
            ForeColor       =   8388608
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"SAB_Balance.frx":04B2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label libYBIASTO0_Solde 
            Caption         =   "Solde"
            Height          =   255
            Left            =   2640
            TabIndex        =   125
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label libYBIASTO0_Total 
            Caption         =   "Total"
            Height          =   255
            Left            =   2640
            TabIndex        =   124
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label libYBIASTO0_Diff 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   8640
            TabIndex        =   123
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label libYBIASTO0_SOLDECEN 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   5160
            TabIndex        =   122
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label libYBIASTO0_YSTOMON 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   5160
            TabIndex        =   121
            Top             =   120
            Width           =   2655
         End
      End
      Begin VB.PictureBox picYAUTE1I0 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00F0FFFF&
         Height          =   7335
         Left            =   -66480
         ScaleHeight     =   7275
         ScaleWidth      =   4995
         TabIndex        =   25
         Top             =   1680
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Frame fraYAUTE1I0 
         Height          =   855
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   3285
         Begin VB.OptionButton optYAUTE1i0_NIV_X 
            Alignment       =   1  'Right Justify
            Caption         =   "NIveau  *"
            Height          =   195
            Left            =   2040
            TabIndex        =   24
            Top             =   600
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optYAUTE1i0_NIV_2 
            Alignment       =   1  'Right Justify
            Caption         =   "NIveau 2"
            Height          =   195
            Left            =   2040
            TabIndex        =   23
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optYAUTE1i0_NIV_1 
            Alignment       =   1  'Right Justify
            Caption         =   "NIveau 1"
            Height          =   195
            Left            =   2040
            TabIndex        =   22
            Top             =   120
            Width           =   1215
         End
         Begin VB.ListBox lstYAUTE1I0 
            Height          =   450
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Frame fraMvt 
         Height          =   8505
         Left            =   -74940
         TabIndex        =   8
         Top             =   390
         Width           =   13635
         Begin MSFlexGridLib.MSFlexGrid fgYBIAMVT0 
            Height          =   8325
            Left            =   45
            TabIndex        =   9
            Top             =   135
            Width           =   13560
            _ExtentX        =   23918
            _ExtentY        =   14684
            _Version        =   393216
            Rows            =   1
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   14737632
            ForeColor       =   8388608
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"SAB_Balance.frx":0552
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame fraTab0 
         Height          =   8550
         Left            =   135
         TabIndex        =   5
         Top             =   315
         Width           =   13590
         Begin VB.Frame fraList 
            Height          =   6900
            Left            =   615
            TabIndex        =   132
            Top             =   1470
            Width           =   7155
            Begin MSFlexGridLib.MSFlexGrid fgList 
               Height          =   6090
               Left            =   930
               TabIndex        =   133
               Top             =   330
               Width           =   7155
               _ExtentX        =   12621
               _ExtentY        =   10742
               _Version        =   393216
               Cols            =   5
               FixedCols       =   0
               BackColor       =   14745599
               BackColorFixed  =   12582912
               ForeColorFixed  =   -2147483633
               BackColorBkg    =   14745599
               WordWrap        =   -1  'True
               FormatString    =   "<Code|<Identifiant         |<Fisc,CO  |<Adresse                                                                    |"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.CommandButton cmdList_Ok 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Imprimer RIB"
               Height          =   600
               Left            =   4725
               Style           =   1  'Graphical
               TabIndex        =   135
               Top             =   6225
               Width           =   1215
            End
            Begin VB.CommandButton cmdList_Quit 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Abandonner"
               Height          =   600
               Left            =   585
               Style           =   1  'Graphical
               TabIndex        =   134
               Top             =   6255
               Width           =   1215
            End
         End
         Begin VB.Frame fraSelect 
            Height          =   990
            Left            =   120
            TabIndex        =   11
            Top             =   165
            Width           =   9180
            Begin VB.CommandButton cmdOptions 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Options de sélection"
               Height          =   600
               Left            =   4320
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   240
               Width           =   2220
            End
            Begin VB.CommandButton cmdSelect_MesComptes 
               BackColor       =   &H00C0FFFF&
               Caption         =   "MesComptes"
               Height          =   600
               Left            =   3000
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   240
               Width           =   1155
            End
            Begin VB.CommandButton cmdSelect 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Exécuter la requête"
               Height          =   600
               Left            =   6720
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   240
               Width           =   2175
            End
            Begin VB.TextBox txtCompte 
               Height          =   315
               Left            =   1170
               TabIndex        =   13
               Top             =   600
               Width           =   1305
            End
            Begin VB.TextBox txtIntitulé 
               Height          =   330
               Left            =   1170
               TabIndex        =   12
               Top             =   195
               Width           =   1275
            End
            Begin VB.Label lblSelect_Compte 
               Caption         =   "%compte%"
               Height          =   195
               Left            =   120
               TabIndex        =   18
               Top             =   600
               Width           =   825
            End
            Begin VB.Label lblSelect_Intitulé 
               Caption         =   "Intitulé"
               Height          =   270
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   660
            End
         End
         Begin VB.Frame fraSelect_Mvt 
            Caption         =   "Mouvements du  au"
            Height          =   930
            Left            =   9360
            TabIndex        =   10
            Top             =   165
            Width           =   3375
            Begin MSComCtl2.DTPicker txtAmjMin 
               Height          =   300
               Left            =   240
               TabIndex        =   0
               Top             =   360
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
               Format          =   97910787
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtAmjMax 
               Height          =   300
               Left            =   1800
               TabIndex        =   1
               Top             =   360
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
               Format          =   97910787
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7320
            Left            =   45
            TabIndex        =   7
            Top             =   1185
            Width           =   13380
            _ExtentX        =   23601
            _ExtentY        =   12912
            _Version        =   393216
            Rows            =   1
            Cols            =   12
            FixedCols       =   0
            RowHeightMin    =   380
            BackColor       =   16448250
            ForeColor       =   8388608
            BackColorFixed  =   8421376
            ForeColorFixed  =   16777215
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"SAB_Balance.frx":0623
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgYAUTE1I0 
         Height          =   7485
         Left            =   -74805
         TabIndex        =   19
         Top             =   1185
         Width           =   13560
         _ExtentX        =   23918
         _ExtentY        =   13203
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         FixedCols       =   0
         RowHeightMin    =   50
         BackColor       =   16777215
         ForeColor       =   4210752
         BackColorFixed  =   8421376
         ForeColorFixed  =   16777215
         BackColorSel    =   12648384
         BackColorBkg    =   16777215
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   3
         FormatString    =   $"SAB_Balance.frx":06C2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblYAUTE1I0 
         Caption         =   "Position COMPTABLE au "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -71550
         TabIndex        =   136
         Top             =   615
         Width           =   10125
      End
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "SAB_Balance.frx":079B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   -30
      Width           =   615
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuAuto_FOTC 
         Caption         =   "Auto : Relevé FOTC"
      End
      Begin VB.Menu mnuAuto_FOTC_CHAPRO 
         Caption         =   "Auto : Etat FOTC : CHA / PRO"
      End
      Begin VB.Menu mnuAuto_SOBI 
         Caption         =   "Auto : Relevé SOBI"
      End
      Begin VB.Menu mnuAuto_Compta_TVA 
         Caption         =   "Auto : Compta TVA (mvt classe 7)"
      End
      Begin VB.Menu mnuAuto_Client_Stat 
         Caption         =   "Auto : stat catégorie client"
      End
      Begin VB.Menu mnuZCOMREF0_Service_Export 
         Caption         =   "Excel : code produit => Compte + Service"
      End
      Begin VB.Menu mnuAuto_Balance_Stock 
         Caption         =   "Auto : Balance par service  + Stock opé"
      End
      Begin VB.Menu mnuAuto_Balance_Service 
         Caption         =   "Auto : Balance par service  31.12.200X"
      End
      Begin VB.Menu mnuContextX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAuto_Groupe_6000 
         Caption         =   "Auto : Relevé groupe 6000 (Libye)"
      End
      Begin VB.Menu mnuSelect_Print_Liste_Xls 
         Caption         =   "Exportation des encours (groupe de racines)"
      End
      Begin VB.Menu mnuSelect_ENG_BEA 
         Caption         =   "Etat des engagements BEA"
      End
      Begin VB.Menu mnuSelect_ENG_LFB 
         Caption         =   "Etat des engagements LFB"
      End
      Begin VB.Menu mnuSelect_ENG_Detail_BEA 
         Caption         =   "Etat détaillé des engagements BEA"
      End
      Begin VB.Menu mnuSelect_ENG_Detail_LFB 
         Caption         =   "Etat détaillé des engagements LFB"
      End
      Begin VB.Menu mnuSelect_CPT_OD 
         Caption         =   "Liste des OD comptabilisées (période)"
      End
      Begin VB.Menu mnuContextX1b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAuto_Autorisation_Echeance 
         Caption         =   "Autorisation  -1 / +3  mois"
      End
      Begin VB.Menu mnuAuto_Autorisation_Dépassement 
         Caption         =   "Autorisation_Dépassement"
      End
      Begin VB.Menu mnuContextX1c 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Print_Liste 
         Caption         =   "Imprimer la liste"
      End
      Begin VB.Menu mnuSelect_Print_Liste_T 
         Caption         =   "Imprimer la liste + Total Racine"
      End
      Begin VB.Menu mnuSelect_Print_Relevé 
         Caption         =   "Imprimer les Relevés"
      End
      Begin VB.Menu mnuSelect_Print_Cumul 
         Caption         =   "Imprimer cumul des mouvements"
      End
      Begin VB.Menu mnuPrint0_X1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect_Print_Balance 
         Caption         =   "Imprimer la  balance"
      End
      Begin VB.Menu mnuSelect_Print_Balance_Stock 
         Caption         =   "Imprimer Balance  + Stock opérations"
      End
      Begin VB.Menu mnuPrint0_X2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect_Print_Client_Stat 
         Caption         =   "Imprimer stat catégorie client"
      End
      Begin VB.Menu mnuSelect_Print_PCI_DC 
         Caption         =   "Imprimer les anomalies de sens des comptes"
      End
      Begin VB.Menu mnuPrint0_X3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect_Print_fgSelect 
         Caption         =   "Export affichage => excel"
      End
      Begin VB.Menu mnuSelect_Print_fgSelect_Mail 
         Caption         =   "Export affichage => mail"
      End
      Begin VB.Menu mnuSelect_Print_Compte 
         Caption         =   "Export YBIACPT0.csv"
      End
      Begin VB.Menu mnuSelect_Print_Adresse 
         Caption         =   "Export  Adresse.csv"
      End
   End
   Begin VB.Menu mnuPrint1 
      Caption         =   "mnuPrint1"
      Visible         =   0   'False
      Begin VB.Menu mnuRelevé_Print 
         Caption         =   "Imprimer Relevé"
      End
      Begin VB.Menu mnuRIB_Print 
         Caption         =   "Imprimer RIB"
      End
      Begin VB.Menu mnuRelevéRIB_Print 
         Caption         =   "Imprimer Relevé +RIB"
      End
   End
   Begin VB.Menu mnuPrint2 
      Caption         =   "mnuPrint2"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint2_Excel 
         Caption         =   "export EXCEL"
      End
      Begin VB.Menu mnuPrint2_Mail 
         Caption         =   "envoi Mail"
      End
   End
   Begin VB.Menu mnufgSelect 
      Caption         =   "mnufgSelect"
      Visible         =   0   'False
      Begin VB.Menu mnufgSelect_fgYBIAMVT0 
         Caption         =   "afficher Mvts (date de Traitement)"
      End
      Begin VB.Menu mnufgSelect_fgYBIAMVT0_MOUVEMDVA 
         Caption         =   "afficher Mvts (date de Valeur)"
      End
      Begin VB.Menu mnufgSelect_fgYBIASTO0 
         Caption         =   "afficher Stock"
      End
      Begin VB.Menu mnufgSelect_fgYAUTE1I0 
         Caption         =   "afficher AUT"
      End
      Begin VB.Menu mnufgSelect_X1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufgSelect_Print_RIB 
         Caption         =   "Imprimer RIB"
      End
      Begin VB.Menu mnufgSelect_Print_Extrait 
         Caption         =   "Imprimer l'extrait"
      End
      Begin VB.Menu mnufgSelect_X2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufgSelect_KYC 
         Caption         =   "afficher Signatures"
      End
      Begin VB.Menu mnuZADRESS0 
         Caption         =   "Afficher les adresses"
      End
   End
End
Attribute VB_Name = "frmSAB_Balance"
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
Dim SAB_Balance_Aut As typeAuthorization
Dim blnAuto As Boolean, blnError As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim fgYBIAMVT0_FormatString As String, fgYBIAMVT0_K As Integer
Dim fgYBIAMVT0_RowDisplay As Integer, fgYBIAMVT0_RowClick As Integer, fgYBIAMVT0_ColClick As Integer
Dim fgYBIAMVT0_ColorClick As Long, fgYBIAMVT0_ColorDisplay As Long
Dim fgYBIAMVT0_Sort1 As Integer, fgYBIAMVT0_Sort2 As Integer
Dim fgYBIAMVT0_SortAD As Integer, fgYBIAMVT0_Sort1_Old As Integer
Dim fgYBIAMVT0_arrIndex As Integer
Dim blnfgYBIAMVT0_DisplayLine As Boolean


Dim fgYAUTE1I0_FormatString As String, fgYAUTE1I0_K As Integer
Dim fgYAUTE1I0_RowDisplay As Integer, fgYAUTE1I0_RowClick As Integer, fgYAUTE1I0_ColClick As Integer
Dim fgYAUTE1I0_ColorClick As Long, fgYAUTE1I0_ColorDisplay As Long
Dim fgYAUTE1I0_Sort1 As Integer, fgYAUTE1I0_Sort2 As Integer
Dim fgYAUTE1I0_SortAD As Integer, fgYAUTE1I0_Sort1_Old As Integer
Dim fgYAUTE1I0_arrIndex As Integer
Dim blnfgYAUTE1I0_DisplayLine As Boolean
Dim xYAUTE1I0 As typeYAUTE1I0
Dim blnfgYAUTE1I0_Display As Boolean, fgYAUTE1I0_RowSelect As Integer

Dim meYBIACPT0 As typeYBIACPT0, xYBIACPT0 As typeYBIACPT0

Dim Nb As Long

Dim mcboPCEC As String, mcboDevise As String, mcboPLANCOPRO As String
Dim mcboSelect_CLIENACAT As String

Dim marrYBIACPT0() As typeYBIACPT0, marrYBIACPT0_Nb As Long
Dim xYBIAMVT0 As typeYBIAMVT0

Dim xAmjMin As String, xAmjMax As String
Dim mAmjMin As String, mAmjmax As String
Dim mAmj_SOLDEDMO As Long, mAmj_COMPTEOUV As Long, mAmj_COMPTECLO As Long

Dim blnMesComptes As Boolean
Dim mYAUTE1I0_AUTE1INIV As Long

Dim blnSelect_Pays As Boolean



Dim fgYBIASTO0_FormatString As String, fgYBIASTO0_K As Integer
Dim fgYBIASTO0_RowDisplay As Integer, fgYBIASTO0_RowClick As Integer, fgYBIASTO0_ColClick As Integer
Dim fgYBIASTO0_ColorClick As Long, fgYBIASTO0_ColorDisplay As Long
Dim fgYBIASTO0_Sort1 As Integer, fgYBIASTO0_Sort2 As Integer
Dim fgYBIASTO0_SortAD As Integer, fgYBIASTO0_Sort1_Old As Integer
Dim fgYBIASTO0_arrIndex As Integer
Dim blnfgYBIASTO0_DisplayLine As Boolean
Dim meYBIASTO0 As typeYBIASTO0, xYBIASTO0 As typeYBIASTO0

Dim marrZCOMREF0() As typeZCOMREF0, marrZCOMREF0_Nb As Long
Dim arrService() As String, arrService_Nb As Long
Dim arrDevise() As String, arrDevise_Nb As Long
Dim wCOMREFCOR() As String

Dim arrYBIASTO0() As typeYBIASTO0, stockYBIACPT0() As typeYBIACPT0, stockCompte_Nb As Long
Dim wYSTOMON() As Currency, wDORCPTDMV() As Long
Dim blnService_Printer As Boolean
Dim blnBalance_Stock_détail As Boolean
Dim arrService_Balance_Cumul() As typeSAB_Balance_Cumul

Dim blnZDORCPT_SQL_ODBC As Boolean
Dim meUnit As typeUnit

Dim blnBalance_Service_Stock As Boolean
Dim curBalance() As Currency
Dim meZADRESS0 As typeZADRESS0, arrZADRESS0() As typeZADRESS0, arrZADRESS0_Nb As Integer, arrZADRESS0_K As Integer

Dim oldYBIACPT0 As typeYBIACPT0, xZPLAN0 As typeZPLAN0
Dim mXls1_Row As Long, mXls1_Col As Long, mXls2_Row As Long, mXls2_Col As Long, mXls2_Row_Cli As Long
Dim mXls2_row_T As Long, mXls1_Row_C As Long, mXls1_Row_T As Long

Dim mXls1_File As Integer
Dim mXls1_Cols As Integer
'Dim mXls2_Cols As Integer, mXls2_Row As Integer

Dim arrDev() As String, arrDev_RowT() As Long, arrDev_Cours() As Double, arrDev_Nb As Integer
Dim xCLIGRPREG As String
Dim sBilan_DB As String, colBilan_DB As Integer, sBilan_CR As String, colBilan_CR As Integer
Dim sHors_Bilan_DB As String, colHors_Bilan_DB As Integer, sHors_Bilan_CR As String, colHors_Bilan_CR As Integer
Dim alfBilan_DB As String, alfBilan_CR As String, alfHors_Bilan_DB As String, alfHors_Bilan_CR As String
Dim arrPays() As typePays, arrPays_NB As Integer

Dim blnGroupe_Filtre As Boolean, arrGroupe_Filtre() As String, arrGroupe_Filtre_Nb As Integer

Dim curBilan_Min As Currency, curHors_Bilan_Min As Currency
Dim blnGroupe_Nostro_Exclus As Boolean

Dim blnfrmSAB_Dossier_DB  As Boolean

Dim fgList_FormatString As String, fgList_K As Integer
Dim fgList_RowDisplay As Integer, fgList_RowClick As Integer, fgList_ColClick As Integer
Dim fgList_ColorClick As Long, fgList_ColorDisplay As Long
Dim fgList_Sort1 As Integer, fgList_Sort2 As Integer
Dim fgList_SortAD As Integer, fgList_Sort1_Old As Integer
Dim fgList_arrIndex As Integer
Dim blnfgList_DisplayLine As Boolean

Dim rsSabX As Recordset
Dim arrService_Code() As String, arrService_Lib() As String, arrService_Code_Nb As Integer

Dim mAUTE1ICLI As String, mAUTE1IDAF_Min8 As String, mAUTE1IDAF_Max8 As String
Dim mAUTE1IDAF_Min7 As Long, mAUTE1IDAF_Max7 As Long
Dim mfgYAUTE1I0_Fct As String
Dim mAUTE1ICLI_Aut As String, arrCellBackColor(3) As Long, arrCellBackColor_K As Integer

Dim mENG_LFB As String, mENG_LFB_6002 As String
Dim blnBalance_Stock_CLIENARES As Boolean
Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook
Dim wsExcel As Excel.Worksheet
Private Sub cmdBalance_Ok_Click_xlsManual()
Dim I As Integer, K As Integer, X As String
Dim wBalance_Ok_Param As String
Dim wsexcelRien As Excel.Worksheet

Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Balance : " & fgSelect.Rows - 1)
' Rechercher tous les comptes (même solde = 0) , n'imprimer que les comptes non soldés à la date
If chkSelect_SoldeZ = "1" Then fgSelect_Display_SoldeZ
wBalance_Ok_Param = cmdBalance_Ok_Param
Call cmdBalance_Ok_Print_xlsManual(wBalance_Ok_Param, " ", -1, wsexcelRien)

Me.Enabled = True: Me.MousePointer = 0

End Sub


Public Sub cmdBalance_Ok_Print_xlsManual(lBalance_Ok_Param As String, lMsg As String, ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim wIdFile As Integer, wFileName As String
Dim V
V = Null

If chkBalance_Pays Then
    Mid$(lBalance_Ok_Param, 11, 1) = "1"     ' balance par pays
    fgSelect_Sort1 = 9: fgSelect_Sort2 = 9 ' TRI  Pays / PCI/ Dev /Compte
Else
    fgSelect_Sort1 = 2: fgSelect_Sort2 = 2 ' TRI Dev / PCI / Compte
End If
fgSelect_SortAD = 6
fgSelect_SortX fgSelect_Sort1


If chkBalance_CSV = "1" Then
    If Mid$(txtBalance_CSV_Folder, Len(txtBalance_CSV_Folder), 1) <> "\" Then txtBalance_CSV_Folder = txtBalance_CSV_Folder & "\"
    wFileName = txtBalance_CSV_Folder & txtBalance_CSV_FileName
    wIdFile = 0
    Mid$(lBalance_Ok_Param, 6, 1) = chkBalance_CSV
    V = File_Export_Monitor("Output", wIdFile, wFileName)
    Mid$(lBalance_Ok_Param, 7, 3) = Format$(wIdFile, "000")
    Call File_Export_Monitor("Print", wIdFile, "COMPTEDEV;COMPTEOBL;COMPTECOM;COMPTEINT;DB_dev;CR_dev;DB_eur;CR_eur;COMPTEOUV;COMPTEFON;SOLDEDMO")

End If

If IsNull(V) Then
    Call prtSAB_Balance_Monitor_xlsManual(lBalance_Ok_Param, xAmjMin, xAmjMax, fgSelect, marrYBIACPT0(), marrYBIACPT0_Nb, lMsg, wYSTOMON(), wDORCPTDMV(), currentRow, wsExcel)
    If chkBalance_CSV = "1" Then V = File_Export_Monitor("Close", wIdFile, wFileName)
End If

fgSelect.Visible = True
fraBalance.Visible = False
Me.Show

End Sub

Public Sub cmdBalance_Ok_Stock_xlsManual()
Dim X As String, K As Long, lstW_Index As Long, I As Long, mService As String
Dim blnOk As Boolean, blnAdd As Boolean
Dim xWhere As String, nbDossier As Long, X20 As String
Dim wService_Name As String, wService_Référence As String, wService_Printer As String, wService_Sxx As String
Dim iDevise As Integer
Dim xSQL As String
Dim wBalance_Ok_Param As String
Dim curSolde As Currency
Dim V
Dim ii As Long
Dim currentSheet As Long
Dim currentRow As Long
Dim wbExcel As Excel.Workbook

'===========================================================================================
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : chargement des services/comptes")

ZCOMREF0_SQL_ODBC
'======================================== Compte sans service => Responsable client ===================
mService = ""
For I = 1 To marrYBIACPT0_Nb
    If Trim(marrYBIACPT0(I).CLIENARES) = "R80" Then
        X = marrYBIACPT0(I).PLANCOPRO
        If X = "CAV" Or X = "LIE" Or X = "LOR" Then wCOMREFCOR(I) = "G7"            '$JPL 2014-11-12
    End If
    If wCOMREFCOR(I) = "" Then
        X = marrYBIACPT0(I).PLANCOPRO
        If X = "CAV" Or X = "LIE" Or X = "LOR" Then
            wCOMREFCOR(I) = Trim(marrYBIACPT0(I).CLIENARES)
            If mService <> wCOMREFCOR(I) Then
                mService = wCOMREFCOR(I)
                blnAdd = True
                For K = 1 To arrService_Nb
                    If mService = arrService(K) Then blnAdd = False: Exit For
                Next K
                If blnAdd Then
                    arrService_Nb = arrService_Nb + 1
                    If arrService_Nb >= UBound(arrService) Then ReDim Preserve arrService(arrService_Nb + 10)
                    arrService(arrService_Nb) = mService
                End If
            End If
        End If
    End If
Next I
'=====================================================================================================
' - 2 : compte ne devant pas avoir de contrats associés
' - 1 : compte devant avoir des contrats , mais aucun contrat associé
' >= 0 : compte avec contrats , vérifier balance = stock

ReDim wYSTOMON(marrYBIACPT0_Nb)
For I = 1 To marrYBIACPT0_Nb
    wYSTOMON(I) = -2
Next I
'===========================================================================================
If Not blnBalance_Service_Stock Then
    ZSOLDE0_SQL_ODBC
Else
    Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : chargement OPENAT_PCI")
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID  = 'OPENAT_PCI'"
    Set rsSab = cnsab.Execute(xSQL)
    
    Do While Not rsSab.EOF
        X = Mid$(rsSab("BIATABTXT"), 4, 5)
        For I = 1 To marrYBIACPT0_Nb
            If X = Mid$(marrYBIACPT0(I).COMPTEOBL, 1, 5) Then
                 wYSTOMON(I) = -1
            End If
        Next I
        rsSab.MoveNext
    Loop
    '===========================================================================================
    Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : chargement du stock")
    xWhere = ""
    Call YBIASTO0_Sql(xWhere, nbDossier, arrYBIASTO0(), stockYBIACPT0(), stockCompte_Nb)
    
    Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : rapprochement")
    
    For K = 1 To stockCompte_Nb
        X20 = stockYBIACPT0(K).COMPTECOM
        For I = 1 To marrYBIACPT0_Nb
            If X20 = marrYBIACPT0(I).COMPTECOM Then
                If wYSTOMON(I) > -1 Then MsgBox marrYBIACPT0(I).COMPTECOM, vbExclamation, "stock déjà affecté  " & X20
                wYSTOMON(I) = arrYBIASTO0(K).YSTOMON
                Exit For
            End If
        Next I
    Next K
    '===========================================================================================
    ' recherche date du dernier mouvement
    Call ZDORCPT_SQL_ODBC
End If
'===========================================================================================
arrService_Balance_Cumul_Z
lstW.Clear
For K = 1 To arrService_Nb
    lstW.AddItem arrService(K)
Next K
'                                               '
Call init_xlsManual
'On recopie le classeur modèle de c:\BIASRV vers c:\temp\imp_pdf
FileCopy paramFolder_Local & "\Modeles\modele_BALANCE_Stock.xlsx", paramIMP_PDF_Path_Temp & "\modele_BALANCE_Stock.xlsx"
'on charge CE classeur dans Excel
Call appExcelPublic.Workbooks.Open(paramIMP_PDF_Path_Temp & "\modele_BALANCE_Stock.xlsx")
Set wbExcel = appExcelPublic.ActiveWorkbook
With wbExcel
    .Title = "BALANCE_Stock"
    .Subject = "BALANCE_Stock"
End With
'                                               '
For lstW_Index = 0 To lstW.ListCount - 1
    'on écrit systématiquement sur Feuil1 car BALANCE_Stock est notre feuille modèle
    wbExcel.Sheets.Add
    currentSheet = 2
    For ii = 1 To wbExcel.Sheets.Count
        If wbExcel.Sheets(ii).Name <> "BALANCE_Stock" Then
            currentSheet = ii
            Exit For
        End If
    Next ii
    'on recopie les 4 premières lignes de BALANCE_Stock vers Feuil2
    wbExcel.Sheets("BALANCE_Stock").Select
    Range("A1:L7").Select
    Selection.Copy
    wbExcel.Sheets(currentSheet).Select
    Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
    Range("A1").Select
    ActiveSheet.Paste
    Range("A8").Select
    currentRow = 7
    K = lstW_Index + 1
    lstW.ListIndex = lstW_Index
    mService = Trim(lstW.Text)
    Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : " & mService)
    If Mid$(mService, 1, 1) = "R" Then
        Call rsYBIATAB0_Read("RESPONSABLE", mService, "", wService_Name)
        wService_Name = Mid$(wService_Name, 34, 12)
        Select Case Mid$(mService, 1, 3)
            Case "R60":        wService_Printer = "DRH"
            Case "R80":         wService_Printer = "DER"
           Case Else: wService_Printer = "DCOM"
        End Select
        wService_Sxx = Table_Unit_SSI("", Trim(wService_Printer))
    Else
        V = rsElpTable_Read("SAb_Param", "Compte_Unit", mService, wService_Name, wService_Sxx)
        If Not IsNull(V) Then
            wService_Name = mService & " : ?????"
            wService_Printer = "CPT"
            wService_Sxx = "S60"
        Else
            wService_Printer = Table_Unit_SSI("S", Trim(wService_Sxx))
        End If
    End If
    wService_Référence = mService & " - " & Trim(wService_Name)
    arrService_Balance_Cumul(K, 0).Id = wService_Référence
    iDevise = 0
    fgSelect_Reset
    fgSelect.Rows = 1
    fgSelect.FormatString = fgSelect_FormatString
    fgSelect.Visible = False
    For I = 1 To marrYBIACPT0_Nb
       If mService = wCOMREFCOR(I) Then
           blnOk = True
            xYBIACPT0 = marrYBIACPT0(I)
            If blnBalance_Service_Stock Then
                curSolde = xYBIACPT0.SOLDECEN
            Else
                curSolde = curBalance(I)
            End If
            If curSolde = 0 Then blnOk = False
            If xYBIACPT0.COMPTEFON = "4" And curSolde = 0 Then blnOk = False
            If blnOk Then
                If fctUser_Classe_Aut(xYBIACPT0.COMPTECLA) Then
                    fgSelect_DisplayLine I
'===========================================================================================
                    If arrService_Balance_Cumul(K, iDevise).Dev <> xYBIACPT0.COMPTEDEV Then
                        For iDevise = 1 To arrDevise_Nb
                            If xYBIACPT0.COMPTEDEV = arrDevise(iDevise) Then
                                arrService_Balance_Cumul(K, iDevise).Id = wService_Référence
                                Exit For
                            End If
                        Next iDevise
                    End If
                    curX = Abs(curSolde)
                    If Mid$(xYBIACPT0.COMPTEOBL, 1, 1) <> "9" Then
                        arrService_Balance_Cumul(K, iDevise).Bilan_Nb = arrService_Balance_Cumul(K, iDevise).Bilan_Nb + 1
                        If curSolde > 0 Then
                            arrService_Balance_Cumul(K, iDevise).Bilan_DB = arrService_Balance_Cumul(K, iDevise).Bilan_DB + curX
                        Else
                            arrService_Balance_Cumul(K, iDevise).Bilan_CR = arrService_Balance_Cumul(K, iDevise).Bilan_CR + curX
                        End If
                    Else
                         arrService_Balance_Cumul(K, iDevise).HorsBilan_Nb = arrService_Balance_Cumul(K, iDevise).HorsBilan_Nb + 1
                        If curSolde > 0 Then
                            arrService_Balance_Cumul(K, iDevise).HorsBilan_DB = arrService_Balance_Cumul(K, iDevise).HorsBilan_DB + curX
                        Else
                            arrService_Balance_Cumul(K, iDevise).HorsBilan_CR = arrService_Balance_Cumul(K, iDevise).HorsBilan_CR + curX
                        End If
                   End If
'===========================================================================================
                End If
            End If
       End If
    Next I
    
    fgSelect.Visible = True
    Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : " & mService & " : " & fgSelect.Rows - 1)
    
    If fgSelect.Rows > 1 And blnBalance_Stock_détail Then
        If mService = "" Then mService = "G0"
        Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : impression " & mService)
        Call frmElpPrt.prtIMP_PDF_NoPaper_Init(wService_Sxx, "BIA-BAL-Stock_" & mService, "Archive")
        optBalance_YSOLDE0_J = True
        chkBalance_Détail = "1"
        chkBalance_Récap = "0"
        chkBalance_Récap_Bilan = "0"
        chkBalance_CSV = "0"
        chkBalance_Pays = "0"
        chkBalance_Compte_Soldé = "1"
        wBalance_Ok_Param = cmdBalance_Ok_Param
        Mid$(wBalance_Ok_Param, 1, 1) = "S"                      ' BALANCE avec contrôle STOCK
        Call cmdBalance_Ok_Print_xlsManual(wBalance_Ok_Param, wService_Référence, currentRow, wbExcel.Worksheets(currentSheet))
        If blnService_Printer Then Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("", "", "Balance / Stock - " & mService)

    End If
Next lstW_Index
Call wbExcel.Close(True)
Set wbExcel = Nothing
Kill paramIMP_PDF_Path_Temp & "\modele_BALANCE_Stock.xlsx"

End Sub

Private Sub cmdBalance_Service_xlsManual()

SSTab1.Tab = 0

blnBalance_Stock_détail = True
blnService_Printer = True

Call cmdBalance_Ok_Stock_xlsManual

Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S60", "BIA-BAL-Stock_Cumul", "Archive")

Call prtSAB_Balance_Cumul_Monitor_xlsManual(arrService_Balance_Cumul(), arrService_Nb, arrDevise_Nb)

Dim xPath As String, X1 As String, X2 As String, objFolder, objFiles, fsoFile As File
Dim wSendMail As typeSendMail

xPath = paramEditionNoPaper_Folder & "PDF\Archive_" & YBIATAB0_DATE_CPT_J
Set objFolder = msFileSystem.GetFolder(xPath)
Set objFiles = objFolder.Files
For Each fsoFile In objFiles
    If InStr(fsoFile.Name, "BIA-BAL-Stock_G3") > 0 Then
        If InStr(fsoFile.Name, "(S32)") > 0 Then
            X1 = xPath & "\" & fsoFile.Name
            X2 = Replace(Replace(X1, "S32", "S10"), "GDC", "SOBI")
            If Dir(X2) <> "" Then Kill X2
            msFileSystem.CopyFile X1, X2
            paramEditionNoPaper_Auto_Lnk = "<span style='font-size:9.0pt;font-family:Calibri'>""" _
                                     & "<A HREF=" & Asc34 & Replace(X2, paramEditionNoPaper_Folder & "PDF\", paramEditionNoPaper_Partage) & Asc34 & ">" _
                                    & "Cliquez ici pour afficher le document : " & X2 & "</A><BR><BR>"
    
            wSendMail.From = currentSSIWINMAIL
            wSendMail.FromDisplayName = "NoPaper BIA-BAL-Stock_G3"
            wSendMail.Recipient = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire("S10")
            wSendMail.CcRecipient = ""
            wSendMail.Subject = "BIA-BAL-Stock_G3 "
            wSendMail.Attachment = ""
            wSendMail.Message = mHtml_Head & "<span style='font-size:10.0pt;font-family:Calibri'>" _
                             & htmlFontColor_Black & "<BR><BR>" & paramEditionNoPaper_Auto_Lnk & "</div></body></html>"
            
             wSendMail.AsHTML = True
             srvSendMail.Monitor wSendMail
        End If
    End If
Next

Exit Sub

End Sub

Private Sub fgList_Display()
Dim wColor As Long, xSQL As String
Dim xCOMPTECOM As String, xTITULACLI As String
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fraList.Visible = False
fgList_Reset

fgList.Rows = 1
fgList.FormatString = fgList_FormatString
fgList.Row = 0

currentAction = "fgList_Display"
ReDim arrZADRESS0(50)
arrZADRESS0_Nb = 0
cmdList_Ok.Visible = False
'
'=======================================================================================
xCOMPTECOM = Trim(xYBIACPT0.COMPTECOM)
xSQL = "select * from " & paramIBM_Library_SAB & ".ZADRESS0" _
     & " where ADRESSNUM = '" & xCOMPTECOM & "'" _
     & " and ADRESSTYP = '2'" _
     & " and ADRESSETA = " & currentZMNURUT0.MNURUTETB _
     & " order by ADRESSCOA"
Set rsSabX = cnsab.Execute(xSQL)

Do While Not rsSabX.EOF
    V = rsZADRESS0_GetBuffer(rsSabX, meZADRESS0)
    
    If Trim(meZADRESS0.ADRESSRA1) = "" Then rsZADRESS0_CLIENARA1 meZADRESS0
    arrZADRESS0_Nb = arrZADRESS0_Nb + 1
    arrZADRESS0(arrZADRESS0_Nb) = meZADRESS0
    
    fgList.Rows = fgList.Rows + 1
    fgList.Row = fgList.Rows - 1
    fgList_DisplayLine arrZADRESS0_Nb
    rsSabX.MoveNext
Loop
'=======================================================================================
xSQL = "select TITULACLI from " & paramIBM_Library_SAB & ".ZTITULA0" _
     & " where TITULACOM = '" & xCOMPTECOM _
     & "' and  TITULATPR = '0'" _
     & " and TITULAETA = " & currentZMNURUT0.MNURUTETB
     
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    xTITULACLI = rsSab("TITULACLI")
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZADRESS0" _
         & " where ADRESSNUM = ' " & xTITULACLI & "'" _
         & " and ADRESSTYP = '1'" _
         & " and ADRESSETA = " & currentZMNURUT0.MNURUTETB _
         & " order by ADRESSCOA"
    Set rsSabX = cnsab.Execute(xSQL)
    
    Do While Not rsSabX.EOF
        V = rsZADRESS0_GetBuffer(rsSabX, meZADRESS0)
        If Trim(meZADRESS0.ADRESSRA1) = "" Then
            meZADRESS0.ADRESSNUM = xCOMPTECOM
            rsZADRESS0_CLIENARA1 meZADRESS0
            meZADRESS0.ADRESSNUM = xTITULACLI
        End If
        
        arrZADRESS0_Nb = arrZADRESS0_Nb + 1
        arrZADRESS0(arrZADRESS0_Nb) = meZADRESS0
        
        fgList.Rows = fgList.Rows + 1
        fgList.Row = fgList.Rows - 1
        fgList_DisplayLine arrZADRESS0_Nb
        rsSabX.MoveNext
    Loop
End If

Set rsSab = Nothing
arrZADRESS0_K = 1
If fgList.Rows = 2 Then cmdList_Ok.Visible = True

fraList.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgList.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub fgList_DisplayLine(lIndex As Integer)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim xSQL As String
On Error Resume Next

fgList.RowHeight(fgList.Row) = 1800
fgList.Col = 0: fgList.Text = meZADRESS0.ADRESSTYP
fgList.Col = 1: fgList.Text = Trim(meZADRESS0.ADRESSNUM)
fgList.Col = 2: fgList.Text = meZADRESS0.ADRESSCOA
fgList.Col = 3: fgList.Text = Trim(meZADRESS0.ADRESSRA1) _
                            & vbCrLf & Trim(meZADRESS0.ADRESSRA2) _
                            & vbCrLf & vbCrLf & Trim(meZADRESS0.ADRESSAD1) _
                            & vbCrLf & Trim(meZADRESS0.ADRESSAD2) _
                            & vbCrLf & Trim(meZADRESS0.ADRESSAD3) _
                            & vbCrLf & Trim(meZADRESS0.ADRESSCOP) & " " & Trim(meZADRESS0.ADRESSVIL) _
                            & vbCrLf & Trim(meZADRESS0.ADRESSPAY)

fgList.Col = 4: fgList.Text = lIndex
End Sub

Public Sub fgList_Reset()
fgList.Clear
fgList_Sort1 = 0: fgList_Sort2 = 0
fgList_Sort1_Old = -1
fgList_RowDisplay = 0: fgList_RowClick = 0
fgList_arrIndex = fgList.Cols - 1
blnfgList_DisplayLine = False
fgList_SortAD = 6
fgList.LeftCol = fgList.FixedCols

End Sub


Public Sub fgList_Sort()
If fgList.Rows > 1 Then
    fgList.Row = 1
    fgList.RowSel = fgList.Rows - 1
    
    If fgList_Sort1_Old = fgList_Sort1 Then
        If fgList_SortAD = 5 Then
            fgList_SortAD = 6
        Else
            fgList_SortAD = 5
        End If
    Else
        fgList_SortAD = 5
    End If
    fgList_Sort1_Old = fgList_Sort1
    
    fgList.Col = fgList_Sort1
    fgList.ColSel = fgList_Sort2
    fgList.Sort = fgList_SortAD
End If

End Sub

Public Sub cmdSelect_SQL_Exportation_Liste()
On Error GoTo Error_Handler
Dim X As String, K As Long, xWhere As String
Dim wFile As String, wFilex As String
Dim blnCALCS As Boolean, blnZCLIGRP0 As Boolean
Dim xSQL As String

On Error GoTo Error_Handler
'===================================================================================
xCLIGRPREG = "6000"
blnZCLIGRP0 = False
'Call rsYBIATAB0_Read("SQL_Client", "Libye", "Embargo", xWhere)
X = InputBox("par défaut : " _
    & vbCrLf & "     =========================" & vbCrLf & xCLIGRPREG _
    & vbCrLf & "     =========================", "Balance des comptes du groupe :", xCLIGRPREG)
If Trim(X) = "" Then Exit Sub
If X = "TEST" Then
    xWhere = "'0011012','0011084','0011085','0011088','0011425','0011540','0050733','0050775','0011220'"
    
    xWhere = "'0011002','0011004','0011005','0011006','0011008','0011009','0011010','0011011','0011012','0011067','0011069'" _
    & ",'0011072','0011073','0011080','0011083','0011084','0011085','0011087','0011088'" _
    & ",'0011106','0011116','0011135','0011165','0011189','0011204','0011220','0011368','0011369','0011377','0011425','0011429'" _
    & ",'0011449','0011477','0011540'" _
    & ",'0012322','0012324','0012444','0012472','0012473','0012474','0012480','0050183','0050222','0050223','0050224','0050228'" _
    & ",'0050229','0050422','0050529','0050533','0050601','0050699'" _
    & ",'0050716','0050759','0050733','0050775','0050776','0050836','0050876','0050887'" _
    & ",'0012338','0012337','0012325','0012456','0012472','0012484','0012323','0012455'"
Else
    K = Val(X)
    If K < 1000 Or K > 9999 Then
        Call MsgBox("le groupe doit être [1000 - 9999]", vbCritical, "cmdSelect_SQL_Exportation_Liste")
        Exit Sub
    Else
        xWhere = ""
        xCLIGRPREG = Format$(K, "0000000")
        xSQL = "select CLIGRPCLI from " & paramIBM_Library_SAB & ".ZCLIGRP0" _
       & " where CLIGRPREG = '" & xCLIGRPREG & "'" _
       & "  order by CLIGRPCLI"
        Set rsSab = cnsab.Execute(xSQL)
        
        Do While Not rsSab.EOF
            xWhere = xWhere & ",'" & rsSab("CLIGRPCLI") & "'"
            rsSab.MoveNext
        Loop
        If xWhere = "" Then
            Call MsgBox("Il n'y a pas de racines rattachées à ce groupe ", vbCritical, "cmdSelect_SQL_Exportation_Liste")
            Exit Sub
        Else
            Mid$(xWhere, 1, 1) = " "
        End If
    End If
End If

    

If blnAuto Then
    X = paramServer("\\CPT_Archive\")
Else
    X = ""
End If
If X = "" Then X = "C:\Temp\"
If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"

blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

wFile = X & Trim("CPT balance comptable du groupe " & xCLIGRPREG & ", au " & dateImp_Amj(YBIATAB0_DATE_CPT_J) & ".xlsx")

If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "Plan comptable : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If

If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile
'_________________________________________

X = MsgBox("Exclure les comptes 'Nostro' et les PCI = '199* 299* 399* 419* 514* 98*'", vbYesNo, "Groupe exportation")

If X = vbYes Then
    blnGroupe_Nostro_Exclus = True
Else
    blnGroupe_Nostro_Exclus = False
End If
'curBilan_Min = 0: curHors_Bilan_Min = 0

'If Not blnGroupe_Nostro_Exclus Then
    curBilan_Min = 975000: curHors_Bilan_Min = 680000
    
    X = InputBox("par défaut " _
        & vbCrLf & "     =========================" & vbCrLf & "(effacer les montants si ce filtre ne doit pas être appliqué)" _
        & vbCrLf & "     =========================", "seuils Bilan et Hors-Bilan", curBilan_Min & " " & curHors_Bilan_Min)
    
    X = Trim(X)
    If X = "" Then
        curBilan_Min = 0
        curHors_Bilan_Min = 0
    Else
        K = InStr(X, " ")
        If K > 0 Then
            curBilan_Min = Val(Mid$(X, 1, K))
            curHors_Bilan_Min = Val(Mid$(X, K, Len(X) - K + 1))
        Else
            curBilan_Min = Val(Mid$(X, 1, Len(X)))
            curHors_Bilan_Min = 0
        End If
           
    End If
'End If
'_________________________________________

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "YBIACPT0"
    .Subject = ""
End With


'Initialisation devise________________________________________________________________________________
arrDev_Nb = 0
ReDim Preserve arrDev(1000)
xSQL = "select distinct COMPTEDEV from " & paramIBM_Library_SABSPE & ".YBIACPT0" _
       & " where CLIENACLI in (" & xWhere & ")" _
       & " and COMPTEFON <> '4' " _
       & "  order by COMPTEDEV"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    arrDev_Nb = arrDev_Nb + 1
    arrDev(arrDev_Nb) = Trim(rsSab("COMPTEDEV"))
    rsSab.MoveNext
Loop
ReDim Preserve arrDev(arrDev_Nb + 1)
ReDim arrDev_RowT(arrDev_Nb + 1), arrDev_Cours(arrDev_Nb + 1)

'Initialisation pays________________________________________________________________________________
Call rsZBASTAB0_Pays(arrPays(), arrPays_NB)

'===================================================================================

mXls1_Row_C = 1
Call cmdSelect_SQL_Exportation_Liste_Init(1)
Call cmdSelect_SQL_Exportation_Liste_Init(2)
lstW.Clear


Call cmdSelect_SQL_Exportation_Liste_Detail(xWhere)
cmdSelect_SQL_Exportation_Liste_Detail_T2

'__________________________________________________________________________________
Exit_sub:
'__________________________________________________________________________________

Set rsSab = Nothing

wbExcel.SaveAs wFile
wbExcel.Close
appExcel.Quit

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents
'_____________________________
Exit Sub

Error_Handler:

If Not blnCALCS Then
    X = "C:\Temp\"
    Resume Next
End If

    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents

End Sub

Public Sub cmdSelect_SQL_Exportation_Liste_Detail(lWhere As String)
On Error GoTo Error_Handler
Dim X As String, xWhere As String
Dim K As Integer, K1 As Integer
Dim curMTD As Currency, curMTE As Currency, dblX As Double
Dim blnOk As Boolean

Dim wColor As Long

Call rsYBIACPT0_Init(oldYBIACPT0)
mXls2_row_T = 0

'==========================================================================================================

Set wsExcel = wbExcel.Sheets(3)
wsExcel.Name = "Détail"

'__________________________________________________________________________________

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .HorizontalAlignment = Excel.xlHAlignCenter
    .WrapText = True
    .Font.Size = 7
    .Font.Name = "Arial Unicode MS"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 85
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14CPT balance comptable du groupe " & xCLIGRPREG & ", au " & dateImp_Amj(YBIATAB0_DATE_CPT_J) _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True
wsExcel.PageSetup.Zoom = 75
wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"


Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

wsExcel.Columns(1).ColumnWidth = 8: wsExcel.Cells(1, 1) = "Racine"
wsExcel.Columns(2).ColumnWidth = 8: wsExcel.Cells(1, 2) = "PCI"
wsExcel.Columns(3).ColumnWidth = 8: wsExcel.Cells(1, 3) = "Produit"
wsExcel.Columns(4).ColumnWidth = 8: wsExcel.Cells(1, 4) = "Devise"
wsExcel.Columns(5).ColumnWidth = 15: wsExcel.Cells(1, 5) = "Solde en devise"
wsExcel.Columns(5).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(6).ColumnWidth = 15: wsExcel.Cells(1, 6) = "CV Solde en €"
wsExcel.Columns(6).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(6).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(7).ColumnWidth = 32: wsExcel.Cells(1, 7) = "Client": wsExcel.Columns(7).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(8).ColumnWidth = 16: wsExcel.Cells(1, 8) = "Compte": wsExcel.Columns(8).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(9).ColumnWidth = 6: wsExcel.Cells(1, 9) = "Blocage"
wsExcel.Columns(10).ColumnWidth = 10: wsExcel.Cells(1, 10) = "Date mvt"
wsExcel.Columns(11).ColumnWidth = 32: wsExcel.Cells(1, 11) = "Intitulé": wsExcel.Columns(11).HorizontalAlignment = Excel.xlHAlignLeft

mXls2_Col = 11

For K = 1 To mXls2_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next
'==========================================================================================================
X = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0, " & paramIBM_Library_SAB & ".ZPLAN0" _
       & " where CLIENACLI in (" & lWhere & ")" _
       & " and COMPTEFON <> '4'  And SOLDECEN <> '000000000000000000' " _
       & " and PLANCOOBL = COMPTEOBL" _
       & " order by CLIENACLI,COMPTEOBL, COMPTEDEV, COMPTECOM"
Set rsSab = cnsab.Execute(X)
mXls2_Row = 1
Do While Not rsSab.EOF
    V = rsYBIACPT0_GetBuffer(rsSab, xYBIACPT0)
    
    blnOk = True
    If blnGroupe_Nostro_Exclus Then
        If Mid$(xYBIACPT0.PLANCOPRO, 1, 2) = "NO" Then
            blnOk = False
        Else
            If Mid$(xYBIACPT0.COMPTEOBL, 1, 2) = "98" Then
                blnOk = False
            Else
                Select Case Mid$(xYBIACPT0.COMPTEOBL, 1, 3)
                    Case "199", "299", "399", "419", "514": blnOk = False
                End Select
            End If
        End If
    End If
    
    If blnOk Then
        xZPLAN0.PLANINTIT = rsSab("PLANINTIT")
        If oldYBIACPT0.CLIENACLI <> xYBIACPT0.CLIENACLI Then
            Call cmdSelect_SQL_Exportation_Liste_Detail_T(1)
        Else
            If oldYBIACPT0.COMPTEOBL <> xYBIACPT0.COMPTEOBL Then
                Call cmdSelect_SQL_Exportation_Liste_Detail_T(2)
            Else
                If oldYBIACPT0.COMPTEDEV <> xYBIACPT0.COMPTEDEV Then Call cmdSelect_SQL_Exportation_Liste_Detail_T(3)
            End If
        End If
        
        mXls2_Row = mXls2_Row + 1
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "> " & xYBIACPT0.CLIENACLI & " " & xYBIACPT0.COMPTECOM): DoEvents
        If fctUser_Classe_Aut(xYBIACPT0.COMPTECLA) Then
            curMTD = -xYBIACPT0.SOLDECEN
        Else
            curMTD = -999999999999.99
        End If
        wsExcel.Cells(mXls2_Row, 1) = xYBIACPT0.CLIENACLI
        wsExcel.Cells(mXls2_Row, 2) = xYBIACPT0.COMPTEOBL
        wsExcel.Cells(mXls2_Row, 3) = xYBIACPT0.PLANCOPRO
        wsExcel.Cells(mXls2_Row, 4) = xYBIACPT0.COMPTEDEV
        wsExcel.Cells(mXls2_Row, 5) = curMTD
        wsExcel.Cells(mXls2_Row, 7) = Trim(xYBIACPT0.CLIENARA1) & " " & Trim(xYBIACPT0.CLIENARA2)
        wsExcel.Cells(mXls2_Row, 8) = xYBIACPT0.COMPTECOM
        wsExcel.Cells(mXls2_Row, 9) = xYBIACPT0.COMPTEFON
        If xYBIACPT0.SOLDEDMO > 0 Then wsExcel.Cells(mXls2_Row, 10) = dateImp10(xYBIACPT0.SOLDEDMO + 19000000)
        wsExcel.Cells(mXls2_Row, 11) = xYBIACPT0.COMPTEINT
    
        If curMTD = 0 Then
            curMTE = 0
        Else
            If xYBIACPT0.COMPTEDEV = "EUR" Then
                curMTE = curMTD
            Else
                Call sqlYBIATAB0_Read("PDC", xYBIACPT0.COMPTEDEV, YBIATAB0_DATE_CPT_J, X)
                If IsNumeric(Mid$(X, 9, 15)) Then
                    dblX = CDbl(Mid$(X, 9, 15) / 1000000000)
                    If dblX <> 0 Then curMTE = Round(curMTD / dblX, 2)
                Else
                    curMTE = 0
                End If
            End If
        End If
    
        wsExcel.Cells(mXls2_Row, 6) = curMTE
        
        'If xYBIACPT0.CLIENACLI = "0050183" Then
        '    Debug.Print 50183
        'End If
        If Mid$(xYBIACPT0.COMPTEOBL, 1, 1) <> "9" Then
            If curMTE < 0 Then
                If sBilan_DB = "" Then
                    sBilan_DB = "Détail!F" & mXls2_Row
                Else
                    sBilan_DB = sBilan_DB & "+Détail!F" & mXls2_Row
                End If
            Else
                If sBilan_CR = "" Then
                    sBilan_CR = "Détail!F" & mXls2_Row
                Else
                    sBilan_CR = sBilan_CR & "+Détail!F" & mXls2_Row
                End If
            End If
        Else
            If Mid$(xYBIACPT0.COMPTEOBL, 1, 2) <> "98" Then
                If curMTE < 0 Then
                    If sHors_Bilan_DB = "" Then
                         sHors_Bilan_DB = "Détail!F" & mXls2_Row
                    Else
                         sHors_Bilan_DB = sHors_Bilan_DB & "+Détail!F" & mXls2_Row
                    End If
                Else
                    If sHors_Bilan_CR = "" Then
                         sHors_Bilan_CR = "Détail!F" & mXls2_Row
                    Else
                         sHors_Bilan_CR = sHors_Bilan_CR & "+Détail!F" & mXls2_Row
                    End If
                End If
            End If
            
        End If
        
        
        If Mid$(xYBIACPT0.COMPTEOBL, 1, 2) = "98" Then
            For K = 1 To mXls2_Col
                wsExcel.Cells(mXls2_Row, K).Interior.Color = RGB(220, 220, 220)
            Next K
        End If
End If


'===================================================================================================
    rsSab.MoveNext
Loop

xYBIACPT0.CLIENACLI = ""
xYBIACPT0.CLIENARA1 = "": xYBIACPT0.CLIENARA2 = ""

Call cmdSelect_SQL_Exportation_Liste_Detail_T(0)

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name


End Sub
Public Sub cmdSendMail(lFileName As String)
Dim wSendMail As typeSendMail
wSendMail.FromDisplayName = "@BAL_6000"
wSendMail.RecipientDisplayName = "CPT"

wSendMail.Subject = "Relevé des mouvements comptables du groupe 6000 (Libye)"
wSendMail.Attachment = lFileName
wSendMail.Message = ""

wSendMail.AsHTML = True
srvSendMail.Monitor wSendMail

End Sub
Public Sub cmdSelect_SQL_Exportation_Liste_Init(lSheet As Integer)
Dim K As Integer, K2 As Integer

Set wsExcel = wbExcel.Sheets(lSheet)

Select Case lSheet
    Case 1: wsExcel.Name = "CPT-PCI"
                wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14CPT balance comptable par racine du groupe " & xCLIGRPREG & ", arrêté au " & dateImp10(YBIATAB0_DATE_CPT_J) _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr

    Case 2: wsExcel.Name = "PCI-CPT"
            wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14CPT balance par rubrique comptable du groupe " & xCLIGRPREG & ", arrêté au " & dateImp10(YBIATAB0_DATE_CPT_J) _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr

End Select

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignRight
    .WrapText = False ' True
    .Font.Size = 8
    .Font.Name = "Arial Unicode MS"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 50
wsExcel.PageSetup.PrintTitleRows = "$A1:$L1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

wsExcel.PageSetup.CenterHorizontally = True

mXls1_Col = arrDev_Nb + 2
mXls1_Row = 1: mXls1_Row_T = 0
wsExcel.PageSetup.PrintTitleRows = "$A1:$" & Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", mXls1_Col, 1) & "1"


wsExcel.Columns(1).ColumnWidth = 7:  wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
Select Case lSheet
    Case 1: wsExcel.Cells(mXls1_Row_C, 1) = "Racine"
    Case 2: wsExcel.Cells(mXls1_Row_C, 1) = "Rubrique"
End Select
wsExcel.Cells(mXls1_Row_C, 1).Interior.Color = mColor_GB: wsExcel.Cells(mXls1_Row_C, 1).Font.Color = mColor_Z0
wsExcel.Cells(mXls1_Row_C, 2).Interior.Color = mColor_GB: wsExcel.Cells(mXls1_Row_C, 2).Font.Color = mColor_Z0

wsExcel.Columns(2).ColumnWidth = 30: wsExcel.Cells(mXls1_Row_C, 2) = "Intitulé":  wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignLeft

mXls1_Col = mXls1_Col + 1: colBilan_DB = mXls1_Col
alfBilan_DB = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", mXls1_Col, 1)
wsExcel.Columns(mXls1_Col).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, mXls1_Col) = "Bilan DB € "
wsExcel.Columns(mXls1_Col).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(mXls1_Row_C, mXls1_Col).Interior.Color = mColor_GB
wsExcel.Cells(mXls1_Row_C, mXls1_Col).Font.Color = mColor_Z0

mXls1_Col = mXls1_Col + 1: colBilan_CR = mXls1_Col
alfBilan_CR = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", mXls1_Col, 1)
wsExcel.Columns(mXls1_Col).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, mXls1_Col) = "Bilan CR € "
wsExcel.Columns(mXls1_Col).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(mXls1_Row_C, mXls1_Col).Interior.Color = mColor_GB
wsExcel.Cells(mXls1_Row_C, mXls1_Col).Font.Color = mColor_Z0


mXls1_Col = mXls1_Col + 1: colHors_Bilan_DB = mXls1_Col
alfHors_Bilan_DB = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", mXls1_Col, 1)
wsExcel.Columns(mXls1_Col).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, mXls1_Col) = "Hors-Bilan DB € "
wsExcel.Columns(mXls1_Col).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(mXls1_Row_C, mXls1_Col).Interior.Color = mColor_GB
wsExcel.Cells(mXls1_Row_C, mXls1_Col).Font.Color = mColor_Z0

mXls1_Col = mXls1_Col + 1: colHors_Bilan_CR = mXls1_Col
alfHors_Bilan_CR = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", mXls1_Col, 1)
wsExcel.Columns(mXls1_Col).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, mXls1_Col) = "Hors-Bilan CR € "
wsExcel.Columns(mXls1_Col).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(mXls1_Row_C, mXls1_Col).Interior.Color = mColor_GB
wsExcel.Cells(mXls1_Row_C, mXls1_Col).Font.Color = mColor_Z0


For K = 1 To arrDev_Nb
    K2 = K + 2
    wsExcel.Cells(mXls1_Row_C, K2) = arrDev(K)
    wsExcel.Cells(mXls1_Row_C, K2).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_C, K2).Font.Color = mColor_Z0
    wsExcel.Columns(K2).ColumnWidth = 13: wsExcel.Columns(K2).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
    
    If arrDev(K) = "EUR" Or arrDev(K) = "USD" Then wsExcel.Columns(K2).ColumnWidth = 16
Next K

End Sub
Public Sub cmdSelect_SQL_Exportation_Liste_Detail_T(lTotal As Integer)
Dim K As Integer, X As String
Dim blnBilan_Min As Boolean, blnHors_Bilan_Min As Boolean
On Error GoTo Error_Handler
Set wsExcel = wbExcel.Sheets(1)

If mXls2_row_T > 0 Then
    For K = 1 To arrDev_Nb
        If oldYBIACPT0.COMPTEDEV = arrDev(K) Then
            wsExcel.Cells(mXls1_Row, K + 2).FormulaLocal = "=SOMME(Détail!E" & mXls2_row_T & ":Détail!E" & mXls2_Row & ")"
            'wsExcel.Cells(mXls1_Row, mXls1_Col).FormulaLocal = "=SOMME(Détail!F" & mXls2_row_T & ":Détail!F" & mXls2_Row & ")"
            If sBilan_DB <> "" Then wsExcel.Cells(mXls1_Row, colBilan_DB).FormulaLocal = "=SOMME(" & sBilan_DB & ")"
            If sBilan_CR <> "" Then wsExcel.Cells(mXls1_Row, colBilan_CR).FormulaLocal = "=SOMME(" & sBilan_CR & ")"
            If sHors_Bilan_DB <> "" Then wsExcel.Cells(mXls1_Row, colHors_Bilan_DB).FormulaLocal = "=SOMME(" & sHors_Bilan_DB & ")"
            If sHors_Bilan_CR <> "" Then wsExcel.Cells(mXls1_Row, colHors_Bilan_CR).FormulaLocal = "=SOMME(" & sHors_Bilan_CR & ")"
            Exit For
        End If
    Next K

    
End If
mXls2_row_T = mXls2_Row + 1

oldYBIACPT0 = xYBIACPT0

If lTotal < 2 Then
    If mXls1_Row_T > 0 Then
        wsExcel.Cells(mXls1_Row_T, colBilan_DB).FormulaLocal = "=SOMME(" & alfBilan_DB & mXls1_Row_T + 1 & ":" & alfBilan_DB & mXls1_Row & ")"
        wsExcel.Cells(mXls1_Row_T, colBilan_CR).FormulaLocal = "=SOMME(" & alfBilan_CR & mXls1_Row_T + 1 & ":" & alfBilan_CR & mXls1_Row & ")"
        wsExcel.Cells(mXls1_Row_T, colHors_Bilan_DB).FormulaLocal = "=SOMME(" & alfHors_Bilan_DB & mXls1_Row_T + 1 & ":" & alfHors_Bilan_DB & mXls1_Row & ")"
        wsExcel.Cells(mXls1_Row_T, colHors_Bilan_CR).FormulaLocal = "=SOMME(" & alfHors_Bilan_CR & mXls1_Row_T + 1 & ":" & alfHors_Bilan_CR & mXls1_Row & ")"
        
        blnBilan_Min = False: blnHors_Bilan_Min = False
        If Abs(wsExcel.Cells(mXls1_Row_T, colBilan_DB)) + wsExcel.Cells(mXls1_Row_T, colBilan_CR) >= curBilan_Min Then
            For K = 1 To mXls1_Col: wsExcel.Cells(mXls1_Row_T, K).Interior.Color = mColor_G0: Next K
            blnBilan_Min = True
        End If
        
        If Abs(wsExcel.Cells(mXls1_Row_T, colHors_Bilan_DB)) + wsExcel.Cells(mXls1_Row_T, colHors_Bilan_CR) >= curHors_Bilan_Min Then
            For K = 1 To mXls1_Col: wsExcel.Cells(mXls1_Row_T, K).Interior.Color = mColor_G0: Next K
            blnHors_Bilan_Min = True
        End If
        If blnBilan_Min Then
            wsExcel.Cells(mXls1_Row_T, colBilan_DB).Interior.Color = mColor_G1
            wsExcel.Cells(mXls1_Row_T, colBilan_CR).Interior.Color = mColor_G1
        End If
        If blnHors_Bilan_Min Then
            wsExcel.Cells(mXls1_Row_T, colHors_Bilan_DB).Interior.Color = mColor_G1
            wsExcel.Cells(mXls1_Row_T, colHors_Bilan_CR).Interior.Color = mColor_G1
        End If
        X = alfBilan_DB & mXls1_Row_T + 1 & ":" & alfBilan_DB & mXls1_Row
        wsExcel.Cells(mXls1_Row_T, colBilan_DB).FormulaLocal = "=SI(SOMME(" & X & ")= 0;" & Asc34 & Asc34 & ";SOMME(" & X & "))"
        wsExcel.Cells(mXls1_Row_T, colBilan_DB).Font.Bold = True
         
        X = alfBilan_CR & mXls1_Row_T + 1 & ":" & alfBilan_CR & mXls1_Row
        wsExcel.Cells(mXls1_Row_T, colBilan_CR).FormulaLocal = "=SI(SOMME(" & X & ")= 0;" & Asc34 & Asc34 & ";SOMME(" & X & "))"
        wsExcel.Cells(mXls1_Row_T, colBilan_CR).Font.Bold = True
        
        X = alfHors_Bilan_DB & mXls1_Row_T + 1 & ":" & alfHors_Bilan_DB & mXls1_Row
        wsExcel.Cells(mXls1_Row_T, colHors_Bilan_DB).FormulaLocal = "=SI(SOMME(" & X & ")= 0;" & Asc34 & Asc34 & ";SOMME(" & X & "))"
        wsExcel.Cells(mXls1_Row_T, colHors_Bilan_DB).Font.Bold = True
         
        X = alfHors_Bilan_CR & mXls1_Row_T + 1 & ":" & alfHors_Bilan_CR & mXls1_Row
        wsExcel.Cells(mXls1_Row_T, colHors_Bilan_CR).FormulaLocal = "=SI(SOMME(" & X & ")= 0;" & Asc34 & Asc34 & ";SOMME(" & X & "))"
        wsExcel.Cells(mXls1_Row_T, colHors_Bilan_CR).Font.Bold = True
   End If
    
    mXls1_Row = mXls1_Row + 1
    For K = 1 To mXls1_Col
    '    wsExcel.Cells(mXls1_Row, K) = wsExcel.Cells(1, K)
    '    wsExcel.Cells(mXls1_Row, K).Interior.Color = wsExcel.Cells(1, K).Interior.Color
    '    wsExcel.Cells(mXls1_Row, K).Font.Color = wsExcel.Cells(1, K).Font.Color
        wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(190, 230, 255) 'RGB(255, 255, 190) 'mColor_Y1
    Next K
    wsExcel.Cells(mXls1_Row, 1) = oldYBIACPT0.CLIENACLI
            wsExcel.Cells(mXls1_Row, 1).Font.Bold = True
            wsExcel.Cells(mXls1_Row, 1).Font.Color = vbBlue
    wsExcel.Cells(mXls1_Row, 2) = Trim(oldYBIACPT0.CLIENARA1) & " " & Trim(oldYBIACPT0.CLIENARA2)
            wsExcel.Cells(mXls1_Row, 2).Font.Color = vbBlue
            wsExcel.Cells(mXls1_Row, 2).Font.Size = 7
            wsExcel.Cells(mXls1_Row, 2).Font.Bold = True
    mXls1_Row_T = mXls1_Row
    
    
    If lTotal > 0 Then
        For K = 1 To arrPays_NB
            If Trim(oldYBIACPT0.CLIENARSD) = arrPays(K).Id Then
                mXls1_Row = mXls1_Row + 1
                wsExcel.Cells(mXls1_Row, 2) = oldYBIACPT0.CLIENARSD & " " & Trim(arrPays(K).Nom)
                wsExcel.Cells(mXls1_Row, 2).Font.Bold = True
                wsExcel.Cells(mXls1_Row, 2).Font.Color = vbBlue
                wsExcel.Cells(mXls1_Row, 2).Font.Size = 7
                Exit For
            End If
        Next K
    End If

End If

If lTotal < 3 And lTotal > 0 Then
    sBilan_DB = "": sBilan_CR = ""
    sHors_Bilan_DB = "": sHors_Bilan_CR = ""

    mXls1_Row = mXls1_Row + 1
    wsExcel.Cells(mXls1_Row, 1) = oldYBIACPT0.COMPTEOBL
    wsExcel.Cells(mXls1_Row, 2) = xZPLAN0.PLANINTIT
    wsExcel.Cells(mXls1_Row, colBilan_DB).Interior.Color = mColor_G0
    wsExcel.Cells(mXls1_Row, colBilan_CR).Interior.Color = mColor_G0
    wsExcel.Cells(mXls1_Row, colHors_Bilan_DB).Interior.Color = mColor_G0
    wsExcel.Cells(mXls1_Row, colHors_Bilan_CR).Interior.Color = mColor_G0
    If Mid$(oldYBIACPT0.COMPTEOBL, 1, 2) = "98" Then
        For K = 1 To mXls1_Col - 4
                wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(220, 220, 220)
        Next K
        wsExcel.Cells(mXls1_Row, colBilan_DB) = ""
        wsExcel.Cells(mXls1_Row, colBilan_CR) = ""
        wsExcel.Cells(mXls1_Row, colHors_Bilan_DB) = ""
        wsExcel.Cells(mXls1_Row, colHors_Bilan_DB) = ""

    End If
    lstW.AddItem oldYBIACPT0.COMPTEOBL & "|" & oldYBIACPT0.CLIENACLI & "|" & mXls1_Row & "|" & mXls1_Row_T
    
End If

'If lTotal = 2 Then
 '   lstW.AddItem oldYBIACPT0.COMPTEOBL & "|" & oldYBIACPT0.CLIENACLI & "|" & mXls1_Row & "|" & mXls1_Row_T
'End If
Set wsExcel = wbExcel.Sheets(3)
Exit Sub

Error_Handler:
End Sub

Private Sub cmdBalance_Service()

SSTab1.Tab = 0


If blnAuto Then
    blnBalance_Stock_détail = True
    blnService_Printer = True
Else
    X = MsgBox("Voulez-vous imprimer le détail par services ?", vbYesNo + vbQuestion + vbDefaultButton1, "Sab_Balance : Impression Balance / Stock")
    If X = vbYes Then
        blnBalance_Stock_détail = True
        
        X = MsgBox("Voulez-vous envoyer un mail aux services avec le lien hypertexte ?", vbYesNo + vbQuestion + vbDefaultButton1, "Sab_Balance : Impression Balance / Stock")
        If X = vbYes Then
            blnService_Printer = True
        Else
            blnService_Printer = False
        End If
    
    Else
        blnBalance_Stock_détail = False
    End If
End If

cmdBalance_Ok_Stock
If blnAuto Then
    Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S60", "BIA-BAL-Stock_Cumul", "Archive")
Else
    Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S60", "BIA-BAL-Stock_Cumul", "Prod")
End If


Call prtSAB_Balance_Cumul_Monitor(arrService_Balance_Cumul(), arrService_Nb, arrDevise_Nb)
'Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("", "", "Balance / Stock - Cumul")

If blnAuto Then
    Dim xPath As String, X1 As String, X2 As String, objFolder, objFiles, fsoFile As File
    Dim wSendMail As typeSendMail
    
    xPath = paramEditionNoPaper_Folder & "PDF\Archive_" & YBIATAB0_DATE_CPT_J
    Set objFolder = msFileSystem.GetFolder(xPath)
    Set objFiles = objFolder.Files
    For Each fsoFile In objFiles
        If InStr(fsoFile.Name, "BIA-BAL-Stock_G3") > 0 Then
            If InStr(fsoFile.Name, "(S32)") > 0 Then
                X1 = xPath & "\" & fsoFile.Name
                X2 = Replace(Replace(X1, "S32", "S10"), "GDC", "SOBI")
                If Dir(X2) <> "" Then Kill X2
                msFileSystem.CopyFile X1, X2
                paramEditionNoPaper_Auto_Lnk = "<span style='font-size:9.0pt;font-family:Calibri'>""" _
                                         & "<A HREF=" & Asc34 & Replace(X2, paramEditionNoPaper_Folder & "PDF\", paramEditionNoPaper_Partage) & Asc34 & ">" _
                                        & "Cliquez ici pour afficher le document : " & X2 & "</A><BR><BR>"
        
                wSendMail.From = currentSSIWINMAIL
                wSendMail.FromDisplayName = "NoPaper BIA-BAL-Stock_G3"
                wSendMail.Recipient = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire("S10")
                wSendMail.CcRecipient = ""
                wSendMail.Subject = "BIA-BAL-Stock_G3 "
                wSendMail.Attachment = ""
                wSendMail.Message = mHtml_Head & "<span style='font-size:10.0pt;font-family:Calibri'>" _
                                 & htmlFontColor_Black & "<BR><BR>" & paramEditionNoPaper_Auto_Lnk & "</div></body></html>"
                
                 wSendMail.AsHTML = True
                 srvSendMail.Monitor wSendMail
            End If
        End If
    Next
End If

Exit Sub


End Sub
Public Sub fgYBIASTO0_Sort()
If fgYBIASTO0.Rows > 1 Then
    fgYBIASTO0.Row = 1
    fgYBIASTO0.RowSel = fgYBIASTO0.Rows - 1
    
    If fgYBIASTO0_Sort1_Old = fgYBIASTO0_Sort1 Then
        If fgYBIASTO0_SortAD = 5 Then
            fgYBIASTO0_SortAD = 6
        Else
            fgYBIASTO0_SortAD = 5
        End If
    Else
        fgYBIASTO0_SortAD = 5
    End If
    fgYBIASTO0_Sort1_Old = fgYBIASTO0_Sort1
    
    fgYBIASTO0.Col = fgYBIASTO0_Sort1
    fgYBIASTO0.ColSel = fgYBIASTO0_Sort2
    fgYBIASTO0.Sort = fgYBIASTO0_SortAD
End If
'cboDevise_Reset
End Sub
Public Sub fgYBIASTO0_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgYBIASTO0.Row

If lRow > 0 And lRow < fgYBIASTO0.Rows Then
    fgYBIASTO0.Row = lRow
    For I = 0 To fgYBIASTO0_arrIndex
        fgYBIASTO0.Col = I: fgYBIASTO0.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgYBIASTO0.Row = mRow
    If fgYBIASTO0.Row > 0 Then
        lRow = fgYBIASTO0.Row
        lColor_Old = fgYBIASTO0.CellBackColor
        For I = 0 To fgYBIASTO0_arrIndex
          fgYBIASTO0.Col = I: fgYBIASTO0.CellBackColor = lColor
        Next I
        fgYBIASTO0.Col = 0
    End If
End If

End Sub
Private Sub fgYBIASTO0_Display()
Dim intReturn As Integer
Dim xSQL As String
Dim curTotal As Currency, nbTotal As Long, curSolde As Currency
fgYBIASTO0_Reset
fgYBIASTO0.Rows = 1
fgYBIASTO0.FormatString = fgYBIASTO0_FormatString
fgYBIAMVT0.Visible = False

Set rsSab = Nothing
libYBIASTO0_Diff = ""
curTotal = 0: nbTotal = 0

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIASTO0 where " _
     & "YSTOPCI like '" & Mid$(xYBIACPT0.COMPTEOBL, 1, 5) & "%'" _
     & "AND YSTODEV = '" & xYBIACPT0.COMPTEDEV & "'" _
     & "AND YSTOCLI = " & Val(xYBIACPT0.CLIENACLI)
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYBIASTO0_GetBuffer(rsSab, xYBIASTO0)
    If Not IsNull(V) Then
        MsgBox V, vbCritical, "frmSAB_Balance.SQL_ODBC"
        Exit Sub
    Else
        fgYBIASTO0_DisplayLine
        curTotal = curTotal + xYBIASTO0.YSTOMON
        nbTotal = nbTotal + 1
    End If
    rsSab.MoveNext
Loop
libYBIASTO0_Total = nbTotal & " dossiers : " & xYBIACPT0.COMPTEDEV
libYBIASTO0_Solde = "Solde " & xYBIACPT0.COMPTEDEV & " au " & dateIBM10(YBIATAB0_DATE_CPT_J, True)

libYBIASTO0_YSTOMON = Format$(curTotal, "### ### ### ###.00")
curSolde = Abs(xYBIACPT0.SOLDECEN)
libYBIASTO0_SOLDECEN = Format$(curSolde, "### ### ### ###.00")

If curTotal = curSolde Then
    libYBIASTO0_Diff = ""
Else
    libYBIASTO0_Diff.ForeColor = vbRed
    libYBIASTO0_Diff = Format$(curTotal - curSolde, "### ### ### ###.00")
End If

'libYBIASTO0 = nbTotal & " dossiers, " & Format$(curTotal, "### ### ### ###.00") & " " & xYBIACPT0.COMPTEDEV
fgYBIASTO0.Visible = True
fgYBIASTO0_Sort1 = -1
End Sub


Public Sub fgYBIASTO0_DisplayLine()
On Error Resume Next
fgYBIASTO0.Rows = fgYBIASTO0.Rows + 1
fgYBIASTO0.Row = fgYBIASTO0.Rows - 1
fgYBIASTO0.Col = 0: fgYBIASTO0.Text = xYBIASTO0.YSTOAPP & " " & xYBIASTO0.YSTOOPE & " " & xYBIASTO0.YSTONUM
fgYBIASTO0.Col = 1: fgYBIASTO0.Text = Format$(xYBIASTO0.YSTOMON, "### ### ### ###.00")
fgYBIASTO0.Col = 2: fgYBIASTO0.Text = dateIBM10(xYBIASTO0.YSTODEB, True)
fgYBIASTO0.Col = 3: fgYBIASTO0.Text = dateIBM10(xYBIASTO0.YSTOFIN, True)

End Sub

Public Sub fgYBIASTO0_Reset()
fgYBIASTO0.Clear
fgYBIASTO0_Sort1 = 0: fgYBIASTO0_Sort2 = 0
fgYBIASTO0_Sort1_Old = -1
fgYBIASTO0_RowDisplay = 0: fgYBIASTO0_RowClick = 0
fgYBIASTO0_arrIndex = fgYBIASTO0.Cols - 1
blnfgYBIASTO0_DisplayLine = False
End Sub


Public Sub fgYAUTE1I0_Display(lFct As String)

Dim xSQL As String
Dim X As String

mfgYAUTE1I0_Fct = lFct
mAUTE1ICLI_Aut = ""
arrCellBackColor(1) = mColor_G2: arrCellBackColor(2) = mColor_B1
arrCellBackColor_K = 2

X = Trim(txtSelect_CLIENARES)
If X <> "" Then X = " : " & X  '& " - " & arrBIA_RCOM_Lib(Val(Mid$(X, 2, 2)))

lblYAUTE1I0 = "Position Comptable / Autorisations au " & dateImp10_S(YBIATAB0_DATE_CPT_J) & X
            
mAUTE1IDAF_Min8 = dateElp("MoisAdd", -1, DSys)
mAUTE1IDAF_Min7 = mAUTE1IDAF_Min8 - 19000000
mAUTE1IDAF_Max8 = dateElp("MoisAdd", 3, DSys)
mAUTE1IDAF_Max7 = mAUTE1IDAF_Max8 - 19000000
fgYAUTE1I0_Reset
fgYAUTE1I0.Rows = 1
fgYAUTE1I0.FormatString = fgYAUTE1I0_FormatString
fgYAUTE1I0.Visible = False
fgYAUTE1I0.Row = 0
fgYAUTE1I0.Col = 3: fgYAUTE1I0.CellAlignment = 1
fgYAUTE1I0.Col = 4: fgYAUTE1I0.CellAlignment = 1
fgYAUTE1I0.Col = 5: fgYAUTE1I0.CellAlignment = 1
fgYAUTE1I0.Col = 6: fgYAUTE1I0.CellAlignment = 1
fgYAUTE1I0.Col = 7: fgYAUTE1I0.CellAlignment = 2

Select Case lFct
    Case "Echeance"
            X = Trim(txtSelect_CLIENARES)
            If X <> "" Then X = " and AUTE1Ires = '" & X & "'"
            xSQL = "select * from " & paramIBM_Library_SAB & ".ZAUTE1I0 , " & paramIBM_Library_SAB & ".ZCLIENA0" _
                 & " where AUTE1IETA = 1 " _
                 & X _
                 & " and AUTE1IDAF >= " & mAUTE1IDAF_Min7 & " and AUTE1IDAF <= " & mAUTE1IDAF_Max7 _
                 & " and   AUTE1ITYP = '8'" _
                 & " and CLIENACLI = AUTE1ICLI " _
                 & " order by AUTE1ICLI , AUTE1IOR1 , AUTE1IOR2 , AUTE1IOR3 desc , AUTE1IOR4"
    Case "Dépassement"
            X = Trim(txtSelect_CLIENARES)
            If X <> "" Then X = " and AUTE1Ires = '" & X & "'"
            xSQL = "select * from " & paramIBM_Library_SAB & ".ZAUTE1I0 , " & paramIBM_Library_SAB & ".ZCLIENA0" _
                 & " where AUTE1IETA = 1 " _
                 & X _
                 & " and AUTE1IMTD <> 0" _
                 & " and CLIENACLI = AUTE1ICLI " _
                 & " order by AUTE1ICLI , AUTE1IOR1 , AUTE1IOR2 , AUTE1IOR3 desc , AUTE1IOR4"
    Case Else
            xSQL = "select * from " & paramIBM_Library_SAB & ".ZAUTE1I0 , " & paramIBM_Library_SAB & ".ZCLIENA0" _
                 & " where AUTE1IETA = 1 and AUTE1IGRP = ''" _
                 & " and AUTE1ICLI = '" & mAUTE1ICLI & "'" _
                 & " and ( AUTE1ITYP = '2'" _
                 & " or   AUTE1ITYP = '8')" _
                 & " and CLIENACLI = AUTE1ICLI " _
                 & " order by AUTE1IOR1 , AUTE1IOR2 , AUTE1IOR3 desc , AUTE1IOR4"
End Select


Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    If rsSab("AUTE1IRES") = "R60" Or rsSab("AUTE1IRES") = "R61" Or rsSab("AUTE1IRES") = "R62" Then
    Else
        fgYAUTE1I0_DisplayLine
    End If
    rsSab.MoveNext
Loop


Call lstErr_Clear(Me.lstErr, Me.cmdContext, "affichage mvt : " & fgYAUTE1I0.Rows - 1)

fgYAUTE1I0.Visible = True
fgYAUTE1I0_Sort1 = -1

End Sub


Public Sub fgYAUTE1I0_Display_GRP(lK1 As String)

Dim intReturn As Integer
Dim lenK1 As Integer
Dim blnDisplay As Boolean

'$JPL 20050420 à revoir
'$$$$$$$$$$$$$$$$$$$$$$$$$$$

fgYAUTE1I0_Reset
picYAUTE1I0.Visible = False

fgYAUTE1I0.Rows = 1
fgYAUTE1I0.FormatString = fgYAUTE1I0_FormatString
fgYAUTE1I0.Visible = False

If optYAUTE1i0_NIV_1 Then mYAUTE1I0_AUTE1INIV = 1
If optYAUTE1i0_NIV_2 Then mYAUTE1I0_AUTE1INIV = 2
If optYAUTE1i0_NIV_X Then mYAUTE1I0_AUTE1INIV = 999
'meYbase.ID = constYAUTE1I0
'meYbase.K1 = lK1
'lenK1 = Len(lK1)
'meYbase.Method = "Seek>"

'Do
'    intReturn = tableYBase_Read(meYbase)
'    If Trim(meYbase.ID) <> constYAUTE1I0 Then intReturn = -1
'    If intReturn = 0 Then
'       If Mid$(meYbase.K1, 1, lenK1) = lK1 Then
'            MsgTxt = Space$(34) & meYbase.Text
'            MsgTxtIndex = 0
'            srvYAUTE1I0_GetBuffer xYAUTE1I0
'            blnDisplay = False
'            If xYAUTE1I0.AUTE1IMDB = 0 And xYAUTE1I0.AUTE1IMCR = 0 And xYAUTE1I0.AUTE1IMAU = 0 Then
'            Else
'                If xYAUTE1I0.AUTE1ITYP = 1 And xYAUTE1I0.AUTE1IDEV = xYAUTE1I0.AUTE1IDBA Then
'                Else
'                   If xYAUTE1I0.AUTE1INIV <= mYAUTE1I0_AUTE1INIV Then blnDisplay = True
'                End If
'            End If
'            If xYAUTE1I0.AUTE1IDEP <> " " Then blnDisplay = True
'            If blnDisplay Then fgYAUTE1I0_DisplayLine
'        Else
'            intReturn = -1
'        End If
'
'    End If
'
'Loop Until intReturn <> 0
'Call lstErr_Clear(Me.lstErr, Me.cmdContext, "affichage mvt : " & fgYAUTE1I0.Rows - 1)

'fgYAUTE1I0.Visible = True
'fgYAUTE1I0_Sort1 = -1
'blnfgYAUTE1I0_Display = True

End Sub




Public Sub fgYBIAMVT0_Reset()
fgYBIAMVT0.Clear
fgYBIAMVT0_Sort1 = 0: fgYBIAMVT0_Sort2 = 0
fgYBIAMVT0_Sort1_Old = -1
fgYBIAMVT0_RowDisplay = 0: fgYBIAMVT0_RowClick = 0
fgYBIAMVT0_arrIndex = fgYBIAMVT0.Cols - 1
blnfgYBIAMVT0_DisplayLine = False
End Sub

Public Sub fgYAUTE1I0_Reset()
fgYAUTE1I0.Clear
fgYAUTE1I0_Sort1 = 0: fgYAUTE1I0_Sort2 = 0
fgYAUTE1I0_Sort1_Old = -1
fgYAUTE1I0_RowDisplay = 0: fgYAUTE1I0_RowClick = 0
fgYAUTE1I0.Cols = 9
fgYAUTE1I0_arrIndex = fgYAUTE1I0.Cols - 1
blnfgYAUTE1I0_DisplayLine = False
fgYAUTE1I0_RowSelect = 0: blnfgYAUTE1I0_Display = False
End Sub

Public Sub fgYAUTE1I0_Sort()
If fgYAUTE1I0.Rows > 1 Then
    fgYAUTE1I0.Row = 1
    fgYAUTE1I0.RowSel = fgYAUTE1I0.Rows - 1
    
    If fgYAUTE1I0_Sort1_Old = fgYAUTE1I0_Sort1 Then
        If fgYAUTE1I0_SortAD = 5 Then
            fgYAUTE1I0_SortAD = 6
        Else
            fgYAUTE1I0_SortAD = 5
        End If
    Else
        fgYAUTE1I0_SortAD = 5
    End If
    fgYAUTE1I0_Sort1_Old = fgYAUTE1I0_Sort1
    
    fgYAUTE1I0.Col = fgYAUTE1I0_Sort1
    fgYAUTE1I0.ColSel = fgYAUTE1I0_Sort2
    fgYAUTE1I0.Sort = fgYAUTE1I0_SortAD
End If
'cboDevise_Reset
End Sub

Public Sub fgYBIAMVT0_Sort()
If fgYBIAMVT0.Rows > 1 Then
    fgYBIAMVT0.Row = 1
    fgYBIAMVT0.RowSel = fgYBIAMVT0.Rows - 1
    
    If fgYBIAMVT0_Sort1_Old = fgYBIAMVT0_Sort1 Then
        If fgYBIAMVT0_SortAD = 5 Then
            fgYBIAMVT0_SortAD = 6
        Else
            fgYBIAMVT0_SortAD = 5
        End If
    Else
        fgYBIAMVT0_SortAD = 5
    End If
    fgYBIAMVT0_Sort1_Old = fgYBIAMVT0_Sort1
    
    fgYBIAMVT0.Col = fgYBIAMVT0_Sort1
    fgYBIAMVT0.ColSel = fgYBIAMVT0_Sort2
    fgYBIAMVT0.Sort = fgYBIAMVT0_SortAD
End If
'cboDevise_Reset
End Sub

Public Sub fgYBIAMVT0_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgYBIAMVT0.Rows - 1
    fgYBIAMVT0.Row = I
    fgYBIAMVT0.Col = lK
    X = Format$(Val(fgYBIAMVT0.Text), "0000000000000000")
    fgYBIAMVT0.Col = fgYBIAMVT0_arrIndex - 1
    fgYBIAMVT0.Text = X
Next I


fgYBIAMVT0_Sort1 = fgYBIAMVT0_arrIndex - 1: fgYBIAMVT0_Sort2 = fgYBIAMVT0_arrIndex - 1
fgYBIAMVT0_Sort




End Sub


Public Sub fgYAUTE1I0_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgYAUTE1I0.Rows - 1
    fgYAUTE1I0.Row = I
    fgYAUTE1I0.Col = lK
    X = Format$(Val(fgYAUTE1I0.Text), "0000000000000000")
    fgYAUTE1I0.Col = fgYAUTE1I0_arrIndex - 1
    fgYAUTE1I0.Text = X
Next I


fgYAUTE1I0_Sort1 = fgYAUTE1I0_arrIndex - 1: fgYAUTE1I0_Sort2 = fgYAUTE1I0_arrIndex - 1
fgYAUTE1I0_Sort




End Sub



Public Sub fgYBIAMVT0_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgYBIAMVT0.Row

If lRow > 0 And lRow < fgYBIAMVT0.Rows Then
    fgYBIAMVT0.Row = lRow
    For I = 0 To fgYBIAMVT0_arrIndex
        fgYBIAMVT0.Col = I: fgYBIAMVT0.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgYBIAMVT0.Row = mRow
    If fgYBIAMVT0.Row > 0 Then
        lRow = fgYBIAMVT0.Row
        lColor_Old = fgYBIAMVT0.CellBackColor
        For I = 0 To fgYBIAMVT0_arrIndex
          fgYBIAMVT0.Col = I: fgYBIAMVT0.CellBackColor = lColor
        Next I
        fgYBIAMVT0.Col = 0
    End If
End If

End Sub
Public Sub fgYAUTE1I0_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgYAUTE1I0.Row

If lRow > 0 And lRow < fgYAUTE1I0.Rows Then
    fgYAUTE1I0.Row = lRow
    For I = 0 To fgYAUTE1I0_arrIndex
        fgYAUTE1I0.Col = I: fgYAUTE1I0.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgYAUTE1I0.Row = mRow
    If fgYAUTE1I0.Row > 0 Then
        lRow = fgYAUTE1I0.Row
        lColor_Old = fgYAUTE1I0.CellBackColor
        For I = 0 To fgYAUTE1I0_arrIndex
          fgYAUTE1I0.Col = I: fgYAUTE1I0.CellBackColor = lColor
        Next I
        fgYAUTE1I0.Col = 0
    End If
End If

End Sub




Public Sub fgSelect_Display()
Dim I As Long, blnOk As Boolean
Dim Nb As Long
Dim xPCEC As String, lenPCEC As Integer
Dim xCLIENACAT As String, lenCLIENACAT As Integer
Dim xPRO As String, lenPRO As Integer
Dim xSelect As String, lenSelect As Integer
Dim xIntitulé As String, lenIntitulé As Integer
Dim K As Integer
Dim xCLIENARES As String, lenCLIENARES As Integer
Dim xCLIENARSD As String, lenCLIENARSD  As Integer
Dim xCOMPTECLA As Long, lenCOMPTECLA  As Integer
Dim wDate As Long
Dim blnSelect_Racine As Boolean, blnGroupe_Ok As Boolean

fraSelect_Options.Visible = False
prtSAB_Balance.blnPrint_Relevé_Total_Mvt = False
fraList.Visible = False

SSTab1.Tab = 0
mcboDevise = Trim(cboDevise.Text)
mcboPLANCOPRO = Trim(cboPLANCOPRO.Text)
xPCEC = Trim(cboPCEC): lenPCEC = Len(xPCEC)
xCLIENACAT = Trim(Mid$(cboSelect_CLIENACAT, 1, 3)): lenCLIENACAT = Len(xCLIENACAT)
xPRO = Trim(cboPLANCOPRO): lenPRO = Len(xPRO)
xSelect = Trim(txtCompte): lenSelect = Len(xSelect)

fraYAUTE1I0.Visible = False

blnSelect_Racine = False: blnGroupe_Filtre = False
If IsNumeric(xSelect) Then
    Select Case lenSelect
        Case 5: blnSelect_Racine = True: mAUTE1ICLI = Format(xSelect, "0000000")
        Case 4: blnGroupe_Filtre = True: Call fgSelect_Display_Groupe(xSelect)
    End Select
    End If

xIntitulé = Trim(txtIntitulé): lenIntitulé = Len(xIntitulé)
Call DTPicker_Amj7(txtSelect_SOLDEDMO, mAmj_SOLDEDMO)
Call DTPicker_Amj7(txtSelect_COMPTEOUV, mAmj_COMPTEOUV)
Call DTPicker_Amj7(txtSelect_COMPTECLO, mAmj_COMPTECLO)
xCLIENARES = Trim(txtSelect_CLIENARES): lenCLIENARES = Len(xCLIENARES)
xCLIENARSD = Trim(txtSelect_CLIENARSD): lenCLIENARSD = Len(xCLIENARSD)
xCOMPTECLA = Val(Trim(txtSelect_COMPTECLA)): lenCOMPTECLA = Len(Trim(txtSelect_COMPTECLA))
If chkSelect = "1" Then
    cboDevise.ListIndex = 0
    cboPCEC.ListIndex = 0
    cboSelect_CLIENACAT.ListIndex = 0
    cboPLANCOPRO.ListIndex = 0
    xPCEC = "": lenPCEC = 0
    xCLIENACAT = "": lenCLIENACAT = 0
    xPRO = "": lenPRO = 0
    xCLIENARES = "": lenCLIENARES = 0
    xCLIENARSD = "": lenCLIENARSD = 0
    xCOMPTECLA = 0: lenCOMPTECLA = 0
    mcboDevise = ""
    txtCompte = "": lenSelect = 0
    txtIntitulé = "": lenIntitulé = 0
Else

    If xPCEC = "" _
    And mcboSelect_CLIENACAT = "" _
    And mcboPLANCOPRO = "" _
    And mcboDevise = "" _
    And lenSelect = 0 _
    And lenCLIENARES = 0 _
    And lenCLIENARSD = 0 _
    And lenCOMPTECLA = 0 _
    And lenIntitulé = 0 Then
'        Exit Sub
    End If
End If

If chkSelect_SOLDEDMO = "1" And chkSelect_DORCPTDMV = "1" Then
    Call ZDORCPT_SQL_ODBC
End If

fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Visible = False
fgSelect.Row = 0
fgSelect.Col = 1: fgSelect.CellAlignment = 1


For I = 1 To marrYBIACPT0_Nb

    blnOk = True
    xYBIACPT0 = marrYBIACPT0(I)
    Select Case xYBIACPT0.SOLDECEN
        Case Is = 0: If chkSelect_SoldeZ = "1" And xYBIACPT0.SOLDEDMO < YBIATAB0_DIBM_CPT_JP1 Then blnOk = False
        Case Is > 0: If chkSelect_SoldeDb = "1" Then blnOk = False
        Case Is < 0: If chkSelect_SoldeCr = "1" Then blnOk = False
    End Select
    
    If blnOk And chkSelect_Annulé = "1" Then
            If xYBIACPT0.COMPTEFON = "4" Then
                blnOk = False
            End If
    End If
    '_______________________________________________________________________________
    If blnOk Then
        If blnGroupe_Filtre Then
            blnGroupe_Ok = False
            For K = 1 To arrGroupe_Filtre_Nb
                If xYBIACPT0.CLIENACLI = arrGroupe_Filtre(K) Then blnGroupe_Ok = True: Exit For
            Next K
            blnOk = blnGroupe_Ok
        Else
            If lenSelect > 0 Then
                If blnSelect_Racine Then
                    K = InStr(1, xYBIACPT0.CLIENACLI, xSelect)
                Else
                    K = InStr(1, xYBIACPT0.COMPTECOM, xSelect)
                End If
                If K = 0 Then
                    blnOk = False
                End If
            End If
        End If
    End If
    '_______________________________________________________________________________
    If blnOk And lenPCEC > 0 Then
            If Mid$(xYBIACPT0.COMPTEOBL, 1, lenPCEC) <> xPCEC Then
                blnOk = False
            End If
    End If
    
    If blnOk And lenCLIENACAT > 0 Then
            If Mid$(xYBIACPT0.CLIENACAT, 1, lenCLIENACAT) <> xCLIENACAT Then
                blnOk = False
            End If
    End If
    
    If blnOk And chkSelect_HB = "1" Then
            If Mid$(xYBIACPT0.COMPTEOBL, 1, 1) = "9" Then
                blnOk = False
            End If
    End If
    
    If blnOk And lenPRO > 0 Then
            If Mid$(xYBIACPT0.PLANCOPRO, 1, lenPRO) <> xPRO Then
                blnOk = False
            End If
    End If
    If blnOk And mcboDevise <> "" Then
        If mcboDevise <> Trim(xYBIACPT0.COMPTEDEV) Then
            blnOk = False
        End If
    End If


    If blnOk And lenIntitulé > 0 Then
            K = InStr(1, xYBIACPT0.CLIENASIG, xIntitulé)    ' COMPTEINT
            If K = 0 Then
                blnOk = False
            End If
    End If
    
    If blnOk And chkSelect_SOLDEDMO = "1" Then
        If chkSelect_DORCPTDMV = "1" Then
            wDate = wDORCPTDMV(I)
            If wDate = 0 Then
                wDate = xYBIACPT0.SOLDEDMO   ' !!! tous les comptes ne sont pas gérés par l'application COMPTES DORMANTS
            End If
        Else
            wDate = xYBIACPT0.SOLDEDMO
        End If
        
        If wDate > mAmj_SOLDEDMO Then
            If optSelect_SOLDEDMO_Inf Then
                blnOk = False
            End If
        Else
            If optSelect_SOLDEDMO_Sup Then
                blnOk = False
            End If
        End If
        
    End If
    
    If blnOk And chkSelect_COMPTEOUV = "1" Then
        If xYBIACPT0.COMPTEOUV < mAmj_COMPTEOUV Then
            blnOk = False
        End If
    End If
    
    If blnOk And chkSelect_COMPTECLO = "1" Then
        If xYBIACPT0.COMPTECLO < mAmj_COMPTECLO Then
            blnOk = False
        End If
    End If
    
    If blnOk And lenCLIENARES > 0 Then
            If Mid$(xYBIACPT0.CLIENARES, 1, lenCLIENARES) <> xCLIENARES Then
                blnOk = False
            End If
    End If
    If blnOk And blnSelect_Pays Then
            If xYBIACPT0.CLIENARSD = "   " Then
                blnOk = False
            End If
    End If
    
    If blnOk And lenCLIENARSD > 0 Then
            If Mid$(xYBIACPT0.CLIENARSD, 1, lenCLIENARSD) <> xCLIENARSD Then
                blnOk = False
            End If
    End If
    
    If blnOk And lenCOMPTECLA > 0 Then
            If xYBIACPT0.COMPTECLA <> xCOMPTECLA Then
                blnOk = False
            End If
    End If

    If blnOk Then
        
        If blnMesComptes Then
            If xYBIACPT0.COMPTECLA >= 60 Then fgSelect_DisplayLine I
       Else
            If fctUser_Classe_Aut(xYBIACPT0.COMPTECLA) Then fgSelect_DisplayLine I
        End If
    End If
    
Next I


blnMesComptes = False

Call lstErr_Clear(Me.lstErr, Me.cmdContext, "affichage : " & fgSelect.Rows - 1 & " / " & marrYBIACPT0_Nb)
fgSelect_Sort1 = -1
fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
fgSelect.Visible = True


End Sub

Private Sub fgYBIAMVT0_Display(lCOMPTECOM As String)
Dim xSQL As String
Dim wIBMAMJMax As String
'On Error GoTo Error_Handle


txtAMJ_Control

fgYBIAMVT0_Reset

fgYBIAMVT0.Rows = 1
fgYBIAMVT0.FormatString = fgYBIAMVT0_FormatString
fgYBIAMVT0.Visible = False


xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
     & " where MOUVEMCOM = '" & lCOMPTECOM & "'" _
     & " and MOUVEMDTR >= " & dateIBM(xAmjMin) _
     & " and MOUVEMDTR <= " & dateIBM(xAmjMax) _
     & " order by MOUVEMDTR,MOUVEMPIE,MOUVEMECR"
     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    Call rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVT0)
    fgYBIAMVT0_DisplayLine
    rsSab.MoveNext
Loop


Call lstErr_Clear(Me.lstErr, Me.cmdContext, "affichage mvt : " & fgYBIAMVT0.Rows - 1)

fgYBIAMVT0.Visible = True
fgYBIAMVT0_Sort1 = -1
End Sub


Public Sub fgSelect_DisplayLine(lIndex As Long)
Dim X As String, xCOMPTECLO As String
Dim wForecolor As Long

On Error Resume Next
           Select Case xYBIACPT0.COMPTEFON
                Case 0: wForecolor = RGB(0, 0, 160): xCOMPTECLO = ""
                Case 4: wForecolor = vbRed: xCOMPTECLO = xYBIACPT0.COMPTEFON & " " & dateIBM10(xYBIACPT0.COMPTECLO, True)
                Case Else: wForecolor = vbMagenta
                            If xYBIACPT0.COMPTEBLO = 0 Then
                                xCOMPTECLO = xYBIACPT0.COMPTEFON
                            Else
                                xCOMPTECLO = xYBIACPT0.COMPTEFON & " " & dateIBM10(xYBIACPT0.COMPTEBLO, True)
                            End If

            End Select
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect.Col = 0: fgSelect.Text = xYBIACPT0.CLIENACLI & "_" & xYBIACPT0.COMPTEINT
            fgSelect.CellForeColor = wForecolor
            fgSelect.Col = 2: fgSelect.Text = xYBIACPT0.COMPTEDEV
            If xYBIACPT0.COMPTEDEV = "EUR" Then
                fgSelect.Text = " €"
            '    fgSelect.CellForeColor = RGB(0, 96, 0) 'wForecolor
            'Else
            '    fgSelect.CellForeColor = RGB(0, 0, 64) 'wForecolor
            End If
            fgSelect.CellForeColor = RGB(0, 96, 0)
            
            fgSelect.Col = 3: fgSelect.Text = xYBIACPT0.CLIENARES & " " & xYBIACPT0.COMREFCOR
            fgSelect.CellForeColor = vbBlue ' wForecolor
            fgSelect.Col = 4: fgSelect.Text = Trim(xYBIACPT0.COMPTEOBL) & " " & xYBIACPT0.PLANCOPRO
            fgSelect.CellForeColor = wForecolor
            
            fgSelect.Col = 5: fgSelect.Text = " " & xYBIACPT0.COMPTECOM
            fgSelect.CellForeColor = RGB(0, 96, 0)
            
            fgSelect.Col = 6: fgSelect.Text = dateIBM10(xYBIACPT0.SOLDEDMO, True)
            fgSelect.CellForeColor = wForecolor
            fgSelect.Col = 7: fgSelect.Text = xCOMPTECLO
            fgSelect.CellForeColor = wForecolor

             fgSelect.Col = 8: fgSelect.Text = dateIBM10(xYBIACPT0.COMPTEOUV, True)
            fgSelect.CellForeColor = wForecolor
            fgSelect.Col = 9: fgSelect.Text = xYBIACPT0.CLIENARSD
            fgSelect.CellForeColor = wForecolor
           fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
            
             fgSelect.Col = 1: fgSelect.Text = Format$(Abs(xYBIACPT0.SOLDECEN), "### ### ### ###.00 ")
           If xYBIACPT0.SOLDECEN > 0 Then
                fgSelect.CellForeColor = &HFF&       'vbRed
            Else
                fgSelect.CellForeColor = vbBlue
            End If
End Sub


Public Sub fgYBIAMVT0_DisplayLine()
Dim X As String
On Error Resume Next
fgYBIAMVT0.Rows = fgYBIAMVT0.Rows + 1
fgYBIAMVT0.Row = fgYBIAMVT0.Rows - 1
fgYBIAMVT0.Col = 0: fgYBIAMVT0.Text = dateIBM10(xYBIAMVT0.MOUVEMDTR, True)
fgYBIAMVT0.Col = 1: fgYBIAMVT0.Text = dateIBM10(xYBIAMVT0.MOUVEMDVA, True)
fgYBIAMVT0.Col = 2: fgYBIAMVT0.Text = xYBIAMVT0.MOUVEMOPE & " " & Format(xYBIAMVT0.MOUVEMNUM, "@@@@@@@@@@") & " " & xYBIAMVT0.MOUVEMEVE
fgYBIAMVT0.Col = 4: fgYBIAMVT0.Text = Trim(xYBIAMVT0.LIBELLIB1) & " " & Trim(xYBIAMVT0.LIBELLIB2) & " " & Trim(xYBIAMVT0.LIBELLIB3) & " " & Trim(xYBIAMVT0.LIBELLIB4)
fgYBIAMVT0.Col = fgYBIAMVT0_arrIndex: fgYBIAMVT0.Text = xYBIAMVT0.MOUVEMPIE & " " & xYBIAMVT0.MOUVEMECR
 
fgYBIAMVT0.Col = 3: fgYBIAMVT0.Text = Format$(Abs(xYBIAMVT0.MOUVEMMON), "### ### ### ###.00 ")
If xYBIAMVT0.MOUVEMMON > 0 Then
     fgYBIAMVT0.CellForeColor = vbRed
 Else
     fgYBIAMVT0.CellForeColor = vbBlue
 End If

End Sub
Public Sub fgYAUTE1I0_DisplayLine()
Dim X As String, I As Integer
Dim wCellBackColor As Long, wCellForeColor As Long, AUTE1INIV_Space As String
Dim wCol As Integer
On Error Resume Next
fgYAUTE1I0.Rows = fgYAUTE1I0.Rows + 1
fgYAUTE1I0.Row = fgYAUTE1I0.Rows - 1
AUTE1INIV_Space = ""

If mAUTE1ICLI_Aut <> rsSab("AUTE1ICLI") Then
    For I = 0 To 7 'fgYAUTE1I0_arrIndex
        fgYAUTE1I0.Col = I
            fgYAUTE1I0.CellBackColor = mColor_G1
            
        fgYAUTE1I0.CellForeColor = vbBlue
    Next I
    fgYAUTE1I0.Col = 0: fgYAUTE1I0.Text = rsSab("AUTE1IRES"): fgYAUTE1I0.CellFontBold = True
    fgYAUTE1I0.Col = 1
    X = Trim(rsSab("AUTE1IGRP"))
    If X <> "" Then
        X = X & " - "
        fgYAUTE1I0.CellBackColor = mColor_Y2
    Else
        fgYAUTE1I0.CellBackColor = mColor_G1
    End If
    fgYAUTE1I0.Col = 1: fgYAUTE1I0.Text = X & rsSab("AUTE1ICLI") & "  " & Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
    fgYAUTE1I0.CellFontSize = 8: fgYAUTE1I0.CellFontBold = True

    fgYAUTE1I0.Rows = fgYAUTE1I0.Rows + 1
    fgYAUTE1I0.Row = fgYAUTE1I0.Rows - 1
End If

Select Case rsSab("AUTE1INIV")
    Case 0: wCol = 1
                wCellBackColor = &HFFFFFF
            
    Case 1: wCol = 1
                
                    wCellBackColor = mColor_G0
                    fgYAUTE1I0.Col = 1: fgYAUTE1I0.CellFontBold = True
    Case 2: wCol = 1
                wCellBackColor = &HFFFFFF: AUTE1INIV_Space = " ... "
    Case Else: wCol = 1
                wCellBackColor = &HFFFFFF: AUTE1INIV_Space = "  ... ... "
End Select
'If rsSab("AUTE1IDEP") <> " " Then
 '    wCellForeColor = vbRed
 'Else
     wCellForeColor = vbBlack
 'End If

fgYAUTE1I0.Col = wCol: fgYAUTE1I0.Text = AUTE1INIV_Space & rsSab("AUTE1IAUT")
fgYAUTE1I0.Col = 2: fgYAUTE1I0.Text = rsSab("AUTE1IDEV")
If rsSab("AUTE1IMDB") <> 0 Then fgYAUTE1I0.Col = 3: fgYAUTE1I0.Text = Format$(Abs(rsSab("AUTE1IMDB")), "### ### ### ###.00 ")
If rsSab("AUTE1IMCR") <> 0 Then fgYAUTE1I0.Col = 4: fgYAUTE1I0.Text = Format$(Abs(rsSab("AUTE1IMCR")), "### ### ### ###.00 ")
'If rsSab("AUTE1IMTD") <> 0 Then fgYAUTE1I0.Col = 5: fgYAUTE1I0.Text = Format$(Abs(rsSab("AUTE1IMTD")), "### ### ### ###.00 ")
If rsSab("AUTE1IMAU") <> 0 Then fgYAUTE1I0.Col = 6: fgYAUTE1I0.Text = Format$(Abs(rsSab("AUTE1IMAU")), "### ### ### ###.00 ")


For I = 0 To 7 'fgYAUTE1I0_arrIndex
    fgYAUTE1I0.Col = I
    If I = 3 Then
        fgYAUTE1I0.CellForeColor = vbRed
    Else
        fgYAUTE1I0.CellForeColor = wCellForeColor
    End If
    
    fgYAUTE1I0.CellBackColor = wCellBackColor
Next I

If rsSab("AUTE1IMTD") <> 0 Then
    fgYAUTE1I0.Col = 5
    fgYAUTE1I0.Text = Format$(Abs(rsSab("AUTE1IMTD")), "### ### ### ###.00 ")
    If rsSab("AUTE1IMTD") > 0 Then
        fgYAUTE1I0.CellBackColor = mColor_W1
        fgYAUTE1I0.Col = wCol: fgYAUTE1I0.CellBackColor = mColor_W1
    End If
End If

If rsSab("AUTE1IDAF") <> 0 Then
    fgYAUTE1I0.Col = 7
    X = dateImp10_S(rsSab("AUTE1IDAF") + 19000000)
    fgYAUTE1I0.Text = X
    Select Case rsSab("AUTE1IDAF") + 19000000
        Case Is > mAUTE1IDAF_Max8
        Case Is <= DSys: fgYAUTE1I0.CellBackColor = mColor_W1
        Case Else: fgYAUTE1I0.CellBackColor = RGB(255, 255, 64)
    End Select
End If

mAUTE1ICLI_Aut = rsSab("AUTE1ICLI")

End Sub


Public Sub fgSelect_Sort()
fgSelect.Visible = False
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
fgSelect.Visible = True
End Sub
Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
Dim mK As Integer
Dim wIndex As Long
mK = lK
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    Select Case lK
        Case 0: '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ TRI Racine / PCEC
                fgSelect.Col = fgSelect_arrIndex
                wIndex = Val(fgSelect.Text)
                X = marrYBIACPT0(wIndex).CLIENACLI & marrYBIACPT0(wIndex).COMPTEOBL & marrYBIACPT0(wIndex).COMPTECOM
                fgSelect.Col = fgSelect_arrIndex - 1
                fgSelect.Text = X
        Case 1: fgSelect.Col = 1: X = fgSelect.Text
                fgSelect.Col = fgSelect_arrIndex - 1
                fgSelect.Text = Format$(Val(X), "000000000000.00")
        Case 2: '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ TRI POUR IMPRESSION BALANCE DEV / PCEC / COMPTE
                fgSelect.Col = fgSelect_arrIndex
                wIndex = Val(fgSelect.Text)
                X = marrYBIACPT0(wIndex).COMPTEDEV & marrYBIACPT0(wIndex).COMPTEOBL & marrYBIACPT0(wIndex).COMPTECOM
                fgSelect.Col = fgSelect_arrIndex - 1
                fgSelect.Text = X
        Case 5: fgSelect.Col = 5: X = Trim(fgSelect.Text)
                fgSelect.Col = fgSelect_arrIndex - 1
                fgSelect.Text = Mid$(X, 10, 1) & X          ' code résidence pos=10
        Case 6:
                fgSelect.Col = fgSelect_arrIndex
                X = marrYBIACPT0(Val(fgSelect.Text)).SOLDEDMO
                fgSelect.Col = fgSelect_arrIndex - 1
                fgSelect.Text = X
        Case 8:
                fgSelect.Col = fgSelect_arrIndex
                X = marrYBIACPT0(Val(fgSelect.Text)).COMPTEOUV
                fgSelect.Col = fgSelect_arrIndex - 1
                fgSelect.Text = X
        Case 9: '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ TRI POUR IMPRESSION BALANCE PAYS /DEV / PCEC / COMPTE
                fgSelect.Col = fgSelect_arrIndex
                wIndex = Val(fgSelect.Text)
                X = marrYBIACPT0(wIndex).CLIENARSD & marrYBIACPT0(wIndex).COMPTEOBL & marrYBIACPT0(wIndex).COMPTEDEV & marrYBIACPT0(wIndex).COMPTECOM
                fgSelect.Col = fgSelect_arrIndex - 1
                fgSelect.Text = X
        Case 9000: '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ TRI POUR IMPRESSION stat client catégorie client / racine
                fgSelect.Col = fgSelect_arrIndex
                wIndex = Val(fgSelect.Text)
                X = marrYBIACPT0(wIndex).PLANCOPRO & marrYBIACPT0(wIndex).CLIENACAT & marrYBIACPT0(wIndex).CLIENACLI
                fgSelect.Col = fgSelect_arrIndex - 1
                fgSelect.Text = X
    End Select
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
fgSelect_Sort1 = mK: fgSelect_Sort2 = mK
End Sub



'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Private Sub mnuAuto_Balance_Stock_Click_xlsManual()

blnBalance_Service_Stock = True
optBalance_YSOLDE0_J = True
Call cmdBalance_Service_xlsManual
Me.Show

End Sub

Private Sub mnuAuto_Groupe_6000_Click_xlsManual()
Me.Enabled = False: Me.MousePointer = vbHourglass

fraSelect_Clear
chkSelect_SoldeZ = "0"
txtCompte = "6000"
fgSelect_Display
fgSelect_SortX 0

Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtAmjMin, YBIATAB0_DATE_CPT_JP0)
txtAMJ_Control

prtSAB_Balance.blnPrint_Relevé_Total_Mvt = False
Call mnuSelect_Print_Relevé_Click_xlsManual

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuAuto_RELEVE_FOTC()
Me.Enabled = False: Me.MousePointer = vbHourglass
fraSelect_Clear
chkSelect_SoldeZ = "1"
cboPLANCOPRO = "PO"
fgSelect_Display
fgSelect_SortX 2

Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtAmjMin, YBIATAB0_DATE_CPT_JP0)
txtAMJ_Control

prtSAB_Balance.blnPrint_Relevé_Total_Mvt = True
'mnuSelect_Print_Relevé_Click ========================================================
Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Relevés : " & fgSelect.Rows - 1)

Call prtSAB_Balance_Monitor_RELEVE_FOTC(xAmjMin, xAmjMax, fgSelect, marrYBIACPT0(), marrYBIACPT0_Nb, "", wYSTOMON(), wDORCPTDMV())
fgSelect.Visible = True
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_PCI_DC_Click_xlsManual()
Dim xFct As String, xR As String
Dim xZCOMREF0 As typeZCOMREF0, xSens As String
Dim xPCI As String
Dim xSQL As String
Dim nbErr_PCI As Long, nbErr_Sens As Long
Dim prtTitleText As String
Dim currentSheet As Long
Dim currentRow As Long
Dim comptageRows As Long
Dim maxRows As Long

Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression contrôle PCI : " & marrYBIACPT0_Nb)
'________________________________________________________________
maxRows = 37
'On recopie le classeur modèle de c:\BIASRV vers c:\temp\imp_pdf
FileCopy paramFolder_Local & "\Modeles\modele_BAL_PCI_DC.xlsx", paramIMP_PDF_Path_Temp & "\modele_BAL_PCI_DC.xlsx"
'on charge CE classeur dans Excel
Call init_xlsManual
Call appExcelPublic.Workbooks.Open(paramIMP_PDF_Path_Temp & "\modele_BAL_PCI_DC.xlsx")
Set wbExcel = appExcelPublic.ActiveWorkbook
With wbExcel
    .Title = .Sheets(1).Name
    .Subject = .Sheets(1).Name
End With
currentSheet = 1
currentRow = 1
prtTitleText = "Comptabilité : Etat des anomalies de sens des comptes / PCI " & " - au : " & dateImp(YBIATAB0_DATE_CPT_J)
wbExcel.Sheets(currentSheet).Cells(currentRow, 5) = prtTitleText
currentRow = 8
comptageRows = 8
'________________________________________________________________
xZCOMREF0.COMREFETA = 1
xZCOMREF0.COMREFPLA = 1
xZCOMREF0.COMREFCOM = ""
xZCOMREF0.COMREFCOR = "DC"

nbErr_PCI = 0: nbErr_Sens = 0

For I = 1 To marrYBIACPT0_Nb
    If marrYBIACPT0(I).SOLDECEN <> 0 Then
        xPCI = Mid$(marrYBIACPT0(I).COMPTEOBL, 1, 6)
        If xPCI <> Mid$(xZCOMREF0.COMREFCOM, 1, 6) Then
            xZCOMREF0.COMREFCOM = xPCI
            xSQL = "select * from " & paramIBM_Library_SAB & ".ZCOMREF0 " _
                & " where COMREFCOM = '" & xZCOMREF0.COMREFCOM & "'" _
                & " and COMREFCOR = '" & xZCOMREF0.COMREFCOR & "'" _
                & " and COMREFETA = " & xZCOMREF0.COMREFETA _
                & " and COMREFPLA = " & xZCOMREF0.COMREFPLA
     
            Set rsSab = cnsab.Execute(xSQL)
            If rsSab.EOF Then
                nbErr_PCI = nbErr_PCI + 1
                Call prtBalance_PCI_DC_Line_xlsManual(marrYBIACPT0(I), "???", currentRow, wbExcel.Sheets(1), comptageRows, maxRows)
                xSens = "N"
            Else
                xSens = Mid$(rsSab("COMREFREF"), 1, 1)
            End If
        End If
        
        If xSens = "D" And marrYBIACPT0(I).SOLDECEN < 0 Then
            nbErr_Sens = nbErr_Sens + 1: Call prtBalance_PCI_DC_Line_xlsManual(marrYBIACPT0(I), xSens, currentRow, wbExcel.Sheets(1), comptageRows, maxRows)
        End If
        If xSens = "C" And marrYBIACPT0(I).SOLDECEN > 0 Then
            nbErr_Sens = nbErr_Sens + 1: Call prtBalance_PCI_DC_Line_xlsManual(marrYBIACPT0(I), xSens, currentRow, wbExcel.Sheets(1), comptageRows, maxRows)
        End If
       
                

    End If
    
Next I

Call prtBalance_PCI_DC_Close_xlsManual(marrYBIACPT0_Nb, nbErr_PCI, nbErr_Sens, currentRow, wbExcel.Sheets(1), comptageRows, maxRows)
'on supprime les 5 lignes modèles
Rows("4:8").Select
Selection.Delete
currentRow = currentRow - 5
Call frmSAB_Balance.zoneImpression_xlsManual(wbExcel.Sheets(currentSheet).Name, currentRow, wbExcel.Sheets(currentSheet))
Call ActiveSheet.ExportAsFixedFormat(xlTypePDF, paramIMP_PDF_Path_Temp & "\" & paramEditionNoPaper_Auto_PgmName & ".pdf")
Call impressions_xlsManual.prtIMP_PDF_Monitor_xlsManual
Call wbExcel.Close(True)
Set wbExcel = Nothing
Kill paramIMP_PDF_Path_Temp & "\modele_BAL_PCI_DC.xlsx"
End Sub

Private Sub mnuSelect_Print_Relevé_Click_xlsManual()
Dim wsexcelRien As Excel.Worksheet
Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Relevés : " & fgSelect.Rows - 1)

Call prtSAB_Balance_Monitor_xlsManual("R", xAmjMin, xAmjMax, fgSelect, marrYBIACPT0(), marrYBIACPT0_Nb, "", wYSTOMON(), wDORCPTDMV(), -1, wsexcelRien)
fgSelect.Visible = True
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub zoneImpression_xlsManual(lFct As String, nbRows As Long, wsheet As Excel.Worksheet)

    Call init_TypePagesetup
    If nbRows > 0 Then
        If Trim(lFct) = "BAL6000" Then
            zoneImpressionPagesetup.Zoom = 90
            wsheet.Range("A1:J" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$J$" & CStr(nbRows)
            zoneImpressionPagesetup.LeftFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "prtSAB_Balance   &D &T  BIA_INFO"
        ElseIf Trim(lFct) = "BAL_B_HB" Then
            zoneImpressionPagesetup.Zoom = 100
            wsheet.Range("A1:J" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$J$" & CStr(nbRows)
            zoneImpressionPagesetup.LeftFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "prtSAB_Balance   &D &T  BIA_INFO"
        ElseIf Trim(lFct) = "BAL_PCI_DC" Then
            zoneImpressionPagesetup.Zoom = 90
            wsheet.Range("A1:I" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$I$" & CStr(nbRows)
            zoneImpressionPagesetup.LeftFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "prtBalance_PCI_DC   &D &T  BIA_INFO"
        ElseIf Trim(lFct) = "BALANCE_Cumul" Then
            zoneImpressionPagesetup.Zoom = 100
            wsheet.Range("A1:G" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$G$" & CStr(nbRows)
            zoneImpressionPagesetup.LeftFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "prtBalance_Cumul   &D &T  BIA_INFO"
        ElseIf Trim(lFct) = "BALANCE_Stock" Then
            zoneImpressionPagesetup.Zoom = 85
            wsheet.Range("A1:L" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$L$" & CStr(nbRows)
            zoneImpressionPagesetup.LeftFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "prtSAB_Compta   &D &T  BIA_INFO"
        End If
    End If
    wsheet.Activate
    zoneImpressionPagesetup.RightFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "&P"
    zoneImpressionPagesetup.Orientation = xlLandscape
    Call SetTypePageSetup(wsheet)

End Sub

Public Sub Msg_Rcv(Msg As String)

mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init("SAB_Balance", SAB_Balance_Aut)

Form_Init
If SAB_Balance_Aut.Rapprocher Then DS_Server_Open

Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@RELEVE_FOTC":
                    blnAuto = True
                    lstService.ListIndex = 0
                    chkSelect = 1
                    Call mnuAuto_RELEVE_FOTC
    
    Case "@BAL_6000":
                    blnAuto = True
                    Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S42", "BIA-BAL-6000", "Archive")
                    lstService.ListIndex = 0
                    chkSelect = 1
                    If xlsManual Then
                        Call mnuAuto_Groupe_6000_Click_xlsManual
                    Else
                        mnuAuto_Groupe_6000_Click
                    End If
    Case "@BAL_B/HB"
                Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S60", "BIA-BAL-B-HB", "Archive")
                blnAuto = True
                SSTab1.Tab = 0
                lstService.ListIndex = 0
                chkSelect = 1
                mnuSelect_Print_Balance_Click
                chkBalance_Détail = "0"
                chkBalance_Récap = "0"
                chkBalance_Récap_Bilan = "1"
                If xlsManual Then
                    Call cmdBalance_Ok_Click_xlsManual
                Else
                    Call cmdBalance_Ok_Click
                End If
                Printer_Reset
    Case "@BAL_PCI_DC"
                Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S60", "BIA-BAL-PCI-DC", "Archive")
                blnAuto = True
                SSTab1.Tab = 0
                lstService.ListIndex = 0
                chkSelect = 1
                If xlsManual Then
                    Call mnuSelect_Print_PCI_DC_Click_xlsManual
                Else
                    Call mnuSelect_Print_PCI_DC_Click
                End If
    Case "@RCOM_AUT": blnAuto = True
                    'If Not IsEmpty(XPrt) Then Set XPrt_Previous = XPrt
                    'Printer_PDF
                    lstService.ListIndex = 0
                    'chkSelect = 1
                    SSTab1.Tab = 2
                    cmdRCOM_AUT
                    'If Not IsEmpty(Xprt_Previous) Then Set XPrt = Xprt_Previous
                    'cmdSendMail prtIMP_PDF_FileName
                    
    Case "@CPT_OD": blnAuto = True
                    Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_J)
                    Call DTPicker_Set(txtAmjMin, YBIATAB0_DATE_CPT_JP0)
                    lstService.ListIndex = 0
                    SSTab1.Tab = 2
                    cmdCPT_OD
    Case "@ENG_BEA_LFB": blnAuto = True
                    Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_J)
                    Call DTPicker_Set(txtAmjMin, YBIATAB0_DATE_CPT_JP0)
                    lstService.ListIndex = 0
                    SSTab1.Tab = 0
                    mnuSelect_ENG_Detail_Auto
'$JPL 2014-11-12 Balance stock automatique en fin de mois
    Case "@BAL_STOCK":
                    blnAuto = True
                    lstService.ListIndex = 0
                    chkSelect = 1
                    If xlsManual Then
                        Call mnuAuto_Balance_Stock_Click_xlsManual
                    Else
                        mnuAuto_Balance_Stock_Click
                    End If
    Case Else: blnAuto = False
End Select
If blnAuto Then
    Unload Me
End If
End Sub


Public Sub Form_Init()
Me.Enabled = False:: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True
SSTab1.Visible = False

If Not SAB_Balance_Aut.Xspécial Then
    chkSelect_Résidence.Visible = "2"
    chkSelect_MOUVEMDCO.Visible = "2"
End If


If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistant", vbCritical, "frmSAB_YBIACPT0.param_init"
    fraTab0.Enabled = False
End If

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fraMvt.Visible = True 'SAB_Balance_Aut.Xspécial
fgYAUTE1I0_FormatString = fgYAUTE1I0.FormatString
fgYBIASTO0_FormatString = fgYBIASTO0.FormatString
cmdReset
''SSTab1.Visible = true

fraList.Visible = False
fraList.Top = 1590
fraList.Left = 5655

fraList.Height = 6900
fgList_FormatString = fgList.FormatString


lstService.Top = 0
lstService.Height = 6100
lstService.Width = 3900
lstService.Visible = True
lstService.AddItem "G* Tous les comptes"
lstService.AddItem "G? Comptes sans affectation"

Call lst_LoadK2("SAb_Param", "Compte_Unit", lstService, True)

lstErr_Clear lstErr, cmdContext, "<== SELECTIONNER un SERVICE"

lblYAUTE1I0 = "Position Comptable / Autorisations au " & dateImp10_S(YBIATAB0_DATE_CPT_J)
lblYAUTE1I0.ForeColor = vbBlue
'========================================================================
mnuSelect_CPT_OD.Visible = SAB_Balance_Aut.X10
'========================================================================


Me.Enabled = True:: Me.MousePointer = 0

End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
Dim X As String
On Error Resume Next

blnControl = False
blnError = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
currentAction = ""

lstService.Top = 1080
lstService.Left = 120

fraBalance.Top = 480
fraBalance.Left = 2400

fraSelect_Options.Top = 2640
fraSelect_Options.Left = 120
fraSelect_Options.Visible = False
fraSelect_Options.BackColor = &HF0FFFF '
usrColor_Container fraSelect_Options, fraSelect_Options.BackColor
fraSelect_SOLDEDMO.BackColor = fraSelect_Options.BackColor
fraSelect_Clear
mcboPCEC = ""
mcboSelect_CLIENACAT = ""
mcboDevise = ""

fgSelect.Font = prtFontName_CourierNew
fgYBIAMVT0.Font = prtFontName_CourierNew
fgYBIASTO0.Font = prtFontName_CourierNew
picYAUTE1I0.Enabled = False

mAmjmax = YBIATAB0_DATE_CPT_J: Call DTPicker_Set(txtAmjMax, mAmjmax)
mAmjMin = Mid$(YBIATAB0_DATE_CPT_J, 1, 6) & "01": Call DTPicker_Set(txtAmjMin, mAmjMin)
mAmjMin = Mid$(YBIATAB0_DATE_CPT_MP1, 1, 6) & "01"
chkSelect_SOLDEDMO = "0"
fraSelect_SOLDEDMO.Visible = False
optSelect_SOLDEDMO_Inf = True
Call DTPicker_Set(txtSelect_SOLDEDMO, mAmjmax)

chkSelect_COMPTEOUV = "0"
txtSelect_COMPTEOUV.Visible = False
Call DTPicker_Set(txtSelect_COMPTEOUV, mAmjmax)
chkSelect_COMPTECLO = "0"
txtSelect_COMPTECLO.Visible = False
Call DTPicker_Set(txtSelect_COMPTECLO, mAmjmax)

optBalance_YSOLDE0_J.Caption = dateImp(YBIATAB0_DATE_CPT_J)
optBalance_YSOLDE0_MP1.Caption = dateImp(YBIATAB0_DATE_CPT_MP1)
optBalance_YSOLDE0_MP2.Caption = dateImp(YBIATAB0_DATE_CPT_MP2)
optBalance_YSOLDE0_AP1.Caption = dateImp(YBIATAB0_DATE_CPT_AP1)
optBalance_YSOLDE0_J.Value = True

chkBalance_Détail = "1"
chkBalance_Récap = "0"
chkBalance_Compte_Soldé = "1"

chkBalance_Print(7) = "1"
chkBalance_Print_FontBold(7) = "1"

txtBalance_Print_Trame(7) = 20

chkBalance_Print(1) = "1"
txtBalance_Print_Trame(1) = 20

chkBalance_Print(0) = "1"
chkBalance_Print_FontBold(0) = "1"
txtBalance_Print_Trame(0) = 40

chkBalance_Print(6) = "1"

txtBalance_CSV_Folder = "U:\"
txtBalance_CSV_FileName = "YBALANCE.csv"

mnuAuto_Balance_Stock.Enabled = SAB_Balance_Aut.Xspécial
ReDim wYSTOMON(1)
ReDim wDORCPTDMV(1)
blnService_Printer = False

fraList.Visible = False

blnControl = True

End Sub



Public Function param_Init()
Dim I As Long

marrYBIACPT0_Nb = 0
ReDim marrYBIACPT0(1000) 'paramYBIACPT0_Nb + 10)
param_Init = Null


mENG_LFB = "0011012;0011084;0011004;0011005;0011006;0011008;0011011;0011067;0011069;0011072;0011087;0011106;0050776"

X = "select CLIGRPCLI from " & paramIBM_Library_SAB & ".ZCLIGRP0 " _
     & " where CLIGRPETB= 1 and CLIGRPREG = '0006002' and CLIGRPREL = 'SFI'" _
     & " order by CLIGRPCLI"
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
        If mENG_LFB_6002 = "" Then
            mENG_LFB_6002 = mENG_LFB_6002 & rsSab("CLIGRPCLI")
        Else
            mENG_LFB_6002 = mENG_LFB_6002 & ";" & rsSab("CLIGRPCLI")
        End If
    rsSab.MoveNext
Loop



Call lstErr_Clear(lstErr, cmdContext, "> SAb_Balance_Import.....attendre 2 minutes !"): DoEvents


fgSelect_Display
Call lstErr_AddItem(lstErr, cmdContext, ". SAb_Balance_Import cbo"): DoEvents

fgSelect.Visible = False


Call rsYBIATAB0_cboK2("DEVISE", "ISO", cboDevise)
arrDevise_Nb = cboDevise.ListCount
ReDim arrDevise(cboDevise.ListCount + 1)
For I = 0 To cboDevise.ListCount - 1
    cboDevise.ListIndex = I
    arrDevise(I + 1) = Trim(cboDevise.Text)
Next I
arrService_Nb = 1
ReDim arrService(1): arrService(0) = "***": arrService(1) = ""
arrService_Balance_Cumul_Z

rsZPLAN0_cboPLANCOOBL cboPCEC
rsZPLAN0_cboPLANCOPRO cboPLANCOPRO

'chkSelect = "0"
cboDevise.AddItem ""
cboPCEC.AddItem ""
cboPLANCOPRO.AddItem ""

cboSelect_CLIENACAT.Clear
rsYBIATAB0_cboK2 "SAB", "CLIENACAT", cboSelect_CLIENACAT
cboSelect_CLIENACAT.AddItem " "
cboSelect_CLIENACAT.ListIndex = 0

fgSelect.Visible = True

Call lstErr_ChangeLastItem(lstErr, cmdContext, "= SAb_Balance_Import"): DoEvents

Me.Enabled = True: Me.MousePointer = 0



End Function

Private Sub cboSelect_CLIENACAT_Click()
If Me.Enabled Then fgSelect.Clear


mcboSelect_CLIENACAT = Trim(cboSelect_CLIENACAT.Text)

End Sub

Private Sub cboSelect_CLIENACAT_GotFocus()
txt_GotFocus cboSelect_CLIENACAT

End Sub

Private Sub cboSelect_CLIENACAT_LostFocus()
txt_LostFocus cboSelect_CLIENACAT

End Sub


Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = 0 To fgSelect_arrIndex
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 0 To fgSelect_arrIndex
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.Col = 0
    End If
End If

End Sub

Public Sub fglist_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgList.Row

If lRow > 0 And lRow < fgList.Rows Then
    fgList.Row = lRow
    For I = 0 To fgList_arrIndex
        fgList.Col = I: fgList.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgList.Row = mRow
    If fgList.Row > 0 Then
        lRow = fgList.Row
        lColor_Old = fgList.CellBackColor
        For I = 0 To fgList_arrIndex
          fgList.Col = I: fgList.CellBackColor = lColor
        Next I
        fgList.Col = 0
    End If
End If

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

Private Sub cboDevise_Change()
If Me.Enabled Then fgSelect.Clear

End Sub

Private Sub cboDevise_Click()
If Me.Enabled Then fgSelect.Clear

If mcboDevise <> cboDevise.Text Then
    mcboDevise = Trim(cboDevise.Text)
End If

End Sub

Private Sub cboDevise_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub cboPCEC_Click()
If Me.Enabled Then fgSelect.Clear


mcboPCEC = Trim(cboPCEC.Text)

End Sub

Private Sub cboPCEC_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub


Private Sub cboPLANCOPRO_Click()
If Me.Enabled Then fgSelect.Clear
mcboPLANCOPRO = Trim(cboPLANCOPRO.Text)

End Sub


Private Sub cboPLANCOPRO_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub chkBalance_CSV_Click()
If chkBalance_CSV = "1" Then
    txtBalance_CSV_Folder.Enabled = True
    txtBalance_CSV_FileName.Enabled = True
Else
    txtBalance_CSV_Folder.Enabled = False
    txtBalance_CSV_FileName.Enabled = False
End If
End Sub

Private Sub chkSelect_Annulé_Click()
If Me.Enabled Then fgSelect.Clear

End Sub

Private Sub chkSelect_Click()
If blnControl Then
    If chkSelect = "1" Then
        'fraSelect_Clear
        chkSelect.BackColor = warnUsrColor
        Me.Enabled = False: Me.MousePointer = vbHourglass
        fgSelect_Display
        Me.Enabled = True: Me.MousePointer = 0
    Else
        chkSelect.BackColor = fraSelect_Options.BackColor
    End If
    
End If
End Sub


Private Sub chkSelect_COMPTECLO_Click()
If Me.Enabled Then fgSelect.Clear
If chkSelect_COMPTECLO = "0" Then
    txtSelect_COMPTECLO.Visible = False
Else
    txtSelect_COMPTECLO.Visible = True
End If

End Sub

Private Sub chkSelect_COMPTEOUV_Click()
If Me.Enabled Then fgSelect.Clear
If chkSelect_COMPTEOUV = "0" Then
    txtSelect_COMPTEOUV.Visible = False
Else
    txtSelect_COMPTEOUV.Visible = True
End If

End Sub

Private Sub chkSelect_HB_Click()
If Me.Enabled Then fgSelect.Clear

End Sub

Private Sub chkSelect_SoldeCr_Click()
If Me.Enabled Then fgSelect.Clear

End Sub


Private Sub chkSelect_SoldeDb_Click()
If Me.Enabled Then fgSelect.Clear

End Sub


Private Sub chkSelect_SOLDEDMO_Click()
If Me.Enabled Then fgSelect.Clear
If chkSelect_SOLDEDMO = "0" Then
    fraSelect_SOLDEDMO.Visible = False
Else
    fraSelect_SOLDEDMO.Visible = True
    chkSelect_DORCPTDMV = "1"
End If


End Sub

Private Sub chkSelect_SoldeZ_Click()
If Me.Enabled Then fgSelect.Clear

End Sub


Private Sub cmdBalance_Ok_Click()
Dim I As Integer, K As Integer, X As String
Dim wBalance_Ok_Param As String

Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Balance : " & fgSelect.Rows - 1)

' Rechercher tous les comptes (même solde = 0) , n'imprimer que les comptes non soldés à la date
If chkSelect_SoldeZ = "1" Then fgSelect_Display_SoldeZ

'''If chkBalance_Pays Then   'test exclure CLIENARSD = "  "
'    blnControl = False
'    fraSelect_Clear
'    chkSelect_SoldeZ = "1"
'    chkSelect = "1"
'    blnSelect_Pays = True
    
'    fgSelect_Display
''''End If

wBalance_Ok_Param = cmdBalance_Ok_Param
cmdBalance_Ok_Print wBalance_Ok_Param, " "

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdBalance_Quit_Click()
cmdContext_Quit
End Sub

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdList_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
fraList.Visible = False
prtRIB_ZADRESS0 xYBIACPT0.COMPTECOM, meZADRESS0
Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdList_Quit_Click()
fraList.Visible = False
End Sub


Private Sub cmdOptions_Click()
fraList.Visible = False
If fraSelect_Options.Visible Then
    fraSelect_Options.Visible = False
Else
    fraSelect_Options.Visible = True
End If

End Sub

Private Sub cmdPrint_Click()
Msg = Space$(50)
Me.Enabled = False: Me.MousePointer = vbHourglass
txtAMJ_Control
Select Case SSTab1.Tab
    Case 0:
            If fgSelect.Rows > 1 Then
                Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
           End If
    Case 1:
'            If fgYBIAMVT0.Rows > 1 Then
'                Me.PopupMenu mnuPrint1, vbPopupMenuLeftButton
'           End If
    Case 2:
            If fgYAUTE1I0.Rows > 1 Then Me.PopupMenu mnuPrint2, vbPopupMenuLeftButton

End Select
Me.Enabled = True: Me.MousePointer = 0



End Sub


Private Sub cmdSelect_MesComptes_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

fgSelect.Clear
chkSelect = "0"
cboPCEC = ""
cboSelect_CLIENACAT = ""
cboPLANCOPRO = ""
mcboDevise = ""
txtCompte = ""
txtIntitulé = currentCLIENASIG
blnMesComptes = True
fgSelect_Display
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuSelect_ENG(lList_Act As String, lList As String)
Dim I As Long, blnOk As Boolean, X As String, curX As Currency
Dim marrYBIACPT0_ENG() As typeYBIACPT0

X = YBIATAB0_DATE_CPT_J
xAmjMin = InputBox("ATTENTION : pas de contrôle de validité de la date : " _
    & vbCrLf & "     =========================" & vbCrLf & Replace(X, ";", vbCrLf) _
    & vbCrLf & "     =========================", "en date de traitement du (AAAAMMJJ)", X)
If Trim(xAmjMin) = "" Then GoTo Exit_sub

xAmjMax = xAmjMin


ReDim marrYBIACPT0_ENG(marrYBIACPT0_Nb + 1)
For I = 0 To marrYBIACPT0_Nb
    marrYBIACPT0_ENG(I) = marrYBIACPT0(I)
Next I

fgSelect_Reset

fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Visible = False
fgSelect.Row = 0
fgSelect.Col = 1: fgSelect.CellAlignment = 1

''xAmjMin = YBIATAB0_DATE_CPT_J
'''xAmjMax = YBIATAB0_DATE_CPT_J

For I = 1 To marrYBIACPT0_Nb
    
    If InStr(lList, marrYBIACPT0_ENG(I).CLIENACLI) > 0 Then
        If Mid$(marrYBIACPT0_ENG(I).COMPTEOBL, 1, 2) = "98" Or marrYBIACPT0_ENG(I).PLANCOPRO = "NOS" Then
        Else
            curX = marrYBIACPT0_ENG(I).SOLDECEN
            If xAmjMin <> YBIATAB0_DATE_CPT_J Then
                Call mnuSelect_ENG_Solde(marrYBIACPT0_ENG(I).COMPTECOM, xAmjMin, curX)
                marrYBIACPT0_ENG(I).SOLDECEN = curX
            End If
            If curX > 0 Then
               xYBIACPT0 = marrYBIACPT0_ENG(I)
                fgSelect_DisplayLine I
            End If
        End If
    End If
    
Next I

fgSelect_SortAD = 6
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_SortX 0

prtSAB_Balance_Monitor "LT-ENG-" & lList_Act, xAmjMin, xAmjMax, fgSelect, marrYBIACPT0_ENG(), marrYBIACPT0_Nb, "", wYSTOMON(), wDORCPTDMV()

'_____________________________________________________________________________

fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Visible = False
fgSelect.Row = 0
fgSelect.Col = 1: fgSelect.CellAlignment = 1
For I = 1 To marrYBIACPT0_Nb
    
    If InStr(lList, marrYBIACPT0_ENG(I).CLIENACLI) > 0 Then
            curX = marrYBIACPT0_ENG(I).SOLDECEN
            If xAmjMin <> YBIATAB0_DATE_CPT_J Then
                Call mnuSelect_ENG_Solde(marrYBIACPT0_ENG(I).COMPTECOM, xAmjMin, curX)
                marrYBIACPT0_ENG(I).SOLDECEN = curX
            End If
            If curX <> 0 Then
                xYBIACPT0 = marrYBIACPT0_ENG(I)
                fgSelect_DisplayLine I
        End If
    End If
    
Next I

fgSelect_SortAD = 6
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_SortX 0

prtSAB_Balance_Monitor "LT-CPT-" & lList_Act, xAmjMin, xAmjMax, fgSelect, marrYBIACPT0_ENG(), marrYBIACPT0_Nb, "", wYSTOMON(), wDORCPTDMV()
fgSelect.Visible = True

Exit_sub:
Me.Show

'________________________________________________________________________________
End Sub
Private Sub mnuSelect_ENG_Detail_Display(lList_Act As String, lList As String, lAMJ_Ouv As Long)
Dim I As Long, blnOk As Boolean, X As String, curX As Currency
Dim marrYBIACPT0_ENG() As typeYBIACPT0
Dim K As Long, wCellForeColor As Long, wCellBackColor As Long
Dim meCV1 As typeCV, meCV2 As typeCV
Dim sDB As Currency, curCPT As Currency
Dim sDB_Gage As Currency, sDB_Gage_Non As Currency
Dim sDB_Gage_CV As Currency, sDB_Gage_Non_CV As Currency
Dim mRow As Long
Dim rsSab_Z As Recordset
Dim xAMJ_Ouv As String

ReDim marrYBIACPT0_ENG(marrYBIACPT0_Nb + 1)
For I = 0 To marrYBIACPT0_Nb
    marrYBIACPT0_ENG(I) = marrYBIACPT0(I)
Next I
xAMJ_Ouv = dateImp10_S(lAMJ_Ouv + 19000000)
fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = "<Comptes  " & lList_Act & "      |<Intitulé                           |<Echéance   " _
                      & "|>Débit            |>Crédit             |<Devise" _
                      & "|>Débit €          |>Crédit  €           "
fgSelect.Visible = False
fgSelect.Row = 0
fgSelect.Col = 3: fgSelect.CellAlignment = 1
fgSelect.Col = 4: fgSelect.CellAlignment = 1
fgSelect.Col = 6: fgSelect.CellAlignment = 1
fgSelect.Col = 7: fgSelect.CellAlignment = 1
fgSelect.Col = 5: fgSelect.CellAlignment = 2

xAmjMin = YBIATAB0_DATE_CPT_J
xAmjMax = YBIATAB0_DATE_CPT_J
wCellForeColor = &H800000
wCellBackColor = &HD0FFFF

meCV1.DeviseN = 0
meCV1.Montant = 0

meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
meCV2.OpéAmj = YBIATAB0_DATE_CPT_J

For I = 1 To marrYBIACPT0_Nb
    
    If InStr(lList, marrYBIACPT0_ENG(I).CLIENACLI) > 0 Then
        curX = marrYBIACPT0_ENG(I).SOLDECEN
        If curX > 0 Then
            xYBIACPT0 = marrYBIACPT0_ENG(I)
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect.Col = 0: fgSelect.Text = xYBIACPT0.COMPTECOM
            fgSelect.Col = 1: fgSelect.Text = xYBIACPT0.COMPTEINT
            fgSelect.Col = 5: fgSelect.Text = xYBIACPT0.COMPTEDEV
            fgSelect.Col = 3: fgSelect.CellForeColor = &HFF&      'vbRed
            meCV1.Montant = Abs(xYBIACPT0.SOLDECEN)
            fgSelect.Text = Format$(meCV1.Montant, "### ### ### ###.00 ")
    
            fgSelect.Col = 6: fgSelect.CellForeColor = &HFF&      'vbRed
            
            If xYBIACPT0.COMPTEDEV <> "EUR" Then
                meCV1.DeviseIso = xYBIACPT0.COMPTEDEV
                
                Call CV_Calc("J  ", meCV1, meCV2)
                sDB = sDB + meCV2.Montant
                fgSelect.Text = Format$(meCV2.Montant, "### ### ### ###.00 ")
            Else
                sDB = sDB + meCV1.Montant
                fgSelect.Text = Format$(meCV1.Montant, "### ### ### ###.00 ")
            End If
                    
        End If
    End If
    
Next I
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
For K = 0 To 7
    fgSelect.Col = K
    fgSelect.CellBackColor = mColor_Y2
Next K
fgSelect.Col = 6: fgSelect.CellForeColor = &HFF&      'vbRed
fgSelect.CellFontBold = True
fgSelect.Text = Format$(sDB, "### ### ### ###.00 ")

fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1

For I = 1 To marrYBIACPT0_Nb
    
    If InStr(lList, marrYBIACPT0_ENG(I).CLIENACLI) > 0 Then
        curX = marrYBIACPT0_ENG(I).SOLDECEN
        If curX > 0 Then
            xYBIACPT0 = marrYBIACPT0_ENG(I)
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            mRow = fgSelect.Row
            fgSelect.Col = 0: fgSelect.Text = xYBIACPT0.COMPTECOM
            fgSelect.Col = 1: fgSelect.Text = xYBIACPT0.COMPTEINT
            fgSelect.Col = 5: fgSelect.Text = xYBIACPT0.COMPTEDEV
            fgSelect.Col = 3: fgSelect.CellForeColor = &HFF&      'vbRed
            meCV1.Montant = Abs(xYBIACPT0.SOLDECEN)
            fgSelect.CellFontBold = True
            curCPT = meCV1.Montant
            fgSelect.Text = Format$(meCV1.Montant, "### ### ### ###.00 ")
            fgSelect.CellFontSize = 8
            fgSelect.Col = 6: fgSelect.CellForeColor = &HFF&      'vbRed
            fgSelect.CellFontSize = 8
            If xYBIACPT0.COMPTEDEV <> "EUR" Then
                meCV1.DeviseIso = xYBIACPT0.COMPTEDEV
                Call CV_Calc("J  ", meCV1, meCV2)
                fgSelect.Text = Format$(meCV2.Montant, "### ### ### ###.00 ")
            Else
                fgSelect.Text = Format$(meCV1.Montant, "### ### ### ###.00 ")
            End If
            If lAMJ_Ouv > 0 Then fgSelect.Rows = fgSelect.Rows + 2
            Select Case Trim(xYBIACPT0.COMPTEOBL)
                Case "911219"
                    X = "select * from " & paramIBM_Library_SAB & ".ZCAUDOS0 " _
                         & " where CAUDOSETB = 1 and CAUDOSAGE = 1 " _
                         & " and CAUDOSSER = '00'  and CAUDOSSSE = '00' and CAUDOSCAU = 'ESSPLC'" _
                         & " and CAUDOSTRA = 2 and CAUDOSDEV = '" & xYBIACPT0.COMPTEDEV & "'" _
                         & " and CAUDOSTCL = ' ' and CAUDOSTIE = '" & xYBIACPT0.CLIENACLI & "'" _
                         & " order by CAUDOSFIN"
                    
                    Set rsSabX = cnsab.Execute(X)
                    
                    sDB = 0: sDB_Gage = 0: sDB_Gage_Non = 0: sDB_Gage_CV = 0: sDB_Gage_Non_CV = 0
                    Do While Not rsSabX.EOF
                    fgSelect.Rows = fgSelect.Rows + 1
                    fgSelect.Row = fgSelect.Rows - 1
                    For K = 0 To 7
                        fgSelect.Col = K
                        fgSelect.CellFontSize = 8
                    Next K
                    'fgSelect.Col = 1: fgSelect.Text = rsSabX("CAUDOSCAU") & " " & rsSabX("CAUDOSDOS")
                    fgSelect.Col = 2: fgSelect.Text = Format$(rsSabX("CAUDOSFIN") + 19000000, "@@@@/@@/@@")
                    fgSelect.Col = 5: fgSelect.Text = rsSabX("CAUDOSDEV")
                    fgSelect.Col = 3: 'fgSelect.CellForeColor = &HFF&      'vbRed
                    meCV1.Montant = Abs(rsSabX("CAUDOSMNT"))
                    
                    X = "select * from " & paramIBM_Library_SAB & ".ZCAUAMO0 " _
                         & " where CAUAMOETB = 1 and CAUAMOAGE = 1 " _
                         & " and CAUAMOSER = '00'  and CAUAMOSSE = '00' and CAUAMODOS = " & rsSabX("CAUDOSDOS") _
                         & " and CAUAMOCLI = '" & xYBIACPT0.CLIENACLI & "'" _
                         & " order by CAUAMODAT desc"
                    
                    Set rsSab_Z = cnsab.Execute(X)
                    If Not rsSab_Z.EOF Then meCV1.Montant = Abs(rsSab_Z("CAUAMORES"))
                    
                    
                    fgSelect.Text = Format$(meCV1.Montant, "### ### ### ###.00 ")
                    sDB = sDB + meCV1.Montant
                    If lAMJ_Ouv > 0 And rsSabX("CAUDOSDEB") >= lAMJ_Ouv Then
                        sDB_Gage_Non = sDB_Gage_Non + meCV1.Montant
                        fgSelect.Col = 1: fgSelect.Text = rsSabX("CAUDOSCAU") & " " & rsSabX("CAUDOSDOS") & "   " & dateImp10_S(rsSabX("CAUDOSDEB") + 19000000)
                         For K = 1 To 7
                            fgSelect.Col = K
                            fgSelect.CellBackColor = mColor_G0
                        Next K
                   Else
                        sDB_Gage = sDB_Gage + meCV1.Montant
                        fgSelect.Col = 1: fgSelect.Text = rsSabX("CAUDOSCAU") & " " & rsSabX("CAUDOSDOS")
                   End If
                    
                        
                    fgSelect.Col = 6
                    
                    If rsSabX("CAUDOSDEV") <> "EUR" Then
                        meCV1.DeviseIso = rsSabX("CAUDOSDEV")
                        
                        Call CV_Calc("J  ", meCV1, meCV2)
                        fgSelect.Text = Format$(meCV2.Montant, "### ### ### ###.00 ")
                    Else
                        fgSelect.Text = Format$(meCV1.Montant, "### ### ### ###.00 ")
                        meCV2.Montant = meCV1.Montant
                    End If
                    If lAMJ_Ouv > 0 And rsSabX("CAUDOSDEB") >= lAMJ_Ouv Then
                        sDB_Gage_Non_CV = sDB_Gage_Non_CV + meCV2.Montant
                   Else
                        sDB_Gage_CV = sDB_Gage_CV + meCV2.Montant
                   End If
                    rsSabX.MoveNext
                    Loop
                Case Else
                
                    X = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDODOS0 " _
                         & " where DOSSLDPCI = '" & xYBIACPT0.COMPTEOBL & "' " _
                         & " and DOSSLDDEV = '" & xYBIACPT0.COMPTEDEV & "' " _
                         & " and DOSSLDCLI = '" & xYBIACPT0.CLIENACLI & "' " _
                         & " and DOSSLDSTA not in ('  ','80','90')" _
                         & " and CDODOSCOP = DOSSLDOPE and CDODOSDOS = DOSSLDNUM" _
                         & " order by CDODOSVAL"
                    
                    Set rsSabX = cnsab.Execute(X)
                    sDB = 0: sDB_Gage = 0: sDB_Gage_Non = 0: sDB_Gage_CV = 0: sDB_Gage_Non_CV = 0
                    Do While Not rsSabX.EOF
                        'V = rsYDOSSLD0_GetBuffer(rsSabX, xYDOSSLD0)
                    fgSelect.Rows = fgSelect.Rows + 1
                    fgSelect.Row = fgSelect.Rows - 1
                    For K = 0 To 7
                        fgSelect.Col = K
                        fgSelect.CellFontSize = 8
                    Next K
                    fgSelect.Col = 1: fgSelect.Text = rsSabX("DOSSLDOPE") & " " & rsSabX("DOSSLDNUM")
                    fgSelect.Col = 2: fgSelect.Text = Format$(rsSabX("CDODOSVAL") + 19000000, "@@@@/@@/@@")
                    fgSelect.Col = 5: fgSelect.Text = rsSabX("DOSSLDDEV")
                    fgSelect.Col = 3: 'fgSelect.CellForeColor = &HFF&      'vbRed
                    meCV1.Montant = Abs(rsSabX("DOSSLDMSD"))
                    fgSelect.Text = Format$(meCV1.Montant, "### ### ### ###.00 ")
                    sDB = sDB + meCV1.Montant
                    
                    fgSelect.Col = 6
                    
                    If rsSabX("DOSSLDDEV") <> "EUR" Then
                        meCV1.DeviseIso = rsSabX("DOSSLDDEV")
                        
                        Call CV_Calc("J  ", meCV1, meCV2)
                        fgSelect.Text = Format$(meCV2.Montant, "### ### ### ###.00 ")
                    Else
                        meCV2.Montant = meCV1.Montant
                        fgSelect.Text = Format$(meCV1.Montant, "### ### ### ###.00 ")
                    End If
                    
                    If lAMJ_Ouv > 0 And rsSabX("CDODOSOUV") >= lAMJ_Ouv Then
                        sDB_Gage_Non = sDB_Gage_Non + meCV1.Montant
                        sDB_Gage_Non_CV = sDB_Gage_Non_CV + meCV2.Montant
                        fgSelect.Col = 1: fgSelect.Text = rsSabX("CDODOSCOP") & " " & rsSabX("CDODOSDOS") & "   " & dateImp10_S(rsSabX("CDODOSOUV") + 19000000)
                         For K = 1 To 7
                            fgSelect.Col = K
                            fgSelect.CellBackColor = mColor_G0
                        Next K
                   Else
                        sDB_Gage = sDB_Gage + meCV1.Montant
                        sDB_Gage_CV = sDB_Gage_CV + meCV2.Montant
                        fgSelect.Col = 1: fgSelect.Text = rsSabX("CDODOSCOP") & " " & rsSabX("CDODOSDOS")
                   End If
                    rsSabX.MoveNext
                    Loop
            End Select
            
            If sDB = 0 Then
                 wCellBackColor = &HD0FFFF
             Else
                 If sDB <> curCPT Then
                      wCellBackColor = mColor_W1
                  Else
                      wCellBackColor = mColor_Y2
                  End If
             End If
             fgSelect.Row = mRow
             For K = 0 To 7
                 fgSelect.Col = K
                 fgSelect.CellBackColor = wCellBackColor
             Next K
             
            If lAMJ_Ouv > 0 Then
                If sDB_Gage_Non <> 0 Then
                    If sDB_Gage <> 0 Then
                        fgSelect.Row = mRow + 1
                        fgSelect.Col = 1: fgSelect.Text = "dont date d'ouv. < au " & xAMJ_Ouv
                        fgSelect.Col = 3: fgSelect.CellForeColor = &HFF&       'vbRed
                        fgSelect.CellFontBold = True
                        fgSelect.Text = Format$(sDB_Gage, "### ### ### ###.00 ")
                        fgSelect.Col = 6: fgSelect.CellForeColor = &HFF&       'vbRed
                        fgSelect.CellFontBold = True
                        fgSelect.Text = Format$(sDB_Gage_CV, "### ### ### ###.00 ")
                    End If
                    fgSelect.Row = mRow + 2
                    fgSelect.Col = 1: fgSelect.Text = "dont date d'ouv. >= au " & xAMJ_Ouv
                    fgSelect.Col = 3: fgSelect.CellForeColor = &HFF&       'vbRed
                    fgSelect.CellFontBold = True
                    fgSelect.Text = Format$(sDB_Gage_Non, "### ### ### ###.00 ")
                    fgSelect.Col = 6: fgSelect.CellForeColor = &HFF&       'vbRed
                    fgSelect.CellFontBold = True
                    fgSelect.Text = Format$(sDB_Gage_Non_CV, "### ### ### ###.00 ")
                    For K = 1 To 7
                        fgSelect.Col = K
                        fgSelect.CellBackColor = mColor_G2
                    Next K
                End If
            End If
         
        End If
    End If
    
Next I

fgSelect.Visible = True

Exit_sub:
Me.Show

'________________________________________________________________________________
End Sub


Private Sub fgList_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

If fgList.Rows > 1 Then
    Call fglist_Color(fgList_RowClick, MouseMoveUsr.BackColor, fgList_ColorClick)
    fgList.Col = fgList_arrIndex
    meZADRESS0 = arrZADRESS0(Val(fgList.Text))
    fgList.LeftCol = 0
    cmdList_Ok.Visible = True
Else
    Dim XX As String
    XX = MsgBox("Ce compte n'a pas d'adresse" & vbCrLf & "Voulez-vous imprimer un rib ?", vbYesNo, "impression d'un RIB")
    If XX = vbYes Then
        Call rsZADRESS0_Init(meZADRESS0)
        cmdList_Ok_Click
    End If
    
End If

End Sub


Private Sub fgYAUTE1I0_Click()
fgYAUTE1I0.LeftCol = 0

End Sub

Private Sub fgYAUTE1I0_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wId As String
On Error Resume Next
If y <= fgYAUTE1I0.RowHeightMin Then
    Select Case fgYAUTE1I0.Col
        Case 0: fgYAUTE1I0_Sort1 = 0: fgYAUTE1I0_Sort2 = 2: fgYAUTE1I0_Sort
        Case 1: fgYAUTE1I0_Sort1 = 1: fgYAUTE1I0_Sort2 = 2: fgYAUTE1I0_Sort
        Case 2:  fgYAUTE1I0_Sort1 = 2: fgYAUTE1I0_Sort2 = 2: fgYAUTE1I0_Sort
    End Select
Else
'    fgYAUTE1i0_Select
End If
   

End Sub


Private Sub fgYAUTE1I0_RowColChange()
If blnfgYAUTE1I0_Display And picYAUTE1I0.Enabled Then
    If fgYAUTE1I0.Row <> fgYAUTE1I0_RowSelect Then fgYAUTE1i0_Select
End If
End Sub

Private Sub fgYBIAMVT0_Click()
fgYBIAMVT0.LeftCol = 0

End Sub

Private Sub fgYBIAMVT0_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wId As String
On Error Resume Next
If y <= fgYBIAMVT0.RowHeightMin Then
    Select Case fgYBIAMVT0.Col
        Case 2:  fgYBIAMVT0_Sort1 = 2: fgYBIAMVT0_Sort2 = 2: fgYBIAMVT0_Sort
        Case 3: fgYBIAMVT0_Sort1 = 3: fgYBIAMVT0_Sort2 = 3: fgYBIAMVT0_SortX fgYBIAMVT0_Sort1
        Case 4: fgYBIAMVT0_Sort1 = 4: fgYBIAMVT0_Sort2 = 4: fgYBIAMVT0_Sort
    End Select
Else
    If fgYBIAMVT0.Rows > 1 Then
        Call fgYBIAMVT0_Color(fgYBIAMVT0_RowClick, MouseMoveUsr.BackColor, fgYBIAMVT0_ColorClick)
        fgYBIAMVT0.Col = fgYBIAMVT0_arrIndex
         wId = fgYBIAMVT0.Text
            fgSelect.LeftCol = 0
        'If Not IsNull(srvYBIAMVT0_Import_Read(wId, xYBIAMVT0)) = 0 Then
        '    srvYBIAMVT0_ElpDisplay xYBIAMVT0
            
        ' Else
        '    Shell_MsgBox "fgYBIAMVT0_MouseDown# " & xMvtP0.ID & " : " & xMvtP0.Err, vbCritical, Me.Caption, False

        'End If
    End If
   End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If blnfrmSAB_Dossier_DB Then frmSAB_Dossier_DB.Hide
End Sub

Public Sub lstService_Click()
Dim xSQL As String
Dim xService As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, ". SAb_Balance_Import_YBIACPT0 : en cours ...."): DoEvents

xService = Trim(Mid$(lstService, 1, 2))
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0" ''' where PLANCOPRO = 'CAT'"
If xService = "G*" Then
   xSQL = xSQL & "  left outer join  " & paramIBM_Library_SAB & ".ZCOMREF0  on COMPTECOM = COMREFCOM  and COMREFCOR like 'G%' ORDER by COMPTECOM"
   YBIACPT0_SQL_ODBC xSQL '& " ORDER by COMPTECOM"
   fraSelect_Options.Visible = True
Else
    If xService = "G?" Then
        xSQL = xSQL & "  left outer join  " & paramIBM_Library_SAB & ".ZCOMREF0  on COMPTECOM = COMREFCOM  and COMREFCOR like 'G%' ORDER by COMPTECOM"
       YBIACPT0_SQL_ODBC xSQL '& " ORDER by COMPTECOM"
       cmdSelect_Compte_nonAffecté
         fraSelect_Clear
         blnControl = False:    chkSelect = "1": chkSelect_Annulé = "1": chkSelect_SoldeZ = "1": blnControl = True
         fgSelect_Display
         blnControl = False:    chkSelect = "0": blnControl = True
    Else

         xSQL = xSQL & " C , " & paramIBM_Library_SAB & ".ZCOMREF0 R where C.COMPTECOM = R.COMREFCOM AND COMREFCOR = '" & xService & "'" & " ORDER by COMPTECOM"
        YBIACPT0_SQL_ODBC xSQL
         fraSelect_Clear
         blnControl = False:    chkSelect = "1": blnControl = True
         fgSelect_Display
         blnControl = False:    chkSelect = "0": blnControl = True
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, ". SAb_Balance_Import_YBIACPT0 :" & marrYBIACPT0_Nb & " comptes"): DoEvents
SSTab1.Caption = Trim(lstService)
lstService.Visible = False
SSTab1.Visible = True
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuAuto_Autorisation_Dépassement_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
fraSelect_Options.Visible = False
SSTab1.Tab = 2
Call fgYAUTE1I0_Display("Dépassement")
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuAuto_Autorisation_Echeance_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
fraSelect_Options.Visible = False
SSTab1.Tab = 2
Call fgYAUTE1I0_Display("Echeance")
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuAuto_Balance_Service_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
blnBalance_Service_Stock = False
optBalance_YSOLDE0_AP1 = True
cmdBalance_Service
Me.Enabled = True: Me.MousePointer = 0
Me.Show

End Sub

Private Sub mnuAuto_Balance_Stock_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
blnBalance_Service_Stock = True
optBalance_YSOLDE0_J = True
cmdBalance_Service
Me.Enabled = True: Me.MousePointer = 0

Me.Show

End Sub

Private Sub mnuAuto_Client_Stat_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

fraSelect_Clear

cboPLANCOPRO = "CAV"
fgSelect_Display

mnuSelect_Print_Client_Stat_Click
 
Me.Enabled = True: Me.MousePointer = 0
Me.Show

End Sub

Private Sub mnuAuto_Compta_TVA_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

fraSelect_Clear
chkSelect_SoldeZ = "1"

cboPLANCOPRO = "PRO"
chkSelect_Résidence = "1"
X = MsgBox("Inclure les journées complémentaires ?", vbQuestion + vbYesNo, "Comptabilité : état de la TVA")
If X = vbYes Then
    chkSelect_MOUVEMDCO = "1"
Else
    chkSelect_MOUVEMDCO = "0"
End If

X = MsgBox("Sortie vers un classeur Excel ?", vbQuestion + vbYesNo, "Balance : Cumul des mouvements")
If X = vbYes Then
    csvManual = True
Else
    csvManual = False
End If

txtCompte = "70"
'Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_MP1)
'If Mid$(YBIATAB0_DATE_CPT_MP1, 5, 2) = "01" Then
'    Call DTPicker_Set(txtAmjMin, Mid$(YBIATAB0_DATE_CPT_MP1, 1, 6) & "02")
'    Call MsgBox("période du " & txtAmjMin & " au " & txtAmjMax, vbInformation, "déclaration TVA de JANVIER")
'Else
'    Call DTPicker_Set(txtAmjMin, Mid$(YBIATAB0_DATE_CPT_MP1, 1, 6) & "01")
'End If

txtAMJ_Control

fgSelect_Display_SoldeZ

mnuSelect_Print_Cumul_Click
 
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuAuto_FOTC_CHAPRO_Click()
Dim blnOk As Boolean
Dim I As Long
Me.Enabled = False: Me.MousePointer = vbHourglass
fraSelect_Options.Visible = False
fraSelect_Clear
fgSelect_Reset
Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtAmjMin, YBIATAB0_DATE_CPT_JP0)
txtAMJ_Control

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Visible = False


For I = 1 To marrYBIACPT0_Nb

    xYBIACPT0 = marrYBIACPT0(I)
    If xYBIACPT0.SOLDECEN <> 0 Then
        If xYBIACPT0.PLANCOPRO = "CHA" Or xYBIACPT0.PLANCOPRO = "PRO" Then
            If xYBIACPT0.COMPTEDEV = "USD" Or xYBIACPT0.COMPTEDEV = "GBP" Or xYBIACPT0.COMPTEDEV = "JPY" Then
                If InStr(xYBIACPT0.COMPTEINT, "INTER") = 0 Then
                    fgSelect_DisplayLine I
                End If
            End If
        End If
    End If
Next I
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "affichage : " & fgSelect.Rows - 1 & " / " & marrYBIACPT0_Nb)
fgSelect_SortX 2

fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Relevés : " & fgSelect.Rows - 1)

prtSAB_Liste_Monitor "FOTC_CHAPRO", xAmjMin, xAmjMax, fgSelect, marrYBIACPT0(), marrYBIACPT0_Nb, "", wYSTOMON(), wDORCPTDMV()
fgSelect.Visible = True
Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuAuto_Groupe_6000_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

fraSelect_Clear
chkSelect_SoldeZ = "0"
txtCompte = "6000"
fgSelect_Display
fgSelect_SortX 0

Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtAmjMin, YBIATAB0_DATE_CPT_JP0)
txtAMJ_Control

prtSAB_Balance.blnPrint_Relevé_Total_Mvt = False
mnuSelect_Print_Relevé_Click

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnufgSelect_fgYAUTE1I0_Click()
SSTab1.Tab = 2
SSTab1_GotFocus
End Sub

Private Sub mnufgSelect_fgYBIAMVT0_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
If xAmjMin > xAmjMax Then
    MsgBox "Date Début > date fin"
Else
    If xAmjMax > mAmjmax Then
        MsgBox "Date fin > date dernière compta"
    Else
        blnfrmSAB_Dossier_DB = True
        Call frmSAB_Dossier_DB.Form_Init("", xYBIACPT0.COMPTECOM, xAmjMin, xAmjMax, "", "", "", 0)
    End If
End If
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnufgSelect_fgYBIAMVT0_MOUVEMDVA_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
blnfrmSAB_Dossier_DB = True
Call frmSAB_Dossier_DB.Form_Init("MOUVEMDVA", xYBIACPT0.COMPTECOM, xAmjMin, xAmjMax, "", "", "", 0)
Me.Enabled = True: Me.MousePointer = 0
End Sub


Private Sub mnufgSelect_fgYBIASTO0_Click()
SSTab1.Tab = 3
SSTab1_GotFocus
End Sub


Private Sub mnufgSelect_KYC_Click()
Dim Nb As Long
Nb = Val(xYBIACPT0.CLIENACLI)
Nb = DS_Document_Load(CStr(Nb), paramDocuShare_Collection_KYC)
Call lstErr_AddItem(lstErr, cmdContext, "nb documents : " & Nb): DoEvents

End Sub

Private Sub mnufgSelect_Print_RIB_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

fgList_Display

Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub mnuPrint2_Excel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "Enregitrement EXCEL en cours ......")

Select Case mfgYAUTE1I0_Fct
    Case "CPT_OD"
            Call MSflexGrid_Excel("", "SAB_CPT_OD", lblYAUTE1I0, fgYAUTE1I0, fgYAUTE1I0.Cols - 1)
    Case Else
            Call MSflexGrid_Excel("", "SAB_Aut", lblYAUTE1I0, fgYAUTE1I0, fgYAUTE1I0.Cols - 1)
End Select
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "Enregitrement EXCEL terminé")
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint2_Mail_Click()
Dim xDest As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "Envoi mail en cours ......")

xDest = currentSSIWINMAIL
Call mnuPrint2_Mail_Send(xDest)
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "Envoi mail terminé")
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuSelect_CPT_OD_Click()
Dim X As String, wCellForeColor As Long, wCellBackColor As Long
Dim mMOUVEMOPE As String, mMOUVEMNUM As Long, K As Integer
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass

mfgYAUTE1I0_Fct = "CPT_OD"

Call DTPicker_Control(txtAmjMin, xAmjMin)
Call DTPicker_Control(txtAmjMax, xAmjMax)
If xAmjMin > xAmjMax Then
    MsgBox "Date Début > date fin"
    GoTo Exit_sub
End If
SSTab1.Tab = 2
lblYAUTE1I0 = "Liste des OD comptablisées du " & dateImp10_S(xAmjMin) & " au " & dateImp10_S(xAmjMax)
fraSelect_Options.Visible = False
fgYAUTE1I0_Reset

fgYAUTE1I0.Rows = 1
fgYAUTE1I0.FormatString = "Date TRT      |<Service|<Référence                          |<Compte                          |<intitulé                                                  " _
                      & "|>Débit                                |>Crédit                               |<Dev    |< Libellé                                                                " _
                      & "|<Saisi par               |<Saisi le                                   |<Validé par              |<Validé le                                |<Comptabilisé par   |<Compta le "
                           
fgYAUTE1I0.Visible = False
fgYAUTE1I0.Row = 0
fgYAUTE1I0.Col = 5: fgYAUTE1I0.CellAlignment = 1
fgYAUTE1I0.Col = 6: fgYAUTE1I0.CellAlignment = 1
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Liste des OD : ")

If arrMNURUTUTI_Nb = 0 Then arrMNURUTUTI_Load

X = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTHD left outer join " _
     & paramIBM_Library_SAB & ".ZCPTODC " _
     & " on cptodetb = mouvemeta and cptodope = MOUVEMope and cptodpie = MOUVEMNUM And MOUVEMDTR = cptoddco" _
     & " where MOUVEMDTR >= " & xAmjMin - 19000000 & " and MOUVEMDTR <= " & xAmjMax - 19000000 & " and substring( MOUVEMOPE , 1 , 1 ) = '*'" _
     & " order by mouvemdtr , mouvemnum , mouvempie , mouvemecr"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    fgYAUTE1I0.Rows = fgYAUTE1I0.Rows + 1
    fgYAUTE1I0.Row = fgYAUTE1I0.Rows - 1
    If rsSab("MOUVEMOPE") = "*Z1" Then
        wCellForeColor = &H800000
        wCellBackColor = &HB0FFFF
    Else
        wCellForeColor = &H606060
        wCellBackColor = &HE0E0E0
    End If
    If mMOUVEMOPE <> rsSab("MOUVEMOPE") Or mMOUVEMNUM <> rsSab("MOUVEMNUM") Then
        'fgYAUTE1I0.RowHeight(fgYAUTE1I0.Row) = 100
    
        fgYAUTE1I0.Col = 0: fgYAUTE1I0.Text = dateImp10(rsSab("MOUVEMDTR") + 19000000)
        fgYAUTE1I0.CellForeColor = wCellForeColor: fgYAUTE1I0.CellBackColor = wCellBackColor
        fgYAUTE1I0.Col = 1: fgYAUTE1I0.Text = rsSab("MOUVEMSER") & " " & rsSab("MOUVEMSSE")
        fgYAUTE1I0.CellBackColor = wCellBackColor: fgYAUTE1I0.CellForeColor = wCellForeColor
        fgYAUTE1I0.Col = 2: fgYAUTE1I0.Text = rsSab("MOUVEMOPE") & " " & rsSab("MOUVEMNUM") & " " & rsSab("MOUVEMEVE")
        fgYAUTE1I0.CellBackColor = wCellBackColor: fgYAUTE1I0.CellForeColor = wCellForeColor
        fgYAUTE1I0.CellFontSize = 8: fgYAUTE1I0.CellFontBold = True
        fgYAUTE1I0.Col = 9: fgYAUTE1I0.Text = arrMNURUTUTI(rsSab("CPTODUTI"))
        fgYAUTE1I0.CellBackColor = wCellBackColor: fgYAUTE1I0.CellForeColor = wCellForeColor
        fgYAUTE1I0.Col = 10: fgYAUTE1I0.Text = dateImp10(rsSab("CPTODDCR") + 19000000) & " " & Format$(rsSab("CPTODHCR") / 1000, "00:00")
        fgYAUTE1I0.CellFontSize = 8
        fgYAUTE1I0.CellForeColor = &H606060
        fgYAUTE1I0.Col = 11: fgYAUTE1I0.Text = arrMNURUTUTI(rsSab("CPTODUMO"))
        fgYAUTE1I0.CellBackColor = wCellBackColor: fgYAUTE1I0.CellForeColor = wCellForeColor
        fgYAUTE1I0.Col = 12: fgYAUTE1I0.Text = dateImp10(rsSab("CPTODDMO") + 19000000) & " " & Format$(rsSab("CPTODHMO") / 1000, "00:00")
        fgYAUTE1I0.CellFontSize = 8
        fgYAUTE1I0.CellForeColor = wCellForeColor: fgYAUTE1I0.CellForeColor = &H606060
        fgYAUTE1I0.Col = 13: fgYAUTE1I0.Text = arrMNURUTUTI(rsSab("CPTODUCO"))
        fgYAUTE1I0.CellBackColor = wCellBackColor: fgYAUTE1I0.CellForeColor = wCellForeColor
        fgYAUTE1I0.Col = 14: fgYAUTE1I0.Text = dateImp10(rsSab("CPTODDCO") + 19000000)
        fgYAUTE1I0.CellForeColor = &H606060
        fgYAUTE1I0.CellFontSize = 8
        For K = 0 To 14
            fgYAUTE1I0.Col = K
            fgYAUTE1I0.CellBackColor = wCellBackColor ' &HFFFFFF
        Next K

        mMOUVEMOPE = rsSab("MOUVEMOPE")
        mMOUVEMNUM = rsSab("MOUVEMNUM")
        fgYAUTE1I0.Rows = fgYAUTE1I0.Rows + 1
        fgYAUTE1I0.Row = fgYAUTE1I0.Rows - 1
    End If

    fgYAUTE1I0.Col = 3: fgYAUTE1I0.Text = Trim(rsSab("MOUVEMCOM"))
    fgYAUTE1I0.CellForeColor = wCellForeColor ': fgYAUTE1I0.CellBackColor = wCellBackColor
    fgYAUTE1I0.CellFontSize = 8: fgYAUTE1I0.CellFontBold = True
    fgYAUTE1I0.Col = 4: fgYAUTE1I0.Text = Trim(rsSab("COMPTEINT"))
    fgYAUTE1I0.CellForeColor = &H606060 'wCellForeColor ': fgYAUTE1I0.CellBackColor = wCellBackColor
    fgYAUTE1I0.CellFontSize = 8
    If rsSab("MOUVEMMON") > 0 Then
        fgYAUTE1I0.Col = 5: fgYAUTE1I0.Text = Format$(rsSab("MOUVEMMON"), "### ### ### ##0.00")
        fgYAUTE1I0.CellForeColor = vbRed
        fgYAUTE1I0.CellBackColor = &HF0FFFF
        fgYAUTE1I0.Col = 6: fgYAUTE1I0.CellBackColor = &HF0FFFF
    Else
        fgYAUTE1I0.Col = 6: fgYAUTE1I0.Text = Format$(Abs(rsSab("MOUVEMMON")), "### ### ### ##0.00")
        fgYAUTE1I0.CellForeColor = vbBlue
        fgYAUTE1I0.CellBackColor = &HF0FFFF
        fgYAUTE1I0.Col = 5: fgYAUTE1I0.CellBackColor = &HF0FFFF
    End If
    fgYAUTE1I0.Col = 7: fgYAUTE1I0.Text = Trim(rsSab("COMPTEDEV"))
    fgYAUTE1I0.CellBackColor = &HD0FFFF: fgYAUTE1I0.CellForeColor = wCellForeColor
    fgYAUTE1I0.CellFontSize = 8: fgYAUTE1I0.CellFontBold = True
    fgYAUTE1I0.Col = 8: fgYAUTE1I0.Text = Trim(rsSab("LIBELLIB1")) & Trim(rsSab("LIBELLIB2")) & Trim(rsSab("LIBELLIB3")) & Trim(rsSab("LIBELLIB4"))
    fgYAUTE1I0.CellForeColor = wCellForeColor ': fgYAUTE1I0.CellBackColor = wCellBackColor
    fgYAUTE1I0.CellFontSize = 8

    rsSab.MoveNext
Loop

If fgYAUTE1I0.Rows = 1 Then

    fgYAUTE1I0.Rows = fgYAUTE1I0.Rows + 1
    fgYAUTE1I0.Row = fgYAUTE1I0.Rows - 1
    fgYAUTE1I0.Col = 0: fgYAUTE1I0.Text = "NEANT"
    fgYAUTE1I0.CellBackColor = vbYellow: fgYAUTE1I0.CellForeColor = vbRed
End If
fgYAUTE1I0.Visible = True

Me.Show

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_ENG_BEA_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Etats des engagements BEA ")

fraSelect_Options.Visible = False
X = "0011001"
X = InputBox("par défaut : " _
    & vbCrLf & "     =========================" & vbCrLf & Replace(X, ";", vbCrLf) _
    & vbCrLf & "     =========================", "Liste des racines à sélectionner", X)
If Trim(X) = "" Then GoTo Exit_sub

Call mnuSelect_ENG("BEA", X)

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuSelect_ENG_Detail_Auto()
Dim xFile As String, xObjet As String, xMesg As String, xDest As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Etats des engagements BEA - LFB")

fraSelect_Options.Visible = False
'xDest = srvSendMail.Exchange_Distribution("CPT", "@ENG_BEA_LFB")
paramEditionNoPaper_Auto_PgmName = "BIA-ENG-BEA"

Call mnuSelect_ENG_Detail_Display("BEA", "0011001", 1130617)   '!!!!!!!!!!!!!!!!!! voir mnuSelect_ENG_Detail_BEA_Click

Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "BEA : Enregistrement EXCEL en cours ......")
xObjet = "Engagements BEA en date du " & dateImp10_S(YBIATAB0_DATE_CPT_J)
xFile = "C:\Temp\ENG_BEA.xlsx"
Call MSflexGrid_Excel(xFile, "ENG BEA ", xObjet, fgSelect, 9)
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "BEA : Enregistrement EXCEL terminé")

Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "BEA : Envoi mail en cours ......")
'xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
' & xObjet
'Call MSFlexGrid_SendMail(xDest, "ENG BEA", xObjet, xMesg, fgSelect, 9, xFile)
Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S51", xFile, "Archive", "BIA-ENG-BEA")
Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("CPT", "@ENG_BEA_LFB", "")

Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "BEA : Envoi mail terminé")

'______________________________________________________________________________________________________________________

paramEditionNoPaper_Auto_PgmName = "BIA-ENG-LFB"
Call mnuSelect_ENG_Detail_Display("LFB", mENG_LFB, 1130718)    '!!!!!!!!!!!!!!!! voir mnuSelect_ENG_Detail_LFB_Click
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "LFB : Enregistrement EXCEL en cours ......")
xObjet = "Engagements LFB en date du " & dateImp10_S(YBIATAB0_DATE_CPT_J)
xFile = "C:\Temp\ENG_LFB.xlsx"
Call MSflexGrid_Excel(xFile, "ENG LFB ", xObjet, fgSelect, 9)
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "LFB : Enregistrement EXCEL terminé")

Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "LFB : Envoi mail en cours ......")
'xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
' & xObjet
'Call MSFlexGrid_SendMail(xDest, "ENG LFB", xObjet, xMesg, fgSelect, 9, xFile)
Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S51", xFile, "Archive", "BIA-ENG-LFB")
Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("CPT", "@ENG_BEA_LFB", "")
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "LFB : Envoi mail terminé")

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_ENG_Detail_BEA_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Etats des engagements BEA ")

fraSelect_Options.Visible = False

Call mnuSelect_ENG_Detail_Display("BEA", "0011001", 1130617)  '!!!!! voir mnuSelect_ENG_Detail_Auto

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_ENG_Detail_LFB_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Etats des engagements LFB ")

fraSelect_Options.Visible = False

Call mnuSelect_ENG_Detail_Display("LFB", mENG_LFB, 1130718) '!!! voir mnuSelect_ENG_Detail_Auto

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuSelect_ENG_LFB_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Etats des engagements  LFB")

fraSelect_Options.Visible = False
X = InputBox("par défaut : " _
    & vbCrLf & "     =========================" & vbCrLf & Replace(mENG_LFB, ";", vbCrLf) _
    & vbCrLf & "     =========================", "Liste des racines à sélectionner", X)
If Trim(X) = "" Then GoTo Exit_sub

Call mnuSelect_ENG("LFB", X)

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_Adresse_Click()
Dim wFileName As String, X As String
Dim wIndex As Long
Dim mADRESSCOA As String
Dim mCLIENACLI As String
Dim nbR As Long, nbW As Long
Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Visible = False

wFileName = paramFolder_Local & "\Adresse.csv"
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Export des adresses => wFileName : " & fgSelect.Rows)
X = MsgBox("Création du fichier : " & wFileName, vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)
If X = vbYes Then
    currentAction = ""
Else
    Exit Sub
End If
X = MsgBox("Extraire les adresses fiscales (OUI) ou courrier (NON) " & wFileName, vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)
If X = vbYes Then
    mADRESSCOA = ""
Else
    mADRESSCOA = "CO"
End If
'20061212   Sélection DISTINCT champ CLIENACLI
'           préfixe Responsable (CLIENARES) et Nationalité (CLIENANAT)

mCLIENACLI = ""
nbR = 0: nbW = 0
Call FEU_ROUGE
Open wFileName For Output As #2

For I = 1 To fgSelect.Rows - 1
    nbR = nbR + 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect.Cols - 1: wIndex = Val(fgSelect.Text)
    meYBIACPT0 = marrYBIACPT0(wIndex)
    If mCLIENACLI <> meYBIACPT0.CLIENACLI Then
        nbW = nbW + 1
        mCLIENACLI = meYBIACPT0.CLIENACLI
        rsZADRESS0_Init meZADRESS0
        meZADRESS0.ADRESSNUM = meYBIACPT0.COMPTECOM
        meZADRESS0.ADRESSCOA = mADRESSCOA
        V = rsZADRESS0_Compte(meZADRESS0)
        Print #2, meYBIACPT0.CLIENARES & ";" _
              & meYBIACPT0.CLIENANAT & ";" _
              & meYBIACPT0.CLIENAETA & ";" _
              & meYBIACPT0.COMPTECOM & ";" _
              & meZADRESS0.ADRESSRA1 & ";" _
              & meZADRESS0.ADRESSRA2 & ";" _
              & meZADRESS0.ADRESSAD1 & ";" _
              & meZADRESS0.ADRESSAD2 & ";" _
              & meZADRESS0.ADRESSAD3 & ";" _
              & meZADRESS0.ADRESSCOP & ";" _
              & meZADRESS0.ADRESSVIL & ";" _
              & meZADRESS0.ADRESSPAY & ";"
    End If
    
Next I

Close
Call FEU_VERT
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "Export des adresses => wFileName : terminé ")
X = MsgBox("nb écrits : " & nbW & " / nb lus : " & nbR, , Me.Caption)

Me.Show
Me.Enabled = True: Me.MousePointer = 0





End Sub

Private Sub mnuSelect_Print_Balance_Click()
Dim X As String
fraBalance.Visible = True

End Sub

Private Sub mnuSelect_Print_Balance_Stock_Click()
blnService_Printer = False
blnBalance_Stock_détail = True
blnBalance_Service_Stock = True
cmdBalance_Ok_Stock

End Sub

Private Sub mnuSelect_Print_Client_Stat_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Stat Catégorie Client : " & fgSelect.Rows - 1)
fgSelect_SortAD = 6
fgSelect.LeftCol = 0
Call fgSelect_SortX(9000)

prtSAB_Client_Stat "  ", xAmjMin, xAmjMax, fgSelect, marrYBIACPT0(), marrYBIACPT0_Nb
fgSelect.Visible = True
Me.Show
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnufgSelect_Print_Extrait_Click()
Dim blnNewPage As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Extrait : " & meYBIACPT0.COMPTECOM)

Call DTPicker_Control(txtAmjMin, xAmjMin)
Call DTPicker_Control(txtAmjMax, xAmjMax)

If xAmjMin > xAmjMax Then
    MsgBox "Date Début > date fin"
Else
    If xAmjMax > mAmjmax Then
        MsgBox "Date fin > date dernière compta"
    Else
        prtYBIAMVT0_A4_OpenX
        prtYBIAMVT0_A4_Extrait xYBIACPT0.COMPTECOM, xAmjMin, xAmjMax, False, lstErr, "*", "", blnNewPage
        prtYBIAMVT0_A4_Close
        
        Me.Show
    End If
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_Compte_Click()
Dim wFileName As String, X As String
Dim wIndex As Long
Dim mADRESSCOA As String

Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Visible = False

wFileName = paramFolder_Local & "\YBIACPT0_" & YBIATAB0_DATE_CPT_J & ".csv"
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Export des comptes => wFileName : " & fgSelect.Rows)
X = MsgBox("Création du fichier : " & wFileName, vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)
If X = vbYes Then
    currentAction = ""
Else
    Exit Sub
End If
Call FEU_ROUGE
Open wFileName For Output As #2

Print #2, "ETABLISSEMENT;" & "NUMERO PLAN;" & "NUMERO COMPTE;" & "COMPTE OBLIGATOIRE;" _
& "INTITULE;" & "AGENCE;" & "TABLES BASE 013;" & "DATE OUVERTURE;" & "DATE CLOTURE;" & "Lori/Nostri/AUTRE;" _
& "O/N;" & "CLASSE SECURITE;" & "TABLES BASE 015;" & "DATE LIMITE BLOCAGE;" _
& "MOTIF BLOCAGE;" & "CODE SENS SOLDE D/C;" & "DATE MODIFICATION;" & "CODE ETABLISSEMENT;" & "NUMERO CLIENT;" _
& "CODE AGENCE;" & "CODE ETAT;" & "NOM OU DESIGNATION;" & "PRENOM/DESIGNATION;" & "SIGLE USUEL;" _
& "NUMERO SIREN;" & "NUMERO SIRET;" & "DATE DE NAISSANCE;" & "SECT ACTIVITE REGLEM;" & "CDE PAYS NATIONALITE;" _
& "CDE PAYS DE RESIDENC;" & "RESPONS/EXPLOITATION;" & "QUALITE/AG ECONOMIQU;" & "COTE ACTIVITE;" & "COTE PAIEMENT;" _
& "COTE CREDIT;" & "COTE ADMISSION;" & "DAT ATRIB/COTAT BDF;" & "AN DERN BIL COMM BDF;" & "CATEGORIE CLIENT;" _
& "COTATION INTERNE;" & "INTERDICTION CHEQUIE;" & "INTERDIT CHEQUIER;" & "SECTEUR D ACTIVITE;" & "SECTEUR GEOGRAPHIQUE;" _
& "ENTREPRISE LIEE;" & "LANGUE MESSAGERIE;" & "DATE ENTREE AU PAYS;" & "NOM DE JEUNE FILLE;" _
& "BILAN DE MOIS;" & "CLIENT DOUTEUX O/N;" & "ZONE LIBRE DE 3 CAR.;" & "ZONE LIBRE DE 2 CAR.;" & "EXTENTION DU NOM;" _
& "0=CLI/COLL=1/AUTRE=2;" & "TIERS DE REFERENCE;" & "CODE SELECTION;" & "CODE PCS;" & "DATE CREATION;" _
& "TABLES BASE 014;" & "DATE DERNIER MVT;" & "SOLDE ENCOURS;" & "EX référence;" & "NUMERO CLIENT;" _
& "0:PRINCIPAL, 1:AUTRE;" & "0:PRINCIPAL, 1:AUTRE"


For I = 1 To fgSelect.Rows - 1
    
    fgSelect.Row = I
    fgSelect.Col = fgSelect.Cols - 1: wIndex = Val(fgSelect.Text)
    meYBIACPT0 = marrYBIACPT0(wIndex)
    Print #2, meYBIACPT0.COMPTEETA & ";" & meYBIACPT0.COMPTEPLA & ";" & meYBIACPT0.COMPTECOM & ";" _
 & meYBIACPT0.COMPTEOBL & ";" & meYBIACPT0.COMPTEINT & ";" & meYBIACPT0.COMPTEAGE & ";" _
 & meYBIACPT0.COMPTEDEV & ";" & dateIBM_AMJ(meYBIACPT0.COMPTEOUV) & ";" & dateIBM_AMJ(meYBIACPT0.COMPTECLO) & ";" _
 & meYBIACPT0.COMPTELOR & ";" & meYBIACPT0.COMPTESUC & ";" & meYBIACPT0.COMPTECLA & ";" & meYBIACPT0.COMPTEFON & ";" _
 & dateIBM_AMJ(meYBIACPT0.COMPTEBLO) & ";" & meYBIACPT0.COMPTEMOT & ";" & meYBIACPT0.COMPTESEN & ";" & dateIBM_AMJ(meYBIACPT0.COMPTEMOD) & ";" _
 & meYBIACPT0.CLIENAETB & ";" & meYBIACPT0.CLIENACLI & ";" & meYBIACPT0.CLIENAAGE & ";" & meYBIACPT0.CLIENAETA & ";" _
 & meYBIACPT0.CLIENARA1 & ";" & meYBIACPT0.CLIENARA2 & ";" & meYBIACPT0.CLIENASIG & ";" & meYBIACPT0.CLIENASRN & ";" _
 & meYBIACPT0.CLIENASRT & ";" & dateIBM_AMJ(meYBIACPT0.CLIENADNA) & ";" & meYBIACPT0.CLIENAREG & ";" & meYBIACPT0.CLIENANAT & ";" _
 & meYBIACPT0.CLIENARSD & ";" & meYBIACPT0.CLIENARES & ";" & meYBIACPT0.CLIENAECO & ";" & meYBIACPT0.CLIENAACT & ";" _
 & meYBIACPT0.CLIENAPAI & ";" & meYBIACPT0.CLIENACRD & ";" & meYBIACPT0.CLIENAADM & ";" & dateIBM_AMJ(meYBIACPT0.CLIENAATR) & ";" _
 & meYBIACPT0.CLIENABIL & ";" & meYBIACPT0.CLIENACAT & ";" & meYBIACPT0.CLIENACOT & ";" & meYBIACPT0.CLIENACHQ & ";" _
 & meYBIACPT0.CLIENADAT & ";" & meYBIACPT0.CLIENASAC & ";" & meYBIACPT0.CLIENAGEO & ";" & meYBIACPT0.CLIENAENT & ";" _
 & meYBIACPT0.CLIENAMES & ";" & dateIBM_AMJ(meYBIACPT0.CLIENAPAY) & ";" & meYBIACPT0.CLIENAFIL & ";" & meYBIACPT0.CLIENABIM & ";" _
 & meYBIACPT0.CLIENADOU & ";" & meYBIACPT0.CLIENALI1 & ";" & meYBIACPT0.CLIENALI2 & ";" & meYBIACPT0.CLIENAEXT & ";" _
 & meYBIACPT0.CLIENACOL & ";" & meYBIACPT0.CLIENATIE & ";" & meYBIACPT0.CLIENASEL & ";" & meYBIACPT0.CLIENAPCS & ";" _
 & dateIBM_AMJ(meYBIACPT0.CLIENACRE) & ";" & meYBIACPT0.PLANCOPRO & ";" & dateIBM_AMJ(meYBIACPT0.SOLDEDMO) & ";" & meYBIACPT0.SOLDECEN & ";" _
 & "" & ";" & meYBIACPT0.TITULACLI & ";" & meYBIACPT0.TITULAPRI & ";" & meYBIACPT0.TITULATPR & ";"

'meYBIACPT0.COMREFREF :non valide
Next I

Close
Call FEU_VERT
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "Export des comptes => wFileName : terminé ")

Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_fgSelect_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "Enregistrement EXCEL en cours ......")

Call MSflexGrid_Excel("", "SAB_Balance", lblYAUTE1I0, fgSelect, 9)
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "Enregistrement EXCEL terminé")
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_fgSelect_Mail_Click()
Dim xObjet As String, xMesg As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "Envoi mail en cours ......")
xObjet = "SAB_Balance " & dateImp10_S(YBIATAB0_DATE_CPT_J)
xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
 & xObjet
Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Balance", xObjet, xMesg, fgSelect, 9)
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "Envoi mail terminé")

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuSelect_Print_Liste_Xls_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "mnuSelect_Print_Liste_Xls : ")
cmdSelect_SQL_Exportation_Liste
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_PCI_DC_Click()
Dim xFct As String, xR As String
Dim xZCOMREF0 As typeZCOMREF0, xSens As String
Dim xPCI As String
Dim xSQL As String
Dim nbErr_PCI As Long, nbErr_Sens As Long

Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression contrôle PCI : " & marrYBIACPT0_Nb)
xZCOMREF0.COMREFETA = 1
xZCOMREF0.COMREFPLA = 1
xZCOMREF0.COMREFCOM = ""
xZCOMREF0.COMREFCOR = "DC"

prtBalance_PCI_DC_Open " - au : " & dateImp(YBIATAB0_DATE_CPT_J)
nbErr_PCI = 0: nbErr_Sens = 0

For I = 1 To marrYBIACPT0_Nb
    If marrYBIACPT0(I).SOLDECEN <> 0 Then
        'xYBIACPT0 = marrYBIACPT0(I)
        xPCI = Mid$(marrYBIACPT0(I).COMPTEOBL, 1, 6)
        If xPCI <> Mid$(xZCOMREF0.COMREFCOM, 1, 6) Then
            xZCOMREF0.COMREFCOM = xPCI
            xSQL = "select * from " & paramIBM_Library_SAB & ".ZCOMREF0 " _
                & " where COMREFCOM = '" & xZCOMREF0.COMREFCOM & "'" _
                & " and COMREFCOR = '" & xZCOMREF0.COMREFCOR & "'" _
                & " and COMREFETA = " & xZCOMREF0.COMREFETA _
                & " and COMREFPLA = " & xZCOMREF0.COMREFPLA
     
            Set rsSab = cnsab.Execute(xSQL)
            If rsSab.EOF Then
                nbErr_PCI = nbErr_PCI + 1
                prtBalance_PCI_DC_Line marrYBIACPT0(I), "???"
                xSens = "N"
            Else
                xSens = Mid$(rsSab("COMREFREF"), 1, 1)
            End If
        End If
        
        If xSens = "D" And marrYBIACPT0(I).SOLDECEN < 0 Then nbErr_Sens = nbErr_Sens + 1: prtBalance_PCI_DC_Line marrYBIACPT0(I), xSens
        If xSens = "C" And marrYBIACPT0(I).SOLDECEN > 0 Then nbErr_Sens = nbErr_Sens + 1: prtBalance_PCI_DC_Line marrYBIACPT0(I), xSens
       
                

    End If
    
Next I

prtBalance_PCI_DC_Close marrYBIACPT0_Nb, nbErr_PCI, nbErr_Sens
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuZADRESS0_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

fgList_Display

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub picYAUTE1I0_Click()
MsgBox "à faire", vbInformation, "srvYAUTE1I0_ElpDisplay xYAUTE1I0"
End Sub

Private Sub lstYAUTE1I0_Click()

fgYAUTE1I0_Display_GRP lstYAUTE1I0.Text

End Sub

Private Sub mnuAuto_FOTC_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

fraSelect_Clear
chkSelect_SoldeZ = "1"
cboPLANCOPRO = "PO"
fgSelect_Display
fgSelect_SortX 2

Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtAmjMin, YBIATAB0_DATE_CPT_JP0)
txtAMJ_Control

prtSAB_Balance.blnPrint_Relevé_Total_Mvt = True
mnuSelect_Print_Relevé_Click

Me.Enabled = True: Me.MousePointer = 0

End Sub
Private Sub mnuAuto_SOBI_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

fraSelect_Clear
chkSelect_SoldeZ = "1"
Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtAmjMin, YBIATAB0_DATE_CPT_JP0)
txtAMJ_Control

txtCompte = "R91120"
fgSelect_Display
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1
fgSelect_Sort
mnuSelect_Print_liste_T_Click
 
txtCompte = "R98050"
fgSelect_Display
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1
fgSelect_Sort
mnuSelect_Print_liste_T_Click

txtCompte = "R90321"
fgSelect_Display
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1
fgSelect_Sort
mnuSelect_Print_liste_T_Click

txtCompte = "R90322"
fgSelect_Display
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1
fgSelect_Sort
mnuSelect_Print_liste_T_Click

txtCompte = "R91130"
fgSelect_Display
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1
fgSelect_Sort
mnuSelect_Print_liste_T_Click

txtCompte = "R91131"
fgSelect_Display
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1
fgSelect_Sort
mnuSelect_Print_liste_T_Click

txtCompte = "R13221"
fgSelect_Display
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1
fgSelect_Sort
mnuSelect_Print_liste_T_Click

txtCompte = "R911329"
fgSelect_Display
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1
fgSelect_Sort
mnuSelect_Print_liste_T_Click

txtCompte = "R999019"
fgSelect_Display
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1
fgSelect_Sort
mnuSelect_Print_liste_T_Click

txtCompte = "R999029"
fgSelect_Display
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1
fgSelect_Sort
mnuSelect_Print_liste_T_Click


Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuRelevé_Print_Click()
Dim xMin As String, xMax As String
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdMvtM  : ")
'Call DTPicker_Control(txtAmjMin, xMin)
'Call DTPicker_Control(txtAmjMax, xMax)

'cmdMvt_Print xMin, xMax
Call lstErr_AddItem(lstErr, cmdContext, "cmdMvtM : ")

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuRIB_Print_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
'cmdYXXXXXX_Import

'prtRIB_Monitor meYBIACPT0.MOUVEMCOM
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
If blnfrmSAB_Dossier_DB Then frmSAB_Dossier_DB.Hide: blnfrmSAB_Dossier_DB = False

blnMesComptes = False
fgSelect_Display
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub fgSelect_Click()
fgSelect.LeftCol = 0

End Sub

Private Sub fgSelect_LeaveCell()
'On Error Resume Next
'fgSelect.CellBackColor = &HE0E0E0
End Sub

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub


Private Sub mnuContextQuitter_Click()
Unload Me
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
lstErr.Clear: lstErr.Height = 200

If fraList.Visible Then fraList.Visible = False: Exit Sub
If fraBalance.Visible Then fraBalance.Visible = False: Exit Sub

If fraSelect_Options.Visible Then fraContextOptions_Exit: Exit Sub
If picYAUTE1I0.Visible Then picYAUTE1I0.Visible = False: Exit Sub


If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
Else
    SSTab1.Tab = SSTab1.Tab - 1
End If


If currentAction = "" Then
   
Else
    X = MsgBox("Voulez-vous réellement abandonner la mise à jour?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)
    If X = vbYes Then
        currentAction = ""
    Else
        Exit Sub
    End If
End If

End Sub
Public Sub fraContextOptions_Exit()
fraSelect_Options.Visible = False

End Sub


Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
    cmdSelect_Click
Else
    SendKeys "{TAB}"
End If
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
On Error Resume Next
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
fgSelect.Clear: fgSelect.Row = 0
fgSelect_FormatString = fgSelect.FormatString
fgYBIAMVT0.Clear: fgYBIAMVT0.Row = 0
fgYBIAMVT0_FormatString = fgYBIAMVT0.FormatString
SSTab1.Visible = False
fraBalance.Visible = False


End Sub





Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xId As String
Dim V
Dim mRib_Clé As String, mRib_IbanE As String

On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
       ' Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_SortX 0
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_SortX fgSelect_Sort1
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_SortX fgSelect_Sort1
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 4:  fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 5:
            If chkSelect_Résidence = "1" Then
                fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_SortX fgSelect_Sort1
            Else
                fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_Sort
            End If
        Case 6:  fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_SortX fgSelect_Sort1
        Case 7:  fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8:  fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_SortX fgSelect_Sort1
        Case 9:  fgSelect_Sort1 = 9: fgSelect_Sort2 = 9: fgSelect_SortX fgSelect_Sort1
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        fraList.Visible = False
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        txtAMJ_Control
        fgSelect.Col = fgSelect_arrIndex
             xYBIACPT0 = marrYBIACPT0(Val(fgSelect.Text))
             
            If fctUser_Classe_Aut(xYBIACPT0.COMPTECLA) Then
                mnufgSelect_fgYBIAMVT0 = True
                mnufgSelect_fgYBIAMVT0_MOUVEMDVA = True
                mnufgSelect_fgYBIASTO0 = True
                mnufgSelect_fgYAUTE1I0 = True
                mnufgSelect_Print_Extrait = True
            Else
                mnufgSelect_fgYBIAMVT0 = False
                mnufgSelect_fgYBIAMVT0_MOUVEMDVA = False
                mnufgSelect_fgYBIASTO0 = False
                mnufgSelect_fgYAUTE1I0 = False
                mnufgSelect_Print_Extrait = False
            End If
            
             
            fgSelect.LeftCol = 0
            
            'srvYBIACPT0_ElpDisplay xYBIACPT0
            If xYBIACPT0.PLANCOPRO = "CAV" _
            Or xYBIACPT0.PLANCOPRO = "LOR" _
            Or xYBIACPT0.PLANCOPRO = "LOB" Then
                mRib_Clé = Format$(RibClé(strSocBdfE, strSocBdfG, Trim(xYBIACPT0.COMPTECOM), mRib_IbanE), "00")
                mnufgSelect_Print_RIB.Caption = "Imprimer RIB : " & mRib_Clé
                mnufgSelect_Print_RIB.Enabled = True
            Else
                mnufgSelect_Print_RIB.Enabled = False
            End If
            If Mid$(xYBIACPT0.COMPTECOM, 1, 4) = "3889" And xYBIACPT0.COMPTEDEV = "EUR" Then
                mnufgSelect_Print_RIB.Enabled = True
            End If
            If Mid$(xYBIACPT0.COMPTECOM, 1, 4) = "3656" And xYBIACPT0.COMPTEDEV = "EUR" Then
                mnufgSelect_Print_RIB.Enabled = True
            End If
             If Mid$(xYBIACPT0.COMPTECOM, 1, 4) = "3616" And xYBIACPT0.COMPTEDEV = "EUR" Then
                mnufgSelect_Print_RIB.Enabled = True
            End If
           If Mid$(xYBIACPT0.COMPTECOM, 1, 3) = "262" And xYBIACPT0.COMPTEDEV = "EUR" Then
                mnufgSelect_Print_RIB.Enabled = True
            End If
           If Mid$(xYBIACPT0.COMPTECOM, 1, 3) = "162" And xYBIACPT0.COMPTEDEV = "EUR" Then
                mnufgSelect_Print_RIB.Enabled = True
            End If
             If SAB_Balance_Aut.Rapprocher And Trim(xYBIACPT0.CLIENACLI) <> "" Then
                mnufgSelect_KYC.Enabled = True
            Else
                mnufgSelect_KYC.Enabled = False
            End If
           
            mnuZADRESS0.Enabled = True
            mnufgSelect_fgYAUTE1I0.Enabled = False
             fraMvt.Visible = False
            fraYAUTE1I0.Visible = False: mAUTE1ICLI = xYBIACPT0.CLIENACLI
            fraYBIASTO0.Visible = False
            'If Button = vbRightButton Then
                Me.PopupMenu mnufgSelect, vbPopupMenuLeftButton
            'End If
       Else
            Shell_MsgBox "fgSelect_MouseDown# ", vbCritical, Me.Caption, False

    End If
   End If
End Sub

Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = fgSelect.Cols - 1
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

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






Private Sub mnuSelect_Print_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Balance : " & fgSelect.Rows - 1)

prtSAB_Balance_Monitor "L ", xAmjMin, xAmjMax, fgSelect, marrYBIACPT0(), marrYBIACPT0_Nb, "", wYSTOMON(), wDORCPTDMV()
fgSelect.Visible = True
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_liste_T_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Balance_T: " & fgSelect.Rows - 1)

prtSAB_Balance_Monitor "LT", xAmjMin, xAmjMax, fgSelect, marrYBIACPT0(), marrYBIACPT0_Nb, "", wYSTOMON(), wDORCPTDMV()
fgSelect.Visible = True
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuSelect_Print_Cumul_Click()
Dim xFct As String, xR As String
Dim currentSheet As Long
Dim currentRow As Long
Dim wbExcel As Excel.Workbook
Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Balance des mouvements : " & fgSelect.Rows - 1)
xFct = "C"
If chkSelect_Résidence = "1" Then
    fgSelect_Sort1_Old = -1
    fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_SortX fgSelect_Sort1
    xR = "R"
Else
    xR = "-"
End If
If chkSelect_MOUVEMDCO = "1" Then
    xFct = "CDCO" & xR
Else
    xFct = "CDTR" & xR
End If
If chkSelect_SoldeZ = "0" Then
    xFct = xFct & "Z"
Else
    xFct = xFct & "-"
End If
If csvManual Then
    prtSAB_Balance_Monitor_csvManual xFct, xAmjMin, xAmjMax, fgSelect, marrYBIACPT0(), marrYBIACPT0_Nb, "", wYSTOMON(), wDORCPTDMV()
ElseIf xlsManual Then
    prtSAB_Balance_Monitor_xlsManual xFct, xAmjMin, xAmjMax, fgSelect, marrYBIACPT0(), marrYBIACPT0_Nb, "", wYSTOMON(), wDORCPTDMV(), currentRow, wbExcel
Else
    prtSAB_Balance_Monitor xFct, xAmjMin, xAmjMax, fgSelect, marrYBIACPT0(), marrYBIACPT0_Nb, "", wYSTOMON(), wDORCPTDMV()
End If
fgSelect.Visible = True
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuSelect_Print_Relevé_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Relevés : " & fgSelect.Rows - 1)

prtSAB_Balance_Monitor "R", xAmjMin, xAmjMax, fgSelect, marrYBIACPT0(), marrYBIACPT0_Nb, "", wYSTOMON(), wDORCPTDMV()
fgSelect.Visible = True
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub optYAUTE1i0_NIV_1_Click()
lstYAUTE1I0_Click
End Sub

Private Sub optYAUTE1i0_NIV_2_Click()
lstYAUTE1I0_Click
End Sub


Private Sub optYAUTE1i0_NIV_X_Click()

lstYAUTE1I0_Click
End Sub


Private Sub SSTab1_GotFocus()
On Error Resume Next

Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
    Case 1: 'fgYBIAMVT0.LeftCol = 0
            If Not fraMvt.Visible Then
                fgYBIAMVT0_Display xYBIACPT0.COMPTECOM
                 fraMvt.Visible = True
            End If
    Case 2:
            
            'If Not fraYAUTE1I0.Visible Then
                Call fgYAUTE1I0_Display("")  ''xYBIACPT0.CLIENACLI
            '    fraYAUTE1I0.Visible = True
            'End If
   Case 3:
            If Not fraYBIASTO0.Visible Then
                fgYBIASTO0_Display
                fraYBIASTO0.Visible = True
            End If
'            YBIASTO0.LeftCol = 0
End Select

End Sub


Private Sub txtBalance_Print_Trame_KeyPress(Index As Integer, KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtCompte_Change()
If Me.Enabled Then fgSelect.Clear
End Sub

Private Sub txtCompte_GotFocus()
txt_GotFocus txtCompte

End Sub

Private Sub txtCompte_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtIntitulé_Change()
If Me.Enabled Then fgSelect.Clear

End Sub

Private Sub txtIntitulé_GotFocus()
txt_GotFocus txtIntitulé

End Sub

Private Sub txtIntitulé_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtIntitulé_LostFocus()
txt_LostFocus txtIntitulé

End Sub

Private Sub txtAmjMax_GotFocus()
DTPicker_GotFocus txtAmjMax

End Sub


Private Sub txtAmjMax_LostFocus()
DTPicker_LostFocus txtAmjMax
txtAMJ_Control
End Sub


Private Sub txtAmjMin_GotFocus()
DTPicker_GotFocus txtAmjMin

End Sub


Private Sub txtAmjMin_LostFocus()
DTPicker_LostFocus txtAmjMin
txtAMJ_Control

End Sub



Public Sub fgSelect_ForeColor(lColor As Long)
For I = 0 To fgSelect_arrIndex
  fgSelect.Col = I: fgSelect.CellForeColor = lColor
Next I

End Sub





Private Sub txtCompte_LostFocus()
txt_LostFocus txtCompte
End Sub



Public Sub txtAMJ_Control()
Call DTPicker_Control(txtAmjMax, xAmjMax)

If xAmjMax > mAmjmax Then
'    xAmjMax = mAmjmax
'    Call DTPicker_Set(txtAmjMax, mAmjmax)
End If
Call DTPicker_Control(txtAmjMin, xAmjMin)
If xAmjMin < mAmjMin Then
'    xAmjMin = mAmjMin
'    Call DTPicker_Set(txtAmjMin, mAmjMin)
End If
End Sub

Public Sub fraSelect_Clear()
txtCompte = ""
txtIntitulé = ""

chkSelect = "0"
cboDevise.ListIndex = 0
cboPCEC.ListIndex = 0
cboSelect_CLIENACAT.ListIndex = 0
cboPLANCOPRO.ListIndex = 0
cboPCEC = ""
cboSelect_CLIENACAT = ""
cboPLANCOPRO = ""
mcboDevise = ""
txtSelect_CLIENARES = ""
txtSelect_CLIENARSD = ""
txtSelect_COMPTECLA = ""

chkSelect_SoldeZ = "0"
chkSelect_SoldeCr = "0"
chkSelect_SoldeDb = "0"
chkSelect_COMPTEOUV = "0"
chkSelect_COMPTECLO = "0"
chkSelect_SOLDEDMO = "0"
chkSelect_MOUVEMDCO = "0"
chkSelect_Résidence = "0"
chkSelect_HB = "0"
chkSelect_Annulé = "1"
blnSelect_Pays = False

fraMvt.Visible = False
fraYAUTE1I0.Visible = False: mAUTE1ICLI = ""
fraYBIASTO0.Visible = False

End Sub

Private Sub txtSelect_CLIENARES_Change()
If Me.Enabled Then fgSelect.Clear
End Sub


Private Sub txtSelect_CLIENARES_GotFocus()
txt_GotFocus txtSelect_CLIENARES
End Sub


Private Sub txtSelect_CLIENARES_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtSelect_CLIENARES_LostFocus()
txt_LostFocus txtSelect_CLIENARES

End Sub


Private Sub txtSelect_CLIENARSD_Change()
If Me.Enabled Then fgSelect.Clear
End Sub


Private Sub txtSelect_CLIENARSD_GotFocus()
txt_GotFocus txtSelect_CLIENARSD
End Sub


Private Sub txtSelect_CLIENARSD_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSelect_CLIENARSD_LostFocus()
txt_LostFocus txtSelect_CLIENARSD

End Sub



Public Sub picYAUTE1I0_Display()
Dim lineY As Integer
Dim col1 As Integer, col2 As Integer, col3 As Integer
Dim colM0 As Integer, colM1 As Integer, colM2 As Integer, colM3 As Integer

picYAUTE1I0.Cls
picYAUTE1I0.BackColor = &HF0FFFF
If picYAUTE1I0.Width < 7000 Then
    picYAUTE1I0.FontSize = 8
    lineY = 350
    col1 = 800: col2 = 1600: col3 = 2400
    
Else
    picYAUTE1I0.FontSize = 12
    lineY = 500
    col1 = 1300: col2 = 2600: col3 = 3900
End If

colM0 = picYAUTE1I0.Width / 2 + 50
colM1 = picYAUTE1I0.Width / 2 + col1
colM2 = picYAUTE1I0.Width / 2 + col2
colM3 = picYAUTE1I0.Width / 2 + col3

picYAUTE1I0.Visible = True
 picYAUTE1I0.FontBold = False
picYAUTE1I0.CurrentY = 50
picYAUTE1I0.CurrentX = 50: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print "Groupe";
picYAUTE1I0.CurrentX = col1: picYAUTE1I0.ForeColor = libUsr.ForeColor: picYAUTE1I0.Print xYAUTE1I0.AUTE1IGRP & "  " & xYAUTE1I0.AUTE1IREL & " _ " & Trim(xYAUTE1I0.AUTE1IRAG);

picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + lineY
picYAUTE1I0.CurrentX = 50: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print "Resp";
If xYAUTE1I0.AUTE1INOP <> 0 Then picYAUTE1I0.CurrentX = col1: picYAUTE1I0.ForeColor = libUsr.ForeColor: picYAUTE1I0.Print xYAUTE1I0.AUTE1IRES;

picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + lineY
picYAUTE1I0.CurrentX = 50: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print "Client";
picYAUTE1I0.CurrentX = col1: picYAUTE1I0.ForeColor = libUsr.ForeColor
picYAUTE1I0.FontBold = True: picYAUTE1I0.Print xYAUTE1I0.AUTE1ICLI;
picYAUTE1I0.FontBold = False: picYAUTE1I0.Print " _ " & Trim(xYAUTE1I0.AUTE1IRA1) & " " & Trim(xYAUTE1I0.AUTE1IRA2);

picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + lineY
picYAUTE1I0.CurrentX = 50: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print "Cot BDF ";
picYAUTE1I0.CurrentX = col1: picYAUTE1I0.ForeColor = libUsr.ForeColor: picYAUTE1I0.Print xYAUTE1I0.AUTE1IBDF;
picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print " Interne ";
picYAUTE1I0.ForeColor = libUsr.ForeColor: picYAUTE1I0.Print xYAUTE1I0.AUTE1ICOT;
If xYAUTE1I0.AUTE1IDOU = "O" Then picYAUTE1I0.ForeColor = vbRed: picYAUTE1I0.Print "DOUTEUX ";
If xYAUTE1I0.AUTE1IICH = "O" Then picYAUTE1I0.ForeColor = vbRed: picYAUTE1I0.Print "Interdit CHQ ";
picYAUTE1I0.ForeColor = libUsr.ForeColor

picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + lineY
picYAUTE1I0.Line (0, picYAUTE1I0.CurrentY)-(picYAUTE1I0.Width, picYAUTE1I0.CurrentY)
picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + 100
picYAUTE1I0.FontBold = False
picYAUTE1I0.CurrentX = 50: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print "Dépassement   ";


If xYAUTE1I0.AUTE1IMTD <> 0 Then
    picYAUTE1I0.FontBold = True:    picYAUTE1I0.ForeColor = vbRed
    picYAUTE1I0.Print Format$(xYAUTE1I0.AUTE1IMTD, "### ### ### ##0.00");
     If xYAUTE1I0.AUTE1IDTD <> 0 Then
        picYAUTE1I0.FontBold = False
        picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print "  depuis le ";
        picYAUTE1I0.ForeColor = libUsr.ForeColor: picYAUTE1I0.Print dateIBM10(xYAUTE1I0.AUTE1IDTD, True);
        End If
End If
picYAUTE1I0.ForeColor = libUsr.ForeColor
picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + lineY
picYAUTE1I0.Line (0, picYAUTE1I0.CurrentY)-(picYAUTE1I0.Width, picYAUTE1I0.CurrentY)
picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + 100
picYAUTE1I0.FontBold = False
picYAUTE1I0.CurrentX = 50: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print "Autorisation   ";
picYAUTE1I0.FontBold = True:
picYAUTE1I0.ForeColor = vbMagenta
If xYAUTE1I0.AUTE1IMAU <> 0 Then
    picYAUTE1I0.Print Format$(xYAUTE1I0.AUTE1IMAU, "### ### ### ##0.00");
     If xYAUTE1I0.AUTE1IDAD <> 0 Then
        picYAUTE1I0.FontBold = False
        picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print "  du ";
        picYAUTE1I0.ForeColor = libUsr.ForeColor: picYAUTE1I0.Print dateIBM10(xYAUTE1I0.AUTE1IDAD, True);
        End If
     If xYAUTE1I0.AUTE1IDAF <> 0 Then
        picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print "  au ";
        picYAUTE1I0.FontBold = True
        picYAUTE1I0.ForeColor = libUsr.ForeColor: picYAUTE1I0.Print dateIBM10(xYAUTE1I0.AUTE1IDAF, True);
    End If
End If

picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + lineY
picYAUTE1I0.CurrentX = 50: picYAUTE1I0.ForeColor = libUsr.ForeColor
picYAUTE1I0.FontBold = True: picYAUTE1I0.Print Trim(xYAUTE1I0.AUTE1IAUT);
picYAUTE1I0.FontBold = False: picYAUTE1I0.Print " _ " & xYAUTE1I0.AUTE1ILAU & " " & xYAUTE1I0.AUTE1ICOP & " " & xYAUTE1I0.AUTE1INOP;

picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + lineY
picYAUTE1I0.CurrentX = 50: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print "Blocage";
picYAUTE1I0.CurrentX = col1: picYAUTE1I0.ForeColor = libUsr.ForeColor: picYAUTE1I0.Print xYAUTE1I0.AUTE1IBLO;
picYAUTE1I0.CurrentX = colM0: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print " Niveau ";
picYAUTE1I0.CurrentX = colM1: picYAUTE1I0.ForeColor = libUsr.ForeColor: picYAUTE1I0.Print xYAUTE1I0.AUTE1INIV;
If xYAUTE1I0.AUTE1IELM = "O" Then picYAUTE1I0.Print " élémentaire ";

picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + lineY
picYAUTE1I0.Line (0, picYAUTE1I0.CurrentY)-(picYAUTE1I0.Width, picYAUTE1I0.CurrentY)
picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + 100

picYAUTE1I0.CurrentX = 50: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.FontBold = False: picYAUTE1I0.Print "Débit";
picYAUTE1I0.FontBold = True:
If xYAUTE1I0.AUTE1IMDB <> 0 Then
    picYAUTE1I0.CurrentX = col1: picYAUTE1I0.ForeColor = vbRed: picYAUTE1I0.Print Format$(xYAUTE1I0.AUTE1IMDB, "### ### ### ##0.00");
    picYAUTE1I0.ForeColor = vbBlack: picYAUTE1I0.Print " " & xYAUTE1I0.AUTE1IDEV;
End If
picYAUTE1I0.CurrentX = colM0: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.FontBold = False: picYAUTE1I0.Print "Crédit";
picYAUTE1I0.FontBold = True:
If xYAUTE1I0.AUTE1IMCR <> 0 Then
    picYAUTE1I0.CurrentX = colM1: picYAUTE1I0.ForeColor = vbBlue: picYAUTE1I0.Print Format$(xYAUTE1I0.AUTE1IMCR, "### ### ### ##0.00");
    picYAUTE1I0.ForeColor = vbBlack: picYAUTE1I0.Print " " & xYAUTE1I0.AUTE1IDEV;
End If
If xYAUTE1I0.AUTE1IDEV <> xYAUTE1I0.AUTE1IDBA Then
    picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + lineY
    picYAUTE1I0.CurrentX = 50: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.FontBold = False: picYAUTE1I0.Print "Débit";
    picYAUTE1I0.FontBold = True:
    If xYAUTE1I0.AUTE1IBDB <> 0 Then
        picYAUTE1I0.CurrentX = col1: picYAUTE1I0.ForeColor = vbRed: picYAUTE1I0.Print Format$(xYAUTE1I0.AUTE1IBDB, "### ### ### ##0.00");
        picYAUTE1I0.ForeColor = vbBlack: picYAUTE1I0.Print " " & xYAUTE1I0.AUTE1IDBA;
    End If
    picYAUTE1I0.CurrentX = colM0: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.FontBold = False: picYAUTE1I0.Print "Crédit";
    picYAUTE1I0.FontBold = True:
    If xYAUTE1I0.AUTE1IBCR <> 0 Then
        picYAUTE1I0.CurrentX = colM1: picYAUTE1I0.ForeColor = vbBlue: picYAUTE1I0.Print Format$(xYAUTE1I0.AUTE1IBCR, "### ### ### ##0.00");
        picYAUTE1I0.ForeColor = vbBlack: picYAUTE1I0.Print " " & xYAUTE1I0.AUTE1IDBA;
    End If
End If
picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + lineY
picYAUTE1I0.Line (0, picYAUTE1I0.CurrentY)-(picYAUTE1I0.Width, picYAUTE1I0.CurrentY)
picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + 100


If xYAUTE1I0.AUTE1IDMO <> 0 Then
    picYAUTE1I0.FontBold = False
    picYAUTE1I0.CurrentX = 50: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print "Der Mvt ";
    picYAUTE1I0.FontBold = True: picYAUTE1I0.CurrentX = col1
    picYAUTE1I0.ForeColor = libUsr.ForeColor: picYAUTE1I0.Print dateIBM10(xYAUTE1I0.AUTE1IDMO, True);
    picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + lineY
End If

picYAUTE1I0.FontBold = False
If xYAUTE1I0.AUTE1INOP <> 0 Then
    picYAUTE1I0.CurrentX = 50: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print "Dossier";
    picYAUTE1I0.CurrentX = col1: picYAUTE1I0.ForeColor = libUsr.ForeColor: picYAUTE1I0.Print xYAUTE1I0.AUTE1ISER & " " & xYAUTE1I0.AUTE1ISRV & " " & xYAUTE1I0.AUTE1ICOP & " " & xYAUTE1I0.AUTE1INOP;
     If xYAUTE1I0.AUTE1IFIN <> 0 Then
        picYAUTE1I0.CurrentX = colM0: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print "Fin";
        picYAUTE1I0.FontBold = True: picYAUTE1I0.CurrentX = colM1
        picYAUTE1I0.ForeColor = libUsr.ForeColor: picYAUTE1I0.Print dateIBM10(xYAUTE1I0.AUTE1IFIN, True);
    End If
    If xYAUTE1I0.AUTE1IDEB <> 0 Then
        picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + lineY
        picYAUTE1I0.FontBold = False
        picYAUTE1I0.CurrentX = colM0: picYAUTE1I0.ForeColor = lblUsr.ForeColor: picYAUTE1I0.Print "Début ";
        picYAUTE1I0.FontBold = True: picYAUTE1I0.CurrentX = colM1
        picYAUTE1I0.ForeColor = libUsr.ForeColor: picYAUTE1I0.Print dateIBM10(xYAUTE1I0.AUTE1IDEB, True);
    End If

    picYAUTE1I0.CurrentY = picYAUTE1I0.CurrentY + lineY
    picYAUTE1I0.CurrentX = 50
    picYAUTE1I0.ForeColor = libUsr.ForeColor: picYAUTE1I0.Print Trim(xYAUTE1I0.AUTE1ILIB)
End If
End Sub

Public Sub fgYAUTE1i0_Select()
If fgYAUTE1I0.Rows > 1 Then
    picYAUTE1I0.Enabled = False
    fgYAUTE1I0_RowSelect = fgYAUTE1I0.Row
    Call fgYAUTE1I0_Color(fgYAUTE1I0_RowClick, vbYellow, fgYAUTE1I0_ColorClick)
    fgYAUTE1I0.Col = fgYAUTE1I0_arrIndex - 2
    MsgBox "à faire", vbInformation, "fgYAUTE1i0_Select"
    'MsgTxt = Space$(34) & fgYAUTE1I0.Text
    'MsgTxtIndex = 0
    'srvYAUTE1I0_GetBuffer xYAUTE1I0
    'picYAUTE1I0_Display
    fgSelect.LeftCol = 0
    picYAUTE1I0.Enabled = True
Else
        Shell_MsgBox "fgYAUTE1I0_MouseDown# ", vbCritical, Me.Caption, False

    End If

End Sub

Public Sub fgSelect_Display_SoldeZ()
' !!! chkSelect_SoldeZ = "0" pour sélectionner tous les comptes
'                      = "1" pour ne pas imprimer les comptes soldés sans mvt
chkSelect_SoldeZ = "0"
fgSelect_Display
''chkSelect_SoldeZ = "1"

End Sub

Public Sub YBIACPT0_SQL_ODBC(xSQL As String)
Dim X As String, K As Long, K0 As Long, Nb As Long

On Error GoTo Exit_sub
Set rsSab = Nothing
marrYBIACPT0_Nb = 0
 
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    marrYBIACPT0_Nb = marrYBIACPT0_Nb + 1
    If marrYBIACPT0_Nb >= UBound(marrYBIACPT0) Then ReDim Preserve marrYBIACPT0(marrYBIACPT0_Nb + 500)
    V = rsYBIACPT0_GetBuffer(rsSab, marrYBIACPT0(marrYBIACPT0_Nb))
    If IsNull(rsSab("COMREFCOR")) Then
        X = marrYBIACPT0(marrYBIACPT0_Nb).PLANCOPRO
        If X = "CAV" Or X = "LIE" Or X = "LOR" Then
            marrYBIACPT0(marrYBIACPT0_Nb).COMREFCOR = "  "
        Else
            marrYBIACPT0(marrYBIACPT0_Nb).COMREFCOR = "G?"
        End If
    Else
        marrYBIACPT0(marrYBIACPT0_Nb).COMREFCOR = rsSab("COMREFCOR")
    End If
    If Not IsNull(V) Then
        MsgBox V & vbCrLf & "ABANDON de la requête", vbCritical, "frmSAB_Balance.SQL_ODBC"
        Exit Sub
    End If
    rsSab.MoveNext
Loop

' 22-03-211 actualisation du champ COMPTEFON

X = "select COMPTECOM,COMPTEFON from " & paramIBM_Library_SAB & ".ZCOMPTE0 order by COMPTECOM"

Set rsSab = cnsab.Execute(X)
K0 = 1
Do While Not rsSab.EOF
    X = rsSab("COMPTECOM")
    For K = K0 To marrYBIACPT0_Nb
        If X = marrYBIACPT0(K).COMPTECOM Then
            marrYBIACPT0(K).COMPTEFON = rsSab("COMPTEFON")
            K0 = K + 1
            'Nb = Nb + 1
            Exit For
        End If
    Next K
    rsSab.MoveNext
Loop

blnZDORCPT_SQL_ODBC = False
ReDim wDORCPTDMV(marrYBIACPT0_Nb)
For I = 1 To marrYBIACPT0_Nb
    wDORCPTDMV(I) = 0
Next I


'________________________________________________________________________________

'Call MsgBox("JPL : spécial 31-07-2012 : à SUPPRIMER")
'X = "select SOLDECOM,SOLDEC01,SOLDEDMO from " & paramIBM_Library_SAB & ".ZSOLDE0 order by SOLDECOM"

'Set rsSab = cnsab.Execute(X)
'K0 = 1
'Do While Not rsSab.EOF
'    X = rsSab("SOLDECOM")
'    For K = K0 To marrYBIACPT0_Nb
'        If X = marrYBIACPT0(K).COMPTECOM Then
'            marrYBIACPT0(K).SOLDECEN = rsSab("SOLDEC01")
'            K0 = K + 1
'            Exit For
'        End If
'    Next K
'    rsSab.MoveNext
'Loop
'______________________________________________________________________________________


Exit_sub:

End Sub



Public Sub ZCOMREF0_SQL_ODBC()
Dim xSQL As String, mService As String
Dim K As Long
Dim X20 As String * 20

Set rsSab = Nothing
ReDim marrZCOMREF0(10000): marrZCOMREF0_Nb = 0
ReDim arrService(20): arrService_Nb = 1: arrService(1) = "": mService = ""
ReDim wCOMREFCOR(marrYBIACPT0_Nb + 1)


xSQL = "select * from " & paramIBM_Library_SAB & ".ZCOMREF0 where COMREFCOR like 'G%' order by COMREFCOR, COMREFCOM"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    marrZCOMREF0_Nb = marrZCOMREF0_Nb + 1
    If marrZCOMREF0_Nb >= UBound(marrZCOMREF0) Then ReDim Preserve marrZCOMREF0(marrZCOMREF0_Nb + 50)
    V = rsZCOMREF0_GetBuffer(rsSab, marrZCOMREF0(marrZCOMREF0_Nb))
    If Not IsNull(V) Then
        MsgBox V, vbCritical, "frmSAB_Balance.SQL_ODBC"
        Exit Sub
    Else
        If mService <> marrZCOMREF0(marrZCOMREF0_Nb).COMREFCOR Then
            arrService_Nb = arrService_Nb + 1
            If arrService_Nb >= UBound(arrService) Then ReDim Preserve arrService(arrService_Nb + 10)
            arrService(arrService_Nb) = marrZCOMREF0(marrZCOMREF0_Nb).COMREFCOR
            mService = arrService(arrService_Nb)
        End If
    End If
    rsSab.MoveNext
Loop

For I = 1 To marrYBIACPT0_Nb
    wCOMREFCOR(I) = ""
Next I


For K = 1 To marrZCOMREF0_Nb
    X20 = marrZCOMREF0(K).COMREFCOM
    For I = 1 To marrYBIACPT0_Nb
        If X20 = marrYBIACPT0(I).COMPTECOM Then
            If Not blnService_Printer Then
                If wCOMREFCOR(I) <> "" Then MsgBox marrYBIACPT0(I).COMPTECOM, vbExclamation, " déjà affecté à " & wCOMREFCOR(I) & "Affectation Compte / service"
            End If
            wCOMREFCOR(I) = marrZCOMREF0(K).COMREFCOR
            Exit For
        End If
    Next I
Next K

End Sub

Public Sub mnuZCOMREF0_Service_Export_click()
Dim xSQL As String, X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

arrService_Load
X = "'ASD' , 'CAI' , 'CBO' , 'CCR' , 'DOR' , 'INT' , 'IMP' , 'LOB' , 'NOB'"
X = InputBox("par défaut : " _
    & vbCrLf & "     =========================" & vbCrLf & X _
    & vbCrLf & "     =========================", "Liste des codes produit à sélectionner", X)
If Trim(X) = "" Then GoTo Exit_sub

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 left outer join " _
     & paramIBM_Library_SAB & ".ZCOMREF0 on COMREFCOM = COMPTECOM and substring(COMREFCOR , 1 , 1 ) = 'G'" _
     & " where COMPTEFON <> '4' and PLANCOPRO in (" & X & ")" _
     & " order by PLANCOPRO , COMPTEOBL , COMPTECOM"

Set rsSab = cnsab.Execute(xSQL)

cmdExcel_Init ("CPT")

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdExcel_Init(lTxt As String)
On Error GoTo Error_Handler
Dim xSQL As String
Dim X As String, wFile As String, wFilex As String
Dim blnCALCS As Boolean

On Error GoTo Error_Handler
'===================================================================================
'______________________________________________'

X = "C:\Temp\"

mXls1_File = mXls1_File + 1

wFile = X & Trim("SAB " & lTxt & " " & DSYS_Time & mXls1_File & ".xlsx")
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "BDF_CRT : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then mXls1_File = mXls1_File - 1: Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If
'_________________________________________


If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile

'=========================================================================================
Call lstErr_AddItem(lstErr, cmdContext, "Fichier excel.... : "): DoEvents

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "SAB_Balance"
    .Subject = "SAB_Balance"
End With

'__________________________________________________________________________________

'appExcel.Worksheets.Add

Set wsExcel = wbExcel.Sheets(1): wsExcel.Name = "SAB_Balance"

Set wsExcel = wbExcel.Sheets(1)

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignRight
    .WrapText = False ' True
    .Font.Size = 8
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 80

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14SAB_Balance, arrêté au " & dateImp10(YBIATAB0_DATE_CPT_J) _
                                 & vbCr

wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

mXls1_Row = 1

Call cmdExcel_YBIACPT0

        


'======================================================================================================

Exit_sub:
'__________________________________________________________________________________
Set rsSab = Nothing


wbExcel.SaveAs wFile

wbExcel.Close

'____________________________________________________________________________________
appExcel.Quit

Set rsSab = Nothing

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents


'_____________________________
Exit Sub

Error_Handler:
    If Not blnCALCS Then
        X = "C:\Temp\"
        Resume Next
    End If
    MsgBox Error, vbCritical, Me.Name
    Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents
    
    wbExcel.SaveAs wFile
    wbExcel.Close
    appExcel.Quit

End Sub
Public Sub cmdExcel_YBIACPT0()
Dim X As String, K As Integer, kService As Integer, curX As Currency
On Error GoTo Error_Handler

'===================================================================================

wsExcel.Name = "Comptes"

With wsExcel.Cells
    .HorizontalAlignment = Excel.xlHAlignLeft
    .Font.Size = 10
    .Font.Name = "Calibri"
End With
wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 75
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14SAB_Balance : Comptes" _
                                & "  (édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$E1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


wsExcel.Columns(1).ColumnWidth = 8: wsExcel.Cells(1, 1) = "Produit "
wsExcel.Columns(2).ColumnWidth = 10: wsExcel.Cells(1, 2) = "PCI"
wsExcel.Columns(3).ColumnWidth = 20: wsExcel.Cells(1, 3) = "Compte"
wsExcel.Columns(4).ColumnWidth = 8: wsExcel.Cells(1, 4) = "R.Com"
wsExcel.Columns(5).ColumnWidth = 20: wsExcel.Cells(1, 5) = "Service"
wsExcel.Columns(6).ColumnWidth = 60: wsExcel.Cells(1, 6) = "Intitulé"
wsExcel.Columns(7).ColumnWidth = 18: wsExcel.Cells(1, 7) = "Solde débiteur": wsExcel.Columns(7).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(7).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(8).ColumnWidth = 18: wsExcel.Cells(1, 8) = "Solde créditeur": wsExcel.Columns(8).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(8).HorizontalAlignment = Excel.xlHAlignRight
mXls1_Cols = 8


For K = 1 To mXls1_Cols
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = vbWhite
Next

'________________________________________________________________________________


Do While Not rsSab.EOF
        
    mXls1_Row = mXls1_Row + 1
    wsExcel.Cells(mXls1_Row, 1) = rsSab("PLANCOPRO")
    wsExcel.Cells(mXls1_Row, 2) = rsSab("COMPTEOBL")
    wsExcel.Cells(mXls1_Row, 3) = rsSab("COMPTECOM")
    wsExcel.Cells(mXls1_Row, 4) = rsSab("CLIENARES")
    If Not IsNull(rsSab("COMREFCOR")) Then
        X = rsSab("COMREFCOR")
        If X <> arrService_Code(kService) Then
            For kService = 1 To arrService_Code_Nb
                If X = arrService_Code(kService) Then Exit For
                
            Next kService
        End If
        wsExcel.Cells(mXls1_Row, 5) = arrService_Code(kService) & " " & arrService_Lib(kService)
    End If
    wsExcel.Cells(mXls1_Row, 6) = rsSab("COMPTEINT")
    
    curX = -CCur(rsSab("SOLDECEN")) / 1000
    If curX < 0 Then
         wsExcel.Cells(mXls1_Row, 7) = curX
     Else
         wsExcel.Cells(mXls1_Row, 8) = curX
     End If

    If mXls1_Row Mod 100 Then
        Call lstErr_ChangeLastItem(lstErr, cmdContext, rsSab("COMPTECOM")): DoEvents
        
    End If
    rsSab.MoveNext
Loop



'======================================================================================================

Exit_sub:
'__________________________________________________________________________________


'_____________________________
Exit Sub

Error_Handler:

End Sub

Public Sub XXX_mnuZCOMREF0_Service_Export_click()
Dim I As Integer, K As Integer
Dim wFileName As String

wFileName = paramFolder_Local & "\Compte_Service.csv"
X = MsgBox("Création du fichier : " & wFileName, vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)
If X = vbYes Then
    currentAction = ""
Else
    Exit Sub
End If

ZCOMREF0_SQL_ODBC

Call FEU_ROUGE
Open wFileName For Output As #2

For I = 1 To marrYBIACPT0_Nb

    Print #2, wCOMREFCOR(I) & ";" _
          & marrYBIACPT0(I).CLIENARES & ";" _
          & marrYBIACPT0(I).COMPTEOBL & ";" _
          & marrYBIACPT0(I).COMPTECOM & ";" _
          & marrYBIACPT0(I).COMPTEDEV & ";" _
          & marrYBIACPT0(I).COMPTEFON & ";" _
          & marrYBIACPT0(I).PLANCOPRO & ";" _
          & marrYBIACPT0(I).COMPTEINT & ";" _
         & marrYBIACPT0(I).CLIENARA1
Next I

Close
Call FEU_VERT
End Sub

Public Function cmdBalance_Ok_Param() As String
Dim X As String
Dim I As Integer, K As Integer

X = "BD0J00000000000" & String(8 * 6, "0")

Mid$(X, 3, 1) = chkBalance_Détail
Mid$(X, 5, 1) = chkBalance_Récap
If optBalance_YSOLDE0_MP1 Then Mid$(X, 4, 1) = "M"
If optBalance_YSOLDE0_MP2 Then Mid$(X, 4, 1) = "2"
If optBalance_YSOLDE0_AP1 Then Mid$(X, 4, 1) = "A"
Mid$(X, 10, 1) = chkBalance_Compte_Soldé
Mid$(X, 12, 1) = chkBalance_Récap_Bilan

For I = 0 To 7
    K = 15 + 6 * I
    Mid$(X, K + 1, 1) = chkBalance_Print(I)
    Mid$(X, K + 2, 1) = chkBalance_Print_FontBold(I)
    Mid$(X, K + 3, 1) = chkBalance_Print_Line(I)
    Mid$(X, K + 4, 3) = Format$(Val(txtBalance_Print_Trame(I)), "000")

Next I
cmdBalance_Ok_Param = X

End Function

Public Sub cmdBalance_Ok_Print(lBalance_Ok_Param As String, lMsg As String)
Dim wIdFile As Integer, wFileName As String
Dim V
V = Null

If chkBalance_Pays Then
    Mid$(lBalance_Ok_Param, 11, 1) = "1"     ' balance par pays
    fgSelect_Sort1 = 9: fgSelect_Sort2 = 9 ' TRI  Pays / PCI/ Dev /Compte
Else
    fgSelect_Sort1 = 2: fgSelect_Sort2 = 2 ' TRI Dev / PCI / Compte
End If
fgSelect_SortAD = 6
fgSelect_SortX fgSelect_Sort1


If chkBalance_CSV = "1" Then
    If Mid$(txtBalance_CSV_Folder, Len(txtBalance_CSV_Folder), 1) <> "\" Then txtBalance_CSV_Folder = txtBalance_CSV_Folder & "\"
    wFileName = txtBalance_CSV_Folder & txtBalance_CSV_FileName
    wIdFile = 0
    Mid$(lBalance_Ok_Param, 6, 1) = chkBalance_CSV
    V = File_Export_Monitor("Output", wIdFile, wFileName)
    Mid$(lBalance_Ok_Param, 7, 3) = Format$(wIdFile, "000")
    Call File_Export_Monitor("Print", wIdFile, "COMPTEDEV;COMPTEOBL;COMPTECOM;COMPTEINT;DB_dev;CR_dev;DB_eur;CR_eur;COMPTEOUV;COMPTEFON;SOLDEDMO")

End If

If IsNull(V) Then
    prtSAB_Balance_Monitor lBalance_Ok_Param, xAmjMin, xAmjMax, fgSelect, marrYBIACPT0(), marrYBIACPT0_Nb, lMsg, wYSTOMON(), wDORCPTDMV()
    If chkBalance_CSV = "1" Then V = File_Export_Monitor("Close", wIdFile, wFileName)
End If

fgSelect.Visible = True
fraBalance.Visible = False
Me.Show

End Sub
Public Sub cmdBalance_Ok_Stock()
Dim X As String, K As Long, lstW_Index As Long, I As Long, mService As String
Dim blnOk As Boolean, blnAdd As Boolean
Dim xWhere As String, nbDossier As Long, X20 As String
Dim wService_Name As String, wService_Référence As String, wService_Printer As String, wService_Sxx As String
Dim iDevise As Integer
Dim xSQL As String
Dim wBalance_Ok_Param As String
Dim curSolde As Currency
Dim V
'===========================================================================================
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : chargement des services/comptes")

ZCOMREF0_SQL_ODBC
'======================================== Compte sans service => Responsable client ===================
mService = ""
For I = 1 To marrYBIACPT0_Nb
    If Trim(marrYBIACPT0(I).CLIENARES) = "R80" Then
        X = marrYBIACPT0(I).PLANCOPRO
        If X = "CAV" Or X = "LIE" Or X = "LOR" Then wCOMREFCOR(I) = "G7"            '$JPL 2014-11-12
    End If
    If wCOMREFCOR(I) = "" Then
        X = marrYBIACPT0(I).PLANCOPRO
        If X = "CAV" Or X = "LIE" Or X = "LOR" Then
            wCOMREFCOR(I) = Trim(marrYBIACPT0(I).CLIENARES)
            If mService <> wCOMREFCOR(I) Then
                mService = wCOMREFCOR(I)
                blnAdd = True
                For K = 1 To arrService_Nb
                    If mService = arrService(K) Then blnAdd = False: Exit For
                Next K
                If blnAdd Then
                    arrService_Nb = arrService_Nb + 1
                    If arrService_Nb >= UBound(arrService) Then ReDim Preserve arrService(arrService_Nb + 10)
                    arrService(arrService_Nb) = mService
                End If
            End If
        End If
    End If
Next I
'=====================================================================================================
' - 2 : compte ne devant pas avoir de contrats associés
' - 1 : compte devant avoir des contrats , mais aucun contrat associé
' >= 0 : compte avec contrats , vérifier balance = stock

ReDim wYSTOMON(marrYBIACPT0_Nb)
For I = 1 To marrYBIACPT0_Nb
    wYSTOMON(I) = -2
Next I
'===========================================================================================
If Not blnBalance_Service_Stock Then
    ZSOLDE0_SQL_ODBC
Else
    Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : chargement OPENAT_PCI")
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID  = 'OPENAT_PCI'"
    Set rsSab = cnsab.Execute(xSQL)
    
    Do While Not rsSab.EOF
        X = Mid$(rsSab("BIATABTXT"), 4, 5)
        For I = 1 To marrYBIACPT0_Nb
            If X = Mid$(marrYBIACPT0(I).COMPTEOBL, 1, 5) Then
                 wYSTOMON(I) = -1
            End If
        Next I
        rsSab.MoveNext
    Loop
    '===========================================================================================
    Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : chargement du stock")
    xWhere = ""
    Call YBIASTO0_Sql(xWhere, nbDossier, arrYBIASTO0(), stockYBIACPT0(), stockCompte_Nb)
    
    Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : rapprochement")
    
    For K = 1 To stockCompte_Nb
        X20 = stockYBIACPT0(K).COMPTECOM
        For I = 1 To marrYBIACPT0_Nb
            If X20 = marrYBIACPT0(I).COMPTECOM Then
                If wYSTOMON(I) > -1 Then MsgBox marrYBIACPT0(I).COMPTECOM, vbExclamation, "stock déjà affecté  " & X20
                wYSTOMON(I) = arrYBIASTO0(K).YSTOMON
                Exit For
            End If
        Next I
    Next K
    '===========================================================================================
    ' recherche date du dernier mouvement
    Call ZDORCPT_SQL_ODBC
End If
'===========================================================================================
arrService_Balance_Cumul_Z
lstW.Clear
For K = 1 To arrService_Nb
    lstW.AddItem arrService(K)
Next K
For lstW_Index = 0 To lstW.ListCount - 1
    K = lstW_Index + 1
    lstW.ListIndex = lstW_Index
    mService = Trim(lstW.Text)
    Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : " & mService)
    If Mid$(mService, 1, 1) = "R" Then
        Call rsYBIATAB0_Read("RESPONSABLE", mService, "", wService_Name)
        ' wService_Name = YBIATAB0_Sql_Responsable(mService, cnSab, rsSab)
        wService_Name = Mid$(wService_Name, 34, 12)
        Select Case Mid$(mService, 1, 3)
            Case "R60":        wService_Printer = "DRH"
            Case "R80":         wService_Printer = "DER"
           Case Else: wService_Printer = "DCOM"
           
        End Select
        wService_Sxx = Table_Unit_SSI("", Trim(wService_Printer))
    Else
        V = rsElpTable_Read("SAb_Param", "Compte_Unit", mService, wService_Name, wService_Sxx)

        If Not IsNull(V) Then
            wService_Name = mService & " : ?????"
            wService_Printer = "CPT"
            wService_Sxx = "S60"
        Else
            wService_Printer = Table_Unit_SSI("S", Trim(wService_Sxx))
        End If
    End If
    wService_Référence = mService & " - " & Trim(wService_Name)
    arrService_Balance_Cumul(K, 0).Id = wService_Référence
    iDevise = 0

    'If blnService_Printer Then Printer_Set wService_Printer
    

    fgSelect_Reset
    fgSelect.Rows = 1
    fgSelect.FormatString = fgSelect_FormatString
    fgSelect.Visible = False

    For I = 1 To marrYBIACPT0_Nb
       If mService = wCOMREFCOR(I) Then
           blnOk = True
            xYBIACPT0 = marrYBIACPT0(I)
            If blnBalance_Service_Stock Then
                curSolde = xYBIACPT0.SOLDECEN
            Else
                curSolde = curBalance(I)
            End If
            
            If curSolde = 0 Then blnOk = False
            If xYBIACPT0.COMPTEFON = "4" And curSolde = 0 Then blnOk = False
            If blnOk Then
                If fctUser_Classe_Aut(xYBIACPT0.COMPTECLA) Then
                    fgSelect_DisplayLine I
'===========================================================================================
                    If arrService_Balance_Cumul(K, iDevise).Dev <> xYBIACPT0.COMPTEDEV Then
                        For iDevise = 1 To arrDevise_Nb
                            If xYBIACPT0.COMPTEDEV = arrDevise(iDevise) Then
                                arrService_Balance_Cumul(K, iDevise).Id = wService_Référence
                                Exit For
                            End If
                        Next iDevise
                    End If
                    curX = Abs(curSolde)
                    If Mid$(xYBIACPT0.COMPTEOBL, 1, 1) <> "9" Then
                        arrService_Balance_Cumul(K, iDevise).Bilan_Nb = arrService_Balance_Cumul(K, iDevise).Bilan_Nb + 1
                        If curSolde > 0 Then
                            arrService_Balance_Cumul(K, iDevise).Bilan_DB = arrService_Balance_Cumul(K, iDevise).Bilan_DB + curX
                        Else
                            arrService_Balance_Cumul(K, iDevise).Bilan_CR = arrService_Balance_Cumul(K, iDevise).Bilan_CR + curX
                        End If
                    Else
                         arrService_Balance_Cumul(K, iDevise).HorsBilan_Nb = arrService_Balance_Cumul(K, iDevise).HorsBilan_Nb + 1
                        If curSolde > 0 Then
                            arrService_Balance_Cumul(K, iDevise).HorsBilan_DB = arrService_Balance_Cumul(K, iDevise).HorsBilan_DB + curX
                        Else
                            arrService_Balance_Cumul(K, iDevise).HorsBilan_CR = arrService_Balance_Cumul(K, iDevise).HorsBilan_CR + curX
                        End If
                   End If
'===========================================================================================
                End If
            End If
       End If
    Next I

    
    fgSelect.Visible = True
    Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : " & mService & " : " & fgSelect.Rows - 1)
    
    If fgSelect.Rows > 1 And blnBalance_Stock_détail Then
        If mService = "" Then mService = "G0"

        Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : impression " & mService)
        
        If blnAuto Then
            'X = Table_Unit_SSI("", Trim(wService_Printer))
            Call frmElpPrt.prtIMP_PDF_NoPaper_Init(wService_Sxx, "BIA-BAL-Stock_" & mService, "Archive")
        Else
            Call frmElpPrt.prtIMP_PDF_NoPaper_Init(wService_Sxx, "BIA-BAL-Stock_" & mService, "PROD")
        End If
        
        optBalance_YSOLDE0_J = True
        chkBalance_Détail = "1"
        chkBalance_Récap = "0"
        chkBalance_Récap_Bilan = "0"
        chkBalance_CSV = "0"
        chkBalance_Pays = "0"
        chkBalance_Compte_Soldé = "1"
        wBalance_Ok_Param = cmdBalance_Ok_Param
        Mid$(wBalance_Ok_Param, 1, 1) = "S"                      ' BALANCE avec contrôle STOCK
        cmdBalance_Ok_Print wBalance_Ok_Param, wService_Référence
        If blnService_Printer Then Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("", "", "Balance / Stock - " & mService)

    End If
Next lstW_Index

End Sub


Public Sub arrService_Balance_Cumul_Z()
Dim I As Integer, K As Integer
ReDim arrService_Balance_Cumul(arrService_Nb, arrDevise_Nb)
    
For I = 0 To arrService_Nb
    For K = 0 To arrDevise_Nb
        arrService_Balance_Cumul(I, K).Id = arrService(I)
        arrService_Balance_Cumul(I, K).Dev = arrDevise(K)
        arrService_Balance_Cumul(I, K).Bilan_Nb = 0
        arrService_Balance_Cumul(I, K).Bilan_DB = 0
        arrService_Balance_Cumul(I, K).Bilan_CR = 0
        arrService_Balance_Cumul(I, K).HorsBilan_Nb = 0
        arrService_Balance_Cumul(I, K).HorsBilan_DB = 0
        arrService_Balance_Cumul(I, K).HorsBilan_CR = 0
    Next K
Next I
End Sub


Public Sub ZDORCPT_SQL_ODBC()
'===========================================================================================
' recherche date du dernier mouvement
' pour les CAV et LOR rechercher dans le fichier ZDORCPT0
' corriger certaines anomalies  de DORCPTDMV
Dim xSQL As String
Dim X20 As String
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : chargement ZDORCPT0")

ReDim wDORCPTDMV(marrYBIACPT0_Nb)
For I = 1 To marrYBIACPT0_Nb
    wDORCPTDMV(I) = 0
Next I

xSQL = "select DORCPTCOM,DORCPTDMV from " & paramIBM_Library_SAB & ".ZDORCPT0  order by DORCPTCOM"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    X20 = rsSab("DORCPTCOM")
    For I = 1 To marrYBIACPT0_Nb
        If X20 = marrYBIACPT0(I).COMPTECOM Then
            wDORCPTDMV(I) = rsSab("DORCPTDMV")
            If wDORCPTDMV(I) > marrYBIACPT0(I).SOLDEDMO Then wDORCPTDMV(I) = marrYBIACPT0(I).SOLDEDMO
            
            Exit For
        End If
    Next I
    rsSab.MoveNext
Loop

blnZDORCPT_SQL_ODBC = True
End Sub
Public Sub ZSOLDE0_SQL_ODBC()
'===========================================================================================
Dim xSQL As String
Dim X20 As String
Dim kSolde As Integer

kSolde = 1 + Val(Mid$(YBIATAB0_DATE_CPT_MP1, 5, 2))

Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : chargement ZSOLDE0")

ReDim curBalance(marrYBIACPT0_Nb)
For I = 1 To marrYBIACPT0_Nb
    curBalance(I) = 0
Next I

xSQL = "select * from " & paramIBM_Library_SAB & ".ZSOLDE0  order by SOLDECOM"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    X20 = rsSab("SOLDECOM")
    For I = 1 To marrYBIACPT0_Nb
        If X20 = marrYBIACPT0(I).COMPTECOM Then
            Call rsZSOLDE0_SOLDEC(rsSab, kSolde, curBalance(I))
            'Select Case kSolde
            '        Case 1: curBalance(I) = rsSab("SOLDEC01")
            '        Case 2: curBalance(I) = rsSab("SOLDEC02")
            '        Case 3: curBalance(I) = rsSab("SOLDEC03")
            '        Case 4: curBalance(I) = rsSab("SOLDEC04")
            '        Case 5: curBalance(I) = rsSab("SOLDEC05")
            '        Case 6: curBalance(I) = rsSab("SOLDEC06")
            '        Case 7: curBalance(I) = rsSab("SOLDEC07")
            '        Case 8: curBalance(I) = rsSab("SOLDEC08")
            '        Case 9: curBalance(I) = rsSab("SOLDEC09")
            '        Case 10: curBalance(I) = rsSab("SOLDEC10")
            '        Case 11: curBalance(I) = rsSab("SOLDEC11")
            '        Case 12: curBalance(I) = rsSab("SOLDEC12")
            'End Select
            
            Exit For
        End If
    Next I
    rsSab.MoveNext
Loop

End Sub


Public Sub cmdSelect_Compte_nonAffecté()
Dim X As String, wNb As Long
'===========================================================================================
Call lstErr_Clear(Me.lstErr, Me.cmdContext, ": chargement des services/comptes")

'$JPL 2010-08-26
'==========================================================================
wNb = 0
For I = 1 To marrYBIACPT0_Nb
    If marrYBIACPT0(I).COMREFCOR = "G?" Then
        wNb = wNb + 1
        marrYBIACPT0(wNb) = marrYBIACPT0(I)
    End If
Next I
marrYBIACPT0_Nb = wNb
Exit Sub
'==========================================================================

ZCOMREF0_SQL_ODBC
'======================================== Compte sans service => Responsable client ===================
For I = 1 To marrYBIACPT0_Nb
    If wCOMREFCOR(I) = "" Then
        X = marrYBIACPT0(I).PLANCOPRO
        If X = "CAV" Or X = "LIE" Or X = "LOR" Then
            wCOMREFCOR(I) = Trim(marrYBIACPT0(I).CLIENARES)
        End If
    End If
Next I
'=====================================================================================================
wNb = 0
For I = 1 To marrYBIACPT0_Nb
    If wCOMREFCOR(I) = "" Then
        wNb = wNb + 1
        marrYBIACPT0(wNb) = marrYBIACPT0(I)
    End If
Next I
marrYBIACPT0_Nb = wNb

'=====================================================================================================

End Sub


Public Sub cmdSelect_SQL_Exportation_Liste_Detail_T2()
Dim K As Long, X As String, I1 As Long, I2 As Long
Dim wRow As Long, wCol As Long
Dim wCOMPTEOBL As String, wCLIENACLI As String
Dim wCOMPTEOBL_Row As Long, WCLIENACLI_Row As Long
Dim mCOMPTEOBL As String
On Error GoTo Error_Handler
'===================================================================================
Set wsExcel = wbExcel.Sheets(2)

wRow = 1
mXls1_Row_T = 0
For K = 0 To lstW.ListCount - 1
    lstW.ListIndex = K
    X = lstW.Text
    I1 = InStr(X, "|"): wCOMPTEOBL = Mid$(X, 1, I1 - 1)
    I2 = I1 + 1
    I1 = InStr(I2, X, "|"): wCLIENACLI = Mid$(X, I2, I1 - I2)
    I2 = I1 + 1
    I1 = InStr(I2, X, "|"): wCOMPTEOBL_Row = Val(Mid$(X, I2, I1 - I2))
    WCLIENACLI_Row = Val(Mid$(X, I1 + 1, Len(X) - I1))


    If mCOMPTEOBL <> wCOMPTEOBL Then
    
        If mXls1_Row_T > 0 Then
            X = alfBilan_DB & mXls1_Row_T + 1 & ":" & alfBilan_DB & wRow
            wsExcel.Cells(mXls1_Row_T, colBilan_DB).FormulaLocal = "=SI(SOMME(" & X & ")= 0;" & Asc34 & Asc34 & ";SOMME(" & X & "))"
            wsExcel.Cells(mXls1_Row_T, colBilan_DB).Font.Bold = True
             
            X = alfBilan_CR & mXls1_Row_T + 1 & ":" & alfBilan_CR & wRow
            wsExcel.Cells(mXls1_Row_T, colBilan_CR).FormulaLocal = "=SI(SOMME(" & X & ")= 0;" & Asc34 & Asc34 & ";SOMME(" & X & "))"
            wsExcel.Cells(mXls1_Row_T, colBilan_CR).Font.Bold = True
            
            X = alfHors_Bilan_DB & mXls1_Row_T + 1 & ":" & alfHors_Bilan_DB & wRow
            wsExcel.Cells(mXls1_Row_T, colHors_Bilan_DB).FormulaLocal = "=SI(SOMME(" & X & ")= 0;" & Asc34 & Asc34 & ";SOMME(" & X & "))"
            wsExcel.Cells(mXls1_Row_T, colHors_Bilan_DB).Font.Bold = True
             
            X = alfHors_Bilan_CR & mXls1_Row_T + 1 & ":" & alfHors_Bilan_CR & wRow
            wsExcel.Cells(mXls1_Row_T, colHors_Bilan_CR).FormulaLocal = "=SI(SOMME(" & X & ")= 0;" & Asc34 & Asc34 & ";SOMME(" & X & "))"
            wsExcel.Cells(mXls1_Row_T, colHors_Bilan_CR).Font.Bold = True
        End If
        mCOMPTEOBL = wCOMPTEOBL
        wRow = wRow + 1
        wsExcel.Cells(wRow, 1) = wCOMPTEOBL
        wsExcel.Cells(wRow, 1).Font.Bold = True
        wsExcel.Cells(wRow, 1).Font.Color = vbBlue
        wsExcel.Cells(wRow, 1).Font.Size = 7
        wsExcel.Cells(wRow, 2) = "='CPT-PCI'!B" & wCOMPTEOBL_Row
        wsExcel.Cells(wRow, 2).Font.Bold = True
        wsExcel.Cells(wRow, 2).Font.Color = vbBlue
        wsExcel.Cells(wRow, 2).Font.Size = 7
        If Mid$(wCOMPTEOBL, 1, 2) = "98" Then
            For I1 = 1 To mXls1_Col: wsExcel.Cells(wRow, I1).Interior.Color = RGB(220, 220, 220): Next I1
        Else
            For I1 = 1 To mXls1_Col: wsExcel.Cells(wRow, I1).Interior.Color = mColor_Y1: Next I1   ' RGB(190, 230, 255)
        End If
        mXls1_Row_T = wRow
    End If
    wRow = wRow + 1
    wsExcel.Cells(wRow, 1) = wCLIENACLI
    wsExcel.Cells(wRow, 2) = "='CPT-PCI'!B" & WCLIENACLI_Row
    For I1 = 3 To mXls1_Col
        X = "'CPT-PCI'!" & Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", I1, 1) & wCOMPTEOBL_Row
       wsExcel.Cells(wRow, I1).FormulaLocal = "=SI(" & X & "=" & Asc34 & Asc34 & ";" & Asc34 & Asc34 & ";" & X & ")"
        'wsExcel.Cells(wRow, I1) = "='CPT-PCI'!" & Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", I1, 1) & wCOMPTEOBL_Row
    Next I1
    For I1 = 0 To 3
        wsExcel.Cells(wRow, mXls1_Col - I1).Interior.Color = mColor_Y0
    Next I1

Next K

If mXls1_Row_T > 0 Then
    X = alfBilan_DB & mXls1_Row_T + 1 & ":" & alfBilan_DB & wRow
    wsExcel.Cells(mXls1_Row_T, colBilan_DB).FormulaLocal = "=SI(SOMME(" & X & ")= 0;" & Asc34 & Asc34 & ";SOMME(" & X & "))"
    wsExcel.Cells(mXls1_Row_T, colBilan_DB).Font.Bold = True
     
    X = alfBilan_CR & mXls1_Row_T + 1 & ":" & alfBilan_CR & wRow
    wsExcel.Cells(mXls1_Row_T, colBilan_CR).FormulaLocal = "=SI(SOMME(" & X & ")= 0;" & Asc34 & Asc34 & ";SOMME(" & X & "))"
    wsExcel.Cells(mXls1_Row_T, colBilan_CR).Font.Bold = True
    
    X = alfHors_Bilan_DB & mXls1_Row_T + 1 & ":" & alfHors_Bilan_DB & wRow
    wsExcel.Cells(mXls1_Row_T, colHors_Bilan_DB).FormulaLocal = "=SI(SOMME(" & X & ")= 0;" & Asc34 & Asc34 & ";SOMME(" & X & "))"
    wsExcel.Cells(mXls1_Row_T, colHors_Bilan_DB).Font.Bold = True
     
    X = alfHors_Bilan_CR & mXls1_Row_T + 1 & ":" & alfHors_Bilan_CR & wRow
    wsExcel.Cells(mXls1_Row_T, colHors_Bilan_CR).FormulaLocal = "=SI(SOMME(" & X & ")= 0;" & Asc34 & Asc34 & ";SOMME(" & X & "))"
    wsExcel.Cells(mXls1_Row_T, colHors_Bilan_CR).Font.Bold = True
End If

Exit Sub

Error_Handler:

End Sub

Public Sub fgSelect_Display_Groupe(lCLIGRPREG As String)
Dim xSQL As String, X As String

X = Format(lCLIGRPREG, "0000000")

xSQL = "select count(*) from " & paramIBM_Library_SAB & ".ZCLIGRP0" _
& " where CLIGRPREG = '" & X & "'"
Set rsSab = cnsab.Execute(xSQL)
ReDim arrGroupe_Filtre(rsSab(0) + 1)

xSQL = "select CLIGRPCLI from " & paramIBM_Library_SAB & ".ZCLIGRP0" _
& " where CLIGRPREG = '" & X & "'" _
& "  order by CLIGRPCLI"
 Set rsSab = cnsab.Execute(xSQL)
 
 arrGroupe_Filtre_Nb = 0
 
 Do While Not rsSab.EOF
     arrGroupe_Filtre_Nb = arrGroupe_Filtre_Nb + 1
      arrGroupe_Filtre(arrGroupe_Filtre_Nb) = rsSab("CLIGRPCLI")
     rsSab.MoveNext
 Loop


End Sub

Public Sub arrService_Load()

If arrService_Code_Nb = 0 Then
    ReDim arrService_Code(lstService.ListCount), arrService_Lib(lstService.ListCount)
   
    Dim X As String
    X = "select * from ElpTable where SNN = 0" _
        & " and id = 'SAb_Param'" _
        & " and K1 = 'Compte_Unit' order by K2"
        
    Set rsMDB = cnMDB.Execute(X)
    Do While Not rsMDB.EOF
            arrService_Code_Nb = arrService_Code_Nb + 1
            arrService_Code(arrService_Code_Nb) = Trim(rsMDB("K2"))
            arrService_Lib(arrService_Code_Nb) = Trim(rsMDB("Name"))
        rsMDB.MoveNext
    Loop

End If

End Sub

Public Sub cmdRCOM_AUT()

Dim rsX As Recordset, xSQL As String, wAUTE1IRES As String, K As Integer, xDest As String, xFile As String
Me.Enabled = False: Me.MousePointer = vbHourglass
'_________________________________________________________________________________________________
If Mid$(YBIATAB0_DATE_CPT_J, 1, 6) <> Mid$(YBIATAB0_DATE_CPT_JS1, 1, 6) Then
    mAUTE1IDAF_Min8 = dateElp("MoisAdd", -1, DSys)
    mAUTE1IDAF_Min7 = mAUTE1IDAF_Min8 - 19000000
    mAUTE1IDAF_Max8 = dateElp("MoisAdd", 1, DSys)
    mAUTE1IDAF_Max7 = mAUTE1IDAF_Max8 - 19000000
    
'_________________________________________________________________________________________________
'2012-06-18 fin de mois => Contrôle Permanent"

        txtSelect_CLIENARES = ""
        Call lstErr_Clear(Me.lstErr, Me.cmdContext, "cmdRCOM_AUT Dépassement global")
        
        Call fgYAUTE1I0_Display("Echeance")
        
        If fgYAUTE1I0.Rows > 1 Then
           ' V = mailAdresse_Production_Control("LEGOUARD", xDest)
           ' If IsNull(V) Then Call mnuPrint2_Mail_Send(xDest)
            xFile = "C:\Temp\AUT_ECH.xlsx"
            paramEditionNoPaper_Auto_PgmName = "BIA-AUT-ECH"
            Call MSflexGrid_Excel(xFile, "AUT_ECH", lblYAUTE1I0, fgYAUTE1I0, fgYAUTE1I0.Cols - 1)
            
            Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S51", xFile, "Archive", "BIA-AUT-ECH")
            Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("", "", "")

        End If

        Call fgYAUTE1I0_Display("Dépassement")
        
        If fgYAUTE1I0.Rows > 1 Then
            'V = mailAdresse_Production_Control("LEGOUARD", xDest)
            'If IsNull(V) Then Call mnuPrint2_Mail_Send(xDest)
            xFile = "C:\Temp\AUT_DEP.xlsx"
            paramEditionNoPaper_Auto_PgmName = "BIA-AUT-DEP"
            Call MSflexGrid_Excel(xFile, "AUT_DEP", lblYAUTE1I0, fgYAUTE1I0, fgYAUTE1I0.Cols - 1)
            
            Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S51", xFile, "Archive", "BIA-AUT-DEP")
            Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("", "", "")
      End If

'_________________________________________________________________________________________________
    xSQL = "select distinct AUTE1IRES  from " & paramIBM_Library_SAB & ".ZAUTE1I0" _
         & " where AUTE1IETA = 1 " _
         & " and AUTE1IDAF >= " & mAUTE1IDAF_Min7 & " and AUTE1IDAF <= " & mAUTE1IDAF_Max7 _
         & " and   AUTE1ITYP = '8'" _
         & " order by AUTE1IRES"
    Set rsX = cnsab.Execute(xSQL)
    
    Do While Not rsX.EOF
        wAUTE1IRES = rsX("AUTE1IRES")
        txtSelect_CLIENARES = wAUTE1IRES
        Call lstErr_Clear(Me.lstErr, Me.cmdContext, "cmdRCOM_AUT Dépassement: " & wAUTE1IRES)
        
        Call fgYAUTE1I0_Display("Echeance")
        
        If fgYAUTE1I0.Rows > 1 Then
            K = Val(Mid$(wAUTE1IRES, 2, 2))
            
            V = mailAdresse_Production_Control(arrBIA_RCOM_Lib(K), xDest)
            If Not IsNull(V) Then V = mailAdresse_Production_Control(arrBIA_RCOM_Lib(0), xDest)
            
            If IsNull(V) Then
                'Call mnuPrint2_Mail_Send(xDest)
                xFile = "C:\Temp\AUT_ECH_" & wAUTE1IRES & ".xlsx"
                paramEditionNoPaper_Auto_PgmName = "BIA-AUT-ECH"
                Call MSflexGrid_Excel(xFile, "AUT_DEP", lblYAUTE1I0, fgYAUTE1I0, fgYAUTE1I0.Cols - 1)
                
                Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S41", xFile, "Prod", "BIA-AUT-ECH-" & wAUTE1IRES)
                
                X = "Veuillez trouver ci-joint la liste des autorisations dont l'échéance est comprise entre le " & dateImp10_S(mAUTE1IDAF_Min8) & " et le " & dateImp10_S(mAUTE1IDAF_Max8)

                Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("", xDest, X)
            End If
        End If
        
        rsX.MoveNext
    Loop
End If
'_________________________________________________________________________________________________

xSQL = "select distinct AUTE1IRES  from " & paramIBM_Library_SAB & ".ZAUTE1I0 " _
     & " where AUTE1IETA = 1 " _
     & " and AUTE1IMTD <> 0" _
     & " order by AUTE1IRES"

Set rsX = cnsab.Execute(xSQL)

Do While Not rsX.EOF
    wAUTE1IRES = rsX("AUTE1IRES")
    txtSelect_CLIENARES = wAUTE1IRES
    Call lstErr_Clear(Me.lstErr, Me.cmdContext, "cmdRCOM_AUT Dépassement: " & wAUTE1IRES)
    
    Call fgYAUTE1I0_Display("Dépassement")
    
    If fgYAUTE1I0.Rows > 1 Then
        K = Val(Mid$(wAUTE1IRES, 2, 2))
        
        V = mailAdresse_Production_Control(arrBIA_RCOM_Lib(K), xDest)
        If Not IsNull(V) Then V = mailAdresse_Production_Control(arrBIA_RCOM_Lib(0), xDest)
        
        If IsNull(V) Then
            'Call mnuPrint2_Mail_Send(xDest)
                xFile = "C:\Temp\AUT_DEP_" & wAUTE1IRES & ".xlsx"
                paramEditionNoPaper_Auto_PgmName = "BIA-AUT-DEP"
                Call MSflexGrid_Excel(xFile, "AUT_DEP", lblYAUTE1I0, fgYAUTE1I0, fgYAUTE1I0.Cols - 1)
                
                Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S41", xFile, "Prod", "BIA-AUT-DEP-" & wAUTE1IRES)
                
                X = "Veuillez trouver ci-joint l'état de dépassement en date du  " & dateImp10_S(YBIATAB0_DATE_CPT_J)

                Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("", xDest, X)
            End If
    End If
    
    rsX.MoveNext
Loop

'____________________________________________________________________________________________________
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0
End Sub
Public Sub cmdCPT_OD()

Dim xFile As String
Me.Enabled = False: Me.MousePointer = vbHourglass

'_________________________________________________________________________________________________

Call lstErr_Clear(Me.lstErr, Me.cmdContext, "cmdCPT_OD")

mnuSelect_CPT_OD_Click

If fgYAUTE1I0.Rows > 1 Then
    'xDest = srvSendMail.Exchange_Distribution("CPT", "@CPT_OD")

    'Call mnuPrint2_Mail_Send(xDest)
    
    xFile = "C:\Temp\CPT_OD.xlsx"
    paramEditionNoPaper_Auto_PgmName = "BIA-SAB-CPT-OD"
    Call MSflexGrid_Excel(xFile, "SAB_CPT_OD ", lblYAUTE1I0, fgYAUTE1I0, fgYAUTE1I0.Cols - 1)
    
    Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S60", xFile, "Archive", "BIA-SAB-CPT-OD")
    Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("CPT", "@CPT_OD", "")

End If


'____________________________________________________________________________________________________
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0
End Sub


Public Sub mnuPrint2_Mail_Send(lDest As String)
Dim X As String, xCLIENARES As String, blnCPT_OD As Boolean

Select Case mfgYAUTE1I0_Fct
    Case "Echeance"
        X = "Veuillez trouver ci-après la liste des autorisations dont l'échéance est comprise entre le " & dateImp10_S(mAUTE1IDAF_Min8) & " et le " & dateImp10_S(mAUTE1IDAF_Max8)

    Case "Dépassement"
        X = "Veuillez trouver ci-après l'état de dépassement en date du  " & dateImp10_S(YBIATAB0_DATE_CPT_J)
    Case "CPT_OD"
        blnCPT_OD = True
        X = lblYAUTE1I0
    Case Else
        X = mfgYAUTE1I0_Fct
End Select

If blnCPT_OD Then
    Call MSFlexGrid_SendMail(lDest, "SAB_CPT_OD", lblYAUTE1I0, X, fgYAUTE1I0, fgYAUTE1I0.Cols - 1)
Else
    xCLIENARES = Trim(txtSelect_CLIENARES)
    If xCLIENARES <> "" Then
        X = X & "<BR>concernant le responsable " & xCLIENARES & " - " & arrBIA_RCOM_Lib(Val(Mid$(xCLIENARES, 2, 2)))
    End If
    
    Call MSFlexGrid_SendMail(lDest, "SAB_Aut_Alerte", lblYAUTE1I0, X, fgYAUTE1I0, fgYAUTE1I0.Cols - 1)
End If

End Sub

Public Sub mnuSelect_ENG_Solde(lMOUVEMCOM As String, lAMJMin As String, lBIAMVTSD0 As Currency)
Dim X As String


X = "select BIAMVTSD0 from " & paramIBM_Library_SABSPE & ".YBIAMVTH" _
  & " where MOUVEMCOM = '" & lMOUVEMCOM & "'" _
     & " and MOUVEMDTR >= " & lAMJMin - 19000000 _
     & " order by MOUVEMDTR "

Set rsSabX = cnsab.Execute(X)

If Not rsSabX.EOF Then
        lBIAMVTSD0 = rsSabX("BIAMVTSD0")
End If

End Sub
