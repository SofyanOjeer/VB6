VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSAB_Balance 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_Balance"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   Icon            =   "SAB_Balance.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6270
   ScaleWidth      =   8685
   Begin VB.ListBox lstW 
      Height          =   255
      Left            =   3120
      Sorted          =   -1  'True
      TabIndex        =   126
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstService 
      Height          =   645
      Left            =   360
      TabIndex        =   114
      Top             =   2040
      Width           =   2505
   End
   Begin VB.Frame fraSelect_Options 
      Height          =   3525
      Left            =   0
      TabIndex        =   86
      Top             =   2760
      Width           =   8475
      Begin VB.CheckBox chkSelect_COMPTEOUV 
         BackColor       =   &H80000004&
         Caption         =   "Date création >="
         Height          =   285
         Left            =   5520
         TabIndex        =   113
         Top             =   3120
         Width           =   1545
      End
      Begin VB.CheckBox chkSelect_Annulé 
         BackColor       =   &H80000004&
         Caption         =   "exclure comptes annulés"
         Height          =   285
         Left            =   2880
         TabIndex        =   111
         Top             =   2280
         Width           =   2235
      End
      Begin VB.CheckBox chkSelect_HB 
         BackColor       =   &H80000004&
         Caption         =   "exclure classe 9"
         Height          =   285
         Left            =   2880
         TabIndex        =   110
         Top             =   1800
         Width           =   2000
      End
      Begin VB.ComboBox cboPCEC 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   103
         Text            =   "PCEC"
         Top             =   1800
         Width           =   1300
      End
      Begin VB.ComboBox cboDevise 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   102
         Text            =   "Devise"
         Top             =   840
         Width           =   1300
      End
      Begin VB.ComboBox cboPLANCOPRO 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   101
         Text            =   "prod"
         Top             =   1320
         Width           =   1300
      End
      Begin VB.CheckBox chkSelect 
         Alignment       =   1  'Right Justify
         Caption         =   "Tous les comptes"
         Height          =   285
         Left            =   120
         TabIndex        =   100
         Top             =   360
         Width           =   1875
      End
      Begin VB.CheckBox chkSelect_Résidence 
         Caption         =   "tri RNR(CHA/PRO)"
         Height          =   300
         Left            =   5520
         TabIndex        =   99
         Top             =   2160
         Width           =   2000
      End
      Begin VB.CheckBox chkSelect_SoldeZ 
         BackColor       =   &H80000004&
         Caption         =   "exclure Solde=0"
         Height          =   285
         Left            =   2880
         TabIndex        =   98
         Top             =   360
         Value           =   1  'Checked
         Width           =   2000
      End
      Begin VB.CheckBox chkSelect_SoldeDb 
         BackColor       =   &H80000004&
         Caption         =   "exclure Débiteur"
         Height          =   255
         Left            =   2880
         TabIndex        =   97
         Top             =   720
         Width           =   2000
      End
      Begin VB.CheckBox chkSelect_SoldeCr 
         BackColor       =   &H80000004&
         Caption         =   "exclure Créditeur"
         Height          =   285
         Left            =   2880
         TabIndex        =   96
         Top             =   1080
         Width           =   2000
      End
      Begin VB.Frame fraSelect_SOLDEDMO 
         Height          =   1335
         Left            =   5400
         TabIndex        =   92
         Top             =   720
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
            Format          =   19595267
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
      End
      Begin VB.CheckBox chkSelect_SOLDEDMO 
         BackColor       =   &H80000004&
         Caption         =   "date dernier mouvement"
         Height          =   285
         Left            =   5520
         TabIndex        =   91
         Top             =   360
         Width           =   2025
      End
      Begin VB.CheckBox chkSelect_MOUVEMDCO 
         Caption         =   "select DCO"
         Height          =   345
         Left            =   5520
         TabIndex        =   90
         Top             =   2640
         Width           =   1110
      End
      Begin VB.TextBox txtSelect_CLIENARES 
         Height          =   285
         Left            =   840
         TabIndex        =   89
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtSelect_CLIENARSD 
         Height          =   285
         Left            =   840
         TabIndex        =   88
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtSelect_COMPTECLA 
         Height          =   300
         Left            =   840
         TabIndex        =   87
         Top             =   3000
         Width           =   615
      End
      Begin MSComCtl2.DTPicker txtSelect_COMPTEOUV 
         Height          =   300
         Left            =   7200
         TabIndex        =   112
         Top             =   3120
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
         Format          =   19595267
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   -328351
      End
      Begin VB.Label lblSelect_Devise 
         Caption         =   "Devise"
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   840
         Width           =   600
      End
      Begin VB.Label lblSelect_PLANCOPRO 
         Caption         =   "Produit"
         Height          =   270
         Left            =   120
         TabIndex        =   108
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label lblSelect_PCEC 
         Caption         =   "PCEC"
         Height          =   210
         Left            =   120
         TabIndex        =   107
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label lblSelect_CLIENARES 
         Caption         =   "Resp."
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblSelect_CLIENARSD 
         Caption         =   "Pays R"
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblSelect_COMPTECLA 
         Caption         =   "Classe S"
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   3120
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
      Left            =   1560
      TabIndex        =   26
      Top             =   1920
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
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   3915
      TabIndex        =   6
      Top             =   0
      Width           =   4260
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5730
      Left            =   0
      TabIndex        =   4
      Top             =   555
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   10107
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Comptes"
      TabPicture(0)   =   "SAB_Balance.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mouvements"
      TabPicture(1)   =   "SAB_Balance.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraMvt"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Autorisations"
      TabPicture(2)   =   "SAB_Balance.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picYAUTE1I0"
      Tab(2).Control(1)=   "fraYAUTE1I0"
      Tab(2).Control(2)=   "fgYAUTE1I0"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Stock"
      TabPicture(3)   =   "SAB_Balance.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraYBIASTO0"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraYBIASTO0 
         Height          =   5265
         Left            =   -74880
         TabIndex        =   119
         Top             =   480
         Width           =   8490
         Begin MSFlexGridLib.MSFlexGrid fgYBIASTO0 
            Height          =   4365
            Left            =   0
            TabIndex        =   120
            Top             =   840
            Width           =   8520
            _ExtentX        =   15028
            _ExtentY        =   7699
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
            FormatString    =   $"SAB_Balance.frx":037A
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
            Left            =   2880
            TabIndex        =   125
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label libYBIASTO0_Total 
            Caption         =   "Total"
            Height          =   255
            Left            =   2880
            TabIndex        =   124
            Top             =   240
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
            Left            =   5640
            TabIndex        =   123
            Top             =   360
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
            Left            =   120
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
            Left            =   120
            TabIndex        =   121
            Top             =   120
            Width           =   2655
         End
      End
      Begin VB.PictureBox picYAUTE1I0 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00F0FFFF&
         Height          =   5055
         Left            =   -71400
         ScaleHeight     =   4995
         ScaleWidth      =   4995
         TabIndex        =   25
         Top             =   480
         Width           =   5055
      End
      Begin VB.Frame fraYAUTE1I0 
         Height          =   855
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   8295
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
            Height          =   645
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.Frame fraMvt 
         Height          =   5265
         Left            =   60
         TabIndex        =   8
         Top             =   390
         Width           =   8490
         Begin MSFlexGridLib.MSFlexGrid fgYBIAMVT0 
            Height          =   5085
            Left            =   45
            TabIndex        =   9
            Top             =   135
            Width           =   8400
            _ExtentX        =   14817
            _ExtentY        =   8969
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
            FormatString    =   $"SAB_Balance.frx":041A
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
      End
      Begin VB.Frame fraTab0 
         Height          =   5310
         Left            =   -74865
         TabIndex        =   5
         Top             =   315
         Width           =   8430
         Begin VB.Frame fraSelect 
            Height          =   990
            Left            =   75
            TabIndex        =   11
            Top             =   165
            Width           =   6420
            Begin VB.CommandButton cmdOptions 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Options"
               Height          =   600
               Left            =   3720
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   240
               Width           =   1140
            End
            Begin VB.CommandButton cmdSelect_MesComptes 
               BackColor       =   &H00C0FFFF&
               Caption         =   "MesComptes"
               Height          =   600
               Left            =   2400
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   240
               Width           =   1155
            End
            Begin VB.CommandButton cmdSelect 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Exécuter"
               Height          =   600
               Left            =   5040
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox txtCompte 
               Height          =   315
               Left            =   960
               TabIndex        =   13
               Top             =   600
               Width           =   1305
            End
            Begin VB.TextBox txtIntitulé 
               Height          =   330
               Left            =   960
               TabIndex        =   12
               Top             =   240
               Width           =   1275
            End
            Begin VB.Label lblSelect_Compte 
               Caption         =   "Compte"
               Height          =   195
               Left            =   120
               TabIndex        =   18
               Top             =   600
               Width           =   690
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
            Height          =   1050
            Left            =   6780
            TabIndex        =   10
            Top             =   150
            Width           =   1575
            Begin MSComCtl2.DTPicker txtAmjMin 
               Height          =   300
               Left            =   300
               TabIndex        =   0
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
               Format          =   19595267
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtAmjMax 
               Height          =   300
               Left            =   300
               TabIndex        =   1
               Top             =   630
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
               Format          =   19595267
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   4080
            Left            =   45
            TabIndex        =   7
            Top             =   1185
            Width           =   8340
            _ExtentX        =   14711
            _ExtentY        =   7197
            _Version        =   393216
            Rows            =   1
            Cols            =   12
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
            FormatString    =   $"SAB_Balance.frx":04EB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgYAUTE1I0 
         Height          =   4245
         Left            =   -74880
         TabIndex        =   19
         Top             =   1320
         Width           =   8400
         _ExtentX        =   14817
         _ExtentY        =   7488
         _Version        =   393216
         Rows            =   1
         Cols            =   12
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
         FormatString    =   $"SAB_Balance.frx":058C
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
      Left            =   8190
      Picture         =   "SAB_Balance.frx":069F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   -30
      Width           =   500
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
         Caption         =   "Export => Compte + Service"
      End
      Begin VB.Menu mnuAuto_Balance_Stock 
         Caption         =   "Auto : Balance par service  + Stock opé"
      End
      Begin VB.Menu mnuContextX1 
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
      Begin VB.Menu mnuSelect_Print_Extrait 
         Caption         =   "Imprimer les Extraits "
      End
      Begin VB.Menu mnuSelect_Print_Cumul 
         Caption         =   "Imprimer cumul des mouvements"
      End
      Begin VB.Menu mnuSelect_Print_Balance 
         Caption         =   "Imprimer la  balance"
      End
      Begin VB.Menu mnuSelect_Print_Balance_Stock 
         Caption         =   "Imprimer Balance  + Stock opérations"
      End
      Begin VB.Menu mnuSelect_Print_Client_Stat 
         Caption         =   "Imprimer stat catégorie client"
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
   Begin VB.Menu mnufgSelect 
      Caption         =   "mnufgSelect"
      Visible         =   0   'False
      Begin VB.Menu mnufgSelect_fgYBIAMVT0 
         Caption         =   "afficher Mvts"
      End
      Begin VB.Menu mnufgSelect_fgYBIASTO0 
         Caption         =   "afficher Stock"
      End
      Begin VB.Menu mnufgSelect_fgYAUTE1I0 
         Caption         =   "afficher AUT"
      End
      Begin VB.Menu mnufgSelect_Print_RIB 
         Caption         =   "Imprimer RIB"
      End
      Begin VB.Menu mnufgSelect_Détail 
         Caption         =   "Afficher détail"
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

Dim xMvtP0 As typeMvtP0

Dim Nb As Long

Dim mcboPCEC As String, mcboDevise As String, mcboPLANCOPRO As String

Dim marrYBIACPT0() As typeYBIACPT0, marrYBIACPT0_Nb As Long
Dim xYBIAMVT0 As typeYBIAMVT0

Dim meYbase As typeYBase, xYbase As typeYBase
Dim xAmjMin As String, xAmjMax As String
Dim mAmjMin As String, mAmjmax As String
Dim mAmj_SOLDEDMO As Long, mAmj_COMPTEOUV As Long

Dim blnMesComptes As Boolean
Dim mYAUTE1I0_AUTE1INIV As Long

Dim cnADO As New ADODB.Connection
Dim rsADO As New ADODB.Recordset

Dim blnSelect_Pays As Boolean



Dim fgYBIASTO0_FormatString As String, fgYBIASTO0_K As Integer
Dim fgYBIASTO0_RowDisplay As Integer, fgYBIASTO0_RowClick As Integer, fgYBIASTO0_ColClick As Integer
Dim fgYBIASTO0_ColorClick As Long, fgYBIASTO0_ColorDisplay As Long
Dim fgYBIASTO0_Sort1 As Integer, fgYBIASTO0_Sort2 As Integer
Dim fgYBIASTO0_SortAD As Integer, fgYBIASTO0_Sort1_Old As Integer
Dim fgYBIASTO0_arrIndex As Integer
Dim blnfgYBIASTO0_DisplayLine As Boolean
Dim meYBIASTO0 As typeYBIASTO0, xYBIASTO0 As typeYBIASTO0

Dim marrYCOMREF0() As typeYCOMREF0, marrYCOMREF0_Nb As Long
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
Dim xSql As String
Dim curTotal As Currency, nbTotal As Long, curSolde As Currency
fgYBIASTO0_Reset
fgYBIASTO0.Rows = 1
fgYBIASTO0.FormatString = fgYBIASTO0_FormatString
fgYBIAMVT0.Visible = False

Set rsADO = Nothing
libYBIASTO0_Diff = ""
curTotal = 0: nbTotal = 0
If blnJPL Then Exit Sub

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIASTO0 where " _
     & "YSTOPCI like '" & mId$(xYBIACPT0.COMPTEOBL, 1, 5) & "%'" _
     & "AND YSTODEV = '" & xYBIACPT0.COMPTEDEV & "'" _
     & "AND YSTOCLI = " & Val(xYBIACPT0.CLIENACLI)
Set rsADO = cnADO.Execute(xSql)

Do While Not rsADO.EOF
    V = srvYBIASTO0_GetBuffer_ODBC(rsADO, xYBIASTO0)
    If Not IsNull(V) Then
        MsgBox V, vbCritical, "frmSAB_Balance.SQL_ODBC"
        Exit Sub
    Else
        fgYBIASTO0_DisplayLine
        curTotal = curTotal + xYBIASTO0.YSTOMON
        nbTotal = nbTotal + 1
    End If
    rsADO.MoveNext
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


Public Sub fgYAUTE1I0_Display(lCLIENACLI As String)

Dim intReturn As Integer
Dim mK1 As String


recYAUTE1I0_Init xYAUTE1I0

lstYAUTE1I0.Clear
lstYAUTE1I0.AddItem Space$(7) & lCLIENACLI


xYbase.ID = constYAUTE1I0_GRP
xYbase.K1 = lCLIENACLI
xYbase.Method = "Seek>"

Do
    intReturn = tableYBase_Read(xYbase)
    If Trim(xYbase.ID) <> constYAUTE1I0_GRP Then intReturn = -1
    If intReturn = 0 Then
        If mId$(xYbase.K1, 1, 7) = lCLIENACLI Then
            lstYAUTE1I0.AddItem mId$(xYbase.K1, 8, 7) & "0"
             lstYAUTE1I0.AddItem mId$(xYbase.K1, 8, 7) & "9999999"
       Else
            intReturn = -1
        End If
        
    End If
        
Loop Until intReturn <> 0


lstYAUTE1I0.ListIndex = 0

End Sub


Public Sub fgYAUTE1I0_Display_GRP(lK1 As String)

Dim intReturn As Integer
Dim lenK1 As Integer
Dim blnDisplay As Boolean

''SSTab1.Tab = 2

fgYAUTE1I0_Reset
picYAUTE1I0.Visible = False

fgYAUTE1I0.Rows = 1
fgYAUTE1I0.FormatString = fgYAUTE1I0_FormatString
fgYAUTE1I0.Visible = False

If optYAUTE1i0_NIV_1 Then mYAUTE1I0_AUTE1INIV = 1
If optYAUTE1i0_NIV_2 Then mYAUTE1I0_AUTE1INIV = 2
If optYAUTE1i0_NIV_X Then mYAUTE1I0_AUTE1INIV = 999
meYbase.ID = constYAUTE1I0
meYbase.K1 = lK1
lenK1 = Len(lK1)
meYbase.Method = "Seek>"

Do
    intReturn = tableYBase_Read(meYbase)
    If Trim(meYbase.ID) <> constYAUTE1I0 Then intReturn = -1
    If intReturn = 0 Then
           '  Debug.Print meYbase.Text
       If mId$(meYbase.K1, 1, lenK1) = lK1 Then
            MsgTxt = Space$(34) & meYbase.Text
            MsgTxtIndex = 0
            srvYAUTE1I0_GetBuffer xYAUTE1I0
            blnDisplay = False
            If xYAUTE1I0.AUTE1IMDB = 0 And xYAUTE1I0.AUTE1IMCR = 0 And xYAUTE1I0.AUTE1IMAU = 0 Then
            Else
                If xYAUTE1I0.AUTE1ITYP = 1 And xYAUTE1I0.AUTE1IDEV = xYAUTE1I0.AUTE1IDBA Then
                Else
                   If xYAUTE1I0.AUTE1INIV <= mYAUTE1I0_AUTE1INIV Then blnDisplay = True
                End If
            End If
            If xYAUTE1I0.AUTE1IDEP <> " " Then blnDisplay = True
            If blnDisplay Then fgYAUTE1I0_DisplayLine
        Else
            intReturn = -1
        End If
        
    End If
        
Loop Until intReturn <> 0
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "affichage mvt : " & fgYAUTE1I0.Rows - 1)

fgYAUTE1I0.Visible = True
fgYAUTE1I0_Sort1 = -1
blnfgYAUTE1I0_Display = True

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




Private Sub fgSelect_Display()
Dim I As Long, blnOk As Boolean
Dim Nb As Long
Dim xPCEC As String, lenPCEC As Integer
Dim xPRO As String, lenPRO As Integer
Dim xSelect As String, lenSelect As Integer
Dim xIntitulé As String, lenIntitulé As Integer
Dim K As Integer
Dim xCLIENARES As String, lenCLIENARES As Integer
Dim xCLIENARSD As String, lenCLIENARSD  As Integer
Dim xCOMPTECLA As Long, lenCOMPTECLA  As Integer
Dim wDate As Long

fraSelect_Options.Visible = False
SSTab1.Tab = 0
mcboDevise = Trim(cboDevise.Text)
mcboPLANCOPRO = Trim(cboPLANCOPRO.Text)
xPCEC = Trim(cboPCEC): lenPCEC = Len(xPCEC)
xPRO = Trim(cboPLANCOPRO): lenPRO = Len(xPRO)
xSelect = Trim(txtCompte): lenSelect = Len(xSelect)
xIntitulé = Trim(txtIntitulé): lenIntitulé = Len(xIntitulé)
Call DTPicker_Amj7(txtSelect_SOLDEDMO, mAmj_SOLDEDMO)
Call DTPicker_Amj7(txtSelect_COMPTEOUV, mAmj_COMPTEOUV)
xCLIENARES = Trim(txtSelect_CLIENARES): lenCLIENARES = Len(xCLIENARES)
xCLIENARSD = Trim(txtSelect_CLIENARSD): lenCLIENARSD = Len(xCLIENARSD)
xCOMPTECLA = Val(Trim(txtSelect_COMPTECLA)): lenCOMPTECLA = Len(Trim(txtSelect_COMPTECLA))
If chkSelect = "1" Then
    cboDevise.ListIndex = 0
    cboPCEC.ListIndex = 0
    cboPLANCOPRO.ListIndex = 0
    xPCEC = "": lenPCEC = 0
    xPRO = "": lenPRO = 0
    xCLIENARES = "": lenCLIENARES = 0
    xCLIENARSD = "": lenCLIENARSD = 0
    xCOMPTECLA = 0: lenCOMPTECLA = 0
    mcboDevise = ""
    txtCompte = "": lenSelect = 0
    txtIntitulé = "": lenIntitulé = 0
Else

    If xPCEC = "" _
    And mcboPLANCOPRO = "" _
    And mcboDevise = "" _
    And lenSelect = 0 _
    And lenCLIENARES = 0 _
    And lenCLIENARSD = 0 _
    And lenCOMPTECLA = 0 _
    And lenIntitulé = 0 Then
        Exit Sub
    End If
End If

If chkSelect_SOLDEDMO = "1" And chkSelect_DORCPTDMV = "1" Then
    Call ZDORCPT_SQL_ODBC
End If

fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Visible = False


For I = 1 To marrYBIACPT0_Nb

    blnOk = True
    xYBIACPT0 = marrYBIACPT0(I)
    Select Case xYBIACPT0.SOLDECEN
        Case Is = 0: If chkSelect_SoldeZ = "1" Then blnOk = False
        Case Is > 0: If chkSelect_SoldeDb = "1" Then blnOk = False
        Case Is < 0: If chkSelect_SoldeCr = "1" Then blnOk = False
    End Select
    
    If blnOk And chkSelect_Annulé = "1" Then
            If xYBIACPT0.COMPTEFON = "4" Then blnOk = False
    End If
    
    If blnOk And lenPCEC > 0 Then
            If mId$(xYBIACPT0.COMPTEOBL, 1, lenPCEC) <> xPCEC Then blnOk = False
    End If
    
    If blnOk And chkSelect_HB = "1" Then
            If mId$(xYBIACPT0.COMPTEOBL, 1, 1) = "9" Then blnOk = False
    End If
    
    If blnOk And lenPRO > 0 Then
            If mId$(xYBIACPT0.PLANCOPRO, 1, lenPRO) <> xPRO Then blnOk = False
    End If
    If blnOk And mcboDevise <> "" Then
        If mcboDevise <> Trim(xYBIACPT0.COMPTEDEV) Then blnOk = False
    End If


    If blnOk And lenSelect > 0 Then
            K = InStr(1, xYBIACPT0.COMPTECOM, xSelect)
            If K = 0 Then blnOk = False
    
    End If
    If blnOk And lenIntitulé > 0 Then
            K = InStr(1, xYBIACPT0.CLIENASIG, xIntitulé)    ' COMPTEINT
            If K = 0 Then blnOk = False
    End If
    
    If blnOk And chkSelect_SOLDEDMO = "1" Then
        If chkSelect_DORCPTDMV = "1" Then
            wDate = wDORCPTDMV(I)
            If wDate = 0 Then wDate = xYBIACPT0.SOLDEDMO   ' !!! tous les comptes ne sont pas gérés par l'application COMPTES DORMANTS
        Else
            wDate = xYBIACPT0.SOLDEDMO
        End If
        
        If wDate > mAmj_SOLDEDMO Then
            If optSelect_SOLDEDMO_Inf Then blnOk = False
        Else
            If optSelect_SOLDEDMO_Sup Then blnOk = False
        End If
        
    End If
    
    If blnOk And chkSelect_COMPTEOUV = "1" Then
        If xYBIACPT0.COMPTEOUV < mAmj_COMPTEOUV Then blnOk = False
    End If
    
    If blnOk And lenCLIENARES > 0 Then
            If mId$(xYBIACPT0.CLIENARES, 1, lenCLIENARES) <> xCLIENARES Then blnOk = False
    End If
    If blnOk And blnSelect_Pays Then
            If xYBIACPT0.CLIENARSD = "   " Then blnOk = False
    End If
    
    If blnOk And lenCLIENARSD > 0 Then
            If mId$(xYBIACPT0.CLIENARSD, 1, lenCLIENARSD) <> xCLIENARSD Then blnOk = False
    End If
    
    If blnOk And lenCOMPTECLA > 0 Then
            If xYBIACPT0.COMPTECLA <> xCOMPTECLA Then blnOk = False
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
Dim intReturn As Integer
Dim wIBMAMJMax As String
'On Error GoTo Error_Handle


txtAMJ_Control

fgYBIAMVT0_Reset

fgYBIAMVT0.Rows = 1
fgYBIAMVT0.FormatString = fgYBIAMVT0_FormatString
fgYBIAMVT0.Visible = False

recYBIAMVT0_Init xYBIAMVT0

meYbase.ID = constYBIAMVT0
meYbase.K1 = lCOMPTECOM & dateIBM(xAmjMin)
wIBMAMJMax = dateIBM(xAmjMax)
meYbase.Method = "Seek>"
'intReturn = tableYBase_Read(meYBase)

Do
    intReturn = tableYBase_Read(meYbase)
    If Trim(meYbase.ID) <> constYBIAMVT0 Then intReturn = -1
    If intReturn = 0 Then
        MsgTxt = Space$(34) & meYbase.Text
        MsgTxtIndex = 0
        srvYBIAMVT0_GetBuffer xYBIAMVT0
        If xYBIAMVT0.MOUVEMCOM = lCOMPTECOM And xYBIAMVT0.MOUVEMDTR <= wIBMAMJMax Then
            fgYBIAMVT0_DisplayLine
        Else
            intReturn = -1
        End If
        
    End If
        
Loop Until intReturn <> 0



Call lstErr_Clear(Me.lstErr, Me.cmdContext, "affichage mvt : " & fgYBIAMVT0.Rows - 1)

fgYBIAMVT0.Visible = True
fgYBIAMVT0_Sort1 = -1
End Sub


Public Sub fgSelect_DisplayLine(lIndex As Long)
Dim X As String
On Error Resume Next
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect.Col = 0: fgSelect.Text = xYBIACPT0.CLIENACLI & "_" & xYBIACPT0.COMPTEINT
            fgSelect.Col = 2: fgSelect.Text = xYBIACPT0.COMPTEDEV
            fgSelect.Col = 3: fgSelect.Text = xYBIACPT0.CLIENARES
            fgSelect.Col = 4: fgSelect.Text = Trim(xYBIACPT0.COMPTEOBL) & " " & xYBIACPT0.PLANCOPRO
            
            fgSelect.Col = 5: fgSelect.Text = xYBIACPT0.COMPTECOM
            
            fgSelect.Col = 6: fgSelect.Text = dateIBM10(xYBIACPT0.SOLDEDMO, True)
            If xYBIACPT0.COMPTEFON <> "0" Then
                If xYBIACPT0.COMPTEBLO = 0 Then
                    X = ""
                Else
                    X = " " & dateIBM10(xYBIACPT0.COMPTEBLO, True)
                End If
                fgSelect.Col = 7: fgSelect.Text = xYBIACPT0.COMPTEFON & X
            End If
             fgSelect.Col = 8: fgSelect.Text = dateIBM10(xYBIACPT0.COMPTEOUV, True)
            fgSelect.Col = 9: fgSelect.Text = xYBIACPT0.CLIENARSD
           fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
            
             fgSelect.Col = 1: fgSelect.Text = Format$(Abs(xYBIACPT0.SOLDECEN), "### ### ### ###.00 ")
           If xYBIACPT0.SOLDECEN > 0 Then
                fgSelect.CellForeColor = vbRed
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
fgYBIAMVT0.Col = fgYBIAMVT0_arrIndex: fgYBIAMVT0.Text = meYbase.K1
 
fgYBIAMVT0.Col = 3: fgYBIAMVT0.Text = Format$(Abs(xYBIAMVT0.MOUVEMMON), "### ### ### ###.00 ")
If xYBIAMVT0.MOUVEMMON > 0 Then
     fgYBIAMVT0.CellForeColor = vbRed
 Else
     fgYBIAMVT0.CellForeColor = vbBlue
 End If

End Sub
Public Sub fgYAUTE1I0_DisplayLine()
Dim X As String, I As Integer
Dim wCellBackColor As Long, wCellForeColor As Long

On Error Resume Next
fgYAUTE1I0.Rows = fgYAUTE1I0.Rows + 1
fgYAUTE1I0.Row = fgYAUTE1I0.Rows - 1
Select Case xYAUTE1I0.AUTE1INIV
    Case 0: wCellBackColor = &HC0FFC0    ' green
    Case 1: wCellBackColor = &HFFC0C0    ' blue
    Case 2: wCellBackColor = &HFFFFC0   ' cyan
    Case Else: wCellBackColor = &HFFC0FF    'magenta
End Select
If xYAUTE1I0.AUTE1IDEP <> " " Then
     wCellForeColor = vbRed
 Else
     fgYAUTE1I0.CellForeColor = vbBlue
 End If

fgYAUTE1I0.Col = 0: fgYAUTE1I0.Text = xYAUTE1I0.AUTE1IGRP
fgYAUTE1I0.Col = 1: fgYAUTE1I0.Text = xYAUTE1I0.AUTE1ICLI
fgYAUTE1I0.Col = 2: fgYAUTE1I0.Text = xYAUTE1I0.AUTE1IAUT
fgYAUTE1I0.Col = 3: fgYAUTE1I0.Text = xYAUTE1I0.AUTE1IDEV
If xYAUTE1I0.AUTE1IMDB <> 0 Then fgYAUTE1I0.Col = 4: fgYAUTE1I0.Text = Format$(Abs(xYAUTE1I0.AUTE1IMDB), "### ### ### ###.00 ")
If xYAUTE1I0.AUTE1IMCR <> 0 Then fgYAUTE1I0.Col = 5: fgYAUTE1I0.Text = Format$(Abs(xYAUTE1I0.AUTE1IMCR), "### ### ### ###.00 ")
If xYAUTE1I0.AUTE1IMTD <> 0 Then fgYAUTE1I0.Col = 6: fgYAUTE1I0.Text = Format$(Abs(xYAUTE1I0.AUTE1IMTD), "### ### ### ###.00 ")
If xYAUTE1I0.AUTE1IMAU <> 0 Then
    fgYAUTE1I0.Col = 7:    fgYAUTE1I0.Text = Format$(Abs(xYAUTE1I0.AUTE1IMAU), "### ### ### ###.00 ")
    fgYAUTE1I0.Col = 8: fgYAUTE1I0.Text = dateIBM10(xYAUTE1I0.AUTE1IDAF, True)
End If

fgYAUTE1I0.Col = fgYAUTE1I0_arrIndex - 2: fgYAUTE1I0.Text = meYbase.Text
fgYAUTE1I0.Col = fgYAUTE1I0_arrIndex: fgYAUTE1I0.Text = meYbase.K1
 

For I = 0 To fgYAUTE1I0_arrIndex
    fgYAUTE1I0.Col = I
    fgYAUTE1I0.CellForeColor = wCellForeColor
    fgYAUTE1I0.CellBackColor = wCellBackColor
    
Next I
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
        Case 1: fgSelect.Col = 1: X = fgSelect.Text
                fgSelect.Col = fgSelect_arrIndex - 1
                fgSelect.Text = Format$(Val(X), "000000000000.00")
        Case 2: '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ TRI POUR IMPRESSION BALANCE DEV / PCEC / COMPTE
                fgSelect.Col = fgSelect_arrIndex
                wIndex = Val(fgSelect.Text)
                X = marrYBIACPT0(wIndex).COMPTEDEV & marrYBIACPT0(wIndex).COMPTEOBL & marrYBIACPT0(wIndex).COMPTECOM
                fgSelect.Col = fgSelect_arrIndex - 1
                fgSelect.Text = X
        Case 5: fgSelect.Col = 5: X = fgSelect.Text
                fgSelect.Col = fgSelect_arrIndex - 1
                fgSelect.Text = mId$(X, 10, 1) & X          ' code résidence pos=10
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


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init("SAB_Balance", SAB_Balance_Aut)

'blnSetfocus = True
Form_Init

Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case "@AUTO_COMPTA"
                'If paramEnvironement = constProduction Then
                    meUnit.ID = "CPT"
                    Table_Unit meUnit
                    Printer_Set meUnit.Printer
                'End If
                blnAuto = True
                SSTab1.Tab = 0
                lstService.ListIndex = 0
                chkSelect = 1
                mnuSelect_Print_Balance_Click
                chkBalance_Détail = "0"
                chkBalance_Récap = "0"
                chkBalance_Récap_Bilan = "1"
                cmdBalance_Ok_Click
                Unload Me

    Case Else: blnAuto = False
End Select

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
    MsgBox "paramétrage inconsistent", vbCritical, "frmSAB_YBIACPT0.param_init"
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

lstService.Top = 0
lstService.Height = 6100
lstService.Width = 3900
lstService.Visible = True
lstService.AddItem "   Tous les comptes"
xElpTable.ID = "SAb_Param"
xElpTable.K1 = "Compte_Unit"
xElpTable.K2 = ""

Call lst_LoadK2(xElpTable, lstService)

lstErr_Clear lstErr, cmdContext, "<== SELECTIONNER un SERVICE"
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

fraSelect_Clear
mcboPCEC = ""
mcboDevise = ""

fgSelect.Font = prtFontName_CourierNew
fgYBIAMVT0.Font = prtFontName_CourierNew
fgYBIASTO0.Font = prtFontName_CourierNew
picYAUTE1I0.Enabled = False

mAmjmax = YBIATAB0_DATE_CPT_J: Call DTPicker_Set(txtAmjMax, mAmjmax)
mAmjMin = mId$(YBIATAB0_DATE_CPT_J, 1, 6) & "01": Call DTPicker_Set(txtAmjMin, mAmjMin)
mAmjMin = mId$(YBIATAB0_DATE_CPT_MP1, 1, 6) & "01"
chkSelect_SOLDEDMO = "0"
fraSelect_SOLDEDMO.Visible = False
optSelect_SOLDEDMO_Inf = True
Call DTPicker_Set(txtSelect_SOLDEDMO, mAmjmax)

chkSelect_COMPTEOUV = "0"
txtSelect_COMPTEOUV.Visible = False
Call DTPicker_Set(txtSelect_COMPTEOUV, mAmjmax)

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

blnControl = True

End Sub



Public Function param_Init()
Dim I As Long


ReDim marrYBIACPT0(paramYBIACPT0_Nb + 10)
param_Init = Null



Call lstErr_Clear(lstErr, cmdContext, "> SAb_Balance_Import.....attendre 2 minutes !"): DoEvents

'chkSelect = "1"

'''srvYBIACPT0_Import_Array marrYBIACPT0_Nb, marrYBIACPT0()
''Call lstErr_AddItem(lstErr, cmdContext, ". SAb_Balance_Import_YBIACPT0_" & marrYBIACPT0_Nb): DoEvents

fgSelect_Display
Call lstErr_AddItem(lstErr, cmdContext, ". SAb_Balance_Import cbo"): DoEvents

fgSelect.Visible = False


srvYBIATAB0_Import_cboDevise cboDevise
arrDevise_Nb = cboDevise.ListCount - 1
ReDim arrDevise(cboDevise.ListCount)
For I = 0 To arrDevise_Nb
    cboDevise.ListIndex = I
    arrDevise(I) = cboDevise.Text
Next I
arrService_Nb = 1
ReDim arrService(1): arrService(0) = "***": arrService(1) = ""
arrService_Balance_Cumul_Z

srvYPLAN0_Import_cboPCEC cboPCEC
srvYPLAN0_Import_cboPLANCOPRO cboPLANCOPRO

'chkSelect = "0"

fgSelect.Visible = True

Call lstErr_ChangeLastItem(lstErr, cmdContext, "= SAb_Balance_Import"): DoEvents

Me.Enabled = True: Me.MousePointer = 0



End Function
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
        chkSelect.ForeColor = warnUsrColor
        Me.Enabled = False: Me.MousePointer = vbHourglass
        fgSelect_Display
        Me.Enabled = True: Me.MousePointer = 0
    Else
        chkSelect.ForeColor = lblUsr.ForeColor
    End If
    
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

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdOptions_Click()
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
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdSelect_MesComptes_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Clear
chkSelect = "0"
cboPCEC = ""
cboPLANCOPRO = ""
mcboDevise = ""
txtCompte = ""
txtIntitulé = currentYBIAUSR0.CLIENASIG
blnMesComptes = True
fgSelect_Display
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdTest_Click()
End Sub

Private Sub fgYAUTE1I0_Click()
fgYAUTE1I0.LeftCol = 0

End Sub

Private Sub fgYAUTE1I0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wId As String
On Error Resume Next
If Y <= fgYAUTE1I0.RowHeightMin Then
    Select Case fgYAUTE1I0.Col
        Case 0: fgYAUTE1I0_Sort1 = 0: fgYAUTE1I0_Sort2 = 2: fgYAUTE1I0_Sort
        Case 1: fgYAUTE1I0_Sort1 = 1: fgYAUTE1I0_Sort2 = 2: fgYAUTE1I0_Sort
        Case 2:  fgYAUTE1I0_Sort1 = 2: fgYAUTE1I0_Sort2 = 2: fgYAUTE1I0_Sort
    End Select
Else
    fgYAUTE1i0_Select
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

Private Sub fgYBIAMVT0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wId As String
On Error Resume Next
If Y <= fgYBIAMVT0.RowHeightMin Then
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
        If Not IsNull(srvYBIAMVT0_Import_Read(wId, xYBIAMVT0)) = 0 Then
            srvYBIAMVT0_ElpDisplay xYBIAMVT0
            
         Else
            Shell_MsgBox "fgYBIAMVT0_MouseDown# " & xMvtP0.ID & " : " & xMvtP0.Err, vbCritical, Me.Caption, False

        End If
    End If
   End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
If Not blnJPL Then
    If blnControl Then
        cnADO.Close
        Set cnADO = Nothing
    End If
End If
End Sub

Private Sub lstService_Click()
Dim xSql As String
Dim xService As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, ". SAb_Balance_Import_YBIACPT0 : en cours ...."): DoEvents

xService = Trim(mId$(lstService, 1, 2))
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0" ''' where PLANCOPRO = 'CAT'"
If xService = "" Then
   YBIACPT0_SQL_ODBC xSql & " ORDER by COMPTECOM"
   fraSelect_Options.Visible = True
Else

    xSql = xSql & " C , ZCOMREF0 R where C.COMPTECOM = R.COMREFCOM AND COMREFCOR = '" & xService & "'" & " ORDER by COMPTECOM"
   YBIACPT0_SQL_ODBC xSql
    fraSelect_Clear
    blnControl = False:    chkSelect = "1": blnControl = True
    fgSelect_Display
End If
Call lstErr_AddItem(lstErr, cmdContext, ". SAb_Balance_Import_YBIACPT0 :" & marrYBIACPT0_Nb & " comptes"): DoEvents
SSTab1.Caption = Trim(lstService)
lstService.Visible = False
SSTab1.Visible = True
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuAuto_Balance_Stock_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
SSTab1.Tab = 0
X = MsgBox("Voulez-vous imprimer le détail par services ?", vbYesNo + vbQuestion + vbDefaultButton1, "Sab_Balance : Impression Balance / Stock")
If X = vbYes Then
    blnBalance_Stock_détail = True
    
    X = MsgBox("Voulez-vous routerles balances vers les imprimantes des services ?", vbYesNo + vbQuestion + vbDefaultButton1, "Sab_Balance : Impression Balance / Stock")
    If X = vbYes Then
        blnService_Printer = True
    Else
        blnService_Printer = False
    End If

Else
    blnBalance_Stock_détail = False
End If

cmdBalance_Ok_Stock
Call prtSAB_Balance_Cumul_Monitor(arrService_Balance_Cumul(), arrService_Nb, arrDevise_Nb)
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub mnuAuto_Client_Stat_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

fraSelect_Clear

cboPLANCOPRO = "CAV"
fgSelect_Display

mnuSelect_Print_Client_Stat_Click
 
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub XXXAuto_Compta_Balance_Pays_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
blnControl = False
fraSelect_Clear
chkSelect_SoldeZ = "1"
chkSelect = "1"
blnSelect_Pays = True

fgSelect_Display
fgSelect_SortX 9
X = cmdBalance_Ok_Param
Mid$(X, 11, 1) = "1"     ' balance par pays
prtSAB_Balance_Monitor X, xAmjMin, xAmjMax, fgSelect, marrYBIACPT0(), marrYBIACPT0_Nb, "", wYSTOMON(), wDORCPTDMV()
Me.Show
Me.Enabled = True: Me.MousePointer = 0
 

End Sub

Private Sub mnuAuto_Compta_TVA_Click()


Me.Enabled = False: Me.MousePointer = vbHourglass

fraSelect_Clear
chkSelect_SoldeZ = "1"

cboPLANCOPRO = "PRO"
chkSelect_Résidence = "1"
chkSelect_MOUVEMDCO = "1"
txtCompte = "70"
Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_MP1)
    Call DTPicker_Set(txtAmjMin, mId$(YBIATAB0_DATE_CPT_MP1, 1, 6) & "01")
'cas particulier SAB JANVIER : ignorer les éctitures de régularisation du 01-01
'---------------------------------------------------------------------------------
If mId$(YBIATAB0_DATE_CPT_MP1, 5, 2) = "01" Then
    Call DTPicker_Set(txtAmjMin, mId$(YBIATAB0_DATE_CPT_MP1, 1, 6) & "02")
Else
    Call DTPicker_Set(txtAmjMin, mId$(YBIATAB0_DATE_CPT_MP1, 1, 6) & "01")
End If

txtAMJ_Control

fgSelect_Display_SoldeZ

mnuSelect_Print_Cumul_Click
 
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnufgSelect_Détail_Click()
srvYBIACPT0_ElpDisplay xYBIACPT0
End Sub

Private Sub mnufgSelect_fgYAUTE1I0_Click()
SSTab1.Tab = 2
SSTab1_GotFocus
End Sub

Private Sub mnufgSelect_fgYBIAMVT0_Click()
SSTab1.Tab = 1
SSTab1_GotFocus
End Sub

Private Sub mnufgSelect_fgYBIASTO0_Click()
SSTab1.Tab = 3
SSTab1_GotFocus
End Sub


Private Sub mnufgSelect_Print_RIB_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

prtRIB_Monitor xYBIACPT0.COMPTECOM
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

Private Sub picYAUTE1I0_Click()
srvYAUTE1I0_ElpDisplay xYAUTE1I0
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
Call DTPicker_Set(txtAmjMin, YBIATAB0_DATE_CPT_J)
txtAMJ_Control

mnuSelect_Print_Relevé_Click

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuAuto_SOBI_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

fraSelect_Clear
chkSelect_SoldeZ = "1"
Call DTPicker_Set(txtAmjMax, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtAmjMin, YBIATAB0_DATE_CPT_J)
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

Private Sub mnuSelect_Print_Click()
'blnTotal = True
'cmdPrint_Journal

End Sub

Private Sub cmdSelect_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
blnMesComptes = False
fgSelect_Display
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub fgSelect_Click()
fgSelect.LeftCol = 0

End Sub

Private Sub fgSelect_LeaveCell()
On Error Resume Next
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

If Not blnJPL Then cnADO.Open paramODBC_DSN_SAB

End Sub





Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xId As String
Dim V

On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_SortX fgSelect_Sort1
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_SortX fgSelect_Sort1
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 4:  fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_SortX fgSelect_Sort1
        Case 5:
            If chkSelect_Résidence = "1" Then
                fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_SortX fgSelect_Sort1
            Else
                fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_Sort
            End If
        Case 4:  fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_SortX fgSelect_Sort1
        Case 6:  fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_SortX fgSelect_Sort1
        Case 7:  fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_SortX fgSelect_Sort1
        Case 9:  fgSelect_Sort1 = 9: fgSelect_Sort2 = 9: fgSelect_SortX fgSelect_Sort1
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex
             xYBIACPT0 = marrYBIACPT0(Val(fgSelect.Text))
            fgSelect.LeftCol = 0
            
            'srvYBIACPT0_ElpDisplay xYBIACPT0
            If xYBIACPT0.PLANCOPRO = "CAV" _
            Or xYBIACPT0.PLANCOPRO = "LOR" _
            Or xYBIACPT0.PLANCOPRO = "LOB" Then
                mnufgSelect_Print_RIB.Enabled = True
            Else
                mnufgSelect_Print_RIB.Enabled = False
            End If
            
             fraMvt.Visible = False
            fraYAUTE1I0.Visible = False
            fraYBIASTO0.Visible = False
            'If Button = vbRightButton Then
                Me.PopupMenu mnufgSelect, vbPopupMenuLeftButton
            'End If
       Else
            Shell_MsgBox "fgSelect_MouseDown# " & xMvtP0.ID & " : " & xMvtP0.Err, vbCritical, Me.Caption, False

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


Public Sub txt_X()
'Call txt_GotFocus(txt)
'KeyAscii = convUCase(KeyAscii)
'Call txt_LostFocus(txt)

'Call txt_GotFocus(txt)
'If XopDevise(2).maxD = 0 Then
'    Call num_KeyAscii(KeyAscii)
'Else
'    Call num_KeyAsciiD(KeyAscii, txt)
'End If
'Call txt_LostFocus(txt)

End Sub




Private Sub mnuSelect_Print_Recap_Click()
'blnTotal = False
'cmdPrint_Journal

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
Me.Enabled = False: Me.MousePointer = vbHourglass
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Relevés : " & fgSelect.Rows - 1)
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
prtSAB_Balance_Monitor xFct, xAmjMin, xAmjMax, fgSelect, marrYBIACPT0(), marrYBIACPT0_Nb, "", wYSTOMON(), wDORCPTDMV()
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
            If Not fraYAUTE1I0.Visible Then
                fgYAUTE1I0_Display xYBIACPT0.CLIENACLI
                fraYAUTE1I0.Visible = True
            End If
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
cboPLANCOPRO.ListIndex = 0
cboPCEC = ""
cboPLANCOPRO = ""
mcboDevise = ""
txtSelect_CLIENARES = ""
txtSelect_CLIENARSD = ""
txtSelect_COMPTECLA = ""

chkSelect_SoldeZ = "0"
chkSelect_SoldeCr = "0"
chkSelect_SoldeDb = "0"
chkSelect_COMPTEOUV = "0"
chkSelect_SOLDEDMO = "0"
chkSelect_MOUVEMDCO = "0"
chkSelect_Résidence = "0"
chkSelect_HB = "0"
chkSelect_Annulé = "0"
blnSelect_Pays = False

fraMvt.Visible = False
fraYAUTE1I0.Visible = False
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
    MsgTxt = Space$(34) & fgYAUTE1I0.Text
    MsgTxtIndex = 0
    srvYAUTE1I0_GetBuffer xYAUTE1I0
    picYAUTE1I0_Display
    fgSelect.LeftCol = 0
    picYAUTE1I0.Enabled = True
Else
        Shell_MsgBox "fgYAUTE1I0_MouseDown# " & xMvtP0.ID & " : " & xMvtP0.Err, vbCritical, Me.Caption, False

    End If

End Sub

Public Sub fgSelect_Display_SoldeZ()
' !!! chkSelect_SoldeZ = "0" pour sélectionner tous les comptes
'                      = "1" pour ne pas imprimer les comptes soldés sans mvt
chkSelect_SoldeZ = "0"
fgSelect_Display
''chkSelect_SoldeZ = "1"

End Sub

Public Sub YBIACPT0_SQL_ODBC(xSql As String)
Set rsADO = Nothing
marrYBIACPT0_Nb = 0
 
If blnJPL Then Exit Sub

Set rsADO = cnADO.Execute(xSql)

Do While Not rsADO.EOF
    marrYBIACPT0_Nb = marrYBIACPT0_Nb + 1
    If marrYBIACPT0_Nb >= UBound(marrYBIACPT0) Then ReDim Preserve marrYBIACPT0(marrYBIACPT0_Nb + 50)
    V = srvYBIACPT0_GetBuffer_ODBC(rsADO, marrYBIACPT0(marrYBIACPT0_Nb))
    If Not IsNull(V) Then
        MsgBox V, vbCritical, "frmSAB_Balance.SQL_ODBC"
        Exit Sub
    End If
    rsADO.MoveNext
Loop

blnZDORCPT_SQL_ODBC = False
ReDim wDORCPTDMV(marrYBIACPT0_Nb)
For I = 1 To marrYBIACPT0_Nb
    wDORCPTDMV(I) = 0
Next I

End Sub

Public Sub ZCOMREF0_SQL_ODBC()
Dim xSql As String, mService As String
Dim K As Long
Dim X20 As String * 20

Set rsADO = Nothing
ReDim marrYCOMREF0(10000): marrYCOMREF0_Nb = 0
ReDim arrService(20): arrService_Nb = 1: arrService(1) = "": mService = ""
ReDim wCOMREFCOR(marrYBIACPT0_Nb + 1)


If blnJPL Then Exit Sub
xSql = "select * from " & paramIBM_Library_SAB & ".ZCOMREF0 where COMREFCOR like 'G%' order by COMREFCOR, COMREFCOM"
Set rsADO = cnADO.Execute(xSql)

Do While Not rsADO.EOF
    marrYCOMREF0_Nb = marrYCOMREF0_Nb + 1
    If marrYCOMREF0_Nb >= UBound(marrYCOMREF0) Then ReDim Preserve marrYCOMREF0(marrYCOMREF0_Nb + 50)
    V = srvYCOMREF0_GetBuffer_ODBC(rsADO, marrYCOMREF0(marrYCOMREF0_Nb))
    If Not IsNull(V) Then
        MsgBox V, vbCritical, "frmSAB_Balance.SQL_ODBC"
        Exit Sub
    Else
        If mService <> marrYCOMREF0(marrYCOMREF0_Nb).COMREFCOR Then
            arrService_Nb = arrService_Nb + 1
            If arrService_Nb >= UBound(arrService) Then ReDim Preserve arrService(arrService_Nb + 10)
            arrService(arrService_Nb) = marrYCOMREF0(marrYCOMREF0_Nb).COMREFCOR
            mService = arrService(arrService_Nb)
        End If
    End If
    rsADO.MoveNext
Loop

For I = 1 To marrYBIACPT0_Nb
    wCOMREFCOR(I) = ""
Next I


For K = 1 To marrYCOMREF0_Nb
    X20 = marrYCOMREF0(K).COMREFCOM
    For I = 1 To marrYBIACPT0_Nb
        If X20 = marrYBIACPT0(I).COMPTECOM Then
            If wCOMREFCOR(I) <> "" Then MsgBox marrYBIACPT0(I).COMPTECOM, vbExclamation & " déjà affecté a " & wCOMREFCOR(I), "Affectation Comppte / service"
            wCOMREFCOR(I) = marrYCOMREF0(K).COMREFCOR
            Exit For
        End If
    Next I
Next K

End Sub

Public Sub mnuZCOMREF0_Service_Export_click()
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
    If mId$(txtBalance_CSV_Folder, Len(txtBalance_CSV_Folder), 1) <> "\" Then txtBalance_CSV_Folder = txtBalance_CSV_Folder & "\"
    wFileName = txtBalance_CSV_Folder & txtBalance_CSV_FileName
    wIdFile = 0
    Mid$(lBalance_Ok_Param, 6, 1) = chkBalance_CSV
    V = File_Export_Monitor("Output", wIdFile, wFileName)
    Mid$(lBalance_Ok_Param, 7, 3) = Format$(wIdFile, "000")
    Call File_Export_Monitor("Print", wIdFile, "COMPTEDEV;COMPTEOBL;COMPTECOM;COMPTEINT;DB_dev;CR_dev;DB_eur;CR_eur")

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
Dim wService_Name As String, wService_Printer As String
Dim iDevise As Integer
Dim xSql As String
Dim wBalance_Ok_Param As String
'===========================================================================================
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : chargement des services/comptes")

ZCOMREF0_SQL_ODBC
'======================================== Compte sans service => Responsable client ===================
mService = ""
For I = 1 To marrYBIACPT0_Nb
    If wCOMREFCOR(I) = "" Then
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
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : chargement OPENAT_PCI")
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID  = 'OPENAT_PCI'"
Set rsADO = cnADO.Execute(xSql)

Do While Not rsADO.EOF
    X = mId$(rsADO("BIATABTXT"), 4, 5)
    For I = 1 To marrYBIACPT0_Nb
        If X = mId$(marrYBIACPT0(I).COMPTEOBL, 1, 5) Then
             wYSTOMON(I) = -1
        End If
    Next I
    rsADO.MoveNext
Loop
'===========================================================================================
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : chargement du stock")
xWhere = ""
Call YBIASTO0_Sql(xWhere, nbDossier, arrYBIASTO0(), stockYBIACPT0(), stockCompte_Nb, cnADO, rsADO)

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
    If mId$(mService, 1, 1) = "R" Then
        wService_Name = YBIATAB0_Sql_Responsable(mService, cnADO, rsADO)
        Select Case mId$(mService, 1, 3)
            Case "R60", "R80":        wService_Printer = "INFO"
            Case Else: wService_Printer = "DCOM"
        End Select
        
    Else
        xElpTable.ID = "SAb_Param"
        xElpTable.K1 = "Compte_Unit"
        xElpTable.K2 = mService
        xElpTable.Method = "Seek="
        If tableElpTable_Read(xElpTable) = 0 Then
            wService_Name = mService & " : " & Trim(xElpTable.Name)
            wService_Printer = Trim(xElpTable.Memo)
        Else
            wService_Name = mService & " : ?????"
            wService_Printer = "INFO"
        End If
    End If
    arrService_Balance_Cumul(K, 0).ID = wService_Name
    iDevise = 0

    If blnService_Printer Then Printer_Set wService_Printer
    fgSelect_Reset
    fgSelect.Rows = 1
    fgSelect.FormatString = fgSelect_FormatString
    fgSelect.Visible = False

    For I = 1 To marrYBIACPT0_Nb
        If mService = wCOMREFCOR(I) Then
            blnOk = True
            xYBIACPT0 = marrYBIACPT0(I)
            If xYBIACPT0.SOLDECEN = 0 Then blnOk = False
            If xYBIACPT0.COMPTEFON = "4" And xYBIACPT0.SOLDECEN = 0 Then blnOk = False
            If blnOk Then
                If fctUser_Classe_Aut(xYBIACPT0.COMPTECLA) Then
                    fgSelect_DisplayLine I
'===========================================================================================
                    If arrService_Balance_Cumul(K, iDevise).Dev <> xYBIACPT0.COMPTEDEV Then
                        For iDevise = 1 To arrDevise_Nb
                            If xYBIACPT0.COMPTEDEV = arrDevise(iDevise) Then
                                arrService_Balance_Cumul(K, iDevise).ID = wService_Name
                                Exit For
                            End If
                        Next iDevise
                    End If
                    curX = Abs(xYBIACPT0.SOLDECEN)
                    If mId$(xYBIACPT0.COMPTEOBL, 1, 1) <> "9" Then
                        arrService_Balance_Cumul(K, iDevise).Bilan_Nb = arrService_Balance_Cumul(K, iDevise).Bilan_Nb + 1
                        If xYBIACPT0.SOLDECEN > 0 Then
                            arrService_Balance_Cumul(K, iDevise).Bilan_DB = arrService_Balance_Cumul(K, iDevise).Bilan_DB + curX
                        Else
                            arrService_Balance_Cumul(K, iDevise).Bilan_CR = arrService_Balance_Cumul(K, iDevise).Bilan_CR + curX
                        End If
                    Else
                         arrService_Balance_Cumul(K, iDevise).HorsBilan_Nb = arrService_Balance_Cumul(K, iDevise).HorsBilan_Nb + 1
                        If xYBIACPT0.SOLDECEN > 0 Then
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
        Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : impression " & mService)

        optBalance_YSOLDE0_J = True
        chkBalance_Détail = "1"
        chkBalance_Récap = "0"
        chkBalance_Récap_Bilan = "0"
        chkBalance_CSV = "0"
        chkBalance_Pays = "0"
        chkBalance_Compte_Soldé = "1"
        wBalance_Ok_Param = cmdBalance_Ok_Param
        Mid$(wBalance_Ok_Param, 1, 1) = "S"                      ' BALANCE avec contrôle STOCK
        cmdBalance_Ok_Print wBalance_Ok_Param, wService_Name
    End If
Next lstW_Index

End Sub

Public Sub arrService_Balance_Cumul_Z()
Dim I As Integer, K As Integer
ReDim arrService_Balance_Cumul(arrService_Nb, arrDevise_Nb)
    
For I = 0 To arrService_Nb
    For K = 0 To arrDevise_Nb
        arrService_Balance_Cumul(I, K).ID = arrService(I)
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
Dim xSql As String
Dim X20 As String
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "cmdBalance_Ok_Stock : chargement ZDORCPT0")

ReDim wDORCPTDMV(marrYBIACPT0_Nb)
For I = 1 To marrYBIACPT0_Nb
    wDORCPTDMV(I) = 0
Next I

xSql = "select DORCPTCOM,DORCPTDMV from " & paramIBM_Library_SAB & ".ZDORCPT0  order by DORCPTCOM"
Set rsADO = cnADO.Execute(xSql)

Do While Not rsADO.EOF
    X20 = rsADO("DORCPTCOM")
    For I = 1 To marrYBIACPT0_Nb
        If X20 = marrYBIACPT0(I).COMPTECOM Then
            wDORCPTDMV(I) = rsADO("DORCPTDMV")
            If wDORCPTDMV(I) > marrYBIACPT0(I).SOLDEDMO Then wDORCPTDMV(I) = marrYBIACPT0(I).SOLDEDMO
            
            Exit For
        End If
    Next I
    rsADO.MoveNext
Loop

blnZDORCPT_SQL_ODBC = True
End Sub
