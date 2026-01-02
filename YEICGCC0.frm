VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmYEICGCC0 
   AutoRedraw      =   -1  'True
   Caption         =   "Gestion des chèques circulants"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13575
   Icon            =   "YEICGCC0.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10530
   ScaleWidth      =   13575
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   0
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9852
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   17383
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Gestion des chèques circulants"
      TabPicture(0)   =   "YEICGCC0.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Paramétrage"
      TabPicture(1)   =   "YEICGCC0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraParam"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "."
      TabPicture(2)   =   "YEICGCC0.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fgList"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraCHQ"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraSelect_Options_L"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fraCHQ_Max"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lstW"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "fgLogV"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "fraLogV"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "fraSuivi"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "fraJRNENT0"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "fraSelect_Options_St"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cmdDetail_CLIBEN"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      Begin VB.CommandButton cmdDetail_CLIBEN 
         BackColor       =   &H00FFFFC0&
         Caption         =   "CLI+BEN ="
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   -63000
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   480
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Frame fraSelect_Options_St 
         BackColor       =   &H00F0FFFF&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   972
         Left            =   -70590
         TabIndex        =   144
         Top             =   4725
         Width           =   8000
         Begin VB.TextBox txtSelect_Options_PCI 
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   6720
            TabIndex        =   159
            Text            =   "162120"
            Top             =   420
            Width           =   972
         End
         Begin MSComCtl2.DTPicker txtSelect_Options_St_AMJMIN 
            Height          =   300
            Left            =   1860
            TabIndex        =   146
            Top             =   375
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   92012547
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin MSComCtl2.DTPicker txtSelect_Options_St_AMJMAX 
            Height          =   300
            Left            =   3300
            TabIndex        =   147
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   92012547
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.Label lblSelect_Options_PCI 
            BackColor       =   &H00F0FFFF&
            Caption         =   "PCI (6 chiffres)     chq BQ:162120"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   4905
            TabIndex        =   158
            Top             =   345
            Width           =   1320
         End
         Begin VB.Label lblSelect_Options_St 
            BackColor       =   &H00F0FFFF&
            Caption         =   "période du ... au ...."
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   240
            TabIndex        =   145
            Top             =   360
            Width           =   1452
         End
      End
      Begin VB.Frame fraJRNENT0 
         BackColor       =   &H00B0F0FF&
         Caption         =   "Journalisation"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   6900
         Left            =   -67740
         TabIndex        =   122
         Top             =   1230
         Visible         =   0   'False
         Width           =   3012
         Begin MSFlexGridLib.MSFlexGrid fgJRNENT0 
            Height          =   5000
            Left            =   120
            TabIndex        =   123
            Top             =   720
            Width           =   2650
            _ExtentX        =   4683
            _ExtentY        =   8811
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            RowHeightMin    =   285
            BackColor       =   16316664
            ForeColor       =   8388608
            BackColorFixed  =   10543359
            ForeColorFixed  =   0
            BackColorSel    =   12648384
            BackColorBkg    =   16777215
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            ScrollBars      =   2
            AllowUserResizing=   3
            FormatString    =   "< Champ          |< Valeur                    "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame fraParam 
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9012
         Left            =   -74760
         TabIndex        =   107
         Top             =   600
         Width           =   12972
         Begin VB.CheckBox chkEICGCCXXX 
            BackColor       =   &H00808080&
            Caption         =   "aide libellé"
            Height          =   280
            Left            =   7560
            TabIndex        =   143
            Top             =   7320
            Value           =   1  'Checked
            Width           =   3372
         End
         Begin VB.Frame fraParam_K 
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6132
            Left            =   7440
            TabIndex        =   110
            Top             =   600
            Width           =   5172
            Begin VB.TextBox txtParam_Action 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   972
               Left            =   1320
               MultiLine       =   -1  'True
               TabIndex        =   121
               Text            =   "YEICGCC0.frx":035E
               Top             =   960
               Width           =   3612
            End
            Begin VB.TextBox libParam_Action 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   1320
               TabIndex        =   118
               Top             =   480
               Width           =   1932
            End
            Begin VB.TextBox txtParam_K 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   1320
               TabIndex        =   116
               Top             =   2880
               Width           =   1932
            End
            Begin VB.CommandButton cmdParam_Delete 
               BackColor       =   &H00FF80FF&
               Caption         =   "Supprimer"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   115
               Top             =   4920
               Width           =   900
            End
            Begin VB.CommandButton cmdParam_Add 
               BackColor       =   &H000080FF&
               Caption         =   "Ajouter"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   1440
               Style           =   1  'Graphical
               TabIndex        =   114
               Top             =   4920
               Width           =   900
            End
            Begin VB.CommandButton cmdParam_Update 
               BackColor       =   &H0080FF80&
               Caption         =   "Enregistrer"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   2760
               Style           =   1  'Graphical
               TabIndex        =   113
               Top             =   4920
               Width           =   900
            End
            Begin VB.CommandButton cmdParam_Quit 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Abandonner"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   112
               Top             =   4920
               Width           =   900
            End
            Begin VB.TextBox txtParam_X 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   972
               Left            =   1320
               MultiLine       =   -1  'True
               TabIndex        =   111
               Text            =   "YEICGCC0.frx":0364
               Top             =   3480
               Width           =   3612
            End
            Begin VB.Label lblParam_K_Aide 
               Caption         =   "libellé associé"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   372
               Left            =   240
               TabIndex        =   120
               Top             =   3480
               Width           =   1092
            End
            Begin VB.Label lblParam_Action 
               Caption         =   "Action"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   120
               TabIndex        =   119
               Top             =   480
               Width           =   852
            End
            Begin VB.Label lblParam_K 
               Caption         =   "Code"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   240
               TabIndex        =   117
               Top             =   2880
               Width           =   1092
            End
         End
         Begin VB.ListBox lstParam_K 
            Height          =   5910
            Left            =   3720
            TabIndex        =   109
            Top             =   600
            Width           =   2700
         End
         Begin VB.ListBox lstParam_Action 
            Height          =   6105
            Left            =   480
            TabIndex        =   108
            Top             =   600
            Width           =   2700
         End
      End
      Begin VB.Frame fraSuivi 
         BackColor       =   &H00B0F0FF&
         Caption         =   "Saisie d'événement"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   2172
         Left            =   -68040
         TabIndex        =   99
         Top             =   1800
         Width           =   6012
         Begin VB.CommandButton cmdSuivi_Update 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Enregistrer"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   105
            Top             =   960
            Width           =   900
         End
         Begin VB.CommandButton cmdSuivi_Quit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Abandonner"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   104
            Top             =   1560
            Width           =   900
         End
         Begin VB.TextBox txtSuivi_Q 
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   2880
            MultiLine       =   -1  'True
            TabIndex        =   101
            Top             =   1080
            Width           =   1572
         End
         Begin VB.ComboBox cboSuivi_K 
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   1080
            Width           =   1212
         End
         Begin VB.Label lblSuivi_K 
            BackColor       =   &H00B0F0FF&
            Caption         =   "Evénement"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   720
            TabIndex        =   103
            Top             =   600
            Width           =   852
         End
         Begin VB.Label lblSuivi_Q 
            BackColor       =   &H00B0F0FF&
            Caption         =   "numéro du chèque"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3000
            TabIndex        =   102
            Top             =   600
            Width           =   1332
         End
      End
      Begin VB.Frame fraLogV 
         BackColor       =   &H00B0F0FF&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2100
         Left            =   -68280
         TabIndex        =   89
         Top             =   6120
         Width           =   6612
         Begin VB.ComboBox cboLogV_K2 
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Left            =   240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   93
            Top             =   600
            Width           =   2652
         End
         Begin VB.TextBox txtLogV_X 
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   612
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   91
            Top             =   960
            Width           =   6132
         End
         Begin VB.ComboBox cboLogV_K 
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Left            =   240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   240
            Width           =   2652
         End
         Begin MSComCtl2.DTPicker txtLogV_E 
            Height          =   300
            Left            =   4200
            TabIndex        =   96
            Top             =   600
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CheckBox        =   -1  'True
            CustomFormat    =   "dd  MM yyy"
            Format          =   92012547
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.Label lblLogV_E 
            BackColor       =   &H00B0F0FF&
            Caption         =   "Echéance"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3120
            TabIndex        =   97
            Top             =   600
            Width           =   852
         End
         Begin VB.Label lblLogV_X 
            BackColor       =   &H0080C0FF&
            Caption         =   "...."
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   492
            Left            =   240
            TabIndex        =   92
            Top             =   1560
            Width           =   6132
            WordWrap        =   -1  'True
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgLogV 
         Height          =   2100
         Left            =   -74880
         TabIndex        =   88
         Top             =   6120
         Width           =   6612
         _ExtentX        =   11668
         _ExtentY        =   3704
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   33023
         ForeColorFixed  =   -2147483633
         BackColorBkg    =   14745599
         FormatString    =   $"YEICGCC0.frx":036A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ListBox lstW 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   -72480
         TabIndex        =   87
         Top             =   2760
         Visible         =   0   'False
         Width           =   4212
      End
      Begin VB.Frame fraCHQ_Max 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4245
         Left            =   -73320
         TabIndex        =   81
         Top             =   3720
         Visible         =   0   'False
         Width           =   11535
         Begin VB.CommandButton cmdCHQ_MAX_Quit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Abandonner"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   9720
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   480
            Width           =   1020
         End
         Begin VB.Label libCHQ_EICGCCECPT_X 
            BackColor       =   &H00E0FFE0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lib"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   612
            Left            =   9240
            TabIndex        =   86
            Top             =   2880
            Width           =   2172
            WordWrap        =   -1  'True
         End
         Begin VB.Label libCHQ_EICGCCECPT 
            BackColor       =   &H00E0FFE0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   9240
            TabIndex        =   85
            Top             =   3720
            Width           =   2172
         End
         Begin VB.Label libCHQ_EICGCCEMT 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00B0F0FF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   9240
            TabIndex        =   84
            Top             =   1440
            Width           =   2172
         End
         Begin VB.Image imgCHQ_Max 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   4008
            Left            =   120
            Stretch         =   -1  'True
            Top             =   120
            Width           =   9000
         End
         Begin VB.Label libCHQ_EICGCCXNOM 
            BackColor       =   &H00B0F0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lib"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   612
            Left            =   9240
            TabIndex        =   83
            Top             =   2040
            Width           =   2172
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraSelect_Options_L 
         BackColor       =   &H00F0FFFF&
         Height          =   972
         Left            =   -68040
         TabIndex        =   78
         Top             =   840
         Width           =   6012
         Begin VB.ComboBox cboEICGCCLOGK 
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   1440
            Sorted          =   -1  'True
            TabIndex        =   79
            Text            =   "cboNOTPAYLOGK"
            Top             =   360
            Width           =   3612
         End
         Begin VB.Label lblEICGCCLOGK 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Actions"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   480
            TabIndex        =   80
            Top             =   480
            Width           =   1452
         End
      End
      Begin VB.Frame fraCHQ 
         BackColor       =   &H00F0FFFF&
         Caption         =   "cliquer sur l'image pour l'agrandir"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   5820
         Left            =   -74880
         TabIndex        =   77
         Top             =   360
         Visible         =   0   'False
         Width           =   6550
         Begin VB.Image imgCHQ 
            Height          =   2532
            Left            =   240
            Stretch         =   -1  'True
            Top             =   3120
            Width           =   5892
         End
         Begin VB.Image imgCHQ_Verso 
            Height          =   2532
            Left            =   240
            Stretch         =   -1  'True
            Top             =   360
            Width           =   5892
         End
      End
      Begin VB.Frame fraTab0 
         Height          =   9420
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   13296
         Begin VB.Frame fraYEICGCCLOG 
            BackColor       =   &H00E0FFFF&
            Height          =   5532
            Left            =   600
            TabIndex        =   124
            Top             =   2040
            Visible         =   0   'False
            Width           =   6036
            Begin VB.Label txtYEICGCCLOGX 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1332
               Left            =   1920
               TabIndex        =   142
               Top             =   3600
               Width           =   3732
               WordWrap        =   -1  'True
            End
            Begin VB.Label txtYEICGCCLOGE 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   1920
               TabIndex        =   141
               Top             =   3200
               Width           =   1812
            End
            Begin VB.Label txtYEICGCCLOGA 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   1920
               TabIndex        =   140
               Top             =   2800
               Width           =   1812
            End
            Begin VB.Label txtYEICGCCLOGI 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   1920
               TabIndex        =   139
               Top             =   2400
               Width           =   1812
            End
            Begin VB.Label txtYEICGCCLOGK 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   1920
               TabIndex        =   138
               Top             =   2000
               Width           =   1812
            End
            Begin VB.Label txtYEICGCCLOGS 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   1920
               TabIndex        =   137
               Top             =   1600
               Width           =   1812
            End
            Begin VB.Label txtYEICGCCLOGU 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   1920
               TabIndex        =   136
               Top             =   1200
               Width           =   1812
            End
            Begin VB.Label txtYEICGCCLOGH 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   1920
               TabIndex        =   135
               Top             =   800
               Width           =   1812
            End
            Begin VB.Label txtYEICGCCLOGD 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   1920
               TabIndex        =   134
               Top             =   400
               Width           =   1812
            End
            Begin VB.Label lblYEICGCCLOGX 
               BackColor       =   &H00E0FFFF&
               Caption         =   "commentaire"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   480
               TabIndex        =   133
               Top             =   3600
               Width           =   1452
            End
            Begin VB.Label lblYEICGCCLOGE 
               BackColor       =   &H00E0FFFF&
               Caption         =   "échéance"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   480
               TabIndex        =   132
               Top             =   3204
               Width           =   1452
            End
            Begin VB.Label lblYEICGCCLOGA 
               BackColor       =   &H00E0FFFF&
               Caption         =   "statut"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   480
               TabIndex        =   131
               Top             =   2800
               Width           =   1452
            End
            Begin VB.Label lblYEICGCCLOGI 
               BackColor       =   &H00E0FFFF&
               Caption         =   "identifiant"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   480
               TabIndex        =   130
               Top             =   2400
               Width           =   1452
            End
            Begin VB.Label lblYEICGCCLOGK 
               BackColor       =   &H00E0FFFF&
               Caption         =   "code événement"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   480
               TabIndex        =   129
               Top             =   2000
               Width           =   1452
            End
            Begin VB.Label lblYEICGCCLOGS 
               BackColor       =   &H00E0FFFF&
               Caption         =   "séquence"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   480
               TabIndex        =   128
               Top             =   1600
               Width           =   1452
            End
            Begin VB.Label lblYEICGCCLOGU 
               BackColor       =   &H00E0FFFF&
               Caption         =   "utilisateur"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   480
               TabIndex        =   127
               Top             =   1200
               Width           =   1452
            End
            Begin VB.Label lblYEICGCCLOGH 
               BackColor       =   &H00E0FFFF&
               Caption         =   "heure"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   480
               TabIndex        =   126
               Top             =   800
               Width           =   1452
            End
            Begin VB.Label lblYEICGCCLOGD 
               BackColor       =   &H00E0FFFF&
               Caption         =   "date"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   480
               TabIndex        =   125
               Top             =   400
               Width           =   1452
            End
         End
         Begin VB.Frame fraDetail 
            BackColor       =   &H00FAFAFA&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7980
            Left            =   6572
            TabIndex        =   10
            Top             =   1440
            Visible         =   0   'False
            Width           =   6036
            Begin VB.CommandButton cmdDetail_UpdateVO 
               BackColor       =   &H00FF80FF&
               Caption         =   "associer 1 vignette"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   3120
               Style           =   1  'Graphical
               TabIndex        =   76
               Top             =   7100
               Width           =   900
            End
            Begin VB.Frame fraDetail_X 
               BackColor       =   &H00D0FFFF&
               Caption         =   "Bénéficiaire"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2172
               Left            =   240
               TabIndex        =   52
               Top             =   3000
               Width           =   5532
               Begin VB.ComboBox cboDetail_EICGCCXECO 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   324
                  Left            =   3360
                  Style           =   2  'Dropdown List
                  TabIndex        =   98
                  Top             =   1800
                  Width           =   2052
               End
               Begin VB.TextBox txtDetail_EICGCCXNOM 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Left            =   960
                  TabIndex        =   54
                  Top             =   1440
                  Width           =   3492
               End
               Begin VB.TextBox txtDetail_EICGCCXECO 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Left            =   960
                  TabIndex        =   53
                  Top             =   1800
                  Width           =   3492
               End
               Begin VB.Label lblDetail_EICRICREF 
                  BackColor       =   &H00D0FFFF&
                  Caption         =   "Référence  remettant"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   612
                  Left            =   120
                  TabIndex        =   157
                  Top             =   600
                  Width           =   732
               End
               Begin VB.Label libDetail_EICRICDO6 
                  BackColor       =   &H00E0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   492
                  Left            =   960
                  TabIndex        =   156
                  Top             =   850
                  Width           =   4332
                  WordWrap        =   -1  'True
               End
               Begin VB.Label libDetail_EICRICREF 
                  BackColor       =   &H00E0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   960
                  TabIndex        =   155
                  Top             =   550
                  Width           =   4332
               End
               Begin VB.Label libDetail_EICGCCXBQ 
                  BackColor       =   &H00E0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   960
                  TabIndex        =   65
                  Top             =   240
                  Width           =   972
               End
               Begin VB.Label lblDetail_EICGCCXBQ 
                  BackColor       =   &H00D0FFFF&
                  Caption         =   "Banque / Compte"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   156
                  TabIndex        =   60
                  Top             =   240
                  Width           =   612
               End
               Begin VB.Label libDetail_EICGCCXCPT 
                  BackColor       =   &H00E0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   2040
                  TabIndex        =   59
                  Top             =   240
                  Width           =   1692
               End
               Begin VB.Label libDetail_EICGCCXID 
                  BackColor       =   &H00E0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   4320
                  TabIndex        =   58
                  Top             =   240
                  Width           =   972
               End
               Begin VB.Label lblDetail_EICGCCXID 
                  BackColor       =   &H00D0FFFF&
                  Caption         =   "Id"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   3960
                  TabIndex        =   57
                  Top             =   240
                  Width           =   216
               End
               Begin VB.Label lblDetail_EICGCCNOM 
                  BackColor       =   &H00D0FFFF&
                  Caption         =   "nom"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   240
                  TabIndex        =   56
                  Top             =   1440
                  Width           =   612
               End
               Begin VB.Label lblDetail_EICGCCXECO 
                  BackColor       =   &H00D0FFFF&
                  Caption         =   "motif éco"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   240
                  TabIndex        =   55
                  Top             =   1800
                  Width           =   732
               End
            End
            Begin VB.ComboBox cboDetail_EICGCCSTA 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   324
               Left            =   4080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   47
               Top             =   7200
               Width           =   1500
            End
            Begin VB.Frame fraDetail_K 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Contrôles"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1932
               Left            =   240
               TabIndex        =   15
               Top             =   5160
               Width           =   5532
               Begin VB.ComboBox cboDetail_EICGCCSTAK 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   324
                  Left            =   3840
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   48
                  Top             =   1440
                  Width           =   1500
               End
               Begin VB.ComboBox cboDetail_EICGCCKLAB 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   324
                  Left            =   1080
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   46
                  Top             =   1500
                  Width           =   1700
               End
               Begin VB.ComboBox cboDetail_EICGCCKEND 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   324
                  Left            =   1080
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   45
                  Top             =   1100
                  Width           =   1700
               End
               Begin VB.ComboBox cboDetail_EICGCCKSIG 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   324
                  Left            =   1080
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   44
                  Top             =   700
                  Width           =   1700
               End
               Begin VB.ComboBox cboDetail_EICGCCKMT 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   324
                  Left            =   1080
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   43
                  Top             =   300
                  Width           =   1700
               End
               Begin MSComCtl2.DTPicker txtDetail_EICGCCEAMJ 
                  Height          =   300
                  Left            =   3840
                  TabIndex        =   41
                  Top             =   600
                  Width           =   1500
                  _ExtentX        =   2646
                  _ExtentY        =   529
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CalendarBackColor=   16777215
                  CalendarForeColor=   0
                  CalendarTitleBackColor=   8421504
                  CalendarTitleForeColor=   16777215
                  CalendarTrailingForeColor=   12632256
                  CheckBox        =   -1  'True
                  CustomFormat    =   "dd  MM yyy"
                  Format          =   92012547
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin VB.Label lblDetail_EICGCCKLAB 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "LAB"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   64
                  Top             =   1560
                  Width           =   975
               End
               Begin VB.Label lblDetail_EICGCCKKEND 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "endos"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   63
                  Top             =   1200
                  Width           =   975
               End
               Begin VB.Label lblDetail_EICGCCKSIG 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "signature"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   62
                  Top             =   840
                  Width           =   975
               End
               Begin VB.Label lblDetail_EICGCCKMT 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "montant"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   150
                  TabIndex        =   61
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label lblDetail_EICGCCSTAK 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "statut contrôles"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   3840
                  TabIndex        =   49
                  Top             =   1080
                  Width           =   1332
               End
               Begin VB.Label lblDetail_EICGCCEAMJ 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "date émission"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   3960
                  TabIndex        =   42
                  Top             =   240
                  Width           =   1092
               End
            End
            Begin VB.Frame fraDetail_V 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Vignette"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1092
               Left            =   240
               TabIndex        =   14
               Top             =   1920
               Width           =   5532
               Begin VB.TextBox txtDetail_EICGCCVEXT 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Left            =   3720
                  TabIndex        =   37
                  Top             =   720
                  Width           =   1572
               End
               Begin VB.TextBox txtDetail_EICGCCVINT 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Left            =   1440
                  TabIndex        =   36
                  Top             =   720
                  Width           =   1692
               End
               Begin VB.Label libDetail_EICGCCVJPG 
                  BackColor       =   &H00E0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   40
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label libDetail_EICGCCVREM 
                  BackColor       =   &H00E0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   4320
                  TabIndex        =   39
                  Top             =   360
                  Width           =   735
               End
               Begin VB.Label libDetail_EICGCCVAMJ 
                  BackColor       =   &H00E0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   960
                  TabIndex        =   38
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label lblDetail_EICGCCVEXT 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "externe"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   3120
                  TabIndex        =   35
                  Top             =   720
                  Width           =   576
               End
               Begin VB.Label lblDetail_EICGCCVINT 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "réf archive interne"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   150
                  TabIndex        =   34
                  Top             =   720
                  Width           =   1452
               End
               Begin VB.Label lblDetail_EICGCCVJPG 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "numéro"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   2040
                  TabIndex        =   33
                  Top             =   360
                  Width           =   615
               End
               Begin VB.Label lblDetail_EICGCCVREM 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "remise"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   32
                  Top             =   360
                  Width           =   615
               End
               Begin VB.Label lblDetail_EICGCCVAMJ 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "date scan"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   150
                  TabIndex        =   31
                  Top             =   360
                  Width           =   732
               End
            End
            Begin VB.CommandButton cmdDetail_Update 
               BackColor       =   &H0080FF80&
               Caption         =   "Enregistrer"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   2160
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   7100
               Width           =   900
            End
            Begin VB.CommandButton cmdDetail_Quit 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Abandonner"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   7100
               Width           =   900
            End
            Begin VB.CommandButton cmdDetail_Action 
               BackColor       =   &H000080FF&
               Caption         =   "Gérer"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   1200
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   7100
               Width           =   900
            End
            Begin VB.Label libDetail_EICGCCECLI_Resp 
               BackColor       =   &H00E0FFFF&
               Caption         =   "lib"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3240
               TabIndex        =   106
               Top             =   1200
               Width           =   2412
            End
            Begin VB.Label lblDetail_EICGCCID 
               Alignment       =   2  'Center
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "ID"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF00FF&
               Height          =   300
               Left            =   360
               TabIndex        =   75
               Top             =   120
               Width           =   732
            End
            Begin VB.Label lblDetail_EICGCCUUSR 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Id"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   240
               TabIndex        =   50
               Top             =   7680
               Width           =   5532
            End
            Begin VB.Label libDetail_EICGCCECPT_X 
               BackColor       =   &H00E0FFFF&
               Caption         =   "lib"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3240
               TabIndex        =   30
               Top             =   840
               Width           =   2415
            End
            Begin VB.Label libDetail_EICGCCECLI_X 
               BackColor       =   &H00E0FFFF&
               Caption         =   "lib"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3240
               TabIndex        =   29
               Top             =   480
               Width           =   2415
            End
            Begin VB.Label libDetail_EICGCCECLI 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1200
               TabIndex        =   28
               Top             =   480
               Width           =   975
            End
            Begin VB.Label lblDetail_EICGCCECLI 
               BackColor       =   &H00FAFAFA&
               Caption         =   "client"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   360
               TabIndex        =   27
               Top             =   480
               Width           =   855
            End
            Begin VB.Label libDetail_EICGCCEIND 
               Alignment       =   2  'Center
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4440
               TabIndex        =   26
               Top             =   1560
               Width           =   375
            End
            Begin VB.Label libDetail_EICGCCECHQ 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1200
               TabIndex        =   25
               Top             =   1560
               Width           =   975
            End
            Begin VB.Label libDetail_EICGCCEMT 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1200
               TabIndex        =   24
               Top             =   1200
               Width           =   1575
            End
            Begin VB.Label libDetail_EICGCCECPT 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1200
               TabIndex        =   23
               Top             =   840
               Width           =   1815
            End
            Begin VB.Label lblDetail_EICGCCEIND 
               BackColor       =   &H00FAFAFA&
               Caption         =   "indice de circulation"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2760
               TabIndex        =   22
               Top             =   1560
               Width           =   1575
            End
            Begin VB.Label lblDetail_EICGCCECHQ 
               BackColor       =   &H00FAFAFA&
               Caption         =   "numéro"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   360
               TabIndex        =   21
               Top             =   1560
               Width           =   855
            End
            Begin VB.Label lblDetail_EICGCCEMT 
               BackColor       =   &H00FAFAFA&
               Caption         =   "montant"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   360
               TabIndex        =   20
               Top             =   1200
               Width           =   855
            End
            Begin VB.Label lblDetail_EICGCCECPT 
               BackColor       =   &H00FAFAFA&
               Caption         =   "compte"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   360
               TabIndex        =   19
               Top             =   840
               Width           =   855
            End
            Begin VB.Label lblDetail_EICGCCAMJ 
               Alignment       =   2  'Center
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "AMJ"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4320
               TabIndex        =   18
               Top             =   156
               Width           =   1215
            End
            Begin VB.Label lblDetail_EICGCCOPE 
               Alignment       =   2  'Center
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "OPE"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   2760
               TabIndex        =   17
               Top             =   156
               Width           =   1332
            End
            Begin VB.Label lblDetail_EICGCCETB 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "ETB"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   1200
               TabIndex        =   16
               Top             =   156
               Width           =   1332
            End
         End
         Begin VB.ComboBox cboSelect_SQL 
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   9240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   240
            Width           =   3732
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   555
            Left            =   10440
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   720
            Width           =   1335
         End
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00F0FFFF&
            Height          =   1212
            Left            =   240
            TabIndex        =   6
            Top             =   120
            Visible         =   0   'False
            Width           =   8832
            Begin VB.Frame fraSelect_Options_1 
               BackColor       =   &H00F0FFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   972
               Left            =   120
               TabIndex        =   66
               Top             =   120
               Width           =   6732
               Begin VB.ComboBox cboSelect_EICGCCEIND 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   312
                  Left            =   3960
                  Sorted          =   -1  'True
                  TabIndex        =   152
                  Text            =   "AUT"
                  Top             =   600
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_EICGCCSTA 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   312
                  Left            =   2880
                  Sorted          =   -1  'True
                  TabIndex        =   149
                  Text            =   "AUT"
                  Top             =   600
                  Width           =   696
               End
               Begin VB.TextBox txtSelect_EICGCCID 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Left            =   5640
                  TabIndex        =   95
                  Top             =   120
                  Width           =   852
               End
               Begin VB.TextBox txtSelect_EICGCCXNOM 
                  Height          =   288
                  Left            =   2880
                  TabIndex        =   74
                  Top             =   120
                  Width           =   1692
               End
               Begin VB.TextBox txtSelect_EICGCCECHQ 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Left            =   5640
                  TabIndex        =   73
                  Top             =   600
                  Width           =   852
               End
               Begin VB.TextBox txtSelect_EICGCCECPT 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Left            =   840
                  TabIndex        =   69
                  Top             =   600
                  Width           =   1092
               End
               Begin VB.TextBox txtSelect_EICGCCECLI 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Left            =   840
                  TabIndex        =   67
                  Top             =   120
                  Width           =   972
               End
               Begin VB.Label lblSelect_EICGCCEIND 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "IC"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   3720
                  TabIndex        =   151
                  Top             =   650
                  Width           =   252
               End
               Begin VB.Label lblSelect_EICGCCSTA 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "statut"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   2040
                  TabIndex        =   148
                  Top             =   650
                  Width           =   612
               End
               Begin VB.Label lblSelect_EICGCCID 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Identifiant"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   4800
                  TabIndex        =   94
                  Top             =   120
                  Width           =   732
               End
               Begin VB.Label lblSelect_EICGCCXNOM 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Bénéficiaire"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   1920
                  TabIndex        =   72
                  Top             =   120
                  Width           =   972
               End
               Begin VB.Label lblSelect_EICGCCECHQ 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "chèque"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   4920
                  TabIndex        =   71
                  Top             =   600
                  Width           =   612
               End
               Begin VB.Label lblSelect_EICGCCXeCPT 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "compte"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   120
                  TabIndex        =   70
                  Top             =   650
                  Width           =   612
               End
               Begin VB.Label lblSelect_EICGCCXECLI 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "client"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   120
                  TabIndex        =   68
                  Top             =   120
                  Width           =   612
               End
            End
            Begin MSComCtl2.DTPicker txtSelect_EICGCCAMJ 
               Height          =   300
               Left            =   7080
               TabIndex        =   9
               Top             =   840
               Width           =   1452
               _ExtentX        =   2566
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CheckBox        =   -1  'True
               CustomFormat    =   "dd  MM yyy"
               Format          =   92012547
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_EICGCCAMJ_Min 
               Height          =   300
               Left            =   7080
               TabIndex        =   150
               Top             =   400
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CheckBox        =   -1  'True
               CustomFormat    =   "dd  MM yyy"
               Format          =   92012547
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_EICGCCAMJ 
               BackColor       =   &H00F0FFFF&
               Caption         =   "date de comptabilisation"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   6960
               TabIndex        =   51
               Top             =   120
               Width           =   1812
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   8028
            Left            =   240
            TabIndex        =   5
            Top             =   1320
            Visible         =   0   'False
            Width           =   12912
            _ExtentX        =   22781
            _ExtentY        =   14155
            _Version        =   393216
            Rows            =   1
            Cols            =   11
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   -2147483633
            ForeColor       =   12582912
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483637
            BackColorSel    =   12648384
            BackColorBkg    =   -2147483633
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"YEICGCC0.frx":0462
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgList 
         Height          =   2700
         Left            =   -68280
         TabIndex        =   153
         Top             =   360
         Visible         =   0   'False
         Width           =   6612
         _ExtentX        =   11668
         _ExtentY        =   4763
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColor       =   14741744
         BackColorFixed  =   12582912
         ForeColorFixed  =   -2147483633
         BackColorBkg    =   14741744
         FormatString    =   "<Date           |> Dossier    |> Montant              |<Motif économique                                        "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
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
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13080
      Picture         =   "YEICGCC0.frx":0548
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
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
   Begin VB.Menu mnufgSelect 
      Caption         =   "mnufgSelect"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuLogV 
      Caption         =   "mnuLogV"
      Visible         =   0   'False
      Begin VB.Menu mnuLogV_Annuler 
         Caption         =   "Annuler"
      End
      Begin VB.Menu mnuLogV_Valider 
         Caption         =   "Valider"
      End
      Begin VB.Menu mnuLogV_Culot 
         Caption         =   "Valider au culot"
      End
   End
End
Attribute VB_Name = "frmYEICGCC0"
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
Dim x As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim YEICGCC0_Aut As typeAuthorization
Dim blnAuto As Boolean, blnError As Boolean
Dim cmdSelect_SQL_K As String

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean


'______________________________________________________________________

Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long
Dim xYEICGCC0 As typeYEICGCC0, newYEICGCC0 As typeYEICGCC0, oldYEICGCC0 As typeYEICGCC0
Dim arrYEICGCC0() As typeYEICGCC0, arrYEICGCC0_Nb As Long, arrYEICGCC0_Max As Long, arrYEICGCC0_Index As Long
Dim lastYEICGCC0 As typeYEICGCC0
Dim selYEICGCC0() As typeYEICGCC0, selYEICGCC0_Nb As Integer, selYEICGCC0_Max As Integer
Dim zYEICGCC0 As typeYEICGCC0
Dim voYEICGCC0 As typeYEICGCC0
Dim listYEICGCC0 As typeYEICGCC0

Dim mEICGCCID As Long, mEICGCCID_0 As Long
Dim mEICGCCXXX As String
Dim oldYEICGCCLOG As typeYEICGCCLOG, newYEICGCCLOG As typeYEICGCCLOG, xYEICGCCLOG As typeYEICGCCLOG
Dim arrYEICGCCLOG() As typeYEICGCCLOG, arrYEICGCCLOG_Nb As Long, arrYEICGCCLOG_Max As Long, arrYEICGCCLOG_Index As Long
Dim updYEICGCCLOG As typeYEICGCCLOG
Dim selYEICGCCLOG() As typeYEICGCCLOG, selYEICGCCLOG_Nb As Long, selYEICGCCLOG_Max As Long, selYEICGCCLOG_Index As Long

Dim blnDetail_UpdateVO As Boolean

Dim blnDos_Log As Boolean

Dim fgLogV_FormatString As String, fgLogV_K As Integer
Dim fgLogV_RowDisplay As Integer, fgLogV_RowClick As Integer, fgLogV_ColClick As Integer
Dim fgLogV_ColorClick As Long, fgLogV_ColorDisplay As Long
Dim fgLogV_Sort1 As Integer, fgLogV_Sort2 As Integer
Dim fgLogV_SortAD As Integer, fgLogV_Sort1_Old As Integer
Dim fgLogV_arrIndex As Integer
Dim blnfgLogV_DisplayLine As Boolean

'______________________________________________________________________
Dim cnAdo_CHQ_ARCHIVE As New ADODB.Connection
Dim xCHQ_SCAN As typeCHQ_SCAN

Dim Old_YBIATAB0 As typeYBIATAB0, New_YBIATAB0 As typeYBIATAB0
Dim Action_YBIATAB0 As typeYBIATAB0

Dim xZBASTAB0 As typeZBASTAB0
Dim mCLIENARES As String, mMNURUTUTI As String, mMNUUTIMAI As String
Dim mCLIENARES_Lib As String
'______________________________________________________________________

Dim xJRNENT0 As typeJRNENT0
Dim arrEICGCCXECO() As String
Dim mAMJ_7Past As String, mAMJ_7Ante As String
Dim lstParam_Action_ListIndex As Long


Dim fgList_FormatString As String, fgList_K As Integer
Dim fgList_RowDisplay As Integer, fgList_RowClick As Integer, fgList_ColClick As Integer
Dim fgList_ColorClick As Long, fgList_ColorDisplay As Long
Dim fgList_Sort1 As Integer, fgList_Sort2 As Integer
Dim fgList_SortAD As Integer, fgList_Sort1_Old As Integer
Dim fgList_arrIndex As Integer
Dim blnfgList_DisplayLine As Boolean

Dim mXls1_Row As Long, mXls1_Col As Long, mXls2_Row As Long, mXls2_Col As Long, mXls2_Row_Cli As Long
Dim mXls2_row_T As Long, mXls1_Row_C As Long, mXls1_Row_T As Long

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel

Private Sub fgList_Display(lEICGCCECLI As String, lEICGCCXNOM As String)
Dim wColor As Long, xSQL As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgList.Visible = False
fgList_Reset

fgList.Rows = 1
fgList.FormatString = fgList_FormatString
fgList.Row = 0

currentAction = "fgList_Display"


xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
     & " where EICGCCECLI = '" & lEICGCCECLI & "'" _
     & " and EICGCCXNOM = '" & Trim(lEICGCCXNOM) & "'" _
     & " order by EICGCCAMJ"
     
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYEICGCC0_GetBuffer(rsSab, listYEICGCC0)
    fgList.Rows = fgList.Rows + 1
    fgList.Row = fgList.Rows - 1
    fgList_DisplayLine I, True
   
    rsSab.MoveNext
Loop
Set rsSab = Nothing


fgList.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYEICGCC0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub fgList_DisplayLine(lIndex As Long, blnYEICGCC0 As Boolean)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim xSQL As String
On Error Resume Next
 Select Case listYEICGCC0.EICGCCSTA
    Case Is = "V", "@": wColor = RGB(32, 96, 32)
    Case Is = "A", "I", "R": wColor = RGB(128, 128, 128)
    Case Else: wColor = RGB(64, 64, 128)
        If listYEICGCC0.EICGCCDOS = 0 Then
            wColor = vbRed
        Else
            If listYEICGCC0.EICGCCVJPG = 0 Then
                If listYEICGCC0.EICGCCAMJ < mAMJ_7Ante Then
                    wColor = vbRed
                Else
                    wColor = vbMagenta 'RGB(255, 96, 255)
                End If
            Else
                If listYEICGCC0.EICGCCSTAK = " " Then
                    Select Case listYEICGCC0.EICGCCSTAK
                        Case "X": XPrt.ForeColor = vbMagenta
                        Case Else: XPrt.ForeColor = vbBlue
                    End Select
                End If
            End If
        End If
End Select

fgList.Col = 0: fgList.Text = dateImp10(listYEICGCC0.EICGCCAMJ)
fgList.CellForeColor = wColor
fgList.Col = 1: fgList.Text = listYEICGCC0.EICGCCID
fgList.CellForeColor = wColor
fgList.Col = 2: fgList.Text = Format(listYEICGCC0.EICGCCEMT, "### ### ### ###.00") & " "
fgList.CellForeColor = wColor
fgList.Col = 3: fgList.Text = listYEICGCC0.EICGCCXECO
fgList.CellForeColor = wColor

End Sub
Public Sub fgList_ForeColor(lColor As Long)
For I = 0 To fgList_arrIndex
  fgList.Col = I: fgList.CellForeColor = lColor
Next I

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

Public Sub fgList_SortX(lK As Integer)
Dim I As Integer, x As String, wIndex As Long

For I = 1 To fgList.Rows - 1
    fgList.Row = I
    fgList.Col = fgList_arrIndex
    wIndex = Val(fgList.Text)
    Select Case lK
        Case 2: x = Format$(arrYEICGCC0(wIndex).EICGCCEMT, "000000000000000.00")
        Case 1: x = Format$(arrYEICGCC0(wIndex).EICGCCID, "00000000000")
    End Select
    fgList.Col = fgList_arrIndex - 1
    fgList.Text = x
Next I

fgList_Sort1 = fgList_arrIndex - 1: fgList_Sort2 = fgList_arrIndex - 1
fgList_Sort
End Sub



Private Sub fgSelect_Display()
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Row = 0

currentAction = "fgSelect_Display"
    
For I = 1 To arrYEICGCC0_Nb
         
    xYEICGCC0 = arrYEICGCC0(I)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine I, True
    
Next I

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYEICGCC0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub cmdYEICGCCLOG_New()
Dim V

App_Debug = "cmdYEICGCCLOG_New"
newYEICGCCLOG.EICGCCLOGH = time_Hms            ' heure maj
newYEICGCCLOG.EICGCCLOGS = newYEICGCCLOG.EICGCCLOGS + 1

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
 V = sqlYEICGCCLOG_Insert(newYEICGCCLOG)

'________________________________________________________________________________

If Not IsNull(V) Then GoTo Error_MsgBox
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, Me.Name & " ~ " & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Sub


Private Sub cmdSelect_SQL_YEICGCCLOG()
Dim V
Dim x As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_YEICGCCLOG"
blnOk = False
xWhere = ""

'Call DTPicker_Control(txtSelect_EICGCCAMJ, wAmjMax)
'If optSelect_EICGCCAMJ_E Then
'    xWhere = " and   EICGCCLOGD = " & wAmjMax
'Else
'    If optSelect_EICGCCAMJ_S Then
'        xWhere = xWhere & " and   EICGCCLOGD >= " & wAmjMax
'    End If
'End If
xWhere = Replace(cmdSelect_SQL_AMJ, "EICGCCAMJ", "EICGCCLOGD")

x = Trim(cboEICGCCLOGK)
If x <> "" Then xWhere = xWhere & " and   EICGCCLOGK = '" & x & "'"


If xWhere = "" Then
    Call MsgBox("Préciser au moins un critère de filtrage", vbExclamation, "EICGCC : recherche")
    Exit Sub
End If
Mid$(xWhere, 1, 7) = " where "
arrYEICGCCLOG_SQL ".YEICGCCLOG " & xWhere & " order by EICGCCLOGD,EICGCCLOGH,EICGCCLOGU,EICGCCLOGS"
   

fgSelect_Display_YEICGCCLOG


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrYEICGCCLOG_SQL(xWhere As String)
Dim V
Dim x As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYEICGCCLOG(101)
arrYEICGCCLOG_Max = 100: arrYEICGCCLOG_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYEICGCCLOG_GetBuffer(rsSab, xYEICGCCLOG)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYEICGCCLOG.fgselect_Display"
        '' Exit Sub
     Else
         arrYEICGCCLOG_Nb = arrYEICGCCLOG_Nb + 1
         If arrYEICGCCLOG_Nb > arrYEICGCCLOG_Max Then
             arrYEICGCCLOG_Max = arrYEICGCCLOG_Max + 100
             ReDim Preserve arrYEICGCCLOG(arrYEICGCCLOG_Max)
         End If
         
         arrYEICGCCLOG(arrYEICGCCLOG_Nb) = xYEICGCCLOG
    End If
    rsSab.MoveNext
Loop


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub fraDetail_Display()
Dim I As Long
Dim xSQL As String
Dim wEICGCCSTA As String
Dim xCLIENARES As String

On Error GoTo Error_Handler

mMNUUTIMAI = ""

blnControl = False
currentAction = "fraDetail_Display"
SSTab1.Tab = 0
fraDetail.Visible = False
fraCHQ.Visible = False
cmdDetail_UpdateVO.Visible = False
cmdDetail_Action.Visible = False
cboDetail_EICGCCXECO.Visible = False
cboDetail_EICGCCXECO.ListIndex = -1
txtDetail_EICGCCXECO.Visible = True

'fraDetail_Update.Enabled = False
lblDetail_EICGCCID = xYEICGCC0.EICGCCID
lblDetail_EICGCCUUSR = xYEICGCC0.EICGCCID & " - mise à jour : " & xYEICGCC0.EICGCCUUSR & "  " & dateImp10(xYEICGCC0.EICGCCUAMJ) & "  " & timeImp8(xYEICGCC0.EICGCCUHMS)
lblDetail_EICGCCETB = "Service : " & xYEICGCC0.EICGCCSER & "  " & xYEICGCC0.EICGCCSSE
lblDetail_EICGCCOPE = xYEICGCC0.EICGCCOPE & "  " & xYEICGCC0.EICGCCDOS
lblDetail_EICGCCAMJ = dateImp10(xYEICGCC0.EICGCCAMJ)


libDetail_EICGCCECLI = xYEICGCC0.EICGCCECLI
xSQL = "select CLIENARA1, COMPTEINT, CLIENARES from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " Where COMPTECOM = '" & xYEICGCC0.EICGCCECPT & "'"
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    libDetail_EICGCCECLI_X = Trim(rsSab("CLIENARA1"))
    libDetail_EICGCCECPT_X = rsSab("COMPTEINT")
    xCLIENARES = rsSab("CLIENARES")
Else
    libDetail_EICGCCECLI_X = "?????"
    libDetail_EICGCCECPT_X = "?????"
    xCLIENARES = ""
End If

Call sqlCLIENARES(xCLIENARES)
libDetail_EICGCCECLI_Resp = xCLIENARES & " - " & mCLIENARES_Lib

If IsNumeric(xYEICGCC0.EICGCCECPT) Then
    libDetail_EICGCCECPT = Format(xYEICGCC0.EICGCCECPT, "@@@@@ @@@ @@@ @@@@@@@")
Else
    libDetail_EICGCCECPT = xYEICGCC0.EICGCCECPT
End If

libDetail_EICGCCEMT = Format$(xYEICGCC0.EICGCCEMT, "### ### ##0.00")
libDetail_EICGCCECHQ = xYEICGCC0.EICGCCECHQ
libDetail_EICGCCEIND = xYEICGCC0.EICGCCEIND

libDetail_EICGCCXBQ = Trim(xYEICGCC0.EICGCCXBQ)
libDetail_EICGCCXCPT = Format(xYEICGCC0.EICGCCXCPT, "@@@@@ @@@ @@@ @@@@@@@")
libDetail_EICGCCXID = xYEICGCC0.EICGCCXID

If xYEICGCC0.EICGCCOPE = "RI0" Then fraDetail_Display_BDF


txtDetail_EICGCCXNOM = Trim(xYEICGCC0.EICGCCXNOM)
If Trim(xYEICGCC0.EICGCCXNOM) = "" Then
    txtDetail_EICGCCXNOM.BackColor = RGB(255, 196, 164)
Else
    txtDetail_EICGCCXNOM.BackColor = RGB(164, 255, 164)

End If
txtDetail_EICGCCXECO = Trim(xYEICGCC0.EICGCCXECO)
If Trim(xYEICGCC0.EICGCCXECO) = "" Then
    txtDetail_EICGCCXECO.BackColor = RGB(255, 196, 164)
Else
    txtDetail_EICGCCXECO.BackColor = RGB(164, 255, 164)

End If

If xYEICGCC0.EICGCCVAMJ <> 0 Then
    libDetail_EICGCCVAMJ = dateImp10(xYEICGCC0.EICGCCVAMJ)
    Call imgCHQ_Load(xYEICGCC0.EICGCCVAMJ, xYEICGCC0.EICGCCVJPG)
Else
    libDetail_EICGCCVAMJ = ""
End If

libDetail_EICGCCVREM = xYEICGCC0.EICGCCVREM
libDetail_EICGCCVJPG = xYEICGCC0.EICGCCVJPG
txtDetail_EICGCCVINT = Trim(xYEICGCC0.EICGCCVINT)
If Trim(xYEICGCC0.EICGCCVINT) = "" Then
    txtDetail_EICGCCVINT.BackColor = RGB(255, 196, 164)
    If xYEICGCC0.EICGCCVAMJ <> 0 Then
        Select Case xYEICGCC0.EICGCCOPE
            Case "RI0": txtDetail_EICGCCVINT = "EIC " & dateImp10(xYEICGCC0.EICGCCVAMJ)
            Case "REM": txtDetail_EICGCCVINT = "JC  " & dateImp10(xYEICGCC0.EICGCCAMJ)
        End Select
    End If
Else
    txtDetail_EICGCCVINT.BackColor = RGB(164, 255, 164)

End If

txtDetail_EICGCCVEXT = Trim(xYEICGCC0.EICGCCVEXT)
If Trim(xYEICGCC0.EICGCCVEXT) = "" Then
    txtDetail_EICGCCVEXT.BackColor = RGB(255, 196, 164)
Else
    txtDetail_EICGCCVEXT.BackColor = RGB(164, 255, 164)

End If

If xYEICGCC0.EICGCCEAMJ <> 0 Then
    wAMJMin = xYEICGCC0.EICGCCEAMJ
    Call DTPicker_Set(txtDetail_EICGCCEAMJ, wAMJMin)
Else
    wAMJMin = xYEICGCC0.EICGCCAMJ
    Call DTPicker_Set(txtDetail_EICGCCEAMJ, wAMJMin)
    txtDetail_EICGCCEAMJ.Value = Null
End If

Call cbo_Scan(xYEICGCC0.EICGCCKLAB, cboDetail_EICGCCKLAB)
If xYEICGCC0.EICGCCKLAB = " " Then
    cboDetail_EICGCCKLAB.BackColor = RGB(255, 196, 164)
Else
    cboDetail_EICGCCKLAB.BackColor = RGB(164, 255, 164)

End If

Call cbo_Scan(xYEICGCC0.EICGCCKSIG, cboDetail_EICGCCKSIG)
If xYEICGCC0.EICGCCKSIG = " " Then
    cboDetail_EICGCCKSIG.BackColor = RGB(255, 196, 164)
Else
    cboDetail_EICGCCKSIG.BackColor = RGB(164, 255, 164)

End If

Call cbo_Scan(xYEICGCC0.EICGCCKEND, cboDetail_EICGCCKEND)
If xYEICGCC0.EICGCCKEND = " " Then
    cboDetail_EICGCCKEND.BackColor = RGB(255, 196, 164)
Else
    cboDetail_EICGCCKEND.BackColor = RGB(164, 255, 164)

End If

Call cbo_Scan(xYEICGCC0.EICGCCKMT, cboDetail_EICGCCKMT)
If xYEICGCC0.EICGCCKMT = " " Then
    cboDetail_EICGCCKMT.BackColor = RGB(255, 196, 164)
Else
    cboDetail_EICGCCKMT.BackColor = RGB(164, 255, 164)

End If

Call cbo_Scan(xYEICGCC0.EICGCCSTAK, cboDetail_EICGCCSTAK)
If xYEICGCC0.EICGCCSTAK = " " Then
    cboDetail_EICGCCSTAK.BackColor = RGB(255, 128, 128)
Else
    cboDetail_EICGCCSTAK.BackColor = RGB(96, 255, 96)

End If
Call cbo_Scan(xYEICGCC0.EICGCCSTA, cboDetail_EICGCCSTA)
If xYEICGCC0.EICGCCSTA = " " Then
    cboDetail_EICGCCSTA.BackColor = RGB(255, 128, 128)
Else
    cboDetail_EICGCCSTA.BackColor = RGB(96, 255, 96)

End If


libCHQ_EICGCCEMT = libDetail_EICGCCEMT
libCHQ_EICGCCECPT = libDetail_EICGCCECPT
libCHQ_EICGCCECPT_X = libDetail_EICGCCECPT_X
libCHQ_EICGCCXNOM = txtDetail_EICGCCXNOM
'________________________________________________________________________

txtDetail_EICGCCVINT.Enabled = False
txtDetail_EICGCCVEXT.Enabled = False
txtDetail_EICGCCXNOM.Enabled = False
txtDetail_EICGCCXECO.Enabled = False
cboDetail_EICGCCKMT.Enabled = False
cboDetail_EICGCCKSIG.Enabled = False
cboDetail_EICGCCKEND.Enabled = False
cboDetail_EICGCCKLAB.Enabled = False
txtDetail_EICGCCEAMJ.Enabled = False
cboDetail_EICGCCSTA.Enabled = False
cboDetail_EICGCCSTAK.Enabled = False
cmdDetail_Update.Visible = False
cmdDetail_Action.Visible = False

'________________________________________________________________________
wEICGCCSTA = xYEICGCC0.EICGCCSTA

If wEICGCCSTA = "I" Then GoTo Exit_sub

'If YEICGCC0_Aut.Rapprocher Then
'    If wEICGCCSTA = "A" Then
'        cmdDetail_Action.Caption = "Reprise"
'        cmdDetail_Action.Visible = True
'        GoTo Exit_Sub
'    Else
'        cmdDetail_Action.Caption = "Annuler"
'        cmdDetail_Action.Visible = True
'        'wEICGCCSTA = " "
'    End If
'End If

If YEICGCC0_Aut.Saisir Or YEICGCC0_Aut.Valider Or YEICGCC0_Aut.Rapprocher Then
    cmdDetail_Action.Visible = True
    cmdDetail_Update.Visible = True
End If

If xYEICGCC0.EICGCCVJPG = 0 And xYEICGCC0.EICGCCSTA = " " Then
    fraDetail.Visible = True
    cmdDetail_UpdateVO.Visible = blnDetail_UpdateVO
    'cmdDetail_CLIBEN.Visible = Not cmdDetail_UpdateVO.Visible
    GoTo Exit_sub
End If


Select Case wEICGCCSTA
    Case "V":
            If YEICGCC0_Aut.Valider Then
                txtDetail_EICGCCVEXT.Enabled = True
            Else
                If YEICGCC0_Aut.Saisir And Trim(txtDetail_EICGCCVEXT) = "" Then
                    txtDetail_EICGCCVEXT.Enabled = True
                End If
            End If
    Case " ":
        If YEICGCC0_Aut.Saisir Then
            txtDetail_EICGCCVINT.Enabled = True
            txtDetail_EICGCCVEXT.Enabled = True
            If xYEICGCC0.EICGCCXBQ <> strSocBdfE Then txtDetail_EICGCCXNOM.Enabled = True
            txtDetail_EICGCCXECO.Enabled = True
        End If
        
        If YEICGCC0_Aut.Valider Then
            cboDetail_EICGCCKMT.Enabled = True
            cboDetail_EICGCCKSIG.Enabled = True
            cboDetail_EICGCCKEND.Enabled = True
            cboDetail_EICGCCKLAB.Enabled = True
            txtDetail_EICGCCEAMJ.Enabled = True
        End If

End Select
'If xYEICGCC0.EICGCCDOS > 0 Then cmdDetail_Update.Visible = True
       
'________________________________________________________________________
GoTo Exit_sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
Exit_sub:
cmdDetail_CLIBEN.Visible = Not cmdDetail_UpdateVO.Visible

blnControl = True
fraDetail.Visible = True

If cmdSelect_SQL_K = "J" Then
    cmdDetail_Action.Visible = False
    cmdDetail_Update.Visible = False
    cmdDetail_UpdateVO.Visible = False
    cboDetail_EICGCCKMT.Enabled = False
    cboDetail_EICGCCKSIG.Enabled = False
    cboDetail_EICGCCKEND.Enabled = False
    cboDetail_EICGCCKLAB.Enabled = False
    txtDetail_EICGCCEAMJ.Enabled = False
    txtDetail_EICGCCVINT.Enabled = False
    txtDetail_EICGCCVEXT.Enabled = False
    txtDetail_EICGCCXECO.Enabled = False

    xSQL = "select * from " & paramIBM_Library_SABJRN & ".JRNENT0 " _
         & " where jorcv = " & oldYEICGCC0.JORCV _
         & " and joSEQN = " & oldYEICGCC0.JOSEQN
         
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        V = srvJRNENT0_GetBuffer_ODBC(rsSab, xJRNENT0)
        If IsNull(V) Then
            xJRNENT0.JOUSER = oldYEICGCC0.EICGCCUUSR
            Call srvJRNENT0_fgX(xJRNENT0, fgJRNENT0)
            fraJRNENT0.Caption = JOENTT_Lib(xJRNENT0.JOENTT)
            fraJRNENT0.ForeColor = vbRed
            fraJRNENT0.Visible = True
        End If
    End If
Else

    fgLogV_Display

End If



End Sub
Public Sub cmdSendMail_EIC_GCC()
Dim wSendMail As typeSendMail
Dim xDétail As String, xHeader As String, mbgColor As String
Dim K As Long, htmlFontColor_K As String
Dim xAlerte As String, xSQL As String

On Error Resume Next

'____________________________________________________________________________________________
xSQL = "select count(*) as Tally    from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
     & " where EICGCCSTA = ' ' and  EICGCCVJPG = 0 and EICGCCAMJ < " & mAMJ_7Ante
Set rsSab = cnsab.Execute(xSQL)
K = rsSab("Tally")
If K > 0 Then
    xAlerte = htmlFontColor_Red & "<B><U>" & K & " vignettes non parvenues concernant des chèques circulants reçus via SIT avant le " & dateImp10(mAMJ_7Ante) & "<BR>(voir pièce jointe à la page : EIC en attente)</B></U><BR><BR>"
End If
'____________________________________________________________________________________________
cmdSelect_SQL_K = "L#"
Call DTPicker_Set(txtSelect_EICGCCAMJ, DSys)
Call DTPicker_Set(txtDetail_EICGCCEAMJ, mAMJ_7Ante)

'optSelect_EICGCCAMJ_S = True
cboEICGCCLOGK = ""

cmdSelect_SQL_YEICGCCLOG

xHeader = "<TR>" _
         & "<TD bgcolor=#0090A0 width=50 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Dossier</TD>" _
         & "<TD bgcolor=#0090A0 width=100 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Action</B></TD>" _
         & "<TD bgcolor=#0090A0 width=550 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Commentaire</TD>" _
         & "<TD bgcolor=#0090A0 width=300 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>mise à jour le</TD>" _
        & "</TR>"

xDétail = ""
mbgColor = "bgcolor = #FAFAD2"
For K = 1 To arrYEICGCCLOG_Nb
    Select Case arrYEICGCCLOG(K).EICGCCLOGA
        Case "A": htmlFontColor_K = htmlFontColor_Red
        Case "V": htmlFontColor_K = htmlFontColor_Green
        Case Else:
            If Trim(arrYEICGCCLOG(K).EICGCCLOGK) = "AI1" Then
                htmlFontColor_K = htmlFontColor_Red
            Else
                htmlFontColor_K = htmlFontColor_Blue
            End If
    End Select
    If arrYEICGCCLOG(K).EICGCCLOGI > 0 Then
        x = arrYEICGCCLOG(K).EICGCCLOGI
    Else
        x = ""
    End If
    xDétail = xDétail _
         & "<TR>" _
         & "<TD " & mbgColor & " width=50 height=5><span style='font-size:8.0pt;font-family:Arial'>" & htmlFontColor_K & x & "</TD>" _
         & "<TD " & mbgColor & " width=100 height=5><span style='font-size:8.0pt;font-family:Arial'>" & htmlFontColor_K & arrYEICGCCLOG(K).EICGCCLOGK & "</TD>" _
         & "<TD " & mbgColor & " width=550 height=5><span style='font-size:8.0pt;font-family:Arial'>" & htmlFontColor_K & arrYEICGCCLOG(K).EICGCCLOGX & "</TD>" _
         & "<TD " & mbgColor & " width=300 height=5><span style='font-size:8.0pt;font-family:Arial'>" & htmlFontColor_K & dateImp10(arrYEICGCCLOG(K).EICGCCLOGD) & "   " & timeImp8(arrYEICGCCLOG(K).EICGCCLOGH) & "   " & arrYEICGCCLOG(K).EICGCCLOGU & "</TD>" _
         & "</TR>"

Next K

wSendMail.FromDisplayName = "@EIC_GCC"
wSendMail.RecipientDisplayName = "EIC_GCC"

wSendMail.Subject = "Traitement EIC_GCC du : " & dateImp10(YBIATAB0_DATE_CPT_J) & " (cf. pièce jointe)"
wSendMail.Attachment = prtIMP_PDF_FileName
wSendMail.Message = "<body bgcolor = #FFFFFF><BR>" _
                    & xAlerte _
                    & "<TABLE   width=1000 border=1 cellpadding=4 ></B>" _
                    & xHeader _
                    & xDétail _
                    & "</TABLE>"

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub


Public Sub imgCHQ_Load(lDate As Long, lImage As Long)
Dim x As String
Dim xPath_ImgCHQ As String
    
x = paramCHQ_SCAN_Image_Archive & "\"
xPath_ImgCHQ = x & Trim(lDate) & "\Archive\" & Format(lImage, "00000000") & ".jpg"
If Dir(xPath_ImgCHQ) <> "" Then
    imgCHQ.Picture = LoadPicture(xPath_ImgCHQ)
    'cmdPrint.Enabled = True
    fraCHQ.Visible = True
Else
    Call MsgBox(xPath_ImgCHQ, vbExclamation, "imgCHQ_Load")
    imgCHQ.Picture = LoadPicture("")
End If
xPath_ImgCHQ = x & Trim(lDate) & "\Archive\ba" & Format(lImage, "00000000") & ".jpg"
If Dir(xPath_ImgCHQ) <> "" Then
    imgCHQ_Verso.Picture = LoadPicture(xPath_ImgCHQ)
Else
    Call MsgBox(xPath_ImgCHQ, vbExclamation, "imgCHQ_Load")
    imgCHQ_Verso.Picture = LoadPicture("")
End If
fraCHQ.Visible = True


End Sub

Public Sub cmdSelect_Reset()
If blnControl Then
    If Me.Enabled Then cmdContext.SetFocus
    lstErr.Clear
    fgSelect.Visible = False
    fraDetail.Visible = False
    fraCHQ.Visible = False
    fraCHQ_Max.Visible = False
    fraLogV.Visible = False
    fgLogV.Visible = False
    fraSuivi.Visible = False
    fgList.Visible = False
    'fraSelect_Options.Visible = False
    fraSelect_Options_1.Visible = False
    fraSelect_Options_L.Visible = False
    fraSelect_Options_St.Visible = False
    lstW.Visible = False
    lblSelect_EICGCCAMJ = "date de comptabilisation"
    cmdSelect_Ok.Visible = True
    cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, 3))
    Select Case cmdSelect_SQL_K
        Case "1":
            fraSelect_Options.Visible = True: fraSelect_Options_1.Visible = True
        Case "1s":
            lblSelect_EICGCCAMJ = "date de la numérisation"
            fraSelect_Options.Visible = True: fraSelect_Options_1.Visible = True
        Case "2e": cmdSelect_Ok.Visible = False: fraSelect_Options.Visible = False: cmdSelect_SQL_2e
        Case "2v": cmdSelect_Ok.Visible = False: fraSelect_Options.Visible = False: cmdSelect_SQL_2v
        Case "2?": cmdSelect_Ok.Visible = False: fraSelect_Options.Visible = False: cmdSelect_SQL_2vo
        Case "E": cmdSelect_Ok.Visible = False: fraSelect_Options.Visible = False: cmdSelect_SQL_Echeancier
        Case "Ie", "Iv":
            fraSelect_Options.Visible = False
        Case "I#": cmdSelect_Ok.Visible = False: fraSelect_Options.Visible = False
                    cboSuivi_K.ListIndex = 0
                    txtSuivi_Q = ""
                    fraSuivi.Visible = True
        Case "L#":
                  lblSelect_EICGCCAMJ = "date des événements":: fraSelect_Options_L.Visible = True
                  fraSelect_Options.Visible = True
                  'cboEICGCCLOGK.Visible = True
        Case "J": fraSelect_Options.Visible = False: cmdSelect_Ok.Visible = False: cmdSelect_SQL_Journalisation
        Case "J#": fraSelect_Options.Visible = False: cmdSelect_Ok.Visible = False: cmdSelect_SQL_Journalisation_Action
        Case "S?": fraSelect_Options.Visible = False: cmdSelect_Ok.Visible = False: cmdSelect_SQL_Stock
        Case "St": fraSelect_Options.Visible = False: fraSelect_Options_St.Width = 4800: fraSelect_Options_St.Visible = True: cmdSelect_Ok.Visible = True
        Case "SBq": fraSelect_Options.Visible = False: fraSelect_Options_St.Width = 8000: fraSelect_Options_St.Visible = True: cmdSelect_Ok.Visible = True
        Case "Xb": cmdSelect_SQL_Exportation_Bénéficiaires
    End Select

End If

End Sub



'______________________________________________________________________

Private Sub fgSelect_Display_YEICGCCLOG()
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler

ReDim selYEICGCCLOG(arrYEICGCCLOG_Nb)
selYEICGCCLOG_Nb = arrYEICGCCLOG_Nb

SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "Date               |<Heure             |<Utilisateur          |>Seq     |" _
                      & "> Dossier |<événement         |>Echéance         |<Libellé                                                                                                                                  |"
fgSelect.Row = 0

currentAction = "fgSelect_Display_YEICGCCLOG"
    
For I = 1 To arrYEICGCCLOG_Nb
         
    xYEICGCCLOG = arrYEICGCCLOG(I)
    selYEICGCCLOG(I) = xYEICGCCLOG
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    
     Select Case xYEICGCCLOG.EICGCCLOGA
        Case Is = "V": wColor = RGB(32, 96, 32)
        Case Is = "A", "I": wColor = RGB(128, 128, 128)
        Case Else:
            If xYEICGCCLOG.EICGCCLOGE > 0 Then
                wColor = vbMagenta
            Else
                wColor = RGB(64, 64, 128)
            End If
    End Select

    
    If cmdSelect_SQL_K <> "J#" Then
        fgSelect.Col = 0: fgSelect.Text = dateImp10(xYEICGCCLOG.EICGCCLOGD)
        fgSelect.CellForeColor = wColor
        fgSelect.Col = 1: fgSelect.Text = timeImp8(xYEICGCCLOG.EICGCCLOGH)
        fgSelect.CellForeColor = wColor
    Else
        fgSelect.Col = 0: fgSelect.Text = dateJma6_Imp10(xYEICGCCLOG.JODATE)
        fgSelect.CellForeColor = vbBlack
        fgSelect.Col = 1: fgSelect.Text = xYEICGCCLOG.JOENTT & " " & xYEICGCCLOG.JOSEQN
        fgSelect.CellForeColor = vbBlack
    End If

    fgSelect.Col = 2: fgSelect.Text = xYEICGCCLOG.EICGCCLOGU: fgSelect.CellForeColor = wColor
    fgSelect.Col = 3: fgSelect.Text = xYEICGCCLOG.EICGCCLOGS & " ": fgSelect.CellForeColor = wColor
    fgSelect.Col = 4: fgSelect.Text = xYEICGCCLOG.EICGCCLOGI & " ": fgSelect.CellForeColor = wColor
    fgSelect.Col = 5: fgSelect.Text = xYEICGCCLOG.EICGCCLOGK & " " & xYEICGCCLOG.EICGCCLOGA: fgSelect.CellForeColor = wColor
    If xYEICGCCLOG.EICGCCLOGE <> 0 Then fgSelect.Col = 6: fgSelect.Text = dateImp10(xYEICGCCLOG.EICGCCLOGE) & "  "
    fgSelect.CellForeColor = wColor
    fgSelect.Col = 7: fgSelect.Text = xYEICGCCLOG.EICGCCLOGX: fgSelect.CellForeColor = wColor

    fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = I: fgSelect.CellForeColor = wColor

Next I

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYEICGCCLOG_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub



Private Sub cmdSelect_SQL_1()
Dim V
Dim x As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdYEICGCC0_SQL"
blnOk = False
xWhere = ""



x = Trim(txtSelect_EICGCCID)
If x <> "" Then
    xWhere = xWhere & " and   EICGCCID = " & Val(x)
Else
    x = Trim(txtSelect_EICGCCECHQ)
    If x <> "" Then
        xWhere = xWhere & " and   EICGCCECHQ ='" & Format(Val(x), "0000000") & "'"
    Else
        xWhere = cmdSelect_SQL_AMJ
        If cmdSelect_SQL_K = "1s" Then xWhere = Replace(xWhere, "EICGCCAMJ", "EICGCCVAMJ")
        x = Trim(txtSelect_EICGCCECLI)
        If x <> "" Then
            xWhere = xWhere & " and   EICGCCECLI ='" & Format(Val(x), "0000000") & "'"
        End If
        
        x = Trim(txtSelect_EICGCCECPT)
        If x <> "" Then
            xWhere = xWhere & " and   EICGCCECPT like'" & x & "%'"
        End If
        
        
        x = Trim(txtSelect_EICGCCXNOM)
        If x <> "" Then
            xWhere = xWhere & " and   EICGCCXNOM like'%" & x & "%'"
        End If
        
        x = Mid$(Trim(cboSelect_EICGCCSTA), 1, 1)
        If x <> "" Then
            xWhere = xWhere & " and   EICGCCSTA  = '" & x & "'"
        End If
        
        x = Trim(txtSelect_EICGCCXNOM)
        If x <> "" Then
            xWhere = xWhere & " and   EICGCCXNOM like'%" & x & "%'"
        End If
        
        x = Mid$(Trim(cboSelect_EICGCCEIND), 1, 1)
        If x <> "" Then
            xWhere = xWhere & " and   EICGCCEIND  = '" & x & "'"
        End If
    End If
End If


If xWhere = "" Then
    Call MsgBox("Préciser au moins un critère de filtrage", vbExclamation, "EICGCC : recherche")
    Exit Sub
End If
Mid$(xWhere, 1, 7) = " where "
arrYEICGCC0_SQL xWhere & " order by EICGCCAMJ , EICGCCID"

fgSelect_Display

If arrYEICGCC0_Nb = 1 Then
    oldYEICGCC0 = arrYEICGCC0(arrYEICGCC0_Nb)
    xYEICGCC0 = oldYEICGCC0
    fraDetail_Display

End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_2e()
Dim V
Dim x As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2e"
blnOk = False
xWhere = ""


xWhere = " where EICGCCSTA = ' '  and EICGCCVJPG = 0"
arrYEICGCC0_SQL xWhere & " order by EICGCCAMJ , EICGCCID"

fgSelect_Display


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_2v()
Dim V
Dim x As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2v"
blnOk = False
xWhere = ""


xWhere = " where EICGCCSTA = ' ' and  EICGCCVJPG > 0 and EICGCCDOS > 0"
arrYEICGCC0_SQL xWhere & " order by EICGCCAMJ , EICGCCID"

fgSelect_Display


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_2vo()
Dim V
Dim x As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2vo"
blnOk = False
xWhere = ""


xWhere = " where EICGCCSTA = ' ' and  EICGCCVJPG > 0 and EICGCCDOS = 0"
arrYEICGCC0_SQL xWhere & " order by EICGCCAMJ , EICGCCID"

fgSelect_Display


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrYEICGCC0_SQL(xWhere As String)
Dim V
Dim x As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYEICGCC0(101)
arrYEICGCC0_Max = 100: arrYEICGCC0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYEICGCC0_GetBuffer(rsSab, xYEICGCC0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYEICGCC0.fgselect_Display"
        '' Exit Sub
     Else
         arrYEICGCC0_Nb = arrYEICGCC0_Nb + 1
         If arrYEICGCC0_Nb > arrYEICGCC0_Max Then
             arrYEICGCC0_Max = arrYEICGCC0_Max + 100
             ReDim Preserve arrYEICGCC0(arrYEICGCC0_Max)
         End If
         
         arrYEICGCC0(arrYEICGCC0_Nb) = xYEICGCC0
    End If
    rsSab.MoveNext
Loop


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

'______________________________________________________________________

Private Sub fgSelect_Display_Echeancier()
Dim wColor As Long, x As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler

SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
x = fgSelect_FormatString
x = Replace(x, "chèque", "échéance    .")
x = Replace(x, "Bénéficiaire                                 ", "événement")
x = Replace(x, "Banque/compte", "en attente                                           ")

fgSelect.FormatString = x
fgSelect.Row = 0

currentAction = "fgSelect_Display"

For I = 1 To arrYEICGCC0_Nb
    xYEICGCC0 = arrYEICGCC0(I)
    
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine I, False
    
Next I
    

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYEICGCC0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
'================================================================




End Sub

Private Sub fgLogV_Display()
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
fgLogV.Visible = False
fgLogV_Reset

fgLogV.Rows = 1
fgLogV.FormatString = fgLogV_FormatString
fgLogV.Row = 0

currentAction = "fgLogV_Display"
Call arrYEICGCCLOG_SQL(".YEICGCCLOV where EICGCCLOGI = " & oldYEICGCC0.EICGCCID _
                       & " order by EICGCCLOGD, EICGCCLOGH, EICGCCLOGU, EICGCCLOGS")
    
For I = 1 To arrYEICGCCLOG_Nb
         
    xYEICGCCLOG = arrYEICGCCLOG(I)
    fgLogV.Rows = fgLogV.Rows + 1
    fgLogV.Row = fgLogV.Rows - 1
    fgLogV_DisplayLine I
    
Next I

fgLogV.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long, blnYEICGCC0 As Boolean)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim xSQL As String
On Error Resume Next
 Select Case xYEICGCC0.EICGCCSTA
    Case Is = "V", "@": wColor = RGB(32, 96, 32)
    Case Is = "A", "I", "R": wColor = RGB(128, 128, 128)
    Case Else: wColor = RGB(64, 64, 128)
        If xYEICGCC0.EICGCCDOS = 0 Then
            wColor = vbRed
        Else
            If xYEICGCC0.EICGCCVJPG = 0 Then
                If xYEICGCC0.EICGCCAMJ < mAMJ_7Ante Then
                    wColor = vbRed
                Else
                    wColor = vbMagenta 'RGB(255, 96, 255)
                End If
            Else
                If xYEICGCC0.EICGCCSTAK = " " Then
                    Select Case xYEICGCC0.EICGCCSTAK
                        Case "X": XPrt.ForeColor = vbMagenta
                        Case Else: XPrt.ForeColor = vbBlue
                    End Select
                End If
            End If
        End If
End Select

If cmdSelect_SQL_K <> "J" Then
    fgSelect.Col = 0: fgSelect.Text = dateImp10(xYEICGCC0.EICGCCAMJ)
    fgSelect.CellForeColor = wColor
    fgSelect.Col = 1: fgSelect.Text = xYEICGCC0.EICGCCOPE
    fgSelect.CellForeColor = wColor
    fgSelect.Col = 2: fgSelect.Text = xYEICGCC0.EICGCCDOS
    fgSelect.CellForeColor = wColor
Else
    fgSelect.Col = 0: fgSelect.Text = dateJma6_Imp10(xYEICGCC0.JODATE)
    fgSelect.CellForeColor = vbBlack
    fgSelect.Col = 1: fgSelect.Text = xYEICGCC0.JOENTT
    fgSelect.CellForeColor = vbBlack
    fgSelect.Col = 2: fgSelect.Text = xYEICGCC0.JOSEQN
    fgSelect.CellForeColor = wColor
End If


fgSelect.Col = 3
xSQL = "select COMPTEINT from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " Where COMPTECOM = '" & xYEICGCC0.EICGCCECPT & "'"
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    fgSelect.Text = rsSab("COMPTEINT")
Else
    fgSelect.Text = xYEICGCC0.EICGCCECLI
End If
fgSelect.CellForeColor = wColor
fgSelect.Col = 4: fgSelect.Text = Format(xYEICGCC0.EICGCCEMT, "### ### ### ###.00") & " "
fgSelect.CellForeColor = wColor
If blnYEICGCC0 Then
    fgSelect.Col = 5: fgSelect.Text = xYEICGCC0.EICGCCECHQ
    fgSelect.CellForeColor = wColor
    fgSelect.Col = 6: fgSelect.Text = xYEICGCC0.EICGCCXNOM
    fgSelect.CellForeColor = wColor
    If xYEICGCC0.EICGCCXBQ = strSocBdfE Then
        fgSelect.Col = 7
        If IsNumeric(xYEICGCC0.EICGCCXCPT) Then
            fgSelect.Text = Format(xYEICGCC0.EICGCCXCPT, "@@@@@ @@@ @@@ @@@@@@@")
        Else
            fgSelect.Text = xYEICGCC0.EICGCCXCPT
        End If
        
    Else
        fgSelect.Col = 7: fgSelect.Text = xYEICGCC0.EICGCCXBQ
    End If
Else
    fgSelect.Col = 5: fgSelect.Text = dateImp10(arrYEICGCCLOG(lIndex).EICGCCLOGE) & "  "
    fgSelect.CellForeColor = wColor
    fgSelect.CellBackColor = RGB(255, 228, 96)
    fgSelect.CellFontBold = True
    fgSelect.Col = 6: fgSelect.Text = arrYEICGCCLOG(lIndex).EICGCCLOGK
    fgSelect.CellForeColor = wColor
    fgSelect.Col = 7: fgSelect.Text = arrYEICGCCLOG(lIndex).EICGCCLOGX
End If
fgSelect.CellForeColor = wColor

fgSelect.Col = 8: fgSelect.Text = xYEICGCC0.EICGCCSTA & xYEICGCC0.EICGCCID
fgSelect.CellForeColor = wColor


fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
End Sub
Public Sub fgLogV_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
On Error Resume Next
 Select Case xYEICGCCLOG.EICGCCLOGA
    Case Is = "V": wColor = RGB(32, 96, 32)
    Case Is = "A", "I": wColor = RGB(128, 128, 128)
    Case Else:
        If xYEICGCCLOG.EICGCCLOGE > 0 Then
            wColor = vbMagenta
        Else
            wColor = RGB(64, 64, 128)
        End If
End Select


fgLogV.Col = 0: fgLogV.Text = xYEICGCCLOG.EICGCCLOGK
fgLogV.CellForeColor = wColor
fgLogV.Col = 2: fgLogV.Text = xYEICGCCLOG.EICGCCLOGX
fgLogV.CellForeColor = wColor
If xYEICGCCLOG.EICGCCLOGE > 0 Then
    fgLogV.Col = 1: fgLogV.Text = " " & dateImp10(xYEICGCCLOG.EICGCCLOGE)
    fgLogV.CellForeColor = wColor
End If
fgLogV.Col = 4: fgLogV.Text = dateImp10(xYEICGCCLOG.EICGCCLOGD) & " " & timeImp8(xYEICGCCLOG.EICGCCLOGH) & " " & xYEICGCCLOG.EICGCCLOGU & " " & xYEICGCCLOG.EICGCCLOGS
fgLogV.CellForeColor = wColor
fgLogV.Col = 3: fgLogV.Text = xYEICGCCLOG.EICGCCLOGA
fgLogV.CellForeColor = wColor


fgLogV.Col = fgLogV_arrIndex: fgLogV.Text = lIndex
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

Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, x As String, wIndex As Long

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    wIndex = Val(fgSelect.Text)
    Select Case lK
        Case 4: x = Format$(arrYEICGCC0(wIndex).EICGCCEMT, "000000000000000.00")
        Case 0: x = arrYEICGCC0(wIndex).EICGCCAMJ
        Case 8: x = Format$(arrYEICGCC0(wIndex).EICGCCID, "00000000000")
        Case 5:
            If cmdSelect_SQL_K = "E" Then
                x = arrYEICGCCLOG(wIndex).EICGCCLOGE
            Else
                x = arrYEICGCC0(wIndex).EICGCCECHQ
            End If
    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = x
Next I

fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub



'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim wFct As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

wFct = UCase$(Trim(Mid$(Msg, 1, 12)))
Call BiaPgmAut_Init(wFct, YEICGCC0_Aut)
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'MsgBox "JPL : lecture PROD / màj EICGCC test  ", vbCritical
'paramIBM_Library_SABSPE_XXX = "SAB073USPE"
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'blnSetfocus = True
Form_Init


Select Case wFct
    Case "@EIC_GCC": blnAuto = True
                    If Not IsEmpty(XPrt) Then Set Xprt_Previous = XPrt
                    Printer_PDF

                    cmdSelect_SQL_Import_EIC
                    cmdSelect_SQL_Import_REM
                    cmdSelect_SQL_Import_Vignettes
                    cmdPrint_Auto
                    cmdSendMail_EIC_GCC
                    Unload Me
                    If Not IsEmpty(Xprt_Previous) Then Set XPrt = Xprt_Previous
    Case Else: blnAuto = False
End Select

End Sub


Public Sub Form_Init()
Dim V, xSQL As String, x As String
Dim K As Long

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True

blnControl = False

cmdReset

lblDetail_EICGCCID.ForeColor = vbMagenta
libDetail_EICGCCECLI_Resp.ForeColor = vbMagenta
fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False

fraDetail.Visible = False
Set fraDetail.Container = fraTab0
fraDetail.Top = fgSelect.Top
fraDetail.Left = fgSelect.Left + fgSelect.Width - fraDetail.Width - 300

fraSelect_Options_1.BorderStyle = 0

Set fraSelect_Options_L.Container = fraSelect_Options
fraSelect_Options_L.Top = fraSelect_Options_1.Top
fraSelect_Options_L.Left = fraSelect_Options_1.Left
fraSelect_Options_L.Height = fraSelect_Options_1.Height
fraSelect_Options_L.Width = fraSelect_Options_1.Width
fraSelect_Options_L.BorderStyle = 0

Set fraSelect_Options_St.Container = fraSelect_Options.Container
fraSelect_Options_St.Top = fraSelect_Options.Top
fraSelect_Options_St.Left = fraSelect_Options.Left
'fraSelect_Options_St.Height = fraSelect_Options.Height
'fraSelect_Options_St.Width = fraSelect_Options.Width
'fraSelect_Options_St.BorderStyle = 0
Call DTPicker_Set(txtSelect_Options_St_AMJMAX, YBIATAB0_DATE_CPT_JS1) '
Call DTPicker_Set(txtSelect_Options_St_AMJMIN, Mid$(YBIATAB0_DATE_CPT_JS1, 1, 4) & "0101") '




lstW.Visible = False
Set lstW.Container = fraDetail
lstW.Top = 5160
lstW.Left = 1440
lstW.ForeColor = vbBlue

fraCHQ.Visible = False
Set fraCHQ.Container = fraTab0
fraCHQ.Top = fgSelect.Top
fraCHQ.Left = fraDetail.Left - fraCHQ.Width
fraCHQ.ForeColor = vbMagenta

fraCHQ_Max.Visible = False
Set fraCHQ_Max.Container = fraTab0
fraCHQ_Max.Top = fgSelect.Top
fraCHQ_Max.Left = fraDetail.Left + fraDetail.Width - fraCHQ_Max.Width


fgLogV.Visible = False
fgLogV_FormatString = fgLogV.FormatString
Set fgLogV.Container = fraTab0
fgLogV.Top = fraCHQ.Top + fraCHQ.Height
fgLogV.Left = fraDetail.Left - fgLogV.Width
fgLogV.Height = 2100

fraLogV.Visible = False
Set fraLogV.Container = fraTab0
fraLogV.Top = fraCHQ.Top + fraCHQ.Height
fraLogV.Left = fraDetail.Left - fraLogV.Width
fraLogV.Height = 2100


fraJRNENT0.Visible = False
Set fraJRNENT0.Container = fraTab0
fraJRNENT0.Top = fgSelect.Top
fraJRNENT0.Left = fraDetail.Left - fraJRNENT0.Width

fraYEICGCCLOG.Visible = False
fraYEICGCCLOG.Top = fgSelect.Top
fraYEICGCCLOG.Left = fraDetail.Left

fgList.Visible = False
fgList_FormatString = fgList.FormatString
Set fgList.Container = fraTab0
fgList.Top = fraCHQ.Top
fgList.Left = fraDetail.Left - fgList.Width
fgList.Height = fraCHQ.Height

Set cmdDetail_CLIBEN.Container = cmdDetail_UpdateVO.Container
cmdDetail_CLIBEN.Top = cmdDetail_UpdateVO.Top
cmdDetail_CLIBEN.Left = cmdDetail_UpdateVO.Left
cmdDetail_CLIBEN.Visible = False

mEICGCCID = 0
Call DTPicker_Set(txtSelect_EICGCCAMJ, YBIATAB0_DATE_CPT_JS1) '
Call DTPicker_Set(txtSelect_EICGCCAMJ_Min, YBIATAB0_DATE_CPT_JP0) '

txtSelect_EICGCCAMJ_Min.Value = Null

cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1  - sélection (filtre)"
cboSelect_SQL.AddItem "1s - sélection (date scan)"
cboSelect_SQL.AddItem "2v - vignettes à contrôler"
cboSelect_SQL.AddItem "2e - EIC | REM en attente de vignettes"
cboSelect_SQL.AddItem "2? - vignettes orphelines "
cboSelect_SQL.AddItem "E  - échéancier"
cboSelect_SQL.AddItem "L#  - Consultation du suivi des événements"
cboSelect_SQL.AddItem "S?  - stock de dossiers"
cboSelect_SQL.AddItem "St  - Statistiques"
cboSelect_SQL.AddItem "SBq  - Statistiques chèques de banque"
If YEICGCC0_Aut.Valider Then
    cboSelect_SQL.AddItem "J  - Consultation de la journalisation"
    cboSelect_SQL.AddItem "J#  - Consultation de la journalisation des événements"
    cboSelect_SQL.AddItem "I# - importation événement AF0 AI1 RQ0"
    cboSelect_SQL.AddItem "Xb  - Exportation des bénéficiaires"
End If
If YEICGCC0_Aut.Xspécial Then
    cboSelect_SQL.AddItem "Ie - importation EIC"
    cboSelect_SQL.AddItem "Ir - importation REM (SAB)"
    cboSelect_SQL.AddItem "Iv - importation vignettes"
End If
cboSelect_SQL.ListIndex = 0

srvCHQ_SCAN_param

cboDetail_EICGCCKLAB.Clear
cboDetail_EICGCCKLAB.AddItem "  - à vérifier"
cboDetail_EICGCCKLAB.AddItem "X - non conforme"
cboDetail_EICGCCKLAB.AddItem "I - ignoré"
cboDetail_EICGCCKLAB.AddItem "V - vérifié"

cboDetail_EICGCCKMT.Clear
cboDetail_EICGCCKMT.AddItem "  - à vérifier"
cboDetail_EICGCCKMT.AddItem "V - vérifié"
cboDetail_EICGCCKMT.AddItem "X - non conforme"

cboDetail_EICGCCKEND.Clear
cboDetail_EICGCCKEND.AddItem "  - à vérifier"
cboDetail_EICGCCKEND.AddItem "V - vérifié"
cboDetail_EICGCCKEND.AddItem "X - non conforme"

cboDetail_EICGCCKSIG.Clear
cboDetail_EICGCCKSIG.AddItem "  - à vérifier"
cboDetail_EICGCCKSIG.AddItem "V - vérifié"
cboDetail_EICGCCKSIG.AddItem "X - non conforme"

cboDetail_EICGCCSTAK.Clear
cboDetail_EICGCCSTAK.AddItem "  - CHQ à vérifier"
cboDetail_EICGCCSTAK.AddItem "V - CHQ conforme"
cboDetail_EICGCCSTAK.AddItem "X - CHQ non conforme"
cboDetail_EICGCCSTAK.AddItem "! - CHQ accepté"

cboDetail_EICGCCSTA.Clear
cboDetail_EICGCCSTA.AddItem "  - EIC en cours"
cboDetail_EICGCCSTA.AddItem "A - EIC annulée"
cboDetail_EICGCCSTA.AddItem "I - ignoré"
cboDetail_EICGCCSTA.AddItem "R - rejeté"
cboDetail_EICGCCSTA.AddItem "V - EIC vérifiée"
cboDetail_EICGCCSTA.AddItem "@ - val automatique"


cboSelect_EICGCCSTA.Clear
cboSelect_EICGCCSTA.AddItem "  - Tous les dossiers"
cboSelect_EICGCCSTA.AddItem "A - EIC annulés"
cboSelect_EICGCCSTA.AddItem "I - ignorés"
cboSelect_EICGCCSTA.AddItem "R - rejetés"
cboSelect_EICGCCSTA.AddItem "V - EIC vérifiés"
cboSelect_EICGCCSTA.AddItem "@ - val automatique"


xSQL = "select count(*) as Tally    from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
     & " where EICGCCSTA = ' ' and  EICGCCVJPG > 0 and EICGCCDOS = 0"
Set rsSab = cnsab.Execute(xSQL)
If rsSab("Tally") = 0 Then
    blnDetail_UpdateVO = False
Else
    blnDetail_UpdateVO = True
End If

lstW.Clear
cboSelect_EICGCCEIND.Clear
cboSelect_EICGCCEIND.AddItem " "
cboSelect_EICGCCEIND.AddItem "0"
cboSelect_EICGCCEIND.AddItem "1"
cboSelect_EICGCCEIND.AddItem "2"
cboSelect_EICGCCEIND.AddItem "3"
cboSelect_EICGCCEIND.AddItem "4"
cboSelect_EICGCCEIND.AddItem "5"
cboSelect_EICGCCEIND.AddItem "6"
cboSelect_EICGCCEIND.AddItem "7"
cboSelect_EICGCCEIND.AddItem "8"
cboSelect_EICGCCEIND.AddItem "9"
 
'Initialisation Log________________________________________________________________________________
rsYEICGCCLOG_Init newYEICGCCLOG

cboEICGCCLOGK.Clear
cboEICGCCLOGK.AddItem ""
xSQL = "select distinct EICGCCLOGK from " & paramIBM_Library_SABSPE_XXX & ".YEICGCCLOG order by EICGCCLOGK"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    cboEICGCCLOGK.AddItem Trim(rsSab("EICGCCLOGK"))
    rsSab.MoveNext
Loop

'Initialisation cbo _______________________________________________________________________________

Call DTPicker_Set(txtLogV_E, YBIATAB0_DATE_CPT_JS1) '
mAMJ_7Past = dateElp("Jour", 7, YBIATAB0_DATE_CPT_JS1)
mAMJ_7Ante = dateElp("Jour", -7, YBIATAB0_DATE_CPT_JS1)

cmdParam_Quit_Click
lstParam_Action.Clear
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YEICGCC0' and BIATABK1 = 'Action' order by BIATABK2"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF

    lstParam_Action.AddItem rsSab("BIATABK2")
    rsSab.MoveNext
Loop
fraParam_K.Enabled = False

If lstParam_Action.ListCount = 0 Then parametrage_Reprise
'Initialisation cbo _______________________________________________________________________________

Set fraSuivi.Container = fraTab0
fraSuivi.Top = 1000
fraSuivi.Left = 7000
fraSuivi.Visible = False
cboSuivi_K.Clear
cboSuivi_K.AddItem "AF0"
cboSuivi_K.AddItem "AI1"
cboSuivi_K.AddItem "RQ0"

cboDetail_EICGCCXECO.Visible = False
cboDetail_EICGCCXECO.Top = txtDetail_EICGCCXECO.Top
cboDetail_EICGCCXECO.Left = txtDetail_EICGCCXECO.Left

cboDetail_EICGCCXECO.Clear

xSQL = "select count(*) as tally from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YEICGCC0' and BIATABK1 = 'Motif Eco'"
Set rsSab = cnsab.Execute(xSQL)
K = 1
If Not rsSab.EOF Then K = rsSab("Tally") + 1
ReDim arrEICGCCXECO(K)
'cboDetail_EICGCCXECO.Sorted = False

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YEICGCC0' and BIATABK1 = 'Motif Eco' order by BIATABK2"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    'V = rsYBIATAB0_GetBuffer(rsSab, xPays_importation)

    cboDetail_EICGCCXECO.AddItem rsSab("BIATABK2")
    arrEICGCCXECO(cboDetail_EICGCCXECO.ListCount - 1) = rsSab("BIATABTXT")
    rsSab.MoveNext
Loop

Set rsSab = Nothing
'____________________________________________________________________
mCLIENARES = "": mMNURUTUTI = "": mMNUUTIMAI = ""
mCLIENARES_Lib = ""

If paramIBM_Library_SAB = "SAB073U" Then
    paramIBM_Library_SABJRN = "SAB073JRN"
    MsgBox "Form_Init : paramIBM_Library_SABJRN=SAB073JRN"
End If
'========================================================================
'========================================================================


Me.Enabled = True
End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------

blnControl = False
blnError = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
currentAction = ""
SSTab1.Tab = 0
'fraDetail.ForeColor = vbBlue
'fraDetail_X.ForeColor = vbBlue
'fraDetail_V.ForeColor = vbBlue
'fraDetail_K.ForeColor = vbBlue
blnControl = True

End Sub



Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgSelect.Visible = False
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = fgSelect_arrIndex To fgSelect.FixedCols Step -1
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = fgSelect_arrIndex To fgSelect.FixedCols Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
    End If
End If
fgSelect.LeftCol = fgSelect.FixedCols
fgSelect.Visible = True
End Sub

Public Sub fgLogV_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgLogV.Visible = False
mRow = fgLogV.Row

If lRow > 0 And lRow < fgLogV.Rows Then
    fgLogV.Row = lRow
    For I = fgLogV_arrIndex To fgLogV.FixedCols Step -1
        fgLogV.Col = I: fgLogV.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgLogV.Row = mRow
    If fgLogV.Row > 0 Then
        lRow = fgLogV.Row
        lColor_Old = fgLogV.CellBackColor
        For I = fgLogV_arrIndex To fgLogV.FixedCols Step -1
          fgLogV.Col = I: fgLogV.CellBackColor = lColor
        Next I
    End If
End If
fgLogV.LeftCol = fgLogV.FixedCols
fgLogV.Visible = True
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

Private Sub cboDetail_EICGCCXECO_Click()
If cboDetail_EICGCCXECO.Visible Then
    cboDetail_EICGCCXECO.Visible = False
    txtDetail_EICGCCXECO = arrEICGCCXECO(cboDetail_EICGCCXECO.ListIndex)
    txtDetail_EICGCCXECO.Visible = True
    'txtDetail_EICGCCXECO.SetFocus
End If
End Sub


Private Sub cboEICGCCLOGK_GotFocus()
txt_GotFocus cboEICGCCLOGK
If fgSelect.Visible Then cmdSelect_Reset

End Sub

Private Sub cboLogV_K_Click()
Dim xSQL As String
Dim wAction As String

Me.Enabled = False: Me.MousePointer = vbHourglass: blnControl = False

wAction = Trim(cboLogV_K)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YEICGCC0' and BIATABK1 = 'Action'  and BIATABK2 = '" & wAction & "'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    lblLogV_X = Trim(rsSab("BIATABTXT"))
Else
    lblLogV_X = ""
End If

cboLogV_K2_Load wAction

Select Case wAction
    Case "Annulation", "Reprise/Ann", "Révision", "CHQ accepté", "CHQ àIgnorer", "non circulan"
            txtLogV_E.Visible = False
    Case "Mail DCOM": txtLogV_E.CheckBox = True: Call DTPicker_Set(txtLogV_E, mAMJ_7Past) '
    Case Else: txtLogV_E.Visible = True: Call DTPicker_Set(txtLogV_E, YBIATAB0_DATE_CPT_JS1) '
End Select


Me.Enabled = True: Me.MousePointer = 0: blnControl = True
End Sub

Private Sub cboLogV_K2_Click()
Dim xSQL As String

Me.Enabled = False: Me.MousePointer = vbHourglass
If blnControl Then
    xSQL = "select BIATABTXT from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
         & " where BIATABID = 'YEICGCC0' and BIATABK1 = '" & Trim(cboLogV_K) & "'  and BIATABK2 = '" & Trim(cboLogV_K2) & "'"
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        lblLogV_X = rsSab("BIATABTXT")
        txtLogV_X = Trim(lblLogV_X) & " : "
    Else
        lblLogV_X = ""
    End If
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cboSelect_EICGCCSTA_GotFocus()
txt_GotFocus cboSelect_EICGCCSTA
If fgSelect.Visible Then cmdSelect_Reset

End Sub


Private Sub cboSelect_EICGCCSTA_LostFocus()
'txt_losttFocus cboSelect_EICGCCSTA

End Sub


Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

End Sub


Private Sub cmdCHQ_MAX_Quit_Click()
fraCHQ_Max.Visible = False
End Sub

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdDetail_Action_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cboLogV_K.Clear

If oldYEICGCC0.EICGCCDOS = 0 Then
    Select Case oldYEICGCC0.EICGCCSTA
        Case " ": cboLogV_K.AddItem "Annulation"
        Case "A": cboLogV_K.AddItem "Reprise/Ann"
    End Select

Else
    cboLogV_K.AddItem "AOCT"
    cboLogV_K.AddItem "Mail DCOM"
    If YEICGCC0_Aut.Rapprocher Then
        If oldYEICGCC0.EICGCCSTA <> "A" Then cboLogV_K.AddItem "Révision"
        
        Select Case oldYEICGCC0.EICGCCSTA
            Case " ":
                cboLogV_K.AddItem "Annulation"
                cboLogV_K.AddItem "CHQ rejeté"
                If oldYEICGCC0.EICGCCSTAK <> "V" Then cboLogV_K.AddItem "CHQ accepté"
            Case "A":
                cboLogV_K.AddItem "Reprise/Ann"
        End Select
    End If
End If
If oldYEICGCC0.EICGCCSTA = " " Then
    cboLogV_K.AddItem "non circulan"
    cboLogV_K.AddItem "CHQ àIgnorer"
End If

cboLogV_K.ListIndex = -1
cboLogV_K2.Clear
txtLogV_X = ""
txtLogV_E.Value = Null
fraLogV.Visible = True

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDetail_CLIBEN_Click()

If Trim(txtDetail_EICGCCXNOM) = "" Then
    Call MsgBox("Préciser le nom du bénéficiaire", vbInformation, "EIC_GCC : recherche Client +bénéficiaire")
Else
    Me.Enabled = False: Me.MousePointer = vbHourglass
    Call lstErr_Clear(lstErr, cmdContext, "> Recherche CLIENT  MÊME BENEFICIAIRE ........"): DoEvents
    
    Call fgList_Display(libDetail_EICGCCECLI, txtDetail_EICGCCXNOM)
    Me.Enabled = True: Me.MousePointer = 0
End If
End Sub

Private Sub cmdDetail_Quit_Click()
If fraLogV.Visible Then
    fraLogV.Visible = False
Else
    fraCHQ_Max.Visible = False
    fraCHQ.Visible = False
    fraJRNENT0.Visible = False
    fraDetail.Visible = False
    fgLogV.Visible = False: fraLogV.Visible = False
    fgList.Visible = False
End If
End Sub

Private Sub cmdDetail_Update_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

Call cmdDetail_Update_Ok("+Log")

Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Function cmdYEICGCC0_Update(lFct As String)
Dim K As Integer
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

'________________________________________________________________________________

If lFct = "Update" Then V = sqlYEICGCC0_Update(newYEICGCC0, oldYEICGCC0, True)
'________________________________________________________________________________
If lFct = "New" Then V = sqlYEICGCC0_Insert(newYEICGCC0)

If lFct = "Import" Then
    For K = 1 To arrYEICGCC0_Nb
        If arrYEICGCC0(K).EICGCCID > 0 Then
            oldYEICGCC0 = arrYEICGCC0(K)
            V = sqlYEICGCC0_Insert(oldYEICGCC0)
            If Not IsNull(V) Then GoTo Error_MsgBox
                newYEICGCCLOG.EICGCCLOGK = "AF0"

        End If
    Next K
End If
'________________________________________________________________________________
If lFct = "Delete" Then V = sqlYEICGCC0_Delete(oldYEICGCC0)
'________________________________________________________________________________

If lFct = "UpdateVO" Then
    V = sqlYEICGCC0_Update(newYEICGCC0, oldYEICGCC0, True)
    
    If Not IsNull(V) Then GoTo Error_MsgBox
    
    newYEICGCC0 = voYEICGCC0
    newYEICGCC0.EICGCCSTA = "I"
    newYEICGCC0.EICGCCXECO = "=> dossier " & dateImp10(oldYEICGCC0.EICGCCAMJ) & " " & oldYEICGCC0.EICGCCOPE & " " & oldYEICGCC0.EICGCCDOS
    V = sqlYEICGCC0_Update(newYEICGCC0, voYEICGCC0, True)
End If
'________________________________________________________________________________

If lFct = "+log" Then
    newYEICGCCLOG.EICGCCLOGH = time_Hms            ' heure maj
    newYEICGCCLOG.EICGCCLOGS = newYEICGCCLOG.EICGCCLOGS + 1
    V = sqlYEICGCCLOG_Insert(newYEICGCCLOG)
End If

If lFct = "Dos+Log" Then
    V = sqlYEICGCC0_Update(newYEICGCC0, oldYEICGCC0, True)
    
    If IsNull(V) Then
        newYEICGCCLOG.EICGCCLOGH = time_Hms            ' heure maj
        newYEICGCCLOG.EICGCCLOGS = newYEICGCCLOG.EICGCCLOGS + 1
        V = sqlYEICGCCLOG_Insert(newYEICGCCLOG)
    End If
End If
'________________________________________________________________________________

If lFct = "#Log" Then V = sqlYEICGCCLOG_Update(updYEICGCCLOG, oldYEICGCCLOG)

If lFct = "Dos#Log" Then
    V = sqlYEICGCC0_Update(newYEICGCC0, oldYEICGCC0, True)
    
    If IsNull(V) Then
        V = sqlYEICGCCLOG_Update(updYEICGCCLOG, oldYEICGCCLOG)
    End If
End If

If lFct = "Dos#Log+Val" Then
    V = sqlYEICGCC0_Update(newYEICGCC0, oldYEICGCC0, True)
    
    If IsNull(V) Then
        V = sqlYEICGCCLOG_Update(updYEICGCCLOG, oldYEICGCCLOG)
    End If
    If IsNull(V) Then
        newYEICGCCLOG.EICGCCLOGH = time_Hms            ' heure maj
        newYEICGCCLOG.EICGCCLOGS = newYEICGCCLOG.EICGCCLOGS + 1
        V = sqlYEICGCCLOG_Insert(newYEICGCCLOG)
    End If
End If

If lFct = "#Log+Val" Then
    V = sqlYEICGCCLOG_Update(updYEICGCCLOG, oldYEICGCCLOG)
    If IsNull(V) Then
        newYEICGCCLOG.EICGCCLOGH = time_Hms            ' heure maj
        newYEICGCCLOG.EICGCCLOGS = newYEICGCCLOG.EICGCCLOGS + 1
        V = sqlYEICGCCLOG_Insert(newYEICGCCLOG)
    End If
End If

If lFct = "Dos#Log#VO" Then
    V = sqlYEICGCC0_Update(newYEICGCC0, oldYEICGCC0, True)
    
    If IsNull(V) Then
        V = sqlYEICGCCLOG_Update(updYEICGCCLOG, oldYEICGCCLOG)
    End If
    If IsNull(V) Then
        newYEICGCC0 = voYEICGCC0
        newYEICGCC0.EICGCCSTA = " "
        V = sqlYEICGCC0_Update(newYEICGCC0, voYEICGCC0, True)
    End If
End If
'________________________________________________________________________________
'________________________________________________________________________________

If Not IsNull(V) Then GoTo Error_MsgBox
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYEICGCC0_Update"
Exit_sub:

    cmdYEICGCC0_Update = V
    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function

Private Sub cmdDetail_UpdateVO_Click()
Dim x As String, xSQL As String

Me.Enabled = False: Me.MousePointer = vbHourglass
    
lstW.Clear
mEICGCCXXX = "EICGCCVJPG"
xSQL = "select *   from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
     & " where EICGCCSTA = ' ' and  EICGCCVJPG > 0 and EICGCCDOS = 0"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
     x = rsSab("EICGCCID") & " - " & rsSab("EICGCCECPT") & " € " & Format$(rsSab("EICGCCEMT"), "### ### ##0.00")
    lstW.AddItem x
    rsSab.MoveNext
Loop


lstW.Visible = True

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub



Private Sub cmdParam_Add_Click()
Dim x As String
Me.Enabled = False: Me.MousePointer = vbHourglass

New_YBIATAB0 = Old_YBIATAB0
x = Trim(txtParam_K)
If x = "" Then
    Call MsgBox("Préciser le code du motif", vbCritical, "EIC_GCC : paramétrage")
Else
    New_YBIATAB0.BIATABK2 = x
    New_YBIATAB0.BIATABTXT = Trim(txtParam_X)
    If IsNull(Parametrage_New) Then lstParam_Action_Load ' 'cmdParam_Quit_Click
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Delete_Click()
Dim x As String
Me.Enabled = False: Me.MousePointer = vbHourglass

If IsNull(Parametrage_Delete) Then lstParam_Action_Load 'cmdParam_Quit_Click

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Quit_Click()

txtParam_Action = ""
libParam_Action = ""
txtParam_K = ""
txtParam_X = ""
lstParam_K.Clear

cmdParam_Quit.Visible = False
cmdParam_Add.Visible = False
cmdParam_Update.Visible = False
cmdParam_Delete.Visible = False

End Sub

Private Function Parametrage_Delete()
Dim xSQL As String
On Error GoTo Error_Handler

Dim V
App_Debug = "Parametrage_Delete"

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
V = sqlYBIATAB0_Delete(Old_YBIATAB0)
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    Parametrage_Delete = V
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function

Private Function Parametrage_New()
Dim xSQL As String
On Error GoTo Error_Handler

Dim V
App_Debug = "Parametrage_New"

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
V = sqlYBIATAB0_Insert(New_YBIATAB0)
If Not IsNull(V) Then GoTo Error_MsgBox
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    Parametrage_New = V
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function


Public Function Parametrage_Update()
Dim V

App_Debug = "Parametrage_Update"
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
V = sqlYBIATAB0_Update(New_YBIATAB0, Old_YBIATAB0)
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    
    Parametrage_Update = V

    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function


Private Sub cmdParam_Update_Click()
Dim x As String
Me.Enabled = False: Me.MousePointer = vbHourglass

New_YBIATAB0 = Old_YBIATAB0
New_YBIATAB0.BIATABTXT = Trim(txtParam_X)
If IsNull(Parametrage_Update) Then lstParam_Action_Load 'cmdParam_Quit_Click

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdPrint_Click()
Dim x As String, I As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass
x = "Gestion des chèques circulants"

Select Case SSTab1.Tab
    Case 0:
        x = cmdPrint_Title
        Select Case cmdSelect_SQL_K
            Case "1", "2e", "2v", "2?": Call cmdPrint_YEICGCC0(x)
            Case "E": Call cmdPrint_YEICGCC0_Echéancier(x)
            Case "L#": Call cmdPrint_YEICGCCLOG(x)
            Case "St": Call cmdPrint_YEICGCC0_Statistiques(x)
        End Select
End Select

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdPrint_Auto()
Dim x As String, xSQL As String, I As Integer
Dim wDSYS_5J As Long
Dim mDsys_1 As Long

Me.Enabled = False: Me.MousePointer = vbHourglass


x = "Gestion des chèques circulants"
wDSYS_5J = Val(dateElp("Ouvré", -5, DSys))

cmdSelect_SQL_K = "L#"
Call DTPicker_Set(txtSelect_EICGCCAMJ, DSys)
'optSelect_EICGCCAMJ_S = True
txtSelect_EICGCCAMJ_Min.Value = Null

cboEICGCCLOGK = ""

cmdSelect_SQL_YEICGCCLOG
x = cmdPrint_Title & " " & dateImp10(DSys)
prtYEICGCC0_Init "YEICGCCLOG", x
prtYEICGCC0_Open                            'prtYEICGCCLOG_Form
For I = 1 To arrYEICGCCLOG_Nb
    prtYEICGCCLOG_Line arrYEICGCCLOG(I)
Next I
prtYEICGCC0_Close False


cmdSelect_SQL_K = "E"
cmdSelect_SQL_Echeancier
x = cmdPrint_Title
prtYEICGCC0_Init "Echéancier", x
frmElpPrt.prtNewPage
prtYEICGCC0_Form_Echéancier
For I = 1 To arrYEICGCC0_Nb
    prtYEICGCC0_Line_Echéancier arrYEICGCC0(I), arrYEICGCCLOG(I)
Next I
prtYEICGCC0_Close False

cmdSelect_SQL_K = "2e"
cmdSelect_SQL_2e
x = cmdPrint_Title
prtYEICGCC0_Init "YEICGCC0", x
frmElpPrt.prtNewPage
prtYEICGCC0_Form
For I = 1 To arrYEICGCC0_Nb
    If arrYEICGCC0(I).EICGCCAGE < wDSYS_5J Then prtYEICGCC0_Line arrYEICGCC0(I)
Next I
prtYEICGCC0_Close False

cmdSelect_SQL_K = "2?"
cmdSelect_SQL_2vo
x = cmdPrint_Title
prtYEICGCC0_Init "YEICGCC0", x
frmElpPrt.prtNewPage
prtYEICGCC0_Form
For I = 1 To arrYEICGCC0_Nb
    prtYEICGCC0_Line arrYEICGCC0(I)
Next I
prtYEICGCC0_Close False

cmdSelect_SQL_K = "2v"
cmdSelect_SQL_2v
x = cmdPrint_Title
prtYEICGCC0_Init "YEICGCC0", x
frmElpPrt.prtNewPage
prtYEICGCC0_Form
For I = 1 To arrYEICGCC0_Nb
    prtYEICGCC0_Line arrYEICGCC0(I)
Next I
prtYEICGCC0_Close False
'_________________________________________________________________________________
mDsys_1 = DSys
xSQL = "select distinct EICGCCLOGD from " & paramIBM_Library_SABSPE_XXX & ".YEICGCCLOG " _
     & " group by EICGCCLOGD order by EICGCCLOGD desc"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF

    mDsys_1 = rsSab("EICGCCLOGD")
    If mDsys_1 < DSys Then Exit Do
    rsSab.MoveNext
Loop

cmdSelect_SQL_K = "L#"
Call DTPicker_Set(txtSelect_EICGCCAMJ, CStr(mDsys_1))
txtSelect_EICGCCAMJ_Min.Value = Null
'optSelect_EICGCCAMJ_E = True
cboEICGCCLOGK = ""

cmdSelect_SQL_YEICGCCLOG
x = cmdPrint_Title & " " & dateImp10(mDsys_1)
prtYEICGCC0_Init "YEICGCCLOG", x
frmElpPrt.prtNewPage
prtYEICGCCLOG_Form
For I = 1 To arrYEICGCCLOG_Nb
    prtYEICGCCLOG_Line arrYEICGCCLOG(I)
Next I

prtYEICGCC0_Close True

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Chèques circulants_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Reset
fgSelect.Visible = False
fraSelect_Options.Visible = False

Select Case cmdSelect_SQL_K
    Case "1", "1s": fraSelect_Options.Visible = True: cmdSelect_SQL_1
    Case "2e": cmdSelect_SQL_2e
    Case "2v": cmdSelect_SQL_2v
    Case "2?": cmdSelect_SQL_2vo
    Case "E": cmdSelect_SQL_Echeancier
    Case "Ie": cmdSelect_SQL_Import_EIC
    Case "Ir": cmdSelect_SQL_Import_REM
    Case "Iv": cmdSelect_SQL_Import_Vignettes
    Case "J": 'cmdSelect_SQL_Journalisation
    Case "L#": fraSelect_Options.Visible = True: cmdSelect_SQL_YEICGCCLOG
    Case "St": cmdSelect_SQL_Statistiques
    Case "SBq": cmdSelect_SQL_Statistiques_ChèquesBanque
    Case "Xb": cmdSelect_SQL_Exportation_Bénéficiaires
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< Chèques circulants_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus

End Sub


Public Sub cmdSelect_SQL_Exportation_Bénéficiaires()
On Error GoTo Error_Handler
Dim x As String, K As Long, xWhere As String
Dim wFile As String, wFilex As String
Dim blnCALCS As Boolean, blnZCLIGRP0 As Boolean
Dim xSQL As String

On Error GoTo Error_Handler
'===================================================================================
If blnAuto Then
    x = paramServer("\\GDMP_Archive\")
Else
    x = ""
End If
If x = "" Then x = "C:\Temp\"
If Mid$(x, Len(x), 1) <> "\" Then x = x & "\"

blnCALCS = False
If Dir(x & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

wFile = x & Trim("GDMP EIC_GCC liste des bénéficiaires " & ", au " & dateImp_Amj(DSys) & ".xlsx")

If Not blnAuto Then
    x = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "GDMP EIC_GCC : nom du fichier d'exportation", wFile)
    If Trim(x) = "" Then Exit Sub
    wFilex = Trim(x)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If

If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile
'_________________________________________


Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "YBIACPT0"
    .Subject = ""
End With


'===================================================================================

mXls1_Row_C = 1
Call cmdSelect_SQL_Exportation_Bénéficiaires_Init(1)
'==========================================================================================================
x = "select EICGCCXNOM ,count(*) from " & paramIBM_Library_SABSPE & ".YEICGCC0 " _
   & "  WHERE EICGCCOPE = 'RI0' and EICGCCSTA = 'V' GROUP BY eicgccxnom ORDER BY EICGCCXNOM"

Set rsSab = cnsab.Execute(x)
mXls2_Row = 1
Do While Not rsSab.EOF
    mXls1_Row = mXls1_Row + 1
    wsExcel.Cells(mXls1_Row, 1) = rsSab(0)
    wsExcel.Cells(mXls1_Row, 2) = rsSab(1)
    
    rsSab.MoveNext
Loop
'===================================================================================================

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
    x = "C:\Temp\"
    Resume Next
End If

    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents

End Sub

Public Sub cmdSelect_SQL_Exportation_Bénéficiaires_Init(lSheet As Integer)
Dim K As Integer, K2 As Integer

Set wsExcel = wbExcel.Sheets(lSheet)

wsExcel.Name = "EIC_GCC"
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14GDMP EIC_GCC liste des bénéficiaires" _
                            & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignLeft
    .WrapText = False ' True
    .Font.Size = 8
    .Font.Name = "Arial Unicode MS"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORPortrait
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 100
wsExcel.PageSetup.PrintTitleRows = "$A1:$B1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

wsExcel.PageSetup.CenterHorizontally = True

mXls1_Col = arrDev_Nb + 2
mXls1_Row = 1: mXls1_Row_T = 0


wsExcel.Columns(1).ColumnWidth = 32
wsExcel.Cells(mXls1_Row_C, 1) = "Bénéficiaires"
wsExcel.Cells(mXls1_Row_C, 1).Interior.Color = mColor_GB: wsExcel.Cells(mXls1_Row_C, 1).Font.Color = mColor_Z0
wsExcel.Cells(mXls1_Row_C, 2).Interior.Color = mColor_GB: wsExcel.Cells(mXls1_Row_C, 2).Font.Color = mColor_Z0

wsExcel.Columns(2).ColumnWidth = 8: wsExcel.Cells(mXls1_Row_C, 2) = "Nombre":  wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(mXls1_Col).NumberFormat = "### ### ###"


End Sub

Private Sub cmdSelect_SQL_Journalisation()
Dim V
Dim x As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_Journalisation"
blnOk = False
   
x = InputBox("indiquer " & vbCrLf & "- un numéro de dossier" & vbCrLf & "- ou une date de recherche(jj/mm/aaaa)", "journalisation : recherche")
If Trim(x) = "" Then GoTo Exit_sub
If IsNumeric(x) Then
    xWhere = " where EICGCCID = " & Val(x)
Else
    Call dateJma10_Amj(x, wAMJMin)
    xWhere = " where EICGCCUAMJ >= " & wAMJMin
End If


arrJEICGCC0_SQL xWhere

fgSelect_Display


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
Exit_sub:
End Sub
Private Sub cmdSelect_SQL_Journalisation_Action()
Dim V
Dim x As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_Journalisation"
blnOk = False
   
x = InputBox("indiquer " & vbCrLf & "- un numéro de dossier" & vbCrLf & "- ou une date de recherche(jj/mm/aaaa)", "journalisation : recherche")
If Trim(x) = "" Then GoTo Exit_sub
If IsNumeric(x) Then
    xWhere = " where EICGCCLOGI = " & Val(x)
Else
    Call dateJma10_Amj(x, wAMJMin)
    xWhere = " where EICGCCLOGD >= " & wAMJMin
End If


arrJEICGCCLOG_SQL xWhere

fgSelect_Display_YEICGCCLOG

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
Exit_sub:
End Sub


Private Sub arrJEICGCC0_SQL(xWhere As String)
Dim V
Dim x As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYEICGCC0(101)
arrYEICGCC0_Max = 100: arrYEICGCC0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABJRN & ".JEICGCC0 D , " _
        & paramIBM_Library_SABJRN & ".JRNENT0 J " _
        & xWhere & " and D.JORCV = J.JORCV and D.JOSEQN = J.JOSEQN" _
        & " order by D.JORCV , D.JOSEQN"

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsJEICGCC0_GetBuffer(rsSab, xYEICGCC0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYEICGCC0.fgselect_Display"
        '' Exit Sub
     Else
         arrYEICGCC0_Nb = arrYEICGCC0_Nb + 1
         If arrYEICGCC0_Nb > arrYEICGCC0_Max Then
             arrYEICGCC0_Max = arrYEICGCC0_Max + 100
             ReDim Preserve arrYEICGCC0(arrYEICGCC0_Max)
         End If
         
         arrYEICGCC0(arrYEICGCC0_Nb) = xYEICGCC0
    End If
    rsSab.MoveNext
Loop


Exit Sub

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrJEICGCCLOG_SQL(xWhere As String)
Dim V
Dim x As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYEICGCCLOG(101)
arrYEICGCCLOG_Max = 100: arrYEICGCCLOG_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABJRN & ".JEICGCCLOG D , " _
        & paramIBM_Library_SABJRN & ".JRNENT0 J " _
        & xWhere & " and D.JORCV = J.JORCV and D.JOSEQN = J.JOSEQN" _
        & " order by D.JORCV , D.JOSEQN"

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsJEICGCCLOG_GetBuffer(rsSab, xYEICGCCLOG)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYEICGCC0.fgselect_Display"
        '' Exit Sub
     Else
         arrYEICGCCLOG_Nb = arrYEICGCCLOG_Nb + 1
         If arrYEICGCCLOG_Nb > arrYEICGCCLOG_Max Then
             arrYEICGCCLOG_Max = arrYEICGCCLOG_Max + 100
             ReDim Preserve arrYEICGCCLOG(arrYEICGCCLOG_Max)
         End If
         
         arrYEICGCCLOG(arrYEICGCCLOG_Nb) = xYEICGCCLOG
    End If
    rsSab.MoveNext
Loop


Exit Sub

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub





Private Sub cmdSuivi_Quit_Click()
fraSuivi.Visible = False

End Sub

Private Sub cmdSuivi_Update_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

If IsNull(fraSuivi_Control) Then
        cmdYEICGCCLOG_New
        fraSuivi.Visible = False
        Call fraDetail_Display_EICGCCID(newYEICGCCLOG.EICGCCLOGI)
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub fgList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim wOrigine As String, xSQL As String
On Error Resume Next


If y <= fgList.RowHeightMin Then
    Select Case fgList.Col
        Case 0: fgList_Sort1 = 0: fgList_Sort2 = 1: fgList_Sort
        Case 1:  fgList_Sort1 = 1: fgList_Sort2 = 1: fgList_Sort ' x 1
        Case 2: fgList_Sort1 = 2: fgList_Sort2 = 2: fgList_Sort   'X 2
        Case 3: fgList_Sort1 = 3: fgList_Sort2 = 3: fgList_Sort
    End Select
Else
    If fgList.Rows > 1 Then
        'Call fgList_Color(fgList_RowClick, MouseMoveUsr.BackColor, fgList_ColorClick)
        fgList.Col = fgList_arrIndex:  arrYEICGCC0_Index = CLng(fgList.Text)
   End If
End If

End Sub


Private Sub fgLogV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim wOrigine As String
On Error Resume Next


If y <= fgLogV.RowHeightMin Then
Else
    If fgLogV.Rows > 1 Then
        blnControl = False
        Call fgLogV_Color(fgLogV_RowClick, MouseMoveUsr.BackColor, fgLogV_ColorClick)
        fgLogV.Col = fgLogV_arrIndex:  arrYEICGCCLOG_Index = CLng(fgLogV.Text)
        oldYEICGCCLOG = arrYEICGCCLOG(arrYEICGCCLOG_Index)
        xYEICGCCLOG = oldYEICGCCLOG
        
        mnuLogV_Annuler.Enabled = False
        mnuLogV_Valider.Enabled = False
        mnuLogV_Culot.Visible = False
        If oldYEICGCCLOG.EICGCCLOGA = " " Then
            mnuLogV_Annuler.Enabled = True
        
            Select Case Trim(oldYEICGCCLOG.EICGCCLOGK)
                Case "AF0", "AI1", "RQ0": mnuLogV_Valider.Enabled = True
                Case "CHQ rejeté": mnuLogV_Valider.Enabled = True
            End Select
            If oldYEICGCCLOG.EICGCCLOGE > 0 Then mnuLogV_Valider.Enabled = True
            If Trim(oldYEICGCCLOG.EICGCCLOGK) = "Mail DCOM" And oldYEICGCCLOG.EICGCCLOGE < YBIATAB0_DATE_CPT_JS1 Then
                mnuLogV_Culot.Visible = True
            End If
        End If
        Me.PopupMenu mnuLogV, vbPopupMenuLeftButton
        blnControl = True

   End If
End If

End Sub


Private Sub imgCHQ_Click()
fraCHQ_Max.Visible = True
imgCHQ_Max.Picture = imgCHQ.Picture
End Sub

Private Sub imgCHQ_Verso_Click()
fraCHQ_Max.Visible = True
imgCHQ_Max.Picture = imgCHQ_Verso.Picture

End Sub

Private Sub lstParam_Action_Click()

lstParam_Action_Load
End Sub


Private Sub lstParam_K_Click()
Dim xSQL As String
Old_YBIATAB0.BIATABK2 = lstParam_K
txtParam_K = Trim(lstParam_K)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YEICGCC0' and BIATABK1 = '" & Trim(lstParam_Action) & "'  and BIATABK2 = '" & Trim(lstParam_K) & "'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    Old_YBIATAB0.BIATABTXT = rsSab("BIATABTXT")

    txtParam_X = Trim(Old_YBIATAB0.BIATABTXT)
    cmdParam_Delete.Visible = YEICGCC0_Aut.Rapprocher
    cmdParam_Update.Visible = YEICGCC0_Aut.Rapprocher
    
Else
    txtParam_X = ""
End If

End Sub


Private Sub lstW_Click()
Select Case mEICGCCXXX
    Case "EICGCCVINT": txtDetail_EICGCCVINT = lstW.Text
    Case "EICGCCVEXT": txtDetail_EICGCCVEXT = lstW.Text
    Case "EICGCCXECO": txtDetail_EICGCCXECO = lstW.Text
    Case "EICGCCXNOM": txtDetail_EICGCCXNOM = lstW.Text
    Case "EICGCCVJPG": cmdDetail_UpdateVO_Ok
End Select
lstW.Visible = False
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
'blnControl = False
lstErr.Clear: lstErr.Height = 200


If fgList.Visible Then
    fgList.Visible = False
    Exit Sub
End If
If fraJRNENT0.Visible Then
    fraJRNENT0.Visible = False: fraDetail.Visible = False
    fraYEICGCCLOG.Visible = False
    Exit Sub
End If

If fraSuivi.Visible Then
    fraSuivi.Visible = False
    Exit Sub
End If

If lstW.Visible Then
    lstW.Visible = False
    Exit Sub
End If

If fraCHQ_Max.Visible Then
    fraCHQ_Max.Visible = False
    Exit Sub
End If
If fraCHQ.Visible Then
    fraCHQ.Visible = False
    Exit Sub
End If

If fraDetail.Visible Then
    fraCHQ.Visible = False
    fraCHQ_Max.Visible = False
    fraDetail.Visible = False: fgLogV.Visible = False: fraLogV.Visible = False
    fgList.Visible = False
    Exit Sub
End If
If fgSelect.Visible Then
    fgSelect.Visible = False
    Exit Sub
End If

If SSTab1.Tab = 0 Then
    Unload Me
End If
    Exit Sub

End Sub
Public Sub cmdContext_Return()
    If SSTab1.Tab = 0 Then
        If cmdSelect_SQL_K <> "J" And cmdSelect_SQL_K <> "J#" Then
            If Not fgSelect.Version Then cmdSelect_Ok_Click
        End If
    Else
        SendKeys "{TAB}"
    End If
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
End Sub





Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim wOrigine As String, xSQL As String
On Error Resume Next


If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_SortX 0
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_SortX 4
        Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_SortX 5
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_SortX 8
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  arrYEICGCC0_Index = CLng(fgSelect.Text)
        
        Select Case cmdSelect_SQL_K
            Case "S?", "St", "SBq"
            Case "J#":
                 oldYEICGCCLOG = arrYEICGCCLOG(arrYEICGCC0_Index)
                 fraJEICGCCLOG_Display
          Case "L#"
                oldYEICGCC0.EICGCCID = selYEICGCCLOG(arrYEICGCC0_Index).EICGCCLOGI
                If oldYEICGCC0.EICGCCID > 0 Then
                    xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
                         & " Where EICGCCID = " & oldYEICGCC0.EICGCCID
                    Set rsSab = cnsab.Execute(xSQL)
                    If Not rsSab.EOF Then
                        V = rsYEICGCC0_GetBuffer(rsSab, oldYEICGCC0)
                        xYEICGCC0 = oldYEICGCC0
                        fraDetail_Display
                    End If
                End If
          Case Else
                oldYEICGCC0 = arrYEICGCC0(arrYEICGCC0_Index)
                xYEICGCC0 = oldYEICGCC0
                fraDetail_Display
        End Select
        
   End If
End If

End Sub

Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = fgSelect.Cols - 1
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = fgSelect.FixedCols

End Sub


Public Sub fgLogV_Reset()
fgLogV.Clear
fgLogV_Sort1 = 0: fgLogV_Sort2 = 0
fgLogV_Sort1_Old = -1
fgLogV_RowDisplay = 0: fgLogV_RowClick = 0
fgLogV_arrIndex = fgLogV.Cols - 1
blnfgLogV_DisplayLine = False
fgLogV_SortAD = 6
fgLogV.LeftCol = fgLogV.FixedCols

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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






Public Sub fgSelect_ForeColor(lColor As Long)
For I = 0 To fgSelect_arrIndex
  fgSelect.Col = I: fgSelect.CellForeColor = lColor
Next I

End Sub







Private Sub mnuLogV_Annuler_Click()
Dim K As Integer, K2 As Integer, xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass

updYEICGCCLOG = oldYEICGCCLOG
updYEICGCCLOG.EICGCCLOGA = "A"


arrYEICGCCLOG(arrYEICGCCLOG_Index) = updYEICGCCLOG
If Trim(oldYEICGCCLOG.EICGCCLOGK) <> "Associer VO" Then
    Call cmdDetail_Update_Ok("#Log")
Else
'______________________________________________________________________
    K = InStr(updYEICGCCLOG.EICGCCLOGX, "-")
    K2 = InStr(updYEICGCCLOG.EICGCCLOGX, ":")
    voYEICGCC0.EICGCCID = Val(Mid$(updYEICGCCLOG.EICGCCLOGX, K + 1, K2 - K - 1))
    xSQL = "select *    from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
         & " where EICGCCID= " & voYEICGCC0.EICGCCID
    Set rsSab = cnsab.Execute(xSQL)
    If rsSab.EOF Then
        Call MsgBox("Enregistrement non trouvé : " & voYEICGCC0.EICGCCID, vbCritical, "Annulation d'une association d'une vignette orpheline")
    Else
        V = rsYEICGCC0_GetBuffer(rsSab, voYEICGCC0)
        Call cmdDetail_Update_Ok("#Log#VO")
    End If
End If


Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub mnuLogV_Culot_Click()
Dim x As String

Me.Enabled = False: Me.MousePointer = vbHourglass

rsYEICGCCLOG_Init newYEICGCCLOG
Select Case Trim(oldYEICGCCLOG.EICGCCLOGK)
    Case "Mail DCOM": newYEICGCCLOG.EICGCCLOGK = "Mail DCOM *"
    Case Else: GoTo Exit_sub
End Select

newYEICGCCLOG.EICGCCLOGA = "V"
sqlCLIENARES_Mail
If mMNUUTIMAI = "" Then
      Call MsgBox("Abandon ( manque l'adresse mail du responsable)", vbExclamation, "Mail DCOM")
Else
    txtDetail_EICGCCXECO = "conforme à l'activité du client"
    
    newYEICGCCLOG.EICGCCLOGI = oldYEICGCCLOG.EICGCCLOGI
    newYEICGCCLOG.EICGCCLOGX = " : " & txtDetail_EICGCCXECO
    
    updYEICGCCLOG = oldYEICGCCLOG
    updYEICGCCLOG.EICGCCLOGA = "V"
    arrYEICGCCLOG(arrYEICGCCLOG_Index) = updYEICGCCLOG
    Call cmdDetail_Update_Ok("#Log+Val")

End If


Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuLogV_Valider_Click()
Dim x As String, blnCOmmentaire As Boolean
Dim xSQL As String

Me.Enabled = False: Me.MousePointer = vbHourglass
blnCOmmentaire = False
rsYEICGCCLOG_Init newYEICGCCLOG
Select Case Trim(oldYEICGCCLOG.EICGCCLOGK)
    Case "AF0": newYEICGCCLOG.EICGCCLOGK = "AF0 *": blnCOmmentaire = True
    Case "AI1": newYEICGCCLOG.EICGCCLOGK = "AI1 *": blnCOmmentaire = True
    Case "RQ0": newYEICGCCLOG.EICGCCLOGK = "RQ0 *": blnCOmmentaire = True
    Case "CHQ rejeté": newYEICGCCLOG.EICGCCLOGK = "CHQ => bénéf": blnCOmmentaire = True
End Select

If Not blnCOmmentaire Then
    updYEICGCCLOG = oldYEICGCCLOG
    updYEICGCCLOG.EICGCCLOGA = "V"
    arrYEICGCCLOG(arrYEICGCCLOG_Index) = updYEICGCCLOG
    Call cmdDetail_Update_Ok("#Log")
Else
    x = InputBox("Commentaire :", "Validation d'une action")
    
    If x <> "" Then
        
        newYEICGCCLOG.EICGCCLOGI = oldYEICGCCLOG.EICGCCLOGI
        newYEICGCCLOG.EICGCCLOGX = " : " & x
        
        updYEICGCCLOG = oldYEICGCCLOG
        updYEICGCCLOG.EICGCCLOGA = "V"
        arrYEICGCCLOG(arrYEICGCCLOG_Index) = updYEICGCCLOG
        Call cmdDetail_Update_Ok("#Log+Val")
        
        xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0" _
        & " where EICGCCID = " & oldYEICGCCLOG.EICGCCLOGI
        Set rsSab = cnsab.Execute(xSQL)
    
        If Not rsSab.EOF Then
            V = rsYEICGCC0_GetBuffer(rsSab, oldYEICGCC0)

            newYEICGCC0 = oldYEICGCC0
            newYEICGCC0.EICGCCSTA = "A"
            Call cmdYEICGCC0_Update("Update")
        End If
    End If
End If
Me.Enabled = True: Me.MousePointer = 0


End Sub


Private Sub txtDetail_EICGCCVEXT_GotFocus()
'txt_GotFocus txtDetail_EICGCCVEXT
End Sub


Private Sub txtDetail_EICGCCVEXT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtDetail_EICGCCVEXT_LostFocus()
'txt_LostFocus txtDetail_EICGCCVEXT
End Sub


Private Sub txtDetail_EICGCCVINT_GotFocus()
'txt_GotFocus txtDetail_EICGCCVINT
End Sub


Private Sub txtDetail_EICGCCVINT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtDetail_EICGCCVINT_LostFocus()
'txt_LostFocus txtDetail_EICGCCVINT
End Sub


Private Sub txtDetail_EICGCCXECO_GotFocus()
'txt_GotFocus txtDetail_EICGCCXECO
If Trim(txtDetail_EICGCCXECO) = "" Then
    'txtDetail_EICGCCXECO.Visible = False
    'cboDetail_EICGCCXECO.ListIndex = 0
    cboDetail_EICGCCXECO.Visible = True
    'cboDetail_EICGCCXECO.SetFocus
End If


End Sub


Private Sub txtDetail_EICGCCXECO_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtDetail_EICGCCXECO_LostFocus()
'txt_LostFocus txtDetail_EICGCCXECO
End Sub


Private Sub txtDetail_EICGCCXNOM_Change()
If blnControl And chkEICGCCXXX = "1" Then Call lstW_SQL(txtDetail_EICGCCXNOM)

End Sub

Private Sub txtDetail_EICGCCXNOM_GotFocus()
'txt_GotFocus txtDetail_EICGCCXNOM
mEICGCCXXX = "EICGCCXNOM"
If chkEICGCCXXX = "1" And Trim(txtDetail_EICGCCXNOM) <> "" Then
    lstW.Clear
    'lstW.Visible = True
End If
End Sub


Private Sub txtDetail_EICGCCXNOM_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtDetail_EICGCCXNOM_LostFocus()
'txt_LostFocus txtDetail_EICGCCXNOM
lstW.Visible = False
End Sub


Private Sub txtLogV_X_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtSelect_EICGCCAMJ_Change()
If fgSelect.Visible Then cmdSelect_Reset

End Sub

Private Sub txtSelect_EICGCCAMJ_KeyPress(KeyAscii As Integer)
cmdSelect_Reset

End Sub

Public Function fraDetail_Control()
Dim V, x As String, blnOk As Boolean, K As Integer, wMsgBox As String
Dim Nb As Long
'fraDetail_Update.Enabled = False

wMsgBox = ""
blnOk = False
fraDetail_Control = Null

newYEICGCC0.EICGCCVINT = Trim(txtDetail_EICGCCVINT)
newYEICGCC0.EICGCCVEXT = Trim(txtDetail_EICGCCVEXT)

newYEICGCC0.EICGCCXNOM = Trim(txtDetail_EICGCCXNOM)
newYEICGCC0.EICGCCXECO = Trim(txtDetail_EICGCCXECO)


newYEICGCC0.EICGCCKMT = cboDetail_EICGCCKMT
newYEICGCC0.EICGCCKSIG = cboDetail_EICGCCKSIG
newYEICGCC0.EICGCCKEND = cboDetail_EICGCCKEND
newYEICGCC0.EICGCCKLAB = cboDetail_EICGCCKLAB

If IsNull(txtDetail_EICGCCEAMJ) Then
    newYEICGCC0.EICGCCEAMJ = 0
Else
    Call DTPicker_Control(txtDetail_EICGCCEAMJ, x)
    newYEICGCC0.EICGCCEAMJ = x
    If newYEICGCC0.EICGCCEAMJ > newYEICGCC0.EICGCCAMJ Then
        wMsgBox = " - date d'émission du chèque > date de comptabilisation" & vbCrLf
    Else
        Nb = DateDiff("d", dateImp10_S(newYEICGCC0.EICGCCEAMJ), dateImp10_S(newYEICGCC0.EICGCCAMJ))
        If Nb > 372 Then wMsgBox = " - date d'émission du chèque > 1 an et 7 jours" & vbCrLf
    End If
End If

fraDetail_Control_EICGCCSTA_Init

Call fraDetail_Control_EICGCCSTA

If wMsgBox <> "" Then
    fraDetail_Control = "?"
    Call MsgBox(wMsgBox, vbCritical, "EIC_GCC : contrôle détail")
End If

Exit_sub:
'fraDetail_Update.Enabled = True
End Function


Public Sub cmdSendMail_DCOM(lTxt As String)
Dim wSendMail As typeSendMail
Dim xHeader As String, xDétail As String, mbgColor As String
Dim wSubject As String
Dim K As Long, I As Long, I2 As Long
Dim wTxt As String
Dim xEch As String, xEch_bgColor As String
On Error Resume Next
'style="text-align: right;">

wTxt = ""
If K <= 60 Then
    wTxt = lTxt
Else
    For I = 1 To K Step 60
        wTxt = wTxt & Mid$(lTxt, I, 60) & "<BR>"
    Next I
End If

If newYEICGCCLOG.EICGCCLOGE = 0 Then
    xEch = ""
    xEch_bgColor = "#FFFFFF"
Else
    xEch = "Répondre avant le <BR> " & dateImp10(newYEICGCCLOG.EICGCCLOGE)
    xEch_bgColor = "bgcolor=#FF8000"
    
End If

mbgColor = "bgcolor = #FFFFFF"
wSendMail.Subject = "EIC_GCC " & dateImp10(oldYEICGCC0.EICGCCAMJ) & " - " & oldYEICGCC0.EICGCCID & " : " & lTxt
wSendMail.Recipient = mMNUUTIMAI
wSendMail.FromDisplayName = "EIC_GCC" '"DCOM"
wSendMail.CcDisplayName = "EIC_GCC"
wSendMail.From = currentZMNUUTI0.MNUUTIMAI
'wSendMail.FromDisplayName = "GDMP : chèques circulants"



wSendMail.Subject = wSendMail.Subject & wSubject
xHeader = "<TR>" _
         & "<TD bgcolor=#0090A0  width=200 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>" _
         & "Compte" & "</TD>" _
         & "<TD bgcolor=#0090A0  width=500 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>" _
         & "Intitulé" _
         & "<TD bgcolor=#0090A0  width=200 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>" _
         & "montant" & "</TD>" _
        & "</TR>"



xDétail = "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:8.0pt;font-family:Arial'>" & htmlFontColor_Blue _
     & libDetail_EICGCCECPT & "</TD>" _
     & "<TD  " & mbgColor & "  width=500 height=5><span style='font-size:10.0pt;font-family:Arial'>" & htmlFontColor_Blue _
     & "<B>" & libDetail_EICGCCECPT_X & "</TD>" _
     & "<TD " & mbgColor & " width=200 height=5><span style='font-size:10.0pt;font-family:Arial'>" & htmlFontColor_Blue _
     & "<B> " & libDetail_EICGCCEMT & "</B/TD>" _
     & "</TR>"
     


xDétail = xDétail & "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:10.0pt;font-family:Arial'>" & htmlFontColor_Black _
     & "objet de la demande" & "</TD>" _
     & "<TD  " & mbgColor & "  width=500 height=15><span style='font-size:10.0pt;font-family:Arial'>" & htmlFontColor_Red _
     & "<B>" & wTxt & "</TD>" _
     & "<TD " & xEch_bgColor & "   width=200 height=5><span style='font-size:10.0pt;font-family:Arial'>" & htmlFontColor_White _
     & xEch & "</B/TD>" _
     & "</TR>"



wSendMail.Attachment = paramCHQ_SCAN_Image_Archive & "\" & Trim(oldYEICGCC0.EICGCCVAMJ) & "\Archive\" & Format(oldYEICGCC0.EICGCCVJPG, "00000000") & ".jpg"

'& "<FONT face=" & Asc34 & "@Arial Unicode MS" & Asc34 & ">" _


wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                    & "<span style='font-size:12.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Blue _
                    & "Bonjour," _
                    & "<BR> Veuillez trouver ci-joint la photocopie du chèque compensation tiré sur l'un de vos clients." _
                    & "<BR> Merci de nous faire parvenir les renseignements demandés ci-après." _
                    & "<BR> Bonne réception." _
                    & "<BR><BR>" & "<TABLE border = 1  width=800 height=5 cellpadding=3 >" _
                    & xHeader _
                    & xDétail _
                    & "</TABLE><BR><BR>" ' _
                    '& " <img src=" & Asc34 & paramCHQ_SCAN_Image_Archive & "\" & Trim(oldYEICGCC0.EICGCCVAMJ) & "\Archive\" & oldYEICGCC0.EICGCCVJPG & ".jpg" & Asc34 _
                    '& " width =640 height = 275 border = 2 hspace = 15>"


wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail



End Sub
Public Sub cmdSendMail_DCOM_Culot(lTxt As String)
Dim wSendMail As typeSendMail
Dim xHeader As String, xDétail As String, mbgColor As String
Dim wSubject As String
Dim K As Long, I As Long, I2 As Long
Dim wTxt As String
Dim xEch As String, xEch_bgColor As String
On Error Resume Next
'style="text-align: right;">

wTxt = ""
If K <= 60 Then
    wTxt = lTxt
Else
    For I = 1 To K Step 60
        wTxt = wTxt & Mid$(lTxt, I, 60) & "<BR>"
    Next I
End If

xEch = "sans réponse au <BR> " & dateImp10(YBIATAB0_DATE_CPT_JS1)
xEch_bgColor = "bgcolor=#FF80FF"
    
mbgColor = "bgcolor = #FFFFFF"
wSendMail.Subject = "EIC_GCC " & dateImp10(oldYEICGCC0.EICGCCAMJ) & " - " & oldYEICGCC0.EICGCCID & " : " & lTxt
wSendMail.Recipient = mMNUUTIMAI
wSendMail.FromDisplayName = "EIC_GCC" '"DCOM"
wSendMail.CcDisplayName = "EIC_GCC"
wSendMail.From = currentZMNUUTI0.MNUUTIMAI
'wSendMail.FromDisplayName = "GDMP : chèques circulants"



wSendMail.Subject = wSendMail.Subject & wSubject
xHeader = "<TR>" _
         & "<TD bgcolor=#0090A0  width=200 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>" _
         & "Compte" & "</TD>" _
         & "<TD bgcolor=#0090A0  width=500 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>" _
         & "Intitulé" _
         & "<TD bgcolor=#0090A0  width=200 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>" _
         & "montant" & "</TD>" _
        & "</TR>"



xDétail = "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:8.0pt;font-family:Arial'>" & htmlFontColor_Blue _
     & libDetail_EICGCCECPT & "</TD>" _
     & "<TD  " & mbgColor & "  width=500 height=5><span style='font-size:10.0pt;font-family:Arial'>" & htmlFontColor_Blue _
     & "<B>" & libDetail_EICGCCECPT_X & "</TD>" _
     & "<TD " & mbgColor & " width=200 height=5><span style='font-size:10.0pt;font-family:Arial'>" & htmlFontColor_Blue _
     & "<B> " & libDetail_EICGCCEMT & "</B/TD>" _
     & "</TR>"
     


xDétail = xDétail & "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:10.0pt;font-family:Arial'>" & htmlFontColor_Black _
     & "motif économique" & "</TD>" _
     & "<TD  " & mbgColor & "  width=500 height=15><span style='font-size:10.0pt;font-family:Arial'>" & htmlFontColor_Red _
     & "<B>" & wTxt & "</TD>" _
     & "<TD " & xEch_bgColor & "   width=200 height=5><span style='font-size:10.0pt;font-family:Arial'>" & htmlFontColor_White _
     & xEch & "</B/TD>" _
     & "</TR>"



wSendMail.Attachment = paramCHQ_SCAN_Image_Archive & "\" & Trim(oldYEICGCC0.EICGCCVAMJ) & "\Archive\" & Format(oldYEICGCC0.EICGCCVJPG, "00000000") & ".jpg"

'& "<FONT face=" & Asc34 & "@Arial Unicode MS" & Asc34 & ">" _


wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                    & "<span style='font-size:12.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Blue _
                    & "Bonjour," _
                    & "<BR> Sans réponse de votre part, le dossier est classé avec la mention :" _
                    & htmlFontColor_Red & "<BR><B> conforme à l'activité du client.</B>" _
                    & htmlFontColor_Blue & "<BR> Bonne réception." _
                    & "<BR><BR>" & "<TABLE border = 1  width=800 height=5 cellpadding=3 >" _
                    & xHeader _
                    & xDétail _
                    & "</TABLE><BR><BR>" ' _
                    '& " <img src=" & Asc34 & paramCHQ_SCAN_Image_Archive & "\" & Trim(oldYEICGCC0.EICGCCVAMJ) & "\Archive\" & oldYEICGCC0.EICGCCVJPG & ".jpg" & Asc34 _
                    '& " width =640 height = 275 border = 2 hspace = 15>"


wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail



End Sub






Private Sub txtSelect_EICGCCECHQ_GotFocus()
txt_GotFocus txtSelect_EICGCCECHQ
If fgSelect.Visible Then cmdSelect_Reset

End Sub


Private Sub txtSelect_EICGCCECHQ_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtSelect_EICGCCECHQ_LostFocus()
txt_LostFocus txtSelect_EICGCCECHQ
End Sub

Private Sub txtSelect_EICGCCECLI_GotFocus()
txt_GotFocus txtSelect_EICGCCECLI
If fgSelect.Visible Then cmdSelect_Reset

End Sub


Private Sub txtSelect_EICGCCECLI_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtSelect_EICGCCECLI_LostFocus()
txt_LostFocus txtSelect_EICGCCECLI
End Sub

Private Sub txtSelect_EICGCCECPT_GotFocus()
txt_GotFocus txtSelect_EICGCCECPT
If fgSelect.Visible Then cmdSelect_Reset

End Sub


Private Sub txtSelect_EICGCCECPT_LostFocus()
txt_LostFocus txtSelect_EICGCCECPT

End Sub


Private Sub txtSelect_EICGCCID_GotFocus()
txt_GotFocus txtSelect_EICGCCID
If fgSelect.Visible Then cmdSelect_Reset

End Sub

Private Sub txtSelect_EICGCCID_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub


Private Sub txtSelect_EICGCCID_LostFocus()
txt_LostFocus txtSelect_EICGCCID

End Sub

Private Sub txtSelect_EICGCCXNOM_Change()

'Dim K As Integer, X As String
'If blnControl Then
'    fraDetail.Visible = False
'    X = Trim(txtSelect_EICGCCXNOM)
'End If
End Sub


Private Sub txtSelect_EICGCCXNOM_GotFocus()
txt_GotFocus txtSelect_EICGCCXNOM
If fgSelect.Visible Then cmdSelect_Reset

End Sub


Private Sub txtSelect_EICGCCXNOM_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub



Public Sub cmdSelect_SQL_Import_EIC()
Dim V
Dim x As String, blnOk As Boolean, K As Long
Dim xSQL As String
Dim wEICGCCAMJ As Long, Nb As Long
On Error GoTo Error_Handler
currentAction = "cmdSelect_SQL_Import_EIC"

'_________________________________________________________________________________
xSQL = "select EICGCCAMJ from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
     & "  Where EICGCCOPE = 'RI0' order by EICGCCAMJ desc"
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    wEICGCCAMJ = 20090101
Else
    wEICGCCAMJ = rsSab("EICGCCAMJ")
End If

'++++++++++++++++++++++++++++++++++++++++++
rsYEICGCCLOG_Init newYEICGCCLOG
newYEICGCCLOG.EICGCCLOGK = "Import EIC"
newYEICGCCLOG.EICGCCLOGX = "date dernière compta : " & wEICGCCAMJ
cmdYEICGCCLOG_New
'++++++++++++++++++++++++++++++++++++++++++

'_________________________________________________________________________________

xSQL = "select * from " & paramIBM_Library_SAB & ".ZEICRIC0 " _
     & " Where EICRICDCP >= " & wEICGCCAMJ - 19000000 _
     & " and EICRICOPE = 'RI0' and EICRICOPR = '160'order by  EICRICNU1"

Set rsSab = cnsab.Execute(xSQL)

Call cmdSelect_SQL_Import_EIC_Update(True, Nb)


'++++++++++++++++++++++++++++++++++++++++++
newYEICGCCLOG.EICGCCLOGK = "Import EIC"
newYEICGCCLOG.EICGCCLOGI = 0
newYEICGCCLOG.EICGCCLOGE = 0
newYEICGCCLOG.EICGCCLOGA = ""
If arrYEICGCC0_Nb = 0 Then
    newYEICGCCLOG.EICGCCLOGX = "pas de chèques circulants trouvés"
    cmdYEICGCCLOG_New
Else
    newYEICGCCLOG.EICGCCLOGX = "nb RI0 ajoutés : " & Nb & " (ID : " & mEICGCCID_0 & "-" & mEICGCCID & ")"
    cmdYEICGCCLOG_New
End If
'++++++++++++++++++++++++++++++++++++++++++

'_________________________________________________________________________________ ' date de création

xSQL = "select * from " & paramIBM_Library_SAB & ".ZEICRIC0 " _
     & " Where EICRICDCR >= " & wEICGCCAMJ - 19000000 _
     & " and EICRICOPE = 'RQ0' and EICRICOPR = '169'order by  EICRICNU1"

Set rsSab = cnsab.Execute(xSQL)

Call cmdSelect_SQL_Import_EIC_Update(True, Nb)


Call cmdSelect_SQL_Import_EIC_560(wEICGCCAMJ)

Call cmdSelect_SQL_Import_EIC_AF0_Ok

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
Exit_sub:
Call lstErr_AddItem(lstErr, cmdContext, "- Chèques circulants ajoutés : " & Nb): DoEvents
cboSelect_SQL.ListIndex = 0

End Sub
Public Sub cmdSelect_SQL_Import_EIC_560(wEICGCCAMJ As Long)
Dim V
Dim x As String, blnOk As Boolean, K As Long
Dim xSQL As String
Dim Nb As Long, Nb_AF0 As Long
On Error GoTo Error_Handler
currentAction = "cmdSelect_SQL_Import_EIC"
Dim arrEICAO2OPE() As String, arrNb As Long
Dim arrEICAO2NU1() As Long
Dim arrEICAO2OPR() As String
Dim arrEICAO2NUR() As Long
'_________________________________________________________________________________

'++++++++++++++++++++++++++++++++++++++++++
newYEICGCCLOG.EICGCCLOGK = "Import 560"
newYEICGCCLOG.EICGCCLOGI = 0
newYEICGCCLOG.EICGCCLOGE = 0
newYEICGCCLOG.EICGCCLOGA = ""
newYEICGCCLOG.EICGCCLOGX = "date dernier envoi : " & wEICGCCAMJ
cmdYEICGCCLOG_New
'++++++++++++++++++++++++++++++++++++++++++
xSQL = "select count(*) as Tally    from " & paramIBM_Library_SAB & ".ZEICAO20 " _
     & " Where EICAO2DNV >= " & wEICGCCAMJ - 19000000
Set rsSab = cnsab.Execute(xSQL)

K = 10
If Not rsSab.EOF Then K = rsSab("Tally") + 10
ReDim arrEICAO2OPE(K)
ReDim arrEICAO2NU1(K)
ReDim arrEICAO2OPR(K)
ReDim arrEICAO2NUR(K)
'_________________________________________________________________________________

xSQL = "select * from " & paramIBM_Library_SAB & ".ZEICAO20 " _
     & " Where EICAO2DNV >= " & wEICGCCAMJ - 19000000 _
     & " order by  EICAO2DNV, EICAO2NU1"

Set rsSab = cnsab.Execute(xSQL)
arrNb = 0
Nb = 0
Nb_AF0 = 0
Do While Not rsSab.EOF
         arrNb = arrNb + 1
         
         arrEICAO2OPE(arrNb) = rsSab("EICAO2OPE")
         arrEICAO2NU1(arrNb) = rsSab("EICAO2NU1")
         arrEICAO2OPR(arrNb) = rsSab("EICAO2OPR")
         arrEICAO2NUR(arrNb) = rsSab("EICAO2NUR")
    rsSab.MoveNext
Loop


For K = 1 To arrNb
    xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
         & " Where EICGCCOPE = '" & arrEICAO2OPR(K) & "'" _
         & " and EICGCCDOS = " & arrEICAO2NUR(K)
         
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        V = rsYEICGCC0_GetBuffer(rsSab, oldYEICGCC0)
        
        xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YEICGCCLOG " _
             & " Where EICGCCLOGI = " & oldYEICGCC0.EICGCCID _
             & " and EICGCCLOGK = '" & arrEICAO2OPE(K) & "'"
             
        Set rsSab = cnsab.Execute(xSQL)
        If rsSab.EOF Then
        
             Nb = Nb + 1
             Select Case arrEICAO2OPE(K)
                 Case "AI1"
                     newYEICGCC0 = oldYEICGCC0
                     newYEICGCC0.EICGCCSTA = "R"
                     
                     '++++++++++++++++++++++++++++++++++++++++++
                     newYEICGCCLOG.EICGCCLOGK = "AI1"
                     newYEICGCCLOG.EICGCCLOGI = oldYEICGCC0.EICGCCID
                     newYEICGCCLOG.EICGCCLOGE = DSys
                     newYEICGCCLOG.EICGCCLOGA = ""
                     newYEICGCCLOG.EICGCCLOGX = ": Chèque circulant rejeté : " & oldYEICGCC0.EICGCCECHQ & " / Id: " & oldYEICGCC0.EICGCCID & " retour vignette au remettant"
                     '++++++++++++++++++++++++++++++++++++++++++
                     
                     V = cmdYEICGCC0_Update("Dos+Log")
                  Case "AF0"
                     
                     '++++++++++++++++++++++++++++++++++++++++++
                     newYEICGCCLOG.EICGCCLOGK = "AF0"
                     newYEICGCCLOG.EICGCCLOGI = oldYEICGCC0.EICGCCID
                     newYEICGCCLOG.EICGCCLOGE = DSys
                     newYEICGCCLOG.EICGCCLOGA = ""
                     newYEICGCCLOG.EICGCCLOGX = ": demande de télécopie : " & oldYEICGCC0.EICGCCECHQ & " / SAB : " & oldYEICGCC0.EICGCCOPE & "  " & oldYEICGCC0.EICGCCDOS & " / Id: " & oldYEICGCC0.EICGCCID
                     cmdYEICGCCLOG_New
                     '++++++++++++++++++++++++++++++++++++++++++
                     
            End Select
        End If
    Else
        If arrEICAO2OPE(K) = "AF0" Then
            V = cmdSelect_SQL_Import_EIC_560_New(arrEICAO2OPR(K), arrEICAO2NUR(K))
            If IsNull(V) Then Nb_AF0 = Nb_AF0 + 1
        End If
    End If
    
Next K


'++++++++++++++++++++++++++++++++++++++++++
newYEICGCCLOG.EICGCCLOGK = "Import 560"
newYEICGCCLOG.EICGCCLOGI = 0
newYEICGCCLOG.EICGCCLOGE = 0
newYEICGCCLOG.EICGCCLOGA = ""
newYEICGCCLOG.EICGCCLOGX = Nb & " + " & Nb_AF0 & " (AF0) / " & arrNb & " concernant les chèques circulants trouvés"
cmdYEICGCCLOG_New
'++++++++++++++++++++++++++++++++++++++++++

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
Exit_sub:
Call lstErr_AddItem(lstErr, cmdContext, "- Chèques circulants ajoutés : " & Nb): DoEvents

End Sub

Public Sub cmdSelect_SQL_Import_EIC_Update(blnFiltre As Boolean, Nb As Long)
Dim V
Dim x As String, blnOk As Boolean, K As Long
Dim xSQL As String
Dim wEICGCCEMT As Currency
Dim wEICGCCEIND As String

On Error GoTo Error_Handler
currentAction = "cmdSelect_SQL_Import_EIC"

ReDim arrYEICGCC0(101)
arrYEICGCC0_Max = 100: arrYEICGCC0_Nb = 0
Nb = 0
'_________________________________________________________________________________
Do While Not rsSab.EOF
    blnOk = False
    wEICGCCEMT = CCur(rsSab("EICRICMON")) / 100
    wEICGCCEIND = Mid$(rsSab("EICRICD71"), 56, 1)
    
    If blnFiltre Then
        If wEICGCCEMT >= 5000 Then blnOk = True
        If wEICGCCEIND <> "0" Then blnOk = True
    Else
        blnOk = True
    End If
    
    If blnOk Then
        rsYEICGCC0_Init newYEICGCC0
        newYEICGCC0.EICGCCETB = rsSab("EICRICETB")
        newYEICGCC0.EICGCCAGE = rsSab("EICRICAGE")
        newYEICGCC0.EICGCCSER = rsSab("EICRICSER")
        newYEICGCC0.EICGCCSSE = rsSab("EICRICSSE")
        newYEICGCC0.EICGCCOPE = rsSab("EICRICOPE")
        newYEICGCC0.EICGCCDOS = rsSab("EICRICNU1")
        newYEICGCC0.EICGCCAMJ = rsSab("EICRICDCP") + 19000000
        If newYEICGCC0.EICGCCAMJ = 19000000 Then newYEICGCC0.EICGCCAMJ = DSys
        
        newYEICGCC0.EICGCCECPT = rsSab("EICRICCPD")
        newYEICGCC0.EICGCCEMT = wEICGCCEMT
        newYEICGCC0.EICGCCECHQ = rsSab("EICRICCHE")
        newYEICGCC0.EICGCCEIND = wEICGCCEIND
        newYEICGCC0.EICGCCXBQ = rsSab("EICRICDON")
                
        arrYEICGCC0_Nb = arrYEICGCC0_Nb + 1
        If arrYEICGCC0_Nb > arrYEICGCC0_Max Then
            arrYEICGCC0_Max = arrYEICGCC0_Max + 100
            ReDim Preserve arrYEICGCC0(arrYEICGCC0_Max)
         End If

        arrYEICGCC0(arrYEICGCC0_Nb) = newYEICGCC0
    End If
        
    rsSab.MoveNext
Loop

Set rsSab = Nothing

If arrYEICGCC0_Nb > 0 Then
    Call cmdSelect_SQL_Import_EICGCCID
    
    For K = 1 To arrYEICGCC0_Nb
        blnOk = True
        'If wEICGCCAMJ = arrYEICGCC0(K).EICGCCAMJ Then
            xSQL = "select EICGCCOPE EICGCCDOS from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
                 & " Where EICGCCOPE = '" & arrYEICGCC0(K).EICGCCOPE & "' and EICGCCDOS =" & arrYEICGCC0(K).EICGCCDOS
            Set rsSab = cnsab.Execute(xSQL)
            If Not rsSab.EOF Then blnOk = False
        'End If
        If blnOk Then
            Nb = Nb + 1
            mEICGCCID = mEICGCCID + 1
            arrYEICGCC0(K).EICGCCID = mEICGCCID
            
            arrYEICGCC0(K).EICGCCECLI = cmdSelect_SQL_Import_EICGCCECPT(arrYEICGCC0(K).EICGCCECPT)
'___________________________________________________________________________
            If arrYEICGCC0(K).EICGCCOPE = "RQ0" Then
                newYEICGCCLOG.EICGCCLOGI = arrYEICGCC0(K).EICGCCID
                newYEICGCCLOG.EICGCCLOGK = "RQ0"
                newYEICGCCLOG.EICGCCLOGA = ""
                newYEICGCCLOG.EICGCCLOGX = ": requête remettant : " & oldYEICGCC0.EICGCCECHQ & " / Id: " & oldYEICGCC0.EICGCCID
                newYEICGCCLOG.EICGCCLOGE = YBIATAB0_DATE_CPT_JS1
                cmdYEICGCCLOG_New
            End If
'____________________________________________________________________________________

        End If
    Next K
    
    Call cmdYEICGCC0_Update("Import")
    
End If


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
Exit_sub:
Call lstErr_AddItem(lstErr, cmdContext, "- Chèques circulants ajoutés : " & Nb): DoEvents
cboSelect_SQL.ListIndex = 0

End Sub


Public Sub cmdSelect_SQL_Import_REM()
Dim V
Dim x As String, blnOk As Boolean, K As Long
Dim xSQL As String
Dim wEICGCCAMJ As Long, wEICGCCEMT As Currency
Dim wEICGCCEIND As String
Dim Nb As Long
On Error GoTo Error_Handler
currentAction = "cmdSelect_SQL_Import_REM"

ReDim arrYEICGCC0(101)
arrYEICGCC0_Max = 100: arrYEICGCC0_Nb = 0
Nb = 0
'_________________________________________________________________________________
xSQL = "select EICGCCAMJ from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
     & "  Where EICGCCOPE = 'REM' order by EICGCCAMJ desc"
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    wEICGCCAMJ = 20090101
Else
    wEICGCCAMJ = rsSab("EICGCCAMJ")
End If

'++++++++++++++++++++++++++++++++++++++++++
rsYEICGCCLOG_Init newYEICGCCLOG
newYEICGCCLOG.EICGCCLOGK = "Import REM"
newYEICGCCLOG.EICGCCLOGX = "date dernière compta : " & wEICGCCAMJ
cmdYEICGCCLOG_New
'++++++++++++++++++++++++++++++++++++++++++

'_________________________________________________________________________________

xSQL = "select * from " & paramIBM_Library_SAB & ".ZGUIRC20, " _
    & paramIBM_Library_SAB & ".ZGUIRC10 where GUIRC1DOS = GUIRC2DOS " _
    & " and  GUIRC1ETA = GUIRC2ETA   and  GUIRC1AGE = GUIRC2AGE " _
    & " and  GUIRC1SER = GUIRC2SER and  GUIRC1SSE = GUIRC2SSE and  GUIRC1OPE = GUIRC2OPE " _
     & " and GUIRC2DT3 >= " & wEICGCCAMJ - 19000000 _
     & " and GUIRC2OPE = 'REM'  and GUIRC1NAT = '001' and GUIRC2CTA = 3 order by  GUIRC2DOS"
     
'     & " and GUIRC2MT1 >= 5000" _


Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    blnOk = True 'False
    wEICGCCEMT = CCur(rsSab("GUIRC2MT1"))
    'If wEICGCCEMT >= 5000 Then blnOk = True
    If blnOk Then
        rsYEICGCC0_Init newYEICGCC0
        newYEICGCC0.EICGCCETB = rsSab("GUIRC2ETA")
        newYEICGCC0.EICGCCAGE = rsSab("GUIRC2AGE")
        newYEICGCC0.EICGCCSER = rsSab("GUIRC2SER")
        newYEICGCC0.EICGCCSSE = rsSab("GUIRC2SSE")
        newYEICGCC0.EICGCCOPE = rsSab("GUIRC2OPE")
        newYEICGCC0.EICGCCDOS = rsSab("GUIRC2DOS")
        newYEICGCC0.EICGCCAMJ = rsSab("GUIRC2DT3") + 19000000
        newYEICGCC0.EICGCCECPT = rsSab("GUIRC2CPT")
        newYEICGCC0.EICGCCEMT = wEICGCCEMT
        newYEICGCC0.EICGCCECHQ = Format$(rsSab("GUIRC2CHQ"), "0000000")
        newYEICGCC0.EICGCCXCPT = rsSab("GUIRC1CP2")
        
        arrYEICGCC0_Nb = arrYEICGCC0_Nb + 1
        If arrYEICGCC0_Nb > arrYEICGCC0_Max Then
            arrYEICGCC0_Max = arrYEICGCC0_Max + 100
            ReDim Preserve arrYEICGCC0(arrYEICGCC0_Max)
         End If

        arrYEICGCC0(arrYEICGCC0_Nb) = newYEICGCC0
    End If
        
    rsSab.MoveNext
Loop

Set rsSab = Nothing

If arrYEICGCC0_Nb > 0 Then
    Call cmdSelect_SQL_Import_EICGCCID
    
    For K = 1 To arrYEICGCC0_Nb
        blnOk = True
        If wEICGCCAMJ = arrYEICGCC0(K).EICGCCAMJ Then
            xSQL = "select EICGCCOPE EICGCCDOS from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
                 & " Where EICGCCOPE = 'REM' and EICGCCDOS =" & arrYEICGCC0(K).EICGCCDOS
            Set rsSab = cnsab.Execute(xSQL)
            If Not rsSab.EOF Then blnOk = False
        End If
        If blnOk Then
            Nb = Nb + 1
            mEICGCCID = mEICGCCID + 1
            arrYEICGCC0(K).EICGCCID = mEICGCCID
            arrYEICGCC0(K).EICGCCXBQ = strSocBdfE

            arrYEICGCC0(K).EICGCCECLI = cmdSelect_SQL_Import_EICGCCECPT(arrYEICGCC0(K).EICGCCECPT)
            
            xSQL = "select CLIENACLI,COMPTEINT from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
                 & " Where COMPTECOM = '" & arrYEICGCC0(K).EICGCCXCPT & "'"
            Set rsSab = cnsab.Execute(xSQL)
            If Not rsSab.EOF Then
                arrYEICGCC0(K).EICGCCXID = rsSab("CLIENACLI")
                arrYEICGCC0(K).EICGCCXNOM = rsSab("COMPTEINT")
                arrYEICGCC0(K).EICGCCKLAB = "I"
            End If

        End If
    Next K
    
    Call cmdYEICGCC0_Update("Import")
    
End If


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
Exit_sub:
Call lstErr_AddItem(lstErr, cmdContext, "- Chèques circulants ajoutés : " & Nb): DoEvents
'++++++++++++++++++++++++++++++++++++++++++
newYEICGCCLOG.EICGCCLOGK = "Import REM"
newYEICGCCLOG.EICGCCLOGI = 0
newYEICGCCLOG.EICGCCLOGE = 0
newYEICGCCLOG.EICGCCLOGA = ""
If arrYEICGCC0_Nb = 0 Then
    newYEICGCCLOG.EICGCCLOGX = "pas de chèques circulants trouvés"
    cmdYEICGCCLOG_New
Else
    newYEICGCCLOG.EICGCCLOGX = "nb REM ajoutés : " & Nb & " (ID : " & mEICGCCID_0 & "-" & mEICGCCID & ")"
    cmdYEICGCCLOG_New
End If
'++++++++++++++++++++++++++++++++++++++++++
cboSelect_SQL.ListIndex = 0

End Sub


Public Sub cmdSelect_SQL_Import_Vignettes()
Dim V
Dim x As String, blnOk As Boolean, K As Long
Dim xSQL As String
Dim wEICGCCVREM As Long, wEICGCCEMT As Currency
Dim wEICGCCEIND As String
Dim Nb_REM As Long, Nb_RI0 As Long
Dim blnGCC As Boolean

On Error GoTo Error_Handler
currentAction = "cmdSelect_SQL_Import_Vignettes"

ReDim arrYEICGCC0(101)


arrYEICGCC0_Max = 100: arrYEICGCC0_Nb = 0
Nb_REM = 0: Nb_RI0 = 0

cnAdo_CHQ_ARCHIVE.Open paramODBC_DSN_CHQ_SCAN_ARCHIVE

x = paramCHQ_SCAN_Appli_Archive & "\CHEQUE"
If UCase$(x) <> UCase$(cnAdo_CHQ_ARCHIVE.DefaultDatabase) Then
    MsgBox x, vbCritical, "DSN 'CHQ_ARCHIVE' non conforme "
    cnAdo_Info cnAdo_CHQ_ARCHIVE
    End
End If

'_________________________________________________________________________________
xSQL = "select EICGCCVREM from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
     & "  where EICGCCVREM > 0 order by EICGCCVREM desc"
Set rsSab = cnsab.Execute(xSQL)
wEICGCCVREM = 0
If Not rsSab.EOF Then wEICGCCVREM = rsSab("EICGCCVREM")

If wEICGCCVREM = 0 Then wEICGCCVREM = 16402      '20090101 ' 20090801 18304 '



'++++++++++++++++++++++++++++++++++++++++++
rsYEICGCCLOG_Init newYEICGCCLOG
newYEICGCCLOG.EICGCCLOGK = "Import V"
newYEICGCCLOG.EICGCCLOGX = "dernier numéro Crem : " & wEICGCCVREM
cmdYEICGCCLOG_New
'++++++++++++++++++++++++++++++++++++++++++

'_________________________________________________________________________________

xSQL = "select * from CHEQUE where CREm > '" & Format(wEICGCCVREM, "00000000") & "' order by Crem, IMAGE"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

Do While Not rsSab.EOF
   ' V = srvCHQ_SCAN_GetBuffer_ODBC(rsSab, xCHQ_SCAN)
    xCHQ_SCAN.Id = rsSab("ID")
    If xCHQ_SCAN.Id = "R" Then
        rsYEICGCC0_Init zYEICGCC0
        zYEICGCC0.EICGCCXCPT = rsSab("COMPTE")
        zYEICGCC0.EICGCCVREM = rsSab("CRem")
        zYEICGCC0.EICGCCVAMJ = rsSab("Date")
        zYEICGCC0.EICGCCAMJ = zYEICGCC0.EICGCCVAMJ
        If rsSab("Nature") = "GCC" Then
            blnGCC = True
        Else
            blnGCC = False
        End If
    Else
        x = rsSab("Zone3")
        'If InStr(X, "2179978") > 0 Then
        If Mid$(x, 6, 4) = "2179" Then
            blnOk = True 'False
        Else
            blnOk = blnGCC
        End If
            
            wEICGCCEMT = CCur(Val(rsSab("Zone1"))) / 100
            ''If wEICGCCEMT >= 5000 Then blnOk = True
            
        '   ?indice de circulation
        
            If blnOk Then
                newYEICGCC0 = zYEICGCC0
                newYEICGCC0.EICGCCETB = 1
                newYEICGCC0.EICGCCAGE = 1
                newYEICGCC0.EICGCCSER = "00"
                newYEICGCC0.EICGCCSSE = "GU"
                newYEICGCC0.EICGCCOPE = "XXX"
                newYEICGCC0.EICGCCDOS = 0
                newYEICGCC0.EICGCCECPT = Mid$(rsSab("zone2"), 2, 11)
                newYEICGCC0.EICGCCEMT = wEICGCCEMT
                newYEICGCC0.EICGCCECHQ = rsSab("zone4")
                newYEICGCC0.EICGCCEIND = Mid$(rsSab("zone3"), 12, 1)
                newYEICGCC0.EICGCCVJPG = rsSab("IMAGE")
                        
                arrYEICGCC0_Nb = arrYEICGCC0_Nb + 1
                If arrYEICGCC0_Nb > arrYEICGCC0_Max Then
                    arrYEICGCC0_Max = arrYEICGCC0_Max + 100
                    ReDim Preserve arrYEICGCC0(arrYEICGCC0_Max)
                End If
        
                arrYEICGCC0(arrYEICGCC0_Nb) = newYEICGCC0
            End If
        End If
    rsSab.MoveNext
Loop



Set rsSab = Nothing

If arrYEICGCC0_Nb > 0 Then
    Call cmdSelect_SQL_Import_EICGCCID
    
    For K = 1 To arrYEICGCC0_Nb
        blnOk = True
        'If arrYEICGCC0(K).EICGCCECHQ = "1002176" Then
        '    Debug.Print
        'End If
        arrYEICGCC0(K).EICGCCECLI = cmdSelect_SQL_Import_EICGCCECPT(arrYEICGCC0(K).EICGCCECPT)

        xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
             & " Where EICGCCECPT = '" & arrYEICGCC0(K).EICGCCECPT _
             & "' and EICGCCECHQ ='" & arrYEICGCC0(K).EICGCCECHQ & "'" _
             & "  and EICGCCVAMJ = 0"
             
        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            blnOk = False
            V = rsYEICGCC0_GetBuffer(rsSab, oldYEICGCC0)
            If Not IsNull(V) Then
                '++++++++++++++++++++++++++++++++++++++++++
                newYEICGCCLOG.EICGCCLOGK = "Import V"
                newYEICGCCLOG.EICGCCLOGX = "# lecture YEICGCC0 " & arrYEICGCC0(K).EICGCCECPT & " chq " & arrYEICGCC0(K).EICGCCECHQ
                cmdYEICGCCLOG_New
                '++++++++++++++++++++++++++++++++++++++++++
            Else
                If oldYEICGCC0.EICGCCEMT <> arrYEICGCC0(K).EICGCCEMT Then
                    '++++++++++++++++++++++++++++++++++++++++++
                    newYEICGCCLOG.EICGCCLOGK = "Import V"
                    newYEICGCCLOG.EICGCCLOGX = "# montant EIC/scan " & arrYEICGCC0(K).EICGCCECPT & " chq " & arrYEICGCC0(K).EICGCCECHQ
                    cmdYEICGCCLOG_New
                    '++++++++++++++++++++++++++++++++++++++++++
                End If
                If oldYEICGCC0.EICGCCVAMJ > 0 Then
                    If oldYEICGCC0.EICGCCVAMJ <> arrYEICGCC0(K).EICGCCVAMJ _
                    Or oldYEICGCC0.EICGCCVJPG <> arrYEICGCC0(K).EICGCCVJPG Then
                       '++++++++++++++++++++++++++++++++++++++++++
                        newYEICGCCLOG.EICGCCLOGK = "Import V"
                        newYEICGCCLOG.EICGCCLOGX = "# EIC/scan déjà rapproché " & arrYEICGCC0(K).EICGCCECPT & " chq " & arrYEICGCC0(K).EICGCCECHQ _
                                                 & " scan : " & arrYEICGCC0(K).EICGCCVAMJ & "-" & arrYEICGCC0(K).EICGCCVJPG _
                                                 & oldYEICGCC0.EICGCCVAMJ & "-" & oldYEICGCC0.EICGCCVJPG
                        cmdYEICGCCLOG_New
                        '++++++++++++++++++++++++++++++++++++++++++
                    End If
                Else
                    newYEICGCC0 = oldYEICGCC0
                    newYEICGCC0.EICGCCVAMJ = arrYEICGCC0(K).EICGCCVAMJ
                    newYEICGCC0.EICGCCVJPG = arrYEICGCC0(K).EICGCCVJPG
                    newYEICGCC0.EICGCCVREM = arrYEICGCC0(K).EICGCCVREM
                    If newYEICGCC0.EICGCCOPE = "REM" Then
                        newYEICGCC0.EICGCCEIND = arrYEICGCC0(K).EICGCCEIND
                        If newYEICGCC0.EICGCCEIND = "8" And newYEICGCC0.EICGCCEMT < 5000 Then
                            newYEICGCC0.EICGCCSTA = "@"
                        End If
                    End If
                    V = cmdYEICGCC0_Update("Update")
                    If IsNull(V) Then
                        Nb_RI0 = Nb_RI0 + 1
                    Else
                        '++++++++++++++++++++++++++++++++++++++++++
                        newYEICGCCLOG.EICGCCLOGK = "Import V"
                        newYEICGCCLOG.EICGCCLOGX = "# màj YEICGCC0 " & arrYEICGCC0(K).EICGCCECPT & " chq " & arrYEICGCC0(K).EICGCCECHQ _
                                             & " scan : " & arrYEICGCC0(K).EICGCCVAMJ & "-" & arrYEICGCC0(K).EICGCCVJPG
                        cmdYEICGCCLOG_New
                        '++++++++++++++++++++++++++++++++++++++++++
                    End If
                End If
            End If
        End If
        If blnOk Then
            Nb_REM = Nb_REM + 1
            mEICGCCID = mEICGCCID + 1
            newYEICGCC0 = arrYEICGCC0(K)
            
            newYEICGCC0.EICGCCID = mEICGCCID

         
            cmdSelect_SQL_Import_ZGUIRC20
            If newYEICGCC0.EICGCCDOS = 0 Then cmdSelect_SQL_Import_ZCHGOPE0
            
            xSQL = "select CLIENACLI,COMPTEINT from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
                 & " Where COMPTECOM = '" & newYEICGCC0.EICGCCXCPT & "'"
            Set rsSab = cnsab.Execute(xSQL)
            If Not rsSab.EOF Then
                newYEICGCC0.EICGCCXID = rsSab("CLIENACLI")
                newYEICGCC0.EICGCCXNOM = rsSab("COMPTEINT")
                newYEICGCC0.EICGCCKLAB = "I"
            End If
            
            newYEICGCC0.EICGCCXBQ = strSocBdfE
            
            
            arrYEICGCC0(K) = newYEICGCC0

       End If
    Next K
    
    Call cmdYEICGCC0_Update("Import")
End If


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
Exit_sub:
Call lstErr_AddItem(lstErr, cmdContext, "- Chèques circulants ajoutés : " & Nb_REM): DoEvents
'++++++++++++++++++++++++++++++++++++++++++
newYEICGCCLOG.EICGCCLOGK = "Import V"
newYEICGCCLOG.EICGCCLOGI = 0
newYEICGCCLOG.EICGCCLOGE = 0
newYEICGCCLOG.EICGCCLOGA = ""
If arrYEICGCC0_Nb = 0 Then
    newYEICGCCLOG.EICGCCLOGX = "pas de chèques circulants trouvés"
    cmdYEICGCCLOG_New
Else
    newYEICGCCLOG.EICGCCLOGX = "nb RI0/REM rapprochés : " & Nb_RI0
    cmdYEICGCCLOG_New
    newYEICGCCLOG.EICGCCLOGX = "nb XXX ajoutés : " & Nb_REM & " (ID : " & mEICGCCID_0 & "-" & mEICGCCID & ")"
    cmdYEICGCCLOG_New
End If
'++++++++++++++++++++++++++++++++++++++++++

cnAdo_CHQ_ARCHIVE.Close
Set cnAdo_CHQ_ARCHIVE = Nothing
cboSelect_SQL.ListIndex = 0

End Sub
Public Sub cmdSelect_SQL_Import_EICGCCID()
Dim xSQL As String
'_________________________________________________________________________________
xSQL = "select count(*) as Tally    from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 "
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    mEICGCCID = 0
Else
    mEICGCCID = rsSab("Tally")
    
    xSQL = "select EICGCCID from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
     & "  Where EICGCCID >= " & mEICGCCID & " order by EICGCCID desc"
    Set rsSab = cnsab.Execute(xSQL)

    If Not rsSab.EOF Then
        mEICGCCID = rsSab("EICGCCID")
    End If
End If
mEICGCCID_0 = mEICGCCID + 1
End Sub

Private Sub txtSelect_EICGCCXNOM_LostFocus()
txt_LostFocus txtSelect_EICGCCXNOM

End Sub



Public Sub cmdSelect_SQL_Import_ZGUIRC20()
Dim V
Dim x As String, blnOk As Boolean, K As Long
Dim xSQL As String
Dim Nb As Integer
Dim wEICGCCAMJ As Long, wEICGCCSER As String, wEICGCCSSE As String
Dim wGUIRC2MT1 As Currency, wGUIRC2CPT As String
Dim wGUIRC2DOS As Long, wGUIRC2CTA As Integer
On Error GoTo Error_Handler

Nb = 0
blnOk = False
newYEICGCCLOG.EICGCCLOGI = 0
newYEICGCCLOG.EICGCCLOGE = 0
newYEICGCCLOG.EICGCCLOGA = ""

xSQL = "select * from " & paramIBM_Library_SAB & ".ZGUIRC20 " _
     & " Where GUIRC2CHQ = " & Val(newYEICGCC0.EICGCCECHQ) & " and GUIRC2OPE ='" & newYEICGCC0.EICGCCOPE & "'" _
     & " and GUIRC2CTA = 3 order by GUIRC2DOS"
Set rsSab = cnsab.Execute(xSQL)
    
Do While Not rsSab.EOF
    Nb = Nb + 1
    blnOk = True
    wGUIRC2MT1 = rsSab("GUIRC2MT1")
    wGUIRC2CPT = Trim(rsSab("GUIRC2CPT"))
    wGUIRC2DOS = rsSab("GUIRC2DOS")
    wEICGCCAMJ = rsSab("GUIRC2DT3")
    wEICGCCSER = rsSab("GUIRC2SER")
    wEICGCCSSE = rsSab("GUIRC2SSE")

    If wGUIRC2MT1 <> newYEICGCC0.EICGCCEMT Then
        blnOk = False
        '++++++++++++++++++++++++++++++++++++++++++
        newYEICGCCLOG.EICGCCLOGK = "Import V"
        newYEICGCCLOG.EICGCCLOGX = "# montant REM/scan " & wGUIRC2MT1 & " / " & newYEICGCC0.EICGCCEMT & " cpt " & Trim(newYEICGCC0.EICGCCECPT) & " chq " & newYEICGCC0.EICGCCECHQ
        cmdYEICGCCLOG_New
        '++++++++++++++++++++++++++++++++++++++++++
    End If
    If wGUIRC2CPT <> Trim(newYEICGCC0.EICGCCECPT) Then
        blnOk = False
        '++++++++++++++++++++++++++++++++++++++++++
        newYEICGCCLOG.EICGCCLOGK = "Import V"
        newYEICGCCLOG.EICGCCLOGX = "# compte REM/scan " & wGUIRC2CPT & " / " & Trim(newYEICGCC0.EICGCCECPT) & " chq " & newYEICGCC0.EICGCCECHQ
        cmdYEICGCCLOG_New
        '++++++++++++++++++++++++++++++++++++++++++
    End If
                   
    rsSab.MoveNext
Loop

If Nb > 1 Then
        '++++++++++++++++++++++++++++++++++++++++++
        newYEICGCCLOG.EICGCCLOGK = "Import V"
        newYEICGCCLOG.EICGCCLOGX = "# ZGUIRC20 : " & Nb & " chèques même N° " & newYEICGCC0.EICGCCECHQ & " , cpt " & newYEICGCC0.EICGCCECPT
        cmdYEICGCCLOG_New
        '++++++++++++++++++++++++++++++++++++++++++
End If

If blnOk Then
    newYEICGCC0.EICGCCAMJ = wEICGCCAMJ + 19000000
    newYEICGCC0.EICGCCSER = wEICGCCSER
    newYEICGCC0.EICGCCSSE = wEICGCCSSE
    newYEICGCC0.EICGCCDOS = wGUIRC2DOS
    cmdSelect_SQL_Import_ZGUIRC10
End If

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & "cmdSelect_SQL_Import_ZGUIRC20"
Exit_sub:

End Sub

Public Sub cmdSelect_SQL_Import_ZCHGOPE0()
Dim V
Dim x As String, blnOk As Boolean, K As Long
Dim xSQL As String
Dim Nb As Integer
Dim wEICGCCAMJ As Long, wEICGCCSER As String, wEICGCCSSE As String
Dim wCHGOPEMO1 As Currency
Dim wCHGOPEDOS As Long, wCHGOPECTA As Integer
On Error GoTo Error_Handler

If Val(newYEICGCC0.EICGCCECHQ) = 0 Then Exit Sub

Nb = 0
blnOk = False
newYEICGCCLOG.EICGCCLOGI = 0
newYEICGCCLOG.EICGCCLOGE = 0
newYEICGCCLOG.EICGCCLOGA = ""

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCHGOPE0 " _
     & " Where CHGOPECHQ = " & Val(newYEICGCC0.EICGCCECHQ) & " and CHGOPEOPE ='TRF'" _
     & " and CHGOPEANN = ' ' order by CHGOPEDOS"
Set rsSab = cnsab.Execute(xSQL)
    
Do While Not rsSab.EOF
    Nb = Nb + 1
    blnOk = True
    wCHGOPEMO1 = rsSab("CHGOPEMO1")
    wCHGOPEDOS = rsSab("CHGOPEDOS")
    wEICGCCAMJ = rsSab("CHGOPEDT1")
    wEICGCCSER = rsSab("CHGOPESER")
    wEICGCCSSE = rsSab("CHGOPESSE")

    If wCHGOPEMO1 <> newYEICGCC0.EICGCCEMT Then
        blnOk = False
        '++++++++++++++++++++++++++++++++++++++++++
        newYEICGCCLOG.EICGCCLOGK = "Import TRF"
        newYEICGCCLOG.EICGCCLOGX = "# montant TRF/scan " & wCHGOPEMO1 & " / " & newYEICGCC0.EICGCCEMT & " cpt " & Trim(newYEICGCC0.EICGCCECPT) & " chq " & newYEICGCC0.EICGCCECHQ
        cmdYEICGCCLOG_New
        '++++++++++++++++++++++++++++++++++++++++++
    End If
                   
    rsSab.MoveNext
Loop

If Nb > 1 Then
        '++++++++++++++++++++++++++++++++++++++++++
        newYEICGCCLOG.EICGCCLOGK = "Import TRF"
        newYEICGCCLOG.EICGCCLOGX = "# ZCHGOPE0 : " & Nb & " chèques même N° " & newYEICGCC0.EICGCCECHQ & " , cpt " & newYEICGCC0.EICGCCECPT
        cmdYEICGCCLOG_New
        '++++++++++++++++++++++++++++++++++++++++++
End If

If blnOk Then
    newYEICGCC0.EICGCCAMJ = wEICGCCAMJ + 19000000
    newYEICGCC0.EICGCCSER = wEICGCCSER
    newYEICGCC0.EICGCCSSE = wEICGCCSSE
    newYEICGCC0.EICGCCOPE = "TRF"
    newYEICGCC0.EICGCCDOS = wCHGOPEDOS
    cmdSelect_SQL_Import_ZCHGDET0
End If

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & "cmdSelect_SQL_Import_ZCHGOPE0"
Exit_sub:

End Sub

Public Sub cmdSelect_SQL_Import_ZGUIRC10()
Dim V
Dim x As String, blnOk As Boolean, K As Long
Dim xSQL As String
On Error GoTo Error_Handler


xSQL = "select * from " & paramIBM_Library_SAB & ".ZGUIRC10 " _
     & " Where GUIRC1DOS = " & newYEICGCC0.EICGCCDOS & " and GUIRC1OPE ='" & newYEICGCC0.EICGCCOPE & "'" _
     & " and GUIRC1SER = '" & newYEICGCC0.EICGCCSER & "' and GUIRC1SSE ='" & newYEICGCC0.EICGCCSSE & "'"
Set rsSab = cnsab.Execute(xSQL)
    
If rsSab.EOF Then

    '++++++++++++++++++++++++++++++++++++++++++
    newYEICGCCLOG.EICGCCLOGK = "Import REM"
    newYEICGCCLOG.EICGCCLOGX = "# dossier inconnu ZGUIRC10 : " & newYEICGCC0.EICGCCOPE & " " & newYEICGCC0.EICGCCDOS
    cmdYEICGCCLOG_New
    '++++++++++++++++++++++++++++++++++++++++++
Else
        x = rsSab("GUIRC1CP2")
    If x <> newYEICGCC0.EICGCCXCPT Then
        blnOk = False
        '++++++++++++++++++++++++++++++++++++++++++
        newYEICGCCLOG.EICGCCLOGK = "Import REM"
        newYEICGCCLOG.EICGCCLOGX = "! Compte à créditer REM/scan " & x & " / " & newYEICGCC0.EICGCCXCPT & " chq " & newYEICGCC0.EICGCCECHQ
        cmdYEICGCCLOG_New
        '++++++++++++++++++++++++++++++++++++++++++
    End If
    newYEICGCC0.EICGCCXCPT = x
End If
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & "cmdSelect_SQL_Import_ZGUIRC10"
Exit_sub:

End Sub
Public Sub cmdSelect_SQL_Import_ZCHGDET0()
Dim V
Dim x As String, blnOk As Boolean, K As Long
Dim wCHGDETSEN As String
Dim xSQL As String
On Error GoTo Error_Handler


xSQL = "select * from " & paramIBM_Library_SAB & ".ZCHGDET0 " _
     & " Where CHGDETDOS = " & newYEICGCC0.EICGCCDOS & " and CHGDETOPE ='" & newYEICGCC0.EICGCCOPE & "'" _
     & " and CHGDETSER = '" & newYEICGCC0.EICGCCSER & "' and CHGDETSSE ='" & newYEICGCC0.EICGCCSSE & "'"
Set rsSab = cnsab.Execute(xSQL)
    
If rsSab.EOF Then
    '++++++++++++++++++++++++++++++++++++++++++
    newYEICGCCLOG.EICGCCLOGK = "Import V"
    newYEICGCCLOG.EICGCCLOGX = "# dossier inconnu ZCHGDET0 : " & newYEICGCC0.EICGCCOPE & " " & newYEICGCC0.EICGCCDOS
    cmdYEICGCCLOG_New
    '++++++++++++++++++++++++++++++++++++++++++
End If
Do While Not rsSab.EOF
    wCHGDETSEN = rsSab("CHGDETSEN")
    x = Trim(rsSab("CHGDETCP1"))
    If wCHGDETSEN = "C" Then
        If x <> Trim(newYEICGCC0.EICGCCXCPT) Then
            blnOk = False
            '++++++++++++++++++++++++++++++++++++++++++
            newYEICGCCLOG.EICGCCLOGK = "Import V"
            newYEICGCCLOG.EICGCCLOGX = "! Compte à créditer TRF/scan " & x & " / " & newYEICGCC0.EICGCCXCPT & " chq " & newYEICGCC0.EICGCCECHQ
            cmdYEICGCCLOG_New
            '++++++++++++++++++++++++++++++++++++++++++
            newYEICGCC0.EICGCCXCPT = x
        End If
    Else
        If x <> Trim(newYEICGCC0.EICGCCECPT) Then
            blnOk = False
            '++++++++++++++++++++++++++++++++++++++++++
            newYEICGCCLOG.EICGCCLOGK = "Import V"
            newYEICGCCLOG.EICGCCLOGX = "# compte à débiter TRF/scan " & x & " / " & Trim(newYEICGCC0.EICGCCECPT) & " chq " & newYEICGCC0.EICGCCECHQ
            cmdYEICGCCLOG_New
            '++++++++++++++++++++++++++++++++++++++++++
             newYEICGCC0.EICGCCECPT = x
       End If
    End If

    rsSab.MoveNext
Loop

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & "cmdSelect_SQL_Import_ZCHGDET0"
Exit_sub:

End Sub



Public Function cmdSelect_SQL_Import_EICGCCECPT(lEICGCCECPT As String) As String
Dim xSQL As String
cmdSelect_SQL_Import_EICGCCECPT = ""
xSQL = "select CLIENACLI from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " Where COMPTECOM = '" & lEICGCCECPT & "'"
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    cmdSelect_SQL_Import_EICGCCECPT = rsSab("CLIENACLI")
Else
    xSQL = "select COMREFCOM from " & paramIBM_Library_SAB & ".ZCOMREF0 " _
         & " Where COMREFREF = '" & lEICGCCECPT & "' and COMREFCOR = 'SI' " _
         & " and COMREFETA = 1 and COMREFPLA = 1"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        lEICGCCECPT = rsSab("COMREFCOM")
         xSQL = "select CLIENACLI from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
             & " Where COMPTECOM = '" & lEICGCCECPT & "'"
        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            cmdSelect_SQL_Import_EICGCCECPT = rsSab("CLIENACLI")
        End If
    End If
       
End If


End Function

Public Sub lstW_SQL(C As Control)
Dim x As String, xSQL As String
x = Trim(C)
If x <> "" Then
    xSQL = "select distinct " & mEICGCCXXX & "  from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
         & " where " & mEICGCCXXX & " like '" & x & "%'" _
         & " group by " & mEICGCCXXX & " order by " & mEICGCCXXX
    Set rsSab = cnsab.Execute(xSQL)
    
    lstW.Clear
    Do While Not rsSab.EOF
        
        lstW.AddItem Trim(rsSab(mEICGCCXXX))
        rsSab.MoveNext
    Loop
End If
If lstW.ListCount > 0 Then
    lstW.Visible = True
Else
    lstW.Visible = False
End If
End Sub

Public Sub cmdDetail_UpdateVO_Ok()
Dim V, x As String, K As Integer, xSQL As String
Dim blnAut As Boolean

Me.Enabled = False: Me.MousePointer = vbHourglass

K = InStr(lstW.Text, " ")
voYEICGCC0.EICGCCID = Val(Mid$(lstW.Text, 1, K))
xSQL = "select *    from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
     & " where EICGCCID= " & voYEICGCC0.EICGCCID
Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    Call MsgBox("Enregistrement non trouvé : " & voYEICGCC0.EICGCCID, vbCritical, "Associer une vignette orpheline")
    GoTo Exit_sub
End If
'______________________________________________________________________
V = rsYEICGCC0_GetBuffer(rsSab, voYEICGCC0)
blnAut = True
x = oldYEICGCC0.EICGCCID & " - " & voYEICGCC0.EICGCCID & vbCrLf
If oldYEICGCC0.EICGCCEMT <> voYEICGCC0.EICGCCEMT Then
    x = x & "- les montants ne sont pas identiques " & vbCrLf
    blnAut = YEICGCC0_Aut.Rapprocher
End If
If oldYEICGCC0.EICGCCECHQ <> voYEICGCC0.EICGCCECHQ Then
    x = x & "- les numéros de chèques ne sont pas identiques " & vbCrLf
    blnAut = YEICGCC0_Aut.Rapprocher
End If
If oldYEICGCC0.EICGCCECPT <> voYEICGCC0.EICGCCECPT Then
    x = x & "- les comptes ne sont pas identiques " & vbCrLf
End If

If blnAut Then
    x = MsgBox("confirmation de l'association : " & x, vbQuestion & vbYesNo, "Associer une vignette orpheline")
Else
    x = MsgBox("vous n'êtes pas autorisé à cette association : " & x, vbNo, "Associer une vignette orpheline")
End If

If x <> vbYes Then GoTo Exit_sub
'______________________________________________________________________

Call fraDetail_Control
newYEICGCC0 = xYEICGCC0
newYEICGCC0.EICGCCVAMJ = voYEICGCC0.EICGCCVAMJ
newYEICGCC0.EICGCCVJPG = voYEICGCC0.EICGCCVJPG
newYEICGCC0.EICGCCVREM = voYEICGCC0.EICGCCVREM
V = cmdYEICGCC0_Update("UpdateVO")


'++++++++++++++++++++++++++++++++++++++++++
If IsNull(V) Then
    rsYEICGCCLOG_Init newYEICGCCLOG
    newYEICGCCLOG.EICGCCLOGK = "Associer VO"
    newYEICGCCLOG.EICGCCLOGX = "lien " & oldYEICGCC0.EICGCCID & " - " & voYEICGCC0.EICGCCID & " : " _
                             & oldYEICGCC0.EICGCCOPE & " " & oldYEICGCC0.EICGCCDOS _
                             & " du " & dateImp10(oldYEICGCC0.EICGCCAMJ) & " " _
                              & " avec VO " & voYEICGCC0.EICGCCID & " du " & dateImp10(voYEICGCC0.EICGCCAMJ)
    newYEICGCCLOG.EICGCCLOGI = oldYEICGCC0.EICGCCID
    cmdYEICGCCLOG_New
End If
'++++++++++++++++++++++++++++++++++++++++++


fraDetail.Visible = False
fraCHQ_Max.Visible = False
fraCHQ.Visible = False

cmdSelect_Ok_Click
'______________________________________________________________________

Exit_sub:
lstW.Visible = False
Me.Enabled = True: Me.MousePointer = 0

End Sub


Public Sub cboLogV_K2_Load(lK1 As String)
Dim xSQL As String
cboLogV_K2.Clear
xSQL = "select BIATABK2 from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YEICGCC0' and BIATABK1 = '" & Trim(lK1) & "' order by BIATABK2"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF

    cboLogV_K2.AddItem rsSab("BIATABK2")
    rsSab.MoveNext
Loop
'If cboLogV_K2.ListCount > 0 Then cboLogV_K2.ListIndex = 0
End Sub

Public Function fraLogV_Control()
Dim x As String, wMsgBox As String
Dim K As Integer

fraLogV_Control = Null
wMsgBox = ""
blnDos_Log = False
'''''newYEICGCC0 = oldYEICGCC0

fraDetail_Control_EICGCCSTA_Init

newYEICGCCLOG.EICGCCLOGK = cboLogV_K
newYEICGCCLOG.EICGCCLOGX = txtLogV_X
If Trim(newYEICGCCLOG.EICGCCLOGX) = "" Then
    wMsgBox = " - préciser un commentaire pour cette action" & vbCrLf
End If
newYEICGCCLOG.EICGCCLOGI = oldYEICGCC0.EICGCCID
newYEICGCCLOG.EICGCCLOGA = ""

If IsNull(txtLogV_E) Then
    newYEICGCCLOG.EICGCCLOGE = 0
Else
    Call DTPicker_Control(txtLogV_E, x)
    newYEICGCCLOG.EICGCCLOGE = x
End If

Select Case Trim(newYEICGCCLOG.EICGCCLOGK)
    Case "Annulation": newYEICGCCLOG.EICGCCLOGE = 0
                       newYEICGCCLOG.EICGCCLOGA = "V"
                    If newYEICGCC0.EICGCCSTA <> " " Then
                        wMsgBox = " - annulation impossible, statut du dossier = " & oldYEICGCC0.EICGCCSTA & vbCrLf
                    Else
                        blnDos_Log = True
                        Call fraDetail_Control_EICGCCSTA
                        newYEICGCC0.EICGCCSTA = "A"
                    End If
                
    Case "Reprise/Ann": newYEICGCCLOG.EICGCCLOGE = 0
                        newYEICGCCLOG.EICGCCLOGA = "V"
                    If newYEICGCC0.EICGCCSTA <> "A" Then
                        wMsgBox = " - reprise impossible, statut du dossier = " & oldYEICGCC0.EICGCCSTA & vbCrLf
                     Else
                        blnDos_Log = True
                        Call fraDetail_Control_EICGCCSTA
                        newYEICGCC0.EICGCCSTA = " "
                   End If
    Case "Mail DCOM": '''newYEICGCCLOG.EICGCCLOGE = 0
                      newYEICGCCLOG.EICGCCLOGA = " " '"V"
                      sqlCLIENARES_Mail
                      If mMNUUTIMAI = "" Then
                            wMsgBox = wMsgBox & " - préciser l'adesse mail du responsable" & vbCrLf
                       End If
                        Call fraDetail_Control_EICGCCSTA
    Case "Révision": newYEICGCCLOG.EICGCCLOGE = 0
                      newYEICGCCLOG.EICGCCLOGA = "V"
                     newYEICGCC0.EICGCCSTA = " "
    Case "CHQ accepté": newYEICGCCLOG.EICGCCLOGE = 0
                    If newYEICGCC0.EICGCCSTAK = "V" Then
                        wMsgBox = " - impossible, statut des contrôles du dossier = " & oldYEICGCC0.EICGCCSTAK & vbCrLf
                    Else
                        blnDos_Log = True
                        Call fraDetail_Control_EICGCCSTA
                        newYEICGCC0.EICGCCSTAK = "!"
                    End If
    Case "CHQ rejeté": newYEICGCCLOG.EICGCCLOGE = YBIATAB0_DATE_CPT_JS1
                        Call fraDetail_Control_EICGCCSTA
                        newYEICGCC0.EICGCCSTA = "R"
    Case "non circulan": newYEICGCCLOG.EICGCCLOGE = 0
                         newYEICGCCLOG.EICGCCLOGA = "V"
                         newYEICGCC0.EICGCCSTA = "@"
    Case "CHQ àIgnorer": newYEICGCCLOG.EICGCCLOGE = 0
                         newYEICGCCLOG.EICGCCLOGA = "V"
                         newYEICGCC0.EICGCCSTA = "I"
End Select


If wMsgBox <> "" Then
    fraLogV_Control = "?"
    Call MsgBox(wMsgBox, vbCritical, "EIC_GCC : contrôle des événements")
End If

End Function

Private Sub txtSelect_Options_PCI_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtSuivi_Q_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Public Function fraSuivi_Control()
Dim x As String, xWhere As String, xSQL As String, wMsgBox As String
Dim Nb As Long
Dim wEICGCCECHQ As String
Debut:

wMsgBox = ""
fraSuivi_Control = Null
Nb = 0
rsYEICGCCLOG_Init newYEICGCCLOG

wEICGCCECHQ = Format(Val(txtSuivi_Q), "0000000")
xWhere = " where EICGCCECHQ ='" & wEICGCCECHQ & "' and EICGCCSTA <> 'A'"
xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYEICGCC0_GetBuffer(rsSab, oldYEICGCC0)
    Nb = Nb + 1
    rsSab.MoveNext
Loop

If Nb = 0 Then
    wMsgBox = "pas de dossier en cours avec ce numéro de chèque." & vbCrLf
    x = MsgBox(wMsgBox & vbCrLf _
               & " Voulez-vous créer un dossier EIC ?", vbYesNo, "EIC_GCC : saisie des événements")
    If x = vbYes Then
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZEICRIC0 " _
             & " Where EICRICCHE = '" & wEICGCCECHQ & "'" _
             & " and EICRICOPE = 'RI0' and EICRICOPR = '160'order by  EICRICNU1"
        
        Set rsSab = cnsab.Execute(xSQL)

        Call cmdSelect_SQL_Import_EIC_Update(False, Nb)
        If Nb = 1 Then
            GoTo Debut
        Else
            wMsgBox = "pas de dossier 'EIC RI0' avec ce numéro de chèque." & vbCrLf
       End If
    End If
Else
    If Nb > 1 Then
        wMsgBox = "plusieurs dossiers en cours avec le même numéro de chèque"
    Else
        newYEICGCCLOG.EICGCCLOGK = cboSuivi_K
        Select Case Trim(newYEICGCCLOG.EICGCCLOGK)
            Case "AF0":     newYEICGCCLOG.EICGCCLOGX = ":Demande de télécopie au remettant"
            Case "AI1":     newYEICGCCLOG.EICGCCLOGX = ":retour vignette au remettant"
            Case "RQ0":     newYEICGCCLOG.EICGCCLOGX = ":requête remettant"
        End Select
        newYEICGCCLOG.EICGCCLOGI = oldYEICGCC0.EICGCCID
        newYEICGCCLOG.EICGCCLOGA = ""
    
        newYEICGCCLOG.EICGCCLOGE = YBIATAB0_DATE_CPT_JS1 'dateElp("Ouvré", 5, DSys)
    End If
End If
If wMsgBox <> "" Then
    fraSuivi_Control = "?"
    Call MsgBox(wMsgBox, vbCritical, "EIC_GCC : saisie des événements")
End If

End Function

Public Sub fraDetail_Control_EICGCCSTA()

If newYEICGCC0.EICGCCSTAK <> "!" Then
    If newYEICGCC0.EICGCCEAMJ > 0 Then
        If newYEICGCC0.EICGCCKMT = "V" _
        And newYEICGCC0.EICGCCKSIG = "V" _
        And newYEICGCC0.EICGCCKEND = "V" Then
            If newYEICGCC0.EICGCCKLAB = "V" Or newYEICGCC0.EICGCCKLAB = "I" Then newYEICGCC0.EICGCCSTAK = "V"
        End If
        If newYEICGCC0.EICGCCKMT = "X" _
        Or newYEICGCC0.EICGCCKSIG = "X" _
        Or newYEICGCC0.EICGCCKEND = "X" _
        Or newYEICGCC0.EICGCCKLAB = "X" Then
            newYEICGCC0.EICGCCSTAK = "X"
        End If
    End If
End If

If newYEICGCC0.EICGCCSTAK = "V" Or newYEICGCC0.EICGCCSTAK = "!" Then
    If Trim(newYEICGCC0.EICGCCVINT) <> "" _
    And Trim(newYEICGCC0.EICGCCXNOM) <> "" _
    And Trim(newYEICGCC0.EICGCCXECO) <> "" Then
        newYEICGCC0.EICGCCSTA = "V"
    End If
End If

End Sub

Public Sub cmdDetail_Update_Ok(lFct As String)
Dim V, xSQL As String
Dim mEICGCCID As Long

newYEICGCC0 = oldYEICGCC0
mEICGCCID = oldYEICGCC0.EICGCCID

If IsNull(fraDetail_Control) Then

    Select Case lFct
        Case "#Log", "#Log+Val"
            Call fraDetail_Control_EICGCCSTA
            V = cmdYEICGCC0_Update("Dos" & lFct)
            
        Case "#Log#VO"
            If Trim(oldYEICGCCLOG.EICGCCLOGK) = "Associer VO" Then
                newYEICGCC0.EICGCCVAMJ = 0
                newYEICGCC0.EICGCCVJPG = 0
                newYEICGCC0.EICGCCVREM = 0
            End If
            Call fraDetail_Control_EICGCCSTA
            V = cmdYEICGCC0_Update("Dos" & lFct)
            
        Case Else
            If fraLogV.Visible Then
                If IsNull(fraLogV_Control) Then V = cmdYEICGCC0_Update("Dos+Log")
            Else
                V = cmdYEICGCC0_Update("Update")
            End If
    End Select
    
End If

If IsNull(V) Then

     If lFct = "+Log" And Trim(newYEICGCCLOG.EICGCCLOGK) = "Mail DCOM" Then Call cmdSendMail_DCOM(Trim(txtLogV_X))
     If lFct = "#Log+Val" And Trim(newYEICGCCLOG.EICGCCLOGK) = "Mail DCOM *" Then Call cmdSendMail_DCOM_Culot(Trim(txtDetail_EICGCCXECO))
     rsYEICGCCLOG_Init newYEICGCCLOG
     newYEICGCCLOG.EICGCCLOGK = ""
     blnControl = True
     cmdSelect_Reset
     'cmdSelect_Ok_Click
     'Call fraDetail_Display_EICGCCID(mEICGCCID)
End If

Exit_sub:

End Sub

Public Sub fraDetail_Display_EICGCCID(lEICGCCID As Long)
Dim V, xSQL As String
 xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
 & "  Where EICGCCID = " & lEICGCCID
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
   V = rsYEICGCC0_GetBuffer(rsSab, oldYEICGCC0)
   xYEICGCC0 = oldYEICGCC0
   fraDetail_Display
End If

End Sub

Public Sub fraDetail_Control_EICGCCSTA_Init()
Dim K As Integer

newYEICGCC0.EICGCCSTAK = " "
newYEICGCC0.EICGCCSTA = " "
For K = 1 To arrYEICGCCLOG_Nb
    If arrYEICGCCLOG(K).EICGCCLOGA = " " Or arrYEICGCCLOG(K).EICGCCLOGA = "V" Then
        Select Case Trim(arrYEICGCCLOG(K).EICGCCLOGK)
            Case "Annulation": newYEICGCC0.EICGCCSTA = "A"
            Case "Reprise/Ann": newYEICGCC0.EICGCCSTA = " "
            Case "Révision": newYEICGCC0.EICGCCSTA = " "
            Case "CHQ rejeté": newYEICGCC0.EICGCCSTA = "R"
            Case "CHQ accepté": newYEICGCC0.EICGCCSTAK = "!"
        End Select
    End If
Next K

End Sub

Public Sub sqlCLIENARES(lCLIENARES As String)
Dim xSQL As String, x As String
If mCLIENARES <> lCLIENARES Then
    mCLIENARES = lCLIENARES
    mMNURUTUTI = ""
    mMNUUTIMAI = ""
    mCLIENARES_Lib = ""
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 " _
         & " Where BASTABETA = 1 and BASTABNUM = 6 and BASTABARG = 'CLI" & lCLIENARES & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        x = rsSab("BASTABLO2") & rsSab("BASTABDON")
        mCLIENARES_Lib = Mid$(x, 1, 30)
        mMNURUTUTI = Mid$(x, 36, 10)
    End If
End If
End Sub
Public Sub sqlCLIENARES_Mail()
Dim xSQL As String, x As String
If mMNUUTIMAI = "" Then
    xSQL = "select MNUUTIMAI from " & paramIBM_Library_SAB & ".ZMNURUT0 , " _
         & paramIBM_Library_SAB & ".ZMNUUTI0 " _
         & " Where MNURUTUTI = '" & mMNURUTUTI & "' and MNURUTCUT = MNUUTICUT"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        mMNUUTIMAI = rsSab("MNUUTIMAI")
    End If
End If
If mMNUUTIMAI = "" Then
    x = InputBox("Le paramétrage de ce responsable ne permet pas de rechercher son adresse 'mail'" & vbCrLf _
                & "saisir son adresse (<nom> <point> <initiales du prénom>)", "SAISIE")
Else
    x = InputBox("Confirmer l'adresse 'mail'du destinataire, SINON" & vbCrLf _
                & "saisir son adresse (<nom> <point> <initiales du prénom>)", "CONFIRMATION", mMNUUTIMAI)
End If
If x <> "" Then
    If InStr(x, "@") > 0 Then
        mMNUUTIMAI = Trim(x)
    Else
        mMNUUTIMAI = Trim(x) & "@bia-paris.fr"
    End If
End If
End Sub



Public Sub parametrage_Reprise()

New_YBIATAB0.BIATABID = "YEICGCC0"
New_YBIATAB0.BIATABK1 = "Action"

New_YBIATAB0.BIATABK2 = "Annulation"
New_YBIATAB0.BIATABTXT = "Annulation du dossier, 'Reprise/ann' pour le réactiver"
Call Parametrage_New

New_YBIATAB0.BIATABK2 = "Reprise/Ann"
New_YBIATAB0.BIATABTXT = "Réactivation du dossier"
Call Parametrage_New

New_YBIATAB0.BIATABK2 = "Révision"
New_YBIATAB0.BIATABTXT = "Restauration du statut du dossier sans vérification"
Call Parametrage_New

New_YBIATAB0.BIATABK2 = "Mail DCOM"
New_YBIATAB0.BIATABTXT = "Envoi d'un mail au responsable commercial + copie GDMP"
Call Parametrage_New

New_YBIATAB0.BIATABK2 = "CHQ rejeté"
New_YBIATAB0.BIATABTXT = "Rejet d'un chèque"
Call Parametrage_New

New_YBIATAB0.BIATABK2 = "CHQ accepté"
New_YBIATAB0.BIATABTXT = "le chèque est accepté, même non conforme"
Call Parametrage_New

New_YBIATAB0.BIATABK2 = "AOCT"
New_YBIATAB0.BIATABTXT = "opération sur chèque"
Call Parametrage_New
'__________________________________________________________________________________________

New_YBIATAB0.BIATABK2 = "Motif Eco"
New_YBIATAB0.BIATABTXT = "liste des motifs économiques"
Call Parametrage_New

New_YBIATAB0.BIATABK1 = "Motif Eco"

New_YBIATAB0.BIATABK2 = "salaire"
New_YBIATAB0.BIATABTXT = "salaire"
Call Parametrage_New
New_YBIATAB0.BIATABK2 = "honoraires"
New_YBIATAB0.BIATABTXT = "honoraires"
Call Parametrage_New
New_YBIATAB0.BIATABK2 = "impôts"
New_YBIATAB0.BIATABTXT = "impôts"
Call Parametrage_New

'__________________________________________________________________________________________

New_YBIATAB0.BIATABK1 = "Annulation"

New_YBIATAB0.BIATABK2 = "erreur SIT"
New_YBIATAB0.BIATABTXT = "erreur SIT"
Call Parametrage_New
New_YBIATAB0.BIATABK2 = "doublon"
New_YBIATAB0.BIATABTXT = "doublon"
Call Parametrage_New
'__________________________________________________________________________________________

New_YBIATAB0.BIATABK1 = "Révision"

New_YBIATAB0.BIATABK2 = "erreur saisie"
New_YBIATAB0.BIATABTXT = "erreur saisie"
Call Parametrage_New

End Sub

Public Sub cmdSelect_SQL_Echeancier()
Dim V
Dim x As String
Dim xSQL As String

ReDim arrYEICGCC0(101)
ReDim arrYEICGCCLOG(101)
arrYEICGCC0_Max = 100: arrYEICGCC0_Nb = 0

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_Echeancier"

xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YEICGCCLOG , " _
     & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 " _
     & " Where  EICGCCLOGE > 0 and EICGCCLOGA = ' '  and EICGCCLOGI= EICGCCID"
     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    V = rsYEICGCC0_GetBuffer(rsSab, xYEICGCC0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "fgSelect_Display_Echeancier"
        '' Exit Sub
     Else
         arrYEICGCC0_Nb = arrYEICGCC0_Nb + 1
         If arrYEICGCC0_Nb > arrYEICGCC0_Max Then
             arrYEICGCC0_Max = arrYEICGCC0_Max + 100
             ReDim Preserve arrYEICGCC0(arrYEICGCC0_Max)
             ReDim Preserve arrYEICGCCLOG(arrYEICGCC0_Max)
         End If
         
       arrYEICGCC0(arrYEICGCC0_Nb) = xYEICGCC0
       V = rsYEICGCCLOG_GetBuffer(rsSab, xYEICGCCLOG)
       arrYEICGCCLOG(arrYEICGCC0_Nb) = xYEICGCCLOG
    
    End If
    rsSab.MoveNext
Loop

fgSelect_Display_Echeancier

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Public Function cmdSelect_SQL_Import_EIC_560_New(lEICAO2OPR As String, lEICAO2NUR As Long)
Dim xSQL As String, Nb As Long

cmdSelect_SQL_Import_EIC_560_New = Null
 xSQL = "select * from " & paramIBM_Library_SAB & ".ZEICRIC0 " _
      & " Where EICRICOPE = '" & lEICAO2OPR & "'" _
      & " and EICRICNU1 =" & lEICAO2NUR
 
 Set rsSab = cnsab.Execute(xSQL)

 Call cmdSelect_SQL_Import_EIC_Update(False, Nb)
 If Nb = 0 Then
    cmdSelect_SQL_Import_EIC_560_New = "?"
    newYEICGCCLOG.EICGCCLOGK = "AF0 ?"
    newYEICGCCLOG.EICGCCLOGI = 0
    newYEICGCCLOG.EICGCCLOGE = 0
    newYEICGCCLOG.EICGCCLOGA = ""
    newYEICGCCLOG.EICGCCLOGX = ": ZEICRIC0 inconnu SAB : " & lEICAO2OPR & "  " & lEICAO2NUR
    cmdYEICGCCLOG_New
 Else
    newYEICGCCLOG.EICGCCLOGK = "AF0"
    newYEICGCCLOG.EICGCCLOGI = oldYEICGCC0.EICGCCID
    newYEICGCCLOG.EICGCCLOGA = ""
    newYEICGCCLOG.EICGCCLOGX = ": demande de télécopie : " & oldYEICGCC0.EICGCCECHQ & " / SAB : " & lEICAO2OPR & "  " & lEICAO2NUR & " / Id: " & oldYEICGCC0.EICGCCID
    newYEICGCCLOG.EICGCCLOGE = YBIATAB0_DATE_CPT_JS1
    cmdYEICGCCLOG_New
End If
End Function

Public Function cmdSelect_SQL_Import_EIC_AF0_Ok()
Dim xSQL As String, K As Long, wAmj As Long
Dim V
Dim x As String

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_Echeancier"

xSQL = ".YEICGCCLOG where EICGCCLOGK = 'AF0' and EICGCCLOGA = ' ' and  EICGCCLOGE > 0 "
arrYEICGCCLOG_SQL xSQL

For K = 1 To arrYEICGCCLOG_Nb

    oldYEICGCCLOG = arrYEICGCCLOG(K)
    
    'xSql = "select EICGCCOPE, EICGCCDOS from " & paramIBM_Library_SABSPE_xxx & ".YEICGCC0" _
    '    & " where EICGCCID = " & oldYEICGCCLOG.EICGCCLOGI
    xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0" _
        & " where EICGCCID = " & oldYEICGCCLOG.EICGCCLOGI
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        V = rsYEICGCC0_GetBuffer(rsSab, oldYEICGCC0)

        xSQL = "select * from " & paramIBM_Library_SAB & ".ZEICRIC0 " _
            & " where EICRICOPE = '" & oldYEICGCC0.EICGCCOPE & "'" _
            & " and EICRICNU1 =" & oldYEICGCC0.EICGCCDOS
'__________________________________________________________
        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            wAmj = rsSab("EICRICTE1")
            If wAmj > 0 Then
                updYEICGCCLOG = oldYEICGCCLOG
                updYEICGCCLOG.EICGCCLOGA = "V"
                
                newYEICGCCLOG.EICGCCLOGK = "AF0 *"
                newYEICGCCLOG.EICGCCLOGI = oldYEICGCCLOG.EICGCCLOGI
                newYEICGCCLOG.EICGCCLOGA = ""
                newYEICGCCLOG.EICGCCLOGX = ": télécopie reçue le " & dateImp10(wAmj + 19000000)
                newYEICGCCLOG.EICGCCLOGE = 0
                If oldYEICGCC0.EICGCCEIND = "0" Then
                    newYEICGCC0 = oldYEICGCC0
                    newYEICGCC0.EICGCCSTA = "A"
                    Call cmdYEICGCC0_Update("Dos#Log+Val")
                Else
                    Call cmdYEICGCC0_Update("#Log+Val")
                End If
            End If
        End If
    End If
Next K

Exit Function

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Function

Public Sub cmdPrint_YEICGCC0(x As String)
prtYEICGCC0_Init "YEICGCC0", x
prtYEICGCC0_Open
For I = 1 To arrYEICGCC0_Nb
    prtYEICGCC0_Line arrYEICGCC0(I)
Next I
prtYEICGCC0_Close True

End Sub
Public Sub cmdPrint_YEICGCC0_Echéancier(x As String)
prtYEICGCC0_Init "Echéancier", x
prtYEICGCC0_Open
For I = 1 To arrYEICGCC0_Nb
    prtYEICGCC0_Line_Echéancier arrYEICGCC0(I), arrYEICGCCLOG(I)
Next I
prtYEICGCC0_Close True

End Sub

Public Sub cmdPrint_YEICGCC0_Statistiques(x As String)
prtYEICGCC0_Init "Statistiques", x & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax)
prtYEICGCC0_Open
prtYEICGCC0_Line_Statistiques fgSelect
'For I = 1 To arrYEICGCC0_Nb
'    prtYEICGCC0_Line_Echéancier arrYEICGCC0(I), arrYEICGCCLOG(I)
'Next I
prtYEICGCC0_Close True

End Sub

Public Sub cmdPrint_YEICGCCLOG(x As String)
prtYEICGCC0_Init "YEICGCCLOG", x
prtYEICGCC0_Open
For I = 1 To arrYEICGCCLOG_Nb
    prtYEICGCCLOG_Line arrYEICGCCLOG(I)
Next I
prtYEICGCC0_Close True

End Sub


Public Function cmdPrint_Title() As String
Dim x As String
cmdPrint_Title = ""
Select Case cmdSelect_SQL_K
    Case "1": x = "Selection (filtre)"
    Case "2e": x = "EIC | REM en attente"
    Case "2v": x = "vignettes à contrôler"
    Case "2?": x = "vignettes orphelines"
    Case "E": x = "Echéancier"
    Case "L#": x = "Liste des événements du"
    Case "St": x = "Comptage pour la période du "
End Select
cmdPrint_Title = x
End Function

Public Sub fraJEICGCCLOG_Display()
Dim xSQL As String
If cmdSelect_SQL_K = "J#" Then
    xSQL = "select * from " & paramIBM_Library_SABJRN & ".JRNENT0 " _
         & " where jorcv = " & oldYEICGCCLOG.JORCV _
         & " and joSEQN = " & oldYEICGCCLOG.JOSEQN
         
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        V = srvJRNENT0_GetBuffer_ODBC(rsSab, xJRNENT0)
        If IsNull(V) Then
            xJRNENT0.JOUSER = oldYEICGCCLOG.EICGCCLOGU
            Call srvJRNENT0_fgX(xJRNENT0, fgJRNENT0)
            fraJRNENT0.Caption = JOENTT_Lib(xJRNENT0.JOENTT)
            fraJRNENT0.ForeColor = vbRed
            fraJRNENT0.Visible = True
            txtYEICGCCLOGD = oldYEICGCCLOG.EICGCCLOGD
            txtYEICGCCLOGH = oldYEICGCCLOG.EICGCCLOGH
            txtYEICGCCLOGU = oldYEICGCCLOG.EICGCCLOGU
            txtYEICGCCLOGS = oldYEICGCCLOG.EICGCCLOGS
            txtYEICGCCLOGK = oldYEICGCCLOG.EICGCCLOGK
            txtYEICGCCLOGI = oldYEICGCCLOG.EICGCCLOGI
            txtYEICGCCLOGA = oldYEICGCCLOG.EICGCCLOGA
            txtYEICGCCLOGE = oldYEICGCCLOG.EICGCCLOGE
            txtYEICGCCLOGX = oldYEICGCCLOG.EICGCCLOGX
            
            fraYEICGCCLOG.Visible = True
        End If
    End If
End If
End Sub


Public Sub lstParam_Action_Load()
Dim xSQL As String

lstParam_Action_ListIndex = lstParam_Action.ListIndex
cmdParam_Quit_Click
Action_YBIATAB0.BIATABID = "YEICGCC0"
Action_YBIATAB0.BIATABK1 = "Action"
Action_YBIATAB0.BIATABK2 = lstParam_Action
libParam_Action = lstParam_Action

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YEICGCC0' and BIATABK1 = 'Action'  and BIATABK2 = '" & Trim(lstParam_Action) & "'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    Action_YBIATAB0.BIATABTXT = Trim(rsSab("BIATABTXT"))
    txtParam_Action = Action_YBIATAB0.BIATABTXT
End If

Old_YBIATAB0.BIATABID = "YEICGCC0"
Old_YBIATAB0.BIATABK1 = Action_YBIATAB0.BIATABK2

lstParam_K.Clear
xSQL = "select BIATABK2 from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YEICGCC0' and BIATABK1 = '" & Trim(lstParam_Action) & "' order by BIATABK2"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF

    lstParam_K.AddItem rsSab("BIATABK2")
    rsSab.MoveNext
Loop
If lstParam_K.ListCount > 0 Then lstParam_K.ListIndex = 0
cmdParam_Quit.Visible = True
cmdParam_Add.Visible = YEICGCC0_Aut.Rapprocher

fraParam_K.Enabled = True
End Sub

Public Sub cmdSelect_SQL_Stock()

Dim Nb As Long
Dim xSQL As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Nature                                                  |>                Nombre"
fgSelect.Row = 0

currentAction = "cmdSelect_SQL_Stock"
    
'____________________________________________________________________________________
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1

xSQL = "select count(*) as Tally from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0" _
     & " where EICGCCSTA = ' ' and EICGCCVJPG = 0"
Set rsSab = cnsab.Execute(xSQL)

Nb = rsSab("Tally")
fgSelect.Col = 0: fgSelect.Text = "Chèques en attente"
fgSelect.Col = 1: fgSelect.Text = Format$(Nb, "### ##0")
'____________________________________________________________________________________
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1

xSQL = "select count(*) as Tally from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0" _
     & " where EICGCCSTA = ' ' and  EICGCCVJPG > 0 and EICGCCDOS > 0"
Set rsSab = cnsab.Execute(xSQL)

Nb = rsSab("Tally")
fgSelect.Col = 0: fgSelect.Text = "Vignettes à contrôler"
fgSelect.Col = 1: fgSelect.Text = Format$(Nb, "### ##0")
'____________________________________________________________________________________
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1

xSQL = "select count(*) as Tally from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0" _
       & " where EICGCCSTA = ' ' and  EICGCCVJPG > 0 and EICGCCDOS = 0"
Set rsSab = cnsab.Execute(xSQL)

Nb = rsSab("Tally")
fgSelect.Col = 0: fgSelect.Text = "Vignettes orphelines"
fgSelect.Col = 1: fgSelect.Text = Format$(Nb, "### ##0")
'____________________________________________________________________________________
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1

xSQL = "select count(*) as Tally from " & paramIBM_Library_SABSPE_XXX & ".YEICGCCLOG" _
       & " where EICGCCLOGK = 'AF0' and EICGCCLOGA = ' ' and  EICGCCLOGE > 0"
Set rsSab = cnsab.Execute(xSQL)

Nb = rsSab("Tally")
fgSelect.Col = 0: fgSelect.Text = "Télécopie en attente"
fgSelect.Col = 1: fgSelect.Text = Format$(Nb, "### ##0")
'____________________________________________________________________________________
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1

xSQL = "select count(*) as Tally from " & paramIBM_Library_SABSPE_XXX & ".YEICGCCLOG" _
       & " where EICGCCLOGK = 'AI1' and EICGCCLOGA = ' ' and  EICGCCLOGE > 0"
Set rsSab = cnsab.Execute(xSQL)

Nb = rsSab("Tally")
fgSelect.Col = 0: fgSelect.Text = "Chèques rejetés à renvoyer"
fgSelect.Col = 1: fgSelect.Text = Format$(Nb, "### ##0")

'____________________________________________________________________________________
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1

xSQL = "select count(*) as Tally from " & paramIBM_Library_SABSPE_XXX & ".YEICGCCLOG" _
       & " where EICGCCLOGK = 'Mail DCOM' and EICGCCLOGA = ' ' and  EICGCCLOGE > 0"
Set rsSab = cnsab.Execute(xSQL)

Nb = rsSab("Tally")
fgSelect.Col = 0: fgSelect.Text = "Réponses DCOM en attente"
fgSelect.Col = 1: fgSelect.Text = Format$(Nb, "### ##0")

fgSelect.Visible = True

'Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYEICGCCLOG_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Public Sub cmdSelect_SQL_Statistiques()

Dim Nb As Long, nbT As Long, K As Long
Dim xSQL As String, x As String, xWhere As String, xWhereLog As String
Dim arrEICGCCEIND(10) As Long, arrEICRICEIND(10) As Long
On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Ventilation par statut des dossiers                       " _
                      & "|>              Total" _
                      & "|>           en cours" _
                      & "|>           vérifiés" _
                      & "|>            annulés" _
                      & "|>          à ignorer" _
                      & "|>     non circulants" _
                      & "|>            rejetés"
fgSelect.Row = 0

currentAction = "cmdSelect_SQL_Statistiques"
    
Call DTPicker_Control(txtSelect_Options_St_AMJMIN, wAMJMin)
Call DTPicker_Control(txtSelect_Options_St_AMJMAX, WAMJMax)
xWhere = " where   EICGCCAMJ >= " & wAMJMin & " and   EICGCCAMJ <= " & WAMJMax
    
'____________________________________________________________________________________
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1

xSQL = "select EICGCCSTA,count(*) as Tally from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0" _
     & xWhere & " and EICGCCOPE = 'RI0' group by EICGCCSTA "
Set rsSab = cnsab.Execute(xSQL)

fgSelect.Col = 0: fgSelect.Text = "Images chèques"
nbT = 0
Do While Not rsSab.EOF
    x = rsSab("EICGCCSTA")
    Nb = rsSab("Tally")
    nbT = nbT + Nb
    Select Case x
        Case " ": fgSelect.Col = 2
        Case "V": fgSelect.Col = 3
        Case "A": fgSelect.Col = 4
        Case "I": fgSelect.Col = 5
        Case "@": fgSelect.Col = 6
        Case "R": fgSelect.Col = 7
        Case Else: fgSelect.Col = 8
    End Select
    fgSelect.Text = Format$(Nb, "### ##0")
   
    rsSab.MoveNext
Loop
fgSelect.Col = 1: fgSelect.Text = Format$(nbT, "### ##0")
fgSelect.CellFontBold = True
fgSelect.CellBackColor = vbCyan
'____________________________________________________________________________________
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1

xSQL = "select EICGCCSTA,count(*) as Tally from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0" _
     & xWhere & " and EICGCCOPE in ('REM','TRF') group by EICGCCSTA "
Set rsSab = cnsab.Execute(xSQL)

fgSelect.Col = 0: fgSelect.Text = "Remises BIA"
nbT = 0
Do While Not rsSab.EOF
    x = rsSab("EICGCCSTA")
    Nb = rsSab("Tally")
    nbT = nbT + Nb
    Select Case x
        Case " ": fgSelect.Col = 2
        Case "V": fgSelect.Col = 3
        Case "A": fgSelect.Col = 4
        Case "I": fgSelect.Col = 5
        Case "@": fgSelect.Col = 6
        Case "R": fgSelect.Col = 7
        Case Else: fgSelect.Col = 8
    End Select
    fgSelect.Text = Format$(Nb, "### ##0")
   
    rsSab.MoveNext
Loop
fgSelect.Col = 1: fgSelect.Text = Format$(nbT, "### ##0")
fgSelect.CellFontBold = True
fgSelect.CellBackColor = vbCyan
'____________________________________________________________________________________
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1

xSQL = "select EICGCCSTA,count(*) as Tally from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0" _
     & xWhere & " and EICGCCOPE not in ('RI0','REM','TRF') group by EICGCCSTA "
Set rsSab = cnsab.Execute(xSQL)

fgSelect.Col = 0: fgSelect.Text = "Divers"
nbT = 0
Do While Not rsSab.EOF
    x = rsSab("EICGCCSTA")
    Nb = rsSab("Tally")
    nbT = nbT + Nb
    Select Case x
        Case " ": fgSelect.Col = 2
        Case "V": fgSelect.Col = 3
        Case "A": fgSelect.Col = 4
        Case "I": fgSelect.Col = 5
        Case "@": fgSelect.Col = 6
        Case "R": fgSelect.Col = 7
        Case Else: fgSelect.Col = 8
    End Select
    fgSelect.Text = Format$(Nb, "### ##0")
   
    rsSab.MoveNext
Loop
fgSelect.Col = 1: fgSelect.Text = Format$(nbT, "### ##0")
fgSelect.CellFontBold = True
fgSelect.CellBackColor = vbCyan
'____________________________________________________________________________________

xWhereLog = " where   EICGCCLOGD >= " & wAMJMin & " and   EICGCCLOGD <= " & WAMJMax
fgSelect.Rows = fgSelect.Rows + 3
fgSelect.Row = fgSelect.Rows - 1

xSQL = "select count(*) as Tally from " & paramIBM_Library_SABSPE_XXX & ".YEICGCCLOG" _
       & xWhereLog & " and EICGCCLOGK = 'AF0' and EICGCCLOGA <> 'A' "
Set rsSab = cnsab.Execute(xSQL)

Nb = rsSab("Tally")
fgSelect.Col = 0: fgSelect.Text = "Nombre de télécopies 'chèque' demandées"
fgSelect.Col = 1: fgSelect.Text = Format$(Nb, "### ##0")
fgSelect.CellFontBold = True
fgSelect.CellBackColor = vbCyan
'____________________________________________________________________________________

fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1

xSQL = "select count(*) as Tally from " & paramIBM_Library_SABSPE_XXX & ".YEICGCCLOG" _
       & xWhereLog & " and EICGCCLOGK = 'AI1' and EICGCCLOGA <> 'A' "
Set rsSab = cnsab.Execute(xSQL)

Nb = rsSab("Tally")
fgSelect.Col = 0: fgSelect.Text = "Nombre de rejets d'IC reçues"
fgSelect.Col = 1: fgSelect.Text = Format$(Nb, "### ##0")
fgSelect.CellFontBold = True
fgSelect.CellBackColor = vbCyan
'____________________________________________________________________________________


fgSelect.Rows = fgSelect.Rows + 2
fgSelect.Row = fgSelect.Rows - 1
fgSelect.Col = 0: fgSelect.Text = "Ventilation RI0 / 160 par indice de circulation": fgSelect.CellBackColor = RGB(255, 139, 83)
fgSelect.Col = 1: fgSelect.Text = "Total": fgSelect.CellBackColor = RGB(255, 139, 83)
fgSelect.Col = 2: fgSelect.Text = "IC = 1": fgSelect.CellBackColor = RGB(255, 139, 83)
fgSelect.Col = 3: fgSelect.Text = "IC = 2": fgSelect.CellBackColor = RGB(255, 139, 83)
fgSelect.Col = 4: fgSelect.Text = "IC = 3": fgSelect.CellBackColor = RGB(255, 139, 83)
fgSelect.Col = 5: fgSelect.Text = "IC = 4": fgSelect.CellBackColor = RGB(255, 139, 83)
fgSelect.Col = 6: fgSelect.Text = "IC = 5": fgSelect.CellBackColor = RGB(255, 139, 83)
fgSelect.Col = 7: fgSelect.Text = "non circulant": fgSelect.CellBackColor = RGB(255, 139, 83)
fgSelect.Col = 8: fgSelect.Text = "???": fgSelect.CellBackColor = RGB(255, 139, 83)
'____________________________________________________________________________________
For K = 0 To 10
    arrEICGCCEIND(K) = 0: arrEICRICEIND(K) = 0
Next K
'____________________________________________________________________________________
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
fgSelect.Col = 0: fgSelect.Text = "SAB date CREATION = période (si comptabilisé)"

xSQL = "select EICRICD71  from " & paramIBM_Library_SAB & ".ZEICRIC0" _
        & " where EICRICOPR = 160 and EICRICDCP <> 0" _
       & " and EICRICDCR >= " & wAMJMin - 19000000 & " and EICRICDCR <= " & WAMJMax - 19000000
         
Set rsSab = cnsab.Execute(xSQL)

nbT = 0
Do While Not rsSab.EOF

    x = Mid$(rsSab("EICRICD71"), 56, 1)
    If IsNumeric(x) Then
        K = Val(x)

        If K < 1 Or K > 5 Then K = 6
    Else
        K = 7
    End If
    arrEICRICEIND(K) = arrEICRICEIND(K) + 1
    
    rsSab.MoveNext
Loop

For K = 1 To 7
    fgSelect.Col = K + 1
    fgSelect.Text = Format$(arrEICRICEIND(K), "### ###")
    nbT = nbT + arrEICRICEIND(K)
Next K
fgSelect.Col = 1: fgSelect.Text = Format$(nbT, "### ##0")
fgSelect.CellFontBold = True
fgSelect.CellBackColor = RGB(255, 169, 128)
'____________________________________________________________________________________
For K = 0 To 10
    arrEICGCCEIND(K) = 0: arrEICRICEIND(K) = 0
Next K
'____________________________________________________________________________________

fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
fgSelect.Col = 0: fgSelect.Text = "SAB date Réglement = période (si comptabilisé)"


'xSql = "select  EICRICNU1 , EICRICRGL , EICRICDCP , EICRICD71  from " & paramIBM_Library_SAB & ".ZEICRIC0" _
'        & " where EICRICOPR = 160 and EICRICDCP <> 0" _
'       & " and EICRICDCP >= " & wAmjMin - 19000000 & " and EICRICDCP <= " & wAmjMax - 19000000

xSQL = "select EICRICD71  from " & paramIBM_Library_SAB & ".ZEICRIC0" _
        & " where EICRICOPR = 160 and EICRICDCP <> 0" _
       & " and EICRICRGL >= " & wAMJMin - 19000000 & " and EICRICRGL <= " & WAMJMax - 19000000
         
Set rsSab = cnsab.Execute(xSQL)

nbT = 0
Do While Not rsSab.EOF

    x = Mid$(rsSab("EICRICD71"), 56, 1)
    If IsNumeric(x) Then
        K = Val(x)

        If K < 1 Or K > 5 Then K = 6
    Else
        K = 7
    End If
    arrEICRICEIND(K) = arrEICRICEIND(K) + 1
    
    rsSab.MoveNext
Loop

For K = 1 To 7
    fgSelect.Col = K + 1
    fgSelect.Text = Format$(arrEICRICEIND(K), "### ###")
    nbT = nbT + arrEICRICEIND(K)
Next K
fgSelect.Col = 1: fgSelect.Text = Format$(nbT, "### ##0")
fgSelect.CellFontBold = True
fgSelect.CellBackColor = RGB(255, 169, 128)

'____________________________________________________________________________________
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1

xSQL = "select EICGCCEIND,count(*) as Tally from " & paramIBM_Library_SABSPE_XXX & ".YEICGCC0" _
     & xWhere & " and EICGCCOPE = 'RI0' group by EICGCCEIND "
Set rsSab = cnsab.Execute(xSQL)

fgSelect.Col = 0: fgSelect.Text = "EIC_GCC : en gestion "
nbT = 0
Do While Not rsSab.EOF
    K = rsSab("EICGCCEIND")
    Nb = rsSab("Tally")
    nbT = nbT + Nb
    If K < 1 Or K > 5 Then K = 6
    arrEICGCCEIND(K) = arrEICGCCEIND(K) + Nb
    
    fgSelect.Col = K + 1
    fgSelect.Text = Format$(Nb, "### ###")
   
    rsSab.MoveNext
Loop
fgSelect.Col = 1: fgSelect.Text = Format$(nbT, "### ##0")
fgSelect.CellFontBold = True
fgSelect.CellBackColor = RGB(255, 169, 128)

'____________________________________________________________________________________

For K = 2 To 6
    fgSelect.Col = K
    fgSelect.Row = fgSelect.Rows - 2
    x = Trim(fgSelect.Text)
    fgSelect.Row = fgSelect.Rows - 1
    If x <> Trim(fgSelect.Text) Then fgSelect.CellBackColor = vbMagenta
  
Next K

'____________________________________________________________________________________

fgSelect.Visible = True

'Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYEICGCCLOG_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Public Sub cmdSelect_SQL_Statistiques_ChèquesBanque()
Dim blnPCI6 As Boolean
Dim Nb As Long
Dim xSQL As String, x As String, xWhere As String, xWhereLog As String
Dim wCur As Currency

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<code opération" _
                      & "|<compte                         " _
                      & "|<Intitulé                                                                     " _
                      & "|>nombre                   " _
                      & "|>montant €                       "
                      
fgSelect.Row = 0

currentAction = "cmdSelect_SQL_Statistiques"
    
Call DTPicker_Control(txtSelect_Options_St_AMJMIN, wAMJMin)
Call DTPicker_Control(txtSelect_Options_St_AMJMAX, WAMJMax)

x = Trim(txtSelect_Options_PCI)
If Len(x) = 6 Then
    blnPCI6 = True
    xWhere = " where EICGCCSTA = 'V' and EICGCCAMJ >= " & wAMJMin & " and   EICGCCAMJ <= " & WAMJMax & " and eicgccecpt = comptecom  and COMPTEOBL = '" & Trim(txtSelect_Options_PCI) & "'"
    xSQL = "select EICGCCOPE , EICGCCECPT ,  count(*) as Nb , sum(eicgccemt) as wCur, compteint from " _
         & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 , " & paramIBM_Library_SAB & ".ZCOMPTE0 " _
         & xWhere & "  group by EICGCCOPE , EICGCCECPT , compteint"
Else
    blnPCI6 = False
    xWhere = " where EICGCCSTA = 'V' and EICGCCAMJ >= " & wAMJMin & " and   EICGCCAMJ <= " & WAMJMax & " and eicgccecpt = comptecom  and COMPTEOBL like '" & Trim(txtSelect_Options_PCI) & "%'"
    xSQL = "select EICGCCOPE , COMPTEOBL ,  count(*) as Nb , sum(eicgccemt) as wCur , PLANINTIT from " _
         & paramIBM_Library_SABSPE_XXX & ".YEICGCC0 , " & paramIBM_Library_SAB & ".ZCOMPTE0 , " & paramIBM_Library_SAB & ".ZPLAN0" _
         & xWhere & "  and COMPTEOBL = PLANCOOBL group by EICGCCOPE , COMPTEOBL , PLANINTIT"
End If
  
    
'____________________________________________________________________________________

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect.Col = 0: fgSelect.Text = rsSab("EICGCCOPE")
    If blnPCI6 Then
        fgSelect.Col = 2: fgSelect.Text = rsSab("COMPTEINT")
        fgSelect.Col = 1: fgSelect.Text = rsSab("EICGCCECPT")
    Else
        fgSelect.Col = 1: fgSelect.Text = rsSab("COMPTEOBL")
        fgSelect.Col = 2: fgSelect.Text = rsSab("PLANINTIT")
    End If
    
   fgSelect.Col = 3: fgSelect.Text = Format$(rsSab("Nb"), "### ##0")
    fgSelect.Col = 4: fgSelect.Text = Format$(rsSab("wCur"), "### ### ### ##0.00")
       
    rsSab.MoveNext
Loop
'____________________________________________________________________________________

fgSelect.Visible = True

'Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYEICGCCLOG_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub




Public Function cmdSelect_SQL_AMJ()
If IsNull(txtSelect_EICGCCAMJ) Then
    WAMJMax = 0
Else
    Call DTPicker_Control(txtSelect_EICGCCAMJ, WAMJMax)
End If
If IsNull(txtSelect_EICGCCAMJ_Min) Then
    wAMJMin = 0
Else
    Call DTPicker_Control(txtSelect_EICGCCAMJ_Min, wAMJMin)
End If

If wAMJMin = 0 And WAMJMax = 0 Then
    cmdSelect_SQL_AMJ = ""
Else
        If wAMJMin = 0 Then
            cmdSelect_SQL_AMJ = " and   EICGCCAMJ = " & WAMJMax
        Else
            If WAMJMax = 0 Then
                cmdSelect_SQL_AMJ = " and   EICGCCAMJ = " & wAMJMin
            Else
            If wAMJMin < WAMJMax Then
                cmdSelect_SQL_AMJ = " and   EICGCCAMJ >= " & wAMJMin & " And EICGCCAMJ <= " & WAMJMax
            Else
                cmdSelect_SQL_AMJ = " and   EICGCCAMJ <= " & wAMJMin & " And EICGCCAMJ >= " & WAMJMax
            End If
        End If
    End If
End If

End Function


Public Sub fraDetail_Display_BDF()
Dim I As Long
Dim xSQL As String
Dim x As String, blnOk As Boolean
Dim mEICRICDO6 As String, K1 As Integer

xSQL = "select EICRICREF , EICRICDO6 from " & paramIBM_Library_SAB & ".ZEICRIC0 " _
     & " Where EICRICNU1 = " & xYEICGCC0.EICGCCDOS
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    libDetail_EICRICREF = Trim(rsSab("EICRICREF"))
    mEICRICDO6 = rsSab("EICRICDO6")
    libDetail_EICRICDO6 = mEICRICDO6
Else
    libDetail_EICRICREF = "?????"
    libDetail_EICRICDO6 = "?????"
End If
blnOk = False
K1 = InStr(mEICRICDO6, xYEICGCC0.EICGCCXBQ)
If K1 > 0 Then
    x = Mid$(mEICRICDO6, K1 + 5, 5)
    xSQL = "select FGDNOETA from " & paramIBM_Library_SAB & ".ZFGDBDF0 " _
         & " Where FGDCOETA = '" & Trim(xYEICGCC0.EICGCCXBQ) & "' and FGDCOGUI = '" & x & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        blnOk = True
        libDetail_EICGCCXCPT = Trim(rsSab("FGDNOETA"))
    End If
End If
If Not blnOk Then
    xSQL = "select FGDNOETA from " & paramIBM_Library_SAB & ".ZFGDBDF0 " _
         & " Where FGDCOETA = '" & Trim(xYEICGCC0.EICGCCXBQ) & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        x = Trim(rsSab("FGDNOETA"))
        K1 = InStr(x, " ")
        libDetail_EICGCCXCPT = Mid$(x, 1, K1)
    End If
End If

End Sub

