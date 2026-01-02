VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSAB_Dossier 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_Dossier"
   ClientHeight    =   12165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16335
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SAB_Dossier.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   12165
   ScaleWidth      =   16335
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   8700
      TabIndex        =   5
      Top             =   75
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   11640
      Left            =   15
      TabIndex        =   3
      Top             =   480
      Width           =   16275
      _ExtentX        =   28707
      _ExtentY        =   20532
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Suivi des dossiers SAB"
      TabPicture(0)   =   "SAB_Dossier.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "."
      TabPicture(1)   =   "SAB_Dossier.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgX"
      Tab(1).Control(1)=   "fraSelect_Options_Scan_Liste"
      Tab(1).Control(2)=   "fraSelect_Options_3uti"
      Tab(1).Control(3)=   "fraSelect_Options_Log"
      Tab(1).Control(4)=   "fraYDOSXOD0"
      Tab(1).Control(5)=   "fgLOG"
      Tab(1).Control(6)=   "fraSwift"
      Tab(1).Control(7)=   "fraSelect_Options_5"
      Tab(1).Control(8)=   "fraSelect_Options_Xc"
      Tab(1).Control(9)=   "fraSelect_Options_6"
      Tab(1).Control(10)=   "fraCompte"
      Tab(1).Control(11)=   "lstW"
      Tab(1).Control(12)=   "fgCPTPIE"
      Tab(1).ControlCount=   13
      Begin MSFlexGridLib.MSFlexGrid fgX 
         Height          =   1215
         Left            =   -71670
         TabIndex        =   108
         Top             =   9825
         Visible         =   0   'False
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   8421376
         ForeColorFixed  =   -2147483633
         BackColorBkg    =   -2147483633
         AllowUserResizing=   3
         FormatString    =   "<fgX                                                                   |||||  "
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
      Begin VB.Frame fraSelect_Options_Scan_Liste 
         BackColor       =   &H00D0F0FF&
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   -72330
         TabIndex        =   105
         Top             =   3660
         Visible         =   0   'False
         Width           =   8220
         Begin MSComCtl2.DTPicker txtSelect_Options_Scan_Liste_AMJ 
            Height          =   300
            Left            =   2415
            TabIndex        =   106
            Top             =   495
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
            Format          =   125239299
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.Label lblSelect_Options_Scan_Liste_AMJ 
            BackColor       =   &H00D0F0FF&
            Caption         =   "date de la numérisation"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   495
            TabIndex        =   107
            Top             =   540
            Width           =   1905
         End
      End
      Begin VB.Frame fraSelect_Options_3uti 
         BackColor       =   &H00F0FFFF&
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   -74745
         TabIndex        =   92
         Top             =   4650
         Visible         =   0   'False
         Width           =   8220
         Begin VB.ComboBox cboSelect_Options_3uti_CDODOSNOT 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5145
            TabIndex        =   99
            Text            =   "cboSelect_Options_3uti_CDODOSNOT"
            Top             =   315
            Width           =   2880
         End
         Begin VB.CheckBox chkSelect_Options_3uti_Swift 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F0FFFF&
            Caption         =   "Afficher la date du dernier message Swift 799 ou 999"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   300
            Left            =   75
            TabIndex        =   98
            Top             =   750
            Width           =   4725
         End
         Begin VB.ComboBox cboSelect_Options_3uti_UTI 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5835
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   720
            Width           =   2250
         End
         Begin MSComCtl2.DTPicker txtSelect_Options_3uti_AmjMin 
            Height          =   300
            Left            =   855
            TabIndex        =   93
            Top             =   315
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
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
            Format          =   138870787
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin MSComCtl2.DTPicker txtSelect_Options_3uti_AmjMax 
            Height          =   300
            Left            =   2430
            TabIndex        =   94
            Top             =   330
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
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
            Format          =   138870787
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.Label lblSelect_Options_3uti_CDODOSNOT 
            BackColor       =   &H00F0FFFF&
            Caption         =   "BQ émettrice"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3900
            TabIndex        =   100
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label lblSelect_Options_3uti_AMJ 
            BackColor       =   &H00F0FFFF&
            Caption         =   "période"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   105
            TabIndex        =   96
            Top             =   390
            Width           =   645
         End
         Begin VB.Label lblSelect_Options_3uti_UTI 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Utilisateur"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4905
            TabIndex        =   95
            Top             =   780
            Width           =   990
         End
      End
      Begin VB.Frame fraSelect_Options_Log 
         BackColor       =   &H00F0FFFF&
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   -74670
         TabIndex        =   79
         Top             =   6825
         Visible         =   0   'False
         Width           =   8220
         Begin VB.TextBox txtSelect_Options_Log_OPE 
            Height          =   330
            Left            =   6060
            MaxLength       =   3
            TabIndex        =   91
            Top             =   465
            Width           =   600
         End
         Begin MSComCtl2.DTPicker txtSelect_Options_Log_AmjMin 
            Height          =   300
            Left            =   1350
            TabIndex        =   81
            Top             =   495
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
            Format          =   138870787
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin MSComCtl2.DTPicker txtSelect_Options_Log_AmjMax 
            Height          =   300
            Left            =   2985
            TabIndex        =   82
            Top             =   495
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
            Format          =   138805251
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.Label lblSelect_Options_Log_OPE 
            BackColor       =   &H00F0FFFF&
            Caption         =   "code opération"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4755
            TabIndex        =   90
            Top             =   525
            Width           =   1380
         End
         Begin VB.Label lblSelect_Options_Log 
            BackColor       =   &H00F0FFFF&
            Caption         =   "période"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   495
            TabIndex        =   80
            Top             =   540
            Width           =   1365
         End
      End
      Begin VB.Frame fraYDOSXOD0 
         BackColor       =   &H00FFE0FF&
         Caption         =   "Modification de l'imputation du dossier"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   3045
         Left            =   -63120
         TabIndex        =   46
         Top             =   840
         Visible         =   0   'False
         Width           =   4500
         Begin VB.CommandButton cmdYDOSXOD0__Update 
            BackColor       =   &H0080FF80&
            Caption         =   "Enregistrer"
            Height          =   480
            Left            =   2715
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   2325
            Width           =   900
         End
         Begin VB.CommandButton cmdYDOSXOD0_Quit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Abandonner"
            Height          =   480
            Left            =   540
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   2295
            Width           =   990
         End
         Begin VB.TextBox txtDOSXODNUM 
            Height          =   330
            Left            =   2610
            TabIndex        =   75
            Top             =   1800
            Width           =   1200
         End
         Begin VB.ComboBox cboDOSXODOPE 
            Height          =   330
            Left            =   2610
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   1110
            Width           =   1200
         End
         Begin VB.Label libYDOSXOD0 
            BackColor       =   &H00C0FFFF&
            Caption         =   "mise à jour"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   315
            TabIndex        =   78
            Top             =   465
            Width           =   3840
         End
         Begin VB.Label lblDOSXODNUM 
            BackColor       =   &H00FFE0FF&
            Caption         =   "numéro opération"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   405
            TabIndex        =   74
            Top             =   1845
            Width           =   2145
         End
         Begin VB.Label lblDOSXODOPE 
            BackColor       =   &H00FFE0FF&
            Caption         =   "Code opération"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   435
            TabIndex        =   48
            Top             =   1170
            Width           =   2100
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgLOG 
         Height          =   1215
         Left            =   -73485
         TabIndex        =   83
         Top             =   6315
         Visible         =   0   'False
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   16777215
         ForeColor       =   16384
         BackColorFixed  =   8421376
         ForeColorFixed  =   -2147483633
         BackColorBkg    =   -2147483633
         AllowUserResizing=   3
         FormatString    =   $"SAB_Dossier.frx":0342
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
      Begin VB.Frame fraSwift 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9500
         Left            =   -69180
         TabIndex        =   85
         Top             =   1710
         Visible         =   0   'False
         Width           =   7200
         Begin VB.CheckBox chkSIDE_DB_Show 
            BackColor       =   &H00C0FFFF&
            Caption         =   "afficher le message et l'historique du traitement SAA"
            Height          =   255
            Left            =   60
            TabIndex        =   86
            Top             =   600
            Width           =   6945
         End
         Begin MSFlexGridLib.MSFlexGrid fgSwift 
            Height          =   8445
            Left            =   60
            TabIndex        =   87
            Top             =   930
            Width           =   7005
            _ExtentX        =   12356
            _ExtentY        =   14896
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16777215
            ForeColor       =   12582912
            BackColorFixed  =   16777168
            ForeColorFixed  =   16711680
            BackColorBkg    =   16777215
            GridColor       =   12632064
            GridColorFixed  =   12632064
            AllowUserResizing=   3
            FormatString    =   $"SAB_Dossier.frx":044C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label libSWIFT_SWISABSWID 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   75
            TabIndex        =   88
            Top             =   210
            Width           =   6960
         End
      End
      Begin VB.Frame fraSelect_Options_5 
         BackColor       =   &H00F0FFFF&
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   -74550
         TabIndex        =   71
         Top             =   375
         Visible         =   0   'False
         Width           =   3780
         Begin VB.ComboBox cboSelect_DOSSLDSTA 
            Height          =   330
            Left            =   960
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   510
            Width           =   1380
         End
         Begin VB.Label lblSelect_DOSSLDSTA 
            BackColor       =   &H00F0FFFF&
            Caption         =   "état du dossier"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   915
            TabIndex        =   73
            Top             =   120
            Width           =   1365
         End
      End
      Begin VB.Frame fraSelect_Options_Xc 
         BackColor       =   &H00F0FFFF&
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   -74730
         TabIndex        =   63
         Top             =   1650
         Visible         =   0   'False
         Width           =   5715
         Begin VB.OptionButton optSelect_DOSCD7DAN_In 
            BackColor       =   &H00F0FFFF&
            Caption         =   "inclure"
            Height          =   210
            Left            =   2150
            TabIndex        =   69
            Top             =   465
            Width           =   1230
         End
         Begin VB.OptionButton optSelect_DOSCD7DAN_Out 
            BackColor       =   &H00F0FFFF&
            Caption         =   "exclure"
            Height          =   210
            Left            =   2150
            TabIndex        =   68
            Top             =   150
            Value           =   -1  'True
            Width           =   1230
         End
         Begin VB.ComboBox cboSelect_DOSCD7DSIT 
            Height          =   330
            Left            =   285
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   600
            Width           =   1380
         End
         Begin MSComCtl2.DTPicker txtSelect_DOSCD7DAN 
            Height          =   300
            Left            =   4470
            TabIndex        =   67
            Top             =   750
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
            Format          =   138805251
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.Label lblSelect_DOSCD7DAN 
            BackColor       =   &H00F0FFFF&
            Caption         =   "les dossiers annulés avant le"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2130
            TabIndex        =   66
            Top             =   810
            Width           =   2070
         End
         Begin VB.Label lblSelect_DOSCD7DSIT 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Date de situation"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   285
            TabIndex        =   65
            Top             =   180
            Width           =   1365
         End
      End
      Begin VB.Frame fraSelect_Options_6 
         BackColor       =   &H00F0FFFF&
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   -74310
         TabIndex        =   55
         Top             =   8340
         Visible         =   0   'False
         Width           =   8265
         Begin VB.TextBox txtSelect_6_PCI 
            Height          =   330
            Left            =   405
            MaxLength       =   5
            TabIndex        =   60
            Text            =   "13221"
            Top             =   510
            Width           =   1140
         End
         Begin VB.TextBox txtSelect_6_CLIEANCLI 
            Height          =   330
            Left            =   2640
            TabIndex        =   57
            Top             =   510
            Width           =   1140
         End
         Begin MSComCtl2.DTPicker txtSelect_6_AMJMin 
            Height          =   300
            Left            =   4515
            TabIndex        =   56
            Top             =   510
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
            Format          =   125960195
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.Label libSelect_6_PCI 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Prov (CDE: 13221, CDI: 25302)"
            Height          =   270
            Left            =   105
            TabIndex        =   70
            Top             =   915
            Width           =   2580
         End
         Begin VB.Label libSelect_6_AMJMax 
            BackColor       =   &H00F0FFFF&
            Caption         =   "au"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6240
            TabIndex        =   62
            Top             =   510
            Width           =   1410
         End
         Begin VB.Label lblSelect_6_PCI 
            BackColor       =   &H00F0FFFF&
            Caption         =   "PCI (5 ou 6)"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   540
            TabIndex        =   61
            Top             =   120
            Width           =   840
         End
         Begin VB.Label lblSelect_6_CLIENACLI 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Client"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2865
            TabIndex        =   59
            Top             =   120
            Width           =   840
         End
         Begin VB.Label lblSelect_6_AMJMin 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Période du"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4515
            TabIndex        =   58
            Top             =   120
            Width           =   1065
         End
      End
      Begin VB.Frame fraCompte 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Compte"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   -67875
         TabIndex        =   25
         Top             =   165
         Visible         =   0   'False
         Width           =   6315
         Begin VB.TextBox txtD_CLIENARSD 
            Height          =   330
            Left            =   2835
            Locked          =   -1  'True
            TabIndex        =   44
            Text            =   "CLIENARSD"
            Top             =   3500
            Width           =   675
         End
         Begin VB.TextBox txtD_CLIENARES 
            Height          =   330
            Left            =   4275
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "CLIENARES"
            Top             =   2500
            Width           =   960
         End
         Begin VB.TextBox txtD_CLIENANAT 
            Height          =   330
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "CLIENARES"
            Top             =   3500
            Width           =   690
         End
         Begin VB.TextBox txtD_CLIENARA1 
            Height          =   330
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "COMPTEINT"
            Top             =   3000
            Width           =   4050
         End
         Begin VB.TextBox txtD_CLIENASIG 
            Height          =   330
            Left            =   3195
            Locked          =   -1  'True
            TabIndex        =   38
            Text            =   "CLIENASIG"
            Top             =   2500
            Width           =   960
         End
         Begin VB.TextBox txtD_CLIENACLI 
            Height          =   330
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "CLIENACLI"
            Top             =   2500
            Width           =   1140
         End
         Begin VB.TextBox txtD_COMPTECLO 
            Height          =   330
            Left            =   4545
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "COMPTECLO"
            Top             =   1500
            Width           =   1215
         End
         Begin VB.TextBox txtD_COMPTEOUV 
            Height          =   330
            Left            =   2955
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "COMPTEOUV"
            Top             =   1485
            Width           =   1140
         End
         Begin VB.TextBox txtD_PLANCOPRO 
            Height          =   330
            Left            =   5505
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "PLANCOPRO"
            Top             =   500
            Width           =   495
         End
         Begin VB.TextBox txtD_COMPTEFON 
            Height          =   330
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   31
            Text            =   "COMPTEFON"
            Top             =   1500
            Width           =   405
         End
         Begin VB.TextBox txtD_COMPTEOBL 
            Height          =   330
            Left            =   4425
            Locked          =   -1  'True
            TabIndex        =   30
            Text            =   "COMPTEOBL"
            Top             =   500
            Width           =   960
         End
         Begin VB.TextBox txtD_COMPTEINT 
            Height          =   330
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "COMPTEINT"
            Top             =   1010
            Width           =   4050
         End
         Begin VB.TextBox txtD_COMPTECOM 
            Height          =   330
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "COMPTECOM"
            Top             =   500
            Width           =   2310
         End
         Begin VB.Label lblD_CLIENANAT 
            BackColor       =   &H00D0F0FF&
            Caption         =   "pays nationalité, rés"
            Height          =   345
            Left            =   180
            TabIndex        =   42
            Top             =   3550
            Width           =   1530
         End
         Begin VB.Label lblD_CLIENARA1 
            BackColor       =   &H00D0F0FF&
            Caption         =   "intitulé"
            Height          =   345
            Left            =   180
            TabIndex        =   39
            Top             =   3050
            Width           =   1530
         End
         Begin VB.Label lblD_CLIENACLI 
            BackColor       =   &H00D0F0FF&
            Caption         =   "client, sigle, resp"
            Height          =   345
            Left            =   180
            TabIndex        =   36
            Top             =   2550
            Width           =   1530
         End
         Begin VB.Label lblD_COMPTEFON 
            BackColor       =   &H00D0F0FF&
            Caption         =   "code fonct,Dcre,Dclo"
            Height          =   345
            Left            =   180
            TabIndex        =   33
            Top             =   1550
            Width           =   1530
         End
         Begin VB.Label lblD_COMPTEINT 
            BackColor       =   &H00D0F0FF&
            Caption         =   "intitulé"
            Height          =   345
            Left            =   165
            TabIndex        =   28
            Top             =   1065
            Width           =   1530
         End
         Begin VB.Label lblD_COMPTECOM 
            BackColor       =   &H00D0F0FF&
            Caption         =   "compte PCI produit"
            Height          =   345
            Left            =   180
            TabIndex        =   26
            Top             =   550
            Width           =   1530
         End
      End
      Begin VB.ListBox lstW 
         BackColor       =   &H00FFFFFF&
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
         Left            =   -68925
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   3270
         Visible         =   0   'False
         Width           =   4212
      End
      Begin VB.Frame fraTab0 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   11235
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   16050
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   4350
            Left            =   120
            TabIndex        =   6
            Top             =   1410
            Visible         =   0   'False
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   7673
            _Version        =   393216
            Rows            =   1
            Cols            =   5
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16777215
            ForeColor       =   16384
            BackColorFixed  =   8421376
            ForeColorFixed  =   16777215
            BackColorSel    =   12648384
            BackColorBkg    =   16777215
            AllowBigSelection=   0   'False
            AllowUserResizing=   3
            FormatString    =   "<Dev   |<Opé    |>Dossier           |<Client               |"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComDlg.CommonDialog CmDialog2 
            Left            =   10890
            Top             =   870
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   5265
            Left            =   105
            TabIndex        =   50
            Top             =   5865
            Visible         =   0   'False
            Width           =   15765
            _ExtentX        =   27808
            _ExtentY        =   9287
            _Version        =   393216
            Tabs            =   6
            Tab             =   5
            TabsPerRow      =   6
            TabHeight       =   520
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Mvts comptables"
            TabPicture(0)   =   "SAB_Dossier.frx":04DB
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "fgBIAMVT"
            Tab(0).Control(1)=   "cmdSAB_Dossier_DB"
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Gestion"
            TabPicture(1)   =   "SAB_Dossier.frx":04F7
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fgDossier"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Swift"
            TabPicture(2)   =   "SAB_Dossier.frx":0513
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "fgYSWISAB0"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Courrier"
            TabPicture(3)   =   "SAB_Dossier.frx":052F
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "cmdSAB_Dossier_CDO"
            Tab(3).Control(1)=   "fgCourrier"
            Tab(3).ControlCount=   2
            TabCaption(4)   =   "Scan"
            TabPicture(4)   =   "SAB_Dossier.frx":054B
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "fgScan"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Commissions"
            TabPicture(5)   =   "SAB_Dossier.frx":0567
            Tab(5).ControlEnabled=   -1  'True
            Tab(5).Control(0)=   "fgCOM"
            Tab(5).Control(0).Enabled=   0   'False
            Tab(5).Control(1)=   "fraECNFPT"
            Tab(5).Control(1).Enabled=   0   'False
            Tab(5).ControlCount=   2
            Begin VB.Frame fraECNFPT 
               BackColor       =   &H0080C0FF&
               Caption         =   "Calcul de la commission prorata temporis"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4500
               Left            =   10965
               TabIndex        =   110
               Top             =   450
               Visible         =   0   'False
               Width           =   4500
               Begin VB.Frame fraECNFPT_TOT_X 
                  BackColor       =   &H00E0FFE0&
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1395
                  Left            =   0
                  TabIndex        =   112
                  Top             =   3100
                  Width           =   4500
                  Begin MSComCtl2.DTPicker txtECNFPT_DREG 
                     Height          =   390
                     Left            =   1890
                     TabIndex        =   125
                     Top             =   150
                     Width           =   1335
                     _ExtentX        =   2355
                     _ExtentY        =   688
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Calibri"
                        Size            =   11.25
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
                     Format          =   138346499
                     CurrentDate     =   38699.44875
                     MaxDate         =   401768
                     MinDate         =   36526.4425347222
                  End
                  Begin VB.Label libECNFPT_TOT_X 
                     Alignment       =   2  'Center
                     BackColor       =   &H00808000&
                     Caption         =   "--"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   300
                     Left            =   1860
                     TabIndex        =   131
                     Top             =   1005
                     Width           =   1410
                  End
                  Begin VB.Label lblECNFPT_TOT_X 
                     BackColor       =   &H00808000&
                     Caption         =   "Total"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   300
                     Left            =   360
                     TabIndex        =   130
                     Top             =   1020
                     Width           =   1410
                  End
                  Begin VB.Label libECNFPT_NBJ_X 
                     BackColor       =   &H00E0FFE0&
                     Caption         =   "--"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   300
                     Left            =   3390
                     TabIndex        =   129
                     Top             =   180
                     Width           =   705
                  End
                  Begin VB.Label libECNFPT_MON_X 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E0FFE0&
                     Caption         =   "--"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   300
                     Left            =   1905
                     TabIndex        =   128
                     Top             =   630
                     Width           =   1380
                  End
                  Begin VB.Label lblECNFPT_DREG 
                     BackColor       =   &H00E0FFE0&
                     Caption         =   "Arrêté au"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   300
                     Left            =   375
                     TabIndex        =   127
                     Top             =   240
                     Width           =   1410
                  End
                  Begin VB.Label lblECNFPT_MON_X 
                     BackColor       =   &H00E0FFE0&
                     Caption         =   "Commission"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   300
                     Left            =   300
                     TabIndex        =   126
                     Top             =   630
                     Width           =   1410
                  End
               End
               Begin VB.Frame fraECNFPT_COM 
                  BackColor       =   &H00E0FFFF&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2800
                  Left            =   0
                  TabIndex        =   111
                  Top             =   350
                  Width           =   4500
                  Begin VB.TextBox txtECNFPT_MON 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   390
                     Left            =   1800
                     TabIndex        =   124
                     Top             =   2300
                     Width           =   2055
                  End
                  Begin VB.TextBox txtECNFPT_NBJ 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   390
                     Left            =   1800
                     TabIndex        =   123
                     Top             =   1900
                     Width           =   855
                  End
                  Begin VB.TextBox txtECNFPT_DDEB 
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   390
                     Left            =   1800
                     TabIndex        =   122
                     Top             =   1500
                     Width           =   2610
                  End
                  Begin VB.TextBox txtECNFPT_PER 
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   390
                     Left            =   1800
                     TabIndex        =   121
                     Top             =   1100
                     Width           =   420
                  End
                  Begin VB.TextBox txtECNFPT_TX1 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   390
                     Left            =   1800
                     TabIndex        =   120
                     Top             =   700
                     Width           =   1050
                  End
                  Begin VB.TextBox txtECNFPT_MTA 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   390
                     Left            =   1800
                     TabIndex        =   119
                     Top             =   300
                     Width           =   2055
                  End
                  Begin VB.Label lblECNFPT_MON 
                     BackColor       =   &H00E0FFFF&
                     Caption         =   "Commission"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   300
                     Left            =   300
                     TabIndex        =   118
                     Top             =   2400
                     Width           =   1410
                  End
                  Begin VB.Label lblECNFPT_NBJ 
                     BackColor       =   &H00E0FFFF&
                     Caption         =   "Nb jours"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   300
                     Left            =   300
                     TabIndex        =   117
                     Top             =   2000
                     Width           =   1410
                  End
                  Begin VB.Label lblECNFPT_DBP 
                     BackColor       =   &H00E0FFFF&
                     Caption         =   "Période"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   300
                     Left            =   300
                     TabIndex        =   116
                     Top             =   1600
                     Width           =   1410
                  End
                  Begin VB.Label lblECNFPT_PER 
                     BackColor       =   &H00E0FFFF&
                     Caption         =   "Périodicité"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   300
                     Left            =   300
                     TabIndex        =   115
                     Top             =   1200
                     Width           =   1410
                  End
                  Begin VB.Label lblECNFPT_TX1 
                     BackColor       =   &H00E0FFFF&
                     Caption         =   "Taux"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   300
                     Left            =   300
                     TabIndex        =   114
                     Top             =   800
                     Width           =   1410
                  End
                  Begin VB.Label lblECNFPT_MTA 
                     BackColor       =   &H00E0FFFF&
                     Caption         =   "Assiette"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   11.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   300
                     Left            =   300
                     TabIndex        =   113
                     Top             =   400
                     Width           =   1410
                  End
               End
            End
            Begin VB.CommandButton cmdSAB_Dossier_CDO 
               BackColor       =   &H00FFFF00&
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   -65160
               Style           =   1  'Graphical
               TabIndex        =   101
               ToolTipText     =   "Cliquer ici pour afficher toutes les écritures comptables concernant ce dossier"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmdSAB_Dossier_DB 
               BackColor       =   &H0080C0FF&
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   -72945
               Style           =   1  'Graphical
               TabIndex        =   89
               ToolTipText     =   "Cliquer ici pour afficher toutes les écritures comptables concernant ce dossier"
               Top             =   15
               Width           =   585
            End
            Begin MSFlexGridLib.MSFlexGrid fgBIAMVT 
               Height          =   4785
               Left            =   -74940
               TabIndex        =   51
               Top             =   660
               Width           =   15450
               _ExtentX        =   27252
               _ExtentY        =   8440
               _Version        =   393216
               Cols            =   8
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16777215
               ForeColor       =   16384
               BackColorFixed  =   8421376
               ForeColorFixed  =   -2147483633
               BackColorBkg    =   -2147483633
               AllowUserResizing=   3
               FormatString    =   $"SAB_Dossier.frx":0583
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid fgYSWISAB0 
               Height          =   4860
               Left            =   -74895
               TabIndex        =   84
               ToolTipText     =   "cliquer pour afficher le détail du message swift"
               Top             =   420
               Visible         =   0   'False
               Width           =   15540
               _ExtentX        =   27411
               _ExtentY        =   8573
               _Version        =   393216
               Rows            =   1
               Cols            =   13
               FixedCols       =   0
               RowHeightMin    =   300
               BackColor       =   16777215
               ForeColor       =   12582912
               BackColorFixed  =   10526720
               ForeColorFixed  =   16777215
               BackColorSel    =   12648384
               BackColorBkg    =   15794175
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
               GridLinesFixed  =   1
               AllowUserResizing=   3
               FormatString    =   $"SAB_Dossier.frx":06A0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid fgDossier 
               Height          =   4755
               Left            =   -74925
               TabIndex        =   102
               Top             =   375
               Visible         =   0   'False
               Width           =   15540
               _ExtentX        =   27411
               _ExtentY        =   8387
               _Version        =   393216
               Cols            =   12
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16777215
               ForeColor       =   4210752
               BackColorFixed  =   8438015
               ForeColorFixed  =   16384
               BackColorBkg    =   -2147483633
               AllowUserResizing=   3
               FormatString    =   $"SAB_Dossier.frx":0799
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid fgCourrier 
               Height          =   4740
               Left            =   -74880
               TabIndex        =   103
               Top             =   480
               Visible         =   0   'False
               Width           =   15570
               _ExtentX        =   27464
               _ExtentY        =   8361
               _Version        =   393216
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16777215
               ForeColor       =   4210752
               BackColorFixed  =   16777088
               ForeColorFixed  =   16384
               BackColorBkg    =   -2147483633
               AllowUserResizing=   3
               FormatString    =   $"SAB_Dossier.frx":089E
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
            Begin MSFlexGridLib.MSFlexGrid fgScan 
               Height          =   4875
               Left            =   -74865
               TabIndex        =   104
               Top             =   390
               Visible         =   0   'False
               Width           =   15510
               _ExtentX        =   27358
               _ExtentY        =   8599
               _Version        =   393216
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16777215
               ForeColor       =   4210752
               BackColorFixed  =   8438015
               ForeColorFixed  =   16384
               BackColorBkg    =   -2147483633
               AllowUserResizing=   3
               FormatString    =   $"SAB_Dossier.frx":09DE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid fgCOM 
               Height          =   4755
               Left            =   105
               TabIndex        =   109
               Top             =   390
               Visible         =   0   'False
               Width           =   15540
               _ExtentX        =   27411
               _ExtentY        =   8387
               _Version        =   393216
               Cols            =   19
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16777215
               ForeColor       =   4210752
               BackColorFixed  =   8438015
               ForeColorFixed  =   16384
               BackColorBkg    =   -2147483633
               AllowUserResizing=   3
               FormatString    =   $"SAB_Dossier.frx":0B1E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.Frame fraDetail 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Left            =   4335
            TabIndex        =   14
            Top             =   1410
            Visible         =   0   'False
            Width           =   11520
            Begin VB.Label libDOSSLDCLI 
               BackColor       =   &H00C0FFFF&
               Caption         =   "client"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   255
               Left            =   3150
               TabIndex        =   53
               Top             =   150
               Width           =   1815
            End
            Begin VB.Label libDOSSLDLIB 
               BackColor       =   &H00C0FFFF&
               Caption         =   "libellé"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   255
               Left            =   5130
               TabIndex        =   52
               Top             =   150
               Width           =   6360
            End
            Begin VB.Label libDOSSLDDEV 
               BackColor       =   &H00C0FFFF&
               Caption         =   "dev"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1110
               TabIndex        =   17
               Top             =   150
               Width           =   720
            End
            Begin VB.Label libDOSSLDNUM 
               BackColor       =   &H00C0FFFF&
               Caption         =   "num"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   255
               Left            =   1920
               TabIndex        =   16
               Top             =   150
               Width           =   1140
            End
            Begin VB.Label libDOSSLDOPE 
               BackColor       =   &H00C0FFFF&
               Caption         =   "dossier"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   45
               TabIndex        =   15
               Top             =   150
               Width           =   1005
            End
         End
         Begin VB.ComboBox cboSelect_SQL 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   10860
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   5040
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   13035
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   900
            Width           =   1335
         End
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00F0FFFF&
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1212
            Left            =   135
            TabIndex        =   7
            Top             =   255
            Visible         =   0   'False
            Width           =   9855
            Begin VB.CheckBox chkSelect_DOSSLDMG 
               BackColor       =   &H00D0F0FF&
               Caption         =   "dossier : solde comptable <> solde gestion "
               Height          =   250
               Left            =   6060
               TabIndex        =   54
               Top             =   165
               Width           =   3570
            End
            Begin VB.CheckBox chkSelect_DOSSLDSVC 
               BackColor       =   &H00F0FFFF&
               Caption         =   "exclure les dossiers' saisie en cours' (01)"
               Height          =   250
               Left            =   6075
               TabIndex        =   45
               Top             =   450
               Width           =   3570
            End
            Begin VB.CheckBox chkSelect_DOSSLDSTA 
               BackColor       =   &H00F0FFFF&
               Caption         =   "exclure les dossiers clôturés(80,90)"
               Height          =   250
               Left            =   6045
               TabIndex        =   24
               Top             =   885
               Value           =   1  'Checked
               Width           =   3555
            End
            Begin VB.ComboBox cboSelect_DOSSLDPCI 
               Height          =   330
               Left            =   4830
               Sorted          =   -1  'True
               TabIndex        =   22
               Text            =   "PCI5"
               Top             =   135
               Width           =   1176
            End
            Begin VB.ComboBox cboSelect_DOSSLDOPE 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2745
               TabIndex        =   20
               Text            =   "cboSelect_DOSSLDOPE"
               Top             =   450
               Width           =   1170
            End
            Begin VB.Frame fraSelect_Options_1 
               BackColor       =   &H00F0FFFF&
               BorderStyle     =   0  'None
               Height          =   1035
               Left            =   270
               TabIndex        =   10
               Top             =   165
               Width           =   2460
               Begin VB.TextBox txtSelect_DOSSLDNUM 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   60
                  TabIndex        =   0
                  Top             =   240
                  Width           =   1050
               End
               Begin VB.ComboBox cboSelect_DOSSLDDEV 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   1290
                  Sorted          =   -1  'True
                  TabIndex        =   18
                  Text            =   "dev"
                  Top             =   255
                  Width           =   705
               End
               Begin VB.Label lblSelect_DOSSLDDEV 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "devise"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1260
                  TabIndex        =   19
                  Top             =   -30
                  Width           =   615
               End
               Begin VB.Label lblSelect_DOSSLDNUM 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Dossier"
                  BeginProperty Font 
                     Name            =   "Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   60
                  TabIndex        =   11
                  Top             =   -30
                  Width           =   615
               End
            End
            Begin VB.Label lblSelect_DOSSLDPCI 
               BackColor       =   &H00D0F0FF&
               Caption         =   "PCI"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   4290
               TabIndex        =   23
               Top             =   180
               Width           =   495
            End
            Begin VB.Label lblSelect_DOSSLDOPE 
               BackColor       =   &H00F0FFFF&
               Caption         =   "opération"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2790
               TabIndex        =   21
               Top             =   150
               Width           =   975
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   3795
            Left            =   4350
            TabIndex        =   12
            Top             =   1935
            Visible         =   0   'False
            Width           =   11595
            _ExtentX        =   20452
            _ExtentY        =   6694
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16777215
            ForeColor       =   16384
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483633
            BackColorBkg    =   -2147483633
            FormatString    =   $"SAB_Dossier.frx":0C07
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgCPTPIE 
         Height          =   3000
         Left            =   -74715
         TabIndex        =   49
         Top             =   3285
         Visible         =   0   'False
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   5292
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   16777215
         ForeColor       =   4210752
         BackColorFixed  =   13693183
         ForeColorFixed  =   16384
         BackColorBkg    =   -2147483633
         AllowUserResizing=   3
         FormatString    =   $"SAB_Dossier.frx":0CC4
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   15810
      Picture         =   "SAB_Dossier.frx":0DBE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   15
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
   Begin VB.Menu mnuPrint 
      Caption         =   "mnuPrint"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint_2_Liste 
         Caption         =   "Imprimer la liste"
      End
      Begin VB.Menu mnuPrint_2_Exportation 
         Caption         =   "Rapprochement Compta / Gestion .xlsx"
      End
   End
   Begin VB.Menu mnuDoc 
      Caption         =   "mnuDoc"
      Visible         =   0   'False
      Begin VB.Menu mnuDoc_Display 
         Caption         =   "Afficher le document"
      End
      Begin VB.Menu mnuDoc_Delete 
         Caption         =   "Supprimer le document"
      End
      Begin VB.Menu mnuDoc_Rename 
         Caption         =   "Renommer le document"
      End
   End
   Begin VB.Menu mnuPrint2 
      Caption         =   "mnuPrint2"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint2_Excel 
         Caption         =   "Excel"
      End
      Begin VB.Menu mnuPrint2_Mail 
         Caption         =   "envoi Mail"
      End
   End
End
Attribute VB_Name = "frmSAB_Dossier"
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
'$BIA_VB_HAB Dim SAB_Dossier_Aut As typeAuthorization
Dim arrHab(19) As Boolean
Dim blnAuto As Boolean, blnError As Boolean
Dim cmdSelect_SQL_K As String
Dim rsSabX As New ADODB.Recordset

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean
Dim fgSelect_Width As Integer, fgSelect_Height As Integer
Dim fgSelect_BackColorFixed As Long, fgSelect_ForeColorFixed As Long, fgSelect_ForeColor As Long, fgSelect_BackColor As Long


'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim xAmjMin As String, xAmjMax As String
Dim wDIBM_Min As Long, wDIBM_Max As Long
Dim wDMS_Min As String, wDMS_Max As String
Dim xYBIACPT0 As typeYBIACPT0, newYBIACPT0 As typeYBIACPT0, oldYBIACPT0 As typeYBIACPT0
Dim arrYBIACPT0() As typeYBIACPT0, arrYBIACPT0_Nb As Long, arrYBIACPT0_Max As Long, arrYBIACPT0_Index As Long

Dim xYDOSSLD0 As typeYDOSSLD0, newYDOSSLD0 As typeYDOSSLD0, oldYDOSSLD0 As typeYDOSSLD0
Dim arrYDOSSLD0() As typeYDOSSLD0, arrYDOSSLD0_Nb As Long, arrYDOSSLD0_Max As Long, arrYDOSSLD0_Index As Long
Dim arrYDOSSLD0_K As Long

Dim xYDOSSLD1 As typeYDOSSLD1, newYDOSSLD1 As typeYDOSSLD1, oldYDOSSLD1 As typeYDOSSLD1
Dim arrYDOSSLD1() As typeYDOSSLD1, arrYDOSSLD1_Nb As Long, arrYDOSSLD1_Max As Long, arrYDOSSLD1_Index As Long

Dim xYDOSMVT0 As typeYDOSMVT0, newYDOSMVT0 As typeYDOSMVT0, oldYDOSMVT0 As typeYDOSMVT0
Dim arrYDOSMVT0() As typeYDOSMVT0, arrYDOSMVT0_Nb As Long, arrYDOSMVT0_Max As Long, arrYDOSMVT0_Index As Long

Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean



Dim fgBIAMVT_FormatString As String, fgBIAMVT_K As Integer
Dim fgBIAMVT_RowDisplay As Integer, fgBIAMVT_RowClick As Integer, fgBIAMVT_ColClick As Integer
Dim fgBIAMVT_ColorClick As Long, fgBIAMVT_ColorDisplay As Long
Dim fgBIAMVT_Sort1 As Integer, fgBIAMVT_Sort2 As Integer
Dim fgBIAMVT_SortAD As Integer, fgBIAMVT_Sort1_Old As Integer
Dim fgBIAMVT_arrIndex As Integer
Dim blnfgBIAMVT_DisplayLine As Boolean

Dim xYBIAMVTH As typeYBIAMVT0, newYBIAMVTH As typeYBIAMVT0, oldYBIAMVTH As typeYBIAMVT0
Dim arrYBIAMVTH() As typeYBIAMVT0, arrYBIAMVTH_Nb As Long, arrYBIAMVTH_Max As Long, arrYBIAMVTH_Index As Long

Dim fgCPTPIE_FormatString As String, fgCPTPIE_K As Integer
Dim fgCPTPIE_RowDisplay As Integer, fgCPTPIE_RowClick As Integer, fgCPTPIE_ColClick As Integer
Dim fgCPTPIE_ColorClick As Long, fgCPTPIE_ColorDisplay As Long
Dim fgCPTPIE_Sort1 As Integer, fgCPTPIE_Sort2 As Integer
Dim fgCPTPIE_SortAD As Integer, fgCPTPIE_Sort1_Old As Integer
Dim fgCPTPIE_arrIndex As Integer
Dim blnfgCPTPIE_DisplayLine As Boolean

Dim xYCPTPIEH As typeYBIAMVT0, newYCPTPIEH As typeYBIAMVT0, oldYCPTPIEH As typeYBIAMVT0
Dim arrYCPTPIEH() As typeYBIAMVT0, arrYCPTPIEH_Nb As Long, arrYCPTPIEH_Max As Long, arrYCPTPIEH_Index As Long

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim arrDev() As String, arrDev_Nb As Integer, arrDev_RowT() As Long, arrDev_Cours() As Double
Dim arrOPE() As String, arrOPE_Nb As Integer, arrOPE_K As Integer

Dim fgDossier_FormatString As String, fgDossier_K As Integer
Dim fgDossier_RowDisplay As Integer, fgDossier_RowClick As Integer, fgDossier_ColClick As Integer
Dim fgDossier_ColorClick As Long, fgDossier_ColorDisplay As Long
Dim fgDossier_Sort1 As Integer, fgDossier_Sort2 As Integer
Dim fgDossier_SortAD As Integer, fgDossier_Sort1_Old As Integer
Dim fgDossier_arrIndex As Integer
Dim blnfgDossier_DisplayLine As Boolean

Dim fgCourrier_FormatString As String, fgCourrier_K As Integer
Dim fgCourrier_RowDisplay As Integer, fgCourrier_RowClick As Integer, fgCourrier_ColClick As Integer
Dim fgCourrier_ColorClick As Long, fgCourrier_ColorDisplay As Long
Dim fgCourrier_Sort1 As Integer, fgCourrier_Sort2 As Integer
Dim fgCourrier_SortAD As Integer, fgCourrier_Sort1_Old As Integer
Dim fgCourrier_arrIndex As Integer
Dim blnfgCourrier_DisplayLine As Boolean

Dim fgScan_FormatString As String, fgScan_K As Integer
Dim fgScan_RowDisplay As Integer, fgScan_RowClick As Integer, fgScan_ColClick As Integer
Dim fgScan_ColorClick As Long, fgScan_ColorDisplay As Long
Dim fgScan_Sort1 As Integer, fgScan_Sort2 As Integer
Dim fgScan_SortAD As Integer, fgScan_Sort1_Old As Integer
Dim fgScan_arrIndex As Integer
Dim blnfgScan_DisplayLine As Boolean

Dim mCDOMODDMO As Long
Dim arrZCDODOS0() As typeZCDODOS0, arrZCDODOS0_Nb As Long, arrZCDODOS0_K As Long, blnZCDODOS0 As Boolean
Dim sqlPCI As String, sqlPCI_Len As Integer, sqlCLI As String
Dim mXls1_Row As Long, mXls1_Col As Long, mXls2_Row As Long, mXls2_Col As Long, mXls2_Row_Cli As Long
Dim mMTD0_Cli As Currency, mMTD9_Cli As Currency, mMTD0_Dos As Currency, mMTD9_Dos As Currency, mMTDJ_Dos As Currency
Dim mCDODOSDOS_Nb As Long, mCDODOSPCC_Nb As Long, mCDODOSPDE_Nb As Long, mCDODOSXXX_Nb As Long, mDOSSLDMSD_Nb As Long, mProv_Nb As Long
Dim mCDODOSVAL_Nb As Long
Dim mAnn_Nb As Long
Dim blnDos_Ok As Boolean, blnCli_Ok As Boolean
Dim wMTD As Currency, wPIE As Long, wECR As Long


Dim blnProvisions_Control As Boolean, mProvisions_Control_Ope As String, mProvisions_Control_PCI As String
Dim wFilex As String, wFile As String

Dim xYDOSCD70 As typeYDOSCD70, newYDOSCD70 As typeYDOSCD70, oldYDOSCD70 As typeYDOSCD70
Dim dosYDOSCD70 As typeYDOSCD70

Dim sMTD_Solde_C As Currency, sMTD_COM_C As Currency, sMTD_UTI_G As Currency
Dim sMTD_COM_G2 As Currency, sMTD_COM_G2Prata As Currency, sMTD_COM_G2PDIF As Currency, sMTD_COM_G3 As Currency
Dim sMTD_TC2 As Currency
Dim tMTD_Solde_C As Currency, tMTD_COM_C As Currency, tMTD_UTI_G As Currency
Dim tMTD_COM_G2 As Currency, tMTD_COM_G2Prata As Currency, tMTD_COM_G2PDIF As Currency, tMTD_COM_G3 As Currency
Dim mDev_R1 As Long, mDev_R2 As Long
Dim blnCDODOSANN As Boolean
Dim mSOLDE_K As Integer
Dim mXls1_Row_C As Long, mXls1_Row_N As Long, mXls1_Row_D As Long, mXls1_Row_T As Long, mXls1_Col_EUR As Long
Dim mXls1_Row_SP As Long, mXls2_Row_D As Long
Dim sMTD_COM_ANN_C As Currency, sMTD_COM_ANN_N As Currency
'Dim  As Long
Dim mHeader_xls As String

Dim xYDOSXOD0 As typeYDOSXOD0, newYDOSXOD0 As typeYDOSXOD0, oldYDOSXOD0 As typeYDOSXOD0

Dim fgLog_FormatString As String, fgLog_K As Integer
Dim fgLog_RowDisplay As Integer, fgLog_RowClick As Integer, fgLog_ColClick As Integer
Dim fgLog_ColorClick As Long, fgLog_ColorDisplay As Long
Dim fgLog_Sort1 As Integer, fgLog_Sort2 As Integer
Dim fgLog_SortAD As Integer, fgLog_Sort1_Old As Integer
Dim fgLog_arrIndex As Integer
Dim blnfgLog_DisplayLine As Boolean


Dim fgYSWISAB0_FormatString As String, fgYSWISAB0_K As Integer
Dim fgYSWISAB0_RowDisplay As Integer, fgYSWISAB0_RowClick As Integer, fgYSWISAB0_ColClick As Integer
Dim fgYSWISAB0_ColorClick As Long, fgYSWISAB0_ColorDisplay As Long
Dim fgYSWISAB0_Sort1 As Integer, fgYSWISAB0_Sort2 As Integer
Dim fgYSWISAB0_SortAD As Integer, fgYSWISAB0_Sort1_Old As Integer
Dim fgYSWISAB0_arrIndex As Integer
Dim blnfgYSWISAB0_DisplayLine As Boolean

Dim xYSWISAB0 As typeYSWISAB0
Dim cnSIDE_DB As New ADODB.Connection, rsSIDE_DB As New ADODB.Recordset
Dim blnSIDE_DB As Boolean
Dim fgSwift_FormatString As String

Dim blnBIAMVT As Boolean
Dim xrText As typerText
Dim oldYSWISAB0 As typeYSWISAB0

Dim arrRow() As Long, arrRow_Err() As Long

Dim arrCDOMODEVE_07() As Long, arrCDOMODEVE_07_Nb As Long

Dim mRacinesExclues As String, LFB_RacinesExclues As String

Dim blnYSWILNK0_Display As Boolean, mYSWILNK0_Display As String

Dim mDoc_Filename As String
Dim mCDODOSOUV_11001 As Long, mCDODOSOUV_11012 As Long

Dim arrECNFPT_DOS() As Long, arrECNFPT_DOS_Nb As Long

Dim blnZCAUDOS0_S01 As Boolean


Dim fgX_FormatString As String, fgX_K As Integer
Dim fgX_RowDisplay As Integer, fgX_RowClick As Integer, fgX_ColClick As Integer
Dim fgX_ColorClick As Long, fgX_ColorDisplay As Long
Dim fgX_Sort1 As Integer, fgX_Sort2 As Integer
Dim fgX_SortAD As Integer, fgX_Sort1_Old As Integer
Dim fgX_arrIndex As Integer
Dim blnfgX_DisplayLine As Boolean


Dim fgCOM_FormatString As String, fgCOM_K As Integer
Dim fgCOM_RowDisplay As Integer, fgCOM_RowClick As Integer, fgCOM_ColClick As Integer
Dim fgCOM_ColorClick As Long, fgCOM_ColorDisplay As Long
Dim fgCOM_Sort1 As Integer, fgCOM_Sort2 As Integer
Dim fgCOM_SortAD As Integer, fgCOM_Sort1_Old As Integer
Dim fgCOM_arrIndex As Integer
Dim blnfgCOM_DisplayLine As Boolean
Dim mCDOCOMMON As Currency, mCDOCOMDOS As Long

Dim mECNFPT_Row As Integer
Dim mECNFPT_MTA As Currency, mECNFPT_TX1 As Double, mECNFPT_PER As String
Dim mECNFPT_DDEB As Long, mECNFPT_DFIN As Long, mECNFPT_DREG As Long, mECNFPT_MON As Currency
Dim mECNFPT_NBJ As Long, mECNFPT_NBJ_X As Long, mECNFPT_Ratio As Long
Dim mECNFPT_MIN As Currency, mECNFPT_TOT As Currency

Dim arrYDOSCD70() As typeYDOSCD70, arrYDOSCD70_Nb As Integer, arrYDOSCD70_Max As Integer
Dim blnECNFPT_CD7 As Boolean

Dim arrZAUTENA0() As typeZAUTENA0, arrZAUTENA0_Nb As Integer, arrZAUTENA0_Max As Integer
Dim xZAUTENA0 As typeZAUTENA0
Public Sub cmdPrint_YDOSSLD0(X As String)
'prtYDOSSLD0_Init "YDOSSLD0", X
'prtYEICGCC0_Open
'For I = 1 To fgSelect.Rows - 1
'    fgSelect.Row = I
'    fgSelect.Col = 0: xYDOSSLD0.DOSSLDDEV = Trim(fgSelect.Text)
'    fgSelect.Col = 1: xYDOSSLD0.DOSSLDOPE = Trim(fgSelect.Text)
'     fgSelect.Col = 2: xYDOSSLD0.DOSSLDNUM = Val(Trim(fgSelect.Text))
'    prtYDOSSLD0_Line arrYDOSSLD0(I)
'Next I
'prtYDOSSLD0_Close True

End Sub

Public Sub cmdSelect_SQL_Xi_Dossier_Echu(lSheet As Integer, lDOSSLDCLI As String, lDOSSLDPCI As String, lCLIENARA1 As String)
On Error GoTo Error_Handler
Const dateEcheanceLimite As Long = 1181231
Dim xSql As String
Dim X As String, K As Integer
Dim DAmjD As Long, DAmjF As Long, Nb1 As Long, Nb2 As Long, xCur As Currency

    '__________________________________________________________________________________
    Call lstErr_AddItem(lstErr, cmdContext, "> Exportation ........ " & lSheet & "-" & lDOSSLDCLI): DoEvents
    Set wsExcel = wbExcel.Sheets(lSheet)
    Call rsYDOSCD70_Init(oldYDOSCD70)
    If lDOSSLDCLI = "0011001" Then
        If lSheet = 10 Then
            xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDODOS0 " _
                 & " where DOSSLDPCI like '" & lDOSSLDPCI & "%' and DOSSLDCLI = '" & lDOSSLDCLI & "'" _
                 & " and DOSSLDSTA not in ('  ','80','90')" _
                 & " and CDODOSCOP = DOSSLDOPE and CDODOSDOS = DOSSLDNUM" _
                 & " and CDODOSOUV >= " & mCDODOSOUV_11001 _
                 & " and CDODOSVAL > " & dateEcheanceLimite _
                 & " order by DOSSLDDEV,DOSSLDOPE,DOSSLDNUM"
        Else
            xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDODOS0 " _
                 & " where DOSSLDPCI like '" & lDOSSLDPCI & "%' and DOSSLDCLI = '" & lDOSSLDCLI & "'" _
                 & " and DOSSLDSTA not in ('  ','80','90')" _
                 & " and CDODOSCOP = DOSSLDOPE and CDODOSDOS = DOSSLDNUM" _
                 & " and CDODOSOUV >= " & mCDODOSOUV_11001 _
                 & " and CDODOSVAL < " & wDIBM_Min _
                 & " order by DOSSLDDEV,DOSSLDOPE,DOSSLDNUM"
        End If
    Else
        If InStr(LFB_RacinesExclues, lDOSSLDCLI) > 0 Then
            If lSheet = 10 Then
                xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDODOS0 " _
                     & " where DOSSLDPCI like '" & lDOSSLDPCI & "%' and DOSSLDCLI = '" & lDOSSLDCLI & "'" _
                     & " and DOSSLDSTA not in ('  ','80','90')" _
                     & " and CDODOSCOP = DOSSLDOPE and CDODOSDOS = DOSSLDNUM" _
                     & " and CDODOSOUV >= " & mCDODOSOUV_11012 _
                     & " and CDODOSVAL > " & dateEcheanceLimite _
                     & " order by DOSSLDDEV,DOSSLDOPE,DOSSLDNUM"
            Else
                xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDODOS0 " _
                     & " where DOSSLDPCI like '" & lDOSSLDPCI & "%' and DOSSLDCLI = '" & lDOSSLDCLI & "'" _
                     & " and DOSSLDSTA not in ('  ','80','90')" _
                     & " and CDODOSCOP = DOSSLDOPE and CDODOSDOS = DOSSLDNUM" _
                     & " and CDODOSOUV >= " & mCDODOSOUV_11012 _
                     & " and CDODOSVAL < " & wDIBM_Min _
                     & " order by DOSSLDDEV,DOSSLDOPE,DOSSLDNUM"
            End If
        Else
            If lSheet = 10 Then
                xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDODOS0 " _
                     & " where DOSSLDPCI like '" & lDOSSLDPCI & "%' and DOSSLDCLI = '" & lDOSSLDCLI & "'" _
                     & " and DOSSLDSTA not in ('  ','80','90')" _
                     & " and CDODOSCOP = DOSSLDOPE and CDODOSDOS = DOSSLDNUM" _
                     & " and CDODOSVAL > " & dateEcheanceLimite _
                     & " order by DOSSLDDEV,DOSSLDOPE,DOSSLDNUM"
            Else
                xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDODOS0 " _
                     & " where DOSSLDPCI like '" & lDOSSLDPCI & "%' and DOSSLDCLI = '" & lDOSSLDCLI & "'" _
                     & " and DOSSLDSTA not in ('  ','80','90')" _
                     & " and CDODOSCOP = DOSSLDOPE and CDODOSDOS = DOSSLDNUM" _
                     & " and CDODOSVAL < " & wDIBM_Min _
                     & " order by DOSSLDDEV,DOSSLDOPE,DOSSLDNUM"
            End If
        End If
    End If
    Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF

    V = rsYDOSSLD0_GetBuffer(rsSab, xYDOSSLD0)
        
    If xYDOSSLD0.DOSSLDDEV <> oldYDOSSLD0.DOSSLDDEV Then
        Call cmdSelect_SQL_Xi_Dossier_T(lCLIENARA1, lDOSSLDPCI)
    End If
    mXls2_Row = mXls2_Row + 1
    DAmjD = rsSab("CDODOSOUV") + 19000000
    DAmjF = rsSab("CDODOSVAL") + 19000000
    
    wsExcel.Cells(mXls2_Row, 1) = xYDOSSLD0.DOSSLDOPE
    wsExcel.Cells(mXls2_Row, 2) = xYDOSSLD0.DOSSLDNUM
    wsExcel.Cells(mXls2_Row, 3) = rsSab("CDODOSCON")
    wsExcel.Cells(mXls2_Row, 4) = xYDOSSLD0.DOSSLDCLI
    wsExcel.Cells(mXls2_Row, 5) = dateImp10(DAmjD)
    wsExcel.Cells(mXls2_Row, 6) = dateImp10(DAmjF)
    wsExcel.Cells(mXls2_Row, 7) = xYDOSSLD0.DOSSLDMSD
    wsExcel.Cells(mXls2_Row, 8) = xYDOSSLD0.DOSSLDDEV
    '//////////////////////////////////
    If xYDOSSLD0.DOSSLDDEV <> "EUR" Then
        wsExcel.Cells(mXls2_Row, 9) = CDbl(xYDOSSLD0.DOSSLDMSD) / arrDev_Cours(retourne_indice_devise(xYDOSSLD0.DOSSLDDEV))
    Else
        wsExcel.Cells(mXls2_Row, 9) = xYDOSSLD0.DOSSLDMSD
    End If
    '//////////////////////////////////
    If DAmjF <= wAmjMin Then
        wsExcel.Cells(mXls2_Row, 6).Font.Color = vbMagenta
        Nb1 = DateDiff("d", dateImp_Amj(DAmjD), wDMS_Min) + 1
    Else
        Nb1 = DateDiff("d", dateImp_Amj(DAmjD), dateImp_Amj(DAmjF)) + 1
    End If
    wsExcel.Cells(mXls2_Row, 10) = Nb1
    If Nb1 >= 93 Then
        If xYDOSSLD0.DOSSLDDEV <> "EUR" Then
            wsExcel.Cells(mXls2_Row, 11) = xYDOSSLD0.DOSSLDMSD
            wsExcel.Cells(mXls2_Row, 12) = CDbl(xYDOSSLD0.DOSSLDMSD) / arrDev_Cours(retourne_indice_devise(xYDOSSLD0.DOSSLDDEV))
        Else
            wsExcel.Cells(mXls2_Row, 11) = xYDOSSLD0.DOSSLDMSD
            wsExcel.Cells(mXls2_Row, 12) = xYDOSSLD0.DOSSLDMSD
        End If
    Else
        If xYDOSSLD0.DOSSLDDEV <> "EUR" Then
            wsExcel.Cells(mXls2_Row, 13) = xYDOSSLD0.DOSSLDMSD
            wsExcel.Cells(mXls2_Row, 14) = CDbl(xYDOSSLD0.DOSSLDMSD) / arrDev_Cours(retourne_indice_devise(xYDOSSLD0.DOSSLDDEV))
        Else
            wsExcel.Cells(mXls2_Row, 13) = xYDOSSLD0.DOSSLDMSD
            wsExcel.Cells(mXls2_Row, 14) = xYDOSSLD0.DOSSLDMSD
        End If
    End If
    rsSab.MoveNext
Loop

xYDOSSLD0.DOSSLDDEV = ""

Call cmdSelect_SQL_Xi_Dossier_T(lCLIENARA1, lDOSSLDPCI)

Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & lSheet & "-" & lDOSSLDCLI): DoEvents


'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & lDOSSLDCLI): DoEvents

End Sub






Public Sub cmdSelect_SQL_Xi_Init_Echu(lSheet As Integer)
Dim K As Integer, K2 As Integer
On Error Resume Next
Set wsExcel = wbExcel.Sheets(lSheet)


With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignRight
    .WrapText = False ' True
    .Font.Size = 9
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 65
'wsExcel.PageSetup.PrintTitleRows = "$A1:$L1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
Select Case lSheet
    Case 7
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14" & mHeader_xls & ", arrêté au " & dateImp10(wAmjMin) _
                                        & vbCr & "&B&U&10(crédits documentaires export échus)"
    Case 9
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14" & mHeader_xls & ", arrêté au " & dateImp10(wAmjMin) _
                                        & vbCr & "&B&U&10(crédits documentaires export échéance sup 31/12/2018)"
End Select

wsExcel.PageSetup.CenterHorizontally = True

wsExcel.Columns(1).ColumnWidth = 7:  wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 30: wsExcel.Cells(mXls1_Row_C, 2) = "Intitulé":  wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignLeft

If lSheet = 7 Then
    wsExcel.Cells(mXls1_Row_C, 1) = "ECHUS"
    wsExcel.Cells(mXls1_Row_C, 1).Interior.Color = mColor_Y1
    wsExcel.Cells(mXls1_Row_C, 2).Interior.Color = mColor_GB: wsExcel.Cells(mXls1_Row_C, 2).Font.Color = mColor_Z0
    wsExcel.Cells(mXls1_Row_N - 1, 1) = "T ECHUS"
End If
If lSheet = 9 Then
    wsExcel.Cells(mXls1_Row_C, 1) = "A ECHOIR"
    wsExcel.Cells(mXls1_Row_C, 1).Interior.Color = mColor_Y1
    wsExcel.Cells(mXls1_Row_C, 2).Interior.Color = mColor_GB: wsExcel.Cells(mXls1_Row_C, 2).Font.Color = mColor_Z0
    wsExcel.Cells(mXls1_Row_N - 1, 1) = "T A ECHOIR"
End If
wsExcel.Cells(mXls1_Row_N + 1, 1) = "Total" 'mXls1_Row_T + 1
wsExcel.Cells(mXls1_Row_N + 1, 1).Interior.Color = mColor_GB 'mXls1_Row_T + 1
wsExcel.Cells(mXls1_Row_N + 1, 2).Interior.Color = mColor_GB 'mXls1_Row_T + 1
wsExcel.Cells(mXls1_Row_N + 1, 2).Font.Color = mColor_Z0
wsExcel.Cells(mXls1_Row_N + 2, 1) = "Cours" 'mXls1_Row_T + 1
wsExcel.Cells(mXls1_Row_N + 3, 1) = "Total dev.":
wsExcel.Cells(mXls1_Row_N + 4, 1) = " € >= 93 J ":
wsExcel.Cells(mXls1_Row_N + 5, 1) = " € < 93 J ":
wsExcel.Cells(mXls1_Row_N + 6, 1) = "Total €":

wsExcel.Cells(mXls1_Row_N, 1).Interior.Color = mColor_GB
wsExcel.Cells(mXls1_Row_N, 2).Interior.Color = mColor_GB

For K = 1 To arrDev_Nb
    K2 = K + 2
    wsExcel.Cells(mXls1_Row_C, K2) = arrDev(K)
    wsExcel.Cells(mXls1_Row_C, K2).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_C, K2).Font.Color = mColor_Z0
    wsExcel.Cells(mXls1_Row_N, K2) = arrDev(K)
    wsExcel.Cells(mXls1_Row_N, K2).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_N, K2).Font.Color = mColor_Z0
    wsExcel.Columns(K2).ColumnWidth = 13: wsExcel.Columns(K2).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
    
    If arrDev(K) = "EUR" Or arrDev(K) = "USD" Then wsExcel.Columns(K2).ColumnWidth = 16
    wsExcel.Cells(mXls1_Row_N + 2, K2).NumberFormat = "### ##0.00000"
    wsExcel.Cells(mXls1_Row_N + 2, K2) = arrDev_Cours(K)
Next K

For K = 1 To mXls1_Col
    wsExcel.Cells(mXls1_Row_N - 1, K).Interior.Color = mColor_G0
    wsExcel.Cells(mXls1_Row_N + 1, K).Interior.Color = mColor_G0
    wsExcel.Cells(mXls1_Row_N + 3, K).Interior.Color = mColor_Y0
    wsExcel.Cells(mXls1_Row_N + 4, K).Interior.Color = mColor_Y0
    wsExcel.Cells(mXls1_Row_N + 5, K).Interior.Color = mColor_Y0
    wsExcel.Cells(mXls1_Row_N + 6, K).Interior.Color = mColor_Y0
Next K

End Sub



'______________________________________________________________________
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

Do While Not rsSab.EOF
    blnOk = True
    If chkSelect_DOSSLDSVC = "1" Then
        If rsSab("DOSSLDSVC") = "01" And rsSab("DOSSLDSTA") = "01" Then blnOk = False
    End If
    If blnOk Then

         fgSelect.Rows = fgSelect.Rows + 1
         fgSelect.Row = fgSelect.Rows - 1
         fgSelect.Col = 0: fgSelect.Text = rsSab("DOSSLDDEV")
         fgSelect.Col = 1: fgSelect.Text = rsSab("DOSSLDOPE")
         fgSelect.Col = 2: fgSelect.Text = rsSab("DOSSLDNUM")
         fgSelect.Col = 3: fgSelect.Text = rsSab("DOSSLDCLI")
         
        
         fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = I
    End If
    rsSab.MoveNext
Loop

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgSelect_Display_ZCDODOS0()
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
Call rsYDOSSLD0_Init(oldYDOSSLD0)
Do While Not rsSab.EOF
    blnOk = True
    If chkSelect_DOSSLDSVC = "1" Then
        If rsSab("CDODOSEVE") = "01" And rsSab("CDODOSETA") = "01" Then blnOk = False
    End If
    If blnOk Then

         fgSelect.Rows = fgSelect.Rows + 1
         fgSelect.Row = fgSelect.Rows - 1
         fgSelect.Col = 0: fgSelect.Text = rsSab("CDODOSDEV")
         fgSelect.Col = 1: fgSelect.Text = rsSab("CDODOSCOP")
         fgSelect.Col = 2: fgSelect.Text = rsSab("CDODOSDOS")
         fgSelect.Col = 3: fgSelect.Text = rsSab("CDODOSNOT")
         
        oldYDOSSLD0.DOSSLDOPE = rsSab("CDODOSCOP")
        oldYDOSSLD0.DOSSLDNUM = rsSab("CDODOSDOS")
        oldYDOSSLD0.DOSSLDDEV = rsSab("CDODOSDEV")
        If oldYDOSSLD0.DOSSLDOPE = "CDI" Then
            oldYDOSSLD0.DOSSLDCLI = rsSab("CDODOSDON")
        Else
            oldYDOSSLD0.DOSSLDCLI = rsSab("CDODOSNOT")
        End If
        oldYDOSSLD0.DOSSLDSTA = rsSab("CDODOSEVE")
        oldYDOSSLD0.DOSSLDSVC = rsSab("CDODOSETA")

         fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = I
    End If
    rsSab.MoveNext
Loop

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgSelect_Display_ZENCCAR0()
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
Call rsYDOSSLD0_Init(oldYDOSSLD0)
Do While Not rsSab.EOF
    blnOk = True
    If chkSelect_DOSSLDSVC = "1" Then
        If rsSab("ENCCARETA") = "01" Then blnOk = False
    End If
    If blnOk Then

         fgSelect.Rows = fgSelect.Rows + 1
         fgSelect.Row = fgSelect.Rows - 1
         fgSelect.Col = 0: fgSelect.Text = rsSab("ENCCARDEV")
         fgSelect.Col = 1: fgSelect.Text = rsSab("ENCCARCOP")
         fgSelect.Col = 2: fgSelect.Text = rsSab("ENCCARDOS")
         fgSelect.Col = 3: fgSelect.Text = rsSab("ENCCARORD")
         
        oldYDOSSLD0.DOSSLDOPE = rsSab("ENCCARCOP")
        oldYDOSSLD0.DOSSLDNUM = rsSab("ENCCARDOS")
        oldYDOSSLD0.DOSSLDDEV = rsSab("ENCCARDEV")
        'If oldYDOSSLD0.DOSSLDOPE = "CDI" Then
        '    oldYDOSSLD0.DOSSLDCLI = rsSab("ENCCARDON")
        'Else
        '    oldYDOSSLD0.DOSSLDCLI = rsSab("ENCCARNOT")
        'End If
        oldYDOSSLD0.DOSSLDCLI = rsSab("ENCCARORD")
        
        'oldYDOSSLD0.DOSSLDSTA = rsSab("ENCCAREVE")
        oldYDOSSLD0.DOSSLDSVC = rsSab("ENCCARETA")

         fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = I
    End If
    rsSab.MoveNext
Loop

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgSelect_Display_3()
Dim X As String, xEVE As String, xETA As String
On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Row = 0

currentAction = "fgSelect_Display_3"

Do While Not rsSab.EOF

         fgSelect.Rows = fgSelect.Rows + 1
         fgSelect.Row = fgSelect.Rows - 1
         fgSelect.Col = 0: fgSelect.Text = rsSab("CDODOSDEV")
         fgSelect.Col = 1: fgSelect.Text = rsSab("CDODOSCOP")
         fgSelect.Col = 2: fgSelect.Text = rsSab("CDODOSDOS")
         xEVE = rsSab("CDODOSEVE")
         xETA = rsSab("CDODOSETA")
         Select Case xEVE
            Case "01": X = "OUV"
            Case "02": X = "MOD"
            Case "07": X = "REO"
            Case "80": X = "ANN"
            Case "90": X = "CLO"
            Case Else: X = xEVE
        End Select
         Select Case xETA
            Case "01": X = X & "-S"
            Case "02": X = X & "-V"
            Case "07": X = X & "-C"
            Case Else: X = X & "-" & xETA
        End Select
         fgSelect.Col = 3: fgSelect.Text = X
        
    rsSab.MoveNext
Loop

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgSelect_Display_5réfext(lFct As String)
Dim X As String, blnOk As Boolean
On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString

currentAction = "fgSelect_Display_5réf"

Do While Not rsSab.EOF
    blnOk = False
   X = "select CDODOSDEV ,CDODOSCOP , CDODOSDOS , CDODOSEXT , CDODOSOUV  from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
         & " where CDODOSCOP = 'CDE' and CDODOSEXT = '" & rsSab("CDODOSEXT") & "' and CDODOSEVE <> '90'" _
         & " order by CDODOSOUV desc"
         
    Set rsSabX = cnsab.Execute(X)
    Do While Not rsSabX.EOF
        If lFct = "" Then
            blnOk = True
        Else
            If rsSabX("CDODOSOUV") = Val(YBIATAB0_DIBM_CPT_J) Then blnOk = True
        End If
        
        If blnOk Then
             fgSelect.Rows = fgSelect.Rows + 1
             fgSelect.Row = fgSelect.Rows - 1
    
            fgSelect.Col = 0: fgSelect.Text = rsSabX("CDODOSDEV")
            fgSelect.Col = 1: fgSelect.Text = rsSabX("CDODOSCOP")
            fgSelect.Col = 2: fgSelect.Text = rsSabX("CDODOSDOS")
            fgSelect.Col = 3: fgSelect.Text = rsSabX("CDODOSEXT")
            fgSelect.CellForeColor = vbMagenta
        End If
        
        rsSabX.MoveNext
    Loop
    
    rsSab.MoveNext
Loop

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgSelect_Display_3uti()
Dim X As String, xEVE As String, xETA As String, K As Integer, wAmj As Long, mAMJ_7j As Long
On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Gestionnaire       |<Dossier                | Date prévue d'util" _
                      & "|< date MT999     |<Date remise       |<C/N/D|>Montant utilisation         |<Devise|<Date de validité|< Banque                                     " _
                      & "|Référence externe            "
fgSelect.Row = 0
fgSelect.Col = 2: fgSelect.CellAlignment = 2
fgSelect.Col = 6: fgSelect.CellAlignment = 1
currentAction = "fgSelect_Display_3uti"

mAMJ_7j = Val(dateElp("Jour", -7, YBIATAB0_DATE_CPT_J))
Do While Not rsSab.EOF

         fgSelect.Rows = fgSelect.Rows + 1
         fgSelect.Row = fgSelect.Rows - 1
         
        If fgSelect.Row Mod 2 = 0 Then
            For K = 0 To 10: fgSelect.Col = K: fgSelect.CellBackColor = RGB(240, 240, 240): Next K
        End If
         
         fgSelect.Col = 0: fgSelect.Text = rsSab("MNURUTUTI")
         wAmj = rsSab("CDOUTIPRE") + 19000000
         fgSelect.Col = 2: fgSelect.Text = "  " & dateAMJ10(wAmj)
                            fgSelect.CellForeColor = vbBlue
         If wAmj < mAMJ_7j Then
            fgSelect.CellBackColor = mColor_W0
        Else
            fgSelect.CellBackColor = mColor_Y0
        End If
        
         
         fgSelect.Col = 1: fgSelect.Text = rsSab("CDOUTICOP") & "  " & rsSab("CDOUTIDOS") & " - " & rsSab("CDOUTIUTI")
         'fgSelect.Col = 3: fgSelect.Text = rsSab("CDOUTIUTI")
         fgSelect.Col = 4: fgSelect.Text = "   " & dateAMJ10(rsSab("CDOUTIDRE") + 19000000)
         fgSelect.Col = 5: fgSelect.Text = "  " & rsSab("CDOUTITMO")
         fgSelect.Col = 6: fgSelect.Text = Format(rsSab("CDOUTIMON"), "### ### ### ##0.00")
         fgSelect.CellForeColor = vbBlue
         fgSelect.Col = 7: fgSelect.Text = "  " & rsSab("CDODOSDEV")
         fgSelect.CellForeColor = vbBlue
         fgSelect.Col = 8: fgSelect.Text = "    " & dateAMJ10(rsSab("CDODOSVAL") + 19000000)
         fgSelect.Col = 9: fgSelect.Text = "  " & rsSab("CDODOSNOT") & "  " & rsSab("CLIENASIG")
         fgSelect.Col = 10: fgSelect.Text = "  " & rsSab("CDODOSEXT")
         If chkSelect_Options_3uti_Swift = "1" Then
             X = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
             & " where SWISABOPEC = '" & rsSab("CDOUTICOP") & "'" _
             & " and SWISABOPEN = " & rsSab("CDOUTIDOS") _
             & " and SWISABWES = 'S' and SWISABWMTK in ( 799 , 999) order by SWISABWAMJ desc"
            Set rsSabX = cnsab.Execute(X)
            If Not rsSabX.EOF Then
                fgSelect.Col = 3: fgSelect.Text = dateAMJ10(rsSabX("SWISABWAMJ"))
                fgSelect.CellForeColor = vbMagenta
            End If
        End If
    rsSab.MoveNext
Loop

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgSelect_Display_5_AUT()
Dim X As String, K As Integer, I As Integer, mColor As Long
On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Opé |>Dossier       |<Code Autorisation           " _
                      & "|<Devise|>Consultation globale|>   SAB_Dossier     |<Code état|<PCI                                                        |<Client                           " _
                      & "|            "
fgSelect.Row = 0
fgSelect.Col = 1: fgSelect.CellAlignment = 2
fgSelect.Col = 4: fgSelect.CellAlignment = 1
fgSelect.Col = 5: fgSelect.CellAlignment = 1
fgSelect.Col = 6: fgSelect.CellAlignment = 2
currentAction = "fgSelect_Display_5_AUT"

For K = 1 To arrZAUTENA0_Nb
    If arrZAUTENA0(K).AUTENAENC <> arrZAUTENA0(K).DOSSLDMSD Then
         fgSelect.Rows = fgSelect.Rows + 1
         fgSelect.Row = fgSelect.Rows - 1
         
         mColor = 0
         
         fgSelect.Col = 0: fgSelect.Text = " " & arrZAUTENA0(K).AUTENAOPE
         fgSelect.Col = 1: fgSelect.Text = Format(arrZAUTENA0(K).AUTENADOS, "000000")
         fgSelect.Col = 2: fgSelect.Text = "_" & arrZAUTENA0(K).AUTENAAUT
         fgSelect.Col = 3: fgSelect.Text = " " & arrZAUTENA0(K).AUTENADEV
         fgSelect.Col = 4:
         fgSelect.Text = Format$(Abs(arrZAUTENA0(K).AUTENAENC), "### ### ### ##0.00")
        
         If arrZAUTENA0(K).AUTENAENC < 0 Then
            fgSelect.CellForeColor = vbRed
         Else
            fgSelect.CellForeColor = vbBlue
         End If

         
         fgSelect.Col = 5:
         fgSelect.Text = Format$(Abs(arrZAUTENA0(K).DOSSLDMSD), "### ### ### ##0.00")
        
         If arrZAUTENA0(K).DOSSLDMSD < 0 Then
            fgSelect.CellForeColor = vbRed
         Else
            fgSelect.CellForeColor = vbBlue
         End If
         fgSelect.Col = 6: fgSelect.Text = " " & arrZAUTENA0(K).DOSSLDSTA
         fgSelect.Col = 7: fgSelect.Text = " " & arrZAUTENA0(K).DOSSLDPCI
         fgSelect.Col = 8: fgSelect.Text = " " & arrZAUTENA0(K).AUTENACLI
         
        If arrZAUTENA0(K).DOSSLDSTA = "90" Then
            mColor = mColor_W1
        Else
            If arrZAUTENA0(K).DOSSLDPCI = "" Then
                mColor = mColor_W0
            Else
                If Mid$(arrZAUTENA0(K).AUTENAAUT, 1, 1) = "?" Then mColor = mColor_Y1
            End If
        End If
        If mColor > 0 Then
          For I = 0 To 8
               fgSelect.Col = I: fgSelect.CellBackColor = mColor
           Next I
        End If
    End If
Next K

If fgSelect.Rows > 1 Then fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgSelect_Display_3RDO()
Dim X As String, xEVE As String, xETA As String, K As Integer, wAmj As Long, mAMJ_7j As Long
On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Gestionnaire       |<Dossier                | Date prévue d'util" _
                      & "|< date MT999     |<Date remise       |<C/N/D|>Montant utilisation         |<Devise|<Date de validité|< Banque                                     " _
                      & "|Référence externe            "
fgSelect.Row = 0
fgSelect.Col = 2: fgSelect.CellAlignment = 2
fgSelect.Col = 6: fgSelect.CellAlignment = 1
currentAction = "fgSelect_Display_3uti"

mAMJ_7j = Val(dateElp("Jour", -7, YBIATAB0_DATE_CPT_J))
Do While Not rsSab.EOF

         fgSelect.Rows = fgSelect.Rows + 1
         fgSelect.Row = fgSelect.Rows - 1
         
        If fgSelect.Row Mod 2 = 0 Then
            For K = 0 To 10: fgSelect.Col = K: fgSelect.CellBackColor = RGB(240, 240, 240): Next K
        End If
         
         fgSelect.Col = 0: fgSelect.Text = rsSab("MNURUTUTI")
         fgSelect.Col = 2: fgSelect.CellForeColor = vbBlue
         If rsSab("ENCCAREC1") <> 0 Then
            wAmj = rsSab("ENCCAREC1") + 19000000
            fgSelect.Text = "  " & dateAMJ10(wAmj)
        Else
            wAmj = rsSab("ENCCARDAR") + 19000000
            fgSelect.Text = "  " & dateAMJ10(wAmj) & " à vue"
        End If
        
                            
         If wAmj < mAMJ_7j Then
            fgSelect.CellBackColor = mColor_W0
        Else
            fgSelect.CellBackColor = mColor_Y0
        End If
        
         
         fgSelect.Col = 1: fgSelect.Text = rsSab("ENCCARCOP") & "  " & rsSab("ENCCARDOS") '& " - " & rsSab("ENCCARREG")
         'fgSelect.Col = 3: fgSelect.Text = rsSab("ENCCARUTI")
         'fgSelect.Col = 4: fgSelect.Text = "   " & dateAMJ10(rsSab("ENCCARDRE") + 19000000)
         'fgSelect.Col = 5: fgSelect.Text = "  " & rsSab("ENCCARTMO")
         fgSelect.Col = 6: fgSelect.Text = Format(rsSab("ENCCARMON"), "### ### ### ##0.00")
         fgSelect.CellForeColor = vbBlue
         fgSelect.Col = 7: fgSelect.Text = "  " & rsSab("ENCCARDEV")
         fgSelect.CellForeColor = vbBlue
         'fgSelect.Col = 8: fgSelect.Text = "    " & dateAMJ10(rsSab("ENCCARDVA") + 19000000)
         'fgSelect.Col = 9: fgSelect.Text = "  " & rsSab("ENCCARNOT") & "  " & rsSab("CLIENASIG")
         'fgSelect.Col = 10: fgSelect.Text = "  " & rsSab("ENCCAREXT")
         If chkSelect_Options_3uti_Swift = "1" Then
             X = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
             & " where SWISABOPEC = '" & rsSab("ENCCARCOP") & "'" _
             & " and SWISABOPEN = " & rsSab("ENCCARDOS") _
             & " and SWISABWES = 'S' and SWISABWMTK in ( 420 ,499 , 999) order by SWISABWAMJ desc"
            Set rsSabX = cnsab.Execute(X)
            If Not rsSabX.EOF Then
                fgSelect.Col = 3: fgSelect.Text = dateAMJ10(rsSabX("SWISABWAMJ"))
                fgSelect.CellForeColor = vbMagenta
            End If
        End If
    rsSab.MoveNext
Loop

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub



Private Sub fgSelect_Display_2()
Dim wColor As Long
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Dev   |<PCI          |<Client           |"
fgSelect.Row = 0

currentAction = "fgSelect_Display"

Do While Not rsSab.EOF
    V = rsYDOSSLD1_GetBuffer(rsSab, xYDOSSLD1)
    
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect.Col = 0: fgSelect.Text = xYDOSSLD1.DOSSLDDEV
    fgSelect.Col = 1: fgSelect.Text = xYDOSSLD1.DOSSLDPCI
    fgSelect.Col = 2: fgSelect.Text = xYDOSSLD1.DOSSLDCLI
    If xYDOSSLD1.DOSSLDMSD <> xYDOSSLD1.DOSSLDGSD Then
        Select Case Mid$(xYDOSSLD1.DOSSLDPCI, 1, 5)
            Case "91120", "91122", "98050", "90312", "98520": fgSelect.CellBackColor = RGB(255, 192, 255)
            Case "70721", "91130", "91131": fgSelect.CellBackColor = RGB(238, 221, 255)
        End Select
    End If
    fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = I
    rsSab.MoveNext
Loop

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub arrYDOSSLD0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYDOSSLD0(101)
arrYDOSSLD0_Max = 100: arrYDOSSLD0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYDOSSLD0_GetBuffer(rsSab, xYDOSSLD0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYDOSSLD0.fgselect_Display"
        '' Exit Sub
     Else
         arrYDOSSLD0_Nb = arrYDOSSLD0_Nb + 1
         If arrYDOSSLD0_Nb > arrYDOSSLD0_Max Then
             arrYDOSSLD0_Max = arrYDOSSLD0_Max + 100
             ReDim Preserve arrYDOSSLD0(arrYDOSSLD0_Max)
         End If
         
         arrYDOSSLD0(arrYDOSSLD0_Nb) = xYDOSSLD0
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
Private Sub arrYDOSSLD1_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYDOSSLD1(101)
arrYDOSSLD1_Max = 100: arrYDOSSLD1_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD1 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYDOSSLD1_GetBuffer(rsSab, xYDOSSLD1)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYDOSSLD1.fgselect_Display"
        '' Exit Sub
     Else
         arrYDOSSLD1_Nb = arrYDOSSLD1_Nb + 1
         If arrYDOSSLD1_Nb > arrYDOSSLD1_Max Then
             arrYDOSSLD1_Max = arrYDOSSLD1_Max + 100
             ReDim Preserve arrYDOSSLD1(arrYDOSSLD1_Max)
         End If
         
         arrYDOSSLD1(arrYDOSSLD1_Nb) = xYDOSSLD1
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


Private Sub arrYBIACPT0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYBIACPT0(101)
arrYBIACPT0_Max = 100: arrYBIACPT0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYBIACPT0_GetBuffer(rsSab, xYBIACPT0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYBIACPT0.fgselect_Display"
        '' Exit Sub
     Else
         arrYBIACPT0_Nb = arrYBIACPT0_Nb + 1
         If arrYBIACPT0_Nb > arrYBIACPT0_Max Then
             arrYBIACPT0_Max = arrYBIACPT0_Max + 100
             ReDim Preserve arrYBIACPT0(arrYBIACPT0_Max)
         End If
         
         arrYBIACPT0(arrYBIACPT0_Nb) = xYBIACPT0
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


Private Sub arrYDOSMVT0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYDOSMVT0(101): arrYDOSMVT0_Max = 100: arrYDOSMVT0_Nb = 0
ReDim arrYBIAMVTH(101): arrYBIAMVTH_Max = 100: arrYBIAMVTH_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSMVT0 , " _
    & paramIBM_Library_SABSPE & ".YBIAMVTH " _
    & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYDOSMVT0_GetBuffer(rsSab, xYDOSMVT0)
    V = rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVTH)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYDOSMVT0.fgselect_Display"
        '' Exit Sub
     Else
         arrYDOSMVT0_Nb = arrYDOSMVT0_Nb + 1
         If arrYDOSMVT0_Nb > arrYDOSMVT0_Max Then
             arrYDOSMVT0_Max = arrYDOSMVT0_Max + 100
             ReDim Preserve arrYDOSMVT0(arrYDOSMVT0_Max)
             arrYBIAMVTH_Max = arrYBIAMVTH_Max + 100
             ReDim Preserve arrYBIAMVTH(arrYBIAMVTH_Max)
         End If
         
         arrYDOSMVT0(arrYDOSMVT0_Nb) = xYDOSMVT0
         arrYBIAMVTH(arrYDOSMVT0_Nb) = xYBIAMVTH
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




Public Sub cmdSelect_Reset()
If blnControl Then
    cmdSelect_Clear
    cmdSelect_Ok.Visible = False 'True
    fraSelect_Options_1.Visible = False
    fraSelect_Options_5.Visible = False
    fraSelect_Options_6.Visible = False
    fraSelect_Options_Xc.Visible = False
    fraSelect_Options_Log.Visible = False
    fraSelect_Options_3uti.Visible = False
    fraSelect_Options_Scan_Liste.Visible = False
    Dim K As Integer
    K = InStr(cboSelect_SQL, "-")
    If K > 0 Then
        cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, K - 1))
    Else
        cmdSelect_SQL_K = "?"
    End If
    fgSelect.Width = fgSelect_Width
    fgSelect.Height = fgSelect_Height
    fgSelect.BackColor = vbWhite
    fgX.Visible = False
    
    Select Case cmdSelect_SQL_K
        Case "1":
            lblSelect_DOSSLDNUM.Caption = "Dossier"
            fraSelect_Options.Visible = True: fraSelect_Options_1.Visible = True
            cmdSelect_Ok.Visible = True
        Case "2":
            'chkSelect_DOSSLDSTA.value = "1"
            lblSelect_DOSSLDNUM.Caption = "Client"
            fraSelect_Options.Visible = True: fraSelect_Options_1.Visible = True
            cmdSelect_Ok.Visible = True
        Case "2#":
            chkSelect_DOSSLDSTA.value = "0"
            chkSelect_DOSSLDSVC.value = "0"
            lblSelect_DOSSLDNUM.Caption = "Client"
            fraSelect_Options.Visible = True: fraSelect_Options_1.Visible = True
            cmdSelect_Ok.Visible = True
        Case "3uti":
            If cboSelect_Options_3uti_UTI.ListCount = 0 Then cboSelect_Options_3uti_UTI_Load
            fgSelect.Width = 12750
            fgSelect.Height = 7800
            fraSelect_Options.Visible = True: fraSelect_Options_3uti.Visible = True
            cmdSelect_Ok.Visible = True
        Case "3RDO":
            If cboSelect_Options_3uti_UTI.ListCount = 0 Then cboSelect_Options_3uti_UTI_Load_RDE
            fgSelect.Width = 12750
            fgSelect.Height = 7800
            fraSelect_Options.Visible = True: fraSelect_Options_3uti.Visible = True
            '                                               '
            lblSelect_Options_3uti_CDODOSNOT.Visible = False
            cboSelect_Options_3uti_CDODOSNOT.Visible = False
            '                                               '
            cmdSelect_Ok.Visible = True
        Case "5":
            fraSelect_Options_5.Visible = True: fraSelect_Options.Visible = False
            cmdSelect_Ok.Visible = True
         Case "6":
            fraSelect_Options_6.Visible = True: fraSelect_Options.Visible = False
            cmdSelect_Ok.Visible = True
         Case "Xc":
            fraSelect_Options_Xc.Visible = True: fraSelect_Options.Visible = False
            cmdSelect_Ok.Visible = True
         Case "zOD":
            fraSelect_Options_Log.Visible = True: fraSelect_Options.Visible = False
            cmdSelect_Ok.Visible = True
         Case "Scan_Liste":
            fraSelect_Options_Scan_Liste.Visible = True: fraSelect_Options.Visible = False
            cmdSelect_Ok.Visible = True
       Case "GAR_Ech"
            fraSelect_Options.Visible = False
            cmdSelect_Ok.Visible = True
                        
       Case Else
            fraSelect_Options.Visible = False
            cmdSelect_Ok.Visible = True
    End Select

End If

End Sub
Public Sub cboSelect_Options_3uti_UTI_Load_RDE()
Dim X As String
     
cboSelect_Options_3uti_UTI.Clear
cboSelect_Options_3uti_UTI.AddItem ""
X = "select MNURUTUTI , MNURUTCUT from " & paramIBM_Library_SAB & ".ZENCCAR0 , " _
                        & paramIBM_Library_SAB & ".ZMNURUT0  " _
     & " where ENCCARCET <'90' and ENCCARCOP = 'RDE'" _
     & " and ENCCARUL1 = MNURUTCUT" _
     & " group by MNURUTUTI , MNURUTCUT order by MNURUTUTI "

Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    cboSelect_Options_3uti_UTI.AddItem rsSab("MNURUTUTI") & " | " & rsSab("MNURUTCUT")
    rsSab.MoveNext
Loop

cboSelect_Options_3uti_CDODOSNOT.Clear
cboSelect_Options_3uti_CDODOSNOT.AddItem ""

End Sub

Public Sub cmdSelect_Clear()
    lstErr.Clear
    fgSelect.Visible = False
    fgDetail.Visible = False: fraDetail.Visible = False
    SSTab2.Visible = False
    lstW.Visible = False
    fraCompte.Visible = False
    fraYDOSXOD0.Visible = False
    fraDetail.Visible = False
    fgDossier.Visible = False: fgCourrier.Visible = False: fgScan.Visible = False
    fgYSWISAB0.Visible = False
    fraSwift.Visible = False
    fgLOG.Visible = False
    fgCOM.Visible = False
    mnuPrint_2_Exportation.Enabled = False
End Sub


Public Sub cmdDetail_Reset()
If blnControl Then
    lstErr.Clear
    If fgDetail.Visible Then
        fgDetail.Visible = False: fraDetail.Visible = False
        SSTab2.Visible = False
        fgCPTPIE.Visible = False
        fgDetail_Display
    End If
End If

End Sub


Private Sub cmdSelect_SQL_1()
Dim V, X As String
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean
Dim xDOSSLDM As String, xDOSSLDG As String, xDOSSLDK As String
Dim xField1 As String, xK As String, xField2 As String

On Error GoTo Error_Handler

currentAction = "cmdYDOSSLD0_SQL"
blnOk = False

arrYBIACPT0_Nb = 0

xWhere = ""
X = Trim(txtSelect_DOSSLDNUM)
If X <> "" Then
    xWhere = "   and DOSSLDNUM = " & Val(X)
    X = Trim(cboSelect_DOSSLDOPE)
    If X <> "" Then xWhere = xWhere & "   and DOSSLDOPE = '" & X & "'"
Else
    X = Trim(cboSelect_DOSSLDPCI)
    If X <> "" Then xWhere = xWhere & "   and DOSSLDPCI = '" & X & "'"
    If chkSelect_DOSSLDMG = "1" Then
        If X = "" Then
            Call MsgBox("Différence de solde, préciser le PCI (91120,91122,98050,90312)", vbCritical, "SAB_Dossier")
            Exit Sub
        Else
            xWhere = xWhere & " and (DOSSLDMSD <> DOSSLDGSD) "
        End If
    End If
    
    
    X = Trim(cboSelect_DOSSLDDEV)
    If X <> "" Then xWhere = xWhere & "   and DOSSLDDEV = '" & X & "'"
    X = Trim(cboSelect_DOSSLDOPE)
    If X <> "" Then xWhere = xWhere & "   and DOSSLDOPE = '" & X & "'"
    If chkSelect_DOSSLDSTA = "1" Then xWhere = xWhere & "   and DOSSLDSTA not in ('  ','80','90')"
    'If chkSelect_DOSSLDSVC = "1" Then xWhere = xWhere & "   and (DOSSLDSVC <> '01' and DOSSLDSTA <> '01')"
End If
If xWhere <> "" Then Mid$(xWhere, 1, 6) = " where"


xSql = "select distinct DOSSLDDEV , DOSSLDOPE , DOSSLDNUM ,DOSSLDCLI,DOSSLDSTA,DOSSLDSVC from " & paramIBM_Library_SABSPE & ".YDOSSLD0 " _
     & xWhere & " order by DOSSLDDEV , DOSSLDOPE , DOSSLDNUM"
Set rsSab = cnsab.Execute(xSql)


fgSelect_Display

If fgSelect.Rows = 2 Then
    fgSelect.Row = fgSelect.Rows - 1
     fgSelect.Col = 0: xYDOSSLD0.DOSSLDDEV = Trim(fgSelect.Text)
     fgSelect.Col = 1: xYDOSSLD0.DOSSLDOPE = Trim(fgSelect.Text)
     fgSelect.Col = 2: xYDOSSLD0.DOSSLDNUM = Val(Trim(fgSelect.Text))
    fgSelect.Col = 3: xYDOSSLD0.DOSSLDCLI = Trim(fgSelect.Text)
    fgDetail_Display
    If blnBIAMVT Then xYDOSSLD0 = oldYDOSSLD0: fgBIAMVT_Display: SSTab2.Tab = 2
Else
    If fgSelect.Rows = 1 Then cmdSelect_SQL_1New
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_1New()
Dim V, X As String, xCOP As String
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean
Dim xDOSSLDM As String, xDOSSLDG As String, xDOSSLDK As String
Dim xField1 As String, xK As String, xField2 As String

On Error GoTo Error_Handler

currentAction = "cmdYDOSSLD0_SQL"
blnOk = False

arrYBIACPT0_Nb = 0

xWhere = ""
xCOP = Trim(cboSelect_DOSSLDOPE)
X = Trim(txtSelect_DOSSLDNUM)
If X <> "" Then
    Select Case xCOP
        Case "CDE", "CDI"
            xWhere = " where CDODOSDOS = " & Val(X)
            If xCOP <> "" Then xWhere = xWhere & "   and CDODOSCOP = '" & xCOP & "'"
        
        
            xSql = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
                 & xWhere & " order by CDODOSCOP , CDODOSDOS"
            Set rsSab = cnsab.Execute(xSql)
            
            fgSelect_Display_ZCDODOS0
            
            
            If fgSelect.Rows = 2 Then
                fgSelect.Row = fgSelect.Rows - 1
                 fgSelect.Col = 0: xYDOSSLD0.DOSSLDDEV = Trim(fgSelect.Text)
                 fgSelect.Col = 1: xYDOSSLD0.DOSSLDOPE = Trim(fgSelect.Text)
                 fgSelect.Col = 2: xYDOSSLD0.DOSSLDNUM = Val(Trim(fgSelect.Text))
                fgSelect.Col = 3: xYDOSSLD0.DOSSLDCLI = Trim(fgSelect.Text)
                'fgYSWISAB0_Display
                SSTab2.Tab = 2: SSTab2.Visible = True
            End If
        Case "RDE", "RDI"
            xWhere = " where ENCCARDOS = " & Val(X)
            X = Trim(cboSelect_DOSSLDOPE)
            If X <> "" Then xWhere = xWhere & "   and ENCCARCOP = '" & xCOP & "'"
        
        
            xSql = "select * from " & paramIBM_Library_SAB & ".ZENCCAR0 " _
                 & xWhere & " order by ENCCARCOP , ENCCARDOS"
            Set rsSab = cnsab.Execute(xSql)
            
            fgSelect_Display_ZENCCAR0
            
            
            If fgSelect.Rows = 2 Then
                fgSelect.Row = fgSelect.Rows - 1
                 fgSelect.Col = 0: xYDOSSLD0.DOSSLDDEV = Trim(fgSelect.Text)
                 fgSelect.Col = 1: xYDOSSLD0.DOSSLDOPE = Trim(fgSelect.Text)
                 fgSelect.Col = 2: xYDOSSLD0.DOSSLDNUM = Val(Trim(fgSelect.Text))
                fgSelect.Col = 3: xYDOSSLD0.DOSSLDCLI = Trim(fgSelect.Text)
                'fgYSWISAB0_Display
                SSTab2.Tab = 2: SSTab2.Visible = True
            End If
        End Select
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_3()
Dim V, X As String
Dim xSql As String

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_3"

xSql = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
     & " where CDODOSETA <> '03' "
Set rsSab = cnsab.Execute(xSql)

fgSelect_Display_3

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_Scan_Importation()
Dim V, X As String, K As Integer, K1 As Integer
Dim wCDODOSCOP As String, blnOk As Boolean, blnSpace As Boolean
Dim xFileName As String, wCDODOSDOS As Long, mDOS_Path As String, mDOS_seq As Long
Dim objFolder, objFiles
Dim fsoFile As File, xFileName_Suite As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_Scan_Importation"

V = sqlYBIATAB0_Read("CREDOC", "WINDOWS_TEMP", "_SCAN", X)
If Not IsNull(V) Then GoTo Error_MsgBox
X = Trim(X)
'X = "_Scan\" & xYDOSSLD0.DOSSLDOPE & "_" & Format(xYDOSSLD0.DOSSLDNUM, "000000") & "\"
If Dir(X) = "" Then
    V = "Répertoire inconnu : " & X
    GoTo Error_MsgBox
Else
    If Mid$(X, Len(X), 1) = "\" Then X = Mid$(X, 1, Len(X) - 1)
    Set objFolder = msFileSystem.GetFolder(X)
    Set objFiles = objFolder.Files
    For Each fsoFile In objFiles
        'If InStr(fsoFile.Type, "Document") > 0 Then
            blnOk = False
            xFileName = UCase$(fsoFile.Name) & "???"
            Select Case Mid$(xFileName, 1, 3)
                Case "CDE": wCDODOSCOP = "CDE": blnOk = True
                Case "CDI": wCDODOSCOP = "CDI": blnOk = True
                Case "RDE": wCDODOSCOP = "RDE": blnOk = True
                Case "RDI": wCDODOSCOP = "RDI": blnOk = True
                Case "ENG": wCDODOSCOP = "ENG": blnOk = True
                Case "GAR": wCDODOSCOP = "GAR": blnOk = True
            End Select
            
            If blnOk Then
                blnSpace = True
                wCDODOSDOS = 0
                If IsNumeric(Mid$(xFileName, 4, 1)) Then
                    K1 = 4
                Else
                    K1 = 5
                End If
                For K = K1 To Len(xFileName)
                    If IsNumeric(Mid$(xFileName, K, 1)) Then
                        blnSpace = False
                        wCDODOSDOS = wCDODOSDOS * 10 + Val(Mid$(xFileName, K, 1))
                    Else
                        If Not blnSpace Then
                            Exit For
                        Else
                            If Mid$(xFileName, K, 1) <> " " Then Exit For
                        End If
                    End If
                Next K
                K1 = Len(fsoFile.Name)
                If K < K1 Then
                    xFileName_Suite = Mid$(fsoFile.Name, K, K1 - K + 1)
                Else
                    xFileName_Suite = "." & fileName_Extension(fsoFile.Name)
                End If
                                
                Select Case wCDODOSCOP
                    Case "CDE", "CDI":
                        X = "select *  from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
                          & " where CDODOSETB = " & currentSAB_ETA & " and CDODOSAGE = " & currentSAB_AGE _
                          & " and CDODOSSER = '00' and CDODOSSSE = '00'" _
                          & " and CDODOSCOP = '" & wCDODOSCOP & "' and CDODOSDOS = " & wCDODOSDOS
                    Case "RDE", "RDI":
                        X = "select *  from " & paramIBM_Library_SAB & ".ZENCCAR0 " _
                          & " where ENCCARETA = " & currentSAB_ETA & " and ENCCARAGE = " & currentSAB_AGE _
                          & " and ENCCARSER = '00' and ENCCARSSE = '00'" _
                          & " and ENCCARCOP = '" & wCDODOSCOP & "' and ENCCARDOS = " & wCDODOSDOS
                    Case "ENG", "GAR":
                        X = "select *  from " & paramIBM_Library_SAB & ".ZCAUDOS0 " _
                          & " where CAUDOSETB = " & currentSAB_ETA & " and CAUDOSAGE = " & currentSAB_AGE _
                          & " and CAUDOSDOS = " & wCDODOSDOS
                End Select
                
                Set rsSabX = cnsab.Execute(X)
                If rsSabX.EOF Then
                    If Not blnAuto Then Call MsgBox("Dossier inconnu : " & wCDODOSCOP & " " & wCDODOSDOS, vbCritical, "cmdSelect_SQL_Scan_Importation")
                Else
            
                    X = wCDODOSCOP & "_" & Format(wCDODOSDOS, "000000")
                    mDOS_Path = paramCDO_Dossier_Path & "_SCAN\" & X
                    If Not msFileSystem.FolderExists(mDOS_Path) Then MkDir mDOS_Path
                    mDOS_seq = mDOS_seq + 1
                    X = mDOS_Path & "\" & X & "_" & DSYS_Time & mDOS_seq & xFileName_Suite ' fileName_Extension(fsoFile.Name)
                    fsoFile.Move X
                End If
                
            End If
        'End If
    Next
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_Scan_Liste()
Dim wJMA10 As String
Dim objFolders As Folder, objFolder As Folder, fsoFile As File
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_Scan_Liste"
On Error GoTo Error_Handler

Call DTPicker_Control(txtSelect_Options_Scan_Liste_AMJ, X)
wJMA10 = dateImp10_S(X)
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Fichier                                           |<le                                     "
fgSelect.Row = 0


X = paramCDO_Dossier_Path & "_SCAN\"
    Set objFolders = msFileSystem.GetFolder(X)
    For Each objFolder In objFolders.SubFolders
        If Mid$(objFolder.DateLastModified, 1, 10) = wJMA10 Then
            Call lstErr_AddItem(lstErr, cmdContext, objFolder.Name): DoEvents

            For Each fsoFile In objFolder.Files
                If Mid$(fsoFile.DateLastModified, 1, 10) = wJMA10 Then
                    If Trim(fsoFile.Name) <> "Thumbs.db" Then
                        fgSelect.Rows = fgSelect.Rows + 1
                        fgSelect.Row = fgSelect.Rows - 1
                        fgSelect.Col = 0: fgSelect.Text = fsoFile.Name
                        fgSelect.Col = 1: fgSelect.Text = fsoFile.DateLastModified
                        fgSelect.Col = 2: fgSelect.Text = fsoFile.path
                    End If
                    'Debug.Print "- "; fsoFile.Name; fsoFile.DateLastModified; fsoFile
                End If
            Next
        End If
    Next
fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_3uti()
Dim V, X As String, K As Integer
Dim xSql As String, xAnd As String

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_3uti"

X = Trim(cboSelect_Options_3uti_UTI.Text)
If X <> "" Then
    K = InStr(X, "|")
    If K > 0 Then
        xAnd = " and CDOUTIVA1 = " & Mid$(X, K + 1, Len(X) - K + 1)
    End If
End If

X = Trim(cboSelect_Options_3uti_CDODOSNOT.Text)
If X <> "" Then
    K = InStr(X, "-")
    If K > 0 Then
        xAnd = xAnd & " and CDODOSNOT = " & Format(Val(Mid$(X, 1, K - 1)), "0000000")
    End If
End If

Call DTPicker_Control(txtSelect_Options_3uti_AmjMin, wAmjMin)
Call DTPicker_Control(txtSelect_Options_3uti_AmjMax, wAmjMax)

xAnd = xAnd & " and CDOUTIPRE >= " & wAmjMin - 19000000 & " and CDOUTIPRE <= " & wAmjMax - 19000000

xSql = "select * from " & paramIBM_Library_SAB & ".ZCDOUTI0 , " _
                        & paramIBM_Library_SAB & ".ZCDODOS0 , " _
                        & paramIBM_Library_SAB & ".ZMNURUT0 , " _
                        & paramIBM_Library_SAB & ".ZCLIENA0 " _
     & " where CDOUTIEVE = '03' and CDOUTIATT = '01' and CDOUTIETA = '02' and CDOUTICOP = 'CDE'" _
     & xAnd _
     & " and CDODOSEVE not in ('  ','80','90')" _
     & " and CDOUTICOP = CDODOSCOP" _
     & " and CDOUTIDOS = CDODOSDOS" _
     & " and CDOUTIVA1 = MNURUTCUT" _
     & " and CDODOSNOT = CLIENACLI" _
     & " order by MNURUTUTI , CDOUTIPRE , CDODOSDOS , CDOUTIUTI"
     
Set rsSab = cnsab.Execute(xSql)

fgSelect_Display_3uti

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_3RDO()
Dim V, X As String, K As Integer
Dim xSql As String, xAnd As String

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_3RDO"

X = Trim(cboSelect_Options_3uti_UTI.Text)
If X <> "" Then
    K = InStr(X, "|")
    If K > 0 Then
        xAnd = " and ENCCARUL1 = " & Mid$(X, K + 1, Len(X) - K + 1)
    End If
End If

X = Trim(cboSelect_Options_3uti_CDODOSNOT.Text)
If X <> "" Then
    K = InStr(X, "-")
    If K > 0 Then
        xAnd = xAnd & " and ENCCARORD = " & Format(Val(Mid$(X, 1, K - 1)), "0000000")
    End If
End If

Call DTPicker_Control(txtSelect_Options_3uti_AmjMin, wAmjMin)
Call DTPicker_Control(txtSelect_Options_3uti_AmjMax, wAmjMax)

'xAnd = xAnd & " and ENCCAREC1 >= " & wAmjMin - 19000000 & " and ENCCAREC1 <= " & wAmjMax - 19000000

xSql = "select * from " & paramIBM_Library_SAB & ".ZENCCAR0 , " _
                        & paramIBM_Library_SAB & ".ZMNURUT0  " _
     & " where ENCCARCET < '90'" _
     & xAnd _
     & " and ENCCARUL1 = MNURUTCUT" _
     & " order by MNURUTUTI , ENCCAREC1 , ENCCARDOS "
     
Set rsSab = cnsab.Execute(xSql)

fgSelect_Display_3RDO

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdSelect_SQL_3uti_Export()
On Error GoTo Error_Handler
Dim Nb As Long, wId As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSql As String
Dim X As String, K As Long, K2 As Long, kMax As Long, K_Nb As Long, K_Mt As Long
Dim xWhere As String, X2 As String
Dim wForecolor As Long, wBackColor As Long
'______________________________________________

wFile = "C:\Temp\SAB_Dossier_3uti " & DSys & " " & time_Hms & ".xlsx"
'______________________________________________

X = InputBox("par défaut : " & wFile _
    & vbCrLf & vbCrLf & "     =========================" _
    & vbCrLf & "     =========================", "SAB_Dossier : nom du fichier d'exportation", wFile)
If Trim(X) = "" Then Exit Sub

wFilex = Trim(X)
'______________________________________________


If Dir(wFile) <> "" Then Kill wFile

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "SAB_Dossier"
    .Subject = "SAB_Dossier"
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "SAB_Dossier"
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
    .Font.Size = 9
    .Font.Name = "Calibri"
End With

wsExcel.Columns(1).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
wsExcel.Columns(2).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
wsExcel.Columns(4).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Columns(7).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 100

wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
Nb = 1
Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : " & Nb & " enregistrements"): DoEvents

'____________________________________________________________________________________________
    Nb = 0
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14SAB_Dossier : utilisations en attente au " & dateImp10_S(DSys)


    For K = 0 To fgSelect.Rows - 1
        fgSelect.Row = K
        
        Nb = Nb + 1
        For K2 = 0 To 10
        
            fgSelect.Col = K2: X = Trim(fgSelect.Text)
            If K = 0 Then
                wsExcel.Columns(K2 + 1).ColumnWidth = fgSelect.CellWidth / 100
                wsExcel.Cells(Nb, K2 + 1).Font.Color = vbWhite
                wsExcel.Cells(Nb, K2 + 1).Interior.Color = mColor_GB
            Else
                wsExcel.Cells(Nb, K2 + 1).Font.Color = colorHex_RGB(fgSelect.CellForeColor)

            End If
            wsExcel.Cells(Nb, K2 + 1) = X
        Next K2
    Next K
'____________________________________________________________________________________________

Call lstErr_ChangeLastItem(lstErr, cmdContext, "Exportation en cours : " & Nb & " enregistrements"): DoEvents
Set rsSab = Nothing


wbExcel.SaveAs wFile

wbExcel.Close

'____________________________________________________________________________________
appExcel.Quit

Set rsSab = Nothing

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing

Call lstErr_AddItem(lstErr, cmdContext, "Exportation terminée"): DoEvents

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub cmdSelect_SQL_YDOSXOD0_Export()
On Error GoTo Error_Handler
Dim Nb As Long, wId As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSql As String
Dim X As String, K As Long, K2 As Long, kMax As Long, K_Nb As Long, K_Mt As Long
Dim xWhere As String, X2 As String
Dim wForecolor As Long, wBackColor As Long
'______________________________________________

wFile = "C:\Temp\SAB_Dossier_zOD " & DSys & " " & time_Hms & ".xlsx"
'______________________________________________

X = InputBox("par défaut : " & wFile _
    & vbCrLf & vbCrLf & "     =========================" _
    & vbCrLf & "     =========================", "SAB_Dossier : nom du fichier d'exportation", wFile)
If Trim(X) = "" Then Exit Sub

wFilex = Trim(X)
'______________________________________________


If Dir(wFile) <> "" Then Kill wFile

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "SAB_Dossier"
    .Subject = "SAB_Dossier"
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "SAB_Dossier"
'__________________________________________________________________________________

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .HorizontalAlignment = Excel.xlHAlignLeft
    .WrapText = True
    .Font.Size = 9
    .Font.Name = "Calibri"
End With

wsExcel.Columns(1).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
wsExcel.Columns(71).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
wsExcel.Columns(10).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
wsExcel.Columns(6).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Columns(11).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 100

wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
Nb = 1
Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : " & Nb & " enregistrements"): DoEvents

'____________________________________________________________________________________________
    Nb = 0
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14SAB_Dossier : liste des OD au " & dateImp10_S(DSys)


    For K = 0 To fgLOG.Rows - 1
        fgLOG.Row = K
        
        Nb = Nb + 1
        For K2 = 0 To 10
        
            fgLOG.Col = K2: X = Trim(fgLOG.Text)
            If K = 0 Then
                wsExcel.Columns(K2 + 1).ColumnWidth = fgLOG.CellWidth / 100
                wsExcel.Cells(Nb, K2 + 1).Font.Color = vbWhite
                wsExcel.Cells(Nb, K2 + 1).Interior.Color = mColor_GB
            'Else
                'wsExcel.Cells(Nb, K2 + 1).Font.Color = mColor_GB

            End If
            wsExcel.Cells(Nb, K2 + 1) = X
        Next K2
    Next K
'____________________________________________________________________________________________

Call lstErr_ChangeLastItem(lstErr, cmdContext, "Exportation en cours : " & Nb & " enregistrements"): DoEvents
Set rsSab = Nothing


wbExcel.SaveAs wFile

wbExcel.Close

'____________________________________________________________________________________
appExcel.Quit

Set rsSab = Nothing

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing

Call lstErr_AddItem(lstErr, cmdContext, "Exportation terminée"): DoEvents

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub


Private Sub cmdSelect_SQL_2()
Dim V, X As String
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean
Dim xDOSSLDM As String, xDOSSLDG As String, xDOSSLDK As String
Dim xField1 As String, xK As String, xField2 As String

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2"
blnOk = False

arrYBIACPT0_Nb = 0

xWhere = ""
X = Trim(txtSelect_DOSSLDNUM)
If X <> "" Then xWhere = "   and DOSSLDCLI = '" & Format(Val(X), "0000000") & "'"


X = Trim(cboSelect_DOSSLDPCI)
If X <> "" Then xWhere = xWhere & "   and DOSSLDPCI = '" & X & "'"
If chkSelect_DOSSLDMG = "1" Then
    If X = "" Then
        Call MsgBox("Différence de solde, préciser le PCI (91120,91122,98050,90312)", vbCritical, "SAB_Dossier")
        Exit Sub
    Else
        xWhere = xWhere & " and (DOSSLDMSD <> DOSSLDGSD) "
    End If
End If




X = Trim(cboSelect_DOSSLDDEV)
If X <> "" Then xWhere = xWhere & "   and DOSSLDDEV = '" & X & "'"
If xWhere <> "" Then Mid$(xWhere, 1, 6) = " where"

xSql = "select *  from " & paramIBM_Library_SABSPE & ".YDOSSLD1 " _
     & xWhere & " order by DOSSLDDEV , DOSSLDPCI , DOSSLDCLI"
Set rsSab = cnsab.Execute(xSql)


fgSelect_Display_2

If fgSelect.Rows = 2 Then
'    fgSelect.Row = fgSelect.Rows - 1
'     fgSelect.Col = 0: xYDOSSLD0.DOSSLDDEV = Trim(fgSelect.Text)
'     fgSelect.Col = 1: xYDOSSLD0.DOSSLDOPE = Trim(fgSelect.Text)
'     fgSelect.Col = 2: xYDOSSLD0.DOSSLDNUM = Val(Trim(fgSelect.Text))
       
'    fgDetail_Display
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_5()
Dim V, X As String
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_5"
blnOk = False

xSql = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
     & "where CDODOSEVE = '" & Mid$(cboSelect_DOSSLDSTA, 1, 2) & "' and CDODOSETA = '" & Mid$(cboSelect_DOSSLDSTA, 6, 2) & "'"
Set rsSab = cnsab.Execute(xSql)

fgSelect_Display_3

'xWhere = "where DOSSLDSTA = '" & Mid$(cboSelect_DOSSLDSTA, 1, 2) & "' and DOSSLDSVC = '" & Mid$(cboSelect_DOSSLDSTA, 6, 2) & "'"
'xSql = "select distinct DOSSLDDEV , DOSSLDOPE , DOSSLDNUM , DOSSLDCLI from " & paramIBM_Library_SABSPE & ".YDOSSLD0 " _
'     & xWhere & " order by DOSSLDDEV , DOSSLDOPE , DOSSLDNUM"
'Set rsSab = cnsab.Execute(xSql)
'fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_5_ECNFPT()
Dim V, X As String
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_5_ECNFPT"
blnOk = False

xSql = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0" _
     & " where CDODOSCOP = 'CDE' and CDODOSEVE <> '90' and CDODOSEVE <> '80'" _
         & " and CDODOSDOS  in ( select distinct(CDOCOMDOS) from " & paramIBM_Library_SAB & ".ZCDOCOM0" _
         & " where CDOCOMCOP = 'CDE' and CDOCOMCOM = 'ECNFPT') order by CDODOSDOS, CDODOSNOT"

Set rsSab = cnsab.Execute(xSql)

fgSelect_Display_ZCDODOS0


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_5_AUT()
Dim V, X As String, K As Long, K0 As Long, wAUTENAAUT As String, Nb As Long, wErr As String
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_5_AUT"
Call lstErr_AddItem(lstErr, cmdContext, "> " & " ........"): DoEvents

blnOk = False

fgSelect.Width = 16000
fgSelect.Height = 8800
'fgSelect.BackColor = RGB(240, 240, 240)
K0 = 1
Call cmdSelect_SQL_5_AUT_arr


xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0" _
     & " where DOSSLDOPE = 'CDE' and DOSSLDPCI > '9' and DOSSLDPCI < '999'" _
     & " and DOSSLDSTA <> '  ' order by DOSSLDNUM, DOSSLDDEV, DOSSLDCLI"

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    xYDOSSLD0.DOSSLDNUM = rsSab("DOSSLDNUM")
    xYDOSSLD0.DOSSLDSTA = rsSab("DOSSLDSTA")
    
    If xYDOSSLD0.DOSSLDSTA = "90" And xYDOSSLD0.DOSSLDNUM < 100000 Then
        'REPRISE
    Else
        xYDOSSLD0.DOSSLDPCI = rsSab("DOSSLDPCI")
    '    Select Case Mid$(xYDOSSLD0.DOSSLDPCI, 1, 5)
    '        Case "91130", "98750", "85052": wAUTENAAUT = "PDI"
    '        Case "91120": wAUTENAAUT = "CEC"
    '        Case "98050": wAUTENAAUT = "CEN"
    '    End Select
        xYDOSSLD0.DOSSLDMSD = rsSab("DOSSLDMSD")
        xYDOSSLD0.DOSSLDCLI = rsSab("DOSSLDCLI")
        xYDOSSLD0.DOSSLDDEV = rsSab("DOSSLDDEV")
        blnOk = False
        wErr = "?"
        Nb = arrZAUTENA0_Nb
        For K = K0 To Nb
            If xYDOSSLD0.DOSSLDNUM > arrZAUTENA0(K).AUTENADOS Then
                K0 = K
            Else
                If xYDOSSLD0.DOSSLDNUM < arrZAUTENA0(K).AUTENADOS Then
                    Exit For
                Else
                    arrZAUTENA0(K).DOSSLDSTA = xYDOSSLD0.DOSSLDSTA
                    If arrZAUTENA0(K).AUTENADEV = xYDOSSLD0.DOSSLDDEV _
                   And arrZAUTENA0(K).AUTENACLI = xYDOSSLD0.DOSSLDCLI Then
                        arrZAUTENA0(K).DOSSLDMSD = arrZAUTENA0(K).DOSSLDMSD - xYDOSSLD0.DOSSLDMSD
                        arrZAUTENA0(K).DOSSLDPCI = arrZAUTENA0(K).DOSSLDPCI & "_" & xYDOSSLD0.DOSSLDPCI
                        blnOk = True
                        Exit For
                    End If
                End If
            End If
            
        Next K
        
        If Not blnOk Then
            If xYDOSSLD0.DOSSLDMSD <> 0 Then
            
                xSql = "select CDODOSBER , CDODOSBEN from " & paramIBM_Library_SAB & ".ZCDODOS0" _
                     & " where CDODOSCOP = 'CDE' and CDODOSDOS = " & xYDOSSLD0.DOSSLDNUM
                
                Set rsSabX = cnsab.Execute(xSql)
                If Not rsSabX.EOF Then
                    If rsSabX("CDODOSBER") = " " And Val(rsSabX("CDODOSBEN")) = xYDOSSLD0.DOSSLDCLI Then blnOk = True
                End If
                
                
            
                If Not blnOk Then
                
                    arrZAUTENA0_Nb = arrZAUTENA0_Nb + 1
                    arrZAUTENA0(arrZAUTENA0_Nb) = arrZAUTENA0(0)
                    arrZAUTENA0(arrZAUTENA0_Nb).AUTENAAUT = wErr
                    arrZAUTENA0(arrZAUTENA0_Nb).AUTENADOS = xYDOSSLD0.DOSSLDNUM
                    arrZAUTENA0(arrZAUTENA0_Nb).DOSSLDSTA = xYDOSSLD0.DOSSLDSTA
                    arrZAUTENA0(arrZAUTENA0_Nb).DOSSLDMSD = xYDOSSLD0.DOSSLDMSD
                    arrZAUTENA0(arrZAUTENA0_Nb).DOSSLDPCI = xYDOSSLD0.DOSSLDPCI
                    arrZAUTENA0(arrZAUTENA0_Nb).AUTENAOPE = rsSab("DOSSLDOPE")
                    arrZAUTENA0(arrZAUTENA0_Nb).AUTENADEV = xYDOSSLD0.DOSSLDDEV
                    arrZAUTENA0(arrZAUTENA0_Nb).AUTENACLI = xYDOSSLD0.DOSSLDCLI
                End If
            End If
        End If
    End If
    
    rsSab.MoveNext
Loop


fgSelect_Display_5_AUT
Call lstErr_AddItem(lstErr, cmdContext, "< "): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_5_ECNFPT_Com()
Dim V, X As String
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_5_ECNFPT_Com"
blnOk = False

'cmdSelect_SQL_5_ECNFPT_arr
xYDOSSLD0.DOSSLDOPE = "CDE"
Call fgCOM_Display
SSTab2.Tab = 5
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub cmdSelect_SQL_5_ECNFPT_arr()
Dim xSql As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_5_ECNFPT_arr"
xSql = "select count(*) from " & paramIBM_Library_SAB & ".ZCDODOS0" _
     & " where CDODOSCOP = 'CDE' and CDODOSEVE <> '90' and CDODOSEVE <> '80'" _
         & " and CDODOSDOS  in ( select distinct(CDOCOMDOS) from " & paramIBM_Library_SAB & ".ZCDOCOM0" _
         & " where CDOCOMCOP = 'CDE' and CDOCOMCOM = 'ECNFPT')" _

Set rsSab = cnsab.Execute(xSql)

ReDim arrECNFPT_DOS(rsSab(0) + 1)
arrECNFPT_DOS_Nb = 0


xSql = "select CDODOSDOS from " & paramIBM_Library_SAB & ".ZCDODOS0" _
     & " where CDODOSCOP = 'CDE' and CDODOSEVE <> '90'" _
         & " and CDODOSDOS  in ( select distinct(CDOCOMDOS) from " & paramIBM_Library_SAB & ".ZCDOCOM0" _
         & " where CDOCOMCOP = 'CDE' and CDOCOMCOM = 'ECNFPT') and CDODOSEVE <> '80'" _

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    arrECNFPT_DOS_Nb = arrECNFPT_DOS_Nb + 1
    arrECNFPT_DOS(arrECNFPT_DOS_Nb) = rsSab("CDODOSDOS")
    rsSab.MoveNext
Loop



Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_5_AUT_arr()
Dim xSql As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_5_AUT_arr"
xSql = "select count(*) from " & paramIBM_Library_SAB & ".ZAUTENA0" _
     & " where AUTENAOPE = 'CDE' "
Set rsSab = cnsab.Execute(xSql)

ReDim arrZAUTENA0(5000)  '(rsSab(0) * 2 + 1)
arrZAUTENA0_Nb = 0
Call rsZAUTENA0_Init(arrZAUTENA0(0))


xSql = "select * from " & paramIBM_Library_SAB & ".ZAUTENA0" _
     & " where AUTENAOPE = 'CDE'   and AUTENADOS > 80000  order by AUTENADOS, AUTENADEV, AUTENACLI, AUTENAAUT"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    
    Call rsZAUTENA0_GetBuffer(rsSab, xZAUTENA0)
    If xZAUTENA0.AUTENADOS = arrZAUTENA0(arrZAUTENA0_Nb).AUTENADOS _
   And xZAUTENA0.AUTENADEV = arrZAUTENA0(arrZAUTENA0_Nb).AUTENADEV _
   And xZAUTENA0.AUTENACLI = arrZAUTENA0(arrZAUTENA0_Nb).AUTENACLI Then
        arrZAUTENA0(arrZAUTENA0_Nb).AUTENAENC = arrZAUTENA0(arrZAUTENA0_Nb).AUTENAENC + xZAUTENA0.AUTENAENC
        arrZAUTENA0(arrZAUTENA0_Nb).AUTENAAUT = arrZAUTENA0(arrZAUTENA0_Nb).AUTENAAUT & "_" & xZAUTENA0.AUTENAAUT
    Else
        arrZAUTENA0_Nb = arrZAUTENA0_Nb + 1
        arrZAUTENA0(arrZAUTENA0_Nb) = xZAUTENA0
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


Private Sub cmdSelect_SQL_5réfext(lFct As String)
Dim V, X As String
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_5réf"
blnOk = False


xSql = "select CDODOSEXT , count(*)   from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
     & " where CDODOSCOP = 'CDE' and CDODOSEVE <> '90'" _
     & " group by CDODOSEXT having count(*) > 1"
Set rsSab = cnsab.Execute(xSql)

Call fgSelect_Display_5réfext(lFct)


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Public Sub cmdSelect_SQL_Surveillance()
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_Surveillance"
blnZCAUDOS0_S01 = False

cmdSelect_SQL_Surveillance_CDO
If blnAuto Then Call cmdSendMail_SAB_Dossier("", "BIA-CDO-Surveillance")
'
cmdSelect_SQL_Surveillance_RDO
If blnAuto Then Call cmdSendMail_SAB_Dossier("", "BIA-RDO-Surveillance")


cmdSelect_SQL_Surveillance_CAU
If blnAuto Then Call cmdSendMail_SAB_Dossier("", "BIA-CAU-Surveillance")

blnZCAUDOS0_S01 = True
cmdSelect_SQL_Surveillance_CAU
If blnAuto Then Call cmdSendMail_SAB_Dossier("", "BIA-CAU-Surveillance-GDMP")

blnZCAUDOS0_S01 = False
'=======================
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_Surveillance_CDO()
Dim V, X As String, XSVC1 As String, XSVC2 As String
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean
Dim wCli As String, wOPE As String, wDOS As Long, wErr As String, wLIB As String, wNAT As String
Dim wMTD As Currency, wDev As String, wAmj As String
Dim mCDODOSVAL_Max As Long
Dim nbD As Long, nbT As Long, wDT As String

Dim wMDB As Currency, wMCR As Currency
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_Surveillance_CDO"
blnOk = False
lstW.Clear
mCDODOSVAL_Max = dateElp("MoisAdd", -3, DSys) - 19000000
'_______________________________________________________________________________________
'wCli = "Surveillance en date du " & dateImp10_S(YBIATAB0_DATE_CPT_J)
'wLIB = "Edité le " & dateImp10_S(DSys) & "  " & Time

'lstW.AddItem "00X|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""

'_______________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 " _
     & " where (DOSSLDMSD <> DOSSLDGSD)   and substring(DOSSLDPCI , 1 , 5) in ('91120','91122','98050','91130','91131','90312')" _
     & " and DOSSLDSTA not in ('  ','80','90')   and DOSSLDSVC <> '01' "
Set rsSab = cnsab.Execute(xSql)
nbT = 0: nbD = 0

wLIB = "DOSSIER  : écart compta / gestion"
Do While Not rsSab.EOF
    wOPE = rsSab("DOSSLDOPE")
    wDOS = rsSab("DOSSLDNUM")
    wCli = rsSab("DOSSLDCLI")
    wMTD = 0 '
    wDev = rsSab("DOSSLDDEV")
    wErr = "# " & rsSab("DOSSLDPCI")
    wAmj = ""
        
    xSql = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
         & " where CDODOSCOP = '" & wOPE & "'" _
         & " and CDODOSDOS = " & wDOS
    Set rsSabX = cnsab.Execute(xSql)
    If Not rsSabX.EOF Then
        wNAT = rsSabX("CDODOSCON")
        wMTD = rsSabX("CDODOSMOT")
        wDev = rsSabX("CDODOSDEV")
        wAmj = dateImp10(rsSabX("CDODOSVAL") + 19000000)
    End If
    wNAT = ""
    wMTD = rsSab("DOSSLDMSD") - rsSab("DOSSLDGSD")
    If wMTD <> 0 Then
        wNAT = "- C#G"
        nbD = nbD + 1
           lstW.AddItem "02D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
        
    End If
    
    rsSab.MoveNext
Loop

wLIB = "Dossiers non clos en écart comptabilité / gestion : " & nbD
wCli = "PCI :91120,91122,98050,91130,91131,90312"
If nbD = 0 Then
    wLIB = "Dossiers non clos en écart comptabilité / gestion : NEANT"
    lstW.AddItem "02S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "02T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If
'_____________________________________________________________________________________________


xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSMVT0 " _
     & " where DOSMVTNUM = 0 and substring(DOSMVTPCI , 1 , 5) in ('91120','91122', '98050','91130','91131','90312') "
Set rsSab = cnsab.Execute(xSql)

wLIB = "mvt comptable non affecté à un dossier "
nbT = 0: nbD = 0

Do While Not rsSab.EOF
    wOPE = rsSab("DOSMVTOPE")
    wDOS = rsSab("DOSMVTNUM")
    wCli = rsSab("DOSMVTCLI")
    wNAT = rsSab("DOSMVTEVE")
    wMTD = rsSab("DOSMVTMTD")
    wDev = rsSab("DOSMVTDEV")
    wErr = "? " & rsSab("DOSMVTPCI")
    wAmj = dateImp10(rsSab("DOSMVTDTR"))
    wLIB = "mvt comptable " & rsSab("DOSMVTPIE") & "-" & rsSab("DOSMVTECR") & " non affecté à un dossier "
    nbD = nbD + 1
    lstW.AddItem "03D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
    rsSab.MoveNext
Loop
wLIB = "Mouvements comptables orphelins, non affectés à un dossier : " & nbD
wCli = "PCI :91120,91122,98050,91130,91131,90312"
If nbD = 0 Then
    wLIB = "Mouvements comptables orphelins, non affectés à un dossier : NEANT"
    lstW.AddItem "03S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "03T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If
'_____________________________________________________________________________________________


xSql = "select *  from " & paramIBM_Library_SABSPE & ".YDOSSLD1 " _
     & " where (DOSSLDMSD <> DOSSLDGSD)   and substring(DOSSLDPCI , 1 , 5) in ('91120','91122','98050','91130','91131','90312')"
Set rsSab = cnsab.Execute(xSql)
nbT = 0: nbD = 0

wLIB = "COMPTE : écart "
Do While Not rsSab.EOF
    wOPE = ""
    wDOS = 0
    wCli = rsSab("DOSSLDCLI")
    wMTD = 0 '
    wDev = rsSab("DOSSLDDEV")
    wErr = "* " & rsSab("DOSSLDPCI")
    wAmj = ""
    
    wNAT = ""
    wMTD = rsSab("DOSSLDMSD") - rsSab("DOSSLDGSD")
    If wMTD <> 0 Then
        wNAT = "? C#G"
        wLIB = "COMPTE : écart - " & rsSab("DOSSLDNBV") & " dossiers"
        nbD = nbD + 1
        lstW.AddItem "04D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
    End If
    rsSab.MoveNext
Loop
wLIB = "total des mouvements comptables (PCI-client) <> total des soldes de gestion : " & nbD
wCli = "PCI :91120,91122,98050,91130,91131,90312"
If nbD = 0 Then
    wLIB = "total des mouvements comptables (PCI-client) <> total des soldes de gestion : NEANT"
    lstW.AddItem "04S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "04T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If
'_____________________________________________________________________________________________

xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 " _
     & " where DOSSLDnum > 70000 and DOSSLDMSD <> 0 and substring(DOSSLDPCI , 1 , 5) in ('91120','91122','98050','91130','91131','90312')" _
     & " and DOSSLDSTA  in ('  ','80','90')  and DOSSLDSVC = '03'"
Set rsSab = cnsab.Execute(xSql)
nbT = 0: nbD = 0

wLIB = "dossier annulé en gestion ayant un solde comptable"
Do While Not rsSab.EOF
    wOPE = rsSab("DOSSLDOPE")
    wDOS = rsSab("DOSSLDNUM")
    wCli = rsSab("DOSSLDCLI")
    wMTD = rsSab("DOSSLDMSD") '
    wDev = rsSab("DOSSLDDEV")
    wErr = "! " & rsSab("DOSSLDPCI")
    wAmj = ""
        
    xSql = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
         & " where CDODOSCOP = '" & wOPE & "'" _
         & " and CDODOSDOS = " & wDOS
    Set rsSabX = cnsab.Execute(xSql)
    If Not rsSabX.EOF Then
        wNAT = rsSabX("CDODOSCON")
        wMTD = rsSabX("CDODOSMOT")
        wDev = rsSabX("CDODOSDEV")
        wAmj = dateImp10(rsSabX("CDODOSVAL") + 19000000)
    End If
    wNAT = ""
    wMTD = rsSab("DOSSLDMSD")
    If wMTD <> 0 Then
        wNAT = "! 90"
        nbD = nbD + 1
        lstW.AddItem "01D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
    End If
    
    rsSab.MoveNext
Loop
wLIB = "Dossiers clos en gestion mais présentant un solde comptable : " & nbD
wCli = "PCI :91120,91122,98050,91130,91131,90312"
If nbD = 0 Then
    wLIB = "Dossiers clos en gestion mais présentant un solde comptable : NEANT"
    lstW.AddItem "01S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "01T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If

'_______________________________________________________________________________________

nbT = 0: nbD = 0

xSql = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
     & " where CDODOSETA <> '03' "
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    wOPE = rsSab("CDODOSCOP")
    If wOPE = "CDE" Then
        wCli = rsSab("CDODOSNOT")
    Else
        wCli = rsSab("CDODOSDON")
    End If
    wDOS = rsSab("CDODOSDOS")
    wNAT = rsSab("CDODOSCON")
    wMTD = rsSab("CDODOSMOT")
    wDev = rsSab("CDODOSDEV")
    wErr = "D " & rsSab("CDODOSEVE")
    wAmj = dateImp10(rsSab("CDODOSVAL") + 19000000)
    Select Case rsSab("CDODOSETA")
        Case "01": XSVC1 = "non validé": XSVC2 = "non validée"
        Case "02": XSVC1 = "non comptabilisé": XSVC2 = "non comptabilisée"
        Case Else: XSVC1 = "SVC ???": XSVC2 = "SVC ???"
    End Select
    Select Case rsSab("CDODOSEVE")
        Case "01": wLIB = "Ouverture " & XSVC2
        Case "02": wLIB = "Modification " & XSVC2
        Case "07": wLIB = "Réouverture " & XSVC2
        Case "80": wLIB = "Annulation de solde " & XSVC2: wMTD = 0
        Case "90": wLIB = "Clôture " & XSVC2: wMTD = 0
        Case Else: wLIB = "événement ??? " & XSVC1
    End Select
    nbD = nbD + 1
    lstW.AddItem "10D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
    rsSab.MoveNext
Loop
xSql = "select count(*) as tally from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
     & " where CDODOSEVE not in ('  ','80','90')"
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then nbT = rsSab(0)

wLIB = "Evénements en attente : " & nbD & " / " & nbT & " dossiers non clos"
wCli = "code état <> comptabilisé"
If nbD = 0 Then
    wLIB = "Evénements  en attente : NEANT / " & nbT & "  dossiers non clos"
    lstW.AddItem "10S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "10T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If

'_______________________________________________________________________________________

xSql = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
     & " where  CDODOSVAL <= " & mCDODOSVAL_Max & " and CDODOSEVE not in ('  ','80','90')  "
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    wOPE = rsSab("CDODOSCOP")
    If wOPE = "CDE" Then
        wCli = rsSab("CDODOSNOT")
    Else
        wCli = rsSab("CDODOSDON")
    End If
    wDOS = rsSab("CDODOSDOS")
    wNAT = rsSab("CDODOSCON")
    wMTD = rsSab("CDODOSMOT")
    wDev = rsSab("CDODOSDEV")
    wErr = "V " & rsSab("CDODOSEVE")
    wAmj = dateImp10(rsSab("CDODOSVAL") + 19000000)
    wLIB = "date validité : " & dateImp10(rsSab("CDODOSVAL") + 19000000)

'en cours de développement
    '    nbD = nbD + 1

    'lstW.AddItem "11D|" &  wERR & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDEV & "|" & wAmj
    rsSab.MoveNext
Loop
'wLIB = "Dossiers < date validité : " & nbD
'lstW.AddItem "11T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""

'_____________________________________________________________________________________________

'_____________________________________________________________________________________________


Call YDOSSLD0_Export_CDO

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_Surveillance_RDO()
Dim V, X As String, XSVC1 As String, XSVC2 As String
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean
Dim wCli As String, wOPE As String, wDOS As Long, wErr As String, wLIB As String, wNAT As String
Dim wMTD As Currency, wDev As String, wAmj As String
Dim nbD As Long, nbT As Long, wDT As String

Dim wMDB As Currency, wMCR As Currency
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_Surveillance"
blnOk = False
lstW.Clear
'_______________________________________________________________________________________
'wCli = "Surveillance en date du " & dateImp10_S(YBIATAB0_DATE_CPT_J)
'wLIB = "Edité le " & dateImp10_S(DSys) & "  " & Time

'lstW.AddItem "00X|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""

'_______________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 " _
     & " where (DOSSLDMSD <> DOSSLDGSD)   and substring(DOSSLDPCI , 1 , 5) in ('98520')" _
     & " and DOSSLDSTA not in ('  ','90')   and DOSSLDSVC <> '01' "
Set rsSab = cnsab.Execute(xSql)
nbT = 0: nbD = 0

wLIB = "DOSSIER  : écart compta / gestion"
Do While Not rsSab.EOF
    wOPE = rsSab("DOSSLDOPE")
    wDOS = rsSab("DOSSLDNUM")
    wCli = rsSab("DOSSLDCLI")
    wMTD = 0 '
    wDev = rsSab("DOSSLDDEV")
    wErr = "# " & rsSab("DOSSLDPCI")
    wAmj = ""
        
    xSql = "select * from " & paramIBM_Library_SAB & ".ZENCCAR0 " _
         & " where ENCCARCOP = '" & wOPE & "'" _
         & " and ENCCARDOS = " & wDOS
    Set rsSabX = cnsab.Execute(xSql)
    If Not rsSabX.EOF Then
        wNAT = rsSabX("ENCCARNAT")
        wMTD = rsSabX("ENCCARMON")
        wDev = rsSabX("ENCCARDEV")
        wAmj = dateImp10(rsSabX("ENCCARDAR") + 19000000)
    End If
    wNAT = ""
    wMTD = rsSab("DOSSLDMSD") - rsSab("DOSSLDGSD")
    If wMTD <> 0 Then
        wNAT = "- C#G"
        nbD = nbD + 1
           lstW.AddItem "02D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
        
    End If
    
    rsSab.MoveNext
Loop

wLIB = "Dossiers non clos en écart comptabilité / gestion : " & nbD
wCli = "PCI :98520"
If nbD = 0 Then
    wLIB = "Dossiers non clos en écart comptabilité / gestion : NEANT"
    lstW.AddItem "02S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "02T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If
'_____________________________________________________________________________________________


xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSMVT0 " _
     & " where DOSMVTNUM = 0 and substring(DOSMVTPCI , 1 , 5) in ('98520') "
Set rsSab = cnsab.Execute(xSql)

wLIB = "mvt comptable non affecté à un dossier "
nbT = 0: nbD = 0

Do While Not rsSab.EOF
    wOPE = rsSab("DOSMVTOPE")
    wDOS = rsSab("DOSMVTNUM")
    wCli = rsSab("DOSMVTCLI")
    wNAT = rsSab("DOSMVTEVE")
    wMTD = rsSab("DOSMVTMTD")
    wDev = rsSab("DOSMVTDEV")
    wErr = "? " & rsSab("DOSMVTPCI")
    wAmj = dateImp10(rsSab("DOSMVTDTR"))
    wLIB = "mvt comptable " & rsSab("DOSMVTPIE") & "-" & rsSab("DOSMVTECR") & " non affecté à un dossier "
    nbD = nbD + 1
    lstW.AddItem "03D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
    rsSab.MoveNext
Loop
wLIB = "Mouvements comptables orphelins, non affectés à un dossier : " & nbD
wCli = "PCI :98520"
If nbD = 0 Then
    wLIB = "Mouvements comptables orphelins, non affectés à un dossier : NEANT"
    lstW.AddItem "03S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "03T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If
'_____________________________________________________________________________________________


xSql = "select *  from " & paramIBM_Library_SABSPE & ".YDOSSLD1 " _
     & " where (DOSSLDMSD <> DOSSLDGSD)   and substring(DOSSLDPCI , 1 , 5) in ('98520')"
Set rsSab = cnsab.Execute(xSql)
nbT = 0: nbD = 0

wLIB = "COMPTE : écart "
Do While Not rsSab.EOF
    wOPE = ""
    wDOS = 0
    wCli = rsSab("DOSSLDCLI")
    wMTD = 0 '
    wDev = rsSab("DOSSLDDEV")
    wErr = "* " & rsSab("DOSSLDPCI")
    wAmj = ""
    
    wNAT = ""
    wMTD = rsSab("DOSSLDMSD") - rsSab("DOSSLDGSD")
    If wMTD <> 0 Then
        wNAT = "? C#G"
        wLIB = "COMPTE : écart - " & rsSab("DOSSLDNBV") & " dossiers"
        nbD = nbD + 1
        lstW.AddItem "04D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
    End If
    rsSab.MoveNext
Loop
wLIB = "total des mouvements comptables (PCI-client) <> total des soldes de gestion : " & nbD
wCli = "PCI :98520"
If nbD = 0 Then
    wLIB = "total des mouvements comptables (PCI-client) <> total des soldes de gestion : NEANT"
    lstW.AddItem "04S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "04T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If
'_____________________________________________________________________________________________

xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 " _
     & " where DOSSLDMSD <> 0 and substring(DOSSLDPCI , 1 , 5) in ('98520')" _
     & " and DOSSLDSTA  in ('  ','80','90')  and DOSSLDSVC = '03'"
Set rsSab = cnsab.Execute(xSql)
nbT = 0: nbD = 0

wLIB = "dossier annulé en gestion ayant un solde comptable"
Do While Not rsSab.EOF
    wOPE = rsSab("DOSSLDOPE")
    wDOS = rsSab("DOSSLDNUM")
    wCli = rsSab("DOSSLDCLI")
    wMTD = rsSab("DOSSLDMSD") '
    wDev = rsSab("DOSSLDDEV")
    wErr = "! " & rsSab("DOSSLDPCI")
    wAmj = ""
    
    xSql = "select * from " & paramIBM_Library_SAB & ".ZENCCAR0 " _
         & " where ENCCARCOP = '" & wOPE & "'" _
         & " and ENCCARDOS = " & wDOS
    Set rsSabX = cnsab.Execute(xSql)
    If Not rsSabX.EOF Then
        wNAT = rsSabX("ENCCARNAT")
        wMTD = rsSabX("ENCCARMON")
        wDev = rsSabX("ENCCARDEV")
        wAmj = dateImp10(rsSabX("ENCCARDAR") + 19000000)
    End If
    wNAT = ""
    wMTD = rsSab("DOSSLDMSD")
    If wMTD <> 0 Then
        wNAT = "! 90"
        nbD = nbD + 1
        lstW.AddItem "01D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
    End If
    
    rsSab.MoveNext
Loop
wLIB = "Dossiers clos en gestion mais présentant un solde comptable : " & nbD
wCli = "PCI :98520"
If nbD = 0 Then
    wLIB = "Dossiers clos en gestion mais présentant un solde comptable : NEANT"
    lstW.AddItem "01S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "01T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If

'_______________________________________________________________________________________

nbT = 0: nbD = 0

xSql = "select * from " & paramIBM_Library_SAB & ".ZENCCAR0 " _
     & " where ENCCARCET not in ('03' , '13' , '83', '93') "
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    wOPE = rsSab("ENCCARCOP")
    wCli = ""
    wDOS = rsSab("ENCCARDOS")
    wNAT = rsSab("ENCCARNAT")
    wMTD = rsSab("ENCCARMON")
    wDev = rsSab("ENCCARDEV")
    wErr = "D " & rsSab("ENCCARCET")
    wAmj = dateImp10(rsSab("ENCCARDAR") + 19000000)
    Select Case rsSab("ENCCARCET")
        Case "01": wLIB = "non validé": XSVC2 = "non validée"
        Case "02": wLIB = "non comptabilisé": XSVC2 = "non comptabilisée"
        Case "91": wLIB = "clos non validé": XSVC2 = "clos non validée"
        Case "92": wLIB = "clos non comptabilisé": XSVC2 = "clos non comptabilisée"
       Case Else: wLIB = rsSab("ENCCARCET") & " ???": XSVC2 = rsSab("ENCCARCET") & " ???"
    End Select
    nbD = nbD + 1
    lstW.AddItem "10D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
    rsSab.MoveNext
Loop
xSql = "select count(*) as tally from " & paramIBM_Library_SAB & ".ZENCCAR0 " _
     & " where ENCCARCET <> '93'"
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then nbT = rsSab(0)

wLIB = "Evénements en attente : " & nbD & " / " & nbT & " dossiers non clos"
wCli = "code état <> comptabilisé"
If nbD = 0 Then
    wLIB = "Evénements  en attente : NEANT / " & nbT & "  dossiers non clos"
    lstW.AddItem "10S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "10T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If

'_______________________________________________________________________________________


'_____________________________________________________________________________________________

'_____________________________________________________________________________________________


Call YDOSSLD0_Export_RDO

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdSelect_SQL_Surveillance_CAU()
Dim V, X As String, XSVC1 As String, XSVC2 As String
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean
Dim wCli As String, wOPE As String, wDOS As Long, wErr As String, wLIB As String, wNAT As String
Dim wMTD As Currency, wDev As String, wAmj As String
Dim nbD As Long, nbT As Long, wDT As String

Dim wMDB As Currency, wMCR As Currency
Dim sqlPCI As String, listPCI As String

Dim whereYDOSSLD0 As String, whereZCAUDOS0 As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_Surveillance"
blnOk = False
lstW.Clear

If blnZCAUDOS0_S01 Then
    whereYDOSSLD0 = " and DOSSLDNUM < 500000 "
    whereZCAUDOS0 = " and CAUDOSDOS < 500000 "
Else
    whereYDOSSLD0 = " and DOSSLDNUM >= 500000 "
    whereZCAUDOS0 = " and CAUDOSDOS >= 500000 "
End If

sqlPCI = "'98720' "
X = "select distinct BIATABTXT from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & "where BIATABID = 'OPENAT_PCI' and substring (biatabk1 , 1 , 3 ) in ('ENG' , 'GAR')"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    sqlPCI = sqlPCI & ", '" & Mid$(rsSab("BIATABTXT"), 4, 5) & "'"
    'listPCI = listPCI & Mid$(rsSab("BIATABTXT"), 4, 5) & " "
    rsSab.MoveNext
Loop
'''''''''''sqlPCI = Replace(sqlPCI, "98710", "98720")
'Mid$(sqlPCI, 1, 1) = " "
listPCI = "ENG - GAR"
'_______________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 " _
     & " where (DOSSLDMSD <> DOSSLDGSD)   and substring(DOSSLDPCI , 1 , 5) in (" & sqlPCI & " )" _
     & " and DOSSLDSTA not in ('  ','90')" & whereYDOSSLD0 '   and DOSSLDSVC <> '01' "
Set rsSab = cnsab.Execute(xSql)
nbT = 0: nbD = 0

wLIB = "DOSSIER  : écart compta / gestion"
Do While Not rsSab.EOF
    wOPE = rsSab("DOSSLDOPE")
    wDOS = rsSab("DOSSLDNUM")
    wCli = rsSab("DOSSLDCLI")
    wMTD = 0 '
    wDev = rsSab("DOSSLDDEV")
    wErr = "# " & rsSab("DOSSLDPCI")
    wAmj = ""
        
    xSql = "select * from " & paramIBM_Library_SAB & ".ZCAUDOS0 " _
         & " where CAUDOSDOS = " & wDOS
    Set rsSabX = cnsab.Execute(xSql)
    If Not rsSabX.EOF Then
        wNAT = rsSabX("CAUDOSCAU")
        wMTD = rsSabX("CAUDOSMNT")
        wDev = rsSabX("CAUDOSDEV")
        wAmj = dateImp10(rsSabX("CAUDOSDEB") + 19000000)
    End If
    wNAT = ""
    wMTD = rsSab("DOSSLDMSD") - rsSab("DOSSLDGSD")
    If wMTD <> 0 Then
        wNAT = "- C#G"
        nbD = nbD + 1
           lstW.AddItem "02D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
        
    End If
    
    rsSab.MoveNext
Loop

wLIB = "Dossiers non clos en écart comptabilité / gestion : " & nbD
wCli = "PCI : " & listPCI
If nbD = 0 Then
    wLIB = "Dossiers non clos en écart comptabilité / gestion : NEANT"
    lstW.AddItem "02S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "02T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If
'_____________________________________________________________________________________________


xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSMVT0 " _
     & " where DOSMVTNUM = 0 and substring(DOSMVTPCI , 1 , 5) in (" & sqlPCI & " )"
Set rsSab = cnsab.Execute(xSql)

wLIB = "mvt comptable non affecté à un dossier "
nbT = 0: nbD = 0

Do While Not rsSab.EOF
    wOPE = rsSab("DOSMVTOPE")
    wDOS = rsSab("DOSMVTNUM")
    wCli = rsSab("DOSMVTCLI")
    wNAT = rsSab("DOSMVTEVE")
    wMTD = rsSab("DOSMVTMTD")
    wDev = rsSab("DOSMVTDEV")
    wErr = "? " & rsSab("DOSMVTPCI")
    wAmj = dateImp10(rsSab("DOSMVTDTR"))
    wLIB = "mvt comptable " & rsSab("DOSMVTPIE") & "-" & rsSab("DOSMVTECR") & " non affecté à un dossier "
    nbD = nbD + 1
    lstW.AddItem "03D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
    rsSab.MoveNext
Loop
wLIB = "Mouvements comptables orphelins, non affectés à un dossier : " & nbD
wCli = "PCI : " & listPCI
If nbD = 0 Then
    wLIB = "Mouvements comptables orphelins, non affectés à un dossier : NEANT"
    lstW.AddItem "03S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "03T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If
'_____________________________________________________________________________________________


xSql = "select *  from " & paramIBM_Library_SABSPE & ".YDOSSLD1 " _
     & " where (DOSSLDMSD <> DOSSLDGSD)   and substring(DOSSLDPCI , 1 , 5) in (" & sqlPCI & " ) "
Set rsSab = cnsab.Execute(xSql)
nbT = 0: nbD = 0

wLIB = "COMPTE : écart "
Do While Not rsSab.EOF
    wOPE = ""
    wDOS = 0
    wCli = rsSab("DOSSLDCLI")
    wMTD = 0 '
    wDev = rsSab("DOSSLDDEV")
    wErr = "* " & rsSab("DOSSLDPCI")
    wAmj = ""
    
    If wCli = "0050319" Then
        Debug.Print wCli
    End If
    wNAT = ""
    wMTD = rsSab("DOSSLDMSD") - rsSab("DOSSLDGSD")
    If wMTD <> 0 Then
        wNAT = "? C#G"
        wLIB = "COMPTE : écart - " & rsSab("DOSSLDNBV") & " dossiers"
        nbD = nbD + 1
        lstW.AddItem "04D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
    End If
    rsSab.MoveNext
Loop
wLIB = "total des mouvements comptables (PCI-client) <> total des soldes de gestion : " & nbD
wCli = "PCI : " & listPCI
If nbD = 0 Then
    wLIB = "total des mouvements comptables (PCI-client) <> total des soldes de gestion : NEANT"
    lstW.AddItem "04S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "04T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If
'_____________________________________________________________________________________________

xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 " _
     & " where DOSSLDOPE in ('ENG' , 'GAR') and DOSSLDMSD <> 0 and substring(DOSSLDPCI , 1 , 5) in (" & sqlPCI & " )" _
     & " and DOSSLDSTA  in ('  ','80','90')  and DOSSLDSVC = '03'" & whereYDOSSLD0
Set rsSab = cnsab.Execute(xSql)
nbT = 0: nbD = 0

wLIB = "dossier annulé en gestion ayant un solde comptable"
Do While Not rsSab.EOF
    wOPE = rsSab("DOSSLDOPE")
    wDOS = rsSab("DOSSLDNUM")
    wCli = rsSab("DOSSLDCLI")
    wMTD = rsSab("DOSSLDMSD") '
    wDev = rsSab("DOSSLDDEV")
    wErr = "! " & rsSab("DOSSLDPCI")
    wAmj = ""
    
    xSql = "select * from " & paramIBM_Library_SAB & ".ZCAUDOS0 " _
         & " where CAUDOSDOS = " & wDOS
    Set rsSabX = cnsab.Execute(xSql)
    If Not rsSabX.EOF Then
        wNAT = rsSabX("CAUDOSCAU")
        wMTD = rsSabX("CAUDOSMNT")
        wDev = rsSabX("CAUDOSDEV")
        wAmj = dateImp10(rsSabX("CAUDOSDEB") + 19000000)
    End If
    wNAT = ""
    wMTD = rsSab("DOSSLDMSD")
    If wMTD <> 0 Then
        wNAT = "! 90"
        nbD = nbD + 1
        lstW.AddItem "01D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
    End If
    
    rsSab.MoveNext
Loop
wLIB = "Dossiers clos en gestion mais présentant un solde comptable : " & nbD
wCli = "PCI : " & listPCI
If nbD = 0 Then
    wLIB = "Dossiers clos en gestion mais présentant un solde comptable : NEANT"
    lstW.AddItem "01S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "01T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If

'_______________________________________________________________________________________

nbT = 0: nbD = 0

xSql = "select * from " & paramIBM_Library_SAB & ".ZCAUDOS0 " _
     & " where CAUDOSTRA not in (2 , 4 , 6 ) " & whereZCAUDOS0
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    wOPE = "ENG - GAR"
    wCli = ""
    wDOS = rsSab("CAUDOSDOS")
    wNAT = rsSab("CAUDOSCAU")
    wMTD = rsSab("CAUDOSMNT")
    wDev = rsSab("CAUDOSDEV")
    wErr = "D " & rsSab("CAUDOSTRA")
    wAmj = dateImp10(rsSab("CAUDOSDEB") + 19000000)
    Select Case rsSab("CAUDOSTRA")
        Case 0: wLIB = "non validé": XSVC2 = "non validée"
        Case 1: wLIB = "non comptabilisé": XSVC2 = "non comptabilisée"
       Case Else: wLIB = rsSab("CAUDOSTRA") & " ???": XSVC2 = rsSab("CAUDOSTRA") & " ???"
    End Select
    nbD = nbD + 1
    lstW.AddItem "10D|" & wErr & "|" & wOPE & "|" & wDOS & "|" & wCli & "|" & wLIB & "|" & wNAT & "|" & wMTD & "|" & wDev & "|" & wAmj
    rsSab.MoveNext
Loop
xSql = "select count(*) as tally from " & paramIBM_Library_SAB & ".ZCAUDOS0 " _
     & " where CAUDOSTRA < 4 "
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then nbT = rsSab(0)

wLIB = "Dossiers en attente : " & nbD & " / " & nbT & " dossiers non clos"
wCli = "code état <> comptabilisé"
If nbD = 0 Then
    wLIB = "Evénements  en attente : NEANT / " & nbT & "  dossiers non clos"
    lstW.AddItem "10S|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
Else
    lstW.AddItem "10T|" & "" & "|" & "" & "|" & "" & "|" & wCli & "|" & wLIB & "|" & "" & "|" & "" & "|" & "" & "|" & ""
End If

'_______________________________________________________________________________________


'_____________________________________________________________________________________________

'_____________________________________________________________________________________________


Call YDOSSLD0_Export_CAU

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub cmdSelect_SQL_2_Exportation_Init()
Dim V, X As String, XSVC1 As String, XSVC2 As String
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean, blnSD0 As Boolean
Dim wCli As String, wOPE As String, wDOS As Long, wErr As String, wLIB As String, wNAT As String
Dim wMTD As Currency, wDev As String, wAmj As String, wPIE As Long, wECR As Long

Dim wMDB As Currency, wMCR As Currency
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2_Exportation_Init"
blnOk = False
lstW.Clear
'_______________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 " _
     & " where DOSSLDPCI like '" & Mid$(oldYBIACPT0.COMPTEOBL, 1, 5) & "%' and DOSSLDCLI = '" & oldYBIACPT0.CLIENACLI & "'" _
     & "  and DOSSLDDEV = '" & oldYBIACPT0.COMPTEDEV & "' and DOSSLDSTA <> '  '"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    If rsSab("DOSSLDSTA") = "01" And rsSab("DOSSLDSVC") = "01" Then
    Else
        wOPE = rsSab("DOSSLDOPE")
        wDOS = rsSab("DOSSLDNUM")
        If wOPE = "-RM" Or wOPE = "-RS" Then
            wOPE = "***": wDOS = 9999
        Else
            If wDOS < 70000 Or wDOS > 900000 Then wOPE = "***"
        End If
        
        lstW.AddItem wOPE & "|" & wDOS & "|9|" & "99999999" & "|" & "G??" & "|" & rsSab("DOSSLDGSD") & "|" & "" & "|" & rsSab("DOSSLDSTA") & "|" & rsSab("DOSSLDSVC") & "|"
    End If
    rsSab.MoveNext
Loop
'_______________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH" _
     & " where  MOUVEMCOM = '" & oldYBIACPT0.COMPTECOM & "'" _
     & " order by MOUVEMDTR,MOUVEMOPE,MOUVEMNUM,MOUVEMPIE,MOUVEMECR"
Set rsSab = cnsab.Execute(xSql)
blnSD0 = False
Do While Not rsSab.EOF
    wOPE = rsSab("MOUVEMOPE")
    wDOS = rsSab("MOUVEMNUM")
    wAmj = rsSab("MOUVEMDTR") + 19000000
    wPIE = rsSab("MOUVEMPIE")
    wECR = rsSab("MOUVEMECR")
    If wOPE = "*Z1" Then
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSXOD0 " _
          & " where DOSXODDTR = '" & wAmj & "'" _
          & " and DOSXODPIE = " & wPIE & " and DOSXODECR = " & wECR
         Set rsSabX = cnsab.Execute(xSql)
         If Not rsSabX.EOF Then
             wOPE = rsSabX("DOSXODOPE")
             wDOS = rsSabX("DOSXODNUM")
         End If
    End If
    If wOPE = "CDE" Or wOPE = "CDI" Then
        If wDOS < 70000 Or wDOS > 900000 Then wOPE = "***"
    Else
        wOPE = "***"
        wDOS = "9999"
    End If
    lstW.AddItem wOPE & "|" & wDOS & "|1|" & wAmj & "|" & rsSab("MOUVEMEVE") & "|" & -rsSab("MOUVEMMON") & "|" & rsSab("MOUVEMANU") & "|" & wPIE & "|" & wECR & "|" & rsSab("LIBELLIB1")
    If Not blnSD0 Then
        blnSD0 = True
        If rsSab("BIAMVTSD0") <> 0 Then
            lstW.AddItem "***" & "|" & "9999" & "|0|" & "20030316" & "|" & "-RS" & "|" & -rsSab("BIAMVTSD0") & "|" & " " & "|" & 0 & "|" & 0 & "|" & "reprise solde"
        End If
        
    End If
    
    rsSab.MoveNext
Loop
'_______________________________________________________________________________________


Call cmdSelect_SQL_2_Exportation_Xlsx
'_______________________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_2_Exportation_JPL()
Dim V, X As String, XSVC1 As String, XSVC2 As String
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean, blnSD0 As Boolean
Dim wCli As String, wOPE As String, wDOS As Long, wErr As String, wLIB As String, wNAT As String
Dim wMTD As Currency, wDev As String, wAmj As String, wPIE As Long, wECR As Long

Dim wMDB As Currency, wMCR As Currency
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2_Exportation_Init"
blnOk = False
lstW.Clear
oldYBIACPT0.COMPTEOBL = "985200"
oldYBIACPT0.COMPTEDEV = "EUR"
oldYBIACPT0.COMPTECOM = "985200EUR1RDE"
'_______________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 " _
     & " where DOSSLDPCI like '" & Mid$(oldYBIACPT0.COMPTEOBL, 1, 5) & "%'" _
     & "  and DOSSLDDEV = '" & oldYBIACPT0.COMPTEDEV & "' and DOSSLDSTA <> '  '"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    If rsSab("DOSSLDSTA") = "01" And rsSab("DOSSLDSVC") = "01" Then
    Else
        wOPE = rsSab("DOSSLDOPE")
        wDOS = rsSab("DOSSLDNUM")
        'If wOPE = "-RM" Or wOPE = "-RS" Then
        '    wOPE = "***": wDOS = 9999
        'Else
        '    If wDOS < 70000 Or wDOS > 900000 Then wOPE = "***"
        'End If
        
        lstW.AddItem wOPE & "|" & wDOS & "|9|" & "99999999" & "|" & "G??" & "|" & rsSab("DOSSLDGSD") & "|" & "" & "|" & rsSab("DOSSLDSTA") & "|" & rsSab("DOSSLDSVC") & "|"
    End If
    rsSab.MoveNext
Loop
'_______________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH" _
     & " where  MOUVEMCOM = '" & oldYBIACPT0.COMPTECOM & "'" _
     & " order by MOUVEMDTR,MOUVEMOPE,MOUVEMNUM,MOUVEMPIE,MOUVEMECR"
Set rsSab = cnsab.Execute(xSql)
blnSD0 = False
Do While Not rsSab.EOF
    wOPE = rsSab("MOUVEMOPE")
    wDOS = Fix(rsSab("MOUVEMNUM") / 100)
    wAmj = rsSab("MOUVEMDTR") + 19000000
    wPIE = rsSab("MOUVEMPIE")
    wECR = rsSab("MOUVEMECR")
    If wOPE = "*Z1" Then
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSXOD0 " _
          & " where DOSXODDTR = '" & wAmj & "'" _
          & " and DOSXODPIE = " & wPIE & " and DOSXODECR = " & wECR
         Set rsSabX = cnsab.Execute(xSql)
         If Not rsSabX.EOF Then
             wOPE = rsSabX("DOSXODOPE")
             wDOS = rsSabX("DOSXODNUM")
         End If
    End If
    'If wOPE = "CDE" Or wOPE = "CDI" Then
    '    If wDOS < 70000 Or wDOS > 900000 Then wOPE = "***"
    'Else
    If wOPE = "RDE" And wDOS < 904000 Then
        wOPE = "***"
        wDOS = "9999"
    End If
    lstW.AddItem wOPE & "|" & wDOS & "|1|" & wAmj & "|" & rsSab("MOUVEMEVE") & "|" & -rsSab("MOUVEMMON") & "|" & rsSab("MOUVEMANU") & "|" & wPIE & "|" & wECR & "|" & rsSab("LIBELLIB1")
    If Not blnSD0 Then
        blnSD0 = True
        If rsSab("BIAMVTSD0") <> 0 Then
            lstW.AddItem "***" & "|" & "9999" & "|0|" & "20030316" & "|" & "-RS" & "|" & -rsSab("BIAMVTSD0") & "|" & " " & "|" & 0 & "|" & 0 & "|" & "reprise solde"
        End If
        
    End If
    
    rsSab.MoveNext
Loop
'_______________________________________________________________________________________


Call cmdSelect_SQL_2_Exportation_Xlsx
'_______________________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub



Private Sub cmdSelect_SQL_6()
Dim V, X As String, XSVC1 As String, XSVC2 As String
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim blnCli As Boolean, blnDos As Boolean
Dim wDev As String, wCli As String, wOPE As String, wDOS As Long, wSTA As String
Dim mDev As String, mCli As String, mOPE As String, mDOS As Long
Dim wMTD As Currency
Dim blnOk As Boolean
Dim wDTR As Long
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_6_Exportation_Init"
blnCli = False
lstW.Clear
Call DTPicker_Control(txtSelect_6_AMJMin, wAmjMin)

wAmjMax = YBIATAB0_DATE_CPT_J
sqlPCI = Trim(txtSelect_6_PCI)
sqlPCI_Len = Len(sqlPCI)
If sqlPCI_Len < 5 Then
    V = "Préciser 5 ou 6 chiffres pour le PCI sélectionné"
    GoTo Error_MsgBox
End If

blnProvisions_Control = False: mProvisions_Control_Ope = "CDE": mProvisions_Control_PCI = ""
Select Case Mid$(sqlPCI, 1, 5)
    Case "13221": blnProvisions_Control = True: mProvisions_Control_PCI = "91120"
    Case "25302": blnProvisions_Control = True: mProvisions_Control_Ope = "CDI": mProvisions_Control_PCI = "90312"
    Case "90312", "98755", "98756": mProvisions_Control_Ope = "CDI"
    
End Select


sqlCLI = Trim(txtSelect_6_CLIEANCLI)
If sqlCLI <> "" Then
    sqlCLI = Format(sqlCLI, "0000000")
    xWhere = " and CLIENACLI = '" & sqlCLI & "'"
End If
mCDODOSDOS_Nb = 0: mCDODOSPCC_Nb = 0: mCDODOSPDE_Nb = 0: mCDODOSXXX_Nb = 0
mCDODOSVAL_Nb = 0

'_______________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '" & sqlPCI & "%'" & xWhere _
     & "  order by COMPTECOM"
Set rsSab = cnsab.Execute(xSql)

'_______________________________________________________________________________________
Do While Not rsSab.EOF
    wDev = rsSab("COMPTEDEV")
    wCli = rsSab("CLIENACLI")
    wOPE = rsSab("COMPTECOM")
    wMTD = -rsSab("SOLDECEN") / 1000
    wDTR = rsSab("SOLDEDMO") + 19000000
    
    lstW.AddItem wCli & "|" & wDev & "|" & "0" & "|" & wOPE & "|" & rsSab("COMPTEINT") & "|" & wDTR & "|" & wMTD & "|" & rsSab("COMPTEFON") & "|0|"
    
    rsSab.MoveNext
Loop


'_______________________________________________________________________________________
If sqlCLI <> "" Then xWhere = " and DOSMVTCLI = '" & sqlCLI & "'"

xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSMVT0 ," _
     & paramIBM_Library_SABSPE & ".YDOSSLD0 " _
     & " where DOSMVTPCI like '" & sqlPCI & "%'" & xWhere _
     & " and DOSSLDOPE = DOSMVTOPE  and DOSSLDNUM = DOSMVTNUM and DOSSLDDEV = DOSMVTDEV and DOSSLDPCI = DOSMVTPCI and DOSSLDCLI = DOSMVTCLI " _
     & "  order by DOSMVTDEV,DOSMVTCLI,DOSMVTOPE,DOSMVTNUM,DOSMVTDTR"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    wDev = rsSab("DOSMVTDEV")
    wCli = rsSab("DOSMVTCLI")
    wOPE = rsSab("DOSMVTOPE")
    wDOS = rsSab("DOSMVTNUM")
    wMTD = rsSab("DOSMVTMTD")
    wDTR = rsSab("DOSMVTDTR")
    wSTA = rsSab("DOSSLDSTA")
    If wDTR <= wAmjMax Then
        If wSTA = "90" Or wSTA = "  " Or wDOS = 0 Then
            lstW.AddItem wCli & "|" & wDev & "|" & "1" & "|" & wOPE & "|" & wDOS & "|" & wDTR & "|" & wMTD & "|" & rsSab("DOSMVTPIE") & "|" & rsSab("DOSMVTECR") & "|"
        Else
            lstW.AddItem wCli & "|" & wDev & "|" & "2" & "|" & wOPE & "|" & wDOS & "|" & wDTR & "|" & wMTD & "|" & rsSab("DOSMVTPIE") & "|" & rsSab("DOSMVTECR") & "|"
        End If
    End If
    
    rsSab.MoveNext
Loop




'_______________________________________________________________________________________
If blnProvisions_Control Then

    xSql = "select * from " & paramIBM_Library_SAB & ".ZCDOREG0 ," _
         & paramIBM_Library_SABSPE & ".YBIACPT0 " _
         & " where CDOREGCRD = 'R'and CDOREGCOP = '" & mProvisions_Control_Ope & "'and CDOREGDCR = 0" _
         & " and CDOREGETA = '02' and CDOREGCOM = COMPTECOM " _
         & "  order by CDOREGCOM"
    Set rsSab = cnsab.Execute(xSql)
    
    Do While Not rsSab.EOF
        wDev = rsSab("CDOREGDEV")
        wOPE = rsSab("CDOREGCOP")
        wDOS = rsSab("CDOREGDOS")
        wMTD = -rsSab("CDOREGMON")
        wDTR = rsSab("CDOREGDRE") + 19000000
        wSTA = "REG"
       
        wCli = rsSab("CLIENACLI")
        If sqlCLI = "" Then
            lstW.AddItem wCli & "|" & wDev & "|" & "2" & "|" & wOPE & "|" & wDOS & "|" & wDTR & "|" & wMTD & "|0|0|"
        Else
            If sqlCLI = wCli Then lstW.AddItem wCli & "|" & wDev & "|" & "2" & "|" & wOPE & "|" & wDOS & "|" & wDTR & "|" & wMTD & "|0|0|"
        End If
        
        rsSab.MoveNext
    Loop
End If
'_______________________________________________________________________________________

Call cmdSelect_SQL_6_Exportation_Xlsx
'_______________________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Public Sub cmdSelect_SQL_6_Exportation_Xlsx()
On Error GoTo Error_Handler
Dim xSql As String
Dim X As String, XX As String, I As Integer
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, kLen As Integer
Dim K5 As Integer, K6 As Integer, K7 As Integer, K8 As Integer, K9 As Integer
Dim K10 As Integer, K11 As Integer
Dim wK As Integer
Dim wColorzSD9 As Long, xCur As Currency
Dim blnCALCS As Boolean

On Error GoTo Error_Handler
'===================================================================================
If blnAuto Then
    X = paramServer("\\CDO_Archive\")
Else
    X = ""
End If
If X = "" Then X = "C:\Temp\"
If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"


blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True


XX = IIf(sqlCLI = "", "", "-" & sqlCLI)

wFile = X & Trim("CDO PCI " & sqlPCI & XX & "-" & YBIATAB0_DATE_CPT_J & ".xlsx")
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut : " & wFile _
        & vbCrLf & vbCrLf & "     =========================" _
        & vbCrLf & "     =========================", "Crédits Documentaires : Etat xxx  : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If
'_________________________________________


If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents
'=========================================================================================

Call rsYDOSMVT0_Init(oldYDOSMVT0)
newYDOSMVT0 = oldYDOSMVT0
xYDOSMVT0 = oldYDOSMVT0
Call rsYBIACPT0_Init(oldYBIACPT0)
mMTD0_Dos = 0: mMTD9_Dos = 0: mMTDJ_Dos = 0: blnDos_Ok = False
mMTD0_Cli = 0: mMTD9_Cli = 0: blnCli_Ok = False
mAnn_Nb = 0: mProv_Nb = 0: mDOSSLDMSD_Nb = 0: mCDODOSPDE_Nb = 0: mCDODOSPCC_Nb = 0: mCDODOSDOS_Nb = 0: mCDODOSXXX_Nb = 0
mXls2_Row_D = 0

xSql = "select count(*) as Tally from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
     & " where CDODOSCOP = '" & mProvisions_Control_Ope & "' and CDODOSEVE <> '90' and CDODOSPPO <> 0"
Set rsSab = cnsab.Execute(xSql)


ReDim arrZCDODOS0(rsSab(0) + 1)

xSql = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
     & " where CDODOSCOP = '" & mProvisions_Control_Ope & "' and CDODOSEVE <> '90' and CDODOSPPO <> 0" _
     & "  order by CDODOSDOS"
Set rsSab = cnsab.Execute(xSql)

arrZCDODOS0_Nb = 0
'_______________________________________________________________________________________
Do While Not rsSab.EOF
    arrZCDODOS0_Nb = arrZCDODOS0_Nb + 1
    V = rsZCDODOS0_GetBuffer(rsSab, arrZCDODOS0(arrZCDODOS0_Nb))
    rsSab.MoveNext
Loop

'=========================================================================================
'=========================================================================================
 
xSql = "select count(*) as Tally from " & paramIBM_Library_SABSPE & ".YDOSSLD0 " _
     & " where substring(DOSSLDPCI , 1 , 5) = '" & mProvisions_Control_PCI & "' and DOSSLDSTA <> '90'"
Set rsSab = cnsab.Execute(xSql)


ReDim arrYDOSSLD0(rsSab(0) + 1)

xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 " _
     & " where substring(DOSSLDPCI , 1 , 5) = '" & mProvisions_Control_PCI & "' and DOSSLDSTA <> '90'" _
     & "  order by DOSSLDNUM"
Set rsSab = cnsab.Execute(xSql)

arrYDOSSLD0_Nb = 0
'_______________________________________________________________________________________
Do While Not rsSab.EOF
    arrYDOSSLD0_Nb = arrYDOSSLD0_Nb + 1
    V = rsYDOSSLD0_GetBuffer(rsSab, arrYDOSSLD0(arrYDOSSLD0_Nb))
    rsSab.MoveNext
Loop

'=========================================================================================



Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "DOSCDO"
    .Subject = ""
End With

'__________________________________________________________________________________

Set wsExcel = wbExcel.Sheets(1)
wsExcel.Name = "Récapitulatif"


With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .WrapText = True
    .Font.Size = 7
    .Font.Name = "Arial Unicode MS"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 75

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Récapitulatif  PCI : " & sqlPCI & ", Client : " & sqlCLI & ", période du " & dateImp10(wAmjMin) & " au " & dateImp10(wAmjMax) _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.CenterHorizontally = True

wsExcel.PageSetup.PrintTitleRows = "$A1:$F1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

wsExcel.Columns(1).ColumnWidth = 6: wsExcel.Cells(1, 1) = "Code"
wsExcel.Columns(2).ColumnWidth = 5: wsExcel.Cells(1, 2) = "n° ligne détail"
wsExcel.Columns(3).ColumnWidth = 15: wsExcel.Cells(1, 3) = "Compte"
wsExcel.Columns(4).ColumnWidth = 30: wsExcel.Cells(1, 4) = "Intitulé"
wsExcel.Columns(5).ColumnWidth = 10: wsExcel.Cells(1, 5) = "Référence"
wsExcel.Columns(6).ColumnWidth = 50: wsExcel.Cells(1, 6) = "Informations"


For K = 1 To 6
    wsExcel.Cells(1, K).Interior.Color = mColor_GB ' RGB(255, 128, 50)
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next K
mXls1_Row = 1
X = ", Période du " & dateImp10(wAmjMin) & " au " & dateImp10(wAmjMax)
XX = "Récapitulatif de la feuille DETAIL ( " & dateImp10(DSys) & " " & Time & ")"
Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("", 0, "PCI : " & sqlPCI, "Client : " & sqlCLI & X, "", XX)


'Set wsExcel = wbExcel.ActiveSheet
Set wsExcel = wbExcel.Sheets(2)

wsExcel.Name = "Détail"

'__________________________________________________________________________________
wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 80

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Liste  PCI : " & sqlPCI & ", Client : " & sqlCLI & ", période du " & dateImp10(wAmjMin) & " au " & dateImp10(wAmjMax) _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.CenterHorizontally = True

wsExcel.PageSetup.PrintTitleRows = "$A1:$H1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .WrapText = True
    .Font.Size = 8
    .Font.Name = "Arial Unicode MS"
    .RowHeight = 17
End With


wsExcel.Columns(1).ColumnWidth = 18: wsExcel.Cells(1, 1) = "Référence"
wsExcel.Columns(2).ColumnWidth = 15: wsExcel.Cells(1, 2) = "Solde " & dateImp10(wAmjMin): wsExcel.Columns(2).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(3).ColumnWidth = 15: wsExcel.Cells(1, 3) = "Credit": wsExcel.Columns(3).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(4).ColumnWidth = 15: wsExcel.Cells(1, 4) = "Débit": wsExcel.Columns(4).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(5).ColumnWidth = 15: wsExcel.Cells(1, 5) = "Solde " & dateImp10(wAmjMax): wsExcel.Columns(5).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(6).ColumnWidth = 12: wsExcel.Cells(1, 6) = "D TRT": wsExcel.Columns(6).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(7).ColumnWidth = 50: wsExcel.Cells(1, 7) = "Libellé"
wsExcel.Columns(8).ColumnWidth = 15: wsExcel.Cells(1, 8) = "Validité": wsExcel.Columns(8).HorizontalAlignment = Excel.xlHAlignCenter


For K = 1 To 8
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next K




mXls2_Row = 1
mXls2_Row_Cli = 1
For I = 0 To lstW.ListCount - 1
    lstW.ListIndex = I
    X = Trim(lstW.Text): kLen = Len(X)
    K1 = InStr(1, X, "|"): newYDOSMVT0.DOSMVTCLI = Mid$(X, 1, K1 - 1)
    K2 = InStr(K1 + 1, X, "|"): newYDOSMVT0.DOSMVTDEV = Mid$(X, K1 + 1, K2 - K1 - 1)
    K3 = InStr(K2 + 1, X, "|"): wK = Mid$(X, K2 + 1, K3 - K2 - 1)
    K4 = InStr(K3 + 1, X, "|"): xYDOSMVT0.DOSMVTOPE = Mid$(X, K3 + 1, K4 - K3 - 1)
    K5 = InStr(K4 + 1, X, "|"):
        If wK = 0 Then
            oldYBIACPT0.COMPTEINT = Mid$(X, K4 + 1, K5 - K4 - 1)
        Else
            xYDOSMVT0.DOSMVTNUM = Mid$(X, K4 + 1, K5 - K4 - 1)
        End If
    K6 = InStr(K5 + 1, X, "|"): newYDOSMVT0.DOSMVTDTR = Mid$(X, K5 + 1, K6 - K5 - 1)
    K7 = InStr(K6 + 1, X, "|"): XX = Mid$(X, K6 + 1, K7 - K6 - 1)
        If Trim(XX) = "" Then
            wMTD = 0
        Else
            wMTD = CCur(XX)
        End If

    K8 = InStr(K7 + 1, X, "|"): wPIE = Mid$(X, K7 + 1, K8 - K7 - 1)
    K9 = InStr(K8 + 1, X, "|"): wECR = Val(Mid$(X, K8 + 1, K9 - K8 - 1))
    ''K10 = InStr(K9 + 1, X, "|"): wECR = Val(Mid$(X, K9 + 1, K10 - K9 - 1))
    
'_______________________________________________________________________________
        Select Case wK
            Case 0: Call lstErr_ChangeLastItem(lstErr, cmdContext, "Compte : " & xYDOSMVT0.DOSMVTOPE): DoEvents

            Case 1: newYDOSMVT0.DOSMVTOPE = "***": newYDOSMVT0.DOSMVTNUM = 0
            Case 2: newYDOSMVT0.DOSMVTOPE = xYDOSMVT0.DOSMVTOPE: newYDOSMVT0.DOSMVTNUM = xYDOSMVT0.DOSMVTNUM
        End Select

'=========================================================================================
    
    If oldYDOSSLD0.DOSSLDDEV <> newYDOSMVT0.DOSMVTDEV Or oldYDOSSLD0.DOSSLDCLI <> newYDOSMVT0.DOSMVTCLI Then
    
        Call cmdSelect_SQL_6_Exportation_Xlsx_Dos

        Call cmdSelect_SQL_6_Exportation_Xlsx_Cli
'_______________________________________________________________________________
        mXls2_Row = mXls2_Row + 2
        mXls2_Row_Cli = mXls2_Row
        wsExcel.Cells(mXls2_Row, 6) = newYDOSMVT0.DOSMVTDEV: wsExcel.Cells(mXls2_Row, 6).Font.Bold = True
        If wK = 0 Then
            oldYBIACPT0.SOLDECEN = wMTD
            oldYBIACPT0.COMPTECOM = xYDOSMVT0.DOSMVTOPE
            wsExcel.Cells(mXls2_Row, 1) = xYDOSMVT0.DOSMVTOPE: wsExcel.Cells(mXls2_Row, 1).Font.Bold = True
            wsExcel.Cells(mXls2_Row, 1).Font.Size = 7
            wsExcel.Cells(mXls2_Row, 5) = wMTD: wsExcel.Cells(mXls2_Row, 5).Font.Bold = True
            wsExcel.Cells(mXls2_Row, 5).Font.Size = 7
            wsExcel.Cells(mXls2_Row, 7) = oldYBIACPT0.COMPTEINT: wsExcel.Cells(mXls2_Row, 7).Font.Bold = True
            wsExcel.Cells(mXls2_Row, 7).Font.Size = 7
        Else

            wsExcel.Cells(mXls2_Row, 1) = newYDOSMVT0.DOSMVTCLI: wsExcel.Cells(mXls2_Row, 1).Font.Bold = True
            wsExcel.Cells(mXls2_Row, 7) = "? COMPTE INCONNU": wsExcel.Cells(mXls2_Row, 3).Font.Bold = True
            Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("#Cpt", mXls2_Row, newYDOSMVT0.DOSMVTDEV & " " & newYDOSMVT0.DOSMVTCLI, "", newYDOSMVT0.DOSMVTOPE & " " & newYDOSMVT0.DOSMVTNUM, "? COMPTE INCONNU")

        End If
        For K = 1 To 7:    wsExcel.Cells(mXls2_Row, K).Interior.Color = RGB(230, 230, 230): Next K
        oldYDOSSLD0.DOSSLDDEV = newYDOSMVT0.DOSMVTDEV: oldYDOSSLD0.DOSSLDCLI = newYDOSMVT0.DOSMVTCLI: mMTD0_Cli = 0: mMTD9_Cli = 0: blnCli_Ok = False
    End If
'_________________________________________________________________________________________
    If wK = 0 Then
        oldYDOSMVT0.DOSMVTOPE = "": oldYDOSMVT0.DOSMVTNUM = 0: mMTD0_Dos = 0: mMTD9_Dos = 0: mMTDJ_Dos = 0: blnDos_Ok = False
    Else
        If oldYDOSMVT0.DOSMVTOPE <> newYDOSMVT0.DOSMVTOPE Or oldYDOSMVT0.DOSMVTNUM <> newYDOSMVT0.DOSMVTNUM Then
            Call cmdSelect_SQL_6_Exportation_Xlsx_Dos
                    
        End If
'_________________________________________________________________________________________

        If newYDOSMVT0.DOSMVTDTR < wAmjMin Then
            mMTD0_Cli = mMTD0_Cli + wMTD
            mMTD0_Dos = mMTD0_Dos + wMTD
        Else
            mMTD9_Dos = mMTD9_Dos + wMTD
            
            mXls2_Row = mXls2_Row + 1
            If Not blnDos_Ok Then wsExcel.Cells(mXls2_Row, 2) = mMTD0_Dos
            
            If mXls2_Row_D = 0 Then
                mXls2_Row_D = mXls2_Row
                If oldYDOSMVT0.DOSMVTOPE = "***" Then
                    wsExcel.Cells(mXls2_Row, 1) = oldYDOSMVT0.DOSMVTOPE & " " & xYDOSMVT0.DOSMVTNUM
                Else
                    wsExcel.Cells(mXls2_Row, 1) = oldYDOSMVT0.DOSMVTOPE & " " & oldYDOSMVT0.DOSMVTNUM
                End If
            End If
            
            wsExcel.Cells(mXls2_Row, 6) = dateImp10(newYDOSMVT0.DOSMVTDTR): wsExcel.Cells(mXls2_Row, 6).Font.Size = 7
            If wMTD < 0 Then
                wsExcel.Cells(mXls2_Row, 3) = wMTD: wsExcel.Cells(mXls2_Row, 3).Font.Size = 7
            Else
                wsExcel.Cells(mXls2_Row, 4) = wMTD: wsExcel.Cells(mXls2_Row, 4).Font.Size = 7
            End If
            blnDos_Ok = True
            If wPIE > 0 Then
                xSql = "select * from " & paramIBM_Library_SAB & ".ZLIBEL0 " _
                 & " where LIBELETA = 1" _
                 & " and LIBELPIE = " & wPIE & " and LIBELECR = " & wECR & " order by LIBELNUM"
                Set rsSabX = cnsab.Execute(xSql)
                X = ""
                
                Do While Not rsSabX.EOF
                    X = X & rsSabX("LIBELLIB")
                    rsSabX.MoveNext
                Loop
                wsExcel.Cells(mXls2_Row, 7) = X
                wsExcel.Cells(mXls2_Row, 7).Font.Size = 7
            End If
            
            'For K = 1 To 8:    wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_Y0: Next K

            If newYDOSMVT0.DOSMVTDTR > YBIATAB0_DATE_CPT_J Then
                wsExcel.Cells(mXls2_Row, 6).Font.Color = mColor_W1
                wsExcel.Cells(mXls2_Row, 6).Interior.Color = mColor_Y1
            Else
                '====================================
                 mMTD9_Cli = mMTD9_Cli + wMTD
                 mMTDJ_Dos = mMTDJ_Dos + wMTD

                '====================================
            End If

        End If
    End If
Next I
'_______________________________________________________________________________
Call cmdSelect_SQL_6_Exportation_Xlsx_Dos
Call cmdSelect_SQL_6_Exportation_Xlsx_Cli

'__________________________________________________________________________________

mXls2_Row = mXls2_Row + 1
For K = 1 To 8
    wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_GB
Next K
'__________________________________________________________________________________
K1 = mXls1_Row - 2

If blnProvisions_Control Then

    If sqlCLI = "" Then cmdSelect_SQL_6_Exportation_Xlsx_Recap_End
End If




Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("", 0, "", "", "", K1 & " informations à contrôler.")

'__________________________________________________________________________________

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

End Sub

Public Sub cmdSelect_SQL_Xc()
On Error GoTo Error_Handler
Dim xSql As String
Dim X As String, XX As String, I As Integer
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, kLen As Integer
Dim mSolde As Currency, mDev As String, xDev As String, mDev_K As Integer
Dim blnOk As Boolean
Dim blnCALCS As Boolean

On Error GoTo Error_Handler
'===================================================================================
'______________________________________________'
If blnAuto Then
    X = paramServer("\\CDO_Archive\")
    wAmjMin = YBIATAB0_DATE_CPT_J
Else
    X = Trim(cboSelect_DOSCD7DSIT)
    wAmjMin = Mid$(X, 1, 4) & Mid$(X, 6, 2) & Mid$(X, 9, 2)
    X = ""
End If

blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

If X = "" Then X = "C:\Temp\"
If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"

Call DTPicker_Control(txtSelect_DOSCD7DAN, wAmjMax)
wDIBM_Max = wAmjMax - 19000000
wDIBM_Min = wAmjMin - 19000000
If wDIBM_Max < wDIBM_Min Then
    Call MsgBox("dates non cohérentes", vbExclamation, "Commissions à provisionner")
    Exit Sub
End If

mSOLDE_K = (Mid$(YBIATAB0_DATE_CPT_JS1, 5, 2) - 1) - (Mid$(wAmjMin, 5, 2) - 1) + (Mid$(YBIATAB0_DATE_CPT_JS1, 1, 4) - Mid$(wAmjMin, 1, 4)) * 12

wDMS_Min = dateImp_Amj(wAmjMin)
wFile = X & Trim("CDO commissions à recevoir au " & wAmjMin & ".xlsx")
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "Crédits Documentaires : Commissions à provisionner  : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If
'_________________________________________


If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile

Call lstErr_AddItem(lstErr, cmdContext, "Dossiers annulés  : "): DoEvents

'=========================================================================================
cmdSelect_SQL_5_ECNFPT_arr  '$JPL 2013-07-09
'=========================================================================================

xSql = "select count(*)  from " & paramIBM_Library_SAB & ".ZCDOMOD0 " _
     & " where CDOMODCOP = 'CDE' and CDOMODEVE = '07'"
Set rsSab = cnsab.Execute(xSql)

arrCDOMODEVE_07_Nb = rsSab(0)
ReDim arrCDOMODEVE_07(arrCDOMODEVE_07_Nb + 1)

xSql = "select CDOMODDOS from " & paramIBM_Library_SAB & ".ZCDOMOD0 " _
     & " where CDOMODCOP = 'CDE' and CDOMODEVE = '07'" _
     & " order by CDOMODDOS"
Set rsSab = cnsab.Execute(xSql)
K = 0
Do While Not rsSab.EOF
    K = K + 1
    arrCDOMODEVE_07(K) = rsSab("CDOMODDOS")
    rsSab.MoveNext
Loop

'=========================================================================================

xSql = "select count(*) as tally from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
     & " where CDODOSEVE in ('80','90') and CDODOSDAN > " & wDIBM_Min & " and CDODOSDAN <= " & wDIBM_Max
Set rsSab = cnsab.Execute(xSql)

arrZCDODOS0_Nb = rsSab(0)
ReDim arrZCDODOS0(arrZCDODOS0_Nb + 1)

xSql = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
     & " where CDODOSEVE in ('80','90') and CDODOSDAN > " & wDIBM_Min & " and CDODOSDAN <= " & wDIBM_Max _
     & " order by CDODOSDOS"
Set rsSab = cnsab.Execute(xSql)
K = 0
Do While Not rsSab.EOF
    K = K + 1
    arrZCDODOS0(K).CDODOSCOP = rsSab("CDODOSCOP")
    arrZCDODOS0(K).CDODOSDOS = rsSab("CDODOSDOS")
    arrZCDODOS0(K).CDODOSDAN = rsSab("CDODOSDAN") + 19000000
    arrZCDODOS0(K).CDODOSMOT = 0
    arrZCDODOS0(K).CDODOSMOC = 0
    rsSab.MoveNext
Loop

For I = 1 To arrZCDODOS0_Nb
        Debug.Print "COP=" & arrZCDODOS0(I).CDODOSCOP
        Debug.Print "DOS=" & arrZCDODOS0(I).CDODOSDOS
        Debug.Print "DAN=" & arrZCDODOS0(I).CDODOSDAN
        Debug.Print "MOC=" & arrZCDODOS0(I).CDODOSMOC
        Debug.Print "MOT=" & arrZCDODOS0(I).CDODOSMOT
        Exit For
Next I



' ! détournement des champs CDODOSMOT et CDODOSMOC pour cumul des commissions perçues
' entre la date d'arrêté et la date limite de la prise en compte des annulations
Call lstErr_AddItem(lstErr, cmdContext, "Commissions / dossiers annulés : "): DoEvents

xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSMVT0 " _
     & " where DOSMVTPCI in ('707210','707212') and DOSMVTDTR > " & wAmjMin & " and DOSMVTDTR <= " & wAmjMax _
     & " order by DOSMVTOPE,DOSMVTNUM"
Set rsSab = cnsab.Execute(xSql)
K = 0
oldYDOSMVT0.DOSMVTOPE = ""
oldYDOSMVT0.DOSMVTNUM = 0

Do While Not rsSab.EOF
    xYDOSMVT0.DOSMVTOPE = rsSab("DOSMVTOPE")
    xYDOSMVT0.DOSMVTNUM = rsSab("DOSMVTNUM")
    If oldYDOSMVT0.DOSMVTOPE <> xYDOSMVT0.DOSMVTOPE Or oldYDOSMVT0.DOSMVTNUM <> xYDOSMVT0.DOSMVTNUM Then
        K = 0: blnOk = False
        For K = 1 To arrZCDODOS0_Nb
            If arrZCDODOS0(K).CDODOSCOP = xYDOSMVT0.DOSMVTOPE And arrZCDODOS0(K).CDODOSDOS = xYDOSMVT0.DOSMVTNUM Then
                blnOk = True
                Exit For
            End If
        Next K
        If Not blnOk Then K = 0
        oldYDOSMVT0.DOSMVTOPE = xYDOSMVT0.DOSMVTOPE
        oldYDOSMVT0.DOSMVTNUM = xYDOSMVT0.DOSMVTNUM
    End If
    
    Select Case Trim(rsSab("DOSMVTPCI"))
        Case "707210": arrZCDODOS0(K).CDODOSMOC = arrZCDODOS0(K).CDODOSMOC + rsSab("DOSMVTMTD")
        Case "707212": arrZCDODOS0(K).CDODOSMOT = arrZCDODOS0(K).CDODOSMOT + rsSab("DOSMVTMTD")
    End Select
    
    rsSab.MoveNext
Loop


'_______________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Fichier excel.... : "): DoEvents


Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "DOSCDO"
    .Subject = "commissions à recevoir"
End With

'__________________________________________________________________________________

appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add

Set wsExcel = wbExcel.Sheets(1): wsExcel.Name = "Recapitulatif"
Set wsExcel = wbExcel.Sheets(2): wsExcel.Name = "ECNF"
Set wsExcel = wbExcel.Sheets(3): wsExcel.Name = "ENOTIF"
Set wsExcel = wbExcel.Sheets(4): wsExcel.Name = "91120"
Set wsExcel = wbExcel.Sheets(5): wsExcel.Name = "98050"
Set wsExcel = wbExcel.Sheets(6): wsExcel.Name = "91130"

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
    .Font.Name = "Arial Unicode MS"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 85

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Liste des commissions (ECNF, ENOTIF, ELVD) à provisionner, arrêté au " & dateImp10(wAmjMin) _
                                & vbCr & "&B&U&10(en excluant les dossiers annulés jusqu'au " & dateImp10(wAmjMax) & ")" & vbCr

If optSelect_DOSCD7DAN_In = True Then
    wsExcel.PageSetup.CenterHeader = Replace(wsExcel.PageSetup.CenterHeader, "excluant", "incluant")
End If

wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$J1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

mXls1_Col = arrDev_Nb + 2
mXls1_Row_T = 1
mXls1_Row_C = 1
mXls1_Row_N = arrDev_Nb + 2
mXls1_Row_D = mXls1_Row_N + arrDev_Nb + 1


'wsExcel.Rows(1).RowHeight = 34
'wsExcel.Rows(1).HorizontalAlignment = Excel.xlHAlignCenter
'wsExcel.Rows(1).VerticalAlignment = Excel.xlVAlignCenter
'wsExcel.Cells(1, 5) = "Etat des commissions (ECNF, ENOTIF, ELVD) à provisionner, arrêté au " & dateImp10(wAmjMin)
'wsExcel.Cells(1, 5).Font.Bold = True

'wsExcel.Rows(2).RowHeight = 25
'wsExcel.Rows(2).HorizontalAlignment = Excel.xlHAlignCenter
'wsExcel.Rows(2).VerticalAlignment = Excel.xlVAlignCenter
'wsExcel.Cells(2, 5) = " (en excluant les dossiers annulés jusqu'au " & dateImp10(wAmjMax) & ")                  "
'For K = 2 To 8
'    wsExcel.Cells(1, K).Interior.Color = mColor_Y1
'    wsExcel.Cells(2, K).Interior.Color = mColor_Y1
'Next K

'Ici Sofyan le 16/05/2025

wsExcel.Columns(1).ColumnWidth = 7: wsExcel.Cells(mXls1_Row_C, 1) = "ECNF": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 14: wsExcel.Cells(mXls1_Row_C, 2) = "Balance " & dateImp10(wAmjMin): wsExcel.Columns(2).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(3).ColumnWidth = 14: wsExcel.Cells(mXls1_Row_C, 3) = "Total Eng": wsExcel.Columns(3).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(4).ColumnWidth = 11: wsExcel.Cells(mXls1_Row_C, 4) = "Total COM": wsExcel.Columns(4).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(5).ColumnWidth = 13: wsExcel.Cells(mXls1_Row_C, 5) = "Total perçu": wsExcel.Columns(5).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(6).ColumnWidth = 11: wsExcel.Cells(mXls1_Row_C, 6) = "Total à percevoir": wsExcel.Columns(6).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(7).ColumnWidth = 11: wsExcel.Cells(mXls1_Row_C, 7) = "G =Total PDIF": wsExcel.Columns(7).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(8).ColumnWidth = 11: wsExcel.Cells(mXls1_Row_C, 8) = "H<= " & dateImp10(wAmjMin): wsExcel.Columns(8).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(9).ColumnWidth = 11: wsExcel.Cells(mXls1_Row_C, 9) = "> " & dateImp10(wAmjMin): wsExcel.Columns(9).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(10).ColumnWidth = 7: wsExcel.Cells(mXls1_Row_C, 10) = "G +% H ": wsExcel.Columns(10).NumberFormat = "##0.00"
wsExcel.Columns(11).ColumnWidth = 11: wsExcel.Cells(mXls1_Row_C, 11) = "à provisionner": wsExcel.Columns(11).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(12).ColumnWidth = 11: wsExcel.Cells(mXls1_Row_C, 12) = "C perçues/ANN": wsExcel.Columns(12).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"


wsExcel.Cells(mXls1_Row_C, 1).Interior.Color = mColor_Y1
wsExcel.Cells(mXls1_Row_N, 1).Interior.Color = mColor_Y1
wsExcel.Cells(mXls1_Row_D, 1).Interior.Color = mColor_Y1

For K = 2 To 12
    wsExcel.Cells(mXls1_Row_C, K).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_C, K).Font.Color = mColor_Z0
    
    wsExcel.Cells(mXls1_Row_N, K) = wsExcel.Cells(mXls1_Row_C, K)
    wsExcel.Cells(mXls1_Row_N, K).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_N, K).Font.Color = mColor_Z0

    wsExcel.Cells(mXls1_Row_D, K).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_D, K).Font.Color = mColor_Z0
Next K

ReDim arrDev_Cours(arrDev_Nb + 1)
For K = 1 To arrDev_Nb
    K2 = K + mXls1_Row_C
    wsExcel.Cells(K2, 1) = arrDev(K)
    wsExcel.Cells(K2, 1).Interior.Color = mColor_GB
    wsExcel.Cells(K2, 1).Font.Color = mColor_Z0
    wsExcel.Cells(K2, 10) = 50
    
    wsExcel.Cells(K + mXls1_Row_N, 1) = wsExcel.Cells(K2, 1)
    wsExcel.Cells(K + mXls1_Row_N, 1).Interior.Color = mColor_GB
    wsExcel.Cells(K + mXls1_Row_N, 1).Font.Color = mColor_Z0
    wsExcel.Cells(K + mXls1_Row_N, 10) = 50
    
    wsExcel.Cells(K + mXls1_Row_D, 1) = wsExcel.Cells(K2, 1)
    wsExcel.Cells(K + mXls1_Row_D, 1).Interior.Color = mColor_GB
    wsExcel.Cells(K + mXls1_Row_D, 1).Font.Color = mColor_Z0
    wsExcel.Cells(K + mXls1_Row_D, 10) = 0.15
    
    arrDev_RowT(K) = 0
    If arrDev(K) = "EUR" Then
        arrDev_Cours(K) = 1
        mXls1_Col_EUR = K2
    Else
        arrDev_Cours(K) = 0
        Call sqlYBIATAB0_Read("PDC", arrDev(K), wAmjMin, X)
        If IsNumeric(Mid$(X, 9, 15)) Then arrDev_Cours(K) = CDbl(Mid$(X, 9, 15) / 1000000000)
    End If
Next K
'__________________________________________________________________________________

Call cmdSelect_SQL_Xc_Dossier(2, "C")

Set wsExcel = wbExcel.Sheets(1)
For K = 1 To arrDev_Nb
    K2 = K + mXls1_Row_C
    wsExcel.Cells(K2, 7).Interior.Color = mColor_Y0
    wsExcel.Cells(K2, 8).Interior.Color = mColor_Y0
    wsExcel.Cells(K2, 11).Interior.Color = mColor_Y1

    If arrDev_RowT(K) > 0 Then
        wsExcel.Cells(K2, 3) = "=ECNF!G" & arrDev_RowT(K)
        'wsExcel.Cells(K2, 3) = "=ECNF!I" & arrDev_RowT(K)
        wsExcel.Cells(K2, 4) = "=ECNF!K" & arrDev_RowT(K)
        wsExcel.Cells(K2, 5) = "=ECNF!L" & arrDev_RowT(K)
        wsExcel.Cells(K2, 6) = "=ECNF!M" & arrDev_RowT(K)
        wsExcel.Cells(K2, 7) = "=ECNF!N" & arrDev_RowT(K)
        wsExcel.Cells(K2, 8) = "=ECNF!O" & arrDev_RowT(K)
        wsExcel.Cells(K2, 9) = "=ECNF!P" & arrDev_RowT(K)
        wsExcel.Cells(K2, 11).FormulaLocal = "= G" & K2 & " + H" & K2 & " * J" & K2 & " / 100"
        wsExcel.Cells(K2, 12) = "=ECNF!S" & arrDev_RowT(K)
        arrDev_RowT(K) = 0
    
    End If
Next K
'__________________________________________________________________________________

Call cmdSelect_SQL_Xc_Dossier(3, "N")

Set wsExcel = wbExcel.Sheets(1)
wsExcel.Cells(mXls1_Row_N, 1) = "ENOTIF"

For K = 1 To arrDev_Nb
    K2 = K + mXls1_Row_N
    wsExcel.Cells(K2, 7).Interior.Color = mColor_Y0
    wsExcel.Cells(K2, 8).Interior.Color = mColor_Y0
    wsExcel.Cells(K2, 11).Interior.Color = mColor_Y1
    If arrDev_RowT(K) > 0 Then
        wsExcel.Cells(K2, 3) = "=ENOTIF!G" & arrDev_RowT(K)
        'wsExcel.Cells(K2, 3) = "=ENOTIF!I" & arrDev_RowT(K)
        wsExcel.Cells(K2, 4) = "=ENOTIF!K" & arrDev_RowT(K)
        wsExcel.Cells(K2, 5) = "=ENOTIF!L" & arrDev_RowT(K)
        wsExcel.Cells(K2, 6) = "=ENOTIF!M" & arrDev_RowT(K)
        wsExcel.Cells(K2, 7) = "=ENOTIF!N" & arrDev_RowT(K)
        wsExcel.Cells(K2, 8) = "=ENOTIF!O" & arrDev_RowT(K)
        wsExcel.Cells(K2, 9) = "=ENOTIF!P" & arrDev_RowT(K)
        wsExcel.Cells(K2, 11).FormulaLocal = "= G" & K2 & " + H" & K2 & " * J" & K2 & " / 100"
        wsExcel.Cells(K2, 12) = "=ENOTIF!S" & arrDev_RowT(K)
        arrDev_RowT(K) = 0
    
    End If
Next K
'__________________________________________________________________________________


Call cmdSelect_SQL_Xc_ZSOLDE0(4, "('91120','91122')")

Set wsExcel = wbExcel.Sheets(1)
For K = 1 To arrDev_Nb
    K2 = K + mXls1_Row_C
    wsExcel.Cells(K2, 2).Interior.Color = mColor_G0
    wsExcel.Cells(K2, 3).Interior.Color = mColor_G0
    If arrDev_RowT(K) > 0 Then
        wsExcel.Cells(K2, 2) = "=91120!D" & arrDev_RowT(K)
        If Abs(wsExcel.Cells(K2, 2) + wsExcel.Cells(K2, 3)) > 0.01 Then
            wsExcel.Cells(K2, 2).Interior.Color = mColor_W0
            wsExcel.Cells(K2, 3).Interior.Color = mColor_W0
        End If
        
        arrDev_RowT(K) = 0
    
    End If
Next K
'__________________________________________________________________________________

Call cmdSelect_SQL_Xc_ZSOLDE0(5, "('98050')")

Set wsExcel = wbExcel.Sheets(1)
For K = 1 To arrDev_Nb
    K2 = K + mXls1_Row_N
    wsExcel.Cells(K2, 2).Interior.Color = mColor_G0
    wsExcel.Cells(K2, 3).Interior.Color = mColor_G0
    If arrDev_RowT(K) > 0 Then
        wsExcel.Cells(K2, 2) = "=98050!D" & arrDev_RowT(K)
        If Abs(wsExcel.Cells(K2, 2) + wsExcel.Cells(K2, 3)) > 0.01 Then
            wsExcel.Cells(K2, 2).Interior.Color = mColor_W0
            wsExcel.Cells(K2, 3).Interior.Color = mColor_W0
        End If
        arrDev_RowT(K) = 0
    
    End If
Next K
'__________________________________________________________________________________

Call cmdSelect_SQL_Xc_ZSOLDE0(6, "('91130','91131')")

Set wsExcel = wbExcel.Sheets(1)

wsExcel.Cells(mXls1_Row_D, 1) = "ELVD"
wsExcel.Cells(mXls1_Row_D, 2) = wsExcel.Cells(mXls1_Row_C, 2)
wsExcel.Cells(mXls1_Row_D, 10) = "% B "
wsExcel.Cells(mXls1_Row_D, 11) = wsExcel.Cells(mXls1_Row_C, 11)


For K = 1 To arrDev_Nb
    K2 = K + mXls1_Row_D
    wsExcel.Cells(K2, 2).Interior.Color = mColor_Y0
    wsExcel.Cells(K2, 11).Interior.Color = mColor_Y1
    If arrDev_RowT(K) > 0 Then
        wsExcel.Cells(K2, 2) = "=91130!D" & arrDev_RowT(K)
        wsExcel.Cells(K2, 11).FormulaLocal = "= - B" & K2 & " * J" & K2 & " / 100"
        arrDev_RowT(K) = 0
    
    End If
Next K
'__________________________________________________________________________________



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
    'If Not blnAuto Then MsgBox Error, vbCritical, Me.Name
    Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents
    
    wbExcel.SaveAs wFile
    wbExcel.Close
    appExcel.Quit

End Sub

Public Sub cmdSelect_SQL_XE1an()
On Error GoTo Error_Handler
Dim xSql As String
Dim X As String, XX As String, I As Integer
Dim K As Integer, K1 As Integer, K2 As Integer, mK2 As Integer, K4 As Integer, kLen As Integer
Dim mSolde As Currency, mDev As String, xDev As String, mDev_K As Integer
Dim blnOk As Boolean
Dim blnCALCS As Boolean

On Error GoTo Error_Handler
'===================================================================================
'______________________________________________'
X = paramServer("\\CDO_Archive\")
wAmjMin = YBIATAB0_DATE_CPT_J
wAmjMax = dateElp("AnAdd", 1, YBIATAB0_DATE_CPT_J)

blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

If X = "" Then X = "C:\Temp\"
If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"

wDIBM_Min = wAmjMin - 19000000

wDMS_Min = dateImp_Amj(wAmjMin)
wFile = X & Trim("CDO Engagements à 1 an " & wAmjMin & ".xlsx")
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "Crédits Documentaires : Engagements à 1 an  : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
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
    .Title = "DOSCDO"
    .Subject = "Engagements"
End With

'__________________________________________________________________________________

appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add

Set wsExcel = wbExcel.Sheets(1): wsExcel.Name = "Recapitulatif"
Set wsExcel = wbExcel.Sheets(2): wsExcel.Name = "91120"
Set wsExcel = wbExcel.Sheets(3): wsExcel.Name = "98050"
Set wsExcel = wbExcel.Sheets(4): wsExcel.Name = "91130"
Set wsExcel = wbExcel.Sheets(5): wsExcel.Name = "91131"
Set wsExcel = wbExcel.Sheets(6): wsExcel.Name = "91122"
Set wsExcel = wbExcel.Sheets(7): wsExcel.Name = "91132"
Set wsExcel = wbExcel.Sheets(8): wsExcel.Name = "98052"

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
    .Font.Name = "Calibri" '"Arial Unicode MS"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 80

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14CREDOC : Engagements à 1 an, arrêté au " & dateImp10(wAmjMin) _
                                 & vbCr

wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

mXls1_Col = arrDev_Nb + 2
mXls1_Row_T = 1
mXls1_Row_C = 1

wsExcel.Columns(1).ColumnWidth = 7: wsExcel.Cells(mXls1_Row_C, 1) = "PCI": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 7: wsExcel.Cells(mXls1_Row_C, 2) = "Devise": wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(3).ColumnWidth = 10: wsExcel.Cells(mXls1_Row_C, 3) = "Client": wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(4).ColumnWidth = 15: wsExcel.Cells(mXls1_Row_C, 4) = "Compte": wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(5).ColumnWidth = 30: wsExcel.Cells(mXls1_Row_C, 5) = "Intitulé": wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(6).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, 6) = "Balance " & dateImp10(wAmjMin): wsExcel.Columns(6).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(7).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, 7) = "Mt devise <= " & dateImp10(wAmjMax): wsExcel.Columns(7).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(8).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, 8) = "Mt devise > " & dateImp10(wAmjMax): wsExcel.Columns(8).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(9).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, 9) = "Total <= " & dateImp10(wAmjMax): wsExcel.Columns(9).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(10).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, 10) = "Total > " & dateImp10(wAmjMax): wsExcel.Columns(10).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(11).ColumnWidth = 7: wsExcel.Cells(mXls1_Row_C, 11) = "!!! ": wsExcel.Columns(11).NumberFormat = "[Yellow]### ### ### ##0"

wsExcel.Cells(mXls1_Row_C, 1).Interior.Color = mColor_Y1

For K = 1 To 11
    wsExcel.Cells(mXls1_Row_C, K).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_C, K).Font.Color = mColor_Z0
    
Next K

'__________________________________________________________________________________

xSql = "select count(*) as tally from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '91120%' and COMPTEFON <> '4'"
Set rsSab = cnsab.Execute(xSql)

arrYBIACPT0_Nb = rsSab(0)
ReDim arrYBIACPT0(arrYBIACPT0_Nb + 1), arrRow(arrYBIACPT0_Nb + 1), arrRow_Err(arrYBIACPT0_Nb + 1)

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '91120%' and COMPTEFON <> '4'" _
     & " order by COMPTEDEV , CLIENACLI , COMPTECOM"
Set rsSab = cnsab.Execute(xSql)
K = 0
Do While Not rsSab.EOF
    K = K + 1
    V = rsYBIACPT0_GetBuffer(rsSab, arrYBIACPT0(K))
    arrRow(K) = 0
    rsSab.MoveNext
Loop

Call cmdSelect_SQL_XE1an_Dossier(2, "91120")

K2 = mXls1_Row_C: mK2 = 0
Call rsYBIACPT0_Init(oldYBIACPT0)

Set wsExcel = wbExcel.Sheets(1)
For K = 1 To arrYBIACPT0_Nb
    If arrYBIACPT0(K).SOLDECEN = 0 And arrRow(K) = 0 Then
    Else
        
        xYBIACPT0 = arrYBIACPT0(K)
        
        If oldYBIACPT0.COMPTEDEV <> xYBIACPT0.COMPTEDEV Then
            If mK2 > 0 Then
                wsExcel.Cells(K2, 2).Font.Bold = True
                wsExcel.Cells(K2, 9).FormulaLocal = "=SOMME(G" & mK2 & ":G" & K2 & ")"
                wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
                wsExcel.Cells(K2, 9).Font.Bold = True

                wsExcel.Cells(K2, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & K2 & ")"
                wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
                wsExcel.Cells(K2, 10).Font.Bold = True
                
                wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)



            End If
            oldYBIACPT0 = xYBIACPT0
            mK2 = K2 + 1
        End If
        
        K2 = K2 + 1
        wsExcel.Cells(K2, 7).Interior.Color = mColor_Y0
        wsExcel.Cells(K2, 8).Interior.Color = mColor_Y1
        wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
        wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
        
        wsExcel.Cells(K2, 1) = xYBIACPT0.COMPTEOBL
        wsExcel.Cells(K2, 2) = xYBIACPT0.COMPTEDEV
        wsExcel.Cells(K2, 3) = xYBIACPT0.CLIENACLI
        wsExcel.Cells(K2, 4) = xYBIACPT0.COMPTECOM
        wsExcel.Cells(K2, 5) = xYBIACPT0.CLIENARA1
        wsExcel.Cells(K2, 6) = xYBIACPT0.SOLDECEN
        If arrRow(K) > 0 Then
            wsExcel.Cells(K2, 7).FormulaLocal = "=91120!J" & arrRow(K)
            wsExcel.Cells(K2, 8).FormulaLocal = "=91120!K" & arrRow(K)
        End If
        If xYBIACPT0.SOLDECEN <> wsExcel.Cells(K2, 7) + wsExcel.Cells(K2, 8) Then
            For K4 = 1 To 6
                wsExcel.Cells(K2, K4).Interior.Color = mColor_W1
            Next K4
        End If
        If arrRow_Err(K) > 0 Then
            wsExcel.Cells(K2, 11) = arrRow_Err(K)
            wsExcel.Cells(K2, 11).Interior.Color = vbRed
        End If
    End If
Next K

If mK2 > 0 Then
    wsExcel.Cells(K2, 2).Font.Bold = True
    wsExcel.Cells(K2, 9).FormulaLocal = "=SOMME(G" & mK2 & ":G" & K2 & ")"
    wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
    wsExcel.Cells(K2, 9).Font.Bold = True
    
    wsExcel.Cells(K2, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & K2 & ")"
    wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
    wsExcel.Cells(K2, 10).Font.Bold = True
    
    wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Weight = xlThick
    wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
End If

'=========================================================================================

mXls1_Row_N = K2 + 2
wsExcel.Cells(mXls1_Row_N, 1).Interior.Color = mColor_Y1
For K = 1 To 11
    
    wsExcel.Cells(mXls1_Row_N, K) = wsExcel.Cells(mXls1_Row_C, K)
    wsExcel.Cells(mXls1_Row_N, K).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_N, K).Font.Color = mColor_Z0

Next K

xSql = "select count(*) as tally from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '98050%' and COMPTEFON <> '4'"
Set rsSab = cnsab.Execute(xSql)

arrYBIACPT0_Nb = rsSab(0)
ReDim arrYBIACPT0(arrYBIACPT0_Nb + 1), arrRow(arrYBIACPT0_Nb + 1), arrRow_Err(arrYBIACPT0_Nb + 1)

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '98050%' and COMPTEFON <> '4'" _
     & " order by COMPTEDEV , CLIENACLI , COMPTECOM"
Set rsSab = cnsab.Execute(xSql)
K = 0
Do While Not rsSab.EOF
    K = K + 1
    V = rsYBIACPT0_GetBuffer(rsSab, arrYBIACPT0(K))
    rsSab.MoveNext
Loop

Call cmdSelect_SQL_XE1an_Dossier(3, "98050")

Set wsExcel = wbExcel.Sheets(1)


K2 = mXls1_Row_N: mK2 = 0
Call rsYBIACPT0_Init(oldYBIACPT0)

Set wsExcel = wbExcel.Sheets(1)
For K = 1 To arrYBIACPT0_Nb
    If arrYBIACPT0(K).SOLDECEN = 0 And arrRow(K) = 0 Then
    Else
        
        xYBIACPT0 = arrYBIACPT0(K)
        
        If oldYBIACPT0.COMPTEDEV <> xYBIACPT0.COMPTEDEV Then
            If mK2 > 0 Then
                wsExcel.Cells(K2, 2).Font.Bold = True
                wsExcel.Cells(K2, 9).FormulaLocal = "=SOMME(G" & mK2 & ":G" & K2 & ")"
                wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
                wsExcel.Cells(K2, 9).Font.Bold = True

                wsExcel.Cells(K2, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & K2 & ")"
                wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
                wsExcel.Cells(K2, 10).Font.Bold = True
                
                wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)

            End If
            oldYBIACPT0 = xYBIACPT0
            mK2 = K2 + 1
        End If
        
        K2 = K2 + 1
        wsExcel.Cells(K2, 7).Interior.Color = mColor_Y0
        wsExcel.Cells(K2, 8).Interior.Color = mColor_Y1
        wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
        wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
        
        wsExcel.Cells(K2, 1) = xYBIACPT0.COMPTEOBL
        wsExcel.Cells(K2, 2) = xYBIACPT0.COMPTEDEV
        wsExcel.Cells(K2, 3) = xYBIACPT0.CLIENACLI
        wsExcel.Cells(K2, 4) = xYBIACPT0.COMPTECOM
        wsExcel.Cells(K2, 5) = xYBIACPT0.CLIENARA1
        wsExcel.Cells(K2, 6) = xYBIACPT0.SOLDECEN
        If arrRow(K) > 0 Then
            wsExcel.Cells(K2, 7).FormulaLocal = "=98050!J" & arrRow(K)
            wsExcel.Cells(K2, 8).FormulaLocal = "=98050!K" & arrRow(K)
        End If
        If xYBIACPT0.SOLDECEN <> wsExcel.Cells(K2, 7) + wsExcel.Cells(K2, 8) Then
            For K4 = 1 To 6
                wsExcel.Cells(K2, K4).Interior.Color = mColor_W1
            Next K4
        End If

        If arrRow_Err(K) > 0 Then
            wsExcel.Cells(K2, 11) = arrRow_Err(K)
            wsExcel.Cells(K2, 11).Interior.Color = vbRed
        End If

    End If
Next K

If mK2 > 0 Then
    wsExcel.Cells(K2, 2).Font.Bold = True
    wsExcel.Cells(K2, 9).FormulaLocal = "=SOMME(G" & mK2 & ":G" & K2 & ")"
    wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
    wsExcel.Cells(K2, 9).Font.Bold = True
    
    wsExcel.Cells(K2, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & K2 & ")"
    wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
    wsExcel.Cells(K2, 10).Font.Bold = True
    
    wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Weight = xlThick
    wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
End If
'=========================================================================================

mXls1_Row_T = K2 + 2
wsExcel.Cells(mXls1_Row_T, 1).Interior.Color = mColor_Y1
For K = 1 To 11
    
    wsExcel.Cells(mXls1_Row_T, K) = wsExcel.Cells(mXls1_Row_C, K)
    wsExcel.Cells(mXls1_Row_T, K).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_T, K).Font.Color = mColor_Z0

Next K

xSql = "select count(*) as tally from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '91130%' and COMPTEFON <> '4'"
Set rsSab = cnsab.Execute(xSql)

arrYBIACPT0_Nb = rsSab(0)
ReDim arrYBIACPT0(arrYBIACPT0_Nb + 1), arrRow(arrYBIACPT0_Nb + 1), arrRow_Err(arrYBIACPT0_Nb + 1)

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '91130%' and COMPTEFON <> '4'" _
     & " order by COMPTEDEV , CLIENACLI , COMPTECOM"
Set rsSab = cnsab.Execute(xSql)
K = 0
Do While Not rsSab.EOF
    K = K + 1
    V = rsYBIACPT0_GetBuffer(rsSab, arrYBIACPT0(K))
    rsSab.MoveNext
Loop

Call cmdSelect_SQL_XE1an_Dossier(4, "91130")

Set wsExcel = wbExcel.Sheets(1)


K2 = mXls1_Row_T: mK2 = 0
Call rsYBIACPT0_Init(oldYBIACPT0)

Set wsExcel = wbExcel.Sheets(1)
For K = 1 To arrYBIACPT0_Nb
    If arrYBIACPT0(K).SOLDECEN = 0 And arrRow(K) = 0 Then
    Else
        
        xYBIACPT0 = arrYBIACPT0(K)
        
        If oldYBIACPT0.COMPTEDEV <> xYBIACPT0.COMPTEDEV Then
            If mK2 > 0 Then
                wsExcel.Cells(K2, 2).Font.Bold = True
                wsExcel.Cells(K2, 9).FormulaLocal = "=SOMME(G" & mK2 & ":G" & K2 & ")"
                wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
                wsExcel.Cells(K2, 9).Font.Bold = True

                wsExcel.Cells(K2, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & K2 & ")"
                wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
                wsExcel.Cells(K2, 10).Font.Bold = True
                
                wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)

            End If
            oldYBIACPT0 = xYBIACPT0
            mK2 = K2 + 1
        End If
        
        K2 = K2 + 1
        wsExcel.Cells(K2, 7).Interior.Color = mColor_Y0
        wsExcel.Cells(K2, 8).Interior.Color = mColor_Y1
        wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
        wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
        
        wsExcel.Cells(K2, 1) = xYBIACPT0.COMPTEOBL
        wsExcel.Cells(K2, 2) = xYBIACPT0.COMPTEDEV
        wsExcel.Cells(K2, 3) = xYBIACPT0.CLIENACLI
        wsExcel.Cells(K2, 4) = xYBIACPT0.COMPTECOM
        wsExcel.Cells(K2, 5) = xYBIACPT0.CLIENARA1
        wsExcel.Cells(K2, 6) = xYBIACPT0.SOLDECEN
        If arrRow(K) > 0 Then
            wsExcel.Cells(K2, 7).FormulaLocal = "=91130!J" & arrRow(K)
            wsExcel.Cells(K2, 8).FormulaLocal = "=91130!K" & arrRow(K)
        End If
        If xYBIACPT0.SOLDECEN <> wsExcel.Cells(K2, 7) + wsExcel.Cells(K2, 8) Then
            For K4 = 1 To 6
                wsExcel.Cells(K2, K4).Interior.Color = mColor_W1
            Next K4
        End If

        If arrRow_Err(K) > 0 Then
            wsExcel.Cells(K2, 11) = arrRow_Err(K)
            wsExcel.Cells(K2, 11).Interior.Color = vbRed
        End If

    End If
Next K

If mK2 > 0 Then
    wsExcel.Cells(K2, 2).Font.Bold = True
    wsExcel.Cells(K2, 9).FormulaLocal = "=SOMME(G" & mK2 & ":G" & K2 & ")"
    wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
    wsExcel.Cells(K2, 9).Font.Bold = True
    
    wsExcel.Cells(K2, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & K2 & ")"
    wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
    wsExcel.Cells(K2, 10).Font.Bold = True
    
    wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Weight = xlThick
    wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
End If

'============================================================================================

mXls1_Row_T = K2 + 2
wsExcel.Cells(mXls1_Row_T, 1).Interior.Color = mColor_Y1
For K = 1 To 11
    
    wsExcel.Cells(mXls1_Row_T, K) = wsExcel.Cells(mXls1_Row_C, K)
    wsExcel.Cells(mXls1_Row_T, K).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_T, K).Font.Color = mColor_Z0

Next K

xSql = "select count(*) as tally from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '91131%' and COMPTEFON <> '4'"
Set rsSab = cnsab.Execute(xSql)

arrYBIACPT0_Nb = rsSab(0)
ReDim arrYBIACPT0(arrYBIACPT0_Nb + 1), arrRow(arrYBIACPT0_Nb + 1), arrRow_Err(arrYBIACPT0_Nb + 1)

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '91131%' and COMPTEFON <> '4'" _
     & " order by COMPTEDEV , CLIENACLI , COMPTECOM"
Set rsSab = cnsab.Execute(xSql)
K = 0
Do While Not rsSab.EOF
    K = K + 1
    V = rsYBIACPT0_GetBuffer(rsSab, arrYBIACPT0(K))
    rsSab.MoveNext
Loop

Call cmdSelect_SQL_XE1an_Dossier(5, "91131")

Set wsExcel = wbExcel.Sheets(1)


K2 = mXls1_Row_T: mK2 = 0
Call rsYBIACPT0_Init(oldYBIACPT0)

Set wsExcel = wbExcel.Sheets(1)
For K = 1 To arrYBIACPT0_Nb
    If arrYBIACPT0(K).SOLDECEN = 0 And arrRow(K) = 0 Then
    Else
        
        xYBIACPT0 = arrYBIACPT0(K)
        
        If oldYBIACPT0.COMPTEDEV <> xYBIACPT0.COMPTEDEV Then
            If mK2 > 0 Then
                wsExcel.Cells(K2, 2).Font.Bold = True
                wsExcel.Cells(K2, 9).FormulaLocal = "=SOMME(G" & mK2 & ":G" & K2 & ")"
                wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
                wsExcel.Cells(K2, 9).Font.Bold = True

                wsExcel.Cells(K2, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & K2 & ")"
                wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
                wsExcel.Cells(K2, 10).Font.Bold = True
                
                wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)

            End If
            oldYBIACPT0 = xYBIACPT0
            mK2 = K2 + 1
        End If
        
        K2 = K2 + 1
        wsExcel.Cells(K2, 7).Interior.Color = mColor_Y0
        wsExcel.Cells(K2, 8).Interior.Color = mColor_Y1
        wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
        wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
        
        wsExcel.Cells(K2, 1) = xYBIACPT0.COMPTEOBL
        wsExcel.Cells(K2, 2) = xYBIACPT0.COMPTEDEV
        wsExcel.Cells(K2, 3) = xYBIACPT0.CLIENACLI
        wsExcel.Cells(K2, 4) = xYBIACPT0.COMPTECOM
        wsExcel.Cells(K2, 5) = xYBIACPT0.CLIENARA1
        wsExcel.Cells(K2, 6) = xYBIACPT0.SOLDECEN
        If arrRow(K) > 0 Then
            wsExcel.Cells(K2, 7).FormulaLocal = "=91131!J" & arrRow(K)
            wsExcel.Cells(K2, 8).FormulaLocal = "=91131!K" & arrRow(K)
        End If
        If xYBIACPT0.SOLDECEN <> wsExcel.Cells(K2, 7) + wsExcel.Cells(K2, 8) Then
            For K4 = 1 To 6
                wsExcel.Cells(K2, K4).Interior.Color = mColor_W1
            Next K4
        End If

        If arrRow_Err(K) > 0 Then
            wsExcel.Cells(K2, 11) = arrRow_Err(K)
            wsExcel.Cells(K2, 11).Interior.Color = vbRed
        End If

    End If
Next K

If mK2 > 0 Then
    wsExcel.Cells(K2, 2).Font.Bold = True
    wsExcel.Cells(K2, 9).FormulaLocal = "=SOMME(G" & mK2 & ":G" & K2 & ")"
    wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
    wsExcel.Cells(K2, 9).Font.Bold = True
    
    wsExcel.Cells(K2, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & K2 & ")"
    wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
    wsExcel.Cells(K2, 10).Font.Bold = True
    
    wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Weight = xlThick
    wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
End If

'=========================================================================================

mXls1_Row_T = K2 + 2
wsExcel.Cells(mXls1_Row_T, 1).Interior.Color = mColor_Y1
For K = 1 To 11
    
    wsExcel.Cells(mXls1_Row_T, K) = wsExcel.Cells(mXls1_Row_C, K)
    wsExcel.Cells(mXls1_Row_T, K).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_T, K).Font.Color = mColor_Z0

Next K

xSql = "select count(*) as tally from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '91122%' and COMPTEFON <> '4'"
Set rsSab = cnsab.Execute(xSql)

arrYBIACPT0_Nb = rsSab(0)
ReDim arrYBIACPT0(arrYBIACPT0_Nb + 1), arrRow(arrYBIACPT0_Nb + 1), arrRow_Err(arrYBIACPT0_Nb + 1)

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '91122%' and COMPTEFON <> '4'" _
     & " order by COMPTEDEV , CLIENACLI , COMPTECOM"
Set rsSab = cnsab.Execute(xSql)
K = 0
Do While Not rsSab.EOF
    K = K + 1
    V = rsYBIACPT0_GetBuffer(rsSab, arrYBIACPT0(K))
    rsSab.MoveNext
Loop

Call cmdSelect_SQL_XE1an_Dossier(6, "91122")

Set wsExcel = wbExcel.Sheets(1)


K2 = mXls1_Row_T: mK2 = 0
Call rsYBIACPT0_Init(oldYBIACPT0)

Set wsExcel = wbExcel.Sheets(1)
For K = 1 To arrYBIACPT0_Nb
    If arrYBIACPT0(K).SOLDECEN = 0 And arrRow(K) = 0 Then
    Else
        
        xYBIACPT0 = arrYBIACPT0(K)
        
        If oldYBIACPT0.COMPTEDEV <> xYBIACPT0.COMPTEDEV Then
            If mK2 > 0 Then
                wsExcel.Cells(K2, 2).Font.Bold = True
                wsExcel.Cells(K2, 9).FormulaLocal = "=SOMME(G" & mK2 & ":G" & K2 & ")"
                wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
                wsExcel.Cells(K2, 9).Font.Bold = True

                wsExcel.Cells(K2, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & K2 & ")"
                wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
                wsExcel.Cells(K2, 10).Font.Bold = True
                
                wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)

            End If
            oldYBIACPT0 = xYBIACPT0
            mK2 = K2 + 1
        End If
        
        K2 = K2 + 1
        wsExcel.Cells(K2, 7).Interior.Color = mColor_Y0
        wsExcel.Cells(K2, 8).Interior.Color = mColor_Y1
        wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
        wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
        
        wsExcel.Cells(K2, 1) = xYBIACPT0.COMPTEOBL
        wsExcel.Cells(K2, 2) = xYBIACPT0.COMPTEDEV
        wsExcel.Cells(K2, 3) = xYBIACPT0.CLIENACLI
        wsExcel.Cells(K2, 4) = xYBIACPT0.COMPTECOM
        wsExcel.Cells(K2, 5) = xYBIACPT0.CLIENARA1
        wsExcel.Cells(K2, 6) = xYBIACPT0.SOLDECEN
        If arrRow(K) > 0 Then
            wsExcel.Cells(K2, 7).FormulaLocal = "=91122!J" & arrRow(K)
            wsExcel.Cells(K2, 8).FormulaLocal = "=91122!K" & arrRow(K)
        End If
        If xYBIACPT0.SOLDECEN <> wsExcel.Cells(K2, 7) + wsExcel.Cells(K2, 8) Then
            For K4 = 1 To 6
                wsExcel.Cells(K2, K4).Interior.Color = mColor_W1
            Next K4
        End If

        If arrRow_Err(K) > 0 Then
            wsExcel.Cells(K2, 11) = arrRow_Err(K)
            wsExcel.Cells(K2, 11).Interior.Color = vbRed
        End If

    End If
Next K

If mK2 > 0 Then
    wsExcel.Cells(K2, 2).Font.Bold = True
    wsExcel.Cells(K2, 9).FormulaLocal = "=SOMME(G" & mK2 & ":G" & K2 & ")"
    wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
    wsExcel.Cells(K2, 9).Font.Bold = True
    
    wsExcel.Cells(K2, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & K2 & ")"
    wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
    wsExcel.Cells(K2, 10).Font.Bold = True
    
    wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Weight = xlThick
    wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
End If

'=========================================================================================

mXls1_Row_T = K2 + 2
wsExcel.Cells(mXls1_Row_T, 1).Interior.Color = mColor_Y1
For K = 1 To 11
    
    wsExcel.Cells(mXls1_Row_T, K) = wsExcel.Cells(mXls1_Row_C, K)
    wsExcel.Cells(mXls1_Row_T, K).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_T, K).Font.Color = mColor_Z0

Next K

xSql = "select count(*) as tally from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '91132%' and COMPTEFON <> '4'"
Set rsSab = cnsab.Execute(xSql)

arrYBIACPT0_Nb = rsSab(0)
ReDim arrYBIACPT0(arrYBIACPT0_Nb + 1), arrRow(arrYBIACPT0_Nb + 1), arrRow_Err(arrYBIACPT0_Nb + 1)

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '91132%' and COMPTEFON <> '4'" _
     & " order by COMPTEDEV , CLIENACLI , COMPTECOM"
Set rsSab = cnsab.Execute(xSql)
K = 0
Do While Not rsSab.EOF
    K = K + 1
    V = rsYBIACPT0_GetBuffer(rsSab, arrYBIACPT0(K))
    rsSab.MoveNext
Loop

Call cmdSelect_SQL_XE1an_Dossier(7, "91132")

Set wsExcel = wbExcel.Sheets(1)


K2 = mXls1_Row_T: mK2 = 0
Call rsYBIACPT0_Init(oldYBIACPT0)

Set wsExcel = wbExcel.Sheets(1)
For K = 1 To arrYBIACPT0_Nb
    If arrYBIACPT0(K).SOLDECEN = 0 And arrRow(K) = 0 Then
    Else
        
        xYBIACPT0 = arrYBIACPT0(K)
        
        If oldYBIACPT0.COMPTEDEV <> xYBIACPT0.COMPTEDEV Then
            If mK2 > 0 Then
                wsExcel.Cells(K2, 2).Font.Bold = True
                wsExcel.Cells(K2, 9).FormulaLocal = "=SOMME(G" & mK2 & ":G" & K2 & ")"
                wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
                wsExcel.Cells(K2, 9).Font.Bold = True

                wsExcel.Cells(K2, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & K2 & ")"
                wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
                wsExcel.Cells(K2, 10).Font.Bold = True
                
                wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)

            End If
            oldYBIACPT0 = xYBIACPT0
            mK2 = K2 + 1
        End If
        
        K2 = K2 + 1
        wsExcel.Cells(K2, 7).Interior.Color = mColor_Y0
        wsExcel.Cells(K2, 8).Interior.Color = mColor_Y1
        wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
        wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
        
        wsExcel.Cells(K2, 1) = xYBIACPT0.COMPTEOBL
        wsExcel.Cells(K2, 2) = xYBIACPT0.COMPTEDEV
        wsExcel.Cells(K2, 3) = xYBIACPT0.CLIENACLI
        wsExcel.Cells(K2, 4) = xYBIACPT0.COMPTECOM
        wsExcel.Cells(K2, 5) = xYBIACPT0.CLIENARA1
        wsExcel.Cells(K2, 6) = xYBIACPT0.SOLDECEN
        If arrRow(K) > 0 Then
            wsExcel.Cells(K2, 7).FormulaLocal = "=91132!J" & arrRow(K)
            wsExcel.Cells(K2, 8).FormulaLocal = "=91132!K" & arrRow(K)
        End If
        If xYBIACPT0.SOLDECEN <> wsExcel.Cells(K2, 7) + wsExcel.Cells(K2, 8) Then
            For K4 = 1 To 6
                wsExcel.Cells(K2, K4).Interior.Color = mColor_W1
            Next K4
        End If

        If arrRow_Err(K) > 0 Then
            wsExcel.Cells(K2, 11) = arrRow_Err(K)
            wsExcel.Cells(K2, 11).Interior.Color = vbRed
        End If

    End If
Next K

If mK2 > 0 Then
    wsExcel.Cells(K2, 2).Font.Bold = True
    wsExcel.Cells(K2, 9).FormulaLocal = "=SOMME(G" & mK2 & ":G" & K2 & ")"
    wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
    wsExcel.Cells(K2, 9).Font.Bold = True
    
    wsExcel.Cells(K2, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & K2 & ")"
    wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
    wsExcel.Cells(K2, 10).Font.Bold = True
    
    wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Weight = xlThick
    wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
End If
'=========================================================================================

mXls1_Row_T = K2 + 2
wsExcel.Cells(mXls1_Row_T, 1).Interior.Color = mColor_Y1
For K = 1 To 11
    
    wsExcel.Cells(mXls1_Row_T, K) = wsExcel.Cells(mXls1_Row_C, K)
    wsExcel.Cells(mXls1_Row_T, K).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_T, K).Font.Color = mColor_Z0

Next K

xSql = "select count(*) as tally from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '98052%' and COMPTEFON <> '4'"
Set rsSab = cnsab.Execute(xSql)

arrYBIACPT0_Nb = rsSab(0)
ReDim arrYBIACPT0(arrYBIACPT0_Nb + 1), arrRow(arrYBIACPT0_Nb + 1), arrRow_Err(arrYBIACPT0_Nb + 1)

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTEOBL like '98052%' and COMPTEFON <> '4'" _
     & " order by COMPTEDEV , CLIENACLI , COMPTECOM"
Set rsSab = cnsab.Execute(xSql)
K = 0
Do While Not rsSab.EOF
    K = K + 1
    V = rsYBIACPT0_GetBuffer(rsSab, arrYBIACPT0(K))
    rsSab.MoveNext
Loop

Call cmdSelect_SQL_XE1an_Dossier(8, "98052")

Set wsExcel = wbExcel.Sheets(1)


K2 = mXls1_Row_T: mK2 = 0
Call rsYBIACPT0_Init(oldYBIACPT0)

Set wsExcel = wbExcel.Sheets(1)
For K = 1 To arrYBIACPT0_Nb
    If arrYBIACPT0(K).SOLDECEN = 0 And arrRow(K) = 0 Then
    Else
        
        xYBIACPT0 = arrYBIACPT0(K)
        
        If oldYBIACPT0.COMPTEDEV <> xYBIACPT0.COMPTEDEV Then
            If mK2 > 0 Then
                wsExcel.Cells(K2, 2).Font.Bold = True
                wsExcel.Cells(K2, 9).FormulaLocal = "=SOMME(G" & mK2 & ":G" & K2 & ")"
                wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
                wsExcel.Cells(K2, 9).Font.Bold = True

                wsExcel.Cells(K2, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & K2 & ")"
                wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
                wsExcel.Cells(K2, 10).Font.Bold = True
                
                wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)

            End If
            oldYBIACPT0 = xYBIACPT0
            mK2 = K2 + 1
        End If
        
        K2 = K2 + 1
        wsExcel.Cells(K2, 7).Interior.Color = mColor_Y0
        wsExcel.Cells(K2, 8).Interior.Color = mColor_Y1
        wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
        wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
        
        wsExcel.Cells(K2, 1) = xYBIACPT0.COMPTEOBL
        wsExcel.Cells(K2, 2) = xYBIACPT0.COMPTEDEV
        wsExcel.Cells(K2, 3) = xYBIACPT0.CLIENACLI
        wsExcel.Cells(K2, 4) = xYBIACPT0.COMPTECOM
        wsExcel.Cells(K2, 5) = xYBIACPT0.CLIENARA1
        wsExcel.Cells(K2, 6) = xYBIACPT0.SOLDECEN
        If arrRow(K) > 0 Then
            wsExcel.Cells(K2, 7).FormulaLocal = "=98052!J" & arrRow(K)
            wsExcel.Cells(K2, 8).FormulaLocal = "=98052!K" & arrRow(K)
        End If
        If xYBIACPT0.SOLDECEN <> wsExcel.Cells(K2, 7) + wsExcel.Cells(K2, 8) Then
            For K4 = 1 To 6
                wsExcel.Cells(K2, K4).Interior.Color = mColor_W1
            Next K4
        End If

        If arrRow_Err(K) > 0 Then
            wsExcel.Cells(K2, 11) = arrRow_Err(K)
            wsExcel.Cells(K2, 11).Interior.Color = vbRed
        End If

    End If
Next K

If mK2 > 0 Then
    wsExcel.Cells(K2, 2).Font.Bold = True
    wsExcel.Cells(K2, 9).FormulaLocal = "=SOMME(G" & mK2 & ":G" & K2 & ")"
    wsExcel.Cells(K2, 9).Interior.Color = mColor_Y0
    wsExcel.Cells(K2, 9).Font.Bold = True
    
    wsExcel.Cells(K2, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & K2 & ")"
    wsExcel.Cells(K2, 10).Interior.Color = mColor_Y1
    wsExcel.Cells(K2, 10).Font.Bold = True
    
    wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Weight = xlThick
    wsExcel.Range("A" & K2 & ":J" & K2).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
End If

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

Public Sub cmdSelect_SQL_Xi()
'On Error GoTo Error_Handler
Dim wFilex As String, wFile As String, xSql As String
Dim X As String, XX As String, I As Integer
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, kLen As Integer
Dim mSolde As Currency, mDev As String, xDev As String, mDev_K As Integer
Dim iRow As Integer, iSheet As Integer
Dim derniere_ligne As Long
Dim totalcvEuro As Currency
Dim totalSup93 As Currency
Dim totalInf93 As Currency
Dim iName() As Long
Dim m1mois As Long
Dim m1mois_IBM As Long
Dim Xd As Double

'______________________________________________

wAmjMin = YBIATAB0_DATE_CPT_J
wDIBM_Min = wAmjMin - 19000000
wDMS_Min = dateImp_Amj(wAmjMin)
m1mois = CLng(DateValue(Now)) - 30
m1mois_IBM = Format(m1mois, "yyyymmdd")
m1mois_IBM = m1mois_IBM - 19000000

wFile = Trim("C:\Temp\CDO engagements Intragroupe au " & dateImp_Amj(wAmjMin) & ".xlsx")
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut :" _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "Crédits Documentaires : engagements Intragroupe  : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If
'_________________________________________


If Dir(wFile) <> "" Then Kill wFile

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'06-03-2012 liste des racines à exclure
mRacinesExclues = ""

'If Not blnAuto Then
'$JPL    mRacinesExclues = "11001 11004 11005 11006 11008 11011 11012 11067 11069 11072 11084 11087 50699 50776"
'$JPL 20130711    mRacinesExclues = "11004 11005 11006 11008 11011 11012 11067 11069 11072 11084 11087 50699 50776"
'$JPL 20130711    mRacinesExclues = InputBox("par défaut :  VERSION 17-06-2013 " _
'$JPL 20130711        & vbCrLf & "     =========================" & vbCrLf & mRacinesExclues _
'$JPL 20130711        & vbCrLf & "     =========================", "Liste des racines à exclure  : ", mRacinesExclues)
'End If
'=========================================================================================

'$JPL 20130711If Trim(mRacinesExclues) = "" Then Exit Sub

mCDODOSOUV_11001 = 1130617
mCDODOSOUV_11012 = 1130718
LFB_RacinesExclues = "0011004 0011005 0011006 0011008 0011011 0011012 0011067 0011069 0011072 0011084 0011087 0050699 0050776"

'$JPL 2014-07-30
mCDODOSOUV_11001 = 0
mCDODOSOUV_11012 = 0
LFB_RacinesExclues = ""

mHeader_xls = "Récapitulatif des engagements INTRAGROUPE"
If mRacinesExclues <> "" Then mHeader_xls = "V 17-06-2013 :" & mHeader_xls & " (" & mRacinesExclues & " exclus)"
xSql = "select count(*) as tally from " & paramIBM_Library_SAB & ".ZREPCOR0 " _
     & " where REPCORETB = 1 and REPCORATR = 'INTRAGROUP' and REPCORVST = 'G'"
Set rsSab = cnsab.Execute(xSql)

arrYBIACPT0_Nb = rsSab(0)
ReDim arrYBIACPT0(arrYBIACPT0_Nb + 1)

xSql = "select distinct(REPCORVCL) from " & paramIBM_Library_SAB & ".ZREPCOR0 " _
     & " where REPCORETB = 1 and REPCORATR = 'INTRAGROUP' and REPCORVST = 'G'" _
     & " order by REPCORVCL"
Set rsSab = cnsab.Execute(xSql)
arrYBIACPT0_Nb = 0
Do While Not rsSab.EOF
    X = Mid$(rsSab("REPCORVCL"), 1, 7)
    XX = Mid$(X, 3, 5)
    If InStr(mRacinesExclues, XX) = 0 Then
        arrYBIACPT0_Nb = arrYBIACPT0_Nb + 1
        arrYBIACPT0(arrYBIACPT0_Nb).CLIENACLI = X
    End If
    rsSab.MoveNext
Loop

''!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'GoTo Suite1   '06-03-2012 sans objet
''!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
'X = MsgBox("Uniquement participation > 10% ? ", vbQuestion + vbYesNo, "Engagements intragroupe")
'If X = vbYes Then
'    mHeader_xls = "Récapitulatif des engagements Groupe ( participation > 10 %)"
'
'    ReDim arrYBIACPT0(24)
'    arrYBIACPT0_Nb = 23
'    arrYBIACPT0(1).CLIENACLI = "0011001"
'    arrYBIACPT0(2).CLIENACLI = "0011002"
'    arrYBIACPT0(3).CLIENACLI = "0011003"
'    arrYBIACPT0(4).CLIENACLI = "0011004"
'    arrYBIACPT0(5).CLIENACLI = "0011005"
'    arrYBIACPT0(6).CLIENACLI = "0011006"
'    arrYBIACPT0(7).CLIENACLI = "0011007"
'    arrYBIACPT0(8).CLIENACLI = "0011008"
'    arrYBIACPT0(9).CLIENACLI = "0011009"
'    arrYBIACPT0(10).CLIENACLI = "0011010"
'    arrYBIACPT0(11).CLIENACLI = "0011011"
'    arrYBIACPT0(12).CLIENACLI = "0011012"
'    arrYBIACPT0(13).CLIENACLI = "0011067"
'    arrYBIACPT0(14).CLIENACLI = "0011069"
'    arrYBIACPT0(15).CLIENACLI = "0011072"
'    arrYBIACPT0(16).CLIENACLI = "0011080"
'    arrYBIACPT0(17).CLIENACLI = "0011087"
'    arrYBIACPT0(18).CLIENACLI = "0011106"
'    arrYBIACPT0(19).CLIENACLI = "0011466"
'    arrYBIACPT0(20).CLIENACLI = "0011474"
'    arrYBIACPT0(21).CLIENACLI = "0011477"
'    arrYBIACPT0(22).CLIENACLI = "0050601"
'    arrYBIACPT0(23).CLIENACLI = "0050776"
'Else
'    X = MsgBox("Uniquement racines 'Embargo Libye' ? ", vbQuestion + vbYesNo, "Engagements intragroupe")
'    If X = vbYes Then
'        mHeader_xls = "Récapitulatif des engagements Groupe ( Embargo Libye)"
'        ReDim arrYBIACPT0(9)
'        arrYBIACPT0_Nb = 9
'        arrYBIACPT0(1).CLIENACLI = "0011012"
'        arrYBIACPT0(2).CLIENACLI = "0011084"
'        arrYBIACPT0(3).CLIENACLI = "0011085"
'        arrYBIACPT0(4).CLIENACLI = "0011088"
'        arrYBIACPT0(5).CLIENACLI = "0011220"
'        arrYBIACPT0(6).CLIENACLI = "0011425"
'        arrYBIACPT0(7).CLIENACLI = "0011540"
'        arrYBIACPT0(8).CLIENACLI = "0050733"
'        arrYBIACPT0(9).CLIENACLI = "0050775"
'    End If
'
'End If
''!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
''!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call lstErr_ChangeLastItem(lstErr, cmdContext, "Intitulé racine : "): DoEvents
For K = 1 To arrYBIACPT0_Nb
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "Intitulé racine : " & arrYBIACPT0(K).CLIENACLI): DoEvents
    xSql = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
         & " where CLIENACLI = '" & arrYBIACPT0(K).CLIENACLI & "'"
    Set rsSab = cnsab.Execute(xSql)
    If Not rsSab.EOF Then
        arrYBIACPT0(K).CLIENARA1 = rsSab("CLIENARA1")
    Else
        arrYBIACPT0(K).CLIENARA1 = "?"
    End If
Next K
'_______________________________________________________________________________________


mXls1_Col = arrDev_Nb + 2
mXls1_Row_C = 1
mXls1_Row_N = arrYBIACPT0_Nb + 3
mXls1_Row_SP = mXls1_Row_N + arrYBIACPT0_Nb + 2
mXls1_Row_T = mXls1_Row_SP + arrYBIACPT0_Nb + 2

ReDim arrDev_Cours(arrDev_Nb + 1)

For K = 1 To arrDev_Nb
    K2 = K + 2
    arrDev_RowT(K) = 0
    If arrDev(K) = "EUR" Then
        arrDev_Cours(K) = 1
        mXls1_Col_EUR = K2
    Else
        arrDev_Cours(K) = 0.00001
        Call sqlYBIATAB0_Read("PDC", arrDev(K), wAmjMin, X)
        If IsNumeric(Mid$(X, 9, 15)) Then arrDev_Cours(K) = CDbl(Mid$(X, 9, 15) / 1000000000)
    End If
Next K
'__________________________________________________________________________________
Set appExcel = CreateObject("Excel.Application")
'appExcel.Visible = True
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "DOSCDO"
    .Subject = "eng actionnaires"
End With

appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add

Set wsExcel = wbExcel.Sheets(1): wsExcel.Name = "CDE " & wDMS_Min
Set wsExcel = wbExcel.Sheets(2): wsExcel.Name = "CDE >= 93 J"
Set wsExcel = wbExcel.Sheets(3): wsExcel.Name = "CDE < 93 J"
Set wsExcel = wbExcel.Sheets(4): wsExcel.Name = "91120"
Set wsExcel = wbExcel.Sheets(5): wsExcel.Name = "91130"
Set wsExcel = wbExcel.Sheets(6): wsExcel.Name = "91121"
Set wsExcel = wbExcel.Sheets(7): wsExcel.Name = "CDE ECHUS " & wDMS_Min
Set wsExcel = wbExcel.Sheets(8): wsExcel.Name = "ECHUS"
Set wsExcel = wbExcel.Sheets(9): wsExcel.Name = "A ECHOIR SUP 31_12_2018"
Set wsExcel = wbExcel.Sheets(10): wsExcel.Name = "SUP31_12_2018"
Set wsExcel = wbExcel.Sheets(11): wsExcel.Name = "REOUVERTS DEPUIS " & Format(m1mois, "dd_mm_yyyy")

'GoTo Suite1

Call cmdSelect_SQL_Xi_Init(1) 'CDE  & wDMS_Min
Call cmdSelect_SQL_Xi_Init(2) 'CDE >= 93 J
Call cmdSelect_SQL_Xi_Init(3) 'CDE < 93 J
Call cmdSelect_SQL_Xi_Init_Echu(7) 'CDE ECHUS & wDMS_Min
Call cmdSelect_SQL_Xi_Init_Echu(9) 'A ECHOIR SUP31_12_2018
'======================================================================================================
Call cmdSelect_SQL_Xi_Dossier_Init(4) '91120
For K = 1 To arrYBIACPT0_Nb
        K2 = mXls1_Row_C + K
        Call cmdSelect_SQL_Xi_Dossier(4, arrYBIACPT0(K).CLIENACLI, "91120", arrYBIACPT0(K).CLIENARA1) 'remplit feuille 4
        Set wsExcel = wbExcel.Sheets(1) 'ecrit feuille 1
        wsExcel.Cells(K2, 1) = arrYBIACPT0(K).CLIENACLI
        wsExcel.Cells(K2, 2) = arrYBIACPT0(K).CLIENARA1
        For K3 = 1 To arrDev_Nb
            If arrDev_RowT(K3) > 0 Then
                wsExcel.Cells(K2, K3 + 2) = "=91120!G" & arrDev_RowT(K3) 'remplit feuille 4 colonne G
                ''''arrDev_RowT(K3) = 0
            End If
        Next K3
        Set wsExcel = wbExcel.Sheets(2) 'ecrit feuille 2
        wsExcel.Cells(K2, 1) = arrYBIACPT0(K).CLIENACLI
        wsExcel.Cells(K2, 2) = arrYBIACPT0(K).CLIENARA1
        For K3 = 1 To arrDev_Nb
            If arrDev_RowT(K3) > 0 Then
                wsExcel.Cells(K2, K3 + 2) = "=91120!K" & arrDev_RowT(K3) 'remplit feuille 4 colonne J
                ''arrDev_RowT(K3) = 0
            End If
        Next K3
        Set wsExcel = wbExcel.Sheets(3) 'ecrit feuille 3
        wsExcel.Cells(K2, 1) = arrYBIACPT0(K).CLIENACLI
        wsExcel.Cells(K2, 2) = arrYBIACPT0(K).CLIENARA1
        For K3 = 1 To arrDev_Nb
            If arrDev_RowT(K3) > 0 Then
                wsExcel.Cells(K2, K3 + 2) = "=91120!M" & arrDev_RowT(K3) 'remplit feuille 4 colonne K
                arrDev_RowT(K3) = 0
            End If
        Next K3
Next K
' Insère formules TOTAUX                                        '
Set wsExcel = wbExcel.Sheets(1) 'ecrit feuille 1
For K = 1 To arrDev_Nb
    K2 = K + 2
    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_N - 1, K2).FormulaLocal = "=SOMME(" & X & mXls1_Row_C + 1 & ":" & X & mXls1_Row_N - 2 & ")"
    wsExcel.Cells(mXls1_Row_N - 1, K2).Font.Bold = True
Next K
Set wsExcel = wbExcel.Sheets(2) 'ecrit feuille 2
For K = 1 To arrDev_Nb
    K2 = K + 2
    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_N - 1, K2).FormulaLocal = "=SOMME(" & X & mXls1_Row_C + 1 & ":" & X & mXls1_Row_N - 2 & ")"
    wsExcel.Cells(mXls1_Row_N - 1, K2).Font.Bold = True
Next K
Set wsExcel = wbExcel.Sheets(3) 'ecrit feuille 3
For K = 1 To arrDev_Nb
    K2 = K + 2
    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_N - 1, K2).FormulaLocal = "=SOMME(" & X & mXls1_Row_C + 1 & ":" & X & mXls1_Row_N - 2 & ")"
    wsExcel.Cells(mXls1_Row_N - 1, K2).Font.Bold = True
Next K
'__________________________________________________________________________________/////////////////////////////////////////
Call cmdSelect_SQL_Xi_Dossier_Init(8) '91120
For K = 1 To arrYBIACPT0_Nb
        K2 = mXls1_Row_C + K
        Call cmdSelect_SQL_Xi_Dossier_Echu(8, arrYBIACPT0(K).CLIENACLI, "91120", arrYBIACPT0(K).CLIENARA1) 'remplit feuille 8
        Set wsExcel = wbExcel.Sheets(7) 'ecrit feuille 7
        wsExcel.Cells(K2, 1) = arrYBIACPT0(K).CLIENACLI
        wsExcel.Cells(K2, 2) = arrYBIACPT0(K).CLIENARA1
        For K3 = 1 To arrDev_Nb
            If arrDev_RowT(K3) > 0 Then
                wsExcel.Cells(K2, K3 + 2) = "=ECHUS!G" & arrDev_RowT(K3) 'remplit feuille 7 colonne G
                arrDev_RowT(K3) = 0
            End If
        Next K3
Next K
'__________________________________________________________________________________/////////////////////////////////////////
Call cmdSelect_SQL_Xi_Dossier_Init(10) '91120
For K = 1 To arrYBIACPT0_Nb
        K2 = mXls1_Row_C + K
        Call cmdSelect_SQL_Xi_Dossier_Echu(10, arrYBIACPT0(K).CLIENACLI, "91120", arrYBIACPT0(K).CLIENARA1) 'remplit feuille 10
        Set wsExcel = wbExcel.Sheets(9) 'ecrit feuille 9
        wsExcel.Cells(K2, 1) = arrYBIACPT0(K).CLIENACLI
        wsExcel.Cells(K2, 2) = arrYBIACPT0(K).CLIENARA1
        For K3 = 1 To arrDev_Nb
            If arrDev_RowT(K3) > 0 Then
                wsExcel.Cells(K2, K3 + 2) = "=SUP31_12_2018!G" & arrDev_RowT(K3) 'remplit feuille 9 colonne G
                arrDev_RowT(K3) = 0
            End If
        Next K3
Next K
' Insère formules TOTAUX                                        '
Set wsExcel = wbExcel.Sheets(7) 'ecrit feuille 7
For K = 1 To arrDev_Nb
    K2 = K + 2
    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_N - 1, K2).FormulaLocal = "=SOMME(" & X & mXls1_Row_C + 1 & ":" & X & mXls1_Row_N - 2 & ")"
    wsExcel.Cells(mXls1_Row_N - 1, K2).Font.Bold = True
Next K
Set wsExcel = wbExcel.Sheets(9) 'ecrit feuille 9
For K = 1 To arrDev_Nb
    K2 = K + 2
    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_N - 1, K2).FormulaLocal = "=SOMME(" & X & mXls1_Row_C + 1 & ":" & X & mXls1_Row_N - 2 & ")"
    wsExcel.Cells(mXls1_Row_N - 1, K2).Font.Bold = True
Next K
'__________________________________________________________________________________/////////////////////////////////////////
'======================================================================================================
Call cmdSelect_SQL_Xi_Dossier_Init(5) '91130
For K = 1 To arrYBIACPT0_Nb
        K2 = mXls1_Row_N + K
        Call cmdSelect_SQL_Xi_Dossier_PDI(5, arrYBIACPT0(K).CLIENACLI, "9113", arrYBIACPT0(K).CLIENARA1)
        Set wsExcel = wbExcel.Sheets(1)
        wsExcel.Cells(K2, 1) = arrYBIACPT0(K).CLIENACLI
        wsExcel.Cells(K2, 2) = arrYBIACPT0(K).CLIENARA1
        For K3 = 1 To arrDev_Nb
            If arrDev_RowT(K3) > 0 Then
                wsExcel.Cells(K2, K3 + 2) = "=91130!G" & arrDev_RowT(K3)
                '''''arrDev_RowT(K3) = 0
            End If
        Next K3

        Set wsExcel = wbExcel.Sheets(2)
        wsExcel.Cells(K2, 1) = arrYBIACPT0(K).CLIENACLI
        wsExcel.Cells(K2, 2) = arrYBIACPT0(K).CLIENARA1
        For K3 = 1 To arrDev_Nb
            If arrDev_RowT(K3) > 0 Then
                wsExcel.Cells(K2, K3 + 2) = "=91130!K" & arrDev_RowT(K3)
                ''arrDev_RowT(K3) = 0
            End If
        Next K3

        Set wsExcel = wbExcel.Sheets(3)
        wsExcel.Cells(K2, 1) = arrYBIACPT0(K).CLIENACLI
        wsExcel.Cells(K2, 2) = arrYBIACPT0(K).CLIENARA1
        For K3 = 1 To arrDev_Nb
            If arrDev_RowT(K3) > 0 Then
                wsExcel.Cells(K2, K3 + 2) = "=91130!M" & arrDev_RowT(K3)
                arrDev_RowT(K3) = 0
            End If
        Next K3
Next K

Set wsExcel = wbExcel.Sheets(1)
For K = 1 To arrDev_Nb
    K2 = K + 2

    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_SP - 1, K2).FormulaLocal = "=SOMME(" & X & mXls1_Row_N + 1 & ":" & X & mXls1_Row_SP - 2 & ")"
    wsExcel.Cells(mXls1_Row_SP - 1, K2).Font.Bold = True
Next K

Set wsExcel = wbExcel.Sheets(2)
For K = 1 To arrDev_Nb
    K2 = K + 2

    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_SP - 1, K2).FormulaLocal = "=SOMME(" & X & mXls1_Row_N + 1 & ":" & X & mXls1_Row_SP - 2 & ")"
    wsExcel.Cells(mXls1_Row_SP - 1, K2).Font.Bold = True
Next K

Set wsExcel = wbExcel.Sheets(3)
For K = 1 To arrDev_Nb
    K2 = K + 2

    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_SP - 1, K2).FormulaLocal = "=SOMME(" & X & mXls1_Row_N + 1 & ":" & X & mXls1_Row_SP - 2 & ")"
    wsExcel.Cells(mXls1_Row_SP - 1, K2).Font.Bold = True
Next K
'======================================================================================================
Call cmdSelect_SQL_Xi_Dossier_Init(6) '91121
For K = 1 To arrYBIACPT0_Nb
        K2 = mXls1_Row_SP + K
        Call cmdSelect_SQL_Xi_ZCAUDOS0(6, arrYBIACPT0(K).CLIENACLI, "91121", arrYBIACPT0(K).CLIENARA1) 'remplit feuille 6

        Set wsExcel = wbExcel.Sheets(1) 'ecrit feuille 1
        wsExcel.Cells(K2, 1) = arrYBIACPT0(K).CLIENACLI
        wsExcel.Cells(K2, 2) = arrYBIACPT0(K).CLIENARA1
        For K3 = 1 To arrDev_Nb
            If arrDev_RowT(K3) > 0 Then
                wsExcel.Cells(K2, K3 + 2) = "=91121!G" & arrDev_RowT(K3) 'remplit feuille 6 colonne G
                '''''arrDev_RowT(K3) = 0
            End If
        Next K3

        Set wsExcel = wbExcel.Sheets(2) 'ecrit feuille 2
        wsExcel.Cells(K2, 1) = arrYBIACPT0(K).CLIENACLI
        wsExcel.Cells(K2, 2) = arrYBIACPT0(K).CLIENARA1
        For K3 = 1 To arrDev_Nb
            If arrDev_RowT(K3) > 0 Then
                wsExcel.Cells(K2, K3 + 2) = "=91121!K" & arrDev_RowT(K3) 'remplit feuille 6 colonne J
                ''arrDev_RowT(K3) = 0
            End If
        Next K3

        Set wsExcel = wbExcel.Sheets(3) 'ecrit feuille 3
        wsExcel.Cells(K2, 1) = arrYBIACPT0(K).CLIENACLI
        wsExcel.Cells(K2, 2) = arrYBIACPT0(K).CLIENARA1
        For K3 = 1 To arrDev_Nb
            If arrDev_RowT(K3) > 0 Then
                wsExcel.Cells(K2, K3 + 2) = "=91121!M" & arrDev_RowT(K3) 'remplit feuille 6 colonne K
                arrDev_RowT(K3) = 0
            End If
        Next K3
Next K
' Insère formules TOTAUX                                        '
Set wsExcel = wbExcel.Sheets(1) 'ecrit feuille 1
For K = 1 To arrDev_Nb
    K2 = K + 2

    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_T - 1, K2).FormulaLocal = "=SOMME(" & X & mXls1_Row_SP + 1 & ":" & X & mXls1_Row_T - 2 & ")"
    wsExcel.Cells(mXls1_Row_T - 1, K2).Font.Bold = True
Next K

Set wsExcel = wbExcel.Sheets(2) 'ecrit feuille 2
For K = 1 To arrDev_Nb
    K2 = K + 2

    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_T - 1, K2).FormulaLocal = "=SOMME(" & X & mXls1_Row_SP + 1 & ":" & X & mXls1_Row_T - 2 & ")"
    wsExcel.Cells(mXls1_Row_T - 1, K2).Font.Bold = True
Next K


Set wsExcel = wbExcel.Sheets(3) 'ecrit feuille 3
For K = 1 To arrDev_Nb
    K2 = K + 2

    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_T - 1, K2).FormulaLocal = "=SOMME(" & X & mXls1_Row_SP + 1 & ":" & X & mXls1_Row_T - 2 & ")"
    wsExcel.Cells(mXls1_Row_T - 1, K2).Font.Bold = True
Next K
' Insère formules TOTAUX                                        '
Set wsExcel = wbExcel.Sheets(1) 'ecrit feuille 1
For K = 1 To arrDev_Nb
    K2 = K + 2

    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_T + 1, K2).FormulaLocal = "=" & X & mXls1_Row_N - 1 & " + " & X & mXls1_Row_SP - 1 & " + " & X & mXls1_Row_T - 1
    wsExcel.Cells(mXls1_Row_T + 3, K2).FormulaLocal = "=" & X & mXls1_Row_T + 1 & " / " & X & mXls1_Row_T + 2
Next K

X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", arrDev_Nb + 2, 1)

wsExcel.Cells(mXls1_Row_T + 6, mXls1_Col_EUR).FormulaLocal = "=SOMME(C" & mXls1_Row_T + 3 & ":" & X & mXls1_Row_T + 3 & ")"
wsExcel.Cells(mXls1_Row_T + 6, mXls1_Col_EUR).Font.Bold = True
wsExcel.Cells(mXls1_Row_T + 6, mXls1_Col_EUR).Interior.Color = mColor_Y1

Set wsExcel = wbExcel.Sheets(2) 'ecrit feuille 2
For K = 1 To arrDev_Nb
    K2 = K + 2

    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_T + 1, K2).FormulaLocal = "=" & X & mXls1_Row_N - 1 & " + " & X & mXls1_Row_SP - 1 & " + " & X & mXls1_Row_T - 1
    wsExcel.Cells(mXls1_Row_T + 3, K2).FormulaLocal = "=" & X & mXls1_Row_T + 1 & " / " & X & mXls1_Row_T + 2
Next K

X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", arrDev_Nb + 2, 1)

wsExcel.Cells(mXls1_Row_T + 6, mXls1_Col_EUR).FormulaLocal = "=SOMME(C" & mXls1_Row_T + 3 & ":" & X & mXls1_Row_T + 3 & ")"
wsExcel.Cells(mXls1_Row_T + 6, mXls1_Col_EUR).Font.Bold = True
wsExcel.Cells(mXls1_Row_T + 6, mXls1_Col_EUR).Interior.Color = mColor_Y1


Set wsExcel = wbExcel.Sheets(3) 'ecrit feuille 3
For K = 1 To arrDev_Nb
    K2 = K + 2

    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_T + 1, K2).FormulaLocal = "=" & X & mXls1_Row_N - 1 & " + " & X & mXls1_Row_SP - 1 & " + " & X & mXls1_Row_T - 1
    wsExcel.Cells(mXls1_Row_T + 3, K2).FormulaLocal = "=" & X & mXls1_Row_T + 1 & " / " & X & mXls1_Row_T + 2
Next K

X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", arrDev_Nb + 2, 1)

wsExcel.Cells(mXls1_Row_T + 6, mXls1_Col_EUR).FormulaLocal = "=SOMME(C" & mXls1_Row_T + 3 & ":" & X & mXls1_Row_T + 3 & ")"
wsExcel.Cells(mXls1_Row_T + 6, mXls1_Col_EUR).Font.Bold = True
wsExcel.Cells(mXls1_Row_T + 6, mXls1_Col_EUR).Interior.Color = mColor_Y1
'__________________________________________________________________________________/////////////////////////////////////////
' Insère formules TOTAUX                                        '
Set wsExcel = wbExcel.Sheets(7)
For K = 1 To arrDev_Nb
    K2 = K + 2
    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_N + 1, K2).FormulaLocal = "=" & X & mXls1_Row_N - 1
    wsExcel.Cells(mXls1_Row_N + 3, K2).FormulaLocal = "=" & X & mXls1_Row_N + 1 & " / " & X & mXls1_Row_N + 2
Next K
Set wsExcel = wbExcel.Sheets(9)
For K = 1 To arrDev_Nb
    K2 = K + 2
    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", K2, 1)
    wsExcel.Cells(mXls1_Row_N + 1, K2).FormulaLocal = "=" & X & mXls1_Row_N - 1
    wsExcel.Cells(mXls1_Row_N + 3, K2).FormulaLocal = "=" & X & mXls1_Row_N + 1 & " / " & X & mXls1_Row_N + 2
Next K
'__________________________________________________________________________________/////////////////////////////////////////
'======================================================================================================
'======================================================================================================
'06-03-2012 total / racine en CV €
For iSheet = 1 To 3 'écrit feuilles 1 à 3
    Set wsExcel = wbExcel.Sheets(iSheet)
    K1 = arrDev_Nb + 3
    wsExcel.Columns(K1).ColumnWidth = 15: wsExcel.Columns(K1).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
    For iRow = 1 To mXls1_Row_T + 1
        If wsExcel.Cells(iRow, 3).Interior.Color <> mColor_GB Then
            mSolde = 0
            For K = 1 To arrDev_Nb
                K2 = K + 2
                mSolde = mSolde + wsExcel.Cells(iRow, K2) / arrDev_Cours(K)
            Next K
            If mSolde <> 0 Then wsExcel.Cells(iRow, K1) = mSolde
            wsExcel.Cells(iRow, K1).Interior.Color = mColor_Y1
       Else
            wsExcel.Cells(iRow, K1) = "cv Euro":
            wsExcel.Cells(iRow, K1).Interior.Color = mColor_GB
            wsExcel.Cells(iRow, K1).Font.Color = mColor_Z0
        End If
    Next iRow
Next iSheet
'__________________________________________________________________________________/////////////////////////////////////////
Set wsExcel = wbExcel.Sheets(7)
K1 = arrDev_Nb + 3
wsExcel.Columns(K1).ColumnWidth = 15: wsExcel.Columns(K1).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
For iRow = 1 To mXls1_Row_N - 1
    If wsExcel.Cells(iRow, 3).Interior.Color <> mColor_GB Then
        mSolde = 0
        For K = 1 To arrDev_Nb
            K2 = K + 2
            mSolde = mSolde + wsExcel.Cells(iRow, K2) / arrDev_Cours(K)
        Next K
        If mSolde <> 0 Then wsExcel.Cells(iRow, K1) = mSolde
        wsExcel.Cells(iRow, K1).Interior.Color = mColor_Y1
   Else
        wsExcel.Cells(iRow, K1) = "cv Euro":
        wsExcel.Cells(iRow, K1).Interior.Color = mColor_GB
        wsExcel.Cells(iRow, K1).Font.Color = mColor_Z0
    End If
Next iRow
'__________________________________________________________________________________/////////////////////////////////////////
Set wsExcel = wbExcel.Sheets(9)
K1 = arrDev_Nb + 3
wsExcel.Columns(K1).ColumnWidth = 15: wsExcel.Columns(K1).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
For iRow = 1 To mXls1_Row_N - 1
    If wsExcel.Cells(iRow, 3).Interior.Color <> mColor_GB Then
        mSolde = 0
        For K = 1 To arrDev_Nb
            K2 = K + 2
            mSolde = mSolde + wsExcel.Cells(iRow, K2) / arrDev_Cours(K)
        Next K
        If mSolde <> 0 Then wsExcel.Cells(iRow, K1) = mSolde
        wsExcel.Cells(iRow, K1).Interior.Color = mColor_Y1
   Else
        wsExcel.Cells(iRow, K1) = "cv Euro":
        wsExcel.Cells(iRow, K1).Interior.Color = mColor_GB
        wsExcel.Cells(iRow, K1).Font.Color = mColor_Z0
    End If
Next iRow
'__________________________________________________________________________________/////////////////////////////////////////
'Total général colonne I feuilles 4-5-6-8-10
    For K1 = 4 To 10
        If K1 < 7 Or K1 = 8 Or K1 = 10 Then
            totalcvEuro = 0
            totalSup93 = 0
            totalInf93 = 0
            Set wsExcel = wbExcel.Sheets(K1)
            derniere_ligne = retourne_dernligne(wsExcel, 4)
            For iRow = 2 To derniere_ligne
                If wsExcel.Cells(iRow, 9).Font.Bold = True Then
                    totalcvEuro = totalcvEuro + wsExcel.Cells(iRow, 9)
                End If
                If wsExcel.Cells(iRow, 12).Font.Bold = True Then
                    totalSup93 = totalSup93 + wsExcel.Cells(iRow, 12)
                End If
                If wsExcel.Cells(iRow, 14).Font.Bold = True Then
                    totalInf93 = totalInf93 + wsExcel.Cells(iRow, 14)
                End If
            Next iRow
            wsExcel.Cells(derniere_ligne, 1) = "Total €"
            wsExcel.Cells(derniere_ligne, 1).Font.Bold = True
            wsExcel.Cells(derniere_ligne, 9) = totalcvEuro
            wsExcel.Cells(derniere_ligne, 9).Font.Bold = True
            wsExcel.Cells(derniere_ligne, 12) = totalSup93
            wsExcel.Cells(derniere_ligne, 12).Font.Bold = True
            wsExcel.Cells(derniere_ligne, 14) = totalInf93
            wsExcel.Cells(derniere_ligne, 14).Font.Bold = True
            For K = 1 To 16
                wsExcel.Cells(derniere_ligne, K).Interior.Color = mColor_Y1
            Next K
        End If
    Next K1
'======================================================================================================
'totaux supplémentaires à garder à cet endroit
    Set wsExcel = wbExcel.Sheets(8)
    derniere_ligne = retourne_dernligne(wsExcel, 4)
    Set wsExcel = wbExcel.Sheets(7)
    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", arrDev_Nb + 2, 1)
    wsExcel.Cells(mXls1_Row_N + 4, mXls1_Col_EUR).FormulaLocal = "=ECHUS!L" & derniere_ligne
    wsExcel.Cells(mXls1_Row_N + 4, mXls1_Col_EUR).Font.Bold = True
    wsExcel.Cells(mXls1_Row_N + 4, mXls1_Col_EUR).Interior.Color = mColor_Y1
    
    wsExcel.Cells(mXls1_Row_N + 5, mXls1_Col_EUR).FormulaLocal = "=ECHUS!N" & derniere_ligne
    wsExcel.Cells(mXls1_Row_N + 5, mXls1_Col_EUR).Font.Bold = True
    wsExcel.Cells(mXls1_Row_N + 5, mXls1_Col_EUR).Interior.Color = mColor_Y1
    
    wsExcel.Cells(mXls1_Row_N + 6, mXls1_Col_EUR).FormulaLocal = "=SOMME(C" & mXls1_Row_N + 3 & ":" & X & mXls1_Row_N + 3 & ")"
    wsExcel.Cells(mXls1_Row_N + 6, mXls1_Col_EUR).Font.Bold = True
    wsExcel.Cells(mXls1_Row_N + 6, mXls1_Col_EUR).Interior.Color = mColor_Y1

    Set wsExcel = wbExcel.Sheets(1)
    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", arrDev_Nb + 2, 1)
    wsExcel.Cells(mXls1_Row_T + 4, mXls1_Col_EUR).value = Sheets(2).Cells(mXls1_Row_T + 6, mXls1_Col_EUR).value
    wsExcel.Cells(mXls1_Row_T + 4, mXls1_Col_EUR).Font.Bold = True
    wsExcel.Cells(mXls1_Row_T + 4, mXls1_Col_EUR).Interior.Color = mColor_Y1

    wsExcel.Cells(mXls1_Row_T + 5, mXls1_Col_EUR).FormulaLocal = Sheets(3).Cells(mXls1_Row_T + 6, mXls1_Col_EUR).value
    wsExcel.Cells(mXls1_Row_T + 5, mXls1_Col_EUR).Font.Bold = True
    wsExcel.Cells(mXls1_Row_T + 5, mXls1_Col_EUR).Interior.Color = mColor_Y1

    Set wsExcel = wbExcel.Sheets(10)
    derniere_ligne = retourne_dernligne(wsExcel, 4)
    Set wsExcel = wbExcel.Sheets(9)
    X = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", arrDev_Nb + 2, 1)
    wsExcel.Cells(mXls1_Row_N + 4, mXls1_Col_EUR).FormulaLocal = "=SUP31_12_2018!L" & derniere_ligne
    wsExcel.Cells(mXls1_Row_N + 4, mXls1_Col_EUR).Font.Bold = True
    wsExcel.Cells(mXls1_Row_N + 4, mXls1_Col_EUR).Interior.Color = mColor_Y1

    wsExcel.Cells(mXls1_Row_N + 5, mXls1_Col_EUR).FormulaLocal = "=SUP31_12_2018!N" & derniere_ligne
    wsExcel.Cells(mXls1_Row_N + 5, mXls1_Col_EUR).Font.Bold = True
    wsExcel.Cells(mXls1_Row_N + 5, mXls1_Col_EUR).Interior.Color = mColor_Y1

    wsExcel.Cells(mXls1_Row_N + 6, mXls1_Col_EUR).FormulaLocal = "=SOMME(C" & mXls1_Row_N + 3 & ":" & X & mXls1_Row_N + 3 & ")"
    wsExcel.Cells(mXls1_Row_N + 6, mXls1_Col_EUR).Font.Bold = True
    wsExcel.Cells(mXls1_Row_N + 6, mXls1_Col_EUR).Interior.Color = mColor_Y1

'======================================================================================================

' 25/02/2019                       Nommage des cellules utiles pour DER classeur auto_intragroup
Set wsExcel = wbExcel.Sheets(1) 'ecrit feuille 1
ReDim iName(1 To 4)
For K = 1 To 4
    iName(K) = 0
Next K
For iRow = 1 To mXls1_Row_T - 1
    If wsExcel.Cells(iRow, 1) = "11001" Then
        iName(1) = iName(1) + 1
        X = Mid(CStr(100 + iName(1)), 2)
        With wsExcel.Cells(iRow, 22)
            .Name = "_CV11001" & X
            .value = wsExcel.Cells(iRow, 22).value
        End With
    ElseIf wsExcel.Cells(iRow, 1) = "11005" Then
        iName(2) = iName(2) + 1
        X = Mid(CStr(100 + iName(2)), 2)
        With wsExcel.Cells(iRow, 22)
            .Name = "_CV11005" & X
            .value = wsExcel.Cells(iRow, 22).value
        End With
    ElseIf wsExcel.Cells(iRow, 1) = "11008" Then
        iName(3) = iName(3) + 1
        X = Mid(CStr(100 + iName(3)), 2)
        With wsExcel.Cells(iRow, 22)
            .Name = "_CV11008" & X
            .value = wsExcel.Cells(iRow, 22).value
        End With
    ElseIf wsExcel.Cells(iRow, 1) = "11084" Then
        iName(4) = iName(4) + 1
        X = Mid(CStr(100 + iName(4)), 2)
        With wsExcel.Cells(iRow, 22)
            .Name = "_CV11084" & X
            .value = wsExcel.Cells(iRow, 22).value
        End With
    End If
Next iRow
'                                                                                               '
Suite1:
' 26/02/2019                       liste des dossiers réouverts depuis moins de 1 mois
'======================================================================================================
Call cmdSelect_SQL_Xi_Dossier_Init(11) 'REOUVERTS DEPUIS
xSql = "select CDOMODDOS,CDOMODOUV,CDOMODCLO,CDOMODDMO,CDODOSNOT,CLIENARA1, CDOMODDON,CDOMODBEN from " & paramIBM_Library_SAB & ".ZCDOMOD0, " _
     & paramIBM_Library_SAB & ".ZCDODOS0," & paramIBM_Library_SAB & ".ZCLIENA0 " _
     & " WHERE CDOMODEVE = '07' and CDOMODDMO >= " & m1mois_IBM _
     & " and CDODOSDOS = CDOMODDOS and CLIENACLI = CDODOSNOT"
Set rsSab = cnsab.Execute(xSql)
iRow = 1
Do While Not rsSab.EOF
    iRow = iRow + 1
    wsExcel.Cells(iRow, 1) = rsSab("CDOMODDOS")
    wsExcel.Cells(iRow, 2) = Trim(rsSab("CDODOSNOT")) & " " & Trim(rsSab("CLIENARA1"))
    wsExcel.Cells(iRow, 2).HorizontalAlignment = Excel.xlHAlignLeft
    wsExcel.Cells(iRow, 3) = retourne_Nom_Beneficiaire_Donneur(Trim(rsSab("CDOMODDON")))
    wsExcel.Cells(iRow, 3).HorizontalAlignment = Excel.xlHAlignLeft
    wsExcel.Cells(iRow, 4) = retourne_Nom_Beneficiaire_Donneur(Trim(rsSab("CDOMODBEN")))
    wsExcel.Cells(iRow, 4).HorizontalAlignment = Excel.xlHAlignLeft
    X = rsSab("CDOMODOUV")
    If X <> "" Then
        Xd = CDbl(X) + 19000000
        X = CStr(Xd)
        wsExcel.Cells(iRow, 5) = "'" & Mid(X, 7) & "/" & Mid(X, 5, 2) & "/" & Left(X, 4)
    End If
    X = rsSab("CDOMODCLO")
    If X <> "" Then
        Xd = CDbl(X) + 19000000
        X = CStr(Xd)
        wsExcel.Cells(iRow, 6) = "'" & Mid(X, 7) & "/" & Mid(X, 5, 2) & "/" & Left(X, 4)
    End If
    X = rsSab("CDOMODDMO")
    If X <> "" Then
        Xd = CDbl(X) + 19000000
        X = CStr(Xd)
        wsExcel.Cells(iRow, 7) = "'" & Mid(X, 7) & "/" & Mid(X, 5, 2) & "/" & Left(X, 4)
    End If
    rsSab.MoveNext
Loop

'                                                                                               '

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

Call MsgBox("Export terminé !")

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
    Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents
    
    wbExcel.SaveAs wFile
    wbExcel.Close
    appExcel.Quit

End Sub

Public Sub cmdSelect_SQL_Xc_Dossier(lSheet As Integer, DOSCD7KCN As String)

'On Error GoTo Error_Handler
Dim xSql As String
Dim X As String, K As Integer
Dim DAmjD As Long, DAmjF As Long, Nb1 As Long, Nb2 As Long, xCur As Currency
'__________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "> Exportation ........ " & lSheet & "-" & DOSCD7KCN): DoEvents

Set wsExcel = wbExcel.Sheets(lSheet)
'wsExcel.Name = DOSCD7KCN & "_PDif"

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 75

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Liste des commissions " & DOSCD7KCN & " à provisionner, arrêté au " & dateImp10(wAmjMin) _
                                & vbCr & "&B&U&10(en excluant les dossiers annulés jusqu'au " & dateImp10(wAmjMax) & ")" & vbCr
wsExcel.PageSetup.CenterHorizontally = True

wsExcel.PageSetup.PrintTitleRows = "$A1:$R1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignRight
    .WrapText = True
    .Font.Size = 8
    .Font.Name = "Arial Unicode MS"
    .RowHeight = 17
End With



wsExcel.Columns(1).ColumnWidth = 5: wsExcel.Cells(1, 1) = "Code": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 7: wsExcel.Cells(1, 2) = "Dossier"
wsExcel.Columns(3).ColumnWidth = 4: wsExcel.Cells(1, 3) = "C/N": wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(4).ColumnWidth = 7: wsExcel.Cells(1, 4) = "Client": wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(5).ColumnWidth = 10: wsExcel.Cells(1, 5) = "Validité du": wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(6).ColumnWidth = 10: wsExcel.Cells(1, 6) = "au": wsExcel.Columns(6).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(7).ColumnWidth = 12: wsExcel.Cells(1, 7) = "Solde": wsExcel.Columns(7).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(8).ColumnWidth = 5: wsExcel.Cells(1, 8) = "Devise": wsExcel.Columns(8).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(9).ColumnWidth = 12: wsExcel.Cells(1, 9) = "Mt RGL": wsExcel.Columns(9).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(10).ColumnWidth = 10: wsExcel.Cells(1, 10) = "Date RGL": wsExcel.Columns(10).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(11).ColumnWidth = 10: wsExcel.Cells(1, 11) = "Mt COM": wsExcel.Columns(11).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(12).ColumnWidth = 10: wsExcel.Cells(1, 12) = "Mt perçu": wsExcel.Columns(12).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(13).ColumnWidth = 10: wsExcel.Cells(1, 13) = "Mt à percevoir": wsExcel.Columns(13).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(14).ColumnWidth = 10: wsExcel.Cells(1, 14) = "Mt PDIF": wsExcel.Columns(14).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(15).ColumnWidth = 10: wsExcel.Cells(1, 15) = "<=" & dateImp10(wAmjMin): wsExcel.Columns(15).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(16).ColumnWidth = 10: wsExcel.Cells(1, 16) = "> " & dateImp10(wAmjMin): wsExcel.Columns(16).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(17).ColumnWidth = 7: wsExcel.Cells(1, 17) = "%": wsExcel.Columns(17).NumberFormat = "### ##0.00"
wsExcel.Columns(18).ColumnWidth = 10: wsExcel.Cells(1, 18) = "# Mt compta": wsExcel.Columns(18).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(19).ColumnWidth = 10: wsExcel.Cells(1, 19) = "perçu / ANN": wsExcel.Columns(19).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"



mXls2_Col = 19
For K = 1 To mXls2_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB ' RGB(255, 128, 50)
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next K

mXls2_Row = 1
'mDev = "": mDEV_STotal = 0
mDev_R1 = 2

Call rsYDOSCD70_Init(oldYDOSCD70)

xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSCD70" _
     & " where DOSCD7DSIT = " & wAmjMin & " and DOSCD7KCN = '" & DOSCD7KCN & "'" _
     & " order by DOSCD7DEV,DOSCD7OPE,DOSCD7NUM,DOSCD7KNAT"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF

    V = rsYDOSCD70_GetBuffer(rsSab, xYDOSCD70)
        
    If xYDOSCD70.DOSCD7OPE <> oldYDOSCD70.DOSCD7OPE Or xYDOSCD70.DOSCD7NUM <> oldYDOSCD70.DOSCD7NUM Then
        Call cmdSelect_SQL_Xc_Dossier_S(DOSCD7KCN)
        blnECNFPT_CD7 = False
        If xYDOSCD70.DOSCD7NUM = 111040 Then
            Debug.Print xYDOSCD70.DOSCD7NUM
        End If
    End If
    If xYDOSCD70.DOSCD7OPE = "CDE" And xYDOSCD70.DOSCD7NUM = 117568 And xYDOSCD70.DOSCD7KNAT = 2 Then
        Debug.Print sMTD_COM_G2; "    "; xYDOSCD70.DOSCD7MTD
    End If
    

    
    Select Case xYDOSCD70.DOSCD7KNAT
        Case "0": sMTD_Solde_C = sMTD_Solde_C - xYDOSCD70.DOSCD7MTD: dosYDOSCD70 = xYDOSCD70
        Case "1": sMTD_COM_C = sMTD_COM_C + xYDOSCD70.DOSCD7MTD
        Case "2": sMTD_COM_G2 = sMTD_COM_G2 + xYDOSCD70.DOSCD7MTD
                  
                  oldYDOSCD70.DOSCD7DAMJ = xYDOSCD70.DOSCD7DAMJ
                  If xYDOSCD70.DOSCD7DAMJ > wAmjMin Then
                        sMTD_COM_G2PDIF = sMTD_COM_G2PDIF + xYDOSCD70.DOSCD7MTD
                  Else
                  


                  
                        If xYDOSCD70.DOSCD7DFIN > wAmjMin Then
                            If xYDOSCD70.DOSCD7DDEB Mod 100 = 0 Then
                                xYDOSCD70.DOSCD7DDEB = xYDOSCD70.DOSCD7DDEB + 1
                            End If

                              Nb1 = DateDiff("d", dateImp_Amj(xYDOSCD70.DOSCD7DDEB), wDMS_Min) + 1
                              Nb2 = DateDiff("d", dateImp_Amj(xYDOSCD70.DOSCD7DDEB), dateImp_Amj(xYDOSCD70.DOSCD7DFIN)) + 1
                              sMTD_COM_G2Prata = sMTD_COM_G2Prata + xYDOSCD70.DOSCD7MTD * Nb1 / Nb2
                         Else
                         If xYDOSCD70.DOSCD7DDEB = "20250425" Then
                            MsgBox ("OK2")
                         End If
                              sMTD_COM_G2Prata = sMTD_COM_G2Prata + xYDOSCD70.DOSCD7MTD
                         End If
                    End If
                    If xYDOSCD70.DOSCD7STA = "@1" Then blnECNFPT_CD7 = True
                    
        Case "3":     sMTD_COM_G3 = sMTD_COM_G3 + xYDOSCD70.DOSCD7MTD
        Case "4":     sMTD_UTI_G = sMTD_UTI_G + xYDOSCD70.DOSCD7MTD: oldYDOSCD70.DOSCD7DAMJ = xYDOSCD70.DOSCD7DAMJ
        Case "5": sMTD_TC2 = xYDOSCD70.DOSCD7MTD
    End Select
       
    
    
    rsSab.MoveNext
Loop

xYDOSCD70.DOSCD7NUM = 0
xYDOSCD70.DOSCD7DEV = ""

Call cmdSelect_SQL_Xc_Dossier_S(DOSCD7KCN)

Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & lSheet & "-" & DOSCD7KCN): DoEvents


'_____________________________
Exit Sub

'Error_Handler:
  '  If Not blnAuto Then MsgBox Error, vbCritical, Me.Name
  '  Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & DOSCD7KCN): DoEvents

End Sub

Public Sub cmdSelect_SQL_XE1an_Dossier(lSheet As Integer, lDOSSLDPCI As String)

On Error GoTo Error_Handler
Dim xSql As String
Dim X As String, K As Integer
Dim DAmjD As Long, DAmjF As Long, Nb1 As Long, Nb2 As Long, xCur As Currency
Dim mK2 As Integer
Dim mRow_err As Long
Dim blnOk As Boolean
'__________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "> Exportation ........ " & lSheet & "-" & lDOSSLDPCI): DoEvents

Set wsExcel = wbExcel.Sheets(lSheet)

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 75

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Liste des dossiers" & lDOSSLDPCI & ", arrêté au " & dateImp10(wAmjMin) _
                                & vbCr
wsExcel.PageSetup.CenterHorizontally = True

wsExcel.PageSetup.PrintTitleRows = "$A1:$L1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignRight
    .WrapText = True
    .Font.Size = 8
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.Columns(1).ColumnWidth = 7: wsExcel.Cells(mXls1_Row_C, 1) = "PCI": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 7: wsExcel.Cells(mXls1_Row_C, 2) = "Devise": wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(3).ColumnWidth = 8: wsExcel.Cells(mXls1_Row_C, 3) = "Client": wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(4).ColumnWidth = 8: wsExcel.Cells(mXls1_Row_C, 4) = "Dossier": wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(5).ColumnWidth = 12: wsExcel.Cells(mXls1_Row_C, 5) = "D.ouverture": wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(6).ColumnWidth = 12: wsExcel.Cells(mXls1_Row_C, 6) = "D.validité": wsExcel.Columns(6).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(7).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, 7) = "Montant": wsExcel.Columns(7).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(8).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, 8) = "MTD <= " & dateImp10(wAmjMax): wsExcel.Columns(8).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(9).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, 9) = "MTD > " & dateImp10(wAmjMax): wsExcel.Columns(9).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"

wsExcel.Columns(10).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, 10) = "Total <= " & dateImp10(wAmjMax): wsExcel.Columns(10).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(11).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, 11) = "Total > " & dateImp10(wAmjMax): wsExcel.Columns(11).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(12).ColumnWidth = 16: wsExcel.Cells(mXls1_Row_C, 12) = "!!!!": wsExcel.Columns(12).HorizontalAlignment = Excel.xlHAlignLeft


Call rsYDOSSLD0_Init(oldYDOSSLD0)
xYDOSSLD0 = oldYDOSSLD0

mXls2_Col = 12
For K = 1 To mXls2_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB ' RGB(255, 128, 50)
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next K

mXls2_Row = 1: mK2 = 0: mRow_err = 0

'If Mid$(lDOSSLDPCI, 1, 4) <> "9113" Then

Select Case Mid$(lDOSSLDPCI, 1, 5)
'==========================================================================================================
    Case "91130", "91131"
'==========================================================================================================
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDOREG0 " _
         & " where DOSSLDPCI like '" & lDOSSLDPCI & "%'" _
         & " and DOSSLDSTA not in ('  ','80','90')" _
         & " and CDOREGCOP = DOSSLDOPE and CDOREGDOS = DOSSLDNUM" _
         & " and CDOREGPAI = 3 and CDOREGETA <> '03' and CDOREGCRD = 'D' and substring(CDOREGCAA , 1 , 3) = 'PDI'" _
         & " order by CDOREGDEV,DOSSLDCLI,DOSSLDNUM"
    Set rsSab = cnsab.Execute(xSql)
    
    Do While Not rsSab.EOF
    
        V = rsYDOSSLD0_GetBuffer(rsSab, xYDOSSLD0)
        
        If oldYDOSSLD0.DOSSLDPCI <> xYDOSSLD0.DOSSLDPCI _
        Or oldYDOSSLD0.DOSSLDDEV <> xYDOSSLD0.DOSSLDDEV _
        Or oldYDOSSLD0.DOSSLDCLI <> xYDOSSLD0.DOSSLDCLI Then
            If mK2 > 0 Then
                wsExcel.Cells(mXls2_Row, 2).Font.Bold = True
                
                wsExcel.Cells(mXls2_Row, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & mXls2_Row & ")"
                wsExcel.Cells(mXls2_Row, 10).Interior.Color = mColor_Y0
                wsExcel.Cells(mXls2_Row, 10).Font.Bold = True
    
                wsExcel.Cells(mXls2_Row, 11).FormulaLocal = "=SOMME(I" & mK2 & ":I" & mXls2_Row & ")"
                wsExcel.Cells(mXls2_Row, 11).Interior.Color = mColor_Y1
                wsExcel.Cells(mXls2_Row, 11).Font.Bold = True
                    
                wsExcel.Range("A" & mXls2_Row & ":K" & mXls2_Row).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & mXls2_Row & ":K" & mXls2_Row).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
    
                For K = 1 To arrYBIACPT0_Nb
                    If oldYDOSSLD0.DOSSLDDEV = arrYBIACPT0(K).COMPTEDEV _
                    And oldYDOSSLD0.DOSSLDCLI = arrYBIACPT0(K).CLIENACLI Then
                        arrRow(K) = mXls2_Row
                        arrRow_Err(K) = mRow_err
                        Exit For
                    End If
                Next K
            End If
            oldYDOSSLD0 = xYDOSSLD0
            mRow_err = 0
            mK2 = mXls2_Row + 1
        End If
        mXls2_Row = mXls2_Row + 1
        
        
        wsExcel.Cells(mXls2_Row, 1) = xYDOSSLD0.DOSSLDPCI
        wsExcel.Cells(mXls2_Row, 2) = xYDOSSLD0.DOSSLDDEV
        wsExcel.Cells(mXls2_Row, 3) = xYDOSSLD0.DOSSLDCLI
        wsExcel.Cells(mXls2_Row, 4) = xYDOSSLD0.DOSSLDNUM
        wsExcel.Cells(mXls2_Row, 5) = dateImp10(rsSab("CDOREGDEN") + 19000000)
        DAmjF = rsSab("CDOREGDRE") + 19000000
        wsExcel.Cells(mXls2_Row, 6) = dateImp10(DAmjF)
        xCur = rsSab("CDOREGMON")
        If DAmjF > wAmjMax Then
            wsExcel.Cells(mXls2_Row, 9) = xCur
        Else
            wsExcel.Cells(mXls2_Row, 8) = xCur
        End If
        
        wsExcel.Cells(mXls2_Row, 8).Interior.Color = mColor_Y0:
        wsExcel.Cells(mXls2_Row, 9).Interior.Color = mColor_Y1:


        rsSab.MoveNext
    Loop
'==========================================================================================================
    Case "91132"
'==========================================================================================================
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDOREG0 " _
         & " where DOSSLDPCI like '" & lDOSSLDPCI & "%'" _
         & " and DOSSLDSTA not in ('  ','80','90')" _
         & " and CDOREGCOP = DOSSLDOPE and CDOREGDOS = DOSSLDNUM" _
         & " and CDOREGPAI = 1 and CDOREGETA <> '03' and CDOREGCRD = 'D'" _
         & " order by CDOREGDEV,DOSSLDCLI,DOSSLDNUM"
    Set rsSab = cnsab.Execute(xSql)
    
    Do While Not rsSab.EOF
    
        V = rsYDOSSLD0_GetBuffer(rsSab, xYDOSSLD0)
        
        If oldYDOSSLD0.DOSSLDPCI <> xYDOSSLD0.DOSSLDPCI _
        Or oldYDOSSLD0.DOSSLDDEV <> xYDOSSLD0.DOSSLDDEV _
        Or oldYDOSSLD0.DOSSLDCLI <> xYDOSSLD0.DOSSLDCLI Then
            If mK2 > 0 Then
                wsExcel.Cells(mXls2_Row, 2).Font.Bold = True
                
                wsExcel.Cells(mXls2_Row, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & mXls2_Row & ")"
                wsExcel.Cells(mXls2_Row, 10).Interior.Color = mColor_Y0
                wsExcel.Cells(mXls2_Row, 10).Font.Bold = True
    
                wsExcel.Cells(mXls2_Row, 11).FormulaLocal = "=SOMME(I" & mK2 & ":I" & mXls2_Row & ")"
                wsExcel.Cells(mXls2_Row, 11).Interior.Color = mColor_Y1
                wsExcel.Cells(mXls2_Row, 11).Font.Bold = True
                    
                wsExcel.Range("A" & mXls2_Row & ":K" & mXls2_Row).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & mXls2_Row & ":K" & mXls2_Row).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
    
                For K = 1 To arrYBIACPT0_Nb
                    If oldYDOSSLD0.DOSSLDDEV = arrYBIACPT0(K).COMPTEDEV _
                    And oldYDOSSLD0.DOSSLDCLI = arrYBIACPT0(K).CLIENACLI Then
                        arrRow(K) = mXls2_Row
                        arrRow_Err(K) = mRow_err
                        Exit For
                    End If
                Next K
            End If
            oldYDOSSLD0 = xYDOSSLD0
            mRow_err = 0
            mK2 = mXls2_Row + 1
        End If
        mXls2_Row = mXls2_Row + 1
        
        
        wsExcel.Cells(mXls2_Row, 1) = xYDOSSLD0.DOSSLDPCI
        wsExcel.Cells(mXls2_Row, 2) = xYDOSSLD0.DOSSLDDEV
        wsExcel.Cells(mXls2_Row, 3) = xYDOSSLD0.DOSSLDCLI
        wsExcel.Cells(mXls2_Row, 4) = xYDOSSLD0.DOSSLDNUM
        wsExcel.Cells(mXls2_Row, 5) = dateImp10(rsSab("CDOREGDEN") + 19000000)
        DAmjF = rsSab("CDOREGDRE") + 19000000
        wsExcel.Cells(mXls2_Row, 6) = dateImp10(DAmjF)
        xCur = rsSab("CDOREGMON")
        If DAmjF > wAmjMax Then
            wsExcel.Cells(mXls2_Row, 9) = xCur
        Else
            wsExcel.Cells(mXls2_Row, 8) = xCur
        End If
        
        wsExcel.Cells(mXls2_Row, 8).Interior.Color = mColor_Y0:
        wsExcel.Cells(mXls2_Row, 9).Interior.Color = mColor_Y1:


        rsSab.MoveNext
    Loop
'==========================================================================================================
    Case "98052"
'==========================================================================================================
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDOREG0 " _
         & " where DOSSLDPCI like '" & lDOSSLDPCI & "%'" _
         & " and DOSSLDSTA not in ('  ','80','90')" _
         & " and CDOREGCOP = DOSSLDOPE and CDOREGDOS = DOSSLDNUM" _
         & " and CDOREGPAI = 3 and CDOREGETA <> '03' and CDOREGCRD = 'D'" _
         & " order by CDOREGDEV,DOSSLDCLI,DOSSLDNUM"
    Set rsSab = cnsab.Execute(xSql)
    
    Do While Not rsSab.EOF
    
        V = rsYDOSSLD0_GetBuffer(rsSab, xYDOSSLD0)
        
        If oldYDOSSLD0.DOSSLDPCI <> xYDOSSLD0.DOSSLDPCI _
        Or oldYDOSSLD0.DOSSLDDEV <> xYDOSSLD0.DOSSLDDEV _
        Or oldYDOSSLD0.DOSSLDCLI <> xYDOSSLD0.DOSSLDCLI Then
            If mK2 > 0 Then
                wsExcel.Cells(mXls2_Row, 2).Font.Bold = True
                
                wsExcel.Cells(mXls2_Row, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & mXls2_Row & ")"
                wsExcel.Cells(mXls2_Row, 10).Interior.Color = mColor_Y0
                wsExcel.Cells(mXls2_Row, 10).Font.Bold = True
    
                wsExcel.Cells(mXls2_Row, 11).FormulaLocal = "=SOMME(I" & mK2 & ":I" & mXls2_Row & ")"
                wsExcel.Cells(mXls2_Row, 11).Interior.Color = mColor_Y1
                wsExcel.Cells(mXls2_Row, 11).Font.Bold = True
                    
                wsExcel.Range("A" & mXls2_Row & ":K" & mXls2_Row).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & mXls2_Row & ":K" & mXls2_Row).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
    
                For K = 1 To arrYBIACPT0_Nb
                    If oldYDOSSLD0.DOSSLDDEV = arrYBIACPT0(K).COMPTEDEV _
                    And oldYDOSSLD0.DOSSLDCLI = arrYBIACPT0(K).CLIENACLI Then
                        arrRow(K) = mXls2_Row
                        arrRow_Err(K) = mRow_err
                        Exit For
                    End If
                Next K
            End If
            oldYDOSSLD0 = xYDOSSLD0
            mRow_err = 0
            mK2 = mXls2_Row + 1
        End If
        mXls2_Row = mXls2_Row + 1
        
        
        wsExcel.Cells(mXls2_Row, 1) = xYDOSSLD0.DOSSLDPCI
        wsExcel.Cells(mXls2_Row, 2) = xYDOSSLD0.DOSSLDDEV
        wsExcel.Cells(mXls2_Row, 3) = xYDOSSLD0.DOSSLDCLI
        wsExcel.Cells(mXls2_Row, 4) = xYDOSSLD0.DOSSLDNUM
        wsExcel.Cells(mXls2_Row, 5) = dateImp10(rsSab("CDOREGDEN") + 19000000)
        DAmjF = rsSab("CDOREGDRE") + 19000000
        wsExcel.Cells(mXls2_Row, 6) = dateImp10(DAmjF)
        xCur = rsSab("CDOREGMON")
        If DAmjF > wAmjMax Then
            wsExcel.Cells(mXls2_Row, 9) = xCur
        Else
            wsExcel.Cells(mXls2_Row, 8) = xCur
        End If
        
        wsExcel.Cells(mXls2_Row, 8).Interior.Color = mColor_Y0:
        wsExcel.Cells(mXls2_Row, 9).Interior.Color = mColor_Y1:


        rsSab.MoveNext
    Loop

'==========================================================================================================
    Case Else
'=======================================================================================================================
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDODOS0 " _
         & " where DOSSLDPCI like '" & lDOSSLDPCI & "%' " _
         & " and DOSSLDSTA not in ('  ','80','90')" _
         & " and CDODOSCOP = DOSSLDOPE and CDODOSDOS = DOSSLDNUM" _
         & " order by DOSSLDDEV, DOSSLDCLI , DOSSLDNUM"
    
    Set rsSab = cnsab.Execute(xSql)
    
    Do While Not rsSab.EOF
        V = rsYDOSSLD0_GetBuffer(rsSab, xYDOSSLD0)
        If oldYDOSSLD0.DOSSLDPCI <> xYDOSSLD0.DOSSLDPCI _
        Or oldYDOSSLD0.DOSSLDDEV <> xYDOSSLD0.DOSSLDDEV _
        Or oldYDOSSLD0.DOSSLDCLI <> xYDOSSLD0.DOSSLDCLI Then
            If mK2 > 0 Then
                wsExcel.Cells(mXls2_Row, 2).Font.Bold = True
                
                wsExcel.Cells(mXls2_Row, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & mXls2_Row & ")"
                wsExcel.Cells(mXls2_Row, 10).Interior.Color = mColor_Y0
                wsExcel.Cells(mXls2_Row, 10).Font.Bold = True
    
                wsExcel.Cells(mXls2_Row, 11).FormulaLocal = "=SOMME(I" & mK2 & ":I" & mXls2_Row & ")"
                wsExcel.Cells(mXls2_Row, 11).Interior.Color = mColor_Y1
                wsExcel.Cells(mXls2_Row, 11).Font.Bold = True
                    
                wsExcel.Range("A" & mXls2_Row & ":K" & mXls2_Row).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & mXls2_Row & ":K" & mXls2_Row).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
                
                blnOk = False
                For K = 1 To arrYBIACPT0_Nb
                    If oldYDOSSLD0.DOSSLDDEV = arrYBIACPT0(K).COMPTEDEV _
                    And oldYDOSSLD0.DOSSLDCLI = arrYBIACPT0(K).CLIENACLI Then
                        arrRow(K) = mXls2_Row
                        arrRow_Err(K) = mRow_err
                        blnOk = True
                        Exit For
                    End If
                Next K
                If Not blnOk Then
                    wsExcel.Cells(mXls2_Row, 12) = "Compte client inconnu"
                    wsExcel.Cells(mXls2_Row, 12).Interior.Color = vbRed
                    wsExcel.Cells(mXls2_Row, 12).Font.Color = vbYellow
                    For K = 1 To 6
                        wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_W1
                    Next K
    
                End If
                
            End If
            oldYDOSSLD0 = xYDOSSLD0
            mRow_err = 0
            mK2 = mXls2_Row + 1
        End If
        
        mXls2_Row = mXls2_Row + 1
        
        wsExcel.Cells(mXls2_Row, 1) = xYDOSSLD0.DOSSLDPCI
        wsExcel.Cells(mXls2_Row, 2) = xYDOSSLD0.DOSSLDDEV
        wsExcel.Cells(mXls2_Row, 3) = xYDOSSLD0.DOSSLDCLI
        wsExcel.Cells(mXls2_Row, 4) = xYDOSSLD0.DOSSLDNUM
        wsExcel.Cells(mXls2_Row, 5) = dateImp10(rsSab("CDODOSOUV") + 19000000)
        DAmjF = rsSab("CDODOSVAL") + 19000000
        wsExcel.Cells(mXls2_Row, 6) = dateImp10(DAmjF)
        wsExcel.Cells(mXls2_Row, 7) = rsSab("CDODOSMON")
        xCur = -xYDOSSLD0.DOSSLDMSD
        If DAmjF > wAmjMax Then
            wsExcel.Cells(mXls2_Row, 9) = xCur
        Else
            wsExcel.Cells(mXls2_Row, 8) = xCur
        End If
        
        wsExcel.Cells(mXls2_Row, 8).Interior.Color = mColor_Y0:
        wsExcel.Cells(mXls2_Row, 9).Interior.Color = mColor_Y1:
        If xYDOSSLD0.DOSSLDMSD <> xYDOSSLD0.DOSSLDGSD Then
            If xYDOSSLD0.DOSSLDSVC = "03" Then
                mRow_err = mXls2_Row
                wsExcel.Cells(mXls2_Row, 12) = "SD Cpt # Ges"
                wsExcel.Cells(mXls2_Row, 12).Interior.Color = vbRed
                wsExcel.Cells(mXls2_Row, 12).Font.Color = vbYellow
            Else
                wsExcel.Cells(mXls2_Row, 12) = "non comptabilisé"
            End If
            For K = 1 To 6
                wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_W1
            Next K
            
        End If
        
        
        rsSab.MoveNext
    Loop
    

    
End Select
'=================================================================================================
If mK2 > 0 Then
            wsExcel.Cells(mXls2_Row, 2).Font.Bold = True
            
            wsExcel.Cells(mXls2_Row, 10).FormulaLocal = "=SOMME(H" & mK2 & ":H" & mXls2_Row & ")"
            wsExcel.Cells(mXls2_Row, 10).Interior.Color = mColor_Y0
            wsExcel.Cells(mXls2_Row, 10).Font.Bold = True

            wsExcel.Cells(mXls2_Row, 11).FormulaLocal = "=SOMME(I" & mK2 & ":I" & mXls2_Row & ")"
            wsExcel.Cells(mXls2_Row, 11).Interior.Color = mColor_Y1
            wsExcel.Cells(mXls2_Row, 11).Font.Bold = True
                
            wsExcel.Range("A" & mXls2_Row & ":K" & mXls2_Row).Borders(xlEdgeBottom).Weight = xlThick
            wsExcel.Range("A" & mXls2_Row & ":K" & mXls2_Row).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
        For K = 1 To arrYBIACPT0_Nb
            If oldYDOSSLD0.DOSSLDDEV = arrYBIACPT0(K).COMPTEDEV _
            And oldYDOSSLD0.DOSSLDCLI = arrYBIACPT0(K).CLIENACLI Then
                arrRow(K) = mXls2_Row
                Exit For
            End If
        Next K
End If

'=================================================================================================================
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & lSheet & "-" & lDOSSLDPCI): DoEvents


'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & lDOSSLDPCI): DoEvents

End Sub


Public Sub cmdSelect_SQL_Xi_Dossier(lSheet As Integer, lDOSSLDCLI As String, lDOSSLDPCI As String, lCLIENARA1 As String)

On Error GoTo Error_Handler
Dim xSql As String
Dim X As String, K As Integer
Dim DAmjD As Long, DAmjF As Long, Nb1 As Long, Nb2 As Long, xCur As Currency
'__________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "> Exportation ........ " & lSheet & "-" & lDOSSLDCLI): DoEvents

Set wsExcel = wbExcel.Sheets(lSheet)


Call rsYDOSCD70_Init(oldYDOSCD70)
If lDOSSLDCLI = "0011001" Then
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDODOS0 " _
         & " where DOSSLDPCI like '" & lDOSSLDPCI & "%' and DOSSLDCLI = '" & lDOSSLDCLI & "'" _
         & " and DOSSLDSTA not in ('  ','80','90')" _
         & " and CDODOSCOP = DOSSLDOPE and CDODOSDOS = DOSSLDNUM" _
         & " and CDODOSOUV >= " & mCDODOSOUV_11001 _
         & " order by DOSSLDDEV,DOSSLDOPE,DOSSLDNUM"
Else
    If InStr(LFB_RacinesExclues, lDOSSLDCLI) > 0 Then

        xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDODOS0 " _
             & " where DOSSLDPCI like '" & lDOSSLDPCI & "%' and DOSSLDCLI = '" & lDOSSLDCLI & "'" _
             & " and DOSSLDSTA not in ('  ','80','90')" _
             & " and CDODOSCOP = DOSSLDOPE and CDODOSDOS = DOSSLDNUM" _
             & " and CDODOSOUV >= " & mCDODOSOUV_11012 _
             & " order by DOSSLDDEV,DOSSLDOPE,DOSSLDNUM"
    Else
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDODOS0 " _
             & " where DOSSLDPCI like '" & lDOSSLDPCI & "%' and DOSSLDCLI = '" & lDOSSLDCLI & "'" _
             & " and DOSSLDSTA not in ('  ','80','90')" _
             & " and CDODOSCOP = DOSSLDOPE and CDODOSDOS = DOSSLDNUM" _
             & " order by DOSSLDDEV,DOSSLDOPE,DOSSLDNUM"
    End If
End If

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF

    V = rsYDOSSLD0_GetBuffer(rsSab, xYDOSSLD0)
        
    If xYDOSSLD0.DOSSLDDEV <> oldYDOSSLD0.DOSSLDDEV Then
        Call cmdSelect_SQL_Xi_Dossier_T(lCLIENARA1, lDOSSLDPCI)
    End If
    mXls2_Row = mXls2_Row + 1
    DAmjD = rsSab("CDODOSOUV") + 19000000
    DAmjF = rsSab("CDODOSVAL") + 19000000
    
    wsExcel.Cells(mXls2_Row, 1) = xYDOSSLD0.DOSSLDOPE
    wsExcel.Cells(mXls2_Row, 2) = xYDOSSLD0.DOSSLDNUM
    wsExcel.Cells(mXls2_Row, 3) = rsSab("CDODOSCON")
    wsExcel.Cells(mXls2_Row, 4) = xYDOSSLD0.DOSSLDCLI
    wsExcel.Cells(mXls2_Row, 5) = dateImp10(DAmjD)
    wsExcel.Cells(mXls2_Row, 6) = dateImp10(DAmjF)
    wsExcel.Cells(mXls2_Row, 7) = xYDOSSLD0.DOSSLDMSD
    wsExcel.Cells(mXls2_Row, 8) = xYDOSSLD0.DOSSLDDEV
    '//////////////////////////////////
    If xYDOSSLD0.DOSSLDDEV <> "EUR" Then
        wsExcel.Cells(mXls2_Row, 9) = CDbl(xYDOSSLD0.DOSSLDMSD) / arrDev_Cours(retourne_indice_devise(xYDOSSLD0.DOSSLDDEV))
    Else
        wsExcel.Cells(mXls2_Row, 9) = xYDOSSLD0.DOSSLDMSD
    End If
    '//////////////////////////////////
    If xYDOSSLD0.DOSSLDNUM = "116908" Then
        MsgBox ("OK")
    End If
    If DAmjF <= wAmjMin Then
        wsExcel.Cells(mXls2_Row, 6).Font.Color = vbMagenta
        Nb1 = DateDiff("d", dateImp_Amj(DAmjD), wDMS_Min) + 1
    Else
        Nb1 = DateDiff("d", dateImp_Amj(DAmjD), dateImp_Amj(DAmjF)) + 1
    End If
    wsExcel.Cells(mXls2_Row, 10) = Nb1
    If Nb1 >= 93 Then
        If xYDOSSLD0.DOSSLDDEV <> "EUR" Then
            wsExcel.Cells(mXls2_Row, 11) = xYDOSSLD0.DOSSLDMSD
            wsExcel.Cells(mXls2_Row, 12) = CDbl(xYDOSSLD0.DOSSLDMSD) / arrDev_Cours(retourne_indice_devise(xYDOSSLD0.DOSSLDDEV))
        Else
            wsExcel.Cells(mXls2_Row, 11) = xYDOSSLD0.DOSSLDMSD
            wsExcel.Cells(mXls2_Row, 12) = xYDOSSLD0.DOSSLDMSD
        End If
    Else
        If xYDOSSLD0.DOSSLDDEV <> "EUR" Then
            wsExcel.Cells(mXls2_Row, 13) = xYDOSSLD0.DOSSLDMSD
            wsExcel.Cells(mXls2_Row, 14) = CDbl(xYDOSSLD0.DOSSLDMSD) / arrDev_Cours(retourne_indice_devise(xYDOSSLD0.DOSSLDDEV))
        Else
            wsExcel.Cells(mXls2_Row, 13) = xYDOSSLD0.DOSSLDMSD
            wsExcel.Cells(mXls2_Row, 14) = xYDOSSLD0.DOSSLDMSD
        End If
    End If
    rsSab.MoveNext
Loop

xYDOSSLD0.DOSSLDDEV = ""

Call cmdSelect_SQL_Xi_Dossier_T(lCLIENARA1, lDOSSLDPCI)

Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & lSheet & "-" & lDOSSLDCLI): DoEvents


'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & lDOSSLDCLI): DoEvents

End Sub
Public Sub cmdSelect_SQL_Xi_Dossier_PDI(lSheet As Integer, lDOSSLDCLI As String, lDOSSLDPCI As String, lCLIENARA1 As String)

On Error GoTo Error_Handler
Dim xSql As String
Dim X As String, K As Integer
Dim DAmjD As Long, DAmjF As Long, Nb1 As Long, Nb2 As Long, xCur As Currency
'__________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "> Exportation ........ " & lSheet & "-" & lDOSSLDCLI): DoEvents

Set wsExcel = wbExcel.Sheets(lSheet)


Call rsYDOSCD70_Init(oldYDOSCD70)
If lDOSSLDCLI = "0011001" Then
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDOREG0 , " _
         & paramIBM_Library_SAB & ".ZCDODOS0 " _
         & " where DOSSLDPCI like '" & lDOSSLDPCI & "%' and DOSSLDCLI = '" & lDOSSLDCLI & "'" _
         & " and DOSSLDSTA not in ('  ','80','90')" _
         & " and CDOREGCOP = DOSSLDOPE and CDOREGDOS = DOSSLDNUM" _
         & " and CDOREGPAI = 3 and CDOREGETA <> '03' and CDOREGCRD = 'D' and substring(CDOREGCAA , 1 , 3) = 'PDI'" _
         & " and CDODOSCOP = DOSSLDOPE and CDODOSDOS = DOSSLDNUM" _
         & " and CDODOSOUV >= " & mCDODOSOUV_11001 _
        & " order by DOSSLDDEV,DOSSLDOPE,DOSSLDNUM"
Else
    If InStr(LFB_RacinesExclues, lDOSSLDCLI) > 0 Then
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDOREG0 , " _
             & paramIBM_Library_SAB & ".ZCDODOS0 " _
             & " where DOSSLDPCI like '" & lDOSSLDPCI & "%' and DOSSLDCLI = '" & lDOSSLDCLI & "'" _
             & " and DOSSLDSTA not in ('  ','80','90')" _
             & " and CDOREGCOP = DOSSLDOPE and CDOREGDOS = DOSSLDNUM" _
             & " and CDOREGPAI = 3 and CDOREGETA <> '03' and CDOREGCRD = 'D' and substring(CDOREGCAA , 1 , 3) = 'PDI'" _
             & " and CDODOSCOP = DOSSLDOPE and CDODOSDOS = DOSSLDNUM" _
             & " and CDODOSOUV >= " & mCDODOSOUV_11012 _
            & " order by DOSSLDDEV,DOSSLDOPE,DOSSLDNUM"
    Else
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD0 , " & paramIBM_Library_SAB & ".ZCDOREG0 " _
             & " where DOSSLDPCI like '" & lDOSSLDPCI & "%' and DOSSLDCLI = '" & lDOSSLDCLI & "'" _
             & " and DOSSLDSTA not in ('  ','80','90')" _
             & " and CDOREGCOP = DOSSLDOPE and CDOREGDOS = DOSSLDNUM" _
             & " and CDOREGPAI = 3 and CDOREGETA <> '03' and CDOREGCRD = 'D' and substring(CDOREGCAA , 1 , 3) = 'PDI'" _
             & " order by DOSSLDDEV,DOSSLDOPE,DOSSLDNUM"
    End If
End If

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF

    V = rsYDOSSLD0_GetBuffer(rsSab, xYDOSSLD0)
        
    If xYDOSSLD0.DOSSLDDEV <> oldYDOSSLD0.DOSSLDDEV Then
        Call cmdSelect_SQL_Xi_Dossier_T(lCLIENARA1, lDOSSLDPCI)
    End If
    mXls2_Row = mXls2_Row + 1
    DAmjD = rsSab("CDOREGDEN") + 19000000
    DAmjF = rsSab("CDOREGDRE") + 19000000
    xCur = -rsSab("CDOREGMON")
    
    wsExcel.Cells(mXls2_Row, 1) = xYDOSSLD0.DOSSLDOPE
    wsExcel.Cells(mXls2_Row, 2) = xYDOSSLD0.DOSSLDNUM
    'wsExcel.Cells(mXls2_Row, 3) = rsSab("CDODOSCON")
    wsExcel.Cells(mXls2_Row, 4) = xYDOSSLD0.DOSSLDCLI
    wsExcel.Cells(mXls2_Row, 5) = dateImp10(DAmjD)
    wsExcel.Cells(mXls2_Row, 6) = dateImp10(DAmjF)
    wsExcel.Cells(mXls2_Row, 7) = xCur
    wsExcel.Cells(mXls2_Row, 8) = rsSab("CDOREGDEV")
    '//////////////////////////////////
    If rsSab("CDOREGDEV") <> "EUR" Then
        wsExcel.Cells(mXls2_Row, 9) = CDbl(xCur) / arrDev_Cours(retourne_indice_devise(rsSab("CDOREGDEV")))
    Else
        wsExcel.Cells(mXls2_Row, 9) = xCur
    End If
    '//////////////////////////////////
    If DAmjF <= wAmjMin Then
        wsExcel.Cells(mXls2_Row, 6).Font.Color = vbMagenta
        Nb1 = DateDiff("d", dateImp_Amj(DAmjD), wDMS_Min) + 1
    Else
        Nb1 = DateDiff("d", dateImp_Amj(DAmjD), dateImp_Amj(DAmjF)) + 1
    End If
    wsExcel.Cells(mXls2_Row, 10) = Nb1 'nb jours ////////////////////////////////////////////////////////
    If Nb1 >= 93 Then
        If rsSab("CDOREGDEV") <> "EUR" Then
            wsExcel.Cells(mXls2_Row, 11) = xCur
            wsExcel.Cells(mXls2_Row, 12) = CDbl(xCur) / arrDev_Cours(retourne_indice_devise(rsSab("CDOREGDEV")))
        Else
            wsExcel.Cells(mXls2_Row, 11) = xCur
            wsExcel.Cells(mXls2_Row, 12) = xCur
        End If
    Else
        If rsSab("CDOREGDEV") <> "EUR" Then
            wsExcel.Cells(mXls2_Row, 13) = xCur
            wsExcel.Cells(mXls2_Row, 14) = CDbl(xCur) / arrDev_Cours(retourne_indice_devise(rsSab("CDOREGDEV")))
        Else
            wsExcel.Cells(mXls2_Row, 13) = xCur
            wsExcel.Cells(mXls2_Row, 14) = xCur
        End If
    End If
    rsSab.MoveNext
Loop

xYDOSSLD0.DOSSLDDEV = ""

Call cmdSelect_SQL_Xi_Dossier_T(lCLIENARA1, lDOSSLDPCI)

Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & lSheet & "-" & lDOSSLDCLI): DoEvents


'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & lDOSSLDCLI): DoEvents

End Sub
Public Sub cmdSelect_SQL_Xi_ZCAUDOS0(lSheet As Integer, lDOSSLDCLI As String, lDOSSLDPCI As String, lCLIENARA1 As String)

On Error GoTo Error_Handler
Dim xSql As String
Dim X As String, K As Integer
Dim DAmjD As Long, DAmjF As Long, Nb1 As Long, Nb2 As Long, xCur As Currency
Dim DAmjX As Long

'__________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "> Exportation ........ " & lSheet & "-" & lDOSSLDCLI): DoEvents

Set wsExcel = wbExcel.Sheets(lSheet)

xYDOSSLD0.DOSSLDCLI = lDOSSLDCLI
oldYDOSSLD0.DOSSLDDEV = ""

If lDOSSLDCLI = "0011001" Then
    xSql = "select * from " & paramIBM_Library_SAB & ".ZCAUDOS0 " _
         & " where CAUDOSCAU = 'ESSPLC'" _
         & " and CAUDOSTRA in (2,3,5)" _
         & " and CAUDOSTIE = '" & lDOSSLDCLI & "'" _
         & " and CAUDOSDEB >= " & mCDODOSOUV_11001 _
         & " order by CAUDOSDEV,CAUDOSDOS"
Else
    If InStr(LFB_RacinesExclues, lDOSSLDCLI) > 0 Then
        xSql = "select * from " & paramIBM_Library_SAB & ".ZCAUDOS0 " _
             & " where CAUDOSCAU = 'ESSPLC'" _
             & " and CAUDOSTRA in (2,3,5)" _
             & " and CAUDOSTIE = '" & lDOSSLDCLI & "'" _
             & " and CAUDOSDEB >= " & mCDODOSOUV_11012 _
             & " order by CAUDOSDEV,CAUDOSDOS"
    Else
        xSql = "select * from " & paramIBM_Library_SAB & ".ZCAUDOS0 " _
             & " where CAUDOSCAU = 'ESSPLC'" _
             & " and CAUDOSTRA in (2,3,5)" _
             & " and CAUDOSTIE = '" & lDOSSLDCLI & "'" _
             & " order by CAUDOSDEV,CAUDOSDOS"
    End If
End If

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF

    xYDOSSLD0.DOSSLDDEV = rsSab("CAUDOSDEV")
    xYDOSSLD0.DOSSLDGSD = -rsSab("CAUDOSMNT")
    xYDOSSLD0.DOSSLDNUM = rsSab("CAUDOSDOS")
    
    If xYDOSSLD0.DOSSLDDEV <> oldYDOSSLD0.DOSSLDDEV Then
        Call cmdSelect_SQL_Xi_Dossier_T(lCLIENARA1, lDOSSLDPCI)
    End If
    
    DAmjD = rsSab("CAUDOSDEB") + 19000000
    DAmjF = rsSab("CAUDOSFIN") + 19000000
'===================================================================================
    xSql = "select * from " & paramIBM_Library_SAB & ".ZCAUAMO0 " _
         & " where CAUAMODOS = " & xYDOSSLD0.DOSSLDNUM _
         & " order by CAUAMODAT"
    Set rsSabX = cnsab.Execute(xSql)
    
    Do While Not rsSabX.EOF
        DAmjX = rsSabX("CAUAMODAT") + 19000000
        xCur = -rsSabX("CAUAMOMON")
        xYDOSSLD0.DOSSLDGSD = xYDOSSLD0.DOSSLDGSD - xCur
        If DAmjX > wAmjMin Then
                mXls2_Row = mXls2_Row + 1
                
                wsExcel.Cells(mXls2_Row, 1) = "CAU"
                wsExcel.Cells(mXls2_Row, 2) = xYDOSSLD0.DOSSLDNUM
                wsExcel.Cells(mXls2_Row, 4) = xYDOSSLD0.DOSSLDCLI
                wsExcel.Cells(mXls2_Row, 5) = dateImp10(DAmjD)
                wsExcel.Cells(mXls2_Row, 6) = dateImp10(DAmjX)
                wsExcel.Cells(mXls2_Row, 7) = xCur
                wsExcel.Cells(mXls2_Row, 8) = xYDOSSLD0.DOSSLDDEV
                '//////////////////////////////////
                If xYDOSSLD0.DOSSLDDEV <> "EUR" Then
                    wsExcel.Cells(mXls2_Row, 9) = CDbl(xCur) / arrDev_Cours(retourne_indice_devise(xYDOSSLD0.DOSSLDDEV))
                Else
                    wsExcel.Cells(mXls2_Row, 9) = xCur
                End If
                '//////////////////////////////////
                Nb1 = DateDiff("d", dateImp_Amj(DAmjD), dateImp_Amj(DAmjX)) + 1
                wsExcel.Cells(mXls2_Row, 10) = Nb1
                If Nb1 >= 93 Then
                    If xYDOSSLD0.DOSSLDDEV <> "EUR" Then
                        wsExcel.Cells(mXls2_Row, 11) = xYDOSSLD0.DOSSLDGSD
                        wsExcel.Cells(mXls2_Row, 12) = CDbl(xCur) / arrDev_Cours(retourne_indice_devise(xYDOSSLD0.DOSSLDDEV))
                    Else
                        wsExcel.Cells(mXls2_Row, 11) = xYDOSSLD0.DOSSLDGSD
                        wsExcel.Cells(mXls2_Row, 12) = xYDOSSLD0.DOSSLDGSD
                    End If
                Else
                    If xYDOSSLD0.DOSSLDDEV <> "EUR" Then
                        wsExcel.Cells(mXls2_Row, 13) = xYDOSSLD0.DOSSLDGSD
                        wsExcel.Cells(mXls2_Row, 14) = CDbl(xCur) / arrDev_Cours(retourne_indice_devise(xYDOSSLD0.DOSSLDDEV))
                    Else
                        wsExcel.Cells(mXls2_Row, 13) = xYDOSSLD0.DOSSLDGSD
                        wsExcel.Cells(mXls2_Row, 14) = xYDOSSLD0.DOSSLDGSD
                    End If
                End If
                    
        End If
            
        rsSabX.MoveNext
    Loop
    
'===================================================================================
    mXls2_Row = mXls2_Row + 1
    
    wsExcel.Cells(mXls2_Row, 1) = "CAU"
    wsExcel.Cells(mXls2_Row, 2) = xYDOSSLD0.DOSSLDNUM
    wsExcel.Cells(mXls2_Row, 4) = xYDOSSLD0.DOSSLDCLI
    wsExcel.Cells(mXls2_Row, 5) = dateImp10(DAmjD)
    wsExcel.Cells(mXls2_Row, 6) = dateImp10(DAmjF)
    wsExcel.Cells(mXls2_Row, 7) = xYDOSSLD0.DOSSLDGSD
    wsExcel.Cells(mXls2_Row, 8) = xYDOSSLD0.DOSSLDDEV
    '//////////////////////////////////
    If xYDOSSLD0.DOSSLDDEV <> "EUR" Then
        wsExcel.Cells(mXls2_Row, 9) = CDbl(xYDOSSLD0.DOSSLDGSD) / arrDev_Cours(retourne_indice_devise(xYDOSSLD0.DOSSLDDEV))
    Else
        wsExcel.Cells(mXls2_Row, 9) = xYDOSSLD0.DOSSLDGSD
    End If
    '//////////////////////////////////
    If DAmjF <= wAmjMin Then
        wsExcel.Cells(mXls2_Row, 6).Font.Color = vbMagenta
        Nb1 = DateDiff("d", dateImp_Amj(DAmjD), wDMS_Min) + 1
    Else
        Nb1 = DateDiff("d", dateImp_Amj(DAmjD), dateImp_Amj(DAmjF)) + 1
    End If
    wsExcel.Cells(mXls2_Row, 10) = Nb1
    If Nb1 >= 93 Then
        If xYDOSSLD0.DOSSLDDEV <> "EUR" Then
            wsExcel.Cells(mXls2_Row, 11) = xYDOSSLD0.DOSSLDGSD
            wsExcel.Cells(mXls2_Row, 12) = CDbl(xYDOSSLD0.DOSSLDGSD) / arrDev_Cours(retourne_indice_devise(xYDOSSLD0.DOSSLDDEV))
        Else
            wsExcel.Cells(mXls2_Row, 11) = xYDOSSLD0.DOSSLDGSD
            wsExcel.Cells(mXls2_Row, 12) = xYDOSSLD0.DOSSLDGSD
        End If
    Else
        If xYDOSSLD0.DOSSLDDEV <> "EUR" Then
            wsExcel.Cells(mXls2_Row, 13) = xYDOSSLD0.DOSSLDGSD
            wsExcel.Cells(mXls2_Row, 14) = CDbl(xYDOSSLD0.DOSSLDGSD) / arrDev_Cours(retourne_indice_devise(xYDOSSLD0.DOSSLDDEV))
        Else
            wsExcel.Cells(mXls2_Row, 13) = xYDOSSLD0.DOSSLDGSD
            wsExcel.Cells(mXls2_Row, 14) = xYDOSSLD0.DOSSLDGSD
        End If
    End If
    rsSab.MoveNext
Loop

xYDOSSLD0.DOSSLDDEV = ""

Call cmdSelect_SQL_Xi_Dossier_T(lCLIENARA1, lDOSSLDPCI)

Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & lSheet & "-" & lDOSSLDCLI): DoEvents


'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & lDOSSLDCLI): DoEvents

End Sub
Public Sub cmdSelect_SQL_Xi_Dossier_Init(lSheet As Integer)

On Error GoTo Error_Handler
Dim xSql As String
Dim X As String, K As Integer
Dim DAmjD As Long, DAmjF As Long, Nb1 As Long, Nb2 As Long, xCur As Currency
'__________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "> Exportation ........ " & lSheet): DoEvents

Set wsExcel = wbExcel.Sheets(lSheet)

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 85
wsExcel.PageSetup.PrintTitleRows = "$A1:$J1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
Select Case lSheet
    Case 4
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Liste des engagements INTRAGROUPE, arrêté au " & dateImp10(wAmjMin) _
                                        & vbCr & "&B&U&10(crédits documentaires export paiements différés - PCI 91120)"
    Case 5
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Liste des engagements INTRAGROUPE, arrêté au " & dateImp10(wAmjMin) _
                                        & vbCr & "&B&U&10(crédits documentaires export paiements différés - PCI 91130)"
    Case 6
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Liste des engagements INTRAGROUPE, arrêté au " & dateImp10(wAmjMin) _
                                        & vbCr & "&B&U&10(crédits documentaires export confirmés - PCI 91121)"
    Case 8
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Liste des engagements INTRAGROUPE, arrêté au " & dateImp10(wAmjMin) _
                                        & vbCr & "&B&U&10(crédits documentaires export échus)"
    Case 10
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Liste des engagements INTRAGROUPE, arrêté au " & dateImp10(wAmjMin) _
                                        & vbCr & "&B&U&10(crédits documentaires export échéance sup 31/12/2018)"
    Case 11
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Liste des dossiers réouverts depuis 1 mois." & " Le " & dateImp10(wAmjMin) _
                                        & vbCr & "&B&U&10(crédits documentaires dossiers réouverts)"
End Select

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignRight
    .WrapText = True
    .Font.Size = 9
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

If lSheet = 11 Then
    wsExcel.Columns(1).ColumnWidth = 7: wsExcel.Cells(1, 1) = "Dossier": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
    wsExcel.Columns(2).ColumnWidth = 7: wsExcel.Cells(1, 2) = "Client": wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignCenter
    wsExcel.Columns(3).ColumnWidth = 7: wsExcel.Cells(1, 3) = "Donneur ordre": wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignCenter
    wsExcel.Columns(4).ColumnWidth = 7: wsExcel.Cells(1, 4) = "Bénéficiaire": wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignCenter
    wsExcel.Columns(5).ColumnWidth = 7: wsExcel.Cells(1, 5) = "Ouvert le": wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignCenter
    wsExcel.Columns(6).ColumnWidth = 7: wsExcel.Cells(1, 6) = "Clos le": wsExcel.Columns(6).HorizontalAlignment = Excel.xlHAlignCenter
    wsExcel.Columns(7).ColumnWidth = 7: wsExcel.Cells(1, 7) = "Réouvert le": wsExcel.Columns(7).HorizontalAlignment = Excel.xlHAlignCenter
    mXls2_Col = 7
Else
    wsExcel.Columns(1).ColumnWidth = 5: wsExcel.Cells(1, 1) = "Code": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
    wsExcel.Columns(2).ColumnWidth = 7: wsExcel.Cells(1, 2) = "Dossier"
    wsExcel.Columns(3).ColumnWidth = 4: wsExcel.Cells(1, 3) = "C/N": wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignCenter
    wsExcel.Columns(4).ColumnWidth = 7: wsExcel.Cells(1, 4) = "Client": wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignCenter
    wsExcel.Columns(5).ColumnWidth = 10: wsExcel.Cells(1, 5) = "Validité du": wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignCenter
    wsExcel.Columns(6).ColumnWidth = 10: wsExcel.Cells(1, 6) = "au": wsExcel.Columns(6).HorizontalAlignment = Excel.xlHAlignCenter
    If lSheet = 5 Then
        wsExcel.Columns(5).ColumnWidth = 10: wsExcel.Cells(1, 5) = "Date utilisation": wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignCenter
        wsExcel.Columns(6).ColumnWidth = 10: wsExcel.Cells(1, 6) = "Date paiement": wsExcel.Columns(6).HorizontalAlignment = Excel.xlHAlignCenter
    End If
    wsExcel.Columns(7).ColumnWidth = 14: wsExcel.Cells(1, 7) = "Solde GESTION": wsExcel.Columns(7).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
    wsExcel.Columns(8).ColumnWidth = 5: wsExcel.Cells(1, 8) = "Devise": wsExcel.Columns(8).HorizontalAlignment = Excel.xlHAlignCenter
    wsExcel.Columns(9).ColumnWidth = 14: wsExcel.Cells(1, 9) = "cv Euro": wsExcel.Columns(9).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00": wsExcel.Cells(1, 9).HorizontalAlignment = Excel.xlHAlignCenter
    
    wsExcel.Columns(10).ColumnWidth = 7: wsExcel.Cells(1, 10) = "Nb J": wsExcel.Columns(10).NumberFormat = "###0"
    wsExcel.Columns(11).ColumnWidth = 14: wsExcel.Cells(1, 11) = "Solde >= 93 J ": wsExcel.Columns(11).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
    wsExcel.Columns(12).ColumnWidth = 14: wsExcel.Cells(1, 12) = "Solde Euro >= 93 J ": wsExcel.Columns(12).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
    wsExcel.Columns(13).ColumnWidth = 14: wsExcel.Cells(1, 13) = "Solde < 93 J ": wsExcel.Columns(13).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
    wsExcel.Columns(14).ColumnWidth = 14: wsExcel.Cells(1, 14) = "Solde Euro < 93 J ": wsExcel.Columns(14).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
    wsExcel.Columns(15).ColumnWidth = 30: wsExcel.Cells(1, 15) = "Intitulé": wsExcel.Columns(15).HorizontalAlignment = Excel.xlHAlignLeft
    wsExcel.Columns(16).ColumnWidth = 14: wsExcel.Cells(1, 16) = "Solde COMPTABLE": wsExcel.Columns(16).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
    
    mXls2_Col = 16
End If

For K = 1 To mXls2_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB ' RGB(255, 128, 50)
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next K

mXls2_Row = 1
mDev_R1 = 2


'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub









Public Function retourne_dernligne(FExcel As Excel.Worksheet, laColonne As Long) As Long
Dim numLig As Long

    For numLig = 2 To 65536
        If Trim(FExcel.Cells(numLig, laColonne)) = "" Then
            retourne_dernligne = numLig
            Exit Function
        End If
    Next numLig
    
End Function


Public Function retourne_indice_devise(dev As String) As Long

    retourne_indice_devise = -1
    For I = 1 To arrDev_Nb
        If UCase(dev) = UCase(arrDev(I)) Then
            retourne_indice_devise = I
            Exit Function
        End If
    Next I
    
End Function

Private Function retourne_Nom_Beneficiaire_Donneur(zNum As String) As String
Dim X As String
Dim rs As ADODB.Recordset

    retourne_Nom_Beneficiaire_Donneur = zNum
    X = "select CDOTIERA1 from " & paramIBM_Library_SAB & ".ZCDOTIE0" _
        & " where CDOTIETIE = '" & zNum & "'"
    Set rs = cnsab.Execute(X)
    Do While Not rs.EOF
        retourne_Nom_Beneficiaire_Donneur = CStr(zNum) & " " & Trim(rs("CDOTIERA1"))
        Exit Do
    Loop
    rs.Close
    Set rs = Nothing

End Function

Public Sub YDOSSLD0_Export_CDO()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wAmjMin As String, wAmjMax As String
Dim X As String, K As Long
Dim blnCALCS As Boolean

On Error GoTo Error_Handler
'===================================================================================
If blnAuto Then
    X = paramServer("\\CDO_Archive\")
Else
    X = ""
End If
If X = "" Then X = "C:\Temp\"
If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"


blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

wFile = X & Trim("CDO Surveillance " & dateImp_Amj(YBIATAB0_DATE_CPT_J) & ".xlsx")
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut : " & wFile _
        & vbCrLf & vbCrLf & "     =========================" _
        & vbCrLf & "     =========================", "Etat de surveillance Crédits Documentaires : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If
'_________________________________________

If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile

'If Dir(wFile) <> "" Then Kill wFile

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "DOSCDO"
    .Subject = ""
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "Detail " & dateImp10(YBIATAB0_DATE_CPT_J)


YDOSSLD0_Export_Detail_CDO

'__________________________________________________________________________________
'__________________________________________________________________________________
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
Call lstErr_AddItem(lstErr, cmdContext, "< CDO Exportation terminée"): DoEvents


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

Public Sub YDOSSLD0_Export_RDO()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wAmjMin As String, wAmjMax As String
Dim X As String, K As Long
Dim blnCALCS As Boolean

On Error GoTo Error_Handler
'===================================================================================
If blnAuto Then
    X = paramServer("\\CDO_Archive\")
Else
    X = ""
End If
If X = "" Then X = "C:\Temp\"
If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"


blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

wFile = X & Trim("RDO Surveillance " & dateImp_Amj(YBIATAB0_DATE_CPT_J) & ".xlsx")
'______________________________________________
'If Not blnAuto Then
'    X = InputBox("par défaut : " & wFile _
'        & vbCrLf & vbCrLf & "     =========================" _
'        & vbCrLf & "     =========================", "Etat de surveillance Remises Documentaires : nom du fichier d'exportation", wFile)
'    If Trim(X) = "" Then Exit Sub
    
'    wFilex = Trim(X)
    '______________________________________________
'    If wFile <> wFilex Then
'        wFile = wFilex
'    End If
'End If
'_________________________________________

If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile

'If Dir(wFile) <> "" Then Kill wFile

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "DOSRDO"
    .Subject = ""
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "Detail " & dateImp10(YBIATAB0_DATE_CPT_J)


YDOSSLD0_Export_Detail_RDO

'__________________________________________________________________________________
'__________________________________________________________________________________
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
Call lstErr_AddItem(lstErr, cmdContext, "< RDO Exportation terminée"): DoEvents

End Sub
Public Sub YDOSSLD0_Export_CAU()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wAmjMin As String, wAmjMax As String
Dim X As String, K As Long
Dim blnCALCS As Boolean

On Error GoTo Error_Handler
'===================================================================================
If blnAuto Then
    X = paramServer("\\CDO_Archive\")
Else
    X = ""
End If
If X = "" Then X = "C:\Temp\"
If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"


blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

If blnZCAUDOS0_S01 Then
    wFile = X & Trim("ENG-GAR Surveillance GDMP" & dateImp_Amj(YBIATAB0_DATE_CPT_J) & ".xlsx")
Else
    wFile = X & Trim("ENG-GAR Surveillance " & dateImp_Amj(YBIATAB0_DATE_CPT_J) & ".xlsx")
End If
'______________________________________________
'If Not blnAuto Then
'    X = InputBox("par défaut : " & wFile _
'        & vbCrLf & vbCrLf & "     =========================" _
'        & vbCrLf & "     =========================", "Etat de surveillance Remises Documentaires : nom du fichier d'exportation", wFile)
'    If Trim(X) = "" Then Exit Sub
    
'    wFilex = Trim(X)
    '______________________________________________
'    If wFile <> wFilex Then
'        wFile = wFilex
'    End If
'End If
'_________________________________________

If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile

'If Dir(wFile) <> "" Then Kill wFile

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "DOSCAU"
    .Subject = ""
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "Detail " & dateImp10(YBIATAB0_DATE_CPT_J)


YDOSSLD0_Export_Detail_CAU

'__________________________________________________________________________________
'__________________________________________________________________________________
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
Call lstErr_AddItem(lstErr, cmdContext, "< ENG-GAR Exportation terminée"): DoEvents


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


Public Sub cmdSendMail_SAB_Dossier(lCC As String, lDocName As String)
Dim wSendMail As typeSendMail
Dim K As Long, htmlFontColor_K As String

On Error Resume Next

'____________________________________________________________________________________________
paramEditionNoPaper_Auto_PgmName = "BIA-SAB-DOSSIER"
If blnZCAUDOS0_S01 Then
    wSendMail.FromDisplayName = "S01"
    wSendMail.RecipientDisplayName = "NoPaper"
    Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S01", wFile, "Archive", lDocName)
Else
    If lDocName = "BIA-CDO-Engagement1an" Then
        wSendMail.FromDisplayName = "S51"
        wSendMail.RecipientDisplayName = "NoPaper"
        Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S51", wFile, "Archive", lDocName)
        '11/02/2020 ajout envoie à S60 = Compta
        wSendMail.FromDisplayName = "S60"
        wSendMail.RecipientDisplayName = "NoPaper"
        Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S60", wFile, "Archive", lDocName)
    Else
        wSendMail.FromDisplayName = "@SAB_DOSSIER"
        wSendMail.RecipientDisplayName = "CDO"
        Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S10", wFile, "Archive", lDocName)
    End If
End If
wSendMail.CcDisplayName = lCC

wSendMail.Subject = lDocName & " du : " & dateImp10(YBIATAB0_DATE_CPT_J)
'wSendMail.Subject = wFile & " (cf. pièce jointe)"
wSendMail.Attachment = "" 'wFile
'wSendMail.Message = "<body bgcolor = #FFFFFF><BR>"

wSendMail.Message = mHtml_Head & "<span style='font-size:10.0pt;font-family:Calibri'>" _
                 & htmlFontColor_Black & "BIA-SAB-DOSSIER" & "<BR><BR>" & paramEditionNoPaper_Auto_Lnk & "</div></body></html>"

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub
Public Sub cmdSelect_SQL_2_Exportation_Xlsx()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFilex As String, wFile As String, xSql As String
Dim wAmjMin As String, wAmjMax As String
Dim X As String, I As Integer
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, kLen As Integer
Dim K5 As Integer, K6 As Integer, K7 As Integer, K8 As Integer, K9 As Integer
Dim wMTD As Currency, wDTR As Long, wPIE As Long, wECR As Long
Dim wOPE As String, wNum As Long, wEVE As String, wANU As String, wK As Integer, wAmj As Long
Dim oldOPE As String, oldNUM As Long, oldMTD As Currency, totalMTD As Currency, gestionMTD As Currency
Dim repMTD As Currency, repGMTD As Currency
Dim wSTA As String, wSVC As String
'______________________________________________

wFile = Trim("C:\Temp\CDO Compte " & Trim(oldYBIACPT0.COMPTECOM) & ".xlsx")
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut : " & wFile _
        & vbCrLf & vbCrLf & "     =========================" _
        & vbCrLf & "     =========================", "Crédits Documentaires : Liste de rapprochement compta / gestion  : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If
'_________________________________________


If Dir(wFile) <> "" Then Kill wFile

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "DOSCDO"
    .Subject = ""
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = oldYBIACPT0.COMPTECOM

'__________________________________________________________________________________
wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 95
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14CDO : Liste de rapprochement compta / gestion " & Trim(oldYBIACPT0.COMPTECOM) & " en date du " & dateImp10_S(YBIATAB0_DATE_CPT_J) _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$G1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .WrapText = True
    .Font.Size = 8
    .Font.Name = "Arial Unicode MS"
    .RowHeight = 17
End With

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

wsExcel.Columns(1).ColumnWidth = 6: wsExcel.Cells(1, 1) = "Code": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 7: wsExcel.Cells(1, 2) = "N° dossier": wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(3).ColumnWidth = 9: wsExcel.Cells(1, 3) = "Date": wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(4).ColumnWidth = 6: wsExcel.Cells(1, 4) = "Evé": wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(5).ColumnWidth = 12: wsExcel.Cells(1, 5) = "Montant": wsExcel.Columns(5).NumberFormat = "### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(6).ColumnWidth = 15: wsExcel.Cells(1, 6) = "Gestion": wsExcel.Columns(6).NumberFormat = "### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(7).ColumnWidth = 30: wsExcel.Cells(1, 7) = "Libellé"


For K = 1 To 7
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 170, 80)
Next K

oldOPE = "": oldNUM = 0: oldMTD = 0: totalMTD = 0: gestionMTD = 0
repMTD = 0: repGMTD = 0
wRow = 1
For I = 0 To lstW.ListCount - 1
    lstW.ListIndex = I
    X = Trim(lstW.Text): kLen = Len(X)
    K1 = InStr(1, X, "|"): wOPE = Mid$(X, 1, K1 - 1)
    K2 = InStr(K1 + 1, X, "|"): wNum = Mid$(X, K1 + 1, K2 - K1 - 1)
    K3 = InStr(K2 + 1, X, "|"): wK = Mid$(X, K2 + 1, K3 - K2 - 1)
    K4 = InStr(K3 + 1, X, "|"): wAmj = Mid$(X, K3 + 1, K4 - K3 - 1)
    K5 = InStr(K4 + 1, X, "|"): wEVE = Mid$(X, K4 + 1, K5 - K4 - 1)
    K6 = InStr(K5 + 1, X, "|"): wMTD = Mid$(X, K5 + 1, K6 - K5 - 1)
    K7 = InStr(K6 + 1, X, "|"): wANU = Mid$(X, K6 + 1, K7 - K6 - 1)
    K8 = InStr(K7 + 1, X, "|"): wSTA = Mid$(X, K7 + 1, K8 - K7 - 1)
    K9 = InStr(K8 + 1, X, "|"): wSVC = Mid$(X, K8 + 1, K9 - K8 - 1)
    If wNum = 905015 Then
        Debug.Print 905015
    End If
'-----------------------------------------------------------------------------------
    If wOPE = "***" Then
        oldOPE = wOPE
        If wK = 9 Then
            repGMTD = repGMTD + wMTD
            gestionMTD = gestionMTD + wMTD
            wRow = wRow + 1
            wsExcel.Cells(wRow, 6) = wMTD
            If wMTD = 0 Then
                wsExcel.Cells(wRow, 6).Interior.Color = RGB(230, 230, 230)
                wsExcel.Cells(wRow, 7).Interior.Color = RGB(230, 230, 230)
            Else
                 wsExcel.Cells(wRow, 6).Interior.Color = RGB(255, 170, 80)
                 wsExcel.Cells(wRow, 7).Interior.Color = RGB(255, 170, 80)
           End If
            
            wsExcel.Cells(wRow, 7) = "Gestion " & wOPE & " " & wNum & " (" & wSTA & "-" & wSVC & ")"
        Else
            repMTD = repMTD + wMTD
            wRow = wRow + 1
            totalMTD = totalMTD + wMTD
            wsExcel.Cells(wRow, 1) = wOPE
            wsExcel.Cells(wRow, 2) = wNum
            wsExcel.Cells(wRow, 3) = dateImp10(wAmj)
            wsExcel.Cells(wRow, 4) = wEVE
            wsExcel.Cells(wRow, 5) = wMTD
            If wANU <> "0" Then wsExcel.Cells(wRow, 6) = wANU
            wsExcel.Cells(wRow, 7) = Mid$(X, K9 + 1, kLen - K9)  '- 1)

        End If
    Else
'-----------------------------------------------------------------------------------

        If wK = 9 Then
            wRow = wRow + 1
            wsExcel.Cells(wRow, 5) = oldMTD
            wsExcel.Cells(wRow, 6) = wMTD
            wsExcel.Cells(wRow, 7) = "Gestion " & wOPE & " " & wNum & " (" & wSTA & "-" & wSVC & ")"
            gestionMTD = gestionMTD + wMTD
            If oldMTD = wMTD Then
                wsExcel.Cells(wRow, 4) = "C=G"
                wsExcel.Cells(wRow, 5).Interior.Color = RGB(192, 255, 192)
                wsExcel.Cells(wRow, 6).Interior.Color = RGB(192, 255, 192)
            Else
                wsExcel.Cells(wRow, 4) = "C#G"
                wsExcel.Cells(wRow, 5).Interior.Color = RGB(255, 192, 192)
                wsExcel.Cells(wRow, 6).Interior.Color = RGB(255, 192, 192)
            End If
            oldOPE = "": oldNUM = 0: oldMTD = 0
        Else
            If oldOPE <> wOPE Or oldNUM <> wNum Then
                 If oldOPE = "***" Then
                    oldNUM = 0
                    wRow = wRow + 1
                    wsExcel.Cells(wRow, 5) = repMTD
                    wsExcel.Cells(wRow, 6) = repGMTD
                    wsExcel.Cells(wRow, 7) = "Total reprise TI"
                    wsExcel.Cells(wRow, 7).Interior.Color = RGB(255, 170, 80)
                    If repMTD = repGMTD Then
                        wsExcel.Cells(wRow, 4) = "C=G"
                        wsExcel.Cells(wRow, 5).Interior.Color = RGB(192, 255, 192)
                        wsExcel.Cells(wRow, 6).Interior.Color = RGB(192, 255, 192)
                    Else
                        wsExcel.Cells(wRow, 4) = "C*G"
                        wsExcel.Cells(wRow, 5).Interior.Color = RGB(255, 192, 192)
                        wsExcel.Cells(wRow, 6).Interior.Color = RGB(255, 192, 192)
                    End If
                 End If
                 If oldNUM <> 0 Then
                    wRow = wRow + 1
                    wsExcel.Cells(wRow, 4) = "C?G"
                    wsExcel.Cells(wRow, 5) = oldMTD
                    wsExcel.Cells(wRow, 7) = "pas de dossier en gestion"
                    wsExcel.Cells(wRow, 5).Interior.Color = RGB(192, 255, 192)
                    wsExcel.Cells(wRow, 6).Interior.Color = RGB(255, 0, 0)
                 End If
            
                oldOPE = wOPE: oldNUM = wNum: oldMTD = 0
            
            End If
            wRow = wRow + 1
            oldMTD = oldMTD + wMTD
            totalMTD = totalMTD + wMTD
            wsExcel.Cells(wRow, 1) = wOPE
            wsExcel.Cells(wRow, 2) = wNum
            wsExcel.Cells(wRow, 3) = dateImp10(wAmj)
            wsExcel.Cells(wRow, 4) = wEVE
            wsExcel.Cells(wRow, 5) = wMTD
            If wANU <> "0" Then wsExcel.Cells(wRow, 6) = wANU
            wsExcel.Cells(wRow, 7) = Mid$(X, K9 + 1, kLen - K9)  '- 1)
        End If
    End If

Next I

'__________________________________________________________________________________

wRow = wRow + 1
For K = 1 To 7
    wsExcel.Cells(wRow, K).Interior.Color = RGB(255, 200, 120)
Next K
wsExcel.Cells(wRow, 5) = totalMTD
wsExcel.Cells(wRow, 5) = totalMTD
wsExcel.Cells(wRow, 1) = "Total"
wsExcel.Cells(wRow, 2) = "mvts"
wsExcel.Cells(wRow, 7) = "(total des soldes de gestion)"
wsExcel.Cells(wRow, 6) = gestionMTD
If gestionMTD = totalMTD Then
    wsExcel.Cells(wRow, 4) = "C=G"
    wsExcel.Cells(wRow, 5).Interior.Color = RGB(192, 255, 192)
    wsExcel.Cells(wRow, 6).Interior.Color = RGB(192, 255, 192)
Else
    wsExcel.Cells(wRow, 4) = "C#G"
    wsExcel.Cells(wRow, 6).Interior.Color = RGB(255, 192, 192)
    wsExcel.Cells(wRow, 5).Interior.Color = RGB(255, 192, 192)
End If


wRow = wRow + 1
For K = 1 To 7
    wsExcel.Cells(wRow, K).Interior.Color = RGB(255, 170, 80)
Next K

wsExcel.Cells(wRow, 1) = "Solde"
wsExcel.Cells(wRow, 2) = "compte"
wsExcel.Cells(wRow, 3) = dateImp10(YBIATAB0_DATE_CPT_J)
wsExcel.Cells(wRow, 6) = oldYBIACPT0.COMPTECOM
wsExcel.Cells(wRow, 7) = oldYBIACPT0.COMPTEINT
wsExcel.Cells(wRow, 5) = -oldYBIACPT0.SOLDECEN
If totalMTD = -oldYBIACPT0.SOLDECEN Then
    wsExcel.Cells(wRow, 4) = "C=S"
    wsExcel.Cells(wRow, 5).Interior.Color = RGB(192, 255, 192)
Else
    wsExcel.Cells(wRow, 4) = "C#S"
    wsExcel.Cells(wRow, 5).Interior.Color = RGB(255, 192, 192)
End If

'__________________________________________________________________________________

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
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents

End Sub

Public Sub YDOSSLD0_Export_Detail_CDO()
On Error GoTo Error_Handler
Dim X As String, XX As String
Dim wRow As Long, wCol As Long, kIndex As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSql As String
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, kLen As Integer
Dim K5 As Integer, K6 As Integer, K7 As Integer, K8 As Integer
Dim wDOSSLDOPEC As String, wDOSSLDOPEN As Long, wCLIEANRA1 As String
Dim wCLIEANCLI As String
Dim wCDODOSMOT As Currency, wCDODOSDEV As String, wCDODOSCON   As String, xCDODOSOUV As String
Dim wCli As String, wOPE As String, wDOS As Long, wErr As String, wLIB As String, wNAT As String
Dim wMTD As Currency, wDev As String, wAmj As String
Dim wDT As String
'______________________________________________

'__________________________________________________________________________________

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlthick
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .WrapText = True
    .Font.Size = 7
    .Font.Name = "Arial Unicode MS"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 80
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14CDO : Surveillance en date du " & dateImp10_S(YBIATAB0_DATE_CPT_J) _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$I1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


Call lstErr_AddItem(lstErr, cmdContext, "CDO Exportation en cours : "): DoEvents

wsExcel.Columns(1).ColumnWidth = 7: wsExcel.Cells(1, 1) = "Banque ": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 29: wsExcel.Cells(1, 2) = "Banque émettrice"
wsExcel.Columns(3).ColumnWidth = 6: wsExcel.Cells(1, 3) = "Code": wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(4).ColumnWidth = 7: wsExcel.Cells(1, 4) = "N° dossier": wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(5).ColumnWidth = 8: wsExcel.Cells(1, 5) = "Contrôle": wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(6).ColumnWidth = 65: wsExcel.Cells(1, 6) = "libellé surveillance"
wsExcel.Columns(7).ColumnWidth = 6: wsExcel.Cells(1, 7) = "Nature": wsExcel.Columns(7).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(8).ColumnWidth = 12: wsExcel.Cells(1, 8) = "Montant": wsExcel.Columns(8).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(8).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(9).ColumnWidth = 6: wsExcel.Cells(1, 9) = "Devise": wsExcel.Columns(9).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(10).ColumnWidth = 9: wsExcel.Cells(1, 10) = "Date": wsExcel.Columns(10).HorizontalAlignment = Excel.xlHAlignCenter

For K = 1 To 10
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 170, 80)
Next
wRow = 1
For kIndex = 0 To lstW.ListCount - 1
    wRow = wRow + 1
    lstW.ListIndex = kIndex
    X = Trim(lstW.Text): kLen = Len(X)
    K = InStr(1, X, "|"): wDT = Mid$(X, 1, K - 1)
    K1 = InStr(K + 1, X, "|"): wErr = Mid$(X, K + 1, K1 - K - 1)
    K2 = InStr(K1 + 1, X, "|"): wOPE = Mid$(X, K1 + 1, K2 - K1 - 1)
    K3 = InStr(K2 + 1, X, "|"): wDOS = Val(Mid$(X, K2 + 1, K3 - K2 - 1))
    K4 = InStr(K3 + 1, X, "|"): wCli = Mid$(X, K3 + 1, K4 - K3 - 1)
    K5 = InStr(K4 + 1, X, "|"): wLIB = Mid$(X, K4 + 1, K5 - K4 - 1)
    K6 = InStr(K5 + 1, X, "|"): wNAT = Mid$(X, K5 + 1, K6 - K5 - 1)
    K7 = InStr(K6 + 1, X, "|"): XX = Mid$(X, K6 + 1, K7 - K6 - 1)
        If Trim(XX) = "" Then
            wMTD = 0
        Else
            wMTD = CCur(XX)
        End If

    K8 = InStr(K7 + 1, X, "|"): wDev = Mid$(X, K7 + 1, K8 - K7 - 1)
    If kLen > K8 Then
        wAmj = Mid$(X, K8 + 1, kLen - K8)
    Else
        wAmj = ""
    End If
    Select Case Mid$(wDT, 3, 1)
        Case "S":
                'wsExcel.Rows(wrow).RowHeight = 30
                wsExcel.Cells(wRow, 2) = wCli: wsExcel.Cells(wRow, 2).Font.Bold = True
                wsExcel.Cells(wRow, 6) = wLIB: wsExcel.Cells(wRow, 6).Font.Bold = True
                For K = 1 To 10
                    wsExcel.Cells(wRow, K).Interior.Color = mColor_G0 'RGB(190, 255, 190)
                Next K
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
               ' wRow = wRow + 1
               ' wsExcel.Rows(wRow).RowHeight = 5
               ' For K = 1 To 10: wsExcel.Cells(wRow, K).Interior.Color = RGB(128, 128, 128): Next K

        Case "T":
                'wsExcel.Rows(wrow).RowHeight = 30
                wsExcel.Cells(wRow, 2) = wCli: wsExcel.Cells(wRow, 2).Font.Bold = True
                wsExcel.Cells(wRow, 6) = wLIB: wsExcel.Cells(wRow, 6).Font.Bold = True
                For K = 1 To 10
                    wsExcel.Cells(wRow, K).Interior.Color = mColor_W0 'RGB(255, 190, 190)
                Next K
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
                'wRow = wRow + 1
                'wsExcel.Rows(wRow).RowHeight = 5
                'For K = 1 To 10: wsExcel.Cells(wRow, K).Interior.Color = RGB(128, 128, 128): Next K
       Case "X":
                'wsExcel.Rows(wrow).RowHeight = 20
                wsExcel.Cells(wRow, 2) = wCli: wsExcel.Cells(wRow, 2).Font.Bold = True
                wsExcel.Cells(wRow, 6) = wLIB
                For K = 1 To 10
                    wsExcel.Cells(wRow, K).Interior.Color = mColor_Y0 'RGB(255, 220, 120)
                Next K
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
               ' wRow = wRow + 1
                'wsExcel.Rows(wRow).RowHeight = 5
                'For K = 1 To 10: wsExcel.Cells(wRow, K).Interior.Color = RGB(128, 128, 128): Next K
        Case Else
        
        xSql = "select  CLIENACLI , CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
                & " where CLIENACLI = '" & wCli & "'"
        Set rsSab = cnsab.Execute(xSql)
        If Not rsSab.EOF Then
            wCLIEANCLI = rsSab("CLIENACLI")
            wCLIEANRA1 = rsSab("CLIENARA1")
        Else
            wCLIEANCLI = ""
            wCLIEANRA1 = ""
        End If
        'kLen = Len(X)
        wsExcel.Cells(wRow, 1) = wCli
        wsExcel.Cells(wRow, 2) = wCLIEANRA1
        wsExcel.Cells(wRow, 3) = wOPE
        wsExcel.Cells(wRow, 4) = wDOS
        wsExcel.Cells(wRow, 5) = wErr
        wsExcel.Cells(wRow, 6) = wLIB
        wsExcel.Cells(wRow, 7) = wNAT
        wsExcel.Cells(wRow, 8) = wMTD
        wsExcel.Cells(wRow, 9) = wDev
        wsExcel.Cells(wRow, 10) = wAmj
    End Select
Next kIndex

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub YDOSSLD0_Export_Detail_RDO()
On Error GoTo Error_Handler
Dim X As String, XX As String
Dim wRow As Long, wCol As Long, kIndex As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSql As String
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, kLen As Integer
Dim K5 As Integer, K6 As Integer, K7 As Integer, K8 As Integer
Dim wDOSSLDOPEC As String, wDOSSLDOPEN As Long, wCLIEANRA1 As String
Dim wCLIEANCLI As String
Dim wCDODOSMOT As Currency, wCDODOSDEV As String, wCDODOSCON   As String, xCDODOSOUV As String
Dim wCli As String, wOPE As String, wDOS As Long, wErr As String, wLIB As String, wNAT As String
Dim wMTD As Currency, wDev As String, wAmj As String
Dim wDT As String
'______________________________________________

'__________________________________________________________________________________

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlthick
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .WrapText = True
    .Font.Size = 7
    .Font.Name = "Arial Unicode MS"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 80
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14RDO : Surveillance en date du " & dateImp10_S(YBIATAB0_DATE_CPT_J) _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$I1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


Call lstErr_AddItem(lstErr, cmdContext, "RDO Exportation en cours : "): DoEvents

wsExcel.Columns(1).ColumnWidth = 7: wsExcel.Cells(1, 1) = "Nature ": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 29: wsExcel.Cells(1, 2) = ""
wsExcel.Columns(3).ColumnWidth = 6: wsExcel.Cells(1, 3) = "Code": wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(4).ColumnWidth = 7: wsExcel.Cells(1, 4) = "N° dossier": wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(5).ColumnWidth = 8: wsExcel.Cells(1, 5) = "Contrôle": wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(6).ColumnWidth = 65: wsExcel.Cells(1, 6) = "libellé surveillance"
wsExcel.Columns(7).ColumnWidth = 6: wsExcel.Cells(1, 7) = "Evénement": wsExcel.Columns(7).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(8).ColumnWidth = 12: wsExcel.Cells(1, 8) = "Montant": wsExcel.Columns(8).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(8).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(9).ColumnWidth = 6: wsExcel.Cells(1, 9) = "Devise": wsExcel.Columns(9).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(10).ColumnWidth = 9: wsExcel.Cells(1, 10) = "Date": wsExcel.Columns(10).HorizontalAlignment = Excel.xlHAlignCenter

For K = 1 To 10
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 170, 80)
Next
wRow = 1
For kIndex = 0 To lstW.ListCount - 1
    wRow = wRow + 1
    lstW.ListIndex = kIndex
    X = Trim(lstW.Text): kLen = Len(X)
    K = InStr(1, X, "|"): wDT = Mid$(X, 1, K - 1)
    K1 = InStr(K + 1, X, "|"): wErr = Mid$(X, K + 1, K1 - K - 1)
    K2 = InStr(K1 + 1, X, "|"): wOPE = Mid$(X, K1 + 1, K2 - K1 - 1)
    K3 = InStr(K2 + 1, X, "|"): wDOS = Val(Mid$(X, K2 + 1, K3 - K2 - 1))
    K4 = InStr(K3 + 1, X, "|"): wCli = Mid$(X, K3 + 1, K4 - K3 - 1)
    K5 = InStr(K4 + 1, X, "|"): wLIB = Mid$(X, K4 + 1, K5 - K4 - 1)
    K6 = InStr(K5 + 1, X, "|"): wNAT = Mid$(X, K5 + 1, K6 - K5 - 1)
    K7 = InStr(K6 + 1, X, "|"): XX = Mid$(X, K6 + 1, K7 - K6 - 1)
        If Trim(XX) = "" Then
            wMTD = 0
        Else
            wMTD = CCur(XX)
        End If

    K8 = InStr(K7 + 1, X, "|"): wDev = Mid$(X, K7 + 1, K8 - K7 - 1)
    If kLen > K8 Then
        wAmj = Mid$(X, K8 + 1, kLen - K8)
    Else
        wAmj = ""
    End If
    Select Case Mid$(wDT, 3, 1)
        Case "S":
                'wsExcel.Rows(wrow).RowHeight = 30
                wsExcel.Cells(wRow, 2) = wCli: wsExcel.Cells(wRow, 2).Font.Bold = True
                wsExcel.Cells(wRow, 6) = wLIB: wsExcel.Cells(wRow, 6).Font.Bold = True
                For K = 1 To 10
                    wsExcel.Cells(wRow, K).Interior.Color = mColor_G0 'RGB(190, 255, 190)
                Next K
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
               ' wRow = wRow + 1
               ' wsExcel.Rows(wRow).RowHeight = 5
               ' For K = 1 To 10: wsExcel.Cells(wRow, K).Interior.Color = RGB(128, 128, 128): Next K

        Case "T":
                'wsExcel.Rows(wrow).RowHeight = 30
                wsExcel.Cells(wRow, 2) = wCli: wsExcel.Cells(wRow, 2).Font.Bold = True
                wsExcel.Cells(wRow, 6) = wLIB: wsExcel.Cells(wRow, 6).Font.Bold = True
                For K = 1 To 10
                    wsExcel.Cells(wRow, K).Interior.Color = mColor_W0 'RGB(255, 190, 190)
                Next K
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
                'wRow = wRow + 1
                'wsExcel.Rows(wRow).RowHeight = 5
                'For K = 1 To 10: wsExcel.Cells(wRow, K).Interior.Color = RGB(128, 128, 128): Next K
       Case "X":
                'wsExcel.Rows(wrow).RowHeight = 20
                wsExcel.Cells(wRow, 2) = wCli: wsExcel.Cells(wRow, 2).Font.Bold = True
                wsExcel.Cells(wRow, 6) = wLIB
                For K = 1 To 10
                    wsExcel.Cells(wRow, K).Interior.Color = mColor_Y0 'RGB(255, 220, 120)
                Next K
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
               ' wRow = wRow + 1
                'wsExcel.Rows(wRow).RowHeight = 5
                'For K = 1 To 10: wsExcel.Cells(wRow, K).Interior.Color = RGB(128, 128, 128): Next K
        Case Else
        
        xSql = "select  CLIENACLI , CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
                & " where CLIENACLI = '" & wCli & "'"
        Set rsSab = cnsab.Execute(xSql)
        If Not rsSab.EOF Then
            wCLIEANCLI = rsSab("CLIENACLI")
            wCLIEANRA1 = rsSab("CLIENARA1")
        Else
            wCLIEANCLI = ""
            wCLIEANRA1 = ""
        End If
        'kLen = Len(X)
        wsExcel.Cells(wRow, 1) = wCli
        wsExcel.Cells(wRow, 2) = wCLIEANRA1
        wsExcel.Cells(wRow, 3) = wOPE
        wsExcel.Cells(wRow, 4) = wDOS
        wsExcel.Cells(wRow, 5) = wErr
        wsExcel.Cells(wRow, 6) = wLIB
        wsExcel.Cells(wRow, 7) = wNAT
        wsExcel.Cells(wRow, 8) = wMTD
        wsExcel.Cells(wRow, 9) = wDev
        wsExcel.Cells(wRow, 10) = wAmj
    End Select
Next kIndex

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub



Public Sub YDOSSLD0_Export_Detail_CAU()
On Error GoTo Error_Handler
Dim X As String, XX As String
Dim wRow As Long, wCol As Long, kIndex As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSql As String
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, kLen As Integer
Dim K5 As Integer, K6 As Integer, K7 As Integer, K8 As Integer
Dim wDOSSLDOPEC As String, wDOSSLDOPEN As Long, wCLIEANRA1 As String
Dim wCLIEANCLI As String
Dim wCDODOSMOT As Currency, wCDODOSDEV As String, wCDODOSCON   As String, xCDODOSOUV As String
Dim wCli As String, wOPE As String, wDOS As Long, wErr As String, wLIB As String, wNAT As String
Dim wMTD As Currency, wDev As String, wAmj As String
Dim wDT As String
'______________________________________________
On Error GoTo Error_Handler

'__________________________________________________________________________________

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlthick
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .WrapText = True
    .Font.Size = 7
    .Font.Name = "Arial Unicode MS"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 80
If blnZCAUDOS0_S01 Then
    X = " (n° dossier < 500 000) "
Else
    X = " (n° dossier >= 500 000) "
End If

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14ENG - GAR " & X & ": Surveillance en date du " & dateImp10_S(YBIATAB0_DATE_CPT_J) _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$I1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


Call lstErr_AddItem(lstErr, cmdContext, "ENG- GAR Exportation en cours : "): DoEvents

wsExcel.Columns(1).ColumnWidth = 7: wsExcel.Cells(1, 1) = "Nature ": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 29: wsExcel.Cells(1, 2) = ""
wsExcel.Columns(3).ColumnWidth = 6: wsExcel.Cells(1, 3) = "Code": wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(4).ColumnWidth = 7: wsExcel.Cells(1, 4) = "N° dossier": wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(5).ColumnWidth = 8: wsExcel.Cells(1, 5) = "Contrôle": wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(6).ColumnWidth = 65: wsExcel.Cells(1, 6) = "libellé surveillance"
wsExcel.Columns(7).ColumnWidth = 6: wsExcel.Cells(1, 7) = "Evénement": wsExcel.Columns(7).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(8).ColumnWidth = 12: wsExcel.Cells(1, 8) = "Montant": wsExcel.Columns(8).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(8).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(9).ColumnWidth = 6: wsExcel.Cells(1, 9) = "Devise": wsExcel.Columns(9).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(10).ColumnWidth = 9: wsExcel.Cells(1, 10) = "Date": wsExcel.Columns(10).HorizontalAlignment = Excel.xlHAlignCenter

For K = 1 To 10
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 170, 80)
Next
wRow = 1
For kIndex = 0 To lstW.ListCount - 1
    wRow = wRow + 1
    lstW.ListIndex = kIndex
    X = Trim(lstW.Text): kLen = Len(X)
    K = InStr(1, X, "|"): wDT = Mid$(X, 1, K - 1)
    K1 = InStr(K + 1, X, "|"): wErr = Mid$(X, K + 1, K1 - K - 1)
    K2 = InStr(K1 + 1, X, "|"): wOPE = Mid$(X, K1 + 1, K2 - K1 - 1)
    K3 = InStr(K2 + 1, X, "|"): wDOS = Val(Mid$(X, K2 + 1, K3 - K2 - 1))
    K4 = InStr(K3 + 1, X, "|"): wCli = Mid$(X, K3 + 1, K4 - K3 - 1)
    K5 = InStr(K4 + 1, X, "|"): wLIB = Mid$(X, K4 + 1, K5 - K4 - 1)
    K6 = InStr(K5 + 1, X, "|"): wNAT = Mid$(X, K5 + 1, K6 - K5 - 1)
    K7 = InStr(K6 + 1, X, "|"): XX = Mid$(X, K6 + 1, K7 - K6 - 1)
        If Trim(XX) = "" Then
            wMTD = 0
        Else
            wMTD = CCur(XX)
        End If

    K8 = InStr(K7 + 1, X, "|"): wDev = Mid$(X, K7 + 1, K8 - K7 - 1)
    If kLen > K8 Then
        wAmj = Mid$(X, K8 + 1, kLen - K8)
    Else
        wAmj = ""
    End If
    Select Case Mid$(wDT, 3, 1)
        Case "S":
                'wsExcel.Rows(wrow).RowHeight = 30
                wsExcel.Cells(wRow, 2) = wCli: wsExcel.Cells(wRow, 2).Font.Bold = True
                wsExcel.Cells(wRow, 6) = wLIB: wsExcel.Cells(wRow, 6).Font.Bold = True
                For K = 1 To 10
                    wsExcel.Cells(wRow, K).Interior.Color = mColor_G0 'RGB(190, 255, 190)
                Next K
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
               ' wRow = wRow + 1
               ' wsExcel.Rows(wRow).RowHeight = 5
               ' For K = 1 To 10: wsExcel.Cells(wRow, K).Interior.Color = RGB(128, 128, 128): Next K

        Case "T":
                'wsExcel.Rows(wrow).RowHeight = 30
                wsExcel.Cells(wRow, 2) = wCli: wsExcel.Cells(wRow, 2).Font.Bold = True
                wsExcel.Cells(wRow, 6) = wLIB: wsExcel.Cells(wRow, 6).Font.Bold = True
                For K = 1 To 10
                    wsExcel.Cells(wRow, K).Interior.Color = mColor_W0 'RGB(255, 190, 190)
                Next K
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
                'wRow = wRow + 1
                'wsExcel.Rows(wRow).RowHeight = 5
                'For K = 1 To 10: wsExcel.Cells(wRow, K).Interior.Color = RGB(128, 128, 128): Next K
       Case "X":
                'wsExcel.Rows(wrow).RowHeight = 20
                wsExcel.Cells(wRow, 2) = wCli: wsExcel.Cells(wRow, 2).Font.Bold = True
                wsExcel.Cells(wRow, 6) = wLIB
                For K = 1 To 10
                    wsExcel.Cells(wRow, K).Interior.Color = mColor_Y0 'RGB(255, 220, 120)
                Next K
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Weight = xlThick
                wsExcel.Range("A" & wRow & ":J" & wRow).Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
               ' wRow = wRow + 1
                'wsExcel.Rows(wRow).RowHeight = 5
                'For K = 1 To 10: wsExcel.Cells(wRow, K).Interior.Color = RGB(128, 128, 128): Next K
        Case Else
        
        xSql = "select  CLIENACLI , CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
                & " where CLIENACLI = '" & wCli & "'"
        Set rsSab = cnsab.Execute(xSql)
        If Not rsSab.EOF Then
            wCLIEANCLI = rsSab("CLIENACLI")
            wCLIEANRA1 = rsSab("CLIENARA1")
        Else
            wCLIEANCLI = ""
            wCLIEANRA1 = ""
        End If
        'kLen = Len(X)
        wsExcel.Cells(wRow, 1) = wCli
        wsExcel.Cells(wRow, 2) = wCLIEANRA1
        wsExcel.Cells(wRow, 3) = wOPE
        wsExcel.Cells(wRow, 4) = wDOS
        wsExcel.Cells(wRow, 5) = wErr
        wsExcel.Cells(wRow, 6) = wLIB
        wsExcel.Cells(wRow, 7) = wNAT
        wsExcel.Cells(wRow, 8) = wMTD
        wsExcel.Cells(wRow, 9) = wDev
        wsExcel.Cells(wRow, 10) = wAmj
    End Select
Next kIndex

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub







Private Sub fgDetail_Display()
Dim wColor As Long
Dim X As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler

blnBIAMVT = False

fgDetail.Visible = False: fraDetail.Visible = False
SSTab2.Visible = False
fgCPTPIE.Visible = False
fraCompte.Visible = False: fraYDOSXOD0.Visible = False
    fraSwift.Visible = False

fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.Row = 0

currentAction = "fgDetail_Display"

libDOSSLDOPE = xYDOSSLD0.DOSSLDOPE
libDOSSLDDEV = xYDOSSLD0.DOSSLDDEV
libDOSSLDNUM = xYDOSSLD0.DOSSLDNUM
libDOSSLDCLI = xYDOSSLD0.DOSSLDCLI
libDOSSLDLIB = ""
Select Case Mid$(xYDOSSLD0.DOSSLDCLI, 1, 1)
    Case "0"
        xWhere = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & xYDOSSLD0.DOSSLDCLI & "'"
        Set rsSab = cnsab.Execute(xWhere)
        If Not rsSab.EOF Then libDOSSLDLIB = rsSab("CLIENARA1")
    Case "R"
        xWhere = "select ENCTIERA1 from " & paramIBM_Library_SAB & ".ZENCTIE0 where ENCTIEETA = 1 and ENCTIETIE = '0" & Mid$(xYDOSSLD0.DOSSLDCLI, 2, 6) & "'"
        Set rsSab = cnsab.Execute(xWhere)
        If Not rsSab.EOF Then libDOSSLDLIB = rsSab("ENCTIERA1")
    Case "T":
        If xYDOSSLD0.DOSSLDOPE = "RDE" Then
            If Len(xYDOSSLD0.DOSSLDCLI) = 8 Then ' ATTENTION longueur = 8 quelques fois !!!!!!!!!!!!!!!!!!!!!
                xWhere = "select ENCTIERA1 from " & paramIBM_Library_SAB & ".ZENCTIE0 where ENCTIEETA = 1 and ENCTIETIE = '" & Mid$(xYDOSSLD0.DOSSLDCLI, 2, 7) & "'"
            Else
                xWhere = "select ENCTIERA1 from " & paramIBM_Library_SAB & ".ZENCTIE0 where ENCTIEETA = 1 and ENCTIETIE = '0" & Mid$(xYDOSSLD0.DOSSLDCLI, 2, 6) & "'"
            End If
            Set rsSab = cnsab.Execute(xWhere)
            If Not rsSab.EOF Then libDOSSLDLIB = rsSab("ENCTIERA1")
        Else
            If Len(xYDOSSLD0.DOSSLDCLI) = 8 Then ' ATTENTION longueur = 8 quelques fois !!!!!!!!!!!!!!!!!!!!!
                xWhere = "select CAUTIRRA1 from " & paramIBM_Library_SAB & ".ZCAUTIR0 where CAUTIRETA = 1 and CAUTIRNUM = '" & Mid$(xYDOSSLD0.DOSSLDCLI, 2, 7) & "'"
            Else
                xWhere = "select CAUTIRRA1 from " & paramIBM_Library_SAB & ".ZCAUTIR0 where CAUTIRETA = 1 and CAUTIRNUM = '0" & Mid$(xYDOSSLD0.DOSSLDCLI, 2, 6) & "'"
            End If
            Set rsSab = cnsab.Execute(xWhere)
            If Not rsSab.EOF Then libDOSSLDLIB = rsSab("CAUTIRRA1")
        End If
End Select

xWhere = " where DOSSLDOPE ='" & xYDOSSLD0.DOSSLDOPE & "'" _
     & " and DOSSLDNUM = " & xYDOSSLD0.DOSSLDNUM _
     & " and DOSSLDDEV = '" & xYDOSSLD0.DOSSLDDEV & "'" _
     & " and DOSSLDCLI = '" & xYDOSSLD0.DOSSLDCLI & "'" _
     & " order by DOSSLDPCI "

Call arrYDOSSLD0_SQL(xWhere)


For I = 1 To arrYDOSSLD0_Nb
         
    xYDOSSLD0 = arrYDOSSLD0(I)
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_DisplayLine I
    'If fgSelect.Rows = 2 And Not blnBIAMVT Then
    If Not blnBIAMVT Then
        Select Case Mid$(xYDOSSLD0.DOSSLDPCI, 1, 5)
            Case "91120", "91122", "98050", "91130", "91131", "90312", "98520": blnBIAMVT = True: oldYDOSSLD0 = xYDOSSLD0
        End Select
    End If
Next I

fgDetail.Visible = True: fraDetail.Visible = True
If arrYDOSSLD0_Nb = 1 Then
    arrYDOSSLD0_Index = 1
        oldYDOSSLD0 = arrYDOSSLD0(arrYDOSSLD0_Index)
        xYDOSSLD0 = oldYDOSSLD0
        fgBIAMVT_Display
        fgDossier.Visible = False
        fgYSWISAB0.Visible = False
        fgCOM.Visible = False
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgDetail_Display_3()
Dim wColor As Long
Dim X As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
fgDetail.Visible = False: fraDetail.Visible = False
SSTab2.Visible = False
fgCPTPIE.Visible = False
fraCompte.Visible = False: fraYDOSXOD0.Visible = False
fraSwift.Visible = False

fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.Row = 0

currentAction = "fgDetail_Display"

libDOSSLDOPE = xYDOSSLD0.DOSSLDOPE
libDOSSLDDEV = xYDOSSLD0.DOSSLDDEV
libDOSSLDNUM = xYDOSSLD0.DOSSLDNUM
libDOSSLDCLI = ""
libDOSSLDLIB = ""

xWhere = " where DOSSLDOPE ='" & xYDOSSLD0.DOSSLDOPE & "'" _
     & " and DOSSLDNUM = " & xYDOSSLD0.DOSSLDNUM _
     & " and DOSSLDDEV = '" & xYDOSSLD0.DOSSLDDEV & "'" _
     & " order by DOSSLDPCI "

Call arrYDOSSLD0_SQL(xWhere)


For I = 1 To arrYDOSSLD0_Nb
         
    xYDOSSLD0 = arrYDOSSLD0(I)
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_DisplayLine I
    
Next I

fgDetail.Visible = True: fraDetail.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub



Private Sub fgDetail_Display_2()
Dim wColor As Long
Dim X As String, xWhere As String, xOPE As String
Dim xSql As String
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String
Dim xDOSSLDM As String, xDOSSLDG As String, xDOSSLDK As String

On Error GoTo Error_Handler
fgDetail.Visible = False: fraDetail.Visible = False
SSTab2.Visible = False
fgCPTPIE.Visible = False
fraCompte.Visible = False: fraYDOSXOD0.Visible = False
fraSwift.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.Row = 0

currentAction = "fgDetail_Display_2"

libDOSSLDOPE = xYDOSSLD1.DOSSLDPCI
libDOSSLDDEV = xYDOSSLD1.DOSSLDDEV
libDOSSLDNUM = ""
libDOSSLDCLI = xYDOSSLD1.DOSSLDCLI

libDOSSLDLIB = ""
xSql = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & xYDOSSLD1.DOSSLDCLI & "'"
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then libDOSSLDLIB = rsSab("CLIENARA1")


xWhere = " where DOSSLDPCI ='" & xYDOSSLD1.DOSSLDPCI & "'" _
     & " and DOSSLDCLI = '" & xYDOSSLD1.DOSSLDCLI & "'" _
     & " and DOSSLDDEV = '" & xYDOSSLD1.DOSSLDDEV & "'"

     
xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSSLD1 " & xWhere & " order by DOSSLDPCI "
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then
    V = rsYDOSSLD1_GetBuffer(rsSab, xYDOSSLD1)
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_DisplayLine_2 I
End If

fgDetail.Rows = fgDetail.Rows + 1
 fgDetail.Row = fgDetail.Rows - 1
For I = 0 To 4: fgDetail.Col = I: fgDetail.Text = "===========================": Next I

xWhere = " where DOSSLDPCI ='" & xYDOSSLD1.DOSSLDPCI & "'" _
     & " and DOSSLDCLI = '" & xYDOSSLD1.DOSSLDCLI & "'" _
     & " and DOSSLDDEV = '" & xYDOSSLD1.DOSSLDDEV & "'"
     
If chkSelect_DOSSLDSTA = "1" Then xWhere = xWhere & "   and DOSSLDSTA not in ('  ','80','90')"
'If chkSelect_DOSSLDSVC = "1" Then xWhere = xWhere & "   and (DOSSLDSVC <> '01' or DOSSLDSTA <> '01')"
If chkSelect_DOSSLDMG = "1" Then xWhere = xWhere & " and (DOSSLDMSD <> DOSSLDGSD) "



Call arrYDOSSLD0_SQL(xWhere & " order by DOSSLDPCI ")


For I = 1 To arrYDOSSLD0_Nb
    xYDOSSLD0 = arrYDOSSLD0(I)
    blnOk = True
    If chkSelect_DOSSLDSVC = "1" Then
        If xYDOSSLD0.DOSSLDSVC = "01" And xYDOSSLD0.DOSSLDSTA = "01" Then blnOk = False
    End If
    If blnOk Then
        xYDOSSLD0 = arrYDOSSLD0(I)
        fgDetail.Rows = fgDetail.Rows + 1
        fgDetail.Row = fgDetail.Rows - 1
        fgDetail_DisplayLine I
    End If
Next I

fgDetail.Visible = True: fraDetail.Visible = True
mnuPrint_2_Exportation.Enabled = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub



Private Sub fgBIAMVT_Display()
Dim wColor As Long
Dim X As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
SSTab2.Visible = False: SSTab2.Tab = 0
fgCPTPIE.Visible = False
fraCompte.Visible = False: fraYDOSXOD0.Visible = False
fraSwift.Visible = False
fgBIAMVT_Reset

fgBIAMVT.Rows = 1
fgBIAMVT.FormatString = fgBIAMVT_FormatString
fgBIAMVT.Row = 0

currentAction = "fgBIAMVT_Display"

xWhere = " where DOSMVTOPE = '" & oldYDOSSLD0.DOSSLDOPE & "'" _
     & " and DOSMVTNUM = " & oldYDOSSLD0.DOSSLDNUM _
     & " and DOSMVTDEV = '" & oldYDOSSLD0.DOSSLDDEV & "'" _
     & " and DOSMVTPCI ='" & oldYDOSSLD0.DOSSLDPCI & "'" _
     & " and DOSMVTCLI ='" & oldYDOSSLD0.DOSSLDCLI & "'" _
     & " and MOUVEMETA = 1" _
     & " and MOUVEMPIE = DOSMVTPIE  and MOUVEMECR = DOSMVTECR" _
     & " order by DOSMVTDTR  "

Call arrYDOSMVT0_SQL(xWhere)


For I = 1 To arrYDOSMVT0_Nb
         
    xYDOSMVT0 = arrYDOSMVT0(I)
    xYBIAMVTH = arrYBIAMVTH(I)
    fgBIAMVT.Rows = fgBIAMVT.Rows + 1
    fgBIAMVT.Row = fgBIAMVT.Rows - 1
    fgBIAMVT_DisplayLine I
    
Next I

SSTab2.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgDossier_Display()
Dim wColor As Long
Dim xSql As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
SSTab2.Visible = False
fgDossier.Visible = False
fgYSWISAB0.Visible = False
fgCOM.Visible = False
fgDossier_Reset

fgDossier.Rows = 1
fgDossier.FormatString = fgDossier_FormatString
fgDossier.Row = 0

currentAction = "fgDossier_Display"
mCDOMODDMO = 0
'__________________________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SAB & ".ZCDOMOD0 " _
     & " where CDOMODCOP = '" & xYDOSSLD0.DOSSLDOPE & "'" _
     & " and   CDOMODDOS = " & xYDOSSLD0.DOSSLDNUM _
     & " order by CDOMODNMO"
Set rsSab = cnsab.Execute(xSql)


Do While Not rsSab.EOF
    fgDossier.Rows = fgDossier.Rows + 1
    fgDossier.Row = fgDossier.Rows - 1
    fgDossier_DisplayLine_ZCDOMOD0 I
    rsSab.MoveNext
Loop
'__________________________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
     & " where CDODOSCOP = '" & xYDOSSLD0.DOSSLDOPE & "'" _
     & " and   CDODOSDOS = " & xYDOSSLD0.DOSSLDNUM
Set rsSab = cnsab.Execute(xSql)


Do While Not rsSab.EOF
    fgDossier.Rows = fgDossier.Rows + 1
    fgDossier.Row = fgDossier.Rows - 1
    fgDossier_DisplayLine_ZCDODOS0 I
    rsSab.MoveNext
Loop
'__________________________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SAB & ".ZCDOUTI0 " _
     & " where CDOUTICOP = '" & xYDOSSLD0.DOSSLDOPE & "'" _
     & " and   CDOUTIDOS = " & xYDOSSLD0.DOSSLDNUM _
     & " order by CDOUTIUTI"
Set rsSab = cnsab.Execute(xSql)


Do While Not rsSab.EOF
    fgDossier.Rows = fgDossier.Rows + 1
    fgDossier.Row = fgDossier.Rows - 1
    fgDossier_DisplayLine_ZCDOUTI0 I
    rsSab.MoveNext
Loop
'__________________________________________________________________________________________________

If fgDossier.Rows > 1 Then
    fgDossier_Sort1 = 0
    fgDossier_Sort2 = 0
    fgDossier_Sort
End If
SSTab2.Visible = True
fgDossier.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgDossier_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
SSTab2.Visible = False: fraDetail.Visible = False
mRow = fgDossier.Row

If lRow > 0 And lRow < fgDossier.Rows Then
    fgDossier.Row = lRow
    For I = fgDossier_arrIndex To fgDossier.FixedCols Step -1
        fgDossier.Col = I: fgDossier.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgDossier.Row = mRow
    If fgDossier.Row > 0 Then
        lRow = fgDossier.Row
        fgDossier.Col = fgDossier_arrIndex
        lColor_Old = fgDossier.CellBackColor
        For I = fgDossier_arrIndex To fgDossier.FixedCols Step -1
          fgDossier.Col = I: fgDossier.CellBackColor = lColor
        Next I
    End If
End If
fgDossier.LeftCol = fgDossier.FixedCols
SSTab2.Visible = True: fraDetail.Visible = True
End Sub
Public Sub fgDossier_DisplayLine_ZCDOMOD0(lIndex As Long)
Dim wColor As Long
Dim wCDOMODEVE As String, X As String
On Error Resume Next

If mCDOMODDMO = 0 Then mCDOMODDMO = rsSab("CDOMODOUV")
wCDOMODEVE = rsSab("CDOMODEVE")
If wCDOMODEVE = "90" Or wCDOMODEVE = "80" Then
    wColor = vbRed
    fgDossier.Col = 8: fgDossier.Text = Format$(rsSab("CDOMODANN"), "### ### ### ##0.00")
    fgDossier.CellForeColor = wColor
    fgDossier.Col = 0: fgDossier.Text = dateImp_Amj(rsSab("CDOMODDAN") + 19000000)
    fgDossier.CellForeColor = wColor
Else
    fgDossier.Col = 0: fgDossier.Text = dateImp_Amj(mCDOMODDMO + 19000000)
    wColor = vbBlue
End If
mCDOMODDMO = rsSab("CDOMODDMO")

fgDossier.Col = 1: fgDossier.Text = "Mod"
fgDossier.Col = 2: fgDossier.Text = rsSab("CDOMODNMO")
fgDossier.Col = 3: fgDossier.Text = wCDOMODEVE
fgDossier.CellForeColor = wColor
fgDossier.Col = 4: fgDossier.Text = rsSab("CDOMODETA")
fgDossier.Col = 5: fgDossier.Text = rsSab("CDOMODCON")

fgDossier.Col = 6: fgDossier.Text = Format$(rsSab("CDOMODMOT"), "### ### ### ##0.00")
fgDossier.CellForeColor = vbBlue
fgDossier.Col = 7: fgDossier.Text = rsSab("CDOMODDEV")

fgDossier.Col = 9
X = ""
If rsSab("CDOMODMOC") > 0 Then X = "Confirmé :" & Format$(rsSab("CDOMODMOC"), "### ### ### ##0.00")
If rsSab("CDOMODMOD") > 0 Then X = X & "     Ducroire :" & Format$(rsSab("CDOMODMOD"), "### ### ### ##0.00")
fgDossier.Text = X
fgDossier.Col = fgDossier_arrIndex: fgDossier.Text = lIndex
End Sub

Public Sub fgDossier_DisplayLine_ZCDOUTI0(lIndex As Long)
Dim wColor As Long, wBackColor As Long
Dim wCDOUTIEVE As String
On Error Resume Next

wCDOUTIEVE = rsSab("CDOUTIEVE")

wBackColor = RGB(255, 255, 192)

    wColor = vbBlue
fgDossier.Col = 0: fgDossier.Col = 0: fgDossier.Text = dateImp_Amj(rsSab("CDOUTIPRE") + 19000000)

    fgDossier.CellBackColor = wBackColor

fgDossier.Col = 1: fgDossier.Text = "UTI"
    fgDossier.CellBackColor = wBackColor
fgDossier.Col = 2: fgDossier.Text = rsSab("CDOUTIUTI")
    fgDossier.CellBackColor = wBackColor
fgDossier.Col = 3: fgDossier.Text = wCDOUTIEVE
    fgDossier.CellBackColor = wBackColor
fgDossier.CellForeColor = wColor
    fgDossier.CellBackColor = wBackColor
fgDossier.Col = 4: fgDossier.Text = rsSab("CDOUTIETA")
    fgDossier.CellBackColor = wBackColor
fgDossier.Col = 5: fgDossier.Text = rsSab("CDOUTITMO")
    fgDossier.CellBackColor = wBackColor

fgDossier.Col = 6: fgDossier.Text = Format$(-rsSab("CDOUTIMON"), "### ### ### ##0.00")
    fgDossier.CellBackColor = wBackColor
fgDossier.CellForeColor = vbRed
fgDossier.Col = 7
    fgDossier.CellBackColor = wBackColor
fgDossier.Col = 0
    fgDossier.CellBackColor = wBackColor
fgDossier.Col = 8
    fgDossier.CellBackColor = wBackColor

fgDossier.Col = 9
    fgDossier.CellBackColor = wBackColor
fgDossier.Text = "date remise :" & dateImp10(rsSab("CDOUTIDRE") + 19000000) _
               & "     Attente :" & rsSab("CDOUTIATT")
fgDossier.Col = fgDossier_arrIndex: fgDossier.Text = lIndex
End Sub

Public Sub fgDossier_DisplayLine_ZCDODOS0(lIndex As Long)
Dim wColor As Long, wBackColor As Long
Dim wCDODOSEVE As String, X As String
On Error Resume Next

wBackColor = RGB(192, 255, 192)
wCDODOSEVE = rsSab("CDODOSEVE")
If wCDODOSEVE = "90" Or wCDODOSEVE = "80" Then
    wColor = vbRed
    fgDossier.Col = 8: fgDossier.Text = Format$(rsSab("CDODOSANN"), "### ### ### ##0.00")
    fgDossier.CellForeColor = wColor
    fgDossier.CellBackColor = wBackColor
    fgDossier.Col = 0: fgDossier.Text = dateImp_Amj(rsSab("CDODOSDAN") + 19000000)
    fgDossier.CellForeColor = wColor
    fgDossier.CellBackColor = wBackColor
Else
    wColor = vbBlue
    fgDossier.Col = 8
    fgDossier.CellBackColor = wBackColor
    fgDossier.Col = 0: fgDossier.Text = dateImp_Amj(mCDOMODDMO + 19000000)
    fgDossier.CellBackColor = wBackColor
End If

fgDossier.Col = 1: fgDossier.Text = "Dos"
fgDossier.CellBackColor = wBackColor
fgDossier.Col = 2
fgDossier.CellBackColor = wBackColor
fgDossier.Col = 3: fgDossier.Text = wCDODOSEVE
fgDossier.CellBackColor = wBackColor
fgDossier.CellForeColor = wColor
fgDossier.Col = 4: fgDossier.Text = rsSab("CDODOSETA")
fgDossier.CellBackColor = wBackColor
fgDossier.Col = 5: fgDossier.Text = rsSab("CDODOSCON")
fgDossier.CellBackColor = wBackColor

fgDossier.Col = 6: fgDossier.Text = Format$(rsSab("CDODOSMOT"), "### ### ### ##0.00")
fgDossier.CellBackColor = wBackColor
fgDossier.CellForeColor = vbBlue
fgDossier.Col = 7: fgDossier.Text = rsSab("CDODOSDEV")
fgDossier.CellBackColor = wBackColor

fgDossier.Col = 9
fgDossier.CellBackColor = wBackColor
X = ""
If rsSab("CDODOSMOC") > 0 Then X = "Confirmé :" & Format$(rsSab("CDODOSMOC"), "### ### ### ##0.00")
If rsSab("CDODOSMOD") > 0 Then X = X & "     Ducroire :" & Format$(rsSab("CDODOSMOD"), "### ### ### ##0.00")
fgDossier.Text = X
fgDossier.Col = fgDossier_arrIndex: fgDossier.Text = lIndex
End Sub

Public Sub fgCOM_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
SSTab2.Visible = False: fraDetail.Visible = False
mRow = fgCOM.Row

If lRow > 0 And lRow < fgCOM.Rows Then
    fgCOM.Row = lRow
    For I = fgCOM_arrIndex To fgCOM.FixedCols Step -1
        fgCOM.Col = I: fgCOM.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgCOM.Row = mRow
    If fgCOM.Row > 0 Then
        lRow = fgCOM.Row
        fgCOM.Col = fgCOM_arrIndex
        lColor_Old = fgCOM.CellBackColor
        For I = fgCOM_arrIndex To fgCOM.FixedCols Step -1
          fgCOM.Col = I: fgCOM.CellBackColor = lColor
        Next I
    End If
End If
fgCOM.LeftCol = fgCOM.FixedCols
SSTab2.Visible = True: fraDetail.Visible = True
End Sub
Private Sub fgCOM_Display()
Dim wColor As Long
Dim xSql As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
SSTab2.Visible = False
fgCOM.Visible = False
fgYSWISAB0.Visible = False
fgCOM.Visible = False
fgCOM_Reset

fraECNFPT.Visible = False
mECNFPT_Row = 0
mECNFPT_MTA = 0

fgCOM.Rows = 1
fgCOM.FormatString = fgCOM_FormatString
fgCOM.Row = 0
fgCOM.Col = 1: fgCOM.CellAlignment = 1
fgCOM.Col = 3: fgCOM.CellAlignment = 1
fgCOM.Col = 10: fgCOM.CellAlignment = 1
fgCOM.Col = 11: fgCOM.CellAlignment = 1
fgCOM.Col = 14: fgCOM.CellAlignment = 1
fgCOM.Col = 15: fgCOM.CellAlignment = 1
fgCOM.Col = 16: fgCOM.CellAlignment = 1
fgCOM.Col = 17: fgCOM.CellAlignment = 1

currentAction = "fgCOM_Display"
Select Case cmdSelect_SQL_K
    Case "5 ECNFPT": xWhere = " and CDOCOMDOS = " & xYDOSSLD0.DOSSLDNUM & " and CDOCOMCOM = 'ECNFPT' and CDOCOMETA <> '03' and CDOCOMMON > 0"
    Case "5 ECNFPT_Com", "5 ECNFPT_CD7": xWhere = " and CDOCOMCOM = 'ECNFPT' and CDOCOMETA <> '03' and CDOCOMMON > 0"
    Case Else: xWhere = " and CDOCOMDOS = " & xYDOSSLD0.DOSSLDNUM
End Select
'__________________________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SAB & ".ZCDOCOM0 " _
        & " left outer join " & paramIBM_Library_SAB & ".ZCDOCO20" _
     & " on CDOCO2ETB = CDOCOMETB and CDOCO2AGE = CDOCOMAGE and CDOCO2SER = CDOCOMSER and CDOCO2SSE = CDOCOMSSE" _
     & " and CDOCO2COP = CDOCOMCOP and CDOCO2DOS = CDOCOMDOS and CDOCO2NUR = CDOCOMNUR and CDOCO2UTI = CDOCOMUTI" _
     & " and CDOCO2EVE = CDOCOMEVE and CDOCO2SEQ = CDOCOMSEQ and CDOCO2SPE = CDOCOMSPE" _
        & " left outer join " & paramIBM_Library_SAB & ".ZCDOTC20" _
     & " on CDOTC2ETB = CDOCOMETB and CDOTC2AGE = CDOCOMAGE and CDOTC2SER = CDOCOMSER and CDOTC2SSE = CDOCOMSSE" _
     & " and CDOTC2COP = CDOCOMCOP and CDOTC2DOS = CDOCOMDOS and CDOTC2NUR = CDOCOMNUR and CDOTC2UTI = CDOCOMUTI" _
     & " and CDOTC2EVE = CDOCOMEVE and CDOTC2SEQ = CDOCOMSEQ" _
     & " where CDOCOMCOP = '" & xYDOSSLD0.DOSSLDOPE & "'" _
     & xWhere _
     & " order by CDOCOMDOS, CDOCOMUTI, CDOCOMCOM, CDOCOMDBP"

 '    & " order by CDOCOMDOS, CDOCOMNUR, CDOCOMUTI, CDOCOMEVE, CDOCOMSEQ, CDOCOMSPE"

Set rsSab = cnsab.Execute(xSql)

mCDOCOMMON = 0
mCDOCOMDOS = 0
mCDOCOMDOS = xYDOSSLD0.DOSSLDNUM
Do While Not rsSab.EOF
    If rsSab("CDOCOMDOS") <> mCDOCOMDOS Then
        fgCom_Display_Total
        mCDOCOMMON = 0
        mECNFPT_MTA = 0
        mCDOCOMDOS = rsSab("CDOCOMDOS")
    
    End If
    
    fgCOM.Rows = fgCOM.Rows + 1
    fgCOM.Row = fgCOM.Rows - 1
    fgCOM_DisplayLine_ZCDOCOM0 I
    rsSab.MoveNext
Loop
'__________________________________________________________________________________________________
fgCom_Display_Total
'__________________________________________________________________________________________________

SSTab2.Visible = True
fgCOM.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgCOM_Reset()
fgCOM.Clear
fgCOM_Sort1 = 0: fgCOM_Sort2 = 0
fgCOM_Sort1_Old = -1
fgCOM_RowDisplay = 0: fgCOM_RowClick = 0
fgCOM_arrIndex = fgCOM.Cols - 1
blnfgCOM_DisplayLine = False
fgCOM_SortAD = 6
fgCOM.LeftCol = fgCOM.FixedCols

End Sub
Public Sub fgCOM_Sort()
If fgCOM.Rows > 1 Then
    fgCOM.Row = 1
    fgCOM.RowSel = fgCOM.Rows - 1
    
    If fgCOM_Sort1_Old = fgCOM_Sort1 Then
        If fgCOM_SortAD = 5 Then
            fgCOM_SortAD = 6
        Else
            fgCOM_SortAD = 5
        End If
    Else
        fgCOM_SortAD = 5
    End If
    fgCOM_Sort1_Old = fgCOM_Sort1
    
    fgCOM.Col = fgCOM_Sort1
    fgCOM.ColSel = fgCOM_Sort2
    fgCOM.Sort = fgCOM_SortAD
End If

End Sub


Public Sub fgCOM_DisplayLine_ZCDOCOM0(lIndex As Long)
Dim wColor As Long, wCellBackColor As Long, curX As Currency, WCDOCOMMON As Currency, WCDOCO2MIN As Currency
Dim K As Integer, X As String, NbJ As Long
On Error Resume Next


fgCOM.Col = 0: fgCOM.Text = rsSab("CDOCOMUTI") & "-" & rsSab("CDOCOMEVE") & "-" & rsSab("CDOCOMSEQ") & "-" & rsSab("CDOCOMSPE")
fgCOM.Col = 1: fgCOM.Text = rsSab("CDOCOMBEN")
If rsSab("CDOCOMBEN") = "N" Then
    wColor = vbRed
    wCellBackColor = mColor_Y2
Else
    wColor = vbBlue
End If

fgCOM.Col = 2: fgCOM.Text = rsSab("CDOCOMCOM")
WCDOCOMMON = rsSab("CDOCOMMON")
If WCDOCOMMON <> 0 Then
    fgCOM.Col = 3: fgCOM.Text = Format$(WCDOCOMMON, "### ### ##0.00")
    fgCOM.CellForeColor = wColor
    If rsSab("CDOCOMETA") <> "03" Then
        For K = 2 To 8
            fgCOM.Col = K: fgCOM.CellFontBold = True
        Next K
    End If
Else
    wCellBackColor = RGB(220, 220, 220)
End If
If rsSab("CDOCOMMTV") <> 0 Then
    fgCOM.Col = 4: fgCOM.Text = Format$(rsSab("CDOCOMMTV"), "### ### ##0.00")
    fgCOM.CellForeColor = wColor
End If
fgCOM.Col = 5: fgCOM.Text = rsSab("CDOCOMDEV")

If rsSab("CDOCOMDBP") > 0 Then
    fgCOM.Col = 6: fgCOM.Text = dateImp10(rsSab("CDOCOMDBP") + 19000000)
End If
If rsSab("CDOCOMFNP") > 0 Then
    fgCOM.Col = 7: fgCOM.Text = dateImp10(rsSab("CDOCOMFNP") + 19000000)
End If
If rsSab("CDOCOMREG") > 0 Then
    fgCOM.Col = 9: fgCOM.Text = dateImp10(rsSab("CDOCOMREG") + 19000000)
End If

fgCOM.Col = 8: fgCOM.Text = rsSab("CDOCOMETA")
fgCOM.Col = 1: fgCOM.Text = rsSab("CDOCOMBEN")


If Not IsNull(rsSab("CDOCO2DOS")) Then
    If rsSab("CDOCO2TVA") <> "O" Then
        If rsSab("CDOCOMMTV") <> 0 Then fgCOM.Col = 4: fgCOM.Text = "N"
    End If
    
    If rsSab("CDOCO2MTA") <> 0 Then
        fgCOM.Col = 10: fgCOM.Text = Format$(rsSab("CDOCO2MTA") / 100, "### ### ##0.00")
        fgCOM.CellForeColor = wColor
    End If
    If rsSab("CDOCO2TX1") <> 0 Then
        fgCOM.Col = 11: fgCOM.Text = Format$(rsSab("CDOCO2TX1"), "##0.00000")
        fgCOM.CellForeColor = wColor
    End If
    WCDOCO2MIN = rsSab("CDOCO2MIN") / 100
    If WCDOCO2MIN <> 0 Then
        fgCOM.Col = 16: fgCOM.Text = Format$(WCDOCO2MIN, "### ### ##0.00")
        fgCOM.CellForeColor = wColor
    End If
    If rsSab("CDOCO2NBJ") <> 0 Then
        fgCOM.Col = 13: fgCOM.Text = Format$(rsSab("CDOCO2NBJ"), "###")
    End If
    If rsSab("CDOCOMCOM") = "ECNFPT" And WCDOCOMMON > 0 And rsSab("CDOCOMBEN") = "O" Then
        If rsSab("CDOCOMETA") <> "03" Then
            Select Case rsSab("CDOTC2PER")
                Case "M": NbJ = 3000
                Case "T": NbJ = 9000
                Case "S": NbJ = 18000
                Case "A": NbJ = 36000
                Case Else: NbJ = 100
            End Select
            curX = Fix(rsSab("CDOCO2MTA") * rsSab("CDOCO2TX1") * rsSab("CDOCO2NBJ") / NbJ + 0.00500001) / 100
            If curX = 0 Then
                fgCOM.Col = 3: fgCOM.CellBackColor = mColor_W0
                curX = WCDOCOMMON
            Else
                 If rsSab("CDOCOMREG") > 0 Then
                    fgCOM.Col = 9: fgCOM.CellBackColor = mColor_G1
                    If curX < WCDOCO2MIN And WCDOCO2MIN > 0 Then curX = WCDOCO2MIN
                End If 'Else
                    'mECNFPT_Row = fgCOM.Row
                    mECNFPT_MTA = rsSab("CDOCO2MTA")
                    mECNFPT_PER = rsSab("CDOCO2PER")
                    mECNFPT_TX1 = rsSab("CDOCO2TX1")
                    mECNFPT_NBJ = rsSab("CDOCO2NBJ")
                    mECNFPT_DDEB = rsSab("CDOCOMDBP") + 19000000
                    mECNFPT_DFIN = rsSab("CDOCOMFNP") + 19000000
                    mECNFPT_MON = curX
                    mECNFPT_Ratio = NbJ
                    mECNFPT_MIN = WCDOCO2MIN
               'End If
                   
                fgCOM.Col = 3: fgCOM.Text = Format$(curX, "### ### ##0.00")
                If Abs(curX - WCDOCOMMON) < 0.1 Then
                    fgCOM.CellBackColor = mColor_G1
                Else
                    
                    fgCOM.CellBackColor = mColor_Y2
                    fgCOM.Col = 17: fgCOM.Text = Format$(WCDOCOMMON, "### ### ##0.00")
                    fgCOM.CellBackColor = mColor_Y2
                    
                    If cmdSelect_SQL_K = "5 ECNFPT_CD7" Then
                        arrYDOSCD70_Nb = arrYDOSCD70_Nb + 1
                        If arrYDOSCD70_Nb > arrYDOSCD70_Max Then
                            arrYDOSCD70_Max = arrYDOSCD70_Max + 100
                            ReDim Preserve arrYDOSCD70(arrYDOSCD70_Max)
                        End If
                        
                        arrYDOSCD70(arrYDOSCD70_Nb).DOSCD7OPE = rsSab("CDOCOMCOP")
                        arrYDOSCD70(arrYDOSCD70_Nb).DOSCD7NUM = rsSab("CDOCOMDOS")
                        arrYDOSCD70(arrYDOSCD70_Nb).DOSCD7DDEB = rsSab("CDOCOMDBP") + 19000000
                        arrYDOSCD70(arrYDOSCD70_Nb).DOSCD7DFIN = rsSab("CDOCOMFNP") + 19000000
                        arrYDOSCD70(arrYDOSCD70_Nb).DOSCD7MTD = curX
                    End If
                End If
            End If
            mCDOCOMMON = mCDOCOMMON + curX
        End If
    End If

End If

If Not IsNull(rsSab("CDOTC2DOS")) Then
    fgCOM.Col = 12: fgCOM.Text = rsSab("CDOTC2PER")

    If rsSab("CDOTC2TX1") <> 0 Then
        fgCOM.Col = 14: fgCOM.Text = Format$(rsSab("CDOTC2TX1"), "##0.00000")
        fgCOM.CellForeColor = wColor
        fgCOM.CellBackColor = mColor_G1
    End If
    If rsSab("CDOTC2MTF") <> 0 Then
        fgCOM.Col = 15: fgCOM.Text = Format$(rsSab("CDOTC2MTF"), "### ### ### ##0.00")
        fgCOM.CellForeColor = wColor
        fgCOM.CellBackColor = mColor_G1
    End If
End If
If wCellBackColor > 0 Then
    For K = 0 To 13
        fgCOM.Col = K: fgCOM.CellBackColor = wCellBackColor
    Next K
End If

fgCOM.Col = fgCOM_arrIndex: fgCOM.Text = lIndex
End Sub




Public Sub fgDossier_Reset()
fgDossier.Clear
fgDossier_Sort1 = 0: fgDossier_Sort2 = 0
fgDossier_Sort1_Old = -1
fgDossier_RowDisplay = 0: fgDossier_RowClick = 0
fgDossier_arrIndex = fgDossier.Cols - 1
blnfgDossier_DisplayLine = False
fgDossier_SortAD = 6
fgDossier.LeftCol = fgDossier.FixedCols

End Sub

Public Sub fgDossier_Sort()
If fgDossier.Rows > 1 Then
    fgDossier.Row = 1
    fgDossier.RowSel = fgDossier.Rows - 1
    
    If fgDossier_Sort1_Old = fgDossier_Sort1 Then
        If fgDossier_SortAD = 5 Then
            fgDossier_SortAD = 6
        Else
            fgDossier_SortAD = 5
        End If
    Else
        fgDossier_SortAD = 5
    End If
    fgDossier_Sort1_Old = fgDossier_Sort1
    
    fgDossier.Col = fgDossier_Sort1
    fgDossier.ColSel = fgDossier_Sort2
    fgDossier.Sort = fgDossier_SortAD
End If

End Sub



Public Sub fgCourrier_Display()
Dim wColor As Long, X As String
Dim objFolder, objFiles
Dim fsoFile As File

On Error GoTo Error_Handler
currentAction = "fgCourrier_Display"
SSTab2.Visible = False
fgCourrier.Visible = False
fgCourrier_Reset

fgCourrier.Rows = 1
fgCourrier.FormatString = fgCourrier_FormatString
fgCourrier.Row = 0


X = xYDOSSLD0.DOSSLDOPE & "_" & Format(xYDOSSLD0.DOSSLDNUM, "000000") & "\"
If xYDOSSLD0.DOSSLDOPE = "RDE" Or xYDOSSLD0.DOSSLDOPE = "RDI" Then
    If Dir(paramRDE_Dossier_Path_DROPI & X) <> "" Then
        Set objFolder = msFileSystem.GetFolder(paramRDE_Dossier_Path_DROPI & X)
        Set objFiles = objFolder.Files
        For Each fsoFile In objFiles
            '$JPL 2014-11-27 If InStr(fsoFile.Type, "Document") > 0 Then
                fgCourrier.Rows = fgCourrier.Rows + 1
                fgCourrier.Row = fgCourrier.Rows - 1
                fgCourrier.Col = 0: fgCourrier.Text = fsoFile.DateCreated
                fgCourrier.Col = 1: fgCourrier.Text = fsoFile.Name
            '$JPL 2014-11-27 End If
        Next
    End If
Else
    If Dir(paramCDO_Dossier_Path_DROPI & X) <> "" Then
        Set objFolder = msFileSystem.GetFolder(paramCDO_Dossier_Path_DROPI & X)
        Set objFiles = objFolder.Files
        For Each fsoFile In objFiles
            '$JPL 2014-11-27 If InStr(fsoFile.Type, "Document") > 0 Then
                fgCourrier.Rows = fgCourrier.Rows + 1
                fgCourrier.Row = fgCourrier.Rows - 1
                fgCourrier.Col = 0: fgCourrier.Text = fsoFile.DateCreated
                fgCourrier.Col = 1: fgCourrier.Text = fsoFile.Name
            '$JPL 2014-11-27 End If
        Next
    End If
End If
'__________________________________________________________________________________________________
SSTab2.Visible = True
fgCourrier.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgCourrier_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
SSTab2.Visible = False: fraDetail.Visible = False
mRow = fgCourrier.Row

If lRow > 0 And lRow < fgCourrier.Rows Then
    fgCourrier.Row = lRow
    For I = fgCourrier_arrIndex To fgCourrier.FixedCols Step -1
        fgCourrier.Col = I: fgCourrier.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgCourrier.Row = mRow
    If fgCourrier.Row > 0 Then
        lRow = fgCourrier.Row
        fgCourrier.Col = fgCourrier_arrIndex
        lColor_Old = fgCourrier.CellBackColor
        For I = fgCourrier_arrIndex To fgCourrier.FixedCols Step -1
          fgCourrier.Col = I: fgCourrier.CellBackColor = lColor
        Next I
    End If
End If
fgCourrier.LeftCol = fgCourrier.FixedCols
SSTab2.Visible = True: fraDetail.Visible = True
End Sub
Public Sub fgCourrier_Reset()
fgCourrier.Clear
fgCourrier_Sort1 = 0: fgCourrier_Sort2 = 0
fgCourrier_Sort1_Old = -1
fgCourrier_RowDisplay = 0: fgCourrier_RowClick = 0
fgCourrier_arrIndex = fgCourrier.Cols - 1
blnfgCourrier_DisplayLine = False
fgCourrier_SortAD = 6
fgCourrier.LeftCol = fgCourrier.FixedCols

End Sub

Public Sub fgCourrier_Sort()
If fgCourrier.Rows > 1 Then
    fgCourrier.Row = 1
    fgCourrier.RowSel = fgCourrier.Rows - 1
    
    If fgCourrier_Sort1_Old = fgCourrier_Sort1 Then
        If fgCourrier_SortAD = 5 Then
            fgCourrier_SortAD = 6
        Else
            fgCourrier_SortAD = 5
        End If
    Else
        fgCourrier_SortAD = 5
    End If
    fgCourrier_Sort1_Old = fgCourrier_Sort1
    
    fgCourrier.Col = fgCourrier_Sort1
    fgCourrier.ColSel = fgCourrier_Sort2
    fgCourrier.Sort = fgCourrier_SortAD
End If

End Sub




Private Sub fgCPTPIE_Display()
Dim wColor As Long
Dim xSql As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
fgCPTPIE.Visible = False: fraDetail.Visible = False
fraCompte.Visible = False: fraYDOSXOD0.Visible = False
fraSwift.Visible = False
fgCPTPIE_Reset

fgCPTPIE.Rows = 1
fgCPTPIE.FormatString = fgCPTPIE_FormatString
fgCPTPIE.Row = 0

currentAction = "fgCPTPIE_Display"

xWhere = " where MOUVEMETA = 1 " _
     & " and MOUVEMPIE = " & oldYDOSMVT0.DOSMVTPIE _
     & " and MOUVEMCOM = COMPTECOM " _
     & " order by MOUVEMECR "


xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH , " & paramIBM_Library_SABSPE & ".YBIACPT0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYBIAMVT0_GetBuffer(rsSab, xYCPTPIEH)
         
    fgCPTPIE.Rows = fgCPTPIE.Rows + 1
    fgCPTPIE.Row = fgCPTPIE.Rows - 1
    fgCPTPIE_DisplayLine I
    
    rsSab.MoveNext
Loop

fgCPTPIE.Visible = True: fraDetail.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub




Public Sub fgDetail_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim blnSolde As Boolean

On Error Resume Next
fgDetail.Col = 0
If cmdSelect_SQL_K = "2" Then
    fgDetail.CellForeColor = RGB(32, 96, 32)
Else
    fgDetail.Text = xYDOSSLD0.DOSSLDPCI & " " & xYDOSSLD0.DOSSLDDEV & " " & xYDOSSLD0.DOSSLDCLI
End If
        Select Case Mid$(xYDOSSLD0.DOSSLDPCI, 1, 5)
            Case "91120", "91122", "98050", "90312", "98520": blnSolde = True: wColor = RGB(255, 208, 255)
            Case "70721", "91130", "91131": blnSolde = True: wColor = RGB(238, 221, 255)
            Case Else: blnSolde = False
                If xYDOSSLD0.DOSSLDOPE = "ENG" Or xYDOSSLD0.DOSSLDOPE = "GAR" Then
                    If Mid$(xYDOSSLD0.DOSSLDPCI, 1, 1) = "9" Then blnSolde = True
                End If
        End Select

fgDetail.Col = 1: fgDetail.Text = xYDOSSLD0.DOSSLDOPE
fgDetail.Col = 2: fgDetail.Text = xYDOSSLD0.DOSSLDNUM
fgDetail.Col = 3: fgDetail.Text = Format$(xYDOSSLD0.DOSSLDMSD, "### ### ### ##0.00")
fgDetail.CellForeColor = IIf(xYDOSSLD0.DOSSLDMSD < 0, vbRed, vbBlue)

If blnSolde Then
    fgDetail.Col = 4: fgDetail.Text = Format$(xYDOSSLD0.DOSSLDGSD, "### ### ### ##0.00")
    fgDetail.CellForeColor = IIf(xYDOSSLD0.DOSSLDGSD < 0, vbRed, vbBlue)

    If xYDOSSLD0.DOSSLDMSD <> xYDOSSLD0.DOSSLDGSD Then
        fgDetail.CellBackColor = wColor
        If xYDOSSLD0.DOSSLDSTA = "  " Or xYDOSSLD0.DOSSLDSTA = "90" Then
            If xYDOSSLD0.DOSSLDNUM < 70000 Or xYDOSSLD0.DOSSLDNUM > 900000 Then fgDetail.CellBackColor = RGB(255, 220, 192)
        End If
    End If
End If

fgDetail.Col = 5: fgDetail.Text = xYDOSSLD0.DOSSLDSTA
fgDetail.Col = 6: fgDetail.Text = xYDOSSLD0.DOSSLDSVC


fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex
End Sub


Public Sub fgDetail_DisplayLine_2(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim blnSolde As Boolean

On Error Resume Next

Select Case Mid$(xYDOSSLD1.DOSSLDPCI, 1, 5)
    Case "91120", "91122", "98050", "90312", "98520": blnSolde = True: wColor = RGB(255, 208, 255)
    Case "70721", "91130", "91131": blnSolde = True: wColor = RGB(238, 221, 255)
    Case Else: blnSolde = False
End Select

fgDetail.Col = 0: fgDetail.Text = xYDOSSLD1.DOSSLDPCI & " " & xYDOSSLD1.DOSSLDDEV & " " & xYDOSSLD1.DOSSLDCLI
fgDetail.CellFontUnderline = True
fgDetail.CellFontBold = True
fgDetail.Col = 1: fgDetail.Text = ""
fgDetail.CellForeColor = vbRed
fgDetail.CellFontBold = True
fgDetail.Col = 2: fgDetail.Text = ""
fgDetail.CellForeColor = vbBlue
fgDetail.CellFontBold = True
fgDetail.Col = 3: fgDetail.Text = Format$(xYDOSSLD1.DOSSLDMSD, "### ### ### ##0.00")
fgDetail.CellForeColor = IIf(xYDOSSLD1.DOSSLDMSD < 0, vbRed, vbBlue)
fgDetail.CellFontBold = True
If blnSolde Then
    fgDetail.Col = 4: fgDetail.Text = Format$(xYDOSSLD1.DOSSLDGSD, "### ### ### ##0.00")
    fgDetail.CellForeColor = IIf(xYDOSSLD1.DOSSLDGSD < 0, vbRed, vbBlue)
    fgDetail.CellFontBold = True
    If xYDOSSLD1.DOSSLDMSD <> xYDOSSLD1.DOSSLDGSD Then fgDetail.CellBackColor = wColor
End If

fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex
End Sub

Public Sub fgBIAMVT_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim blnSolde As Boolean

On Error Resume Next
fgBIAMVT.Col = 0: fgBIAMVT.Text = xYBIAMVTH.MOUVEMSER & " " & xYBIAMVTH.MOUVEMSSE & " " & xYBIAMVTH.MOUVEMOPE & " " & xYBIAMVTH.MOUVEMEVE & " " & xYBIAMVTH.MOUVEMNUM
'fgBIAMVT.Col = 0: fgBIAMVT.Text = dateImp10_S(xYDOSMVT0.DOSMVTDTR)

fgBIAMVT.Col = 1: fgBIAMVT.Text = xYBIAMVTH.MOUVEMCOM


'fgBIAMVT.Col = IIf(xYDOSMVT0.DOSMVTKDC = "D", 2, 3)
fgBIAMVT.Col = IIf(xYDOSMVT0.DOSMVTMTD < 0, 2, 3)

fgBIAMVT.Text = Format$(Abs(xYDOSMVT0.DOSMVTMTD), "### ### ### ##0.00")

If xYDOSMVT0.DOSMVTMTD < 0 Then
    fgBIAMVT.CellForeColor = vbRed
Else
    fgBIAMVT.CellForeColor = vbBlue
End If

fgBIAMVT.Col = 4: fgBIAMVT.Text = xYBIAMVTH.LIBELLIB1 & xYBIAMVTH.LIBELLIB2 & xYBIAMVTH.LIBELLIB3 & xYBIAMVTH.LIBELLIB4
fgBIAMVT.Col = 5: fgBIAMVT.Text = dateImp10_S(xYDOSMVT0.DOSMVTDTR)
fgBIAMVT.Col = 6: fgBIAMVT.Text = Format$(xYDOSMVT0.DOSMVTPIE, "##### ##0") & "-" & Format$(xYDOSMVT0.DOSMVTECR, "### ##0")
fgBIAMVT.Col = fgBIAMVT_arrIndex: fgBIAMVT.Text = lIndex
End Sub


Public Sub fgCPTPIE_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wFontSize As Integer
Dim blnFontBold As Boolean

On Error Resume Next
If oldYDOSMVT0.DOSMVTECR = xYCPTPIEH.MOUVEMECR Then
    blnFontBold = True
    wFontSize = 6
Else
    blnFontBold = False
    wFontSize = 7
End If

fgCPTPIE.Col = 0: fgCPTPIE.Text = Trim(xYCPTPIEH.MOUVEMCOM)
fgCPTPIE.CellFontBold = blnFontBold
If xYCPTPIEH.MOUVEMMON > 0 Then
    fgCPTPIE.Col = 1: fgCPTPIE.Text = Format$(xYCPTPIEH.MOUVEMMON, "### ### ### ##0.00")
    fgCPTPIE.CellForeColor = vbRed
Else
    fgCPTPIE.Col = 2: fgCPTPIE.Text = Format$(Abs(xYCPTPIEH.MOUVEMMON), "### ### ### ##0.00")
    fgCPTPIE.CellForeColor = vbBlue
End If

fgCPTPIE.Col = 3: fgCPTPIE.Text = Trim(rsSab("COMPTEINT"))
fgCPTPIE.CellFontBold = blnFontBold
fgCPTPIE.CellFontSize = wFontSize
fgCPTPIE.Col = 4: fgCPTPIE.Text = Trim(xYCPTPIEH.LIBELLIB1) & Trim(xYCPTPIEH.LIBELLIB2) & Trim(xYCPTPIEH.LIBELLIB3) & Trim(xYCPTPIEH.LIBELLIB4)
fgCPTPIE.Col = 5: fgCPTPIE.Text = dateIBM10(xYCPTPIEH.MOUVEMDTR, True)
fgCPTPIE.Col = 6: fgCPTPIE.Text = Format$(xYCPTPIEH.MOUVEMPIE, "##### ##0") & "-" & Format$(xYCPTPIEH.MOUVEMECR, "### ##0")
fgCPTPIE.Col = 7: fgCPTPIE.Text = xYCPTPIEH.MOUVEMANU

fgCPTPIE.Col = fgCPTPIE_arrIndex: fgCPTPIE.Text = lIndex
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

Public Sub fgdetail_Sort()
If fgDetail.Rows > 1 Then
    fgDetail.Row = 1
    fgDetail.RowSel = fgDetail.Rows - 1
    
    If fgDetail_Sort1_Old = fgDetail_Sort1 Then
        If fgDetail_SortAD = 5 Then
            fgDetail_SortAD = 6
        Else
            fgDetail_SortAD = 5
        End If
    Else
        fgDetail_SortAD = 5
    End If
    fgDetail_Sort1_Old = fgDetail_Sort1
    
    fgDetail.Col = fgDetail_Sort1
    fgDetail.ColSel = fgDetail_Sort2
    fgDetail.Sort = fgDetail_SortAD
End If

End Sub


Public Sub fgBIAMVT_Sort()
If fgBIAMVT.Rows > 1 Then
    fgBIAMVT.Row = 1
    fgBIAMVT.RowSel = fgBIAMVT.Rows - 1
    
    If fgBIAMVT_Sort1_Old = fgBIAMVT_Sort1 Then
        If fgBIAMVT_SortAD = 5 Then
            fgBIAMVT_SortAD = 6
        Else
            fgBIAMVT_SortAD = 5
        End If
    Else
        fgBIAMVT_SortAD = 5
    End If
    fgBIAMVT_Sort1_Old = fgBIAMVT_Sort1
    
    fgBIAMVT.Col = fgBIAMVT_Sort1
    fgBIAMVT.ColSel = fgBIAMVT_Sort2
    fgBIAMVT.Sort = fgBIAMVT_SortAD
End If

End Sub


Public Sub fgCPTPIE_Sort()
If fgCPTPIE.Rows > 1 Then
    fgCPTPIE.Row = 1
    fgCPTPIE.RowSel = fgCPTPIE.Rows - 1
    
    If fgCPTPIE_Sort1_Old = fgCPTPIE_Sort1 Then
        If fgCPTPIE_SortAD = 5 Then
            fgCPTPIE_SortAD = 6
        Else
            fgCPTPIE_SortAD = 5
        End If
    Else
        fgCPTPIE_SortAD = 5
    End If
    fgCPTPIE_Sort1_Old = fgCPTPIE_Sort1
    
    fgCPTPIE.Col = fgCPTPIE_Sort1
    fgCPTPIE.ColSel = fgCPTPIE_Sort2
    fgCPTPIE.Sort = fgCPTPIE_SortAD
End If

End Sub



Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    wIndex = Val(fgSelect.Text)
    Select Case lK
    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I

fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub

Public Sub fglog_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgLOG.Rows - 1
    fgLOG.Row = I
    fgLOG.Col = lK
    'wIndex = Val(fgLOG.Text)
    Select Case lK
        Case 0: Call dateJMA_AMJ(fgLOG.Text, X): fgLOG.Col = 1: X = X & Trim(fgLOG.Text)
        Case 5:  X = Format$(CCur(fgLOG.Text), "0000000000000.00")
        Case 6:  X = Trim(fgLOG.Text)
                fgLOG.Col = 5: X = X & Format$(CCur(fgLOG.Text), "0000000000000.00")
        Case 9: Call dateJMA_AMJ(fgLOG.Text, X)
        Case 10: X = Format$(Val(fgLOG.Text), "0000000000000")
    End Select
    fgLOG.Col = fgLog_arrIndex - 1
    fgLOG.Text = X
Next I

fgLog_Sort1 = fgLog_arrIndex - 1: fgLog_Sort2 = fgLog_arrIndex - 1
fgLog_Sort
End Sub


Public Sub fgLog_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgLOG.Visible = False
mRow = fgLOG.Row

If lRow > 0 And lRow < fgLOG.Rows Then
    fgLOG.Row = lRow
    For I = 0 To fgLOG.Cols - 1 's  Step -1
        fgLOG.Col = I: fgLOG.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgLOG.Row = mRow
    If fgLOG.Row > 0 Then
        lRow = fgLOG.Row
        fgLOG.Col = fgLog_arrIndex
        lColor_Old = fgLOG.CellBackColor
        For I = 0 To fgLOG.Cols - 1 ' Step -1
          fgLOG.Col = I: fgLOG.CellBackColor = lColor
        Next I
    End If
End If
fgLOG.LeftCol = fgLOG.FixedCols
fgLOG.Visible = True
End Sub
Public Sub fgLog_Reset()
fgLOG.Clear
fgLog_Sort1 = 0: fgLog_Sort2 = 0
fgLog_Sort1_Old = -1
fgLog_RowDisplay = 0: fgLog_RowClick = 0
fgLog_arrIndex = fgLOG.Cols - 1
blnfgLog_DisplayLine = False
fgLog_SortAD = 6
fgLOG.LeftCol = fgLOG.FixedCols

End Sub
Public Sub fgX_Reset()
fgX.Clear
fgX_Sort1 = 0: fgX_Sort2 = 0
fgX_Sort1_Old = -1
fgX_RowDisplay = 0: fgX_RowClick = 0
fgX_arrIndex = fgX.Cols - 1
blnfgX_DisplayLine = False
fgX_SortAD = 6
fgX.LeftCol = fgX.FixedCols

End Sub

Public Sub fgLog_Sort()
If fgLOG.Rows > 1 Then
    fgLOG.Row = 1
    fgLOG.RowSel = fgLOG.Rows - 1
    
    If fgLog_Sort1_Old = fgLog_Sort1 Then
        If fgLog_SortAD = 5 Then
            fgLog_SortAD = 6
        Else
            fgLog_SortAD = 5
        End If
    Else
        fgLog_SortAD = 5
    End If
    fgLog_Sort1_Old = fgLog_Sort1
    
    fgLOG.Col = fgLog_Sort1
    fgLOG.ColSel = fgLog_Sort2
    fgLOG.Sort = fgLog_SortAD
End If

End Sub


Public Sub fgX_Sort()
If fgX.Rows > 1 Then
    fgX.Row = 1
    fgX.RowSel = fgX.Rows - 1
    
    If fgX_Sort1_Old = fgX_Sort1 Then
        If fgX_SortAD = 5 Then
            fgX_SortAD = 6
        Else
            fgX_SortAD = 5
        End If
    Else
        fgX_SortAD = 5
    End If
    fgX_Sort1_Old = fgX_Sort1
    
    fgX.Col = fgX_Sort1
    fgX.ColSel = fgX_Sort2
    fgX.Sort = fgX_SortAD
End If

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
'$BIA_VB_HAB Call BiaPgmAut_Init(wFct, SAB_Dossier_Aut)
Call BIA_VB_HAB(wFct, arrHab(), cboSelect_SQL)

'blnSetfocus = True
Form_Init


Select Case wFct

    Case "@CDO_SCAN": blnAuto = True
        cmdSelect_SQL_Scan_Importation
        Unload Me

    Case "@SAB_DOSSIER": blnAuto = True
        Me.Enabled = False: Me.MousePointer = vbHourglass
        
        cmdSelect_SQL_K = "5 ECNFPT_CD7"
        cmdSelect_SQL_5_ECNFPT_CD7
        '$JPL 2015-09-13 cmdSelect_SQL_5_ECNFPT_CD7 màj YDOSCD70 pour les commissions ECNFPT

        cmdSelect_SQL_Surveillance
        '$JPL 2014-06-12 cmdSendMail_SAB_Dossier ""
        
        If Mid$(YBIATAB0_DATE_CPT_J, 1, 6) <> Mid$(YBIATAB0_DATE_CPT_JS1, 1, 6) Then
        '_________________________________________________________________________
            txtSelect_6_PCI = "13221"
            txtSelect_6_CLIEANCLI = ""
            wAmjMin = dateElp("Jour", 1, YBIATAB0_DATE_CPT_MP1)
            Call DTPicker_Set(txtSelect_6_AMJMin, wAmjMin) '

            cmdSelect_SQL_6
            cmdSendMail_SAB_Dossier "CDO_SQL_6", "BIA-CDO-Prov-13221"
        '_________________________________________________________________________
            txtSelect_6_PCI = "25302"
            txtSelect_6_CLIEANCLI = ""
            wAmjMin = dateElp("Jour", 1, YBIATAB0_DATE_CPT_MP1)
            Call DTPicker_Set(txtSelect_6_AMJMin, wAmjMin) '

            cmdSelect_SQL_6
            cmdSendMail_SAB_Dossier "CDO_SQL_6", "BIA-CDO-Prov-25302"
        '_________________________________________________________________________
            wAmjMin = dateElp("Jour", 1, YBIATAB0_DATE_CPT_J)
            Call DTPicker_Set(txtSelect_DOSCD7DAN, wAmjMin) '

            cmdSelect_SQL_Xc
            cmdSendMail_SAB_Dossier "", "BIA-CDO-Commissions"
         '_________________________________________________________________________
 '$JPL 2015-12-09
            wAmjMin = dateElp("Jour", 1, YBIATAB0_DATE_CPT_J)
            Call DTPicker_Set(txtSelect_DOSCD7DAN, wAmjMin) '

            cmdSelect_SQL_XE1an
            cmdSendMail_SAB_Dossier "CDO_SQL_XE1", "BIA-CDO-Engagement"
       '_________________________________________________________________________
            wAmjMin = dateElp("Jour", 1, YBIATAB0_DATE_CPT_J)
            Call DTPicker_Set(txtSelect_DOSCD7DAN, wAmjMin) '

            cmdSelect_SQL_XE1an
            cmdSendMail_SAB_Dossier "CDO_SQL_XE1", "BIA-CDO-Engagement1an"
        '_________________________________________________________________________
            
        End If
        
        Dim X As String, xDest As String
        mailAdresse_Production_Load
        
        Call cmdSelect_SQL_5réfext("Jour")
        If fgSelect.Rows > 1 Then
            'Call mailAdresse_Production_Control("CHIBAB;CHARTIER;REOL_CH;OURY", xDest)
             xDest = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire("S10")

            X = "Liste des dossiers CDO créés le " & dateImp10_S(YBIATAB0_DATE_CPT_J) _
              & ",<BR> dont la référence externe est identique à celle de dossiers non clos."
            Call MSFlexGrid_SendMail(xDest, "CDO_Doublon", "Liste des nouveaux dossiers CDO 'doublons' - " & dateImp10_S(YBIATAB0_DATE_CPT_J), X, fgSelect, 4)
        End If

        
        Me.Enabled = True: Me.MousePointer = 0
        Unload Me

    Case Else: blnAuto = False
End Select

End Sub


Public Sub Form_Init()
Dim V, xSql As String, X As String
Dim K As Long

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True

If arrMT_Nb = 0 Then arrMT_Load

cmdReset
blnControl = False
libDOSSLDDEV.ForeColor = vbBlack
libDOSSLDOPE.ForeColor = vbBlack
libDOSSLDNUM.ForeColor = vbBlack

fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False
fgSelect_Width = fgSelect.Width
fgSelect_Height = fgSelect.Height
fgSelect_BackColorFixed = fgSelect.BackColorFixed
fgSelect_ForeColorFixed = fgSelect.ForeColorFixed
fgSelect_ForeColor = fgSelect.ForeColor
fgSelect_BackColor = fgSelect.BackColor


fgLOG.Visible = False
Set fgLOG.Container = fraTab0
fgLOG.Top = fgSelect.Top
fgLOG.Left = fgSelect.Left
fgLOG.Width = fgBIAMVT.Width
fgLOG.Height = fraTab0.Height - fgSelect.Top
fgLog_FormatString = fgLOG.FormatString
fgLOG.Visible = False

fraSelect_Options_1.BorderStyle = 0

lstW.Visible = False

fgDetail.Visible = False: fraDetail.Visible = False
fgDetail_FormatString = fgDetail.FormatString

SSTab2.Tab = 0
SSTab2.Visible = False
fgBIAMVT_FormatString = fgBIAMVT.FormatString

fgDossier_FormatString = fgDossier.FormatString
fgCOM_FormatString = fgCOM.FormatString

Set fgCPTPIE.Container = SSTab2
fgCPTPIE.Top = fgBIAMVT.Top
fgCPTPIE.Height = fgBIAMVT.Height
fgCPTPIE.Left = 2000
fgCPTPIE.Visible = False
fgCPTPIE_FormatString = fgCPTPIE.FormatString

fraCompte.Visible = False
Set fraCompte.Container = fraTab0
fraCompte.Top = 1440
fraCompte.Left = fraTab0.Left + fraTab0.Width - fraCompte.Width

fgX.Visible = False
Set fgX.Container = fraTab0
fgX.Top = fgSelect.Top
fgX.Left = fgSelect.Left
fgX.Height = 9500
fgX.Width = 15540


fraYDOSXOD0.Visible = False
Set fraYDOSXOD0.Container = fraTab0
fraYDOSXOD0.Top = 3000
fraYDOSXOD0.Left = fraTab0.Left + fraTab0.Width - fraYDOSXOD0.Width - 200
fraYDOSXOD0.ForeColor = vbRed

fraECNFPT.Visible = False
lblECNFPT_TOT_X.ForeColor = mColor_Z0
libECNFPT_TOT_X.ForeColor = mColor_Z0


'$BIA_VB_HAB
'cboSelect_SQL.Clear
'cboSelect_SQL.AddItem "1  - sélection / dossier"
'cboSelect_SQL.AddItem "2  - sélection / client"
'cboSelect_SQL.AddItem "3  - Evénement en attente de validation"
'cboSelect_SQL.AddItem "5  - sélection / code état"
'If SAB_Dossier_Aut.Rapprocher Then
'    cboSelect_SQL.AddItem "6  - Etat PCI-Client-Dossier"
'    cboSelect_SQL.AddItem "X#  - Etat de surveillance (.xls) "
'    cboSelect_SQL.AddItem "Xc - Commissions à recevoir (.xls) "
'End If
'If SAB_Dossier_Aut.Comptabiliser Then
'    cboSelect_SQL.AddItem "Xi - engagements Intragroupe (.xls) "
'    cboSelect_SQL.AddItem "XE1an - engagements +/- 1 an (.xls) "
    
'End If

'If SAB_Dossier_Aut.Valider Then
'    cboSelect_SQL.AddItem "2# - sélection / client + Mise à jour"
'    cboSelect_SQL.AddItem "zOD - Liste des OD"
'    cboSelect_SQL.AddItem "zSD - Liste des ajustements de soldes CPT / GES"
'End If
'$BIA_VB_HAB
If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0
cmdSelect_SQL_K = "1"


fgCourrier_FormatString = fgCourrier.FormatString
fgScan_FormatString = fgScan.FormatString

lstW.Clear




'Initialisation opération________________________________________________________________________________
arrOPE_Nb = 0
ReDim Preserve arrOPE(1000)

cboSelect_DOSSLDOPE.Clear
cboSelect_DOSSLDOPE.AddItem "CDE"
cboSelect_DOSSLDOPE.AddItem "CDI"
cboSelect_DOSSLDOPE.AddItem "RDE"
cboSelect_DOSSLDOPE.AddItem "RDI"
cboSelect_DOSSLDOPE.AddItem "ENG"
cboSelect_DOSSLDOPE.AddItem "GAR"
cboSelect_DOSSLDOPE.AddItem ""
xSql = "select distinct DOSSLDOPE from " & paramIBM_Library_SABSPE & ".YDOSSLD0 order by DOSSLDOPE"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    arrOPE_Nb = arrOPE_Nb + 1
    arrOPE(arrOPE_Nb) = Trim(rsSab("DOSSLDOPE"))
    cboSelect_DOSSLDOPE.AddItem Trim(rsSab("DOSSLDOPE"))
    rsSab.MoveNext
Loop
ReDim Preserve arrOPE(arrOPE_Nb + 1)
Call cbo_Scan("CDE", cboSelect_DOSSLDOPE)

'Initialisation devise________________________________________________________________________________
arrDev_Nb = 0
ReDim Preserve arrDev(1000)

cboSelect_DOSSLDDEV.Clear
cboSelect_DOSSLDDEV.AddItem ""
xSql = "select distinct DOSSLDDEV from " & paramIBM_Library_SABSPE & ".YDOSSLD0 order by DOSSLDDEV"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    arrDev_Nb = arrDev_Nb + 1
    arrDev(arrDev_Nb) = Trim(rsSab("DOSSLDDEV"))
    cboSelect_DOSSLDDEV.AddItem Trim(rsSab("DOSSLDDEV"))
    rsSab.MoveNext
Loop
ReDim Preserve arrDev(arrDev_Nb + 1)
ReDim arrDev_RowT(arrDev_Nb + 1)
'Initialisation PCI_______________________________________________________________________________

cboSelect_DOSSLDPCI.Clear
cboSelect_DOSSLDPCI.AddItem ""
xSql = "select distinct(DOSSLDPCI) from " & paramIBM_Library_SABSPE & ".YDOSSLD0"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_DOSSLDPCI.AddItem Trim(rsSab("DOSSLDPCI"))
    rsSab.MoveNext
Loop


'Initialisation code état___________________________________________________________________________

fraSelect_Options_5.Visible = False
Set fraSelect_Options_5.Container = fraTab0
fraSelect_Options_5.Top = fraSelect_Options.Top
fraSelect_Options_5.Left = fraSelect_Options.Left
fraSelect_Options_5.Width = fraSelect_Options.Width
cboSelect_DOSSLDSTA.Clear
xSql = "select distinct DOSSLDSTA , DOSSLDSVC from " & paramIBM_Library_SABSPE & ".YDOSSLD0 group by DOSSLDSTA , DOSSLDSVC"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    If Trim(rsSab(0)) <> "" Then cboSelect_DOSSLDSTA.AddItem rsSab(0) & " - " & rsSab(1)
    rsSab.MoveNext
Loop
If cboSelect_DOSSLDSTA.ListCount > 0 Then cboSelect_DOSSLDSTA.ListIndex = 0
'___________________________________________________________________________

fraSelect_Options_3uti.Visible = False
Set fraSelect_Options_3uti.Container = fraTab0
fraSelect_Options_3uti.Top = fraSelect_Options.Top
fraSelect_Options_3uti.Left = fraSelect_Options.Left
fraSelect_Options_3uti.Width = fraSelect_Options.Width

Call DTPicker_Set(txtSelect_Options_3uti_AmjMin, "20000101") '
Call DTPicker_Set(txtSelect_Options_3uti_AmjMax, DSys) '

'___________________________________________________________________________

fraSelect_Options_6.Visible = False
Set fraSelect_Options_6.Container = fraTab0
fraSelect_Options_6.Top = fraSelect_Options.Top
fraSelect_Options_6.Left = fraSelect_Options.Left
fraSelect_Options_6.Width = fraSelect_Options.Width

wAmjMin = dateElp("Jour", 1, YBIATAB0_DATE_CPT_MP1)
Call DTPicker_Set(txtSelect_6_AMJMin, wAmjMin) '
libSelect_6_AMJMax = "au " & dateImp10_S(YBIATAB0_DATE_CPT_J)

'___________________________________________________________________________

fraSelect_Options_Xc.Visible = False
Set fraSelect_Options_Xc.Container = fraTab0
fraSelect_Options_Xc.Top = fraSelect_Options.Top
fraSelect_Options_Xc.Left = fraSelect_Options.Left
fraSelect_Options_Xc.Width = fraSelect_Options.Width

wAmjMin = dateElp("Jour", 1, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtSelect_DOSCD7DAN, wAmjMin) '

xSql = "select distinct DOSCD7DSIT  from " & paramIBM_Library_SABSPE & ".YDOSCD70 order by DOSCD7DSIT desc"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    If Trim(rsSab(0)) <> "" Then cboSelect_DOSCD7DSIT.AddItem dateImp_Amj(rsSab(0))
    rsSab.MoveNext
Loop
If cboSelect_DOSCD7DSIT.ListCount > 0 Then cboSelect_DOSCD7DSIT.ListIndex = 0
'___________________________________________________________________________

fraSelect_Options_Log.Visible = False
Set fraSelect_Options_Log.Container = fraTab0
fraSelect_Options_Log.Top = fraSelect_Options.Top
fraSelect_Options_Log.Left = fraSelect_Options.Left
fraSelect_Options_Log.Width = fraSelect_Options.Width

Call DTPicker_Set(txtSelect_Options_Log_AmjMin, YBIATAB0_DATE_CPT_J) '
Call DTPicker_Set(txtSelect_Options_Log_AmjMax, YBIATAB0_DATE_CPT_J) '
'___________________________________________________________________________

fraSelect_Options_Scan_Liste.Visible = False
Set fraSelect_Options_Scan_Liste.Container = fraTab0
fraSelect_Options_Scan_Liste.Top = fraSelect_Options.Top
fraSelect_Options_Scan_Liste.Left = fraSelect_Options.Left
fraSelect_Options_Scan_Liste.Width = fraSelect_Options.Width

Call DTPicker_Set(txtSelect_Options_Scan_Liste_AMJ, YBIATAB0_DATE_CPT_J) '

'___________________________________________________________________________
fgYSWISAB0_FormatString = fgYSWISAB0.FormatString

fgSwift_FormatString = fgSwift.FormatString
fgSwift.ForeColor = vbBlue
fraSwift.Visible = False
Set fraSwift.Container = fraTab0
fraSwift.Top = fraDetail.Top
fraSwift.Left = fraTab0.Left + fraTab0.Width - fraSwift.Width - 200
'fraSwift.Height = 7300
blnSIDE_DB = False

'___________________________________________________________________________
fgDossier_FormatString = fgDossier.FormatString

'___________________________________________________________________________
fraSelect_Options.Visible = True
blnControl = True

Me.Enabled = True
'Denis ROSILLETTE le 23/10/2012
'txtSelect_DOSSLDNUM.SetFocus
If Me.Visible Then
    txtSelect_DOSSLDNUM.SetFocus
End If
End Sub

Public Sub fgScan_Display()
Dim wColor As Long, X As String
Dim objFolder, objFiles
Dim fsoFile As File

On Error GoTo Error_Handler
currentAction = "fgScan_Display"
SSTab2.Visible = False
fgScan.Visible = False
fgScan_Reset

fgScan.Rows = 1
fgScan.FormatString = fgScan_FormatString
fgScan.Row = 0


X = "_Scan\" & xYDOSSLD0.DOSSLDOPE & "_" & Format(xYDOSSLD0.DOSSLDNUM, "000000") & "\"
If Dir(paramCDO_Dossier_Path_DROPI & X) <> "" Then
    Set objFolder = msFileSystem.GetFolder(paramCDO_Dossier_Path_DROPI & X)
    Set objFiles = objFolder.Files
    For Each fsoFile In objFiles
        '$JPL 2014-11-27 If InStr(fsoFile.Type, "Document") > 0 Then
            fgScan.Rows = fgScan.Rows + 1
            fgScan.Row = fgScan.Rows - 1
            fgScan.Col = 0: fgScan.Text = fsoFile.DateCreated
            fgScan.Col = 1: fgScan.Text = fsoFile.Name
        '$JPL 2014-11-27 End If
    Next
End If
'__________________________________________________________________________________________________
SSTab2.Visible = True
fgScan.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub fgX_Display()
Dim wColor As Long, X As String
Dim objFolder, objFiles
Dim fsoFile As File
Dim newMontant As Double

On Error GoTo Error_Handler
currentAction = "fgX_Display"
fgX.Visible = False
fgX_Reset

fgX.Rows = 1
fgX.FormatString = "<Nature    |<Dossier|<Client       |<Intitulé                                                                        " _
                 & " |<Montant           |<Devise " _
                 & "|<Début             |<Fin                |<Référence acte              |< Référence Client                               "
fgX.Row = 0
fgX.Col = 1: fgX.CellAlignment = 1
fgX.Col = 4: fgX.CellAlignment = 1

Do While Not rsSab.EOF

    fgX.Rows = fgX.Rows + 1
    fgX.Row = fgX.Rows - 1
    'fgX_DisplayLine I
    fgX.Col = 1: fgX.Text = Format$(rsSab("CAUDOSDOS"), "### ##0")
    fgX.Col = 0: fgX.Text = rsSab("CAUDOSCAU")
    fgX.Col = 2: fgX.Text = rsSab("CAUDOSBEN")
    fgX.Col = 3: fgX.Text = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA1"))
    'Voir si le montant de l'encours a été bougé (encours restant)
    newMontant = 0
    If retourne_encours_restant(rsSab("CAUDOSDOS"), rsSab("CAUDOSBEN"), rsSab("CAUDOSETB"), newMontant) = False Then
        newMontant = CDbl(rsSab("CAUDOSMNT"))
    End If
    '
    fgX.Col = 4: fgX.Text = Format$(newMontant, "### ### ### ##0.00")
    fgX.Col = 5: fgX.Text = rsSab("CAUDOSDEV")
    fgX.Col = 6: fgX.Text = dateImp_Amj(rsSab("CAUDOSDEB") + 19000000)
    If rsSab("CAUDOSFIN") > 0 Then fgX.Col = 7: fgX.Text = dateImp_Amj(rsSab("CAUDOSFIN") + 19000000)
    fgX.Col = 8: fgX.Text = rsSab("CAUDOSACT")
    fgX.Col = 9: fgX.Text = rsSab("CAUDOSREF")
    rsSab.MoveNext

Loop
'__________________________________________________________________________________________________
SSTab2.Visible = True
fgX.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Function retourne_encours_restant(zDossier As Long, zClient As String, zEtb As Long, ByRef zMontant As Double) As Boolean
Dim xSql As String
Dim rs As ADODB.Recordset

    retourne_encours_restant = False
    zMontant = 0
    xSql = "SELECT CAUAGARES FROM " & paramIBM_Library_SAB & ".ZCAUAGA0 WHERE CAUAGACLI = '" & zClient & "'"
    xSql = xSql & " AND CAUAGADOS = " & zDossier & " and CAUAGAETB = " & zEtb
    xSql = xSql & " ORDER BY CAUAGADAT DESC"
    Set rs = cnsab.Execute(xSql)
    If Not rs.EOF Then
        zMontant = CDbl(rs("CAUAGARES"))
        retourne_encours_restant = True
    End If
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing

End Function


Public Sub fgScan_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
SSTab2.Visible = False: fraDetail.Visible = False
mRow = fgScan.Row

If lRow > 0 And lRow < fgScan.Rows Then
    fgScan.Row = lRow
    For I = fgScan_arrIndex To fgScan.FixedCols Step -1
        fgScan.Col = I: fgScan.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgScan.Row = mRow
    If fgScan.Row > 0 Then
        lRow = fgScan.Row
        fgScan.Col = fgScan_arrIndex
        lColor_Old = fgScan.CellBackColor
        For I = fgScan_arrIndex To fgScan.FixedCols Step -1
          fgScan.Col = I: fgScan.CellBackColor = lColor
        Next I
    End If
End If
fgScan.LeftCol = fgScan.FixedCols
SSTab2.Visible = True: fraDetail.Visible = True
End Sub
Public Sub fgScan_Reset()
fgScan.Clear
fgScan_Sort1 = 0: fgScan_Sort2 = 0
fgScan_Sort1_Old = -1
fgScan_RowDisplay = 0: fgScan_RowClick = 0
fgScan_arrIndex = fgScan.Cols - 1
blnfgScan_DisplayLine = False
fgScan_SortAD = 6
fgScan.LeftCol = fgScan.FixedCols

End Sub

Public Sub fgScan_Sort()
If fgScan.Rows > 1 Then
    fgScan.Row = 1
    fgScan.RowSel = fgScan.Rows - 1
    
    If fgScan_Sort1_Old = fgScan_Sort1 Then
        If fgScan_SortAD = 5 Then
            fgScan_SortAD = 6
        Else
            fgScan_SortAD = 5
        End If
    Else
        fgScan_SortAD = 5
    End If
    fgScan_Sort1_Old = fgScan_Sort1
    
    fgScan.Col = fgScan_Sort1
    fgScan.ColSel = fgScan_Sort2
    fgScan.Sort = fgScan_SortAD
End If

End Sub




Private Sub fgCOM_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)


If cmdSelect_SQL_K = 1 And fgCOM.Row = mECNFPT_Row Then
    If mECNFPT_MTA = 0 Then
        Call MsgBox("paramétres de calcul indéterminés", vbExclamation, "Commission prorata temporis")
    Else
        txtECNFPT_MTA = Format$(mECNFPT_MTA / 100, "### ### ##0.00")
        txtECNFPT_TX1 = Format$(mECNFPT_TX1, "##0.00000")
        txtECNFPT_PER = mECNFPT_PER
        txtECNFPT_NBJ = Format$(mECNFPT_NBJ, "##0")
        txtECNFPT_DDEB = dateImp10(mECNFPT_DDEB) & "  -  " & dateImp10(mECNFPT_DFIN)
        txtECNFPT_MON = Format$(mECNFPT_MON, "### ### ##0.00")
        
        libECNFPT_TOT_X = Format$(mECNFPT_TOT, "### ### ##0.00")
    
        libECNFPT_MON_X = ""
        libECNFPT_NBJ_X = ""
        Call DTPicker_Set(txtECNFPT_DREG, CStr(mECNFPT_DFIN))
        txtECNFPT_DREG_Change
        fraECNFPT.Visible = True
    End If
Else
    fraECNFPT.Visible = False
End If
End Sub


Private Sub fgScan_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
If y <= fgScan.RowHeightMin Then
Else
    If fgScan.Rows > 1 Then
        Call fgScan_Color(fgScan_RowClick, MouseMoveUsr.BackColor, fgScan_ColorClick)
        fgScan.Col = 1
        mDoc_Filename = "_SCAN\" & xYDOSSLD0.DOSSLDOPE & "_" & Format(xYDOSSLD0.DOSSLDNUM, "000000") & "\" & fgScan.Text
      If arrHab(16) Then
            mnuDoc_Rename.Visible = True
            Me.PopupMenu mnuDoc, vbPopupMenuLeftButton
        Else
            mnuDoc_Display_Click
        End If
   End If
End If
Me.Enabled = True: Me.MousePointer = 0
'Windows_Display_File
End Sub

Private Sub fgYSWISAB0_Display()
Dim xSql As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgYSWISAB0.Visible = False
fgYSWISAB0_Reset

fgYSWISAB0.Rows = 1
fgYSWISAB0.FormatString = fgYSWISAB0_FormatString
fgYSWISAB0.Row = 0

currentAction = "fgYSWISAB0_Display"
mYSWILNK0_Display = ""
'__________________________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABOPEC = '" & xYDOSSLD0.DOSSLDOPE & "'" _
     & " and   SWISABOPEN = " & xYDOSSLD0.DOSSLDNUM _
     & " order by SWISABWAMJ , SWISABWHMS"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF

    fgYSWISAB0.Rows = fgYSWISAB0.Rows + 1
    fgYSWISAB0.Row = fgYSWISAB0.Rows - 1
    fgYSWISAB0_DisplayLine I

    If rsSab("SWISABXGOS") = "G" Or rsSab("SWISABXEVE") = "G" Then
        mYSWILNK0_Display = mYSWILNK0_Display & " " & rsSab("SWISABSWID") & " ,"
    End If
        
    rsSab.MoveNext

Loop
         
If mYSWILNK0_Display <> "" Then
    fgYSWISAB0_Display_YSWISAB0_YSWILNK0
End If
'DR 20/03/2014
fgYSWISAB0_Sort1 = 12: fgYSWISAB0_Sort2 = 12: fgYSWISAB0_Sort

fgYSWISAB0.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgYSWISAB0.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub fgYSWISAB0_Display_YSWISAB0_YSWILNK0()
On Error GoTo Error_Handler
Dim X As String

Mid$(mYSWILNK0_Display, Len(mYSWILNK0_Display), 1) = ")"
X = "select distinct SWILNKAPPN from " & paramIBM_Library_SABSPE & ".YSWILNK0 " _
    & " where SWILNKAPPC = 'GOS' and SWILNKSTA = '' and SWILNKSWID in (" & mYSWILNK0_Display
Set rsSabX = cnsab.Execute(X)


Do While Not rsSabX.EOF
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
        & " where SWISABOPEC = 'GOS' and SWISABOPEN = " & rsSabX("SWILNKAPPN") _
        & " order by SWISABSWID"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
    
        fgYSWISAB0.Rows = fgYSWISAB0.Rows + 1
        fgYSWISAB0.Row = fgYSWISAB0.Rows - 1
        fgYSWISAB0_DisplayLine 0
        
        rsSab.MoveNext
    
    Loop

    
    rsSabX.MoveNext
Loop

If fgYSWISAB0.Rows > 2 Then fgYSWISAB0_Sort1 = 11: fgYSWISAB0_Sort2 = 11: fgYSWISAB0_Sort

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub fgYSWISAB0_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long
Dim X As String, X2 As String

On Error Resume Next

If rsSab("SWISABWES") = "S" Then
    X = rsSab("SWISABWMTK") & " S"
    wColor = RGB(16, 96, 16)
Else
    X = rsSab("SWISABWMTK") & " E"
    wColor = vbBlue
End If

  

fgYSWISAB0.Col = 0: fgYSWISAB0.Text = X
fgYSWISAB0.CellForeColor = wColor

fgYSWISAB0.Col = 1: fgYSWISAB0.Text = rsSab("SWISABWBIC")
fgYSWISAB0.CellForeColor = wColor
fgYSWISAB0.Col = 2: fgYSWISAB0.Text = rsSab("SWISABWL20")
fgYSWISAB0.CellForeColor = wColor
Select Case rsSab("SWISABK20")
    Case "!": fgYSWISAB0.CellBackColor = RGB(220, 220, 255)
    Case Is <> " ": fgYSWISAB0.CellBackColor = RGB(220, 255, 220)
End Select

fgYSWISAB0.Col = 3: fgYSWISAB0.Text = Format$(CCur(rsSab("SWISABWMTD")), "### ### ### ##0.00")
fgYSWISAB0.CellForeColor = vbRed
fgYSWISAB0.CellFontBold = True
fgYSWISAB0.Col = 4: fgYSWISAB0.Text = rsSab("SWISABWDEV")
fgYSWISAB0.CellForeColor = wColor
fgYSWISAB0.Col = 5
If rsSab("SWISABKPDE") <> " " Then
    fgYSWISAB0.Text = rsSab("SWISABKPDE") & " PDE"
    Select Case rsSab("SWISABK20")
        Case "!", "?": fgYSWISAB0.CellBackColor = RGB(255, 220, 220)
        Case Else: fgYSWISAB0.CellBackColor = RGB(220, 255, 220)
    End Select
Else
    If rsSab("SWISABK20") <> " " Then fgYSWISAB0.Text = rsSab("SWISABK20") & " réf ="
End If
    fgYSWISAB0.CellForeColor = wColor
fgYSWISAB0.Col = 6: fgYSWISAB0.Text = dateImp10_S(rsSab("SWISABWAMJ")) & " " & timeImp8(rsSab("SWISABWHMS"))
fgYSWISAB0.CellForeColor = RGB(80, 80, 80)

fgYSWISAB0.Col = fgYSWISAB0_arrIndex: fgYSWISAB0.Text = lIndex
fgYSWISAB0.Col = 8: fgYSWISAB0.Text = rsSab("SWISABWID1")
fgYSWISAB0.Col = 9: fgYSWISAB0.Text = rsSab("SWISABWIDL")
fgYSWISAB0.Col = 10: fgYSWISAB0.Text = rsSab("SWISABWIDH")
fgYSWISAB0.Col = 11: fgYSWISAB0.Text = rsSab("SWISABSWID")

'DR 20/03/2014
fgYSWISAB0.Col = 12: fgYSWISAB0.ColWidth(12) = 1: fgYSWISAB0.Text = rsSab("SWISABWAMJ")

If rsSab("SWISABWSTA") <> "V" Then
    For K = 0 To 11
        fgYSWISAB0.Col = K
        fgYSWISAB0.CellBackColor = mColor_W1
    Next K
End If



fgYSWISAB0.Col = 7
K = Val(Mid$(rsSab("SWISABKSRV"), 2, 2))
'fgYSWISAB0.Text = rsSab("SWISABWN20")
fgYSWISAB0.CellForeColor = wColor

X = ""
If rsSab("SWISABXGOS") <> " " Then
    X = " # GOS"
    fgYSWISAB0.CellBackColor = vbYellow

Else
    Select Case rsSab("SWISABXEVE")
        Case " ", "="
        Case "G": X = " # EVE": fgYSWISAB0.CellBackColor = RGB(245, 222, 131)
        Case "*":
            
            If rsSab("SWISABK999") = "I" Then
                fgYSWISAB0.CellBackColor = RGB(220, 220, 220)
            'Else
                 'fgYSWISAB0.CellBackColor = mColor_B0
           End If
            
        Case Else: X = " ???": fgYSWISAB0.CellBackColor = mColor_W1
    End Select
End If
fgYSWISAB0.Text = rsSab("SWISABWN20") & X

End Sub

Public Sub fgYSWISAB0_Reset()
fgYSWISAB0.Clear
fgYSWISAB0_Sort1 = 0: fgYSWISAB0_Sort2 = 0
fgYSWISAB0_Sort1_Old = -1
fgYSWISAB0_RowDisplay = 0: fgYSWISAB0_RowClick = 0
fgYSWISAB0_arrIndex = fgYSWISAB0.Cols - 1
blnfgYSWISAB0_DisplayLine = False
fgYSWISAB0_SortAD = 6
fgYSWISAB0.LeftCol = fgYSWISAB0.FixedCols

End Sub

Public Sub fgYSWISAB0_Sort()
If fgYSWISAB0.Rows > 1 Then
    fgYSWISAB0.Row = 1
    fgYSWISAB0.RowSel = fgYSWISAB0.Rows - 1
    
    If fgYSWISAB0_Sort1_Old = fgYSWISAB0_Sort1 Then
        If fgYSWISAB0_SortAD = 5 Then
            fgYSWISAB0_SortAD = 6
        Else
            fgYSWISAB0_SortAD = 5
        End If
    Else
        fgYSWISAB0_SortAD = 5
    End If
    fgYSWISAB0_Sort1_Old = fgYSWISAB0_Sort1
    
    fgYSWISAB0.Col = fgYSWISAB0_Sort1
    fgYSWISAB0.ColSel = fgYSWISAB0_Sort2
    fgYSWISAB0.Sort = fgYSWISAB0_SortAD
End If

End Sub

Public Sub fgYSWISAB0_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgYSWISAB0.Rows - 1
    fgYSWISAB0.Row = I
    fgYSWISAB0.Col = lK
    Select Case lK
        Case 3: fgYSWISAB0.Col = 3: X = Format$(Val(fgYSWISAB0.Text), "000000000000000.00")
        Case 4:
            fgYSWISAB0.Col = 4: X = Trim(fgYSWISAB0.Text)
            fgYSWISAB0.Col = 3: X = X & Format$(Val(fgYSWISAB0.Text), "000000000000000.00")
    End Select
    fgYSWISAB0.Col = fgYSWISAB0_arrIndex - 1
    fgYSWISAB0.Text = X
Next I

fgYSWISAB0_Sort1 = fgYSWISAB0_arrIndex - 1: fgYSWISAB0_Sort2 = fgYSWISAB0_arrIndex - 1
fgYSWISAB0_Sort
End Sub


Public Sub fgYSWISAB0_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgYSWISAB0.Visible = False
mRow = fgYSWISAB0.Row

If lRow > 0 And lRow < fgYSWISAB0.Rows Then
    fgYSWISAB0.Row = lRow
    For I = fgYSWISAB0_arrIndex To fgYSWISAB0.FixedCols Step -1
        fgYSWISAB0.Col = I: fgYSWISAB0.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgYSWISAB0.Row = mRow
    If fgYSWISAB0.Row > 0 Then
        lRow = fgYSWISAB0.Row
        lColor_Old = fgYSWISAB0.CellBackColor
        For I = fgYSWISAB0_arrIndex To fgYSWISAB0.FixedCols Step -1
          fgYSWISAB0.Col = I: fgYSWISAB0.CellBackColor = lColor
        Next I
    End If
End If
fgYSWISAB0.LeftCol = fgYSWISAB0.FixedCols
fgYSWISAB0.Visible = True
End Sub

Private Sub cboSelect_Options_3uti_CDODOSNOT_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub cboSelect_Options_3uti_UTI_Click()
cmdSelect_Clear

End Sub

Private Sub chkSelect_Options_3uti_Swift_Click()
cmdSelect_Clear

End Sub

Private Sub chkSIDE_DB_Show_Click()
On Error Resume Next
Dim K As Integer
If fraSwift.Visible = True Then
    If chkSIDE_DB_Show = "1" Then
        'K = InStr(libSWIFT_SWISABSWID, " ")
        If K > 0 Then frmSIDE_DB.fgSwift_Display Val(Mid$(libSWIFT_SWISABSWID, 1, K)), 0, 0, 0
        Call frmSIDE_DB.fgSwift_Display(oldYSWISAB0.SWISABSWID, 0, 0, 0)
    Else
        frmSIDE_DB.Hide
    End If
End If

End Sub

Private Sub cmdSAB_Dossier_CDO_Click()
Dim s() As String
Dim choix As String

    choix = ""
    s = Split(cboSelect_SQL.Text, "-")
    If UBound(s) > 0 Then
        choix = Trim(s(0))
    End If
    If LCase(choix) = "3rdo" Or LCase(choix) = "3uti" Then
        Call MsgBox("Cette fonction ne peut pas être utilisée ici !")
        Exit Sub
    End If

    If Trim(UCase(cboSelect_DOSSLDOPE.Text)) = "RDE" Or Trim(UCase(cboSelect_DOSSLDOPE.Text)) = "RDI" Then
        Call frmSAB_Dossier_RDE.Form_Init("", oldYDOSSLD0.DOSSLDOPE, oldYDOSSLD0.DOSSLDNUM)
    Else
        Call frmSAB_Dossier_CDO.Form_Init("", oldYDOSSLD0.DOSSLDOPE, oldYDOSSLD0.DOSSLDNUM)
    End If
    
End Sub

Private Sub cmdSAB_Dossier_DB_Click()
Call frmSAB_Dossier_DB.Form_Init("", "", "", "", "00", "00", oldYDOSSLD0.DOSSLDOPE, oldYDOSSLD0.DOSSLDNUM)
End Sub

Private Sub fgCourrier_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
If y <= fgCourrier.RowHeightMin Then
Else
    If fgCourrier.Rows > 1 Then
        Call fgCourrier_Color(fgCourrier_RowClick, MouseMoveUsr.BackColor, fgCourrier_ColorClick)
        fgCourrier.Col = 1
        mDoc_Filename = xYDOSSLD0.DOSSLDOPE & "_" & Format(xYDOSSLD0.DOSSLDNUM, "000000") & "\" & fgCourrier.Text

        If arrHab(16) Then
            mnuDoc_Rename.Visible = False
            Me.PopupMenu mnuDoc, vbPopupMenuLeftButton
        Else
            mnuDoc_Display_Click
        End If
   End If
End If
Me.Enabled = True: Me.MousePointer = 0
'Windows_Display_File
End Sub


Private Sub fgX_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xMsgbox As String, xSql As String, Nb As Long
On Error Resume Next


If y <= fgX.RowHeightMin Then

        Select Case fgX.Col
            Case 0: fgX_Sort1 = 0: fgX_Sort2 = 1: fgX_Sort
            Case 1:  fgX_Sort1 = 1: fgX_Sort2 = 2: fgX_Sort
            Case 2: fgX_Sort1 = 2: fgX_Sort2 = 2: fgX_Sort
            Case 3:  fgX_Sort1 = 3: fgX_Sort2 = 4: fgX_Sort
            Case 4: fgX_Sort1 = 4: fgX_Sort2 = 4: fgX_Sort
            Case 5:  fgX_Sort1 = 5: fgX_Sort2 = 5: fgX_Sort
            Case 6: fgX_Sort1 = 6: fgX_Sort2 = 6: fgX_Sort
            Case 7:  fgX_Sort1 = 7: fgX_Sort2 = 7: fgX_Sort
'            Case fgX_arrIndex:  fgX_SortX fgX_arrIndex
        End Select
Else
'    If fgX.Rows > 1 Then
'        Call fgX_Color(fgX_RowClick, MouseMoveUsr.BackColor, fgX_ColorClick)
'   End If


End If

End Sub


Private Sub fgYSWISAB0_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xMTK As String, xBIC As String
On Error Resume Next


If y <= fgYSWISAB0.RowHeightMin Then
    Select Case fgYSWISAB0.Col
        Case 0: fgYSWISAB0_Sort1 = 0: fgYSWISAB0_Sort2 = 2: fgYSWISAB0_Sort
        Case 1:  fgYSWISAB0_Sort1 = 1: fgYSWISAB0_Sort2 = 2: fgYSWISAB0_Sort
        Case 2:  fgYSWISAB0_Sort1 = 2: fgYSWISAB0_Sort2 = 2: fgYSWISAB0_Sort
        Case 3:  fgYSWISAB0_Sort1 = 3: fgYSWISAB0_Sort2 = 3: fgYSWISAB0_SortX 3
        Case 4:  fgYSWISAB0_Sort1 = 4: fgYSWISAB0_Sort2 = 4: fgYSWISAB0_SortX 4
        Case 5:  fgYSWISAB0_Sort1 = 5: fgYSWISAB0_Sort2 = 5: fgYSWISAB0_Sort
        
        'DR 20/03/2014
        'Case 6:  fgYSWISAB0_Sort1 = 6: fgYSWISAB0_Sort2 = 6: fgYSWISAB0_Sort
        Case 6:  fgYSWISAB0_Sort1 = 12: fgYSWISAB0_Sort2 = 12: fgYSWISAB0_Sort
        
        Case 7:  fgYSWISAB0_Sort1 = 7: fgYSWISAB0_Sort2 = 7: fgYSWISAB0_Sort
        Case 8:  fgYSWISAB0_Sort1 = 8: fgYSWISAB0_Sort2 = 8: fgYSWISAB0_Sort
        Case 9:  fgYSWISAB0_Sort1 = 9: fgYSWISAB0_Sort2 = 9: fgYSWISAB0_Sort
        Case 10: fgYSWISAB0_Sort1 = 10: fgYSWISAB0_Sort2 = 10: fgYSWISAB0_Sort
        Case fgYSWISAB0_arrIndex:  fgYSWISAB0_SortX fgYSWISAB0_arrIndex
    End Select
Else
    If fgYSWISAB0.Rows > 1 Then
        Call fgYSWISAB0_Color(fgYSWISAB0_RowClick, MouseMoveUsr.BackColor, fgYSWISAB0_ColorClick)
        
        fgYSWISAB0.Col = 11: xYSWISAB0.SWISABSWID = fgYSWISAB0.Text
        fgYSWISAB0.Col = 0: xMTK = Trim(fgYSWISAB0.Text)
        fgYSWISAB0.Col = 1: xBIC = Trim(fgYSWISAB0.Text)
        fgYSWISAB0.Col = 6: xBIC = xBIC & "   swift du " & Trim(fgYSWISAB0.Text)
        
        Call fgSwift_Display(xYSWISAB0.SWISABSWID, xMTK, xBIC)
        
        If X > 9350 Then
            Dim xSql As String
            xSql = "select  SWILNKAPPN from " & paramIBM_Library_SABSPE & ".YSWILNK0 " _
                & " where SWILNKAPPC = 'GOS' and SWILNKSTA = '' and SWILNKSWID = " & xYSWISAB0.SWISABSWID
            Set rsSabX = cnsab.Execute(xSql)
            If Not rsSabX.EOF Then Call frmYGOSDOS0_SQL_3(rsSabX("SWILNKAPPN"))
        End If
   End If
End If
Wait_SS 0
fgYSWISAB0.LeftCol = 0
End Sub


Private Sub fgSwift_Display(lSWISABSWID As Long, lMTK As String, lBIC As String)
Dim wColor As Long, wColorFixed As Long
'Dim X As String, xWhere As String, xOPE As String
Dim xSql As String
'Dim I As Long
'Dim blnOk As Boolean, blnDisplay As Boolean
'Dim wAmj As String

On Error GoTo Error_Handler
fraSwift.Visible = False
'fgswift_Reset
If Not blnSIDE_DB Then
    cnSIDE_DB.Open paramODBC_DSN_SIDE_DB
    blnSIDE_DB = True
End If
fgSwift.Rows = 1
'fgSwift.FormatString = fgSwift_FormatString
fgSwift.FormatString = "<" & lMTK & "    |<" & lBIC & "                                                       ||"
fgSwift.Row = 0
fgSwift.Col = 0: fgSwift.CellFontBold = True
fgSwift.Col = 1: fgSwift.CellFontBold = True
currentAction = "fgswift_Display"


xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABSWID = " & lSWISABSWID
Set rsSab = cnsab.Execute(xSql)
'___________________________________________________________________
 If Not rsSab.EOF Then
    Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)

    If oldYSWISAB0.SWISABWES = "E" Then
        X = "reçu de "
        wColor = RGB(190, 240, 255)
        wColorFixed = vbBlue
    Else
        X = "émis vers "
        wColor = RGB(220, 255, 220)
        wColorFixed = RGB(0, 64, 0)
    End If
    libSWIFT_SWISABSWID = "Dossier : " & Trim(oldYSWISAB0.SWISABOPEC) & " " & Format(oldYSWISAB0.SWISABOPEN, "### ###")
    fgSwift.Col = 0: fgSwift.Text = oldYSWISAB0.SWISABWMTK
    fgSwift.CellFontBold = True: fgSwift.CellBackColor = wColor
    fgSwift.ForeColorFixed = wColorFixed
    fgSwift.Col = 1: fgSwift.Text = X & oldYSWISAB0.SWISABWBIC & " le " & dateImp10(oldYSWISAB0.SWISABWAMJ) & " " & timeImp8(oldYSWISAB0.SWISABWHMS)
    fgSwift.CellFontBold = True: fgSwift.CellBackColor = wColor
    fgSwift.ForeColorFixed = wColorFixed
    fraSwift.BackColor = wColor

'If Not rsSab.EOF Then
'    libSWIFT_SWISABSWID = lSWISABSWID & " - " & rsSab("SWISABOPEC") & " " & rsSab("SWISABOPEN")

'    If rsSab("SWISABWES") = "E" Then
'        fgSwift.Col = 0: fgSwift.CellBackColor = RGB(32, 160, 255)
'        fgSwift.Col = 1: fgSwift.CellBackColor = RGB(32, 160, 255)
'        wColor = RGB(190, 240, 255)
'    Else
'        fgSwift.Col = 0: fgSwift.CellBackColor = RGB(32, 230, 190)
'        fgSwift.Col = 1: fgSwift.CellBackColor = RGB(32, 230, 190)
'        wColor = mColor_G0
'    End If
    xSql = "select * from rtextField " _
        & "where Aid = " & rsSab("SWISABWID1") _
        & " and text_s_umidl = " & rsSab("SWISABWIDL") _
        & " and text_s_umidh  =  " & rsSab("SWISABWIDH") _
        & " order by field_cnt"
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    If Not rsSIDE_DB.EOF Then
        Do While Not rsSIDE_DB.EOF
        
            fgSwift.Rows = fgSwift.Rows + 1
            fgSwift.Row = fgSwift.Rows - 1
        
            fgSwift_DisplayLine fgSwift.Row, wColor, wColorFixed
        
            rsSIDE_DB.MoveNext
        
        Loop
    Else
        xSql = "select * from rtext " _
            & "where Aid = " & rsSab("SWISABWID1") _
            & " and text_s_umidl = " & rsSab("SWISABWIDL") _
            & " and text_s_umidh  =  " & rsSab("SWISABWIDH")
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
        If Not rsSIDE_DB.EOF Then
            Call srvrText_GetBuffer_ODBC(rsSIDE_DB, xrText)
            fgSwift_DisplayLine_rText fgSwift.Row, wColor, wColorFixed
        End If
    End If
    
    fraSwift.Visible = True

    If chkSIDE_DB_Show Then frmSIDE_DB.fgSwift_Display lSWISABSWID, 0, 0, 0

End If

'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub fgSwift_DisplayLine(lIndex As Long, lCellBackColor As Long, lColorFixed As Long)
Dim K As Integer, iAsc13 As Integer, iLen As Integer
Dim blnAsc13 As Boolean
Dim xValue As String, V

On Error Resume Next
fgSwift.Col = 0: fgSwift.Text = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
fgSwift.CellBackColor = lCellBackColor
fgSwift.CellForeColor = lColorFixed
fgSwift.Col = 1
fgSwift.CellForeColor = lColorFixed

        Select Case rsSIDE_DB("field_code")
            Case "45", "46", "47", "77":
                V = rsSIDE_DB("value_memo")
                If IsNull(V) Then V = rsSIDE_DB("value")
            Case Else:
                    V = rsSIDE_DB("value")
        End Select
        If IsNull(V) Then
            xValue = ""
        Else
            xValue = V
        End If

 iLen = Len(xValue)
 K = 1
 Do
    iAsc13 = InStr(K, xValue, Asc13)
    If iAsc13 > 0 Then
        fgSwift.Text = Trim(Mid$(xValue, K, iAsc13 - K))
        fgSwift.CellForeColor = lColorFixed
        K = iAsc13 + 2
        fgSwift.Rows = fgSwift.Rows + 1
        fgSwift.Row = fgSwift.Rows - 1
    End If
 Loop Until iAsc13 = 0

fgSwift.Text = Trim(Mid$(xValue, K, iLen - K + 1))
fgSwift.CellForeColor = lColorFixed
fgSwift.Col = fgSwift.Cols - 1: fgSwift.Text = rsSIDE_DB("field_cnt")




End Sub



Public Sub fgSwift_DisplayLine_rText(lIndex As Long, lCellBackColor As Long, lColorFixed As Long)
Dim K As Integer, iAsc13 As Integer, iLen As Integer
Dim blnAsc13 As Boolean
Dim xValue As String, X As String, K2 As Integer

On Error Resume Next

xValue = xrText.text_data_block & Asc13
iLen = Len(xValue)
If Mid$(xValue, 1, 3) = Asc13 & Asc10 & ":" Then
    K = 3
Else
    K = 1
End If
Do
    iAsc13 = InStr(K, xValue, Asc13)
    If iAsc13 > 0 Then
        fgSwift.Rows = fgSwift.Rows + 1
        fgSwift.Row = fgSwift.Rows - 1
        X = Trim(Mid$(xValue, K, iAsc13 - K))
        fgSwift.Col = 1
        fgSwift.CellForeColor = lColorFixed
        If Mid$(X, 1, 1) <> ":" Then
            fgSwift.Text = Trim(Mid$(xValue, K, iAsc13 - K))
        Else
            K2 = InStr(2, X, ":")
            If K2 > 0 Then
                fgSwift.Text = Trim(Mid$(X, K2 + 1, Len(X) - K2))
                fgSwift.Col = 0: fgSwift.Text = Trim(Mid$(X, 2, K2 - 2))
                fgSwift.CellBackColor = lCellBackColor
                fgSwift.CellForeColor = lColorFixed
            Else
                fgSwift.Text = Trim(Mid$(xValue, K, iAsc13 - K))
            End If
        End If
        
        K = iAsc13 + 2
    End If
 Loop Until iAsc13 = 0

'fgSwift.Text = Trim(Mid$(xValue, K, iLen - K + 1))
'fgSwift.Col = fgSwift.Cols - 1: fgSwift.Text = rsSIDE_DB("field_cnt")


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
blnControl = True

End Sub



Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgSelect.Visible = False
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = 1 To fgSelect.FixedCols Step -1
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        fgSelect.Col = fgSelect_arrIndex
        lColor_Old = fgSelect.CellBackColor
        For I = 1 To fgSelect.FixedCols Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
    End If
End If
fgSelect.LeftCol = fgSelect.FixedCols
fgSelect.Visible = True
End Sub

Public Sub fgDetail_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgDetail.Visible = False: fraDetail.Visible = False
mRow = fgDetail.Row

If lRow > 0 And lRow < fgDetail.Rows Then
    fgDetail.Row = lRow
    For I = 2 To fgDetail.FixedCols Step -1
        fgDetail.Col = I: fgDetail.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgDetail.Row = mRow
    If fgDetail.Row > 0 Then
        lRow = fgDetail.Row
        fgDetail.Col = fgDetail_arrIndex
        lColor_Old = fgDetail.CellBackColor
        For I = 2 To fgDetail.FixedCols Step -1
          fgDetail.Col = I: fgDetail.CellBackColor = lColor
        Next I
    End If
End If
fgDetail.LeftCol = fgDetail.FixedCols
fgDetail.Visible = True: fraDetail.Visible = True
End Sub

Public Sub fgBIAMVT_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
SSTab2.Visible = False: fraDetail.Visible = False
mRow = fgBIAMVT.Row

If lRow > 0 And lRow < fgBIAMVT.Rows Then
    fgBIAMVT.Row = lRow
    For I = fgBIAMVT_arrIndex To fgBIAMVT.FixedCols Step -1
        fgBIAMVT.Col = I: fgBIAMVT.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgBIAMVT.Row = mRow
    If fgBIAMVT.Row > 0 Then
        lRow = fgBIAMVT.Row
        fgBIAMVT.Col = fgBIAMVT_arrIndex
        lColor_Old = fgBIAMVT.CellBackColor
        For I = fgBIAMVT_arrIndex To fgBIAMVT.FixedCols Step -1
          fgBIAMVT.Col = I: fgBIAMVT.CellBackColor = lColor
        Next I
    End If
End If
fgBIAMVT.LeftCol = fgBIAMVT.FixedCols
SSTab2.Visible = True: fraDetail.Visible = True
End Sub

Public Sub fgCPTPIE_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgCPTPIE.Visible = False: fraDetail.Visible = False
mRow = fgCPTPIE.Row

If lRow > 0 And lRow < fgCPTPIE.Rows Then
    fgCPTPIE.Row = lRow
    For I = fgCPTPIE_arrIndex To fgCPTPIE.FixedCols Step -1
        fgCPTPIE.Col = I: fgCPTPIE.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgCPTPIE.Row = mRow
    If fgCPTPIE.Row > 0 Then
        lRow = fgCPTPIE.Row
        fgCPTPIE.Col = fgCPTPIE_arrIndex
        lColor_Old = fgCPTPIE.CellBackColor
        For I = fgCPTPIE_arrIndex To fgCPTPIE.FixedCols Step -1
          fgCPTPIE.Col = I: fgCPTPIE.CellBackColor = lColor
        Next I
    End If
End If
fgCPTPIE.LeftCol = fgCPTPIE.FixedCols
fgCPTPIE.Visible = True: fraDetail.Visible = True
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





















Private Sub cboSelect_DOSSLDOPE_Change()
cmdSelect_Clear

End Sub

Private Sub cboSelect_DOSSLDPCI_Change()
cmdSelect_Clear

End Sub


Private Sub cboSelect_DOSSLDPCI_Click()
cmdSelect_Clear

End Sub


Private Sub cboSelect_DOSSLDSTA_Change()
cmdSelect_Clear

End Sub

Private Sub cboSelect_DOSSLDSTA_Click()
cmdSelect_Clear

End Sub

Private Sub cboSelect_SQL_Click()
'cmdSelect_Clear
cmdSelect_Reset
End Sub





Private Sub cboSelect_DOSSLDDEV_Change()
cmdSelect_Clear

End Sub

Private Sub cboSelect_DOSSLDDEV_Click()
cmdSelect_Clear

End Sub

Private Sub cboSelect_DOSSLDOPE_Click()
cmdDetail_Reset
End Sub


Private Sub chkSelect_DOSSLDSTA_Click()
cmdSelect_Clear

End Sub

Private Sub chkSelect_DOSSLDSVC_Click()
cmdSelect_Clear

End Sub


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub



Private Sub cmdPrint_Click()
Dim X As String, I As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass

Select Case SSTab1.Tab
    Case 0:
        Select Case cmdSelect_SQL_K
            Case "1":
                If SSTab2.Visible Then Me.PopupMenu mnuPrint2, vbPopupMenuLeftButton
            Case "3uti", "3RDO", "3", "5", "zSD", "GAR_Ech", "5 ECNFPT_Com", "5 AUT":
                Me.PopupMenu mnuPrint2, vbPopupMenuLeftButton
            
            Case "zOD": cmdSelect_SQL_YDOSXOD0_Export
        Case Else: Me.PopupMenu mnuPrint, vbPopupMenuLeftButton
        End Select
    End Select

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAB_Dossier_cmdSelect_Ok ........"): DoEvents

'If fgSelect.Visible Then
cmdSelect_Reset
fgSelect.Visible = False
'fraSelect_Options.Visible = False

Select Case cmdSelect_SQL_K
    Case "1": fraSelect_Options.Visible = True: cmdSelect_SQL_1
    Case "2", "2#": fraSelect_Options.Visible = True: cmdSelect_SQL_2
    Case "3":  cmdSelect_SQL_3
    Case "3uti":  cmdSelect_SQL_3uti
    Case "3RDO": cmdSelect_SQL_3RDO
    Case "5":  cmdSelect_SQL_5
    Case "5 AUT":  cmdSelect_SQL_5_AUT
    Case "5 réf ext =":  cmdSelect_SQL_5réfext ""
    Case "5 ECNFPT": cmdSelect_SQL_5_ECNFPT
    Case "5 ECNFPT_Com": cmdSelect_SQL_5_ECNFPT_Com
    Case "5 ECNFPT_CD7": cmdSelect_SQL_5_ECNFPT_CD7
    Case "6":  cmdSelect_SQL_6
    Case "X#":  cmdSelect_SQL_Surveillance
    Case "Xc":  cmdSelect_SQL_Xc
    Case "Xi":  cmdSelect_SQL_Xi
    Case "XE1an":  cmdSelect_SQL_XE1an
    Case "zOD":  cmdSelect_SQL_YDOSXOD0
    Case "zSD":  cmdSelect_SQL_YDOSNOK0
    Case "Scan": cmdSelect_SQL_Scan_Importation
    Case "Scan_Liste": cmdSelect_SQL_Scan_Liste
    Case "JPL":  cmdSelect_SQL_5_AUT 'cmdSelect_SQL_JPL 'cmdSelect_SQL_2_Exportation_JPL 'cmdSelect_SQL_JPL
    Case "GAR_Ech": Call cmdSelect_SQL_GAR_ECH
End Select
    
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_Dossier_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
'If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
End Sub


Private Sub cmdYDOSXOD0__Update_Click()
Dim V
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
newYDOSXOD0 = oldYDOSXOD0
newYDOSXOD0.DOSXODOPE = Trim(cboDOSXODOPE)
newYDOSXOD0.DOSXODNUM = Val(Trim(txtDOSXODNUM))

V = sqlYDOSXOD0_Update(newYDOSXOD0, oldYDOSXOD0)
If IsNull(V) Then
    fraYDOSXOD0.Visible = False

    Me.Enabled = True: Me.MousePointer = 0
    Exit Sub
End If
Error_Handler:
    Call MsgBox(V, vbCritical, "Mise à jour YDOSXOD0")
    Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdYDOSXOD0_Quit_Click()
fraYDOSXOD0.Visible = False

End Sub

Private Sub fgBIAMVT_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next

If y <= fgBIAMVT.RowHeightMin Then
Else
    If fgBIAMVT.Rows > 1 Then
        Call fgBIAMVT_Color(fgBIAMVT_RowClick, MouseMoveUsr.BackColor, fgBIAMVT_ColorClick)
        fgBIAMVT.Col = fgBIAMVT_arrIndex:  arrYDOSMVT0_Index = CLng(fgBIAMVT.Text)
        oldYDOSMVT0 = arrYDOSMVT0(arrYDOSMVT0_Index)
        xYDOSMVT0 = oldYDOSMVT0
        If cmdSelect_SQL_K = "2#" Then
            fraYDOSXOD0_Display
        Else

            fgCPTPIE_Display
        End If
        
   End If
End If


End Sub


Private Sub fgCPTPIE_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next

If y <= fgCPTPIE.RowHeightMin Then
Else
    If fgCPTPIE.Rows > 1 Then
        Call fgCPTPIE_Color(fgCPTPIE_RowClick, MouseMoveUsr.BackColor, fgCPTPIE_ColorClick)
        fgCPTPIE.Col = 0
        fraCompte_display Trim(fgCPTPIE.Text)
   End If
End If
End Sub


Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next


If y <= fgDetail.RowHeightMin Then
Else
    If fgDetail.Rows > 1 Then
        Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
        fgDetail.Col = fgDetail_arrIndex:  arrYDOSSLD0_Index = CLng(fgDetail.Text)
        oldYDOSSLD0 = arrYDOSSLD0(arrYDOSSLD0_Index)
        xYDOSSLD0 = oldYDOSSLD0
        fgBIAMVT_Display
        fgDossier.Visible = False
        fgYSWISAB0.Visible = False
        fgCOM.Visible = False
   End If
End If

End Sub


Private Sub fgLOG_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xMsgbox As String, xSql As String, Nb As Long
On Error Resume Next


If y <= fgLOG.RowHeightMin Then
    If cmdSelect_SQL_K = "zOD" Or cmdSelect_SQL_K = "zSD" Then

        Select Case fgLOG.Col
            Case 0: fgLog_Sort1 = 0: fgLog_Sort2 = 0: fglog_SortX 0
            Case 1:  fgLog_Sort1 = 1: fgLog_Sort2 = 2: fgLog_Sort
            Case 2: fgLog_Sort1 = 2: fgLog_Sort2 = 2: fgLog_Sort
            Case 3:  fgLog_Sort1 = 3: fgLog_Sort2 = 4: fgLog_Sort
            Case 4: fgLog_Sort1 = 4: fgLog_Sort2 = 4: fgLog_Sort
            Case 5:  fgLog_Sort1 = 5: fgLog_Sort2 = 5: fglog_SortX 5
            Case 6: fgLog_Sort1 = 6: fgLog_Sort2 = 6: fglog_SortX 6
            Case 7:  fgLog_Sort1 = 7: fgLog_Sort2 = 7: fgLog_Sort
            Case 8:  fgLog_Sort1 = 8: fgLog_Sort2 = 8: fgLog_Sort
            Case 9:  fgLog_Sort1 = 9: fgLog_Sort2 = 9: fglog_SortX 9
            Case 10:  fgLog_Sort1 = 10: fgLog_Sort2 = 10: fglog_SortX 10
            Case fgLog_arrIndex:  fglog_SortX fgLog_arrIndex
        End Select
    End If
Else
    If fgLOG.Rows > 1 Then
        Call fgLog_Color(fgLog_RowClick, MouseMoveUsr.BackColor, fgLog_ColorClick)
        'fgLOG.Col = fgLog_arrIndex:  arrYBIACPT0_Index = CLng(fgLOG.Text)
         If cmdSelect_SQL_K = "zOD" Then
                fraYDOSXOD0_Display
        Else
            If cmdSelect_SQL_K = "zSD" And arrHab(3) Then
                fgLOG.Col = 0: xYDOSMVT0.DOSMVTOPE = Trim(fgLOG.Text)
                fgLOG.Col = 1: xYDOSMVT0.DOSMVTNUM = CLng(fgLOG.Text)
                fgLOG.Col = 2: xYDOSMVT0.DOSMVTDEV = Trim(fgLOG.Text)
                fgLOG.Col = 3: xYDOSMVT0.DOSMVTPCI = Trim(fgLOG.Text)
                fgLOG.Col = 4: xYDOSMVT0.DOSMVTCLI = Trim(fgLOG.Text)
                xSql = xYDOSMVT0.DOSMVTOPE & "   " & xYDOSMVT0.DOSMVTNUM & "   " & xYDOSMVT0.DOSMVTPCI & "   " & xYDOSMVT0.DOSMVTCLI
                xMsgbox = MsgBox("Voulez-vous annuler cette régularisation non comptable ?" & vbCrLf & xSql, vbQuestion + vbYesNo, "frmSAB_Dossier : fgLog_Mousedown")
                If xMsgbox = vbYes Then
                    
                    xSql = "Update " & paramIBM_Library_SABSPE & ".YDOSNOK0 " _
                         & " set DOSNOKMSD = 0, DOSNOKGSD = 0 , DOSNOKUUSR = '" & usrName_UCase & "' , DOSNOKUAMJ = " & DSys & "  , DOSNOKUHMS = " & time_Hms _
                         & " where DOSNOKOPE = '" & xYDOSMVT0.DOSMVTOPE & "'" _
                         & " and   DOSNOKNUM = " & xYDOSMVT0.DOSMVTNUM _
                         & " and   DOSNOKDEV = '" & xYDOSMVT0.DOSMVTDEV & "'" _
                         & " and   DOSNOKPCI = '" & xYDOSMVT0.DOSMVTPCI & "'" _
                         & " and   DOSNOKCLI = '" & xYDOSMVT0.DOSMVTCLI & "'"
                    Set rsSab = cnsab.Execute(xSql, Nb)
                    
                    ' Tester si la mise à jour a été effectuée
                    '===================================================================================
                    
                    If Nb = 0 Then
                        Call MsgBox("Erreur mise à jour YDOSNOK0", vbCritical, "frmSAB_Dossier : fgLog_Mousedown")
                    End If
                    cmdSelect_SQL_YDOSNOK0
                End If
            End If
        End If
   End If


End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If chkSIDE_DB_Show = "1" Then frmSIDE_DB.Hide
If blnSIDE_DB Then
    cnSIDE_DB.Close
    Set cnSIDE_DB = Nothing
End If
End Sub

Private Sub fraTab0_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

    If fgSelect.Visible And fgYSWISAB0.Visible Then
        fgSelect.ZOrder 0
        fgYSWISAB0.ZOrder 1
    End If

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

If cmdSelect_SQL_K = "5 AUT" Then
    fgDetail.Visible = False
    fraDetail.Visible = False
    SSTab2.Visible = False
    fgSelect.BackColor = vbWhite
    fgSelect.Width = 16000
    fgSelect.Height = 8800
    Exit Sub
 
End If
If fraECNFPT.Visible Then
    fraECNFPT.Visible = False
    Exit Sub
End If

If fraSwift.Visible Then
    fraSwift.Visible = False
    Exit Sub
End If

If fraYDOSXOD0.Visible Then
    fraYDOSXOD0.Visible = False
    Exit Sub
End If

If fgLOG.Visible Then
    fgLOG.Visible = False
    Exit Sub
End If
If fraCompte.Visible Then
    fraCompte.Visible = False
    Exit Sub
End If
If fgCPTPIE.Visible Then
    fgCPTPIE.Visible = False
    Exit Sub
End If

If SSTab2.Visible Then
    SSTab2.Visible = False
    Exit Sub
End If
If fgDetail.Visible Then
    fgDetail.Visible = False
    fraDetail.Visible = False
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





Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String, xSql As String
On Error Resume Next


If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        If cmdSelect_SQL_K = "5 AUT" Then
                fgSelect.Col = 3: xYDOSSLD0.DOSSLDDEV = Trim(fgSelect.Text)
                fgSelect.Col = 0: xYDOSSLD0.DOSSLDOPE = Trim(fgSelect.Text)
                fgSelect.Col = 1: xYDOSSLD0.DOSSLDNUM = Val(Trim(fgSelect.Text))
                fgSelect.Col = 8: xYDOSSLD0.DOSSLDCLI = Trim(fgSelect.Text)
                cmdSelect_Clear
                fgDetail_Display
                If blnBIAMVT Then xYDOSSLD0 = oldYDOSSLD0: fgBIAMVT_Display: SSTab2.Tab = 0
                fgSelect.BackColor = RGB(220, 220, 220)
                fgSelect.Width = fgSelect_Width
                fgSelect.Height = fgSelect_Height
                fgSelect.Visible = True
        Else
        
            Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
            fgSelect.Col = fgSelect_arrIndex:  arrYBIACPT0_Index = CLng(fgSelect.Text)
            Select Case cmdSelect_SQL_K
                Case "1", "5 ECNFPT"
                    fgSelect.Col = 0: xYDOSSLD0.DOSSLDDEV = Trim(fgSelect.Text)
                    fgSelect.Col = 1: xYDOSSLD0.DOSSLDOPE = Trim(fgSelect.Text)
                    fgSelect.Col = 2: xYDOSSLD0.DOSSLDNUM = Val(Trim(fgSelect.Text))
                    fgSelect.Col = 3: xYDOSSLD0.DOSSLDCLI = Trim(fgSelect.Text)
                    fgDetail_Display
                    If blnBIAMVT Then xYDOSSLD0 = oldYDOSSLD0: fgBIAMVT_Display: SSTab2.Tab = 2
    
                Case "2", "2#":
                     fgSelect.Col = 0: xYDOSSLD1.DOSSLDDEV = Trim(fgSelect.Text)
                     fgSelect.Col = 1: xYDOSSLD1.DOSSLDPCI = Trim(fgSelect.Text)
                     fgSelect.Col = 2: xYDOSSLD1.DOSSLDCLI = Trim(fgSelect.Text)
                    
                     fgDetail_Display_2
                Case "3", "5", "5 réf ext ="
    
                    fgSelect.Col = 0: xYDOSSLD0.DOSSLDDEV = Trim(fgSelect.Text)
                    fgSelect.Col = 1: xYDOSSLD0.DOSSLDOPE = Trim(fgSelect.Text)
                    fgSelect.Col = 2: xYDOSSLD0.DOSSLDNUM = Val(Trim(fgSelect.Text))
                    fgDossier_Display
                    SSTab2.Tab = 1
                Case "3uti", "3RDO"
                    fgSelect.Col = 1: wOrigine = Trim(fgSelect.Text)
                    Dim K As Integer
                    K = InStr(wOrigine, " ")
                    xYDOSSLD0.DOSSLDOPE = Mid$(wOrigine, 1, K - 1)
                    xYDOSSLD0.DOSSLDNUM = Val(Mid$(wOrigine, K + 1, Len(wOrigine) - K))
                    fgYSWISAB0_Display
                    fgDossier.Visible = False
                    fgBIAMVT.Visible = False
                    fgCOM.Visible = False
                    SSTab2.Visible = True
                    SSTab2.Tab = 2
                    'SSTab2.ZOrder 0
                    fgSelect.ZOrder 1
                Case "Scan_Liste"
                    fgSelect.Col = 2
                    mDoc_Filename = fgSelect.Text
                    Call frmElpPrt.Windows_Display_File(mDoc_Filename)
            End Select
        End If
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


Public Sub fgDetail_Reset()
fgDetail.Clear
fgDetail_Sort1 = 0: fgDetail_Sort2 = 0
fgDetail_Sort1_Old = -1
fgDetail_RowDisplay = 0: fgDetail_RowClick = 0
fgDetail_arrIndex = fgDetail.Cols - 1
blnfgDetail_DisplayLine = False
fgDetail_SortAD = 6
fgDetail.LeftCol = fgDetail.FixedCols

End Sub




Public Sub fgBIAMVT_Reset()
fgBIAMVT.Clear
fgBIAMVT_Sort1 = 0: fgBIAMVT_Sort2 = 0
fgBIAMVT_Sort1_Old = -1
fgBIAMVT_RowDisplay = 0: fgBIAMVT_RowClick = 0
fgBIAMVT_arrIndex = fgBIAMVT.Cols - 1
blnfgBIAMVT_DisplayLine = False
fgBIAMVT_SortAD = 6
fgBIAMVT.LeftCol = fgBIAMVT.FixedCols

End Sub

Public Sub fgCPTPIE_Reset()
fgCPTPIE.Clear
fgCPTPIE_Sort1 = 0: fgCPTPIE_Sort2 = 0
fgCPTPIE_Sort1_Old = -1
fgCPTPIE_RowDisplay = 0: fgCPTPIE_RowClick = 0
fgCPTPIE_arrIndex = fgCPTPIE.Cols - 1
blnfgCPTPIE_DisplayLine = False
fgCPTPIE_SortAD = 6
fgCPTPIE.LeftCol = fgCPTPIE.FixedCols

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













Private Sub mnuPrint_Detail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdPrint_YDOSSLD0 True

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint_Recap_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdPrint_YDOSSLD0 False

Me.Enabled = True: Me.MousePointer = 0
End Sub


























Private Sub mnuDoc_Delete_Click()
On Error GoTo Error_Handler
If MsgBox("Confirmez-vous la suppression du document ?", vbYesNo, mDoc_Filename) = vbYes Then

Me.Enabled = False: Me.MousePointer = vbHourglass

    If InStr(mDoc_Filename, "_SCAN") > 0 Then
        msFileSystem.DeleteFile paramCDO_Dossier_Path & mDoc_Filename
    ElseIf frmSAB_Dossier.cboSelect_DOSSLDOPE.Text = "RDE" Or frmSAB_Dossier.cboSelect_DOSSLDOPE.Text = "RDI" Then
        msFileSystem.DeleteFile paramRDE_Dossier_Path & mDoc_Filename
    Else
        msFileSystem.DeleteFile paramCDO_Dossier_Path & mDoc_Filename
    End If
    Select Case SSTab2.Tab
        Case 3: fgCourrier_Display
        Case 4: fgScan_Display
    End Select
End If
Me.Enabled = True: Me.MousePointer = 0
Exit Sub
Error_Handler:
    Call MsgBox(Error, vbCritical, "mnuDoc_Delete_Click")
    Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuDoc_Display_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

'ON LIT TOUS LES DOCUMENTS SCANNES DANS LE MEME REPERTOIRE
'soit --> (\\docsrv2013\.biadoc$\RisquesOpérationnels\CREDOC\Production\_SCAN

'If cboSelect_DOSSLDOPE.Text = "RDE" Then
'    Call frmElpPrt.Windows_Display_File(paramRDE_Dossier_Path_DROPI & mDoc_Filename)
'Else
If InStr(mDoc_Filename, "_SCAN") > 0 Then
    Call frmElpPrt.Windows_Display_File(paramCDO_Dossier_Path_DROPI & mDoc_Filename)
ElseIf frmSAB_Dossier.cboSelect_DOSSLDOPE.Text = "RDE" Or frmSAB_Dossier.cboSelect_DOSSLDOPE.Text = "RDI" Then
    Call frmElpPrt.Windows_Display_File(paramRDE_Dossier_Path_DROPI & mDoc_Filename)
Else
    Call frmElpPrt.Windows_Display_File(paramCDO_Dossier_Path_DROPI & mDoc_Filename)
End If
'End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuDoc_Rename_Click()
Dim X As String, oldCDODOSCOP As String, oldCDODOSDOS As Long
Dim newCDODOSCOP As String, newCDODOSDOS As Long
Dim K As Integer, K2 As Integer
Dim mDOS_Path As String, newDoc_Filename As String

On Error GoTo Error_Handler

Me.Enabled = False: Me.MousePointer = vbHourglass

K = InStr(mDoc_Filename, "CD")
If K > 0 Then
    oldCDODOSCOP = Mid$(mDoc_Filename, K, 3)
    K2 = InStr(K, mDoc_Filename, "\")
    If K2 > 0 Then
        oldCDODOSDOS = Val(Mid$(mDoc_Filename, K + 4, K2 - K - 4))
    End If
End If
K = InStr(mDoc_Filename, "RD")
If K > 0 Then
    oldCDODOSCOP = Mid$(mDoc_Filename, K, 3)
    K2 = InStr(K, mDoc_Filename, "\")
    If K2 > 0 Then
        oldCDODOSDOS = Val(Mid$(mDoc_Filename, K + 4, K2 - K - 4))
    End If
End If

newCDODOSCOP = UCase(Trim(InputBox("Préciser le code opération de dossier : CDE, CDI, RDE ou RDI " _
    , "SAB_Dossier : renommer un document numérisé", oldCDODOSCOP)))
If newCDODOSCOP = "" Then Exit Sub

If newCDODOSCOP <> "CDE" And newCDODOSCOP <> "CDI" And newCDODOSCOP <> "RDE" And newCDODOSCOP <> "RDI" Then
    Call MsgBox("Le code opération du dossier doit être 'CDE','CDI','RDE' ou 'RDI'", vbCritical, "SAB_Dossier : renommer un document numérisé")
    Exit Sub
End If

X = Trim(InputBox("Préciser le nouveau numéro de dossier : " _
    , "SAB_Dossier : renommer un document numérisé", oldCDODOSDOS))
If X = "" Then Exit Sub
If Not IsNumeric(X) Then
    Call MsgBox("Le numéro du dossier doit être numérique", vbCritical, "SAB_Dossier : renommer un document numérisé")
    Exit Sub
End If
newCDODOSDOS = Val(X)
If newCDODOSCOP = "CDE" Or newCDODOSCOP = "CDI" Then
    X = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
          & " where CDODOSCOP = '" & newCDODOSCOP & "' and CDODOSDOS = " & newCDODOSDOS
ElseIf newCDODOSCOP = "RDE" Or newCDODOSCOP = "RDI" Then
    X = "select * from " & paramIBM_Library_SAB & ".ZENCCAR0 " _
          & " where ENCCARCOP = '" & newCDODOSCOP & "' and ENCCARDOS = " & newCDODOSDOS
End If
 Set rsSab = cnsab.Execute(X)

If rsSab.EOF Then
    Call MsgBox("Ce dossier est inconnu : " & newCDODOSCOP & " " & newCDODOSDOS, vbCritical, "SAB_Dossier : renommer un document numérisé")
    Exit Sub
End If

X = newCDODOSCOP & "_" & Format(newCDODOSDOS, "000000")
mDOS_Path = paramCDO_Dossier_Path & "_SCAN\" & X
If Not msFileSystem.FolderExists(mDOS_Path) Then MkDir mDOS_Path

newDoc_Filename = Replace(mDoc_Filename, oldCDODOSCOP & "_" & Format(oldCDODOSDOS, "000000"), X)
    msFileSystem.MoveFile paramCDO_Dossier_Path & mDoc_Filename, paramCDO_Dossier_Path & newDoc_Filename
    Select Case SSTab2.Tab
        Case 3: fgCourrier_Display
        Case 4: fgScan_Display
    End Select
'End If
Me.Enabled = True: Me.MousePointer = 0
Exit Sub
Error_Handler:
    Call MsgBox(Error, vbCritical, "mnuDoc_Delete_Click")
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint_2_Exportation_Click()
Dim xWhere As String, xSql As String
Dim rsSABY As New ADODB.Recordset

Me.Enabled = False: Me.MousePointer = vbHourglass

xWhere = " where COMPTEOBL like'" & xYDOSSLD1.DOSSLDPCI & "%'" _
     & " and CLIENACLI = '" & xYDOSSLD1.DOSSLDCLI & "'" _
     & " and COMPTEDEV = '" & xYDOSSLD1.DOSSLDDEV & "'"

     
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " & xWhere
Set rsSABY = cnsab.Execute(xSql)

Do While Not rsSABY.EOF
    V = rsYBIACPT0_GetBuffer(rsSABY, oldYBIACPT0)
    Call cmdSelect_SQL_2_Exportation_Init
    rsSABY.MoveNext
Loop


Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint_2_Liste_Click()
Dim I As Long

    On Error GoTo errHandler
    If fgSelect.Rows > 1 Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        For I = 0 To fgSelect.Cols - 1
            fgSelect.TextMatrix(fgSelect.Rows - 1, I) = "Le " & Format(Now, "dd mmm yyyy hh:nn:ss")
        Next I
        fgSelect.MergeCells = flexMergeFree
        fgSelect.MergeRow(fgSelect.Rows - 1) = True
        CmDialog2.PrinterDefault = True
        CmDialog2.CancelError = True
        CmDialog2.flags = cdlPDReturnDC + cdlPDNoPageNums + cdlPDDisablePrintToFile
        CmDialog2.ShowPrinter
        Printer.PaintPicture fgSelect.Picture, 0, 0
        Printer.EndDoc
        MsgBox "Fin de l'impression..."
        Call cmdSelect_Ok_Click
        Exit Sub
    End If
errHandler:
If Err = 32755 Then
    MsgBox "Impression annulée !"
Else
    MsgBox "Impression impossible !"
End If

End Sub


Private Sub mnuPrint2_Excel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String

Select Case cmdSelect_SQL_K
    Case "1":
        Select Case SSTab2.Tab
            Case 0:
                 X = "SAB_Dossier : " & Trim(txtSelect_DOSSLDNUM) & " - Comptabilité au " & dateImp10_S(DSys)
                Call MSflexGrid_Excel("", "SAB_Dossier", X, fgBIAMVT, fgBIAMVT.Cols - 1)
            Case 1:
                 X = "SAB_Dossier : " & Trim(txtSelect_DOSSLDNUM) & " - Dossier au " & dateImp10_S(DSys)
                Call MSflexGrid_Excel("", "SAB_Dossier", X, fgDossier, fgDossier.Cols - 1)
            Case 2:
                 X = "SAB_Dossier : " & Trim(txtSelect_DOSSLDNUM) & " - SWIFT au " & dateImp10_S(DSys)
                Call MSflexGrid_Excel("", "SAB_Dossier", X, fgYSWISAB0, fgYSWISAB0.Cols - 1)
            Case 3:
                 X = "SAB_Dossier : " & Trim(txtSelect_DOSSLDNUM) & " - Courrier au " & dateImp10_S(DSys)
                Call MSflexGrid_Excel("", "SAB_Dossier", X, fgCourrier, fgCourrier.Cols - 1)
            Case 4:
                 X = "SAB_Dossier : " & Trim(txtSelect_DOSSLDNUM) & " - Scan au " & dateImp10_S(DSys)
                Call MSflexGrid_Excel("", "SAB_Dossier", X, fgScan, fgScan.Cols - 1)
            Case 5:
                 X = "SAB_Dossier : " & Trim(txtSelect_DOSSLDNUM) & " - Commissions au " & dateImp10_S(DSys)
                Call MSflexGrid_Excel("", "SAB_Dossier", X, fgCOM, fgCOM.Cols - 1)
       End Select
        
    Case "3uti":
        X = "SAB_Dossier : utilisations en attente au " & dateImp10_S(DSys)
        Call MSflexGrid_Excel("", "SAB_Dossier", X, fgSelect, fgSelect.Cols - 1)
    Case "3RDO":
        X = "SAB_Dossier : RDO en attente au " & dateImp10_S(DSys)
        Call MSflexGrid_Excel("", "SAB_Dossier", X, fgSelect, fgSelect.Cols - 1)
    Case "3":
        X = "SAB_Dossier : Evénements à valider au " & dateImp10_S(DSys)
        Call MSflexGrid_Excel("", "SAB_Dossier", X, fgSelect, fgSelect.Cols - 1)
    Case "zSD":
        X = cboSelect_SQL.Text & "  au " & dateImp10_S(DSys)
        Call MSflexGrid_Excel("", "SAB_Dossier", X, fgLOG, fgLOG.Cols - 1)
    Case "zSD":
        X = cboSelect_SQL.Text & "  au " & dateImp10_S(DSys)
        Call MSflexGrid_Excel("", "SAB_Dossier", X, fgSelect, fgSelect.Cols - 1)
    Case "GAR_Ech":
        X = "Echéancier des garanties" & "  au " & dateImp10_S(DSys)
        Call MSflexGrid_Excel("", "SAB_Dossier_GAR_Ech", X, fgX, 9)
    Case "5 ECNFPT_Com":
        X = "Liste des commissions ECNFPT" & "  au " & dateImp10_S(DSys)
        Call MSflexGrid_Excel("", "SAB_Dossier_ECNFPT_Com", X, fgCOM, 17)
    Case "5 AUT":
        X = "Rapprochement SAB_Dossier / Consultation globale" & "  au " & dateImp10_S(DSys)
        Call MSflexGrid_Excel("", "SAB_Dossier_5_AUT", X, fgSelect, 8)
End Select

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint2_Mail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String
Select Case cmdSelect_SQL_K
    Case "1":
        Select Case SSTab2.Tab
            Case 0:
                 X = "SAB_Dossier : " & Trim(txtSelect_DOSSLDNUM) & " - Comptabilité au " & dateImp10_S(DSys)
                Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Dossier", X, X, fgBIAMVT, fgBIAMVT.Cols - 1)
            Case 1:
                 X = "SAB_Dossier : " & Trim(txtSelect_DOSSLDNUM) & " - Dossier au " & dateImp10_S(DSys)
                Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Dossier", X, X, fgDossier, fgDossier.Cols - 1)
            Case 2:
                 X = "SAB_Dossier : " & Trim(txtSelect_DOSSLDNUM) & " - SWIFT au " & dateImp10_S(DSys)
                Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Dossier", X, X, fgYSWISAB0, fgYSWISAB0.Cols - 1)
            Case 3:
                 X = "SAB_Dossier : " & Trim(txtSelect_DOSSLDNUM) & " - Courrier au " & dateImp10_S(DSys)
                Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Dossier", X, X, fgCourrier, fgCourrier.Cols - 1)
            Case 4:
                 X = "SAB_Dossier : " & Trim(txtSelect_DOSSLDNUM) & " - Scan au " & dateImp10_S(DSys)
                Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Dossier", X, X, fgScan, fgScan.Cols - 1)
            Case 5:
                 X = "SAB_Dossier : " & Trim(txtSelect_DOSSLDNUM) & " - Commissions au " & dateImp10_S(DSys)
                Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Dossier", X, X, fgCOM, fgCOM.Cols - 1)
       End Select
    Case "3uti":
        X = "SAB_Dossier : utilisations en attente au " & dateImp10_S(DSys)
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Dossier", X, X, fgSelect, fgSelect.Cols - 1)
    Case "3RDO":
        X = "SAB_Dossier : RDO en attente au " & dateImp10_S(DSys)
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Dossier", X, X, fgSelect, fgSelect.Cols - 1)
    Case "3":
        X = "SAB_Dossier : Evénements à valider au " & dateImp10_S(DSys)
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Dossier", X, X, fgSelect, fgSelect.Cols - 1)
    Case "zSD":
        X = cboSelect_SQL.Text & "  au " & dateImp10_S(DSys)
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Dossier", X, X, fgLOG, fgLOG.Cols - 1)
    Case "zSD":
        X = cboSelect_SQL.Text & "  au " & dateImp10_S(DSys)
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Dossier", X, X, fgSelect, fgSelect.Cols - 1)
    Case "GAR_Ech":
        X = cboSelect_SQL.Text & "  au " & dateImp10_S(DSys)
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Dossier : GAR_Ech", X, X, fgX, 9)
    Case "5 ECNFPT_Com":
        X = cboSelect_SQL.Text & "  au " & dateImp10_S(DSys)
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Dossier : ECNFPT_Com", X, X, fgCOM, 17)
    Case "5 AUT":
         X = "Rapprochement SAB_Dossier / Consultation globale" & "  au " & dateImp10_S(DSys)
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SAB_Dossier : 5_AUT", X, X, fgSelect, 8)

End Select



Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub SSTab2_Click(PreviousTab As Integer)

If SSTab2.Tab = 0 Then
    If Me.WindowState = 2 Then
        cmdSAB_Dossier_DB.Left = SSTab2.Left + 2800
    Else
        cmdSAB_Dossier_DB.Left = SSTab2.Left + 2055
    End If
End If
If SSTab2.Tab = 1 And Not fgDossier.Visible Then fgDossier_Display
If SSTab2.Tab = 2 And Not fgYSWISAB0.Visible Then fgYSWISAB0_Display
If SSTab2.Tab = 3 Then
    If xYDOSSLD0.DOSSLDOPE = "CDE" Or xYDOSSLD0.DOSSLDOPE = "CDI" Or xYDOSSLD0.DOSSLDOPE = "RDE" Or xYDOSSLD0.DOSSLDOPE = "RDI" Then
        cmdSAB_Dossier_CDO.Visible = True
    Else
        cmdSAB_Dossier_CDO.Visible = False
    End If
    
    If Me.WindowState = 2 Then
        cmdSAB_Dossier_CDO.Left = SSTab2.Width - 6800
    Else
        cmdSAB_Dossier_CDO.Left = SSTab2.Width - 5800
    End If
    
    fgCourrier_Display
End If
If SSTab2.Tab = 4 And Not fgScan.Visible Then fgScan_Display
If SSTab2.Tab = 5 And Not fgCOM.Visible Then
    If xYDOSSLD0.DOSSLDOPE = "CDE" Or xYDOSSLD0.DOSSLDOPE = "CDI" Then fgCOM_Display
End If
End Sub

Private Sub txtDOSXODNUM_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtECNFPT_DREG_Change()

Call DTPicker_Control(txtECNFPT_DREG, wAmjMin)
mECNFPT_DREG = Val(wAmjMin)
If mECNFPT_DREG < mECNFPT_DDEB Then
    libECNFPT_NBJ_X = "????"
    libECNFPT_MON_X = ""
    libECNFPT_TOT_X = ""
Else
    'If mECNFPT_DREG = mECNFPT_DFIN Then
    '    mECNFPT_NBJ_X = mECNFPT_NBJ
    'Else
        If mECNFPT_DREG > mECNFPT_DFIN Then
             mECNFPT_NBJ_X = mECNFPT_NBJ + DateDiff("d", dateImp_Amj(mECNFPT_DFIN), dateImp_Amj(mECNFPT_DREG))
        Else
             mECNFPT_NBJ_X = DateDiff("d", dateImp_Amj(mECNFPT_DDEB), dateImp_Amj(mECNFPT_DREG))
             If mECNFPT_NBJ_X > mECNFPT_NBJ Then mECNFPT_NBJ_X = mECNFPT_NBJ
        End If
        
    'End If
    libECNFPT_NBJ_X = mECNFPT_NBJ_X
    curX = Fix(mECNFPT_MTA * mECNFPT_TX1 * mECNFPT_NBJ_X / mECNFPT_Ratio + 0.00500001) / 100
    libECNFPT_MON_X = Trim(Format$(curX, "### ### ##0.00"))
    curX = mECNFPT_TOT - mECNFPT_MON + curX
    If curX < mECNFPT_MIN Then curX = mECNFPT_MIN
    
    libECNFPT_TOT_X = Trim(Format$(curX, "### ### ##0.00"))

End If

End Sub

Private Sub txtECNFPT_Dreg_GotFocus()
libECNFPT_MON_X = ""
libECNFPT_NBJ_X = ""
End Sub


Private Sub txtSelect_6_CLIEANCLI_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub

Private Sub txtSelect_6_PCI_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub

Private Sub txtSelect_DOSSLDNUM_Change()
cmdSelect_Clear
End Sub

Private Sub txtSelect_DOSSLDNUM_GotFocus()
Call txt_GotFocus(txtSelect_DOSSLDNUM)

End Sub


Private Sub txtSelect_DOSSLDNUM_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub



Public Sub fraCompte_display(lCOMPTECOM As String)
Dim xSql As String

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '" & lCOMPTECOM & "'"
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then
    V = rsYBIACPT0_GetBuffer(rsSab, xYBIACPT0)
    If IsNull(V) Then
        txtD_COMPTECOM = xYBIACPT0.COMPTECOM
        txtD_COMPTEINT = xYBIACPT0.COMPTEINT
        txtD_COMPTEOBL = xYBIACPT0.COMPTEOBL
        txtD_COMPTEFON = xYBIACPT0.COMPTEFON
        txtD_PLANCOPRO = xYBIACPT0.PLANCOPRO
        If xYBIACPT0.COMPTEOUV = 0 Then
            txtD_COMPTEOUV = ""
        Else
            txtD_COMPTEOUV = dateIBM10(xYBIACPT0.COMPTEOUV, True)
        End If
        If xYBIACPT0.COMPTECLO = 0 Then
            txtD_COMPTECLO = ""
        Else
            txtD_COMPTECLO = dateIBM10(xYBIACPT0.COMPTECLO, True)
        End If

        txtD_CLIENACLI = xYBIACPT0.CLIENACLI
        txtD_CLIENASIG = xYBIACPT0.CLIENASIG
        txtD_CLIENARES = xYBIACPT0.CLIENARES
        txtD_CLIENARA1 = xYBIACPT0.CLIENARA1
        txtD_CLIENANAT = xYBIACPT0.CLIENANAT
        txtD_CLIENARSD = xYBIACPT0.CLIENARSD


        fraCompte.Visible = True
    End If
End If
End Sub

Public Sub cmdSelect_SQL_6_Exportation_Xlsx_Recap(lK As String, lRow2 As Long, lCpt As String, lInt As String, lRef As String, lTxt As String)
Dim K As Integer, wColor As Long
On Error GoTo Error_Handler

Set wsExcel = wbExcel.Sheets(1)

mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = lK
If lRow2 <> 0 Then wsExcel.Cells(mXls1_Row, 2) = lRow2
wsExcel.Cells(mXls1_Row, 3) = lCpt
wsExcel.Cells(mXls1_Row, 4) = lInt
wsExcel.Cells(mXls1_Row, 5) = lRef
wsExcel.Cells(mXls1_Row, 6) = lTxt

If lK = "" Then
    wColor = mColor_G0
Else
    Select Case Mid$(lK, 1, 1)
        Case "#": wColor = mColor_W1
        Case "?": wColor = RGB(255, 245, 255)
        Case "!": wColor = mColor_Y1
        Case "+": wColor = mColor_W0: wsExcel.Cells(mXls1_Row, 1) = Mid$(lK, 2, Len(lK) - 1)
        Case "-": wColor = mColor_G0: wsExcel.Cells(mXls1_Row, 1) = Mid$(lK, 2, Len(lK) - 1)
        Case "_": wColor = mColor_GB: wsExcel.Cells(mXls1_Row, 1) = Mid$(lK, 2, Len(lK) - 1)
        Case Else: wColor = mColor_Y0
    End Select
End If
For K = 1 To 6
    wsExcel.Cells(mXls1_Row, K).Interior.Color = wColor
Next K
Set wsExcel = wbExcel.Sheets(2)

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub cmdSelect_SQL_6_Exportation_Control(lRow As Long, lMTD As Currency)
Dim K As Long, X As String, wDOSSLDMSD As Currency, wPC As Double, wCDODOSPPO As Double, wCDODOSPDE As Currency, xCur As Currency
Dim wDOSSLDSTA As String, xDOSSLDSTA_SVC As String, xCDODOSProv As String, xCDODOSDEV As String, xCDODOSVAL As String, wCDODOSVAL As Long

wDOSSLDMSD = 0: wPC = 0: wCDODOSPPO = 0: wCDODOSPDE = 0: xCDODOSDEV = "???"
If mXls2_Row_D = 0 Then mXls2_Row_D = lRow
arrZCDODOS0_K = 0: xCDODOSProv = "": xCDODOSVAL = "": wCDODOSVAL = 0
For K = 1 To arrZCDODOS0_Nb
   If oldYDOSMVT0.DOSMVTOPE = arrZCDODOS0(K).CDODOSCOP And oldYDOSMVT0.DOSMVTNUM = arrZCDODOS0(K).CDODOSDOS Then
        arrZCDODOS0_K = K
        arrZCDODOS0(K).CDODOSETB = -1
        wCDODOSPPO = arrZCDODOS0(arrZCDODOS0_K).CDODOSPPO: wCDODOSPDE = arrZCDODOS0(arrZCDODOS0_K).CDODOSPDE
        xCDODOSDEV = arrZCDODOS0(arrZCDODOS0_K).CDODOSDEV
        xCDODOSProv = "   (SAB " & wCDODOSPPO & " % => " & Format$(wCDODOSPDE, "### ### ##0.00") & ")"
        wCDODOSVAL = arrZCDODOS0(arrZCDODOS0_K).CDODOSVAL + 19000000
        xCDODOSVAL = dateImp_Amj(wCDODOSVAL)
        Exit For
    End If
Next K
        
'_______________________________________________________________________________________
If arrZCDODOS0_K = 0 And oldYDOSMVT0.DOSMVTNUM > 0 Then

   X = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
         & " where CDODOSCOP = '" & oldYDOSMVT0.DOSMVTOPE & "' and CDODOSDOS = " & oldYDOSMVT0.DOSMVTNUM
    Set rsSab = cnsab.Execute(X)

If Not rsSab.EOF Then
        wCDODOSPPO = rsSab("CDODOSPPO"): wCDODOSPDE = rsSab("CDODOSPDE")
        xCDODOSDEV = rsSab("CDODOSDEV")
        xCDODOSProv = "   (Prov " & wCDODOSPPO & " % => " & Format$(wCDODOSPDE, "### ### ##0.00") & ")"
End If
'_______________________________________________________________________________________


End If

arrYDOSSLD0_K = 0: xDOSSLDSTA_SVC = "": wDOSSLDSTA = ""
For K = 1 To arrYDOSSLD0_Nb
   If oldYDOSMVT0.DOSMVTOPE = arrYDOSSLD0(K).DOSSLDOPE And oldYDOSMVT0.DOSMVTNUM = arrYDOSSLD0(K).DOSSLDNUM Then
   'And oldYDOSSLD0.DOSSLDDEV = arrYDOSSLD0(K).DOSSLDDEV Then
        arrYDOSSLD0_K = K
        wDOSSLDSTA = arrYDOSSLD0(K).DOSSLDSTA
        xDOSSLDSTA_SVC = "   (" & arrYDOSSLD0(K).DOSSLDSTA & "-" & arrYDOSSLD0(K).DOSSLDSVC & ")"
        wsExcel.Cells(mXls2_Row_D, 1) = Trim(wsExcel.Cells(mXls2_Row_D, 1)) & xDOSSLDSTA_SVC
        wsExcel.Cells(mXls2_Row_D, 1).Font.Color = RGB(0, 64, 64)

        arrYDOSSLD0(K).DOSSLDSVC = "**"
        Exit For
    End If
Next K


X = ""

If arrYDOSSLD0_K > 0 Then
    wDOSSLDMSD = -arrYDOSSLD0(arrYDOSSLD0_K).DOSSLDMSD
    X = "solde du crédit : " & Format$(wDOSSLDMSD, "### ### ##0.00") & " " & xCDODOSDEV & xCDODOSProv
    If oldYDOSSLD0.DOSSLDDEV <> xCDODOSDEV Then wsExcel.Cells(lRow, 7).Interior.Color = mColor_W1

End If
wsExcel.Cells(lRow, 7) = X
wsExcel.Cells(lRow, 8) = xCDODOSVAL
X = ""

If wDOSSLDMSD = 0 Then
    If lMTD = 0 Then
        wsExcel.Cells(lRow, 6).Interior.Color = mColor_G0
    Else
        If wDOSSLDSTA = "80" Or wDOSSLDSTA = "90" Then
            wsExcel.Cells(lRow, 6) = "????"
            wsExcel.Cells(lRow, 6).Interior.Color = mColor_W0
            mDOSSLDMSD_Nb = mDOSSLDMSD_Nb + 1
            X = "? dossier utilisé en totalité, mais une provision de " & Format$(lMTD, "### ### ##0.00") & oldYDOSSLD0.DOSSLDDEV & " est comptabilisée."
            Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("?Prov 0", lRow, oldYBIACPT0.COMPTECOM, oldYBIACPT0.COMPTEINT, oldYDOSMVT0.DOSMVTOPE & " " & oldYDOSMVT0.DOSMVTNUM, X)
        Else
            wsExcel.Cells(lRow, 6).Interior.Color = mColor_G0
        End If
    End If
Else
    xCur = wDOSSLDMSD * wCDODOSPPO / 100
    
    If Abs(lMTD - xCur) < 10 Then
        wsExcel.Cells(lRow, 6) = Format$(lMTD * 100 / wDOSSLDMSD, "##") & " %" 'wPC = lMTD * 100 / wDOSSLDMSD
    
        wsExcel.Cells(lRow, 6).Interior.Color = mColor_G0
    Else
        If wCDODOSPPO > 1 Then
            wsExcel.Cells(lRow, 6) = Format$(xCur, "### ### ##0.00")
        Else
            wsExcel.Cells(lRow, 6) = "????"
        End If
        
        wsExcel.Cells(lRow, 6).Interior.Color = mColor_W0
        mProv_Nb = mProv_Nb + 1
        X = "? prov calculée  : " & Format$(xCur, "### ### ##0.00") & xCDODOSDEV & " mais " & Format$(lMTD, "### ### ##0.00") & " " & oldYDOSSLD0.DOSSLDDEV & " provisionnés "
        Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("?Prov #", lRow, oldYBIACPT0.COMPTECOM, oldYBIACPT0.COMPTEINT, oldYDOSMVT0.DOSMVTOPE & " " & oldYDOSMVT0.DOSMVTNUM, X)
    End If
End If
'--------------------------------------------------------------------------------------
If arrZCDODOS0_K = 0 Then
    If oldYDOSMVT0.DOSMVTNUM > 0 And mMTDJ_Dos <> 0 Then
        mCDODOSDOS_Nb = mCDODOSDOS_Nb + 1
        wsExcel.Cells(lRow, 6).Interior.Color = mColor_W0
        Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("?Dos", lRow, oldYBIACPT0.COMPTECOM, oldYBIACPT0.COMPTEINT, oldYDOSMVT0.DOSMVTOPE & " " & oldYDOSMVT0.DOSMVTNUM, oldYBIACPT0.COMPTECOM & "dossier annulé ou sans % de provisions")
    End If
Else
    If Trim(oldYBIACPT0.COMPTECOM) <> Trim(arrZCDODOS0(arrZCDODOS0_K).CDODOSPCC) Then
        If wDOSSLDMSD > 0 Then
            mCDODOSPCC_Nb = mCDODOSPCC_Nb + 1
            wsExcel.Cells(lRow, 6).Interior.Color = mColor_W0
            Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("?Cpt", lRow, oldYBIACPT0.COMPTECOM, oldYBIACPT0.COMPTEINT, oldYDOSMVT0.DOSMVTOPE & " " & oldYDOSMVT0.DOSMVTNUM, "! compte de provisions (compta : " & oldYBIACPT0.COMPTECOM & ") # (" & arrZCDODOS0(arrZCDODOS0_K).CDODOSPCC & " gestion CDODOSPCC)")
        End If
    End If
    'If mMTDJ_Dos <> arrZCDODOS0(arrZCDODOS0_K).CDODOSPDE Then
    '    wsExcel.Cells(lRow, 7).Font.Color = mColor_W1
    '    mCDODOSPDE_Nb = mCDODOSPDE_Nb + 1
    '    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("#Prov", lRow, oldYBIACPT0.COMPTECOM, oldYBIACPT0.COMPTEINT, oldYDOSMVT0.DOSMVTOPE & " " & oldYDOSMVT0.DOSMVTNUM, "# solde de provisions (compta : " & Format$(mMTDJ_Dos, "### ### ##0.00") & ") # (" & Format$(arrZCDODOS0(arrZCDODOS0_K).CDODOSPDE, "### ### ##0.00") & " gestion CDODOSPDE) ")
    'End If
End If
wsExcel.Cells(mXls2_Row_D, 1).Interior.Color = wsExcel.Cells(lRow, 6).Interior.Color
wsExcel.Cells(mXls2_Row_D, 2).Interior.Color = wsExcel.Cells(lRow, 6).Interior.Color
wsExcel.Cells(lRow, 5).Interior.Color = wsExcel.Cells(lRow, 6).Interior.Color
wsExcel.Cells(lRow, 7).Interior.Color = wsExcel.Cells(lRow, 6).Interior.Color
wsExcel.Cells(lRow, 6).Font.Color = RGB(0, 64, 64)
wsExcel.Cells(lRow, 7).Font.Color = RGB(0, 64, 64)

wsExcel.Cells(lRow, 8).Interior.Color = wsExcel.Cells(lRow, 6).Interior.Color
If wCDODOSVAL < YBIATAB0_DATE_CPT_JS1 And lMTD <> 0 Then
    If wCDODOSVAL <= YBIATAB0_DATE_CPT_MP2 Then
        mCDODOSVAL_Nb = mCDODOSVAL_Nb + 1
        wsExcel.Cells(lRow, 8).Interior.Color = mColor_W1
    Else
        wsExcel.Cells(lRow, 8).Interior.Color = mColor_W0
    End If
End If

wsExcel.Cells(lRow, 8).Font.Color = RGB(0, 64, 64)

End Sub

Public Sub cmdSelect_SQL_6_Exportation_Xlsx_Dos()
Dim K As Integer

mMTD9_Dos = mMTD0_Dos + mMTD9_Dos
mMTDJ_Dos = mMTD0_Dos + mMTDJ_Dos
'________________________________________________________________________________
If mMTD9_Dos = 0 Then
    If Not blnDos_Ok Then
        For K = 1 To arrZCDODOS0_Nb
           If oldYDOSMVT0.DOSMVTOPE = arrZCDODOS0(K).CDODOSCOP And oldYDOSMVT0.DOSMVTNUM = arrZCDODOS0(K).CDODOSDOS Then
                blnDos_Ok = True
                Exit For
            End If
        Next K
    End If
End If
'________________________________________________________________________________


If blnDos_Ok Or mMTD9_Dos <> 0 Then
    blnCli_Ok = True
    mXls2_Row = mXls2_Row + 1
    If mXls2_Row_D = 0 Then mXls2_Row_D = mXls2_Row
    wsExcel.Cells(mXls2_Row_D, 1) = oldYDOSMVT0.DOSMVTOPE & " " & oldYDOSMVT0.DOSMVTNUM ': wsExcel.Cells(mXls2_Row, 4).Font.Bold = True

    If Not blnDos_Ok Then wsExcel.Cells(mXls2_Row, 2) = mMTD0_Dos
    wsExcel.Cells(mXls2_Row, 5) = mMTD9_Dos
    If blnProvisions_Control Then
        If oldYDOSMVT0.DOSMVTOPE = "***" And mMTD9_Dos <> 0 Then
             For K = 1 To 6:    wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_W0: Next K
             mAnn_Nb = mAnn_Nb + 1
             Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("!Ann", mXls2_Row, oldYBIACPT0.COMPTECOM, oldYBIACPT0.COMPTEINT, "", "Dossiers annulés en gestion présentant un solde en comptabilité")
        Else
            Call cmdSelect_SQL_6_Exportation_Control(mXls2_Row, mMTD9_Dos)
        End If
    End If
End If

oldYDOSMVT0.DOSMVTOPE = newYDOSMVT0.DOSMVTOPE: oldYDOSMVT0.DOSMVTNUM = newYDOSMVT0.DOSMVTNUM: mMTD0_Dos = 0: mMTD9_Dos = 0: mMTDJ_Dos = 0: blnDos_Ok = False
mXls2_Row_D = 0
End Sub

Public Sub cmdSelect_SQL_6_Exportation_Xlsx_Cli()
Dim K As Integer

'_______________________________________________________________________________
mMTD9_Cli = mMTD0_Cli + mMTD9_Cli
If blnCli_Ok Or mMTD9_Cli <> 0 Then
    If mMTD9_Cli = oldYBIACPT0.SOLDECEN Then
       ' For K = 1 To 6:    wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_G1: Next K
    Else
        mXls2_Row = mXls2_Row + 1
        wsExcel.Cells(mXls2_Row, 2) = mMTD0_Cli ': wsExcel.Cells(mXls2_Row, 2).Font.Bold = True
        wsExcel.Cells(mXls2_Row, 5) = mMTD9_Cli ': wsExcel.Cells(mXls2_Row, 5).Font.Bold = True
        wsExcel.Cells(mXls2_Row, 1) = oldYDOSSLD0.DOSSLDCLI: wsExcel.Cells(mXls2_Row, 1).Font.Bold = True
        wsExcel.Cells(mXls2_Row, 6) = oldYDOSSLD0.DOSSLDDEV ': wsExcel.Cells(mXls2_Row, 6).Font.Bold = True
        For K = 1 To 6:    wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_W1: wsExcel.Cells(mXls2_Row_Cli, K).Interior.Color = mColor_W1: Next K
        Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("#solde", mXls2_Row, oldYBIACPT0.COMPTECOM, oldYBIACPT0.COMPTEINT, " ", "Ecart entre le solde comptable et total par dossier")
    End If
End If


End Sub

Public Sub cmdSelect_SQL_6_Exportation_Xlsx_Recap_End()
Dim K As Long, K1 As Long
For K = 1 To arrZCDODOS0_Nb
   If arrZCDODOS0(K).CDODOSETB <> -1 Then
        If arrZCDODOS0(K).CDODOSEVE = "01" And arrZCDODOS0(K).CDODOSETA = "01" Then
        Else
            blnDos_Ok = False
            For K1 = 1 To arrYDOSSLD0_Nb
               If arrZCDODOS0(K).CDODOSCOP = arrYDOSSLD0(K1).DOSSLDOPE And arrZCDODOS0(K).CDODOSDOS = arrYDOSSLD0(K1).DOSSLDNUM Then
                    If arrYDOSSLD0(K1).DOSSLDMSD = 0 Then blnDos_Ok = True
                    Exit For
                End If
            Next K1
            If Not blnDos_Ok Then
                mCDODOSXXX_Nb = mCDODOSXXX_Nb + 1
                Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("?Prov", 0, arrZCDODOS0(K).CDODOSPCC, "", arrZCDODOS0(K).CDODOSCOP & " " & arrZCDODOS0(K).CDODOSDOS, " provision non comptabilisée.")
            End If
        End If
   End If
Next K
'__________________________________________________________________________________
        
Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("_****", 0, "", "", "", "")

If mCDODOSPCC_Nb = 0 Then
    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("-#Cpt", 0, "", "", "", "aucun dossier non concordant compta /gestion (compte de provisions).")
Else
    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("+#Cpt", 0, "", "", "", mCDODOSPCC_Nb & " dossiers non concordants compta /gestion (compte de provisions).")
End If

If mCDODOSXXX_Nb = 0 Then
    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("-?Prov", 0, "", "", "", "aucun dossier non comptabilisé (compte de provisions).")
Else
    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("+?Prov", 0, "", "", "", mCDODOSXXX_Nb & " dossiers non comptabilisés (compte de provisions).")
End If

'__________________________________________________________________________________
'(lColor As Long, lRow2 As Long, lCpt As String, lInt As String, lRef As String, lTxt As String)
If mCDODOSDOS_Nb = 0 Then
    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("-?Dos", 0, "", "", "", "aucun dossier sans % de provisions dans SAB/ZCDODOS0.")
Else
    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("+?Dos", 0, "", "", "", mCDODOSDOS_Nb & " dossiers annulés ou sans % de provisions.")
End If
'If mCDODOSPDE_Nb = 0 Then
'    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("-#Prov", 0, "", "", "", "aucun dossier en écart compta /gestion (montant de provisions).")
'Else
'    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("+#Prov", 0, "", "", "", mCDODOSPDE_Nb & " dossiers en écart compta /gestion (montant de provisions).")
'End If
If mDOSSLDMSD_Nb = 0 Then
    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("-?Prov 0", 0, "", "", "", "aucun dossier utilisé en totalité ayant une provision comptabilisée.")
Else
    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("+?Prov 0", 0, "", "", "", mDOSSLDMSD_Nb & " dossiers utilisés en totalité, mais une provision comptabilisée.")
End If
If mProv_Nb = 0 Then
    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("-?Prov #", 0, "", "", "", "aucun dossier utilisé en totalité ayant une provision comptabilisée.")
Else
    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("+?Prov #", 0, "", "", "", mProv_Nb & " dossiers dont la provison calculée est différente de la provision comptabilisée.")
End If
If mAnn_Nb = 0 Then
    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("-?ann #", 0, "", "", "", "aucune alerte pour dossier annulé en gestion présentant un solde en comptabilité.")
Else
    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("+?Ann #", 0, "", "", "", mAnn_Nb & " alertes pour dossiers annulés en gestion présentant un solde en comptabilité.")
End If
If mCDODOSVAL_Nb = 0 Then
    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("-?Dos", 0, "", "", "", "aucun dossier échu au " & dateImp10(YBIATAB0_DATE_CPT_MP2) & " présentant un solde comptable de provisions.")
Else
    Call cmdSelect_SQL_6_Exportation_Xlsx_Recap("+?Dos", 0, "", "", "", mCDODOSVAL_Nb & " dossiers  échus au " & dateImp10(YBIATAB0_DATE_CPT_MP2) & " présentant un solde comptable de provisions.")
End If
End Sub

Public Sub cmdSelect_SQL_Xc_Dossier_S(lDOSCD7KCN As String)
Dim curX As Currency, K As Integer
On Error GoTo Error_Handler
If oldYDOSCD70.DOSCD7NUM = 117568 Then
MsgBox ("OK")
End If
If oldYDOSCD70.DOSCD7NUM <> 0 Then
    mXls2_Row = mXls2_Row + 1
    
If oldYDOSCD70.DOSCD7NUM = 117296 Then
    Debug.Print ">>> TRACE 117296 <<<"
    Debug.Print "Avant reset: COM_G2=" & sMTD_COM_G2 & " G2PDIF=" & sMTD_COM_G2PDIF
    Debug.Print "ANN=" & blnCDODOSANN & " Opt=" & optSelect_DOSCD7DAN_Out
    Debug.Print "COM_ANN_C=" & sMTD_COM_ANN_C & " COM_ANN_N=" & sMTD_COM_ANN_N
End If


    
    If blnCDODOSANN And optSelect_DOSCD7DAN_Out Then
        sMTD_UTI_G = 0:  sMTD_COM_G2Prata = 0: sMTD_COM_G2PDIF = 0
        Select Case lDOSCD7KCN
            Case "C": sMTD_COM_G2PDIF = sMTD_COM_ANN_C
            Case "N": sMTD_COM_G2PDIF = sMTD_COM_ANN_N
        End Select
        If sMTD_COM_G2PDIF = 0 Then sMTD_COM_G2 = 0
    End If
    
    wsExcel.Cells(mXls2_Row, 1) = oldYDOSCD70.DOSCD7OPE
    wsExcel.Cells(mXls2_Row, 2) = oldYDOSCD70.DOSCD7NUM
    
    wsExcel.Cells(mXls2_Row, 3) = oldYDOSCD70.DOSCD7KCN
    wsExcel.Cells(mXls2_Row, 4) = oldYDOSCD70.DOSCD7CLI
    wsExcel.Cells(mXls2_Row, 5) = dateImp10(oldYDOSCD70.DOSCD7DDEB)
    If oldYDOSCD70.DOSCD7NUM = "117239" Then
        MsgBox ("OK")
    End If
    wsExcel.Cells(mXls2_Row, 6) = dateImp10(oldYDOSCD70.DOSCD7DFIN)
    wsExcel.Cells(mXls2_Row, 7) = sMTD_Solde_C
    wsExcel.Cells(mXls2_Row, 8) = oldYDOSCD70.DOSCD7DEV
    If sMTD_UTI_G <> 0 Then wsExcel.Cells(mXls2_Row, 9) = sMTD_UTI_G
    If oldYDOSCD70.DOSCD7DAMJ > 0 Then wsExcel.Cells(mXls2_Row, 10) = dateImp10(oldYDOSCD70.DOSCD7DAMJ)
    
    
    Debug.Print "DOSSIER:", oldYDOSCD70.DOSCD7NUM, _
            "ANN?", blnCDODOSANN, _
            "Option:", optSelect_DOSCD7DAN_Out, _
            "COM_G2:", sMTD_COM_G2, _
            "COM_G2PDIF:", sMTD_COM_G2PDIF, _
            "COM_G3:", sMTD_COM_G3

    
    curX = sMTD_COM_G2 + sMTD_COM_G3
    

    If curX <> 0 Then wsExcel.Cells(mXls2_Row, 11) = curX
    If sMTD_COM_G3 <> 0 Then wsExcel.Cells(mXls2_Row, 12) = sMTD_COM_G3
    If sMTD_COM_G2 <> 0 Then wsExcel.Cells(mXls2_Row, 13) = sMTD_COM_G2
    If sMTD_COM_G2PDIF <> 0 Then wsExcel.Cells(mXls2_Row, 14) = sMTD_COM_G2PDIF
    If sMTD_COM_G2Prata <> 0 Then
        wsExcel.Cells(mXls2_Row, 15) = sMTD_COM_G2Prata
        curX = sMTD_COM_G2 - sMTD_COM_G2Prata - sMTD_COM_G2PDIF
        If curX <> 0 Then wsExcel.Cells(mXls2_Row, 16) = curX
    End If
    If sMTD_TC2 <> 0 Then wsExcel.Cells(mXls2_Row, 17) = sMTD_TC2
    
    If blnCDODOSANN Then
            For K = 1 To mXls2_Col
                wsExcel.Cells(mXls2_Row, K).Interior.Color = RGB(240, 240, 240)
            Next K
            wsExcel.Cells(mXls2_Row, 10) = dateImp10(arrZCDODOS0(arrZCDODOS0_K).CDODOSDAN)
            wsExcel.Cells(mXls2_Row, 10).Font.Color = vbRed
            wsExcel.Cells(mXls2_Row, 10).Interior.Color = mColor_Y1
    Else
        If oldYDOSCD70.DOSCD7DFIN <= wAmjMin Then
            For K = 1 To mXls2_Col
                wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_Y0
            Next K
            wsExcel.Cells(mXls2_Row, 6).Font.Color = vbMagenta
        End If
   End If
   
    wsExcel.Cells(mXls2_Row, 14).Interior.Color = RGB(210, 240, 240)
    If blnCDODOSANN And sMTD_COM_G2PDIF <> 0 Then
        wsExcel.Cells(mXls2_Row, 14).Interior.Color = mColor_Y1
        wsExcel.Cells(mXls2_Row, 19) = sMTD_COM_G2PDIF
        wsExcel.Cells(mXls2_Row, 19).Interior.Color = mColor_Y1
    End If
    If sMTD_COM_G3 <> sMTD_COM_C Then wsExcel.Cells(mXls2_Row, 12).Interior.Color = vbMagenta: wsExcel.Cells(mXls2_Row, 18) = sMTD_COM_C
    
    tMTD_Solde_C = tMTD_Solde_C + sMTD_Solde_C
    tMTD_COM_C = tMTD_COM_C + sMTD_COM_C
    tMTD_UTI_G = tMTD_UTI_G + sMTD_UTI_G
    tMTD_COM_G2 = tMTD_COM_G2 + sMTD_COM_G2
    tMTD_COM_G2Prata = tMTD_COM_G2Prata + sMTD_COM_G2Prata
    tMTD_COM_G3 = tMTD_COM_G3 + sMTD_COM_G3
    tMTD_COM_G2PDIF = tMTD_COM_G2PDIF + sMTD_COM_G2PDIF
    
'______________________________________________________________________________________________________________________
    If dosYDOSCD70.DOSCD7STA = "07" Then
        wsExcel.Cells(mXls2_Row, 1) = oldYDOSCD70.DOSCD7OPE & " r"
        wsExcel.Cells(mXls2_Row, 1).Interior.Color = RGB(255, 255, 96)
        wsExcel.Cells(mXls2_Row, 2).Interior.Color = RGB(255, 255, 96)
    Else
        For K = 1 To arrCDOMODEVE_07_Nb
            If oldYDOSCD70.DOSCD7NUM = arrCDOMODEVE_07(K) Then
                wsExcel.Cells(mXls2_Row, 1) = oldYDOSCD70.DOSCD7OPE & " *"
                wsExcel.Cells(mXls2_Row, 1).Interior.Color = RGB(255, 255, 96)
                wsExcel.Cells(mXls2_Row, 2).Interior.Color = RGB(255, 255, 96)
                Exit For
            Else
                If oldYDOSCD70.DOSCD7NUM < arrCDOMODEVE_07(K) Then Exit For

            End If
        Next K
    End If

'______________________________________________________________________________________________________________________
    If xYDOSCD70.DOSCD7DEV <> oldYDOSCD70.DOSCD7DEV Then cmdSelect_SQL_Xc_Dossier_T
'$jpl 2013-07-09 ___________________________________________________________________________
        For K = 1 To arrECNFPT_DOS_Nb
            If oldYDOSCD70.DOSCD7NUM = arrECNFPT_DOS(K) Then
                If blnECNFPT_CD7 Then
                    wsExcel.Cells(mXls2_Row, 11).Interior.Color = RGB(0, 255, 0)
                    wsExcel.Cells(mXls2_Row, 17).Interior.Color = RGB(0, 255, 0)
                Else
                    wsExcel.Cells(mXls2_Row, 11).Interior.Color = RGB(0, 255, 0)
                End If
                Exit For
            End If
        Next K
'____________________________________________________________________________________________
'______________________________________________________________________________________________________________________
End If




oldYDOSCD70 = xYDOSCD70
sMTD_Solde_C = 0: sMTD_COM_C = 0: sMTD_UTI_G = 0: sMTD_COM_G2 = 0: sMTD_COM_G3 = 0: sMTD_COM_G2Prata = 0: sMTD_COM_G2PDIF = 0
sMTD_TC2 = 0
blnCDODOSANN = False
sMTD_COM_ANN_C = 0: sMTD_COM_ANN_N = 0
For arrZCDODOS0_K = 1 To arrZCDODOS0_Nb
    If oldYDOSCD70.DOSCD7OPE = arrZCDODOS0(arrZCDODOS0_K).CDODOSCOP _
   And oldYDOSCD70.DOSCD7NUM = arrZCDODOS0(arrZCDODOS0_K).CDODOSDOS Then
   
   If arrZCDODOS0(arrZCDODOS0_K).CDODOSDOS = 117296 Then
    Debug.Print ">>> TRACE 117296 <<<"
    Debug.Print "Avant reset: COM_G2=" & sMTD_COM_G2 & " G2PDIF=" & sMTD_COM_G2PDIF
    Debug.Print "ANN=" & blnCDODOSANN & " Opt=" & optSelect_DOSCD7DAN_Out
    Debug.Print "COM_ANN_C=" & sMTD_COM_ANN_C & " COM_ANN_N=" & sMTD_COM_ANN_N
End If
   
        blnCDODOSANN = True
        sMTD_COM_ANN_C = arrZCDODOS0(arrZCDODOS0_K).CDODOSMOC: sMTD_COM_ANN_N = arrZCDODOS0(arrZCDODOS0_K).CDODOSMOT
        Exit For
    End If
Next arrZCDODOS0_K

Exit Sub

Error_Handler:
    If Not blnAuto Then MsgBox Error, vbCritical, Me.Name
End Sub
Public Sub cmdSelect_SQL_Xc_Dossier_T()
Dim K As Integer
On Error GoTo Error_Handler

mXls2_Row = mXls2_Row + 1
    
wsExcel.Cells(mXls2_Row, 7).FormulaLocal = "=SOMME(G" & mDev_R1 & ":G" & mXls2_Row - 1 & ")"
wsExcel.Cells(mXls2_Row, 9).FormulaLocal = "=SOMME(I" & mDev_R1 & ":I" & mXls2_Row - 1 & ")"
wsExcel.Cells(mXls2_Row, 11).FormulaLocal = "=SOMME(K" & mDev_R1 & ":K" & mXls2_Row - 1 & ")"
wsExcel.Cells(mXls2_Row, 12).FormulaLocal = "=SOMME(L" & mDev_R1 & ":L" & mXls2_Row - 1 & ")"
wsExcel.Cells(mXls2_Row, 13).FormulaLocal = "=SOMME(M" & mDev_R1 & ":M" & mXls2_Row - 1 & ")"
wsExcel.Cells(mXls2_Row, 14).FormulaLocal = "=SOMME(N" & mDev_R1 & ":N" & mXls2_Row - 1 & ")"
wsExcel.Cells(mXls2_Row, 15).FormulaLocal = "=SOMME(O" & mDev_R1 & ":O" & mXls2_Row - 1 & ")"
wsExcel.Cells(mXls2_Row, 16).FormulaLocal = "=SOMME(P" & mDev_R1 & ":P" & mXls2_Row - 1 & ")"
wsExcel.Cells(mXls2_Row, 19).FormulaLocal = "=SOMME(S" & mDev_R1 & ":S" & mXls2_Row - 1 & ")"

For K = 1 To mXls2_Col

    wsExcel.Cells(mXls2_Row, K).Interior.Color = RGB(190, 220, 220)
Next K

tMTD_Solde_C = 0: tMTD_COM_C = 0: tMTD_UTI_G = 0: tMTD_COM_G2 = 0: tMTD_COM_G3 = 0: tMTD_COM_G2Prata = 0: tMTD_COM_G2PDIF = 0

mDev_R1 = mXls2_Row + 1

For K = 1 To arrDev_Nb
    If oldYDOSCD70.DOSCD7DEV = arrDev(K) Then
        arrDev_RowT(K) = mXls2_Row
        Exit For
    End If
Next K

Exit Sub

Error_Handler:
    If Not blnAuto Then MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub cmdSelect_SQL_Xi_Dossier_T(lCLIENARA1 As String, lDOSSLDPCI As String)
Dim K As Integer
Dim curSolde As Currency, xSql As String

If oldYDOSSLD0.DOSSLDDEV <> "" Then
    mXls2_Row = mXls2_Row + 1
    wsExcel.Cells(mXls2_Row, 4) = oldYDOSSLD0.DOSSLDCLI: wsExcel.Cells(mXls2_Row, 4).Font.Bold = True
    wsExcel.Cells(mXls2_Row, 8) = oldYDOSSLD0.DOSSLDDEV: wsExcel.Cells(mXls2_Row, 8).Font.Bold = True
    wsExcel.Cells(mXls2_Row, 7).FormulaLocal = "=SOMME(G" & mDev_R1 & ":G" & mXls2_Row - 1 & ")"
    wsExcel.Cells(mXls2_Row, 7).Font.Bold = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    wsExcel.Cells(mXls2_Row, 9).FormulaLocal = "=SOMME(I" & mDev_R1 & ":I" & mXls2_Row - 1 & ")"
    wsExcel.Cells(mXls2_Row, 9).Font.Bold = True
    '///////////////////////////////////////////////////////////////////////////////////////////
    
    wsExcel.Cells(mXls2_Row, 11).FormulaLocal = "=SOMME(K" & mDev_R1 & ":K" & mXls2_Row - 1 & ")"
    wsExcel.Cells(mXls2_Row, 11).Font.Bold = True
    wsExcel.Cells(mXls2_Row, 12).FormulaLocal = "=SOMME(L" & mDev_R1 & ":L" & mXls2_Row - 1 & ")"
    wsExcel.Cells(mXls2_Row, 12).Font.Bold = True
    wsExcel.Cells(mXls2_Row, 13).FormulaLocal = "=SOMME(M" & mDev_R1 & ":M" & mXls2_Row - 1 & ")"
    wsExcel.Cells(mXls2_Row, 13).Font.Bold = True
    wsExcel.Cells(mXls2_Row, 14).FormulaLocal = "=SOMME(N" & mDev_R1 & ":N" & mXls2_Row - 1 & ")"
    wsExcel.Cells(mXls2_Row, 14).Font.Bold = True
    wsExcel.Cells(mXls2_Row, 15) = lCLIENARA1
    For K = 1 To mXls2_Col
        wsExcel.Cells(mXls2_Row, K).Interior.Color = RGB(190, 220, 220)
    Next K
    
    mDev_R1 = mXls2_Row + 1
    
    For K = 1 To arrDev_Nb
        If oldYDOSSLD0.DOSSLDDEV = arrDev(K) Then
            arrDev_RowT(K) = mXls2_Row
            Exit For
        End If
    Next K
    
'===================================================================================
    curSolde = 0
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
         & " where COMPTEOBL like '" & lDOSSLDPCI & "%'" _
         & " and COMPTEDEV = '" & oldYDOSSLD0.DOSSLDDEV & "'" _
         & " and CLIENACLI = '" & oldYDOSSLD0.DOSSLDCLI & "'"
         
    Set rsSabX = cnsab.Execute(xSql)
    
    Do While Not rsSabX.EOF
        curSolde = curSolde - CCur(rsSabX("SOLDECEN")) / 1000
        rsSabX.MoveNext
    Loop
    wsExcel.Cells(mXls2_Row, 16) = curSolde
    If curSolde <> CCur(wsExcel.Cells(mXls2_Row, 7)) Then wsExcel.Cells(mXls2_Row, 16).Interior.Color = mColor_W1
'===================================================================================

End If
oldYDOSSLD0 = xYDOSSLD0
End Sub


Public Sub cmdSelect_SQL_Xc_ZSOLDE0(lSheet As Integer, lPCI As String)
On Error GoTo Error_Handler
Dim xSql As String
Dim X As String, K As Integer
Dim curX As Currency, xDev As String, mDev As String
'__________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "> Exportation ........ " & lSheet & "-" & lPCI): DoEvents

Set wsExcel = wbExcel.Sheets(lSheet)

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Solde des comptes PCI " & lPCI & " , arrêté au " & dateImp10(wAmjMin) _
                                & vbCr & "&B&U&10" & vbCr
wsExcel.PageSetup.CenterHorizontally = True

wsExcel.PageSetup.PrintTitleRows = "$A1:$D1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignLeft
    .WrapText = True
    .Font.Size = 8
    .Font.Name = "Arial Unicode MS"
    .RowHeight = 17
End With



wsExcel.Columns(1).ColumnWidth = 5: wsExcel.Cells(1, 1) = "Devise": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 17: wsExcel.Cells(1, 2) = "Compte"
wsExcel.Columns(3).ColumnWidth = 32: wsExcel.Cells(1, 3) = "Intitulé"
wsExcel.Columns(4).ColumnWidth = 17: wsExcel.Cells(1, 4) = "Solde  ": wsExcel.Columns(4).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignRight

mXls2_Row = 1
mDev_R1 = 2
mXls2_Col = 4

For K = 1 To mXls2_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB ' RGB(255, 128, 50)
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next K

xSql = "select * from " & paramIBM_Library_SAB & ".ZCOMPTE0 , " & paramIBM_Library_SAB & ".ZSOLDE0" _
     & " where substring(COMPTEOBL , 1 , 5) in " & lPCI _
     & " and  COMPTECOM = SOLDECOM" _
     & " order by COMPTEDEV"
Set rsSab = cnsab.Execute(xSql)
mDev = ""
Do While Not rsSab.EOF
    xDev = rsSab("COMPTEDEV")
    If mDev <> xDev Then Call cmdSelect_SQL_Xc_ZSOLDE0_T(xDev, mDev)

    Select Case mSOLDE_K
        Case 0:    curX = rsSab("SOLDECEN")
        Case 1:    curX = rsSab("SOLDEC01")
        Case 2:    curX = rsSab("SOLDEC02")
        Case 3:    curX = rsSab("SOLDEC03")
        Case 4:    curX = rsSab("SOLDEC04")
        Case 5:    curX = rsSab("SOLDEC05")
        Case 6:    curX = rsSab("SOLDEC06")
        Case 7:    curX = rsSab("SOLDEC07")
        Case 8:    curX = rsSab("SOLDEC08")
        Case 9:    curX = rsSab("SOLDEC09")
        Case 10:   curX = rsSab("SOLDEC10")
        Case 11:   curX = rsSab("SOLDEC11")
        Case 12:   curX = rsSab("SOLDEC12")
        Case Else: curX = 0
    End Select
    
    mXls2_Row = mXls2_Row + 1

    wsExcel.Cells(mXls2_Row, 1) = xDev
    wsExcel.Cells(mXls2_Row, 2) = rsSab("COMPTECOM")
    wsExcel.Cells(mXls2_Row, 3) = rsSab("COMPTEINT")
    wsExcel.Cells(mXls2_Row, 4) = -curX

    rsSab.MoveNext
Loop
'==============================================================================================

Call cmdSelect_SQL_Xc_ZSOLDE0_T("", mDev)

Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & lSheet & "-" & lPCI): DoEvents


'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & lPCI): DoEvents


End Sub

Public Sub cmdSelect_SQL_Xc_ZSOLDE0_T(xDev As String, mDev As String)
Dim K As Integer

If mDev <> "" Then
    mXls2_Row = mXls2_Row + 1
    wsExcel.Cells(mXls2_Row, 1) = mDev: wsExcel.Cells(mXls2_Row, 1).Font.Bold = True
    wsExcel.Cells(mXls2_Row, 4).FormulaLocal = "=SOMME(D" & mDev_R1 & ":D" & mXls2_Row - 1 & ")"
    wsExcel.Cells(mXls2_Row, 4).Font.Bold = True
    For K = 1 To mXls2_Col
        wsExcel.Cells(mXls2_Row, K).Interior.Color = RGB(190, 220, 220)
    Next K
    
    For K = 1 To arrDev_Nb
        If mDev = arrDev(K) Then
            arrDev_RowT(K) = mXls2_Row
            Exit For
        End If
    Next K

End If
mDev_R1 = mXls2_Row + 1
mDev = xDev


End Sub

Public Sub cmdSelect_SQL_Xi_Init(lSheet As Integer)
Dim K As Integer, K2 As Integer
On Error Resume Next
Set wsExcel = wbExcel.Sheets(lSheet)


With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignRight
    .WrapText = False ' True
    .Font.Size = 9
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 65
'wsExcel.PageSetup.PrintTitleRows = "$A1:$L1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
Select Case lSheet
    Case 1
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14" & mHeader_xls & ", arrêté au " & dateImp10(wAmjMin) _
                                        & vbCr & "&B&U&10(crédits documentaires export confirmés et paiements différés)"
    Case 2
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14" & mHeader_xls & " > 92 jours, arrêté au " & dateImp10(wAmjMin) _
                                        & vbCr & "&B&U&10(crédits documentaires export paiements différés >= 93 jours)"
    Case 3
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Liste des engagements INTRAGROUPE, arrêté au " & dateImp10(wAmjMin) _
                                        & vbCr & "&B&U&10(crédits documentaires export paiements différés < 93 jours)"
End Select

wsExcel.PageSetup.CenterHorizontally = True



wsExcel.Columns(1).ColumnWidth = 9:  wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 30: wsExcel.Cells(mXls1_Row_C, 2) = "Intitulé":  wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignLeft

wsExcel.Cells(mXls1_Row_C, 1) = "91120"
wsExcel.Cells(mXls1_Row_C, 1).Interior.Color = mColor_Y1
wsExcel.Cells(mXls1_Row_C, 2).Interior.Color = mColor_GB: wsExcel.Cells(mXls1_Row_C, 2).Font.Color = mColor_Z0

wsExcel.Cells(mXls1_Row_N - 1, 1) = "T 91120"
wsExcel.Cells(mXls1_Row_N, 1) = "9113*":
wsExcel.Cells(mXls1_Row_N, 1).Interior.Color = mColor_Y1
wsExcel.Cells(mXls1_Row_N, 2).Interior.Color = mColor_GB:    wsExcel.Cells(mXls1_Row_N, 2).Font.Color = mColor_Z0

wsExcel.Cells(mXls1_Row_SP - 1, 1) = "T 9113*"
wsExcel.Cells(mXls1_Row_SP, 1) = "91121":
wsExcel.Cells(mXls1_Row_SP, 1).Interior.Color = mColor_Y1
wsExcel.Cells(mXls1_Row_SP, 2).Interior.Color = mColor_GB:    wsExcel.Cells(mXls1_Row_N, 2).Font.Color = mColor_Z0

wsExcel.Cells(mXls1_Row_T - 1, 1) = "T 91121"
wsExcel.Cells(mXls1_Row_T + 1, 1) = "Total":
wsExcel.Cells(mXls1_Row_T, 1).Interior.Color = mColor_GB
wsExcel.Cells(mXls1_Row_T, 2).Interior.Color = mColor_GB:    wsExcel.Cells(mXls1_Row_T, 2).Font.Color = mColor_Z0
wsExcel.Cells(mXls1_Row_T + 2, 1) = "Cours":
wsExcel.Cells(mXls1_Row_T + 3, 1) = " Total Dev."
If lSheet <> 2 And lSheet <> 3 Then
    wsExcel.Cells(mXls1_Row_T + 4, 1) = " € >= 93 J "
    wsExcel.Cells(mXls1_Row_T + 5, 1) = " € < 93 J "
End If
wsExcel.Cells(mXls1_Row_T + 6, 1) = " Total €"

For K = 1 To arrDev_Nb
    K2 = K + 2
    wsExcel.Cells(mXls1_Row_C, K2) = arrDev(K)
    wsExcel.Cells(mXls1_Row_C, K2).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_C, K2).Font.Color = mColor_Z0
    wsExcel.Cells(mXls1_Row_N, K2) = arrDev(K)
    wsExcel.Cells(mXls1_Row_N, K2).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_N, K2).Font.Color = mColor_Z0
    wsExcel.Cells(mXls1_Row_SP, K2) = arrDev(K)
    wsExcel.Cells(mXls1_Row_SP, K2).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_SP, K2).Font.Color = mColor_Z0
    wsExcel.Cells(mXls1_Row_T, K2) = arrDev(K)
    wsExcel.Cells(mXls1_Row_T, K2).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row_T, K2).Font.Color = mColor_Z0
    wsExcel.Columns(K2).ColumnWidth = 13: wsExcel.Columns(K2).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
    
    If arrDev(K) = "EUR" Or arrDev(K) = "USD" Then wsExcel.Columns(K2).ColumnWidth = 16
    wsExcel.Cells(mXls1_Row_T + 2, K2).NumberFormat = "### ##0.00000"
    wsExcel.Cells(mXls1_Row_T + 2, K2) = arrDev_Cours(K)
Next K

For K = 1 To mXls1_Col
    wsExcel.Cells(mXls1_Row_N - 1, K).Interior.Color = mColor_G0
    wsExcel.Cells(mXls1_Row_SP - 1, K).Interior.Color = mColor_G0
    wsExcel.Cells(mXls1_Row_T - 1, K).Interior.Color = mColor_G0
    wsExcel.Cells(mXls1_Row_T + 1, K).Interior.Color = mColor_G0
    wsExcel.Cells(mXls1_Row_T + 3, K).Interior.Color = mColor_Y0
    If lSheet <> 2 And lSheet <> 3 Then
        wsExcel.Cells(mXls1_Row_T + 4, K).Interior.Color = mColor_Y0
        wsExcel.Cells(mXls1_Row_T + 5, K).Interior.Color = mColor_Y0
    End If
    wsExcel.Cells(mXls1_Row_T + 6, K).Interior.Color = mColor_Y0
Next K


End Sub
Public Sub fraYDOSXOD0_Display()
Dim X As String, K As Integer, wAmj As String
Dim V
On Error GoTo Error_Handler

If cmdSelect_SQL_K = "zOD" Then
    fgLOG.Col = 0: X = Trim(fgLOG.Text): Call dateJma10_Amj(X, wAmj)
    xYDOSXOD0.DOSXODDTR = Val(wAmj)
    fgLOG.Col = 10:  X = Trim(fgLOG.Text)

Else
    fgBIAMVT.Col = 5: X = Trim(fgBIAMVT.Text): Call dateJma10_Amj(X, wAmj)
    xYDOSXOD0.DOSXODDTR = Val(wAmj)
    fgBIAMVT.Col = 6:  X = Trim(fgBIAMVT.Text)
End If

K = InStr(X, "-")
If K = 0 Then
    V = "Erreur de décodage : " & X
    GoTo Error_Handler
End If

xYDOSXOD0.DOSXODPIE = Val(Mid$(X, 1, K - 1))
xYDOSXOD0.DOSXODECR = Val(Mid$(X, K + 1, Len(X) - K))

X = "select * from " & paramIBM_Library_SABSPE & ".YDOSXOD0 " _
   & " where DOSXODDTR = " & xYDOSXOD0.DOSXODDTR _
   & " and   DOSXODPIE = " & xYDOSXOD0.DOSXODPIE _
   & " and   DOSXODecr = " & xYDOSXOD0.DOSXODECR
   
Set rsSab = cnsab.Execute(X)

If rsSab.EOF Then Exit Sub

Call rsYDOSXOD0_GetBuffer(rsSab, oldYDOSXOD0)
libYDOSXOD0 = "màj par " & Trim(oldYDOSXOD0.DOSXODUUSR) & " le " & dateImp10(oldYDOSXOD0.DOSXODUAMJ) & " à " & timeImp8(oldYDOSXOD0.DOSXODUHMS)
lblDOSXODOPE = "Code opération : " & oldYDOSXOD0.DOSXODOPE
lblDOSXODNUM = "Numéro opération : " & oldYDOSXOD0.DOSXODNUM
If oldYDOSXOD0.DOSXODNUM = 0 Then
    txtDOSXODNUM = ""
Else
    txtDOSXODNUM = oldYDOSXOD0.DOSXODNUM
End If

cboDOSXODOPE.Clear
cboDOSXODOPE.AddItem "CDE"
cboDOSXODOPE.AddItem "CDI"
cboDOSXODOPE.AddItem "ENG"
cboDOSXODOPE.AddItem "GAR"
cboDOSXODOPE.AddItem "RDE"
cboDOSXODOPE.AddItem "RDI"
Select Case oldYDOSXOD0.DOSXODOPE
    Case "CDI": cboDOSXODOPE.ListIndex = 1
    Case Else: cboDOSXODOPE.ListIndex = 0
End Select
fraYDOSXOD0.Visible = True
Exit Sub

Error_Handler:

Call MsgBox(V, vbCritical, "Mise à jour YDOSXOD0")

End Sub

Public Sub cmdSelect_SQL_YDOSXOD0()
Dim V, X As String
Dim xSql As String, xWhere As String

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_YDOSXOD0"

Call DTPicker_Control(txtSelect_Options_Log_AmjMin, wAmjMin)
Call DTPicker_Control(txtSelect_Options_Log_AmjMax, wAmjMax)

X = Trim(txtSelect_Options_Log_OPE)
If X <> "" Then xWhere = " and MOUVEMOPE = '" & X & "'"

xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSXOD0, " & paramIBM_Library_SABSPE & ".YBIAMVTHP" _
     & " where DOSXODDTR >= " & wAmjMin & " and DOSXODDTR <= " & wAmjMax _
     & " and mouvemeta = 1 and dosxodpie = mouvempie and dosxodecr = mouvemecr " & xWhere _
     & " order by DOSXODDTR,MOUVEMCOM"
Set rsSab = cnsab.Execute(xSql)

fgLOG_YDOSXOD0_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdSelect_SQL_YDOSNOK0()
Dim V, X As String
Dim xSql As String

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_YDOSNOK0"


xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSNOK0 " _
     & " order by DOSNOKOPE,DOSNOKNUM"
Set rsSab = cnsab.Execute(xSql)

fgLOG_YDOSNOK0_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub fgLOG_YDOSXOD0_Display()
Dim wColor As Long
Dim curX As Currency
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgLOG.Visible = False
fgLog_Reset

fgLOG.Rows = 1
fgLOG.FormatString = fgLog_FormatString
fgLOG.Row = 0

currentAction = "fgLog_Display"

Do While Not rsSab.EOF

    fgLOG.Rows = fgLOG.Rows + 1
    fgLOG.Row = fgLOG.Rows - 1
    fgLOG.Col = 0: fgLOG.Text = dateImp10(rsSab("DOSXODDTR"))
    fgLOG.Col = 1: fgLOG.Text = rsSab("MOUVEMCOM")
    fgLOG.Col = 2: fgLOG.Text = rsSab("MOUVEMOPE") & " " & Val(rsSab("MOUVEMNUM"))
    fgLOG.Col = 3: fgLOG.Text = rsSab("MOUVEMEVE")
    If rsSab("DOSXODNUM") <> 0 Then
       fgLOG.Col = 4: fgLOG.Text = rsSab("DOSXODOPE") & " " & rsSab("DOSXODNUM")
       fgLOG.CellForeColor = vbMagenta
    End If
    curX = -CCur(rsSab("MOUVEMMON")) '/ 1000
    fgLOG.Col = 5: fgLOG.Text = Format(Abs(curX), "### ### ### ##0.00")
    If curX >= 0 Then
       fgLOG.CellForeColor = vbBlue
    Else
       fgLOG.CellForeColor = vbRed
    End If
    
    fgLOG.Col = 6: fgLOG.Text = rsSab("COMPTEDEV")
    fgLOG.Col = 7: fgLOG.Text = rsSab("DOSXODLIB")
    fgLOG.Col = 8: fgLOG.Text = rsSab("DOSXODUUSR")
    If Trim(rsSab("DOSXODUUSR")) <> "BIA_AUTO" Then fgLOG.CellForeColor = vbMagenta
    fgLOG.Col = 9: fgLOG.Text = dateImp10(rsSab("DOSXODUAMJ")) & " " & timeImp8(rsSab("DOSXODUHMS"))
    fgLOG.Col = 10: fgLOG.Text = rsSab("DOSXODPIE") & "-" & rsSab("DOSXODECR")
    
    fgLOG.Col = fgLog_arrIndex: fgLOG.Text = I
    rsSab.MoveNext
Loop

fgLOG.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgLOG.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub fgLOG_YDOSNOK0_Display()
Dim wColor As Long
Dim curX As Currency
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgLOG.Visible = False
fgLog_Reset

fgLOG.Rows = 1
fgLOG.FormatString = "<Opé |>Numéro       |<Devise|<PCI          |<Client          |>Correction Compta|>Correction Gestion|<Commentaire                                  |<Màj par                          |< le                                    ||"
fgLOG.Row = 0

currentAction = "fgLog_Display"

Do While Not rsSab.EOF

    fgLOG.Rows = fgLOG.Rows + 1
    fgLOG.Row = fgLOG.Rows - 1
    fgLOG.Col = 0: fgLOG.Text = rsSab("DOSNOKOPE")
    fgLOG.Col = 1: fgLOG.Text = rsSab("DOSNOKNUM")
    fgLOG.Col = 2: fgLOG.Text = rsSab("DOSNOKDEV")
    fgLOG.Col = 3: fgLOG.Text = rsSab("DOSNOKPCI")
    fgLOG.Col = 4: fgLOG.Text = rsSab("DOSNOKCLI")
    curX = CCur(rsSab("DOSNOKMSD"))
    fgLOG.Col = 5: fgLOG.Text = Format(Abs(curX), "### ### ### ##0.00")
    If curX >= 0 Then
       fgLOG.CellForeColor = vbBlue
    Else
       fgLOG.CellForeColor = vbRed
    End If
    curX = CCur(rsSab("DOSNOKGSD"))
    fgLOG.Col = 6: fgLOG.Text = Format(Abs(curX), "### ### ### ##0.00")
    If curX >= 0 Then
       fgLOG.CellForeColor = vbBlue
    Else
       fgLOG.CellForeColor = vbRed
    End If
    fgLOG.Col = 7: fgLOG.Text = rsSab("DOSNOKUTXT")
    fgLOG.Col = 8: fgLOG.Text = rsSab("DOSNOKUUSR")
    fgLOG.Col = 9: fgLOG.Text = dateImp10(rsSab("DOSNOKUAMJ")) & " " & timeImp8(rsSab("DOSNOKUHMS"))
    
    fgLOG.Col = fgLog_arrIndex: fgLOG.Text = I
    rsSab.MoveNext
Loop

fgLOG.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgLOG.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub txtSelect_DOSSLDNUM_LostFocus()
Call txt_LostFocus(txtSelect_DOSSLDNUM)

End Sub

Private Sub txtSelect_Options_3uti_AmjMax_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_Options_3uti_AmjMin_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_Options_Log_OPE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub



Public Sub cboSelect_Options_3uti_UTI_Load()
Dim X As String
     
cboSelect_Options_3uti_UTI.Clear
cboSelect_Options_3uti_UTI.AddItem ""
' X = "select MNURUTUTI,MNURUTCUT from " & paramIBM_Library_SAB_P & ".ZMNURUT0 , " & paramIBM_Library_SAB_P & ".ZMNUUTI0" _
'     & " where MNURUTLOG = 'O' and mnuutigr2 like 'G_SOBI%' and MNUUTICUT = MNURUTCUT order by mnurututi"
X = "select MNURUTUTI , MNURUTCUT from " & paramIBM_Library_SAB & ".ZCDOUTI0 , " _
                        & paramIBM_Library_SAB & ".ZMNURUT0  " _
     & " where CDOUTIEVE = '03' and CDOUTIATT = '01' and CDOUTIETA = '02' and CDOUTICOP = 'CDE'" _
     & " and CDOUTIVA1 = MNURUTCUT" _
     & " group by MNURUTUTI , MNURUTCUT order by MNURUTUTI "

Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    cboSelect_Options_3uti_UTI.AddItem rsSab("MNURUTUTI") & " | " & rsSab("MNURUTCUT")
    rsSab.MoveNext
Loop

cboSelect_Options_3uti_CDODOSNOT.Clear
cboSelect_Options_3uti_CDODOSNOT.AddItem ""
 X = "select distinct CDODOSNOT  , CLIENASIG from " & paramIBM_Library_SAB_P & ".ZCDODOS0 ," & paramIBM_Library_SAB_P & ".ZCLIENA0" _
     & " where CDODOSEVE <> '90'" _
     & " and CDODOSNOT = CLIENACLI order by CLIENASIG"

Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    cboSelect_Options_3uti_CDODOSNOT.AddItem Val(rsSab("CDODOSNOT")) & " - " & rsSab("CLIENASIG")
    rsSab.MoveNext
Loop



End Sub

Public Sub cmdSelect_SQL_JPL()
On Error Resume Next
Dim X As String, xDest As String

            cmdSelect_SQL_XE1an
            cmdSendMail_SAB_Dossier "CDO_SQL_XE1", "BIA-CDO-Engagement1an"
Exit Sub

Call cmdSelect_SQL_5réfext("Jour")
'If fgSelect.Rows > 1 Then
    'Call mailAdresse_Production_Control("CHIBAB;CHARTIER;REOL_CH;OURY", xDest)
    xDest = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire("S10")

    X = "Liste des dossiers CDO créés le " & dateImp10_S(YBIATAB0_DATE_CPT_J) _
      & ",<BR> dont la référence externe est identique à celle de dossiers non clos."
    Call MSFlexGrid_SendMail(xDest, "CDO_Doublon", "Liste des nouveaux dossiers CDO 'doublons' - " & dateImp10_S(YBIATAB0_DATE_CPT_J), X, fgSelect, 4)
'End If

Exit Sub




If frmYGOSDOS0.hwnd = 0 Then
    Bia_swift_Monitor.frmYGOSDOS0_Show
End If
frmYGOSDOS0.Hide
frmYGOSDOS0.Msg_Rcv "BIA_GOS     " & "SQL_3: 201"

frmYGOSDOS0.Show


End Sub
Public Sub frmYGOSDOS0_SQL_3(lGOSDOSIDD As Long)
On Error Resume Next
If frmYGOSDOS0.hwnd = 0 Then
    Bia_swift_Monitor.frmYGOSDOS0_Show
End If
frmYGOSDOS0.Hide
frmYGOSDOS0.Msg_Rcv "BIA_GOS     " & "SQL_3: " & lGOSDOSIDD

frmYGOSDOS0.Show


End Sub

Private Sub txtSelect_Options_Scan_Liste_AMJ_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_Options_Scan_Liste_AMJ_Click()
cmdSelect_Clear

End Sub



Public Sub cmdSelect_SQL_GAR_ECH()
Dim V, X As String
Dim xSql As String

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_GAR_ECH"

xSql = "select * from " & paramIBM_Library_SAB & ".ZCAUDOS0, " & paramIBM_Library_SAB & ".ZCLIENA0 " _
     & " where CAUDOSTRA < 4 and CLIENACLI = CAUDOSBEN order by CAUDOSDOS"
Set rsSab = cnsab.Execute(xSql)

fgX_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Public Sub fgCom_Display_Total()
If mCDOCOMMON > 0 Then
    fgCOM.Rows = fgCOM.Rows + 1
    fgCOM.Row = fgCOM.Rows - 1
    fgCOM.Col = 0: fgCOM.Text = mCDOCOMDOS
    fgCOM.CellFontBold = True: fgCOM.CellBackColor = mColor_GB: fgCOM.CellForeColor = mColor_Z0
    fgCOM.Col = 2: fgCOM.Text = "ECNFPT"
    fgCOM.CellFontBold = True: fgCOM.CellBackColor = mColor_GB: fgCOM.CellForeColor = mColor_Z0
    fgCOM.Col = 1: fgCOM.Text = "O"
    fgCOM.CellFontBold = True: fgCOM.CellBackColor = mColor_GB: fgCOM.CellForeColor = mColor_Z0
    fgCOM.Col = 3: fgCOM.Text = Format$(mCDOCOMMON, "### ### ##0.00")
    fgCOM.CellFontBold = True: fgCOM.CellBackColor = mColor_GB: fgCOM.CellForeColor = mColor_Z0
    mECNFPT_TOT = mCDOCOMMON
    mECNFPT_Row = fgCOM.Row
End If
End Sub

Public Sub cmdSelect_SQL_5_ECNFPT_CD7()
Dim K As Integer, xSql As String, X As String
Dim V, wDSIT As Long

On Error GoTo Error_Handler

Call lstErr_AddItem(lstErr, cmdContext, "> cmdSelect_SQL_5_ECNFPT_CD7 Début"): DoEvents

'calcul des commissions ECNFPT + màj YDOSCD70
wDSIT = YBIATAB0_DATE_CPT_J
If Mid$(YBIATAB0_DATE_CPT_J, 5, 2) <> Mid$(YBIATAB0_DATE_CPT_JS1, 5, 2) Then
    V = rsYBIATAB0_Read("DATE", "CAL", "M", X)
    If IsNull(V) Then wDSIT = X
End If


ReDim arrYDOSCD70(100)
arrYDOSCD70_Nb = 0
arrYDOSCD70_Max = 99
Call cmdSelect_SQL_5_ECNFPT_Com
SSTab2.Visible = False
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Call lstErr_AddItem(lstErr, cmdContext, "- cmdSelect_SQL_5_ECNFPT_CD7 màj YDOSCD70"): DoEvents

'V = cnSAB_Transaction("BeginTrans")
'If Not IsNull(V) Then GoTo Error_MsgBox

For K = 1 To arrYDOSCD70_Nb
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YDOSCD70 " _
     & " where DOSCD7DSIT= " & wDSIT _
     & " and DOSCD7OPE = '" & arrYDOSCD70(K).DOSCD7OPE & "'" _
     & " and DOSCD7NUM = " & arrYDOSCD70(K).DOSCD7NUM _
     & " and DOSCD7KCN = 'C'" _
     & " and DOSCD7KNAT = '2'" _
     & " and DOSCD7PCI = '707210'" _
     & " and DOSCD7DDEB = " & arrYDOSCD70(K).DOSCD7DDEB _
     & " and DOSCD7DFIN = " & arrYDOSCD70(K).DOSCD7DFIN
           
     Set rsSab = cnsab.Execute(xSql)
    If Not rsSab.EOF Then
        Call rsYDOSCD70_GetBuffer(rsSab, oldYDOSCD70)
        If oldYDOSCD70.DOSCD7MTD <> arrYDOSCD70(K).DOSCD7MTD Then
            X = "set DOSCD7STA = '@1' , DOSCD7MTD = " & cur_P(arrYDOSCD70(K).DOSCD7MTD)
            V = sqlYDOSCD70_Update_Field(oldYDOSCD70, X)
        End If
    End If
Next K

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdSelect_SQL_5_ECNFPT_CD7"
Exit_sub:
   
   ' If Not IsNull(V) Then
   '     V = cnSAB_Transaction("Rollback")
   ' Else
   '     V = cnSAB_Transaction("Commit")
   ' End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Call lstErr_AddItem(lstErr, cmdContext, "< cmdSelect_SQL_5_ECNFPT_CD7 terminé"): DoEvents

End Sub
