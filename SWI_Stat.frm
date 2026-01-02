VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSWI_Stat 
   AutoRedraw      =   -1  'True
   Caption         =   "SWI_Stat"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13575
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SWI_Stat.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10530
   ScaleWidth      =   13575
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
      Height          =   270
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
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "statistiques SWIFT"
      TabPicture(0)   =   "SWI_Stat.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "."
      TabPicture(1)   =   "SWI_Stat.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraSelect_Options_3"
      Tab(1).Control(1)=   "fraSelect_Options_2"
      Tab(1).Control(2)=   "txtFg"
      Tab(1).Control(3)=   "fraSwift"
      Tab(1).Control(4)=   "lstW"
      Tab(1).ControlCount=   5
      Begin VB.Frame fraSelect_Options_3 
         BackColor       =   &H00F0FFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2052
         Left            =   -74535
         TabIndex        =   80
         Top             =   6570
         Width           =   8085
         Begin VB.TextBox txtSelect3_MOUVEMCOM 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4770
            TabIndex        =   81
            Top             =   375
            Width           =   2010
         End
         Begin MSComCtl2.DTPicker txtSelect3_SWISABWAMJ_Min 
            Height          =   300
            Left            =   2205
            TabIndex        =   82
            Top             =   330
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
            Format          =   49086467
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin MSComCtl2.DTPicker txtSelect3_SWISABWAMJ_Max 
            Height          =   300
            Left            =   2205
            TabIndex        =   83
            Top             =   870
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
            Format          =   49086467
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.Label lblSelect3_X 
            Caption         =   "sélection des mouvements comptables : 00 TR TRF  Débit"
            Height          =   405
            Left            =   300
            TabIndex        =   86
            Top             =   1410
            Width           =   7305
         End
         Begin VB.Label lblSelect3_SWISABWAMJ 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Période comptable"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   285
            TabIndex        =   85
            Top             =   360
            Width           =   1380
         End
         Begin VB.Label lblSelect3_MOUVEMCOM 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Compte"
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
            Left            =   3570
            TabIndex        =   84
            Top             =   375
            Width           =   1005
         End
      End
      Begin VB.Frame fraSelect_Options_2 
         BackColor       =   &H00F0FFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2052
         Left            =   -74535
         TabIndex        =   70
         Top             =   3645
         Width           =   8085
         Begin VB.TextBox txtSelect2_SWISABWUSR 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4710
            TabIndex        =   77
            Top             =   990
            Width           =   2010
         End
         Begin VB.TextBox txtSelect2_SWISABWMTK 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4725
            MaxLength       =   3
            TabIndex        =   75
            Top             =   405
            Width           =   630
         End
         Begin MSComCtl2.DTPicker txtSelect2_SWISABWAMJ_Min 
            Height          =   300
            Left            =   1770
            TabIndex        =   72
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
            Format          =   49086467
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin MSComCtl2.DTPicker txtSelect2_SWISABWAMJ_Max 
            Height          =   300
            Left            =   1740
            TabIndex        =   73
            Top             =   960
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
            Format          =   49086467
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.Label lblSelect2_SWISABWUSR 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Utilisateur"
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
            Left            =   3720
            TabIndex        =   76
            Top             =   1005
            Width           =   1005
         End
         Begin VB.Label lblSelect2_SWISABWMTK 
            BackColor       =   &H00D0F0FF&
            Caption         =   "MTxxx"
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
            Left            =   3720
            TabIndex        =   74
            Top             =   390
            Width           =   1035
         End
         Begin VB.Label lblSelect2_SWISABWAMJ 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Date SAA"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   285
            TabIndex        =   71
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.TextBox txtFg 
         Height          =   1260
         Left            =   -74475
         MultiLine       =   -1  'True
         TabIndex        =   69
         Top             =   1290
         Visible         =   0   'False
         Width           =   6732
      End
      Begin VB.Frame fraSwift 
         BackColor       =   &H00C0E0FF&
         Height          =   7050
         Left            =   -68070
         TabIndex        =   64
         Top             =   1005
         Visible         =   0   'False
         Width           =   6200
         Begin VB.CheckBox chkSIDE_DB_Show 
            BackColor       =   &H00C0FFFF&
            Caption         =   "afficher le message et l'historique du traitement SAA"
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
            Left            =   60
            TabIndex        =   67
            Top             =   600
            Width           =   6050
         End
         Begin VB.CheckBox chkSAB_Dossier_DB_Show 
            BackColor       =   &H0080C0FF&
            Caption         =   "afficher les écrirures comptables"
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
            Left            =   60
            TabIndex        =   66
            Top             =   930
            Width           =   6050
         End
         Begin MSFlexGridLib.MSFlexGrid fgSwift 
            Height          =   5610
            Left            =   60
            TabIndex        =   65
            Top             =   1260
            Width           =   6045
            _ExtentX        =   10663
            _ExtentY        =   9895
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16777215
            ForeColor       =   12582912
            BackColorFixed  =   16777168
            ForeColorFixed  =   16711680
            BackColorBkg    =   16777215
            GridColor       =   12632064
            GridColorFixed  =   12632064
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   "<Code |<Valeur                                                                                                |"
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
         Begin VB.Label libSWIFT_SWISABSWID 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   68
            Top             =   210
            Width           =   6050
         End
      End
      Begin VB.ListBox lstW 
         BackColor       =   &H00E0F0FF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1500
         Left            =   -73455
         TabIndex        =   12
         Top             =   345
         Visible         =   0   'False
         Width           =   3015
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
         Height          =   9420
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   13296
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   6400
            Left            =   690
            TabIndex        =   11
            Top             =   3075
            Visible         =   0   'False
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   11298
            _Version        =   393216
            Cols            =   17
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   15794175
            ForeColor       =   16711680
            BackColorFixed  =   8421504
            ForeColorFixed  =   -2147483633
            BackColorBkg    =   -2147483633
            FormatString    =   $"SWI_Stat.frx":0342
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
         Begin VB.ComboBox cboSelect_SQL 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   7365
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   2265
            Width           =   4590
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Rechercher"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   11880
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2265
            Width           =   1212
         End
         Begin VB.Frame fraSelect_Options_1 
            BackColor       =   &H00F0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2052
            Left            =   0
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   13152
            Begin VB.ComboBox cboSelect_SWISAB_5 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   324
               Left            =   11400
               Style           =   2  'Dropdown List
               TabIndex        =   53
               Top             =   1600
               Width           =   1572
            End
            Begin VB.ComboBox cboSelect_SWISAB_4 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   324
               Left            =   11400
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Top             =   1300
               Width           =   1572
            End
            Begin VB.ComboBox cboSelect_SWISAB_3 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   324
               Left            =   11400
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Top             =   1000
               Width           =   1572
            End
            Begin VB.ComboBox cboSelect_SWISAB_2 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   324
               Left            =   11400
               Style           =   2  'Dropdown List
               TabIndex        =   42
               Top             =   700
               Width           =   1572
            End
            Begin VB.ComboBox cboSelect_SWISAB_1 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   324
               Left            =   11400
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   400
               Width           =   1572
            End
            Begin VB.Frame fraSelect_Options_1A 
               BackColor       =   &H00D0F0FF&
               BorderStyle     =   0  'None
               Height          =   1812
               Left            =   120
               TabIndex        =   9
               Top             =   120
               Width           =   10932
               Begin VB.ComboBox cboSelect_SWISABWBIC_Pays 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   1200
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   79
                  Top             =   1395
                  Width           =   1250
               End
               Begin VB.ComboBox cboSelect_SWISABWMTK 
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
                  Left            =   3960
                  Sorted          =   -1  'True
                  TabIndex        =   63
                  Text            =   "WMTK"
                  Top             =   105
                  Width           =   990
               End
               Begin VB.ComboBox cboSelect_SWISABWES 
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
                  Left            =   1200
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   62
                  Top             =   870
                  Width           =   1250
               End
               Begin VB.ComboBox cboSelect_SWISABW59Z_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   9240
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   58
                  Top             =   1440
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWISABW50Z_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   9240
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   57
                  Top             =   480
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWISABW59Z 
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
                  Left            =   9840
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   56
                  Top             =   1440
                  Width           =   972
               End
               Begin VB.ComboBox cboSelect_SWISABW50Z 
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
                  Left            =   9840
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   55
                  Top             =   480
                  Width           =   972
               End
               Begin VB.ComboBox cboSelect_SWISABW57A 
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
                  Left            =   6600
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   40
                  Top             =   1440
                  Width           =   1812
               End
               Begin VB.ComboBox cboSelect_SWISABW52A 
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
                  Left            =   6600
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   39
                  Top             =   1080
                  Width           =   1812
               End
               Begin VB.ComboBox cboSelect_SWISABWBIC 
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
                  Left            =   6600
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   38
                  Top             =   120
                  Width           =   1812
               End
               Begin VB.ComboBox cboSelect_SWISABWDEV 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   3960
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   37
                  Top             =   555
                  Width           =   990
               End
               Begin VB.ComboBox cboSelect_SWISABW59P 
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
                  Left            =   9840
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   36
                  Top             =   1080
                  Width           =   972
               End
               Begin VB.ComboBox cboSelect_SWISABW50P 
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
                  Left            =   9840
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   35
                  Top             =   120
                  Width           =   972
               End
               Begin VB.ComboBox cboSelect_SWISABWEBA 
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
                  Left            =   6600
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   34
                  Top             =   480
                  Width           =   972
               End
               Begin VB.ComboBox cboSelect_SWISABWDEV_K 
                  ForeColor       =   &H00C00000&
                  Height          =   330
                  Left            =   3360
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   33
                  Top             =   585
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWISABW59P_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   9240
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   32
                  Top             =   1080
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWISABW50P_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   9240
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   31
                  Top             =   120
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWISABWEBA_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   6000
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   30
                  Top             =   480
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWISABW57A_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   6000
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   29
                  Top             =   1440
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWISABW52A_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   6000
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   28
                  Top             =   1080
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWISABWBIC_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   6000
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   27
                  Top             =   120
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWISABSSE 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   3960
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   19
                  Top             =   1440
                  Width           =   990
               End
               Begin VB.ComboBox cboSelect_SWISABSSE_K 
                  ForeColor       =   &H00C00000&
                  Height          =   330
                  Left            =   3360
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   18
                  Top             =   1440
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWISABOPEC_K 
                  ForeColor       =   &H00C00000&
                  Height          =   312
                  Left            =   3360
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   960
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWISABWMTK_K 
                  ForeColor       =   &H00C00000&
                  Height          =   330
                  Left            =   3360
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   15
                  Top             =   90
                  Width           =   576
               End
               Begin VB.ComboBox cboSelect_SWISABOPEC 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   3960
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   13
                  Top             =   960
                  Width           =   990
               End
               Begin MSComCtl2.DTPicker txtSelect_SWISABWAMJ_Min 
                  Height          =   300
                  Left            =   1200
                  TabIndex        =   50
                  Top             =   120
                  Width           =   1250
                  _ExtentX        =   2196
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
                  Format          =   49086467
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin MSComCtl2.DTPicker txtSelect_SWISABWAMJ_Max 
                  Height          =   300
                  Left            =   1200
                  TabIndex        =   51
                  Top             =   480
                  Width           =   1250
                  _ExtentX        =   2196
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
                  Format          =   49086467
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin VB.Label lblSelect_SWISABWBIC_Pays 
                  Alignment       =   2  'Center
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "Pays E/S"
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
                  Left            =   120
                  TabIndex        =   78
                  Top             =   1440
                  Width           =   900
               End
               Begin VB.Label lblSelect_SWISABWES 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0FF&
                  Caption         =   "Sens"
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
                  Left            =   120
                  TabIndex        =   61
                  Top             =   900
                  Width           =   900
               End
               Begin VB.Label lblSelect_SWISABW59Z 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "FR UE **"
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
                  Left            =   8520
                  TabIndex        =   60
                  Top             =   1500
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWISABW50Z 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "FR UE **"
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
                  Left            =   8520
                  TabIndex        =   59
                  Top             =   550
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWISABWAMJ 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFC0FF&
                  Caption         =   "Date SAA"
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
                  Left            =   120
                  TabIndex        =   49
                  Top             =   120
                  Width           =   900
               End
               Begin VB.Label lblSelect_SWISABWDEV 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "Devise"
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
                  Left            =   2600
                  TabIndex        =   26
                  Top             =   615
                  Width           =   615
               End
               Begin VB.Label lblSelect_SWISABW59P 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "Pays BEN"
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
                  Left            =   8520
                  TabIndex        =   25
                  Top             =   1150
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWISABW50P 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "Pays DO"
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
                  Left            =   8520
                  TabIndex        =   24
                  Top             =   200
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWISABWEBA 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "Routage"
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
                  Left            =   5280
                  TabIndex        =   23
                  Top             =   600
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWISABW57A 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "BIC BEN"
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
                  Left            =   5280
                  TabIndex        =   22
                  Top             =   1500
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWISABW52A 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "BIC DO"
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
                  Left            =   5280
                  TabIndex        =   21
                  Top             =   1080
                  Width           =   612
               End
               Begin VB.Label lblSelect_SWISABWBIC 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "BIC E/S"
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
                  Left            =   5280
                  TabIndex        =   20
                  Top             =   240
                  Width           =   732
               End
               Begin VB.Label lblSelect_SWISABSSE 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "Service"
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
                  Left            =   2600
                  TabIndex        =   17
                  Top             =   1500
                  Width           =   612
               End
               Begin VB.Label lblSelect_SWISABOPEC 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "Code opé"
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
                  Left            =   2600
                  TabIndex        =   14
                  Top             =   990
                  Width           =   735
               End
               Begin VB.Label lblSelect_SWISABWMTK 
                  BackColor       =   &H00D0F0FF&
                  Caption         =   "MTxxx"
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
                  Left            =   2600
                  TabIndex        =   10
                  Top             =   165
                  Width           =   615
               End
            End
            Begin VB.Label libSelect_SWISAB 
               Alignment       =   2  'Center
               BackColor       =   &H00F0FFFF&
               Caption         =   "Tri- Ventilation "
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   11400
               TabIndex        =   54
               Top             =   120
               Width           =   1452
            End
            Begin VB.Label libSelect_SWIDOS_5 
               BackColor       =   &H00F0FFFF&
               Caption         =   "5 - "
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   11160
               TabIndex        =   52
               Top             =   1650
               Width           =   252
            End
            Begin VB.Label libSelect_SWIDOS_4 
               BackColor       =   &H00F0FFFF&
               Caption         =   "4 - "
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   11160
               TabIndex        =   48
               Top             =   1350
               Width           =   252
            End
            Begin VB.Label libSelect_SWIDOS_3 
               BackColor       =   &H00F0FFFF&
               Caption         =   "3 - "
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   11160
               TabIndex        =   47
               Top             =   1050
               Width           =   252
            End
            Begin VB.Label libSelect_SWIDOS_2 
               BackColor       =   &H00F0FFFF&
               Caption         =   "2 - "
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   11160
               TabIndex        =   46
               Top             =   750
               Width           =   252
            End
            Begin VB.Label libSelect_SWIDOS_1 
               BackColor       =   &H00F0FFFF&
               Caption         =   "1 - "
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   11160
               TabIndex        =   45
               Top             =   450
               Width           =   252
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6700
            Left            =   15
            TabIndex        =   5
            Top             =   2655
            Visible         =   0   'False
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   11827
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16777215
            ForeColor       =   16711680
            BackColorFixed  =   8421376
            ForeColorFixed  =   12648447
            BackColorSel    =   12648384
            BackColorBkg    =   15790320
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   "< 1    |< 2       |> Nb               |> Montant           "
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
      End
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
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
      Left            =   13080
      Picture         =   "SWI_Stat.frx":0456
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
   Begin VB.Menu mnuPrint 
      Caption         =   "mnuPrint"
      Visible         =   0   'False
      Begin VB.Menu mnuExportation 
         Caption         =   "Exportation .xlsx"
      End
   End
   Begin VB.Menu mnuPrint2 
      Caption         =   "mnuPrint2"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint2_Excel 
         Caption         =   "Excel"
      End
      Begin VB.Menu mnuPrint2_Mail 
         Caption         =   "Envoi mail"
      End
   End
End
Attribute VB_Name = "frmSWI_Stat"
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
Dim arrHab(19) As Boolean
Dim blnAuto As Boolean, blnError As Boolean
Dim cmdSelect_SQL_K As String

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean
Dim fgSelect_BackColorFixed As Long, fgSelect_ForeColorFixed As Long, fgSelect_ForeColor As Long, fgSelect_BackColor As Long

'______________________________________________________________________
Dim wAmj8_tiret As String, xAmj8_from_crea_date_time As String, xAmj8_to_crea_date_time As String


Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim xYSWISAB0 As typeYSWISAB0, newYSWISAB0 As typeYSWISAB0, oldYSWISAB0 As typeYSWISAB0
Dim arrYSWISAB0() As typeYSWISAB0, arrYSWISAB0_Nb As Long, arrYSWISAB0_Max As Long, arrYSWISAB0_Index As Long
Dim xYSWISAB1 As typeYSWISAB1

Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean


Dim arrSWISAB_Field(13) As String, arrSWISAB_Lib(13) As String, arrSWISAB_Field_Nb As Integer
Dim arrSWISAB_Group(13) As Integer, arrSWISAB_Group_Nb As Integer
Dim arrSWISAB1(13) As Boolean
Dim mGroupBy As String
Dim xWhere_SQL As String
Dim blnDevise_Sum As Boolean

Dim cnSIDE_DB As New ADODB.Connection, rsSIDE_DB As New ADODB.Recordset
Dim fgSwift_FormatString As String
Dim xrText As typerText

Dim HeightOfLine As Long, LinesOfText As Long
Dim Mesg_aid As Long, mesg_s_umidl As Long, mesg_s_umidh As Long
Dim mSWISABSWID As Long
Dim mMOUVEMSER As String, mMOUVEMSSE As String, mMOUVEMOPE As String, mMOUVEMNUM As Long
Dim mSWISABSWID_Xd As Long
Dim cmdSelect_SQL_1_rText As String, blnSelect_SQL_1_rText As Boolean

Dim rtextField_Value As Variant

Dim Nb_E(20) As Long, Nb_S(20) As Long

Dim rsSabX As New ADODB.Recordset

'______________________________________________________________________
Private Sub fgSelect_Display()
Dim wColor As Long

Dim I As Long, K As Integer
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect.Cols = arrSWISAB_Group_Nb + 2
fgSelect_Reset

fgSelect.Rows = 1
fgSelect_FormatString = "<" & arrSWISAB_Lib(arrSWISAB_Group(1))
For K = 2 To arrSWISAB_Group_Nb
    fgSelect_FormatString = fgSelect_FormatString & "|<" & arrSWISAB_Lib(arrSWISAB_Group(K))
Next K

fgSelect.FormatString = fgSelect_FormatString & "|>        Nombre |>                      Montant"
'fgSelect_FormatString
fgSelect.Width = Len(fgSelect.FormatString) * 80
fgSelect.Row = 0

currentAction = "fgSelect_Display"

Do While Not rsSab.EOF
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine I, True

    rsSab.MoveNext
Loop
    

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYSWISAB0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long, blnYSWISAB0 As Boolean)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim xSql As String, xCur As Currency
On Error Resume Next

For K = 0 To arrSWISAB_Group_Nb - 1

    fgSelect.Col = K: fgSelect.Text = rsSab(K)
Next K
fgSelect.Col = arrSWISAB_Group_Nb: fgSelect.Text = Format(rsSab(arrSWISAB_Group_Nb), "##### ###")
xCur = rsSab(arrSWISAB_Group_Nb + 1)
If blnDevise_Sum Then
    If xCur <> 0 Then fgSelect.Col = arrSWISAB_Group_Nb + 1: fgSelect.Text = Format(xCur, "### ### ### ##0.00")
End If

'fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
End Sub

Private Sub fgSelect2_Display()
Dim wColor As Long

Dim I As Long, K As Integer
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1

fgSelect.FormatString = "Utilisateur         |>   Type|>          Nombre "
'fgSelect_FormatString
fgSelect.Width = Len(fgSelect.FormatString) * 80
fgSelect.Row = 0

currentAction = "fgSelect_Display"

Do While Not rsSIDE_DB.EOF
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect2_DisplayLine

    rsSIDE_DB.MoveNext
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
Private Sub fgSelect3_Display()
Dim wColor As Long

Dim I As Long, K As Integer
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1

fgSelect.FormatString = "<Compte              |<date             |<N/réf                        |<L/réf                            |>Montant        |<Devise" _
    & " |<DO                      |<DO intitulé                                                |<DO |<DO Banque   " _
    & " |<BEN Banque                        |<BEN |<BEN                                     |<BEN intitulé                                                      " _
    & " |<Motif                        |<Motif intitulé                                               " _
    & " |<Libellé comptable                                                          "
fgSelect.Width = fraTab0.Width - 200
fgSelect.Row = 0
fgSelect.Col = 4: fgSelect.CellAlignment = 1
currentAction = "fgSelect_Display"

Do While Not rsSab.EOF
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect3_DisplayLine

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




Private Sub fgSelect2sf_Display()
Dim wColor As Long

Dim I As Long, K As Integer
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 3

fgSelect.FormatString = "<Sens                     |>Total =   |>Tech |>Non Ctrl|>No Violation|>Détection SafeWatch|>Violation Accepted|>Acceptation SecFin |>Rejet SecFin|>NAK|>Annulés|>Live"
'fgSelect_FormatString
fgSelect.Width = Len(fgSelect.FormatString) * 80
fgSelect.Row = 0

currentAction = "fgSelect2ef_Display"

fgSelect.Row = 1
fgSelect.Col = 0: fgSelect.Text = "Sortants"
fgSelect.Col = 1: fgSelect.Text = Nb_S(1)
fgSelect.Col = 2: fgSelect.Text = Nb_S(2)
fgSelect.Col = 3: fgSelect.Text = Nb_S(8)
fgSelect.Col = 4: fgSelect.Text = Nb_S(3)
fgSelect.Col = 5: fgSelect.Text = Nb_S(4)
fgSelect.Col = 6: fgSelect.Text = Nb_S(5)
fgSelect.Col = 7: fgSelect.Text = Nb_S(6)
fgSelect.Col = 8: fgSelect.Text = Nb_S(7)
fgSelect.Col = 9: fgSelect.Text = Nb_S(9)
fgSelect.Col = 10: fgSelect.Text = Nb_S(10)
fgSelect.Col = 11: fgSelect.Text = Nb_S(11)


fgSelect.Row = 2
fgSelect.Col = 0: fgSelect.Text = "Entrants"
fgSelect.Col = 1: fgSelect.Text = Nb_E(1)
fgSelect.Col = 2: fgSelect.Text = Nb_E(2)
fgSelect.Col = 4: fgSelect.Text = Nb_E(3)
fgSelect.Col = 5: fgSelect.Text = Nb_E(4)
fgSelect.Col = 6: fgSelect.Text = Nb_E(5)
fgSelect.Col = 7: fgSelect.Text = Nb_E(6)

fgSelect.Col = 11: fgSelect.Text = Nb_E(11)

'________________________________________________________________________


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgSelect2sf_Display_Suite(lTxt As String)
Dim wColor As Long

Dim I As Long, K As Integer
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler

currentAction = "fgSelect2ef_Display_Suite"


'__________________________________________________________________________
fgSelect.Rows = fgSelect.Rows + 2
fgSelect.Row = fgSelect.Rows - 1
For K = 0 To 11
    fgSelect.Col = K
    fgSelect.CellBackColor = mColor_GB
    fgSelect.CellForeColor = vbWhite
Next K

fgSelect.Col = 0: fgSelect.Text = lTxt
fgSelect.Col = 1: fgSelect.Text = "Type MT"


Do While Not rsSIDE_DB.EOF
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect.Col = 0: If Not IsNull(rsSIDE_DB(0)) Then fgSelect.Text = rsSIDE_DB(0)
    fgSelect.Col = 1: If Not IsNull(rsSIDE_DB(1)) Then fgSelect.Text = rsSIDE_DB(1)
    fgSelect.Col = 2: If Not IsNull(rsSIDE_DB(2)) Then fgSelect.Text = rsSIDE_DB(2)

    rsSIDE_DB.MoveNext
Loop
'________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub arrYSWISAB0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYSWISAB0(101)
arrYSWISAB0_Max = 100: arrYSWISAB0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYSWISAB0_GetBuffer(rsSab, xYSWISAB0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYSWISAB0.fgselect_Display"
        '' Exit Sub
     Else
         arrYSWISAB0_Nb = arrYSWISAB0_Nb + 1
         If arrYSWISAB0_Nb > arrYSWISAB0_Max Then
             arrYSWISAB0_Max = arrYSWISAB0_Max + 100
             ReDim Preserve arrYSWISAB0(arrYSWISAB0_Max)
         End If
         
         arrYSWISAB0(arrYSWISAB0_Nb) = xYSWISAB0
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
Dim K As Integer
If blnControl Then
    cmdSelect_Clear
    cmdSelect_Ok.Visible = True
    K = InStr(cboSelect_SQL, "-")
    If K > 1 Then
        cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, K - 1))
    Else
        cmdSelect_SQL_K = "???"
    End If
    Select Case cmdSelect_SQL_K
        Case "1":
            fraSelect_Options_1.Visible = True: fraSelect_Options_1A.Visible = True: fraSelect_Options_2.Visible = False
            Case "2", "2sf": fraSelect_Options_2.Visible = True: fraSelect_Options_1.Visible = False
            Case "3trf": fraSelect_Options_3.Visible = True: fraSelect_Options_1.Visible = False
    End Select

End If

End Sub



Private Sub fgSwift_Display(lSWISABSWID As Long)
Dim wColor As Long, wColorFixed As Long
Dim X As String, xWhere As String, xOPE As String
Dim xSql As String
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String
Dim xUUMID As String
On Error GoTo Error_Handler


fraSwift.Visible = False
'fgswift_Reset

fgSwift.Rows = 1
fgSwift.FormatString = fgSwift_FormatString
fgSwift.Row = 0
fgSwift.RowHeight(0) = 700
currentAction = "fgswift_Display"
'mSWISABSWID = lSWISABSWID
'----------------------------------------------------------------
blnOk = True
If lSWISABSWID > 0 Then
    blnOk = False
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABSWID = " & lSWISABSWID
Else
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABWID1 = " & Mesg_aid _
         & " and SWISABWIDL = " & mesg_s_umidl _
         & " and SWISABWIDH = " & mesg_s_umidh
End If

    Set rsSab = cnsab.Execute(xSql)
    
    If Not rsSab.EOF Then

        blnOk = True
        Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
        Mesg_aid = oldYSWISAB0.SWISABWID1
        mesg_s_umidl = oldYSWISAB0.SWISABWIDL
        mesg_s_umidh = oldYSWISAB0.SWISABWIDH
        mMOUVEMSER = oldYSWISAB0.SWISABSER
        mMOUVEMSSE = oldYSWISAB0.SWISABSSE
        mMOUVEMOPE = oldYSWISAB0.SWISABOPEC
        mMOUVEMNUM = oldYSWISAB0.SWISABOPEN

    End If

'----------------------------------------------------------------
If Not blnOk Then
    Call rsYSWISAB0_Init(oldYSWISAB0)
    libSWIFT_SWISABSWID = " !!! inconnu dans YSWISAB0 et SAA !!!!!!!!!!!!!!!"
    xSql = "select * from rMesg " _
        & "where Aid = " & Mesg_aid _
        & " and Mesg_s_umidl = " & mesg_s_umidl _
        & " and Mesg_s_umidh  =  " & mesg_s_umidh
   Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
   
    If Not rsSIDE_DB.EOF Then
         If Not IsNull(rsSIDE_DB("mesg_type")) Then
            oldYSWISAB0.SWISABWMTK = rsSIDE_DB("mesg_type")
        Else
            oldYSWISAB0.SWISABWMTK = "XXX"
         End If
         xUUMID = rsSIDE_DB("mesg_uumid")
         If Mid$(xUUMID, 1, 1) = "I" Then
             oldYSWISAB0.SWISABWES = "S"
         Else
             oldYSWISAB0.SWISABWES = "E"
         End If
        oldYSWISAB0.SWISABWBIC = Mid$(xUUMID, 2, 11)
        Call dateJma10_Amj(Mid$(rsSIDE_DB("mesg_crea_date_time"), 1, 10), X)
        oldYSWISAB0.SWISABWAMJ = Val(X)
        X = Mid$(rsSIDE_DB("mesg_crea_date_time"), 12, 8)
        oldYSWISAB0.SWISABWHMS = Val(Mid$(X, 1, 2) & Mid$(X, 4, 2) & Mid$(X, 7, 2))

    End If
End If
'--------------------------------------------------------------
    If oldYSWISAB0.SWISABWES = "E" Then
        X = "reçu de "
        wColor = RGB(190, 240, 255)
        wColorFixed = vbBlue
    Else
        X = "émis vers "
        wColor = RGB(220, 255, 220)
        wColorFixed = RGB(0, 64, 0)
    End If
    libSWIFT_SWISABSWID = "SAB : " & Trim(oldYSWISAB0.SWISABOPEC) & " " & Format(oldYSWISAB0.SWISABOPEN, "### ###")
    
    If cmdSelect_SQL_K = "1trf" Then
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB1 where SWISAB1ID = " & lSWISABSWID
        Set rsSab = cnsab.Execute(xSql)
    
        If Not rsSab.EOF Then
            libSWIFT_SWISABSWID = "D.Ordre : " & rsSab("SWISABW50P") & "  " & rsSab("SWISABW50Z") & "  " & rsSab("SWISABW52A") _
                                & "  - Bénéficiaire : " & rsSab("SWISABW59P") & "  " & rsSab("SWISABW59Z") & "  " & rsSab("SWISABW57A")
        End If
    End If
    fgSwift.Col = 0: fgSwift.Text = oldYSWISAB0.SWISABWMTK
    fgSwift.CellFontBold = True: fgSwift.CellBackColor = wColor
    fgSwift.ForeColorFixed = wColorFixed
    fgSwift.Col = 1: fgSwift.Text = X & oldYSWISAB0.SWISABWBIC & " le " & dateImp10(oldYSWISAB0.SWISABWAMJ) & " " & timeImp8(oldYSWISAB0.SWISABWHMS) _
                                  & vbCrLf & ZSWIBIC0_Select(oldYSWISAB0.SWISABWBIC)
    fgSwift.CellFontBold = True: fgSwift.CellBackColor = wColor
    fgSwift.ForeColorFixed = wColorFixed
    fraSwift.BackColor = wColor
    
   ' xSQL = "select field_code , field_option , field_cnt , cast(value as varchar) as value from rtextField " _

    xSql = "select *  from rtextField  " _
        & "where Aid = " & Mesg_aid _
        & " and text_s_umidl = " & mesg_s_umidl _
        & " and text_s_umidh  =  " & mesg_s_umidh _
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
            & "where Aid = " & Mesg_aid _
            & " and text_s_umidl = " & mesg_s_umidl _
            & " and text_s_umidh  =  " & mesg_s_umidh
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
        If Not rsSIDE_DB.EOF Then
            Call srvrText_GetBuffer_ODBC(rsSIDE_DB, xrText)
            fgSwift_DisplayLine_rText fgSwift.Row, wColor, wColorFixed
        End If
    End If
    fraSwift.Visible = True
'End If


    If chkSIDE_DB_Show Then Call frmSIDE_DB.fgSwift_Display(lSWISABSWID, Mesg_aid, mesg_s_umidl, mesg_s_umidh)
    
    If mMOUVEMNUM = 0 Then
        chkSAB_Dossier_DB_Show.Enabled = False
    Else
        chkSAB_Dossier_DB_Show.Enabled = True
        If chkSAB_Dossier_DB_Show Then Call frmSAB_Dossier_DB.Form_Init("", "", "", "", mMOUVEMSER, mMOUVEMSSE, mMOUVEMOPE, mMOUVEMNUM)
    End If


'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Function ZSWIBIC0_Select(lMsg As String) As String
Dim xSql As String
xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIBIC0 where SWIBICBIC like '" & Trim(lMsg) & "%' order by SWIBICBIC"
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then
    ZSWIBIC0_Select = Trim(rsSab("SWIBICIN1")) & "  " & Trim(rsSab("SWIBICVIL")) & "  " & Trim(rsSab("SWIBICCOM"))
Else
    ZSWIBIC0_Select = ""
End If

End Function


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
        If Len(fgSwift.Text) > 50 Then fgSwift.RowHeight(fgSwift.Row) = 500
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




Private Sub cmdSelect_SQL_1()
Dim V
Dim X As String, K As Integer
Dim xAnd As String, xSql As String
Dim blnOk As Boolean, blnYSWISAB1 As Boolean

On Error GoTo Error_Handler

currentAction = "cmdYSWISAB0_SQL"
blnOk = False
blnYSWISAB1 = False
blnDevise_Sum = False

Call DTPicker_Control(txtSelect_SWISABWAMJ_Min, wAmjMin)
Call DTPicker_Control(txtSelect_SWISABWAMJ_Max, wAmjMax)

xWhere_SQL = " Where SWISABWAMJ >= " & wAmjMin & " and SWISABWAMJ <= " & wAmjMax

If Trim(cboSelect_SWISABWMTK_K) <> "" Then
    X = Trim(cboSelect_SWISABWMTK)
    If InStr(X, "%") Then
        If Trim(cboSelect_SWISABWMTK_K) = "<>" Then
            xWhere_SQL = xWhere_SQL & " and SWISABWMTK NOT LIKE '" & X & "'"
        Else
            xWhere_SQL = xWhere_SQL & " and SWISABWMTK LIKE '" & X & "'"
        End If
    Else
        If InStr(X, ",") Then
            If Trim(cboSelect_SWISABWMTK_K) = "<>" Then
                xWhere_SQL = xWhere_SQL & " and SWISABWMTK  NOT IN ('" & Replace(X, ",", "','") & "')"
            Else
                xWhere_SQL = xWhere_SQL & " and SWISABWMTK  IN ('" & Replace(X, ",", "','") & "')"
            End If
        Else
            xWhere_SQL = xWhere_SQL & " and SWISABWMTK " & Trim(cboSelect_SWISABWMTK_K) & "'" & X & "'"
        End If
    End If
End If

If Trim(cboSelect_SWISABWES) <> "" Then xWhere_SQL = xWhere_SQL & " and   SWISABWES = '" & Mid$(cboSelect_SWISABWES, 1, 1) & "'"

X = Trim(cboSelect_SWISABWBIC_Pays)
If X <> "" Then
    xWhere_SQL = xWhere_SQL & " and   substring(SWISABWBIC,5,2) ='" & Mid$(X, 1, 2) & "'"
End If

If Trim(cboSelect_SWISABOPEC_K) <> "" Then xWhere_SQL = xWhere_SQL & " and   SWISABOPEC " & Trim(cboSelect_SWISABOPEC_K) & "'" & cboSelect_SWISABOPEC & "'"
If Trim(cboSelect_SWISABWDEV_K) <> "" Then xWhere_SQL = xWhere_SQL & " and   SWISABWDEV " & Trim(cboSelect_SWISABWDEV_K) & "'" & cboSelect_SWISABWDEV & "'"
If Trim(cboSelect_SWISABSSE_K) <> "" Then xWhere_SQL = xWhere_SQL & " and   SWISABSSE " & Trim(cboSelect_SWISABSSE_K) & "'" & cboSelect_SWISABSSE & "'"
X = Trim(cboSelect_SWISABWBIC_K)
Select Case X
    Case "":
    Case "=", "<>": xWhere_SQL = xWhere_SQL & " and   SWISABWBIC " & Trim(cboSelect_SWISABWBIC_K) & "'" & cboSelect_SWISABWBIC & "'"
    Case "=4": xWhere_SQL = xWhere_SQL & " and   SWISABWBIC like '" & Mid$(cboSelect_SWISABWBIC, 1, 4) & "%'"
    Case "=6": xWhere_SQL = xWhere_SQL & " and   SWISABWBIC like '" & Mid$(cboSelect_SWISABWBIC, 1, 6) & "%'"
    Case "=8": xWhere_SQL = xWhere_SQL & " and   SWISABWBIC like '" & Mid$(cboSelect_SWISABWBIC, 1, 8) & "%'"
End Select

X = Trim(cboSelect_SWISABW52A_K)
If X <> "" Then
    Select Case X
        Case "=", "<>": xWhere_SQL = xWhere_SQL & " and   SWISABW52A " & Trim(cboSelect_SWISABW52A_K) & "'" & cboSelect_SWISABW52A & "'"
        Case "=4": xWhere_SQL = xWhere_SQL & " and   SWISABW52A like '" & Mid$(cboSelect_SWISABW52A, 1, 4) & "%'"
        Case "=6": xWhere_SQL = xWhere_SQL & " and   SWISABW52A like '" & Mid$(cboSelect_SWISABW52A, 1, 6) & "%'"
        Case "=8": xWhere_SQL = xWhere_SQL & " and   SWISABW52A like '" & Mid$(cboSelect_SWISABW52A, 1, 8) & "%'"
    End Select
    blnYSWISAB1 = True
End If

X = Trim(cboSelect_SWISABW57A_K)
If X <> "" Then
    Select Case X
        Case "=", "<>": xWhere_SQL = xWhere_SQL & " and   SWISABW57A " & Trim(cboSelect_SWISABW57A_K) & "'" & cboSelect_SWISABW57A & "'"
        Case "=4": xWhere_SQL = xWhere_SQL & " and   SWISABW57A like '" & Mid$(cboSelect_SWISABW57A, 1, 4) & "%'"
        Case "=6": xWhere_SQL = xWhere_SQL & " and   SWISABW57A like '" & Mid$(cboSelect_SWISABW57A, 1, 6) & "%'"
        Case "=8": xWhere_SQL = xWhere_SQL & " and   SWISABW57A like '" & Mid$(cboSelect_SWISABW57A, 1, 8) & "%'"
    End Select
    blnYSWISAB1 = True
End If

If Trim(cboSelect_SWISABWEBA_K) <> "" Then
    xWhere_SQL = xWhere_SQL & " and   SWISABWEBA " & Trim(cboSelect_SWISABWEBA_K) & "'" & cboSelect_SWISABWEBA & "'"
    blnYSWISAB1 = True
End If
If Trim(cboSelect_SWISABW50P_K) <> "" Then
    xWhere_SQL = xWhere_SQL & " and   SWISABW50P " & Trim(cboSelect_SWISABW50P_K) & "'" & cboSelect_SWISABW50P & "'"
    blnYSWISAB1 = True
End If
If Trim(cboSelect_SWISABW59P_K) <> "" Then
    xWhere_SQL = xWhere_SQL & " and   SWISABW59P " & Trim(cboSelect_SWISABW59P_K) & "'" & cboSelect_SWISABW59P & "'"
    blnYSWISAB1 = True
End If
If Trim(cboSelect_SWISABW50Z_K) <> "" Then
    xWhere_SQL = xWhere_SQL & " and   SWISABW50Z " & Trim(cboSelect_SWISABW50Z_K) & "'" & cboSelect_SWISABW50Z & "'"
    blnYSWISAB1 = True
End If
If Trim(cboSelect_SWISABW59Z_K) <> "" Then
    xWhere_SQL = xWhere_SQL & " and   SWISABW59Z " & Trim(cboSelect_SWISABW59Z_K) & "'" & cboSelect_SWISABW59Z & "'"
    blnYSWISAB1 = True
End If
   
blnOk = False
arrSWISAB_Group_Nb = 0
For K = 1 To arrSWISAB_Field_Nb
    If Trim(cboSelect_SWISAB_1) = arrSWISAB_Lib(K) Then
        arrSWISAB_Group_Nb = arrSWISAB_Group_Nb + 1
        arrSWISAB_Group(1) = K
        If arrSWISAB1(K) Then blnYSWISAB1 = True
        Exit For
    End If
Next K

If Trim(cboSelect_SWISAB_2) = "" Then
    blnOk = True
Else
    For K = 1 To arrSWISAB_Field_Nb
        If Trim(cboSelect_SWISAB_2) = arrSWISAB_Lib(K) Then
            arrSWISAB_Group_Nb = arrSWISAB_Group_Nb + 1
            arrSWISAB_Group(arrSWISAB_Group_Nb) = K
            If arrSWISAB1(K) Then blnYSWISAB1 = True
            Exit For
        End If
    Next K
End If
If Trim(cboSelect_SWISAB_3) = "" Then
    blnOk = True
Else
    For K = 1 To arrSWISAB_Field_Nb
        If Trim(cboSelect_SWISAB_3) = arrSWISAB_Lib(K) Then
            arrSWISAB_Group_Nb = arrSWISAB_Group_Nb + 1
            arrSWISAB_Group(arrSWISAB_Group_Nb) = K
            If arrSWISAB1(K) Then blnYSWISAB1 = True
            Exit For
        End If
    Next K
End If
If Trim(cboSelect_SWISAB_4) = "" Then
    blnOk = True
Else
    For K = 1 To arrSWISAB_Field_Nb
        If Trim(cboSelect_SWISAB_4) = arrSWISAB_Lib(K) Then
            arrSWISAB_Group_Nb = arrSWISAB_Group_Nb + 1
            arrSWISAB_Group(arrSWISAB_Group_Nb) = K
            If arrSWISAB1(K) Then blnYSWISAB1 = True
            Exit For
        End If
    Next K
End If
If Trim(cboSelect_SWISAB_5) = "" Then
    blnOk = True
Else
    For K = 1 To arrSWISAB_Field_Nb
        If Trim(cboSelect_SWISAB_5) = arrSWISAB_Lib(K) Then
            arrSWISAB_Group_Nb = arrSWISAB_Group_Nb + 1
            arrSWISAB_Group(arrSWISAB_Group_Nb) = K
            If arrSWISAB1(K) Then blnYSWISAB1 = True
            Exit For
        End If
    Next K
End If
   
mGroupBy = arrSWISAB_Field(arrSWISAB_Group(1))
For K = 2 To arrSWISAB_Group_Nb
    mGroupBy = mGroupBy & " , " & arrSWISAB_Field(arrSWISAB_Group(K))
Next K
If InStr(mGroupBy, "SWISABWDEV") > 0 Then blnDevise_Sum = True

If blnYSWISAB1 Then
    xSql = "select " & mGroupBy & " , count(*) , SUM(SWISABWMTD) from " & paramIBM_Library_SABSPE & ".YSWISAB0," & paramIBM_Library_SABSPE & ".YSWISAB1 " _
         & xWhere_SQL & " and SWISAB1ID = SWISABSWID" _
         & " group by  " & mGroupBy _
         & " order by " & mGroupBy

Else
    xSql = "select " & mGroupBy & " , count(*) , SUM(SWISABWMTD) from " & paramIBM_Library_SABSPE & ".YSWISAB0" _
         & xWhere_SQL _
         & " group by  " & mGroupBy _
         & " order by " & mGroupBy
End If
Set rsSab = cnsab.Execute(xSql)

fgSelect_Display


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_2()
Dim V
Dim X As String, K As Integer
Dim xAnd As String, xSql As String
Dim blnOk As Boolean, blnYSWISAB1 As Boolean

On Error GoTo Error_Handler

currentAction = "cmdYSWISAB0_SQL_2"
blnOk = False

Call DTPicker_Amj8_tiret(txtSelect2_SWISABWAMJ_Min, wAmj8_tiret)
xAmj8_from_crea_date_time = wAmj8_tiret
Call DTPicker_Amj8_tiret(txtSelect2_SWISABWAMJ_Max, wAmj8_tiret)
xAmj8_to_crea_date_time = wAmj8_tiret

xAnd = ""

If Trim(txtSelect2_SWISABWMTK) <> "" Then xAnd = " and mesg_type = '" & Trim(txtSelect2_SWISABWMTK) & "'"
If Trim(txtSelect2_SWISABWUSR) <> "" Then xAnd = " and intv_oper_nickname = '" & Trim(txtSelect2_SWISABWUSR) & "'"

mGroupBy = "intv_oper_nickname , mesg_type"
xSql = "select " & mGroupBy & " , count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & xAnd _
        & " and substring(Intv_merged_text,1,55) =  'Routed from rp [_MP_authorisation] to rp [_SI_to_SWIFT]'" _
        & " group by  " & mGroupBy _
        & " order by " & mGroupBy
   
   Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
   
'
fgSelect2_Display


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_3trf()
Dim V
Dim X As String, K As Integer
Dim xAnd As String, xSql As String
Dim blnOk As Boolean, blnYSWISAB1 As Boolean

On Error GoTo Error_Handler

currentAction = "cmdYSWISAB0_SQL_2"
blnOk = False

Call DTPicker_Control(txtSelect3_SWISABWAMJ_Min, wAmjMin)
Call DTPicker_Control(txtSelect3_SWISABWAMJ_Max, wAmjMax)
X = Trim(txtSelect3_MOUVEMCOM)
If X = "" Then
    V = "Préciser le compte"
    GoTo Error_MsgBox
End If

xAnd = " where MOUVEMCOM = '" & X & "' and MOUVEMDTR >= " & wAmjMin - 19000000 & " And MOUVEMDTR <= " & wAmjMax - 19000000 _
     & " and MOUVEMOPE = 'TRF' and MOUVEMSER = '00' and MOUVEMSSE = 'TR' and MOUVEMMON > 0"


'If Trim(txtSelect2_SWISABWMTK) <> "" Then xAnd = " and mesg_type = '" & Trim(txtSelect2_SWISABWMTK) & "'"
'If Trim(txtSelect2_SWISABWUSR) <> "" Then xAnd = " and intv_oper_nickname = '" & Trim(txtSelect2_SWISABWUSR) & "'"

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH" _
    & xAnd _
    & " order by MOUVEMDTR , MOUVEMNUM"
Set rsSab = cnsab.Execute(xSql)
   
'
fgSelect3_Display


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_JPL()
Dim V
Dim X As String, K As Integer
Dim xAnd As String, xSql As String
Dim blnOk As Boolean, blnYSWISAB1 As Boolean

On Error GoTo Error_Handler

currentAction = "cmdYSWISAB0_SQL_2"
blnOk = False
Call DTPicker_Amj8_tiret(txtSelect2_SWISABWAMJ_Min, wAmj8_tiret)
xAmj8_from_crea_date_time = wAmj8_tiret
Call DTPicker_Amj8_tiret(txtSelect2_SWISABWAMJ_Max, wAmj8_tiret)
xAmj8_to_crea_date_time = wAmj8_tiret


    xSql = "SELECT    count(*) From sysobjects, syscolumns " _
          & " WHERE  ( sysobjects.id = syscolumns.id) And  (sysobjects.xtype = 'U') " _
          & " AND sysobjects.name LIKE 'rappe'"
          
    xSql = "SELECT    * From sysobjects" _
          & " WHERE   (sysobjects.xtype = 'U') "
          
    xSql = "SELECT    syscolumns.name From sysobjects, syscolumns " _
          & " WHERE  ( sysobjects.id = syscolumns.id) And  (sysobjects.xtype = 'U') " _
          & " AND sysobjects.name LIKE 'sysobjects' ORDER BY syscolumns.colorder"
          
    xSql = "SELECT    syscolumns.name From sysobjects, syscolumns " _
          & " WHERE  ( sysobjects.id = syscolumns.id)  " _
          & " AND sysobjects.name LIKE 'rJrnl' ORDER BY syscolumns.colorder"
          
'xSQL = "select * from sysobjects where name like '%user%'"

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
Do While Not rsSIDE_DB.EOF
    Debug.Print rsSIDE_DB(0) ', rsSIDE_DB(1)     ', rsSIDE_DB(2), rsSIDE_DB(3), rsSIDE_DB(4), rsSIDE_DB(5), rsSIDE_DB(6), rsSIDE_DB(7)
    rsSIDE_DB.MoveNext
Loop
Exit Sub

xSql = "select distinct intv_oper_nickname from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 09:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _


   Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
Do While Not rsSIDE_DB.EOF
    Debug.Print rsSIDE_DB("intv_oper_nickname")
    rsSIDE_DB.MoveNext
Loop
'Exit Sub


'============================================================================================

'mGroupBy = "intv_oper_nickname , mesg_type"
xSql = "select * from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 09:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and intv_oper_nickname  =  'LAGARDE'" _
        
        '& " and  (substring(Intv_merged_text,1,51) =  'Routed from rp [OFCS_Validate] to rp [_SI_to_SWIFT]'" _
        '& " or  substring(Intv_merged_text,1,54) = 'Routed from rp [OFCS_Validate] to rp [_MP_verification]')"

      '          & " and Intv_inty_name  =  'Toolkit intervention'" _
      '  & " and Intv_mpfn_name  =  'OFCS_Detect'"
'& " and substring(Intv_merged_text,1,115) =  'Routed from rp [OFCS_IN] to rp [_SI_to_SWIFT]; On Processing by Function OFCS_Detect with result Violation_Accepted'"
'============================================================================================
   Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
   

Do While Not rsSIDE_DB.EOF
    Debug.Print rsSIDE_DB("mesg_uumid"), rsSIDE_DB("mesg_trn_ref"), rsSIDE_DB("mesg_crea_date_time")

    rsSIDE_DB.MoveNext
Loop


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_2sf()
Dim V
Dim X As String, K As Integer
Dim xAnd As String, xSql As String



On Error GoTo Error_Handler

currentAction = "cmdYSWISAB0_SQL_2"

Call DTPicker_Amj8_tiret(txtSelect2_SWISABWAMJ_Min, wAmj8_tiret)
xAmj8_from_crea_date_time = wAmj8_tiret
Call DTPicker_Amj8_tiret(txtSelect2_SWISABWAMJ_Max, wAmj8_tiret)
xAmj8_to_crea_date_time = wAmj8_tiret

cmdSelect_SQL_2sf_Init

'          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
 '         & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _

'_________________________________________________________________________________________________
mGroupBy = "mesg_sub_format "
xSql = "select " & mGroupBy & " , count(*) from rMesg  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
         & " group by  " & mGroupBy _
         & " order by " & mGroupBy

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

Do While Not rsSIDE_DB.EOF
    If rsSIDE_DB(0) = "INPUT" Then
        Nb_S(1) = rsSIDE_DB(1)
    Else
        Nb_E(1) = rsSIDE_DB(1)
    End If

    rsSIDE_DB.MoveNext
Loop

'_________________________________________________________________________________________________
mGroupBy = "mesg_sub_format "
xSql = "select " & mGroupBy & " , count(*) from rMesg  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
          & " and substring(mesg_type,1,1) = '0'" _
         & " group by  " & mGroupBy _
         & " order by " & mGroupBy

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

Do While Not rsSIDE_DB.EOF
    If rsSIDE_DB(0) = "INPUT" Then
        Nb_S(2) = rsSIDE_DB(1)
    Else
        Nb_E(2) = rsSIDE_DB(1)
    End If

    rsSIDE_DB.MoveNext
Loop
'_________________________________________________________________________________________________
mGroupBy = "mesg_sub_format "
xSql = "select " & mGroupBy & " , count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and Intv_inty_name  like   'Toolkit intervention %'" _
        & " and Intv_mpfn_name  =  'OFCS_Detect'" _
        & " group by  " & mGroupBy _
        & " order by " & mGroupBy

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

Do While Not rsSIDE_DB.EOF
    If rsSIDE_DB(0) = "INPUT" Then
        Nb_S(4) = rsSIDE_DB(1)
    Else
        Nb_E(4) = rsSIDE_DB(1)
    End If

    rsSIDE_DB.MoveNext
Loop
'_________________________________________________________________________________________________
xSql = "select count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and substring(Intv_merged_text,1,109) =  'Routed from rp [OFCS_IN] to rp [_SI_to_SWIFT]; On Processing by Function OFCS_Detect with result No_Violation'"


Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then Nb_S(3) = rsSIDE_DB(0)

'_________________________________________________________________________________________________
xSql = "select count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and substring(Intv_merged_text,1,115) =  'Routed from rp [OFCS_IN] to rp [_SI_to_SWIFT]; On Processing by Function OFCS_Detect with result Violation_Accepted'"

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then Nb_S(5) = rsSIDE_DB(0)
'_________________________________________________________________________________________________
xSql = "select count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and substring(Intv_merged_text,1,119) =  'Routed from rp [OFCS_IN] to rp [_MP_verification]; On Processing by Function OFCS_Detect with result Violation_Accepted'"

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then Nb_S(5) = Nb_S(5) + rsSIDE_DB(0)

'_________________________________________________________________________________________________
xSql = "select count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and  substring(Intv_merged_text,1,51) =  'Routed from rp [OFCS_Validate] to rp [_SI_to_SWIFT]'" '


Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then Nb_S(6) = rsSIDE_DB(0)
'_________________________________________________________________________________________________
xSql = "select count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and  substring(Intv_merged_text,1,55) = 'Routed from rp [OFCS_Validate] to rp [_MP_verification]'"


Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then Nb_S(6) = Nb_S(6) + rsSIDE_DB(0)
'_________________________________________________________________________________________________
xSql = "select count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and  substring(Intv_merged_text,1,56) = 'Routed from rp [OFCS_Validate] to rp [_MP_authorisation]'"


Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then Nb_S(6) = Nb_S(6) + rsSIDE_DB(0)

'_________________________________________________________________________________________________
xSql = "select count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and substring(Intv_merged_text,1,51) =  'Routed from rp [OFCS_Validate] to rp [_MP_mod_text]'"

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then Nb_S(7) = rsSIDE_DB(0)
'_________________________________________________________________________________________________
xSql = "select count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and substring(Intv_merged_text,1,54) =  'Disposed from rp [_AI_from_APPLI] to rp [_SI_to_SWIFT]'"

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then Nb_S(8) = rsSIDE_DB(0)
'_________________________________________________________________________________________________
xSql = "select count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and substring(Intv_merged_text,1,50) =  'Routed from rp [_SI_to_SWIFT] to rp [_MP_mod_text]'"

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then Nb_S(9) = rsSIDE_DB(0)

'_________________________________________________________________________________________________
xSql = "select count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and substring(Intv_merged_text,1,21) =  'Completed at rp [_MP_'"

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then Nb_S(10) = rsSIDE_DB(0)
'_________________________________________________________________________________________________

mGroupBy = "mesg_sub_format "
xSql = "select " & mGroupBy & " , count(*) from rMesg  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
         & " and mesg_status  <>  'COMPLETED'" _
         & " group by  " & mGroupBy _
         & " order by " & mGroupBy

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

Do While Not rsSIDE_DB.EOF
    If rsSIDE_DB(0) = "INPUT" Then
        Nb_S(11) = rsSIDE_DB(1)
    Else
        Nb_E(11) = rsSIDE_DB(1)
    End If

    rsSIDE_DB.MoveNext
Loop


'_________________________________________________________________________________________________
xSql = "select count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and substring(Intv_merged_text,1,107) =  'Routed from rp [OFCS_OUT] to rp [AutoRcvOK]; On Processing by Function OFCS_Detect with result No_Violation'"


Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then Nb_E(3) = rsSIDE_DB(0)


'_________________________________________________________________________________________________
xSql = "select count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and substring(Intv_merged_text,1,113) =  'Routed from rp [OFCS_OUT] to rp [AutoRcvOK]; On Processing by Function OFCS_Detect with result Violation_Accepted'"


Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then Nb_E(5) = rsSIDE_DB(0)
'_________________________________________________________________________________________________
xSql = "select count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and substring(Intv_merged_text,1,47) =  'Routed from rp [OFCS_OUT] to rp [AutoRcvPbOFAC]'"


Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then Nb_E(6) = rsSIDE_DB(0)


fgSelect2sf_Display
'_________________________________________________________________________________________________

mGroupBy = "x_inst0_unit_name , mesg_type"
xSql = "select " & mGroupBy & " , count(*) from rMesg  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
          & " and mesg_crea_oper_nickname = 'SYSTEM'" _
          & " and mesg_mod_oper_nickname <> 'SYSTEM'" _
         & " group by  " & mGroupBy _
         & " order by " & mGroupBy
   Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

fgSelect2sf_Display_Suite "Msg SAB Modifiés SAA"

'_________________________________________________________________________________________________

mGroupBy = "x_inst0_unit_name , mesg_type"
xSql = "select " & mGroupBy & " , count(*) from rMesg  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
          & " and mesg_crea_rp_name = '_MP_creation'" _
         & " group by  " & mGroupBy _
         & " order by " & mGroupBy
   Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)


fgSelect2sf_Display_Suite "Msg créés SAA"
'_________________________________________________________________________________________________
mGroupBy = "intv_oper_nickname , mesg_type"
xSql = "select " & mGroupBy & " , count(*) from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and substring(Intv_merged_text,1,55) = 'Routed from rp [_MP_authorisation] to rp [_SI_to_SWIFT]'" _
         & " group by  " & mGroupBy _
         & " order by " & mGroupBy


Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
fgSelect2sf_Display_Suite "Msg autorisés SAA"
'_________________________________________________________________________________________________
fgSelect.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
    Resume Next
End Sub


Private Sub fgDetail_Display()
Dim wColor As Long
Dim xWhere As String, xSql As String

Dim K As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.Row = 0

currentAction = "fgDetail_Display"

fgSelect.Col = arrSWISAB_Group_Nb
If Val(fgSelect.Text) > 200 Then
    If vbNo = MsgBox("Il y a " & Val(fgSelect.Text) & " messages," & vbCrLf & "voulez-vous continuer ?", vbQuestion & vbYesNo, "SWI_Stat : affichage détail") Then Exit Sub
End If

xWhere = ""
mGroupBy = arrSWISAB_Field(arrSWISAB_Group(1))
For K = 1 To arrSWISAB_Group_Nb
    fgSelect.Col = K - 1
    xWhere = xWhere & " and " & arrSWISAB_Field(arrSWISAB_Group(K)) & " ='" & Trim(fgSelect.Text) & "'"
Next K

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0" _
     & " left outer join " & paramIBM_Library_SABSPE & ".YSWISAB1 on SWISAB1ID = SWISABSWID " _
     & xWhere_SQL & xWhere

Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    V = rsYSWISAB0_GetBuffer(rsSab, xYSWISAB0)
    If xYSWISAB0.SWISABWMTK = "103" Or xYSWISAB0.SWISABWMTK = "202" Then
        V = rsYSWISAB1_GetBuffer(rsSab, xYSWISAB1)
    Else
        rsYSWISAB1_Init xYSWISAB1
    End If
    
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_DisplayLine
    
    rsSab.MoveNext
Loop

fgDetail.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgDetail2_Display()
Dim wColor As Long
Dim X0 As String, X1 As String, xSql As String

Dim K As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = "<Type  |> Montant                            |<Devise |> Date                                             |<BIC                         |<N/Réf                                      |<L/Réf                                  |||"
fgDetail.Row = 0

currentAction = "fgDetail_Display"

fgSelect.Col = 2
If Val(fgSelect.Text) > 200 Then
    If vbNo = MsgBox("Il y a " & Val(fgSelect.Text) & " messages," & vbCrLf & "voulez-vous continuer ?", vbQuestion & vbYesNo, "SWI_Stat : affichage détail") Then Exit Sub
End If

fgSelect.Col = 0: X0 = Trim(fgSelect.Text)
fgSelect.Col = 1: X1 = Trim(fgSelect.Text)

xSql = "select * from rMesg , rIntv  " _
          & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
          & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
         & " and mesg_type  =  '" & X1 & "'" _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and Intv_oper_nickname  = '" & X0 & "'" _
        & " and substring(Intv_merged_text,1,55) =  'Routed from rp [_MP_authorisation] to rp [_SI_to_SWIFT]'"
   
   Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
Do While Not rsSIDE_DB.EOF
    
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail2_DisplayLine
    
    rsSIDE_DB.MoveNext
Loop

fgDetail.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


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

Public Sub fgSelect2_DisplayLine()
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim xSql As String, xCur As Currency
On Error Resume Next

fgSelect.Col = 0: fgSelect.Text = rsSIDE_DB(0)
fgSelect.Col = 1: fgSelect.Text = rsSIDE_DB(1)
fgSelect.Col = 2: fgSelect.Text = rsSIDE_DB(2)
End Sub

Public Sub fgSelect3_DisplayLine()
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim xSql As String, xCur As Currency
Dim blnEntrant As Boolean
'On Error Resume Next

fgSelect.Col = 0: fgSelect.Text = Trim(rsSab("MOUVEMCOM"))
fgSelect.Col = 1: fgSelect.Text = dateImp10(rsSab("MOUVEMDTR") + 19000000)
fgSelect.Col = 2: fgSelect.Text = rsSab("MOUVEMSER") & " " & rsSab("MOUVEMSSE") & " " & rsSab("MOUVEMOPE") & " " & rsSab("MOUVEMNUM")
fgSelect.Col = 4: fgSelect.Text = Format$(rsSab("MOUVEMMON"), "###########0.00")
If rsSab("MOUVEMMON") > 0 Then fgSelect.CellForeColor = vbRed
fgSelect.Col = 5: fgSelect.Text = Trim(rsSab("COMPTEDEV"))
fgSelect.Col = 16: fgSelect.Text = Trim(rsSab("LIBELLIB1")) & Trim(rsSab("LIBELLIB2")) & Trim(rsSab("LIBELLIB3")) & Trim(rsSab("LIBELLIB4"))

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 , " & paramIBM_Library_SABSPE & ".YSWISAB1" _
     & " where SWISABWMTK = '103' and SWISABSER = '" & rsSab("MOUVEMSER") & "' and SWISABSSE = '" & rsSab("MOUVEMSSE") _
     & "' and SWISABOPEC = '" & rsSab("MOUVEMOPE") & "' and SWISABOPEN = " & rsSab("MOUVEMNUM") _
     & " and SWISAB1ID = SWISABSWID" _
     & " order by SWISABSWID"
Set rsSabX = cnsab.Execute(xSql)

blnEntrant = False
Do While Not rsSabX.EOF
    'If rsSabX("SWISABWES") = "S" Then
        fgSelect.Col = 9: fgSelect.Text = rsSabX("SWISABW52A")
        fgSelect.Col = 8: fgSelect.Text = rsSabX("SWISABW50P")
        fgSelect.Col = 10: fgSelect.Text = rsSabX("SWISABW57A")
        fgSelect.Col = 11: fgSelect.Text = rsSabX("SWISABW59P")
    If rsSabX("SWISABWES") = "E" Then
        fgSelect.Col = 3: fgSelect.Text = rsSabX("SWISABWL20")
        Call fgSelect3_DisplayLine_103
        blnEntrant = True
    Else
        If Not blnEntrant Then
            fgSelect.Col = 3: fgSelect.Text = rsSabX("SWISABWL20")
           Call fgSelect3_DisplayLine_103
        End If
    End If
    rsSabX.MoveNext
Loop



End Sub


Public Sub fgSelect3_DisplayLine_103()
Dim xSql As String, X As String, X1 As String, K As Integer, K2 As Integer, K3 As Integer
Dim mField As String
Dim wText_Data_Block As String
Dim x50 As String, x70 As String, x59 As String
On Error GoTo Error_Handler
'==================================================================
xSql = "select *  from rtextField  " _
& "where Aid = " & rsSabX("SWISABWID1") _
& " and text_s_umidl = " & rsSabX("SWISABWIDL") _
& " and text_s_umidh  =  " & rsSabX("SWISABWIDH") _
& " order by field_cnt"
Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then
    Do While Not rsSIDE_DB.EOF
        mField = rsSIDE_DB("field_code") ' & rsSIDE_DB("field_option")
        X = rsSIDE_DB("value") '& Asc13
        Select Case mField
            Case "50": x50 = X
            Case "59": x59 = X
            Case "70": x70 = X

        End Select
        rsSIDE_DB.MoveNext
    
    Loop
Else
    xSql = "select * from rtext " _
        & "where Aid = " & rsSabX("SWISABWID1") _
        & " and text_s_umidl = " & rsSabX("SWISABWIDL") _
        & " and text_s_umidh  =  " & rsSabX("SWISABWIDH")
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    If Not rsSIDE_DB.EOF Then
        V = rsSIDE_DB("text_data_block"): wText_Data_Block = IIf(IsNull(V), "", V & vbCrLf & ":")
    '_____________________________________________________
         K = InStr(wText_Data_Block, ":50") + 3
        If K > 0 Then
            K2 = InStr(K, wText_Data_Block, ":") + 1
            K3 = InStr(K2, wText_Data_Block, ":") + 1
            x50 = Mid$(wText_Data_Block, K2, K3 - K2 + 1)
        End If
         K = InStr(wText_Data_Block, ":59") + 3
        If K > 0 Then
            K2 = InStr(K, wText_Data_Block, ":") + 1
            K3 = InStr(K2, wText_Data_Block, ":") + 1
            x59 = Mid$(wText_Data_Block, K2, K3 - K2 + 1)
        End If
         K = InStr(wText_Data_Block, ":70") + 3
        If K > 0 Then
            K2 = InStr(K, wText_Data_Block, ":") + 1
            K3 = InStr(K2, wText_Data_Block, ":") + 1
            x70 = Mid$(wText_Data_Block, K2, K3 - K2 + 1)
        End If
    End If
End If
'_____________________________________________________
If Mid$(x50, 1, 1) = "/" Then
    K = InStr(x50, vbCrLf)
    If K > 0 Then
        fgSelect.Col = 6: fgSelect.Text = Mid$(x50, 1, K - 1)
        fgSelect.Col = 7: fgSelect.Text = Replace(Mid$(x50, K + 2, Len(x50) - K - 1), vbCrLf, " ")
    Else
        fgSelect.Col = 7: fgSelect.Text = Replace(x50, vbCrLf, " ")
    End If
Else
    fgSelect.Col = 7: fgSelect.Text = Replace(x50, vbCrLf, " ")
End If

If Mid$(x59, 1, 1) = "/" Then
    K = InStr(x59, vbCrLf)
    If K > 0 Then
        fgSelect.Col = 12: fgSelect.Text = Replace(Mid$(x59, 1, K - 1), vbCrLf, " ")
        fgSelect.Col = 13: fgSelect.Text = Replace(Mid$(x59, K + 2, Len(x59) - K - 1), vbCrLf, " ")
    Else
        fgSelect.Col = 13: fgSelect.Text = Replace(x59, vbCrLf, " ")
    End If
Else
    fgSelect.Col = 13: fgSelect.Text = Replace(x59, vbCrLf, " ")
End If

K = InStr(x70, "P/O")
If K > 0 Then
    fgSelect.Col = 14: fgSelect.Text = Replace(Mid$(x70, K, Len(x70) - K - 1), vbCrLf, " ")
    fgSelect.Col = 15: fgSelect.Text = Replace(Mid$(x70, 1, K - 1), vbCrLf, " ")
Else
    fgSelect.Col = 15: fgSelect.Text = Replace(x70, vbCrLf, " ")
End If

GoTo Exit_sub

'==================================================================
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation_SAB_YSWISAB1"
Exit_sub:

End Sub

Public Sub fgDetail_DisplayLine()
Dim wColor As Long

If xYSWISAB0.SWISABWES = "S" Then
    wColor = RGB(16, 96, 16)
Else
    wColor = vbBlue
End If

On Error Resume Next
'wColor = vbBlue: wColor_Row = vbWhite
fgDetail.Col = 0: fgDetail.Text = xYSWISAB0.SWISABSER & " " & xYSWISAB0.SWISABSSE
fgDetail.CellForeColor = wColor
fgDetail.Col = 1: fgDetail.Text = xYSWISAB0.SWISABOPEC & " " & xYSWISAB0.SWISABOPEN
fgDetail.CellForeColor = wColor
fgDetail.Col = 2: fgDetail.Text = xYSWISAB0.SWISABWMTK
fgDetail.CellForeColor = wColor

fgDetail.Col = 3:
Select Case xYSWISAB1.SWISABWEBA
    Case "E": fgDetail.Text = "EBA"
    Case "T": fgDetail.Text = "TGT"
End Select
fgDetail.CellForeColor = wColor
If xYSWISAB0.SWISABWMTD <> 0 Then
    fgDetail.Col = 4: fgDetail.Text = Format(xYSWISAB0.SWISABWMTD, "### ### ### ##0.00")
    fgDetail.CellForeColor = wColor
End If
fgDetail.Col = 5: fgDetail.Text = xYSWISAB0.SWISABWDEV
fgDetail.CellForeColor = wColor
fgDetail.Col = 6: fgDetail.Text = dateImp10_S(xYSWISAB0.SWISABWAMJ)
fgDetail.CellForeColor = wColor
fgDetail.Col = 7: fgDetail.Text = xYSWISAB0.SWISABWBIC
fgDetail.CellForeColor = wColor
fgDetail.Col = 8: fgDetail.Text = xYSWISAB1.SWISABW52A
fgDetail.CellForeColor = wColor
fgDetail.Col = 9: fgDetail.Text = xYSWISAB1.SWISABW50P
fgDetail.CellForeColor = wColor
fgDetail.Col = 10: fgDetail.Text = xYSWISAB1.SWISABW57A
fgDetail.CellForeColor = wColor
fgDetail.Col = 11: fgDetail.Text = xYSWISAB1.SWISABW59P
fgDetail.CellForeColor = wColor
fgDetail.Col = 12: fgDetail.Text = xYSWISAB0.SWISABWN20
fgDetail.CellForeColor = wColor
fgDetail.Col = 13: fgDetail.Text = xYSWISAB0.SWISABWL20
fgDetail.CellForeColor = wColor
fgDetail.Col = 14: fgDetail.Text = xYSWISAB0.SWISABSWID
fgDetail.CellForeColor = wColor


'fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex
End Sub

Public Sub fgDetail2_DisplayLine()
Dim wColor As Long, curX As Currency

If rsSIDE_DB("mesg_sub_format") = "INPUT" Then
    wColor = RGB(16, 96, 16)
Else
    wColor = vbBlue
End If

On Error Resume Next
fgDetail.Col = 0: fgDetail.Text = rsSIDE_DB("mesg_type")
fgDetail.CellForeColor = wColor

curX = CCur(rsSIDE_DB("x_fin_amount"))
If curX <> 0 Then
    fgDetail.Col = 1: fgDetail.Text = Format(curX, "### ### ### ##0.00")
    fgDetail.CellForeColor = wColor
End If
fgDetail.Col = 2: fgDetail.Text = rsSIDE_DB("x_fin_ccy")
fgDetail.CellForeColor = wColor
fgDetail.Col = 3: fgDetail.Text = rsSIDE_DB("last_update")
fgDetail.CellForeColor = wColor
fgDetail.Col = 4: fgDetail.Text = Mid$(rsSIDE_DB("mesg_uumid"), 2, 11)
fgDetail.CellForeColor = wColor
fgDetail.Col = 5: fgDetail.Text = rsSIDE_DB("mesg_trn_ref")
fgDetail.CellForeColor = wColor
fgDetail.Col = 6: fgDetail.Text = rsSIDE_DB("mesg_rel_trn_ref")
fgDetail.CellForeColor = wColor
fgDetail.Col = 7: fgDetail.Text = rsSIDE_DB("rMesg.Aid")
fgDetail.CellForeColor = wColor
fgDetail.Col = 8: fgDetail.Text = rsSIDE_DB("Mesg_s_umidl")
fgDetail.CellForeColor = wColor
fgDetail.Col = 9: fgDetail.Text = rsSIDE_DB("Mesg_s_umidh")
fgDetail.CellForeColor = wColor


'fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex
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



Public Sub fgDetail_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgDetail.Rows - 1
    fgDetail.Row = I
    fgDetail.Col = lK
    Select Case lK
        Case 4:
            fgDetail.Col = 5: X = Trim(fgDetail.Text)
            fgDetail.Col = 4: X = Format$(Val(fgDetail.Text), "000000000000000.00") & X
        Case 5:
            fgDetail.Col = 4: X = Format$(Val(fgDetail.Text), "000000000000000.00")
            fgDetail.Col = 5: X = Trim(fgDetail.Text) & X
        Case 6:
            fgDetail.Col = 6: X = fgDetail.Text: X = Mid$(X, 7, 4) & Mid$(X, 4, 2) & Mid$(X, 1, 2)
        Case 7, 8, 9, 10, 11, 12, 13:
            fgDetail.Col = 6: X = fgDetail.Text: X = Mid$(X, 7, 4) & Mid$(X, 4, 2) & Mid$(X, 1, 2)
            fgDetail.Col = lK: X = Trim(fgDetail.Text) & X
        Case 14:
            fgDetail.Col = 14: X = Format$(Val(fgDetail.Text), "000000000000000")
    End Select
    fgDetail.Col = fgDetail_arrIndex - 1
    fgDetail.Text = X
Next I

fgDetail_Sort1 = fgDetail_arrIndex - 1: fgDetail_Sort2 = fgDetail_arrIndex - 1
fgdetail_Sort
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
Call BIA_VB_HAB(wFct, arrHab(), cboSelect_SQL)

'blnSetfocus = True
Form_Init


blnAuto = False


End Sub


Public Sub Form_Init()
Dim V, xSql As String, X As String
Dim K As Long

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True

blnControl = False

cmdReset


fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False
fgSelect_BackColorFixed = fgSelect.BackColorFixed
fgSelect_ForeColorFixed = fgSelect.ForeColorFixed
fgSelect_ForeColor = fgSelect.ForeColor
fgSelect_BackColor = fgSelect.BackColor


fraSelect_Options_1A.BorderStyle = 0

fraSelect_Options_2.Visible = False
Set fraSelect_Options_2.Container = fraTab0
fraSelect_Options_2.Top = fraSelect_Options_1.Top
fraSelect_Options_2.Left = fraSelect_Options_1.Left
fraSelect_Options_2.Height = fraSelect_Options_1.Height
fraSelect_Options_2.Width = fraSelect_Options_1.Width


fraSelect_Options_3.Visible = False
Set fraSelect_Options_3.Container = fraTab0
fraSelect_Options_3.Top = fraSelect_Options_1.Top
fraSelect_Options_3.Left = fraSelect_Options_1.Left
fraSelect_Options_3.Height = fraSelect_Options_1.Height
fraSelect_Options_3.Width = fraSelect_Options_1.Width

lstW.Visible = False
Set lstW.Container = fraTab0
lstW.Top = fgDetail.Top + 300
lstW.Left = fgDetail.Left + fgDetail.Width - lstW.Width - 300
lstW.Height = fgDetail.Height - 300
lstW.BackColor = &HFAFAFA
lstW.ForeColor = vbBlack '&H4080&

fgDetail_FormatString = fgDetail.FormatString
fgDetail.Left = 1035
fgDetail.Top = 2700

Set fraSwift.Container = fraTab0
fraSwift.Top = fgSelect.Top
fraSwift.Left = fraTab0.Left + fraTab0.Width - fraSwift.Width - 200
fgSwift_FormatString = fgSwift.FormatString

wAmjMin = Mid$(YBIATAB0_DATE_CPT_J, 1, 4) & "0101"
wAmjMax = YBIATAB0_DATE_CPT_J
Call DTPicker_Set(txtSelect_SWISABWAMJ_Max, wAmjMax) '
Call DTPicker_Set(txtSelect_SWISABWAMJ_Min, wAmjMin) '
wAmjMin = Mid$(YBIATAB0_DATE_CPT_J, 1, 6) & "01"
Call DTPicker_Set(txtSelect3_SWISABWAMJ_Min, wAmjMax) '
Call DTPicker_Set(txtSelect3_SWISABWAMJ_Max, wAmjMax) '
wAmjMin = Mid$(YBIATAB0_DATE_CPT_J, 1, 6) & "01"
Call DTPicker_Set(txtSelect3_SWISABWAMJ_Min, wAmjMax) '
wAmjMax = YBIATAB0_DATE_CPT_MP1
Call DTPicker_Set(txtSelect2_SWISABWAMJ_Max, wAmjMax) '
wAmjMin = Mid$(YBIATAB0_DATE_CPT_MP1, 1, 6) & "01"
Call DTPicker_Set(txtSelect2_SWISABWAMJ_Min, wAmjMin) '

lstW.Clear

param_Init

If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0


Me.Enabled = True
End Sub

Private Sub cboSelect_SWISAB_1_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISAB_2_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISAB_3_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISAB_4_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISAB_5_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABOPEC_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABSSE_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABW50P_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABW50P_K_Click()
If blnControl Then cmdSelect_Clear
If Trim(cboSelect_SWISABW50P_K) = "" Then cboSelect_SWISABW50P.ListIndex = 0

End Sub


Private Sub cboSelect_SWISABW50Z_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABW52A_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABW52A_K_Click()
If blnControl Then cmdSelect_Clear
If Trim(cboSelect_SWISABW52A_K) = "" Then cboSelect_SWISABW52A.ListIndex = 0

End Sub


Private Sub cboSelect_SWISABW57A_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABW57A_K_Click()
If blnControl Then cmdSelect_Clear
If Trim(cboSelect_SWISABW57A_K) = "" Then cboSelect_SWISABW57A.ListIndex = 0

End Sub


Private Sub cboSelect_SWISABW59P_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABW59P_K_Click()
If blnControl Then cmdSelect_Clear
If Trim(cboSelect_SWISABW59P_K) = "" Then cboSelect_SWISABW59P.ListIndex = 0

End Sub


Private Sub cboSelect_SWISABW59Z_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABW59Z_K_Click()
If blnControl Then cmdSelect_Clear
If Trim(cboSelect_SWISABW59Z_K) = "" Then cboSelect_SWISABW59Z.ListIndex = 0

End Sub

Private Sub cboSelect_SWISABWBIC_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABWBIC_Pays_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABWDEV_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABWDEV_K_Click()
If blnControl Then cmdSelect_Clear
If Trim(cboSelect_SWISABWDEV_K) = "" Then cboSelect_SWISABWDEV.ListIndex = 0

End Sub


Private Sub cboSelect_SWISABW50Z_K_Click()
If blnControl Then cmdSelect_Clear
If Trim(cboSelect_SWISABW50Z_K) = "" Then cboSelect_SWISABW50Z.ListIndex = 0

End Sub

Private Sub cboSelect_SWISABOPEC_K_Click()
If blnControl Then cmdSelect_Clear

If Trim(cboSelect_SWISABOPEC_K) = "" Then cboSelect_SWISABOPEC.ListIndex = 0

End Sub

Private Sub cboSelect_SWISABWBIC_K_Click()
If blnControl Then cmdSelect_Clear
If Trim(cboSelect_SWISABWBIC_K) = "" Then cboSelect_SWISABWBIC.ListIndex = 0

End Sub


Private Sub cboSelect_SWISABWEBA_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABWEBA_K_Click()
If blnControl Then cmdSelect_Clear
If Trim(cboSelect_SWISABWEBA_K) = "" Then cboSelect_SWISABWEBA.ListIndex = 0

End Sub


Private Sub cboSelect_SWISABSSE_K_Click()
If Trim(cboSelect_SWISABSSE_K) = "" Then cboSelect_SWISABSSE.ListIndex = 0
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABWES_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_SWISABWMTK_K_Click()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub chkSAB_Dossier_DB_Show_Click()
On Error Resume Next
Dim K As Integer
If fraSwift.Visible = True Then
    If chkSAB_Dossier_DB_Show = "1" Then
        If mMOUVEMNUM > 0 Then Call frmSAB_Dossier_DB.Form_Init("", "", "", "", mMOUVEMSER, mMOUVEMSSE, mMOUVEMOPE, mMOUVEMNUM)
    Else
        frmSAB_Dossier_DB.Hide
    End If
End If

End Sub

Private Sub chkSIDE_DB_Show_Click()
On Error Resume Next
Dim K As Integer
If fraSwift.Visible = True Then
    If chkSIDE_DB_Show = "1" Then
        K = InStr(libSWIFT_SWISABSWID, " ")
        'If K > 0 Then frmSIDE_DB.fgSwift_Display Val(Mid$(libSWIFT_SWISABSWID, 1, K))
        If K > 0 Then Call frmSIDE_DB.fgSwift_Display(mSWISABSWID, Mesg_aid, mesg_s_umidl, mesg_s_umidh)
    Else
        frmSIDE_DB.Hide
    End If
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If chkSIDE_DB_Show = "1" Then frmSIDE_DB.Hide
If chkSAB_Dossier_DB_Show = "1" Then frmSAB_Dossier_DB.Hide
    cnSIDE_DB.Close
    Set cnSIDE_DB = Nothing

End Sub

Private Sub mnuExportation_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Exportation en cours ......"): DoEvents

'cmdSelect_SQL_1

YSWISAB0_Export


Me.Enabled = True: Me.MousePointer = 0
End Sub

Public Sub YSWISAB0_Export()
On Error GoTo Error_Handler
Dim Nb As Long, wId As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSql As String
Dim wAmjMin As String, wAmjMax As String
Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim X As String, K As Long, K2 As Long, kMax As Long, K_Nb As Long, K_Mt As Long
Dim xWhere As String, X2 As String
Dim wForecolor As Long, wBackColor As Long
Dim s_Solde As Currency, s_Prov As Currency
Dim t_Solde As Currency, t_Prov As Currency, x_Prov As Currency
Dim mSWISABDOS As Long, mSWISABWAMJ As String
'______________________________________________
Call DTPicker_Control(txtSelect_SWISABWAMJ_Min, wAmjMin)
Call DTPicker_Control(txtSelect_SWISABWAMJ_Max, wAmjMax)

wFile = "C:\Temp\YSWISAB0 " & DSys & " " & time_Hms & ".xlsx"
'______________________________________________

X = InputBox("par défaut : " & wFile _
    & vbCrLf & vbCrLf & "     =========================" _
    & vbCrLf & "     =========================", "SWI_STAT : nom du fichier d'exportation", wFile)
If Trim(X) = "" Then Exit Sub

wFilex = Trim(X)
'______________________________________________


If Dir(wFile) <> "" Then Kill wFile

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "SWI_STAT"
    .Subject = "SWI_STAT"
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "SWI_STAT"
'__________________________________________________________________________________

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .WrapText = True
    .Font.Size = 9
    .Font.Name = "Calibri"
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 100

wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
Nb = 1
Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : " & Nb & " enregistrements"): DoEvents

Select Case cmdSelect_SQL_K
    Case "1"
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Messages SWIFT : Statistiques du " & dateImp10(wAmjMin) & " au " & dateImp10(wAmjMax)

'____________________________________________________________________________________________
        For K = 1 To arrSWISAB_Group_Nb
            wsExcel.Cells(Nb, K) = arrSWISAB_Lib(arrSWISAB_Group(K)): wsExcel.Columns(K).ColumnWidth = 15
            wsExcel.Columns(K).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        Next K
        
        K_Nb = arrSWISAB_Group_Nb + 1
        wsExcel.Cells(Nb, K_Nb) = "Nombre": wsExcel.Columns(K_Nb).ColumnWidth = 10: wsExcel.Columns(K_Nb).NumberFormat = "#######"
        wsExcel.Columns(K_Nb).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        
        K_Mt = arrSWISAB_Group_Nb + 2
        wsExcel.Cells(Nb, K_Mt) = "Montant"
        wsExcel.Columns(K_Mt).ColumnWidth = 20: wsExcel.Columns(K_Mt).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
        wsExcel.Columns(K_Mt).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        
        
        For K = 1 To K_Mt
            wsExcel.Columns(K).VerticalAlignment = Excel.XlVAlign.xlVAlignTop
            wsExcel.Cells(1, K).Interior.Color = RGB(255, 255, 153)
        Next K
        
        
        For K = 1 To fgSelect.Rows - 1
            fgSelect.Row = K
            Nb = Nb + 1
            For K2 = 1 To arrSWISAB_Group_Nb
                fgSelect.Col = K2 - 1
                wsExcel.Cells(Nb, K2) = fgSelect.Text
            Next K2
            fgSelect.Col = K_Nb - 1
            wsExcel.Cells(Nb, K_Nb) = Val(fgSelect.Text)
            fgSelect.Col = K_Mt - 1
            wsExcel.Cells(Nb, K_Mt) = CCur(num_CDec(fgSelect.Text))
        
        Next K
        
    Case Else
'____________________________________________________________________________________________
    Nb = 0
        wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Messages SWIFT : Statistiques du " & xAmj8_from_crea_date_time & " au " & xAmj8_to_crea_date_time


    For K = 0 To fgSelect.Rows - 1
        fgSelect.Row = K
        
        wForecolor = fgSelect.CellForeColor
        If wForecolor = 0 Then
            If K = 0 Then
                wForecolor = fgSelect_ForeColorFixed
            Else
                wForecolor = fgSelect_ForeColor
            End If
        End If
        
        wBackColor = fgSelect.CellBackColor
        If wBackColor = 0 Then
            If K = 0 Then
                wBackColor = fgSelect_BackColorFixed
            Else
                wBackColor = fgSelect_BackColor
            End If
        End If

        Nb = Nb + 1
        For K2 = 0 To 11
        
            fgSelect.Col = K2: X = Trim(fgSelect.Text)
            If K = 0 Then
                wsExcel.Columns(K2 + 1).ColumnWidth = fgSelect.CellWidth / 100
                If K2 > 0 Then wsExcel.Columns(K2 + 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            End If
            wsExcel.Cells(Nb, K2 + 1) = X
            wsExcel.Cells(Nb, K2 + 1).Font.Color = wForecolor
            wsExcel.Cells(Nb, K2 + 1).Interior.Color = wBackColor
        Next K2
    Next K
End Select
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

Public Sub fgDetail_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgDetail.Visible = False
mRow = fgDetail.Row

If lRow > 0 And lRow < fgDetail.Rows Then
    fgDetail.Row = lRow
    For I = fgDetail_arrIndex To fgDetail.FixedCols Step -1
        fgDetail.Col = I: fgDetail.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgDetail.Row = mRow
    If fgDetail.Row > 0 Then
        lRow = fgDetail.Row
        lColor_Old = fgDetail.CellBackColor
        For I = fgDetail_arrIndex To fgDetail.FixedCols Step -1
          fgDetail.Col = I: fgDetail.CellBackColor = lColor
        Next I
    End If
End If
fgDetail.LeftCol = fgDetail.FixedCols
fgDetail.Visible = True
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





Private Sub cboSelect_SWISABWMTK_Click()
If blnControl Then cmdSelect_Clear
End Sub

Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

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
            Case "3trf":
                Me.PopupMenu mnuPrint2, vbPopupMenuLeftButton
            Case Else
                Me.PopupMenu mnuPrint, vbPopupMenuLeftButton
        End Select
    End Select

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdPrint_YSWISAB0(blnDetail As Boolean)
Dim X As String, xSql As String, I As Integer, K As Integer
Dim wAmj As String, xWhere As String
Dim soldeD As typeYSWISAB0, soldeF As typeYSWISAB0, Total As typeYSWISAB0
Dim blnXprt_Line As Boolean
Dim Nb_Detail As Long


End Sub

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SWI_STAT début du traitement........"): DoEvents

If fgSelect.Visible Then cmdSelect_Reset
fgSelect.Visible = False

Select Case cmdSelect_SQL_K
    Case "1":  cmdSelect_SQL_1
    Case "2":  cmdSelect_SQL_2
    Case "2sf":  cmdSelect_SQL_2sf
    Case "3trf":  cmdSelect_SQL_3trf
    Case "JPL":  cmdSelect_SQL_JPL
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< SWI_STAT traitement terminé"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus

End Sub


Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String
On Error Resume Next

fraSwift.Visible = False

If y <= fgDetail.RowHeightMin Then
    Select Case fgDetail.Col
        Case 0: fgDetail_Sort1 = 0: fgDetail_Sort2 = 2: fgdetail_Sort
        Case 1:  fgDetail_Sort1 = 1: fgDetail_Sort2 = 2: fgdetail_Sort
        Case 2: fgDetail_Sort1 = 2: fgDetail_Sort2 = 2: fgdetail_Sort
        Case 3: fgDetail_Sort1 = 3: fgDetail_Sort2 = 3: fgdetail_Sort
        Case 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14: fgDetail_SortX fgDetail.Col
    End Select
Else
    If fgDetail.Rows > 1 Then
       ' blnControl = False
       Select Case cmdSelect_SQL_K
            Case "1"
                Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
                fgDetail.Col = 14 ':  arrYSWISAB0_Index = CLng(fgDetail.Text)
                fgSwift_Display CLng(fgDetail.Text)
            Case "2"
                fgDetail.Col = 7: Mesg_aid = CLng(fgDetail.Text)
                fgDetail.Col = 8: mesg_s_umidl = CLng(fgDetail.Text)
                fgDetail.Col = 9: mesg_s_umidh = CLng(fgDetail.Text)
                fgSwift_Display 0
        End Select
   End If
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
    Case Is = 27: cmdContext_Quit: KeyCode = 0
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select


End Sub

Public Sub cmdContext_Quit()
'blnControl = False
lstErr.Clear: lstErr.Height = 200

If SSTab1.Tab <> 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If
If fraSwift.Visible Then fraSwift.Visible = False:       Exit Sub

If lstW.Visible Then lstW.Visible = False:       Exit Sub

If fgDetail.Visible Then
    fgDetail.Visible = False
    Exit Sub
End If
If fgSelect.Visible Then
    fgSelect.Visible = False
    Exit Sub
End If

If SSTab1.Tab = 0 Then
    Unload Me
End If

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
cnSIDE_DB.Open paramODBC_DSN_SIDE_DB
fgSelect.Clear: fgSelect.Row = 0
End Sub





Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String, xSql As String
On Error Resume Next

fgDetail.Visible = False
fraSwift.Visible = False
lstW.Visible = False
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  arrYSWISAB0_Index = CLng(fgSelect.Text)
        
        Select Case cmdSelect_SQL_K
            Case "1": fgDetail_Display
            Case "2": fgDetail2_Display
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






Private Sub mnuPrint2_Excel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String
Call lstErr_AddItem(lstErr, cmdContext, "< SWI_STAT Export en cours "): DoEvents

Select Case cmdSelect_SQL_K
    Case "3trf":
        X = "Sélection du " & dateImp10_S(DSys)
        Call MSflexGrid_Excel("", "SWI_Stat", X, fgSelect, fgSelect.Cols - 1)
End Select
Call lstErr_AddItem(lstErr, cmdContext, "< SWI_STAT Export terminé"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint2_Mail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String

Call lstErr_AddItem(lstErr, cmdContext, "< SWI_STAT Export en cours "): DoEvents

Select Case cmdSelect_SQL_K
    Case "3trf":
        X = "Sélection du " & dateImp10_S(DSys)
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "SWI_Stat", X, X, fgSelect, fgSelect.Cols - 1)
End Select


Call lstErr_AddItem(lstErr, cmdContext, "< SWI_STAT Export terminé"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub


'Public Sub fgSelect_ForeColor(lColor As Long)
'For I = 0 To fgSelect_arrIndex
'  fgSelect.Col = I: fgSelect.CellForeColor = lColor
'Next I

'End Sub
















Private Sub txtSelect_SWISABWAMJ_Max_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub txtSelect_SWISABWAMJ_Max_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub txtSelect_SWISABWAMJ_Max_KeyPress(KeyAscii As Integer)
cmdSelect_Reset

End Sub

Private Sub txtSelect_SWISABWAMJ_Min_Change()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub txtSelect_SWISABWAMJ_Min_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub txtSelect_SWISABWAMJ_Min_KeyPress(KeyAscii As Integer)
cmdSelect_Reset

End Sub

Public Sub param_Init()
Dim xSql As String
Dim K As Integer

arrSWISAB_Field(0) = "": arrSWISAB_Lib(1) = "": arrSWISAB1(0) = False
arrSWISAB_Field(1) = "SWISABOPEC": arrSWISAB_Lib(1) = "Code Opé": arrSWISAB1(1) = False
arrSWISAB_Field(2) = "SWISABWMTK": arrSWISAB_Lib(2) = "Type MT": arrSWISAB1(2) = False
arrSWISAB_Field(3) = "SWISABWES": arrSWISAB_Lib(3) = "E / S": arrSWISAB1(3) = False
arrSWISAB_Field(4) = "SWISABWDEV": arrSWISAB_Lib(4) = "Devise": arrSWISAB1(4) = False
arrSWISAB_Field(5) = "SWISABSSE": arrSWISAB_Lib(5) = "Service": arrSWISAB1(5) = False
arrSWISAB_Field(6) = "SWISABWBIC": arrSWISAB_Lib(6) = "BIC d'échange": arrSWISAB1(6) = False
arrSWISAB_Field(7) = "SWISABW52A": arrSWISAB_Lib(7) = "BIC BQ  D.O.": arrSWISAB1(7) = True
arrSWISAB_Field(8) = "SWISABW57A": arrSWISAB_Lib(8) = "BIC BQ  BEN": arrSWISAB1(8) = True
arrSWISAB_Field(9) = "SWISABWEBA": arrSWISAB_Lib(9) = "Route": arrSWISAB1(9) = True
arrSWISAB_Field(10) = "SWISABW50P": arrSWISAB_Lib(10) = "Pays DO": arrSWISAB1(10) = True
arrSWISAB_Field(11) = "SWISABW59P": arrSWISAB_Lib(11) = "Pays BEN": arrSWISAB1(11) = True
arrSWISAB_Field(12) = "SWISABW50Z": arrSWISAB_Lib(12) = "Zone Pays DO": arrSWISAB1(12) = True
arrSWISAB_Field(13) = "SWISABW59Z": arrSWISAB_Lib(13) = "Zone Pays BEN": arrSWISAB1(13) = True
arrSWISAB_Field_Nb = 13

For K = 0 To arrSWISAB_Field_Nb
    If K > 0 Then cboSelect_SWISAB_1.AddItem arrSWISAB_Lib(K)
    cboSelect_SWISAB_2.AddItem arrSWISAB_Lib(K)
    cboSelect_SWISAB_3.AddItem arrSWISAB_Lib(K)
    cboSelect_SWISAB_4.AddItem arrSWISAB_Lib(K)
    cboSelect_SWISAB_5.AddItem arrSWISAB_Lib(K)
Next K
cboSelect_SWISAB_1.ListIndex = 0
arrSWISAB_Group(1) = 0
arrSWISAB_Group_Nb = 0

'Initialisation WES_______________________________________________________________________________
cboSelect_SWISABWES.Clear
cboSelect_SWISABWES.AddItem ""
cboSelect_SWISABWES.AddItem "Entrant"
cboSelect_SWISABWES.AddItem "Sortant"
cboSelect_SWISABWES.ListIndex = 0

'Initialisation Pays_______________________________________________________________________________
cboSelect_SWISABWBIC_Pays.Clear
cboSelect_SWISABWBIC_Pays.AddItem ""

 xSql = "select * from CT " _
     & "order by  CountryCode"

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
   
Do While Not rsSIDE_DB.EOF
    cboSelect_SWISABWBIC_Pays.AddItem rsSIDE_DB(0) & " " & LCase(Trim(rsSIDE_DB(1)))
    rsSIDE_DB.MoveNext

Loop


'

'_______________________________________________________________________________
cboSelect_SWISABOPEC_K.Clear
cboSelect_SWISABOPEC_K.AddItem ""
cboSelect_SWISABOPEC_K.AddItem "="
cboSelect_SWISABOPEC_K.AddItem "<>"

cboSelect_SWISABOPEC.Clear
cboSelect_SWISABOPEC.AddItem ""
xSql = "select distinct SWISABOPEC from " & paramIBM_Library_SABSPE & ".YSWISAB0 order by SWISABOPEC"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWISABOPEC.AddItem Trim(rsSab("SWISABOPEC"))
    rsSab.MoveNext
Loop

'_______________________________________________________________________________
'Initialisation MTK_______________________________________________________________________________

cboSelect_SWISABWMTK_K.Clear
cboSelect_SWISABWMTK_K.AddItem ""
cboSelect_SWISABWMTK_K.AddItem "="
cboSelect_SWISABWMTK_K.AddItem "<>"

cboSelect_SWISABWMTK.Clear
cboSelect_SWISABWMTK.AddItem ""
cboSelect_SWISABWMTK.AddItem "103"
cboSelect_SWISABWMTK.AddItem "103,202"
cboSelect_SWISABWMTK.AddItem "202"
cboSelect_SWISABWMTK.AddItem "700"
cboSelect_SWISABWMTK.AddItem "700,701"
cboSelect_SWISABWMTK.AddItem "%99"
cboSelect_SWISABWMTK.ListIndex = 0
'_______________________________________________________________________________
cboSelect_SWISABWDEV_K.Clear
cboSelect_SWISABWDEV_K.AddItem ""
cboSelect_SWISABWDEV_K.AddItem "="
cboSelect_SWISABWDEV_K.AddItem "<>"

cboSelect_SWISABWDEV.Clear
cboSelect_SWISABWDEV.AddItem ""
xSql = "select distinct SWISABWDEV from " & paramIBM_Library_SABSPE & ".YSWISAB0 order by SWISABWDEV"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWISABWDEV.AddItem Trim(rsSab("SWISABWDEV"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWISABSSE_K.Clear
cboSelect_SWISABSSE_K.AddItem ""
cboSelect_SWISABSSE_K.AddItem "="
cboSelect_SWISABSSE_K.AddItem "<>"

cboSelect_SWISABSSE.Clear
cboSelect_SWISABSSE.AddItem ""
xSql = "select distinct SWISABSSE from " & paramIBM_Library_SABSPE & ".YSWISAB0 order by SWISABSSE"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWISABSSE.AddItem Trim(rsSab("SWISABSSE"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWISABWBIC_K.Clear
cboSelect_SWISABWBIC_K.AddItem ""
cboSelect_SWISABWBIC_K.AddItem "="
cboSelect_SWISABWBIC_K.AddItem "=4"
cboSelect_SWISABWBIC_K.AddItem "=6"
cboSelect_SWISABWBIC_K.AddItem "=8"
cboSelect_SWISABWBIC_K.AddItem "<>"

cboSelect_SWISABWBIC.Clear
cboSelect_SWISABWBIC.AddItem ""
xSql = "select distinct SWISABWBIC from " & paramIBM_Library_SABSPE & ".YSWISAB0 order by SWISABWBIC"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWISABWBIC.AddItem Trim(rsSab("SWISABWBIC"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWISABW52A_K.Clear
cboSelect_SWISABW52A_K.AddItem ""
cboSelect_SWISABW52A_K.AddItem "="
cboSelect_SWISABW52A_K.AddItem "=4"
cboSelect_SWISABW52A_K.AddItem "=6"
cboSelect_SWISABW52A_K.AddItem "=8"
cboSelect_SWISABW52A_K.AddItem "<>"

cboSelect_SWISABW52A.Clear
cboSelect_SWISABW52A.AddItem ""
xSql = "select distinct SWISABW52A from " & paramIBM_Library_SABSPE & ".YSWISAB1 order by SWISABW52A"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWISABW52A.AddItem Trim(rsSab("SWISABW52A"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWISABW57A_K.Clear
cboSelect_SWISABW57A_K.AddItem ""
cboSelect_SWISABW57A_K.AddItem "="
cboSelect_SWISABW57A_K.AddItem "=4"
cboSelect_SWISABW57A_K.AddItem "=6"
cboSelect_SWISABW57A_K.AddItem "=8"
cboSelect_SWISABW57A_K.AddItem "<>"

cboSelect_SWISABW57A.Clear
cboSelect_SWISABW57A.AddItem ""
xSql = "select distinct SWISABW57A from " & paramIBM_Library_SABSPE & ".YSWISAB1 order by SWISABW57A"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWISABW57A.AddItem Trim(rsSab("SWISABW57A"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWISABWEBA_K.Clear
cboSelect_SWISABWEBA_K.AddItem ""
cboSelect_SWISABWEBA_K.AddItem "="
cboSelect_SWISABWEBA_K.AddItem "<>"

cboSelect_SWISABWEBA.Clear
cboSelect_SWISABWEBA.AddItem ""
xSql = "select distinct SWISABWEBA from " & paramIBM_Library_SABSPE & ".YSWISAB1 order by SWISABWEBA"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWISABWEBA.AddItem Trim(rsSab("SWISABWEBA"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWISABW50P_K.Clear
cboSelect_SWISABW50P_K.AddItem ""
cboSelect_SWISABW50P_K.AddItem "="
cboSelect_SWISABW50P_K.AddItem "<>"

cboSelect_SWISABW50P.Clear
cboSelect_SWISABW50P.AddItem ""
xSql = "select distinct SWISABW50P from " & paramIBM_Library_SABSPE & ".YSWISAB1 order by SWISABW50P"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWISABW50P.AddItem Trim(rsSab("SWISABW50P"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWISABW59P_K.Clear
cboSelect_SWISABW59P_K.AddItem ""
cboSelect_SWISABW59P_K.AddItem "="
cboSelect_SWISABW59P_K.AddItem "<>"

cboSelect_SWISABW59P.Clear
cboSelect_SWISABW59P.AddItem ""
xSql = "select distinct SWISABW59P from " & paramIBM_Library_SABSPE & ".YSWISAB1 order by SWISABW59P"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWISABW59P.AddItem Trim(rsSab("SWISABW59P"))
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
cboSelect_SWISABW50Z_K.Clear
cboSelect_SWISABW50Z_K.AddItem ""
cboSelect_SWISABW50Z_K.AddItem "="
cboSelect_SWISABW50Z_K.AddItem "<>"

cboSelect_SWISABW50Z.Clear
cboSelect_SWISABW50Z.AddItem ""
cboSelect_SWISABW50Z.AddItem "FR"
cboSelect_SWISABW50Z.AddItem "UE"
cboSelect_SWISABW50Z.AddItem "**"
'_______________________________________________________________________________
cboSelect_SWISABW59Z_K.Clear
cboSelect_SWISABW59Z_K.AddItem ""
cboSelect_SWISABW59Z_K.AddItem "="
cboSelect_SWISABW59Z_K.AddItem "<>"

cboSelect_SWISABW59Z.Clear
cboSelect_SWISABW59Z.AddItem ""
cboSelect_SWISABW59Z.AddItem "FR"
cboSelect_SWISABW59Z.AddItem "UE"
cboSelect_SWISABW59Z.AddItem "**"

End Sub


Public Sub cmdSelect_Clear()
    lstErr.Clear
    lstW.Visible = False
    fgSelect.Visible = False
    fgDetail.Visible = False
    fraSwift.Visible = False

End Sub

Public Sub cmdSelect_SQL_2sf_Init()
Dim K As Integer
For K = 1 To 20
    Nb_E(K) = 0: Nb_S(K) = 0
Next K
End Sub

Private Sub txtSelect2_SWISABWAMJ_Max_Change()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub txtSelect2_SWISABWAMJ_Max_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub txtSelect2_SWISABWAMJ_Max_KeyPress(KeyAscii As Integer)
cmdSelect_Clear

End Sub


Private Sub txtSelect2_SWISABWAMJ_Min_Change()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub txtSelect2_SWISABWAMJ_Min_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub txtSelect2_SWISABWAMJ_Min_KeyPress(KeyAscii As Integer)
cmdSelect_Clear

End Sub


Private Sub txtSelect2_SWISABWMTK_Change()
If blnControl Then cmdSelect_Clear
End Sub

Private Sub txtSelect2_SWISABWMTK_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtSelect2_SWISABWUSR_Change()
If blnControl Then cmdSelect_Clear
End Sub


Private Sub txtSelect2_SWISABWUSR_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


