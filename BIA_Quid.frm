VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBIA_Quid 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA_Quid"
   ClientHeight    =   10305
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BIA_Quid.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10305
   ScaleWidth      =   13530
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
      Left            =   6120
      TabIndex        =   2
      Top             =   0
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9720
      Left            =   0
      TabIndex        =   3
      Top             =   495
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   17145
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Rechercher"
      TabPicture(0)   =   "BIA_Quid.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "BIA_Quid.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtRTF"
      Tab(1).Control(1)=   "txtFg"
      Tab(1).Control(2)=   "fraSelect_Options_ZFICBDF0"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "BIA_Quid.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame fraSelect_Options_ZFICBDF0 
         BackColor       =   &H00F0FFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   -74670
         TabIndex        =   21
         Top             =   705
         Visible         =   0   'False
         Width           =   9375
         Begin VB.TextBox txtSelect_FICBDFCPT 
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
            Left            =   5445
            MaxLength       =   11
            TabIndex        =   39
            Top             =   810
            Width           =   1485
         End
         Begin VB.TextBox txtSelect_FICBDFBVIL 
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
            Left            =   3500
            TabIndex        =   26
            Top             =   800
            Width           =   1020
         End
         Begin VB.TextBox txtSelect_FICBDFBNOM 
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
            Left            =   3500
            TabIndex        =   25
            Top             =   315
            Width           =   1065
         End
         Begin VB.TextBox txtSelect_FICBDFGVIL 
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
            Left            =   8200
            TabIndex        =   31
            Top             =   525
            Width           =   1035
         End
         Begin VB.TextBox txtSelect_FICBDFGUI 
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
            Left            =   5490
            MaxLength       =   5
            TabIndex        =   27
            Top             =   270
            Width           =   930
         End
         Begin VB.TextBox txtSelect_FICBDFGCP 
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
            Left            =   8200
            TabIndex        =   28
            Top             =   120
            Width           =   1050
         End
         Begin VB.TextBox txtSelect_FICBDFBCIB 
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
            Left            =   1400
            TabIndex        =   22
            Top             =   300
            Width           =   800
         End
         Begin VB.TextBox txtSelect_FICBDFGNOM 
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
            Left            =   8200
            TabIndex        =   30
            Top             =   900
            Width           =   1065
         End
         Begin VB.TextBox txtSelect_FICBDFBCP 
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
            Left            =   1400
            TabIndex        =   24
            Top             =   800
            Width           =   800
         End
         Begin VB.Label libSelect_FICBDFCPT 
            BackColor       =   &H00F0FFFF&
            Caption         =   "compte"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4800
            TabIndex        =   40
            Top             =   825
            Width           =   600
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF00FF&
            BorderWidth     =   5
            X1              =   4650
            X2              =   4650
            Y1              =   120
            Y2              =   1230
         End
         Begin VB.Label Label1 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Ville"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2235
            TabIndex        =   37
            Top             =   800
            Width           =   1125
         End
         Begin VB.Label lblSelect_FICBDFBNOM 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Dénomination"
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
            Left            =   2205
            TabIndex        =   36
            Top             =   300
            Width           =   1200
         End
         Begin VB.Label lblSelect_FICBDFGVIL 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Ville"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   7005
            TabIndex        =   35
            Top             =   585
            Width           =   1125
         End
         Begin VB.Label lblSelect_FICBDFGCP 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Code postal"
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
            Left            =   7005
            TabIndex        =   34
            Top             =   225
            Width           =   1065
         End
         Begin VB.Label lblSelect_FICBDFBCIB 
            BackColor       =   &H00F0FFFF&
            Caption         =   "code banque"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   33
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label lblSelect_FICBDFGNOM 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Domiciliation"
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
            Left            =   7005
            TabIndex        =   32
            Top             =   885
            Width           =   1155
         End
         Begin VB.Label lblSelect_FICBDFBCP 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Code Postal"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   135
            TabIndex        =   29
            Top             =   800
            Width           =   1125
         End
         Begin VB.Label lblSelect_FICBDFGUI 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Guichet"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4800
            TabIndex        =   23
            Top             =   300
            Width           =   600
         End
      End
      Begin VB.TextBox txtFg 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   -69030
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Text            =   "BIA_Quid.frx":035E
         Top             =   1155
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Frame fraSelect 
         BackColor       =   &H00E0E0E0&
         Height          =   9630
         Left            =   -135
         TabIndex        =   4
         Top             =   495
         Width           =   13425
         Begin MSFlexGridLib.MSFlexGrid fgYSAAJRN0 
            Height          =   2805
            Left            =   3840
            TabIndex        =   41
            Top             =   6090
            Visible         =   0   'False
            Width           =   9360
            _ExtentX        =   16510
            _ExtentY        =   4948
            _Version        =   393216
            Cols            =   9
            FixedCols       =   0
            RowHeightMin    =   400
            BackColor       =   12648447
            ForeColor       =   4210752
            BackColorFixed  =   8438015
            ForeColorFixed  =   0
            BackColorBkg    =   12648447
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   $"BIA_Quid.frx":0366
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
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   4600
            Left            =   3810
            TabIndex        =   19
            Top             =   1425
            Visible         =   0   'False
            Width           =   9360
            _ExtentX        =   16510
            _ExtentY        =   8123
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   400
            BackColor       =   15790320
            ForeColor       =   4210752
            BackColorFixed  =   8421504
            ForeColorFixed  =   16777215
            BackColorBkg    =   15790320
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   $"BIA_Quid.frx":043B
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
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   11820
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   705
            Width           =   1335
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
            Height          =   330
            Left            =   9840
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   300
            Width           =   3435
         End
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00F0FFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1305
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Visible         =   0   'False
            Width           =   9375
            Begin VB.CheckBox chkSelect_SWIBKUBIC 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0FFFF&
               Caption         =   "uniquement ..RMA"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   3210
               TabIndex        =   38
               Top             =   945
               Width           =   1860
            End
            Begin VB.CheckBox chkSelect_SWIBICBIC 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0FFFF&
               Caption         =   "uniquement ......XXX"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   75
               TabIndex        =   20
               Top             =   945
               Value           =   1  'Checked
               Width           =   1980
            End
            Begin VB.ComboBox cboSelect_SWIBICPays 
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
               Left            =   3210
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   525
               Width           =   1860
            End
            Begin VB.TextBox txtSelect_SWIBICVIL 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   18
               Top             =   765
               Width           =   1230
            End
            Begin VB.TextBox txtSelect_SWIBICIN1 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6630
               TabIndex        =   17
               Top             =   195
               Width           =   2400
            End
            Begin VB.TextBox txtSelect_SWIBICBIC 
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
               Left            =   870
               TabIndex        =   15
               Top             =   540
               Width           =   1230
            End
            Begin VB.Label lblSelect_SWIBICPPays 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Pays"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   2460
               TabIndex        =   14
               Top             =   540
               Width           =   495
            End
            Begin VB.Label lblSelect_SWIBICVIL 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Ville"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   5640
               TabIndex        =   13
               Top             =   720
               Width           =   495
            End
            Begin VB.Label lblSelect_SWIBICIN1 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Intitulé"
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
               Left            =   5595
               TabIndex        =   12
               Top             =   180
               Width           =   660
            End
            Begin VB.Label lblSelect_SWIBICBIC 
               BackColor       =   &H00F0FFFF&
               Caption         =   "BIC"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   105
               TabIndex        =   11
               Top             =   540
               Width           =   495
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7710
            Left            =   120
            TabIndex        =   10
            Top             =   1425
            Width           =   13140
            _ExtentX        =   23178
            _ExtentY        =   13600
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   400
            BackColor       =   16777215
            ForeColor       =   16711680
            BackColorFixed  =   8421376
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   $"BIA_Quid.frx":0584
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
      Begin RichTextLib.RichTextBox txtRTF 
         Height          =   5610
         Left            =   -69525
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3450
         Visible         =   0   'False
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   9895
         _Version        =   393217
         BackColor       =   15790320
         HideSelection   =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"BIA_Quid.frx":06CA
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
      Picture         =   "BIA_Quid.frx":074A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "mnuPrint"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuRIB 
      Caption         =   "mnuRIB"
      Visible         =   0   'False
      Begin VB.Menu mnuRIB_Clé 
         Caption         =   "RIB_Clé"
      End
      Begin VB.Menu mnuRIB_IBAN 
         Caption         =   "IBAN"
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
Attribute VB_Name = "frmBIA_Quid"
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
Dim mSelect_Id As String

Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean
Dim mDetail_Id As String


Dim fgYSAAJRN0_FormatString As String, fgYSAAJRN0_K As Integer
Dim fgYSAAJRN0_RowDisplay As Integer, fgYSAAJRN0_RowClick As Integer, fgYSAAJRN0_ColClick As Integer
Dim fgYSAAJRN0_ColorClick As Long, fgYSAAJRN0_ColorDisplay As Long
Dim fgYSAAJRN0_Sort1 As Integer, fgYSAAJRN0_Sort2 As Integer
Dim fgYSAAJRN0_SortAD As Integer, fgYSAAJRN0_Sort1_Old As Integer
Dim fgYSAAJRN0_arrIndex As Integer
Dim blnfgYSAAJRN0_DisplayLine As Boolean

'______________________________________________________________________

Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long



Dim HeightOfLine As Long, LinesOfText As Long

Dim txtRTF_prtForeColor_Header As Long

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim mXls2_Cols As Integer, mXls2_Row As Integer

Dim blnAliasName As Boolean, mAliasName As String

'______________________________________________________________________
Dim rsSabX As New ADODB.Recordset
Dim arrJrnl_Event_Id() As String, arrJrnl_Event_Lib() As String, arrJrnl_Event_Sta() As String, arrJrnl_Event_Nb As Integer, arrJrnl_Event_K As Integer
Dim arrSWIBKUBIC_8() As String, arrSWIBKUBIC() As String, arrSWIBICIN1() As String, arrSWIBKUBIC_Sta() As String
Dim arrSWIBKUBIC_6() As String
Public Sub cmdSelect_SQL_ZBASTAB0_23_Exportation()

On Error GoTo Error_Handler
Dim X As String, K As Long, K2 As Long, xSQL As String, xWhere As String, wNum As Long
Dim wFile As String, wFilex As String
Dim blnCALCS As Boolean
Dim xLib As String
On Error GoTo Error_Handler
currentAction = "cmdSelect_SQL_ZBASTAB0_23_Exportation"
'===================================================================================
'If blnAuto Then
'    X = paramServer("\\CPT_Archive\")
'Else
    X = ""
'End If
If X = "" Then X = "C:\Temp\"
If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"

blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

xLib = "SAB OPE-EVE "

wFile = X & xLib & " " & dateImp_Amj(DSys) & ".xlsx"

'If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "SAB : liste des codes opération et événement: nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
'End If

If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile
'_________________________________________


Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "OPE_EVE"
    .Subject = ""
End With

xWhere = ""

xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
    & " where BASTABETA =  1 and BASTABNUM = 24 order by BASTABARG"
Set rsSab = cnsab.Execute(xSQL)

Call cmdSelect_SQL_ZBASTAB0_23_Exportation_Page(1, xLib)
Call cmdSelect_SQL_ZBASTAB0_23_Exportation_Detail



wbExcel.SaveAs wFile
wbExcel.Close
appExcel.Quit
'===================================================================================================
Exit_sub:
'__________________________________________________________________________________

Set rsSab = Nothing


Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
Call lstErr_AddItem(lstErr, frmElp.cmdContext, "< Exportation terminée"): DoEvents
'_____________________________
Exit Sub

Error_Handler:

If Not blnCALCS Then
    X = "C:\Temp\"
    Resume Next
End If
MsgBox Error, vbCritical, currentAction
Call lstErr_AddItem(lstErr, frmElp.cmdContext, "< Exportation terminée"): DoEvents

End Sub

Public Sub cmdSelect_SQL_ZBASTAB0_23_Exportation_Page(lSheet As Integer, lLib As String)

On Error GoTo Error_Handler
Dim K As Integer

'==========================================================================================================

Set wsExcel = wbExcel.Sheets(lSheet)
wsExcel.Name = lSheet & "-" & lLib

'__________________________________________________________________________________

With wsExcel.Cells
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(160, 160, 160)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(220, 220, 220)
    .VerticalAlignment = Excel.xlVAlignCenter
    .HorizontalAlignment = Excel.xlHAlignLeft
    .WrapText = True
    .Font.Size = 8
    .Font.Name = "Calibri"
    .Font.Color = RGB(0, 64, 128)
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14" & "SAB : liste des codes OPERATION / EVENEMENT" _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True
wsExcel.PageSetup.PrintTitleRows = "$A1:$E1"

wsExcel.PageSetup.Zoom = 100

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : " & lLib): DoEvents

mXls2_Cols = 5
mXls2_Row = 1



wsExcel.Cells(1, 1) = "Opération": wsExcel.Columns(1).ColumnWidth = 8
wsExcel.Cells(1, 2) = "Application": wsExcel.Columns(2).ColumnWidth = 8
wsExcel.Cells(1, 3) = "Evénement": wsExcel.Columns(3).ColumnWidth = 8
wsExcel.Cells(1, 4) = "Libellé opération": wsExcel.Columns(4).ColumnWidth = 45
wsExcel.Cells(1, 5) = "Libellé événement": wsExcel.Columns(5).ColumnWidth = 45


For K = 1 To mXls2_Cols
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next

'==========================================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdSelect_SQL_Exportation_Detail"


End Sub



Public Sub cmdSelect_SQL_ZBASTAB0_23_Exportation_Detail()
Dim X As String, K As Integer, K2 As Integer, mCol As Integer, xSQL As String
Dim rsSabX As ADODB.Recordset
Dim mOPE As String, mOPE_Lib As String, mOPE_App As String
On Error GoTo Error_Handler
'==========================================================================================================

Do While Not rsSab.EOF

    X = rsSab("BASTABARG")
    If mOPE <> Mid$(X, 1, 3) Then
'==========================================================================================================
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
            & " where BASTABETA =  1 and BASTABNUM = 23 and BASTABARG = '" & Mid$(X, 1, 3) & "'"
        Set rsSabX = cnsab.Execute(xSQL)
        If Not rsSabX.EOF Then
            mOPE = Trim(rsSabX("BASTABARG"))
            mOPE_App = Trim(rsSabX("BASTABLO1"))
            mOPE_Lib = Mid$(rsSabX("BASTABDON"), 1, 32)
        Else
            mOPE = Mid$(X, 1, 3)
            mOPE_App = "???"
            mOPE_Lib = "???"
        End If
        

'==========================================================================================================
    End If
    
    mXls2_Row = mXls2_Row + 1
    wsExcel.Cells(mXls2_Row, 1) = mOPE
    wsExcel.Cells(mXls2_Row, 1).Font.Color = vbBlue
    wsExcel.Cells(mXls2_Row, 2) = mOPE_App
    wsExcel.Cells(mXls2_Row, 2).Font.Color = vbMagenta
    wsExcel.Cells(mXls2_Row, 3) = Mid$(X, 4, 3)
    wsExcel.Cells(mXls2_Row, 4) = mOPE_Lib
    wsExcel.Cells(mXls2_Row, 4).Font.Color = vbBlue
    wsExcel.Cells(mXls2_Row, 5) = Mid$(rsSab("BASTABDON"), 13, 32)
'_________________________________________________________________________________________________
    rsSab.MoveNext
Loop

'==========================================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdSelect_SQL_Exportation_Detail"


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

Private Sub fgDetail_ZTCHCOR0_Display(lBIC As String)
Dim xSQL As String

On Error GoTo Error_Handler
currentAction = "fgDetail_ZTCHCOR0_Display"
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.Row = 0
'___________________________________________________________________________

xSQL = "select * from " & paramIBM_Library_SAB & ".ZTCHCOR0 , " & paramIBM_Library_SAB & ".ZSWIBIC0" _
     & " where TCHCORETB = 1 and TCHCORCOD = '001' and TCHCORRGP = 'T' and TCHCORTYP = 'B' and TCHCORCLI = '' " _
     & " and TCHCORBIC = '" & lBIC & "' and SWIBICBIC = TCHCORBI1 order by TCHCORDEV , TCHCORBI1"

Set rsSab = cnsab.Execute(xSQL)
  
Do While Not rsSab.EOF

    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_ZTCHCOR0_Display_Line
    
    rsSab.MoveNext

Loop

fgDetail.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgDetail.Rows - 1): DoEvents

'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgYSAAJRN0_Display(lBIC As String)
Dim xSQL As String, K As Integer

On Error GoTo Error_Handler
currentAction = "fgYSAAJRN0_YSAAJRN0_Display"
fgYSAAJRN0.Visible = False
'fgYSAAJRN0_Reset

fgYSAAJRN0.Rows = 1
fgYSAAJRN0.FormatString = fgYSAAJRN0_FormatString
fgYSAAJRN0.Row = 0
'___________________________________________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSAAJRN0 " _
          & " where SAAJRNAID = 0 " _
          & " and SAAJRNTOPX = '" & Mid$(lBIC, 1, 8) & "'" _
          & " order by SAAJRNAID , SAAJRNAMJH desc , SAAJRNSEQ"


Set rsSab = cnsab.Execute(xSQL)
  
Do While Not rsSab.EOF

    fgYSAAJRN0.Rows = fgYSAAJRN0.Rows + 1
    fgYSAAJRN0.Row = fgYSAAJRN0.Rows - 1
    'fgYSAAJRN0_YSAAJRN0_Display_Line
    
    
    V = DateAdd("s", 1200802192 - rsSab("SAAJRNAMJH"), "01/01/2000 00:00:00")
    
    fgYSAAJRN0.Col = 0:  fgYSAAJRN0.Text = V
    X = rsSab("SAAJRNEVEC") & " " & rsSab("SAAJRNEVEN")
    fgYSAAJRN0.Col = 1: fgYSAAJRN0.Text = X
    fgYSAAJRN0.Col = 2: fgYSAAJRN0.Text = fgYSAAJRN0_Display_Lib(X)
    fgYSAAJRN0.Col = 3:  fgYSAAJRN0.Text = rsSab("SAAJRNTOPK")
    fgYSAAJRN0.Col = 4:  fgYSAAJRN0.Text = Trim(rsSab("SAAJRNTOPX"))
    fgYSAAJRN0.Col = 5:  fgYSAAJRN0.Text = rsSab("SAAJRNAID")
    fgYSAAJRN0.Col = 6:  fgYSAAJRN0.Text = rsSab("SAAJRNAMJH")
    fgYSAAJRN0.Col = 7:  fgYSAAJRN0.Text = rsSab("SAAJRNSEQ")
    fgYSAAJRN0.Col = 8:  fgYSAAJRN0.Text = rsSab("SAAJRNSUFX")
    Select Case arrJrnl_Event_Sta(arrJrnl_Event_K)
        Case "A": For K = 0 To 8: fgYSAAJRN0.Col = K: fgYSAAJRN0.CellBackColor = mColor_G2: Next K
        Case "R": For K = 0 To 8: fgYSAAJRN0.Col = K: fgYSAAJRN0.CellBackColor = mColor_W1: Next K
    End Select
    
    rsSab.MoveNext

Loop

If fgYSAAJRN0.Rows > 5 Then fgYSAAJRN0.TopRow = fgYSAAJRN0.Rows - 5
fgYSAAJRN0.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgYSAAJRN0.Rows - 1): DoEvents

'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Function fgYSAAJRN0_Display_Lib(lK As String) As String

If lK <> arrJrnl_Event_Id(arrJrnl_Event_K) Then
    For arrJrnl_Event_K = 1 To arrJrnl_Event_Nb
        If lK = arrJrnl_Event_Id(arrJrnl_Event_K) Then Exit For
    Next arrJrnl_Event_K
End If
fgYSAAJRN0_Display_Lib = arrJrnl_Event_Lib(arrJrnl_Event_K)

End Function

Private Sub fgDetail_ZBASTAB0_23_Display(lOPE As String)
Dim xSQL As String

On Error GoTo Error_Handler
currentAction = "fgDetail_ZBASTAB0_23_Display"
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = "<Code  |<Intitulé                                                                                                          ||"
fgDetail.Row = 0
'___________________________________________________________________________

xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
     & " Where BASTABETA = 1 and BASTABNUM = 24 " _
     & " and BASTABARG like '" & lOPE & "%' order by BASTABARG"

Set rsSab = cnsab.Execute(xSQL)
  
Do While Not rsSab.EOF

    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail.Col = 0: fgDetail.Text = Mid$(rsSab("BASTABARG"), 4, 3)
    fgDetail.Col = 1: fgDetail.Text = Mid$(rsSab("BASTABDON"), 13, 32)
    
    rsSab.MoveNext

Loop

fgDetail.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgDetail.Rows - 1): DoEvents

'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgDetail_ZFICBDF0_Display(lCIB As String)
Dim X As String, xWhere As String
Dim xSQL As String

On Error GoTo Error_Handler

currentAction = "fgDetail_ZFICBDF0_Display"
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = "<Guichet  |<BIC                     |<Nom du guichet                      |<CP Ville                                              |<Adresse                                                                                                 |"
fgDetail.Row = 0
'___________________________________________________________________________

xWhere = " Where substring(ZFICBDF0 , 1 , 6 ) ='3" & lCIB & "' and substring (ZFICBDF0 , 299 , 1 ) not in ('A' , 'P')"

X = Format(Trim(txtSelect_FICBDFGUI), "00000")
If X <> "" Then xWhere = xWhere & " and substring(ZFICBDF0 , 7 , 5 ) ='" & X & "'"

X = Trim(txtSelect_FICBDFGNOM)
If X <> "" Then xWhere = xWhere & " and substring(ZFICBDF0 , 53, 20 ) like '%" & X & "%'"

X = Trim(txtSelect_FICBDFGCP)
If X <> "" Then xWhere = xWhere & " and substring(ZFICBDF0 , 203 , 5) like '" & X & "%'"

X = Trim(txtSelect_FICBDFGVIL)
If X <> "" Then xWhere = xWhere & " and substring(ZFICBDF0 , 203 , 32 ) like '%" & X & "%'"
    
xSQL = "select * from " & paramIBM_Library_SABSPE & ".ZFICBDF0 " & xWhere & " order by substring(ZFICBDF0 , 1 , 6 )"

Set rsSab = cnsab.Execute(xSQL)
  
Do While Not rsSab.EOF

    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_ZFICBDF0_Display_Line
    
    rsSab.MoveNext

Loop

fgDetail.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgDetail.Rows - 1): DoEvents

'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub



Public Sub fgDetail_ZTCHCOR0_Display_Line()
Dim blnBold As Boolean

On Error Resume Next
If rsSab("TCHCORMTR") = 1 Then
    blnBold = True
Else
    blnBold = False
End If

fgDetail.Col = 0: fgDetail.Text = rsSab("TCHCORDEV")
fgDetail.CellFontBold = blnBold
fgDetail.Col = 1: fgDetail.Text = rsSab("TCHCORMTR")

fgDetail.Col = 2: fgDetail.Text = rsSab("SWIBICBIC")
fgDetail.CellFontBold = blnBold
fgDetail.Col = 3: fgDetail.Text = Trim(rsSab("SWIBICIN1")) & Trim(rsSab("SWIBICIN2")) & Trim(rsSab("SWIBICIN3"))
fgDetail.Col = 4: fgDetail.Text = Trim(rsSab("SWIBICVIL"))
fgDetail.Col = 5: fgDetail.Text = Trim(rsSab("SWIBICCOM"))


End Sub


Public Sub fgDetail_ZFICBDF0_Display_Line()
Dim X As String, wColor As Long

On Error Resume Next
X = rsSab("ZFICBDF0")

fgDetail.Col = 0: fgDetail.Text = Mid$(X, 7, 5)
fgDetail.Col = 1: fgDetail.Text = Trim(Mid$(X, 235, 11))
fgDetail.Col = 2: fgDetail.Text = Trim(Mid$(X, 53, 20))
fgDetail.Col = 3: fgDetail.Text = Trim(Mid$(X, 203, 32))
fgDetail.Col = 4: fgDetail.Text = Trim(Mid$(X, 107, 32 * 3))


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




Public Sub fgDetail_Sort()
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



Public Sub Form_Init()
Dim V, xSQL As String, X As String
Dim K As Long

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True


cmdReset
blnControl = False

fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False

fgDetail_FormatString = fgDetail.FormatString
fgDetail.Enabled = True
fgDetail.Visible = False
fgDetail.Top = fgSelect.Top
fgDetail.Left = 3500

fgYSAAJRN0_FormatString = fgYSAAJRN0.FormatString
fgYSAAJRN0.Enabled = True
fgYSAAJRN0.Visible = False
fgYSAAJRN0.Top = fgSelect.Top + 4800
fgYSAAJRN0.Left = 3500

fraSelect_Options_ZFICBDF0.Visible = False
Set fraSelect_Options_ZFICBDF0.Container = fraSelect
fraSelect_Options_ZFICBDF0.Top = fraSelect_Options.Top
fraSelect_Options_ZFICBDF0.Left = fraSelect_Options.Left



'Initialisation PAYS ______________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdPrint, "Initialisation pays ")

cboSelect_SWIBICPays.Clear
cboSelect_SWIBICPays.AddItem "  "

X = "select * from  " & paramIBM_Library_SAB & ".ZBASTAB0 " _
    & " where BASTABETA = 1 and BASTABNUM = 11 order by BASTABARG"
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    cboSelect_SWIBICPays.AddItem Mid$(rsSab("BASTABARG"), 4, 2) & " - " & Mid$(rsSab("BASTABLO2"), 4, 16)
    rsSab.MoveNext
Loop

'______________________________________________________________________
xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'SAA' and BIATABK1 = 'Jrnl_Event' and BIATABK2 like 'RMS  %'"
Set rsSab = cnsab.Execute(xSQL)
arrJrnl_Event_Nb = rsSab(0) + 1
ReDim arrJrnl_Event_Id(arrJrnl_Event_Nb), arrJrnl_Event_Lib(arrJrnl_Event_Nb), arrJrnl_Event_Sta(arrJrnl_Event_Nb)
arrJrnl_Event_Nb = 0

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'SAA' and BIATABK1 = 'Jrnl_Event' and BIATABK2 like 'RMS  %'" _
     & "  order by BIATABK2"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrJrnl_Event_Nb = arrJrnl_Event_Nb + 1
    arrJrnl_Event_Id(arrJrnl_Event_Nb) = Trim(rsSab("BIATABK2"))
    arrJrnl_Event_Lib(arrJrnl_Event_Nb) = Trim(Mid$(rsSab("BIATABTXT"), 1, 99))
    arrJrnl_Event_Sta(arrJrnl_Event_Nb) = Trim(Mid$(rsSab("BIATABTXT"), 104, 1))
    rsSab.MoveNext
Loop

fraSelect_Options.Visible = True



If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0

blnControl = True


cmdSelect_Reset
Me.Enabled = True

End Sub



'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
currentActiveControl_Name = C.Name
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
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

'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub


Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = lK
    Select Case lK
'        Case 3: fgSelect.Col = 3: X = Format$(Val(fgSelect.Text), "000000000000000.00")

    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I

fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
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



'______________________________________________________________________
Private Sub fgSelect_ZSWIBIC0_Display()

Dim K As Long

On Error GoTo Error_Handler
currentAction = "fgSelect_ZSWIBIC0_Display"
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Row = 0

Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_ZSWIBIC0_Display_Line
    
    rsSab.MoveNext

Loop

fgSelect.Visible = True

If fgSelect.Rows = 2 Then
    fgSelect.Col = 0
    Call fgDetail_ZTCHCOR0_Display(Trim(fgSelect.Text))
    Call fgYSAAJRN0_Display(Trim(fgSelect.Text))
End If

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgSelect_ZFICBDF0_Display()

Dim K As Long

On Error GoTo Error_Handler
currentAction = "fgSelect_ZFICBDF0_Display"
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<CIB        |<Dénomination                                                                                |< Adresse                                                                                                          |<CP Ville                                         "
                 
fgSelect.Row = 0

Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_ZFICBDF0_Display_Line
    
    rsSab.MoveNext

Loop

fgSelect.Visible = True

If fgSelect.Rows = 2 Then
    fgSelect.Col = 0: mSelect_Id = Trim(fgSelect.Text)

    Call fgDetail_ZFICBDF0_Display(mSelect_Id)
End If

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgSelect_ZBASTAB0_23_Display()

Dim K As Long

On Error GoTo Error_Handler
currentAction = "fgSelect_ZFICBDF0_Display"
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Code  |<Application|<Intitulé                                                                                                          ||"
                 
fgSelect.Row = 0

Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect.Col = 0: fgSelect.Text = Trim(rsSab("BASTABARG"))
    fgSelect.Col = 1: fgSelect.Text = Trim(rsSab("BASTABLO1"))
    fgSelect.Col = 2: fgSelect.Text = Mid$(rsSab("BASTABDON"), 1, 32)
    
    rsSab.MoveNext

Loop

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_ZSWIBIC0_Display_Line()

Dim X As String, wColor As Long, wColor2 As Long

On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = rsSab("SWIBICBIC")
If IsNull(rsSab("SWIBKUBIC")) Then
    wColor = vbMagenta
    wColor2 = vbBlue
    fgSelect.CellFontBold = False
Else
    wColor = RGB(0, 96, 32)
    wColor2 = RGB(0, 96, 32)
    fgSelect.CellFontBold = True
End If

fgSelect.CellForeColor = wColor
fgSelect.Col = 1: fgSelect.Text = Trim(rsSab("SWIBICIN1")) & Trim(rsSab("SWIBICIN2")) & Trim(rsSab("SWIBICIN3"))
    'DR 05/12/2014
    fgSelect.CellForeColor = wColor2
    fgSelect.Col = 2: fgSelect.Text = Retourne_PaysISO(Mid(Trim(rsSab("SWIBICBIC")), 5, 2))
fgSelect.CellForeColor = wColor2
fgSelect.Col = 3: fgSelect.Text = Trim(rsSab("SWIBICVIL"))
fgSelect.CellForeColor = wColor2
fgSelect.Col = 4: fgSelect.Text = Trim(rsSab("SWIBICCOM"))
fgSelect.CellForeColor = wColor2
End Sub


Public Sub fgSelect_ZFICBDF0_Display_Line()
Dim X As String

On Error Resume Next
X = rsSab("ZFICBDF0")
fgSelect.Col = 0: fgSelect.Text = Mid$(X, 2, 5)
fgSelect.Col = 1: fgSelect.Text = Mid$(X, 13, 40)
fgSelect.Col = 3: fgSelect.Text = Mid$(X, 186, 32)
fgSelect.Col = 2: fgSelect.Text = Trim(Mid$(X, 90, 32 * 3))
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim wFct As String

mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

wFct = UCase$(Trim(Mid$(Msg, 1, 12)))
Call BIA_VB_HAB(wFct, arrHab(), cboSelect_SQL)

Form_Init
Select Case wFct
    Case "@RMA_CTL": blnAuto = True
        cmdSelect_SQL_K = "RMA_CTL"
        Call cmdSelect_SQL_RMA_CTL
        Call mnuPrint2_Mail_Click
        Unload Me
    Case Else: blnAuto = False

End Select
End Sub



Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgSelect.Visible = False
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = 1 To 0 Step -1
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 1 To 0 Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
    End If
End If
fgSelect.LeftCol = fgSelect.FixedCols
fgSelect.Visible = True
End Sub


Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

End Sub

Private Sub chkSelect_SWIBICBIC_Click()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub chkSelect_SWIBKUBIC_Click()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cmdPrint_Click()

Select Case cmdSelect_SQL_K
    Case "1": 'DR 05/12/2014 sortie vers fichier CSV
              Call fgSelect_Print_CSV
    Case "3": cmdSelect_SQL_ZBASTAB0_23_Exportation
    Case "RMA_CTL", "RMA_SAA_CTL", "Swift_Alias": Me.PopupMenu mnuPrint2, vbPopupMenuLeftButton
End Select
End Sub

Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim mRib_Clé  As String, mRib_IbanE As String
On Error Resume Next


If y <= fgDetail.RowHeightMin Then
    fgDetail.Visible = False
    Select Case fgDetail.Col
        Case 0: fgDetail_Sort1 = 0: fgDetail_Sort2 = 3: fgDetail_Sort
        Case 1:  fgDetail_Sort1 = 1: fgDetail_Sort2 = 3: fgDetail_Sort
        Case 2: fgDetail_Sort1 = 2: fgDetail_Sort2 = 2: fgDetail_Sort
        Case 3: fgDetail_Sort1 = 3: fgDetail_Sort2 = 3: fgDetail_Sort
        Case 4: fgDetail_Sort1 = 4: fgDetail_Sort2 = 4: fgDetail_Sort
    End Select
    fgDetail.Visible = True
Else
    If fgDetail.Rows > 1 Then
        Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
       If cmdSelect_SQL_K = "2" Then
            If Trim(txtSelect_FICBDFCPT) <> "" Then
                fgSelect.Col = 0: mDetail_Id = Trim(fgDetail.Text)
                mRib_Clé = Format$(RibClé(mSelect_Id, mDetail_Id, Trim(txtSelect_FICBDFCPT), mRib_IbanE), "00")
                mnuRIB_Clé.Caption = "Clé : " & mRib_Clé
                mnuRIB_IBAN.Caption = "IBAN : " & Iban_Print(mRib_IbanE)
                Me.PopupMenu mnuRIB, vbPopupMenuLeftButton
                'Debug.Print mRib_Clé, mRib_IbanE
            End If
        End If
        

   End If
End If
fgDetail.LeftCol = 0


End Sub
Private Sub mnuPrint2_Excel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String

Select Case cmdSelect_SQL_K
    Case "Swift_Alias":
        X = mAliasName
        Call MSflexGrid_Excel("", "Swift_Alias", X, fgSelect, fgSelect.Cols - 1)
    Case "RMA_CTL":
        X = "Contrôle RMA : SAB / SAA (journal des événements YSAAJRN0)"
        Call MSflexGrid_Excel("", "RMA_CTL", X, fgSelect, 3)
    Case "RMA_SAA_CTL":
        X = "Contrôle RMA : SAB / SAA (fichier.txt des RMA généré sur la plateforme Alliance)"
        Call MSflexGrid_Excel("", "RMA_SAA_CTL", X, fgSelect, 3)
    End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub



Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next


If y <= fgSelect.RowHeightMin Then
    fgSelect.Visible = False
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
    End Select
    fgSelect.Visible = True
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        Select Case cmdSelect_SQL_K
            Case "1"
                fgSelect.Col = 0: mSelect_Id = Trim(fgSelect.Text)
                Call fgDetail_ZTCHCOR0_Display(mSelect_Id)
                Call fgYSAAJRN0_Display(mSelect_Id)
            Case "2"
                fgSelect.Col = 0: mSelect_Id = Trim(fgSelect.Text)
                Call fgDetail_ZFICBDF0_Display(mSelect_Id)
             Case "3"
                fgSelect.Col = 0: mSelect_Id = Trim(fgSelect.Text)
                Call fgDetail_ZBASTAB0_23_Display(mSelect_Id)
       End Select
        
   End If
End If
fgSelect.LeftCol = 0


End Sub

Private Sub Form_Activate()
Set XForm = Me

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

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
blnControl = True

End Sub

Public Sub cmdSelect_Clear()

lstErr.Clear
fgSelect.Visible = False
fgDetail.Visible = False: fgYSAAJRN0.Visible = False
cmdSelect_Ok.BackColor = vbGreen

End Sub

Public Sub cmdSelect_Reset()
Dim K As Integer
If blnControl Then
    cmdSelect_Clear
    K = InStr(cboSelect_SQL, "-")
    If K > 1 Then
        cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, K - 1))
    Else
        cmdSelect_SQL_K = "???"
    End If
    
    fraSelect_Options.Visible = False
    fraSelect_Options_ZFICBDF0.Visible = False
    fgDetail.Height = 7500
    
    Select Case cmdSelect_SQL_K
        Case "1": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True: fgDetail.Height = 4600
        Case "2": cmdSelect_Ok.Visible = True: fraSelect_Options_ZFICBDF0.Visible = True
        Case "3"
        Case "Swift_Alias"
    End Select

End If
End Sub


Private Sub cmdSelect_SQL_ZSWIBIC0()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"
xWhere = ""

X = Trim(txtSelect_SWIBICBIC)
If X <> "" Then xWhere = xWhere & " and SWIBICBIC like '" & X & "%'"

If chkSelect_SWIBICBIC = "1" Then xWhere = xWhere & " and SWIBICBIC like '%XXX'"
    
X = Trim(Mid$(cboSelect_SWIBICPays, 1, 2))
If X <> "" Then xWhere = xWhere & " and substring(SWIBICBIC,5,2) = '" & X & "'"

X = Trim(txtSelect_SWIBICIN1)
If X <> "" Then xWhere = xWhere & " and SWIBICIN1 like '%" & X & "%'"
    
X = Trim(txtSelect_SWIBICVIL)
If X <> "" Then xWhere = xWhere & " and SWIBICVIL like '%" & X & "%'"

xWhere = Replace(xWhere, "and", "where", 1, 1)

If chkSelect_SWIBKUBIC = "1" Then

    If xWhere <> "" Then 'DR 05/12/2014
            xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIBIC0 " _
                 & " , " & paramIBM_Library_SAB & ".ZSWIBKU0 " & xWhere & " and SWIBICBIC = SWIBKUBIC order by SWIBICBIC"
        Else
            xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIBIC0 " _
                 & " , " & paramIBM_Library_SAB & ".ZSWIBKU0 where SWIBICBIC = SWIBKUBIC order by SWIBICBIC"
        End If
    Else
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIBIC0 " _
             & " left outer join " & paramIBM_Library_SAB & ".ZSWIBKU0 on SWIBICBIC = SWIBKUBIC " & xWhere & " order by SWIBICBIC"
    End If


Set rsSab = cnsab.Execute(xSQL)
  

fgSelect_ZSWIBIC0_Display

Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_RMA_CTL()
Dim X As String, K As Integer, mSAAJRNTOPX As String, arrSWIBKUBIC_Loop As Integer, arrSWIBKUBIC_K As Integer
Dim xSQL As String, Nb As Long, xSta As String, iCol As Integer
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_RMA_CTL"
'_______________________________________________________________________________________
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Commentaire                                                                    |<BIC                                |< Intitulé                                                                                                           |||"
fgSelect.Row = 0


xSQL = "select count(*) from " & paramIBM_Library_SAB & ".ZSWIBKU0"

Set rsSab = cnsab.Execute(xSQL)
Nb = rsSab(0) + 10
ReDim arrSWIBKUBIC_8(Nb), arrSWIBKUBIC(Nb), arrSWIBKUBIC_Sta(Nb)

Nb = 0
xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIBKU0" _
     & " left outer join " & paramIBM_Library_SAB & ".ZSWIBIC0 on SWIBICBIC = SWIBKUBIC " _
     & " order by SWIBKUBIC"


Set rsSab = cnsab.Execute(xSQL)
  
Do While Not rsSab.EOF
    If IsNull(rsSab("SWIBICBIC")) Then
        If rsSab("SWIBKUBIC") <> "SOGEFRPPTGV" Then
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect.Col = 1: fgSelect.Text = rsSab("SWIBKUBIC"): fgSelect.CellForeColor = vbRed: fgSelect.CellFontBold = True
            fgSelect.Col = 0: fgSelect.Text = "BIC inconnu, RMA à supprimer dans SAB": fgSelect.CellForeColor = vbRed
        End If
    Else
        Nb = Nb + 1
        arrSWIBKUBIC_8(Nb) = Mid$(rsSab("SWIBKUBIC"), 1, 8)
        arrSWIBKUBIC(Nb) = Trim(rsSab("SWIBKUBIC"))
        arrSWIBKUBIC_Sta(Nb) = ""
    End If
    rsSab.MoveNext

Loop

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSAAJRN0 " _
          & " where SAAJRNAID = 0 " _
          & " and SAAJRNEVEC = 'RMS'  and SAAJRNTOPK = 'B'" _
          & " order by SAAJRNTOPX, SAAJRNAID , SAAJRNAMJH desc , SAAJRNSEQ"


Set rsSab = cnsab.Execute(xSQL)
  
Do While Not rsSab.EOF
    X = Trim(rsSab("SAAJRNTOPX"))
    If mSAAJRNTOPX <> X Then
        If arrSWIBKUBIC_K = 0 Then
            If arrSWIBKUBIC_Sta(0) = "A" Then
                xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIBIC0 " _
                     & " where SWIBICBIC like '" & mSAAJRNTOPX & "%'"
                
                Set rsSabX = cnsab.Execute(xSQL)
                  
                If Not rsSabX.EOF Then
                    fgSelect.Rows = fgSelect.Rows + 1
                    fgSelect.Row = fgSelect.Rows - 1
                    fgSelect.Col = 1: fgSelect.Text = mSAAJRNTOPX: fgSelect.CellForeColor = vbBlue: fgSelect.CellFontBold = True
                    fgSelect.Col = 2: fgSelect.Text = Trim(rsSabX("SWIBICIN1")): fgSelect.CellForeColor = vbBlue
                    fgSelect.Col = 0: fgSelect.Text = "RMA à ajouter dans SAB": fgSelect.CellForeColor = vbBlue

               '     Debug.Print "manquant SAB: "; mSAAJRNTOPX, Trim(rsSabX("SWIBICIN1"))
               ' Else
                   Debug.Print "RMA SAA obsolète : "; mSAAJRNTOPX
                End If
            End If
        'Else
        '    arrSWIBKUBIC_Sta(arrSWIBKUBIC_K) = xSta
        End If
        mSAAJRNTOPX = X
        arrSWIBKUBIC_K = 0
        For K = arrSWIBKUBIC_Loop + 1 To Nb
            If mSAAJRNTOPX = arrSWIBKUBIC_8(K) Then
                arrSWIBKUBIC_Loop = K
                arrSWIBKUBIC_K = K
                Exit For
            Else
                If mSAAJRNTOPX < arrSWIBKUBIC_8(K) Then Exit For
            End If
        
        Next K
        arrSWIBKUBIC_Sta(arrSWIBKUBIC_K) = ""
    End If
    
    X = fgYSAAJRN0_Display_Lib(rsSab("SAAJRNEVEC") & " " & rsSab("SAAJRNEVEN"))
    
    Select Case arrJrnl_Event_Sta(arrJrnl_Event_K)
        Case "A": arrSWIBKUBIC_Sta(arrSWIBKUBIC_K) = "A"
        Case "R": arrSWIBKUBIC_Sta(arrSWIBKUBIC_K) = "R"
    End Select

    rsSab.MoveNext

Loop


For K = 2 To Nb
    If Mid$(arrSWIBKUBIC(K), 1, 8) = Mid$(arrSWIBKUBIC(K - 1), 1, 8) Then
        arrSWIBKUBIC_Sta(K) = arrSWIBKUBIC_Sta(K - 1)
    End If

Next K

For K = 1 To Nb
    If arrSWIBKUBIC_Sta(K) <> "A" Then
    
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIBIC0 " _
             & " where SWIBICBIC like '" & arrSWIBKUBIC(K) & "%'"
        
        Set rsSabX = cnsab.Execute(xSQL)
          
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect.Col = 1: fgSelect.Text = arrSWIBKUBIC(K): fgSelect.CellFontBold = True
        If rsSabX.EOF Then
            'Debug.Print "SAB : BIC Inconnu : "; arrSWIBKUBIC(K)
            fgSelect.CellForeColor = vbBlack
            fgSelect.Col = 2: fgSelect.Text = "": fgSelect.CellForeColor = vbBlack
            fgSelect.Col = 0: fgSelect.Text = "BIC inconnu à supprimer dans SAB": fgSelect.CellForeColor = vbBlack
        Else
            Select Case arrSWIBKUBIC_Sta(K)
                Case "R": 'Debug.Print "SAB : à Révoquer : "; arrSWIBKUBIC(K), Trim(rsSabX("SWIBICIN1"))
                    fgSelect.CellForeColor = vbRed
                    fgSelect.Col = 2: fgSelect.Text = Trim(rsSabX("SWIBICIN1")): fgSelect.CellForeColor = vbRed
                    fgSelect.Col = 0: fgSelect.Text = "RMA révoqué, BIC à supprimer dans SAB": fgSelect.CellForeColor = vbRed
                
                Case Else: ' Debug.Print "SAB : pas de RMA échangé : "; arrSWIBKUBIC(K), Trim(rsSabX("SWIBICIN1"))
                    fgSelect.CellForeColor = vbMagenta
                    fgSelect.Col = 2: fgSelect.Text = Trim(rsSabX("SWIBICIN1")): fgSelect.CellForeColor = vbMagenta
                    fgSelect.Col = 0: fgSelect.Text = "pas de RMA échangé , BIC à supprimer dans SAB": fgSelect.CellForeColor = vbMagenta
            End Select
        End If
    End If
Next K

Set rsSab = Nothing
Set rsSabX = Nothing
fgSelect.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub cmdSelect_SQL_ZFICBDF0()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_ZFICBDF0"
xWhere = " Where substring(ZFICBDF0 , 1 , 1 ) ='1'"

X = Format(Trim(txtSelect_FICBDFBCIB), "00000")
If X <> "" Then xWhere = xWhere & " and substring(ZFICBDF0 , 2 , 5 ) ='" & X & "'"

X = Trim(txtSelect_FICBDFBNOM)
If X <> "" Then xWhere = xWhere & " and substring(ZFICBDF0 , 13 , 40 ) like '%" & X & "%'"

X = Trim(txtSelect_FICBDFBCP)
If X <> "" Then xWhere = xWhere & " and substring(ZFICBDF0 , 186 , 5) like '" & X & "%'"

X = Trim(txtSelect_FICBDFBVIL)
If X <> "" Then xWhere = xWhere & " and substring(ZFICBDF0 , 90 , 32 *4) like '%" & X & "%'"
    
xSQL = "select * from " & paramIBM_Library_SABSPE & ".ZFICBDF0 " & xWhere & " order by substring(ZFICBDF0 , 1 , 6 )"

Set rsSab = cnsab.Execute(xSQL)
  

fgSelect_ZFICBDF0_Display

Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_ZBASTAB0_23()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_ZBASTAB0_23"
    
xWhere = " Where BASTABETA = 1 and BASTABNUM = 23"
xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 " & xWhere & " order by BASTABARG"

Set rsSab = cnsab.Execute(xSQL)
  
fgSelect_ZBASTAB0_23_Display

Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub

Private Sub fgSelect_Print_CSV()
Dim FicSortie As String
Dim lngFicSortie As Long
Dim ligOut As String
Dim lig As Long

    If fgSelect.Rows > 2 Then
        FicSortie = paramTemp_Folder & "\" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhnnss") & "_ListeBic.csv"
        FicSortie = InputBox("Veuillez confirmer le nom du fichier de sortie.", "Liste des BIC", FicSortie)
        If FicSortie <> "" Then
            lngFicSortie = FreeFile
            Open FicSortie For Output As #lngFicSortie
            ligOut = "BIC;Intitulé;Pays;Ville;Commentaires"
            Print #lngFicSortie, ligOut
            For lig = 1 To fgSelect.Rows - 1
                fgSelect.Row = lig
                fgSelect.Col = 0: ligOut = fgSelect.Text & ";"
                fgSelect.Col = 1: ligOut = ligOut & fgSelect.Text & ";"
                fgSelect.Col = 2: ligOut = ligOut & fgSelect.Text & ";"
                fgSelect.Col = 3: ligOut = ligOut & fgSelect.Text
                Print #lngFicSortie, ligOut
            Next lig
            Close #lngFicSortie
            MsgBox ("Le fichier a été généré !")
        End If
    Else
        MsgBox ("Il n'y a aucune ligne à exporter !")
    End If
    

End Sub

Private Function Retourne_PaysISO(iso As String) As String
Dim xSQL As String
Dim newRs As ADODB.Recordset
Static mISO As String, mPaysIso As String

If mISO <> iso Then
    mISO = iso
    mPaysIso = iso
    xSQL = "SELECT SUBSTR(BASTABLO1, 4, 9) AS RISO FROM " & paramIBM_Library_SAB & ".ZBASTAB0"
    xSQL = xSQL & " WHERE BASTABNUM = 11"
    xSQL = xSQL & " AND SUBSTR(BASTABARG, 4, 2)  = '" & UCase(iso) & "'"
    Set newRs = cnsab.Execute(xSQL)
    If Not newRs.EOF Then
        mPaysIso = newRs("RISO")
    End If
    newRs.Close
    Set newRs = Nothing
End If
Retourne_PaysISO = mPaysIso
End Function

Public Sub cmdContext_Return()
    If SSTab1.Tab = 0 Then
        cmdSelect_Ok_Click
    Else
        SendKeys "{TAB}"
    End If
End Sub


Public Sub cmdContext_Quit()
lstErr.Clear: lstErr.Height = 200

If txtRTF.Visible Then
    txtRTF.Visible = False
    Exit Sub
End If

If txtFg.Visible Then
    txtFg.Visible = False
    Exit Sub
End If

If fgDetail.Visible Then
    fgDetail.Visible = False: fgYSAAJRN0.Visible = False
    Exit Sub
End If

If fgSelect.Visible Then
    fgSelect.Visible = False
    Exit Sub
End If


Unload Me

End Sub

Private Sub Form_Load()


mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False

End Sub


Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub lstErr_Click()
If lstErr.Height > 500 Then
    lstErr.Height = 480
Else
    lstErr.Height = lstErr.ListCount * 200 + 300
End If

End Sub

Private Sub mnuPrint2_Mail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim xObjet As String, xMesg As String, xDest As String
Select Case cmdSelect_SQL_K
    Case "Swift_Alias":
            xObjet = mAliasName
            xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
             & xObjet
    
            Call MSFlexGrid_SendMail(currentSSIWINMAIL, "Swift_Alias", xObjet, xMesg, fgSelect, fgSelect.Cols - 1)

    Case "RMA_CTL":
            xObjet = "Contrôle RMA : SAB / SAA - " & Date & " " & Time
            xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
             & xObjet
            If blnAuto Then
                xDest = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire("S60_RMA")
            Else
                xDest = currentSSIWINMAIL
            End If
            Call MSFlexGrid_SendMail(xDest, "RMA_CTL", xObjet, xMesg, fgSelect, 3)
    Case "RMA_SAA_CTL":
            xObjet = "Contrôle RMA : SAB / SAA (fichier.txt) - " & Date & " " & Time
            xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
             & xObjet
            If blnAuto Then
                xDest = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire("S60_RMA")
            Else
                xDest = currentSSIWINMAIL
            End If
            Call MSFlexGrid_SendMail(xDest, "RMA_SAA_CTL", xObjet, xMesg, fgSelect, 3)
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub txtSelect_FICBDFCPT_GotFocus()
txt_GotFocus txtSelect_FICBDFCPT

End Sub

Private Sub txtSelect_FICBDFCPT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelect_FICBDFCPT_LostFocus()
txt_LostFocus txtSelect_FICBDFCPT

End Sub

Private Sub txtSelect_FICBDFGCP_GotFocus()
txt_GotFocus txtSelect_FICBDFGCP
End Sub

Private Sub txtSelect_FICBDFGCP_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub


Private Sub txtSelect_FICBDFGCP_LostFocus()
txt_LostFocus txtSelect_FICBDFGCP
End Sub

Private Sub txtSelect_FICBDFGNOM_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub txtSelect_FICBDFGNOM_GotFocus()
txt_GotFocus txtSelect_FICBDFGNOM
End Sub


Private Sub txtSelect_FICBDFGNOM_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelect_FICBDFGNOM_LostFocus()
txt_LostFocus txtSelect_FICBDFGNOM
End Sub

Private Sub txtSelect_FICBDFBCIB_Change()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub txtSelect_FICBDFBCIB_GotFocus()
txt_GotFocus txtSelect_FICBDFBCIB
End Sub


Private Sub txtSelect_FICBDFBCIB_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtSelect_FICBDFBCIB_LostFocus()
txt_LostFocus txtSelect_FICBDFBCIB
End Sub

Private Sub txtSelect_FICBDFBCP_Change()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub txtSelect_FICBDFBCP_GotFocus()
txt_GotFocus txtSelect_FICBDFBCP
End Sub


Private Sub txtSelect_FICBDFBCP_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtSelect_FICBDFBCP_LostFocus()
txt_LostFocus txtSelect_FICBDFBCP
End Sub

Private Sub txtSelect_FICBDFGUI_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub txtSelect_FICBDFGUI_GotFocus()
txt_GotFocus txtSelect_FICBDFGUI
End Sub


Private Sub txtSelect_FICBDFGUI_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtSelect_FICBDFGUI_LostFocus()
txt_LostFocus txtSelect_FICBDFGUI
End Sub

Private Sub txtSelect_FICBDFBNOM_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub txtSelect_FICBDFBNOM_GotFocus()
txt_GotFocus txtSelect_FICBDFBNOM
End Sub


Private Sub txtSelect_FICBDFBNOM_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtSelect_FICBDFBNOM_LostFocus()
txt_LostFocus txtSelect_FICBDFBNOM
End Sub

Private Sub txtSelect_FICBDFBVIL_Change()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub txtSelect_FICBDFBVIL_GotFocus()
txt_GotFocus txtSelect_FICBDFBVIL
End Sub


Private Sub txtSelect_FICBDFBVIL_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtSelect_FICBDFBVIL_LostFocus()
txt_LostFocus txtSelect_FICBDFBVIL
End Sub

Private Sub txtSelect_FICBDFGVIL_GotFocus()
txt_GotFocus txtSelect_FICBDFGVIL
End Sub


Private Sub txtSelect_FICBDFGVIL_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSelect_FICBDFGVIL_LostFocus()
txt_LostFocus txtSelect_FICBDFGVIL
End Sub

Private Sub txtSelect_SWIBICBIC_Change()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub txtSelect_SWIBICBIC_GotFocus()
txt_GotFocus txtSelect_SWIBICBIC
End Sub


Private Sub txtSelect_SWIBICBIC_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelect_SWIBICBIC_LostFocus()
txt_LostFocus txtSelect_SWIBICBIC
End Sub

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_Quid_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Clear

Select Case cmdSelect_SQL_K
    Case "1": cmdSelect_SQL_ZSWIBIC0
    Case "2": cmdSelect_SQL_ZFICBDF0
    Case "3": cmdSelect_SQL_ZBASTAB0_23
    Case "Swift_Alias": cmdSelect_SQL_Swift_Alias
    Case "RMA_CTL": cmdSelect_SQL_RMA_CTL
    Case "RMA_SAA_CTL": cmdSelect_SQL_RMA_SAA_CTL
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_Quid_cmdSelect_Ok"): DoEvents
lstErr.Height = 480
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
cmdSelect_Ok.BackColor = fgSelect.BackColorFixed
End Sub



Private Sub txtSelect_SWIBICIN1_Change()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub txtSelect_SWIBICIN1_GotFocus()
txt_GotFocus txtSelect_SWIBICIN1
End Sub


Private Sub txtSelect_SWIBICIN1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelect_SWIBICIN1_LostFocus()
txt_LostFocus txtSelect_SWIBICIN1
End Sub

Private Sub txtSelect_SWIBICVIL_Change()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub txtSelect_SWIBICVIL_GotFocus()
txt_GotFocus txtSelect_SWIBICVIL
End Sub


Private Sub txtSelect_SWIBICVIL_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelect_SWIBICVIL_LostFocus()
txt_LostFocus txtSelect_SWIBICVIL
End Sub



Public Sub cmdSelect_SQL_Swift_Alias()

On Error GoTo Error_Handler

Dim X As String, wFile As String, wFilex As String
Dim K As Integer
currentAction = "cmdSelect_SQL_Swift_Alias"

wFile = "C:\Temp\Swift Alias.txt"

X = InputBox("par défaut : " _
    & vbCrLf & "     =========================" & vbCrLf & wFile _
    & vbCrLf & "     =========================", "SSwift_Alias : nom du fichier à traiter", wFile)
If Trim(X) = "" Then Exit Sub
wFilex = Trim(X)
'______________________________________________
If wFile <> wFilex Then
    wFile = wFilex
End If

If Dir(wFile) = "" Then V = "Ce fichier n'exite pas": GoTo Error_MsgBox
'_______________________________________________________________________________________
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Row = 0

Open wFile For Input As #1

Do Until EOF(1)
    Line Input #1, X
    X = Trim(X)
    If Not blnAliasName Then
        If InStr(X, "Alias Name") > 0 Then
            blnAliasName = True
            mAliasName = Replace(X, " ", "")
        End If
    Else
        If InStr(X, "Institution") > 0 Then
            K = InStr(12, X, "=")
            X = Trim(Mid$(X, K + 1, Len(X) - K))
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect.Col = 0: fgSelect.Text = X
            X = "select * from " & paramIBM_Library_SAB & ".ZSWIBIC0 where SWIBICBIC = '" & X & "'"
            Set rsSab = cnsab.Execute(X)
            
            If Not rsSab.EOF Then
                fgSelect.Col = 1: fgSelect.Text = Trim(rsSab("SWIBICIN1")) & Trim(rsSab("SWIBICIN2")) & Trim(rsSab("SWIBICIN3"))
                fgSelect.Col = 2: fgSelect.Text = Trim(rsSab("SWIBICVIL"))
                fgSelect.Col = 3: fgSelect.Text = Trim(rsSab("SWIBICCOM"))
            End If
        End If
    End If

Loop

fgSelect.Visible = True

Close #1

'_______________________________________________________________________________________
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Public Sub cmdSelect_SQL_RMA_SAA_CTL()
'
' lecture d'un fichier issu de SAA - Relationship Management (Details : all)
' Own BIC:           BIARFRPP

On Error GoTo Error_Handler

Dim X As String, wFile As String, wFilex As String
Dim K As Integer

Dim blnBic_BIARFRPP As Boolean, blnBic_Correspondent As Boolean, blnReceive As Boolean, blnSend As Boolean
Dim wBIC As String, wReceive As String, wSend As String
Dim wBIC_6 As String

currentAction = "cmdSelect_SQL_RMA_SAA_CTL"
'________________________________________________________________________________________________________________
wFile = "C:\Temp\RMA_SAA.txt"

X = InputBox("par défaut : " _
    & vbCrLf & "     =========================" & vbCrLf & wFile _
    & vbCrLf & "     =========================", "SRMA_SAA_CTL : nom du fichier à traiter", wFile)
If Trim(X) = "" Then Exit Sub
wFilex = Trim(X)
'______________________________________________
If wFile <> wFilex Then
    wFile = wFilex
End If

If Dir(wFile) = "" Then V = "Ce fichier n'exite pas": GoTo Error_MsgBox
'_______________________________________________________________________________________
Dim mSAAJRNTOPX As String, arrSWIBKUBIC_Loop As Integer, arrSWIBKUBIC_K As Integer
Dim xSQL As String, Nb As Long, xSta As String, iCol As Integer

'_______________________________________________________________________________________
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Commentaire                                                                    |<BIC                                |< Intitulé                                                                                                           |||"
fgSelect.Row = 0


xSQL = "select count(*) from " & paramIBM_Library_SAB & ".ZSWIBKU0"

Set rsSab = cnsab.Execute(xSQL)
Nb = rsSab(0) + 10
ReDim arrSWIBKUBIC_8(Nb), arrSWIBKUBIC(Nb), arrSWIBKUBIC_Sta(Nb), arrSWIBKUBIC_6(Nb)

Nb = 0
xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIBKU0" _
     & " left outer join " & paramIBM_Library_SAB & ".ZSWIBIC0 on SWIBICBIC = SWIBKUBIC " _
     & " order by SWIBKUBIC"


Set rsSab = cnsab.Execute(xSQL)
  
Do While Not rsSab.EOF
    If IsNull(rsSab("SWIBICBIC")) Then
        If rsSab("SWIBKUBIC") <> "SOGEFRPPTGV" Then
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect.Col = 1: fgSelect.Text = rsSab("SWIBKUBIC"): fgSelect.CellForeColor = vbRed: fgSelect.CellFontBold = True
            fgSelect.Col = 0: fgSelect.Text = "BIC inconnu, RMA à supprimer dans SAB": fgSelect.CellForeColor = vbRed
        End If
    Else
        Nb = Nb + 1
        arrSWIBKUBIC_8(Nb) = Mid$(rsSab("SWIBKUBIC"), 1, 8)
        arrSWIBKUBIC_6(Nb) = Mid$(rsSab("SWIBKUBIC"), 1, 6)
        arrSWIBKUBIC(Nb) = Trim(rsSab("SWIBKUBIC"))
        arrSWIBKUBIC_Sta(Nb) = ""
    End If
    rsSab.MoveNext

Loop



'________________________________________________________________________________________________

Open wFile For Input As #1

Do Until EOF(1)
    Line Input #1, X
    X = Trim(X)
    If InStr(X, "End of Report") > 0 Then X = "Own BIC:"
    
    If InStr(X, "Own BIC:") > 0 Then
        If blnBic_Correspondent Then
            'Debug.Print wBIC, wReceive, wSend
'________________________________________________________________________________________________
            If wBIC = "BYLADE77" Then
                Debug.Print wBIC
            End If
                wBIC_6 = Mid$(wBIC, 1, 6)
                mSAAJRNTOPX = wBIC  ''X
                arrSWIBKUBIC_K = 0
                'For K = arrSWIBKUBIC_Loop + 1 To Nb
                For K = 1 To Nb
                    If mSAAJRNTOPX = arrSWIBKUBIC_8(K) Then
                        arrSWIBKUBIC_Loop = K
                        arrSWIBKUBIC_K = K
                        Exit For
                    Else
                        If wBIC_6 < arrSWIBKUBIC_6(K) Then Exit For
                    End If
                
                Next K
                
                arrSWIBKUBIC_Sta(arrSWIBKUBIC_K) = wReceive & wSend
'________________________________________________________________________________________________
            
            
        End If
        blnBic_BIARFRPP = False
        blnBic_Correspondent = False: blnReceive = False: blnSend = False
        wBIC = "": wReceive = "?": wSend = "?"
        If InStr(X, "BIARFRPP") > 0 Then blnBic_BIARFRPP = True
    Else
        If blnBic_BIARFRPP Then
            If InStr(X, "Correspondent BIC:") > 0 Then
                blnBic_Correspondent = True
                wBIC = Mid$(X, 20, 8)
            Else
                If blnBic_Correspondent Then
                    If InStr(X, "Authorisation to receive") > 0 Then
                        blnReceive = True: blnSend = False
                    Else
                        If InStr(X, "Authorisation to send") > 0 Then
                            blnReceive = False: blnSend = True
                        Else
                            If InStr(X, "Authorisation status: Enabled") > 0 Then
                                If blnReceive Then wReceive = "R"
                                If blnSend Then wSend = "S"
                            Else
                                If InStr(X, "This authorisation has been revoked") > 0 Then
                                    If blnReceive Then wReceive = "X"
                                    If blnSend Then wSend = "X"
                                Else
                                    If InStr(X, "This authorisation has been rejected") > 0 Then
                                        If blnReceive Then wReceive = "X"
                                        If blnSend Then wSend = "X"
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    

   '________________________________________________________________________________

Loop

Close #1



For K = 2 To Nb
    If Mid$(arrSWIBKUBIC(K), 1, 8) = Mid$(arrSWIBKUBIC(K - 1), 1, 8) Then
        arrSWIBKUBIC_Sta(K) = arrSWIBKUBIC_Sta(K - 1)
    End If

Next K

For K = 1 To Nb
    If arrSWIBKUBIC_Sta(K) <> "RS" Then
    
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZSWIBIC0 " _
             & " where SWIBICBIC like '" & arrSWIBKUBIC(K) & "%'"
        
        Set rsSabX = cnsab.Execute(xSQL)
          
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect.Col = 1: fgSelect.Text = arrSWIBKUBIC(K): fgSelect.CellFontBold = True
        If rsSabX.EOF Then
            'Debug.Print "SAB : BIC Inconnu : "; arrSWIBKUBIC(K)
            fgSelect.CellForeColor = vbBlack
            fgSelect.Col = 2: fgSelect.Text = "": fgSelect.CellForeColor = vbBlack
            fgSelect.Col = 0: fgSelect.Text = "BIC inconnu à supprimer dans SAB": fgSelect.CellForeColor = vbBlack
        Else
            Select Case arrSWIBKUBIC_Sta(K)
                Case "": 'Debug.Print "SAB : à Révoquer : "; arrSWIBKUBIC(K), Trim(rsSabX("SWIBICIN1"))
                    fgSelect.CellForeColor = vbRed
                    fgSelect.Col = 2: fgSelect.Text = Trim(rsSabX("SWIBICIN1")): fgSelect.CellForeColor = vbRed
                    fgSelect.Col = 0: fgSelect.Text = "RMA disparu après reprise BKE, ": fgSelect.CellForeColor = vbRed
                Case "XX": 'Debug.Print "SAB : à Révoquer : "; arrSWIBKUBIC(K), Trim(rsSabX("SWIBICIN1"))
                    fgSelect.CellForeColor = vbRed
                    fgSelect.Col = 2: fgSelect.Text = Trim(rsSabX("SWIBICIN1")): fgSelect.CellForeColor = vbRed
                    fgSelect.Col = 0: fgSelect.Text = "XX : RMA révoqué, BIC à supprimer dans SAB": fgSelect.CellForeColor = vbRed
                
                Case "??": ' Debug.Print "SAB : pas de RMA échangé : "; arrSWIBKUBIC(K), Trim(rsSabX("SWIBICIN1"))
                    fgSelect.CellForeColor = vbMagenta
                    fgSelect.Col = 2: fgSelect.Text = Trim(rsSabX("SWIBICIN1")): fgSelect.CellForeColor = vbMagenta
                    fgSelect.Col = 0: fgSelect.Text = "?? : pas de RMA actif , BIC à supprimer dans SAB": fgSelect.CellForeColor = vbMagenta
                Case Else: ' Debug.Print "SAB : pas de RMA échangé : "; arrSWIBKUBIC(K), Trim(rsSabX("SWIBICIN1"))
                    fgSelect.CellForeColor = vbMagenta
                    fgSelect.Col = 2: fgSelect.Text = Trim(rsSabX("SWIBICIN1")): fgSelect.CellForeColor = vbMagenta
                    fgSelect.Col = 0: fgSelect.Text = arrSWIBKUBIC_Sta(K) & " à analyser": fgSelect.CellForeColor = vbMagenta
            End Select
        End If
    End If
Next K

Set rsSab = Nothing
Set rsSabX = Nothing
fgSelect.Visible = True

'_______________________________________________________________________________________
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdSelect_SQL_SAA_Operator()

On Error GoTo Error_Handler

Dim X As String, wFile As String, wFilex As String
Dim K As Integer, blnOperateurID As Boolean
Dim mProfile As String
currentAction = "cmdSelect_SQL_SAA_Operator"

wFile = "C:\Temp\SAA_Operators_121127.txt"

X = InputBox("par défaut : " _
    & vbCrLf & "     =========================" & vbCrLf & wFile _
    & vbCrLf & "     =========================", "SSwift_Alias : nom du fichier à traiter", wFile)
If Trim(X) = "" Then Exit Sub
wFilex = Trim(X)
'______________________________________________
If wFile <> wFilex Then
    wFile = wFilex
End If

If Dir(wFile) = "" Then V = "Ce fichier n'exite pas": GoTo Error_MsgBox
'_______________________________________________________________________________________
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Opérateur        |<Profil                             |Enabled   |Aproved   |Modifié le                    |" _
                      & " None |BOTC |COBK |DAFI |DCOM |DGAL |DGLI |INFO |ORPA |SCLE |SGCP |SOBF |SOBI |SOPT |STLX"
                      
                      
fgSelect.Row = 0

fgSelect.Col = 1: fgSelect.CellAlignment = 0
For K = 2 To 19
    fgSelect.Col = K: fgSelect.CellAlignment = 2

Next K
fgSelect.Col = 1: fgSelect.CellAlignment = 1

Open wFile For Input As #1

Do Until EOF(1)
    Line Input #1, X
    X = Trim(X)
    
    If InStr(X, "Operator ID") > 0 Then
        blnOperateurID = True
        mProfile = ""
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        K = InStr(X, "=")
        X = Trim(Mid$(X, K + 1, Len(X) - K))
        fgSelect.Col = 0: fgSelect.Text = X
    Else
        If blnOperateurID Then
            
            If InStr(X, "Active profile") > 0 Then
                K = InStr(X, "=")
                X = Trim(Mid$(X, K + 1, Len(X) - K))
                If mProfile = "" Then
                    mProfile = X
                Else
                    mProfile = mProfile & " / " & X
                End If
                fgSelect.Col = 1: fgSelect.Text = mProfile
            Else
                If InStr(X, "Enable status") > 0 Then
                    K = InStr(X, "=")
                    X = Trim(Mid$(X, K + 1, Len(X) - K))
                    fgSelect.Col = 2: fgSelect.Text = X
                Else
                    If InStr(X, "Approval status") > 0 Then
                        K = InStr(X, "=")
                        X = Trim(Mid$(X, K + 1, Len(X) - K))
                        fgSelect.Col = 3: fgSelect.Text = X
                    Else
                         If InStr(X, "Last changed") > 0 Then
                            K = InStr(X, "=")
                            X = Trim(Mid$(X, K + 1, Len(X) - K))
                            fgSelect.Col = 4: fgSelect.Text = X
                        Else
                          If InStr(X, "Assigned unit") > 0 Then
                            K = InStr(X, "=")
                            X = Trim(Mid$(X, K + 1, Len(X) - K))
                            Select Case X
                                Case "None": fgSelect.Col = 5: fgSelect.Text = "•"
                                Case "BOTC": fgSelect.Col = 6: fgSelect.Text = "•"
                                Case "COBK": fgSelect.Col = 7: fgSelect.Text = "•"
                                Case "DAFI": fgSelect.Col = 8: fgSelect.Text = "•"
                                Case "DCOM": fgSelect.Col = 9: fgSelect.Text = "•"
                                Case "DGAL": fgSelect.Col = 10: fgSelect.Text = "•"
                                Case "DGLI": fgSelect.Col = 11: fgSelect.Text = "•"
                                Case "INFO": fgSelect.Col = 12: fgSelect.Text = "•"
                                Case "ORPA": fgSelect.Col = 13: fgSelect.Text = "•"
                                Case "SCLE": fgSelect.Col = 14: fgSelect.Text = "•"
                                Case "SGCP": fgSelect.Col = 15: fgSelect.Text = "•"
                                Case "SOBF": fgSelect.Col = 16: fgSelect.Text = "•"
                                Case "SOBI": fgSelect.Col = 17: fgSelect.Text = "•"
                                Case "SOPT": fgSelect.Col = 18: fgSelect.Text = "•"
                                Case "STLX": fgSelect.Col = 19: fgSelect.Text = "•"
                            End Select
                        End If
                            
                       End If
                   End If
                End If
            End If
            
            
            
        End If
    End If
    
    

Loop

fgSelect.Visible = True

Close #1

'_______________________________________________________________________________________
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


