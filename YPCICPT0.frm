VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmYPCICPT0 
   AutoRedraw      =   -1  'True
   Caption         =   "Plan Comptable"
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
   Icon            =   "YPCICPT0.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10530
   ScaleWidth      =   13575
   Begin VB.ListBox lstErr 
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   0
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9855
      Left            =   0
      TabIndex        =   2
      Top             =   435
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   17383
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
      TabCaption(0)   =   "Plan comptable"
      TabPicture(0)   =   "YPCICPT0.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "."
      TabPicture(1)   =   "YPCICPT0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstW"
      Tab(1).Control(1)=   "fraCompte"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraCompte 
         BackColor       =   &H00C0C0C0&
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
         Left            =   -68265
         TabIndex        =   13
         Top             =   645
         Visible         =   0   'False
         Width           =   6315
         Begin VB.TextBox txtD_CLIENARSD_Fiscal 
            Height          =   330
            Left            =   5310
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   3465
            Width           =   675
         End
         Begin VB.TextBox txtD_COMPTEDEV 
            Height          =   330
            Left            =   5445
            Locked          =   -1  'True
            TabIndex        =   47
            Text            =   "COMPTEDEV"
            Top             =   1995
            Width           =   510
         End
         Begin VB.TextBox txtD_SOLDECEN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3230
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "SOLDECEN"
            Top             =   2010
            Width           =   1860
         End
         Begin VB.TextBox txtD_SOLDEDMO 
            Height          =   330
            Left            =   1935
            Locked          =   -1  'True
            TabIndex        =   45
            Text            =   "SOLDEDMO"
            Top             =   2010
            Width           =   1215
         End
         Begin VB.TextBox txtD_CLIENARSD 
            Height          =   330
            Left            =   4410
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "CLIENARSD"
            Top             =   3465
            Width           =   675
         End
         Begin VB.TextBox txtD_CLIENARES 
            Height          =   330
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   31
            Text            =   "CLIENARES"
            Top             =   2490
            Width           =   960
         End
         Begin VB.TextBox txtD_CLIENANAT 
            Height          =   330
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "CLIENARES"
            Top             =   3500
            Width           =   690
         End
         Begin VB.TextBox txtD_CLIENARA1 
            Height          =   330
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "COMPTEINT"
            Top             =   3000
            Width           =   4050
         End
         Begin VB.TextBox txtD_CLIENASIG 
            Height          =   330
            Left            =   3230
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "CLIENASIG"
            Top             =   2500
            Width           =   1395
         End
         Begin VB.TextBox txtD_CLIENACLI 
            Height          =   330
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "CLIENACLI"
            Top             =   2500
            Width           =   1245
         End
         Begin VB.TextBox txtD_COMPTECLO 
            Height          =   330
            Left            =   4725
            Locked          =   -1  'True
            TabIndex        =   23
            Text            =   "COMPTECLO"
            Top             =   1455
            Width           =   1215
         End
         Begin VB.TextBox txtD_COMPTEOUV 
            Height          =   330
            Left            =   3230
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "COMPTEOUV"
            Top             =   1455
            Width           =   1140
         End
         Begin VB.TextBox txtD_PLANCOPRO 
            Height          =   330
            Left            =   5505
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "PLANCOPRO"
            Top             =   500
            Width           =   495
         End
         Begin VB.TextBox txtD_COMPTEFON 
            Height          =   330
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "COMPTEFON"
            Top             =   1500
            Width           =   405
         End
         Begin VB.TextBox txtD_COMPTEOBL 
            Height          =   330
            Left            =   4425
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "COMPTEOBL"
            Top             =   500
            Width           =   960
         End
         Begin VB.TextBox txtD_COMPTEINT 
            Height          =   330
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "COMPTEINT"
            Top             =   1010
            Width           =   4065
         End
         Begin VB.TextBox txtD_COMPTECOM 
            Height          =   330
            Left            =   1905
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "COMPTECOM"
            Top             =   500
            Width           =   2310
         End
         Begin VB.Label lblD_CLIENARSD 
            BackColor       =   &H00C0C0C0&
            Caption         =   "pays  de résidence, zone Fiscale"
            Height          =   405
            Left            =   2940
            TabIndex        =   49
            Top             =   3420
            Width           =   1425
         End
         Begin VB.Label lblD_SOLDECEN 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Dmvt,solde,devise"
            Height          =   345
            Left            =   180
            TabIndex        =   48
            Top             =   2010
            Width           =   1530
         End
         Begin VB.Label lblD_CLIENANAT 
            BackColor       =   &H00C0C0C0&
            Caption         =   "pays de nationalité"
            Height          =   225
            Left            =   180
            TabIndex        =   30
            Top             =   3555
            Width           =   1695
         End
         Begin VB.Label lblD_CLIENARA1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "intitulé"
            Height          =   345
            Left            =   180
            TabIndex        =   27
            Top             =   3050
            Width           =   1530
         End
         Begin VB.Label lblD_CLIENACLI 
            BackColor       =   &H00C0C0C0&
            Caption         =   "client, sigle, resp"
            Height          =   345
            Left            =   180
            TabIndex        =   24
            Top             =   2550
            Width           =   1530
         End
         Begin VB.Label lblD_COMPTEFON 
            BackColor       =   &H00C0C0C0&
            Caption         =   "code fonct,Dcre,Dclo"
            Height          =   345
            Left            =   180
            TabIndex        =   21
            Top             =   1550
            Width           =   1530
         End
         Begin VB.Label lblD_COMPTEINT 
            BackColor       =   &H00C0C0C0&
            Caption         =   "intitulé"
            Height          =   345
            Left            =   165
            TabIndex        =   16
            Top             =   1065
            Width           =   1530
         End
         Begin VB.Label lblD_COMPTECOM 
            BackColor       =   &H00C0C0C0&
            Caption         =   "compte, PCI, produit"
            Height          =   345
            Left            =   180
            TabIndex        =   14
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
         TabIndex        =   11
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
         Height          =   9450
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   13305
         Begin VB.CommandButton cmdUpdate_Ok 
            BackColor       =   &H00C0C000&
            Caption         =   "Enregistrer"
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
            Left            =   12240
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   765
            Width           =   900
         End
         Begin VB.Frame fraList1 
            BackColor       =   &H00C0C0C0&
            Height          =   3510
            Left            =   120
            TabIndex        =   41
            Top             =   5865
            Visible         =   0   'False
            Width           =   13095
            Begin MSFlexGridLib.MSFlexGrid fgList1 
               Height          =   3180
               Left            =   120
               TabIndex        =   42
               Top             =   195
               Width           =   12885
               _ExtentX        =   22728
               _ExtentY        =   5609
               _Version        =   393216
               Cols            =   15
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16777215
               ForeColor       =   16384
               BackColorFixed  =   8421504
               ForeColorFixed  =   16777215
               BackColorBkg    =   -2147483633
               AllowUserResizing=   3
               FormatString    =   $"YPCICPT0.frx":0342
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
            Height          =   345
            Left            =   9990
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   360
            Width           =   3225
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
            Left            =   10740
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   795
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
            TabIndex        =   6
            Top             =   210
            Visible         =   0   'False
            Width           =   9855
            Begin VB.CheckBox chkSelect_COMPTEFON 
               BackColor       =   &H00F0FFFF&
               Caption         =   "afficher les comptes clos"
               Height          =   270
               Left            =   5010
               TabIndex        =   40
               Top             =   660
               Width           =   2325
            End
            Begin VB.ComboBox cboSelect_DEV 
               Height          =   330
               Left            =   3705
               Sorted          =   -1  'True
               TabIndex        =   36
               Text            =   "dev"
               Top             =   630
               Width           =   945
            End
            Begin VB.Frame fraSelect_Options_1 
               BackColor       =   &H00F0FFFF&
               Height          =   915
               Left            =   105
               TabIndex        =   9
               Top             =   180
               Width           =   3135
               Begin VB.ComboBox cboSelect_PLANCOPRO 
                  Height          =   330
                  Left            =   1695
                  Sorted          =   -1  'True
                  TabIndex        =   38
                  Text            =   "PRO"
                  Top             =   495
                  Width           =   945
               End
               Begin VB.TextBox txtSelect_Where 
                  Height          =   345
                  Left            =   315
                  TabIndex        =   12
                  Top             =   480
                  Width           =   1050
               End
               Begin VB.Label lblSelect_PLANCOPRO 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Produit"
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
                  Left            =   1830
                  TabIndex        =   39
                  Top             =   180
                  Width           =   615
               End
               Begin VB.Label lblSelect_Where 
                  BackColor       =   &H00F0FFFF&
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
                  Left            =   495
                  TabIndex        =   10
                  Top             =   165
                  Width           =   615
               End
            End
            Begin VB.Label lblSelect_DEV 
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
               Left            =   3840
               TabIndex        =   37
               Top             =   195
               Width           =   615
            End
         End
         Begin VB.Frame fraDetail 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   5625
            Left            =   4275
            TabIndex        =   33
            Top             =   1395
            Visible         =   0   'False
            Width           =   8970
            Begin VB.OptionButton optPCICPTAUTO_I 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Ignorer"
               Height          =   210
               Left            =   1875
               TabIndex        =   53
               Top             =   675
               Width           =   855
            End
            Begin VB.OptionButton optPCICPTAUTO_M 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Manuel"
               Height          =   210
               Left            =   870
               TabIndex        =   52
               Top             =   675
               Width           =   900
            End
            Begin VB.OptionButton optPCICPTAUTO_A 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Auto"
               Height          =   210
               Left            =   150
               TabIndex        =   51
               Top             =   675
               Width           =   750
            End
            Begin VB.TextBox txtPCICPTSUFX 
               BackColor       =   &H00FFFFFF&
               Height          =   570
               Left            =   2880
               MultiLine       =   -1  'True
               TabIndex        =   43
               Top             =   240
               Width           =   5910
            End
            Begin VB.TextBox txtPCICPTMETA 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   105
               TabIndex        =   35
               Text            =   "12345689"
               Top             =   255
               Width           =   2670
            End
            Begin VB.TextBox txtPCICPTTXT 
               Height          =   885
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   34
               Top             =   4635
               Width           =   8730
            End
            Begin MSFlexGridLib.MSFlexGrid fgDetail 
               Height          =   3570
               Left            =   90
               TabIndex        =   44
               Top             =   960
               Visible         =   0   'False
               Width           =   8730
               _ExtentX        =   15399
               _ExtentY        =   6297
               _Version        =   393216
               Cols            =   11
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16777215
               ForeColor       =   16384
               BackColorFixed  =   12632064
               ForeColorFixed  =   16777215
               BackColorBkg    =   14737632
               FormatString    =   $"YPCICPT0.frx":044F
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
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   5610
            Left            =   135
            TabIndex        =   5
            Top             =   1410
            Visible         =   0   'False
            Width           =   7365
            _ExtentX        =   12991
            _ExtentY        =   9895
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
            BackColorBkg    =   -2147483637
            AllowBigSelection=   0   'False
            AllowUserResizing=   3
            FormatString    =   $"YPCICPT0.frx":04E7
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
      Picture         =   "YPCICPT0.frx":0571
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
   End
End
Attribute VB_Name = "frmYPCICPT0"
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
Dim SAB_Dossier_Aut As typeAuthorization
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

'______________________________________________________________________

Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long
Dim xAmjMin As String, xAmjMax As String
Dim wDIBM_Min As Long, wDIBM_Max As Long
Dim wDMS_Min As String, wDMS_Max As String
Dim xYBIACPT0 As typeYBIACPT0, newYBIACPT0 As typeYBIACPT0, oldYBIACPT0 As typeYBIACPT0
Dim arrYBIACPT0() As typeYBIACPT0, arrYBIACPT0_Nb As Long, arrYBIACPT0_Max As Long, arrYBIACPT0_Index As Long

Dim xYPCICPT0 As typeYPCICPT0, newYPCICPT0 As typeYPCICPT0, oldYPCICPT0 As typeYPCICPT0
Dim arrYPCICPT0() As typeYPCICPT0, arrYPCICPT0_Nb As Long, arrYPCICPT0_Max As Long, arrYPCICPT0_Index As Long
Dim arrYPCICPT0_K As Long

Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean



Dim fgList1_FormatString As String, fgList1_K As Integer
Dim fgList1_RowDisplay As Integer, fgList1_RowClick As Integer, fgList1_ColClick As Integer
Dim fgList1_ColorClick As Long, fgList1_ColorDisplay As Long
Dim fgList1_Sort1 As Integer, fgList1_Sort2 As Integer
Dim fgList1_SortAD As Integer, fgList1_Sort1_Old As Integer
Dim fgList1_arrIndex As Integer
Dim blnfgList1_DisplayLine As Boolean

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim arrDev() As String, arrDev_Nb As Integer, arrDev_RowT() As Long, arrDev_Cours() As Double
Dim arrDev_Num() As String, arrDev_EUR As Integer

Dim mXls1_Row As Long, mXls1_Col As Long, mXls2_Row As Long, mXls2_Col As Long, mXls2_Row_Cli As Long
Dim wMTD As Currency, wPIE As Long, wECR As Long


Dim wFilex As String, wFile As String
Dim xZCOMPTE0 As typeZCOMPTE0, oldZCOMPTE0 As typeZCOMPTE0
Dim xZRELEVE0 As typeZRELEVE0, oldZRELEVE0 As typeZRELEVE0
Dim mMsgBox_Err As String

Dim xZPLAN0 As typeZPLAN0, newZPLAN0 As typeZPLAN0, oldZPLAN0 As typeZPLAN0
Dim arrZPLAN0() As typeZPLAN0, arrZPLAN0_Nb As Long, arrZPLAN0_Max As Long, arrZPLAN0_Index As Long
Dim arrZPLAN0_K As Long
Dim arrZPLAN0_Lnk() As Long, arrYPCILNK0_Lnk() As Long

Dim arrPays() As typePays, arrPays_Nb As Integer
Dim arrZSOLDE0() As typeZSOLDE0, arrZSOLDE0_Nb As Long, xZSOLDE0 As typeZSOLDE0
Dim arrZCOMPTE0() As typeZCOMPTE0, arrZCOMPTE0_Nb As Long

Dim arrCLIEANRES() As String, arrCLIEANRES_Nb As Integer
Public Sub cmdSelect_SQL_Exportation_PCICPTBASE(blnCompte As Boolean, blnCompte_All As Boolean)
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wAMJMin As String, WAMJMax As String
Dim X As String, K As Long
Dim blnCALCS As Boolean
On Error GoTo Error_Handler
'===================================================================================
If blnAuto Then
    X = paramServer("\\CPT_Archive\")
Else
    X = ""
End If
If X = "" Then X = "C:\Temp\"
If mId$(X, Len(X), 1) <> "\" Then X = X & "\"

blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

If Not blnCompte Then
    wFile = X & Trim("CPT plan comptable au " & dateImp_Amj(YBIATAB0_DATE_CPT_J) & ".xlsx")
Else
    If blnCompte_All Then
        wFile = X & Trim("CPT plan comptable détaillé au " & dateImp_Amj(YBIATAB0_DATE_CPT_J) & ".xlsx")
    Else
        wFile = X & Trim("CPT plan comptable anomalies au " & dateImp_Amj(YBIATAB0_DATE_CPT_J) & ".xlsx")
    End If
End If

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

'_________________________________________

If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile


Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "YPCICPT0"
    .Subject = ""
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "PCI- " & dateImp10(YBIATAB0_DATE_CPT_J)

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
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14CPT : Plan comptable en date du " & dateImp10_S(YBIATAB0_DATE_CPT_J) _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$L1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

wsExcel.Columns(1).ColumnWidth = 8: wsExcel.Cells(1, 1) = "PCI*": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(2).ColumnWidth = 8: wsExcel.Cells(1, 2) = "Produit"
wsExcel.Columns(3).ColumnWidth = 25: wsExcel.Cells(1, 3) = "Structure PCI / compte": wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(4).ColumnWidth = 32: wsExcel.Cells(1, 4) = "Suffixe PCI / solde + devise du compte": wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(5).ColumnWidth = 6: wsExcel.Cells(1, 5) = "Sécurité"
wsExcel.Columns(6).ColumnWidth = 6: wsExcel.Cells(1, 6) = "Blocage"
wsExcel.Columns(7).ColumnWidth = 6: wsExcel.Cells(1, 7) = "Sens"
wsExcel.Columns(8).ColumnWidth = 6: wsExcel.Cells(1, 8) = "G.dépassement"
wsExcel.Columns(9).ColumnWidth = 6: wsExcel.Cells(1, 9) = "Tiers obl"
wsExcel.Columns(10).ColumnWidth = 6: wsExcel.Cells(1, 10) = "Clientèle"
wsExcel.Columns(11).ColumnWidth = 6: wsExcel.Cells(1, 11) = "Longueur": wsExcel.Columns(8).NumberFormat = "##0"
wsExcel.Columns(11).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(12).ColumnWidth = 32: wsExcel.Cells(1, 12) = "Intitulé": wsExcel.Columns(12).HorizontalAlignment = Excel.xlHAlignLeft
'mXls1_Col = 12

'If blnCompte Then
    wsExcel.Columns(13).ColumnWidth = 25: wsExcel.Cells(1, 13) = "Commentaire": wsExcel.Columns(13).HorizontalAlignment = Excel.xlHAlignLeft
    mXls1_Col = 13
    wsExcel.PageSetup.Zoom = 70
    wsExcel.PageSetup.PrintTitleRows = "$A1:$M1"
'End If

For K = 1 To mXls1_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next
'__________________________________________________________________________________

cmdSelect_SQL_Exportation_PCICPTBASE_Detail blnCompte, blnCompte_All

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

Public Sub cmdSelect_SQL_Exportation_PCICPTBASE_Detail(blnCompte As Boolean, blnCompte_All As Boolean)
On Error GoTo Error_Handler
Dim X As String, xWhere As String
Dim wRow As Long, wCol As Long, mRow_PCI As Long
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, kLen As Integer
Dim K5 As Integer, K6 As Integer, K7 As Integer, K8 As Integer
Dim wAmj As String, mSOLDEDMO_Min As String
Dim wColor As Long
Dim xPCICPTMETA_Control As String, blnPCICPTMETA_Control As Boolean
mSOLDEDMO_Min = dateElp("MoisAdd", -18, YBIATAB0_DATE_CPT_J)
If blnCompte Then cmdSelect_SQL_Exportation_PCICPTBASE_Detail_Compte


xWhere = " where PLANETABL = 1 and PLANPLAN = 1 and PLANCOOBL = PCICPTLNK"
X = Trim(txtSelect_Where)
If X <> "" Then xWhere = xWhere & " and PCICPTBASE like '" & X & "%'"

X = Trim(cboSelect_PLANCOPRO)
If X <> "" Then xWhere = xWhere & " and PLANCOPRO = '" & X & "'"


X = "select * from " & paramIBM_Library_SABSPE & ".YPCICPT0, " & paramIBM_Library_SAB & ".ZPLAN0 " _
       & xWhere _
       & " order by  PCICPTBASE"
Set rsSab = cnsab.Execute(X)

wRow = 1
Do While Not rsSab.EOF
    V = rsYPCICPT0_GetBuffer(rsSab, xYPCICPT0)
    V = rsZPLAN0_GetBuffer(rsSab, xZPLAN0)
    wRow = wRow + 1
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "> " & xZPLAN0.PLANCOPRO & " " & xYPCICPT0.PCICPTBASE): DoEvents
    If blnCompte Then
        wsExcel.Cells(wRow, 1).Font.Color = RGB(0, 64, 64): wsExcel.Cells(wRow, 1).Font.Bold = True
        wsExcel.Cells(wRow, 2).Font.Color = RGB(0, 64, 64): wsExcel.Cells(wRow, 2).Font.Bold = True
        wsExcel.Cells(wRow, 12).Font.Color = RGB(0, 64, 64)
        For K = 1 To mXls1_Col:    wsExcel.Cells(wRow, K).Interior.Color = mColor_Y0: Next
    End If
    If InStr(xYPCICPT0.PCICPTMETA, "?") Then wsExcel.Cells(wRow, 2).Interior.Color = mColor_W1
    
    xYPCICPT0.PCICPTBASE = Trim(xYPCICPT0.PCICPTBASE)
    If xYPCICPT0.PCICPTLEN = 5 Then
        X = xYPCICPT0.PCICPTBASE & "*"
    Else
        X = xYPCICPT0.PCICPTBASE
    End If
    Select Case xYPCICPT0.PCICPTAUTO
        Case "M": X = X & " M": wsExcel.Cells(wRow, 1).Interior.Color = RGB(0, 255, 0)
        Case "I": X = X & " I": wsExcel.Cells(wRow, 1).Interior.Color = RGB(0, 255, 0)
    End Select
    
    wsExcel.Cells(wRow, 1) = X
    wsExcel.Cells(wRow, 2) = xZPLAN0.PLANCOPRO
    wsExcel.Cells(wRow, 4) = Replace(Trim(xYPCICPT0.PCICPTSUFX), " ", "   "): wsExcel.Cells(wRow, 4).Font.Color = RGB(0, 64, 64)
    X = PCICPTMETA_Display(xYPCICPT0.PCICPTMETA)
    wsExcel.Cells(wRow, 3) = X
    wsExcel.Cells(wRow, 3).Font.Color = RGB(0, 96, 96): wsExcel.Cells(wRow, 3).Font.Bold = True: wsExcel.Cells(wRow, 3).Font.Size = 8
    wsExcel.Cells(wRow, 5) = xZPLAN0.PLANCLASS
    
    
    Select Case xZPLAN0.PLANFONCT
        Case 0
        Case 1: wsExcel.Cells(wRow, 6) = "DB": wsExcel.Cells(wRow, 6).Interior.Color = mColor_Y1
        Case 2: wsExcel.Cells(wRow, 6) = "CR": wsExcel.Cells(wRow, 6).Interior.Color = mColor_Y1
        Case 3: wsExcel.Cells(wRow, 6) = "DB-CR": wsExcel.Cells(wRow, 6).Interior.Color = mColor_Y1
        Case 4: wsExcel.Cells(wRow, 6) = "clos": wsExcel.Cells(wRow, 6).Interior.Color = mColor_Y1
        Case Else: wsExcel.Cells(wRow, 6) = xZPLAN0.PLANFONCT: wsExcel.Cells(wRow, 6).Interior.Color = mColor_W1
    End Select
    
    Select Case xZPLAN0.PLANSESOL
        Case " "
        Case "D": wsExcel.Cells(wRow, 7) = "DB": wsExcel.Cells(wRow, 7).Interior.Color = mColor_Y1
        Case "C": wsExcel.Cells(wRow, 7) = "CR": wsExcel.Cells(wRow, 7).Interior.Color = mColor_Y1
        Case Else: wsExcel.Cells(wRow, 7) = xZPLAN0.PLANSESOL: wsExcel.Cells(wRow, 7).Interior.Color = mColor_W1
    End Select
    
    Select Case xZPLAN0.PLANGEDEP
    Case "N"
        Case "O": wsExcel.Cells(wRow, 8) = "oui": wsExcel.Cells(wRow, 8).Interior.Color = mColor_Y1
        Case Else: wsExcel.Cells(wRow, 8) = xZPLAN0.PLANGEDEP: wsExcel.Cells(wRow, 8).Interior.Color = mColor_W1
    End Select
    
    Select Case xZPLAN0.PLANTIERS
    Case "N"
        Case "O": wsExcel.Cells(wRow, 9) = "oui": wsExcel.Cells(wRow, 9).Interior.Color = mColor_Y1
        Case Else: wsExcel.Cells(wRow, 9) = xZPLAN0.PLANTIERS: wsExcel.Cells(wRow, 9).Interior.Color = mColor_W1
    End Select
    
    Select Case xZPLAN0.PLANFICOB
    Case "N"
        Case "O": wsExcel.Cells(wRow, 10) = "oui": wsExcel.Cells(wRow, 10).Interior.Color = mColor_Y1
        Case Else: wsExcel.Cells(wRow, 10) = xZPLAN0.PLANFICOB: wsExcel.Cells(wRow, 10).Interior.Color = mColor_W1
    End Select
    
    wsExcel.Cells(wRow, 11) = xZPLAN0.PLANCARAC
    wsExcel.Cells(wRow, 12) = xZPLAN0.PLANINTIT
    wsExcel.Cells(wRow, 13) = Trim(xYPCICPT0.PCICPTTXT)
'===================================================================================================
    If blnCompte Then
        mRow_PCI = wRow
        For K = 1 To arrZCOMPTE0_Nb
            If xYPCICPT0.PCICPTBASE = mId$(arrZCOMPTE0(K).COMPTEOBL, 1, xYPCICPT0.PCICPTLEN) Then
                xZCOMPTE0 = arrZCOMPTE0(K)
                
                blnPCICPTMETA_Control = PCICPTMETA_Control(xPCICPTMETA_Control)
                If xYPCICPT0.PCICPTAUTO = "I" Then blnPCICPTMETA_Control = True
                
                If blnCompte_All Or Not blnPCICPTMETA_Control Then
                
                    wRow = wRow + 1
                    wAmj = arrZSOLDE0(K).SOLDEDMO + 19000000
                    wsExcel.Cells(wRow, 2) = dateImp10(wAmj)
                    wsExcel.Cells(wRow, 3) = xZCOMPTE0.COMPTECOM
                    If arrZSOLDE0(K).SOLDECEN = 0 Then
                        wsExcel.Cells(wRow, 4) = xZCOMPTE0.COMPTEDEV & " "
                        If wAmj < mSOLDEDMO_Min Then
                            Select Case xZPLAN0.PLANCOPRO
                                Case "ICC", "CHA", "PRO": wColor = RGB(230, 230, 230)
                                Case "CAV", "LOR", "DAT": wColor = mColor_W0
                                Case Else: wColor = RGB(200, 200, 200)
                            End Select
                            wsExcel.Cells(wRow, 2).Interior.Color = wColor
                            wsExcel.Cells(wRow, 3).Interior.Color = wColor
                            wsExcel.Cells(wRow, 12).Interior.Color = wColor
                        End If
                    Else
                        wsExcel.Cells(wRow, 4) = Format$(-arrZSOLDE0(K).SOLDECEN, "### ### ### ##0.00") & "  " & xZCOMPTE0.COMPTEDEV & " "
                    End If
                    wsExcel.Cells(wRow, 4).HorizontalAlignment = Excel.xlHAlignRight
                    wsExcel.Cells(wRow, 12) = xZCOMPTE0.COMPTEINT
                    
                    If Not blnPCICPTMETA_Control Then
                        wsExcel.Cells(wRow, 3).Interior.Color = mColor_W0
                        wsExcel.Cells(wRow, 13) = xPCICPTMETA_Control
                        wsExcel.Cells(wRow, 13).Interior.Color = mColor_W0
                    End If
    
    '__________________________________________________________________________________________________
                    wsExcel.Cells(wRow, 5) = xZCOMPTE0.COMPTECLA
                    If xZCOMPTE0.COMPTECLA < xZPLAN0.PLANCLASS Then wsExcel.Cells(wRow, 5).Interior.Color = mColor_W1
    '__________________________________________________________________________________________________
                    Select Case xZCOMPTE0.COMPTEFON
                        Case 0
                        Case 1: wsExcel.Cells(wRow, 6) = "DB": wsExcel.Cells(wRow, 6).Interior.Color = mColor_Y1
                        Case 2: wsExcel.Cells(wRow, 6) = "CR": wsExcel.Cells(wRow, 6).Interior.Color = mColor_Y1
                        Case 3: wsExcel.Cells(wRow, 6) = "DB-CR": wsExcel.Cells(wRow, 6).Interior.Color = mColor_Y1
                        Case 4: wsExcel.Cells(wRow, 6) = "clos": wsExcel.Cells(wRow, 6).Interior.Color = mColor_Y1
                        Case Else: wsExcel.Cells(wRow, 6) = xZPLAN0.PLANFONCT: wsExcel.Cells(wRow, 6).Interior.Color = mColor_W1
                    End Select
                    If xZPLAN0.PLANFONCT <> 0 Then
                        If xZCOMPTE0.COMPTEFON <> "4" Then
                            If xZCOMPTE0.COMPTEFON <> xZPLAN0.PLANFONCT Then wsExcel.Cells(wRow, 6) = mColor_W1
                        End If
                    End If
     '__________________________________________________________________________________________________
                     Select Case xZCOMPTE0.COMPTESEN
                        Case " "
                        Case "D": wsExcel.Cells(wRow, 7) = "DB": wsExcel.Cells(wRow, 7).Interior.Color = mColor_Y1
                        Case "C": wsExcel.Cells(wRow, 7) = "CR": wsExcel.Cells(wRow, 7).Interior.Color = mColor_Y1
                        Case Else: wsExcel.Cells(wRow, 7) = xZPLAN0.PLANSESOL: wsExcel.Cells(wRow, 7).Interior.Color = mColor_W1
                    End Select
                    If Trim(xZPLAN0.PLANSESOL) <> "" Then
                        If xZCOMPTE0.COMPTESEN <> xZPLAN0.PLANSESOL Then wsExcel.Cells(wRow, 7).Interior.Color = mColor_W1
                    End If
     '__________________________________________________________________________________________________
                     If xZCOMPTE0.COMPTESUC = "O" Then wsExcel.Cells(wRow, 8) = "Succession": wsExcel.Cells(wRow, 8).Interior.Color = mColor_Y1: wsExcel.Cells(wRow, 8).Font.Size = 6
     '__________________________________________________________________________________________________
                      Select Case xZCOMPTE0.COMPTELOR
                        Case " "
                        Case "N": wsExcel.Cells(wRow, 9) = "Nostro": wsExcel.Cells(wRow, 9).Interior.Color = mColor_Y1
                        Case "L": wsExcel.Cells(wRow, 9) = "Loro": wsExcel.Cells(wRow, 9).Interior.Color = mColor_Y1
                        Case Else: wsExcel.Cells(wRow, 9) = xZCOMPTE0.COMPTELOR: wsExcel.Cells(wRow, 9).Interior.Color = mColor_W1
                    End Select
                End If
'__________________________________________________________________________________________________

           End If
        Next K
        
        If Not blnCompte_All Then
            If wRow = mRow_PCI Then wRow = wRow - 1
        End If

    End If
'===================================================================================================
    rsSab.MoveNext
Loop

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name


End Sub

Public Sub cmdSelect_SQL_Surveillance_ZPLAN0()
Dim X As String, mPLANCOPRO As String, mPLANFONCT As String
Dim K As Long, Nb_Lu As Long, Nb_Err As Long

 On Error GoTo Error_Handler

'==============================================================================
Call lstErr_AddItem(lstErr, cmdContext, "> contrôle : ZPLAN0"): DoEvents

Set wsExcel = wbExcel.Sheets(2)

Nb_Err = 0: Nb_Lu = 0

For arrZPLAN0_Index = 1 To arrZPLAN0_Nb
    xZPLAN0 = arrZPLAN0(arrZPLAN0_Index)
    Nb_Lu = Nb_Lu + 1
    
    If xZPLAN0.PLANETABL <> 1 Then
        mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
        wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0"
        wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
        wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
        wsExcel.Cells(mXls2_Row, 5) = "code établissement : " & xZPLAN0.PLANETABL
    End If
     
    If xZPLAN0.PLANPLAN <> 1 Then
        mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
        wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0"
        wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
        wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
        wsExcel.Cells(mXls2_Row, 5) = "code agence : " & xZPLAN0.PLANPLAN
    End If
   
     
    If Len(Trim(xZPLAN0.PLANCOOBL)) <> 6 Then
        mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
        wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0"
        wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
        wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
        wsExcel.Cells(mXls2_Row, 5) = "longueur de la rubrique comptable <> 6"
    End If
   
     
    If Trim(xZPLAN0.PLANINTIT) = "" Then
        mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
        wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0"
        wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
        wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
        wsExcel.Cells(mXls2_Row, 5) = "intitulé non défini"
    End If
    
    If mPLANCOPRO <> xZPLAN0.PLANCOPRO Then
        X = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 " _
       & " where BASTABETA = 1 and BASTABNUM = 14 and BASTABARG = '" & xZPLAN0.PLANCOPRO & "'"
        Set rsSabX = cnsab.Execute(X)
        
        If Not rsSabX.EOF Then
            mPLANCOPRO = xZPLAN0.PLANCOPRO
        Else
            mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
            wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0"
            wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
            wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
            wsExcel.Cells(mXls2_Row, 5) = "code produit inconnu (table 14) : " & xZPLAN0.PLANCOPRO
        End If
    End If
     
    If xZPLAN0.PLANCLASS = 0 Then
        mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
        wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0"
        wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
        wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
        wsExcel.Cells(mXls2_Row, 5) = "classe de sécurité non définie"
    End If
    
    If mPLANCOPRO <> xZPLAN0.PLANFONCT Then
        X = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 " _
       & " where BASTABETA = 1 and BASTABNUM = 15 and BASTABARG = '" & xZPLAN0.PLANFONCT & "'"
        Set rsSabX = cnsab.Execute(X)
        
        If Not rsSabX.EOF Then
            mPLANFONCT = xZPLAN0.PLANFONCT
        Else
            mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
            wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0"
            wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
            wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
            wsExcel.Cells(mXls2_Row, 5) = "code fonctionnement inconnu (table 15) : " & xZPLAN0.PLANFONCT
        End If
    End If
     
    If xZPLAN0.PLANSESOL <> " " And xZPLAN0.PLANSESOL <> "C" And xZPLAN0.PLANSESOL <> "D" Then
        mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
        wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0"
        wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
        wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
        wsExcel.Cells(mXls2_Row, 5) = "code sens du solde inconnu { |C|D} non défini : " & xZPLAN0.PLANSESOL
    End If
     
    If xZPLAN0.PLANGEDEP <> "O" And xZPLAN0.PLANGEDEP <> "N" Then
        mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
        wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0"
        wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
        wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
        wsExcel.Cells(mXls2_Row, 5) = "code gestion du dépassement inconnu {O|N} non défini : " & xZPLAN0.PLANGEDEP
    End If
     
    If xZPLAN0.PLANTIERS <> "O" And xZPLAN0.PLANTIERS <> "N" Then
        mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
        wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0"
        wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
        wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
        wsExcel.Cells(mXls2_Row, 5) = "code saisie d'un titulaire inconnu {O|N} non défini : " & xZPLAN0.PLANTIERS
    End If
     
    If xZPLAN0.PLANFICOB <> "O" And xZPLAN0.PLANFICOB <> "N" Then
        mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
        wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0"
        wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
        wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
        wsExcel.Cells(mXls2_Row, 5) = "code compte de la clientèle inconnu {O|N} non défini : " & xZPLAN0.PLANFICOB
    End If
     
    If xZPLAN0.PLANCARAC < 3 Or xZPLAN0.PLANCARAC > 20 Then
        mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
        wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0"
        wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
        wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
        wsExcel.Cells(mXls2_Row, 5) = "longueur du compte hors limites (3 à 20) : " & xZPLAN0.PLANCARAC
    End If
    
Next arrZPLAN0_Index

'==================================================================================================
Set wsExcel = wbExcel.Sheets(1)
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "ZPLAN0"

If Nb_Err = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* Contrôle de cohérence du fichier ZPLAN0 / Référentiel : Ok (" & arrZPLAN0_Nb & " enregistrements lus)."
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* Contrôle de cohérence du fichier ZPLAN0  / Référentiel : " & Nb_Err & " anomalies (" & arrZPLAN0_Nb & " enregistrements lus)."
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If
'==================================================================================================

Exit Sub

Error_Handler:
If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub cmdSelect_SQL_Surveillance_YBIACPT0()
Dim X As String
Dim K As Long, Nb_Lu As Long, Nb_Err As Long
Dim Nb_Err_COMPTEOBL As Long, Nb_Err_PLANFONCT As Long, Nb_Err_PLANSESOL As Long
Dim Nb_Err_PLANCLASS As Long, Nb_Err_COMPTESEN As Long, Nb_Err_SOLDECEN As Long, Nb_Err_999999 As Long
Dim Nb_Err_CLIENARSD As Long, Nb_Err_COMPTEDEV As Long, Nb_Err_Balance As Long
Dim Nb_Err_CLIENACLI As Long, Nb_Err_CLIENARES As Long
Dim blnOk As Boolean, blnCLIENARSD As Boolean
Dim iFiscal_Meta As Integer, mFiscal_Code As String

 On Error GoTo Error_Handler

'==============================================================================
Call lstErr_AddItem(lstErr, cmdContext, "> contrôle : YBIACPT0"): DoEvents
Set wsExcel = wbExcel.Sheets(2)

ReDim arrB_DB(arrDev_Nb + 1), arrB_CR(arrDev_Nb + 1)
ReDim arrHB_DB(arrDev_Nb + 1), arrHB_CR(arrDev_Nb + 1)
Dim arrDev_K As Integer, mDev_Code As String, mDev_Num As String, iDev_Meta As Integer
Dim blnHB As Boolean

Nb_Err = 0: Nb_Lu = 0
Nb_Err_COMPTEOBL = 0: Nb_Err_PLANFONCT = 0: Nb_Err_PLANSESOL = 0
Nb_Err_PLANCLASS = 0: Nb_Err_COMPTESEN = 0: Nb_Err_SOLDECEN = 0: Nb_Err_999999 = 0
Nb_Err_CLIENARSD = 0
Nb_Err_CLIENACLI = 0: Nb_Err_CLIENARES = 0
arrZPLAN0_Index = 0: arrYPCICPT0_Index = 0
iFiscal_Meta = 0

Call rsZPLAN0_Init(oldZPLAN0)
Call rsYPCICPT0_Init(oldYPCICPT0)
'=======================================================================================================
X = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
   & " where COMPTEFON <> '4' " _
       & " order by COMPTEOBL,COMPTEDEV,COMPTECOM"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    V = rsYBIACPT0_GetBuffer(rsSab, xYBIACPT0)
    Nb_Lu = Nb_Lu + 1
    
    If Trim(xYBIACPT0.COMPTEOBL) = "" Then
        blnOk = False
        'mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
        'wsExcel.Cells(mXls2_Row, 1) = "YBIACPT0"
        'wsExcel.Cells(mXls2_Row, 2) = xYBIACPT0.COMPTECOM
        'wsExcel.Cells(mXls2_Row, 3) = xYBIACPT0.COMPTEINT
        'wsExcel.Cells(mXls2_Row, 5) = "Rubrique comptable manquante : " & xYBIACPT0.COMPTEOBL
        'wsExcel.Cells(mXls2_Row, 4).Interior.Color = mColor_W0
    Else
        If xYBIACPT0.COMPTEOBL <> oldZPLAN0.PLANCOOBL Then
            blnOk = False
            blnHB = IIf(mId$(xYBIACPT0.COMPTEOBL, 1, 1) = 9, True, False)
            For arrZPLAN0_Index = 1 To arrZPLAN0_Nb
                If xYBIACPT0.COMPTEOBL = arrZPLAN0(arrZPLAN0_Index).PLANCOOBL Then
                   oldZPLAN0 = arrZPLAN0(arrZPLAN0_Index)
                   oldYPCICPT0 = arrYPCICPT0(arrZPLAN0_Lnk(arrZPLAN0_Index))
                   iFiscal_Meta = InStr(oldYPCICPT0.PCICPTMETA, "#")
                   iDev_Meta = InStr(oldYPCICPT0.PCICPTMETA, "$$$")
                    blnOk = True

                    Exit For
                End If
            Next arrZPLAN0_Index
        End If
    End If
    
    
    If Not blnOk Then
        Set wsExcel = wbExcel.Sheets(1)
        mXls1_Row = mXls1_Row + 1: Nb_Err_COMPTEOBL = Nb_Err_COMPTEOBL + 1
        wsExcel.Cells(mXls1_Row, 1) = "YBIACPT0"
        wsExcel.Cells(mXls1_Row, 2) = xYBIACPT0.COMPTECOM
        wsExcel.Cells(mXls1_Row, 3) = xYBIACPT0.COMPTEINT
        wsExcel.Cells(mXls1_Row, 4) = -xYBIACPT0.SOLDECEN
        wsExcel.Cells(mXls1_Row, 5) = xYBIACPT0.COMPTEDEV & " - Rubrique comptable inconnue : " & xYBIACPT0.COMPTEOBL
        For K = 1 To 5
            wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W1
            wsExcel.Cells(mXls1_Row, K).Font.Color = mColor_Y1
            wsExcel.Cells(mXls1_Row, K).Font.Bold = True

        Next K
        Set wsExcel = wbExcel.Sheets(2)
    Else
'___________________________________________________________________________________________
        
        If xYBIACPT0.COMPTEDEV <> mDev_Code Then
            mDev_Code = "": mDev_Num = ""
            For arrDev_K = 1 To arrDev_Nb
                If xYBIACPT0.COMPTEDEV = arrDev(arrDev_K) Then
                   mDev_Code = arrDev(arrDev_K)
                   mDev_Num = arrDev_Num(arrDev_K)
                    Exit For
                End If
            Next arrDev_K
        End If
        
        If mDev_Code = "" Then
            mXls2_Row = mXls2_Row + 1: Nb_Err_COMPTEDEV = Nb_Err_COMPTEDEV + 1
            wsExcel.Cells(mXls2_Row, 1) = "YBIACPT0"
            wsExcel.Cells(mXls2_Row, 2) = xYBIACPT0.COMPTECOM
            wsExcel.Cells(mXls2_Row, 3) = xYBIACPT0.COMPTEINT
            wsExcel.Cells(mXls2_Row, 4) = -xYBIACPT0.SOLDECEN
            wsExcel.Cells(mXls2_Row, 5) = "la devise du compte est inconnue: " & mDev_Code & " (rubrique comptable " & xYBIACPT0.COMPTEOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
            
            arrDev_K = 0: mDev_Code = "?#"
            
            For K = 1 To 5
                wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_Y1
                wsExcel.Cells(mXls2_Row, K).Font.Color = vbRed
            Next K
        End If
          
        If iDev_Meta > 0 Then
            If mDev_Code <> mId$(xYBIACPT0.COMPTECOM, iDev_Meta, 3) _
            And mDev_Num <> mId$(xYBIACPT0.COMPTECOM, iDev_Meta, 3) Then
                If mId$(xYBIACPT0.COMPTECOM, iDev_Meta, 3) <> "CVD" Then
                    mXls2_Row = mXls2_Row + 1: Nb_Err_COMPTEDEV = Nb_Err_COMPTEDEV + 1
                    wsExcel.Cells(mXls2_Row, 1) = "YBIACPT0"
                    wsExcel.Cells(mXls2_Row, 2) = xYBIACPT0.COMPTECOM
                    wsExcel.Cells(mXls2_Row, 3) = xYBIACPT0.COMPTEINT
                    wsExcel.Cells(mXls2_Row, 4) = -xYBIACPT0.SOLDECEN
                    wsExcel.Cells(mXls2_Row, 5) = "la devise de l'identifaint du compte est # de la devise du compte : " & mDev_Code & " (rubrique comptable " & xYBIACPT0.COMPTEOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
                    For K = 1 To 5
                        wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_Y1
                        wsExcel.Cells(mXls2_Row, K).Font.Color = vbRed
                    Next K
                End If
            End If
        End If
'___________________________________________________________________________________________
        If blnHB Then
            If xYBIACPT0.SOLDECEN < 0 Then
                arrHB_CR(arrDev_K) = arrHB_CR(arrDev_K) - xYBIACPT0.SOLDECEN
            Else
                arrHB_DB(arrDev_K) = arrHB_DB(arrDev_K) - xYBIACPT0.SOLDECEN
            End If
        Else
            If xYBIACPT0.SOLDECEN < 0 Then
                arrB_CR(arrDev_K) = arrB_CR(arrDev_K) - xYBIACPT0.SOLDECEN
            Else
                arrB_DB(arrDev_K) = arrB_DB(arrDev_K) - xYBIACPT0.SOLDECEN
            End If
        End If

'_________  __________________________________________________________________________________
        If iFiscal_Meta > 0 And Trim(xYBIACPT0.CLIENARSD) <> "" Then
            mFiscal_Code = fraCompte_Fiscal(Trim(xYBIACPT0.CLIENARSD))
            If mFiscal_Code <> mId$(xYBIACPT0.COMPTECOM, iFiscal_Meta, 1) Then
                mXls2_Row = mXls2_Row + 1: Nb_Err_CLIENARSD = Nb_Err_CLIENARSD + 1
                wsExcel.Cells(mXls2_Row, 1) = "YBIACPT0"
                wsExcel.Cells(mXls2_Row, 2) = xYBIACPT0.COMPTECOM
                wsExcel.Cells(mXls2_Row, 3) = xYBIACPT0.COMPTEINT
                wsExcel.Cells(mXls2_Row, 4) = -xYBIACPT0.SOLDECEN
                wsExcel.Cells(mXls2_Row, 5) = "Zone fiscale du PCI non conforme avec le pays du client : " & xYBIACPT0.CLIENARSD & " => " & mFiscal_Code & " (rubrique comptable " & xYBIACPT0.COMPTEOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
                For K = 1 To 5
                    wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_Y1
                    wsExcel.Cells(mXls2_Row, K).Font.Color = vbRed
                Next K
            End If
        End If
'___________________________________________________________________________________________
        If xYBIACPT0.COMPTECLA < oldZPLAN0.PLANCLASS Then
            mXls2_Row = mXls2_Row + 1: Nb_Err_PLANCLASS = Nb_Err_PLANCLASS + 1
            wsExcel.Cells(mXls2_Row, 1) = "YBIACPT0"
            wsExcel.Cells(mXls2_Row, 2) = xYBIACPT0.COMPTECOM
            wsExcel.Cells(mXls2_Row, 3) = xYBIACPT0.COMPTEINT
            wsExcel.Cells(mXls2_Row, 4) = -xYBIACPT0.SOLDECEN
            wsExcel.Cells(mXls2_Row, 5) = "code 'classe de sécurité' du compte : " & xYBIACPT0.COMPTECLA & " < " & oldZPLAN0.PLANCLASS & " (rubrique comptable " & xYBIACPT0.COMPTEOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
        End If
'___________________________________________________________________________________________
        If xYBIACPT0.COMPTEFON <> oldZPLAN0.PLANFONCT Then
            If oldZPLAN0.PLANFONCT <> "0" Then
                mXls2_Row = mXls2_Row + 1: Nb_Err_PLANFONCT = Nb_Err_PLANFONCT + 1
                wsExcel.Cells(mXls2_Row, 1) = "YBIACPT0"
                wsExcel.Cells(mXls2_Row, 2) = xYBIACPT0.COMPTECOM
                wsExcel.Cells(mXls2_Row, 3) = xYBIACPT0.COMPTEINT
                wsExcel.Cells(mXls2_Row, 4) = -xYBIACPT0.SOLDECEN
                wsExcel.Cells(mXls2_Row, 5) = "code 'fonctionnement' du compte : " & xYBIACPT0.COMPTEFON & " # " & oldZPLAN0.PLANFONCT & " (rubrique comptable " & xYBIACPT0.COMPTEOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
                For K = 1 To 5: wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_Y0: Next K
            End If
        End If
'___________________________________________________________________________________________
        If xYBIACPT0.COMPTESEN <> oldZPLAN0.PLANSESOL Then
            If oldZPLAN0.PLANSESOL <> " " Then
                mXls2_Row = mXls2_Row + 1: Nb_Err_PLANSESOL = Nb_Err_PLANSESOL + 1
                wsExcel.Cells(mXls2_Row, 1) = "YBIACPT0"
                wsExcel.Cells(mXls2_Row, 2) = xYBIACPT0.COMPTECOM
                wsExcel.Cells(mXls2_Row, 3) = xYBIACPT0.COMPTEINT
                wsExcel.Cells(mXls2_Row, 4) = -xYBIACPT0.SOLDECEN
                wsExcel.Cells(mXls2_Row, 5) = "code 'sens' du compte : " & xYBIACPT0.COMPTESEN & " # " & oldZPLAN0.PLANSESOL & " (rubrique comptable " & xYBIACPT0.COMPTEOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
                For K = 1 To 5: wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_Y0: Next K
            End If
        End If
'___________________________________________________________________________________________
        
        Select Case xYBIACPT0.COMPTESEN
            Case " "
            Case "D"
                If xYBIACPT0.SOLDECEN < 0 Then
                    mXls2_Row = mXls2_Row + 1: Nb_Err_SOLDECEN = Nb_Err_SOLDECEN + 1
                    wsExcel.Cells(mXls2_Row, 1) = "YBIACPT0"
                    wsExcel.Cells(mXls2_Row, 2) = xYBIACPT0.COMPTECOM
                    wsExcel.Cells(mXls2_Row, 3) = xYBIACPT0.COMPTEINT
                    wsExcel.Cells(mXls2_Row, 4) = -xYBIACPT0.SOLDECEN
                    wsExcel.Cells(mXls2_Row, 5) = "Ce compte devrait être débiteur : " & " " & xYBIACPT0.COMPTEDEV
                    For K = 1 To 5
                        wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_Y1
                        wsExcel.Cells(mXls2_Row, K).Font.Color = vbRed
                    Next K
                End If
            Case "C"
                If xYBIACPT0.SOLDECEN > 0 Then
                    mXls2_Row = mXls2_Row + 1: Nb_Err_SOLDECEN = Nb_Err_SOLDECEN + 1
                    wsExcel.Cells(mXls2_Row, 1) = "YBIACPT0"
                    wsExcel.Cells(mXls2_Row, 2) = xYBIACPT0.COMPTECOM
                    wsExcel.Cells(mXls2_Row, 3) = xYBIACPT0.COMPTEINT
                    wsExcel.Cells(mXls2_Row, 4) = -xYBIACPT0.SOLDECEN
                    wsExcel.Cells(mXls2_Row, 5) = "Ce compte devrait être créditeur : " & " " & xYBIACPT0.COMPTEDEV
                    For K = 1 To 5
                        wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_Y1
                        wsExcel.Cells(mXls2_Row, K).Font.Color = vbRed
                    Next K
                End If
            End Select
'___________________________________________________________________________________________
        If xYBIACPT0.COMPTEOBL > "999990    " And xYBIACPT0.COMPTEOBL < "999999    " Then
                mXls2_Row = mXls2_Row + 1: Nb_Err_999999 = Nb_Err_999999 + 1
                wsExcel.Cells(mXls2_Row, 1) = "YBIACPT0"
                wsExcel.Cells(mXls2_Row, 2) = xYBIACPT0.COMPTECOM
                wsExcel.Cells(mXls2_Row, 3) = xYBIACPT0.COMPTEINT
                wsExcel.Cells(mXls2_Row, 4) = -xYBIACPT0.SOLDECEN
                wsExcel.Cells(mXls2_Row, 5) = "compte à clôturer : " & xYBIACPT0.COMPTEDEV & " -   (rubrique comptable " & xYBIACPT0.COMPTEOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
                For K = 1 To 5: wsExcel.Cells(mXls2_Row, K).Interior.Color = mColor_Y0: Next K
        End If
'___________________________________________________________________________________________

        
    End If
    rsSab.MoveNext
Loop

Set wsExcel = wbExcel.Sheets(1)

For arrDev_K = 0 To arrDev_Nb
    If arrB_DB(arrDev_K) + arrB_CR(arrDev_K) <> 0 Then
        mXls1_Row = mXls1_Row + 1: Nb_Err_Balance = Nb_Err_Balance + 1
        wsExcel.Cells(mXls1_Row, 1) = "YBIACPT0"
        wsExcel.Cells(mXls1_Row, 4) = arrB_DB(arrDev_K) + arrB_CR(arrDev_K)
        wsExcel.Cells(mXls1_Row, 5) = arrDev(arrDev_K) & " - Balance Bilan non équilibrée"
        For K = 1 To 5
            wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W1
            wsExcel.Cells(mXls1_Row, K).Font.Color = mColor_Y1
            wsExcel.Cells(mXls1_Row, K).Font.Bold = True
        Next K
    End If
    If arrHB_DB(arrDev_K) + arrHB_CR(arrDev_K) <> 0 Then
        mXls1_Row = mXls1_Row + 1: Nb_Err_Balance = Nb_Err_Balance + 1
        wsExcel.Cells(mXls1_Row, 1) = "YBIACPT0"
        wsExcel.Cells(mXls1_Row, 4) = arrHB_DB(arrDev_K) + arrHB_CR(arrDev_K)
        wsExcel.Cells(mXls1_Row, 5) = arrDev(arrDev_K) & " - Balance Hors-Bilan non équilibrée"
        For K = 1 To 5
            wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W1
            wsExcel.Cells(mXls1_Row, K).Font.Color = mColor_Y1
            wsExcel.Cells(mXls1_Row, K).Font.Bold = True
        Next K
    End If

Next arrDev_K

'_________________________________________________________________________________________________________

'_________________________________________________________________________________________________________
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "YBIACPT0"
If Nb_Err_Balance = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* balances Bilan & Hors-bilan en devise equilibrées : Ok (" & Nb_Lu & " enregistrements lus)."
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* balances Bilan & Hors-bilan en devise : " & Nb_Err_Balance & " anomalies (" & Nb_Lu & " enregistrements lus)."
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W1
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If
'_________________________________________________________________________________________________________
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "YBIACPT0"
If Nb_Err_COMPTEOBL = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* contrôle des rubriques comptables associées aux comptes : Ok "
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* comptes dont la rubrique comptable est inconnue : " & Nb_Err_COMPTEOBL & " anomalies "
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If

'_________________________________________________________________________________________________________
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "YBIACPT0"
If Nb_Err_COMPTEDEV = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* contrôle de la devise / l'identifiant du compte : Ok"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* contrôle de la devise / l'identifiant du compte : " & Nb_Err_COMPTEDEV & " anomalies"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0
    Next K
End If

'_________________________________________________________________________________________________________
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "YBIACPT0"
If Nb_Err_CLIENARSD = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* contrôle du code zone fiscale du PCI conforme avec le pays du client : Ok"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* contrôle du code zone fiscale du PCI non conforme avec le pays du client : " & Nb_Err_CLIENARSD & " anomalies"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0
    Next K
End If

'_________________________________________________________________________________________________________
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "YBIACPT0"
If Nb_Err_PLANCLASS = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* contrôle du code 'classe de sécurité' du compte / rubrique comptable : Ok"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* contrôle du code 'classe de sécurité' du compte / rubrique comptable  : " & Nb_Err_PLANCLASS & " anomalies"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If
'_________________________________________________________________________________________________________
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "YBIACPT0"
If Nb_Err_PLANFONCT = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* contrôle du code 'fonctionnement' du compte / rubrique comptable : Ok"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* contrôle du code 'fonctionnement' du compte / rubrique comptable : " & Nb_Err_PLANFONCT & " anomalies"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If
'_________________________________________________________________________________________________________
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "YBIACPT0"
If Nb_Err_PLANSESOL = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* contrôle du code 'sens' du compte / rubrique comptable : Ok"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* contrôle du code 'sens' du compte / rubrique comptable : " & Nb_Err_PLANSESOL & " anomalies"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If
'_________________________________________________________________________________________________________
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "YBIACPT0"
If Nb_Err_COMPTESEN = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* contrôle du solde des comptes conforme au sens comptable : Ok"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* contrôle du solde des comptes NON conforme au sens comptable : " & Nb_Err_COMPTESEN & " anomalies"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If
'_________________________________________________________________________________________________________
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "YBIACPT0"
If Nb_Err_999999 = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* comptes à clôturer (PCI : 99999*) : Néant"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* comptes à clôturer  (PCI : 99999*): " & Nb_Err_999999
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If

Exit Sub

Error_Handler:
If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub cmdSelect_SQL_Surveillance_ZCOMREF0()
Dim X As String
Dim K As Long
Dim Nb_Err_COMREFCOM As Long

 On Error GoTo Error_Handler

'==============================================================================
Call lstErr_AddItem(lstErr, cmdContext, "> contrôle : ZCOMREF0"): DoEvents
'Set wsExcel = wbExcel.Sheets(2)
Set wsExcel = wbExcel.Sheets(1)

'=======================================================================================================
X = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
   & " where COMPTEFON <> '4' and CLIENACLI = '' and COMPTECOM not in " _
   & " ( select COMREFCOM from " & paramIBM_Library_SAB & ".ZCOMREF0 where COMREFCOR like 'G%')" _
       & " order by COMPTECOM"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    
'___________________________________________________________________________________________
    mXls1_Row = mXls1_Row + 1: Nb_Err_COMREFCOM = Nb_Err_COMREFCOM + 1
    wsExcel.Cells(mXls1_Row, 1) = "ZCOMREF0"
    wsExcel.Cells(mXls1_Row, 2) = rsSab("COMPTECOM")
    wsExcel.Cells(mXls1_Row, 3) = rsSab("COMPTEINT")
    wsExcel.Cells(mXls1_Row, 4) = -CCur(rsSab("SOLDECEN")) / 1000
    wsExcel.Cells(mXls1_Row, 5) = rsSab("COMPTEDEV") & " - compte à affecter à un service G? "
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_Y0: Next K
'___________________________________________________________________________________________

        
    rsSab.MoveNext
Loop
'_________________________________________________________________________________________________________
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "YBIACPT0"
If Nb_Err_COMREFCOM = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* tous les comptes généraux sont affectés à un service G* : Ok "
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* il y a des comptes généraux à affecter à un service G* : " & Nb_Err_COMREFCOM & " anomalies "
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W1
    Next K
End If

Exit Sub

Error_Handler:
If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub cmdSelect_SQL_Surveillance_COMPTEOUV()
Dim X As String
Dim K As Long, wIBMMin As Long
Dim Nb_Err_COMPTEOUV As Long
Dim blnOk As Boolean
 On Error GoTo Error_Handler

'==============================================================================
Call lstErr_AddItem(lstErr, cmdContext, "> contrôle : COMPTEOUV"): DoEvents
Set wsExcel = wbExcel.Sheets(1)
wIBMMin = YBIATAB0_DATE_CPT_J - 19000000
'=======================================================================================================
X = "select * from " & paramIBM_Library_SAB & ".ZCOMPTE0 " _
   & " where COMPTEOUV >= " & wIBMMin _
       & " order by COMPTEOBL,COMPTECOM"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    
'___________________________________________________________________________________________
    V = rsZCOMPTE0_GetBuffer(rsSab, xZCOMPTE0)

    If xZCOMPTE0.COMPTEOBL <> oldZPLAN0.PLANCOOBL Then
        blnOk = False
        For arrZPLAN0_Index = 1 To arrZPLAN0_Nb
            If xZCOMPTE0.COMPTEOBL = arrZPLAN0(arrZPLAN0_Index).PLANCOOBL Then
               oldZPLAN0 = arrZPLAN0(arrZPLAN0_Index)
               oldYPCICPT0 = arrYPCICPT0(arrZPLAN0_Lnk(arrZPLAN0_Index))
               blnOk = True
                Exit For
            End If
        Next arrZPLAN0_Index
        If Not blnOk Then
            oldZPLAN0.PLANINTIT = "?????????": oldZPLAN0.PLANCOOBL = xZCOMPTE0.COMPTEOBL
            oldYPCICPT0.PCICPTAUTO = "I"
        End If
    End If
    mXls1_Row = mXls1_Row + 1: Nb_Err_COMPTEOUV = Nb_Err_COMPTEOUV + 1
    wsExcel.Cells(mXls1_Row, 1) = "ZCOMPTE0"
    wsExcel.Cells(mXls1_Row, 2) = xZCOMPTE0.COMPTECOM
    wsExcel.Cells(mXls1_Row, 3) = xZCOMPTE0.COMPTEINT
    'wsExcel.Cells(mXls1_Row, 4) = -xZCOMPTE0.SOLDECEN
    wsExcel.Cells(mXls1_Row, 5) = xZCOMPTE0.COMPTEDEV & " - compte ouvert le " & dateImp10(YBIATAB0_DATE_CPT_J) & " (rubrique comptable " & oldZPLAN0.PLANCOOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
    If Not blnOk Then For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0: Next K
    If oldYPCICPT0.PCICPTAUTO <> "I" Then
        xYPCICPT0 = oldYPCICPT0
        If Not PCICPTMETA_Control(X) Then
            mXls1_Row = mXls1_Row + 1
            wsExcel.Cells(mXls1_Row, 1) = "ZCOMPTE0"
            wsExcel.Cells(mXls1_Row, 2) = xZCOMPTE0.COMPTECOM
            wsExcel.Cells(mXls1_Row, 3) = xZCOMPTE0.COMPTEINT
            'wsExcel.Cells(mXls1_Row, 4) = -xZCOMPTE0.SOLDECEN
            wsExcel.Cells(mXls1_Row, 5) = " - structure non conforme au PCI : " & Trim(X)
            For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0: Next K
        End If
    End If

'___________________________________________________________________________________________

        
    rsSab.MoveNext
Loop
'_________________________________________________________________________________________________________
Set wsExcel = wbExcel.Sheets(1)
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "ZCOMPTE0"
If Nb_Err_COMPTEOUV = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* pas de compte ouvert le " & dateImp10(YBIATAB0_DATE_CPT_J)
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(96, 190, 255)
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* " & Nb_Err_COMPTEOUV & " comptes ouverts le " & dateImp10(YBIATAB0_DATE_CPT_J)
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(96, 190, 255)
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If

Exit Sub

Error_Handler:
If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub cmdSelect_SQL_Surveillance_COMPTEFON()
Dim X As String, xWhere As String
Dim K As Long
Dim Nb_Err_COMPTEFON As Long, Nb_COMPTEFON As Long
Dim blnOk As Boolean
 On Error GoTo Error_Handler
 
'Call rsYBIATAB0_Read("SQL_Client", "Libye", "Embargo", xWhere)
X = "select CLIGRPCLI from " & paramIBM_Library_SAB & ".ZCLIGRP0" _
& " where CLIGRPREG = '0006000'" _
& "  order by CLIGRPCLI"
 Set rsSab = cnsab.Execute(X)
 
 Do While Not rsSab.EOF
     xWhere = xWhere & ",'" & rsSab("CLIGRPCLI") & "'"
     rsSab.MoveNext
 Loop
If xWhere <> "" Then Mid$(xWhere, 1, 1) = " "

'==============================================================================
Call lstErr_AddItem(lstErr, cmdContext, "> contrôle : COMPTEFON"): DoEvents
Set wsExcel = wbExcel.Sheets(1)
'=======================================================================================================
'X = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
'   & " where COMPTEFON <>'4' And CLIENACLI in (" & xWhere & ")" _
'       & " order by COMPTEOBL,CLIENACLI,COMPTECOM"
'Set rsSab = cnsab.Execute(X)


X = "select * from " & paramIBM_Library_SAB & ".ZTITULA0, " & paramIBM_Library_SAB & ".ZCOMPTE0" _
  & "  where TITULACLI in (" & xWhere & ")" _
  & "  and TITULACOM = COMPTECOM and TITULATPR = '0' and COMPTEFON <>'4'" _
  & " order by COMPTEOBL,TITULACLI,COMPTECOM"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    
'___________________________________________________________________________________________
    Nb_COMPTEFON = Nb_COMPTEFON + 1
    
    V = rsZCOMPTE0_GetBuffer(rsSab, xZCOMPTE0)
    For arrZPLAN0_Index = 1 To arrZPLAN0_Nb
        If xZCOMPTE0.COMPTEOBL = arrZPLAN0(arrZPLAN0_Index).PLANCOOBL Then
           oldZPLAN0 = arrZPLAN0(arrZPLAN0_Index)
           oldYPCICPT0 = arrYPCICPT0(arrZPLAN0_Lnk(arrZPLAN0_Index))
           blnOk = True
            Exit For
        End If
    Next arrZPLAN0_Index

    If xZCOMPTE0.COMPTEFON <> "3" Then
   
        mXls1_Row = mXls1_Row + 1
        wsExcel.Cells(mXls1_Row, 1) = "COMPTEFON"
        wsExcel.Cells(mXls1_Row, 2) = xZCOMPTE0.COMPTECOM
        wsExcel.Cells(mXls1_Row, 3) = xZCOMPTE0.COMPTEINT
        'wsExcel.Cells(mXls1_Row, 4) = -xZCOMPTE0.SOLDECEN
        If mId$(xZCOMPTE0.COMPTEOBL, 1, 5) <> "12123" Then
            Nb_Err_COMPTEFON = Nb_Err_COMPTEFON + 1
            wsExcel.Cells(mXls1_Row, 5) = xZCOMPTE0.COMPTEDEV & " - ce compte n'est pas bloqué. Code fonctionnement = " & xZCOMPTE0.COMPTEFON _
                                    & " (PCI : " & xZCOMPTE0.COMPTEOBL & " " & Trim(oldZPLAN0.PLANINTIT) & ")"
            For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0: Next K
        Else
            wsExcel.Cells(mXls1_Row, 5) = xZCOMPTE0.COMPTEDEV & " - groupe Libye (6000). Code fonctionnement = " & xZCOMPTE0.COMPTEFON _
                                    & " (PCI : " & xZCOMPTE0.COMPTEOBL & " " & Trim(oldZPLAN0.PLANINTIT) & ")"
            For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_Y0: Next K
        
        End If
    Else
         If mId$(xZCOMPTE0.COMPTEOBL, 1, 5) = "12123" Then
            mXls1_Row = mXls1_Row + 1
            wsExcel.Cells(mXls1_Row, 1) = "COMPTEFON"
            wsExcel.Cells(mXls1_Row, 2) = xZCOMPTE0.COMPTECOM
            wsExcel.Cells(mXls1_Row, 3) = xZCOMPTE0.COMPTEINT

            Nb_Err_COMPTEFON = Nb_Err_COMPTEFON + 1
            wsExcel.Cells(mXls1_Row, 5) = xZCOMPTE0.COMPTEDEV & " - ce compte ne doit pas être bloqué. Code fonctionnement = " & xZCOMPTE0.COMPTEFON _
                                    & " (PCI : " & xZCOMPTE0.COMPTEOBL & " " & Trim(oldZPLAN0.PLANINTIT) & ")"
            For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0: Next K
        End If
    End If

'___________________________________________________________________________________________

        
    rsSab.MoveNext
Loop
'_________________________________________________________________________________________________________
Set wsExcel = wbExcel.Sheets(1)
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "ZCOMPTE0"
If Nb_Err_COMPTEFON = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* les " & Nb_COMPTEFON & " comptes du groupe 'Libye' sont bloqués"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* " & Nb_Err_COMPTEFON & " / " & Nb_COMPTEFON & " comptes du groupe 'Libye' ne sont pas bloqués"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W1
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If

Exit Sub

Error_Handler:
If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub cmdSelect_SQL_Surveillance_YPCICPT0_Update()
Dim X As String, X1 As String
Dim K As Long, K1 As Long
Dim Nb_Err_YPCICPT0 As Long
Dim blnOk As Boolean
 On Error GoTo Error_Handler

'==============================================================================
Call lstErr_AddItem(lstErr, cmdContext, "> contrôle : YPCICPT0"): DoEvents
Set wsExcel = wbExcel.Sheets(1)
'=======================================================================================================
For K = 1 To arrYPCICPT0_Nb
    If arrYPCICPT0(K).PCICPTUAMJ >= YBIATAB0_DATE_CPT_J Then
        mXls1_Row = mXls1_Row + 1: Nb_Err_YPCICPT0 = Nb_Err_YPCICPT0 + 1
        wsExcel.Cells(mXls1_Row, 1) = "YPCICPT0"
        X1 = arrYPCICPT0(K).PCICPTBASE
        Select Case arrYPCICPT0(K).PCICPTAUTO
            Case "M": X1 = X1 & " M"
            Case "I": X1 = X1 & " I"
        End Select
        wsExcel.Cells(mXls1_Row, 2) = X1
        wsExcel.Cells(mXls1_Row, 3) = arrYPCICPT0(K).PCICPTMETA
        If arrYPCICPT0(K).PCICPTUSEQ = 0 Then
            X = "Création par "
        Else
            X = "Modification par "
        End If
        
        wsExcel.Cells(mXls1_Row, 5) = X & arrYPCICPT0(K).PCICPTUUSR & " le " & dateImp10(arrYPCICPT0(K).PCICPTUAMJ) & "  " & timeImp(arrYPCICPT0(K).PCICPTUHMS)
        For K1 = 1 To 5: wsExcel.Cells(mXls1_Row, K1).Interior.Color = RGB(164, 220, 255): Next K1
        
        If Trim(arrYPCICPT0(K).PCICPTSUFX) = "" And Trim(arrYPCICPT0(K).PCICPTTXT) = "" Then
        Else
            mXls1_Row = mXls1_Row + 1
            wsExcel.Cells(mXls1_Row, 1) = "YPCICPT0"
            wsExcel.Cells(mXls1_Row, 2) = X1
            wsExcel.Cells(mXls1_Row, 3) = "suffixe : " & Trim(arrYPCICPT0(K).PCICPTSUFX)
            wsExcel.Cells(mXls1_Row, 5) = "" & Trim(arrYPCICPT0(K).PCICPTTXT)
            For K1 = 1 To 5: wsExcel.Cells(mXls1_Row, K1).Interior.Color = RGB(190, 236, 255): Next K1
        End If
      'arrYPCILNK0_Lnk(K)
    End If
Next K
    
'___________________________________________________________________________________________
'_________________________________________________________________________________________________________
Set wsExcel = wbExcel.Sheets(1)
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "YPCICPT0"
If Nb_Err_YPCICPT0 = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* pas de modification du plan comptable le " & dateImp10(YBIATAB0_DATE_CPT_J)
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(96, 190, 255)
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* " & Nb_Err_YPCICPT0 & " modifications du plan comptable le " & dateImp10(YBIATAB0_DATE_CPT_J)
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(96, 190, 255)
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If

Exit Sub

Error_Handler:
If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub cmdSelect_SQL_Surveillance_ZTITULA0()
Dim X As String
Dim K As Long, wIBMMin As Long
Dim Nb_Err_TITULACLI As Long
Dim blnOk As Boolean
 On Error GoTo Error_Handler

'==============================================================================
Call lstErr_AddItem(lstErr, cmdContext, "> contrôle :ZTITULA0"): DoEvents
Set wsExcel = wbExcel.Sheets(1)
wIBMMin = YBIATAB0_DATE_CPT_J - 19000000
'=======================================================================================================
X = "select * from " & paramIBM_Library_SAB & ".ZTITULA0 , " & paramIBM_Library_SAB & ".ZCOMPTE0" _
   & " WHERE (titulacli > '9900000' or  titulacli < '0010000')" _
   & " and titulacom = comptecom and comptefon <> '4'"

Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    
'___________________________________________________________________________________________
    mXls1_Row = mXls1_Row + 1: Nb_Err_TITULACLI = Nb_Err_TITULACLI + 1
    wsExcel.Cells(mXls1_Row, 1) = "ZTITULA0"
    wsExcel.Cells(mXls1_Row, 2) = rsSab("COMPTECOM")
    wsExcel.Cells(mXls1_Row, 3) = rsSab("COMPTEINT")
    wsExcel.Cells(mXls1_Row, 5) = rsSab("COMPTEDEV") & " - anomalie : le titulaire " & rsSab("TITULACLI") & " est hors plage ('0010000'-'0099999')"
    For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0: Next K
'___________________________________________________________________________________________

        
    rsSab.MoveNext
Loop
'_________________________________________________________________________________________________________
Set wsExcel = wbExcel.Sheets(1)
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "ZTITULA0"
If Nb_Err_TITULACLI = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* pas d'anomalie / plage de titulaires ('0010000'-'0099999')"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* " & Nb_Err_TITULACLI & " anomalies / plage de titulaires ('0010000'-'0099999')"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W1
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If

Exit Sub

Error_Handler:
If Not blnAuto Then MsgBox Error, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub cmdSelect_SQL_Surveillance_YBIAMVTH()
Dim X As String
Dim K As Long, wIBMMin As Long
Dim Nb_Err_YBIAMVTH As Long
Dim blnRupture As Boolean
Dim mCOMPTEDEV As String, mMOUVEMPIE As Long, blnBilan As Boolean, curSolde As Currency
Dim xYBIAMVT0 As typeYBIAMVT0
 On Error GoTo Error_Handler

'==============================================================================
Call lstErr_AddItem(lstErr, cmdContext, "> contrôle :YBIAMVTH"): DoEvents
Set wsExcel = wbExcel.Sheets(1)
wIBMMin = YBIATAB0_DATE_CPT_J - 19000000
mCOMPTEDEV = "": mMOUVEMPIE = 0: curSolde = 0: blnBilan = True
'=======================================================================================================
X = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTHD  " _
   & " WHERE MOUVEMDTR = " & wIBMMin _
   & " order by MOUVEMPIE, COMPTEDEV, COMPTEOBL"

Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF

    blnRupture = False
    If mMOUVEMPIE <> rsSab("MOUVEMPIE") Or mCOMPTEDEV <> rsSab("COMPTEDEV") Then
        blnRupture = True
    Else
        If mId$(rsSab("COMPTEOBL"), 1, 1) = "9" Then
            If blnBilan Then blnRupture = True
        Else
            If Not blnBilan Then blnRupture = True
        End If
    End If
    
     If blnRupture Then
        If curSolde <> 0 Then
            mXls1_Row = mXls1_Row + 1: Nb_Err_YBIAMVTH = Nb_Err_YBIAMVTH + 1
            wsExcel.Cells(mXls1_Row, 1) = "B / HB"
            wsExcel.Cells(mXls1_Row, 2) = "Pièce " & mMOUVEMPIE
            wsExcel.Cells(mXls1_Row, 3) = xYBIAMVT0.MOUVEMSER & "-" & xYBIAMVT0.MOUVEMSSE & " " & xYBIAMVT0.MOUVEMOPE & " " & xYBIAMVT0.MOUVEMEVE & " " & xYBIAMVT0.MOUVEMNUM
            wsExcel.Cells(mXls1_Row, 4) = Format$(curSolde, "### ### ### ##0.00") & " " & mCOMPTEDEV
            If blnBilan Then
                wsExcel.Cells(mXls1_Row, 5) = "écart Bilan " & dateImp10(YBIATAB0_DATE_CPT_J)
            Else
                wsExcel.Cells(mXls1_Row, 5) = "écart Hors-Bilan " & dateImp10(YBIATAB0_DATE_CPT_J)
            End If
            For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0: Next K
        End If
        Call rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVT0)
        
        mMOUVEMPIE = rsSab("MOUVEMPIE")
        mCOMPTEDEV = rsSab("COMPTEDEV")
        If mId$(rsSab("COMPTEOBL"), 1, 1) = "9" Then
            blnBilan = False
        Else
            blnBilan = True
        End If
        curSolde = 0
    End If
    
    curSolde = curSolde + rsSab("MOUVEMMON")
    If xYBIAMVT0.MOUVEMOPE = "ECH" Then
        X = rsSab("MOUVEMCOM")
        If InStr(X, "DTX") > 0 And mId$(rsSab("COMPTEOBL"), 1, 1) <> "9" Then
        
            mXls1_Row = mXls1_Row + 1
            wsExcel.Cells(mXls1_Row, 1) = "B / HB"
            wsExcel.Cells(mXls1_Row, 2) = X
            wsExcel.Cells(mXls1_Row, 3) = rsSab("COMPTEINT")
            wsExcel.Cells(mXls1_Row, 4) = Format$(rsSab("MOUVEMMON"), "### ### ### ##0.00") & " " & mCOMPTEDEV
            wsExcel.Cells(mXls1_Row, 5) = "écriture ECH : compte DTX" & rsSab("COMPTEOBL")
        End If
    End If
    
    rsSab.MoveNext
Loop
'_________________________________________________________________________________________________________
Set wsExcel = wbExcel.Sheets(1)
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "YBIAMVTH"
If Nb_Err_YBIAMVTH = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* pas d'anomalie d'équilibre des écritures comptables au " & dateImp10(YBIATAB0_DATE_CPT_J)
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* " & Nb_Err_YBIAMVTH & " anomalies d'équilibre des écritures comptables au " & dateImp10(YBIATAB0_DATE_CPT_J)
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W1
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If

Exit Sub

Error_Handler:
If Not blnAuto Then MsgBox Error, vbCritical, Me.Name & " : " & currentAction


End Sub



Public Sub cmdSelect_SQL_Surveillance_ZCLIENA0()
Dim X As String
Dim K As Long, wIBMMin As Long
Dim Nb_Err_CLIENACLI As Long
Dim blnOk As Boolean
 On Error GoTo Error_Handler

'==============================================================================
Call lstErr_AddItem(lstErr, cmdContext, "> contrôle :ZCLIENA0"): DoEvents
Set wsExcel = wbExcel.Sheets(1)
'=======================================================================================================
X = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
   & "order by CLIENARES"

Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    
    If IsNull(arrCLIEANRES_Control(rsSab("CLIENARES"))) Then
'___________________________________________________________________________________________
        mXls1_Row = mXls1_Row + 1: Nb_Err_CLIENACLI = Nb_Err_CLIENACLI + 1
        wsExcel.Cells(mXls1_Row, 1) = "ZCLIENA0"
        wsExcel.Cells(mXls1_Row, 2) = rsSab("CLIENACLI")
        wsExcel.Cells(mXls1_Row, 3) = rsSab("CLIENARA1")
        wsExcel.Cells(mXls1_Row, 5) = "le responsable " & rsSab("CLIENARES") & " n'est pas référencé dans la table 6 de ZBASTAB0"
        For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0: Next K
'___________________________________________________________________________________________

    End If
        
    rsSab.MoveNext
Loop
'_________________________________________________________________________________________________________
Set wsExcel = wbExcel.Sheets(1)
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "ZCLIENA0"
If Nb_Err_CLIENACLI = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* pas d'anomalie Client / table des responsables"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* " & Nb_Err_CLIENACLI & " anomalies Client / table des responsables"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W1
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If

Exit Sub

Error_Handler:
If Not blnAuto Then MsgBox Error, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub cmdSendMail_Surveillance()
Dim wSendMail As typeSendMail
Dim xDétail As String, xHeader As String, mbgColor As String
Dim K As Long, htmlFontColor_K As String
Dim iRow As Integer, iCol As Integer, X As String, xTD As String
Dim wForecolor As String, wBackColor As String, xColor As String
Dim html_B_On As String, html_B_Off As String
Dim blnEnd As Boolean
On Error Resume Next

Me.Enabled = False: Me.MousePointer = vbHourglass

cmdSelect_SQL_Surveillance
'______________________________________________

Set appExcel = CreateObject("Excel.Application")
Set wbExcel = appExcel.Workbooks.Open(wFile)
Set wsExcel = wbExcel.Worksheets(1)
'__________________________________________________________________________________

Call lstErr_AddItem(lstErr, cmdContext, "Importation :  " & wFile): DoEvents

blnEnd = False
K = 0
xHeader = ""
xDétail = ""
mbgColor = "bgcolor = #E0E0E0"

Do
    K = K + 1
    X = Trim(wsExcel.Cells(K, 1))
    If X = "Fin du traitement" Then
        blnEnd = True
    Else
    
        wBackColor = cmdSendMail_Cell_Color(wsExcel.Cells(K, 1).Interior.Color) '   "#FFFFF0" 'htmlFontColor_Blue
        wForecolor = cmdSendMail_Cell_Color(wsExcel.Cells(K, 1).Font.Color) '"#FF0000"
        If wsExcel.Cells(K, 1).Font.Bold Then
            html_B_On = "<B>": html_B_Off = "</B>"
        Else
             html_B_On = "": html_B_Off = ""
       End If
        
        xHeader = xHeader & "<TR>" _
         & "<TD bgcolor=" & wBackColor & " width=50 ALIGN=LEFT><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=" & wForecolor & ">" _
         & html_B_On & X & "</TD>" _
         & "<TD bgcolor=" & wBackColor & " width=150 ALIGN=LEFT><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=" & wForecolor & ">" _
         & html_B_On & Trim(wsExcel.Cells(K, 2)) & html_B_Off & "</TD>" _
         & "<TD bgcolor=" & wBackColor & " width=200 ALIGN=LEFT><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=" & wForecolor & ">" _
         & html_B_On & Trim(wsExcel.Cells(K, 3)) & html_B_Off & "</TD>" _
         & "<TD bgcolor=" & wBackColor & " width=100 ALIGN=RIGHT><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=" & wForecolor & ">" _
         & html_B_On & Trim(wsExcel.Cells(K, 4)) & html_B_Off & "</TD>" _
         & "<TD bgcolor=" & wBackColor & " width=500 ALIGN=LEFT><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=" & wForecolor & ">" _
         & html_B_On & Trim(wsExcel.Cells(K, 5)) & html_B_Off & "</TD>" _
        & "</TR>"

    End If
    
Loop Until blnEnd = True

wbExcel.Saved = True
'____________________________________________________________________________________
appExcel.Quit

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing

'=======================================================================================

wSendMail.From = currentSSIWINMAIL
wSendMail.AsHTML = True

If blnAuto Then
    paramEditionNoPaper_Auto_PgmName = "BIA_PCI_COMPTE"
            
    wSendMail.FromDisplayName = "@PCI_COMPTE"
    If cmdSelect_SQL_K = "X#c" Then
        wSendMail.RecipientDisplayName = "GSOP"
        wSendMail.Subject = "BIA-PCI-COMPTE-GSOP " & dateImp_Amj(YBIATAB0_DATE_CPT_J)
        Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S11", wFile, "Prod", "BIA-PCI-COMPTE-GSOP")
    Else
        wSendMail.RecipientDisplayName = "CPT"
        wSendMail.Subject = "BIA-PCI-COMPTE-CPT " & dateImp_Amj(YBIATAB0_DATE_CPT_J)
        Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S60", wFile, "Archive", "BIA-PCI-COMPTE-CPT")
    End If

    wSendMail.Attachment = ""
    wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                        & "<span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & "<Font color = #404040>" _
                        & htmlFontColor_Black & "<BR>" & paramEditionNoPaper_Auto_Lnk & "<BR>" _
                        & "<TABLE   width=1000 border=1 cellpadding=5 ></B>" _
                        & xHeader _
                        & "</TABLE>"

Else
    wSendMail.Recipient = currentSSIWINMAIL

    wSendMail.Attachment = wFile
    
    wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                        & "<span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & "<Font color = #404040>" _
                        & "<TABLE   width=1000 border=1 cellpadding=5 ></B>" _
                        & xHeader _
                        & "</TABLE>"

End If


srvSendMail.Monitor wSendMail
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdSendMail_PCI_COMPTE()
Dim wSendMail As typeSendMail
Dim K As Long, htmlFontColor_K As String

On Error Resume Next

'____________________________________________________________________________________________
wSendMail.From = currentSSIWINMAIL
wSendMail.AsHTML = True

If blnAuto Then
    paramEditionNoPaper_Auto_PgmName = "BIA-PCI-COMPTE"
        Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S60", wFile, "Archive", "BIA-PCI-COMPTE-CPT-M")
    wSendMail.FromDisplayName = "@PCI_COMPTE"
    wSendMail.RecipientDisplayName = "CPT"
    wSendMail.Subject = "BIA-PCI-COMPTE fin de mois"
    wSendMail.Attachment = ""
    wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                        & "<span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & "<Font color = #404040>" _
                        & htmlFontColor_Black & "<BR>" & paramEditionNoPaper_Auto_Lnk & "<BR>" _

Else
    wSendMail.Recipient = currentSSIWINMAIL
    wSendMail.Subject = "BIA-PCI-COMPTE-CPT-M " & dateImp_Amj(YBIATAB0_DATE_CPT_J)
    wSendMail.Attachment = ""
    wSendMail.Message = "<body bgcolor = #FFFFFF><BR>"
End If


srvSendMail.Monitor wSendMail
'==================================================================================================


End Sub

Public Function cmdSendMail_Cell_Color(lColor As Long) As String
Dim xColor As String, X As String
xColor = Hex(lColor)
Select Case Len(xColor)
    Case 6:
    Case 2: xColor = "0000" & xColor
    Case 4: xColor = "00" & xColor
    Case 1: xColor = "00000" & xColor
    Case 3: xColor = "000" & xColor
End Select

cmdSendMail_Cell_Color = " #" & mId$(xColor, 5, 2) & mId$(xColor, 3, 2) & mId$(xColor, 1, 2)
End Function

Public Sub cmdSelect_SQL_Surveillance_YPCICPT0()
Dim X As String, mPLANCOPRO As String, mPLANFONCT As String
Dim K As Long, K1 As Long, Nb_Err As Long

 On Error GoTo Error_Handler

'==============================================================================
Call lstErr_AddItem(lstErr, cmdContext, "> contrôle : YPCICPT0"): DoEvents
Set wsExcel = wbExcel.Sheets(2)

Nb_Err = 0
Call rsZPLAN0_Init(oldZPLAN0)
Call rsYPCICPT0_Init(oldYPCICPT0)

For arrZPLAN0_Index = 1 To arrZPLAN0_Nb
    xZPLAN0 = arrZPLAN0(arrZPLAN0_Index)
    
    If arrZPLAN0_Lnk(arrZPLAN0_Index) = 0 Then
        mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
        wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0 *"
        wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
        wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
        wsExcel.Cells(mXls2_Row, 5) = "enregistrement orphelin ZPLAN0 => YPCICPT0"
    Else
        K1 = arrYPCILNK0_Lnk(arrZPLAN0_Lnk(arrZPLAN0_Index))
        If K1 <= 0 Then
            mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
            wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0 *"
            wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
            wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
            wsExcel.Cells(mXls2_Row, 5) = "erreur lien ZPLAN0 => YPCICPT0 => ZPLAN0"
        Else

            oldZPLAN0 = arrZPLAN0(K1)
            If xZPLAN0.PLANCLASS <> oldZPLAN0.PLANCLASS Then
                mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
                wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0 *"
                wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
                wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
                wsExcel.Cells(mXls2_Row, 5) = "classe de sécurité " & xZPLAN0.PLANCLASS & " # " & oldZPLAN0.PLANCLASS & "   (" & oldZPLAN0.PLANCOOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
            End If
            
            If xZPLAN0.PLANFONCT <> oldZPLAN0.PLANFONCT Then
                mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
                wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0 *"
                wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
                wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
                wsExcel.Cells(mXls2_Row, 5) = "code fonctionnement " & xZPLAN0.PLANFONCT & " # " & oldZPLAN0.PLANFONCT & "   (" & oldZPLAN0.PLANCOOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
            End If
            If xZPLAN0.PLANSESOL <> oldZPLAN0.PLANSESOL Then
                mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
                wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0 *"
                wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
                wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
                wsExcel.Cells(mXls2_Row, 5) = "sens du solde " & xZPLAN0.PLANSESOL & " # " & oldZPLAN0.PLANSESOL & "   (" & oldZPLAN0.PLANCOOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
            End If
            If xZPLAN0.PLANGEDEP <> oldZPLAN0.PLANGEDEP Then
                mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
                wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0 *"
                wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
                wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
                wsExcel.Cells(mXls2_Row, 5) = "code 'gestion du dépassement' " & xZPLAN0.PLANGEDEP & " # " & oldZPLAN0.PLANGEDEP & "   (" & oldZPLAN0.PLANCOOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
            End If
            If xZPLAN0.PLANTIERS <> oldZPLAN0.PLANTIERS Then
                mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
                wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0 *"
                wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
                wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
                wsExcel.Cells(mXls2_Row, 5) = "code 'saisie du titulaire' " & xZPLAN0.PLANTIERS & " # " & oldZPLAN0.PLANTIERS & "   (" & oldZPLAN0.PLANCOOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
            End If
            If xZPLAN0.PLANFICOB <> oldZPLAN0.PLANFICOB Then
                mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
                wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0 *"
                wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
                wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
                wsExcel.Cells(mXls2_Row, 5) = "code 'compte de la clientèle' " & xZPLAN0.PLANFICOB & " # " & oldZPLAN0.PLANFICOB & "   (" & oldZPLAN0.PLANCOOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
            End If
            If xZPLAN0.PLANCARAC <> oldZPLAN0.PLANCARAC Then
                mXls2_Row = mXls2_Row + 1: Nb_Err = Nb_Err + 1
                wsExcel.Cells(mXls2_Row, 1) = "ZPLAN0 *"
                wsExcel.Cells(mXls2_Row, 2) = xZPLAN0.PLANCOOBL
                wsExcel.Cells(mXls2_Row, 3) = xZPLAN0.PLANINTIT
                wsExcel.Cells(mXls2_Row, 5) = "longueur du compte " & xZPLAN0.PLANCARAC & " # " & oldZPLAN0.PLANCARAC & "   (" & oldZPLAN0.PLANCOOBL & "  " & Trim(oldZPLAN0.PLANINTIT) & ")"
            End If
        End If
    End If
    
Next arrZPLAN0_Index

'==============================================================================
Set wsExcel = wbExcel.Sheets(1)

mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "ZPLAN0 *"

If Nb_Err = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* Contrôle de cohérence des rubriques : Ok (" & arrYPCICPT0_Nb & " enregistrements lus)."
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* Contrôle de cohérence des rubriques : " & Nb_Err & " anomalies, (" & arrYPCICPT0_Nb & " enregistrements lus)."
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If

Exit Sub

Error_Handler:
If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub





Public Sub cmdPrint_YPCICPT0(X As String)
'prtYPCICPT0_Init "YPCICPT0", X
'prtYEICGCC0_Open
'For I = 1 To fgSelect.Rows - 1
'    fgSelect.Row = I
'    fgSelect.Col = 0: xYPCICPT0.DOSSLDDEV = Trim(fgSelect.Text)
'    fgSelect.Col = 1: xYPCICPT0.DOSSLDOPE = Trim(fgSelect.Text)
'     fgSelect.Col = 2: xYPCICPT0.DOSSLDNUM = Val(Trim(fgSelect.Text))
'    prtYPCICPT0_Line arrYPCICPT0(I)
'Next I
'prtYPCICPT0_Close True

End Sub

'______________________________________________________________________
Private Sub fgSelect_Display()
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
Call lstErr_Clear(lstErr, cmdPrint, "Recherche ...... ")

SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Row = 0

currentAction = "fgSelect_Display"

Do While Not rsSab.EOF

     fgSelect.Rows = fgSelect.Rows + 1
     fgSelect.Row = fgSelect.Rows - 1
     fgSelect.Col = 0: fgSelect.Text = rsSab("PCICPTBASE")
     fgSelect.Col = 1: fgSelect.Text = PCICPTMETA_Display(rsSab("PCICPTMETA"))
                       fgSelect.CellFontBold = True
                       fgSelect.CellBackColor = RGB(0, 192, 192)
                       fgSelect.CellForeColor = RGB(255, 255, 255)
     fgSelect.Col = 2: fgSelect.Text = rsSab("PLANINTIT")
    
     fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = I
    rsSab.MoveNext
Loop

fgSelect.Visible = True

Call lstErr_Clear(lstErr, cmdPrint, " > nb PCI* : " & fgSelect.Row)

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub arrYPCICPT0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYPCICPT0(501)
arrYPCICPT0_Max = 500: arrYPCICPT0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YPCICPT0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYPCICPT0_GetBuffer(rsSab, xYPCICPT0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYPCICPT0.fgselect_Display"
        '' Exit Sub
     Else
         arrYPCICPT0_Nb = arrYPCICPT0_Nb + 1
         If arrYPCICPT0_Nb > arrYPCICPT0_Max Then
             arrYPCICPT0_Max = arrYPCICPT0_Max + 100
             ReDim Preserve arrYPCICPT0(arrYPCICPT0_Max)
         End If
         
         arrYPCICPT0(arrYPCICPT0_Nb) = xYPCICPT0
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
Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYBIACPT0(101)
arrYBIACPT0_Max = 100: arrYBIACPT0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

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


Public Sub cmdSelect_Reset()
If blnControl Then
    cmdSelect_Clear
    cmdSelect_Ok.Visible = False 'True
    fraSelect_Options_1.Visible = False
    cmdSelect_SQL_K = Trim(mId$(cboSelect_SQL, 1, 3))
    Select Case cmdSelect_SQL_K
        Case "1", "6c", "6s":
            lblSelect_Where.Caption = "PCI"
            fraSelect_Options.Visible = True: fraSelect_Options_1.Visible = True
            cmdSelect_Ok.Visible = True
        Case "2":
            lblSelect_Where.Caption = "Client"
            fraSelect_Options.Visible = True: fraSelect_Options_1.Visible = True
            cmdSelect_Ok.Visible = True
       Case Else
            lblSelect_Where.Caption = "PCI"
            fraSelect_Options.Visible = True: fraSelect_Options_1.Visible = True
            cmdSelect_Ok.Visible = True
    End Select

End If

End Sub

Public Sub cmdSelect_Clear()
    lstErr.Clear
    fgSelect.Visible = False
    fgDetail.Visible = False: fraDetail.Visible = False
    fraList1.Visible = False
    lstW.Visible = False
    fraCompte.Visible = False
    fraDetail.Visible = False
    cmdPrint.Visible = False
    cmdUpdate_Ok.Visible = False

End Sub


Public Sub cmdDetail_Reset()
If blnControl Then
    lstErr.Clear
    If fgDetail.Visible Then
        fgDetail.Visible = False: fraDetail.Visible = False: cmdUpdate_Ok.Visible = False
        fraList1.Visible = False
        fgDetail_Display
    End If
End If

End Sub


Private Sub cmdSelect_SQL_1()
Dim V, X As String
Dim xSQL As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean
Dim xDOSSLDM As String, xDOSSLDG As String, xDOSSLDK As String
Dim xField1 As String, xK As String, xField2 As String

On Error GoTo Error_Handler
Call lstErr_Clear(lstErr, cmdPrint, "Recherche ..... ")

currentAction = "cmdYPCICPT0_SQL_1"
blnOk = False


xWhere = " where PCICPTLNK = PLANCOOBL"
X = Trim(txtSelect_Where)
If X <> "" Then xWhere = xWhere & " and PCICPTBASE like '" & X & "%'"

X = Trim(cboSelect_PLANCOPRO)
If X <> "" Then xWhere = xWhere & " and PLANCOPRO = '" & X & "'"


xSQL = "select PCICPTBASE,PCICPTMETA,PLANINTIT from " & paramIBM_Library_SABSPE & ".YPCICPT0, " & paramIBM_Library_SAB & ".ZPLAN0 " _
     & xWhere & " order by PCICPTBASE"
Set rsSab = cnsab.Execute(xSQL)


fgSelect_Display

If fgSelect.Rows = 2 Then
    fgSelect.Row = 1
    fgSelect.Col = 0: xYPCICPT0.PCICPTBASE = Trim(fgSelect.Text)
    fgDetail_Display
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_2()
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub fgDetail_Display()
Dim wColor As Long
Dim X As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
Call lstErr_Clear(lstErr, cmdPrint, "Recherche ...... ")

fgDetail.Visible = False: fraDetail.Visible = False: cmdUpdate_Ok.Visible = False
fraList1.Visible = False
fraCompte.Visible = False
fgDetail_Reset
txtPCICPTMETA.Locked = True
txtPCICPTTXT.Locked = True
txtPCICPTSUFX.Locked = True
optPCICPTAUTO_A.Enabled = False
optPCICPTAUTO_M.Enabled = False
optPCICPTAUTO_I.Enabled = False

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.Row = 0

currentAction = "fgDetail_Display"

fraDetail.Caption = xYPCICPT0.PCICPTBASE
txtPCICPTMETA = ""
txtPCICPTTXT = ""
optPCICPTAUTO_A.Value = True

xWhere = "select * from " & paramIBM_Library_SABSPE & ".YPCICPT0 where PCICPTBASE = '" & xYPCICPT0.PCICPTBASE & "'"
Set rsSab = cnsab.Execute(xWhere)

If Not rsSab.EOF Then
    V = rsYPCICPT0_GetBuffer(rsSab, oldYPCICPT0)
    xYPCICPT0 = oldYPCICPT0
    txtPCICPTMETA = PCICPTMETA_Display(xYPCICPT0.PCICPTMETA)
    txtPCICPTTXT = Trim(xYPCICPT0.PCICPTTXT)
    txtPCICPTSUFX = Trim(xYPCICPT0.PCICPTSUFX)
    Select Case xYPCICPT0.PCICPTAUTO
        Case "M": optPCICPTAUTO_M.Value = True
        Case "I": optPCICPTAUTO_I.Value = True
        Case Else: optPCICPTAUTO_A.Value = True
    End Select
Else
    V = "Impossible de lire l'enregistrement : " & xWhere
    GoTo Error_MsgBox
End If


xWhere = "select * from " & paramIBM_Library_SAB & ".ZPLAN0 " _
       & "  where PLANETABL = 1 and PLANPLAN = 1 and PLANCOOBL like '" & Trim(xYPCICPT0.PCICPTBASE) & "%'"
Set rsSab = cnsab.Execute(xWhere)

Do While Not rsSab.EOF
    V = rsZPLAN0_GetBuffer(rsSab, xZPLAN0)
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_DisplayLine I
    
    rsSab.MoveNext
Loop

Call lstErr_Clear(lstErr, cmdPrint, " > nb PCI : " & fgDetail.Row)

fgDetail.Visible = True: fraDetail.Visible = True
fraDetail.Enabled = True

Select Case cmdSelect_SQL_K
    Case "6s"
        cmdUpdate_Ok.Visible = True
        txtPCICPTMETA.Locked = False
        txtPCICPTTXT.Locked = False
        txtPCICPTSUFX.Locked = False
        optPCICPTAUTO_A.Enabled = True
        optPCICPTAUTO_M.Enabled = True
        optPCICPTAUTO_I.Enabled = True
    Case "6c"
        cmdUpdate_Ok.Visible = True
        txtPCICPTTXT.Locked = False
End Select
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgList1_Display()
Dim wColor As Long
Dim X As String, xWhere As String

Dim I As Long
Dim blnZSOLDE0 As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
Call lstErr_Clear(lstErr, cmdPrint, "Recherche ...... ")

mMsgBox_Err = ""

fraList1.Visible = False ': SSTab2.Tab = 0
fraCompte.Visible = False
fgList1_Reset

fgList1.Rows = 1
fgList1.FormatString = fgList1_FormatString
fgList1.Row = 0

currentAction = "fgList1_Display"

xWhere = "select * from " & paramIBM_Library_SAB & ".ZPLAN0 where PLANCOOBL = '" & Trim(xZPLAN0.PLANCOOBL) & "'"
Set rsSab = cnsab.Execute(xWhere)

If rsSab.EOF Then
    Call MsgBox("erreur de lecture ZPLAN0 : " & Trim(xZPLAN0.PLANCOOBL), vbCritical, "PCI_Compte")
Else
    V = rsZPLAN0_GetBuffer(rsSab, xZPLAN0)
End If
xWhere = "select * from " & paramIBM_Library_SAB & ".ZCOMPTE0 LEFT OUTER JOIN " & paramIBM_Library_SAB & ".ZSOLDE0" _
       & " on COMPTECOM = SOLDECOM" _
       & "  where COMPTEETA = 1 and COMPTEPLA = 1 and COMPTEOBL = '" & Trim(xZPLAN0.PLANCOOBL) & "'"
       
If Trim(cboSelect_DEV) <> "" Then xWhere = xWhere & " and  COMPTEDEV = '" & Trim(cboSelect_DEV) & "'"
If chkSelect_COMPTEFON <> "1" Then xWhere = xWhere & " and  COMPTEFON <> '4'"
Set rsSab = cnsab.Execute(xWhere & " order by COMPTECOM")
blnZSOLDE0 = True
Do While Not rsSab.EOF
    V = rsZCOMPTE0_GetBuffer(rsSab, xZCOMPTE0)
    
    xZSOLDE0.SOLDECEN = rsSab("SOLDECEN")
    fgList1.Rows = fgList1.Rows + 1
    fgList1.Row = fgList1.Rows - 1
    fgList1_DisplayLine I
    
    rsSab.MoveNext
Loop
Call lstErr_Clear(lstErr, cmdPrint, " > nb comptes : " & fgList1.Row)

fraList1.Visible = True

If mMsgBox_Err <> "" Then Call MsgBox(mMsgBox_Err, vbExclamation, "Contrôles de cohérence ZPLAN0 / ZCOMPTE0")

Exit Sub

Error_Handler:
    If blnZSOLDE0 Then
        xZSOLDE0.SOLDECEN = 0
        Resume Next
    End If
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgDetail_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long

On Error Resume Next
fgDetail.Col = 0

fgDetail.Col = 0: fgDetail.Text = xZPLAN0.PLANCOOBL
fgDetail.Col = 1: fgDetail.Text = xZPLAN0.PLANCOPRO
fgDetail.Col = 2: fgDetail.Text = xZPLAN0.PLANCLASS

fgDetail.Col = 3
Select Case xZPLAN0.PLANFONCT
 Case 0
 Case 1: fgDetail.Text = "DB": fgDetail.CellBackColor = mColor_Y1
 Case 2: fgDetail.Text = "CR": fgDetail.CellBackColor = mColor_Y1
 Case 3: fgDetail.Text = "DB-CR": fgDetail.CellBackColor = mColor_Y1
 Case 4: fgDetail.Text = "clos": fgDetail.CellBackColor = mColor_Y1
 Case Else: fgDetail.Text = xZPLAN0.PLANFONCT: fgDetail.CellBackColor = mColor_W1
End Select
fgDetail.Col = 4
Select Case xZPLAN0.PLANSESOL
 Case " "
 Case "D": fgDetail.Text = "DB": fgDetail.CellBackColor = mColor_Y1
 Case "C": fgDetail.Text = "CR": fgDetail.CellBackColor = mColor_Y1
 Case Else: fgDetail.Text = xZPLAN0.PLANSESOL: fgDetail.CellBackColor = mColor_W1
End Select
fgDetail.Col = 5
Select Case xZPLAN0.PLANGEDEP
 Case "N"
 Case "O": fgDetail.Text = "oui": fgDetail.CellBackColor = mColor_Y1
 Case Else: fgDetail.Text = xZPLAN0.PLANGEDEP: fgDetail.CellBackColor = mColor_W1
End Select
fgDetail.Col = 6
Select Case xZPLAN0.PLANTIERS
 Case "N"
 Case "O": fgDetail.Text = "oui": fgDetail.CellBackColor = mColor_Y1
 Case Else: fgDetail.Text = xZPLAN0.PLANTIERS: fgDetail.CellBackColor = mColor_W1
End Select
fgDetail.Col = 7
Select Case xZPLAN0.PLANFICOB
 Case "N"
 Case "O": fgDetail.Text = "oui": fgDetail.CellBackColor = mColor_Y1
 Case Else: fgDetail.Text = xZPLAN0.PLANFICOB: fgDetail.CellBackColor = mColor_W1
End Select

fgDetail.Col = 8: fgDetail.Text = xZPLAN0.PLANCARAC
fgDetail.Col = 9: fgDetail.Text = xZPLAN0.PLANINTIT


fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex
End Sub


Public Sub fgList1_DisplayLine(lIndex As Long)
Dim X As String

fgList1.Col = 0: fgList1.Text = xZCOMPTE0.COMPTEOBL
fgList1.Col = 1: fgList1.Text = xZCOMPTE0.COMPTECOM
If xYPCICPT0.PCICPTAUTO <> "I" Then
    If Not PCICPTMETA_Control(X) Then
        fgList1.CellBackColor = mColor_W1
        fgList1.Col = 13: fgList1.Text = X
        fgList1.CellBackColor = mColor_W1
    End If
End If
fgList1.Col = 2: fgList1.Text = xZCOMPTE0.COMPTEDEV
fgList1.Col = 5: fgList1.Text = xZCOMPTE0.COMPTECLA
If xZCOMPTE0.COMPTECLA < xZPLAN0.PLANCLASS Then
    mMsgBox_Err = mMsgBox_Err & " - " & xZCOMPTE0.COMPTECOM & " COMPTECLA < PLANCLASS" & vbCrLf
    fgList1.CellBackColor = mColor_W1
    
End If

fgList1.Col = 3
Select Case xZCOMPTE0.COMPTELOR
 Case " "
 Case "L": fgList1.Text = "L": fgList1.CellBackColor = mColor_Y1
 Case "N": fgList1.Text = "N": fgList1.CellBackColor = mColor_Y1
 Case Else: fgList1.Text = xZCOMPTE0.COMPTELOR: fgList1.CellBackColor = mColor_W1
End Select

fgList1.Col = 4
Select Case xZCOMPTE0.COMPTESUC
 Case "N"
 Case "O": fgList1.Text = "oui": fgList1.CellBackColor = mColor_Y1
 Case Else: fgList1.Text = xZCOMPTE0.COMPTESUC: fgList1.CellBackColor = mColor_W1
End Select

fgList1.Col = 6
Select Case xZCOMPTE0.COMPTEFON
 Case 0
 Case 1: fgList1.Text = "DB": fgList1.CellBackColor = mColor_Y1
 Case 2: fgList1.Text = "CR": fgList1.CellBackColor = mColor_Y1
 Case 3: fgList1.Text = "DB-CR": fgList1.CellBackColor = mColor_Y1
 Case 4: fgList1.Text = "clos": fgList1.CellBackColor = mColor_Y1
 Case Else: fgList1.Text = xZCOMPTE0.COMPTEFON: fgList1.CellBackColor = mColor_W1
End Select
If xZPLAN0.PLANFONCT <> 0 Then
    If xZCOMPTE0.COMPTEFON <> "4" Then
        If xZCOMPTE0.COMPTEFON <> xZPLAN0.PLANFONCT Then
            mMsgBox_Err = mMsgBox_Err & " - " & xZCOMPTE0.COMPTECOM & " COMPTEFON <> PLANFONCT" & vbCrLf
            fgList1.CellBackColor = mColor_W1
        End If
    End If
End If

If xZCOMPTE0.COMPTEBLO <> 0 Then fgList1.Col = 7: fgList1.Text = dateImp10(xZCOMPTE0.COMPTEBLO + 19000000)

fgList1.Col = 8
Select Case xZCOMPTE0.COMPTESEN
 Case " "
 Case "D": fgList1.Text = "DB": fgList1.CellBackColor = mColor_Y1
 Case "C": fgList1.Text = "CR": fgList1.CellBackColor = mColor_Y1
 Case Else: fgList1.Text = xZCOMPTE0.COMPTESEN: fgList1.CellBackColor = mColor_W1
End Select
If xZCOMPTE0.COMPTESEN <> xZPLAN0.PLANSESOL Then
    mMsgBox_Err = mMsgBox_Err & " - " & xZCOMPTE0.COMPTECOM & " COMPTESEN <> PLANSESOL" & vbCrLf
    fgList1.CellBackColor = mColor_W1
End If

If xZCOMPTE0.COMPTEOUV <> 0 Then fgList1.Col = 9: fgList1.Text = dateImp10(xZCOMPTE0.COMPTEOUV + 19000000)
If xZCOMPTE0.COMPTECLO <> 0 Then
    fgList1.Col = 10: fgList1.Text = dateImp10(xZCOMPTE0.COMPTECLO + 19000000)
Else
    If fctUser_Classe_Aut(xZCOMPTE0.COMPTECLA) Then
        fgList1.Col = 10: fgList1.Text = Format$(-(xZSOLDE0.SOLDECEN), "### ### ### ##0.00")
        If xZSOLDE0.SOLDECEN > 0 Then
            fgList1.CellForeColor = vbRed
        Else
            fgList1.CellForeColor = vbBlue
        End If
    End If
End If

If xZCOMPTE0.COMPTEMOD <> 0 Then fgList1.Col = 11: fgList1.Text = dateImp10(xZCOMPTE0.COMPTEMOD + 19000000)

fgList1.Col = 12: fgList1.Text = xZCOMPTE0.COMPTEINT

On Error Resume Next
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


Public Sub fgList1_Sort()
If fgList1.Rows > 1 Then
    fgList1.Row = 1
    fgList1.RowSel = fgList1.Rows - 1
    
    If fgList1_Sort1_Old = fgList1_Sort1 Then
        If fgList1_SortAD = 5 Then
            fgList1_SortAD = 6
        Else
            fgList1_SortAD = 5
        End If
    Else
        fgList1_SortAD = 5
    End If
    fgList1_Sort1_Old = fgList1_Sort1
    
    fgList1.Col = fgList1_Sort1
    fgList1.ColSel = fgList1_Sort2
    fgList1.Sort = fgList1_SortAD
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

wFct = UCase$(Trim(mId$(Msg, 1, 12)))
Call BiaPgmAut_Init(wFct, SAB_Dossier_Aut)

'blnSetfocus = True
Form_Init


Select Case wFct
    Case "@PCI_COMPTE": blnAuto = True
        Me.Enabled = False: Me.MousePointer = vbHourglass

        cmdSelect_SQL_Update_PCICPTBASE
        cmdSelect_SQL_Update_PCICPTMETA
        
        cmdSelect_SQL_K = "X#"
        cmdSendMail_Surveillance
        
        If mId$(YBIATAB0_DATE_CPT_J, 1, 6) <> mId$(YBIATAB0_DATE_CPT_JS1, 1, 6) Then
        '_________________________________________________________________________
            cmdSelect_SQL_Exportation_PCICPTBASE True, False
            cmdSendMail_PCI_COMPTE
        End If
        
        cmdSelect_SQL_K = "X#c"
        cmdSendMail_Surveillance

        Me.Enabled = True: Me.MousePointer = 0

        Unload Me
    
    Case Else: blnAuto = False
End Select

End Sub


Public Sub Form_Init()
Dim V, xSQL As String, X As String
Dim K As Long

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True


cmdReset
cmdSelect_Clear

blnControl = False

txtPCICPTMETA.Font.Bold = True
txtPCICPTMETA.ForeColor = mColor_Z0
txtPCICPTMETA.BackColor = RGB(0, 192, 192)

txtPCICPTSUFX.Font.Bold = True
txtPCICPTSUFX.ForeColor = mColor_Z0
txtPCICPTSUFX.BackColor = RGB(0, 192, 192)

txtPCICPTTXT.ForeColor = vbBlue
fraDetail.ForeColor = mColor_GB

fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False

fraSelect_Options_1.BorderStyle = 0

lstW.Visible = False

fgDetail_FormatString = fgDetail.FormatString

'SSTab2.Tab = 0
fraList1.Visible = False
fgList1_FormatString = fgList1.FormatString





fraCompte.Visible = False
Set fraCompte.Container = fraTab0
fraCompte.Top = 1440
fraCompte.Left = fraTab0.Left + fraTab0.Width - fraCompte.Width

cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1  - sélection / PCI"
cboSelect_SQL.AddItem "2  - sélection / client"
cboSelect_SQL.AddItem "X#c  - état de surveillance Client"

If SAB_Dossier_Aut.Saisir Then cboSelect_SQL.AddItem "6c  - mise à jour du commentaire"
If SAB_Dossier_Aut.Valider Then cboSelect_SQL.AddItem "6s  - mise à jour de la structure"

If SAB_Dossier_Aut.Rapprocher Then
    cboSelect_SQL.AddItem "Mb  - mise à jour BASE"
    cboSelect_SQL.AddItem "Ms  - mise à jour structure"
End If
If SAB_Dossier_Aut.Comptabiliser Then
    cboSelect_SQL.AddItem "X#  - état de surveillance"
    cboSelect_SQL.AddItem "Xb  - exportation BASE"
    cboSelect_SQL.AddItem "Xbc  - exportation BASE + Comptes"
    cboSelect_SQL.AddItem "Xba  - exportation BASE + anomalies"

End If
If SAB_Dossier_Aut.Xspécial Then cboSelect_SQL.AddItem "JPL  - màj COMPTECLA auto"
cboSelect_SQL.ListIndex = 0
cmdSelect_SQL_K = "1"



lstW.Clear

'Initialisation PLANCOPRO________________________________________________________________________________
cboSelect_PLANCOPRO.Clear
cboSelect_PLANCOPRO.AddItem ""
xSQL = "select distinct PLANCOPRO from " & paramIBM_Library_SAB & ".ZPLAN0 order by PLANCOPRO"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    cboSelect_PLANCOPRO.AddItem Trim(rsSab("PLANCOPRO"))
    rsSab.MoveNext
Loop


'Initialisation devise________________________________________________________________________________
arrDev_Nb = 0
xSQL = "select count(*) from " & paramIBM_Library_SAB & ".ZBASDVS0 "
Set rsSab = cnsab.Execute(xSQL)
arrDev_Nb = rsSab(0)
ReDim Preserve arrDev(arrDev_Nb + 1), arrDev_Num(arrDev_Nb + 1), arrDev_RowT(arrDev_Nb + 1)

cboSelect_DEV.Clear
cboSelect_DEV.AddItem ""
xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASDVS0 order by BASDVSDEV"
Set rsSab = cnsab.Execute(xSQL)
arrDev_Nb = 0
Do While Not rsSab.EOF
    arrDev_Nb = arrDev_Nb + 1
    arrDev(arrDev_Nb) = Trim(rsSab("BASDVSDEV"))
    arrDev_Num(arrDev_Nb) = Trim(rsSab("BASDVSNUM"))
    Select Case arrDev(arrDev_Nb)
        Case "USD": arrDev_Num(arrDev_Nb) = "400"
        Case "CHF": arrDev_Num(arrDev_Nb) = "036"
        Case "AED": arrDev_Num(arrDev_Nb) = "647"
        Case "EUR":
            arrDev_EUR = arrDev_Nb
            arrDev_Num(arrDev_Nb) = "978"
    
    End Select
    cboSelect_DEV.AddItem Trim(rsSab("BASDVSDEV"))
    rsSab.MoveNext
Loop


'___________________________________________________________________________

Call fraCompte_Fiscal("FR")

fraSelect_Options.Visible = True
blnControl = True

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
blnControl = True

End Sub



Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgSelect.Visible = False
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = fgSelect_arrIndex To fgSelect.FixedCols Step -1
        If I <> 1 Then fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        fgSelect.Col = fgSelect_arrIndex
        lColor_Old = fgSelect.CellBackColor
        For I = fgSelect_arrIndex To fgSelect.FixedCols Step -1
          If I <> 1 Then fgSelect.Col = I: fgSelect.CellBackColor = lColor
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
fgDetail.Visible = True: fraDetail.Visible = True
End Sub

Public Sub fgList1_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fraList1.Visible = False: fraDetail.Visible = False
mRow = fgList1.Row

If lRow > 0 And lRow < fgList1.Rows Then
    fgList1.Row = lRow
    For I = fgList1_arrIndex To fgList1.FixedCols Step -1
        fgList1.Col = I: fgList1.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgList1.Row = mRow
    If fgList1.Row > 0 Then
        lRow = fgList1.Row
        lColor_Old = fgList1.CellBackColor
        For I = fgList1_arrIndex To fgList1.FixedCols Step -1
          fgList1.Col = I: fgList1.CellBackColor = lColor
        Next I
    End If
End If
fgList1.LeftCol = fgList1.FixedCols
fraList1.Visible = True: fraDetail.Visible = True
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

























Private Sub cboSelect_DEV_Change()
fgList1_Display
End Sub

Private Sub cboSelect_PLANCOPRO_Change()
cmdSelect_Clear

End Sub

Private Sub cboSelect_PLANCOPRO_Click()
cmdSelect_Clear

End Sub

Private Sub cboSelect_PLANCOPRO_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub cboSelect_SQL_Click()
'cmdSelect_Clear
cmdSelect_Reset
End Sub





Private Sub cboSelect_DEV_Click()
fgList1_Display
End Sub







Private Sub chkSelect_COMPTEFON_Click()
fgList1_Display

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
        Me.PopupMenu mnuPrint, vbPopupMenuLeftButton
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
    Case "1", "6c", "6s": fraSelect_Options.Visible = True: cmdSelect_SQL_1
    Case "2": fraSelect_Options.Visible = True: cmdSelect_SQL_2
    Case "X#", "X#c": cmdSendMail_Surveillance ': cmdSelect_SQL_Surveillance
    Case "Mb": cmdSelect_SQL_Update_PCICPTBASE
    Case "Ms": cmdSelect_SQL_Update_PCICPTMETA
    Case "Xb": cmdSelect_SQL_Exportation_PCICPTBASE False, False
    Case "Xbc": cmdSelect_SQL_Exportation_PCICPTBASE True, True
    Case "Xba": cmdSelect_SQL_Exportation_PCICPTBASE True, False
    Case "JPL": cmdSelect_SQL_JPL
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_Dossier_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
End Sub


Private Sub cmdUpdate_Ok_Click()
Dim V

Me.Enabled = False: Me.MousePointer = vbHourglass

newYPCICPT0 = oldYPCICPT0
newYPCICPT0.PCICPTMETA = Replace(Trim(txtPCICPTMETA), " ", "")
newYPCICPT0.PCICPTSUFX = Trim(txtPCICPTSUFX)
newYPCICPT0.PCICPTTXT = Trim(txtPCICPTTXT)
If optPCICPTAUTO_A Then newYPCICPT0.PCICPTAUTO = " "
If optPCICPTAUTO_M Then newYPCICPT0.PCICPTAUTO = "M"
If optPCICPTAUTO_I Then newYPCICPT0.PCICPTAUTO = "I"


cnSab_Update.Open paramODBC_DSN_SAB
V = sqlYPCICPT0_Update(newYPCICPT0, oldYPCICPT0, True)
If Not IsNull(V) Then
    Call MsgBox(V & vbCrLf & oldYPCICPT0.PCICPTBASE, vbCritical, cmdUpdate_Ok)
End If
cnSab_Update.Close
cmdUpdate_Ok.Visible = False
cmdSelect_Clear
cmdSelect_SQL_1
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub fgList1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next

If y <= fgList1.RowHeightMin Then
Else
    If fgList1.Rows > 1 Then
        Call fgList1_Color(fgList1_RowClick, MouseMoveUsr.BackColor, fgList1_ColorClick)
        fgList1.Col = 1:  xYBIACPT0.COMPTECOM = fgList1.Text
       fraCompte_display xYBIACPT0.COMPTECOM

   End If
End If


End Sub


Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next


If y <= fgDetail.RowHeightMin Then
Else
    If fgDetail.Rows > 1 Then
        Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
        fgDetail.Col = 0:  xZPLAN0.PLANCOOBL = CLng(fgDetail.Text)
        fgList1_Display

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
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select


End Sub

Public Sub cmdContext_Quit()
'blnControl = False
lstErr.Clear: lstErr.Height = 200


If fraCompte.Visible Then
    fraCompte.Visible = False
    Exit Sub
End If

If fraList1.Visible Then
    fraList1.Visible = False
    Exit Sub
End If
If fgDetail.Visible Then
    fgDetail.Visible = False: fraDetail.Visible = False: cmdUpdate_Ok.Visible = False
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
Dim wOrigine As String, xSQL As String
On Error Resume Next


If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  arrYBIACPT0_Index = CLng(fgSelect.Text)
        Select Case cmdSelect_SQL_K
            Case "1", "6c", "6s"
                fgSelect.Col = 0: xYPCICPT0.PCICPTBASE = Trim(fgSelect.Text)
                fgDetail_Display
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




Public Sub fgList1_Reset()
fgList1.Clear
fgList1_Sort1 = 0: fgList1_Sort2 = 0
fgList1_Sort1_Old = -1
fgList1_RowDisplay = 0: fgList1_RowClick = 0
fgList1_arrIndex = fgList1.Cols - 1
blnfgList1_DisplayLine = False
fgList1_SortAD = 6
fgList1.LeftCol = fgList1.FixedCols

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






Public Sub fgSelect_ForeColor(lColor As Long)
For I = 0 To fgSelect_arrIndex
  fgSelect.Col = I: fgSelect.CellForeColor = lColor
Next I

End Sub

































Private Sub mnuPrint_2_Exportation_Click()
Dim xWhere As String, xSQL As String
Dim rsSABY As New ADODB.Recordset

Me.Enabled = False: Me.MousePointer = vbHourglass



Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub txtPCICPTSUFX_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtSelect_Where_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_Where_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub



Public Sub fraCompte_display(lCOMPTECOM As String)
Dim xSQL As String, K As Integer, X As String

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '" & lCOMPTECOM & "'"
Set rsSab = cnsab.Execute(xSQL)

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
        
        If xYBIACPT0.SOLDEDMO = 0 Then
            txtD_SOLDEDMO = ""
        Else
            txtD_SOLDEDMO = dateIBM10(xYBIACPT0.SOLDEDMO, True)
        End If
        txtD_COMPTEDEV = xYBIACPT0.COMPTEDEV
        
        txtD_SOLDECEN = Format$(Abs(xYBIACPT0.SOLDECEN), "### ### ### ##0.00")
        If xYBIACPT0.SOLDECEN > 0 Then
            txtD_SOLDECEN.ForeColor = vbRed
        Else
            txtD_SOLDECEN.ForeColor = vbBlue
        End If
        txtD_SOLDECEN.Visible = fctUser_Classe_Aut(xYBIACPT0.COMPTECLA)
        
        txtD_CLIENACLI = xYBIACPT0.CLIENACLI
        txtD_CLIENASIG = xYBIACPT0.CLIENASIG
        txtD_CLIENARES = xYBIACPT0.CLIENARES
        txtD_CLIENARA1 = xYBIACPT0.CLIENARA1 & " " & xYBIACPT0.CLIENARA2
        txtD_CLIENANAT = xYBIACPT0.CLIENANAT
        txtD_CLIENARSD = xYBIACPT0.CLIENARSD
        txtD_CLIENARSD_Fiscal = fraCompte_Fiscal(Trim(xYBIACPT0.CLIENARSD))
        
        
        fraCompte.Visible = True
    End If
End If
End Sub


Public Sub cmdSelect_SQL_Update_PCICPTBASE()
Dim V, X As String
Dim xSQL As String
Dim xWhere As String
Dim K As Long, blnExiste As Boolean, blnEnAttente As Boolean

xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YPCICPT0"
Set rsSab = cnsab.Execute(xSQL)

arrYPCICPT0_Max = rsSab(0): arrYPCICPT0_Nb = 0
ReDim arrYPCICPT0(arrYPCICPT0_Max + 1)

xSQL = "select PCICPTBASE,PCICPTLEN from " & paramIBM_Library_SABSPE & ".YPCICPT0 order by PCICPTBASE"

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    arrYPCICPT0_Nb = arrYPCICPT0_Nb + 1
    arrYPCICPT0(arrYPCICPT0_Nb).PCICPTBASE = rsSab("PCICPTBASE")
    arrYPCICPT0(arrYPCICPT0_Nb).PCICPTLEN = rsSab("PCICPTLEN")
    rsSab.MoveNext
Loop

'_____________________________________________________________________________________
cnSab_Update.Open paramODBC_DSN_SAB

oldZPLAN0.PLANCOOBL = ""
rsYPCICPT0_Init newYPCICPT0
blnEnAttente = False

xSQL = "select * from " & paramIBM_Library_SAB & ".ZPLAN0 order by PLANCOOBL"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    xZPLAN0.PLANCOOBL = rsSab("PLANCOOBL")
    'If mId$(xZPLAN0.PLANCOOBL, 1, 5) = "38845" Then
    '    Debug.Print "38845"
    'End If
    blnExiste = False
    For K = 1 To arrYPCICPT0_Nb
        If arrYPCICPT0(K).PCICPTBASE = xZPLAN0.PLANCOOBL Then
            blnExiste = True: Exit For
        Else
            If mId$(arrYPCICPT0(K).PCICPTBASE, 1, arrYPCICPT0(K).PCICPTLEN) = mId$(xZPLAN0.PLANCOOBL, 1, arrYPCICPT0(K).PCICPTLEN) Then
                blnExiste = True: Exit For
            End If
        End If
    Next K
    If Not blnExiste Then
        Call rsZPLAN0_GetBuffer(rsSab, xZPLAN0)
        xZPLAN0.PLANINTIT = Text_LCase(xZPLAN0.PLANINTIT)
        
        If Not blnEnAttente Then
            blnEnAttente = True
            oldZPLAN0 = xZPLAN0
            newYPCICPT0.PCICPTBASE = oldZPLAN0.PLANCOOBL
            newYPCICPT0.PCICPTLEN = 6
            newYPCICPT0.PCICPTLNK = oldZPLAN0.PLANCOOBL
        Else
            If mId$(oldZPLAN0.PLANCOOBL, 1, 5) = mId$(xZPLAN0.PLANCOOBL, 1, 5) _
            And oldZPLAN0.PLANINTIT = xZPLAN0.PLANINTIT _
            And oldZPLAN0.PLANCARAC = xZPLAN0.PLANCARAC Then
                newYPCICPT0.PCICPTBASE = mId$(oldZPLAN0.PLANCOOBL, 1, 5)
                newYPCICPT0.PCICPTLEN = 5
            Else
                 V = sqlYPCICPT0_Insert(newYPCICPT0)

                oldZPLAN0 = xZPLAN0
                newYPCICPT0.PCICPTBASE = oldZPLAN0.PLANCOOBL
                newYPCICPT0.PCICPTLEN = 6
                newYPCICPT0.PCICPTLNK = oldZPLAN0.PLANCOOBL
            End If
        End If
        
    End If
    rsSab.MoveNext
Loop

 If blnEnAttente Then V = sqlYPCICPT0_Insert(newYPCICPT0)

cnSab_Update.Close
Set cnSab_Update = Nothing

Set rsSab = Nothing



End Sub
Public Sub cmdSelect_SQL_Update_PCICPTMETA()
Dim V, X As String
Dim xSQL As String
Dim xWhere As String
Dim K As Long, K1 As Long, blnInit As Boolean
Dim xPCICPTMETA As String, xPCICBASE As String, xCOMPTECOM As String, xCLIENACLI As String, xCOMPTEDEV As String
Dim xPLANCARAC As Integer
Dim arrPCICPTSUFX(300) As String, arrPCICPTSUFX_NB As Integer, arrPCICPTSUFX_K As Integer
Dim PCICPTSUFX_K1 As Integer, PCICPTSUFX_L As Integer, blnPCICPTSUFX As Boolean, blnPCICPTSUFX_New As Boolean
Dim xPCICPTSUFX As String
Dim blnErr As Boolean

Call lstErr_Clear(lstErr, cmdContext, "> cmdSelect_SQL_Update_PCICPTMETA"): DoEvents

xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YPCICPT0 " 'where PCICPTMETA = ''"
Set rsSab = cnsab.Execute(xSQL)

arrYPCICPT0_Max = rsSab(0): arrYPCICPT0_Nb = 0
ReDim arrYPCICPT0(arrYPCICPT0_Max + 1)

'xSql = "select * from " & paramIBM_Library_SABSPE & ".YPCICPT0 where PCICPTMETA = '' order by PCICPTBASE"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YPCICPT0  order by PCICPTBASE"

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    arrYPCICPT0_Nb = arrYPCICPT0_Nb + 1
    Call rsYPCICPT0_GetBuffer(rsSab, arrYPCICPT0(arrYPCICPT0_Nb))
    rsSab.MoveNext
Loop

'_____________________________________________________________________________________

For arrYPCICPT0_Index = 1 To arrYPCICPT0_Nb
    blnInit = False: blnErr = False
    arrPCICPTSUFX_NB = 0: xPCICPTSUFX = ""
    xYPCICPT0 = arrYPCICPT0(arrYPCICPT0_Index)
    xPCICBASE = Trim(xYPCICPT0.PCICPTBASE)
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "> " & xPCICBASE): DoEvents

    xSQL = "select * from " & paramIBM_Library_SAB & ".ZPLAN0 where PLANCOOBL = '" & Trim(xYPCICPT0.PCICPTLNK) & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        xPLANCARAC = rsSab("PLANCARAC")
        xPCICPTMETA = String$(xPLANCARAC, "?")
    Else
        xPLANCARAC = 20
        xPCICPTMETA = String$(20, "?")
    End If
    
    'If xPCICBASE = "25302" Then
    '    Debug.Print xPCICBASE
    'End If
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0" _
         & " where COMPTEOBL like '" & xPCICBASE & "%' and COMPTEFON <> '4' order by COMPTECOM"
    Set rsSab = cnsab.Execute(xSQL)
    
    Do While Not rsSab.EOF
        xCOMPTECOM = rsSab("COMPTECOM")
        If Not blnInit Or blnErr Then
            blnInit = True
            Select Case mId$(xCOMPTECOM, 1, 1)
                Case "R": Mid$(xPCICPTMETA, 1, 1) = "R"
                Case "N": Mid$(xPCICPTMETA, 1, 1) = "N"
            End Select
            '________________________________________________________________
            
            K = InStr(xCOMPTECOM, xPCICBASE)
            If K > 0 Then
                Mid$(xPCICPTMETA, K, xYPCICPT0.PCICPTLEN) = String$(xYPCICPT0.PCICPTLEN, "*") ' xPCICBASE
                If mId$(xCOMPTECOM, 1, 1) = "R" And xYPCICPT0.PCICPTLEN = 5 Then Mid$(xPCICPTMETA, K + xYPCICPT0.PCICPTLEN, 1) = "#"
            Else
                If xYPCICPT0.PCICPTLEN > 5 Then
                    X = mId$(xPCICBASE, 1, 5)
                    K = InStr(xCOMPTECOM, X)
                    If K > 0 Then
                        Mid$(xPCICPTMETA, K, 5) = "*****"
                        If mId$(xCOMPTECOM, K + 5, 1) = "0" Then Mid$(xPCICPTMETA, K + 5, 1) = "0"
                    End If
                End If
                If K = 0 Then
                    K = InStr(xCOMPTECOM, mId$(xPCICBASE, 1, 4))
                    If K > 0 Then Mid$(xPCICPTMETA, K, 4) = "****"
                End If
                If K = 0 Then
                    K = InStr(xCOMPTECOM, mId$(xPCICBASE, 1, 3))
                    If K > 0 Then Mid$(xPCICPTMETA, K, 3) = "***"
                End If
            End If
            If mId$(xPCICPTMETA, 1, 6) = "*****?" And mId$(xCOMPTECOM, 6, 1) = "0" Then Mid$(xPCICPTMETA, 6, 1) = "0"
            '________________________________________________________________
            xCOMPTEDEV = rsSab("COMPTEDEV")
            K = InStr(xCOMPTECOM, xCOMPTEDEV)
            If K > 0 Then
                Mid$(xPCICPTMETA, K, 3) = "$$$"

                X = mId$(xCOMPTECOM, K + 3, 3)
                If X = "EUR" Or X = "AED" Or X = "USD" Or X = "RES" Then Mid$(xPCICPTMETA, K + 3, 3) = "<=>"
            Else
                K = InStr(xCOMPTECOM, "CVD")
                If K > 0 Then
                    Mid$(xPCICPTMETA, K, 3) = "$$$"
                Else
                    For K1 = 1 To arrDev_Nb
                        If arrDev(K1) = rsSab("COMPTEDEV") Then
                            K = InStr(xCOMPTECOM, arrDev_Num(K1))
                            If K > 0 Then Mid$(xPCICPTMETA, K, 3) = "§§§"
                            Exit For
                        End If
                    Next K1
                End If
                
            End If
            '________________________________________________________________
            xCLIENACLI = Trim(rsSab("CLIENACLI"))
            If xCLIENACLI <> "" Then
                K = InStr(xCOMPTECOM, mId$(xCLIENACLI, 3, 5))
                If K > 0 Then Mid$(xPCICPTMETA, K, 5) = "ccccc"
            End If
'_____________________________________________________________________________________________________________
            Select Case rsSab("PLANCOPRO")
                Case "CAV", "LOR", "LOB", "LIE", "DOR", "DTT", "CBO", "LDT", "LDX": Mid$(xPCICPTMETA, 9, 3) = "+++"
                Case "NOS", "NOB", "BDF": Mid$(xPCICPTMETA, 10, 3) = "+++"
                Case "PRO", "CHA": Mid$(xPCICPTMETA, 10, 6) = "#act++"
                Case "CHB":
                            If mId$(xPCICPTMETA, 11, 5) = "?????" Then Mid$(xPCICPTMETA, 11, 5) = "ccccc"
                            'If mId$(xPCICBASE, 1, 2) = "70" Then Mid$(xPCICPTMETA, 10, 1) = "#"
                Case "ICC":
                            If mId$(xPCICPTMETA, 1, 1) <> "R" Then
                                If mId$(xPCICPTMETA, 10, 4) = "????" Then
                                    Mid$(xPCICPTMETA, 10, 4) = "#act"
                                    If mId$(xPCICPTMETA, 14, 3) = "???" Then
                                        If mId$(xCOMPTECOM, 14, 3) = "LOR" Or mId$(xCOMPTECOM, 14, 3) = "NOS" Then
                                            Mid$(xPCICPTMETA, 14, 3) = "nol"
                                        Else
                                            Mid$(xPCICPTMETA, 14, 3) = "nat"
                                        End If
                                    End If
                               End If
                            End If
                            
                                
                Case "IMP": Mid$(xPCICPTMETA, 9, 12) = "nat(dossier)"
                Case "CCR", "PBD", "PBI", "PKD", "PKI"
                            If mId$(xPCICPTMETA, 10, 7) = "???????" Then Mid$(xPCICPTMETA, 10, 7) = "#actnat"
            End Select
            
            Select Case xPCICBASE
                  Case "101011", "101101", "101111": Mid$(xPCICPTMETA, 6, 1) = "0"
                  
                  Case "13224": Mid$(xPCICPTMETA, 1, 7) = "R25301#"
                  Case "162301": Mid$(xPCICPTMETA, 1, 7) = "R16210#"
                  Case "26215": Mid$(xPCICPTMETA, 1, 7) = "R26210#"
                  Case "206179": Mid$(xPCICPTMETA, 1, 7) = "R20616#"
                  Case "29150": Mid$(xPCICPTMETA, 1, 7) = "R20610#"
                  Case "673010": Mid$(xPCICPTMETA, 1, 6) = "673090"
                  Case "388930": Mid$(xPCICPTMETA, 1, 15) = "388940$$$opénat"
                  Case "388931": Mid$(xPCICPTMETA, 1, 13) = "388opé$$$++++"
                  Case "388991": Mid$(xPCICPTMETA, 1, 16) = "388940$$$opé???"
                  Case "519111": Mid$(xPCICPTMETA, 1, 7) = "R51910#"
                  Case "93600": Mid$(xPCICPTMETA, 1, 6) = "936001"
                  Case "702600": Mid$(xPCICPTMETA, 1, 6) = "****90"
                  Case "703400": Mid$(xPCICPTMETA, 1, 15) = "703420$$$#act++"
                  Case "977100": Mid$(xPCICPTMETA, 10, 6) = "??????"
                  Case "977700": Mid$(xPCICPTMETA, 10, 6) = "#+++++"
                  Case "978100": Mid$(xPCICPTMETA, 10, 7) = "#act???"
                  Case "985011": Mid$(xPCICPTMETA, 10, 6) = "#act??"
                  Case "999990": Mid$(xPCICPTMETA, 1, 13) = "999999$$$#opé"
                  Case "987560": Mid$(xPCICPTMETA, 11, 5) = "ccccc"
                  Case "388941", "388971", "388981": Mid$(xPCICPTMETA, 10, 4) = "1opé"
                  Case "901196", "901199": Mid$(xPCICPTMETA, 16, 3) = "nat"
                  Case "901196", "901199": Mid$(xPCICPTMETA, 16, 3) = "nat"
                  
                  Case "388101", "388111", "38830", "38831", "38832", "38841", "38842": Mid$(xPCICPTMETA, 10, 7) = "#actopé"
                  Case "977000": Mid$(xPCICPTMETA, 10, 4) = "1nat"
                 
                  Case "16211": Mid$(xPCICPTMETA, 9, 6) = "+++SMF"
                  
                  Case "29111", "29711", "291131": Mid$(xPCICPTMETA, 9, 3) = "+++"
                  
                  Case "191529", "29112", "291169", "291521", "291529", "49700", "206169", "291561", "391529", "161109"
                                Mid$(xPCICPTMETA, 9, 12) = "nat(dossier)"
                  
                  Case "365631", "365681": Mid$(xPCICPTMETA, 10, 4) = "#+++"
                  
                  Case "36566": Mid$(xPCICPTMETA, 10, 1) = "#"
                  
                  Case "38820", "38840": Mid$(xPCICPTMETA, 10, 7) = "#actopé"
                  
                  Case "388941": Mid$(xPCICPTMETA, 9, 4) = "#nat"
                  
                  Case "90319": Mid$(xPCICPTMETA, 16, 3) = "nat"
                  
                  Case "978000": Mid$(xPCICPTMETA, 10, 7) = "#actnat"
                  
                  Case "98520": Mid$(xPCICPTMETA, 10, 4) = "#nat"
                  
                  Case "999990", "999999": Mid$(xPCICPTMETA, 10, 4) = "#opé"
            End Select
 '_____________________________________________________________________________________________________________
            If Not blnErr Then
                If mId$(xPCICPTMETA, xPLANCARAC, 1) <> "?" Then
                       blnPCICPTSUFX = False
                Else
                   blnPCICPTSUFX = True
                   For K = xPLANCARAC To 1 Step -1
                       If mId$(xPCICPTMETA, K, 1) <> "?" Then PCICPTSUFX_K1 = K + 1: Exit For
                   Next K
                   PCICPTSUFX_L = xPLANCARAC - PCICPTSUFX_K1 + 1
                   Mid$(xPCICPTMETA, PCICPTSUFX_K1, PCICPTSUFX_L) = String(PCICPTSUFX_L, "_")
                End If
            End If
        End If
        
        If blnPCICPTSUFX Then
            blnPCICPTSUFX_New = True
            X = Trim(mId$(xCOMPTECOM, PCICPTSUFX_K1, PCICPTSUFX_L))
            For K = 1 To arrPCICPTSUFX_NB
                If arrPCICPTSUFX(K) = X Then blnPCICPTSUFX_New = False: Exit For
            Next K
            If blnPCICPTSUFX_New Then
                arrPCICPTSUFX_NB = arrPCICPTSUFX_NB + 1
                arrPCICPTSUFX(arrPCICPTSUFX_NB) = X
                xPCICPTSUFX = xPCICPTSUFX & X & " "
            End If
        End If
        
        If InStr(xPCICPTMETA, "??") > 0 Then
            blnErr = True
        Else
            blnErr = False
        End If
        
        rsSab.MoveNext
    Loop

    If blnInit Then
        'If arrPCICPTSUFX_NB = 1 Then
        '    Mid$(xPCICPTMETA, PCICPTSUFX_K1, PCICPTSUFX_L) = xPCICPTSUFX
        'End If
        
        If Trim(arrYPCICPT0(arrYPCICPT0_Index).PCICPTMETA) <> Trim(xPCICPTMETA) _
        Or Trim(arrYPCICPT0(arrYPCICPT0_Index).PCICPTSUFX) <> Trim(xPCICPTSUFX) Then
       
            arrYPCICPT0(arrYPCICPT0_Index).PCICPTUUSR = ""
            arrYPCICPT0(arrYPCICPT0_Index).PCICPTMETA = xPCICPTMETA
            If Len(xPCICPTSUFX) > 64 Then
                arrYPCICPT0(arrYPCICPT0_Index).PCICPTSUFX = mId$(xPCICPTSUFX, 1, 64)
            Else
                arrYPCICPT0(arrYPCICPT0_Index).PCICPTSUFX = xPCICPTSUFX
            End If
            
        End If
    End If
Next arrYPCICPT0_Index
'_____________________________________________________________________________________
cnSab_Update.Open paramODBC_DSN_SAB
For arrYPCICPT0_Index = 1 To arrYPCICPT0_Nb

    If arrYPCICPT0(arrYPCICPT0_Index).PCICPTAUTO = " " Then
        If Trim(arrYPCICPT0(arrYPCICPT0_Index).PCICPTUUSR) = "" Then
            oldYPCICPT0 = arrYPCICPT0(arrYPCICPT0_Index)
            oldYPCICPT0.PCICPTMETA = "màj à faire"
            oldYPCICPT0.PCICPTSUFX = "màj à faire"
            V = sqlYPCICPT0_Update(arrYPCICPT0(arrYPCICPT0_Index), oldYPCICPT0, True)
            If Not IsNull(V) Then
                Call MsgBox(V & vbCrLf & arrYPCICPT0(arrYPCICPT0_Index).PCICPTSUFX, vbCritical, oldYPCICPT0.PCICPTBASE)
            End If
            
        End If
    End If

Next arrYPCICPT0_Index

cnSab_Update.Close

'_____________________________________________________________________________________


Set rsSab = Nothing



End Sub


Public Sub cmdSelect_SQL_Surveillance()
 
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wAMJMin As String, WAMJMax As String
Dim X As String, K As Long, K1 As Long, K2 As Long
Dim blnCALCS As Boolean
On Error GoTo Error_Handler
'===================================================================================
If blnAuto Then
    X = paramServer("\\CPT_Archive\")
Else
    X = ""
End If
If X = "" Then X = "C:\Temp\"
If mId$(X, Len(X), 1) <> "\" Then X = X & "\"
blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

If cmdSelect_SQL_K = "X#c" Then
    wFile = X & Trim("CPT surveillance des clients au " & dateImp_Amj(YBIATAB0_DATE_CPT_J) & ".xlsx")
Else
    wFile = X & Trim("CPT surveillance du plan comptable au " & dateImp_Amj(YBIATAB0_DATE_CPT_J) & ".xlsx")
End If

If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "Surveillance du plan comptable : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If


If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile
'====================================================================================


Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "CPT"
    .Subject = ""
End With

'Set wsExcel = wbExcel.ActiveSheet
Set wsExcel = wbExcel.Sheets(1)
wsExcel.Name = "CPT récap " & dateImp10(YBIATAB0_DATE_CPT_J)
cmdSelect_SQL_Surveillance_Init
mXls1_Row = 1
'__________________________________________________________________________________

Set wsExcel = wbExcel.Sheets(2)
wsExcel.Name = "CPT détail " & dateImp10(YBIATAB0_DATE_CPT_J)
cmdSelect_SQL_Surveillance_Init
mXls2_Row = 1

'===========================================================================

Call lstErr_AddItem(lstErr, cmdContext, "> initialisation : ZPLAN0"): DoEvents


X = "select count(*) from " & paramIBM_Library_SAB & ".ZPLAN0 "
Set rsSab = cnsab.Execute(X)
If Not rsSab.EOF Then
    ReDim arrZPLAN0(rsSab(0) + 1)
    ReDim arrZPLAN0_Lnk(rsSab(0) + 1)
Else
    ReDim arrZPLAN0(1)
    ReDim arrZPLAN0_Lnk(1)
End If

arrZPLAN0_Nb = 0
X = "select * from " & paramIBM_Library_SAB & ".ZPLAN0 " _
       & " order by PLANETABL,PLANPLAN,PLANCOOBL"
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    arrZPLAN0_Nb = arrZPLAN0_Nb + 1
    V = rsZPLAN0_GetBuffer(rsSab, arrZPLAN0(arrZPLAN0_Nb))
    arrZPLAN0_Lnk(arrZPLAN0_Nb) = 0
    rsSab.MoveNext
Loop
'__________________________________________________________________________________

Call lstErr_AddItem(lstErr, cmdContext, "> initialisation : YPCICPT0"): DoEvents

X = "select count(*) from " & paramIBM_Library_SABSPE & ".YPCICPT0 "
Set rsSab = cnsab.Execute(X)
If Not rsSab.EOF Then
    ReDim arrYPCICPT0(rsSab(0) + 1)
    ReDim arrYPCILNK0_Lnk(rsSab(0) + 1)
Else
    ReDim arrYPCICPT0(1)
    ReDim arrYPCILNK0_Lnk(1)
End If

arrYPCICPT0_Nb = 0
X = "select * from " & paramIBM_Library_SABSPE & ".YPCICPT0 " _
       & " order by PCICPTLNK"
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    arrYPCICPT0_Nb = arrYPCICPT0_Nb + 1
    V = rsYPCICPT0_GetBuffer(rsSab, arrYPCICPT0(arrYPCICPT0_Nb))
    arrYPCILNK0_Lnk(arrYPCICPT0_Nb) = 0
    rsSab.MoveNext
Loop

For K = 1 To arrZPLAN0_Nb
    For K1 = 1 To arrYPCICPT0_Nb
        If arrZPLAN0(K).PLANCOOBL = arrYPCICPT0(K1).PCICPTLNK Then
           arrZPLAN0_Lnk(K) = K1
           arrYPCILNK0_Lnk(K1) = K
           Exit For
        End If
    Next K1
Next K

For K = 1 To arrZPLAN0_Nb
    If arrZPLAN0_Lnk(K) = 0 Then
        For K1 = 1 To arrYPCICPT0_Nb
            If mId$(arrZPLAN0(K).PLANCOOBL, 1, 5) = Trim(arrYPCICPT0(K1).PCICPTBASE) Then
               arrZPLAN0_Lnk(K) = K1
              Exit For
            End If
        Next K1
    End If
Next K

'__________________________________________________________________________________

If cmdSelect_SQL_K = "X#c" Then
    cmdSelect_SQL_Surveillance_COMPTEFON
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

    cmdSelect_SQL_Surveillance_COMPTEOUV
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

    cmdSelect_SQL_Surveillance_ZCLIENA0
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

    cmdSelect_SQL_Surveillance_ZTITULA0
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

    cmdSelect_SQL_Surveillance_ZRELEVE0
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

Else


    cmdSelect_SQL_Surveillance_COMPTEFON
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

    cmdSelect_SQL_Surveillance_YPCICPT0_Update
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

    cmdSelect_SQL_Surveillance_COMPTEOUV
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

    cmdSelect_SQL_Surveillance_ZCLIENA0
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

    cmdSelect_SQL_Surveillance_ZCOMREF0
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

    cmdSelect_SQL_Surveillance_ZTITULA0
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

    cmdSelect_SQL_Surveillance_ZRELEVE0
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

    cmdSelect_SQL_Surveillance_YBIACPT0
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

    cmdSelect_SQL_Surveillance_YBIAMVTH
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

    cmdSelect_SQL_Surveillance_ZPLAN0
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

    cmdSelect_SQL_Surveillance_YPCICPT0
    mXls1_Row = mXls1_Row + 1 ': wsExcel.Rows(mXls1_Row).RowHeight = 10
    'For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(196, 196, 196): Next K

End If

mXls1_Row = mXls1_Row + 1: wsExcel.Cells(mXls1_Row, 1) = "Fin du traitement"
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
'Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents


End Sub

Public Function PCICPTMETA_Display(lX As String) As String
PCICPTMETA_Display = lX
PCICPTMETA_Display = Replace(PCICPTMETA_Display, "act", " act", , 1)
PCICPTMETA_Display = Replace(PCICPTMETA_Display, "nat", " nat", , 1)
PCICPTMETA_Display = Replace(PCICPTMETA_Display, "nol", " nol", , 1)
PCICPTMETA_Display = Replace(PCICPTMETA_Display, "opé", " opé", , 1)
PCICPTMETA_Display = Replace(PCICPTMETA_Display, "<=>", " <=>", , 1)
PCICPTMETA_Display = Replace(PCICPTMETA_Display, "*", " *", , 1)
PCICPTMETA_Display = Replace(PCICPTMETA_Display, "(", " (", , 1)
PCICPTMETA_Display = Replace(PCICPTMETA_Display, "§", " §", , 1)
PCICPTMETA_Display = Replace(PCICPTMETA_Display, "#", " #", , 1)
PCICPTMETA_Display = Replace(PCICPTMETA_Display, "$", " $", , 1)
PCICPTMETA_Display = Replace(PCICPTMETA_Display, "cc", " cc", , 1)
PCICPTMETA_Display = Replace(PCICPTMETA_Display, "+", " +", , 1)
PCICPTMETA_Display = Replace(PCICPTMETA_Display, "_", " _", , 1)
PCICPTMETA_Display = Replace(PCICPTMETA_Display, "?", " ?", , 1)
PCICPTMETA_Display = Trim(PCICPTMETA_Display)
End Function

Public Function PCICPTMETA_Control(wCOMPTECOM As String) As Boolean
'____________________________________
'===> dim xZCOMPTE,xPCICPT0
'____________________________________
Dim X As String, X1 As String, K As Integer, K1 As Integer, K2 As Integer
Dim wCLIENACLI As String, wCLIENARSD As String, wCLIENARSD_Fiscal As String
Dim blnAnomalie As Boolean
On Error GoTo Error_Handler

PCICPTMETA_Control = False
blnAnomalie = False
wCOMPTECOM = xZCOMPTE0.COMPTECOM
'______________________________________________________________________
X1 = mId$(xYPCICPT0.PCICPTMETA, 1, 1)
If X1 = "R" Or X1 = "N" Then
    If X1 = mId$(wCOMPTECOM, 1, 1) Then Mid$(wCOMPTECOM, 1, 1) = " "
End If
'______________________________________________________________________

K = InStr(xYPCICPT0.PCICPTMETA, "******")
If K > 0 Then
    If mId$(xYPCICPT0.PCICPTBASE, 1, 6) = mId$(xZCOMPTE0.COMPTECOM, K, 6) Then Mid$(wCOMPTECOM, K, 6) = "      "
Else
    K = InStr(xYPCICPT0.PCICPTMETA, "*****")
    If K > 0 Then
        If mId$(xYPCICPT0.PCICPTBASE, 1, 5) = mId$(xZCOMPTE0.COMPTECOM, K, 5) Then Mid$(wCOMPTECOM, K, 5) = "     "
    Else
        K = InStr(xYPCICPT0.PCICPTMETA, "****")
        If K > 0 Then
            If mId$(xYPCICPT0.PCICPTBASE, 1, 4) = mId$(xZCOMPTE0.COMPTECOM, K, 4) Then Mid$(wCOMPTECOM, K, 4) = "    "
        Else
            K = InStr(xYPCICPT0.PCICPTMETA, "***")
            If K > 0 Then
                If mId$(xYPCICPT0.PCICPTBASE, 1, 3) = mId$(xZCOMPTE0.COMPTECOM, K, 3) Then Mid$(wCOMPTECOM, K, 3) = "   "
            End If
        End If
    End If

End If
'______________________________________________________________________
K = InStr(xYPCICPT0.PCICPTMETA, "$$$")
If K > 0 Then
    If mId$(wCOMPTECOM, K, 3) = xZCOMPTE0.COMPTEDEV Then
        Mid$(wCOMPTECOM, K, 3) = "   "
    Else
        If mId$(wCOMPTECOM, K, 3) = "CVD" Then
            Mid$(wCOMPTECOM, K, 3) = "   "
        Else
            X1 = mId$(wCOMPTECOM, K, 3)
            For K1 = 1 To arrDev_Nb
                If arrDev_Num(K1) = X1 Then
                    Mid$(wCOMPTECOM, K, 3) = "   "
                    Exit For
                End If
            Next K1

        End If
        
    End If
    
End If
K = InStr(xYPCICPT0.PCICPTMETA, "§§§")
If K > 0 Then
    
    For K1 = 1 To arrDev_Nb
        If arrDev(K1) = xZCOMPTE0.COMPTEDEV Then
            If mId$(wCOMPTECOM, K, 3) = arrDev_Num(K1) Then Mid$(wCOMPTECOM, K, 3) = "   "
            Exit For
        End If
    Next K1
End If

K = InStr(xYPCICPT0.PCICPTMETA, "<=>")
If K > 0 Then
    X1 = mId$(xZCOMPTE0.COMPTECOM, K, 3)
    If X1 = "CVD" Or X1 = "RES" Then
        Mid$(wCOMPTECOM, K, 3) = "   "
    Else
        For K1 = 1 To arrDev_Nb
            If arrDev(K1) = X1 Then
                Mid$(wCOMPTECOM, K, 3) = "   "
                'If mId$(wCOMPTECOM, K, 3) = arrDev_Num(K1) Then Mid$(wCOMPTECOM, K, 3) = "   "
                Exit For
            End If
        Next K1
    End If
End If

'______________________________________________________________________
K = InStr(xYPCICPT0.PCICPTMETA, "ccccc")
If K = 0 Then
'______________________________________________________________________
    K = InStr(xYPCICPT0.PCICPTMETA, "#")
    If K > 0 Then
        If IsNumeric(mId$(wCOMPTECOM, K, 1)) Then Mid$(wCOMPTECOM, K, 1) = " "
    End If
Else
    X = "select * from " & paramIBM_Library_SAB & ".ZTITULA0, " & paramIBM_Library_SAB & ".ZCLIENA0" _
           & "  where TITULACOM = '" & xZCOMPTE0.COMPTECOM & "' and TITULATPR = '0' and TITULACLI = CLIENACLI"
    Set rsSabX = cnsab.Execute(X)
    
    If Not rsSabX.EOF Then
        wCLIENACLI = rsSabX("CLIENACLI")
        If mId$(wCOMPTECOM, K, 5) = mId$(wCLIENACLI, 3, 5) Then
            Mid$(wCOMPTECOM, K, 5) = "     "
        Else
            If wCLIENACLI = "0010000" Then Mid$(wCOMPTECOM, K, 5) = "     "
        End If
        wCLIENARSD = Trim(rsSabX("CLIENARSD"))
        wCLIENARSD_Fiscal = fraCompte_Fiscal(wCLIENARSD)
        K = InStr(xYPCICPT0.PCICPTMETA, "#")
        If K > 0 Then
             If mId$(wCOMPTECOM, K, 1) = wCLIENARSD_Fiscal Then Mid$(wCOMPTECOM, K, 1) = " "
        Else
            If xYPCICPT0.PCICPTLEN = 5 Then
                If mId$(xZCOMPTE0.COMPTEOBL, 6, 1) <> wCLIENARSD_Fiscal Then blnAnomalie = True
            End If
        End If
    Else
    End If
End If
'______________________________________________________________________
K = InStr(xYPCICPT0.PCICPTMETA, "act")
If K > 0 Then
    If mId$(wCOMPTECOM, K, 3) = "000" Or mId$(wCOMPTECOM, K, 3) = "ACT" Then Mid$(wCOMPTECOM, K, 3) = "   "
End If
'______________________________________________________________________
K = InStr(xYPCICPT0.PCICPTMETA, "nol")
If K > 0 Then
    X1 = mId$(wCOMPTECOM, K, 3)
    If X1 = "LOR" Or X1 = "NOS" Or X1 = "LOB" Or X1 = "NOB" Then Mid$(wCOMPTECOM, K, 3) = "   "
End If
'______________________________________________________________________
K = InStr(xYPCICPT0.PCICPTMETA, "nat")
If K > 0 Then Mid$(wCOMPTECOM, K, 3) = "   "
'______________________________________________________________________
K = InStr(xYPCICPT0.PCICPTMETA, "opé")
If K > 0 Then Mid$(wCOMPTECOM, K, 3) = "   "
'______________________________________________________________________
K = InStr(xYPCICPT0.PCICPTMETA, "(dossier)")
If K > 0 Then
    If IsNumeric(mId$(wCOMPTECOM, K, 3)) Then Mid$(wCOMPTECOM, K, 9) = "         "
End If
'______________________________________________________________________

K = InStr(xYPCICPT0.PCICPTMETA, "+++++")
If K > 0 Then
    X1 = mId$(xZCOMPTE0.COMPTECOM, K, 5)
    If X1 >= "00000" And X1 <= "99999" Then Mid$(wCOMPTECOM, K, 5) = "     "
End If
K = InStr(xYPCICPT0.PCICPTMETA, "++++")
If K > 0 Then
    X1 = mId$(xZCOMPTE0.COMPTECOM, K, 4)
    If X1 >= "0000" And X1 <= "9999" Then Mid$(wCOMPTECOM, K, 4) = "    "
End If
K = InStr(xYPCICPT0.PCICPTMETA, "+++")
If K > 0 Then
    X1 = mId$(xZCOMPTE0.COMPTECOM, K, 3)
    If X1 >= "000" And X1 <= "999" Then Mid$(wCOMPTECOM, K, 3) = "   "
Else
    K = InStr(xYPCICPT0.PCICPTMETA, "++")
    If K > 0 Then
        X1 = mId$(xZCOMPTE0.COMPTECOM, K, 2)
        If X1 >= "00" And X1 <= "99" Then Mid$(wCOMPTECOM, K, 2) = "  "
    End If
End If
'______________________________________________________________________

K = InStr(xYPCICPT0.PCICPTMETA, "_")
If K > 0 Then
    K1 = Len(xYPCICPT0.PCICPTMETA) - K + 1
    X1 = Trim(mId$(xZCOMPTE0.COMPTECOM, K, K1))
    K2 = InStr(xYPCICPT0.PCICPTSUFX, X1)
    If K2 > 0 Then Mid$(wCOMPTECOM, K, K1) = String$(K1, " ")
End If

'______________________________________________________________________
If Trim(wCOMPTECOM) <> "" Then
    For K = 1 To Len(xYPCICPT0.PCICPTMETA)
        If mId$(wCOMPTECOM, K, 1) = mId$(xYPCICPT0.PCICPTMETA, K, 1) Then Mid$(wCOMPTECOM, K, 1) = " "
    Next K
End If

If Trim(wCOMPTECOM) = "" Then PCICPTMETA_Control = True
wCOMPTECOM = Replace(wCOMPTECOM, " ", "-")

If blnAnomalie Then
    PCICPTMETA_Control = False
    wCOMPTECOM = "PCI : " & xZCOMPTE0.COMPTEOBL
End If

If wCLIENACLI <> "" Then
    wCOMPTECOM = wCOMPTECOM & " # " & wCLIENACLI
    If wCLIENARSD_Fiscal <> "" Then wCOMPTECOM = wCOMPTECOM & " : " & wCLIENARSD & " => " & wCLIENARSD_Fiscal
End If

Exit Function

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Function

Public Function fraCompte_Fiscal(lPays As String) As String
Static K As Integer

If arrPays_Nb = 0 Then
    Call rsZBASTAB0_Pays(arrPays(), arrPays_Nb)
    K = 0: arrPays(0).Id = "?"
End If
'___________________________________________________________________________
If lPays = arrPays(K).Id Then
    fraCompte_Fiscal = arrPays(K).Fiscal
Else
    fraCompte_Fiscal = ""
    For K = 1 To arrPays_Nb
        If lPays = arrPays(K).Id Then fraCompte_Fiscal = arrPays(K).Fiscal: Exit For
    Next K
    If K > arrPays_Nb Then K = 1
End If
End Function

Public Sub cmdSelect_SQL_Surveillance_Init()
Dim K As Integer

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
    .Font.Size = 7
    .Font.Name = "Arial Unicode MS"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 80
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14CPT : Surveillance du plan comptable en date du " & dateImp10_S(YBIATAB0_DATE_CPT_J) _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$C1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

wsExcel.Columns(1).ColumnWidth = 10: wsExcel.Cells(1, 1) = "Contrôle"
wsExcel.Columns(2).ColumnWidth = 20: wsExcel.Cells(1, 2) = "Identification"
wsExcel.Columns(3).ColumnWidth = 32: wsExcel.Cells(1, 3) = "Intitulé du compte"
wsExcel.Columns(4).ColumnWidth = 15: wsExcel.Cells(1, 4) = "solde ": wsExcel.Columns(4).NumberFormat = "### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(5).ColumnWidth = 80: wsExcel.Cells(1, 5) = "nature du contrôle"

For K = 1 To 5
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next

End Sub

Public Sub cmdSelect_SQL_Exportation_PCICPTBASE_Detail_Compte()
Dim xWhere As String

xWhere = " where COMPTEFON <> '4'"
X = Trim(txtSelect_Where)
If X <> "" Then xWhere = xWhere & " and COMPTEOBL like '" & X & "%'"

X = Trim(cboSelect_PLANCOPRO)
If X <> "" Then xWhere = xWhere & " and COMPTEOBL in (select PLANCOOBL from " & paramIBM_Library_SAB & ".zplan0 where PLANCOPRO = '" & X & "')"

'__________________________________________________________________________________
X = "select count(*) from " & paramIBM_Library_SAB & ".ZCOMPTE0" _
       & xWhere
Set rsSab = cnsab.Execute(X)

ReDim arrZCOMPTE0(rsSab(0) + 1), arrZSOLDE0(rsSab(0) + 1)

'__________________________________________________________________________________

xWhere = xWhere & " and COMPTEETA = SOLDEETA and COMPTEPLA = SOLDEPLA and COMPTECOM = SOLDECOM "

X = "select * from " & paramIBM_Library_SAB & ".ZCOMPTE0, " & paramIBM_Library_SAB & ".ZSOLDE0" _
       & xWhere _
       & " order by COMPTEOBL,COMPTECOM"
Set rsSab = cnsab.Execute(X)

arrZCOMPTE0_Nb = 0
Do While Not rsSab.EOF
    arrZCOMPTE0_Nb = arrZCOMPTE0_Nb + 1
    V = rsZCOMPTE0_GetBuffer(rsSab, arrZCOMPTE0(arrZCOMPTE0_Nb))
    arrZSOLDE0(arrZCOMPTE0_Nb).SOLDEDMO = rsSab("SOLDEDMO")
    arrZSOLDE0(arrZCOMPTE0_Nb).SOLDECEN = rsSab("SOLDECEN")
    rsSab.MoveNext
Loop

End Sub

Public Function arrCLIEANRES_Control(lCLIEANRES As String)
Static mCLIEANRES_K As Integer

arrCLIEANRES_Control = Null

If arrCLIEANRES_Nb = 0 Then
    Dim X As String
    X = "select count(*) from " & paramIBM_Library_SAB & ".ZBASTAB0 " _
       & " where BASTABETA = 1 and BASTABNUM = 6 "
    Set rsSabX = cnsab.Execute(X)
    ReDim arrCLIEANRES(rsSabX(0) + 1)
    
    arrCLIEANRES_Nb = 0
    
    X = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 " _
       & " where BASTABETA = 1 and BASTABNUM = 6 order by BASTABARG "
        Set rsSabX = cnsab.Execute(X)
        
    Do While Not rsSabX.EOF
        arrCLIEANRES_Nb = arrCLIEANRES_Nb + 1
        arrCLIEANRES(arrCLIEANRES_Nb) = mId$(rsSabX("BASTABARG"), 4, 3)
        
        rsSabX.MoveNext
    Loop
    mCLIEANRES_K = 0
End If
If arrCLIEANRES(mCLIEANRES_K) = lCLIEANRES Then
    arrCLIEANRES_Control = mCLIEANRES_K
Else
    For mCLIEANRES_K = 1 To arrCLIEANRES_Nb
       If arrCLIEANRES(mCLIEANRES_K) = lCLIEANRES Then arrCLIEANRES_Control = mCLIEANRES_K: Exit For
    Next mCLIEANRES_K
End If

End Function

Public Sub cmdSelect_SQL_Surveillance_ZRELEVE0()

Dim X As String
Dim K As Long, wIBMMin As Long
Dim Nb_Err_RELEVECOM As Long
Dim blnOk As Boolean
 On Error GoTo Error_Handler

'==============================================================================
Call lstErr_AddItem(lstErr, cmdContext, "> contrôle : ZRELEVE0"): DoEvents
Set wsExcel = wbExcel.Sheets(1)

rsZCOMPTE0_Init oldZCOMPTE0
rsZRELEVE0_Init oldZRELEVE0
blnOk = True

'=======================================================================================================
X = "select * from " & paramIBM_Library_SAB & ".ZCOMPTE0 " _
  & " left outer join " & paramIBM_Library_SAB & ".ZPLAN0 on compteobl = plancoobl" _
  & " left outer join " & paramIBM_Library_SAB & ".ztitula0 on comptecom = titulacom" _
  & " left outer join " & paramIBM_Library_SAB & ".zreleve0 on relevecom = comptecom" _
  & " where comptefon <> '4'  and plancopro in ('CAV','DOR','CBO','LOR','LOB','LIE')" _
  & " and RELEVEREL <> 'W' order by comptecom, titulacli"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    
'___________________________________________________________________________________________
    xZCOMPTE0.COMPTECOM = rsSab("COMPTECOM")
    If oldZCOMPTE0.COMPTECOM <> xZCOMPTE0.COMPTECOM Then
        If Not blnOk Then
            mXls1_Row = mXls1_Row + 1: Nb_Err_RELEVECOM = Nb_Err_RELEVECOM + 1
            wsExcel.Cells(mXls1_Row, 1) = "ZRELEVE0"
            wsExcel.Cells(mXls1_Row, 2) = oldZCOMPTE0.COMPTECOM
            wsExcel.Cells(mXls1_Row, 3) = oldZCOMPTE0.COMPTEINT
            wsExcel.Cells(mXls1_Row, 5) = oldZCOMPTE0.COMPTEDEV & " - le destinataire du relevé n'est pas un titulaire # " & oldZRELEVE0.RELEVENUM
            If Not blnOk Then For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0: Next K
        End If
        V = rsZCOMPTE0_GetBuffer(rsSab, oldZCOMPTE0)
        blnOk = False
    End If
    V = rsSab("RELEVENUM")
    If IsNull(V) Then
            mXls1_Row = mXls1_Row + 1: Nb_Err_RELEVECOM = Nb_Err_RELEVECOM + 1
            wsExcel.Cells(mXls1_Row, 1) = "ZRELEVE0"
            wsExcel.Cells(mXls1_Row, 2) = oldZCOMPTE0.COMPTECOM
            wsExcel.Cells(mXls1_Row, 3) = oldZCOMPTE0.COMPTEINT
            wsExcel.Cells(mXls1_Row, 5) = oldZCOMPTE0.COMPTEDEV & " - destinataire du relevé non précisé"
            If Not blnOk Then For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0: Next K
    Else
        oldZRELEVE0.RELEVENUM = Trim(V)
        Select Case rsSab("RELEVETYP")
            Case "1":
                V = rsSab("TITULACLI")
                If IsNull(V) Then
                    mXls1_Row = mXls1_Row + 1: Nb_Err_RELEVECOM = Nb_Err_RELEVECOM + 1
                    wsExcel.Cells(mXls1_Row, 1) = "ZRELEVE0"
                    wsExcel.Cells(mXls1_Row, 2) = oldZCOMPTE0.COMPTECOM
                    wsExcel.Cells(mXls1_Row, 3) = oldZCOMPTE0.COMPTEINT
                    wsExcel.Cells(mXls1_Row, 5) = oldZCOMPTE0.COMPTEDEV & " - il n'y a pas de titulaire pour ce compte"
                    If Not blnOk Then For K = 1 To 5: wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_W0: Next K
                Else
                    If Trim(oldZRELEVE0.RELEVENUM) = Trim(V) Then blnOk = True
                End If
            Case Else
                    If oldZCOMPTE0.COMPTECOM = oldZRELEVE0.RELEVENUM Then blnOk = True
        End Select
    End If
'___________________________________________________________________________________________

        
    rsSab.MoveNext
Loop
'_________________________________________________________________________________________________________
Set wsExcel = wbExcel.Sheets(1)
mXls1_Row = mXls1_Row + 1
wsExcel.Cells(mXls1_Row, 1) = "ZRELEVE0"
If Nb_Err_RELEVECOM = 0 Then
    wsExcel.Cells(mXls1_Row, 5) = "* pas d'anomalies de destinataire de relevé de compte "
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(96, 190, 255)
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
Else
    wsExcel.Cells(mXls1_Row, 5) = "* " & Nb_Err_RELEVECOM & " anomalies de destinataire de relevé de compte"
    For K = 1 To 5
        wsExcel.Cells(mXls1_Row, K).Interior.Color = RGB(96, 190, 255)
        wsExcel.Cells(mXls1_Row, K).Font.Bold = True
    Next K
End If

Exit Sub

Error_Handler:
If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Public Sub cmdSelect_SQL_JPL()
Dim xSQL As String, Nb As Integer, Nb_Maj As Integer
Dim rsADO As ADODB.Recordset
Call lstErr_Clear(lstErr, cmdPrint, "cmdSelect_SQL_JPL ..... ")

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCOMPTE0, " & paramIBM_Library_SAB & ".ZPLAN0 " _
     & " where COMPTEOBL = PLANCOOBL  order by COMPTECOM"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    
    If rsSab("COMPTECLA") < rsSab("PLANCLASS") Then
        If rsSab("PLANCLASS") < 9 Then
            Nb = Nb + 1
            
            'Debug.Print Nb, rsSab("COMPTECOM"), rsSab("COMPTECLA"), rsSab("PLANCLASS")
            xSQL = "update " & paramIBM_Library_SAB & ".ZCOMPTE0 set COMPTECLA = " & rsSab("PLANCLASS") _
                 & " where COMPTEETA = " & rsSab("COMPTEETA") & " and COMPTEPLA = " & rsSab("COMPTEPLA") _
                 & " and COMPTECOM ='" & rsSab("COMPTECOM") & "'"
            Call FEU_ROUGE
            Set rsADO = cnsab.Execute(xSQL, Nb_Maj)
            Call FEU_VERT
            If Nb_Maj <> 1 Then
                Debug.Print "# " & Nb, rsSab("COMPTECOM"), rsSab("COMPTECLA"), rsSab("PLANCLASS")
            End If
            If Nb Mod 100 = 0 Then
                Debug.Print
            End If
        End If
    End If
    rsSab.MoveNext
Loop

Call lstErr_Clear(lstErr, cmdPrint, "cmdSelect_SQL_JPL : " & Nb)


End Sub
