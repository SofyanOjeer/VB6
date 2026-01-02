VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAB_Dossier_DB 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_Dossier: base de données"
   ClientHeight    =   10305
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SAB_Dossier_DB.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10305
   ScaleWidth      =   13530
   Begin VB.Frame fraSelect 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9660
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   13350
      Begin TabDlg.SSTab SSTab1 
         Height          =   9015
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   15901
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Relevé comptable du dossier"
         TabPicture(0)   =   "SAB_Dossier_DB.frx":030A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fgBIAMVT"
         Tab(0).Control(1)=   "fraCompte"
         Tab(0).Control(2)=   "fgCPTPIE"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Extrait de compte"
         TabPicture(1)   =   "SAB_Dossier_DB.frx":0326
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "fgExtrait"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "fgCPTPIE_2"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "."
         TabPicture(2)   =   "SAB_Dossier_DB.frx":0342
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtRTF"
         Tab(2).Control(1)=   "txtFg"
         Tab(2).ControlCount=   2
         Begin MSFlexGridLib.MSFlexGrid fgCPTPIE 
            Height          =   3000
            Left            =   -74880
            TabIndex        =   9
            ToolTipText     =   "Cliquer pour obtenir la fiche compte et un extrait"
            Top             =   5730
            Visible         =   0   'False
            Width           =   12795
            _ExtentX        =   22569
            _ExtentY        =   5292
            _Version        =   393216
            Cols            =   11
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16777215
            ForeColor       =   4210752
            BackColorFixed  =   8438015
            ForeColorFixed  =   16384
            BackColorBkg    =   -2147483633
            AllowUserResizing=   3
            FormatString    =   $"SAB_Dossier_DB.frx":035E
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
         Begin MSFlexGridLib.MSFlexGrid fgCPTPIE_2 
            Height          =   3000
            Left            =   120
            TabIndex        =   35
            ToolTipText     =   "cliquer ici pour afficher les écritures comptables du dossier sélectionné"
            Top             =   5490
            Visible         =   0   'False
            Width           =   12795
            _ExtentX        =   22569
            _ExtentY        =   5292
            _Version        =   393216
            Cols            =   11
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16777215
            ForeColor       =   4210752
            BackColorFixed  =   8438015
            ForeColorFixed  =   16384
            BackColorBkg    =   -2147483633
            AllowUserResizing=   3
            FormatString    =   $"SAB_Dossier_DB.frx":048B
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
         Begin VB.Frame fraCompte 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Compte"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4575
            Left            =   -68865
            TabIndex        =   10
            Top             =   990
            Visible         =   0   'False
            Width           =   6345
            Begin VB.Frame Frame1 
               BackColor       =   &H0080C0FF&
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
               Height          =   585
               Left            =   180
               TabIndex        =   30
               Top             =   3900
               Width           =   6000
               Begin VB.CommandButton cmdD_Extrait 
                  BackColor       =   &H0080FF80&
                  Caption         =   "Extrait de compte du ... au ...."
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   500
                  Left            =   375
                  MaskColor       =   &H80000000&
                  Style           =   1  'Graphical
                  TabIndex        =   32
                  ToolTipText     =   "cliquer ici pour afficher un extrait de compte de la période"
                  Top             =   60
                  Width           =   1380
               End
               Begin MSComCtl2.DTPicker txtD_Extrait_AMJMIn 
                  Height          =   300
                  Left            =   2490
                  TabIndex        =   31
                  Top             =   180
                  Width           =   1290
                  _ExtentX        =   2275
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
                  Format          =   105709571
                  CurrentDate     =   36299
                  MaxDate         =   401768
                  MinDate         =   -328351
               End
               Begin MSComCtl2.DTPicker txtD_Extrait_AMJMax 
                  Height          =   300
                  Left            =   4185
                  TabIndex        =   33
                  Top             =   195
                  Width           =   1320
                  _ExtentX        =   2328
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
                  Format          =   105709571
                  CurrentDate     =   36299
                  MaxDate         =   401768
                  MinDate         =   -328351
               End
            End
            Begin VB.TextBox txtD_COMPTECOM 
               Height          =   330
               Left            =   1905
               Locked          =   -1  'True
               TabIndex        =   23
               Text            =   "COMPTECOM"
               Top             =   500
               Width           =   2310
            End
            Begin VB.TextBox txtD_COMPTEINT 
               Height          =   330
               Left            =   1905
               Locked          =   -1  'True
               TabIndex        =   22
               Text            =   "COMPTEINT"
               Top             =   1010
               Width           =   4050
            End
            Begin VB.TextBox txtD_COMPTEOBL 
               Height          =   330
               Left            =   4425
               Locked          =   -1  'True
               TabIndex        =   21
               Text            =   "COMPTEOBL"
               Top             =   500
               Width           =   960
            End
            Begin VB.TextBox txtD_COMPTEFON 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1905
               Locked          =   -1  'True
               TabIndex        =   20
               Text            =   "COMPTEFON"
               Top             =   1500
               Width           =   405
            End
            Begin VB.TextBox txtD_PLANCOPRO 
               Height          =   330
               Left            =   5505
               Locked          =   -1  'True
               TabIndex        =   19
               Text            =   "PLANCOPRO"
               Top             =   500
               Width           =   495
            End
            Begin VB.TextBox txtD_COMPTEOUV 
               Height          =   330
               Left            =   2955
               Locked          =   -1  'True
               TabIndex        =   18
               Text            =   "COMPTEOUV"
               Top             =   1485
               Width           =   1140
            End
            Begin VB.TextBox txtD_COMPTECLO 
               Height          =   330
               Left            =   4545
               Locked          =   -1  'True
               TabIndex        =   17
               Text            =   "COMPTECLO"
               Top             =   1500
               Width           =   1215
            End
            Begin VB.TextBox txtD_CLIENACLI 
               Height          =   330
               Left            =   1905
               Locked          =   -1  'True
               TabIndex        =   16
               Text            =   "CLIENACLI"
               Top             =   2500
               Width           =   1305
            End
            Begin VB.TextBox txtD_CLIENASIG 
               Height          =   330
               Left            =   3375
               Locked          =   -1  'True
               TabIndex        =   15
               Text            =   "CLIENASIG"
               Top             =   2475
               Width           =   1230
            End
            Begin VB.TextBox txtD_CLIENARA1 
               Height          =   330
               Left            =   1905
               Locked          =   -1  'True
               TabIndex        =   14
               Text            =   "COMPTEINT"
               Top             =   3000
               Width           =   4050
            End
            Begin VB.TextBox txtD_CLIENANAT 
               Height          =   330
               Left            =   1905
               Locked          =   -1  'True
               TabIndex        =   13
               Text            =   "CLIENARES"
               Top             =   3500
               Width           =   690
            End
            Begin VB.TextBox txtD_CLIENARES 
               Height          =   330
               Left            =   4695
               Locked          =   -1  'True
               TabIndex        =   12
               Text            =   "CLIENARES"
               Top             =   2460
               Width           =   1170
            End
            Begin VB.TextBox txtD_CLIENARSD 
               Height          =   330
               Left            =   2835
               Locked          =   -1  'True
               TabIndex        =   11
               Text            =   "CLIENARSD"
               Top             =   3500
               Width           =   675
            End
            Begin VB.Label lblD_COMPTECOM 
               BackColor       =   &H00D0F0FF&
               Caption         =   "compte PCI produit"
               Height          =   345
               Left            =   180
               TabIndex        =   29
               Top             =   550
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
            Begin VB.Label lblD_COMPTEFON 
               BackColor       =   &H00D0F0FF&
               Caption         =   "code fonct,Dcre,Dclo"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   180
               TabIndex        =   27
               Top             =   1550
               Width           =   1530
            End
            Begin VB.Label lblD_CLIENACLI 
               BackColor       =   &H00D0F0FF&
               Caption         =   "client, sigle, resp"
               Height          =   345
               Left            =   180
               TabIndex        =   26
               Top             =   2550
               Width           =   1530
            End
            Begin VB.Label lblD_CLIENARA1 
               BackColor       =   &H00D0F0FF&
               Caption         =   "intitulé"
               Height          =   345
               Left            =   180
               TabIndex        =   25
               Top             =   3060
               Width           =   1530
            End
            Begin VB.Label lblD_CLIENANAT 
               BackColor       =   &H00D0F0FF&
               Caption         =   "pays nationalité, rés"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   180
               TabIndex        =   24
               Top             =   3550
               Width           =   1530
            End
         End
         Begin VB.TextBox txtFg 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3015
            Left            =   -68115
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   7
            Text            =   "SAB_Dossier_DB.frx":05B8
            Top             =   1260
            Visible         =   0   'False
            Width           =   5595
         End
         Begin RichTextLib.RichTextBox txtRTF 
            Height          =   8385
            Left            =   -74880
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   420
            Visible         =   0   'False
            Width           =   12795
            _ExtentX        =   22569
            _ExtentY        =   14790
            _Version        =   393217
            BackColor       =   15790320
            HideSelection   =   0   'False
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"SAB_Dossier_DB.frx":05C0
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
         Begin MSFlexGridLib.MSFlexGrid fgBIAMVT 
            Height          =   8220
            Left            =   -74880
            TabIndex        =   8
            ToolTipText     =   "Cliquer pour obtenir le détail d'une pièce comptable"
            Top             =   525
            Width           =   12795
            _ExtentX        =   22569
            _ExtentY        =   14499
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
            FormatString    =   $"SAB_Dossier_DB.frx":0640
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
         Begin MSFlexGridLib.MSFlexGrid fgExtrait 
            Height          =   8370
            Left            =   120
            TabIndex        =   34
            ToolTipText     =   "cliquer ici pour afficher la pièce comptable (puis le dossier)"
            Top             =   435
            Width           =   12795
            _ExtentX        =   22569
            _ExtentY        =   14764
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16777215
            ForeColor       =   16384
            BackColorFixed  =   12632064
            ForeColorFixed  =   16777215
            BackColorBkg    =   -2147483643
            AllowUserResizing=   3
            FormatString    =   $"SAB_Dossier_DB.frx":0736
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
      Begin VB.Label libSelect1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4545
         TabIndex        =   36
         Top             =   210
         Width           =   8610
      End
      Begin VB.Label libSelect0 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   195
         TabIndex        =   4
         Top             =   195
         Width           =   4260
      End
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "cliquer ici pour fermer cette fenêtre (ESC = fermeture progressive)"
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
      Picture         =   "SAB_Dossier_DB.frx":0828
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "cliquer ici pour exporter les informations dans un fichier Excel"
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
End
Attribute VB_Name = "frmSAB_Dossier_DB"
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
Dim YGOSDOS0_Aut As typeAuthorization
Dim blnAuto As Boolean, blnError As Boolean
Dim cmdSelect_SQL_K As String
Dim wAmjMin As String, wAmjMax As String

Dim HeightOfLine As Long, LinesOfText As Long

Dim txtRTF_prtForeColor_Header As Long
Dim Mesg_aid As Long, mesg_s_umidl As Long, mesg_s_umidh As Long

Dim fgBIAMVT_FormatString As String, fgBIAMVT_K As Integer
Dim fgBIAMVT_RowDisplay As Integer, fgBIAMVT_RowClick As Integer, fgBIAMVT_ColClick As Integer
Dim fgBIAMVT_ColorClick As Long, fgBIAMVT_ColorDisplay As Long
Dim fgBIAMVT_Sort1 As Integer, fgBIAMVT_Sort2 As Integer
Dim fgBIAMVT_SortAD As Integer, fgBIAMVT_Sort1_Old As Integer
Dim fgBIAMVT_arrIndex As Integer
Dim blnfgBIAMVT_DisplayLine As Boolean

Dim memoYBIAMVTH As typeYBIAMVT0
Dim xYBIAMVTH As typeYBIAMVT0, newYBIAMVTH As typeYBIAMVT0, oldYBIAMVTH As typeYBIAMVT0
Dim mYBIAMVTH_Rupture As typeYBIAMVT0

Dim fgCPTPIE_FormatString As String, fgCPTPIE_K As Integer
Dim fgCPTPIE_RowDisplay As Integer, fgCPTPIE_RowClick As Integer, fgCPTPIE_ColClick As Integer
Dim fgCPTPIE_ColorClick As Long, fgCPTPIE_ColorDisplay As Long
Dim fgCPTPIE_Sort1 As Integer, fgCPTPIE_Sort2 As Integer
Dim fgCPTPIE_SortAD As Integer, fgCPTPIE_Sort1_Old As Integer
Dim fgCPTPIE_arrIndex As Integer
Dim blnfgCPTPIE_DisplayLine As Boolean


Dim fgCPTPIE_2_FormatString As String, fgCPTPIE_2_K As Integer
Dim fgCPTPIE_2_RowDisplay As Integer, fgCPTPIE_2_RowClick As Integer, fgCPTPIE_2_ColClick As Integer
Dim fgCPTPIE_2_ColorClick As Long, fgCPTPIE_2_ColorDisplay As Long
Dim fgCPTPIE_2_Sort1 As Integer, fgCPTPIE_2_Sort2 As Integer
Dim fgCPTPIE_2_SortAD As Integer, fgCPTPIE_2_Sort1_Old As Integer
Dim fgCPTPIE_2_arrIndex As Integer
Dim blnfgCPTPIE_2_DisplayLine As Boolean

Dim fgExtrait_FormatString As String, fgExtrait_K As Integer
Dim fgExtrait_RowDisplay As Integer, fgExtrait_RowClick As Integer, fgExtrait_ColClick As Integer
Dim fgExtrait_ColorClick As Long, fgExtrait_ColorDisplay As Long
Dim fgExtrait_Sort1 As Integer, fgExtrait_Sort2 As Integer
Dim fgExtrait_SortAD As Integer, fgExtrait_Sort1_Old As Integer
Dim fgExtrait_arrIndex As Integer
Dim blnfgExtrait_DisplayLine As Boolean


Dim xYCPTPIEH As typeYBIAMVT0, newYCPTPIEH As typeYBIAMVT0, oldYCPTPIEH As typeYBIAMVT0
Dim oldYBIACPT0 As typeYBIACPT0

Dim mSQL_Dossier_YBIAMVTHN As String, mSQL_Dossier_YDOSXODN As String
Dim mSQL_Dossier_Pièce As String
Dim mSQL_Extrait_YBIAMVTHN As String
Dim mSQL_Extrait_Pièce As String

Dim mSQL_Extrait_Fct As String, mSQL_Extrait_AMJMin As String, mSQL_Extrait_AMJMax As String
Public Sub fgBIAMVT_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
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
End Sub

Public Sub fgBIAMVT_Display(lMOUVEMSER As String, lMOUVEMSSE As String, lMOUVEMOPE As String, lMOUVEMNUM As Long)
Dim wColor As Long
Dim xSql As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String, X As String


On Error GoTo Error_Handler
SSTab1.Tab = 0
'fgswift_Reset
'libSelect0 = "Dossier : " & lMOUVEMSER & " " & lMOUVEMSSE & " " & lMOUVEMOPE & " " & lMOUVEMNUM
libSelect0 = "Dossier : " & " " & lMOUVEMOPE & " " & lMOUVEMNUM

fgCPTPIE.Visible = False
fraCompte.Visible = False
fraSelect.Visible = False
fgBIAMVT_Reset

fgBIAMVT.Rows = 1
fgBIAMVT.FormatString = fgBIAMVT_FormatString
fgBIAMVT.Row = 0

currentAction = "fgBIAMVT_Display"

If lMOUVEMOPE = "RDE" Or lMOUVEMOPE = "RDI" Then
    X = " where MOUVEMNUM between " & lMOUVEMNUM * 100 & " and " & lMOUVEMNUM * 100 + 99
Else
    X = " where MOUVEMNUM = " & lMOUVEMNUM
End If

'________________________________________________________________________________
'xWhere = X _
'     & " and MOUVEMOPE = '" & lMOUVEMOPE & "'" _
'     & " and MOUVEMSER = '" & lMOUVEMSER & "'" _
'     & " and MOUVEMSSE = '" & lMOUVEMSSE & "'" _
'     & " order by MOUVEMDTR, MOUVEMPIE,MOUVEMECR "
xWhere = X _
     & " and MOUVEMOPE = '" & lMOUVEMOPE & "'" _
     & " order by MOUVEMDTR, MOUVEMPIE,MOUVEMECR "


mSQL_Dossier_YBIAMVTHN = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTHN  " & xWhere
Set rsSab = cnsab.Execute(mSQL_Dossier_YBIAMVTHN)

Do While Not rsSab.EOF
    V = rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVTH)
         
    fgBIAMVT.Rows = fgBIAMVT.Rows + 1
    fgBIAMVT.Row = fgBIAMVT.Rows - 1
    If fctUser_Classe_Aut(xYBIAMVTH.COMPTECLA) Then fgBIAMVT_DisplayLine ""
    
    rsSab.MoveNext
Loop
'________________________________________________________________________________
xWhere = " where DOSXODNUM = " & lMOUVEMNUM _
     & " and DOSXODOPE = '" & lMOUVEMOPE & "'" _
     & " and MOUVEMETA = 1 and MOUVEMPIE = DOSXODPIE and MOUVEMECR = DOSXODECR " _
     & " and MOUVEMSER = '" & lMOUVEMSER & "'" _
     & " and MOUVEMSSE = '" & lMOUVEMSSE & "'" _
     & " order by MOUVEMDTR, MOUVEMPIE,MOUVEMECR "

' and MOUVEMDTR = DOSXODDTR

mSQL_Dossier_YDOSXODN = "select * from " & paramIBM_Library_SABSPE & ".YDOSXOD0N , " & paramIBM_Library_SABSPE & ".YBIAMVTHP  " & xWhere
Set rsSab = cnsab.Execute(mSQL_Dossier_YDOSXODN)

Do While Not rsSab.EOF
    V = rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVTH)
    fgBIAMVT.Rows = fgBIAMVT.Rows + 1
    fgBIAMVT.Row = fgBIAMVT.Rows - 1
    If fctUser_Classe_Aut(xYBIAMVTH.COMPTECLA) Then fgBIAMVT_DisplayLine "XOD"
    
    rsSab.MoveNext
Loop


If fgBIAMVT.Rows > 1 Then fgBIAMVT_Sort1 = fgBIAMVT_arrIndex: fgBIAMVT_Sort2 = fgBIAMVT_arrIndex: fgBIAMVT_Sort
fgBIAMVT.Visible = True
fraSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub fgBIAMVT_DisplayLine(lFct As String)
'Dim K As Integer
'Dim wColor As Long, wColor_Row As Long
Dim X As String
On Error Resume Next
fgBIAMVT.Col = 0: fgBIAMVT.Text = "  " & dateImp10_S(xYBIAMVTH.MOUVEMDTR + 19000000)
fgBIAMVT.Col = 1:
If lFct = "XOD" Then
    fgBIAMVT.Text = xYBIAMVTH.MOUVEMSER & " " & xYBIAMVTH.MOUVEMSSE & " " & rsSab("DOSXODOPE") & " " & rsSab("DOSXODNUM") & " " & xYBIAMVTH.MOUVEMEVE _
                   & " (" & xYBIAMVTH.MOUVEMOPE & " " & xYBIAMVTH.MOUVEMNUM & ")"
    fgBIAMVT.CellForeColor = vbMagenta
Else
    fgBIAMVT.Text = xYBIAMVTH.MOUVEMSER & " " & xYBIAMVTH.MOUVEMSSE & " " & xYBIAMVTH.MOUVEMOPE & " " & xYBIAMVTH.MOUVEMNUM & " " & xYBIAMVTH.MOUVEMEVE
End If

fgBIAMVT.Col = 2: fgBIAMVT.Text = xYBIAMVTH.MOUVEMCOM

fgBIAMVT.Col = IIf(xYBIAMVTH.MOUVEMMON > 0, 3, 4)

fgBIAMVT.Text = Format$(Abs(xYBIAMVTH.MOUVEMMON), "### ### ### ##0.00")

If xYBIAMVTH.MOUVEMMON > 0 Then
    fgBIAMVT.CellForeColor = vbRed
Else
    fgBIAMVT.CellForeColor = vbBlue
End If

fgBIAMVT.Col = 5: fgBIAMVT.Text = Trim(xYBIAMVTH.LIBELLIB1) & Trim(xYBIAMVTH.LIBELLIB2) & Trim(xYBIAMVTH.LIBELLIB3) & Trim(xYBIAMVTH.LIBELLIB4)
fgBIAMVT.Col = 6:
X = Format$(xYBIAMVTH.MOUVEMPIE, "##### ##0") & "-" & Format$(xYBIAMVTH.MOUVEMECR, "### ##0")
fgBIAMVT.Text = X
fgBIAMVT.Col = fgBIAMVT_arrIndex
    fgBIAMVT.Text = xYBIAMVTH.MOUVEMDTR & X
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


Public Sub fgExtrait_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgExtrait.Row

If lRow > 0 And lRow < fgExtrait.Rows Then
    fgExtrait.Row = lRow
    For I = fgExtrait_arrIndex To fgExtrait.FixedCols Step -1
        fgExtrait.Col = I: fgExtrait.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgExtrait.Row = mRow
    If fgExtrait.Row > 0 Then
        lRow = fgExtrait.Row
        fgExtrait.Col = fgExtrait_arrIndex
        lColor_Old = fgExtrait.CellBackColor
        For I = fgExtrait_arrIndex To fgExtrait.FixedCols Step -1
          fgExtrait.Col = I: fgExtrait.CellBackColor = lColor
        Next I
    End If
End If
fgExtrait.LeftCol = fgExtrait.FixedCols
End Sub

Public Sub fgExtrait_Display(lMOUVEMCOM As String, lAMJMin As String, lAMJMax As String)
Dim K As Integer, curBIAMVTSD0 As Currency
Dim xSql As String, xWhere As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

fgExtrait_Reset
libSelect1 = "Extrait de compte : " & lMOUVEMCOM & " du " & dateImp10_S(lAMJMin) & " au " & dateImp10_S(lAMJMax) & " - " & oldYBIACPT0.COMPTEINT

fgExtrait.Rows = 1
fgExtrait.FormatString = fgExtrait_FormatString
fgExtrait.Row = 0
For K = 0 To 6: fgExtrait.Col = K: fgExtrait.CellFontBold = True: Next K
currentAction = "fgExtrait_Display"

'============================================================
If Not fctUser_Classe_Aut(oldYBIACPT0.COMPTECLA) Then Exit Sub
'============================================================


xWhere = " where MOUVEMCOM = '" & lMOUVEMCOM & "'" _
     & " and MOUVEMDTR >= " & lAMJMin - 19000000 _
     & " and MOUVEMDTR <= " & lAMJMax - 19000000 _
     & " order by MOUVEMDTR, MOUVEMOPE, MOUVEMNUM, MOUVEMPIE,MOUVEMECR "


mSQL_Extrait_YBIAMVTHN = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH  left outer join " & paramIBM_Library_SABSPE & ".YDOSXOD0 on DOSXODPIE = MOUVEMPIE and DOSXODECR = MOUVEMECR " & xWhere



Set rsSab = cnsab.Execute(mSQL_Extrait_YBIAMVTHN)

Do While Not rsSab.EOF
    V = rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVTH)
    If fgExtrait.Rows = 1 Then
        curBIAMVTSD0 = xYBIAMVTH.BIAMVTSD0
        fgExtrait_DisplayLine_MTD xYBIAMVTH.BIAMVTSD0
        mYBIAMVTH_Rupture.MOUVEMDTR = xYBIAMVTH.MOUVEMDTR
    Else
        If mYBIAMVTH_Rupture.MOUVEMDTR <> xYBIAMVTH.MOUVEMDTR Then
            fgExtrait.Col = 6
            fgExtrait.Text = Format$(-curBIAMVTSD0, "### ### ### ##0.00")
            
            If curBIAMVTSD0 > 0 Then
                fgExtrait.CellForeColor = vbRed
                fgExtrait.CellBackColor = mColor_W0
                fgExtrait.Col = 2: fgExtrait.CellBackColor = mColor_W0
            Else
                fgExtrait.CellForeColor = vbBlue
                fgExtrait.CellBackColor = mColor_B0
                fgExtrait.Col = 0: fgExtrait.CellBackColor = mColor_B0
            End If
            mYBIAMVTH_Rupture.MOUVEMDTR = xYBIAMVTH.MOUVEMDTR
        End If
    End If
    curBIAMVTSD0 = curBIAMVTSD0 + xYBIAMVTH.MOUVEMMON
    fgExtrait.Rows = fgExtrait.Rows + 1
    fgExtrait.Row = fgExtrait.Rows - 1
    If fctUser_Classe_Aut(xYBIAMVTH.COMPTECLA) Then fgExtrait_DisplayLine I
    
    rsSab.MoveNext
Loop

fgExtrait.Rows = fgExtrait.Rows + 1
fgExtrait.Row = fgExtrait.Rows - 1
fgExtrait.Col = 1: fgExtrait.Text = "SOLDE FINAL"
fgExtrait_DisplayLine_MTD curBIAMVTSD0
For K = 0 To 6: fgExtrait.Col = K
    fgExtrait.CellFontBold = True
    fgExtrait.CellBackColor = fgExtrait.BackColorFixed
Next K

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub fgExtrait_Display_MOUVEMDVA(lMOUVEMCOM As String, lAMJMin As String, lAMJMax As String)
Dim K As Integer, curBIAMVTSD0 As Currency
Dim xSql As String, xWhere As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String, wAmjMin7 As Long, wAmjMax7 As Long
On Error GoTo Error_Handler

wAmjMin7 = lAMJMin - 19000000
wAmjMax7 = lAMJMax - 19000000

fgExtrait_Reset
libSelect1 = "Extrait en DATE DE VALEUR du compteCompte : " & lMOUVEMCOM & " du " & dateImp10_S(lAMJMin) & " au " & dateImp10_S(lAMJMax) & " - " & oldYBIACPT0.COMPTEINT

fgExtrait.Rows = 1
fgExtrait.FormatString = fgExtrait_FormatString
fgExtrait.Row = 0
For K = 0 To 6: fgExtrait.Col = K: fgExtrait.CellFontBold = True: Next K
currentAction = "fgExtrait_Display_MOUVEMDVA"

'============================================================
If Not fctUser_Classe_Aut(oldYBIACPT0.COMPTECLA) Then Exit Sub
'============================================================

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
     & " where MOUVEMCOM = '" & lMOUVEMCOM & "'" _
     & " order by MOUVEMDTR, MOUVEMPIE, MOUVEMECR "

Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then
    Call rsYBIAMVT0_GetBuffer(rsSab, mYBIAMVTH_Rupture)
    curBIAMVTSD0 = rsSab("BIAMVTSD0")
Else
    V = "Néant"
    GoTo Error_MsgBox
End If

xWhere = " where MOUVEMCOM = '" & lMOUVEMCOM & "'" _
     & " order by MOUVEMDVA, MOUVEMDTR, MOUVEMOPE, MOUVEMNUM, MOUVEMPIE,MOUVEMECR "


mSQL_Extrait_YBIAMVTHN = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH  left outer join " & paramIBM_Library_SABSPE & ".YDOSXOD0 on DOSXODPIE = MOUVEMPIE and DOSXODECR = MOUVEMECR " & xWhere



Set rsSab = cnsab.Execute(mSQL_Extrait_YBIAMVTHN)

Do While Not rsSab.EOF
    If rsSab("MOUVEMDVA") < wAmjMin7 Then
        curBIAMVTSD0 = curBIAMVTSD0 + rsSab("MOUVEMMON")
        mYBIAMVTH_Rupture.MOUVEMDVA = rsSab("MOUVEMDVA")
    Else
        If rsSab("MOUVEMDVA") > wAmjMax7 Then Exit Do
    
        V = rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVTH)
        If fgExtrait.Rows = 1 Then
            fgExtrait_DisplayLine_MTD curBIAMVTSD0
            mYBIAMVTH_Rupture.MOUVEMDVA = xYBIAMVTH.MOUVEMDVA
        Else
            If mYBIAMVTH_Rupture.MOUVEMDVA <> xYBIAMVTH.MOUVEMDVA Then
                fgExtrait.Col = 6
                fgExtrait.Text = Format$(-curBIAMVTSD0, "### ### ### ##0.00")
                
                If curBIAMVTSD0 > 0 Then
                    fgExtrait.CellForeColor = vbRed
                    fgExtrait.CellBackColor = mColor_W0
                    fgExtrait.Col = 2: fgExtrait.CellBackColor = mColor_W0
                Else
                    fgExtrait.CellForeColor = vbBlue
                    fgExtrait.CellBackColor = mColor_B0
                    fgExtrait.Col = 2: fgExtrait.CellBackColor = mColor_B0
                End If
                mYBIAMVTH_Rupture.MOUVEMDVA = xYBIAMVTH.MOUVEMDVA
            End If

        End If
        curBIAMVTSD0 = curBIAMVTSD0 + xYBIAMVTH.MOUVEMMON
        fgExtrait.Rows = fgExtrait.Rows + 1
        fgExtrait.Row = fgExtrait.Rows - 1
        'If fctUser_Classe_Aut(xYBIAMVTH.COMPTECLA) Then
        
        fgExtrait_DisplayLine I
        
        If rsSab("MOUVEMDTR") > wAmjMax7 Then
            fgExtrait.Col = 0: fgExtrait.CellBackColor = mColor_Y2
        Else
                'dateMOUVEMDTR = Date_VB(CLng(rsSab("MOUVEMDTR") + 19000000), 0)
                'dateMOUVEMDVA = Date_VB(CLng(rsSab("MOUVEMDVA") + 19000000), 0)
                
                If Abs(DateDiff("d", Date_VB(CLng(rsSab("MOUVEMDVA") + 19000000), 0), Date_VB(CLng(rsSab("MOUVEMDTR") + 19000000), 0))) > 7 Then fgExtrait.Col = 0: fgExtrait.CellBackColor = mColor_Y2
        End If
    End If
    rsSab.MoveNext
Loop

fgExtrait.Rows = fgExtrait.Rows + 1
fgExtrait.Row = fgExtrait.Rows - 1
fgExtrait.Col = 1: fgExtrait.Text = "SOLDE en VALEUR"
fgExtrait_DisplayLine_MTD curBIAMVTSD0
For K = 0 To 6: fgExtrait.Col = K
    fgExtrait.CellFontBold = True
    fgExtrait.CellBackColor = fgExtrait.BackColorFixed
Next K

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgExtrait_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim blnSolde As Boolean

On Error Resume Next
fgExtrait.Col = 0: fgExtrait.Text = "   " & dateImp10_S(xYBIAMVTH.MOUVEMDTR + 19000000)
fgExtrait.Col = 1
If Not IsNull(rsSab("DOSXODOPE")) Then
    fgExtrait.Text = xYBIAMVTH.MOUVEMSER & " " & xYBIAMVTH.MOUVEMSSE & " " & rsSab("DOSXODOPE") & " " & rsSab("DOSXODNUM") & " " & xYBIAMVTH.MOUVEMEVE _
                   & " (" & xYBIAMVTH.MOUVEMOPE & " " & xYBIAMVTH.MOUVEMNUM & ")"
    fgExtrait.CellForeColor = vbMagenta
Else
    fgExtrait.Text = xYBIAMVTH.MOUVEMSER & " " & xYBIAMVTH.MOUVEMSSE & " " & xYBIAMVTH.MOUVEMOPE & " " & xYBIAMVTH.MOUVEMNUM & " " & xYBIAMVTH.MOUVEMEVE
End If

fgExtrait.Col = 2: fgExtrait.Text = "   " & dateImp10_S(xYBIAMVTH.MOUVEMDVA + 19000000)
fgExtrait_DisplayLine_MTD xYBIAMVTH.MOUVEMMON

fgExtrait.Col = 5: fgExtrait.Text = Trim(xYBIAMVTH.LIBELLIB1) & Trim(xYBIAMVTH.LIBELLIB2) & Trim(xYBIAMVTH.LIBELLIB3) & Trim(xYBIAMVTH.LIBELLIB4)
fgExtrait.Col = 7: fgExtrait.Text = Format$(xYBIAMVTH.MOUVEMPIE, "##### ##0") & "-" & Format$(xYBIAMVTH.MOUVEMECR, "### ##0")
'fgExtrait.Col = fgExtrait_arrIndex: fgExtrait.Text = lIndex
End Sub


Public Sub fgExtrait_DisplayLine_MTD(lMOUVEMMON As Currency)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim blnSolde As Boolean

On Error Resume Next
fgExtrait.Col = IIf(lMOUVEMMON > 0, 3, 4)

fgExtrait.Text = Format$(Abs(lMOUVEMMON), "### ### ### ##0.00")

If lMOUVEMMON > 0 Then
    fgExtrait.CellForeColor = vbRed
Else
    fgExtrait.CellForeColor = vbBlue
End If

End Sub


Public Sub fgExtrait_Reset()
fgExtrait.Clear
fgExtrait_Sort1 = 0: fgExtrait_Sort2 = 0
fgExtrait_Sort1_Old = -1
fgExtrait_RowDisplay = 0: fgExtrait_RowClick = 0
fgExtrait_arrIndex = fgExtrait.Cols - 1
blnfgExtrait_DisplayLine = False
fgExtrait_SortAD = 6
fgExtrait.LeftCol = fgExtrait.FixedCols

End Sub

Public Sub fgExtrait_Sort()
If fgExtrait.Rows > 1 Then
    fgExtrait.Row = 1
    fgExtrait.RowSel = fgExtrait.Rows - 1
    
    If fgExtrait_Sort1_Old = fgExtrait_Sort1 Then
        If fgExtrait_SortAD = 5 Then
            fgExtrait_SortAD = 6
        Else
            fgExtrait_SortAD = 5
        End If
    Else
        fgExtrait_SortAD = 5
    End If
    fgExtrait_Sort1_Old = fgExtrait_Sort1
    
    fgExtrait.Col = fgExtrait_Sort1
    fgExtrait.ColSel = fgExtrait_Sort2
    fgExtrait.Sort = fgExtrait_SortAD
End If

End Sub



Public Sub fgCPTPIE_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgCPTPIE.Visible = False
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
fgCPTPIE.Visible = True
End Sub


Private Sub fgCPTPIE_Display()
Dim xSql As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
fgCPTPIE.Visible = False
fraCompte.Visible = False
fgCPTPIE_Reset

fgCPTPIE.Rows = 1
fgCPTPIE.FormatString = fgCPTPIE_FormatString
fgCPTPIE.Row = 0

currentAction = "fgCPTPIE_Display"

xWhere = " where MOUVEMETA = 1 " _
     & " and MOUVEMPIE = " & oldYBIAMVTH.MOUVEMPIE _
     & " and MOUVEMCOM = COMPTECOM " _
     & " order by MOUVEMECR "


mSQL_Dossier_Pièce = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH , " & paramIBM_Library_SABSPE & ".YBIACPT0 " & xWhere
Set rsSab = cnsab.Execute(mSQL_Dossier_Pièce)

Do While Not rsSab.EOF
    V = rsYBIAMVT0_GetBuffer(rsSab, xYCPTPIEH)
         
    fgCPTPIE.Rows = fgCPTPIE.Rows + 1
    fgCPTPIE.Row = fgCPTPIE.Rows - 1
    If fctUser_Classe_Aut(xYCPTPIEH.COMPTECLA) Then fgCPTPIE_DisplayLine I
    
    rsSab.MoveNext
Loop

fgCPTPIE.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub




Public Sub fgCPTPIE_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wFontSize As Integer
Dim blnFontBold As Boolean

On Error Resume Next
fgCPTPIE.Col = 0: fgCPTPIE.Text = dateIBM10(xYCPTPIEH.MOUVEMDTR, True)
If oldYBIAMVTH.MOUVEMECR = xYCPTPIEH.MOUVEMECR Then
    blnFontBold = True
    wFontSize = 8 '6
Else
    blnFontBold = False
    wFontSize = 8
End If
fgCPTPIE.Col = 1: fgCPTPIE.Text = xYCPTPIEH.MOUVEMSER & " " & xYCPTPIEH.MOUVEMSSE & " " & xYCPTPIEH.MOUVEMOPE & " " & xYCPTPIEH.MOUVEMNUM & " " & xYCPTPIEH.MOUVEMEVE

fgCPTPIE.Col = 2: fgCPTPIE.Text = Trim(xYCPTPIEH.MOUVEMCOM)
fgCPTPIE.CellFontBold = blnFontBold
If xYCPTPIEH.MOUVEMMON > 0 Then
    fgCPTPIE.Col = 3: fgCPTPIE.Text = Format$(xYCPTPIEH.MOUVEMMON, "### ### ### ##0.00")
    fgCPTPIE.CellForeColor = vbRed
Else
    fgCPTPIE.Col = 4: fgCPTPIE.Text = Format$(Abs(xYCPTPIEH.MOUVEMMON), "### ### ### ##0.00")
    fgCPTPIE.CellForeColor = vbBlue
End If

fgCPTPIE.Col = 5: fgCPTPIE.Text = Trim(rsSab("COMPTEINT"))
fgCPTPIE.CellFontBold = blnFontBold
fgCPTPIE.CellFontSize = wFontSize
fgCPTPIE.Col = 6: fgCPTPIE.Text = Trim(xYCPTPIEH.LIBELLIB1) & Trim(xYCPTPIEH.LIBELLIB2) & Trim(xYCPTPIEH.LIBELLIB3) & Trim(xYCPTPIEH.LIBELLIB4)
fgCPTPIE.Col = 7: fgCPTPIE.Text = Format$(xYCPTPIEH.MOUVEMPIE, "##### ##0") & "-" & Format$(xYCPTPIEH.MOUVEMECR, "### ##0")
fgCPTPIE.Col = 8: fgCPTPIE.Text = xYCPTPIEH.MOUVEMANU

fgCPTPIE.Col = fgCPTPIE_arrIndex: fgCPTPIE.Text = lIndex
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



Public Sub fgCPTPIE_2_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgCPTPIE_2.Visible = False
mRow = fgCPTPIE_2.Row

If lRow > 0 And lRow < fgCPTPIE_2.Rows Then
    fgCPTPIE_2.Row = lRow
    For I = fgCPTPIE_2_arrIndex To fgCPTPIE_2.FixedCols Step -1
        fgCPTPIE_2.Col = I: fgCPTPIE_2.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgCPTPIE_2.Row = mRow
    If fgCPTPIE_2.Row > 0 Then
        lRow = fgCPTPIE_2.Row
        fgCPTPIE_2.Col = fgCPTPIE_2_arrIndex
        lColor_Old = fgCPTPIE_2.CellBackColor
        For I = fgCPTPIE_2_arrIndex To fgCPTPIE_2.FixedCols Step -1
          fgCPTPIE_2.Col = I: fgCPTPIE_2.CellBackColor = lColor
        Next I
    End If
End If
fgCPTPIE_2.LeftCol = fgCPTPIE_2.FixedCols
fgCPTPIE_2.Visible = True
End Sub


Private Sub fgCPTPIE_2_Display()
Dim xSql As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
fgCPTPIE_2.Visible = False
fraCompte.Visible = False
fgCPTPIE_2_Reset

fgCPTPIE_2.Rows = 1
fgCPTPIE_2.FormatString = fgCPTPIE_2_FormatString
fgCPTPIE_2.Row = 0

currentAction = "fgCPTPIE_2_Display"

xWhere = " where MOUVEMETA = 1 " _
     & " and MOUVEMPIE = " & xYBIAMVTH.MOUVEMPIE _
     & " and MOUVEMCOM = COMPTECOM " _
     & " order by MOUVEMECR "


mSQL_Extrait_Pièce = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH , " & paramIBM_Library_SABSPE & ".YBIACPT0 " & xWhere
Set rsSab = cnsab.Execute(mSQL_Extrait_Pièce)

Do While Not rsSab.EOF
    V = rsYBIAMVT0_GetBuffer(rsSab, xYCPTPIEH)
         
    fgCPTPIE_2.Rows = fgCPTPIE_2.Rows + 1
    fgCPTPIE_2.Row = fgCPTPIE_2.Rows - 1
    If fctUser_Classe_Aut(xYCPTPIEH.COMPTECLA) Then fgCPTPIE_2_DisplayLine I
    
    rsSab.MoveNext
Loop

fgCPTPIE_2.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub




Public Sub fgCPTPIE_2_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wFontSize As Integer
Dim blnFontBold As Boolean

On Error Resume Next
fgCPTPIE_2.Col = 0: fgCPTPIE_2.Text = dateIBM10(xYCPTPIEH.MOUVEMDTR, True)
If oldYBIAMVTH.MOUVEMECR = xYCPTPIEH.MOUVEMECR Then
    blnFontBold = True
    wFontSize = 8 '6
Else
    blnFontBold = False
    wFontSize = 8
End If
fgCPTPIE_2.Col = 1: fgCPTPIE_2.Text = xYCPTPIEH.MOUVEMSER & " " & xYCPTPIEH.MOUVEMSSE & " " & xYCPTPIEH.MOUVEMOPE & " " & xYCPTPIEH.MOUVEMNUM & " " & xYCPTPIEH.MOUVEMEVE

fgCPTPIE_2.Col = 2: fgCPTPIE_2.Text = Trim(xYCPTPIEH.MOUVEMCOM)
fgCPTPIE_2.CellFontBold = blnFontBold
If xYCPTPIEH.MOUVEMMON > 0 Then
    fgCPTPIE_2.Col = 3: fgCPTPIE_2.Text = Format$(xYCPTPIEH.MOUVEMMON, "### ### ### ##0.00")
    fgCPTPIE_2.CellForeColor = vbRed
Else
    fgCPTPIE_2.Col = 4: fgCPTPIE_2.Text = Format$(Abs(xYCPTPIEH.MOUVEMMON), "### ### ### ##0.00")
    fgCPTPIE_2.CellForeColor = vbBlue
End If

fgCPTPIE_2.Col = 5: fgCPTPIE_2.Text = Trim(rsSab("COMPTEINT"))
fgCPTPIE_2.CellFontBold = blnFontBold
fgCPTPIE_2.CellFontSize = wFontSize
fgCPTPIE_2.Col = 6: fgCPTPIE_2.Text = Trim(xYCPTPIEH.LIBELLIB1) & Trim(xYCPTPIEH.LIBELLIB2) & Trim(xYCPTPIEH.LIBELLIB3) & Trim(xYCPTPIEH.LIBELLIB4)
fgCPTPIE_2.Col = 7: fgCPTPIE_2.Text = Format$(xYCPTPIEH.MOUVEMPIE, "##### ##0") & "-" & Format$(xYCPTPIEH.MOUVEMECR, "### ##0")
fgCPTPIE_2.Col = 8: fgCPTPIE_2.Text = xYCPTPIEH.MOUVEMANU

fgCPTPIE_2.Col = fgCPTPIE_2_arrIndex: fgCPTPIE_2.Text = lIndex
End Sub



Public Sub fgCPTPIE_2_Reset()
fgCPTPIE_2.Clear
fgCPTPIE_2_Sort1 = 0: fgCPTPIE_2_Sort2 = 0
fgCPTPIE_2_Sort1_Old = -1
fgCPTPIE_2_RowDisplay = 0: fgCPTPIE_2_RowClick = 0
fgCPTPIE_2_arrIndex = fgCPTPIE_2.Cols - 1
blnfgCPTPIE_2_DisplayLine = False
fgCPTPIE_2_SortAD = 6
fgCPTPIE_2.LeftCol = fgCPTPIE_2.FixedCols

End Sub


Public Sub fgCPTPIE_2_Sort()
If fgCPTPIE_2.Rows > 1 Then
    fgCPTPIE_2.Row = 1
    fgCPTPIE_2.RowSel = fgCPTPIE_2.Rows - 1
    
    If fgCPTPIE_2_Sort1_Old = fgCPTPIE_2_Sort1 Then
        If fgCPTPIE_2_SortAD = 5 Then
            fgCPTPIE_2_SortAD = 6
        Else
            fgCPTPIE_2_SortAD = 5
        End If
    Else
        fgCPTPIE_2_SortAD = 5
    End If
    fgCPTPIE_2_Sort1_Old = fgCPTPIE_2_Sort1
    
    fgCPTPIE_2.Col = fgCPTPIE_2_Sort1
    fgCPTPIE_2.ColSel = fgCPTPIE_2_Sort2
    fgCPTPIE_2.Sort = fgCPTPIE_2_SortAD
End If

End Sub









Private Sub cmdContext_Click()
Unload Me

End Sub

Private Sub cmdD_Extrait_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass
fgExtrait.Visible = False

Call DTPicker_Control(txtD_Extrait_AMJMIn, wAmjMin)
Call DTPicker_Control(txtD_Extrait_AMJMax, wAmjMax)


fgExtrait_FormatString = "< Date TRT         |<" & oldYBIACPT0.COMPTECOM & "            |< Date valeur    " _
                      & "|>           Débit  " & oldYBIACPT0.COMPTEDEV & "|>           Crédit   " & oldYBIACPT0.COMPTEDEV & "|<Libellé   " _
                      & "                                                                   |>Solde                  |"

If mSQL_Extrait_Fct = "MOUVEMDVA" Then
    Call fgExtrait_Display_MOUVEMDVA(oldYBIAMVTH.MOUVEMCOM, wAmjMin, wAmjMax)
Else
    Call fgExtrait_Display(oldYBIAMVTH.MOUVEMCOM, wAmjMin, wAmjMax)
End If
SSTab1.Tab = 1
fgExtrait.Visible = True
fraSelect.Visible = True

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdPrint_Click()
If Not txtRTF.Visible Then
    cmdPrint_Display
Else
    prtOrientation = vbPRORLandscape
    prtEdition_Open
    prtTitleText = libSelect0
    prtPgmName = "frmSAB_Dossier_DB"
    prtHeaderHeight = 300
    prtFormType = ""
    'prtForeColor_Header = txtRTF_prtForeColor_Header
    frmElpPrt.prtStdInit
    XPrt.CurrentY = prtMinY + prtHeaderHeight
    Call frmElpPrt.prtRTF(txtRTF.TextRTF)
    
    XPrt.DrawWidth = 5

    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)

    prtEdition_Close
End If



End Sub
Private Sub cmdPrint_Display()

'Dim mSQL_Dossier_YBIAMVTHN As String, mSQL_Dossier_YDOSXODN As String
'Dim mSQL_Dossier_Pièce As String
'Dim mSQL_Extrait_YBIAMVTHN As String
'Dim mSQL_Extrait_Pièce As String
Call lstErr_Clear(lstErr, cmdContext, "> Exportation ......"): DoEvents

Call YBIAMVTH_Exportation(mSQL_Extrait_Fct, mSQL_Extrait_AMJMin, mSQL_Extrait_AMJMax, lstErr, oldYBIAMVTH, mSQL_Dossier_YBIAMVTHN, mSQL_Dossier_YDOSXODN, mSQL_Dossier_Pièce, mSQL_Extrait_YBIAMVTHN, mSQL_Extrait_Pièce)

End Sub

Private Sub fgBIAMVT_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Dim K As Integer, XX As String
If y <= fgBIAMVT.RowHeightMin Then
    Select Case fgBIAMVT.Col
        Case 0: fgBIAMVT_Sort1 = 0: fgBIAMVT_Sort2 = 1: fgBIAMVT_Sort
        Case 1:  fgBIAMVT_Sort1 = 1: fgBIAMVT_Sort2 = 2: fgBIAMVT_Sort
        Case 2: fgBIAMVT_Sort1 = 2: fgBIAMVT_Sort2 = 2: fgBIAMVT_Sort
        Case 3: fgBIAMVT_Sort1 = 3: fgBIAMVT_Sort2 = 3: fgBIAMVT_Sort
        Case 4: fgBIAMVT_Sort1 = 4: fgBIAMVT_Sort2 = 4: fgBIAMVT_Sort
        Case 5: fgBIAMVT_Sort1 = 5: fgBIAMVT_Sort2 = 5: fgBIAMVT_Sort
    End Select
Else
    If fgBIAMVT.Rows > 1 Then
        Call fgBIAMVT_Color(fgBIAMVT_RowClick, MouseMoveUsr.BackColor, fgBIAMVT_ColorClick)
        'fgBIAMVT.Col = 1:  oldYBIAMVTH.MOUVEMCOM = Trim(fgBIAMVT.Text)
        fgBIAMVT.Col = 6:
        XX = Trim(fgBIAMVT.Text)
        K = InStr(1, XX, "-")
        oldYBIAMVTH.MOUVEMPIE = Val(Mid$(XX, 1, K - 1))
        oldYBIAMVTH.MOUVEMECR = Val(Mid$(XX, K + 1, Len(XX) - K))
        
        fgCPTPIE_Display
        
   End If
End If

End Sub


Private Sub fgCPTPIE_2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Dim XX As String, K1 As Integer
If y <= fgCPTPIE_2.RowHeightMin Then
Else
    If fgCPTPIE_2.Rows > 1 Then
        Call fgCPTPIE_2_Color(fgCPTPIE_2_RowClick, MouseMoveUsr.BackColor, fgCPTPIE_2_ColorClick)
        fgCPTPIE_2.Col = 1
        XX = Trim(fgCPTPIE_2.Text)
        
        oldYBIAMVTH.MOUVEMSER = Mid$(XX, 1, 2)
        oldYBIAMVTH.MOUVEMSSE = Mid$(XX, 4, 2)
        oldYBIAMVTH.MOUVEMOPE = Mid$(XX, 7, 3)
        K1 = InStr(11, XX, " "): oldYBIAMVTH.MOUVEMNUM = Val(Mid$(XX, 11, K1 - 11))

        Call fgBIAMVT_Display(oldYBIAMVTH.MOUVEMSER, oldYBIAMVTH.MOUVEMSSE, oldYBIAMVTH.MOUVEMOPE, oldYBIAMVTH.MOUVEMNUM)
   End If
End If


End Sub


Private Sub fgCPTPIE_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Dim XX As String
If y <= fgCPTPIE.RowHeightMin Then
Else
    If fgCPTPIE.Rows > 1 Then
        Call fgCPTPIE_Color(fgCPTPIE_RowClick, MouseMoveUsr.BackColor, fgCPTPIE_ColorClick)
        fgCPTPIE.Col = 2
        oldYBIAMVTH.MOUVEMCOM = Trim(fgCPTPIE.Text)
        fgCPTPIE.Col = 0: XX = Trim(fgCPTPIE.Text)
         Call dateJMA_AMJ(XX, wAmjMin)
        Call DTPicker_Set(txtD_Extrait_AMJMIn, wAmjMin)
        Call DTPicker_Set(txtD_Extrait_AMJMax, wAmjMin)

        fraCompte_display oldYBIAMVTH.MOUVEMCOM
   End If
End If

End Sub


Public Sub fraCompte_display(lCOMPTECOM As String)
Dim xSql As String

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '" & lCOMPTECOM & "'"
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then
    V = rsYBIACPT0_GetBuffer(rsSab, oldYBIACPT0)
    If IsNull(V) Then
        txtD_COMPTECOM = oldYBIACPT0.COMPTECOM
        txtD_COMPTEINT = oldYBIACPT0.COMPTEINT
        txtD_COMPTEOBL = oldYBIACPT0.COMPTEOBL
        txtD_COMPTEFON = oldYBIACPT0.COMPTEFON
        txtD_PLANCOPRO = oldYBIACPT0.PLANCOPRO
        If oldYBIACPT0.COMPTEOUV = 0 Then
            txtD_COMPTEOUV = ""
        Else
            txtD_COMPTEOUV = dateIBM10(oldYBIACPT0.COMPTEOUV, True)
        End If
        If oldYBIACPT0.COMPTECLO = 0 Then
            txtD_COMPTECLO = ""
        Else
            txtD_COMPTECLO = dateIBM10(oldYBIACPT0.COMPTECLO, True)
        End If

        txtD_CLIENACLI = oldYBIACPT0.CLIENACLI
        txtD_CLIENASIG = oldYBIACPT0.CLIENASIG
        txtD_CLIENARES = oldYBIACPT0.CLIENARES
        txtD_CLIENARA1 = oldYBIACPT0.CLIENARA1
        txtD_CLIENANAT = oldYBIACPT0.CLIENANAT
        txtD_CLIENARSD = oldYBIACPT0.CLIENARSD


        fraCompte.Visible = True
    End If
End If
End Sub


Private Sub fgExtrait_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Dim K As Integer, XX As String
If y <= fgExtrait.RowHeightMin Then
    If mSQL_Extrait_Fct = "" Then
        Select Case fgExtrait.Col
            Case 0: fgExtrait_Sort1 = 0: fgExtrait_Sort2 = 1: fgExtrait_Sort
            Case 1:  fgExtrait_Sort1 = 1: fgExtrait_Sort2 = 2: fgExtrait_Sort
            Case 2: fgExtrait_Sort1 = 2: fgExtrait_Sort2 = 2: fgExtrait_Sort
            Case 3: fgExtrait_Sort1 = 3: fgExtrait_Sort2 = 3: fgExtrait_Sort
            Case 4: fgExtrait_Sort1 = 4: fgExtrait_Sort2 = 4: fgExtrait_Sort
            Case 5: fgExtrait_Sort1 = 5: fgExtrait_Sort2 = 5: fgExtrait_Sort
        End Select
    End If
Else
    If fgExtrait.Rows > 1 Then
        Call fgExtrait_Color(fgExtrait_RowClick, MouseMoveUsr.BackColor, fgExtrait_ColorClick)
        'fgextrait.Col = 1:  oldYBIAMVTH.MOUVEMCOM = Trim(fgextrait.Text)
        fgExtrait.Col = 7:
        XX = Trim(fgExtrait.Text)
        K = InStr(1, XX, "-")
        xYBIAMVTH.MOUVEMPIE = Val(Mid$(XX, 1, K - 1))
        xYBIAMVTH.MOUVEMECR = Val(Mid$(XX, K + 1, Len(XX) - K))
        
        fgCPTPIE_2_Display
        
   End If
End If

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

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Return()
        SendKeys "{TAB}"
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
If fgCPTPIE_2.Visible Then
    fgCPTPIE_2.Visible = False
    Exit Sub
End If

If fgExtrait.Visible Then
    fgExtrait.Visible = False
    SSTab1.Tab = 0
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



'If fraSelect.Visible Then
'    fraSelect.Visible = False
'    Exit Sub
'End If

Unload Me

End Sub

Private Sub Form_Load()

frmSAB_Dossier_DB_Show

Set XForm = Me
Me.Left = 5000
KeyPreview = True

blnControl = False
'mWindowState = Me.WindowState
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate


fgBIAMVT_FormatString = fgBIAMVT.FormatString
fgBIAMVT.Enabled = True
fgBIAMVT.Visible = False

fgCPTPIE_FormatString = fgCPTPIE.FormatString
fgCPTPIE.Enabled = True
fgCPTPIE.Visible = False: txtFg.Visible = False


fgCPTPIE_2_FormatString = fgCPTPIE_2.FormatString
fgCPTPIE_2.Enabled = True
fgCPTPIE_2.Visible = False: txtFg.Visible = False

fgBIAMVT_FormatString = fgExtrait.FormatString
fgExtrait.Enabled = True
fgExtrait.Visible = False

libSelect0.BackColor = mColor_Y1
libSelect0.ForeColor = RGB(128, 64, 0)
libSelect1.BackColor = mColor_Y1
libSelect1.ForeColor = RGB(128, 64, 0)

Call DTPicker_Set(txtD_Extrait_AMJMIn, YBIATAB0_DATE_CPT_MP1)
Call DTPicker_Set(txtD_Extrait_AMJMax, YBIATAB0_DATE_CPT_J)
End Sub


Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub


Public Sub Form_Init(lFct As String, lMOUVEMCOM As String, lAMJMin As String, lAMJMax As String, lMOUVEMSER As String, lMOUVEMSSE As String, lMOUVEMOPE As String, lMOUVEMNUM As Long)
'___________________________________________________________
On Error Resume Next

If Not frmSAB_Dossier_DB.Visible Then frmSAB_Dossier_DB.Visible = True
frmSAB_Dossier_DB.Show

fraSelect.Visible = False: fgBIAMVT.Visible = False: fgCPTPIE.Visible = False: txtFg.Visible = False
fgCPTPIE_2.Visible = False
txtRTF.Visible = False

mSQL_Extrait_Fct = lFct
mSQL_Extrait_AMJMin = lAMJMin
mSQL_Extrait_AMJMax = lAMJMax

mSQL_Dossier_YBIAMVTHN = "": mSQL_Dossier_YDOSXODN = ""
mSQL_Dossier_Pièce = ""
mSQL_Extrait_YBIAMVTHN = ""
mSQL_Extrait_Pièce = ""

rsYBIAMVT0_Init oldYBIAMVTH
oldYBIAMVTH.MOUVEMNUM = lMOUVEMNUM
oldYBIAMVTH.MOUVEMOPE = lMOUVEMOPE
oldYBIAMVTH.MOUVEMSER = lMOUVEMSER
oldYBIAMVTH.MOUVEMSSE = lMOUVEMSSE
oldYBIAMVTH.MOUVEMCOM = lMOUVEMCOM

memoYBIAMVTH = oldYBIAMVTH

X = Trim(frmSAB_Dossier_DB.Caption)
AppActivate X

If lMOUVEMNUM > 0 Then
    Call fgBIAMVT_Display(lMOUVEMSER, lMOUVEMSSE, lMOUVEMOPE, lMOUVEMNUM)
Else
    Call DTPicker_Set(txtD_Extrait_AMJMIn, lAMJMin)
    Call DTPicker_Set(txtD_Extrait_AMJMax, lAMJMax)

    fraCompte_display oldYBIAMVTH.MOUVEMCOM

    cmdD_Extrait_Click
End If
End Sub

