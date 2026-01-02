VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBIA_GUIMAD 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_GUIMAD : mise à disposition"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "BIA_GUIMAD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   45
      Width           =   5055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "BIA_GUIMAD.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Statistiques"
      TabPicture(1)   =   "BIA_GUIMAD.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgStatistiques"
      Tab(1).Control(1)=   "lstW"
      Tab(1).ControlCount=   2
      Begin VB.ListBox lstW 
         Height          =   255
         Left            =   -67800
         Sorted          =   -1  'True
         TabIndex        =   45
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame fraSelect 
         Height          =   8445
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13560
         Begin VB.ListBox lstSelect 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5310
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   58
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Frame fraGUIMADLIEN 
            BackColor       =   &H00D0D0D0&
            Height          =   7095
            Left            =   3960
            TabIndex        =   53
            Top             =   1150
            Width           =   4455
            Begin VB.ListBox lstGUIMADLIEN_Display 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1740
               Left            =   360
               Sorted          =   -1  'True
               TabIndex        =   61
               Top             =   3600
               Width           =   3615
            End
            Begin VB.ListBox lstGUIMADLIEN 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2130
               Left            =   360
               Sorted          =   -1  'True
               TabIndex        =   59
               Top             =   720
               Width           =   3615
            End
            Begin VB.TextBox txtGUIMADLIEN_Scan 
               Height          =   285
               Left            =   360
               TabIndex        =   56
               Top             =   5880
               Width           =   3615
            End
            Begin VB.CommandButton cmdGUIMADLIEN_Load 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Rechercher"
               Height          =   525
               Left            =   2400
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   6360
               Width           =   1575
            End
            Begin VB.CommandButton cmdGUIMADLIEN_Quit 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Abandonner"
               Height          =   525
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   6360
               Width           =   1575
            End
            Begin VB.Label lblGUIMADLIEN_Display 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "MAD déjà rapprochées sur ce dossier"
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
               Left            =   360
               TabIndex        =   62
               Top             =   3240
               Width           =   3615
            End
            Begin VB.Label lblGUIMADLIEN 
               Alignment       =   2  'Center
               BackColor       =   &H00F0F0F0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Sélectionner un bénéficiaire "
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
               Left            =   360
               TabIndex        =   57
               Top             =   360
               Width           =   3615
            End
         End
         Begin VB.Frame fraUpdate 
            BackColor       =   &H00F0F0F0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   7215
            Left            =   8400
            TabIndex        =   11
            Top             =   1150
            Width           =   5175
            Begin VB.Frame fraUpdate_B 
               BackColor       =   &H00D0D0D0&
               Height          =   3975
               Left            =   120
               TabIndex        =   16
               Top             =   3120
               Width           =   4935
               Begin VB.CheckBox chkUpdate_GUIMADLIEN 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "Effacer le lien bénéficiaire"
                  Height          =   375
                  Left            =   2400
                  TabIndex        =   60
                  Top             =   2520
                  Width           =   2295
               End
               Begin VB.TextBox txtUpdate_GUIMADMOT 
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   51
                  Top             =   2160
                  Width           =   3735
               End
               Begin VB.CheckBox chkUpdate_GUIMADTIN 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "pour compte de"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   50
                  Top             =   1080
                  Width           =   1455
               End
               Begin VB.CommandButton cmdUpdate_Annuler 
                  BackColor       =   &H000000FF&
                  Caption         =   "Annuler/Reprendre"
                  Height          =   525
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   44
                  Top             =   2640
                  Width           =   1575
               End
               Begin VB.CommandButton cmdUpdate_Quit 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Abandonner"
                  Height          =   525
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   37
                  Top             =   3240
                  Width           =   1575
               End
               Begin VB.CommandButton cmdUpdate_Ok 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Enregistrer"
                  Height          =   645
                  Left            =   2880
                  Style           =   1  'Graphical
                  TabIndex        =   36
                  Top             =   2880
                  Width           =   1575
               End
               Begin VB.ComboBox cboUpdate_GUIMADMOT 
                  Height          =   315
                  Left            =   1080
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   35
                  Top             =   1680
                  Width           =   3255
               End
               Begin VB.TextBox txtUpdate_GUIMADTIN 
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   33
                  Top             =   1320
                  Width           =   3735
               End
               Begin VB.TextBox txtUpdate_GUIMADTDO 
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   31
                  Top             =   240
                  Width           =   3735
               End
               Begin VB.TextBox txtUpdate_GUIESPTI1 
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   30
                  Top             =   720
                  Width           =   3735
               End
               Begin VB.Label lblUpdate_GUIMADMOT 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "Motif"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   34
                  Top             =   1800
                  Width           =   855
               End
               Begin VB.Label lblUpdate_GUIMADTDO 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "D.Ordre"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   32
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label lblUpdate_GUIESPTI1 
                  BackColor       =   &H00D0D0D0&
                  Caption         =   "Bénéficiaire"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   29
                  Top             =   720
                  Width           =   975
               End
            End
            Begin VB.Frame fraUpdate_A 
               BackColor       =   &H00F0F0F0&
               Height          =   3015
               Left            =   120
               TabIndex        =   12
               Top             =   120
               Width           =   4935
               Begin VB.TextBox txtUpdate_GUIMADLIEN 
                  Height          =   285
                  Left            =   4200
                  TabIndex        =   27
                  Top             =   2160
                  Width           =   495
               End
               Begin VB.TextBox txtUpdate_GUIMADSTA 
                  Height          =   285
                  Left            =   2760
                  TabIndex        =   26
                  Top             =   2160
                  Width           =   1215
               End
               Begin MSComCtl2.DTPicker txtUpdate_GUIESPDJO 
                  Height          =   300
                  Left            =   1320
                  TabIndex        =   25
                  Top             =   2160
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   16777215
                  CalendarForeColor=   0
                  CalendarTitleBackColor=   8421504
                  CalendarTitleForeColor=   16777215
                  CalendarTrailingForeColor=   12632256
                  CustomFormat    =   "dd  MM yyy"
                  Format          =   121700355
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin VB.TextBox txtUpdate_GUIESPCP1 
                  Height          =   285
                  Left            =   2640
                  TabIndex        =   23
                  Top             =   1200
                  Width           =   2055
               End
               Begin VB.TextBox txtUpdate_GUIESPCL1 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   22
                  Top             =   1200
                  Width           =   855
               End
               Begin VB.TextBox txtUpdate_GUIESPMON 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   2640
                  TabIndex        =   20
                  Top             =   720
                  Width           =   2055
               End
               Begin VB.TextBox txtUpdate_GUIESPDEV 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   19
                  Top             =   720
                  Width           =   615
               End
               Begin VB.TextBox txtUpdate_GUIESPDOS 
                  Height          =   285
                  Left            =   3000
                  TabIndex        =   17
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.TextBox txtUpdate_GUIESPOPE 
                  Height          =   285
                  Left            =   1320
                  TabIndex        =   14
                  Top             =   240
                  Width           =   615
               End
               Begin VB.TextBox txtUpdate_GUIESPNAT 
                  Height          =   285
                  Left            =   2160
                  TabIndex        =   13
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label lblUpdate_GUIMADUSR 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "MàJ par"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   64
                  Top             =   2640
                  Width           =   1095
               End
               Begin VB.Label libUpdate_GUIMADUSR 
                  BackColor       =   &H00F0F0F0&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   255
                  Left            =   1320
                  TabIndex        =   63
                  Top             =   2640
                  Width           =   2775
               End
               Begin VB.Label lblUpdate_GUIESPDJO 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Créé le"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   28
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.Label libUpdate_GUIESPCL1 
                  BackColor       =   &H00D0D0D0&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "x"
                  Height          =   255
                  Left            =   1320
                  TabIndex        =   24
                  Top             =   1680
                  Width           =   3375
               End
               Begin VB.Label lblUpdate_GUIESPCL1 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Client"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   21
                  Top             =   1320
                  Width           =   975
               End
               Begin VB.Label lblUpdate_GUIESPMON 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Montant"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   18
                  Top             =   840
                  Width           =   975
               End
               Begin VB.Label lblUpdate_GUIESPDOS 
                  BackColor       =   &H00F0F0F0&
                  Caption         =   "Dossier"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF80FF&
                  Height          =   495
                  Left            =   120
                  TabIndex        =   15
                  Top             =   240
                  Width           =   1095
               End
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7185
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Width           =   13440
            _ExtentX        =   23707
            _ExtentY        =   12674
            _Version        =   393216
            Rows            =   1
            Cols            =   12
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777210
            ForeColor       =   8388608
            BackColorFixed  =   16776921
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"BIA_GUIMAD.frx":0044
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
         Begin VB.ComboBox cboSelect_SQL 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   9
            Text            =   "cboSelect_SQL"
            Top             =   260
            Width           =   3015
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Rechercher"
            Height          =   525
            Left            =   11640
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1815
         End
         Begin VB.Frame fraSelect_Options_1 
            Height          =   1005
            Left            =   3240
            TabIndex        =   6
            Top             =   120
            Width           =   8355
            Begin VB.CheckBox chkSelect_GUIMADMOT 
               Caption         =   "Cumul motifs"
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox txtSelect_GUIMADTIN 
               Height          =   285
               Left            =   6480
               TabIndex        =   48
               Top             =   600
               Width           =   1815
            End
            Begin VB.CheckBox chkSelect_GUIESPDJO 
               Caption         =   "Période de création"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox txtSelect_GUIESPTI1 
               Height          =   285
               Left            =   6480
               TabIndex        =   40
               Top             =   240
               Width           =   1815
            End
            Begin VB.TextBox txtSelect_GUIESPCL1 
               Height          =   285
               Left            =   3720
               TabIndex        =   39
               Top             =   600
               Width           =   1215
            End
            Begin MSComCtl2.DTPicker txtSelect_GUIESPDJO 
               Height          =   300
               Left            =   2040
               TabIndex        =   38
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   121896963
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_GUIESPDJO_Max 
               Height          =   300
               Left            =   2040
               TabIndex        =   46
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CustomFormat    =   "dd  MM yyy"
               Format          =   121896963
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_GUIMADTIN 
               Caption         =   "pour compte de (%)"
               Height          =   255
               Left            =   4920
               TabIndex        =   47
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label lblSelect_GUIESPTI1 
               Caption         =   "Bénéficiaire (%)"
               Height          =   255
               Left            =   5040
               TabIndex        =   42
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lblSelect_GUIESPCL1 
               Caption         =   "Racine client"
               Height          =   255
               Left            =   3720
               TabIndex        =   41
               Top             =   240
               Width           =   975
            End
         End
         Begin MSComCtl2.DTPicker txtSelect_AmjMin 
            Height          =   300
            Left            =   11760
            TabIndex        =   10
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CheckBox        =   -1  'True
            CustomFormat    =   "dd  MM yyy"
            Format          =   121765891
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgStatistiques 
         Height          =   8145
         Left            =   -74880
         TabIndex        =   49
         Top             =   600
         Width           =   13440
         _ExtentX        =   23707
         _ExtentY        =   14367
         _Version        =   393216
         Rows            =   1
         Cols            =   17
         FixedCols       =   0
         RowHeightMin    =   300
         BackColor       =   16777210
         ForeColor       =   8388608
         BackColorFixed  =   16776921
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   16777210
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         TextStyleFixed  =   4
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   1
         AllowUserResizing=   3
         FormatString    =   $"BIA_GUIMAD.frx":00F5
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
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "BIA_GUIMAD.frx":01E3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label libRéférenceInterne 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContext_x1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuselect 
      Caption         =   "mnuSelect"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Quit 
         Caption         =   "Abandonner"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint0_All 
         Caption         =   "Imprimer TOUS les courriers"
      End
   End
End
Attribute VB_Name = "frmBIA_GUIMAD"
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
Dim BIA_GUIMAD_Aut As typeAuthorization
Dim blnTransaction As Boolean
Dim blnAuto As Boolean, blnAuto_Ok As Boolean
Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long
Dim wAmjMin7 As Long, wAmjMax7 As Long


Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnSetfocus As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim mCLIENARA1 As String

'______________________________________________________________________

Dim xYGUIMAD0 As typeYGUIMAD0, meYGUIMAD0 As typeYGUIMAD0
Dim newYGUIMAD0 As typeYGUIMAD0, oldYGUIMAD0 As typeYGUIMAD0
Dim arrYGUIMAD0() As typeYGUIMAD0, arrYGUIMAD0_Nb As Long, arrYGUIMAD0_Max As Long, arrYGUIMAD0_Index As Long
Dim selYGUIMAD0() As typeYGUIMAD0, selYGUIMAD0_Nb As Long, selYGUIMAD0_Max As Long, selYGUIMAD0_Index As Long
Dim xZCLIENA0 As typeZCLIENA0

Dim xYBIAMVT0 As typeYBIAMVT0
Dim cmdSelect_Ok_Caption As String
Dim cmdSelect_SQL_K As String

Dim curDB As Currency, curCR As Currency
Dim selZCHGOPE0() As typeZCHGOPE0, selZCHGOPE0_Nb As Long, selZCHGOPE0_Max As Long, selZCHGOPE0_Index As Long
Dim xZCHGOPE0 As typeZCHGOPE0

Dim rsSabX As New ADODB.Recordset

Dim arrGUIESPMON() As Currency, arrGUIESPNB() As Long
Dim wMM As Integer, wAAAA As Integer, arrGUIESPMON_Dev As Integer

Dim fgStatistiques_FormatString As String, fgStatistiques_K As Integer
Dim fgStatistiques_RowDisplay As Integer, fgStatistiques_RowClick As Integer, fgStatistiques_ColClick As Integer
Dim fgStatistiques_ColorClick As Long, fgStatistiques_ColorDisplay As Long
Dim fgStatistiques_Sort1 As Integer, fgStatistiques_Sort2 As Integer
Dim fgStatistiques_SortAD As Integer, fgStatistiques_Sort1_Old As Integer
Dim fgStatistiques_arrIndex As Integer
Dim blnfgStatistiques_DisplayLine As Boolean

Dim meCV1 As typeCV, meCV2 As typeCV

'______________________________________________________________________

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
        For I = fgSelect_arrIndex To 0 Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.LeftCol = 0
    End If
End If

End Sub
Private Sub fgSelect_Display()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
cmdPrint.Enabled = False
currentAction = "fgselect_Display"

For I = 1 To arrYGUIMAD0_Nb

        xYGUIMAD0 = arrYGUIMAD0(I)
    
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
Next I

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgSelect.Rows - 1): DoEvents
If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
    cmdPrint.Enabled = True
End If
fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Private Sub fgStatistiques_Bénéficiaire_Display()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = "<Bénéficiaire                        |> Nombre      |>           Montant  |                                    ||"
cmdPrint.Enabled = False
currentAction = "fgselect_Display"

For I = 1 To arrYGUIMAD0_Nb

        xYGUIMAD0 = arrYGUIMAD0(I)
    
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgStatistiques_Bénéficiaire_DisplayLine I
Next I

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgSelect.Rows - 1): DoEvents
If fgSelect.Rows > 1 Then
    fgSelect_Sort1_Old = fgSelect_arrIndex - 1
    fgSelect_SortAD = 5
    fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgStatistiques_Bénéficiaire_SortX 2
    cmdPrint.Enabled = True
End If
fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub fgStatistiques_Display()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long
Dim xMM As Integer, xAAAA As Integer

On Error GoTo Error_Handler
SSTab1.Tab = 1
fgStatistiques.Visible = False
fgStatistiques_Reset
fgStatistiques.Rows = 1
wAAAA = wAMJMin / 10000 + 1900
wMM = wAMJMin / 100 Mod 100
xMM = wMM
xAAAA = wAAAA
fgStatistiques_FormatString = "Client    |Dev |"

For I = 1 To 12
    fgStatistiques_FormatString = fgStatistiques_FormatString & ">" & Format$(xMM, "00") & "." & xAAAA & "        |"
    xMM = xMM + 1
    If xMM = 13 Then xMM = 1: xAAAA = xAAAA + 1
Next I
fgStatistiques_FormatString = fgStatistiques_FormatString & ">Total            |Dev   |Client    |"
fgStatistiques.FormatString = fgStatistiques_FormatString
cmdPrint.Enabled = False
currentAction = "fgStatistiques_Display"

For I = 1 To arrYGUIMAD0_Nb

     xYGUIMAD0 = arrYGUIMAD0(I)

        fgStatistiques.Rows = fgStatistiques.Rows + 1
        fgStatistiques.Row = fgStatistiques.Rows - 1
        fgStatistiques_DisplayLine I
Next I

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgStatistiques.Rows - 1): DoEvents
If fgStatistiques.Rows > 1 Then
'    fgStatistiques_Sort1 = 0: fgStatistiques_Sort2 = 2: fgStatistiques_Sort
    cmdPrint.Enabled = True
End If
fgStatistiques.LeftCol = 9

fgStatistiques.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


Private Sub lstSelect_Load_1()
Dim I As Long, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_1"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True
chkSelect_GUIMADMOT.Visible = False

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub lstSelect_Load_3()
Dim I As Long, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_3"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Private Sub lstSelect_Load_4()
Dim I As Long, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_4"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


Private Sub lstSelect_Load_7()
Dim I As Long, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_7"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True
chkSelect_GUIESPDJO = "1"

Call DTPicker_Set(txtSelect_GUIESPDJO_Max, YBIATAB0_DATE_CPT_JS1)
Call DTPicker_Set(txtSelect_GUIESPDJO, dateElp("Ouvré", -10, YBIATAB0_DATE_CPT_JS1))

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
Dim X As String, wColor As Long

On Error Resume Next
Select Case xYGUIMAD0.GUIMADSTA
    Case Is = "V": wColor = vbBlue
    Case Is = "A": wColor = vbGrayText
    Case Else: wColor = vbMagenta
End Select

fgSelect.Col = 0: fgSelect.Text = xYGUIMAD0.GUIMADID
fgSelect.CellForeColor = wColor
X = xYGUIMAD0.GUIESPOPE & " " & xYGUIMAD0.GUIESPNAT & " " & Format$(xYGUIMAD0.GUIESPDOS, "#####0#")
fgSelect.Col = 1: fgSelect.Text = X
fgSelect.CellForeColor = wColor
X = Format$(Abs(xYGUIMAD0.GUIESPMON), "### ### ### ###.00")
fgSelect.Col = 2: fgSelect.Text = X
fgSelect.CellForeColor = vbRed
fgSelect.Col = 3: fgSelect.Text = xYGUIMAD0.GUIESPDEV
fgSelect.CellForeColor = wColor
fgSelect.Col = 4: fgSelect.Text = xYGUIMAD0.GUIESPCL1
fgSelect.CellForeColor = wColor

fgSelect.Col = 6: fgSelect.Text = dateIBM10(xYGUIMAD0.GUIESPDJO, True)
fgSelect.CellForeColor = wColor
fgSelect.Col = 7: fgSelect.Text = xYGUIMAD0.GUIESPTI1
fgSelect.CellForeColor = wColor
fgSelect.Col = 8: fgSelect.Text = Format$(xYGUIMAD0.GUIMADLIEN, "####")
fgSelect.CellForeColor = wColor
fgSelect.Col = 9: fgSelect.Text = xYGUIMAD0.GUIMADMOT
fgSelect.CellForeColor = wColor

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

If Trim(xYGUIMAD0.GUIESPCL1) = "" Or xYGUIMAD0.GUIESPCL1 = "0010000" Then
    fgSelect.Col = 5: fgSelect.Text = xYGUIMAD0.GUIMADTDO
Else
    X = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & xYGUIMAD0.GUIESPCL1 & "'"
    Set rsSab = cnsab.Execute(X)
    
    If Not rsSab.EOF Then fgSelect.Col = 5: fgSelect.Text = rsSab("CLIENARA1")
End If
fgSelect.CellForeColor = wColor
End Sub

Public Sub fgStatistiques_Bénéficiaire_DisplayLine(lIndex As Long)
Dim X As String, wColor As Long, xWhere As String, xSQL As String
On Error Resume Next


'wColor = vbBlue

fgSelect.Col = 0: fgSelect.Text = xYGUIMAD0.GUIESPTI1
'fgSelect.CellForeColor = wColor
fgSelect.Col = 1: fgSelect.Text = Format$(xYGUIMAD0.GUIESPDOS, "### ### ##0")
'fgSelect.CellForeColor = wColor
X = Format$(xYGUIMAD0.GUIMADMON, "### ### ### ###.00")
fgSelect.Col = 2: fgSelect.Text = X
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
End Sub



Public Sub fgStatistiques_DisplayLine(lIndex As Long)
Dim X As String
Dim K As Integer
Dim xCol1 As String
On Error Resume Next

fgStatistiques.Col = 0: fgStatistiques.Text = xYGUIMAD0.GUIESPCL1
If chkSelect_GUIMADMOT <> "1" Then
    xCol1 = xYGUIMAD0.GUIESPDEV
Else
     xCol1 = xYGUIMAD0.GUIESPDEV & "-" & Left$(xYGUIMAD0.GUIMADMOT, 1)
End If
fgStatistiques.Col = 1: fgStatistiques.Text = xCol1
For K = 1 To 13

    If arrGUIESPMON(lIndex, K) <> 0 Then
        If K = 13 Then
            fgStatistiques.Col = 14
        Else
            fgStatistiques.Col = K + 2 - wMM
            If fgStatistiques.Col < 2 Then fgStatistiques.Col = fgStatistiques.Col + 12
        End If
        X = Trim(Format$(Abs(arrGUIESPMON(lIndex, K)), "### ### ### ###"))
        fgStatistiques.Text = X
    End If
Next K
fgStatistiques.Col = 16: fgStatistiques.Text = xYGUIMAD0.GUIESPCL1
fgStatistiques.Col = 15: fgStatistiques.Text = xCol1
End Sub


Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = fgSelect.Cols - 1
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

End Sub


Public Sub fgStatistiques_Reset()
fgStatistiques.Clear
fgStatistiques_Sort1 = 0: fgStatistiques_Sort2 = 0
fgStatistiques_Sort1_Old = -1
fgStatistiques_RowDisplay = 0: fgStatistiques_RowClick = 0
fgStatistiques_arrIndex = fgStatistiques.Cols - 1
blnfgStatistiques_DisplayLine = False
fgStatistiques_SortAD = 6
fgStatistiques.LeftCol = 0

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
Dim I As Integer, X As String
Dim wIndex As Integer
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    wIndex = Val(fgSelect.Text)
    Select Case lK
        Case 0: X = Format$(arrYGUIMAD0(wIndex).GUIMADID, "000000000")
        Case 2: X = Format$(arrYGUIMAD0(wIndex).GUIESPMON, "000000000000000.00")
        Case 3: X = arrYGUIMAD0(wIndex).GUIESPDEV & Format$(arrYGUIMAD0(wIndex).GUIESPMON, "000000000000000.00")
        Case 6: X = arrYGUIMAD0(wIndex).GUIESPDJO
    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub
Public Sub fgStatistiques_Bénéficiaire_SortX(lK As Integer)
Dim I As Integer, X As String
Dim wIndex As Integer
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = lK
    Select Case lK
        Case 1: X = Format$(Val(fgSelect.Text), "000000000")
        Case 2: X = Format$(Val(fgSelect.Text), "000000000000000.00")
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


Public Sub cmdContext_Quit()
If fraGUIMADLIEN.Visible Then fraGUIMADLIEN.Visible = False: Exit Sub
If fraUpdate.Visible Then fraUpdate.Visible = False: Exit Sub
If fgSelect.Visible Then fgSelect.Visible = False: cmdSelect_Ok.Caption = "Extraire les mouvements": Exit Sub
Unload Me
End Sub




Private Sub cboSelect_SQL_Click()
cmdSelect_SQL_K = Mid$(cboSelect_SQL, 1, 1)
If blnControl Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    lstSelect.Visible = False
    txtSelect_AMJMIN.Visible = False
    fraSelect_Options_1.Visible = False
    fraUpdate.Visible = False
    fgStatistiques.Visible = False
    Select Case cmdSelect_SQL_K
        Case "1": lstSelect_Load_1
        Case "2": lstSelect_Load_2
        Case "3": lstSelect_Load_3
        Case "4": lstSelect_Load_4
        Case "6": lstSelect_Load_6
        Case "7": lstSelect_Load_7
        Case "8": lstSelect_Load_8
        Case "9": lstSelect_Load_9
    End Select
    Me.Enabled = True: Me.MousePointer = 0
End If
End Sub


Private Sub chkSelect_GUIESPDJO_Click()
If chkSelect_GUIESPDJO = "1" Then
    If cmdSelect_SQL_K = "1" Or cmdSelect_SQL_K = "7" Then txtSelect_GUIESPDJO.Visible = True
    txtSelect_GUIESPDJO_Max.Visible = True
Else
    txtSelect_GUIESPDJO.Visible = False
    txtSelect_GUIESPDJO_Max.Visible = False
End If


End Sub

Private Sub chkUpdate_GUIMADTIN_Click()
If chkUpdate_GUIMADTIN = "1" Then
    txtUpdate_GUIMADTIN.Visible = True
Else
    txtUpdate_GUIMADTIN.Visible = False
End If

End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdGUIMADLIEN_Load_Click()
Dim X As String, K As Integer, blnOk As Boolean
Dim xWhere As String, xSQL As String
Dim wGUIMADID As Long

Me.Enabled = False: Me.MousePointer = vbHourglass
txtGUIMADLIEN_Scan = Replace(txtGUIMADLIEN_Scan, "'", " ")
lstGUIMADLIEN.Clear

xWhere = " where GUIMADLIEN = 0 "
K = 0
blnOk = False
Do
    X = Space_Scan(txtGUIMADLIEN_Scan, K)
    If X = "" Then
        blnOk = True
    Else
        xWhere = xWhere & " and GUIESPTI1 like '%" & X & "%'"
    End If
Loop Until blnOk


Set rsSab = Nothing

xSQL = "select GUIMADID,GUIESPTI1 from " & paramIBM_Library_SABSPE & ".YGUIMAD0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    wGUIMADID = rsSab("GUIMADID")
    If wGUIMADID <> xYGUIMAD0.GUIMADID Then
        lstGUIMADLIEN.AddItem Format$(wGUIMADID, "000000") & " " & rsSab("GUIESPTI1")
    End If
    rsSab.MoveNext

Loop

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdGUIMADLIEN_Quit_Click()
fraGUIMADLIEN.Visible = False
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
Dim I As Integer

blnControl = False
usrColor_Set

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False
currentAction = ""

blnAuto = False
blnAuto_Ok = False
lstSelect.Visible = False
cmdSelect_Ok.Caption = "Extraire les mouvements"

libRéférenceInterne = ""
cboSelect_SQL.ListIndex = 0
fgSelect.Visible = False
Call DTPicker_Set(txtSelect_AMJMIN, YBIATAB0_DATE_CPT_JS1)
Call DTPicker_Set(txtSelect_GUIESPDJO, YBIATAB0_DATE_CPT_JS1)
Call DTPicker_Set(txtSelect_GUIESPDJO_Max, YBIATAB0_DATE_CPT_JS1)
fraUpdate.Visible = False
cboSelect_SQL.ListIndex = 1
blnControl = True
cboSelect_SQL.ListIndex = 0
chkUpdate_GUIMADTIN = "0"
chkSelect_GUIMADMOT.Visible = False
End Sub
Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

SSTab1.Tab = 0

blnControl = False

fgSelect_FormatString = fgSelect.FormatString
fgStatistiques_FormatString = fgStatistiques.FormatString
cmdSelect_Ok.Visible = False
fraSelect_Options_1.Visible = False
fraGUIMADLIEN.Visible = False
lstGUIMADLIEN_Display.Enabled = False
lstGUIMADLIEN_Display.ForeColor = vbMagenta
txtSelect_GUIESPDJO.Visible = False
txtSelect_GUIESPDJO_Max.Visible = False
cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1 - Consultation des MAD"
cboSelect_SQL.AddItem "2 - Import des MAD : SAB => BIA_GUIMAD"
cboSelect_SQL.AddItem "3 - Liste des MAD à compléter"
cboSelect_SQL.AddItem "4 - Liste des bénéficiaires"
cboSelect_SQL.AddItem "6 - Stat DO par année civile"
cboSelect_SQL.AddItem "7 - Statistisques bénéficiaires"
If BIA_GUIMAD_Aut.Xspécial Then
    cboSelect_SQL.AddItem "8 - Alertes MAD incomplètes ou annulées"
    cboSelect_SQL.AddItem "9 - Informations auprès des gestionnaires"
End If
If BIA_GUIMAD_Aut.Valider Then cboSelect_SQL.AddItem "M - modification Texte"

cboUpdate_GUIMADMOT.Clear
Call cbo_Load("GDMP", "MAD_Motif", cboUpdate_GUIMADMOT, 3)
cboUpdate_GUIMADMOT.AddItem ""
cmdReset



End Sub

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


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Dim Msg As String
Dim I As Integer

Me.Enabled = False: Me.MousePointer = vbHourglass
    Select Case cmdSelect_SQL_K
        Case "1": cmdPrint_Ok_1
        Case "4": cmdPrint_Ok_1
        Case "6": cmdPrint_Ok_6
        Case "7": cmdPrint_Ok_7
    End Select

Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

Me.Enabled = False: Me.MousePointer = vbHourglass
    blnOk = Not fgSelect.Visible
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CDR_cmdSelect_Ok ........"): DoEvents
cmdSelect_Ok.Visible = False
fraUpdate.Visible = False
fraSelect_Options_1.Enabled = False
txtSelect_AMJMIN.Enabled = False
fgSelect.Clear
DoEvents
If blnOk Then
    cmdSelect_Ok.Caption = "Modifier les options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    lstSelect.BackColor = &H8000000F
    Call usrColor_Container(lstSelect, lstSelect.BackColor)
    Select Case cmdSelect_SQL_K
        Case "1": cmdSelect_SQL
        Case "2": cmdSelect_SQL_2
        Case "3": cmdSelect_SQL_3
        Case "4": cmdSelect_SQL_4
        Case "6": cmdSelect_SQL_6
        Case "7": cmdSelect_SQL_7
        Case "8": cmdSelect_SQL_8
        Case "9": cmdSelect_SQL_9
        Case "M": cmdSelect_SQL_M
    End Select

    fgSelect.Enabled = True
Else
    cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
    cmdSelect_Ok.BackColor = &HC0FFC0
    lstSelect.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(lstSelect, lstSelect.BackColor)
    fgSelect.Visible = False
    fgSelect.Enabled = False
    fraSelect_Options_1.Enabled = True
    txtSelect_AMJMIN.Enabled = True

End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CDR_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
cmdSelect_Ok.Visible = True

End Sub


Private Sub cmdSelect_SQL()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL"
xWhere = " where GUIMADID > 0 "

Set rsSab = Nothing
Call DTPicker_Control(txtSelect_GUIESPDJO, wAMJMin)
Call DTPicker_Control(txtSelect_GUIESPDJO_Max, WAMJMax)

If chkSelect_GUIESPDJO = "1" Then
    xWhere = xWhere & " and GUIESPDJO >= " & wAMJMin - 19000000 _
                    & " and GUIESPDJO <= " & WAMJMax - 19000000
End If
X = Trim(txtSelect_GUIESPCL1)
If X <> "" Then xWhere = xWhere & " and GUIESPCL1 like '%" & X & "%'"
X = Trim(txtSelect_GUIESPTI1)
If X <> "" Then xWhere = xWhere & " and GUIESPTI1 like '%" & X & "%'"
X = Trim(txtSelect_GUIMADTIN)
If X <> "" Then xWhere = xWhere & " and GUIMADTIN like '%" & X & "%'"

arrYGUIMAD0_SQL xWhere & " order by GUIESPDOS,GUIMADID"
    
fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_6()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
Dim lDate As Long
On Error GoTo Error_Handler

Set rsSab = Nothing
currentAction = "cmdSelect_SQL_6"
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents

currentAction = "cmdSelect_SQL_6"
xWhere = " where GUIMADID > 0 and GUIMADSTA <> 'A'"

Set rsSab = Nothing
'Call DTPicker_Control(txtSelect_GUIESPDJO, wAmjMin)
Call DTPicker_Control(txtSelect_GUIESPDJO_Max, WAMJMax)
wAMJMin = Val(Mid$(WAMJMax, 1, 4) + "0100" - 19000000)

lDate = Val(Mid$(WAMJMax, 1, 6)) * 100 - 19000000 + 99
WAMJMax = lDate
'$20070903_$JPL wAmjMin = lDate - 10000

'If chkSelect_GUIESPDJO = "1" Then
    xWhere = xWhere & " and GUIESPDJO >= " & wAMJMin _
                    & " and GUIESPDJO <= " & WAMJMax
'End If
X = Trim(txtSelect_GUIESPCL1)
If X <> "" Then xWhere = xWhere & " and GUIESPCL1 like '%" & X & "%'"
X = Trim(txtSelect_GUIESPTI1)
If X <> "" Then xWhere = xWhere & " and GUIESPTI1 like '%" & X & "%'"
X = Trim(txtSelect_GUIMADTIN)
If X <> "" Then xWhere = xWhere & " and GUIMADTIN like '%" & X & "%'"

xSQL = "select count(distinct concat(GUIESPCL1,GUIESPDEV)) as Tally  from " & paramIBM_Library_SABSPE & ".YGUIMAD0 " & xWhere

Set rsSab = cnsab.Execute(xSQL)
K = rsSab("Tally") + 20
If chkSelect_GUIMADMOT = "1" Then K = K * 10
ReDim arrGUIESPMON(K, 13)
ReDim arrGUIESPNB(K, 13)
ReDim arrYGUIMAD0(K)

Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YGUIMAD0 " & xWhere & " order by GUIESPCL1,GUIESPDEV"
If chkSelect_GUIMADMOT = "1" Then xSQL = xSQL & ",substr(GUIMADMOT,1,4)"
Set rsSab = cnsab.Execute(xSQL)
    
cmdSelect_SQL_6_Cumul

fgStatistiques_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub cmdSelect_SQL_6_Cumul()
Dim V
Dim xSQL As String, K As Long, I As Long
Dim xWhere As String, xAnd As String
Dim kDev As Integer, kEUR As Integer, kUSD As Integer
Dim blnRupture As Boolean
Dim wGUIMADMOT As String

On Error GoTo Error_Handler

wGUIMADMOT = "?"
arrYGUIMAD0_Nb = 0
Do While Not rsSab.EOF
        blnRupture = False
        If xYGUIMAD0.GUIESPCL1 <> rsSab("GUIESPCL1") _
        Or xYGUIMAD0.GUIESPDEV <> rsSab("GUIESPDEV") Then blnRupture = True
        If chkSelect_GUIMADMOT = "1" Then
            If wGUIMADMOT <> Left$(rsSab("GUIMADMOT"), 1) Then blnRupture = True
        End If
        If blnRupture Then
            arrYGUIMAD0_Nb = arrYGUIMAD0_Nb + 1
            V = rsYGUIMAD0_GetBuffer(rsSab, xYGUIMAD0)

            If Not IsNull(V) Then
                MsgBox V, vbCritical, "cmdSelect_SQL_6_Cumul"
            Else
                wGUIMADMOT = Left$(xYGUIMAD0.GUIMADMOT, 1)
                arrYGUIMAD0(arrYGUIMAD0_Nb) = xYGUIMAD0
            End If
        End If
        
     K = rsSab("GUIESPDJO") / 100 Mod 100
     arrGUIESPMON(arrYGUIMAD0_Nb, K) = arrGUIESPMON(arrYGUIMAD0_Nb, K) + rsSab("GUIESPMON")
     arrGUIESPNB(arrYGUIMAD0_Nb, K) = arrGUIESPNB(arrYGUIMAD0_Nb, K) + 1
    rsSab.MoveNext

Loop

arrGUIESPMON_Dev = arrYGUIMAD0_Nb + 2
kEUR = arrYGUIMAD0_Nb + 1
kUSD = arrYGUIMAD0_Nb + 2
arrYGUIMAD0(kEUR).GUIESPDEV = "EUR": arrYGUIMAD0(kEUR).GUIESPCL1 = ""
arrYGUIMAD0(kUSD).GUIESPDEV = "USD": arrYGUIMAD0(kUSD).GUIESPCL1 = ""


For I = 1 To arrYGUIMAD0_Nb
    Select Case arrYGUIMAD0(I).GUIESPDEV
        Case "EUR": kDev = kEUR
        Case "USD": kDev = kUSD
        Case Else: kDev = cmdSelect_SQL_6_Cumul_Dev(arrYGUIMAD0(I).GUIESPDEV)
    End Select
    For K = 1 To 12
        If arrGUIESPMON(I, K) <> 0 Then
            arrGUIESPMON(I, K) = Round(arrGUIESPMON(I, K), 0)
            arrGUIESPMON(I, 13) = arrGUIESPMON(I, 13) + arrGUIESPMON(I, K)
            arrGUIESPNB(I, 13) = arrGUIESPNB(I, 13) + arrGUIESPNB(I, K)
            
            arrGUIESPMON(kDev, K) = arrGUIESPMON(kDev, K) + arrGUIESPMON(I, K)
            arrGUIESPMON(kDev, 13) = arrGUIESPMON(kDev, 13) + arrGUIESPMON(I, K)
            arrGUIESPNB(kDev, K) = arrGUIESPNB(kDev, K) + arrGUIESPNB(I, K)
            arrGUIESPNB(kDev, 13) = arrGUIESPNB(kDev, 13) + arrGUIESPNB(I, K)
       End If
    Next K
Next I



Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_3()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL"
xWhere = " where GUIMADID > 0 and GUIMADSTA = ' '"

Set rsSab = Nothing

arrYGUIMAD0_SQL xWhere & " order by GUIESPDOS,GUIMADID"
    
fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_4()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL"
xWhere = " where GUIMADID > 0 and GUIMADLIEN = 0"

Call DTPicker_Control(txtSelect_GUIESPDJO, wAMJMin)
Call DTPicker_Control(txtSelect_GUIESPDJO_Max, WAMJMax)

If chkSelect_GUIESPDJO = "1" Then
    xWhere = xWhere & " and GUIESPDJO >= " & wAMJMin - 19000000 _
                    & " and GUIESPDJO <= " & WAMJMax - 19000000
End If
X = Trim(txtSelect_GUIESPCL1)
If X <> "" Then xWhere = xWhere & " and GUIESPCL1 like '%" & X & "%'"
X = Trim(txtSelect_GUIESPTI1)
If X <> "" Then xWhere = xWhere & " and GUIESPTI1 like '%" & X & "%'"
X = Trim(txtSelect_GUIMADTIN)
If X <> "" Then xWhere = xWhere & " and GUIMADTIN like '%" & X & "%'"

arrYGUIMAD0_SQL xWhere & " order by GUIESPDOS,GUIMADID"
    
fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_7()
Dim V, blnNew As Boolean
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL"
xWhere = " where GUIMADID > 0 "

Call DTPicker_Control(txtSelect_GUIESPDJO, wAMJMin)
Call DTPicker_Control(txtSelect_GUIESPDJO_Max, WAMJMax)

If chkSelect_GUIESPDJO = "1" Then
    xWhere = xWhere & " and GUIESPDJO >= " & wAMJMin - 19000000 _
                    & " and GUIESPDJO <= " & WAMJMax - 19000000
End If
X = Trim(txtSelect_GUIESPCL1)
If X <> "" Then xWhere = xWhere & " and GUIESPCL1 like '%" & X & "%'"
X = Trim(txtSelect_GUIESPTI1)
If X <> "" Then xWhere = xWhere & " and GUIESPTI1 like '%" & X & "%'"
X = Trim(txtSelect_GUIMADTIN)
If X <> "" Then xWhere = xWhere & " and GUIMADTIN like '%" & X & "%'"

ReDim arrYGUIMAD0(101): arrYGUIMAD0_Max = 100: arrYGUIMAD0_Nb = 0

Set rsSab = Nothing
rsYGUIMAD0_Init xYGUIMAD0
xSQL = "select GUIMADID,GUIMADLIEN,GUIMADMON,GUIESPTI1 from " & paramIBM_Library_SABSPE & ".YGUIMAD0 " & xWhere & " order by GUIMADLIEN,GUIMADID"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    xYGUIMAD0.GUIMADID = rsSab("GUIMADID")
    xYGUIMAD0.GUIMADLIEN = rsSab("GUIMADLIEN")
    xYGUIMAD0.GUIMADMON = rsSab("GUIMADMON")
    xYGUIMAD0.GUIESPTI1 = rsSab("GUIESPTI1")
    If xYGUIMAD0.GUIMADMON = 0 Then MsgBox "Cv Eur = 0", vbCritical, "Dosier : " & xYGUIMAD0.GUIMADID
    If xYGUIMAD0.GUIMADLIEN = 0 Then xYGUIMAD0.GUIMADLIEN = xYGUIMAD0.GUIMADID
    blnNew = True
    For K = 1 To arrYGUIMAD0_Nb
        If arrYGUIMAD0(K).GUIMADLIEN = xYGUIMAD0.GUIMADLIEN Then
            arrYGUIMAD0(K).GUIESPDOS = arrYGUIMAD0(K).GUIESPDOS + 1
            arrYGUIMAD0(K).GUIMADMON = arrYGUIMAD0(K).GUIMADMON + xYGUIMAD0.GUIMADMON
            blnNew = False
            Exit For
        End If
    Next K
    If blnNew Then
        arrYGUIMAD0_Nb = arrYGUIMAD0_Nb + 1
         If arrYGUIMAD0_Nb > arrYGUIMAD0_Max Then
             arrYGUIMAD0_Max = arrYGUIMAD0_Max + 50
             ReDim Preserve arrYGUIMAD0(arrYGUIMAD0_Max)
         End If
         xYGUIMAD0.GUIESPDOS = 1
         arrYGUIMAD0(arrYGUIMAD0_Nb) = xYGUIMAD0
    End If
    rsSab.MoveNext

Loop
    
    
    
fgStatistiques_Bénéficiaire_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_8()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL"
xWhere = " where GUIMADID > 0 and GUIMADSTA = ' '"

Set rsSab = Nothing

arrYGUIMAD0_SQL xWhere & " order by GUIESPDOS,GUIMADID"
    
'fgSelect_Display
'If arrYGUIMAD0_Nb > 0 Then
cmdSelect_SQL_8_Mail

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub cmdSelect_SQL_9()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
Dim mCLIENACLI As String, mCLIENARES As String, mCLIENARA1 As String
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL_9"
Call DTPicker_Control(txtSelect_AMJMIN, wAMJMin)
xWhere = "where GUIESPDJO = " & wAMJMin - 19000000 & " order by GUIESPCP1,GUIMADID"

Set rsSab = Nothing

arrYGUIMAD0_SQL xWhere

mCLIENACLI = "": mCLIENARES = "": mCLIENARA1 = ""
Set rsSabX = Nothing
lstW.Clear
For I = 1 To arrYGUIMAD0_Nb

    If mCLIENACLI <> arrYGUIMAD0(I).GUIESPCL1 Then
        mCLIENACLI = arrYGUIMAD0(I).GUIESPCL1
        If mCLIENACLI = "0010000" Then
            mCLIENARES = "XXX"
            mCLIENARA1 = arrYGUIMAD0(I).GUIMADTDO
        Else
            xWhere = " where CLIENACLI = '" & arrYGUIMAD0(I).GUIESPCL1 & "'"
            xSQL = "select CLIENARES , CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 " & xWhere
            Set rsSabX = cnsab.Execute(xSQL)
            
            If Not rsSabX.EOF Then
                mCLIENARES = rsSabX("CLIENARES")
                mCLIENARA1 = Trim(rsSabX("CLIENARA1"))
            Else
                mCLIENARES = "XXX"
                mCLIENARA1 = "???"
            End If
        End If
    End If
    If mCLIENACLI = "0010000" Then
        mCLIENARES = "XXX"
        mCLIENARA1 = arrYGUIMAD0(I).GUIMADTDO
    End If
    lstW.AddItem mCLIENARES & ";" & arrYGUIMAD0(I).GUIESPCP1 & ";" & I & ";" & mCLIENARA1
Next I
fgSelect_Display
cmdSelect_SQL_9_Mail

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_2()
Dim V
On Error GoTo Error_Handler

Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_2"): DoEvents

currentAction = "cmdSelect_SQL_2"
Call DTPicker_Control(txtSelect_AMJMIN, wAMJMin)
If wAMJMin = "00000000" Then
    MsgBox "Préciser la date", vbInformation, "Import des MAD à une date"
    Exit Sub
End If
    
cmdSelect_SQL_2_ZGUIESP0
cmdSelect_SQL_2_ZCHGOPE0
    
cmdSelect_SQL_3

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_2_ZGUIESP0()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String, X As String
Dim blnOk As Boolean
On Error GoTo Error_Handler

    
xWhere = "where GUIESPDJO = " & wAMJMin - 19000000 & " order by GUIESPDOS,GUIMADID"
arrYGUIMAD0_SQL xWhere

Set rsSab = Nothing
xWhere = " where GUIESPDJO = '" & wAMJMin - 19000000 & " '" _
       & " and GUIESPOPE = 'RE 'and ( GUIESPNAT = '001' or GUIESPNAT = '003' or GUIESPNAT = '009' or GUIESPNAT = '010')"
xSQL = "select * from " & paramIBM_Library_SAB & ".ZGUIESP0 " & xWhere & " order by GUIESPDOS"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    blnOk = False
    X = Mid$(rsSab("GUIESPTI1"), 1, 3)
    If X = "MAD" Or X = "mad" Then
        blnOk = True
        xYGUIMAD0.GUIESPDOS = rsSab("GUIESPDOS")
        xYGUIMAD0.GUIESPOPE = rsSab("GUIESPOPE")
        xYGUIMAD0.GUIESPNAT = rsSab("GUIESPNAT")
        For K = 1 To arrYGUIMAD0_Nb
            If xYGUIMAD0.GUIESPDOS < arrYGUIMAD0(K).GUIESPDOS Then Exit For
            If xYGUIMAD0.GUIESPDOS = arrYGUIMAD0(K).GUIESPDOS Then
                If xYGUIMAD0.GUIESPOPE = arrYGUIMAD0(K).GUIESPOPE _
                And xYGUIMAD0.GUIESPNAT = arrYGUIMAD0(K).GUIESPNAT Then blnOk = False: Exit For
            End If
        Next K
        If blnOk Then YGUIMAD0_Add_ZGUIESP0
    End If
    
    rsSab.MoveNext

Loop
    
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub cmdSelect_SQL_2_ZCHGOPE0()
Dim V
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String, X As String
Dim blnOk As Boolean
On Error GoTo Error_Handler

    
xWhere = "where GUIESPDJO = " & wAMJMin - 19000000 & " order by GUIESPDOS,GUIMADID"
arrYGUIMAD0_SQL xWhere

Set rsSab = Nothing
xWhere = " where CHGOPECRE = '" & wAMJMin - 19000000 & " '" _
       & " and CHGOPEOPE = 'CPT'and ( CHGOPENAT = '110' or CHGOPENAT = '111' )" _
       & " and CHGOPESER = '00' and CHGOPESSE = 'GU'"
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCHGOPE0 " & xWhere & " order by CHGOPEDOS"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    blnOk = False
        blnOk = True
        xYGUIMAD0.GUIESPDOS = rsSab("CHGOPEDOS")
        xYGUIMAD0.GUIESPOPE = rsSab("CHGOPEOPE")
        xYGUIMAD0.GUIESPNAT = rsSab("CHGOPENAT")
        For K = 1 To arrYGUIMAD0_Nb
            If xYGUIMAD0.GUIESPDOS < arrYGUIMAD0(K).GUIESPDOS Then Exit For
            If xYGUIMAD0.GUIESPDOS = arrYGUIMAD0(K).GUIESPDOS Then
                If xYGUIMAD0.GUIESPOPE = arrYGUIMAD0(K).GUIESPOPE _
                And xYGUIMAD0.GUIESPNAT = arrYGUIMAD0(K).GUIESPNAT Then blnOk = False: Exit For
            End If
        Next K
        If blnOk Then YGUIMAD0_Add_ZCHGOPE0
    
    rsSab.MoveNext

Loop
    
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub arrYGUIMAD0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrYGUIMAD0(101)
arrYGUIMAD0_Max = 100: arrYGUIMAD0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YGUIMAD0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYGUIMAD0_GetBuffer(rsSab, xYGUIMAD0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgselect_Display"
        '' Exit Sub
     Else
         arrYGUIMAD0_Nb = arrYGUIMAD0_Nb + 1
         If arrYGUIMAD0_Nb > arrYGUIMAD0_Max Then
             arrYGUIMAD0_Max = arrYGUIMAD0_Max + 50
             ReDim Preserve arrYGUIMAD0(arrYGUIMAD0_Max)
         End If
         
         arrYGUIMAD0(arrYGUIMAD0_Nb) = xYGUIMAD0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdUpdate_Annuler_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

newYGUIMAD0 = oldYGUIMAD0
If newYGUIMAD0.GUIMADSTA = "A" Then
    If Trim(newYGUIMAD0.GUIMADMOT) = "" Then
        newYGUIMAD0.GUIMADSTA = " "
    Else
        newYGUIMAD0.GUIMADSTA = "V"
    End If
Else
    newYGUIMAD0.GUIMADSTA = "A"
End If


    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    V = cmdUpdate_Ok_Transaction
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        arrYGUIMAD0(arrYGUIMAD0_Index) = newYGUIMAD0
        xYGUIMAD0 = newYGUIMAD0
        fgSelect_DisplayLine arrYGUIMAD0_Index
        fraUpdate.Visible = False
        fraGUIMADLIEN.Visible = False

    Else
        MsgBox V, vbCritical, Me.Name & " : cmdUpdate_Ok"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdUpdate_Ok_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

If IsNull(fraUpdate_Control) Then
    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    V = cmdUpdate_Ok_Transaction
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        arrYGUIMAD0(arrYGUIMAD0_Index) = newYGUIMAD0
        xYGUIMAD0 = newYGUIMAD0
        fgSelect_DisplayLine arrYGUIMAD0_Index
        fraUpdate.Visible = False
        fraGUIMADLIEN.Visible = False
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdUpdate_Ok"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdUpdate_Quit_Click()
fraGUIMADLIEN.Visible = False
fraUpdate.Visible = False
End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
Me.Enabled = False
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    If cmdSelect_SQL_K = "7" Then
        Select Case fgSelect.Col
            Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_Sort
            Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgStatistiques_Bénéficiaire_SortX 1
            Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgStatistiques_Bénéficiaire_SortX 2
        End Select
    Else
        Select Case fgSelect.Col
            Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_SortX 0
            Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
            Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_SortX 2
            Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_SortX 3
            Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
            Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_Sort
            Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_SortX 6
            Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
            Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
           Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
        End Select
    End If
Else
    If fgSelect.Rows > 1 Then
        fgSelect.Col = 5:  mCLIENARA1 = Trim(fgSelect.Text)
        fgSelect.Col = fgSelect_arrIndex:  arrYGUIMAD0_Index = CLng(fgSelect.Text)
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        xYGUIMAD0 = arrYGUIMAD0(arrYGUIMAD0_Index)
        oldYGUIMAD0 = xYGUIMAD0
        fraUpdate_Display
        If cmdSelect_SQL_K = "2" Or cmdSelect_SQL_K = "3" Or cmdSelect_SQL_K = "4" Then lstGUIMADLIEN_Load
   End If
End If
Me.Enabled = True: Me.MousePointer = 0
End Sub


Public Function fraUpdate_Control()
Dim blnUpdate_Control As Boolean
Dim X As String
blnUpdate_Control = True
Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données "): DoEvents
newYGUIMAD0 = oldYGUIMAD0
X = Trim(txtUpdate_GUIESPTI1)
If X = "" Then
    blnUpdate_Control = False
    txtUpdate_GUIESPTI1.BackColor = errUsr.BackColor
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le bénéficiaire")
Else
    txtUpdate_GUIESPTI1.BackColor = txtUsr.BackColor
End If
newYGUIMAD0.GUIESPTI1 = X
If Trim(xYGUIMAD0.GUIESPCL1) = "" Or xYGUIMAD0.GUIESPCL1 = "0010000" Then
    X = Trim(txtUpdate_GUIMADTDO)
    If X = "" Then
        blnUpdate_Control = False
        txtUpdate_GUIMADTDO.BackColor = errUsr.BackColor
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le donneur d'ordre")
    Else
        txtUpdate_GUIESPCL1.BackColor = txtUsr.BackColor
    End If
    newYGUIMAD0.GUIMADTDO = X
End If
If chkUpdate_GUIMADTIN = "1" Then
    newYGUIMAD0.GUIMADTIN = Trim(txtUpdate_GUIMADTIN)
Else
    newYGUIMAD0.GUIMADTIN = ""
End If

X = Trim(cboUpdate_GUIMADMOT)
If X = "" Then
    blnUpdate_Control = False
    cboUpdate_GUIMADMOT.BackColor = errUsr.BackColor
    Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le motif")
Else
    cboUpdate_GUIMADMOT.BackColor = txtUsr.BackColor
End If
newYGUIMAD0.GUIMADMOT = Mid$(X, 1, 4) & Trim(txtUpdate_GUIMADMOT)
        
txtUpdate_GUIMADMOT.BackColor = txtUsr.BackColor
If Mid$(newYGUIMAD0.GUIMADMOT, 1, 1) = "9" Then
    X = Trim(txtUpdate_GUIMADMOT)
    If X = "" Then
        blnUpdate_Control = False
        txtUpdate_GUIMADMOT.BackColor = errUsr.BackColor
        Call lstErr_AddItem(lstErr, cmdContext, "?_________Préciser le motif Divers")
    End If
End If

If newYGUIMAD0.GUIMADSTA = " " Then newYGUIMAD0.GUIMADSTA = "V"
If chkUpdate_GUIMADLIEN = "1" Then
    newYGUIMAD0.GUIMADLIEN = 0
Else
    newYGUIMAD0.GUIMADLIEN = CLng(txtUpdate_GUIMADLIEN)
End If

If blnUpdate_Control Then
    fraUpdate_Control = Null
Else
    fraUpdate_Control = "?_________fraUpdate_Control"
End If
End Function

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
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
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

Private Sub lstGUIMADLIEN_Click()
If lstGUIMADLIEN.ListIndex >= 0 Then
    txtUpdate_GUIMADLIEN = CLng(Left(lstGUIMADLIEN.Text, 6))
    fraGUIMADLIEN.Visible = False
End If
End Sub

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub

Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim meUnit As typeUnit, X As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), BIA_GUIMAD_Aut)

blnSetfocus = True
Form_Init
blnAuto = False

Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@AUTO_GUIMAD": blnAuto = True
                Call cbo_Scan("2 -", cboSelect_SQL)
                Call DTPicker_Set(txtSelect_AMJMIN, YBIATAB0_DATE_CPT_J)
                cmdSelect_Ok_Click
                Call cbo_Scan("9 -", cboSelect_SQL)
                Call DTPicker_Set(txtSelect_AMJMIN, YBIATAB0_DATE_CPT_J)
                cmdSelect_Ok_Click
                Call cbo_Scan("8 -", cboSelect_SQL)
                cmdSelect_Ok_Click
                Unload Me

    Case Else: blnAuto = False
End Select


End Sub


Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
    If fraGUIMADLIEN.Visible Then cmdGUIMADLIEN_Load_Click: Exit Sub

    If fraGUIMADLIEN.Visible Then fraGUIMADLIEN.Visible = False: Exit Sub
    If fraUpdate.Visible _
    And fraUpdate_B.Enabled _
    And cmdUpdate_Ok.Enabled Then cmdUpdate_Ok_Click: Exit Sub
Else
    If currentAction = "" Then
        If SSTab1.Tab > 0 Then
            SSTab1.Tab = 0
        Else
           'SendKeys "{TAB}"
           ' cmdSelect_Click
        End If
    End If
End If
End Sub









Private Sub mnuPrint0_All_Click()
Dim I As Long, K As Long
Me.Enabled = False: Me.MousePointer = vbHourglass
    
For I = 1 To arrYGUIMAD0_Nb
    fgSelect.Row = I
    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
    xYGUIMAD0 = arrYGUIMAD0(K)
    'prtSAB_CDR_Monitor xYGUIMAD0
Next I

Me.Show

Me.Enabled = True: Me.MousePointer = 0



End Sub




Public Sub cmdPrint_Ok()
Dim K As Long, X As String, xSQL As String
Dim wMOUVEMCOM As String
lstSelect.Visible = False


End Sub
Public Sub cmdPrint_Ok_1()
Dim K As Long, X As String
Dim wIndex As Integer

fgSelect.Visible = False
prtBIA_GUIMAD_Open 1, "Liste des mises à disposition", wMM, wAAAA
For K = 1 To fgSelect.Rows - 1
        fgSelect.Row = K
        fgSelect.Col = fgSelect_arrIndex
        wIndex = Val(fgSelect.Text)
        xYGUIMAD0 = arrYGUIMAD0(wIndex)
        
        prtBIA_GUIMAD_NewLine 1
        XPrt.CurrentX = prtMinX + 50: XPrt.Print dateIBM10(xYGUIMAD0.GUIESPDJO, True);
        
        XPrt.CurrentX = prtMinX + 1050: XPrt.Print xYGUIMAD0.GUIESPOPE;
        XPrt.CurrentX = prtMinX + 1350: XPrt.Print xYGUIMAD0.GUIESPNAT;
        X = Format$(xYGUIMAD0.GUIESPDOS, "### ##0")
        XPrt.CurrentX = prtMinX + 2050 - XPrt.TextWidth(X)
        XPrt.Print X;
         X = Format$(Abs(xYGUIMAD0.GUIESPMON), "### ### ### ###.00")
        XPrt.CurrentX = prtMinX + 3700 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.CurrentX = prtMinX + 3750: XPrt.Print xYGUIMAD0.GUIESPDEV;
        XPrt.CurrentX = prtMinX + 4200: XPrt.Print xYGUIMAD0.GUIESPCL1;
        XPrt.CurrentX = prtMinX + 5000
        If Trim(xYGUIMAD0.GUIESPCL1) = "" Or xYGUIMAD0.GUIESPCL1 = "0010000" Then
            XPrt.Print Trim(xYGUIMAD0.GUIMADTDO);
        Else
            X = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & xYGUIMAD0.GUIESPCL1 & "'"
            Set rsSab = cnsab.Execute(X)
            
            If Not rsSab.EOF Then XPrt.Print Trim(rsSab("CLIENARA1"));
        End If
        XPrt.CurrentX = prtMinX + 8000: XPrt.Print xYGUIMAD0.GUIESPTI1;
        
        X = Trim(xYGUIMAD0.GUIMADMOT)
        cbo_Scan Left$(X, 3), cboUpdate_GUIMADMOT
        XPrt.CurrentX = prtMinX + 11000: XPrt.Print cboUpdate_GUIMADMOT.Text;
       If Len(X) > 4 Then XPrt.CurrentX = prtMinX + 13000: XPrt.Print Right$(X, Len(X) - 4);

        X = Format$(xYGUIMAD0.GUIMADID, "### ##0")
        XPrt.CurrentX = prtMinX + 16000 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.CurrentX = prtMinX + 16050: XPrt.Print xYGUIMAD0.GUIMADSTA;
        
        If Trim(xYGUIMAD0.GUIMADTIN) <> "" Then
            prtBIA_GUIMAD_NewLine 1
            XPrt.FontItalic = True
            XPrt.CurrentX = prtMinX + 8000: XPrt.Print xYGUIMAD0.GUIMADTIN;
            XPrt.FontItalic = False
        End If
         
Next K
prtBIA_GUIMAD_Close 1

fgSelect.Visible = True


End Sub


Public Sub cmdPrint_Ok_7()
Dim K As Long, X As String
Dim wIndex As Integer

fgSelect.Visible = False
prtBIA_GUIMAD_Open 7, "MAD : Statistiques par bénéficiaire, période du " & dateImp(wAMJMin) & "-" & dateImp(WAMJMax), wMM, wAAAA
For K = 1 To fgSelect.Rows - 1
        fgSelect.Row = K
        fgSelect.Col = fgSelect_arrIndex
        wIndex = Val(fgSelect.Text)
        xYGUIMAD0 = arrYGUIMAD0(wIndex)
        
        prtBIA_GUIMAD_NewLine 1
        XPrt.CurrentX = prtMinX + 50: XPrt.Print xYGUIMAD0.GUIESPTI1;
        X = Format$(xYGUIMAD0.GUIESPDOS, "### ##0")
        XPrt.CurrentX = prtMinX + 6000 - XPrt.TextWidth(X)
        XPrt.Print X;
         X = Format$(Abs(xYGUIMAD0.GUIMADMON), "### ### ### ###.00")
        XPrt.CurrentX = prtMinX + 9500 - XPrt.TextWidth(X)
        XPrt.Print X;
Next K
prtBIA_GUIMAD_Close 7

fgSelect.Visible = True


End Sub

Public Sub cmdPrint_Ok_6()
Dim K As Long, X As String, K2 As Long
Dim wIndex As Integer
Dim blnRupture As Boolean
Dim xName As String, xMemo As String
rsYGUIMAD0_Init meYGUIMAD0
fgSelect.Visible = False

prtBIA_GUIMAD_Open 6, "Synthèse des mises à disposition", wMM, wAAAA
For wIndex = 1 To arrYGUIMAD0_Nb

    xYGUIMAD0 = arrYGUIMAD0(wIndex)
        prtBIA_GUIMAD_NewLine 6
        blnRupture = False
        If meYGUIMAD0.GUIESPDEV <> xYGUIMAD0.GUIESPDEV Then blnRupture = True
        If meYGUIMAD0.GUIESPCL1 <> xYGUIMAD0.GUIESPCL1 Then blnRupture = True
        
        If blnRupture Then
            XPrt.CurrentX = prtMinX + 50: XPrt.Print xYGUIMAD0.GUIESPDEV;
            XPrt.CurrentX = prtMinX + 400: XPrt.Print xYGUIMAD0.GUIESPCL1;
            XPrt.CurrentX = prtMinX + 1000
                X = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & xYGUIMAD0.GUIESPCL1 & "'"
                Set rsSab = cnsab.Execute(X)
                
                If Not rsSab.EOF Then XPrt.Print Mid$(rsSab("CLIENARA1"), 1, 30);
        End If
        
        If chkSelect_GUIMADMOT = "1" Then
            If blnRupture Then prtBIA_GUIMAD_NewLine 6
            Call rsElpTable_Read("GDMP", "MAD_Stat", Mid$(xYGUIMAD0.GUIMADMOT, 1, 1), xName, xMemo)
            If xName = "" Then xName = "???"
            XPrt.CurrentX = prtMinX + 1000: XPrt.Print xName;
        End If

    For K = 1 To 13
    
        If arrGUIESPMON(wIndex, K) <> 0 Then
            If K = 13 Then
                K2 = 13
            Else
                K2 = K + 1 - wMM
                'If K2 < 2 Then K2 = K2 + 12
            End If
            X = Trim(Format$(Abs(arrGUIESPMON(wIndex, K)), "### ### ### ###"))
            XPrt.CurrentX = prtMaxX - (13 - K2) * 900 - XPrt.TextWidth(X) - 50
            XPrt.Print X;
        End If
    Next K
    meYGUIMAD0 = xYGUIMAD0
Next wIndex


For wIndex = arrYGUIMAD0_Nb + 1 To arrGUIESPMON_Dev
    prtBIA_GUIMAD_NewLine 6
    XPrt.DrawWidth = 6
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
    xYGUIMAD0 = arrYGUIMAD0(wIndex)
    XPrt.CurrentX = prtMinX + 50: XPrt.Print xYGUIMAD0.GUIESPDEV;
    XPrt.CurrentY = XPrt.CurrentY + 100

    For K = 1 To 13
    
        If arrGUIESPMON(wIndex, K) <> 0 Then
            If K = 13 Then
                K2 = 13
            Else
                K2 = K + 1 - wMM
               ' If K2 < 2 Then K2 = K2 + 12
            End If
            X = Trim(Format$(Abs(arrGUIESPMON(wIndex, K)), "### ### ### ###"))
            XPrt.CurrentX = prtMaxX - (13 - K2) * 900 - XPrt.TextWidth(X) - 50
            XPrt.Print X;
        End If
    Next K
    
    prtBIA_GUIMAD_NewLine 6
    XPrt.DrawWidth = 2
    
    For K = 1 To 13
    
        If arrGUIESPMON(wIndex, K) <> 0 Then
            If K = 13 Then
                K2 = 13
            Else
                K2 = K + 1 - wMM
               ' If K2 < 2 Then K2 = K2 + 12
            End If
            X = Trim(Format$(Abs(arrGUIESPNB(wIndex, K)), "### ### ### ###"))
            XPrt.CurrentX = prtMaxX - (13 - K2) * 900 - XPrt.TextWidth(X) - 50
            XPrt.Print X;
        End If

    Next K

Next wIndex

XPrt.DrawWidth = 5
prtBIA_GUIMAD_Close 6

fgSelect.Visible = True


End Sub

Public Function YGUIMAD0_Add_ZGUIESP0()
Dim V, X As String, xSQL As String
Dim Nb As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdUpdate_Add_New"
'-------------------------------------------------------
'trace debug
Dim fTrace As Integer
fTrace = FreeFile
Call FEU_ROUGE
Open "c:\temp\YGUIMAD0_Add_ZGUIESP0.txt" For Output As fTrace

YGUIMAD0_Add_ZGUIESP0 = Null
mMsgBox = xYGUIMAD0.GUIESPOPE & " " & xYGUIMAD0.GUIESPNAT & " " & xYGUIMAD0.GUIESPDOS
Print #fTrace, mMsgBox & " " & Format(Now, "dd/mm/yyyy hh:nn:ss")

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
Print #fTrace, "Après cnSAB_Transaction('BeginTrans') " & Format(Now, "dd/mm/yyyy hh:nn:ss")

If Not IsNull(V) Then GoTo Error_MsgBox

Call lstErr_AddItem(lstErr, cmdContext, "Ajout MAD : " & mMsgBox): DoEvents
'________________________________________________________________________________
rsYGUIMAD0_Init meYGUIMAD0
Print #fTrace, "Après rsYGUIMAD0_Init meYGUIMAD0 " & Format(Now, "dd/mm/yyyy hh:nn:ss")

V = sqlYGUIMAD0_Init(meYGUIMAD0)
Print #fTrace, "Après sqlYGUIMAD0_Init(meYGUIMAD0) " & Format(Now, "dd/mm/yyyy hh:nn:ss")

If Not IsNull(V) Then GoTo Error_MsgBox
meYGUIMAD0.GUIESPOPE = rsSab("GUIESPOPE")
meYGUIMAD0.GUIESPDOS = rsSab("GUIESPDOS")
meYGUIMAD0.GUIESPNAT = rsSab("GUIESPNAT")
meYGUIMAD0.GUIESPMON = rsSab("GUIESPMON")
meYGUIMAD0.GUIESPDEV = rsSab("GUIESPDEV")
meYGUIMAD0.GUIESPCP1 = rsSab("GUIESPCP1")
meYGUIMAD0.GUIESPCL1 = rsSab("GUIESPCL1")
X = UCase$(rsSab("GUIESPTI1"))
X = Replace(X, "MAD F/", "")
X = Replace(X, "MAD FAV/", "")
X = Replace(X, "MAD ", "")
meYGUIMAD0.GUIESPTI1 = X
meYGUIMAD0.GUIESPDJO = rsSab("GUIESPDJO")
If rsSab("GUIESPCTA") = "4" Then meYGUIMAD0.GUIMADSTA = "A"
If meYGUIMAD0.GUIESPDEV = "EUR" Then
    meYGUIMAD0.GUIMADMON = meYGUIMAD0.GUIESPMON
Else
    meCV1.DeviseIso = meYGUIMAD0.GUIESPDEV
    meCV1.Montant = meYGUIMAD0.GUIESPMON
    meCV1.OpéAmj = meYGUIMAD0.GUIESPDJO + 19000000
    Call CV_Calc("J", meCV1, meCV2)
    Print #fTrace, "Après Call CV_Calc('J', meCV1, meCV2) " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    
    meYGUIMAD0.GUIMADMON = meCV2.Montant

End If

V = sqlYGUIMAD0_Insert(meYGUIMAD0)
Print #fTrace, "Après sqlYGUIMAD0_Insert(meYGUIMAD0) " & Format(Now, "dd/mm/yyyy hh:nn:ss")

If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
        Print #fTrace, "Après cnSAB_Transaction('Rollback') " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    Else
        V = cnSAB_Transaction("Commit")
        Print #fTrace, "Après cnSAB_Transaction('Commit') " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    End If
    
    YGUIMAD0_Add_ZGUIESP0 = V
    Print #fTrace, "Après YGUIMAD0_Add_ZGUIESP0 = V " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    
    '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    Close #fTrace
    Call FEU_VERT

End Function

Public Function YGUIMAD0_Add_ZCHGOPE0()
Dim V, X As String, xSQL As String, xWhere As String
Dim Nb As Long
Dim mMsgBox As String
Dim wDev As String, curDev As Currency, curEur As Currency
Dim wGUIESPTI1 As String

On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "YGUIMAD0_Add_ZCHGOPE0"
'-------------------------------------------------------

YGUIMAD0_Add_ZCHGOPE0 = Null
mMsgBox = xYGUIMAD0.GUIESPOPE & " " & xYGUIMAD0.GUIESPNAT & " " & xYGUIMAD0.GUIESPDOS

xWhere = " where CHGMESDOS = " & rsSab("CHGOPEDOS") _
       & " and CHGMESOPE = '" & rsSab("CHGOPEOPE") & "'" _
       & " and CHGMESDON = '" & rsSab("CHGOPECON") & "'" _
       & " and CHGMESSER = '" & rsSab("CHGOPESER") & "' and CHGMESSSE = '" & rsSab("CHGOPESSE") & "'"
            
xSQL = "select CHGMESMO1 from " & paramIBM_Library_SAB & ".ZCHGMES0 " & xWhere
Set rsSabX = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    X = UCase$(rsSabX("CHGMESMO1"))
Else
    X = ""
End If
If InStr(X, "MAD") = 0 Then Exit Function
'---------------------------------------------------- uniquement si MAD
X = Replace(X, "MAD F/", "")
X = Replace(X, "MAD FAV/", "")
X = Replace(X, "MAD ", "")
wGUIESPTI1 = X


'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

Call lstErr_AddItem(lstErr, cmdContext, "Ajout MAD : " & mMsgBox): DoEvents
'________________________________________________________________________________
rsYGUIMAD0_Init meYGUIMAD0
V = sqlYGUIMAD0_Init(meYGUIMAD0)
If Not IsNull(V) Then GoTo Error_MsgBox
meYGUIMAD0.GUIESPOPE = rsSab("CHGOPEOPE")
meYGUIMAD0.GUIESPDOS = rsSab("CHGOPEDOS")
meYGUIMAD0.GUIESPNAT = rsSab("CHGOPENAT")

wDev = rsSab("CHGOPEDE1")
If wDev = "EUR" Then
    curEur = rsSab("CHGOPEMO1")
    curDev = rsSab("CHGOPEMO2")
    If curDev = 0 Then
        curDev = curEur
    Else
        wDev = rsSab("CHGOPEDE2")
    End If
Else
    curEur = rsSab("CHGOPEMO2")
    curDev = rsSab("CHGOPEMO1")

End If

meYGUIMAD0.GUIESPDEV = wDev
meYGUIMAD0.GUIESPMON = curDev
meYGUIMAD0.GUIMADMON = curEur
meYGUIMAD0.GUIESPCL1 = rsSab("CHGOPECON")
meYGUIMAD0.GUIESPDJO = rsSab("CHGOPECRE")

Set rsSabX = Nothing
xWhere = " where CHGDETDOS = " & meYGUIMAD0.GUIESPDOS _
       & " and CHGDETOPE = '" & meYGUIMAD0.GUIESPOPE & "'" _
       & " and CHGDETCL1 = '" & meYGUIMAD0.GUIESPCL1 & "'" _
       & " and CHGDETSER = '" & rsSab("CHGOPESER") & "' and CHGDETSSE = '" & rsSab("CHGOPESSE") & "'"
            
xSQL = "select CHGDETCP1 from " & paramIBM_Library_SAB & ".ZCHGDET0 " & xWhere
Set rsSabX = cnsab.Execute(xSQL)

If Not rsSabX.EOF Then meYGUIMAD0.GUIESPCP1 = rsSabX("CHGDETCP1")

'------------------------------------------
meYGUIMAD0.GUIESPTI1 = wGUIESPTI1
'------------------------------------------
V = sqlYGUIMAD0_Insert(meYGUIMAD0)
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    YGUIMAD0_Add_ZCHGOPE0 = V
    '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function

Public Function cmdUpdate_Ok_Transaction()
Dim V, X As String, xSQL As String
Dim Nb As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdUpdate_Ok_Transaction"
'-------------------------------------------------------
cmdUpdate_Ok_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
V = sqlYGUIMAD0_Update(newYGUIMAD0, oldYGUIMAD0)
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    cmdUpdate_Ok_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function


Public Sub lstSelect_Load_2()
cmdSelect_Ok_Caption = "Importer les MAD de SAB=> GUIMAD "
'cmdSelect_Ok.BackColor = &HC0FFC0
txtSelect_AMJMIN.Visible = True
txtSelect_AMJMIN.Enabled = True
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
End Sub

Public Sub lstSelect_Load_9()
cmdSelect_Ok_Caption = "courriels aux gestionnaires "
txtSelect_AMJMIN.Visible = True
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
End Sub

Public Sub lstSelect_Load_8()
cmdSelect_Ok_Caption = "Alertes MAD incomplètes ou annulées "
txtSelect_AMJMIN.Visible = False
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
End Sub

Public Sub lstSelect_Load_6()
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_1"
cmdSelect_Ok_Caption = "Lancer la requête"
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True
chkSelect_GUIESPDJO = "1"
    txtSelect_GUIESPDJO.Visible = False
    txtSelect_GUIESPDJO_Max.Visible = True

chkSelect_GUIMADMOT.Visible = True
End Sub

Public Sub fraUpdate_Display()
Dim X As String

Call lstErr_Clear(lstErr, cmdContext, ">Afficahge du détail dossier"): DoEvents
lblUpdate_GUIESPDOS = "MAD n° " & xYGUIMAD0.GUIMADID
lblUpdate_GUIESPDOS.ForeColor = vbMagenta
'fraUpdate.ForeColor = vbMagenta
txtUpdate_GUIESPOPE = xYGUIMAD0.GUIESPOPE
txtUpdate_GUIESPNAT = xYGUIMAD0.GUIESPNAT
txtUpdate_GUIESPDOS = xYGUIMAD0.GUIESPDOS
txtUpdate_GUIESPMON = Format(xYGUIMAD0.GUIESPMON, "### ### ### ##0.00")
txtUpdate_GUIESPDEV = xYGUIMAD0.GUIESPDEV
txtUpdate_GUIESPCP1 = xYGUIMAD0.GUIESPCP1
txtUpdate_GUIESPCL1 = xYGUIMAD0.GUIESPCL1
txtUpdate_GUIESPTI1 = Trim(xYGUIMAD0.GUIESPTI1)
txtUpdate_GUIESPTI1.BackColor = txtUsr.BackColor

Call DTPicker_Set(txtUpdate_GUIESPDJO, xYGUIMAD0.GUIESPDJO + 19000000)

txtUpdate_GUIMADTDO = Trim(xYGUIMAD0.GUIMADTDO)
txtUpdate_GUIMADTDO.BackColor = txtUsr.BackColor
X = Trim(xYGUIMAD0.GUIMADTIN)
If X <> "" Then
    chkUpdate_GUIMADTIN = "1"
    txtUpdate_GUIMADTIN.Visible = True
Else
    chkUpdate_GUIMADTIN = "0"
    txtUpdate_GUIMADTIN.Visible = False
End If
txtUpdate_GUIMADTIN = X
txtUpdate_GUIMADTIN.BackColor = txtUsr.BackColor

txtUpdate_GUIMADMOT = ""
X = Trim(xYGUIMAD0.GUIMADMOT)
If X = "" Then
    cboUpdate_GUIMADMOT.ListIndex = 0
Else
    cbo_Scan Left$(X, 3), cboUpdate_GUIMADMOT
    cboUpdate_GUIMADMOT.BackColor = txtUsr.BackColor
    If Len(X) > 4 Then txtUpdate_GUIMADMOT = Right$(X, Len(X) - 4)
End If

txtUpdate_GUIMADLIEN = xYGUIMAD0.GUIMADLIEN

fraUpdate.Visible = True
fraUpdate_A.Enabled = False
fraUpdate_B.Enabled = BIA_GUIMAD_Aut.Saisir
If Trim(xYGUIMAD0.GUIESPCL1) = "" Or xYGUIMAD0.GUIESPCL1 = "0010000" Then
    libUpdate_GUIESPCL1 = ""
   txtUpdate_GUIMADTDO.Enabled = True
Else
    libUpdate_GUIESPCL1 = mCLIENARA1
   txtUpdate_GUIMADTDO.Enabled = False
End If

Select Case xYGUIMAD0.GUIMADSTA
    Case " ": txtUpdate_GUIMADSTA = "Import": cmdUpdate_Ok.Enabled = BIA_GUIMAD_Aut.Saisir
    Case "V": txtUpdate_GUIMADSTA = "Validé": cmdUpdate_Ok.Enabled = BIA_GUIMAD_Aut.Valider
    Case "A": txtUpdate_GUIMADSTA = "Annulé": cmdUpdate_Ok.Enabled = BIA_GUIMAD_Aut.Valider
    Case Else: txtUpdate_GUIMADSTA = xYGUIMAD0.GUIMADSTA
End Select

chkUpdate_GUIMADLIEN = "0"
If xYGUIMAD0.GUIMADLIEN = 0 Then
    chkUpdate_GUIMADLIEN.Enabled = False
Else
    chkUpdate_GUIMADLIEN.Enabled = BIA_GUIMAD_Aut.Valider
End If
libUpdate_GUIMADUSR = Trim(xYGUIMAD0.GUIMADUSR)

cmdUpdate_Annuler.Enabled = BIA_GUIMAD_Aut.Valider
End Sub

Private Sub txtGUIMADLIEN_Scan_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSelect_GUIESPCL1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSelect_GUIESPTI1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub cboUpdate_GUIMADMOT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelect_GUIMADTIN_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtUpdate_GUIMADTDO_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtUpdate_GUIMADTIN_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub



Public Sub cmdSelect_SQL_9_Mail()
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim xDétail As String, xSubject As String
Dim wNb As Long, xNb As String
Dim mCLIENACLI As String, mCLIENARES As String, mCLIENARA1 As String
Dim xCLIENARES As String
Dim iCol As Integer, K As Integer, X As String
Dim xGUIMADMOT As String

mCLIENARES = ""

For K = 0 To lstW.ListCount - 1
    lstW.ListIndex = K
    X = lstW.Text
    iCol = 0
    xCLIENARES = CSV_Scan(X, iCol)
    Call CSV_Scan(X, iCol)
    arrYGUIMAD0_Index = CSV_Scan(X, iCol)
    mCLIENARA1 = CSV_Scan(X, iCol)
    
    
    If mCLIENARES <> xCLIENARES Then
        If mCLIENARES <> "" Then cmdSelect_SQL_9_Mail_Send mCLIENARES, xDétail, xSubject
        mCLIENARES = xCLIENARES
        Call rsYBIATAB0_Read("RESPONSABLE", mCLIENARES, "", X)
        xSubject = "MAD du " & dateImp10(wAMJMin) & " - Gestionnaire " & mCLIENARES & " : " & Trim(Mid$(X, 46, 30))
        xDétail = "<BR><U></CENTER><B>" & htmlFontColor("MAGENTA") & xSubject & "</B></U><BR><BR>" & "<TABLE width= 90% border=1>"
    End If
    xYGUIMAD0 = arrYGUIMAD0(arrYGUIMAD0_Index)
    
    cbo_Scan Left$(xYGUIMAD0.GUIMADMOT, 3), cboUpdate_GUIMADMOT
    xGUIMADMOT = Trim(xYGUIMAD0.GUIMADMOT)
    If Len(xGUIMADMOT) > 4 Then
        xGUIMADMOT = cboUpdate_GUIMADMOT.Text & " " & Right$(xGUIMADMOT, Len(xGUIMADMOT) - 4)
    Else
        xGUIMADMOT = cboUpdate_GUIMADMOT.Text
    End If
    
    xDétail = xDétail & "<TR><TD width = 10%>" & htmlFontColor("BLUE") & "<FONT SIZE=-1>" & xYGUIMAD0.GUIESPCP1 & "</FONT SIZE></TD>" _
                      & "<TD width = 30%><FONT SIZE=-1>" & htmlFontColor("BLUE") & mCLIENARA1 & "</FONT></TD>" _
                      & "<TD width = 20%><B>" & htmlFontColor("RED") & Format$(xYGUIMAD0.GUIESPMON, "### ### ### ##0.00") & xYGUIMAD0.GUIESPDEV & "</B></TD>" _
                      & "<TD width = 20%><FONT SIZE=-1>" & htmlFontColor("MAGENTA") & xYGUIMAD0.GUIESPTI1 & "</TD>" _
                      & "<TD width = 20%><FONT SIZE=-1>" & htmlFontColor("MAGENTA") & xGUIMADMOT & "</FONT></TD></TR>"
Next K

If mCLIENARES <> "" Then cmdSelect_SQL_9_Mail_Send mCLIENARES, xDétail, xSubject

 

End Sub
Public Sub cmdSelect_SQL_8_Mail()
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim xDétail As String, xSubject As String
Dim wNb As Long, xNb As String
Dim iCol As Integer, K As Integer, X As String
Dim mCLIENACLI As String, mCLIENARA1 As String
Dim xSQL As String

mCLIENACLI = ""
xSubject = "MAD à compléter"
xDétail = "<BR><U></CENTER><B>" & htmlFontColor("BLUE") & xSubject & "</B></U><BR><BR>" & "<TABLE width= 90% border=1>"
For K = 1 To arrYGUIMAD0_Nb

    xYGUIMAD0 = arrYGUIMAD0(K)
    
    If mCLIENACLI <> xYGUIMAD0.GUIESPCL1 Then
        mCLIENACLI = xYGUIMAD0.GUIESPCL1
        xSQL = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & mCLIENACLI & "'"
        Set rsSabX = cnsab.Execute(xSQL)
        
        If Not rsSabX.EOF Then
            mCLIENARA1 = Trim(rsSabX("CLIENARA1"))
        Else
            mCLIENARA1 = "???"
        End If
    End If

    xDétail = xDétail & "<TR><TD width = 15%>" & htmlFontColor("BLUE") & "<FONT SIZE=-1>" & xYGUIMAD0.GUIESPCP1 & "</FONT SIZE></TD>" _
                      & "<TD width = 35%><FONT SIZE=-1>" & htmlFontColor("BLUE") & mCLIENARA1 & "</FONT></TD>" _
                      & "<TD width = 20%><B>" & htmlFontColor("BLACK") & Format$(xYGUIMAD0.GUIESPMON, "### ### ### ##0.00") & xYGUIMAD0.GUIESPDEV & "</B></TD>" _
                      & "<TD width = 30%><FONT SIZE=-1>" & htmlFontColor("BLACK") & xYGUIMAD0.GUIESPTI1 & "</FONT></TD></TR>"

Next K

cmdSelect_SQL_8_Mail_Send xDétail, xSubject

End Sub

Public Sub cmdSelect_SQL_9_Mail_Send(lCLIENARES As String, lDétail As String, lSubject As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim iCol As Integer, K As Integer

wSendMail.FromDisplayName = "MAD_" & lCLIENARES
wSendMail.RecipientDisplayName = "GUIMAD"

bgColor = "CYAN"
wSendMail.Subject = lSubject
wSendMail.Attachment = ""
wSendMail.Message = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
                    & "<FONT face=" & Asc34 & prtFontName_Arial & Asc34 & ">" _
                    & htmlFontColor("BLUE") & lDétail

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

 

End Sub



Public Sub cmdSelect_SQL_8_Mail_Send(lDétail As String, lSubject As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim iCol As Integer, K As Integer

wSendMail.FromDisplayName = "ALERTE"
wSendMail.RecipientDisplayName = "GUIMAD"

bgColor = "MAGENTA"
wSendMail.Subject = lSubject
wSendMail.Attachment = ""
wSendMail.Message = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
                    & "<FONT face=" & Asc34 & prtFontName_Arial & Asc34 & ">" _
                    & htmlFontColor("BLUE") & lDétail

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

 

End Sub

Public Function cmdSelect_SQL_6_Cumul_Dev(lDev As String)
Dim K As Integer
For K = arrYGUIMAD0_Nb + 3 To arrGUIESPMON_Dev

    If arrYGUIMAD0(K).GUIESPDEV = lDev Then cmdSelect_SQL_6_Cumul_Dev = K: Exit Function
Next K
arrGUIESPMON_Dev = arrGUIESPMON_Dev + 1
arrYGUIMAD0(arrGUIESPMON_Dev).GUIESPDEV = lDev
arrYGUIMAD0(arrGUIESPMON_Dev).GUIESPCL1 = ""
cmdSelect_SQL_6_Cumul_Dev = arrGUIESPMON_Dev
End Function

Public Sub lstGUIMADLIEN_Load()
Dim X As String, K As Integer, blnOk As Boolean
Dim xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass
K = 0
txtGUIMADLIEN_Scan = ""
blnOk = False
Do
    X = Space_Scan(xYGUIMAD0.GUIESPTI1, K)
    If X = "" Then
        blnOk = True
    Else
        If Len(X) > 2 Then txtGUIMADLIEN_Scan = txtGUIMADLIEN_Scan & Left$(X, 3) & " "
    End If
Loop Until blnOk
cmdGUIMADLIEN_Load_Click

lstGUIMADLIEN_Display.Clear
xSQL = "select GUIMADID,GUIESPTI1 from " & paramIBM_Library_SABSPE & ".YGUIMAD0 where GUIMADLIEN = " & xYGUIMAD0.GUIMADID
Set rsSabX = cnsab.Execute(xSQL)

Do While Not rsSabX.EOF
    lstGUIMADLIEN_Display.AddItem Format$(rsSabX("GUIMADID"), "000000") & " " & rsSabX("GUIESPTI1")
    rsSabX.MoveNext

Loop

If lstGUIMADLIEN_Display.ListCount > 0 Then
    lstGUIMADLIEN.Enabled = False
Else
    lstGUIMADLIEN.Enabled = cmdUpdate_Ok.Enabled
End If
fraGUIMADLIEN.Visible = True
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdSelect_SQL_M()
Dim V, X As String, wGUIESPDOS As Long, wGUIESPTI1 As String

X = InputBox("Indiquer le n° de dossier: " _
    & "ZGUIESP0 (1,1,00,GU,RE) ")
wGUIESPDOS = Val(Trim(X))
If wGUIESPDOS = 0 Then Exit Sub

X = "select * from " & paramIBM_Library_SAB & ".ZGUIESP0 " _
   & " where GUIESPETA = 1 and GUIESPAGE = 1 and GUIESPSER = '00' and GUIESPSSE = 'GU' and  GUIESPOPE = 'RE '" _
   & " and GUIESPDOS = " & wGUIESPDOS
Set rsSab = cnsab.Execute(X)

If rsSab.EOF Then
    Call MsgBox("Dossier inconnu", vbCritical, "BIA_GUIMAD : modification Texte")
    Exit Sub
End If
wGUIESPTI1 = Trim(rsSab("GUIESPTI1"))
wGUIESPTI1 = InputBox("Modifier le libellé : " _
    & vbCrLf & "     =========================", "BIA_GUIMAD : modification Texte", wGUIESPTI1)

wGUIESPTI1 = Trim(wGUIESPTI1)
If wGUIESPTI1 = "" Then
    Call MsgBox("modification abandonnée", vbCritical, "BIA_GUIMAD : modification Texte")
    Exit Sub
End If
If Len(wGUIESPTI1) > 30 Then wGUIESPTI1 = Mid$(wGUIESPTI1, 1, 30)
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
X = "Update " & paramIBM_Library_SAB & ".ZGUIESP0 " _
   & " set GUIESPTI1 = '" & wGUIESPTI1 & "'" _
   & " where GUIESPETA = 1 and GUIESPAGE = 1 and GUIESPSER = '00' and GUIESPSSE = 'GU' and  GUIESPOPE = 'RE '" _
   & " and GUIESPDOS = " & wGUIESPDOS
Call FEU_ROUGE
Set rsSab = cnsab.Execute(X)
Call FEU_VERT

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, "BIA_GUIMAD : modification Texte"
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Sub

Public Sub cmdSelect_JPL()
Dim xSQL As String
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YGUIMAD0 where GUIESPDEV <> 'EUR' and GUIMADMON > 999999"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call rsYGUIMAD0_GetBuffer(rsSab, oldYGUIMAD0)
    
    meCV1.DeviseIso = oldYGUIMAD0.GUIESPDEV
    meCV1.Montant = oldYGUIMAD0.GUIESPMON
    meCV1.OpéAmj = oldYGUIMAD0.GUIESPDJO + 19000000
    Call CV_Calc("J", meCV1, meCV2)
    newYGUIMAD0 = oldYGUIMAD0
    newYGUIMAD0.GUIMADMON = meCV2.Montant
    Debug.Print oldYGUIMAD0.GUIMADID; oldYGUIMAD0.GUIESPMON; newYGUIMAD0.GUIMADMON; oldYGUIMAD0.GUIMADMON
    cmdUpdate_Ok_Transaction
    rsSab.MoveNext

Loop

End Sub
