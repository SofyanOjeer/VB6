VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmYGOSDOS0_Annexe 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E6E6E6&
   Caption         =   "Gestion des Opérations en Suspens ANNEXE"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "YGOSDOS0_Annexe.frx":0000
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
      Height          =   270
      Left            =   6120
      TabIndex        =   4
      Top             =   0
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9855
      Left            =   0
      TabIndex        =   2
      Top             =   465
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   17383
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Gestion des Opérations en Suspens"
      TabPicture(0)   =   "YGOSDOS0_Annexe.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Paramétrage"
      TabPicture(1)   =   "YGOSDOS0_Annexe.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "."
      TabPicture(2)   =   "YGOSDOS0_Annexe.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   9285
         Left            =   -75030
         TabIndex        =   9
         Top             =   390
         Width           =   13425
         _ExtentX        =   23680
         _ExtentY        =   16378
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "fraSelect"
         TabPicture(0)   =   "YGOSDOS0_Annexe.frx":035E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fgFree"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtFg"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lstW"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "fraSwift"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "fraSelect_Options_J"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "-"
         TabPicture(1)   =   "YGOSDOS0_Annexe.frx":037A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraDetail"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "-"
         TabPicture(2)   =   "YGOSDOS0_Annexe.frx":0396
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.Frame fraDetail 
            BackColor       =   &H00D8DFD8&
            Height          =   8025
            Left            =   -74265
            TabIndex        =   35
            Top             =   645
            Width           =   12600
            Begin TabDlg.SSTab tabDetail 
               Height          =   7575
               Left            =   165
               TabIndex        =   36
               Top             =   210
               Width           =   12240
               _ExtentX        =   21590
               _ExtentY        =   13361
               _Version        =   393216
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   520
               ForeColor       =   192
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Dossier"
               TabPicture(0)   =   "YGOSDOS0_Annexe.frx":03B2
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "libDetail_SWISABSWID"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "fgDetail"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "fraList"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).ControlCount=   3
               TabCaption(1)   =   "Evénements"
               TabPicture(1)   =   "YGOSDOS0_Annexe.frx":03CE
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "fraEVE_C"
               Tab(1).ControlCount=   1
               Begin VB.Frame fraList 
                  BackColor       =   &H00D8DFD8&
                  Height          =   7155
                  Left            =   6465
                  TabIndex        =   38
                  Top             =   350
                  Visible         =   0   'False
                  Width           =   5940
               End
               Begin VB.Frame fraEVE_C 
                  Height          =   7455
                  Left            =   -74910
                  TabIndex        =   37
                  Top             =   330
                  Width           =   12015
               End
               Begin MSFlexGridLib.MSFlexGrid fgDetail 
                  Height          =   6795
                  Left            =   120
                  TabIndex        =   39
                  Top             =   780
                  Visible         =   0   'False
                  Width           =   6195
                  _ExtentX        =   10927
                  _ExtentY        =   11986
                  _Version        =   393216
                  Cols            =   3
                  FixedCols       =   0
                  RowHeightMin    =   350
                  BackColor       =   15794175
                  ForeColor       =   8192
                  BackColorFixed  =   8421376
                  ForeColorFixed  =   16777215
                  BackColorBkg    =   15794175
                  AllowUserResizing=   3
                  FormatString    =   "<Code |<Valeur                                                                                                      |"
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
               Begin VB.Label libDetail_SWISABSWID 
                  BackColor       =   &H00E0FFFF&
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   90
                  TabIndex        =   40
                  Top             =   330
                  Width           =   5985
               End
            End
         End
         Begin VB.Frame fraSelect_Options_J 
            BackColor       =   &H00F0FFFF&
            BorderStyle     =   0  'None
            Height          =   1305
            Left            =   210
            TabIndex        =   18
            Top             =   1995
            Visible         =   0   'False
            Width           =   9375
            Begin VB.TextBox txtSelect_J_merged_text 
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
               TabIndex        =   41
               Top             =   870
               Width           =   1260
            End
            Begin VB.ComboBox cboSelect_J_event_name_Top 
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
               Left            =   3795
               Sorted          =   -1  'True
               TabIndex        =   25
               Top             =   90
               Visible         =   0   'False
               Width           =   3660
            End
            Begin VB.ComboBox cboSelect_J_appl_serv_name 
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
               Left            =   1300
               Sorted          =   -1  'True
               TabIndex        =   24
               Top             =   925
               Width           =   1290
            End
            Begin VB.ComboBox cboSelect_J_event_severity 
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
               Left            =   1300
               Sorted          =   -1  'True
               TabIndex        =   23
               Top             =   500
               Width           =   1320
            End
            Begin VB.ComboBox cboSelect_J_event_class 
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
               Left            =   3615
               Sorted          =   -1  'True
               TabIndex        =   22
               Top             =   555
               Width           =   1710
            End
            Begin VB.ComboBox cboSelect_J_event_name 
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
               Left            =   3345
               Sorted          =   -1  'True
               TabIndex        =   21
               Top             =   90
               Width           =   3660
            End
            Begin VB.ComboBox cboSelect_J_comp_name 
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
               Left            =   1300
               Sorted          =   -1  'True
               TabIndex        =   20
               Top             =   75
               Width           =   840
            End
            Begin VB.TextBox txtSelect_J_oper_nickname 
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
               Left            =   4020
               TabIndex        =   19
               Top             =   885
               Width           =   1260
            End
            Begin MSComCtl2.DTPicker txtSelect_J_AMJMin 
               Height          =   300
               Left            =   7100
               TabIndex        =   26
               Top             =   500
               Width           =   2145
               _ExtentX        =   3784
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
               CustomFormat    =   "dd/MM/yyyy   HH:mm:ss"
               Format          =   111673347
               CurrentDate     =   40898.9999884259
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_J_AMJMax 
               Height          =   300
               Left            =   7100
               TabIndex        =   27
               Top             =   925
               Width           =   2145
               _ExtentX        =   3784
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
               CustomFormat    =   "dd/MM/yyyy   HH:mm:ss"
               Format          =   111869955
               CurrentDate     =   40898.9999884259
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_J_merged_text 
               BackColor       =   &H00F0FFFF&
               Caption         =   "texte"
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
               Left            =   5865
               TabIndex        =   42
               Top             =   525
               Width           =   690
            End
            Begin VB.Label lblSelect_J_appl_serv_name 
               BackColor       =   &H00F0FFFF&
               Caption         =   "application"
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
               Left            =   135
               TabIndex        =   34
               Top             =   925
               Width           =   1005
            End
            Begin VB.Label lblSelect_J_event_class 
               BackColor       =   &H00F0FFFF&
               Caption         =   "classe "
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
               Left            =   2835
               TabIndex        =   33
               Top             =   600
               Width           =   690
            End
            Begin VB.Label lblSelect_J_event_name 
               BackColor       =   &H00F0FFFF&
               Caption         =   "événement"
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
               Left            =   2250
               TabIndex        =   32
               Top             =   135
               Width           =   1005
            End
            Begin VB.Label lblSelect_J_event_severity 
               BackColor       =   &H00F0FFFF&
               Caption         =   "severity"
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
               TabIndex        =   31
               Top             =   500
               Width           =   1005
            End
            Begin VB.Label lblSelect_J_comp_name 
               BackColor       =   &H00F0FFFF&
               Caption         =   "composant"
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
               TabIndex        =   30
               Top             =   75
               Width           =   1000
            End
            Begin VB.Label lblSelect_J_AMJ 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               Caption         =   "période"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7100
               TabIndex        =   29
               Top             =   135
               Width           =   2025
            End
            Begin VB.Label lblSelect_J_oper_nickname 
               BackColor       =   &H00F0FFFF&
               Caption         =   "opérateur(%)"
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
               Left            =   2730
               TabIndex        =   28
               Top             =   915
               Width           =   1230
            End
         End
         Begin VB.Frame fraSwift 
            BackColor       =   &H00C0E0FF&
            Height          =   7320
            Left            =   7275
            TabIndex        =   12
            Top             =   1590
            Visible         =   0   'False
            Width           =   6200
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
               TabIndex        =   14
               Top             =   930
               Width           =   6050
            End
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
               TabIndex        =   13
               Top             =   600
               Width           =   6050
            End
            Begin MSFlexGridLib.MSFlexGrid fgSwift 
               Height          =   6000
               Left            =   60
               TabIndex        =   15
               Top             =   1260
               Width           =   6050
               _ExtentX        =   10663
               _ExtentY        =   10583
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
               TabIndex        =   16
               Top             =   210
               Width           =   6050
            End
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
            Height          =   2040
            Left            =   9480
            TabIndex        =   11
            Top             =   345
            Visible         =   0   'False
            Width           =   4212
         End
         Begin VB.TextBox txtFg 
            Height          =   1260
            Left            =   90
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   570
            Visible         =   0   'False
            Width           =   6732
         End
         Begin MSFlexGridLib.MSFlexGrid fgFree 
            Height          =   3465
            Left            =   2475
            TabIndex        =   17
            Top             =   5760
            Visible         =   0   'False
            Width           =   3570
            _ExtentX        =   6297
            _ExtentY        =   6112
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            BackColorFixed  =   8421504
            ForeColorFixed  =   16777215
            BackColorBkg    =   14737632
            WordWrap        =   -1  'True
            FormatString    =   $"YGOSDOS0_Annexe.frx":03EA
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
      Begin VB.Frame fraTab0 
         Height          =   9420
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   13296
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
            Left            =   9510
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   300
            Width           =   3705
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   10485
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   780
            Width           =   1335
         End
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00F0FFFF&
            Height          =   1305
            Left            =   90
            TabIndex        =   6
            Top             =   135
            Visible         =   0   'False
            Width           =   9375
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7890
            Left            =   100
            TabIndex        =   5
            Top             =   1470
            Visible         =   0   'False
            Width           =   13080
            _ExtentX        =   23072
            _ExtentY        =   13917
            _Version        =   393216
            Rows            =   1
            Cols            =   14
            FixedCols       =   0
            RowHeightMin    =   450
            BackColor       =   16777215
            ForeColor       =   12582912
            BackColorFixed  =   8421504
            ForeColorFixed  =   16777215
            BackColorSel    =   12648384
            BackColorBkg    =   15794175
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"YGOSDOS0_Annexe.frx":04AB
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
      Height          =   500
      Left            =   13080
      Picture         =   "YGOSDOS0_Annexe.frx":05BA
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
      Begin VB.Menu mnuPrint_Recap 
         Caption         =   "Etat récapitulatif"
      End
      Begin VB.Menu mnuPrint_Detail 
         Caption         =   "Etat détaillé"
      End
   End
   Begin VB.Menu mnuZSWIBIC0 
      Caption         =   "BIC"
      Visible         =   0   'False
      Begin VB.Menu mnuSWIBICBIC 
         Caption         =   "BIC"
      End
   End
End
Attribute VB_Name = "frmYGOSDOS0_Annexe"
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
Dim intReturn As Integer, wFile As String
'JPL HAB Dim YGOSDOS0_Aut As typeAuthorization
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


Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean
Dim fgDetail_50 As String, fgDetail_59 As String, fgDetail_57 As String
Dim fgDetail_70 As String, fgDetail_72 As String

'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long


Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim mXls1_Row As Long, mXls1_File As Integer


Dim cnSIDE_DB As New ADODB.Connection, rsSIDE_DB As New ADODB.Recordset, rsSIDE_Loop As New ADODB.Recordset
Dim rsSIDE_X As New ADODB.Recordset

Dim xrMesg As typerMesg, xrIntv As typerIntv
Dim wAmj8_tiret As String, xAmj8_from_crea_date_time As String, xAmj8_to_crea_date_time As String

Dim mfraDetail_Width As Integer
Dim mGOSDOSIDD_Last As Long
Dim arrService_Code(100) As String, arrService_Lib(100) As String, arrService_Mail(100, 2) As String
Dim arrService_Code_SAA(100) As String, mParam_Mail_K As Integer


Dim mYGOSDOS0_Fct As String, mYGOSEVE0_Fct  As String, blnYGOSDOS0_Display As Boolean, blnYGOSDOS0_Update As Boolean
Dim paramGOSDOS_Path As String, paramGOSDOS_Path_DROPI As String
Dim oldFileName As String, newFileName As String, newDirPath As String, newFileExtension As String

Dim xYSWISAB0 As typeYSWISAB0, newYSWISAB0 As typeYSWISAB0, oldYSWISAB0 As typeYSWISAB0

Dim Old_YBIATAB0 As typeYBIATAB0, New_YBIATAB0 As typeYBIATAB0


Dim blnSwift_Display As Boolean

Dim oldZSWIENA0 As typeZSWIENA0, newZSWIENA0 As typeZSWIENA0
Dim mZSWIENA0_Fct As String

Dim rsSabX As New ADODB.Recordset
Dim importSWISABSWID As Long, autoSWISABSWID As Long, autoSWISABZSWI As Long

Dim HeightOfLine As Long, LinesOfText As Long
Dim Mesg_aid As Long, mesg_s_umidl As Long, mesg_s_umidh As Long
Dim mSWISABSWID As Long

Dim rtextField_Value As Variant
Dim fgSwift_FormatString As String

Dim xParam As typeYGOSEVE0, newParam As typeYGOSEVE0, oldParam As typeYGOSEVE0
Dim blnYGOSDOS0_New As Boolean, cmdSelect_SQL_Kbis As String


Dim arrJrnl_Nb As Long, arrJrnl_Comp_Name() As String, arrJrnl_Event_Num() As Long, arrJrnl_Alerte() As String, arrJrnl_Top() As String
Dim xrJrnl As typerJrnl
Dim newYSAAJRN0 As typeYSAAJRN0, oldYSAAJRN0 As typeYSAAJRN0
Dim arrJrnl_Event_Id() As String, arrJrnl_Event_Lib() As String, arrJrnl_Event_Nb As Integer, arrJrnl_Event_K As Integer

Dim xrText As typerText
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


Private Sub fgDetail_Display()
Dim wColor As Long, wColorFixed As Long
Dim X As String, xWhere As String, xOPE As String
Dim xSql As String
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String


On Error GoTo Error_Handler
currentAction = "fgDetail_Display"
fgDetail.Visible = False: fraDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.BackColorFixed = fgSelect.BackColorFixed
fgDetail.Row = 0
fgDetail_50 = "": fgDetail_59 = "": fgDetail_57 = ""
fgDetail_70 = "": fgDetail_72 = ""


If cmdSelect_SQL_K = "1trf" Then
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0," & paramIBM_Library_SABSPE & ".YSWISAB1 where SWISABWID1 = " & Mesg_aid _
    & " and SWISABWIDL = " & mesg_s_umidl _
    & " and SWISABWIDH = " & mesg_s_umidh _
    & " and SWISAB1ID = SWISABSWID"

Else
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABWID1 = " & Mesg_aid _
    & " and SWISABWIDL = " & mesg_s_umidl _
    & " and SWISABWIDH = " & mesg_s_umidh
End If

Set rsSab = cnsab.Execute(xSql)
If rsSab.EOF Then

    If Mid$(cmdSelect_SQL_K, 1, 1) = "1" Or Mid$(cmdSelect_SQL_K, 1, 1) = "J" Then Call fgSwift_Display(0)
Else
    Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
    If Mid$(cmdSelect_SQL_K, 1, 1) = "1" Or Mid$(cmdSelect_SQL_K, 1, 1) = "J" Then
        fgSwift_Display rsSab("SWISABSWID")
    Else
        If oldYSWISAB0.SWISABWES = "E" Then
            X = "reçu de "
            wColor = RGB(190, 240, 255)
            wColorFixed = vbBlue
        Else
            X = "émis vers "
            wColor = RGB(220, 255, 220)
            wColorFixed = RGB(0, 64, 0)
        End If
        libDetail_SWISABSWID = "SAB : " & Trim(oldYSWISAB0.SWISABOPEC) & " " & Format(oldYSWISAB0.SWISABOPEN, "### ###")
        
    
        libDetail_SWISABSWID.BackColor = wColor
        fgDetail.Col = 0: fgDetail.Text = oldYSWISAB0.SWISABWMTK
        fgDetail.CellFontBold = True: fgDetail.CellBackColor = wColor
        fgDetail.ForeColorFixed = wColorFixed
        fgDetail.Col = 1: fgDetail.Text = X & oldYSWISAB0.SWISABWBIC & " le " & dateImp10(oldYSWISAB0.SWISABWAMJ) & " " & timeImp8(oldYSWISAB0.SWISABWHMS)
        fgDetail.CellFontBold = True: fgDetail.CellBackColor = wColor
        fgDetail.ForeColorFixed = wColorFixed
        
        
        xSql = "select * from rtextField " _
            & "where Aid = " & Mesg_aid _
            & " and text_s_umidl = " & mesg_s_umidl _
            & " and text_s_umidh  =  " & mesg_s_umidh _
            & " order by field_cnt"

        Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
        If Not rsSIDE_DB.EOF Then
            Do While Not rsSIDE_DB.EOF
            
                fgDetail.Rows = fgDetail.Rows + 1
                fgDetail.Row = fgDetail.Rows - 1
            
                fgDetail_DisplayLine fgDetail.Row
            
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
                fgDetail_DisplayLine_rText fgDetail.Row   ' , wColor, wColorFixed
            End If
        End If
        tabDetail.Tab = 0
        fgDetail.Visible = True
    End If
End If



'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


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
mSWISABSWID = lSWISABSWID
'----------------------------------------------------------------
blnOk = False
If lSWISABSWID > 0 Then
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABSWID = " & lSWISABSWID
    Set rsSab = cnsab.Execute(xSql)
    
    If Not rsSab.EOF Then
        blnOk = True
        Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
        Mesg_aid = oldYSWISAB0.SWISABWID1
        mesg_s_umidl = oldYSWISAB0.SWISABWIDL
        mesg_s_umidh = oldYSWISAB0.SWISABWIDH

    End If
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

If Mid$(cmdSelect_SQL_K, 1, 1) = "1" Or Mid$(cmdSelect_SQL_K, 1, 1) = "J" Then
    Set fraSwift.Container = fraTab0
    fraSwift.Top = fraTab0.Top + 1600
    fraSwift.Left = fraTab0.Width - fraSwift.Width - 50

    If chkSIDE_DB_Show Then Call frmSIDE_DB.fgSwift_Display(lSWISABSWID, Mesg_aid, mesg_s_umidl, mesg_s_umidh)
    
Else
    Set fraSwift.Container = fraDetail
    fraSwift.Top = 600
    fraSwift.Left = fraDetail.Width - 6000
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
Public Sub fgDetail_DisplayLine(lIndex As Long)
Dim K As Integer, iAsc13 As Integer, iLen As Integer
Dim blnAsc13 As Boolean
Dim xValue As String

On Error Resume Next
fgDetail.Col = 0: fgDetail.Text = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
fgDetail.Col = 1: xValue = Trim(rsSIDE_DB("value"))
Select Case rsSIDE_DB("field_code")
    Case "50": fgDetail_50 = xValue
    Case "57": fgDetail_57 = xValue
    Case "59": fgDetail_59 = xValue
    Case "70": fgDetail_70 = xValue
    Case "72": fgDetail_72 = xValue
End Select
 iLen = Len(xValue)
 K = 1
 Do
    iAsc13 = InStr(K, xValue, Asc13)
    If iAsc13 > 0 Then
        fgDetail.Text = Trim(Mid$(xValue, K, iAsc13 - K))
        K = iAsc13 + 2
        fgDetail.Rows = fgDetail.Rows + 1
        fgDetail.Row = fgDetail.Rows - 1
    End If
 Loop Until iAsc13 = 0

fgDetail.Text = Trim(Mid$(xValue, K, iLen - K + 1))
fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = rsSIDE_DB("field_cnt")



End Sub



Public Sub fgDetail_DisplayLine_rText(lIndex As Long)
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
        fgDetail.Rows = fgDetail.Rows + 1
        fgDetail.Row = fgDetail.Rows - 1
        X = Trim(Mid$(xValue, K, iAsc13 - K))
        fgDetail.Col = 1
        'fgDetail.CellForeColor = lColorFixed
        If Mid$(X, 1, 1) <> ":" Then
            fgDetail.Text = Trim(Mid$(xValue, K, iAsc13 - K))
        Else
            K2 = InStr(2, X, ":")
            If K2 > 0 Then
                fgDetail.Text = Trim(Mid$(X, K2 + 1, Len(X) - K2))
                fgDetail.Col = 0: fgDetail.Text = Trim(Mid$(X, 2, K2 - 2))
                'fgdetail.CellBackColor = lCellBackColor
                'fgdetail.CellForeColor = lColorFixed
            Else
                fgDetail.Text = Trim(Mid$(xValue, K, iAsc13 - K))
            End If
        End If
        
        K = iAsc13 + 2
    End If
 Loop Until iAsc13 = 0

'fgDetail.Text = Trim(Mid$(xValue, K, iLen - K + 1))
'fgDetail.Col = fgDetail.Cols - 1: fgDetail.Text = rsSIDE_DB("field_cnt")

K = InStr(xValue, ":50") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_50 = Mid$(xValue, K, K2 - K - 3)
    End If
End If
K = InStr(xValue, ":59") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_59 = Mid$(xValue, K, K2 - K - 3)
    End If
End If

K = InStr(xValue, ":57") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_57 = Mid$(xValue, K, K2 - K - 3)
    End If
End If

K = InStr(xValue, ":70") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_70 = Mid$(xValue, K, K2 - K - 3)
    End If
End If

K = InStr(xValue, ":72") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_72 = Mid$(xValue, K, K2 - K - 3)
    End If
End If

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

Public Sub fgdetail_Sort()
If fgDetail.Rows > 1 Then
        fgDetail.Visible = False

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
        fgDetail.Visible = True
End If

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
fgSelect.FormatString = Replace(fgSelect_FormatString, "Information", "Date valeur")
fgSelect.Row = 0

currentAction = "fgSelect_Display"
    
Do While Not rsSIDE_DB.EOF
    'V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine I


    rsSIDE_DB.MoveNext

Loop
         
         
If fgSelect.Rows > 2 Then fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgSelect_Display_rJrnl()
Dim wColor As Long, K As Integer
Dim V, mComp_Name
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler ' Resume Next  '
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<App|<Evénement                       |< Horodatage             " _
                      & "|<Libellé                                                                                                                                                                    " _
                      & "|<Opérateur       |<Classe                 |<Sévérité           " _
                      & "|<Application                   |<Fonction        " _
                      & " |<Alarm                     |||"
fgSelect.Row = 0
wColor = 0
currentAction = "fgSelect_Display_rJrnl"
    
Do While Not rsSIDE_DB.EOF
    'V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    'Debug.Print rsSIDE_DB("jrnl_rev_date_time")
    fgSelect.Col = 3:  V = rsSIDE_DB("Jrnl_merged_text"): If Not IsNull(V) Then fgSelect.Text = V
    fgSelect.Col = 0:  mComp_Name = rsSIDE_DB("Jrnl_comp_name"): If Not IsNull(V) Then fgSelect.Text = mComp_Name
    fgSelect.Col = 1:  V = rsSIDE_DB("Jrnl_event_name"): If Not IsNull(V) Then fgSelect.Text = rsSIDE_DB("Jrnl_event_num") & " " & V
    fgSelect.Col = 2:  V = rsSIDE_DB("Jrnl_date_time"): If Not IsNull(V) Then fgSelect.Text = V
    fgSelect.Col = 4:  V = rsSIDE_DB("Jrnl_oper_nickname"): If Not IsNull(V) Then fgSelect.Text = V
    fgSelect.Col = 5:  V = rsSIDE_DB("Jrnl_event_class"): If Not IsNull(V) Then fgSelect.Text = V
    fgSelect.Col = 6:  V = rsSIDE_DB("Jrnl_event_severity")
    If Not IsNull(V) Then
        fgSelect.Text = V
        Select Case V
            Case "FATAL"
                For K = 0 To 9: fgSelect.Col = K: fgSelect.CellBackColor = vbRed: Next K
            Case "SEVERE"
                For K = 0 To 9: fgSelect.Col = K: fgSelect.CellForeColor = vbRed: Next K
            Case "WARNING"
                If mComp_Name <> "OFCS" Then
                    For K = 0 To 9: fgSelect.Col = K: fgSelect.CellForeColor = vbMagenta: Next K
                End If
        End Select
    End If
    
    fgSelect.Col = 7:  V = rsSIDE_DB("Jrnl_appl_serv_name"): If Not IsNull(V) Then fgSelect.Text = V
    fgSelect.Col = 8:  V = rsSIDE_DB("Jrnl_func_name"): If Not IsNull(V) Then fgSelect.Text = V
    fgSelect.Col = 9:  V = rsSIDE_DB("Jrnl_alarm_status"): If Not IsNull(V) Then fgSelect.Text = V
    fgSelect.Col = 10:  V = rsSIDE_DB("Aid"): If Not IsNull(V) Then fgSelect.Text = V
    fgSelect.Col = 11:  V = rsSIDE_DB("Jrnl_rev_date_time"): If Not IsNull(V) Then fgSelect.Text = V
    fgSelect.Col = 12:  V = rsSIDE_DB("Jrnl_seq_nbr"): If Not IsNull(V) Then fgSelect.Text = V
    rsSIDE_DB.MoveNext

Loop
         
fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
fgSelect.Visible = True
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgSelect_Display_YSAAJRN0()
Dim K As Integer, X As String
Dim V

Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler ' Resume Next  '
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Date                             |<Evénement       " _
                      & "|<Libellé                                                       " _
                      & "| = |<Informations                                                                       |||"
fgSelect.Row = 0

currentAction = "fgSelect_Display_YSAAJRN0"
    
Do While Not rsSab.EOF
  
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    V = DateAdd("s", 1200802192 - rsSab("SAAJRNAMJH"), "01/01/2000 00:00:00")
    
    fgSelect.Col = 0:  fgSelect.Text = V
    X = rsSab("SAAJRNEVEC") & " " & rsSab("SAAJRNEVEN")
    fgSelect.Col = 1: fgSelect.Text = X
    fgSelect.Col = 2: fgSelect.Text = fgSelect_Display_YSAAJRN0_Lib(X)
    fgSelect.Col = 3:  fgSelect.Text = rsSab("SAAJRNTOPK")
    fgSelect.Col = 4:  fgSelect.Text = Trim(rsSab("SAAJRNTOPX"))
    fgSelect.Col = 5:  fgSelect.Text = rsSab("SAAJRNAID")
    fgSelect.Col = 6:  fgSelect.Text = rsSab("SAAJRNAMJH")
    fgSelect.Col = 7:  fgSelect.Text = rsSab("SAAJRNSEQ")
    fgSelect.Col = 8:  fgSelect.Text = rsSab("SAAJRNSUFX")
    rsSab.MoveNext

Loop
         
fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
fgSelect.Visible = True
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


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
    
    fraSelect_Options.Visible = True
    fraSelect_Options_J.Visible = False
    Select Case cmdSelect_SQL_K
       Case "J*"
           fraSelect_Options.Visible = False
           fraSelect_Options_J.Visible = True
           cboSelect_J_event_name_Top.Visible = False
           cboSelect_J_event_severity.Visible = True
           cboSelect_J_event_class.Visible = True
           cboSelect_J_appl_serv_name.Visible = True
           txtSelect_J_oper_nickname.Visible = True
           cmdSelect_Ok.Visible = True
        Case "J="
           fraSelect_Options.Visible = False
           fraSelect_Options_J.Visible = True
           cboSelect_J_event_name_Top.Visible = True
           cboSelect_J_event_severity.Visible = False
           cboSelect_J_event_class.Visible = False
           cboSelect_J_appl_serv_name.Visible = False
           txtSelect_J_oper_nickname.Visible = False
           cmdSelect_Ok.Visible = True
      Case Else
           fraSelect_Options.Visible = False
           cmdSelect_Ok.Visible = True
    End Select

End If
End Sub

Private Sub cmdSelect_SQL_JPL()
Dim V
Dim xSql As String, X As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean
Dim wSwift_address_K As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_JPL"
blnOk = False
xWhere = ""
Dim wDateFrom, wDateTo

V = cmdSelect_SQL_rJrnl_Date("02/01/2012 00:00:00", "16/01/2012 23:00:00", wDateFrom, wDateTo)

xSql = "select count(*)  from rJrnl " _
          & "where Aid = 0 " _
          & " and jrnl_rev_date_time >= " & wDateTo _
          & " and jrnl_rev_date_time <= " & wDateFrom

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
        If Not rsSIDE_DB.EOF Then Debug.Print rsSIDE_DB(0)
Exit Sub

'=====================================================================================
xSql = "select distinct jrnl_comp_name , jrnl_event_num, jrnl_event_name from rJrnl " _
          & "where Aid = 0 " _
          & " and jrnl_rev_date_time >= " & 829000000 _
          & " and jrnl_rev_date_time <= " & 950000000 & " and jrnl_comp_name = 'RMS'" _
          & " order by jrnl_comp_name , jrnl_event_num , jrnl_event_name"
          
'          & " and substring(jrnl_display_text,1,7) = 'Message'" _

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
        Do While Not rsSIDE_DB.EOF
            Debug.Print Trim(rsSIDE_DB(0)) & " - " & Trim(rsSIDE_DB(1)) & " - " & Trim(rsSIDE_DB(2))
           ' If rsSIDE_DB(0) = "RMS" Then
          '      V = ""
         '   End If
            rsSIDE_DB.MoveNext
        
        Loop
Exit Sub



xSql = "select * from  rAppe " _
    & "where appe_network_delivery_status = 'DLV_NACKED' and appe_inst_num = 0" _
    & " order by appe_date_time , appe_seq_nbr"
Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    
    If Not rsSIDE_DB.EOF Then
        Do While Not rsSIDE_DB.EOF
            Debug.Print rsSIDE_DB("appe_date_time"), rsSIDE_DB("appe_seq_nbr")
            'fgSwift.Rows = fgSwift.Rows + 1
            'fgSwift.Row = fgSwift.Rows - 1
        
            'fgSwift_DisplayLine fgSwift.Row, wColor, wColorFixed
        
            rsSIDE_DB.MoveNext
        
        Loop
    End If
Exit Sub

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wBackColor As Long
Dim xUUMID As String
Dim X As String

On Error Resume Next
xUUMID = rsSIDE_DB("mesg_uumid")
If Mid$(xUUMID, 1, 1) = "I" Then
    X = rsSIDE_DB("mesg_type") & " S"
    wColor = RGB(16, 96, 16)
    wBackColor = mColor_W0
Else
    X = rsSIDE_DB("mesg_type") & " E"
    wColor = vbBlue
    wBackColor = mColor_B0
End If

fgSelect.Col = 0: fgSelect.Text = X
fgSelect.CellForeColor = wColor

fgSelect.Col = 1: fgSelect.Text = Mid$(xUUMID, 2, 11)
fgSelect.CellForeColor = wColor
fgSelect.Col = 5: fgSelect.Text = rsSIDE_DB("x_fin_value_date")
fgSelect.CellForeColor = wColor
fgSelect.Col = 4: fgSelect.Text = rsSIDE_DB("x_fin_ccy")
fgSelect.CellForeColor = wColor
fgSelect.Col = 3: fgSelect.Text = Format$(CCur(rsSIDE_DB("x_fin_amount")), "### ### ### ##0.00")
fgSelect.CellForeColor = vbRed
fgSelect.CellFontBold = True
fgSelect.Col = 2: fgSelect.Text = rsSIDE_DB("mesg_trn_ref")
fgSelect.CellForeColor = wColor
fgSelect.Col = 6
X = rsSIDE_DB("mesg_crea_date_time")
fgSelect.Text = Mid$(X, 7, 4) & "-" & Mid$(X, 4, 2) & "-" & Mid$(X, 1, 2) & "    " & Mid$(X, 12, 8)
fgSelect.CellForeColor = RGB(80, 80, 80)


fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
fgSelect.Col = 8: fgSelect.Text = rsSIDE_DB("aid")
fgSelect.Col = 9: fgSelect.Text = rsSIDE_DB("mesg_s_umidl")
fgSelect.Col = 10: fgSelect.Text = rsSIDE_DB("mesg_s_umidh")

If Trim(rsSIDE_DB("mesg_status")) = "COMPLETED" Then
    fgSelect.Col = 7: fgSelect.Text = Trim(rsSIDE_DB("inst_mpfn_name")) 'xUUMID
    fgSelect.CellForeColor = wColor
Else
    fgSelect.Col = 7: fgSelect.Text = " " & Trim(rsSIDE_DB("inst_mpfn_name")) & "  =>  " & Trim(rsSIDE_DB("inst_rp_name"))
    fgSelect.CellForeColor = wColor
    fgSelect.CellFontBold = True
    For K = 0 To 10
        fgSelect.Col = K: fgSelect.CellBackColor = wBackColor
    Next K
End If

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


'Select Case rsSIDE_DB("field_code")
'    Case "45", "46", "47":   xValue = rsSIDE_DB("value_memo")
'    Case Else:    xValue = rsSIDE_DB("value")
'End Select

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


Public Sub fgSelect_Sort()
If fgSelect.Rows > 1 Then
    fgSelect.Visible = False
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
    fgSelect.Visible = True
End If

End Sub

Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long
fgSelect.Visible = False

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = lK
    Select Case lK
        Case 3: fgSelect.Col = 3: X = Format$(Val(fgSelect.Text), "000000000000000.00")
        Case 4:
            fgSelect.Col = 4: X = Trim(fgSelect.Text)
            fgSelect.Col = 3: X = X & Format$(Val(fgSelect.Text), "000000000000000.00")
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

wFct = UCase$(Trim(Mid$(Msg, 1, 12)))
'JPL HAB Call BiaPgmAut_Init(wFct, YGOSDOS0_Aut)
'JPL HAB Call BiaPgmAut_Init("YSWISAB0", YSWISAB0_Aut)
Call BIA_VB_HAB(wFct, arrHab(), cboSelect_SQL)

'paramIBM_Library_JPL073 = "JPLTST"
'paramIBM_Library_JPL073SPE = "JPLTST"

'blnSetfocus = True

Select Case wFct
    Case Else: blnAuto = False: Form_Init

End Select
End Sub


Public Sub Form_Init()
Dim V, xSql As String, X As String
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

fgDetail.Visible = False: fraDetail.Visible = False
fgDetail_FormatString = fgDetail.FormatString


Set fraSwift.Container = fraDetail
fraSwift.Visible = False
fraSwift.Top = 1430
fraSwift.Height = 7200
fraSwift.Left = fraDetail.Width - fraSwift.Width - 500
libSWIFT_SWISABSWID.BackColor = mColor_Y0
libSWIFT_SWISABSWID.ForeColor = vbMagenta 'RGB(128, 64, 0)
fgSwift_FormatString = fgSwift.FormatString


lstW.Visible = False
lstW.Clear

'Initialisation Service ______________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdPrint, "Initialisation Services ")

arrService_Code(0) = "S00": arrService_Lib(0) = "none"
arrService_Code_SAA(0) = "SOBF" 'arrService_Code(K)
arrService_Mail(0, 1) = "": arrService_Mail(0, 2) = ""

For K = 1 To 99
     arrService_Code(K) = "S" & Format$(K, "00"): arrService_Lib(K) = arrService_Code(K)
     arrService_Code_SAA(K) = "SOBF" 'arrService_Code(K)
     arrService_Mail(K, 1) = "": arrService_Mail(K, 2) = ""
Next K


xSql = "select *from " & paramIBM_Library_SABSPE & ".YSSIUSR0 where SSIUSRNAT= 'S' order by SSIUSRUNIT"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    X = rsSab("SSIUSRUNIT") & " - " & rsSab("SSIUSRPRFX")
    K = Val(Mid$(rsSab("SSIUSRUNIT"), 2, 2))
    arrService_Code(K) = rsSab("SSIUSRUNIT")
    arrService_Lib(K) = Trim(rsSab("SSIUSRPRFX"))
    rsSab.MoveNext
Loop


'__________________________________________________________________________________________________________


'__________________________________________________________________________________________________________

'________________________________________________________________________________________________________
fraSelect_Options.Visible = True

If arrHab(15) Then Form_Init_Options_J


If cboSelect_SQL.ListCount > 0 Then Call cbo_Scan("J=", cboSelect_SQL)
blnControl = True


cmdSelect_Reset
Me.Enabled = True
On Error Resume Next
cmdSelect_Ok.SetFocus

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






Private Sub cboSelect_SQL_Click()
cmdSelect_Reset
End Sub


Private Sub chkSAB_Dossier_DB_Show_Click()
On Error Resume Next
Dim K As Integer
If fraSwift.Visible = True Then
    If chkSAB_Dossier_DB_Show = "1" Then
        'If mMOUVEMNUM > 0 Then Call frmSAB_Dossier_DB.Form_Init("","", "", "", mMOUVEMSER, mMOUVEMSSE, mMOUVEMOPE, mMOUVEMNUM)
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

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Dim X As String, I As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, ">cmdPrint : Initialisation ")

Select Case SSTab1.Tab
            Case Else: cmdPrint_Excel
        End Select
        
Call lstErr_AddItem(lstErr, cmdPrint, "<cmdPrint : terminé ")

Me.Enabled = True: Me.MousePointer = 0
End Sub

Public Sub cmdPrint_Excel()
On Error GoTo Error_Handler
Dim xSql As String
Dim X As String, wFilex As String
Dim blnCALCS As Boolean

On Error GoTo Error_Handler
'===================================================================================
'______________________________________________'
X = paramServer("\\CDO_Archive\")
wAmjMin = DSys

blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

If X = "" Then X = "C:\Temp\"
If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"

mXls1_File = mXls1_File + 1

wFile = X & Trim("Swift " & DSYS_Time & mXls1_File & ".xlsx")
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "Swift : nom du fichier d'exportation", wFile)
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
    .Title = "Swift"
    .Subject = "Messages"
End With

'__________________________________________________________________________________

'appExcel.Worksheets.Add

Set wsExcel = wbExcel.Sheets(1): wsExcel.Name = "SWIFT"

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

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Swift, arrêté au " & dateImp10(wAmjMin) _
                                 & vbCr

wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

mXls1_Row = 1

Select Case SSTab1.Tab
    Case 0:
        Select Case cmdSelect_SQL_K
            Case "J*", "J=":   cmdPrint_Excel_rJrnl
4            Case Else:    cmdPrint_Excel_rMesg
        End Select
    Case 1
        

    End Select
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
Public Sub cmdPrint_Excel_rMesg()
On Error GoTo Error_Handler
Dim xSql As String
Dim X As String
Dim K As Integer
Dim wForecolor As Long

'On Error GoTo Error_Handler
'===================================================================================


wsExcel.Columns(1).ColumnWidth = 5: wsExcel.Cells(mXls1_Row, 1) = "Type": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 5: wsExcel.Cells(mXls1_Row, 2) = "Sens": wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(3).ColumnWidth = 15: wsExcel.Cells(mXls1_Row, 3) = "BIC": wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(4).ColumnWidth = 15: wsExcel.Cells(mXls1_Row, 4) = "Référence": wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(5).ColumnWidth = 15: wsExcel.Cells(mXls1_Row, 5) = "Montant": wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(6).ColumnWidth = 5: wsExcel.Cells(mXls1_Row, 6) = "Devise": wsExcel.Columns(6).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(7).ColumnWidth = 11: wsExcel.Cells(mXls1_Row, 7) = "Information": wsExcel.Columns(7).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(8).ColumnWidth = 16: wsExcel.Cells(mXls1_Row, 8) = "Date de réception SAA": wsExcel.Columns(8).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(9).ColumnWidth = 12: wsExcel.Cells(mXls1_Row, 9) = "Service": wsExcel.Columns(9).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(10).ColumnWidth = 15: wsExcel.Cells(mXls1_Row, 10) = "référence SAB": wsExcel.Columns(10).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(11).ColumnWidth = 35: wsExcel.Cells(mXls1_Row, 11) = "Message": wsExcel.Columns(11).HorizontalAlignment = Excel.xlHAlignLeft


For K = 1 To 11
    wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row, K).Font.Color = mColor_Z0
    
Next K

'__________________________________________________________________________________

For K = 1 To fgSelect.Rows - 1
    fgSelect.Row = K
    mXls1_Row = mXls1_Row + 1
    X = Trim(fgSelect.Text)
    fgSelect.Col = 0:
    wForecolor = fgSelect.CellForeColor
    'wBackColor = cmdSendMail_xls_Color(fgSelect.CellBackColor)

    X = Trim(fgSelect.Text)
    wsExcel.Cells(mXls1_Row, 1) = Mid$(X, 1, 3): wsExcel.Cells(mXls1_Row, 1).Font.Color = wForecolor
    wsExcel.Cells(mXls1_Row, 2) = Mid$(X, 5, 1): wsExcel.Cells(mXls1_Row, 2).Font.Color = wForecolor
                      
    fgSelect.Col = 1: wsExcel.Cells(mXls1_Row, 3) = Trim(fgSelect.Text): wsExcel.Cells(mXls1_Row, 3).Font.Color = wForecolor
    fgSelect.Col = 2: wsExcel.Cells(mXls1_Row, 4) = Trim(fgSelect.Text): wsExcel.Cells(mXls1_Row, 4).Font.Color = wForecolor
    fgSelect.Col = 3: wsExcel.Cells(mXls1_Row, 5) = Trim(fgSelect.Text): wsExcel.Cells(mXls1_Row, 5).Font.Color = wForecolor
    fgSelect.Col = 4: wsExcel.Cells(mXls1_Row, 6) = Trim(fgSelect.Text): wsExcel.Cells(mXls1_Row, 6).Font.Color = wForecolor
    fgSelect.Col = 5: wsExcel.Cells(mXls1_Row, 7) = Trim(fgSelect.Text): wsExcel.Cells(mXls1_Row, 7).Font.Color = wForecolor
    fgSelect.Col = 6: wsExcel.Cells(mXls1_Row, 8) = Trim(fgSelect.Text): wsExcel.Cells(mXls1_Row, 8).Font.Color = wForecolor
    
    fgSelect.Col = 8: Mesg_aid = Val(fgSelect.Text)
    fgSelect.Col = 9: mesg_s_umidl = Val(fgSelect.Text)
    fgSelect.Col = 10: mesg_s_umidh = Val(fgSelect.Text)
'____________________________________________________________________________________________
    
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABWID1 = " & Mesg_aid _
        & " and SWISABWIDL = " & mesg_s_umidl _
        & " and SWISABWIDH = " & mesg_s_umidh

        Set rsSab = cnsab.Execute(xSql)
        If Not rsSab.EOF Then
            wsExcel.Cells(mXls1_Row, 9) = arrService_Lib(Val(Mid$(rsSab("SWISABKSRV"), 2, 2)))
            wsExcel.Cells(mXls1_Row, 9).Font.Color = wForecolor
            wsExcel.Cells(mXls1_Row, 10) = rsSab("SWISABSER") & " " & rsSab("SWISABSSE") & " " & rsSab("SWISABOPEC") & " " & rsSab("SWISABOPEN")
            wsExcel.Cells(mXls1_Row, 10).Font.Color = wForecolor
        End If
'____________________________________________________________________________________________
    X = ""
    xSql = "select * from rtextField " _
        & "where Aid = " & Mesg_aid _
        & " and text_s_umidl = " & mesg_s_umidl _
        & " and text_s_umidh  =  " & mesg_s_umidh _
        & " order by field_cnt"
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    If Not rsSIDE_DB.EOF Then
        Do While Not rsSIDE_DB.EOF
        
            X = X & rsSIDE_DB("field_code") & rsSIDE_DB("field_option") & " : " & Trim(rsSIDE_DB("value")) & vbCrLf
        
            rsSIDE_DB.MoveNext
        
        Loop
    Else
        xSql = "select * from rtext " _
            & "where Aid = " & Mesg_aid _
            & " and text_s_umidl = " & mesg_s_umidl _
            & " and text_s_umidh  =  " & mesg_s_umidh
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
        If Not rsSIDE_DB.EOF Then
            X = X & rsSIDE_DB("text_data_block")
        End If
    End If

    wsExcel.Cells(mXls1_Row, 11) = X
    wsExcel.Cells(mXls1_Row, 11).Font.Color = wForecolor

Next K

'======================================================================================================

Exit_sub:
'__________________________________________________________________________________


'_____________________________
Exit Sub

Error_Handler:
End Sub
Public Sub cmdPrint_Excel_rJrnl()
On Error GoTo Error_Handler
Dim xSql As String
Dim X As String
Dim K As Integer, K2 As Integer
Dim wForecolor As Long, wBackColor As Long

'On Error GoTo Error_Handler
'===================================================================================
wsExcel.Cells.HorizontalAlignment = Excel.xlHAlignLeft
fgSelect.Visible = False

    For K = 0 To fgSelect.Rows - 1
        fgSelect.Row = K
        
        wForecolor = fgSelect.CellForeColor
        If wForecolor = 0 Then
            If K = 0 Then
                wForecolor = fgSelect.ForeColorFixed
            Else
                wForecolor = fgSelect.ForeColor
            End If
        End If
        
        wBackColor = fgSelect.CellBackColor
        If wBackColor = 0 Then
            If K = 0 Then
                wBackColor = fgSelect.BackColorFixed
            Else
                wBackColor = fgSelect.BackColor
            End If
        End If

        mXls1_Row = mXls1_Row + 1
        For K2 = 0 To 9
        
            fgSelect.Col = K2: X = Trim(fgSelect.Text)
            If K = 0 Then
                wsExcel.Columns(K2 + 1).ColumnWidth = fgSelect.CellWidth / 100
                'If K2 > 0 Then wsExcel.Columns(K2 + 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            End If
            wsExcel.Cells(mXls1_Row, K2 + 1) = X
            wsExcel.Cells(mXls1_Row, K2 + 1).Font.Color = wForecolor
            wsExcel.Cells(mXls1_Row, K2 + 1).Interior.Color = wBackColor
        Next K2
    Next K


'======================================================================================================

Exit_sub:
'__________________________________________________________________________________

fgSelect.Visible = True
'_____________________________
Exit Sub

Error_Handler:
fgSelect.Visible = True
End Sub

Public Sub cmdPrint_Excel_YBIATAB0_SAA()
Dim xSql As String, X As String, K As Long
On Error GoTo Error_Handler


'On Error GoTo Error_Handler
'===================================================================================
With wsExcel.Cells
    .HorizontalAlignment = Excel.xlHAlignLeft
    .Font.Size = 9
    .Font.Name = "Courier New"
End With
wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 100
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14SAA_Alerte : paramétrage" _
                                & "  (édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$E1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


wsExcel.Columns(1).ColumnWidth = 10: wsExcel.Cells(1, 1) = "Id "
wsExcel.Columns(2).ColumnWidth = 15: wsExcel.Cells(1, 2) = "Nature"
wsExcel.Columns(3).ColumnWidth = 15: wsExcel.Cells(1, 3) = "Code"
wsExcel.Columns(4).ColumnWidth = 10: wsExcel.Cells(1, 4) = ""
wsExcel.Columns(5).ColumnWidth = 65: wsExcel.Cells(1, 5) = "Libellé"

For K = 1 To 5
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = vbWhite
Next

xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'SAA'" _
     & " order by BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF

    mXls1_Row = mXls1_Row + 1
    wsExcel.Cells(mXls1_Row, 1) = rsSab("BIATABID")
    wsExcel.Cells(mXls1_Row, 2) = rsSab("BIATABK1")
    wsExcel.Cells(mXls1_Row, 3) = rsSab("BIATABK2")
    wsExcel.Cells(mXls1_Row, 5) = Trim(Mid$(rsSab("BIATABTXT"), 1, 99))
    
        
    Select Case Trim(rsSab("BIATABK1"))
        Case "Amount", "Approval"
            wsExcel.Cells(mXls1_Row, 4) = Format(CCur(Trim(rsSab("BIATABTXT"))), "### ### ### ### ##0")
            wsExcel.Cells(mXls1_Row, 4).Interior.Color = mColor_Y1
            wsExcel.Cells(mXls1_Row, 4).HorizontalAlignment = Excel.xlHAlignRight
        Case "Jrnl_Event"
            wsExcel.Cells(mXls1_Row, 4) = Mid$(rsSab("BIATABTXT"), 100, 4)
            wsExcel.Cells(mXls1_Row, 4).Interior.Color = mColor_Y1
    End Select
    rsSab.MoveNext
Loop



'======================================================================================================

Exit_sub:
'__________________________________________________________________________________


'_____________________________
Exit Sub

Error_Handler:

End Sub

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_GOS_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Clear

Select Case cmdSelect_SQL_K
    Case "J*": cmdSelect_SQL_rJrnl
    Case "J=": cmdSelect_SQL_YSAAJRN0
    Case "JPL":
    
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_GOS_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
cmdSelect_Ok.BackColor = fgSelect.BackColorFixed
End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, xUUMID As String
On Error Resume Next


If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2:  fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3:  fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_SortX 3
        Case 4:  fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_SortX 4
        Case 5:  fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_Sort
        Case 6:  fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7:  fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8:  fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
        Case 9:  fgSelect_Sort1 = 9: fgSelect_Sort2 = 9: fgSelect_Sort
        Case 10: fgSelect_Sort1 = 10: fgSelect_Sort2 = 10: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fraSwift.Visible = False
        fgFree.Visible = False
        
        blnYGOSDOS0_New = False
        
        Select Case cmdSelect_SQL_K
                
            Case "J=":
                fgSelect.Col = 4:  oldYSAAJRN0.SAAJRNTOPX = fgSelect.Text
                fgSelect.Col = 5:  oldYSAAJRN0.SAAJRNAID = fgSelect.Text
                fgSelect.Col = 6:  oldYSAAJRN0.SAAJRNAMJH = fgSelect.Text
                fgSelect.Col = 7:  oldYSAAJRN0.SAAJRNSEQ = fgSelect.Text
                fgSelect.Col = 8:  oldYSAAJRN0.SAAJRNSUFX = fgSelect.Text
                wX = "select * from rJrnl " _
                        & " where Aid = " & oldYSAAJRN0.SAAJRNAID _
                        & " and jrnl_rev_date_time = " & oldYSAAJRN0.SAAJRNAMJH _
                        & " and jrnl_seq_nbr = " & oldYSAAJRN0.SAAJRNSEQ

                Set rsSIDE_DB = cnSIDE_DB.Execute(wX)
                If Not rsSIDE_DB.EOF Then
                    Call srvrJrnl_GetBuffer_ODBC(rsSIDE_DB, xrJrnl)
                    fgSelect.Col = 3
                    If fgSelect.Text = "U" Then
                        If IsNull(cmdSelect_SQL_YSAAJRN0_rMesg(oldYSAAJRN0.SAAJRNTOPX, oldYSAAJRN0.SAAJRNSUFX)) Then fgDetail_Display
                    Else
                        fgFree_Display_rJrnl
                    End If
                End If
             Case "J*":
                fgSelect.Col = 10:  oldYSAAJRN0.SAAJRNAID = fgSelect.Text
                fgSelect.Col = 11:  oldYSAAJRN0.SAAJRNAMJH = fgSelect.Text
                fgSelect.Col = 12:  oldYSAAJRN0.SAAJRNSEQ = fgSelect.Text
                
                wX = "select * from rJrnl " _
                        & " where Aid = " & oldYSAAJRN0.SAAJRNAID _
                        & " and jrnl_rev_date_time = " & oldYSAAJRN0.SAAJRNAMJH _
                        & " and jrnl_seq_nbr = " & oldYSAAJRN0.SAAJRNSEQ

                Set rsSIDE_DB = cnSIDE_DB.Execute(wX)
                If Not rsSIDE_DB.EOF Then
                    Call srvrJrnl_GetBuffer_ODBC(rsSIDE_DB, xrJrnl)

                    Importation_Jrnl_Top False
                    If newYSAAJRN0.SAAJRNTOPK = "U" Then
                        fgFree_Display_rJrnl
                        oldYSAAJRN0 = newYSAAJRN0
                        If IsNull(cmdSelect_SQL_YSAAJRN0_rMesg(newYSAAJRN0.SAAJRNTOPX, newYSAAJRN0.SAAJRNSUFX)) Then fgDetail_Display
                    Else
                        fgFree_Display_rJrnl
                    End If
                End If
                
        End Select
        
   End If
End If
Wait_SS 0
fgSelect.LeftCol = 0

End Sub

Public Sub Importation_Jrnl_Top(blnYSAAJRN0_New As Boolean)
Dim K1 As Integer, K2 As Integer
On Error GoTo Error_Handler

newYSAAJRN0.SAAJRNTOPK = ""
newYSAAJRN0.SAAJRNTOPX = ""
newYSAAJRN0.SAAJRNSUFX = 0

If blnYSAAJRN0_New Then
    If xrJrnl.jrnl_comp_name = "BSA" And newYSAAJRN0.SAAJRNEVEN = 3000 Then
        Select Case xrJrnl.jrnl_oper_nickname
            Case "SUPER", "RSO", "LSO":
                newYSAAJRN0.SAAJRNTOPK = "O"
                newYSAAJRN0.SAAJRNTOPX = xrJrnl.jrnl_oper_nickname
                Call sqlYSAAJRN0_Insert(newYSAAJRN0)
        End Select
        Exit Sub
'###############
    End If
End If

K1 = InStr(xrJrnl.jrnl_merged_text, "perator ")
If K1 > 0 Then
    newYSAAJRN0.SAAJRNTOPK = "O"
    K1 = K1 + 8
    K2 = InStr(K1, xrJrnl.jrnl_merged_text, ":")
    If K2 > K1 Then
        newYSAAJRN0.SAAJRNTOPX = Mid$(xrJrnl.jrnl_merged_text, K1, K2 - K1)
    End If
    If blnYSAAJRN0_New Then Call sqlYSAAJRN0_Insert(newYSAAJRN0)
    Exit Sub
'###############
End If

If xrJrnl.jrnl_comp_name = "RMS" Then
    K1 = InStr(xrJrnl.jrnl_merged_text, "orrespondent BIC: ")
    If K1 > 0 Then
        newYSAAJRN0.SAAJRNTOPK = "B"
        K1 = K1 + 18
        K2 = InStr(K1, xrJrnl.jrnl_merged_text, "\")
        If K2 > K1 Then
            newYSAAJRN0.SAAJRNTOPX = Mid$(xrJrnl.jrnl_merged_text, K1, K2 - K1)
        End If
        If blnYSAAJRN0_New Then Call sqlYSAAJRN0_Insert(newYSAAJRN0)
        Exit Sub
    '###############
    End If
End If

If xrJrnl.jrnl_comp_name = "SIS" Then
    K1 = InStr(xrJrnl.jrnl_merged_text, "UMID ")
    If K1 > 0 Then
        newYSAAJRN0.SAAJRNTOPK = "U"
        K1 = K1 + 5
        K2 = InStr(K1, xrJrnl.jrnl_merged_text, ",")
        If K2 > K1 Then
            newYSAAJRN0.SAAJRNTOPX = Mid$(xrJrnl.jrnl_merged_text, K1, K2 - K1)
        End If
    
        K1 = InStr(K2, xrJrnl.jrnl_merged_text, "uffix ")
        If K1 > 0 Then
            K1 = K1 + 6
            K2 = InStr(K1, xrJrnl.jrnl_merged_text, ":")
            If K2 > K1 Then
                newYSAAJRN0.SAAJRNSUFX = Mid$(xrJrnl.jrnl_merged_text, K1, K2 - K1)
            End If
        End If
    End If
    If blnYSAAJRN0_New Then Call sqlYSAAJRN0_Insert(newYSAAJRN0)
    Exit Sub
    '###############
End If

If xrJrnl.jrnl_comp_name = "MXS" Then
    K1 = InStr(xrJrnl.jrnl_merged_text, "UMID ")
    If K1 > 0 Then
        newYSAAJRN0.SAAJRNTOPK = "U"
        K2 = K1 + 5
        K1 = InStr(K2, xrJrnl.jrnl_merged_text, ": ")
        K1 = K1 + 2
        K2 = InStr(K1, xrJrnl.jrnl_merged_text, "\")
        If K2 > K1 Then
            newYSAAJRN0.SAAJRNTOPX = Mid$(xrJrnl.jrnl_merged_text, K1, K2 - K1)
        End If

    
        K1 = InStr(K2, xrJrnl.jrnl_merged_text, "uffix ")
        If K1 > 0 Then
            K2 = K1 + 5
            K1 = InStr(K2, xrJrnl.jrnl_merged_text, ": ")
            K1 = K1 + 2
            K2 = InStr(K1, xrJrnl.jrnl_merged_text, "\")
            If K2 > K1 Then
                newYSAAJRN0.SAAJRNSUFX = Mid$(xrJrnl.jrnl_merged_text, K1, K2 - K1)
            End If
        End If
    End If
    If blnYSAAJRN0_New Then Call sqlYSAAJRN0_Insert(newYSAAJRN0)
    Exit Sub
    '###############
End If
GoTo Exit_sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

Exit_sub:
End Sub


Public Sub fgFree_Display_rJrnl()
Dim xSql As String, K As Integer

On Error GoTo Error_Handler
'______________________________________________________________________
If arrJrnl_Field_Nb = 0 Then
    xSql = "SELECT    count(*) From sysobjects, syscolumns " _
          & " WHERE  ( sysobjects.id = syscolumns.id) And  (sysobjects.xtype = 'U') " _
          & " AND sysobjects.name LIKE 'rJrnl'"
    Set rsSIDE_X = cnSIDE_DB.Execute(xSql)
    arrJrnl_Field_Nb = rsSIDE_X(0)
    ReDim arrJrnl_Field(arrJrnl_Field_Nb + 1)
    K = 0
    
    xSql = "SELECT    syscolumns.name From sysobjects, syscolumns " _
          & " WHERE  ( sysobjects.id = syscolumns.id) And  (sysobjects.xtype = 'U') " _
          & " AND sysobjects.name LIKE 'rJrnl' ORDER BY syscolumns.colorder"
    Set rsSIDE_X = cnSIDE_DB.Execute(xSql)
    Do While Not rsSIDE_X.EOF
            arrJrnl_Field(K) = rsSIDE_X(0)
            K = K + 1
        rsSIDE_X.MoveNext
    Loop
End If
'______________________________________________________________________


Set fgFree.Container = fraTab0
fgFree.Top = 1900
fgFree.Height = 7100
fgFree.Width = 6630
fgFree.Left = 6250
fgFree.Visible = True
fgFree.Clear
fgFree.Rows = 1
fgFree.FormatString = "Champ                        |<Valeur                                                                                     |"


    'xrJrnl.jrnl_merged_text = rsSIDE_DB("Jrnl_merged_text")
    For K = 0 To 20
        fgFree.Rows = fgFree.Rows + 1
        fgFree.Row = fgFree.Rows - 1
        
        fgFree.Col = 0: fgFree.Text = arrJrnl_Field(K)
        
        If Not IsNull(rsSIDE_DB(K)) Then fgFree.Col = 1: fgFree.Text = rsSIDE_DB(K)
        If K = 6 Or K = 13 Then fgFree.CellForeColor = vbMagenta
    Next K
    fgFree.RowHeight(fgFree.Row) = 1200
    fgFree.Col = 1: fgFree.Text = xrJrnl.jrnl_merged_text
    fgFree.CellForeColor = vbBlue

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub


Private Sub fgSwift_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xField As String, xSql As String
If fgSwift.Rows > 1 Then
    fgSwift.Col = 0
    xField = Trim(fgSwift.Text)
        fgSwift.Col = 1
        If ZSWIBIC0_Select(fgSwift.Text) <> "" Then
            mnuSWIBICBIC.Caption = Trim(rsSab("SWIBICIN1")) & "  " & Trim(rsSab("SWIBICVIL")) & "  " & Trim(rsSab("SWIBICCOM"))
            Me.PopupMenu mnuZSWIBIC0, vbPopupMenuLeftButton
        End If
End If
fgSwift.Col = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If chkSIDE_DB_Show = "1" Then frmSIDE_DB.Hide
If chkSAB_Dossier_DB_Show = "1" Then frmSAB_Dossier_DB.Hide
    cnSIDE_DB.Close
    Set cnSIDE_DB = Nothing

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



If fraSwift.Visible Then
    fraSwift.Visible = False
    Exit Sub
End If
If fgFree.Visible Then fgFree.Visible = False:       Exit Sub

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
            If Not fgSelect.Visible Then cmdSelect_Ok_Click
        Else
            'SendKeys "{TAB}"
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


























Private Sub txtSelect_J_oper_nickname_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Public Sub cmdSelect_Clear()

lstErr.Clear
fgSelect.Visible = False
lstW.Visible = False
fraSwift.Visible = False
fgFree.Visible = False


blnSwift_Display = False
cmdSelect_Ok.BackColor = vbGreen
End Sub

Public Sub cmdSelect_SQL_rJrnl()
Dim xSql As String, xWhere As String, X As String
Dim wDateFrom, wDateTo
Dim wEté As String, wHiver As String, K As Integer, K2 As Integer
On Error GoTo Error_Handler

V = cmdSelect_SQL_rJrnl_Date(txtSelect_J_AMJMin, txtSelect_J_AMJMax, wDateFrom, wDateTo)

If Not IsNull(V) Then
    Call MsgBox(V, vbCritical, "BIA_GOS : J*")
    Exit Sub
End If
    
If Mid$(txtSelect_J_AMJMin, 7, 4) <> Mid$(txtSelect_J_AMJMax, 7, 4) Then
    Call MsgBox("les millésimes sont différents", vbCritical, "BIA_GOS : J*")
    Exit Sub
End If

xWhere = ""

X = Trim(cboSelect_J_comp_name)
If X <> "" Then
    If InStr(X, "%") > 0 Then
        xWhere = xWhere & " and jrnl_comp_name like '" & X & "'"
    Else
        xWhere = xWhere & " and jrnl_comp_name = '" & X & "'"
    End If
End If

X = Trim(cboSelect_J_event_name)
If X <> "" Then
    K = InStr(X, " ")
    If K > 0 Then
        K2 = InStr(K + 1, X, "-")
        If K2 > 0 Then
            xWhere = xWhere & " and jrnl_comp_name = '" & Trim(Mid$(X, 1, K - 1)) & "'" _
                   & " and jrnl_event_num = '" & Trim(Mid$(X, K + 1, K2 - K - 1)) & "'"
        End If
    End If
End If

X = Trim(cboSelect_J_event_class)
If X <> "" Then
    If InStr(X, "%") > 0 Then
        xWhere = xWhere & " and jrnl_event_class like '" & X & "'"
    Else
        xWhere = xWhere & " and jrnl_event_class = '" & X & "'"
    End If
End If

X = Trim(cboSelect_J_event_severity)
If X <> "" Then
    If InStr(X, "%") > 0 Then
        xWhere = xWhere & " and jrnl_event_severity like '" & X & "'"
    Else
        xWhere = xWhere & " and jrnl_event_severity = '" & X & "'"
    End If
End If

X = Trim(cboSelect_J_appl_serv_name)
If X <> "" Then
    If InStr(X, "%") > 0 Then
        xWhere = xWhere & " and jrnl_appl_serv_name like '" & X & "'"
    Else
        xWhere = xWhere & " and jrnl_appl_serv_name = '" & X & "'"
    End If
End If

X = Trim(txtSelect_J_oper_nickname)
If X <> "" Then
    If InStr(X, "%") > 0 Then
        xWhere = xWhere & " and jrnl_oper_nickname like '" & X & "'"
    Else
        xWhere = xWhere & " and jrnl_oper_nickname = '" & X & "'"
    End If
End If

X = Trim(txtSelect_J_merged_text)
If X <> "" Then
    xWhere = xWhere & " and jrnl_merged_text like '%" & X & "%'"
End If

    
xSql = "select * from rJrnl " _
          & "where Aid = 0 " _
          & " and jrnl_rev_date_time >= " & wDateTo _
          & " and jrnl_rev_date_time <= " & wDateFrom _
          & xWhere _
          & " order by  jrnl_rev_date_time desc , jrnl_seq_nbr desc"
          
'          & " and substring(jrnl_display_text,1,7) = 'Message'" _

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
  
fgSelect_Display_rJrnl

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Public Function cmdSelect_SQL_rJrnl_Date(lAMJMin, lAMJMax, wDateFrom, wDateTo)
Dim X As String

Dim wEté As String, wHiver As String, K As Integer
On Error GoTo Error_Handler

cmdSelect_SQL_rJrnl_Date = Null

wDateFrom = 1200802192 - DateDiff("s", "01/01/2000 00:00:00", lAMJMin)
wDateTo = 1200802192 - DateDiff("s", "01/01/2000 00:00:00", lAMJMax)

'X = DateAdd("s", 1200802192 - wDateFrom, "01/01/2000 00:00:00")

X = "31/03/" & Mid$(txtSelect_J_AMJMin, 7, 4)
K = Weekday(X)
If K > 1 Then X = DateAdd("d", -K + 1, X)
Call dateJMA_AMJ(X, wEté)

X = "31/10/" & Mid$(txtSelect_J_AMJMin, 7, 4)
K = Weekday(X)
If K > 1 Then X = DateAdd("d", -K + 1, X)
Call dateJMA_AMJ(X, wHiver)

Call dateJMA_AMJ(txtSelect_J_AMJMin, X)

If X >= wEté And X <= wHiver Then
    wDateFrom = wDateFrom + 3600
    lblSelect_J_AMJ.BackColor = &H80FFFF
Else
    lblSelect_J_AMJ.BackColor = &HE0E0E0
End If


Call dateJMA_AMJ(txtSelect_J_AMJMax, X)
If X >= wEté And X <= wHiver Then wDateTo = wDateTo + 3600
If wDateFrom < wDateTo Then
    cmdSelect_SQL_rJrnl_Date = "Date/heure de début < Date/heure de fin"
    Exit Function
End If


Exit Function

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Function

Public Sub cmdSelect_SQL_rCorr()
Dim xSql As String, K As Integer
Dim wDateFrom, wDateTo
On Error GoTo Error_Handler

xSql = "select  corr_x1 from rCorr" _
     & " where corr_x1 like 'B%' and corr_status = 'CORR_INACTIVE'"
Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

K = 0
Do While Not rsSIDE_DB.EOF
    'For K = 0 To 28
        Debug.Print K, rsSIDE_DB(K)
    'Next K
    rsSIDE_DB.MoveNext

Loop

  

'fgSelect_Display_rCorr

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub



Public Sub Form_Init_Options_J()
Dim xSql As String

fraSelect_Options_J.Visible = False
Set fraSelect_Options_J.Container = fraTab0
fraSelect_Options_J.Top = fraSelect_Options.Top
fraSelect_Options_J.Left = fraSelect_Options.Left
txtSelect_J_AMJMin = DSys_S & "  " & Time
txtSelect_J_AMJMax = txtSelect_J_AMJMin

cboSelect_J_event_name_Top.Top = cboSelect_J_event_name.Top
cboSelect_J_event_name_Top.Left = cboSelect_J_event_name.Left
cboSelect_J_event_name_Top.Visible = False

'''param_Init_Options_J
'''param_Init_Options_J_Event
'______________________________________________________________________
cboSelect_J_comp_name.Clear
cboSelect_J_comp_name.AddItem ""
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'SAA' and BIATABK1 = 'Jrnl_Comp'" _
     & "  order by BIATABK2"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    cboSelect_J_comp_name.AddItem Trim(Mid$(rsSab("BIATABTXT"), 1, 99))
    rsSab.MoveNext
Loop

'______________________________________________________________________
xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'SAA' and BIATABK1 = 'Jrnl_Event'"
Set rsSab = cnsab.Execute(xSql)
arrJrnl_Event_Nb = rsSab(0) + 1
ReDim arrJrnl_Event_Id(arrJrnl_Event_Nb), arrJrnl_Event_Lib(arrJrnl_Event_Nb)
arrJrnl_Event_Nb = 0

cboSelect_J_event_name.Clear
cboSelect_J_event_name.AddItem ""
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'SAA' and BIATABK1 = 'Jrnl_Event'" _
     & "  order by BIATABK2"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    arrJrnl_Event_Nb = arrJrnl_Event_Nb + 1
    arrJrnl_Event_Id(arrJrnl_Event_Nb) = Trim(rsSab("BIATABK2"))
    arrJrnl_Event_Lib(arrJrnl_Event_Nb) = Trim(Mid$(rsSab("BIATABTXT"), 1, 99))
    
    cboSelect_J_event_name.AddItem Trim(rsSab("BIATABK2")) & " - " & Trim(Mid$(rsSab("BIATABTXT"), 1, 99))
    rsSab.MoveNext
Loop
'______________________________________________________________________
cboSelect_J_event_name_Top.Clear
cboSelect_J_event_name_Top.AddItem ""
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'SAA' and BIATABK1 = 'Jrnl_Event' and substring(BIATABTXT,103,1) <> ''" _
     & "  order by BIATABK2"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    cboSelect_J_event_name_Top.AddItem Trim(rsSab("BIATABK2")) & " - " & Trim(Mid$(rsSab("BIATABTXT"), 1, 99))
    rsSab.MoveNext
Loop


'______________________________________________________________________
cboSelect_J_event_class.Clear
cboSelect_J_event_class.AddItem ""
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'SAA' and BIATABK1 = 'Jrnl_Class'" _
     & "  order by BIATABK2"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    cboSelect_J_event_class.AddItem Trim(Mid$(rsSab("BIATABTXT"), 1, 99))
    rsSab.MoveNext
Loop

'______________________________________________________________________
cboSelect_J_appl_serv_name.Clear
cboSelect_J_appl_serv_name.AddItem ""
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'SAA' and BIATABK1 = 'Jrnl_Serv'" _
     & "  order by BIATABK2"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    cboSelect_J_appl_serv_name.AddItem Trim(Mid$(rsSab("BIATABTXT"), 1, 99))
    rsSab.MoveNext
Loop

'______________________________________________________________________
cboSelect_J_event_severity.Clear
cboSelect_J_event_severity.AddItem ""
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'SAA' and BIATABK1 = 'Jrnl_Severit'" _
     & "  order by BIATABK2"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    cboSelect_J_event_severity.AddItem Trim(Mid$(rsSab("BIATABTXT"), 1, 99))
    rsSab.MoveNext
Loop

End Sub

Public Sub cmdSelect_SQL_YSAAJRN0()
Dim xSql As String, xWhere As String, K As Long, K2 As Long
Dim wDateFrom, wDateTo

V = cmdSelect_SQL_rJrnl_Date(txtSelect_J_AMJMin, txtSelect_J_AMJMax, wDateFrom, wDateTo)

If Not IsNull(V) Then
    Call MsgBox(V, vbCritical, "BIA_GOS : J=")
    Exit Sub
End If
xWhere = ""

X = Trim(cboSelect_J_comp_name)
If X <> "" Then
    If InStr(X, "%") > 0 Then
        xWhere = xWhere & " and SAAJRNEVEC like '" & X & "'"
    Else
        xWhere = xWhere & " and SAAJRNEVEC = '" & X & "'"
    End If
End If

X = Trim(cboSelect_J_event_name_Top)
If X <> "" Then
    K = InStr(X, " ")
    If K > 0 Then
        K2 = InStr(K + 1, X, "-")
        If K2 > 0 Then
            xWhere = xWhere & " and SAAJRNEVEC = '" & Trim(Mid$(X, 1, K - 1)) & "'" _
                   & " and SAAJRNEVEN= " & Trim(Mid$(X, K + 1, K2 - K - 1))
        End If
    End If
End If
'
X = Trim(txtSelect_J_merged_text)
If X <> "" Then
    xWhere = xWhere & " and SAAJRNTOPX like '" & X & "%'"
    'If InStr(X, "%") > 0 Then
    '    xWhere = xWhere & " and SAAJRNTOPX like '" & X & "'"
    'Else
    '    xWhere = xWhere & " and SAAJRNTOPX = '" & X & "'"
    'End If
End If

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSAAJRN0 " _
          & "where SAAJRNAID = 0 " _
          & " and SAAJRNAMJH >= " & wDateTo _
          & " and SAAJRNAMJH <= " & wDateFrom _
          & xWhere _
           & " order by SAAJRNAID , SAAJRNAMJH desc , SAAJRNSEQ"


Set rsSab = cnsab.Execute(xSql)

fgSelect_Display_YSAAJRN0

End Sub

Public Function fgSelect_Display_YSAAJRN0_Lib(lK As String) As String

If lK <> arrJrnl_Event_Id(arrJrnl_Event_K) Then
    For arrJrnl_Event_K = 1 To arrJrnl_Event_Nb
        If lK = arrJrnl_Event_Id(arrJrnl_Event_K) Then Exit For
    Next arrJrnl_Event_K
End If
fgSelect_Display_YSAAJRN0_Lib = arrJrnl_Event_Lib(arrJrnl_Event_K)

End Function

Public Function cmdSelect_SQL_YSAAJRN0_rMesg(lSAAJRNTOPX As String, lSAAJRNSUFX As Double)
Dim xSql As String, X As String, wUmid_suffix As Long
Dim wDateFrom, wDateTo

cmdSelect_SQL_YSAAJRN0_rMesg = "?"
X = CStr(lSAAJRNSUFX)

If Len(X) < 6 Then
    wUmid_suffix = lSAAJRNSUFX
    wDateFrom = 1200802192
    wDateTo = 0
Else
    If Len(X) > 6 Then
        wUmid_suffix = Val(Mid$(X, 7, Len(X) - 6))
    Else
        wUmid_suffix = 0
    End If
    X = Mid$(X, 5, 2) & "/" & Mid$(X, 3, 2) & "/20" & Mid$(X, 1, 2)
    
    Call cmdSelect_SQL_rJrnl_Date(X & " 00:00:00", X & " 23:59:59", wDateFrom, wDateTo)
End If

xSql = "select * from rMesg " _
      & "where mesg_uumid = '" & Trim(lSAAJRNTOPX) & "' and mesg_uumid_suffix = " & wUmid_suffix _
      & " and mesg_s_umidl < " & wDateFrom & " and mesg_s_umidl > " & wDateTo _
      & " order by rMesg.Aid , mesg_s_umidl desc ,mesg_s_umidh desc"

      
Set rsSIDE_X = cnSIDE_DB.Execute(xSql)

If rsSIDE_X.EOF Then
    xSql = "select * from rMesg " _
          & "where mesg_uumid = '" & Trim(lSAAJRNTOPX) & "'" _
          & " order by rMesg.Aid , mesg_s_umidl desc ,mesg_s_umidh desc"
    Set rsSIDE_X = cnSIDE_DB.Execute(xSql)
End If

If Not rsSIDE_X.EOF Then

    Mesg_aid = rsSIDE_X("Aid")
    mesg_s_umidl = rsSIDE_X("mesg_s_umidl")
    mesg_s_umidh = rsSIDE_X("mesg_s_umidh")
    cmdSelect_SQL_YSAAJRN0_rMesg = Null
End If


End Function

