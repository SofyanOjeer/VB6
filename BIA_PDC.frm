VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBIA_PDC 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA_PDC"
   ClientHeight    =   12165
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   15615
   Icon            =   "BIA_PDC.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   12165
   ScaleWidth      =   15615
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   8580
      TabIndex        =   19
      Top             =   15
      Width           =   5205
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   11640
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   15510
      _ExtentX        =   27358
      _ExtentY        =   20532
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Position de change"
      TabPicture(0)   =   "BIA_PDC.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "TERME"
      TabPicture(1)   =   "BIA_PDC.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTab1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Paramétrage"
      TabPicture(2)   =   "BIA_PDC.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraPDCOPE"
      Tab(2).Control(1)=   "fraSelect_Options_Y"
      Tab(2).Control(2)=   "fraSelect_Comment_Xls"
      Tab(2).Control(3)=   "fraSelect_Options_xls"
      Tab(2).Control(4)=   "chkSelect_PDCMVTKCUT"
      Tab(2).Control(5)=   "fraPDC_Param"
      Tab(2).Control(6)=   "fraTab2"
      Tab(2).ControlCount=   7
      Begin VB.Frame fraPDCOPE 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   -71520
         TabIndex        =   106
         Top             =   4320
         Visible         =   0   'False
         Width           =   8000
         Begin VB.CommandButton cmdPDCOPE_Update_Ref 
            BackColor       =   &H0080C0FF&
            Caption         =   "Modifier N° ticket"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   6030
            Style           =   1  'Graphical
            TabIndex        =   151
            Top             =   6375
            Width           =   1400
         End
         Begin VB.CommandButton cmdPDCOPE_Quit 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Abandonner"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   2175
            Style           =   1  'Graphical
            TabIndex        =   150
            Top             =   6345
            Width           =   1400
         End
         Begin VB.CommandButton cmdPDCOPE_Update 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Enregistrer"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   4125
            Style           =   1  'Graphical
            TabIndex        =   149
            Top             =   6360
            Width           =   1400
         End
         Begin VB.Frame fraPDCOPE_S 
            BackColor       =   &H00C0FFFF&
            Height          =   3735
            Left            =   0
            TabIndex        =   123
            Top             =   1230
            Width           =   8000
            Begin VB.ComboBox cboPDCOPESER 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1845
               Style           =   2  'Dropdown List
               TabIndex        =   135
               Top             =   225
               Width           =   900
            End
            Begin VB.TextBox txtPDCOPEREF 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4185
               TabIndex        =   134
               Top             =   240
               Width           =   1215
            End
            Begin VB.ComboBox cboPDCOPEOPEC 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1875
               Style           =   2  'Dropdown List
               TabIndex        =   133
               Top             =   750
               Width           =   900
            End
            Begin VB.ComboBox cboPDCOPEOPET 
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
               Left            =   5490
               Style           =   2  'Dropdown List
               TabIndex        =   132
               Top             =   645
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.TextBox txtPDCOPEOPEN 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4155
               TabIndex        =   131
               Top             =   740
               Width           =   1215
            End
            Begin VB.TextBox txtPDCOPECLI 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1860
               TabIndex        =   130
               Top             =   2385
               Width           =   900
            End
            Begin VB.ComboBox cboPDCOPEDEV1 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1860
               Style           =   2  'Dropdown List
               TabIndex        =   129
               Top             =   1335
               Width           =   900
            End
            Begin VB.ComboBox cboPDCOPESENS 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2985
               Style           =   2  'Dropdown List
               TabIndex        =   128
               Top             =   1335
               Width           =   2385
            End
            Begin VB.TextBox txtPDCOPEMTD1 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   5655
               TabIndex        =   127
               Top             =   1305
               Width           =   2052
            End
            Begin VB.ComboBox cboPDCOPEDEV2 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1875
               Style           =   2  'Dropdown List
               TabIndex        =   126
               Top             =   1850
               Width           =   900
            End
            Begin VB.TextBox txtPDCOPEITXT 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   3600
               MaxLength       =   64
               MultiLine       =   -1  'True
               TabIndex        =   125
               Text            =   "BIA_PDC.frx":0496
               Top             =   2880
               Width           =   4215
            End
            Begin VB.CheckBox chkPDCOPEINFO 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0080C0FF&
               Caption         =   "demande pour information ?"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   645
               TabIndex        =   124
               Top             =   3330
               Width           =   2535
            End
            Begin MSComCtl2.DTPicker txtPDCOPEDVA 
               Height          =   300
               Left            =   1860
               TabIndex        =   136
               Top             =   2865
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
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
               Format          =   93257731
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblPDCOPEOPEN 
               BackColor       =   &H00C0FFFF&
               Caption         =   "N° dossier"
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
               Left            =   3045
               TabIndex        =   148
               Top             =   780
               Width           =   930
            End
            Begin VB.Label lblPDCOPESER 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Service"
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
               Left            =   700
               TabIndex        =   147
               Top             =   285
               Width           =   735
            End
            Begin VB.Label lblPDCOPEREF 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Ticket"
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
               Left            =   3030
               TabIndex        =   146
               Top             =   270
               Width           =   615
            End
            Begin VB.Label lblPDCOPEOPEC 
               BackColor       =   &H00C0FFFF&
               Caption         =   "opération"
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
               Left            =   700
               TabIndex        =   145
               Top             =   765
               Width           =   1020
            End
            Begin VB.Label lblPDCOPECLI 
               BackColor       =   &H00C0FFFF&
               Caption         =   "contrepartie"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   645
               TabIndex        =   144
               Top             =   2430
               Width           =   990
            End
            Begin VB.Label libPDCOPECLI 
               BackColor       =   &H00A0FFFF&
               Caption         =   "contrepartie"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3015
               TabIndex        =   143
               Top             =   2415
               Width           =   4215
            End
            Begin VB.Label lblPDCOPEDEV1 
               BackColor       =   &H00FF80FF&
               Caption         =   "devise principale"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   142
               Top             =   1335
               Width           =   1700
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblPDCOPEMTD1 
               BackColor       =   &H00C0FFFF&
               Caption         =   "montant"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6705
               TabIndex        =   141
               Top             =   930
               Width           =   735
            End
            Begin VB.Label lblPDCOPEDEV2 
               BackColor       =   &H00FF80FF&
               Caption         =   "devise secondaire"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   140
               Top             =   1890
               Width           =   1695
               WordWrap        =   -1  'True
            End
            Begin VB.Label libPDCOPEMTD2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00A0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "montant dev secondai"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   5655
               TabIndex        =   139
               Top             =   1850
               Width           =   2055
            End
            Begin VB.Label lblPDCOPEDVA 
               BackColor       =   &H00FF00FF&
               Caption         =   "date valeur"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   645
               TabIndex        =   138
               Top             =   2865
               Width           =   990
            End
            Begin VB.Label libPDCOPETAUX 
               BackColor       =   &H00A0FFFF&
               Caption         =   " -"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   3030
               TabIndex        =   137
               Top             =   1850
               Width           =   2295
            End
         End
         Begin VB.Frame fraPDCOPE_V 
            BackColor       =   &H00C0FFFF&
            Height          =   1350
            Left            =   0
            TabIndex        =   117
            Top             =   4920
            Width           =   8000
            Begin VB.TextBox txtPDCOPEFIXING 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
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
               Left            =   1860
               TabIndex        =   120
               Text            =   "0"
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox txtPDCOPETAUX 
               Alignment       =   1  'Right Justify
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
               Left            =   1860
               TabIndex        =   119
               Text            =   "0"
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtPDCOPEVTXT 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   3615
               MaxLength       =   64
               MultiLine       =   -1  'True
               TabIndex        =   118
               Text            =   "BIA_PDC.frx":049C
               Top             =   225
               Width           =   4215
            End
            Begin VB.Label lblPDCOPETAUX 
               BackColor       =   &H0080C0FF&
               Caption         =   "cours au certain"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   105
               TabIndex        =   122
               Top             =   345
               Width           =   1560
            End
            Begin VB.Label lblPDCOPEFIXING 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Fixing"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   195
               TabIndex        =   121
               Top             =   870
               Width           =   1515
            End
         End
         Begin VB.Frame fraPDCOPE_R 
            BackColor       =   &H00A0FFFF&
            Enabled         =   0   'False
            Height          =   1335
            Left            =   45
            TabIndex        =   108
            Top             =   60
            Width           =   8000
            Begin VB.Label libPDCOPEIUSR 
               BackColor       =   &H00A0FFFF&
               Caption         =   "saisi par"
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
               Left            =   3840
               TabIndex        =   116
               Top             =   600
               Width           =   3645
            End
            Begin VB.Label libPDCOPEVUSR 
               BackColor       =   &H00A0FFFF&
               Caption         =   "validé par"
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
               Left            =   3840
               TabIndex        =   115
               Top             =   960
               Width           =   3840
            End
            Begin VB.Label libPDCOPEID 
               BackColor       =   &H00A0FFFF&
               Caption         =   "ID"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   120
               TabIndex        =   114
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label libPDCOPESTA 
               Appearance      =   0  'Flat
               BackColor       =   &H00A0FFFF&
               Caption         =   "statut"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   345
               Left            =   2490
               TabIndex        =   113
               Top             =   120
               Width           =   4335
            End
            Begin VB.Label lblPDCOPESTA2 
               BackColor       =   &H00A0FFFF&
               Caption         =   "statut 2"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   112
               Top             =   600
               Width           =   2055
            End
            Begin VB.Label lblPDCOPESTA3 
               BackColor       =   &H00A0FFFF&
               Caption         =   "statut 3"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   111
               Top             =   960
               Width           =   2055
            End
            Begin VB.Label lblPDCOPEIUSR 
               BackColor       =   &H00A0FFFF&
               Caption         =   "saisi par"
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
               Left            =   2520
               TabIndex        =   110
               Top             =   600
               Width           =   855
            End
            Begin VB.Label lblPDCOPEVUSR 
               BackColor       =   &H00A0FFFF&
               Caption         =   "mise à jour  par"
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
               Left            =   2520
               TabIndex        =   109
               Top             =   960
               Width           =   1215
            End
         End
         Begin VB.CommandButton cmdPDCOPE_Annulation 
            BackColor       =   &H000000FF&
            Caption         =   "Annuler l'opération"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   107
            Top             =   6400
            Width           =   1400
         End
      End
      Begin VB.Frame fraSelect_Options_Y 
         Caption         =   "Date du duplicata"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -61200
         TabIndex        =   104
         Top             =   2595
         Width           =   2760
         Begin MSComCtl2.DTPicker txtSelect_AMJ_Y 
            Height          =   300
            Left            =   1215
            TabIndex        =   105
            Top             =   315
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
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
            Format          =   93257731
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
      End
      Begin VB.Frame fraSelect_Comment_Xls 
         BackColor       =   &H0080C0FF&
         Caption         =   "Commentaire à inclure dans le mail"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3812
         Left            =   -64860
         TabIndex        =   87
         Top             =   4770
         Width           =   5172
         Begin VB.TextBox txtSelect_Comment_xls 
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3000
            Left            =   360
            MultiLine       =   -1  'True
            TabIndex        =   88
            Top             =   480
            Width           =   4332
         End
      End
      Begin VB.Frame fraSelect_Options_xls 
         Height          =   1230
         Left            =   -74745
         TabIndex        =   77
         Top             =   480
         Visible         =   0   'False
         Width           =   11295
         Begin VB.TextBox txtSelect_Sheet_xls 
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
            TabIndex        =   86
            Text            =   "RECAPEUR"
            Top             =   600
            Width           =   1212
         End
         Begin VB.TextBox txtSelect_File_xls 
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
            TabIndex        =   84
            Top             =   240
            Width           =   3612
         End
         Begin VB.CommandButton cmdSelect_Ok_xls 
            BackColor       =   &H0000FF00&
            Caption         =   "validation du contrôle"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   600
            Width           =   2292
         End
         Begin VB.CheckBox chkSelect_Suspens_Out_xls 
            Caption         =   "Exclure les suspens FOTC"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   225
            Left            =   120
            TabIndex        =   79
            Top             =   600
            Width           =   2412
         End
         Begin VB.CheckBox chkSelect_HB_xls 
            Caption         =   "Exclure HB >"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   120
            TabIndex        =   78
            Top             =   960
            Width           =   1452
         End
         Begin MSComCtl2.DTPicker txtSelect_AMJ_xls 
            Height          =   300
            Left            =   1560
            TabIndex        =   81
            Top             =   240
            Width           =   1332
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
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
            Format          =   93257731
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin MSComCtl2.DTPicker txtSelect_AMJ_HB_xls 
            Height          =   300
            Left            =   1560
            TabIndex        =   98
            Top             =   900
            Width           =   1332
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
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
            Format          =   93257731
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label libSelect_Report_xls 
            BackColor       =   &H0000FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "opé reportées ?"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   7860
            TabIndex        =   89
            Top             =   240
            Width           =   3105
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSelect_Sheet_xls 
            Caption         =   "feuille "
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
            TabIndex        =   85
            Top             =   600
            Width           =   492
         End
         Begin VB.Label lblSelect_File_xls 
            Caption         =   "fichier .xls"
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
            TabIndex        =   83
            Top             =   240
            Width           =   732
         End
         Begin VB.Label lblSelect_AMJ_xls 
            Caption         =   "date"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   252
            Left            =   840
            TabIndex        =   80
            Top             =   240
            Width           =   492
         End
      End
      Begin VB.CheckBox chkSelect_PDCMVTKCUT 
         Alignment       =   1  'Right Justify
         Caption         =   "Exclure mvts comptables non soldés"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   225
         Left            =   -65145
         TabIndex        =   75
         Top             =   1050
         Width           =   3132
      End
      Begin VB.Frame fraPDC_Param 
         Caption         =   "PDC_Param"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   -69765
         TabIndex        =   71
         Top             =   1500
         Width           =   7575
         Begin VB.CommandButton cmdPDC_Param 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Enregistrer"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   5640
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   360
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker txtPDC_Param 
            Height          =   300
            Left            =   3720
            TabIndex        =   73
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
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
            Format          =   93257731
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblPDC_Param 
            Caption         =   "date de modification des schémas comptables TERME"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   495
            Left            =   240
            TabIndex        =   72
            Top             =   480
            Width           =   2655
         End
      End
      Begin VB.Frame fraTab1 
         Height          =   10710
         Left            =   -74895
         TabIndex        =   46
         Top             =   600
         Width           =   15750
         Begin VB.Frame fraReport 
            BackColor       =   &H000000FF&
            Caption         =   "Annulation d'une opération PDC reportée et du report"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   5172
            Left            =   435
            TabIndex        =   91
            Top             =   1635
            Visible         =   0   'False
            Width           =   5412
            Begin VB.CommandButton cmdReport_Quit 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Abandonner"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   94
               Top             =   4320
               Width           =   1695
            End
            Begin VB.CommandButton cmdReport_Update 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Enregistrer"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   612
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   93
               Top             =   4320
               Width           =   1575
            End
            Begin VB.TextBox txtReport_Comment 
               BackColor       =   &H00C0FFC0&
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1200
               Left            =   240
               MultiLine       =   -1  'True
               TabIndex        =   92
               Text            =   "BIA_PDC.frx":04A2
               Top             =   2880
               Width           =   4812
            End
            Begin VB.Label libReport 
               BackColor       =   &H00C0C0FF&
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1452
               Left            =   240
               TabIndex        =   96
               Top             =   480
               Width           =   4812
            End
            Begin VB.Label lblReport_Comment 
               BackColor       =   &H000000FF&
               Caption         =   "Précisez le motif de l'annulation"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   372
               Left            =   1080
               TabIndex        =   95
               Top             =   2280
               Width           =   2892
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgTerme 
            Height          =   4620
            Left            =   150
            TabIndex        =   47
            Top             =   330
            Visible         =   0   'False
            Width           =   14880
            _ExtentX        =   26247
            _ExtentY        =   8149
            _Version        =   393216
            Rows            =   1
            Cols            =   13
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   14745599
            ForeColor       =   8388608
            BackColorFixed  =   8438015
            ForeColorFixed  =   -2147483641
            BackColorSel    =   16776960
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"BIA_PDC.frx":04E2
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
         Begin MSFlexGridLib.MSFlexGrid fgTermeEch 
            Height          =   5205
            Left            =   210
            TabIndex        =   48
            Top             =   5250
            Visible         =   0   'False
            Width           =   14880
            _ExtentX        =   26247
            _ExtentY        =   9181
            _Version        =   393216
            Rows            =   1
            Cols            =   14
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16777215
            ForeColor       =   8388608
            BackColorFixed  =   8438015
            ForeColorFixed  =   0
            BackColorSel    =   12640511
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"BIA_PDC.frx":05CD
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
      Begin VB.Frame fraTab2 
         Height          =   8205
         Left            =   -74760
         TabIndex        =   20
         Top             =   1800
         Visible         =   0   'False
         Width           =   13290
         Begin VB.Frame fraSuspens 
            BackColor       =   &H00FFE0FF&
            Caption         =   "Gestion des suspens"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7005
            Left            =   60
            TabIndex        =   23
            Top             =   5100
            Width           =   9135
            Begin VB.CommandButton cmdSuspens_Quit 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Abandonner"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   2730
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   6255
               Width           =   1695
            End
            Begin VB.CommandButton cmdSuspens_Annulation 
               BackColor       =   &H000000FF&
               Caption         =   "Effacer"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   225
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   6270
               Width           =   1455
            End
            Begin VB.CommandButton cmdSuspens_Update 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Enregistrer"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   612
               Left            =   5265
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   6270
               Width           =   1575
            End
            Begin VB.Frame fraSuspens_S 
               BackColor       =   &H00FFE0FF&
               Height          =   1890
               Left            =   120
               TabIndex        =   33
               Top             =   600
               Width           =   8820
               Begin VB.OptionButton optSuspens_XXX 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "non déterminé"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   70
                  Top             =   1100
                  Value           =   -1  'True
                  Visible         =   0   'False
                  Width           =   1935
               End
               Begin VB.OptionButton optSuspens_XXT 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "terme"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   67
                  Top             =   1400
                  Width           =   1455
               End
               Begin VB.OptionButton optSuspens_XXC 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "comptant"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   66
                  Top             =   800
                  Width           =   1335
               End
               Begin VB.TextBox txtSuspens_PDCMVTMTD1 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   6120
                  TabIndex        =   38
                  Text            =   "MTD1"
                  Top             =   320
                  Width           =   2175
               End
               Begin VB.TextBox txtSuspens_PDCMVTMTD2 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   6150
                  TabIndex        =   37
                  Text            =   "MTD2"
                  Top             =   1365
                  Width           =   2175
               End
               Begin VB.ComboBox cboSuspens_PDCMVTSENS 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   1800
                  Style           =   2  'Dropdown List
                  TabIndex        =   36
                  Top             =   320
                  Width           =   2295
               End
               Begin VB.ComboBox cboSuspens_PDCMVTDEV1 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   240
                  Style           =   2  'Dropdown List
                  TabIndex        =   35
                  Top             =   320
                  Width           =   1200
               End
               Begin VB.ComboBox cboSuspens_PDCMVTDEV2 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   240
                  Style           =   2  'Dropdown List
                  TabIndex        =   34
                  Top             =   1230
                  Width           =   1200
               End
               Begin VB.Label libSuspens_PDCMVTTAUX 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFE0FF&
                  Caption         =   " -"
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
                  Left            =   6100
                  TabIndex        =   39
                  Top             =   930
                  Width           =   2175
               End
            End
            Begin VB.Frame fraSuspens_M 
               BackColor       =   &H00FFE0FF&
               Height          =   1875
               Left            =   120
               TabIndex        =   24
               Top             =   2400
               Width           =   8925
               Begin VB.TextBox txtSuspens_Comment 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   555
                  Left            =   1155
                  MaxLength       =   64
                  TabIndex        =   102
                  Top             =   1200
                  Width           =   6555
               End
               Begin VB.TextBox txtSuspens_PDCMVTCLI 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   1150
                  TabIndex        =   27
                  Top             =   750
                  Width           =   1260
               End
               Begin VB.ComboBox cboSuspens_PDCMVTSER 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   1150
                  Style           =   2  'Dropdown List
                  TabIndex        =   26
                  Top             =   300
                  Width           =   1260
               End
               Begin VB.CommandButton cmdSuspens_Modification 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Modifier"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   612
                  Left            =   7200
                  Style           =   1  'Graphical
                  TabIndex        =   25
                  Top             =   210
                  Width           =   1575
               End
               Begin MSComCtl2.DTPicker txtSuspens_PDCMVTDVA 
                  Height          =   300
                  Left            =   4905
                  TabIndex        =   28
                  Top             =   270
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   529
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
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
                  Format          =   93257731
                  CurrentDate     =   36299
                  MaxDate         =   401768
                  MinDate         =   -328351
               End
               Begin VB.Label libSuspens_Comment 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "commentaire"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   103
                  Top             =   1305
                  Width           =   990
               End
               Begin VB.Label lblSuspens_PDCMVTcli 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "contrepartie"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   32
                  Top             =   825
                  Width           =   855
               End
               Begin VB.Label libSuspens_PDCMVTCLI 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "contrepartie"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   2670
                  TabIndex        =   31
                  Top             =   825
                  Width           =   3615
               End
               Begin VB.Label lblSuspens_PDCMVTSER 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "Service"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   30
                  Top             =   360
                  Width           =   735
               End
               Begin VB.Label lblSuspens_PDCMVTDVA 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "date valeur"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   3570
                  TabIndex        =   29
                  Top             =   330
                  Width           =   1095
               End
            End
            Begin MSFlexGridLib.MSFlexGrid fgSuspens_Log 
               Height          =   1740
               Left            =   120
               TabIndex        =   101
               Top             =   4455
               Width           =   8790
               _ExtentX        =   15505
               _ExtentY        =   3069
               _Version        =   393216
               Rows            =   1
               Cols            =   4
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   14745599
               ForeColor       =   8388608
               BackColorFixed  =   8438015
               ForeColorFixed  =   -2147483641
               BackColorSel    =   16776960
               BackColorBkg    =   16777210
               WordWrap        =   -1  'True
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
               GridLinesFixed  =   1
               AllowUserResizing=   3
               FormatString    =   $"BIA_PDC.frx":06A9
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
            Begin VB.Label libSuspens_PDCMVTDTR 
               Alignment       =   2  'Center
               BackColor       =   &H00FFE0FF&
               Caption         =   "libSuspens_PDCMVTDTR"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1110
               TabIndex        =   43
               Top             =   285
               Width           =   6135
            End
         End
         Begin VB.Frame fraPDCOPE_Options 
            Height          =   885
            Left            =   10290
            TabIndex        =   44
            Top             =   0
            Width           =   3075
         End
         Begin VB.Frame fraSuspens_Options 
            Height          =   885
            Left            =   1560
            TabIndex        =   21
            Top             =   360
            Width           =   5595
            Begin VB.CheckBox chkSuspens_All 
               Alignment       =   1  'Right Justify
               Caption         =   "Afficher l'historique des suspens"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   720
               TabIndex        =   22
               Top             =   360
               Width           =   3015
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgPDCOPE 
            Height          =   5145
            Left            =   1050
            TabIndex        =   45
            Top             =   1185
            Visible         =   0   'False
            Width           =   14880
            _ExtentX        =   26247
            _ExtentY        =   9075
            _Version        =   393216
            Rows            =   1
            Cols            =   12
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   14745599
            ForeColor       =   8388608
            BackColorFixed  =   11599871
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"BIA_PDC.frx":0750
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
         Height          =   11190
         Left            =   120
         TabIndex        =   4
         Top             =   345
         Width           =   15120
         Begin VB.CommandButton cmdPDCOPE_CONF_CALL 
            BackColor       =   &H0000C000&
            Caption         =   "CONF CALL"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   10890
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   100
            Top             =   675
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdPDCOPE_New 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Saisir une opération"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   12195
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   645
            Width           =   1155
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
            Left            =   11055
            Sorted          =   -1  'True
            TabIndex        =   8
            Text            =   "cboSelect_SQL"
            Top             =   255
            Width           =   3750
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   1230
            Left            =   120
            TabIndex        =   6
            Top             =   15
            Width           =   10770
            Begin VB.CheckBox chkSelect_Suspens_Out 
               Alignment       =   1  'Right Justify
               Caption         =   "Exclure les suspens FOTC"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF00FF&
               Height          =   225
               Left            =   2655
               TabIndex        =   76
               Top             =   480
               Width           =   2172
            End
            Begin VB.CheckBox chkSelect_ZCHGOPE0_NC 
               Alignment       =   1  'Right Justify
               Caption         =   "Afficher les opérations SAB non validées"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   8000
               TabIndex        =   65
               Top             =   495
               Width           =   2535
            End
            Begin VB.CheckBox chkSelect_ZCHGOPE0 
               Alignment       =   1  'Right Justify
               Caption         =   "Afficher les opérations SAB"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   8000
               TabIndex        =   64
               Top             =   255
               Width           =   2535
            End
            Begin VB.CheckBox chkSelect_Terme 
               Alignment       =   1  'Right Justify
               Caption         =   "inclure TERME"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3255
               TabIndex        =   49
               Top             =   180
               Width           =   1572
            End
            Begin VB.CheckBox chkSelect_Suspens 
               Alignment       =   1  'Right Justify
               Caption         =   "Afficher les suspens FOTC"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF00FF&
               Height          =   225
               Left            =   5295
               TabIndex        =   18
               Top             =   585
               Width           =   2412
            End
            Begin VB.CheckBox chkSelect_Ope 
               Alignment       =   1  'Right Justify
               Caption         =   "Afficher les tickets saisis"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   5295
               TabIndex        =   17
               Top             =   270
               Width           =   2412
            End
            Begin VB.CheckBox chkSelect_Log 
               Alignment       =   1  'Right Justify
               Caption         =   "Afficher suivi des traitements"
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
               Left            =   5295
               TabIndex        =   15
               Top             =   840
               Width           =   2412
            End
            Begin VB.CheckBox chkSelect_HB 
               Alignment       =   1  'Right Justify
               Caption         =   "Exclure HB >"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   348
               Left            =   2460
               TabIndex        =   14
               Top             =   780
               Width           =   1212
            End
            Begin VB.ComboBox cboSelect_Devise 
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
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   600
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker txtSelect_AMJ 
               Height          =   300
               Left            =   840
               TabIndex        =   9
               Top             =   200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
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
               Format          =   93257731
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtSelect_AMJ_HB 
               Height          =   300
               Left            =   3735
               TabIndex        =   99
               Top             =   825
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "@Arial Unicode MS"
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
               Format          =   93257731
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label libSelect_Report 
               BackColor       =   &H0000FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "opé reportées ?"
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
               Left            =   8000
               TabIndex        =   90
               Top             =   915
               Width           =   2535
               WordWrap        =   -1  'True
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00C0C0C0&
               BorderWidth     =   2
               X1              =   7875
               X2              =   7875
               Y1              =   165
               Y2              =   1245
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00C0C0C0&
               BorderWidth     =   2
               X1              =   5070
               X2              =   5070
               Y1              =   135
               Y2              =   1215
            End
            Begin VB.Label lblSelect_AMJ 
               Caption         =   "date"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   495
            End
            Begin VB.Label lblSelect_Devise 
               Caption         =   "Devise"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   120
               TabIndex        =   12
               Top             =   650
               Width           =   612
            End
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Exécuter la requête"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   13575
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   615
            Width           =   1275
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   4650
            Left            =   105
            TabIndex        =   7
            Top             =   1350
            Visible         =   0   'False
            Width           =   14880
            _ExtentX        =   26247
            _ExtentY        =   8202
            _Version        =   393216
            Rows            =   1
            Cols            =   12
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   14745599
            ForeColor       =   8388608
            BackColorFixed  =   16777088
            ForeColorFixed  =   -2147483641
            BackColorSel    =   14737632
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"BIA_PDC.frx":0840
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
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   5000
            Left            =   75
            TabIndex        =   10
            Top             =   6100
            Visible         =   0   'False
            Width           =   14850
            _ExtentX        =   26194
            _ExtentY        =   8811
            _Version        =   393216
            Rows            =   1
            Cols            =   12
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16777210
            ForeColor       =   8388608
            BackColorFixed  =   13684944
            ForeColorFixed  =   0
            BackColorSel    =   12648384
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"BIA_PDC.frx":092D
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
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
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
      Height          =   500
      Left            =   14280
      Picture         =   "BIA_PDC.frx":0A20
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   -90
      Width           =   500
   End
   Begin VB.Frame fraPrint 
      Caption         =   "options d'impression"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   12150
      TabIndex        =   50
      Top             =   10170
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkPrint_Suspens_Out 
         Caption         =   "Exclure les suspens FOTC"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   840
         TabIndex        =   97
         Top             =   360
         Width           =   2500
      End
      Begin VB.CheckBox chkPrint_PDCOPEDTR 
         Caption         =   "Opérations PDC (date création)"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   69
         Top             =   4800
         Width           =   3345
      End
      Begin VB.CheckBox chkPrint_CHGOPECRE 
         Caption         =   "Opérations SAB (date création)"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   68
         Top             =   5200
         Width           =   3345
      End
      Begin VB.CheckBox chkPrint_ZCHGOPE0 
         Caption         =   "Opérations SAB non comptabilisées"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   63
         Top             =   5600
         Width           =   3345
      End
      Begin VB.CommandButton cmdPrint_Quit 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Abandonner"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   6360
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint_Ok 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Imprimer"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   6360
         Width           =   2532
      End
      Begin VB.CheckBox chkPrint_Comptant 
         Caption         =   "Position comptant"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   59
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2500
      End
      Begin VB.CheckBox chkPrint_Terme_Echéancier 
         Caption         =   "échéancier terme"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   58
         Top             =   2800
         Width           =   2500
      End
      Begin VB.CheckBox chkPrint_PDCMVTKCUT 
         Caption         =   "liste des mvts comptables non soldés"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   57
         Top             =   3600
         Width           =   3105
      End
      Begin VB.CheckBox chkPrint_Suspens 
         Caption         =   "liste des suspens"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   56
         Top             =   3200
         Width           =   2500
      End
      Begin VB.CheckBox chkPrint_YPDCLOG0 
         Caption         =   "historique des traitements"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   55
         Top             =   4400
         Width           =   2500
      End
      Begin VB.CheckBox chkPrint_YPDCMVT0 
         Caption         =   "liste des mouvements comptables"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   54
         Top             =   4000
         Width           =   3345
      End
      Begin VB.CheckBox chkPrint_Exclure_PDCMVTKCUT 
         Caption         =   "Exclure mvts comptables non soldés"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   840
         TabIndex        =   53
         Top             =   1080
         Width           =   3105
      End
      Begin VB.CheckBox chkPrint_Exclure_HB 
         Caption         =   "Exclure les mouvements HB"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   840
         TabIndex        =   52
         Top             =   720
         Width           =   2500
      End
      Begin VB.CheckBox chkPrint_Comptant_Terme 
         Caption         =   "Position comptant + terme"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   51
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2500
      End
      Begin VB.Label libPrint_Etat 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Etats au jj/mm/aaaa"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   480
         TabIndex        =   62
         Top             =   2280
         Width           =   3255
      End
   End
   Begin VB.Label libSelect 
      BackColor       =   &H00FFFED9&
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
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   0
      Width           =   7005
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
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
      Begin VB.Menu mnuSelect_Print_YPDCPOS0 
         Caption         =   "Imprimer l'historique des positions "
      End
      Begin VB.Menu mnuSelect_Print_YPDCMVT0 
         Caption         =   "Imprimer l'historique des positions  +  mouvements"
      End
   End
   Begin VB.Menu mnuPDCMVTKCUT 
      Caption         =   "mnuPDCMVTKCUT"
      Visible         =   0   'False
      Begin VB.Menu mnuPDCMVTKCUT_Update 
         Caption         =   "TOPER ce mvt comptable comme non soldé"
      End
      Begin VB.Menu mnuPDCMVTKCUT_Quit 
         Caption         =   "abandonner"
      End
   End
End
Attribute VB_Name = "frmBIA_PDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'Habilitations :
'==============
'Consulter  : affichage PDC
'Saisir     : saisie des opérations du jour ( par service)
'Valider    : saisie des cours (FOTC => service)
'Compta     : import des mouvements comptables
'Rapprocher : saisie des opérations par BOTC (rappro avec FOTC)
'X_Spécial : annulation, / Reprise
'---------------------------------------------------------

Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String, currentError As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim BIA_PDC_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency
Dim blnAuto As Boolean

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

Dim blnTransaction As Boolean
Dim cmdSelect_SQL_K As String, cmdSelect_SQL_Where As String, cmdSelect_SQL_RA1 As String
'______________________________________________________________________

Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long

Dim xYPDCPOS0 As typeYPDCPOS0, newYPDCPOS0 As typeYPDCPOS0, oldYPDCPOS0 As typeYPDCPOS0
Dim arrYPDCPOS0() As typeYPDCPOS0, arrYPDCPOS0_Nb As Long, arrYPDCPOS0_Max As Long, arrYPDCPOS0_Index As Long
Dim selYPDCPOS0() As typeYPDCPOS0
Dim fixingJ_1() As typeYPDCPOS0
Dim FixingJ_1_AMJ As String
Dim arrPPJ() As Currency

Dim xYPDCMVT0 As typeYPDCMVT0, newYPDCMVT0 As typeYPDCMVT0, oldYPDCMVT0 As typeYPDCMVT0
Dim arrYPDCMVT0() As typeYPDCMVT0, arrYPDCMVT0_Nb As Long, arrYPDCMVT0_Max As Long, arrYPDCMVT0_Index As Long
Dim memoYPDCMVT0 As typeYPDCMVT0

Dim xZSOLDE0 As typeZSOLDE0, newZSOLDE0 As typeZSOLDE0, oldZSOLDE0 As typeZSOLDE0
Dim xZMOUVEA0 As typeZMOUVEA0, newZMOUVEA0 As typeZMOUVEA0, oldZMOUVEA0 As typeZMOUVEA0

Dim arrDev, arrDev_Nb As Integer, blnHB As Boolean
Dim arrDev_Row
Dim blnDeviseU As Boolean

Dim xYPDCLOG0 As typeYPDCLOG0, newYPDCLOG0 As typeYPDCLOG0, oldYPDCLOG0 As typeYPDCLOG0
Dim arrYPDCLOG0() As typeYPDCLOG0, arrYPDCLOG0_Nb As Long, arrYPDCLOG0_Max As Long
Dim fgYPDCLOG0_FormatString As String, mPDCLOGUSEQ As Long
Dim mPDCLOGDTR_Min As String

'___________________________________________________________________________________________

Dim fgPDCOPE_FormatString As String, fgPDCOPE_K As Integer
Dim fgPDCOPE_RowDisplay As Integer, fgPDCOPE_RowClick As Integer, fgPDCOPE_ColClick As Integer
Dim fgPDCOPE_ColorClick As Long, fgPDCOPE_ColorDisplay As Long
Dim fgPDCOPE_Sort1 As Integer, fgPDCOPE_Sort2 As Integer
Dim fgPDCOPE_SortAD As Integer, fgPDCOPE_Sort1_Old As Integer
Dim fgPDCOPE_arrIndex As Integer
Dim blnfgPDCOPE_DisplayLine As Boolean

Dim xYPDCOPE0 As typeYPDCOPE0, newYPDCOPE0 As typeYPDCOPE0, oldYPDCOPE0 As typeYPDCOPE0
Dim arrYPDCOPE0() As typeYPDCOPE0, arrYPDCOPE0_Nb As Long, arrYPDCOPE0_Max As Long, arrYPDCOPE0_Index As Long
Dim selYPDCOPE0() As typeYPDCOPE0, selYPDCOPE0_Nb As Long
Dim memoYPDCOPE0 As typeYPDCOPE0
Dim blnPDCOPEDEVU As Boolean
Dim mCLIENACLI As String, mCLIENARA1 As String, mCLIENARES As String, mMNURUTUTI As String
Dim mFIXING_DEV As String, mFixing_AMJ As String, mFIXING_Cours As Double
Dim mPDCOPEDVA_2J As String, mPDCOPEDVA_5J As String
Dim blnPDCOPE_Control_S As Boolean, blnPDCOPE_Control_V As Boolean
Dim wCellBackColor As Long, mPDCPOSDTR As String, blnPDC_Instant As Boolean
Dim mPDCOPEMTD1 As Currency
Dim mRecipient_FOTC As String, mSQL_Unit As String
Dim mRecipient_BOTC As String
Dim mRecipient_CONF_CALL As String
Dim blnPDCOPE_Control_Ok As Boolean, blnSuspens_Control_Ok As Boolean

Dim xZCHGOPE0 As typeZCHGOPE0

'____________________________________________________________________________________

Dim fgTerme_FormatString As String, fgTerme_K As Integer
Dim fgTerme_RowDisplay As Integer, fgTerme_RowClick As Integer, fgTerme_ColClick As Integer
Dim fgTerme_ColorClick As Long, fgTerme_ColorDisplay As Long
Dim fgTerme_Sort1 As Integer, fgTerme_Sort2 As Integer
Dim fgTerme_SortAD As Integer, fgTerme_Sort1_Old As Integer
Dim fgTerme_arrIndex As Integer
Dim blnfgTerme_DisplayLine As Boolean


Dim fgTermeEch_FormatString As String, fgTermeEch_K As Integer
Dim fgTermeEch_RowDisplay As Integer, fgTermeEch_RowClick As Integer, fgTermeEch_ColClick As Integer
Dim fgTermeEch_ColorClick As Long, fgTermeEch_ColorDisplay As Long
Dim fgTermeEch_Sort1 As Integer, fgTermeEch_Sort2 As Integer
Dim fgTermeEch_SortAD As Integer, fgTermeEch_Sort1_Old As Integer
Dim fgTermeEch_arrIndex As Integer
Dim blnfgTermeEch_DisplayLine As Boolean

Dim arrTerme_DB() As typeYPDCPOS0
Dim arrTerme_CR() As typeYPDCPOS0
Dim arrKCUT() As typeYPDCPOS0
Dim arrSuspens_Dev() As typeYPDCMVT0
Dim arrSWP_Dev() As typeYPDCMVT0

Dim arrZCHGOPE0() As typeZCHGOPE0, arrZCHGOPE0_Max As Long, arrZCHGOPE0_Nb As Long

Dim paramIBM_Library_SABXXX As String
Dim mfgDetail_Top As Integer, mfgDetail_Height As Integer

Dim arrCHGOPEVAL() As String
Dim mTER382100_Amj As String
Dim xYPDCMAIL As typeYPDCMAIL, oldYPDCMAIL As typeYPDCMAIL

Dim BOTC_File_xls As String, BOTC_Sheet_xls As String
Dim mAMJ_xls As String, mAMJ_JP0 As String
Dim mPDCPOSDTR_PNL As String
Dim wAmjMin_HB As String

Dim blnPDCOPE_CONF_CALL_Visible As Boolean, blnPDCOPE_CONF_CALL_Saisie As Boolean

Dim arrDevF_ISO() As String, arrDevF_AMJ() As Long, arrDevF_Nb As Integer, arrDevF_Max As Integer

Dim fgSuspens_Log_FormatString As String

Dim localUnit As String
Public Sub cmdPrint_CHGOPECRE_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim K As Integer, I As Integer
Dim xSQL As String, mTrame As String
Dim X As String
Dim comptageRows As Long
Dim maxRows As Long
Dim maxRowsPlus As Long

    prtBIA_PDC_Etat_Init "Liste des opérations SAB saisies le " & dateImp(wAMJMin)
    wsExcel.Cells(currentRow, 5) = "Liste des opérations SAB saisies le " & dateImp(wAMJMin)
    currentRow = 6
    comptageRows = currentRow
    maxRows = 45
    maxRowsPlus = 3
    cmdSelect_SQL_2_ZCHGOPE0
    For I = 1 To arrZCHGOPE0_Nb
        If arrZCHGOPE0(I).CHGOPEVAL <> "O" Then
            X = "non validé"
        Else
            X = ""
            If arrZCHGOPE0(I).CHGOPESER <> "TC" Then
                If arrZCHGOPE0(I).CHGOPEMDA <> "MAD" Then
                    X = "non compta"
                End If
            End If
        End If
        Call prtBIA_PDCOPE_ZCHGOPE0_xlsManual(arrZCHGOPE0(I), X, currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    Next I
    If currentRow >= maxRows + maxRowsPlus Then
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
            comptageRows = 3
            currentRow = currentRow + 3
        End If
    End If
    comptageRows = comptageRows + 1
    currentRow = currentRow + 1
    Range("A4:K4").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste
    wsExcel.Rows(currentRow).RowHeight = 6
    If currentRow >= maxRows + maxRowsPlus Then
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
            comptageRows = 3
            currentRow = currentRow + 3
        End If
    End If
    comptageRows = comptageRows + 1
    currentRow = currentRow + 1
    Range("A6:J6").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste

End Sub

Public Sub cmdPrint_Ok_Etat_xlsManual()
Dim K As Integer, I As Integer
Dim blnOk As Boolean, xSQL As String
Dim mEtat_PDCPOS As String, mEtat_Exclure As String, mEtat_Exclure_PDCMVTKCUT As String
Dim mEtat_Suspens As String

Dim currentRow As Long
Dim wbExcel2 As Excel.Workbook
Dim ar() As String
Dim ii As Long
Dim nbSheetRows() As Long
Dim nombreFeuilles As Long

blnControl = False
'_________________________________________________
mEtat_Exclure = ""
If chkPrint_Exclure_HB = "1" Then mEtat_Exclure = "(opérations de change en hors-bilan exclues)"
If chkPrint_Suspens_Out = "1" Then mEtat_Exclure = mEtat_Exclure & " (suspens FOTC exclus)"
If chkPrint_Exclure_PDCMVTKCUT = "1" Then
    mEtat_Exclure_PDCMVTKCUT = " (mouvements comptables non soldés exclus)"
Else
    mEtat_Exclure_PDCMVTKCUT = ""
End If
chkSelect_HB = chkPrint_Exclure_HB
chkSelect_Suspens_Out = chkPrint_Suspens_Out
chkSelect_PDCMVTKCUT = chkPrint_Exclure_PDCMVTKCUT
'                                               '
nombreFeuilles = 10
ReDim nbSheetRows(1 To nombreFeuilles)
Call init_xlsManual
'On recopie le classeur modèle de c:\BIASRV vers c:\temp\imp_pdf
FileCopy paramFolder_Local & "\Modeles\modele_PDC.xlsx", paramIMP_PDF_Path_Temp & "\modele_PDC.xlsx"
'on charge CE classeur dans Excel
Call appExcelPublic.Workbooks.Open(paramIMP_PDF_Path_Temp & "\modele_PDC.xlsx")
Set wbExcel2 = appExcelPublic.ActiveWorkbook
With wbExcel2
    .Title = "PDC"
    .Subject = "PDC"
End With
'                                               '
'___________________________________________________________________________
If chkPrint_Comptant = "1" Then
    chkSelect_Terme = "0"
    cmdSelect_SQL_1
    mEtat_PDCPOS = "Position de change comptant au  " & dateImp(wAMJMin)
    Call prtBIA_PDC_Init(mEtat_PDCPOS, mEtat_Exclure & mEtat_Suspens, mEtat_Exclure_PDCMVTKCUT)
    currentRow = 1
    wbExcel2.Sheets("COMPTANT").Activate
    Call prtBIA_PDCPOS_Form_xlsManual(currentRow, wbExcel2.Sheets("COMPTANT"))
    currentRow = 6
    Call prtBIA_PDCPOS_Line_xlsManual(fgSelect, currentRow, wbExcel2.Sheets("COMPTANT"))
    'on supprime les 3 lignes du modèle
    Rows("4:6").Select
    Selection.Delete
    currentRow = currentRow - 3
    nbSheetRows(1) = currentRow
End If
'___________________________________________________________________________
If chkPrint_Comptant_Terme = "1" Then
    chkSelect_Terme = "1"
    cmdSelect_SQL_1
    mEtat_PDCPOS = "Position de change comptant + terme au  " & dateImp(wAMJMin)
    Call prtBIA_PDC_Init(mEtat_PDCPOS, mEtat_Exclure & mEtat_Suspens, mEtat_Exclure_PDCMVTKCUT)
    currentRow = 1
    wbExcel2.Sheets("TERME").Activate
    Call prtBIA_PDCPOS_Form_xlsManual(currentRow, wbExcel2.Sheets("TERME"))
    currentRow = 6
    Call prtBIA_PDCPOS_Line_xlsManual(fgSelect, currentRow, wbExcel2.Sheets("TERME"))
    'on supprime les 3 lignes du modèle
    Rows("4:6").Select
    Selection.Delete
    currentRow = currentRow - 3
    nbSheetRows(2) = currentRow
End If
'___________________________________________________________________________
If chkPrint_Terme_Echéancier = "1" Then
    currentRow = 1
    wbExcel2.Sheets("ECHEANCIER").Activate
    Call prtBIA_PDCTER_Line_xlsManual(fgTermeEch, currentRow, wbExcel2.Sheets("ECHEANCIER"))
    'on supprime les 3 lignes du modèle
    Rows("4:6").Select
    Selection.Delete
    currentRow = currentRow - 3
    nbSheetRows(3) = currentRow
End If
If chkPrint_Suspens = "1" Then
    currentRow = 1
    wbExcel2.Sheets("SUSPENS").Activate
    Call cmdPrint_Suspens_xlsManual(currentRow, wbExcel2.Sheets("SUSPENS"))
    'on supprime les 3 lignes du modèle
    Rows("4:6").Select
    Selection.Delete
    currentRow = currentRow - 3
    nbSheetRows(4) = currentRow
End If
If chkPrint_PDCMVTKCUT = "1" Then
    currentRow = 1
    wbExcel2.Sheets("CPTS").Activate
    Call cmdPrint_PDCMVTKCUT_xlsManual(currentRow, wbExcel2.Sheets("CPTS"))
    'on supprime les 3 lignes du modèle
    Rows("4:6").Select
    Selection.Delete
    currentRow = currentRow - 3
    nbSheetRows(5) = currentRow
End If
If chkPrint_YPDCMVT0 = "1" Then
    currentRow = 1
    wbExcel2.Sheets("MVTS").Activate
    Call cmdPrint_YPDCMVT0_xlsManual(currentRow, wbExcel2.Sheets("MVTS"))
    'on supprime les 4 lignes du modèle
    Rows("4:7").Select
    Selection.Delete
    currentRow = currentRow - 4
    nbSheetRows(6) = currentRow
End If
If chkPrint_YPDCLOG0 = "1" Then
    currentRow = 1
    wbExcel2.Sheets("HISTO").Activate
    Call cmdPrint_YPDCLOG0_xlsManual(currentRow, wbExcel2.Sheets("HISTO"))
    'on supprime les 3 lignes du modèle
    Rows("4:6").Select
    Selection.Delete
    currentRow = currentRow - 3
    nbSheetRows(7) = currentRow
End If
If chkPrint_ZCHGOPE0 = "1" Then
    currentRow = 1
    wbExcel2.Sheets("OPENSAISIE").Activate
    Call cmdPrint_ZCHGOPE0_xlsManual(currentRow, wbExcel2.Sheets("OPENSAISIE"))
    'on supprime les 3 lignes du modèle
    Rows("4:6").Select
    Selection.Delete
    currentRow = currentRow - 3
    nbSheetRows(8) = currentRow
End If
If chkPrint_CHGOPECRE = "1" Then
    currentRow = 1
    wbExcel2.Sheets("OPESAISIE").Activate
    Call cmdPrint_CHGOPECRE_xlsManual(currentRow, wbExcel2.Sheets("OPESAISIE"))
    'on supprime les 3 lignes du modèle
    Rows("4:6").Select
    Selection.Delete
    currentRow = currentRow - 3
    nbSheetRows(9) = currentRow
End If
If chkPrint_PDCOPEDTR = "1" Then
    currentRow = 1
    wbExcel2.Sheets("PDCSAISIE").Activate
    Call cmdPrint_PDCOPEDTR_xlsManual(currentRow, wbExcel2.Sheets("PDCSAISIE"))
    'on supprime les 3 lignes du modèle
    Rows("4:6").Select
    Selection.Delete
    currentRow = currentRow - 3
    nbSheetRows(10) = currentRow
End If
'_________________________________________________
ReDim ar(1 To nombreFeuilles)
For ii = 1 To nombreFeuilles
    Call zoneImpression_xlsManual(wbExcel2.Sheets(ii).Name, nbSheetRows(ii), wbExcel2.Sheets(ii))
    ar(ii) = wbExcel2.Sheets(ii).Name
Next ii
wbExcel2.Sheets(ar).Select
Call ActiveSheet.ExportAsFixedFormat(xlTypePDF, paramIMP_PDF_Path_Temp & "\" & paramEditionNoPaper_Auto_PgmName & ".pdf")
Call impressions_xlsManual.prtIMP_PDF_Monitor_xlsManual
Call wbExcel2.Close(False)
Set wbExcel2 = Nothing
Kill paramIMP_PDF_Path_Temp & "\modele_PDC.xlsx"
blnControl = True
End Sub
Public Sub cmdPrint_PDCMVTKCUT_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim K As Integer, I As Integer
Dim xSQL As String, mTrame As String
Dim wKCUT As String
Dim comptageRows As Long
Dim maxRows As Long
Dim maxRowsPlus As Long

prtBIA_PDC_Etat_Init "Liste des mouvements comptables non soldés"
wsExcel.Cells(currentRow, 4) = "Liste des mouvements comptables non soldés"
currentRow = 6
comptageRows = currentRow
maxRows = 45
maxRowsPlus = 3
cmdSelect_SQL_1PDCMVTKCUT

For I = 1 To arrYPDCMVT0_Nb
    For K = 1 To arrDev_Nb
        If arrKCUT(K).PDCPOSDEV = xYPDCMVT0.PDCMVTDEV Then
            wKCUT = "cut " & arrKCUT(K).PDCPOSPRIX
            Exit For
        End If
    Next K
    Call prtBIA_PDCMVT_Line_xlsManual(xYPDCMVT0, wKCUT, currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
Next I
If currentRow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
        comptageRows = 3
        currentRow = currentRow + 3
    End If
End If
comptageRows = comptageRows + 1
currentRow = currentRow + 1
Range("A4:K4").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
wsExcel.Rows(currentRow).RowHeight = 6
If currentRow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
        comptageRows = 3
        currentRow = currentRow + 3
    End If
End If
comptageRows = comptageRows + 1
currentRow = currentRow + 1
Range("A6:I6").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste

End Sub

Public Sub cmdPrint_PDCOPEDTR_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim K As Integer, I As Integer
Dim xSQL As String, mTrame As String
Dim comptageRows As Long
Dim maxRows As Long
Dim maxRowsPlus As Long

    prtBIA_PDC_Etat_Init "Liste des opérations BIA_PDC saisies le " & dateImp(wAMJMin)
    wsExcel.Cells(currentRow, 4) = "Liste des opérations BIA_PDC saisies le " & dateImp(wAMJMin)
    currentRow = 6
    comptageRows = currentRow
    maxRows = 45
    maxRowsPlus = 3
    cmdSelect_SQL_Where = "where PDCOPEIAMJ = '" & wAMJMin & "' or PDCOPEDTR = '" & wAMJMin & "'"
    Call arrYPDCOPE0_SQL(cmdSelect_SQL_Where)
    For I = 1 To arrYPDCOPE0_Nb
        Call prtBIA_PDCOPE_YPDCOPE0_xlsManual(arrYPDCOPE0(I), currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    Next I
    If currentRow >= maxRows + maxRowsPlus Then
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
            comptageRows = 3
            currentRow = currentRow + 3
        End If
    End If
    comptageRows = comptageRows + 1
    currentRow = currentRow + 1
    Range("A4:K4").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste
    wsExcel.Rows(currentRow).RowHeight = 6
    If currentRow >= maxRows + maxRowsPlus Then
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
            comptageRows = 3
            currentRow = currentRow + 3
        End If
    End If
    comptageRows = comptageRows + 1
    currentRow = currentRow + 1
    Range("A6:J6").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste

End Sub

Public Sub cmdPrint_Suspens_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim K As Integer, I As Integer
Dim xSQL As String, mTrame As String
Dim comptageRows As Long
Dim maxRows As Long
Dim maxRowsPlus As Long

prtBIA_PDC_Etat_Init "Liste des suspens TC"
wsExcel.Cells(currentRow, 4) = "Liste des suspens TC"
currentRow = 6
comptageRows = currentRow
maxRows = 45
maxRowsPlus = 3
xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCMVT0 where PDCMVTOPEC like 'XX%' and PDCMVTSTA2 = ' ' order by PDCMVTDEV"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    V = rsYPDCMVT0_GetBuffer(rsSab, xYPDCMVT0)
    Call prtBIA_PDCMVT_Line_xlsManual(xYPDCMVT0, "", currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    rsSab.MoveNext
Loop
If currentRow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
        comptageRows = 3
        currentRow = currentRow + 3
    End If
End If
comptageRows = comptageRows + 1
currentRow = currentRow + 1
Range("A4:K4").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
wsExcel.Rows(currentRow).RowHeight = 6
If currentRow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
        comptageRows = 3
        currentRow = currentRow + 3
    End If
End If
comptageRows = comptageRows + 1
currentRow = currentRow + 1
Range("A6:I6").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste

End Sub

Public Sub cmdPrint_YPDCLOG0_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim K As Integer, I As Integer
Dim xSQL As String, mTrame As String
Dim comptageRows As Long
Dim maxRows As Long
Dim maxRowsPlus As Long

prtBIA_PDC_Etat_Init "Historique des traitements de la journée comptable du " & dateImp(wAMJMin)
wsExcel.Cells(currentRow, 3) = "Historique des traitements de la journée comptable du " & dateImp(wAMJMin)
currentRow = 6
comptageRows = currentRow
maxRows = 45
maxRowsPlus = 3
xSQL = "where PDCLOGDTR = '" & wAMJMin & "'"
Call arrYPDCLOG0_SQL(xSQL)
For I = 1 To arrYPDCLOG0_Nb
    Call prtBIA_PDCLOG_Line_xlsManual(arrYPDCLOG0(I), currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
Next I
If currentRow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
        comptageRows = 3
        currentRow = currentRow + 3
    End If
End If
comptageRows = comptageRows + 1
currentRow = currentRow + 1
Range("A4:K4").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
wsExcel.Rows(currentRow).RowHeight = 6
If currentRow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
        comptageRows = 3
        currentRow = currentRow + 3
    End If
End If
comptageRows = comptageRows + 1
currentRow = currentRow + 1
Range("A6:F6").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste

End Sub

Public Sub cmdPrint_YPDCMVT0_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim K As Integer, I As Integer
Dim xSQL As String, mTrame As String
Dim comptageRows As Long
Dim maxRows As Long
Dim maxRowsPlus As Long

prtBIA_PDC_Etat_Init "détail des mouvements"
wsExcel.Cells(currentRow, 5) = "Détail des mouvements"
currentRow = 7
maxRows = 45
maxRowsPlus = 4
For I = 1 To arrYPDCPOS0_Nb
    newYPDCPOS0 = arrYPDCPOS0(I)
    xSQL = "where PDCMVTDTR = '" & newYPDCPOS0.PDCPOSDTR & "' and PDCMVTDEV = '" & newYPDCPOS0.PDCPOSDEV & "'"
    Call arrYPDCMVT0_SQL(xSQL)
    mTrame = " "
    If arrYPDCMVT0_Nb > 0 Then
        mTrame = "B"
        oldYPDCPOS0 = newYPDCPOS0
        oldYPDCPOS0.PDCPOSDTR = DateComptablePrecedente(newYPDCPOS0.PDCPOSDTR)
        If oldYPDCPOS0.PDCPOSDTR <> "00000000" Then
            xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCPOS0 " _
                 & " where PDCPOSDTR ='" & oldYPDCPOS0.PDCPOSDTR & "'" _
                 & " and    PDCPOSDEV ='" & oldYPDCPOS0.PDCPOSDEV & "'"
            Set rsSab = cnsab.Execute(xSQL)
            If Not rsSab.EOF Then
                V = rsYPDCPOS0_GetBuffer(rsSab, oldYPDCPOS0)
                Call prtBIA_PDCMVT_POS_xlsManual(oldYPDCPOS0, "0", currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
            End If
        End If
        For K = 1 To arrYPDCMVT0_Nb
            Call prtBIA_PDCMVT_Line_xlsManual(arrYPDCMVT0(K), "", currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
        Next K
    End If
    Call prtBIA_PDCMVT_POS_xlsManual(newYPDCPOS0, mTrame, currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
Next I
If currentRow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
        comptageRows = 3
        currentRow = currentRow + 3
    End If
End If
comptageRows = comptageRows + 1
currentRow = currentRow + 1
Range("A4:K4").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
wsExcel.Rows(currentRow).RowHeight = 6
If currentRow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
        comptageRows = 3
        currentRow = currentRow + 3
    End If
End If
comptageRows = comptageRows + 1
currentRow = currentRow + 1
Range("A6:I6").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste

End Sub

Public Sub cmdPrint_ZCHGOPE0_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim K As Integer, I As Integer
Dim xSQL As String, mTrame As String
Dim comptageRows As Long
Dim maxRows As Long
Dim maxRowsPlus As Long

    prtBIA_PDC_Etat_Init "Liste des opérations SAB saisies et non validées"
    wsExcel.Cells(currentRow, 5) = "Liste des opérations SAB saisies et non validées"
    currentRow = 6
    comptageRows = currentRow
    maxRows = 45
    maxRowsPlus = 3
    xSQL = "where CHGOPEDE2 <> '   ' and CHGOPEVAL = '1' and CHGOPEANN = ' '"
    Call arrZCHGOPE0_SQL(xSQL)
    For I = 1 To arrZCHGOPE0_Nb
        Call prtBIA_PDCOPE_ZCHGOPE0_xlsManual(arrZCHGOPE0(I), "non validée", currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    Next I
    If currentRow >= maxRows + maxRowsPlus Then
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
            comptageRows = 3
            currentRow = currentRow + 3
        End If
    End If
    comptageRows = comptageRows + 1
    currentRow = currentRow + 1
    Range("A4:K4").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste
    wsExcel.Rows(currentRow).RowHeight = 6
    If currentRow >= maxRows + maxRowsPlus Then
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
            comptageRows = 3
            currentRow = currentRow + 3
        End If
    End If
    comptageRows = comptageRows + 1
    currentRow = currentRow + 1
    Range("A6:J6").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste

End Sub

Public Sub zoneImpression_xlsManual(lFct As String, nbRows As Long, wsheet As Excel.Worksheet)

    Call init_TypePagesetup
    If nbRows > 0 Then
        If Trim(lFct) = "COMPTANT" Then
            wsheet.Activate
            zoneImpressionPagesetup.Zoom = 95
            wsheet.Range("A1:K" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$K$" & CStr(nbRows)
        ElseIf Trim(lFct) = "TERME" Then
            wsheet.Activate
            zoneImpressionPagesetup.Zoom = 90
            wsheet.Range("A1:K" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$K$" & CStr(nbRows)
        ElseIf Trim(lFct) = "ECHEANCIER" Then
            wsheet.Activate
            zoneImpressionPagesetup.Zoom = 90
            wsheet.Range("A1:K" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$K$" & CStr(nbRows)
        ElseIf Trim(lFct) = "SUSPENS" Then
            wsheet.Activate
            zoneImpressionPagesetup.Zoom = 95
            wsheet.Range("A1:I" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$I$" & CStr(nbRows)
        ElseIf Trim(lFct) = "CPTS" Then
            wsheet.Activate
            zoneImpressionPagesetup.Zoom = 95
            wsheet.Range("A1:I" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$I$" & CStr(nbRows)
        ElseIf Trim(lFct) = "MVTS" Then
            wsheet.Activate
            zoneImpressionPagesetup.Zoom = 95
            wsheet.Range("A1:I" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$I$" & CStr(nbRows)
        ElseIf Trim(lFct) = "HISTO" Then
            wsheet.Activate
            zoneImpressionPagesetup.Zoom = 95
            wsheet.Range("A1:F" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$F$" & CStr(nbRows)
        ElseIf Trim(lFct) = "OPENSAISIE" Then
            wsheet.Activate
            zoneImpressionPagesetup.Zoom = 95
            wsheet.Range("A1:J" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$J$" & CStr(nbRows)
        ElseIf Trim(lFct) = "OPESAISIE" Then
            wsheet.Activate
            zoneImpressionPagesetup.Zoom = 95
            wsheet.Range("A1:J" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$J$" & CStr(nbRows)
        ElseIf Trim(lFct) = "PDCSAISIE" Then
            wsheet.Activate
            zoneImpressionPagesetup.Zoom = 95
            wsheet.Range("A1:J" & CStr(nbRows)).Select
            zoneImpressionPagesetup.PrintArea = "$A$1:$J$" & CStr(nbRows)
        End If
    End If
    wsheet.Activate
    zoneImpressionPagesetup.RightFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "&P"
    zoneImpressionPagesetup.Orientation = xlLandscape
    zoneImpressionPagesetup.LeftFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "prtBIA_PDC   &D &T  BIA_INFO"
    Call SetTypePageSetup(wsheet)
    
End Sub

Public Sub cmdSendMail_BIA_PDC_xlsManual()
Dim wSendMail As typeSendMail
Dim xDétail As String, xHeader As String, mbgColor As String
Dim K As Long, htmlFontColor_K As String
'On Error Resume Next

blnControl = False
chkSelect_Suspens_Out.Value = "0"

'$jpl 2014-10-10 Printer_PDF
Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S54", "BIA-PDC", "Archive")

chkPrint_Comptant = "1"
chkPrint_Comptant_Terme = "1"
chkPrint_Terme_Echéancier = "1"
chkPrint_Suspens = "1"
chkPrint_PDCMVTKCUT = "1"
chkPrint_YPDCMVT0 = "1"
chkPrint_YPDCLOG0 = "1"
chkPrint_ZCHGOPE0 = "1"
chkPrint_CHGOPECRE = "1"
chkPrint_PDCOPEDTR = "1"

Call cmdPrint_Ok_Etat_xlsManual

fgYPDCLOG0_FormatString = "<Date Compta   |<mise à jour le                               |<Nature|<libellé                                                                                                                          |<Pièce                  |<Màj par                |||||"

xHeader = "<TR>" _
         & "<TD bgcolor=#0090A0 width=100 height=5><span style='font-size:9.0pt;font-family:Calibri'><Font color=#FFFFFF>Date Compta</TD>" _
         & "<TD bgcolor=#0090A0 width=50 height=5><span style='font-size:9.0pt;font-family:Calibri'><Font color=#FFFFFF>Nature</B></TD>" _
         & "<TD bgcolor=#0090A0 width=550 height=5><span style='font-size:9.0pt;font-family:Calibri'><Font color=#FFFFFF>libellé</TD>" _
         & "<TD bgcolor=#0090A0 width=300 height=5><span style='font-size:9.0pt;font-family:Calibri'><Font color=#FFFFFF>mise à jour le</TD>" _
        & "</TR>"

xDétail = ""
mbgColor = "bgcolor = #FAFAD2"
For K = 1 To arrYPDCLOG0_Nb
    htmlFontColor_K = htmlFontColor_Blue
    If Mid$(arrYPDCLOG0(K).PDCLOGNAT, 3, 1) <> " " Then
        htmlFontColor_K = htmlFontColor_Red
    End If
    If Mid$(arrYPDCLOG0(K).PDCLOGNAT, 1, 1) = "7" Then
        htmlFontColor_K = htmlFontColor_Red
    End If

    xDétail = xDétail _
         & "<TR>" _
         & "<TD " & mbgColor & " width=100 height=5><span style='font-size:8.0pt;font-family:Calibri'>" & htmlFontColor_K & dateImp10(arrYPDCLOG0(K).PDCLOGDTR) & "</TD>" _
         & "<TD " & mbgColor & " width=50 height=5><span style='font-size:8.0pt;font-family:Calibri'>" & htmlFontColor_K & arrYPDCLOG0(K).PDCLOGNAT & "</TD>" _
         & "<TD " & mbgColor & " width=550 height=5><span style='font-size:8.0pt;font-family:Calibri'>" & htmlFontColor_K & arrYPDCLOG0(K).PDCLOGTXT & "</TD>" _
         & "<TD " & mbgColor & " width=300 height=5><span style='font-size:8.0pt;font-family:Calibri'>" & htmlFontColor_K & dateImp10(arrYPDCLOG0(K).PDCLOGUAMJ) & "  " & arrYPDCLOG0(K).PDCLOGUHMS & "  " & arrYPDCLOG0(K).PDCLOGUUSR & "</TD>" _
         & "</TR>"

Next K

wSendMail.FromDisplayName = "@BIA_PDC"
wSendMail.RecipientDisplayName = "BIA_PDC"
Call DTPicker_Control(txtSelect_AMJ, wAMJMin)

wSendMail.Subject = "Calcul de la position de change au : " & dateImp10(wAMJMin) & " (cf. pièce jointe)"
wSendMail.Attachment = "" ' prtIMP_PDF_FileName

wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                    & paramEditionNoPaper_Auto_Lnk _
                    & "<TABLE   width=1000 border=1 cellpadding=4 ></B>" _
                    & xHeader _
                    & xDétail _
                    & "</TABLE>"

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail
End Sub

Public Sub DevF_Load()
Dim X As String, xSQL As String
ReDim arrDevF_ISO(100), arrDevF_AMJ(100)

arrDevF_Max = 100
arrDevF_Nb = 0
xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 where BASTABNUM = 36 order by BASTABARG"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    X = rsSab("BASTABARG")
    wAMJMin = 19000000 + convX2P(Mid$(X, 6, 4))
    If wAMJMin > YBIATAB0_DATE_CPT_J Then
        arrDevF_Nb = arrDevF_Nb + 1
        If arrDevF_Nb > arrDevF_Max Then
            arrDevF_Max = arrDevF_Max + 100
            ReDim Preserve arrDevF_ISO(arrDevF_Max), arrDevF_AMJ(arrDevF_Max)
        End If
        arrDevF_ISO(arrDevF_Nb) = Mid$(X, 3, 3)
        arrDevF_AMJ(arrDevF_Nb) = wAMJMin
    End If
    rsSab.MoveNext
Loop
End Sub

Public Sub cmdSendMail_BIA_PDC()
Dim wSendMail As typeSendMail
Dim xDétail As String, xHeader As String, mbgColor As String
Dim K As Long, htmlFontColor_K As String
On Error Resume Next

blnControl = False
chkSelect_Suspens_Out.Value = "0"

'$jpl 2014-10-10 Printer_PDF
Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S54", "BIA-PDC", "Archive")

chkPrint_Comptant = "1"
chkPrint_Comptant_Terme = "1"
chkPrint_Terme_Echéancier = "1"
chkPrint_Suspens = "1"
chkPrint_PDCMVTKCUT = "1"
chkPrint_YPDCMVT0 = "1"
chkPrint_YPDCLOG0 = "1"
chkPrint_ZCHGOPE0 = "1"
chkPrint_CHGOPECRE = "1"
chkPrint_PDCOPEDTR = "1"

cmdPrint_Ok_Etat

fgYPDCLOG0_FormatString = "<Date Compta   |<mise à jour le                               |<Nature|<libellé                                                                                                                          |<Pièce                  |<Màj par                |||||"

xHeader = "<TR>" _
         & "<TD bgcolor=#0090A0 width=100 height=5><span style='font-size:9.0pt;font-family:Calibri'><Font color=#FFFFFF>Date Compta</TD>" _
         & "<TD bgcolor=#0090A0 width=50 height=5><span style='font-size:9.0pt;font-family:Calibri'><Font color=#FFFFFF>Nature</B></TD>" _
         & "<TD bgcolor=#0090A0 width=550 height=5><span style='font-size:9.0pt;font-family:Calibri'><Font color=#FFFFFF>libellé</TD>" _
         & "<TD bgcolor=#0090A0 width=300 height=5><span style='font-size:9.0pt;font-family:Calibri'><Font color=#FFFFFF>mise à jour le</TD>" _
        & "</TR>"

xDétail = ""
mbgColor = "bgcolor = #FAFAD2"
For K = 1 To arrYPDCLOG0_Nb
    htmlFontColor_K = htmlFontColor_Blue
    If Mid$(arrYPDCLOG0(K).PDCLOGNAT, 3, 1) <> " " Then
        htmlFontColor_K = htmlFontColor_Red
    End If
    If Mid$(arrYPDCLOG0(K).PDCLOGNAT, 1, 1) = "7" Then
        htmlFontColor_K = htmlFontColor_Red
    End If

    xDétail = xDétail _
         & "<TR>" _
         & "<TD " & mbgColor & " width=100 height=5><span style='font-size:8.0pt;font-family:Calibri'>" & htmlFontColor_K & dateImp10(arrYPDCLOG0(K).PDCLOGDTR) & "</TD>" _
         & "<TD " & mbgColor & " width=50 height=5><span style='font-size:8.0pt;font-family:Calibri'>" & htmlFontColor_K & arrYPDCLOG0(K).PDCLOGNAT & "</TD>" _
         & "<TD " & mbgColor & " width=550 height=5><span style='font-size:8.0pt;font-family:Calibri'>" & htmlFontColor_K & arrYPDCLOG0(K).PDCLOGTXT & "</TD>" _
         & "<TD " & mbgColor & " width=300 height=5><span style='font-size:8.0pt;font-family:Calibri'>" & htmlFontColor_K & dateImp10(arrYPDCLOG0(K).PDCLOGUAMJ) & "  " & arrYPDCLOG0(K).PDCLOGUHMS & "  " & arrYPDCLOG0(K).PDCLOGUUSR & "</TD>" _
         & "</TR>"

Next K

wSendMail.FromDisplayName = "@BIA_PDC"
wSendMail.RecipientDisplayName = "BIA_PDC"
Call DTPicker_Control(txtSelect_AMJ, wAMJMin)

wSendMail.Subject = "Calcul de la position de change au : " & dateImp10(wAMJMin) & " (cf. pièce jointe)"
wSendMail.Attachment = "" ' prtIMP_PDF_FileName

wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                    & paramEditionNoPaper_Auto_Lnk _
                    & "<TABLE   width=1000 border=1 cellpadding=4 ></B>" _
                    & xHeader _
                    & xDétail _
                    & "</TABLE>"

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub
Public Sub cmdSendMail_xls()
Dim wSendMail As typeSendMail
Dim xDétail As String, xHeader As String, mbgColor As String
Dim K As Long, htmlFontColor_K As String
Dim iRow As Integer, iCol As Integer, X As String, xTD As String
Dim wForecolor As String, wBackColor As String, xColor As String
Dim xFOTC As String
On Error Resume Next


xHeader = ""
For iRow = 0 To fgSelect.Rows - 1
    fgSelect.Row = iRow
    xTD = ""
    For iCol = 0 To 9
        If iCol <> 3 Then 'ignorer le prix de position
            fgSelect.Col = iCol
            X = Trim(fgSelect.Text)
            If iRow = 0 Then
                wForecolor = cmdSendMail_xls_Color(fgSelect.ForeColorFixed)
                wBackColor = cmdSendMail_xls_Color(fgSelect.BackColorFixed)
                xTD = xTD _
                     & "<TD bgcolor=" & wBackColor & " width=270px><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=" & wForecolor & "><B>" _
                     & X & "</B/TD>"
            Else
                If fgSelect.CellForeColor <> 0 Then
                    wForecolor = cmdSendMail_xls_Color(fgSelect.CellForeColor)
                Else
                    wForecolor = cmdSendMail_xls_Color(fgSelect.ForeColor)
                End If
                
                If fgSelect.CellBackColor <> 0 Then
                    wBackColor = cmdSendMail_xls_Color(fgSelect.CellBackColor)
                Else
                    wBackColor = cmdSendMail_xls_Color(fgSelect.BackColor)
                End If
                
                xTD = xTD _
                     & "<TD bgcolor=" & wBackColor & " width=270px><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=" & wForecolor & ">" _
                     & X & "</TD>"
            End If
        End If
    Next iCol
    xHeader = xHeader & "<TR>" & xTD & "</TR>"

Next iRow

xDétail = ""
For iRow = 0 To fgDetail.Rows - 1
    fgDetail.Row = iRow
    xTD = ""
    For iCol = 0 To 8
        If iCol <> 3 Then 'ignorer le prix de position
         fgDetail.Col = iCol
         X = fgDetail.Text
         If iRow = 0 Then
                wForecolor = cmdSendMail_xls_Color(fgDetail.ForeColorFixed)
                wBackColor = cmdSendMail_xls_Color(fgDetail.BackColorFixed)
              xTD = xTD _
                  & "<TD bgcolor=" & wBackColor & " width=270px><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=" & wForecolor & "><B>" _
                  & X & "</B/TD>"
        Else
                If fgDetail.CellForeColor <> 0 Then
                    wForecolor = cmdSendMail_xls_Color(fgDetail.CellForeColor)
                Else
                    wForecolor = cmdSendMail_xls_Color(fgDetail.ForeColor)
                End If
                
                If fgDetail.CellBackColor <> 0 Then
                    wBackColor = cmdSendMail_xls_Color(fgDetail.CellBackColor)
                Else
                    wBackColor = cmdSendMail_xls_Color(fgDetail.BackColor)
                End If
                
                
             xTD = xTD _
                  & "<TD bgcolor=" & wBackColor & " width=270px><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=" & wForecolor & ">" _
                  & X & "</TD>"
         End If
        End If
    Next iCol
    xDétail = xDétail & "<TR >" & xTD & "</TR>"

Next iRow
mbgColor = "bgcolor = #E0E0E0"

wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = "BOTC_xls"
wSendMail.RecipientDisplayName = "BIA_PDC"

wSendMail.Subject = "Contrôle de la position de change  PDC / BOTC.xls au " & dateImp10(mAMJ_xls)
wSendMail.Attachment = ""
X = Replace("<U>" & Trim(libSelect_Report_xls) & "</U>" & Trim(txtSelect_Comment_xls), vbCr, "<BR>")
xFOTC = ""
If chkSelect_Suspens_Out_xls.Value = "1" Then xFOTC = xFOTC & " - (hors suspens FOTC)"
If chkSelect_HB_xls.Value = "1" Then xFOTC = xFOTC & " - (hors-bilan exclus)"

wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                    & "<span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & "<Font color = #404040>" _
                    & "Veuillez trouver ci-après le contrôle des tableaux récapitulatifs de la position de change au " & dateImp10(mAMJ_xls) _
                    & "<BR>" & htmlFontColor_Blue & X _
                    & "<BR><Font color = #303030>  1 - <U>calculé par l'application 'BIA_PDC'</U>" & htmlFontColor_Gray & xFOTC _
                    & "<BR><BR>" _
                    & "<TABLE   width=2430px border=1 cellpadding=9 ></B>" _
                    & "<div align=" & Asc34 & "right" & Asc34 _
                    & xHeader _
                    & "</div></TABLE>" _
                    & "<BR>" _
                    & "<BR><Font color = #303030>  2 - <U>établi par le service BOTC </U>" _
                    & "<span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Gray & "(" & Trim(txtSelect_File_xls) & " | " & Trim(txtSelect_Sheet_xls) _
                    & " - " & usrName_UCase & " - " & dateImp10(DSys) & " " & Time & ")." _
                    & "<BR><BR>" _
                    & "<TABLE   width=2430px border=1 cellpadding=8 ></B>" _
                    & "<div align=" & Asc34 & "right" & Asc34 _
                    & xDétail _
                    & "</div></TABLE>"

wSendMail.AsHTML = True
xYPDCMAIL.PDCMAILDTR = mAMJ_xls
xYPDCMAIL.PDCMAILTXT = wSendMail.Message
V = sqlYPDCMAIL_Insert(xYPDCMAIL)
If Not IsNull(V) Then Error_Route (V)


srvSendMail.Monitor wSendMail

End Sub
Public Sub cmdSendMail_xls_Duplicata()
Dim wSendMail As typeSendMail
On Error Resume Next


wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = "BOTC_xls"
wSendMail.Recipient = currentSSIWINMAIL

wSendMail.Subject = "DUPLICATA : Contrôle de la position de change  PDC / BOTC.xls au " & dateImp10(oldYPDCMAIL.PDCMAILDTR)
wSendMail.Attachment = ""

wSendMail.AsHTML = True
wSendMail.Message = oldYPDCMAIL.PDCMAILTXT
srvSendMail.Monitor wSendMail

End Sub


Public Sub cmdSendMail_Cours()
Dim wSendMail As typeSendMail
Dim xHeader As String, xDétail As String, mbgColor As String
Dim wSubject As String
Dim K As Long
Dim xSQL As String

On Error Resume Next

If newYPDCOPE0.PDCOPESER = "TC" Then Exit Sub

If Trim(currentSSIWINMAIL) = "" Then
    Call MsgBox("Vous n'avez pas d'adresse mail enregistrée dans SAB", vbCritical, "BIA_PDC : gestion des opérations")
    Exit Sub
End If
mbgColor = "bgcolor = #FFFFFF"
wSendMail.Subject = "PDC " & dateImp10(newYPDCOPE0.PDCOPEDTR) & " - " & newYPDCOPE0.PDCOPEID & " : "
'wSendMail.FromDisplayName = "COURS"
'wSendMail.RecipientDisplayName = "BIA_PDC"
wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = "BIA_PDC - " & usrName_UCase
wSendMail.CcRecipient = srvSendMail.Exchange_Distribution("BIA_PDC", newYPDCOPE0.PDCOPESSE)

Select Case newYPDCOPE0.PDCOPESTA
    Case "0": wSubject = "Demande de cours": mbgColor = "bgcolor = #FAFAD2"
              wSendMail.Recipient = mRecipient_FOTC
    Case "V": wSubject = " Réponse à demande de cours": mbgColor = "bgcolor = #D0FFD0"
              wSendMail.Recipient = mailAdresse_Production(newYPDCOPE0.PDCOPEIUSR) & ";" & mRecipient_BOTC & ";" & mRecipient_FOTC
    Case "A": wSubject = " Annulation d'une demande de cours": mbgColor = "bgcolor = #FFD0D0"
              wSendMail.Recipient = mailAdresse_Production(newYPDCOPE0.PDCOPEIUSR) & ";" & mRecipient_FOTC
              If newYPDCOPE0.PDCOPETAUX <> 0 Then wSendMail.Recipient = wSendMail.Recipient & ";" & mRecipient_BOTC
              
    Case "I": wSubject = " Demande de cours pour information": mbgColor = "bgcolor = #E0E0E0"
              If newYPDCOPE0.PDCOPETAUX = 0 Then
                wSendMail.Recipient = mRecipient_FOTC
              Else
                wSendMail.Recipient = mailAdresse_Production(newYPDCOPE0.PDCOPEIUSR) & ";" & mRecipient_FOTC
              End If
    Case Else: wSubject = "statut : " & newYPDCOPE0.PDCOPESTA: mbgColor = "bgcolor = #E0E0E0"
               wSendMail.Recipient = mailAdresse_Production(newYPDCOPE0.PDCOPEIUSR) & ";" & mRecipient_FOTC
   
End Select

'__________________________________________________________________________

If blnPDCOPE_CONF_CALL_Saisie Or Trim(newYPDCOPE0.PDCOPEVTXT) = "CONF_CALL" Then
    If newYPDCOPE0.PDCOPESTA = "V" Then wSubject = " Confirmation CONF_CALL"
    
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 " _
         & " Where BASTABETA = 1 and BASTABNUM = 6 and BASTABARG = 'CLI" & mCLIENARES & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        X = rsSab("BASTABLO2") & rsSab("BASTABDON")
        mMNURUTUTI = Mid$(X, 36, 10)
        xSQL = "select MNUUTIMAI from " & paramIBM_Library_SAB & ".ZMNURUT0 , " _
         & paramIBM_Library_SAB & ".ZMNUUTI0 " _
         & " Where MNURUTUTI = '" & mMNURUTUTI & "' and MNURUTCUT = MNUUTICUT"
        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            wSendMail.Recipient = mRecipient_CONF_CALL & ";" & rsSab("MNUUTIMAI")
        Else
            wSendMail.Recipient = mRecipient_CONF_CALL
        End If
    End If

End If
'__________________________________________________________________________

wSendMail.Subject = wSendMail.Subject & wSubject
xHeader = "<TR>" _
         & "<TD bgcolor=#0090A0  width=200 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF>" _
         & dateImp10(newYPDCOPE0.PDCOPEDTR) & " - " & newYPDCOPE0.PDCOPEID & "</TD>" _
         & "<TD bgcolor=#0090A0  width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF>" _
         & "Service : " & newYPDCOPE0.PDCOPESER & " " & newYPDCOPE0.PDCOPESSE _
         & " - " & newYPDCOPE0.PDCOPEOPEC & " " & newYPDCOPE0.PDCOPEOPET & Format$(newYPDCOPE0.PDCOPEOPEN, "### ###") & "</TD>" _
         & "<TD bgcolor=#0090A0  width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF>" _
         & wSubject & "</TD>" _
        & "</TR>"


If newYPDCOPE0.PDCOPESENS = "A" Then
    cboPDCOPESENS.ListIndex = 0
Else
    cboPDCOPESENS.ListIndex = cboPDCOPESENS.ListCount - 1
End If

Call ZCLIEAN0_SQL(newYPDCOPE0.PDCOPECLI)


xDétail = "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "Devise principale" & "</TD>" _
     & "<TD  " & mbgColor & "  width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Red _
     & "<B><div align=" & Asc34 & "right" & Asc34 & ">" & Format$(Abs(newYPDCOPE0.PDCOPEMTD1), "### ### ### ##0.00") & " " & newYPDCOPE0.PDCOPEDEV1 & "</div></TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Red _
     & cboPDCOPESENS.Text & "</B/TD>" _
     & "</TR>" _
     & "<TR>" _
     & "<TD  bgcolor=#BFFFFF width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "Devise secondaire" & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Blue _
     & "<B><div align=" & Asc34 & "right" & Asc34 & ">" & Format$(Abs(newYPDCOPE0.PDCOPEMTD2), "### ### ### ##0.00") & " " & newYPDCOPE0.PDCOPEDEV2 & "</div></TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Red _
     & "Cours " & lblPDCOPETAUX & " : " & Format$(newYPDCOPE0.PDCOPETAUX, "### ##0.000 000") & "</B/TD>" _
     & "</TR>"
     
xDétail = xDétail _
     & "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "Contrepartie" & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:7.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & newYPDCOPE0.PDCOPECLI & " " & mCLIENARA1 & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:9.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Blue _
     & "Date valeur : " & dateImp10(newYPDCOPE0.PDCOPEDVA) & "</TD>" _
     & "</TR>" _
     & "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "Saisie par" & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:7.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & newYPDCOPE0.PDCOPEIUSR & "  " & dateImp10(newYPDCOPE0.PDCOPEIAMJ) & "  " & timeImp8(newYPDCOPE0.PDCOPEIHMS) & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:7.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Magenta _
     & "." & newYPDCOPE0.PDCOPEITXT & "</TD>" _
     & "</TR>" _
     & "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "mise à jour par" & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & newYPDCOPE0.PDCOPEVUSR & "  " & dateImp10(newYPDCOPE0.PDCOPEVAMJ) & "  " & timeImp8(newYPDCOPE0.PDCOPEVHMS) & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:7.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Magenta _
     & "." & newYPDCOPE0.PDCOPEVTXT & "</TD>" _
     & "</TR>"



wSendMail.Attachment = ""
wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                    & "<BR>" & "<TABLE border = 1  width=800 height=5 cellpadding=3 >" _
                    & xHeader _
                    & xDétail _
                    & "</TABLE>"


wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail



End Sub

Public Sub cmdSendMail_Report()
Dim wSendMail As typeSendMail
Dim xHeader As String, xDétail As String, mbgColor As String
Dim wSubject As String
Dim K As Long
On Error Resume Next


wSendMail.Subject = "PDC " & dateImp10(newYPDCOPE0.PDCOPEDTR) & " - " & newYPDCOPE0.PDCOPEID & " : "

wSendMail.FromDisplayName = "@BIA_PDC"
wSendMail.RecipientDisplayName = "BIA_PDC"
wSubject = " report d'une opération non rapprochée"
wSendMail.Subject = wSendMail.Subject & wSubject
mbgColor = "bgcolor = #FFD0D0"
xHeader = "<TR>" _
         & "<TD bgcolor=#0090A0  width=200 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF>" _
         & dateImp10(newYPDCOPE0.PDCOPEDTR) & " - " & newYPDCOPE0.PDCOPEID & "</TD>" _
         & "<TD bgcolor=#0090A0  width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF>" _
         & "Service : " & newYPDCOPE0.PDCOPESER & " " & newYPDCOPE0.PDCOPESSE _
         & " - " & newYPDCOPE0.PDCOPEOPEC & " " & newYPDCOPE0.PDCOPEOPET & Format$(newYPDCOPE0.PDCOPEOPEN, "### ###") & "</TD>" _
         & "<TD bgcolor=#0090A0  width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF>" _
         & wSubject & "</TD>" _
        & "</TR>"


If newYPDCOPE0.PDCOPESENS = "A" Then
    cboPDCOPESENS.ListIndex = 0
Else
    cboPDCOPESENS.ListIndex = cboPDCOPESENS.ListCount - 1
End If
Call ZCLIEAN0_SQL(newYPDCOPE0.PDCOPECLI)


xDétail = "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "Devise principale" & "</TD>" _
     & "<TD  " & mbgColor & "  width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Red _
     & "<B><div align=" & Asc34 & "right" & Asc34 & ">" & Format$(Abs(newYPDCOPE0.PDCOPEMTD1), "### ### ### ##0.00") & " " & newYPDCOPE0.PDCOPEDEV1 & "</div></TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Red _
     & cboPDCOPESENS.Text & "</B/TD>" _
     & "</TR>" _
     & "<TR>" _
     & "<TD  bgcolor=#BFFFFF width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "Devise secondaire" & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Blue _
     & "<B><div align=" & Asc34 & "right" & Asc34 & ">" & Format$(Abs(newYPDCOPE0.PDCOPEMTD2), "### ### ### ##0.00") & " " & newYPDCOPE0.PDCOPEDEV2 & "</div></TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Red _
     & "Cours " & lblPDCOPETAUX & " : " & Format$(newYPDCOPE0.PDCOPETAUX, "### ##0.000 000") & "</B/TD>" _
     & "</TR>"
     
xDétail = xDétail _
     & "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "Contrepartie" & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:7.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & newYPDCOPE0.PDCOPECLI & " " & mCLIENARA1 & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:9.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Blue _
     & "Date valeur : " & dateImp10(newYPDCOPE0.PDCOPEDVA) & "</TD>" _
     & "</TR>" _
     & "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "Saisie par" & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:7.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & newYPDCOPE0.PDCOPEIUSR & "  " & dateImp10(newYPDCOPE0.PDCOPEIAMJ) & "  " & timeImp8(newYPDCOPE0.PDCOPEIHMS) & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:7.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Magenta _
     & "." & newYPDCOPE0.PDCOPEITXT & "</TD>" _
     & "</TR>" _
     & "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "mise à jour par" & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & newYPDCOPE0.PDCOPEVUSR & "  " & dateImp10(newYPDCOPE0.PDCOPEVAMJ) & "  " & timeImp8(newYPDCOPE0.PDCOPEVHMS) & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:7.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Magenta _
     & "." & newYPDCOPE0.PDCOPEVTXT & "</TD>" _
     & "</TR>"



wSendMail.Attachment = ""
wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                    & "<BR>" & "<TABLE border = 1  width=800 height=5 cellpadding=3 >" _
                    & xHeader _
                    & xDétail _
                    & "</TABLE>"


wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail



End Sub
Public Sub cmdSendMail_Alerte(lMsg As String)
Dim wSendMail As typeSendMail
Dim xHeader As String, xDétail As String, mbgColor As String
Dim wSubject As String
Dim K As Long
On Error Resume Next


wSendMail.Subject = "PDC Alerte : pas de traitement automatique @BIA_PDC (cours manquants au " & dateImp10(YBIATAB0_DATE_CPT_J) & ")"

wSendMail.FromDisplayName = "Alerte"
wSendMail.RecipientDisplayName = "@BIA_PDC"


X = Replace(lMsg, vbCr, "<BR>")

wSendMail.Attachment = ""
wSendMail.Message = "<body bgcolor = #FF0000>" _
                    & "<BR>""" _
                    & X

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail



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
Public Sub fgTerme_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgTerme.Row

If lRow > 0 And lRow < fgTerme.Rows Then
    fgTerme.Row = lRow
    For I = 0 To fgTerme_arrIndex
        fgTerme.Col = I: fgTerme.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgTerme.Row = mRow
    If fgTerme.Row > 0 Then
        lRow = fgTerme.Row
        lColor_Old = fgTerme.CellBackColor
        For I = 0 To fgTerme_arrIndex
          fgTerme.Col = I: fgTerme.CellBackColor = lColor
        Next I
        fgTerme.Col = 0
    End If
End If

End Sub

Public Sub fgPDCOPE_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgPDCOPE.Row

If lRow > 0 And lRow < fgPDCOPE.Rows Then
    fgPDCOPE.Row = lRow
    For I = 0 To fgPDCOPE_arrIndex
        fgPDCOPE.Col = I: fgPDCOPE.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgPDCOPE.Row = mRow
    If fgPDCOPE.Row > 0 Then
        lRow = fgPDCOPE.Row
        lColor_Old = fgPDCOPE.CellBackColor
        For I = 0 To fgPDCOPE_arrIndex
          fgPDCOPE.Col = I: fgPDCOPE.CellBackColor = lColor
        Next I
        fgPDCOPE.Col = 0
    End If
End If

End Sub
Public Sub fgTermeEch_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgTermeEch.Row

If lRow > 0 And lRow < fgTermeEch.Rows Then
    fgTermeEch.Row = lRow
    For I = 0 To fgTermeEch_arrIndex
        fgTermeEch.Col = I: fgTermeEch.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgTermeEch.Row = mRow
    If fgTermeEch.Row > 0 Then
        lRow = fgTermeEch.Row
        lColor_Old = fgTermeEch.CellBackColor
        For I = 0 To fgTermeEch_arrIndex
          fgTermeEch.Col = I: fgTermeEch.CellBackColor = lColor
        Next I
        fgTermeEch.Col = 0
    End If
End If

End Sub

Public Sub fgDetail_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgDetail.Row

If lRow > 0 And lRow < fgDetail.Rows Then
    fgDetail.Row = lRow
    For I = 0 To fgDetail_arrIndex
        fgDetail.Col = I: fgDetail.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgDetail.Row = mRow
    If fgDetail.Row > 0 Then
        lRow = fgDetail.Row
        lColor_Old = fgDetail.CellBackColor
        For I = 0 To fgDetail_arrIndex
          fgDetail.Col = I: fgDetail.CellBackColor = lColor
        Next I
        fgDetail.Col = 0
    End If
End If

End Sub

Private Sub fgSelect_Display_1()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = "fgselect_Display"
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
X = fgSelect_FormatString

fgSelect.FormatString = Replace(X, "Fixing J-1     ", "Fixing " & Mid$(dateImp10_S(FixingJ_1_AMJ), 1, 5))
'2010-10-28 : <Devise              |>Position €              |>Position Dev            |> Prix Pos         |>Fixing              |>PP Devises         |>Réeval Jour           |>PP Jour €           |>PP J-1   €         |>RPC              |

    If Not blnPDC_Instant Then
        fgSelect.BackColorFixed = &HFFFF80
        wCellBackColor = &HFFFFF0
    Else
        fgSelect.BackColorFixed = &HB0FFFF   ' &H80C0FF
        wCellBackColor = &HE0FFFF
    End If
arrYPDCPOS0(0).PDCPOSDEV = "TOT"
rsYPDCPOS0_Init arrYPDCPOS0(0)

For I = 1 To arrYPDCPOS0_Nb
         
    xYPDCPOS0 = arrYPDCPOS0(I)
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine_1 I
Next I
If blnDeviseU Then
Else
    If BIA_PDC_Aut.Avis Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        For I = 0 To fgSelect.Cols - 1
            fgSelect.Col = I
            fgSelect.CellBackColor = fgSelect.BackColorFixed
        Next I
    
        fgSelect.Col = 1
        fgSelect.Text = Format$(arrYPDCPOS0(0).PDCPOSPOSE, "### ### ### ##0.00")
        If arrYPDCPOS0(0).PDCPOSPOSE >= 0 Then
            fgSelect.CellForeColor = vbBlue
        Else
            fgSelect.CellForeColor = vbRed
        End If
        
        If chkSelect_Terme <> "1" Then
            fgSelect.Col = 8
            fgSelect.Text = Format$(arrYPDCPOS0(0).PDCPOSPNL, "### ### ### ##0.00")
            If arrYPDCPOS0(0).PDCPOSPNL >= 0 Then
                fgSelect.CellForeColor = vbBlue
            Else
                fgSelect.CellForeColor = vbRed
            End If
        End If
        fgSelect.Col = 7
        fgSelect.Text = Format$(arrYPDCPOS0(0).PDCPOSPOSD, "### ### ### ##0.00")
        If arrYPDCPOS0(0).PDCPOSPOSD >= 0 Then
            fgSelect.CellForeColor = vbBlue
        Else
            fgSelect.CellForeColor = vbRed
        End If
    End If
    'fgSelect.Col = 9: fgSelect.CellBackColor = RGB(230, 230, 230)
    'fgSelect.Col = 10: fgSelect.CellBackColor = RGB(240, 240, 240)
   
End If

fgSelect.Visible = BIA_PDC_Aut.Avis
Call lstErr_AddItem(lstErr, cmdContext, "Nb de comptes : " & arrYPDCPOS0_Nb): DoEvents
'If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
'End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Public Sub cmdSelect_SQL_xls()
On Error GoTo Error_Handler
Dim K As Long, K2 As Long, I As Long, Nb As Long, iRow As Long, nbErr As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSQL As String
Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim X As String
Dim wCellBackColor As Long, wCours_format As String
Dim xCur As Currency, xCur2 As Currency
Dim blnControl_Ok As Boolean, blnControl_Err As Boolean, blnControl_Err_Msg As Boolean
Dim blnDevise_Gérée As Boolean
'______________________________________________
On Error GoTo Error_Handler


fraSelect_Comment_Xls.Visible = False
cmdSelect_Ok_xls.Visible = False
cmdSelect_Ok_xls.BackColor = &HFF00&     '&HC0FFC0
cmdSelect_Ok_xls.Caption = "validation du contrôle"
nbErr = 6
Me.Enabled = False: Me.MousePointer = vbHourglass

Call DTPicker_Control(txtSelect_AMJ_xls, mAMJ_xls)

Call arrYPDCMAIL_SQL(mAMJ_xls)
If Trim(oldYPDCMAIL.PDCMAILTXT) <> "" Then
    X = MsgBox("Contrôle déjà validé, voulez-vous VOUS envoyer un duplicata ?", vbYesNo, "BIA_PDC : contrôle PDC / BOTC.xls")
    If X = vbYes Then cmdSendMail_xls_Duplicata: GoTo Exit_sub
End If

fgPDCOPE.Visible = False
fgDetail.Visible = False
fgSelect.Visible = False

'______________________________________________


wFilex = Trim(txtSelect_File_xls)

'______________
blnControl = False
chkSelect_Terme = "0"
chkSelect_HB = "0"
chkSelect_Suspens_Out.Value = chkSelect_Suspens_Out_xls.Value
chkSelect_HB.Value = chkSelect_HB_xls.Value
blnControl = False
blnControl_Err_Msg = False

cmdSelect_SQL_Where = "where PDCPOSDTR = '" & mAMJ_xls & "'"
Call arrYPDCPOS0_SQL(cmdSelect_SQL_Where)
Call cmdSelect_SQL_1_YPDCOPE0_R(mAMJ_xls)
If arrYPDCOPE0_Nb = 0 Then
    libSelect_Report_xls = ""
    libSelect_Report_xls.Visible = False
Else
    libSelect_Report_xls = "Opérations reportées incluses:" & vbCrLf
    For K = 1 To arrYPDCOPE0_Nb
     libSelect_Report_xls = libSelect_Report_xls _
          & arrYPDCOPE0(K).PDCOPESENS & " " & arrYPDCOPE0(K).PDCOPEDEV1 & " " & Trim(Format(arrYPDCOPE0(K).PDCOPEMTD1, "### ### ### ##0.00")) _
          & " / " & arrYPDCOPE0(K).PDCOPEDEV2 & " " & Trim(Format(arrYPDCOPE0(K).PDCOPEMTD2, "### ### ### ##0.00")) _
          & vbCrLf
    Next K
    
    libSelect_Report_xls.Visible = True
End If
wAMJMin = mAMJ_xls
Call DTPicker_Control(txtSelect_AMJ_HB_xls, wAmjMin_HB)

cmdSelect_SQL_1HB

fgSelect_Display_1

fgSelect.BackColorFixed = RGB(160, 255, 160)
fgSelect.ForeColorFixed = RGB(0, 0, 128)
fgSelect.Row = 0: fgSelect.Col = 0: fgSelect.Text = "BIA_PDC" '& dateImp10_S(mAMJ_xls)
fgSelect.Row = fgSelect.Rows - 2: fgSelect.Col = 9: wCellBackColor = fgSelect.CellBackColor
fgSelect.Row = fgSelect.Rows - 1: fgSelect.Col = 9:  fgSelect.CellBackColor = wCellBackColor
fgSelect.Row = fgSelect.Rows - 2: fgSelect.Col = 10: wCellBackColor = fgSelect.CellBackColor
fgSelect.Row = fgSelect.Rows - 1: fgSelect.Col = 10:  fgSelect.CellBackColor = wCellBackColor


fgDetail_Reset

fgDetail.Rows = fgSelect.Rows
X = Replace(fgSelect_FormatString, ">Fixing J-1       |>RPC             ", "")
X = Replace(X, "Devise  ", "BOTC.xls")
fgDetail.FormatString = X
fgDetail.BackColorFixed = RGB(255, 172, 89)
fgDetail.ForeColorFixed = RGB(0, 0, 128)
wCellBackColor = &HFFFFF0

'______________________________________________

Set appExcel = CreateObject("Excel.Application")
Set wbExcel = appExcel.Workbooks.Open(wFilex)
Set wsExcel = wbExcel.Worksheets(Trim(txtSelect_Sheet_xls))
'__________________________________________________________________________________

Call lstErr_AddItem(lstErr, cmdContext, "Importation arbitrage.xls "): DoEvents

rsYPDCPOS0_Init xYPDCPOS0
xYPDCPOS0.PDCPOSDEV = YBIATAB0_DATE_CPT_J
Nb = 0
nbErr = 0

For I = 1 To 16 '15
    X = wsExcel.Cells(I, 1)
    K = InStr(X, "EUR/")
    If K > 0 Then
        xYPDCPOS0.PDCPOSDEV = Mid$(X, K + 4, 3)
         If xYPDCPOS0.PDCPOSDEV = "YEN" Then
            xYPDCPOS0.PDCPOSDEV = "JPY"
            wCours_format = "#### ##0.00"
        Else
            wCours_format = "### ##0.000000"
        End If
        blnDevise_Gérée = False
        For K2 = 1 To arrDev_Nb
            If xYPDCPOS0.PDCPOSDEV = arrDev(K2) Then
                iRow = arrDev_Row(K2)
                blnDevise_Gérée = True
                Exit For
            End If
        Next K2
        
        If blnDevise_Gérée Then
         fgDetail.Row = iRow
        fgDetail.Col = 0: fgDetail.Text = "EUR / " & xYPDCPOS0.PDCPOSDEV
           
            xCur = CCur(wsExcel.Cells(I, 2))
            fgDetail.Col = 1: fgDetail.Text = Format$(xCur, "### ### ### ##0.00")
            If xCur >= 0 Then
                 fgDetail.CellForeColor = vbBlue
             Else
                 fgDetail.CellForeColor = vbRed
             End If
    
            xCur = CCur(wsExcel.Cells(I, 3))
            fgDetail.Col = 2: fgDetail.Text = Format$(xCur, "### ### ### ##0.00")
            If xCur >= 0 Then
                 fgDetail.CellForeColor = vbBlue
             Else
                 fgDetail.CellForeColor = vbRed
             End If
    
            'fgDetail.Col = 3: fgDetail.Text = Format$(-CDbl(wsExcel.Cells(I, 4)), wCours_format)
            fgDetail.Col = 4: fgDetail.Text = Format$(CDbl(wsExcel.Cells(I, 6)), wCours_format)
            
            xCur = CCur(wsExcel.Cells(I, 7))
            fgDetail.Col = 5: fgDetail.Text = Format$(xCur, "### ### ### ##0.00")
            If xCur >= 0 Then
                 fgDetail.CellForeColor = vbBlue
             Else
                 fgDetail.CellForeColor = vbRed
             End If
    
            xCur = CCur(wsExcel.Cells(I, 8))
            fgDetail.Col = 6: fgDetail.Text = Format$(xCur, "### ### ### ##0.00")
            If xCur >= 0 Then
                 fgDetail.CellForeColor = vbBlue
             Else
                 fgDetail.CellForeColor = vbRed
             End If
    
            xCur = CCur(wsExcel.Cells(I, 9))
            fgDetail.Col = 7: fgDetail.Text = Format$(xCur, "### ### ### ##0.00")
            If xCur >= 0 Then
                 fgDetail.CellForeColor = vbBlue
             Else
                 fgDetail.CellForeColor = vbRed
             End If
    
            xCur = CCur(wsExcel.Cells(I, 10))
            fgDetail.Col = 8: fgDetail.Text = Format$(xCur, "### ### ### ##0.00")
            If xCur >= 0 Then
                 fgDetail.CellForeColor = vbBlue
             Else
                 fgDetail.CellForeColor = vbRed
             End If
        End If

    Else
        If InStr(X, "Total") > 0 Then
            fgDetail.Row = arrDev_Nb + 1 '10
            xCur = CCur(wsExcel.Cells(I, 2))
            fgDetail.Col = 1: fgDetail.Text = Format$(xCur, "### ### ### ##0.00")
            If xCur >= 0 Then
                 fgDetail.CellForeColor = vbBlue
             Else
                 fgDetail.CellForeColor = vbRed
            End If
            xCur = CCur(wsExcel.Cells(I, 9))
            fgDetail.Col = 7: fgDetail.Text = Format$(xCur, "### ### ### ##0.00")
            If xCur >= 0 Then
                 fgDetail.CellForeColor = vbBlue
             Else
                 fgDetail.CellForeColor = vbRed
            End If
            xCur = CCur(wsExcel.Cells(I, 10))
            fgDetail.Col = 8: fgDetail.Text = Format$(xCur, "### ### ### ##0.00")
            If xCur >= 0 Then
                 fgDetail.CellForeColor = vbBlue
             Else
                 fgDetail.CellForeColor = vbRed
            End If
        End If
    End If
Next I


'

'wbExcel.Close
wbExcel.Saved = True
'____________________________________________________________________________________
appExcel.Quit


Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgDetail.Row = I
    fgDetail.Col = 0
    If Trim(fgDetail.Text) <> "" Then
        For K = 0 To 8
            fgSelect.Col = K
            fgDetail.Col = K
            If fgSelect.Text = fgDetail.Text Then
                'fgDetail.CellBackColor = fgSelect.CellBackColor
                fgSelect.CellBackColor = RGB(230, 255, 230)
                fgDetail.CellBackColor = RGB(230, 255, 230)
            Else
                xCur = CCur(fgSelect.Text)
                xCur2 = CCur(fgDetail.Text)
                blnControl_Ok = False
                blnControl_Err = True
                Select Case K
                    Case 1, 2:
                                If Abs(xCur - xCur2) <= 0.05 Then
                                    blnControl_Err = False
                                Else
                                    blnControl_Err_Msg = True
                                End If
                    Case 4:
                    Case 3: If Abs(xCur - xCur2) <= 0.001 Then blnControl_Ok = True
                    Case Else: If Abs(xCur - xCur2) < 2 Then blnControl_Ok = True
                End Select
                If blnControl_Ok Then
                    fgSelect.CellBackColor = RGB(230, 255, 230)
                    fgDetail.CellBackColor = RGB(230, 255, 230)
                Else
                    If blnControl_Err Then
                        fgSelect.CellBackColor = RGB(255, 180, 255)
                        fgDetail.CellBackColor = RGB(255, 180, 255)
                    Else
                        fgSelect.CellBackColor = RGB(255, 255, 200)
                        fgDetail.CellBackColor = RGB(255, 255, 200)
                    End If
                End If
            End If
            
        Next K
    End If
Next I

fraSelect_Options.Visible = False
If mAMJ_xls = YBIATAB0_DATE_CPT_J Then
    cmdSelect_Ok_xls.Visible = True
Else
    If BIA_PDC_Aut.Xspécial Then cmdSelect_Ok_xls.Visible = True
End If
fgSelect.Visible = True
fgDetail.Visible = True

'MsgBox "cmdSelect_SQL_xls"
'GoTo Exit_sub

If BOTC_File_xls <> txtSelect_File_xls Then
    
    V = sqlYPDCMAIL_DeleteW(" where PDCMAILDTR = 1 and PDCMAILSEQ = 0")
    xYPDCMAIL.PDCMAILDTR = 1
    xYPDCMAIL.PDCMAILSEQ = 0
    xYPDCMAIL.PDCMAILTXT = txtSelect_File_xls
    V = sqlYPDCMAIL_Insert(xYPDCMAIL)
    If IsNull(V) Then BOTC_File_xls = txtSelect_File_xls

End If
If BOTC_Sheet_xls <> txtSelect_Sheet_xls Then
    
    V = sqlYPDCMAIL_DeleteW(" where PDCMAILDTR = 1 and PDCMAILSEQ = 1")
    xYPDCMAIL.PDCMAILDTR = 1
    xYPDCMAIL.PDCMAILSEQ = 1
    xYPDCMAIL.PDCMAILTXT = txtSelect_Sheet_xls
    V = sqlYPDCMAIL_Insert(xYPDCMAIL)
    If IsNull(V) Then BOTC_Sheet_xls = txtSelect_Sheet_xls

End If

If blnControl_Err_Msg And libSelect_Report_xls.Visible Then
    Call MsgBox("Vérifier si le report des opérations est justifié, " & vbCrLf _
                & " sinon ANNULER le report automatique du " & dateImp10(YBIATAB0_DATE_CPT_J), vbInformation, "Contrôle de la position de change")
End If

'_____________________________
GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
    If nbErr > 5 Then
        nbErr = nbErr + 1
        Resume Next
    End If
Exit_sub:
blnControl = True
Me.Enabled = True: Me.MousePointer = 0

End Sub


Public Sub cmdSelect_SQL_Y()
On Error GoTo Error_Handler
Dim K As Long, K2 As Long, I As Long, Nb As Long, iRow As Long, nbErr As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSQL As String
Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim X As String
Dim wCellBackColor As Long, wCours_format As String
Dim xCur As Currency, xCur2 As Currency
Dim blnControl_Ok As Boolean, blnControl_Err As Boolean, blnControl_Err_Msg As Boolean
Dim blnDevise_Gérée As Boolean
'______________________________________________
On Error GoTo Error_Handler


Me.Enabled = False: Me.MousePointer = vbHourglass

Call DTPicker_Control(txtSelect_AMJ_Y, mAMJ_xls)

Call arrYPDCMAIL_SQL(mAMJ_xls)
If Trim(oldYPDCMAIL.PDCMAILTXT) = "" Then
    X = MsgBox("Contrôle non validé", vbInformation, "BIA_PDC : contrôle PDC / BOTC.xls")
Else
    cmdSendMail_xls_Duplicata

End If

'_____________________________
GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Exit_sub:
blnControl = True
Me.Enabled = True: Me.MousePointer = 0

End Sub



Private Sub fgTerme_Display_1()
Dim I As Long, kDev As Integer
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = "fgTerme_Display"
SSTab1.Tab = 1
SSTab1.Caption = "TERME " & dateImp(wAMJMin)
SSTab1.Tab = 0
fgTerme.Visible = False

fgTerme.Rows = 1
fgTerme_Reset
fgTerme.FormatString = fgTerme_FormatString

For I = 1 To arrYPDCPOS0_Nb
         
    xYPDCPOS0 = arrYPDCPOS0(I)
    For kDev = 1 To arrDev_Nb
        If xYPDCPOS0.PDCPOSDEV = arrDev(kDev) Then Exit For
    Next kDev
    If xYPDCPOS0.PDCPOSTERD <> 0 Or xYPDCPOS0.PDCPOSSWPD <> 0 _
    Or arrTerme_DB(kDev).PDCPOSUPDS <> 0 Then
            fgTerme.Rows = fgTerme.Rows + 1
            fgTerme.Row = fgTerme.Rows - 1
            fgTerme_DisplayLine_1TER I, kDev
            
            fgTerme.Rows = fgTerme.Rows + 1
            fgTerme.Row = fgTerme.Rows - 1
            fgTerme_DisplayLine_1SWP I, kDev
    End If
Next I

fgTerme.Visible = BIA_PDC_Aut.Avis

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub fgPDCOPE_Display_1()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = "fgPDCOPE_Display"
SSTab1.Tab = 0
fgDetail.Visible = False
fgPDCOPE.Visible = False
fgPDCOPE_Reset


fgPDCOPE.Rows = 1
fgPDCOPE.FormatString = fgPDCOPE_FormatString

For I = 1 To arrYPDCOPE0_Nb
         
    xYPDCOPE0 = arrYPDCOPE0(I)
        fgPDCOPE.Rows = fgPDCOPE.Rows + 1
        fgPDCOPE.Row = fgPDCOPE.Rows - 1
        fgPDCOPE_DisplayLine_1 I
Next I

fgPDCOPE.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Nb d'opérations : " & arrYPDCOPE0_Nb): DoEvents
'If fgPDCOPE.Rows > 1 Then
'    fgPDCOPE_Sort1 = 0: fgPDCOPE_Sort2 = 2: fgPDCOPE_Sort
'End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub fgPDCOPE_ZCHGOPE0()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = "fgPDCOPE_ZCHGOPE0"
SSTab1.Tab = 0
fgDetail.Visible = False
fgPDCOPE.Visible = False
fgPDCOPE_Reset


fgPDCOPE.Rows = 1
fgPDCOPE.FormatString = fgPDCOPE_FormatString

For I = 1 To arrZCHGOPE0_Nb
         
    xZCHGOPE0 = arrZCHGOPE0(I)
        fgPDCOPE.Rows = fgPDCOPE.Rows + 1
        fgPDCOPE.Row = fgPDCOPE.Rows - 1
        fgPDCOPE_ZCHGOPE0_Line I
Next I

fgPDCOPE.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Nb d'opérations : " & arrZCHGOPE0_Nb): DoEvents
'If fgPDCOPE.Rows > 1 Then
'    fgPDCOPE_Sort1 = 0: fgPDCOPE_Sort2 = 2: fgPDCOPE_Sort
'End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub fgTermeEch_Display_1()
Dim I As Long, wColor As Long

Dim blnTER As Boolean, blnSWP As Boolean
Dim xSQL As String, xSql2 As String, wIBMAMJ As Long

Dim wMT_Eur As Currency, wMT_Dev As Currency
Dim wCol_Eur As Integer, wCol_Dev As Integer
Dim wDev As String, kDev As Integer
Dim wCHGOPEMO4 As Currency
Dim wPDCPOSFIXT    As Double, xCV As String
Dim iAMJSWP As Long
On Error GoTo Error_Handler
currentAction = "fgTermeEch_Display"
SSTab1.Tab = 1
SSTab1.Caption = "TERME    " & dateImp(wAMJMin)
SSTab1.Tab = 0
fgTermeEch.Visible = False
fgTermeEch_Reset
wIBMAMJ = wAMJMin - 19000000

fgTermeEch.Rows = 1
fgTermeEch.FormatString = fgTermeEch_FormatString
wColor = &HC0E0FF
iAMJSWP = CLng(wAMJMin) - 19000000

If blnPDC_Instant Then
    xSql2 = " and CHGOPEDT1 >= " & wIBMAMJ
Else
    If wAMJMin < mTER382100_Amj Then                ' changement de schèma comptable
        xSql2 = " and CHGOPEDT1 >= " & wIBMAMJ
    Else
        xSql2 = " and CHGOPEDT1 > " & wIBMAMJ
    End If
End If

xSQL = " where CHGOPEOPE in ('TER','SWP') and CHGOPEANN = ' '" _
    & " and CHGOPECRE <= " & wIBMAMJ & xSql2


Call arrZCHGOPE0_SQL(xSQL)

rsYPDCPOS0_Init xYPDCPOS0
rsYPDCMVT0_Init xYPDCMVT0

For kDev = 1 To arrDev_Nb
    arrTerme_DB(kDev) = xYPDCPOS0
    arrTerme_CR(kDev) = xYPDCPOS0
    arrSWP_Dev(kDev) = xYPDCMVT0
    arrSWP_Dev(kDev).PDCMVTDEV = arrDev(kDev)
    For I = 1 To arrYPDCPOS0_Nb
        If arrYPDCPOS0(I).PDCPOSDEV = arrDev(kDev) Then
            arrTerme_DB(kDev).PDCPOSFIXT = arrYPDCPOS0(I).PDCPOSFIXT
            Exit For
        End If
    Next I
Next kDev
        

For kDev = 1 To arrDev_Nb
    'arrTerme_DB(kDev) = xYPDCPOS0
    'arrTerme_CR(kDev) = xYPDCPOS0
    blnTER = False: blnSWP = False
    For I = 1 To arrZCHGOPE0_Nb
    
       xZCHGOPE0 = arrZCHGOPE0(I)
       
        
        If xZCHGOPE0.CHGOPEDE1 = "EUR" Then
        
            wDev = xZCHGOPE0.CHGOPEDE2
            wMT_Eur = xZCHGOPE0.CHGOPEMO1
            If xZCHGOPE0.CHGOPESEN = "A" Then
                wCol_Eur = 7: wCol_Dev = 9
                wCHGOPEMO4 = xZCHGOPE0.CHGOPEMO4
            Else
                wCol_Eur = 6: wCol_Dev = 10
                wCHGOPEMO4 = -xZCHGOPE0.CHGOPEMO4
           End If
            If xZCHGOPE0.CHGOPEOPE = "TER" Then
                wMT_Dev = xZCHGOPE0.CHGOPEMO2 + wCHGOPEMO4
            Else
                wMT_Dev = xZCHGOPE0.CHGOPEMO2
            End If
            
        Else
        
            If xZCHGOPE0.CHGOPEDE2 <> "EUR" Then
          '      Call MsgBox("opération croisée non traitée par ce programme", vbCritical, "BIA_PDC TERME")
            End If
            wDev = xZCHGOPE0.CHGOPEDE1
            wMT_Dev = xZCHGOPE0.CHGOPEMO1
            If xZCHGOPE0.CHGOPESEN = "A" Then
                wCol_Eur = 6: wCol_Dev = 10
                wCHGOPEMO4 = xZCHGOPE0.CHGOPEMO4
            Else
                wCol_Eur = 7: wCol_Dev = 9
                wCHGOPEMO4 = -xZCHGOPE0.CHGOPEMO4
           End If
            If xZCHGOPE0.CHGOPEOPE = "TER" Then
                wMT_Eur = xZCHGOPE0.CHGOPEMO2 + wCHGOPEMO4
            Else
                wMT_Eur = xZCHGOPE0.CHGOPEMO2
            End If
        End If
        
        If wDev = arrDev(kDev) Then
            
            For arrYPDCPOS0_Index = 1 To arrYPDCPOS0_Nb
                If wDev = arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSDEV Then Exit For
            Next arrYPDCPOS0_Index

'______________________________________________________________________________________________
            arrTerme_DB(kDev).PDCPOSUPDS = arrTerme_DB(kDev).PDCPOSUPDS + 1
            Select Case xZCHGOPE0.CHGOPEOPE
                Case "TER": blnTER = True
                     If blnSWP Then blnSWP = False: fgTermeEch_DisplayTotal_1 kDev, "SWP"
                   
                   If wCol_Eur = 6 Then
                        arrTerme_DB(kDev).PDCPOSTERE = arrTerme_DB(kDev).PDCPOSTERE + wMT_Eur
                        arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSTERE = arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSTERE + wMT_Eur
                   Else
                        arrTerme_CR(kDev).PDCPOSTERE = arrTerme_CR(kDev).PDCPOSTERE + wMT_Eur
                        arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSTERE = arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSTERE - wMT_Eur
                   End If
                    If wCol_Dev = 9 Then
                        arrTerme_DB(kDev).PDCPOSTERD = arrTerme_DB(kDev).PDCPOSTERD + wMT_Dev
                        arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSTERD = arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSTERD + wMT_Dev
                  Else
                        arrTerme_CR(kDev).PDCPOSTERD = arrTerme_CR(kDev).PDCPOSTERD + wMT_Dev
                        arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSTERD = arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSTERD - wMT_Dev
                    End If
                Case "SWP": blnSWP = True

                    If wCol_Eur = 6 Then
                        arrTerme_DB(kDev).PDCPOSSWPE = arrTerme_DB(kDev).PDCPOSSWPE + wMT_Eur
                        'arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSSWPE = arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSSWPE + wMT_Eur
                  Else
                        arrTerme_CR(kDev).PDCPOSSWPE = arrTerme_CR(kDev).PDCPOSSWPE + wMT_Eur
                        'arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSSWPE = arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSSWPE - wMT_Eur
                    End If
                    If wCol_Dev = 9 Then
                        arrTerme_DB(kDev).PDCPOSSWPD = arrTerme_DB(kDev).PDCPOSSWPD + wMT_Dev
                        'arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSSWPD = arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSSWPD + wMT_Dev
                    Else
                        arrTerme_CR(kDev).PDCPOSSWPD = arrTerme_CR(kDev).PDCPOSSWPD + wMT_Dev
                        'arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSSWPD = arrYPDCPOS0(arrYPDCPOS0_Index).PDCPOSSWPD - wMT_Dev
                    End If
                    
                    'SWP comptant HB à inclure pour contrôle du solde
                    
                    If iAMJSWP >= xZCHGOPE0.CHGOPEENG And iAMJSWP <= xZCHGOPE0.CHGOPEDT2 Then
                        
                        
                         If wCol_Eur = 6 Then
                            arrSWP_Dev(kDev).PDCMVTMTE = arrSWP_Dev(kDev).PDCMVTMTE - wMT_Eur
                        Else
                            arrSWP_Dev(kDev).PDCMVTMTE = arrSWP_Dev(kDev).PDCMVTMTE + wMT_Eur
                        End If
                        If wCol_Dev = 9 Then
                            arrSWP_Dev(kDev).PDCMVTMTD = arrSWP_Dev(kDev).PDCMVTMTD - wMT_Dev
                        Else
                            arrSWP_Dev(kDev).PDCMVTMTD = arrSWP_Dev(kDev).PDCMVTMTD + wMT_Dev
                        End If
                  End If
                    
            End Select
    '______________________________________________________________________________________________
            If arrTerme_DB(kDev).PDCPOSFIXT = 0 Then
                xCV = ""
            Else
                xCV = Format$(Round(wMT_Dev / arrTerme_DB(kDev).PDCPOSFIXT, 2), "### ### ### ##0.00")
            End If
            fgTermeEch.Rows = fgTermeEch.Rows + 1
            fgTermeEch.Row = fgTermeEch.Rows - 1
            Call fgTermeEch_DisplayLine_1(I, wCol_Eur, wMT_Eur, wCol_Dev, wMT_Dev, xCV)


        End If
    Next I
'______________________________________________________________________________________________

    If arrTerme_DB(kDev).PDCPOSUPDS > 0 Then
        If blnTER Then fgTermeEch_DisplayTotal_1 kDev, "TER"
        If blnSWP Then fgTermeEch_DisplayTotal_1 kDev, "SWP"
        If wAMJMin >= YBIATAB0_DATE_CPT_J Then fgTermeEch_DisplayTotal_1 kDev, "TOT"
        
    End If
'______________________________________________________________________________________________

Next kDev


fgTermeEch.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Nb d'opérations : " & arrYPDCOPE0_Nb): DoEvents
'If fgTermeEch.Rows > 1 Then
'    fgTermeEch_Sort1 = 0: fgTermeEch_Sort2 = 2: fgTermeEch_Sort
'End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


Private Sub fgDetail_Display_YPDCMVT0()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = "fgDetail_Display"
SSTab1.Tab = 0
fgPDCOPE.Visible = False
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
    
For I = 1 To arrYPDCMVT0_Nb
         
    xYPDCMVT0 = arrYPDCMVT0(I)
        fgDetail.Rows = fgDetail.Rows + 1
        fgDetail.Row = fgDetail.Rows - 1
        fgDetail_DisplayLine_YPDCMVT0 I
Next I

fgDetail.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Nb de mouvements : " & arrYPDCMVT0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub fgDetail_Display_YPDCLOG0()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = "fgDetail_Display"
SSTab1.Tab = 0
fgPDCOPE.Visible = False
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgYPDCLOG0_FormatString

    
For I = 1 To arrYPDCLOG0_Nb
         
    xYPDCLOG0 = arrYPDCLOG0(I)
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_DisplayLine_YPDCLOG0 I
Next I

fgDetail.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Nb de mouvements : " & arrYPDCLOG0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub arrYPDCPOS0_SQL(xWhere As String)
Dim V, I As Integer, K As Integer, I2 As Integer
Dim X As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrYPDCPOS0(101)
arrYPDCPOS0_Max = 100: arrYPDCPOS0_Nb = 0

rsYPDCPOS0_Init arrYPDCPOS0(0)

xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCPOS0 " & xWhere & " order by PDCPOSDTR,PDCPOSDEV"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYPDCPOS0_GetBuffer(rsSab, xYPDCPOS0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmBIA_PDC.fgselect_Display"
        '' Exit Sub
     Else
         arrYPDCPOS0_Nb = arrYPDCPOS0_Nb + 1
         If arrYPDCPOS0_Nb > arrYPDCPOS0_Max Then
             arrYPDCPOS0_Max = arrYPDCPOS0_Max + 50
             ReDim Preserve arrYPDCPOS0(arrYPDCPOS0_Max)
         End If
         arrYPDCPOS0(arrYPDCPOS0_Nb) = xYPDCPOS0
    End If
    rsSab.MoveNext

Loop


Set rsSab = Nothing
If Not blnDeviseU Then

    'If arrYPDCPOS0_Nb = arrDev_Nb Then
        ReDim selYPDCPOS0(arrYPDCPOS0_Nb + arrDev_Nb)
        
        ReDim blnOk(arrYPDCPOS0_Nb + arrDev_Nb) As Boolean
        If arrYPDCPOS0_Nb > 0 Then
            I2 = arrDev_Nb
        Else
            I2 = 0
        End If
        
        For I = 1 To arrYPDCPOS0_Nb 'arrDev_Nb
            selYPDCPOS0(I) = arrYPDCPOS0(I)
            arrYPDCPOS0(I) = arrYPDCPOS0(0)
        Next I
        For I = 1 To arrDev_Nb
            arrYPDCPOS0(arrDev_Row(I)).PDCPOSDEV = arrDev(I)
        Next I
        For K = 1 To arrYPDCPOS0_Nb
            For I = 1 To arrDev_Nb
                If arrYPDCPOS0(I).PDCPOSDEV = selYPDCPOS0(K).PDCPOSDEV Then
                    arrYPDCPOS0(I) = selYPDCPOS0(K)
                    blnOk(K) = True
                    Exit For
                End If
            Next I
        Next K
         For K = 1 To arrYPDCPOS0_Nb
            If Not blnOk(K) Then
                I2 = I2 + 1
                arrYPDCPOS0(I2) = selYPDCPOS0(K)
            End If
        Next K
        arrYPDCPOS0_Nb = I2
   'End If
    
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrYPDCPOS0_SQL_Old(xWhere As String)
Dim V, I As Integer, K As Integer
Dim X As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrYPDCPOS0(101)
arrYPDCPOS0_Max = 100: arrYPDCPOS0_Nb = 0

rsYPDCPOS0_Init arrYPDCPOS0(0)

xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCPOS0 " & xWhere & " order by PDCPOSDTR,PDCPOSDEV"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYPDCPOS0_GetBuffer(rsSab, xYPDCPOS0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmBIA_PDC.fgselect_Display"
        '' Exit Sub
     Else
         arrYPDCPOS0_Nb = arrYPDCPOS0_Nb + 1
         If arrYPDCPOS0_Nb > arrYPDCPOS0_Max Then
             arrYPDCPOS0_Max = arrYPDCPOS0_Max + 50
             ReDim Preserve arrYPDCPOS0(arrYPDCPOS0_Max)
         End If
         arrYPDCPOS0(arrYPDCPOS0_Nb) = xYPDCPOS0
    End If
    rsSab.MoveNext

Loop


Set rsSab = Nothing
If Not blnDeviseU Then

    'If arrYPDCPOS0_Nb = arrDev_Nb Then
        ReDim selYPDCPOS0(arrYPDCPOS0_Nb + arrDev_Nb)
        
        
        For I = 1 To arrDev_Nb
            selYPDCPOS0(I) = arrYPDCPOS0(I)
            arrYPDCPOS0(I) = arrYPDCPOS0(0)
        Next I
        For I = 1 To arrDev_Nb
            arrYPDCPOS0(arrDev_Row(I)).PDCPOSDEV = arrDev(I)
        Next I
        For K = 1 To arrYPDCPOS0_Nb
            For I = 1 To arrDev_Nb
                If arrYPDCPOS0(I).PDCPOSDEV = selYPDCPOS0(K).PDCPOSDEV Then
                    arrYPDCPOS0(I) = selYPDCPOS0(K)
                    Exit For
                End If
            Next I
        Next K
    'End If
    
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub arrYPDCPOS0_SQL_FixingJ_1(lAMJMin As String)
Dim V, I As Integer
Dim X As String, xSQL As String
Dim blnOk As Boolean
On Error GoTo Error_Handler

blnOk = False
FixingJ_1_AMJ = lAMJMin

Do

    FixingJ_1_AMJ = dateElp("Ouvré", -1, FixingJ_1_AMJ)
    
    
    xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCPOS0 where PDCPOSDTR = '" & FixingJ_1_AMJ & "'" & " order by PDCPOSDEV"
    Set rsSab = cnsab.Execute(xSQL)

    Do While Not rsSab.EOF
        V = rsYPDCPOS0_GetBuffer(rsSab, xYPDCPOS0)
        For I = 1 To arrDev_Nb
            If xYPDCPOS0.PDCPOSDEV = arrDev(I) Then
                fixingJ_1(arrDev_Row(I)) = xYPDCPOS0
                blnOk = True
                Exit For
            End If
        Next I
    
        rsSab.MoveNext
    
    Loop
Loop Until blnOk

Set rsSab = Nothing
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrYPDCMVT0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
Dim wCli As Long
On Error GoTo Error_Handler
ReDim arrYPDCMVT0(101)
arrYPDCMVT0_Max = 100: arrYPDCMVT0_Nb = 0


xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCMVT0 " & xWhere & " order by PDCMVTDTR  , PDCMVTOPEN"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYPDCMVT0_GetBuffer(rsSab, xYPDCMVT0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmBIA_PDC.fgselect_Display"
        '' Exit Sub
     Else
         arrYPDCMVT0_Nb = arrYPDCMVT0_Nb + 1
         If arrYPDCMVT0_Nb > arrYPDCMVT0_Max Then
             arrYPDCMVT0_Max = arrYPDCMVT0_Max + 50
             ReDim Preserve arrYPDCMVT0(arrYPDCMVT0_Max)
         End If

         arrYPDCMVT0(arrYPDCMVT0_Nb) = xYPDCMVT0
    End If
    rsSab.MoveNext

Loop


Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrYPDCLOG0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
Dim wCli As Long
On Error GoTo Error_Handler
ReDim arrYPDCLOG0(101)
arrYPDCLOG0_Max = 100: arrYPDCLOG0_Nb = 0


xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCLOG0 " & xWhere & " order by  PDCLOGUAMJ , PDCLOGUHMS,PDCLOGUSEQ"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYPDCLOG0_GetBuffer(rsSab, xYPDCLOG0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmBIA_PDC.fgselect_Display"
        '' Exit Sub
     Else
         arrYPDCLOG0_Nb = arrYPDCLOG0_Nb + 1
         If arrYPDCLOG0_Nb > arrYPDCLOG0_Max Then
             arrYPDCLOG0_Max = arrYPDCLOG0_Max + 50
             ReDim Preserve arrYPDCLOG0(arrYPDCLOG0_Max)
         End If

         arrYPDCLOG0(arrYPDCLOG0_Nb) = xYPDCLOG0
    End If
    rsSab.MoveNext

Loop


Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrYPDCMAIL_SQL(lAMJ As String)
Dim V
Dim X As String, xSQL As String
Dim blnOk As Boolean
On Error GoTo Error_Handler

Call rsYPDCMAIL_Init(oldYPDCMAIL)
xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCMAIL where PDCMAILDTR = '" & lAMJ & "' order by  PDCMAILSEQ"
Set rsSab = cnsab.Execute(xSQL)
blnOk = False
Do While Not rsSab.EOF
    If Not blnOk Then
        V = rsYPDCMAIL_GetBuffer(rsSab, oldYPDCMAIL)
        blnOk = True
    Else
        oldYPDCMAIL.PDCMAILTXT = oldYPDCMAIL.PDCMAILTXT & rsSab("PDCMAILTXT")
    End If
    rsSab.MoveNext

Loop
oldYPDCMAIL.PDCMAILTXT = Replace(oldYPDCMAIL.PDCMAILTXT, Chr$(26), Chr$(128))
oldYPDCMAIL.PDCMAILTXT = Replace(oldYPDCMAIL.PDCMAILTXT, "|", "'")

Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrZCHGOPE0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
Dim wCli As Long
On Error GoTo Error_Handler
ReDim arrZCHGOPE0(101)
arrZCHGOPE0_Max = 100: arrZCHGOPE0_Nb = 0


xSQL = "select * from " & paramIBM_Library_SAB & ".ZCHGOPE0 " & xWhere & " order by  CHGOPEOPE, CHGOPEDT1, CHGOPEDOS"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsZCHGOPE0_GetBuffer(rsSab, xZCHGOPE0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "arrZCHGOPE0_SQL"
        '' Exit Sub
     Else
         arrZCHGOPE0_Nb = arrZCHGOPE0_Nb + 1
         If arrZCHGOPE0_Nb > arrZCHGOPE0_Max Then
             arrZCHGOPE0_Max = arrZCHGOPE0_Max + 50
             ReDim Preserve arrZCHGOPE0(arrZCHGOPE0_Max)
         End If

         arrZCHGOPE0(arrZCHGOPE0_Nb) = xZCHGOPE0
    End If
    rsSab.MoveNext

Loop


Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub ZCLIEAN0_SQL(lCLIENACLI As String)
Dim X As String, xSQL As String
On Error GoTo Error_Handler

If mCLIENACLI <> lCLIENACLI Then
    xSQL = "select CLIENARA1, CLIENARES from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & lCLIENACLI & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If rsSab.EOF Then
        mCLIENARA1 = ""
    Else
        mCLIENACLI = lCLIENACLI
        mCLIENARA1 = rsSab("CLIENARA1")
        mCLIENARES = rsSab("CLIENARES")
    End If
    Set rsSab = Nothing
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : ZCLIEAN0_SQL " & lCLIENACLI

End Sub

Private Sub arrYPDCOPE0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
Dim wCli As Long
On Error GoTo Error_Handler
ReDim arrYPDCOPE0(101)
arrYPDCOPE0_Max = 100: arrYPDCOPE0_Nb = 0


xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCOPE0 " & xWhere & " order by PDCOPEDTR , PDCOPEID"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYPDCOPE0_GetBuffer(rsSab, xYPDCOPE0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmBIA_PDC.fgselect_Display"
        '' Exit Sub
     Else
         arrYPDCOPE0_Nb = arrYPDCOPE0_Nb + 1
         If arrYPDCOPE0_Nb > arrYPDCOPE0_Max Then
             arrYPDCOPE0_Max = arrYPDCOPE0_Max + 50
             ReDim Preserve arrYPDCOPE0(arrYPDCOPE0_Max)
         End If

         arrYPDCOPE0(arrYPDCOPE0_Nb) = xYPDCOPE0
    End If
    rsSab.MoveNext

Loop


Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Public Sub fgSelect_DisplayLine_1(lIndex As Long)
Dim devFixing As Currency, devPP As Currency, eurPP As Currency, curPNL As Currency
Dim wPDCPOSPOSE As Currency, wPDCPOSPOSD As Currency
Dim wCours_format As String
Dim xSQL As String
On Error Resume Next
If blnDeviseU Then
    fgSelect.Col = 0: fgSelect.Text = dateImp10(xYPDCPOS0.PDCPOSDTR)
Else
    fgSelect.Col = 0: fgSelect.Text = "EUR / " & xYPDCPOS0.PDCPOSDEV
End If
fgSelect.CellBackColor = fgSelect.BackColorFixed  '&HFFFFA0

If chkSelect_Terme <> "1" Then
    wPDCPOSPOSE = -xYPDCPOS0.PDCPOSPOSE
    wPDCPOSPOSD = -xYPDCPOS0.PDCPOSPOSD
Else
    wPDCPOSPOSE = -xYPDCPOS0.PDCPOSPOSE - xYPDCPOS0.PDCPOSTERE - xYPDCPOS0.PDCPOSSWPE
    wPDCPOSPOSD = -xYPDCPOS0.PDCPOSPOSD - xYPDCPOS0.PDCPOSTERD - xYPDCPOS0.PDCPOSSWPD
End If

arrYPDCPOS0(0).PDCPOSPOSE = arrYPDCPOS0(0).PDCPOSPOSE + wPDCPOSPOSE
fgSelect.Col = 1: fgSelect.Text = Format$(wPDCPOSPOSE, "### ### ### ##0.00")
fgSelect.CellBackColor = wCellBackColor
If wPDCPOSPOSE >= 0 Then
     fgSelect.CellForeColor = vbBlue
 Else
     fgSelect.CellForeColor = vbRed
 End If
 
fgSelect.Col = 2: fgSelect.Text = Format$(wPDCPOSPOSD, "### ### ### ##0.00")
fgSelect.CellBackColor = wCellBackColor
If wPDCPOSPOSD >= 0 Then
     fgSelect.CellForeColor = vbBlue
 Else
     fgSelect.CellForeColor = vbRed
 End If
 If xYPDCPOS0.PDCPOSDEV = "JPY" Then
    wCours_format = "#### ##0.00"
Else
    wCours_format = "### ##0.000000"
End If

If cmdSelect_SQL_K <> "X" Then
    fgSelect.Col = 3:
    If chkSelect_Terme <> "1" Then
        If xYPDCPOS0.PDCPOSPOSE = 0 Then
            xYPDCPOS0.PDCPOSPRIX = 0
        Else
            xYPDCPOS0.PDCPOSPRIX = Round(Abs(xYPDCPOS0.PDCPOSPOSD / xYPDCPOS0.PDCPOSPOSE), 6)
            If xYPDCPOS0.PDCPOSPRIX > 999 Then xYPDCPOS0.PDCPOSPRIX = 0
        End If
    
        fgSelect.Text = Format$(xYPDCPOS0.PDCPOSPRIX, wCours_format)
    End If
End If

fgSelect.CellBackColor = wCellBackColor

fgSelect.Col = 4: fgSelect.Text = Format$(xYPDCPOS0.PDCPOSFIXT, wCours_format)
If xYPDCPOS0.PDCPOSDTR = xYPDCPOS0.PDCPOSFIXD Then
    fgSelect.CellBackColor = wCellBackColor
Else
    fgSelect.CellBackColor = RGB(230, 230, 230)
End If


'devFixing = -Round(wPDCPOSPOSE * xYPDCPOS0.PDCPOSFIXT, 2) ' Commentaire Kokou : J'ai enlevé la fonction ROUND pour ne pas arrondir le résultat à 2 chiffres après la virgule

devFixing = -(wPDCPOSPOSE * xYPDCPOS0.PDCPOSFIXT) ' Commentaire Kokou : J'ai enlevé la fonction ROUND pour ne pas arrondir le résultat à 2 chiffres après la virgule

fgSelect.Col = 6: fgSelect.Text = Format$(devFixing, "### ### ### ##0.00")
'fgSelect.CellBackColor = &HF0FFFF
If devFixing >= 0 Then
     fgSelect.CellForeColor = vbBlue
 Else
     fgSelect.CellForeColor = vbRed
 End If

devPP = wPDCPOSPOSD - devFixing
fgSelect.Col = 5: fgSelect.Text = Format$(devPP, "### ### ### ##0.00")
'fgSelect.CellBackColor = &HF0FFFF
If devPP >= 0 Then
     fgSelect.CellForeColor = vbBlue
 Else
     fgSelect.CellForeColor = vbRed
 End If

'eurPP = Round(devPP / xYPDCPOS0.PDCPOSFIXT, 2) ' Commentaire Kokou : J'ai enlevé la fonction ROUND pour ne pas arrondir le résultat à 2 chiffres après la virgule

eurPP = (devPP / xYPDCPOS0.PDCPOSFIXT) ' Commentaire Kokou : J'ai enlevé la fonction ROUND pour ne pas arrondir le résultat à 2 chiffres après la virgule

arrYPDCPOS0(0).PDCPOSPOSD = arrYPDCPOS0(0).PDCPOSPOSD + eurPP   ' !!!!!!!!!! cumul pp jour

fgSelect.Col = 7: fgSelect.Text = Format$(eurPP, "### ### ### ##0.00")
arrPPJ(lIndex) = eurPP
'fgSelect.CellBackColor = &HF0FFFF
If eurPP >= 0 Then
     fgSelect.CellForeColor = vbBlue
 Else
     fgSelect.CellForeColor = vbRed
 End If

'fin de mois non ouvrée : PNL pour contrôle BOTC
'----------------------------------------------------------
curPNL = xYPDCPOS0.PDCPOSPNL
If cmdSelect_SQL_K = "X" And mAMJ_xls <> mAMJ_JP0 Then
    xSQL = "select PDCPOSPNL from " & paramIBM_Library_SABSPE_XXX & ".YPDCPOS0 " _
         & "Where PDCPOSDTR = " & mAMJ_JP0 _
         & " and PDCPOSDEV = '" & xYPDCPOS0.PDCPOSDEV & "'"
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then curPNL = rsSab("PDCPOSPNL")

End If

arrYPDCPOS0(0).PDCPOSPNL = arrYPDCPOS0(0).PDCPOSPNL + curPNL
fgSelect.Col = 8
If chkSelect_Terme <> "1" Then fgSelect.Text = Format$(curPNL, "### ### ### ##0.00")
'fgSelect.CellBackColor = &HF0FFFF
If curPNL >= 0 Then
     fgSelect.CellForeColor = vbBlue
 Else
     fgSelect.CellForeColor = vbRed
 End If
 
If Not blnDeviseU Then
    If fixingJ_1(lIndex).PDCPOSDEV = xYPDCPOS0.PDCPOSDEV Then
        fgSelect.Col = 9: fgSelect.Text = Format$(fixingJ_1(lIndex).PDCPOSFIXT, wCours_format)
        fgSelect.CellBackColor = RGB(230, 230, 230)
    End If
End If

arrYPDCPOS0(0).PDCPOSRPC = arrYPDCPOS0(0).PDCPOSRPC + xYPDCPOS0.PDCPOSRPC
fgSelect.Col = 10: fgSelect.Text = Format$(xYPDCPOS0.PDCPOSRPC, "### ### ### ##0.00")
fgSelect.CellBackColor = RGB(240, 240, 240) ' wCellBackColor
If xYPDCPOS0.PDCPOSRPC >= 0 Then
     fgSelect.CellForeColor = vbBlue
 Else
     fgSelect.CellForeColor = vbRed
 End If

Exit_sub:

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex



End Sub
Public Sub fgTerme_DisplayLine_1TER(lIndex As Long, kDev As Integer)
Dim devFixing As Currency, devPP As Currency, eurPP As Currency
Dim wPDCPOSTERE As Currency, wPDCPOSTERD As Currency
Dim wCours_format As String

On Error Resume Next
If blnDeviseU Then
    fgTerme.Col = 0: fgTerme.Text = dateImp10(xYPDCPOS0.PDCPOSDTR)
Else
    fgTerme.Col = 0: fgTerme.Text = "terme / " & xYPDCPOS0.PDCPOSDEV
End If
fgTerme.CellForeColor = vbBlue
fgTerme.CellBackColor = &HC0E0FF

wPDCPOSTERE = -xYPDCPOS0.PDCPOSTERE
arrYPDCPOS0(0).PDCPOSTERE = arrYPDCPOS0(0).PDCPOSTERE + wPDCPOSTERE
fgTerme.Col = 1
If wPDCPOSTERE <> 0 Then fgTerme.Text = Format$(wPDCPOSTERE, "### ### ### ##0.00")
fgTerme.CellBackColor = &HC0E0FF
If wPDCPOSTERE >= 0 Then
     fgTerme.CellForeColor = vbBlue
 Else
     fgTerme.CellForeColor = vbRed
 End If
 
wPDCPOSTERD = -xYPDCPOS0.PDCPOSTERD
fgTerme.Col = 2
If wPDCPOSTERD <> 0 Then fgTerme.Text = Format$(wPDCPOSTERD, "### ### ### ##0.00")
fgTerme.CellBackColor = &HC0E0FF
If wPDCPOSTERD >= 0 Then
     fgTerme.CellForeColor = vbBlue
 Else
     fgTerme.CellForeColor = vbRed
 End If
 If xYPDCPOS0.PDCPOSDEV = "JPY" Then
    wCours_format = "#### ##0.00"
Else
    wCours_format = "### ##0.000000"
End If

fgTerme.Col = 3
fgTerme.CellBackColor = &HC0E0FF

fgTerme.Col = 4: fgTerme.Text = Format$(xYPDCPOS0.PDCPOSFIXT, wCours_format)
fgTerme.CellBackColor = &HC0E0FF

fgTerme.Col = 6
If arrTerme_DB(kDev).PDCPOSTERE <> 0 Then fgTerme.Text = Format$(arrTerme_DB(kDev).PDCPOSTERE, "### ### ### ##0.00")
fgTerme.CellForeColor = vbRed
fgTerme.CellBackColor = &HC0E0FF
fgTerme.Col = 7
If arrTerme_CR(kDev).PDCPOSTERE <> 0 Then fgTerme.Text = Format$(arrTerme_CR(kDev).PDCPOSTERE, "### ### ### ##0.00")
fgTerme.CellForeColor = vbBlue
fgTerme.CellBackColor = &HC0E0FF
fgTerme.Col = 8: fgTerme.Text = xYPDCPOS0.PDCPOSDEV
fgTerme.CellBackColor = &HC0E0FF
fgTerme.Col = 9
If arrTerme_DB(kDev).PDCPOSTERD <> 0 Then fgTerme.Text = Format$(arrTerme_DB(kDev).PDCPOSTERD, "### ### ### ##0.00")
fgTerme.CellForeColor = vbRed
fgTerme.CellBackColor = &HC0E0FF
fgTerme.Col = 10
If arrTerme_CR(kDev).PDCPOSTERD <> 0 Then fgTerme.Text = Format$(arrTerme_CR(kDev).PDCPOSTERD, "### ### ### ##0.00")
fgTerme.CellForeColor = vbBlue
fgTerme.CellBackColor = &HC0E0FF


fgTerme.Col = fgTerme_arrIndex: fgTerme.Text = lIndex



End Sub

Public Sub fgTerme_DisplayLine_1SWP(lIndex As Long, kDev As Integer)
Dim devFixing As Currency, devPP As Currency, eurPP As Currency
Dim wPDCPOSSWPE As Currency, wPDCPOSSWPD As Currency
Dim wCours_format As String

On Error Resume Next
If blnDeviseU Then
    fgTerme.Col = 0: fgTerme.Text = dateImp10(xYPDCPOS0.PDCPOSDTR)
Else
    fgTerme.Col = 0: fgTerme.Text = "swap / " & xYPDCPOS0.PDCPOSDEV
End If
fgTerme.CellForeColor = vbBlue
fgTerme.CellBackColor = &HE0FFFF

wPDCPOSSWPE = -xYPDCPOS0.PDCPOSSWPE
arrYPDCPOS0(0).PDCPOSSWPE = arrYPDCPOS0(0).PDCPOSSWPE + wPDCPOSSWPE
fgTerme.Col = 1
If wPDCPOSSWPE <> 0 Then fgTerme.Text = Format$(wPDCPOSSWPE, "### ### ### ##0.00")
fgTerme.CellBackColor = &HE0FFFF
If wPDCPOSSWPE >= 0 Then
     fgTerme.CellForeColor = vbBlue
 Else
     fgTerme.CellForeColor = vbRed
 End If
 
wPDCPOSSWPD = -xYPDCPOS0.PDCPOSSWPD
fgTerme.Col = 2
If wPDCPOSSWPD <> 0 Then fgTerme.Text = Format$(wPDCPOSSWPD, "### ### ### ##0.00")
fgTerme.CellBackColor = &HE0FFFF
If wPDCPOSSWPD >= 0 Then
     fgTerme.CellForeColor = vbBlue
 Else
     fgTerme.CellForeColor = vbRed
 End If
 If xYPDCPOS0.PDCPOSDEV = "JPY" Then
    wCours_format = "#### ##0.00"
Else
    wCours_format = "### ##0.000000"
End If

fgTerme.Col = 3
fgTerme.CellBackColor = &HE0FFFF

fgTerme.Col = 4: fgTerme.Text = Format$(xYPDCPOS0.PDCPOSFIXT, wCours_format)
fgTerme.CellBackColor = &HE0FFFF

fgTerme.Col = 6
If arrTerme_DB(kDev).PDCPOSSWPE <> 0 Then fgTerme.Text = Format$(arrTerme_DB(kDev).PDCPOSSWPE, "### ### ### ##0.00")
fgTerme.CellForeColor = vbRed
fgTerme.CellBackColor = &HE0FFFF
fgTerme.Col = 7
If arrTerme_CR(kDev).PDCPOSSWPE <> 0 Then fgTerme.Text = Format$(arrTerme_CR(kDev).PDCPOSSWPE, "### ### ### ##0.00")
fgTerme.CellForeColor = vbBlue
fgTerme.CellBackColor = &HE0FFFF
fgTerme.Col = 8: fgTerme.Text = xYPDCPOS0.PDCPOSDEV
fgTerme.CellBackColor = &HE0FFFF
fgTerme.CellBackColor = &HE0FFFF
fgTerme.Col = 9
If arrTerme_DB(kDev).PDCPOSSWPD <> 0 Then fgTerme.Text = Format$(arrTerme_DB(kDev).PDCPOSSWPD, "### ### ### ##0.00")
fgTerme.CellForeColor = vbRed
fgTerme.CellBackColor = &HE0FFFF
fgTerme.Col = 10
If arrTerme_CR(kDev).PDCPOSSWPD <> 0 Then fgTerme.Text = Format$(arrTerme_CR(kDev).PDCPOSSWPD, "### ### ### ##0.00")
fgTerme.CellForeColor = vbBlue
fgTerme.CellBackColor = &HE0FFFF


fgTerme.Col = fgTerme_arrIndex: fgTerme.Text = lIndex



End Sub

Public Sub fgPDCOPE_DisplayLine_1(lIndex As Long)
Dim X As String, wColor As Long, wBackColor As Long
Dim blnAnnulation As Boolean
Dim wMTD1 As Currency, wMTD2 As Currency
Dim wCol_Eur As Integer, wCol_Dev As Integer
Dim xDev1_display As String, xDev2_display As String
On Error Resume Next
blnAnnulation = False
 wBackColor = &HE0FFFF
Select Case xYPDCOPE0.PDCOPESTA
    Case Is = "V", "T": wColor = &H6000&
    Case Is = "A": wColor = &H808080: blnAnnulation = True
    Case Is = "I": wColor = &HC0C000: blnAnnulation = True
    Case Is = "0", "B": wColor = vbMagenta
    Case Is = "R": wColor = vbRed: wBackColor = vbYellow ' &HC0C0FF
    Case Else: wColor = vbBlack
End Select


wMTD1 = -xYPDCOPE0.PDCOPEMTD1
'wBackColor = &HE0FFFF

xDev1_display = " *"
Select Case xYPDCOPE0.PDCOPEDEV1
    Case "EUR":  xDev1_display = " €"
    Case "USD": xDev1_display = " $"
    Case "GBP": xDev1_display = " £"
End Select

xDev2_display = " *"
Select Case xYPDCOPE0.PDCOPEDEV2
    Case "EUR":  xDev2_display = " €"
    Case "USD": xDev2_display = " $"
    Case "GBP": xDev2_display = " £"
End Select

'If xYPDCOPE0.PDCOPEDEV1 = "EUR" Then
    wCol_Eur = 1
    wCol_Dev = 2
'Else
'    If xYPDCOPE0.PDCOPEDEV2 = "EUR" Then
'        wCol_Eur = 2
'        wCol_Dev = 1
'    Else
'        wCol_Eur = 1
'        wCol_Dev = 2
'        wBackColor = &HA0D0FF
'    End If
'End If
fgPDCOPE.Col = 0: fgPDCOPE.Text = xYPDCOPE0.PDCOPESENS & "  " & xYPDCOPE0.PDCOPEDEV1 & " / " & xYPDCOPE0.PDCOPEDEV2
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor

fgPDCOPE.Col = wCol_Eur: fgPDCOPE.Text = Format$(wMTD1, "### ### ### ##0.00") & xDev1_display
fgPDCOPE.CellBackColor = wBackColor
If blnAnnulation Then
    fgPDCOPE.CellForeColor = wColor
Else
    If wMTD1 >= 0 Then
         fgPDCOPE.CellForeColor = vbBlue
     Else
         fgPDCOPE.CellForeColor = vbRed
     End If
End If
wMTD2 = -xYPDCOPE0.PDCOPEMTD2
fgPDCOPE.Col = wCol_Dev: fgPDCOPE.Text = Format$(wMTD2, "### ### ### ##0.00") & xDev2_display
fgPDCOPE.CellBackColor = wBackColor
If blnAnnulation Then
    fgPDCOPE.CellForeColor = wColor
Else
    If wMTD2 >= 0 Then
         fgPDCOPE.CellForeColor = vbBlue
     Else
         fgPDCOPE.CellForeColor = vbRed
     End If
End If
fgPDCOPE.Col = 3: fgPDCOPE.Text = Format$(xYPDCOPE0.PDCOPETAUX, "### ##0.000000")
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor
fgPDCOPE.Col = 4: fgPDCOPE.Text = dateImp10(xYPDCOPE0.PDCOPEDVA)
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor
fgPDCOPE.Col = 5: fgPDCOPE.Text = xYPDCOPE0.PDCOPECLI
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor
fgPDCOPE.Col = 6: fgPDCOPE.Text = dateImp10(xYPDCOPE0.PDCOPEDTR) & " - " & xYPDCOPE0.PDCOPEID
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor
fgPDCOPE.Col = 7: fgPDCOPE.Text = xYPDCOPE0.PDCOPEREF
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor
fgPDCOPE.Col = 8: fgPDCOPE.Text = xYPDCOPE0.PDCOPESER & " " & xYPDCOPE0.PDCOPESSE & " " & xYPDCOPE0.PDCOPEOPEC & " " & Trim(xYPDCOPE0.PDCOPEOPET) & " " & xYPDCOPE0.PDCOPEOPEN
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor
fgPDCOPE.Col = 9: fgPDCOPE.Text = xYPDCOPE0.PDCOPESTA & xYPDCOPE0.PDCOPESTA2 & xYPDCOPE0.PDCOPESTA3
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor


fgPDCOPE.Col = fgPDCOPE_arrIndex: fgPDCOPE.Text = lIndex



End Sub
Public Sub fgPDCOPE_ZCHGOPE0_Line(lIndex As Long)
Dim X As String, wColor As Long, wBackColor As Long
Dim blnAnnulation As Boolean
Dim wMTD1 As Currency, wMTD2 As Currency
Dim wCol_Eur As Integer, wCol_Dev As Integer
Dim xDev1_display As String, xDev2_display As String
On Error Resume Next
blnAnnulation = False
wBackColor = &HE0FFFF
wColor = vbBlack
If xZCHGOPE0.CHGOPEVAL = "O" Then wColor = &H6000&
If xZCHGOPE0.CHGOPEANN <> " " Then wColor = &H808080: blnAnnulation = True

If xZCHGOPE0.CHGOPESEN = "A" Then
    wMTD1 = xZCHGOPE0.CHGOPEMO1
    wMTD2 = -xZCHGOPE0.CHGOPEMO2
Else
    wMTD1 = -xZCHGOPE0.CHGOPEMO1
    wMTD2 = xZCHGOPE0.CHGOPEMO2
End If

wBackColor = &HE0FFFF

xDev1_display = " *"
Select Case xZCHGOPE0.CHGOPEDE1
    Case "EUR":  xDev1_display = " €"
    Case "USD": xDev1_display = " $"
    Case "GBP": xDev1_display = " £"
End Select

xDev2_display = " *"
Select Case xZCHGOPE0.CHGOPEDE2
    Case "EUR":  xDev2_display = " €"
    Case "USD": xDev2_display = " $"
    Case "GBP": xDev2_display = " £"
End Select

'If xZCHGOPE0.CHGOPEde1 = "EUR" Then
    wCol_Eur = 1
    wCol_Dev = 2
'Else
'    If xZCHGOPE0.CHGOPEde2 = "EUR" Then
'        wCol_Eur = 2
'        wCol_Dev = 1
'    Else
'        wCol_Eur = 1
'        wCol_Dev = 2
'        wBackColor = &HA0D0FF
'    End If
'End If
fgPDCOPE.Col = 0: fgPDCOPE.Text = xZCHGOPE0.CHGOPESEN & "  " & xZCHGOPE0.CHGOPEDE1 & " / " & xZCHGOPE0.CHGOPEDE2
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor

fgPDCOPE.Col = wCol_Eur: fgPDCOPE.Text = Format$(wMTD1, "### ### ### ##0.00") & xDev1_display
fgPDCOPE.CellBackColor = wBackColor
If blnAnnulation Then
    fgPDCOPE.CellForeColor = wColor
Else
    If wMTD1 >= 0 Then
         fgPDCOPE.CellForeColor = vbBlue
     Else
         fgPDCOPE.CellForeColor = vbRed
     End If
End If

fgPDCOPE.Col = wCol_Dev: fgPDCOPE.Text = Format$(wMTD2, "### ### ### ##0.00") & xDev2_display
fgPDCOPE.CellBackColor = wBackColor
If blnAnnulation Then
    fgPDCOPE.CellForeColor = wColor
Else
    If wMTD2 >= 0 Then
         fgPDCOPE.CellForeColor = vbBlue
     Else
         fgPDCOPE.CellForeColor = vbRed
     End If
End If
fgPDCOPE.Col = 3
If xZCHGOPE0.CHGOPECO3 <> 0 Then
    fgPDCOPE.Text = Format$(xZCHGOPE0.CHGOPECO3, "### ##0.000000")
Else
    fgPDCOPE.Text = Format$(xZCHGOPE0.CHGOPECO1, "### ##0.000000")
End If
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor
fgPDCOPE.Col = 4: fgPDCOPE.Text = dateImp10(xZCHGOPE0.CHGOPEENG + 19000000)
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor
fgPDCOPE.Col = 5: fgPDCOPE.Text = xZCHGOPE0.CHGOPECON
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor
fgPDCOPE.Col = 6: fgPDCOPE.Text = dateImp(xZCHGOPE0.CHGOPECRE + 19000000)
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor
fgPDCOPE.Col = 7: fgPDCOPE.Text = ""
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor
fgPDCOPE.Col = 8: fgPDCOPE.Text = xZCHGOPE0.CHGOPESER & " " & xZCHGOPE0.CHGOPESSE & " " & xZCHGOPE0.CHGOPEOPE & "    " & xZCHGOPE0.CHGOPEDOS
fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor
fgPDCOPE.Col = 9:
If xZCHGOPE0.CHGOPEVAL <> "O" Then
    fgPDCOPE.Text = "non validé"
Else
    If xZCHGOPE0.CHGOPESER <> "TC" Then
        If xZCHGOPE0.CHGOPEMDA <> "MAD" Then
            fgPDCOPE.Text = "non compta"
        End If
    End If
End If

fgPDCOPE.CellForeColor = wColor: fgPDCOPE.CellBackColor = wBackColor


fgPDCOPE.Col = fgPDCOPE_arrIndex: fgPDCOPE.Text = lIndex



End Sub

Public Sub fgTermeEch_DisplayLine_1(lIndex As Long, wCol_Eur As Integer, wMT_Eur As Currency, wCol_Dev As Integer, wMT_Dev As Currency, xCV As String)
Dim X As String, wColor As Long
Dim blnAnnulation As Boolean
Dim wDev As String, kDev As Integer, wAmj As String
Dim wSens As String

On Error Resume Next
'If xZCHGOPE0.CHGOPEOPE = "TER" Then
'    wColor = &HC0E0FF
'Else
    wColor = &HE0FFFF
'End If
If xZCHGOPE0.CHGOPEOPE = "TER" Then
     wSens = xZCHGOPE0.CHGOPESEN
 Else
     If xZCHGOPE0.CHGOPESEN = "A" Then
         wSens = "V"
     Else
         wSens = "A"
     End If
 End If

fgTermeEch.Col = 0: fgTermeEch.Text = wSens & " - " & xZCHGOPE0.CHGOPEDE1 & " / " & xZCHGOPE0.CHGOPEDE2
fgTermeEch.CellBackColor = wColor
If wSens = "V" Then
    fgTermeEch.CellForeColor = vbRed
Else
    fgTermeEch.CellForeColor = vbBlue
End If
fgTermeEch.Col = 1: fgTermeEch.Text = xZCHGOPE0.CHGOPEOPE & "    " & xZCHGOPE0.CHGOPEDOS
fgTermeEch.CellBackColor = wColor
fgTermeEch.Col = 2: fgTermeEch.Text = dateImp(xZCHGOPE0.CHGOPEENG + 19000000)
fgTermeEch.CellBackColor = wColor
wAmj = xZCHGOPE0.CHGOPEDT1 + 19000000
fgTermeEch.Col = 3: fgTermeEch.Text = dateImp(wAmj)
fgTermeEch.CellForeColor = vbBlue
If wAmj > YBIATAB0_DATE_CPT_J And wAmj < mPDCOPEDVA_5J Then
    fgTermeEch.CellBackColor = &HFFA0FF   '&HFFC0FF
Else
    fgTermeEch.CellBackColor = wColor
End If
fgTermeEch.Col = 4: fgTermeEch.CellBackColor = wColor

If xZCHGOPE0.CHGOPEDE1 = "EUR" Then
    fgTermeEch.Text = Format$(xZCHGOPE0.CHGOPECO2, "###.00000")
Else
    fgTermeEch.Text = Format$(xZCHGOPE0.CHGOPECO4, "###.00000")
End If



fgTermeEch.Col = 8:  fgTermeEch.CellBackColor = wColor

fgTermeEch.Col = wCol_Eur: fgTermeEch.Text = Format$(wMT_Eur, "### ### ### ##0.00")
fgTermeEch.Col = wCol_Dev: fgTermeEch.Text = Format$(wMT_Dev, "### ### ### ##0.00")
fgTermeEch.Col = 6: fgTermeEch.CellForeColor = vbRed: fgTermeEch.CellBackColor = wColor
fgTermeEch.Col = 7: fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
fgTermeEch.Col = 9: fgTermeEch.CellForeColor = vbRed: fgTermeEch.CellBackColor = wColor
fgTermeEch.Col = 10: fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor


fgTermeEch.Col = 11: fgTermeEch.Text = xCV: fgTermeEch.CellBackColor = &HA0D0FF
fgTermeEch.CellForeColor = IIf(wCol_Eur = 6, vbRed, vbBlue)
fgTermeEch.Col = fgTermeEch_arrIndex: fgTermeEch.Text = lIndex



End Sub


Public Sub fgDetail_DisplayLine_YPDCMVT0(lIndex As Long)
Dim devFixing As Currency, devPP As Currency, eurPP As Currency
Dim wPDCMVTMTE As Currency, wPDCMVTMTD As Currency
Dim wCellBackColor As Long
Dim K As Integer
On Error Resume Next
Select Case xYPDCMVT0.PDCMVTOPEC
    Case "TER": wCellBackColor = &HA0C0FF
    Case "SWP": wCellBackColor = &HA0FFFF
    Case Else: wCellBackColor = &HE0FFFF    '&HFE0F0F0    '&HFFFFFF
End Select
If xYPDCMVT0.PDCMVTKCUT <> " " Then wCellBackColor = &HFFA0FF
fgDetail.Col = 0: fgDetail.Text = dateImp10(xYPDCMVT0.PDCMVTDTR)
fgDetail.CellBackColor = &HD0D0D0

wPDCMVTMTE = -xYPDCMVT0.PDCMVTMTE
fgDetail.Col = 1: fgDetail.Text = Format$(wPDCMVTMTE, "### ### ### ##0.00")
fgDetail.CellBackColor = wCellBackColor '&HE0E0E0
If wPDCMVTMTE >= 0 Then
     fgDetail.CellForeColor = vbBlue
 Else
     fgDetail.CellForeColor = vbRed
 End If

wPDCMVTMTD = -xYPDCMVT0.PDCMVTMTD
fgDetail.Col = 2: fgDetail.Text = Format$(wPDCMVTMTD, "### ### ### ##0.00")
fgDetail.CellBackColor = wCellBackColor '&HE0E0E0
If wPDCMVTMTD > 0 Then
     fgDetail.CellForeColor = vbBlue
 Else
     fgDetail.CellForeColor = vbRed
 End If
fgDetail.Col = 3: fgDetail.Text = Format$(xYPDCMVT0.PDCMVTTAUX, "### ##0.000000")
fgDetail.CellBackColor = wCellBackColor

fgDetail.Col = 4: fgDetail.Text = dateImp10(xYPDCMVT0.PDCMVTDVA) & "  "
fgDetail.CellBackColor = wCellBackColor
fgDetail.Col = 5: fgDetail.Text = " " & xYPDCMVT0.PDCMVTDEV
fgDetail.CellBackColor = wCellBackColor
fgDetail.Col = 6: fgDetail.Text = xYPDCMVT0.PDCMVTSTA2 & " " & xYPDCMVT0.PDCMVTSER & "." & xYPDCMVT0.PDCMVTSSE & "  " & xYPDCMVT0.PDCMVTOPEC & "   " & xYPDCMVT0.PDCMVTOPEN
fgDetail.CellBackColor = wCellBackColor
fgDetail.Col = 7
If xYPDCMVT0.PDCMVTKCUT = " " Then
    fgDetail.Text = xYPDCMVT0.PDCMVTCLI
    fgDetail.CellBackColor = wCellBackColor
Else
    fgDetail.CellBackColor = &HE0FFFF
    For K = 1 To arrDev_Nb
        If arrKCUT(K).PDCPOSDEV = xYPDCMVT0.PDCMVTDEV Then
            fgDetail.Text = "cut " & arrKCUT(K).PDCPOSPRIX
            Exit For
        End If
    Next K
End If
'    fgDetail.CellBackColor = wCellBackColor
fgDetail.Col = 8: fgDetail.Text = xYPDCMVT0.PDCMVTSTA & " " & xYPDCMVT0.PDCMVTPIE & "-" & xYPDCMVT0.PDCMVTECR
fgDetail.CellBackColor = wCellBackColor
fgDetail.Col = 9: fgDetail.Text = xYPDCMVT0.PDCMVTCPT
fgDetail.CellBackColor = wCellBackColor



fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex



End Sub

Public Sub fgDetail_DisplayLine_YPDCLOG0(lIndex As Long)

On Error Resume Next

fgDetail.Col = 0: fgDetail.Text = dateImp10(xYPDCLOG0.PDCLOGDTR)
fgDetail.CellBackColor = &HD0D0D0

fgDetail.Col = 1: fgDetail.Text = dateImp10(xYPDCLOG0.PDCLOGUAMJ) & "  " & timeImp8(xYPDCLOG0.PDCLOGUHMS) & " - " & xYPDCLOG0.PDCLOGUSEQ
fgDetail.Col = 2: fgDetail.Text = xYPDCLOG0.PDCLOGSTA & " " & xYPDCLOG0.PDCLOGNAT
fgDetail.Col = 3: fgDetail.Text = xYPDCLOG0.PDCLOGTXT
fgDetail.CellForeColor = vbBlue
Select Case Mid$(xYPDCLOG0.PDCLOGNAT, 3, 1)
    Case " "
    Case "!": fgDetail.CellForeColor = vbMagenta
    Case Else: fgDetail.CellForeColor = vbRed
End Select
Select Case Mid$(xYPDCLOG0.PDCLOGNAT, 1, 1)
    Case "4", "7": fgDetail.CellForeColor = vbMagenta
End Select
If Mid$(xYPDCLOG0.PDCLOGNAT, 1, 2) = "5=" Then fgDetail.CellForeColor = vbMagenta
If xYPDCLOG0.PDCLOGPIE <> 0 Then fgDetail.Col = 4: fgDetail.Text = xYPDCLOG0.PDCLOGPIE & "-" & xYPDCLOG0.PDCLOGECR
fgDetail.Col = 5: fgDetail.Text = xYPDCLOG0.PDCLOGUUSR



fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex



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
Public Sub fgTerme_Reset()
fgTerme.Clear
fgTerme_Sort1 = 0: fgTerme_Sort2 = 0
fgTerme_Sort1_Old = -1
fgTerme_RowDisplay = 0: fgTerme_RowClick = 0
fgTerme_arrIndex = fgTerme.Cols - 1
blnfgTerme_DisplayLine = False
fgTerme_SortAD = 6
fgTerme.LeftCol = 0

End Sub

Public Sub fgPDCOPE_Reset()
fgPDCOPE.Clear
fgPDCOPE_Sort1 = 0: fgPDCOPE_Sort2 = 0
fgPDCOPE_Sort1_Old = -1
fgPDCOPE_RowDisplay = 0: fgPDCOPE_RowClick = 0
fgPDCOPE_arrIndex = fgPDCOPE.Cols - 1
blnfgPDCOPE_DisplayLine = False
fgPDCOPE_SortAD = 6
fgPDCOPE.LeftCol = 0

End Sub

Public Sub fgTermeEch_Reset()
fgTermeEch.Clear
fgTermeEch_Sort1 = 0: fgTermeEch_Sort2 = 0
fgTermeEch_Sort1_Old = -1
fgTermeEch_RowDisplay = 0: fgTermeEch_RowClick = 0
fgTermeEch_arrIndex = fgTermeEch.Cols - 1
blnfgTermeEch_DisplayLine = False
fgTermeEch_SortAD = 6
fgTermeEch.LeftCol = 0

End Sub


Public Sub fgDetail_Reset()
fgDetail.Clear
fgDetail_Sort1 = 0: fgDetail_Sort2 = 0
fgDetail_Sort1_Old = -1
fgDetail_RowDisplay = 0: fgDetail_RowClick = 0
fgDetail_arrIndex = fgDetail.Cols - 1
blnfgDetail_DisplayLine = False
fgDetail_SortAD = 6
fgDetail.LeftCol = 0

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
Public Sub fgTerme_Sort()
If fgTerme.Rows > 1 Then
    fgTerme.Row = 1
    fgTerme.RowSel = fgTerme.Rows - 1
    
    If fgTerme_Sort1_Old = fgTerme_Sort1 Then
        If fgTerme_SortAD = 5 Then
            fgTerme_SortAD = 6
        Else
            fgTerme_SortAD = 5
        End If
    Else
        fgTerme_SortAD = 5
    End If
    fgTerme_Sort1_Old = fgTerme_Sort1
    
    fgTerme.Col = fgTerme_Sort1
    fgTerme.ColSel = fgTerme_Sort2
    fgTerme.Sort = fgTerme_SortAD
End If

End Sub

Public Sub fgPDCOPE_Sort()
If fgPDCOPE.Rows > 1 Then
    fgPDCOPE.Row = 1
    fgPDCOPE.RowSel = fgPDCOPE.Rows - 1
    
    If fgPDCOPE_Sort1_Old = fgPDCOPE_Sort1 Then
        If fgPDCOPE_SortAD = 5 Then
            fgPDCOPE_SortAD = 6
        Else
            fgPDCOPE_SortAD = 5
        End If
    Else
        fgPDCOPE_SortAD = 5
    End If
    fgPDCOPE_Sort1_Old = fgPDCOPE_Sort1
    
    fgPDCOPE.Col = fgPDCOPE_Sort1
    fgPDCOPE.ColSel = fgPDCOPE_Sort2
    fgPDCOPE.Sort = fgPDCOPE_SortAD
End If

End Sub
Public Sub fgTermeEch_Sort()
If fgTermeEch.Rows > 1 Then
    fgTermeEch.Row = 1
    fgTermeEch.RowSel = fgTermeEch.Rows - 1
    
    If fgTermeEch_Sort1_Old = fgTermeEch_Sort1 Then
        If fgTermeEch_SortAD = 5 Then
            fgTermeEch_SortAD = 6
        Else
            fgTermeEch_SortAD = 5
        End If
    Else
        fgTermeEch_SortAD = 5
    End If
    fgTermeEch_Sort1_Old = fgTermeEch_Sort1
    
    fgTermeEch.Col = fgTermeEch_Sort1
    fgTermeEch.ColSel = fgTermeEch_Sort2
    fgTermeEch.Sort = fgTermeEch_SortAD
End If

End Sub


Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    If lK = 2 Then
        fgSelect.Col = 2
        X = fgSelect.Text
    Else
        X = ""
    End If
    
    fgSelect.Col = 3
    X = X & Format$(Val(fgSelect.Text), "000000000000000.00")
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub
Public Sub fgTerme_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgTerme.Rows - 1
    fgTerme.Row = I
    If lK = 2 Then
        fgTerme.Col = 2
        X = fgTerme.Text
    Else
        X = ""
    End If
    
    fgTerme.Col = 3
    X = X & Format$(Val(fgTerme.Text), "000000000000000.00")
    fgTerme.Col = fgTerme_arrIndex - 1
    fgTerme.Text = X
Next I


fgTerme_Sort1 = fgTerme_arrIndex - 1: fgTerme_Sort2 = fgTerme_arrIndex - 1
fgTerme_Sort
End Sub


Public Sub fgPDCOPE_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgPDCOPE.Rows - 1
    fgPDCOPE.Row = I
    fgPDCOPE.Col = fgPDCOPE_arrIndex:  arrYPDCOPE0_Index = CLng(fgPDCOPE.Text)
    Select Case lK
        Case 1: fgPDCOPE.Col = 1: X = Format$(num_CDec(fgPDCOPE.Text), "000000000000000.00")
        Case 2: fgPDCOPE.Col = 2: X = Format$(num_CDec(fgPDCOPE.Text), "000000000000000.00")
        Case 3: X = Format$(arrYPDCOPE0(arrYPDCOPE0_Index).PDCOPETAUX, "0000000.000000000")
        Case 4: X = arrYPDCOPE0(arrYPDCOPE0_Index).PDCOPEDVA
        Case 6: X = Format(arrYPDCOPE0(arrYPDCOPE0_Index).PDCOPEID, "000000000")
        Case 7: X = Format(arrYPDCOPE0(arrYPDCOPE0_Index).PDCOPEREF, "000000000")
        Case 8: X = Format(arrYPDCOPE0(arrYPDCOPE0_Index).PDCOPEOPEN, "000000000")
    End Select
    fgPDCOPE.Col = fgPDCOPE_arrIndex - 1
    fgPDCOPE.Text = X
Next I


fgPDCOPE_Sort1 = fgPDCOPE_arrIndex - 1: fgPDCOPE_Sort2 = fgPDCOPE_arrIndex - 1
fgPDCOPE_Sort

End Sub


Public Sub fgTermeEch_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgTermeEch.Rows - 1
    fgTermeEch.Row = I
    fgTermeEch.Col = fgTermeEch_arrIndex:  arrYPDCOPE0_Index = CLng(fgTermeEch.Text)
    Select Case lK
       ' Case 1: fgTermeEch.Col = 1: X = Format$(num_CDec(fgTermeEch.Text), "000000000000000.00")
    End Select
    fgTermeEch.Col = fgTermeEch_arrIndex - 1
    fgTermeEch.Text = X
Next I


fgTermeEch_Sort1 = fgTermeEch_arrIndex - 1: fgTermeEch_Sort2 = fgTermeEch_arrIndex - 1
fgTermeEch_Sort

End Sub



Public Sub fgDetail_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgDetail.Rows - 1
    fgDetail.Row = I
    fgDetail.Col = fgDetail_arrIndex:  arrYPDCMVT0_Index = CLng(fgDetail.Text)
    Select Case lK
        Case 0: X = arrYPDCMVT0(arrYPDCMVT0_Index).PDCMVTDTR
        Case 1:  X = Format$(arrYPDCMVT0(arrYPDCMVT0_Index).PDCMVTMTE, "000000000000000.00")
        Case 2:  X = Format$(arrYPDCMVT0(arrYPDCMVT0_Index).PDCMVTMTD, "000000000000000.00")
        Case 3: X = Format$(arrYPDCMVT0(arrYPDCMVT0_Index).PDCMVTTAUX, "0000000.000000000")
        Case 4: X = arrYPDCMVT0(arrYPDCMVT0_Index).PDCMVTDVA
        Case 6: X = Format(arrYPDCMVT0(arrYPDCMVT0_Index).PDCMVTOPEN, "000000000")
        Case 8: X = Format(arrYPDCMVT0(arrYPDCMVT0_Index).PDCMVTPIE, "000000000") & Format(arrYPDCMVT0(arrYPDCMVT0_Index).PDCMVTECR, "000000")
    End Select
    fgDetail.Col = fgDetail_arrIndex - 1
    fgDetail.Text = X
Next I


fgDetail_Sort1 = fgDetail_arrIndex - 1: fgDetail_Sort2 = fgDetail_arrIndex - 1
fgDetail_Sort
End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(Msg As String)

mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), BIA_PDC_Aut)

localUnit = Trim(idemUser.Unit)

paramIBM_Library_SABXXX = paramIBM_Library_SABSPE

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'MsgBox "JPL : lecture PROD / màj PDC test  ", vbCritical
'YBIATAB0_DATE_CPT_J = "20151103"
'YBIATAB0_DATE_CPT_JS1 = "20151104"

'paramIBM_Library_SABSPE_XXX = "SAB073USPE"
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! ne pas déplacer
Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@BIA_PDC": blnAuto = True
    Case Else: blnAuto = False
End Select

Form_Init
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! ne pas déplacer

Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@BIA_PDC": blnAuto = True
                    blnControl = False
                    Call cbo_Scan("5", cboSelect_SQL)
                    blnControl = True
                    cmdSelect_SQL_K = "5"
                    cmdSelect_Ok_Click
                    If YBIATAB0_DATE_CPT_JP0 <> YBIATAB0_DATE_CPT_J Then
                        Call DTPicker_Set(txtSelect_AMJ, YBIATAB0_DATE_CPT_JP0)
                        If xlsManual Then
                            Call cmdSendMail_BIA_PDC_xlsManual
                        Else
                            cmdSendMail_BIA_PDC
                        End If
                    End If
                    Call DTPicker_Set(txtSelect_AMJ, YBIATAB0_DATE_CPT_J)
                    If xlsManual Then
                        Call cmdSendMail_BIA_PDC_xlsManual
                    Else
                        cmdSendMail_BIA_PDC
                    End If
                    Unload Me
           
    Case Else: blnAuto = False
                    

End Select


End Sub


Public Sub Form_Init()
Dim X As String, xSQL As String
Dim xMemo As String, xDev As String
Dim blnPDC_Cours_Connu As Boolean, K As Integer
Dim blnOk As Boolean
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
blnControl = False

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
DevF_Load

''$JPL 20-01-2015 ajout CNY
'arrDev_Nb = 11 '10 '9
'arrDev = Array("   ", "AED", "ARS", "CAD", "CHF", "DKK", "GBP", "JPY", "SEK", "USD", "KWD", "CNY")
'arrDev_Row = Array(0, 9, 7, 3, 5, 6, 2, 4, 8, 1, 10, 11)

'$JPL 03-10-2015 suppression JPY
'$JPL 03-10-2015 delete last YPDCPOS0
arrDev_Nb = 11 '11 '10 '9
arrDev = Array("   ", "AED", "ARS", "CAD", "CHF", "DKK", "GBP", "SEK", "USD", "KWD", "CNY", "SAR")
arrDev_Row = Array(0, 8, 6, 3, 4, 5, 2, 7, 1, 9, 10, 11)

ReDim arrTerme_DB(arrDev_Nb), arrTerme_CR(arrDev_Nb), arrKCUT(arrDev_Nb), arrSuspens_Dev(arrDev_Nb), arrSWP_Dev(arrDev_Nb)
ReDim fixingJ_1(arrDev_Nb)
ReDim arrPPJ(arrDev_Nb)
blnPDC_Cours_Connu = True
For K = 0 To arrDev_Nb
    rsYPDCPOS0_Init fixingJ_1(K)
    arrPPJ(K) = 0
Next K

xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YBIATAB0 where BIATABID = 'PDC' and BIATABK2 = '" & YBIATAB0_DATE_CPT_J & "' order by BIATABK1"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    xDev = Trim(rsSab("BIATABK1"))
    xMemo = rsSab("BIATABTXT")
    For K = 1 To arrDev_Nb
        If xDev = arrDev(K) Then
            fixingJ_1(K).PDCPOSDEV = xDev
            If IsNumeric(Mid$(xMemo, 9, 15)) Then
                fixingJ_1(K).PDCPOSFIXT = CDbl(Mid$(xMemo, 9, 15) / 1000000000)
                fixingJ_1(K).PDCPOSFIXD = Mid$(xMemo, 1, 8)
            End If
            Exit For
        End If
    Next K
    
    V = rsYPDCPOS0_GetBuffer(rsSab, xYPDCPOS0)
    rsSab.MoveNext
    
Loop

blnPDC_Cours_Connu = True
X = "BIA_PDC : Cours manquants au " & dateImp10_S(YBIATAB0_DATE_CPT_J) _
  & vbCrLf & vbCrLf & "(SAB073SPE/YBIATAB0 : PDC|<devise>|" & YBIATAB0_DATE_CPT_J & ")" & vbCrLf & vbCrLf

For K = 1 To arrDev_Nb
    If fixingJ_1(K).PDCPOSFIXT = 0 Then
        blnPDC_Cours_Connu = False
        X = X & arrDev(K) & vbCrLf
    End If
Next K

If Not blnPDC_Cours_Connu Then
    If Not blnAuto Then
        Call MsgBox(X, vbCritical, "BIA_PDC")
        BIA_PDC_Aut.Saisir = False
        BIA_PDC_Aut.Valider = False
        BIA_PDC_Aut.Rapprocher = False
        BIA_PDC_Aut.Comptabiliser = False
        BIA_PDC_Aut.Xspécial = False
    Else
        Call cmdSendMail_Alerte(X)
        Unload Me
   End If
End If
'____________________________________________________________
If paramEnvironnement = constProduction Then
    blnOk = False
    xSQL = "select MONFILE from " & paramIBM_Library_SABSPE_XXX & ".YBIAMON7 where MONAPP = 'COMPTA' and MONFLUX = '@BIA_PDC'"
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        If YBIATAB0_DATE_CPT_J <= Trim(rsSab("MONFILE")) Then blnOk = True
    End If
    If Not blnOk Then
        If Not blnAuto Then
            Call MsgBox("Le traitement automatique n'a pas été effectué en date du " & dateImp10(YBIATAB0_DATE_CPT_J) _
                        & vbCrLf & vbCrLf & "(SAB073SPE/YBIAMON7 : COMPTA|@BIA_PDC = " & rsSab("MONFILE") & ")", vbCritical, "BIA_PDC")
            BIA_PDC_Aut.Saisir = False
            BIA_PDC_Aut.Valider = False
            BIA_PDC_Aut.Rapprocher = False
            BIA_PDC_Aut.Comptabiliser = False
            BIA_PDC_Aut.Xspécial = False
        End If
    End If
End If
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DTPicker_Set(txtSelect_AMJ, YBIATAB0_DATE_CPT_JS1)

fgSelect_FormatString = fgSelect.FormatString
fgSuspens_Log_FormatString = fgSuspens_Log.FormatString

fgSelect.Visible = False
fgDetail_FormatString = fgDetail.FormatString
mfgDetail_Top = fgDetail.Top: mfgDetail_Height = fgDetail.Height
fgYPDCLOG0_FormatString = "<Date Compta   |<mise à jour le                               |<Nature|<libellé                                                                                                                          |<Pièce                  |<Màj par                |||||"
mPDCLOGUSEQ = 0
fgDetail.Visible = False
fgTerme_FormatString = fgTerme.FormatString
fgTerme.Visible = False
fgTermeEch_FormatString = fgTermeEch.FormatString
fgTermeEch.Visible = False: fgTermeEch.Top = 360: fgTermeEch.Height = 7560
fraPrint.Visible = False: fraPrint.Top = 480: fraPrint.Left = 9480

fgPDCOPE_FormatString = fgPDCOPE.FormatString
fgPDCOPE.Visible = False
Set fgPDCOPE.Container = fraTab0
fgPDCOPE.Top = fgDetail.Top
fgPDCOPE.Left = fgDetail.Left
fgPDCOPE.Height = fgDetail.Height
fraPDCOPE.Visible = False
Set fraPDCOPE.Container = fraTab0
fraPDCOPE.Top = fgSelect.Top
fraPDCOPE.Left = fgSelect.Left + fgSelect.Width - fraPDCOPE.Width - 300
fraPDCOPE.ZOrder 0

fraSuspens.Visible = False
Set fraSuspens.Container = fraTab0
fraSuspens.Top = fgSelect.Top
fraSuspens.Left = fgSelect.Left + fgSelect.Width - fraSuspens.Width
fraSuspens_Options.Visible = False
Set fraSuspens_Options.Container = fraTab0

fraReport.Visible = False
Set fraReport.Container = fraTab0
fraReport.Top = fgSelect.Top
fraReport.Left = fgSelect.Left

fraTab1.Visible = BIA_PDC_Aut.Avis
mPDCOPEDVA_2J = dateAdd_On("EUR", 2, YBIATAB0_DATE_CPT_JS1)
mPDCOPEDVA_5J = dateAdd_On("EUR", 5, YBIATAB0_DATE_CPT_JS1)

X = dateFinDeMois(YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtSelect_AMJ_HB_xls, X): txtSelect_AMJ_HB_xls.Enabled = False
Call DTPicker_Set(txtSelect_AMJ_HB, X): txtSelect_AMJ_HB.Enabled = False


fraSelect_Options_xls.Visible = False
Set fraSelect_Options_xls.Container = fraSelect_Options.Container
fraSelect_Options_xls.Top = fraSelect_Options.Top
fraSelect_Options_xls.Left = fraSelect_Options.Left
Call DTPicker_Set(txtSelect_AMJ_xls, YBIATAB0_DATE_CPT_J)
chkSelect_Suspens_Out_xls = "1"

fraSelect_Comment_Xls.Visible = False
Set fraSelect_Comment_Xls.Container = fraTab0
fraSelect_Comment_Xls.Top = fgSelect.Top
fraSelect_Comment_Xls.Left = fgSelect.Left + fgSelect.Width - fraSelect_Comment_Xls.Width - 3000

fraSelect_Options_Y.Visible = False
Set fraSelect_Options_Y.Container = fraSelect_Options.Container
fraSelect_Options_Y.Top = fraSelect_Options.Top
fraSelect_Options_Y.Left = fraSelect_Options.Left
Call DTPicker_Set(txtSelect_AMJ_Y, YBIATAB0_DATE_CPT_J)

'_____________________________________________________________
blnPDCOPE_Control_V = BIA_PDC_Aut.Valider
cmdPDCOPE_New.Visible = BIA_PDC_Aut.Saisir
mRecipient_FOTC = srvSendMail.Exchange_Distribution("BIA_PDC", "COURS")
If Trim(mRecipient_FOTC) = "" Then
    Call MsgBox("Table vbsendmail non renseignée pour : BIA_PDC | COURS" & vbCrLf & "FOTC@bia-paris.fr par défaut", vbCritical, "BIA_PDC : gestion des opérations")
    mRecipient_FOTC = "FOTC@bia-paris.fr"
End If
mRecipient_BOTC = srvSendMail.Exchange_Distribution("BIA_PDC", "BOTC")
mRecipient_CONF_CALL = srvSendMail.Exchange_Distribution("BIA_PDC", "CONF_CALL")


cboPDCOPEOPET.Clear
cboPDCOPEOPET.AddItem "   "
blnPDCOPE_CONF_CALL_Visible = False
Select Case localUnit
    Case "FOTC": blnPDCOPE_CONF_CALL_Visible = BIA_PDC_Aut.Saisir
            cmdPDCOPE_CONF_CALL.Visible = blnPDCOPE_CONF_CALL_Visible
    Case "GDC": mSQL_Unit = "" '" and PDCOPESSE = 'TC' and PDCOPESTA3 = 'B'"
            blnPDCOPE_Control_V = BIA_PDC_Aut.Rapprocher
            BIA_PDC_Aut.Saisir = True
            cmdPDCOPE_New.Caption = "Contrôle opération TC"
            cmdPDCOPE_New.Visible = BIA_PDC_Aut.Rapprocher

    Case "GDMP": mSQL_Unit = " and PDCOPESSE in ('GU','TR')"
    Case "SOBI": mSQL_Unit = " and PDCOPESSE = 'CD'"
    Case Else: mSQL_Unit = ""
End Select

fraPDC_Param.Visible = True

cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1 - position de change"
If BIA_PDC_Aut.Rapprocher Then
    cboSelect_SQL.AddItem "3 - contrôle des tickets TC"
    cboSelect_SQL.AddItem "Xls - contrôle du tableau d'arbitrage"
End If
If BIA_PDC_Aut.Valider Then
    cboSelect_SQL.AddItem "4 - gestion des suspens FOTC"
    fraPDC_Param.Visible = True
End If

If BIA_PDC_Aut.Avis Then cboSelect_SQL.AddItem "Y - Duplicata du contrôle du tableau d'arbitrage"

'If BIA_PDC_Aut.Comptabiliser Then cboSelect_SQL.AddItem "5 - Traitement journée comptable"
'If BIA_PDC_Aut.Saisir Then cboSelect_SQL.AddItem "2 - Liste des opérations"
If BIA_PDC_Aut.Xspécial Then
    cboSelect_SQL.AddItem "5 - Traitement journée comptable"
    cboSelect_SQL.AddItem "7 - Annulation POS + MVT >= date"
    cboSelect_SQL.AddItem "9 - Initialisation PDC au 01-01-2009"
End If

cboSelect_SQL.ListIndex = 0

cboSelect_Devise.Clear
cboSelect_Devise.AddItem "   "
cboPDCOPEDEV1.Clear
cboPDCOPEDEV1.AddItem "   "
cboPDCOPEDEV1.AddItem "EUR"
cboPDCOPEDEV2.Clear
cboPDCOPEDEV2.AddItem "   "
cboPDCOPEDEV2.AddItem "EUR"
cboSuspens_PDCMVTDEV1.Clear
cboSuspens_PDCMVTDEV1.AddItem "EUR"
cboSuspens_PDCMVTDEV2.Clear
cboSuspens_PDCMVTDEV2.AddItem "EUR"
For I = 1 To arrDev_Nb
    X = arrDev(arrDev_Row(I))
    cboSelect_Devise.AddItem X
    cboPDCOPEDEV1.AddItem X
    cboPDCOPEDEV2.AddItem X
    cboSuspens_PDCMVTDEV1.AddItem X
    cboSuspens_PDCMVTDEV2.AddItem X
Next I
cboSelect_Devise.ListIndex = 0

blnHB = True

fgSelect.Enabled = True

cboPDCOPESER.Clear
cboPDCOPESER.AddItem "00 00"
cboPDCOPESER.AddItem "00 CD"
cboPDCOPESER.AddItem "00 GU"
cboPDCOPESER.AddItem "00 MP"
cboPDCOPESER.AddItem "00 TR"
cboPDCOPESER.AddItem "CP CP"
cboPDCOPESER.AddItem "TC TC"

cboSuspens_PDCMVTSER.Clear
cboSuspens_PDCMVTSER.AddItem "00 00"
cboSuspens_PDCMVTSER.AddItem "00 CD"
cboSuspens_PDCMVTSER.AddItem "00 GU"
cboSuspens_PDCMVTSER.AddItem "00 MP"
cboSuspens_PDCMVTSER.AddItem "00 TR"
cboSuspens_PDCMVTSER.AddItem "CP CP"
cboSuspens_PDCMVTSER.AddItem "TC TC"

cboPDCOPESENS.Clear
cboPDCOPEOPEC.Clear
If localUnit = "SOBI" Then
    cboPDCOPESENS.AddItem "Vendue au client par la BIA"
    cboPDCOPESENS.Enabled = False
    lblPDCOPEDEV1.Caption = "Devise de paiement"
    lblPDCOPEDEV2.Caption = "Devise du débit client"
    libPDCOPETAUX.Visible = False
    chkPDCOPEINFO.Top = libPDCOPETAUX.Top
    chkPDCOPEINFO.Left = libPDCOPETAUX.Left
    cboPDCOPEOPEC.AddItem "CDE"
    cboPDCOPEOPEC.AddItem "CDI"

Else
    cboPDCOPESENS.AddItem "Achetée par la BIA"
    cboPDCOPESENS.AddItem "Vendue  par la BIA"
    cboPDCOPEOPEC.AddItem "CPT"
    cboPDCOPEOPEC.AddItem "TER"
    cboPDCOPEOPEC.AddItem "SWP"
    cboPDCOPEOPEC.AddItem "CDE"
    cboPDCOPEOPEC.AddItem "CDI"
    cboPDCOPEOPEC.AddItem "OD "
End If

cboSuspens_PDCMVTSENS.Clear
cboSuspens_PDCMVTSENS.AddItem "Achetée par la BIA"
cboSuspens_PDCMVTSENS.AddItem "Vendue par la BIA"

mCLIENACLI = "": mCLIENARA1 = ""
mFIXING_DEV = "":  mFixing_AMJ = "": mFIXING_Cours = 0

Call sqlYBIATAB0_Read("PDC_PARAM", "TER", "382100", xMemo)
If IsNumeric(Mid$(xMemo, 1, 8)) Then
    mTER382100_Amj = Mid$(xMemo, 1, 8)
Else
    mTER382100_Amj = "99999999"
End If
Call DTPicker_Set(txtPDC_Param, mTER382100_Amj)

cmdReset

Me.Enabled = True
Me.MousePointer = 0
End Sub





Private Sub cboPDCOPEDEV1_Click()
fraPDCOPE_Control_DEV_Certain
lblPDCOPEDEV1.BackColor = fraPDCOPE_S.BackColor ' &HC0FFFF

End Sub

Private Sub cboPDCOPEDEV1_GotFocus()
Call txt_GotFocus(cboPDCOPEDEV1)
End Sub


Private Sub cboPDCOPEDEV1_LostFocus()
Call txt_LostFocus(cboPDCOPEDEV1)


End Sub


Private Sub cboPDCOPEDEV2_Click()
fraPDCOPE_Control_DEV_Certain

If localUnit <> "SOBI" Then lblPDCOPEDEV2.BackColor = fraPDCOPE_S.BackColor '&HC0FFFF

End Sub

Private Sub cboPDCOPEDEV2_GotFocus()
Call txt_GotFocus(cboPDCOPEDEV2)
End Sub


Private Sub cboPDCOPEDEV2_LostFocus()
Call txt_LostFocus(cboPDCOPEDEV2)

End Sub


Private Sub cboPDCOPEOPEC_Click()
If cboPDCOPEOPEC.Text = "SWP" Then
    txtPDCOPEFIXING.Enabled = True
    lblPDCOPEFIXING.Caption = "cours à terme"
    txtPDCOPEVTXT.Enabled = True
    txtPDCOPEVTXT.Locked = False
    txtPDCOPEVTXT.FontBold = True
Else
    txtPDCOPEFIXING.Enabled = False
    lblPDCOPEFIXING.Caption = "fixing"
    txtPDCOPEFIXING = ""
    txtPDCOPEVTXT = ""
End If

End Sub

Private Sub cboPDCOPEOPEC_GotFocus()
Call txt_GotFocus(cboPDCOPEOPEC)

End Sub

Private Sub cboPDCOPEOPEC_LostFocus()
Call txt_LostFocus(cboPDCOPEOPEC)

End Sub


Private Sub cboPDCOPEOPET_GotFocus()
Call txt_GotFocus(cboPDCOPEOPET)

End Sub


Private Sub cboPDCOPEOPET_LostFocus()
Call txt_LostFocus(cboPDCOPEOPET)

End Sub


Private Sub cboPDCOPESENS_Click()
fraPDCOPE_Display_Montant Mid$(cboPDCOPESENS, 1, 1)

End Sub

Private Sub cboPDCOPESENS_GotFocus()
Call txt_GotFocus(cboPDCOPESENS)

End Sub


Private Sub cboPDCOPESENS_LostFocus()
Call txt_LostFocus(cboPDCOPESENS)

End Sub


Private Sub cboPDCOPESER_GotFocus()
Call txt_GotFocus(cboPDCOPESER)

End Sub


Private Sub cboPDCOPESER_LostFocus()
Call txt_LostFocus(cboPDCOPESER)

End Sub


Private Sub cboSelect_Devise_Click()
cmdSelect_Reset
If Trim(cboSelect_Devise.Text) = "" Then
    chkSelect_HB.Visible = True
    chkSelect_Suspens_Out.Visible = True
    chkSelect_PDCMVTKCUT.Visible = True
    chkSelect_Log.Visible = True
Else
    chkSelect_HB.Visible = False
    chkSelect_HB.Value = "0"
    chkSelect_Suspens_Out.Visible = False
    chkSelect_Suspens_Out.Value = "0"
    chkSelect_PDCMVTKCUT.Visible = False
    chkSelect_PDCMVTKCUT.Value = "0"
    chkSelect_Log.Visible = False
    chkSelect_Log.Value = "0"
    chkSelect_Ope.Visible = False
    chkSelect_Ope.Value = "0"
    
    chkSelect_Suspens = False
    chkSelect_Suspens = "0"
    chkSelect_ZCHGOPE0 = False
    chkSelect_ZCHGOPE0 = "0"
    chkSelect_ZCHGOPE0_NC = False
    chkSelect_ZCHGOPE0_NC = "0"

End If
End Sub


Private Sub cboSelect_SQL_Click()
cmdSelect_Reset
 End Sub



'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lblSelect_AMJ.ForeColor = vbRed
lblSelect_Devise.ForeColor = vbBlue
chkSelect_HB.ForeColor = vbRed: chkPrint_Exclure_HB.ForeColor = vbMagenta
chkSelect_Terme.ForeColor = vbRed
chkSelect_Suspens_Out.ForeColor = vbMagenta: chkPrint_Suspens_Out.ForeColor = vbMagenta
chkSelect_PDCMVTKCUT.ForeColor = vbMagenta: chkPrint_Exclure_PDCMVTKCUT.ForeColor = vbMagenta
chkSelect_Suspens.ForeColor = vbMagenta
chkSelect_Log.ForeColor = vbBlue
chkSelect_Ope.ForeColor = vbBlue
libPDCOPESTA.ForeColor = vbMagenta
lblPDCOPESTA2.ForeColor = vbMagenta
lblPDCOPESTA3.ForeColor = vbMagenta
libPDCOPETAUX.ForeColor = vbMagenta
'lstErr.Visible = False
currentAction = ""
fraSelect_Options.Enabled = True
'cmdSelect_Ok_Click
fraPDCOPE_Options.Enabled = True
SSTab1.Tab = 0


blnControl = True
chkSelect_Ope = "1"
'cboSelect_SQL_Click
cmdSelect_SQL_2X
SSTab1.Tab = 0
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



Private Sub cboSuspens_PDCMVTDEV1_Click()
Call fraSuspens_Display_PDCMVTTAUX

End Sub


Private Sub cboSuspens_PDCMVTDEV2_Click()
Call fraSuspens_Display_PDCMVTTAUX

End Sub


Private Sub cboSuspens_PDCMVTSENS_Click()
fraSuspens_Display_Montant

End Sub

Private Sub cboSuspens_PDCMVTSENS_GotFocus()
Call txt_GotFocus(cboSuspens_PDCMVTSENS)

End Sub


Private Sub cboSuspens_PDCMVTSENS_LostFocus()
Call txt_LostFocus(cboSuspens_PDCMVTSENS)

End Sub


Private Sub cboSuspens_PDCMVTSER_GotFocus()
Call txt_GotFocus(cboSuspens_PDCMVTSER)

End Sub


Private Sub cboSuspens_PDCMVTSER_LostFocus()
Call txt_LostFocus(cboSuspens_PDCMVTSER)

End Sub




Private Sub chkPDCOPEINFO_Click()
If localUnit = "SOBI" Then
    If chkPDCOPEINFO = "1" Then
        lblPDCOPEDEV2 = "Devise du dossier"
    Else
        lblPDCOPEDEV2 = "Devise du débit client"
    End If
    

End If

End Sub

Private Sub chkSelect_HB_Click()
If chkSelect_HB = "1" Then
    blnControl = False
    txtSelect_AMJ_HB.Enabled = True
    chkSelect_Ope.Value = "0"
    chkSelect_Log.Value = "0"
    chkSelect_Suspens.Value = "0"
    chkSelect_ZCHGOPE0.Value = "0"
    chkSelect_ZCHGOPE0_NC.Value = "0"
    If cmdSelect_SQL_K <> "X" Then blnControl = True
Else
    txtSelect_AMJ_HB.Enabled = False
End If
cmdSelect_Reset
End Sub

Private Sub chkSelect_HB_xls_Click()
'cmdSelect_Reset
If chkSelect_HB_xls = "1" Then
    txtSelect_AMJ_HB_xls.Enabled = True
Else
    txtSelect_AMJ_HB_xls.Enabled = False
End If

End Sub

Private Sub chkSelect_Log_Click()
If chkSelect_Log = "1" Then
    blnControl = False
    chkSelect_Ope.Value = "0"
    chkSelect_HB.Value = "0"
    chkSelect_Suspens_Out.Value = "0"
    chkSelect_PDCMVTKCUT.Value = "0"
    chkSelect_Suspens.Value = "0"
    chkSelect_ZCHGOPE0.Value = "0"
    chkSelect_ZCHGOPE0_NC.Value = "0"
    blnControl = True
End If
cmdSelect_Reset

End Sub

Private Sub chkSelect_Ope_Click()
If chkSelect_Ope = "1" Then
    blnControl = False
    chkSelect_HB.Value = "0"
    chkSelect_Suspens_Out.Value = "0"
    chkSelect_PDCMVTKCUT.Value = "0"
    chkSelect_Log.Value = "0"
    chkSelect_Suspens.Value = "0"
    chkSelect_ZCHGOPE0.Value = "0"
    chkSelect_ZCHGOPE0_NC.Value = "0"
    blnControl = True
End If
cmdSelect_Reset

End Sub


Private Sub chkSelect_PDCMVTKCUT_Click()
If chkSelect_PDCMVTKCUT = "1" Then
    blnControl = False
    chkSelect_Ope.Value = "0"
    chkSelect_Log.Value = "0"
    chkSelect_Suspens.Value = "0"
    chkSelect_ZCHGOPE0.Value = "0"
    chkSelect_ZCHGOPE0_NC.Value = "0"
    blnControl = True
End If
cmdSelect_Reset

End Sub

Private Sub chkSelect_Suspens_Click()
If chkSelect_Suspens = "1" Then
    blnControl = False
    chkSelect_HB.Value = "0"
    chkSelect_Suspens_Out.Value = "0"
    chkSelect_PDCMVTKCUT.Value = "0"
    chkSelect_Log.Value = "0"
    chkSelect_Ope.Value = "0"
    chkSelect_ZCHGOPE0.Value = "0"
    chkSelect_ZCHGOPE0_NC.Value = "0"
    blnControl = True
End If
cmdSelect_Reset

End Sub


Private Sub chkSelect_Suspens_Out_Click()
If chkSelect_Suspens_Out = "1" Then
    blnControl = False
    chkSelect_Ope.Value = "0"
    chkSelect_Log.Value = "0"
    chkSelect_Suspens.Value = "0"
    chkSelect_ZCHGOPE0.Value = "0"
    chkSelect_ZCHGOPE0_NC.Value = "0"
    If cmdSelect_SQL_K <> "X" Then blnControl = True
End If
cmdSelect_Reset

End Sub


Private Sub chkSelect_Suspens_Out_xls_Click()
'cmdSelect_Reset

End Sub

Private Sub chkSelect_Terme_Click()
If chkSelect_Terme = "1" Then
    blnControl = False
    chkSelect_Ope.Value = "0"
    chkSelect_Log.Value = "0"
    chkSelect_Suspens.Value = "0"
    chkSelect_ZCHGOPE0.Value = "0"
    chkSelect_ZCHGOPE0_NC.Value = "0"
    If cmdSelect_SQL_K <> "X" Then blnControl = True
End If
cmdSelect_Reset

End Sub

Private Sub chkSelect_ZCHGOPE0_Click()
If chkSelect_ZCHGOPE0 = "1" Then
    blnControl = False
    chkSelect_HB.Value = "0"
    chkSelect_Suspens_Out.Value = "0"
    chkSelect_PDCMVTKCUT.Value = "0"
    chkSelect_Log.Value = "0"
    chkSelect_Suspens.Value = "0"
    chkSelect_Ope.Value = "0"
    chkSelect_ZCHGOPE0_NC.Value = "0"
    blnControl = True
End If
cmdSelect_Reset

End Sub

Private Sub chkSelect_ZCHGOPE0_NC_Click()
If chkSelect_ZCHGOPE0_NC = "1" Then
    blnControl = False
    chkSelect_HB.Value = "0"
    chkSelect_Suspens_Out.Value = "0"
    chkSelect_PDCMVTKCUT.Value = "0"
    chkSelect_Log.Value = "0"
    chkSelect_Suspens.Value = "0"
    chkSelect_Ope.Value = "0"
    chkSelect_ZCHGOPE0.Value = "0"
    blnControl = True
End If
cmdSelect_Reset

End Sub


Private Sub chkSuspens_All_Click()
cmdSelect_Reset

End Sub

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPDC_Param_Click()
Dim V, X8 As String
Dim newY As typeYBIATAB0, oldY As typeYBIATAB0

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents
Call DTPicker_Control(txtPDC_Param, X8)
newY.BIATABID = "PDC_PARAM"
newY.BIATABK1 = "TER"
newY.BIATABK2 = "382100"
newY.BIATABTXT = X8
oldY = newY
oldY.BIATABTXT = mTER382100_Amj

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
 V = sqlYBIATAB0_Update(newY, oldY)
    

If Not IsNull(V) Then GoTo Error_MsgBox
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        mTER382100_Amj = X8
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdPDCOPE_Annulation_Click()
Dim V

If oldYPDCOPE0.PDCOPESTA = "R" Then
    cmdPDCOPE_Report_Annulation
    Exit Sub
End If

Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    newYPDCOPE0 = oldYPDCOPE0
    newYPDCOPE0.PDCOPESTA = "A"
    V = cmdPDCOPE_Transaction
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        fraPDCOPE.Visible = False

        cmdSelect_SQL_1
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdPDCOPE_Update"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdPDCOPE_New_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Saisie d'une opération"): DoEvents
blnPDCOPE_CONF_CALL_Saisie = False
Select Case cmdSelect_SQL_K
    Case 1: fraPDCOPE_Init
    Case 4: fraSuspens_Init
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdPDCOPE_CONF_CALL_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Saisie d'une opération"): DoEvents
blnPDCOPE_CONF_CALL_Saisie = True
fraPDCOPE_Init
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdPDCOPE_Update_Ref_Click()
Dim V

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

xYPDCOPE0 = oldYPDCOPE0
If xYPDCOPE0.PDCOPESTA2 <> " " Then GoTo Exit_sub
'______________________________________________
If xYPDCOPE0.PDCOPESTA3 = "B" Then
    X = InputBox("Préciser le nouveau numéro  : ", "PDC : modification du numéro du numéro d'opération", oldYPDCOPE0.PDCOPEOPEN)
    If Trim(X) = "" Then GoTo Exit_sub
    If Not IsNumeric(X) Then
        Call MsgBox("Le numéro doit être numérique", vbCritical, "PDC : modification du numéro d'opération")
        GoTo Exit_sub
    End If
    xYPDCOPE0.PDCOPEOPEN = Val(X)
    X = fraPDCOPE_Control_PDCOPEOPEN
    
    If X <> "" Then
        Call MsgBox(X, vbCritical, "PDC : modification du numéro d'opération")
        GoTo Exit_sub
    End If
Else
    X = InputBox("Préciser le nouveau numéro  : ", "PDC : modification du numéro de ticket", oldYPDCOPE0.PDCOPEREF)
    If Trim(X) = "" Then GoTo Exit_sub
    If Not IsNumeric(X) Then
        Call MsgBox("Le numéro doit être numérique", vbCritical, "PDC : modification du numéro de ticket")
        GoTo Exit_sub
    End If
    xYPDCOPE0.PDCOPEREF = Val(X)
    X = fraPDCOPE_Control_Ticket
    
    If X <> "" Then
        Call MsgBox(X, vbCritical, "PDC : modification du numéro de ticket")
        GoTo Exit_sub
    End If
End If

'______________________________________________

newYPDCOPE0 = xYPDCOPE0
    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    V = cmdPDCOPE_Transaction
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        'cmdPDCOPE.Visible = BIA_PDC_Aut.Saisir
        fraPDCOPE.Visible = False
        Call DTPicker_Set(txtSelect_AMJ, YBIATAB0_DATE_CPT_JS1)

        cmdSelect_SQL_1
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdPDCOPE_Update_Ref"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If

Exit_sub:

Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdPrint_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
If BIA_PDC_Aut.Avis Then
    If blnDeviseU Then
        If fgSelect.Rows > 1 Then Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
    Else
        cmdPrint_Ok.Caption = " imprimer >  " & Printer.Devicename
        libPrint_Etat.Caption = " Etats au " & dateImp(wAMJMin)
        chkPrint_Exclure_HB = chkSelect_HB
        chkPrint_Suspens_Out = chkSelect_Suspens_Out
        chkPrint_Exclure_PDCMVTKCUT = chkSelect_PDCMVTKCUT
        If wAMJMin >= YBIATAB0_DATE_CPT_J Then
            chkPrint_Suspens.Enabled = True
            chkPrint_PDCMVTKCUT.Enabled = True
        Else
            chkPrint_Suspens.Enabled = False
            chkPrint_PDCMVTKCUT.Enabled = False
        End If
        fraPrint.Visible = True
    End If
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL_1()
Dim V
Dim X As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents
Call DTPicker_Control(txtSelect_AMJ, wAMJMin)
Call DTPicker_Control(txtSelect_AMJ_HB, wAmjMin_HB)


If wAMJMin > YBIATAB0_DATE_CPT_JS1 Then
    Call MsgBox("date > au " & dateImp10(YBIATAB0_DATE_CPT_JS1), vbExclamation, "BIa_PDC")
    Exit Sub
Else
    If wAMJMin = YBIATAB0_DATE_CPT_JS1 Then
        blnPDC_Instant = True
        mPDCPOSDTR = wAMJMin
        wAMJMin = YBIATAB0_DATE_CPT_J
        'chkSelect_Ope.Visible = False
        'chkSelect_Log.Visible = False
    Else
        blnPDC_Instant = False
        chkSelect_Ope.Visible = True
        chkSelect_Log.Visible = True
        mPDCPOSDTR = wAMJMin
    End If
End If

X = Trim(cboSelect_Devise)
If X = "" Then
    blnDeviseU = False
    cmdSelect_SQL_Where = "where PDCPOSDTR = '" & wAMJMin & "'"
Else
    blnDeviseU = True
    cmdSelect_SQL_Where = "where PDCPOSDTR >= '" & wAMJMin & "' and PDCPOSDEV = '" & X & "'"
End If

Call arrYPDCPOS0_SQL(cmdSelect_SQL_Where)

If Not blnDeviseU Then
    Call arrYPDCPOS0_SQL_FixingJ_1(mPDCPOSDTR)

    If blnPDC_Instant Then
        cmdSelect_SQL_1Instant
    Else
        Call cmdSelect_SQL_1_YPDCOPE0_R(wAMJMin)
        If arrYPDCOPE0_Nb > 0 Then
            libSelect_Report.BackColor = vbYellow '&H0000FFFF&
            libSelect_Report = arrYPDCOPE0_Nb & " opérations reportées incluses"
            libSelect_Report.Visible = True
        End If
    End If
    
    If chkSelect_HB = "1" Then
        cmdSelect_SQL_1HB
        fgDetail_Display_YPDCMVT0
    Else
        If chkSelect_Suspens_Out = "1" Then
            cmdSelect_SQL_1HB
            fgDetail_Display_YPDCMVT0
        End If
    End If
    
    If chkSelect_PDCMVTKCUT = "1" Then
            cmdSelect_SQL_1PDCMVTKCUT
            fgDetail_Display_YPDCMVT0
    End If
    If chkSelect_Ope.Visible And chkSelect_Ope = "1" Then cmdSelect_SQL_2
    If chkSelect_ZCHGOPE0 = "1" Then cmdSelect_SQL_2_ZCHGOPE0
    If chkSelect_ZCHGOPE0_NC = "1" Then cmdSelect_SQL_2_ZCHGOPE0


    If chkSelect_Log.Visible And chkSelect_Log = "1" Then
        If blnAuto Then
            cmdSelect_SQL_Where = "where PDCLOGDTR > '" & mPDCLOGDTR_Min & "'"
        Else
            cmdSelect_SQL_Where = "where PDCLOGDTR = '" & wAMJMin & "'"
        End If
        Call arrYPDCLOG0_SQL(cmdSelect_SQL_Where)
        fgDetail_Display_YPDCLOG0
    End If
    If chkSelect_Suspens = "1" Then
        Call arrYPDCMVT0_SQL("where PDCMVTOPEC like 'XX%' and PDCMVTSTA2 = ' '")
        
        fgDetail_Display_YPDCMVT0
    End If
End If

    
'__________________ !!!!!!!!!! ne pas modifier la séquence de traitement : fgTermeEch_Display_1 puis fgSelect_Display_1

fgTermeEch_Display_1

fgSelect_Display_1

'______________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub cmdSelect_SQL_2()
Dim V
Dim X As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2"
Call DTPicker_Control(txtSelect_AMJ, wAMJMin)
cmdSelect_SQL_2X
cboSelect_SQL.ListIndex = 0
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub cmdSelect_SQL_2_ZCHGOPE0()
Dim V
Dim X As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2_ZCHGOPE0"
Call DTPicker_Control(txtSelect_AMJ, wAMJMin)
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents
blnPDCOPEDEVU = False
If chkSelect_ZCHGOPE0_NC = "1" Then
    cmdSelect_SQL_Where = "where CHGOPEDE2 <> '   ' and CHGOPEVAL = '1' and CHGOPEANN = ' '"
Else
    cmdSelect_SQL_Where = "where CHGOPECRE = " & wAMJMin - 19000000 & " and   CHGOPEDE2 <> '   '"
End If
Call arrZCHGOPE0_SQL(cmdSelect_SQL_Where)

fgPDCOPE_ZCHGOPE0


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_4()
Dim V
Dim X As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_4"

If chkSuspens_All = "1" Then
    Call arrYPDCMVT0_SQL("where PDCMVTOPEC like 'XX%'")
Else
    Call arrYPDCMVT0_SQL("where PDCMVTOPEC like 'XX%' and PDCMVTSTA2 = ' '")
End If

fgDetail_Display_YPDCMVT0
cmdPDCOPE_New.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_3()
Dim V, I As Integer
Dim X As String
Dim Nb_T As Integer, Nb_Ok As Integer, Nb_Err As Integer
Dim wMsg As String, xSQL As String
Dim blnOk As Boolean, blnIdem As Boolean
On Error GoTo Error_Handler

Nb_T = 0: Nb_Ok = 0: Nb_Err = 0
currentAction = "BIA_PDC : rapprochement des tickets FOTC / BOTC du " & dateImp10(wAMJMin)
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents
Call DTPicker_Control(txtSelect_AMJ, wAMJMin)
blnPDCOPEDEVU = False
cmdSelect_SQL_Where = "where PDCOPEDTR = '" & wAMJMin & "' and PDCOPESSE = 'TC' and PDCOPESTA = 'T'"
Call arrYPDCOPE0_SQL(cmdSelect_SQL_Where)

For I = 1 To arrYPDCOPE0_Nb
    Nb_T = Nb_T + 1
    If arrYPDCOPE0(I).PDCOPESTA3 = "B" Then
        Nb_Ok = Nb_Ok + 1
    Else
        Nb_Err = Nb_Err + 1
    End If
Next I
If Nb_T = 0 Then
    Call MsgBox("Il n'y a pas d'opérations TC saisies", vbExclamation, currentAction)
Else
    If Nb_Err = 0 Then
        Call MsgBox(Nb_T & " opérations rapprochées", vbInformation, currentAction)
    Else
        Call MsgBox("Il reste " & Nb_Err & " / " & Nb_T & " opérations à rapprocher ", vbCritical, currentAction)
    End If
End If
'_________________________________________________________________________
wMsg = ""
For I = 1 To arrYPDCOPE0_Nb
    xYPDCOPE0 = arrYPDCOPE0(I)
    If xYPDCOPE0.PDCOPEOPEN <> 0 Then
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZCHGOPE0 " _
             & " where CHGOPEOPE ='" & xYPDCOPE0.PDCOPEOPEC & "'" _
             & " and    CHGOPEDOS =" & xYPDCOPE0.PDCOPEOPEN _
             & " and    CHGOPESER ='" & xYPDCOPE0.PDCOPESER & "'" _
             & " and    CHGOPESSE ='" & xYPDCOPE0.PDCOPESSE & "'"
        Set rsSab = cnsab.Execute(xSQL)
        If rsSab.EOF Then
            wMsg = wMsg & "- " & xYPDCOPE0.PDCOPEOPEN & " Inconnu " & vbCrLf
        Else
            blnOk = True: blnIdem = True
            wMsg = wMsg & "- " & xYPDCOPE0.PDCOPEOPEN & "--------------------" & xYPDCOPE0.PDCOPEID & vbCrLf
          
            If xYPDCOPE0.PDCOPEDEV1 = rsSab("CHGOPEDE1") Then
                If newYPDCOPE0.PDCOPEDEV2 <> rsSab("CHGOPEDE2") Then wMsg = wMsg & "+ devises différentes " & xYPDCOPE0.PDCOPEDEV2 & vbCrLf
                If Abs(xYPDCOPE0.PDCOPEMTD1) <> CCur(rsSab("CHGOPEMO1")) Then wMsg = wMsg & "+ montants différents " & xYPDCOPE0.PDCOPEDEV1 & vbCrLf
                If Abs(xYPDCOPE0.PDCOPEMTD2) <> CCur(rsSab("CHGOPEMO2")) Then wMsg = wMsg & "+ montants différents " & xYPDCOPE0.PDCOPEDEV2 & vbCrLf
            Else
                If newYPDCOPE0.PDCOPEDEV2 = rsSab("CHGOPEDE2") Then
                    If newYPDCOPE0.PDCOPEDEV1 <> rsSab("CHGOPEDE1") Then wMsg = wMsg & "+ devises différentes " & xYPDCOPE0.PDCOPEDEV1 & vbCrLf
                    If Abs(newYPDCOPE0.PDCOPEMTD2) <> CCur(rsSab("CHGOPEMO1")) Then wMsg = wMsg & "+ montants différents " & xYPDCOPE0.PDCOPEDEV2 & vbCrLf
                    If Abs(newYPDCOPE0.PDCOPEMTD1) <> CCur(rsSab("CHGOPEMO2")) Then wMsg = wMsg & "+ montants différents " & xYPDCOPE0.PDCOPEDEV1 & vbCrLf
                End If
            End If
            If xYPDCOPE0.PDCOPECLI <> rsSab("CHGOPECON") Then
                wMsg = wMsg & "+ clients différents" & vbCrLf
            End If
            'TODO If newYPDCOPE0.PDCOPEDVA <> CLng(rsSab("CHGOPEDT1")) + 19000000 Then wMsg = wMsg & "+ dates valeur différentes " & vbCrLf
        End If
    End If
Next I

'_____________________________________________
If wMsg <> "" Then
        Call MsgBox(wMsg, vbCritical, "BIA_PDC : rapprochement SAB / ZCHGOPE0")
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_2X()
Dim V
Dim X As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2X"
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents
Call DTPicker_Control(txtSelect_AMJ, wAMJMin)
X = Trim(cboSelect_Devise)

If X = "" Then
    blnPDCOPEDEVU = False
    cmdSelect_SQL_Where = "where PDCOPEDTR = '" & wAMJMin & "'" & mSQL_Unit
Else
    blnPDCOPEDEVU = True
    cmdSelect_SQL_Where = "where PDCOPEDTR >= '" & wAMJMin & "' and PDCOPEDEV1 = '" & X & "'" & mSQL_Unit
End If

Call arrYPDCOPE0_SQL(cmdSelect_SQL_Where)
fgPDCOPE_Display_1

If wAMJMin = YBIATAB0_DATE_CPT_JS1 Then
    cmdPDCOPE_New.Visible = BIA_PDC_Aut.Saisir
    cmdPDCOPE_CONF_CALL.Visible = blnPDCOPE_CONF_CALL_Visible
Else
    cmdPDCOPE_New.Visible = False
    cmdPDCOPE_CONF_CALL.Visible = False
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_5()
Dim V
Dim X As String, xSQL As String, xCur As Currency
Dim blnOk As Boolean, I As Long, K As Long, blnDevOk As Boolean
Dim blnYPDCLOG0_Write As Boolean
Dim Ks As Long
Dim wAMJSWP As String
On Error GoTo Error_Handler

'===================================================
currentAction = "cmdSelect_SQL_5"
Call lstErr_Clear(lstErr, cmdContext, "recherche date dernier traitement :"): DoEvents
blnOk = False
blnYPDCLOG0_Write = True
wAMJMin = YBIATAB0_DATE_CPT_J
mPDCLOGDTR_Min = DateComptablePrecedente(YBIATAB0_DATE_CPT_J)

Call lstErr_AddItem(lstErr, cmdContext, wAMJMin): DoEvents
Do
    Call lstErr_ChangeLastItem(lstErr, cmdContext, wAMJMin): DoEvents
    cmdSelect_SQL_Where = "where PDCPOSDTR = '" & wAMJMin & "'"
    Call arrYPDCPOS0_SQL(cmdSelect_SQL_Where)
    If arrYPDCPOS0_Nb > 0 Then
        blnOk = True
    Else
        wAMJMin = dateElp("Jour", -1, wAMJMin)
        If wAMJMin < "20081231" Then
            If Not blnAuto Then MsgBox "Il manque la position de départ au 31-12-2009", vbCritical, "recherche date dernier traitement"
            Exit Sub
        End If
    End If
Loop Until blnOk

Call lstErr_AddItem(lstErr, cmdContext, "Contrôle cohérence / devise"): DoEvents
For I = 1 To arrDev_Nb
    blnDevOk = False
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "Contrôle cohérence / devise" & arrDev(I)): DoEvents
    For K = 1 To arrYPDCPOS0_Nb
        If arrDev(I) = arrYPDCPOS0(K).PDCPOSDEV Then blnDevOk = True: Exit For
    Next K
    If Not blnDevOk Then
        xYPDCLOG0.PDCLOGPIE = 0
        xYPDCLOG0.PDCLOGECR = 0
        xYPDCLOG0.PDCLOGNAT = "5P?"
        xYPDCLOG0.PDCLOGTXT = "Il manque la position " & arrDev(I) & " au " & arrYPDCPOS0(K).PDCPOSDTR
        Call YPDCLOG0_AddItem

        If Not blnAuto Then MsgBox xYPDCLOG0.PDCLOGTXT, vbCritical, "Contrôle cohérence / devise"
        blnOk = False
    End If
Next I

If wAMJMin = YBIATAB0_DATE_CPT_J Then
    If Not blnAuto Then MsgBox "traitement déjà effectué au " & wAMJMin, vbInformation, "recherche date dernier traitement"
    blnYPDCLOG0_Write = False
    GoTo Controle_Solde
End If

Call lstErr_AddItem(lstErr, cmdContext, "traitement PDC du "): DoEvents
chkSelect_Log.Value = "1"
mPDCLOGDTR_Min = wAMJMin
Do
    WAMJMax = DateComptableSuivanteR(wAMJMin)
    If WAMJMax = "00000000" Then
        If Not blnAuto Then MsgBox "Pas de date comptable suivante après le " & wAMJMin, vbCritical, "recherche date dernier traitement"
        Exit Sub
    End If
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "traitement PDC du " & WAMJMax): DoEvents
    Call cmdSelect_SQL_5P(wAMJMin, WAMJMax)
    Call cmdSelect_SQL_5Suspens(WAMJMax)
    wAMJMin = WAMJMax
Loop Until WAMJMax = YBIATAB0_DATE_CPT_J


'_____________________________________________________________________________
Controle_Solde:
'_____________________________________________________________________________

Call YPDCLOG0_Init(wAMJMin)

'_________________________________________________
' liste des suspens en cours => à exclure du contrôle /SAB
'_________________________________________________
rsYPDCMVT0_Init xYPDCMVT0
For Ks = 1 To arrDev_Nb
    arrSuspens_Dev(Ks) = xYPDCMVT0
    arrSuspens_Dev(Ks).PDCMVTDEV = arrDev(Ks)
Next Ks

Call arrYPDCMVT0_SQL("where PDCMVTOPEC like 'XXC' and PDCMVTSTA2 = ' '")
For I = 1 To arrYPDCMVT0_Nb
    xYPDCMVT0 = arrYPDCMVT0(I)
    For Ks = 1 To arrDev_Nb
        If arrSuspens_Dev(Ks).PDCMVTDEV = xYPDCMVT0.PDCMVTDEV Then
            arrSuspens_Dev(Ks).PDCMVTMTE = arrSuspens_Dev(Ks).PDCMVTMTE + xYPDCMVT0.PDCMVTMTE
            arrSuspens_Dev(Ks).PDCMVTMTD = arrSuspens_Dev(Ks).PDCMVTMTD + xYPDCMVT0.PDCMVTMTD
            Exit For
        End If
    Next Ks

    
    xYPDCLOG0.PDCLOGPIE = xYPDCMVT0.PDCMVTPIE
    xYPDCLOG0.PDCLOGECR = xYPDCMVT0.PDCMVTECR
    xYPDCLOG0.PDCLOGNAT = "5M+"
    xYPDCLOG0.PDCLOGTXT = "Suspens FOTC » " & xYPDCMVT0.PDCMVTDEV & " " & Trim(Format$(xYPDCMVT0.PDCMVTMTD, "### ### ### ##0.00")) & " / EUR " & Trim(Format$(xYPDCMVT0.PDCMVTMTE, "### ### ### ##0.00"))
    Call YPDCLOG0_AddItem
Next I
'_________________________________________________
' cumul des jambes comptant HB des swaps en cours => à exclure du contrôle /SAB
'_________________________________________________

wAMJSWP = dateElp("Ouvré", -5, YBIATAB0_DATE_CPT_J)
rsYPDCMVT0_Init xYPDCMVT0
For Ks = 1 To arrDev_Nb
    arrSWP_Dev(Ks) = xYPDCMVT0
    arrSWP_Dev(Ks).PDCMVTDEV = arrDev(Ks)
Next Ks

'-----------------------------------------
'Call arrYPDCMVT0_SQL("where PDCMVTOPEC like 'SWP' and PDCMVTDTR >= '" & wAMJSWP & "'")
'$JPL 2011-07-18  cumuler tous les swaps
'-----------------------------------------
Call arrYPDCMVT0_SQL("where PDCMVTOPEC like 'SWP'")
For I = 1 To arrYPDCMVT0_Nb
    xYPDCMVT0 = arrYPDCMVT0(I)
    For Ks = 1 To arrDev_Nb
        If arrSWP_Dev(Ks).PDCMVTDEV = xYPDCMVT0.PDCMVTDEV Then
            arrSWP_Dev(Ks).PDCMVTMTE = arrSWP_Dev(Ks).PDCMVTMTE + xYPDCMVT0.PDCMVTMTE
            arrSWP_Dev(Ks).PDCMVTMTD = arrSWP_Dev(Ks).PDCMVTMTD + xYPDCMVT0.PDCMVTMTD
            Exit For
        End If
    Next Ks

Next I

For Ks = 1 To arrDev_Nb
    If arrSWP_Dev(Ks).PDCMVTMTD <> 0 Then
        xYPDCLOG0.PDCLOGPIE = 0
        xYPDCLOG0.PDCLOGECR = 0
        xYPDCLOG0.PDCLOGNAT = "4-"
        xYPDCLOG0.PDCLOGTXT = "SWP comptant HB » " & arrSWP_Dev(Ks).PDCMVTDEV & " " & Trim(Format$(arrSWP_Dev(Ks).PDCMVTMTD, "### ### ### ##0.00")) & " / EUR " & Trim(Format$(arrSWP_Dev(Ks).PDCMVTMTE, "### ### ### ##0.00"))
        Call YPDCLOG0_AddItem
    End If
Next Ks

'_____________________________________________________________________________
'Controle_Solde
'_____________________________________________________________________________


xYPDCLOG0.PDCLOGPIE = 0
xYPDCLOG0.PDCLOGECR = 0
xYPDCLOG0.PDCLOGNAT = "5  "
xYPDCLOG0.PDCLOGTXT = "Controle_Solde " & YBIATAB0_DATE_CPT_J
Call YPDCLOG0_AddItem


Call lstErr_AddItem(lstErr, cmdContext, "Contrôle des soldes EUR /DEV "): DoEvents
For I = 1 To arrDev_Nb

    rsYPDCPOS0_Init newYPDCPOS0
    
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "Contrôle des soldes EUR / " & arrDev(I)): DoEvents

    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '382100" & arrDev(I) & "EUR'"
    Set rsSab = cnsab.Execute(xSQL)

    If rsSab.EOF Then
        xYPDCLOG0.PDCLOGPIE = 0
        xYPDCLOG0.PDCLOGECR = 0
        xYPDCLOG0.PDCLOGNAT = "5P?"
        xYPDCLOG0.PDCLOGTXT = xSQL
        Call YPDCLOG0_AddItem

        If Not blnAuto Then MsgBox xSQL, vbCritical, "YBIACPT0 - B"
        Exit For
    End If
    newYPDCPOS0.PDCPOSPOSD = rsSab("SOLDECEN") / 1000
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '382110" & arrDev(I) & "EUR'"
    Set rsSab = cnsab.Execute(xSQL)

    If Not rsSab.EOF Then newYPDCPOS0.PDCPOSSWPD = -rsSab("SOLDECEN") / 1000
    
    
    
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '382100EUR" & arrDev(I) & "'"
    Set rsSab = cnsab.Execute(xSQL)

    If rsSab.EOF Then
        xYPDCLOG0.PDCLOGPIE = 0
        xYPDCLOG0.PDCLOGECR = 0
        xYPDCLOG0.PDCLOGNAT = "5P?"
        xYPDCLOG0.PDCLOGTXT = xSQL
        Call YPDCLOG0_AddItem
        If Not blnAuto Then MsgBox xSQL, vbCritical, "YBIACPT0 - B-EUR"
        Exit For
    End If
    newYPDCPOS0.PDCPOSPOSE = rsSab("SOLDECEN") / 1000
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '382110EUR" & arrDev(I) & "'"
    Set rsSab = cnsab.Execute(xSQL)

    If Not rsSab.EOF Then newYPDCPOS0.PDCPOSSWPE = -rsSab("SOLDECEN") / 1000
    
    If blnHB Then
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '931000" & arrDev(I) & "EUR'"
        Set rsSab = cnsab.Execute(xSQL)
    
        If rsSab.EOF Then
            xYPDCLOG0.PDCLOGPIE = 0
            xYPDCLOG0.PDCLOGECR = 0
            xYPDCLOG0.PDCLOGNAT = "5P?"
            xYPDCLOG0.PDCLOGTXT = xSQL
            Call YPDCLOG0_AddItem
            If Not blnAuto Then MsgBox xSQL, vbInformation, "YBIACPT0 - HB"
        Else
            newYPDCPOS0.PDCPOSPOSD = newYPDCPOS0.PDCPOSPOSD + rsSab("SOLDECEN") / 1000
        End If
        
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '931000EUR" & arrDev(I) & "'"
        Set rsSab = cnsab.Execute(xSQL)
        
        If rsSab.EOF Then
            xYPDCLOG0.PDCLOGPIE = 0
            xYPDCLOG0.PDCLOGECR = 0
            xYPDCLOG0.PDCLOGNAT = "5P?"
            xYPDCLOG0.PDCLOGTXT = xSQL
            Call YPDCLOG0_AddItem
            If Not blnAuto Then MsgBox xSQL, vbInformation, "YBIACPT0 - HB-EUR"
        Else
            newYPDCPOS0.PDCPOSPOSE = newYPDCPOS0.PDCPOSPOSE + rsSab("SOLDECEN") / 1000
        End If
    End If

    xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCPOS0 where PDCPOSDTR = '" & YBIATAB0_DATE_CPT_J & "' and PDCPOSDEV = '" & arrDev(I) & "'"
    Set rsSab = cnsab.Execute(xSQL)

    If rsSab.EOF Then
        xYPDCLOG0.PDCLOGPIE = 0
        xYPDCLOG0.PDCLOGECR = 0
        xYPDCLOG0.PDCLOGNAT = "5P?"
        xYPDCLOG0.PDCLOGTXT = xSQL
        Call YPDCLOG0_AddItem
        If Not blnAuto Then MsgBox xSQL, vbCritical, "YPDCPOS0"
        Exit For
    End If
    xCur = rsSab("PDCPOSPOSE") + rsSab("PDCPOSRPC") - arrSuspens_Dev(I).PDCMVTMTE + arrSWP_Dev(I).PDCMVTMTE
    If newYPDCPOS0.PDCPOSPOSE <> xCur Then
            xYPDCLOG0.PDCLOGPIE = 0
            xYPDCLOG0.PDCLOGECR = 0
            xYPDCLOG0.PDCLOGNAT = "5P?"
            xYPDCLOG0.PDCLOGTXT = arrDev(I) & " comptant CV calculée EUR " & Trim(Format$(xCur, "### ### ### ##0.00")) & " SAB " & Trim(Format$(newYPDCPOS0.PDCPOSPOSE, "### ### ### ##0.00"))
            Call YPDCLOG0_AddItem
            If Not blnAuto Then MsgBox xYPDCLOG0.PDCLOGTXT, vbCritical, "Contrôle des soldes EUR / " & arrDev(I)
    End If

    xCur = rsSab("PDCPOSPOSD") - arrSuspens_Dev(I).PDCMVTMTD + arrSWP_Dev(I).PDCMVTMTD
    If newYPDCPOS0.PDCPOSPOSD <> xCur Then
            xYPDCLOG0.PDCLOGPIE = 0
            xYPDCLOG0.PDCLOGECR = 0
            xYPDCLOG0.PDCLOGNAT = "5P?"
            xYPDCLOG0.PDCLOGTXT = arrDev(I) & " comptant Pos calculée " & arrDev(I) & " " & Trim(Format$(xCur, "### ### ### ##0.00")) & " SAB " & Trim(Format$(newYPDCPOS0.PDCPOSPOSD, "### ### ### ##0.00"))
            Call YPDCLOG0_AddItem
            If Not blnAuto Then MsgBox xYPDCLOG0.PDCLOGTXT, vbCritical, "Contrôle soldes EUR / " & arrDev(I)
    End If
Next I

'_________________________________________________
' Matching PDC / SAB + report opérations
'_________________________________________________

Call cmdSelect_SQL_5OPE(YBIATAB0_DATE_CPT_JP0, YBIATAB0_DATE_CPT_JS1)
Call cmdSelect_SQL_5OPE_Report
'_________________________________________________
' Liste des opérations SAB saisies et non comptabilisées
'_________________________________________________

'"where CHGOPEDE2 <> '   ' and CHGOPEVAL = '1' and CHGOPEANN = ' '"
If blnAuto Then
    xSQL = "where CHGOPEDE2 <> '   ' and CHGOPEVAL = '1' and CHGOPEANN = ' '"
    Call arrZCHGOPE0_SQL(xSQL)
    If arrZCHGOPE0_Nb > 0 Then
        xYPDCLOG0.PDCLOGPIE = 0
        xYPDCLOG0.PDCLOGECR = 0
        xYPDCLOG0.PDCLOGNAT = "5!!"
        xYPDCLOG0.PDCLOGTXT = "Opérations SAB saisies et non comptabilisées : " & arrZCHGOPE0_Nb
        Call YPDCLOG0_AddItem
    End If
End If
'______________________________________________________________________________________
If blnYPDCLOG0_Write Then YPDCLOG0_Write

Call DTPicker_Set(txtSelect_AMJ, YBIATAB0_DATE_CPT_J)
blnControl = True
cmdSelect_SQL_1
cboSelect_SQL.ListIndex = 0

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " " & currentAction


End Sub

Private Sub cmdSelect_SQL_7()
Dim V
Dim X As String, wNb As Long
On Error GoTo Error_Handler


'$2009-07-17 reprise suspens TERME
'MsgBox "cmdSelect_SQL_7XXT", vbInformation, "reprise POSPDCTER"
'Call cmdSelect_SQL_7XXT("USD")
'Call cmdSelect_SQL_7XXT("GBP")

'Exit Sub


currentAction = "cmdSelect_SQL_7"
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents
Call DTPicker_Control(txtSelect_AMJ, wAMJMin)

Call lstErr_AddItem(lstErr, cmdContext, currentAction): DoEvents
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Call YPDCLOG0_Init(wAMJMin)
xYPDCLOG0.PDCLOGNAT = "7  "
xYPDCLOG0.PDCLOGTXT = "Annulation POS + MVT »= " & wAMJMin
Call YPDCLOG0_AddItem

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________


V = sqlYPDCPOS0_DeleteW("where PDCPOSDTR >= '" & wAMJMin & "'", wNb)
If Not IsNull(V) Then GoTo Error_MsgBox
xYPDCLOG0.PDCLOGNAT = "7P "
xYPDCLOG0.PDCLOGTXT = wNb & " enregistrements supprimés YPDCPOS0"
Call YPDCLOG0_AddItem

V = sqlYPDCMVT0_DeleteW("where PDCMVTDTR >= '" & wAMJMin & "'", wNb)
If Not IsNull(V) Then GoTo Error_MsgBox
xYPDCLOG0.PDCLOGNAT = "7M "
xYPDCLOG0.PDCLOGTXT = wNb & " enregistrements supprimés YPDCMVT0"
Call YPDCLOG0_AddItem


Set rsSab = Nothing

GoTo Exit_sub



'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
        X = "ERREUR - Rollback"
    Else
        V = cnSAB_Transaction("Commit")
        X = "Terminée"
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
xYPDCLOG0.PDCLOGNAT = "7* "
xYPDCLOG0.PDCLOGTXT = "Annulation POS + MVT »= " & wAMJMin & " » " & X
Call YPDCLOG0_AddItem

YPDCLOG0_Write

cboSelect_SQL.ListIndex = 0

End Sub

Private Sub cmdSelect_SQL_7XXT(lDev As String)
Dim V, xSQL As String
Dim X As String, wNb As Long, K As Long

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_7XXT"
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents
Call DTPicker_Control(txtSelect_AMJ, wAMJMin)

Call lstErr_AddItem(lstErr, cmdContext, currentAction): DoEvents

ReDim arrYPDCPOS0(300), selYPDCPOS0(300)

arrYPDCPOS0_Nb = 0
xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCPOS0 Where PDCPOSDEV = '" & lDev & "' order by PDCPOSDTR"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYPDCPOS0_GetBuffer(rsSab, xYPDCPOS0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmBIA_PDC.fgselect_Display"
        '' Exit Sub
     Else
         arrYPDCPOS0_Nb = arrYPDCPOS0_Nb + 1
         selYPDCPOS0(arrYPDCPOS0_Nb) = xYPDCPOS0
         xYPDCPOS0.PDCPOSTERD = 0: xYPDCPOS0.PDCPOSTERE = 0
         xYPDCPOS0.PDCPOSSWPD = 0: xYPDCPOS0.PDCPOSSWPE = 0
         
         arrYPDCPOS0(arrYPDCPOS0_Nb) = xYPDCPOS0
    End If
    rsSab.MoveNext

Loop

'___________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCMVT0 Where PDCMVTOPEC = 'XXT' and PDCMVTDEV = '" & lDev & "' order by PDCMVTDTR"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYPDCMVT0_GetBuffer(rsSab, xYPDCMVT0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmBIA_PDC.fgselect_Display"
        '' Exit Sub
     Else
        For K = 1 To arrYPDCPOS0_Nb
            If arrYPDCPOS0(K).PDCPOSDTR >= xYPDCMVT0.PDCMVTDTR Then
                arrYPDCPOS0(K).PDCPOSTERD = arrYPDCPOS0(K).PDCPOSTERD + xYPDCMVT0.PDCMVTMTD
                arrYPDCPOS0(K).PDCPOSTERE = arrYPDCPOS0(K).PDCPOSTERE + xYPDCMVT0.PDCMVTMTE
            End If
        Next K
    End If
    rsSab.MoveNext

Loop

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
For K = 1 To arrYPDCPOS0_Nb

    V = sqlYPDCPOS0_Update(arrYPDCPOS0(K), selYPDCPOS0(K))

Next K

Set rsSab = Nothing

GoTo Exit_sub



'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
        X = "ERREUR - Rollback"
    Else
        V = cnSAB_Transaction("Commit")
        X = "Terminée"
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
cboSelect_SQL.ListIndex = 0

End Sub

Private Sub YPDCLOG0_Write()
Dim V
Dim X As String, K As Long
On Error GoTo Error_Handler

currentAction = "YPDCLOG0_Write"
Call lstErr_AddItem(lstErr, cmdContext, currentAction): DoEvents
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________


For K = 1 To arrYPDCLOG0_Nb
    V = sqlYPDCLOG0_Insert(arrYPDCLOG0(K))
    If Not IsNull(V) Then GoTo Error_MsgBox
Next K


Set rsSab = Nothing

GoTo Exit_sub



'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Sub

Private Sub cmdSelect_SQL_9()
Dim V, I As Long, K As Long
Dim X As String, xSQL As String, xMemo As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_9"
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents

ReDim arrYPDCPOS0(101):
arrYPDCPOS0_Max = 100: arrYPDCPOS0_Nb = 0
ReDim arrYPDCMVT0(10001)
Call lstErr_AddItem(lstErr, cmdContext, "YBIACPT0 - A"): DoEvents


wAMJMin = "20131028"
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
'For I = 1 To arrDev_Nb
For I = 10 To 10
    rsYPDCPOS0_Init newYPDCPOS0
    'newYPDCPOS0.PDCPOSDTR = "20081231"
    newYPDCPOS0.PDCPOSDTR = wAMJMin
    newYPDCPOS0.PDCPOSDEV = arrDev(I)
    Call sqlYBIATAB0_Read("PDC", newYPDCPOS0.PDCPOSDEV, newYPDCPOS0.PDCPOSDTR, xMemo)
    newYPDCPOS0.PDCPOSFIXT = CDbl(Mid$(xMemo, 9, 15) / 1000000000)
    newYPDCPOS0.PDCPOSFIXD = Mid$(xMemo, 1, 8)
    '________________________________________________________________________________
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '382100" & newYPDCPOS0.PDCPOSDEV & "EUR'"
    Set rsSab = cnsab.Execute(xSQL)

    If rsSab.EOF Then
        MsgBox xSQL, vbCritical, "YBIACPT0 - B"
        Exit For
    End If
    arrYPDCPOS0_Nb = I
    'newYPDCPOS0.PDCPOSDEV = rsSab("COMPTEDEV")
    newYPDCPOS0.PDCPOSPOSD = rsSab("SOLDECEN") / 1000
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '382100EUR" & newYPDCPOS0.PDCPOSDEV & "'"
    Set rsSab = cnsab.Execute(xSQL)

    If rsSab.EOF Then
        MsgBox xSQL, vbCritical, "YBIACPT0 - B-EUR"
        Exit For
    End If
    newYPDCPOS0.PDCPOSPOSE = rsSab("SOLDECEN") / 1000
    
    If blnHB Then
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '931000" & newYPDCPOS0.PDCPOSDEV & "EUR'"
        Set rsSab = cnsab.Execute(xSQL)
    
        If rsSab.EOF Then
            MsgBox xSQL, vbInformation, "YBIACPT0 - HB"
        Else
            newYPDCPOS0.PDCPOSPOSD = newYPDCPOS0.PDCPOSPOSD + rsSab("SOLDECEN") / 1000
        End If
        
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '931000EUR" & newYPDCPOS0.PDCPOSDEV & "'"
        Set rsSab = cnsab.Execute(xSQL)
        
        If rsSab.EOF Then
            MsgBox xSQL, vbInformation, "YBIACPT0 - HB-EUR"
        Else
            newYPDCPOS0.PDCPOSPOSE = newYPDCPOS0.PDCPOSPOSE + rsSab("SOLDECEN") / 1000
        End If
    End If
    
    
'_______________________________________________________________________________________
    arrYPDCMVT0_Nb = 0
    
    Call cmdSelect_SQL_9M("382100")
    If blnHB Then Call cmdSelect_SQL_9M("931000")
    If newYPDCPOS0.PDCPOSPOSE = 0 Then
        newYPDCPOS0.PDCPOSPRIX = 0
    Else
        newYPDCPOS0.PDCPOSPRIX = Round(Abs(newYPDCPOS0.PDCPOSPOSD / newYPDCPOS0.PDCPOSPOSE), 6)
    End If
    '___________________________________________________ terme
    Call cmdSelect_SQL_9_Terme("382110")
    Call cmdSelect_SQL_9_Terme("933000")
'_______________________________________________________________________________________
    V = sqlYPDCPOS0_Insert(newYPDCPOS0)
    If Not IsNull(V) Then GoTo Error_MsgBox
    'For K = 1 To arrYPDCMVT0_Nb
    '    V = sqlYPDCMVT0_Insert(arrYPDCMVT0(K))
    '    If Not IsNull(V) Then GoTo Error_Msgbox
    'Next K
Next I
Call lstErr_AddItem(lstErr, cmdContext, "YBIACPT0 - A : " & arrYPDCPOS0_Nb): DoEvents
Set rsSab = Nothing

GoTo Exit_sub



'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'Call DTPicker_Set(txtSelect_AMJ, "20081231")
Call DTPicker_Set(txtSelect_AMJ, wAMJMin)

cmdSelect_SQL_1
cboSelect_SQL.ListIndex = 0

End Sub
Private Sub cmdSelect_SQL_5P(lPDCPOSDTR_1 As String, lPDCPOSDTR As String)
Dim V, I As Long, K As Long
Dim X As String, xSQL As String, xMemo As String
Dim wAmj As Long, mMOUVEMDTR As Long
Dim wSUSPENS_Dev As Currency, wSUSPENS_EUR As Currency
Dim meYBIAMON0 As typeYBIAMON0
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_5P"
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents

ReDim arrYPDCMVT0(1001)
Call lstErr_AddItem(lstErr, cmdContext, "YBIACPT0 - A"): DoEvents
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Call YPDCLOG0_Init(lPDCPOSDTR)
xYPDCLOG0.PDCLOGNAT = "5  "
xYPDCLOG0.PDCLOGTXT = "Calcul PDC au " & lPDCPOSDTR & " (veille » " & lPDCPOSDTR_1 & " )"
Call YPDCLOG0_AddItem

'_________________________________________________________________________________
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

mPDCPOSDTR_PNL = lPDCPOSDTR_1

'________________________________________________________________________________
For I = 1 To arrDev_Nb
    Call lstErr_AddItem(lstErr, cmdContext, lPDCPOSDTR & " - " & arrDev(I)): DoEvents
    
    wSUSPENS_Dev = 0: wSUSPENS_EUR = 0
    '---------------------------------
    xSQL = "select PDCMVTMTE,PDCMVTMTD from " & paramIBM_Library_SABSPE_XXX & ".YPDCMVT0 " _
         & "where PDCMVTOPEC = 'XXC' and PDCMVTSTA2 = ' ' and PDCMVTDEV = '" & arrDev(I) & "'"
   
    Set rsSab = cnsab.Execute(xSQL)
    
    Do While Not rsSab.EOF
        wSUSPENS_Dev = wSUSPENS_Dev + rsSab("PDCMVTMTD")
        wSUSPENS_EUR = wSUSPENS_EUR + rsSab("PDCMVTMTE")
        rsSab.MoveNext
    Loop

    xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCPOS0 " _
         & " where PDCPOSDTR ='" & lPDCPOSDTR_1 & "'" _
         & " and    PDCPOSDEV ='" & arrDev(I) & "'"
    Set rsSab = cnsab.Execute(xSQL)

    If rsSab.EOF Then
        xYPDCLOG0.PDCLOGPIE = 0
        xYPDCLOG0.PDCLOGECR = 0
        xYPDCLOG0.PDCLOGNAT = "5P?"
        xYPDCLOG0.PDCLOGTXT = "Enregistrement non trouvé YPDCPOS0 » " & lPDCPOSDTR & " " & arrDev(I)
        Call YPDCLOG0_AddItem

        'MsgBox xSql, vbCritical, currentAction
        Exit For
    End If
    V = rsYPDCPOS0_GetBuffer(rsSab, oldYPDCPOS0)
    
    newYPDCPOS0 = oldYPDCPOS0
    
    newYPDCPOS0.PDCPOSDTR = lPDCPOSDTR
    newYPDCPOS0.PDCPOSPOSE = newYPDCPOS0.PDCPOSPOSE + newYPDCPOS0.PDCPOSRPC
    newYPDCPOS0.PDCPOSRPC = 0
    newYPDCPOS0.PDCPOSPNL = -Round(((newYPDCPOS0.PDCPOSPOSD - wSUSPENS_Dev) + (newYPDCPOS0.PDCPOSPOSE - wSUSPENS_EUR) * newYPDCPOS0.PDCPOSFIXT) / newYPDCPOS0.PDCPOSFIXT, 2)
    '________________________________________________________________________________
    newYPDCPOS0.PDCPOSFIXD = lPDCPOSDTR
    'Call sqlYBIATAB0_Read("FIXING", newYPDCPOS0.PDCPOSDEV, "MP1", xMemo)
    Call sqlYBIATAB0_Read("PDC", newYPDCPOS0.PDCPOSDEV, lPDCPOSDTR, xMemo)
    If IsNumeric(Mid$(xMemo, 9, 15)) Then
        newYPDCPOS0.PDCPOSFIXT = CDbl(Mid$(xMemo, 9, 15) / 1000000000)
        newYPDCPOS0.PDCPOSFIXD = Mid$(xMemo, 1, 8)
    End If
'_______________________________________________________________________________________
    
    arrYPDCMVT0_Nb = 0
    mMOUVEMDTR = Val(lPDCPOSDTR) - 19000000

    Call cmdSelect_SQL_5M("382100", mMOUVEMDTR)
    If blnHB Then Call cmdSelect_SQL_5M("931000", mMOUVEMDTR)

    Call cmdSelect_SQL_5M_Terme("382110", mMOUVEMDTR)
    If blnHB Then Call cmdSelect_SQL_5M_Terme("933000", mMOUVEMDTR)
    
    If newYPDCPOS0.PDCPOSPOSE = 0 Then
        newYPDCPOS0.PDCPOSPRIX = 0
    Else
        newYPDCPOS0.PDCPOSPRIX = Round(Abs((newYPDCPOS0.PDCPOSPOSD - wSUSPENS_Dev) / (newYPDCPOS0.PDCPOSPOSE - wSUSPENS_EUR)), 6)
        If newYPDCPOS0.PDCPOSPRIX > 999 Then newYPDCPOS0.PDCPOSPRIX = 0
    End If
    For K = 1 To arrYPDCMVT0_Nb
        If arrYPDCMVT0(K).PDCMVTOPEC = "CPT" Then
            newYPDCMVT0 = arrYPDCMVT0(K)
            Call cmdSelect_SQL_5CPT
            arrYPDCMVT0(K) = newYPDCMVT0
        End If
    Next K
    
    V = sqlYPDCPOS0_Insert(newYPDCPOS0)
    If Not IsNull(V) Then GoTo Error_MsgBox
    For K = 1 To arrYPDCMVT0_Nb
        V = sqlYPDCMVT0_Insert(arrYPDCMVT0(K))
        If Not IsNull(V) Then GoTo Error_MsgBox
    Next K
Next I
    xYPDCLOG0.PDCLOGPIE = 0
    xYPDCLOG0.PDCLOGECR = 0
    xYPDCLOG0.PDCLOGNAT = "5M "
    xYPDCLOG0.PDCLOGTXT = arrYPDCMVT0_Nb & " enregistrements ajoutés YPDCMVT0"
    Call YPDCLOG0_AddItem

Call lstErr_AddItem(lstErr, cmdContext, "YBIACPT0 - A » "): DoEvents
Set rsSab = Nothing

'___________________________________________________________________________________________
GoTo Exit_sub



'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
        X = "ERREUR - Rollback"
    Else
        V = cnSAB_Transaction("Commit")
        X = "Terminée"
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
xYPDCLOG0.PDCLOGPIE = 0
xYPDCLOG0.PDCLOGECR = 0
xYPDCLOG0.PDCLOGNAT = "5* "
xYPDCLOG0.PDCLOGTXT = "Calcul PDC au " & lPDCPOSDTR & " » " & X
Call YPDCLOG0_AddItem

Call cmdSelect_SQL_5Control(lPDCPOSDTR)

YPDCLOG0_Write

End Sub

Private Sub cmdSelect_SQL_5OPE(lPDCPOSDTR_1 As String, lPDCPOSDTR As String)
Dim V, I As Long, Nb As Long, K As Long
Dim X As String, xSQL As String, xMemo As String
Dim mPDCOPEID As Long
Dim xSql_PDCOPEREF As String
Dim nbTC_SAB As Long, nbTC_SAB_Ok As Long, nbTC_PDC As Long, nbTC_PDC_Ok As Long
Dim nbXX_SAB As Long, nbXX_SAB_Ok As Long, nbXX_PDC As Long, nbXX_PDC_Ok As Long
Dim blnAReporter As Boolean

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_5OPE"
nbTC_SAB = 0: nbTC_SAB_Ok = 0: nbTC_PDC = 0: nbTC_PDC_Ok = 0
nbXX_SAB = 0: nbXX_SAB_Ok = 0: nbXX_PDC = 0: nbXX_PDC_Ok = 0
mPDCOPEID = 0

Call arrYPDCOPE0_SQL(" where PDCOPEDTR ='" & lPDCPOSDTR & "'")
For I = 1 To arrYPDCOPE0_Nb
    If arrYPDCOPE0(I).PDCOPEID > mPDCOPEID Then mPDCOPEID = arrYPDCOPE0(I).PDCOPEID
Next I

Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents
Call arrYPDCOPE0_SQL(" where PDCOPEDTR ='" & lPDCPOSDTR_1 & "' and PDCOPESTA in ('V','T','0')")

ReDim selYPDCOPE0(arrYPDCOPE0_Nb + 10)
For I = 1 To arrYPDCOPE0_Nb
    selYPDCOPE0(I) = arrYPDCOPE0(I)
Next I

xSQL = " where CHGOPECRE =" & lPDCPOSDTR_1 - 19000000 & " and   CHGOPEDE2 <> '   '" _
     & " and CHGOPEVAL = 'O' and CHGOPEANN = ' ' and CHGOPESSE <> 'GU'"
Call arrZCHGOPE0_SQL(xSQL)

ReDim arrCHGOPEVAL(arrZCHGOPE0_Nb + 1)

For I = 1 To arrZCHGOPE0_Nb
    xZCHGOPE0 = arrZCHGOPE0(I)
    arrCHGOPEVAL(I) = xZCHGOPE0.CHGOPEVAL
    
    If xZCHGOPE0.CHGOPESSE = "TC" Then
        nbTC_SAB = nbTC_SAB + 1
        V = cmdSelect_SQL_5OPE_TC
        If IsNull(V) Then nbTC_SAB_Ok = nbTC_SAB_Ok + 1
    Else
        nbXX_SAB = nbXX_SAB + 1
        V = cmdSelect_SQL_5OPE_XX
        If IsNull(V) Then nbXX_SAB_Ok = nbXX_SAB_Ok + 1
   End If

Next I

For I = 1 To arrYPDCOPE0_Nb
    If arrYPDCOPE0(I).PDCOPESTA2 = " " Then
    
        If arrYPDCOPE0(I).PDCOPESSE = "TC" Then
        
            xSQL = " where CHGOPEOPE = '" & arrYPDCOPE0(I).PDCOPEOPEC & "'" _
                 & " and   CHGOPEDOS = " & arrYPDCOPE0(I).PDCOPEOPEN _
                 & " and   CHGOPESSE = 'TC'"
            Call arrZCHGOPE0_SQL(xSQL)
    
            For K = 1 To arrZCHGOPE0_Nb
                xZCHGOPE0 = arrZCHGOPE0(K)
                V = cmdSelect_SQL_5OPE_TC
                If IsNull(V) Then nbTC_SAB = nbTC_SAB + 1: nbTC_SAB_Ok = nbTC_SAB_Ok + 1
            Next K
        Else
            If arrYPDCOPE0(I).PDCOPESSE = "TR" Then
                xSQL = " where CHGOPECRE >=" & arrYPDCOPE0(I).PDCOPEREF - 19000000 & " and   CHGOPEDE2 <> '   '" _
                     & " and   CHGOPECON = " & arrYPDCOPE0(I).PDCOPECLI _
                     & " and CHGOPEVAL = 'O' and CHGOPEANN = ' '  and CHGOPESSE <> 'GU' and   CHGOPESSE <> 'TC'"
                Call arrZCHGOPE0_SQL(xSQL)
        
                For K = 1 To arrZCHGOPE0_Nb
                    xZCHGOPE0 = arrZCHGOPE0(K)
                    V = cmdSelect_SQL_5OPE_XX
                    If IsNull(V) Then nbXX_SAB = nbXX_SAB + 1: nbXX_SAB_Ok = nbXX_SAB_Ok + 1
                Next K
            End If
        End If
   End If
    
Next I

'__________________________________________________________________________
'$JPL 20110719 - rapprochementCDE / CDI - report OPE validée mais non comptabilisée

Call arrYPDCMVT0_SQL("where PDCMVTDTR = '" & lPDCPOSDTR_1 & "'")

For I = 1 To arrYPDCOPE0_Nb
    xYPDCOPE0 = arrYPDCOPE0(I)
    For K = 1 To arrYPDCMVT0_Nb
    
        If xYPDCOPE0.PDCOPEOPEC = arrYPDCMVT0(K).PDCMVTOPEC _
        And xYPDCOPE0.PDCOPEOPEN = arrYPDCMVT0(K).PDCMVTOPEN Then
        
            If xYPDCOPE0.PDCOPEDEV1 = arrYPDCMVT0(K).PDCMVTDEV Then
                xYPDCOPE0.PDCOPEMTD1 = xYPDCOPE0.PDCOPEMTD1 - arrYPDCMVT0(K).PDCMVTMTD
                If xYPDCOPE0.PDCOPEDEV2 = "EUR" Then
                    xYPDCOPE0.PDCOPEMTD2 = xYPDCOPE0.PDCOPEMTD2 - arrYPDCMVT0(K).PDCMVTMTE
                End If
            Else
                If xYPDCOPE0.PDCOPEDEV2 = arrYPDCMVT0(K).PDCMVTDEV Then
                    xYPDCOPE0.PDCOPEMTD2 = xYPDCOPE0.PDCOPEMTD2 - arrYPDCMVT0(K).PDCMVTMTD
                    If xYPDCOPE0.PDCOPEDEV1 = "EUR" Then
                        xYPDCOPE0.PDCOPEMTD1 = xYPDCOPE0.PDCOPEMTD1 - arrYPDCMVT0(K).PDCMVTMTE
                    End If
                End If
            End If
        End If
    Next K

'$JPL 2011-11-15    If Abs(xYPDCOPE0.PDCOPEMTD1) < 0.01 And Abs(xYPDCOPE0.PDCOPEMTD2) < 0.01 Then
    If Abs(xYPDCOPE0.PDCOPEMTD1) < 1 And Abs(xYPDCOPE0.PDCOPEMTD2) < 1 Then
        
        If arrYPDCOPE0(I).PDCOPESTA2 = " " Then arrYPDCOPE0(I).PDCOPESTA2 = "="
    Else
        If arrYPDCOPE0(I).PDCOPESTA2 = "=" Then
            arrYPDCOPE0(I).PDCOPESTA2 = " "
            
            xYPDCLOG0.PDCLOGPIE = 0
            xYPDCLOG0.PDCLOGECR = 0
            xYPDCLOG0.PDCLOGNAT = "5=#"
            xYPDCLOG0.PDCLOGTXT = "Ecart saisie PDC / Compta : " & xYPDCOPE0.PDCOPEOPEC & " " & xYPDCOPE0.PDCOPEOPEN _
                       & " " & xYPDCOPE0.PDCOPEDEV1 & " " & Format$(xYPDCOPE0.PDCOPEMTD1, "### ### ### ##0.00") _
                       & " " & xYPDCOPE0.PDCOPEDEV2 & " " & Format$(xYPDCOPE0.PDCOPEMTD2, "### ### ### ##0.00")
            Call YPDCLOG0_AddItem
        End If
    End If
Next I

'___________________________________________________________________________
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'_________________________________________________________________________________
V = cnSAB_Transaction("BeginTrans")
'If Not IsNull(V) Then GoTo Error_MsgBox

'________________________________________________________________________________

For I = 1 To arrYPDCOPE0_Nb
    blnAReporter = False
    newYPDCOPE0 = arrYPDCOPE0(I)
    If newYPDCOPE0.PDCOPESSE = "TC" Then
        nbTC_PDC = nbTC_PDC + 1
    Else
        nbXX_PDC = nbXX_PDC + 1
    End If
    
    If newYPDCOPE0.PDCOPESTA2 = "=" Then
        If newYPDCOPE0.PDCOPESSE = "TC" Then
            nbTC_PDC_Ok = nbTC_PDC_Ok + 1
        Else
            nbXX_PDC_Ok = nbXX_PDC_Ok + 1
        End If
        V = sqlYPDCOPE0_Update(newYPDCOPE0, selYPDCOPE0(I))
    Else
        If newYPDCOPE0.PDCOPESSE = "TC" Then
            xYPDCLOG0.PDCLOGPIE = 0
            xYPDCLOG0.PDCLOGECR = 0
            xYPDCLOG0.PDCLOGNAT = "5=#"
            If newYPDCOPE0.PDCOPESTA2 = "#" Then
                xYPDCLOG0.PDCLOGTXT = "TC : ticket " & newYPDCOPE0.PDCOPEREF & " différent dans SAB (opération : " & newYPDCOPE0.PDCOPEOPEN & ")"
            Else
                blnAReporter = True
                xYPDCLOG0.PDCLOGTXT = "TC : ticket " & newYPDCOPE0.PDCOPEREF & " inconnu dans SAB (opération : " & newYPDCOPE0.PDCOPEOPEN & ")"
            End If
            Call YPDCLOG0_AddItem
        Else
            xYPDCLOG0.PDCLOGPIE = 0
            xYPDCLOG0.PDCLOGECR = 0
            xYPDCLOG0.PDCLOGNAT = "5=!"
            xYPDCLOG0.PDCLOGTXT = newYPDCOPE0.PDCOPESSE & " report de la transaction " _
                    & Trim(Format$(newYPDCOPE0.PDCOPEMTD1, "### ### ### ##0.00")) & " " & newYPDCOPE0.PDCOPEDEV1 & " / " _
                    & Trim(Format$(newYPDCOPE0.PDCOPEMTD2, "### ### ### ##0.00")) & " " & newYPDCOPE0.PDCOPEDEV2
            Call YPDCLOG0_AddItem
            blnAReporter = True
            If newYPDCOPE0.PDCOPEREF = 0 Then
                xSql_PDCOPEREF = " , PDCOPEREF = " & newYPDCOPE0.PDCOPEDTR
            Else
                xSql_PDCOPEREF = ""
            End If
        End If
'____________________________________________________________________________________________
        If blnAReporter Then
            If newYPDCOPE0.PDCOPESTA = "V" Or newYPDCOPE0.PDCOPESTA = "T" Then
                newYPDCOPE0.PDCOPESTA = "R"
                V = sqlYPDCOPE0_Update(newYPDCOPE0, selYPDCOPE0(I))
            End If
            
            mPDCOPEID = mPDCOPEID + 1
            xYPDCOPE0 = selYPDCOPE0(I)
            xYPDCOPE0.PDCOPEID = mPDCOPEID
            xYPDCOPE0.PDCOPEDTR = lPDCPOSDTR
            xYPDCOPE0.PDCOPESTA2 = " "
            V = sqlYPDCOPE0_Insert(xYPDCOPE0)
            
            newYPDCOPE0 = xYPDCOPE0
            Call cmdSendMail_Report
            
           ' xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YPDCOPE0" _
           '     & " set PDCOPEDTR = '" & lPDCPOSDTR & "'" _
           '     & " , PDCOPEID = " & mPDCOPEID _
           '     & xSql_PDCOPEREF _
           '     & " , PDCOPEUPDS = " & newYPDCOPE0.PDCOPEUPDS + 1 _
           '     & " where PDCOPEDTR = '" & newYPDCOPE0.PDCOPEDTR & "' And PDCOPEID = " & newYPDCOPE0.PDCOPEID _
           '        & " and PDCOPEUPDS = " & newYPDCOPE0.PDCOPEUPDS
           ' Set rsSab_Update = cnSab_Update.Execute(xSql, Nb)
    
            ' Tester si la mise à jour a été effectuée
            '===================================================================================
            
            'If Nb = 0 Then
            '    V = "Erreur màj : " & xSql
            '    Error_Route V
            'Else
            '    Call cmdSendMail_Report
            'End If
        End If
    End If

Next I

Call lstErr_AddItem(lstErr, cmdContext, "cmdSelect_SQL_5OPE"): DoEvents
Set rsSab = Nothing

GoTo Exit_sub



'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
        X = "ERREUR - Rollback"
    Else
        V = cnSAB_Transaction("Commit")
        X = "Terminée"
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Exit_Log:
xYPDCLOG0.PDCLOGPIE = 0
xYPDCLOG0.PDCLOGECR = 0
If nbTC_SAB_Ok = nbTC_SAB Then
    xYPDCLOG0.PDCLOGNAT = "5X"
Else
    xYPDCLOG0.PDCLOGNAT = "5X?"
End If

xYPDCLOG0.PDCLOGTXT = "TC BIA_PDC / SAB (opé saisies & rapprochées) : " & nbTC_SAB_Ok & " / " & nbTC_SAB
Call YPDCLOG0_AddItem
If nbTC_PDC <> nbTC_PDC_Ok Then
    xYPDCLOG0.PDCLOGNAT = "5X?"
    xYPDCLOG0.PDCLOGTXT = "TC BIA_PDC (opé saisies & non rapprochées)  : " & (nbTC_PDC - nbTC_PDC_Ok) & " / " & nbTC_PDC
    Call YPDCLOG0_AddItem
End If
xYPDCLOG0.PDCLOGPIE = 0
xYPDCLOG0.PDCLOGECR = 0
If nbXX_SAB_Ok = nbXX_SAB Then
    xYPDCLOG0.PDCLOGNAT = "5X"
Else
    xYPDCLOG0.PDCLOGNAT = "5X?"
End If
xYPDCLOG0.PDCLOGTXT = "XX BIA_PDC / SAB (opé saisies & rapprochées) : " & nbXX_SAB_Ok & " / " & nbXX_SAB
Call YPDCLOG0_AddItem
If nbXX_PDC <> nbXX_PDC_Ok Then
    xYPDCLOG0.PDCLOGNAT = "5X?"
    xYPDCLOG0.PDCLOGTXT = "XX BIA_PDC (opé non rapprochées /opé saisies)  : " & (nbXX_PDC - nbXX_PDC_Ok) & " / " & nbXX_PDC
    Call YPDCLOG0_AddItem
End If


End Sub

Public Function cmdPDCOPE_Transaction()
Dim V, X As String, xSQL As String
Dim Nb As Long
Dim mMsgBox As String
Dim blnInsert As Boolean
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdPDCOPE_Transaction"
'-------------------------------------------------------
cmdPDCOPE_Transaction = Null
If newYPDCOPE0.PDCOPEID = 0 Then
    blnInsert = True
    X = "select count(*) as Tally  from " & paramIBM_Library_SABSPE_XXX & ".YPDCOPE0 " _
        & " where PDCOPEDTR = '" & newYPDCOPE0.PDCOPEDTR & "'"
    Set rsSab = cnsab.Execute(X)
    newYPDCOPE0.PDCOPEID = rsSab("Tally") + 1
Else
    blnInsert = False
End If

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
If blnInsert Then
    newYPDCOPE0.PDCOPEIUSR = usrName
    newYPDCOPE0.PDCOPEIAMJ = DSys
    newYPDCOPE0.PDCOPEIHMS = time_Hms

    V = sqlYPDCOPE0_Insert(newYPDCOPE0)
Else
    newYPDCOPE0.PDCOPEVUSR = usrName
    newYPDCOPE0.PDCOPEVAMJ = DSys
    newYPDCOPE0.PDCOPEVHMS = time_Hms
    V = sqlYPDCOPE0_Update(newYPDCOPE0, oldYPDCOPE0)
End If

If Not IsNull(V) Then GoTo Error_MsgBox
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        Call cmdSendMail_Cours
    End If
    
    cmdPDCOPE_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function

Public Function cmdSuspens_Transaction(lFct As String)
Dim V, X As String, xSQL As String
Dim Nb As Long
Dim mMsgBox As String
Dim blnInsert As Boolean, blnContrepassation As Boolean
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdSuspens_Transaction"
'-------------------------------------------------------
cmdSuspens_Transaction = Null
If newYPDCMVT0.PDCMVTECR = 0 Then
    blnInsert = True
    X = "select PDCMVTECR  from " & paramIBM_Library_SABSPE_XXX & ".YPDCMVT0 " _
        & " where PDCMVTDTR = '" & newYPDCMVT0.PDCMVTDTR & "' and PDCMVTOPEC like 'XX%' order by PDCMVTECR desc"
    Set rsSab = cnsab.Execute(X)
    
    If rsSab.EOF Then
        newYPDCMVT0.PDCMVTECR = 1
    Else
        newYPDCMVT0.PDCMVTECR = rsSab("PDCMVTECR") + 1
    End If
    If newYPDCMVT0.PDCMVTOPEN <> 0 Then
        blnContrepassation = True
    Else
        blnContrepassation = False
        X = "select PDCMVTECR  from " & paramIBM_Library_SABSPE_XXX & ".YPDCMVT0 " _
            & " where PDCMVTOPEC like 'XX%'and PDCMVTSTA = '+'  order by PDCMVTECR desc"
        Set rsSab = cnsab.Execute(X)
        newYPDCMVT0.PDCMVTOPEN = rsSab("PDCMVTECR") + 1
    End If
Else
    blnInsert = False
End If

rsYPDCLOG0_Init newYPDCLOG0
X = Time
mPDCLOGUSEQ = mPDCLOGUSEQ + 1
xYPDCLOG0.PDCLOGUSEQ = mPDCLOGUSEQ

With newYPDCLOG0
    .PDCLOGDTR = newYPDCMVT0.PDCMVTDTR
    .PDCLOGUAMJ = DSys
    .PDCLOGUHMS = Mid$(X, 1, 2) + Mid$(X, 4, 2) + Mid$(X, 7, 2)
    .PDCLOGUUSR = usrName
    .PDCLOGUSEQ = mPDCLOGUSEQ
    .PDCLOGPIE = newYPDCMVT0.PDCMVTPIE
    .PDCLOGECR = newYPDCMVT0.PDCMVTECR
    .PDCLOGNAT = "4" & newYPDCMVT0.PDCMVTSTA
    .PDCLOGTXT = newYPDCMVT0.PDCMVTCPT & "  " & Trim(Format$(newYPDCMVT0.PDCMVTMTE, "### ### ### ##0.00")) & " EUR/" & newYPDCMVT0.PDCMVTDEV & " " & Trim(Format$(newYPDCMVT0.PDCMVTMTD, "### ### ### ##0.00"))
End With

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
If lFct = constModifier Then
    V = sqlYPDCMVT0_Update(newYPDCMVT0, oldYPDCMVT0)
Else
    If lFct = constDelete Then
        xSQL = " where PDCMVTDTR = '" & newYPDCMVT0.PDCMVTDTR & "' and PDCMVTPIE = " & newYPDCMVT0.PDCMVTPIE & " and PDCMVTECR = " & newYPDCMVT0.PDCMVTECR
    
        V = sqlYPDCMVT0_DeleteW(xSQL, Nb)
        If IsNull(V) Then V = cmdSuspens_Transaction_YPDCPOS0(lFct)
        If IsNull(V) Then
                newYPDCLOG0.PDCLOGNAT = "4X"
                newYPDCLOG0.PDCLOGTXT = "effacement " & newYPDCLOG0.PDCLOGTXT
                V = sqlYPDCLOG0_Insert(newYPDCLOG0)
        End If
    
    Else
        If blnInsert Then
        
            V = sqlYPDCMVT0_Insert(newYPDCMVT0)
            If IsNull(V) Then V = cmdSuspens_Transaction_YPDCPOS0(lFct)
            If IsNull(V) Then V = sqlYPDCLOG0_Insert(newYPDCLOG0)
            If blnContrepassation Then
                If IsNull(V) Then V = sqlYPDCMVT0_Update(memoYPDCMVT0, oldYPDCMVT0)
            End If
        Else
            V = sqlYPDCMVT0_Update(newYPDCMVT0, oldYPDCMVT0)
        End If
    End If
End If
If Not IsNull(V) Then GoTo Error_MsgBox

If Trim(txtSuspens_Comment) <> "" Then
    mPDCLOGUSEQ = mPDCLOGUSEQ + 1
    newYPDCLOG0.PDCLOGUSEQ = mPDCLOGUSEQ
    newYPDCLOG0.PDCLOGNAT = "4c"
    newYPDCLOG0.PDCLOGTXT = Trim(txtSuspens_Comment)
    V = sqlYPDCLOG0_Insert(newYPDCLOG0)
End If

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    cmdSuspens_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function
Public Function cmdYPDCMVT0_Transaction(lFct As String)
Dim V, X As String, xSQL As String
Dim Nb As Long
Dim mMsgBox As String
Dim blnInsert As Boolean, blnContrepassation As Boolean
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdYPDCMVT0_Transaction"
'-------------------------------------------------------
cmdYPDCMVT0_Transaction = Null

rsYPDCLOG0_Init newYPDCLOG0
X = Time
mPDCLOGUSEQ = mPDCLOGUSEQ + 1
xYPDCLOG0.PDCLOGUSEQ = mPDCLOGUSEQ

With newYPDCLOG0
    .PDCLOGDTR = newYPDCMVT0.PDCMVTDTR
    .PDCLOGUAMJ = DSys
    .PDCLOGUHMS = Mid$(X, 1, 2) + Mid$(X, 4, 2) + Mid$(X, 7, 2)
    .PDCLOGUUSR = usrName
    .PDCLOGUSEQ = mPDCLOGUSEQ
    .PDCLOGPIE = newYPDCMVT0.PDCMVTPIE
    .PDCLOGECR = newYPDCMVT0.PDCMVTECR
    .PDCLOGNAT = "4M*"
    .PDCLOGTXT = newYPDCMVT0.PDCMVTCPT & "  " & Trim(Format$(newYPDCMVT0.PDCMVTMTE, "### ### ### ##0.00")) & " EUR/" & Trim(Format$(newYPDCMVT0.PDCMVTDEV & " " & newYPDCMVT0.PDCMVTMTD, "### ### ### ##0.00"))
End With

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
If lFct = constModifier Then
    V = sqlYPDCMVT0_Update(newYPDCMVT0, oldYPDCMVT0)
End If
If Not IsNull(V) Then GoTo Error_MsgBox
If newYPDCMVT0.PDCMVTKCUT = "*" Then
    newYPDCLOG0.PDCLOGTXT = "TOP mvt = non soldé " & newYPDCLOG0.PDCLOGTXT
Else
    newYPDCLOG0.PDCLOGTXT = "annulation TOP mvt = non soldé " & newYPDCLOG0.PDCLOGTXT
End If
V = sqlYPDCLOG0_Insert(newYPDCLOG0)

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    cmdYPDCMVT0_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function

Public Function cmdSuspens_Transaction_YPDCPOS0(lFct As String)
Dim V, X As String, xSQL As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdSuspens_Transaction_YPDCPOS0"
'-------------------------------------------------------
cmdSuspens_Transaction_YPDCPOS0 = Null
'________________________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCPOS0 where PDCPOSDTR = '" & newYPDCMVT0.PDCMVTDTR & "' and PDCPOSDEV = '" & newYPDCMVT0.PDCMVTDEV & "'"
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    V = xSQL
Else
    V = rsYPDCPOS0_GetBuffer(rsSab, oldYPDCPOS0)
    newYPDCPOS0 = oldYPDCPOS0
'_____________________________________________________________________________________
    If newYPDCMVT0.PDCMVTOPEC = "XXC" Then
        If lFct = constDelete Then
            newYPDCPOS0.PDCPOSPOSE = newYPDCPOS0.PDCPOSPOSE - newYPDCMVT0.PDCMVTMTE
            newYPDCPOS0.PDCPOSPOSD = newYPDCPOS0.PDCPOSPOSD - newYPDCMVT0.PDCMVTMTD
        Else
            newYPDCPOS0.PDCPOSPOSE = newYPDCPOS0.PDCPOSPOSE + newYPDCMVT0.PDCMVTMTE
            newYPDCPOS0.PDCPOSPOSD = newYPDCPOS0.PDCPOSPOSD + newYPDCMVT0.PDCMVTMTD
        End If
        If newYPDCPOS0.PDCPOSPOSE = 0 Then
            newYPDCPOS0.PDCPOSPRIX = 0
        Else
            newYPDCPOS0.PDCPOSPRIX = Round(Abs(newYPDCPOS0.PDCPOSPOSD / newYPDCPOS0.PDCPOSPOSE), 6)
        End If
    Else
        If lFct = constDelete Then
            newYPDCPOS0.PDCPOSTERE = newYPDCPOS0.PDCPOSTERE - newYPDCMVT0.PDCMVTMTE
            newYPDCPOS0.PDCPOSTERD = newYPDCPOS0.PDCPOSTERD - newYPDCMVT0.PDCMVTMTD
        Else
            newYPDCPOS0.PDCPOSTERE = newYPDCPOS0.PDCPOSTERE + newYPDCMVT0.PDCMVTMTE
            newYPDCPOS0.PDCPOSTERD = newYPDCPOS0.PDCPOSTERD + newYPDCMVT0.PDCMVTMTD
        End If
    End If
    
'_____________________________________________________________________________________________
    V = sqlYPDCPOS0_Update(newYPDCPOS0, oldYPDCPOS0)
End If
If Not IsNull(V) Then GoTo Error_MsgBox
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    cmdSuspens_Transaction_YPDCPOS0 = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function


Private Sub cmdPrint_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
SSTab1.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Etat : " & fgSelect.Rows - 1)

cmdPrint_Ok_Etat

fraPrint.Visible = False
SSTab1.Visible = True
Me.Show
Me.Enabled = True: Me.MousePointer = 0
cmdSelect_Ok_Click

End Sub

Private Sub cmdPrint_Quit_Click()
fraPrint.Visible = False

End Sub

Private Sub cmdReport_Quit_Click()
fraReport.Visible = False
End Sub

Private Sub cmdReport_Update_Click()
Dim V
Dim I As Integer


Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "> cmdReport_Update :début du traitement"): DoEvents
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
newYPDCOPE0 = oldYPDCOPE0
newYPDCOPE0.PDCOPESTA = "A"
newYPDCOPE0.PDCOPEVUSR = usrName
newYPDCOPE0.PDCOPEVAMJ = DSys
newYPDCOPE0.PDCOPEVHMS = time_Hms
newYPDCOPE0.PDCOPEVTXT = "ANNULATION de l'opération reportée automatiquement"
V = sqlYPDCOPE0_Update(newYPDCOPE0, oldYPDCOPE0)
If Not IsNull(V) Then GoTo Error_MsgBox

For I = 1 To selYPDCOPE0_Nb
    xYPDCOPE0 = selYPDCOPE0(I)
    newYPDCOPE0 = xYPDCOPE0
    newYPDCOPE0.PDCOPESTA = "A"
    newYPDCOPE0.PDCOPEVUSR = usrName
    newYPDCOPE0.PDCOPEVAMJ = DSys
    newYPDCOPE0.PDCOPEVHMS = time_Hms
    newYPDCOPE0.PDCOPEVTXT = "ANNULATION du REPORT automatique"
    V = sqlYPDCOPE0_Update(newYPDCOPE0, xYPDCOPE0)
    If Not IsNull(V) Then GoTo Error_MsgBox

Next I

GoTo Exit_sub
'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        Call cmdSendMail_Report_Annulation
        fraPDCOPE.Visible = False
        fraReport.Visible = False
        cmdSelect_SQL_1
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Call lstErr_AddItem(lstErr, cmdContext, "< cmdReport_Update : fin du Traitement"): DoEvents


Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_PDC_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
fgDetail.Clear
fgDetail.Visible = False
If blnOk Then
    'cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True 'False
    Select Case cmdSelect_SQL_K
        Case "1":   cmdSelect_SQL_1
        Case "2":   cmdSelect_SQL_2
        Case "3":   cmdSelect_SQL_3
        Case "4":   cmdSelect_SQL_4
        Case "5":   cmdSelect_SQL_5
        Case "7":   cmdSelect_SQL_7
        Case "9":   cmdSelect_SQL_9
        Case "X":  cmdSelect_SQL_xls
        Case "Y":  cmdSelect_SQL_Y
    End Select
Else
    'cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_PDC_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

'Me.Enabled = False: Me.MousePointer = vbHourglass
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdPDCOPE_Update_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

If IsNull(fraPDCOPE_Control) Then
    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    V = cmdPDCOPE_Transaction
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        'cmdPDCOPE.Visible = BIA_PDC_Aut.Saisir
        fraPDCOPE.Visible = False
        Call DTPicker_Set(txtSelect_AMJ, YBIATAB0_DATE_CPT_JS1)

        cmdSelect_SQL_1
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdPDCOPE_Update"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0



End Sub

Public Function fraPDCOPE_Control()
Dim X As String, wMsg As String, xSQL As String
Dim K As Integer

blnPDCOPE_Control_Ok = True
Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données "): DoEvents
xYPDCOPE0 = oldYPDCOPE0
fraPDCOPE_Control_DEV_Certain

wMsg = ""
If blnPDCOPE_Control_S Then
    xYPDCOPE0.PDCOPEREF = Val(txtPDCOPEREF)
    xYPDCOPE0.PDCOPEOPEC = cboPDCOPEOPEC
    xYPDCOPE0.PDCOPEOPEN = Val(txtPDCOPEOPEN)
    xYPDCOPE0.PDCOPEOPET = cboPDCOPEOPET
    xYPDCOPE0.PDCOPESENS = Mid$(cboPDCOPESENS, 1, 1)
    xYPDCOPE0.PDCOPEDEV1 = cboPDCOPEDEV1
    xYPDCOPE0.PDCOPEMTD1 = CCur(num_CDec(txtPDCOPEMTD1))
    xYPDCOPE0.PDCOPEDEV2 = cboPDCOPEDEV2
    
    If xYPDCOPE0.PDCOPEDEV1 = "   " Then
        blnPDCOPE_Control_Ok = False
        wMsg = wMsg & "- préciser la devise principale" & vbCrLf
    End If
    If xYPDCOPE0.PDCOPEDEV2 = "   " Then
        blnPDCOPE_Control_Ok = False
        wMsg = wMsg & "- préciser la devise secondaire" & vbCrLf
    End If

    If xYPDCOPE0.PDCOPEDEV1 = xYPDCOPE0.PDCOPEDEV2 Then
        blnPDCOPE_Control_Ok = False
        wMsg = wMsg & "- les deux devises doivent être différentes " & xYPDCOPE0.PDCOPEDEV1 & vbCrLf
    End If
    
    If xYPDCOPE0.PDCOPEDEV1 = "EUR" Then
        xYPDCOPE0.PDCOPESENX = "1"
    Else
        If xYPDCOPE0.PDCOPEDEV2 = "EUR" Then
            xYPDCOPE0.PDCOPESENX = "2"
        Else
            xYPDCOPE0.PDCOPESENX = "3"
            'blnPDCOPE_Control_Ok = False
            'wMsg = wMsg & "- une des deux devises doit être EUR" & vbCrLf
        End If
    End If
    Call DTPicker_Control(txtPDCOPEDVA, xYPDCOPE0.PDCOPEDVA)
    xYPDCOPE0.PDCOPECLI = Format$(txtPDCOPECLI, "0000000")
    xYPDCOPE0.PDCOPESER = Mid$(cboPDCOPESER, 1, 2)
    xYPDCOPE0.PDCOPESSE = Mid$(cboPDCOPESER, 4, 2)
    If localUnit = "GDMP" Then
        Select Case cboPDCOPESER
            Case "00 TR" ', "00 GU"
            Case Else
                blnPDCOPE_Control_Ok = False
                wMsg = wMsg & "- sous-service interdit(uniquement TR)" & vbCrLf
        End Select
    End If
    xYPDCOPE0.PDCOPEITXT = txtPDCOPEITXT
    If xYPDCOPE0.PDCOPESSE <> "TC" Then
        xYPDCOPE0.PDCOPESTA = "0"
    Else
        xYPDCOPE0.PDCOPESTA = "T"
        If xYPDCOPE0.PDCOPEREF = 0 Then
            blnPDCOPE_Control_Ok = False
            wMsg = wMsg & "- préciser le n° du ticket" & vbCrLf
        Else
            wMsg = wMsg & fraPDCOPE_Control_Ticket

        End If
    End If
    If xYPDCOPE0.PDCOPESSE = "CD" Then
        If xYPDCOPE0.PDCOPEOPEC <> "CDE" And xYPDCOPE0.PDCOPEOPEC <> "CDI" Then
            blnPDCOPE_Control_Ok = False
            wMsg = wMsg & "- Code opération : CDE ou CDI" & vbCrLf
        End If
    
        If xYPDCOPE0.PDCOPEOPEN = 0 Then
            blnPDCOPE_Control_Ok = False
            wMsg = wMsg & "- préciser le n° du dossier" & vbCrLf
        End If
    End If
    
    If xYPDCOPE0.PDCOPEMTD1 = 0 Then
        blnPDCOPE_Control_Ok = False
        wMsg = wMsg & "- préciser le montant" & vbCrLf
    End If
    If xYPDCOPE0.PDCOPEDVA < YBIATAB0_DATE_CPT_JS1 Then
        blnPDCOPE_Control_Ok = False
        wMsg = wMsg & "- date valeur < date du jour" & vbCrLf
    End If
       
    For K = 1 To arrDevF_Nb
        
        arrDevF_ISO(arrDevF_Nb) = Mid$(X, 3, 3)
        If arrDevF_AMJ(K) = xYPDCOPE0.PDCOPEDVA Then
            If arrDevF_ISO(K) = xYPDCOPE0.PDCOPEDEV1 Then
                blnPDCOPE_Control_Ok = False
                wMsg = wMsg & "- date valeur = jour férié " & xYPDCOPE0.PDCOPEDEV1 & vbCrLf
            End If
            If arrDevF_ISO(K) = xYPDCOPE0.PDCOPEDEV2 Then
                blnPDCOPE_Control_Ok = False
                wMsg = wMsg & "- date valeur = jour férié " & xYPDCOPE0.PDCOPEDEV2 & vbCrLf
            End If
        End If
    Next K

    K = Weekday(dateImp(xYPDCOPE0.PDCOPEDVA))
    If K = 1 Then
        blnPDCOPE_Control_Ok = False
        wMsg = wMsg & "- date valeur = Dimanche " & vbCrLf
    Else
        If K = 7 Then
            blnPDCOPE_Control_Ok = False
            wMsg = wMsg & "- date valeur = Samedi " & vbCrLf
        End If
    End If

    If Trim(xYPDCOPE0.PDCOPECLI) = "" Then
        blnPDCOPE_Control_Ok = False
        wMsg = wMsg & "- préciser la contrepartie" & vbCrLf
    Else
        Call fraPDCOPE_Display_CLI(xYPDCOPE0.PDCOPECLI)
        If mCLIENARA1 = "" Then
            blnPDCOPE_Control_Ok = False
            wMsg = wMsg & "- contrepartie inconnue" & vbCrLf
        End If
    End If
    
End If
If chkPDCOPEINFO = "1" Then xYPDCOPE0.PDCOPESTA = "I"
If blnPDCOPE_Control_V Then
    xYPDCOPE0.PDCOPEVTXT = Trim(txtPDCOPEVTXT)
    xYPDCOPE0.PDCOPETAUX = CDbl(num_CDec(txtPDCOPETAUX))
    If xYPDCOPE0.PDCOPETAUX = 0 Then
        blnPDCOPE_Control_Ok = False
        wMsg = wMsg & "- préciser le cours" & vbCrLf
    Else
        If xYPDCOPE0.PDCOPESTA = "0" Then xYPDCOPE0.PDCOPESTA = "V"
        If xYPDCOPE0.PDCOPETAUX = 0 Then
            xYPDCOPE0.PDCOPEMTD2 = 0
        Else
            xYPDCOPE0.PDCOPEMTD2 = CCur(num_CDec(libPDCOPEMTD2))
        End If
    End If
End If
Call fraPDCOPE_Display_Montant(xYPDCOPE0.PDCOPESENS)
If blnPDCOPE_Control_Ok And BIA_PDC_Aut.Rapprocher Then
    wMsg = wMsg & fraPDCOPE_Control_memoYPDCOPE0
    wMsg = wMsg & fraPDCOPE_Control_ZCHGOPE0
End If


'__________________________________________________________________________

If blnPDCOPE_Control_Ok Then
    fraPDCOPE_Control = Null
    If Not BIA_PDC_Aut.Rapprocher Then
        newYPDCOPE0 = xYPDCOPE0
    Else
         oldYPDCOPE0 = memoYPDCOPE0
         newYPDCOPE0 = memoYPDCOPE0
         newYPDCOPE0.PDCOPESTA3 = "B"
         newYPDCOPE0.PDCOPEOPEN = xYPDCOPE0.PDCOPEOPEN
         newYPDCOPE0.PDCOPEVTXT = xYPDCOPE0.PDCOPEVTXT
   End If

Else
    Call MsgBox(wMsg, vbCritical, "BIA_PDC : Gestion des opérations du jour")
    fraPDCOPE_Control = "?_________fraPDCOPE_Control"
End If

End Function


Private Sub cmdPDCOPE_Quit_Click()
fraReport.Visible = False
fraPDCOPE.Visible = False

End Sub


Private Sub cmdSelect_Ok_xls_Click()
Dim V
On Error Resume Next
If Not fraSelect_Comment_Xls.Visible Then
    If Trim(oldYPDCMAIL.PDCMAILTXT) <> "" Then
        X = MsgBox("Contrôle déjà validé, voulez-vous l'ANNULER et le REMPLACER ?", vbYesNo, "BIA_PDC : contrôle PDC / BOTC.xls")
        If X = vbNo Then Exit Sub
        V = sqlYPDCMAIL_DeleteW(" where PDCMAILDTR ='" & oldYPDCMAIL.PDCMAILDTR & "'")
        If Not IsNull(V) Then Error_Route (V)

    End If
    cmdSelect_Ok_xls.Caption = "envoyer le mail"
    cmdSelect_Ok_xls.BackColor = &H80C0FF
    fraSelect_Comment_Xls.Visible = True
    txtSelect_Comment_xls.SetFocus
    Exit Sub
End If
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BOTC_xls Début du traitement"): DoEvents

fraSelect_Comment_Xls.Visible = False

cmdSendMail_xls
cmdSelect_Ok_xls.Caption = "mail envoyé"

Call lstErr_AddItem(lstErr, cmdContext, "< BOTC_xls Fin du Traitement"): DoEvents
'cmdSelect_SQL_1
cboSelect_SQL.ListIndex = 0

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSuspens_Annulation_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

    Call lstErr_AddItem(lstErr, cmdContext, ">_________Effacement des données "): DoEvents
    newYPDCMVT0 = oldYPDCMVT0
    V = cmdSuspens_Transaction(constDelete)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        fraSuspens.Visible = False

        cmdSelect_SQL_4
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdSuspens_Annulation"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSuspens_Modification_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

If IsNull(fraSuspens_Control) Then
    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    newYPDCMVT0 = xYPDCMVT0
    V = cmdSuspens_Transaction(constModifier)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        fraSuspens.Visible = False

        cmdSelect_SQL_4
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdSuspens_Annulation"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdSuspens_Quit_Click()
fraSuspens.Visible = False

End Sub


Private Sub cmdSuspens_Update_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

If IsNull(fraSuspens_Control) Then
    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    V = cmdSuspens_Transaction(constUpdate)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        fraSuspens.Visible = False
        cmdSelect_SQL_4
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdSUspens_Update"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0



End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
Dim xSQL As String
On Error Resume Next
If y <= fgDetail.RowHeightMin Then
    fgDetail.Visible = False
    Select Case fgDetail.Col
        Case 0: fgDetail_Sort1 = 0: fgDetail_Sort2 = 0: fgDetail_SortX 0
        Case 1:  fgDetail_Sort1 = 1: fgDetail_Sort2 = 1: fgDetail_SortX 1
        Case 2: fgDetail_Sort1 = 2: fgDetail_Sort2 = 2: fgDetail_SortX 2
        Case 3: fgDetail_Sort1 = 3: fgDetail_Sort2 = 3: fgDetail_SortX 3
        Case 4: fgDetail_Sort1 = 4: fgDetail_Sort2 = 4: fgDetail_SortX 4
        Case 5: fgDetail_Sort1 = 5: fgDetail_Sort2 = 5: fgDetail_Sort
        Case 6: fgDetail_Sort1 = 6: fgDetail_Sort2 = 6: fgDetail_SortX 6
        Case 7: fgDetail_Sort1 = 7: fgDetail_Sort2 = 7: fgDetail_Sort
        Case 8: fgDetail_Sort1 = 8: fgDetail_Sort2 = 8: fgDetail_SortX 8
        Case 9: fgDetail_Sort1 = 9: fgDetail_Sort2 = 9: fgDetail_Sort
       Case fgDetail_arrIndex:  fgDetail_SortX fgDetail_arrIndex
    End Select
    fgDetail.Visible = True
Else
    If cmdSelect_SQL_K = "4" And fgDetail.Rows > 1 Then
        Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
        fgDetail.Col = fgDetail_arrIndex:  arrYPDCMVT0_Index = CLng(fgDetail.Text)
        oldYPDCMVT0 = arrYPDCMVT0(arrYPDCMVT0_Index)
        xYPDCMVT0 = oldYPDCMVT0
        fraSuspens_Display
   End If
   If fgDetail.FormatString = fgDetail_FormatString And BIA_PDC_Aut.Valider Then
        Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
        fgDetail.Col = fgDetail_arrIndex:  arrYPDCMVT0_Index = CLng(fgDetail.Text)
        oldYPDCMVT0 = arrYPDCMVT0(arrYPDCMVT0_Index)
        newYPDCMVT0 = oldYPDCMVT0
        If oldYPDCMVT0.PDCMVTKCUT = " " Then
           newYPDCMVT0.PDCMVTKCUT = "*"
           mnuPDCMVTKCUT_Update.Caption = "TOPER ce mvt comptable comme non soldé"
        Else
           newYPDCMVT0.PDCMVTKCUT = " "
           mnuPDCMVTKCUT_Update.Caption = "ANNULER le top : mvt comptable non soldé"
        End If
       If Mid$(oldYPDCMVT0.PDCMVTOPEC, 1, 2) <> "XX" Then Me.PopupMenu mnuPDCMVTKCUT, vbPopupMenuLeftButton
   End If
End If
fgDetail.LeftCol = -1
fgDetail.LeftCol = 0

End Sub


Private Sub fgPDCOPE_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
Dim xSQL As String
On Error Resume Next
If y <= fgPDCOPE.RowHeightMin Then
    fgPDCOPE.Visible = False
    Select Case fgPDCOPE.Col
        Case 0: fgPDCOPE_Sort1 = 0: fgPDCOPE_Sort2 = 2: fgPDCOPE_Sort
        Case 1:  fgPDCOPE_Sort1 = 1: fgPDCOPE_Sort2 = 1: fgPDCOPE_SortX 1
        Case 2: fgPDCOPE_Sort1 = 2: fgPDCOPE_Sort2 = 2: fgPDCOPE_SortX 2
        Case 3: fgPDCOPE_Sort1 = 3: fgPDCOPE_Sort2 = 3: fgPDCOPE_SortX 3
        Case 4: fgPDCOPE_Sort1 = 4: fgPDCOPE_Sort2 = 4: fgPDCOPE_SortX 4
        Case 5: fgPDCOPE_Sort1 = 5: fgPDCOPE_Sort2 = 5: fgPDCOPE_Sort
        Case 6: fgPDCOPE_Sort1 = 6: fgPDCOPE_Sort2 = 6: fgPDCOPE_SortX 6
        Case 7: fgPDCOPE_Sort1 = 7: fgPDCOPE_Sort2 = 7: fgPDCOPE_SortX 7
        Case 8: fgPDCOPE_Sort1 = 8: fgPDCOPE_Sort2 = 8: fgPDCOPE_SortX 8
        Case 9: fgPDCOPE_Sort1 = 9: fgPDCOPE_Sort2 = 9: fgPDCOPE_Sort
       Case fgPDCOPE_arrIndex:  fgPDCOPE_SortX fgPDCOPE_arrIndex
    End Select
    fgPDCOPE.Visible = True
Else
    If fgPDCOPE.Rows > 1 And chkSelect_Ope = "1" Then
        Call fgPDCOPE_Color(fgPDCOPE_RowClick, MouseMoveUsr.BackColor, fgPDCOPE_ColorClick)
        fgPDCOPE.Col = fgPDCOPE_arrIndex:  arrYPDCOPE0_Index = CLng(fgPDCOPE.Text)
        oldYPDCOPE0 = arrYPDCOPE0(arrYPDCOPE0_Index)
        xYPDCOPE0 = oldYPDCOPE0
        fraPDCOPE_Display
        
   End If
End If
fgPDCOPE.LeftCol = -1
fgPDCOPE.LeftCol = 0

End Sub


Public Sub fraPDCOPE_Display()
Dim V, K As Integer
Dim X As String, X1 As String
Dim blnPDCOPEDTR As Boolean, blnAnnulation As Boolean
On Error GoTo Error_Handler


fraPDCOPE.Visible = False
fraReport.Visible = False

blnPDCOPE_Control_Ok = True
If xYPDCOPE0.PDCOPEDTR = YBIATAB0_DATE_CPT_JS1 Then
    blnPDCOPEDTR = True
Else
    blnPDCOPEDTR = False
End If
fraPDCOPE_R.Visible = True
fraPDCOPE_S.Visible = True
fraPDCOPE_V.Visible = True

'fraPDCOPE_R.Enabled = False
'fraPDCOPE_S.Enabled = False
'fraPDCOPE_V.Enabled = False

cmdPDCOPE_Quit.Visible = True
cmdPDCOPE_Update.Visible = False
cmdPDCOPE_Update_Ref.Visible = False
cmdPDCOPE_Annulation.Visible = False
 Select Case xYPDCOPE0.PDCOPESTA
    Case Is = "V", "T": fraPDCOPE.BackColor = &HE0FFE0
    Case Is = "A": fraPDCOPE.BackColor = &H8000000F
    Case Is = "I": fraPDCOPE.BackColor = &HE0E000
    Case Is = "0", "B": fraPDCOPE.BackColor = &HEFEFFF
    Case Else: fraPDCOPE.BackColor = &HC0FFFF
End Select

fraPDCOPE_R.BackColor = fraPDCOPE.BackColor
Call usrColor_Container(fraPDCOPE_R, fraPDCOPE.BackColor)
fraPDCOPE_S.BackColor = fraPDCOPE.BackColor
Call usrColor_Container(fraPDCOPE_S, fraPDCOPE.BackColor)
fraPDCOPE_V.BackColor = fraPDCOPE.BackColor
Call usrColor_Container(fraPDCOPE_V, fraPDCOPE.BackColor)

libPDCOPEID = dateImp10(xYPDCOPE0.PDCOPEDTR) & " - " & xYPDCOPE0.PDCOPEID
libPDCOPEIUSR = xYPDCOPE0.PDCOPEIUSR & "  " & dateImp10(xYPDCOPE0.PDCOPEIAMJ) & "  " & timeImp8(xYPDCOPE0.PDCOPEIHMS)
libPDCOPEVUSR = xYPDCOPE0.PDCOPEVUSR & "  " & dateImp10(xYPDCOPE0.PDCOPEVAMJ) & "  " & timeImp8(xYPDCOPE0.PDCOPEVHMS)
chkPDCOPEINFO.Visible = False ' = "0"
chkPDCOPEINFO.Value = "0"

blnAnnulation = False
If blnPDCOPEDTR And BIA_PDC_Aut.Saisir Then
    Select Case localUnit
        Case "FOTC": blnAnnulation = BIA_PDC_Aut.Valider ' True
        Case "GDC"
        Case "GDMP"
            If xYPDCOPE0.PDCOPESSE = "TR" Or xYPDCOPE0.PDCOPESSE = "MP" Or xYPDCOPE0.PDCOPESSE = "GU" Then blnAnnulation = True
        Case "SOBI"
            If xYPDCOPE0.PDCOPESSE = "CD" Then blnAnnulation = True
        Case "CPT"
            If xYPDCOPE0.PDCOPESSE = "CP" Then blnAnnulation = True
    End Select
End If

Select Case xYPDCOPE0.PDCOPESTA
    Case "T": libPDCOPESTA = "FOTC": cmdPDCOPE_Annulation.Visible = blnAnnulation
            If BIA_PDC_Aut.Valider And blnPDCOPEDTR And xYPDCOPE0.PDCOPESTA2 = " " And xYPDCOPE0.PDCOPESTA3 = " " Then
                cmdPDCOPE_Update_Ref.Caption = "Modifier le n° du ticket"
                cmdPDCOPE_Update_Ref.Visible = True
            End If
    Case "V": libPDCOPESTA = "demande traitée": If localUnit = "FOTC" Then cmdPDCOPE_Annulation.Visible = blnAnnulation
    Case "0": libPDCOPESTA = "demande en attente": cmdPDCOPE_Annulation.Visible = blnAnnulation
    Case "A": libPDCOPESTA = "demande annulée"
    Case "I": libPDCOPESTA = "pour information": cmdPDCOPE_Annulation.Visible = blnAnnulation
                chkPDCOPEINFO.BackColor = vbMagenta '&HFF
                chkPDCOPEINFO.Caption = "POUR INFORMATION"
                chkPDCOPEINFO = "1"
                chkPDCOPEINFO.Visible = True
    Case "B": libPDCOPESTA = "BOTC": cmdPDCOPE_Annulation.Visible = blnAnnulation
    Case " ": libPDCOPESTA = "saisie en cours"
    Case "R": libPDCOPESTA = "opération reportée"
              If paramEnvironnement = constProduction Then
                    If xYPDCOPE0.PDCOPEDTR = YBIATAB0_DATE_CPT_JP0 _
                    And localUnit = "GDC" _
                    And BIA_PDC_Aut.Saisir Then
                        cmdPDCOPE_Annulation.Visible = True
                    End If
              Else
                    If BIA_PDC_Aut.Saisir Then cmdPDCOPE_Annulation.Visible = True
              End If
    Case Else: libPDCOPESTA = xYPDCOPE0.PDCOPESTA
End Select

If blnPDCOPE_CONF_CALL_Saisie Then fraPDCOPE.BackColor = RGB(0, 128, 0): libPDCOPESTA = "saisie CONF CALL"

Select Case xYPDCOPE0.PDCOPESTA3
    Case "B": lblPDCOPESTA3 = "contrôle BOTC : ok"
        If BIA_PDC_Aut.Rapprocher And blnPDCOPEDTR And xYPDCOPE0.PDCOPESTA2 = " " And xYPDCOPE0.PDCOPESTA = "T" Then
            cmdPDCOPE_Update_Ref.Caption = "Modifier le n° d'opération SAB"
            cmdPDCOPE_Update_Ref.Visible = True
        End If

    Case Else: lblPDCOPESTA3 = xYPDCOPE0.PDCOPESTA3
End Select
lblPDCOPESTA2 = xYPDCOPE0.PDCOPESTA2
cboPDCOPESER = xYPDCOPE0.PDCOPESER & " " & xYPDCOPE0.PDCOPESSE

txtPDCOPEREF = IIf(xYPDCOPE0.PDCOPEREF = 0, "", xYPDCOPE0.PDCOPEREF)
cboPDCOPEOPEC = xYPDCOPE0.PDCOPEOPEC
'cboPDCOPEOPET = xYPDCOPE0.PDCOPEOPET
txtPDCOPEOPEN = IIf(xYPDCOPE0.PDCOPEOPEN = 0, "", xYPDCOPE0.PDCOPEOPEN)
Call fraPDCOPE_Display_Montant(xYPDCOPE0.PDCOPESENS)

Call cbo_Scan(xYPDCOPE0.PDCOPEDEV1, cboPDCOPEDEV1)

Call cbo_Scan(xYPDCOPE0.PDCOPEDEV2, cboPDCOPEDEV2)
'cboPDCOPEDEV1 = xYPDCOPE0.PDCOPEDEV1
'cboPDCOPEDEV2 = xYPDCOPE0.PDCOPEDEV2
Call fraPDCOPE_Display_CLI(xYPDCOPE0.PDCOPECLI)
Call DTPicker_Set(txtPDCOPEDVA, xYPDCOPE0.PDCOPEDVA)

lblPDCOPEFIXING = "Fixing"
txtPDCOPEFIXING = ""


If xYPDCOPE0.PDCOPETAUX = 0 Then
    txtPDCOPETAUX = ""
Else
    txtPDCOPETAUX = Format$(xYPDCOPE0.PDCOPETAUX, "### ##.000000")
    If xYPDCOPE0.PDCOPEDEV1 = "EUR" Then
        Call YBIATAB0_Fixing(xYPDCOPE0.PDCOPEDEV2, xYPDCOPE0.PDCOPEDTR)
        lblPDCOPEFIXING = "Fixing au " & dateImp10(mFixing_AMJ)
        txtPDCOPEFIXING = Format$(mFIXING_Cours, "### ##.000000")
    Else
        If xYPDCOPE0.PDCOPEDEV2 = "EUR" Then
            Call YBIATAB0_Fixing(xYPDCOPE0.PDCOPEDEV1, xYPDCOPE0.PDCOPEDTR)
            lblPDCOPEFIXING = "Fixing au " & dateImp10(mFixing_AMJ)
            txtPDCOPEFIXING = Format$(mFIXING_Cours, "### ##.000000")
        End If
    End If
End If

libPDCOPETAUX = Trim(lblPDCOPETAUX) & " " & txtPDCOPETAUX
txtPDCOPEITXT = Trim(xYPDCOPE0.PDCOPEITXT)
txtPDCOPEVTXT = Trim(xYPDCOPE0.PDCOPEVTXT)

'If xYPDCOPE0.PDCOPEOPEC = "SWP" Then
'    lblPDCOPEFIXING.Caption = "cours terme"
'    txtPDCOPEFIXING = ""
'End If
Call lstErr_Clear(lstErr, cmdContext, ">Affichage du détail d'une opération"): DoEvents

fraPDCOPE_R.Enabled = False
fraPDCOPE_S.Enabled = False
fraPDCOPE_V.Enabled = False

fraPDCOPE.Visible = True
If BIA_PDC_Aut.Valider Then
    If xYPDCOPE0.PDCOPESTA = "0" Or xYPDCOPE0.PDCOPESTA = "I" Then
        fraPDCOPE_V.Enabled = blnPDCOPEDTR
        cmdPDCOPE_Update.Visible = blnPDCOPEDTR
        If txtPDCOPETAUX.Enabled Then txtPDCOPETAUX.SetFocus
    End If
End If

If BIA_PDC_Aut.Rapprocher Then cmdPDCOPE_Update.Visible = blnPDCOPEDTR ': txtPDCOPETAUX.SetFocus
fraPDCOPE.ZOrder 0
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

fraPDCOPE.Visible = True
fraPDCOPE.ZOrder 0
'______________

End Sub

Public Sub fraSuspens_Display()
Dim V
Dim X As String, X1 As String
On Error GoTo Error_Handler


blnSuspens_Control_Ok = True
fraSuspens_S.Enabled = False
fraSuspens_M.Enabled = False
cmdSuspens_Quit.Visible = True
cmdSuspens_Update.Visible = False
cmdSuspens_Modification.Visible = False
cmdSuspens_Annulation.Visible = False
optSuspens_XXX.Visible = False
optSuspens_XXX.ForeColor = vbBlue
optSuspens_XXC.ForeColor = vbBlue
optSuspens_XXT.ForeColor = vbBlue

Call lstErr_Clear(lstErr, cmdContext, ">Affichage du détail d'une opération"): DoEvents
Select Case xYPDCMVT0.PDCMVTOPEC
    Case "XXC": optSuspens_XXC = True
    Case "XXT": optSuspens_XXT = True
    Case Else: optSuspens_XXX = True
End Select


If xYPDCMVT0.PDCMVTSTA = "+" Then
    If xYPDCMVT0.PDCMVTSTA2 = " " Then
        libSuspens_PDCMVTDTR.ForeColor = vbBlue
        libSuspens_PDCMVTDTR = "Suspens FOTC : " & dateImp10(xYPDCMVT0.PDCMVTDTR)
    Else
        libSuspens_PDCMVTDTR.ForeColor = vbMagenta
        libSuspens_PDCMVTDTR = "Suspens FOTC contrepassé: " & dateImp10(xYPDCMVT0.PDCMVTDTR)
    End If
Else
        libSuspens_PDCMVTDTR.ForeColor = vbRed
        libSuspens_PDCMVTDTR = "Contrepassation d'un suspens FOTC : " & dateImp10(xYPDCMVT0.PDCMVTDTR)
End If

cboSuspens_PDCMVTDEV1 = Mid$(xYPDCMVT0.PDCMVTCPT, 7, 3)
cboSuspens_PDCMVTDEV2 = Mid$(xYPDCMVT0.PDCMVTCPT, 11, 3)
If Mid$(xYPDCMVT0.PDCMVTCPT, 1, 1) = "A" Then
    cboSuspens_PDCMVTSENS.ListIndex = 0
Else
    cboSuspens_PDCMVTSENS.ListIndex = 1
End If
If Mid$(xYPDCMVT0.PDCMVTCPT, 7, 3) = "EUR" Then
    txtSuspens_PDCMVTMTD1 = IIf(xYPDCMVT0.PDCMVTMTE = 0, "", Format$(Abs(xYPDCMVT0.PDCMVTMTE), "### ### ### ##0.00"))
    txtSuspens_PDCMVTMTD2 = IIf(xYPDCMVT0.PDCMVTMTD = 0, "", Format$(Abs(xYPDCMVT0.PDCMVTMTD), "### ### ### ##0.00"))
Else
    txtSuspens_PDCMVTMTD2 = IIf(xYPDCMVT0.PDCMVTMTE = 0, "", Format$(Abs(xYPDCMVT0.PDCMVTMTE), "### ### ### ##0.00"))
    txtSuspens_PDCMVTMTD1 = IIf(xYPDCMVT0.PDCMVTMTD = 0, "", Format$(Abs(xYPDCMVT0.PDCMVTMTD), "### ### ### ##0.00"))
End If

fraSuspens_Display_Montant
fraSuspens_Display_PDCMVTTAUX
If xYPDCMVT0.PDCMVTOPEN > 0 And xYPDCMVT0.PDCMVTSTA2 = " " Then
    If xYPDCMVT0.PDCMVTDTR = YBIATAB0_DATE_CPT_J Then cmdSuspens_Annulation.Visible = True
    If xYPDCMVT0.PDCMVTSTA = "+" Then
        cmdSuspens_Update.Visible = True
        cmdSuspens_Update.Caption = "Contrepasser"
        cmdSuspens_Modification.Visible = True: fraSuspens_M.Enabled = True
    End If
End If

'_______________
cboSuspens_PDCMVTSER = xYPDCMVT0.PDCMVTSER & " " & xYPDCMVT0.PDCMVTSSE

Call DTPicker_Set(txtSuspens_PDCMVTDVA, xYPDCMVT0.PDCMVTDVA)
Call fraSuspens_Display_CLI(xYPDCMVT0.PDCMVTCLI)
Call fraSuspens_Display_Log



fraSuspens.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

fraSuspens.Visible = True

End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
Dim xSQL As String
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        If cmdSelect_SQL_K <> "X" Then
            Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
            fgSelect.Col = fgSelect_arrIndex:  arrYPDCPOS0_Index = CLng(fgSelect.Text)
            oldYPDCPOS0 = arrYPDCPOS0(arrYPDCPOS0_Index)
            xSQL = "where PDCMVTDTR = '" & oldYPDCPOS0.PDCPOSDTR & "' and PDCMVTDEV = '" & oldYPDCPOS0.PDCPOSDEV & "'"
    
            Call arrYPDCMVT0_SQL(xSQL)
            
            fgDetail_Display_YPDCMVT0
        End If
   End If
End If
fgSelect.LeftCol = -1
fgSelect.LeftCol = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

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
    Case Is = 13: If Not fraSelect_Comment_Xls.Visible Then KeyCode = 0: cmdContext_Return
    Case Is = 27: KeyCode = 0: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select


End Sub

Public Sub cmdContext_Quit()
'blnControl = False
lstErr.Clear: lstErr.Height = 200
If SSTab1.Tab = 0 Then
    If fraReport.Visible Then
        fraReport.Visible = False
        Exit Sub
    End If
    If fraSelect_Comment_Xls.Visible Then
        fraSelect_Comment_Xls.Visible = False
        Exit Sub
    End If

    If fraSuspens.Visible Then
        fraSuspens.Visible = False
    Else
        If fraPDCOPE.Visible Then
            fraPDCOPE.Visible = False
        Else
            If fgPDCOPE.Visible Then
                fgPDCOPE.Visible = False
            Else
                If fgDetail.Visible Then
                    fgDetail.Visible = False
                Else
                    Unload Me
                End If
            End If
        End If
    End If
Else
       SSTab1.Tab = SSTab1.Tab - 1
End If
End Sub

Public Sub cmdContext_Return()
On Error Resume Next

If fraPDCOPE.Visible Or fraSuspens.Visible Then
Else
    SendKeys "{TAB}"
End If
'___________________________
'If fraPDCOPE.Visible And cmdPDCOPE_Update.Enabled Then
'    If blnPDCOPE_Control_Ok Then
'        cmdPDCOPE_Update_Click
'    Else
'        blnPDCOPE_Control_Ok = True
'    End If
'Else
'    If fraSuspens.Visible And cmdSuspens_Update.Enabled Then
'        If blnSuspens_Control_Ok Then
'            cmdSuspens_Update_Click
'        Else
'            blnSuspens_Control_Ok = True
'        End If
'    Else
'        SendKeys "{TAB}"
'    End If
'End If
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Dim V
Dim xName  As String, xMemo As String
On Error GoTo Error_Handler

mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False


Exit Sub

Error_Handler:

blnControl = False
If Not blnAuto Then MsgBox Error
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
'Call txt_GotFocus(txt)
'Call txt_LostFocus(txt)

'MsgBox "BIA_PDC : msg_rcv"
'YBIATAB0_DATE_CPT_JS1 = "20090304"
'YBIATAB0_DATE_CPT_J = "20090303"

'MsgBox "cmdSelect_SQL_5"
'YPDCLOG0_Init (YBIATAB0_DATE_CPT_J)
'Call cmdSelect_SQL_5OPE(YBIATAB0_DATE_CPT_J, YBIATAB0_DATE_CPT_JS1)
'YPDCLOG0_Write
'Exit Sub

End Sub






Private Sub mnuPDCMVTKCUT_Update_Click()
Dim V, xSQL As String
Dim wYPDCPOS0 As typeYPDCPOS0

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents

wYPDCPOS0 = oldYPDCPOS0
V = cmdYPDCMVT0_Transaction(constModifier)
    
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then

        cmdSelect_SQL_1
        'xSql = "where PDCMVTDTR = '" & wYPDCPOS0.PDCPOSDTR & "' and PDCMVTDEV = '" & wYPDCPOS0.PDCPOSDEV & "'"

        'Call arrYPDCMVT0_SQL(xSql)
        
        'fgDetail_Display_YPDCMVT0

    Else
        MsgBox V, vbCritical, Me.Name & " : mnuPDCMVTKCUT_Update_Click"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_YPDCLOG0_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
'cmdPrint_Ok "LOG"
Me.Enabled = True: Me.MousePointer = 0
Me.Show

End Sub

Private Sub mnuSelect_Print_YPDCMVT0_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdPrint_Historique_YPDCPOS0 "MVT"

Me.Enabled = True: Me.MousePointer = 0
Me.Show

End Sub

Private Sub mnuSelect_Print_YPDCPOS0_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdPrint_Historique_YPDCPOS0 "POS"

Me.Enabled = True: Me.MousePointer = 0
Me.Show

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then cmdSelect_Ok.SetFocus

End Sub














Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
End Select
End Sub


Public Sub cmdPrint_Ok_Etat()
Dim K As Integer, I As Integer
Dim blnOk As Boolean, xSQL As String
Dim mEtat_PDCPOS As String, mEtat_Exclure As String, mEtat_Exclure_PDCMVTKCUT As String
Dim mEtat_Suspens As String

blnControl = False
'_________________________________________________
mEtat_Exclure = ""
If chkPrint_Exclure_HB = "1" Then mEtat_Exclure = "(opérations de change en hors-bilan exclues)"

If chkPrint_Suspens_Out = "1" Then mEtat_Exclure = mEtat_Exclure & " (suspens FOTC exclus)"


If chkPrint_Exclure_PDCMVTKCUT = "1" Then
    mEtat_Exclure_PDCMVTKCUT = " (mouvements comptables non soldés exclus)"
Else
    mEtat_Exclure_PDCMVTKCUT = ""
End If

'If chkSelect_Suspens_Out.Value = "1" Then
'    mEtat_Suspens = "  (suspens FOTC exclus)"
'Else
'    mEtat_Suspens = "  (suspens FOTC inclus)"
'End If
'chkSelect_Suspens_Out.Value = "0"

chkSelect_HB = chkPrint_Exclure_HB
chkSelect_Suspens_Out = chkPrint_Suspens_Out
chkSelect_PDCMVTKCUT = chkPrint_Exclure_PDCMVTKCUT

Call prtBIA_PDC_Open

'___________________________________________________________________________
If chkPrint_Comptant = "1" Then
    chkSelect_Terme = "0"
    cmdSelect_SQL_1
    mEtat_PDCPOS = "Position de change comptant au  " & dateImp(wAMJMin)
    Call prtBIA_PDC_Init(mEtat_PDCPOS, mEtat_Exclure & mEtat_Suspens, mEtat_Exclure_PDCMVTKCUT)
    prtBIA_PDCPOS_Form
    prtBIA_PDCPOS_Line fgSelect
End If
'___________________________________________________________________________
If chkPrint_Comptant_Terme = "1" Then
    chkSelect_Terme = "1"
    cmdSelect_SQL_1

    mEtat_PDCPOS = "Position de change comptant + terme au  " & dateImp(wAMJMin)
    Call prtBIA_PDC_Init(mEtat_PDCPOS, mEtat_Exclure & mEtat_Suspens, mEtat_Exclure_PDCMVTKCUT)
    prtBIA_PDCPOS_Form
    prtBIA_PDCPOS_Line fgSelect
End If

'___________________________________________________________________________
If chkPrint_Terme_Echéancier = "1" Then prtBIA_PDCTER_Line fgTermeEch

If chkPrint_Suspens = "1" Then cmdPrint_Suspens
If chkPrint_PDCMVTKCUT = "1" Then cmdPrint_PDCMVTKCUT
If chkPrint_YPDCMVT0 = "1" Then cmdPrint_YPDCMVT0
If chkPrint_YPDCLOG0 = "1" Then cmdPrint_YPDCLOG0

If chkPrint_ZCHGOPE0 = "1" Then cmdPrint_ZCHGOPE0
If chkPrint_CHGOPECRE = "1" Then cmdPrint_CHGOPECRE
If chkPrint_PDCOPEDTR = "1" Then cmdPrint_PDCOPEDTR

prtBIA_PDC_Close
'_________________________________________________
blnControl = True


End Sub

Private Sub cmdPrint_Historique_YPDCPOS0(lFct As String)
Dim mEtat_PDCPOS As String
 
Call prtBIA_PDC_Open

   mEtat_PDCPOS = "Historique des positions de change comptant au  " & dateImp(wAMJMin)
    Call prtBIA_PDC_Init(mEtat_PDCPOS, "", "")
    prtBIA_PDCPOS_Form
   
    prtBIA_PDCPOS_Line fgSelect
    
    If lFct = "MVT" Then cmdPrint_YPDCMVT0

prtBIA_PDC_Close

End Sub

Public Sub cmdPrint_YPDCMVT0()
Dim K As Integer, I As Integer
Dim xSQL As String, mTrame As String

prtBIA_PDC_Etat_Init "détail des mouvements"
prtBIA_PDCMVT_Form
For I = 1 To arrYPDCPOS0_Nb
    newYPDCPOS0 = arrYPDCPOS0(I)
    xSQL = "where PDCMVTDTR = '" & newYPDCPOS0.PDCPOSDTR & "' and PDCMVTDEV = '" & newYPDCPOS0.PDCPOSDEV & "'"
    Call arrYPDCMVT0_SQL(xSQL)
    mTrame = " "
    If arrYPDCMVT0_Nb > 0 Then
        mTrame = "B"
        oldYPDCPOS0 = newYPDCPOS0
        oldYPDCPOS0.PDCPOSDTR = DateComptablePrecedente(newYPDCPOS0.PDCPOSDTR)
        If oldYPDCPOS0.PDCPOSDTR <> "00000000" Then
            xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCPOS0 " _
                 & " where PDCPOSDTR ='" & oldYPDCPOS0.PDCPOSDTR & "'" _
                 & " and    PDCPOSDEV ='" & oldYPDCPOS0.PDCPOSDEV & "'"
            Set rsSab = cnsab.Execute(xSQL)

            If Not rsSab.EOF Then
                V = rsYPDCPOS0_GetBuffer(rsSab, oldYPDCPOS0)
                Call prtBIA_PDCMVT_POS(oldYPDCPOS0, "0")
            End If
        End If
        For K = 1 To arrYPDCMVT0_Nb
         
            prtBIA_PDCMVT_Line arrYPDCMVT0(K), ""
        Next K
    End If
    Call prtBIA_PDCMVT_POS(newYPDCPOS0, mTrame)

Next I
prtBIA_PDCMVT_End


End Sub
Public Sub cmdPrint_PDCMVTKCUT()
Dim K As Integer, I As Integer
Dim xSQL As String, mTrame As String
Dim wKCUT As String

prtBIA_PDC_Etat_Init "Liste des mouvements comptables non soldés"
prtBIA_PDCMVT_Form

cmdSelect_SQL_1PDCMVTKCUT

For I = 1 To arrYPDCMVT0_Nb

    For K = 1 To arrDev_Nb
        If arrKCUT(K).PDCPOSDEV = xYPDCMVT0.PDCMVTDEV Then
            wKCUT = "cut " & arrKCUT(K).PDCPOSPRIX
            Exit For
        End If
    Next K

    prtBIA_PDCMVT_Line xYPDCMVT0, wKCUT

Next I
prtBIA_PDCMVT_End

End Sub

Public Sub cmdPrint_Suspens()
Dim K As Integer, I As Integer
Dim xSQL As String, mTrame As String

prtBIA_PDC_Etat_Init "Liste des suspens TC"
prtBIA_PDCMVT_Form

xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCMVT0 where PDCMVTOPEC like 'XX%' and PDCMVTSTA2 = ' ' order by PDCMVTDEV"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYPDCMVT0_GetBuffer(rsSab, xYPDCMVT0)
    prtBIA_PDCMVT_Line xYPDCMVT0, ""

    rsSab.MoveNext

Loop
prtBIA_PDCMVT_End

End Sub
Public Sub cmdPrint_YPDCLOG0()
Dim K As Integer, I As Integer
Dim xSQL As String, mTrame As String

prtBIA_PDC_Etat_Init "Historique des traitements de la journée comptable du " & dateImp(wAMJMin)
prtBIA_PDCLOG_Form
'If blnAuto Then
'    xSql = "where PDCLOGDTR > '" & mPDCLOGDTR_Min & "'"
'Else
    xSQL = "where PDCLOGDTR = '" & wAMJMin & "'"
'End If
Call arrYPDCLOG0_SQL(xSQL)

For I = 1 To arrYPDCLOG0_Nb
    prtBIA_PDCLOG_Line arrYPDCLOG0(I)
Next I
prtBIA_PDCLOG_End
End Sub

Public Sub cmdPrint_ZCHGOPE0()
Dim K As Integer, I As Integer
Dim xSQL As String, mTrame As String

prtBIA_PDC_Etat_Init "Liste des opérations SAB saisies et non validées"
prtBIA_PDCOPE_Form

xSQL = "where CHGOPEDE2 <> '   ' and CHGOPEVAL = '1' and CHGOPEANN = ' '"
Call arrZCHGOPE0_SQL(xSQL)

For I = 1 To arrZCHGOPE0_Nb
    prtBIA_PDCOPE_ZCHGOPE0 arrZCHGOPE0(I), "non validée"
Next I
prtBIA_PDCOPE_End

End Sub

Public Sub cmdPrint_CHGOPECRE()
Dim K As Integer, I As Integer
Dim xSQL As String, mTrame As String
Dim X As String
prtBIA_PDC_Etat_Init "Liste des opérations SAB saisies le " & dateImp(wAMJMin)
prtBIA_PDCOPE_Form

cmdSelect_SQL_2_ZCHGOPE0

For I = 1 To arrZCHGOPE0_Nb
    If arrZCHGOPE0(I).CHGOPEVAL <> "O" Then
        X = "non validé"
    Else
        X = ""
        If arrZCHGOPE0(I).CHGOPESER <> "TC" Then
            If arrZCHGOPE0(I).CHGOPEMDA <> "MAD" Then
                X = "non compta"
            End If
        End If
    End If

    prtBIA_PDCOPE_ZCHGOPE0 arrZCHGOPE0(I), X
Next I
prtBIA_PDCOPE_End

End Sub
Public Sub cmdPrint_PDCOPEDTR()
Dim K As Integer, I As Integer
Dim xSQL As String, mTrame As String

prtBIA_PDC_Etat_Init "Liste des opérations BIA_PDC saisies le " & dateImp(wAMJMin)
prtBIA_PDCOPE_Form

cmdSelect_SQL_Where = "where PDCOPEIAMJ = '" & wAMJMin & "' or PDCOPEDTR = '" & wAMJMin & "'"
Call arrYPDCOPE0_SQL(cmdSelect_SQL_Where)

For I = 1 To arrYPDCOPE0_Nb
    prtBIA_PDCOPE_YPDCOPE0 arrYPDCOPE0(I)
Next I
prtBIA_PDCOPE_End

End Sub

Public Sub Error_Route(V)

currentError = CStr(V) & "             ( " & Me.Name & " ~ " & App_Debug & " )"
If blnAuto Then
  '  Call cmdSendMail_Alerte(Me.Name & " ~ " & App_Debug, CStr(V))
Else
    MsgBox V, vbCritical, Me.Name & " ~ " & App_Debug
End If

End Sub








Public Sub cmdSelect_SQL_9M(lPCI As String)
On Error GoTo Error_Handler
Dim xSQL As String, devI As Long, eurI As Long, eurK As Long, eurNb As Long
Dim wMOUVEMCOM As String
Dim eurPDC(5000) As typeYPDCMVT0, eurPDC_Nb As Long
Dim devPDC(5000) As typeYPDCMVT0, devPDC_Nb As Long
Dim wCur As Currency, wPDCMVTMTD_CV As Currency
Dim blnCV As Boolean, xMemo As String

wMOUVEMCOM = lPCI & newYPDCPOS0.PDCPOSDEV & "EUR"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH where MOUVEMCOM = '" & wMOUVEMCOM & "' and MOUVEMDTR > 1081231 order by MOUVEMDTR , MOUVEMPIE , MOUVEMECR"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    rsYPDCMVT0_Init xYPDCMVT0
    xYPDCMVT0.PDCMVTCPT = rsSab("MOUVEMCOM")
    xYPDCMVT0.PDCMVTOPEC = rsSab("MOUVEMOPE")
    xYPDCMVT0.PDCMVTOPEN = rsSab("MOUVEMNUM")
    xYPDCMVT0.PDCMVTDTR = rsSab("MOUVEMDTR")
    xYPDCMVT0.PDCMVTPIE = rsSab("MOUVEMPIE")
    xYPDCMVT0.PDCMVTECR = rsSab("MOUVEMECR")
    xYPDCMVT0.PDCMVTDEV = rsSab("COMPTEDEV")
    xYPDCMVT0.PDCMVTMTD = rsSab("MOUVEMMON")
    xYPDCMVT0.PDCMVTDVA = rsSab("MOUVEMDVA") + 19000000
    xYPDCMVT0.PDCMVTSTA = "*"
    devPDC_Nb = devPDC_Nb + 1
    
    devPDC(devPDC_Nb) = xYPDCMVT0
    
    rsSab.MoveNext

Loop

eurPDC_Nb = 0
wMOUVEMCOM = lPCI & "EUR" & newYPDCPOS0.PDCPOSDEV
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH where MOUVEMCOM = '" & wMOUVEMCOM & "' and MOUVEMDTR > 1081231 order by MOUVEMDTR , MOUVEMPIE , MOUVEMECR"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    rsYPDCMVT0_Init xYPDCMVT0
    xYPDCMVT0.PDCMVTCPT = rsSab("MOUVEMCOM")
    xYPDCMVT0.PDCMVTOPEC = rsSab("MOUVEMOPE")
    xYPDCMVT0.PDCMVTOPEN = rsSab("MOUVEMNUM")
    xYPDCMVT0.PDCMVTDTR = rsSab("MOUVEMDTR")
    xYPDCMVT0.PDCMVTPIE = rsSab("MOUVEMPIE")
    xYPDCMVT0.PDCMVTECR = rsSab("MOUVEMECR")
    xYPDCMVT0.PDCMVTDEV = rsSab("COMPTEDEV")
    xYPDCMVT0.PDCMVTMTE = rsSab("MOUVEMMON")
    xYPDCMVT0.PDCMVTDVA = rsSab("MOUVEMDVA") + 19000000
    xYPDCMVT0.PDCMVTSTA = "*"
    eurPDC_Nb = eurPDC_Nb + 1
    
    eurPDC(eurPDC_Nb) = xYPDCMVT0
    
    rsSab.MoveNext

Loop


For devI = 1 To devPDC_Nb
    xYPDCMVT0 = devPDC(devI)
    eurK = 0: eurNb = 0
    If xYPDCMVT0.PDCMVTOPEC = "PPD" Then
        blnCV = True
        Call sqlYBIATAB0_Read("PDC", newYPDCPOS0.PDCPOSDEV, xYPDCMVT0.PDCMVTDVA, xMemo)
        wPDCMVTMTD_CV = -xYPDCMVT0.PDCMVTMTD / CDbl(Mid$(xMemo, 9, 15) / 1000000000)
    Else
        blnCV = False
        wPDCMVTMTD_CV = 0
    End If

    For eurI = 1 To eurPDC_Nb
        If xYPDCMVT0.PDCMVTDTR = eurPDC(eurI).PDCMVTDTR _
        And xYPDCMVT0.PDCMVTSTA = "*" _
        And xYPDCMVT0.PDCMVTOPEC = eurPDC(eurI).PDCMVTOPEC Then
            Select Case eurPDC(eurI).PDCMVTOPEC
                Case "PPD":
                        If Abs(eurPDC(eurI).PDCMVTMTE - wPDCMVTMTD_CV) < 0.05 Then eurK = eurI: eurNb = 1: Exit For
                Case "*Z1":
                    eurK = eurI: eurNb = eurNb + 1
                Case Else:
                    If xYPDCMVT0.PDCMVTPIE = eurPDC(eurI).PDCMVTPIE _
                    And xYPDCMVT0.PDCMVTOPEN = eurPDC(eurI).PDCMVTOPEN Then eurK = eurI: eurNb = eurNb + 1
            End Select
        End If
    Next eurI
    If eurNb = 1 Then
        devPDC(devI).PDCMVTMTE = eurPDC(eurK).PDCMVTMTE
        devPDC(devI).PDCMVTTAUX = Round(Abs(devPDC(devI).PDCMVTMTD / devPDC(devI).PDCMVTMTE), 6)
        devPDC(devI).PDCMVTSTA = " "
        eurPDC(eurK).PDCMVTSTA = " "
    Else
        'MsgBox eurNb & "  CV : " & xYPDCMVT0.PDCMVTOPEC & " " & xYPDCMVT0.PDCMVTDTR & " " & xYPDCMVT0.PDCMVTPIE & " " & xYPDCMVT0.PDCMVTECR, vbCritical, xYPDCMVT0.PDCMVTCPT
    End If
           
Next devI

For devI = 1 To devPDC_Nb
    arrYPDCMVT0_Nb = arrYPDCMVT0_Nb + 1
    arrYPDCMVT0(arrYPDCMVT0_Nb) = devPDC(devI)
    arrYPDCMVT0(arrYPDCMVT0_Nb).PDCMVTDTR = arrYPDCMVT0(arrYPDCMVT0_Nb).PDCMVTDTR + 19000000
    newYPDCPOS0.PDCPOSPOSD = newYPDCPOS0.PDCPOSPOSD - devPDC(devI).PDCMVTMTD
    If devPDC(devI).PDCMVTSTA <> " " Then
        xYPDCMVT0 = devPDC(devI)
        'MsgBox "?dev: " & xYPDCMVT0.PDCMVTOPEC & " " & xYPDCMVT0.PDCMVTDTR & " " & xYPDCMVT0.PDCMVTPIE & " " & xYPDCMVT0.PDCMVTECR & vbCrLf & xYPDCMVT0.PDCMVTMTD, vbCritical, xYPDCMVT0.PDCMVTCPT
    End If
Next devI
For eurI = 1 To eurPDC_Nb
    newYPDCPOS0.PDCPOSPOSE = newYPDCPOS0.PDCPOSPOSE - eurPDC(eurI).PDCMVTMTE
    If eurPDC(eurI).PDCMVTSTA <> " " Then
        arrYPDCMVT0_Nb = arrYPDCMVT0_Nb + 1
        arrYPDCMVT0(arrYPDCMVT0_Nb) = eurPDC(eurI)
        arrYPDCMVT0(arrYPDCMVT0_Nb).PDCMVTDTR = arrYPDCMVT0(arrYPDCMVT0_Nb).PDCMVTDTR + 19000000
        xYPDCMVT0 = eurPDC(eurI)
        If xYPDCMVT0.PDCMVTOPEC <> "RPC" Then
            'MsgBox "?eur: " & xYPDCMVT0.PDCMVTOPEC & " " & xYPDCMVT0.PDCMVTDTR & " " & xYPDCMVT0.PDCMVTPIE & " " & xYPDCMVT0.PDCMVTECR & vbCrLf & xYPDCMVT0.PDCMVTMTE, vbCritical, xYPDCMVT0.PDCMVTCPT
        End If
    End If
Next eurI

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdSelect_SQL_9M"

End Sub

Public Sub cmdSelect_SQL_5M(lPCI As String, lMOUVEMDTR As Long)
On Error GoTo Error_Handler
Dim xSQL As String, devI As Long, eurI As Long, eurK As Long, eurNb As Long
Dim wMOUVEMCOM As String
Dim eurPDC(1000) As typeYPDCMVT0, eurPDC_Nb As Long
Dim devPDC(1000) As typeYPDCMVT0, devPDC_Nb As Long
Dim wCur As Currency, wPDCMVTMTD_CV As Currency
Dim blnCV As Boolean, xMemo As String
Dim wK2 As String
Dim wPDCMVTOPEN As Long
On Error GoTo Error_Handler

wMOUVEMCOM = lPCI & newYPDCPOS0.PDCPOSDEV & "EUR"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH where MOUVEMCOM = '" & wMOUVEMCOM & "' and MOUVEMDTR = " & lMOUVEMDTR & " and mouvemanu <> 3 order by MOUVEMDTR , MOUVEMPIE , MOUVEMECR"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    rsYPDCMVT0_Init xYPDCMVT0
    xYPDCMVT0.PDCMVTCPT = rsSab("MOUVEMCOM")
    xYPDCMVT0.PDCMVTOPEC = rsSab("MOUVEMOPE")
    xYPDCMVT0.PDCMVTOPEN = rsSab("MOUVEMNUM")
    xYPDCMVT0.PDCMVTDTR = rsSab("MOUVEMDTR")
    xYPDCMVT0.PDCMVTPIE = rsSab("MOUVEMPIE")
    xYPDCMVT0.PDCMVTECR = rsSab("MOUVEMECR")
    xYPDCMVT0.PDCMVTDEV = rsSab("COMPTEDEV")
    xYPDCMVT0.PDCMVTMTD = rsSab("MOUVEMMON")
    xYPDCMVT0.PDCMVTDVA = rsSab("MOUVEMDVA") + 19000000
    xYPDCMVT0.PDCMVTSTA = "*"
    xYPDCMVT0.PDCMVTSTA2 = " "
    xYPDCMVT0.PDCMVTSER = rsSab("MOUVEMSER")
    xYPDCMVT0.PDCMVTSSE = rsSab("MOUVEMSSE")
    If xYPDCMVT0.PDCMVTOPEC = "-TR" Then
        X = rsSab("LIBELLIB2")
        wPDCMVTOPEN = cmdSelect_SQL_5M_Trilog(X)
        If wPDCMVTOPEN > 0 Then xYPDCMVT0.PDCMVTOPEN = wPDCMVTOPEN
    End If
'    If xYPDCMVT0.PDCMVTOPEC <> "SWP" Then
        devPDC_Nb = devPDC_Nb + 1
        devPDC(devPDC_Nb) = xYPDCMVT0
 '   End If
    rsSab.MoveNext

Loop

eurPDC_Nb = 0
wMOUVEMCOM = lPCI & "EUR" & newYPDCPOS0.PDCPOSDEV
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH where MOUVEMCOM = '" & wMOUVEMCOM & "' and MOUVEMDTR = " & lMOUVEMDTR & " order by MOUVEMDTR , MOUVEMPIE , MOUVEMECR"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    rsYPDCMVT0_Init xYPDCMVT0
    xYPDCMVT0.PDCMVTCPT = rsSab("MOUVEMCOM")
    xYPDCMVT0.PDCMVTOPEC = rsSab("MOUVEMOPE")
    xYPDCMVT0.PDCMVTOPEN = rsSab("MOUVEMNUM")
    xYPDCMVT0.PDCMVTDTR = rsSab("MOUVEMDTR")
    xYPDCMVT0.PDCMVTPIE = rsSab("MOUVEMPIE")
    xYPDCMVT0.PDCMVTECR = rsSab("MOUVEMECR")
    xYPDCMVT0.PDCMVTDEV = newYPDCPOS0.PDCPOSDEV
    xYPDCMVT0.PDCMVTMTE = rsSab("MOUVEMMON")
    xYPDCMVT0.PDCMVTDVA = rsSab("MOUVEMDVA") + 19000000
    xYPDCMVT0.PDCMVTSTA = "*"
    xYPDCMVT0.PDCMVTSTA2 = " "
    xYPDCMVT0.PDCMVTSER = rsSab("MOUVEMSER")
    xYPDCMVT0.PDCMVTSSE = rsSab("MOUVEMSSE")
    If xYPDCMVT0.PDCMVTOPEC = "-TR" Then
        X = rsSab("LIBELLIB2")
        wPDCMVTOPEN = cmdSelect_SQL_5M_Trilog(X)
        If wPDCMVTOPEN > 0 Then xYPDCMVT0.PDCMVTOPEN = wPDCMVTOPEN
    End If
'    If xYPDCMVT0.PDCMVTOPEC <> "SWP" Then
        eurPDC_Nb = eurPDC_Nb + 1
        eurPDC(eurPDC_Nb) = xYPDCMVT0
'    End If
    rsSab.MoveNext

Loop


For devI = 1 To devPDC_Nb
    xYPDCMVT0 = devPDC(devI)
    eurK = 0: eurNb = 0
    If xYPDCMVT0.PDCMVTOPEC = "PPD" Then
        blnCV = True
        Call sqlYBIATAB0_Read("PDC", newYPDCPOS0.PDCPOSDEV, xYPDCMVT0.PDCMVTDVA, xMemo)
        If Not IsNumeric(Mid$(xMemo, 9, 15)) Then
            xYPDCLOG0.PDCLOGPIE = xYPDCMVT0.PDCMVTPIE
            xYPDCLOG0.PDCLOGECR = xYPDCMVT0.PDCMVTECR
            xYPDCLOG0.PDCLOGNAT = "5M "
            xYPDCLOG0.PDCLOGTXT = "Fixing non trouvé » " & xYPDCMVT0.PDCMVTDEV & " " & xYPDCMVT0.PDCMVTDVA
            Call YPDCLOG0_AddItem

            'V = "Fixing non trouvé : " & xYPDCMVT0.PDCMVTDEV & " " & xYPDCMVT0.PDCMVTDVA
            'If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdSelect_SQL_9M"
            xYPDCMVT0.PDCMVTTAUX = newYPDCPOS0.PDCPOSFIXT
        Else
            xYPDCMVT0.PDCMVTTAUX = CDbl(Mid$(xMemo, 9, 15) / 1000000000)
        End If
        
        wPDCMVTMTD_CV = -xYPDCMVT0.PDCMVTMTD / xYPDCMVT0.PDCMVTTAUX
    Else
        blnCV = False
        wPDCMVTMTD_CV = 0
    End If

    For eurI = 1 To eurPDC_Nb
        If xYPDCMVT0.PDCMVTDTR = eurPDC(eurI).PDCMVTDTR _
        And eurPDC(eurI).PDCMVTSTA = "*" _
        And xYPDCMVT0.PDCMVTOPEC = eurPDC(eurI).PDCMVTOPEC Then
            Select Case eurPDC(eurI).PDCMVTOPEC
                Case "PPD":
                        If Abs(eurPDC(eurI).PDCMVTMTE - wPDCMVTMTD_CV) < 0.05 Then eurK = eurI: eurNb = 1: Exit For
                Case "*Z1":
                    eurK = eurI: eurNb = eurNb + 1
                Case Else:
                    If xYPDCMVT0.PDCMVTPIE = eurPDC(eurI).PDCMVTPIE _
                    And xYPDCMVT0.PDCMVTOPEN = eurPDC(eurI).PDCMVTOPEN Then eurK = eurI: eurNb = eurNb + 1
            End Select
        End If
    Next eurI
    If eurNb = 1 Then
        devPDC(devI).PDCMVTMTE = eurPDC(eurK).PDCMVTMTE
        devPDC(devI).PDCMVTTAUX = Round(Abs(devPDC(devI).PDCMVTMTD / devPDC(devI).PDCMVTMTE), 6)
        devPDC(devI).PDCMVTSTA = " "
        eurPDC(eurK).PDCMVTSTA = " "
    Else
            xYPDCLOG0.PDCLOGPIE = xYPDCMVT0.PDCMVTPIE
            xYPDCLOG0.PDCLOGECR = xYPDCMVT0.PDCMVTECR
            xYPDCLOG0.PDCLOGNAT = "5M?"
            xYPDCLOG0.PDCLOGTXT = "CV EUR non trouvée : " & Trim(Format$(xYPDCMVT0.PDCMVTMTD, "### ### ### ##0.00")) & " " & Mid$(xYPDCMVT0.PDCMVTCPT, 7, 3) & " ( " & xYPDCMVT0.PDCMVTCPT & " )"
            Call YPDCLOG0_AddItem
      ' MsgBox eurNb & "  CV » " & xYPDCMVT0.PDCMVTOPEC & " " & xYPDCMVT0.PDCMVTDTR & " " & xYPDCMVT0.PDCMVTPIE & " " & xYPDCMVT0.PDCMVTECR, vbCritical, xYPDCMVT0.PDCMVTCPT
    End If
          
Next devI

For devI = 1 To devPDC_Nb
    arrYPDCMVT0_Nb = arrYPDCMVT0_Nb + 1
    arrYPDCMVT0(arrYPDCMVT0_Nb) = devPDC(devI)
    arrYPDCMVT0(arrYPDCMVT0_Nb).PDCMVTDTR = arrYPDCMVT0(arrYPDCMVT0_Nb).PDCMVTDTR + 19000000
    If devPDC(devI).PDCMVTOPEC <> "SWP" Then
        newYPDCPOS0.PDCPOSPOSD = newYPDCPOS0.PDCPOSPOSD + devPDC(devI).PDCMVTMTD
    End If
    If devPDC(devI).PDCMVTSTA <> " " Then
        xYPDCMVT0 = devPDC(devI)
        xYPDCLOG0.PDCLOGPIE = xYPDCMVT0.PDCMVTPIE
        xYPDCLOG0.PDCLOGECR = xYPDCMVT0.PDCMVTECR
        xYPDCLOG0.PDCLOGNAT = "5M?"
        xYPDCLOG0.PDCLOGTXT = "Mvt non rapproché : " & Trim(Format$(xYPDCMVT0.PDCMVTMTD, "### ### ### ##0.00")) & " " & Mid$(xYPDCMVT0.PDCMVTCPT, 7, 3) & " ( " & xYPDCMVT0.PDCMVTCPT & " )"
        Call YPDCLOG0_AddItem

        'MsgBox "?dev: " & xYPDCMVT0.PDCMVTOPEC & " " & xYPDCMVT0.PDCMVTDTR & " " & xYPDCMVT0.PDCMVTPIE & " " & xYPDCMVT0.PDCMVTECR & vbCrLf & xYPDCMVT0.PDCMVTMTD, vbCritical, xYPDCMVT0.PDCMVTCPT
    End If
Next devI
For eurI = 1 To eurPDC_Nb
    If eurPDC(eurI).PDCMVTOPEC = "RPC" Then
        newYPDCPOS0.PDCPOSRPC = newYPDCPOS0.PDCPOSRPC + eurPDC(eurI).PDCMVTMTE
    Else
        If eurPDC(eurI).PDCMVTOPEC <> "SWP" Then
            newYPDCPOS0.PDCPOSPOSE = newYPDCPOS0.PDCPOSPOSE + eurPDC(eurI).PDCMVTMTE
        End If
    End If
    If eurPDC(eurI).PDCMVTSTA <> " " Then
        arrYPDCMVT0_Nb = arrYPDCMVT0_Nb + 1
        arrYPDCMVT0(arrYPDCMVT0_Nb) = eurPDC(eurI)
        arrYPDCMVT0(arrYPDCMVT0_Nb).PDCMVTDTR = arrYPDCMVT0(arrYPDCMVT0_Nb).PDCMVTDTR + 19000000
        If eurPDC(eurI).PDCMVTOPEC <> "RPC" Then
            xYPDCMVT0 = eurPDC(eurI)
            xYPDCLOG0.PDCLOGPIE = xYPDCMVT0.PDCMVTPIE
            xYPDCLOG0.PDCLOGECR = xYPDCMVT0.PDCMVTECR
            xYPDCLOG0.PDCLOGNAT = "5M?"
            xYPDCLOG0.PDCLOGTXT = "Mvt non rapproché : " & Trim(Format$(xYPDCMVT0.PDCMVTMTE, "### ### ### ##0.00")) & " " & Mid$(xYPDCMVT0.PDCMVTCPT, 7, 3) & " ( " & xYPDCMVT0.PDCMVTCPT & " )"
            Call YPDCLOG0_AddItem
            'MsgBox "?eur: " & xYPDCMVT0.PDCMVTOPEC & " " & xYPDCMVT0.PDCMVTDTR & " " & xYPDCMVT0.PDCMVTPIE & " " & xYPDCMVT0.PDCMVTECR & vbCrLf & xYPDCMVT0.PDCMVTMTE, vbCritical, xYPDCMVT0.PDCMVTCPT
        End If
    End If
Next eurI

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " » cmdSelect_SQL_5M"
    xYPDCLOG0.PDCLOGPIE = 0
    xYPDCLOG0.PDCLOGECR = 0
    xYPDCLOG0.PDCLOGNAT = "5M?"
    xYPDCLOG0.PDCLOGTXT = V
    Call YPDCLOG0_AddItem

End Sub
Public Sub cmdSelect_SQL_5M_Terme(lPCI As String, lMOUVEMDTR As Long)
On Error GoTo Error_Handler
Dim xSQL As String, devI As Long, eurI As Long, eurK As Long, eurNb As Long
Dim wMOUVEMCOM As String
Dim eurPDC(1000) As typeYPDCMVT0, eurPDC_Nb As Long
Dim devPDC(1000) As typeYPDCMVT0, devPDC_Nb As Long
Dim wCur As Currency, wPDCMVTMTD_CV_Min As Currency, wPDCMVTMTD_CV_Max As Currency
Dim blnCV As Boolean, xMemo As String
Dim wK2 As String
On Error GoTo Error_Handler

'obsolète
'$$$$$$$$$
Exit Sub


wMOUVEMCOM = lPCI & newYPDCPOS0.PDCPOSDEV & "EUR"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH where MOUVEMCOM = '" & wMOUVEMCOM & "' and MOUVEMDTR = " & lMOUVEMDTR & " and mouvemanu <> 3 order by MOUVEMDTR , MOUVEMPIE , MOUVEMECR"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    rsYPDCMVT0_Init xYPDCMVT0
    xYPDCMVT0.PDCMVTCPT = rsSab("MOUVEMCOM")
    xYPDCMVT0.PDCMVTOPEC = rsSab("MOUVEMOPE")
    xYPDCMVT0.PDCMVTOPEN = rsSab("MOUVEMNUM")
    xYPDCMVT0.PDCMVTDTR = rsSab("MOUVEMDTR")
    xYPDCMVT0.PDCMVTPIE = rsSab("MOUVEMPIE")
    xYPDCMVT0.PDCMVTECR = rsSab("MOUVEMECR")
    xYPDCMVT0.PDCMVTDEV = rsSab("COMPTEDEV")
    xYPDCMVT0.PDCMVTMTD = rsSab("MOUVEMMON")
    xYPDCMVT0.PDCMVTDVA = rsSab("MOUVEMDVA") + 19000000
    xYPDCMVT0.PDCMVTSTA = "*"
    xYPDCMVT0.PDCMVTSTA2 = " "
    xYPDCMVT0.PDCMVTSER = rsSab("MOUVEMSER")
    xYPDCMVT0.PDCMVTSSE = rsSab("MOUVEMSSE")
    devPDC_Nb = devPDC_Nb + 1
    devPDC(devPDC_Nb) = xYPDCMVT0
    rsSab.MoveNext

Loop

eurPDC_Nb = 0
wMOUVEMCOM = lPCI & "EUR" & newYPDCPOS0.PDCPOSDEV
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH where MOUVEMCOM = '" & wMOUVEMCOM & "' and MOUVEMDTR = " & lMOUVEMDTR & " order by MOUVEMDTR , MOUVEMPIE , MOUVEMECR"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    rsYPDCMVT0_Init xYPDCMVT0
    xYPDCMVT0.PDCMVTCPT = rsSab("MOUVEMCOM")
    xYPDCMVT0.PDCMVTOPEC = rsSab("MOUVEMOPE")
    xYPDCMVT0.PDCMVTOPEN = rsSab("MOUVEMNUM")
    xYPDCMVT0.PDCMVTDTR = rsSab("MOUVEMDTR")
    xYPDCMVT0.PDCMVTPIE = rsSab("MOUVEMPIE")
    xYPDCMVT0.PDCMVTECR = rsSab("MOUVEMECR")
    xYPDCMVT0.PDCMVTDEV = newYPDCPOS0.PDCPOSDEV
    xYPDCMVT0.PDCMVTMTE = rsSab("MOUVEMMON")
    xYPDCMVT0.PDCMVTDVA = rsSab("MOUVEMDVA") + 19000000
    xYPDCMVT0.PDCMVTSTA = "*"
    xYPDCMVT0.PDCMVTSTA2 = " "
    xYPDCMVT0.PDCMVTSER = rsSab("MOUVEMSER")
    xYPDCMVT0.PDCMVTSSE = rsSab("MOUVEMSSE")
    eurPDC_Nb = eurPDC_Nb + 1
    eurPDC(eurPDC_Nb) = xYPDCMVT0
    rsSab.MoveNext

Loop


For devI = 1 To devPDC_Nb
    xYPDCMVT0 = devPDC(devI)
    eurK = 0: eurNb = 0
    If newYPDCPOS0.PDCPOSFIXT <> 0 Then
        wCur = Abs(xYPDCMVT0.PDCMVTMTD) / newYPDCPOS0.PDCPOSFIXT
        wPDCMVTMTD_CV_Min = wCur - wCur * 0.1
        wPDCMVTMTD_CV_Max = wCur + wCur * 0.1
    Else
        wPDCMVTMTD_CV_Min = 0
        wPDCMVTMTD_CV_Max = 0
    End If

    For eurI = 1 To eurPDC_Nb
        If xYPDCMVT0.PDCMVTDTR = eurPDC(eurI).PDCMVTDTR _
        And eurPDC(eurI).PDCMVTSTA = "*" _
        And xYPDCMVT0.PDCMVTOPEC = eurPDC(eurI).PDCMVTOPEC Then
            Select Case eurPDC(eurI).PDCMVTOPEC
                Case "*Z1":
                    eurK = eurI: eurNb = eurNb + 1
                Case Else:
                    If xYPDCMVT0.PDCMVTPIE = eurPDC(eurI).PDCMVTPIE _
                    And xYPDCMVT0.PDCMVTOPEN = eurPDC(eurI).PDCMVTOPEN Then
                        If Abs(eurPDC(eurI).PDCMVTMTE) >= wPDCMVTMTD_CV_Min _
                        And Abs(eurPDC(eurI).PDCMVTMTE) <= wPDCMVTMTD_CV_Max _
                            Then eurK = eurI: eurNb = 1: Exit For
                    End If
            End Select
        End If
    Next eurI
    If eurNb = 1 Then
        devPDC(devI).PDCMVTMTE = eurPDC(eurK).PDCMVTMTE
        devPDC(devI).PDCMVTTAUX = Round(Abs(devPDC(devI).PDCMVTMTD / devPDC(devI).PDCMVTMTE), 6)
        devPDC(devI).PDCMVTSTA = " "
        eurPDC(eurK).PDCMVTSTA = " "
    Else
            xYPDCLOG0.PDCLOGPIE = xYPDCMVT0.PDCMVTPIE
            xYPDCLOG0.PDCLOGECR = xYPDCMVT0.PDCMVTECR
            xYPDCLOG0.PDCLOGNAT = "5M?"
            xYPDCLOG0.PDCLOGTXT = "CV EUR non trouvée : " & Trim(Format$(xYPDCMVT0.PDCMVTMTD, "### ### ### ##0.00")) & " " & Mid$(xYPDCMVT0.PDCMVTCPT, 7, 3) & " ( " & xYPDCMVT0.PDCMVTCPT & " )"
            Call YPDCLOG0_AddItem
      ' MsgBox eurNb & "  CV » " & xYPDCMVT0.PDCMVTOPEC & " " & xYPDCMVT0.PDCMVTDTR & " " & xYPDCMVT0.PDCMVTPIE & " " & xYPDCMVT0.PDCMVTECR, vbCritical, xYPDCMVT0.PDCMVTCPT
    End If
           
Next devI

For devI = 1 To devPDC_Nb
    arrYPDCMVT0_Nb = arrYPDCMVT0_Nb + 1
    arrYPDCMVT0(arrYPDCMVT0_Nb) = devPDC(devI)
    arrYPDCMVT0(arrYPDCMVT0_Nb).PDCMVTDTR = arrYPDCMVT0(arrYPDCMVT0_Nb).PDCMVTDTR + 19000000
    If eurPDC(eurI).PDCMVTOPEC = "TER" Then
        newYPDCPOS0.PDCPOSTERD = newYPDCPOS0.PDCPOSTERD + eurPDC(eurI).PDCMVTMTD
    Else
        newYPDCPOS0.PDCPOSSWPD = newYPDCPOS0.PDCPOSSWPD + eurPDC(eurI).PDCMVTMTD
    End If
    If devPDC(devI).PDCMVTSTA <> " " Then
        xYPDCMVT0 = devPDC(devI)
        xYPDCLOG0.PDCLOGPIE = xYPDCMVT0.PDCMVTPIE
        xYPDCLOG0.PDCLOGECR = xYPDCMVT0.PDCMVTECR
        xYPDCLOG0.PDCLOGNAT = "5M?"
        xYPDCLOG0.PDCLOGTXT = "Mvt non rapproché : " & Trim(Format$(xYPDCMVT0.PDCMVTMTD, "### ### ### ##0.00")) & " " & Mid$(xYPDCMVT0.PDCMVTCPT, 7, 3) & " ( " & xYPDCMVT0.PDCMVTCPT & " )"
        Call YPDCLOG0_AddItem

    End If
Next devI
For eurI = 1 To eurPDC_Nb
    If eurPDC(eurI).PDCMVTOPEC = "TER" Then
        newYPDCPOS0.PDCPOSTERE = newYPDCPOS0.PDCPOSTERE + eurPDC(eurI).PDCMVTMTE
    Else
        newYPDCPOS0.PDCPOSSWPE = newYPDCPOS0.PDCPOSSWPE + eurPDC(eurI).PDCMVTMTE
    End If
    If eurPDC(eurI).PDCMVTSTA <> " " Then
        arrYPDCMVT0_Nb = arrYPDCMVT0_Nb + 1
        arrYPDCMVT0(arrYPDCMVT0_Nb) = eurPDC(eurI)
        arrYPDCMVT0(arrYPDCMVT0_Nb).PDCMVTDTR = arrYPDCMVT0(arrYPDCMVT0_Nb).PDCMVTDTR + 19000000
            xYPDCMVT0 = eurPDC(eurI)
            xYPDCLOG0.PDCLOGPIE = xYPDCMVT0.PDCMVTPIE
            xYPDCLOG0.PDCLOGECR = xYPDCMVT0.PDCMVTECR
            xYPDCLOG0.PDCLOGNAT = "5M?"
            xYPDCLOG0.PDCLOGTXT = "Mvt non rapproché : " & Trim(Format$(xYPDCMVT0.PDCMVTMTE, "### ### ### ##0.00")) & " " & Mid$(xYPDCMVT0.PDCMVTCPT, 7, 3) & " ( " & xYPDCMVT0.PDCMVTCPT & " )"
            Call YPDCLOG0_AddItem
    End If
Next eurI

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " » cmdSelect_SQL_5M_Terme"
    xYPDCLOG0.PDCLOGPIE = 0
    xYPDCLOG0.PDCLOGECR = 0
    xYPDCLOG0.PDCLOGNAT = "5M?"
    xYPDCLOG0.PDCLOGTXT = V
    Call YPDCLOG0_AddItem

End Sub

Private Sub txtPDCOPECLI_GotFocus()
Call txt_GotFocus(txtPDCOPECLI)

End Sub

Private Sub txtPDCOPECLI_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtPDCOPECLI_LostFocus()
Call ZCLIEAN0_SQL(Format(Trim(txtPDCOPECLI), "0000000"))
libPDCOPECLI = mCLIENARA1
Call txt_LostFocus(txtPDCOPECLI)

End Sub


Private Sub txtPDCOPEDVA_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
lblPDCOPEDVA.BackColor = fraPDCOPE_S.BackColor '&HC0FFFF
End Sub

Private Sub txtPDCOPEFIXING_Change()
If txtPDCOPEFIXING.Enabled Then fraPDCOPE_Control_DEV_Certain_SWP
End Sub

Private Sub txtPDCOPEFIXING_GotFocus()
Call txt_GotFocus(txtPDCOPEFIXING)

End Sub

Private Sub txtPDCOPEFIXING_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtPDCOPEFIXING)

End Sub


Private Sub txtPDCOPEFIXING_LostFocus()
Call txt_LostFocus(txtPDCOPEFIXING)

End Sub

Private Sub txtPDCOPEITXT_GotFocus()
Call txt_GotFocus(txtPDCOPEITXT)


End Sub


Private Sub txtPDCOPEITXT_LostFocus()
Call txt_LostFocus(txtPDCOPEITXT)

End Sub


Private Sub txtPDCOPEMTD1_Change()
Dim xCur As Currency
xCur = CCur(Val(txtPDCOPEMTD1))
If xCur <> mPDCOPEMTD1 Then
    mPDCOPEMTD1 = xCur
    fraPDCOPE_Control_DEV_Certain

End If

End Sub

Private Sub txtPDCOPEMTD1_GotFocus()
txtPDCOPEMTD1.BackColor = focusUsr.BackColor

End Sub

Private Sub txtPDCOPEMTD1_KeyPress(KeyAscii As Integer)
    Call num_Montant(KeyAscii, txtPDCOPEMTD1)

End Sub

Private Sub txtPDCOPEMTD1_LostFocus()
txtPDCOPEMTD1.BackColor = txtUsr.BackColor

End Sub

Private Sub txtPDCOPEOPEN_GotFocus()
Call txt_GotFocus(txtPDCOPEOPEN)

End Sub

Private Sub txtPDCOPEOPEN_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub

Private Sub txtPDCOPEOPEN_LostFocus()
Call txt_LostFocus(txtPDCOPEOPEN)

End Sub

Private Sub txtPDCOPEREF_GotFocus()
Call txt_GotFocus(txtPDCOPEREF)

End Sub

Private Sub txtPDCOPEREF_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtPDCOPEREF_LostFocus()
Call txt_LostFocus(txtPDCOPEREF)

End Sub

Private Sub txtPDCOPETAUX_Change()
fraPDCOPE_Control_DEV_Certain
End Sub

Private Sub txtPDCOPETAUX_GotFocus()
Call txt_GotFocus(txtPDCOPETAUX)

End Sub

Private Sub txtPDCOPETAUX_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtPDCOPETAUX)
End Sub


Private Sub txtPDCOPETAUX_LostFocus()
Call txt_LostFocus(txtPDCOPETAUX)

End Sub

Private Sub txtPDCOPEVTXT_GotFocus()
Call txt_GotFocus(txtPDCOPEVTXT)

End Sub


Private Sub txtPDCOPEVTXT_LostFocus()
Call txt_LostFocus(txtPDCOPEVTXT)

End Sub


Private Sub txtSelect_AMJ_Change()
cmdSelect_Reset
'cmdSelect_Ok_Click

End Sub


Public Sub cmdSelect_SQL_5CPT()
Dim blnOk As Boolean, blnIdem As Boolean
Dim xSQL As String, wCHGOPEDT1 As String
Dim wDbl As Double
Dim wCHGOPEDE1 As String, wCHGOPEDE2 As String
Dim X As String
On Error GoTo Error_Handler

newYPDCMVT0.PDCMVTSTA2 = "?"

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCHGOPE0 " _
     & " where CHGOPEOPE ='" & newYPDCMVT0.PDCMVTOPEC & "'" _
     & " and    CHGOPEDOS =" & newYPDCMVT0.PDCMVTOPEN _
     & " and    CHGOPESER ='" & newYPDCMVT0.PDCMVTSER & "'" _
     & " and    CHGOPESSE ='" & newYPDCMVT0.PDCMVTSSE & "'"
Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then Exit Sub
blnOk = True: blnIdem = True
newYPDCMVT0.PDCMVTCLI = rsSab("CHGOPECON")
wCHGOPEDE1 = rsSab("CHGOPEDE1")
wCHGOPEDE2 = rsSab("CHGOPEDE2")
If newYPDCMVT0.PDCMVTDEV = wCHGOPEDE1 Then
    If Abs(newYPDCMVT0.PDCMVTMTD) <> CCur(rsSab("CHGOPEMO1")) Then blnOk = False
    If Abs(newYPDCMVT0.PDCMVTMTE) <> CCur(rsSab("CHGOPEMO2")) Then blnOk = False
Else
    If newYPDCMVT0.PDCMVTDEV = wCHGOPEDE2 Then
        If Abs(newYPDCMVT0.PDCMVTMTE) <> CCur(rsSab("CHGOPEMO1")) Then blnOk = False
        If Abs(newYPDCMVT0.PDCMVTMTD) <> CCur(rsSab("CHGOPEMO2")) Then blnOk = False
    Else
            blnOk = False
    End If
End If
'wPDCMVTDVA = newYPDCMVT0.PDCMVTDVA - 19000000
wCHGOPEDT1 = CLng(rsSab("CHGOPEDT1")) + 19000000
If Mid$(newYPDCMVT0.PDCMVTCPT, 1, 1) = "9" Then
    'If wPDCMVTDVA <> CLng(rsSab("CHGOPEENG")) And wPDCMVTDVA <> CLng(rsSab("CHGOPEDT1")) Then blnIdem = False
    newYPDCMVT0.PDCMVTDVA = wCHGOPEDT1
Else
    If newYPDCMVT0.PDCMVTDVA <> wCHGOPEDT1 Then blnIdem = False
End If
wDbl = CDbl(rsSab("CHGOPECO3"))
If wDbl = 0 Then wDbl = CDbl(rsSab("CHGOPECO1"))
If Abs(wDbl - newYPDCMVT0.PDCMVTTAUX) > 0.001 Then
    blnOk = False
End If
If blnOk = False Then
    If wCHGOPEDE1 = "EUR" Or wCHGOPEDE2 = "EUR" Then
        xYPDCLOG0.PDCLOGNAT = "5X#"
        newYPDCMVT0.PDCMVTSTA2 = "#"
        X = " compta/opé ± écart taux et montants"
    Else
        xYPDCLOG0.PDCLOGNAT = "5Xx"
        newYPDCMVT0.PDCMVTSTA2 = "x"
        X = " cours croisés"
   End If
    xYPDCLOG0.PDCLOGPIE = newYPDCMVT0.PDCMVTPIE
    xYPDCLOG0.PDCLOGECR = newYPDCMVT0.PDCMVTECR
    xYPDCLOG0.PDCLOGTXT = newYPDCMVT0.PDCMVTOPEC & " " & newYPDCMVT0.PDCMVTOPEN _
                        & " " & wCHGOPEDE1 & " / " & wCHGOPEDE2 _
                        & X
                        
    If newYPDCMVT0.PDCMVTSTA2 = "#" Then Call YPDCLOG0_AddItem

Else
    newYPDCMVT0.PDCMVTTAUX = wDbl
    If blnIdem = False Then
        newYPDCMVT0.PDCMVTSTA2 = "~"
    Else
        newYPDCMVT0.PDCMVTSTA2 = " "
    End If
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdSelect_SQL_9M"

End Sub

Public Sub cmdSelect_Reset()
On Error Resume Next

If blnControl Then
    If Not blnAuto And Me.Enabled Then cmdContext.SetFocus
    fraPrint.Visible = False
    lstErr.Clear
    fgSelect.Visible = False
    fgDetail.Visible = False
    fgTerme.Visible = False
    fgTermeEch.Visible = False
    fraPDCOPE.Visible = False
    fraSuspens.Visible = False
    fraSuspens_Options.Visible = False
    fraReport.Visible = False

    fgDetail.Visible = False: fgDetail.Top = mfgDetail_Top: fgDetail.Height = mfgDetail_Height
    fgPDCOPE.Visible = True
    cmdPDCOPE_New.Visible = True
    cmdPDCOPE_CONF_CALL.Visible = True
    cmdSelect_SQL_K = Mid$(cboSelect_SQL, 1, 1)
    
    fraSelect_Options.Enabled = True
    fraSelect_Options.Visible = True
    fraSelect_Options_xls.Visible = False
    fraSelect_Comment_Xls.Visible = False
    fraSelect_Options_Y.Visible = False
    libSelect_Report = ""
    libSelect_Report.Visible = False
    
    Select Case cmdSelect_SQL_K
        Case 1: Me.Enabled = True: fraSelect_Options.Enabled = True: cmdSelect_Ok_Click
        Case 3: cmdSelect_SQL_3
        Case 4: fgDetail.Top = fgSelect.Top:: fgDetail.Height = 6660
                fraSelect_Options.Visible = False: fraSuspens_Options.Visible = True
                cmdSelect_SQL_4
        Case 7: Call MsgBox("Attention : gestion manuelle des suspens à faire en fonction de la date de reprise" _
                            & vbCrLf & " - PDCMVTSTA2 à blanc si mvt d'annulation généré >= à cette date" & vbCrLf _
                            & " - resaisir  les suspens créés >= à cette date" & vbCrLf _
                            & " - ATTENTION aux opérations reportées et au calcul PDCPOSPNL" _
                            , vbInformation, "BIA_PDC : Reprise des traitements")
        Case "X": fraSelect_Options.Visible = False: cmdSelect_SQL_xls_Init
        Case "Y": fraSelect_Options.Visible = False: fraSelect_Options_Y.Visible = True
    End Select
End If
End Sub

Public Sub cmdSelect_SQL_2Reset()
If blnControl Then
    lstErr.Clear
    fgPDCOPE.Visible = False
    fraPDCOPE.Visible = False
   
    fraPDCOPE_Options.Enabled = True
    Me.Enabled = True: cmdSelect_SQL_2X
End If

End Sub

Public Sub cmdSelect_SQL_1HB()
Dim wDev As String, xSQL As String
Dim I As Integer, xWhere As String, X As String

ReDim arrYPDCMVT0(101)

X = ""
xWhere = " where PDCMVTDTR = '?????'"
If chkSelect_HB = "1" Then
    If chkSelect_Suspens_Out = "0" Then X = " and PDCMVTOPEC not like 'XX%' "
    
    xWhere = " where PDCMVTDTR <= '" & wAMJMin & "' and PDCMVTDVA > '" & wAmjMin_HB _
    & "'  and PDCMVTSTA2 in (' ', 'x','#','~')" & X

Else
    If chkSelect_Suspens_Out = "1" Then xWhere = " where PDCMVTOPEC like 'XX%'"
End If

arrYPDCMVT0_Max = 100: arrYPDCMVT0_Nb = 0
xSQL = "select count(*) as Tally  from " & paramIBM_Library_SABSPE_XXX & ".YPDCMVT0" & xWhere
Set rsSab = cnsab.Execute(xSQL)
ReDim arrYPDCMVT0(rsSab("Tally") + 1)

xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCMVT0" & xWhere & " order by PDCMVTDEV"
Set rsSab = cnsab.Execute(xSQL)



Do While Not rsSab.EOF
    V = rsYPDCMVT0_GetBuffer(rsSab, xYPDCMVT0)
    'wDev = rsSab("PDCMVTDEV")
    arrYPDCMVT0_Nb = arrYPDCMVT0_Nb + 1
    arrYPDCMVT0(arrYPDCMVT0_Nb) = xYPDCMVT0

    For I = 1 To arrYPDCPOS0_Nb
         If arrYPDCPOS0(I).PDCPOSDEV = xYPDCMVT0.PDCMVTDEV Then
            arrYPDCPOS0(I).PDCPOSPOSE = arrYPDCPOS0(I).PDCPOSPOSE - xYPDCMVT0.PDCMVTMTE ' rsSab("PDCMVTMTE")
            arrYPDCPOS0(I).PDCPOSPOSD = arrYPDCPOS0(I).PDCPOSPOSD - xYPDCMVT0.PDCMVTMTD 'rsSab("PDCMVTMTd")
            Exit For
         End If
    Next I
    rsSab.MoveNext

Loop

'For I = 1 To arrYPDCPOS0_Nb
'    arrYPDCPOS0(I).PDCPOSPNL = -Round((arrYPDCPOS0(I).PDCPOSPOSD + arrYPDCPOS0(I).PDCPOSPOSE * arrYPDCPOS0(I).PDCPOSFIXT) / arrYPDCPOS0(I).PDCPOSFIXT, 2)
'Next I

Set rsSab = Nothing

End Sub
Public Sub cmdSelect_SQL_1_YPDCOPE0_R(lAMJ As String)
Dim xSQL As String
Dim I As Integer, xWhere As String, X As String
Dim wMTE As Currency, wMTD As Currency, wDev As String, wOPEC As String

xWhere = " where PDCOPEDTR = '" & lAMJ & "' and PDCOPESTA = 'R'"

Call arrYPDCOPE0_SQL(xWhere)


For I = 1 To arrYPDCOPE0_Nb

    If arrYPDCOPE0(I).PDCOPESTA = "R" Then
        wOPEC = arrYPDCOPE0(I).PDCOPEOPEC
         If arrYPDCOPE0(I).PDCOPEDEV1 = "EUR" Then
             wDev = arrYPDCOPE0(I).PDCOPEDEV2
             wMTE = arrYPDCOPE0(I).PDCOPEMTD1
             wMTD = arrYPDCOPE0(I).PDCOPEMTD2
             Call cmdSelect_SQL_1Instant_YPDCPOS0(wOPEC, wDev, wMTD, wMTE)
         Else
            If arrYPDCOPE0(I).PDCOPEDEV2 = "EUR" Then
                wDev = arrYPDCOPE0(I).PDCOPEDEV1
                wMTE = arrYPDCOPE0(I).PDCOPEMTD2
                wMTD = arrYPDCOPE0(I).PDCOPEMTD1
                Call cmdSelect_SQL_1Instant_YPDCPOS0(wOPEC, wDev, wMTD, wMTE)
            Else
                    
                wDev = arrYPDCOPE0(I).PDCOPEDEV1
                wMTE = 0
                wMTD = arrYPDCOPE0(I).PDCOPEMTD1
                Call cmdSelect_SQL_1Instant_YPDCPOS0(wOPEC, wDev, wMTD, wMTE)
                wDev = arrYPDCOPE0(I).PDCOPEDEV2
                wMTE = -wMTE
                wMTD = arrYPDCOPE0(I).PDCOPEMTD2
                Call cmdSelect_SQL_1Instant_YPDCPOS0(wOPEC, wDev, wMTD, wMTE)
            End If
        End If
    End If

Next I


End Sub

Public Sub cmdSelect_SQL_1PDCMVTKCUT()
Dim wDev As String, xSQL As String
Dim I As Integer, X As String

ReDim arrYPDCMVT0(101)
arrYPDCMVT0_Max = 100: arrYPDCMVT0_Nb = 0
For I = 1 To arrDev_Nb
    arrKCUT(I).PDCPOSDEV = ""
    arrKCUT(I).PDCPOSPOSE = 0
    arrKCUT(I).PDCPOSPOSD = 0
Next I


xSQL = "select count(*) as Tally  from " & paramIBM_Library_SABSPE_XXX & ".YPDCMVT0 where PDCMVTKCUT <> ' '"
Set rsSab = cnsab.Execute(xSQL)
ReDim arrYPDCMVT0(rsSab("Tally") + 1)

xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCMVT0 where PDCMVTKCUT <> ' '  order by PDCMVTDEV"
Set rsSab = cnsab.Execute(xSQL)



Do While Not rsSab.EOF
    V = rsYPDCMVT0_GetBuffer(rsSab, xYPDCMVT0)
    'wDev = rsSab("PDCMVTDEV")
    arrYPDCMVT0_Nb = arrYPDCMVT0_Nb + 1
    arrYPDCMVT0(arrYPDCMVT0_Nb) = xYPDCMVT0

    For I = 1 To arrYPDCPOS0_Nb
         If arrYPDCPOS0(I).PDCPOSDEV = xYPDCMVT0.PDCMVTDEV Then
            arrYPDCPOS0(I).PDCPOSPOSE = arrYPDCPOS0(I).PDCPOSPOSE - xYPDCMVT0.PDCMVTMTE ' rsSab("PDCMVTMTE")
            arrYPDCPOS0(I).PDCPOSPOSD = arrYPDCPOS0(I).PDCPOSPOSD - xYPDCMVT0.PDCMVTMTD 'rsSab("PDCMVTMTd")
            arrKCUT(I).PDCPOSDEV = xYPDCMVT0.PDCMVTDEV
            arrKCUT(I).PDCPOSPOSE = arrKCUT(I).PDCPOSPOSE + xYPDCMVT0.PDCMVTMTE
            arrKCUT(I).PDCPOSPOSD = arrKCUT(I).PDCPOSPOSD + xYPDCMVT0.PDCMVTMTD
            If arrKCUT(I).PDCPOSPOSD = 0 Then
                arrKCUT(I).PDCPOSPRIX = 0
            Else
                arrKCUT(I).PDCPOSPRIX = Round(Abs(arrKCUT(I).PDCPOSPOSD / arrKCUT(I).PDCPOSPOSE), 5)
            End If
            
            Exit For
         End If
    Next I
    rsSab.MoveNext

Loop


Set rsSab = Nothing

End Sub

Public Sub YPDCLOG0_Init(lPDCLOGDTR As String)
Dim X As String
ReDim arrYPDCLOG0(101): arrYPDCLOG0_Nb = 0: arrYPDCLOG0_Max = 100
rsYPDCLOG0_Init arrYPDCLOG0(0)
arrYPDCLOG0(0).PDCLOGDTR = lPDCLOGDTR
arrYPDCLOG0(0).PDCLOGUAMJ = DSys
X = Time
arrYPDCLOG0(0).PDCLOGUHMS = Mid$(X, 1, 2) + Mid$(X, 4, 2) + Mid$(X, 7, 2)
arrYPDCLOG0(0).PDCLOGUUSR = usrName
xYPDCLOG0 = arrYPDCLOG0(0)

fgDetail_Display_YPDCLOG0
End Sub

Public Sub YPDCLOG0_AddItem()
arrYPDCLOG0_Nb = arrYPDCLOG0_Nb + 1
If arrYPDCLOG0_Nb > arrYPDCLOG0_Max Then
    arrYPDCLOG0_Max = arrYPDCLOG0_Max + 50
    ReDim Preserve arrYPDCLOG0(arrYPDCLOG0_Max)
End If
mPDCLOGUSEQ = mPDCLOGUSEQ + 1
xYPDCLOG0.PDCLOGUSEQ = mPDCLOGUSEQ
arrYPDCLOG0(arrYPDCLOG0_Nb) = xYPDCLOG0

fgDetail.Rows = fgDetail.Rows + 1
fgDetail.Row = fgDetail.Rows - 1
Call fgDetail_DisplayLine_YPDCLOG0(arrYPDCLOG0_Nb)
End Sub

Public Sub YBIATAB0_Fixing(lFIXING_DEV As String, lFIXING_AMJ As String)
Dim X As String, xMemo As String
On Error GoTo Exit_sub
If mFIXING_DEV = lFIXING_DEV And mFixing_AMJ = lFIXING_AMJ Then
Else
    mFIXING_DEV = lFIXING_DEV
    Call sqlYBIATAB0_Read("PDC", lFIXING_DEV, lFIXING_AMJ, xMemo)
    If IsNumeric(Mid$(xMemo, 9, 15)) Then
        mFixing_AMJ = lFIXING_AMJ
        mFIXING_Cours = CDbl(Mid$(xMemo, 9, 15) / 1000000000)
    Else
        X = DateComptablePrecedente(lFIXING_AMJ)
        Call sqlYBIATAB0_Read("PDC", lFIXING_DEV, X, xMemo)
        If IsNumeric(Mid$(xMemo, 9, 15)) Then
            mFixing_AMJ = X
            mFIXING_Cours = CDbl(Mid$(xMemo, 9, 15) / 1000000000)
        Else
            mFixing_AMJ = ""
            mFIXING_Cours = 0
        End If
    End If

End If
Exit_sub:
End Sub

Public Sub fraPDCOPE_Display_Montant(lPDCOPESENS As String)
If lPDCOPESENS = "A" Then
    cboPDCOPESENS.ListIndex = 0
    txtPDCOPEMTD1.ForeColor = vbBlue
    libPDCOPEMTD2.ForeColor = vbRed
    xYPDCOPE0.PDCOPEMTD1 = -Abs(xYPDCOPE0.PDCOPEMTD1)
    xYPDCOPE0.PDCOPEMTD2 = Abs(xYPDCOPE0.PDCOPEMTD2)
Else
    txtPDCOPEMTD1.ForeColor = vbRed
    libPDCOPEMTD2.ForeColor = vbBlue
    cboPDCOPESENS.ListIndex = cboPDCOPESENS.ListCount - 1
    xYPDCOPE0.PDCOPEMTD1 = Abs(xYPDCOPE0.PDCOPEMTD1)
    xYPDCOPE0.PDCOPEMTD2 = -Abs(xYPDCOPE0.PDCOPEMTD2)
End If
txtPDCOPEMTD1 = IIf(xYPDCOPE0.PDCOPEMTD1 = 0, "", Format$(Abs(xYPDCOPE0.PDCOPEMTD1), "### ### ### ##0.00"))
libPDCOPEMTD2 = IIf(xYPDCOPE0.PDCOPEMTD2 = 0, "", Format$(Abs(xYPDCOPE0.PDCOPEMTD2), "### ### ### ##0.00"))

If cboPDCOPEOPEC = "SWP" Then txtPDCOPEVTXT.ForeColor = libPDCOPEMTD2.ForeColor

End Sub
Public Sub fraSuspens_Display_Montant()

If Mid$(cboSuspens_PDCMVTSENS, 1, 1) = "A" Then
    txtSuspens_PDCMVTMTD1.ForeColor = vbBlue
    txtSuspens_PDCMVTMTD2.ForeColor = vbRed
Else
    txtSuspens_PDCMVTMTD1.ForeColor = vbRed
    txtSuspens_PDCMVTMTD2.ForeColor = vbBlue
End If


End Sub

Public Sub fraPDCOPE_Display_CLI(lPDCOPECLI As String)
If Trim(lPDCOPECLI) = "" Then
    txtPDCOPECLI = ""
    libPDCOPECLI = ""
Else
    txtPDCOPECLI = xYPDCOPE0.PDCOPECLI
    Call ZCLIEAN0_SQL(xYPDCOPE0.PDCOPECLI)
    libPDCOPECLI = mCLIENARA1
End If

End Sub

Public Sub fraSuspens_Display_CLI(lPDCMVTCLI As String)
If Trim(lPDCMVTCLI) = "" Then
    txtSuspens_PDCMVTCLI = ""
    libSuspens_PDCMVTCLI = ""
Else
    txtSuspens_PDCMVTCLI = xYPDCMVT0.PDCMVTCLI
    Call ZCLIEAN0_SQL(xYPDCMVT0.PDCMVTCLI)
    libSuspens_PDCMVTCLI = mCLIENARA1
End If

End Sub

Public Sub cmdSelect_SQL_1Instant()
Dim I As Integer, xMemo As String, K As Integer
Dim wMTE As Currency, wMTD As Currency, wDev As String, wOPEC As String

For I = 1 To arrYPDCPOS0_Nb
    xYPDCPOS0 = arrYPDCPOS0(I)
    xYPDCPOS0.PDCPOSDTR = mPDCPOSDTR
    xYPDCPOS0.PDCPOSPOSE = xYPDCPOS0.PDCPOSPOSE + xYPDCPOS0.PDCPOSRPC
    xYPDCPOS0.PDCPOSRPC = 0
    If xYPDCPOS0.PDCPOSFIXT = 0 Then
        xYPDCPOS0.PDCPOSPNL = 0
    Else
        xYPDCPOS0.PDCPOSPNL = -Round((xYPDCPOS0.PDCPOSPOSD + xYPDCPOS0.PDCPOSPOSE * xYPDCPOS0.PDCPOSFIXT) / xYPDCPOS0.PDCPOSFIXT, 2)
    End If
    '________________________________________________________________________________
    Call sqlYBIATAB0_Read("PDC", xYPDCPOS0.PDCPOSDEV, mPDCPOSDTR, xMemo)
    If IsNumeric(Mid$(xMemo, 9, 15)) Then
        xYPDCPOS0.PDCPOSFIXT = CDbl(Mid$(xMemo, 9, 15) / 1000000000)
        xYPDCPOS0.PDCPOSFIXD = Mid$(xMemo, 1, 8)
   End If
   arrYPDCPOS0(I) = xYPDCPOS0
Next I

cmdSelect_SQL_2
For I = 1 To arrYPDCOPE0_Nb
    If arrYPDCOPE0(I).PDCOPESTA = "T" Or arrYPDCOPE0(I).PDCOPESTA = "V" Then
        wOPEC = arrYPDCOPE0(I).PDCOPEOPEC
         If arrYPDCOPE0(I).PDCOPEDEV1 = "EUR" Then
             wDev = arrYPDCOPE0(I).PDCOPEDEV2
             wMTE = arrYPDCOPE0(I).PDCOPEMTD1
             wMTD = arrYPDCOPE0(I).PDCOPEMTD2
             Call cmdSelect_SQL_1Instant_YPDCPOS0(wOPEC, wDev, wMTD, wMTE)
         Else
            If arrYPDCOPE0(I).PDCOPEDEV2 = "EUR" Then
                wDev = arrYPDCOPE0(I).PDCOPEDEV1
                wMTE = arrYPDCOPE0(I).PDCOPEMTD2
                wMTD = arrYPDCOPE0(I).PDCOPEMTD1
                Call cmdSelect_SQL_1Instant_YPDCPOS0(wOPEC, wDev, wMTD, wMTE)
            Else
                    
                wDev = arrYPDCOPE0(I).PDCOPEDEV1
                wMTE = 0
                wMTD = arrYPDCOPE0(I).PDCOPEMTD1
                Call cmdSelect_SQL_1Instant_YPDCPOS0(wOPEC, wDev, wMTD, wMTE)
                wDev = arrYPDCOPE0(I).PDCOPEDEV2
                wMTE = -wMTE
                wMTD = arrYPDCOPE0(I).PDCOPEMTD2
                Call cmdSelect_SQL_1Instant_YPDCPOS0(wOPEC, wDev, wMTD, wMTE)
            End If
        End If
    End If
Next I

End Sub

Public Sub cmdSelect_SQL_1Instant_YPDCPOS0(lOPEC As String, lDev As String, lMTD As Currency, lMTE As Currency)
Dim I As Integer, xMemo As String, K As Integer

For K = 1 To arrYPDCPOS0_Nb
    If arrYPDCPOS0(K).PDCPOSDEV = lDev Then
        If lMTE = 0 Then
            If arrYPDCPOS0(K).PDCPOSFIXT <> 0 Then lMTE = -Round(lMTD / arrYPDCPOS0(K).PDCPOSFIXT, 2)
        End If
        Select Case lOPEC
            Case "TER": arrYPDCPOS0(K).PDCPOSTERD = arrYPDCPOS0(K).PDCPOSTERD + lMTD
                        arrYPDCPOS0(K).PDCPOSTERE = arrYPDCPOS0(K).PDCPOSTERE + lMTE
            Case "SWP": arrYPDCPOS0(K).PDCPOSSWPD = arrYPDCPOS0(K).PDCPOSSWPD + lMTD
                        arrYPDCPOS0(K).PDCPOSSWPE = arrYPDCPOS0(K).PDCPOSSWPE + lMTE
            Case Else
                arrYPDCPOS0(K).PDCPOSPOSE = arrYPDCPOS0(K).PDCPOSPOSE + lMTE
                arrYPDCPOS0(K).PDCPOSPOSD = arrYPDCPOS0(K).PDCPOSPOSD + lMTD
                If arrYPDCPOS0(K).PDCPOSPOSE = 0 Then
                    arrYPDCPOS0(K).PDCPOSPRIX = 0
                Else
                    arrYPDCPOS0(K).PDCPOSPRIX = Round(Abs(arrYPDCPOS0(K).PDCPOSPOSD / arrYPDCPOS0(K).PDCPOSPOSE), 6)
                End If
            End Select
        Exit For
    End If
Next K

End Sub

Public Function fraPDCOPE_Control_Ticket() As String
Dim wMsg As String, xSQL As String
Dim blnOk As Boolean
blnOk = False
wMsg = ""
xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCOPE0 where PDCOPEREF = " & xYPDCOPE0.PDCOPEREF _
     & " and PDCOPEDTR = '" & xYPDCOPE0.PDCOPEDTR & "'" _
     & " and PDCOPESSE = 'TC'"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    If rsSab("PDCOPESTA") <> "A" Then
        If BIA_PDC_Aut.Rapprocher Then
            blnOk = True
            Call rsYPDCOPE0_GetBuffer(rsSab, memoYPDCOPE0)
        Else
            blnPDCOPE_Control_Ok = False
            wMsg = "- ticket déjà saisi" & vbCrLf
        End If
        Exit Do
            
    
    End If
    rsSab.MoveNext
Loop
If BIA_PDC_Aut.Rapprocher And Not blnOk Then
    blnPDCOPE_Control_Ok = False
    wMsg = "- ticket inconnu" & vbCrLf
End If
fraPDCOPE_Control_Ticket = wMsg
End Function

Public Function fraPDCOPE_Control_PDCOPEOPEN() As String
Dim wMsg As String, xSQL As String
Dim blnOk As Boolean
blnOk = False
wMsg = ""
xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCOPE0 where PDCOPEOPEN = " & xYPDCOPE0.PDCOPEOPEN _
     & " and PDCOPEOPEC = '" & xYPDCOPE0.PDCOPEOPEC & "'" _
     & " and PDCOPESER = '" & xYPDCOPE0.PDCOPESER & "'" _
     & " and PDCOPESSE = '" & xYPDCOPE0.PDCOPESSE & "'"
     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    If rsSab("PDCOPESTA") <> "A" Then
            blnPDCOPE_Control_Ok = False
            wMsg = "- numéro d'opération déjà saisi" & vbCrLf
        Exit Do
    End If
    rsSab.MoveNext
Loop
fraPDCOPE_Control_PDCOPEOPEN = wMsg
End Function

Public Function fraPDCOPE_Control_memoYPDCOPE0() As String
Dim wMsg As String
wMsg = ""
If xYPDCOPE0.PDCOPEOPEC <> memoYPDCOPE0.PDCOPEOPEC Then
    blnPDCOPE_Control_Ok = False
    wMsg = wMsg & "- natures différentes : CPT, TER, SWP" & vbCrLf
End If
If xYPDCOPE0.PDCOPEDTR <> memoYPDCOPE0.PDCOPEDTR Then
    blnPDCOPE_Control_Ok = False
    wMsg = wMsg & "- dates comptables différentes" & vbCrLf
End If
If xYPDCOPE0.PDCOPEREF <> memoYPDCOPE0.PDCOPEREF Then
    blnPDCOPE_Control_Ok = False
    wMsg = wMsg & "- tickets différents" & vbCrLf
End If
If xYPDCOPE0.PDCOPEOPEN = 0 Then
    blnPDCOPE_Control_Ok = False
    wMsg = wMsg & "- préciser le n° d'opération SAB" & vbCrLf
Else
    wMsg = wMsg & fraPDCOPE_Control_PDCOPEOPEN
End If
If xYPDCOPE0.PDCOPESENS <> memoYPDCOPE0.PDCOPESENS Then
    blnPDCOPE_Control_Ok = False
    wMsg = wMsg & "- sens différents" & vbCrLf
End If
If xYPDCOPE0.PDCOPEDEV1 <> memoYPDCOPE0.PDCOPEDEV1 Then
    blnPDCOPE_Control_Ok = False
    wMsg = wMsg & "- devises principales différentes" & vbCrLf
End If
If xYPDCOPE0.PDCOPEDEV2 <> memoYPDCOPE0.PDCOPEDEV2 Then
    blnPDCOPE_Control_Ok = False
    wMsg = wMsg & "- devises secondaires différentes" & vbCrLf
End If
If xYPDCOPE0.PDCOPEMTD1 <> memoYPDCOPE0.PDCOPEMTD1 Then
    blnPDCOPE_Control_Ok = False
    wMsg = wMsg & "- montants différents" & vbCrLf
End If
If xYPDCOPE0.PDCOPEDVA <> memoYPDCOPE0.PDCOPEDVA Then
    blnPDCOPE_Control_Ok = False
    wMsg = wMsg & "- dates de valeur différentes" & vbCrLf
End If
If xYPDCOPE0.PDCOPETAUX <> memoYPDCOPE0.PDCOPETAUX Then
    blnPDCOPE_Control_Ok = False
    wMsg = wMsg & "- cours différents" & vbCrLf
End If
If xYPDCOPE0.PDCOPECLI <> memoYPDCOPE0.PDCOPECLI Then
    blnPDCOPE_Control_Ok = False
    wMsg = wMsg & "- clients différents" & vbCrLf
End If
If xYPDCOPE0.PDCOPEVTXT <> memoYPDCOPE0.PDCOPEVTXT Then
    blnPDCOPE_Control_Ok = False
    wMsg = wMsg & "- cours à terme différents" & vbCrLf
End If

fraPDCOPE_Control_memoYPDCOPE0 = wMsg
End Function

Public Sub fraPDCOPE_Init()
Dim blnPDCOPEOPEC_Enabled As Boolean


Call rsYPDCOPE0_Init(oldYPDCOPE0)
oldYPDCOPE0.PDCOPEDTR = YBIATAB0_DATE_CPT_JS1
oldYPDCOPE0.PDCOPESENS = "A"
oldYPDCOPE0.PDCOPESENX = "1"
oldYPDCOPE0.PDCOPEDEV1 = "   "
oldYPDCOPE0.PDCOPEDEV2 = "   "
oldYPDCOPE0.PDCOPEDVA = mPDCOPEDVA_2J


cboPDCOPESER.Enabled = True
txtPDCOPEVTXT.Enabled = False
txtPDCOPEFIXING.Enabled = False

mPDCOPEMTD1 = 0

blnPDCOPEOPEC_Enabled = False
'______________________________________
' pour TC
'___________________________

If blnPDCOPE_CONF_CALL_Saisie Then
            oldYPDCOPE0.PDCOPEOPEC = "CPT"
            oldYPDCOPE0.PDCOPEOPET = "   "
            oldYPDCOPE0.PDCOPESER = "00"
            oldYPDCOPE0.PDCOPESSE = "TR"
            cboPDCOPESER.Enabled = False
            blnPDCOPEOPEC_Enabled = True
            cboPDCOPEOPET.Enabled = False
            txtPDCOPEOPEN.Enabled = False
            txtPDCOPEREF.Enabled = False
            oldYPDCOPE0.PDCOPEITXT = ""
            oldYPDCOPE0.PDCOPEVTXT = "CONF_CALL"

Else
    Select Case localUnit
        Case "FOTC"
            oldYPDCOPE0.PDCOPEOPEC = "CPT"
            oldYPDCOPE0.PDCOPEOPET = "   "
            oldYPDCOPE0.PDCOPESER = "TC"
            oldYPDCOPE0.PDCOPESSE = "TC"
            cboPDCOPESER.Enabled = False
            blnPDCOPEOPEC_Enabled = True
            cboPDCOPEOPET.Enabled = False
            txtPDCOPEOPEN.Enabled = False
            txtPDCOPEREF.Enabled = True
        Case "GDC"
            oldYPDCOPE0.PDCOPEOPEC = "CPT"
            oldYPDCOPE0.PDCOPEOPET = "   "
            oldYPDCOPE0.PDCOPESER = "TC"
            oldYPDCOPE0.PDCOPESSE = "TC"
            cboPDCOPESER.Enabled = False
            blnPDCOPEOPEC_Enabled = True
            cboPDCOPEOPET.Enabled = False
            txtPDCOPEREF.Enabled = True
    
        Case "GDMP"
            oldYPDCOPE0.PDCOPEOPEC = "CPT"
            oldYPDCOPE0.PDCOPEOPET = "   "
            oldYPDCOPE0.PDCOPESER = "00"
            oldYPDCOPE0.PDCOPESSE = "TR"
            cboPDCOPESER.Enabled = False
            cboPDCOPEOPET.Enabled = False
            txtPDCOPEREF.Enabled = False
        Case "SOBI"
            oldYPDCOPE0.PDCOPEOPEC = "CDE"
            oldYPDCOPE0.PDCOPEOPET = "   "
            oldYPDCOPE0.PDCOPESER = "00"
            oldYPDCOPE0.PDCOPESSE = "CD"
            oldYPDCOPE0.PDCOPESENS = "V"
            oldYPDCOPE0.PDCOPEDEV2 = "EUR"
            
            cboPDCOPESER.Enabled = False
            blnPDCOPEOPEC_Enabled = True 'False
            cboPDCOPEOPET.Enabled = False
            txtPDCOPEREF.Enabled = False
        Case "CPT"
            cboPDCOPEOPEC.AddItem "OD "
            oldYPDCOPE0.PDCOPEOPEC = "OD "
            oldYPDCOPE0.PDCOPEOPET = "   "
            oldYPDCOPE0.PDCOPESER = "CP"
            oldYPDCOPE0.PDCOPESSE = "CP"
            cboPDCOPESER.Enabled = False
            cboPDCOPEOPET.Enabled = False
            txtPDCOPEREF.Enabled = False
        Case Else
            oldYPDCOPE0.PDCOPEOPEC = "CPT"
            oldYPDCOPE0.PDCOPEOPET = "   "
            oldYPDCOPE0.PDCOPESER = "00"
            oldYPDCOPE0.PDCOPESSE = "00"
            cboPDCOPESER.Enabled = False
            cboPDCOPEOPET.Enabled = False
            txtPDCOPEREF.Enabled = False
    End Select
End If
'_____________________________________
xYPDCOPE0 = oldYPDCOPE0
fraPDCOPE_Display


lblPDCOPEDVA.BackColor = &HFFA0FF ' magenta&H00FFE0FF&
lblPDCOPEDEV1.BackColor = &HFFA0FF ' magenta
If localUnit = "SOBI" Then
    lblPDCOPEDEV2.BackColor = &H80C0FF
Else
    lblPDCOPEDEV2.BackColor = &HFFA0FF ' magenta
End If
chkPDCOPEINFO.BackColor = &HFFA0FF
chkPDCOPEINFO.Value = "0"
chkPDCOPEINFO.Caption = "demande pour information ?"
chkPDCOPEINFO.Visible = True
lblPDCOPETAUX.ForeColor = vbMagenta
fraPDCOPE_Control_DEV_Certain

cmdPDCOPE_Update.Visible = True
blnPDCOPE_Control_S = True
Me.Enabled = True: Me.MousePointer = 0
fraPDCOPE_S.Enabled = True
fraPDCOPE_V.Enabled = blnPDCOPE_Control_V

'$JPL_2012-07-10 : saisie du cours par stagiaire TC
If localUnit = "FOTC" Then fraPDCOPE_V.Enabled = True

blnPDCOPEOPEC_Enabled = blnPDCOPEOPEC_Enabled

'If localUnit = "FOTC" Or localUnit = "GDC " Then
If txtPDCOPEREF.Enabled Then
    txtPDCOPEREF.SetFocus
Else
'    If txtPDCOPEMTD1.Enabled Then txtPDCOPEMTD1.SetFocus
    If cboPDCOPEDEV1.Enabled Then cboPDCOPEDEV1.SetFocus
End If

End Sub
Public Sub fraSuspens_Init()
Call rsYPDCMVT0_Init(oldYPDCMVT0)
oldYPDCMVT0.PDCMVTDTR = YBIATAB0_DATE_CPT_J
oldYPDCMVT0.PDCMVTOPEC = "" '"XXC"
oldYPDCMVT0.PDCMVTCPT = "Achat USD/EUR Suspens FOTC"
oldYPDCMVT0.PDCMVTSTA = "+"
oldYPDCMVT0.PDCMVTDEV = ""
oldYPDCMVT0.PDCMVTDVA = YBIATAB0_DATE_CPT_M
oldYPDCMVT0.PDCMVTSER = "TC"
oldYPDCMVT0.PDCMVTSSE = "TC"
xYPDCMVT0 = oldYPDCMVT0

fgSuspens_Log.Clear
txtSuspens_Comment = ""

fraSuspens_Display
cmdSuspens_Update.Caption = constAjouter
cmdSuspens_Update.Visible = True
cmdSuspens_Modification.Visible = False
fraSuspens_S.Enabled = True
fraSuspens_M.Enabled = True
Me.Enabled = True: Me.MousePointer = 0
fraSuspens_S.Enabled = True
If txtSuspens_PDCMVTMTD1.Enabled Then
    txtSuspens_PDCMVTMTD1.SetFocus
End If

End Sub


Private Sub txtSelect_AMJ_HB_Change()
cmdSelect_Reset

End Sub


Private Sub txtSelect_AMJ_xls_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_File_xls_GotFocus()
Call txt_GotFocus(txtSelect_File_xls)

End Sub

Private Sub txtSelect_File_xls_LostFocus()
Call txt_LostFocus(txtSelect_File_xls)

End Sub

Private Sub txtSelect_Sheet_xls_GotFocus()
Call txt_GotFocus(txtSelect_Sheet_xls)

End Sub

Private Sub txtSelect_Sheet_xls_LostFocus()
Call txt_LostFocus(txtSelect_Sheet_xls)

End Sub

Private Sub txtSuspens_PDCMVTCLI_GotFocus()
Call txt_GotFocus(txtSuspens_PDCMVTCLI)

End Sub

Private Sub txtSuspens_PDCMVTCLI_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtSuspens_PDCMVTCLI_LostFocus()
Call txt_LostFocus(txtSuspens_PDCMVTCLI)

End Sub

Private Sub txtSuspens_PDCMVTMTD1_Change()
Call fraSuspens_Display_PDCMVTTAUX

End Sub

Private Sub txtSuspens_PDCMVTMTD2_Change()
Call fraSuspens_Display_PDCMVTTAUX

End Sub

Private Sub txtSuspens_PDCMVTMTD2_GotFocus()
txtSuspens_PDCMVTMTD2.BackColor = focusUsr.BackColor

End Sub


Private Sub txtSuspens_PDCMVTMTD2_KeyPress(KeyAscii As Integer)
Call num_Montant(KeyAscii, txtSuspens_PDCMVTMTD2)

End Sub


Private Sub txtSuspens_PDCMVTMTD2_LostFocus()
txtSuspens_PDCMVTMTD2.BackColor = txtUsr.BackColor

End Sub


Private Sub txtSuspens_PDCMVTMTD1_GotFocus()
txtSuspens_PDCMVTMTD1.BackColor = focusUsr.BackColor

End Sub


Private Sub txtSuspens_PDCMVTMTD1_KeyPress(KeyAscii As Integer)
Call num_Montant(KeyAscii, txtSuspens_PDCMVTMTD1)

End Sub


Private Sub txtSuspens_PDCMVTMTD1_LostFocus()
txtSuspens_PDCMVTMTD1.BackColor = txtUsr.BackColor

End Sub



Public Function fraSuspens_Control()
Dim X As String, wMsg As String, xSQL As String, K As Integer
Dim wPDCMVTSENS As String, wPDCMVTDEV1 As String, wPDCMVTDEV2 As String
blnSuspens_Control_Ok = True
Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données "): DoEvents
xYPDCMVT0 = oldYPDCMVT0
wPDCMVTSENS = Mid$(cboSuspens_PDCMVTSENS, 1, 1)
wPDCMVTDEV1 = cboSuspens_PDCMVTDEV1
wPDCMVTDEV2 = cboSuspens_PDCMVTDEV2

If wPDCMVTDEV1 = wPDCMVTDEV2 Then
    blnSuspens_Control_Ok = False
    wMsg = wMsg & "- les deux devises doivent être différentes " & wPDCMVTDEV1 & vbCrLf
End If

If wPDCMVTSENS = "A" Then
    xYPDCMVT0.PDCMVTCPT = "Achat " & wPDCMVTDEV1 & "/" & wPDCMVTDEV2 & "suspens"
Else
    xYPDCMVT0.PDCMVTCPT = "Vente " & wPDCMVTDEV1 & "/" & wPDCMVTDEV2 & "suspens"
End If

If optSuspens_XXC Then
    xYPDCMVT0.PDCMVTOPEC = "XXC"
Else
    If optSuspens_XXT Then
        xYPDCMVT0.PDCMVTOPEC = "XXT"
    Else
        blnSuspens_Control_Ok = False
        wMsg = wMsg & "- préciser comptant ou terme" & vbCrLf
    End If
End If

If wPDCMVTDEV1 = "EUR" Then
    xYPDCMVT0.PDCMVTMTE = CCur(num_CDec(txtSuspens_PDCMVTMTD1))
    xYPDCMVT0.PDCMVTMTD = CCur(num_CDec(txtSuspens_PDCMVTMTD2))
    If wPDCMVTSENS = "A" Then
        xYPDCMVT0.PDCMVTMTE = -Abs(xYPDCMVT0.PDCMVTMTE)
        xYPDCMVT0.PDCMVTMTD = Abs(xYPDCMVT0.PDCMVTMTD)
    Else
        xYPDCMVT0.PDCMVTMTE = Abs(xYPDCMVT0.PDCMVTMTE)
        xYPDCMVT0.PDCMVTMTD = -Abs(xYPDCMVT0.PDCMVTMTD)
    End If
    
    xYPDCMVT0.PDCMVTDEV = wPDCMVTDEV2
Else
    If wPDCMVTDEV2 = "EUR" Then
        xYPDCMVT0.PDCMVTMTD = CCur(num_CDec(txtSuspens_PDCMVTMTD1))
        xYPDCMVT0.PDCMVTMTE = CCur(num_CDec(txtSuspens_PDCMVTMTD2))
        xYPDCMVT0.PDCMVTDEV = wPDCMVTDEV1
        If wPDCMVTSENS = "A" Then
            xYPDCMVT0.PDCMVTMTE = Abs(xYPDCMVT0.PDCMVTMTE)
            xYPDCMVT0.PDCMVTMTD = -Abs(xYPDCMVT0.PDCMVTMTD)
        Else
            xYPDCMVT0.PDCMVTMTE = -Abs(xYPDCMVT0.PDCMVTMTE)
            xYPDCMVT0.PDCMVTMTD = Abs(xYPDCMVT0.PDCMVTMTD)
        End If
    Else
        blnSuspens_Control_Ok = False
        wMsg = wMsg & "- une des deux devises doit être EUR" & vbCrLf
    End If
End If
fraSuspens_Display_Montant
Call fraSuspens_Display_PDCMVTTAUX
If xYPDCMVT0.PDCMVTMTE = 0 Then
    blnSuspens_Control_Ok = False
    wMsg = wMsg & "- préciser le montant en EUR" & vbCrLf
End If
If xYPDCMVT0.PDCMVTMTD = 0 Then
    blnSuspens_Control_Ok = False
    wMsg = wMsg & "- préciser le montant en " & xYPDCMVT0.PDCMVTDEV & vbCrLf
End If

    '________________________________
xYPDCMVT0.PDCMVTSER = Mid$(cboSuspens_PDCMVTSER, 1, 2)
xYPDCMVT0.PDCMVTSSE = Mid$(cboSuspens_PDCMVTSER, 4, 2)
Call DTPicker_Control(txtSuspens_PDCMVTDVA, xYPDCMVT0.PDCMVTDVA)

'If cmdSuspens_Update.Caption <> constAjouter Then

    If xYPDCMVT0.PDCMVTDVA < YBIATAB0_DATE_CPT_JS1 Then
        blnSuspens_Control_Ok = False
        wMsg = wMsg & "- date valeur < date du jour" & vbCrLf
    End If
    For K = 1 To arrDevF_Nb
        
        arrDevF_ISO(arrDevF_Nb) = Mid$(X, 3, 3)
        If arrDevF_AMJ(K) = xYPDCMVT0.PDCMVTDVA Then
            If arrDevF_ISO(K) = wPDCMVTDEV1 Then
                blnSuspens_Control_Ok = False
                wMsg = wMsg & "- date valeur = jour férié " & wPDCMVTDEV1 & vbCrLf
            End If
            If arrDevF_ISO(K) = wPDCMVTDEV2 Then
                blnSuspens_Control_Ok = False
                wMsg = wMsg & "- date valeur = jour férié " & wPDCMVTDEV2 & vbCrLf
            End If
        End If
    Next K

    K = Weekday(dateImp(xYPDCMVT0.PDCMVTDVA))
    If K = 1 Then
        blnSuspens_Control_Ok = False
        wMsg = wMsg & "- date valeur = Dimanche " & vbCrLf
    Else
        If K = 7 Then
            blnSuspens_Control_Ok = False
            wMsg = wMsg & "- date valeur = Samedi " & vbCrLf
        End If
    End If
    
    
'End If
xYPDCMVT0.PDCMVTCLI = Format$(txtSuspens_PDCMVTCLI, "0000000")
If Trim(xYPDCMVT0.PDCMVTCLI) <> "" Then
    Call fraSuspens_Display_CLI(xYPDCMVT0.PDCMVTCLI)
    If mCLIENARA1 = "" Then
        blnSuspens_Control_Ok = False
        wMsg = wMsg & "- contrepartie inconnue" & vbCrLf
    End If
End If


If blnSuspens_Control_Ok Then
    fraSuspens_Control = Null
    newYPDCMVT0 = xYPDCMVT0
    If cmdSuspens_Update.Caption <> constAjouter Then
        newYPDCMVT0.PDCMVTSTA2 = "="
        memoYPDCMVT0 = newYPDCMVT0
        newYPDCMVT0.PDCMVTDTR = YBIATAB0_DATE_CPT_J
        newYPDCMVT0.PDCMVTDVA = newYPDCMVT0.PDCMVTDTR
        newYPDCMVT0.PDCMVTECR = 0
        newYPDCMVT0.PDCMVTSTA = "-"
        newYPDCMVT0.PDCMVTCPT = Replace(newYPDCMVT0.PDCMVTCPT, "suspens", "annul")
        newYPDCMVT0.PDCMVTMTE = -newYPDCMVT0.PDCMVTMTE
        newYPDCMVT0.PDCMVTMTD = -newYPDCMVT0.PDCMVTMTD
    End If
        
Else
    Call MsgBox(wMsg, vbCritical, "BIA_PDC : Gestion des suspens")
    fraSuspens_Control = "?_________fraSuspens_Control"
End If

End Function

Public Sub fraSuspens_Display_PDCMVTTAUX()
Dim xCur1 As Currency, xCur2 As Currency, xDbl As Double
xCur1 = num_CDec(txtSuspens_PDCMVTMTD1)
xCur2 = num_CDec(txtSuspens_PDCMVTMTD2)
If xCur1 = 0 Or xCur2 = 0 Then
    xDbl = 0
Else
    If cboSuspens_PDCMVTDEV1 = "EUR" Then
         xDbl = Round((xCur2 / xCur1), 6)
    Else
         xDbl = Round((xCur1 / xCur2), 6)
    End If
End If
libSuspens_PDCMVTTAUX = "Cours : " & Format$(xDbl, "### ###.### ###")
xYPDCMVT0.PDCMVTTAUX = xDbl
End Sub

Public Function cmdSelect_SQL_5OPE_TC()
Dim K As Long, blnOk As Boolean
Dim curX1 As Currency, curX2 As Currency

blnOk = False
cmdSelect_SQL_5OPE_TC = "?"

For I = 1 To arrYPDCOPE0_Nb
    If xZCHGOPE0.CHGOPEDOS = arrYPDCOPE0(I).PDCOPEOPEN _
   And xZCHGOPE0.CHGOPEOPE = arrYPDCOPE0(I).PDCOPEOPEC _
   And arrYPDCOPE0(I).PDCOPESSE = "TC" Then ' _
    'And arrYPDCOPE0(I).PDCOPESTA2 = " " Then
        arrYPDCOPE0(I).PDCOPESTA2 = "#"
        curX1 = (Abs(Abs(xZCHGOPE0.CHGOPEMO1) - Abs(arrYPDCOPE0(I).PDCOPEMTD1)))
        curX2 = (Abs(Abs(xZCHGOPE0.CHGOPEMO2) - Abs(arrYPDCOPE0(I).PDCOPEMTD2)))
        If xZCHGOPE0.CHGOPEDE1 = arrYPDCOPE0(I).PDCOPEDEV1 _
        And xZCHGOPE0.CHGOPEDE2 = arrYPDCOPE0(I).PDCOPEDEV2 _
        And curX1 < 1 And curX2 < 1 Then
            blnOk = True
            Exit For
        Else
        
            curX1 = (Abs(Abs(xZCHGOPE0.CHGOPEMO1) - Abs(arrYPDCOPE0(I).PDCOPEMTD2)))
            curX2 = (Abs(Abs(xZCHGOPE0.CHGOPEMO2) - Abs(arrYPDCOPE0(I).PDCOPEMTD1)))
            If xZCHGOPE0.CHGOPEDE1 = arrYPDCOPE0(I).PDCOPEDEV2 _
            And xZCHGOPE0.CHGOPEDE2 = arrYPDCOPE0(I).PDCOPEDEV1 _
            And curX1 < 1 And curX2 < 1 Then
                blnOk = True
                Exit For

            End If
        End If
    End If
Next I
If blnOk Then
    cmdSelect_SQL_5OPE_TC = Null
    arrYPDCOPE0(I).PDCOPESTA2 = "="
End If
    
End Function

Public Function fraPDCOPE_Control_ZCHGOPE0()
Dim K As Long, blnOk As Boolean
Dim curX1 As Currency, curX2 As Currency
Dim wMsg As String, xSQL As String
wMsg = ""
blnOk = False
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCHGOPE0 " _
     & " where  CHGOPEDOS = " & xYPDCOPE0.PDCOPEOPEN _
     & " and  CHGOPEOPE = '" & xYPDCOPE0.PDCOPEOPEC & "'" _
     & " and  CHGOPESER = '" & xYPDCOPE0.PDCOPESER & "'" _
     & " and  CHGOPESSE = '" & xYPDCOPE0.PDCOPESSE & "'"

Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    blnPDCOPE_Control_Ok = False
    wMsg = "- n° d'opération inconnu dans SAB" & vbCrLf
Else
    
    V = rsZCHGOPE0_GetBuffer(rsSab, xZCHGOPE0)

    curX1 = (Abs(Abs(xZCHGOPE0.CHGOPEMO1) - Abs(xYPDCOPE0.PDCOPEMTD1)))
    curX2 = (Abs(Abs(xZCHGOPE0.CHGOPEMO2) - Abs(xYPDCOPE0.PDCOPEMTD2)))
    If xZCHGOPE0.CHGOPEDE1 = xYPDCOPE0.PDCOPEDEV1 _
    And xZCHGOPE0.CHGOPEDE2 = xYPDCOPE0.PDCOPEDEV2 _
    And curX1 < 1 And curX2 < 1 Then
        blnOk = True
    Else
    
        curX1 = (Abs(Abs(xZCHGOPE0.CHGOPEMO1) - Abs(xYPDCOPE0.PDCOPEMTD2)))
        curX2 = (Abs(Abs(xZCHGOPE0.CHGOPEMO2) - Abs(xYPDCOPE0.PDCOPEMTD1)))
        If xZCHGOPE0.CHGOPEDE1 = xYPDCOPE0.PDCOPEDEV2 _
        And xZCHGOPE0.CHGOPEDE2 = xYPDCOPE0.PDCOPEDEV1 _
        And curX1 < 1 And curX2 < 1 Then
            blnOk = True
        End If
    End If
    
    If Not blnOk Then
        blnPDCOPE_Control_Ok = False
        wMsg = "- opération SAB " & xYPDCOPE0.PDCOPEOPEN & " = " _
             & xYPDCOPE0.PDCOPEDEV1 & " " & Format$(Abs(xZCHGOPE0.CHGOPEMO1), "##### ### ###.00") & " / " _
             & xYPDCOPE0.PDCOPEDEV2 & " " & Format$(Abs(xZCHGOPE0.CHGOPEMO2), "##### ### ###.00") _
             & vbCrLf
   End If
End If
fraPDCOPE_Control_ZCHGOPE0 = wMsg
End Function

Public Function cmdSelect_SQL_5OPE_XX()
Dim K As Long, blnOk As Boolean
Dim curX1 As Currency, curX2 As Currency

blnOk = False
cmdSelect_SQL_5OPE_XX = "?"
For I = 1 To arrYPDCOPE0_Nb
        If xZCHGOPE0.CHGOPECON = arrYPDCOPE0(I).PDCOPECLI _
        And arrYPDCOPE0(I).PDCOPESSE = "TR" _
        And xZCHGOPE0.CHGOPEOPE = arrYPDCOPE0(I).PDCOPEOPEC _
        And arrYPDCOPE0(I).PDCOPESTA2 = " " Then
        
        'And xZCHGOPE0.CHGOPESSE = arrYPDCOPE0(I).PDCOPESSE Then ' _
        'And arrYPDCOPE0(I).PDCOPESTA2 = " " Then
        curX1 = (Abs(Abs(xZCHGOPE0.CHGOPEMO1) - Abs(arrYPDCOPE0(I).PDCOPEMTD1)))
        curX2 = (Abs(Abs(xZCHGOPE0.CHGOPEMO2) - Abs(arrYPDCOPE0(I).PDCOPEMTD2)))
        If xZCHGOPE0.CHGOPEDE1 = arrYPDCOPE0(I).PDCOPEDEV1 _
        And xZCHGOPE0.CHGOPEDE2 = arrYPDCOPE0(I).PDCOPEDEV2 _
        And curX1 < 1 And curX2 < 1 Then
            blnOk = True
            Exit For
        Else
        
            curX1 = (Abs(Abs(xZCHGOPE0.CHGOPEMO1) - Abs(arrYPDCOPE0(I).PDCOPEMTD2)))
            curX2 = (Abs(Abs(xZCHGOPE0.CHGOPEMO2) - Abs(arrYPDCOPE0(I).PDCOPEMTD1)))
            If xZCHGOPE0.CHGOPEDE1 = arrYPDCOPE0(I).PDCOPEDEV2 _
            And xZCHGOPE0.CHGOPEDE2 = arrYPDCOPE0(I).PDCOPEDEV1 _
            And curX1 < 1 And curX2 < 1 Then
                blnOk = True
                Exit For

            End If
        End If
    End If
Next I
If blnOk Then
    cmdSelect_SQL_5OPE_XX = Null
    arrYPDCOPE0(I).PDCOPESTA2 = "="
    arrYPDCOPE0(I).PDCOPEOPEN = xZCHGOPE0.CHGOPEDOS
End If

End Function

Public Sub cmdSelect_SQL_9_Terme(lPCI As String)
Dim wMOUVEMCOM As String
Dim xSQL As String

wMOUVEMCOM = lPCI & newYPDCPOS0.PDCPOSDEV & "EUR"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH where MOUVEMCOM = '" & wMOUVEMCOM & "'" _
     & " and MOUVEMDTR > 1080000  and MOUVEMDTR < 1089999" _
     & " order by MOUVEMDTR , MOUVEMPIE , MOUVEMECR"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    xYPDCMVT0.PDCMVTMTD = rsSab("MOUVEMMON")
    If rsSab("MOUVEMOPE") = "TER" Then
        newYPDCPOS0.PDCPOSTERD = newYPDCPOS0.PDCPOSTERD + xYPDCMVT0.PDCMVTMTD
    Else
        If rsSab("MOUVEMNUM") <> 13 Then
            newYPDCPOS0.PDCPOSSWPD = newYPDCPOS0.PDCPOSSWPD + xYPDCMVT0.PDCMVTMTD
        End If
    End If
    rsSab.MoveNext

Loop
'_____________________________________________________________________________________
wMOUVEMCOM = lPCI & "EUR" & newYPDCPOS0.PDCPOSDEV
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH where MOUVEMCOM = '" & wMOUVEMCOM & "'" _
     & " and MOUVEMDTR > 1080000  and MOUVEMDTR < 1089999" _
     & " order by MOUVEMDTR , MOUVEMPIE , MOUVEMECR"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    xYPDCMVT0.PDCMVTMTD = rsSab("MOUVEMMON")
    If rsSab("MOUVEMOPE") = "TER" Then
        newYPDCPOS0.PDCPOSTERE = newYPDCPOS0.PDCPOSTERE + xYPDCMVT0.PDCMVTMTD
    Else
        If rsSab("MOUVEMNUM") <> 13 Then
            newYPDCPOS0.PDCPOSSWPE = newYPDCPOS0.PDCPOSSWPE + xYPDCMVT0.PDCMVTMTD
        End If
    End If
    rsSab.MoveNext

Loop

End Sub

Public Sub fgTermeEch_DisplayTotal_1(kDev As Integer, lCHGOPEOPE As String)
Dim wColor As Long, wDev As String
Dim curEUR_DB As Currency, curEUR_CR As Currency
Dim curDEV_DB As Currency, curDEV_CR As Currency
Dim curEUR_B As Currency, curEUR_HB As Currency
Dim curDEV_B As Currency, curDEV_HB As Currency
Dim curEur As Currency, curDev As Currency
Dim curEur_T As Currency, curDev_T As Currency
Dim curEur_SWP_HB As Currency, curDev_SWP_HB As Currency
Dim wText As String, xSQL As String
Dim xCV As String, wCur As Currency
wDev = arrDev(kDev)
If lCHGOPEOPE <> "TOT" Then

    Select Case lCHGOPEOPE
        Case "TER": wColor = &HA0D0FF: wText = "terme"
                    curEUR_DB = arrTerme_DB(kDev).PDCPOSTERE
                    curEUR_CR = arrTerme_CR(kDev).PDCPOSTERE
                    curDEV_DB = arrTerme_DB(kDev).PDCPOSTERD
                    curDEV_CR = arrTerme_CR(kDev).PDCPOSTERD
        Case "SWP": wColor = &HA0D0FF: wText = "swap"
                    curEUR_DB = arrTerme_DB(kDev).PDCPOSSWPE
                    curEUR_CR = arrTerme_CR(kDev).PDCPOSSWPE
                    curDEV_DB = arrTerme_DB(kDev).PDCPOSSWPD
                    curDEV_CR = arrTerme_CR(kDev).PDCPOSSWPD
        'Case "TOT": wColor = &H80C0FF: wText = "TOTAL"
        '            curEUR_DB = arrTerme_DB(kdev).PDCPOSSWPE + arrTerme_DB(kdev).PDCPOSTERE
        '            curEUR_CR = arrTerme_CR(kdev).PDCPOSSWPE + arrTerme_CR(kdev).PDCPOSTERE
        '            curDEV_DB = arrTerme_DB(kdev).PDCPOSSWPD + arrTerme_DB(kdev).PDCPOSTERD
        '            curDEV_CR = arrTerme_CR(kdev).PDCPOSSWPD + arrTerme_CR(kdev).PDCPOSTERD
      End Select
      
    fgTermeEch.Rows = fgTermeEch.Rows + 1
    fgTermeEch.Row = fgTermeEch.Rows - 1
    fgTermeEch.Col = 3: fgTermeEch.Text = wText & " / " & wDev & "  ": fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 4: fgTermeEch.Text = "fixing " & arrTerme_DB(kDev).PDCPOSFIXT: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 8: fgTermeEch.Text = wDev: fgTermeEch.CellBackColor = wColor
    
    fgTermeEch.Col = 6: If curEUR_DB <> 0 Then fgTermeEch.Text = Format$(curEUR_DB, "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbRed: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 7: If curEUR_CR <> 0 Then fgTermeEch.Text = Format$(curEUR_CR, "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 9: If curDEV_DB <> 0 Then fgTermeEch.Text = Format$(curDEV_DB, "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbRed: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 10: If curDEV_CR <> 0 Then fgTermeEch.Text = Format$(curDEV_CR, "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
    
    wCur = curDEV_DB - curDEV_CR
    If arrTerme_DB(kDev).PDCPOSFIXT = 0 Then
        xCV = ""
    Else
        xCV = Format$(Round(Abs(wCur) / arrTerme_DB(kDev).PDCPOSFIXT, 2), "### ### ### ##0.00")
    End If
    fgTermeEch.Col = 11: fgTermeEch.Text = xCV: fgTermeEch.CellBackColor = &HA0D0FF
    fgTermeEch.CellForeColor = IIf(wCur < 0, vbRed, vbBlue)
'_________________________________________________________________________________
Else
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '382110" & wDev & "EUR'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then curDEV_B = rsSab("SOLDECEN") / 1000
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '382110EUR" & wDev & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then curEUR_B = rsSab("SOLDECEN") / 1000
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '933000" & wDev & "EUR'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then curDEV_HB = rsSab("SOLDECEN") / 1000
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '933000EUR" & wDev & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then curEUR_HB = rsSab("SOLDECEN") / 1000
    
    fgTermeEch.Rows = fgTermeEch.Rows + 1
    fgTermeEch.Row = fgTermeEch.Rows - 1
    wColor = &HFFFFD0: wText = "933000"
    
    fgTermeEch.Col = 4: fgTermeEch.Text = wText & " / " & wDev: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 8: fgTermeEch.Text = wDev: fgTermeEch.CellBackColor = wColor
    
    fgTermeEch.Col = 6: If curEUR_HB > 0 Then fgTermeEch.Text = Format$(curEUR_HB, "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbRed: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 7: If curEUR_HB < 0 Then fgTermeEch.Text = Format$(Abs(curEUR_HB), "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 9: If curDEV_HB > 0 Then fgTermeEch.Text = Format$(curDEV_HB, "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbRed: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 10: If curDEV_HB < 0 Then fgTermeEch.Text = Format$(Abs(curDEV_HB), "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
    
    If arrTerme_DB(kDev).PDCPOSFIXT = 0 Then
        xCV = ""
    Else
        xCV = Format$(Round(Abs(curDEV_HB) / arrTerme_DB(kDev).PDCPOSFIXT, 2), "### ### ### ##0.00")
    End If
    fgTermeEch.Col = 11: fgTermeEch.Text = xCV: fgTermeEch.CellBackColor = &HA0D0FF
    fgTermeEch.CellForeColor = IIf(curDEV_HB < 0, vbRed, vbBlue)
'_________________________________________________________________________________

    fgTermeEch.Rows = fgTermeEch.Rows + 1
    fgTermeEch.Row = fgTermeEch.Rows - 1
    wColor = &HFFFFC0: wText = "382110"
    
    fgTermeEch.Col = 4: fgTermeEch.Text = wText & " / " & wDev: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 8: fgTermeEch.Text = wDev: fgTermeEch.CellBackColor = wColor
    
    fgTermeEch.Col = 6: If curEUR_B > 0 Then fgTermeEch.Text = Format$(curEUR_B, "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbRed: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 7: If curEUR_B < 0 Then fgTermeEch.Text = Format$(Abs(curEUR_B), "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 9: If curDEV_B > 0 Then fgTermeEch.Text = Format$(curDEV_B, "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbRed: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 10: If curDEV_B < 0 Then fgTermeEch.Text = Format$(Abs(curDEV_B), "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
    
    If arrTerme_DB(kDev).PDCPOSFIXT = 0 Then
        xCV = ""
    Else
        xCV = Format$(Round(Abs(curDEV_B) / arrTerme_DB(kDev).PDCPOSFIXT, 2), "### ### ### ##0.00")
    End If
    fgTermeEch.Col = 11: fgTermeEch.Text = xCV: fgTermeEch.CellBackColor = &HA0D0FF
    fgTermeEch.CellForeColor = IIf(curDEV_B < 0, vbRed, vbBlue)
'_________________________________________________________________________________

    curEur = curEUR_B + curEUR_HB
    curDev = curDEV_B + curDEV_HB
    
    fgTermeEch.Rows = fgTermeEch.Rows + 1
    fgTermeEch.Row = fgTermeEch.Rows - 1
    wText = "solde"
    wColor = &HFFFF90

    fgTermeEch.Col = 4: fgTermeEch.Text = wText & " / " & wDev: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 8: fgTermeEch.Text = wDev: fgTermeEch.CellBackColor = wColor
    
    fgTermeEch.Col = 6: If curEur > 0 Then fgTermeEch.Text = Format$(curEur, "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbRed: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 7: If curEur < 0 Then fgTermeEch.Text = Format$(Abs(curEur), "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 9: If curDev > 0 Then fgTermeEch.Text = Format$(curDev, "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbRed: fgTermeEch.CellBackColor = wColor
    fgTermeEch.Col = 10: If curDev < 0 Then fgTermeEch.Text = Format$(Abs(curDev), "### ### ### ##0.00")
    fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
    
    If arrTerme_DB(kDev).PDCPOSFIXT = 0 Then
        xCV = ""
    Else
        xCV = Format$(Round(Abs(curDev) / arrTerme_DB(kDev).PDCPOSFIXT, 2), "### ### ### ##0.00")
    End If
    fgTermeEch.Col = 11: fgTermeEch.Text = xCV: fgTermeEch.CellBackColor = &HA0D0FF
    fgTermeEch.CellForeColor = IIf(curDev < 0, vbRed, vbBlue)
'_________________________________________________________________________________

    curEur_T = arrTerme_DB(kDev).PDCPOSTERE - arrTerme_CR(kDev).PDCPOSTERE + arrSWP_Dev(kDev).PDCMVTMTE
    curDev_T = arrTerme_DB(kDev).PDCPOSTERD - arrTerme_CR(kDev).PDCPOSTERD + arrSWP_Dev(kDev).PDCMVTMTD
    If curEur = curEur_T And curDev = curDev_T Then
        wColor = &HC0FFC0
        fgTermeEch.Col = 2: fgTermeEch.Text = "Contrôle OPE": fgTermeEch.CellBackColor = wColor
        fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
        fgTermeEch.Col = 3: fgTermeEch.Text = "= COMPTA   ": fgTermeEch.CellBackColor = wColor
        fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
   Else
        wColor = &HFFC0FF
        fgTermeEch.Rows = fgTermeEch.Rows + 1
        fgTermeEch.Row = fgTermeEch.Rows - 1
        wText = "contrôle"
        
        fgTermeEch.Col = 4: fgTermeEch.Text = wText & " / " & wDev: fgTermeEch.CellBackColor = wColor
        fgTermeEch.Col = 8: fgTermeEch.Text = wDev: fgTermeEch.CellBackColor = wColor
        
        fgTermeEch.Col = 6: If curEur_T > 0 Then fgTermeEch.Text = Format$(curEur_T, "### ### ### ##0.00")
        fgTermeEch.CellForeColor = vbRed: fgTermeEch.CellBackColor = wColor
        fgTermeEch.Col = 7: If curEur_T < 0 Then fgTermeEch.Text = Format$(Abs(curEur_T), "### ### ### ##0.00")
        fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
        fgTermeEch.Col = 9: If curDev_T > 0 Then fgTermeEch.Text = Format$(curDev_T, "### ### ### ##0.00")
        fgTermeEch.CellForeColor = vbRed: fgTermeEch.CellBackColor = wColor
        fgTermeEch.Col = 10: If curDev_T < 0 Then fgTermeEch.Text = Format$(Abs(curDev_T), "### ### ### ##0.00")
        fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
        fgTermeEch.Col = 11: fgTermeEch.Text = "!": fgTermeEch.CellBackColor = wColor
'_________________________________________________________________________________
    End If

    If arrSWP_Dev(kDev).PDCMVTMTD <> 0 Then
        curEur_SWP_HB = arrSWP_Dev(kDev).PDCMVTMTE
        curDev_SWP_HB = arrSWP_Dev(kDev).PDCMVTMTD

        wColor = &HFFFF00
        fgTermeEch.Rows = fgTermeEch.Rows + 1
        fgTermeEch.Row = fgTermeEch.Rows - 1
        fgTermeEch.Col = 2: fgTermeEch.Text = "dont SWP": fgTermeEch.CellBackColor = wColor
        fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
        fgTermeEch.Col = 3: fgTermeEch.Text = "comptant": fgTermeEch.CellBackColor = wColor
        fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
       
        wText = "HB"
        
        fgTermeEch.Col = 4: fgTermeEch.Text = wText & " / " & wDev: fgTermeEch.CellBackColor = wColor
        fgTermeEch.Col = 8: fgTermeEch.Text = wDev: fgTermeEch.CellBackColor = wColor
        
        fgTermeEch.Col = 6: If curEur_SWP_HB > 0 Then fgTermeEch.Text = Format$(curEur_SWP_HB, "### ### ### ##0.00")
        fgTermeEch.CellForeColor = vbRed: fgTermeEch.CellBackColor = wColor
        fgTermeEch.Col = 7: If curEur_SWP_HB < 0 Then fgTermeEch.Text = Format$(Abs(curEur_SWP_HB), "### ### ### ##0.00")
        fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
        fgTermeEch.Col = 9: If curDev_SWP_HB > 0 Then fgTermeEch.Text = Format$(curDev_SWP_HB, "### ### ### ##0.00")
        fgTermeEch.CellForeColor = vbRed: fgTermeEch.CellBackColor = wColor
        fgTermeEch.Col = 10: If curDev_SWP_HB < 0 Then fgTermeEch.Text = Format$(Abs(curDev_SWP_HB), "### ### ### ##0.00")
        fgTermeEch.CellForeColor = vbBlue: fgTermeEch.CellBackColor = wColor
        fgTermeEch.Col = 11: fgTermeEch.Text = "!": fgTermeEch.CellBackColor = wColor
'_________________________________________________________________________________
    End If

End If


End Sub



Public Sub fraPDCOPE_Control_DEV_Certain()
Dim xCur1 As Currency, xCur2 As Currency, xDbl As Double
xCur1 = num_CDec(txtPDCOPEMTD1)
xDbl = num_CDec(txtPDCOPETAUX)
xCur2 = 0

If cboPDCOPEDEV1 = "EUR" Then
    lblPDCOPETAUX = "EUR / " & cboPDCOPEDEV2
    xCur2 = Round(xCur1 * xDbl, 2)
Else
    If cboPDCOPEDEV2 = "EUR" Then
        lblPDCOPETAUX = "EUR / " & cboPDCOPEDEV1
        If xDbl <> 0 Then xCur2 = Round(xCur1 / xDbl, 2)
    Else
        If cboPDCOPEDEV1 = "GBP" Then
            lblPDCOPETAUX = "GBP / " & cboPDCOPEDEV2
            xCur2 = Round(xCur1 * xDbl, 2)
        Else
            If cboPDCOPEDEV2 = "GBP" Then
                lblPDCOPETAUX = "GBP / " & cboPDCOPEDEV1
                If xDbl <> 0 Then xCur2 = Round(xCur1 / xDbl, 2)
            Else
                If cboPDCOPEDEV1 = "USD" Then
                    lblPDCOPETAUX = "USD / " & cboPDCOPEDEV2
                    xCur2 = Round(xCur1 * xDbl, 2)
                Else
                    If cboPDCOPEDEV2 = "USD" Then
                        lblPDCOPETAUX = "USD / " & cboPDCOPEDEV1
                        If xDbl <> 0 Then xCur2 = Round(xCur1 / xDbl, 2)
                    Else
                        lblPDCOPETAUX = "" & cboPDCOPEDEV1 & " / " & cboPDCOPEDEV2
                        xCur2 = Round(xCur1 * xDbl, 2)
                    End If
                End If
            End If
        End If
    End If
End If
xYPDCOPE0.PDCOPEMTD1 = xCur1
xYPDCOPE0.PDCOPEMTD2 = xCur2
libPDCOPEMTD2 = Format$(xCur2, "### ### ### ###.00")
If cboPDCOPEOPEC = "SWP" Then fraPDCOPE_Control_DEV_Certain_SWP
End Sub

Public Sub fraPDCOPE_Control_DEV_Certain_SWP()
Dim xCur1 As Currency, xCur2 As Currency, xDbl As Double
xCur1 = num_CDec(txtPDCOPEMTD1)
xDbl = num_CDec(txtPDCOPEFIXING)
xCur2 = 0

If cboPDCOPEDEV1 = "EUR" Then
    xCur2 = Round(xCur1 * xDbl, 2)
Else
    If cboPDCOPEDEV2 = "EUR" Then
        If xDbl <> 0 Then xCur2 = Round(xCur1 / xDbl, 2)
    Else
        If cboPDCOPEDEV1 = "GBP" Then
            xCur2 = Round(xCur1 * xDbl, 2)
        Else
            If cboPDCOPEDEV2 = "GBP" Then
                If xDbl <> 0 Then xCur2 = Round(xCur1 / xDbl, 2)
            Else
                If cboPDCOPEDEV1 = "USD" Then
                    xCur2 = Round(xCur1 * xDbl, 2)
                Else
                    If cboPDCOPEDEV2 = "USD" Then
                        If xDbl <> 0 Then xCur2 = Round(xCur1 / xDbl, 2)
                    Else
                        xCur2 = Round(xCur1 * xDbl, 2)
                    End If
                End If
            End If
        End If
    End If
End If
txtPDCOPEVTXT = "cours à Terme : " & Format$(xDbl, "##0.000000") & vbCrLf & cboPDCOPEDEV2 & " à terme    : " & Format$(xCur2, "### ### ### ###.00")
txtPDCOPEVTXT.ForeColor = libPDCOPEMTD2.ForeColor

End Sub

Public Sub cmdSelect_SQL_5Suspens(lPDCPOSDTR As String)
'_________________________________________________________________________
' contrepassation automatique SUSPENS
'________________________________________________________________________
Call arrYPDCMVT0_SQL("where PDCMVTOPEC like 'XX%' and PDCMVTSTA2 = ' ' and PDCMVTDVA <= '" & lPDCPOSDTR & "'")
For I = 1 To arrYPDCMVT0_Nb
    oldYPDCMVT0 = arrYPDCMVT0(I)

    newYPDCMVT0 = oldYPDCMVT0
    newYPDCMVT0.PDCMVTSTA2 = "="
    memoYPDCMVT0 = newYPDCMVT0
    newYPDCMVT0.PDCMVTDTR = lPDCPOSDTR
    newYPDCMVT0.PDCMVTDVA = lPDCPOSDTR
    newYPDCMVT0.PDCMVTECR = 0
    newYPDCMVT0.PDCMVTSTA = "-"
    newYPDCMVT0.PDCMVTCPT = Replace(newYPDCMVT0.PDCMVTCPT, "suspens", "annul")
    newYPDCMVT0.PDCMVTMTE = -newYPDCMVT0.PDCMVTMTE
    newYPDCMVT0.PDCMVTMTD = -newYPDCMVT0.PDCMVTMTD
   
    Call cmdSuspens_Transaction(constUpdate)

Next I
End Sub

Public Function cmdSelect_SQL_5M_Trilog(lTxt As String) As Long
Dim K As Integer, K2 As Integer
Dim X As String
X = Trim(lTxt)
K2 = Len(X)
For K = K2 To 2 Step -1
    If Mid$(X, K, 1) = " " Then Exit For
Next K
X = Mid$(X, K + 1, K2 - K)
If IsNumeric(X) Then
    cmdSelect_SQL_5M_Trilog = Val(X)
Else
    cmdSelect_SQL_5M_Trilog = 0
End If
End Function

Public Sub cmdSelect_SQL_5Control(lPDCPOSDTR_1 As String)
Dim V, I As Long, Nb As Long, K As Long
Dim xSQL As String
Dim nbPDCMVTOPEC_X As Long, arrCHGOPEDOS_Ok() As Boolean
Dim blnOk As Boolean

xSQL = " where CHGOPECRE =" & lPDCPOSDTR_1 - 19000000 & " and   CHGOPEDE2 <> '   '" _
     & " and CHGOPEVAL = 'O' and CHGOPEANN = ' ' and CHGOPESSE <> 'GU'"
Call arrZCHGOPE0_SQL(xSQL)

ReDim arrCHGOPEDOS_Ok(arrZCHGOPE0_Nb + 1)

For I = 1 To arrZCHGOPE0_Nb
    arrCHGOPEDOS_Ok(I) = False
Next I
'___________________________________________________________________________
Call arrYPDCMVT0_SQL("where PDCMVTDTR = '" & lPDCPOSDTR_1 & "'")
nbPDCMVTOPEC_X = 0
For K = 1 To arrYPDCMVT0_Nb
    Select Case arrYPDCMVT0(K).PDCMVTOPEC
        Case "-TR", "CDE", "CDI", "CPT", "PPD", "REM", "RPC", "SWP", "TER", "XXC", "XXT"
        Case Else: nbPDCMVTOPEC_X = nbPDCMVTOPEC_X + 1
    End Select
    If arrYPDCMVT0(K).PDCMVTOPEC = "CPT" Then
        blnOk = False
        For I = 1 To arrZCHGOPE0_Nb
            If arrYPDCMVT0(K).PDCMVTOPEN = arrZCHGOPE0(I).CHGOPEDOS Then
                arrCHGOPEDOS_Ok(I) = True
                blnOk = True
                Exit For
            End If
        Next I
        If Not blnOk Then
               xSQL = "select * from " & paramIBM_Library_SAB & ".ZCHGOPE0 where CHGOPESER = '" & arrYPDCMVT0(K).PDCMVTSER _
                    & "' and CHGOPESSE = '" & arrYPDCMVT0(K).PDCMVTSSE _
                    & "' and CHGOPEOPE = '" & arrYPDCMVT0(K).PDCMVTOPEC _
                    & "' and CHGOPEDOS = " & arrYPDCMVT0(K).PDCMVTOPEN
                Set rsSab = cnsab.Execute(xSQL)
                
                If rsSab.EOF Then
                    xYPDCLOG0.PDCLOGNAT = "5X?"
                    xYPDCLOG0.PDCLOGTXT = "Mvt Comptable sans opération : " & arrYPDCMVT0(K).PDCMVTOPEC & " " & arrYPDCMVT0(K).PDCMVTOPEN _
                               & " " & arrYPDCMVT0(K).PDCMVTDEV & " " & Format$(arrYPDCMVT0(K).PDCMVTMTE, "### ### ### ##0.00")
                    Call YPDCLOG0_AddItem
                End If
        End If
    End If
Next K

If nbPDCMVTOPEC_X <> 0 Then
    xYPDCLOG0.PDCLOGNAT = "5X!"
    xYPDCLOG0.PDCLOGTXT = nbPDCMVTOPEC_X & " Ecritures Comptables manuelles (*Z1....)"
    Call YPDCLOG0_AddItem
End If

For I = 1 To arrZCHGOPE0_Nb
    If Not arrCHGOPEDOS_Ok(I) Then
        xYPDCLOG0.PDCLOGNAT = "5X?"
        xYPDCLOG0.PDCLOGTXT = "opé sans compta : " & arrZCHGOPE0(I).CHGOPEOPE & " " & arrZCHGOPE0(I).CHGOPEDOS _
                            & " " & arrZCHGOPE0(I).CHGOPEDE1 & " " & Format$(arrZCHGOPE0(I).CHGOPEMO1, "### ### ### ##0.00") & " / " _
                            & " " & arrZCHGOPE0(I).CHGOPEDE2 & " " & Format$(arrZCHGOPE0(I).CHGOPEMO2, "### ### ### ##0.00")

        Call YPDCLOG0_AddItem
    End If
Next I

'___________________________________________________________________________



End Sub

Public Sub cmdSelect_SQL_xls_Init()
Dim xSQL As String, wAMJ_FDM As String

txtSelect_Sheet_xls = "RECAPEUR"
If paramEnvironnement = constProduction Then
    txtSelect_File_xls = "\\DOCSRV2013\_GROUPS\GDC-BOTC\BOTC 1\arbitrage\arbitrage ter.xlsm"
Else
    txtSelect_File_xls = "C:\Temp\PDC\arbitrage.xlsm"
End If

xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCMAIL where PDCMAILDTR = 1 order by  PDCMAILSEQ"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
        V = rsYPDCMAIL_GetBuffer(rsSab, xYPDCMAIL)
        If xYPDCMAIL.PDCMAILSEQ = 0 Then BOTC_File_xls = Trim(xYPDCMAIL.PDCMAILTXT): txtSelect_File_xls = BOTC_File_xls
        If xYPDCMAIL.PDCMAILSEQ = 1 Then BOTC_Sheet_xls = Trim(xYPDCMAIL.PDCMAILTXT): txtSelect_Sheet_xls = BOTC_Sheet_xls
    rsSab.MoveNext

Loop

fraSelect_Options_xls.Visible = True
blnControl = False
'Call DTPicker_Set(txtSelect_AMJ_xls, YBIATAB0_DATE_CPT_J)
Call DTPicker_Control(txtSelect_AMJ_xls, mAMJ_xls)

cmdSelect_Ok_xls.Visible = False
libSelect_Report_xls = ""
libSelect_Report_xls.Visible = False

mAMJ_JP0 = mAMJ_xls
wAMJ_FDM = dateFinDeMois(mAMJ_xls)
Call DTPicker_Set(txtSelect_AMJ_HB_xls, wAMJ_FDM)
If wAMJ_FDM = mAMJ_xls Then

    mAMJ_JP0 = DateComptableJP0(mAMJ_xls)

'V = rsYBIATAB0_Read("DATE", "CAL", "M", wAMJ_FDM)
'If wAMJ_FDM = YBIATAB0_DATE_CPT_J And wAMJ_FDM = mAMJ_xls Then
    chkSelect_HB_xls = "1"
    MsgBox "Fin de mois SAB : " & dateImp(wAMJ_FDM) & vbCrLf & vbCrLf & "=> exclure les mouvements hors-bilan", vbInformation, frmElp_Caption & " : Contrôle de la position de change"
End If
blnControl = True

End Sub


Public Sub cmdSelect_SQL_5OPE_Report()
Dim V, K As Integer, xSQL As String
blnControl = False
Call DTPicker_Set(txtSelect_AMJ, YBIATAB0_DATE_CPT_JP1)

chkSelect_Terme = "0"
chkSelect_HB = "0"
chkSelect_Suspens_Out.Value = chkSelect_Suspens_Out_xls.Value
chkSelect_HB.Value = chkSelect_HB_xls.Value
blnControl = False

Call cmdSelect_SQL_1

For K = 1 To arrYPDCPOS0_Nb
    xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCPOS0 " _
         & " where PDCPOSDTR = '" & YBIATAB0_DATE_CPT_JP0 & "' and PDCPOSDEV = '" & arrYPDCPOS0(K).PDCPOSDEV & "'"
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        V = rsYPDCPOS0_GetBuffer(rsSab, oldYPDCPOS0)
        If oldYPDCPOS0.PDCPOSPNL <> arrPPJ(K) Then
            '$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            '_________________________________________________________________________________
            V = cnSAB_Transaction("BeginTrans")
            'If Not IsNull(V) Then GoTo Error_MsgBox
            newYPDCPOS0 = oldYPDCPOS0
            newYPDCPOS0.PDCPOSPNL = arrPPJ(K)
            V = sqlYPDCPOS0_Update(newYPDCPOS0, oldYPDCPOS0)
    
            '________________________________________________________________________________
    
    
            If Not IsNull(V) Then
                V = cnSAB_Transaction("Rollback")
            Else
                V = cnSAB_Transaction("Commit")
            End If
                        
            '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        End If
    End If
Next K
blnControl = True
End Sub

Public Sub cmdPDCOPE_Report_Annulation()
Dim xSQL As String, xWhere As String
Dim I As Long
fraReport.ForeColor = vbYellow
lblReport_Comment.ForeColor = vbYellow
libReport = "Annulation de l'opération reportée, réf :  " & dateImp10(oldYPDCOPE0.PDCOPEDTR) & " - " & oldYPDCOPE0.PDCOPEID & vbCrLf _
          & " saisie par " & Trim(oldYPDCOPE0.PDCOPEIUSR) & " le " & dateImp10(oldYPDCOPE0.PDCOPEIAMJ) & " - " & timeImp8(oldYPDCOPE0.PDCOPEIHMS) & vbCrLf _
          & "*****************************************************" & vbCrLf

xWhere = " where PDCOPEDTR > " & oldYPDCOPE0.PDCOPEDTR _
    & " and PDCOPEIUSR = '" & oldYPDCOPE0.PDCOPEIUSR & "'" _
    & " and PDCOPEIAMJ = " & oldYPDCOPE0.PDCOPEIAMJ _
    & " and PDCOPEIHMS = " & oldYPDCOPE0.PDCOPEIHMS

xSQL = "select count(*) as Tally  from " & paramIBM_Library_SABSPE_XXX & ".YPDCOPE0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)
ReDim selYPDCOPE0(rsSab("Tally") + 1)

xSQL = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCOPE0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)
selYPDCOPE0_Nb = 0
Do While Not rsSab.EOF
    V = rsYPDCOPE0_GetBuffer(rsSab, xYPDCOPE0)
    selYPDCOPE0_Nb = selYPDCOPE0_Nb + 1
    selYPDCOPE0(selYPDCOPE0_Nb) = xYPDCOPE0
    libReport = libReport & "+ Annulation du report, réf : " & dateImp10(xYPDCOPE0.PDCOPEDTR) & " - " & xYPDCOPE0.PDCOPEID & vbCrLf
    rsSab.MoveNext
Loop
          

fraReport.Visible = True
End Sub
Public Sub cmdSendMail_Report_Annulation()
Dim wSendMail As typeSendMail
Dim xHeader As String, xDétail As String, mbgColor As String
Dim wSubject As String
Dim K As Long, X As String
On Error Resume Next


wSendMail.Subject = "PDC " & dateImp10(oldYPDCOPE0.PDCOPEDTR) & " - " & oldYPDCOPE0.PDCOPEID & " : "

wSendMail.FromDisplayName = "@BIA_PDC"
wSendMail.RecipientDisplayName = "BIA_PDC"
wSubject = " ANNULATION du report automatique d'une opération non rapprochée"
wSendMail.Subject = wSendMail.Subject & wSubject
mbgColor = "bgcolor = #FFD0D0"
xHeader = "<TR>" _
         & "<TD bgcolor=#0090A0  width=200 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF>" _
         & dateImp10(oldYPDCOPE0.PDCOPEDTR) & " - " & oldYPDCOPE0.PDCOPEID & "</TD>" _
         & "<TD bgcolor=#0090A0  width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF>" _
         & "Service : " & oldYPDCOPE0.PDCOPESER & " " & oldYPDCOPE0.PDCOPESSE _
         & " - " & oldYPDCOPE0.PDCOPEOPEC & " " & oldYPDCOPE0.PDCOPEOPET & Format$(oldYPDCOPE0.PDCOPEOPEN, "### ###") & "</TD>" _
         & "<TD bgcolor=#0090A0  width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF>" _
         & wSubject & "</TD>" _
        & "</TR>"


If oldYPDCOPE0.PDCOPESENS = "A" Then
    cboPDCOPESENS.ListIndex = 0
Else
    cboPDCOPESENS.ListIndex = cboPDCOPESENS.ListCount - 1
End If
Call ZCLIEAN0_SQL(oldYPDCOPE0.PDCOPECLI)


xDétail = "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "Devise principale" & "</TD>" _
     & "<TD  " & mbgColor & "  width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Red _
     & "<B><div align=" & Asc34 & "right" & Asc34 & ">" & Format$(Abs(oldYPDCOPE0.PDCOPEMTD1), "### ### ### ##0.00") & " " & oldYPDCOPE0.PDCOPEDEV1 & "</div></TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Red _
     & cboPDCOPESENS.Text & "</B/TD>" _
     & "</TR>" _
     & "<TR>" _
     & "<TD  bgcolor=#BFFFFF width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "Devise secondaire" & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Blue _
     & "<B><div align=" & Asc34 & "right" & Asc34 & ">" & Format$(Abs(oldYPDCOPE0.PDCOPEMTD2), "### ### ### ##0.00") & " " & oldYPDCOPE0.PDCOPEDEV2 & "</div></TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Red _
     & "Cours " & lblPDCOPETAUX & " : " & Format$(oldYPDCOPE0.PDCOPETAUX, "### ##0.000 000") & "</B/TD>" _
     & "</TR>"
     
xDétail = xDétail _
     & "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "Contrepartie" & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:7.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & oldYPDCOPE0.PDCOPECLI & " " & mCLIENARA1 & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:9.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Blue _
     & "Date valeur : " & dateImp10(oldYPDCOPE0.PDCOPEDVA) & "</TD>" _
     & "</TR>" _
     & "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "Saisie par" & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:7.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & oldYPDCOPE0.PDCOPEIUSR & "  " & dateImp10(oldYPDCOPE0.PDCOPEIAMJ) & "  " & timeImp8(oldYPDCOPE0.PDCOPEIHMS) & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:7.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Magenta _
     & "." & oldYPDCOPE0.PDCOPEITXT & "</TD>" _
     & "</TR>" _
     & "<TR>" _
     & "<TD bgcolor=#BAFAFA  width=200 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & "mise à jour par" & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Black _
     & oldYPDCOPE0.PDCOPEVUSR & "  " & dateImp10(oldYPDCOPE0.PDCOPEVAMJ) & "  " & timeImp8(oldYPDCOPE0.PDCOPEVHMS) & "</TD>" _
     & "<TD " & mbgColor & " width=300 height=5><span style='font-size:7.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Magenta _
     & "." & oldYPDCOPE0.PDCOPEVTXT & "</TD>" _
     & "</TR>"


X = Replace(Trim(txtReport_Comment), vbCr, "<BR>")

wSendMail.Attachment = ""
wSendMail.Message = "<body bgcolor = #FF80FF>" _
                    & "<span style='font-size:10.0pt;font-family:Arial Unicode MS Unicode MS'>" & "<Font color = #404040>" _
                    & X _
                    & "<BR><BR>" & "<TABLE border = 1  width=800 height=5 cellpadding=3 >" _
                    & xHeader _
                    & xDétail _
                    & "</TABLE>"


wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail



End Sub



Public Function cmdSendMail_xls_Color(lColor As Long) As String
Dim xColor As String, X As String
xColor = Hex(lColor)
Select Case Len(xColor)
    Case 6:
    Case 2: xColor = "0000" & xColor
    Case 4: xColor = "00" & xColor
    Case 1: xColor = "00000" & xColor
    Case 3: xColor = "000" & xColor
End Select

cmdSendMail_xls_Color = " #" & Mid$(xColor, 5, 2) & Mid$(xColor, 3, 2) & Mid$(xColor, 1, 2)
End Function



Public Sub fraSuspens_Display_Log()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean '

On Error Resume Next
currentAction = "fraSuspens_Display_Log"

fgSuspens_Log.Visible = False
fgSuspens_Log.Rows = 1
fgSuspens_Log.FormatString = fgSuspens_Log_FormatString
X = "select * from " & paramIBM_Library_SABSPE_XXX & ".YPDCLOG0 " _
    & "where PDCLOGDTR = '" & xYPDCMVT0.PDCMVTDTR & "' and PDCLOGPIE = " & xYPDCMVT0.PDCMVTPIE & " and  PDCLOGECR = " & xYPDCMVT0.PDCMVTECR _
    & " and PDCLOGNAT like '4%'" _
    & " order by  PDCLOGUAMJ , PDCLOGUHMS,PDCLOGUSEQ"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    fgSuspens_Log.Rows = fgSuspens_Log.Rows + 1
    fgSuspens_Log.Row = fgSuspens_Log.Rows - 1
    fgSuspens_Log.Col = 0: fgSuspens_Log.Text = dateImp10_S(rsSab("PDCLOGUAMJ")) & " " & timeImp8(rsSab("PDCLOGUHMS"))
    fgSuspens_Log.Col = 1: fgSuspens_Log.Text = rsSab("PDCLOGUUSR")
    fgSuspens_Log.Col = 2: fgSuspens_Log.Text = rsSab("PDCLOGNAT")
    fgSuspens_Log.Col = 3: fgSuspens_Log.Text = rsSab("PDCLOGTXT")
    rsSab.MoveNext

Loop


Set rsSab = Nothing

fgSuspens_Log.Visible = True


End Sub

Public Function dateAdd_On(lDev As String, lNb As Integer, lAMJ As String) As String
Dim K As Integer, wNb As Integer, wJMA As String, wAmj As String
Dim blnOk As Boolean

wJMA = dateImp10_S(lAMJ)
Do
    wJMA = DateAdd("d", 1, wJMA)
    K = Weekday(wJMA)
    If K = 1 Or K = 7 Then
    Else
        wNb = wNb + 1
        wAmj = Mid$(wJMA, 7, 4) & Mid$(wJMA, 4, 2) & Mid$(wJMA, 1, 2)
        For K = 1 To arrDevF_Nb
    
            If arrDevF_AMJ(K) = wAmj Then
                If arrDevF_ISO(K) = lDev Then wNb = wNb - 1
            End If
        Next K
    End If
        

Loop Until wNb = lNb

dateAdd_On = Mid$(wJMA, 7, 4) & Mid$(wJMA, 4, 2) & Mid$(wJMA, 1, 2)
End Function
