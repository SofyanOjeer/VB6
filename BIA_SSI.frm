VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmBIA_SSI 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA_SSI"
   ClientHeight    =   12165
   ClientLeft      =   120
   ClientTop       =   420
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
   Icon            =   "BIA_SSI.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   12165
   ScaleWidth      =   16335
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
      Left            =   8970
      TabIndex        =   2
      Top             =   60
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   11640
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   16275
      _ExtentX        =   28707
      _ExtentY        =   20532
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
      TabPicture(0)   =   "BIA_SSI.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Paramétrage"
      TabPicture(1)   =   "BIA_SSI.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraParam"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Réservé informatique"
      TabPicture(2)   =   "BIA_SSI.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraParam 
         Height          =   11100
         Left            =   -74925
         TabIndex        =   104
         Top             =   390
         Width           =   16035
         Begin TabDlg.SSTab SSTabParam 
            Height          =   10830
            Left            =   75
            TabIndex        =   149
            Top             =   210
            Width           =   15900
            _ExtentX        =   28046
            _ExtentY        =   19103
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "paramétrage envoi automatique de courriels"
            TabPicture(0)   =   "BIA_SSI.frx":035E
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lstParam_SSIMELUIDX"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lstParam_SSIMELNAT"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "fraParam_SSIMELUIDX"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "paramétrage général BIA_SSI"
            TabPicture(1)   =   "BIA_SSI.frx":037A
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fraParam_K2"
            Tab(1).Control(1)=   "lstParam_K1"
            Tab(1).Control(2)=   "lstParam_K2"
            Tab(1).ControlCount=   3
            TabCaption(2)   =   "Tab 2"
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
            Begin VB.Frame fraParam_SSIMELUIDX 
               BackColor       =   &H00F0FFF0&
               Height          =   10140
               Left            =   7395
               TabIndex        =   160
               Top             =   465
               Width           =   7800
               Begin VB.TextBox txtParam_SSIMELINFO 
                  BackColor       =   &H00D0FFD0&
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
                  Height          =   660
                  Left            =   525
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  TabIndex        =   170
                  Top             =   1485
                  Width           =   6945
               End
               Begin VB.CommandButton cmdParam_SSIMELUIDX_New 
                  BackColor       =   &H000080FF&
                  Caption         =   "Ajouter"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   3975
                  Style           =   1  'Graphical
                  TabIndex        =   169
                  Top             =   2220
                  Width           =   1320
               End
               Begin VB.CommandButton cmdParam_SSIMELUIDX_Quit 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Abandonner"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   690
                  Style           =   1  'Graphical
                  TabIndex        =   165
                  Top             =   2190
                  Width           =   1125
               End
               Begin VB.CommandButton cmdParam_SSIMELUIDX_Update 
                  BackColor       =   &H0080FF80&
                  Caption         =   "Enregistrer"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   5805
                  Style           =   1  'Graphical
                  TabIndex        =   164
                  Top             =   2205
                  Width           =   1110
               End
               Begin VB.TextBox txtParam_SSIMELUIDX 
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   2430
                  TabIndex        =   163
                  Top             =   255
                  Width           =   3705
               End
               Begin VB.TextBox txtParam_SSIMELUNOM 
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
                  Height          =   465
                  Left            =   510
                  MaxLength       =   64
                  MultiLine       =   -1  'True
                  TabIndex        =   162
                  Top             =   900
                  Width           =   6945
               End
               Begin VB.ListBox lstParam_SSIMELUNOM 
                  BackColor       =   &H00E0FFE0&
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00008000&
                  Height          =   6810
                  Left            =   225
                  Style           =   1  'Checkbox
                  TabIndex        =   161
                  Top             =   2715
                  Width           =   7290
               End
               Begin VB.Label lblParam_SSIMELUIDX 
                  BackColor       =   &H00F0FFF0&
                  Caption         =   "Application.fonction"
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
                  Left            =   270
                  TabIndex        =   166
                  Top             =   360
                  Width           =   1995
               End
            End
            Begin VB.ListBox lstParam_SSIMELNAT 
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1260
               Left            =   315
               TabIndex        =   168
               Top             =   660
               Width           =   6990
            End
            Begin VB.ListBox lstParam_SSIMELUIDX 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7740
               Left            =   285
               TabIndex        =   167
               Top             =   2340
               Width           =   15435
            End
            Begin VB.ListBox lstParam_K2 
               BackColor       =   &H00E0FFFF&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   6780
               Left            =   -74655
               TabIndex        =   159
               Top             =   3495
               Width           =   4890
            End
            Begin VB.ListBox lstParam_K1 
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2700
               Left            =   -74640
               TabIndex        =   158
               Top             =   555
               Width           =   4785
            End
            Begin VB.Frame fraParam_K2 
               BackColor       =   &H00E0FFFF&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   8835
               Left            =   -67485
               TabIndex        =   150
               Top             =   840
               Visible         =   0   'False
               Width           =   6585
               Begin VB.CommandButton cmdParam_K2_Quit 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Abandonner"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   510
                  Style           =   1  'Graphical
                  TabIndex        =   156
                  Top             =   8200
                  Width           =   990
               End
               Begin VB.CommandButton cmdParam_K2_Update 
                  BackColor       =   &H0080FF80&
                  Caption         =   "Enregistrer"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   3645
                  Style           =   1  'Graphical
                  TabIndex        =   155
                  Top             =   8200
                  Width           =   900
               End
               Begin VB.CommandButton cmdParam_K2_Add 
                  BackColor       =   &H000080FF&
                  Caption         =   "Ajouter"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   2160
                  Style           =   1  'Graphical
                  TabIndex        =   154
                  Top             =   8200
                  Width           =   900
               End
               Begin VB.CommandButton cmdParam_K2_Delete 
                  BackColor       =   &H00FF80FF&
                  Caption         =   "Supprimer"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   5055
                  Style           =   1  'Graphical
                  TabIndex        =   153
                  Top             =   8200
                  Width           =   900
               End
               Begin VB.TextBox txtParam_K2_Code 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   350
                  Left            =   2535
                  MaxLength       =   10
                  TabIndex        =   152
                  Top             =   255
                  Width           =   2040
               End
               Begin VB.TextBox txtParam_K2_Info 
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   5000
                  Left            =   630
                  MultiLine       =   -1  'True
                  TabIndex        =   151
                  Text            =   "BIA_SSI.frx":0396
                  Top             =   1740
                  Width           =   5355
               End
               Begin VB.Label lblParam_K2_Code 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Code (10car)"
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
                  Left            =   285
                  TabIndex        =   157
                  Top             =   255
                  Width           =   1290
               End
            End
         End
      End
      Begin VB.Frame fraSelect 
         BackColor       =   &H00E0E0E0&
         Height          =   11055
         Left            =   -105
         TabIndex        =   4
         Top             =   495
         Width           =   16290
         Begin VB.CommandButton cmdSSIUSR_New 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Ajouter un utilisateur"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   12300
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   780
            Visible         =   0   'False
            Width           =   1200
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
            Left            =   14640
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   810
            Width           =   1335
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
            Left            =   12315
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   285
            Width           =   3690
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
            Width           =   12075
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   9585
            Left            =   180
            TabIndex        =   8
            Top             =   1365
            Width           =   16095
            _ExtentX        =   28390
            _ExtentY        =   16907
            _Version        =   393216
            Cols            =   11
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16777215
            ForeColor       =   8388608
            BackColorFixed  =   12632064
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   $"BIA_SSI.frx":03CB
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
      Begin TabDlg.SSTab SSTab2 
         Height          =   11175
         Left            =   -74895
         TabIndex        =   9
         Top             =   375
         Width           =   16125
         _ExtentX        =   28443
         _ExtentY        =   19711
         _Version        =   393216
         Tabs            =   6
         Tab             =   1
         TabsPerRow      =   6
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
         TabCaption(0)   =   "fraOptions"
         TabPicture(0)   =   "BIA_SSI.frx":04FC
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fraSelect_Options_4"
         Tab(0).Control(1)=   "lstW"
         Tab(0).Control(2)=   "fraSelect_Options_J"
         Tab(0).Control(3)=   "fraSelect_Options_1"
         Tab(0).Control(4)=   "txtFg"
         Tab(0).Control(5)=   "txtRTF"
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "fraDetail"
         TabPicture(1)   =   "BIA_SSI.frx":0518
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "fraDetail"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "fraProfil"
         TabPicture(2)   =   "BIA_SSI.frx":0534
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraYSSIDIV0"
         Tab(2).Control(1)=   "fraProfil_Update_DIV"
         Tab(2).Control(2)=   "fraYSSIDOM0"
         Tab(2).Control(3)=   "fraProfil"
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "fraCompteH"
         TabPicture(3)   =   "BIA_SSI.frx":0550
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "fraCompteH"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Tab 4"
         TabPicture(4)   =   "BIA_SSI.frx":056C
         Tab(4).ControlEnabled=   0   'False
         Tab(4).ControlCount=   0
         TabCaption(5)   =   "Tab 5"
         TabPicture(5)   =   "BIA_SSI.frx":0588
         Tab(5).ControlEnabled=   0   'False
         Tab(5).ControlCount=   0
         Begin VB.Frame fraYSSIDIV0 
            BackColor       =   &H0080FFFF&
            Caption         =   "DIV : mise à jour manuelle des comptes"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3720
            Left            =   -68310
            TabIndex        =   121
            Top             =   1395
            Visible         =   0   'False
            Width           =   8025
            Begin VB.CheckBox chkSSIDIVPRFK 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0080FFFF&
               Caption         =   "Ce compte est clos dans le domaine"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   135
               TabIndex        =   124
               Top             =   1140
               Width           =   3645
            End
            Begin VB.CommandButton cmdYSSIDIV0_Quit 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Abandonner"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   6675
               Style           =   1  'Graphical
               TabIndex        =   130
               Top             =   1080
               Width           =   1200
            End
            Begin VB.TextBox txtSSIDIVINFO 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1020
               Left            =   225
               MaxLength       =   1024
               MultiLine       =   -1  'True
               TabIndex        =   126
               Top             =   2610
               Width           =   7515
            End
            Begin VB.TextBox txtSSIDIVUIDX 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2025
               TabIndex        =   123
               Top             =   585
               Width           =   2970
            End
            Begin VB.TextBox txtSSIDIVUNOM 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2025
               MaxLength       =   20
               TabIndex        =   125
               Top             =   1815
               Width           =   5775
            End
            Begin VB.CommandButton cmdYSSIDIV0_Update 
               BackColor       =   &H0000FF00&
               Caption         =   "Enregistrer"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   6735
               Style           =   1  'Graphical
               TabIndex        =   122
               Top             =   270
               Width           =   1110
            End
            Begin VB.Label lblSSIDIVINFO 
               BackColor       =   &H0080FFFF&
               Caption         =   "Informations complémentaires"
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
               Left            =   225
               TabIndex        =   129
               Top             =   2250
               Width           =   3405
            End
            Begin VB.Label lblSSIDIVUIDX 
               BackColor       =   &H0080FFFF&
               Caption         =   "Id du compte"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   225
               TabIndex        =   128
               Top             =   570
               Width           =   1380
            End
            Begin VB.Label lblSSIDIVUNOM 
               BackColor       =   &H0080FFFF&
               Caption         =   "Intitulé du compte"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   225
               TabIndex        =   127
               Top             =   1800
               Width           =   1575
            End
         End
         Begin VB.Frame fraProfil_Update_DIV 
            BackColor       =   &H00FFC0FF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2715
            Left            =   -74745
            TabIndex        =   108
            Top             =   2295
            Visible         =   0   'False
            Width           =   8150
            Begin VB.TextBox txtProfil_PRFX_DIV 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   5355
               MaxLength       =   20
               TabIndex        =   110
               Top             =   270
               Width           =   2730
            End
            Begin VB.TextBox txtProfil_UNOM_DIV 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1710
               MaxLength       =   32
               TabIndex        =   113
               Top             =   1350
               Width           =   5445
            End
            Begin VB.TextBox txtProfil_DIDK_DIV 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1755
               MaxLength       =   5
               TabIndex        =   109
               Top             =   240
               Width           =   1260
            End
            Begin VB.TextBox txtProfil_UIDD_DIV 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1725
               TabIndex        =   111
               Top             =   810
               Width           =   1260
            End
            Begin VB.TextBox txtProfil_IDX_DIV 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   5370
               MaxLength       =   20
               TabIndex        =   112
               Top             =   765
               Width           =   2730
            End
            Begin VB.TextBox txtProfil_Info_DIV 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   750
               Left            =   1710
               MaxLength       =   1024
               TabIndex        =   115
               Top             =   1830
               Width           =   5520
            End
            Begin VB.Label lblProfil_PRFX_DIV 
               BackColor       =   &H0000FFFF&
               Caption         =   "Code du sous-domaine"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3210
               TabIndex        =   131
               Top             =   255
               Width           =   1965
            End
            Begin VB.Label lblProfil_UNOM_DIV 
               BackColor       =   &H00FFC0FF&
               Caption         =   "Nom du profil "
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               TabIndex        =   119
               Top             =   1380
               Width           =   1470
            End
            Begin VB.Label lblProfil_DIDK_DIV 
               BackColor       =   &H0000FFFF&
               Caption         =   "Sous-domaine"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   90
               TabIndex        =   118
               Top             =   255
               Width           =   1470
            End
            Begin VB.Label lblProfil_UIDD_DIV 
               BackColor       =   &H0000FF00&
               Caption         =   "N° du profil DIV"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   135
               TabIndex        =   117
               Top             =   855
               Width           =   1605
            End
            Begin VB.Label lblProfil_IDX_DIV 
               BackColor       =   &H0000FF00&
               Caption         =   "Code BIA_SSI"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3240
               TabIndex        =   116
               Top             =   825
               Width           =   1965
            End
            Begin VB.Label lblProfil_Info_DIV 
               BackColor       =   &H00FFC0FF&
               Caption         =   "Commentaire"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   180
               TabIndex        =   114
               Top             =   1920
               Width           =   1260
            End
         End
         Begin VB.Frame fraSelect_Options_4 
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
            Left            =   -71715
            TabIndex        =   91
            Top             =   4545
            Visible         =   0   'False
            Width           =   12075
            Begin VB.ComboBox cboSelect_Options_4_SSIDOMSTAK 
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
               Left            =   10170
               Style           =   2  'Dropdown List
               TabIndex        =   103
               Top             =   465
               Width           =   1515
            End
            Begin VB.CheckBox chkSelect_Options_4_SSIDOMDIDX 
               BackColor       =   &H00F0FFFF&
               Caption         =   "inclure l'historique"
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
               Left            =   2865
               TabIndex        =   100
               Top             =   345
               Value           =   1  'Checked
               Width           =   1935
            End
            Begin VB.TextBox txtSelect_Options_4_SSIDOMUIDD 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6855
               TabIndex        =   99
               Top             =   825
               Width           =   2115
            End
            Begin VB.ComboBox cboSelect_Options_4_SSIDOMNAT 
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
               Left            =   1245
               Style           =   2  'Dropdown List
               TabIndex        =   97
               Top             =   810
               Width           =   3405
            End
            Begin VB.ComboBox cboSelect_Options_4_SSIDOMDIDX 
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
               Left            =   1260
               Style           =   2  'Dropdown List
               TabIndex        =   93
               Top             =   300
               Width           =   1350
            End
            Begin VB.TextBox txtSelect_Options_4_SSIDOMUIDX 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6855
               TabIndex        =   92
               Top             =   300
               Width           =   2100
            End
            Begin VB.Label lblSelect_Options_4_SSIDOMSTAK 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Actif"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   9540
               TabIndex        =   102
               Top             =   495
               Width           =   540
            End
            Begin VB.Label lblSelect_Options_4_SSIDOMUIDD 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Identifiant N"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   5295
               TabIndex        =   98
               Top             =   825
               Width           =   1035
            End
            Begin VB.Label lblSelect_Options_4_SSIDOMNAT 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Nature"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   195
               TabIndex        =   96
               Top             =   810
               Width           =   720
            End
            Begin VB.Label lblSelect_Options_4_SSIDOMDIDX 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Domaine"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   195
               TabIndex        =   95
               Top             =   345
               Width           =   1005
            End
            Begin VB.Label lblSelect_Options_4_SSIDOMUIDX 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Identifiant X"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   5295
               TabIndex        =   94
               Top             =   345
               Width           =   1395
            End
         End
         Begin VB.ListBox lstW 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8700
            Left            =   -66090
            TabIndex        =   83
            Top             =   4350
            Visible         =   0   'False
            Width           =   5700
         End
         Begin VB.Frame fraSelect_Options_J 
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
            Left            =   -71805
            TabIndex        =   76
            Top             =   2955
            Visible         =   0   'False
            Width           =   12075
            Begin VB.TextBox txtSelect_Options_J_SSIUSRUIDX 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6945
               TabIndex        =   82
               Top             =   540
               Width           =   1695
            End
            Begin VB.ComboBox cboSelect_Options_J_SSIDOMDIDX 
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
               Left            =   3870
               Style           =   2  'Dropdown List
               TabIndex        =   79
               Top             =   525
               Width           =   1725
            End
            Begin MSComCtl2.DTPicker txtSelect_Options_J_SSITXTYMAJ 
               Height          =   300
               Left            =   975
               TabIndex        =   77
               Top             =   510
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   16777215
               CalendarForeColor=   0
               CalendarTitleBackColor=   8421504
               CalendarTitleForeColor=   16777215
               CalendarTrailingForeColor=   12632256
               CheckBox        =   -1  'True
               CustomFormat    =   "dd  MM yyy"
               Format          =   87162883
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label libSelect_Options_J_SSIUSRUIDX 
               BackColor       =   &H00F0FFFF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6975
               TabIndex        =   84
               Top             =   900
               Width           =   1620
            End
            Begin VB.Label lblSelect_Options_J_SSIUSRUIDX 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Nom"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6270
               TabIndex        =   81
               Top             =   555
               Width           =   570
            End
            Begin VB.Label lblSelect_Options_J_SSIDOMDIDX 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Domaine"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2835
               TabIndex        =   80
               Top             =   540
               Width           =   1005
            End
            Begin VB.Label lblSelect_Options_J_SSITXTYMAJ 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Date"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   210
               TabIndex        =   78
               Top             =   555
               Width           =   1005
            End
         End
         Begin VB.Frame fraCompteH 
            BackColor       =   &H00F0F0F0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   9570
            Left            =   -73200
            TabIndex        =   67
            Top             =   720
            Width           =   8265
            Begin VB.CommandButton cmdCompteH_Update 
               BackColor       =   &H0000FF00&
               Caption         =   "Vu Historique du compte"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   5415
               Style           =   1  'Graphical
               TabIndex        =   73
               Top             =   8700
               Visible         =   0   'False
               Width           =   2355
            End
            Begin VB.CommandButton cmdCompteH_Quit 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Abandonner"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   72
               Top             =   8745
               Width           =   1200
            End
            Begin VB.TextBox txtCompteH_SSITXTINFO 
               BackColor       =   &H00D0FFD0&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1395
               Left            =   120
               MaxLength       =   1024
               MultiLine       =   -1  'True
               TabIndex        =   70
               Top             =   7140
               Width           =   8000
            End
            Begin MSFlexGridLib.MSFlexGrid fgCompteH 
               Height          =   6120
               Left            =   120
               TabIndex        =   68
               Top             =   465
               Width           =   8000
               _ExtentX        =   14102
               _ExtentY        =   10795
               _Version        =   393216
               Cols            =   4
               FixedCols       =   0
               RowHeightMin    =   300
               BackColor       =   16448250
               ForeColor       =   4210752
               BackColorFixed  =   12632256
               ForeColorFixed  =   0
               BackColorBkg    =   16448250
               GridColor       =   10526720
               GridColorFixed  =   10526720
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               FormatString    =   $"BIA_SSI.frx":05A4
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
            Begin VB.Label lblCompteH_SSITXTINFO 
               BackColor       =   &H00FFFFFF&
               Caption         =   "saisir un commentaire :"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C000C0&
               Height          =   330
               Left            =   135
               TabIndex        =   71
               Top             =   6765
               Width           =   2700
            End
         End
         Begin VB.Frame fraSelect_Options_1 
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
            Left            =   -72345
            TabIndex        =   57
            Top             =   1005
            Visible         =   0   'False
            Width           =   12075
            Begin VB.TextBox txtSelect_Options_1_SSIDOMUIDX 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6735
               TabIndex        =   148
               Top             =   825
               Width           =   2505
            End
            Begin VB.ComboBox cboSelect_Options_1_SSIDOMDIDX 
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
               Left            =   3795
               Style           =   2  'Dropdown List
               TabIndex        =   75
               Top             =   510
               Width           =   1725
            End
            Begin VB.ComboBox cboSelect_Options_1_SSIUSRSTAK 
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
               Left            =   10395
               Style           =   2  'Dropdown List
               TabIndex        =   61
               Top             =   510
               Width           =   1515
            End
            Begin VB.ComboBox cboSelect_Options_1_SSIDOMPRFX 
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
               Left            =   6735
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   285
               Width           =   2505
            End
            Begin VB.TextBox txtSelect_Options_1_SSIUSRUIDX 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   945
               TabIndex        =   59
               Top             =   510
               Width           =   1695
            End
            Begin VB.Label lblSelect_Options_1_SSIDOMUIDX 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Compte"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   5805
               TabIndex        =   147
               Top             =   855
               Width           =   840
            End
            Begin VB.Label lblSelect_Options_1_SSIDOMDIDX 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Domaine"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2835
               TabIndex        =   74
               Top             =   550
               Width           =   1005
            End
            Begin VB.Label lblSelect_Options_1_SSIDOMPRFX 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Profil"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   5790
               TabIndex        =   63
               Top             =   330
               Width           =   840
            End
            Begin VB.Label lblSelect_Options_1_SSIUSRSTAK 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Actif"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   9720
               TabIndex        =   62
               Top             =   550
               Width           =   540
            End
            Begin VB.Label lblSelect_Options_1_SSIUSRUIDX 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Nom"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   150
               TabIndex        =   58
               Top             =   550
               Width           =   1005
            End
         End
         Begin VB.Frame fraYSSIDOM0 
            BackColor       =   &H00E0FFE0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   7365
            Left            =   -65430
            TabIndex        =   37
            Top             =   3045
            Width           =   8145
            Begin VB.ComboBox cboSSIDOMUNIT 
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
               Left            =   1250
               Style           =   2  'Dropdown List
               TabIndex        =   146
               Top             =   1320
               Width           =   2505
            End
            Begin VB.CheckBox chkSSIDOMDECH 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0FFE0&
               Caption         =   "date de fin du contrat"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   3540
               TabIndex        =   65
               Top             =   855
               Value           =   1  'Checked
               Width           =   2145
            End
            Begin VB.CommandButton cmdCompte_All 
               BackColor       =   &H00FF00FF&
               Caption         =   "+"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2415
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   4710
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.ComboBox cboSSIDOMPRFK 
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
               Left            =   6600
               Style           =   2  'Dropdown List
               TabIndex        =   47
               Top             =   2175
               Width           =   1200
            End
            Begin VB.TextBox txtSSIDOMPRFX 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1250
               TabIndex        =   45
               Top             =   1830
               Width           =   2070
            End
            Begin VB.ComboBox cboSSIDOMSTAK 
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
               Left            =   1250
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   765
               Width           =   2010
            End
            Begin VB.TextBox txtSSIDOMTXT 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1395
               Left            =   270
               MaxLength       =   1024
               MultiLine       =   -1  'True
               TabIndex        =   38
               Top             =   2865
               Width           =   7590
            End
            Begin MSComCtl2.DTPicker txtSSIDOMDECH 
               Height          =   300
               Left            =   6510
               TabIndex        =   41
               Top             =   900
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
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
               Format          =   87162883
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSFlexGridLib.MSFlexGrid fgCompte 
               Height          =   2115
               Left            =   195
               TabIndex        =   54
               Top             =   4665
               Width           =   7695
               _ExtentX        =   13573
               _ExtentY        =   3731
               _Version        =   393216
               Cols            =   3
               FixedCols       =   0
               RowHeightMin    =   300
               BackColor       =   16448250
               ForeColor       =   4210752
               BackColorFixed  =   8454016
               ForeColorFixed  =   0
               BackColorBkg    =   16448250
               GridColor       =   10526720
               GridColorFixed  =   10526720
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               FormatString    =   $"BIA_SSI.frx":0668
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
            Begin VB.Label lblSSIDOMUNIT 
               BackColor       =   &H00E0FFE0&
               Caption         =   "Service"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   240
               TabIndex        =   145
               Top             =   1380
               Width           =   1005
            End
            Begin VB.Label libSSIDOMPRFD 
               BackColor       =   &H00E0FFE0&
               Caption         =   "date du contrôle"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1245
               TabIndex        =   52
               Top             =   2175
               Width           =   2085
            End
            Begin VB.Label libSSIDOMUIDD 
               BackColor       =   &H00E0FFE0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "identification du compte"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1250
               TabIndex        =   51
               Top             =   255
               Width           =   5340
            End
            Begin VB.Label libSSIDOMPRFX 
               BackColor       =   &H00E0FFE0&
               Caption         =   "profil"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3600
               TabIndex        =   50
               Top             =   1830
               Width           =   3990
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblSSIDOMYUSR 
               BackColor       =   &H00E0FFE0&
               Caption         =   "mise à jour"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   210
               TabIndex        =   49
               Top             =   6915
               Width           =   7455
            End
            Begin VB.Label lblSSIDOMPRFD 
               BackColor       =   &H00E0FFE0&
               Caption         =   "Contrôlé le"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   240
               TabIndex        =   48
               Top             =   2175
               Width           =   1095
            End
            Begin VB.Label lblSSIDOMPRFK 
               BackColor       =   &H00E0FFE0&
               Caption         =   "habilitations conformes au profil ?"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3600
               TabIndex        =   46
               Top             =   2175
               Width           =   3045
            End
            Begin VB.Label lblSSIDOMPRTFX 
               BackColor       =   &H00E0FFE0&
               Caption         =   "Profil"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   44
               Top             =   1830
               Width           =   900
            End
            Begin VB.Label lblSSIDOMUIDD 
               BackColor       =   &H00E0FFE0&
               Caption         =   "Compte"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   240
               TabIndex        =   43
               Top             =   300
               Width           =   1005
            End
            Begin VB.Label lblSSIDOMTXT 
               BackColor       =   &H00E0FFE0&
               Caption         =   "Commentaire"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   255
               TabIndex        =   42
               Top             =   2595
               Width           =   1275
            End
            Begin VB.Label lblSSIDOMSTAK 
               BackColor       =   &H00E0FFE0&
               Caption         =   "Actif"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   240
               TabIndex        =   39
               Top             =   870
               Width           =   1005
            End
         End
         Begin VB.Frame fraProfil 
            BackColor       =   &H00F0FFF0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   9570
            Left            =   -74460
            TabIndex        =   22
            Top             =   300
            Visible         =   0   'False
            Width           =   8265
            Begin VB.CommandButton cmdProfil_Change 
               BackColor       =   &H0000FFFF&
               Caption         =   "Modifier ce profil"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   6210
               Style           =   1  'Graphical
               TabIndex        =   132
               Top             =   8415
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CommandButton cmdProfil_Update_DIV 
               BackColor       =   &H0000FFFF&
               Caption         =   "gestion manuelle du compte"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   4590
               Style           =   1  'Graphical
               TabIndex        =   120
               Top             =   8715
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CommandButton cmdProfil_Excel 
               BackColor       =   &H0080FFFF&
               Caption         =   "Edition des profils => Excel"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   2955
               Style           =   1  'Graphical
               TabIndex        =   107
               Top             =   8775
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CheckBox chkProfil_DOM 
               BackColor       =   &H00F0FFF0&
               Caption         =   "tous les profils"
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
               Left            =   6150
               TabIndex        =   101
               Top             =   570
               Width           =   1680
            End
            Begin VB.CommandButton cmdProfil_Print 
               BackColor       =   &H0080FFFF&
               Caption         =   "Edition des profils => word"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   1500
               Style           =   1  'Graphical
               TabIndex        =   90
               Top             =   8730
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CommandButton cmdCompte_Val 
               BackColor       =   &H00FF00FF&
               Caption         =   "Valider ce profil non conforme"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   4875
               Style           =   1  'Graphical
               TabIndex        =   69
               Top             =   8355
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CommandButton cmdProfil_Histo 
               BackColor       =   &H00FFFF00&
               Caption         =   "Afficher l'historique"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   1815
               Style           =   1  'Graphical
               TabIndex        =   66
               Top             =   8730
               Width           =   1200
            End
            Begin VB.CommandButton cmdProfil_Delete 
               BackColor       =   &H000000FF&
               Caption         =   "Détacher"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   3270
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   8730
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.Frame fraProfil_Update 
               BackColor       =   &H00F0FFF0&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2640
               Left            =   75
               TabIndex        =   29
               Top             =   5715
               Visible         =   0   'False
               Width           =   8150
               Begin VB.TextBox txtProfil_UPTEXT 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   750
                  Left            =   2790
                  MaxLength       =   50
                  TabIndex        =   36
                  Top             =   1680
                  Width           =   5010
               End
               Begin VB.CommandButton cmdProfil_Display 
                  BackColor       =   &H0080C0FF&
                  Caption         =   "Afficher le compte"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   615
                  Left            =   6750
                  Style           =   1  'Graphical
                  TabIndex        =   34
                  Top             =   300
                  Width           =   1110
               End
               Begin VB.TextBox txtProfil_IDX 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2790
                  MaxLength       =   20
                  TabIndex        =   33
                  Top             =   1050
                  Width           =   2730
               End
               Begin VB.TextBox txtProfil_UIDD 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2760
                  TabIndex        =   31
                  Top             =   300
                  Width           =   1965
               End
               Begin VB.Label lblProfil_UPTEXT 
                  BackColor       =   &H00F0FFF0&
                  Caption         =   "Commentaire"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   150
                  TabIndex        =   35
                  Top             =   1845
                  Width           =   2685
               End
               Begin VB.Label lblProfil_IDX 
                  BackColor       =   &H00F0FFF0&
                  Caption         =   "Nom du profil à créer"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   165
                  TabIndex        =   32
                  Top             =   1065
                  Width           =   2685
               End
               Begin VB.Label lblProfil_UIDD 
                  BackColor       =   &H00F0FFF0&
                  Caption         =   "Id utilisateur dans le domaine"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   200
                  TabIndex        =   30
                  Top             =   285
                  Width           =   2685
               End
            End
            Begin VB.ComboBox cboProfil_DOM 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   3690
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   510
               Width           =   2220
            End
            Begin VB.CommandButton cmdProfil_Update 
               BackColor       =   &H0000FF00&
               Caption         =   "Enregistrer"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   6480
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   8715
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CommandButton cmdProfil_New 
               BackColor       =   &H00FF00FF&
               Caption         =   "Créer un profil"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   4830
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   8730
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CommandButton cmdProfil_Quit 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Abandonner"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   150
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   8730
               Width           =   1200
            End
            Begin MSFlexGridLib.MSFlexGrid fgProfil 
               Height          =   7425
               Left            =   45
               TabIndex        =   23
               Top             =   1020
               Width           =   8145
               _ExtentX        =   14367
               _ExtentY        =   13097
               _Version        =   393216
               Cols            =   5
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16448250
               ForeColor       =   4210752
               BackColorFixed  =   12632064
               ForeColorFixed  =   0
               BackColorBkg    =   16448250
               GridColor       =   10526720
               GridColorFixed  =   10526720
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               FormatString    =   $"BIA_SSI.frx":0701
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
            Begin VB.Label lblProfil_DOM 
               BackColor       =   &H00F0FFF0&
               Caption         =   "Domaine"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2385
               TabIndex        =   28
               Top             =   510
               Width           =   1005
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
            Left            =   -73995
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   16
            Text            =   "BIA_SSI.frx":07C1
            Top             =   6615
            Visible         =   0   'False
            Width           =   5775
         End
         Begin VB.Frame fraDetail 
            BackColor       =   &H00F0FFF0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   9630
            Left            =   1335
            TabIndex        =   10
            Top             =   465
            Visible         =   0   'False
            Width           =   11400
            Begin VB.CommandButton cmdSSIUSR_Delete 
               BackColor       =   &H008080FF&
               Caption         =   "Supprimer"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   8325
               Style           =   1  'Graphical
               TabIndex        =   105
               Top             =   375
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.Frame fraDetail_Update 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   3660
               Left            =   60
               TabIndex        =   18
               Top             =   345
               Width           =   9300
               Begin VB.Frame fraDetail_Update_SRV 
                  BackColor       =   &H00C0C0C0&
                  BorderStyle     =   0  'None
                  Caption         =   "Frame1"
                  Height          =   465
                  Left            =   1500
                  TabIndex        =   138
                  Top             =   375
                  Visible         =   0   'False
                  Width           =   7305
                  Begin VB.TextBox txtSSIUSRUNIT_N 
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
                     Left            =   1200
                     MaxLength       =   2
                     TabIndex        =   142
                     Top             =   105
                     Width           =   675
                  End
                  Begin VB.TextBox txtSSIUSRUNIT_X 
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
                     Left            =   3210
                     MaxLength       =   20
                     TabIndex        =   140
                     Top             =   105
                     Width           =   2040
                  End
                  Begin VB.Label lblSSIUSRUNIT_N 
                     BackColor       =   &H00F0FFF0&
                     Caption         =   "Numéro"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Left            =   285
                     TabIndex        =   141
                     Top             =   165
                     Width           =   795
                  End
                  Begin VB.Label lbllSSIUSRUNIT_X 
                     BackColor       =   &H00F0FFF0&
                     Caption         =   "Code"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Left            =   2175
                     TabIndex        =   139
                     Top             =   165
                     Width           =   795
                  End
               End
               Begin VB.Frame fraDetail_Update_STAK 
                  BackColor       =   &H00C0C0C0&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   990
                  Left            =   300
                  TabIndex        =   133
                  Top             =   540
                  Width           =   8880
                  Begin VB.ComboBox cboSSIUSRUNIT 
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
                     Left            =   1230
                     Style           =   2  'Dropdown List
                     TabIndex        =   144
                     Top             =   615
                     Width           =   2505
                  End
                  Begin VB.CheckBox chkSSIUSRDECH 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00F0FFF0&
                     Caption         =   "date de fin du contrat"
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
                     Left            =   3825
                     TabIndex        =   136
                     Top             =   270
                     Value           =   1  'Checked
                     Width           =   2460
                  End
                  Begin VB.ComboBox cboSSIUSRSTAK 
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
                     Left            =   1245
                     Style           =   2  'Dropdown List
                     TabIndex        =   135
                     Top             =   240
                     Width           =   2460
                  End
                  Begin MSComCtl2.DTPicker txtSSIUSRDECH 
                     Height          =   300
                     Left            =   7035
                     TabIndex        =   137
                     Top             =   285
                     Width           =   1395
                     _ExtentX        =   2461
                     _ExtentY        =   529
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Calibri"
                        Size            =   9.75
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
                     Format          =   87162883
                     CurrentDate     =   38699.44875
                     MaxDate         =   401768
                     MinDate         =   36526.4425347222
                  End
                  Begin VB.Label lblSSIUSRUNIT 
                     BackColor       =   &H00F0FFF0&
                     Caption         =   "Service"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Left            =   195
                     TabIndex        =   143
                     Top             =   690
                     Width           =   1005
                  End
                  Begin VB.Label lblSSIUSRSTAK 
                     BackColor       =   &H00F0FFF0&
                     Caption         =   "Actif"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Left            =   210
                     TabIndex        =   134
                     Top             =   300
                     Width           =   1005
                  End
               End
               Begin VB.Frame fraDetail_Update_PRF 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   525
                  Left            =   285
                  TabIndex        =   85
                  Top             =   1560
                  Width           =   8880
                  Begin VB.ComboBox cboSSIUSRPRFK 
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
                     Left            =   7035
                     Style           =   2  'Dropdown List
                     TabIndex        =   89
                     Top             =   60
                     Width           =   1455
                  End
                  Begin VB.ComboBox cboSSIUSRPRFX 
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
                     Left            =   1245
                     Style           =   2  'Dropdown List
                     TabIndex        =   87
                     Top             =   60
                     Width           =   2505
                  End
                  Begin VB.Label lblSSIUSRPRFK 
                     BackColor       =   &H00F0FFF0&
                     Caption         =   "habilitations conformes au modèle ?"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Left            =   3825
                     TabIndex        =   88
                     Top             =   100
                     Width           =   3075
                  End
                  Begin VB.Label lblSSIUSRPRFX 
                     BackColor       =   &H00F0FFF0&
                     Caption         =   "modèle BIA"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Left            =   210
                     TabIndex        =   86
                     Top             =   100
                     Width           =   1005
                  End
               End
               Begin VB.TextBox txtSSIUSRUIDX 
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
                  Left            =   1380
                  MaxLength       =   20
                  TabIndex        =   21
                  Top             =   120
                  Width           =   4410
               End
               Begin VB.TextBox txtSSIUSRTXT 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1350
                  Left            =   315
                  MaxLength       =   1024
                  MultiLine       =   -1  'True
                  TabIndex        =   19
                  Top             =   2160
                  Width           =   8640
               End
               Begin VB.Label lblSSIUSRUIDX 
                  BackColor       =   &H00F0FFF0&
                  Caption         =   "Nom"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   405
                  TabIndex        =   20
                  Top             =   180
                  Width           =   1005
               End
            End
            Begin VB.CommandButton cmdSSIUSR_Quit 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Abandonner"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   9800
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   3105
               Width           =   1200
            End
            Begin VB.CommandButton cmdSSIUSR_Update 
               BackColor       =   &H0080FF80&
               Caption         =   "Enregistrer"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   9800
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   345
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CommandButton cmdSSIUSR_Histo 
               BackColor       =   &H00FFFF00&
               Caption         =   "Afficher l'historique"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   9800
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   1185
               Width           =   1200
            End
            Begin VB.CommandButton cmdSSIUSR_PRF 
               BackColor       =   &H0000FF00&
               Caption         =   "Ajouter un profil"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   9780
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   2175
               Visible         =   0   'False
               Width           =   1200
            End
            Begin MSFlexGridLib.MSFlexGrid fgDetail 
               Height          =   4890
               Left            =   105
               TabIndex        =   15
               Top             =   4080
               Width           =   11175
               _ExtentX        =   19711
               _ExtentY        =   8625
               _Version        =   393216
               Cols            =   10
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16448250
               ForeColor       =   16711680
               BackColorFixed  =   12632064
               ForeColorFixed  =   16777215
               BackColorBkg    =   16448250
               GridColor       =   0
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               FormatString    =   $"BIA_SSI.frx":07C9
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
            Begin VB.Label lblSSIUSRYUSR 
               BackColor       =   &H00C0FFC0&
               Caption         =   "mise à jour"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   180
               TabIndex        =   64
               Top             =   9135
               Width           =   11025
            End
         End
         Begin RichTextLib.RichTextBox txtRTF 
            Height          =   10200
            Left            =   -74835
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   810
            Visible         =   0   'False
            Width           =   7995
            _ExtentX        =   14102
            _ExtentY        =   17992
            _Version        =   393217
            BackColor       =   16448250
            HideSelection   =   0   'False
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"BIA_SSI.frx":08AB
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
      Left            =   15795
      Picture         =   "BIA_SSI.frx":0927
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   15
      Width           =   500
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   1185
      TabIndex        =   106
      Top             =   -15
      Visible         =   0   'False
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
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
      Begin VB.Menu mnuPrint_Excel 
         Caption         =   "Excel"
      End
      Begin VB.Menu mnuPrint_Mail 
         Caption         =   "Mail"
      End
      Begin VB.Menu mnuPrint_RTF 
         Caption         =   "Enregistrer le détail (.doc)"
      End
      Begin VB.Menu mnuPrint_RTF_USR 
         Caption         =   "Fiche utilisateur (.doc)"
      End
   End
End
Attribute VB_Name = "frmBIA_SSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'txtRTF :
'  1 : noir
'  2 : rouge foncé
'  3 : vert
'  4 : marron clair
'  5 : bleu foncé
'  6 : violet
'  7 : bleu vert
'  8 : gris
'  9 : argenté
' 10 : rouge
' 11 : vert
' 12 : jaune
' 13 : bleue
' 14 : fuschia
' 15 : cyan
' 16 : blanc


Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim arrHab(19) As Boolean, blnHab2_SécuritéPhysique As Boolean
Dim blnAuto As Boolean, blnError As Boolean
Dim cmdSelect_SQL_K As String
Dim rsSab_X As New ADODB.Recordset
Dim mMail_Destinataires As String

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

Dim fgProfil_FormatString As String, fgProfil_K As Integer
Dim fgProfil_RowDisplay As Integer, fgProfil_RowClick As Integer, fgProfil_ColClick As Integer
Dim fgProfil_ColorClick As Long, fgProfil_ColorDisplay As Long
Dim fgProfil_Sort1 As Integer, fgProfil_Sort2 As Integer
Dim fgProfil_SortAD As Integer, fgProfil_Sort1_Old As Integer
Dim fgProfil_arrIndex As Integer
Dim blnfgProfil_DisplayLine As Boolean


Dim fgCompte_FormatString As String, fgCompte_K As Integer
Dim fgCompte_RowDisplay As Integer, fgCompte_RowClick As Integer, fgCompte_ColClick As Integer
Dim fgCompte_ColorClick As Long, fgCompte_ColorDisplay As Long
Dim fgCompte_Sort1 As Integer, fgCompte_Sort2 As Integer
Dim fgCompte_SortAD As Integer, fgCompte_Sort1_Old As Integer
Dim fgCompte_arrIndex As Integer
Dim blnfgCompte_DisplayLine As Boolean

Dim fgCompteH_FormatString As String, fgCompteH_K As Integer
Dim fgCompteH_RowDisplay As Integer, fgCompteH_RowClick As Integer, fgCompteH_ColClick As Integer
Dim fgCompteH_ColorClick As Long, fgCompteH_ColorDisplay As Long
Dim fgCompteH_Sort1 As Integer, fgCompteH_Sort2 As Integer
Dim fgCompteH_SortAD As Integer, fgCompteH_Sort1_Old As Integer
Dim fgCompteH_arrIndex As Integer
Dim blnfgCompteH_DisplayLine As Boolean
'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long

Dim HeightOfLine As Long, LinesOfText As Long

Dim txtRTF_prtForeColor_Header As Long, VB_RTF_Modèle As String, VB_RTF_Suite As String
Dim txtRTF_Replace
Dim rsSab_2 As New ADODB.Recordset, rsSab_3 As New ADODB.Recordset

Dim xYSSIUSR0 As typeYSSIUSR0, oldYSSIUSR0 As typeYSSIUSR0, newYSSIUSR0 As typeYSSIUSR0
Dim hYSSIUSR0 As typeYSSIUSR0
Dim xYSSIDOM0 As typeYSSIDOM0, oldYSSIDOM0 As typeYSSIDOM0, newYSSIDOM0 As typeYSSIDOM0
Dim hYSSIDOM0 As typeYSSIDOM0
Dim arrYSSIDOM0_BIA() As typeYSSIDOM0, arrYSSIDOM0_BIA_Nb As Integer
Dim arrYSSIDOM0_Usr() As typeYSSIDOM0, arrYSSIDOM0_Usr_Nb As Integer

Dim xYSSITXT0 As typeYSSITXT0, oldYSSITXT0_XXX As typeYSSITXT0, newYSSITXT0 As typeYSSITXT0
Dim oldYSSITXT0_USR As typeYSSITXT0, oldYSSITXT0_DOM As typeYSSITXT0, oldYSSITXT0_Histo As typeYSSITXT0
Dim newYSSITXT0_JRN As typeYSSITXT0

Dim hYSSITXT0 As typeYSSITXT0, paramYSSITXT0 As typeYSSITXT0
Dim xYSSIIBM0 As typeYSSIIBM0, oldYSSIIBM0 As typeYSSIIBM0, newYSSIIBM0 As typeYSSIIBM0
Dim hYSSIIBM0 As typeYSSIIBM0, prfYSSIIBM0 As typeYSSIIBM0
Dim oldYSSIIBMH As typeYSSIIBM0, newYSSIIBMH As typeYSSIIBM0

Dim usrYSSIIBM0 As typeYSSIIBM0, usrRTF(30) As String, arrRTF(30) As String

Dim mYSSIUSR0_Update As String, mYSSIDOM0_Update As String, mYSSITXT0_Update As String
Dim mYSSIIBM0_Update As String, mYSSIIBMH_Update As String, mYSSITXT0_JRN_Update As String
Dim mSSIUSRNAT As String
Dim mYSSIUSR0_Update_CMD As String, mYSSIDOM0_Update_CMD As String
Dim mYSSIDOM0_Update_CMD_2 As String

Dim arrSSIUSRSTAK() As String, arrSSIUSRPRFX() As String, arrSSIUSRPRFK() As String
Dim arrSSIUSRSTAK_UB As Integer, arrSSIUSRPRFK_UB As Integer, arrSSIUSRPRFX_UB As Integer

Dim arrSSIDOMDIDX() As String, arrSSIDOMDIDX_UB As Integer, mSSIDOMDIDX As String

Dim blnSelect_Options_1_SSIDOMPRFX As Boolean

Dim blnSelect_SQL_3_Rupture As Boolean

Dim mYSSISAA0_Update As String, mYSSISAAH_Update As String
Dim xYSSISAA0 As typeYSSISAA0, oldYSSISAA0 As typeYSSISAA0, newYSSISAA0 As typeYSSISAA0
Dim usrYSSISAA0 As typeYSSISAA0, prfYSSISAA0 As typeYSSISAA0
Dim xYSSISAAH As typeYSSISAA0, oldYSSISAAH As typeYSSISAA0, newYSSISAAH As typeYSSISAA0

Dim arrSSIDOMPRFX_D() As String, arrSSIDOMPRFX_P() As String, arrSSIDOMPRFX_Nb As Long
Dim arrSAA_UNIT_Code() As String, arrSAA_UNIT_Lib() As String, arrSAA_UNIT_Nb As Long
Dim arrSAA_App_Code() As String, arrSAA_App_Nb As Integer, arrSAA_App_K As Integer
Dim arrSAA_Function_Code() As String, arrSAA_Function_Nb As Integer, arrSAA_Function_K As Integer
Dim arrSAA_Profil_Code() As String, arrSAA_Profil_Lib() As String, arrSAA_Profil_Nb As Long
Dim arrSAA_Function() As String, blnSAA_Function() As Boolean

Dim mFile As String, mSSIDOMUIDD_Unique As Long
Dim arrCtl_Nb(50) As Long, arrCtl_Lib(50) As String, arrCtl_K As Integer
Dim mBIA_SSI_Archives As String
Dim mImport_Nb As Long, mImport_In As Long, mImport_New As Long, mImport_Update As Long, mImport_Ok As Long, mImport_Ann As Long
Dim mImport_PRFD As Long, mImport_PRFH As Long

Dim cnSIDE_DB As New ADODB.Connection, rsSIDE_DB As New ADODB.Recordset, rsSIDE_Loop As New ADODB.Recordset
Dim rsSIDE_X As New ADODB.Recordset


Dim mYSSISAB0_Update As String, mYSSISABH_Update As String
Dim xYSSISAB0 As typeYSSISAB0, oldYSSISAB0 As typeYSSISAB0, newYSSISAB0 As typeYSSISAB0
Dim usrYSSISAB0 As typeYSSISAB0, prfYSSISAB0 As typeYSSISAB0
Dim xYSSISABH As typeYSSISAB0, oldYSSISABH As typeYSSISAB0, newYSSISABH As typeYSSISAB0

Dim arrMNURUTCUT() As String, arrMNURUTCUT_Nb As Integer
Dim arrMNURCLABR(100) As String
Dim mRTF As String
Dim arrJRN_Origine(99) As String

Dim arrMNUGRPNOM() As String, arrMNUGRPNOM_Nb As Integer

'___________________________________________________________________________________________
Dim appWord As Word.Application
Dim docWord As Word.Document
Dim hwndWord As Long
Dim mClipBoard

'___________________________________________________________________________________________

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim mXls2_Cols As Integer, mXls2_Col As Integer, mXls2_Row As Integer
Dim arrProfil_GX() As String, arrProfil_GN() As Long, arrProfil_Nb As Integer
Dim arrProfil_PX() As String, arrProfil_DX() As String, arrProfil_MX() As String
Dim arrYSSIMNU0() As typeWSSIMNU0, arrYSSIMNU0_Nb As Integer

Dim blnProfil_Excel_All As Boolean
Dim blnZMNUOPT0 As Boolean
'___________________________________________________________________________________________

Dim objRootDSE
Dim objDomain
Dim objContainer
Dim objOrganizationalUnit
Dim oAD

Dim mYSSIWIN0_Update As String, mYSSIWINH_Update As String
Dim xYSSIWIN0 As typeYSSIWIN0, oldYSSIWIN0 As typeYSSIWIN0, newYSSIWIN0 As typeYSSIWIN0
Dim usrYSSIWIN0 As typeYSSIWIN0, prfYSSIWIN0 As typeYSSIWIN0
Dim xYSSIWINH As typeYSSIWIN0, oldYSSIWINH As typeYSSIWIN0, newYSSIWINH As typeYSSIWIN0
Dim rtfYSSIWIN0 As typeYSSIWIN0

Dim blnYSSIWIN0_OU_Filter As Boolean, mSWIWINUIDD As Integer
Dim arrYSSIWIN0_OU() As typeYSSIWIN0, arrYSSIWIN0_OU_Nb As Integer, blnYSSIWIN0_OU_New As Boolean
Dim arrYSSIWIN0_User() As Long, arrYSSIWIN0_User_Nb As Long
Dim arrUAC_Lib(21) As String, arrUAC_Val(21) As Long
'___________________________________________________________________________________________

Dim mYSSIDIV0_Update As String, mYSSIDIVH_Update As String
Dim xYSSIDIV0 As typeYSSIDIV0, oldYSSIDIV0 As typeYSSIDIV0, newYSSIDIV0 As typeYSSIDIV0
Dim usrYSSIDIV0 As typeYSSIDIV0, prfYSSIDIV0 As typeYSSIDIV0
Dim xYSSIDIVH As typeYSSIDIV0, oldYSSIDIVH As typeYSSIDIV0, newYSSIDIVH As typeYSSIDIV0
Dim rtfYSSIDIV0 As typeYSSIDIV0

Dim arrSSIUSRUNIT_Code(100) As String
'___________________________________________________________________________________________
Dim olSession
Dim olAPP As New Outlook.Application
Dim olAddressList As Outlook.AddressList
Dim olAddressEntries As Outlook.AddressEntries
Dim olAddressEntry As Outlook.AddressEntry
Dim olAddressEntry_G As Outlook.AddressEntry
Dim olExchangeUser As Outlook.ExchangeUser

Dim arrSSIMELPRFX() As String, newSSIMELPRFX_UIDD As Integer, oldSSIMELPRFX_UIDD As Integer
Dim kSSIMELPRFX_UIDD As Integer
Dim mYSSIMEL0_Update As String, mYSSIMELH_Update As String
Dim xYSSIMEL0 As typeYSSIMEL0, oldYSSIMEL0 As typeYSSIMEL0, newYSSIMEL0 As typeYSSIMEL0
Dim usrYSSIMEL0 As typeYSSIMEL0, prfYSSIMEL0 As typeYSSIMEL0
Dim xYSSIMELH As typeYSSIMEL0, oldYSSIMELH As typeYSSIMEL0, newYSSIMELH As typeYSSIMEL0
Dim rtfYSSIMEL0 As typeYSSIMEL0

Dim mParam_SSIMELNAT_K As Integer
Dim arrSSIMELUIDX() As String, arrSSIMELUNOM() As String, blnSSIMELUNOM() As Boolean, arrSSIMELUNOM_Nb As Integer
'_________________________________________________________________________________________

Dim mYSSITIC0_Update As String, mYSSITICH_Update As String
Dim xYSSITIC0 As typeYSSITIC0, oldYSSITIC0 As typeYSSITIC0, newYSSITIC0 As typeYSSITIC0
Dim usrYSSITIC0 As typeYSSITIC0, prfYSSITIC0 As typeYSSITIC0
Dim xYSSITICH As typeYSSITIC0, oldYSSITICH As typeYSSITIC0, newYSSITICH As typeYSSITIC0
Dim rtfYSSITIC0 As typeYSSITIC0
'_________________________________________________________________________________________

Dim xYSSISAM0 As typeYSSISAM0, oldYSSISAM0 As typeYSSISAM0, newYSSISAM0 As typeYSSISAM0
Dim xYSSISAMH As typeYSSISAM0, oldYSSISAMH As typeYSSISAM0, newYSSISAMH As typeYSSISAM0
Dim rtfYSSISAM0 As typeYSSISAM0

Sub cmdSelect_SQL_9_MEL()
Dim xSQL As String, K As Long, X As String, blnSSIMELPFRX_New As Boolean, wSSIMELPFRX As String
Dim blnYSSIWIN0_Ok As Boolean
Dim mGRP_Old_Nb As Long, mGRP_New_Nb As Long
Dim mUSR_Old_Nb As Long, mUSR_New_Nb As Long, mUSR_Ignore_Nb As Long
Dim wSSIMELUIDX As String
'__________________________________________________________________
Dim oAE As Outlook.AddressEntry
Dim oAEs As Outlook.AddressEntries
Dim oEU As Outlook.ExchangeUser
Dim oDL As Outlook.ExchangeDistributionList
Dim oLists As Outlook.AddressLists
Dim oList As Outlook.AddressList

'__________________________________________________________________


mImport_PRFD = DSys
mImport_PRFH = time_Hms

Call rsYSSIMEL0_Init(usrYSSIMEL0)
usrYSSIMEL0.SSIMELYFCT = "CRE"
usrYSSIMEL0.SSIMELYUSR = usrName_UCase
usrYSSIMEL0.SSIMELYAMJ = mImport_PRFD
usrYSSIMEL0.SSIMELYHMS = mImport_PRFH

Call rsYSSIDOM0_Init(xYSSIDOM0)
xYSSIDOM0.SSIDOMDIDX = "MEL"
xYSSIDOM0.SSIDOMYFCT = "CRE"
xYSSIDOM0.SSIDOMYUSR = usrName_UCase
xYSSIDOM0.SSIDOMYAMJ = mImport_PRFD
xYSSIDOM0.SSIDOMYHMS = mImport_PRFH

xSQL = "select count(*)  from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
     & " where SSIMELNAT = ' '" _
     & " group by SSIMELUIDD"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    mUSR_Old_Nb = mUSR_Old_Nb + rsSab(0)
    rsSab.MoveNext
Loop

xSQL = "select SSIMELUIDD  from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
     & " where SSIMELNAT = '$'" _
     & " order by SSIMELUIDD desc FETCH FIRST 1 ROWS ONLY"
Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    newSSIMELPRFX_UIDD = newSSIMELPRFX_UIDD + 1
    newYSSIMEL0 = usrYSSIMEL0
    newYSSIMEL0.SSIMELNAT = "$"
    newYSSIMEL0.SSIMELUIDX = "Exchange"
    newYSSIMEL0.SSIMELPRFX = "Exchange"
    newYSSIMEL0.SSIMELUNOM = "Exchange@bia-paris.fr"
    newYSSIMEL0.SSIMELUIDD = newSSIMELPRFX_UIDD
    Call cmdUpdate_Init: mYSSIMEL0_Update = "New"
    Call cmdSSIJRN_MEL("<PRFX: |" & newYSSIMEL0.SSIMELUIDX & ">")
    Call cmdUpdate

Else
    newSSIMELPRFX_UIDD = rsSab("SSIMELUIDD")
End If
oldSSIMELPRFX_UIDD = newSSIMELPRFX_UIDD
ReDim arrSSIMELPRFX(newSSIMELPRFX_UIDD + 1)
'_____________________________________________________________________________________

xSQL = "select SSIMELUIDX ,  SSIMELUIDD from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
     & " where SSIMELNAT = '$'" _
     & " order by SSIMELUIDD"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    mGRP_Old_Nb = mGRP_Old_Nb + 1
    K = rsSab("SSIMELUIDD")
    arrSSIMELPRFX(K) = StrConv(Trim(rsSab("SSIMELUIDX")), vbProperCase)

    rsSab.MoveNext
Loop

Call cmdUpdate_Init

If Not blnAuto Then
    newYSSIMEL0 = usrYSSIMEL0
    Call cmdSSIJRN_TXT("MEL", "<ORIG:44><FCT:9-MEL><UID:importation Exchange >" _
                       & " <X:Profils lus : " & mGRP_Old_Nb & ", Utilisateurs lus : " & mUSR_Old_Nb & ">")
    
    Call cmdUpdate
End If
Set olSession = olAPP.Session
'_____________________________________________________________________________________

For Each olAddressList In olSession.AddressLists
    
    If olAddressList.Name = "Tous les groupes" Then
        For Each olAddressEntry In olAddressList.AddressEntries
            Set olExchangeUser = olAddressEntry.GetExchangeUser
            blnSSIMELPFRX_New = True
            wSSIMELPFRX = StrConv(Trim(olAddressEntry.Name), vbProperCase)
            wSSIMELPFRX = Replace(wSSIMELPFRX, Chr(150), Chr(45))
            
            Call lstErr_AddItem(lstErr, cmdContext, "> cmdSelect_SQL_9_MEL : " & wSSIMELPFRX): DoEvents
            For kSSIMELPRFX_UIDD = 1 To newSSIMELPRFX_UIDD
                If wSSIMELPFRX = arrSSIMELPRFX(kSSIMELPRFX_UIDD) Then blnSSIMELPFRX_New = False: Exit For
            Next kSSIMELPRFX_UIDD
            If blnSSIMELPFRX_New Then
                newSSIMELPRFX_UIDD = newSSIMELPRFX_UIDD + 1
                kSSIMELPRFX_UIDD = newSSIMELPRFX_UIDD
                ReDim Preserve arrSSIMELPRFX(newSSIMELPRFX_UIDD + 1)
                
                arrSSIMELPRFX(newSSIMELPRFX_UIDD) = wSSIMELPFRX
                newYSSIMEL0 = usrYSSIMEL0
                newYSSIMEL0.SSIMELNAT = "$"
                newYSSIMEL0.SSIMELUIDX = wSSIMELPFRX
                newYSSIMEL0.SSIMELPRFX = wSSIMELPFRX
                newYSSIMEL0.SSIMELUNOM = wSSIMELPFRX & "@bia-paris.fr"
                newYSSIMEL0.SSIMELUIDD = newSSIMELPRFX_UIDD
                Call cmdUpdate_Init: mYSSIMEL0_Update = "New"
                Call cmdSSIJRN_MEL("<PRFX: |" & newYSSIMEL0.SSIMELUIDX & ">")
                Call cmdUpdate
                mGRP_New_Nb = mGRP_New_Nb + 1
            End If
        '_____________________________________________________________________________________
            If olAddressEntry.AddressEntryUserType = olExchangeDistributionListAddressEntry Then
                Set oDL = olAddressEntry.GetExchangeDistributionList
                Set oAEs = oDL.GetExchangeDistributionListMembers
                For Each oAE In oAEs
                    If oAE.AddressEntryUserType = olExchangeUserAddressEntry _
                        Or oAE.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then
                        Set olExchangeUser = oAE.GetExchangeUser
'______________________________________________________________________________________________
                         wSSIMELUIDX = cmdSelect_SQL_9_MEL_SSIMELUIDX
                         
                         Call lstErr_AddItem(lstErr, cmdContext, "> cmdSelect_SQL_9_MEL : " & wSSIMELPFRX & " " & wSSIMELUIDX): DoEvents
                         
                         usrYSSIMEL0.SSIMELUNOM = StrConv(Trim(olExchangeUser.PrimarySmtpAddress), vbProperCase)
                         xSQL = "select SSIWINUIDD from " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
                              & " where SSIWINNAT = ' ' and SSIWINMAIL = '" & olExchangeUser.PrimarySmtpAddress & "'"
                         Set rsSab = cnsab.Execute(xSQL)
                         
                         If rsSab.EOF Then
                             blnYSSIWIN0_Ok = False
                             mUSR_Ignore_Nb = mUSR_Ignore_Nb + 1
                    
                             Call cmdUpdate_Init
                             newYSSIMEL0 = usrYSSIMEL0
                             Call cmdSSIJRN_TXT_Once("MEL", "<ORIG:43><FCT:9-???><UID:" & usrYSSIMEL0.SSIMELUNOM & "><X:compte Exchange ignoré>")
                             Call cmdUpdate
                        Else
                            
                             blnYSSIWIN0_Ok = True
                             usrYSSIMEL0.SSIMELUIDD = rsSab("SSIWINUIDD")
                             usrYSSIMEL0.SSIMELPRFX = arrSSIMELPRFX(kSSIMELPRFX_UIDD)
                             usrYSSIMEL0.SSIMELUIDX = StrConv(wSSIMELUIDX, vbProperCase) & "_" & kSSIMELPRFX_UIDD
                             cmdSelect_SQL_9_MEL_User
                            
                    
                         End If
'______________________________________________________________________________________________
                Else
                        If oAE.AddressEntryUserType = olExchangeDistributionListAddressEntry Then
                            Call cmdSelect_SQL_9_MEL_ExchangeDistributionList(oAE)
                        End If
                    End If
                Next
            End If
        '_____________________________________________________________________________________
        Next
    End If
    
'$JPL 2015-12-07

    If olAddressList.Name = "Tous les utilisateurs" Then   ' olAddressList.Name = "Utilisateurs BIA" Then
        For Each olAddressEntry In olAddressList.AddressEntries
             Set olExchangeUser = olAddressEntry.GetExchangeUser
             wSSIMELUIDX = cmdSelect_SQL_9_MEL_SSIMELUIDX
             
             Call lstErr_AddItem(lstErr, cmdContext, "> cmdSelect_SQL_9_MEL : " & wSSIMELUIDX): DoEvents
             
             usrYSSIMEL0.SSIMELUNOM = StrConv(Trim(olExchangeUser.PrimarySmtpAddress), vbProperCase)
             xSQL = "select SSIWINUIDD from " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
                  & " where SSIWINNAT = ' ' and SSIWINMAIL = '" & olExchangeUser.PrimarySmtpAddress & "'"
             Set rsSab = cnsab.Execute(xSQL)
             
             If rsSab.EOF Then
                 blnYSSIWIN0_Ok = False
                 mUSR_Ignore_Nb = mUSR_Ignore_Nb + 1
        
                 Call cmdUpdate_Init
                 newYSSIMEL0 = usrYSSIMEL0
                 Call cmdSSIJRN_TXT_Once("MEL", "<ORIG:43><FCT:9-???><UID:" & usrYSSIMEL0.SSIMELUNOM & "><X:compte Exchange ignoré>")
                 Call cmdUpdate
            Else
                 blnYSSIWIN0_Ok = True
                 usrYSSIMEL0.SSIMELUIDD = rsSab("SSIWINUIDD")
                 kSSIMELPRFX_UIDD = 1
                 usrYSSIMEL0.SSIMELPRFX = arrSSIMELPRFX(kSSIMELPRFX_UIDD)
                 usrYSSIMEL0.SSIMELUIDX = wSSIMELUIDX
                 cmdSelect_SQL_9_MEL_User
             End If
             
'________________________________________________________________________________________________
            Set olAddressEntries = olExchangeUser.GetMemberOfList
             
             For Each olAddressEntry_G In olAddressEntries
                 blnSSIMELPFRX_New = True
                 X = StrConv(Trim(olAddressEntry_G.Name), vbProperCase)
                 X = Replace(X, Chr(150), Chr(45))
                 For kSSIMELPRFX_UIDD = 1 To newSSIMELPRFX_UIDD
                     If X = arrSSIMELPRFX(kSSIMELPRFX_UIDD) Then blnSSIMELPFRX_New = False: Exit For
                 Next kSSIMELPRFX_UIDD
                 If blnSSIMELPFRX_New Then
                     newSSIMELPRFX_UIDD = newSSIMELPRFX_UIDD + 1
                     kSSIMELPRFX_UIDD = newSSIMELPRFX_UIDD
                     ReDim Preserve arrSSIMELPRFX(newSSIMELPRFX_UIDD + 1)
                     
                     arrSSIMELPRFX(newSSIMELPRFX_UIDD) = X
                     newYSSIMEL0 = usrYSSIMEL0
                     newYSSIMEL0.SSIMELNAT = "$"
                     newYSSIMEL0.SSIMELUIDX = X
                     newYSSIMEL0.SSIMELPRFX = X
                     newYSSIMEL0.SSIMELUNOM = X & "@bia-paris.fr"
                     newYSSIMEL0.SSIMELUIDD = newSSIMELPRFX_UIDD
                     Call cmdUpdate_Init: mYSSIMEL0_Update = "New"
                     Call cmdSSIJRN_MEL("<PRFX: |" & newYSSIMEL0.SSIMELUIDX & ">")
                     Call cmdUpdate
                     mGRP_New_Nb = mGRP_New_Nb + 1

                 End If
                 If blnYSSIWIN0_Ok Then

                     usrYSSIMEL0.SSIMELUIDX = StrConv(wSSIMELUIDX, vbProperCase) & "_" & kSSIMELPRFX_UIDD
                     usrYSSIMEL0.SSIMELPRFX = arrSSIMELPRFX(kSSIMELPRFX_UIDD)
                     cmdSelect_SQL_9_MEL_User
                 End If
             Next

        Next
    End If
Next
'_____________________________________________________________________________________

'xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0," & paramIBM_Library_SABSPE & ".YSSIMEL0" _
'     & " where SSIDOMNAT = ' ' and SSIDOMDIDX = 'MEL' and SSIDOMPRFX = 'Exchange'" _
'     & " and ( SSIDOMPRFD <> " & mImport_PRFD & " or  SSIDOMPRFH <> " & mImport_PRFH & ")" _
'     & " and SSIMELNAT = SSIDOMNAT and SSIMELUIDX = SSIDOMUIDX"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0," & paramIBM_Library_SABSPE & ".YSSIMEL0" _
     & " where SSIDOMNAT = ' ' and SSIDOMDIDX = 'MEL' and SSIDOMPRFK <> 'X'" _
     & " and SSIDOMPRFD <> " & mImport_PRFD _
     & " and SSIMELNAT = SSIDOMNAT and SSIMELUIDX = SSIDOMUIDX"

Set rsSab_X = cnsab.Execute(xSQL)

Do While Not rsSab_X.EOF
    Call cmdUpdate_Init
    Call rsYSSIDOM0_GetBuffer(rsSab_X, oldYSSIDOM0)
    newYSSIDOM0 = oldYSSIDOM0
    newYSSIDOM0.SSIDOMPRFK = "X"
    newYSSIDOM0.SSIDOMYFCT = "SUP"
    newYSSIDOM0.SSIDOMYUSR = usrName_UCase
    newYSSIDOM0.SSIDOMYAMJ = mImport_PRFD
    newYSSIDOM0.SSIDOMYHMS = mImport_PRFH
    mYSSIDOM0_Update = "Update+H"
    
    Call rsYSSIMEL0_GetBuffer(rsSab_X, oldYSSIMEL0)
    newYSSIMEL0 = oldYSSIMEL0
    newYSSIMEL0.SSIMELPRFK = "X"
    newYSSIMEL0.SSIMELYFCT = "SUP"
    newYSSIMEL0.SSIMELYUSR = usrName_UCase
    newYSSIMEL0.SSIMELYAMJ = mImport_PRFD
    newYSSIMEL0.SSIMELYHMS = mImport_PRFH
    newYSSIMEL0.SSIMELINFO = newYSSIMEL0.SSIMELINFO & "Supprimé"
    mYSSIMEL0_Update = "Update+H"
    Call cmdSSIJRN_MEL("<PRFX: |" & newYSSIMEL0.SSIMELPRFX & ">")
    Call cmdUpdate
    
    Dim rsSab_Y As New ADODB.Recordset
    Call cmdUpdate_Init
     X = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
         & " where SSIMELNAT = '@'  and SSIMELINFO like '%" & newYSSIMEL0.SSIMELUNOM & "%'" _
         & " order by SSIMELUIDX"
    Set rsSab_Y = cnsab.Execute(X)
    
    Do While Not rsSab_Y.EOF
        Call cmdSSIJRN_MEL("<X: à supprimer de l'envoi " & rsSab_Y("SSIMELUIDX") & ">")
        Call cmdUpdate
        rsSab_Y.MoveNext
    Loop
   
    rsSab_X.MoveNext
Loop
'_____________________________________________________________________________________

Call cmdUpdate_Init
newYSSIMEL0 = usrYSSIMEL0
xSQL = "select count(*)  from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
     & " where SSIMELNAT = ' '" _
     & " group by SSIMELUIDD"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    mUSR_New_Nb = mUSR_New_Nb + rsSab(0)
    rsSab.MoveNext
Loop

If Not blnAuto Then
    X = "Profils créés : " & mGRP_New_Nb _
      & ", Utilisateurs créés : " & mUSR_New_Nb - mUSR_Old_Nb _
      & ", Comptes Ignorés : " & mUSR_Ignore_Nb
    Call cmdSSIJRN_TXT("MEL", "<ORIG:44><FCT:9-MEL><UID:importation Exchange ><X:" & X & ">")
    Call cmdUpdate
End If
End Sub

Sub cmdSelect_SQL_9_MEL_ExchangeDistributionList(oAddress As AddressEntry)
    Dim oAE As Outlook.AddressEntry
    Dim oAEs As Outlook.AddressEntries
    'Dim oEU As Outlook.ExchangeUser
    Dim oDL As Outlook.ExchangeDistributionList
    Dim wSSIMELUIDX As String, xSQL As String
    Set oDL = oAddress.GetExchangeDistributionList
    Set oAEs = oDL.GetExchangeDistributionListMembers
    For Each oAE In oAEs
        If oAE.AddressEntryUserType = olExchangeUserAddressEntry _
            Or oAE.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then
            Set olExchangeUser = oAE.GetExchangeUser
'______________________________________________________________________________________________
             wSSIMELUIDX = cmdSelect_SQL_9_MEL_SSIMELUIDX
             
             Call lstErr_AddItem(lstErr, cmdContext, "> cmdSelect_SQL_9_MEL : " & arrSSIMELPRFX(kSSIMELPRFX_UIDD) & " " & wSSIMELUIDX): DoEvents
             
             usrYSSIMEL0.SSIMELUNOM = StrConv(Trim(olExchangeUser.PrimarySmtpAddress), vbProperCase)
             xSQL = "select SSIWINUIDD from " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
                  & " where SSIWINNAT = ' ' and SSIWINMAIL = '" & olExchangeUser.PrimarySmtpAddress & "'"
             Set rsSab = cnsab.Execute(xSQL)
             
             If rsSab.EOF Then
                ' blnYSSIWIN0_Ok = False
                ' mUSR_Ignore_Nb = mUSR_Ignore_Nb + 1
        
                 Call cmdUpdate_Init
                 newYSSIMEL0 = usrYSSIMEL0
                 Call cmdSSIJRN_TXT_Once("MEL", "<ORIG:43><FCT:9-???><UID:" & usrYSSIMEL0.SSIMELUNOM & "><X:compte Exchange ignoré>")
                 Call cmdUpdate
            Else
                
                 'blnYSSIWIN0_Ok = True
                 usrYSSIMEL0.SSIMELUIDD = rsSab("SSIWINUIDD")
                 usrYSSIMEL0.SSIMELPRFX = arrSSIMELPRFX(kSSIMELPRFX_UIDD)
                 usrYSSIMEL0.SSIMELUIDX = StrConv(wSSIMELUIDX, vbProperCase) & "_" & kSSIMELPRFX_UIDD
                 cmdSelect_SQL_9_MEL_User
        
             End If
'______________________________________________________________________________________________
        End If
    Next
End Sub

Sub JPL()
On Error Resume Next
Dim xSQL As String, K As Long, Nb As Long

Dim xTerena As String, xIn As String
Dim lenTerena As Long, blnExit As Boolean
Dim K1 As Long, K10 As Long

Open "c:\temp2\terena.txt" For Input As 1

Do Until EOF(1)
    Line Input #1, xTerena
Loop
Close 1
lenTerena = Len(xTerena)
blnExit = False
K10 = 1

Do Until blnExit
    K = InStr(K10, xTerena, Chr(10))
    If K > 0 Then
        xIn = Mid$(xTerena, K10, K - K10)
        K10 = K + 1
        K = InStr(1, xIn, Chr(9)): Debug.Print "Id : "; Mid$(xIn, 1, K - 1)
        K1 = K + 1: K = InStr(K1, xIn, Chr(9)): Debug.Print "auto : "; Mid$(xIn, K1, K - K1)
        K1 = K + 1: K = InStr(K1, xIn, Chr(9)): Debug.Print "auto : "; Mid$(xIn, K1, K - K1)
        K1 = K + 1: K = InStr(K1, xIn, Chr(9)): Debug.Print "JF : "; Mid$(xIn, K1, K - K1)
        K1 = K + 1: K = InStr(K1, xIn, Chr(9)): Debug.Print "date début : "; Mid$(xIn, K1, K - K1)
        K1 = K + 1: K = InStr(K1, xIn, Chr(9)): Debug.Print "datefin : "; Mid$(xIn, K1, K - K1)
        K1 = K + 1: K = InStr(K1, xIn, Chr(9)): Debug.Print "zone : "; Mid$(xIn, K1, K - K1)
        K1 = K + 1: K = InStr(K1, xIn, Chr(9)): Debug.Print "grp L 1 2 : "; Mid$(xIn, K1, K - K1)
        K1 = K + 1: K = InStr(K1, xIn, Chr(9)): Debug.Print "grp H 1 2 : "; Mid$(xIn, K1, K - K1)
    Else
        blnExit = True
    End If
Loop

Exit Sub
'====================================================================

'xSQL = "select * from qsys2.systables "
xSQL = "select * from qsys2.systistat "
'xSQL = "select * from qsys2.syscolumns"
Set rsSab = cnsab.Execute(xSQL)

'Open "C:\TEMP\qsys2_systables.txt" For Output As 3
Open "C:\TEMP\qsys2_systistat.txt" For Output As 3
'Open "C:\TEMP\qsys2_syscolumns.txt" For Output As 3

Do While Not rsSab.EOF
    Nb = Nb + 1
    If Nb Mod 1000 = 0 Then Call lstErr_ChangeLastItem(lstErr, cmdContext, "JPL : " & Nb): DoEvents

    'For K = 0 To 29
    For K = 0 To 64
    'For K = 0 To 38
        Print #3, Trim(rsSab(K)) & "|";
    Next K
    Print #3, "|"
    rsSab.MoveNext
Loop

Close 3





Exit Sub
'====================================================================
Dim olSession
Dim olAPP As New Outlook.Application
Dim olAddressList As Outlook.AddressList

Dim olAddressEntries As Outlook.AddressEntries
Dim olAddressEntry As Outlook.AddressEntry
Dim olAddressEntry_G As Outlook.AddressEntry
Dim olExchangeUser As Outlook.ExchangeUser


Set olSession = olAPP.Session
For Each olAddressList In olSession.AddressLists
    Debug.Print olAddressList.Name
    
    If olAddressList.Name = "Utilisateurs BIA" Then
        For Each olAddressEntry In olAddressList.AddressEntries
                Set olExchangeUser = olAddressEntry.GetExchangeUser
               Debug.Print olExchangeUser.PrimarySmtpAddress; olExchangeUser.Name; olExchangeUser.Type; olExchangeUser.Address
                Set olAddressEntries = olExchangeUser.GetMemberOfList
                    For Each olAddressEntry_G In olAddressEntries
                        Debug.Print "......"; olAddressEntry_G.Name
                    Next
        Next
    End If
Next

End Sub

Public Sub cmdSelect_SQL_9_WIN()
Dim K As Long

Call lstErr_AddItem(lstErr, cmdContext, "> cmdSelect_SQL_9_WIN"): DoEvents

Call rsYSSIWIN0_Init(xYSSIWIN0)

Set objRootDSE = GetObject("LDAP://RootDSE")
Set objDomain = GetObject("LDAP://" & objRootDSE.Get("DefaultNamingContext"))

Call lstErr_AddItem(lstErr, cmdContext, "> cmdSelect_SQL_9_WIN_OU"): DoEvents
blnYSSIWIN0_OU_Filter = True
Call cmdSelect_SQL_9_WIN_EnumOUs(objDomain.ADsPath)

Call lstErr_AddItem(lstErr, cmdContext, "> cmdSelect_SQL_9_WIN_User"): DoEvents
Call YSSIWIN0_OU_Load
Call YSSIWIN0_User_Load
blnYSSIWIN0_OU_Filter = False
Call rsYSSIWIN0_Init(xYSSIWIN0)
Call cmdSelect_SQL_9_WIN_EnumOUs(objDomain.ADsPath)

For K = 1 To arrYSSIWIN0_User_Nb
    If arrYSSIWIN0_User(K) <> 0 Then
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
             & " where SSIWINNAT = ' ' and SSIWINUIDD = " & arrYSSIWIN0_User(K)
        Set rsSab = cnsab.Execute(X)
        
        If Not rsSab.EOF Then
            Call rsYSSIWIN0_GetBuffer(rsSab, oldYSSIWIN0)
            Call lstErr_AddItem(lstErr, cmdContext, "> cmdSelect_SQL_9_WIN_User_Supprimé : " & rsSab("SSIWINUIDX")): DoEvents
            mYSSIWIN0_Update = "Update+H"
            oldYSSIWINH = oldYSSIWIN0
            newYSSIWIN0 = oldYSSIWIN0
            newYSSIWIN0.SSIWINPRFK = "S"
            newYSSIWIN0.SSIWINYFCT = "SUP"
            newYSSIWIN0.SSIWINYUSR = usrName_UCase
            newYSSIWIN0.SSIWINYAMJ = DSys
            newYSSIWIN0.SSIWINYHMS = time_Hms
                      
            Call cmdSSIJRN_WIN("<X:CTL automate compte WIN supprimé>")
            Call cmdUpdate
        End If
    End If
Next K

End Sub

Sub cmdSelect_SQL_9_WIN_EnumOUs(sADsPath)
    Set objContainer = GetObject(sADsPath)
    If blnYSSIWIN0_OU_Filter Then objContainer.Filter = Array("organizationalUnit")
    For Each objOrganizationalUnit In objContainer
        cmdSelect_SQL_9_WIN_EnumUsers (objOrganizationalUnit.ADsPath)
        cmdSelect_SQL_9_WIN_EnumOUs (objOrganizationalUnit.ADsPath)
        
    Next
    
End Sub


Sub cmdSelect_SQL_9_WIN_EnumUsers(sADsPath)
    Dim K As Integer, strGUID As String, xInfo As String, X As String, blnOk As Boolean
    Dim xName As String, xAccountExpirationDate As String, blnAccountExpirationDate As Boolean
    Set objContainer = GetObject(sADsPath)
    If blnYSSIWIN0_OU_Filter Then objContainer.Filter = Array("organizationalUnit")
    For Each oAD In objContainer
        If oAD.Class = "user" Or oAD.Class = "computer" Or oAD.Class = "printQueue" _
        Or oAD.Class = "group" Or oAD.Class = "organizationalUnit" Then
        
        'Or oAD.Class = "publicFolder"
        
            strGUID = ""
            For K = 1 To Len(oAD.Guid)
                strGUID = strGUID & Chr(Asc(Mid$(oAD.Guid, K, 1)))
            Next K
            
           xName = oAD.Name
           blnAccountExpirationDate = False
           If oAD.Class = "user" Then
                Call dateJMA_AMJ(oAD.AccountExpirationDate, xAccountExpirationDate)
                If xAccountExpirationDate <= 19700101 Then
                    xAccountExpirationDate = ""
                Else
                    If xAccountExpirationDate <= DSys Then blnAccountExpirationDate = True
                End If
           Else
                xAccountExpirationDate = " "
            End If
           xInfo = Replace(xName, "|", "") & "|" _
                  & oAD.userAccountControl & "|" _
                  & oAD.Class & "|" _
                  & Replace(oAD.CN, "|", "") & "|" _
                  & Replace(oAD.SN, "|", "") & "|" _
                  & Replace(oAD.givenName, "|", "") & "|" _
                  & Replace(oAD.DisplayName, "|", "") & "|" _
                  & Replace(oAD.userPrincipalName, "|", "") & "|" _
                  & Replace(oAD.company, "|", "") & "|" _
                  & Replace(oAD.Department, "|", "") & "|" _
                  & Replace(oAD.scriptPath, "|", "") & "|" _
                  & oAD.whenCreated & "|" _
                  & Replace(oAD.Description, "|", "") & "|" _
                  & Replace(oAD.distinguishedName, "|", "") & "|" _
                  & Replace(oAD.mail, "|", "") & "|" _
                  & Replace(oAD.mailnickname, "|", "") & "|" _
                  & Replace(oAD.physicalDeliveryOfficeName, "|", "") & "|" _
                  & Replace(oAD.sAMAccountName, "|", "") & "|" _
                  & Replace(oAD.textEncodedORAddress, "|", "") & "|" _
                  & Replace(oAD.userWorkstations, "|", "") & "|" _
                  & xAccountExpirationDate & "|"
                  
             xInfo = Replace(xInfo, Chr$(146), Chr$(39))
             
            xYSSIWIN0.SSIWINGUID = strGUID
            
            xYSSIWIN0.SSIWINMAIL = Trim(oAD.mail)
            If Len(xYSSIWIN0.SSIWINMAIL) > 32 Then xYSSIWIN0.SSIWINMAIL = ""
            
            xYSSIWIN0.SSIWININFO = Replace(xInfo, Chr(150), Chr(45))
            If Len(xYSSIWIN0.SSIWININFO) > 512 Then xYSSIWIN0.SSIWININFO = Mid$(xYSSIWIN0.SSIWININFO, 1, 512)
            
            If oAD.Class = "organizationalUnit" Then
                If blnYSSIWIN0_OU_Filter Then
                    xYSSIWIN0.SSIWINNAT = "$"
                    xYSSIWIN0.SSIWINUIDX = Replace(xName, "OU=", "")
                    If Len(xYSSIWIN0.SSIWINUIDX) > 20 Then xYSSIWIN0.SSIWINUIDX = Mid$(xYSSIWIN0.SSIWINUIDX, 1, 20)
                    xYSSIWIN0.SSIWINPRFX = xYSSIWIN0.SSIWINUIDX
                    xYSSIWIN0.SSIWINUNOM = Replace(oAD.distinguishedName, ",DC=bia-paris,DC=lan", "")
                    If Len(xYSSIWIN0.SSIWINUNOM) > 64 Then xYSSIWIN0.SSIWINUNOM = Mid$(xYSSIWIN0.SSIWINUNOM, 1, 64)
                    Call cmdSelect_SQL_9_WIN_Control_OU
                End If
            Else
                xYSSIWIN0.SSIWINUIDX = Replace(xName, "CN=", "")
                If Len(xYSSIWIN0.SSIWINUIDX) > 20 Then xYSSIWIN0.SSIWINUIDX = Mid$(xYSSIWIN0.SSIWINUIDX, 1, 20)
                xYSSIWIN0.SSIWINUNOM = Replace(oAD.DisplayName, "'", " ")
                If Len(xYSSIWIN0.SSIWINUNOM) > 64 Then xYSSIWIN0.SSIWINUNOM = Mid$(xYSSIWIN0.SSIWINUNOM, 1, 64)
                
                xYSSIWIN0.SSIWINNAT = " "
                xYSSIWIN0.SSIWINPRFX = arrYSSIWIN0_OU(1).SSIWINUIDX
                For K = 1 To arrYSSIWIN0_OU_Nb
                    If InStr(oAD.distinguishedName, arrYSSIWIN0_OU(K).SSIWINUNOM) Then
                        xYSSIWIN0.SSIWINPRFX = arrYSSIWIN0_OU(K).SSIWINUIDX
                        Exit For
                    End If
                Next K
                
                
                Select Case oAD.userAccountControl
                    Case 512: xYSSIWIN0.SSIWINPRFK = " "
                    Case 514: xYSSIWIN0.SSIWINPRFK = "X"
                    Case Else:
                            If oAD.userAccountControl / 2 Mod 2 = 0 Then
                                If oAD.Class = "user" And InStr(xYSSIWIN0.SSIWINPRFX, "Actifs") > 0 Then
                                    xYSSIWIN0.SSIWINPRFK = "N"
                                Else
                                    xYSSIWIN0.SSIWINPRFK = " "
                                End If
                            Else
                                xYSSIWIN0.SSIWINPRFK = "X"
                            End If
                End Select
                If blnAccountExpirationDate Then
                    xYSSIWIN0.SSIWINPRFK = "X"
                    'Debug.Print xName
                End If
                Call cmdSelect_SQL_9_WIN_Control_User
            End If
        End If
    Next
End Sub






Private Sub cmdPrint_Word_PDF(lFileName As String)
Dim X As String, wFile As String
Dim mWord_PDF_Path As String
On Error GoTo Error_Handler

Call lstErr_Clear(lstErr, cmdContext, "cmdPrint_Word_PDF ...."): DoEvents
lstErr.Height = 510

wFile = lFileName & ".rtf"
If Dir(wFile) <> "" Then Kill wFile
txtRTF.SaveFile wFile

'__________________________________________________________________

ProgressBar1.Visible = True
ProgressBar1.Min = 0: ProgressBar1.Max = 5
ProgressBar1.Value = 1

Set appWord = New Word.Application

hwndWord = FindWindow(vbNullString, "Microsoft Word")
If hwndWord <> 0 Then
    Dim hwnd As Long
    hwnd = SetForegroundWindow(hwndWord)
Else
   MsgBox "Impossible de trouver la fenêtre Word!", vbExclamation
End If
'Sleep 2000

appWord.Documents.Add wFile

ProgressBar1.Value = ProgressBar1.Value + 1

mWord_PDF_Path = Environ("commonprogramfiles") & "\Microsoft Shared\OFFICE" _
& Format(Val(appWord.Version), "00") & "\EXP_PDF.DLL"

ProgressBar1.Value = ProgressBar1.Value + 1
wFile = lFileName & ".pdf"
    
Call appWord.ActiveDocument.ExportAsFixedFormat(wFile, wdExportFormatPDF, False, wdExportOptimizeForPrint)
            
ProgressBar1.Visible = False
appWord.Quit False
   
GoTo Exit_sub

Error_Handler:
    MsgBox Error
    appWord.Quit False
Exit_sub:
   Set docWord = Nothing
   Set appWord = Nothing
   DestroyWindow hwndWord

    Call lstErr_AddItem(lstErr, cmdContext, "Fermeture Word " & hwndWord): DoEvents

End Sub





Public Sub fgCompteH_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgCompteH.Visible = False
mRow = fgCompteH.Row

If lRow > 0 And lRow < fgCompteH.Rows Then
    fgCompteH.Row = lRow
    For I = 0 To 0 Step -1
        fgCompteH.Col = I: fgCompteH.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgCompteH.Row = mRow
    If fgCompteH.Row > 0 Then
        lRow = fgCompteH.Row
        lColor_Old = fgCompteH.CellBackColor
        For I = 0 To 0 Step -1
          fgCompteH.Col = I: fgCompteH.CellBackColor = lColor
        Next I
    End If
End If
fgCompteH.LeftCol = fgCompteH.FixedCols
fgCompteH.Visible = True
End Sub

Public Sub fgCompteH_Reset()
fgCompteH.Clear
fgCompteH_Sort1 = 0: fgCompteH_Sort2 = 0
fgCompteH_Sort1_Old = -1
fgCompteH_RowDisplay = 0: fgCompteH_RowClick = 0
fgCompteH_arrIndex = fgCompteH.Cols - 1
blnfgCompteH_DisplayLine = False
fgCompteH_SortAD = 6
fgCompteH.LeftCol = fgCompteH.FixedCols

End Sub

Public Sub fgCompteH_Sort()

If fgCompteH.Rows > 1 Then
    fgCompteH.Row = 1
    fgCompteH.RowSel = fgCompteH.Rows - 1
    
    If fgCompteH_Sort1_Old = fgCompteH_Sort1 Then
        If fgCompteH_SortAD = 5 Then
            fgCompteH_SortAD = 6
        Else
            fgCompteH_SortAD = 5
        End If
    Else
        fgCompteH_SortAD = 5
    End If
    fgCompteH_Sort1_Old = fgCompteH_Sort1
    
    fgCompteH.Col = fgCompteH_Sort1
    fgCompteH.ColSel = fgCompteH_Sort2
    fgCompteH.Sort = fgCompteH_SortAD
End If

End Sub

Public Sub fgCompteH_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgCompteH.Rows - 1
    fgCompteH.Row = I
    fgCompteH.Col = lK
    Select Case lK
'        Case 3: fgCompteH.Col = 3: X = Format$(Val(fgCompteH.Text), "000000000000000.00")

    End Select
    fgCompteH.Col = fgCompteH_arrIndex - 1
    fgCompteH.Text = X
Next I

fgCompteH_Sort1 = fgCompteH_arrIndex - 1: fgCompteH_Sort2 = fgCompteH_arrIndex - 1
fgCompteH_Sort
End Sub






Public Sub fgCompte_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgCompte.Visible = False
mRow = fgCompte.Row

If lRow > 0 And lRow < fgCompte.Rows Then
    fgCompte.Row = lRow
    For I = 1 To 0 Step -1
        fgCompte.Col = I: fgCompte.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgCompte.Row = mRow
    If fgCompte.Row > 0 Then
        lRow = fgCompte.Row
        lColor_Old = fgCompte.CellBackColor
        For I = 1 To 0 Step -1
          fgCompte.Col = I: fgCompte.CellBackColor = lColor
        Next I
    End If
End If
fgCompte.LeftCol = fgCompte.FixedCols
fgCompte.Visible = True
End Sub

Private Sub fgCompte_Display_IBM(lColor As Long)

Dim K As Long, xSQL As String, blnRupture As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgCompte_Display_IBM"
fgCompte.Visible = False
fgCompte_Reset
fgCompte.FormatString = fgCompte_FormatString
fgCompte.Rows = 1
                 
fgCompte.Row = 0
Do While Not rsSab.EOF
    If fgCompte.Row = 0 Then Call rsYSSIIBM0_GetBuffer(rsSab, usrYSSIIBM0)
    fgCompte.Rows = fgCompte.Rows + 1
    fgCompte.Row = fgCompte.Rows - 1
    fgCompte_Display_IBM_Line
    fgCompte.Col = 1: fgCompte.CellFontBold = True: fgCompte.CellForeColor = lColor

    rsSab.MoveNext
Loop

fgCompte.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgCompte.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Private Sub fgCompte_Display_SAA(lColor As Long)

Dim K As Long, xSQL As String, blnRupture As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgCompte_Display_SAA"
fgCompte.Visible = False
fgCompte_Reset
fgCompte.FormatString = fgCompte_FormatString
fgCompte.Rows = 1
                 
fgCompte.Row = 0
Do While Not rsSab.EOF
    If fgCompte.Row = 0 Then Call rsYSSISAA0_GetBuffer(rsSab, usrYSSISAA0)
    fgCompte.Rows = fgCompte.Rows + 1
    fgCompte.Row = fgCompte.Rows - 1
    fgCompte_Display_SAA_Line
    fgCompte.Col = 1: fgCompte.CellFontBold = True: fgCompte.CellForeColor = lColor

    rsSab.MoveNext
Loop

fgCompte.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgCompte.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


Private Sub fgCompte_Display_WIN(lColor As Long)

Dim K As Long, xSQL As String, blnRupture As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgCompte_Display_WIN"
fgCompte.Visible = False
fgCompte_Reset
fgCompte.FormatString = fgCompte_FormatString
fgCompte.Rows = 1
                 
fgCompte.Row = 0
Do While Not rsSab.EOF
    If fgCompte.Row = 0 Then Call rsYSSIWIN0_GetBuffer(rsSab, usrYSSIWIN0)
    fgCompte.Rows = fgCompte.Rows + 1
    fgCompte.Row = fgCompte.Rows - 1
    fgCompte_Display_WIN_Line
    fgCompte.Col = 1: fgCompte.CellFontBold = True: fgCompte.CellForeColor = lColor

    rsSab.MoveNext
Loop

fgCompte.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgCompte.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub fgCompte_Display_DIV(lColor As Long)

Dim K As Long, xSQL As String, blnRupture As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgCompte_Display_DIV"
fgCompte.Visible = False
fgCompte_Reset
fgCompte.FormatString = fgCompte_FormatString
fgCompte.Rows = 1
                 
fgCompte.Row = 0
Do While Not rsSab.EOF
    'If fgCompte.Row = 0 Then Call rsYSSIDIV0_GetBuffer(rsSab, usrYSSIDIV0)
    fgCompte.Rows = fgCompte.Rows + 1
    fgCompte.Row = fgCompte.Rows - 1
    fgCompte_Display_DIV_Line
    fgCompte.Col = 1: fgCompte.CellFontBold = True: fgCompte.CellForeColor = lColor

    rsSab.MoveNext
Loop

fgCompte.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgCompte.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


Private Sub fgCompte_Display_MEL(lColor As Long)

Dim K As Long, xSQL As String, blnRupture As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgCompte_Display_MEL"
fgCompte.Visible = False
fgCompte_Reset
fgCompte.FormatString = fgCompte_FormatString
fgCompte.Rows = 1
                 
fgCompte.Row = 0
Do While Not rsSab.EOF
    'If fgCompte.Row = 0 Then Call rsYSSIMEL0_GetBuffer(rsSab, usrYSSIMEL0)
    fgCompte.Rows = fgCompte.Rows + 1
    fgCompte.Row = fgCompte.Rows - 1
    fgCompte_Display_MEL_Line
    fgCompte.Col = 1: fgCompte.CellFontBold = True: fgCompte.CellForeColor = lColor

    rsSab.MoveNext
Loop

fgCompte.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgCompte.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Private Sub fgCompte_Display_TIC(lColor As Long)

Dim K As Long, xSQL As String, blnRupture As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgCompte_Display_TIC"
fgCompte.Visible = False
fgCompte_Reset
fgCompte.FormatString = fgCompte_FormatString
fgCompte.Rows = 1
                 
fgCompte.Row = 0
Do While Not rsSab.EOF
    'If fgCompte.Row = 0 Then Call rsYSSITIC0_GetBuffer(rsSab, usrYSSITIC0)
    fgCompte.Rows = fgCompte.Rows + 1
    fgCompte.Row = fgCompte.Rows - 1
    fgCompte_Display_TIC_Line
    fgCompte.Col = 1: fgCompte.CellFontBold = True: fgCompte.CellForeColor = lColor

    rsSab.MoveNext
Loop

fgCompte.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgCompte.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub









Private Sub fgCompte_Display_SAB(lColor As Long)

Dim K As Long, xSQL As String, blnRupture As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgCompte_Display_SAB"
fgCompte.Visible = False
fgCompte_Reset
fgCompte.FormatString = fgCompte_FormatString
fgCompte.Rows = 1
                 
fgCompte.Row = 0
Do While Not rsSab.EOF
    If fgCompte.Row = 0 Then Call rsYSSISAB0_GetBuffer(rsSab, usrYSSISAB0)
    fgCompte.Rows = fgCompte.Rows + 1
    fgCompte.Row = fgCompte.Rows - 1
    fgCompte_Display_SAB_Line
    fgCompte.Col = 1: fgCompte.CellFontBold = True: fgCompte.CellForeColor = lColor

    rsSab.MoveNext
Loop

fgCompte.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgCompte.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub



Private Sub fgCompteH_Display(lDIDX As String)

Dim K As Long, xSQL As String, blnRupture As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgCompteH_Display"
fgCompteH.Visible = False
fgCompteH_Reset
fgCompteH.FormatString = fgCompteH_FormatString
fgCompteH.Rows = 1
                 
fgCompteH.Row = 0
Do While Not rsSab.EOF
    'Call rsYSSIUSR0_GetBuffer(rsSab, xYSSIUSR0)
    fgCompteH.Rows = fgCompteH.Rows + 1
    fgCompteH.Row = fgCompteH.Rows - 1
    Select Case lDIDX
        Case "IBM": fgCompteH_Display_IBM_Line
        Case "SAA": fgCompteH_Display_SAA_Line
        Case "SAB": fgCompteH_Display_SAB_Line
        Case "WIN": fgCompteH_Display_WIN_Line
        Case "DIV": fgCompteH_Display_DIV_Line
        Case "MEL": fgCompteH_Display_MEL_Line
        Case "TIC": fgCompteH_Display_TIC_Line
    End Select
    rsSab.MoveNext
Loop

fgCompteH.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgCompteH.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub fgCompte_Display_IBM_Line()
On Error Resume Next

fgCompte.Col = 0: fgCompte.Text = rsSab("SSIIBMUIDD")
fgCompte.Col = 1: fgCompte.Text = rsSab("UPUPRF")
fgCompte.Col = 2: fgCompte.Text = rsSab("UPTEXT")
fgCompte.CellFontSize = 8


End Sub

Public Sub fgCompte_Display_SAA_Line()
On Error Resume Next

fgCompte.Col = 0: fgCompte.Text = rsSab("SSISAAUIDD")
fgCompte.Col = 1: fgCompte.Text = rsSab("SSISAAUIDX")
fgCompte.Col = 2: fgCompte.Text = rsSab("SSISAAUNOM")
fgCompte.CellFontSize = 8


End Sub

Public Sub fgCompte_Display_WIN_Line()
On Error Resume Next

fgCompte.Col = 0: fgCompte.Text = rsSab("SSIWINUIDD")
fgCompte.Col = 1: fgCompte.Text = rsSab("SSIWINUIDX")
fgCompte.Col = 2: fgCompte.Text = rsSab("SSIWINPRFX")
fgCompte.CellFontSize = 8


End Sub


Public Sub fgCompte_Display_DIV_Line()
On Error Resume Next

fgCompte.Col = 0: fgCompte.Text = rsSab("SSIDIVUIDD")
fgCompte.Col = 1: fgCompte.Text = rsSab("SSIDIVUIDX")
fgCompte.Col = 2: fgCompte.Text = rsSab("SSIDIVPRFX")
fgCompte.CellFontSize = 8


End Sub

Public Sub fgCompte_Display_MEL_Line()
On Error Resume Next

fgCompte.Col = 0: fgCompte.Text = rsSab("SSIMELUIDD")
fgCompte.Col = 1: fgCompte.Text = rsSab("SSIMELUIDX")
fgCompte.Col = 2: fgCompte.Text = rsSab("SSIMELPRFX")
fgCompte.CellFontSize = 8


End Sub


Public Sub fgCompte_Display_TIC_Line()
On Error Resume Next

fgCompte.Col = 0: fgCompte.Text = rsSab("SSITICUIDD")
fgCompte.Col = 1: fgCompte.Text = rsSab("SSITICUIDX")
fgCompte.Col = 2: fgCompte.Text = rsSab("SSITICPRFX")
fgCompte.CellFontSize = 8


End Sub



Public Sub fgCompte_Display_SAB_Line()
On Error Resume Next

fgCompte.Col = 0: fgCompte.Text = rsSab("SSISABUIDD")
fgCompte.Col = 1: fgCompte.Text = rsSab("SSISABUIDX")
fgCompte.Col = 2: fgCompte.Text = rsSab("SSISABUNOM")
fgCompte.CellFontSize = 8


End Sub


Public Sub fgCompteH_Display_SAA_Line()
On Error Resume Next
fgCompteH.Col = 0: fgCompteH.Text = rsSab("SSISAANAT") & "|" & rsSab("SSISAAUIDX") & "|" & rsSab("SSISAAYVER")
fgCompteH.Col = 1: fgCompteH.Text = rsSab("SSISAAPRFK")
Select Case rsSab("SSISAAPRFK")
    Case " ": fgCompteH.CellBackColor = mColor_G1
    Case Else: fgCompteH.CellBackColor = mColor_W0
                fgCompteH.Col = 2
                Select Case rsSab("SSISAAYFCT")
                    Case "VU ": fgCompteH.CellBackColor = mColor_G1
                    Case Else: fgCompteH.CellBackColor = mColor_W0
                End Select
End Select

fgCompteH.Col = 2: fgCompteH.Text = rsSab("SSISAAYFCT") & " - " & Trim(rsSab("SSISAAYUSR")) _
                & " " & dateImp10_S(rsSab("SSISAAYAMJ")) & " " & timeImp8(rsSab("SSISAAYHMS")) _
                & " / " & rsSab("SSISAAYVER")

fgCompteH.Col = 3: fgCompteH.Text = rsSab("SSISAAPRFX")
fgCompteH.CellFontSize = 8


End Sub

Public Sub fgCompteH_Display_WIN_Line()
On Error Resume Next
fgCompteH.Col = 0: fgCompteH.Text = rsSab("SSIWINNAT") & "|" & rsSab("SSIWINGUID") & "|" & rsSab("SSIWINYVER")
fgCompteH.Col = 1: fgCompteH.Text = rsSab("SSIWINPRFK")
Select Case rsSab("SSIWINPRFK")
    Case " ": fgCompteH.CellBackColor = mColor_G1
    Case Else: fgCompteH.CellBackColor = mColor_W0
                fgCompteH.Col = 2
                Select Case rsSab("SSIWINYFCT")
                    Case "VU ": fgCompteH.CellBackColor = mColor_G1
                    Case Else: fgCompteH.CellBackColor = mColor_W0
                End Select
End Select

fgCompteH.Col = 2: fgCompteH.Text = rsSab("SSIWINYFCT") & " - " & Trim(rsSab("SSIWINYUSR")) _
                & " " & dateImp10_S(rsSab("SSIWINYAMJ")) & " " & timeImp8(rsSab("SSIWINYHMS")) _
                & " / " & rsSab("SSIWINYVER")

fgCompteH.Col = 3: fgCompteH.Text = rsSab("SSIWINPRFX")
fgCompteH.CellFontSize = 8


End Sub


Public Sub fgCompteH_Display_DIV_Line()
On Error Resume Next
fgCompteH.Col = 0: fgCompteH.Text = rsSab("SSIDIVNAT") & "|" & rsSab("SSIDIVUIDX") & "|" & rsSab("SSIDIVUIDD") & "|" & rsSab("SSIDIVYVER") & "|"
fgCompteH.Col = 1: fgCompteH.Text = rsSab("SSIDIVPRFK")
Select Case rsSab("SSIDIVPRFK")
    Case " ": fgCompteH.CellBackColor = mColor_G1
    Case Else: fgCompteH.CellBackColor = mColor_W0
                fgCompteH.Col = 2
                Select Case rsSab("SSIDIVYFCT")
                    Case "VU ": fgCompteH.CellBackColor = mColor_G1
                    Case Else: fgCompteH.CellBackColor = mColor_W0
                End Select
End Select

fgCompteH.Col = 2: fgCompteH.Text = rsSab("SSIDIVYFCT") & " - " & Trim(rsSab("SSIDIVYUSR")) _
                & " " & dateImp10_S(rsSab("SSIDIVYAMJ")) & " " & timeImp8(rsSab("SSIDIVYHMS")) _
                & " / " & rsSab("SSIDIVYVER")

fgCompteH.Col = 3: fgCompteH.Text = rsSab("SSIDIVPRFX")
fgCompteH.CellFontSize = 8


End Sub

Public Sub fgCompteH_Display_MEL_Line()
On Error Resume Next
fgCompteH.Col = 0: fgCompteH.Text = rsSab("SSIMELNAT") & "|" & rsSab("SSIMELUIDX") & "|" & rsSab("SSIMELUIDD") & "|" & rsSab("SSIMELYVER") & "|"
fgCompteH.Col = 1: fgCompteH.Text = rsSab("SSIMELPRFK")
Select Case rsSab("SSIMELPRFK")
    Case " ": fgCompteH.CellBackColor = mColor_G1
    Case Else: fgCompteH.CellBackColor = mColor_W0
                fgCompteH.Col = 2
                Select Case rsSab("SSIMELYFCT")
                    Case "VU ": fgCompteH.CellBackColor = mColor_G1
                    Case Else: fgCompteH.CellBackColor = mColor_W0
                End Select
End Select

fgCompteH.Col = 2: fgCompteH.Text = rsSab("SSIMELYFCT") & " - " & Trim(rsSab("SSIMELYUSR")) _
                & " " & dateImp10_S(rsSab("SSIMELYAMJ")) & " " & timeImp8(rsSab("SSIMELYHMS")) _
                & " / " & rsSab("SSIMELYVER")

fgCompteH.Col = 3: fgCompteH.Text = rsSab("SSIMELPRFX")
fgCompteH.CellFontSize = 8


End Sub

Public Sub fgCompteH_Display_TIC_Line()
On Error Resume Next
fgCompteH.Col = 0: fgCompteH.Text = rsSab("SSITICNAT") & "|" & rsSab("SSITICUIDX") & "|" & rsSab("SSITICUIDD") & "|" & rsSab("SSITICYVER") & "|"
fgCompteH.Col = 1: fgCompteH.Text = rsSab("SSITICPRFK")
Select Case rsSab("SSITICPRFK")
    Case " ": fgCompteH.CellBackColor = mColor_G1
    Case Else: fgCompteH.CellBackColor = mColor_W0
                fgCompteH.Col = 2
                Select Case rsSab("SSITICYFCT")
                    Case "VU ": fgCompteH.CellBackColor = mColor_G1
                    Case Else: fgCompteH.CellBackColor = mColor_W0
                End Select
End Select

fgCompteH.Col = 2: fgCompteH.Text = rsSab("SSITICYFCT") & " - " & Trim(rsSab("SSITICYUSR")) _
                & " " & dateImp10_S(rsSab("SSITICYAMJ")) & " " & timeImp8(rsSab("SSITICYHMS")) _
                & " / " & rsSab("SSITICYVER")

fgCompteH.Col = 3: fgCompteH.Text = rsSab("SSITICPRFX")
fgCompteH.CellFontSize = 8


End Sub


Public Sub fgCompteH_Display_SAB_Line()
On Error Resume Next
fgCompteH.Col = 0: fgCompteH.Text = rsSab("SSISABNAT") & "|" & rsSab("SSISABUIDX") & "|" & rsSab("SSISABYVER")
fgCompteH.Col = 1: fgCompteH.Text = rsSab("SSISABPRFK")
Select Case rsSab("SSISABPRFK")
    Case " ": fgCompteH.CellBackColor = mColor_G1
    Case Else: fgCompteH.CellBackColor = mColor_W0
                fgCompteH.Col = 2
                Select Case rsSab("SSISABYFCT")
                    Case "VU ": fgCompteH.CellBackColor = mColor_G1
                    Case Else: fgCompteH.CellBackColor = mColor_W0
                End Select
End Select

fgCompteH.Col = 2: fgCompteH.Text = rsSab("SSISABYFCT") & " - " & Trim(rsSab("SSISABYUSR")) _
                & " " & dateImp10_S(rsSab("SSISABYAMJ")) & " " & timeImp8(rsSab("SSISABYHMS")) _
                & " / " & rsSab("SSISABYVER")

fgCompteH.Col = 3: fgCompteH.Text = rsSab("SSISABPRFX")
fgCompteH.CellFontSize = 8


End Sub

Public Sub fgCompteH_Display_IBM_Line()
On Error Resume Next
fgCompteH.Col = 0: fgCompteH.Text = rsSab("SSIIBMNAT") & "|" & rsSab("SSIIBMUIDD") & "|" & rsSab("SSIIBMYVER")
fgCompteH.Col = 1: fgCompteH.Text = rsSab("SSIIBMPRFK")
Select Case rsSab("SSIIBMPRFK")
    Case " ": fgCompteH.CellBackColor = mColor_G1
    Case Else: fgCompteH.CellBackColor = mColor_W0
                fgCompteH.Col = 2
                Select Case rsSab("SSIIBMYFCT")
                    Case "VU ": fgCompteH.CellBackColor = mColor_G1
                    Case Else: fgCompteH.CellBackColor = mColor_W0
                End Select
End Select

fgCompteH.Col = 2: fgCompteH.Text = rsSab("SSIIBMYFCT") & " - " & Trim(rsSab("SSIIBMYUSR")) _
                & " " & dateImp10_S(rsSab("SSIIBMYAMJ")) & " " & timeImp8(rsSab("SSIIBMYHMS")) _
                & " / " & rsSab("SSIIBMYVER")

fgCompteH.Col = 3: fgCompteH.Text = rsSab("UPTEXT") '"pgm : " & rsSab("UPINPG") & " / jobd : " & rsSab("UPJBDS") & " / groupe :" & rsSab("UPGRPF")
fgCompteH.CellFontSize = 8
End Sub

Public Sub fgCompte_Reset()
fgCompte.Clear
fgCompte_Sort1 = 0: fgCompte_Sort2 = 0
fgCompte_Sort1_Old = -1
fgCompte_RowDisplay = 0: fgCompte_RowClick = 0
fgCompte_arrIndex = fgCompte.Cols - 1
blnfgCompte_DisplayLine = False
fgCompte_SortAD = 6
fgCompte.LeftCol = fgCompte.FixedCols

End Sub

Public Sub fgCompte_Sort()

If fgCompte.Rows > 1 Then
    fgCompte.Row = 1
    fgCompte.RowSel = fgCompte.Rows - 1
    
    If fgCompte_Sort1_Old = fgCompte_Sort1 Then
        If fgCompte_SortAD = 5 Then
            fgCompte_SortAD = 6
        Else
            fgCompte_SortAD = 5
        End If
    Else
        fgCompte_SortAD = 5
    End If
    fgCompte_Sort1_Old = fgCompte_Sort1
    
    fgCompte.Col = fgCompte_Sort1
    fgCompte.ColSel = fgCompte_Sort2
    fgCompte.Sort = fgCompte_SortAD
End If

End Sub

Public Sub fgCompte_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgCompte.Rows - 1
    fgCompte.Row = I
    fgCompte.Col = lK
    Select Case lK
'        Case 3: fgCompte.Col = 3: X = Format$(Val(fgCompte.Text), "000000000000000.00")

    End Select
    fgCompte.Col = fgCompte_arrIndex - 1
    fgCompte.Text = X
Next I

fgCompte_Sort1 = fgCompte_arrIndex - 1: fgCompte_Sort2 = fgCompte_arrIndex - 1
fgCompte_Sort
End Sub






Public Sub fgProfil_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgProfil.Visible = False
mRow = fgProfil.Row

If lRow > 0 And lRow < fgProfil.Rows Then
    fgProfil.Row = lRow
    For I = 2 To 0 Step -1
        fgProfil.Col = I: fgProfil.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgProfil.Row = mRow
    If fgProfil.Row > 0 Then
        lRow = fgProfil.Row
        lColor_Old = fgProfil.CellBackColor
        For I = 2 To 0 Step -1
          fgProfil.Col = I: fgProfil.CellBackColor = lColor
        Next I
    End If
End If
fgProfil.LeftCol = fgProfil.FixedCols
fgProfil.Visible = True
End Sub


Private Sub fgProfil_Display_IBM()

Dim xSQL As String

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgProfil_Display_IBM"
fraProfil_Update.Visible = False
fraProfil_Update_DIV.Visible = False
fgProfil.Visible = False
fgProfil_Reset
fgProfil.FormatString = fgProfil_FormatString
fgProfil.Rows = 1
fgProfil.Row = 0

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0 " _
     & " where SSIIBMNAT = '$'" _
     & " order by UPUPRF"

Set rsSab = cnsab.Execute(xSQL)
  

Do While Not rsSab.EOF
    fgProfil.Rows = fgProfil.Rows + 1
    fgProfil.Row = fgProfil.Rows - 1
    fgProfil_Display_IBM_Line
    rsSab.MoveNext
Loop

fgProfil.Visible = True: fgProfil.Enabled = True

If cmdSelect_SQL_K = "2_D" Then
    cmdProfil_New.Visible = arrHab(3): cmdProfil_Print.Visible = arrHab(3) ': cmdProfil_Excel.Visible = arrHab(3)
Else
    cmdProfil_New.Visible = False: cmdProfil_Print.Visible = False ': cmdProfil_Excel.Visible = False
End If


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgProfil.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgProfil_Display_SAA()

Dim xSQL As String, xWhere As String

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgProfil_Display_SAA"
fraProfil_Update.Visible = False
fraProfil_Update_DIV.Visible = False
fgProfil.Visible = False
fgProfil_Reset
fgProfil.FormatString = fgProfil_FormatString
fgProfil.Rows = 1
fgProfil.Row = 0
If chkProfil_DOM <> "1" Then xWhere = " and SSISAASTAK = ' '"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
     & " where SSISAANAT = '$' and SSISAAUSEQ = 0" & xWhere _
     & " order by SSISAAUIDX"

Set rsSab = cnsab.Execute(xSQL)
  

Do While Not rsSab.EOF
    fgProfil.Rows = fgProfil.Rows + 1
    fgProfil.Row = fgProfil.Rows - 1
    fgProfil_Display_SAA_Line
    rsSab.MoveNext
Loop

fgProfil.Visible = True: fgProfil.Enabled = True

If cmdSelect_SQL_K = "2_D" Then
    cmdProfil_Print.Visible = arrHab(3) ': cmdProfil_Excel.Visible = arrHab(3)
Else
    cmdProfil_Print.Visible = False ': cmdProfil_Excel.Visible = False
End If


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgProfil.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgProfil_Display_WIN()

Dim xSQL As String, xWhere As String

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgProfil_Display_WIN"
fraProfil_Update.Visible = False
fraProfil_Update_DIV.Visible = False
fgProfil.Visible = False
fgProfil_Reset
fgProfil.FormatString = fgProfil_FormatString
fgProfil.Rows = 1
fgProfil.Row = 0
If chkProfil_DOM <> "1" Then xWhere = " and SSIWINSTAK = ' '"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
     & " where SSIWINNAT = '$'" & xWhere _
     & " order by SSIWINPRFX"

Set rsSab = cnsab.Execute(xSQL)
  

Do While Not rsSab.EOF
    fgProfil.Rows = fgProfil.Rows + 1
    fgProfil.Row = fgProfil.Rows - 1
    fgProfil_Display_WIN_Line
    rsSab.MoveNext
Loop

fgProfil.Visible = True: fgProfil.Enabled = True

If cmdSelect_SQL_K = "2_D" Then
    cmdProfil_Print.Visible = arrHab(3): cmdProfil_Excel.Visible = False
Else
    cmdProfil_Print.Visible = False: cmdProfil_Excel.Visible = False
End If


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgProfil.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgProfil_Display_MEL()

Dim xSQL As String, xWhere As String

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgProfil_Display_MEL"
fraProfil_Update.Visible = False
fraProfil_Update_DIV.Visible = False
fgProfil.Visible = False
fgProfil_Reset
fgProfil.FormatString = fgProfil_FormatString
fgProfil.Rows = 1
fgProfil.Row = 0
If chkProfil_DOM <> "1" Then xWhere = " and SSIMELSTAK = ' '"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0 " _
     & " where SSIMELNAT = '$'" & xWhere _
     & " order by SSIMELPRFX"

Set rsSab = cnsab.Execute(xSQL)
  

Do While Not rsSab.EOF
    fgProfil.Rows = fgProfil.Rows + 1
    fgProfil.Row = fgProfil.Rows - 1
    fgProfil_Display_MEL_Line
    rsSab.MoveNext
Loop

fgProfil.Visible = True: fgProfil.Enabled = True

If cmdSelect_SQL_K = "2_D" Then
    cmdProfil_Print.Visible = arrHab(3): cmdProfil_Excel.Visible = False
Else
    cmdProfil_Print.Visible = False: cmdProfil_Excel.Visible = False
End If


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgProfil.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgProfil_Display_TIC()

Dim xSQL As String, xWhere As String

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgProfil_Display_TIC"
fraProfil_Update.Visible = False
fraProfil_Update_DIV.Visible = False
fgProfil.Visible = False
fgProfil_Reset
fgProfil.FormatString = fgProfil_FormatString
fgProfil.Rows = 1
fgProfil.Row = 0
If chkProfil_DOM <> "1" Then xWhere = " and SSITICSTAK = ' '"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSITIC0 " _
     & " where SSITICNAT = '$'" & xWhere _
     & " order by SSITICPRFX"

Set rsSab = cnsab.Execute(xSQL)
  

Do While Not rsSab.EOF
    fgProfil.Rows = fgProfil.Rows + 1
    fgProfil.Row = fgProfil.Rows - 1
    fgProfil_Display_TIC_Line
    rsSab.MoveNext
Loop

fgProfil.Visible = True: fgProfil.Enabled = True

If cmdSelect_SQL_K = "2_D" Then
    cmdProfil_Print.Visible = arrHab(3): cmdProfil_Excel.Visible = False
Else
    cmdProfil_Print.Visible = False: cmdProfil_Excel.Visible = False
End If


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgProfil.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub



Private Sub fgProfil_Display_DIV()

Dim xSQL As String, xWhere As String

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgProfil_Display_DIV"
fraProfil_Update.Visible = False
fraProfil_Update_DIV.Visible = False
fgProfil.Visible = False
fgProfil_Reset
fgProfil.FormatString = fgProfil_FormatString
fgProfil.Rows = 1
fgProfil.Row = 0
If chkProfil_DOM <> "1" Then xWhere = " and SSIDIVSTAK = ' '"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
     & " where SSIDIVNAT = '$'" & xWhere _
     & " order by SSIDIVUIDX"

Set rsSab = cnsab.Execute(xSQL)
  

Do While Not rsSab.EOF
    fgProfil.Rows = fgProfil.Rows + 1
    fgProfil.Row = fgProfil.Rows - 1
    fgProfil_Display_DIV_Line
    rsSab.MoveNext
Loop

fgProfil.Visible = True: fgProfil.Enabled = True

If cmdSelect_SQL_K = "2_D" Then
    cmdProfil_New.Visible = arrHab(3) Or arrHab(5)
    cmdProfil_Print.Visible = arrHab(3) Or arrHab(5): cmdProfil_Excel.Visible = False
Else
    cmdProfil_Print.Visible = False: cmdProfil_Excel.Visible = False
End If


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgProfil.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub





Public Sub fgProfil_Display_IBM_Line()
On Error Resume Next

fgProfil.Col = 0: fgProfil.Text = rsSab("SSIIBMUIDD")
fgProfil.Col = 1: fgProfil.Text = rsSab("UPUPRF")
fgProfil.CellFontBold = True: fgProfil.CellForeColor = vbBlue
fgProfil.Col = 2: fgProfil.Text = rsSab("UPTEXT")

End Sub

Public Sub fgProfil_Display_SAA_Line()
On Error Resume Next

'fgProfil.Col = 0: fgProfil.Text = rsSab("SSISAAUIDD")
fgProfil.Col = 1: fgProfil.Text = rsSab("SSISAAUIDX")
fgProfil.CellFontBold = True: fgProfil.CellForeColor = vbBlue
If rsSab("SSISAASTAK") <> " " Then fgProfil.CellBackColor = RGB(192, 192, 192)
fgProfil.Col = 2: fgProfil.Text = rsSab("SSISAAUNOM")

End Sub


Public Sub fgProfil_Display_WIN_Line()
On Error Resume Next

fgProfil.Col = 0: fgProfil.Text = rsSab("SSIWINUIDD")
fgProfil.Col = 1: fgProfil.Text = rsSab("SSIWINUIDX")
fgProfil.CellFontBold = True: fgProfil.CellForeColor = vbBlue
If rsSab("SSIWINSTAK") <> " " Then fgProfil.CellBackColor = RGB(192, 192, 192)
fgProfil.Col = 2: fgProfil.Text = rsSab("SSIWINUNOM")
fgProfil.Col = 3: fgProfil.Text = rsSab("SSIWINGUID")

End Sub

Public Sub fgProfil_Display_MEL_Line()
On Error Resume Next

fgProfil.Col = 0: fgProfil.Text = rsSab("SSIMELUIDD")
fgProfil.Col = 1: fgProfil.Text = rsSab("SSIMELUIDX")
fgProfil.CellFontBold = True: fgProfil.CellForeColor = vbBlue
If rsSab("SSIMELSTAK") <> " " Then fgProfil.CellBackColor = RGB(192, 192, 192)
fgProfil.Col = 2: fgProfil.Text = rsSab("SSIMELUNOM")
fgProfil.Col = 3: fgProfil.Text = rsSab("SSIMELGUID")

End Sub


Public Sub fgProfil_Display_TIC_Line()
On Error Resume Next

fgProfil.Col = 0: fgProfil.Text = rsSab("SSITICUIDD")
fgProfil.Col = 1: fgProfil.Text = rsSab("SSITICUIDX")
fgProfil.CellFontBold = True: fgProfil.CellForeColor = vbBlue
If rsSab("SSITICSTAK") <> " " Then fgProfil.CellBackColor = RGB(192, 192, 192)
fgProfil.Col = 2: fgProfil.Text = rsSab("SSITICUNOM")
fgProfil.Col = 3: fgProfil.Text = rsSab("SSITICGUID")

End Sub




Public Sub fgProfil_Display_DIV_Line()
On Error Resume Next

fgProfil.Col = 0: fgProfil.Text = rsSab("SSIDIVUIDD")
fgProfil.Col = 1: fgProfil.Text = rsSab("SSIDIVUIDX")
fgProfil.CellFontBold = True: fgProfil.CellForeColor = vbBlue
If rsSab("SSIDIVSTAK") <> " " Then fgProfil.CellBackColor = RGB(192, 192, 192)
fgProfil.Col = 2: fgProfil.Text = Trim(rsSab("SSIDIVUNOM"))
fgProfil.Col = 3: fgProfil.Text = Trim(rsSab("SSIDIVPRFX"))

End Sub




Public Sub fgProfil_Reset()
fgProfil.Clear
fgProfil_Sort1 = 0: fgProfil_Sort2 = 0
fgProfil_Sort1_Old = -1
fgProfil_RowDisplay = 0: fgProfil_RowClick = 0
fgProfil_arrIndex = fgProfil.Cols - 1
blnfgProfil_DisplayLine = False
fgProfil_SortAD = 6
fgProfil.LeftCol = fgProfil.FixedCols

End Sub



Public Sub fgProfil_Sort()

If fgProfil.Rows > 1 Then
    fgProfil.Row = 1
    fgProfil.RowSel = fgProfil.Rows - 1
    
    If fgProfil_Sort1_Old = fgProfil_Sort1 Then
        If fgProfil_SortAD = 5 Then
            fgProfil_SortAD = 6
        Else
            fgProfil_SortAD = 5
        End If
    Else
        fgProfil_SortAD = 5
    End If
    fgProfil_Sort1_Old = fgProfil_Sort1
    
    fgProfil.Col = fgProfil_Sort1
    fgProfil.ColSel = fgProfil_Sort2
    fgProfil.Sort = fgProfil_SortAD
End If

End Sub

Public Sub fgProfil_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgProfil.Rows - 1
    fgProfil.Row = I
    fgProfil.Col = lK
    Select Case lK
        Case 0: fgProfil.Col = 0: X = Format$(Val(fgProfil.Text), "000000000000000")

    End Select
    fgProfil.Col = fgProfil_arrIndex - 1
    fgProfil.Text = X
Next I

fgProfil_Sort1 = fgProfil_arrIndex - 1: fgProfil_Sort2 = fgProfil_arrIndex - 1
fgProfil_Sort
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

Private Sub fgDetail_Display()
Dim X As String, xWhere As String
Dim xSQL As String

On Error GoTo Error_Handler

currentAction = currentAction & "-> fgDetail_Display"
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.Row = 0

Do While Not rsSab.EOF
    Call rsYSSIDOM0_GetBuffer(rsSab, xYSSIDOM0)
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_Display_Line
    rsSab.MoveNext
Loop


fgDetail.Visible = True

If cmdSelect_SQL_K = "2" And fgSelect.Visible Then
    Call fgDetail_Display_lstW
    cmdSSIUSR_Delete.Caption = "Supprimer ce modèle BIA"
    cmdSSIUSR_Delete.Visible = arrHab(3)
Else
    lstW.Visible = False
    cmdSSIUSR_Delete.Visible = False
End If

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgDetail.Rows - 1): DoEvents

'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub



Public Sub fgDetail_Display_Line()
On Error Resume Next
Dim K As Integer, wBackColor As Long, wForecolor As Long

Select Case Trim(xYSSIDOM0.SSIDOMDIDX)
    Case "IBM":  wBackColor = mColor_G0: wForecolor = RGB(0, 96, 0)
    Case "SAA":  wBackColor = mColor_Y2: wForecolor = RGB(255, 0, 255) ' vbMagenta
    Case "SAB":  wBackColor = mColor_B0: wForecolor = vbBlue 'RGB(0, 96, 0)
    Case "SAB_W":  wBackColor = mColor_B0: wForecolor = RGB(96, 0, 230)
    Case "WIN": wBackColor = mColor_Y2: wForecolor = RGB(255, 32, 0)
    Case "DIV":  wBackColor = mColor_Y2: wForecolor = RGB(255, 80, 0) 'RGB(139, 0, 139)
    Case "MEL": wBackColor = mColor_Y2: wForecolor = RGB(220, 96, 96) ' RGB(160, 44, 220) '
    Case "TIC": wBackColor = mColor_Y2: wForecolor = RGB(96, 0, 96) '
End Select

fgDetail.Col = 0: fgDetail.Text = " " & xYSSIDOM0.SSIDOMDIDX: fgDetail.CellFontBold = True
fgDetail.CellForeColor = wForecolor
fgDetail.Col = 1: 'fgDetail.CellBackColor = wForecolor
If Len(xYSSIDOM0.SSIDOMUNIT) = 3 Then
    K = Val(Mid$(xYSSIDOM0.SSIDOMUNIT, 2, 2))
    fgDetail.Text = xYSSIDOM0.SSIDOMUNIT & " : " & arrSSIUSRUNIT_Code(K)
    'fgDetail.CellFontSize = 8
    If xYSSIDOM0.SSIDOMUNIT <> xYSSIUSR0.SSIUSRUNIT Then
        fgDetail.CellBackColor = mColor_Y2
        fgDetail.CellForeColor = wForecolor
    Else
        fgDetail.CellForeColor = RGB(128, 128, 128)
    End If
End If

If xYSSIDOM0.SSIDOMSTAK = "N" Then
    fgDetail.Col = 0: fgDetail.CellBackColor = RGB(230, 230, 230)
    For K = 2 To 9: fgDetail.Col = K: fgDetail.CellBackColor = RGB(230, 230, 230): Next K
    wBackColor = RGB(96, 96, 96)
End If
For K = 2 To 9: fgDetail.Col = K: fgDetail.CellForeColor = wForecolor: Next K

fgDetail.Col = 2: fgDetail.CellFontBold = True
fgDetail.CellForeColor = wForecolor
If Trim(xYSSIDOM0.SSIDOMUIDX) = "" Then
    fgDetail.CellBackColor = mColor_Y2
Else
    fgDetail.Text = xYSSIDOM0.SSIDOMUIDX
End If

fgDetail.Col = 3: fgDetail.Text = xYSSIDOM0.SSIDOMUIDD
fgDetail.Col = 4: fgDetail.Text = xYSSIDOM0.SSIDOMPRFX
    'fgDetail.CellForeColor = fgDetail.ForeColorFixed
fgDetail.Col = 5: fgDetail.Text = xYSSIDOM0.SSIDOMSTAK
If xYSSIDOM0.SSIDOMDECH <> 0 Then
    fgDetail.Col = 6: fgDetail.Text = dateImp10_S(xYSSIDOM0.SSIDOMDECH)
End If
If xYSSIDOM0.SSIDOMPRFD = 0 Then
    fgDetail.Col = 7: fgDetail.Text = xYSSIDOM0.SSIDOMPRFK & " à contrôler"
Else
    fgDetail.Col = 7: fgDetail.Text = xYSSIDOM0.SSIDOMPRFK & " " & dateImp10_S(xYSSIDOM0.SSIDOMPRFD)
End If
Select Case xYSSIDOM0.SSIDOMPRFK
    Case " "
    Case "N": fgDetail.CellBackColor = mColor_W1
    Case "X": fgDetail.CellBackColor = RGB(192, 192, 192)
    Case "!": fgDetail.CellBackColor = mColor_Y1
    Case Else: fgDetail.CellBackColor = mColor_Y2
End Select

If xYSSIDOM0.SSIDOMTLNK > 0 Then fgDetail.Col = 8: fgDetail.Text = xYSSIDOM0.SSIDOMTLNK
fgDetail.Col = 9: fgDetail.Text = xYSSIDOM0.SSIDOMYFCT & " : " & xYSSIDOM0.SSIDOMYUSR _
   & " " & dateImp10_S(xYSSIDOM0.SSIDOMYAMJ) & " " & timeImp8(xYSSIDOM0.SSIDOMYHMS)
fgDetail.CellForeColor = RGB(128, 128, 128)
fgDetail.CellFontSize = 8

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
On Error GoTo Error_Handler

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

fraSelect_Options.Visible = True
fraSelect_Options_1.Visible = True

Set fraSelect_Options_1.Container = fraSelect
fraSelect_Options_1.Top = fraSelect_Options.Top
fraSelect_Options_1.Left = fraSelect_Options.Left


Set fraSelect_Options_J.Container = fraSelect
fraSelect_Options_J.Top = fraSelect_Options.Top
fraSelect_Options_J.Left = fraSelect_Options.Left
Call DTPicker_Set(txtSelect_Options_J_SSITXTYMAJ, DSys)


Set fraSelect_Options_4.Container = fraSelect
fraSelect_Options_4.Top = fraSelect_Options.Top
fraSelect_Options_4.Left = fraSelect_Options.Left

fraDetail.Visible = False
Set fraDetail.Container = fraSelect
fraDetail.Top = fgSelect.Top
fraDetail.Height = fgSelect.Height
fraDetail.Left = fgSelect.Left + fgSelect.Width - fraDetail.Width + 100
fraDetail.ForeColor = vbRed
fraDetail_Update_PRF.BackColor = fraDetail.BackColor
fraDetail_Update_STAK.BackColor = fraDetail.BackColor
fraDetail_Update.BackColor = fraDetail.BackColor

fraDetail_Update_SRV.Visible = False
fraDetail_Update_SRV.BackColor = fraDetail.BackColor
fraDetail_Update_SRV.Left = fraDetail_Update_STAK.Left
fraDetail_Update_SRV.Top = fraDetail_Update_STAK.Top

fraProfil.Visible = False
Set fraProfil.Container = fraSelect
fraProfil.Top = fgSelect.Top
fraProfil.Height = fgSelect.Height
fraProfil.Left = fgSelect.Left + fgSelect.Width - fraProfil.Width + 100
fraProfil.ForeColor = vbMagenta
fgProfil_FormatString = fgProfil.FormatString
chkProfil_DOM.ForeColor = vbMagenta
chkProfil_DOM.Visible = False

cmdCompte_Val.Top = cmdProfil_New.Top
cmdCompte_Val.Left = cmdProfil_New.Left
cmdProfil_Change.Top = cmdProfil_Update.Top
cmdProfil_Change.Left = cmdProfil_Update.Left

fraYSSIDOM0.Visible = False
Set fraYSSIDOM0.Container = fraProfil
fraYSSIDOM0.Top = fgProfil.Top
fraYSSIDOM0.Left = fgProfil.Left
fraYSSIDOM0.ForeColor = vbMagenta

fraCompteH.Visible = False
Set fraCompteH.Container = fraProfil
fraCompteH.Top = 0 'fraProfil.Top
fraCompteH.Left = 0 'fraProfil.Left
fraCompteH.ForeColor = vbRed
fgCompteH_FormatString = fgCompteH.FormatString
lblCompteH_SSITXTINFO.ForeColor = vbMagenta
'___________________________________________________________________

fgCompte_FormatString = fgCompte.FormatString

txtRTF.Visible = False
Set txtRTF.Container = fraSelect
txtRTF.Top = fgSelect.Top
txtRTF.Height = fgSelect.Height
txtRTF.Left = fgSelect.Left

fraProfil_Update_DIV.Visible = False
Set fraProfil_Update_DIV.Container = fraProfil
fraProfil_Update_DIV.Top = fraProfil_Update.Top
fraProfil_Update_DIV.Left = fraProfil_Update.Left

fraYSSIDIV0.Visible = False
Set fraYSSIDIV0.Container = fraProfil
fraYSSIDIV0.Top = fraProfil_Update.Top
fraYSSIDIV0.Left = fraProfil_Update.Left
'

'txtRTF.LoadFile (paramEditionFiligrane_Folder & "\VB_RTF_Modèle.rtf")
'txtRTF.LoadFile ("c:\temp\VB_RTF_Modèle.rtf")
X = paramFile("BIA", 1): X = paramServer(X)
txtRTF.LoadFile X

VB_RTF_Modèle = txtRTF.TextRTF

lstW.Visible = False
Set lstW.Container = fraSelect
lstW.Top = fgSelect.Top
lstW.Height = fgSelect.Height
lstW.Left = fgSelect.Left + fgSelect.Width - lstW.Width - 200

'_________________________________________________________________________________________
ReDim arrSSIUSRSTAK(2): arrSSIUSRSTAK_UB = 2
arrSSIUSRSTAK(0) = " Oui": cboSSIUSRSTAK.AddItem arrSSIUSRSTAK(0): cboSSIDOMSTAK.AddItem arrSSIUSRSTAK(0)
arrSSIUSRSTAK(1) = "Non": cboSSIUSRSTAK.AddItem arrSSIUSRSTAK(1): cboSSIDOMSTAK.AddItem arrSSIUSRSTAK(1)

cboSelect_Options_1_SSIUSRSTAK.AddItem arrSSIUSRSTAK(0)
cboSelect_Options_1_SSIUSRSTAK.AddItem arrSSIUSRSTAK(1)
cboSelect_Options_1_SSIUSRSTAK.AddItem "* tous"
cboSelect_Options_1_SSIUSRSTAK.ListIndex = 0

cboSelect_Options_4_SSIDOMSTAK.AddItem "* tous"
cboSelect_Options_4_SSIDOMSTAK.AddItem arrSSIUSRSTAK(0)
cboSelect_Options_4_SSIDOMSTAK.AddItem arrSSIUSRSTAK(1)
cboSelect_Options_4_SSIDOMSTAK.ListIndex = 0

ReDim arrSSIUSRPRFK(5): arrSSIUSRPRFK_UB = 4
arrSSIUSRPRFK(0) = " Oui": cboSSIUSRPRFK.AddItem arrSSIUSRPRFK(0): cboSSIDOMPRFK.AddItem arrSSIUSRPRFK(0)
arrSSIUSRPRFK(1) = "Non": cboSSIUSRPRFK.AddItem arrSSIUSRPRFK(1): cboSSIDOMPRFK.AddItem arrSSIUSRPRFK(1)
arrSSIUSRPRFK(2) = "? en attente": cboSSIUSRPRFK.AddItem arrSSIUSRPRFK(2): cboSSIDOMPRFK.AddItem arrSSIUSRPRFK(2)
arrSSIUSRPRFK(3) = "! échéance": cboSSIUSRPRFK.AddItem arrSSIUSRPRFK(3): cboSSIDOMPRFK.AddItem arrSSIUSRPRFK(3)
arrSSIUSRPRFK(4) = "X exit_grp": cboSSIUSRPRFK.AddItem arrSSIUSRPRFK(4): cboSSIDOMPRFK.AddItem arrSSIUSRPRFK(4)
'arrSSIUSRPRFK(5) = "$ système": cboSSIUSRPRFK.AddItem arrSSIUSRPRFK(5): cboSSIDOMPRFK.AddItem arrSSIUSRPRFK(5)
'__________________________________________________________________________________________________

Call paramSSIUSRPRFX_Load


'__________________________________________________________________________________________________

ReDim arrSSIDOMDIDX(9): arrSSIDOMDIDX_UB = 9
arrSSIDOMDIDX(0) = "": cboProfil_DOM.AddItem arrSSIDOMDIDX(0)
cboSelect_Options_1_SSIDOMDIDX.AddItem arrSSIDOMDIDX(0)
cboSelect_Options_J_SSIDOMDIDX.AddItem arrSSIDOMDIDX(0)
arrSSIDOMDIDX(1) = "IBM": cboProfil_DOM.AddItem arrSSIDOMDIDX(1)
cboSelect_Options_1_SSIDOMDIDX.AddItem arrSSIDOMDIDX(1)
cboSelect_Options_J_SSIDOMDIDX.AddItem arrSSIDOMDIDX(1)
arrSSIDOMDIDX(2) = "SAA": cboProfil_DOM.AddItem arrSSIDOMDIDX(2)
cboSelect_Options_1_SSIDOMDIDX.AddItem arrSSIDOMDIDX(2)
cboSelect_Options_J_SSIDOMDIDX.AddItem arrSSIDOMDIDX(2)
arrSSIDOMDIDX(3) = "SAB": cboProfil_DOM.AddItem arrSSIDOMDIDX(3)
cboSelect_Options_1_SSIDOMDIDX.AddItem arrSSIDOMDIDX(3)
cboSelect_Options_J_SSIDOMDIDX.AddItem arrSSIDOMDIDX(3)
arrSSIDOMDIDX(4) = "WIN": cboProfil_DOM.AddItem arrSSIDOMDIDX(4)
cboSelect_Options_1_SSIDOMDIDX.AddItem arrSSIDOMDIDX(4)
cboSelect_Options_J_SSIDOMDIDX.AddItem arrSSIDOMDIDX(4)
arrSSIDOMDIDX(5) = "DIV": cboProfil_DOM.AddItem arrSSIDOMDIDX(5)
cboSelect_Options_1_SSIDOMDIDX.AddItem arrSSIDOMDIDX(5)
cboSelect_Options_J_SSIDOMDIDX.AddItem arrSSIDOMDIDX(5)
arrSSIDOMDIDX(6) = "MEL": cboProfil_DOM.AddItem arrSSIDOMDIDX(6)
cboSelect_Options_1_SSIDOMDIDX.AddItem arrSSIDOMDIDX(6)
cboSelect_Options_J_SSIDOMDIDX.AddItem arrSSIDOMDIDX(6)
arrSSIDOMDIDX(7) = "TIC": cboProfil_DOM.AddItem arrSSIDOMDIDX(7)
cboSelect_Options_1_SSIDOMDIDX.AddItem arrSSIDOMDIDX(7)
cboSelect_Options_J_SSIDOMDIDX.AddItem arrSSIDOMDIDX(7)
arrSSIDOMDIDX(8) = "SAB_W": cboProfil_DOM.AddItem arrSSIDOMDIDX(8)
cboSelect_Options_1_SSIDOMDIDX.AddItem arrSSIDOMDIDX(8)
cboSelect_Options_J_SSIDOMDIDX.AddItem arrSSIDOMDIDX(8)

cboSelect_Options_1_SSIDOMDIDX.ListIndex = 0
cboSelect_Options_J_SSIDOMDIDX.ListIndex = 0

cboSelect_Options_4_SSIDOMDIDX.AddItem arrSSIDOMDIDX(3)
cboSelect_Options_4_SSIDOMDIDX.AddItem arrSSIDOMDIDX(2)
cboSelect_Options_4_SSIDOMDIDX.AddItem arrSSIDOMDIDX(1)
cboSelect_Options_4_SSIDOMDIDX.AddItem "USR"
cboSelect_Options_4_SSIDOMDIDX.AddItem "DOM"
cboSelect_Options_4_SSIDOMDIDX.AddItem arrSSIDOMDIDX(4)
cboSelect_Options_4_SSIDOMDIDX.AddItem arrSSIDOMDIDX(5)
cboSelect_Options_4_SSIDOMDIDX.AddItem arrSSIDOMDIDX(6)
cboSelect_Options_4_SSIDOMDIDX.AddItem arrSSIDOMDIDX(7)
cboSelect_Options_4_SSIDOMDIDX.AddItem arrSSIDOMDIDX(8)
cboSelect_Options_4_SSIDOMDIDX.ListIndex = 0
chkSelect_Options_4_SSIDOMDIDX.ForeColor = vbMagenta

'__________________________________________________________________________________________________
xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSIIBM0 " _
     & " where SSIIBMNAT = '$'"
Set rsSab = cnsab.Execute(xSQL)
arrSSIDOMPRFX_Nb = rsSab(0)
xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
     & " where SSISABNAT = '$'"
Set rsSab = cnsab.Execute(xSQL)
arrSSIDOMPRFX_Nb = arrSSIDOMPRFX_Nb + rsSab(0)
xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
     & " where SSIWINNAT = '$'"
Set rsSab = cnsab.Execute(xSQL)
arrSSIDOMPRFX_Nb = arrSSIDOMPRFX_Nb + rsSab(0) + 3 '"SAB_W"

xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
     & " where SSIDIVNAT = '$'"
Set rsSab = cnsab.Execute(xSQL)
arrSSIDOMPRFX_Nb = arrSSIDOMPRFX_Nb + rsSab(0)
ReDim arrSSIDOMPRFX_D(arrSSIDOMPRFX_Nb + 1), arrSSIDOMPRFX_P(arrSSIDOMPRFX_Nb + 1)

xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSIMEL0 " _
     & " where SSIMELNAT = '$'"
Set rsSab = cnsab.Execute(xSQL)
arrSSIDOMPRFX_Nb = arrSSIDOMPRFX_Nb + rsSab(0)
ReDim arrSSIDOMPRFX_D(arrSSIDOMPRFX_Nb + 1), arrSSIDOMPRFX_P(arrSSIDOMPRFX_Nb + 1)

xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSITIC0 " _
     & " where SSITICNAT = '$'"
Set rsSab = cnsab.Execute(xSQL)
arrSSIDOMPRFX_Nb = arrSSIDOMPRFX_Nb + rsSab(0)
ReDim arrSSIDOMPRFX_D(arrSSIDOMPRFX_Nb + 1), arrSSIDOMPRFX_P(arrSSIDOMPRFX_Nb + 1)

xSQL = "select UPUPRF from " & paramIBM_Library_SABSPE & ".YSSIIBM0 " _
     & " where SSIIBMNAT = '$'" _
     & " order by UPUPRF"

Set rsSab = cnsab.Execute(xSQL)
arrSSIDOMPRFX_Nb = 0
Do While Not rsSab.EOF
    arrSSIDOMPRFX_Nb = arrSSIDOMPRFX_Nb + 1
    arrSSIDOMPRFX_D(arrSSIDOMPRFX_Nb) = "IBM"
    arrSSIDOMPRFX_P(arrSSIDOMPRFX_Nb) = rsSab(0)
    rsSab.MoveNext
Loop

xSQL = "select SSISABUIDX from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
     & " where SSISABNAT = '$' and SSISABSTAK = ' '" _
     & " order by SSISABUIDX"

Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrSSIDOMPRFX_Nb = arrSSIDOMPRFX_Nb + 1
    arrSSIDOMPRFX_D(arrSSIDOMPRFX_Nb) = "SAB"
    arrSSIDOMPRFX_P(arrSSIDOMPRFX_Nb) = rsSab(0)
    rsSab.MoveNext
Loop
arrSSIDOMPRFX_Nb = arrSSIDOMPRFX_Nb + 1
arrSSIDOMPRFX_D(arrSSIDOMPRFX_Nb) = "SAB_W"
arrSSIDOMPRFX_P(arrSSIDOMPRFX_Nb) = "Vérification"
arrSSIDOMPRFX_Nb = arrSSIDOMPRFX_Nb + 1
arrSSIDOMPRFX_D(arrSSIDOMPRFX_Nb) = "SAB_W"
arrSSIDOMPRFX_P(arrSSIDOMPRFX_Nb) = "Validation unique"
arrSSIDOMPRFX_Nb = arrSSIDOMPRFX_Nb + 1
arrSSIDOMPRFX_D(arrSSIDOMPRFX_Nb) = "SAB_W"
arrSSIDOMPRFX_P(arrSSIDOMPRFX_Nb) = "Validation globale"

xSQL = "select SSIWINUIDX from " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
     & " where SSIWINNAT = '$' and SSIWINSTAK = ' '" _
     & " order by SSIWINPRFX"

Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrSSIDOMPRFX_Nb = arrSSIDOMPRFX_Nb + 1
    arrSSIDOMPRFX_D(arrSSIDOMPRFX_Nb) = "WIN"
    arrSSIDOMPRFX_P(arrSSIDOMPRFX_Nb) = rsSab(0)
    rsSab.MoveNext
Loop

xSQL = "select SSIDIVUIDX from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
     & " where SSIDIVNAT = '$' and SSIDIVSTAK = ' '" _
     & " order by SSIDIVUIDX"

Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrSSIDOMPRFX_Nb = arrSSIDOMPRFX_Nb + 1
    arrSSIDOMPRFX_D(arrSSIDOMPRFX_Nb) = "DIV"
    arrSSIDOMPRFX_P(arrSSIDOMPRFX_Nb) = rsSab(0)
    rsSab.MoveNext
Loop

xSQL = "select SSIMELUIDX from " & paramIBM_Library_SABSPE & ".YSSIMEL0 " _
     & " where SSIMELNAT = '$' and SSIMELSTAK = ' '" _
     & " order by SSIMELUIDX"

Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrSSIDOMPRFX_Nb = arrSSIDOMPRFX_Nb + 1
    arrSSIDOMPRFX_D(arrSSIDOMPRFX_Nb) = "MEL"
    arrSSIDOMPRFX_P(arrSSIDOMPRFX_Nb) = rsSab(0)
    rsSab.MoveNext
Loop

xSQL = "select SSITICUIDX from " & paramIBM_Library_SABSPE & ".YSSITIC0 " _
     & " where SSITICNAT = '$' and SSITICSTAK = ' '" _
     & " order by SSITICUIDX"

Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrSSIDOMPRFX_Nb = arrSSIDOMPRFX_Nb + 1
    arrSSIDOMPRFX_D(arrSSIDOMPRFX_Nb) = "TIC"
    arrSSIDOMPRFX_P(arrSSIDOMPRFX_Nb) = rsSab(0)
    rsSab.MoveNext
Loop

cboSelect_Options_1_SSIDOMPRFX.AddItem ""
cboSelect_Options_1_SSIDOMPRFX.ListIndex = 0
'__________________________________________________________________________________________________

xSQL = "select * from " & paramIBM_Library_SAB & ".ZMNURUTA order by MNURUTCUT desc"

Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    If arrMNURUTCUT_Nb = 0 Then
        arrMNURUTCUT_Nb = rsSab("MNURUTCUT")
        ReDim arrMNURUTCUT(arrMNURUTCUT_Nb + 1)
    End If
    arrMNURUTCUT(rsSab("MNURUTCUT")) = rsSab("MNURUTUTI")
    rsSab.MoveNext
Loop

For K = 1 To arrMNURUTCUT_Nb
    If arrMNURUTCUT(K) = "" Then arrMNURUTCUT(K) = "? " & K
Next K
'__________________________________________________________________________________________________
For K = 0 To 99
    arrMNURCLABR(K) = "? classe " & K
Next K
xSQL = "select * from " & paramIBM_Library_SAB & ".ZMNURCL0 "

Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrMNURCLABR(rsSab("MNURCLCLA")) = rsSab("MNURCLABR")
    rsSab.MoveNext
Loop



Call paramSAA_Load
'__________________________________________________________________________________________________

Call rsYSSITXT0_Init(newYSSITXT0_JRN)
newYSSITXT0_JRN.SSITXTNAT = "J"
arrJRN_Origine(1) = " _SSI Utilisateur"
arrJRN_Origine(2) = " _SSI Domaine"
arrJRN_Origine(3) = " _SSI Modèle BIA"
arrJRN_Origine(4) = " _SSI Modèle DOM"
arrJRN_Origine(5) = " _SSI Service"
arrJRN_Origine(6) = " _SSI Automate"

arrJRN_Origine(10) = " _IBM"
arrJRN_Origine(20) = " _SAA"
arrJRN_Origine(21) = " _SAA Utilisateur"
arrJRN_Origine(22) = "U_SAA Unit"
arrJRN_Origine(23) = "A_SAA Application"
arrJRN_Origine(24) = "F_SAA Fonction"
arrJRN_Origine(25) = "$_SAA Profil"
arrJRN_Origine(26) = "P_SAA Param profil"

arrJRN_Origine(29) = "W_SAB Hab Swift"
arrJRN_Origine(30) = " _SAB Utilisateur"
arrJRN_Origine(31) = "2_SAB GRP Menus"
arrJRN_Origine(32) = "3_SAB GRP Données"
arrJRN_Origine(33) = "4_SAB GRP Métiers"
arrJRN_Origine(34) = "C_SAB Hab Classes"
arrJRN_Origine(35) = "D_SAB Hab Services"
arrJRN_Origine(36) = "M_SAB Hab Options"
arrJRN_Origine(37) = "H_SAB Hab Lots"
arrJRN_Origine(38) = "$_SAB Profils"
arrJRN_Origine(39) = "G_SAB Hab Métiers"
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! en dur dans fgSelect_Display_H_SAB_Line
arrJRN_Origine(40) = " _WIN"
arrJRN_Origine(41) = " _WIN Utilisateur"
arrJRN_Origine(42) = "$_WIN profil"
arrJRN_Origine(46) = " _MEL Utilisateur"
arrJRN_Origine(44) = "$_MEL profil"
arrJRN_Origine(48) = "@_MEL Auto"
arrJRN_Origine(50) = " _DIV"
arrJRN_Origine(51) = " _DIV Utilisateur"
arrJRN_Origine(52) = "$_DIV Profil"
arrJRN_Origine(55) = " _TIC"
arrJRN_Origine(56) = " _TIC Utilisateur"
arrJRN_Origine(57) = "$_TIC Profil"
arrJRN_Origine(58) = "D_TIC Droits"
arrJRN_Origine(59) = "R_TIC Rôles"


blnControl = True

fraParam.Visible = arrHab(18)
lstParam_K1.AddItem "SAA"
lstParam_K1.AddItem "BIA"

paramUAC_Lib

cboSSIUSRUNIT_Load

lstParam_SSIMELNAT.AddItem "1-destinataires des traitements automatiques"
lstParam_SSIMELNAT.AddItem "2-destinataires par service de l'application BIA_GOS"
lstParam_SSIMELNAT.AddItem "3-destinataires par service des alertes SAA"
lstParam_SSIMELNAT.AddItem "4-destinataires complémentaires RCOM"
lstParam_SSIMELNAT.AddItem "5-destinataires des alertes DROPI"
lstParam_SSIMELNAT.AddItem "6-destinataires des états 'NoPaper' par service"

lstParam_SSIMELUNOM.BackColor = &HF0FFF0
lstParam_SSIMELUNOM.ForeColor = &H4000&
lstParam_SSIMELNAT.BackColor = &HC0E0FF
lstParam_SSIMELUIDX.BackColor = &HE0FFFF
txtParam_SSIMELINFO.ForeColor = vbMagenta
cmdSelect_Reset
Me.Enabled = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

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



Private Sub fgSelect_Display_1()

Dim K As Long, xSQL As String, blnRupture As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Rows = 1
fgSelect.Row = 0

oldYSSIUSR0.SSIUSRUIDN = 0
Do While Not rsSab.EOF
    Call rsYSSIUSR0_GetBuffer(rsSab, xYSSIUSR0)
    If oldYSSIUSR0.SSIUSRUIDN <> xYSSIUSR0.SSIUSRUIDN Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        If blnSelect_Options_1_SSIDOMPRFX Then
            For K = 0 To 8: fgSelect.Col = K
                fgSelect.CellBackColor = RGB(210, 250, 250): fgSelect.CellForeColor = vbBlue: fgSelect.CellFontBold = True
            Next K
        End If

        fgSelect_Display_1_Line
        oldYSSIUSR0.SSIUSRUIDN = xYSSIUSR0.SSIUSRUIDN
        fgSelect.Col = 9
        fgSelect.Text = xYSSIUSR0.SSIUSRNAT & "|" & xYSSIUSR0.SSIUSRUIDN & "| | | |"
    End If
    If blnSelect_Options_1_SSIDOMPRFX Then
        Call rsYSSIDOM0_GetBuffer(rsSab, xYSSIDOM0)
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
    
        fgSelect_Display_1_Line_YSSIDOM0
        fgSelect.Col = 9
        fgSelect.Text = xYSSIDOM0.SSIDOMNAT & "|" & xYSSIDOM0.SSIDOMUIDN & "|" & xYSSIDOM0.SSIDOMDIDX & "|" _
                      & xYSSIDOM0.SSIDOMUIDX & "|" & xYSSIDOM0.SSIDOMUIDD & "|"
    End If

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
Private Sub fgSelect_Display_J()

Dim K As Long, xSQL As String, blnRupture As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_J"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = "<Origine                             |<Eve     |<Identification                                                                                 |" _
     & "<Champ                     |" _
     & "<Libellé / Valeur modifiée                                                               |" _
     & "<Valeur initiale                                                                        |<Date de l'événement                              |||||"
    
fgSelect.Rows = 1
fgSelect.Row = 0

Do While Not rsSab.EOF
    Call rsYSSITXT0_GetBuffer(rsSab, xYSSITXT0)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1

    fgSelect_Display_J_Line

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


Private Sub fgSelect_Display_H_SAB()

Dim xSQL As String, xSQL_H As String, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_H_SAB"
fgSelect.Visible = False
fgSelect_Reset

fgSelect.FormatString = "<Nature                                  |<Identifiant                                                                   |" _
     & ">Lot |<Actif |" _
     & "<Libellé                                                                                |" _
     & "<Conforme? |<Profil                                     |Mise à jour par         le                                                 ||||"
fgSelect.Rows = 1
fgSelect.Row = 0

X = Mid$(cboSelect_Options_4_SSIDOMNAT, 1, 1)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSISABNAT = '" & X & "'"
If X <> "G" Then
'_______________________________________________________________________________________
    X = Trim(txtSelect_Options_4_SSIDOMUIDX)
    If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSISABUIDX like '%" & X & "%'"
    
    X = Trim(txtSelect_Options_4_SSIDOMUIDD)
    If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSISABUIDD = " & Val(X)
    
    If Not blnDisplay Then
        Call MsgBox("Précisez au moins un critère de recherche", vbExclamation, "4 - Détail")
        Exit Sub
    End If
    Select Case Mid$(cboSelect_Options_4_SSIDOMSTAK, 1, 1)
        Case " ": xSQL = xSQL & " and SSISABSTAK = ' '"
        Case "N": xSQL = xSQL & " and SSISABSTAK = 'N'"
    End Select
    
    xSQL = Replace(xSQL, "and", "where", 1, 1)
    
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0" & xSQL
    If chkSelect_Options_4_SSIDOMDIDX = "1" Then xSQL_H = " union " & Replace(xSQL, "YSSISAB0", "YSSISABH")
    
    X = xSQL & xSQL_H _
         & " order by SSISABNAT , SSISABUIDX , SSISABULOT , SSISABYVER"
    
    Set rsSab = cnsab.Execute(X)
    
    
    Do While Not rsSab.EOF
        Call rsYSSISAB0_GetBuffer(rsSab, xYSSISAB0)
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
    
        fgSelect_Display_H_SAB_Line
    
        rsSab.MoveNext
    Loop
Else
'_______________________________________________________________________________________
    xSQL = ""
    X = Trim(txtSelect_Options_4_SSIDOMUIDX)
    If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSISAMUIDX like '%" & X & "%'"
    
    X = Trim(txtSelect_Options_4_SSIDOMUIDD)
    If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSISAMUIDD = " & Val(X)
    
    If Not blnDisplay Then
        Call MsgBox("Précisez au moins un critère de recherche", vbExclamation, "4 - Détail")
        Exit Sub
    End If
    
    xSQL = Replace(xSQL, "and", "where", 1, 1)
    
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAM0" & xSQL
    If chkSelect_Options_4_SSIDOMDIDX = "1" Then xSQL_H = " union " & Replace(xSQL, "YSSISAM0", "YSSISAMH")
    
    X = xSQL & xSQL_H _
         & " order by SSISAMUIDX , SSISAMUIDD , SSISAMYVER"
    
    Set rsSab = cnsab.Execute(X)
    
    
    Do While Not rsSab.EOF
        Call rsYSSISAM0_GetBuffer(rsSab, xYSSISAM0)
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
    
        fgSelect_Display_H_SAM_Line
    
        rsSab.MoveNext
    Loop

End If
'_______________________________________________________________________________________

fgSelect.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgSelect_Display_H_SAA()

Dim xSQL As String, xSQL_H As String, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_H_SAA"
fgSelect.Visible = False
fgSelect_Reset

fgSelect.FormatString = "<Nature                                  |<Identifiant                                                                   |" _
     & ">Séquence |<Actif |" _
     & "<Libellé                                                                                |" _
     & "<Conforme? |<Profil                                     |Mise à jour par         le                                                 ||||"
fgSelect.Rows = 1
fgSelect.Row = 0

X = Mid$(cboSelect_Options_4_SSIDOMNAT, 1, 1)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSISAANAT = '" & X & "'"
  
X = Trim(txtSelect_Options_4_SSIDOMUIDX)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSISAAUIDX like '%" & X & "%'"

X = Trim(txtSelect_Options_4_SSIDOMUIDD)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSISAAUSEQ = " & Val(X)

If Not blnDisplay Then
    Call MsgBox("Précisez au moins un critère de recherche", vbExclamation, "4 - Détail")
    Exit Sub
End If
Select Case Mid$(cboSelect_Options_4_SSIDOMSTAK, 1, 1)
    Case " ": xSQL = xSQL & " and SSISAASTAK = ' '"
    Case "N": xSQL = xSQL & " and SSISAASTAK = 'N'"
End Select
xSQL = Replace(xSQL, "and", "where", 1, 1)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0" & xSQL
If chkSelect_Options_4_SSIDOMDIDX = "1" Then xSQL_H = " union " & Replace(xSQL, "YSSISAA0", "YSSISAAH")

X = xSQL & xSQL_H _
     & " order by SSISAANAT , SSISAAUIDX , SSISAAUSEQ , SSISAAYVER"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    Call rsYSSISAA0_GetBuffer(rsSab, xYSSISAA0)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1

    fgSelect_Display_H_SAA_Line

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
Private Sub fgSelect_Display_H_WIN()

Dim xSQL As String, xSQL_H As String, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_H_WIN"
fgSelect.Visible = False
fgSelect_Reset

fgSelect.FormatString = "<Nature                                  |<Identifiant                                                                   |" _
     & ">Séquence |<Actif |" _
     & "<Libellé                                                                                |" _
     & "<Conforme? |<Profil                                     |Mise à jour par         le                                                 ||||"
fgSelect.Rows = 1
fgSelect.Row = 0

X = Mid$(cboSelect_Options_4_SSIDOMNAT, 1, 1)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIWINNAT = '" & X & "'"
  
X = Trim(txtSelect_Options_4_SSIDOMUIDX)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIWINUIDX like '%" & X & "%'"

X = Trim(txtSelect_Options_4_SSIDOMUIDD)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIWINUIDD = " & Val(X)

If Not blnDisplay Then
    Call MsgBox("Précisez au moins un critère de recherche", vbExclamation, "4 - Détail")
    Exit Sub
End If
Select Case Mid$(cboSelect_Options_4_SSIDOMSTAK, 1, 1)
    Case " ": xSQL = xSQL & " and SSIWINSTAK = ' '"
    Case "N": xSQL = xSQL & " and SSIWINSTAK = 'N'"
End Select
xSQL = Replace(xSQL, "and", "where", 1, 1)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0" & xSQL
If chkSelect_Options_4_SSIDOMDIDX = "1" Then xSQL_H = " union " & Replace(xSQL, "YSSIWIN0", "YSSIWINH")

X = xSQL & xSQL_H _
     & " order by SSIWINNAT , SSIWINUIDX , SSIWINUIDD , SSIWINYVER"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    Call rsYSSIWIN0_GetBuffer(rsSab, xYSSIWIN0)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1

    fgSelect_Display_H_WIN_Line

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
Private Sub fgSelect_Display_H_DIV()

Dim xSQL As String, xSQL_H As String, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_H_DIV"
fgSelect.Visible = False
fgSelect_Reset

fgSelect.FormatString = "<Nature                                  |<Identifiant                                                                   |" _
     & ">Séquence |<Actif |" _
     & "<Libellé                                                                                |" _
     & "<Conforme? |<Profil                                     |Mise à jour par         le                                                 ||||"
fgSelect.Rows = 1
fgSelect.Row = 0

X = Mid$(cboSelect_Options_4_SSIDOMNAT, 1, 1)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIDIVNAT = '" & X & "'"
  
X = Trim(txtSelect_Options_4_SSIDOMUIDX)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIDIVUIDX like '%" & X & "%'"

X = Trim(txtSelect_Options_4_SSIDOMUIDD)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIDIVUIDD = " & Val(X)

If Not blnDisplay Then
    Call MsgBox("Précisez au moins un critère de recherche", vbExclamation, "4 - Détail")
    Exit Sub
End If
Select Case Mid$(cboSelect_Options_4_SSIDOMSTAK, 1, 1)
    Case " ": xSQL = xSQL & " and SSIDIVSTAK = ' '"
    Case "N": xSQL = xSQL & " and SSIDIVSTAK = 'N'"
End Select
xSQL = Replace(xSQL, "and", "where", 1, 1)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0" & xSQL
If chkSelect_Options_4_SSIDOMDIDX = "1" Then xSQL_H = " union " & Replace(xSQL, "YSSIDIV0", "YSSIDIVH")

X = xSQL & xSQL_H _
     & " order by SSIDIVNAT , SSIDIVUIDX , SSIDIVUIDD , SSIDIVYVER"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    Call rsYSSIDIV0_GetBuffer(rsSab, xYSSIDIV0)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1

    fgSelect_Display_H_DIV_Line

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
Private Sub fgSelect_Display_H_MEL()

Dim xSQL As String, xSQL_H As String, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_H_MEL"
fgSelect.Visible = False
fgSelect_Reset

fgSelect.FormatString = "<Nature                                  |<Identifiant                                                                   |" _
     & ">Séquence |<Actif |" _
     & "<Libellé                                                                                |" _
     & "<Conforme? |<Profil                                     |Mise à jour par         le                                                 ||||"
fgSelect.Rows = 1
fgSelect.Row = 0

X = Mid$(cboSelect_Options_4_SSIDOMNAT, 1, 1)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIMELNAT = '" & X & "'"
  
X = Trim(txtSelect_Options_4_SSIDOMUIDX)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIMELUIDX like '%" & X & "%'"

X = Trim(txtSelect_Options_4_SSIDOMUIDD)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIMELUIDD = " & Val(X)

If Not blnDisplay Then
    Call MsgBox("Précisez au moins un critère de recherche", vbExclamation, "4 - Détail")
    Exit Sub
End If
Select Case Mid$(cboSelect_Options_4_SSIDOMSTAK, 1, 1)
    Case " ": xSQL = xSQL & " and SSIMELSTAK = ' '"
    Case "N": xSQL = xSQL & " and SSIMELSTAK = 'N'"
End Select
xSQL = Replace(xSQL, "and", "where", 1, 1)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0" & xSQL
If chkSelect_Options_4_SSIDOMDIDX = "1" Then xSQL_H = " union " & Replace(xSQL, "YSSIMEL0", "YSSIMELH")

X = xSQL & xSQL_H _
     & " order by SSIMELNAT , SSIMELUIDX , SSIMELUIDD , SSIMELYVER"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    Call rsYSSIMEL0_GetBuffer(rsSab, xYSSIMEL0)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1

    fgSelect_Display_H_MEL_Line

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

Private Sub fgSelect_Display_H_TIC()

Dim xSQL As String, xSQL_H As String, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_H_TIC"
fgSelect.Visible = False
fgSelect_Reset

fgSelect.FormatString = "<Nature                                  |<Identifiant                                                                   |" _
     & ">Séquence |<Actif |" _
     & "<Libellé                                                                                |" _
     & "<Conforme? |<Profil                                     |Mise à jour par         le                                                 ||||"
fgSelect.Rows = 1
fgSelect.Row = 0

X = Mid$(cboSelect_Options_4_SSIDOMNAT, 1, 1)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSITICNAT = '" & X & "'"
  
X = Trim(txtSelect_Options_4_SSIDOMUIDX)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSITICUIDX like '%" & X & "%'"

X = Trim(txtSelect_Options_4_SSIDOMUIDD)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSITICUIDD = " & Val(X)

If Not blnDisplay Then
    Call MsgBox("Précisez au moins un critère de recherche", vbExclamation, "4 - Détail")
    Exit Sub
End If
Select Case Mid$(cboSelect_Options_4_SSIDOMSTAK, 1, 1)
    Case " ": xSQL = xSQL & " and SSITICSTAK = ' '"
    Case "N": xSQL = xSQL & " and SSITICSTAK = 'N'"
End Select
xSQL = Replace(xSQL, "and", "where", 1, 1)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSITIC0" & xSQL
If chkSelect_Options_4_SSIDOMDIDX = "1" Then xSQL_H = " union " & Replace(xSQL, "YSSITIC0", "YSSITICH")

If Mid$(cboSelect_Options_4_SSIDOMNAT, 1, 1) = "D" Then
    X = xSQL & xSQL_H _
         & " order by SSITICNAT  , SSITICUIDD , SSITICYVER"
Else
    X = xSQL & xSQL_H _
         & " order by SSITICNAT , SSITICUIDX , SSITICUIDD , SSITICYVER"
End If

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    Call rsYSSITIC0_GetBuffer(rsSab, xYSSITIC0)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1

    fgSelect_Display_H_TIC_Line

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







Private Sub fgSelect_Display_H_USR()

Dim xSQL As String, xSQL_H As String, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_H_USR"
fgSelect.Visible = False
fgSelect_Reset

fgSelect.FormatString = "<Nature                                  |<Identifiant                                                                   |" _
     & ">Lot |<Actif |" _
     & "<Libellé                                                                                |" _
     & "<Conforme? |<Profil                                     |Mise à jour par         le                                                 ||||"
fgSelect.Rows = 1
fgSelect.Row = 0

X = Mid$(cboSelect_Options_4_SSIDOMNAT, 1, 1)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIUSRNAT = '" & X & "'"
  
X = Trim(txtSelect_Options_4_SSIDOMUIDX)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIUSRUIDX like '%" & X & "%'"

X = Trim(txtSelect_Options_4_SSIDOMUIDD)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIUSRUSEQ = " & Val(X)

If Not blnDisplay Then
    Call MsgBox("Précisez au moins un critère de recherche", vbExclamation, "4 - Détail")
    Exit Sub
End If
Select Case Mid$(cboSelect_Options_4_SSIDOMSTAK, 1, 1)
    Case " ": xSQL = xSQL & " and SSIUSRSTAK = ' '"
    Case "N": xSQL = xSQL & " and SSIUSRSTAK = 'N'"
End Select
xSQL = Replace(xSQL, "and", "where", 1, 1)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0" & xSQL
If chkSelect_Options_4_SSIDOMDIDX = "1" Then xSQL_H = " union " & Replace(xSQL, "YSSIUSR0", "YSSIUSRH")

X = xSQL & xSQL_H _
     & " order by SSIUSRNAT , SSIUSRUIDX  , SSIUSRYVER"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    Call rsYSSIUSR0_GetBuffer(rsSab, xYSSIUSR0)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1

    fgSelect_Display_H_USR_Line

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

Private Sub fgSelect_Display_H_DOM()

Dim xSQL As String, xSQL_H As String, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_H_DOM"
fgSelect.Visible = False
fgSelect_Reset

fgSelect.FormatString = "<Nature                                  |<Identifiant                                                                   |" _
     & ">Lot |<Actif |" _
     & "<Libellé                                                                                |" _
     & "<Conforme? |<Profil                                     |Mise à jour par         le                                                 ||||"
fgSelect.Rows = 1
fgSelect.Row = 0

X = Mid$(cboSelect_Options_4_SSIDOMNAT, 1, 1)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIDOMNAT = '" & X & "'"
  
X = Trim(txtSelect_Options_4_SSIDOMUIDX)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIDOMUIDX like '%" & X & "%'"

X = Trim(txtSelect_Options_4_SSIDOMUIDD)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIDOMUSEQ = " & Val(X)

If Not blnDisplay Then
    Call MsgBox("Précisez au moins un critère de recherche", vbExclamation, "4 - Détail")
    Exit Sub
End If
Select Case Mid$(cboSelect_Options_4_SSIDOMSTAK, 1, 1)
    Case " ": xSQL = xSQL & " and SSIDOMSTAK = ' '"
    Case "N": xSQL = xSQL & " and SSIDOMSTAK = 'N'"
End Select
xSQL = Replace(xSQL, "and", "where", 1, 1)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0" & xSQL
If chkSelect_Options_4_SSIDOMDIDX = "1" Then xSQL_H = " union " & Replace(xSQL, "YSSIDOM0", "YSSIDOMH")

X = xSQL & xSQL_H _
     & " order by SSIDOMNAT , SSIDOMUIDN , SSIDOMDIDX  , SSIDOMUIDX , SSIDOMYVER"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    Call rsYSSIDOM0_GetBuffer(rsSab, xYSSIDOM0)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1

    fgSelect_Display_H_DOM_Line

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

Public Sub fgSelect_Display_H_DOM_Line()
On Error Resume Next
Dim K As Integer
fgSelect.Col = 0
fgSelect.Text = arrJRN_Origine(2)
'fgSelect.Col = 2: fgSelect.Text = xYSSIDOM0.SSIDOMDIDX
fgSelect.Col = 1: fgSelect.Text = xYSSIDOM0.SSIDOMUIDN & " - " & Trim(xYSSIDOM0.SSIDOMDIDX) & " - " & xYSSIDOM0.SSIDOMUIDD
fgSelect.Col = 3: fgSelect.Text = " " & xYSSIDOM0.SSIDOMSTAK
fgSelect.Col = 4: fgSelect.Text = Trim(xYSSIDOM0.SSIDOMUIDX)
fgSelect.Col = 5: fgSelect.Text = Trim(xYSSIDOM0.SSIDOMPRFK)
fgSelect.Col = 6: fgSelect.Text = Trim(xYSSIDOM0.SSIDOMPRFX)
fgSelect.Col = 7
fgSelect.Text = xYSSIDOM0.SSIDOMYFCT & " - " & Trim(xYSSIDOM0.SSIDOMYUSR) _
    & " - " & dateImp10_S(xYSSIDOM0.SSIDOMYAMJ) & " - " & timeImp8(xYSSIDOM0.SSIDOMYHMS) & " - V" & xYSSIDOM0.SSIDOMYVER

fgSelect.Col = 9: fgSelect.Text = "DOM|" & xYSSIDOM0.SSIDOMNAT & "|" & xYSSIDOM0.SSIDOMUIDN & "|" & Trim(xYSSIDOM0.SSIDOMDIDX) & "|" _
                                & Trim(xYSSIDOM0.SSIDOMUIDX) & "|" & xYSSIDOM0.SSIDOMUIDD & "|" & xYSSIDOM0.SSIDOMYVER & "|"

End Sub



Public Sub fgSelect_Display_H_USR_Line()
On Error Resume Next
Dim K As Integer
fgSelect.Col = 0
fgSelect.Text = arrJRN_Origine(1)
'fgSelect.Col = 2: fgSelect.Text = xYSSIUSR0.SSIUSRUIDN
fgSelect.Col = 1: fgSelect.Text = xYSSIUSR0.SSIUSRUIDN
fgSelect.Col = 3: fgSelect.Text = " " & xYSSIUSR0.SSIUSRSTAK
fgSelect.Col = 4: fgSelect.Text = Trim(xYSSIUSR0.SSIUSRUIDX)
fgSelect.Col = 5: fgSelect.Text = Trim(xYSSIUSR0.SSIUSRPRFK)
fgSelect.Col = 6: fgSelect.Text = Trim(xYSSIUSR0.SSIUSRPRFX)
fgSelect.Col = 7
fgSelect.Text = xYSSIUSR0.SSIUSRYFCT & " - " & Trim(xYSSIUSR0.SSIUSRYUSR) _
    & " - " & dateImp10_S(xYSSIUSR0.SSIUSRYAMJ) & " - " & timeImp8(xYSSIUSR0.SSIUSRYHMS) & " - V" & xYSSIUSR0.SSIUSRYVER

fgSelect.Col = 9: fgSelect.Text = "USR|" & xYSSIUSR0.SSIUSRNAT & "|" _
                                & Trim(xYSSIUSR0.SSIUSRUIDN) & "|" & xYSSIUSR0.SSIUSRYVER & "|"

End Sub

Private Sub fgSelect_Display_H_IBM()

Dim xSQL As String, xSQL_H As String, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_H_IBM"
fgSelect.Visible = False
fgSelect_Reset

fgSelect.FormatString = "<Nature                                  |<Identifiant                                                                   |" _
     & ">Lot |<Actif |" _
     & "<Libellé                                                                                |" _
     & "<Conforme? |<Profil                                     |Mise à jour par         le                                                 ||||"
fgSelect.Rows = 1
fgSelect.Row = 0

X = Mid$(cboSelect_Options_4_SSIDOMNAT, 1, 1)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIIBMNAT = '" & X & "'"
  
X = Trim(txtSelect_Options_4_SSIDOMUIDX)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and UPUPRF like '%" & X & "%'"

X = Trim(txtSelect_Options_4_SSIDOMUIDD)
If X <> "" Then blnDisplay = True: xSQL = xSQL & " and SSIIBMUIDD = " & Val(X)

If Not blnDisplay Then
    Call MsgBox("Précisez au moins un critère de recherche", vbExclamation, "4 - Détail")
    Exit Sub
End If
Select Case Mid$(cboSelect_Options_4_SSIDOMSTAK, 1, 1)
    Case " ": xSQL = xSQL & " and UPSTAT = '*ENABLED'"
    Case "N": xSQL = xSQL & " and UPSTAT = '*DISABLED'"
End Select
xSQL = Replace(xSQL, "and", "where", 1, 1)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0" & xSQL
If chkSelect_Options_4_SSIDOMDIDX = "1" Then xSQL_H = " union " & Replace(xSQL, "YSSIIBM0", "YSSIIBMH")

X = xSQL & xSQL_H _
     & " order by SSIIBMNAT , UPUPRF , SSIIBMYVER"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    Call rsYSSIIBM0_GetBuffer(rsSab, xYSSIIBM0)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1

    fgSelect_Display_H_IBM_Line

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

Private Sub fgSelect_Display_3()

Dim K As Long, xSQL As String, blnRupture As Boolean, blnDisplay As Boolean
Dim wLnk As String, Nb As Integer
On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_3"

oldYSSIUSR0.SSIUSRUIDN = 0
Do While Not rsSab.EOF
    Call rsYSSIUSR0_GetBuffer(rsSab, xYSSIUSR0)
        
    Call rsYSSIDOM0_GetBuffer(rsSab, xYSSIDOM0)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    Nb = Nb + 1
    If oldYSSIUSR0.SSIUSRUIDN <> xYSSIUSR0.SSIUSRUIDN Then
        oldYSSIUSR0.SSIUSRUIDN = xYSSIUSR0.SSIUSRUIDN
        oldYSSIDOM0.SSIDOMDIDX = xYSSIDOM0.SSIDOMDIDX
        fgSelect.Col = 0: fgSelect.Text = Trim(xYSSIUSR0.SSIUSRUIDX): fgSelect.CellBackColor = mColor_Y2
        fgSelect.Col = 1: fgSelect.Text = Trim(xYSSIDOM0.SSIDOMDIDX): fgSelect.CellBackColor = mColor_Y2
    Else
        If oldYSSIDOM0.SSIDOMDIDX <> xYSSIDOM0.SSIDOMDIDX Then
            oldYSSIDOM0.SSIDOMDIDX = xYSSIDOM0.SSIDOMDIDX
          '  fgSelect.Col = 1: fgSelect.Text = Trim(xYSSIDOM0.SSIDOMDIDX): fgSelect.CellBackColor = mColor_Y2
        End If
            fgSelect.Col = 1: fgSelect.Text = Trim(xYSSIDOM0.SSIDOMDIDX): fgSelect.CellBackColor = mColor_Y2
    End If
    fgSelect.Col = 9
    wLnk = xYSSIDOM0.SSIDOMNAT & "|" & xYSSIDOM0.SSIDOMUIDN & "|" & xYSSIDOM0.SSIDOMDIDX & "|" _
                  & xYSSIDOM0.SSIDOMUIDX & "|" & xYSSIDOM0.SSIDOMUIDD & "|"
    
    fgSelect.Text = wLnk
    fgSelect.Col = 2: fgSelect.CellBackColor = &HC0E0FF: fgSelect.Text = Trim(xYSSIDOM0.SSIDOMUIDX)

    fgSelect.CellFontBold = True
    fgSelect.Col = 3: fgSelect.Text = Trim(xYSSIDOM0.SSIDOMPRFX): fgSelect.CellBackColor = &HC0E0FF
    fgSelect.CellForeColor = vbBlue
     fgSelect.Col = 4
   Select Case xYSSIDOM0.SSIDOMPRFK
        Case " ":  fgSelect.CellBackColor = &HC0E0FF
        Case "N": fgSelect.Text = "Non conforme": fgSelect.CellBackColor = &HC0E0FF
        Case "X": fgSelect.Text = "EXIT_GRP": fgSelect.CellBackColor = RGB(192, 192, 192)
        Case "!": fgSelect.Text = "état Domaine # SSI": fgSelect.CellBackColor = mColor_W0
        Case "?": fgSelect.Text = "En attente": fgSelect.CellBackColor = mColor_W0
        Case Else: fgSelect.Text = xYSSIDOM0.SSIDOMPRFK & "cas non traité": fgSelect.CellBackColor = mColor_W1
    End Select
    Select Case Trim(xYSSIDOM0.SSIDOMDIDX)
        Case "IBM": Call fgSelect_Display_3_Line_YSSIIBM0(wLnk)
        Case "SAB":
            'fgSelect.Col = 5: fgSelect.Text = Trim(rsSab("SSISABPRFX")): fgSelect.CellBackColor = mColor_W1
        '???Case "SAA": Call fgSelect_Display_3_Line_YSSISAA0(wLnk)
    End Select
    
    rsSab.MoveNext
Loop
arrCtl_Nb(arrCtl_K) = Nb

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgSelect_Display_3_SSIUSRPRFK()

Dim K As Long, xSQL As String, blnRupture As Boolean, blnDisplay As Boolean
Dim wLnk As String, Nb As Integer
On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_3_SSIUSRPRFK"

oldYSSIUSR0.SSIUSRUIDN = 0
Do While Not rsSab.EOF
    Call rsYSSIUSR0_GetBuffer(rsSab, xYSSIUSR0)
        
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    Nb = Nb + 1
        oldYSSIUSR0.SSIUSRUIDN = xYSSIUSR0.SSIUSRUIDN
        fgSelect.Col = 0: fgSelect.Text = Trim(xYSSIUSR0.SSIUSRUIDX): fgSelect.CellBackColor = mColor_Y2
    fgSelect.Col = 9
    wLnk = xYSSIUSR0.SSIUSRNAT & "|" & xYSSIUSR0.SSIUSRUIDN & "|" & "|" & "|"
    
    fgSelect.Text = wLnk
    'fgSelect.Col = 2: fgSelect.Text = Trim(xYSSIUSR0.SSIUSRUIDX): fgSelect.CellBackColor = &HC0E0FF
    'fgSelect.CellFontBold = True
    fgSelect.Col = 3: fgSelect.Text = Trim(xYSSIUSR0.SSIUSRPRFX): fgSelect.CellBackColor = &HC0E0FF
    fgSelect.CellForeColor = vbBlue
     fgSelect.Col = 4
   Select Case xYSSIUSR0.SSIUSRPRFK
        Case "", " ": fgSelect.CellBackColor = &HC0E0FF
        Case "N": fgSelect.Text = "Non conforme": fgSelect.CellBackColor = &HC0E0FF
        Case "X": fgSelect.Text = "EXIT_GRP": fgSelect.CellBackColor = RGB(192, 192, 192)
        Case "?": fgSelect.Text = "En attente": fgSelect.CellBackColor = mColor_W0
       Case "!": fgSelect.Text = "état IBM # SSI": fgSelect.CellBackColor = mColor_W0
        Case Else: fgSelect.Text = xYSSIUSR0.SSIUSRPRFX & "cas non traité": fgSelect.CellBackColor = mColor_W1
    End Select
    
    rsSab.MoveNext
Loop
arrCtl_Nb(arrCtl_K) = Nb


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgSelect_Display_3_Inactifs(wComment As String)

Dim K As Long, xSQL As String, blnRupture As Boolean, blnDisplay As Boolean
Dim wLnk As String, Nb As Integer
On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_3"

oldYSSIUSR0.SSIUSRUIDN = 0


Do While Not rsSab.EOF
    Call rsYSSIUSR0_GetBuffer(rsSab, xYSSIUSR0)
        
    Call rsYSSIDOM0_GetBuffer(rsSab, xYSSIDOM0)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    Nb = Nb + 1
    If oldYSSIUSR0.SSIUSRUIDN <> xYSSIUSR0.SSIUSRUIDN Then
        oldYSSIUSR0.SSIUSRUIDN = xYSSIUSR0.SSIUSRUIDN
        oldYSSIDOM0.SSIDOMDIDX = xYSSIDOM0.SSIDOMDIDX
        fgSelect.Col = 0: fgSelect.Text = Trim(xYSSIUSR0.SSIUSRUIDX): fgSelect.CellBackColor = mColor_W0
        fgSelect.Col = 1: fgSelect.Text = Trim(xYSSIDOM0.SSIDOMDIDX): fgSelect.CellBackColor = mColor_W0
    Else
        If oldYSSIDOM0.SSIDOMDIDX <> xYSSIDOM0.SSIDOMDIDX Then
            oldYSSIDOM0.SSIDOMDIDX = xYSSIDOM0.SSIDOMDIDX
            'fgSelect.Col = 1: fgSelect.Text = Trim(xYSSIDOM0.SSIDOMDIDX): fgSelect.CellBackColor = mColor_W0
        End If
            fgSelect.Col = 1: fgSelect.Text = Trim(xYSSIDOM0.SSIDOMDIDX): fgSelect.CellBackColor = mColor_W0
    End If
    fgSelect.Col = 9
    wLnk = xYSSIDOM0.SSIDOMNAT & "|" & xYSSIDOM0.SSIDOMUIDN & "|" & xYSSIDOM0.SSIDOMDIDX & "|" _
                  & xYSSIDOM0.SSIDOMUIDX & "|" & xYSSIDOM0.SSIDOMUIDD & "|"
    
    fgSelect.Text = wLnk
    fgSelect.Col = 2: fgSelect.CellBackColor = mColor_W0
    fgSelect.Text = Trim(xYSSIDOM0.SSIDOMUIDX)
    fgSelect.CellFontBold = True
    fgSelect.Col = 3: fgSelect.Text = Trim(xYSSIDOM0.SSIDOMPRFX): fgSelect.CellBackColor = mColor_W0
    fgSelect.CellForeColor = vbBlue
    fgSelect.Col = 4: fgSelect.Text = wComment: fgSelect.CellBackColor = mColor_W1

     fgSelect.Col = 5
   Select Case xYSSIDOM0.SSIDOMPRFK
        Case " ": fgSelect.Text = xYSSIDOM0.SSIDOMDIDX & " : compte actif":  fgSelect.CellBackColor = &HC0E0FF
        Case "N": fgSelect.Text = xYSSIDOM0.SSIDOMDIDX & " : compte actif : Non conforme": fgSelect.CellBackColor = &HC0E0FF
        Case "X": fgSelect.Text = xYSSIDOM0.SSIDOMDIDX & " : compte actif : EXIT_GRP": fgSelect.CellBackColor = RGB(192, 192, 192)
         Case "?": fgSelect.Text = xYSSIDOM0.SSIDOMDIDX & " : En attente": fgSelect.CellBackColor = mColor_W0
       Case "!": fgSelect.Text = xYSSIDOM0.SSIDOMDIDX & " : compte actif : état compte # SSI": fgSelect.CellBackColor = mColor_W0
        Case Else: fgSelect.Text = xYSSIDOM0.SSIDOMDIDX & " : " & xYSSIDOM0.SSIDOMPRFK & "cas non traité": fgSelect.CellBackColor = mColor_W1
    End Select
    
    rsSab.MoveNext
Loop

arrCtl_Nb(arrCtl_K) = Nb

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgSelect_Display_3_Orphelins(lSSIDOMDIDX As String)

Dim Nb As Long, wLnk As String, K As Long
On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_3_Orphelins"
Do While Not rsSab.EOF
    Nb = Nb + 1
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect.Col = 1: fgSelect.Text = lSSIDOMDIDX: fgSelect.CellBackColor = mColor_W0
    fgSelect.Col = 4:
    Select Case lSSIDOMDIDX
        Case "IBM_S"
            fgSelect.Text = "SUPPRIME": fgSelect.CellBackColor = vbRed: fgSelect.CellForeColor = vbYellow
                fgSelect.Col = 1: fgSelect.CellBackColor = vbRed: fgSelect.CellForeColor = vbYellow
        Case "IBM_H"
            fgSelect.Text = "Orphelin (historique)": fgSelect.CellForeColor = vbRed: fgSelect.CellBackColor = vbYellow
                fgSelect.Col = 1: fgSelect.CellForeColor = vbRed: fgSelect.CellBackColor = vbYellow
         Case "WIN_S"
            fgSelect.Text = "SUPPRIME": fgSelect.CellBackColor = vbRed: fgSelect.CellForeColor = vbYellow
                fgSelect.Col = 1: fgSelect.CellBackColor = vbRed: fgSelect.CellForeColor = vbYellow
      Case Else
            fgSelect.Text = "Orphelin": fgSelect.CellBackColor = mColor_W1
        End Select
        
    fgSelect.Col = 0: fgSelect.Text = "": fgSelect.CellBackColor = mColor_W0
    Select Case lSSIDOMDIDX
        Case "WIN", "WIN_S": Call rsYSSIWIN0_GetBuffer(rsSab, xYSSIWIN0)
            fgSelect.Col = 3: fgSelect.Text = xYSSIWIN0.SSIWINPRFX: fgSelect.CellBackColor = mColor_W0
            fgSelect.Col = 2: fgSelect.Text = xYSSIWIN0.SSIWINUIDX: fgSelect.CellBackColor = mColor_W0
            fgSelect.Col = 9
            wLnk = "?| |" & lSSIDOMDIDX & "|" & xYSSIWIN0.SSIWINGUID & "|" & xYSSIWIN0.SSIWINUIDD & "|"
            
            fgSelect.Text = wLnk
            fgSelect.Col = 6: fgSelect.Text = "Id " & xYSSIWIN0.SSIWINUIDD: fgSelect.CellBackColor = mColor_Y2
        Case "IBM", "IBM_H", "IBM_S": Call rsYSSIIBM0_GetBuffer(rsSab, xYSSIIBM0)
            fgSelect.Col = 3: fgSelect.Text = xYSSIIBM0.UPINPG: fgSelect.CellBackColor = mColor_W0
            fgSelect.Col = 2: fgSelect.Text = xYSSIIBM0.UPUPRF: fgSelect.CellBackColor = mColor_W0
            fgSelect.Col = 9
            wLnk = "?| |" & lSSIDOMDIDX & "|" & xYSSIIBM0.UPUPRF & "|" & xYSSIIBM0.SSIIBMUIDD & "|"
            
            fgSelect.Text = wLnk
            fgSelect.Col = 6: fgSelect.Text = "Créé le " & dateImp10_S(xYSSIIBM0.UPCRTD): fgSelect.CellBackColor = mColor_Y2
            fgSelect.Col = 5: fgSelect.Text = "Dernière connexion le " & dateImp10_S(xYSSIIBM0.UPPSOD): fgSelect.CellBackColor = &HC0E0FF
            fgSelect.CellFontBold = True
        Case "SAA": Call rsYSSISAA0_GetBuffer(rsSab, xYSSISAA0)
            fgSelect.Col = 3: fgSelect.Text = xYSSISAA0.SSISAAPRFX: fgSelect.CellBackColor = mColor_W0
            fgSelect.Col = 2: fgSelect.Text = xYSSISAA0.SSISAAUIDX: fgSelect.CellBackColor = mColor_W0
            fgSelect.Col = 9
            wLnk = "?|SAA|" & xYSSISAA0.SSISAAUIDX & "| | |"
            wLnk = "?| |SAA|" & xYSSISAA0.SSISAAUIDX & "|" & xYSSISAA0.SSISAAUIDD & "|"
            fgSelect.Text = wLnk
            fgSelect.Col = 6: fgSelect.Text = Trim(xYSSISAA0.SSISAAUNOM): fgSelect.CellBackColor = mColor_Y2
            fgSelect.Col = 5:
                If xYSSISAA0.SSISAASTAK = " " Then
                    fgSelect.Text = "Actif": fgSelect.CellBackColor = mColor_Y2
                Else
                    fgSelect.Text = "Inactif": fgSelect.CellBackColor = RGB(190, 190, 190)
                End If
                
            fgSelect.CellFontBold = True
        Case "SAB", "SAB_H": Call rsYSSISAB0_GetBuffer(rsSab, xYSSISAB0)
            fgSelect.Col = 3: fgSelect.Text = xYSSISAB0.SSISABPRFX: fgSelect.CellBackColor = mColor_W0
            fgSelect.Col = 2: fgSelect.Text = xYSSISAB0.SSISABUIDX: fgSelect.CellBackColor = mColor_W0
            fgSelect.Col = 9
            wLnk = "?| |" & lSSIDOMDIDX & "|" & xYSSISAB0.SSISABUIDX & "|" & xYSSISAB0.SSISABUIDD & "|"
            fgSelect.Text = wLnk
            fgSelect.Col = 6: fgSelect.Text = Trim(xYSSISAB0.SSISABUNOM): fgSelect.CellBackColor = mColor_Y2
            fgSelect.Col = 5:
                If xYSSISAB0.SSISABSTAK = " " Then
                    fgSelect.Text = "Actif": fgSelect.CellBackColor = mColor_Y2
                Else
                    fgSelect.Text = "Inactif": fgSelect.CellBackColor = RGB(190, 190, 190)
                End If
                
            fgSelect.CellFontBold = True
         Case "DIV": Call rsYSSIDIV0_GetBuffer(rsSab, xYSSIDIV0)
            fgSelect.Col = 3: fgSelect.Text = xYSSIDIV0.SSIDIVPRFX: fgSelect.CellBackColor = mColor_W0
            fgSelect.Col = 2: fgSelect.Text = xYSSIDIV0.SSIDIVUIDX: fgSelect.CellBackColor = mColor_W0
            fgSelect.Col = 9
            wLnk = "?| |" & lSSIDOMDIDX & "|" & xYSSIDIV0.SSIDIVUIDX & "|" & xYSSIDIV0.SSIDIVUIDD & "|"
            
            fgSelect.Text = wLnk
            fgSelect.Col = 6: fgSelect.Text = "Id " & xYSSIDIV0.SSIDIVUIDD: fgSelect.CellBackColor = mColor_Y2
        Case "TIC", "TIC_S": Call rsYSSITIC0_GetBuffer(rsSab, xYSSITIC0)
            fgSelect.Col = 3: fgSelect.Text = xYSSITIC0.SSITICPRFX: fgSelect.CellBackColor = mColor_W0
            fgSelect.Col = 2: fgSelect.Text = xYSSITIC0.SSITICUIDX: fgSelect.CellBackColor = mColor_W0
            fgSelect.Col = 9
            wLnk = "?| |" & lSSIDOMDIDX & "|" & xYSSITIC0.SSITICUIDX & "|" & xYSSITIC0.SSITICUIDD & "|"
            
            fgSelect.Text = wLnk
            fgSelect.Col = 6: fgSelect.Text = "Id " & xYSSITIC0.SSITICUIDD: fgSelect.CellBackColor = mColor_Y2
   End Select
    
    rsSab.MoveNext
Loop
arrCtl_Nb(arrCtl_K) = Nb

Call lstErr_AddItem(lstErr, cmdContext, lSSIDOMDIDX & " : " & Nb & " enregistrements"): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_Display_1_Line()
On Error Resume Next
Dim K As Integer
fgSelect.Col = 0: fgSelect.Text = xYSSIUSR0.SSIUSRNAT
fgSelect.Col = 1: fgSelect.Text = xYSSIUSR0.SSIUSRUIDN
fgSelect.Col = 2: fgSelect.Text = Trim(xYSSIUSR0.SSIUSRUIDX) ': fgSelect.Font.Bold = True
fgSelect.Col = 3: fgSelect.Text = " " & xYSSIUSR0.SSIUSRSTAK
Select Case xYSSIUSR0.SSIUSRSTAK
    Case " "
    Case "N": fgSelect.CellBackColor = RGB(192, 192, 192)
    Case Else: fgSelect.CellBackColor = mColor_W1
End Select

fgSelect.Col = 4
If xYSSIUSR0.SSIUSRDECH <> 0 Then
    fgSelect.Text = "  " & dateImp10_S(xYSSIUSR0.SSIUSRDECH)
    If xYSSIUSR0.SSIUSRSTAK = " " Then
        If xYSSIUSR0.SSIUSRDECH < DSys Then
            fgSelect.CellBackColor = mColor_W1
        Else
            fgSelect.CellBackColor = mColor_Y1
        End If
    Else
        fgSelect.CellBackColor = RGB(220, 220, 220)
    End If
        
End If
fgSelect.Col = 5: fgSelect.Text = xYSSIUSR0.SSIUSRPRFX: fgSelect.CellForeColor = RGB(0, 96, 0)

If Len(xYSSIUSR0.SSIUSRUNIT) = 3 Then
    K = Val(Mid$(xYSSIUSR0.SSIUSRUNIT, 2, 2))
    X = " " & xYSSIUSR0.SSIUSRUNIT & " : " & arrSSIUSRUNIT_Code(K)
Else
    X = ""
End If

If xYSSIUSR0.SSIUSRPRFD = 0 Then
    fgSelect.Col = 6: fgSelect.Text = xYSSIUSR0.SSIUSRPRFK & X
Else
    fgSelect.Col = 6: fgSelect.Text = xYSSIUSR0.SSIUSRPRFK & "  " & dateImp10_S(xYSSIUSR0.SSIUSRPRFD) & X
End If

If xYSSIUSR0.SSIUSRPRFK <> " " Then fgSelect.CellBackColor = mColor_W1
If xYSSIUSR0.SSIUSRTLNK > 0 Then fgSelect.Col = 7: fgSelect.Text = xYSSIUSR0.SSIUSRTLNK
fgSelect.Col = 8: fgSelect.Text = xYSSIUSR0.SSIUSRYFCT & " par " & xYSSIUSR0.SSIUSRYUSR _
  & " le " & dateImp10_S(xYSSIUSR0.SSIUSRYAMJ) & "  " & timeImp8(xYSSIUSR0.SSIUSRYHMS) & " (" & xYSSIUSR0.SSIUSRYVER & ")"
fgSelect.CellFontSize = 8
fgSelect.CellForeColor = RGB(128, 128, 128)

If xYSSIUSR0.SSIUSRSTAK = "N" Then
    For K = 0 To 9: fgSelect.Col = K: fgSelect.CellForeColor = RGB(96, 96, 96): Next K
End If


End Sub

Public Sub fgSelect_Display_H_SAB_Line()
On Error Resume Next
Dim K As Integer
fgSelect.Col = 0
Select Case xYSSISAB0.SSISABNAT
    Case " ": fgSelect.Text = " _SAB Utilisateurs"
    Case "$": fgSelect.Text = "$_SAB Profils"
    Case "2": fgSelect.Text = "2_SAB GRP Menus"
    Case "3": fgSelect.Text = "3_SAB GRP Données"
    Case "4": fgSelect.Text = "4_SAB GRP Métiers"
    Case "C": fgSelect.Text = "C_SAB Hab Classes"
    Case "D": fgSelect.Text = "D_SAB Hab Services"
    Case "M": fgSelect.Text = "M_SAB Hab Options"
    Case "H": fgSelect.Text = "H_SAB Hab Lots"
    Case "W": fgSelect.Text = "W_SAB Hab Swift"
    Case Else: fgSelect.Text = xYSSISAB0.SSISABNAT
End Select
fgSelect.Col = 2: fgSelect.Text = xYSSISAB0.SSISABULOT
fgSelect.Col = 1: fgSelect.Text = Trim(xYSSISAB0.SSISABUIDX)
fgSelect.Col = 3: fgSelect.Text = " " & xYSSISAB0.SSISABSTAK
fgSelect.Col = 4: fgSelect.Text = Trim(xYSSISAB0.SSISABUNOM)
fgSelect.Col = 5: fgSelect.Text = Trim(xYSSISAB0.SSISABPRFK)
fgSelect.Col = 6: fgSelect.Text = Trim(xYSSISAB0.SSISABPRFX)
fgSelect.Col = 7
fgSelect.Text = xYSSISAB0.SSISABYFCT & " - " & Trim(xYSSISAB0.SSISABYUSR) _
    & " - " & dateImp10_S(xYSSISAB0.SSISABYAMJ) & " - " & timeImp8(xYSSISAB0.SSISABYHMS) & " - V" & xYSSISAB0.SSISABYVER


fgSelect.Col = 9: fgSelect.Text = "SAB|" & xYSSISAB0.SSISABNAT & "|" _
                                & Trim(xYSSISAB0.SSISABUIDX) & "|" & xYSSISAB0.SSISABULOT & "|" & xYSSISAB0.SSISABYVER & "|"

End Sub


Public Sub fgSelect_Display_H_SAM_Line()
On Error Resume Next
Dim K As Integer
'fgSelect.Col = 0
fgSelect.Col = 2: fgSelect.Text = xYSSISAM0.SSISAMREF
fgSelect.Col = 1: fgSelect.Text = Trim(xYSSISAM0.SSISAMUIDX)
fgSelect.Col = 0: fgSelect.Text = xYSSISAM0.SSISAMUIDD
fgSelect.Col = 7
fgSelect.Text = xYSSISAM0.SSISAMYFCT & " - " & Trim(xYSSISAM0.SSISAMYUSR) _
    & " - " & dateImp10_S(xYSSISAM0.SSISAMYAMJ) & " - " & timeImp8(xYSSISAM0.SSISAMYHMS) & " - V" & xYSSISAM0.SSISAMYVER


fgSelect.Col = 9: fgSelect.Text = "SAB_M|" & xYSSISAM0.SSISAMUIDD & "|" & xYSSISAM0.SSISAMYVER & "|"

End Sub



Public Sub fgSelect_Display_H_SAA_Line()
On Error Resume Next

fgSelect.Col = 0
fgSelect.Text = xYSSISAA0.SSISAANAT
fgSelect.Col = 2: fgSelect.Text = xYSSISAA0.SSISAAUSEQ
If xYSSISAA0.SSISAAUIDD = 0 Then
    fgSelect.Col = 1: fgSelect.Text = Trim(xYSSISAA0.SSISAAUIDX)
Else
    fgSelect.Col = 1: fgSelect.Text = Trim(xYSSISAA0.SSISAAUIDX) & " = " & xYSSISAA0.SSISAAUIDD
End If
fgSelect.Col = 3: fgSelect.Text = " " & xYSSISAA0.SSISAASTAK
fgSelect.Col = 4: fgSelect.Text = Trim(xYSSISAA0.SSISAAUNOM)
If xYSSISAA0.SSISAANAT = "P" Then
    Dim K1 As Integer, K2 As Integer
    K2 = xYSSISAA0.SSISAAUSEQ
    K1 = Fix(K2 / 1000)
    K2 = K2 - K1 * 1000
    fgSelect.Text = arrSAA_App_Code(K1) & " - " & arrSAA_Function_Code(K2)
Else
    X = ""

End If

fgSelect.Col = 5: fgSelect.Text = Trim(xYSSISAA0.SSISAAPRFK)
fgSelect.Col = 6: fgSelect.Text = Trim(xYSSISAA0.SSISAAPRFX)
fgSelect.Col = 7
fgSelect.Text = xYSSISAA0.SSISAAYFCT & " - " & Trim(xYSSISAA0.SSISAAYUSR) _
    & " - " & dateImp10_S(xYSSISAA0.SSISAAYAMJ) & " - " & timeImp8(xYSSISAA0.SSISAAYHMS) & " - V" & xYSSISAA0.SSISAAYVER

fgSelect.Col = 9: fgSelect.Text = "SAA|" & xYSSISAA0.SSISAANAT & "|" _
                                & Trim(xYSSISAA0.SSISAAUIDX) & "|" & xYSSISAA0.SSISAAUSEQ & "|" & xYSSISAA0.SSISAAYVER & "|"

End Sub

Public Sub fgSelect_Display_H_WIN_Line()
On Error Resume Next

fgSelect.Col = 0
fgSelect.Text = xYSSIWIN0.SSIWINNAT
'fgSelect.Col = 2: fgSelect.Text = xYSSIWIN0.SSIWINUSEQ
If xYSSIWIN0.SSIWINUIDD = 0 Then
    fgSelect.Col = 1: fgSelect.Text = Trim(xYSSIWIN0.SSIWINUIDX)
Else
    fgSelect.Col = 1: fgSelect.Text = Trim(xYSSIWIN0.SSIWINUIDX) & " = " & xYSSIWIN0.SSIWINUIDD
End If
fgSelect.Col = 3: fgSelect.Text = " " & xYSSIWIN0.SSIWINSTAK
fgSelect.Col = 4: fgSelect.Text = Trim(xYSSIWIN0.SSIWINUNOM)
   X = ""


fgSelect.Col = 5: fgSelect.Text = Trim(xYSSIWIN0.SSIWINPRFK)
fgSelect.Col = 6: fgSelect.Text = Trim(xYSSIWIN0.SSIWINPRFX)
fgSelect.Col = 7
fgSelect.Text = xYSSIWIN0.SSIWINYFCT & " - " & Trim(xYSSIWIN0.SSIWINYUSR) _
    & " - " & dateImp10_S(xYSSIWIN0.SSIWINYAMJ) & " - " & timeImp8(xYSSIWIN0.SSIWINYHMS) & " - V" & xYSSIWIN0.SSIWINYVER

fgSelect.Col = 9: fgSelect.Text = "WIN|" & xYSSIWIN0.SSIWINNAT & "|" _
                                & Trim(xYSSIWIN0.SSIWINUIDX) & "|" & xYSSIWIN0.SSIWINYVER & "|"

End Sub
Public Sub fgSelect_Display_H_DIV_Line()
On Error Resume Next

fgSelect.Col = 0
fgSelect.Text = xYSSIDIV0.SSIDIVNAT
fgSelect.Col = 2: fgSelect.Text = xYSSIDIV0.SSIDIVDIDK
If xYSSIDIV0.SSIDIVUIDD = 0 Then
    fgSelect.Col = 1: fgSelect.Text = Trim(xYSSIDIV0.SSIDIVUIDX)
Else
    fgSelect.Col = 1: fgSelect.Text = Trim(xYSSIDIV0.SSIDIVUIDX) & " = " & xYSSIDIV0.SSIDIVUIDD
End If
fgSelect.Col = 3: fgSelect.Text = " " & xYSSIDIV0.SSIDIVSTAK
fgSelect.Col = 4: fgSelect.Text = Trim(xYSSIDIV0.SSIDIVUNOM)
   X = ""


fgSelect.Col = 5: fgSelect.Text = Trim(xYSSIDIV0.SSIDIVPRFK)
fgSelect.Col = 6: fgSelect.Text = Trim(xYSSIDIV0.SSIDIVPRFX)
fgSelect.Col = 7
fgSelect.Text = xYSSIDIV0.SSIDIVYFCT & " - " & Trim(xYSSIDIV0.SSIDIVYUSR) _
    & " - " & dateImp10_S(xYSSIDIV0.SSIDIVYAMJ) & " - " & timeImp8(xYSSIDIV0.SSIDIVYHMS) & " - V" & xYSSIDIV0.SSIDIVYVER

fgSelect.Col = 9: fgSelect.Text = "DIV|" & xYSSIDIV0.SSIDIVNAT & "|" _
                                & Trim(xYSSIDIV0.SSIDIVUIDX) & "|" & xYSSIDIV0.SSIDIVUIDD & "|" & "|" & xYSSIDIV0.SSIDIVYVER & "|"

End Sub



Public Sub fgSelect_Display_H_TIC_Line()
On Error Resume Next

fgSelect.Col = 0
fgSelect.Text = xYSSITIC0.SSITICNAT
If xYSSITIC0.SSITICNAT = "D" Then
    fgSelect.Col = 1: fgSelect.Text = Trim(xYSSITIC0.SSITICUIDX)

Else

    If xYSSITIC0.SSITICUIDD = 0 Then
        fgSelect.Col = 1: fgSelect.Text = Trim(xYSSITIC0.SSITICUIDX)
    Else
        fgSelect.Col = 1: fgSelect.Text = Trim(xYSSITIC0.SSITICUIDX) & " = " & xYSSITIC0.SSITICUIDD
    End If
End If


fgSelect.Col = 3: fgSelect.Text = " " & xYSSITIC0.SSITICSTAK
fgSelect.Col = 4: fgSelect.Text = Trim(xYSSITIC0.SSITICUNOM)
   X = ""


fgSelect.Col = 5: fgSelect.Text = Trim(xYSSITIC0.SSITICPRFK)
fgSelect.Col = 6: fgSelect.Text = Trim(xYSSITIC0.SSITICPRFX)
fgSelect.Col = 7
fgSelect.Text = xYSSITIC0.SSITICYFCT & " - " & Trim(xYSSITIC0.SSITICYUSR) _
    & " - " & dateImp10_S(xYSSITIC0.SSITICYAMJ) & " - " & timeImp8(xYSSITIC0.SSITICYHMS) & " - V" & xYSSITIC0.SSITICYVER

fgSelect.Col = 9: fgSelect.Text = "TIC|" & xYSSITIC0.SSITICNAT & "|" _
                                & Trim(xYSSITIC0.SSITICUIDX) & "|" & xYSSITIC0.SSITICUIDD & "|" & xYSSITIC0.SSITICYVER & "|"

End Sub

Public Sub fgSelect_Display_H_MEL_Line()
On Error Resume Next

fgSelect.Col = 0
fgSelect.Text = xYSSIMEL0.SSIMELNAT
If xYSSIMEL0.SSIMELUIDD = 0 Then
    fgSelect.Col = 1: fgSelect.Text = Trim(xYSSIMEL0.SSIMELUIDX)
Else
    fgSelect.Col = 1: fgSelect.Text = Trim(xYSSIMEL0.SSIMELUIDX) & " = " & xYSSIMEL0.SSIMELUIDD
End If
fgSelect.Col = 3: fgSelect.Text = " " & xYSSIMEL0.SSIMELSTAK
fgSelect.Col = 4: fgSelect.Text = Trim(xYSSIMEL0.SSIMELUNOM)
   X = ""


fgSelect.Col = 5: fgSelect.Text = Trim(xYSSIMEL0.SSIMELPRFK)
fgSelect.Col = 6: fgSelect.Text = Trim(xYSSIMEL0.SSIMELPRFX)
fgSelect.Col = 7
fgSelect.Text = xYSSIMEL0.SSIMELYFCT & " - " & Trim(xYSSIMEL0.SSIMELYUSR) _
    & " - " & dateImp10_S(xYSSIMEL0.SSIMELYAMJ) & " - " & timeImp8(xYSSIMEL0.SSIMELYHMS) & " - V" & xYSSIMEL0.SSIMELYVER

fgSelect.Col = 9: fgSelect.Text = "MEL|" & xYSSIMEL0.SSIMELNAT & "|" _
                                & Trim(xYSSIMEL0.SSIMELUIDX) & "|" & xYSSIMEL0.SSIMELUIDD & "|" & xYSSIMEL0.SSIMELYVER & "|"

End Sub





Public Sub fgSelect_Display_H_IBM_Line()
On Error Resume Next
Dim K As Integer
fgSelect.Col = 0
fgSelect.Text = xYSSIIBM0.SSIIBMNAT
fgSelect.Col = 2: fgSelect.Text = xYSSIIBM0.SSIIBMUIDD
fgSelect.Col = 1: fgSelect.Text = Trim(xYSSIIBM0.UPUPRF)
'fgSelect.Col = 3: fgSelect.Text = " " & xYSSIIBM0.SSIIBMSTAK
fgSelect.Col = 4: fgSelect.Text = Trim(xYSSIIBM0.UPTEXT)
fgSelect.Col = 5: fgSelect.Text = Trim(xYSSIIBM0.SSIIBMPRFK)
'fgSelect.Col = 6: fgSelect.Text = Trim(xYSSIIBM0.SSIIBMPRFX)
fgSelect.Col = 7
fgSelect.Text = xYSSIIBM0.SSIIBMYFCT & " - " & Trim(xYSSIIBM0.SSIIBMYUSR) _
    & " - " & dateImp10_S(xYSSIIBM0.SSIIBMYAMJ) & " - " & timeImp8(xYSSIIBM0.SSIIBMYHMS) & " - V" & xYSSIIBM0.SSIIBMYVER


fgSelect.Col = 9: fgSelect.Text = "IBM|" & xYSSIIBM0.SSIIBMNAT & "|" _
                                & xYSSIIBM0.SSIIBMUIDD & "|" & xYSSIIBM0.SSIIBMYVER & "|"

End Sub


Public Sub fgSelect_Display_J_Line()
 On Error Resume Next
Dim K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, X As String, X1 As String, blnExit As Boolean
Dim xField As String, blnRow_BIM As Boolean, blnRow_Next As Boolean
Dim xCol8 As String, xCol9 As String
Dim kCol As Integer
'fgSelect.Col = 0: fgSelect.Text = xYSSITXT0.SSITXTUIDN & " - " & xYSSITXT0.SSITXTDIDX & " - " & xYSSITXT0.SSITXTDIDS & " - " & xYSSITXT0.SSITXTUIDD
fgSelect.Col = 8: fgSelect.Text = xCol8
xCol8 = xYSSITXT0.SSITXTNAT & "|" & xYSSITXT0.SSITXTUIDN & "|" & Trim(xYSSITXT0.SSITXTDIDX) & "|" _
              & Trim(xYSSITXT0.SSITXTUIDX) & "|" & xYSSITXT0.SSITXTUIDD & "|" & xYSSITXT0.SSITXTTLNK & "|"
fgSelect.Col = 8: fgSelect.Text = xCol8
fgSelect.Col = 6
fgSelect.Text = dateImp10_S(xYSSITXT0.SSITXTYAMJ) & " - " & timeImp8(xYSSITXT0.SSITXTYHMS) & " - " & xYSSITXT0.SSITXTTLNK _
                & " - " & Trim(xYSSITXT0.SSITXTYUSR)
fgSelect.CellFontSize = 8: fgSelect.CellForeColor = RGB(96, 96, 96)
fgSelect.Col = 3: fgSelect.CellBackColor = mColor_Y0
X = Trim(xYSSITXT0.SSITXTINFO) & ">"
K4 = 0
blnExit = False: blnRow_Next = False
Do
    blnRow_BIM = False
    K1 = InStr(K4 + 1, X, "<")
    If K1 > 0 Then
        K2 = InStr(K1 + 1, X, ":")
        K3 = InStr(K1 + 1, X, "|")
        K4 = InStr(K1 + 1, X, ">")
        If K3 = 0 Or K3 > K4 Then K3 = K4
        Select Case Mid$(X, K1, K2 - K1)
            Case "<UID": X1 = Mid$(X, K2 + 1, K4 - K2 - 1)
                fgSelect.Col = 2: fgSelect.Text = X1: fgSelect.CellForeColor = vbBlue ': fgSelect.CellBackColor = mColor_Y1
            Case "<FCT": fgSelect.Col = 1: fgSelect.Text = Mid$(X, K2 + 1, K4 - K2 - 1)
                    Select Case fgSelect.Text
                        Case "CRE": fgSelect.CellForeColor = RGB(0, 96, 0)
                            For kCol = 0 To 9: fgSelect.Col = kCol: fgSelect.CellBackColor = mColor_G0: Next kCol
                        Case "SUP", "9-???": fgSelect.CellForeColor = RGB(96, 0, 0)
                            For kCol = 0 To 9: fgSelect.Col = kCol: fgSelect.CellBackColor = mColor_W0: Next kCol
                        Case "9-SAA", "9-MEL", "9-TIC", "9-IBM", "9-WIN", "9-UGM": fgSelect.CellForeColor = vbBlue
                            For kCol = 0 To 9: fgSelect.Col = kCol: fgSelect.CellBackColor = mColor_B0: Next kCol
                        Case Else:
                                fgSelect.CellBackColor = mColor_Y0: fgSelect.CellForeColor = vbMagenta
                    End Select
            Case "<ORIG": X1 = Mid$(X, K2 + 1, K4 - K2 - 1)
                fgSelect.Col = 0: fgSelect.Text = arrJRN_Origine(Val(X1)): fgSelect.CellBackColor = mColor_Y0
                If X1 = "6" Then fgSelect.CellBackColor = vbYellow
            Case "<UNOM": X1 = Mid$(X, K2 + 1, K4 - K2 - 1)
                fgSelect.Col = 4: fgSelect.Text = X1: fgSelect.CellBackColor = mColor_Y0
            Case "<X": X1 = Mid$(X, K2 + 1, K4 - K2 - 1): blnRow_Next = True
                 fgSelect.Col = 4: fgSelect.Text = X1 ': fgSelect.CellBackColor = mColor_Y0
            Case "<Y": xCol9 = Mid$(X, K2 + 1, K4 - K2 - 1)
                 fgSelect.Col = 9: fgSelect.Text = xCol9
           Case "<STAK": xField = "code état": blnRow_BIM = True
            Case "<UIDX": xField = "Identifiant": blnRow_BIM = True
            Case "<DECH": xField = "Echéance": blnRow_BIM = True
            Case "<PRFK": xField = "Conforme ?": blnRow_BIM = True
            Case "<PRFX": xField = "Profil": blnRow_BIM = True
            Case "<UIDD": xField = "Id " & xYSSITXT0.SSITXTDIDX
            
            Case "<X":
             xField = "" '"commentaire"
           Case Else: xField = Mid$(X, K1 + 1, K2 - K1 - 1)
        End Select
        If blnRow_BIM Then
            If blnRow_Next Then
                fgSelect.Rows = fgSelect.Rows + 1
                fgSelect.Row = fgSelect.Rows - 1
            Else
                blnRow_Next = True
            End If
            fgSelect.Col = 3: fgSelect.Text = xField ': fgSelect.CellBackColor = mColor_Y0
            fgSelect.Col = 5: fgSelect.Text = Mid$(X, K2 + 1, K3 - K2 - 1): fgSelect.CellForeColor = vbMagenta
            fgSelect.Col = 4: fgSelect.Text = Mid$(X, K3 + 1, K4 - K3 - 1)
            fgSelect.Col = 8: fgSelect.Text = xCol8
            fgSelect.Col = 9: fgSelect.Text = xCol9
        End If
        

    Else
       blnExit = True
    End If
    
Loop Until blnExit


End Sub


Public Sub fgSelect_Display_1_Line_YSSIDOM0()
On Error Resume Next
Dim K As Integer, wColor As Long

Select Case Trim(xYSSIDOM0.SSIDOMDIDX)
    Case "IBM": wColor = RGB(0, 96, 0)
    Case "SAA": wColor = vbMagenta
    Case "SAB": wColor = vbBlue
    Case "WIN": wColor = RGB(255, 32, 0)
    Case "TIC": wColor = RGB(255, 32, 0)
End Select


If xYSSIDOM0.SSIDOMSTAK = "N" Then
    For K = 0 To 9: fgSelect.Col = K: fgSelect.CellForeColor = RGB(96, 96, 96): Next K
Else
    For K = 0 To 9: fgSelect.Col = K: fgSelect.CellForeColor = wColor: Next K
End If
fgSelect.Col = 1: fgSelect.Text = "   " & Trim(xYSSIDOM0.SSIDOMDIDX): fgSelect.CellFontBold = True
fgSelect.Col = 2: fgSelect.Text = "            " & Trim(xYSSIDOM0.SSIDOMUIDX): fgSelect.CellFontBold = True
fgSelect.Col = 3
fgSelect.Text = " " & xYSSIDOM0.SSIDOMSTAK
Select Case xYSSIDOM0.SSIDOMSTAK
    Case " "
    Case "N": fgSelect.CellBackColor = RGB(192, 192, 192)
    Case Else: fgSelect.CellBackColor = mColor_W1
End Select

If xYSSIDOM0.SSIDOMDECH <> 0 Then
    fgSelect.Col = 4: fgSelect.Text = "  " & dateImp10_S(xYSSIDOM0.SSIDOMDECH)
    If xYSSIDOM0.SSIDOMSTAK = " " And xYSSIDOM0.SSIDOMDECH < DSys Then fgSelect.CellBackColor = mColor_W1
End If
fgSelect.Col = 5: fgSelect.Text = xYSSIDOM0.SSIDOMPRFX
If xYSSIDOM0.SSIDOMPRFD = 0 Then
    fgSelect.Col = 6: fgSelect.Text = xYSSIDOM0.SSIDOMPRFK
Else
    fgSelect.Col = 6: fgSelect.Text = xYSSIDOM0.SSIDOMPRFK & "  " & dateImp10_S(xYSSIDOM0.SSIDOMPRFD) & "  " & timeImp8(xYSSIDOM0.SSIDOMPRFH)
End If
Select Case xYSSIDOM0.SSIDOMPRFK
    Case " "
    Case "N": fgSelect.CellBackColor = mColor_W0
    Case "X": fgDetail.CellBackColor = RGB(192, 192, 192)
    Case "!": fgSelect.CellBackColor = mColor_Y1
    Case Else: fgSelect.CellBackColor = mColor_W1
End Select
If xYSSIDOM0.SSIDOMTLNK > 0 Then fgSelect.Col = 7: fgSelect.Text = xYSSIDOM0.SSIDOMTLNK
fgSelect.Col = 8: fgSelect.Text = xYSSIDOM0.SSIDOMYFCT & " par " & xYSSIDOM0.SSIDOMYUSR _
  & " le " & dateImp10_S(xYSSIDOM0.SSIDOMYAMJ) & "  " & timeImp8(xYSSIDOM0.SSIDOMYHMS) & " (" & xYSSIDOM0.SSIDOMYVER & ")"
fgSelect.CellFontSize = 8
fgSelect.CellForeColor = RGB(128, 128, 128)


End Sub


Public Function cmdUpdate()
Dim K As Integer, xSQL As String, X As String
Dim V
Dim blnYBIADTAQ As Boolean
Dim wSSIUSRUIDN As Long

On Error GoTo Error_Handler


If mYSSIUSR0_Update = "New" And newYSSIUSR0.SSIUSRUIDN = 0 Then
    xSQL = "select SSIUSRUIDN from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
         & "  Where SSIUSRNAT = '" & newYSSIUSR0.SSIUSRNAT & "' order by SSIUSRUIDN desc FETCH FIRST 1 ROWS ONLY"
    Set rsSab = cnsab.Execute(xSQL)
    
    If rsSab.EOF Then
        If newYSSIUSR0.SSIUSRNAT = "$" Then
            newYSSIUSR0.SSIUSRUIDN = 1
        Else
            newYSSIUSR0.SSIUSRUIDN = 1001
        End If
    Else
        newYSSIUSR0.SSIUSRUIDN = rsSab("SSIUSRUIDN") + 1
    End If
    newYSSITXT0.SSITXTUIDN = newYSSIUSR0.SSIUSRUIDN
    newYSSIDOM0.SSIDOMUIDN = newYSSIUSR0.SSIUSRUIDN
    newYSSITXT0_JRN.SSITXTUIDN = newYSSIUSR0.SSIUSRUIDN
    wSSIUSRUIDN = newYSSIDOM0.SSIDOMUIDN
Else
     wSSIUSRUIDN = oldYSSIDOM0.SSIDOMUIDN
End If
        


If mYSSIDOM0_Update = "New" Then
    X = "select SSIDOMYVER from " & paramIBM_Library_SABSPE & ".YSSIDOMH " _
         & "  Where SSIDOMNAT = '" & newYSSIDOM0.SSIDOMNAT & "' and SSIDOMUIDN = " & newYSSIDOM0.SSIDOMUIDN _
         & " and SSIDOMDIDX = '" & newYSSIDOM0.SSIDOMDIDX & "'  and SSIDOMUIDX = '" & Replace(newYSSIDOM0.SSIDOMUIDX, "'", "''") & "'" _
         & " order by SSIDOMYVER desc"
    Set rsSab = cnsab.Execute(X)

    If Not rsSab.EOF Then newYSSIDOM0.SSIDOMYVER = rsSab("SSIDOMYVER") + 1
End If

If mYSSITXT0_Update = "New" Then
    xSQL = "select SSITXTTLNK from " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
         & "  Where SSITXTNAT = '" & newYSSITXT0.SSITXTNAT & "'" _
         & "  and SSITXTUIDN = " & newYSSITXT0.SSITXTUIDN _
         & "  and SSITXTDIDX = '" & newYSSITXT0.SSITXTDIDX & "'" _
         & "  and SSITXTUIDX = '" & Replace(newYSSITXT0.SSITXTUIDX, "'", "''") & "'" _
         & "  and SSITXTUIDD = " & newYSSITXT0.SSITXTUIDD _
         & " order by SSITXTTLNK desc"
    Set rsSab = cnsab.Execute(xSQL)
    
    If rsSab.EOF Then
        newYSSITXT0.SSITXTTLNK = 1
    Else
        newYSSITXT0.SSITXTTLNK = rsSab("SSITXTTLNK") + 1
    End If
    newYSSIUSR0.SSIUSRTLNK = newYSSITXT0.SSITXTTLNK
    newYSSIDOM0.SSIDOMTLNK = newYSSITXT0.SSITXTTLNK
    newYSSIIBMH.SSIIBMTLNK = newYSSITXT0.SSITXTTLNK
    newYSSISAAH.SSISAATLNK = newYSSITXT0.SSITXTTLNK
    newYSSIWINH.SSIWINTLNK = newYSSITXT0.SSITXTTLNK
    newYSSIDIVH.SSIDIVTLNK = newYSSITXT0.SSITXTTLNK
End If

If mYSSIIBM0_Update = "New$" Then
    xSQL = "select SSIIBMUIDD from " & paramIBM_Library_SABSPE & ".YSSIIBM0 " _
         & "  Where SSIIBMNAT = '$'  order by SSIIBMUIDD desc FETCH FIRST 1 ROWS ONLY"
    Set rsSab = cnsab.Execute(xSQL)
    
    If rsSab.EOF Then
            newYSSIIBM0.SSIIBMUIDD = 1
    Else
        newYSSIIBM0.SSIIBMUIDD = rsSab("SSIIBMUIDd") + 1
    End If
End If
'________________________________________________________________________________

If mYSSITXT0_JRN_Update = "New" Then
    xSQL = "select SSITXTTLNK from " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
         & "  Where SSITXTNAT = 'J'" _
         & " order by SSITXTTLNK desc FETCH FIRST 1 ROWS ONLY"
    Set rsSab = cnsab.Execute(xSQL)
    
    If rsSab.EOF Then
        newYSSITXT0_JRN.SSITXTTLNK = 1
    Else
        newYSSITXT0_JRN.SSITXTTLNK = rsSab("SSITXTTLNK") + 1
    End If
End If

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

'________________________________________________________________________________

Select Case mYSSIUSR0_Update
    Case ""
    Case "Update+H":
                V = sqlYSSIUSR0_Update(newYSSIUSR0, oldYSSIUSR0)
                If IsNull(V) Then V = sqlYSSIUSRH_Insert(oldYSSIUSR0)
    Case "Update":
                V = sqlYSSIUSR0_Update(newYSSIUSR0, oldYSSIUSR0)
    Case "New": V = sqlYSSIUSR0_Insert(newYSSIUSR0)
    Case "CMD": V = sqlYSSIUSR0_Update_CMD(mYSSIUSR0_Update_CMD)
    
    Case "Delete": V = sqlYSSIUSR0_Delete(oldYSSIUSR0)
    
    Case "Delete+CMD": V = sqlYSSIUSR0_Delete(oldYSSIUSR0)
            If IsNull(V) Then V = sqlYSSIUSR0_Update_CMD(mYSSIUSR0_Update_CMD)
End Select

If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

Select Case mYSSIDOM0_Update
    Case ""
    Case "Update+H": wSSIUSRUIDN = newYSSIDOM0.SSIDOMUIDN
                V = sqlYSSIDOM0_Update(newYSSIDOM0, oldYSSIDOM0)
                If IsNull(V) Then V = sqlYSSIDOMH_Insert(oldYSSIDOM0)
    Case "Update": wSSIUSRUIDN = newYSSIDOM0.SSIDOMUIDN
                V = sqlYSSIDOM0_Update(newYSSIDOM0, oldYSSIDOM0)
    Case "New": wSSIUSRUIDN = newYSSIDOM0.SSIDOMUIDN
                V = sqlYSSIDOM0_Insert(newYSSIDOM0)
    Case "Delete": wSSIUSRUIDN = oldYSSIDOM0.SSIDOMUIDN
                V = sqlYSSIDOM0_Delete(oldYSSIDOM0)
    Case "Delete+H": wSSIUSRUIDN = oldYSSIDOM0.SSIDOMUIDN
                V = sqlYSSIDOM0_Delete(oldYSSIDOM0)
                If IsNull(V) Then V = sqlYSSIDOMH_Insert(oldYSSIDOM0)
                ''If IsNull(V) Then V = sqlYSSIDOMH_Insert(newYSSIDOM0)
                
    Case "CMD": V = sqlYSSIDOM0_Update_CMD(mYSSIDOM0_Update_CMD)
    
    Case "PRFX$":
        Dim wSSIDOMUIDD As Long

        X = "select SSIDOMuidd from " & paramIBM_Library_SABSPE & ".YSSIDOMH " _
             & "  Where SSIDOMNAT = '" & newYSSIDOM0.SSIDOMNAT & "' and SSIDOMUIDN = " & wSSIUSRUIDN _
             & " and SSIDOMuidd < 0 order by SSIDOMUIDD"
        Set rsSab = cnsab.Execute(X)
        If Not rsSab.EOF Then wSSIDOMUIDD = rsSab("SSIDOMUIDD")
        For K = 1 To arrYSSIDOM0_BIA_Nb
            If arrYSSIDOM0_BIA(K).SSIDOMNAT = " " Then
                arrYSSIDOM0_BIA(K).SSIDOMUIDN = wSSIUSRUIDN
                arrYSSIDOM0_BIA(K).SSIDOMUNIT = newYSSIUSR0.SSIUSRUNIT
                X = "select SSIDOMYVER from " & paramIBM_Library_SABSPE & ".YSSIDOMH " _
                     & "  Where SSIDOMNAT = '" & newYSSIDOM0.SSIDOMNAT & "' and SSIDOMUIDN = " & newYSSIDOM0.SSIDOMUIDN _
                     & " and SSIDOMDIDX = '" & newYSSIDOM0.SSIDOMDIDX & "'  and SSIDOMUIDX = '" & newYSSIDOM0.SSIDOMUIDX & "'" _
                     & " order by SSIDOMYVER desc"
                Set rsSab = cnsab.Execute(X)
            
                If Not rsSab.EOF Then arrYSSIDOM0_BIA(K).SSIDOMYVER = rsSab("SSIDOMYVER") + 1
                wSSIDOMUIDD = wSSIDOMUIDD - 1
                arrYSSIDOM0_BIA(K).SSIDOMUIDD = wSSIDOMUIDD
                V = sqlYSSIDOM0_Insert(arrYSSIDOM0_BIA(K))
                If Not IsNull(V) Then GoTo Error_MsgBox
            End If
        Next K
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
If mYSSIDOM0_Update_CMD_2 <> "" Then
    V = sqlYSSIDOM0_Update_CMD(mYSSIDOM0_Update_CMD_2)
    If Not IsNull(V) Then GoTo Error_MsgBox
End If

'________________________________________________________________________________
Select Case mYSSITXT0_Update
    Case ""
    Case "Update":     V = sqlYSSITXT0_Update(newYSSITXT0, oldYSSITXT0_XXX)
    Case "New", "NewP": V = sqlYSSITXT0_Insert(newYSSITXT0)
    Case "Delete": V = sqlYSSITXT0_Delete(newYSSITXT0)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox

'________________________________________________________________________________
Select Case mYSSIIBM0_Update
    Case ""
    Case "Update":     V = sqlYSSIIBM0_Update(newYSSIIBM0, oldYSSIIBM0)
    Case "New$": V = sqlYSSIIBM0_Profil_Insert(newYSSIIBM0, oldYSSIIBM0)
    'Case "Delete": V = sqlYSSIibm0_Delete(oldYSSITXT0_XXX)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYSSIIBMH_Update
     Case ""
   Case "Update":     V = sqlYSSIIBMH_Update(newYSSIIBMH, oldYSSIIBMH)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYSSITXT0_JRN_Update
    Case ""
    Case "New": V = sqlYSSITXT0_Insert(newYSSITXT0_JRN)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

Select Case mYSSISAA0_Update
    Case ""
    Case "Update+H":
                V = sqlYSSISAA0_Update(newYSSISAA0, oldYSSISAA0)
                If IsNull(V) Then V = sqlYSSISAAH_Insert(oldYSSISAA0)
    Case "Update":
                V = sqlYSSISAA0_Update(newYSSISAA0, oldYSSISAA0)
    Case "New": V = sqlYSSISAA0_Insert(newYSSISAA0)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYSSISAAH_Update
     Case ""
   Case "Update":     V = sqlYSSISAAH_Update(newYSSISAAH, oldYSSISAAH)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYSSISAB0_Update
    Case ""
    Case "Update+H":
                V = sqlYSSISAB0_Update(newYSSISAB0, oldYSSISAB0)
                If IsNull(V) Then V = sqlYSSISABH_Insert(oldYSSISAB0)
    Case "Update":
                V = sqlYSSISAB0_Update(newYSSISAB0, oldYSSISAB0)
    Case "New": V = sqlYSSISAB0_Insert(newYSSISAB0)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYSSISABH_Update
    Case ""
    Case "Update":     V = sqlYSSISABH_Update(newYSSISABH, oldYSSISABH)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYSSIWIN0_Update
    Case ""
    Case "Update+H":
                V = sqlYSSIWIN0_Update(newYSSIWIN0, oldYSSIWIN0)
                If IsNull(V) Then V = sqlYSSIWINH_Insert(oldYSSIWIN0)
    Case "Update":
                V = sqlYSSIWIN0_Update(newYSSIWIN0, oldYSSIWIN0)
    Case "New": V = sqlYSSIWIN0_Insert(newYSSIWIN0)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYSSIWINH_Update
    Case ""
    Case "Update":     V = sqlYSSIWINH_Update(newYSSIWINH, oldYSSIWINH)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYSSIDIV0_Update
    Case ""
    Case "Update+H":
                V = sqlYSSIDIV0_Update(newYSSIDIV0, oldYSSIDIV0)
                If IsNull(V) Then V = sqlYSSIDIVH_Insert(oldYSSIDIV0)
    Case "Update":
                V = sqlYSSIDIV0_Update(newYSSIDIV0, oldYSSIDIV0)
    Case "New": V = sqlYSSIDIV0_Insert(newYSSIDIV0)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYSSIDIVH_Update
    Case ""
    Case "Update":     V = sqlYSSIDIVH_Update(newYSSIDIVH, oldYSSIDIVH)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYSSIMEL0_Update
    Case ""
    Case "Update+H":
                V = sqlYSSIMEL0_Update(newYSSIMEL0, oldYSSIMEL0)
                If IsNull(V) Then V = sqlYSSIMELH_Insert(oldYSSIMEL0)
    Case "Update":
                V = sqlYSSIMEL0_Update(newYSSIMEL0, oldYSSIMEL0)
    Case "New": V = sqlYSSIMEL0_Insert(newYSSIMEL0)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYSSIMELH_Update
    Case ""
    Case "Update":     V = sqlYSSIMELH_Update(newYSSIMELH, oldYSSIMELH)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYSSITIC0_Update
    Case ""
    Case "Update+H":
                V = sqlYSSITIC0_Update(newYSSITIC0, oldYSSITIC0)
                If IsNull(V) Then V = sqlYSSITICH_Insert(oldYSSITIC0)
    Case "Update":
                V = sqlYSSITIC0_Update(newYSSITIC0, oldYSSITIC0)
    Case "New": V = sqlYSSITIC0_Insert(newYSSITIC0)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYSSITICH_Update
    Case ""
    Case "Update":     V = sqlYSSITICH_Update(newYSSITICH, oldYSSITICH)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

If wSSIUSRUIDN > 0 Then
    V = cmdUpdate_SSIUSRPRFK_DOM(wSSIUSRUIDN)
    If Not IsNull(V) Then GoTo Error_MsgBox
End If
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdUpdate"
Exit_sub:

    cmdUpdate = V
    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        

    End If
    Call cmdUpdate_Init
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    

End Function


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim wFct As String

mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

If cboSelect_SQL.ListCount = 0 Then
    wFct = UCase$(Trim(Mid$(Msg, 1, 12)))
    Call BIA_VB_HAB(wFct, arrHab(), cboSelect_SQL)
    
   ' paramFile_Init
    
    Call Form_Init
    
    mMail_Destinataires = currentSSIWINMAIL

    Select Case wFct
        Case "@BIA_SSI":    blnAuto = True
                        If Not blnBIA_SSI_Automate Then
                            
                           Call cmdUpdate_Init
                           If blnTimer_Enabled Then
                                X = "<ORIG:6><FCT:ON><UID:TIMER BIA_Audit @BIA_SSI><X:" & usrIdNT & ">"
                                blnBIA_SSI_Automate = True
                            Else
                                X = "<ORIG:6><FCT:exe><UID:BIA_Audit @BIA_SSI><X:" & usrIdNT & ">"
                            End If
                            Call cmdSSIJRN_TXT("@SSI", X)
                            newYSSITXT0_JRN.SSITXTUIDX = "BIA_Audit"
                            Call cmdUpdate
                            
                        End If
                        
                        cmdSelect_SQL_K = "9_IBM"
                        cmdSelect_Ok_Click
                        cmdSelect_SQL_K = "9_SAB"
                        cmdSelect_Ok_Click
                        cmdSelect_SQL_K = "9_SAA"
                        cmdSelect_Ok_Click
                        cmdSelect_SQL_K = "9_WIN"
                        cmdSelect_Ok_Click
                        cmdSelect_SQL_K = "9_MEL"
                        cmdSelect_Ok_Click
                        Unload Me
        Case "@BIA_SSI_JRN":    blnAuto = True
                        mMail_Destinataires = srvSendMail.Exchange_Distribution("BIA_SSI", "@BIA_SSI_JRN")

                        cmdUpdate_PRFK_DECH
                        cmdSelect_SQL_K = "3"
                        cmdSelect_Ok_Click
                        mnuPrint_Mail_Click
                        
                        Call DTPicker_Set(txtSelect_Options_J_SSITXTYMAJ, YBIATAB0_DATE_CPT_JP0)
                        cmdSelect_SQL_K = "J"
                        cmdSelect_Ok_Click
                        mnuPrint_Mail_Click
                       Unload Me
        Case Else: blnAuto = False: cboSelect_SQL.ListIndex = 0
    
    End Select
End If
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


Private Sub cboProfil_DOM_Click()
mSSIDOMDIDX = Trim(cboProfil_DOM.Text)
txtRTF.Visible = False
If Not cboProfil_DOM.Locked Then
    chkProfil_DOM.Visible = True
    Select Case mSSIDOMDIDX
        Case "IBM": Call fgProfil_Display_IBM
        Case "SAA": Call fgProfil_Display_SAA
        Case "SAB": Call fgProfil_Display_SAB
        Case "WIN": Call fgProfil_Display_WIN
        Case "DIV": Call fgProfil_Display_DIV
        Case "MEL": Call fgProfil_Display_MEL
        Case "TIC": Call fgProfil_Display_TIC
    End Select
End If
End Sub


Private Sub fgProfil_Display_SAB()

Dim xSQL As String, xWhere As String


On Error GoTo Error_Handler
currentAction = currentAction & "-> fgProfil_Display_SAB"
fraProfil_Update.Visible = False
fraProfil_Update_DIV.Visible = False
fgProfil.Visible = False
fgProfil_Reset
fgProfil.FormatString = fgProfil_FormatString
fgProfil.Rows = 1
fgProfil.Row = 0
If chkProfil_DOM <> "1" Then xWhere = " and SSISABSTAK = ' '"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
     & " where SSISABNAT = '$'" & xWhere _
     & " order by SSISABUIDX"

Set rsSab = cnsab.Execute(xSQL)
  

Do While Not rsSab.EOF
    fgProfil.Rows = fgProfil.Rows + 1
    fgProfil.Row = fgProfil.Rows - 1
    fgProfil_Display_SAB_Line
    rsSab.MoveNext
Loop

fgProfil.Visible = True: fgProfil.Enabled = True

If cmdSelect_SQL_K = "2_D" Then
    cmdProfil_Print.Visible = arrHab(3): cmdProfil_Excel.Visible = arrHab(3)
Else
    cmdProfil_Print.Visible = False: cmdProfil_Excel.Visible = False
End If


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgProfil.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub fgProfil_Display_SAB_Line()
On Error Resume Next

'fgProfil.Col = 0: fgProfil.Text = rsSab("SSISABUIDD")
fgProfil.Col = 1: fgProfil.Text = rsSab("SSISABUIDX")
fgProfil.CellFontBold = True: fgProfil.CellForeColor = vbBlue
If rsSab("SSISABSTAK") <> " " Then fgProfil.CellBackColor = RGB(192, 192, 192)
fgProfil.Col = 2: fgProfil.Text = Mid$(rsSab("SSISABUNOM"), 1, 10) & "  " & Mid$(rsSab("SSISABUNOM"), 11, 10) & "  " & Mid$(rsSab("SSISABUNOM"), 21, 10)

End Sub


Private Sub cboSelect_Options_1_SSIDOMDIDX_Change()
cboSelect_Options_1_SSIDOMPRFX_Init
End Sub

Private Sub cboSelect_Options_1_SSIDOMDIDX_Click()
cboSelect_Options_1_SSIDOMPRFX_Init
End Sub


Private Sub cboSelect_Options_1_SSIDOMPRFX_Click()
cmdSelect_Clear
End Sub

Private Sub cboSelect_Options_1_SSIUSRSTAK_Click()
cmdSelect_Clear
End Sub

Private Sub cboSelect_Options_4_SSIDOMDIDX_Change()
cmdSelect_Clear
cboSelect_Options_4_SSIDOMNAT_Init

End Sub

Private Sub cboSelect_Options_4_SSIDOMDIDX_Click()
cboSelect_Options_4_SSIDOMNAT_Init

End Sub


Private Sub cboSelect_Options_4_SSIDOMNAT_Click()
cmdSelect_Clear

End Sub

Private Sub cboSelect_Options_J_SSIDOMDIDX_Click()
cmdSelect_Clear

End Sub

Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

End Sub


Private Sub cboSSIUSRPRFK_Click()
cmdSSIUSR_PRF.Visible = False

End Sub

Private Sub cboSSIUSRPRFX_Click()
cmdSSIUSR_PRF.Visible = False
If cboSSIUSRPRFX <> "" And cboSSIUSRUNIT.Text = "" Then
    X = "select SSIUSRUNIT from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
         & " where SSIUSRNAT = '$' and SSIUSRUIDX = '" & Trim(cboSSIUSRPRFX) & "'"
    Set rsSab = cnsab.Execute(X)
    If Not rsSab.EOF Then
            Call cbo_Scan(rsSab("SSIUSRUNIT"), cboSSIUSRUNIT)
    End If
End If

End Sub

Private Sub cboSSIUSRSTAK_Click()
cmdSSIUSR_PRF.Visible = False

End Sub

Private Sub chkProfil_DOM_Click()
If fgProfil.Visible Then cboProfil_DOM_Click
End Sub

Private Sub chkSSIDOMDECH_Click()
If chkSSIDOMDECH = "1" Then
    txtSSIDOMDECH.Visible = True
Else
    txtSSIDOMDECH.Visible = False
End If

End Sub

Private Sub chkSSIUSRDECH_Click()
If chkSSIUSRDECH = "1" Then
    txtSSIUSRDECH.Visible = True
Else
    txtSSIUSRDECH.Visible = False
End If
End Sub

Private Sub cmdCompte_All_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass

Select Case mSSIDOMDIDX
    Case "IBM"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0 " _
             & " where SSIIBMNAT = ' '" _
             & " and SSIIBMUIDD not in (select SSIDOMUIDD from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
             & " where SSIDOMDIDX = '" & mSSIDOMDIDX & "')"
        Set rsSab = cnsab.Execute(X)
        Call fgCompte_Display_IBM(vbMagenta)
    Case "SAA"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
             & " where SSISAANAT = ' ' and SSISAAPRFK = '?'" '_
             '& " and SSISAAUIDX not in (select SSIDOMUIDX from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
             '& " where SSIDOMDIDX = '" & mSSIDOMDIDX & "')"
        Set rsSab = cnsab.Execute(X)
        Call fgCompte_Display_SAA(vbMagenta)
    Case "SAB"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
             & " where SSISABNAT = ' '" _
             & " and SSISABUIDX not in (select SSIDOMUIDX from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
             & " where SSIDOMDIDX = '" & mSSIDOMDIDX & "')"
        Set rsSab = cnsab.Execute(X)
        Call fgCompte_Display_SAB(vbMagenta)
    Case "WIN"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
             & " where SSIWINNAT = ' ' and ( SSIWINPRFK = '?'" _
             & " or  SSIWINPRFK = ' ' and SSIWINUIDX not in (select SSIDOMUIDX from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
             & " where SSIDOMDIDX = '" & mSSIDOMDIDX & "') )"
        Set rsSab = cnsab.Execute(X)
        Call fgCompte_Display_WIN(vbMagenta)
    Case "DIV"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
             & " where SSIDIVNAT = ' ' and SSIDIVPRFK = '?'" _
             & " and SSIDIVDIDK = '" & prfYSSIDIV0.SSIDIVDIDK & "' order by SSIDIVUIDX"
             '& " and SSIDIVUIDX not in (select SSIDOMUIDX from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
             '& " where SSIDOMDIDX = '" & mSSIDOMDIDX & "') order by SSIDIVUIDX"
        Set rsSab = cnsab.Execute(X)
        Call fgCompte_Display_DIV(vbMagenta)
    Case "TIC"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSITIC0 " _
             & " where SSITICNAT = ' '" _
             & " and SSITICUIDX not in (select SSIDOMUIDX from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
             & " where SSIDOMDIDX = '" & mSSIDOMDIDX & "')"
        Set rsSab = cnsab.Execute(X)
        Call fgCompte_Display_TIC(vbMagenta)
End Select
    
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdCompteH_Quit_Click()
On Error Resume Next
fraCompteH.Visible = False
txtRTF.Visible = False
End Sub

Private Sub cmdCompteH_Update_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdUpdate_Init

newYSSITXT0 = oldYSSITXT0_Histo
newYSSITXT0.SSITXTINFO = Trim(txtCompteH_SSITXTINFO)
If Trim(oldYSSITXT0_Histo.SSITXTINFO) <> newYSSITXT0.SSITXTINFO Then
    mYSSITXT0_Update = "New"
    newYSSITXT0.SSITXTYAMJ = DSys
    newYSSITXT0.SSITXTYHMS = time_Hms
    newYSSITXT0.SSITXTYUSR = usrName_UCase
End If

Select Case Trim(oldYSSIDOM0.SSIDOMDIDX)
    Case "IBM"
        mYSSIIBMH_Update = "Update"
        newYSSIIBMH = oldYSSIIBMH
        If oldYSSIIBMH.SSIIBMYFCT <> "VU " Then
            newYSSIIBMH.SSIIBMYFCT = "VU "
            newYSSIIBMH.SSIIBMYUSR = usrName_UCase
            newYSSIIBMH.SSIIBMYAMJ = DSys
            newYSSIIBMH.SSIIBMYHMS = time_Hms
        End If
        Call cmdSSIJRN_IBM("<X:contrôle des modifications (Historique)>")
        If IsNull(cmdUpdate) Then
            fraCompteH_Display
        End If
    Case "SAA"
        mYSSISAAH_Update = "Update"
        newYSSISAAH = oldYSSISAAH
        If oldYSSISAAH.SSISAAYFCT <> "VU " Then
            newYSSISAAH.SSISAASTAK = " "
            newYSSISAAH.SSISAAYFCT = "VU "
            newYSSISAAH.SSISAAYUSR = usrName_UCase
            newYSSISAAH.SSISAAYAMJ = DSys
            newYSSISAAH.SSISAAYHMS = time_Hms
            If oldYSSISAAH.SSISAAPRFK = "?" Then newYSSISAAH.SSISAAPRFK = " "
        End If
        Call cmdSSIJRN_SAA("<X:contrôle des modifications (Historique)>")
        If IsNull(cmdUpdate) Then
            fraCompteH_Display
        End If
    Case "SAB", "SAB_W"
        mYSSISABH_Update = "Update"
        newYSSISABH = oldYSSISABH
        If oldYSSISABH.SSISABYFCT <> "VU " Then
            newYSSISABH.SSISABSTAK = "  "
            newYSSISABH.SSISABYFCT = "VU "
            newYSSISABH.SSISABYUSR = usrName_UCase
            newYSSISABH.SSISABYAMJ = DSys
            newYSSISABH.SSISABYHMS = time_Hms
        End If
        Call cmdSSIJRN_SAB("<X:contrôle des modifications (Historique)>")
        If IsNull(cmdUpdate) Then
            fraCompteH_Display
        End If
    Case "WIN"
        mYSSIWINH_Update = "Update"
        newYSSIWINH = oldYSSIWINH
        If oldYSSIWINH.SSIWINYFCT <> "VU " Then
            newYSSIWINH.SSIWINSTAK = " "
            newYSSIWINH.SSIWINYFCT = "VU "
            newYSSIWINH.SSIWINYUSR = usrName_UCase
            newYSSIWINH.SSIWINYAMJ = DSys
            newYSSIWINH.SSIWINYHMS = time_Hms
        End If
        Call cmdSSIJRN_WIN("<X:contrôle des modifications (Historique)>")
        If IsNull(cmdUpdate) Then
            fraCompteH_Display
        End If
    Case "DIV"
        mYSSIDIVH_Update = "Update"
        newYSSIDIVH = oldYSSIDIVH
        If oldYSSIDIVH.SSIDIVYFCT <> "VU " Then
            newYSSIDIVH.SSIDIVSTAK = " "
            newYSSIDIVH.SSIDIVYFCT = "VU "
            newYSSIDIVH.SSIDIVYUSR = usrName_UCase
            newYSSIDIVH.SSIDIVYAMJ = DSys
            newYSSIDIVH.SSIDIVYHMS = time_Hms
        End If
        Call cmdSSIJRN_DIV("<X:contrôle des modifications (Historique)>")
        If IsNull(cmdUpdate) Then
            fraCompteH_Display
        End If
    Case "MEL"
        mYSSIMELH_Update = "Update"
        newYSSIMELH = oldYSSIMELH
        If oldYSSIMELH.SSIMELYFCT <> "VU " Then
            newYSSIMELH.SSIMELSTAK = " "
            newYSSIMELH.SSIMELYFCT = "VU "
            newYSSIMELH.SSIMELYUSR = usrName_UCase
            newYSSIMELH.SSIMELYAMJ = DSys
            newYSSIMELH.SSIMELYHMS = time_Hms
        End If
        Call cmdSSIJRN_MEL("<X:contrôle des modifications (Historique)>")
        If IsNull(cmdUpdate) Then
            fraCompteH_Display
        End If
    Case "TIC"
        mYSSITICH_Update = "Update"
        newYSSITICH = oldYSSITICH
        If oldYSSITICH.SSITICYFCT <> "VU " Then
            newYSSITICH.SSITICSTAK = " "
            newYSSITICH.SSITICYFCT = "VU "
            newYSSITICH.SSITICYUSR = usrName_UCase
            newYSSITICH.SSITICYAMJ = DSys
            newYSSITICH.SSITICYHMS = time_Hms
        End If
        Call cmdSSIJRN_TIC("<X:contrôle des modifications (Historique)>")
        If IsNull(cmdUpdate) Then
            fraCompteH_Display
        End If
End Select
    

Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdParam_K2_Add_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
Call cmdUpdate_Init
mYSSITXT0_Update = "NewP"
newYSSITXT0 = paramYSSITXT0
newYSSITXT0.SSITXTTLNK = Val(txtParam_K2_Code)
newYSSITXT0.SSITXTINFO = Trim(txtParam_K2_Info)

If IsNull(cmdUpdate) Then
    fraParam_K2.Visible = False
    Call lstParam_K1_Load(lstParam_K1.Text)
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_K2_Delete_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
Call cmdUpdate_Init
mYSSITXT0_Update = "Delete"
newYSSITXT0 = paramYSSITXT0
newYSSITXT0.SSITXTTLNK = Val(txtParam_K2_Code)

If newYSSITXT0.SSITXTTLNK <> paramYSSITXT0.SSITXTTLNK Then
    Call MsgBox("Le code du paramétre sélectionné a été modifié", vbCritical, "Paramétrage : suppression d'un paramètre")
Else
    If IsNull(cmdUpdate) Then
        fraParam_K2.Visible = False
        Call lstParam_K1_Load(lstParam_K1.Text)
    End If
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdParam_K2_Quit_Click()
fraParam_K2.Visible = False
End Sub

Private Sub cmdParam_SSIMELUIDX_New_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
Call cmdUpdate_Init
rsYSSIMEL0_Init newYSSIMEL0
newYSSIMEL0.SSIMELNAT = "@"
newYSSIMEL0.SSIMELUIDX = Trim(txtParam_SSIMELUIDX)
newYSSIMEL0.SSIMELUNOM = Trim(txtParam_SSIMELUNOM)
newYSSIMEL0.SSIMELINFO = Trim(txtParam_SSIMELINFO)
newYSSIMEL0.SSIMELYFCT = "CRE"
newYSSIMEL0.SSIMELYUSR = usrName_UCase
newYSSIMEL0.SSIMELYAMJ = DSys
newYSSIMEL0.SSIMELYHMS = time_Hms
Call cmdUpdate_Init: mYSSIMEL0_Update = "New"
Call cmdSSIJRN_MEL("<X:" & txtParam_SSIMELINFO & ">")
Call cmdUpdate
fraParam_SSIMELUIDX.Visible = False
lstParam_SSIMELNAT_Click
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_SSIMELUIDX_Quit_Click()
fraParam_SSIMELUIDX.Visible = False
End Sub

Private Sub cmdParam_SSIMELUIDX_Update_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
Call cmdUpdate_Init
newYSSIMEL0 = oldYSSIMEL0
newYSSIMEL0.SSIMELUNOM = Trim(txtParam_SSIMELUNOM)
newYSSIMEL0.SSIMELINFO = Trim(txtParam_SSIMELINFO)
newYSSIMEL0.SSIMELYFCT = "MOD"
newYSSIMEL0.SSIMELYUSR = usrName_UCase
newYSSIMEL0.SSIMELYAMJ = DSys
newYSSIMEL0.SSIMELYHMS = time_Hms
Call cmdUpdate_Init: mYSSIMEL0_Update = "Update+H"
Call cmdSSIJRN_MEL("<X:" & txtParam_SSIMELINFO & ">")
Call cmdUpdate
fraParam_SSIMELUIDX.Visible = False
lstParam_SSIMELNAT_Click
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdPrint_Click()
Dim X As String, I As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass

Select Case SSTab1.Tab
    Case 0:
        mnuPrint_RTF.Visible = txtRTF.Visible
        If fraDetail.Visible And Not fraYSSIDOM0.Visible Then
            mnuPrint_RTF_USR.Visible = True
        Else
            mnuPrint_RTF_USR.Visible = False
        End If
        
        Me.PopupMenu mnuPrint, vbPopupMenuLeftButton
    End Select

Me.Enabled = True: Me.MousePointer = 0



End Sub

Private Sub cmdProfil_Change_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
oldYSSIDIV0 = rtfYSSIDIV0
txtProfil_DIDK_DIV.Enabled = False
txtProfil_UIDD_DIV.Enabled = False
txtProfil_IDX_DIV.Enabled = False

txtProfil_DIDK_DIV = oldYSSIDIV0.SSIDIVDIDK
txtProfil_UIDD_DIV = oldYSSIDIV0.SSIDIVUIDD
txtProfil_IDX_DIV = oldYSSIDIV0.SSIDIVUIDX
txtProfil_PRFX_DIV = oldYSSIDIV0.SSIDIVPRFX
txtProfil_UNOM_DIV = oldYSSIDIV0.SSIDIVUNOM
txtProfil_Info_DIV = oldYSSIDIV0.SSIDIVINFO
cmdProfil_Change.Visible = False
cmdProfil_Update.Visible = True
fraProfil_Update_DIV.Caption = "Modification d'un profil"
fraProfil_Update_DIV.Visible = True
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdProfil_Delete_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass

Select Case cmdSelect_SQL_K
    Case "1", "2":
            If Trim(oldYSSIDOM0.SSIDOMUIDX) = "" Then
                mYSSIDOM0_Update = "Delete+H"
                
                newYSSIDOM0 = oldYSSIDOM0
                newYSSIDOM0.SSIDOMYFCT = "SUP"
                newYSSIDOM0.SSIDOMYVER = newYSSIDOM0.SSIDOMYVER + 1
                Call cmdSSIJRN_DOM("<X:suppession du profil " & oldYSSIDOM0.SSIDOMPRFX & ">")

                If IsNull(cmdUpdate) Then
                    If cmdSelect_SQL_K = "1" Then fraDetail_Control_SSIUSRPRFK
                    Call cmdProfil_Quit_Click
                    Call fraDetail_Load
                End If
            Else
                newYSSIDOM0 = oldYSSIDOM0
                newYSSIDOM0.SSIDOMYFCT = "SUP"
                newYSSIDOM0.SSIDOMYVER = newYSSIDOM0.SSIDOMYVER + 1
                newYSSIDOM0.SSIDOMYAMJ = DSys
                newYSSIDOM0.SSIDOMYHMS = time_Hms
                newYSSIDOM0.SSIDOMYUSR = usrName_UCase
                mYSSIDOM0_Update = "Delete+H"  '"Update+H" '
                
                If oldYSSIDOM0.SSIDOMNAT = " " Then
                     Select Case Trim(oldYSSIDOM0.SSIDOMDIDX)
                         Case "IBM"
                             oldYSSIIBM0 = usrYSSIIBM0
                             newYSSIIBM0 = usrYSSIIBM0
                             newYSSIIBM0.SSIIBMPRFK = "?"
                             mYSSIIBM0_Update = "Update"
                         Case "SAA"
                             oldYSSISAA0 = usrYSSISAA0
                             newYSSISAA0 = usrYSSISAA0
                             newYSSISAA0.SSISAAPRFK = "?"
                             mYSSISAA0_Update = "Update"
                         Case "SAB"
                             oldYSSISAB0 = usrYSSISAB0
                             newYSSISAB0 = usrYSSISAB0
                             newYSSISAB0.SSISABPRFK = "?"
                             mYSSISAB0_Update = "Update"
                          Case "WIN"
                             oldYSSIWIN0 = usrYSSIWIN0
                             newYSSIWIN0 = usrYSSIWIN0
                             newYSSIWIN0.SSIWINPRFK = "?"
                             mYSSIWIN0_Update = "Update"
                            ''Call cmdSSIJRN_WIN("<X:Modification 'User'>")
                          Case "DIV"
                             oldYSSIDIV0 = usrYSSIDIV0
                             newYSSIDIV0 = usrYSSIDIV0
                             newYSSIDIV0.SSIDIVPRFK = "?"
                             mYSSIDIV0_Update = "Update"
                          Case "TIC"
                             oldYSSITIC0 = usrYSSITIC0
                             newYSSITIC0 = usrYSSITIC0
                             newYSSITIC0.SSITICPRFK = "?"
                             mYSSITIC0_Update = "Update"
                    End Select
                End If
                Call cmdSSIJRN_DOM("<X:suppession du Profil/compte " & oldYSSIDOM0.SSIDOMPRFX & " - " & oldYSSIDOM0.SSIDOMUIDX & ">")

                If IsNull(cmdUpdate) Then
                    If cmdSelect_SQL_K = "1" Then fraDetail_Control_SSIUSRPRFK
                    Call cmdProfil_Quit_Click
                    Call fraDetail_Load
                End If
            End If
            
            
    Case "2":
End Select

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdProfil_Display_Click()
Dim X As String, wUIDD As Long
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> cmdProfil_Display ........"): DoEvents

oldYSSIIBM0.SSIIBMNAT = " "
oldYSSIIBM0.SSIIBMUIDD = Val(txtProfil_UIDD)
 
If IsNull(cmdSSIIBM_Detail_Display("", "")) Then
    oldYSSIIBM0 = xYSSIIBM0
    cmdProfil_Update.Visible = arrHab(3)
Else
    cmdProfil_Update.Visible = False
End If


Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub
Private Function cmdSSISAA_Detail_Load(lFct As String)
Dim X As String, blnParam As Boolean, whighlight As Integer
On Error GoTo Exit_sub
cmdSSISAA_Detail_Load = "?"
whighlight = 11


'_____________________________________________________________________________________________
If Trim(usrYSSISAA0.SSISAAUIDX) <> "" Then

    If lFct = "YSSISAA*" Then
        blnParam = True
        X = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSISAAH " _
          & " where SSISAANAT = '" & usrYSSISAA0.SSISAANAT & " ' and SSISAAUIDX = '" & usrYSSISAA0.SSISAAUIDX & "'" _
          & " and SSISAAUSEQ = " & usrYSSISAA0.SSISAAUSEQ & "  and SSISAAYVER = " & usrYSSISAA0.SSISAAYVER
        Set rsSab = cnsab.Execute(X)
        If rsSab(0) > 0 Then
            lFct = "YSSISAAH"
        Else
            lFct = "SAA"
        End If
    End If
    If lFct = "YSSISAAH" Then
        whighlight = 9
       X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAAH " _
             & " where SSISAANAT = '" & usrYSSISAA0.SSISAANAT & " ' and SSISAAUIDX = '" & usrYSSISAA0.SSISAAUIDX & "'" _
         & " and SSISAAUSEQ = " & usrYSSISAA0.SSISAAUSEQ & "  and SSISAAYVER = " & usrYSSISAA0.SSISAAYVER
    Else
        If lFct = "SSISAAUIDD" Then
           X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
                 & " where SSISAANAT = '" & usrYSSISAA0.SSISAANAT & " ' and SSISAAUIDX = '" & usrYSSISAA0.SSISAAUIDX & "'" _
             & " and SSISAAUIDD = " & usrYSSISAA0.SSISAAUIDD & " order by SSISAAUSEQ desc"
        Else
           X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
                 & " where SSISAANAT = '" & usrYSSISAA0.SSISAANAT & " ' and SSISAAUIDX = '" & usrYSSISAA0.SSISAAUIDX & "'" _
             & " and SSISAAUSEQ = " & usrYSSISAA0.SSISAAUSEQ
        End If
    End If

    Set rsSab = cnsab.Execute(X)
      
    If rsSab.EOF Then
        Call MsgBox(usrYSSISAA0.SSISAAUIDX & " : Id utilisateur inconnue dans le domaine " & mSSIDOMDIDX, vbCritical, "cmdSSISAA_Detail_Load")
        GoTo Exit_sub
    End If
    
    Call rsYSSISAA0_GetBuffer(rsSab, usrYSSISAA0)
    
    
End If

If blnParam Then
    Select Case usrYSSISAA0.SSISAANAT
        Case "U": mRTF = cmdSSISAA_Detail_txtRTF_Param("Unit ", whighlight)
        Case "A": mRTF = cmdSSISAA_Detail_txtRTF_Param("Application ", whighlight)
        Case "F": mRTF = cmdSSISAA_Detail_txtRTF_Param("Fonction ", whighlight)
        Case "P": mRTF = cmdSSISAA_Detail_txtRTF_Param("Paramétrage du profil", whighlight)
        Case Else: mRTF = cmdSSISAA_Detail_txtRTF_Param(usrYSSISAA0.SSISAANAT, whighlight)
    End Select
    cmdSSISAA_Detail_Load = Null
'_____________________________________________________________________________________________
Else
    If lFct = "SAA" Then
        prfYSSISAA0.SSISAANAT = "$"
        prfYSSISAA0.SSISAAUIDX = Trim(usrYSSISAA0.SSISAAPRFX)
        prfYSSISAA0.SSISAAUSEQ = 0
    End If
    
    If lFct = "YSSISAAH" Then
        oldYSSIDOM0.SSIDOMPRFX = usrYSSISAA0.SSISAAPRFX
    Else
        If Trim(prfYSSISAA0.SSISAAUIDX) = "" Then
            prfYSSISAA0 = usrYSSISAA0
        Else
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
                 & " where SSISAANAT = '" & prfYSSISAA0.SSISAANAT & " ' and SSISAAUIDX = '" & prfYSSISAA0.SSISAAUIDX & "'" _
                 & " and SSISAAUSEQ = " & prfYSSISAA0.SSISAAUSEQ
            Set rsSab = cnsab.Execute(X)
            
            If rsSab.EOF Then
                Call MsgBox(oldYSSISAA0.SSISAAUIDX & " : Profil inconnu dans le domaine " & mSSIDOMDIDX, vbCritical, "cmdSSISAA_Detail_Load")
                GoTo Exit_sub
            End If
            
            Call rsYSSISAA0_GetBuffer(rsSab, prfYSSISAA0)
        End If
    End If
        '_____________________________________________________________________________________________
        
        
        mRTF = cmdSSISAA_Detail_txtRTF(lFct, whighlight)
        cmdSSISAA_Detail_Load = Null
End If


Exit_sub:

End Function
Private Function cmdSSISAA_Detail_Display(lFct As String)

On Error GoTo Exit_sub
cmdSSISAA_Detail_Display = "?"

If IsNull(cmdSSISAA_Detail_Load(lFct)) Then

    txtRTF.TextRTF = VB_RTF_Modèle
    
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", mRTF)
    
    Call txtRTF_Visible
    
    cmdSSISAA_Detail_Display = Null
End If
Exit_sub:


End Function

Private Function cmdSSIWIN_Detail_Display(lFct As String)
Dim whighlight As Integer
On Error GoTo Exit_sub
cmdSSIWIN_Detail_Display = "?"
mRTF = ""
'If IsNull(cmdSSIWIN_Detail_Load(lFct)) Then
Select Case lFct
    
    Case "YSSIWIN0_UIDD"
        whighlight = 11
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
             & " where SSIWINNAT = '" & rtfYSSIWIN0.SSIWINNAT & "' and SSIWINUIDD = '" & rtfYSSIWIN0.SSIWINUIDD & "'"
        Set rsSab_X = cnsab.Execute(X)
    
    Case "YSSIWINH"
        whighlight = 9
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIWINH " _
             & " where SSIWINNAT = '" & rtfYSSIWIN0.SSIWINNAT & "' and SSIWINGUID = '" & rtfYSSIWIN0.SSIWINGUID & "'" _
              & " and SSIWINYVER = " & rtfYSSIWIN0.SSIWINYVER
       Set rsSab_X = cnsab.Execute(X)
    
    Case "YSSIWIN*"
        whighlight = 9
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIWINH " _
             & " where SSIWINNAT = '" & rtfYSSIWIN0.SSIWINNAT & "' and SSIWINUIDX = '" & rtfYSSIWIN0.SSIWINUIDX & "'" _
             & " and SSIWINYVER = " & rtfYSSIWIN0.SSIWINYVER
        Set rsSab_X = cnsab.Execute(X)
        If rsSab_X.EOF Then
            whighlight = 11
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
                 & " where SSIWINNAT = '" & rtfYSSIWIN0.SSIWINNAT & "' and SSIWINUIDX = '" & rtfYSSIWIN0.SSIWINUIDX & "'" _
                 & " and SSIWINYVER = " & rtfYSSIWIN0.SSIWINYVER
            Set rsSab_X = cnsab.Execute(X)
        End If
    Case "YSSIWIN0_UIDX"
            whighlight = 11
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
                 & " where SSIWINNAT = '" & rtfYSSIWIN0.SSIWINNAT & "' and SSIWINUIDX = '" & rtfYSSIWIN0.SSIWINUIDX & "'"
            Set rsSab_X = cnsab.Execute(X)
    Case Else
        whighlight = 11
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
             & " where SSIWINNAT = '" & rtfYSSIWIN0.SSIWINNAT & "' and SSIWINGUID = '" & rtfYSSIWIN0.SSIWINGUID & "'"
        Set rsSab_X = cnsab.Execute(X)
End Select
If Not rsSab_X.EOF Then
    Call rsYSSIWIN0_GetBuffer(rsSab_X, rtfYSSIWIN0)
    mRTF = mRTF & cmdSSIWIN_Detail_txtRTF(whighlight)
Else
    mRTF = "\highlight12 Erreur de lecture : \highlight0\par " & X
End If

txtRTF.TextRTF = VB_RTF_Modèle

txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", mRTF)

Call txtRTF_Visible

cmdSSIWIN_Detail_Display = Null
'End If
Exit_sub:


End Function

Private Function cmdSSIMEL_Detail_Display(lFct As String)
Dim whighlight As Integer
On Error GoTo Exit_sub
cmdSSIMEL_Detail_Display = "?"
mRTF = ""
'If IsNull(cmdSSIMEL_Detail_Load(lFct)) Then
Select Case lFct
    
    Case "YSSIMELH"
        whighlight = 9
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIMELH " _
             & " where SSIMELNAT = '" & rtfYSSIMEL0.SSIMELNAT & "' and SSIMELUIDX = '" & rtfYSSIMEL0.SSIMELUIDX & "'" _
              & " and SSIMELYVER = " & rtfYSSIMEL0.SSIMELYVER
       Set rsSab_X = cnsab.Execute(X)
    
    Case "YSSIMEL*"
        whighlight = 9
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIMELH " _
             & " where SSIMELNAT = '" & rtfYSSIMEL0.SSIMELNAT & "' and SSIMELUIDX = '" & rtfYSSIMEL0.SSIMELUIDX & "'" _
             & " and SSIMELYVER = " & rtfYSSIMEL0.SSIMELYVER
        Set rsSab_X = cnsab.Execute(X)
        If rsSab_X.EOF Then
            whighlight = 11
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0 " _
                 & " where SSIMELNAT = '" & rtfYSSIMEL0.SSIMELNAT & "' and SSIMELUIDX = '" & rtfYSSIMEL0.SSIMELUIDX & "'" _
                 & " and SSIMELYVER = " & rtfYSSIMEL0.SSIMELYVER
            Set rsSab_X = cnsab.Execute(X)
        End If
    Case "YSSIMEL0_UIDX", "YSSIMEL0"
            whighlight = 11
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0 " _
                 & " where SSIMELNAT = '" & rtfYSSIMEL0.SSIMELNAT & "' and SSIMELUIDX = '" & rtfYSSIMEL0.SSIMELUIDX & "'"
            Set rsSab_X = cnsab.Execute(X)
End Select
If Not rsSab_X.EOF Then
    Call rsYSSIMEL0_GetBuffer(rsSab_X, rtfYSSIMEL0)
    mRTF = mRTF & cmdSSIMEL_Detail_txtRTF(whighlight)
Else
    mRTF = "\highlight12 Erreur de lecture : \highlight0\par " & X
End If

txtRTF.TextRTF = VB_RTF_Modèle

txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", mRTF)

Call txtRTF_Visible

cmdSSIMEL_Detail_Display = Null
'End If
Exit_sub:


End Function

Private Function cmdSSITIC_Detail_Display(lFct As String)
Dim whighlight As Integer
On Error GoTo Exit_sub
cmdSSITIC_Detail_Display = "?"
mRTF = ""
'If IsNull(cmdSSITIC_Detail_Load(lFct)) Then
Select Case lFct
    
    Case "YSSITICH"
        whighlight = 9
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSITICH " _
             & " where SSITICNAT = '" & rtfYSSITIC0.SSITICNAT & "' and SSITICUIDX = '" & rtfYSSITIC0.SSITICUIDX & "'" _
              & " and SSITICYVER = " & rtfYSSITIC0.SSITICYVER
       Set rsSab_X = cnsab.Execute(X)
    
    Case "YSSITIC*"
        whighlight = 9
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSITICH " _
             & " where SSITICNAT = '" & rtfYSSITIC0.SSITICNAT & "' and SSITICUIDX = '" & rtfYSSITIC0.SSITICUIDX & "'" _
             & " and SSITICYVER = " & rtfYSSITIC0.SSITICYVER
        Set rsSab_X = cnsab.Execute(X)
        If rsSab_X.EOF Then
            whighlight = 11
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSITIC0 " _
                 & " where SSITICNAT = '" & rtfYSSITIC0.SSITICNAT & "' and SSITICUIDX = '" & rtfYSSITIC0.SSITICUIDX & "'" _
                 & " and SSITICYVER = " & rtfYSSITIC0.SSITICYVER
            Set rsSab_X = cnsab.Execute(X)
        End If
    Case "YSSITIC0_UIDX", "YSSITIC0"
            whighlight = 11
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSITIC0 " _
                 & " where SSITICNAT = '" & rtfYSSITIC0.SSITICNAT & "' and SSITICUIDX = '" & rtfYSSITIC0.SSITICUIDX & "'"
            Set rsSab_X = cnsab.Execute(X)
End Select
If Not rsSab_X.EOF Then
    Call rsYSSITIC0_GetBuffer(rsSab_X, rtfYSSITIC0)
    mRTF = mRTF & cmdSSITIC_Detail_txtRTF(whighlight)
Else
    mRTF = "\highlight12 Erreur de lecture : \highlight0\par " & X
End If

txtRTF.TextRTF = VB_RTF_Modèle

txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", mRTF)

Call txtRTF_Visible

cmdSSITIC_Detail_Display = Null
'End If
Exit_sub:


End Function


Private Function cmdSSISAM_Detail_Display(lFct As String)
Dim whighlight As Integer
On Error GoTo Exit_sub
cmdSSISAM_Detail_Display = "?"
mRTF = ""
'If IsNull(cmdSSISAM_Detail_Load(lFct)) Then
Select Case lFct
    Case "YSSISAM0"
            whighlight = 11
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAM0 " _
                 & " where SSISAMUID = '" & rtfYSSISAM0.SSISAMUIDD & "'"
            Set rsSab_X = cnsab.Execute(X)
    
    Case "YSSISAMH"
        whighlight = 9
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAMH " _
             & " where SSISAMUIDD = '" & rtfYSSISAM0.SSISAMUIDD & "'" _
              & " and SSISAMYVER = " & rtfYSSISAM0.SSISAMYVER
       Set rsSab_X = cnsab.Execute(X)
    
    Case "YSSISAM*"
        whighlight = 9
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAMH " _
             & " where SSISAMUIDd = '" & rtfYSSISAM0.SSISAMUIDD & "'" _
             & " and SSISAMYVER = " & rtfYSSISAM0.SSISAMYVER
        Set rsSab_X = cnsab.Execute(X)
        If rsSab_X.EOF Then
            whighlight = 11
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAM0 " _
                 & " where SSISAMUIDd = '" & rtfYSSISAM0.SSISAMUIDD & "'" _
                 & " and SSISAMYVER = " & rtfYSSISAM0.SSISAMYVER
            Set rsSab_X = cnsab.Execute(X)
        End If
End Select
If Not rsSab_X.EOF Then
    Call rsYSSISAM0_GetBuffer(rsSab_X, rtfYSSISAM0)
    mRTF = mRTF & cmdSSISAM_Detail_txtRTF(whighlight)
Else
    mRTF = "\highlight12 Erreur de lecture : \highlight0\par " & X
End If

txtRTF.TextRTF = VB_RTF_Modèle

txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", mRTF)

Call txtRTF_Visible

cmdSSISAM_Detail_Display = Null
'End If
Exit_sub:


End Function

Private Function cmdSSISAW_Detail_Display(lFct As String)
Dim whighlight As Integer
On Error GoTo Exit_sub
cmdSSISAW_Detail_Display = "?"
mRTF = ""
'If IsNull(cmdSSISAB_Detail_Load(lFct)) Then
Select Case lFct
    Case "YSSISAB0"
            whighlight = 11
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
                 & " where SSISABNAT = 'W' and SSISABUIDX = '" & usrYSSISAB0.SSISABUIDX & "' and SSISABULOT = 0"
            Set rsSab_X = cnsab.Execute(X)
    
    Case "YSSISABH"
        whighlight = 9
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSISABH " _
             & " where SSISABNAT = 'W' and SSISABUIDX = '" & usrYSSISAB0.SSISABUIDX & "' and SSISABULOT = 0" _
              & " and SSISABYVER = " & usrYSSISAB0.SSISABYVER
       Set rsSab_X = cnsab.Execute(X)
    
    Case "YSSISAB*"
        whighlight = 9
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSISABH " _
             & " where SSISABNAT = 'W' and SSISABUIDX = '" & usrYSSISAB0.SSISABUIDX & "' and SSISABULOT = 0" _
             & " and SSISABYVER = " & usrYSSISAB0.SSISABYVER
        Set rsSab_X = cnsab.Execute(X)
        If rsSab_X.EOF Then
            whighlight = 11
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
                 & " where SSISABNAT = 'W' and SSISABUIDX = '" & usrYSSISAB0.SSISABUIDX & "' and SSISABULOT = 0" _
                 & " and SSISABYVER = " & usrYSSISAB0.SSISABYVER
            Set rsSab_X = cnsab.Execute(X)
        End If
End Select
If Not rsSab_X.EOF Then
    Call rsYSSISAB0_GetBuffer(rsSab_X, usrYSSISAB0)
    mRTF = mRTF & cmdSSISAW_Detail_txtRTF(whighlight)
Else
    mRTF = "\highlight12 Erreur de lecture : \highlight0\par " & X
End If

txtRTF.TextRTF = VB_RTF_Modèle

txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", mRTF)

Call txtRTF_Visible

cmdSSISAW_Detail_Display = Null
'End If
Exit_sub:


End Function



Private Function cmdSSIDIV_Detail_Display(lFct As String)
Dim whighlight As Integer
On Error GoTo Exit_sub
cmdSSIDIV_Detail_Display = "?"
mRTF = ""

Select Case lFct
    
    'Case "YSSIDIV0_UIDD"
     '   whighlight = 11
     '   X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
     '        & " where SSIDIVNAT = '" & rtfYSSIDIV0.SSIDIVNAT & "' and SSIDIVUIDD = '" & rtfYSSIDIV0.SSIDIVUIDD & "'"
     '   Set rsSab_X = cnsab.Execute(X)
    
    Case "YSSIDIVH"
        whighlight = 9
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIVH " _
             & " where SSIDIVNAT = '" & rtfYSSIDIV0.SSIDIVNAT & "' and SSIDIVUIDX = '" & rtfYSSIDIV0.SSIDIVUIDX & "'" _
             & " and SSIDIVUIDD = " & rtfYSSIDIV0.SSIDIVUIDD & " and SSIDIVYVER = " & rtfYSSIDIV0.SSIDIVYVER
        Set rsSab_X = cnsab.Execute(X)
    
    Case "YSSIDIV*"
        whighlight = 9
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIVH " _
             & " where SSIDIVNAT = '" & rtfYSSIDIV0.SSIDIVNAT & "' and SSIDIVUIDX = '" & rtfYSSIDIV0.SSIDIVUIDX & "'" _
             & " and SSIDIVUIDD = " & rtfYSSIDIV0.SSIDIVUIDD _
             & " and SSIDIVYVER = " & rtfYSSIDIV0.SSIDIVYVER
        Set rsSab_X = cnsab.Execute(X)
        If rsSab_X.EOF Then
            whighlight = 11
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
                 & " where SSIDIVNAT = '" & rtfYSSIDIV0.SSIDIVNAT & "' and SSIDIVUIDX = '" & rtfYSSIDIV0.SSIDIVUIDX & "'" _
                 & " and SSIDIVUIDD = " & rtfYSSIDIV0.SSIDIVUIDD _
                 & " and SSIDIVYVER = " & rtfYSSIDIV0.SSIDIVYVER
            Set rsSab_X = cnsab.Execute(X)
        End If
    Case Else
            whighlight = 11
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
                 & " where SSIDIVNAT = '" & rtfYSSIDIV0.SSIDIVNAT & "' and SSIDIVUIDX = '" & rtfYSSIDIV0.SSIDIVUIDX & "'" _
                 & " and SSIDIVUIDD = " & rtfYSSIDIV0.SSIDIVUIDD

            Set rsSab_X = cnsab.Execute(X)
End Select
If Not rsSab_X.EOF Then
    Call rsYSSIDIV0_GetBuffer(rsSab_X, rtfYSSIDIV0)
    mRTF = mRTF & cmdSSIDIV_Detail_txtRTF(whighlight)
Else
    If rtfYSSIDIV0.SSIDIVUIDX = "" Then
        mRTF = "\highlight12 identifiant absent \highlight0\par "
    Else
        mRTF = "\highlight12 Erreur de lecture : \highlight0\par " & X
    End If
    Call rsYSSIDIV0_Init(rtfYSSIDIV0)
End If

txtRTF.TextRTF = VB_RTF_Modèle

txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", mRTF)

Call txtRTF_Visible

cmdSSIDIV_Detail_Display = Null
'End If
Exit_sub:


End Function


Private Function cmdSSIIBM_Detail_Display(lUPUPRF As String, lUSR As String)
Dim X As String

On Error GoTo Exit_sub
cmdSSIIBM_Detail_Display = "?"

If IsNull(cmdSSIIBM_Detail_Load(lUPUPRF, lUSR)) Then


    txtRTF.TextRTF = VB_RTF_Modèle
    
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", mRTF & X)
    
    Call txtRTF_Visible
    
    cmdSSIIBM_Detail_Display = Null
End If
Exit_sub:


End Function

Private Function cmdSSIIBM_Detail_Load(lUPUPRF As String, lUSR As String)
Dim X As String, blnSQL As Boolean, whighlight As Integer
On Error GoTo Exit_sub
cmdSSIIBM_Detail_Load = "?"
blnSQL = True
whighlight = 11
'_____________________________________________________________________________________________
If lUPUPRF = "YSSIIBM*" Then
    X = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSIIBMH " _
      & " where SSIIBMNAT = '" & xYSSIIBM0.SSIIBMNAT & " ' and SSIIBMUIDD = " & xYSSIIBM0.SSIIBMUIDD _
      & " and SSIIBMYVER = " & xYSSIIBM0.SSIIBMYVER
    Set rsSab = cnsab.Execute(X)
    If rsSab(0) > 0 Then
        lUPUPRF = "YSSIIBMH"
    Else
        lUPUPRF = ""
        oldYSSIIBM0 = xYSSIIBM0
    End If
End If

Select Case lUPUPRF
    Case "UPUPRF"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0 " _
             & " where SSIIBMNAT = '" & oldYSSIIBM0.SSIIBMNAT & " ' and UPUPRF = '" & oldYSSIIBM0.UPUPRF & "'"

    Case ""
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0 " _
             & " where SSIIBMNAT = '" & oldYSSIIBM0.SSIIBMNAT & " ' and SSIIBMUIDD = " & oldYSSIIBM0.SSIIBMUIDD

    Case "YSSIIBMH"
        whighlight = 9
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBMH " _
             & " where SSIIBMNAT = '" & xYSSIIBM0.SSIIBMNAT & " ' and SSIIBMUIDD = " & xYSSIIBM0.SSIIBMUIDD _
             & "  and SSIIBMYVER = " & xYSSIIBM0.SSIIBMYVER
             
    Case Else
            Call rsYSSIIBM0_Init(xYSSIIBM0)
            blnSQL = False
End Select

If blnSQL Then
        Set rsSab = cnsab.Execute(X)
        
        If rsSab.EOF Then
            Call MsgBox(oldYSSIIBM0.SSIIBMUIDD & " : Id utilisateur inconnue dans le domaine " & mSSIDOMDIDX, vbCritical, "cmdProfil_Display_Click")
            GoTo Exit_sub
        End If
    
    Call rsYSSIIBM0_GetBuffer(rsSab, xYSSIIBM0)
End If
'_____________________________________________________________________________________________
If lUSR = "USR" Then

    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0 " _
         & " where SSIIBMNAT = '" & usrYSSIIBM0.SSIIBMNAT & " ' and SSIIBMUIDD = " & usrYSSIIBM0.SSIIBMUIDD
    
    Set rsSab = cnsab.Execute(X)
      
    If rsSab.EOF Then
        Call MsgBox(usrYSSIIBM0.SSIIBMUIDD & " : Id utilisateur inconnue dans le domaine " & mSSIDOMDIDX, vbCritical, "cmdProfil_Display_Click")
        GoTo Exit_sub
    End If

Call rsYSSIIBM0_GetBuffer(rsSab, usrYSSIIBM0)


End If
'_____________________________________________________________________________________________

If xYSSIIBM0.SSIIBMPRFK = "S" Then whighlight = 14
mRTF = cmdSSIIBM_Detail_txtRTF(lUSR, whighlight)



cmdSSIIBM_Detail_Load = Null

Exit_sub:

End Function



Private Sub cmdProfil_Excel_Click()
Dim X As String, wFile As String, K As Integer
On Error GoTo Error_Handler

Me.Enabled = False: Me.MousePointer = vbHourglass
currentAction = "cmdProfil_Excel_Click"
Call lstErr_Clear(lstErr, cmdContext, "> cmdProfil_Excel_Click ........"): DoEvents

If cmdProfil_Excel.BackColor = &H80FF80 Then
    If lstW.ListCount = 0 Then
        Call MsgBox("Préciser au moins un profil", vbCritical, currentAction)
        GoTo Exit_sub
    Else
        lstW.Visible = False
        cmdProfil_Excel.BackColor = &H80FFFF
    End If
Else
    X = MsgBox("Voulez-vous sélectionner certains profils (OUI)" & vbCrLf & " ou inclure tous les profils (NON) ? ", vbYesNoCancel, "Excel : sélection des profils")
    Select Case X
        Case vbYes: blnProfil_Excel_All = False: cmdProfil_Excel_Select: arrProfil_Nb = 0: GoTo Exit_sub
        Case vbNo: blnProfil_Excel_All = True
        Case vbCancel: GoTo Exit_sub
    End Select
End If


wFile = "C:\Temp\BIA_SSI " & mSSIDOMDIDX & " " & DSYS_Time
'______________________________________________
    wFile = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "Nom du fichier d'exportation ", wFile)
If Trim(wFile) = "" Then GoTo Exit_sub

If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile


Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "BIA_SSI " & mSSIDOMDIDX
    .Subject = ""
End With
    
Select Case mSSIDOMDIDX
    Case "SAB": Call cmdProfil_Excel_SAB
End Select

wbExcel.SaveAs wFile
wbExcel.Close
appExcel.Quit
cmdProfil_Excel.BackColor = &H80FFFF
'===================================================================================================
Exit_sub:
'__________________________________________________________________________________

Set rsSab = Nothing

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
'__________________________________________________________________________________
Call lstErr_AddItem(lstErr, frmElp.cmdContext, "< Exportation terminée"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
 
    Call lstErr_AddItem(lstErr, cmdContext, "< cmdProfil_Excel_Click  terminé"): DoEvents
    Me.Enabled = True: Me.MousePointer = 0
cmdProfil_Excel.BackColor = &H80FFFF

End Sub

Private Sub cmdProfil_Histo_Click()
Dim xSQL As String, X As String, blnHisto_Init As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> cmdSSIDOM_Histo ........"): DoEvents
On Error GoTo Error_Handler

Call fraCompteH_Display

currentAction = "cmdProfil_Histo_Click"
txtRTF.TextRTF = VB_RTF_Modèle

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOMH " _
     & " left outer join " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
     & " on SSITXTNAT = SSIDOMNAT and SSITXTUIDN = SSIDOMUIDN and SSITXTDIDX = SSIDOMDIDX" _
     & " and SSITXTUIDX = SSIDOMUIDX and SSITXTTLNK = SSIDOMTLNK" _
     & " where SSIDOMNAT = '" & oldYSSIDOM0.SSIDOMNAT & "' and SSIDOMUIDN = " & oldYSSIDOM0.SSIDOMUIDN _
     & " and SSIDOMDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "' and SSIDOMUIDX = '" & oldYSSIDOM0.SSIDOMUIDX & "'" _
     & " and SSIDOMUIDD = " & oldYSSIDOM0.SSIDOMUIDD & "  order by SSIDOMYAMJ, SSIDOMYHMS, SSIDOMYVER"
     
Set rsSab = cnsab.Execute(xSQL)
  
Do While Not rsSab.EOF
    Call rsYSSIDOM0_GetBuffer(rsSab, xYSSIDOM0)
    If xYSSIDOM0.SSIDOMTLNK = 0 Then
        xYSSITXT0.SSITXTINFO = ""
    Else
        If Not IsNull(rsSab("SSITXTINFO")) Then
            xYSSITXT0.SSITXTINFO = rsSab("SSITXTINFO")
        Else
            xYSSITXT0.SSITXTINFO = ""
        End If
    End If
    
    If Not blnHisto_Init Then
        blnHisto_Init = True
        hYSSIDOM0 = xYSSIDOM0
        hYSSITXT0 = xYSSITXT0
    End If
    
    
    X = X & cmdSSIDOM_Histo_txtRTF
    
    hYSSIDOM0 = xYSSIDOM0
    hYSSITXT0 = xYSSITXT0
    rsSab.MoveNext
Loop

Set rsSab = Nothing
xYSSIDOM0 = oldYSSIDOM0
xYSSITXT0 = oldYSSITXT0_DOM
X = X & cmdSSIDOM_Histo_txtRTF

txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", X)
Call txtRTF_Visible
GoTo Exit_sub

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
Exit_sub:
 
    Call lstErr_AddItem(lstErr, cmdContext, "< cmdSSIDOM_Histo terminé"): DoEvents
    Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdProfil_New_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> cmdProfil_New ........"): DoEvents
cmdProfil_New.Visible = False
chkSSIDOMDECH.Value = "0"
fgProfil.Enabled = False
txtProfil_UIDD = ""
txtProfil_IDX = ""
txtRTF.Visible = False

 
 
Select Case mSSIDOMDIDX
    Case "IBM": fraProfil_Update.Visible = True
    Case "DIV":
        Call rsYSSIDIV0_Init(oldYSSIDIV0)
        fraProfil_Update_DIV.Caption = "Création d'un profil"
        txtProfil_DIDK_DIV.Enabled = True
        txtProfil_UIDD_DIV.Enabled = True
        txtProfil_IDX_DIV.Enabled = True
        fraProfil_Update_DIV.Visible = True
        cmdProfil_Update.Visible = True
End Select


Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdProfil_Print_Click()
Dim X As String, K As Integer, mRTF_All As String

Me.Enabled = False: Me.MousePointer = vbHourglass

X = "C:\Temp\BIA_SSI Profil " & DSYS_Time
'______________________________________________
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & X _
        & vbCrLf & "     =========================", "Nom du fichier d'exportation ", X)
    If Trim(X) <> "" Then
    
        For K = 1 To fgProfil.Rows - 1
            fgProfil.Row = K
            Select Case mSSIDOMDIDX
                Case "IBM"
                    fgProfil.Col = 0: oldYSSIIBM0.SSIIBMUIDD = Val(fgProfil.Text)
                    oldYSSIIBM0.SSIIBMNAT = "$"
                    Call cmdSSIIBM_Detail_Display("", "")
                Case "SAA"
                    usrYSSISAA0.SSISAAUIDX = ""
                    fgProfil.Col = 1: prfYSSISAA0.SSISAAUIDX = Trim(fgProfil.Text)
                    prfYSSISAA0.SSISAANAT = "$"
                    Call cmdSSISAA_Detail_Display("")
                Case "SAB"
                    usrYSSISAB0.SSISABUIDX = ""
                    fgProfil.Col = 1: prfYSSISAB0.SSISABUIDX = Trim(fgProfil.Text)
                    prfYSSISAB0.SSISABNAT = "$"
                    Call cmdSSISAB_Detail_Display("")
                Case "WIN"
                    usrYSSIWIN0.SSIWINUIDX = ""
                    fgProfil.Col = 3: rtfYSSIWIN0.SSIWINGUID = Trim(fgProfil.Text)
                    rtfYSSIWIN0.SSIWINNAT = "$"
                    Call cmdSSIWIN_Detail_Display("YSSIWIN0_GUID")
                Case "DIV"
                    usrYSSIDIV0.SSIDIVUIDX = ""
                    fgProfil.Col = 0: rtfYSSIDIV0.SSIDIVUIDD = Val(fgProfil.Text)
                    fgProfil.Col = 1: rtfYSSIDIV0.SSIDIVUIDX = Trim(fgProfil.Text)
                    rtfYSSIDIV0.SSIDIVNAT = "$"
                    Call cmdSSIDIV_Detail_Display("YSSIDIV0")
                Case "TIC"
                    usrYSSITIC0.SSITICUIDX = ""
                    fgProfil.Col = 1: rtfYSSITIC0.SSITICUIDX = Trim(fgProfil.Text)
                    rtfYSSITIC0.SSITICNAT = "$"
                    Call cmdSSITIC_Detail_Display("YSSITIC0")
            End Select
            mRTF_All = mRTF_All & mRTF
        Next K
     txtRTF.TextRTF = VB_RTF_Modèle
    
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", mRTF_All)
    
    Call txtRTF_Visible
     ''txtRTF.SaveFile X
     Call cmdPrint_Word_PDF(X)
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdProfil_Quit_Click()
On Error Resume Next
cboProfil_DOM.ListIndex = 0
txtRTF.Visible = False
fraProfil_Update.Visible = False
fraProfil_Update_DIV.Visible = False
fraYSSIDOM0.Visible = False
fraYSSIDIV0.Visible = False
fgProfil.Visible = False: chkProfil_DOM.Visible = False
fraProfil.Visible = False
cmdProfil_Update.Visible = False
cmdProfil_Update_DIV.Visible = False
cmdProfil_New.Visible = False
cmdProfil_Change.Visible = False
cmdProfil_Print.Visible = False: cmdProfil_Excel.Visible = False
cmdProfil_Delete.Visible = False
cmdCompte_Val.Visible = False
lstW.Visible = False
cmdProfil_Excel.BackColor = &H80FFFF

fraDetail.Enabled = True
fgSelect.Enabled = True


End Sub

Private Sub cmdProfil_Update_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass

Select Case cmdSelect_SQL_K
    Case "1", "2", "3", "3_H", "2_S":
            If IsNull(fraYSSIDOM0_Control) Then
                Call cmdSSIJRN_DOM("")
                If IsNull(cmdUpdate) Then
                    If cmdSelect_SQL_K = "1" Then fraDetail_Control_SSIUSRPRFK
                    Call cmdProfil_Quit_Click
                    fraDetail_Load
                End If
            End If
            
    Case "2_D":
            If fraProfil_Update.Visible Or fraProfil_Update_DIV.Visible Then
                
                Select Case mSSIDOMDIDX
                    Case "IBM": V = fraProfil_Control_IBM
                              If IsNull(V) Then
                                 mYSSIIBM0_Update = "New$"
                                 Call cmdSSIJRN_IBM("<X: IBM màj des profils>")
                             End If
                  Case "DIV": V = fraProfil_Control_DIV
                              If IsNull(V) Then
                                 Call cmdSSIJRN_DIV("<X: DIV màj des profils>")
                              End If
                    Case Else: V = "cmdProfil_Update_Click : non programmé"
                End Select
                
                If IsNull(V) Then
                    If IsNull(cmdUpdate) Then
                        cmdProfil_Update.Visible = False
                        cmdProfil_Update_DIV.Visible = False
                        Call cboProfil_DOM_Click
                    End If
                End If
            End If
End Select

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdCompte_Val_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass

    cmdUpdate_Init

If IsNull(fraYSSIDOM0_Control) Then
    mYSSIDOM0_Update = "Update+H"
    
    newYSSIDOM0 = oldYSSIDOM0
    Select Case newYSSIDOM0.SSIDOMDIDX
        Case "IBM"
        Case "SAA": newYSSIDOM0.SSIDOMPRFX = usrYSSISAA0.SSISAAPRFX
        Case "SAB": newYSSIDOM0.SSIDOMPRFX = usrYSSISAB0.SSISABPRFX
        Case "WIN": newYSSIDOM0.SSIDOMPRFX = usrYSSIWIN0.SSIWINPRFX
        Case "TIC": newYSSIDOM0.SSIDOMPRFX = usrYSSITIC0.SSITICPRFX
    End Select
    newYSSIDOM0.SSIDOMPRFK = " "
    newYSSIDOM0.SSIDOMYFCT = "VAL"
    newYSSIDOM0.SSIDOMYAMJ = DSys
    newYSSIDOM0.SSIDOMYHMS = time_Hms
    newYSSIDOM0.SSIDOMYUSR = usrName_UCase
    
    Call cmdSSIJRN_DOM("")
    
    If IsNull(cmdUpdate) Then
    
        Call cmdProfil_Quit_Click
        fraDetail_Load
    End If
End If
Me.Enabled = True: Me.MousePointer = 0


End Sub


Private Sub cmdProfil_Update_DIV_Click()
fraYSSIDIV0_Display
End Sub

Private Sub cmdSSIUSR_Delete_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass

If cmdSelect_SQL_K = "2" Then

    mYSSIUSR0_Update = "Delete+CMD"

    mYSSIUSR0_Update_CMD = "Update " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
                         & " set SSIUSRPRFX = '' " _
                         & " Where SSIUSRNAT = ' ' and SSIUSRPRFX = '" & oldYSSIUSR0.SSIUSRUIDX & "'"

    oldYSSIUSR0.SSIUSRYFCT = "SUP"
    newYSSIUSR0 = oldYSSIUSR0
    Call cmdSSIJRN_USR("")
    
    If IsNull(cmdUpdate) Then
        Call cmdSSIUSR_Quit_Click
        Select Case cmdSelect_SQL_K
            Case "1": Call cmdSelect_SQL_1
            Case "2": paramSSIUSRPRFX_Load: Call cmdSelect_SQL_2
        End Select
    End If
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSSIUSR_Histo_Click()
Dim xSQL As String, X As String, blnHisto_Init As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> cmdSSIUSR_Histo ........"): DoEvents
On Error GoTo Error_Handler

currentAction = "cmdSSIUSR_Histo_Click"
txtRTF.TextRTF = VB_RTF_Modèle
'______________________________________________________________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSRH " _
     & " left outer join " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
     & " on SSITXTNAT = SSIUSRNAT and SSITXTUIDN = SSIUSRUIDN and SSITXTDIDX = '' and SSITXTUIDX = '' " _
     & " and SSITXTTLNK = SSIUSRTLNK" _
     & " where SSIUSRNAT = '" & oldYSSIUSR0.SSIUSRNAT & "' and SSIUSRUIDN = " & oldYSSIUSR0.SSIUSRUIDN _
     & "  order by SSIUSRYVER"
     
Set rsSab = cnsab.Execute(xSQL)
  
Do While Not rsSab.EOF
    Call rsYSSIUSR0_GetBuffer(rsSab, xYSSIUSR0)
    If xYSSIUSR0.SSIUSRTLNK = 0 Then
        xYSSITXT0.SSITXTINFO = ""
    Else
        xYSSITXT0.SSITXTINFO = rsSab("SSITXTINFO")
    End If
    
    If Not blnHisto_Init Then
        blnHisto_Init = True
        hYSSIUSR0 = xYSSIUSR0
        hYSSITXT0 = xYSSITXT0
    End If
    
    
    X = X & cmdSSIUSR_Histo_txtRTF
    
    hYSSIUSR0 = xYSSIUSR0
    hYSSITXT0 = xYSSITXT0
    rsSab.MoveNext
Loop

Set rsSab = Nothing
xYSSIUSR0 = oldYSSIUSR0
xYSSITXT0 = oldYSSITXT0_USR
X = X & cmdSSIUSR_Histo_txtRTF
'______________________________________________________________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOMH " _
     & " left outer join " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
     & " on SSITXTNAT = SSIDOMNAT and SSITXTUIDN = SSIDOMUIDN and SSITXTDIDX = SSIDOMDIDX and SSITXTUIDX = SSIDOMUIDX " _
     & " and SSITXTTLNK = SSIDOMTLNK" _
     & " where SSIDOMNAT = '" & oldYSSIDOM0.SSIDOMNAT & "' and SSIDOMUIDN = " & oldYSSIDOM0.SSIDOMUIDN _
     & "  order by SSIDOMYAMJ , SSIDOMYHMS , SSIDOMYVER"
     
Set rsSab = cnsab.Execute(xSQL)
  
Do While Not rsSab.EOF
    Call rsYSSIDOM0_GetBuffer(rsSab, xYSSIDOM0)
    If xYSSIDOM0.SSIDOMTLNK = 0 Then
        xYSSITXT0.SSITXTINFO = ""
    Else
        If Not IsNull(rsSab("SSITXTINFO")) Then
            xYSSITXT0.SSITXTINFO = rsSab("SSITXTINFO")
        Else
            xYSSITXT0.SSITXTINFO = ""
        End If
        
    End If
    
    If Not blnHisto_Init Then
        blnHisto_Init = True
        hYSSIDOM0 = xYSSIDOM0
        hYSSITXT0 = xYSSITXT0
    End If
    
    
    X = X & cmdSSIDOM_Histo_txtRTF
    
    hYSSIDOM0 = xYSSIDOM0
    hYSSITXT0 = xYSSITXT0
    rsSab.MoveNext
Loop

Set rsSab = Nothing
xYSSIDOM0 = oldYSSIDOM0
xYSSITXT0 = oldYSSITXT0_DOM
If oldYSSIDOM0.SSIDOMYAMJ <> 0 Then X = X & cmdSSIDOM_Histo_txtRTF
'______________________________________________________________________________________________________________________

txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", X)
Call txtRTF_Visible
GoTo Exit_sub

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
Exit_sub:
 
    Call lstErr_AddItem(lstErr, cmdContext, "< cmdSSIUSR_Histo terminé"): DoEvents
    Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Function cmdSSIUSR_Histo_txtRTF() As String
Dim xRTF As String, X As String, xAttibut1 As String, xAttibut2 As String
    
'___________________________________________________________________________________________________
    If hYSSIUSR0.SSIUSRUIDX = xYSSIUSR0.SSIUSRUIDX Then
        xAttibut1 = "": xAttibut2 = ""
    Else
        xAttibut1 = "\highlight12 ": xAttibut2 = "\highlight0 "
    End If
    xRTF = "\fs18\cf13\b" & xYSSIUSR0.SSIUSRNAT & " " & xYSSIUSR0.SSIUSRUIDN & "-" & xYSSIUSR0.SSIUSRYVER & " " & xAttibut1 & xYSSIUSR0.SSIUSRUIDX & xAttibut2 & "\b0 " _
        & "\tab \tab \tab \fs16\cf6" & " màj : " & xYSSIUSR0.SSIUSRYUSR & " " & dateImp10_S(xYSSIUSR0.SSIUSRYAMJ) & " " & timeImp8(xYSSIUSR0.SSIUSRYHMS) & "\cf1\fs18\par "
'___________________________________________________________________________________________________
    If hYSSIUSR0.SSIUSRDECH = xYSSIUSR0.SSIUSRDECH Then
        xAttibut1 = "": xAttibut2 = ""
    Else
        xAttibut1 = "\highlight12 ": xAttibut2 = "\highlight0 "
    End If
    If xYSSIUSR0.SSIUSRDECH = 0 Then
        X = "      "
    Else
        X = dateImp10_S(xYSSIUSR0.SSIUSRDECH)
    End If
    
    X = " date limite d'activité : \cf13 " & xAttibut1 & X & xAttibut2 & "\cf0 "
'___________________________________________________________________________________________________
    If hYSSIUSR0.SSIUSRSTAK = xYSSIUSR0.SSIUSRSTAK Then
        xAttibut1 = "": xAttibut2 = ""
    Else
        xAttibut1 = "\highlight12 ": xAttibut2 = "\highlight0 "
    End If
    Select Case xYSSIUSR0.SSIUSRSTAK
        Case " ":
                xRTF = xRTF & " - " & xAttibut1 & "Utilisateur   : \cf13\b ACTIF \b0\cf0 " & xAttibut2 & X & "\par "
        Case Else
                xRTF = xRTF & " - " & xAttibut1 & "Utilisateur   : \cf10\b INACTIF \b0\cf0 " & xAttibut2 & X & "\par "
    End Select
'___________________________________________________________________________________________________
    If hYSSIUSR0.SSIUSRPRFX = xYSSIUSR0.SSIUSRPRFX Then
        xAttibut1 = "": xAttibut2 = ""
    Else
        xAttibut1 = "\highlight12 ": xAttibut2 = "\highlight0 "
    End If
    xRTF = xRTF & " - Profil        : \cf13\b " & xAttibut1 & xYSSIUSR0.SSIUSRPRFX & xAttibut2 & "\b0\cf0\par "
    
'___________________________________________________________________________________________________
    
     If hYSSIUSR0.SSIUSRPRFD = xYSSIUSR0.SSIUSRPRFD Then
        xAttibut1 = "": xAttibut2 = ""
    Else
        xAttibut1 = "\highlight12 ": xAttibut2 = "\highlight0 "
    End If
   If xYSSIUSR0.SSIUSRPRFD = 0 Then
        X = ""
    Else
        X = " en date du \cf13 " & xAttibut1 & dateImp10_S(xYSSIUSR0.SSIUSRPRFD) & xAttibut2 & "\cf0 "
    End If
'___________________________________________________________________________________________________
    
     If hYSSIUSR0.SSIUSRPRFK = xYSSIUSR0.SSIUSRPRFK Then
        xAttibut1 = "": xAttibut2 = ""
    Else
        xAttibut1 = "\highlight12 ": xAttibut2 = "\highlight0 "
    End If
    Select Case xYSSIUSR0.SSIUSRPRFK
        Case " ":
                xRTF = xRTF & " - habilitations :" & xAttibut1 & "\cf13\b conformes \b0\cf0 " & xAttibut2 & X & "\par "
        Case "N":
                xRTF = xRTF & " - habilitations :" & xAttibut1 & "\cf10\b  NON conformes \b0\cf0 " & xAttibut2 & X & "\par "
         Case "#":
                xRTF = xRTF & " - habilitations :" & xAttibut1 & "\cf14  NON définies \cf0 " & xAttibut2 & X & "\par "
        Case "X":
                xRTF = xRTF & " - habilitations :" & xAttibut1 & "\cf14 X compte clos" & xAttibut2 & "\cf0\par "
       Case Else
                xRTF = xRTF & " - " & xAttibut1 & "\cf10\b habilitations : code inconnu = " & xYSSIUSR0.SSIUSRPRFK & xAttibut2 & " \b0\cf0\par "
    End Select
    
    '___________________________________________________________________________________________________
    If xYSSIUSR0.SSIUSRTLNK <> 0 Then
         If hYSSITXT0.SSITXTINFO = xYSSITXT0.SSITXTINFO Then
            xAttibut1 = "": xAttibut2 = ""
        Else
            xAttibut1 = "\highlight12": xAttibut2 = "\highlight0 "
        End If
        xRTF = xRTF & xAttibut1 & "\tab \cf13 " & Replace(xYSSITXT0.SSITXTINFO, vbCrLf, "\par \tab ") & xAttibut2 & "\cf0\par "
    End If
    xRTF = xRTF & "\cf8 ______________________________________________________________________\par  "

    cmdSSIUSR_Histo_txtRTF = xRTF
    
'"\cf2\highlight10 annulé le :" & " \cf0\cf0  "
End Function

Private Function cmdSSIDOM_Histo_txtRTF() As String
Dim xRTF As String, X As String, xAttibut1 As String, xAttibut2 As String
    
'___________________________________________________________________________________________________
    If hYSSIDOM0.SSIDOMUIDX = xYSSIDOM0.SSIDOMUIDX Then
        xAttibut1 = "": xAttibut2 = ""
    Else
        xAttibut1 = "\highlight12 ": xAttibut2 = "\highlight0 "
    End If
    xRTF = "\fs18\cf13\b" & xYSSIDOM0.SSIDOMNAT & " " & xYSSIDOM0.SSIDOMUIDN & "-" & xYSSIDOM0.SSIDOMYVER & " " & xAttibut1 & xYSSIDOM0.SSIDOMUIDX & xAttibut2 & "\b0 " _
        & "\tab \tab \tab \fs16\cf6\highlight12" & xYSSIDOM0.SSIDOMYFCT & "\highlight0 : " & xYSSIDOM0.SSIDOMYUSR & " " & dateImp10_S(xYSSIDOM0.SSIDOMYAMJ) & " " & timeImp8(xYSSIDOM0.SSIDOMYHMS) & "\cf1\fs18\par\par "
'___________________________________________________________________________________________________
    If hYSSIDOM0.SSIDOMDECH = xYSSIDOM0.SSIDOMDECH Then
        xAttibut1 = "": xAttibut2 = ""
    Else
        xAttibut1 = "\highlight12 ": xAttibut2 = "\highlight0 "
    End If
    If xYSSIDOM0.SSIDOMDECH = 0 Then
        X = "      "
    Else
        X = dateImp10_S(xYSSIDOM0.SSIDOMDECH)
    End If
    
    X = " date limite d'activité : \cf13 " & xAttibut1 & X & xAttibut2 & "\cf0 "
'___________________________________________________________________________________________________
    If hYSSIDOM0.SSIDOMSTAK = xYSSIDOM0.SSIDOMSTAK Then
        xAttibut1 = "": xAttibut2 = ""
    Else
        xAttibut1 = "\highlight12 ": xAttibut2 = "\highlight0 "
    End If
    Select Case xYSSIDOM0.SSIDOMSTAK
        Case " ":
                xRTF = xRTF & " - " & xAttibut1 & "Utilisateur   : \cf13\b ACTIF \b0\cf0 " & xAttibut2 & X & "\par "
        Case Else
                xRTF = xRTF & " - " & xAttibut1 & "Utilisateur   : \cf10\b INACTIF \b0\cf0 " & xAttibut2 & X & "\par "
    End Select
'___________________________________________________________________________________________________

     If hYSSIDOM0.SSIDOMUIDD = xYSSIDOM0.SSIDOMUIDD Then
        xAttibut1 = "\cf13 ": xAttibut2 = ""
    Else
        xAttibut1 = "\highlight12\cf10\b ": xAttibut2 = "\b0\highlight0 "
    End If
   If xYSSIDOM0.SSIDOMUIDD = 0 Then
        X = ""
    Else
        X = "  " & xAttibut1 & xYSSIDOM0.SSIDOMUIDD & " / " & Trim(xYSSIDOM0.SSIDOMUIDX) & xAttibut2 & "\cf0 "
    End If

    If hYSSIDOM0.SSIDOMPRFX = xYSSIDOM0.SSIDOMPRFX Then
        xAttibut1 = "": xAttibut2 = ""
    Else
        xAttibut1 = "\highlight12 ": xAttibut2 = "\highlight0 "
    End If
    xRTF = xRTF & " - Profil \highlight15\cf2\b " & xYSSIDOM0.SSIDOMDIDX & "\highlight0    : \cf13 " & xAttibut1 & xYSSIDOM0.SSIDOMPRFX & xAttibut2 & "\b0\cf0 " & X & "\par "
    
'___________________________________________________________________________________________________
    
     If hYSSIDOM0.SSIDOMPRFD = xYSSIDOM0.SSIDOMPRFD Then
        xAttibut1 = "": xAttibut2 = ""
    Else
        xAttibut1 = "\highlight12 ": xAttibut2 = "\highlight0 "
    End If
   If xYSSIDOM0.SSIDOMPRFD = 0 Then
        X = ""
    Else
        X = " en date du \cf14 " & xAttibut1 & dateImp10_S(xYSSIDOM0.SSIDOMPRFD) & xAttibut2 & "\cf0 "
    End If
'___________________________________________________________________________________________________
    
     If hYSSIDOM0.SSIDOMPRFK = xYSSIDOM0.SSIDOMPRFK Then
        xAttibut1 = "": xAttibut2 = ""
    Else
        xAttibut1 = "\highlight12 ": xAttibut2 = "\highlight0 "
    End If
    Select Case xYSSIDOM0.SSIDOMPRFK
        Case " ":
                xRTF = xRTF & " - habilitations :" & xAttibut1 & "\cf13\b Conformes \b0\cf0 " & xAttibut2 & X & "\par "
        Case "N":
                xRTF = xRTF & " - habilitations :" & xAttibut1 & "\cf10\b  NON Conformes \b0\cf0 " & xAttibut2 & X & "\par "
         Case "#":
                xRTF = xRTF & " - habilitations :" & xAttibut1 & "\cf14  NON définies \cf0 " & xAttibut2 & X & "\par "
        Case "X":
                xRTF = xRTF & " - habilitations :" & xAttibut1 & "\cf14 X compte clos" & xAttibut2 & "\cf0\par "
       Case Else
                xRTF = xRTF & " - " & xAttibut1 & "\cf10\b habilitations : code inconnu = " & xYSSIDOM0.SSIDOMPRFK & xAttibut2 & " \b0\cf0\par "
    End Select
    
    '___________________________________________________________________________________________________
    If xYSSIDOM0.SSIDOMTLNK <> 0 Then
         If hYSSITXT0.SSITXTINFO = xYSSITXT0.SSITXTINFO Then
            xAttibut1 = "": xAttibut2 = ""
        Else
            xAttibut1 = "\highlight12": xAttibut2 = "\highlight0 "
        End If
        xRTF = xRTF & xAttibut1 & "\tab \cf13 " & Replace(xYSSITXT0.SSITXTINFO, vbCrLf, "\par \tab ") & xAttibut2 & "\cf0\par "
    End If
    xRTF = xRTF & "\cf8 ______________________________________________________________________\par  "

    cmdSSIDOM_Histo_txtRTF = xRTF
    
'"\cf2\highlight10 annulé le :" & " \cf0\cf0  "
End Function


Private Function cmdSSIIBM_Detail_txtRTF(lFct As String, lhighlight As Integer) As String
Dim xRTF As String, X As String, XX As String
Dim K As Integer

Select Case lFct
    Case "USR"
            Call cmdSSIIBM_Detail_usrRTF
    Case Else:
            For K = 1 To 30: usrRTF(K) = "\par ": Next K
End Select
'___________________________________________________________________________________________________
X = Trim(xYSSIIBM0.UPUPRF) & " "
arrRTF(1) = "\fs18\cf13\highlight" & lhighlight & " " & X & "\highlight0 " & Space$(22 - Len(X))


X = Trim(xYSSIIBM0.UPUSCL)
If X = "*USER" Then
    arrRTF(2) = "\fs18\cf13 " & X & Space$(22 - Len(X))
Else
    arrRTF(2) = "\fs18\cf10 " & X & Space$(22 - Len(X))
End If

X = Val(xYSSIIBM0.UPPWEI)
If Val(xYSSIIBM0.UPPWEI) > 0 Then
    arrRTF(3) = "\fs18\cf13 " & X & Space$(22 - Len(X))
Else
    arrRTF(3) = "\fs18\cf10 " & X & Space$(22 - Len(X))
End If

X = Trim(xYSSIIBM0.UPPWON)
If X = "*NO" Then
    arrRTF(4) = "\fs18\cf13 " & X & Space$(22 - Len(X))
Else
    arrRTF(4) = "\fs18\cf10 " & X & Space$(22 - Len(X))
End If

X = Trim(xYSSIIBM0.UPPWEX)
If X = "*NO" Then
    arrRTF(18) = "\fs18\cf13 " & X & Space$(22 - Len(X))
Else
    arrRTF(18) = "\fs18\cf10 " & X & Space$(22 - Len(X))
End If

X = Trim(xYSSIIBM0.UPSPAU)
If Len(X) >= 22 Then
    XX = X
Else
    XX = X & Space$(22 - Len(X))
End If

If X = "*NONE" Or X = "*SPLCTL" Then
    arrRTF(5) = "\fs18\cf13 " & XX
Else
    arrRTF(5) = "\fs18\cf10 " & XX
End If


X = Trim(xYSSIIBM0.UPLTCP)
If X = "*YES" Then
    arrRTF(14) = "\fs18\cf13 " & X & Space$(22 - Len(X))
Else
    arrRTF(14) = "\fs18\cf10 " & X & Space$(22 - Len(X))
End If

X = Trim(xYSSIIBM0.UPSTAT)
If X = "*ENABLED" Then
    arrRTF(16) = "\fs18\cf13 " & X & Space$(22 - Len(X))
Else
    arrRTF(16) = "\fs18\cf10 " & X & Space$(22 - Len(X))
End If
If xYSSIIBM0.SSIIBMNAT = "$" Then
    arrRTF(19) = "\fs18\cf13 " & Space$(22)
    arrRTF(20) = "\fs18\cf13 " & Space$(22)
    arrRTF(21) = "\fs18\cf13 " & Space$(22)
    arrRTF(22) = "\fs18\cf13 " & Space$(22)
    X = Trim(xYSSIIBM0.SSIIBMUIDD) & " / " & xYSSIIBM0.SSIIBMYVER: arrRTF(17) = "\fs18\cf13 " & X & Space$(22 - Len(X))
Else
    arrRTF(19) = "\fs18\cf13 " & dateImp10_S(xYSSIIBM0.UPCRTD) & Space$(12)
    arrRTF(20) = "\fs18\cf13 " & dateImp10_S(xYSSIIBM0.UPCHGD) & Space$(12)
    arrRTF(21) = "\fs18\cf13 " & dateImp10_S(xYSSIIBM0.UPPSOD) & Space$(12)
    arrRTF(22) = "\fs18\cf13 " & dateImp10_S(xYSSIIBM0.UPPWCD) & Space$(12)
    X = Trim(xYSSIIBM0.UPUID) & " / " & xYSSIIBM0.SSIIBMYVER: arrRTF(17) = "\fs18\cf13 " & X & Space$(22 - Len(X))
End If

X = Trim(xYSSIIBM0.UPTEXT)
If Len(X) >= 21 Then
    arrRTF(10) = "\fs18\cf13 " & X
Else
    arrRTF(10) = "\fs18\cf13 " & X & Space$(22 - Len(X))
End If

X = Trim(xYSSIIBM0.UPINPL) & "/" & Trim(xYSSIIBM0.UPINPG): arrRTF(6) = "\fs18\cf13 " & X & Space$(22 - Len(X))
X = Trim(xYSSIIBM0.UPJBDL) & "/" & Trim(xYSSIIBM0.UPJBDS): arrRTF(7) = "\fs18\cf13 " & X & Space$(22 - Len(X))
X = Trim(xYSSIIBM0.UPGRPF): arrRTF(8) = "\fs18\cf13 " & X & Space$(22 - Len(X))
X = Trim(xYSSIIBM0.UPCRLB): arrRTF(12) = "\fs18\cf13 " & X & Space$(22 - Len(X))
X = Trim(xYSSIIBM0.UPINML) & "/" & Trim(xYSSIIBM0.UPINMN): arrRTF(13) = "\fs18\cf13 " & X & Space$(22 - Len(X))
X = Trim(xYSSIIBM0.UPATPL) & "/" & Trim(xYSSIIBM0.UPATPG): arrRTF(15) = "\fs18\cf13 " & X & Space$(22 - Len(X))
X = Trim(xYSSIIBM0.UPSPEN): arrRTF(11) = "\fs18\cf13 " & X & Space$(22 - Len(X))
X = Trim(xYSSIIBM0.UPGRAU): arrRTF(9) = "\fs18\cf13 " & X & Space$(22 - Len(X))
arrRTF(23) = "\par\highlight" & lhighlight & "\fs18\cf13 " & xYSSIIBM0.SSIIBMPRFK _
    & " " & xYSSIIBM0.SSIIBMYFCT _
    & " " & xYSSIIBM0.SSIIBMYUSR _
    & " " & xYSSIIBM0.SSIIBMYAMJ & " " & xYSSIIBM0.SSIIBMYHMS & " " & xYSSIIBM0.SSIIBMYVER

xRTF = "\fs16\cf1 User profile name   : " & arrRTF(1) & usrRTF(1) _
     & "\fs16\cf1 User ID number      : " & arrRTF(17) & usrRTF(17) _
     & "\fs16\cf1 Text description    : " & arrRTF(10) & usrRTF(10) _
     & "\fs16\cf1 User class          : " & arrRTF(2) & usrRTF(2) _
     & "\fs16\cf1 Pwd expir. interval : " & arrRTF(3) & usrRTF(3) _
     & "\fs16\cf1 Pwd *None *Yes *No  : " & arrRTF(4) & usrRTF(4) _
     & "\fs16\cf1 Pwd set expired     : " & arrRTF(18) & usrRTF(18) _
     & "\fs16\cf1 Special authorities : " & arrRTF(5) & usrRTF(5) _
     & "\fs16\cf1 Initial program     : " & arrRTF(6) & usrRTF(6) _
     & "\fs16\cf1 Job description     : " & arrRTF(7) & usrRTF(7) _
     & "\fs16\cf1 Group profile       : " & arrRTF(8) & usrRTF(8) _
     & "\fs16\cf1 Current library     : " & arrRTF(12) & usrRTF(12) _
     & "\fs16\cf1 Initial menu        : " & arrRTF(13) & usrRTF(13) _
     & "\fs16\cf1 Limited capability  : " & arrRTF(14) & usrRTF(14) _
     & "\fs16\cf1 Attention program   : " & arrRTF(15) & usrRTF(15) _
     & "\fs16\cf1 Status              : " & arrRTF(16) & usrRTF(16) _
     & "\fs16\cf1 Special environ.    : " & arrRTF(11) & usrRTF(11) _
     & "\fs16\cf1 Group authority     : " & arrRTF(9) & usrRTF(9) _
     & "\fs16\cf1 Creation date       : " & arrRTF(19) & usrRTF(19) _
     & "\fs16\cf1 Previous sign-on    : " & arrRTF(21) & usrRTF(21) _
     & "\fs16\cf1 Password change     : " & arrRTF(22) & usrRTF(22) _
     & "\fs16\cf1 Change date         : " & arrRTF(20) & usrRTF(20) _
        & "\par\cf8 _______________________________________________________________________\par" _
        & arrRTF(23) & usrRTF(23) & "\highlight0 "

'___________________________________________________________________________________________________

If cmdSelect_SQL_K <> "1" Then
    xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight11 Utilisateurs actifs :\highlight0\ulnone"
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIIBM0 " _
         & " where SSIDOMNAT = ' '  and SSIDOMDIDX = 'IBM' and SSIDOMPRFX = '" & xYSSIIBM0.UPUPRF & "'" _
         & " and SSIDOMPRFK <> 'X' and SSIIBMNAT = ' ' and SSIIBMUIDD = SSIDOMUIDD order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        'kNAT_U = kNAT_U + 1
        
        xRTF = xRTF & "\par\tab\fs16\cf13\b " & Mid$(rsSab("SSIDOMUIDX"), 1, 12) _
                    & "\b0\cf2  : " & rsSab("UPTEXT")
        rsSab.MoveNext
    Loop
    xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight9 Utilisateurs INACTIFS :\highlight0\ulnone"
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIIBM0 " _
         & " where SSIDOMNAT = ' '  and SSIDOMDIDX = 'IBM' and SSIDOMPRFX = '" & xYSSIIBM0.UPUPRF & "'" _
         & " and SSIDOMPRFK = 'X' and SSIIBMNAT = ' ' and SSIIBMUIDD = SSIDOMUIDD order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        'kNAT_U = kNAT_U + 1
        
        xRTF = xRTF & "\par\tab\fs16\cf13\b " & Mid$(rsSab("SSIDOMUIDX"), 1, 12) _
                    & "\b0\cf2  : " & rsSab("UPTEXT")
        rsSab.MoveNext
    Loop
End If
xRTF = xRTF & "\par\cf8 _______________________________________________________________________\par"

Exit_sub:
    cmdSSIIBM_Detail_txtRTF = xRTF
End Function
Private Function cmdSSISAA_Detail_txtRTF(lFct As String, lhighlight As Integer) As String
Dim xRTF As String, X As String, xAttibut1 As String, xAttibut2 As String
Dim K1 As Integer, K2 As Integer, prfhighlight As Integer

If Trim(usrYSSISAA0.SSISAAUIDX) <> "" Then
'=======================================================================================================
     If Trim(usrYSSISAA0.SSISAAPRFX) <> Trim(prfYSSISAA0.SSISAAUIDX) Then
        X = "\par\tab\cf8\highlight12 Profil SAA   : \cf10\b " & Trim(usrYSSISAA0.SSISAAPRFX) _
                    & " <> " & Trim(prfYSSISAA0.SSISAAUIDX) & "\b0\cf1  (profil RSSI)\highlight0 "
        prfhighlight = 12
    Else
        prfhighlight = 11
        X = "\par\tab\cf8 Profil SAA   : \cf13\b\highlight11 " & Trim(usrYSSISAA0.SSISAAPRFX) & "\b0\cf0\highlight0 "
    End If
    
    xRTF = "\fs16\ul\cf1 Compte SAA : \fs18\b\cf13\highlight" & lhighlight & " " & Trim(usrYSSISAA0.SSISAAUIDX) _
         & "  \highlight0\cf1   (" & usrYSSISAA0.SSISAAUIDD & ") " _
         & "\cf2  => - " & Trim(usrYSSISAA0.SSISAAUNOM) & "\b0\ulnone\par " _
         & "\par\tab\cf8 Actif        : \cf7 " & cmdSSISTAK_Detail_txtRTF(xYSSISAA0.SSISAASTAK, 0) _
         & X _
         & "\par\tab\cf8 Evénement    : \cf7 " & usrYSSISAA0.SSISAAYFCT _
         & "  (v" & usrYSSISAA0.SSISAAYVER & ")" _
         & "\par\tab\cf8 màj par      : \cf7 " & usrYSSISAA0.SSISAAYUSR _
         & " le " & dateImp10_S(usrYSSISAA0.SSISAAYAMJ) & " " & timeImp8(usrYSSISAA0.SSISAAYHMS) _
         & "\par\par\tab " & cmdSSISAA_Detail_txtRTF_SSISAAINFO(usrYSSISAA0.SSISAAINFO, "=") _
         & "\par\cf8 _______________________________________________________________________\par "
End If

'============================================================================================
If lFct = "YSSISAAH" Then GoTo Exit_sub
'============================================================================================

xRTF = xRTF & "\fs16\ul\cf1 Profil SAA : \fs18\b\cf13\highlight15  " & Trim(prfYSSISAA0.SSISAAUIDX) _
     & "  \highlight0\cf2  =>  - " & Trim(prfYSSISAA0.SSISAAUNOM) & "\b0\ulnone\par "
'____________________________________________________________________________________________

If cmdSelect_SQL_K <> "1" Then
    xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight11 Utilisateurs actifs :\highlight0\ulnone"
    
   ' X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
   '      & " where SSISAANAT = ' ' and SSISAAPRFX = '" & prfYSSISAA0.SSISAAUIDX & "'" _
   '      & " and SSISAAPRFK <> 'X' order by SSISAAUIDX"
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
         & " where SSIDOMNAT = ' '  and SSIDOMDIDX = 'SAA' and SSIDOMPRFX = '" & prfYSSISAA0.SSISAAUIDX & "'" _
         & " and SSIDOMPRFK <> 'X' and SSISAANAT = ' ' and SSISAAUIDD = SSIDOMUIDD order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
       ' kNAT_U = kNAT_U + 1
        
        xRTF = xRTF & "\par\tab\fs16\cf13\b " & Mid$(rsSab("SSISAAUIDX"), 1, 12) _
                    & "\b0\cf2  : " & rsSab("SSISAAUNOM")
        rsSab.MoveNext
    Loop
    xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight9 Utilisateurs INACTIFS :\highlight0\ulnone"
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
         & " where SSIDOMNAT = ' '  and SSIDOMDIDX = 'SAA' and SSIDOMPRFX = '" & prfYSSISAA0.SSISAAUIDX & "'" _
         & " and SSIDOMPRFK = 'X' and SSISAANAT = ' ' and SSISAAUIDD = SSIDOMUIDD order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
       ' kNAT_U = kNAT_U + 1
        
        xRTF = xRTF & "\par\tab\fs16\cf13\b " & Mid$(rsSab("SSISAAUIDX"), 1, 12) _
                    & "\b0\cf2  : " & rsSab("SSISAAUNOM")
        rsSab.MoveNext
    Loop
End If

'____________________________________________________________________________________________
xRTF = xRTF & "\par\par\tab\fs16\ul\cf1\highlight15 Application-Function:\highlight0\ulnone"

X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
     & " where SSISAANAT = 'P' and SSISAAUIDX = '" & prfYSSISAA0.SSISAAUIDX & "'" _
     & " order by SSISAAUSEQ"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    K2 = rsSab("SSISAAUSEQ")
    K1 = Fix(K2 / 1000)
    K2 = K2 - K1 * 1000
    xRTF = xRTF & "\par\tab\cf6 Application  : \cf13 " & arrSAA_App_Code(K1) _
                & "\par\tab\cf6 Function     : \cf13 " & arrSAA_Function_Code(K2) _
                & "\par\tab\cf8 Evénement    : \cf7 " & rsSab("SSISAAYFCT") _
                & "  (v" & rsSab("SSISAAYVER") & ")" _
                & "\par\tab\cf8 màj par      : \cf7 " & rsSab("SSISAAYUSR") _
                & " le " & dateImp10_S(rsSab("SSISAAYAMJ")) & " " & timeImp8(rsSab("SSISAAYHMS")) _
               & "\par\par\tab\cf6 " & cmdSSISAA_Detail_txtRTF_SSISAAINFO(rsSab("SSISAAINFO"), ":")
    

    rsSab.MoveNext
Loop
xRTF = xRTF & "\par\cf8 _______________________________________________________________________\par "


'============================================================================================

Exit_sub:
cmdSSISAA_Detail_txtRTF = xRTF
End Function

Private Function cmdSSISAA_Detail_txtRTF_Param(lFct As String, lhighlight As Integer) As String
Dim xRTF As String, X As String

If usrYSSISAA0.SSISAANAT = "P" Then
    Dim K1 As Integer, K2 As Integer
    K2 = usrYSSISAA0.SSISAAUSEQ
    K1 = Fix(K2 / 1000)
    K2 = K2 - K1 * 1000
    X = "\par\tab\cf6 Application  : \cf5 " & arrSAA_App_Code(K1) _
      & "\par\tab\cf6 Function     : \cf5 " & arrSAA_Function_Code(K2)
Else
    X = ""

End If

'=======================================================================================================
    xRTF = "\fs16\cf1 " & lFct & "        : " _
         & "\fs18\cf13\highlight" & lhighlight & " " & Trim(usrYSSISAA0.SSISAAUIDX) & "  /" & usrYSSISAA0.SSISAAUSEQ & "\highlight0 " _
         & " - " & Trim(usrYSSISAA0.SSISAAUNOM) & "\par " _
         & "\par\tab\cf8 Id SSI       : \cf7 " & usrYSSISAA0.SSISAAUIDD _
         & "\par\tab\cf8 Evénement    : \cf7 " & usrYSSISAA0.SSISAAYFCT _
         & "  (v" & usrYSSISAA0.SSISAAYVER & ")" _
         & "\par\tab\cf8 màj par      : \cf7 " & usrYSSISAA0.SSISAAYUSR _
         & " le " & dateImp10_S(usrYSSISAA0.SSISAAYAMJ) & " " & timeImp8(usrYSSISAA0.SSISAAYHMS) _
         & "\par " & X & "\par " _
         & "\par\tab\cf7 " & cmdSSISAA_Detail_txtRTF_SSISAAINFO(usrYSSISAA0.SSISAAINFO, ":") _
         & "\par\cf8 _______________________________________________________________________\par  "
'___________________________________________________________________________________________________



'============================================================================================

Exit_sub:
cmdSSISAA_Detail_txtRTF_Param = xRTF
End Function


Private Function cmdSSISAB_Detail_Display(lFct As String)

On Error GoTo Exit_sub
cmdSSISAB_Detail_Display = "?"

If IsNull(cmdSSISAB_Detail_Load(lFct)) Then

    txtRTF.TextRTF = VB_RTF_Modèle
    
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", mRTF)
    
    Call txtRTF_Visible
    
    cmdSSISAB_Detail_Display = Null
End If
Exit_sub:

End Function
Private Function cmdSSISAB_Detail_Load(lFct As String)
Dim X As String, whighlight As Integer
On Error GoTo Exit_sub
cmdSSISAB_Detail_Load = "?"
whighlight = 11
txtRTF.Visible = False
'_____________________________________________________________________________________________
If Trim(usrYSSISAB0.SSISABUIDX) <> "" Then
    
    If lFct = "YSSISAB*" Then
        X = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSISABH " _
          & " where SSISABNAT = '" & usrYSSISAB0.SSISABNAT & " ' and SSISABUIDX = '" & usrYSSISAB0.SSISABUIDX & "'" _
          & " and SSISABULOT = 0  and SSISABYVER = " & usrYSSISAB0.SSISABYVER
        Set rsSab = cnsab.Execute(X)
        If rsSab(0) > 0 Then
            lFct = "YSSISABH"
        Else
            lFct = "SAB"
        End If
    End If
    Select Case lFct
        Case "YSSISABH"
             whighlight = 9
             X = "select * from " & paramIBM_Library_SABSPE & ".YSSISABH " _
               & " where SSISABNAT = '" & usrYSSISAB0.SSISABNAT & " ' and SSISABUIDX = '" & usrYSSISAB0.SSISABUIDX & "'" _
               & " and SSISABULOT = 0  and SSISABYVER = " & usrYSSISAB0.SSISABYVER
       Case Else
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
              & " where SSISABNAT = '" & usrYSSISAB0.SSISABNAT & " ' and SSISABUIDX = '" & usrYSSISAB0.SSISABUIDX & "'" _
              & " and SSISABULOT = 0"
    End Select

    Set rsSab = cnsab.Execute(X)
      
    If rsSab.EOF Then
        Call MsgBox(usrYSSISAB0.SSISABUIDX & " : Id utilisateur inconnue dans le domaine SAB " & mSSIDOMDIDX, vbCritical, "cmdSSISAB_Detail_Display")
        GoTo Exit_sub
    End If
    
    Call rsYSSISAB0_GetBuffer(rsSab, usrYSSISAB0)
    
    
End If
'_____________________________________________________________________________________________
If lFct = "SAB" Then
    prfYSSISAB0.SSISABNAT = "$"
    prfYSSISAB0.SSISABUIDX = Trim(usrYSSISAB0.SSISABPRFX)
End If

If usrYSSISAB0.SSISABNAT = "$" Then
    prfYSSISAB0 = usrYSSISAB0
Else
    If lFct <> "YSSISABH" And prfYSSISAB0.SSISABUIDX <> "" Then
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
             & " where SSISABNAT = '" & prfYSSISAB0.SSISABNAT & " ' and SSISABUIDX = '" & prfYSSISAB0.SSISABUIDX & "'" _
             & " and SSISABULOT = 0"
        Set rsSab = cnsab.Execute(X)
        
        If rsSab.EOF Then
            Call MsgBox(usrYSSISAB0.SSISABUIDX & " : Profil inconnu dans le domaine " & mSSIDOMDIDX, vbCritical, "cmdSSISAB_Detail_Display")
            GoTo Exit_sub
        End If
        
        Call rsYSSISAB0_GetBuffer(rsSab, prfYSSISAB0)
    End If
End If
'_____________________________________________________________________________________________
If oldYSSIDOM0.SSIDOMNAT = " " Then
    If oldYSSIDOM0.SSIDOMDIDX = "SAB" _
    And oldYSSIDOM0.SSIDOMUIDX = usrYSSISAB0.SSISABUIDX And oldYSSIDOM0.SSIDOMUIDD = usrYSSISAB0.SSISABUIDD Then
    
    Else
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
             & " where SSIDOMNAT = ' ' and SSIDOMDIDX = 'SAB'  and SSIDOMUIDX = '" & usrYSSISAB0.SSISABUIDX & "'" _
             & " and SSIDOMUIDD = " & usrYSSISAB0.SSISABUIDD
        Set rsSab = cnsab.Execute(X)
        
        If rsSab.EOF Then
            oldYSSISAB0.SSISABPRFX = "??????"
        Else
            Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
    
        End If
    End If
End If


'_____________________________________________________________________________________________


mRTF = cmdSSISAB_Detail_txtRTF(lFct, whighlight)


cmdSSISAB_Detail_Load = Null

Exit_sub:

End Function

Private Function cmdSSISAB_Display(lFct As String)
Dim X As String, blnSQL As Boolean, whighlight As Integer
On Error GoTo Exit_sub

If xYSSISAB0.SSISABNAT = " " Then
    usrYSSISAB0.SSISABUIDX = xYSSISAB0.SSISABUIDX
    Call cmdSSISAB_Detail_Display("SAB")
    GoTo Exit_sub
End If

cmdSSISAB_Display = "?"
'_____________________________________________________________________________________________
If lFct = "YSSISABH" Then
    whighlight = 9
   X = "select * from " & paramIBM_Library_SABSPE & ".YSSISABH " _
         & " where SSISABNAT = '" & xYSSISAB0.SSISABNAT & " ' and SSISABUIDX = '" & xYSSISAB0.SSISABUIDX & "'" _
     & " and SSISABULOT = " & xYSSISAB0.SSISABULOT & "  and SSISABYVER = " & xYSSISAB0.SSISABYVER
    Set rsSab = cnsab.Execute(X)
Else
    whighlight = 11
   X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
         & " where SSISABNAT = '" & xYSSISAB0.SSISABNAT & " ' and SSISABUIDX = '" & xYSSISAB0.SSISABUIDX & "'" _
     & " and SSISABULOT = " & xYSSISAB0.SSISABULOT & "  and SSISABYVER = " & xYSSISAB0.SSISABYVER
    Set rsSab = cnsab.Execute(X)
    If rsSab.EOF Then
        whighlight = 9
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSISABH " _
              & " where SSISABNAT = '" & xYSSISAB0.SSISABNAT & " ' and SSISABUIDX = '" & xYSSISAB0.SSISABUIDX & "'" _
          & " and SSISABULOT = " & xYSSISAB0.SSISABULOT & "  and SSISABYVER = " & xYSSISAB0.SSISABYVER
         Set rsSab = cnsab.Execute(X)
    End If
End If

''''Set rsSab = cnsab.Execute(X)
  
If rsSab.EOF Then
    Call MsgBox(xYSSISAB0.SSISABUIDX & " :  inconnu ", vbCritical, "cmdSSISAB_Detail_Display")
    GoTo Exit_sub
End If

Call rsYSSISAB0_GetBuffer(rsSab, xYSSISAB0)
    
    

'_____________________________________________________________________________________________
Select Case xYSSISAB0.SSISABNAT
    Case "$": mRTF = mRTF & cmdSSISAB_Display_ZMNUGRP0
    Case "H": mRTF = mRTF & cmdSSISAB_Display_ZMNUHLA0
    Case "2", "3", "4": mRTF = mRTF & cmdSSISAB_Display_ZMNUHLB0
    Case "C": mRTF = mRTF & cmdSSISAB_Display_ZMNUUTP0
    Case "D": mRTF = mRTF & cmdSSISAB_Display_ZMNUUTO0
    Case "M": mRTF = mRTF & cmdSSISAB_Display_ZMNUMEN0
    Case "W": usrYSSISAB0 = xYSSISAB0: mRTF = mRTF & cmdSSISAW_Detail_txtRTF(whighlight)
End Select
'_____________________________________________________________________________________________


'mRTF = cmdSSISAB_Detail_txtRTF(whighlight)

txtRTF.TextRTF = VB_RTF_Modèle

txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", mRTF)

Call txtRTF_Visible

cmdSSISAB_Display = Null

Exit_sub:

End Function

Private Function cmdSSITXT_Detail_txtRTF() As String
Dim X As String, K1 As Integer, K2 As Integer
On Error GoTo Exit_sub
cmdSSITXT_Detail_txtRTF = "? cmdSSITXT_Detail_txtRTF"
'_____________________________________________________________________________________________
   X = "select * from " & paramIBM_Library_SABSPE & ".YSSITXT0" _
         & " where SSITXTNAT = '" & xYSSITXT0.SSITXTNAT & "'" _
         & " and SSITXTUIDN = " & xYSSITXT0.SSITXTUIDN _
         & " and SSITXTDIDX = '" & xYSSITXT0.SSITXTDIDX & "'" _
         & " and SSITXTUIDX = '" & xYSSITXT0.SSITXTUIDX & "'" _
         & " and SSITXTUIDD = " & xYSSITXT0.SSITXTUIDD _
         & " and SSITXTTLNK = " & xYSSITXT0.SSITXTTLNK

Set rsSab = cnsab.Execute(X)
  
If rsSab.EOF Then
    Call MsgBox(xYSSITXT0.SSITXTDIDX & " " & xYSSITXT0.SSITXTUIDX & " :  inconnu ", vbCritical, "cmdSSITXT_Detail_Load")
    GoTo Exit_sub
End If

Call rsYSSITXT0_GetBuffer(rsSab, xYSSITXT0)

If xYSSITXT0.SSITXTNAT = "J" Then
    X = ""
    K1 = InStr(1, xYSSITXT0.SSITXTINFO, "<ORIG")
    If K1 > 0 Then
        K2 = InStr(K1, xYSSITXT0.SSITXTINFO, ">")
        If K2 > 0 Then X = "\cf13 " & arrJRN_Origine(Val(Mid$(xYSSITXT0.SSITXTINFO, K1 + 6, K2 - K1 - 6)))
    End If
    K1 = InStr(1, xYSSITXT0.SSITXTINFO, "<FCT")
    If K1 > 0 Then
        K2 = InStr(K1, xYSSITXT0.SSITXTINFO, ">")
        If K2 > 0 Then X = X & " -\cf14 " & Mid$(xYSSITXT0.SSITXTINFO, K1 + 5, K2 - K1 - 5)
    End If
    cmdSSITXT_Detail_txtRTF = "\fs18\cf13\highlight12\ul " & "Journal des évènements SSI : " & "\ulnone\highlight0\b  " & X & "\b0" _
            & "\par\tab\cf8 Utilisateur  : \cf7 " & xYSSITXT0.SSITXTUIDN _
            & "\par\tab\cf8 Domaine      : \cf7 " & xYSSITXT0.SSITXTDIDX _
            & "\par\tab\cf8 Code X       : \cf7 " & xYSSITXT0.SSITXTUIDX _
            & "\par\tab\cf8 Code N       : \cf7 " & xYSSITXT0.SSITXTUIDD _
            & "\par\tab\cf8 Lien TXT     : \cf7 " & xYSSITXT0.SSITXTTLNK _
            & "\par\tab\cf8 màj par      : \cf7 " & xYSSITXT0.SSITXTYUSR _
            & " le " & dateImp10_S(xYSSITXT0.SSITXTYAMJ) & " " & timeImp8(xYSSITXT0.SSITXTYHMS) _
            & "\par\cf8 _______________________________________________________________________\par"

Else
    cmdSSITXT_Detail_txtRTF = "\par\par\tab\cf14\highlight12\ul Annotation   : \highlight0\cf7 " & xYSSITXT0.SSITXTYUSR _
         & " le " & dateImp10_S(xYSSITXT0.SSITXTYAMJ) & " " & timeImp8(xYSSITXT0.SSITXTYHMS) & "\ulnone\par" _
         & "\par\tab\fs18\cf13\highlight12 " & Replace(Trim(xYSSITXT0.SSITXTINFO), vbCrLf, "\highlight0\par\tab ")
End If
    

'_____________________________________________________________________________________________


Exit_sub:

End Function


Private Function cmdSSISAB_Detail_txtRTF(lFct As String, lhighlight As Integer) As String
Dim xRTF As String, X As String, xAttibut1 As String, xAttibut2 As String
Dim K1 As Integer, K2 As Integer, prfhighlight As Integer
Dim kNAT_U As Integer, kNAT_2 As Integer, kNAT_3 As Integer, kNAT_C As Integer, kNAT_D As Integer
Dim blnEnd As Boolean

prfhighlight = 12
If Trim(usrYSSISAB0.SSISABUIDX) <> "" Then
'=======================================================================================================
    'X = "\fs18\cf13\highlight" & lhighlight & " " & Trim(usrYSSISAB0.SSISABUIDX) & "\highlight0 " & " - " & Trim(usrYSSISAB0.SSISABUNOM) & "\par "
    'xRTF = "\fs16\cf1 Utilisateur " & usrYSSISAB0.SSISABUIDD & "        : " & X
    xRTF = "\fs18\ul\cf1 Compte SAB : \cf13\highlight" & lhighlight & "\b " & Trim(usrYSSISAB0.SSISABUIDX) _
        & "  \highlight0\cf1   (" & usrYSSISAB0.SSISABUIDD & ") " & "\cf2  => " _
         & Trim(usrYSSISAB0.SSISABUNOM) & "\b0\ulnone"
'___________________________________________________________________________________________________
    If usrYSSISAB0.SSISABSTAK = " " Then
        xAttibut1 = "": xAttibut2 = ""
    Else
        xAttibut1 = "\highlight12 ": xAttibut2 = "\highlight0 "
    End If
    Select Case usrYSSISAB0.SSISABSTAK
        Case " ":
                xRTF = xRTF & "\par\tab\cf1" & xAttibut1 & "Utilisateur  : \cf13\b ACTIF \b0\cf0 " & xAttibut2
        Case Else
                xRTF = xRTF & "\par\tab\cf1" & xAttibut1 & "Utilisateur : \cf10\b INACTIF \b0\cf0 " & xAttibut2
    End Select
    If usrYSSISAB0.SSISABPRFX <> oldYSSIDOM0.SSIDOMPRFX Then ' prfYSSISAB0.SSISABUIDX Then
        xRTF = xRTF & "\par\tab\cf1\highlight12 Profil SAB   : \cf10\b " & Trim(usrYSSISAB0.SSISABPRFX) _
                    & " <> " & Trim(oldYSSIDOM0.SSIDOMPRFX) & "\b0\cf1    (profil RSSI)\highlight0"
        prfhighlight = 12
    Else
        prfhighlight = 11
        xRTF = xRTF & "\par\tab\cf1\highlight11 Profil SAB   : \cf13\b " & Trim(usrYSSISAB0.SSISABPRFX) & "\b0\cf0\highlight0"
    End If


    If Mid$(usrYSSISAB0.SSISABINFO, 1, 1) = "O" Then
        xRTF = xRTF & "\par\tab\cf1 Accès SAB    : \cf13 OUI"
    Else
        xRTF = xRTF & "\par\tab\cf1 Accès SAB    : \cf10 NON"
    End If
    xRTF = xRTF & "\par\tab\cf1 Groupe MENU  : \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 2, 10) _
                & "\par\tab\cf1 Groupe DROITS: \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 12, 10) _
                & "\par\tab\cf1 Groupe METIER: \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 22, 10) _
                & "\par\tab\cf1 file attente : \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 32, 10) _
                & "\par\tab\cf1 Langue       : \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 42, 1) _
                & "\par\tab\cf1 Menu service : \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 43, 1) _
                & "\par\tab\cf1 Agence défaut: \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 44, 3) _
                & "\par\tab\cf1 Service défa : \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 47, 2) _
                & "\par\tab\cf1 Sous-Service : \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 49, 2) _
                & "\par\tab\cf1 Grp Menu Srv : \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 51, 10) _
                & "\par\tab\cf1 Code générique \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 61, 1) _
                & "\par\tab\cf1 Poste travail: \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 62, 10) _
                & "\par\tab\cf1 Adresse mail : \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 72, 50) _
                & "\par " _
                & "\par\tab\cf8 SSI événement : \cf7 " & usrYSSISAB0.SSISABYFCT _
                & "  (v" & usrYSSISAB0.SSISABYVER & ")" _
                & "\par\tab\cf8 SSI màj par   : \cf7 " & usrYSSISAB0.SSISABYUSR _
                & " le " & dateImp10_S(usrYSSISAB0.SSISABYAMJ) & " " & timeImp8(usrYSSISAB0.SSISABYHMS) _
                & "\par\cf8 _______________________________________________________________________\par"
                
End If

'============================================================================================
If lFct = "YSSISABH" Then GoTo Exit_sub
'============================================================================================
X = "\fs18\ul\cf1 Profil SAB \cf13\highlight15" & "\b " & Trim(prfYSSISAB0.SSISABUIDX) & "  \highlight0 " & "\cf2  => " _
        & Mid$(prfYSSISAB0.SSISABUNOM, 1, 10) & "  " _
        & Mid$(prfYSSISAB0.SSISABUNOM, 11, 10) & "  " _
        & Mid$(prfYSSISAB0.SSISABUNOM, 21, 10) & "  " _
        & "\b0\ulnone\par " _
        & "\par\tab\cf8 Actif        : \cf13 " & cmdSSISTAK_Detail_txtRTF(prfYSSISAB0.SSISABSTAK, 0) _
        & "\par\tab\cf8 Id SSI       : \cf13 " & prfYSSISAB0.SSISABUIDD _
        & "\par\tab\cf8 Evénement    : \cf7 " & prfYSSISAB0.SSISABYFCT _
        & "  (v" & prfYSSISAB0.SSISABYVER & ")" _
        & "\par\tab\cf8 màj par      : \cf7 " & prfYSSISAB0.SSISABYUSR _
        & " le " & dateImp10_S(prfYSSISAB0.SSISABYAMJ) & " " & timeImp8(prfYSSISAB0.SSISABYHMS) _
        & "\par\cf8 _______________________________________________________________________" _

xRTF = xRTF & X
'___________________________________________________________________________________________________

If cmdSelect_SQL_K = "1" Then
    kNAT_U = -1
Else
    xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight11 Utilisateurs actifs :\highlight0\ulnone"
    
  '  X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
  '       & " where SSISABNAT = ' ' and SSISABPRFX = '" & prfYSSISAB0.SSISABUIDX & "'" _
  '       & " and SSISABPRFK <> 'X' order by SSISABUIDX"
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
         & " where SSIDOMNAT = ' '  and SSIDOMDIDX = 'SAB' and SSIDOMPRFX = '" & prfYSSISAB0.SSISABUIDX & "'" _
         & " and SSIDOMPRFK <> 'X' and SSISABNAT = ' ' and SSISABUIDD = SSIDOMUIDD order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        kNAT_U = kNAT_U + 1
        
        xRTF = xRTF & "\par\tab\fs16\cf13\b " & Mid$(rsSab("SSISABUIDX"), 1, 12) _
                    & "\b0\cf2  : " & rsSab("SSISABUNOM")
        rsSab.MoveNext
    Loop
    xRTF = xRTF & "\par\cf8 _______________________________________________________________________"
    xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight9 Utilisateurs INACTIFS :\highlight0\ulnone"
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
         & " where SSIDOMNAT = ' '  and SSIDOMDIDX = 'SAB' and SSIDOMPRFX = '" & prfYSSISAB0.SSISABUIDX & "'" _
         & " and SSIDOMPRFK = 'X' and SSISABNAT = ' ' and SSISABUIDD = SSIDOMUIDD order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        kNAT_U = kNAT_U + 1
        
        xRTF = xRTF & "\par\tab\fs16\cf13\b " & Mid$(rsSab("SSISABUIDX"), 1, 12) _
                    & "\b0\cf2  : " & rsSab("SSISABUNOM")
        rsSab.MoveNext
    Loop
    xRTF = xRTF & "\par\cf8 _______________________________________________________________________"
End If

'___________________________________________________________________________________________________

X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
     & " where SSISABNAT = '2' and SSISABUIDX = '" & Mid$(prfYSSISAB0.SSISABUNOM, 1, 10) & "'" _
     & " and SSISABSTAK = ' ' order by SSISABULOT desc"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    kNAT_2 = kNAT_2 + 1
    If kNAT_2 = 1 Then
        xAttibut1 = "\cf14\highlight15"
    Else
        xAttibut1 = "\cf12\highlight10"
    End If
    
    xRTF = xRTF & "\par\fs16\ul\cf1 Menus " & xAttibut1 & "\b lot " & rsSab("SSISABULOT") & "\cf13  " & Mid$(rsSab("SSISABUIDX"), 1, 14) _
                & "\highlight0\cf2  => " & Trim(rsSab("SSISABUNOM")) & "\b0\ulnone"
    X = rsSab("SSISABINFO")
    If Mid$(X, 1, 1) = "0" Then
        xRTF = xRTF & "\par\tab\cf1 Validé       : \cf10 NON"
    Else
        xRTF = xRTF & "\par\tab\cf1 Validé       : \cf13 OUI"
    End If
    xRTF = xRTF & "\par\tab\cf1 Date début   : \cf13 " & dateImp10_S(Mid$(X, 2, 7) + 19000000) & " " & timeImp8(Mid$(X, 9, 6))
    
    If Val(Mid$(X, 15, 7)) > 0 Then
        xRTF = xRTF & "\par\tab\cf1 Date fin     : \cf13 " & dateImp10_S(Mid$(X, 15, 7) + 19000000) & " " & timeImp8(Mid$(X, 22, 6))
    End If
        xRTF = xRTF & "\par\tab\cf1 màj par      : \cf13 " & arrMNURUTCUT(Val(Mid$(X, 28, 4))) & " \cf1 le \cf13 " & dateImp10_S(Mid$(X, 32, 7) + 19000000) & " " & timeImp8(Mid$(X, 39, 6))
    rsSab.MoveNext
Loop

X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
     & " where SSISABNAT in ('3','C','D') and SSISABUIDX = '" & Mid$(prfYSSISAB0.SSISABUNOM, 11, 10) & "'" _
     & " and SSISABSTAK = ' ' order by SSISABULOT desc , SSISABNAT desc"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    Select Case rsSab("SSISABNAT")
        Case "3": kNAT_3 = kNAT_3 + 1
                If kNAT_3 = 1 Then
                    xAttibut1 = "\cf14\highlight15"
                Else
                    xAttibut1 = "\cf12\highlight10"
                End If
                
                xRTF = xRTF & "\par\par\par\fs16\ul\cf1 Données " & xAttibut1 & "\b lot " & rsSab("SSISABULOT") & "\cf13 - " & Mid$(rsSab("SSISABUIDX"), 1, 14) _
                        & "\highlight0\cf2 => " & Trim(rsSab("SSISABUNOM")) & "\b0\ulnone"
                X = rsSab("SSISABINFO")
                If Mid$(X, 1, 1) = "0" Then
                    xRTF = xRTF & "\par\tab\cf1 Validé       : \cf10 NON"
                Else
                    xRTF = xRTF & "\par\tab\cf1 Validé       : \cf13 OUI"
                End If
                xRTF = xRTF & "\par\tab\cf1 Date début   : \cf13 " & dateImp10_S(Mid$(X, 2, 7) + 19000000) & " " & timeImp8(Mid$(X, 9, 6))
                
                If Val(Mid$(X, 15, 7)) > 0 Then
                    xRTF = xRTF & "\par\tab\cf1 Date fin     : \cf13 " & dateImp10_S(Mid$(X, 15, 7) + 19000000) & " " & timeImp8(Mid$(X, 22, 6))
                End If
        xRTF = xRTF & "\par\tab\cf1 màj par      : \cf13 " & arrMNURUTCUT(Val(Mid$(X, 28, 4))) & " \cf1 le \cf13 " & dateImp10_S(Mid$(X, 32, 7) + 19000000) & " " & timeImp8(Mid$(X, 39, 6))
        Case "C": kNAT_C = kNAT_C + 1
                If kNAT_C = 1 Then
                    xAttibut1 = "\cf1\highlight15"
                Else
                    xAttibut1 = "\cf12\highlight10"
                End If
                xRTF = xRTF & "\par\par\tab\fs16\ul " & xAttibut1 & "\b lot " & rsSab("SSISABULOT") & "\cf13 - " & Mid$(rsSab("SSISABUIDX"), 1, 14) _
                        & "\highlight0\cf1 Droits des données communes  \b0\ulnone "
                X = Trim(rsSab("SSISABINFO"))
                For K1 = 1 To 99
                    Select Case Mid$(X, K1, 1)
                        Case Is = "1"
                            xRTF = xRTF & "\par\tab\cf13 " & arrMNURCLABR(K1) & " : \cf7 consultation"
                        Case Is = "2"
                            xRTF = xRTF & "\par\tab\cf13 " & arrMNURCLABR(K1) & " : \cf14 consultation + mise à jour autorisée"
                         Case Is = "3"
                            xRTF = xRTF & "\par\tab\cf13 " & arrMNURCLABR(K1) & " : \cf10 mise à jour autorisée sans consultation"
                   End Select
                Next K1
        Case "D": kNAT_D = kNAT_D + 1
                 If kNAT_D = 1 Then
                    xAttibut1 = "\cf1\highlight15"
                Else
                    xAttibut1 = "\cf12\highlight10"
                End If
               xRTF = xRTF & "\par\par\tab\fs16\ul " & xAttibut1 & "\b lot " & rsSab("SSISABULOT") & "\cf13 - " & Mid$(rsSab("SSISABUIDX"), 1, 14) _
                        & "\highlight0\cf1 Droits des données opérations \b0\ulnone "
                X = Trim(rsSab("SSISABINFO"))
                For K1 = 1 To Len(X) Step 6
                    If Mid$(X, K1 + 4, 1) = "O" Then
                        xRTF = xRTF & "\par\tab\cf13 " & Mid$(X, K1, 4) & "         : \cf14 mise à jour autorisée"
                    Else
                        xRTF = xRTF & "\par\tab\cf13 " & Mid$(X, K1, 4) & "         : \cf7 consultation"
                    End If
                Next K1
    End Select
    rsSab.MoveNext
Loop


'============================================================================================
'xRTF = xRTF & "\par\cf1 ___________________________________________________________________\par"

Select Case kNAT_2
    Case 1:
    Case 0: xRTF = xRTF & "\par\fs16\cf12\highlight10 " & prfYSSISAB0.SSISABUIDX & " ? aucun lot 'Menu' actif\highlight0"
    Case Else: xRTF = xRTF & "\par\fs16\cf12\highlight10 " & prfYSSISAB0.SSISABUIDX & " ? " & kNAT_2 & " lots 'Menu' actifs\highlight0"
End Select
Select Case kNAT_3
    Case 1
    Case 0: xRTF = xRTF & "\par\fs16\cf12\highlight10 " & prfYSSISAB0.SSISABUIDX & " ? aucun 'groupe Droit/données' actif\highlight0"
    Case Else: xRTF = xRTF & "\par\fs16\cf12\highlight10 " & prfYSSISAB0.SSISABUIDX & " ? " & kNAT_3 & " lots 'groupe Droit/données' actifs\highlight0"
End Select
Select Case kNAT_C
    Case 1
    Case 0: xRTF = xRTF & "\par\fs16\cf12\highlight10 " & prfYSSISAB0.SSISABUIDX & " ? aucun 'Classe Droit/données' actif\highlight0"
    Case Else: xRTF = xRTF & "\par\fs16\cf12\highlight10 " & prfYSSISAB0.SSISABUIDX & " ? " & kNAT_C & " lots 'Classe Droit/données' actifs\highlight0"
End Select
Select Case kNAT_D
    Case 1
    Case 0: xRTF = xRTF & "\par\fs16\cf12\highlight10 " & prfYSSISAB0.SSISABUIDX & " ? aucun 'Service Droit/données' actif\highlight0"
    Case Else: xRTF = xRTF & "\par\fs16\cf12\highlight10 " & prfYSSISAB0.SSISABUIDX & " ? " & kNAT_D & " lots 'Service Droit/données' actifs\highlight0"
End Select
Select Case kNAT_U
    Case 0: xRTF = xRTF & "\par\fs16\cf12\highlight10 " & prfYSSISAB0.SSISABUIDX & " ? aucun utilisateur actif\highlight0"
End Select
xRTF = xRTF & "\par\cf1 ___________________________________________________________________\par"

'============================================================================================

xRTF = xRTF & "\par\fs16\ul\cf1 Métiers " & xAttibut1 & "\b\cf13  " & Mid$(prfYSSISAB0.SSISABUNOM, 21, 10) _
            & "\highlight0\cf2  => Habilitations métier \b0\ulnone"

X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAM0 " _
     & " where SSISAMETA = 1 and SSISAMGRP = '" & Mid$(prfYSSISAB0.SSISABUNOM, 21, 10) & "'" _
     & " order by SSISAMREF,  SSISAMGRP, SSISAMCLA, SSISAMAPP, SSISAMCOD, SSISAMAGE, SSISAMSER, SSISAMSSE, SSISAMOPE, SSISAMNAT, SSISAMPRD, SSISAMAUT"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    Call rsYSSISAM0_GetBuffer(rsSab, rtfYSSISAM0)

    xRTF = xRTF & "\par\tab\cf1 App/code/Ag/Srv/Opé/Nat/.. " _
         & "\cf2 " & rtfYSSISAM0.SSISAMUIDD & " : " _
         & "\cf13 " & rtfYSSISAM0.SSISAMAPP & " " & rtfYSSISAM0.SSISAMCOD _
         & " " & rtfYSSISAM0.SSISAMAGE & " " & rtfYSSISAM0.SSISAMSER & " " & rtfYSSISAM0.SSISAMSSE _
         & " " & rtfYSSISAM0.SSISAMOPE & " " & rtfYSSISAM0.SSISAMNAT & " " & rtfYSSISAM0.SSISAMPRD & " " & rtfYSSISAM0.SSISAMAUT  ' _

    rsSab.MoveNext
Loop

'____________________________________________________________________________________________
Exit_sub:

cmdSSISAB_Detail_txtRTF = xRTF

End Function



Private Function cmdSSIIBM_Detail_usrRTF()
usrRTF(1) = ": " & "\fs18\cf6\highlight15 " & Trim(usrYSSIIBM0.UPUPRF) & " \highlight0\par "

If usrYSSIIBM0.UPUSCL = xYSSIIBM0.UPUSCL Then
    usrRTF(2) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPUSCL) & "\par\par "
Else
    usrRTF(2) = ": " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPUSCL) & "\highlight0\par\par "
End If

If usrYSSIIBM0.UPPWEI = xYSSIIBM0.UPPWEI Then
    usrRTF(3) = ": " & "\fs18\cf6 " & usrYSSIIBM0.UPPWEI & "\par\par "
Else
    usrRTF(3) = ": " & "\fs18\cf6\highlight12 " & usrYSSIIBM0.UPPWEI & "\highlight0\par\par "
End If

If usrYSSIIBM0.UPPWEX = xYSSIIBM0.UPPWEX Then
    usrRTF(18) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPPWEX) & "\par\par "
Else
    usrRTF(18) = ": " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPPWEX) & "\highlight0\par\par "
End If

If usrYSSIIBM0.UPPWON = xYSSIIBM0.UPPWON Then
    usrRTF(4) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPPWON) & "\par\par "
Else
    usrRTF(4) = ": " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPPWON) & "\highlight0\par\par "
End If


If usrYSSIIBM0.UPSPAU = xYSSIIBM0.UPSPAU Then
    'usrRTF(5) = "\par\tab\tab\tab\tab : " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPSPAU) & "\par\par "
    usrRTF(5) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPSPAU) & "\par\par "
Else
    'usrRTF(5) = "\par\tab\tab\tab\tab : " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPSPAU) & "\highlight0\par\par "
    usrRTF(5) = ": " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPSPAU) & "\highlight0\par\par "
End If

If usrYSSIIBM0.UPINPG = xYSSIIBM0.UPINPG And usrYSSIIBM0.UPINPL = xYSSIIBM0.UPINPL Then
    usrRTF(6) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPINPL) & "/" & Trim(usrYSSIIBM0.UPINPG) & "\par\par "
Else
    usrRTF(6) = ": " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPINPL) & "/" & Trim(usrYSSIIBM0.UPINPG) & "\highlight0\par\par "
End If

If usrYSSIIBM0.UPJBDS = xYSSIIBM0.UPJBDS And usrYSSIIBM0.UPJBDL = xYSSIIBM0.UPJBDL Then
    usrRTF(7) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPJBDL) & "/" & Trim(usrYSSIIBM0.UPJBDS) & "\par\par "
Else
    usrRTF(7) = ": " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPJBDL) & "/" & Trim(usrYSSIIBM0.UPJBDS) & "\highlight0\par\par "
End If

If usrYSSIIBM0.UPGRPF = xYSSIIBM0.UPGRPF Then
    usrRTF(8) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPGRPF) & "\par\par "
Else
    usrRTF(8) = ": " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPGRPF) & "\highlight0\par\par "
End If

If usrYSSIIBM0.UPGRAU = xYSSIIBM0.UPGRAU Then
    usrRTF(9) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPGRAU) & "\par\par "
Else
    usrRTF(9) = ": " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPGRAU) & "\highlight0\par\par "
End If

If usrYSSIIBM0.UPTEXT = xYSSIIBM0.UPTEXT Then
    usrRTF(10) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPTEXT) & "\par\par "
Else
    usrRTF(10) = ": " & "\fs18\cf6\highlight15 " & Trim(usrYSSIIBM0.UPTEXT) & "\highlight0\par\par "
End If

If usrYSSIIBM0.UPSPEN = xYSSIIBM0.UPSPEN Then
    usrRTF(11) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPSPEN) & "\par\par "
Else
    usrRTF(11) = ": " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPSPEN) & "\highlight0\par\par"
End If

If usrYSSIIBM0.UPCRLB = xYSSIIBM0.UPCRLB Then
    usrRTF(12) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPCRLB) & "\par\par "
Else
    usrRTF(12) = ": " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPCRLB) & "\highlight0\par\par "
End If

If usrYSSIIBM0.UPINMN = xYSSIIBM0.UPINMN And usrYSSIIBM0.UPINML = xYSSIIBM0.UPINML Then
    usrRTF(13) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPINML) & "/" & Trim(usrYSSIIBM0.UPINMN) & "\par\par "
Else
    usrRTF(13) = ": " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPINML) & "/" & Trim(usrYSSIIBM0.UPINMN) & "\highlight0\par\par "
End If

If usrYSSIIBM0.UPLTCP = xYSSIIBM0.UPLTCP Then
    usrRTF(14) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPLTCP) & "\par\par "
Else
    usrRTF(14) = ": " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPLTCP) & "\highlight0\par\par "
End If

If usrYSSIIBM0.UPATPG = xYSSIIBM0.UPATPG And usrYSSIIBM0.UPATPL = xYSSIIBM0.UPATPL Then
    usrRTF(15) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPATPL) & "/" & Trim(usrYSSIIBM0.UPATPG) & "\par\par "
Else
    usrRTF(15) = ": " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPATPL) & "/" & Trim(usrYSSIIBM0.UPATPG) & "\highlight0\par\par "
End If

If usrYSSIIBM0.UPSTAT = xYSSIIBM0.UPSTAT Then
    usrRTF(16) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPSTAT) & "\par\par "
Else
    usrRTF(16) = ": " & "\fs18\cf6\highlight12 " & Trim(usrYSSIIBM0.UPSTAT) & "\highlight0\par\par "
End If

'If usrYSSIIBM0.UPUID = xYSSIIBM0.UPUID Then
'    usrRTF(17) = ": " & "\fs18\cf6 " & Trim(usrYSSIIBM0.UPUID) & "\par\par "
'Else
    usrRTF(17) = ": " & "\fs18\cf6\highlight15 " & Trim(usrYSSIIBM0.UPUID) & " / " & usrYSSIIBM0.SSIIBMYVER & "\highlight0\par "
'End If

usrRTF(19) = ": " & "\fs18\cf6 " & dateImp10_S(usrYSSIIBM0.UPCRTD) & "\par "
usrRTF(20) = ": " & "\fs18\cf6 " & dateImp10_S(usrYSSIIBM0.UPCHGD) & "\par "
usrRTF(21) = ": " & "\fs18\cf6 " & dateImp10_S(usrYSSIIBM0.UPPSOD) & "\par "
usrRTF(22) = ": " & "\fs18\cf6 " & dateImp10_S(usrYSSIIBM0.UPPWCD) & "\par "

usrRTF(23) = "\tab\fs18\cf6\highlight15 " & usrYSSIIBM0.SSIIBMPRFK _
    & " " & usrYSSIIBM0.SSIIBMYFCT _
    & " " & usrYSSIIBM0.SSIIBMYUSR _
    & " " & usrYSSIIBM0.SSIIBMYAMJ & " " & usrYSSIIBM0.SSIIBMYHMS & " " & usrYSSIIBM0.SSIIBMYVER & "\highlight0"

End Function





Private Sub cmdSSIUSR_New_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> cmdSSIUSR_New ........"): DoEvents

Call rsYSSIUSR0_Init(newYSSIUSR0)
newYSSIUSR0.SSIUSRNAT = mSSIUSRNAT
oldYSSIUSR0 = newYSSIUSR0

Call rsYSSITXT0_Init(newYSSITXT0)
newYSSITXT0.SSITXTNAT = mSSIUSRNAT
oldYSSITXT0_USR = newYSSITXT0

txtSSIUSRUNIT_N.Locked = False

fraDetail_Display
    
Call lstErr_AddItem(lstErr, cmdContext, "< cmdSSIUSR_New terminé"): DoEvents
lstErr.Height = 480
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSSIUSR_PRF_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_SSI_cmdSSIUSR_PRF......"): DoEvents
cboProfil_DOM.Locked = False
txtRTF.Visible = False
lstW.Visible = False
cboProfil_DOM.ListIndex = 0
cmdSSIUSR_Update.Visible = False
fgSelect.Enabled = False
fraDetail.Enabled = False
fgProfil.Visible = False: chkProfil_DOM.Visible = False

fraProfil.Caption = oldYSSIUSR0.SSIUSRUIDX & " : ajouter un profil"
fraProfil.Visible = True
Call rsYSSIDIV0_Init(oldYSSIDIV0)
    
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_SSI_cmdSelect_Ok terminé"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSSIUSR_Quit_Click()
lstW.Visible = False
txtRTF.Visible = False
fraDetail.Visible = False
End Sub

Private Sub cmdSSIUSR_Update_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass


If IsNull(fraDetail_Control) Then
    Call cmdSSIJRN_USR("")
    If IsNull(cmdUpdate) Then
        Call cmdSSIUSR_Quit_Click
        Select Case cmdSelect_SQL_K
            Case "1": Call cmdSelect_SQL_1
            Case "2": paramSSIUSRPRFX_Load: Call cmdSelect_SQL_2
            Case "2_S": cboSSIUSRUNIT_Load: Call cmdSelect_SQL_2_S
        End Select
    End If
End If
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdYSSIDIV0_Quit_Click()
On Error Resume Next
fraYSSIDIV0.Visible = False

End Sub

Private Sub cmdYSSIDIV0_Update_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass

Select Case cmdSelect_SQL_K
    Case "1":
            If IsNull(fraYSSIDIV0_Control) Then
                Call cmdSSIJRN_DIV("<X: DIV màj compte>")
                If IsNull(cmdUpdate) Then
                    Call cmdProfil_Quit_Click
                    fraDetail_Load
                End If
            End If
            
End Select

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub fgCompte_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String
On Error Resume Next


If y <= fgCompte.RowHeightMin Then
    fgCompte.Visible = False
    Select Case fgCompte.Col
        Case 0: fgCompte_Sort1 = 0: fgCompte_Sort2 = 0: fgCompte_Sort
        Case 1:  fgCompte_Sort1 = 1: fgCompte_Sort2 = 1: fgCompte_Sort
    End Select
    fgCompte.Visible = True
Else
    If fgCompte.Rows > 1 Then
        Call fgCompte_Color(fgCompte_RowClick, MouseMoveUsr.BackColor, fgCompte_ColorClick)
        Select Case cmdSelect_SQL_K
            Case "1"
                Select Case mSSIDOMDIDX
                    Case "IBM"
                        usrYSSIIBM0.SSIIBMNAT = oldYSSIUSR0.SSIUSRNAT
                        fgCompte.Col = 0: usrYSSIIBM0.SSIIBMUIDD = Val(fgCompte.Text)
                        Call cmdSSIIBM_Detail_Display("", "USR")
                    Case "SAA"
                        usrYSSISAA0.SSISAANAT = oldYSSIUSR0.SSIUSRNAT
                        fgCompte.Col = 1: usrYSSISAA0.SSISAAUIDX = Trim(fgCompte.Text)
                        'Call cmdSSISAA_Detail_Display("SAA")
                        fgCompte.Col = 0: usrYSSISAA0.SSISAAUIDD = Trim(fgCompte.Text)
                        Call cmdSSISAA_Detail_Display("SSISAAUIDD")
                    Case "SAB"
                        usrYSSISAB0.SSISABNAT = oldYSSIUSR0.SSIUSRNAT
                        fgCompte.Col = 1: usrYSSISAB0.SSISABUIDX = Trim(fgCompte.Text)
                        Call cmdSSISAB_Detail_Display("SAB")
                    Case "WIN"
                        rtfYSSIWIN0.SSIWINNAT = oldYSSIUSR0.SSIUSRNAT
                       'fgCompte.Col = 1: rtfYSSIWIN0.SSIWINUIDX = Trim(fgCompte.Text)
                       ' Call cmdSSIWIN_Detail_Display("YSSIWIN0_UIDX")
                        fgCompte.Col = 0: rtfYSSIWIN0.SSIWINUIDD = Trim(fgCompte.Text)
                         Call cmdSSIWIN_Detail_Display("YSSIWIN0_UIDD")
                        usrYSSIWIN0 = rtfYSSIWIN0
                    Case "DIV"
                        rtfYSSIDIV0.SSIDIVNAT = oldYSSIUSR0.SSIUSRNAT
                        fgCompte.Col = 0: rtfYSSIDIV0.SSIDIVUIDX = Val(fgCompte.Text)
                        fgCompte.Col = 1: rtfYSSIDIV0.SSIDIVUIDX = Trim(fgCompte.Text)
                        Call cmdSSIDIV_Detail_Display("YSSIDIV0")
                        usrYSSIDIV0 = rtfYSSIDIV0
                        oldYSSIDIV0 = rtfYSSIDIV0
                    Case "TIC"
                        rtfYSSITIC0.SSITICNAT = oldYSSIUSR0.SSIUSRNAT
                        fgCompte.Col = 0: rtfYSSITIC0.SSITICUIDD = Val(fgCompte.Text)
                        fgCompte.Col = 1: rtfYSSITIC0.SSITICUIDX = Trim(fgCompte.Text)
                        Call cmdSSITIC_Detail_Display("YSSITIC0")
                        usrYSSITIC0 = rtfYSSITIC0
                        oldYSSITIC0 = rtfYSSITIC0
                        If oldYSSIDOM0.SSIDOMUIDX = "" Then
                            If prfYSSITIC0.SSITICUIDX <> usrYSSITIC0.SSITICPRFX Then
                                Call MsgBox("Profil sélectionné   : " & prfYSSITIC0.SSITICUIDX & vbCrLf _
                                           & "Profil de production : " & usrYSSITIC0.SSITICPRFX _
                                           , vbCritical, "ATHIC : incohérence des profils")
                            End If
                        End If
                End Select
                Select Case oldYSSIDIV0.SSIDIVDIDK
                    Case "TEREN", "SG", "UGM": cmdProfil_Update.Visible = arrHab(5)
                    Case Else: cmdProfil_Update.Visible = arrHab(2)
                End Select
             '   cmdProfil_Update.Visible = arrHab(2)
            Case "2"
        End Select
        
   End If
End If
fgCompte.LeftCol = 0

End Sub

Private Sub fgCompteH_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next


If y <= fgCompteH.RowHeightMin Then
Else
    If fgCompteH.Rows > 1 Then
        Call fgCompteH_Color(fgCompteH_RowClick, MouseMoveUsr.BackColor, fgCompteH_ColorClick)
        
        Select Case Trim(oldYSSIDOM0.SSIDOMDIDX)
            Case "IBM": Call fraCompteH_Display_IBM
            Case "SAA": Call fraCompteH_Display_SAA
            Case "SAB": Call fraCompteH_Display_SAB
            Case "SAB_W": fraCompteH_Display_SAB
            Case "WIN": Call fraCompteH_Display_WIN
            Case "DIV": Call fraCompteH_Display_DIV
            Case "MEL": Call fraCompteH_Display_MEL
            Case "TIC": Call fraCompteH_Display_TIC
        End Select
   End If
End If
fgCompteH.LeftCol = 0

End Sub


Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, K As Integer
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
        Select Case cmdSelect_SQL_K
            Case "1", "2"
                oldYSSIDOM0.SSIDOMNAT = oldYSSIUSR0.SSIUSRNAT
                oldYSSIDOM0.SSIDOMUIDN = oldYSSIUSR0.SSIUSRUIDN
                fgDetail.Col = 0: oldYSSIDOM0.SSIDOMDIDX = Trim(fgDetail.Text)
                fgDetail.Col = 3: oldYSSIDOM0.SSIDOMUIDD = Val(fgDetail.Text)
                fgDetail.Col = 2: oldYSSIDOM0.SSIDOMUIDX = Trim(fgDetail.Text)
                Call fraYSSIDOM0_Load
                'Call cbo_Scan(Trim(oldYSSIDOM0.SSIDOMDIDX), cboProfil_DOM)
            'Case "2"
        End Select
   End If
End If
fgDetail.LeftCol = 0


End Sub




Private Sub fgProfil_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String
On Error Resume Next
cmdProfil_Update.Visible = False
cmdProfil_Update_DIV.Visible = False
fraProfil_Update_DIV.Visible = False
If y <= fgProfil.RowHeightMin Then
    fgProfil.Visible = False
    
    Select Case fgProfil.Col
        Case 0: fgProfil_Sort1 = 0: fgProfil_Sort2 = 0: fgProfil_SortX 0
        Case 1:  fgProfil_Sort1 = 1: fgProfil_Sort2 = 1: fgProfil_Sort
        Case 2: fgProfil_Sort1 = 2: fgProfil_Sort2 = 2: fgProfil_Sort
        Case 3: fgProfil_Sort1 = 3: fgProfil_Sort2 = 3: fgProfil_Sort
    End Select
    fgProfil.Visible = True
Else
    If fgProfil.Rows > 1 Then Call fgProfil_Row_Click
End If

End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next


If y <= fgSelect.RowHeightMin Then
    If cmdSelect_SQL_K = "J" Then
    Else
        fgSelect.Visible = False
        Select Case fgSelect.Col
            Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_Sort
            Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
            Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
            Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        End Select
        fgSelect.Visible = True
    End If
Else
    If fgSelect.Rows > 1 Then Call fgSelect_Row_Click
        
End If
fgSelect.LeftCol = 0


End Sub

Private Sub Form_Activate()
Set XForm = Me

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: KeyCode = 0: cmdContext_Quit
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
fraDetail.Visible = False: lstW.Visible = False
fraProfil.Visible = False
fraProfil.Caption = ""
txtRTF.Visible = False: txtRTF = ""

cmdSelect_Ok.BackColor = vbGreen
cmdSSIUSR_New.Visible = False
lstW.Visible = False
cmdProfil_Print.Visible = False: cmdProfil_Excel.Visible = False
chkProfil_DOM.Visible = False
chkSSIDOMDECH = "0"
fraYSSIDIV0.Visible = False

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
    fraSelect_Options_1.Visible = False
    fraSelect_Options_J.Visible = False
    fraSelect_Options_4.Visible = False
    Select Case cmdSelect_SQL_K
        Case "1": cmdSelect_Ok.Visible = True: fraSelect_Options_1.Visible = True
                    mSSIUSRNAT = " "
                    cmdSSIUSR_New.Caption = "Ajouter un utilisateur"
                    cmdSSIUSR_New.Visible = arrHab(2)
                    cmdProfil_Histo.Visible = True
                    fgDetail.BackColorFixed = &HC0C000    'mColor_GB '&H80FF&
                    fraDetail_Update_PRF.Visible = True: fraDetail_Update_STAK.Visible = True
                    fraDetail_Update_SRV.Visible = False
        Case "2": cmdSelect_Ok.Visible = True: fraSelect_Options_1.Visible = True
                    mSSIUSRNAT = "$"
                    cmdSSIUSR_New.Caption = "Ajouter un modèle BIA"
                    cmdSSIUSR_New.Visible = arrHab(3)
                    cmdProfil_Histo.Visible = True
                    fgDetail.BackColorFixed = mColor_W1 'vbRed
                    fraDetail_Update_PRF.Visible = False: fraDetail_Update_STAK.Visible = True
                    fraDetail_Update_SRV.Visible = False
        Case "2_D": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
                   cmdProfil_Histo.Visible = False
        Case "2_S": cmdSelect_Ok.Visible = True
                    cmdSSIUSR_New.Caption = "Ajouter un service"
                    cmdSSIUSR_New.Visible = arrHab(2)
                    fraDetail_Update_PRF.Visible = False: fraDetail_Update_STAK.Visible = False
                    fraDetail_Update_SRV.Visible = True
                    txtSSIUSRUNIT_N.Locked = True
        Case "3": cmdSelect_Ok.Visible = True: fraSelect_Options_1.Visible = True
        Case "3_H": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
        Case "H": cmdSelect_Ok.Visible = True: fraSelect_Options_4.Visible = True
        Case "9_IBM", "9_SAA": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
        Case "J": cmdSelect_Ok.Visible = True: fraSelect_Options_J.Visible = True
    End Select

End If
End Sub


Private Sub cmdSelect_SQL_2()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2"

xWhere = " where SSIUSRNAt = '$'"

X = Trim(txtSelect_Options_1_SSIUSRUIDX)
If X <> "" Then xWhere = xWhere & " and SSIUSRUIDX like '" & X & "%'"

Select Case Mid$(cboSelect_Options_1_SSIUSRSTAK, 1, 1)
    Case " ": xWhere = xWhere & " and SSIUSRSTAK = ' '"
    Case "N": xWhere = xWhere & " and SSIUSRSTAK = 'N'"
End Select

X = Trim(cboSelect_Options_1_SSIDOMDIDX)
If X = "" And Trim(cboSelect_Options_1_SSIDOMPRFX) = "" Then
    blnSelect_Options_1_SSIDOMPRFX = False
    
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
         & xWhere _
         & " order by SSIUSRUIDX"
Else
    xWhere = xWhere & " and SSIDOMUIDN = SSIUSRUIDN "
    If X <> "" Then xWhere = xWhere & " and SSIDOMDIDX = '" & X & "'"
    X = Trim(cboSelect_Options_1_SSIDOMPRFX)
    If X <> "" Then
        blnSelect_Options_1_SSIDOMPRFX = True
        If Mid$(X, 1, 1) <> "*" Then
            xWhere = xWhere & " and SSIDOMPRFX = '" & X & "'"
        End If
    End If
    Select Case Mid$(cboSelect_Options_1_SSIUSRSTAK, 1, 1)
        Case " ": xWhere = xWhere & " and SSIDOMSTAK = ' '"
        Case "N": xWhere = xWhere & " and SSIDOMSTAK = 'N'"
    End Select
    
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0, " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
         & xWhere _
         & " order by SSIUSRUIDX , SSIDOMDIDX, SSIDOMPRFX , SSIDOMUIDX"
End If

Set rsSab = cnsab.Execute(xSQL)
   

fgSelect_Display_1

Set rsSab = Nothing
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_2_S()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2_S"

xWhere = " where SSIUSRNAt = 'S'"

    
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
     & xWhere _
     & " order by SSIUSRUIDN"

Set rsSab = cnsab.Execute(xSQL)
   

fgSelect_Display_1

Set rsSab = Nothing
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub cmdSelect_SQL_2_D()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler


currentAction = "cmdSelect_SQL_2_D"
Call cmdProfil_Quit_Click
cboProfil_DOM.Locked = False
cmdProfil_Excel.BackColor = &H80FFFF
fraProfil.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub cmdSelect_SQL_9_SSIIBMPRFK()
Dim V, X As String, K As Integer, XX As String
Dim xSQL As String, xWhere As String

Dim rsSab As New ADODB.Recordset

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_9_SSIIBMPRFK"

Call cmdUpdate_Init
Call rsYSSIIBM0_Init(prfYSSIIBM0)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0, " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & " where SSIIBMPRFK = '#'  And SSIDOMNAT = ' ' And SSIDOMDIDX = 'IBM' and SSIIBMUIDD = SSIDOMUIDD" _
     & " and SSIUSRNAT = ' ' and SSIUSRUIDN = SSIDOMUIDN " _
     & " order by SSIDOMPRFX"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
    newYSSIDOM0 = oldYSSIDOM0
    
    Call rsYSSIIBM0_GetBuffer(rsSab, oldYSSIIBM0)
    newYSSIIBM0 = oldYSSIIBM0
    Call rsYSSIUSR0_GetBuffer(rsSab, oldYSSIUSR0)
    newYSSIUSR0 = oldYSSIUSR0
    
    X = rsSab("SSIDOMPRFX")
    newYSSIIBM0.SSIIBMPRFK = cmdSelect_SQL_9_SSIIBMPRFK_Control(X)
    
    newYSSIIBM0.SSIIBMYFCT = "CTL"
    newYSSIIBM0.SSIIBMYUSR = usrName_UCase
    newYSSIIBM0.SSIIBMYAMJ = DSys
    newYSSIIBM0.SSIIBMYHMS = time_Hms
    
    newYSSIDOM0.SSIDOMPRFK = newYSSIIBM0.SSIIBMPRFK
    newYSSIDOM0.SSIDOMPRFD = newYSSIIBM0.SSIIBMYAMJ
    newYSSIDOM0.SSIDOMPRFH = newYSSIIBM0.SSIIBMYHMS
   
    If newYSSIDOM0.SSIDOMPRFK = " " Then
' conforme / profil
        If oldYSSIDOM0.SSIDOMPRFK = " " Then
            mYSSIDOM0_Update = "Update": mYSSIIBM0_Update = "Update"
        Else
            newYSSIDOM0.SSIDOMYFCT = "CTL"
            mYSSIDOM0_Update = "Update+H": mYSSIIBM0_Update = "Update"
        End If
    
    Else
' NON conforme / profil
        newYSSIDOM0.SSIDOMYFCT = "CTL"
        mYSSIDOM0_Update = "Update+H": mYSSIIBM0_Update = "Update"
    
    End If
    'If oldYSSIDOM0.SSIDOMSTAK = "N" Then
        If Trim(oldYSSIIBM0.UPSTAT) = "*DISABLED" _
       And Trim(oldYSSIIBM0.UPINMN) = "*SIGNOFF" _
       And Trim(oldYSSIIBM0.UPINPG) = "*NONE" Then '_
       'And Trim(oldYSSIIBM0.UPGRPF) = "EXIT_GRP" Then
            newYSSIDOM0.SSIDOMPRFK = "X"
            newYSSIIBM0.SSIIBMPRFK = "X"
            newYSSIDOM0.SSIDOMYFCT = "CTL"
            mYSSIDOM0_Update = "Update+H": mYSSIIBM0_Update = "Update"
      End If
    'End If

    '''Call cmdSSIJRN_IBM("<X:importation du compte modifié (Historique)>")
    Call cmdSSIJRN_DOM("<X:voir Historique : IBM " & newYSSIIBM0.UPUPRF & ">")
    Call cmdUpdate
    
    rsSab.MoveNext
Loop

Set rsSab = Nothing

'________________________________________________________________
Call cmdSelect_SQL_9_SSIIBMPRFK_Inactif
'________________________________________________________________

Exit Sub

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_9_SAB_ZMNUUTI0()
Dim V, X As String, K As Integer, XX As String
Dim xSQL As String, xWhere As String, blnP_Min As Boolean

Dim rsSab As New ADODB.Recordset

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_9_SAB_ZMNUUTI0"

Call cmdUpdate_Init
Call rsYSSISAB0_Init(prfYSSISAB0)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0, " & paramIBM_Library_SABSPE & ".YSSIDOM0," _
     & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & " where SSISABNAT = ' ' And SSISABPRFK = '#' And SSIDOMDIDX = 'SAB' and SSISABUIDD = SSIDOMUIDD" _
     & " and SSIUSRNAT = ' ' and SSIUSRUIDN = SSIDOMUIDN" _
     & " order by SSIDOMPRFX"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
    newYSSIDOM0 = oldYSSIDOM0
    
    Call rsYSSISAB0_GetBuffer(rsSab, oldYSSISAB0)
    newYSSISAB0 = oldYSSISAB0
    
    Call rsYSSIUSR0_GetBuffer(rsSab, oldYSSIUSR0)
    newYSSIUSR0 = oldYSSIUSR0
    
    If oldYSSISAB0.SSISABPRFX = "P_MIN" Then
        newYSSISAB0.SSISABPRFK = "X"
        
        newYSSIDOM0.SSIDOMPRFX = "P_MIN"
        newYSSIDOM0.SSIDOMPRFK = "X"
        newYSSIDOM0.SSIDOMSTAK = "N"
        mYSSIDOM0_Update = "Update+H": mYSSISAB0_Update = "Update"
    Else
        If oldYSSISAB0.SSISABPRFX = oldYSSIDOM0.SSIDOMPRFX Then
             ' conforme / profil
            newYSSISAB0.SSISABPRFK = " "
            newYSSIDOM0.SSIDOMPRFK = ""
            If oldYSSIDOM0.SSIDOMPRFK = " " Then
                mYSSIDOM0_Update = "Update": mYSSISAB0_Update = "Update"
            Else
                mYSSIDOM0_Update = "Update+H": mYSSISAB0_Update = "Update"
            End If
       Else
             ' NON conforme / profil
            newYSSISAB0.SSISABPRFK = "N"
            newYSSIDOM0.SSIDOMPRFK = "N"
            mYSSIDOM0_Update = "Update+H": mYSSISAB0_Update = "Update"
        End If
     End If
     
    newYSSISAB0.SSISABYFCT = "CTL"
    newYSSISAB0.SSISABYUSR = usrName_UCase
    newYSSISAB0.SSISABYAMJ = DSys
    newYSSISAB0.SSISABYHMS = time_Hms

    newYSSIDOM0.SSIDOMYFCT = "CTL"
    newYSSIDOM0.SSIDOMYUSR = usrName_UCase
    newYSSIDOM0.SSIDOMPRFD = newYSSISAB0.SSISABYAMJ
    newYSSIDOM0.SSIDOMPRFH = newYSSISAB0.SSISABYHMS
   

    Call cmdSSIJRN_DOM("")
    Call cmdUpdate
    
    rsSab.MoveNext
Loop

Set rsSab = Nothing

'________________________________________________________________
'________________________________________________________________

Exit Sub

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_9_SSIIBMPRFK_Inactif()
Dim V, X As String, K As Integer, XX As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

Dim rsSab As New ADODB.Recordset

currentAction = "cmdSelect_SQL_9_SSIIBMPRFK_Inactif"

Call cmdUpdate_Init
    


Call rsYSSIIBM0_Init(xYSSIIBM0)
'______________________________________________________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIIBM0" _
     & " where SSIDOMSTAK = 'N' and SSIDOMPRFK = ' ' And UPSTAT = '*ENABLED' and SSIIBMUIDD = SSIDOMUIDD" _
     & " order by SSIDOMPRFX"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
    newYSSIDOM0 = oldYSSIDOM0
    
    newYSSIDOM0.SSIDOMPRFK = "!"
    Call cmdSSIJRN_DOM("")
    mYSSIDOM0_Update = "Update"
    Call cmdUpdate
    
    rsSab.MoveNext
Loop
'______________________________________________________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIIBM0" _
     & " where SSIDOMDECH > 0 and  SSIDOMDECH < " & DSys & " and SSIDOMPRFK = ' ' And UPSTAT = '*ENABLED' and SSIIBMUIDD = SSIDOMUIDD" _
     & " order by SSIDOMPRFX"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
    newYSSIDOM0 = oldYSSIDOM0
    
    newYSSIDOM0.SSIDOMPRFK = "!"
    Call cmdSSIJRN_DOM("")
    mYSSIDOM0_Update = "Update"
    Call cmdUpdate
    
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


Private Sub cmdSelect_SQL_1()
Dim V, X As String, K As Integer, XX As String
Dim xSQL As String, xWhere As String, blnSTAK_Oui As Boolean
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"

xWhere = " where SSIUSRNAT = ' '"

X = Trim(txtSelect_Options_1_SSIUSRUIDX)
If X <> "" Then xWhere = xWhere & " and SSIUSRUIDX like '" & X & "%'"

Select Case Mid$(cboSelect_Options_1_SSIUSRSTAK, 1, 1)
    Case " ": xWhere = xWhere & " and SSIUSRSTAK = ' '": blnSTAK_Oui = True
    Case "N": xWhere = xWhere & " and SSIUSRSTAK = 'N'"
End Select

X = Trim(cboSelect_Options_1_SSIDOMDIDX)
If X = "" And Trim(cboSelect_Options_1_SSIDOMPRFX) = "" And Trim(txtSelect_Options_1_SSIDOMUIDX) = "" Then
    blnSelect_Options_1_SSIDOMPRFX = False
    
    'xWhere = Replace(xWhere, "and", "where", 1, 1)
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
         & xWhere _
         & " order by SSIUSRUIDX"
Else
    xWhere = xWhere & " and SSIDOMUIDN = SSIUSRUIDN "
    If X <> "" Then xWhere = xWhere & " and SSIDOMDIDX = '" & X & "'"
    X = Trim(cboSelect_Options_1_SSIDOMPRFX)
    If X <> "" Then
        blnSelect_Options_1_SSIDOMPRFX = True
        If Mid$(X, 1, 1) <> "*" Then
            xWhere = xWhere & " and SSIDOMPRFX = '" & X & "'"
        End If
    End If
    X = Trim(txtSelect_Options_1_SSIDOMUIDX)
    If X <> "" Then xWhere = xWhere & " and SSIDOMUIDX like '%" & X & "%'"
    Select Case Mid$(cboSelect_Options_1_SSIUSRSTAK, 1, 1)
        Case " ": xWhere = xWhere & " and SSIDOMSTAK = ' '": blnSTAK_Oui = True
        Case "N": xWhere = xWhere & " and SSIDOMSTAK = 'N'"
    End Select
    
    
    'xWhere = Replace(xWhere, "and", "where", 1, 1)
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0, " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
         & xWhere _
         & " order by SSIUSRUIDX , SSIDOMDIDX, SSIDOMUIDX"
End If

Set rsSab = cnsab.Execute(xSQL)

fgSelect_Display_1
If fgSelect.Rows = 1 And blnSTAK_Oui Then
    xSQL = Replace(xSQL, "and SSIUSRSTAK = ' '", "")
    Set rsSab = cnsab.Execute(xSQL)
    fgSelect_Display_1
End If

If fgSelect.Rows = 2 Then fgSelect.Row = 1: fgSelect_Row_Click
Set rsSab = Nothing

Exit Sub

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub



Private Sub paramIBM_Init()
Dim V, X As String, xIn As String
Dim xSQL As String, xWhere As String, Nb As Long, K As Integer
Dim xUIDD As String, wPRFX As String, xALPHA As String, wUPUID As Long
On Error GoTo Error_Handler


GoTo Import_IBM

Exit_GRP:


Open "C:\TEMP\BIA_SSI\BIA_SSI_IBM_LMTCPB.txt" For Output As 3
X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIIBM0" _
  & " where substring( SSIDOMPRFX , 1 , 3 ) = 'COB'" _
  & " and UPSPAU like '%SPLCTL%' and SSIIBMUIDD = SSIDOMUIDD order by UPUPRF"
Set rsSab = cnsab.Execute(X)
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    
    Print #3, "CHGUSRPRF USRPRF(" & Trim(rsSab("UPUPRF")) & ") SPCAUT(*NONE)"
    rsSab.MoveNext
Loop

Close 3


Exit Sub

Open "C:\TEMP\BIA_SSI\BIA_SSI_IBM_PWD.txt" For Output As 3
X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIIBM0" _
  & " where substring( SSIDOMPRFX , 1 , 3 ) = 'SAB'" _
  & " and UPPWEI <> 35 and UPPWEI >= 0 and SSIIBMUIDD = SSIDOMUIDD order by UPUPRF"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    
    Print #3, "CHGUSRPRF USRPRF(" & Trim(rsSab("UPUPRF")) & ") PWDEXPITV(35)"
    rsSab.MoveNext
Loop

Close 3

Open "C:\TEMP\BIA_SSI\BIA_SSI_IBM_LMTCPB.txt" For Output As 3
X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0 where UPLTCP = '*PARTIAL' order by UPUPRF"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    
    Print #3, "CHGUSRPRF USRPRF(" & Trim(rsSab("UPUPRF")) & ") LMTCPB(*YES)"
    rsSab.MoveNext
Loop

Close 3


Exit Sub
'====================================================================

Open "C:\TEMP\BIA_SSI\BIA_SSI_IBM_Exit_GRP.txt" For Output As 3
X = "select * from " & paramIBM_Library_SABSPE & ".YSSIdom0 where SSIdomstak = 'N' order by ssidomuidn"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    
    Print #3, "CALL       PGM(BIASABOBJ/S_EXIT_GRP) PARM('" & Trim(rsSab("SSIDOMUIDX")) & "')"
    rsSab.MoveNext
Loop

Close 3
 

'CHGUSRPRF USRPRF(JUNIER) PWDEXPITV(35)
'CHGUSRPRF USRPRF(T_BELHADDA) LMTCPB(*YES)

Exit Sub
'====================================================================

Export:

Open "C:\TEMP\BIA_SSI\BIA_SSI_IBM.txt" For Output As 3
X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0 where SSIIBMNAT = ' ' order by UPUPRF"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    xUIDD = "*"
    xALPHA = rsSab("UPUPRF")
    If Mid$(xALPHA, 2, 1) = "_" Then xALPHA = Mid$(xALPHA & "        ", 3, 7)
    
    Select Case Trim(rsSab("UPINPG"))
        Case "SAB073": wPRFX = "SAB_PROD"
        Case "SAB073U": wPRFX = "SAB_TEST"
        Case Else: wPRFX = Trim(rsSab("UPINPG"))
    End Select
    
     Select Case Mid$(rsSab("UPUPRF"), 1, 2)
        Case "B_": wPRFX = "COBANQUE"
        Case "T_": wPRFX = "SAB_TEST"
        Case "P_": wPRFX = "INFO_PROD"
        Case "I_": wPRFX = "INFO_TEST"
        Case "X_": wPRFX = "SAB_GAP_P"
    End Select
   
    If Mid$(xALPHA, 1, 1) = "Q" Then xUIDD = "1001": wPRFX = "*NONE"
    
    xALPHA = Replace(xALPHA, "_CDO", "")
    xALPHA = Replace(xALPHA, "_B", "")
    xALPHA = Replace(xALPHA, "_C", "")
    xALPHA = Replace(xALPHA, "_TC", "")
    xALPHA = Replace(xALPHA, "_K", "")
    xALPHA = Replace(xALPHA, "_D", "")
    xALPHA = Replace(xALPHA, "_SO", "")
    
    Print #3, xUIDD & ";" & rsSab("SSIIBMUIDD") & ";" & rsSab("UPUPRF") & "; " & Trim(xALPHA) & "x; " & wPRFX; ""
    rsSab.MoveNext
Loop

Close 3
 


Exit Sub
'====================================================================
SSIIBMDECH:


Open "C:\TEMP\BIA_SSI\BIA_SSI_IBM V0.txt" For Input As 1
Open "C:\TEMP\BIA_SSI\BIA_SSI_IBM V1.txt" For Output As 3


Do Until EOF(1)
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        K = 0
        X = CSV_Scan(xIn, K)
        wUPUID = Val(CSV_Scan(xIn, K))
        xWhere = Trim(CSV_Scan(xIn, K))
        xALPHA = CSV_Scan(xIn, K)
        wPRFX = Trim(CSV_Scan(xIn, K))
        If X = "*" And wPRFX = "SAB_PROD" Then
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0 where SSIIBMNAT = '' and SSIIBMUIDD = " & wUPUID
            Set rsSab = cnsab.Execute(X)
            Nb = Val(rsSab("UPPSOC") & rsSab("UPPSOD")) + 19000000
            If Nb < 20120701 Then xIn = Replace(xIn, "*", "*" & Nb)
        End If
    End If
    Print #3, xIn
Loop

Close
Exit Sub
'====================================================================

Import_IBM:

currentAction = "paramIBM_Init"


Call rsYSSIUSR0_Init(newYSSIUSR0)

newYSSIUSR0.SSIUSRYFCT = "INI"
newYSSIUSR0.SSIUSRYAMJ = DSys
newYSSIUSR0.SSIUSRYHMS = time_Hms
newYSSIUSR0.SSIUSRYUSR = usrName_UCase

Call rsYSSIIBM0_Init(newYSSIIBM0)
newYSSIIBM0.SSIIBMNAT = "$"
newYSSIIBM0.SSIIBMYAMJ = DSys
newYSSIIBM0.SSIIBMYHMS = time_Hms
newYSSIIBM0.SSIIBMYUSR = usrName_UCase


Call rsYSSIDOM0_Init(newYSSIDOM0)
newYSSIDOM0.SSIDOMDIDX = "IBM"
newYSSIDOM0.SSIDOMYFCT = "INI"
newYSSIDOM0.SSIDOMYAMJ = DSys
newYSSIDOM0.SSIDOMYHMS = time_Hms
newYSSIDOM0.SSIDOMYUSR = usrName_UCase

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

'xSQL = "delete from " & paramIBM_Library_SABSPE & ".YSSIUSR0"
'Set rsSab = cnSab_Update.Execute(xSQL, Nb)

'xSQL = "delete from " & paramIBM_Library_SABSPE & ".YSSIUSRH"
'Set rsSab = cnSab_Update.Execute(xSQL, Nb)

'xSQL = "delete from " & paramIBM_Library_SABSPE & ".YSSIDOM0"
'Set rsSab = cnSab_Update.Execute(xSQL, Nb)
'xSQL = "delete from " & paramIBM_Library_SABSPE & ".YSSIDOMH"
'Set rsSab = cnSab_Update.Execute(xSQL, Nb)

'xSQL = "delete from " & paramIBM_Library_SABSPE & ".YSSIIBM0 where SSIIBMNAT = '$'"
'Set rsSab = cnSab_Update.Execute(xSQL, Nb)
'xSQL = "delete from " & paramIBM_Library_SABSPE & ".YSSIIBMH"
'Set rsSab = cnSab_Update.Execute(xSQL, Nb)

'xSQL = "delete from " & paramIBM_Library_SABSPE & ".YSSITXT0"
'Set rsSab = cnSab_Update.Execute(xSQL, Nb)


Profils:
'____________________________________________________________________

oldYSSIIBM0.SSIIBMUIDD = 0
newYSSIIBM0.SSIIBMUIDD = 1: newYSSIIBM0.UPUPRF = "QSECOFR": newYSSIIBM0.UPTEXT = "IBM : tous pouvoirs"
V = sqlYSSIIBM0_Profil_Insert(newYSSIIBM0, oldYSSIIBM0): If Not IsNull(V) Then GoTo Error_MsgBox

oldYSSIIBM0.SSIIBMUIDD = 102
newYSSIIBM0.SSIIBMUIDD = 2: newYSSIIBM0.UPUPRF = "SYSTEME": newYSSIIBM0.UPTEXT = "IBM : système"
V = sqlYSSIIBM0_Profil_Insert(newYSSIIBM0, oldYSSIIBM0): If Not IsNull(V) Then GoTo Error_MsgBox

oldYSSIIBM0.SSIIBMUIDD = 345
newYSSIIBM0.SSIIBMUIDD = 3: newYSSIIBM0.UPUPRF = "QSYSOPR": newYSSIIBM0.UPTEXT = "IBM : opérateur"
V = sqlYSSIIBM0_Profil_Insert(newYSSIIBM0, oldYSSIIBM0): If Not IsNull(V) Then GoTo Error_MsgBox

oldYSSIIBM0.SSIIBMUIDD = 256 '245
newYSSIIBM0.SSIIBMUIDD = 4: newYSSIIBM0.UPUPRF = "SAB_PROD": newYSSIIBM0.UPTEXT = newYSSIIBM0.UPUPRF
V = sqlYSSIIBM0_Profil_Insert(newYSSIIBM0, oldYSSIIBM0): If Not IsNull(V) Then GoTo Error_MsgBox

oldYSSIIBM0.SSIIBMUIDD = 233
newYSSIIBM0.SSIIBMUIDD = 5: newYSSIIBM0.UPUPRF = "SAB_TEST": newYSSIIBM0.UPTEXT = newYSSIIBM0.UPUPRF
V = sqlYSSIIBM0_Profil_Insert(newYSSIIBM0, oldYSSIIBM0): If Not IsNull(V) Then GoTo Error_MsgBox

oldYSSIIBM0.SSIIBMUIDD = 251
newYSSIIBM0.SSIIBMUIDD = 6: newYSSIIBM0.UPUPRF = "INFO_PROD": newYSSIIBM0.UPTEXT = newYSSIIBM0.UPUPRF
V = sqlYSSIIBM0_Profil_Insert(newYSSIIBM0, oldYSSIIBM0): If Not IsNull(V) Then GoTo Error_MsgBox

oldYSSIIBM0.SSIIBMUIDD = 145
newYSSIIBM0.SSIIBMUIDD = 7: newYSSIIBM0.UPUPRF = "INFO_TEST": newYSSIIBM0.UPTEXT = newYSSIIBM0.UPUPRF
V = sqlYSSIIBM0_Profil_Insert(newYSSIIBM0, oldYSSIIBM0): If Not IsNull(V) Then GoTo Error_MsgBox

oldYSSIIBM0.SSIIBMUIDD = 567
newYSSIIBM0.SSIIBMUIDD = 8: newYSSIIBM0.UPUPRF = "SAB_GAP_P": newYSSIIBM0.UPTEXT = newYSSIIBM0.UPUPRF
V = sqlYSSIIBM0_Profil_Insert(newYSSIIBM0, oldYSSIIBM0): If Not IsNull(V) Then GoTo Error_MsgBox

oldYSSIIBM0.SSIIBMUIDD = 571
newYSSIIBM0.SSIIBMUIDD = 9: newYSSIIBM0.UPUPRF = "SAB_GAP_T": newYSSIIBM0.UPTEXT = newYSSIIBM0.UPUPRF
V = sqlYSSIIBM0_Profil_Insert(newYSSIIBM0, oldYSSIIBM0): If Not IsNull(V) Then GoTo Error_MsgBox

oldYSSIIBM0.SSIIBMUIDD = 514
newYSSIIBM0.SSIIBMUIDD = 10: newYSSIIBM0.UPUPRF = "COBANQUE": newYSSIIBM0.UPTEXT = newYSSIIBM0.UPUPRF
V = sqlYSSIIBM0_Profil_Insert(newYSSIIBM0, oldYSSIIBM0): If Not IsNull(V) Then GoTo Error_MsgBox

oldYSSIIBM0.SSIIBMUIDD = 122
newYSSIIBM0.SSIIBMUIDD = 11: newYSSIIBM0.UPUPRF = "SABTELEM": newYSSIIBM0.UPTEXT = newYSSIIBM0.UPUPRF
V = sqlYSSIIBM0_Profil_Insert(newYSSIIBM0, oldYSSIIBM0): If Not IsNull(V) Then GoTo Error_MsgBox

'Utilisateurs spéciaux
'____________________________________________________________________

newYSSIUSR0.SSIUSRUIDN = 1000

Open "C:\TEMP\BIA_SSI\BIA_SSI_IBM V1.txt" For Input As 1


Do Until EOF(1)
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        K = 0
        X = CSV_Scan(xIn, K)
        wUPUID = Val(CSV_Scan(xIn, K))
        xWhere = Trim(CSV_Scan(xIn, K))
        xALPHA = CSV_Scan(xIn, K)
        wPRFX = Trim(CSV_Scan(xIn, K))
        If Mid$(X, 1, 1) = "*" Then
            Nb = 0
            newYSSIUSR0.SSIUSRUIDN = newYSSIUSR0.SSIUSRUIDN + 1
            newYSSIUSR0.SSIUSRUIDX = xWhere
            If Len(X) = 9 Then
                newYSSIUSR0.SSIUSRSTAK = "N"
                newYSSIUSR0.SSIUSRDECH = Mid$(X, 2, 8)
            Else
                newYSSIUSR0.SSIUSRSTAK = " "
                newYSSIUSR0.SSIUSRDECH = 0
            End If
            
            V = sqlYSSIUSR0_Insert(newYSSIUSR0): If Not IsNull(V) Then GoTo Error_MsgBox
        End If
        Nb = Nb + 1
        If Not IsNull(paramIBM_Init_Compte(newYSSIUSR0.SSIUSRUIDN, wPRFX, wUPUID)) Then GoTo Error_MsgBox
    End If
Loop

Close


'____________________________________________________________________


'____________________________________________________________________



If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        

    End If
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Set rsSab = Nothing

Exit Sub


End Sub
Private Sub paramSAA_Init()
Dim V, X As String, xIn As String
Dim xSQL As String, xWhere As String, Nb As Long, K As Integer
Dim xUIDD As String, wPRFX As String, xALPHA As String, wUPUID As Long
On Error GoTo Error_Handler


Import_SAA:

currentAction = "paramSAA_Init"
Call cmdUpdate_Init

Call rsYSSISAA0_Init(newYSSISAA0)
newYSSISAA0.SSISAANAT = "$"
newYSSISAA0.SSISAAYAMJ = DSys
newYSSISAA0.SSISAAYHMS = time_Hms
newYSSISAA0.SSISAAYUSR = usrName_UCase


Call rsYSSIDOM0_Init(newYSSIDOM0)
newYSSIDOM0.SSIDOMDIDX = "SAA"
newYSSIDOM0.SSIDOMYAMJ = DSys
newYSSIDOM0.SSIDOMYHMS = time_Hms
newYSSIDOM0.SSIDOMYUSR = usrName_UCase
'_________________________________________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0, " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
     & " where SSISAANAT = ' ' and SSISAAPRFK = '?' " _
     & " and SSIDOMUIDX = SSISAAUIDX and SSIDOMNAt = ' ' and SSIDOMPRFX = 'SAB_PROD' order by SSISAAUIDX"

Set rsSab = cnsab.Execute(xSQL)

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________



Do While Not rsSab.EOF
    Call rsYSSISAA0_GetBuffer(rsSab, oldYSSISAA0)
    newYSSISAA0 = oldYSSISAA0
    newYSSISAA0.SSISAAPRFK = " "
    mYSSISAA0_Update = "Update"

    Call rsYSSIDOM0_GetBuffer(rsSab, newYSSIDOM0)
    mYSSIDOM0_Update = "New"
    'UIDN
    newYSSIDOM0.SSIDOMDIDX = "SAA"
   ' If Trim(oldYSSISAA0.SSISAAUIDX) = "TAN" Or Trim(oldYSSISAA0.SSISAAUIDX) = "ZABAY" Then
   '     newYSSIDOM0.SSIDOMDIDS = 1
   ' Else
   '     newYSSIDOM0.SSIDOMDIDS = 0
   ' End If
    newYSSIDOM0.SSIDOMSTAK = oldYSSISAA0.SSISAASTAK
    newYSSIDOM0.SSIDOMDECH = 0
    newYSSIDOM0.SSIDOMUIDD = oldYSSISAA0.SSISAAUIDD
    newYSSIDOM0.SSIDOMUIDX = oldYSSISAA0.SSISAAUIDX
    newYSSIDOM0.SSIDOMPRFX = oldYSSISAA0.SSISAAPRFX
    If oldYSSISAA0.SSISAASTAK = " " Then
        newYSSIDOM0.SSIDOMPRFK = " "
    Else
        newYSSIDOM0.SSIDOMPRFK = "X"
    End If
    newYSSIDOM0.SSIDOMPRFD = oldYSSISAA0.SSISAAYAMJ
    newYSSIDOM0.SSIDOMPRFH = oldYSSISAA0.SSISAAYHMS
    newYSSIDOM0.SSIDOMTLNK = 0
    newYSSIDOM0.SSIDOMYFCT = "INI"
    newYSSIDOM0.SSIDOMYAMJ = DSys
    newYSSIDOM0.SSIDOMYHMS = time_Hms
    newYSSIDOM0.SSIDOMYUSR = usrName_UCase
    newYSSIDOM0.SSIDOMYVER = 0
    
   V = sqlYSSISAA0_Update(newYSSISAA0, oldYSSISAA0)
    If Not IsNull(V) Then GoTo Error_MsgBox
   V = sqlYSSIDOM0_Insert(newYSSIDOM0)
    If Not IsNull(V) Then GoTo Error_MsgBox

    rsSab.MoveNext
Loop

'____________________________________________________________________
Call rsYSSIUSR0_Init(newYSSIUSR0)

newYSSIUSR0.SSIUSRYFCT = "INI"
newYSSIUSR0.SSIUSRYAMJ = DSys
newYSSIUSR0.SSIUSRYHMS = time_Hms
newYSSIUSR0.SSIUSRYUSR = usrName_UCase

'____________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0" _
     & " where SSISAANAT = ' ' and SSISAAPRFK = '?' " _
     & " order by SSISAAUIDX"

Set rsSab = cnsab.Execute(xSQL)


Do While Not rsSab.EOF
    Call rsYSSISAA0_GetBuffer(rsSab, oldYSSISAA0)
    newYSSISAA0 = oldYSSISAA0
    newYSSISAA0.SSISAAPRFK = " "
    mYSSISAA0_Update = "Update"
    'newYSSIDOM0.SSIDOMDIDS = 0
    
    Select Case Trim(oldYSSISAA0.SSISAAUIDX)
        Case "interim1": X = "INTERIM1"
        Case "interim2": X = "INTERIM2"
        Case "interim3": X = "INTERIM3"
        Case "interim4": X = "INTERIM4"
        Case "interim6": X = "INTERIM6"
        Case "ALLARDCLE": X = "ALLARD" '                   cmdProfil_Histo.Visible = False = 1
        Case "BOUDROUAZ": X = "BOUDROUAZ"
        Case "DAOUDTC": X = "DAOUD" '                   cmdProfil_Histo.Visible = False = 1
        Case "DARMON_K": X = "DARMON" '                   cmdProfil_Histo.Visible = False = 1
        Case "DUCHESNE2": X = "DUCHESNE" '                   cmdProfil_Histo.Visible = False = 1
        Case "FEBVRE": X = "FEBVRE_JP"
        Case "FEBVRE2": X = "FEBVRE_JP" '                   cmdProfil_Histo.Visible = False = 1
        Case "FONTANA": X = "FONTANA"
        Case "GRACHEHATC": X = "GRACHEHA" '                   cmdProfil_Histo.Visible = False = 1
        Case "RATSITOARIVONY": X = "RATSITOARI"
        Case "REOL": X = "REOL_CH"
        Case "SUPER": X = "ADMIN"
        Case "SUPER1": X = "ADMIN" '                   cmdProfil_Histo.Visible = False = 1
        Case "VAISSIER3": X = "VAISSIER" '                   cmdProfil_Histo.Visible = False = 1
        Case "BENHAMOU": X = "BENHAMOU"
        Case "BOUDOUAZ": X = "BOUDOUAZ"
        Case "CHRYSANT1": X = "CHRYSANTOS"
        Case "CHRYSANT2": X = "CHRYSANTOS" '                   cmdProfil_Histo.Visible = False = 1
        Case "FERRIEREBO": X = "FERRIERE" '                   cmdProfil_Histo.Visible = False = 1
        Case "HAFIZ": X = "HAFIZ"
        Case "HELIOUI": X = "HELIOUI"
        Case "HOSTINGUE": X = "HOSTINGUE"
        Case "HRABI": X = "HRABI"
        Case "LECCIA": X = "LECCIA"
        Case "MARCOT": X = "MARCOT"
        Case "MERABIA": X = "MERABIA"
        Case "SHALLOUF": X = "SHALLOUF"
        Case "TERF": X = "TERF"
        Case Else: X = "???"
    End Select

     xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
         & " where SSIUSRNAT = ' ' and SSIUSRUIDX = '" & X & "'"
    
    Set rsSab_X = cnsab.Execute(xSQL)
    If Not rsSab_X.EOF Then
        mYSSIDOM0_Update = "New"
        newYSSIDOM0.SSIDOMUIDN = rsSab_X("SSIUSRUIDN")
        newYSSIDOM0.SSIDOMDIDX = "SAA"
        newYSSIDOM0.SSIDOMSTAK = oldYSSISAA0.SSISAASTAK
        newYSSIDOM0.SSIDOMDECH = 0
        newYSSIDOM0.SSIDOMUIDD = oldYSSISAA0.SSISAAUIDD
        newYSSIDOM0.SSIDOMUIDX = oldYSSISAA0.SSISAAUIDX
        newYSSIDOM0.SSIDOMPRFX = oldYSSISAA0.SSISAAPRFX
        If oldYSSISAA0.SSISAASTAK = " " Then
            newYSSIDOM0.SSIDOMPRFK = " "
        Else
            newYSSIDOM0.SSIDOMPRFK = "X"
        End If
        newYSSIDOM0.SSIDOMPRFD = oldYSSISAA0.SSISAAYAMJ
        newYSSIDOM0.SSIDOMPRFH = oldYSSISAA0.SSISAAYHMS
        newYSSIDOM0.SSIDOMTLNK = 0
        newYSSIDOM0.SSIDOMYFCT = "INI"
        newYSSIDOM0.SSIDOMYAMJ = DSys
        newYSSIDOM0.SSIDOMYHMS = time_Hms
        newYSSIDOM0.SSIDOMYUSR = usrName_UCase
        newYSSIDOM0.SSIDOMYVER = 0
    
        V = sqlYSSISAA0_Update(newYSSISAA0, oldYSSISAA0)
         If Not IsNull(V) Then GoTo Error_MsgBox
        V = sqlYSSIDOM0_Insert(newYSSIDOM0)
         If Not IsNull(V) Then GoTo Error_MsgBox
    End If

    rsSab.MoveNext
Loop


'________________________________________________________________________________


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        

    End If
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Set rsSab = Nothing

Exit Sub


End Sub
Private Sub paramWIN_Init()
Dim V, X As String, xIn As String, blnOk As Boolean
Dim xSQL As String, xWhere As String, Nb As Long, K As Integer
Dim xUIDD As String, wPRFX As String, xALPHA As String, wUPUID As Long
On Error GoTo Error_Handler


Import_WIN:

currentAction = "paramWIN_Init"
Call cmdUpdate_Init

Call rsYSSIWIN0_Init(newYSSIWIN0)
newYSSIWIN0.SSIWINNAT = "$"
newYSSIWIN0.SSIWINYAMJ = DSys
newYSSIWIN0.SSIWINYHMS = time_Hms
newYSSIWIN0.SSIWINYUSR = usrName_UCase


Call rsYSSIDOM0_Init(newYSSIDOM0)
newYSSIDOM0.SSIDOMDIDX = "WIN"
newYSSIDOM0.SSIDOMYAMJ = DSys
newYSSIDOM0.SSIDOMYHMS = time_Hms
newYSSIDOM0.SSIDOMYUSR = usrName_UCase
'_________________________________________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0, " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
     & " where SSIWINNAT = ' ' and SSIWINPRFK = '?' " _
     & " and SSIDOMUIDX = SSIWINUIDX and SSIDOMNAT = ' ' and SSIDOMPRFX = 'SAB_PROD' order by SSIWINUIDX"

Set rsSab = cnsab.Execute(xSQL)

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

Do While Not rsSab.EOF
    Call rsYSSIWIN0_GetBuffer(rsSab, oldYSSIWIN0)
    newYSSIWIN0 = oldYSSIWIN0
    newYSSIWIN0.SSIWINPRFK = " "
    mYSSIWIN0_Update = "Update"

    Call rsYSSIDOM0_GetBuffer(rsSab, newYSSIDOM0)
    mYSSIDOM0_Update = "New"
    newYSSIDOM0.SSIDOMDIDX = "WIN"
    newYSSIDOM0.SSIDOMSTAK = oldYSSIWIN0.SSIWINSTAK
    newYSSIDOM0.SSIDOMDECH = 0
    newYSSIDOM0.SSIDOMUIDD = oldYSSIWIN0.SSIWINUIDD
    newYSSIDOM0.SSIDOMUIDX = oldYSSIWIN0.SSIWINUIDX
    newYSSIDOM0.SSIDOMPRFX = oldYSSIWIN0.SSIWINPRFX
    newYSSIDOM0.SSIDOMPRFK = " "
    newYSSIDOM0.SSIDOMPRFD = oldYSSIWIN0.SSIWINYAMJ
    newYSSIDOM0.SSIDOMPRFH = oldYSSIWIN0.SSIWINYHMS
    newYSSIDOM0.SSIDOMTLNK = 0
    newYSSIDOM0.SSIDOMYFCT = "INI"
    newYSSIDOM0.SSIDOMYAMJ = DSys
    newYSSIDOM0.SSIDOMYHMS = time_Hms
    newYSSIDOM0.SSIDOMYUSR = usrName_UCase
    newYSSIDOM0.SSIDOMYVER = 0
    
    V = sqlYSSIWIN0_Update(newYSSIWIN0, oldYSSIWIN0)
    If Not IsNull(V) Then GoTo Error_MsgBox
    V = sqlYSSIDOM0_Insert(newYSSIDOM0)
    If Not IsNull(V) Then GoTo Error_MsgBox

    rsSab.MoveNext
Loop

'____________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
     & " where SSIWINNAT = ' ' and SSIWINPRFK = '?' "

Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    'blnOk = False
    X = Trim(rsSab("SSIWINUIDX"))
    If X = "_BEN-MALEK.perso" Then X = "BENMALEK"
    If X = "_FEBVRE_M.Perso" Then X = "FEBVRE_M"
    If X = "_REOL_Ch.Perso" Then X = "REOL_CH"
    
    If Mid$(X, 1, 1) = "_" Then
        blnOk = True
        X = Replace(X, "_", "")
        X = Replace(X, ".Perso", "")
    Else
    If InStr(X, "Scan") > 0 Then X = "BIA_SCAN"
    If InStr(X, "SCAN") > 0 Then X = "BIA_SCAN"
    If InStr(X, "SRV") > 0 Then X = "BIA_SERVEURS"
    End If
        Select Case X
            Case "CARLOTTI2": X = "CARLOTTI": blnOk = True
            Case "LOULERGU": X = "LOULERGUE": blnOk = True
            Case "REOL_Ch": X = "REOL_CH": blnOk = True
            Case "CONSULTANT2": X = "CONSULTANT": blnOk = True
            Case "ARIAS PEREZ": X = "ARIAS": blnOk = True
            Case "BENCHAHDA": X = "FEBVRE_M": blnOk = True
            Case "EFFIO-JAIMES": X = "EFFIO": blnOk = True
            Case "JPLTST": X = "LOULERGUE": blnOk = True
            Case "REOL_Ca": X = "REOL_CA": blnOk = True
            Case "VAISSIER_F": X = "VAISSIER": blnOk = True
            Case "Roberto FERREIRA": X = "FERREIRA": blnOk = True
            Case "DTB": X = "DATABAIL": blnOk = True
            Case "TRADEIN1": X = "BIA_REUTERS": blnOk = True
            Case "Trader_DGA": X = "BIA_REUTERS": blnOk = True
            Case "Trader_TC1": X = "BIA_REUTERS": blnOk = True
            Case "Trader_TC2": X = "BIA_REUTERS": blnOk = True
            Case "RECRUTEMENT": X = "BIA_WINDOWS": blnOk = True
            Case "FORMATION": X = "BIA_WINDOWS": blnOk = True
            Case "Compta_toto": X = "BIA_WINDOWS": blnOk = True
            Case "E-MID": X = "BIA_WINDOWS": blnOk = True
            Case "INFOGEN": X = "BIA_WINDOWS": blnOk = True
            Case "BO": X = "BIA_WINDOWS": blnOk = True
            Case "SUNCHEQUE": X = "BIA_WINDOWS": blnOk = True
            Case "DocushareLDAP": X = "BIA_WINDOWS": blnOk = True
            Case "LDAPRead": X = "BIA_WINDOWS": blnOk = True
            Case "RSA ENVISION": X = "BIA_WINDOWS": blnOk = True
            Case "sms-monitor": X = "BIA_WINDOWS": blnOk = True
            Case "XTIMEMAIL": X = "BIA_WINDOWS": blnOk = True
            Case "Admin": X = "WINDOWS": blnOk = True
            Case "Administrateur": X = "WINDOWS": blnOk = True
            Case "Administrator": X = "WINDOWS": blnOk = True
            Case "anonymous": X = "WINDOWS": blnOk = True
            Case "ARCSERVE": X = "BIA_WINDOWS": blnOk = True
            Case "BRIGHTSTOR": X = "BIA_WINDOWS": blnOk = True
             Case "DB2ADMIN": X = "BIA_WINDOWS": blnOk = True
             Case "SIDE_DB": X = "BIA_WINDOWS": blnOk = True
           
            Case "SUNGARD": X = "SUNGARD": blnOk = True
            Case "DATAREADY": X = "SUNGARD": blnOk = True
            Case "BIA_FTP": X = "BIA": blnOk = True
            Case "BIA_INTRANET": X = "BIA": blnOk = True
            Case "BIAINSTALL": X = "BIA": blnOk = True
            Case "INSTALLIMP": X = "BIA": blnOk = True
            Case "BO": X = "BIA": blnOk = True
            Case "bsabadmin": X = "SAB": blnOk = True
            Case "TESTTN TESTTN": X = "NGUON": blnOk = True
            Case "GUEST": X = "WINDOWS": blnOk = True
            Case "SQLExecutiveCmdExec": X = "BIA_WINDOWS": blnOk = True
            Case "ALLIANCE": X = "BIA_SERVEURS": blnOk = True
            Case "biaxtime": X = "BIA_SERVEURS": blnOk = True
            Case "CORONACS": X = "BIA_SERVEURS": blnOk = True
            Case "EXCHANGE": X = "BIA_SERVEURS": blnOk = True
            Case "INTERPEL2013": X = "BIA_SERVEURS": blnOk = True
            Case "PSP2013": X = "BIA_SERVEURS": blnOk = True
            Case "SIDE2010": X = "BIA_SERVEURS": blnOk = True
            Case "SYSPERTEC": X = "BIA_SERVEURS": blnOk = True
            Case "Antoine Berthier": X = "DATABAIL": blnOk = True
            
        End Select
    'If blnOk Then
    X = Replace(X, "'", "''")
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
             & " where SSIUSRNAT = ' ' and SSIUSRUIDX = '" & X & "'"
        Set rsSab_X = cnsab.Execute(xSQL)
        If Not rsSab_X.EOF Then
            Call rsYSSIWIN0_GetBuffer(rsSab, oldYSSIWIN0)

            newYSSIWIN0 = oldYSSIWIN0
            newYSSIWIN0.SSIWINPRFK = " "
            mYSSIWIN0_Update = "Update"
        
            mYSSIDOM0_Update = "New"
            newYSSIDOM0.SSIDOMNAT = " "
            newYSSIDOM0.SSIDOMSTAK = " "
            newYSSIDOM0.SSIDOMUNIT = " "
            newYSSIDOM0.SSIDOMUIDN = rsSab_X("SSIUSRUIDN")
            newYSSIDOM0.SSIDOMDIDX = "WIN"
            newYSSIDOM0.SSIDOMSTAK = oldYSSIWIN0.SSIWINSTAK
            newYSSIDOM0.SSIDOMDECH = 0
            newYSSIDOM0.SSIDOMUIDD = oldYSSIWIN0.SSIWINUIDD
            newYSSIDOM0.SSIDOMUIDX = oldYSSIWIN0.SSIWINUIDX
            newYSSIDOM0.SSIDOMPRFX = oldYSSIWIN0.SSIWINPRFX
            newYSSIDOM0.SSIDOMPRFK = " "
            newYSSIDOM0.SSIDOMPRFD = oldYSSIWIN0.SSIWINYAMJ
            newYSSIDOM0.SSIDOMPRFH = oldYSSIWIN0.SSIWINYHMS
            newYSSIDOM0.SSIDOMTLNK = 0
            newYSSIDOM0.SSIDOMYFCT = "INI"
            newYSSIDOM0.SSIDOMYAMJ = DSys
            newYSSIDOM0.SSIDOMYHMS = time_Hms
            newYSSIDOM0.SSIDOMYUSR = usrName_UCase
            newYSSIDOM0.SSIDOMYVER = 0
            
            V = sqlYSSIWIN0_Update(newYSSIWIN0, oldYSSIWIN0)
            If Not IsNull(V) Then GoTo Error_MsgBox
            V = sqlYSSIDOM0_Insert(newYSSIDOM0)
            If Not IsNull(V) Then GoTo Error_MsgBox
        'End If
    End If
    rsSab.MoveNext
Loop


'________________________________________________________________________________


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        

    End If
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Set rsSab = Nothing

Exit Sub


End Sub
Private Sub paramDIV_Init_UGM_UIDX()
Dim V, X As String, xIn As String, blnOk As Boolean
Dim xSQL As String, xWhere As String, Nb As Long, K As Integer
Dim xUIDD As String, wPRFX As String, xALPHA As String, wUPUID As Long
On Error GoTo Error_Handler


Import_DIV:

currentAction = "paramDIV_Init_UGM_UIDX"
Call cmdUpdate_Init

Call rsYSSIDIV0_Init(newYSSIDIV0)
newYSSIDIV0.SSIDIVNAT = "$"
newYSSIDIV0.SSIDIVYAMJ = DSys
newYSSIDIV0.SSIDIVYHMS = time_Hms
newYSSIDIV0.SSIDIVYUSR = usrName_UCase


Call rsYSSIDOM0_Init(newYSSIDOM0)
newYSSIDOM0.SSIDOMDIDX = "DIV"
newYSSIDOM0.SSIDOMYAMJ = DSys
newYSSIDOM0.SSIDOMYHMS = time_Hms
newYSSIDOM0.SSIDOMYUSR = usrName_UCase
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________


'____________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0" _
     & " where SSIDIVNAT = ' ' and SSIDIVDIDK = 'UGM' and SSIDIVPRFK = '?' "

Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    'blnOk = False
    X = Trim(rsSab("SSIDIVUIDX"))
        Select Case X
            Case "BASTE-850": X = "BASTE"
            Case "BELBAHI-753": X = "BELBAHI"
            Case "BENMALEK-841": X = "BENMALEK"
            Case "BENMALEK-866": X = "BENMALEK"
            Case "BENTLEY-472": X = "BENTLEY"
            Case "BENTLEY-815": X = "BENTLEY"
            Case "BENTLEY-835": X = "BENTLEY"
            Case "BERCHE-667": X = "BERCHE"
            Case "BERCHE-875": X = "BERCHE"
            Case "CLERC-1982": X = "CLERC"
            Case "CLERC-XXXX": X = "CLERC"
            Case "CULPIN-828": X = "CULPIN"
            Case "DAOUD-1962": X = "DAOUD"
            Case "DJOUZI-248": X = "DJOUZI"
            Case "FEBVRE-472": X = "FEBVRE_M"
            Case "FOUCART-540": X = "FOUCART"
            Case "GRACHEHA-XXXX": X = "GRACHEHA"
            Case "HAJJAR-633": X = "HAJJAR"
            Case "HANDICHI-886": X = "HANDICHI"
            Case "HAUDOYER-853": X = "HAUDOYER"
            Case "JUNIER-587": X = "JUNIER"
            Case "LAGARDE-849": X = "LAGARDE"
            Case "LAGARDEFOTC-XXXX": X = "LAGARDE"
            Case "LECOCQ-547": X = "LECOCQ"
            Case "LECOCQ-862": X = "LECOCQ"
            Case "LEGOUARD-838": X = "LEGOUARD"
            Case "LIGOT-713": X = "LIGOT"
            Case "MENAGE-XXXX": X = "SAMSIC"
            Case "DIOP-XXXX": X = "SAMSIC"
            Case "METIDJI-735": X = "METIDJI"
            Case "METIDJI-863": X = "METIDJI"
            Case "MORICEAU-669": X = "MORICEAU"
            Case "MORICEAUCLAV-XXXX": X = "MORICEAU"
            Case "MOSTEFA-695": X = "MOSTEFA"
            Case "NAVAILLESCLAVIE-XXXX": X = "CARLOS"
            Case "CARLOSCLAVIER-XXXX": X = "CARLOS"
            Case "CARLOS-842": X = "CARLOS"
            Case "NGUON-725": X = "NGUON"
            Case "RICHARDIERE-761": X = "RICHARDIERE"
            Case "RICHARDIERE-868": X = "RICHARDIERE"
            Case "RABIA-878": X = "RABIA"
            Case "SALLE-803": X = "SALLE"
            Case "STAMBOULI-XXXX": X = "STAMBOULI"
            Case "STAMBOULICARTE-680": X = "STAMBOULI"
            Case "TABTI-833": X = "TABTI"
            Case "VAISSIER-802": X = "VAISSIER"
            Case "VAISSIER-879": X = "VAISSIER"
           
        End Select
    'If blnOk Then
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
             & " where SSIUSRNAT = ' ' and SSIUSRUIDX = '" & X & "'"
        Set rsSab_X = cnsab.Execute(xSQL)
        If Not rsSab_X.EOF Then
            Call rsYSSIDIV0_GetBuffer(rsSab, oldYSSIDIV0)

            newYSSIDIV0 = oldYSSIDIV0
            newYSSIDIV0.SSIDIVPRFK = " "
            mYSSIDIV0_Update = "Update"
        
            mYSSIDOM0_Update = "New"
            newYSSIDOM0.SSIDOMNAT = " "
            newYSSIDOM0.SSIDOMSTAK = " "
            newYSSIDOM0.SSIDOMUNIT = " "
            newYSSIDOM0.SSIDOMUIDN = rsSab_X("SSIUSRUIDN")
            newYSSIDOM0.SSIDOMDIDX = "DIV"
            newYSSIDOM0.SSIDOMSTAK = oldYSSIDIV0.SSIDIVSTAK
            newYSSIDOM0.SSIDOMDECH = 0
            newYSSIDOM0.SSIDOMUIDD = oldYSSIDIV0.SSIDIVUIDD
            newYSSIDOM0.SSIDOMUIDX = oldYSSIDIV0.SSIDIVUIDX
            newYSSIDOM0.SSIDOMPRFX = oldYSSIDIV0.SSIDIVPRFX
            newYSSIDOM0.SSIDOMPRFK = " "
            newYSSIDOM0.SSIDOMPRFD = oldYSSIDIV0.SSIDIVYAMJ
            newYSSIDOM0.SSIDOMPRFH = oldYSSIDIV0.SSIDIVYHMS
            newYSSIDOM0.SSIDOMTLNK = 0
            newYSSIDOM0.SSIDOMYFCT = "INI"
            newYSSIDOM0.SSIDOMYAMJ = DSys
            newYSSIDOM0.SSIDOMYHMS = time_Hms
            newYSSIDOM0.SSIDOMYUSR = usrName_UCase
            newYSSIDOM0.SSIDOMYVER = 0
            
            V = sqlYSSIDIV0_Update(newYSSIDIV0, oldYSSIDIV0)
            If Not IsNull(V) Then GoTo Error_MsgBox
            V = sqlYSSIDOM0_Insert(newYSSIDOM0)
            If Not IsNull(V) Then GoTo Error_MsgBox
        'End If
    End If
    rsSab.MoveNext
Loop


'________________________________________________________________________________


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        

    End If
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Set rsSab = Nothing

Exit Sub


End Sub


Private Sub paramMEL_Init()
Dim V, X As String, xIn As String
Dim xSQL As String, xWhere As String, Nb As Long
Dim rsMDB As New ADODB.Recordset
Dim rsSab_Local As New ADODB.Recordset

Dim K As Long, xUsr As String
Dim kLen As Integer, I As Integer, K1 As Integer, blnOk As Boolean, blnExit As Boolean

On Error GoTo Error_Handler

Call lstParam_SSIMELUNOM_Load("")

Import_MEL:

currentAction = "paramMEL_Init"
Call cmdUpdate_Init

Call rsYSSIMEL0_Init(newYSSIMEL0)
newYSSIMEL0.SSIMELNAT = "@"
newYSSIMEL0.SSIMELYAMJ = DSys
newYSSIMEL0.SSIMELYHMS = time_Hms
newYSSIMEL0.SSIMELYUSR = usrName_UCase

'Call MsgBox("goto ROPDOSMAIL")
'GoTo ROPDOSMAIL

X = "select * from ElpTable where SNN = 0" _
    & " and id = 'vbSendMail' order by K1,K2"
    
Set rsMDB = cnMDB.Execute(X)
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'________________________________________________________________________________

Do While Not rsMDB.EOF
    If Not IsNull(rsMDB("Memo")) Then
        newYSSIMEL0.SSIMELUIDX = Trim(rsMDB("K1")) & "." & Trim(rsMDB("K2"))
        newYSSIMEL0.SSIMELUNOM = Trim(rsMDB("Name"))
        newYSSIMEL0.SSIMELINFO = ""
        X = UCase$(Trim(rsMDB("Memo")))
        xUsr = Replace(X, "@BIA-PARIS.FR", "@bia-paris.fr")
        
        kLen = Len(xUsr)
        K1 = 1
        blnExit = False
        Do
            K = InStr(K1, xUsr, ";")
            If K > 0 Then
                blnOk = False
                X = StrConv(Trim(Mid$(xUsr, K1, K - K1)), vbProperCase)
                For I = 1 To arrSSIMELUNOM_Nb
                    If X = arrSSIMELUNOM(I) Then blnOk = True: newYSSIMEL0.SSIMELINFO = newYSSIMEL0.SSIMELINFO & X & ";": Exit For
                Next I
                If Not blnOk Then Call MsgBox(newYSSIMEL0.SSIMELUIDX & " ? " & X): Debug.Print newYSSIMEL0.SSIMELUIDX & " ? " & X
               K1 = K + 1
            Else
                X = StrConv(Trim(Mid$(xUsr, K1, kLen - K1 + 1)), vbProperCase)
                For I = 1 To arrSSIMELUNOM_Nb
                    If X = arrSSIMELUNOM(I) Then blnOk = True: newYSSIMEL0.SSIMELINFO = newYSSIMEL0.SSIMELINFO & X: Exit For
                Next I
                If Not blnOk Then Call MsgBox(newYSSIMEL0.SSIMELUIDX & " ? " & X): Debug.Print newYSSIMEL0.SSIMELUIDX & " ? " & X
                blnExit = True 'Exit Do
            End If
     
        Loop Until blnExit
        mYSSIMEL0_Update = "New"
        Call cmdUpdate
    End If
    rsMDB.MoveNext
Loop

'________________________________________________________________________________

GOSDOSMAIL:

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'GOSDOSMAIL' order by BIATABK1"
Set rsSab_Local = cnsab.Execute(X)
    
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'________________________________________________________________________________SSIWINMAIL

Do While Not rsSab_Local.EOF
    If Not IsNull(rsSab_Local("BIATABTXT")) Then
        newYSSIMEL0.SSIMELUIDX = "BIA_GOS" & "." & Trim(rsSab_Local("BIATABK1"))
        newYSSIMEL0.SSIMELUNOM = arrSSIUSRUNIT_Code(Mid$(rsSab_Local("BIATABK1"), 2, 2))
        newYSSIMEL0.SSIMELINFO = ""
        X = UCase$(Trim(rsSab_Local("BIATABTXT")))
        xUsr = Replace(X, "@BIA-PARIS.FR", "@bia-paris.fr")
        
        kLen = Len(xUsr)
        K1 = 1
        blnExit = False
        Do
            K = InStr(K1, xUsr, ";")
            If K > 0 Then
                blnOk = False
                X = mailAdresse_Production(Trim(Mid$(xUsr, K1, K - K1)))
                X = StrConv(X, vbProperCase)
                For I = 1 To arrSSIMELUNOM_Nb
                    If X = arrSSIMELUNOM(I) Then blnOk = True: newYSSIMEL0.SSIMELINFO = newYSSIMEL0.SSIMELINFO & X & ";": Exit For
                Next I
                If Not blnOk Then Call MsgBox(newYSSIMEL0.SSIMELUIDX & " ? " & X): Debug.Print newYSSIMEL0.SSIMELUIDX & " ? " & X
                K1 = K + 1
            Else
                X = mailAdresse_Production(Trim(Mid$(xUsr, K1, kLen - K1 + 1)))
                X = StrConv(X, vbProperCase)
               For I = 1 To arrSSIMELUNOM_Nb
                    If X = arrSSIMELUNOM(I) Then blnOk = True: newYSSIMEL0.SSIMELINFO = newYSSIMEL0.SSIMELINFO & X: Exit For
                Next I
                If Not blnOk Then Call MsgBox(newYSSIMEL0.SSIMELUIDX & " ? " & X): Debug.Print newYSSIMEL0.SSIMELUIDX & " ? " & X
               blnExit = True 'Exit Do
            End If
     
        Loop Until blnExit
        mYSSIMEL0_Update = "New"
        Call cmdUpdate
    End If
    rsSab_Local.MoveNext
Loop

'________________________________________________________________________________

SAA_Alerte:

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'SAA_Alerte' and BIATABK1 = 'Mail' order by BIATABK2"
Set rsSab_Local = cnsab.Execute(X)
    
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'________________________________________________________________________________SSIWINMAIL

Do While Not rsSab_Local.EOF
    If Not IsNull(rsSab_Local("BIATABTXT")) Then
        newYSSIMEL0.SSIMELUIDX = "SAA_Alerte" & "." & Trim(rsSab_Local("BIATABK2"))
        newYSSIMEL0.SSIMELUNOM = arrSSIUSRUNIT_Code(Mid$(rsSab_Local("BIATABK2"), 2, 2))
        newYSSIMEL0.SSIMELINFO = ""
        X = UCase$(Trim(rsSab_Local("BIATABTXT")))
        xUsr = Replace(X, "@BIA-PARIS.FR", "@bia-paris.fr")
        
        kLen = Len(xUsr)
        K1 = 1
        blnExit = False
        Do
            K = InStr(K1, xUsr, ";")
            If K > 0 Then
                blnOk = False
                X = Trim(Mid$(xUsr, K1, K - K1))
                Select Case X
                    Case "RICHARDIER": X = "RICHARDIERE"
                    Case "REOL_CH": X = "REOL_Ch"
                End Select
                X = mailAdresse_Production(X)
                X = StrConv(X, vbProperCase)
                For I = 1 To arrSSIMELUNOM_Nb
                    If X = arrSSIMELUNOM(I) Then blnOk = True: newYSSIMEL0.SSIMELINFO = newYSSIMEL0.SSIMELINFO & X & ";": Exit For
                Next I
                If Not blnOk Then Call MsgBox(newYSSIMEL0.SSIMELUIDX & " ? " & X): Debug.Print newYSSIMEL0.SSIMELUIDX & " ? " & X
                K1 = K + 1
            Else
                X = mailAdresse_Production(Trim(Mid$(xUsr, K1, kLen - K1 + 1)))
                X = StrConv(X, vbProperCase)
               For I = 1 To arrSSIMELUNOM_Nb
                    If X = arrSSIMELUNOM(I) Then blnOk = True: newYSSIMEL0.SSIMELINFO = newYSSIMEL0.SSIMELINFO & X: Exit For
                Next I
                If Not blnOk Then Call MsgBox(newYSSIMEL0.SSIMELUIDX & " ? " & X): Debug.Print newYSSIMEL0.SSIMELUIDX & " ? " & X
               blnExit = True 'Exit Do
            End If
     
        Loop Until blnExit
        mYSSIMEL0_Update = "New"
        Call cmdUpdate
    End If
    rsSab_Local.MoveNext
Loop
'________________________________________________________________________________

GOSDOSMAIL_R:

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'GOSDOSMAIL_R' order by BIATABK1"
Set rsSab_Local = cnsab.Execute(X)
    
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'________________________________________________________________________________SSIWINMAIL

Do While Not rsSab_Local.EOF
    If Not IsNull(rsSab_Local("BIATABTXT")) Then
        newYSSIMEL0.SSIMELUIDX = "RCOM" & "." & Trim(rsSab_Local("BIATABK1"))
        newYSSIMEL0.SSIMELUNOM = Trim(newYSSIMEL0.SSIMELUIDX)
        newYSSIMEL0.SSIMELINFO = ""
        X = UCase$(Trim(rsSab_Local("BIATABTXT")))
        xUsr = Replace(X, "@BIA-PARIS.FR", "@bia-paris.fr")
        
        kLen = Len(xUsr)
        K1 = 1
        blnExit = False
        Do
            K = InStr(K1, xUsr, ";")
            If K > 0 Then
                blnOk = False
                X = mailAdresse_Production(Trim(Mid$(xUsr, K1, K - K1)))
                X = StrConv(X, vbProperCase)
                For I = 1 To arrSSIMELUNOM_Nb
                    If X = arrSSIMELUNOM(I) Then blnOk = True: newYSSIMEL0.SSIMELINFO = newYSSIMEL0.SSIMELINFO & X & ";": Exit For
                Next I
                If Not blnOk Then Call MsgBox(newYSSIMEL0.SSIMELUIDX & " ? " & X): Debug.Print newYSSIMEL0.SSIMELUIDX & " ? " & X
                K1 = K + 1
            Else
                X = mailAdresse_Production(Trim(Mid$(xUsr, K1, kLen - K1 + 1)))
                X = StrConv(X, vbProperCase)
               For I = 1 To arrSSIMELUNOM_Nb
                    If X = arrSSIMELUNOM(I) Then blnOk = True: newYSSIMEL0.SSIMELINFO = newYSSIMEL0.SSIMELINFO & X: Exit For
                Next I
                If Not blnOk Then Call MsgBox(newYSSIMEL0.SSIMELUIDX & " ? " & X): Debug.Print newYSSIMEL0.SSIMELUIDX & " ? " & X
               blnExit = True 'Exit Do
            End If
     
        Loop Until blnExit
        mYSSIMEL0_Update = "New"
        Call cmdUpdate
    End If
    rsSab_Local.MoveNext
Loop
'________________________________________________________________________________

ROPDOSMAIL:

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'ROPDOSMAIL' order by BIATABK1"
Set rsSab_Local = cnsab.Execute(X)
    
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'________________________________________________________________________________SSIWINMAIL

Do While Not rsSab_Local.EOF
    If Not IsNull(rsSab_Local("BIATABTXT")) Then
        newYSSIMEL0.SSIMELUIDX = "DROPI" & "." & Trim(rsSab_Local("BIATABK1"))
        newYSSIMEL0.SSIMELUNOM = Trim(newYSSIMEL0.SSIMELUIDX)
        newYSSIMEL0.SSIMELINFO = ""
        X = UCase$(Trim(rsSab_Local("BIATABTXT")))
        xUsr = Replace(X, "@BIA-PARIS.FR", "@bia-paris.fr")
        
        kLen = Len(xUsr)
        K1 = 1
        blnExit = False
        Do
            K = InStr(K1, xUsr, ";")
            If K > 0 Then
                blnOk = False
                X = mailAdresse_Production(Trim(Mid$(xUsr, K1, K - K1)))
                X = StrConv(X, vbProperCase)
                For I = 1 To arrSSIMELUNOM_Nb
                    If X = arrSSIMELUNOM(I) Then blnOk = True: newYSSIMEL0.SSIMELINFO = newYSSIMEL0.SSIMELINFO & X & ";": Exit For
                Next I
                If Not blnOk Then Call MsgBox(newYSSIMEL0.SSIMELUIDX & " ? " & X): Debug.Print newYSSIMEL0.SSIMELUIDX & " ? " & X
                K1 = K + 1
            Else
                X = mailAdresse_Production(Trim(Mid$(xUsr, K1, kLen - K1 + 1)))
                X = StrConv(X, vbProperCase)
               For I = 1 To arrSSIMELUNOM_Nb
                    If X = arrSSIMELUNOM(I) Then blnOk = True: newYSSIMEL0.SSIMELINFO = newYSSIMEL0.SSIMELINFO & X: Exit For
                Next I
                If Not blnOk Then Call MsgBox(newYSSIMEL0.SSIMELUIDX & " ? " & X): Debug.Print newYSSIMEL0.SSIMELUIDX & " ? " & X
               blnExit = True 'Exit Do
            End If
     
        Loop Until blnExit
        mYSSIMEL0_Update = "New"
        Call cmdUpdate
    End If
    rsSab_Local.MoveNext
Loop

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : paramMEL_Init"
Exit_sub:

    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Set rsSab_Local = Nothing

Exit Sub


End Sub






Private Sub paramSAB_Init()
Dim V, X As String, xIn As String
Dim xSQL As String, xWhere As String, Nb As Long, K As Integer
Dim xUIDD As String, wPRFX As String, xALPHA As String, wUPUID As Long

On Error GoTo Error_Handler


Import_SAB:

currentAction = "paramSAB_Init"
Call cmdUpdate_Init

Call rsYSSIDOM0_Init(newYSSIDOM0)
newYSSIDOM0.SSIDOMDIDX = "SAB"
newYSSIDOM0.SSIDOMYAMJ = DSys
newYSSIDOM0.SSIDOMYHMS = time_Hms
newYSSIDOM0.SSIDOMYUSR = usrName_UCase
'_________________________________________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0, " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
     & " where SSISABNAT = ' ' and SSISABPRFK = '?' " _
     & " and SSIDOMUIDX = SSISABUIDX and SSIDOMNAt = ' ' and SSIDOMPRFX = 'SAB_PROD' order by SSISABUIDX"

Set rsSab = cnsab.Execute(xSQL)

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________



Do While Not rsSab.EOF
    Call rsYSSISAB0_GetBuffer(rsSab, oldYSSISAB0)
    newYSSISAB0 = oldYSSISAB0
    newYSSISAB0.SSISABPRFK = " "
    mYSSISAB0_Update = "Update"

    Call rsYSSIDOM0_GetBuffer(rsSab, newYSSIDOM0)
    mYSSIDOM0_Update = "New"
    newYSSIDOM0.SSIDOMDIDX = "SAB"
    newYSSIDOM0.SSIDOMSTAK = oldYSSISAB0.SSISABSTAK
    newYSSIDOM0.SSIDOMDECH = 0
    newYSSIDOM0.SSIDOMUIDD = oldYSSISAB0.SSISABUIDD
    newYSSIDOM0.SSIDOMUIDX = oldYSSISAB0.SSISABUIDX
    newYSSIDOM0.SSIDOMPRFX = oldYSSISAB0.SSISABPRFX
    If oldYSSISAB0.SSISABSTAK = " " Then
        newYSSIDOM0.SSIDOMPRFK = " "
    Else
        newYSSIDOM0.SSIDOMPRFK = "X"
    End If
    newYSSIDOM0.SSIDOMPRFD = oldYSSISAB0.SSISABYAMJ
    newYSSIDOM0.SSIDOMPRFH = oldYSSISAB0.SSISABYHMS
    newYSSIDOM0.SSIDOMTLNK = 0
    newYSSIDOM0.SSIDOMYFCT = "INI"
    newYSSIDOM0.SSIDOMYAMJ = DSys
    newYSSIDOM0.SSIDOMYHMS = time_Hms
    newYSSIDOM0.SSIDOMYUSR = usrName_UCase
    newYSSIDOM0.SSIDOMYVER = 0
    
   V = sqlYSSISAB0_Update(newYSSISAB0, oldYSSISAB0)
    If Not IsNull(V) Then GoTo Error_MsgBox
   V = sqlYSSIDOM0_Insert(newYSSIDOM0)
    If Not IsNull(V) Then GoTo Error_MsgBox

    rsSab.MoveNext
Loop

'____________________________________________________________________
 xSQL = "select SSIUSRUIDN from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & " where  SSIUSRUIDX = 'EXIT_GRP'"

Set rsSab_X = cnsab.Execute(xSQL)
Dim mSSIUSRUIDN_X As Long, wSSIUSRUIDN As Long
If rsSab_X.EOF Then
    Call MsgBox("Créer un utilisateur 'EXIT_GRP' = poubelle", vbCritical, "paramSAB_init")
    GoTo Exit_sub
End If
mSSIUSRUIDN_X = rsSab_X("SSIUSRUIDN")
'____________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0" _
     & " where SSISABNAT = ' ' and SSISABPRFK = '?' " _
     & " order by SSISABUIDX"

Set rsSab = cnsab.Execute(xSQL)


Do While Not rsSab.EOF
    Call rsYSSISAB0_GetBuffer(rsSab, oldYSSISAB0)
    mYSSISAB0_Update = ""
    
    If Mid$(oldYSSISAB0.SSISABUIDX, 1, 2) = "G_" Then mYSSISAB0_Update = "Update": wSSIUSRUIDN = mSSIUSRUIDN_X
    If Mid$(oldYSSISAB0.SSISABUIDX, 1, 1) = "$" Then mYSSISAB0_Update = "Update": wSSIUSRUIDN = mSSIUSRUIDN_X
    Select Case Trim(oldYSSISAB0.SSISABUIDX)
        Case "ALIS_C":  X = "ALIS": mYSSISAB0_Update = "Update": wSSIUSRUIDN = 0
        Case "CHAUVEAU":  mYSSISAB0_Update = "Update": wSSIUSRUIDN = mSSIUSRUIDN_X
        Case "HALTEBOURG":  X = "HALTEBOU": mYSSISAB0_Update = "Update": wSSIUSRUIDN = 0
        Case "P_ST1":  mYSSISAB0_Update = "Update": wSSIUSRUIDN = mSSIUSRUIDN_X
        Case "ZZZZZZZZZZ":  mYSSISAB0_Update = "Update": wSSIUSRUIDN = mSSIUSRUIDN_X
        Case "FONTANA": X = "FONTANA": mYSSISAB0_Update = "Update": wSSIUSRUIDN = 0
        Case "SABTELE": X = "SAB": mYSSISAB0_Update = "Update": wSSIUSRUIDN = 0
        Case "X_DARMON": X = "DARMON": mYSSISAB0_Update = "Update": wSSIUSRUIDN = 0
        Case "X_CHAUMERE": X = "CHAUMERET": mYSSISAB0_Update = "Update": wSSIUSRUIDN = 0
        Case "X_DELALAND": X = "DELALANDE": mYSSISAB0_Update = "Update": wSSIUSRUIDN = 0
        Case "X_VIGNERON": X = "VIGNERON": mYSSISAB0_Update = "Update": wSSIUSRUIDN = 0
    End Select
    
    If mYSSISAB0_Update = "Update" Then
        If wSSIUSRUIDN = 0 Then
            xSQL = "select SSIUSRUIDN from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
                & " where  SSIUSRUIDX = '" & X & "'"

            Set rsSab_X = cnsab.Execute(xSQL)
            wSSIUSRUIDN = rsSab_X("SSIUSRUIDN")
        End If
        newYSSISAB0 = oldYSSISAB0
        newYSSISAB0.SSISABPRFK = " "
        mYSSIDOM0_Update = "New"
        newYSSIDOM0.SSIDOMNAT = ""
        newYSSIDOM0.SSIDOMUIDN = wSSIUSRUIDN
        newYSSIDOM0.SSIDOMDIDX = "SAB"
        newYSSIDOM0.SSIDOMSTAK = oldYSSISAB0.SSISABSTAK
        newYSSIDOM0.SSIDOMDECH = 0
        newYSSIDOM0.SSIDOMUIDD = oldYSSISAB0.SSISABUIDD
        newYSSIDOM0.SSIDOMUIDX = oldYSSISAB0.SSISABUIDX
        newYSSIDOM0.SSIDOMPRFX = oldYSSISAB0.SSISABPRFX
        If oldYSSISAB0.SSISABSTAK = " " Then
            newYSSIDOM0.SSIDOMPRFK = " "
        Else
            newYSSIDOM0.SSIDOMPRFK = "X"
        End If
        newYSSIDOM0.SSIDOMPRFD = oldYSSISAB0.SSISABYAMJ
        newYSSIDOM0.SSIDOMPRFH = oldYSSISAB0.SSISABYHMS
        newYSSIDOM0.SSIDOMTLNK = 0
        newYSSIDOM0.SSIDOMYFCT = "INI"
        newYSSIDOM0.SSIDOMYAMJ = DSys
        newYSSIDOM0.SSIDOMYHMS = time_Hms
        newYSSIDOM0.SSIDOMYUSR = usrName_UCase
        newYSSIDOM0.SSIDOMYVER = 0
    
        V = sqlYSSISAB0_Update(newYSSISAB0, oldYSSISAB0)
         If Not IsNull(V) Then GoTo Error_MsgBox
        V = sqlYSSIDOM0_Insert(newYSSIDOM0)
         If Not IsNull(V) Then GoTo Error_MsgBox
    End If

    rsSab.MoveNext
Loop


'________________________________________________________________________________


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        

    End If
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Set rsSab = Nothing

Exit Sub


End Sub

Private Sub paramSSIUSRUNIT_Init()
Dim V, X As String, xIn As String
Dim xSQL As String, xWhere As String, Nb As Long, K As Integer
Dim xSSIUSRUNIT As String, blnUpdate As Boolean
On Error GoTo Error_Handler


Import_SAB:

currentAction = "paramSSIUSRUNIT_Init"
Call cmdUpdate_Init

Call rsYSSIUSR0_Init(newYSSIUSR0)
newYSSIUSR0.SSIUSRNAT = "S"
newYSSIUSR0.SSIUSRYFCT = "INI"
newYSSIUSR0.SSIUSRYAMJ = DSys
newYSSIUSR0.SSIUSRYHMS = time_Hms
newYSSIUSR0.SSIUSRYUSR = usrName_UCase
'_________________________________________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'ROPDOSISRV'" _
     & " order by BIATABK1"
Call FEU_ROUGE
Set rsSab = cnsab.Execute(xSQL)
Call FEU_VERT
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
'Call MsgBox("GoTo ROPDOSGUSR", vbExclamation, "paramSSIUSRUNIT_Init")
'******************
'GoTo SSIUSRUNIT_S99
'******************

Do While Not rsSab.EOF
    mYSSIUSR0_Update = "New"
    X = Trim(rsSab("BIATABK1"))
    newYSSIUSR0.SSIUSRUNIT = X
    newYSSIUSR0.SSIUSRUIDN = Val(Mid$(X, 2, 2))
    X = Trim(rsSab("BIATABTXT"))
    newYSSIUSR0.SSIUSRUIDX = Mid$(X, 13, 20)
    newYSSIUSR0.SSIUSRPRFX = Mid$(X, 1, 12)

    Select Case newYSSIUSR0.SSIUSRUNIT
        Case "S01": newYSSIUSR0.SSIUSRUIDX = "Moyens de paiement  "
        Case "S10": newYSSIUSR0.SSIUSRUIDX = "CréditsDocumentaires"
        Case "S11": newYSSIUSR0.SSIUSRUIDX = "Soutien opérationnel"
        Case "S12": newYSSIUSR0.SSIUSRUIDX = "Sécurité            "
        Case "S20": newYSSIUSR0.SSIUSRUIDX = "Direction générale  "
        Case "S21": newYSSIUSR0.SSIUSRUIDX = "Inspection          "
        Case "S22": newYSSIUSR0.SSIUSRUIDX = "Contrôle de gestion "
        Case "S30": newYSSIUSR0.SSIUSRUIDX = "Dir des opérations  "
        Case "S31": newYSSIUSR0.SSIUSRUIDX = "Organisation        "
        Case "S32": newYSSIUSR0.SSIUSRUIDX = "Gestion crédits/BOTC"
        Case "S33": newYSSIUSR0.SSIUSRUIDX = "Contrôle permanent  ": newYSSIUSR0.SSIUSRPRFX = "Ctrl permanent"
        Case "S34": newYSSIUSR0.SSIUSRUIDX = "Maîtrise d'ouvrage  "
        Case "S40": newYSSIUSR0.SSIUSRUIDX = "Informatique        "
        Case "S41": newYSSIUSR0.SSIUSRUIDX = "Dir commerciale     "
        Case "S42": newYSSIUSR0.SSIUSRUIDX = "Sécurité financière "
        Case "S50": newYSSIUSR0.SSIUSRUIDX = "Dir générale adjoint"
        Case "S51": newYSSIUSR0.SSIUSRUIDX = "Dép des risques     "
        Case "S52": newYSSIUSR0.SSIUSRUIDX = "Juridique           ": newYSSIUSR0.SSIUSRPRFX = "Juridique"
        Case "S53": newYSSIUSR0.SSIUSRUIDX = "Contentieux         "
        Case "S54": newYSSIUSR0.SSIUSRUIDX = "Front office TC     "
        Case "S60": newYSSIUSR0.SSIUSRUIDX = "Comptabilité        "
        Case "S61": newYSSIUSR0.SSIUSRUIDX = "Relations humaines  "
        Case "S62": newYSSIUSR0.SSIUSRUIDX = "Services généraux   "
        Case "S97": newYSSIUSR0.SSIUSRUIDX = "destinataires Alerte"
        Case "S98": newYSSIUSR0.SSIUSRUIDX = "Utilisateurs divers "
        Case "S99": newYSSIUSR0.SSIUSRUIDX = "Utilisateurs non HAB"
        
    End Select
   V = sqlYSSIUSR0_Insert(newYSSIUSR0)
    If Not IsNull(V) Then GoTo Error_MsgBox

    rsSab.MoveNext
Loop

'_________________________________________________________________________________________________
Modèle:

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & " where SSIUSRNAT = '$'  order by SSIUSRUIDX"

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Select Case Trim(rsSab("SSIUSRUIDX"))
        Case "BIA_ADMIN": X = "S40"
        Case "BIA_ADMIN   /1": X = "S40"
        Case "BIA_ADMIN   /2": X = "S40"
        Case "BIA_ADMIN   /3": X = "S40"
        Case "BIA_CAC": X = "S98"
        Case "BIA_CGES": X = "S22"
        Case "BIA_CPT_V": X = "S60"
        Case "BIA_CPT_V1": X = "S60"
        Case "BIA_DAFI_S": X = "S32"
        Case "BIA_DAFI_V": X = "S32"
        Case "BIA_DCOM": X = "S41"
        Case "BIA_DCOM    /1": X = "S41"
        Case "BIA_DEON": X = "S42"
        Case "BIA_DER": X = "S51"
        Case "BIA_DER_V": X = "S51"
        Case "BIA_DER_1": X = "S51"
        Case "BIA_DRH_S": X = "S61"
        Case "BIA_FOTC": X = "S54"
        Case "BIA_GSOP_S": X = "S11"
        Case "BIA_INSPECTA": X = "S21"
        Case "BIA_INSPECTI": X = "S21"
        Case "BIA_JURIDIQU": X = "S52"
        Case "BIA_SOBF_SV": X = "S01"
        Case "BIA_SOBF_S2": X = "S01"
        Case "BIA_SOBF_S3": X = "S01"
        Case "BIA_SOBI_SV": X = "S10"
        Case "BIA_SOBI_S3": X = "S10"
        Case "BIA_SOBI_S3B": X = "S10"
        Case "BIA_SWI": X = "S62"
        Case "STAGIAIRE DCOM": X = "S41"
        Case "STAGIAIRE DER": X = "S51"
        Case "STAGIAIRE GSOP": X = "S11"
        Case Else:
            X = ""
    End Select
    If X <> "" Then
        xSQL = "Update " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
         & " set SSIUSRUNIT = '" & X & "' " _
         & "  Where SSIUSRNAT = '$' and SSIUSRUIDN = " & rsSab("SSIUSRUIDN")
        Call FEU_ROUGE
        V = sqlYSSIUSR0_Update_CMD(xSQL)
        Call FEU_VERT
        If Not IsNull(V) Then GoTo Error_MsgBox
    End If
    
    rsSab.MoveNext
Loop
'_________________________________________________________________________________________________
ROPDOSGUSR:

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'ROPDOSGUSR'" _
     & " order by BIATABK1"

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
         & " where SSIUSRNAT = ' ' and SSIUSRUIDX = '" & Trim(rsSab("BIATABK1")) & "'"
    Set rsSab_X = cnsab.Execute(xSQL)

    If Not rsSab_X.EOF Then
        xSSIUSRUNIT = Mid$(rsSab("BIATABTXT"), 26, 3)
        xSQL = "Update " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
         & " set SSIUSRUNIT = '" & xSSIUSRUNIT & "' " _
         & "  Where SSIUSRNAT = ' ' and SSIUSRUIDN = " & rsSab_X("SSIUSRUIDN")
        Call FEU_ROUGE
        V = sqlYSSIUSR0_Update_CMD(xSQL)
        Call FEU_VERT
        If Not IsNull(V) Then GoTo Error_MsgBox
        
         xSQL = "Update " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
         & " set SSIDOMUNIT = '" & xSSIUSRUNIT & "' " _
         & "  Where SSIDOMNAT = ' ' and SSIDOMUIDN = " & rsSab_X("SSIUSRUIDN")
        Call FEU_ROUGE
        V = sqlYSSIDOM0_Update_CMD(xSQL)
        Call FEU_VERT
        If Not IsNull(V) Then GoTo Error_MsgBox
   Else
        Debug.Print Trim(rsSab("BIATABK1"))
    End If
    rsSab.MoveNext
Loop

'________________________________________________________________________________
SSIUSRUNIT_S99:
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & " where SSIUSRNAT = ' ' and SSIUSRUNIT = '' "
Set rsSab_X = cnsab.Execute(xSQL)
Do While Not rsSab_X.EOF
    If rsSab_X("SSIUSRUIDN") >= 1279 And rsSab_X("SSIUSRUIDN") <= 1292 Then
        blnUpdate = False
    Else
        blnUpdate = True
    End If
    
    If rsSab_X("SSIUSRSTAK") = "N" Then
        xSSIUSRUNIT = "S99"
    Else
        xSSIUSRUNIT = "S98"
    End If
    Select Case rsSab_X("SSIUSRUIDN")
        Case 1001, 1004, 1010, 1011: blnUpdate = False
        Case 1002: xSSIUSRUNIT = "S98"
        Case 1003: xSSIUSRUNIT = "S40"
        Case 1325: xSSIUSRUNIT = "S51"
    End Select
    If blnUpdate Then
        xSQL = "Update " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
         & " set SSIUSRUNIT = '" & xSSIUSRUNIT & "' " _
         & "  Where SSIUSRNAT = ' ' and SSIUSRUIDN = " & rsSab_X("SSIUSRUIDN")
        Call FEU_ROUGE
        V = sqlYSSIUSR0_Update_CMD(xSQL)
        Call FEU_VERT
        If Not IsNull(V) Then GoTo Error_MsgBox
        
         xSQL = "Update " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
         & " set SSIDOMUNIT = '" & xSSIUSRUNIT & "' " _
         & "  Where SSIDOMNAT = ' ' and SSIDOMUIDN = " & rsSab_X("SSIUSRUIDN")
        Call FEU_ROUGE
        V = sqlYSSIDOM0_Update_CMD(xSQL)
        Call FEU_VERT
        If Not IsNull(V) Then GoTo Error_MsgBox
   Else
        'Debug.Print Trim(rsSab("BIATABK1"))
    End If
    rsSab_X.MoveNext
Loop
'________________________________________________________________________________

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        

    End If
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Set rsSab = Nothing

Exit Sub


End Sub



Private Sub paramSAA_Init_Compte()
Dim V, X As String, xIn As String
Dim xSQL As String, xWhere As String, Nb As Long, K As Integer
Dim xUIDD As String, wPRFX As String, xALPHA As String, wUPUID As Long
On Error GoTo Error_Handler




currentAction = "paramSAA_Init_Compte"
Call cmdUpdate_Init

'GoTo Param_SAA
'____________________________________________________________________
Call rsYSSIUSR0_Init(newYSSIUSR0)


newYSSIUSR0.SSIUSRYFCT = "INI"
newYSSIUSR0.SSIUSRYAMJ = DSys
newYSSIUSR0.SSIUSRYHMS = time_Hms
newYSSIUSR0.SSIUSRYUSR = usrName_UCase
newYSSIUSR0.SSIUSRSTAK = "N"
newYSSIUSR0.SSIUSRDECH = 0
newYSSIUSR0.SSIUSRPRFK = "X"

newYSSIUSR0.SSIUSRUIDX = "BENHAMOU": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "BOUDROUAZ": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "CHRYSANTOS": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "HAFIZ": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "HELIOUI": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "HOSTINGUE": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "HRABI": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "LECCIA": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "MARCOT": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "MERABIA": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "SHALLOUF": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "TERF": mYSSIUSR0_Update = "New": Call cmdUpdate

'____________________________________________________________________


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

    
Set rsSab = Nothing

Exit Sub


End Sub
Private Sub paramWIN_Init_Compte()
Dim V, X As String, xIn As String
Dim xSQL As String, xWhere As String, Nb As Long, K As Integer
Dim xUIDD As String, wPRFX As String, xALPHA As String, wUPUID As Long
On Error GoTo Error_Handler

currentAction = "paramWIN_Init_Compte"
Call cmdUpdate_Init

Call rsYSSIDOM0_Init(newYSSIDOM0)

newYSSIDOM0.SSIDOMNAT = "$"

newYSSIDOM0.SSIDOMYFCT = "INI"
newYSSIDOM0.SSIDOMYAMJ = DSys
newYSSIDOM0.SSIDOMYHMS = time_Hms
newYSSIDOM0.SSIDOMYUSR = usrName_UCase
newYSSIDOM0.SSIDOMSTAK = " "

newYSSIDOM0.SSIDOMDIDX = "WIN"
newYSSIDOM0.SSIDOMUIDD = -1
newYSSIDOM0.SSIDOMUIDX = "21-Actifs"
newYSSIDOM0.SSIDOMPRFX = "21-Actifs"

For K = 1 To 32
    newYSSIDOM0.SSIDOMUIDN = K
    mYSSIDOM0_Update = "New": Call cmdUpdate
Next K

'GoTo Exit_sub


Call rsYSSIWIN0_Init(newYSSIWIN0)

newYSSIWIN0.SSIWINNAT = "$"
newYSSIWIN0.SSIWINUIDD = 0
newYSSIWIN0.SSIWININFO = "||||||||||||||||||||||"

newYSSIWIN0.SSIWINYFCT = "INI"
newYSSIWIN0.SSIWINYAMJ = DSys
newYSSIWIN0.SSIWINYHMS = time_Hms
newYSSIWIN0.SSIWINYUSR = usrName_UCase
newYSSIWIN0.SSIWINSTAK = " "

newYSSIWIN0.SSIWINUIDD = 0
newYSSIWIN0.SSIWINGUID = "divers"
newYSSIWIN0.SSIWINUIDX = "0-DIVERS "
newYSSIWIN0.SSIWINUNOM = "divers"
newYSSIWIN0.SSIWINPRFX = "Divers"
mYSSIWIN0_Update = "New": Call cmdUpdate

newYSSIWIN0.SSIWINUIDD = 1
newYSSIWIN0.SSIWINGUID = "OU=Domain Controllers"
newYSSIWIN0.SSIWINUIDX = "1-Domain Controller"
newYSSIWIN0.SSIWINUNOM = "OU=Domain Controllers"
newYSSIWIN0.SSIWINPRFX = "Domain Controllers"
mYSSIWIN0_Update = "New": Call cmdUpdate

newYSSIWIN0.SSIWINUIDD = 2
newYSSIWIN0.SSIWINGUID = "OU=Serveurs BIA"
newYSSIWIN0.SSIWINUIDX = "2-Serveurs BIA"
newYSSIWIN0.SSIWINUNOM = "OU=Serveurs BIA"
newYSSIWIN0.SSIWINPRFX = "Serveurs BIA"
mYSSIWIN0_Update = "New": Call cmdUpdate

newYSSIWIN0.SSIWINUIDD = 3
newYSSIWIN0.SSIWINGUID = "OU=SYSTEME,OU=Groupes BIA"
newYSSIWIN0.SSIWINUIDX = "3-SYSTEME,Groupes"
newYSSIWIN0.SSIWINUNOM = "OU=SYSTEME,OU=Groupes BIA"
newYSSIWIN0.SSIWINPRFX = "SYSTEME,Groupes"
mYSSIWIN0_Update = "New": Call cmdUpdate

newYSSIWIN0.SSIWINUIDD = 4
newYSSIWIN0.SSIWINGUID = "OU=Groupes BIA"
newYSSIWIN0.SSIWINUIDX = "4-Groupes BIA"
newYSSIWIN0.SSIWINUNOM = "OU=Groupes BIA"
newYSSIWIN0.SSIWINPRFX = "Groupes BIA"
mYSSIWIN0_Update = "New": Call cmdUpdate

newYSSIWIN0.SSIWINUIDD = 5
newYSSIWIN0.SSIWINGUID = "OU=Login Specific Scan"
newYSSIWIN0.SSIWINUIDX = "5-Login Scan"
newYSSIWIN0.SSIWINUNOM = "OU=Login Specific Scan,OU=Utilisateurs BIA"
newYSSIWIN0.SSIWINPRFX = "Login Scan"
mYSSIWIN0_Update = "New": Call cmdUpdate

'____________________________________________________________________
Call rsYSSIUSR0_Init(newYSSIUSR0)
Call cmdUpdate_Init


newYSSIUSR0.SSIUSRYFCT = "INI"
newYSSIUSR0.SSIUSRYAMJ = DSys
newYSSIUSR0.SSIUSRYHMS = time_Hms
newYSSIUSR0.SSIUSRYUSR = usrName_UCase
newYSSIUSR0.SSIUSRSTAK = " "
newYSSIUSR0.SSIUSRDECH = 0
newYSSIUSR0.SSIUSRPRFK = " "


newYSSIUSR0.SSIUSRUIDX = "WIN_SERVEURS": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "WIN_IMPRIMANTES": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "WIN_PC": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "BIA_GROUPES": mYSSIUSR0_Update = "New": Call cmdUpdate
''''newYSSIUSR0.SSIUSRUIDX = "BIA_FOLDER": mYSSIUSR0_Update = "New": Call cmdUpdate


newYSSIUSR0.SSIUSRUIDX = "BIA_WINDOWS": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "BIA_SCAN": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "BIA_REUTERS": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "BIA_SERVEURS": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "VAISSIER FLORIAN": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "DATABAIL": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "STANDARD": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "WINDOWS": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "SUNGARD": mYSSIUSR0_Update = "New": Call cmdUpdate

newYSSIUSR0.SSIUSRSTAK = "N"
newYSSIUSR0.SSIUSRPRFK = "X"
newYSSIUSR0.SSIUSRUIDX = "CHERGUI": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "AUDISOFT": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "BEDJAOUI": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "BELABED": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "BENNAI": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "BRESNU": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "BRESSION": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "ELGHAZI": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "FAUTRAT": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "FERCATI": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "GATELLIER": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "HADJ": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "HASNIOU": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "HOUCHI": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "JOSEPH-PARFAITE": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "LABORDE": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "LAIDI": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "LALLET": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "LEVACHER": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "LUSSIGNY": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "NGUYEN": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "SAIDANE": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "SAUZE": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "SEGHAYER": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "SOULIERS": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "TALEB": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "TAMINE": mYSSIUSR0_Update = "New": Call cmdUpdate
newYSSIUSR0.SSIUSRUIDX = "FERREIRA": mYSSIUSR0_Update = "New": Call cmdUpdate

'____________________________________________________________________


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

    
Set rsSab = Nothing

Exit Sub


End Sub

Private Sub paramDIV_Init_UGM_PRFX()
Dim V, X As String, xIn As String
Dim xSQL As String, xWhere As String, Nb As Long, K As Integer
Dim xUIDD As String, wPRFX As String, xALPHA As String, wUPUID As Long
On Error GoTo Error_Handler

currentAction = "paramDIV_Init_UGM"
Call cmdUpdate_Init

Call rsYSSITXT0_Init(newYSSITXT0)


newYSSITXT0.SSITXTNAT = "P"
newYSSITXT0.SSITXTDIDX = "UGM"
newYSSITXT0.SSITXTYAMJ = DSys
newYSSITXT0.SSITXTYHMS = time_Hms
newYSSITXT0.SSITXTYUSR = usrName_UCase
newYSSITXT0.SSITXTTLNK = 1: newYSSITXT0.SSITXTINFO = "C:\Temp\ugm.txt": mYSSITXT0_Update = "New": Call cmdUpdate


Call rsYSSIDIV0_Init(newYSSIDIV0)

newYSSIDIV0.SSIDIVNAT = "$"
newYSSIDIV0.SSIDIVUIDD = 0
newYSSIDIV0.SSIDIVDIDK = "UGM"

newYSSIDIV0.SSIDIVYFCT = "INI"
newYSSIDIV0.SSIDIVYAMJ = DSys
newYSSIDIV0.SSIDIVYHMS = time_Hms
newYSSIDIV0.SSIDIVYUSR = usrName_UCase
newYSSIDIV0.SSIDIVSTAK = " "

newYSSIDIV0.SSIDIVUIDD = 1
newYSSIDIV0.SSIDIVUIDX = "UGM_GDMP"
newYSSIDIV0.SSIDIVPRFX = "Caisse"
newYSSIDIV0.SSIDIVUNOM = "lecteur bagde GDMP"
newYSSIDIV0.SSIDIVINFO = newYSSIDIV0.SSIDIVUNOM
mYSSIDIV0_Update = "New": Call cmdUpdate

newYSSIDIV0.SSIDIVUIDD = 2
newYSSIDIV0.SSIDIVUIDX = "UGM_INFO"
newYSSIDIV0.SSIDIVPRFX = "Informatique"
newYSSIDIV0.SSIDIVUNOM = "lecteur bagde Informatique"
newYSSIDIV0.SSIDIVINFO = newYSSIDIV0.SSIDIVUNOM
mYSSIDIV0_Update = "New": Call cmdUpdate

newYSSIDIV0.SSIDIVUIDD = 3
newYSSIDIV0.SSIDIVUIDX = "UGM_FOTC badge"
newYSSIDIV0.SSIDIVPRFX = "Badge 5 EME"
newYSSIDIV0.SSIDIVUNOM = "lecteur bagde FOTC"
newYSSIDIV0.SSIDIVINFO = newYSSIDIV0.SSIDIVUNOM
mYSSIDIV0_Update = "New": Call cmdUpdate

newYSSIDIV0.SSIDIVUIDD = 4
newYSSIDIV0.SSIDIVUIDX = "UGM_FOTC clavier"
newYSSIDIV0.SSIDIVPRFX = "Clavier 5 EME"
newYSSIDIV0.SSIDIVUNOM = "clavier FOTC"
newYSSIDIV0.SSIDIVINFO = newYSSIDIV0.SSIDIVUNOM
mYSSIDIV0_Update = "New": Call cmdUpdate


'____________________________________________________________________


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

    
Set rsSab = Nothing

Exit Sub


End Sub


Private Sub paramFile_Init()
Dim V, X As String, xIn As String
Dim xSQL As String, xWhere As String, Nb As Long, K As Integer
Dim xUIDD As String, wPRFX As String, xALPHA As String, wUPUID As Long
On Error GoTo Error_Handler

 


currentAction = "paramFile_Init"
Call cmdUpdate_Init

Param_SAA:
Call rsYSSITXT0_Init(newYSSITXT0)


newYSSITXT0.SSITXTNAT = "P"
newYSSITXT0.SSITXTDIDX = "SAA"
newYSSITXT0.SSITXTYAMJ = DSys
newYSSITXT0.SSITXTYHMS = time_Hms
newYSSITXT0.SSITXTYUSR = usrName_UCase
newYSSITXT0.SSITXTTLNK = 1: newYSSITXT0.SSITXTINFO = "C:\Temp\SAA Unit.txt": mYSSITXT0_Update = "New": Call cmdUpdate
newYSSITXT0.SSITXTTLNK = 2: newYSSITXT0.SSITXTINFO = "C:\Temp\SAA Profile.txt": mYSSITXT0_Update = "New": Call cmdUpdate
newYSSITXT0.SSITXTTLNK = 3: newYSSITXT0.SSITXTINFO = "C:\Temp\SAA Operator.txt": mYSSITXT0_Update = "New": Call cmdUpdate
newYSSITXT0.SSITXTTLNK = 4: newYSSITXT0.SSITXTINFO = "C:\Temp\BIA_SSI_Archives": mYSSITXT0_Update = "New": Call cmdUpdate

newYSSITXT0.SSITXTDIDX = "BIA"
newYSSITXT0.SSITXTTLNK = 1
newYSSITXT0.SSITXTINFO = "\\BiaDoc\Filigrane\VB_RTF_Modèle.rtf" '"c:\temp\VB_RTF_Modèle.rtf"
mYSSITXT0_Update = "New": Call cmdUpdate

'____________________________________________________________________


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

    
Set rsSab = Nothing

Exit Sub


End Sub

Private Sub cmdSelect_SQL_9_SAA_Unit()
Dim V, X As String, xIn As String
Dim xSQL As String, xWhere As String, Nb As Long, K As Integer
Dim blnUnit As Boolean
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_9_SAA_Unit"
Call lstErr_Clear(lstErr, cmdContext, "> " & currentAction & " ........"): DoEvents

 Call paramSAA_Load
 
 '____________________________________________________________________
Call cmdUpdate_Init
Call rsYSSISAA0_Init(xYSSISAA0)

xYSSISAA0.SSISAANAT = "U"

Open mFile For Input As 1

blnUnit = False
Do Until EOF(1)
    Line Input #1, xIn
    xIn = Trim(xIn)
    If xIn <> "" Then
    
        I = InStr(1, xIn, "Number of entries:")
        If I > 0 Then
            mImport_Nb = Val(Mid$(xIn, I + 18, Len(xIn) - I - 17))
        End If
        
        
        I = InStr(1, xIn, "Unit Name")
        If I > 0 Then
            If blnUnit Then
                 Call cmdSelect_SQL_9_SAA_Update
                 mImport_In = mImport_In + 1
                 blnUnit = False
            End If
            
            blnUnit = True
            xYSSISAA0.SSISAAUIDX = Mid$(xIn, I + 17, Len(xIn) - I - 16)
            xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
                 & " where SSISAANAT = 'U' and SSISAAUIDX = '" & xYSSISAA0.SSISAAUIDX & "'"
            
            Set rsSab = cnsab.Execute(xSQL)
            If Not rsSab.EOF Then
                mYSSISAA0_Update = "Update+H"
                Call rsYSSISAA0_GetBuffer(rsSab, oldYSSISAA0)
                newYSSISAA0 = oldYSSISAA0
                newYSSISAA0.SSISAAYFCT = "MOD"

            Else
                mYSSISAA0_Update = "New"
                newYSSISAA0 = xYSSISAA0
                newYSSISAA0.SSISAAYFCT = "CRE"
                
            End If
        End If
        
        If blnUnit Then
            I = InStr(1, xIn, "Approval state")
            If I > 0 Then
                I = InStr(I + 17, xIn, "Approved")
                If I > 0 Then
                    newYSSISAA0.SSISAAPRFK = " "
                Else
                    newYSSISAA0.SSISAAPRFK = "N"
                End If
            End If
            
            I = InStr(1, xIn, "Description")
            If I > 0 Then
                newYSSISAA0.SSISAAUNOM = Mid$(xIn, I + 17, Len(xIn) - I - 16)
            End If
        End If
    End If
Loop

If blnUnit Then
     Call cmdSelect_SQL_9_SAA_Update
     mImport_In = mImport_In + 1
End If
            

Close

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

Set rsSab = Nothing
Call paramSAA_Load

Exit Sub


End Sub
Private Sub cmdSelect_SQL_9_TIC_USR()

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' nb roles par utilisateur : 50
' nb droits par roles      : 500
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Const nbRoles As Integer = 50
Const nbDroits As Integer = 500

Dim V, X As String, xIn As String
Dim xSQL As String, xWhere As String, Nb As Long, K As Integer, K2 As Integer
Dim blnEnd As Boolean, blnFile_Ok As Boolean, blnInput_Skip As Boolean
Dim blnUsers_Ok As Boolean, blnRoles_Ok As Boolean, blnDroits_Ok As Boolean
Dim blnExiste As Boolean

Dim arrRoles_UIDX() As String, arrRoles_Droits() As String, arrRoles_Info() As String, arrRoles_Nb As Integer, arrRoles_K As Integer
Dim arrRoles_Ordre() As Integer, arrRoles_Ordre_K As Integer
'Dim arrUsers_UIDX() As String, arrUsers_UNOM() As String, arrUsers_Roles() As String, arrUsers_Info() As String, arrUsers_Nb As Integer
'Dim arrUsers_PRFK() As String, arrUsers_UNOM_xIn() As String
Dim oldUsers() As typeYSSITIC0, newUsers() As typeYSSITIC0, arrUsers_Nb As Integer
Dim arrPRFX() As typeYSSITIC0, arrPRFX_Nb As Integer
Dim arrDroits_UNOM() As String, arrDroits_Info() As String, arrDroits_Nb As Integer

Dim wSSITICUNOM As String, xDFin As String, mUsers As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_9_ATHIC_USR"
Call lstErr_Clear(lstErr, cmdContext, "> " & currentAction & " ........"): DoEvents

'__________________________________________________________________
Call rsYSSITIC0_Init(usrYSSITIC0)
usrYSSITIC0.SSITICYFCT = "CRE"
usrYSSITIC0.SSITICYUSR = usrName_UCase
usrYSSITIC0.SSITICYAMJ = mImport_PRFD
usrYSSITIC0.SSITICYHMS = mImport_PRFH

Call rsYSSIDOM0_Init(xYSSIDOM0)
xYSSIDOM0.SSIDOMDIDX = "TIC"
xYSSIDOM0.SSIDOMYFCT = "CRE"
xYSSIDOM0.SSIDOMYUSR = usrName_UCase
xYSSIDOM0.SSIDOMYAMJ = mImport_PRFD
xYSSIDOM0.SSIDOMYHMS = mImport_PRFH
'__________________________________________________________________

xSQL = "select SSITICUIDD  from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
     & " where SSITICNAT = 'R' order by SSITICUIDD desc FETCH FIRST 1 ROWS ONLY"
Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    arrRoles_Nb = 0
Else
    arrRoles_Nb = rsSab(0)
End If
ReDim arrRoles_UIDX(arrRoles_Nb + 1), arrRoles_Droits(arrRoles_Nb + 1) As String, arrRoles_Info(arrRoles_Nb + 1), arrRoles_Ordre(arrRoles_Nb + 1)

xSQL = "select *  from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
     & " where SSITICNAT = 'R'" _
     & " order by SSITICUIDD"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    K = rsSab("SSITICUIDD")
    arrRoles_UIDX(K) = Trim(rsSab("SSITICUIDX"))
    arrRoles_Info(K) = Trim(rsSab("SSITICINFO"))
    arrRoles_Droits(K) = Space(nbDroits)
    rsSab.MoveNext
Loop
'__________________________________________________________________

xSQL = "select SSITICUIDD  from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
     & " where SSITICNAT = 'D' order by SSITICUIDD desc FETCH FIRST 1 ROWS ONLY"
Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    arrDroits_Nb = 0
Else
    arrDroits_Nb = rsSab(0)
End If
ReDim arrDroits_UNOM(arrDroits_Nb + 1), arrDroits_Info(arrDroits_Nb + 1)

xSQL = "select *  from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
     & " where SSITICNAT = 'D'" _
     & " order by SSITICUIDD"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    K = rsSab("SSITICUIDD")
    arrDroits_UNOM(K) = Trim(rsSab("SSITICUNOM"))
    rsSab.MoveNext
Loop
'__________________________________________________________________

xSQL = "select SSITICUIDD  from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
     & " where SSITICNAT = ' ' order by SSITICUIDD desc FETCH FIRST 1 ROWS ONLY"
Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    arrUsers_Nb = 0
Else
    arrUsers_Nb = rsSab(0)
End If
ReDim oldUsers(arrUsers_Nb + 1), newUsers(arrUsers_Nb + 1)

xSQL = "select *  from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
     & " where SSITICNAT = ' '" _
     & " order by SSITICUIDD"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    K = rsSab("SSITICUIDD")
    Call rsYSSITIC0_GetBuffer(rsSab, oldUsers(K))
    newUsers(K) = oldUsers(K)
    newUsers(K).SSITICINFO = Space(nbRoles)
    rsSab.MoveNext
Loop
 '____________________________________________________________________
Call cmdUpdate_Init
Open mFile For Input As 1

Do Until EOF(1)
    If blnInput_Skip Then
        blnInput_Skip = False
    Else
        Line Input #1, xIn
        xIn = Trim(xIn)
    End If
    If xIn <> "" Then
        If Not blnFile_Ok Then
            If InStr(1, xIn, "Gestion des rôles") > 0 Then blnFile_Ok = True
        Else
            If Not blnRoles_Ok Then
                If InStr(1, xIn, "Tous/Aucun") > 0 Then
  '=======================================================================================
                   blnRoles_Ok = True
                    blnEnd = False
                    Do Until blnEnd
                        Line Input #1, xIn
                        If Trim(xIn) = "" Then
                            blnEnd = True
                        Else
                            K = 0
                            Call Space_Scan(xIn, K)
                            X = Space_Scan(xIn, K)
                            blnExiste = False
                            For K = 1 To arrRoles_Nb
                                If arrRoles_UIDX(K) = X Then blnExiste = True: Exit For
                            Next K
                            If Not blnExiste Then
                                arrRoles_Nb = arrRoles_Nb + 1
                                K = arrRoles_Nb
                                ReDim Preserve arrRoles_UIDX(arrRoles_Nb + 1), arrRoles_Droits(arrRoles_Nb + 1) As String _
                                             , arrRoles_Info(arrRoles_Nb + 1), arrRoles_Ordre(arrRoles_Nb + 1)

                                arrRoles_UIDX(arrRoles_Nb) = X
                                arrRoles_Info(arrRoles_Nb) = "New"
                                arrRoles_Droits(arrRoles_Nb) = Space(nbDroits)
                            End If
                            arrRoles_Ordre_K = arrRoles_Ordre_K + 1
                            arrRoles_Ordre(K) = arrRoles_Ordre_K
                            
                        End If
                    Loop
                    '=====================
                    arrRoles_Ordre_K = 0
                    '=====================
                End If
            Else
                If Not blnUsers_Ok Then
                    If InStr(1, xIn, "Prénom") > 0 Then blnUsers_Ok = True: blnDroits_Ok = False
 '=======================================================================================
                    If blnUsers_Ok Then
                        blnEnd = False: mUsers = ""
                        Do Until blnEnd
                            Line Input #1, xIn
                            xIn = Trim(xIn) & " "
                            If xIn = " " Then
                                blnEnd = True
                            Else
                                If InStr(1, xIn, "Prénom") = 0 And InStr(1, xIn, "None") = 0 Then
                                    K = 0
                                    wSSITICUNOM = xIn
                                    Call Space_Scan(xIn, K)
                                    Call Space_Scan(xIn, K)
                                    X = Space_Scan(xIn, K)
                                    Call Space_Scan(xIn, K)
                                    xDFin = Space_Scan(xIn, K)
                                    blnExiste = False
                                    
                                    For K = 1 To arrUsers_Nb
                                        If newUsers(K).SSITICUIDX = X Then blnExiste = True: Exit For
                                    Next K
                                    If Not blnExiste Then
                                        arrUsers_Nb = arrUsers_Nb + 1
                                        K = arrUsers_Nb
                                        ReDim Preserve oldUsers(arrUsers_Nb + 1), newUsers(arrUsers_Nb + 1)
                                        newUsers(arrUsers_Nb) = usrYSSITIC0
                                        newUsers(arrUsers_Nb).SSITICUIDX = X
                                        newUsers(arrUsers_Nb).SSITICUIDD = arrUsers_Nb
                                        newUsers(arrUsers_Nb).SSITICYFCT = "New"
                                        newUsers(arrUsers_Nb).SSITICINFO = Space(nbRoles)
                                   End If
                                    newUsers(K).SSITICUNOM = wSSITICUNOM
                                    
                                    If xDFin = "" Then
                                        newUsers(K).SSITICPRFK = " "
                                    Else
                                        If DSys > Mid$(xDFin, 7, 4) & Mid$(xDFin, 4, 2) & Mid$(xDFin, 1, 2) Then
                                            newUsers(K).SSITICPRFK = "X"
                                        Else
                                            newUsers(K).SSITICPRFK = " "
                                        End If
    
                                    End If
                                    
                                    mUsers = mUsers & " " & K
                                End If
                            End If
                        Loop
                    End If
                Else
                    If Not blnDroits_Ok Then
                        If InStr(1, xIn, "Les droits de l'application") > 0 Then
 '=======================================================================================
                            blnUsers_Ok = False: blnDroits_Ok = True
                            blnEnd = False
                            Line Input #1, xIn
                            arrRoles_Ordre_K = arrRoles_Ordre_K + 1
                            For arrRoles_K = 1 To arrRoles_Nb
                                If arrRoles_Ordre(arrRoles_K) = arrRoles_Ordre_K Then
                                    'arrRoles_Droits(arrRoles_K) = Space(nbDroits)
                                    Exit For
                                End If
                            Next arrRoles_K
                            
                            K = 0
                            Do
                                K2 = Val(Space_Scan(mUsers, K))
                                If K2 = 0 Then
                                    blnEnd = True
                                Else
                                    Mid$(newUsers(K2).SSITICINFO, arrRoles_K, 1) = "R"
                                End If
                            Loop Until blnEnd
                            
                            blnEnd = False
                            Do Until blnEnd
                                Line Input #1, xIn
                                xIn = Trim(xIn)
                                If xIn <> "" Then
                                    If InStr(1, xIn, "Prénom") > 0 Then
 '=======================================================================================
                                        blnInput_Skip = True
                                        blnEnd = True
                                    Else
                                        If InStr(1, xIn, "Copyright Athic") > 0 Then Exit Do
  '=======================================================================================
                                       blnExiste = False
                                        For K = 1 To arrDroits_Nb
                                            If arrDroits_UNOM(K) = xIn Then blnExiste = True: Exit For
                                        Next K
                                        If Not blnExiste Then
                                            arrDroits_Nb = arrDroits_Nb + 1
                                            K = arrDroits_Nb
                                            ReDim Preserve arrDroits_UNOM(arrDroits_Nb + 1), arrDroits_Info(arrDroits_Nb + 1)
            
                                            arrDroits_UNOM(arrDroits_Nb) = xIn
                                            arrDroits_Info(arrDroits_Nb) = "New"
                                        End If
                                        Mid$(arrRoles_Droits(arrRoles_K), K, 1) = "D"
                                    End If
                                End If
                            Loop
                        End If
                    End If
                    '
                End If
                '=============
            End If
        End If
    End If
Loop




If Not blnFile_Ok Then V = "Ce  fichier n'est pas conforme à ATHIC : Gestion des rôles": GoTo Error_MsgBox
'------------------------------------------
' Roles
'------------------------------------------

For K = 1 To arrRoles_Nb
    If RTrim(arrRoles_Droits(K)) <> arrRoles_Info(K) Then
        Call cmdUpdate_Init
        
        If arrRoles_Info(K) = "New" Then
            mYSSITIC0_Update = "New"
            newYSSITIC0 = usrYSSITIC0
            newYSSITIC0.SSITICNAT = "R"
            newYSSITIC0.SSITICUIDX = arrRoles_UIDX(K)
            newYSSITIC0.SSITICUIDD = K
            newYSSITIC0.SSITICINFO = arrRoles_Droits(K)
            newYSSITIC0.SSITICUNOM = newYSSITIC0.SSITICUIDX
            Call cmdSSIJRN_TIC("") '("<X:" & RTrim(newYSSITIC0.SSITICINFO) & ">")
            Call cmdUpdate
        Else
            xSQL = "select *  from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
                 & " where SSITICNAT = 'R' and SSITICUIDX = '" & arrRoles_UIDX(K) & "'"
            Set rsSab = cnsab.Execute(xSQL)
            If Not rsSab.EOF Then
                mYSSITIC0_Update = "Update+H"
                Call rsYSSITIC0_GetBuffer(rsSab, oldYSSITIC0)
                newYSSITIC0 = oldYSSITIC0
                newYSSITIC0.SSITICINFO = arrRoles_Droits(K)
                newYSSITIC0.SSITICYFCT = "MOD"
                newYSSITIC0.SSITICYUSR = usrName_UCase
                newYSSITIC0.SSITICYAMJ = mImport_PRFD
                newYSSITIC0.SSITICYHMS = mImport_PRFH
                Call cmdSSIJRN_TIC("") '("<X:" & RTrim(newYSSITIC0.SSITICINFO) & ">")
               Call cmdUpdate
            End If
        End If
    End If
Next K
'------------------------------------------
'Droits
'------------------------------------------
For K = 1 To arrDroits_Nb
    If arrDroits_Info(K) = "New" Then
        Call cmdUpdate_Init
        
        mYSSITIC0_Update = "New"
        newYSSITIC0 = usrYSSITIC0
        newYSSITIC0.SSITICNAT = "D"
        newYSSITIC0.SSITICUIDX = K
        newYSSITIC0.SSITICUIDD = K
        newYSSITIC0.SSITICINFO = ""
        newYSSITIC0.SSITICUNOM = arrDroits_UNOM(K)
        Call cmdSSIJRN_TIC("<X:" & RTrim(newYSSITIC0.SSITICINFO) & ">")
        Call cmdUpdate
    End If
Next K


'------------------------------------------
'PRFX
'------------------------------------------

xSQL = "select SSITICUIDD  from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
     & " where SSITICNAT = '$' order by SSITICUIDD desc FETCH FIRST 1 ROWS ONLY"
Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    arrPRFX_Nb = 0
Else
    arrPRFX_Nb = rsSab(0)
End If
ReDim arrPRFX(arrPRFX_Nb + 1)

xSQL = "select *  from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
     & " where SSITICNAT = '$'" _
     & " order by SSITICUIDD"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    K = rsSab("SSITICUIDD")
    Call rsYSSITIC0_GetBuffer(rsSab, arrPRFX(K))
    rsSab.MoveNext
Loop


For K = 1 To arrUsers_Nb
    blnExiste = False
     X = RTrim(newUsers(K).SSITICINFO)
     If X <> "" Then
        For K2 = 1 To arrPRFX_Nb
            If arrPRFX(K2).SSITICINFO = X Then
               newUsers(K).SSITICPRFX = arrPRFX(K2).SSITICUIDX
               blnExiste = True: Exit For
           End If
        Next K2
        If Not blnExiste Then
           Call cmdUpdate_Init
            arrPRFX_Nb = arrPRFX_Nb + 1
            ReDim Preserve arrPRFX(arrDroits_Nb + 1)
            arrPRFX(arrPRFX_Nb).SSITICINFO = X
            
            mYSSITIC0_Update = "New"
            newYSSITIC0 = usrYSSITIC0
            newYSSITIC0.SSITICNAT = "$"
            newYSSITIC0.SSITICUIDD = arrPRFX_Nb
            Select Case arrPRFX_Nb
               Case 1: newYSSITIC0.SSITICUIDX = "Support"
               Case 2: newYSSITIC0.SSITICUIDX = "Athic"
               Case 3: newYSSITIC0.SSITICUIDX = "Athic +"
               Case 4: newYSSITIC0.SSITICUIDX = "Admin"
               Case 5: newYSSITIC0.SSITICUIDX = "GDMP +"
               Case 6: newYSSITIC0.SSITICUIDX = "GDMP"
               Case 7: newYSSITIC0.SSITICUIDX = "Contrôle"
               Case 8: newYSSITIC0.SSITICUIDX = "Admin BIA"
               Case 9: newYSSITIC0.SSITICUIDX = "Consultation"
               Case Else: newYSSITIC0.SSITICUIDX = "Athic_" & arrPRFX_Nb
            End Select
            newYSSITIC0.SSITICINFO = X
            For K2 = 1 To Len(X)
               If Mid$(X, K2, 1) = "R" Then newYSSITIC0.SSITICUNOM = newYSSITIC0.SSITICUNOM & arrRoles_UIDX(K2) & ","
            Next K2
            If Len(newYSSITIC0.SSITICUNOM) > 128 Then newYSSITIC0.SSITICUNOM = Mid$(newYSSITIC0.SSITICUNOM, 1, 128)
            
            Call cmdSSIJRN_TIC("<X:" & RTrim(newYSSITIC0.SSITICINFO) & ">")
            Call cmdUpdate
            
            arrPRFX(arrPRFX_Nb) = newYSSITIC0
            newUsers(K).SSITICPRFX = newYSSITIC0.SSITICUIDX
        End If
    End If
Next K

'------------------------------------------
'Users
'------------------------------------------
mImport_Nb = arrUsers_Nb
For K = 1 To arrUsers_Nb
     X = RTrim(newUsers(K).SSITICINFO)
     If X <> "" Then
        Call cmdUpdate_Init
   
        If oldUsers(K).SSITICPRFK = "?" Then newUsers(K).SSITICPRFK = "?" ': mYSSITIC0_Update = "Update"
                
        If oldUsers(K).SSITICINFO = RTrim(newUsers(K).SSITICINFO) _
        And oldUsers(K).SSITICUNOM = Trim(newUsers(K).SSITICUNOM) _
        And oldUsers(K).SSITICPRFX = newUsers(K).SSITICPRFX _
        And oldUsers(K).SSITICPRFK = newUsers(K).SSITICPRFK Then
        
            mImport_Ok = mImport_Ok + 1
            xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
            & " where SSIDOMNAT = ' '" _
            & " and SSIDOMDIDX = 'TIC'" _
            & " and SSIDOMUIDX = '" & newUsers(K).SSITICUIDX & "'" _
            & " and SSIDOMUIDD = " & newUsers(K).SSITICUIDD
'$JPL 2015-11-17
'            & " and SSIDOMUIDX = '" & newYSSITIC0.SSITICUIDX & "'" _
'            & " and SSIDOMUIDD = " & newYSSITIC0.SSITICUIDD
            
            Set rsSab = cnsab.Execute(xSQL)
            If Not rsSab.EOF Then
                Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
                newYSSIDOM0 = oldYSSIDOM0
                newYSSIDOM0.SSIDOMPRFD = mImport_PRFD
                newYSSIDOM0.SSIDOMPRFH = mImport_PRFH
                mYSSIDOM0_Update = "Update"
            End If
            
        Else
            
            If newUsers(K).SSITICYFCT = "New" Then
                mImport_New = mImport_New + 1
                mYSSITIC0_Update = "New"
                newYSSITIC0 = newUsers(K)
                newYSSITIC0.SSITICPRFK = "?"
                newYSSITIC0.SSITICYFCT = "CRE"
                Call cmdSSIJRN_TIC("<X:" & RTrim(newYSSITIC0.SSITICPRFX) & ">")
                Call cmdUpdate
            Else
                mImport_Update = mImport_Update + 1
               mYSSITIC0_Update = "Update+H"
                oldYSSITIC0 = oldUsers(K)
                newYSSITIC0 = newUsers(K)
                newYSSITIC0.SSITICYFCT = "MOD"
                newYSSITIC0.SSITICYUSR = usrName_UCase
                newYSSITIC0.SSITICYAMJ = mImport_PRFD
                newYSSITIC0.SSITICYHMS = mImport_PRFH
                
                xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
                & " where SSIDOMNAT = ' '" _
                & " and SSIDOMDIDX = 'TIC'" _
                & " and SSIDOMUIDX = '" & newYSSITIC0.SSITICUIDX & "'" _
                & " and SSIDOMUIDD = " & newYSSITIC0.SSITICUIDD
                
                Set rsSab = cnsab.Execute(xSQL)
                If Not rsSab.EOF Then
                    If Trim(rsSab("SSIDOMPRFX")) <> newUsers(K).SSITICPRFX Then
                        If newUsers(K).SSITICPRFK = " " Then newYSSITIC0.SSITICPRFK = "N"
                    End If
                    
                    If rsSab("SSIDOMPRFK") = newYSSITIC0.SSITICPRFK And rsSab("SSIDOMPRFD") = mImport_PRFD And rsSab("SSIDOMPRFH") = mImport_PRFH Then
                    Else
                        Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
                        newYSSIDOM0 = oldYSSIDOM0
                        newYSSIDOM0.SSIDOMPRFK = newYSSITIC0.SSITICPRFK
                        newYSSIDOM0.SSIDOMPRFD = mImport_PRFD
                        newYSSIDOM0.SSIDOMPRFH = mImport_PRFH
                        mYSSIDOM0_Update = "Update"
                    End If
                End If
                
                Call cmdSSIJRN_TIC("<X:" & RTrim(newYSSITIC0.SSITICPRFX) & ">")
                
                Call cmdUpdate
            End If
        End If
    End If
Next K
'_____________________________________________________________________________________
' Comptes supprimés
'_____________________________________________________________________________________


xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0," & paramIBM_Library_SABSPE & ".YSSITIC0" _
     & " where SSIDOMNAT = ' ' and SSIDOMDIDX = 'TIC' and SSIDOMPRFK <> 'X'" _
     & " and SSIDOMPRFD < " & mImport_PRFD _
     & " and SSITICNAT = SSIDOMNAT and SSITICUIDX = SSIDOMUIDX and SSITICUIDD = SSIDOMUIDD"

Set rsSab_X = cnsab.Execute(xSQL)

Do While Not rsSab_X.EOF
    mImport_Ann = mImport_Ann + 1
    Call cmdUpdate_Init
    Call rsYSSIDOM0_GetBuffer(rsSab_X, oldYSSIDOM0)
    newYSSIDOM0 = oldYSSIDOM0
    newYSSIDOM0.SSIDOMPRFK = "X"
    newYSSIDOM0.SSIDOMYFCT = "SUP"
    newYSSIDOM0.SSIDOMYUSR = usrName_UCase
    newYSSIDOM0.SSIDOMYAMJ = mImport_PRFD
    newYSSIDOM0.SSIDOMYHMS = mImport_PRFH
    mYSSIDOM0_Update = "Update+H"
    
    Call rsYSSITIC0_GetBuffer(rsSab_X, oldYSSITIC0)
    newYSSITIC0 = oldYSSITIC0
    newYSSITIC0.SSITICPRFK = "X"
    newYSSITIC0.SSITICYFCT = "SUP"
    newYSSITIC0.SSITICYUSR = usrName_UCase
    newYSSITIC0.SSITICYAMJ = mImport_PRFD
    newYSSITIC0.SSITICYHMS = mImport_PRFH
    newYSSITIC0.SSITICINFO = newYSSITIC0.SSITICINFO & "Supprimé"
    mYSSITIC0_Update = "Update+H"
    Call cmdSSIJRN_TIC("<PRFX: |" & newYSSITIC0.SSITICPRFX & ">")
    Call cmdUpdate
      
    rsSab_X.MoveNext
Loop
'_____________________________________________________________________________________

'___________________________________________________________________________________________________________________________

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

Close
Set rsSab = Nothing

Exit Sub


End Sub

Private Sub cmdSelect_SQL_9_SAA_Profil()
Dim V, X As String, xIn As String
Dim xSQL As String, xWhere As String, Nb As Long, K As Integer
Dim blnProfil As Boolean, blnFunction As Boolean
Dim wSSISAAUIDD As Long
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_9_SAA_Profil"
Call lstErr_Clear(lstErr, cmdContext, "> " & currentAction & " ........"): DoEvents

 Call paramSAA_Load
 
 xSQL = "select SSISAAUIDD from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
     & " where SSISAANAT = '$'" _
     & " order by SSISAAUIDD desc"

Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then wSSISAAUIDD = rsSab("SSISAAUIDD")


 '____________________________________________________________________
Call cmdUpdate_Init
Call rsYSSISAA0_Init(xYSSISAA0)
xYSSISAA0.SSISAANAT = "$"

Open mFile For Input As 1

blnProfil = False
Do Until EOF(1)
    Line Input #1, xIn
    xIn = Trim(xIn)
    If xIn <> "" Then
    
        I = InStr(1, xIn, "Number of entries:")
        If I > 0 Then
            mImport_Nb = Val(Mid$(xIn, I + 18, Len(xIn) - I - 17))
        End If
        
        I = InStr(1, xIn, "Name        =")
        If I > 0 Then
            If blnProfil Then
                 Call cmdSelect_SQL_9_SAA_Profil_Update
                 mImport_In = mImport_In + 1
                 blnProfil = False
            End If
            
            blnProfil = True: blnFunction = False
            ReDim arrSAA_Function(arrSAA_Function_Nb), blnSAA_Function(arrSAA_Function_Nb)

            xYSSISAA0.SSISAAUIDX = Mid$(xIn, I + 14, Len(xIn) - I - 13)
            xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
                 & " where SSISAANAT = '$' and SSISAAUIDX = '" & xYSSISAA0.SSISAAUIDX & "'" _
                 & " and SSISAAUSEQ = 0"
            
            Set rsSab = cnsab.Execute(xSQL)
            If Not rsSab.EOF Then
                mYSSISAA0_Update = "Update+H"
                Call rsYSSISAA0_GetBuffer(rsSab, oldYSSISAA0)
                newYSSISAA0 = oldYSSISAA0
                newYSSISAA0.SSISAAYFCT = "MOD"
                newYSSISAA0.SSISAAINFO = ""

            Else
                mYSSISAA0_Update = "New"
                newYSSISAA0 = xYSSISAA0
                newYSSISAA0.SSISAAYFCT = "CRE"
                wSSISAAUIDD = wSSISAAUIDD + 1
                newYSSISAA0.SSISAAUIDD = wSSISAAUIDD
            End If
        End If
        
        If blnProfil Then
            I = InStr(1, xIn, "Application =")
            If I > 0 Then
                X = Mid$(xIn, I + 14, Len(xIn) - I - 13)
                For arrSAA_App_K = 1 To arrSAA_App_Nb
                    If X = arrSAA_App_Code(arrSAA_App_K) Then Exit For
                Next arrSAA_App_K
            Else
                I = InStr(1, xIn, "Function =")
                If I > 0 Then
                    X = Mid$(xIn, I + 11, Len(xIn) - I - 10)
                    For arrSAA_Function_K = 1 To arrSAA_Function_Nb
                        If X = arrSAA_Function_Code(arrSAA_Function_K) Then blnFunction = True: blnSAA_Function(arrSAA_Function_K) = True: Exit For
                    Next arrSAA_Function_K
                Else
                    If blnFunction Then
                        arrSAA_Function(arrSAA_Function_K) = arrSAA_Function(arrSAA_Function_K) & xIn & vbCrLf
                    End If
                End If
            End If
            
                
        End If
    End If
Loop

If blnProfil Then
     Call cmdSelect_SQL_9_SAA_Profil_Update
     mImport_In = mImport_In + 1
End If
            

Close

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

Set rsSab = Nothing
Call paramSAA_Load

Exit Sub


End Sub

Private Sub cmdSelect_SQL_9_DIV_UGM_Load()
Dim V, X As String, xIn As String, blnSSIDIVUIDX As Boolean
Dim K1 As Integer, K2 As Integer, wBadge As Long
Dim kDEB As Integer, kFIN As Integer, wDEB As String, wFIN As String
Dim arrBadge_K(6) As Integer, arrBadge_Id(6) As Long
Dim arrYSSIDIV0() As typeYSSIDIV0, arrYSSIDIV0_Nb As Integer
Dim blnUpdate As Boolean, blnUIDX_XXXX As Boolean

Dim xSQL As String, xWhere As String, Nb As Long, K As Integer
Dim wSSIDIVUIDX As String
On Error GoTo Error_Handler
currentAction = "cmdSelect_SQL_9_DIV_UGM_Load"
Call lstErr_Clear(lstErr, cmdContext, "> " & currentAction & " ........"): DoEvents

Call rsYSSIDIV0_Init(xYSSIDIV0)
xYSSIDIV0.SSIDIVDIDK = "UGM"

newYSSIDIV0 = xYSSIDIV0
Call cmdUpdate_Init


xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
     & " where SSIDIVNAT = '$' and SSIDIVDIDK = 'UGM' and SSIDIVSTAK = ' '"
Set rsSab = cnsab.Execute(xSQL)
ReDim arrYSSIDIV0(rsSab(0) + 1)
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
     & " where SSIDIVNAT = '$' and SSIDIVDIDK = 'UGM' and SSIDIVSTAK = ' '" _
     & " order by SSIDIVUIDD"

Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrYSSIDIV0_Nb = arrYSSIDIV0_Nb + 1
    Call rsYSSIDIV0_GetBuffer(rsSab, arrYSSIDIV0(arrYSSIDIV0_Nb))
    rsSab.MoveNext
Loop

 '____________________________________________________________________
Call cmdUpdate_Init
Call rsYSSIDIV0_Init(xYSSIDIV0)
xYSSIDIV0.SSIDIVDIDK = "UGM"
'''xYSSIDIV0.SSIDIVNAT = "$"


Open mFile For Input As 1

Do Until EOF(1)
    Line Input #1, xIn
    xIn = Trim(xIn)
    If xIn <> "" Then
        If Not blnSSIDIVUIDX Then
            If InStr(xIn, "Identifiant") > 0 Then
                blnSSIDIVUIDX = True
                kDEB = InStr(xIn, "Date Début")
                kFIN = InStr(xIn, "Date Fin")
                arrBadge_K(1) = InStr(xIn, "Badge 1")
                If arrBadge_K(1) <= 0 Then V = "Erreur : InStr(xIn, ""Badge 1"")": GoTo Error_MsgBox
                arrBadge_K(2) = InStr(xIn, "Badge 2")
                If arrBadge_K(2) <= 0 Then V = "Erreur : InStr(xIn, ""Badge 2"")": GoTo Error_MsgBox
                arrBadge_K(3) = InStr(xIn, "Badge 3")
                If arrBadge_K(3) <= 0 Then V = "Erreur : InStr(xIn, ""Badge 3"")": GoTo Error_MsgBox
                arrBadge_K(4) = InStr(xIn, "Badge 4")
                If arrBadge_K(4) <= 0 Then V = "Erreur : InStr(xIn, ""Badge 4"")": GoTo Error_MsgBox
                arrBadge_K(5) = InStr(xIn, "Badge 5")
                If arrBadge_K(5) <= 0 Then V = "Erreur : InStr(xIn, ""Badge 5"")": GoTo Error_MsgBox
                arrBadge_K(6) = InStr(xIn, "Badge 6")
                If arrBadge_K(6) <= 0 Then V = "Erreur : InStr(xIn, ""Badge 6"")": GoTo Error_MsgBox
                
                K2 = InStr(xIn, "Lecteurs")
                If K2 <= 0 Then V = "Erreur : K2 = InStr(xIn, ""Lecteurs"")": GoTo Error_MsgBox
            End If
        Else
            mImport_In = mImport_In + 1
            K = InStr(xIn, " ")
            wSSIDIVUIDX = Mid$(xIn, 1, K - 1)
            If Len(wSSIDIVUIDX) > 15 Then wSSIDIVUIDX = Mid$(wSSIDIVUIDX, 1, 15)
            Call dateJma10_Amj(Mid$(xIn, kDEB, 10), wDEB)
            Call dateJma10_Amj(Mid$(xIn, kFIN, 10), wFIN)
            xYSSIDIV0.SSIDIVPRFK = " "
            If wDEB > DSys Then
                xYSSIDIV0.SSIDIVPRFK = "X"
            Else
                If wFIN < DSys Then xYSSIDIV0.SSIDIVPRFK = "X"
            End If
            
            For K = 1 To 6
                V = Val(Mid$(xIn, arrBadge_K(K), 15))
                arrBadge_Id(K) = IIf(V < 99999, V, 0)
            Next K
            X = Replace(xIn, "                      ", "|")
            X = Replace(X, "                     ", "|")
            X = Replace(X, "                    ", "|")
            X = Replace(X, "                   ", "|")
            X = Replace(X, "                  ", "|")
            X = Replace(X, "                 ", "|")
            X = Replace(X, "                ", "|")
            X = Replace(X, "               ", "|")
            X = Replace(X, "              ", "|")
            X = Replace(X, "            ", "|")
            X = Replace(X, "           ", "|")
            X = Replace(X, "          ", "|")
            X = Replace(X, "         ", "|")
            X = Replace(X, "        ", "|")
            X = Replace(X, "       ", "|")
            
            xYSSIDIV0.SSIDIVINFO = X
           
            If Len(xIn) < K2 Then
                Call MsgBox(xIn, vbCritical, "cmdSelect_SQL_9_DIV_UGM_Load : aucun lecteur affecté")
                Call cmdUpdate_Init
                mYSSITXT0_JRN_Update = "New"
                Call rsYSSITXT0_Init(newYSSITXT0_JRN)
                newYSSITXT0_JRN.SSITXTNAT = "J"
                newYSSITXT0_JRN.SSITXTDIDX = "DIV"
                newYSSITXT0_JRN.SSITXTUIDX = wSSIDIVUIDX
                newYSSITXT0_JRN.SSITXTINFO = "<ORIG:50><Y:DIV| |" & wSSIDIVUIDX & ">" _
                                         & "<UID: - " & wSSIDIVUIDX & ">" _
                                         & "<FCT:???><X:aucun lecteur affecté " & xIn
                
                newYSSITXT0_JRN.SSITXTYAMJ = DSys
                newYSSITXT0_JRN.SSITXTYHMS = time_Hms
                newYSSITXT0_JRN.SSITXTYUSR = usrName_UCase
                Call cmdUpdate
            Else
                X = Mid$(xIn, K2, Len(xIn) - K2 + 1)
                For K = 1 To arrYSSIDIV0_Nb
                    If InStr(X, arrYSSIDIV0(K).SSIDIVPRFX) > 0 Then
                       If arrYSSIDIV0(K).SSIDIVUIDX = "UGM_FOTC clavier" Then
                            blnUIDX_XXXX = True
                            For K1 = 1 To 6
                                If arrBadge_Id(K1) > 0 Then
                                    xYSSIDIV0.SSIDIVINFO = Replace(xYSSIDIV0.SSIDIVINFO, arrBadge_Id(K1), "XXXX")
                                End If
                            Next K1
                       Else
                            blnUIDX_XXXX = False
                       End If
                       
                       For K1 = 1 To 6
                            If arrBadge_Id(K1) > 0 Then
                                If blnUIDX_XXXX Then
                                    xYSSIDIV0.SSIDIVUIDX = wSSIDIVUIDX & "-XXXX"
                                    'xYSSIDIV0.SSIDIVINFO = Replace(xYSSIDIV0.SSIDIVINFO, arrBadge_Id(K1), "XXXX")
                                Else
                                    xYSSIDIV0.SSIDIVUIDX = wSSIDIVUIDX & "-" & arrBadge_Id(K1)
                                End If
                               xYSSIDIV0.SSIDIVUIDD = arrYSSIDIV0(K).SSIDIVUIDD
                                xYSSIDIV0.SSIDIVPRFX = arrYSSIDIV0(K).SSIDIVUIDX
                                xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
                                    & " where SSIDIVNAT = ' ' and SSIDIVUIDX = '" & xYSSIDIV0.SSIDIVUIDX & "'" _
                                    & "  and SSIDIVUIDD = " & xYSSIDIV0.SSIDIVUIDD & " and SSIDIVDIDK = 'UGM'"

                                Set rsSab = cnsab.Execute(xSQL)
                                If rsSab.EOF Then
                                    mImport_New = mImport_New + 1
                                    mYSSIDIV0_Update = "New"
                                    newYSSIDIV0 = xYSSIDIV0
                                    newYSSIDIV0.SSIDIVSTAK = " "
                                    newYSSIDIV0.SSIDIVPRFK = "?"
                                    newYSSIDIV0.SSIDIVUNOM = wSSIDIVUIDX
                                    newYSSIDIV0.SSIDIVYFCT = "CRE"
                                    newYSSIDIV0.SSIDIVYAMJ = DSys
                                    newYSSIDIV0.SSIDIVYHMS = time_Hms
                                    newYSSIDIV0.SSIDIVYUSR = usrName_UCase
                                    newYSSIDIV0.SSIDIVINFO = xYSSIDIV0.SSIDIVINFO
                                    Call cmdSSIJRN_DIV("<X:Création " & xYSSIDIV0.SSIDIVUIDX & ">")
                                   Call cmdUpdate
                                Else
                                    Call rsYSSIDIV0_GetBuffer(rsSab, oldYSSIDIV0)
                                    If xYSSIDIV0.SSIDIVINFO <> Trim(rsSab("SSIDIVINFO")) Then
                                        mImport_Update = mImport_Update + 1
                                        blnUpdate = True
                                    Else
                                        mImport_Ok = mImport_Ok + 1
                                        blnUpdate = False
                                        If xYSSIDIV0.SSIDIVPRFK <> rsSab("SSIDIVPRFK") _
                                        And rsSab("SSIDIVPRFK") <> "?" Then blnUpdate = True
                                    End If
                                    
                                    If blnUpdate Then
                                        
                                        mYSSIDIV0_Update = "Update+H"
                                        oldYSSIDIVH = oldYSSIDIV0
                                        newYSSIDIV0 = oldYSSIDIV0
                                       '' newYSSIDIV0.SSIDIVUIDX = xYSSIDIV0.SSIDIVUIDX
                                       '' newYSSIDIV0.SSIDIVUNOM = xYSSIDIV0.SSIDIVUNOM
                                        newYSSIDIV0.SSIDIVPRFX = xYSSIDIV0.SSIDIVPRFX
                                        newYSSIDIV0.SSIDIVINFO = xYSSIDIV0.SSIDIVINFO
                                        newYSSIDIV0.SSIDIVYFCT = "CTL"
                                        newYSSIDIV0.SSIDIVYUSR = usrName_UCase
                                        newYSSIDIV0.SSIDIVYAMJ = DSys
                                        newYSSIDIV0.SSIDIVYHMS = time_Hms
                                        
                                        If oldYSSIDIV0.SSIDIVPRFK <> "?" Then
                                            If oldYSSIDIV0.SSIDIVPRFK <> xYSSIDIV0.SSIDIVPRFK Then
                                                newYSSIDIV0.SSIDIVPRFK = xYSSIDIV0.SSIDIVPRFK
                                                xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
                                                    & " Where SSIDOMNAT = ' ' and SSIDOMDIDX = 'DIV'" _
                                                    & " and SSIDOMUIDX = '" & oldYSSIDIV0.SSIDIVUIDX & "'" _
                                                    & " and SSIDOMUIDD = " & oldYSSIDIV0.SSIDIVUIDD _
                                                    & " and SSIUSRNAT = ' ' and SSIUSRUIDN = SSIDOMUIDN"
                                
                                                Set rsSab = cnsab.Execute(xSQL)
                                            
                                                If Not rsSab.EOF Then
                                                    If rsSab("SSIDOMPRFK") <> newYSSIDIV0.SSIDIVPRFK Then
                                                        Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
                                                        mYSSIDOM0_Update = "Update"
                                                        newYSSIDOM0 = oldYSSIDOM0
                                                        newYSSIDOM0.SSIDOMPRFK = newYSSIDIV0.SSIDIVPRFK
                                                        If rsSab("SSIUSRSTAK") = "N" Then newYSSIDOM0.SSIDOMSTAK = "N"
                                                    End If
                                                End If
                                            End If
                                        End If
                                        
                                        Call cmdSSIJRN_DIV("<X:Contrôle " & xYSSIDIV0.SSIDIVUIDX & ">")
                                        Call cmdUpdate
                                    End If
                                    
                                    Call cmdUpdate_Init
        '______________________________________________________________________________________________
                                    If oldYSSIDIV0.SSIDIVPRFK <> "?" Then
                                        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
                                        & " where SSIDOMNAT = ' '" _
                                        & " and SSIDOMDIDX = 'DIV'" _
                                        & " and SSIDOMUIDX = '" & xYSSIDIV0.SSIDIVUIDX & "'" _
                                        & " and SSIDOMUIDD = " & xYSSIDIV0.SSIDIVUIDD
                                             
                                        Set rsSab = cnsab.Execute(xSQL)
                                        If Not rsSab.EOF Then
                                            Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
                                            newYSSIDOM0 = oldYSSIDOM0
                                            newYSSIDOM0.SSIDOMPRFD = mImport_PRFD
                                            newYSSIDOM0.SSIDOMPRFH = mImport_PRFH
                                            
                                            mYSSIDOM0_Update = "Update"
                                            Call cmdUpdate
                                        Else
                                            Call MsgBox(xSQL, vbCritical, "cmdSelect_SQL_9_DIV_UGM_Load : inconnu")
                                        End If
                                    End If

                                
                                End If

                            End If
                        Next K1

                    End If
                Next K
            End If
        End If
    End If
Loop



Close

FIN:
Call cmdUpdate_Init
Call rsYSSITXT0_Init(newYSSITXT0_JRN)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIDIV0" _
& " where SSIDOMNAT = ' '" _
& " and SSIDOMDIDX = 'DIV' " _
& " and (SSIDOMPRFD <> " & mImport_PRFD & " or SSIDOMPRFH <> " & mImport_PRFH & ")" _
& " and SSIDIVNAT = SSIDOMNAT  and SSIDIVUIDX = SSIDOMUIDX and SSIDIVUIDD = SSIDOMUIDD and SSIDIVDIDK = 'UGM' " _


Set rsSab_X = cnsab.Execute(xSQL)
Do While Not rsSab_X.EOF
    Call rsYSSIDOM0_GetBuffer(rsSab_X, oldYSSIDOM0)
    Call MsgBox(oldYSSIDOM0.SSIDOMUIDX & " - " & oldYSSIDOM0.SSIDOMPRFX, vbCritical, "cmdSelect_SQL_9_DIV_UGM_Load : non contrôlé, absent de la liste UGM.txt ")
    
    mYSSITXT0_JRN_Update = "New"
    newYSSITXT0_JRN.SSITXTNAT = "J"
    newYSSITXT0_JRN.SSITXTDIDX = "DIV"
    newYSSITXT0_JRN.SSITXTUIDX = oldYSSIDOM0.SSIDOMUIDX
    newYSSITXT0_JRN.SSITXTINFO = "<ORIG:50><Y:DIV| |" & Trim(oldYSSIDOM0.SSIDOMUIDX) _
                         & "|" & oldYSSIDOM0.SSIDOMYVER & "|>" _
                         & "<UID:" & oldYSSIDOM0.SSIDOMNAT & " - " & oldYSSIDOM0.SSIDOMUIDX & " - " & oldYSSIDOM0.SSIDOMPRFX & ">" _
                         & "<FCT:???><X:absent de la liste UGM.txt>"
    
    newYSSITXT0_JRN.SSITXTYAMJ = DSys
    newYSSITXT0_JRN.SSITXTYHMS = time_Hms
    newYSSITXT0_JRN.SSITXTYUSR = usrName_UCase
    Call cmdUpdate

    rsSab_X.MoveNext
Loop
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " :" & currentAction
Exit_sub:

Set rsSab = Nothing
Close
Exit Sub


End Sub

Private Sub cmdSelect_SQL_9_DIV_TEREN_Load()
Dim V, X As String, xIn As String, blnSSIDIVUIDX As Boolean
Dim K1 As Integer, K2 As Integer, wBadge As Long
Dim kDEB As Integer, kFIN As Integer, wDEB As String, wFIN As String
Dim arrBadge_K(6) As Integer, arrBadge_Id(6) As Long
Dim arrYSSIDIV0() As typeYSSIDIV0, arrYSSIDIV0_Nb As Integer
Dim blnUpdate As Boolean, blnUIDX_XXXX As Boolean

Dim xSQL As String, xWhere As String, Nb As Long, K As Integer
Dim wSSIDIVUIDX As String, wSSIDIVPRFX As String, wSSIDIVPRFK As String

Dim xTerena As String
Dim lenTerena As Long, blnExit As Boolean, K10 As Long


On Error GoTo Error_Handler
currentAction = "cmdSelect_SQL_9_DIV_TEREN_Load"
Call lstErr_Clear(lstErr, cmdContext, "> " & currentAction & " ........"): DoEvents

Call rsYSSIDIV0_Init(xYSSIDIV0)
xYSSIDIV0.SSIDIVDIDK = "TEREN"

newYSSIDIV0 = xYSSIDIV0
Call cmdUpdate_Init


xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
     & " where SSIDIVNAT = '$' and SSIDIVDIDK = 'TEREN' and SSIDIVSTAK = ' '"
Set rsSab = cnsab.Execute(xSQL)
ReDim arrYSSIDIV0(rsSab(0) + 1)
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
     & " where SSIDIVNAT = '$' and SSIDIVDIDK = 'TEREN' and SSIDIVSTAK = ' '" _
     & " order by SSIDIVUIDD"

Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrYSSIDIV0_Nb = arrYSSIDIV0_Nb + 1
    Call rsYSSIDIV0_GetBuffer(rsSab, arrYSSIDIV0(arrYSSIDIV0_Nb))
    rsSab.MoveNext
Loop

 '____________________________________________________________________
Call cmdUpdate_Init
Call rsYSSIDIV0_Init(xYSSIDIV0)
xYSSIDIV0.SSIDIVDIDK = "TEREN"
'''xYSSIDIV0.SSIDIVNAT = "$"


Open mFile For Input As 1

'xTerena = ""
'Do Until EOF(1)
'    Line Input #1, xIn
'    xTerena = xTerena & xIn
'Loop
'Close 1
'lenTerena = Len(xTerena)
blnExit = False
K10 = 1

'Do Until blnExit
'    K = InStr(K10, xTerena, Chr(10))
'    If K > 0 Then
'        xIn = Mid$(xTerena, K10, K - K10)
'        K10 = K + 1
Do Until EOF(1)
    Line Input #1, xIn

        K = InStr(1, xIn, Chr(9)): xYSSIDIV0.SSIDIVUIDX = Trim(Mid$(xIn, 1, K - 1))
        If xYSSIDIV0.SSIDIVUIDX <> "Identifiant" Then
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)): wSSIDIVPRFK = Mid$(xIn, K1, K - K1)
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)): Call dateJma10_Amj(Mid$(xIn, K1, K - K1), wDEB)
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)): Call dateJma10_Amj(Mid$(xIn, K1, K - K1), wFIN)
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)): wSSIDIVPRFX = Mid$(xIn, K1, 2)
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            If IsNumeric(Mid$(xIn, K1, K - K1)) And Val(Mid$(xIn, K1, K - K1)) <> 0 Then Mid$(xIn, K1, K - K1) = "00000000xxxx"
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            If IsNumeric(Mid$(xIn, K1, K - K1)) And Val(Mid$(xIn, K1, K - K1)) <> 0 Then Mid$(xIn, K1, K - K1) = "00000000xxxx"
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            If IsNumeric(Mid$(xIn, K1, K - K1)) And Val(Mid$(xIn, K1, K - K1)) <> 0 Then Mid$(xIn, K1, K - K1) = "00000000xxxx"
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            If IsNumeric(Mid$(xIn, K1, K - K1)) And Val(Mid$(xIn, K1, K - K1)) <> 0 Then Mid$(xIn, K1, K - K1) = "00000000xxxx"
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            If IsNumeric(Mid$(xIn, K1, K - K1)) And Val(Mid$(xIn, K1, K - K1)) <> 0 Then Mid$(xIn, K1, K - K1) = "00000000xxxx"
            K1 = K + 1: K = InStr(K1, xIn, Chr(9)):
            If IsNumeric(Mid$(xIn, K1, K - K1)) And Val(Mid$(xIn, K1, K - K1)) <> 0 Then Mid$(xIn, K1, K - K1) = "00000000xxxx"
            
            Call cmdUpdate_Init
            xYSSIDIV0.SSIDIVUIDD = 100 '$JPL 2014-11-06
            xYSSIDIV0.SSIDIVINFO = xIn
            xYSSIDIV0.SSIDIVPRFK = " "
            If wSSIDIVPRFK = "0" Then
                xYSSIDIV0.SSIDIVPRFK = "X"
            Else
                xYSSIDIV0.SSIDIVPRFK = " "
                If wDEB > DSys Then
                    xYSSIDIV0.SSIDIVPRFK = "X"
                Else
                    If wFIN < DSys Then xYSSIDIV0.SSIDIVPRFK = "X"
                End If
            End If
            xYSSIDIV0.SSIDIVPRFX = "code accès : " & wSSIDIVPRFX
            For K = 1 To arrYSSIDIV0_Nb
                If arrYSSIDIV0(K).SSIDIVPRFX = wSSIDIVPRFX Then
                    '$JPL 2014-11-06 xYSSIDIV0.SSIDIVUIDD = arrYSSIDIV0(K).SSIDIVUIDD
                    xYSSIDIV0.SSIDIVPRFX = arrYSSIDIV0(K).SSIDIVUIDX
                   Exit For
                End If
            Next K
            
             xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
                 & " where SSIDIVNAT = ' ' and SSIDIVUIDX = '" & xYSSIDIV0.SSIDIVUIDX & "'" _
                 & "  and SSIDIVUIDD = " & xYSSIDIV0.SSIDIVUIDD & " and SSIDIVDIDK = 'TEREN'"

             Set rsSab = cnsab.Execute(xSQL)
             If rsSab.EOF Then
                mImport_New = mImport_New + 1
                mYSSIDIV0_Update = "New"
                newYSSIDIV0 = xYSSIDIV0
                newYSSIDIV0.SSIDIVSTAK = " "
                newYSSIDIV0.SSIDIVPRFK = "?"
                newYSSIDIV0.SSIDIVUNOM = xYSSIDIV0.SSIDIVUIDX
                newYSSIDIV0.SSIDIVYFCT = "CRE"
                newYSSIDIV0.SSIDIVYAMJ = DSys
                newYSSIDIV0.SSIDIVYHMS = time_Hms
                newYSSIDIV0.SSIDIVYUSR = usrName_UCase
                newYSSIDIV0.SSIDIVINFO = xYSSIDIV0.SSIDIVINFO
                Call cmdSSIJRN_DIV("<X:Création " & xYSSIDIV0.SSIDIVUIDX & xYSSIDIV0.SSIDIVUIDD & ">")
                Call cmdUpdate
            Else
                Call rsYSSIDIV0_GetBuffer(rsSab, oldYSSIDIV0)
                If xYSSIDIV0.SSIDIVINFO <> Trim(rsSab("SSIDIVINFO")) _
                Or xYSSIDIV0.SSIDIVPRFX <> Trim(rsSab("SSIDIVPRFX")) Then
                    mImport_Update = mImport_Update + 1
                    blnUpdate = True
                Else
                    mImport_Ok = mImport_Ok + 1
                    blnUpdate = False
                    If xYSSIDIV0.SSIDIVPRFK <> rsSab("SSIDIVPRFK") _
                    And rsSab("SSIDIVPRFK") <> "?" Then blnUpdate = True
                End If
                
                If blnUpdate Then
                    
                    mYSSIDIV0_Update = "Update+H"
                    oldYSSIDIVH = oldYSSIDIV0
                    newYSSIDIV0 = oldYSSIDIV0
                    newYSSIDIV0.SSIDIVPRFX = xYSSIDIV0.SSIDIVPRFX
                    newYSSIDIV0.SSIDIVINFO = xYSSIDIV0.SSIDIVINFO
                    newYSSIDIV0.SSIDIVYFCT = "MAJ"
                    newYSSIDIV0.SSIDIVYUSR = usrName_UCase
                    newYSSIDIV0.SSIDIVYAMJ = DSys
                    newYSSIDIV0.SSIDIVYHMS = time_Hms
                    
                    If oldYSSIDIV0.SSIDIVPRFK <> "?" Then
                        'If oldYSSIDIV0.SSIDIVPRFK <> xYSSIDIV0.SSIDIVPRFK Then
                            newYSSIDIV0.SSIDIVPRFK = xYSSIDIV0.SSIDIVPRFK
                            xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
                                & " Where SSIDOMNAT = ' ' and SSIDOMDIDX = 'DIV'" _
                                & " and SSIDOMUIDX = '" & oldYSSIDIV0.SSIDIVUIDX & "'" _
                                & " and SSIDOMUIDD = " & oldYSSIDIV0.SSIDIVUIDD _
                                & " and SSIUSRNAT = ' ' and SSIUSRUIDN = SSIDOMUIDN"
            
                            Set rsSab = cnsab.Execute(xSQL)
                        
                            If Not rsSab.EOF Then
                                If Trim(rsSab("SSIDOMPRFX")) = Trim(newYSSIDIV0.SSIDIVPRFX) _
                                And rsSab("SSIDOMPRFK") = newYSSIDIV0.SSIDIVPRFK Then
                                Else
                                    Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
                                    mYSSIDOM0_Update = "Update"
                                    newYSSIDOM0 = oldYSSIDOM0
                                    newYSSIDOM0.SSIDOMPRFK = newYSSIDIV0.SSIDIVPRFK
                                    If newYSSIDIV0.SSIDIVPRFK <> "X" Then
                                        If Trim(rsSab("SSIDOMPRFX")) <> Trim(newYSSIDIV0.SSIDIVPRFX) Then newYSSIDOM0.SSIDOMPRFK = "N"
                                    End If
                                    'newYSSIDOM0.SSIDOMPRFX = newYSSIDIV0.SSIDIVPRFX
                                    'If rsSab("SSIUSRSTAK") = "N" Then newYSSIDOM0.SSIDOMSTAK = "N"
                                End If
                            End If
                        'End If
                    End If
                    
                    'Call cmdSSIJRN_DIV("<X:Contrôle " & xYSSIDIV0.SSIDIVUIDX & ">")
                    Call cmdSSIJRN_DIV("")
                    Call cmdUpdate
                End If
                
                Call cmdUpdate_Init
'______________________________________________________________________________________________
                If oldYSSIDIV0.SSIDIVPRFK <> "?" Then
                    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
                    & " where SSIDOMNAT = ' '" _
                    & " and SSIDOMDIDX = 'DIV'" _
                    & " and SSIDOMUIDX = '" & xYSSIDIV0.SSIDIVUIDX & "'" _
                    & " and SSIDOMUIDD = " & xYSSIDIV0.SSIDIVUIDD
                         
                    Set rsSab = cnsab.Execute(xSQL)
                    If Not rsSab.EOF Then
                        Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
                        newYSSIDOM0 = oldYSSIDOM0
                        newYSSIDOM0.SSIDOMPRFD = mImport_PRFD
                        newYSSIDOM0.SSIDOMPRFH = mImport_PRFH
                        
                        mYSSIDOM0_Update = "Update"
                        Call cmdUpdate
                    Else
                        Call MsgBox(xSQL, vbCritical, "cmdSelect_SQL_9_DIV_TEREN_Load : inconnu")
                    End If
                End If

            End If
        End If
    'Else
    '    blnExit = True
    'End If
Loop


FIN:
Close 1

Call cmdUpdate_Init
Call rsYSSITXT0_Init(newYSSITXT0_JRN)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIDIV0" _
& " where SSIDOMNAT = ' '" _
& " and SSIDOMDIDX = 'DIV' " _
& " and (SSIDOMPRFD <> " & mImport_PRFD & " or SSIDOMPRFH <> " & mImport_PRFH & ")" _
& " and SSIDIVNAT = SSIDOMNAT  and SSIDIVUIDX = SSIDOMUIDX and SSIDIVUIDD = SSIDOMUIDD and SSIDIVDIDK = 'TEREN' " _


Set rsSab_X = cnsab.Execute(xSQL)
Do While Not rsSab_X.EOF
    Call rsYSSIDOM0_GetBuffer(rsSab_X, oldYSSIDOM0)
    Call MsgBox(oldYSSIDOM0.SSIDOMUIDX & " - " & oldYSSIDOM0.SSIDOMPRFX, vbCritical, "cmdSelect_SQL_9_DIV_TEREN_Load : non contrôlé, absent de la liste TEREN.txt ")
    
    mYSSITXT0_JRN_Update = "New"
    newYSSITXT0_JRN.SSITXTNAT = "J"
    newYSSITXT0_JRN.SSITXTDIDX = "DIV"
    newYSSITXT0_JRN.SSITXTUIDX = oldYSSIDOM0.SSIDOMUIDX
    newYSSITXT0_JRN.SSITXTINFO = "<ORIG:50><Y:DIV| |" & Trim(oldYSSIDOM0.SSIDOMUIDX) & oldYSSIDOM0.SSIDOMUIDD _
                         & "|" & oldYSSIDOM0.SSIDOMYVER & "|>" _
                         & "<UID:" & oldYSSIDOM0.SSIDOMNAT & " - " & oldYSSIDOM0.SSIDOMUIDX & " - " & oldYSSIDOM0.SSIDOMPRFX & ">" _
                         & "<FCT:???><X:absent de la liste TERENA.txt>"
    
    newYSSITXT0_JRN.SSITXTYAMJ = DSys
    newYSSITXT0_JRN.SSITXTYHMS = time_Hms
    newYSSITXT0_JRN.SSITXTYUSR = usrName_UCase
    Call cmdUpdate

    rsSab_X.MoveNext
Loop
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " :" & currentAction
Exit_sub:

Set rsSab = Nothing
Close
Exit Sub


End Sub



Private Sub cmdSelect_SQL_9_SAA_App()
Dim V, X As String, xIn As String, K As Integer
Dim blnSAA_Function As Boolean
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_9_SAA_App"
Call lstErr_Clear(lstErr, cmdContext, "> " & currentAction & " ........"): DoEvents

 Call paramSAA_Load
 
 '____________________________________________________________________
Call cmdUpdate_Init

Call rsYSSISAA0_Init(newYSSISAA0)
newYSSISAA0.SSISAANAT = "A"
    newYSSISAA0.SSISAAYAMJ = DSys
    newYSSISAA0.SSISAAYHMS = time_Hms
    newYSSISAA0.SSISAAYUSR = usrName_UCase
    
Open mFile For Input As 1

blnSAA_Function = False
Do Until EOF(1)
    Line Input #1, xIn
    xIn = Trim(xIn)
    If xIn <> "" Then
    
        I = InStr(1, xIn, "Application =")
        If I > 0 Then
            X = Mid$(xIn, I + 14, Len(xIn) - I - 13)
            blnSAA_Function = False
            For K = 1 To arrSAA_App_Nb
                If X = arrSAA_App_Code(K) Then blnSAA_Function = True: Exit For
            Next K
            If Not blnSAA_Function Then
                arrSAA_App_Nb = arrSAA_App_Nb + 1
                ReDim Preserve arrSAA_App_Code(arrSAA_App_Nb + 1)
                arrSAA_App_Code(arrSAA_App_Nb) = X
                newYSSISAA0.SSISAAUIDX = X
                newYSSISAA0.SSISAAUIDD = arrSAA_App_Nb
                mYSSISAA0_Update = "New"
                newYSSISAA0.SSISAAYFCT = "CRE"
                Call cmdSSIJRN_SAA("")
                Call cmdSelect_SQL_9_SAA_Update
            End If
         End If
    End If
Loop

Close

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

Set rsSab = Nothing
Call paramSAA_Load

Exit Sub


End Sub

Private Sub cmdSelect_SQL_9_SAA_Function()
Dim V, X As String, xIn As String, K As Integer
Dim blnSAA_Function As Boolean
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_9_SAA_Function"
Call lstErr_Clear(lstErr, cmdContext, "> " & currentAction & " ........"): DoEvents

 Call paramSAA_Load
 '____________________________________________________________________
Call cmdUpdate_Init

Call rsYSSISAA0_Init(newYSSISAA0)
newYSSISAA0.SSISAANAT = "F"
    newYSSISAA0.SSISAAYAMJ = DSys
    newYSSISAA0.SSISAAYHMS = time_Hms
    newYSSISAA0.SSISAAYUSR = usrName_UCase
    
Open mFile For Input As 1

blnSAA_Function = False
Do Until EOF(1)
    Line Input #1, xIn
    xIn = Trim(xIn)
    If xIn <> "" Then
    
        I = InStr(1, xIn, "Function =")
        If I > 0 Then
            X = Mid$(xIn, I + 11, Len(xIn) - I - 10)
            blnSAA_Function = False
            For K = 1 To arrSAA_Function_Nb
                If X = arrSAA_Function_Code(K) Then blnSAA_Function = True: Exit For
            Next K
            If Not blnSAA_Function Then
                arrSAA_Function_Nb = arrSAA_Function_Nb + 1
                ReDim Preserve arrSAA_Function_Code(arrSAA_Function_Nb + 1)
                arrSAA_Function_Code(arrSAA_Function_Nb) = X
                newYSSISAA0.SSISAAUIDX = X
                newYSSISAA0.SSISAAUIDD = arrSAA_Function_Nb
                mYSSISAA0_Update = "New"
                newYSSISAA0.SSISAAYFCT = "CRE"
                Call cmdSSIJRN_SAA("")
                Call cmdSelect_SQL_9_SAA_Update
            End If
         End If
    End If
Loop

Close

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

Set rsSab = Nothing
Call paramSAA_Load

Exit Sub


End Sub

Private Sub cmdSelect_SQL_9_SAA_Operator()
Dim V, X As String, xIn As String, K As Integer, xSQL As String
Dim blnOperator As Boolean, blnYSSIDOM0 As Boolean, blnSSISAAINFO As Boolean
Dim mSSISAAUIDD As Long
Dim SSISAAPRFX_Loop As Integer, blnSSISAAPRFX_Loop As Boolean
Dim SSISAAPRFX_Occurs As Integer

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_9_SAA_Operator"
Call lstErr_Clear(lstErr, cmdContext, "> " & currentAction & " ........"): DoEvents

blnSSISAAPRFX_Loop = True

Do While blnSSISAAPRFX_Loop

    blnSSISAAPRFX_Loop = False
    SSISAAPRFX_Loop = SSISAAPRFX_Loop + 1
    Call paramSAA_Load
     
    xSQL = "select SSISAAUIDD from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
         & " where SSISAANAT = ' '" _
         & " order by SSISAAUIDD desc"
    
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then mSSISAAUIDD = rsSab(0)
    '____________________________________________________________________
    Call cmdUpdate_Init
    
    Call rsYSSISAA0_Init(xYSSISAA0)
    
    newYSSISAA0.SSISAANAT = " "
        
    Open mFile For Input As 1
    
    blnOperator = False: blnYSSIDOM0 = False
    Do Until EOF(1)
        Line Input #1, xIn
        xIn = Trim(xIn)
        
            I = InStr(1, xIn, "Number of entries:")
            If I > 0 Then
                mImport_Nb = Val(Mid$(xIn, I + 18, Len(xIn) - I - 17))
            End If
            
            
        If xIn = "" Or xIn = "End of Report" Then
        
        Else
            If Mid$(xIn, 1, 1) = "_" Then
            Else
            
                If InStr(1, xIn, "End of Report") > 0 Then
                Else
                    I = InStr(1, xIn, "Operator ID           =")
                    If I > 0 Then
                        If blnOperator Then
                            If mYSSISAA0_Update <> "" Then
                                mImport_In = mImport_In + 1
                                newYSSISAA0.SSISAAYAMJ = DSys
                                newYSSISAA0.SSISAAYHMS = time_Hms
                                newYSSISAA0.SSISAAYUSR = usrName_UCase
                                newYSSISAA0.SSISAAINFO = xYSSISAA0.SSISAAINFO
                                newYSSISAA0.SSISAASTAK = xYSSISAA0.SSISAASTAK
                                newYSSISAA0.SSISAAUNOM = xYSSISAA0.SSISAAUNOM
                                Call cmdSelect_SQL_9_SAA_Update
                                cmdUpdate_Init
                                Call cmdSelect_SQL_9_SAA_Operator_YSSIDOM0
                            End If
                             blnOperator = False
                        End If
                        
                        Call cmdUpdate_Init
                        Call rsYSSISAA0_Init(xYSSISAA0)
                        blnOperator = True: blnYSSIDOM0 = False
                        SSISAAPRFX_Occurs = 0
           
                        xYSSISAA0.SSISAAUIDX = Mid$(xIn, I + 24, Len(xIn) - I - 23)
                        'If InStr(xIn, "SALLE") > 0 Then
                        '    Debug.Print "Debug"
                        'End If
                    End If
                    
                    If blnOperator Then
                        blnSSISAAINFO = True
                        I = InStr(1, xIn, "Approval status")
                        If I > 0 Then
                            X = Mid$(xIn, I + 24, Len(xIn) - I - 23)
                            If X = "Approved" Then
                                xYSSISAA0.SSISAASTAK = " "
                            Else
                                xYSSISAA0.SSISAASTAK = "N"
                            End If
                        End If
                        If InStr(1, xIn, "Last changed") > 0 Then blnSSISAAINFO = False
                        If InStr(1, xIn, "Calculated pwd") > 0 Then blnSSISAAINFO = False
                        If InStr(1, xIn, "Last sign-on") > 0 Then blnSSISAAINFO = False
                        If InStr(1, xIn, "Name                  =") > 0 Then
                            If Len(xIn) > I + 25 Then
                                xYSSISAA0.SSISAAUNOM = Mid$(xIn, I + 25, Len(xIn) - I - 24)
                            Else
                                xYSSISAA0.SSISAAUNOM = ""
                            End If
                            
                            If Len(xYSSISAA0.SSISAAUNOM) > 32 Then xYSSISAA0.SSISAAUNOM = Mid$(xYSSISAA0.SSISAAUNOM, 1, 32)
                        End If
                        
                        If InStr(1, xIn, "Active profile        =") > 0 Then
                            SSISAAPRFX_Occurs = SSISAAPRFX_Occurs + 1
                            If SSISAAPRFX_Loop < SSISAAPRFX_Occurs Then blnSSISAAPRFX_Loop = True

                            If SSISAAPRFX_Loop = SSISAAPRFX_Occurs Then
                                xYSSISAA0.SSISAAPRFX = Mid$(xIn, I + 25, Len(xIn) - I - 24)
                                
                                xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
                                     & " where SSISAANAT = ' ' and SSISAAUIDX = '" & xYSSISAA0.SSISAAUIDX & "'" _
                                     & " and SSISAAPRFX = '" & xYSSISAA0.SSISAAPRFX & "'"
                                
                                Set rsSab = cnsab.Execute(xSQL)
                                If Not rsSab.EOF Then
                                    mYSSISAA0_Update = "Update+H"
                                    Call rsYSSISAA0_GetBuffer(rsSab, oldYSSISAA0)
                                    newYSSISAA0 = oldYSSISAA0
                                    newYSSISAA0.SSISAAYFCT = "MOD"
                                    newYSSISAA0.SSISAAINFO = ""
                    
                                Else
                                    mYSSISAA0_Update = "New"
                                    newYSSISAA0 = xYSSISAA0
                                    newYSSISAA0.SSISAAYFCT = "CRE"
                                    newYSSISAA0.SSISAAPRFK = "?"
                                    mSSISAAUIDD = mSSISAAUIDD + 1
                                    newYSSISAA0.SSISAAUIDD = mSSISAAUIDD
                                    xSQL = "select SSISAAUSEQ from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
                                     & " where SSISAANAT = ' ' and SSISAAUIDX = '" & xYSSISAA0.SSISAAUIDX & "'" _
                                     & " order by SSISAAUSEQ desc"
                                
                                    Set rsSab = cnsab.Execute(xSQL)
                                    If Not rsSab.EOF Then newYSSISAA0.SSISAAUSEQ = rsSab("SSISAAUSEQ") + 1
                              End If
                            End If

                        End If
                        
                        If blnSSISAAINFO Then
                            xYSSISAA0.SSISAAINFO = xYSSISAA0.SSISAAINFO & xIn & vbCrLf
                        End If
                    End If
                End If
                            
            End If
        End If
    Loop
    
    If blnOperator Then
        If mYSSISAA0_Update <> "" Then
            mImport_In = mImport_In + 1
            newYSSISAA0.SSISAAYAMJ = DSys
            newYSSISAA0.SSISAAYHMS = time_Hms
            newYSSISAA0.SSISAAYUSR = usrName_UCase
            newYSSISAA0.SSISAAINFO = xYSSISAA0.SSISAAINFO
            newYSSISAA0.SSISAASTAK = xYSSISAA0.SSISAASTAK
            newYSSISAA0.SSISAAUNOM = xYSSISAA0.SSISAAUNOM
            Call cmdSelect_SQL_9_SAA_Update
            Call cmdUpdate_Init
            Call cmdSelect_SQL_9_SAA_Operator_YSSIDOM0
        End If
        Call cmdUpdate_Init
    End If
    
    Close
Loop
'________________________________________________________________________
'$JPL 2014-11-05 : profil opérateur supprimé

X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
     & " where SSIDOMNAT = ' '" _
     & " and SSIDOMDIDX = 'SAA' and SSIDOMprfd < " & mImport_PRFD _
     & " and SSIDOMPRFK <> 'X'"

Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    Call cmdUpdate_Init
    Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
    newYSSIDOM0 = oldYSSIDOM0
    newYSSIDOM0.SSIDOMPRFD = mImport_PRFD
    newYSSIDOM0.SSIDOMPRFH = mImport_PRFH
    newYSSIDOM0.SSIDOMPRFK = "X": newYSSIDOM0.SSIDOMSTAK = "N"
    
    mYSSIDOM0_Update = "Update+H"
    Call cmdUpdate
    
    rsSab.MoveNext
Loop
Call cmdUpdate_Init
'________________________________________________________________________

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYSSIUSR0_Update"
Exit_sub:

Set rsSab = Nothing
Call paramSAA_Load

Exit Sub


End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Return()
    If SSTab1.Tab = 0 Then
        'If cmdProfil_Update_DIV.Visible And cmdProfil_Update.Visible Then cmdProfil_Update_Click
        'If cmdYSSIDIV0_Update.Visible Then cmdYSSIDIV0_Update_Click
        If Not fgSelect.Visible Then cmdSelect_Ok_Click
    Else
        SendKeys "{TAB}"
    End If
End Sub


Public Sub cmdContext_Quit()
lstErr.Clear: lstErr.Height = 200



If cmdSelect_SQL_K = "3" Then
    fraDetail.Visible = False: lstW.Visible = False
    fraProfil.Visible = False
    fraProfil.Caption = ""
    txtRTF.Visible = False
    fraYSSIDIV0.Visible = False
    lstW.Visible = False
    Exit Sub
End If
If SSTab1.Tab = 1 Then
    If SSTabParam = 0 Then
        If fraParam_SSIMELUIDX.Visible Then
                fraParam_SSIMELUIDX.Visible = False
                Exit Sub
        End If
    End If
    If fraParam_K2.Visible Then
        fraParam_K2.Visible = False
        Exit Sub
    Else
        SSTab1.Tab = 0
        Exit Sub
    End If
End If
If fraYSSIDIV0.Visible Then
    fraYSSIDIV0.Visible = False
    Exit Sub
End If
If lstW.Visible Then
    lstW.Visible = False
    Exit Sub
End If

If fraCompteH.Visible Then
    fraCompteH.Visible = False
    txtRTF.Visible = False
    Exit Sub
End If

If txtRTF.Visible Then
    txtRTF.Visible = False
    Exit Sub
End If


If txtFg.Visible Then
    txtFg.Visible = False
    Exit Sub
End If
If fraYSSIDOM0.Visible Then
    cmdProfil_Quit_Click
    Exit Sub
End If


If fraProfil.Visible Then
    cmdProfil_Quit_Click
    Exit Sub
End If
If fraDetail.Visible Then
    cmdSSIUSR_Quit_Click
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





Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_SSI Traitement en cours ....."): DoEvents

'If fgSelect.Visible Then cmdSelect_Clear
'cmdSSIUSR_New.Visible = False
cmdSelect_Clear
Select Case cmdSelect_SQL_K
    Case "1": cmdSelect_SQL_1
    Case "2": cmdSelect_SQL_2
    Case "2_D": cmdSelect_SQL_2_D
    Case "2_S": cmdSelect_SQL_2_S
    Case "3": cmdSelect_SQL_3
    Case "3_H": Call cmdSelect_SQL_3_H("")
    Case "9_IBM": cmdSelect_SQL_9_SSIIBMPRFK
    Case "9_SAA": cmdSelect_SQL_9_SAA
    Case "9_SAB": cmdSelect_SQL_9_SAB
    Case "9_WIN": cmdSelect_SQL_9_WIN
    Case "9_TERENA": cmdSelect_SQL_9_DIV_TEREN 'cmdSelect_SQL_9_DIV_UGM
    Case "9_MEL": cmdSelect_SQL_9_MEL
    Case "9_TIC": cmdSelect_SQL_9_TIC
    Case "H": cmdSelect_SQL_H
    Case "J": cmdSelect_SQL_J
    Case "DS": JPL_DS
    Case "JPL":
        '   Call JPL_DS
      Call JPL
      ' Call cmdSelect_SQL_9_TIC
     '   Call paramMEL_Init
     '  Call paramSSIUSRUNIT_Init
     '________________________________
     '   paramDIV_Init_UGM_PRFX
     '   cmdSelect_SQL_9_DIV_UGM
     '   paramDIV_Init_UGM_UIDX
     '________________________________
     '   paramWIN_Init_Compte
     '   cmdSelect_SQL_9_WIN
     '   paramWIN_Init
     '   cmdSelect_SQL_9_WIN
    ' ________________________________
        'paramIBM_Init
       ' paramSAA_Init_Compte
       ' paramSAA_Init
       ' paramSAB_Init
        'cmdSelect_SQL_9_SAA_Profil_Inactif
       ' paramModèle_Init
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_SSI_cmdSelect_Ok terminé"): DoEvents
lstErr.Height = 480
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
cmdSelect_Ok.BackColor = fgSelect.BackColorFixed
End Sub









Private Sub lstParam_K1_Click()
Dim X As String
Select Case lstParam_K1.Text
    Case "SAA": Call lstParam_K1_Load("SAA")
    Case "BIA":  Call lstParam_K1_Load("BIA")
End Select

End Sub

Private Sub lstParam_K2_Click()
If IsNull(lstParam_K2_Load) Then
    If lstParam_K1.Text = "BIA" Then
        cmdParam_K2_Add.Visible = arrHab(19)
        cmdParam_K2_Add.Visible = arrHab(19)
        cmdParam_K2_Add.Visible = arrHab(19)
    Else
        cmdParam_K2_Add.Visible = arrHab(18)
        cmdParam_K2_Add.Visible = arrHab(18)
        cmdParam_K2_Add.Visible = arrHab(18)
    End If
    fraParam_K2.Visible = True
End If
End Sub


Private Sub lstParam_SSIMELNAT_Click()
Dim K As Integer, xSQL As String
fraParam_SSIMELUIDX.Visible = False
lstParam_SSIMELUIDX.Clear

mParam_SSIMELNAT_K = Mid$(lstParam_SSIMELNAT, 1, 1)
Select Case mParam_SSIMELNAT_K
    Case 1: Call lstParam_SSIMELUNOM_Load("")
    Case 2: Call lstParam_SSIMELUNOM_Load(" and SSIMELUIDX like 'BIA_GOS.%'")
    Case 3: Call lstParam_SSIMELUNOM_Load(" and SSIMELUIDX like 'SAA_Alerte.%'")
    Case 4: Call lstParam_SSIMELUNOM_Load(" and SSIMELUIDX like 'RCOM.%'")
    Case 5: Call lstParam_SSIMELUNOM_Load(" and SSIMELUIDX like 'DROPI.%'")
    Case 6: Call lstParam_SSIMELUNOM_Load(" and SSIMELUIDX like 'NoPaper.%'")
End Select

mParam_SSIMELNAT_K = lstParam_SSIMELNAT.ListIndex + 1


End Sub

Private Sub lstParam_SSIMELUIDX_Click()
Dim xSQL As String, K As Integer, K1 As Integer, I As Integer, kLen As Integer
Dim X As String, xUsr As String, blnOk As Boolean, blnExit As Boolean

cmdParam_SSIMELUIDX_Update.Visible = False
txtParam_SSIMELUNOM.Locked = Not arrHab(18)
txtParam_SSIMELUIDX.Locked = Not arrHab(18)
txtParam_SSIMELUNOM.Locked = Not arrHab(18)
txtParam_SSIMELUNOM = ""
lstParam_SSIMELUNOM.Visible = False
lstParam_SSIMELUNOM.Clear
For I = 1 To arrSSIMELUNOM_Nb
    blnSSIMELUNOM(I) = False
Next I
K = InStr(lstParam_SSIMELUIDX.Text, "|") - 1
X = Trim(Mid$(lstParam_SSIMELUIDX.Text, 1, K))
txtParam_SSIMELUIDX = X
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0 " _
             & " where SSIMELNAT = '@' and SSIMELUIDX  = '" & X & "'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    Call rsYSSIMEL0_GetBuffer(rsSab, oldYSSIMEL0)
    txtParam_SSIMELUNOM = Trim(oldYSSIMEL0.SSIMELUNOM)
'____________________________________________________________________


    xUsr = Trim(oldYSSIMEL0.SSIMELINFO)
    txtParam_SSIMELINFO = xUsr
    kLen = Len(xUsr)
    K1 = 1
    If kLen = 0 Then
        blnExit = True
    Else
        blnExit = False
    End If
    
    Do Until blnExit
        K = InStr(K1, xUsr, ";")
        If K > 0 Then
            blnOk = False
            X = StrConv(Trim(Mid$(xUsr, K1, K - K1)), vbProperCase)
            For I = 1 To arrSSIMELUNOM_Nb
                If X = arrSSIMELUNOM(I) Then blnSSIMELUNOM(I) = True: blnOk = True: Exit For
            Next I
            K1 = K + 1
        Else
            X = StrConv(Trim(Mid$(xUsr, K1, kLen - K1 + 1)), vbProperCase)
            For I = 1 To arrSSIMELUNOM_Nb
                If X = arrSSIMELUNOM(I) Then blnSSIMELUNOM(I) = True: blnOk = True: Exit For
            Next I
            blnExit = True 'Exit Do
        End If
        If blnOk = False Then
            Call MsgBox(X & " adresse mail inconnue", vbCritical, "lstParam_SSIMELUIDX")
        End If
    Loop
    
    Call lstParam_SSIMELUNOM_Display

'____________________________________________________________________
 
Else
    txtParam_SSIMELUNOM = ""
End If
cmdParam_SSIMELUIDX_New.Visible = arrHab(18)
cmdParam_SSIMELUIDX_Update.Visible = arrHab(18)
lstParam_SSIMELUNOM.Visible = True
fraParam_SSIMELUIDX.Visible = True

End Sub


Private Sub lstParam_SSIMELUNOM_Click()

If fraParam_SSIMELUIDX.Visible Then
    fraParam_SSIMELUIDX.Visible = False

    Dim K As Integer, I As Integer
    For K = 1 To arrSSIMELUNOM_Nb
        blnSSIMELUNOM(K) = False
    Next K
    For K = 0 To lstParam_SSIMELUNOM.ListCount - 1
    
        If lstParam_SSIMELUNOM.Selected(K) Then
            txtParam_SSIMELINFO = ""
            lstParam_SSIMELUNOM.ListIndex = K
            For I = 1 To arrSSIMELUNOM_Nb
                If lstParam_SSIMELUNOM = arrSSIMELUNOM(I) Then
                    blnSSIMELUNOM(I) = True: Exit For
                End If
            Next I
        End If
    Next K
    Call lstParam_SSIMELUNOM_Display
    fraParam_SSIMELUIDX.Visible = True
End If

End Sub


Private Sub lstW_Click()
Dim wUIDX As String
X = Trim(lstW.Text)
I = InStr(1, X, "|")
If I > 0 Then
    wUIDX = Val(Mid$(X, 1, I - 1))
    txtSelect_Options_J_SSIUSRUIDX = Mid$(X, I + 1, Len(X) - I)
    lstW.Visible = False
    libSelect_Options_J_SSIUSRUIDX = wUIDX
End If
End Sub

Private Sub mnuPrint_Excel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String
Call lstErr_AddItem(lstErr, cmdContext, "> BIA_SSI : export Excel ...." & fgSelect.Rows & " lignes"): DoEvents
If fraDetail.Visible Then
    Select Case cmdSelect_SQL_K
        Case "1":
            X = "Utilisateur " & oldYSSIUSR0.SSIUSRUIDX & " : situation au " & dateImp10_S(DSys) & " " & Time
        Case "2":
            X = "Modèle " & oldYSSIUSR0.SSIUSRUIDX & " : situation au " & dateImp10_S(DSys) & " " & Time
    End Select
    Call MSFlexGrid_SendMail(currentSSIWINMAIL, "BIA_SSI", X, X, fgDetail, 8)
Else
    Select Case cmdSelect_SQL_K
        Case "1":
            X = "Situation au " & dateImp10_S(DSys) & " " & Time
            Call MSflexGrid_Excel("", "BIA_SSI", X, fgSelect, 8)
             X = "Modèle " & oldYSSIUSR0.SSIUSRUIDX & " : situation au " & dateImp10_S(DSys) & " " & Time
         Case "2":
            X = "Modèle " & oldYSSIUSR0.SSIUSRUIDX & " : situation au " & dateImp10_S(DSys) & " " & Time
            Call MSflexGrid_Excel("", "BIA_SSI", X, fgSelect, 8)
         Case "2_D":
            X = "Profil " & oldYSSIUSR0.SSIUSRUIDX & " : situation au " & dateImp10_S(DSys) & " " & Time
            Call MSflexGrid_Excel("", "BIA_SSI", X, fgProfil, 3)
         Case "2_S":
            X = "Service " & oldYSSIUSR0.SSIUSRUIDX & " : situation au " & dateImp10_S(DSys) & " " & Time
            Call MSflexGrid_Excel("", "BIA_SSI", X, fgSelect, 8)
        Case "3", "3_H":
            X = "Comptes non conformes au " & dateImp10_S(DSys) & " " & Time
            Call MSflexGrid_Excel("", "BIA_SSI", X, fgSelect, 7)
        Case "J":
            X = "Journal des événements au " & dateImp10_S(DSys) & " " & Time
            Call MSflexGrid_Excel("", "BIA_SSI", X, fgSelect, 6)
         Case Else:
            X = "BIA_SSI " & oldYSSIUSR0.SSIUSRUIDX & " : situation au " & dateImp10_S(DSys) & " " & Time
            Call MSflexGrid_Excel("", "BIA_SSI", X, fgSelect, 8)
    End Select
End If

Call lstErr_AddItem(lstErr, cmdContext, "< BIA_SSI : export Excel terminé"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuPrint_Mail_Click()
Dim X As String

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_AddItem(lstErr, cmdContext, "> BIA_SSI : export mail ...."): DoEvents
If fraDetail.Visible Then
    Select Case cmdSelect_SQL_K
        Case "1":
            X = "Utilisateur " & oldYSSIUSR0.SSIUSRUIDX & " : situation au " & dateImp10_S(DSys) & " " & Time
            Call MSFlexGrid_SendMail(mMail_Destinataires, "BIA_SSI", X, X, fgDetail, 8)
         Case "2":
            X = "Modèle " & oldYSSIUSR0.SSIUSRUIDX & " : situation au " & dateImp10_S(DSys) & " " & Time
            Call MSFlexGrid_SendMail(currentSSIWINMAIL, "BIA_SSI", X, X, fgDetail, 8)
   End Select
Else
    Select Case cmdSelect_SQL_K
        Case "1":
            X = "Situation au " & dateImp10_S(DSys) & " " & Time
            Call MSFlexGrid_SendMail(mMail_Destinataires, "BIA_SSI", X, X, fgSelect, 8)
         Case "2":
            X = "Modèle " & oldYSSIUSR0.SSIUSRUIDX & " : situation au " & dateImp10_S(DSys) & " " & Time
            Call MSFlexGrid_SendMail(mMail_Destinataires, "BIA_SSI", X, X, fgSelect, 8)
         Case "2_D":
            X = "Profil " & oldYSSIUSR0.SSIUSRUIDX & " : situation au " & dateImp10_S(DSys) & " " & Time
            Call MSFlexGrid_SendMail(mMail_Destinataires, "BIA_SSI", X, X, fgProfil, 3)
         Case "2_S":
            X = "Service " & oldYSSIUSR0.SSIUSRUIDX & " : situation au " & dateImp10_S(DSys) & " " & Time
            Call MSFlexGrid_SendMail(mMail_Destinataires, "BIA_SSI", X, X, fgSelect, 8)
        Case "3", "3_H":
            X = "Comptes non conformes au " & dateImp10_S(DSys) & " " & Time
            Call MSFlexGrid_SendMail(mMail_Destinataires, "BIA_SSI", X, X, fgSelect, 7)
        Case "J":
            If Not IsNull(txtSelect_Options_J_SSITXTYMAJ) Then
                X = "Journal des événements du " & Mid$(txtSelect_Options_J_SSITXTYMAJ, 1, 10) & ", édité le " & dateImp10_S(DSys) & " " & Time
            Else
                X = "Journal des événements " & ", édité le " & dateImp10_S(DSys) & " " & Time
            End If
            Call MSFlexGrid_SendMail(mMail_Destinataires, "BIA_SSI", X, X, fgSelect, 6)
         Case Else:
            X = "BIA_SSI " & oldYSSIUSR0.SSIUSRUIDX & " : situation au " & dateImp10_S(DSys) & " " & Time
            Call MSFlexGrid_SendMail(mMail_Destinataires, "BIA_SSI", X, X, fgSelect, 8)
    End Select
End If

Call lstErr_AddItem(lstErr, cmdContext, "< BIA_SSI : export mail terminé"): DoEvents


Me.Enabled = True: Me.MousePointer = 0

End Sub



Public Sub fraDetail_Display()
Dim X As String

On Error GoTo Error_Handler
currentAction = currentAction & "-> fraDetail_Display"

fraDetail.Caption = oldYSSIUSR0.SSIUSRUIDX

fraDetail_Update.Enabled = arrHab(2)
'cboSSIUSRPRFX.Locked = True
cboSSIUSRPRFK.Locked = True

txtSSIUSRUIDX = Trim(oldYSSIUSR0.SSIUSRUIDX)
If cmdSelect_SQL_K = "2_S" Then
    txtSSIUSRUNIT_X = Trim(oldYSSIUSR0.SSIUSRPRFX)
    txtSSIUSRUNIT_N = oldYSSIUSR0.SSIUSRUIDN
Else
    Call cbo_Scan(oldYSSIUSR0.SSIUSRSTAK, cboSSIUSRSTAK)
    If oldYSSIUSR0.SSIUSRSTAK <> " " Then
        cboSSIUSRSTAK.BackColor = mColor_W1
    Else
        cboSSIUSRSTAK.BackColor = mColor_G2
    End If
    
    Call cbo_Scan(Trim(oldYSSIUSR0.SSIUSRPRFX), cboSSIUSRPRFX)
    Call cbo_Scan(oldYSSIUSR0.SSIUSRPRFK, cboSSIUSRPRFK)
    Call cbo_Scan(oldYSSIUSR0.SSIUSRUNIT, cboSSIUSRUNIT)
    If oldYSSIUSR0.SSIUSRPRFK <> " " Then
        cboSSIUSRPRFK.BackColor = mColor_W1
    Else
        cboSSIUSRPRFK.BackColor = mColor_G2
    End If
    
    If oldYSSIUSR0.SSIUSRDECH = 0 Then
        Call DTPicker_Set(txtSSIUSRDECH, DSys)
        chkSSIUSRDECH.Value = "0"
    Else
        Call DTPicker_Set(txtSSIUSRDECH, CStr(oldYSSIUSR0.SSIUSRDECH))
        chkSSIUSRDECH.Value = "1"
    End If
End If
txtSSIUSRTXT = Trim(oldYSSITXT0_USR.SSITXTINFO)

If Trim(oldYSSIUSR0.SSIUSRYUSR) <> "" Then
    lblSSIUSRYUSR = oldYSSIUSR0.SSIUSRYFCT & " par " & Trim(oldYSSIUSR0.SSIUSRYUSR) _
                        & " le " & dateImp10_S(oldYSSIUSR0.SSIUSRYAMJ) & " " & timeImp8(oldYSSIUSR0.SSIUSRYHMS) _
                        & " (" & oldYSSIUSR0.SSIUSRYVER & ")"

Else
    lblSSIUSRYUSR = ""
End If

cmdSSIUSR_Update.Visible = arrHab(2)
If oldYSSIUSR0.SSIUSRUIDN = 0 Then
    cmdSSIUSR_PRF.Visible = False
    cmdSSIUSR_Histo.Visible = False
    cboSSIUSRPRFK.Enabled = False
Else
    If oldYSSIUSR0.SSIUSRSTAK = " " Then
        cmdSSIUSR_PRF.Visible = arrHab(2) Or arrHab(5)
    Else
        cmdSSIUSR_PRF.Visible = False
    End If
    cmdSSIUSR_Histo.Visible = True
    cboSSIUSRPRFK.Enabled = True
End If
If cmdSelect_SQL_K = "2_S" Then cmdSSIUSR_PRF.Visible = False

X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
     & " where SSIDOMNAT = '" & oldYSSIUSR0.SSIUSRNAT & "' and SSIDOMUIDN = " & oldYSSIUSR0.SSIUSRUIDN _
     & " order by SSIDOMSTAK, SSIDOMDIDX  , SSIDOMUIDX , SSIDOMPRFX"

Set rsSab = cnsab.Execute(X)
Call fgDetail_Display

fraDetail.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub fraYSSIDOM0_Display()
Dim X As String

On Error GoTo Error_Handler
currentAction = currentAction & "-> fraYSSIDOM0_Display"

cboSSIDOMPRFK.Locked = True
txtSSIDOMPRFX.Locked = True
cmdCompte_Val.Visible = False

libSSIDOMUIDD = oldYSSIDOM0.SSIDOMUIDD & " / " & oldYSSIDOM0.SSIDOMUIDX
libSSIDOMPRFX = ""

Call cbo_Scan(oldYSSIDOM0.SSIDOMSTAK, cboSSIDOMSTAK)
If oldYSSIDOM0.SSIDOMSTAK <> " " Then
    cboSSIDOMSTAK.BackColor = RGB(192, 192, 192)
Else
    cboSSIDOMSTAK.BackColor = mColor_G2
End If


txtSSIDOMPRFX = Trim(oldYSSIDOM0.SSIDOMPRFX)

Call cbo_Scan(oldYSSIDOM0.SSIDOMPRFK, cboSSIDOMPRFK)
Call cbo_Scan(oldYSSIDOM0.SSIDOMUNIT, cboSSIDOMUNIT)

Select Case oldYSSIDOM0.SSIDOMPRFK
    Case " ": cboSSIDOMPRFK.BackColor = mColor_G2
    Case "N": cboSSIDOMPRFK.BackColor = mColor_W1
    Case "X": cboSSIDOMPRFK.BackColor = RGB(192, 192, 192)
    Case "!": cboSSIDOMPRFK.BackColor = mColor_Y2
    Case Else: cboSSIDOMPRFK.BackColor = mColor_W1
End Select

If oldYSSIDOM0.SSIDOMDECH = 0 Then
    Call DTPicker_Set(txtSSIDOMDECH, DSys)
    chkSSIDOMDECH.Value = "0"
Else
    Call DTPicker_Set(txtSSIDOMDECH, CStr(oldYSSIDOM0.SSIDOMDECH))
    chkSSIDOMDECH.Value = "1"
End If

txtSSIDOMTXT = Trim(oldYSSITXT0_DOM.SSITXTINFO)
If oldYSSIDOM0.SSIDOMPRFD <> 0 Then
    libSSIDOMPRFD = dateImp10_S(oldYSSIDOM0.SSIDOMPRFD) & "  " & timeImp8(oldYSSIDOM0.SSIDOMPRFH)
Else
    libSSIDOMPRFD = ""
End If

If Trim(oldYSSIDOM0.SSIDOMYUSR) <> "" Then
    lblSSIDOMYUSR = oldYSSIDOM0.SSIDOMYFCT & " par " & Trim(oldYSSIDOM0.SSIDOMYUSR) _
                  & " le " & dateImp10_S(oldYSSIDOM0.SSIDOMYAMJ) & " " & timeImp8(oldYSSIDOM0.SSIDOMYHMS) _
                  & " (" & oldYSSIDOM0.SSIDOMYVER & ")"
Else
    lblSSIDOMYUSR = ""
End If

fgCompte.Visible = False: cmdCompte_All.Visible = False
fgCompte.Clear
If cmdSelect_SQL_K = "1" Then
    If oldYSSIDOM0.SSIDOMNAT = " " Then
        Select Case oldYSSIDOM0.SSIDOMDIDX
            Case "IBM": fraYSSIDOM0_Display_IBM
            Case "SAA": fraYSSIDOM0_Display_SAA
            Case "SAB": fraYSSIDOM0_Display_SAB
            Case "WIN": fraYSSIDOM0_Display_WIN
            Case "DIV": fraYSSIDOM0_Display_DIV
            Case "MEL": fraYSSIDOM0_Display_MEL
            Case "TIC": fraYSSIDOM0_Display_TIC
        End Select
    
        If oldYSSIDOM0.SSIDOMUIDD > 0 And oldYSSIDOM0.SSIDOMPRFK = "N" Then cmdCompte_Val.Visible = blnHab2_SécuritéPhysique   'arrHab(2)
       ' cmdProfil_Update.Visible = False
    End If
End If
fraYSSIDOM0.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub





Private Sub mnuPrint_RTF_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

X = "C:\Temp\BIA_SSI " & DSYS_Time
'______________________________________________
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & X _
        & vbCrLf & "     =========================", "Nom du fichier d'exportation ", X)
 If Trim(X) <> "" Then Call cmdPrint_Word_PDF(X)
    


Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint_RTF_USR_Click()
Dim K As Integer, mRTF_All As String, X As String

Me.Enabled = False: Me.MousePointer = vbHourglass


xYSSIUSR0 = oldYSSIUSR0
mRTF_All = cmdSSIUSR_Detail_txtRTF(11)

For K = 1 To fgDetail.Rows - 1
    fgDetail.Row = K
        
     xYSSIDOM0.SSIDOMNAT = mSSIUSRNAT
     xYSSIDOM0.SSIDOMUIDN = oldYSSIUSR0.SSIUSRUIDN
     fgDetail.Col = 0: xYSSIDOM0.SSIDOMDIDX = Trim(fgDetail.Text)
     fgDetail.Col = 2: xYSSIDOM0.SSIDOMUIDX = Trim(fgDetail.Text)
     fgDetail.Col = 3: xYSSIDOM0.SSIDOMUIDD = Val(fgDetail.Text)
     X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
          & " where SSIDOMNAT = '" & xYSSIDOM0.SSIDOMNAT & "' and SSIDOMUIDN = " & xYSSIDOM0.SSIDOMUIDN _
          & " and SSIDOMDIDX = '" & xYSSIDOM0.SSIDOMDIDX & "' and SSIDOMUIDX = '" & xYSSIDOM0.SSIDOMUIDX & "'" _
          & " and SSIDOMUIDD = " & xYSSIDOM0.SSIDOMUIDD
     Set rsSab = cnsab.Execute(X)

    If Not rsSab.EOF Then
         Call rsYSSIDOM0_GetBuffer(rsSab, xYSSIDOM0)
         mRTF_All = mRTF_All & cmdSSIDOM_Detail_txtRTF(11)
         Select Case xYSSIDOM0.SSIDOMDIDX
            Case "SAB"
                usrYSSISAB0.SSISABNAT = xYSSIDOM0.SSIDOMNAT
                usrYSSISAB0.SSISABUIDX = xYSSIDOM0.SSIDOMUIDX
                Call cmdSSISAB_Detail_Load("SAB")
                mRTF_All = mRTF_All & mRTF
                    
             Case "SAA"
                usrYSSISAA0.SSISAANAT = xYSSIDOM0.SSIDOMNAT
                usrYSSISAA0.SSISAAUIDX = xYSSIDOM0.SSIDOMUIDX
                Call cmdSSISAA_Detail_Load("SAA")
                mRTF_All = mRTF_All & mRTF
             Case "IBM"
                oldYSSIIBM0.UPUPRF = xYSSIDOM0.SSIDOMPRFX
                oldYSSIIBM0.SSIIBMNAT = "$"
                usrYSSIIBM0.SSIIBMNAT = xYSSIDOM0.SSIDOMNAT
                usrYSSIIBM0.SSIIBMUIDD = xYSSIDOM0.SSIDOMUIDD
                Call cmdSSIIBM_Detail_Load("UPUPRF", "USR")
                mRTF_All = mRTF_All & mRTF
             Case "WIN"
                rtfYSSIWIN0.SSIWINNAT = xYSSIDOM0.SSIDOMNAT
                rtfYSSIWIN0.SSIWINUIDD = xYSSIDOM0.SSIDOMUIDD
                Call cmdSSIWIN_Detail_Display("YSSIWIN0_UIDD")
                mRTF_All = mRTF_All & mRTF
             Case "DIV"
                rtfYSSIDIV0.SSIDIVNAT = xYSSIDOM0.SSIDOMNAT
                rtfYSSIDIV0.SSIDIVUIDD = xYSSIDOM0.SSIDOMUIDD
                Call cmdSSIDIV_Detail_Display("YSSIDIV0_UIDD")
                mRTF_All = mRTF_All & mRTF
             Case "MEL"
                rtfYSSIMEL0.SSIMELNAT = xYSSIDOM0.SSIDOMNAT
                rtfYSSIMEL0.SSIMELUIDX = xYSSIDOM0.SSIDOMUIDX
                Call cmdSSIMEL_Detail_Display("YSSIDIV0_UIDX")
                mRTF_All = mRTF_All & mRTF
             Case "TIC"
                rtfYSSITIC0.SSITICNAT = xYSSIDOM0.SSIDOMNAT
                rtfYSSITIC0.SSITICUIDX = xYSSIDOM0.SSIDOMUIDX
                Call cmdSSITIC_Detail_Display("YSSITIC0")
                mRTF_All = mRTF_All & mRTF
       End Select

     End If
     
Next K
txtRTF.TextRTF = VB_RTF_Modèle

txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", mRTF_All)

Call txtRTF_Visible
X = "C:\Temp\BIA_SSI " & oldYSSIUSR0.SSIUSRUIDX & " " & DSYS_Time
'______________________________________________
X = InputBox("par défaut : " _
    & vbCrLf & "     =========================" & vbCrLf & X _
    & vbCrLf & "     =========================", "Nom du fichier d'exportation ", X)
 If Trim(X) <> "" Then Call cmdPrint_Word_PDF(X)

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub txtParam_K2_Code_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtProfil_DIDK_DIV_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtProfil_UIDD_DIV_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtProfil_UIDD_KeyPress(KeyAscii As Integer)
Select Case mSSIDOMDIDX
    Case "IBM": Call num_KeyAscii(KeyAscii)
End Select

End Sub


Private Sub txtSelect_Options_1_SSIUSRUIDX_Change()
cmdSelect_Clear
End Sub

Private Sub txtSelect_Options_1_SSIUSRUIDX_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelect_Options_4_SSIDOMUIDD_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_Options_4_SSIDOMUIDD_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtSelect_Options_4_SSIDOMUIDX_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_Options_4_SSIDOMUIDX_KeyPress(KeyAscii As Integer)
'KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelect_Options_J_SSITXTYMAJ_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_Options_J_SSITXTYMAJ_Click()
cmdSelect_Clear

End Sub

Private Sub txtSelect_Options_J_SSIUSRUIDX_Change()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdSelect_Clear
libSelect_Options_J_SSIUSRUIDX = ""
lstW.Visible = False
lstW.Clear

X = "select SSIUSRUIDX ,SSIUSRUIDN from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
         & " where SSIUSRUIDX like '" & Trim(txtSelect_Options_J_SSIUSRUIDX) & "%'" _
         & " order by SSIUSRUIDX , SSIUSRUIDN"

Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    lstW.AddItem rsSab("SSIUSRUIDN") & "|" & Trim(rsSab("SSIUSRUIDx"))
    rsSab.MoveNext
Loop


lstW.Visible = True

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub txtSelect_Options_J_SSIUSRUIDX_Click()
cmdSelect_Clear

End Sub

Private Sub txtSelect_Options_J_SSIUSRUIDX_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtSSIUSRDECH_Change()
cmdSSIUSR_PRF.Visible = False
End Sub


Private Sub txtSSIUSRUNIT_N_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtSSIUSRTXT_Change()
cmdSSIUSR_PRF.Visible = False

End Sub


Private Sub txtSSIUSRUIDX_Change()
cmdSSIUSR_PRF.Visible = False
If cmdSelect_SQL_K <> "2_S" Then txtSSIUSRUIDX = UCase(txtSSIUSRUIDX)

End Sub

Private Sub txtSSIUSRUIDX_KeyPress(KeyAscii As Integer)
If cmdSelect_SQL_K <> "2_S" Then KeyAscii = convUCase(KeyAscii)
End Sub


Public Function fraDetail_Control()
Dim X As String, wMsgBox As String
Dim wAMJ As Long, wSSITXTINFO As String
On Error GoTo Exit_sub

currentAction = "fraDetail_Control"
Call cmdUpdate_Init

newYSSIUSR0 = oldYSSIUSR0
newYSSITXT0 = oldYSSITXT0_USR
oldYSSITXT0_XXX = oldYSSITXT0_USR
'==========================
wMsgBox = ""
fraDetail_Control = "?"

If cmdSelect_SQL_K = "2_S" Then
    Call fraDetail_Control_SSIUSRUNIT(wMsgBox)
Else

    If newYSSIUSR0.SSIUSRUIDN = 0 Then
        mYSSIUSR0_Update = "New"
        newYSSIUSR0.SSIUSRYFCT = "CRE"
       X = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
             & " where SSIUSRNAT = '" & oldYSSIUSR0.SSIUSRNAT & "' and SSIUSRUIDX = '" & Trim(txtSSIUSRUIDX) & "'"
        
        Set rsSab = cnsab.Execute(X)
        If Not rsSab.EOF Then
            Select Case oldYSSIUSR0.SSIUSRNAT
                Case "$": wMsgBox = wMsgBox & " - Ce modèle BIA existe déjà" & vbCrLf
                Case Else: wMsgBox = wMsgBox & " - Cet utilisateur existe déjà" & vbCrLf
            End Select
        End If
    Else
        mYSSIUSR0_Update = "Update+H"
        newYSSIUSR0.SSIUSRYFCT = "MOD"
    End If
    
    newYSSIUSR0.SSIUSRUIDX = Trim(txtSSIUSRUIDX)
    If newYSSIUSR0.SSIUSRUIDX = "" Then wMsgBox = wMsgBox & " - préciser le nom" & vbCrLf
    
    newYSSIUSR0.SSIUSRSTAK = Mid$(cboSSIUSRSTAK.Text, 1, 1)
    newYSSIUSR0.SSIUSRPRFX = Trim(cboSSIUSRPRFX)
    newYSSIUSR0.SSIUSRPRFK = Mid$(cboSSIUSRPRFK, 1, 1)
    If newYSSIUSR0.SSIUSRPRFX <> "" And newYSSIUSR0.SSIUSRPRFX <> oldYSSIUSR0.SSIUSRPRFX Then wMsgBox = wMsgBox & fraDetail_Control_SSIUSRPRFX
    If cboSSIUSRUNIT = "" Then
        newYSSIUSR0.SSIUSRUNIT = ""
    Else
        newYSSIUSR0.SSIUSRUNIT = Mid$(cboSSIUSRUNIT, 1, 3)
    End If
    If newYSSIUSR0.SSIUSRUNIT <> oldYSSIUSR0.SSIUSRUNIT Then
        mYSSIDOM0_Update_CMD_2 = "Update " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
             & " set SSIDOMUNIT = '" & newYSSIUSR0.SSIUSRUNIT & "'" _
             & "  Where SSIDOMNAT = '" & newYSSIUSR0.SSIUSRNAT & "' and SSIDOMUIDN = " & newYSSIUSR0.SSIUSRUIDN
    End If
    If chkSSIUSRDECH.Value = "1" Then
        Call DTPicker_Control(txtSSIUSRDECH, X)
        wAMJ = CLng(X)
        If wAMJ <> oldYSSIUSR0.SSIUSRDECH Then
            If wAMJ < DSys Then
                wMsgBox = wMsgBox & " - date échéance  < aujourd'hui" & vbCrLf
            Else
                newYSSIUSR0.SSIUSRDECH = wAMJ
                If newYSSIUSR0.SSIUSRPRFK = "!" Then newYSSIUSR0.SSIUSRPRFK = " "
            End If
        End If
    Else
        newYSSIUSR0.SSIUSRDECH = 0
    End If
End If

'____________________________________________________________________________________
newYSSIUSR0.SSIUSRYAMJ = DSys
newYSSIUSR0.SSIUSRYHMS = time_Hms
newYSSIUSR0.SSIUSRYUSR = usrName_UCase

'____________________________________________________________________________________
wSSITXTINFO = Trim(txtSSIUSRTXT)
newYSSITXT0.SSITXTINFO = wSSITXTINFO
newYSSITXT0.SSITXTYAMJ = DSys
newYSSITXT0.SSITXTYHMS = time_Hms
newYSSITXT0.SSITXTYUSR = usrName_UCase
'____________________________________________________________________________________

If wMsgBox <> "" Then
    fraDetail_Control = "?"
    Call MsgBox(wMsgBox, vbCritical, currentAction)
Else
    fraDetail_Control = Null
    
    If Trim(oldYSSITXT0_USR.SSITXTINFO) <> Trim(newYSSITXT0.SSITXTINFO) Then mYSSITXT0_Update = "New"
           
    If newYSSIUSR0.SSIUSRSTAK = "N" And oldYSSIUSR0.SSIUSRSTAK <> "N" Then
         If newYSSIUSR0.SSIUSRUIDX <> oldYSSIUSR0.SSIUSRUIDX Then
            fraDetail_Control = "?"
            Call MsgBox("Ne pas modifier le nom et le code état dans la même transaction", vbCritical, currentAction)
         Else
            mYSSIDOM0_Update = "CMD"
            mYSSIDOM0_Update_CMD = "Update " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
             & " set SSIDOMSTAK = 'N' " _
             & "  Where SSIDOMNAT = '" & newYSSIUSR0.SSIUSRNAT & "' and SSIDOMUIDN = " & newYSSIUSR0.SSIUSRUIDN
        End If
    End If
    If newYSSIUSR0.SSIUSRNAT = "$" Then
        If newYSSIUSR0.SSIUSRUIDX <> oldYSSIUSR0.SSIUSRUIDX Then
            If Trim(oldYSSIUSR0.SSIUSRUIDX) <> "" Then
                mYSSIDOM0_Update = "CMD"
                mYSSIDOM0_Update_CMD = "Update " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
                 & " set SSIDOMPRFX = '" & newYSSIUSR0.SSIUSRUIDX & "' " _
                 & "  Where SSIDOMNAT = ' ' and SSIDOMPRFX = '" & oldYSSIUSR0.SSIUSRUIDX & "'"
            End If
        End If
    End If
    
End If

Exit_sub:

End Function
Public Function fraDetail_Control_SSIUSRPRFX()
Dim X As String, K1 As Integer, K2 As Integer, blnExist As Boolean
Dim xWhere As String, wSSIUSRSTAK As String
On Error GoTo Exit_sub

currentAction = "fraDetail_Control_BIA"

fraDetail_Control_SSIUSRPRFX = ""

X = "select SSIUSRUIDN from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
     & " where SSIUSRNAT = '$' and SSIUSRUIDX = '" & newYSSIUSR0.SSIUSRPRFX & "'"

Set rsSab = cnsab.Execute(X)
If rsSab.EOF Then
    fraDetail_Control_SSIUSRPRFX = " - modèle inconnu : " & newYSSIUSR0.SSIUSRPRFX & vbCrLf
    Exit Function
End If

xWhere = " where SSIDOMNAT = '$' and SSIDOMUIDn = " & rsSab("SSIUSRUIDN") & " and SSIDOMSTAK = ' '"
X = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " & xWhere
     

Set rsSab = cnsab.Execute(X)
If rsSab(0) = 0 Then
    fraDetail_Control_SSIUSRPRFX = " - modèle sans profils associés : " & newYSSIUSR0.SSIUSRPRFX & vbCrLf
    Exit Function
End If

ReDim arrYSSIDOM0_BIA(rsSab(0) + 1)
arrYSSIDOM0_BIA_Nb = 0
X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " & xWhere

Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    arrYSSIDOM0_BIA_Nb = arrYSSIDOM0_BIA_Nb + 1
    Call rsYSSIDOM0_GetBuffer(rsSab, arrYSSIDOM0_BIA(arrYSSIDOM0_BIA_Nb))
    rsSab.MoveNext
Loop
'__________________________________________________________________________________________________________

 xWhere = " where SSIDOMNAT = ' ' and SSIDOMUIDN = " & newYSSIUSR0.SSIUSRUIDN

X = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " & xWhere

     
Set rsSab = cnsab.Execute(X)

ReDim arrYSSIDOM0_Usr(rsSab(0) + 1)
arrYSSIDOM0_Usr_Nb = 0
X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " & xWhere
     
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    arrYSSIDOM0_Usr_Nb = arrYSSIDOM0_Usr_Nb + 1
    Call rsYSSIDOM0_GetBuffer(rsSab, arrYSSIDOM0_Usr(arrYSSIDOM0_Usr_Nb))
    rsSab.MoveNext
Loop

'__________________________________________________________________________________________________________


For K1 = 1 To arrYSSIDOM0_BIA_Nb
    blnExist = False
    For K2 = 1 To arrYSSIDOM0_Usr_Nb
        If arrYSSIDOM0_BIA(K1).SSIDOMDIDX = arrYSSIDOM0_Usr(K2).SSIDOMDIDX _
        And arrYSSIDOM0_BIA(K1).SSIDOMPRFX = arrYSSIDOM0_Usr(K2).SSIDOMPRFX Then blnExist = True: Exit For
    Next K2
    If Not blnExist Then
        mYSSIDOM0_Update = "PRFX$"
        arrYSSIDOM0_BIA(K1).SSIDOMNAT = " "
        arrYSSIDOM0_BIA(K1).SSIDOMUIDN = newYSSIUSR0.SSIUSRUIDN
        'arrYSSIDOM0_BIA(K1).SSIDOMUIDD = 0
        arrYSSIDOM0_BIA(K1).SSIDOMUIDX = ""
        arrYSSIDOM0_BIA(K1).SSIDOMSTAK = newYSSIUSR0.SSIUSRSTAK
        arrYSSIDOM0_BIA(K1).SSIDOMDECH = newYSSIUSR0.SSIUSRDECH
        arrYSSIDOM0_BIA(K1).SSIDOMPRFK = "?"
        arrYSSIDOM0_BIA(K1).SSIDOMPRFD = 0
        arrYSSIDOM0_BIA(K1).SSIDOMPRFH = 0
        arrYSSIDOM0_BIA(K1).SSIDOMTLNK = 0
        arrYSSIDOM0_BIA(K1).SSIDOMYVER = 0
        arrYSSIDOM0_BIA(K1).SSIDOMYFCT = "INI"
        arrYSSIDOM0_BIA(K1).SSIDOMYAMJ = DSys
        arrYSSIDOM0_BIA(K1).SSIDOMYHMS = time_Hms
        arrYSSIDOM0_BIA(K1).SSIDOMYUSR = usrName_UCase
   
   End If
Next K1

newYSSIUSR0.SSIUSRPRFK = " "
For K2 = 1 To arrYSSIDOM0_Usr_Nb
    blnExist = False
    For K1 = 1 To arrYSSIDOM0_BIA_Nb
        If arrYSSIDOM0_BIA(K1).SSIDOMDIDX = arrYSSIDOM0_Usr(K2).SSIDOMDIDX _
        And arrYSSIDOM0_BIA(K1).SSIDOMPRFX = arrYSSIDOM0_Usr(K2).SSIDOMPRFX Then blnExist = True: Exit For
    Next K1
    If Not blnExist Then
        newYSSIUSR0.SSIUSRPRFK = "N"
        newYSSIUSR0.SSIUSRYAMJ = DSys
        newYSSIUSR0.SSIUSRYHMS = time_Hms
    End If
Next K2

Exit_sub:

End Function

Public Function fraProfil_Control_IBM()
Dim X As String, wMsgBox As String
On Error GoTo Exit_sub

currentAction = "fraProfil_Control_IBM"
Call cmdUpdate_Init

newYSSIIBM0 = oldYSSIIBM0
'==========================
wMsgBox = ""
fraProfil_Control_IBM = "?"


If Trim(txtProfil_IDX) = "" Then
    wMsgBox = wMsgBox & " - préciser le nom du profil" & vbCrLf
Else
    newYSSIIBM0.UPUPRF = Trim(txtProfil_IDX)

    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0 " _
         & " where SSIIBMNAT = '$'  and UPUPRF = '" & newYSSIIBM0.UPUPRF & "'"
    
    Set rsSab = cnsab.Execute(X)
      
    If Not rsSab.EOF Then
        wMsgBox = wMsgBox & " - ce nom du profil est déjà utilisé" & vbCrLf
    End If
End If
If Trim(txtProfil_UPTEXT) = "" Then
    wMsgBox = wMsgBox & " - préciser un commentaire" & vbCrLf
Else
    newYSSIIBM0.UPTEXT = Trim(txtProfil_UPTEXT)
End If
'____________________________________________________________________________________

If wMsgBox <> "" Then
    fraProfil_Control_IBM = "?"
    Call MsgBox(wMsgBox, vbCritical, currentAction)
Else
    newYSSIIBM0.SSIIBMYAMJ = DSys
    newYSSIIBM0.SSIIBMYHMS = time_Hms
    newYSSIIBM0.SSIIBMYUSR = usrName_UCase
    newYSSIIBM0.SSIIBMYVER = 0
    newYSSIIBM0.SSIIBMPRFK = ""
   fraProfil_Control_IBM = Null
    mYSSIIBM0_Update = "New"
End If

Exit_sub:

End Function

Public Function fraProfil_Control_DIV()
Dim X As String, wMsgBox As String
On Error GoTo Exit_sub

currentAction = "fraProfil_Control_DIV"
Call cmdUpdate_Init
wMsgBox = ""
fraProfil_Control_DIV = "?"

Select Case Trim(txtProfil_DIDK_DIV)
    Case "TEREN", "SG", "UGM":
        If Not arrHab(5) Then wMsgBox = wMsgBox & " - vous n'êtes pas habilité à ce sous-domaine" & vbCrLf
    Case Else:
        If Not arrHab(2) Then wMsgBox = wMsgBox & " - vous n'êtes pas habilité à ce sous-domaine" & vbCrLf
End Select

'==========================
If fraProfil_Update_DIV.Caption <> "Création d'un profil" Then
    newYSSIDIV0 = oldYSSIDIV0
    mYSSIDIV0_Update = "Update"
Else
    mYSSIDIV0_Update = "New"
    Call rsYSSIDIV0_Init(newYSSIDIV0)
    newYSSIDIV0.SSIDIVNAT = "$"
    newYSSIDIV0.SSIDIVYFCT = "CRE"
    
    
    If Trim(txtProfil_DIDK_DIV) = "" Then
        wMsgBox = wMsgBox & " - préciser le sous-domaine" & vbCrLf
    Else
        newYSSIDIV0.SSIDIVDIDK = Trim(txtProfil_DIDK_DIV)
    End If
        
    If Val(txtProfil_UIDD_DIV) = 0 Then
        wMsgBox = wMsgBox & " - préciser le numéro du profil" & vbCrLf
    Else
        newYSSIDIV0.SSIDIVUIDD = Val(txtProfil_UIDD_DIV)
        If Trim(txtProfil_DIDK_DIV) = "TEREN" Then
            If newYSSIDIV0.SSIDIVUIDD <> 100 Then
                wMsgBox = wMsgBox & " - TEREN : le numéro de profil = 100 (JPL 2014-11-06)" & vbCrLf
            End If
        Else
            newYSSIDIV0.SSIDIVUIDD = Val(txtProfil_UIDD_DIV)
        
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
                 & " where SSIDIVNAT = '$'  and SSIDIVUIDD = " & newYSSIDIV0.SSIDIVUIDD
            
            Set rsSab = cnsab.Execute(X)
              
            If Not rsSab.EOF Then
                wMsgBox = wMsgBox & " - ce numéro de profil est déjà utilisé" & vbCrLf
            End If
        End If
    End If
    
    
    If Trim(txtProfil_IDX_DIV) = "" Then
        wMsgBox = wMsgBox & " - préciser le nom du profil" & vbCrLf
    Else
        newYSSIDIV0.SSIDIVUIDX = Trim(txtProfil_IDX_DIV)
        newYSSIDIV0.SSIDIVPRFX = newYSSIDIV0.SSIDIVUIDX
    
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
             & " where SSIDIVNAT = '$'  and SSIDIVUIDX = '" & newYSSIDIV0.SSIDIVUIDX & "'"
        
        Set rsSab = cnsab.Execute(X)
          
        If Not rsSab.EOF Then
            wMsgBox = wMsgBox & " - ce CODE du profil est déjà utilisé" & vbCrLf
        End If
    End If
End If
'_________________________________________________________________________________


If Trim(txtProfil_UNOM_DIV) = "" Then
    wMsgBox = wMsgBox & " - préciser le nom" & vbCrLf
Else
    newYSSIDIV0.SSIDIVUNOM = Trim(txtProfil_UNOM_DIV)
End If
'If Trim(txtProfil_Info_DIV) = "" Then
'    wMsgBox = wMsgBox & " - préciser un commentaire" & vbCrLf
'Else
    newYSSIDIV0.SSIDIVINFO = Trim(txtProfil_Info_DIV)
    newYSSIDIV0.SSIDIVPRFX = Trim(txtProfil_PRFX_DIV)
'End If
'____________________________________________________________________________________

If wMsgBox <> "" Then
    fraProfil_Control_DIV = "?"
    Call MsgBox(wMsgBox, vbCritical, currentAction)
Else
    newYSSIDIV0.SSIDIVYAMJ = DSys
    newYSSIDIV0.SSIDIVYHMS = time_Hms
    newYSSIDIV0.SSIDIVYUSR = usrName_UCase
    newYSSIDIV0.SSIDIVPRFK = ""
    fraProfil_Control_DIV = Null
End If

Exit_sub:

End Function

Public Function fraYSSIDIV0_Control()
Dim X As String, wMsgBox As String
On Error GoTo Exit_sub

currentAction = "fraYSSIDIV0_Control"
Call cmdUpdate_Init

'==========================
wMsgBox = ""
fraYSSIDIV0_Control = "?"

newYSSIDOM0 = oldYSSIDOM0

If Trim(oldYSSIDIV0.SSIDIVUIDX) <> "" Then
    newYSSIDIV0 = oldYSSIDIV0
Else
    mYSSIDIV0_Update = "New"
    Call rsYSSIDIV0_Init(newYSSIDIV0)
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
         & " where SSIDIVNAT = '$' and SSIDIVUIDX = '" & oldYSSIDOM0.SSIDOMPRFX & "'"
    Set rsSab = cnsab.Execute(X)
    If rsSab.EOF Then
        Call MsgBox(X, vbCritical, "fraYSSIDIV0_Control: profil inconnu : " & oldYSSIDOM0.SSIDOMPRFX)
    Else
        newYSSIDIV0.SSIDIVDIDK = rsSab("SSIDIVDIDK")
        newYSSIDIV0.SSIDIVUIDD = rsSab("SSIDIVUIDD")
        newYSSIDIV0.SSIDIVPRFX = oldYSSIDOM0.SSIDOMPRFX
        newYSSIDIV0.SSIDIVYFCT = "cpt"
    End If
        
End If


If Trim(txtSSIDIVUIDX) = "" Then
    wMsgBox = wMsgBox & " - préciser l'identifiant du compte" & vbCrLf
Else
    newYSSIDIV0.SSIDIVUIDX = Trim(txtSSIDIVUIDX)
End If

If Trim(txtSSIDIVUNOM) = "" Then
    wMsgBox = wMsgBox & " - préciser le nom" & vbCrLf
Else
    newYSSIDIV0.SSIDIVUNOM = Trim(txtSSIDIVUNOM)
End If
newYSSIDIV0.SSIDIVINFO = Trim(txtSSIDIVINFO)

If chkSSIDIVPRFK.Visible Then
    If chkSSIDIVPRFK.Value = "1" Then
        newYSSIDIV0.SSIDIVPRFK = "X"
    Else
        newYSSIDIV0.SSIDIVPRFK = " "
    End If
    newYSSIDOM0.SSIDOMPRFK = newYSSIDIV0.SSIDIVPRFK
End If
   
newYSSIDIV0.SSIDIVYAMJ = DSys
newYSSIDIV0.SSIDIVYHMS = time_Hms
newYSSIDIV0.SSIDIVYUSR = usrName_UCase

mYSSIDOM0_Update = "Update"
newYSSIDOM0.SSIDOMYFCT = "vu"
newYSSIDOM0.SSIDOMPRFD = DSys
newYSSIDOM0.SSIDOMPRFH = time_Hms
newYSSIDOM0.SSIDOMYAMJ = DSys
newYSSIDOM0.SSIDOMYHMS = time_Hms
newYSSIDOM0.SSIDOMYUSR = usrName_UCase
If mYSSIDIV0_Update = "New" Then
    'mYSSIDOM0_Update = "CMD"
    'mYSSIDOM0_Update_CMD = "Update " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
    ' & " set SSIDOMUIDX  = '" & newYSSIDIV0.SSIDIVUIDX & "', SSIDOMUIDD =" & newYSSIDIV0.SSIDIVUIDD & ", SSIDOMPRFK ='" & newYSSIDIV0.SSIDIVPRFK & "'" _
    ' & " , SSIDOMYFCT = 'cpt' , SSIDOMYUSR = '" & usrName_UCase & "', SSIDOMYAMJ = " & DSys & ", SSIDOMYHMS = " & time_Hms _
    ' & "  Where SSIDOMNAT = '" & oldYSSIDOM0.SSIDOMNAT & "' and SSIDOMUIDN = " & oldYSSIDOM0.SSIDOMUIDN _
    ' & "  and SSIDOMDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "' and SSIDOMUIDX = '" & oldYSSIDOM0.SSIDOMUIDX & "'" _
    ' & " and SSIDOMUIDD= " & oldYSSIDOM0.SSIDOMUIDD
    
    mYSSIDOM0_Update = "New"
    newYSSIDOM0.SSIDOMUIDD = newYSSIDIV0.SSIDIVUIDD
    newYSSIDOM0.SSIDOMUIDX = newYSSIDIV0.SSIDIVUIDX
    newYSSIDOM0.SSIDOMPRFX = newYSSIDIV0.SSIDIVPRFX
    newYSSIDOM0.SSIDOMPRFK = " "
    newYSSIDOM0.SSIDOMYFCT = "cpt"
    
    X = "select SSIDIVYVER from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
         & " where SSIDIVNAT = ' ' and SSIDIVUIDX = '" & newYSSIDIV0.SSIDIVUIDX & "'" _
         & " order by SSIDIVYVER desc FETCH FIRST 1 ROWS ONLY"
    Set rsSab = cnsab.Execute(X)
    If Not rsSab.EOF Then newYSSIDIV0.SSIDIVYVER = rsSab("SSIDIVYVER")
Else
    
    If oldYSSIDOM0.SSIDOMPRFK <> newYSSIDOM0.SSIDOMPRFK Then
        mYSSIDOM0_Update = "Update+H"
        newYSSIDOM0.SSIDOMYFCT = "mod"
    End If
    If oldYSSIDIV0.SSIDIVUNOM <> newYSSIDIV0.SSIDIVUNOM _
    Or oldYSSIDIV0.SSIDIVPRFK <> newYSSIDIV0.SSIDIVPRFK _
    Or oldYSSIDIV0.SSIDIVINFO <> newYSSIDIV0.SSIDIVINFO Then
        mYSSIDIV0_Update = "Update+H"
        newYSSIDIV0.SSIDIVYFCT = "mod"
    End If

End If

'____________________________________________________________________________________

If wMsgBox <> "" Then
    fraYSSIDIV0_Control = "?"
    Call MsgBox(wMsgBox, vbCritical, currentAction)
Else
    fraYSSIDIV0_Control = Null
End If

Exit_sub:

End Function



Public Sub fraDetail_Load()
Dim xSQL As String

fraDetail.Visible = False
txtRTF.Visible = False
fraProfil.Visible = False
lstW.Visible = False
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
     & " where SSIUSRNAT = '" & oldYSSIUSR0.SSIUSRNAT & "' and SSIUSRUIDN = " & oldYSSIUSR0.SSIUSRUIDN

Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    Call MsgBox(xSQL, vbCritical, "fraDetail_Load : inconnu")
    Exit Sub
End If
Call rsYSSIUSR0_GetBuffer(rsSab, oldYSSIUSR0)

If oldYSSIUSR0.SSIUSRTLNK <> 0 Then
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
         & " where SSITXTNAT = '" & oldYSSIUSR0.SSIUSRNAT & "' and SSITXTUIDN = " & oldYSSIUSR0.SSIUSRUIDN _
         & " and SSITXTDIDX = '' and SSITXTUIDX= '' and SSITXTTLNK = " & oldYSSIUSR0.SSIUSRTLNK
    
    Set rsSab = cnsab.Execute(xSQL)
    If rsSab.EOF Then
        Call MsgBox(xSQL, vbCritical, "fraDetail_Load : inconnu")
    Else
        Call rsYSSITXT0_GetBuffer(rsSab, oldYSSITXT0_USR)
    End If
Else
        Call rsYSSITXT0_Init(oldYSSITXT0_USR)
        oldYSSITXT0_USR.SSITXTNAT = oldYSSIUSR0.SSIUSRNAT
        oldYSSITXT0_USR.SSITXTUIDN = oldYSSIUSR0.SSIUSRUIDN
End If

fraDetail_Display

End Sub

Public Sub fraYSSIDOM0_Load()
Dim xSQL As String, X As String, K As Integer

fraYSSIDOM0.Visible = False
txtRTF.Visible = False: txtRTF = ""
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
         & " where SSIDOMNAT = '" & oldYSSIDOM0.SSIDOMNAT & "' and SSIDOMUIDN = " & oldYSSIDOM0.SSIDOMUIDN _
         & " and SSIDOMDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "' and SSIDOMUIDX = '" & oldYSSIDOM0.SSIDOMUIDX & "'" _
         & " and SSIDOMUIDD = " & oldYSSIDOM0.SSIDOMUIDD
Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    Call MsgBox(xSQL, vbCritical, "fraYSSIDOM0_Load : YSSIDOM0 inconnu")
    Exit Sub
End If
Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)

If oldYSSIDOM0.SSIDOMTLNK <> 0 Then
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
         & " where SSITXTNAT = '" & oldYSSIDOM0.SSIDOMNAT & "' and SSITXTUIDN = " & oldYSSIDOM0.SSIDOMUIDN _
         & " and SSITXTDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "' and SSITXTUIDX = '" & oldYSSIDOM0.SSIDOMUIDX & "'" _
         & " and SSITXTUIDD = " & oldYSSIDOM0.SSIDOMUIDD & " and SSITXTTLNK = " & oldYSSIDOM0.SSIDOMTLNK
    
    Set rsSab = cnsab.Execute(xSQL)
    If rsSab.EOF Then
        Call MsgBox(xSQL, vbCritical, "fraYSSIDOM0_Load : YSSITXT0 inconnu")
        Call rsYSSITXT0_Init(oldYSSITXT0_DOM)
    Else
        Call rsYSSITXT0_GetBuffer(rsSab, oldYSSITXT0_DOM)
    End If
Else
        Call rsYSSITXT0_Init(oldYSSITXT0_DOM)
End If


blnHab2_SécuritéPhysique = arrHab(2)

Select Case oldYSSIDOM0.SSIDOMDIDX
    Case "IBM": Call fraYSSIDOM0_Load_IBM
    Case "SAA": Call fraYSSIDOM0_Load_SAA
    Case "SAB": Call fraYSSIDOM0_Load_SAB
    Case "SAB_W": Call fraYSSIDOM0_Load_SAB_W
    Case "WIN": Call fraYSSIDOM0_Load_WIN
    Case "DIV": Call fraYSSIDOM0_Load_DIV
        Select Case oldYSSIDIV0.SSIDIVDIDK
            Case "TEREN", "SG", "UGM": cmdProfil_Update_DIV.Visible = arrHab(5): blnHab2_SécuritéPhysique = arrHab(5)
            Case Else: cmdProfil_Update_DIV.Visible = arrHab(2)
        End Select
    Case "MEL": Call fraYSSIDOM0_Load_MEL
    Case "TIC": Call fraYSSIDOM0_Load_TIC
End Select
'____________________________________________________________________________________________________________
If blnHab2_SécuritéPhysique Then
    cmdProfil_Update.Visible = blnHab2_SécuritéPhysique
    If oldYSSIDOM0.SSIDOMUIDD <= 0 And Trim(oldYSSIDOM0.SSIDOMUIDX) = "" Then
        cmdProfil_Delete.Caption = "Supprimer le profil"
        cmdProfil_Delete.BackColor = vbRed
        cmdProfil_Delete.Visible = blnHab2_SécuritéPhysique
    Else
        If oldYSSIDOM0.SSIDOMDIDX <> "MEL" And oldYSSIDOM0.SSIDOMDIDX <> "SAB_W" Then
            cmdProfil_Delete.Visible = blnHab2_SécuritéPhysique
            cmdProfil_Delete.BackColor = mColor_W1
            cmdProfil_Delete.Caption = "Détacher le compte " & Trim(oldYSSIDOM0.SSIDOMUIDX)
        End If
    End If
End If
X = Trim(oldYSSIDOM0.SSIDOMDIDX)
fraProfil.Caption = oldYSSIUSR0.SSIUSRUIDX & "  :  " & X & "  :  " & oldYSSIDOM0.SSIDOMPRFX
cboProfil_DOM.Locked = True
For K = 1 To arrSSIDOMDIDX_UB
    If X = arrSSIDOMDIDX(K) Then cboProfil_DOM.ListIndex = K: Exit For
Next K
'____________________________________________________________________________________________________________
fraProfil.Visible = True
fraYSSIDOM0_Display
txtRTF.Visible = True
End Sub
Public Sub fraYSSIDOM0_Load_IBM()
Dim xSQL As String, X As String, K As Integer

'____________________________________________________________________________________________________________
If Trim(oldYSSIDOM0.SSIDOMPRFX) = "" Then
    oldYSSIIBM0.SSIIBMNAT = oldYSSIUSR0.SSIUSRNAT
    oldYSSIIBM0.SSIIBMUIDD = oldYSSIDOM0.SSIDOMUIDD
    usrYSSIIBM0.SSIIBMNAT = oldYSSIUSR0.SSIUSRNAT
    usrYSSIIBM0.SSIIBMUIDD = oldYSSIDOM0.SSIDOMUIDD
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBMH " _
         & " where SSIIBMNAT = '" & oldYSSIIBM0.SSIIBMNAT & " ' and SSIIBMUIDD = " & oldYSSIIBM0.SSIIBMUIDD _
         & " order by SSIIBMYVER desc"
    Set rsSab = cnsab.Execute(X)
    If rsSab.EOF Then
        Call cmdSSIIBM_Detail_Display("", "")
    Else
        xYSSIIBM0.SSIIBMNAT = rsSab("SSIIBMNAT")
        xYSSIIBM0.SSIIBMUIDD = rsSab("SSIIBMUIDD")
        xYSSIIBM0.SSIIBMYVER = rsSab("SSIIBMYVER")
        Call cmdSSIIBM_Detail_Display("YSSIIBMH", "USR")
    End If
    oldYSSIIBM0 = xYSSIIBM0

Else
    oldYSSIIBM0.UPUPRF = oldYSSIDOM0.SSIDOMPRFX
    oldYSSIIBM0.SSIIBMNAT = "$"
    usrYSSIIBM0.SSIIBMNAT = oldYSSIUSR0.SSIUSRNAT
    usrYSSIIBM0.SSIIBMUIDD = oldYSSIDOM0.SSIDOMUIDD
    'If Trim(oldYSSIDOM0.SSIDOMUIDX) = "" Then
    If oldYSSIDOM0.SSIDOMUIDD <= 0 Then
        Call cmdSSIIBM_Detail_Display("UPUPRF", "")
    Else
        Call cmdSSIIBM_Detail_Display("UPUPRF", "USR")
    End If
    oldYSSIIBM0 = xYSSIIBM0
End If
'____________________________________________________________________________________________________________

End Sub
Public Sub fraYSSIDOM0_Load_SAA()
Dim xSQL As String, X As String, K As Integer

'____________________________________________________________________________________________________________
If Trim(oldYSSIDOM0.SSIDOMPRFX) = "" Then
Else
    prfYSSISAA0.SSISAAUIDX = oldYSSIDOM0.SSIDOMPRFX
    prfYSSISAA0.SSISAANAT = "$"
    prfYSSISAA0.SSISAAUSEQ = 0
    usrYSSISAA0.SSISAANAT = oldYSSIUSR0.SSIUSRNAT
    usrYSSISAA0.SSISAAUIDX = oldYSSIDOM0.SSIDOMUIDX
    usrYSSISAA0.SSISAAUIDD = oldYSSIDOM0.SSIDOMUIDD
    
    xSQL = "select SSISAAUSEQ from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
         & " where SSISAANAT = '" & usrYSSISAA0.SSISAANAT & " ' and SSISAAUIDX = '" & usrYSSISAA0.SSISAAUIDX & "'" _
         & " and SSISAAUIDD = '" & usrYSSISAA0.SSISAAUIDD & "'"
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        usrYSSISAA0.SSISAAUSEQ = rsSab("SSISAAUSEQ")
    Else
        usrYSSISAA0.SSISAAUSEQ = 0
        Call MsgBox(xSQL, vbCritical, "Erreur lecture fraYSSIDOM0_Load_SAA")
    End If
    

    Call cmdSSISAA_Detail_Display("SAA")
End If
'____________________________________________________________________________________________________________

End Sub

Public Sub fraYSSIDOM0_Load_WIN()
Dim xSQL As String, X As String, K As Integer

'____________________________________________________________________________________________________________
If Trim(oldYSSIDOM0.SSIDOMPRFX) = "" Then
Else
    'prfYSSIWIN0.SSIWINUIDX = oldYSSIDOM0.SSIDOMPRFX
    'prfYSSIWIN0.SSIWINNAT = "$"
    rtfYSSIWIN0.SSIWINNAT = oldYSSIUSR0.SSIUSRNAT
    rtfYSSIWIN0.SSIWINUIDX = oldYSSIDOM0.SSIDOMUIDX
    rtfYSSIWIN0.SSIWINUIDD = oldYSSIDOM0.SSIDOMUIDD
    Call cmdSSIWIN_Detail_Display("YSSIWIN0_UIDD")
    usrYSSIWIN0 = rtfYSSIWIN0
End If
'____________________________________________________________________________________________________________

End Sub


Public Sub fraYSSIDOM0_Load_MEL()
Dim xSQL As String, X As String, K As Integer

'____________________________________________________________________________________________________________
If Trim(oldYSSIDOM0.SSIDOMPRFX) = "" Then
Else
    rtfYSSIMEL0.SSIMELNAT = oldYSSIUSR0.SSIUSRNAT
    rtfYSSIMEL0.SSIMELUIDX = oldYSSIDOM0.SSIDOMUIDX
    rtfYSSIMEL0.SSIMELUIDD = oldYSSIDOM0.SSIDOMUIDD
    Call cmdSSIMEL_Detail_Display("YSSIMEL0_UIDX")
    usrYSSIMEL0 = rtfYSSIMEL0
End If
'____________________________________________________________________________________________________________

End Sub

Public Sub fraYSSIDOM0_Load_TIC()
Dim xSQL As String, X As String, K As Integer

'____________________________________________________________________________________________________________
If Trim(oldYSSIDOM0.SSIDOMPRFX) = "" Then
Else
    rtfYSSITIC0.SSITICNAT = oldYSSIUSR0.SSIUSRNAT
    rtfYSSITIC0.SSITICUIDX = oldYSSIDOM0.SSIDOMUIDX
    rtfYSSITIC0.SSITICUIDD = oldYSSIDOM0.SSIDOMUIDD
    Call cmdSSITIC_Detail_Display("YSSITIC0_UIDX")
    usrYSSITIC0 = rtfYSSITIC0
End If
'____________________________________________________________________________________________________________

End Sub


Public Sub fraYSSIDOM0_Load_DIV()
Dim xSQL As String, X As String, K As Integer

'____________________________________________________________________________________________________________
If Trim(oldYSSIDOM0.SSIDOMPRFX) = "" Then
Else
    'prfYSSIDIV0.SSIDIVUIDX = oldYSSIDOM0.SSIDOMPRFX
    'prfYSSIDIV0.SSIDIVNAT = "$"
    rtfYSSIDIV0.SSIDIVNAT = oldYSSIUSR0.SSIUSRNAT
    rtfYSSIDIV0.SSIDIVUIDX = oldYSSIDOM0.SSIDOMUIDX
    rtfYSSIDIV0.SSIDIVUIDD = oldYSSIDOM0.SSIDOMUIDD
    Call cmdSSIDIV_Detail_Display("YSSIDIV0")
    usrYSSIDIV0 = rtfYSSIDIV0
    oldYSSIDIV0 = rtfYSSIDIV0
End If
'____________________________________________________________________________________________________________

End Sub

Public Sub fraYSSIDOM0_Load_SAB()
Dim xSQL As String, X As String, K As Integer

'____________________________________________________________________________________________________________
If Trim(oldYSSIDOM0.SSIDOMPRFX) = "" Then
Else
    prfYSSISAB0.SSISABUIDX = oldYSSIDOM0.SSIDOMPRFX
    prfYSSISAB0.SSISABNAT = "$"
    prfYSSISAB0.SSISABULOT = 0
    usrYSSISAB0.SSISABNAT = oldYSSIUSR0.SSIUSRNAT
    usrYSSISAB0.SSISABUIDX = oldYSSIDOM0.SSIDOMUIDX
    usrYSSISAB0.SSISABULOT = 0
    Call cmdSSISAB_Detail_Display("SAB")
End If
'____________________________________________________________________________________________________________

End Sub
Public Sub fraYSSIDOM0_Load_SAB_W()
Dim xSQL As String, X As String, K As Integer

'____________________________________________________________________________________________________________
If Trim(oldYSSIDOM0.SSIDOMPRFX) = "" Then
Else
    prfYSSISAB0.SSISABUIDX = oldYSSIDOM0.SSIDOMPRFX
    prfYSSISAB0.SSISABNAT = "$"
    prfYSSISAB0.SSISABULOT = 0
    usrYSSISAB0.SSISABNAT = oldYSSIUSR0.SSIUSRNAT
    usrYSSISAB0.SSISABUIDX = oldYSSIDOM0.SSIDOMUIDX
    usrYSSISAB0.SSISABULOT = 0
    Call cmdSSISAW_Detail_Display("YSSISAB0")
End If
'____________________________________________________________________________________________________________

End Sub




Public Sub cmdUpdate_Init()
mYSSIUSR0_Update = "": mYSSIDOM0_Update = "": mYSSITXT0_Update = ""
mYSSIIBM0_Update = "": mYSSIIBMH_Update = ""
mYSSITXT0_JRN_Update = ""
mYSSISAA0_Update = "": mYSSISAAH_Update = ""
mYSSIDOM0_Update_CMD = ""
mYSSISAB0_Update = "": mYSSISABH_Update = ""
mYSSIWIN0_Update = "": mYSSIWINH_Update = ""
mYSSIDIV0_Update = "": mYSSIDIVH_Update = ""
mYSSIMEL0_Update = "": mYSSIMELH_Update = ""
mYSSITIC0_Update = "": mYSSITICH_Update = ""
mYSSIDOM0_Update_CMD = "": mYSSIDOM0_Update_CMD_2 = ""
End Sub

Public Sub txtRTF_Visible()
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "{\f0\fswiss\fprq2\fcharset0 Calibri;}", "{\f0\fmodern\fprq1\fcharset0 Courier New;}")
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "\cf1\f0\fs20 a\cf2 b\cf3 c\cf4 d\cf5 e\cf6 f\cf7 g\cf8 h\cf9 i\cf10 j\cf11 k\cf12 l\cf13 m\cf14 n\cf15 o\cf16 p", "")
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", "")
txtRTF.Visible = True

End Sub

Public Function fraYSSIDOM0_Control()
Dim X As String, wMsgBox As String, wAMJ As Long
On Error GoTo Exit_sub

currentAction = "fraYSSIDOM0_Control"
Call cmdUpdate_Init

Call rsYSSIDOM0_Init(newYSSIDOM0)
'==========================
wMsgBox = ""
fraYSSIDOM0_Control = "?"
newYSSIDOM0 = oldYSSIDOM0
newYSSITXT0 = oldYSSITXT0_DOM
oldYSSITXT0_XXX = oldYSSITXT0_DOM

newYSSIDOM0.SSIDOMSTAK = Mid$(cboSSIDOMSTAK.Text, 1, 1)
newYSSIDOM0.SSIDOMPRFK = Mid$(cboSSIDOMPRFK, 1, 1)
If cboSSIDOMUNIT = "" Then
    newYSSIDOM0.SSIDOMUNIT = ""
Else
    newYSSIDOM0.SSIDOMUNIT = Mid$(cboSSIDOMUNIT, 1, 3)
End If

If chkSSIDOMDECH.Value = "1" Then
    Call DTPicker_Control(txtSSIDOMDECH, X)
    wAMJ = CLng(X)
    If wAMJ <> oldYSSIDOM0.SSIDOMDECH Then
        If wAMJ < DSys Then
            wMsgBox = wMsgBox & " - date échéance  < aujourd'hui" & vbCrLf
        Else
            newYSSIDOM0.SSIDOMDECH = wAMJ
            If newYSSIDOM0.SSIDOMPRFK = "!" Then newYSSIDOM0.SSIDOMPRFK = " "
        End If
    End If
Else
    newYSSIDOM0.SSIDOMDECH = 0
End If
If newYSSIDOM0.SSIDOMNAT = " " Then
    Select Case newYSSIDOM0.SSIDOMDIDX
        Case "IBM": fraYSSIDOM0_Control_IBM
        Case "SAA": fraYSSIDOM0_Control_SAA
        Case "SAB": fraYSSIDOM0_Control_SAB
        Case "WIN": fraYSSIDOM0_Control_WIN
        Case "DIV": fraYSSIDOM0_Control_DIV
        Case "TIC": fraYSSIDOM0_Control_TIC
    End Select
End If
newYSSITXT0.SSITXTINFO = Trim(txtSSIDOMTXT)
'____________________________________________________________________________________

If wMsgBox <> "" Then
    fraYSSIDOM0_Control = "?"
    Call MsgBox(wMsgBox, vbCritical, currentAction)
Else
    newYSSIDOM0.SSIDOMYAMJ = DSys
    newYSSIDOM0.SSIDOMYHMS = time_Hms
    newYSSIDOM0.SSIDOMYUSR = usrName_UCase
    fraYSSIDOM0_Control = Null
    If oldYSSIDOM0.SSIDOMUIDD = 0 Then  'oldYSSIDOM0.SSIDOMUIDX = "" Then
        newYSSIDOM0.SSIDOMYFCT = "CRE"
        mYSSIDOM0_Update = "New"
        If newYSSIDOM0.SSIDOMUIDD = 0 Then
        
            X = "select SSIDOMUIDD from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
         & "  Where SSIDOMNAT = '" & newYSSIDOM0.SSIDOMNAT & "' and SSIDOMUIDN = " & newYSSIDOM0.SSIDOMUIDN _
         & " and SSIDOMDIDX = '" & newYSSIDOM0.SSIDOMDIDX & "'  and SSIDOMUIDD< 0" _
         & " order by SSIDOMUIDD "
            Set rsSab = cnsab.Execute(X)

            If Not rsSab.EOF Then
                newYSSIDOM0.SSIDOMUIDD = rsSab("SSIDOMUIDD") - 1
            Else
                newYSSIDOM0.SSIDOMUIDD = -1
            End If
            Select Case newYSSIDOM0.SSIDOMDIDX
                Case "IBM":  mYSSIIBM0_Update = ""
                Case "SAA":  mYSSISAA0_Update = ""
                Case "SAB":  mYSSISAB0_Update = ""
                Case "WIN":  mYSSIWIN0_Update = ""
                Case "DIV":  mYSSIDIV0_Update = ""
                Case "TIC":  mYSSITIC0_Update = ""
            End Select
            
        
            'Select Case newYSSIDOM0.SSIDOMDIDX
            '    Case "IBM": newYSSIDOM0.SSIDOMUIDD = oldYSSIIBM0.SSIIBMUIDD: mYSSIIBM0_Update = ""
            '    Case "SAA": newYSSIDOM0.SSIDOMUIDD = prfYSSISAA0.SSISAAUIDD: mYSSISAA0_Update = ""
            '    Case "SAB": newYSSIDOM0.SSIDOMUIDD = -1: mYSSISAB0_Update = ""
            'End Select
        End If
        If newYSSIDOM0.SSIDOMNAT = "$" Then
            newYSSIDOM0.SSIDOMPRFK = " "
            newYSSIDOM0.SSIDOMUIDX = newYSSIDOM0.SSIDOMPRFX
        End If

    Else
        newYSSIDOM0.SSIDOMYFCT = "MOD"
        ''newYSSIDOM0.SSIDOMYVER = newYSSIDOM0.SSIDOMYVER + 1
        mYSSIDOM0_Update = "Update+H"
    End If
    If Trim(oldYSSITXT0_DOM.SSITXTINFO) <> Trim(newYSSITXT0.SSITXTINFO) Then
        mYSSITXT0_Update = "New"
        newYSSITXT0.SSITXTNAT = newYSSIDOM0.SSIDOMNAT
        newYSSITXT0.SSITXTUIDN = newYSSIDOM0.SSIDOMUIDN
        newYSSITXT0.SSITXTDIDX = newYSSIDOM0.SSIDOMDIDX
        newYSSITXT0.SSITXTUIDX = newYSSIDOM0.SSIDOMUIDX
        newYSSITXT0.SSITXTUIDD = newYSSIDOM0.SSIDOMUIDD
        newYSSITXT0.SSITXTYAMJ = DSys
        newYSSITXT0.SSITXTYHMS = time_Hms
        newYSSITXT0.SSITXTYUSR = usrName_UCase
    End If
    
End If

Exit_sub:


End Function

Public Function paramIBM_Init_Compte(lSSIDOMUIDN As Long, lSSIDOMPRFX As String, lSSIDOMUIDD As Long)
newYSSIDOM0.SSIDOMUIDN = lSSIDOMUIDN
If Trim(lSSIDOMPRFX) = "*NONE" Then
    newYSSIDOM0.SSIDOMPRFX = ""
    newYSSIDOM0.SSIDOMPRFK = "N"
Else
    newYSSIDOM0.SSIDOMPRFX = lSSIDOMPRFX
End If
newYSSIDOM0.SSIDOMUIDD = lSSIDOMUIDD
newYSSIDOM0.SSIDOMSTAK = newYSSIUSR0.SSIUSRSTAK
newYSSIDOM0.SSIDOMDECH = newYSSIUSR0.SSIUSRDECH

X = "select UPUPRF , UPSTAT from " & paramIBM_Library_SABSPE & ".YSSIIBM0 where SSIIBMNAT = '' and SSIIBMUIDD = " & newYSSIDOM0.SSIDOMUIDD
Set rsSab = cnsab.Execute(X)
If Trim(rsSab("UPSTAT")) = "*DISABLED" Then newYSSIDOM0.SSIDOMSTAK = "N"
newYSSIDOM0.SSIDOMUIDX = Trim(rsSab("UPUPRF"))

paramIBM_Init_Compte = sqlYSSIDOM0_Insert(newYSSIDOM0)

'If Trim(newYSSIDOM0.SSIDOMPRFX) = "" Or Trim(newYSSIDOM0.SSIDOMPRFX) = "SYSTEME" Then
If Trim(newYSSIDOM0.SSIDOMPRFX) = "SYSTEME" Then
    X = " "
Else
    X = "#"
End If
Dim Nb As Long
    X = "update " & paramIBM_Library_SABSPE & ".YSSIIBM0" _
         & " set SSIIBMPRFK= '" & X & "'" _
         & " where SSIIBMNAT = '' and SSIIBMUIDD = " & newYSSIDOM0.SSIDOMUIDD
        Call FEU_ROUGE
        Set rsSab = cnSab_Update.Execute(X, Nb)
        Call FEU_VERT
        If Nb = 0 Then
            Call MsgBox(Error, vbCritical, X)
        End If
    

End Function

Public Sub fraCompteH_Display()
'On Error Resume Next

txtCompteH_SSITXTINFO = ""
txtCompteH_SSITXTINFO.Locked = True
cmdCompteH_Update.Visible = False
lblCompteH_SSITXTINFO.Visible = False

Select Case Trim(oldYSSIDOM0.SSIDOMDIDX)
    Case "IBM"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBMH " _
             & " where SSIIBMNAT = ' '" _
             & " and SSIIBMUIDD  = " & oldYSSIDOM0.SSIDOMUIDD _
             & " order by SSIIBMYVER"
        Set rsSab = cnsab.Execute(X)
        
        Call fgCompteH_Display("IBM")
    Case "SAA"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAAH " _
             & " where SSISAANAT = ' '" _
             & " and SSISAAUIDX  = '" & oldYSSIDOM0.SSIDOMUIDX & "'" _
             & " order by SSISAAYVER"
        Set rsSab = cnsab.Execute(X)
        
        Call fgCompteH_Display("SAA")
    Case "SAB"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSISABH " _
             & " where SSISABNAT = ' '" _
             & " and SSISABUIDX  = '" & oldYSSIDOM0.SSIDOMUIDX & "'" _
             & " order by SSISABYVER"
        Set rsSab = cnsab.Execute(X)
        
        Call fgCompteH_Display("SAB")
    Case "SAB_W"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSISABH " _
             & " where SSISABNAT = 'W'" _
             & " and SSISABUIDX  = '" & oldYSSIDOM0.SSIDOMUIDX & "'" _
             & " order by SSISABYVER"
        Set rsSab = cnsab.Execute(X)
        
        Call fgCompteH_Display("SAB")
    Case "WIN"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIWINH " _
             & " where SSIWINNAT = ' '" _
             & " and SSIWINUIDD  = " & oldYSSIDOM0.SSIDOMUIDD _
             & " order by SSIWINYVER"
             
             '& " and SSIWINUIDX  like '%" & oldYSSIDOM0.SSIDOMUIDX & "%'" _

        Set rsSab = cnsab.Execute(X)
        
        Call fgCompteH_Display("WIN")
    Case "DIV"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIVH " _
             & " where SSIDIVNAT = ' '" _
             & " and SSIDIVUIDX  = '" & oldYSSIDOM0.SSIDOMUIDX & "'" _
             & " order by SSIDIVYVER"
        Set rsSab = cnsab.Execute(X)
        
        Call fgCompteH_Display("DIV")
    Case "MEL"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIMELH " _
             & " where SSIMELNAT = ' '" _
             & " and SSIMELUIDX  = '" & oldYSSIDOM0.SSIDOMUIDX & "'" _
             & " order by SSIMELYVER"
        Set rsSab = cnsab.Execute(X)
        
        Call fgCompteH_Display("MEL")
    Case "TIC"
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSITICH " _
             & " where SSITICNAT = ' '" _
             & " and SSITICUIDX  = '" & oldYSSIDOM0.SSIDOMUIDX & "'" _
             & " order by SSITICYVER"
        Set rsSab = cnsab.Execute(X)
        
        Call fgCompteH_Display("TIC")
End Select

fraCompteH.Visible = True
'fraCompteH.ZOrder 0
fraCompteH.Caption = "Historique du compte " & fraProfil.Caption
End Sub

Public Sub cmdSelect_SQL_3()
Dim V, X As String, K As Integer, XX As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_3"

arrCtl_K = 0

'_________________________________________________________________________________________________
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = "<Identifiant              |<Domaine                  |<Compte                             |<Profil                            " _
           & "|<Champ                                |<Non Conforme                                                                     |<Référence                                                                  |||"
fgSelect.Rows = 1
fgSelect.Row = 0
'================================================================================================

arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "comptes Windows SUPPRIMES"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
     & " where SSIWINPRFK = 'S' and SSIWINYFCT = 'SUP' order by SSIWINUIDX"

Set rsSab = cnsab.Execute(xSQL)

Call fgSelect_Display_3_Orphelins("WIN_S")
Call fgSelect_Display_3_Total(vbRed)
'================================================================================================

arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "comptes Windows orphelins"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
     & " where SSIWINPRFK = '?'  order by SSIWINUIDX"

Set rsSab = cnsab.Execute(xSQL)

Call fgSelect_Display_3_Orphelins("WIN")
Call fgSelect_Display_3_Total(mColor_W1)
'================================================================================================

arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "comptes IBM SUPPRIMES"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0" _
     & " where SSIIBMPRFK = 'S' and SSIIBMYFCT = 'SUP' order by UPUPRF"

Set rsSab = cnsab.Execute(xSQL)

Call fgSelect_Display_3_Orphelins("IBM_S")
Call fgSelect_Display_3_Total(vbRed)
'================================================================================================
xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSIIBMH" _
     & " where SSIIBMPRFK = '?'"

Set rsSab = cnsab.Execute(xSQL)
If rsSab(0) > 0 Then
    arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "comptes IBM_H orphelins"
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBMH" _
         & " where SSIIBMPRFK = '?' order by UPUPRF "
    
    Set rsSab = cnsab.Execute(xSQL)
    
    Call fgSelect_Display_3_Orphelins("IBM_H")
    Call fgSelect_Display_3_Total(mColor_Y2)
End If
'================================================================================================

arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "comptes IBM orphelins"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0" _
     & " where SSIIBMPRFK = '?'  order by UPUPRF"

Set rsSab = cnsab.Execute(xSQL)

Call fgSelect_Display_3_Orphelins("IBM")
Call fgSelect_Display_3_Total(mColor_W1)
'================================================================================================
arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "comptes SAA orphelins"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0" _
     & " where SSISAAPRFK = '?'" _
     & " union select * from " & paramIBM_Library_SABSPE & ".YSSISAAH" _
     & " where SSISAAPRFK = '?' order by SSISAAUIDX"

Set rsSab = cnsab.Execute(xSQL)

Call fgSelect_Display_3_Orphelins("SAA")
Call fgSelect_Display_3_Total(mColor_W1)
'================================================================================================
arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "comptes SAB_H orphelins"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISABH" _
     & " where SSISABnat = ' ' and SSISABPRFK = '?' " _
     & " order by SSISABUIDX"

Set rsSab = cnsab.Execute(xSQL)

Call fgSelect_Display_3_Orphelins("SAB_H")
Call fgSelect_Display_3_Total(mColor_Y1)

arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "comptes SAB orphelins"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0" _
     & " where SSISABnat = ' ' and SSISABPRFK = '?' " _
     & " order by SSISABUIDX"

Set rsSab = cnsab.Execute(xSQL)

Call fgSelect_Display_3_Orphelins("SAB")
Call fgSelect_Display_3_Total(mColor_W1)
'================================================================================================

arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "comptes DIV orphelins"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0" _
     & " where SSIDIVPRFK = '?' order by SSIDIVDIDK , SSIDIVUIDX"

Set rsSab = cnsab.Execute(xSQL)

Call fgSelect_Display_3_Orphelins("DIV")
Call fgSelect_Display_3_Total(mColor_W1)

'================================================================================================
arrCtl_K = arrCtl_K + 1
 arrCtl_Lib(arrCtl_K) = "comptes à désactiver" ' (X compte clos)"
'_________________________________________________________________________________________________
'xWhere = " where SSIDOMSTAK <> ' ' and SSIDOMPRFK <> 'X' and SSIDOMPRFX <> ''" _
       & " and SSIDOMUIDN = SSIUSRUIDN"
 'xWhere = " where SSIDOMSTAK = 'N' and SSIDOMPRFK <> 'X' and SSIDOMDIDX <> 'IBM'" _
 '      & " and SSIDOMUIDN = SSIUSRUIDN"
 xWhere = " where SSIUSRSTAK <> 'N' and SSIDOMSTAK = 'N' and SSIDOMPRFK <> 'X'" _
       & " and SSIDOMUIDN = SSIUSRUIDN"
      
X = Trim(txtSelect_Options_1_SSIUSRUIDX)
If X <> "" Then xWhere = xWhere & " and SSIUSRUIDX like '" & X & "%'"

X = Trim(cboSelect_Options_1_SSIDOMDIDX)
If X <> "" Then xWhere = xWhere & " and SSIDOMDIDX = '" & X & "'"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 , " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & xWhere _
     & " order by SSIUSRUIDX , SSIDOMDIDX , SSIDOMUIDX"

Set rsSab = cnsab.Execute(xSQL)

Call fgSelect_Display_3_Inactifs("compte SSI inactif")
Call fgSelect_Display_3_Total(mColor_W1)


'================================================================================================
arrCtl_K = arrCtl_K + 1
arrCtl_Lib(arrCtl_K) = "comptes SSI désactivés"
'_________________________________________________________________________________________________
xWhere = " where SSIUSRNAT = ' ' and SSIUSRSTAK = 'N' and SSIDOMPRFK <> 'X'" _
       & " and SSIDOMUIDN = SSIUSRUIDN"
       
X = Trim(txtSelect_Options_1_SSIUSRUIDX)
If X <> "" Then xWhere = xWhere & " and SSIUSRUIDX like '" & X & "%'"

X = Trim(cboSelect_Options_1_SSIDOMDIDX)
If X <> "" Then xWhere = xWhere & " and SSIDOMDIDX = '" & X & "'"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 , " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & xWhere _
     & " order by SSIUSRUIDX , SSIDOMDIDX , SSIDOMUIDX"

Set rsSab = cnsab.Execute(xSQL)

Call fgSelect_Display_3_Inactifs("comptes SSI désactivés")
Call fgSelect_Display_3_Total(mColor_W1)
'================================================================================================
arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "utilisateurs non conformes au modèle BIA"

xWhere = " where SSIUSRPRFK not in (' ' , 'X') "
       
X = Trim(txtSelect_Options_1_SSIUSRUIDX)
If X <> "" Then xWhere = xWhere & " and SSIUSRUIDX like '" & X & "%'"


xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & xWhere _
     & " order by SSIUSRUIDX "


Set rsSab = cnsab.Execute(xSQL)

Call fgSelect_Display_3_SSIUSRPRFK
Call fgSelect_Display_3_Total(mColor_W1)
'================================================================================================
arrCtl_K = arrCtl_K + 1
arrCtl_Lib(arrCtl_K) = "profils à affecter"
'_________________________________________________________________________________________________
xWhere = " where SSIDOMSTAK = ' ' and SSIDOMUIDX = ''" _
       & " and SSIDOMUIDN = SSIUSRUIDN"
       
X = Trim(txtSelect_Options_1_SSIUSRUIDX)
If X <> "" Then xWhere = xWhere & " and SSIUSRUIDX like '" & X & "%'"

X = Trim(cboSelect_Options_1_SSIDOMDIDX)
If X <> "" Then xWhere = xWhere & " and SSIDOMDIDX = '" & X & "'"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 , " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & xWhere _
     & " order by SSIUSRUIDX , SSIDOMDIDX , SSIDOMUIDX"

Set rsSab = cnsab.Execute(xSQL)

Call fgSelect_Display_3_Inactifs("profils à affecter")
Call fgSelect_Display_3_Total(mColor_W1)


'================================================================================================
arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "comptes non conformes au profil du domaine"
'____________________________________________________________________________________________
xWhere = " where SSIDOMUIDN = SSIUSRUIDN " _
       & " and SSIDOMPRFK not in ( ' ', 'X')"
       
X = Trim(txtSelect_Options_1_SSIUSRUIDX)
If X <> "" Then xWhere = xWhere & " and SSIUSRUIDX like '" & X & "%'"

X = Trim(cboSelect_Options_1_SSIDOMDIDX)
If X <> "" Then xWhere = xWhere & " and SSIDOMDIDX = '" & X & "'"

X = Trim(cboSelect_Options_1_SSIDOMPRFX)
If X <> "" Then
    blnSelect_Options_1_SSIDOMPRFX = True
    If Mid$(X, 1, 1) <> "*" Then
        xWhere = xWhere & " and SSIDOMPRFX = '" & X & "'"
    End If
End If

Select Case Mid$(cboSelect_Options_1_SSIUSRSTAK, 1, 1)
    Case " ": xWhere = xWhere & " and SSIDOMSTAK = ' '"
    Case "N": xWhere = xWhere & " and SSIDOMSTAK = 'N'"
End Select

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0, " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
     & xWhere _
     & " order by SSIUSRUIDX , SSIDOMPRFX"


Set rsSab = cnsab.Execute(xSQL)

fgSelect_Display_3

Call fgSelect_Display_3_Total(mColor_W1)
'================================================================================================

arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "paramétrage envoi automatique de courriels"

Call lstParam_SSIMELUNOM_Load("")

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
     & " Where SSIMELNAT = '@' and SSIMELPRFK <> 'X'" _
     & " order by SSIMELUIDX "
Set rsSab = cnsab.Execute(xSQL)

fgSelect_Display_3_MEL

Call fgSelect_Display_3_Total(mColor_W1)
'================================================================================================

arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "comptes Athic SUPPRIMES"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
     & " where SSITICPRFK = 'S' and SSITICYFCT = 'SUP' order by SSITICUIDX"

Set rsSab = cnsab.Execute(xSQL)

Call fgSelect_Display_3_Orphelins("TIC_S")
Call fgSelect_Display_3_Total(vbRed)
'================================================================================================

arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "comptes Athic orphelins"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
     & " where SSITICPRFK = '?'  order by SSITICUIDX"

Set rsSab = cnsab.Execute(xSQL)

Call fgSelect_Display_3_Orphelins("TIC")
Call fgSelect_Display_3_Total(mColor_W1)
'================================================================================================



Call cmdSelect_SQL_3_H("Suite")

fgSelect.Visible = True

Set rsSab = Nothing

Exit Sub

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub cmdSelect_SQL_rJrnl()
Dim xSQL As String, xWhere As String, X As String
Dim wDateFrom, wDateTo
On Error GoTo Error_Handler
X = dateImp10_S(wAmjMin)
wDateFrom = 1200802192 - DateDiff("s", "01/01/2000 00:00:00", X & " 00:00:00")

wDateTo = 1200802192 - DateDiff("s", "01/01/2000 00:00:00", X & " 23:59:59")

xWhere = " and jrnl_comp_name = 'BSA'"

cnSIDE_DB.Open paramODBC_DSN_SIDE_DB
    
xSQL = "select * from rJrnl " _
          & "where Aid = 0 " _
          & " and jrnl_rev_date_time >= " & wDateTo _
          & " and jrnl_rev_date_time <= " & wDateFrom _
          & xWhere _
          & " order by  jrnl_rev_date_time desc , jrnl_seq_nbr desc"
          
'          & " and substring(jrnl_display_text,1,7) = 'Message'" _

Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
  
fgSelect_Display_rJrnl

GoTo Exit_sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
Exit_sub:
    cnSIDE_DB.Close
    Set cnSIDE_DB = Nothing

End Sub

Private Sub fgSelect_Display_rJrnl()
Dim wColor As Long, K As Integer
Dim V, mComp_Name, blnDisplay As Boolean, X As String
Dim I As Long

On Error GoTo Error_Handler ' Resume Next  '
'SSTab1.Tab = 0
fgSelect.Visible = False
currentAction = "fgSelect_Display_rJrnl"
    
Do While Not rsSIDE_DB.EOF
    I = rsSIDE_DB("Jrnl_event_num")
    V = rsSIDE_DB("Jrnl_oper_nickname")
    If Not IsNull(V) Then
        X = V
    Else
        X = ""
    End If
    blnDisplay = True
    If I < 3005 Then
       If X = "LSO" Or X = "RSO" Or X = "SUPER" Then
       Else
            If InStr(1, X, "SUPER") > 0 Then
            Else
                blnDisplay = False
            End If
        End If
    End If
    If blnDisplay Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        'fgSelect.Col = 6:  V = rsSIDE_DB("Jrnl_event_severity")
        
        
        fgSelect.Col = 0:  fgSelect.Text = "SAA Journal"
        fgSelect.Col = 1: fgSelect.Text = I
        
        
        fgSelect.Col = 2:  V = rsSIDE_DB("Jrnl_oper_nickname"): If Not IsNull(V) Then fgSelect.Text = V
        V = rsSIDE_DB("Jrnl_event_name"): If Not IsNull(V) Then fgSelect.Text = fgSelect.Text & " " & V
        fgSelect.Col = 3:  V = rsSIDE_DB("Jrnl_event_class"): If Not IsNull(V) Then fgSelect.Text = V
        fgSelect.Col = 4:  V = rsSIDE_DB("Jrnl_merged_text"): If Not IsNull(V) Then fgSelect.Text = V
        'fgSelect.Col = 5:  V = rsSIDE_DB("Jrnl_event_name"): If Not IsNull(V) Then fgSelect.Text = V
        fgSelect.Col = 6:  V = rsSIDE_DB("Jrnl_date_time"): If Not IsNull(V) Then fgSelect.Text = V

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
        'fgSelect.Col = 9:  V = rsSIDE_DB("Jrnl_alarm_status"): If Not IsNull(V) Then fgSelect.Text = V
        'fgSelect.Col = 10:  V = rsSIDE_DB("Aid"): If Not IsNull(V) Then fgSelect.Text = V
        'fgSelect.Col = 11:  V = rsSIDE_DB("Jrnl_rev_date_time"): If Not IsNull(V) Then fgSelect.Text = V
        'fgSelect.Col = 12:  V = rsSIDE_DB("Jrnl_seq_nbr"): If Not IsNull(V) Then fgSelect.Text = V
    End If
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

Public Sub cmdSelect_SQL_J()
Dim V, X As String, K As Integer, XX As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler


currentAction = "cmdSelect_SQL_J"
xWhere = ""

If Not IsNull(txtSelect_Options_J_SSITXTYMAJ.Value) Then
    Call DTPicker_Control(txtSelect_Options_J_SSITXTYMAJ, wAmjMin)
    xWhere = xWhere & " and SSITXTYAMJ = " & wAmjMin
End If

X = Trim(cboSelect_Options_J_SSIDOMDIDX)
If X <> "" Then xWhere = xWhere & " and SSITXTDIDX = '" & X & "'"
  

X = Trim(libSelect_Options_J_SSIUSRUIDX)
If X <> "" Then
    xWhere = xWhere & " and SSITXTUIDN = " & Val(X)
Else
    xWhere = xWhere & " and SSITXTUIDX like '%" & Trim(txtSelect_Options_J_SSIUSRUIDX) & "%'"
    
End If

If xWhere = "" Then
    V = "Préciser les critères de recherche"
    GoTo Error_MsgBox
End If

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSITXT0" _
     & "  where SSITXTNAT = 'J'" & xWhere _
     & " order by SSITXTYAMJ , SSITXTYHMS , SSITXTTLNK"


Set rsSab = cnsab.Execute(xSQL)
   

fgSelect_Display_J

If Not IsNull(txtSelect_Options_J_SSITXTYMAJ.Value) Then Call cmdSelect_SQL_rJrnl

Set rsSab = Nothing

Exit Sub

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub cmdSelect_SQL_H()
Dim V, X As String, K As Integer, XX As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler


currentAction = "cmdSelect_SQL_H"
xWhere = ""

   
Select Case Trim(cboSelect_Options_4_SSIDOMDIDX)
    Case "IBM": fgSelect_Display_H_IBM
    Case "SAA": fgSelect_Display_H_SAA
    Case "SAB": fgSelect_Display_H_SAB
    Case "WIN": fgSelect_Display_H_WIN
    Case "USR": fgSelect_Display_H_USR
    Case "DOM": fgSelect_Display_H_DOM
    Case "DIV": fgSelect_Display_H_DIV
    Case "MEL": fgSelect_Display_H_MEL
    Case "TIC": fgSelect_Display_H_TIC
End Select

Set rsSab = Nothing

Exit Sub

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub cmdSelect_SQL_3_H(lFct As String)
Dim V, X As String, K As Integer, XX As String, Nb As Integer
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_3_H"
'================================================================================================

arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "Modifications WIN archivées NON VALIDEES"
    
'_________________________________________________________________________________________________
If lFct = "" Then
    fgSelect.Visible = False
    fgSelect_Reset
    fgSelect.FormatString = "<Identifiant              |<Domaine                  |<Compte                             |<Profil                            " _
           & "|<Champ                                |<Non Conforme                                               |<Référence                                                                  |||"
    fgSelect.Rows = 1
    fgSelect.Row = 0
'Else

End If
'================================================================================================

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIWINH, " & paramIBM_Library_SABSPE & ".YSSIDOM0, " _
     & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & " where SSIWINPRFK <> ' ' and SSIWINYFCT <> 'VU'" _
     & " and SSIDOMNAT = SSIWINNAT  and SSIDOMDIDX = 'WIN' and SSIDOMUIDD = SSIWINUIDD " _
     & " and SSIUSRNAT = SSIWINNAT and SSIUSRUIDN = SSIDOMUIDN " _
     & " order by SSIUSRUIDX , SSIDOMPRFX"

Set rsSab = cnsab.Execute(xSQL)

fgSelect_Display_3
Call fgSelect_Display_3_Total(vbYellow)

'================================================================================================
arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "Modifications IBM archivées NON VALIDEES"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBMH, " & paramIBM_Library_SABSPE & ".YSSIDOM0, " _
     & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & " where SSIIBMPRFK <> ' ' and SSIIBMYFCT <> 'VU'" _
     & " and SSIDOMNAT = SSIIBMNAT  and SSIDOMDIDX = 'IBM' and SSIDOMUIDD = SSIIBMUIDD " _
     & " and SSIUSRNAT = SSIIBMNAT and SSIUSRUIDN = SSIDOMUIDN " _
     & " order by SSIUSRUIDX , SSIDOMPRFX"

Set rsSab = cnsab.Execute(xSQL)

fgSelect_Display_3
Call fgSelect_Display_3_Total(vbYellow)
'================================================================================================
arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "Modifications SAB archivées NON VALIDEES"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISABH, " & paramIBM_Library_SABSPE & ".YSSIDOM0, " _
     & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & " where SSISABPRFK <> ' ' and SSISABYFCT <> 'VU'" _
     & " and SSIDOMNAT = SSISABNAT  and SSIDOMDIDX = 'SAB' and SSIDOMUIDD = SSISABUIDD " _
     & " and SSIUSRNAT = SSISABNAT and SSIUSRUIDN = SSIDOMUIDN " _
     & " order by SSIUSRUIDX , SSIDOMPRFX"

Set rsSab = cnsab.Execute(xSQL)

fgSelect_Display_3
Call fgSelect_Display_3_Total(vbYellow)
'================================================================================================
arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "Modifications SAB_W (SWIFT) archivées NON VALIDEES"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISABH, " & paramIBM_Library_SABSPE & ".YSSIDOM0, " _
     & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & " where SSISABNAT = 'W' and SSISABPRFK <> ' ' and SSISABYFCT <> 'VU'" _
     & " and SSIDOMDIDX = 'SAB_W' and SSIDOMUIDD = SSISABUIDD " _
     & " and SSIUSRUIDN = SSIDOMUIDN " _
     & " order by SSIUSRUIDX , SSIDOMPRFX"

Set rsSab = cnsab.Execute(xSQL)

fgSelect_Display_3
Call fgSelect_Display_3_Total(vbYellow)

'================================================================================================
arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "Modifications SAA archivées NON VALIDEES"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAAH, " & paramIBM_Library_SABSPE & ".YSSIDOM0, " _
     & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & " where SSISAAPRFK <> ' ' and SSISAAYFCT <> 'VU'" _
     & " and SSIDOMNAT = SSISAANAT  and SSIDOMDIDX = 'SAA' and SSIDOMUIDD = SSISAAUIDD " _
     & " and SSIUSRNAT = SSISAANAT and SSIUSRUIDN = SSIDOMUIDN " _
     & " order by SSIUSRUIDX , SSIDOMPRFX"

Set rsSab = cnsab.Execute(xSQL)

fgSelect_Display_3
Call fgSelect_Display_3_Total(vbYellow)
'================================================================================================
arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "Modifications DIV archivées NON VALIDEES"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIVH, " & paramIBM_Library_SABSPE & ".YSSIDOM0, " _
     & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & " where SSIDIVPRFK <> ' ' and SSIDIVYFCT <> 'VU'" _
     & " and SSIDOMNAT = SSIDIVNAT  and SSIDOMDIDX = 'DIV' and SSIDOMUIDX = SSIDIVUIDX and SSIDOMUIDD = SSIDIVUIDD " _
     & " and SSIUSRNAT = SSIDIVNAT and SSIUSRUIDN = SSIDOMUIDN " _
     & " order by SSIUSRUIDX , SSIDOMPRFX"

Set rsSab = cnsab.Execute(xSQL)

fgSelect_Display_3
Call fgSelect_Display_3_Total(vbYellow)
'================================================================================================
arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "Modifications MEL archivées NON VALIDEES"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIMELH, " & paramIBM_Library_SABSPE & ".YSSIDOM0, " _
     & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & " where SSIMELPRFK <> ' ' and SSIMELYFCT <> 'VU'" _
     & " and SSIDOMNAT = SSIMELNAT  and SSIDOMDIDX = 'MEL' and SSIDOMUIDX = SSIMELUIDX and SSIDOMUIDD = SSIMELUIDD " _
     & " and SSIUSRNAT = SSIMELNAT and SSIUSRUIDN = SSIDOMUIDN " _
     & " order by SSIUSRUIDX , SSIDOMPRFX"

Set rsSab = cnsab.Execute(xSQL)

fgSelect_Display_3
Call fgSelect_Display_3_Total(vbYellow)

'================================================================================================
arrCtl_K = arrCtl_K + 1: arrCtl_Lib(arrCtl_K) = "Modifications TIC archivées NON VALIDEES"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSITICH, " & paramIBM_Library_SABSPE & ".YSSIDOM0, " _
     & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & " where SSITICPRFK <> ' ' and SSITICYFCT <> 'VU'" _
     & " and SSIDOMNAT = SSITICNAT  and SSIDOMDIDX = 'TIC' and SSIDOMUIDD = SSITICUIDD " _
     & " and SSIUSRNAT = SSITICNAT and SSIUSRUIDN = SSIDOMUIDN " _
     & " order by SSIUSRUIDX , SSIDOMPRFX"

Set rsSab = cnsab.Execute(xSQL)

fgSelect_Display_3
Call fgSelect_Display_3_Total(vbYellow)

'================================================================================================
fgSelect.Visible = True

Set rsSab = Nothing

Exit Sub

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub fgSelect_Display_3_Line_YSSIIBM0(lLnk As String)
Dim X As String
If xYSSIDOM0.SSIDOMPRFX = "" Then
    Call rsYSSIIBM0_Init(xYSSIIBM0)
Else
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0 " _
         & " where SSIIBMNAT = '$' and UPUPRF = '" & xYSSIDOM0.SSIDOMPRFX & "'"
    
    Set rsSab_X = cnsab.Execute(X)
    
    If rsSab.EOF Then
        If Not blnAuto Then
            Call MsgBox(xYSSIDOM0.SSIDOMPRFX & " : Profil inconnu dans le domaine " & vbCrLf & X, vbCritical, "fgSelect_Display_3_Line_YSSIIBM0")
        End If
        GoTo Exit_sub
    End If
    Call rsYSSIIBM0_GetBuffer(rsSab_X, xYSSIIBM0)
End If

X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0 where SSIIBMNAT = '' and SSIIBMUIDD = " & xYSSIDOM0.SSIDOMUIDD
Set rsSab_X = cnsab.Execute(X)

If rsSab.EOF Then
    If Not blnAuto Then
        Call MsgBox(xYSSIDOM0.SSIDOMUIDX & " : Compte inconnu dans le domaine " & vbCrLf & X, vbCritical, "fgSelect_Display_3_Line_YSSIIBM0")
    End If
    GoTo Exit_sub
End If
Call rsYSSIIBM0_GetBuffer(rsSab_X, usrYSSIIBM0)
fgSelect.Col = 6: fgSelect.Text = usrYSSIIBM0.UPTEXT: fgSelect.CellBackColor = mColor_Y2
fgSelect.Col = 5: fgSelect.Text = "Connexion " & dateImp10_S(usrYSSIIBM0.UPPSOD) & ", Créé " & dateImp10_S(usrYSSIIBM0.UPCRTD): fgSelect.CellBackColor = &HC0E0FF
'fgSelect.CellFontBold = True

If usrYSSIIBM0.UPUSCL <> xYSSIIBM0.UPUSCL Then Call fgSelect_Display_3_Line(lLnk, "User class", usrYSSIIBM0.UPUSCL, xYSSIIBM0.UPUSCL)
If usrYSSIIBM0.UPPWEI <> xYSSIIBM0.UPPWEI Then Call fgSelect_Display_3_Line(lLnk, "Pwd expir. interval", CStr(usrYSSIIBM0.UPPWEI), CStr(xYSSIIBM0.UPPWEI))
If usrYSSIIBM0.UPPWEX <> xYSSIIBM0.UPPWEX Then Call fgSelect_Display_3_Line(lLnk, "Pwd *None *Yes *No", usrYSSIIBM0.UPPWEX, xYSSIIBM0.UPPWEX)
If usrYSSIIBM0.UPPWON <> xYSSIIBM0.UPPWON Then Call fgSelect_Display_3_Line(lLnk, "Pwd set expired", usrYSSIIBM0.UPPWON, xYSSIIBM0.UPPWON)
If usrYSSIIBM0.UPSPAU <> xYSSIIBM0.UPSPAU Then Call fgSelect_Display_3_Line(lLnk, "Special authorities ", usrYSSIIBM0.UPSPAU, xYSSIIBM0.UPSPAU)
If usrYSSIIBM0.UPINPG <> xYSSIIBM0.UPINPG Then Call fgSelect_Display_3_Line(lLnk, "Initial program", usrYSSIIBM0.UPINPG, xYSSIIBM0.UPINPG)
If usrYSSIIBM0.UPINPL <> xYSSIIBM0.UPINPL Then Call fgSelect_Display_3_Line(lLnk, "Initial program lib", usrYSSIIBM0.UPINPL, xYSSIIBM0.UPINPL)
If usrYSSIIBM0.UPJBDS <> xYSSIIBM0.UPJBDS Then Call fgSelect_Display_3_Line(lLnk, "Job description", usrYSSIIBM0.UPJBDS, xYSSIIBM0.UPJBDS)
If usrYSSIIBM0.UPJBDL <> xYSSIIBM0.UPJBDL Then Call fgSelect_Display_3_Line(lLnk, "Job description lib", usrYSSIIBM0.UPJBDL, xYSSIIBM0.UPJBDL)
If usrYSSIIBM0.UPGRPF <> xYSSIIBM0.UPGRPF Then Call fgSelect_Display_3_Line(lLnk, "Group profile ", usrYSSIIBM0.UPGRPF, xYSSIIBM0.UPGRPF)
If usrYSSIIBM0.UPGRAU <> xYSSIIBM0.UPGRAU Then Call fgSelect_Display_3_Line(lLnk, "Group authority", usrYSSIIBM0.UPGRAU, xYSSIIBM0.UPGRAU)
If usrYSSIIBM0.UPSPEN <> xYSSIIBM0.UPSPEN Then Call fgSelect_Display_3_Line(lLnk, "Special environ.", usrYSSIIBM0.UPSPEN, xYSSIIBM0.UPSPEN)
If usrYSSIIBM0.UPCRLB <> xYSSIIBM0.UPCRLB Then Call fgSelect_Display_3_Line(lLnk, "Current library", usrYSSIIBM0.UPCRLB, xYSSIIBM0.UPCRLB)
If usrYSSIIBM0.UPINMN <> xYSSIIBM0.UPINMN Then Call fgSelect_Display_3_Line(lLnk, "Initial menu", usrYSSIIBM0.UPINMN, xYSSIIBM0.UPINMN)
If usrYSSIIBM0.UPINML <> xYSSIIBM0.UPINML Then Call fgSelect_Display_3_Line(lLnk, "Initial menu lib", usrYSSIIBM0.UPINML, xYSSIIBM0.UPINML)
If usrYSSIIBM0.UPLTCP <> xYSSIIBM0.UPLTCP Then Call fgSelect_Display_3_Line(lLnk, "Limited capability", usrYSSIIBM0.UPLTCP, xYSSIIBM0.UPLTCP)
If usrYSSIIBM0.UPATPG <> xYSSIIBM0.UPATPG Then Call fgSelect_Display_3_Line(lLnk, "Attention program ", usrYSSIIBM0.UPATPG, xYSSIIBM0.UPATPG)
If usrYSSIIBM0.UPATPL <> xYSSIIBM0.UPATPL Then Call fgSelect_Display_3_Line(lLnk, "Attention program lib", usrYSSIIBM0.UPATPL, xYSSIIBM0.UPATPL)

Exit_sub:

End Sub


Public Sub fgSelect_Display_3_Line(lLnk As String, lField As String, lV1 As String, lV2 As String)

fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
fgSelect.Col = 4: fgSelect.Text = lField
fgSelect.Col = 5: fgSelect.Text = lV1: fgSelect.CellBackColor = mColor_Y0
fgSelect.CellForeColor = vbMagenta

fgSelect.Col = 6: fgSelect.Text = lV2 ': fgSelect.CellBackColor = mColor_G1
fgSelect.CellForeColor = vbBlue

fgSelect.Col = 9: fgSelect.Text = lLnk
End Sub

Public Function cmdSelect_SQL_9_SSIIBMPRFK_Control(lSSIDOMPRFX As String) As String
Dim xSQL As String

    cmdSelect_SQL_9_SSIIBMPRFK_Control = "N"
    If Trim(lSSIDOMPRFX) = "" Then
        Call rsYSSIIBM0_Init(prfYSSIIBM0)
    Else
        If lSSIDOMPRFX <> prfYSSIIBM0.UPUPRF Then
            xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0" _
                 & " where  SSIIBMNAT = '$' and UPUPRF = '" & lSSIDOMPRFX & "'"
            Set rsSab_X = cnsab.Execute(xSQL)
            If Not rsSab_X.EOF Then
                Call rsYSSIIBM0_GetBuffer(rsSab_X, prfYSSIIBM0)
            Else
                Call rsYSSIIBM0_Init(prfYSSIIBM0)
            End If
        End If
    End If
    
    If oldYSSIDOM0.SSIDOMPRFX = "" Then
    Else
        If Trim(oldYSSIIBM0.UPINPL) = "*LIBL" Or Trim(oldYSSIIBM0.UPINPL) = "QSYS" Then
            xYSSIIBM0.UPINPL = oldYSSIIBM0.UPINPL
        Else
            xYSSIIBM0.UPINPL = prfYSSIIBM0.UPINPL
        End If
        
        '2013-07-22 And oldYSSIIBM0.UPPWEX = prfYSSIIBM0.UPPWEX _
        '2013-07-22 And oldYSSIIBM0.UPSTAT = prfYSSIIBM0.UPSTAT Then
        If oldYSSIIBM0.UPUSCL = prfYSSIIBM0.UPUSCL _
        And oldYSSIIBM0.UPPWEI = prfYSSIIBM0.UPPWEI _
        And oldYSSIIBM0.UPPWON = prfYSSIIBM0.UPPWON _
        And oldYSSIIBM0.UPSPAU = prfYSSIIBM0.UPSPAU _
        And oldYSSIIBM0.UPINPG = prfYSSIIBM0.UPINPG _
        And oldYSSIIBM0.UPINPL = xYSSIIBM0.UPINPL _
        And oldYSSIIBM0.UPJBDS = prfYSSIIBM0.UPJBDS _
        And oldYSSIIBM0.UPJBDL = prfYSSIIBM0.UPJBDL _
        And oldYSSIIBM0.UPGRPF = prfYSSIIBM0.UPGRPF _
        And oldYSSIIBM0.UPGRAU = prfYSSIIBM0.UPGRAU _
        And oldYSSIIBM0.UPSPEN = prfYSSIIBM0.UPSPEN _
        And oldYSSIIBM0.UPCRLB = prfYSSIIBM0.UPCRLB _
        And oldYSSIIBM0.UPINMN = prfYSSIIBM0.UPINMN _
        And oldYSSIIBM0.UPINML = prfYSSIIBM0.UPINML _
        And oldYSSIIBM0.UPLTCP = prfYSSIIBM0.UPLTCP _
        And oldYSSIIBM0.UPATPG = prfYSSIIBM0.UPATPG _
        And oldYSSIIBM0.UPATPL = prfYSSIIBM0.UPATPL Then
            cmdSelect_SQL_9_SSIIBMPRFK_Control = " "
        Else
            cmdSelect_SQL_9_SSIIBMPRFK_Control = "N"
        End If
    End If
    
        If Trim(oldYSSIIBM0.UPSTAT) = "*DISABLED" _
       And Trim(oldYSSIIBM0.UPINMN) = "*SIGNOFF" _
       And Trim(oldYSSIIBM0.UPINPG) = "*NONE" Then '_
            cmdSelect_SQL_9_SSIIBMPRFK_Control = "X"
        End If

End Function

Public Sub fraCompteH_Display_IBM()
Dim X As String, K1 As Integer, K2 As Integer


fgCompteH.Col = 0: X = fgCompteH.Text & "|"
K1 = InStr(1, X, "|") + 1
xYSSIIBM0.SSIIBMNAT = Mid$(X, 1, K1 - 2)
K2 = InStr(K1, X, "|") + 1
xYSSIIBM0.SSIIBMUIDD = Val(Mid$(X, K1, K2 - K1 - 1))
K1 = InStr(K2, X, "|")
xYSSIIBM0.SSIIBMYVER = Val(Mid$(X, K2, K1 - K2))

Call cmdSSIIBM_Detail_Display("YSSIIBMH", "USR")
oldYSSIIBMH = xYSSIIBM0

cmdCompteH_Update.Visible = arrHab(2)
lblCompteH_SSITXTINFO.Visible = arrHab(2)
txtCompteH_SSITXTINFO.Locked = Not arrHab(2)

If oldYSSIIBMH.SSIIBMTLNK <> 0 Then
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
         & " where SSITXTNAT = '" & oldYSSIIBMH.SSIIBMNAT & "' and SSITXTUIDN = " & oldYSSIDOM0.SSIDOMUIDN _
         & " and SSITXTDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "' and SSITXTUIDX = '" & oldYSSIIBMH.UPUPRF & "'" _
         & " and SSITXTUIDD = " & oldYSSIIBMH.SSIIBMUIDD & " and SSITXTTLNK = " & oldYSSIIBMH.SSIIBMTLNK
    
    Set rsSab = cnsab.Execute(X)
    If rsSab.EOF Then
        Call MsgBox(X, vbCritical, "fgCompteH : inconnu")
    Else
        Call rsYSSITXT0_GetBuffer(rsSab, oldYSSITXT0_Histo)
        
    End If
Else
        Call rsYSSITXT0_Init(oldYSSITXT0_Histo)
        oldYSSITXT0_Histo.SSITXTNAT = oldYSSIIBMH.SSIIBMNAT
        oldYSSITXT0_Histo.SSITXTUIDN = oldYSSIDOM0.SSIDOMUIDN
        oldYSSITXT0_Histo.SSITXTDIDX = oldYSSIDOM0.SSIDOMDIDX
        oldYSSITXT0_Histo.SSITXTUIDX = oldYSSIIBMH.UPUPRF
        oldYSSITXT0_Histo.SSITXTUIDD = oldYSSIDOM0.SSIDOMUIDD

End If
txtCompteH_SSITXTINFO = Trim(oldYSSITXT0_Histo.SSITXTINFO)
End Sub

Public Sub fraCompteH_Display_SAA()
Dim X As String, K1 As Integer, K2 As Integer


fgCompteH.Col = 0: X = fgCompteH.Text & "|"
K1 = InStr(1, X, "|") + 1
usrYSSISAA0.SSISAANAT = Mid$(X, 1, K1 - 2)
K2 = InStr(K1, X, "|") + 1
usrYSSISAA0.SSISAAUIDX = Mid$(X, K1, K2 - K1 - 1)
K1 = InStr(K2, X, "|")
usrYSSISAA0.SSISAAYVER = Val(Mid$(X, K2, K1 - K2))

Call cmdSSISAA_Detail_Display("YSSISAAH")
oldYSSISAAH = usrYSSISAA0

cmdCompteH_Update.Visible = arrHab(2)
lblCompteH_SSITXTINFO.Visible = arrHab(2)
txtCompteH_SSITXTINFO.Locked = Not arrHab(2)

If oldYSSISAAH.SSISAATLNK <> 0 Then
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
         & " where SSITXTNAT = '" & oldYSSISAAH.SSISAANAT & "' and SSITXTUIDN = " & oldYSSIDOM0.SSIDOMUIDN _
         & " and SSITXTDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "' and SSITXTDIDS = 0  and SSITXTUIDx = '" & oldYSSISAAH.SSISAAUIDX & "'" _
         & " and SSITXTUIDD = " & oldYSSISAAH.SSISAAUIDD & " and SSITXTTLNK = " & oldYSSISAAH.SSISAATLNK
    
    Set rsSab = cnsab.Execute(X)
    If rsSab.EOF Then
        Call MsgBox(X, vbCritical, "fgCompteH : inconnu")
    Else
        Call rsYSSITXT0_GetBuffer(rsSab, oldYSSITXT0_Histo)
        
    End If
Else
        Call rsYSSITXT0_Init(oldYSSITXT0_Histo)
        oldYSSITXT0_Histo.SSITXTNAT = oldYSSISAAH.SSISAANAT
        oldYSSITXT0_Histo.SSITXTUIDN = oldYSSIDOM0.SSIDOMUIDN
        oldYSSITXT0_Histo.SSITXTDIDX = oldYSSIDOM0.SSIDOMDIDX
        oldYSSITXT0_Histo.SSITXTUIDX = oldYSSISAAH.SSISAAUIDX
        oldYSSITXT0_Histo.SSITXTUIDD = oldYSSISAAH.SSISAAUIDD

End If
txtCompteH_SSITXTINFO = Trim(oldYSSITXT0_Histo.SSITXTINFO)
End Sub

Public Sub fraCompteH_Display_WIN()
Dim X As String, K1 As Integer, K2 As Integer


fgCompteH.Col = 0: X = fgCompteH.Text & "|"
K1 = InStr(1, X, "|") + 1
rtfYSSIWIN0.SSIWINNAT = Mid$(X, 1, K1 - 2)
K2 = InStr(K1, X, "|") + 1
rtfYSSIWIN0.SSIWINGUID = Mid$(X, K1, K2 - K1 - 1)
K1 = InStr(K2, X, "|")
rtfYSSIWIN0.SSIWINYVER = Val(Mid$(X, K2, K1 - K2))

Call cmdSSIWIN_Detail_Display("YSSIWINH")
oldYSSIWINH = rtfYSSIWIN0

cmdCompteH_Update.Visible = arrHab(2)
lblCompteH_SSITXTINFO.Visible = arrHab(2)
txtCompteH_SSITXTINFO.Locked = Not arrHab(2)

If oldYSSIWINH.SSIWINTLNK <> 0 Then
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
         & " where SSITXTNAT = '" & oldYSSIWINH.SSIWINNAT & "' and SSITXTUIDN = " & oldYSSIDOM0.SSIDOMUIDN _
         & " and SSITXTDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "' and SSITXTDIDS = 0  and SSITXTUIDx = '" & oldYSSIWINH.SSIWINUIDX & "'" _
         & " and SSITXTUIDD = " & oldYSSIWINH.SSIWINUIDD & " and SSITXTTLNK = " & oldYSSIWINH.SSIWINTLNK
    
    Set rsSab = cnsab.Execute(X)
    If rsSab.EOF Then
        Call MsgBox(X, vbCritical, "fgCompteH : inconnu")
    Else
        Call rsYSSITXT0_GetBuffer(rsSab, oldYSSITXT0_Histo)
        
    End If
Else
        Call rsYSSITXT0_Init(oldYSSITXT0_Histo)
        oldYSSITXT0_Histo.SSITXTNAT = oldYSSIWINH.SSIWINNAT
        oldYSSITXT0_Histo.SSITXTUIDN = oldYSSIDOM0.SSIDOMUIDN
        oldYSSITXT0_Histo.SSITXTDIDX = oldYSSIDOM0.SSIDOMDIDX
        oldYSSITXT0_Histo.SSITXTUIDX = oldYSSIWINH.SSIWINUIDX
        oldYSSITXT0_Histo.SSITXTUIDD = oldYSSIWINH.SSIWINUIDD

End If
txtCompteH_SSITXTINFO = Trim(oldYSSITXT0_Histo.SSITXTINFO)
End Sub


Public Sub fraCompteH_Display_DIV()
Dim X As String, K1 As Integer, K2 As Integer


fgCompteH.Col = 0: X = fgCompteH.Text & "|"
K1 = InStr(1, X, "|") + 1
rtfYSSIDIV0.SSIDIVNAT = Mid$(X, 1, K1 - 2)
K2 = InStr(K1, X, "|") + 1
rtfYSSIDIV0.SSIDIVUIDX = Mid$(X, K1, K2 - K1 - 1)
K1 = InStr(K2, X, "|") + 1
rtfYSSIDIV0.SSIDIVUIDD = Val(Mid$(X, K2, K1 - K2 - 1))
K2 = InStr(K1, X, "|") + 1
rtfYSSIDIV0.SSIDIVYVER = Val(Mid$(X, K1, K2 - K1 - 1))

Call cmdSSIDIV_Detail_Display("YSSIDIVH")
oldYSSIDIVH = rtfYSSIDIV0

Select Case oldYSSIDIV0.SSIDIVDIDK
    Case "TEREN", "SG", "UGM": cmdProfil_Update_DIV.Visible = arrHab(5)
    Case Else: cmdProfil_Update_DIV.Visible = arrHab(2)
End Select

cmdCompteH_Update.Visible = cmdProfil_Update_DIV.Visible 'arrHab(2)
lblCompteH_SSITXTINFO.Visible = cmdProfil_Update_DIV.Visible 'arrHab(2)
txtCompteH_SSITXTINFO.Locked = Not cmdProfil_Update_DIV.Visible 'arrHab(2)

If oldYSSIDIVH.SSIDIVTLNK <> 0 Then
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
         & " where SSITXTNAT = '" & oldYSSIDIVH.SSIDIVNAT & "' and SSITXTUIDN = " & oldYSSIDOM0.SSIDOMUIDN _
         & " and SSITXTDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "' and SSITXTDIDS = 0  and SSITXTUIDx = '" & oldYSSIDIVH.SSIDIVUIDX & "'" _
         & " and SSITXTUIDD = " & oldYSSIDIVH.SSIDIVUIDD & " and SSITXTTLNK = " & oldYSSIDIVH.SSIDIVTLNK
    
    Set rsSab = cnsab.Execute(X)
    If rsSab.EOF Then
        Call MsgBox(X, vbCritical, "fgCompteH : inconnu")
    Else
        Call rsYSSITXT0_GetBuffer(rsSab, oldYSSITXT0_Histo)
        
    End If
Else
        Call rsYSSITXT0_Init(oldYSSITXT0_Histo)
        oldYSSITXT0_Histo.SSITXTNAT = oldYSSIDIVH.SSIDIVNAT
        oldYSSITXT0_Histo.SSITXTUIDN = oldYSSIDOM0.SSIDOMUIDN
        oldYSSITXT0_Histo.SSITXTDIDX = oldYSSIDOM0.SSIDOMDIDX
        oldYSSITXT0_Histo.SSITXTUIDX = oldYSSIDIVH.SSIDIVUIDX
        oldYSSITXT0_Histo.SSITXTUIDD = oldYSSIDIVH.SSIDIVUIDD

End If
txtCompteH_SSITXTINFO = Trim(oldYSSITXT0_Histo.SSITXTINFO)
End Sub


Public Sub fraCompteH_Display_MEL()
Dim X As String, K1 As Integer, K2 As Integer


fgCompteH.Col = 0: X = fgCompteH.Text & "|"
K1 = InStr(1, X, "|") + 1
rtfYSSIMEL0.SSIMELNAT = Mid$(X, 1, K1 - 2)
K2 = InStr(K1, X, "|") + 1
rtfYSSIMEL0.SSIMELUIDX = Mid$(X, K1, K2 - K1 - 1)
K1 = InStr(K2, X, "|") + 1
rtfYSSIMEL0.SSIMELUIDD = Val(Mid$(X, K2, K1 - K2 - 1))
K2 = InStr(K1, X, "|") + 1
rtfYSSIMEL0.SSIMELYVER = Val(Mid$(X, K1, K2 - K1 - 1))

Call cmdSSIMEL_Detail_Display("YSSIMELH")
oldYSSIMELH = rtfYSSIMEL0

cmdCompteH_Update.Visible = arrHab(2)
lblCompteH_SSITXTINFO.Visible = arrHab(2)
txtCompteH_SSITXTINFO.Locked = Not arrHab(2)

If oldYSSIMELH.SSIMELTLNK <> 0 Then
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
         & " where SSITXTNAT = '" & oldYSSIMELH.SSIMELNAT & "' and SSITXTUIDN = " & oldYSSIDOM0.SSIDOMUIDN _
         & " and SSITXTDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "' and SSITXTDIDS = 0  and SSITXTUIDx = '" & oldYSSIMELH.SSIMELUIDX & "'" _
         & " and SSITXTUIDD = " & oldYSSIMELH.SSIMELUIDD & " and SSITXTTLNK = " & oldYSSIMELH.SSIMELTLNK
    
    Set rsSab = cnsab.Execute(X)
    If rsSab.EOF Then
        Call MsgBox(X, vbCritical, "fgCompteH : inconnu")
    Else
        Call rsYSSITXT0_GetBuffer(rsSab, oldYSSITXT0_Histo)
        
    End If
Else
        Call rsYSSITXT0_Init(oldYSSITXT0_Histo)
        oldYSSITXT0_Histo.SSITXTNAT = oldYSSIMELH.SSIMELNAT
        oldYSSITXT0_Histo.SSITXTUIDN = oldYSSIDOM0.SSIDOMUIDN
        oldYSSITXT0_Histo.SSITXTDIDX = oldYSSIDOM0.SSIDOMDIDX
        oldYSSITXT0_Histo.SSITXTUIDX = oldYSSIMELH.SSIMELUIDX
        oldYSSITXT0_Histo.SSITXTUIDD = oldYSSIMELH.SSIMELUIDD

End If
txtCompteH_SSITXTINFO = Trim(oldYSSITXT0_Histo.SSITXTINFO)
End Sub



Public Sub fraCompteH_Display_TIC()
Dim X As String, K1 As Integer, K2 As Integer


fgCompteH.Col = 0: X = fgCompteH.Text & "|"
K1 = InStr(1, X, "|") + 1
rtfYSSITIC0.SSITICNAT = Mid$(X, 1, K1 - 2)
K2 = InStr(K1, X, "|") + 1
rtfYSSITIC0.SSITICUIDX = Mid$(X, K1, K2 - K1 - 1)
K1 = InStr(K2, X, "|") + 1
rtfYSSITIC0.SSITICUIDD = Val(Mid$(X, K2, K1 - K2 - 1))
K2 = InStr(K1, X, "|") + 1
rtfYSSITIC0.SSITICYVER = Val(Mid$(X, K1, K2 - K1 - 1))

Call cmdSSITIC_Detail_Display("YSSITICH")
oldYSSITICH = rtfYSSITIC0

cmdCompteH_Update.Visible = arrHab(2)
lblCompteH_SSITXTINFO.Visible = arrHab(2)
txtCompteH_SSITXTINFO.Locked = Not arrHab(2)

If oldYSSITICH.SSITICTLNK <> 0 Then
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
         & " where SSITXTNAT = '" & oldYSSITICH.SSITICNAT & "' and SSITXTUIDN = " & oldYSSIDOM0.SSIDOMUIDN _
         & " and SSITXTDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "' and SSITXTDIDS = 0  and SSITXTUIDx = '" & oldYSSITICH.SSITICUIDX & "'" _
         & " and SSITXTUIDD = " & oldYSSITICH.SSITICUIDD & " and SSITXTTLNK = " & oldYSSITICH.SSITICTLNK
    
    Set rsSab = cnsab.Execute(X)
    If rsSab.EOF Then
        Call MsgBox(X, vbCritical, "fgCompteH : inconnu")
    Else
        Call rsYSSITXT0_GetBuffer(rsSab, oldYSSITXT0_Histo)
        
    End If
Else
        Call rsYSSITXT0_Init(oldYSSITXT0_Histo)
        oldYSSITXT0_Histo.SSITXTNAT = oldYSSITICH.SSITICNAT
        oldYSSITXT0_Histo.SSITXTUIDN = oldYSSIDOM0.SSIDOMUIDN
        oldYSSITXT0_Histo.SSITXTDIDX = oldYSSIDOM0.SSIDOMDIDX
        oldYSSITXT0_Histo.SSITXTUIDX = oldYSSITICH.SSITICUIDX
        oldYSSITXT0_Histo.SSITXTUIDD = oldYSSITICH.SSITICUIDD

End If
txtCompteH_SSITXTINFO = Trim(oldYSSITXT0_Histo.SSITXTINFO)
End Sub





Public Sub fraCompteH_Display_SAB()
Dim X As String, K1 As Integer, K2 As Integer


fgCompteH.Col = 0: X = fgCompteH.Text & "|"
K1 = InStr(1, X, "|") + 1
usrYSSISAB0.SSISABNAT = Mid$(X, 1, K1 - 2)
K2 = InStr(K1, X, "|") + 1
usrYSSISAB0.SSISABUIDX = Mid$(X, K1, K2 - K1 - 1)
K1 = InStr(K2, X, "|")
usrYSSISAB0.SSISABYVER = Val(Mid$(X, K2, K1 - K2))



If usrYSSISAB0.SSISABNAT = "W" Then
    Call cmdSSISAW_Detail_Display("YSSISABH") '
Else
    Call cmdSSISAB_Detail_Display("YSSISABH")
End If
oldYSSISABH = usrYSSISAB0

cmdCompteH_Update.Visible = arrHab(2)
lblCompteH_SSITXTINFO.Visible = arrHab(2)
txtCompteH_SSITXTINFO.Locked = Not arrHab(2)

If oldYSSISABH.SSISABTLNK <> 0 Then
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
         & " where SSITXTNAT = '" & oldYSSISABH.SSISABNAT & "' and SSITXTUIDN = " & oldYSSIDOM0.SSIDOMUIDN _
         & " and SSITXTDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "' and SSITXTDIDS = 0  and SSITXTUIDx = '" & oldYSSISABH.SSISABUIDX & "'" _
         & " and SSITXTUIDD = " & oldYSSISABH.SSISABUIDD & " and SSITXTTLNK = " & oldYSSISABH.SSISABTLNK
    
    Set rsSab = cnsab.Execute(X)
    If rsSab.EOF Then
        Call MsgBox(X, vbCritical, "fgCompteH : inconnu")
    Else
        Call rsYSSITXT0_GetBuffer(rsSab, oldYSSITXT0_Histo)
        
    End If
Else
        Call rsYSSITXT0_Init(oldYSSITXT0_Histo)
        oldYSSITXT0_Histo.SSITXTNAT = oldYSSISABH.SSISABNAT
        oldYSSITXT0_Histo.SSITXTUIDN = oldYSSIDOM0.SSIDOMUIDN
        oldYSSITXT0_Histo.SSITXTDIDX = oldYSSIDOM0.SSIDOMDIDX
        oldYSSITXT0_Histo.SSITXTUIDX = oldYSSISABH.SSISABUIDX
        oldYSSITXT0_Histo.SSITXTUIDD = oldYSSISABH.SSISABUIDD

End If
txtCompteH_SSITXTINFO = Trim(oldYSSITXT0_Histo.SSITXTINFO)
End Sub


Public Sub cboSelect_Options_1_SSIDOMPRFX_Init()
Dim K As Integer

cboSelect_Options_1_SSIDOMPRFX.Clear
cboSelect_Options_1_SSIDOMPRFX.AddItem ""
cboSelect_Options_1_SSIDOMPRFX.AddItem "* - Tous"

Select Case cboSelect_Options_1_SSIDOMDIDX.Text
    Case "SAA"
        For K = 1 To arrSAA_Profil_Nb
            cboSelect_Options_1_SSIDOMPRFX.AddItem arrSAA_Profil_Code(K)
        Next K
    Case Else
        For K = 1 To arrSSIDOMPRFX_Nb
            If arrSSIDOMPRFX_D(K) = cboSelect_Options_1_SSIDOMDIDX.Text Then
                cboSelect_Options_1_SSIDOMPRFX.AddItem arrSSIDOMPRFX_P(K)
            End If
            
        Next K

End Select
cboSelect_Options_1_SSIDOMPRFX.ListIndex = 0

End Sub

Public Sub cboSelect_Options_4_SSIDOMNAT_Init()
Dim K As Integer

cboSelect_Options_4_SSIDOMNAT.Clear
cboSelect_Options_4_SSIDOMNAT.AddItem ""
Select Case cboSelect_Options_4_SSIDOMDIDX.Text
    Case "SAB"
        For K = 29 To 39
            cboSelect_Options_4_SSIDOMNAT.AddItem arrJRN_Origine(K)
        Next K
    Case "SAA"
        For K = 20 To 29
            cboSelect_Options_4_SSIDOMNAT.AddItem arrJRN_Origine(K)
        Next K
    Case "WIN"
        For K = 40 To 45
            cboSelect_Options_4_SSIDOMNAT.AddItem arrJRN_Origine(K)
        Next K
    Case "MEL"
        For K = 46 To 49
            cboSelect_Options_4_SSIDOMNAT.AddItem arrJRN_Origine(K)
        Next K
    Case "DIV"
        For K = 50 To 54
            cboSelect_Options_4_SSIDOMNAT.AddItem arrJRN_Origine(K)
        Next K
    Case "TIC"
        For K = 55 To 59
            cboSelect_Options_4_SSIDOMNAT.AddItem arrJRN_Origine(K)
        Next K

End Select
cboSelect_Options_4_SSIDOMNAT.ListIndex = 0

End Sub


Public Sub cmdSSIJRN_USR(lTxt As String)
Dim wYVER As Long, wORIG As String
If mYSSIUSR0_Update = "Update" Or mYSSIUSR0_Update = "Update+H" Then wYVER = newYSSIUSR0.SSIUSRYVER + 1

mYSSITXT0_JRN_Update = "New"
Call rsYSSITXT0_Init(newYSSITXT0_JRN)
newYSSITXT0_JRN.SSITXTNAT = "J"
newYSSITXT0_JRN.SSITXTUIDN = newYSSIUSR0.SSIUSRUIDN
Select Case newYSSIUSR0.SSIUSRNAT
    Case "$": wORIG = "<ORIG:3>"
    Case "S": wORIG = "<ORIG:5>"
    Case Else: wORIG = "<ORIG:1>"
End Select

newYSSITXT0_JRN.SSITXTINFO = wORIG & "<Y:USR|" & newYSSIUSR0.SSIUSRNAT & "|" & Trim(newYSSIUSR0.SSIUSRUIDN) & "|" & wYVER & "|>" _
                            & "<UID:" & newYSSIUSR0.SSIUSRUIDN & "-" & Trim(newYSSIUSR0.SSIUSRUIDX) & ">" _
                            & "<FCT:" & newYSSIUSR0.SSIUSRYFCT & ">" & lTxt

newYSSITXT0_JRN.SSITXTYAMJ = DSys
newYSSITXT0_JRN.SSITXTYHMS = time_Hms
newYSSITXT0_JRN.SSITXTYUSR = usrName_UCase

If newYSSIUSR0.SSIUSRUIDX <> oldYSSIUSR0.SSIUSRUIDX Then
    newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<UIDX:" & Trim(oldYSSIUSR0.SSIUSRUIDX) & " | " & Trim(newYSSIUSR0.SSIUSRUIDX) & ">"
End If
If newYSSIUSR0.SSIUSRSTAK <> oldYSSIUSR0.SSIUSRSTAK Then
    newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<STAK:" & Trim(oldYSSIUSR0.SSIUSRSTAK) & " | " & Trim(newYSSIUSR0.SSIUSRSTAK) & ">"
End If
If newYSSIUSR0.SSIUSRDECH <> oldYSSIUSR0.SSIUSRDECH Then
    newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<DECH:" & Trim(oldYSSIUSR0.SSIUSRDECH) & " | " & Trim(newYSSIUSR0.SSIUSRDECH) & ">"
End If
If newYSSIUSR0.SSIUSRPRFX <> oldYSSIUSR0.SSIUSRPRFX Then
    newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<PRFX:" & Trim(oldYSSIUSR0.SSIUSRPRFX) & " | " & Trim(newYSSIUSR0.SSIUSRPRFX) & ">"
End If
If newYSSIUSR0.SSIUSRPRFK <> oldYSSIUSR0.SSIUSRPRFK Then
    newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<PRFK:" & Trim(oldYSSIUSR0.SSIUSRPRFK) & " | " & Trim(newYSSIUSR0.SSIUSRPRFK) & ">"
End If

End Sub
Public Sub cmdSSIJRN_DOM(lTxt As String)
Dim wYVER As Long, wORIG As String
If mYSSIDOM0_Update = "Update" Or mYSSIDOM0_Update = "Update+H" Then wYVER = newYSSIDOM0.SSIDOMYVER + 1

mYSSITXT0_JRN_Update = "New"
Call rsYSSITXT0_Init(newYSSITXT0_JRN)
newYSSITXT0_JRN.SSITXTNAT = "J"
newYSSITXT0_JRN.SSITXTUIDN = newYSSIDOM0.SSIDOMUIDN
newYSSITXT0_JRN.SSITXTDIDX = newYSSIDOM0.SSIDOMDIDX
newYSSITXT0_JRN.SSITXTUIDX = newYSSIDOM0.SSIDOMUIDX
newYSSITXT0_JRN.SSITXTUIDD = newYSSIDOM0.SSIDOMUIDD

If newYSSIDOM0.SSIDOMNAT = "$" Then
    wORIG = "<ORIG:4>"
Else
    wORIG = "<ORIG:2>"
End If

newYSSITXT0_JRN.SSITXTINFO = wORIG & "<Y:DOM|" & newYSSIDOM0.SSIDOMNAT & "|" & newYSSIDOM0.SSIDOMUIDN & "|" & Trim(newYSSIDOM0.SSIDOMDIDX) _
                            & "|" & Trim(newYSSIDOM0.SSIDOMUIDX) & "|" & newYSSIDOM0.SSIDOMUIDD & "|" & wYVER & "|>" _
                            & "<UID:" & oldYSSIUSR0.SSIUSRUIDN & "-" & Trim(oldYSSIUSR0.SSIUSRUIDX) _
                            & " | " & Trim(newYSSIDOM0.SSIDOMDIDX) & "-" & newYSSIDOM0.SSIDOMUIDD & "-" & Trim(newYSSIDOM0.SSIDOMUIDX) _
                            & " | " & Trim(newYSSIDOM0.SSIDOMPRFX) & ">" _
                            & "<FCT:" & newYSSIDOM0.SSIDOMYFCT & ">" & lTxt

newYSSITXT0_JRN.SSITXTYAMJ = DSys
newYSSITXT0_JRN.SSITXTYHMS = time_Hms
newYSSITXT0_JRN.SSITXTYUSR = usrName_UCase

If newYSSIDOM0.SSIDOMUIDD <> oldYSSIDOM0.SSIDOMUIDD Then
    newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<UIDD:" & Trim(oldYSSIDOM0.SSIDOMUIDD) & " | " & Trim(newYSSIDOM0.SSIDOMUIDD) & ">"
End If
If newYSSIDOM0.SSIDOMUIDX <> oldYSSIDOM0.SSIDOMUIDX Then
    newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<UIDX:" & Trim(oldYSSIDOM0.SSIDOMUIDX) & " | " & Trim(newYSSIDOM0.SSIDOMUIDX) & ">"
End If
If newYSSIDOM0.SSIDOMSTAK <> oldYSSIDOM0.SSIDOMSTAK Then
    newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<STAK:" & Trim(oldYSSIDOM0.SSIDOMSTAK) & " | " & Trim(newYSSIDOM0.SSIDOMSTAK) & ">"
End If
If newYSSIDOM0.SSIDOMDECH <> oldYSSIDOM0.SSIDOMDECH Then
    newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<DECH:" & Trim(oldYSSIDOM0.SSIDOMDECH) & " | " & Trim(newYSSIDOM0.SSIDOMDECH) & ">"
End If
If newYSSIDOM0.SSIDOMPRFX <> oldYSSIDOM0.SSIDOMPRFX Then
    newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<PRFX:" & Trim(oldYSSIDOM0.SSIDOMPRFX) & " | " & Trim(newYSSIDOM0.SSIDOMPRFX) & ">"
End If
If newYSSIDOM0.SSIDOMPRFK <> oldYSSIDOM0.SSIDOMPRFK Then
    newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<PRFK:" & Trim(oldYSSIDOM0.SSIDOMPRFK) & " | " & Trim(newYSSIDOM0.SSIDOMPRFK) & ">"
End If

End Sub


Public Sub cmdSSIJRN_IBM(lTxt As String)
Dim wYVER As Long
If mYSSIIBM0_Update = "Update" Or mYSSIIBM0_Update = "Update+H" Then wYVER = newYSSIIBM0.SSIIBMYVER + 1

mYSSITXT0_JRN_Update = "New"
Call rsYSSITXT0_Init(newYSSITXT0_JRN)
newYSSITXT0_JRN.SSITXTNAT = "J"
newYSSITXT0_JRN.SSITXTUIDN = 0
newYSSITXT0_JRN.SSITXTDIDX = "IBM"
newYSSITXT0_JRN.SSITXTUIDX = newYSSIIBM0.UPUPRF
newYSSITXT0_JRN.SSITXTUIDD = newYSSIIBM0.UPUID


newYSSITXT0_JRN.SSITXTINFO = "<ORIG:10><Y:IBM|" & newYSSIIBM0.SSIIBMNAT & "|" & newYSSIIBM0.SSIIBMUIDD & "|" & wYVER & "|>" _
                           & "<UID:" & newYSSIIBM0.SSIIBMUIDD & "-" & Trim(newYSSIIBM0.UPUPRF) & ">" _
                           & "<FCT:" & newYSSIIBM0.SSIIBMYFCT & ">" & lTxt
                        
newYSSITXT0_JRN.SSITXTYAMJ = DSys
newYSSITXT0_JRN.SSITXTYHMS = time_Hms
newYSSITXT0_JRN.SSITXTYUSR = usrName_UCase


End Sub

Public Sub cmdSSIJRN_SAA(lTxt As String)
Dim xOrig As String
Dim wYVER As Long
If mYSSISAA0_Update = "Update" Or mYSSISAA0_Update = "Update+H" Then wYVER = newYSSISAA0.SSISAAYVER + 1

mYSSITXT0_JRN_Update = "New"
Call rsYSSITXT0_Init(newYSSITXT0_JRN)
newYSSITXT0_JRN.SSITXTNAT = "J"
newYSSITXT0_JRN.SSITXTUIDN = 0
newYSSITXT0_JRN.SSITXTDIDX = "SAA"
newYSSITXT0_JRN.SSITXTUIDX = newYSSISAA0.SSISAAUIDX
newYSSITXT0_JRN.SSITXTUIDD = newYSSISAA0.SSISAAUIDD
Select Case oldYSSISAA0.SSISAANAT
    Case " ": xOrig = "<ORIG:21>"
    Case "U": xOrig = "<ORIG:22>"
    Case "A": xOrig = "<ORIG:23>"
    Case "F": xOrig = "<ORIG:24>"
    Case "$": xOrig = "<ORIG:25>"
    Case "P": xOrig = "<ORIG:26>"
    Case Else: xOrig = "<ORIG:20>"
End Select

newYSSITXT0_JRN.SSITXTINFO = xOrig & "<Y:SAA|" & newYSSISAA0.SSISAANAT & "|" & Trim(newYSSISAA0.SSISAAUIDX) _
                         & "|" & newYSSISAA0.SSISAAUSEQ & "|" & wYVER & "|>" _
                         & "<UID:" & newYSSISAA0.SSISAANAT & "-" & newYSSISAA0.SSISAAUIDX & "-" & Trim(newYSSISAA0.SSISAAUNOM) & ">" _
                         & "<FCT:" & newYSSISAA0.SSISAAYFCT & ">" & lTxt

newYSSITXT0_JRN.SSITXTYAMJ = DSys
newYSSITXT0_JRN.SSITXTYHMS = time_Hms
newYSSITXT0_JRN.SSITXTYUSR = usrName_UCase
    
If newYSSITXT0_JRN.SSITXTNAT = " " Then
    If newYSSISAA0.SSISAAUIDD <> oldYSSISAA0.SSISAAUIDD Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<UIDD:" & Trim(oldYSSISAA0.SSISAAUIDD) & " | " & Trim(newYSSISAA0.SSISAAUIDD) & ">"
    End If
    If newYSSISAA0.SSISAAPRFK <> oldYSSISAA0.SSISAAPRFK Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<PRFK:" & Trim(oldYSSISAA0.SSISAAPRFK) & " | " & Trim(newYSSISAA0.SSISAAPRFK) & ">"
    End If
    If newYSSISAA0.SSISAAUNOM <> oldYSSISAA0.SSISAAUNOM Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<NOM:" & Trim(oldYSSISAA0.SSISAAUNOM) & " | " & Trim(newYSSISAA0.SSISAAUNOM) & ">"
    End If
    If newYSSISAA0.SSISAAINFO <> oldYSSISAA0.SSISAAINFO Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<INFO: voir historique" & " | " & ">"
    End If
End If

End Sub
Public Sub cmdSSIJRN_TXT_Once(lDIDX As String, lTxt As String)
Dim xSQL As String

Call cmdSSIJRN_TXT(lDIDX, lTxt)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSITXT0" _
      & " where SSITXTNAT = '" & newYSSITXT0_JRN.SSITXTNAT & "'" _
      & " and SSITXTUIDN = " & newYSSITXT0_JRN.SSITXTUIDN _
      & " and SSITXTDIDX = '" & newYSSITXT0_JRN.SSITXTDIDX & "'" _
      & " and SSITXTUIDX = '" & newYSSITXT0_JRN.SSITXTUIDX & "'" _
      & " and SSITXTUIDD = " & newYSSITXT0_JRN.SSITXTUIDD _
      & " and SSITXTYAMJ = " & newYSSITXT0_JRN.SSITXTYAMJ _
      & " and SSITXTINFO = '" & newYSSITXT0_JRN.SSITXTINFO & "'"

Set rsSab = cnsab.Execute(xSQL)
  
If Not rsSab.EOF Then mYSSITXT0_JRN_Update = ""
  

End Sub

Public Sub cmdSSIJRN_TXT(lDIDX As String, lTxt As String)

mYSSITXT0_JRN_Update = "New"
Call rsYSSITXT0_Init(newYSSITXT0_JRN)
newYSSITXT0_JRN.SSITXTNAT = "J"
newYSSITXT0_JRN.SSITXTUIDN = 0
newYSSITXT0_JRN.SSITXTDIDX = lDIDX

newYSSITXT0_JRN.SSITXTINFO = lTxt

newYSSITXT0_JRN.SSITXTYAMJ = DSys
newYSSITXT0_JRN.SSITXTYHMS = time_Hms
newYSSITXT0_JRN.SSITXTYUSR = usrName_UCase
    

End Sub

Public Sub cmdSSIJRN_WIN(lTxt As String)
Dim xOrig As String
Dim wYVER As Long

If mYSSIWIN0_Update = "Update" Or mYSSIWIN0_Update = "Update+H" Then wYVER = newYSSIWIN0.SSIWINYVER + 1

mYSSITXT0_JRN_Update = "New"
Call rsYSSITXT0_Init(newYSSITXT0_JRN)
newYSSITXT0_JRN.SSITXTNAT = "J"
newYSSITXT0_JRN.SSITXTUIDN = 0
newYSSITXT0_JRN.SSITXTDIDX = "WIN"
newYSSITXT0_JRN.SSITXTUIDX = newYSSIWIN0.SSIWINUIDX
newYSSITXT0_JRN.SSITXTUIDD = newYSSIWIN0.SSIWINUIDD
Select Case newYSSIWIN0.SSIWINNAT
    Case "", " ": xOrig = "<ORIG:41>"
    Case "$": xOrig = "<ORIG:42>"
    Case Else: xOrig = "<ORIG:40>"
End Select

newYSSITXT0_JRN.SSITXTINFO = xOrig & "<Y:WIN|" & newYSSIWIN0.SSIWINNAT & "|" & Trim(newYSSIWIN0.SSIWINUIDX) _
                         & "|" & wYVER & "|>" _
                         & "<UID:" & newYSSIWIN0.SSIWINNAT & " - " & newYSSIWIN0.SSIWINUIDX & " - " & Trim(newYSSIWIN0.SSIWINUNOM) & ">" _
                         & "<FCT:" & newYSSIWIN0.SSIWINYFCT & ">" & lTxt

newYSSITXT0_JRN.SSITXTYAMJ = DSys
newYSSITXT0_JRN.SSITXTYHMS = time_Hms
newYSSITXT0_JRN.SSITXTYUSR = usrName_UCase
    
If newYSSITXT0_JRN.SSITXTNAT = " " Then
    If newYSSIWIN0.SSIWINUIDD <> oldYSSIWIN0.SSIWINUIDD Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<UIDD:" & Trim(oldYSSIWIN0.SSIWINUIDD) & " | " & Trim(newYSSIWIN0.SSIWINUIDD) & ">"
    End If
    If newYSSIWIN0.SSIWINPRFK <> oldYSSIWIN0.SSIWINPRFK Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<PRFK:" & Trim(oldYSSIWIN0.SSIWINPRFK) & " | " & Trim(newYSSIWIN0.SSIWINPRFK) & ">"
    End If
    If newYSSIWIN0.SSIWINUNOM <> oldYSSIWIN0.SSIWINUNOM Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<NOM:" & Trim(oldYSSIWIN0.SSIWINUNOM) & " | " & Trim(newYSSIWIN0.SSIWINUNOM) & ">"
    End If
    If newYSSIWIN0.SSIWININFO <> oldYSSIWIN0.SSIWININFO Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<INFO: voir historique" & " | " & ">"
    End If
End If

End Sub
Public Sub cmdSSIJRN_MEL(lTxt As String)
Dim xOrig As String
Dim wYVER As Long

If mYSSIMEL0_Update = "Update" Or mYSSIMEL0_Update = "Update+H" Then wYVER = newYSSIMEL0.SSIMELYVER + 1

mYSSITXT0_JRN_Update = "New"
Call rsYSSITXT0_Init(newYSSITXT0_JRN)
newYSSITXT0_JRN.SSITXTNAT = "J"
newYSSITXT0_JRN.SSITXTUIDN = 0
newYSSITXT0_JRN.SSITXTDIDX = "MEL"
newYSSITXT0_JRN.SSITXTUIDX = newYSSIMEL0.SSIMELUIDX
newYSSITXT0_JRN.SSITXTUIDD = newYSSIMEL0.SSIMELUIDD
Select Case newYSSIMEL0.SSIMELNAT
    Case "$": xOrig = "<ORIG:47>"
    Case "@": xOrig = "<ORIG:48>"
    Case Else: xOrig = "<ORIG:46>"
End Select

newYSSITXT0_JRN.SSITXTINFO = xOrig & "<Y:MEL|" & newYSSIMEL0.SSIMELNAT & "|" & Trim(newYSSIMEL0.SSIMELUIDX) _
                         & "|" & wYVER & "|>" _
                         & "<UID:" & newYSSIMEL0.SSIMELNAT & " - " & newYSSIMEL0.SSIMELUIDX & " - " & newYSSIMEL0.SSIMELUIDD & " - " & Trim(newYSSIMEL0.SSIMELUNOM) & ">" _
                         & "<FCT:" & newYSSIMEL0.SSIMELYFCT & ">" & lTxt

newYSSITXT0_JRN.SSITXTYAMJ = DSys
newYSSITXT0_JRN.SSITXTYHMS = time_Hms
newYSSITXT0_JRN.SSITXTYUSR = usrName_UCase
    
If newYSSITXT0_JRN.SSITXTNAT = " " Then
     If newYSSIMEL0.SSIMELUIDX <> oldYSSIMEL0.SSIMELUIDX Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<UIDX:" & Trim(oldYSSIMEL0.SSIMELUIDX) & " | " & Trim(newYSSIMEL0.SSIMELUIDX) & ">"
    End If
   If newYSSIMEL0.SSIMELUIDD <> oldYSSIMEL0.SSIMELUIDD Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<UIDD:" & Trim(oldYSSIMEL0.SSIMELUIDD) & " | " & Trim(newYSSIMEL0.SSIMELUIDD) & ">"
    End If
    If newYSSIMEL0.SSIMELPRFK <> oldYSSIMEL0.SSIMELPRFK Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<PRFK:" & Trim(oldYSSIMEL0.SSIMELPRFK) & " | " & Trim(newYSSIMEL0.SSIMELPRFK) & ">"
    End If
    If newYSSIMEL0.SSIMELUNOM <> oldYSSIMEL0.SSIMELUNOM Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<NOM:" & Trim(oldYSSIMEL0.SSIMELUNOM) & " | " & Trim(newYSSIMEL0.SSIMELUNOM) & ">"
    End If
    If newYSSIMEL0.SSIMELINFO <> oldYSSIMEL0.SSIMELINFO Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<INFO: voir historique" & " | " & ">"
    End If
End If

End Sub

Public Sub cmdSSIJRN_TIC(lTxt As String)
Dim xOrig As String
Dim wYVER As Long

If mYSSITIC0_Update = "Update" Or mYSSITIC0_Update = "Update+H" Then wYVER = newYSSITIC0.SSITICYVER + 1

mYSSITXT0_JRN_Update = "New"
Call rsYSSITXT0_Init(newYSSITXT0_JRN)
newYSSITXT0_JRN.SSITXTNAT = "J"
newYSSITXT0_JRN.SSITXTUIDN = 0
newYSSITXT0_JRN.SSITXTDIDX = "TIC"
newYSSITXT0_JRN.SSITXTUIDX = newYSSITIC0.SSITICUIDX
newYSSITXT0_JRN.SSITXTUIDD = newYSSITIC0.SSITICUIDD
Select Case newYSSITIC0.SSITICNAT
    Case " ": xOrig = "<ORIG:56>"
    Case "$": xOrig = "<ORIG:57>"
    Case "D": xOrig = "<ORIG:58>"
    Case "R": xOrig = "<ORIG:59>"
    Case Else: xOrig = "<ORIG:55>"
End Select
newYSSITXT0_JRN.SSITXTINFO = xOrig & "<Y:TIC|" & newYSSITIC0.SSITICNAT & "|" & Trim(newYSSITIC0.SSITICUIDX) _
                         & "|" & wYVER & "|>" _
                         & "<UID:" & newYSSITIC0.SSITICNAT & " - " & newYSSITIC0.SSITICUIDX & " - " & newYSSITIC0.SSITICUIDD & " - " & Trim(newYSSITIC0.SSITICUNOM) & ">" _
                         & "<FCT:" & newYSSITIC0.SSITICYFCT & ">" & lTxt

newYSSITXT0_JRN.SSITXTYAMJ = DSys
newYSSITXT0_JRN.SSITXTYHMS = time_Hms
newYSSITXT0_JRN.SSITXTYUSR = usrName_UCase
    
If newYSSITXT0_JRN.SSITXTNAT = " " Then
     If newYSSITIC0.SSITICUIDX <> oldYSSITIC0.SSITICUIDX Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<UIDX:" & Trim(oldYSSITIC0.SSITICUIDX) & " | " & Trim(newYSSITIC0.SSITICUIDX) & ">"
    End If
   If newYSSITIC0.SSITICUIDD <> oldYSSITIC0.SSITICUIDD Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<UIDD:" & Trim(oldYSSITIC0.SSITICUIDD) & " | " & Trim(newYSSITIC0.SSITICUIDD) & ">"
    End If
    If newYSSITIC0.SSITICPRFK <> oldYSSITIC0.SSITICPRFK Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<PRFK:" & Trim(oldYSSITIC0.SSITICPRFK) & " | " & Trim(newYSSITIC0.SSITICPRFK) & ">"
    End If
    If newYSSITIC0.SSITICPRFX <> oldYSSITIC0.SSITICPRFX Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<PRFX:" & Trim(oldYSSITIC0.SSITICPRFX) & " | " & Trim(newYSSITIC0.SSITICPRFX) & ">"
    End If
    If newYSSITIC0.SSITICUNOM <> oldYSSITIC0.SSITICUNOM Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<NOM:" & Trim(oldYSSITIC0.SSITICUNOM) & " | " & Trim(newYSSITIC0.SSITICUNOM) & ">"
    End If
    If newYSSITIC0.SSITICINFO <> oldYSSITIC0.SSITICINFO Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<INFO: voir historique" & " | " & ">"
    End If
End If

End Sub


Public Sub cmdSSIJRN_DIV(lTxt As String)
Dim xOrig As String
Dim wYVER As Long

If mYSSIDIV0_Update = "Update" Or mYSSIDIV0_Update = "Update+H" Then wYVER = newYSSIDIV0.SSIDIVYVER + 1

mYSSITXT0_JRN_Update = "New"
Call rsYSSITXT0_Init(newYSSITXT0_JRN)
newYSSITXT0_JRN.SSITXTNAT = "J"
newYSSITXT0_JRN.SSITXTUIDN = 0
newYSSITXT0_JRN.SSITXTDIDX = "DIV"
newYSSITXT0_JRN.SSITXTUIDX = newYSSIDIV0.SSIDIVUIDX
newYSSITXT0_JRN.SSITXTUIDD = newYSSIDIV0.SSIDIVUIDD
Select Case newYSSIDIV0.SSIDIVNAT
    Case "", " ": xOrig = "<ORIG:51>"
    Case "$": xOrig = "<ORIG:52>"
    Case Else: xOrig = "<ORIG:50>"
End Select

newYSSITXT0_JRN.SSITXTINFO = xOrig & "<Y:DIV|" & newYSSIDIV0.SSIDIVNAT & "|" & Trim(newYSSIDIV0.SSIDIVUIDX) _
                         & "|" & wYVER & "|>" _
                         & "<UID:" & newYSSIDIV0.SSIDIVNAT & " - " & newYSSIDIV0.SSIDIVUIDX & " - " & Trim(newYSSIDIV0.SSIDIVUNOM) & ">" _
                         & "<FCT:" & newYSSIDIV0.SSIDIVYFCT & ">" & lTxt

newYSSITXT0_JRN.SSITXTYAMJ = DSys
newYSSITXT0_JRN.SSITXTYHMS = time_Hms
newYSSITXT0_JRN.SSITXTYUSR = usrName_UCase
    
'If newYSSITXT0_JRN.SSITXTNAT = " " Then
    If newYSSIDIV0.SSIDIVUIDD <> oldYSSIDIV0.SSIDIVUIDD Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<UIDD:" & Trim(oldYSSIDIV0.SSIDIVUIDD) & " | " & Trim(newYSSIDIV0.SSIDIVUIDD) & ">"
    End If
    If newYSSIDIV0.SSIDIVPRFX <> oldYSSIDIV0.SSIDIVPRFX Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<PRFX:" & Trim(oldYSSIDIV0.SSIDIVPRFX) & " | " & Trim(newYSSIDIV0.SSIDIVPRFX) & ">"
    End If
    If newYSSIDIV0.SSIDIVPRFK <> oldYSSIDIV0.SSIDIVPRFK Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<PRFK:" & Trim(oldYSSIDIV0.SSIDIVPRFK) & " | " & Trim(newYSSIDIV0.SSIDIVPRFK) & ">"
    End If
    If newYSSIDIV0.SSIDIVUNOM <> oldYSSIDIV0.SSIDIVUNOM Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<NOM:" & Trim(oldYSSIDIV0.SSIDIVUNOM) & " | " & Trim(newYSSIDIV0.SSIDIVUNOM) & ">"
    End If
    If newYSSIDIV0.SSIDIVINFO <> oldYSSIDIV0.SSIDIVINFO Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<INFO: voir historique" & " | " & ">"
    End If
'End If

End Sub







Public Sub cmdSSIJRN_SAB(lTxt As String)
Dim wYVER As Long
If mYSSISAB0_Update = "Update" Or mYSSISAB0_Update = "Update+H" Then wYVER = newYSSISAB0.SSISABYVER + 1

mYSSITXT0_JRN_Update = "New"
Call rsYSSITXT0_Init(newYSSITXT0_JRN)
newYSSITXT0_JRN.SSITXTNAT = "J"
newYSSITXT0_JRN.SSITXTUIDN = 0
newYSSITXT0_JRN.SSITXTDIDX = "SAB"
newYSSITXT0_JRN.SSITXTUIDX = newYSSISAB0.SSISABUIDX
newYSSITXT0_JRN.SSITXTUIDD = newYSSISAB0.SSISABUIDD


newYSSITXT0_JRN.SSITXTINFO = "<ORIG:30><Y:SAB|" & newYSSISAB0.SSISABNAT & "|" & Trim(newYSSISAB0.SSISABUIDX) _
                        & "|" & newYSSISAB0.SSISABULOT & "|" & wYVER & "|>" _
                        & "<UID:" & newYSSISAB0.SSISABNAT & "-" & newYSSISAB0.SSISABUIDX & "-" & newYSSISAB0.SSISABULOT _
                        & "-" & Trim(newYSSISAB0.SSISABUNOM) & ">" & "<FCT:" & newYSSISAB0.SSISABYFCT & ">" & lTxt

newYSSITXT0_JRN.SSITXTYAMJ = DSys
newYSSITXT0_JRN.SSITXTYHMS = time_Hms
newYSSITXT0_JRN.SSITXTYUSR = usrName_UCase
    
If newYSSITXT0_JRN.SSITXTNAT = " " Then
    If newYSSISAB0.SSISABUIDD <> oldYSSISAB0.SSISABUIDD Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<UIDD:" & Trim(oldYSSISAB0.SSISABUIDD) & " | " & Trim(newYSSISAB0.SSISABUIDD) & ">"
    End If
    If newYSSISAB0.SSISABPRFK <> oldYSSISAB0.SSISABPRFK Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<PRFK:" & Trim(oldYSSISAB0.SSISABPRFK) & " | " & Trim(newYSSISAB0.SSISABPRFK) & ">"
    End If
    If newYSSISAB0.SSISABUNOM <> oldYSSISAB0.SSISABUNOM Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<NOM:" & Trim(oldYSSISAB0.SSISABUNOM) & " | " & Trim(newYSSISAB0.SSISABUNOM) & ">"
    End If
    If newYSSISAB0.SSISABINFO <> oldYSSISAB0.SSISABINFO Then
        newYSSITXT0_JRN.SSITXTINFO = newYSSITXT0_JRN.SSITXTINFO & "<INFO: voir historique" & " | " & ">"
    End If
End If

End Sub

Public Sub paramSAA_Load()
Dim xSQL As String, blnREDIM As Boolean
'__________________________________________________________________________________________________
xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
     & " where SSISAANAT = 'U'"
Set rsSab = cnsab.Execute(xSQL)
arrSAA_UNIT_Nb = rsSab(0)

ReDim arrSAA_UNIT_Code(arrSAA_UNIT_Nb + 1), arrSAA_UNIT_Lib(arrSAA_UNIT_Nb + 1)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
     & " where SSISAANAT = 'U'" _
     & " order by SSISAAUIDX"

Set rsSab = cnsab.Execute(xSQL)
arrSAA_UNIT_Nb = 0
Do While Not rsSab.EOF
    arrSAA_UNIT_Nb = arrSAA_UNIT_Nb + 1
    arrSAA_UNIT_Code(arrSAA_UNIT_Nb) = Trim(rsSab("SSISAAUIDX"))
    arrSAA_UNIT_Lib(arrSAA_UNIT_Nb) = Trim(rsSab("SSISAAUNOM"))
    rsSab.MoveNext
Loop

'__________________________________________________________________________________________________

blnREDIM = False
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
     & " where SSISAANAT = 'A'" _
     & " order by SSISAAUIDD desc"

Set rsSab = cnsab.Execute(xSQL)
arrSAA_App_Nb = 0
Do While Not rsSab.EOF
    I = rsSab("SSISAAUIDD")
    If Not blnREDIM Then
        blnREDIM = True
        arrSAA_App_Nb = I
        ReDim arrSAA_App_Code(arrSAA_App_Nb + 1)
    End If
    arrSAA_App_Code(I) = Trim(rsSab("SSISAAUIDX"))
    rsSab.MoveNext
Loop
'__________________________________________________________________________________________________

blnREDIM = False
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
     & " where SSISAANAT = 'F'" _
     & " order by SSISAAUIDD desc"

Set rsSab = cnsab.Execute(xSQL)
arrSAA_Function_Nb = 0
Do While Not rsSab.EOF
    I = rsSab("SSISAAUIDD")
    If Not blnREDIM Then
        blnREDIM = True
        arrSAA_Function_Nb = I
        ReDim arrSAA_Function_Code(arrSAA_Function_Nb + 1)
    End If
    arrSAA_Function_Code(I) = Trim(rsSab("SSISAAUIDX"))
    rsSab.MoveNext
Loop
'__________________________________________________________________________________________________
xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
     & " where SSISAANAT = '$'"
Set rsSab = cnsab.Execute(xSQL)
arrSAA_Profil_Nb = rsSab(0)

ReDim arrSAA_Profil_Code(arrSAA_Profil_Nb + 1), arrSAA_Profil_Lib(arrSAA_Profil_Nb + 1)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
     & " where SSISAANAT = '$'" _
     & " order by SSISAAUIDX"

Set rsSab = cnsab.Execute(xSQL)
arrSAA_Profil_Nb = 0
Do While Not rsSab.EOF
    arrSAA_Profil_Nb = arrSAA_Profil_Nb + 1
    arrSAA_Profil_Code(arrSAA_Profil_Nb) = Trim(rsSab("SSISAAUIDX"))
    arrSAA_Profil_Lib(arrSAA_Profil_Nb) = Trim(rsSab("SSISAAUNOM"))
    rsSab.MoveNext
Loop


End Sub

Public Sub cmdSelect_SQL_9_SAA_Update()
Dim X As String
newYSSISAA0.SSISAAINFO = Trim(newYSSISAA0.SSISAAINFO)
If Len(newYSSISAA0.SSISAAINFO) > 1024 Then newYSSISAA0.SSISAAINFO = Mid$(newYSSISAA0.SSISAAINFO, 1, 104)
' Détecter les modifications
'===================================================================================
Call lstErr_AddItem(lstErr, cmdContext, newYSSISAA0.SSISAANAT & "-" & newYSSISAA0.SSISAAUIDX & "-" & newYSSISAA0.SSISAAUSEQ): DoEvents

If newYSSISAA0.SSISAANAT = oldYSSISAA0.SSISAANAT _
And newYSSISAA0.SSISAAUIDX = oldYSSISAA0.SSISAAUIDX _
And newYSSISAA0.SSISAAUSEQ = oldYSSISAA0.SSISAAUSEQ _
And newYSSISAA0.SSISAAUIDD = oldYSSISAA0.SSISAAUIDD _
And newYSSISAA0.SSISAASTAK = oldYSSISAA0.SSISAASTAK _
And newYSSISAA0.SSISAAPRFX = oldYSSISAA0.SSISAAPRFX _
And newYSSISAA0.SSISAAPRFK = oldYSSISAA0.SSISAAPRFK _
And Trim(newYSSISAA0.SSISAAUNOM) = Trim(oldYSSISAA0.SSISAAUNOM) _
And newYSSISAA0.SSISAATLNK = oldYSSISAA0.SSISAATLNK _
And Trim(newYSSISAA0.SSISAAINFO) = Trim(oldYSSISAA0.SSISAAINFO) Then
    mImport_Ok = mImport_Ok + 1
Else
    If mYSSISAA0_Update = "New" Then
        mImport_New = mImport_New + 1
    Else
        mImport_Update = mImport_Update + 1
    End If

    Call cmdSSIJRN_SAA("")
    newYSSISAA0.SSISAAYAMJ = DSys
    newYSSISAA0.SSISAAYHMS = time_Hms
    newYSSISAA0.SSISAAYUSR = usrName_UCase
    Call cmdUpdate
End If

End Sub

Public Sub cmdSelect_SQL_9_SAA_Profil_Update()
Dim I As Integer, xSQL As String

Call cmdSelect_SQL_9_SAA_Update


For I = 1 To arrSAA_Function_Nb
    If blnSAA_Function(I) Then
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
             & " where SSISAANAT = 'P' and SSISAAUIDX = '" & xYSSISAA0.SSISAAUIDX & "'" _
             & " and SSISAAUSEQ = " & arrSAA_App_K * 1000 + I
        
        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            mYSSISAA0_Update = "Update+H"
            Call rsYSSISAA0_GetBuffer(rsSab, oldYSSISAA0)
            newYSSISAA0 = oldYSSISAA0
            newYSSISAA0.SSISAAYFCT = "MOD"
            newYSSISAA0.SSISAAINFO = arrSAA_Function(I)
        Else
            mYSSISAA0_Update = "New"
            newYSSISAA0 = xYSSISAA0
            newYSSISAA0.SSISAAYFCT = "CRE"
            newYSSISAA0.SSISAAUSEQ = arrSAA_App_K * 1000 + I
            newYSSISAA0.SSISAAINFO = arrSAA_Function(I)

        End If
        newYSSISAA0.SSISAANAT = "P"
        newYSSISAA0.SSISAAUNOM = arrSAA_App_Code(arrSAA_App_K) & " - " & arrSAA_Function_Code(I)
        If Len(newYSSISAA0.SSISAAUNOM) > 32 Then newYSSISAA0.SSISAAUNOM = Mid$(newYSSISAA0.SSISAAUNOM, 1, 32)
        Call cmdSelect_SQL_9_SAA_Update

    End If

Next I

End Sub

Public Sub cmdSelect_SQL_9_SAA_Operator_YSSIDOM0()
If newYSSISAA0.SSISAAPRFK <> "?" Then
    ''''If newYSSISAA0.SSISAAPRFX <> oldYSSISAA0.SSISAAPRFX Or newYSSISAA0.SSISAASTAK <> oldYSSISAA0.SSISAASTAK Then
    
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
             & " where SSIDOMNAT = ' '" _
             & " and SSIDOMDIDX = 'SAA' and SSIDOMUIDX = '" & newYSSISAA0.SSISAAUIDX & "'" _
             & " and SSIDOMUIDD = " & newYSSISAA0.SSISAAUIDD
        
        Set rsSab = cnsab.Execute(X)
        If Not rsSab.EOF Then
            Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
            newYSSIDOM0 = oldYSSIDOM0
            newYSSIDOM0.SSIDOMPRFD = mImport_PRFD
            newYSSIDOM0.SSIDOMPRFH = mImport_PRFH
            If oldYSSIDOM0.SSIDOMPRFX = newYSSISAA0.SSISAAPRFX Then
                newYSSIDOM0.SSIDOMPRFK = " "
            Else
                newYSSIDOM0.SSIDOMPRFK = "N"
            End If
            If newYSSISAA0.SSISAASTAK <> " " Then newYSSIDOM0.SSIDOMPRFK = "X": newYSSIDOM0.SSIDOMSTAK = "N"
            
            mYSSIDOM0_Update = "Update"
            Call cmdUpdate
        Else
            Call MsgBox(X, vbCritical, "cmdSelect_SQL_9_SAA_Operator_YSSIDOM0 : inconnu")
        End If
    '''End If
End If

'''Call cmdSelect_SQL_9_SAA_Update

End Sub

Public Sub fraYSSIDOM0_Display_IBM()
Dim X As String

libSSIDOMPRFX = Trim(oldYSSIIBM0.UPTEXT)

If Trim(oldYSSIDOM0.SSIDOMUIDX) <> "" Then
    cmdCompte_All.Visible = False
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0 " _
         & " where SSIIBMNAT = ' ' and SSIIBMUIDD = " & oldYSSIDOM0.SSIDOMUIDD
    Set rsSab = cnsab.Execute(X)
    Call fgCompte_Display_IBM(vbBlue)

Else
    cmdCompte_All.Visible = True
    X = Trim(oldYSSIUSR0.SSIUSRUIDX) & "                    "
    X = Trim(Mid$(X, 1, 8))
    Call rsYSSIIBM0_Init(usrYSSIIBM0)
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIIBM0 " _
         & " where SSIIBMNAT = ' ' and UPUPRF like '%" & X & "%'" _
         & " and SSIIBMUIDD not in (select SSIDOMUIDD from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
         & " where SSIDOMDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "')"
    Set rsSab = cnsab.Execute(X)
    Call fgCompte_Display_IBM(vbMagenta)
End If

End Sub
Public Sub fraYSSIDOM0_Display_SAA()
Dim X As String

libSSIDOMPRFX = Trim(oldYSSISAA0.SSISAAUNOM)

If Trim(oldYSSIDOM0.SSIDOMUIDX) <> "" Then
    cmdCompte_All.Visible = False
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
         & " where SSISAANAT = ' ' and SSISAAUIDX = '" & oldYSSIDOM0.SSIDOMUIDX & "' and SSISAAUSEQ = 0"
    Set rsSab = cnsab.Execute(X)
    Call fgCompte_Display_SAA(vbBlue)

Else
    cmdCompte_All.Visible = True
    X = Trim(oldYSSIUSR0.SSIUSRUIDX) & "                    "
    X = Trim(Mid$(X, 1, 8))
    Call rsYSSISAA0_Init(usrYSSISAA0)
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
         & " where SSISAANAT = ' ' and SSISAAUIDX like '%" & X & "%'" _
         & " and SSISAAUIDD not in (select SSIDOMUIDD from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
         & " where SSIDOMDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "')"
    Set rsSab = cnsab.Execute(X)
    Call fgCompte_Display_SAA(vbMagenta)
End If

End Sub

Public Sub fraYSSIDOM0_Display_WIN()
Dim X As String

libSSIDOMPRFX = Trim(oldYSSIWIN0.SSIWINUNOM)

If Trim(oldYSSIDOM0.SSIDOMUIDX) <> "" Then
    cmdCompte_All.Visible = False
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
         & " where SSIWINNAT = ' ' and SSIWINUIDX = '" & oldYSSIDOM0.SSIDOMUIDX & "'"
    Set rsSab = cnsab.Execute(X)
    Call fgCompte_Display_WIN(vbBlue)

Else
    cmdCompte_All.Visible = True
    X = Trim(oldYSSIUSR0.SSIUSRUIDX) & "                    "
    X = Trim(Mid$(X, 1, 8))
    Call rsYSSIWIN0_Init(usrYSSIWIN0)
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
         & " where SSIWINNAT = ' ' and SSIWINUIDX like '%" & X & "%'" _
         & " and SSIWINUIDD not in (select SSIDOMUIDD from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
         & " where SSIDOMDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "')"
    Set rsSab = cnsab.Execute(X)
    Call fgCompte_Display_WIN(vbMagenta)
End If

End Sub

Public Sub fraYSSIDOM0_Display_DIV()
Dim X As String

libSSIDOMPRFX = Trim(oldYSSIDIV0.SSIDIVUNOM)

If Trim(oldYSSIDOM0.SSIDOMUIDX) <> "" Then
    cmdCompte_All.Visible = False
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
         & " where SSIDIVNAT = ' ' and SSIDIVUIDX = '" & oldYSSIDOM0.SSIDOMUIDX & "'" _
         & " and SSIDIVUIDD = " & oldYSSIDOM0.SSIDOMUIDD
    Set rsSab = cnsab.Execute(X)
    Call fgCompte_Display_DIV(vbBlue)

Else
    cmdCompte_All.Visible = True
    X = Trim(oldYSSIUSR0.SSIUSRUIDX) & "                    "
    X = Trim(Mid$(X, 1, 8))
   ' Call rsYSSIDIV0_Init(usrYSSIDIV0)
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
         & " where SSIDIVNAT = ' ' and SSIDIVUIDX like '%" & X & "%' and SSIDIVPRFK = '?'" _
         & " and SSIDIVUIDD = " & prfYSSIDIV0.SSIDIVUIDD
        ' & " and SSIDIVUIDX not in (select SSIDOMUIDX from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
        ' & " where SSIDOMDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "')"
    Set rsSab = cnsab.Execute(X)
    Call fgCompte_Display_DIV(vbMagenta)
End If

End Sub


Public Sub fraYSSIDOM0_Display_MEL()
Dim X As String

libSSIDOMPRFX = Trim(oldYSSIMEL0.SSIMELUNOM)

If Trim(oldYSSIDOM0.SSIDOMUIDX) <> "" Then
    cmdCompte_All.Visible = False
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0 " _
         & " where SSIMELNAT = ' ' and SSIMELUIDX = '" & oldYSSIDOM0.SSIDOMUIDX & "'" _
         & " and SSIMELUIDD = " & oldYSSIDOM0.SSIDOMUIDD
    Set rsSab = cnsab.Execute(X)
    Call fgCompte_Display_MEL(vbBlue)

Else
    cmdCompte_All.Visible = True
    X = Trim(oldYSSIUSR0.SSIUSRUIDX) & "                    "
    X = Trim(Mid$(X, 1, 8))
   ' Call rsYSSIMEL0_Init(usrYSSIMEL0)
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0 " _
         & " where SSIMELNAT = ' ' and SSIMELUIDX like '%" & X & "%' and SSIMELPRFK = '?'" _
         & " and SSIMELUIDD = " & prfYSSIMEL0.SSIMELUIDD
        ' & " and SSIMELUIDX not in (select SSIDOMUIDX from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
        ' & " where SSIDOMDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "')"
    Set rsSab = cnsab.Execute(X)
    Call fgCompte_Display_MEL(vbMagenta)
End If

End Sub


Public Sub fraYSSIDOM0_Display_TIC()

Dim X As String

libSSIDOMPRFX = Trim(oldYSSITIC0.SSITICUNOM)

If Trim(oldYSSIDOM0.SSIDOMUIDX) <> "" Then
    cmdCompte_All.Visible = False
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSITIC0 " _
         & " where SSITICNAT = ' ' and SSITICUIDX = '" & oldYSSIDOM0.SSIDOMUIDX & "'"
    Set rsSab = cnsab.Execute(X)
    Call fgCompte_Display_TIC(vbBlue)

Else
    cmdCompte_All.Visible = True
    X = Trim(oldYSSIUSR0.SSIUSRUIDX) & "                    "
    X = Trim(Mid$(X, 1, 8))
    Call rsYSSITIC0_Init(usrYSSITIC0)
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSITIC0 " _
         & " where SSITICNAT = ' ' and SSITICUIDX like '%" & X & "%'" _
         & " and SSITICUIDD not in (select SSIDOMUIDD from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
         & " where SSIDOMDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "')"
    Set rsSab = cnsab.Execute(X)
    Call fgCompte_Display_TIC(vbMagenta)
End If

End Sub


Public Sub fraYSSIDOM0_Display_SAB()
Dim X As String

libSSIDOMPRFX = Trim(oldYSSISAB0.SSISABUNOM)

If Trim(oldYSSIDOM0.SSIDOMUIDX) <> "" Then
    cmdCompte_All.Visible = False
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
         & " where SSISABNAT = ' ' and SSISABUIDX = '" & oldYSSIDOM0.SSIDOMUIDX & "' and SSISABULOT = 0"
    Set rsSab = cnsab.Execute(X)
    Call fgCompte_Display_SAB(vbBlue)

Else
    cmdCompte_All.Visible = True
    X = Trim(oldYSSIUSR0.SSIUSRUIDX) & "                    "
    X = Trim(Mid$(X, 1, 8))
    Call rsYSSISAB0_Init(usrYSSISAB0)
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
         & " where SSISABNAT = ' ' and SSISABUIDX like '%" & X & "%'" _
         & " and SSISABUIDD not in (select SSIDOMUIDD from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
         & " where SSIDOMDIDX = '" & oldYSSIDOM0.SSIDOMDIDX & "')"
    Set rsSab = cnsab.Execute(X)
    Call fgCompte_Display_SAB(vbMagenta)
End If

End Sub





Public Sub fraYSSIDOM0_Control_IBM()
oldYSSIIBM0 = usrYSSIIBM0
newYSSIIBM0 = oldYSSIIBM0
If newYSSIDOM0.SSIDOMUIDX = "" Then
    If usrYSSIIBM0.SSIIBMUIDD <> 0 Then
        newYSSIDOM0.SSIDOMUIDD = usrYSSIIBM0.SSIIBMUIDD
        newYSSIDOM0.SSIDOMUIDX = usrYSSIIBM0.UPUPRF
        
        mYSSIIBM0_Update = "Update"
        newYSSIIBM0.SSIIBMPRFK = cmdSelect_SQL_9_SSIIBMPRFK_Control(newYSSIDOM0.SSIDOMPRFX)
        newYSSIIBM0.SSIIBMYFCT = "CTL"
        newYSSIIBM0.SSIIBMYUSR = usrName_UCase
        newYSSIIBM0.SSIIBMYAMJ = DSys
        newYSSIIBM0.SSIIBMYHMS = time_Hms
        
        newYSSIDOM0.SSIDOMPRFK = newYSSIIBM0.SSIIBMPRFK
        newYSSIDOM0.SSIDOMPRFD = newYSSIIBM0.SSIIBMYAMJ
        newYSSIDOM0.SSIDOMPRFH = newYSSIIBM0.SSIIBMYHMS
    End If
Else
    If newYSSIDOM0.SSIDOMPRFK <> "X" Then
        newYSSIDOM0.SSIDOMPRFK = cmdSelect_SQL_9_SSIIBMPRFK_Control(newYSSIDOM0.SSIDOMPRFX)
    End If
End If

If newYSSIDOM0.SSIDOMSTAK = "N" And oldYSSIIBM0.UPSTAT = "*ENABLED" Then newYSSIDOM0.SSIDOMPRFK = "!"

End Sub

Public Sub fraYSSIDOM0_Control_SAA()
If newYSSIDOM0.SSIDOMUIDX = "" Then
    If usrYSSISAA0.SSISAAUIDD <> 0 Then
        newYSSIDOM0.SSIDOMUIDD = usrYSSISAA0.SSISAAUIDD
        newYSSIDOM0.SSIDOMUIDX = usrYSSISAA0.SSISAAUIDX
        
        mYSSISAA0_Update = "Update"
        oldYSSISAA0 = usrYSSISAA0
        newYSSISAA0 = oldYSSISAA0
        newYSSISAA0.SSISAAYFCT = "CTL"
        newYSSISAA0.SSISAAYUSR = usrName_UCase
        newYSSISAA0.SSISAAYAMJ = DSys
        newYSSISAA0.SSISAAYHMS = time_Hms
        
        newYSSIDOM0.SSIDOMPRFK = newYSSISAA0.SSISAAPRFK
        newYSSIDOM0.SSIDOMPRFD = newYSSISAA0.SSISAAYAMJ
        newYSSIDOM0.SSIDOMPRFH = newYSSISAA0.SSISAAYHMS
    End If
End If

If usrYSSISAA0.SSISAAPRFX = newYSSIDOM0.SSIDOMPRFX Then
    newYSSIDOM0.SSIDOMPRFK = " "
    If newYSSIDOM0.SSIDOMSTAK = "N" And oldYSSISAA0.SSISAASTAK = " " Then newYSSIDOM0.SSIDOMPRFK = "!"
Else
    newYSSIDOM0.SSIDOMPRFK = "N"
End If
If newYSSISAA0.SSISAASTAK <> " " Then newYSSIDOM0.SSIDOMPRFK = "X"

If usrYSSISAA0.SSISAAPRFK <> " " Then
    mYSSISAA0_Update = "Update"
    oldYSSISAA0 = usrYSSISAA0
    newYSSISAA0 = usrYSSISAA0
    newYSSISAA0.SSISAAPRFK = " "
End If

End Sub

Public Sub fraYSSIDOM0_Control_WIN()
If newYSSIDOM0.SSIDOMUIDX = "" Then
    If usrYSSIWIN0.SSIWINUIDD <> 0 Then
        newYSSIDOM0.SSIDOMUIDD = usrYSSIWIN0.SSIWINUIDD
        newYSSIDOM0.SSIDOMUIDX = usrYSSIWIN0.SSIWINUIDX
        
        mYSSIWIN0_Update = "Update"
        oldYSSIWIN0 = usrYSSIWIN0
        newYSSIWIN0 = oldYSSIWIN0
        newYSSIWIN0.SSIWINYFCT = "CTL"
        newYSSIWIN0.SSIWINYUSR = usrName_UCase
        newYSSIWIN0.SSIWINYAMJ = DSys
        newYSSIWIN0.SSIWINYHMS = time_Hms
        newYSSIWIN0.SSIWINPRFK = " "
        newYSSIDOM0.SSIDOMPRFK = newYSSIWIN0.SSIWINPRFK
        newYSSIDOM0.SSIDOMPRFD = newYSSIWIN0.SSIWINYAMJ
        newYSSIDOM0.SSIDOMPRFH = newYSSIWIN0.SSIWINYHMS
    End If
End If

If usrYSSIWIN0.SSIWINPRFK <> "X" Then
    If usrYSSIWIN0.SSIWINPRFX = newYSSIDOM0.SSIDOMPRFX Then
        'newYSSIDOM0.SSIDOMPRFK = " "
        If newYSSIDOM0.SSIDOMSTAK = "N" And oldYSSIWIN0.SSIWINSTAK = " " Then newYSSIDOM0.SSIDOMPRFK = "!"
    Else
        newYSSIDOM0.SSIDOMPRFK = "N"
    End If
End If
If cmdSSIWIN_UAC_PRTFK(usrYSSIWIN0.SSIWININFO) = "X" Then newYSSIDOM0.SSIDOMPRFK = "X"

'!!!!!!!!!!!!!!! A revoir
'__________________________
'If newYSSIWIN0.SSIWINSTAK <> " " Then newYSSIDOM0.SSIDOMPRFK = "X"

'If usrYSSIWIN0.SSIWINPRFK <> " " Then
'    mYSSIWIN0_Update = "Update"
'    oldYSSIWIN0 = usrYSSIWIN0
'    newYSSIWIN0 = usrYSSIWIN0
'    newYSSIWIN0.SSIWINPRFK = " "
'End If

End Sub


Public Sub fraYSSIDOM0_Control_DIV()
If newYSSIDOM0.SSIDOMUIDX = "" Then
    If prfYSSIDIV0.SSIDIVUIDD <> 0 Then
        newYSSIDOM0.SSIDOMUIDD = prfYSSIDIV0.SSIDIVUIDD '!!!!!!!!!!!!!!!!!!!!!!!
        newYSSIDOM0.SSIDOMUIDX = usrYSSIDIV0.SSIDIVUIDX
        mYSSIDIV0_Update = "Update"
        'oldYSSIDIV0 = usrYSSIDIV0
        newYSSIDIV0 = oldYSSIDIV0
        newYSSIDIV0.SSIDIVYFCT = "CTL"
        newYSSIDIV0.SSIDIVYUSR = usrName_UCase
        newYSSIDIV0.SSIDIVYAMJ = DSys
        newYSSIDIV0.SSIDIVYHMS = time_Hms
        If newYSSIDIV0.SSIDIVPRFK = "?" Then newYSSIDIV0.SSIDIVPRFK = " "

        newYSSIDOM0.SSIDOMPRFK = newYSSIDIV0.SSIDIVPRFK
        newYSSIDOM0.SSIDOMPRFD = newYSSIDIV0.SSIDIVYAMJ
        newYSSIDOM0.SSIDOMPRFH = newYSSIDIV0.SSIDIVYHMS
    End If
End If

If usrYSSIDIV0.SSIDIVPRFX = newYSSIDOM0.SSIDOMPRFX Then
    newYSSIDOM0.SSIDOMPRFK = " "
    If newYSSIDOM0.SSIDOMSTAK = "N" And oldYSSIDIV0.SSIDIVSTAK = " " Then newYSSIDOM0.SSIDOMPRFK = "!"
Else
    newYSSIDOM0.SSIDOMPRFK = "N"
End If


If usrYSSIDIV0.SSIDIVPRFK <> " " Then
    mYSSIDIV0_Update = "Update"
    oldYSSIDIV0 = usrYSSIDIV0
    newYSSIDIV0 = usrYSSIDIV0
    newYSSIDIV0.SSIDIVPRFK = " "
End If

End Sub

Public Sub fraYSSIDOM0_Control_TIC()
If newYSSIDOM0.SSIDOMUIDX = "" Then
    If usrYSSITIC0.SSITICUIDD <> 0 Then
        newYSSIDOM0.SSIDOMUIDD = usrYSSITIC0.SSITICUIDD '!!!!!!!!!!!!!!!!!!!!!!!
        newYSSIDOM0.SSIDOMUIDX = usrYSSITIC0.SSITICUIDX
        mYSSITIC0_Update = "Update"
        'oldYSSITIC0 = usrYSSITIC0
        newYSSITIC0 = oldYSSITIC0
        newYSSITIC0.SSITICYFCT = "CTL"
        newYSSITIC0.SSITICYUSR = usrName_UCase
        newYSSITIC0.SSITICYAMJ = DSys
        newYSSITIC0.SSITICYHMS = time_Hms
        If newYSSITIC0.SSITICPRFK = "?" Then newYSSITIC0.SSITICPRFK = " "

        newYSSIDOM0.SSIDOMPRFK = newYSSITIC0.SSITICPRFK
        newYSSIDOM0.SSIDOMPRFD = newYSSITIC0.SSITICYAMJ
        newYSSIDOM0.SSIDOMPRFH = newYSSITIC0.SSITICYHMS
    End If
End If

If usrYSSITIC0.SSITICPRFX = newYSSIDOM0.SSIDOMPRFX Then
    newYSSIDOM0.SSIDOMPRFK = " "
    If newYSSIDOM0.SSIDOMSTAK = "N" And oldYSSITIC0.SSITICSTAK = " " Then newYSSIDOM0.SSIDOMPRFK = "!"
Else
    newYSSIDOM0.SSIDOMPRFK = "N"
End If


If usrYSSITIC0.SSITICPRFK <> " " Then
    mYSSITIC0_Update = "Update"
    oldYSSITIC0 = usrYSSITIC0
    newYSSITIC0 = usrYSSITIC0
    newYSSITIC0.SSITICPRFK = " "
End If

End Sub


Public Sub fraYSSIDOM0_Control_SAB()
If newYSSIDOM0.SSIDOMUIDX = "" Then
    If usrYSSISAB0.SSISABUIDD <> 0 Then
        newYSSIDOM0.SSIDOMUIDD = usrYSSISAB0.SSISABUIDD
        newYSSIDOM0.SSIDOMUIDX = usrYSSISAB0.SSISABUIDX
        
        mYSSISAB0_Update = "Update"
        oldYSSISAB0 = usrYSSISAB0
        newYSSISAB0 = oldYSSISAB0
        newYSSISAB0.SSISABYFCT = "CTL"
        newYSSISAB0.SSISABYUSR = usrName_UCase
        newYSSISAB0.SSISABYAMJ = DSys
        newYSSISAB0.SSISABYHMS = time_Hms
        
        newYSSIDOM0.SSIDOMPRFK = newYSSISAB0.SSISABPRFK
        newYSSIDOM0.SSIDOMPRFD = newYSSISAB0.SSISABYAMJ
        newYSSIDOM0.SSIDOMPRFH = newYSSISAB0.SSISABYHMS
    End If
End If

If usrYSSISAB0.SSISABPRFX = newYSSIDOM0.SSIDOMPRFX Then
    newYSSIDOM0.SSIDOMPRFK = " "
    If newYSSIDOM0.SSIDOMSTAK = "N" And oldYSSISAB0.SSISABSTAK = " " Then newYSSIDOM0.SSIDOMPRFK = "!"
Else
    newYSSIDOM0.SSIDOMPRFK = "N"
End If
If Trim(newYSSISAB0.SSISABSTAK) <> "" Then newYSSIDOM0.SSIDOMPRFK = "X"

If Trim(usrYSSISAB0.SSISABPRFX) = "P_MIN" Then
    newYSSIDOM0.SSIDOMPRFK = "X"
    newYSSIDOM0.SSIDOMSTAK = "N"
End If

If Trim(usrYSSISAB0.SSISABPRFK) <> "" Then
    mYSSISAB0_Update = "Update"
    oldYSSISAB0 = usrYSSISAB0
    newYSSISAB0 = usrYSSISAB0
    newYSSISAB0.SSISABPRFK = " "
End If

End Sub

Public Function paramFile(lDIDX As String, lTLNK As Long) As String
Dim xSQL As String
Call rsYSSITXT0_Init(xYSSITXT0)

xYSSITXT0.SSITXTNAT = "P"
xYSSITXT0.SSITXTDIDX = lDIDX
xYSSITXT0.SSITXTTLNK = lTLNK
xSQL = "select SSITXTINFO from " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
     & "  Where SSITXTNAT = 'P'" _
     & "  and SSITXTUIDN = 0" _
     & "  and SSITXTDIDX = '" & lDIDX & "'" _
     & "  and SSITXTUIDX = ''" _
     & "  and SSITXTUIDd = 0" _
     & "  and SSITXTTLNK = " & lTLNK
Set rsSab_X = cnsab.Execute(xSQL)
    
If Not rsSab_X.EOF Then
    paramFile = Trim(rsSab_X("SSITXTINFO"))
Else
    Call MsgBox("Paramétrage absent pour : " & lDIDX & " " & lTLNK, vbCritical, "ParamFile")
    paramFile = "C:\temp"
End If

End Function

Public Sub lstParam_K1_Load(lDIDX As String)
Dim xSQL As String

fraParam_K2.Visible = False

Call rsYSSITXT0_Init(paramYSSITXT0)
paramYSSITXT0.SSITXTNAT = "P"
paramYSSITXT0.SSITXTDIDX = lDIDX

lstParam_K2.Clear
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
     & "  Where SSITXTNAT = 'P'" _
     & "  and SSITXTUIDN = 0" _
     & "  and SSITXTDIDX = '" & lDIDX & "'" _
     & "  and SSITXTUIDX = ''" _
     & "  and SSITXTUIDd = 0" _
     & "  order by SSITXTTLNK  "
Set rsSab_X = cnsab.Execute(xSQL)
    
Do While Not rsSab_X.EOF
    lstParam_K2.AddItem rsSab_X("SSITXTTLNK") & " - " & Trim(rsSab_X("SSITXTINFO"))
    rsSab_X.MoveNext
Loop
End Sub


Public Function lstParam_K2_Load()
Dim xSQL As String, K As Long
On Error GoTo Error_Handler

lstParam_K2_Load = "?"

K = InStr(lstParam_K2, "-")
If K < 1 Then K = 1
K = Val(Mid$(lstParam_K2, 1, K - 1))

txtParam_K2_Code = ""
txtParam_K2_Info = ""

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSITXT0 " _
     & "  Where SSITXTNAT = 'P'" _
     & "  and SSITXTUIDN = 0" _
     & "  and SSITXTDIDX = '" & paramYSSITXT0.SSITXTDIDX & "'" _
     & "  and SSITXTUIDX = ''" _
     & "  and SSITXTUIDd = 0" _
     & "  and SSITXTTLNK = " & K
Set rsSab_X = cnsab.Execute(xSQL)
    
If Not rsSab_X.EOF Then
    Call rsYSSITXT0_GetBuffer(rsSab_X, paramYSSITXT0)
    txtParam_K2_Code = paramYSSITXT0.SSITXTTLNK
    txtParam_K2_Info = paramYSSITXT0.SSITXTINFO
    lstParam_K2_Load = Null
Else
    V = "Enregistrement inconnu dans YSSITXT0"
    GoTo Error_MsgBox
End If

Exit Function

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & "lstParam_K2_Load"
End Function

Public Sub cmdSelect_SQL_9_SAA()
Dim blnSAA_Load As Boolean, X As String, objFile As File
On Error GoTo Error_Handler

mBIA_SSI_Archives = paramFile("BIA", 2)
If mBIA_SSI_Archives = "C:\temp" Then mBIA_SSI_Archives = "C:\temp\BIA_SSI_Archives"
If Not msFileSystem.FolderExists(mBIA_SSI_Archives) Then MkDir mBIA_SSI_Archives

mFile = paramFile("SAA", 1)
If Dir(mFile) = "" Then
    If Not blnAuto Then Call MsgBox(mFile, vbInformation, "Fichier non trouvé")
Else
    Call cmdUpdate_Init
    Call cmdSSIJRN_TXT("SAA", "<ORIG:20><FCT:9-SAA><UID:" & mFile & "><X:importation SAA Unit >")
    Call cmdUpdate
    Set objFile = msFileSystem.GetFile(mFile)
    X = objFile.DateLastModified
    mImport_PRFD = Val(Mid$(X, 7, 4) & Mid$(X, 4, 2) & Mid$(X, 1, 2))
    mImport_PRFH = Val(Mid$(X, 12, 2) & Mid$(X, 15, 2) & Mid$(X, 18, 2))
    mImport_Nb = 0: mImport_In = 0: mImport_New = 0: mImport_Update = 0: mImport_Ok = 0
'================================
    cmdSelect_SQL_9_SAA_Unit
'================================
    Call cmdUpdate_Init
    X = "à Importer : " & mImport_Nb & "  Lus : " & mImport_In & "  Ok : " & mImport_Ok & "  Créés : " & mImport_New & "  Modifiés : " & mImport_Update

    Call cmdSSIJRN_TXT("SAA", "<ORIG:20><FCT:9-SAA><UID:" & mFile & "><X:" & X & ">")
    Call cmdUpdate
    Call lstErr_AddItem(lstErr, cmdContext, "SAA_Unit : " & X): DoEvents
    If mImport_Nb <> mImport_In Then Call MsgBox(X, vbCritical, "Erreur Import " & mFile)
    
    blnSAA_Load = True
    msFileSystem.MoveFile mFile, mBIA_SSI_Archives & "\SAA Unit " & DSys & "_" & time_Hms & ".txt"
    'If blnAuto Then Kill mFile
End If

mFile = paramFile("SAA", 2)
If Dir(mFile) = "" Then
    If Not blnAuto Then Call MsgBox(mFile, vbInformation, "Fichier non trouvé")
Else
    Call cmdUpdate_Init
    Call cmdSSIJRN_TXT("SAA", "<ORIG:20><FCT:9-SAA><UID:" & mFile & "><X:importation SAA Profile >")
    Call cmdUpdate
    Set objFile = msFileSystem.GetFile(mFile)
    X = objFile.DateLastModified
    mImport_PRFD = Val(Mid$(X, 7, 4) & Mid$(X, 4, 2) & Mid$(X, 1, 2))
    mImport_PRFH = Val(Mid$(X, 12, 2) & Mid$(X, 15, 2) & Mid$(X, 18, 2))
    mImport_Nb = 0: mImport_In = 0: mImport_New = 0: mImport_Update = 0: mImport_Ok = 0
'================================
    
    cmdSelect_SQL_9_SAA_App
    cmdSelect_SQL_9_SAA_Function
    
    mImport_Nb = 0: mImport_In = 0: mImport_New = 0: mImport_Update = 0: mImport_Ok = 0
    cmdSelect_SQL_9_SAA_Profil
'================================
    Call cmdUpdate_Init
    X = "à Importer : " & mImport_Nb & "  Lus : " & mImport_In & "  Ok : " & mImport_Ok & "  Créés : " & mImport_New & "  Modifiés : " & mImport_Update

    Call cmdSSIJRN_TXT("SAA", "<ORIG:20><FCT:9-SAA><UID:" & mFile & "><X:" & X & ">")
    Call cmdUpdate
    Call lstErr_AddItem(lstErr, cmdContext, "SAA_Unit : " & X): DoEvents
    If mImport_Nb <> mImport_In Then Call MsgBox(X, vbCritical, "Erreur Import " & mFile)
    
    blnSAA_Load = True
    msFileSystem.MoveFile mFile, mBIA_SSI_Archives & "\SAA Profile " & DSys & "_" & time_Hms & ".txt"
    'If blnAuto Then Kill mFile
End If

mFile = paramFile("SAA", 3)
If Dir(mFile) = "" Then
    If Not blnAuto Then Call MsgBox(mFile, vbInformation, "Fichier non trouvé")
Else
    Call cmdUpdate_Init
    Call cmdSSIJRN_TXT("SAA", "<ORIG:20><FCT:9-SAA><UID:" & mFile & "><X:importation SAA Operator >")
    Call cmdUpdate
    Set objFile = msFileSystem.GetFile(mFile)
    X = objFile.DateLastModified
    mImport_PRFD = Val(Mid$(X, 7, 4) & Mid$(X, 4, 2) & Mid$(X, 1, 2))
    mImport_PRFH = Val(Mid$(X, 12, 2) & Mid$(X, 15, 2) & Mid$(X, 18, 2))
    mImport_Nb = 0: mImport_In = 0: mImport_New = 0: mImport_Update = 0: mImport_Ok = 0
'================================
    cmdSelect_SQL_9_SAA_Operator
'================================
    Call cmdUpdate_Init
    X = "à Importer : " & mImport_Nb & "  Lus : " & mImport_In & "  Ok : " & mImport_Ok & "  Créés : " & mImport_New & "  Modifiés : " & mImport_Update

    Call cmdSSIJRN_TXT("SAA", "<ORIG:20><FCT:9-SAA><UID:" & mFile & "><X:" & X & ">")
    Call cmdUpdate
    Call lstErr_AddItem(lstErr, cmdContext, "SAA_Unit : " & X): DoEvents
    If mImport_Nb <> mImport_In Then Call MsgBox(X, vbCritical, "Erreur Import " & mFile)

    blnSAA_Load = True
    msFileSystem.MoveFile mFile, mBIA_SSI_Archives & "\SAA Operator " & DSys & "_" & time_Hms & ".txt"
    'If blnAuto Then Kill mFile
'End If

    cmdSelect_SQL_9_SAA_Profil_Inactif
End If


If blnSAA_Load Then Call paramSAA_Load
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdSelect_SQL_9_SAA"
Exit_sub:

Set rsSab = Nothing
Call paramSAA_Load

Exit Sub

End Sub

Public Sub cmdSelect_SQL_9_TIC()
Dim blnATHIC_Load As Boolean, X As String, objFile As File
On Error GoTo Error_Handler

mBIA_SSI_Archives = paramFile("BIA", 2)
If mBIA_SSI_Archives = "C:\temp" Then mBIA_SSI_Archives = "C:\temp\BIA_SSI_Archives"
If Not msFileSystem.FolderExists(mBIA_SSI_Archives) Then MkDir mBIA_SSI_Archives

'mFile = paramFile("ATHIC", 1)
mFile = "c:\Temp\athic.txt"
If Dir(mFile) = "" Then
    
    If Not blnAuto Then Call MsgBox(mFile, vbInformation, "Fichier non trouvé")
Else
    Call cmdUpdate_Init
    Call cmdSSIJRN_TXT("ATHIC", "<ORIG:55><FCT:9-TIC><UID:" & mFile & "><X:importation ATHIC Unit >")
    Call cmdUpdate
    Set objFile = msFileSystem.GetFile(mFile)
    X = objFile.DateLastModified
    mImport_PRFD = Val(Mid$(X, 7, 4) & Mid$(X, 4, 2) & Mid$(X, 1, 2))
    mImport_PRFH = Val(Mid$(X, 12, 2) & Mid$(X, 15, 2) & Mid$(X, 18, 2))
    mImport_Nb = 0: mImport_In = 0: mImport_New = 0: mImport_Update = 0: mImport_Ok = 0: mImport_Ann = 0
    X = MsgBox("C:\Temp\ATHIC.txt  du " & objFile.DateLastModified & vbCrLf & vbCrLf _
            & "archivé après traitement : " & mBIA_SSI_Archives, vbYesNo, "Confirmation de l'importation")
    If X = vbNo Then GoTo Exit_sub
'================================
    cmdSelect_SQL_9_TIC_USR
'================================

'    X = "select count(*)  from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
'         & " where SSITICNAT = ' ' and SSITICPRFK = 'X'"
'    Set rsSab = cnsab.Execute(X)
'    If rsSab.EOF Then
'        mImport_Ann = 0
'    Else
'        mImport_Ann = rsSab(0)
'    End If

    Call cmdUpdate_Init
    X = "Comptes : " & mImport_Nb & "  Idem : " & mImport_Ok & "  Créés : " & mImport_New & "  Modifiés : " & mImport_Update & "  Annulés : " & mImport_Ann

    Call cmdSSIJRN_TXT("ATHIC", "<ORIG:55><FCT:9-TIC><UID:" & mFile & "><X:" & X & ">")
    Call cmdUpdate
    Call lstErr_AddItem(lstErr, cmdContext, "ATHIC : " & X): DoEvents
    Call MsgBox(X, vbInformation, "Athic importation " & mFile)
    
    blnATHIC_Load = True
    msFileSystem.MoveFile mFile, mBIA_SSI_Archives & "\ATHIC " & DSys & "_" & time_Hms & ".txt"
    If blnAuto Then Kill mFile
End If

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdSelect_SQL_9_ATHIC"
Exit_sub:

Set rsSab = Nothing

Exit Sub

End Sub


Public Sub cmdSelect_SQL_9_DIV_UGM()
Dim blnSAA_Load As Boolean, objFile As File


If Not arrHab(5) Then
    Call MsgBox("vous n'êtes pas habilité à cette fonction", vbExclamation, "Sécurité physique")
    GoTo Exit_sub
End If

mBIA_SSI_Archives = paramFile("BIA", 2)
If mBIA_SSI_Archives = "C:\temp" Then mBIA_SSI_Archives = "C:\temp\BIA_SSI_Archives"
If Not msFileSystem.FolderExists(mBIA_SSI_Archives) Then MkDir mBIA_SSI_Archives

mFile = paramFile("UGM", 1)
If Dir(mFile) = "" Then
    If Not blnAuto Then Call MsgBox(mFile, vbInformation, "Fichier non trouvé")
Else
    Set objFile = msFileSystem.GetFile(mFile)
    X = objFile.DateLastModified
    mImport_PRFD = Val(Mid$(X, 7, 4) & Mid$(X, 4, 2) & Mid$(X, 1, 2))
    mImport_PRFH = Val(Mid$(X, 12, 2) & Mid$(X, 15, 2) & Mid$(X, 18, 2))
    
    Call cmdUpdate_Init
    Call cmdSSIJRN_TXT("DIV", "<ORIG:50><FCT:9-UGM><UID:" & mFile & "><X:importation DIV_UGM >")
    Call cmdUpdate
    
    mImport_Nb = 0: mImport_In = 0: mImport_New = 0: mImport_Update = 0: mImport_Ok = 0
'================================
    cmdSelect_SQL_9_DIV_UGM_Load
'================================
    Call cmdUpdate_Init
    X = "Lus : " & mImport_In & "  Ok : " & mImport_Ok & "  Créés : " & mImport_New & "  Modifiés : " & mImport_Update

    Call cmdSSIJRN_TXT("DIV", "<ORIG:50><FCT:9-UGM><UID:" & mFile & "><X:" & X & ">")
    Call cmdUpdate
    Call lstErr_AddItem(lstErr, cmdContext, "DIV_UGM : " & X): DoEvents
    msFileSystem.MoveFile mFile, mBIA_SSI_Archives & "\UGM " & DSys & "_" & time_Hms & ".txt"
    'If blnAuto Then Kill mFile
End If

Exit_sub:

End Sub

Public Sub cmdSelect_SQL_9_DIV_TEREN()
Dim blnSAA_Load As Boolean, objFile As File

If Not arrHab(5) Then
    Call MsgBox("vous n'êtes pas habilité à cette fonction", vbExclamation, "Sécurité physique")
    GoTo Exit_sub
End If

mBIA_SSI_Archives = paramFile("BIA", 2)
If mBIA_SSI_Archives = "C:\temp" Then mBIA_SSI_Archives = "C:\temp\BIA_SSI_Archives"
If Not msFileSystem.FolderExists(mBIA_SSI_Archives) Then MkDir mBIA_SSI_Archives

mFile = paramFile("TEREN", 1)
mFile = InputBox("Fichier TERENA à importer :" _
    & vbCrLf & "     =========================" & vbCrLf & "Archivage : " & mBIA_SSI_Archives _
    & vbCrLf & "     =========================", " ", mFile)
If Trim(mFile) = "" Then GoTo Exit_sub

If Dir(mFile) = "" Then
    If Not blnAuto Then Call MsgBox(mFile, vbInformation, "Fichier non trouvé")
Else
    Set objFile = msFileSystem.GetFile(mFile)
    X = objFile.DateLastModified
    mImport_PRFD = Val(Mid$(X, 7, 4) & Mid$(X, 4, 2) & Mid$(X, 1, 2))
    mImport_PRFH = Val(Mid$(X, 12, 2) & Mid$(X, 15, 2) & Mid$(X, 18, 2))
    
    Call cmdUpdate_Init
    Call cmdSSIJRN_TXT("DIV", "<ORIG:50><FCT:9-TEREN><UID:" & mFile & "><X:importation DIV_TEREN >")
    Call cmdUpdate
    
    mImport_Nb = 0: mImport_In = 0: mImport_New = 0: mImport_Update = 0: mImport_Ok = 0
'================================
    cmdSelect_SQL_9_DIV_TEREN_Load
'================================
    Call cmdUpdate_Init
    X = "Lus : " & mImport_In & "  Ok : " & mImport_Ok & "  Créés : " & mImport_New & "  Modifiés : " & mImport_Update

    Call cmdSSIJRN_TXT("DIV", "<ORIG:50><FCT:9-TEREN><UID:" & mFile & "><X:" & X & ">")
    Call cmdUpdate
    Call lstErr_AddItem(lstErr, cmdContext, "DIV_TEREN : " & X): DoEvents
    msFileSystem.MoveFile mFile, mBIA_SSI_Archives & "\TEREN " & DSys & "_" & time_Hms & ".txt"
    'If blnAuto Then Kill mFile
End If

Exit_sub:

End Sub

Public Sub fgSelect_Display_3_Total(lColor As Long)
Dim K As Integer, wColor As Long, wRowHeight As Integer
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
fgSelect.RowHeight(fgSelect.Row) = 350
fgSelect.Col = 5:
fgSelect.CellFontBold = True
If arrCtl_Nb(arrCtl_K) = 0 Then
    fgSelect.Text = arrCtl_Lib(arrCtl_K) & " : Néant"
    'fgSelect.CellForeColor = mColor_Z0
    fgSelect.CellBackColor = mColor_G0 'fgSelect.BackColorFixed
    wColor = fgSelect.BackColorFixed
    wRowHeight = 10
Else
    fgSelect.Text = arrCtl_Nb(arrCtl_K) & " " & arrCtl_Lib(arrCtl_K)
    fgSelect.CellBackColor = lColor 'mColor_W1
    wColor = vbMagenta
    wRowHeight = 50
End If

'fgSelect.Col = 5: fgSelect.Text = arrCtl_Lib(arrCtl_K): fgSelect.CellForeColor = mColor_Z0: fgSelect.CellFontBold = True
fgSelect.Col = 9: fgSelect.Text = "X-"
'For K = 0 To 10
 '   fgSelect.Col = K: fgSelect.CellBackColor = fgSelect.BackColorFixed 'RGB(230, 230, 230)
    
'Next K
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
fgSelect.RowHeight(fgSelect.Row) = wRowHeight
For K = 0 To 9
    fgSelect.Col = K
    fgSelect.CellBackColor = wColor
Next K
End Sub

Public Sub fraDetail_Control_SSIUSRPRFK()
If Trim(oldYSSIUSR0.SSIUSRPRFX) <> "" Then

    newYSSIUSR0 = oldYSSIUSR0
    Call fraDetail_Control_SSIUSRPRFX
    mYSSIDOM0_Update = ""
    If newYSSIUSR0.SSIUSRPRFK <> oldYSSIUSR0.SSIUSRPRFK Then
        mYSSIUSR0_Update = "Update"
        Call cmdUpdate
    End If
End If
End Sub

Public Sub fgSelect_Row_Click()
Dim K1 As Integer, K2 As Integer, wX As String, X As String
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        txtRTF.Visible = False
        Select Case cmdSelect_SQL_K
            Case "1", "2", "3", "3_H", "2_S"
                fgSelect.Col = 9
                K1 = InStr(1, fgSelect.Text, "|") + 1
                wX = Mid$(fgSelect.Text, 1, K1 - 2)
                Select Case Trim(wX)
                    Case "", "$", "S"
                        oldYSSIUSR0.SSIUSRNAT = wX
                        K2 = InStr(K1, fgSelect.Text, "|") + 1
                        oldYSSIUSR0.SSIUSRUIDN = Val(Mid$(fgSelect.Text, K1, K2 - K1 - 1))
                        
                        oldYSSIDOM0.SSIDOMNAT = oldYSSIUSR0.SSIUSRNAT
                        oldYSSIDOM0.SSIDOMUIDN = oldYSSIUSR0.SSIUSRUIDN
                        K1 = InStr(K2, fgSelect.Text, "|") + 1
                        oldYSSIDOM0.SSIDOMDIDX = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                        K2 = InStr(K1, fgSelect.Text, "|") + 1
                        oldYSSIDOM0.SSIDOMUIDX = Trim(Mid$(fgSelect.Text, K1, K2 - K1 - 1))
                        K1 = InStr(K2, fgSelect.Text, "|") + 1
                        oldYSSIDOM0.SSIDOMUIDD = Val(Mid$(fgSelect.Text, K2, K1 - K2 - 1))
                        
                        Call fraDetail_Load
                        If oldYSSIDOM0.SSIDOMUIDX <> "" Then
                            Call fraYSSIDOM0_Load
                            If cmdSelect_SQL_K = "3_H" Then Call cmdProfil_Histo_Click
                        End If
                    Case "?"
                        K2 = InStr(K1, fgSelect.Text, "|") + 1
                        K1 = InStr(K2, fgSelect.Text, "|") + 1
                        wX = Trim(Mid$(fgSelect.Text, K2, K1 - K2 - 1))
                        Select Case wX
                            Case "IBM", "IBM_S"
                                oldYSSIIBM0.SSIIBMNAT = ""
                                K2 = InStr(K1, fgSelect.Text, "|") + 1
                                K1 = InStr(K2, fgSelect.Text, "|") + 1
                                oldYSSIIBM0.SSIIBMUIDD = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                                Call cmdSSIIBM_Detail_Display("", "")
                                 If arrHab(2) Then
                                    
                                    If wX = "IBM_S" Then
                                        X = MsgBox("Vu ?", vbYesNo, "Compte SUPPRIME par l'administrateur IBM")
                                    Else
                                        X = MsgBox("Voulez-vous le supprimer ?", vbYesNo, "Compte IBM orphelin")
                                    End If
                                    
                                    If X = vbYes Then
                                        cmdUpdate_Init
                                        mYSSIIBM0_Update = "Update"
                                        oldYSSIIBM0 = xYSSIIBM0
                                        newYSSIIBM0 = oldYSSIIBM0
                                        newYSSIIBM0.SSIIBMPRFK = "S"
                                        newYSSIIBM0.SSIIBMYFCT = "VU "
                                        newYSSIIBM0.SSIIBMYUSR = usrName_UCase
                                        newYSSIIBM0.SSIIBMYAMJ = DSys
                                        newYSSIIBM0.SSIIBMYHMS = time_Hms
                                        Call cmdSSIJRN_IBM("<X:VU compte IBM supprimé>")
                                        cmdUpdate
                                    End If
                                    txtRTF.Visible = False
                                End If
                           Case "IBM_H"
                                oldYSSIIBM0.SSIIBMNAT = ""
                                K2 = InStr(K1, fgSelect.Text, "|") + 1
                                K1 = InStr(K2, fgSelect.Text, "|") + 1
                                '  '$JPL 2015-05-12----------------------------------------------------------------------
                                'oldYSSIIBM0.SSIIBMUIDD = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                                xYSSIIBM0.SSIIBMUIDD = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                                xYSSIIBM0.SSIIBMNAT = " "
                                xYSSIIBM0.SSIIBMYVER = 0
                                '  '$JPL 2015-05-12----------------------------------------------------------------------
                                
                                Call cmdSSIIBM_Detail_Display("YSSIIBM*", "")
                                If arrHab(2) Then
                                    X = MsgBox("Vu ?", vbYesNo, "Nouveau compte IBM archivé")
                                    If X = vbYes Then
                                        cmdUpdate_Init
                                        mYSSIIBMH_Update = "Update"
                                        oldYSSIIBMH = xYSSIIBM0
                                        newYSSIIBMH = oldYSSIIBMH
                                        newYSSIIBMH.SSIIBMPRFK = " "
                                        newYSSIIBMH.SSIIBMYFCT = "VU "
                                        newYSSIIBMH.SSIIBMYUSR = usrName_UCase
                                        newYSSIIBMH.SSIIBMYAMJ = DSys
                                        newYSSIIBMH.SSIIBMYHMS = time_Hms
                                        oldYSSIIBM0 = oldYSSIIBMH
                                        newYSSIIBM0 = newYSSIIBMH
                                        Call cmdSSIJRN_IBM("<X:contrôle des modifications (Historique)>")
                                        cmdUpdate
                                    End If
                                    txtRTF.Visible = False
                                End If
                            Case "SAA"
                                usrYSSISAA0.SSISAANAT = ""
                                K2 = InStr(K1, fgSelect.Text, "|") + 1
                                usrYSSISAA0.SSISAAUIDX = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                                Call cmdSSISAA_Detail_Display("SAA")
                            Case "SAB"
                                usrYSSISAB0.SSISABNAT = ""
                                K2 = InStr(K1, fgSelect.Text, "|") + 1
                                usrYSSISAB0.SSISABUIDX = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                                Call cmdSSISAB_Detail_Display("SAB")
                           Case "SAB_H"
                                oldYSSISAB0.SSISABNAT = ""

                                Call cmdSSISAB_Detail_Display("YSSISAB*")
                                If arrHab(2) Then
                                    X = MsgBox("Vu ?", vbYesNo, "Nouveau compte SAB archivé")
                                    If X = vbYes Then
                                        cmdUpdate_Init
                                        mYSSISABH_Update = "Update"
                                        oldYSSISABH = xYSSISAB0
                                        newYSSISABH = oldYSSISABH
                                        newYSSISABH.SSISABPRFK = " "
                                        newYSSISABH.SSISABYFCT = "VU "
                                        newYSSISABH.SSISABYUSR = usrName_UCase
                                        newYSSISABH.SSISABYAMJ = DSys
                                        newYSSISABH.SSISABYHMS = time_Hms
                                        oldYSSISAB0 = oldYSSISABH
                                        newYSSISAB0 = newYSSISABH
                                        Call cmdSSIJRN_SAB("<X:contrôle des modifications (Historique)>")
                                        cmdUpdate
                                    End If
                                    txtRTF.Visible = False
                                End If
                            Case "WIN", "WIN_S"
                                rtfYSSIWIN0.SSIWINNAT = ""
                                K2 = InStr(K1, fgSelect.Text, "|") + 1
                                rtfYSSIWIN0.SSIWINGUID = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                                Call cmdSSIWIN_Detail_Display("YSSIWIN0")
                                 If arrHab(2) Then
                                    
                                    If wX = "WIN_S" Then
                                        X = MsgBox("Vu ?", vbYesNo, "Compte SUPPRIME par l'administrateur WIN")
                                    Else
                                        X = MsgBox("Voulez-vous le supprimer ?", vbYesNo, "Compte WIN orphelin")
                                    End If
                                    
                                    If X = vbYes Then
                                        cmdUpdate_Init
                                        mYSSIWIN0_Update = "Update"
                                        oldYSSIWIN0 = rtfYSSIWIN0
                                        newYSSIWIN0 = oldYSSIWIN0
                                        newYSSIWIN0.SSIWINPRFK = "S"
                                        newYSSIWIN0.SSIWINYFCT = "VU "
                                        newYSSIWIN0.SSIWINYUSR = usrName_UCase
                                        newYSSIWIN0.SSIWINYAMJ = DSys
                                        newYSSIWIN0.SSIWINYHMS = time_Hms
                                        Call cmdSSIJRN_WIN("<X:VU compte WIN supprimé>")
                                        cmdUpdate
                                    End If
                                    txtRTF.Visible = False
                                End If
                            Case "DIV"
                                rtfYSSIDIV0.SSIDIVNAT = ""
                                K2 = InStr(K1, fgSelect.Text, "|") + 1
                                rtfYSSIDIV0.SSIDIVUIDX = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                                K1 = InStr(K2, fgSelect.Text, "|") + 1
                                rtfYSSIDIV0.SSIDIVUIDD = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                                Call cmdSSIDIV_Detail_Display("YSSIDIV0")
                            Case "MEL"
                                rtfYSSIMEL0.SSIMELNAT = ""
                                K2 = InStr(K1, fgSelect.Text, "|") + 1
                                rtfYSSIMEL0.SSIMELUIDX = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                                K1 = InStr(K2, fgSelect.Text, "|") + 1
                                rtfYSSIMEL0.SSIMELUIDD = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                                Call cmdSSIMEL_Detail_Display("YSSIMEL0")
                            Case "TIC"
                                rtfYSSITIC0.SSITICNAT = ""
                                K2 = InStr(K1, fgSelect.Text, "|") + 1
                                rtfYSSITIC0.SSITICUIDX = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                                K1 = InStr(K2, fgSelect.Text, "|") + 1
                                rtfYSSITIC0.SSITICUIDD = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                                Call cmdSSITIC_Detail_Display("YSSITIC0")
                        End Select
                End Select
             Case "J", "H"
                If cmdSelect_SQL_K = "J" Then
                    Dim wFct As String
                    fgSelect.Col = 1: wFct = Trim(fgSelect.Text)
                    fgSelect.Col = 8
                    K1 = InStr(1, fgSelect.Text, "|") + 1
                    xYSSITXT0.SSITXTNAT = Mid$(fgSelect.Text, 1, K1 - 2)
                    K2 = InStr(K1, fgSelect.Text, "|") + 1
                    xYSSITXT0.SSITXTUIDN = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                    K1 = InStr(K2, fgSelect.Text, "|") + 1
                    xYSSITXT0.SSITXTDIDX = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                    K2 = InStr(K1, fgSelect.Text, "|") + 1
                    xYSSITXT0.SSITXTUIDX = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                    K1 = InStr(K2, fgSelect.Text, "|") + 1
                    xYSSITXT0.SSITXTUIDD = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                    K2 = InStr(K1, fgSelect.Text, "|") + 1
                    xYSSITXT0.SSITXTTLNK = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                    Call cmdSSITXT_Detail_txtRTF
                End If
                
                fgSelect.Col = 9
                K1 = InStr(1, fgSelect.Text, "|") + 1
                wX = Mid$(fgSelect.Text, 1, K1 - 2)
                Select Case Trim(wX)
                    Case "USR"
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       oldYSSIUSR0.SSIUSRNAT = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                       K1 = InStr(K2, fgSelect.Text, "|") + 1
                       oldYSSIUSR0.SSIUSRUIDN = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       oldYSSIUSR0.SSIUSRYVER = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                       If cmdSelect_SQL_K = "J" Then
                             If oldYSSIUSR0.SSIUSRUIDN = 0 Then oldYSSIUSR0.SSIUSRUIDN = xYSSITXT0.SSITXTUIDN
                            Call fraDetail_Load
                      Else
                            xYSSIUSR0 = oldYSSIUSR0
                            Call cmdSSIUSR_Detail_Load("YSSIUSR*")
                       End If
                    Case "DOM"
                        K2 = InStr(K1, fgSelect.Text, "|") + 1
                        oldYSSIUSR0.SSIUSRNAT = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                        K1 = InStr(K2, fgSelect.Text, "|") + 1
                        oldYSSIUSR0.SSIUSRUIDN = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                        If oldYSSIUSR0.SSIUSRUIDN = 0 Then oldYSSIUSR0.SSIUSRUIDN = xYSSITXT0.SSITXTUIDN
                        
                        oldYSSIDOM0.SSIDOMNAT = oldYSSIUSR0.SSIUSRNAT
                        oldYSSIDOM0.SSIDOMUIDN = oldYSSIUSR0.SSIUSRUIDN
                        K2 = InStr(K1, fgSelect.Text, "|") + 1
                        oldYSSIDOM0.SSIDOMDIDX = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                        K1 = InStr(K2, fgSelect.Text, "|") + 1
                        oldYSSIDOM0.SSIDOMUIDX = Trim(Mid$(fgSelect.Text, K2, K1 - K2 - 1))
                        K2 = InStr(K1, fgSelect.Text, "|") + 1
                        oldYSSIDOM0.SSIDOMUIDD = Val(Mid$(fgSelect.Text, K1, K2 - K1 - 1))
                       If cmdSelect_SQL_K = "J" Then
                            Call fraDetail_Load
                            If wFct <> "SUP" And oldYSSIDOM0.SSIDOMUIDX <> "" Then Call fraYSSIDOM0_Load
                       Else
                            xYSSIDOM0 = oldYSSIDOM0
                            Call cmdSSIDOM_Detail_Load("YSSIDOM*")
                       End If
                  Case "SAB"
                        K2 = InStr(K1, fgSelect.Text, "|") + 1
                        xYSSISAB0.SSISABNAT = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                        K1 = InStr(K2, fgSelect.Text, "|") + 1
                        xYSSISAB0.SSISABUIDX = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                        K2 = InStr(K1, fgSelect.Text, "|") + 1
                        xYSSISAB0.SSISABULOT = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                        K1 = InStr(K2, fgSelect.Text, "|") + 1
                        xYSSISAB0.SSISABYVER = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                        usrYSSISAB0 = xYSSISAB0
                        If cmdSelect_SQL_K = "J" Then
                            Call cmdSSISAB_Detail_Display("YSSISAB*")
                        Else
                            mRTF = ""
                            Call cmdSSISAB_Display("YSSISAB*")
                        End If

                    Case "SAA"
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       usrYSSISAA0.SSISAANAT = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                       K1 = InStr(K2, fgSelect.Text, "|") + 1
                       usrYSSISAA0.SSISAAUIDX = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       usrYSSISAA0.SSISAAUSEQ = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                       K1 = InStr(K2, fgSelect.Text, "|") + 1
                       usrYSSISAA0.SSISAAYVER = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                       Call cmdSSISAA_Detail_Display("YSSISAA*")
 
                    Case "IBM"
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       xYSSIIBM0.SSIIBMNAT = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                       K1 = InStr(K2, fgSelect.Text, "|") + 1
                       xYSSIIBM0.SSIIBMUIDD = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       xYSSIIBM0.SSIIBMYVER = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                       Call cmdSSIIBM_Detail_Display("YSSIIBM*", "")
                                          
 
                    Case "WIN"
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       rtfYSSIWIN0.SSIWINNAT = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                       K1 = InStr(K2, fgSelect.Text, "|") + 1
                       rtfYSSIWIN0.SSIWINUIDX = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       rtfYSSIWIN0.SSIWINYVER = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                       Call cmdSSIWIN_Detail_Display("YSSIWIN*")
                    Case "DIV"
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       rtfYSSIDIV0.SSIDIVNAT = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                       K1 = InStr(K2, fgSelect.Text, "|") + 1
                       rtfYSSIDIV0.SSIDIVUIDX = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       rtfYSSIDIV0.SSIDIVUIDD = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                       K1 = InStr(K2, fgSelect.Text, "|") + 1
                       rtfYSSIDIV0.SSIDIVYVER = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                      Call cmdSSIDIV_Detail_Display("YSSIDIV*")
                    Case "MEL"
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       rtfYSSIMEL0.SSIMELNAT = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                       K1 = InStr(K2, fgSelect.Text, "|") + 1
                       rtfYSSIMEL0.SSIMELUIDX = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       rtfYSSIMEL0.SSIMELUIDD = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                       K1 = InStr(K2, fgSelect.Text, "|") + 1
                       rtfYSSIMEL0.SSIMELYVER = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                      Call cmdSSIMEL_Detail_Display("YSSIMEL*")
                    Case "TIC"
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       rtfYSSITIC0.SSITICNAT = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                       K1 = InStr(K2, fgSelect.Text, "|") + 1
                       rtfYSSITIC0.SSITICUIDX = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       rtfYSSITIC0.SSITICUIDD = Mid$(fgSelect.Text, K1, K2 - K1 - 1)
                       K1 = InStr(K2, fgSelect.Text, "|") + 1
                       rtfYSSITIC0.SSITICYVER = Mid$(fgSelect.Text, K2, K1 - K2 - 1)
                      Call cmdSSITIC_Detail_Display("YSSITIC*")
                    Case "SAB_M"
                       K2 = InStr(K1, fgSelect.Text, "|") + 1
                       rtfYSSISAM0.SSISAMUIDD = Val(Mid$(fgSelect.Text, K1, K2 - K1 - 1))
                       K1 = InStr(K2, fgSelect.Text, "|") + 1
                       rtfYSSISAM0.SSISAMYVER = Val(Mid$(fgSelect.Text, K2, K1 - K2 - 1))
                      Call cmdSSISAM_Detail_Display("YSSISAM*")
             End Select
        End Select

End Sub

Public Sub cmdSelect_SQL_9_SAB()
Call cmdSelect_SQL_9_SAB_ZMNUUTI0
Call cmdSelect_SQL_9_SAB_ZMNU
End Sub

Public Sub cmdSelect_SQL_9_SAB_ZMNU()
Dim xSQL As String, xFCT As String, xOrig As String

Dim rsSab As New ADODB.Recordset

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_9_SAB_ZMNUHLA0"

Call cmdUpdate_Init
Call rsYSSISAB0_Init(prfYSSISAB0)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0" _
     & " where SSISABNAT <> ' ' and SSISABPRFK in ('?','#')" _
     & " order by SSISABNAT , SSISABULOT , SSISABUIDX"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call rsYSSISAB0_GetBuffer(rsSab, oldYSSISAB0)
    newYSSISAB0 = oldYSSISAB0
    newYSSISAB0.SSISABPRFK = " "
    mYSSISAB0_Update = "Update"
    Select Case oldYSSISAB0.SSISABPRFK
        Case "#": xFCT = "<FCT:MOD>"
        Case "?": xFCT = "<FCT:CRE>"
        Case Else: xFCT = "<FCT:" & oldYSSISAB0.SSISABPRFK & ">"
    End Select
    
    Select Case oldYSSISAB0.SSISABNAT
        Case "H": xOrig = "<ORIG:37>"
        Case "2": xOrig = "<ORIG:31>"
        Case "3": xOrig = "<ORIG:32>"
        Case "4": xOrig = "<ORIG:33>"
        Case "C": xOrig = "<ORIG:34>"
        Case "D": xOrig = "<ORIG:35>"
        Case "M": xOrig = "<ORIG:36>"
        Case "$": xOrig = "<ORIG:38>"
        Case Else: xOrig = "<ORIG:" & oldYSSISAB0.SSISABNAT & ">"
    End Select

        mYSSITXT0_JRN_Update = "New"
        Call rsYSSITXT0_Init(newYSSITXT0_JRN)
        newYSSITXT0_JRN.SSITXTNAT = "J"
        newYSSITXT0_JRN.SSITXTUIDN = 0
        newYSSITXT0_JRN.SSITXTDIDX = "SAB"
        newYSSITXT0_JRN.SSITXTUIDX = newYSSISAB0.SSISABUIDX
        newYSSITXT0_JRN.SSITXTUIDD = newYSSISAB0.SSISABUIDD
        
        newYSSITXT0_JRN.SSITXTINFO = "<Y:SAB|" & newYSSISAB0.SSISABNAT & "|" & Trim(newYSSISAB0.SSISABUIDX) & "|" & newYSSISAB0.SSISABULOT & "|>" _
                                    & "<UID:" & newYSSISAB0.SSISABUIDX & ">" _
                                    & xFCT & xOrig & "<UNOM:" & Trim(newYSSISAB0.SSISABUNOM) & ">"
        newYSSITXT0_JRN.SSITXTYAMJ = DSys
        newYSSITXT0_JRN.SSITXTYHMS = time_Hms
        newYSSITXT0_JRN.SSITXTYUSR = usrName_UCase
        
    Call cmdUpdate
    
    rsSab.MoveNext
Loop

Set rsSab = Nothing

'________________________________________________________________

Exit Sub

Error_Handler:
    
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Public Function cmdSSISAB_Display_ZMNUGRP0() As String
Dim xRTF As String
'____________________________________________________________________________________________
xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight11 Utilisateurs actifs :\highlight0\ulnone"

X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
     & " where SSISABNAT = ' ' and SSISABPRFX = '" & xYSSISAB0.SSISABUIDX & "'" _
     & " and SSISABPRFK <> 'X' order by SSISABUIDX"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    
    xRTF = xRTF & "\par\tab\fs16\cf13\b " & Mid$(rsSab("SSISABUIDX"), 1, 12) _
                & "\b0\cf2  : " & rsSab("SSISABUNOM")
    rsSab.MoveNext
Loop
'____________________________________________________________________________________________
xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight9 Utilisateurs INACTIFS :\highlight0\ulnone"

X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
     & " where SSISABNAT = ' ' and SSISABPRFX = '" & xYSSISAB0.SSISABUIDX & "'" _
     & " and SSISABPRFK = 'X' order by SSISABUIDX"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    
    xRTF = xRTF & "\par\tab\fs16\cf13\b " & Mid$(rsSab("SSISABUIDX"), 1, 12) _
                & "\b0\cf2  : " & rsSab("SSISABUNOM")
    rsSab.MoveNext
Loop

cmdSSISAB_Display_ZMNUGRP0 = "\fs18\cf13\highlight12\ul\b " & xYSSISAB0.SSISABUIDX & "\b0\ulnone\highlight0\cf1             (Profil) " & "\cf2  => " _
        & Mid$(xYSSISAB0.SSISABUNOM, 1, 10) & "  " _
        & Mid$(xYSSISAB0.SSISABUNOM, 11, 10) & "  " _
        & Mid$(xYSSISAB0.SSISABUNOM, 21, 10) & "  " _
        & "\par " _
        & "\par\tab\cf8 Actif        : \cf7 " & cmdSSISTAK_Detail_txtRTF(xYSSISAB0.SSISABSTAK, 0) _
        & "\par\tab\cf8 Id SSI       : \cf7 " & xYSSISAB0.SSISABUIDD _
        & "\par\tab\cf8 Evénement    : \cf7 " & xYSSISAB0.SSISABYFCT _
        & "  (v" & xYSSISAB0.SSISABYVER & ")" _
        & "\par\tab\cf8 màj par      : \cf7 " & xYSSISAB0.SSISABYUSR _
        & " le " & dateImp10_S(xYSSISAB0.SSISABYAMJ) & " " & timeImp8(xYSSISAB0.SSISABYHMS) _
        & "\par\cf8 _______________________________________________________________________\par" _
        & xRTF _
        & "\par\cf8 _______________________________________________________________________\par"

End Function
Public Sub cmdSSIUSR_Detail_Load(lFct As String)
Dim xRTF As String, whighlight As Integer, wSSIUSRDECH As String, wSSIUSRPRFD As String
'____________________________________________________________________________________________
whighlight = 11
If lFct = "YSSIUSR*" Then
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSRH " _
         & " where SSIUSRNAT = ' ' and SSIUSRUIDN = " & xYSSIUSR0.SSIUSRUIDN _
         & " and SSIUSRYVER = " & xYSSIUSR0.SSIUSRYVER
    Set rsSab = cnsab.Execute(X)
    If Not rsSab.EOF Then
        whighlight = 9
    Else
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
             & " where SSIUSRNAT = ' ' and SSIUSRUIDN = " & xYSSIUSR0.SSIUSRUIDN
        Set rsSab = cnsab.Execute(X)
    End If
Else
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
         & " where SSIUSRNAT = ' ' and SSIUSRUIDN = " & xYSSIUSR0.SSIUSRUIDN
    Set rsSab = cnsab.Execute(X)
End If

If Not rsSab.EOF Then
    Call rsYSSIUSR0_GetBuffer(rsSab, xYSSIUSR0)
    'Call cmdSSIUSR_Detail_txtRTF(whighlight)

    txtRTF.TextRTF = VB_RTF_Modèle
    
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", cmdSSIUSR_Detail_txtRTF(whighlight))
    
    Call txtRTF_Visible

End If
End Sub

Public Function cmdSSIUSR_Detail_txtRTF(lhighlight As Integer) As String
Dim xRTF As String, xTLNK As String
'____________________________________________________________________________________________
If xYSSIUSR0.SSIUSRTLNK <> 0 Then
         xYSSITXT0.SSITXTNAT = xYSSIUSR0.SSIUSRNAT
         xYSSITXT0.SSITXTUIDN = xYSSIUSR0.SSIUSRUIDN
         xYSSITXT0.SSITXTDIDX = ""
         xYSSITXT0.SSITXTUIDX = ""
         xYSSITXT0.SSITXTUIDD = 0
         xYSSITXT0.SSITXTTLNK = xYSSIUSR0.SSIUSRTLNK
    xTLNK = cmdSSITXT_Detail_txtRTF
End If

 xRTF = "\fs18\cf1\ul\highlight12 Utilisateur :\cf13\highlight" & lhighlight & "\b  " & Trim(xYSSIUSR0.SSIUSRUIDX) & "  \b0\ulnone\highlight0" _
         & "\par " _
         & "\par\tab\cf8 Actif        : \cf7 " & cmdSSISTAK_Detail_txtRTF(xYSSIUSR0.SSIUSRSTAK, xYSSIUSR0.SSIUSRDECH) _
         & "\par\tab\cf8 Id SSI       : \cf7 " & xYSSIUSR0.SSIUSRUIDN _
         & "\par\tab\cf8 Nom          : \cf7 " & xYSSIUSR0.SSIUSRUIDX _
         & "\par\tab\cf8 Profil BIA   : \cf7 " & Trim(xYSSIUSR0.SSIUSRPRFX) _
         & "\par\tab\cf8 Conforme ?   : \cf7 " & cmdSSIPRFK_Detail_txtRTF(xYSSIUSR0.SSIUSRPRFK, xYSSIUSR0.SSIUSRPRFD, xYSSIUSR0.SSIUSRPRFH) _
         & "\par\tab\cf8 Commentaire °: \cf7 " & xYSSIUSR0.SSIUSRTLNK _
         & "\par\tab\cf8 Evénement    : \cf7 " & xYSSIUSR0.SSIUSRYFCT _
         & "  (v" & xYSSIUSR0.SSIUSRYVER & ")" _
         & "\par\tab\cf8 màj par      : \cf7 " & xYSSIUSR0.SSIUSRYUSR _
         & " le " & dateImp10_S(xYSSIUSR0.SSIUSRYAMJ) & " " & timeImp8(xYSSIUSR0.SSIUSRYHMS) _
         & xTLNK _
         & "\par\cf8\b ______________________________________________________________________ \b0\par"
         

cmdSSIUSR_Detail_txtRTF = xRTF

End Function


Public Function cmdSSIWIN_Detail_txtRTF(lhighlight As Integer) As String
Dim xRTF As String, xTLNK As String, K1 As Integer, K2 As Integer, X As String, wUAC As Long, xUAC As String, xUAC2 As String, wUAC2 As Long
'____________________________________________________________________________________________
If rtfYSSIWIN0.SSIWINTLNK <> 0 Then
         xYSSITXT0.SSITXTNAT = rtfYSSIWIN0.SSIWINNAT
         'xYSSITXT0.SSITXTUIDN = rtfYSSIWIN0.SSIwinUIDN
         xYSSITXT0.SSITXTDIDX = "WIN"
         xYSSITXT0.SSITXTUIDX = rtfYSSIWIN0.SSIWINUIDX
         xYSSITXT0.SSITXTUIDD = rtfYSSIWIN0.SSIWINUIDD
         xYSSITXT0.SSITXTTLNK = rtfYSSIWIN0.SSIWINTLNK
    xTLNK = cmdSSITXT_Detail_txtRTF
End If

 xRTF = "\fs18\cf1\ul\highlight12 Compte Windows :\cf13\highlight" & lhighlight & "\b  " & Trim(rtfYSSIWIN0.SSIWINUIDX) & "  \b0\ulnone\highlight0" _
         & "\par " _
         & "\par\tab\cf8 Actif        : \cf7 " & cmdSSISTAK_Detail_txtRTF(rtfYSSIWIN0.SSIWINSTAK, 0) _
         & "\par\tab\cf8 Id SSI       : \cf7 " & rtfYSSIWIN0.SSIWINUIDD _
         & "\par\tab\cf8 Nom          : \cf7 " & rtfYSSIWIN0.SSIWINUIDX _
         & "\par\tab\cf8 Profil BIA   : \cf7 " & Trim(rtfYSSIWIN0.SSIWINPRFX) _
         & "\par\tab\cf8 Conforme ?   : \cf7 " & cmdSSIPRFK_Detail_txtRTF(rtfYSSIWIN0.SSIWINPRFK, 0, 0) _
         & "\par\tab\cf8 Commentaire °: \cf7 " & rtfYSSIWIN0.SSIWINTLNK _
         & "\par\tab\cf8 Evénement    : \cf7 " & rtfYSSIWIN0.SSIWINYFCT _
         & "  (v" & rtfYSSIWIN0.SSIWINYVER & ")" _
         & "\par\tab\cf8 màj par      : \cf7 " & rtfYSSIWIN0.SSIWINYUSR _
         & " le " & dateImp10_S(rtfYSSIWIN0.SSIWINYAMJ) & " " & timeImp8(rtfYSSIWIN0.SSIWINYHMS) _
         & xTLNK _
         & "\par\cf8\b ______________________________________________________________________ \b0\par"
         

X = rtfYSSIWIN0.SSIWININFO
K1 = InStr(X, "|") + 1
K2 = InStr(K1, X, "|")
xUAC = Mid$(X, K1, K2 - K1)
wUAC = Val(xUAC)
Select Case wUAC
    Case 512: xUAC2 = "\highlight11 " & arrUAC_Lib(9)
    Case 514: xUAC2 = "\highlight9\cf1" & arrUAC_Lib(9) & arrUAC_Lib(2)
    Case Else
        xUAC2 = "\highlight12\cf10 "
        wUAC2 = wUAC
        For K1 = 21 To 1 Step -1
            If wUAC2 >= arrUAC_Val(K1) Then
                xUAC2 = xUAC2 & arrUAC_Lib(K1) & " "
                wUAC2 = wUAC2 - arrUAC_Val(K1)
            End If
            
        Next K1
        
    
End Select
If xUAC <> "" Then X = Replace(X, xUAC, xUAC2 & "\highlight0", 1, 1)

X = "\par\tab\cf8 Name         : \cf13 " & X
X = Replace(X, "|", "\par\tab\cf8 UAC          : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 Class        : \cf13 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 Common name  : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 Last name    : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 Given Name   : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 Display name : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 UPN          : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 Company      : \cf14 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 Department   : \cf2 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 scriptPath   : \cf13 ", 1, 1)
X = Replace(X, ".cmd", ".cmd \cf2 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 whenCreated  : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 Description  : \cf2 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 DN           : \cf7 ", 1, 1)
X = Replace(X, "OU=", "\par\tab              : OU=", 1, 1)
X = Replace(X, "DC=", "\par\tab              : DC=", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 mail         : \cf13 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 mailnickname : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 physicalOffic: \cf2 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 sAMAccountNam: \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 txtEncodedAdd: \cf7 ", 1, 1)
X = Replace(X, ";o=Exchange", "\par\tab              : ;o=Exchange", 1, 1)

X = Replace(X, "|", "\par\tab\cf8 userWorkstat : \cf2 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 accountExpire: \highlight12\cf2 ", 1, 1)
X = Replace(X, "|", "\par\highlight0\tab\cf8              : \cf2 ", 1, 1)

xRTF = Replace(xRTF, "{", " ") & X
'___________________________________________________________________________________________________

If cmdSelect_SQL_K <> "1" Then
    xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight11 Utilisateurs actifs :\highlight0\ulnone"
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
         & " where SSIDOMNAT = ' '  and SSIDOMDIDX = 'WIN' and SSIDOMPRFX = '" & rtfYSSIWIN0.SSIWINUIDX & "'" _
         & " and SSIDOMPRFK <> 'X' and SSIWINNAT = ' ' and SSIWINUIDD = SSIDOMUIDD order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        'kNAT_U = kNAT_U + 1
        
        xRTF = xRTF & "\par\tab\fs16\cf13\b " & rsSab("SSIDOMUIDX") _
                    & "\b0\cf2  : " & Trim(rsSab("SSIWINUNOM"))
        rsSab.MoveNext
    Loop

    xRTF = xRTF & "\par\cf8 _______________________________________________________________________\par"
    xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight9 Utilisateurs INACTIFS :\highlight0\ulnone"
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
         & " where SSIDOMNAT = ' '  and SSIDOMDIDX = 'WIN' and SSIDOMPRFX = '" & rtfYSSIWIN0.SSIWINUIDX & "'" _
         & " and SSIDOMPRFK = 'X' and SSIWINNAT = ' ' and SSIWINUIDD = SSIDOMUIDD order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        'kNAT_U = kNAT_U + 1
        
        xRTF = xRTF & "\par\tab\fs16\cf13\b " & rsSab("SSIDOMUIDX") _
                    & "\b0\cf2  : " & Trim(rsSab("SSIWINUNOM"))
        rsSab.MoveNext
    Loop
End If
xRTF = xRTF & "\par\cf8 _______________________________________________________________________\par"

cmdSSIWIN_Detail_txtRTF = xRTF

End Function

Public Function cmdSSIMEL_Detail_txtRTF(lhighlight As Integer) As String
Dim xRTF As String, xTLNK As String, K1 As Integer, K2 As Integer, X As String, wSSIMELPRFX As String
'____________________________________________________________________________________________
If rtfYSSIMEL0.SSIMELTLNK <> 0 Then
         xYSSITXT0.SSITXTNAT = rtfYSSIMEL0.SSIMELNAT
         'xYSSITXT0.SSITXTUIDN = rtfYSSIMEL0.SSIMELUIDN
         xYSSITXT0.SSITXTDIDX = "MEL"
         xYSSITXT0.SSITXTUIDX = rtfYSSIMEL0.SSIMELUIDX
         xYSSITXT0.SSITXTUIDD = rtfYSSIMEL0.SSIMELUIDD
         xYSSITXT0.SSITXTTLNK = rtfYSSIMEL0.SSIMELTLNK
    xTLNK = cmdSSITXT_Detail_txtRTF
End If

 xRTF = "\fs18\cf1\ul\highlight12 Compte MEL :\cf13\highlight" & lhighlight & "\b  " & Trim(rtfYSSIMEL0.SSIMELUIDX) & "  \b0\ulnone\highlight0" _
         & "\par " _
         & "\par\tab\cf8 Actif        : \cf7 " & cmdSSISTAK_Detail_txtRTF(rtfYSSIMEL0.SSIMELSTAK, 0) _
         & "\par\tab\cf8 Id SSI       : \cf7 " & rtfYSSIMEL0.SSIMELUIDD _
         & "\par\tab\cf8 Nom          : \cf7 " & rtfYSSIMEL0.SSIMELUIDX _
         & "\par\tab\cf8 Profil BIA   : \cf13\highlight12 " & Trim(rtfYSSIMEL0.SSIMELPRFX) & "\highlight0 " _
         & "\par\tab\cf8 Conforme ?   : \cf7 " & cmdSSIPRFK_Detail_txtRTF(rtfYSSIMEL0.SSIMELPRFK, 0, 0) _
         & "\par\tab\cf8 Commentaire °: \cf7 " & rtfYSSIMEL0.SSIMELTLNK _
         & "\par\tab\cf8 Evénement    : \cf7 " & rtfYSSIMEL0.SSIMELYFCT _
         & "  (v" & rtfYSSIMEL0.SSIMELYVER & ")" _
         & "\par\tab\cf8 màj par      : \cf7 " & rtfYSSIMEL0.SSIMELYUSR _
         & " le " & dateImp10_S(rtfYSSIMEL0.SSIMELYAMJ) & " " & timeImp8(rtfYSSIMEL0.SSIMELYHMS) _
         & xTLNK _
         & "\par\cf8\b ______________________________________________________________________ \b0\par"
         

X = rtfYSSIMEL0.SSIMELINFO

X = "\par\tab\cf8 PrimarySmtpAddress: \cf13 " & X
X = Replace(X, "|", "\par\tab\cf8 Alias             : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 LastName          : \cf13 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 FirstName         : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 Name              : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 JobTitle          : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 OfficeLocation    : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 BusinessTelNumber : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 CompanyName       : \cf14 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 Department        : \cf2 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 MobileTelNumber   : \cf13 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 PostalCode        : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 StreetAddress     : \cf2 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 Comments          : \cf7 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 Address           : \cf13 ", 1, 1)
X = Replace(X, "|", "\par\tab\cf8 id : \cf7 ", 1, 1)

xRTF = Replace(xRTF, "{", " ") & X

'___________________________________________________________________________________________________

If cmdSelect_SQL_K = "1" Then
    xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight11 paramétrage envoi automatique de courriels :\highlight0\ulnone\par"
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
         & " where SSIMELNAT = '@'  and SSIMELINFO like '%" & rtfYSSIMEL0.SSIMELUNOM & "%'" _
         & " order by SSIMELUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        xRTF = xRTF & "\par\tab\tab\fs18\cf13\b " & Trim(rsSab("SSIMELUIDX")) _
                    & "\b0\cf2  : " & Trim(rsSab("SSIMELUNOM"))
        rsSab.MoveNext
    Loop
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
         & " where SSIMELNAT = ' '  and SSIMELUIDD =" & rtfYSSIMEL0.SSIMELUIDD _
         & " order by SSIMELUIDX"
    Set rsSab_X = cnsab.Execute(X)
    
    Do While Not rsSab_X.EOF
        wSSIMELPRFX = Trim(rsSab_X("SSIMELPRFX"))
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
             & " where SSIMELNAT = '@'  and SSIMELINFO like '%" & wSSIMELPRFX & "%'" _
             & " order by SSIMELUIDX"
        Set rsSab = cnsab.Execute(X)
        
        Do While Not rsSab.EOF
            xRTF = xRTF & "\par\fs18\cf7\b      " & wSSIMELPRFX & "    \cf13 " & Trim(rsSab("SSIMELUIDX")) _
                        & "\b0\cf2\tab  : " & Trim(rsSab("SSIMELUNOM"))
            rsSab.MoveNext
        Loop
        rsSab_X.MoveNext
    Loop
Else
    xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight11 Utilisateurs actifs :\highlight0\ulnone"
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIMEL0 " _
         & " where SSIDOMNAT = ' '  and SSIDOMDIDX = 'MEL' and SSIDOMPRFX = '" & rtfYSSIMEL0.SSIMELUIDX & "'" _
         & " and SSIDOMPRFK <> 'X' and SSIMELNAT = ' ' and SSIMELUIDX = SSIDOMUIDX order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        'kNAT_U = kNAT_U + 1
        
        xRTF = xRTF & "\par\tab\fs18\cf13\b " & rsSab("SSIDOMUIDX") _
                    & "\b0\cf2  : " & Trim(rsSab("SSIMELUNOM"))
        rsSab.MoveNext
    Loop

    xRTF = xRTF & "\par\cf8 _______________________________________________________________________\par"
    xRTF = xRTF & "\par\par\fs18\ul\cf1\highlight9 Utilisateurs INACTIFS :\highlight0\ulnone"
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIMEL0 " _
         & " where SSIDOMNAT = ' '  and SSIDOMDIDX = 'MEL' and SSIDOMPRFX = '" & rtfYSSIMEL0.SSIMELUIDX & "'" _
         & " and SSIDOMPRFK = 'X' and SSIMELNAT = ' ' and SSIMELUIDX = SSIDOMUIDX order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        'kNAT_U = kNAT_U + 1
        
        xRTF = xRTF & "\par\tab\fs18\cf13\b " & rsSab("SSIDOMUIDX") _
                    & "\b0\cf2  : " & Trim(rsSab("SSIMELUNOM"))
        rsSab.MoveNext
    Loop
End If
xRTF = xRTF & "\par\cf8 _______________________________________________________________________\par"

cmdSSIMEL_Detail_txtRTF = xRTF

End Function
Public Function cmdSSITIC_Detail_txtRTF(lhighlight As Integer) As String
Dim xRTF As String, xTLNK As String, K1 As Integer, K2 As Integer, X As String, wSSITICINFO As String
Dim mRoles As String, blnEnd As Boolean
'____________________________________________________________________________________________
If rtfYSSITIC0.SSITICTLNK <> 0 Then
         xYSSITXT0.SSITXTNAT = rtfYSSITIC0.SSITICNAT
         'xYSSITXT0.SSITXTUIDN = rtfYSSITIC0.SSITICUIDN
         xYSSITXT0.SSITXTDIDX = "TIC"
         xYSSITXT0.SSITXTUIDX = rtfYSSITIC0.SSITICUIDX
         xYSSITXT0.SSITXTUIDD = rtfYSSITIC0.SSITICUIDD
         xYSSITXT0.SSITXTTLNK = rtfYSSITIC0.SSITICTLNK
    xTLNK = cmdSSITXT_Detail_txtRTF
End If

 xRTF = "\fs18\cf1\ul\highlight12 Compte TIC :\cf13\highlight" & lhighlight & "\b  " & Trim(rtfYSSITIC0.SSITICUIDX) & "  \b0\ulnone\highlight0" _
         & "\par " _
         & "\par\tab\cf8 Actif        : \cf7 " & cmdSSISTAK_Detail_txtRTF(rtfYSSITIC0.SSITICSTAK, 0) _
         & "\par\tab\cf8 Id SSI       : \cf7 " & rtfYSSITIC0.SSITICUIDD _
         & "\par\tab\cf8 Compte       : \cf7 " & rtfYSSITIC0.SSITICUIDX _
         & "\par\tab\cf8 Nom          : \cf7 " & rtfYSSITIC0.SSITICUNOM _
         & "\par\tab\cf8 Profil BIA   : \cf13\highlight12 " & Trim(rtfYSSITIC0.SSITICPRFX) & "\highlight0 " _
         & "\par\tab\cf8 Conforme ?   : \cf7 " & cmdSSIPRFK_Detail_txtRTF(rtfYSSITIC0.SSITICPRFK, 0, 0) _
         & "\par\tab\cf8 Commentaire °: \cf7 " & rtfYSSITIC0.SSITICTLNK _
         & "\par\tab\cf8 Evénement    : \cf7 " & rtfYSSITIC0.SSITICYFCT _
         & "  (v" & rtfYSSITIC0.SSITICYVER & ")" _
         & "\par\tab\cf8 màj par      : \cf7 " & rtfYSSITIC0.SSITICYUSR _
         & " le " & dateImp10_S(rtfYSSITIC0.SSITICYAMJ) & " " & timeImp8(rtfYSSITIC0.SSITICYHMS) _
         & xTLNK _
         & "\par\cf8\b ______________________________________________________________________ \b0\par"
         

wSSITICINFO = rtfYSSITIC0.SSITICINFO


xRTF = Replace(xRTF, "{", " ") & wSSITICINFO

'___________________________________________________________________________________________________

Select Case rtfYSSITIC0.SSITICNAT
    Case " ", "$"
        xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight12 ATHIC : Rôles :\highlight0\ulnone\par"
        
        For K1 = 1 To Len(wSSITICINFO)
            If Mid$(wSSITICINFO, K1, 1) = "R" Then
                X = "select SSITICUNOM from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
                     & " where SSITICNAT = 'R'  and SSITICUIDD = " & K1
                Set rsSab = cnsab.Execute(X)
                If Not rsSab.EOF Then
                    xRTF = xRTF & "\par\tab\fs18\cf13\b " & K1 _
                               & "\b0\cf2  : " & Trim(rsSab("SSITICUNOM"))
                End If
            End If
        Next K1
    Case "D"
        mRoles = ""
        xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight12 ATHIC : Rôles :\highlight0\ulnone\par"
         X = "select SSITICUIDD , SSITICUNOM from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
             & " where SSITICNAT = 'R'  and substring(SSITICINFO," & rtfYSSITIC0.SSITICUIDD & ",1) = 'D'" _
             & " order by SSITICUIDD"
        Set rsSab = cnsab.Execute(X)
        Do Until rsSab.EOF
            mRoles = mRoles & rsSab("SSITICUIDD") & " "
            xRTF = xRTF & "\par\tab\fs18\cf13\b " & rsSab("SSITICUIDD") _
                       & "\b0\cf2  : " & Trim(rsSab("SSITICUNOM"))
            rsSab.MoveNext
        Loop
        
        xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight12 ATHIC : Profils :\highlight0\ulnone\par"
        K1 = 0: blnEnd = False
        Do
            K2 = Val(Space_Scan(mRoles, K1))
            If K2 = 0 Then
                blnEnd = True
            Else
                  X = "select  SSITICUIDX from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
                      & " where SSITICNAT = '$'  and substring(SSITICINFO," & K2 & ",1) = 'R'" _
                      & " and SSITICPRFK <> 'X' order by SSITICUIDD"
                 Set rsSab = cnsab.Execute(X)
                 Do Until rsSab.EOF
                
                     xRTF = xRTF & "\par\tab\fs18\cf13\b " & K2 _
                                & "\b0\cf2  : " & Trim(rsSab("SSITICUIDX"))
                     rsSab.MoveNext
                 Loop
            End If
        Loop Until blnEnd
        
        xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight12 ATHIC : Comptes :\highlight0\ulnone\par"
        K1 = 0: blnEnd = False
        Do
            K2 = Val(Space_Scan(mRoles, K1))
            If K2 = 0 Then
                blnEnd = True
            Else
                  X = "select  SSITICUIDX from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
                      & " where SSITICNAT = ' '  and substring(SSITICINFO," & K2 & ",1) = 'R'" _
                      & " and SSITICPRFK <> 'X' order by SSITICUIDD"
                 Set rsSab = cnsab.Execute(X)
                 Do Until rsSab.EOF
                
                     xRTF = xRTF & "\par\tab\fs18\cf13\b " & K2 _
                                & "\b0\cf2  : " & Trim(rsSab("SSITICUIDX"))
                     rsSab.MoveNext
                 Loop
            End If
        Loop Until blnEnd
    Case "R"
        xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight12 ATHIC : Profils :\highlight0\ulnone\par"
          X = "select SSITICUIDD, SSITICUIDX from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
              & " where SSITICNAT = '$'  and substring(SSITICINFO," & rtfYSSITIC0.SSITICUIDD & ",1) = 'R'" _
              & " and SSITICPRFK <> 'X' order by SSITICUIDX"
         Set rsSab = cnsab.Execute(X)
         Do Until rsSab.EOF
        
             xRTF = xRTF & "\par\tab\fs18\cf13\b " & rsSab("SSITICUIDD") _
                        & "\b0\cf2  : " & Trim(rsSab("SSITICUIDX"))
             rsSab.MoveNext
         Loop
        
        xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight12 ATHIC : Comptes :\highlight0\ulnone\par"
          X = "select  SSITICPRFX, SSITICUIDX from " & paramIBM_Library_SABSPE & ".YSSITIC0" _
              & " where SSITICNAT = ' '  and substring(SSITICINFO," & rtfYSSITIC0.SSITICUIDD & ",1) = 'R'" _
              & " and SSITICPRFK <> 'X' order by SSITICPRFX"
         Set rsSab = cnsab.Execute(X)
         Do Until rsSab.EOF
        
             xRTF = xRTF & "\par\tab\fs18\cf13\b " & rsSab("SSITICPRFX") _
                        & "\b0\cf2  : " & Trim(rsSab("SSITICUIDX"))
             rsSab.MoveNext
         Loop

End Select

'___________________________________________________________________________________________________


'If cmdSelect_SQL_K <> "1" Then
'Else
If rtfYSSITIC0.SSITICNAT = "$" Then
    xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight11 Utilisateurs actifs :\highlight0\ulnone"
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSITIC0 " _
         & " where SSIDOMNAT = ' '  and SSIDOMDIDX = 'TIC' and SSIDOMPRFX = '" & rtfYSSITIC0.SSITICUIDX & "'" _
         & " and SSIDOMPRFK <> 'X' and SSITICNAT = ' ' and SSITICUIDX = SSIDOMUIDX order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        'kNAT_U = kNAT_U + 1
        
        xRTF = xRTF & "\par\tab\fs18\cf13\b " & rsSab("SSIDOMUIDX") _
                    & "\b0\cf2  : " & Trim(rsSab("SSITICUNOM"))
        rsSab.MoveNext
    Loop

    xRTF = xRTF & "\par\cf8 _______________________________________________________________________\par"
    xRTF = xRTF & "\par\par\fs18\ul\cf1\highlight9 Utilisateurs INACTIFS :\highlight0\ulnone"
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSITIC0 " _
         & " where SSIDOMNAT = ' '  and SSIDOMDIDX = 'TIC' and SSIDOMPRFX = '" & rtfYSSITIC0.SSITICUIDX & "'" _
         & " and SSIDOMPRFK = 'X' and SSITICNAT = ' ' and SSITICUIDX = SSIDOMUIDX order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        'kNAT_U = kNAT_U + 1
        
        xRTF = xRTF & "\par\tab\fs18\cf13\b " & rsSab("SSIDOMUIDX") _
                    & "\b0\cf2  : " & Trim(rsSab("SSITICUNOM"))
        rsSab.MoveNext
    Loop
End If
xRTF = xRTF & "\par\cf8 _______________________________________________________________________\par"

cmdSSITIC_Detail_txtRTF = xRTF

End Function

Public Function cmdSSISAM_Detail_txtRTF(lhighlight As Integer) As String
Dim xRTF As String, xTLNK As String, K1 As Integer, K2 As Integer, X As String, wHBTI01LIB As String
'Dim mRoles As String, blnEnd As Boolean
'____________________________________________________________________________________________
If rtfYSSISAM0.SSISAMTLNK <> 0 Then
         xYSSITXT0.SSITXTNAT = rtfYSSISAM0.SSISAMNAT
         'xYSSITXT0.SSITXTUIDN = rtfYSSISAM0.SSISAMUIDN
         xYSSITXT0.SSITXTDIDX = "SAB_M"
         xYSSITXT0.SSITXTUIDX = rtfYSSISAM0.SSISAMUIDX
         xYSSITXT0.SSITXTUIDD = rtfYSSISAM0.SSISAMUIDD
         xYSSITXT0.SSITXTTLNK = rtfYSSISAM0.SSISAMTLNK
    xTLNK = cmdSSITXT_Detail_txtRTF
End If

wHBTI01LIB = "select HBTI01LIB from " & paramIBM_Library_SAB & ".ZHBTI010 " _
     & " where HBTI01LAN = '1' and HBTI01APP = '" & rtfYSSISAM0.SSISAMAPP & "'" _
     & " and HBTI01COD ='" & rtfYSSISAM0.SSISAMCOD & "'"
Set rsSab = cnsab.Execute(wHBTI01LIB)
If rsSab.EOF Then
    wHBTI01LIB = Trim(rtfYSSISAM0.SSISAMUIDX)
Else
    wHBTI01LIB = Trim(rsSab("HBTI01LIB"))
End If

 xRTF = "\fs18\cf1\ul\highlight12 SAB Métier         :\cf13\highlight" & lhighlight & "\b  " & wHBTI01LIB & "  \b0\ulnone\highlight0" _
         & "\par " _
         & "\par\tab\cf8 Id SSI               : \cf7 " & rtfYSSISAM0.SSISAMUIDD _
         & "\par\tab\cf8 Habilitation         : \cf13 " & rtfYSSISAM0.SSISAMUIDX _
         & "\par\tab\cf8 Commentaire °        : \cf7 " & rtfYSSISAM0.SSISAMTLNK _
         & "\par\tab\cf8 Evénement            : \cf7 " & rtfYSSISAM0.SSISAMYFCT _
         & "  (v" & rtfYSSISAM0.SSISAMYVER & ")" _
         & "\par\tab\cf8 màj par              : \cf7 " & rtfYSSISAM0.SSISAMYUSR _
         & " le " & dateImp10_S(rtfYSSISAM0.SSISAMYAMJ) & " " & timeImp8(rtfYSSISAM0.SSISAMYHMS) _
         & xTLNK _
         & "\par\cf8\b ______________________________________________________________________ \b0\par" _
         & "\par\tab\cf1 Eta/Lot/GRP/classe   : \cf13 " & rtfYSSISAM0.SSISAMETA & " " & rtfYSSISAM0.SSISAMREF & " " & rtfYSSISAM0.SSISAMGRP & " " & rtfYSSISAM0.SSISAMCLA _
         & "\par\tab\cf1 Application/code     : \cf13 " & rtfYSSISAM0.SSISAMAPP & " " & rtfYSSISAM0.SSISAMCOD _
         & "\par\tab\cf1 Ag/Srv/Sous-service  : \cf13 " & rtfYSSISAM0.SSISAMAGE & " " & rtfYSSISAM0.SSISAMSER & " " & rtfYSSISAM0.SSISAMSSE _
         & "\par\tab\cf1 Opé/nat/Compte/autres: \cf13 " & rtfYSSISAM0.SSISAMOPE & " " & rtfYSSISAM0.SSISAMNAT & " " & rtfYSSISAM0.SSISAMPRD & " " & rtfYSSISAM0.SSISAMAUT
         

       X = "\par\tab\cf1 Droits fonctionnalité: \cf13 " & rtfYSSISAM0.SSISAMFON _
         & "\par\tab\cf1 Droits données       : \cf13 " & rtfYSSISAM0.SSISAMDON _
         & "\par\tab\cf1 Droits caisse s      : \cf13 " & rtfYSSISAM0.SSISAMCAI _
         & "\par\tab\cf1 Mt plafond           : \cf13 " & Format(rtfYSSISAM0.SSISAMMON, "### ### ##0.00") & " " & rtfYSSISAM0.SSISAMDEV _
         & "\par\tab\cf1 Droits délais        : \cf13 " & rtfYSSISAM0.SSISAMDLY _
         & "\par\tab\cf1 Profil               : \cf13 " & rtfYSSISAM0.SSISAMPRO _
         & "\par\tab\cf1 Client (*5)          : \cf13 " & rtfYSSISAM0.SSISAMCLI _
         & "\par\tab\cf1 Droits utilisateur   : \cf13 " & rtfYSSISAM0.SSISAMEIC _
         & "\par\tab\cf1 Droits mandat        : \cf13 " & rtfYSSISAM0.SSISAMSDD _
         & "\par\tab\cf1 Droits opération     : \cf13 " & rtfYSSISAM0.SSISAMDRO _
         & "\par\tab\cf1 Sup commissions      : \cf13 " & rtfYSSISAM0.SSISAMSUC _
         & "\par\tab\cf1 % commissions        : \cf13 " & rtfYSSISAM0.SSISAMPRO _
         & "\par\tab\cf1 Nbj marge +          : \cf13 " & rtfYSSISAM0.SSISAMNJ1 _
         & "\par\tab\cf1 Type nbj marge +     : \cf13 " & rtfYSSISAM0.SSISAMTJ1 _
         & "\par\tab\cf1 Nbj marge -          : \cf13 " & rtfYSSISAM0.SSISAMNJ2 _
         & "\par\tab\cf1 Type nbj marge -    : \cf13 " & rtfYSSISAM0.SSISAMTJ2 _
         & "\par\tab\cf1 % mt arrondi         : \cf13 " & rtfYSSISAM0.SSISAMPRA _
         & "\par\tab\cf1 Droits échelles      : \cf13 " & rtfYSSISAM0.SSISAMECH
 xRTF = xRTF & X
'___________________________________________________________________________________________________


cmdSSISAM_Detail_txtRTF = xRTF

End Function



Public Function cmdSSISAW_Detail_txtRTF(lhighlight As Integer) As String
Dim xRTF As String, xTLNK As String, K1 As Integer, K2 As Integer, X As String, wHBTI01LIB As String
Dim xREEL As String, xRECEP As String, xSAISI As String, xMODIF As String, xSUPPR As String, xENVOI As String
Dim xDETAI As String, xPAGE As String, xEDREF As String, xEDVAL As String, xEDMES As String, xEDDEV As String
Dim xMONTA As String, xVALID As String, xMECON As String, xMONT2 As String, xEDGEN As String, xDRCOP As String
Dim XControl As String
Dim xRTF_1 As String
'____________________________________________________________________________________________
If usrYSSISAB0.SSISABTLNK <> 0 Then
         xYSSITXT0.SSITXTNAT = usrYSSISAB0.SSISABNAT
         'xYSSITXT0.SSITXTUIDN = usrYSSISAB0.SSISABUIDN
         xYSSITXT0.SSITXTDIDX = "SAB_M"
         xYSSITXT0.SSITXTUIDX = usrYSSISAB0.SSISABUIDX
         xYSSITXT0.SSITXTUIDD = usrYSSISAB0.SSISABUIDD
         xYSSITXT0.SSITXTTLNK = usrYSSISAB0.SSISABTLNK
    xTLNK = cmdSSITXT_Detail_txtRTF
End If

'  UTILI      PIC X(10).
'  SERVI      PIC X(02).
'  SSERV      PIC X(02).
'  REEL-X        PIC X.
'  RECEP         PIC X.

'  DETAI         PIC X.
'  PAGE-X        PIC X.
'  EDREF         PIC 9.
'  EDVAL         PIC 9.
'  EDMES         PIC 9.
'  EDDEV         PIC 9.
'  SAISI         PIC X.
'  MODIF         PIC X.
'  SUPPR         PIC X.

'  MONTA         PIC 9(10).
'  VALID-X       PIC X.
'  MECON         PIC X(120).
'  ENVOI         PIC X.
'  MONT2         PIC 9(10).
'  EDGEN         PIC X.
'  DRCOP         PIC X.

xREEL = IIf(Mid$(usrYSSISAB0.SSISABINFO, 15, 1) = "1", "\cf13 PROD", "\cf16\highlight14 Test \highlight0")
xRECEP = IIf(Mid$(usrYSSISAB0.SSISABINFO, 16, 1) = "O", "\cf13 Oui", "\cf7 Non")
xDETAI = IIf(Mid$(usrYSSISAB0.SSISABINFO, 17, 1) = "O", "\cf13 Oui", "\cf7 Non")
xPAGE = IIf(Mid$(usrYSSISAB0.SSISABINFO, 18, 1) = "O", "\cf13 Oui", "\cf7 Non")
xEDREF = Mid$(usrYSSISAB0.SSISABINFO, 19, 1)
xEDVAL = Mid$(usrYSSISAB0.SSISABINFO, 20, 1)
xEDMES = Mid$(usrYSSISAB0.SSISABINFO, 21, 1)
xEDDEV = Mid$(usrYSSISAB0.SSISABINFO, 22, 1)
xSAISI = IIf(Mid$(usrYSSISAB0.SSISABINFO, 23, 1) = "O", "\cf13 Oui", "\cf7 Non")
xMODIF = IIf(Mid$(usrYSSISAB0.SSISABINFO, 24, 1) = "O", "\cf13 Oui", "\cf7 Non")
xSUPPR = IIf(Mid$(usrYSSISAB0.SSISABINFO, 25, 1) = "O", "\cf13 Oui", "\cf7 Non")
xMONTA = Format(Mid$(usrYSSISAB0.SSISABINFO, 26, 10), "### ### ### ###")
Select Case Mid$(usrYSSISAB0.SSISABINFO, 36, 1)
    Case "3": xVALID = "\cf13 Vérification"
    Case "2": xVALID = "\cf13\highlight12 Validation unique \highlight0"
    Case "1": xVALID = "\cf16\highlight14 Validation globale \highlight0"
    Case Else: xVALID = "\cf12\highlight2 code validation inconnu " & Mid$(usrYSSISAB0.SSISABINFO, 36, 1) & " \highlight0"
End Select

xMECON = Mid$(usrYSSISAB0.SSISABINFO, 37, 120)
xENVOI = IIf(Mid$(usrYSSISAB0.SSISABINFO, 157, 1) = "O", "\highlight12\cf13 Oui \highlight0", "\cf7 Non")
xMONT2 = Format(Mid$(usrYSSISAB0.SSISABINFO, 158, 10), "### ### ### ###")
xEDGEN = IIf(Mid$(usrYSSISAB0.SSISABINFO, 168, 1) = "O", "\cf13 Oui", "\cf7 Non")
xDRCOP = IIf(Mid$(usrYSSISAB0.SSISABINFO, 169, 1) = "1", "\cf13 Oui", "\cf7 Non")

If Mid$(usrYSSISAB0.SSISABINFO, 23, 1) = "O" _
And Mid$(usrYSSISAB0.SSISABINFO, 36, 1) <> "3" Then
    XControl = "\cf12\highlight10 incompatibilité des droits : saisir / valider \highlight0"
    xSAISI = "\cf16\highlight14 Oui \highlight0"
End If
If Mid$(usrYSSISAB0.SSISABINFO, 157, 1) = "O" _
And Mid$(usrYSSISAB0.SSISABINFO, 36, 1) = "3" Then
    XControl = "\cf12\highlight10 incompatibilité des droits : vérification / envoi \highlight0"
    xENVOI = "\cf16\highlight14 Oui \highlight0"
End If
xRTF_1 = "\par\par\tab\cf1 Production/Test      : \cf13 " & xREEL _
         & "\par\tab\cf1 Droit à réception    : " & xRECEP _
         & "\par\tab\cf1 Droit suppression    : " & xSUPPR _
         & "\par\tab\cf1 Droit modification   : " & xMODIF _
         & "\par\tab\cf1 Droit copie msg      : " & xDRCOP _
         & "\par\par\tab\cf1 Edition détaillée    : " & xDETAI _
         & "\par\tab\cf1 Edition 1 msg / page : " & xPAGE _
         & "\par\tab\cf1 Edition à génération : " & xEDGEN _
         & "\par\tab\cf1 Edition tri référence: \cf13 " & xEDREF _
         & "\par\tab\cf1 Edition tri date val : \cf13 " & xEDVAL _
         & "\par\tab\cf1 Edition tri type msg : \cf13 " & xEDMES _
         & "\par\tab\cf1 Edition tri devise   : \cf13 " & xEDDEV _
         & "\par"



xRTF = "\fs18\cf1\ul\highlight12 SAB droits SWIFT   :\cf13\highlight" & lhighlight & "\b  " & usrYSSISAB0.SSISABUIDX & "  \b0\ulnone\highlight0" _
         & "\par " _
         & "\par\tab\cf8 Id SSI               : \cf7 " & usrYSSISAB0.SSISABUIDD _
         & "\par\tab\cf8 Habilitation         : \cf13 " & usrYSSISAB0.SSISABUIDX _
         & "\par\tab\cf8 Commentaire °        : \cf7 " & usrYSSISAB0.SSISABTLNK _
         & "\par\tab\cf8 Evénement            : \cf7 " & usrYSSISAB0.SSISABYFCT _
         & "  (v" & usrYSSISAB0.SSISABYVER & ")" _
         & "\par\tab\cf8 màj par              : \cf7 " & usrYSSISAB0.SSISABYUSR _
         & " le " & dateImp10_S(usrYSSISAB0.SSISABYAMJ) & " " & timeImp8(usrYSSISAB0.SSISABYHMS) _
         & xTLNK _
         & "\par\cf8\b ______________________________________________________________________ \b0\par" _
         & "\par\tab\cf1 Compte               : \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 1, 10) _
         & "\par\tab\cf1 Service/sous-service : \cf13 " & Mid$(usrYSSISAB0.SSISABINFO, 11, 2) & "  " & Mid$(usrYSSISAB0.SSISABINFO, 13, 2) _
         & "\par\tab" & XControl _
         & "\par\tab\cf1 Droit saisir/générer : " & xSAISI _
         & "\par\tab\cf1 Droit validation     : \cf13 " & xVALID _
         & "\par\tab\cf1 Droit d'envoi        : " & xENVOI _
         & "\par\tab\cf1 Droit type MT        : \cf13 " & Trim(xMECON) _
         & "\par\tab\cf1 Montant maximum      : \cf2 " & xMONTA _
         & "\par\tab\cf1 Mt exonéré 2ème valid: \cf2 " & xMONT2 _
         & xRTF_1 _
         & "\par\cf8\b ______________________________________________________________________ \b0\par" _


'xRTF = xRTF & X
'___________________________________________________________________________________________________


cmdSSISAW_Detail_txtRTF = xRTF

End Function




Public Function cmdSSIDIV_Detail_txtRTF(lhighlight As Integer) As String
Dim xRTF As String, xTLNK As String, X As String
'____________________________________________________________________________________________
If rtfYSSIDIV0.SSIDIVTLNK <> 0 Then
         xYSSITXT0.SSITXTNAT = rtfYSSIDIV0.SSIDIVNAT
         'xYSSITXT0.SSITXTUIDN = rtfYSSIDIV0.SSIDIVUIDN
         xYSSITXT0.SSITXTDIDX = "DIV"
         xYSSITXT0.SSITXTUIDX = rtfYSSIDIV0.SSIDIVUIDX
         xYSSITXT0.SSITXTUIDD = rtfYSSIDIV0.SSIDIVUIDD
         xYSSITXT0.SSITXTTLNK = rtfYSSIDIV0.SSIDIVTLNK
    xTLNK = cmdSSITXT_Detail_txtRTF
End If

 xRTF = "\fs18\cf1\ul\highlight12 Compte DIV " & rtfYSSIDIV0.SSIDIVDIDK & " :\cf13\highlight" & lhighlight & "\b  " & Trim(rtfYSSIDIV0.SSIDIVUIDX) & "  \b0\ulnone\highlight0" _
         & "\par " _
         & "\par\tab\cf8 Actif        : \cf7 " & cmdSSISTAK_Detail_txtRTF(rtfYSSIDIV0.SSIDIVSTAK, 0) _
         & "\par\tab\cf8 Id SSI       : \cf7 " & rtfYSSIDIV0.SSIDIVUIDD _
         & "\par\tab\cf8 Profil SSI   : \cf13 " & rtfYSSIDIV0.SSIDIVUIDX _
         & "\par\tab\cf8 Profil \cf2" & rtfYSSIDIV0.SSIDIVDIDK & "   : " & Trim(rtfYSSIDIV0.SSIDIVPRFX) _
         & "\par\tab\cf8 Nom          : \cf13 " & rtfYSSIDIV0.SSIDIVUNOM _
         & "\par\tab\cf8 Conforme ?   : \cf7 " & cmdSSIPRFK_Detail_txtRTF(rtfYSSIDIV0.SSIDIVPRFK, 0, 0) _
         & "\par\tab\cf8 Commentaire °: \cf7 " & rtfYSSIDIV0.SSIDIVTLNK _
         & "\par\tab\cf8 Evénement    : \cf7 " & rtfYSSIDIV0.SSIDIVYFCT _
         & "  (v" & rtfYSSIDIV0.SSIDIVYVER & ")" _
         & "\par\tab\cf8 màj par      : \cf7 " & rtfYSSIDIV0.SSIDIVYUSR _
         & " le " & dateImp10_S(rtfYSSIDIV0.SSIDIVYAMJ) & " " & timeImp8(rtfYSSIDIV0.SSIDIVYHMS) _
         & xTLNK _
         & "\par\cf8\b ______________________________________________________________________ \b0\par"
         
If Trim(rtfYSSIDIV0.SSIDIVDIDK) = "TEREN" Then
    X = rtfYSSIDIV0.SSIDIVINFO
    X = "\par\tab\cf8 Identifiant             : \cf13 " & X
    X = Replace(X, vbTab, "\par\tab\cf8 Autorisation            : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 Jours fériés            : \cf13 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 date début              : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 date fin                : \cf13 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 heure fin               : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 zone comptage           : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 groupe lecteurs 1 et 2  : \cf2 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 groupe horaires 1 et 2  : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 comptage lecteur 1 et 3 : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 Horaire individuel      : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 LMMJVSD                 : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 ascenseurs              : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 lecteurs                : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 badge 1                 : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 badge 2                 : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 badge 3                 : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 badge 4                 : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 badge 5                 : \cf7 ", 1, 1)
    X = Replace(X, vbTab, "\par\tab\cf8 badge 6                 : \cf7 ", 1, 1)
Else
    If Trim(rtfYSSIDIV0.SSIDIVDIDK) = "UGM" Then
        X = rtfYSSIDIV0.SSIDIVINFO
        X = "\par\tab\cf8 Identifiant    : \cf13 " & X
        X = Replace(X, "|", "\par\tab\cf8 Date début     : \cf7 ", 1, 1)
        X = Replace(X, "|", "\par\tab\cf8 Date fin       : \cf13 ", 1, 1)
        X = Replace(X, "|", "\par\tab\cf8 Groupe horaire : \cf7 ", 1, 1)
        X = Replace(X, "|", "\par\tab\cf8 Badge1         : \cf13 ", 1, 1)
        X = Replace(X, "|", "\par\tab\cf8 Badge2         : \cf7 ", 1, 1)
        X = Replace(X, "|", "\par\tab\cf8 Badge3         : \cf7 ", 1, 1)
        X = Replace(X, "|", "\par\tab\cf8 Badge4         : \cf7 ", 1, 1)
        X = Replace(X, "|", "\par\tab\cf8 Badge5         : \cf7 ", 1, 1)
        X = Replace(X, "|", "\par\tab\cf8 Badge6         : \cf7 ", 1, 1)
        X = Replace(X, "|", "\par\tab\cf8 Lecteurs       : \cf2 ", 1, 1)
    Else
        X = "\par\tab\cf8 Info : \cf13 " & Replace(Trim(rtfYSSIDIV0.SSIDIVINFO), vbCrLf, "\highlight0\par\tab\tab ")
    End If
End If
xRTF = Replace(xRTF, "{", " ") & "\par\cf13 " & X
'___________________________________________________________________________________________________

If cmdSelect_SQL_K = "2_D" Then
    xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight11 Utilisateurs actifs :\highlight0\ulnone"
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
         & " where SSIDOMNAT = ' '  and SSIDOMDIDX = 'DIV' and SSIDOMPRFX = '" & rtfYSSIDIV0.SSIDIVUIDX & "'" _
         & " and SSIDOMPRFK <> 'X' and SSIDIVNAT = ' ' and SSIDIVUIDD = SSIDOMUIDD  and SSIDIVUIDX = SSIDOMUIDX order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        'kNAT_U = kNAT_U + 1
        
        xRTF = xRTF & "\par\tab\fs16\cf13\b " & rsSab("SSIDOMUIDX") _
                    & "\b0\cf2  : " & Trim(rsSab("SSIDIVUNOM"))
        rsSab.MoveNext
    Loop
    xRTF = xRTF & "\par\par\fs16\ul\cf1\highlight9 Utilisateurs INACTIFS :\highlight0\ulnone"
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIDIV0 " _
         & " where SSIDOMNAT = ' '  and SSIDOMDIDX = 'DIV' and SSIDOMPRFX = '" & rtfYSSIDIV0.SSIDIVUIDX & "'" _
         & " and SSIDOMPRFK = 'X' and SSIDIVNAT = ' ' and SSIDIVUIDD = SSIDOMUIDD   and SSIDIVUIDX = SSIDOMUIDX order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        'kNAT_U = kNAT_U + 1
        
        xRTF = xRTF & "\par\tab\fs16\cf13\b " & rsSab("SSIDOMUIDX") _
                    & "\b0\cf2  : " & Trim(rsSab("SSIDIVUNOM"))
        rsSab.MoveNext
    Loop
End If
xRTF = xRTF & "\par\cf8 _______________________________________________________________________\par"

cmdSSIDIV_Detail_txtRTF = xRTF

End Function


Public Function cmdSSIDOM_Detail_Load(lFct As String) As String
Dim xRTF As String, whighlight As Integer, wSSIDOMDECH As String, wSSIDOMPRFD As String
'____________________________________________________________________________________________
whighlight = 11
If lFct = "YSSIDOM*" Then
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOMH " _
         & " where SSIDOMNAT = ' ' and SSIDOMUIDN = " & xYSSIDOM0.SSIDOMUIDN _
         & " and SSIDOMDIDX = '" & xYSSIDOM0.SSIDOMDIDX & "' and SSIDOMUIDX = '" & xYSSIDOM0.SSIDOMUIDX & "'" _
         & " and SSIDOMUIDD = " & xYSSIDOM0.SSIDOMUIDD & " And SSIDOMYVER = " & xYSSIDOM0.SSIDOMYVER
    Set rsSab = cnsab.Execute(X)
    If Not rsSab.EOF Then
        whighlight = 9
    Else
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
         & " where SSIDOMNAT = ' ' and SSIDOMUIDN = " & xYSSIDOM0.SSIDOMUIDN _
         & " and SSIDOMDIDX = '" & xYSSIDOM0.SSIDOMDIDX & "' and SSIDOMUIDX = '" & xYSSIDOM0.SSIDOMUIDX & "'" _
         & " and SSIDOMUIDD = " & xYSSIDOM0.SSIDOMUIDD
        Set rsSab = cnsab.Execute(X)
    End If
Else
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
         & " where SSIDOMNAT = ' ' and SSIDOMUIDN = " & xYSSIDOM0.SSIDOMUIDN _
         & " and SSIDOMDIDX = '" & xYSSIDOM0.SSIDOMDIDX & "' and SSIDOMUIDX = '" & xYSSIDOM0.SSIDOMUIDX & "'" _
         & " and SSIDOMUIDD = " & xYSSIDOM0.SSIDOMUIDD
    Set rsSab = cnsab.Execute(X)
End If

If Not rsSab.EOF Then
    Call rsYSSIDOM0_GetBuffer(rsSab, xYSSIDOM0)


    txtRTF.TextRTF = VB_RTF_Modèle
    
    txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", cmdSSIDOM_Detail_txtRTF(whighlight))
    
    Call txtRTF_Visible

    cmdSSIDOM_Detail_Load = Null
End If
End Function

Public Function cmdSSIDOM_Detail_txtRTF(lhighlight As Integer) As String
Dim xRTF As String, xTLNK As String
'____________________________________________________________________________________________

If xYSSIDOM0.SSIDOMTLNK <> 0 Then
         xYSSITXT0.SSITXTNAT = xYSSIDOM0.SSIDOMNAT
         xYSSITXT0.SSITXTUIDN = xYSSIDOM0.SSIDOMUIDN
         xYSSITXT0.SSITXTDIDX = xYSSIDOM0.SSIDOMDIDX
         xYSSITXT0.SSITXTUIDX = xYSSIDOM0.SSIDOMUIDX
         xYSSITXT0.SSITXTUIDD = xYSSIDOM0.SSIDOMUIDD
         xYSSITXT0.SSITXTTLNK = xYSSIDOM0.SSIDOMTLNK
    xTLNK = cmdSSITXT_Detail_txtRTF
End If

 xRTF = "\fs18\cf1\ul\highlight12 Domaine " & Trim(xYSSIDOM0.SSIDOMDIDX) & " :\cf13\highlight" & lhighlight & "\b " & xYSSIDOM0.SSIDOMUIDX & "  \b0\ulnone\highlight0" _
         & "\par " _
         & "\par\tab\cf8 Actif        : \cf7 " & cmdSSISTAK_Detail_txtRTF(xYSSIDOM0.SSIDOMSTAK, xYSSIDOM0.SSIDOMDECH) _
         & "\par\tab\cf8 Id SSI       : \cf7 " & xYSSIDOM0.SSIDOMUIDN & " - " & xYSSIDOM0.SSIDOMDIDX & " - " & xYSSIDOM0.SSIDOMUIDD _
         & "\par\tab\cf8 Nom          : \cf7 " & xYSSIDOM0.SSIDOMUIDX _
         & "\par\tab\cf8 Profil       : \cf13 " & Trim(xYSSIDOM0.SSIDOMPRFX) _
         & "\par\tab\cf8 Conforme ?   : \cf7 " & cmdSSIPRFK_Detail_txtRTF(xYSSIDOM0.SSIDOMPRFK, xYSSIDOM0.SSIDOMPRFD, xYSSIDOM0.SSIDOMPRFH) _
         & "\par\tab\cf8 Commentaire °: \cf7 " & xYSSIDOM0.SSIDOMTLNK _
         & "\par\tab\cf8 Evénement    : \cf7 " & xYSSIDOM0.SSIDOMYFCT _
         & "  (v" & xYSSIDOM0.SSIDOMYVER & ")" _
         & "\par\tab\cf8 màj par      : \cf7 " & xYSSIDOM0.SSIDOMYUSR _
         & " le " & dateImp10_S(xYSSIDOM0.SSIDOMYAMJ) & " " & timeImp8(xYSSIDOM0.SSIDOMYHMS) _
         & xTLNK _
         & "\highlight0\par\cf8 _______________________________________________________________________\par"


cmdSSIDOM_Detail_txtRTF = xRTF
End Function




Public Function cmdSSISAB_Display_ZMNUHLA0() As String
Dim X0 As String, xDBD As String, xFID As String, xCRE As String, xVAL As String, xMOD As String, xFIN0 As String

If Mid$(xYSSISAB0.SSISABINFO, 1, 1) = "1" Then
    X0 = "OUI"
Else
    X0 = "NON"
End If
If Val(Mid$(xYSSISAB0.SSISABINFO, 1, 7)) > 0 Then
    xDBD = dateImp10_S(Mid$(xYSSISAB0.SSISABINFO, 2, 7) + 19000000) & " " & timeImp8(Mid$(xYSSISAB0.SSISABINFO, 9, 6))
End If
If Val(Mid$(xYSSISAB0.SSISABINFO, 15, 7)) > 0 Then
    xFID = dateImp10_S(Mid$(xYSSISAB0.SSISABINFO, 15, 7) + 19000000) & " " & timeImp8(Mid$(xYSSISAB0.SSISABINFO, 22, 6))
End If
If Val(Mid$(xYSSISAB0.SSISABINFO, 32, 7)) > 0 Then
    xCRE = dateImp10_S(Mid$(xYSSISAB0.SSISABINFO, 32, 7) + 19000000) & " " & timeImp8(Mid$(xYSSISAB0.SSISABINFO, 39, 6)) & " par " & arrMNURUTCUT(Val(Mid$(xYSSISAB0.SSISABINFO, 28, 4)))
End If
If Val(Mid$(xYSSISAB0.SSISABINFO, 49, 7)) > 0 Then
    xMOD = dateImp10_S(Mid$(xYSSISAB0.SSISABINFO, 49, 7) + 19000000) & " " & timeImp8(Mid$(xYSSISAB0.SSISABINFO, 56, 6)) & " par " & arrMNURUTCUT(Val(Mid$(xYSSISAB0.SSISABINFO, 45, 4)))
End If
If Val(Mid$(xYSSISAB0.SSISABINFO, 66, 7)) > 0 Then
    xVAL = dateImp10_S(Mid$(xYSSISAB0.SSISABINFO, 66, 7) + 19000000) & " " & timeImp8(Mid$(xYSSISAB0.SSISABINFO, 73, 6)) & " par " & arrMNURUTCUT(Val(Mid$(xYSSISAB0.SSISABINFO, 62, 4)))
End If
If Val(Mid$(xYSSISAB0.SSISABINFO, 79, 7)) > 0 Then
    xFIN0 = dateImp10_S(Mid$(xYSSISAB0.SSISABINFO, 79, 7) + 19000000) & " " & timeImp8(Mid$(xYSSISAB0.SSISABINFO, 86, 6))
End If

cmdSSISAB_Display_ZMNUHLA0 = "\fs18\cf13\highlight12\ul\b ZMNUHLA0 lot : " & xYSSISAB0.SSISABULOT & "\b0\ulnone\highlight0\cf1 " & "\cf2  => " _
        & xYSSISAB0.SSISABUNOM _
        & "\par " _
        & "\par\tab\cf8 Actif        : \cf7 " & cmdSSISTAK_Detail_txtRTF(xYSSISAB0.SSISABSTAK, 0) _
        & "\par\tab\cf8 Id SSI       : \cf7 " & xYSSISAB0.SSISABUIDD _
        & "\par\tab\cf8 Evénement    : \cf7 " & xYSSISAB0.SSISABYFCT _
        & "  (v" & xYSSISAB0.SSISABYVER & ")" _
        & "\par\tab\cf8 màj par      : \cf7 " & xYSSISAB0.SSISABYUSR _
        & " le " & dateImp10_S(xYSSISAB0.SSISABYAMJ) & " " & timeImp8(xYSSISAB0.SSISABYHMS) _
        & "\par" _
        & "\par\tab\cf1 Validé       : \cf13 " & X0 _
        & "\par\tab\cf1 date de début: \cf13 " & xDBD _
        & "\par\tab\cf1 date de fin  : \cf10 " & xFID _
        & "\par\tab\cf1 saisi le     : \cf13 " & xCRE _
        & "\par\tab\cf1 validé le    : \cf13 " & xVAL _
        & "\par\tab\cf1 modifié le   : \cf13 " & xMOD _
        & "\par\tab\cf1 fin précédent:\cf13 " & xFIN0 _
        & "\par\cf8 _______________________________________________________________________\par"


End Function


Public Function cmdSSISAB_Display_ZMNUHLB0() As String
Dim X0 As String, xDBD As String, xFID As String, xCRE As String
If Mid$(xYSSISAB0.SSISABINFO, 1, 1) = "1" Then
    X0 = "OUI"
Else
    X0 = "NON"
End If
If Val(Mid$(xYSSISAB0.SSISABINFO, 1, 7)) > 0 Then
    xDBD = dateImp10_S(Mid$(xYSSISAB0.SSISABINFO, 2, 7) + 19000000) & " " & timeImp8(Mid$(xYSSISAB0.SSISABINFO, 9, 6))
End If
If Val(Mid$(xYSSISAB0.SSISABINFO, 15, 7)) > 0 Then
    xFID = dateImp10_S(Mid$(xYSSISAB0.SSISABINFO, 15, 7) + 19000000) & " " & timeImp8(Mid$(xYSSISAB0.SSISABINFO, 22, 6))
End If
If Val(Mid$(xYSSISAB0.SSISABINFO, 32, 7)) > 0 Then
    xCRE = dateImp10_S(Mid$(xYSSISAB0.SSISABINFO, 32, 7) + 19000000) & " " & timeImp8(Mid$(xYSSISAB0.SSISABINFO, 39, 6)) & " par " & arrMNURUTCUT(Val(Mid$(xYSSISAB0.SSISABINFO, 28, 4)))
End If

cmdSSISAB_Display_ZMNUHLB0 = "\fs18\cf13\highlight12\ul\b ZMNUHLB0 lot : " & xYSSISAB0.SSISABULOT & "\b0\ulnone\highlight0\cf1 " & "\cf2  => " _
        & xYSSISAB0.SSISABUNOM _
        & "\par " _
        & "\par\tab\cf8 Actif        : \cf7 " & cmdSSISTAK_Detail_txtRTF(xYSSISAB0.SSISABSTAK, 0) _
        & "\par\tab\cf8 Id SSI       : \cf7 " & xYSSISAB0.SSISABUIDD _
        & "\par\tab\cf8 Evénement    : \cf7 " & xYSSISAB0.SSISABYFCT _
        & "  (v" & xYSSISAB0.SSISABYVER & ")" _
        & "\par\tab\cf8 màj par      : \cf7 " & xYSSISAB0.SSISABYUSR _
        & " le " & dateImp10_S(xYSSISAB0.SSISABYAMJ) & " " & timeImp8(xYSSISAB0.SSISABYHMS) _
        & "\par" _
        & "\par\tab\cf1 Validé       : \cf13 " & X0 _
        & "\par\tab\cf1 date de début: \cf13 " & xDBD _
        & "\par\tab\cf1 date de fin  : \cf10 " & xFID _
        & "\par\tab\cf1 saisi le     : \cf13 " & xCRE _
        & "\par\cf8 _______________________________________________________________________\par"


End Function

Public Function cmdSSISAB_Display_ZMNUUTP0() As String
Dim X As String, xRTF As String, K1 As Integer
X = Trim(rsSab("SSISABINFO"))
For K1 = 1 To 99
    Select Case Mid$(X, K1, 1)
        Case Is = "1"
            xRTF = xRTF & "\par\tab\cf13 " & arrMNURCLABR(K1) & " : \cf7 consultation"
        Case Is = "2"
            xRTF = xRTF & "\par\tab\cf13 " & arrMNURCLABR(K1) & " : \cf14 consultation + mise à jour autorisée"
         Case Is = "3"
            xRTF = xRTF & "\par\tab\cf13 " & arrMNURCLABR(K1) & " : \cf10 mise à jour autorisée sans consultation"
    End Select
Next K1
cmdSSISAB_Display_ZMNUUTP0 = "\fs18\cf13\highlight12\ul\b lot - groupe : " & xYSSISAB0.SSISABULOT & "\cf13 - " & Mid$(rsSab("SSISABUIDX"), 1, 14) _
                           & "\b0\ulnone\highlight0\cf1 " & "\cf2  => (droits des données communes)" _
        & xYSSISAB0.SSISABUNOM _
        & "\par " _
        & "\par\tab\cf8 Actif        : \cf7 " & cmdSSISTAK_Detail_txtRTF(xYSSISAB0.SSISABSTAK, 0) _
        & "\par\tab\cf8 Evénement    : \cf7 " & xYSSISAB0.SSISABYFCT _
        & "  (v" & xYSSISAB0.SSISABYVER & ")" _
        & "\par\tab\cf8 màj par      : \cf7 " & xYSSISAB0.SSISABYUSR _
        & " le " & dateImp10_S(xYSSISAB0.SSISABYAMJ) & " " & timeImp8(xYSSISAB0.SSISABYHMS) _
        & "\par" _
        & xRTF _
        & "\par\cf8 _______________________________________________________________________\par"


End Function
Public Function cmdSSISAB_Display_ZMNUUTO0() As String
Dim X As String, xRTF As String, K1 As Integer
X = Trim(rsSab("SSISABINFO"))
For K1 = 1 To Len(X) Step 6
    If Mid$(X, K1 + 4, 1) = "O" Then
        xRTF = xRTF & "\par\tab\cf13 " & Mid$(X, K1, 4) & "         : \cf14 mise à jour autorisée"
    Else
        xRTF = xRTF & "\par\tab\cf13 " & Mid$(X, K1, 4) & "         : \cf7 consultation"
    End If
Next K1
cmdSSISAB_Display_ZMNUUTO0 = "\fs18\cf13\highlight12\ul\b lot - groupe : " & xYSSISAB0.SSISABULOT & "\cf13 - " & Mid$(rsSab("SSISABUIDX"), 1, 14) _
                           & "\b0\ulnone\highlight0\cf1 " & "\cf2  => (droits des données opérations)" _
        & xYSSISAB0.SSISABUNOM _
        & "\par " _
        & "\par\tab\cf8 Actif        : \cf7 " & cmdSSISTAK_Detail_txtRTF(xYSSISAB0.SSISABSTAK, 0) _
        & "\par\tab\cf8 Evénement    : \cf7 " & xYSSISAB0.SSISABYFCT _
        & "  (v" & xYSSISAB0.SSISABYVER & ")" _
        & "\par\tab\cf8 màj par      : \cf7 " & xYSSISAB0.SSISABYUSR _
        & " le " & dateImp10_S(xYSSISAB0.SSISABYAMJ) & " " & timeImp8(xYSSISAB0.SSISABYHMS) _
        & "\par" _
        & xRTF _
        & "\par\cf8 _______________________________________________________________________\par"


End Function
Public Function cmdSSISAB_Display_ZMNUMEN0() As String
Dim X As String, xRTF As String, K1 As Integer
On Error Resume Next
If arrMNUGRPNOM_Nb = 0 Then paramMNUGRPNOM_Load

X = xYSSISAB0.SSISABINFO
For K1 = 1 To arrMNUGRPNOM_Nb
    If Mid$(X, K1, 1) = "O" Then
        xRTF = xRTF & "\par\tab\cf13 " & arrMNUGRPNOM(K1)
    End If
Next K1
cmdSSISAB_Display_ZMNUMEN0 = "\fs18\cf13\highlight12\ul\b lot " & xYSSISAB0.SSISABULOT & "\cf13  -  option : " & xYSSISAB0.SSISABUIDX _
                           & "\b0\ulnone\highlight0\cf1 " & "\cf2  => " & Trim(xYSSISAB0.SSISABUNOM) _
        & "\par " _
        & "\par\tab\cf8 Actif        : \cf7 " & cmdSSISTAK_Detail_txtRTF(xYSSISAB0.SSISABSTAK, 0) _
        & "\par\tab\cf8 Evénement    : \cf7 " & xYSSISAB0.SSISABYFCT _
        & "  (v" & xYSSISAB0.SSISABYVER & ")" _
        & "\par\tab\cf8 màj par      : \cf7 " & xYSSISAB0.SSISABYUSR _
        & " le " & dateImp10_S(xYSSISAB0.SSISABYAMJ) & " " & timeImp8(xYSSISAB0.SSISABYHMS) _
        & "\par" _
        & xRTF _
        & "\par\cf8 _______________________________________________________________________\par"


End Function



Public Sub paramMNUGRPNOM_Load()
X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0" _
     & " where SSISABNAT = '2' order by SSISABUIDD desc"

Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    If arrMNUGRPNOM_Nb = 0 Then
        arrMNUGRPNOM_Nb = rsSab("SSISABUIDD")
        ReDim arrMNUGRPNOM(arrMNUGRPNOM_Nb + 1)
    End If
    arrMNUGRPNOM(rsSab("SSISABUIDD")) = Mid$(rsSab("SSISABUIDX"), 1, 12) & " : " & Trim(rsSab("SSISABUNOM"))
    rsSab.MoveNext
Loop

For I = 1 To arrMNUGRPNOM_Nb
    If arrMNUGRPNOM(I) = "" Then arrMNUGRPNOM(I) = "?grp " & I
Next I

End Sub

Public Function cmdSSISAA_Detail_txtRTF_SSISAAINFO(lTxt, lX As String) As String
Dim X As String, K As Integer
X = Replace(lTxt, vbCrLf, "\par\tab\cf1 ")
'Approved
K = InStr(X, "CCY/Amount") + 15
If K > 15 Then
    Dim blnExit As Boolean, K1 As Integer
    K1 = InStr(K, X, lX) - 35
    Do
        K = InStr(K, X, "cf1")
        If K = 0 Or K > K1 Then
            blnExit = True
        Else
            Mid$(X, K, 3) = "cf*"
        End If
        
    Loop Until blnExit
    X = Replace(X, "\cf*", "\tab\cf13       ")
End If
X = Replace(X, "Approved", "\highlight11 Approved \highlight0 ")
X = Replace(X, "Unapproved", "\highlight9 Unapproved \highlight0 ")
cmdSSISAA_Detail_txtRTF_SSISAAINFO = Replace(X, lX, lX & "\cf13   ")
End Function

Public Sub cmdSelect_SQL_9_SAA_Profil_Inactif()
Dim xSQL As String
Dim rsSab_X As New ADODB.Recordset

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
     & " where SSISAANAT = '$' and SSISAASTAK = ' '" _
     & "and SSISAAUIDX not in (select distinct SSIDOMPRFX from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
     & " where SSIDOMNAT = ' ' and SSIDOMDIDX = 'SAA'" _
     & " and SSIDOMSTAK = ' ')"

Set rsSab_X = cnsab.Execute(xSQL)
Do Until rsSab_X.EOF
    mYSSISAA0_Update = "Update"
    Call rsYSSISAA0_GetBuffer(rsSab_X, oldYSSISAA0)
    newYSSISAA0 = oldYSSISAA0
    newYSSISAA0.SSISAAYFCT = "MOD"
    newYSSISAA0.SSISAASTAK = "N"
    
    Call cmdSSIJRN_SAA("<X:profil inutilisé>")
    newYSSISAA0.SSISAAYAMJ = DSys
    newYSSISAA0.SSISAAYHMS = time_Hms
    newYSSISAA0.SSISAAYUSR = usrName_UCase
    Call cmdUpdate
    
    rsSab_X.MoveNext
Loop


'xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
'     & " where SSISAANAT = '$' and SSISAASTAK = 'N'" _
'     & "and SSISAAUIDX in (select distinct SSIDOMPRFX from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
'     & " where SSIDOMNAT = ' ' and SSIDOMDIDX = 'SAA'" _
'     & " and SSIDOMSTAK = ' ')"

'$JPL 2014-11-19

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
     & " where SSISAANAT = '$' and SSISAASTAK = 'N'" _
     & "and SSISAAUIDX in (select distinct SSISAAPRFX from " & paramIBM_Library_SABSPE & ".YSSISAA0 " _
     & " where SSISAANAT = ' ' and SSISAASTAK <> 'N')"

Set rsSab_X = cnsab.Execute(xSQL)
Do Until rsSab_X.EOF
    mYSSISAA0_Update = "Update"
    Call rsYSSISAA0_GetBuffer(rsSab_X, oldYSSISAA0)
    newYSSISAA0 = oldYSSISAA0
    newYSSISAA0.SSISAAYFCT = "MOD"
    newYSSISAA0.SSISAASTAK = " "
    
    Call cmdSSIJRN_SAA("<X:profil réactivé>")
    newYSSISAA0.SSISAAYAMJ = DSys
    newYSSISAA0.SSISAAYHMS = time_Hms
    newYSSISAA0.SSISAAYUSR = usrName_UCase
    Call cmdUpdate
    
    rsSab_X.MoveNext
Loop


End Sub

Public Sub fgProfil_Display()

End Sub

Public Function cmdSSIPRFK_Detail_txtRTF(lPRFK As String, lPRFD As Long, lPRFH As Long) As String
Dim X As String

If lPRFD > 0 Then
    X = "\cf8 contrôlé le \cf13 " & dateImp10_S(lPRFD) & " " & timeImp8(lPRFH)
End If

Select Case lPRFK
    Case " ": cmdSSIPRFK_Detail_txtRTF = "\cf13\highlight11 Oui\highlight0   " & X
    Case "N": cmdSSIPRFK_Detail_txtRTF = "\cf10\highlight12 Non\highlight0   " & X
    Case "?": cmdSSIPRFK_Detail_txtRTF = "\cf1\highlight14 ? en attente\highlight0   " & X
    Case "!": cmdSSIPRFK_Detail_txtRTF = "\cf1\highlight14 ! échéance\highlight0   " & X
    Case "X": cmdSSIPRFK_Detail_txtRTF = "\cf1\highlight9 X exit_grp\highlight0   " & X
    Case Else: cmdSSIPRFK_Detail_txtRTF = "\cf2\highlight12 " & lPRFK & "\highlight0   " & X
End Select


End Function
Public Function cmdSSISTAK_Detail_txtRTF(lSTAK As String, lDECH As Long) As String
Dim X As String

If lDECH > 0 Then
    If lDECH > DSys Then
        X = "\cf8 Echéance \cf13 " & dateImp10_S(lDECH)
    Else
        X = "\cf8 Echéance \cf10 " & dateImp10_S(lDECH)
    End If
    
End If

Select Case lSTAK
    Case " ", "": cmdSSISTAK_Detail_txtRTF = "\cf13\highlight11 Oui\highlight0   " & X
    Case "N": cmdSSISTAK_Detail_txtRTF = "\cf10\highlight12 Non\highlight0   " & X
    Case Else: cmdSSISTAK_Detail_txtRTF = "\cf2\highlight12 " & lSTAK & "\highlight0   " & X
End Select


End Function


Public Sub fgProfil_Row_Click()
        Call fgProfil_Color(fgProfil_RowClick, MouseMoveUsr.BackColor, fgProfil_ColorClick)
        Select Case cmdSelect_SQL_K
            Case "1", "2"
                fgProfil.Enabled = False
                Call rsYSSIDOM0_Init(oldYSSIDOM0)
                oldYSSIDOM0.SSIDOMNAT = mSSIUSRNAT
                oldYSSIDOM0.SSIDOMUIDN = oldYSSIUSR0.SSIUSRUIDN
                oldYSSIDOM0.SSIDOMSTAK = oldYSSIUSR0.SSIUSRSTAK
                oldYSSIDOM0.SSIDOMDECH = oldYSSIUSR0.SSIUSRDECH
                oldYSSIDOM0.SSIDOMDIDX = mSSIDOMDIDX
                fgProfil.Col = 1
                oldYSSIDOM0.SSIDOMPRFX = Trim(fgProfil.Text)
                oldYSSIDOM0.SSIDOMPRFK = "?"
                oldYSSIDOM0.SSIDOMUNIT = oldYSSIUSR0.SSIUSRUNIT
                
                Call rsYSSITXT0_Init(oldYSSITXT0_DOM)
                oldYSSITXT0_DOM.SSITXTNAT = mSSIUSRNAT
                
                Select Case mSSIDOMDIDX
                    Case "IBM"
                        fgProfil.Col = 0: oldYSSIIBM0.SSIIBMUIDD = Val(fgProfil.Text)
                        oldYSSIIBM0.SSIIBMNAT = "$"
                        Call cmdSSIIBM_Detail_Display("", "")
                        oldYSSIIBM0 = xYSSIIBM0
                        fraYSSIDOM0_Display
                    Case "SAA"
                        fgProfil.Col = 1: prfYSSISAA0.SSISAAUIDX = Trim(fgProfil.Text)
                        prfYSSISAA0.SSISAANAT = "$"
                        usrYSSISAA0.SSISAAUIDX = ""
                        Call cmdSSISAA_Detail_Display("")
                        oldYSSISAA0 = xYSSISAA0
                        fraYSSIDOM0_Display
                    Case "SAB"
                        fgProfil.Col = 1: prfYSSISAB0.SSISABUIDX = Trim(fgProfil.Text)
                        prfYSSISAB0.SSISABNAT = "$"
                        usrYSSISAB0.SSISABUIDX = ""
                        Call cmdSSISAB_Detail_Display("")
                        oldYSSISAB0 = xYSSISAB0
                        fraYSSIDOM0_Display
                     Case "WIN"
                        fgProfil.Col = 1: rtfYSSIWIN0.SSIWINUIDX = Trim(fgProfil.Text)
                        fgProfil.Col = 3: rtfYSSIWIN0.SSIWINGUID = Trim(fgProfil.Text)
                        rtfYSSIWIN0.SSIWINNAT = "$"
                        'usrYSSIWIN0.SSIWINUIDX = ""
                        Call cmdSSIWIN_Detail_Display("YSSIWIN0")
                        oldYSSIWIN0 = rtfYSSIWIN0
                        fraYSSIDOM0_Display
                      Case "DIV"
                        fgProfil.Col = 0: rtfYSSIDIV0.SSIDIVUIDD = Val(fgProfil.Text)
                        fgProfil.Col = 1: rtfYSSIDIV0.SSIDIVUIDX = Trim(fgProfil.Text)
                        rtfYSSIDIV0.SSIDIVNAT = "$"
                        Call cmdSSIDIV_Detail_Display("YSSIDIV0")
                        'usrYSSIDIV0 = rtfYSSIDIV0
                        'oldYSSIDIV0 = rtfYSSIDIV0
                        prfYSSIDIV0 = rtfYSSIDIV0
                        fraYSSIDOM0_Display
                        Select Case prfYSSIDIV0.SSIDIVDIDK
                            Case "TEREN", "SG", "UGM": cmdProfil_Update_DIV.Visible = arrHab(5)
                            Case Else: cmdProfil_Update_DIV.Visible = arrHab(2)
                        End Select
                      Case "MEL"
                        fgProfil.Col = 0: rtfYSSIMEL0.SSIMELUIDD = Val(fgProfil.Text)
                        fgProfil.Col = 1: rtfYSSIMEL0.SSIMELUIDX = Trim(fgProfil.Text)
                        rtfYSSIMEL0.SSIMELNAT = "$"
                        Call cmdSSIMEL_Detail_Display("YSSIMEL0")
                        prfYSSIMEL0 = rtfYSSIMEL0
                        fraYSSIDOM0_Display
                      Case "TIC"
                        fgProfil.Col = 0: rtfYSSITIC0.SSITICUIDD = Val(fgProfil.Text)
                        fgProfil.Col = 1: rtfYSSITIC0.SSITICUIDX = Trim(fgProfil.Text)
                        rtfYSSITIC0.SSITICNAT = "$"
                        Call cmdSSITIC_Detail_Display("YSSITIC0")
                        prfYSSITIC0 = rtfYSSITIC0
                        fraYSSIDOM0_Display
                        
             End Select
                        
            Select Case mSSIDOMDIDX
                Case "MEL": cmdProfil_Update.Visible = False
                Case "DIV": cmdProfil_Update.Visible = arrHab(2) Or arrHab(5)
                Case Else: cmdProfil_Update.Visible = arrHab(2)
             End Select
            Case "2_D"
                Select Case mSSIDOMDIDX
                    Case "IBM"
                        fgProfil.Col = 0: oldYSSIIBM0.SSIIBMUIDD = Val(fgProfil.Text)
                        oldYSSIIBM0.SSIIBMNAT = "$"
                        Call cmdSSIIBM_Detail_Display("", "")
                    Case "SAA"
                        usrYSSISAA0.SSISAAUIDX = ""
                        fgProfil.Col = 1: prfYSSISAA0.SSISAAUIDX = Trim(fgProfil.Text)
                        prfYSSISAA0.SSISAANAT = "$"
                        Call cmdSSISAA_Detail_Display("")
                    Case "SAB"
                        'usrYSSISAB0.SSISABUIDX = ""
                        fgProfil.Col = 1: usrYSSISAB0.SSISABUIDX = Trim(fgProfil.Text)
                        usrYSSISAB0.SSISABNAT = "$"
                        Call cmdSSISAB_Detail_Display("")
                        If currentAction = "cmdProfil_Excel_Click" Then
                            lstW.AddItem usrYSSISAB0.SSISABUIDX
                        End If
                    Case "WIN"
                        rtfYSSIWIN0.SSIWINUIDX = ""
                        fgProfil.Col = 3: rtfYSSIWIN0.SSIWINGUID = Trim(fgProfil.Text)
                        rtfYSSIWIN0.SSIWINNAT = "$"
                        Call cmdSSIWIN_Detail_Display("YSSIWIN0")
                    Case "DIV"
                        rtfYSSIDIV0.SSIDIVUIDX = ""
                        fgProfil.Col = 0: rtfYSSIDIV0.SSIDIVUIDD = Val(fgProfil.Text)
                        fgProfil.Col = 1: rtfYSSIDIV0.SSIDIVUIDX = Trim(fgProfil.Text)
                        rtfYSSIDIV0.SSIDIVNAT = "$"
                        Call cmdSSIDIV_Detail_Display("YSSIDIV0")
                        
                        cmdProfil_Change.Visible = arrHab(2)
                        If Not arrHab(2) Then cmdProfil_Change.Visible = arrHab(5)
                    Case "MEL"
                        rtfYSSIMEL0.SSIMELUIDX = ""
                        fgProfil.Col = 1: rtfYSSIMEL0.SSIMELUIDX = Trim(fgProfil.Text)
                        rtfYSSIMEL0.SSIMELNAT = "$"
                        Call cmdSSIMEL_Detail_Display("YSSIMEL0")
                    Case "TIC"
                        rtfYSSITIC0.SSITICUIDX = ""
                        fgProfil.Col = 1: rtfYSSITIC0.SSITICUIDX = Trim(fgProfil.Text)
                        rtfYSSITIC0.SSITICNAT = "$"
                        Call cmdSSITIC_Detail_Display("YSSITIC0")
                End Select
            End Select

End Sub

Public Sub paramSSIUSRPRFX_Load()
Dim xSQL As String
Dim rsSab As New ADODB.Recordset

cboSSIUSRPRFX.Clear
cboSSIUSRPRFX.AddItem ""

xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
     & " where SSIUSRNAT = '$'"
Set rsSab = cnsab.Execute(xSQL)
arrSSIUSRPRFX_UB = rsSab(0)

ReDim arrSSIUSRPRFX(arrSSIUSRPRFX_UB + 1)

xSQL = "select SSIUSRUIDX from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
     & " where SSIUSRNAT = '$'" _
     & " order by SSIUSRUIDX"

Set rsSab = cnsab.Execute(xSQL)
arrSSIUSRPRFX_UB = 0

Do While Not rsSab.EOF
    arrSSIUSRPRFX_UB = arrSSIUSRPRFX_UB + 1
    arrSSIUSRPRFX(arrSSIUSRPRFX_UB) = Trim(rsSab(0))
    cboSSIUSRPRFX.AddItem arrSSIUSRPRFX(arrSSIUSRPRFX_UB)
    rsSab.MoveNext
Loop
cboSSIUSRPRFX.ListIndex = 0

End Sub

Public Function cmdUpdate_SSIUSRPRFK_DOM(lSSIUSRUIDN As Long)

Dim V, xSQL As String

cmdUpdate_SSIUSRPRFK_DOM = Null
xSQL = "select count(*)  from " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
     & "  Where SSIDOMNAT = ' ' and SSIDOMUIDN = " & lSSIUSRUIDN & " and SSIDOMPRFK not in (' ','X') "
Set rsSab = cnsab.Execute(xSQL)

If rsSab(0) > 0 Then
    mYSSIUSR0_Update_CMD = "Update " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
                         & " set SSIUSRPRFK = 'N' " _
                         & " Where SSIUSRNAT = ' ' and SSIUSRUIDN = " & lSSIUSRUIDN & " and SSIUSRPRFK = ' '"
    Call FEU_ROUGE
    V = sqlYSSIUSR0_Update_CMD(mYSSIUSR0_Update_CMD)
    Call FEU_VERT
Else
    mYSSIUSR0_Update_CMD = "Update " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
                         & " set SSIUSRPRFK = ' ' " _
                         & " Where SSIUSRNAT = ' ' and SSIUSRUIDN = " & lSSIUSRUIDN & " and SSIUSRPRFK <> ' '"
    Call FEU_ROUGE
    V = sqlYSSIUSR0_Update_CMD(mYSSIUSR0_Update_CMD)
    Call FEU_VERT

End If
cmdUpdate_SSIUSRPRFK_DOM = V
End Function
Public Sub cmdUpdate_PRFK_DECH()

Dim xSQL As String

mYSSIUSR0_Update = "CMD"
mYSSIUSR0_Update_CMD = "Update " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
                     & " set SSIUSRPRFK = '!' " _
                     & " Where SSIUSRNAT = ' '  and SSIUSRPRFK = ' ' and SSIUSRSTAK = ' ' and SSIUSRDECH > 0 and SSIUSRDECH < " & DSys

mYSSIDOM0_Update = "CMD"
mYSSIDOM0_Update_CMD = "Update " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
                     & " set SSIDOMPRFK = '!' " _
                     & " Where SSIDOMNAT = ' '  and SSIDOMPRFK = ' ' and SSIDOMSTAK = ' ' and SSIDOMDECH > 0 and SSIDOMDECH < " & DSys
Call cmdUpdate

End Sub



Public Sub paramModèle_Init()
Dim xIn As String, K As Integer, X As String, X1 As String, X2 As String
Dim xSQL As String, blnSAB As Boolean, wSSIDOMPRFX_SAB As String, blnSAA As Boolean, wSSIDOMPRFX_SAA As String
Dim rsSab_X As New ADODB.Recordset

Call rsYSSIUSR0_Init(newYSSIUSR0)
newYSSIUSR0.SSIUSRYFCT = "INI"
newYSSIUSR0.SSIUSRYAMJ = DSys
newYSSIUSR0.SSIUSRYHMS = time_Hms
newYSSIUSR0.SSIUSRYUSR = usrName_UCase

Call rsYSSITXT0_Init(newYSSITXT0)
newYSSITXT0.SSITXTYAMJ = newYSSIUSR0.SSIUSRYAMJ
newYSSITXT0.SSITXTYHMS = newYSSIUSR0.SSIUSRYHMS
newYSSITXT0.SSITXTYUSR = newYSSIUSR0.SSIUSRYUSR

Call rsYSSIDOM0_Init(newYSSIDOM0)
newYSSIDOM0.SSIDOMYFCT = "INI"
newYSSIDOM0.SSIDOMYAMJ = newYSSIUSR0.SSIUSRYAMJ
newYSSIDOM0.SSIDOMYHMS = newYSSIUSR0.SSIUSRYHMS
newYSSIDOM0.SSIDOMYUSR = newYSSIUSR0.SSIUSRYUSR

'X = "select SSIUSRUIDN from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
'     & "  Where SSIUSRNAT = '$' order by SSIUSRUIDN desc"
'Set rsSab = cnsab.Execute(xSQL)


Open "C:\TEMP\BIA_SSI\BIA_SSI Modèle.txt" For Input As 1


Do Until EOF(1)
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        K = 0
        X = Trim(CSV_Scan(xIn, K))
        X1 = Trim(CSV_Scan(xIn, K))
        X2 = Trim(CSV_Scan(xIn, K))
        If X2 <> "" Then
        
            xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
                 & " where SSIDOMNAT = ' ' and SSIDOMDIDX = 'SAB' and SSIDOMUIDX = '" & X1 & "'"
            
            Set rsSab_X = cnsab.Execute(xSQL)
            If Not rsSab_X.EOF Then
              Call rsYSSIDOM0_GetBuffer(rsSab_X, xYSSIDOM0)
        
                newYSSIUSR0.SSIUSRNAT = "$"
                newYSSIUSR0.SSIUSRUIDX = X
                newYSSITXT0.SSITXTINFO = X2
                newYSSITXT0.SSITXTNAT = "$"
                mYSSIUSR0_Update = "New": mYSSITXT0_Update = "New": Call cmdUpdate
                newYSSIDOM0.SSIDOMNAT = "$"
                newYSSIDOM0.SSIDOMUIDD = 0
                newYSSIDOM0.SSIDOMTLNK = 0
                newYSSIDOM0.SSIDOMUIDN = newYSSIUSR0.SSIUSRUIDN
               blnSAA = False: blnSAB = False
               Select Case X
                    Case "BIA_ADMIN": blnSAB = True: wSSIDOMPRFX_SAB = "P_ADMIN"
                    Case "BIA_ADMIN   /3": blnSAB = True: wSSIDOMPRFX_SAB = "P_ADMIN   /3"
                    Case "BIA_CAC": blnSAB = True: wSSIDOMPRFX_SAB = "P_CAC"
                    Case "BIA_QSYSOPR"
                    Case "BIA_DAFI_S":
                        blnSAB = True: wSSIDOMPRFX_SAB = "P_DAFI_S"
                        blnSAA = True: wSSIDOMPRFX_SAA = "BIA_SV"
                    Case "BIA_SOBF_S3":
                        blnSAB = True: wSSIDOMPRFX_SAB = "P_SOBF_S3"
                        blnSAA = True: wSSIDOMPRFX_SAA = "BIA_SV"
                End Select
                
                If blnSAA Then
                        newYSSIDOM0.SSIDOMDIDX = "SAA"
                        newYSSIDOM0.SSIDOMPRFX = wSSIDOMPRFX_SAA
                        newYSSIDOM0.SSIDOMUIDX = wSSIDOMPRFX_SAA
                        newYSSIDOM0.SSIDOMUIDD = newYSSIDOM0.SSIDOMUIDD - 1
                        mYSSIDOM0_Update = "New": mYSSITXT0_Update = "": Call cmdUpdate
                End If
                
                If blnSAB Then
                        newYSSIDOM0.SSIDOMDIDX = "IBM"
                        newYSSIDOM0.SSIDOMPRFX = "SAB_PROD"
                        newYSSIDOM0.SSIDOMUIDX = "SAB_PROD"
                        newYSSIDOM0.SSIDOMUIDD = newYSSIDOM0.SSIDOMUIDD - 1
                        mYSSIDOM0_Update = "New": mYSSITXT0_Update = "": Call cmdUpdate
                        newYSSIDOM0.SSIDOMDIDX = "IBM"
                        newYSSIDOM0.SSIDOMPRFX = "SAB_TEST"
                        newYSSIDOM0.SSIDOMUIDX = "SAB_TEST"
                        newYSSIDOM0.SSIDOMUIDD = newYSSIDOM0.SSIDOMUIDD - 1
                        mYSSIDOM0_Update = "New": mYSSITXT0_Update = "": Call cmdUpdate
                        newYSSIDOM0.SSIDOMDIDX = "SAB"
                        newYSSIDOM0.SSIDOMPRFX = wSSIDOMPRFX_SAB
                        newYSSIDOM0.SSIDOMUIDX = wSSIDOMPRFX_SAB
                        newYSSIDOM0.SSIDOMUIDD = newYSSIDOM0.SSIDOMUIDD - 1
                        mYSSIDOM0_Update = "New": mYSSITXT0_Update = "": Call cmdUpdate
                Else
                
                    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
                         & " where  SSIDOMNAT = ' ' and  SSIDOMUIDN = " & xYSSIDOM0.SSIDOMUIDN _
                         & " and SSIDOMPRFX <> 'X' order by SSIDOMDIDX , SSIDOMUIDX"
                    
                    Set rsSab_X = cnsab.Execute(xSQL)
                    Do While Not rsSab_X.EOF
                        Call rsYSSIDOM0_GetBuffer(rsSab_X, xYSSIDOM0)
                        If xYSSIDOM0.SSIDOMPRFK = " " Then
                            newYSSIDOM0.SSIDOMDIDX = xYSSIDOM0.SSIDOMDIDX
                            newYSSIDOM0.SSIDOMPRFX = xYSSIDOM0.SSIDOMPRFX
                            newYSSIDOM0.SSIDOMUIDX = xYSSIDOM0.SSIDOMPRFX
                            newYSSIDOM0.SSIDOMUIDD = newYSSIDOM0.SSIDOMUIDD - 1
                            mYSSIDOM0_Update = "New": mYSSITXT0_Update = "": Call cmdUpdate
                        End If
                        rsSab_X.MoveNext
                    Loop
                End If
            End If
        End If
    End If
Loop

Close
'______________________________________________________________________________________________
Open "C:\TEMP\BIA_SSI\BIA_SSI Modèle.txt" For Input As 1
Do Until EOF(1)
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        K = 0
        X = Trim(CSV_Scan(xIn, K))
        X1 = Trim(CSV_Scan(xIn, K))
        
         xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
              & " where SSIDOMNAT = ' ' and SSIDOMDIDX = 'SAB' and SSIDOMUIDX = '" & X1 & "'"
         
         Set rsSab_X = cnsab.Execute(xSQL)
         If Not rsSab_X.EOF Then
              xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
                  & " where SSIUSRNAT = ' ' and SSIUSRUIDN = " & rsSab_X("SSIDOMUIDN")
             
             Set rsSab_X = cnsab.Execute(xSQL)
             If Not rsSab_X.EOF Then
               Call rsYSSIUSR0_GetBuffer(rsSab_X, oldYSSIUSR0)
    
                 newYSSIUSR0 = oldYSSIUSR0
                 
                 newYSSIUSR0.SSIUSRPRFX = X
                 mYSSIUSR0_Update = "Update": mYSSITXT0_Update = "": Call cmdUpdate
             End If
        End If
    End If
Loop

Close

End Sub

Public Sub fgDetail_Display_lstW()
Dim xSQL As String
Set lstW.Container = fraSelect
lstW.ZOrder 0
lstW.Height = fgDetail.Height
lstW.Top = fraSelect.Top + fraSelect.Height - lstW.Height - 1000
lstW.Width = 4000
lstW.Left = fraDetail.Left + fraDetail.Width - lstW.Width - 300
lstW.BackColor = mColor_G0
lstW.Clear

xSQL = "select SSIUSRUIDX from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
 & " where SSIUSRNAT = ' ' and SSIUSRPRFX = '" & oldYSSIUSR0.SSIUSRUIDX & "'" _
 & " order by SSIUSRUIDX "


Set rsSab_X = cnsab.Execute(xSQL)
Do While Not rsSab_X.EOF
    lstW.AddItem rsSab_X("SSIUSRUIDX")
    rsSab_X.MoveNext
Loop
lstW.Visible = True

End Sub

Public Sub cmdProfil_Excel_SAB()
On Error GoTo Error_Handler
Dim X As String, K As Long, I As Long, xPRE As String, wSSISABULOT As Long
Dim blnOk As Boolean, Iter As Integer
Dim blnYSSISAM0 As Boolean

Call lstErr_AddItem(lstErr, cmdContext, "cmdProfil_Excel_SAB"): DoEvents



If blnProfil_Excel_All Then
    X = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
         & " where SSISABNAT = '$' and SSISABSTAK = ' ' "
    Set rsSab = cnsab.Execute(X)
    ReDim arrProfil_PX(rsSab(0) + 1), arrProfil_DX(rsSab(0) + 1), arrProfil_MX(rsSab(0) + 1), arrProfil_GX(rsSab(0) + 1), arrProfil_GN(rsSab(0) + 1)

   arrProfil_Nb = 0
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
         & " where SSISABNAT = '$' and SSISABSTAK = ' ' order by SSISABUIDX "
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        arrProfil_Nb = arrProfil_Nb + 1
        arrProfil_PX(arrProfil_Nb) = Trim(rsSab("SSISABUIDX"))
        arrProfil_GX(arrProfil_Nb) = Mid$(rsSab("SSISABUNOM"), 1, 10)
        arrProfil_DX(arrProfil_Nb) = Mid$(rsSab("SSISABUNOM"), 11, 10)
        arrProfil_MX(arrProfil_Nb) = Mid$(rsSab("SSISABUNOM"), 21, 10)
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
             & " where SSISABNAT = '2' and SSISABUIDX ='" & arrProfil_GX(arrProfil_Nb) & "' and SSISABSTAK = ' ' "
        Set rsSab_X = cnsab.Execute(X)
        If Not rsSab_X.EOF Then arrProfil_GN(arrProfil_Nb) = rsSab_X("SSISABUIDD")
        rsSab.MoveNext
    Loop
Else
    arrProfil_Nb = lstW.ListCount
    ReDim arrProfil_PX(arrProfil_Nb), arrProfil_DX(arrProfil_Nb), arrProfil_GX(arrProfil_Nb), arrProfil_GN(arrProfil_Nb), arrProfil_MX(arrProfil_Nb)
    For K = 1 To lstW.ListCount
        lstW.ListIndex = K - 1
        arrProfil_PX(K) = lstW.Text
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
             & " where SSISABNAT = '$' and SSISABUIDX ='" & arrProfil_PX(K) & "' and SSISABSTAK = ' ' "
        Set rsSab = cnsab.Execute(X)
        If Not rsSab.EOF Then
            arrProfil_GX(K) = Mid$(rsSab("SSISABUNOM"), 1, 10)
            arrProfil_DX(K) = Mid$(rsSab("SSISABUNOM"), 11, 10)
            arrProfil_MX(K) = Mid$(rsSab("SSISABUNOM"), 21, 10)
            X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
                 & " where SSISABNAT = '2' and SSISABUIDX ='" & arrProfil_GX(K) & "' and SSISABSTAK = ' ' "
            Set rsSab_X = cnsab.Execute(X)
            If Not rsSab_X.EOF Then arrProfil_GN(K) = rsSab_X("SSISABUIDD")
        End If
    Next K
    
End If


Call cmdProfil_Excel_SAB_Init(1, "SAB options de menu")
'==========================================================================================================
Call cmdProfil_Excel_SAB_USR
'==========================================================================================================
X = MsgBox("Voulez-vous extraire les options de menu (OUI) ?", vbYesNoCancel, "Excel : sélection des options de menu")
Select Case X
    Case vbYes: blnZMNUOPT0 = True
    Case vbNo: blnZMNUOPT0 = False
    Case vbCancel: Exit Sub
End Select

If blnZMNUOPT0 Then
    X = "select count(*) from " & paramIBM_Library_SAB & ".ZMNUOPT0 "
    Set rsSab = cnsab.Execute(X)
    arrYSSIMNU0_Nb = 0
    ReDim arrYSSIMNU0(rsSab(0) + 1)
    
    X = "select * from " & paramIBM_Library_SAB & ".ZMNUOPT0, " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
         & " where SSISABNAT = 'M' and SSISABSTAK = ' ' and ssisabuidx = mnuoptcod order by  mnuoptcod "
    Set rsSab = cnsab.Execute(X)
    Do While Not rsSab.EOF
        If arrYSSIMNU0_Nb = 0 Then wSSISABULOT = rsSab("SSISABULOT")
        arrYSSIMNU0_Nb = arrYSSIMNU0_Nb + 1
        arrYSSIMNU0(arrYSSIMNU0_Nb).SSIMNUCOD = Val(rsSab("SSISABUIDX"))
        arrYSSIMNU0(arrYSSIMNU0_Nb).SSIMNUARB = ""
        arrYSSIMNU0(arrYSSIMNU0_Nb).SSIMNULIB = rsSab("MNUOPTLIB")
        arrYSSIMNU0(arrYSSIMNU0_Nb).SSIMNUINFO = rsSab("SSISABINFO")
        arrYSSIMNU0(arrYSSIMNU0_Nb).SSIMNUENS = rsSab("MNUOPTENS")
        rsSab.MoveNext
    Loop
    '____________________________________________________________________________________________
    Call lstErr_AddItem(lstErr, cmdContext, "ZMNUMEN0 : G_ADMIN"): DoEvents
    X = "select * from " & paramIBM_Library_SAB & ".ZMNUMEN0" _
         & " where MNUMENGRP = 'G_ADMIN' and MNUMENCOD > 0 " _
         & " and   MNUMENREF = " & wSSISABULOT _
         & " and   MNUMENETB = 1 " _
         & " order by MNUMENPRE, MNUMENORD"
    Set rsSab = cnsab.Execute(X)
    Do While Not rsSab.EOF
        For K = 1 To arrYSSIMNU0_Nb
            If rsSab("MNUMENCOD") = arrYSSIMNU0(K).SSIMNUCOD Then
                arrYSSIMNU0(K).SSIMNUPRE = rsSab("MNUMENPRE")
                arrYSSIMNU0(K).SSIMNUORD = rsSab("MNUMENORD")
                If arrYSSIMNU0(K).SSIMNUPRE = 0 Then arrYSSIMNU0(K).SSIMNUARB = Format$(rsSab("MNUMENORD"), "0000000")
                Exit For
            End If
        Next K
        rsSab.MoveNext
    Loop
    '____________________________________________________________________________________________
    Call lstErr_AddItem(lstErr, cmdContext, "ZMNUMEN0 : SABTELE"): DoEvents
    X = "select * from " & paramIBM_Library_SAB & ".ZMNUMEN0" _
         & " where MNUMENGRP = 'SABTELE' and MNUMENCOD > 0 " _
         & " and   MNUMENREF = " & wSSISABULOT _
         & " and   MNUMENETB = 1 " _
         & " order by MNUMENPRE, MNUMENORD"
    Set rsSab = cnsab.Execute(X)
    Do While Not rsSab.EOF
        For K = 1 To arrYSSIMNU0_Nb
            If rsSab("MNUMENCOD") = arrYSSIMNU0(K).SSIMNUCOD Then
                If arrYSSIMNU0(K).SSIMNUPRE = 0 Then
                    arrYSSIMNU0(K).SSIMNUPRE = rsSab("MNUMENPRE")
                    arrYSSIMNU0(K).SSIMNUORD = rsSab("MNUMENORD")
                    If arrYSSIMNU0(K).SSIMNUPRE = 0 Then arrYSSIMNU0(K).SSIMNUARB = Format$(rsSab("MNUMENORD"), "0000000")
                End If
                Exit For
            End If
        Next K
        rsSab.MoveNext
    Loop
    
    
    '____________________________________________________________________________________________
    Call lstErr_AddItem(lstErr, cmdContext, "Arborescence"): DoEvents
    Do
        blnOk = True
        For I = 1 To arrYSSIMNU0_Nb
            If arrYSSIMNU0(I).SSIMNUPRE > 0 And arrYSSIMNU0(I).SSIMNUARB = "" Then
                For K = 1 To arrYSSIMNU0_Nb
                    If arrYSSIMNU0(I).SSIMNUPRE = arrYSSIMNU0(K).SSIMNUCOD Then
                        If arrYSSIMNU0(K).SSIMNUARB = "" Then
                            blnOk = False
                        Else
                            arrYSSIMNU0(I).SSIMNUARB = arrYSSIMNU0(K).SSIMNUARB & Format$(arrYSSIMNU0(I).SSIMNUORD, "0000000")
                        End If
                        Exit For
                    End If
                Next K
            End If
        Next I
        
    Iter = Iter + 1: If Iter > 7 Then blnOk = True
    Loop Until blnOk
    
    
    Call cmdProfil_Excel_SAB_Detail("OPT")
End If
'==========================================================================================================
For K = 1 To arrProfil_Nb - 1
    If arrProfil_MX(K) <> "" Then
        For I = K + 1 To arrProfil_Nb
            If arrProfil_MX(K) = arrProfil_MX(I) Then arrProfil_MX(I) = ""
        Next I
    End If
Next K


Call cmdProfil_Excel_SAB_M_Init
oldYSSISAM0.SSISAMCLA = -1



X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAM0, " & paramIBM_Library_SAB & ".ZHBTI010" _
     & " where HBTI01LAN = '1' and HBTI01APP = SSISAMAPP and HBTI01COD = SSISAMCOD " _
     & " order by  SSISAMCLA,SSISAMAPP,SSISAMCOD,SSISAMAGE,SSISAMSER,SSISAMSSE,SSISAMOPE,SSISAMNAT,SSISAMPRD,SSISAMAUT,SSISAMREF,SSISAMGRP "
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    For K = 1 To arrProfil_Nb
        If rsSab("SSISAMGRP") = arrProfil_MX(K) Then
            mXls2_Row = mXls2_Row + 1
            Call rsYSSISAM0_GetBuffer(rsSab, xYSSISAM0)
            wsExcel.Cells(mXls2_Row, 6) = xYSSISAM0.SSISAMGRP
            If xYSSISAM0.SSISAMCLA = oldYSSISAM0.SSISAMCLA _
            And xYSSISAM0.SSISAMAPP = oldYSSISAM0.SSISAMAPP _
            And xYSSISAM0.SSISAMCOD = oldYSSISAM0.SSISAMCOD _
            And xYSSISAM0.SSISAMAGE = oldYSSISAM0.SSISAMAGE _
            And xYSSISAM0.SSISAMSER = oldYSSISAM0.SSISAMSER _
            And xYSSISAM0.SSISAMSSE = oldYSSISAM0.SSISAMSSE _
            And xYSSISAM0.SSISAMOPE = oldYSSISAM0.SSISAMOPE _
            And xYSSISAM0.SSISAMNAT = oldYSSISAM0.SSISAMNAT _
            And xYSSISAM0.SSISAMPRD = oldYSSISAM0.SSISAMPRD _
            And xYSSISAM0.SSISAMAUT = oldYSSISAM0.SSISAMAUT _
            And xYSSISAM0.SSISAMREF = oldYSSISAM0.SSISAMREF Then
                'wsExcel.Cells(mXls2_Row, 5).Interior.Color = mColor_G1
            Else
                oldYSSISAM0 = xYSSISAM0
                wsExcel.Cells(mXls2_Row, 1) = oldYSSISAM0.SSISAMCLA & "_" & oldYSSISAM0.SSISAMUIDD
                wsExcel.Cells(mXls2_Row, 2) = Trim(rsSab("HBTI01LIB"))
                wsExcel.Cells(mXls2_Row, 3) = oldYSSISAM0.SSISAMAPP & " " & oldYSSISAM0.SSISAMCOD
                wsExcel.Cells(mXls2_Row, 4) = oldYSSISAM0.SSISAMOPE & " " & oldYSSISAM0.SSISAMNAT & " " & oldYSSISAM0.SSISAMPRD
                wsExcel.Cells(mXls2_Row, 5) = Format(oldYSSISAM0.SSISAMAGE, "##") & oldYSSISAM0.SSISAMSER & oldYSSISAM0.SSISAMSSE & " " & oldYSSISAM0.SSISAMAUT & " " & oldYSSISAM0.SSISAMREF
                For I = 1 To mXls2_Cols: wsExcel.Cells(mXls2_Row, I).Interior.Color = mColor_G0: Next I
            End If
            wsExcel.Cells(mXls2_Row, 7) = xYSSISAM0.SSISAMFON
            wsExcel.Cells(mXls2_Row, 8) = xYSSISAM0.SSISAMDON
            wsExcel.Cells(mXls2_Row, 9) = xYSSISAM0.SSISAMCAI
            If xYSSISAM0.SSISAMMON <> 0 Then wsExcel.Cells(mXls2_Row, 10) = Format(xYSSISAM0.SSISAMMON, "### ### ### ###.##") & " " & xYSSISAM0.SSISAMDEV
            wsExcel.Cells(mXls2_Row, 11) = xYSSISAM0.SSISAMDLY
            wsExcel.Cells(mXls2_Row, 12) = xYSSISAM0.SSISAMPRO
            wsExcel.Cells(mXls2_Row, 13) = xYSSISAM0.SSISAMCLI
            wsExcel.Cells(mXls2_Row, 14) = xYSSISAM0.SSISAMEIC
            wsExcel.Cells(mXls2_Row, 15) = xYSSISAM0.SSISAMSDD
            wsExcel.Cells(mXls2_Row, 16) = xYSSISAM0.SSISAMDRO
            wsExcel.Cells(mXls2_Row, 17) = xYSSISAM0.SSISAMSUC
            wsExcel.Cells(mXls2_Row, 18) = xYSSISAM0.SSISAMPRC
            wsExcel.Cells(mXls2_Row, 19) = xYSSISAM0.SSISAMNJ1
            wsExcel.Cells(mXls2_Row, 20) = xYSSISAM0.SSISAMTJ1
            wsExcel.Cells(mXls2_Row, 21) = xYSSISAM0.SSISAMNJ2
            wsExcel.Cells(mXls2_Row, 22) = xYSSISAM0.SSISAMTJ2
            wsExcel.Cells(mXls2_Row, 23) = xYSSISAM0.SSISAMPRA
            wsExcel.Cells(mXls2_Row, 24) = xYSSISAM0.SSISAMECH
        End If
        
    Next K
    
    rsSab.MoveNext
Loop

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdProfil_Excel_SAB"
End Sub
Public Sub cmdProfil_Excel_SAB_Init(lSheet As Integer, lLib As String)

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
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14" & lLib _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True
wsExcel.PageSetup.PrintTitleRows = "$A1:$G1"

wsExcel.PageSetup.Zoom = 80

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : " & lLib): DoEvents

mXls2_Cols = 4 + arrProfil_Nb
mXls2_Row = 1

wsExcel.Rows(1).Orientation = xlVertical
wsExcel.Rows(1).RowHeight = 130
'wsExcel.Rows(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Rows(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Rows(1).VerticalAlignment = Excel.xlVAlignTop

For K = 1 To arrProfil_Nb
    wsExcel.Cells(1, 4 + K) = arrProfil_GX(K): wsExcel.Columns(4 + K).ColumnWidth = 2
    wsExcel.Columns(4 + K).HorizontalAlignment = Excel.xlHAlignCenter
    'wsExcel.Columns(4 + K).VerticalAlignment = Excel.xlVAlignTop
Next K

wsExcel.Cells(1, 1) = "Application": wsExcel.Columns(1).ColumnWidth = 12
wsExcel.Cells(1, 1).Orientation = xlHorizontal
wsExcel.Cells(1, 1).VerticalAlignment = Excel.xlVAlignCenter
wsExcel.Cells(1, 2) = "X": wsExcel.Columns(2).ColumnWidth = 8
wsExcel.Cells(1, 2).Orientation = xlHorizontal
wsExcel.Cells(1, 2).VerticalAlignment = Excel.xlVAlignCenter
wsExcel.Cells(1, 3) = "Code": wsExcel.Columns(3).ColumnWidth = 10
wsExcel.Cells(1, 3).Orientation = xlHorizontal
wsExcel.Cells(1, 3).VerticalAlignment = Excel.xlVAlignCenter
wsExcel.Cells(1, 4) = "Libellé": wsExcel.Columns(4).ColumnWidth = 55
wsExcel.Cells(1, 4).Orientation = xlHorizontal
wsExcel.Cells(1, 4).VerticalAlignment = Excel.xlVAlignCenter


For K = 1 To mXls2_Cols
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next

'==========================================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdProfil_Excel_SAB_Init"


End Sub
Public Sub cmdProfil_Excel_SAB_M_Init()

On Error GoTo Error_Handler
Dim K As Integer

'==========================================================================================================

Set wsExcel = wbExcel.Sheets(2)
wsExcel.Name = 2 & "-métiers"

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
    .Font.Name = "Courier New"
    .Font.Color = vbBlack 'RGB(0, 64, 128)
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14" & "Habilitations métiers" _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True
wsExcel.PageSetup.PrintTitleRows = "$A1:$G1"

wsExcel.PageSetup.Zoom = 80

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : Habilitations métiers"): DoEvents

mXls2_Cols = 24
mXls2_Row = 1

wsExcel.Rows(1).HorizontalAlignment = Excel.xlHAlignCenter

wsExcel.Cells(1, 1) = "Classe_N°": wsExcel.Columns(1).ColumnWidth = 8
wsExcel.Cells(1, 2) = "Libellé": wsExcel.Columns(2).ColumnWidth = 30: wsExcel.Columns(2).Font.Color = vbBlue
wsExcel.Cells(1, 3) = "Application": wsExcel.Columns(3).ColumnWidth = 8
wsExcel.Cells(1, 4) = "Opé/nat/type cpt": wsExcel.Columns(4).ColumnWidth = 15
wsExcel.Cells(1, 5) = "Ag/service": wsExcel.Columns(5).ColumnWidth = 8

wsExcel.Cells(1, 6) = "Groupe": wsExcel.Columns(6).ColumnWidth = 12: wsExcel.Columns(6).Font.Color = vbMagenta
wsExcel.Columns(6).Font.Bold = True

wsExcel.Cells(1, 7) = "Fonctions": wsExcel.Columns(7).ColumnWidth = 10
wsExcel.Cells(1, 8) = "Données": wsExcel.Columns(8).ColumnWidth = 10
wsExcel.Cells(1, 9) = "Caisses": wsExcel.Columns(9).ColumnWidth = 10
wsExcel.Cells(1, 10) = "Mt plafond": wsExcel.Columns(10).ColumnWidth = 24
wsExcel.Cells(1, 11) = "Délais": wsExcel.Columns(11).ColumnWidth = 5
wsExcel.Cells(1, 12) = "Profil": wsExcel.Columns(12).ColumnWidth = 10
wsExcel.Cells(1, 13) = "Clients": wsExcel.Columns(13).ColumnWidth = 15
wsExcel.Cells(1, 14) = "Utilisateur": wsExcel.Columns(14).ColumnWidth = 13
wsExcel.Cells(1, 15) = "Mandat": wsExcel.Columns(15).ColumnWidth = 23
wsExcel.Cells(1, 16) = "Opération": wsExcel.Columns(16).ColumnWidth = 6
wsExcel.Cells(1, 17) = "Sup com": wsExcel.Columns(17).ColumnWidth = 5
wsExcel.Cells(1, 18) = "% com": wsExcel.Columns(18).ColumnWidth = 5
wsExcel.Cells(1, 19) = "Nbj marge +": wsExcel.Columns(19).ColumnWidth = 5
wsExcel.Cells(1, 20) = "Type marge +": wsExcel.Columns(20).ColumnWidth = 5
wsExcel.Cells(1, 21) = "Nbj marge -": wsExcel.Columns(21).ColumnWidth = 5
wsExcel.Cells(1, 22) = "Type marge -": wsExcel.Columns(22).ColumnWidth = 5
wsExcel.Cells(1, 23) = "% mt arrondi": wsExcel.Columns(23).ColumnWidth = 5
wsExcel.Cells(1, 24) = "Echelles": wsExcel.Columns(24).ColumnWidth = 10

For K = 1 To mXls2_Cols
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next

'==========================================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdProfil_Excel_SAB_Init"


End Sub

Public Sub cmdProfil_Excel_SAB_Detail(lFct As String)
Dim X As String, K As Integer, K2 As Integer, mRow As Integer, xSQL As String, blnMnu_Ok As Boolean
Dim I As Integer
On Error GoTo Error_Handler
'==========================================================================================================
Call lstErr_AddItem(lstErr, cmdContext, "cmdProfil_Excel_SAB_Detail"): DoEvents
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Rows = arrYSSIMNU0_Nb + 1
fgSelect.Row = 0
For K = 1 To arrYSSIMNU0_Nb
    fgSelect.Row = fgSelect.Row + 1
    If arrYSSIMNU0(K).SSIMNUARB = "" Then
        fgSelect.Col = 0: fgSelect.Text = "9999999"
    Else
        fgSelect.Col = 0: fgSelect.Text = arrYSSIMNU0(K).SSIMNUARB
    End If
    fgSelect.Col = 2: fgSelect.Text = K
    fgSelect.Col = 1: fgSelect.Text = arrYSSIMNU0(K).SSIMNUCOD

Next K
fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort

For mRow = 1 To fgSelect.Rows - 1
    blnMnu_Ok = blnProfil_Excel_All
    fgSelect.Row = mRow
    fgSelect.Col = 2: K2 = Val(fgSelect.Text)
    X = arrYSSIMNU0(K2).SSIMNUINFO
    If Not blnProfil_Excel_All Then
        For K = 1 To arrProfil_Nb
            If Mid$(X, arrProfil_GN(K), 1) <> " " Then blnMnu_Ok = True: Exit For
        Next K
    End If
    
    
    If blnMnu_Ok Then
        mXls2_Row = mXls2_Row + 1
        I = Len(arrYSSIMNU0(K2).SSIMNUARB) / 7
        wsExcel.Cells(mXls2_Row, 1) = Trim(arrYSSIMNU0(K2).SSIMNUENS)
        If arrYSSIMNU0(K2).SSIMNUARB = "" Then wsExcel.Cells(mXls2_Row, 3).Interior.Color = mColor_Y2
        wsExcel.Cells(mXls2_Row, 3) = Trim(arrYSSIMNU0(K2).SSIMNUCOD)
        wsExcel.Cells(mXls2_Row, 4) = String(I, ". ") & Trim(arrYSSIMNU0(K2).SSIMNULIB)
        If InStr(arrYSSIMNU0(K2).SSIMNULIB, "MENU") > 0 Then
            Call lstErr_ChangeLastItem(lstErr, cmdContext, arrYSSIMNU0(K2).SSIMNULIB): DoEvents
            For K = 1 To arrProfil_Nb + 4
                wsExcel.Cells(mXls2_Row, K).Font.Color = vbBlue
                wsExcel.Cells(mXls2_Row, K).Font.Bold = True
            Next K
        End If
        For K = 1 To arrProfil_Nb
            If Mid$(X, arrProfil_GN(K), 1) <> " " Then
                wsExcel.Cells(mXls2_Row, K + 4) = Mid$(X, arrProfil_GN(K), 1)
                wsExcel.Cells(mXls2_Row, K + 4).Interior.Color = mColor_G1
            End If
            
        Next K
    End If

'_________________________________________________________________________________________________
Next mRow


'==========================================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdProfil_Excel_SAB_Detail"

End Sub





Public Sub cmdProfil_Excel_SAB_USR()
On Error GoTo Error_Handler
Dim X As String, K As Long, K1 As Long, wSSISABUNOM As String
Dim mProfil_PX As String

Call lstErr_AddItem(lstErr, cmdContext, "cmdProfil_Excel_SAB_USR"): DoEvents

 For K = 1 To arrProfil_Nb
    Call lstErr_ChangeLastItem(lstErr, cmdContext, arrProfil_GX(K)): DoEvents
   ' X = "select SSISABUNOM from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
   '      & " where SSISABNAT = '$' and SSISABUIDX = '" & arrProfil_PX(K) & "'" _
   '      & " and SSISABSTAK = ' '"
   ' Set rsSab = cnsab.Execute(X)
   ' If Not rsSab.EOF Then
   '     wSSISABUNOM = rsSab("SSISABUNOM")
   ' Else
   '     wSSISABUNOM = Space(30)
   ' End If
    
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
         & " where SSISABNAT = ' ' and SSISABPRFX = '" & arrProfil_PX(K) & "'" _
         & " and SSISABSTAK = ' ' order by SSISABUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        If mProfil_PX <> arrProfil_PX(K) Then
            mProfil_PX = arrProfil_PX(K)
            mXls2_Row = mXls2_Row + 1
            wsExcel.Cells(mXls2_Row, 1) = arrProfil_PX(K)
            wsExcel.Cells(mXls2_Row, 1).Font.Bold = True
            wsExcel.Cells(mXls2_Row, 2) = "Profil SAB"
            wsExcel.Cells(mXls2_Row, 1).Font.Color = vbMagenta
            wsExcel.Cells(mXls2_Row, 4) = arrProfil_GX(K) & "  |  " & arrProfil_DX(K) & "  |  " & arrProfil_MX(K)
            wsExcel.Cells(mXls2_Row, 4).Font.Color = vbMagenta
            wsExcel.Cells(mXls2_Row, 4).Font.Bold = True
        End If
        
        mXls2_Row = mXls2_Row + 1
        
        wsExcel.Cells(mXls2_Row, 1) = arrProfil_PX(K)
        wsExcel.Cells(mXls2_Row, 1).Font.Bold = True
        wsExcel.Cells(mXls2_Row, 1).Font.Color = vbBlue
        wsExcel.Cells(mXls2_Row, 2) = "Utilisateur"
        wsExcel.Cells(mXls2_Row, 3) = rsSab("SSISABUIDD")
        wsExcel.Cells(mXls2_Row, 4) = Trim(rsSab("SSISABUIDX"))
        wsExcel.Cells(mXls2_Row, 4).Font.Bold = True
        wsExcel.Cells(mXls2_Row, 4).Font.Color = vbBlue
        wsExcel.Cells(mXls2_Row, K + 4) = "X"
        wsExcel.Cells(mXls2_Row, K + 4).Interior.Color = mColor_G1
       rsSab.MoveNext
    Loop
'___________________________________________________________________________________________________________________
 '   mXls2_Row = mXls2_Row + 1


    X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
         & " where SSISABNAT = 'C' and SSISABUIDX = '" & arrProfil_DX(K) & "'" _
         & " and SSISABSTAK = ' ' order by SSISABUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        X = rsSab("SSISABINFO")
        For K1 = 1 To 99
            If Mid$(X, K1, 1) <> " " Then
                mXls2_Row = mXls2_Row + 1
                wsExcel.Cells(mXls2_Row, 1) = arrProfil_DX(K)
                wsExcel.Cells(mXls2_Row, 2) = "Classe"
                wsExcel.Cells(mXls2_Row, 3) = K1
                 Select Case Mid$(X, K1, 1)
                     Case Is = "1"
                         wsExcel.Cells(mXls2_Row, 4) = arrMNURCLABR(K1) & " : consultation"
                         wsExcel.Cells(mXls2_Row, K + 4) = "C"
                         wsExcel.Cells(mXls2_Row, K + 4).Interior.Color = mColor_G1
                    Case Is = "2"
                         wsExcel.Cells(mXls2_Row, 4) = arrMNURCLABR(K1) & " : consultation + mise à jour autorisée"
                          wsExcel.Cells(mXls2_Row, K + 4) = "*"
                         wsExcel.Cells(mXls2_Row, K + 4).Interior.Color = mColor_W1
                     Case Is = "3"
                         wsExcel.Cells(mXls2_Row, 4) = arrMNURCLABR(K1) & " : mise à jour autorisée sans consultation"
                         wsExcel.Cells(mXls2_Row, K + 4) = "#"
                         wsExcel.Cells(mXls2_Row, K + 4).Interior.Color = mColor_Y1
                End Select
            End If
        Next K1
        
        rsSab.MoveNext
    Loop
 '___________________________________________________________________________________________________________________
 '   mXls2_Row = mXls2_Row + 1


    X = "select * from " & paramIBM_Library_SABSPE & ".YSSISAB0 " _
         & " where SSISABNAT = 'D' and SSISABUIDX = '" & arrProfil_DX(K) & "'" _
         & " and SSISABSTAK = ' ' order by SSISABUIDX"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        X = Trim(rsSab("SSISABINFO"))
        For K1 = 1 To Len(X) Step 6
            If Mid$(X, K1, 4) <> "    " Then
                mXls2_Row = mXls2_Row + 1
                wsExcel.Cells(mXls2_Row, 1) = arrProfil_DX(K)
                wsExcel.Cells(mXls2_Row, 2) = "Service"
                wsExcel.Cells(mXls2_Row, 3) = K1
                 Select Case Mid$(X, K1 + 4, 1)
                     Case Is = "O"
                         wsExcel.Cells(mXls2_Row, 4) = Mid$(X, K1, 4) & " : mise à jour autorisée"
                         wsExcel.Cells(mXls2_Row, K + 4) = "*"
                         wsExcel.Cells(mXls2_Row, K + 4).Interior.Color = mColor_W1
                    Case Is = "N"
                         wsExcel.Cells(mXls2_Row, 4) = Mid$(X, K1, 4) & " : consultation"
                          wsExcel.Cells(mXls2_Row, K + 4) = "*"
                         wsExcel.Cells(mXls2_Row, K + 4).Interior.Color = mColor_W1
                     Case Else
                         wsExcel.Cells(mXls2_Row, 4) = Mid$(X, K1, 4) & " : ?????????????????"
                         wsExcel.Cells(mXls2_Row, K + 4) = "?"
                         wsExcel.Cells(mXls2_Row, K + 4).Interior.Color = mColor_Y1
                End Select
            End If
        Next K1
        
        rsSab.MoveNext
    Loop
   mXls2_Row = mXls2_Row + 1
Next K
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdProfil_Excel_SAB"

End Sub

Public Sub cmdProfil_Excel_Select()
Set lstW.Container = fraSelect
lstW.ZOrder 0
lstW.Height = fgDetail.Height
lstW.Top = fraSelect.Top + fraSelect.Height - lstW.Height - 1000
lstW.Width = 4000
lstW.Left = fraDetail.Left + fraDetail.Width - lstW.Width - 300
lstW.BackColor = mColor_G0
lstW.Clear
lstW.Visible = True
'cmdProfil_Excel.Caption = "Edition des profils => Excel"
cmdProfil_Excel.BackColor = &H80FF80

End Sub

Public Sub YSSIWIN0_OU_Load()
Dim xSQL As String

blnYSSIWIN0_OU_New = False
xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
     & " where SSIWINNAT = '$'"
Set rsSab = cnsab.Execute(xSQL)
ReDim arrYSSIWIN0_OU(rsSab(0) + 1)

arrYSSIWIN0_OU_Nb = 0

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
     & " where SSIWINNAT = '$'" _
     & " order by SSIWINUIDD"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    arrYSSIWIN0_OU_Nb = arrYSSIWIN0_OU_Nb + 1
    Call rsYSSIWIN0_GetBuffer(rsSab, arrYSSIWIN0_OU(arrYSSIWIN0_OU_Nb))
    rsSab.MoveNext
Loop

End Sub

Public Sub YSSIWIN0_User_Load()
Dim xSQL As String

xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
     & " where SSIWINNAT = ' '"
Set rsSab = cnsab.Execute(xSQL)
ReDim arrYSSIWIN0_User(rsSab(0) + 1)

arrYSSIWIN0_User_Nb = 0

xSQL = "select SSIWINUIDD from " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
     & " where SSIWINNAT = ' ' and SSIWINPRFK <> 'S'" _
     & " order by SSIWINUIDD"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
        arrYSSIWIN0_User_Nb = arrYSSIWIN0_User_Nb + 1
        arrYSSIWIN0_User(arrYSSIWIN0_User_Nb) = rsSab("SSIWINUIDD")
    rsSab.MoveNext
Loop

End Sub

Public Sub cmdSelect_SQL_9_WIN_Control_OU()
Dim xSQL As String, blnUIDX As Boolean, blnGUID As Boolean, X As String

Call cmdUpdate_Init

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
     & " where SSIWINNAT = '$' and SSIWINGUID = '" & xYSSIWIN0.SSIWINGUID & "'" _
     & " order by SSIWINUIDD"
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    xSQL = "select SSIWINUIDD from " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
         & "  Where SSIWINNAT = '$' order by SSIWINUIDD desc FETCH FIRST 1 ROWS ONLY"
    Set rsSab = cnsab.Execute(xSQL)
    
    If rsSab.EOF Then
        mSWIWINUIDD = 0
    Else
        mSWIWINUIDD = rsSab("SSIWINUIDD")
    End If

    mYSSIWIN0_Update = "New"
    mSWIWINUIDD = mSWIWINUIDD + 1
    
    newYSSIWIN0 = xYSSIWIN0
    
    newYSSIWIN0.SSIWINUIDD = mSWIWINUIDD
    newYSSIWIN0.SSIWINUIDX = newYSSIWIN0.SSIWINUIDD & "-" & xYSSIWIN0.SSIWINUIDX
    If Len(newYSSIWIN0.SSIWINUIDX) > 20 Then newYSSIWIN0.SSIWINUIDX = Mid$(newYSSIWIN0.SSIWINUIDX, 1, 20)

    newYSSIWIN0.SSIWINYFCT = "CRE"
    newYSSIWIN0.SSIWINYUSR = usrName_UCase
    newYSSIWIN0.SSIWINYAMJ = DSys
    newYSSIWIN0.SSIWINYHMS = time_Hms
    
    Call cmdSSIJRN_WIN("<X:Création '" & oAD.Class & "'>")
    Call cmdUpdate
Else
    If xYSSIWIN0.SSIWININFO <> Trim(rsSab("SSIWININFO")) Then
        Call rsYSSIWIN0_GetBuffer(rsSab, oldYSSIWIN0)
        mYSSIWIN0_Update = "Update+H"
        oldYSSIWINH = oldYSSIWIN0
        newYSSIWIN0 = oldYSSIWIN0
        newYSSIWIN0.SSIWINUIDX = newYSSIWIN0.SSIWINUIDD & "-" & xYSSIWIN0.SSIWINUIDX
        If Len(newYSSIWIN0.SSIWINUIDX) > 20 Then newYSSIWIN0.SSIWINUIDX = Mid$(newYSSIWIN0.SSIWINUIDX, 1, 20)
        newYSSIWIN0.SSIWINPRFX = xYSSIWIN0.SSIWINPRFX
        newYSSIWIN0.SSIWINUNOM = xYSSIWIN0.SSIWINUNOM
        newYSSIWIN0.SSIWININFO = xYSSIWIN0.SSIWININFO
        newYSSIWIN0.SSIWINYFCT = "CTL"
        newYSSIWIN0.SSIWINYUSR = usrName_UCase
        newYSSIWIN0.SSIWINYAMJ = DSys
        newYSSIWIN0.SSIWINYHMS = time_Hms
        
        Call cmdSSIJRN_WIN("<X:Contrôle 'OrganizationalUnit'>")
        Call cmdUpdate
        
    End If

End If

End Sub
Public Sub cmdSelect_SQL_9_WIN_Control_User()
Dim xSQL As String, blnUIDX As Boolean, blnGUID As Boolean, X As String
Dim blnUpdate As Boolean, K As Long
Call cmdUpdate_Init

Call lstErr_ChangeLastItem(lstErr, cmdContext, "> cmdSelect_SQL_9_WIN_User : " & xYSSIWIN0.SSIWINUIDX): DoEvents

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
     & " where SSIWINNAT = ' ' and SSIWINGUID = '" & xYSSIWIN0.SSIWINGUID & "'" _
     & " order by SSIWINUIDD"
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    xSQL = "select SSIWINUIDD from " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
         & "  Where SSIWINNAT = ' ' order by SSIWINUIDD desc FETCH FIRST 1 ROWS ONLY"
    Set rsSab = cnsab.Execute(xSQL)
    
    If rsSab.EOF Then
        mSWIWINUIDD = 0
    Else
        mSWIWINUIDD = rsSab("SSIWINUIDD")
    End If

    mYSSIWIN0_Update = "New"
    mSWIWINUIDD = mSWIWINUIDD + 1
    
    newYSSIWIN0 = xYSSIWIN0
    newYSSIWIN0.SSIWINPRFK = "?"
    newYSSIWIN0.SSIWINUIDD = mSWIWINUIDD
    newYSSIWIN0.SSIWINYFCT = "CRE"
    newYSSIWIN0.SSIWINYUSR = usrName_UCase
    newYSSIWIN0.SSIWINYAMJ = DSys
    newYSSIWIN0.SSIWINYHMS = time_Hms
    
    Call cmdSSIJRN_WIN("<X:Création '" & oAD.Class & "'>")
    Select Case oAD.Class
        Case "group": cmdSelect_SQL_9_WIN_Control_Class ("BIA_GROUPES")
        Case "computer":
            If InStr(oAD.distinguishedName, "OU=Serveurs BIA") > 0 Then
                cmdSelect_SQL_9_WIN_Control_Class ("WIN_SERVEURS")
            Else
                If InStr(oAD.distinguishedName, "OU=Domain Controllers") > 0 Then
                    cmdSelect_SQL_9_WIN_Control_Class ("WIN_SERVEURS")
                Else
                    cmdSelect_SQL_9_WIN_Control_Class ("WIN_PC")
                End If
            End If
        Case "printQueue": cmdSelect_SQL_9_WIN_Control_Class ("WIN_IMPRIMANTES")
        Case "publicFolder": cmdSelect_SQL_9_WIN_Control_Class ("BIA_FOLDER")
    End Select
    Call cmdUpdate
Else
    For K = 1 To arrYSSIWIN0_User_Nb
        If arrYSSIWIN0_User(K) = rsSab("SSIWINUIDD") Then
            arrYSSIWIN0_User(K) = 0
            Exit For
        End If
    Next K
    If Trim(xYSSIWIN0.SSIWININFO) <> Trim(rsSab("SSIWININFO")) Then
        blnUpdate = True
    Else
        blnUpdate = False
        If xYSSIWIN0.SSIWINPRFK <> rsSab("SSIWINPRFK") _
        And rsSab("SSIWINPRFK") <> "?" Then blnUpdate = True
    End If
    If xYSSIWIN0.SSIWINMAIL <> Trim(rsSab("SSIWINMAIL")) Then blnUpdate = True
   
    If blnUpdate Then
        Call rsYSSIWIN0_GetBuffer(rsSab, oldYSSIWIN0)
        mYSSIWIN0_Update = "Update+H"
        oldYSSIWINH = oldYSSIWIN0
        newYSSIWIN0 = oldYSSIWIN0
        newYSSIWIN0.SSIWINUIDX = xYSSIWIN0.SSIWINUIDX
        newYSSIWIN0.SSIWINUNOM = xYSSIWIN0.SSIWINUNOM
        newYSSIWIN0.SSIWINPRFX = xYSSIWIN0.SSIWINPRFX
        newYSSIWIN0.SSIWININFO = xYSSIWIN0.SSIWININFO
        newYSSIWIN0.SSIWINMAIL = xYSSIWIN0.SSIWINMAIL
        newYSSIWIN0.SSIWINYFCT = "CTL"
        newYSSIWIN0.SSIWINYUSR = usrName_UCase
        newYSSIWIN0.SSIWINYAMJ = DSys
        newYSSIWIN0.SSIWINYHMS = time_Hms
        
        If oldYSSIWIN0.SSIWINPRFK <> "?" Then
            If oldYSSIWIN0.SSIWINPRFK <> xYSSIWIN0.SSIWINPRFK Then
                newYSSIWIN0.SSIWINPRFK = xYSSIWIN0.SSIWINPRFK
                xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0, " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
                    & " Where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN'" _
                    & " and SSIDOMUIDX = '" & oldYSSIWIN0.SSIWINUIDX & "'" _
                    & " and SSIDOMUIDD = " & oldYSSIWIN0.SSIWINUIDD _
                    & " and SSIUSRNAT = ' ' and SSIUSRUIDN = SSIDOMUIDN"

                Set rsSab = cnsab.Execute(xSQL)
            
                If Not rsSab.EOF Then
                    If rsSab("SSIDOMPRFK") <> newYSSIWIN0.SSIWINPRFK Then
                        Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
                        mYSSIDOM0_Update = "Update"
                        newYSSIDOM0 = oldYSSIDOM0
                        newYSSIDOM0.SSIDOMPRFK = newYSSIWIN0.SSIWINPRFK
                        newYSSIDOM0.SSIDOMYFCT = "CTL"
                        newYSSIDOM0.SSIDOMYUSR = usrName_UCase
                        newYSSIDOM0.SSIDOMYAMJ = DSys
                        newYSSIDOM0.SSIDOMYHMS = time_Hms
                        'newYSSIDOM0.SSIDOMPRFD = DSys
                        'newYSSIDOM0.SSIDOMPRFH = time_Hms
                       If rsSab("SSIUSRSTAK") = "N" Then newYSSIDOM0.SSIDOMSTAK = "N"
                    End If
                End If
            End If
        End If
        
        Call cmdSSIJRN_WIN("<X:Contrôle " & oAD.Class & ">")
        Call cmdUpdate
        
    End If

End If

End Sub



Public Sub cmdSelect_SQL_9_WIN_Control_Class(lSSIUSRUIDX)

Dim xSQL As String

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
     & " where SSIUSRNAT = ' ' and SSIUSRUIDX = '" & lSSIUSRUIDX & "'"
Set rsSab_X = cnsab.Execute(xSQL)
If Not rsSab_X.EOF Then
    
    newYSSIWIN0.SSIWINPRFK = " "
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    mYSSIDOM0_Update = "New"
    newYSSIDOM0.SSIDOMNAT = " "
    newYSSIDOM0.SSIDOMSTAK = " "
    newYSSIDOM0.SSIDOMUNIT = " "
    newYSSIDOM0.SSIDOMUIDN = rsSab_X("SSIUSRUIDN")
    newYSSIDOM0.SSIDOMDIDX = "WIN"
    newYSSIDOM0.SSIDOMSTAK = newYSSIWIN0.SSIWINSTAK
    newYSSIDOM0.SSIDOMDECH = 0
    newYSSIDOM0.SSIDOMUIDD = newYSSIWIN0.SSIWINUIDD
    newYSSIDOM0.SSIDOMUIDX = newYSSIWIN0.SSIWINUIDX
    newYSSIDOM0.SSIDOMPRFX = newYSSIWIN0.SSIWINPRFX
    newYSSIDOM0.SSIDOMPRFK = " "
    newYSSIDOM0.SSIDOMPRFD = newYSSIWIN0.SSIWINYAMJ
    newYSSIDOM0.SSIDOMPRFH = newYSSIWIN0.SSIWINYHMS
    newYSSIDOM0.SSIDOMTLNK = 0
    newYSSIDOM0.SSIDOMYFCT = "INI"
    newYSSIDOM0.SSIDOMYAMJ = DSys
    newYSSIDOM0.SSIDOMYHMS = time_Hms
    newYSSIDOM0.SSIDOMYUSR = usrName_UCase
    newYSSIDOM0.SSIDOMYVER = 0
    
End If

End Sub

Public Sub paramUAC_Lib()
arrUAC_Lib(21) = "TRUSTED_TO_AUTH_FOR_DELEGATION "
arrUAC_Lib(20) = "PASSWORD_EXPIRED "
arrUAC_Lib(19) = "DONT_REQ_PREAUTH "
arrUAC_Lib(18) = "USE_DES_KEY_ONLY "
arrUAC_Lib(17) = "NOT_DELEGATED "
arrUAC_Lib(16) = "TRUSTED_FOR_DELEGATION "
arrUAC_Lib(15) = "SMARTCARD_REQUIRED "
arrUAC_Lib(14) = "MNS_LOGON_ACCOUNT "
arrUAC_Lib(13) = "DONT_EXPIRE_PASSWORD "
arrUAC_Lib(12) = "SERVER_TRUST_ACCOUNT "
arrUAC_Lib(11) = "WORKSTATION_TRUST_ACCOUNT "
arrUAC_Lib(10) = "INTERDOMAIN_TRUST_ACCOUNT "
arrUAC_Lib(9) = "NORMAL_ACCOUNT "
arrUAC_Lib(8) = "TEMP_DUPLICATE_ACCOUNT "
arrUAC_Lib(7) = "ENCRYPTED_TEXT_PWD_ALLOWED "
arrUAC_Lib(6) = "PASSWD_CANT_CHANGE "
arrUAC_Lib(5) = "PASSWD_NOTREQD "
arrUAC_Lib(4) = "LOCKOUT "
arrUAC_Lib(3) = "HOMEDIR_REQUIRED "
arrUAC_Lib(2) = "ACCOUNTDISABLE "
arrUAC_Lib(1) = "SCRIPT "

arrUAC_Val(21) = 16777216
arrUAC_Val(20) = 8388608
arrUAC_Val(19) = 4194304
arrUAC_Val(18) = 2097152
arrUAC_Val(17) = 1048576
arrUAC_Val(16) = 524288
arrUAC_Val(15) = 262144
arrUAC_Val(14) = 131072
arrUAC_Val(13) = 65536
arrUAC_Val(12) = 8192
arrUAC_Val(11) = 4096
arrUAC_Val(10) = 2048
arrUAC_Val(9) = 512
arrUAC_Val(8) = 256
arrUAC_Val(7) = 128
arrUAC_Val(6) = 64
arrUAC_Val(5) = 32
arrUAC_Val(4) = 16
arrUAC_Val(3) = 8
arrUAC_Val(2) = 2
arrUAC_Val(1) = 1


End Sub

Public Function cmdSSIWIN_UAC_PRTFK(lSSIWININFO As String) As String
Dim K1 As Integer, K2 As Integer
K1 = InStr(lSSIWININFO, "|") + 1
K2 = InStr(K1, lSSIWININFO, "|")
If Val(Mid$(lSSIWININFO, K1, K2 - K1)) / 2 Mod 2 = 0 Then
    cmdSSIWIN_UAC_PRTFK = " "
Else
    cmdSSIWIN_UAC_PRTFK = "X"
End If

End Function

Public Sub fraYSSIDIV0_Display()

txtSSIDIVUIDX = Trim(oldYSSIDIV0.SSIDIVUIDX)
If Trim(oldYSSIDIV0.SSIDIVUIDX) = "" Then
    txtSSIDIVUIDX.Enabled = True
Else
    txtSSIDIVUIDX.Enabled = False
End If

chkSSIDIVPRFK.Visible = False
Select Case oldYSSIDIV0.SSIDIVPRFK
    Case "X": chkSSIDIVPRFK.Value = "1": chkSSIDIVPRFK.Visible = True
    Case " ": chkSSIDIVPRFK.Value = "0": chkSSIDIVPRFK.Visible = True
End Select

txtSSIDIVUNOM = Trim(oldYSSIDIV0.SSIDIVUNOM)
txtSSIDIVINFO = Trim(oldYSSIDIV0.SSIDIVINFO)

fraYSSIDIV0.Visible = True
End Sub

Public Sub fraDetail_Control_SSIUSRUNIT(lMsgBox As String)

newYSSIUSR0.SSIUSRNAT = "S"
If newYSSIUSR0.SSIUSRUIDN = 0 Then
    mYSSIUSR0_Update = "New"
    newYSSIUSR0.SSIUSRYFCT = "CRE"
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
         & " where SSIUSRNAT = 'S' and SSIUSRUIDN = " & Val(txtSSIUSRUNIT_N)
    
    Set rsSab = cnsab.Execute(X)
    If Not rsSab.EOF Then lMsgBox = lMsgBox & " - Ce numéro service existe déjà" & vbCrLf
Else
    mYSSIUSR0_Update = "Update+H"
    newYSSIUSR0.SSIUSRYFCT = "MOD"
End If

newYSSIUSR0.SSIUSRUIDX = Trim(txtSSIUSRUIDX)
If newYSSIUSR0.SSIUSRUIDX = "" Then lMsgBox = lMsgBox & " - préciser le nom" & vbCrLf

newYSSIUSR0.SSIUSRPRFX = Trim(txtSSIUSRUNIT_X)
If newYSSIUSR0.SSIUSRPRFX = "" Then lMsgBox = lMsgBox & " - préciser le code du service" & vbCrLf

newYSSIUSR0.SSIUSRUIDN = Val(txtSSIUSRUNIT_N)
If newYSSIUSR0.SSIUSRUIDN < 1 Or newYSSIUSR0.SSIUSRUIDN > 99 Then lMsgBox = lMsgBox & " - le code du service est compris entre 1 et 99" & vbCrLf

newYSSITXT0.SSITXTNAT = newYSSIUSR0.SSIUSRNAT
newYSSITXT0.SSITXTUIDN = newYSSIUSR0.SSIUSRUIDN



End Sub

Public Sub cboSSIUSRUNIT_Load()
Dim xSQL As String, K As Integer
For K = 0 To 99: arrSSIUSRUNIT_Code(K) = "": Next K
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
     & " where SSIUSRNAT = 'S'" _
     & " order by SSIUSRUIDN"

Set rsSab = cnsab.Execute(xSQL)
cboSSIUSRUNIT.Clear
cboSSIUSRUNIT.AddItem ""
cboSSIDOMUNIT.Clear
cboSSIDOMUNIT.AddItem ""
Do While Not rsSab.EOF
    If Trim(rsSab("SSIUSRSTAK")) = "" Then
        cboSSIUSRUNIT.AddItem Trim(rsSab("SSIUSRUNIT")) & "-" & Trim(rsSab("SSIUSRUIDX"))
        cboSSIDOMUNIT.AddItem Trim(rsSab("SSIUSRUNIT")) & "-" & Trim(rsSab("SSIUSRUIDX"))
    End If
    arrSSIUSRUNIT_Code(rsSab("SSIUSRUIDN")) = Trim(rsSab("SSIUSRPRFX"))
    rsSab.MoveNext
Loop

End Sub

Public Sub cmdSelect_SQL_9_MEL_User()
Dim xInfo As String, xSQL As String

'If InStr(usrYSSIMEL0.SSIMELUNOM, "oulerg") > 0 Then
'    Debug.Print "cmdSelect_SQL_9_MEL_User"
'End If
usrYSSIMEL0.SSIMELPRFK = "N"
If InStr(usrYSSIMEL0.SSIMELUNOM, "@bia-paris.fr") > 0 Then usrYSSIMEL0.SSIMELPRFK = " "
If InStr(usrYSSIMEL0.SSIMELUNOM, "@bia-paris.int") > 0 Then usrYSSIMEL0.SSIMELPRFK = " "
If InStr(usrYSSIMEL0.SSIMELUNOM, "@bia-paris.xx") > 0 Then usrYSSIMEL0.SSIMELPRFK = "X"
If usrYSSIMEL0.SSIMELUNOM = "" Then usrYSSIMEL0.SSIMELPRFK = "X"

xInfo = Trim(olExchangeUser.PrimarySmtpAddress) & "|" _
      & Trim(olExchangeUser.Alias) & "|" _
      & Trim(olExchangeUser.LastName) & "|" _
      & Trim(olExchangeUser.FirstName) & "|" _
      & Trim(olExchangeUser.Name) & "|" _
      & Trim(olExchangeUser.JobTitle) & "|" _
      & Trim(olExchangeUser.OfficeLocation) & "|" _
      & Trim(olExchangeUser.BusinessTelephoneNumber) & "|" _
      & Trim(olExchangeUser.CompanyName) & "|" _
      & Trim(olExchangeUser.Department) & "|" _
      & Trim(olExchangeUser.MobileTelephoneNumber) & "|" _
      & Trim(olExchangeUser.PostalCode) & "|" _
      & Trim(olExchangeUser.StreetAddress) & "|" _
      & Trim(olExchangeUser.Comments) & "|" _
      & Trim(olExchangeUser.Address) & "|" _
      & Trim(olExchangeUser.Id) & "|"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
     & " where SSIMELNAT = ' ' and SSIMELUIDX = '" & usrYSSIMEL0.SSIMELUIDX & "' and SSIMELUIDD = " & usrYSSIMEL0.SSIMELUIDD
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    If usrYSSIMEL0.SSIMELUNOM <> "" Then
        newYSSIMEL0 = usrYSSIMEL0
        newYSSIMEL0.SSIMELINFO = xInfo
        Call cmdUpdate_Init: mYSSIMEL0_Update = "New"
        Call cmdSSIJRN_MEL("<X:Exchange groupe : " & usrYSSIMEL0.SSIMELPRFX & ">")
        Call cmdUpdate
    End If
Else
    If xInfo = Trim(rsSab("SSIMELINFO")) Then
        mYSSIMEL0_Update = ""
    Else
        Call rsYSSIMEL0_GetBuffer(rsSab, oldYSSIMEL0)
        newYSSIMEL0 = oldYSSIMEL0
        newYSSIMEL0.SSIMELUIDX = usrYSSIMEL0.SSIMELUIDX
        newYSSIMEL0.SSIMELPRFX = usrYSSIMEL0.SSIMELPRFX
        newYSSIMEL0.SSIMELPRFK = usrYSSIMEL0.SSIMELPRFK
        newYSSIMEL0.SSIMELINFO = xInfo
        newYSSIMEL0.SSIMELYFCT = "MOD"
        newYSSIMEL0.SSIMELYUSR = usrName_UCase
        newYSSIMEL0.SSIMELYAMJ = mImport_PRFD
        newYSSIMEL0.SSIMELYHMS = mImport_PRFH
        Call cmdUpdate_Init: mYSSIMEL0_Update = "Update+H"
        Call cmdSSIJRN_MEL("<X:Exchange groupe : " & usrYSSIMEL0.SSIMELPRFX & ">")
        Call cmdUpdate
    End If
    
End If

'___________________________________________________________________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
& " where SSIDOMNAT = ' '" _
& " and SSIDOMDIDX = 'MEL'" _
& " and SSIDOMUIDX = '" & usrYSSIMEL0.SSIMELUIDX & "'" _
& " and SSIDOMUIDD = " & usrYSSIMEL0.SSIMELUIDD

Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    If rsSab("SSIDOMPRFK") = usrYSSIMEL0.SSIMELPRFK And rsSab("SSIDOMPRFD") = mImport_PRFD Then
    Else
        Call rsYSSIDOM0_GetBuffer(rsSab, oldYSSIDOM0)
        newYSSIDOM0 = oldYSSIDOM0
        newYSSIDOM0.SSIDOMPRFK = usrYSSIMEL0.SSIMELPRFK
        newYSSIDOM0.SSIDOMPRFD = mImport_PRFD
        newYSSIDOM0.SSIDOMPRFH = mImport_PRFH
        mYSSIDOM0_Update = "Update"
        Call cmdUpdate
    End If
Else
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 " _
    & " where SSIDOMNAT = ' '" _
    & " and SSIDOMDIDX = 'WIN'" _
    & " and SSIDOMUIDD = " & usrYSSIMEL0.SSIMELUIDD
         
    Set rsSab = cnsab.Execute(xSQL)
    If rsSab.EOF Then
        Call cmdUpdate_Init
        newYSSIMEL0 = usrYSSIMEL0
        Call cmdSSIJRN_TXT_Once("MEL", "<ORIG:43><FCT:9-???><UID:" & usrYSSIMEL0.SSIMELUNOM & "><X:compte Exchange non affecté WIN>")
        Call cmdUpdate
    Else
        If usrYSSIMEL0.SSIMELUNOM <> "" Then

            newYSSIDOM0 = xYSSIDOM0
            newYSSIDOM0.SSIDOMUIDN = rsSab("SSIDOMUIDN")
            newYSSIDOM0.SSIDOMUIDX = usrYSSIMEL0.SSIMELUIDX
            newYSSIDOM0.SSIDOMUIDD = usrYSSIMEL0.SSIMELUIDD
            newYSSIDOM0.SSIDOMUNIT = rsSab("SSIDOMUNIT")
            newYSSIDOM0.SSIDOMSTAK = rsSab("SSIDOMSTAK")
            newYSSIDOM0.SSIDOMDECH = rsSab("SSIDOMDECH")
            newYSSIDOM0.SSIDOMPRFX = usrYSSIMEL0.SSIMELPRFX
            newYSSIDOM0.SSIDOMPRFK = usrYSSIMEL0.SSIMELPRFK
            newYSSIDOM0.SSIDOMPRFD = mImport_PRFD
            newYSSIDOM0.SSIDOMPRFH = mImport_PRFH
            mYSSIDOM0_Update = "New"
           Call cmdSSIJRN_DOM("")
            Call cmdSSIJRN_DOM("<X:Exchange groupe : " & usrYSSIMEL0.SSIMELPRFX & ">")
           Call cmdUpdate
        End If
    End If
    
End If
                    

End Sub

Public Sub lstParam_SSIMELUNOM_Load(lWhere As String)
Dim xSQL As String

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
     & " where SSIMELNAT = '@' " & lWhere _
     & " order by SSIMELUIDX"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    lstParam_SSIMELUIDX.AddItem rsSab("SSIMELUIDX") & " | " & Trim(rsSab("SSIMELUNOM"))
    rsSab.MoveNext
Loop
'______________________________________________________________________________________________
xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
     & " where SSIMELNAT in ('$',' ')"
Set rsSab = cnsab.Execute(xSQL)
arrSSIMELUNOM_Nb = rsSab(0) + 1
ReDim arrSSIMELUIDX(arrSSIMELUNOM_Nb), arrSSIMELUNOM(arrSSIMELUNOM_Nb), blnSSIMELUNOM(arrSSIMELUNOM_Nb)


arrSSIMELUNOM_Nb = 0

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
     & " where SSIMELNAT = '$'  order by  SSIMELUIDX"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    arrSSIMELUNOM_Nb = arrSSIMELUNOM_Nb + 1
    arrSSIMELUIDX(arrSSIMELUNOM_Nb) = rsSab("SSIMELUIDX")
    arrSSIMELUNOM(arrSSIMELUNOM_Nb) = StrConv(Trim(rsSab("SSIMELUNOM")), vbProperCase)
    blnSSIMELUNOM(arrSSIMELUNOM_Nb) = False
    rsSab.MoveNext
Loop

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
     & " where SSIMELNAT = ' ' and SSIMELPRFX = 'Exchange' order by  SSIMELUIDX"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    arrSSIMELUNOM_Nb = arrSSIMELUNOM_Nb + 1
    arrSSIMELUIDX(arrSSIMELUNOM_Nb) = Trim(rsSab("SSIMELUIDX"))
    arrSSIMELUNOM(arrSSIMELUNOM_Nb) = StrConv(Trim(rsSab("SSIMELUNOM")), vbProperCase)
    blnSSIMELUNOM(arrSSIMELUNOM_Nb) = False
    rsSab.MoveNext
Loop

End Sub

Public Sub lstParam_SSIMELUNOM_Display()
Dim I As Integer
fraParam_SSIMELUIDX.Visible = False
txtParam_SSIMELINFO = ""
lstParam_SSIMELUNOM.Clear
For I = 1 To arrSSIMELUNOM_Nb
    If blnSSIMELUNOM(I) Then
        lstParam_SSIMELUNOM.AddItem arrSSIMELUNOM(I)
        lstParam_SSIMELUNOM.Selected(lstParam_SSIMELUNOM.ListCount - 1) = True
        If txtParam_SSIMELINFO = "" Then
            txtParam_SSIMELINFO = arrSSIMELUNOM(I)
        Else
            txtParam_SSIMELINFO = txtParam_SSIMELINFO & ";" & arrSSIMELUNOM(I)
        End If
    End If
Next I
For I = 1 To arrSSIMELUNOM_Nb
    If Not blnSSIMELUNOM(I) Then
        lstParam_SSIMELUNOM.AddItem arrSSIMELUNOM(I)
    End If
Next I
lstParam_SSIMELUNOM.ListIndex = 0
fraParam_SSIMELUIDX.Visible = True
End Sub

Public Sub fgSelect_Display_3_MEL()

Dim K As Long, xSQL As String, X As String, xUsr As String
Dim kLen As Integer, I As Integer, K1 As Integer, blnOk As Boolean, blnExit As Boolean
Dim Nb As Integer
On Error GoTo Error_Handler
currentAction = currentAction & "-> fgSelect_Display_3_MEL"


Do While Not rsSab.EOF
    xUsr = Trim(rsSab("SSIMELINFO"))
    kLen = Len(xUsr)
    K1 = 1
    blnExit = False
    Do
        K = InStr(K1, xUsr, ";")
        If K > 0 Then
            blnOk = False
            X = StrConv(Trim(Mid$(xUsr, K1, K - K1)), vbProperCase)
            For I = 1 To arrSSIMELUNOM_Nb
                If X = arrSSIMELUNOM(I) Then blnOk = True: Exit For
            Next I
            K1 = K + 1
        Else
            X = StrConv(Trim(Mid$(xUsr, K1, kLen - K1 + 1)), vbProperCase)
            For I = 1 To arrSSIMELUNOM_Nb
                If X = arrSSIMELUNOM(I) Then blnOk = True: Exit For
            Next I
            blnExit = True 'Exit Do
        End If
        If blnOk = False Then
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            Nb = Nb + 1
            fgSelect.Col = 0: fgSelect.Text = "paramétrage"
            fgSelect.Col = 1: fgSelect.Text = "MEL"
            fgSelect.Col = 2: fgSelect.Text = X: fgSelect.CellBackColor = mColor_W0
            fgSelect.Col = 3: fgSelect.Text = rsSab("SSIMELUIDX"): fgSelect.CellBackColor = mColor_W0
            fgSelect.Col = 6: fgSelect.Text = rsSab("SSIMELUNOM")
            fgSelect.Col = 5: fgSelect.Text = "adresse mail inconnue": fgSelect.CellBackColor = mColor_Y0
            fgSelect.Col = 4: fgSelect.Text = "destinataire"
        End If
    Loop Until blnExit
    
    rsSab.MoveNext
Loop

arrCtl_Nb(arrCtl_K) = Nb


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction




End Sub

Public Function cmdSelect_SQL_9_MEL_SSIMELUIDX()
Dim wSSIMELUIDX As String
wSSIMELUIDX = Trim(olExchangeUser.LastName)
If wSSIMELUIDX = "" Then wSSIMELUIDX = Trim(olExchangeUser.Alias)
If wSSIMELUIDX = "" Then wSSIMELUIDX = Trim(olExchangeUser.FirstName)
If wSSIMELUIDX = "" Then wSSIMELUIDX = Trim(olExchangeUser.Name)

cmdSelect_SQL_9_MEL_SSIMELUIDX = wSSIMELUIDX
End Function

Public Sub JPL_DS()
Dim K As Integer
Call DS_Server_Open

X = InputBox("collection " & paramDocuShare_Collection_SI_Doc & " : Service informatique documentation")
'    & vbCrLf & "     =========================" & vbCrLf & paramDocuShare_Collection_SI_Doc _
'    & vbCrLf & "     =========================", "Document ", "")
If Trim(X) = "" Then GoTo Exit_sub

'paramDocuShare_Collection_SI_Doc =  2858'X
Call DS_Document_Load("*" & X & "*", paramDocuShare_Collection_SI_Doc)

'For K = 2000 To 2012
'    Call lstErr_Clear(lstErr, cmdContext, "DS_Document_Load  " & Str(K)): DoEvents
'   Call DS_Document_Load("J*", paramDocuShare_Collection_SI_Doc)
'   Call DS_Document_Load("L*", paramDocuShare_Collection_SI_Doc)
'   Call DS_Document_Load("M*", paramDocuShare_Collection_SI_Doc)
'   Call DS_Document_Load("N*", paramDocuShare_Collection_SI_Doc)
'   Call DS_Document_Load("P*", paramDocuShare_Collection_SI_Doc)
'   Call DS_Document_Load("R*", paramDocuShare_Collection_SI_Doc)
'   Call DS_Document_Load("S*", paramDocuShare_Collection_SI_Doc)
'   Call DS_Document_Load("T*", paramDocuShare_Collection_SI_Doc)
'Next K

Exit_sub:

End Sub
