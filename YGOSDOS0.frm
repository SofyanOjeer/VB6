VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmYGOSDOS0 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E6E6E6&
   Caption         =   "Gestion des Opérations en Suspens"
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
   Icon            =   "YGOSDOS0.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10305
   ScaleWidth      =   13530
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   585
      Left            =   4815
      TabIndex        =   217
      Top             =   1245
      Visible         =   0   'False
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   1032
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdMail_MT 
      Height          =   375
      Left            =   12510
      Picture         =   "YGOSDOS0.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   45
      Visible         =   0   'False
      Width           =   500
   End
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
      Tab             =   2
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
      TabPicture(0)   =   "YGOSDOS0.frx":0614
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Paramétrage"
      TabPicture(1)   =   "YGOSDOS0.frx":0630
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tabParam"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "."
      TabPicture(2)   =   "YGOSDOS0.frx":064C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   9285
         Left            =   -15
         TabIndex        =   63
         Top             =   390
         Visible         =   0   'False
         Width           =   13425
         _ExtentX        =   23680
         _ExtentY        =   16378
         _Version        =   393216
         Tabs            =   10
         Tab             =   2
         TabsPerRow      =   10
         TabHeight       =   520
         TabCaption(0)   =   "fraSelect"
         TabPicture(0)   =   "YGOSDOS0.frx":0668
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fraSelect_Options_7"
         Tab(0).Control(1)=   "fraSelect_Options_Stat"
         Tab(0).Control(2)=   "fraSelect_Options_4"
         Tab(0).Control(3)=   "lstW"
         Tab(0).Control(4)=   "fraSelect_Options_1a"
         Tab(0).Control(5)=   "fraSelect_Options_1b"
         Tab(0).Control(6)=   "fraSelect_Options_3"
         Tab(0).Control(7)=   "fraSelect_Options_9"
         Tab(0).Control(8)=   "txtFg"
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "-"
         TabPicture(1)   =   "YGOSDOS0.frx":0684
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "fraMail_MT"
         TabPicture(2)   =   "YGOSDOS0.frx":06A0
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "fraMail_MT"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "fraSwift"
         TabPicture(3)   =   "YGOSDOS0.frx":06BC
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "fraSWISABKSRV"
         Tab(3).Control(1)=   "fraSwift"
         Tab(3).Control(2)=   "fgFree"
         Tab(3).ControlCount=   3
         TabCaption(4)   =   "fraDetail"
         TabPicture(4)   =   "YGOSDOS0.frx":06D8
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "fraDetail"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "fraDetail (suite)"
         TabPicture(5)   =   "YGOSDOS0.frx":06F4
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "fraEVE"
         Tab(5).Control(1)=   "fraDetail_C"
         Tab(5).ControlCount=   2
         TabCaption(6)   =   "fraEVE_Swift"
         TabPicture(6)   =   "YGOSDOS0.frx":0710
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "fraEVE_Swift"
         Tab(6).Control(1)=   "fgModèle"
         Tab(6).ControlCount=   2
         TabCaption(7)   =   "fraPJ"
         TabPicture(7)   =   "YGOSDOS0.frx":072C
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "fraPJ"
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "Tab 8"
         TabPicture(8)   =   "YGOSDOS0.frx":0748
         Tab(8).ControlEnabled=   0   'False
         Tab(8).ControlCount=   0
         TabCaption(9)   =   "Tab 9"
         TabPicture(9)   =   "YGOSDOS0.frx":0764
         Tab(9).ControlEnabled=   0   'False
         Tab(9).ControlCount=   0
         Begin VB.Frame fraSelect_Options_7 
            BackColor       =   &H00F0FFFF&
            BorderStyle     =   0  'None
            Height          =   1260
            Left            =   -74715
            TabIndex        =   234
            Top             =   6360
            Visible         =   0   'False
            Width           =   9375
            Begin VB.ComboBox cboSelect_7_SRV 
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
               Left            =   435
               Sorted          =   -1  'True
               TabIndex        =   235
               Top             =   600
               Width           =   2355
            End
         End
         Begin VB.Frame fraSelect_Options_Stat 
            BackColor       =   &H00F0FFFF&
            BorderStyle     =   0  'None
            Height          =   1260
            Left            =   -74715
            TabIndex        =   222
            Top             =   4950
            Visible         =   0   'False
            Width           =   9375
            Begin MSComCtl2.DTPicker txtSelect_Stat_AMJMin 
               Height          =   300
               Left            =   2730
               TabIndex        =   223
               Top             =   615
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
               Format          =   92930051
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_Stat_AMJMax 
               Height          =   300
               Left            =   4500
               TabIndex        =   225
               Top             =   630
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
               Format          =   92930051
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_Stat_AMJMin 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Période du ...... au ......"
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
               TabIndex        =   224
               Top             =   615
               Width           =   2220
            End
         End
         Begin VB.Frame fraSelect_Options_4 
            BackColor       =   &H00F0FFFF&
            BorderStyle     =   0  'None
            Height          =   1260
            Left            =   -74790
            TabIndex        =   201
            Top             =   3480
            Visible         =   0   'False
            Width           =   9375
            Begin VB.ComboBox cboSelect_4_GOSDOSGSRV 
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
               Left            =   3435
               Sorted          =   -1  'True
               TabIndex        =   203
               Top             =   585
               Width           =   2355
            End
            Begin VB.ComboBox cboSelect_4_GOSDOSISRV 
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
               Left            =   435
               Sorted          =   -1  'True
               TabIndex        =   202
               Top             =   600
               Width           =   2355
            End
            Begin MSComCtl2.DTPicker txtSelect_4_GOSDOSECHD 
               Height          =   300
               Left            =   7065
               TabIndex        =   206
               Top             =   570
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
               Format          =   92930051
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_4_GOSDOSECHD 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Date échéance <= au"
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
               Left            =   6825
               TabIndex        =   207
               Top             =   240
               Width           =   2220
            End
            Begin VB.Label lblSelect_4_GOSDOSGSRV 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Service gestionnaire"
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
               Left            =   3495
               TabIndex        =   205
               Top             =   195
               Width           =   2625
            End
            Begin VB.Label lblSelect_4_GOSDOSISRV 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Service initiateur"
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
               Left            =   495
               TabIndex        =   204
               Top             =   180
               Width           =   2460
            End
         End
         Begin VB.Frame fraEVE_Swift 
            BackColor       =   &H00D8DFD8&
            Caption         =   "En-tête du message Swift"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4560
            Left            =   -74730
            TabIndex        =   189
            Top             =   570
            Width           =   8190
            Begin VB.ComboBox cboEVE_Swift_20 
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
               Left            =   2280
               TabIndex        =   228
               Text            =   "cboEVE_Swift_20"
               Top             =   2250
               Width           =   2745
            End
            Begin VB.ComboBox cboEVE_Swift_MT 
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
               Left            =   2300
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   193
               Top             =   840
               Width           =   1620
            End
            Begin VB.ComboBox cboEVE_Swift_BIC 
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
               Left            =   2300
               Sorted          =   -1  'True
               TabIndex        =   192
               Text            =   "cboEVE_Swift_BIC"
               Top             =   1605
               Width           =   2715
            End
            Begin VB.TextBox txtEVE_Swift_21 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Left            =   2300
               MaxLength       =   16
               TabIndex        =   191
               Top             =   3075
               Width           =   2700
            End
            Begin VB.CommandButton cmdEVE_Swift_Ok 
               BackColor       =   &H0080FF80&
               Caption         =   "Suite (champ 79)"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   885
               Left            =   6180
               Style           =   1  'Graphical
               TabIndex        =   190
               Top             =   2385
               Width           =   1400
            End
            Begin VB.Label lblEVE_Swift_MT 
               BackColor       =   &H00D8DFD8&
               Caption         =   "Type de message"
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
               Left            =   255
               TabIndex        =   197
               Top             =   915
               Width           =   1815
            End
            Begin VB.Label lblEVE_Swift_BIC 
               BackColor       =   &H00D8DFD8&
               Caption         =   "BIC du destinataire"
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
               Left            =   195
               TabIndex        =   196
               Top             =   1590
               Width           =   1860
            End
            Begin VB.Label lblEVE_Swift_20 
               BackColor       =   &H00D8DFD8&
               Caption         =   "Champ 20"
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
               Left            =   285
               TabIndex        =   195
               Top             =   2325
               Width           =   1605
            End
            Begin VB.Label lblEVE_Swift_21 
               BackColor       =   &H00D8DFD8&
               Caption         =   "Champ 21"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   240
               TabIndex        =   194
               Top             =   3090
               Width           =   1455
            End
         End
         Begin VB.Frame fraEVE 
            BackColor       =   &H00D8DFD8&
            Height          =   7000
            Left            =   -70200
            TabIndex        =   178
            Top             =   765
            Visible         =   0   'False
            Width           =   8730
            Begin VB.CommandButton cmdEVE_Dupliquer 
               BackColor       =   &H0000FFFF&
               Caption         =   "Dupliquer le Message"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   780
               Left            =   4935
               Style           =   1  'Graphical
               TabIndex        =   216
               Top             =   4095
               Visible         =   0   'False
               Width           =   1400
            End
            Begin VB.CommandButton cmdEVE_Quit 
               BackColor       =   &H00C0C0FF&
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
               Height          =   825
               Left            =   6570
               Style           =   1  'Graphical
               TabIndex        =   187
               Top             =   5025
               Width           =   1400
            End
            Begin VB.Frame fraEVE_S 
               BackColor       =   &H00E0FFFF&
               Height          =   6060
               Left            =   150
               TabIndex        =   179
               Top             =   240
               Width           =   8400
               Begin VB.CommandButton cmdEVE_Ignore 
                  BackColor       =   &H008080FF&
                  Caption         =   "Annulation de l'événement"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   870
                  Left            =   4845
                  Style           =   1  'Graphical
                  TabIndex        =   180
                  Top             =   1845
                  Visible         =   0   'False
                  Width           =   1400
               End
               Begin VB.CommandButton cmdEVE_Ok_àClôturer 
                  BackColor       =   &H0080C0FF&
                  Caption         =   "Enregistrer + dossier à clôturer"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   780
                  Left            =   6270
                  Style           =   1  'Graphical
                  TabIndex        =   212
                  Top             =   2835
                  Visible         =   0   'False
                  Width           =   1400
               End
               Begin VB.CommandButton cmdEVE_Ok_Clôture 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Enregistrer + Clôture du dossier"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   780
                  Left            =   6285
                  Style           =   1  'Graphical
                  TabIndex        =   211
                  Top             =   1890
                  Visible         =   0   'False
                  Width           =   1400
               End
               Begin VB.CommandButton cmdEVE_Ok 
                  BackColor       =   &H0000FF00&
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
                  Height          =   780
                  Left            =   6255
                  Style           =   1  'Graphical
                  TabIndex        =   185
                  Top             =   3825
                  Visible         =   0   'False
                  Width           =   1400
               End
               Begin VB.CheckBox chkGOSEVEGSRV 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Service en charge de la prochaine intervention"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   435
                  Left            =   5745
                  TabIndex        =   184
                  Top             =   240
                  Width           =   2415
               End
               Begin VB.ComboBox cboGOSEVENAT 
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
                  Left            =   2925
                  Style           =   2  'Dropdown List
                  TabIndex        =   183
                  Top             =   225
                  Width           =   2535
               End
               Begin VB.ComboBox cboGOSEVEGSRV 
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
                  Left            =   5835
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   182
                  Top             =   885
                  Width           =   2400
               End
               Begin VB.TextBox txtGOSEVETXT 
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   5100
                  Left            =   300
                  MultiLine       =   -1  'True
                  TabIndex        =   181
                  Top             =   800
                  Width           =   5355
               End
               Begin MSComCtl2.DTPicker txtGOSEVEECHD 
                  Height          =   300
                  Left            =   6990
                  TabIndex        =   200
                  Top             =   1380
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
                  Format          =   92930051
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin VB.Label lblGOSEVEECHD 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "  Echéance"
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
                  Left            =   5820
                  TabIndex        =   199
                  Top             =   1365
                  Width           =   2450
               End
               Begin VB.Label lblGOSEVENAT 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Nature de l'événement / action"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Left            =   300
                  TabIndex        =   186
                  Top             =   210
                  Width           =   2475
               End
            End
            Begin VB.Label lblGOSEVEUAMJ 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "maj par le"
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
               Left            =   150
               TabIndex        =   188
               Top             =   6420
               Width           =   8400
            End
         End
         Begin VB.Frame fraDetail_C 
            BackColor       =   &H00E0FFFF&
            Height          =   7100
            Left            =   -74940
            TabIndex        =   156
            Top             =   945
            Visible         =   0   'False
            Width           =   5310
            Begin VB.Frame fraDetail_LAB 
               BackColor       =   &H00F0FFFF&
               Height          =   5730
               Left            =   150
               TabIndex        =   159
               Top             =   630
               Width           =   5000
               Begin VB.CommandButton cmdDetail_Lab_Link 
                  BackColor       =   &H0080C0FF&
                  Caption         =   "associer à un dossier GOS"
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
                  Left            =   240
                  Style           =   1  'Graphical
                  TabIndex        =   219
                  Top             =   5115
                  Width           =   1440
               End
               Begin VB.Frame fraDetail_EVE 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Tâche initiale à réaliser par le service"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2745
                  Left            =   75
                  TabIndex        =   165
                  Top             =   2355
                  Width           =   4875
                  Begin VB.ComboBox cboGOSDOSGSRV 
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
                     Left            =   60
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   167
                     Top             =   390
                     Width           =   2400
                  End
                  Begin VB.TextBox txtGOSDOSTXT 
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1785
                     Left            =   60
                     MultiLine       =   -1  'True
                     TabIndex        =   166
                     Top             =   795
                     Width           =   4740
                  End
                  Begin MSComCtl2.DTPicker txtGOSDOSECHD 
                     Height          =   300
                     Left            =   3465
                     TabIndex        =   168
                     Top             =   375
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
                     Format          =   92930051
                     CurrentDate     =   38699.44875
                     MaxDate         =   401768
                     MinDate         =   36526.4425347222
                  End
                  Begin VB.Label lblGOSDOSECHD 
                     BackColor       =   &H00E0FFFF&
                     Caption         =   "Echéance"
                     BeginProperty Font 
                        Name            =   "Arial Unicode MS"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Left            =   2535
                     TabIndex        =   169
                     Top             =   420
                     Width           =   915
                  End
               End
               Begin VB.TextBox txtGOSDOSCLI 
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
                  Left            =   1000
                  TabIndex        =   164
                  Top             =   840
                  Width           =   900
               End
               Begin VB.ComboBox cboGOSDOSPAYS 
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
                  Left            =   1000
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   163
                  Top             =   1800
                  Width           =   2325
               End
               Begin VB.ComboBox cboGOSDOSLABK 
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
                  Left            =   1035
                  Style           =   2  'Dropdown List
                  TabIndex        =   162
                  Top             =   360
                  Width           =   2085
               End
               Begin VB.CommandButton cmdDetail_LAB_Ok 
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
                  Height          =   555
                  Left            =   3735
                  Style           =   1  'Graphical
                  TabIndex        =   161
                  Top             =   5115
                  Width           =   1212
               End
               Begin VB.ComboBox cboGOSDOSRCOM 
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
                  Left            =   1000
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   160
                  Top             =   1305
                  Width           =   2370
               End
               Begin VB.Label libGOSDOSLABK 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "libGOSDOSLABK"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   3315
                  TabIndex        =   175
                  Top             =   420
                  Visible         =   0   'False
                  Width           =   2580
               End
               Begin VB.Label libGOSDOSCLI 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "libGOSDOSCLI"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   255
                  Left            =   2055
                  TabIndex        =   174
                  Top             =   900
                  Width           =   2580
               End
               Begin VB.Label lblGOSDOSPAYS 
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
                  Height          =   255
                  Left            =   180
                  TabIndex        =   173
                  Top             =   1800
                  Width           =   960
               End
               Begin VB.Label lblGOSDOSLABK 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Motif"
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
                  Left            =   195
                  TabIndex        =   172
                  Top             =   420
                  Width           =   750
               End
               Begin VB.Label lblGOSDOSCLI 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Client"
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
                  Left            =   180
                  TabIndex        =   171
                  Top             =   900
                  Width           =   750
               End
               Begin VB.Label lblGOSDOSRCOM 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Resp com"
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
                  Left            =   180
                  TabIndex        =   170
                  Top             =   1425
                  Width           =   750
               End
            End
            Begin VB.Frame fraDetail_Y 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   0  'None
               Height          =   492
               Left            =   120
               TabIndex        =   157
               Top             =   195
               Width           =   4995
               Begin VB.Label lblGOSDOSSTAD 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "lblGOSDOSSTAD"
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
                  Left            =   105
                  TabIndex        =   158
                  Top             =   90
                  Width           =   4800
               End
            End
            Begin VB.Label lblGOSDOSUAMJ 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "maj par le"
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
               Left            =   150
               TabIndex        =   177
               Top             =   6705
               Width           =   4995
            End
            Begin VB.Label lblGOSDOSIAMJ 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "créé par le"
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
               Left            =   135
               TabIndex        =   176
               Top             =   6375
               Width           =   4995
            End
         End
         Begin VB.Frame fraPJ 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Pièce Jointe"
            Height          =   5850
            Left            =   -74400
            TabIndex        =   148
            Top             =   1065
            Visible         =   0   'False
            Width           =   9975
            Begin RichTextLib.RichTextBox rtfPJ 
               Height          =   2025
               Left            =   4515
               TabIndex        =   153
               TabStop         =   0   'False
               Top             =   3400
               Width           =   5100
               _ExtentX        =   8996
               _ExtentY        =   3572
               _Version        =   393217
               BackColor       =   14737632
               HideSelection   =   0   'False
               ScrollBars      =   3
               AutoVerbMenu    =   -1  'True
               TextRTF         =   $"YGOSDOS0.frx":0780
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
            Begin VB.CommandButton cmdGOSDOSPJ 
               BackColor       =   &H000080FF&
               Caption         =   "Mémoriser le chemin d'accès au répertoire"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   900
               Style           =   1  'Graphical
               TabIndex        =   152
               Top             =   3300
               Width           =   2625
            End
            Begin VB.DirListBox dirListBox 
               Height          =   2115
               Left            =   120
               TabIndex        =   151
               Top             =   960
               Width           =   4000
            End
            Begin VB.DriveListBox DriveListBox 
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
               Left            =   120
               TabIndex        =   150
               Top             =   360
               Width           =   4000
            End
            Begin VB.FileListBox filDoc 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   2820
               Left            =   4455
               Pattern         =   "*.doc;*.pdf;*.rtf;*.xls;*.txt"
               TabIndex        =   149
               Top             =   360
               Width           =   5235
            End
            Begin VB.Label librtfPJ 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Click droit pour copier/coller ==>"
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
               Left            =   1125
               TabIndex        =   154
               Top             =   4290
               Width           =   2865
            End
         End
         Begin VB.Frame fraSWISABKSRV 
            BackColor       =   &H00FFE0FF&
            Caption         =   "Modification de l'affectation d'un dossier"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   3960
            Left            =   -67545
            TabIndex        =   136
            Top             =   630
            Visible         =   0   'False
            Width           =   4500
            Begin VB.ComboBox cboSWISABKSRV 
               Height          =   315
               Left            =   2600
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   142
               Top             =   990
               Width           =   1200
            End
            Begin VB.ComboBox cboSWISABSER 
               Height          =   315
               Left            =   2600
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   141
               Top             =   1590
               Width           =   1200
            End
            Begin VB.ComboBox cboSWISABOPEC 
               Height          =   315
               Left            =   2600
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   140
               Top             =   2115
               Width           =   1200
            End
            Begin VB.TextBox txtSWISABOPEN 
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
               Left            =   2600
               TabIndex        =   139
               Top             =   2655
               Width           =   1200
            End
            Begin VB.CommandButton cmdSWISABKSRV_Quit 
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
               Left            =   495
               Style           =   1  'Graphical
               TabIndex        =   138
               Top             =   3195
               Width           =   1155
            End
            Begin VB.CommandButton cmdSWISABKSRV__Update 
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
               Left            =   2565
               Style           =   1  'Graphical
               TabIndex        =   137
               Top             =   3210
               Width           =   1170
            End
            Begin VB.Label lblSWISABKSRV 
               BackColor       =   &H00FFE0FF&
               Caption         =   "Service BIA"
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
               Left            =   495
               TabIndex        =   147
               Top             =   1080
               Width           =   1995
            End
            Begin VB.Label lblSWISABSER 
               BackColor       =   &H00FFE0FF&
               Caption         =   "Service/ sous service SAB"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   495
               TabIndex        =   146
               Top             =   1650
               Width           =   1890
            End
            Begin VB.Label lblSWISABOPEC 
               BackColor       =   &H00FFE0FF&
               Caption         =   "Code opération"
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
               Left            =   495
               TabIndex        =   145
               Top             =   2205
               Width           =   2010
            End
            Begin VB.Label lblSWISABOPEN 
               BackColor       =   &H00FFE0FF&
               Caption         =   "numéro opération"
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
               Left            =   495
               TabIndex        =   144
               Top             =   2655
               Width           =   2010
            End
            Begin VB.Label libSWISABKSRV 
               BackColor       =   &H00C0FFFF&
               Caption         =   "mise à jour"
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
               Left            =   315
               TabIndex        =   143
               Top             =   465
               Width           =   3840
            End
         End
         Begin VB.Frame fraSwift 
            BackColor       =   &H00C0E0FF&
            Height          =   7320
            Left            =   -74505
            TabIndex        =   131
            Top             =   900
            Visible         =   0   'False
            Width           =   6200
            Begin VB.CommandButton cmdSwift_Print 
               BackColor       =   &H00E0E0E0&
               Height          =   500
               Left            =   5595
               MaskColor       =   &H000080FF&
               Picture         =   "YGOSDOS0.frx":07FC
               Style           =   1  'Graphical
               TabIndex        =   218
               Top             =   660
               Width           =   500
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
               TabIndex        =   133
               Top             =   600
               Width           =   5370
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
               TabIndex        =   132
               Top             =   930
               Width           =   5400
            End
            Begin MSFlexGridLib.MSFlexGrid fgSwift 
               Height          =   6000
               Left            =   60
               TabIndex        =   134
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
               TabIndex        =   135
               Top             =   210
               Width           =   6050
            End
         End
         Begin VB.Frame fraMail_MT 
            BackColor       =   &H00F0FFF0&
            Caption         =   "Envoi courriel"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8025
            Left            =   495
            TabIndex        =   119
            Top             =   735
            Visible         =   0   'False
            Width           =   12600
            Begin VB.ListBox lstMail_MT_Message 
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
               Height          =   2760
               Left            =   2895
               TabIndex        =   121
               Top             =   3375
               Visible         =   0   'False
               Width           =   5370
            End
            Begin VB.CommandButton cmdMail_MT_NOK 
               BackColor       =   &H0080C0FF&
               Caption         =   "Ne pas envoyer"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   825
               Left            =   7050
               Style           =   1  'Graphical
               TabIndex        =   208
               Top             =   6645
               Width           =   1770
            End
            Begin VB.ListBox lstMail_MT_CC 
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
               Height          =   4620
               Left            =   5400
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   122
               Top             =   2850
               Visible         =   0   'False
               Width           =   3450
            End
            Begin VB.ListBox lstMail_MT_To 
               BackColor       =   &H00D0F0FF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5820
               Left            =   3300
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   130
               Top             =   1635
               Visible         =   0   'False
               Width           =   3450
            End
            Begin VB.TextBox txtMail_MT_To 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   720
               Left            =   2900
               MultiLine       =   -1  'True
               TabIndex        =   129
               Top             =   1020
               Width           =   9000
            End
            Begin VB.TextBox txtMail_MT_CC 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   720
               Left            =   2900
               MultiLine       =   -1  'True
               TabIndex        =   128
               Top             =   2130
               Width           =   9000
            End
            Begin VB.TextBox txtMail_MT_Message 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2880
               Left            =   2900
               MultiLine       =   -1  'True
               TabIndex        =   127
               Top             =   3375
               Width           =   9000
            End
            Begin VB.CommandButton cmdMail_MT_Ok 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Envoyer"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   825
               Left            =   9525
               Style           =   1  'Graphical
               TabIndex        =   126
               Top             =   6585
               Width           =   1770
            End
            Begin VB.CommandButton cmdMail_MT_Quit 
               BackColor       =   &H00C0C0FF&
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
               Height          =   825
               Left            =   1380
               Style           =   1  'Graphical
               TabIndex        =   125
               Top             =   6555
               Width           =   1770
            End
            Begin VB.CommandButton cmdMail_MT_To 
               BackColor       =   &H00D0F0FF&
               Caption         =   "Afficher la liste des destinataires"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   825
               Left            =   400
               Style           =   1  'Graphical
               TabIndex        =   124
               Top             =   990
               Width           =   1770
            End
            Begin VB.CommandButton cmdMail_MT_CC 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Afficher la liste pour copie"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   825
               Left            =   400
               Style           =   1  'Graphical
               TabIndex        =   123
               Top             =   2190
               Width           =   1770
            End
            Begin VB.CommandButton cmdMail_MT_Message 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Afficher la liste des messages"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   825
               Left            =   400
               Style           =   1  'Graphical
               TabIndex        =   120
               Top             =   3435
               Width           =   1770
            End
         End
         Begin VB.Frame fraDetail 
            BackColor       =   &H00D8DFD8&
            Height          =   8025
            Left            =   -74580
            TabIndex        =   96
            Top             =   675
            Width           =   12600
            Begin TabDlg.SSTab tabDetail 
               Height          =   7575
               Left            =   150
               TabIndex        =   97
               Top             =   195
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
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Dossier"
               TabPicture(0)   =   "YGOSDOS0.frx":08FE
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "libDetail_SWISABSWID"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "fgDetail"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "fraList"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).ControlCount=   3
               TabCaption(1)   =   "Evénements"
               TabPicture(1)   =   "YGOSDOS0.frx":091A
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "fraEVE_C"
               Tab(1).ControlCount=   1
               Begin VB.Frame fraEVE_C 
                  Height          =   7455
                  Left            =   -74910
                  TabIndex        =   110
                  Top             =   330
                  Width           =   12015
                  Begin VB.CommandButton cmdEVE_Invalidation 
                     BackColor       =   &H0000FFFF&
                     Caption         =   "Invalidation du dossier"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   735
                     Left            =   1860
                     Style           =   1  'Graphical
                     TabIndex        =   215
                     Top             =   6420
                     Visible         =   0   'False
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmdEVE_Clôture 
                     BackColor       =   &H00C0C0C0&
                     Caption         =   "Clôture du dossier"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   735
                     Left            =   7380
                     MaskColor       =   &H00000000&
                     Style           =   1  'Graphical
                     TabIndex        =   209
                     Top             =   6420
                     Visible         =   0   'False
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmdEVE_New 
                     BackColor       =   &H0080C0FF&
                     Caption         =   "Ajouter un événement"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   735
                     Left            =   10125
                     Style           =   1  'Graphical
                     TabIndex        =   115
                     Top             =   6420
                     Width           =   1650
                  End
                  Begin VB.CommandButton cmdEVE_Restauration 
                     BackColor       =   &H000000FF&
                     Caption         =   "Restauration du dossier"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   735
                     Left            =   8790
                     Style           =   1  'Graphical
                     TabIndex        =   114
                     Top             =   6435
                     Visible         =   0   'False
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmdEVE_Validation 
                     BackColor       =   &H00C0FFC0&
                     Caption         =   "Validation du dossier"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   735
                     Left            =   6060
                     Style           =   1  'Graphical
                     TabIndex        =   113
                     Top             =   6435
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmdEVE_Rejet 
                     BackColor       =   &H00FF80FF&
                     Caption         =   "Rejet du dossier"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   735
                     Left            =   4695
                     Style           =   1  'Graphical
                     TabIndex        =   112
                     Top             =   6400
                     Width           =   1200
                  End
                  Begin VB.CommandButton cmdEVE_Annulation 
                     BackColor       =   &H008080FF&
                     Caption         =   "Annulation du dossier"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   735
                     Left            =   3270
                     Style           =   1  'Graphical
                     TabIndex        =   111
                     Top             =   6400
                     Width           =   1200
                  End
                  Begin MSFlexGridLib.MSFlexGrid fgEVE 
                     Height          =   6195
                     Left            =   60
                     TabIndex        =   116
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   11895
                     _ExtentX        =   20981
                     _ExtentY        =   10927
                     _Version        =   393216
                     Rows            =   1
                     Cols            =   7
                     FixedCols       =   0
                     RowHeightMin    =   350
                     BackColor       =   15794175
                     ForeColor       =   8388608
                     BackColorFixed  =   8421376
                     ForeColorFixed  =   16777215
                     BackColorSel    =   14737632
                     BackColorBkg    =   16777210
                     WordWrap        =   -1  'True
                     AllowBigSelection=   0   'False
                     FocusRect       =   2
                     HighLight       =   0
                     GridLinesFixed  =   1
                     AllowUserResizing=   3
                     FormatString    =   $"YGOSDOS0.frx":0936
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
               Begin VB.Frame fraList 
                  BackColor       =   &H00D8DFD8&
                  Height          =   7155
                  Left            =   6465
                  TabIndex        =   98
                  Top             =   350
                  Visible         =   0   'False
                  Width           =   5940
                  Begin VB.Frame fraList_Options 
                     BackColor       =   &H00C0E0FF&
                     Height          =   2535
                     Left            =   80
                     TabIndex        =   100
                     Top             =   4605
                     Width           =   5775
                     Begin VB.CommandButton cmdList_Display 
                        BackColor       =   &H0080FFFF&
                        Caption         =   "Afficher le dossier"
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
                        Left            =   150
                        Style           =   1  'Graphical
                        TabIndex        =   227
                        Top             =   375
                        Width           =   1890
                     End
                     Begin VB.TextBox txtList_Add 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00FFFF80&
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
                        Left            =   2085
                        TabIndex        =   226
                        Top             =   885
                        Width           =   735
                     End
                     Begin VB.CommandButton cmdDetail_LAB_Quit 
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
                        Height          =   555
                        Left            =   3330
                        Style           =   1  'Graphical
                        TabIndex        =   108
                        Top             =   1890
                        Width           =   2280
                     End
                     Begin VB.CommandButton cmdList_Add 
                        BackColor       =   &H00FFFF80&
                        Caption         =   "Ajouter ce message  au dossier"
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
                        Left            =   120
                        Style           =   1  'Graphical
                        TabIndex        =   107
                        Top             =   750
                        Width           =   1890
                     End
                     Begin VB.CommandButton cmdList_New 
                        BackColor       =   &H00C0FFC0&
                        Caption         =   "Créer un nouveau dossier"
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
                        Left            =   3330
                        Style           =   1  'Graphical
                        TabIndex        =   106
                        Top             =   700
                        Width           =   2280
                     End
                     Begin VB.CommandButton cmdList_Ignore 
                        BackColor       =   &H0000C000&
                        Caption         =   "Vu"
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
                        Left            =   3330
                        Style           =   1  'Graphical
                        TabIndex        =   105
                        Top             =   1290
                        Width           =   2280
                     End
                     Begin VB.ComboBox cboList_SWISABKSRV 
                        Height          =   315
                        Left            =   3315
                        Sorted          =   -1  'True
                        Style           =   2  'Dropdown List
                        TabIndex        =   104
                        Top             =   225
                        Width           =   2280
                     End
                     Begin VB.CommandButton cmdList_SWISABKSRV 
                        BackColor       =   &H00C0C0C0&
                        Caption         =   "affecter ce message  au service =>"
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
                        Left            =   105
                        Style           =   1  'Graphical
                        TabIndex        =   103
                        Top             =   120
                        Width           =   2730
                     End
                     Begin VB.CommandButton cmdList_SAB_Annulation 
                        BackColor       =   &H004040FF&
                        Caption         =   "Annuler le message dans SAB"
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
                        Left            =   105
                        Style           =   1  'Graphical
                        TabIndex        =   102
                        Top             =   1935
                        Width           =   2700
                     End
                     Begin VB.CommandButton cmdList_SAB_Modification 
                        BackColor       =   &H0080C0FF&
                        Caption         =   "Modifier la référence du message"
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
                        Left            =   135
                        Style           =   1  'Graphical
                        TabIndex        =   101
                        Top             =   1350
                        Width           =   2700
                     End
                  End
                  Begin VB.CommandButton cmdList_Quit 
                     BackColor       =   &H00C0C0FF&
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
                     Height          =   500
                     Left            =   3360
                     Style           =   1  'Graphical
                     TabIndex        =   99
                     Top             =   6585
                     Width           =   2280
                  End
                  Begin MSFlexGridLib.MSFlexGrid fgList 
                     Height          =   4350
                     Left            =   45
                     TabIndex        =   109
                     Top             =   120
                     Width           =   5775
                     _ExtentX        =   10186
                     _ExtentY        =   7673
                     _Version        =   393216
                     Cols            =   7
                     FixedCols       =   0
                     RowHeightMin    =   350
                     BackColor       =   16777215
                     ForeColor       =   8192
                     BackColorFixed  =   33023
                     ForeColorFixed  =   16777215
                     BackColorBkg    =   16777215
                     Redraw          =   -1  'True
                     AllowUserResizing=   3
                     FormatString    =   "<Type  |<BIC émis/reçu            |<référence                           |>Montant               |<Dev    ||"
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
               Begin MSFlexGridLib.MSFlexGrid fgDetail 
                  Height          =   6795
                  Left            =   120
                  TabIndex        =   117
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
                  TabIndex        =   118
                  Top             =   330
                  Width           =   5985
               End
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
            Left            =   -66255
            TabIndex        =   95
            Top             =   6795
            Visible         =   0   'False
            Width           =   4212
         End
         Begin VB.Frame fraSelect_Options_1a 
            Appearance      =   0  'Flat
            BackColor       =   &H00F0FFFF&
            BorderStyle     =   0  'None
            Caption         =   "-"
            ForeColor       =   &H80000008&
            Height          =   1260
            Left            =   -66075
            TabIndex        =   87
            Top             =   2025
            Width           =   2130
            Begin VB.OptionButton optSelect_rTextField_OR 
               BackColor       =   &H00F0FFFF&
               Caption         =   "OU"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   600
               TabIndex        =   93
               Top             =   660
               Width           =   500
            End
            Begin VB.OptionButton optSelect_rTextField_AND 
               BackColor       =   &H00F0FFFF&
               Caption         =   "ET"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   105
               TabIndex        =   92
               Top             =   645
               Value           =   -1  'True
               Width           =   450
            End
            Begin VB.CheckBox chkSelect_rTextField 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0FFFF&
               Caption         =   "Recherche Texte (MAJ)"
               Height          =   195
               Left            =   120
               TabIndex        =   91
               Top             =   30
               Value           =   1  'Checked
               Width           =   1890
            End
            Begin VB.TextBox txtSelect_rTextField_Code 
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
               Left            =   1095
               MaxLength       =   2
               TabIndex        =   90
               Top             =   915
               Width           =   450
            End
            Begin VB.TextBox txtSelect_rTextField2 
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
               Left            =   1080
               TabIndex        =   89
               Top             =   600
               Width           =   900
            End
            Begin VB.TextBox txtSelect_rTextField1 
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
               Left            =   135
               TabIndex        =   88
               Top             =   300
               Width           =   1830
            End
            Begin VB.Label Label1 
               BackColor       =   &H00F0FFFF&
               Caption         =   "N° champ"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   165
               TabIndex        =   94
               Top             =   945
               Width           =   825
            End
         End
         Begin VB.Frame fraSelect_Options_1b 
            Appearance      =   0  'Flat
            BackColor       =   &H00F0FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1260
            Left            =   -66120
            TabIndex        =   82
            Top             =   495
            Width           =   2130
            Begin VB.CheckBox chkSelect_GOSDOSKSRV 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0080C0FF&
               Caption         =   "+ service NONE"
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
               Left            =   270
               TabIndex        =   85
               Top             =   75
               Value           =   1  'Checked
               Width           =   1500
            End
            Begin VB.ComboBox cboSelect_GOSDOSKSRV 
               Height          =   315
               Left            =   255
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Top             =   390
               Width           =   1620
            End
            Begin VB.ComboBox cboSelect_SWISABWSTA 
               Height          =   315
               Left            =   930
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Top             =   810
               Width           =   975
            End
            Begin VB.Label lblSelect_SWISABWSTA 
               BackColor       =   &H00F0FFFF&
               Caption         =   "état SAA"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   180
               TabIndex        =   86
               Top             =   840
               Width           =   705
            End
         End
         Begin VB.Frame fraSelect_Options_3 
            BackColor       =   &H00F0FFFF&
            BorderStyle     =   0  'None
            Height          =   1260
            Left            =   -74685
            TabIndex        =   71
            Top             =   480
            Visible         =   0   'False
            Width           =   7845
            Begin VB.TextBox txtSelect_3_GOSDOSIDD 
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
               Height          =   315
               Left            =   120
               TabIndex        =   214
               Top             =   630
               Width           =   930
            End
            Begin VB.ComboBox cboSelect_3_GOSDOSSTAG 
               Height          =   315
               Left            =   6450
               Sorted          =   -1  'True
               TabIndex        =   210
               Top             =   675
               Width           =   1215
            End
            Begin VB.ComboBox cboSelect_3_GOSDOSGSRV 
               Height          =   315
               Left            =   2000
               Sorted          =   -1  'True
               TabIndex        =   76
               Top             =   150
               Width           =   1400
            End
            Begin VB.ComboBox cboSelect_3_GOSDOSRCOM 
               Height          =   315
               Left            =   2000
               Sorted          =   -1  'True
               TabIndex        =   75
               Top             =   660
               Width           =   1400
            End
            Begin VB.ComboBox cboSelect_3_GOSDOSWBIC 
               Height          =   315
               Left            =   4300
               Sorted          =   -1  'True
               TabIndex        =   74
               Top             =   165
               Width           =   1300
            End
            Begin VB.ComboBox cboSelect_3_GOSDOSCLI 
               Height          =   315
               Left            =   4300
               Sorted          =   -1  'True
               TabIndex        =   73
               Top             =   675
               Width           =   1300
            End
            Begin VB.ComboBox cboSelect_3_GOSDOSSTAD 
               Height          =   315
               Left            =   6435
               Sorted          =   -1  'True
               TabIndex        =   72
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label lblSelect_3_GOSDOSIDD 
               Alignment       =   2  'Center
               BackColor       =   &H0080C0FF&
               Caption         =   "N° GOS"
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
               TabIndex        =   213
               Top             =   150
               Width           =   930
            End
            Begin VB.Label lblSelect_3_GOSDOSGSRV 
               BackColor       =   &H00F0FFFF&
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
               Left            =   1260
               TabIndex        =   81
               Top             =   210
               Width           =   585
            End
            Begin VB.Label lblSelect_3_GOSDOSCLI 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Client"
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
               Left            =   3630
               TabIndex        =   80
               Top             =   690
               Width           =   525
            End
            Begin VB.Label lblSelect_3_GOSDOSRCOM 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Resp com"
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
               Left            =   1305
               TabIndex        =   79
               Top             =   690
               Width           =   570
            End
            Begin VB.Label lblSelect_3_GOSDOSWBIC 
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
               Height          =   255
               Left            =   3660
               TabIndex        =   78
               Top             =   270
               Width           =   615
            End
            Begin VB.Label lblSelect_3_GOSDOSSTAD 
               BackColor       =   &H00F0FFFF&
               Caption         =   "code état"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   510
               Left            =   5820
               TabIndex        =   77
               Top             =   90
               Width           =   540
            End
         End
         Begin VB.Frame fraSelect_Options_9 
            BackColor       =   &H00F0FFFF&
            BorderStyle     =   0  'None
            Height          =   1260
            Left            =   -74700
            TabIndex        =   65
            Top             =   1995
            Visible         =   0   'False
            Width           =   7845
            Begin VB.CheckBox chkSelect_9_SWISABKSRV 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0080C0FF&
               Caption         =   "inclure le service NONE"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4410
               TabIndex        =   68
               Top             =   60
               Value           =   1  'Checked
               Width           =   2190
            End
            Begin VB.ComboBox cboSelect_9_SWISABKSTA 
               Height          =   315
               Left            =   1095
               TabIndex        =   67
               Top             =   645
               Width           =   2355
            End
            Begin VB.ComboBox cboSelect_9_SWISABKSRV 
               Height          =   315
               Left            =   4635
               Sorted          =   -1  'True
               TabIndex        =   66
               Top             =   645
               Width           =   2130
            End
            Begin VB.Label lblSelect_9_SWISABKSTA 
               BackColor       =   &H00F0FFFF&
               Caption         =   "code état"
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
               TabIndex        =   70
               Top             =   675
               Width           =   720
            End
            Begin VB.Label lblSelect_9_SWISABKSRV 
               BackColor       =   &H00F0FFFF&
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
               Left            =   3795
               TabIndex        =   69
               Top             =   675
               Width           =   585
            End
         End
         Begin VB.TextBox txtFg 
            Height          =   1260
            Left            =   -74700
            MultiLine       =   -1  'True
            TabIndex        =   64
            Top             =   7665
            Visible         =   0   'False
            Width           =   6732
         End
         Begin MSFlexGridLib.MSFlexGrid fgFree 
            Height          =   3465
            Left            =   -67000
            TabIndex        =   155
            Top             =   5000
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
            FormatString    =   $"YGOSDOS0.frx":0A14
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
         Begin MSFlexGridLib.MSFlexGrid fgModèle 
            Height          =   4680
            Left            =   -69480
            TabIndex        =   198
            Top             =   4350
            Visible         =   0   'False
            Width           =   7710
            _ExtentX        =   13600
            _ExtentY        =   8255
            _Version        =   393216
            Cols            =   3
            FixedCols       =   0
            RowHeightMin    =   500
            BackColor       =   16777215
            ForeColor       =   8192
            BackColorFixed  =   12632064
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            WordWrap        =   -1  'True
            Redraw          =   -1  'True
            AllowUserResizing=   3
            FormatString    =   $"YGOSDOS0.frx":0AD5
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
      Begin TabDlg.SSTab tabParam 
         Height          =   9465
         Left            =   -74880
         TabIndex        =   23
         Top             =   360
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   16695
         _Version        =   393216
         Tabs            =   4
         Tab             =   1
         TabsPerRow      =   4
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
         TabCaption(0)   =   "paramétrage GOS"
         TabPicture(0)   =   "YGOSDOS0.frx":0B77
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lstParam"
         Tab(0).Control(1)=   "fraParam_GOSDOSLABK"
         Tab(0).Control(2)=   "lstParam_GOSDOSLABK"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Mail / service"
         TabPicture(1)   =   "YGOSDOS0.frx":0B93
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "libPARAM_MAil"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "SAA : paramétrage"
         TabPicture(2)   =   "YGOSDOS0.frx":0BAF
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraParam_SAA"
         Tab(2).Control(1)=   "lstParam_SAA_K1"
         Tab(2).Control(2)=   "lstParam_SAA_Id"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "-"
         TabPicture(3)   =   "YGOSDOS0.frx":0BCB
         Tab(3).ControlEnabled=   0   'False
         Tab(3).ControlCount=   0
         Begin VB.Frame fraParam_SAA 
            BackColor       =   &H00E0FFFF&
            Height          =   5160
            Left            =   -68550
            TabIndex        =   46
            Top             =   3855
            Visible         =   0   'False
            Width           =   6585
            Begin VB.Frame fraParam_SAA_Jrnl_Event 
               BackColor       =   &H00C0E0FF&
               Height          =   2700
               Left            =   0
               TabIndex        =   55
               Top             =   1305
               Width           =   6555
               Begin VB.OptionButton optParam_SAA_TXT_RMA_X 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "autre(Query...)"
                  Height          =   195
                  Left            =   4890
                  TabIndex        =   233
                  Top             =   2205
                  Width           =   1350
               End
               Begin VB.OptionButton optParam_SAA_TXT_RMA_R 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Révocation"
                  Height          =   195
                  Left            =   3765
                  TabIndex        =   232
                  Top             =   2235
                  Width           =   1350
               End
               Begin VB.OptionButton optParam_SAA_TXT_RMA_A 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Autorisation"
                  Height          =   195
                  Left            =   2565
                  TabIndex        =   231
                  Top             =   2205
                  Width           =   1350
               End
               Begin VB.ComboBox cboParam_SAA_TopK 
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
                  Left            =   3645
                  Sorted          =   -1  'True
                  TabIndex        =   61
                  Top             =   960
                  Width           =   2565
               End
               Begin VB.ComboBox cboParam_SAA_Alerte 
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
                  Left            =   3645
                  Sorted          =   -1  'True
                  TabIndex        =   58
                  Top             =   1665
                  Width           =   2655
               End
               Begin VB.TextBox txtParam_SAA_TXT 
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
                  Left            =   1095
                  TabIndex        =   56
                  Top             =   300
                  Width           =   5130
               End
               Begin VB.Label lblParam_SAA_TXT_RMA 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "RMA : nature de l'évènement"
                  Height          =   330
                  Left            =   375
                  TabIndex        =   230
                  Top             =   2190
                  Width           =   2655
               End
               Begin VB.Label lblParam_SAA_TopK_2 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "* => tous les événements     # => cas particuliers (pgm)"
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
                  Left            =   450
                  TabIndex        =   62
                  Top             =   1065
                  Width           =   2370
               End
               Begin VB.Label lblParam_SAA_TopK 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Historiser pour recherches rapides (J=)"
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
                  Left            =   225
                  TabIndex        =   60
                  Top             =   855
                  Width           =   3240
               End
               Begin VB.Label lblParam_SAA_Alerte 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Service à alerter par courriel"
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
                  TabIndex        =   59
                  Top             =   1650
                  Width           =   2760
               End
               Begin VB.Label lblParam_SAA_TXT 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Libellé"
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
                  Left            =   315
                  TabIndex        =   57
                  Top             =   330
                  Width           =   675
               End
            End
            Begin VB.TextBox txtParam_SAA_K2 
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
               Left            =   2300
               TabIndex        =   52
               Top             =   255
               Width           =   2040
            End
            Begin VB.CommandButton cmdParam_SAA_Delete 
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
               Left            =   5070
               Style           =   1  'Graphical
               TabIndex        =   51
               Top             =   4200
               Width           =   900
            End
            Begin VB.CommandButton cmdParam_SAA_Add 
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
               Left            =   2175
               Style           =   1  'Graphical
               TabIndex        =   50
               Top             =   4200
               Width           =   900
            End
            Begin VB.CommandButton cmdParam_SAA_Update 
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
               TabIndex        =   49
               Top             =   4200
               Width           =   900
            End
            Begin VB.CommandButton cmdParam_SAA_Quit 
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
               Left            =   315
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   4200
               Width           =   990
            End
            Begin VB.TextBox txtParam_SAA_MTD 
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
               Left            =   2300
               MaxLength       =   12
               TabIndex        =   47
               Top             =   870
               Width           =   2055
            End
            Begin VB.Label lblParam_SAA_K2 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Code "
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
               TabIndex        =   54
               Top             =   255
               Width           =   1290
            End
            Begin VB.Label lblParam_SAA_MTD 
               BackColor       =   &H00E0FFFF&
               Caption         =   "montant autorisé €"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   270
               TabIndex        =   53
               Top             =   810
               Width           =   1995
            End
         End
         Begin VB.ListBox lstParam_SAA_K1 
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
            Height          =   6060
            Left            =   -74835
            TabIndex        =   45
            Top             =   3090
            Visible         =   0   'False
            Width           =   8160
         End
         Begin VB.ListBox lstParam_SAA_Id 
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
            Height          =   1980
            Left            =   -74745
            TabIndex        =   44
            Top             =   570
            Visible         =   0   'False
            Width           =   4785
         End
         Begin VB.ListBox lstParam 
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
            Height          =   1980
            Left            =   -74655
            TabIndex        =   42
            Top             =   585
            Width           =   4785
         End
         Begin VB.Frame fraParam_GOSDOSLABK 
            BackColor       =   &H00E0FFFF&
            Height          =   8835
            Left            =   -68865
            TabIndex        =   25
            Top             =   405
            Visible         =   0   'False
            Width           =   6585
            Begin VB.ComboBox cboParam_StaC 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2550
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   220
               Top             =   2025
               Width           =   3705
            End
            Begin VB.ComboBox cboParam_GOSDOSLABK_GSrv 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2565
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   1410
               Width           =   2400
            End
            Begin VB.TextBox txtParam_GOSDOSLABK_J 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2550
               MaxLength       =   3
               TabIndex        =   39
               Top             =   855
               Width           =   705
            End
            Begin VB.CommandButton cmdParam_GOSDOSLABK_Quit 
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
               TabIndex        =   31
               Top             =   8200
               Width           =   990
            End
            Begin VB.CommandButton cmdParam_GOSDOSLABK_Update 
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
               TabIndex        =   30
               Top             =   8200
               Width           =   900
            End
            Begin VB.CommandButton cmdParam_GOSDOSLABK_Add 
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
               TabIndex        =   29
               Top             =   8200
               Width           =   900
            End
            Begin VB.CommandButton cmdParam_GOSDOSLABK_Delete 
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
               TabIndex        =   28
               Top             =   8200
               Width           =   900
            End
            Begin VB.TextBox txtParam_GOSDOSLABK 
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
               TabIndex        =   26
               Top             =   255
               Width           =   2040
            End
            Begin VB.TextBox txtParam_GOSDOSLABK_Lib 
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
               Left            =   645
               MultiLine       =   -1  'True
               TabIndex        =   27
               Text            =   "YGOSDOS0.frx":0BE7
               Top             =   3000
               Width           =   5355
            End
            Begin VB.Label lblParam_StaC 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Code statistique"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   405
               TabIndex        =   221
               Top             =   2115
               Width           =   1725
            End
            Begin VB.Label lblParam_GOSDOSLABK_GSrv 
               BackColor       =   &H00E0FFFF&
               Caption         =   "service gestionnaire"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   300
               TabIndex        =   40
               Top             =   1530
               Width           =   2070
            End
            Begin VB.Label lblParam_GOSDOSLABK_J 
               BackColor       =   &H00E0FFFF&
               Caption         =   "nb jours calendaires => échéance"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   270
               TabIndex        =   38
               Top             =   810
               Width           =   1995
            End
            Begin VB.Label lblParam_GOSDOSLABK 
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
               TabIndex        =   32
               Top             =   255
               Width           =   1290
            End
         End
         Begin VB.ListBox lstParam_GOSDOSLABK 
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
            Height          =   6060
            Left            =   -74700
            TabIndex        =   24
            Top             =   3030
            Width           =   4890
         End
         Begin VB.Label libPARAM_MAil 
            Caption         =   "màj BIA_AUDIT > BIA_SSI > Param"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   3870
            TabIndex        =   229
            Top             =   1830
            Width           =   4815
         End
      End
      Begin VB.Frame fraTab0 
         Height          =   9420
         Left            =   -74880
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
            Begin VB.CheckBox chkSelect_GOSDOSIAMJ 
               BackColor       =   &H00F0FFFF&
               Caption         =   "date création"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   8010
               TabIndex        =   22
               Top             =   135
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker txtSelect_GOSDOSIAMJ_Max 
               Height          =   300
               Left            =   8025
               TabIndex        =   9
               Top             =   840
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
               Format          =   92930051
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_GOSDOSIAMJ_Min 
               Height          =   300
               Left            =   8025
               TabIndex        =   12
               Top             =   465
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
               Format          =   92930051
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Frame fraSelect_Options_1 
               BackColor       =   &H00F0FFFF&
               BorderStyle     =   0  'None
               Height          =   1260
               Left            =   15
               TabIndex        =   10
               Top             =   0
               Width           =   7845
               Begin VB.TextBox txtSelect_GOSDOSWMTD 
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
                  Left            =   2200
                  TabIndex        =   37
                  Top             =   870
                  Width           =   1035
               End
               Begin VB.ComboBox cboSelect_SWISABOPEC 
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
                  Left            =   315
                  Sorted          =   -1  'True
                  TabIndex        =   35
                  Text            =   "WMTK"
                  Top             =   450
                  Width           =   780
               End
               Begin VB.TextBox txtSelect_SWISABOPEN 
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
                  Left            =   135
                  TabIndex        =   33
                  Top             =   885
                  Width           =   1200
               End
               Begin VB.ComboBox cboSelect_GOSDOSWES 
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
                  Left            =   4080
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   20
                  Top             =   45
                  Width           =   930
               End
               Begin VB.TextBox txtSelect_GOSDOSWTRN 
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
                  Left            =   4080
                  TabIndex        =   18
                  Top             =   915
                  Width           =   1620
               End
               Begin VB.ComboBox cboSelect_GOSDOSWMTK 
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
                  Left            =   2200
                  Sorted          =   -1  'True
                  TabIndex        =   16
                  Text            =   "WMTK"
                  Top             =   30
                  Width           =   1050
               End
               Begin VB.ComboBox cboSelect_GOSDOSWBIC 
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
                  Left            =   4080
                  Sorted          =   -1  'True
                  TabIndex        =   15
                  Text            =   "BIC sender"
                  Top             =   435
                  Width           =   1665
               End
               Begin VB.ComboBox cboSelect_GOSDOSWDEV 
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
                  Left            =   2200
                  Sorted          =   -1  'True
                  TabIndex        =   13
                  Text            =   "dev"
                  Top             =   450
                  Width           =   900
               End
               Begin VB.Label lblSelect_GOSDOSWMTD 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "montant"
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
                  Left            =   1485
                  TabIndex        =   36
                  Top             =   885
                  Width           =   780
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H000080FF&
                  BorderWidth     =   3
                  X1              =   1400
                  X2              =   1400
                  Y1              =   0
                  Y2              =   1200
               End
               Begin VB.Label lblSelect_SWISABOPEN 
                  BackColor       =   &H0080C0FF&
                  Caption         =   "  code + n° opé"
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
                  Left            =   45
                  TabIndex        =   34
                  Top             =   60
                  Width           =   1305
               End
               Begin VB.Label lblSelect_GOSDOSWER 
                  BackColor       =   &H00F0FFFF&
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
                  Left            =   3435
                  TabIndex        =   21
                  Top             =   105
                  Width           =   435
               End
               Begin VB.Label lblSelect_GOSDOSWTRN 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Champ 20-21"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   3400
                  TabIndex        =   19
                  Top             =   810
                  Width           =   750
               End
               Begin VB.Label lblSelect_GOSDOSWMTK 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "MT xx%"
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
                  Left            =   1485
                  TabIndex        =   17
                  Top             =   80
                  Width           =   570
               End
               Begin VB.Label lblSelect_GOSDOSWDEV 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "devise"
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
                  Left            =   1470
                  TabIndex        =   14
                  Top             =   480
                  Width           =   615
               End
               Begin VB.Label lblSelect_GOSDOSWSND 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "BIC Emis/reçu"
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
                  Left            =   3400
                  TabIndex        =   11
                  Top             =   465
                  Width           =   645
               End
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7890
            Left            =   120
            TabIndex        =   5
            Top             =   1440
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
            FormatString    =   $"YGOSDOS0.frx":0C1C
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
      Picture         =   "YGOSDOS0.frx":0D22
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
   Begin VB.Menu mnuPrint2 
      Caption         =   "mnuPrint2"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint2_Mail 
         Caption         =   "envoi mail"
      End
      Begin VB.Menu mnuPrint2_Excel 
         Caption         =   "excel"
      End
   End
   Begin VB.Menu mnuYSWIECH0 
      Caption         =   "mnuYSWIECH0"
      Visible         =   0   'False
      Begin VB.Menu mnuYSWIECH0_Ann 
         Caption         =   "Annuler cet enregistrement (A)"
      End
      Begin VB.Menu mnuYSWIECH0_Res 
         Caption         =   "Restaurer cet enregistrement (#)"
      End
   End
End
Attribute VB_Name = "frmYGOSDOS0"
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
Dim arrHab(19) As Boolean, blnHab_YGOSEVE0_New  As Boolean
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

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long

Dim xYGOSDOS0 As typeYGOSDOS0, newYGOSDOS0 As typeYGOSDOS0, oldYGOSDOS0 As typeYGOSDOS0
Dim arrYGOSDOS0() As typeYGOSDOS0, arrYGOSDOS0_Nb As Long, arrYGOSDOS0_Max As Long, arrYGOSDOS0_Index As Long

Dim xYGOSEVE0 As typeYGOSEVE0, newYGOSEVE0 As typeYGOSEVE0, oldYGOSEVE0 As typeYGOSEVE0
Dim arrYGOSEVE0() As typeYGOSEVE0, arrYGOSEVE0_Nb As Long, arrYGOSEVE0_Max As Long, arrYGOSEVE0_Index As Long

Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean
Dim fgDetail_50 As String, fgDetail_59 As String, fgDetail_57 As String
Dim fgDetail_32B As String, fgDetail_33B As String, fgDetail_36 As String
Dim fgDetail_30V As String, fgDetail_37G As String, fgDetail_34E As String
Dim fgDetail_30T As String, fgDetail_30P As String, fgDetail_82A As String, fgDetail_87A As String
Dim fgDetail_22C As String
Dim fgDetail_70 As String, fgDetail_72 As String

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim mXls1_Row As Long, mXls1_Cols As Long, mXls1_File As Integer, mXls1_Cols_WMTK As Long, mWMTK_Col As String
Dim mXls1_Col_1 As Long, mXls1_Col_2 As Long
Dim mXls1_Col As Long, mXls2_Row As Long, mXls2_Col As Long

Dim cnSIDE_DB As New ADODB.Connection, rsSIDE_DB As New ADODB.Recordset, rsSIDE_Loop As New ADODB.Recordset
Dim rsSIDE_X As New ADODB.Recordset

Dim mSWISABSWID_MT700 As Long

Dim xrMesg As typerMesg, xrIntv As typerIntv
Dim wAmj8_tiret As String, xAmj8_from_crea_date_time As String, xAmj8_to_crea_date_time As String
Dim xAmj8_1Live As String

Dim blnGOSEVE_Mail As Boolean, blnGOSDOSSTAD_X As Boolean, blnGOSDOSSTAD_C As Boolean

Dim mfraDetail_Width As Integer
Dim mGOSDOSIDD_Last As Long
Dim arrService_Code(100) As String, arrService_Lib(100) As String, arrService_Mail(100, 2) As String
Dim arrService_Code_SAA(100) As String, mParam_Mail_K As Integer

Dim arrRCOM_Code(100) As String, arrRCOM_Lib(100) As String, arrRCOM_Mail(100) As String

Dim fgEVE_FormatString As String, fgEVE_K As Integer
Dim fgEVE_RowDisplay As Integer, fgEVE_RowClick As Integer, fgEVE_ColClick As Integer
Dim fgEVE_ColorClick As Long, fgEVE_ColorDisplay As Long
Dim fgEVE_Sort1 As Integer, fgEVE_Sort2 As Integer
Dim fgEVE_SortAD As Integer, fgEVE_Sort1_Old As Integer
Dim fgEVE_arrIndex As Integer
Dim blnfgEVE_DisplayLine As Boolean

Dim mYGOSDOS0_Fct As String, mYGOSEVE0_Fct  As String, blnYGOSDOS0_Display As Boolean, blnYGOSDOS0_Update As Boolean
Dim paramGOSDOS_Path As String, paramGOSDOS_Path_DROPI As String
Dim oldFileName As String, newFileName As String, newDirPath As String, newFileExtension As String


Dim Old_YBIATAB0 As typeYBIATAB0, New_YBIATAB0 As typeYBIATAB0

Dim mYSWILNK0_Fct As String
Dim xYSWILNK0 As typeYSWILNK0, newYSWILNK0 As typeYSWILNK0, oldYSWILNK0 As typeYSWILNK0
Dim arrYSWILNK0() As typeYSWILNK0, arrYSWILNK0_Nb As Long, arrYSWILNK0_Max As Long, arrYSWILNK0_Index As Long

Dim mYSWISAB0_Fct As String
Dim xYSWISAB0 As typeYSWISAB0, newYSWISAB0 As typeYSWISAB0, oldYSWISAB0 As typeYSWISAB0
Dim arrYSWISAB0() As typeYSWISAB0, arrYSWISAB0_Nb As Long, arrYSWISAB0_Max As Long, arrYSWISAB0_Index As Long
Dim matchYSWISAB0 As typeYSWISAB0
Dim m999_YSWISAB0 As typeYSWISAB0

Dim xYSWISAB1 As typeYSWISAB1, newYSWISAB1 As typeYSWISAB1, oldYSWISAB1 As typeYSWISAB1

Dim mList_Row As Long, mList_YGOSDOS0 As typeYGOSDOS0, m999_YGOSDOS0 As typeYGOSDOS0
Dim m999_YGOSEVE0 As typeYGOSEVE0, duplic_YGOSEVE0 As typeYGOSEVE0

Dim fglist_FormatString As String, fglist_K As Integer
Dim fglist_RowDisplay As Integer, fglist_RowClick As Integer, fglist_ColClick As Integer
Dim fglist_ColorClick As Long, fglist_ColorDisplay As Long
Dim fglist_Sort1 As Integer, fglist_Sort2 As Integer
Dim fglist_SortAD As Integer, fglist_Sort1_Old As Integer
Dim fglist_arrIndex As Integer
Dim blnfglist_DisplayLine As Boolean

Dim blnSwift_Display As Boolean

Dim oldZSWIENA0 As typeZSWIENA0, newZSWIENA0 As typeZSWIENA0
Dim mZSWIENA0_Fct As String
 Dim newYSWIMON0 As typeYSWIMON0, oldYSWIMON0 As typeYSWIMON0
Dim rsSabX As New ADODB.Recordset
Dim importSWISABSWID As Long, autoSWISABSWID As Long, autoSWISABZSWI As Long

Dim newYBIADTAQ As typeYBIADTAQ
Dim xrText As typerText

Dim HeightOfLine As Long, LinesOfText As Long
Dim Mesg_aid As Long, mesg_s_umidl As Long, mesg_s_umidh As Long
Dim mSWISABSWID As Long

Dim mMOUVEMSER As String, mMOUVEMSSE As String, mMOUVEMOPE As String, mMOUVEMNUM As Long
Dim mSWISABSWID_Xd As Long
Dim cmdSelect_SQL_1_rText As String, blnSelect_SQL_1_rText As Boolean

Dim rtextField_Value As Variant
Dim fgSwift_FormatString As String
Dim xParam As typeYGOSEVE0, newParam As typeYGOSEVE0, oldParam As typeYGOSEVE0
Dim blnYGOSDOS0_New As Boolean, cmdSelect_SQL_Kbis As String

Dim mfilDoc_Path As String, blnfilDoc_Path As Boolean

Dim arrSAA_Usr_Id() As String, arrSAA_Usr_MTD() As Currency, arrSAA_Usr_Nb As Integer
Dim curSAA_103_EUR As Currency
Dim curSAA_202_EUR As Currency
Dim curSAA_202_BOTC_EUR As Currency
Dim newSAA_YBIATAB0 As typeYBIATAB0

Dim cours_X As Double

Dim arrJrnl_Nb As Long, arrJrnl_Comp_Name() As String, arrJrnl_Event_Num() As Long, arrJrnl_Alerte() As String, arrJrnl_Top() As String
Dim xrJrnl As typerJrnl
Dim newYSAAJRN0 As typeYSAAJRN0, oldYSAAJRN0 As typeYSAAJRN0
Dim arrJrnl_Event_Id() As String, arrJrnl_Event_Lib() As String, arrJrnl_Event_Nb As Integer, arrJrnl_Event_K As Integer

Dim wMT_BIC_E As String, wMT_BIC_S As String
Dim wMT_50A As String, wMT_50P As String
Dim wMT_52A As String, wMT_52P As String
Dim wMT_54A As String
Dim wMT_57A As String, wMT_57P As String
Dim wMT_58A As String, wMT_58P As String
Dim wMT_59A As String, wMT_59P As String
Dim wMT_82A As String
Dim wMT_87A As String
Dim wMT_20 As String, wMT_21 As String
Dim wMT_22A As String, wMT_22B As String
Dim wMT_22C As String
Dim wMT_30V As String, wMT_30V_JS1 As String, wMT_30X As String
Dim wMT_53A_B1 As String, wMT_53A_B2 As String
Dim wMT_57A_B1 As String, wMT_57A_B2 As String
Dim wMT_32_DEV As String, wMT_33_DEV As String, wMT_34_DEV As String
Dim wMT_32_MTD As Currency, wMT_33_MTD As Currency, wMT_34_MTD As Currency
Dim wMT_34_N As String, wMT_17R As String



Dim arrStaP() As typeX_Stat, arrStaP_Nb As Long
Dim arrStaC() As typeX_Stat, arrStaC_Nb As Long
Dim arrDOS() As typeX_Stat, arrDOS_Nb As Long

Dim htmlFontColor_rText As String
Dim blnYSWILNK0_Display As Boolean, mYSWILNK0_Display As String

Dim mAMJ_SQL As String
Dim mSelect_SQL_Listindex_3 As Integer
'_________________________________________________________________________________________________

Dim newYSWIRAM0 As typeYSWIRAM0, oldYSWIRAM0 As typeYSWIRAM0, xYSWIRAM0 As typeYSWIRAM0
Dim mMTK As String, mMTK_Seq As String
Dim mYSWIRAM0_Fct As String, mYSWIRAM0_SQL_Set As String, mYSWIRAM0_Match_XOPE As String
Dim blnYSWIRAM0_Match_CONF As Boolean, blnYSWIRAM0_Match_Retry As Boolean
Dim xField_K As String, xField_K2 As String, xField_V As String

Dim wField_57A As String, blnReprise As Boolean

Dim xYSWIECH0 As typeYSWIECH0, newYSWIECH0 As typeYSWIECH0, oldYSWIECH0  As typeYSWIECH0, mYSWIECH0_Offset  As typeYSWIECH0
Dim xYSWIECH1 As typeYSWIECH1, newYSWIECH1 As typeYSWIECH1, oldYSWIECH1  As typeYSWIECH1
Dim xYSWI950 As typeYSWI950, newYSWI950 As typeYSWI950, oldYSWI950  As typeYSWI950
Dim mYSWIECH0_Fct As String, mYSWIECH1_Fct As String, mYSWI950_Fct As String
Dim mSWIECHSWIL As Integer
Dim blnK115 As Boolean, mK115_SWISABSWID As Long
Public Sub YSWIRAM0_Importation()
Dim xSql As String

On Error GoTo Error_Handler
currentAction = "YSWIRAM0_Importation "

'________________________________________________________________________
Call lstErr_Clear(lstErr, cmdContext, currentAction & " : 1"): DoEvents
'________________________________________________________________________

If mYSWIRAM0_SWISABSWID = 0 Then
    xSql = "select SWIRAMXID  from " & paramIBM_Library_SABSPE & ".YSWIRAM0 " _
         & "order by SWIRAMXID desc FETCH FIRST 1 ROWS ONLY"
    Set rsSab = cnsab.Execute(xSql)
    
    If Not rsSab.EOF Then mYSWIRAM0_SWISABSWID = rsSab(0)
    If mYSWIRAM0_SWISABSWID < 1167441 Then blnReprise = True: mYSWIRAM0_SWISABSWID = 1186430 '1081355  '1163240 ''1167441 '$JPL test
End If

blnYSWIRAM0_Match_Retry = False
xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABSWID > " & mYSWIRAM0_SWISABSWID & " and SWISABWMTK in ( '300' , '320','399') order by SWISABSWID"


Call YSWIRAM0_Importation_MT(xSql)


blnYSWIRAM0_Match_Retry = True
xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIRAM0," & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWIRAMSTA = '?' and SWISABSWID = SWIRAMXID order by SWISABSWID"

Call YSWIRAM0_Importation_MT(xSql)

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, currentAction & " : exit"): DoEvents
'________________________________________________________________________

End Sub

Public Sub YSWIECH0_Importation()
Dim xSql As String

On Error GoTo Error_Handler
currentAction = "YSWIECH0_Importation "

'________________________________________________________________________
Call lstErr_Clear(lstErr, cmdContext, currentAction & " : 1"): DoEvents
'________________________________________________________________________

If mYSWIECH0_SWISABSWID = 0 Then
    xSql = "select SWIECHSWID  from " & paramIBM_Library_SABSPE & ".YSWIECH0 " _
         & "order by SWIECHSWID desc FETCH FIRST 1 ROWS ONLY"
    Set rsSab = cnsab.Execute(xSql)
    
    If Not rsSab.EOF Then mYSWIECH0_SWISABSWID = rsSab(0)
    If mYSWIECH0_SWISABSWID < 1298720 Then blnReprise = True: mYSWIECH0_SWISABSWID = 1298720 '1286254
    
    xSql = "select SWIEC1SWID  from " & paramIBM_Library_SABSPE & ".YSWIECH1 " _
         & "order by SWIEC1SWID desc FETCH FIRST 1 ROWS ONLY"
    Set rsSab = cnsab.Execute(xSql)
    
    If Not rsSab.EOF Then
        If mYSWIECH0_SWISABSWID < rsSab(0) Then mYSWIECH0_SWISABSWID = rsSab(0)
    End If
End If

oldYSWIECH0_SWISABSWID = mYSWIECH0_SWISABSWID

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABSWID > " & mYSWIECH0_SWISABSWID & " and SWISABWMTK in ( '103', '300', '320')" _
     & " and SWISABWES = 'S' order by SWISABSWID"


Call YSWIECH0_Importation_MT(xSql)

If xYSWIECH0.SWIECHSWID > mYSWIECH0_SWISABSWID Then mYSWIECH0_SWISABSWID = xYSWIECH0.SWIECHSWID


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & currentAction
Exit_sub:

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, currentAction & " : exit"): DoEvents
'________________________________________________________________________

End Sub
Public Sub YSWIECH0_Importation_2()
Dim xSql As String

On Error GoTo Error_Handler
currentAction = "YSWIECH0_Importation_MT202 "

'________________________________________________________________________
Call lstErr_Clear(lstErr, cmdContext, currentAction & " : 1"): DoEvents
'________________________________________________________________________


xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABSWID > " & oldYSWIECH0_SWISABSWID & " And SWISABSWID <= " & mYSWIECH0_SWISABSWID _
     & " and SWISABWMTK = '202' and SWISABWES = 'S'" _
     & " and swisabswid not in ( select swiechswiX from " & paramIBM_Library_SABSPE & ".yswiech0)" _
     & " and swisabswid not in ( select swiechswiD from " & paramIBM_Library_SABSPE & ".yswiech0)" _
     & " order by SWISABSWID"


Call YSWIECH0_Importation_MT(xSql)

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & currentAction
Exit_sub:

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, currentAction & " : exit"): DoEvents
'________________________________________________________________________

End Sub


Public Sub YSWI950_Importation()
Dim xSql As String

On Error GoTo Error_Handler
currentAction = "YSWI950_Importation "

'________________________________________________________________________
Call lstErr_Clear(lstErr, cmdContext, currentAction & " : 1"): DoEvents
'________________________________________________________________________


If mYSWI950_SWISABSWID = 0 Then
    xSql = "select SWI950SWID  from " & paramIBM_Library_SABSPE & ".YSWI950 " _
         & "order by SWI950SWID desc FETCH FIRST 1 ROWS ONLY"
    Set rsSab = cnsab.Execute(xSql)
    
    If Not rsSab.EOF Then mYSWI950_SWISABSWID = rsSab(0)
    If mYSWI950_SWISABSWID < 1298720 Then blnReprise = True: mYSWI950_SWISABSWID = 1298720 ' 1286467 '2015-10-01 ' test EMP 1201400 '
    
End If

Call rsYSWI950_Init(newYSWI950)

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABSWID > " & mYSWI950_SWISABSWID & " and SWISABWMTK = '950'" _
     & " order by SWISABSWID"
        
Set rsSab = cnsab.Execute(xSql)

Call YSWI950_Importation_MT



GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & currentAction
Exit_sub:

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, currentAction & " : exit"): DoEvents
'________________________________________________________________________

End Sub


Public Sub YSWI950_Importation_MT()
Dim blnExit As Boolean, K As Integer, K2 As Integer, K3 As Integer
Dim X As String, X2 As String

On Error GoTo Error_Handler

    
Do While Not rsSab.EOF
    newYSWI950.SWI950SWIL = 0
    mYSWI950_SWISABSWID = rsSab("SWISABSWID")
    newYSWI950.SWI950SWID = rsSab("SWISABSWID")
    newYSWI950.SWI950WBIC = Mid$(rsSab("SWISABWBIC"), 1, 8)
    newYSWI950.SWI950WES = rsSab("SWISABWES")
    newYSWI950.SWI950WDEV = rsSab("SWISABWDEV")
    Call YSWIRAM0_Fields(rsSab("SWISABWID1"), rsSab("SWISABWIDL"), rsSab("SWISABWIDH"))
    
    For K = 0 To fgDetail.Rows - 1
         fgDetail.Row = K
         fgDetail.Col = 0: xField_K = fgDetail.Text
         If xField_K <> "" Then newYSWI950.SWI950SWIL = newYSWI950.SWI950SWIL + 1
         If xField_K = "61" Then
            fgDetail.Col = 1: xField_V = fgDetail.Text
            X = Mid$(xField_V, 1, 6)
            If IsNumeric(X) Then
                newYSWI950.SWI950WVAL = Val(X) + 20000000
            Else
                newYSWI950.SWI950WVAL = rsSab("SWISABWAMJ")
            End If
            
            K2 = IIf(IsNumeric(Mid$(xField_V, 7, 1)), 11, 7)
            
            newYSWI950.SWI950SENS = Mid$(xField_V, K2, 1)
            K2 = K2 + 1
            K3 = K2
            blnExit = False
            Do
            
                If IsNumeric(Mid$(xField_V, K3, 1)) Then
                    K3 = K3 + 1
                Else
                    If K3 = K2 Then
                        K2 = K2 + 1
                        K3 = K2
                    Else
                        If Mid$(xField_V, K3, 1) = "," Then
                            K3 = K3 + 1
                        Else
                            blnExit = True
                        End If
                    End If
                End If
           Loop Until blnExit
            
           newYSWI950.SWI950WMTD = CCur(Mid$(xField_V, K2, K3 - K2))
            
           newYSWI950.SWI950WN20 = ""
           newYSWI950.SWI950WL20 = ""
           K2 = K3 + 4
           K3 = InStr(K2, xField_V, "//")
           If K3 > 0 Then
              X = Mid$(xField_V, K2, K3 - K2)
              If Trim(X) = "NONREF" Then X = ""
              X2 = Mid$(xField_V, K3 + 2, Len(xField_V) - K3 - 1)
           Else
              X2 = ""
              X = Mid$(xField_V, K2, Len(xField_V) - K2 + 1)
           End If
            
            If newYSWI950.SWI950WES = "E" Then
                newYSWI950.SWI950WN20 = X
                newYSWI950.SWI950WL20 = X2
            Else
                newYSWI950.SWI950WN20 = X2
                newYSWI950.SWI950WL20 = X
            End If
            
            mYSWI950_Fct = "New"
            Call YSWIECH0_Update
        End If
    Next K
        
    rsSab.MoveNext

Loop

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & currentAction
Exit_sub:

End Sub




Public Sub YSWIECH0_Match()
Dim xSql As String

On Error GoTo Error_Handler
currentAction = "YSWIECH0_Match"

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, currentAction & " : OPE_20"): DoEvents
'________________________________________________________________________

xSql = "select *  from " & paramIBM_Library_SABSPE & ".YSWIECH0E " _
     & " where SWIECHSTA = '#' and SWIECHDECH <= " & DSys _
     & " and SWIECHWMTK <> '950'" _
     & " order by SWIECHDECH, SWIECHOPEC, SWIECHOPEN"

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    Call rsYSWIECH0_GetBuffer(rsSab, oldYSWIECH0)
    Call YSWIECH0_Match_OPE_20
    rsSab.MoveNext

Loop

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, currentAction & " : OPE_21"): DoEvents
'________________________________________________________________________

xSql = "select *  from " & paramIBM_Library_SABSPE & ".YSWIECH0E " _
     & " where SWIECHSTA = '#' and SWIECHDECH <= " & DSys _
     & " and SWIECHWMTK <> '950'  and SWIECHW22C <> ''" _
     & " order by SWIECHDECH, SWIECHOPEC, SWIECHOPEN"

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    Call rsYSWIECH0_GetBuffer(rsSab, oldYSWIECH0)
    Call YSWIECH0_Match_OPE_21
    rsSab.MoveNext

Loop

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, currentAction & " : 950"): DoEvents
'________________________________________________________________________

xSql = "select *  from " & paramIBM_Library_SABSPE & ".YSWIECH0E " _
     & " where SWIECHSTA = '#' and SWIECHDECH <= " & DSys _
     & " and SWIECHWMTK = '950'" _
     & " order by SWIECHDECH, SWIECHOPEC, SWIECHOPEN"

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    Call rsYSWIECH0_GetBuffer(rsSab, oldYSWIECH0)
    Call YSWIECH0_Match_950
    rsSab.MoveNext

Loop

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, currentAction & " : Terminé"): DoEvents
'________________________________________________________________________


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & currentAction
Exit_sub:

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, currentAction & " : exit"): DoEvents
'________________________________________________________________________

End Sub

Public Sub YSWIECH0_Match_Ignore()
Dim xSql As String

On Error GoTo Error_Handler
currentAction = "YSWIECH0_Match_Ignore"

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, currentAction & " : 900 / 910"): DoEvents
'________________________________________________________________________

xSql = "select *  from " & paramIBM_Library_SABSPE & ".YSWIECH0E " _
     & " where SWIECHSTA = '#' and SWIECHDECH <= " & YBIATAB0_DATE_CPT_J _
     & " and SWIECHWMTK in ('900' , '910')" _
     & " order by SWIECHDECH, SWIECHOPEC, SWIECHOPEN"

'     & " and SWIECHWES = 'E' and SWIECHWMTK in ('900' , '910')" _

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    Call rsYSWIECH0_GetBuffer(rsSab, oldYSWIECH0)
    newYSWIECH0 = oldYSWIECH0
    newYSWIECH0.SWIECHSTA = "I"
    newYSWIECH0.SWIECHYAMJ = DSys
    newYSWIECH0.SWIECHYHMS = time_Hms
    newYSWIECH0.SWIECHYUSR = usrName_UCase
    mYSWIECH0_Fct = "Update"
    
    Call YSWIECH1_Info(newYSWIECH0.SWIECHSTA, dateImp10_S(oldYSWIECH0.SWIECHDECH) & " <= " & dateImp10_S(YBIATAB0_DATE_CPT_J))
    
    Call YSWIECH0_Update
    rsSab.MoveNext

Loop

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, currentAction & " : 950 < " & YBIATAB0_DATE_CPT_JP1): DoEvents
'________________________________________________________________________


mYSWI950_SQL_Set = " set SWI950SWIX = -1 where SWI950WVAL < " & YBIATAB0_DATE_CPT_JP1
mYSWI950_Fct = "SQL_Table"
'''''Call YSWIECH0_Update

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, currentAction & " : Terminé"): DoEvents
'________________________________________________________________________


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & currentAction
Exit_sub:

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, currentAction & " : exit"): DoEvents
'________________________________________________________________________

End Sub



Public Sub YSWIECH0_Match_OPE_20()
Dim xSql As String

On Error GoTo Error_Handler

If oldYSWIECH0.SWIECHWMTK = "198" Or oldYSWIECH0.SWIECHWMTK = "298" Then
    xSql = "select *  from " & paramIBM_Library_SABSPE & ".YSWISAB0N " _
         & " where SWISABOPEN = " & oldYSWIECH0.SWIECHOPEN & " and SWISABOPEC = '" & oldYSWIECH0.SWIECHOPEC & "'" _
         & " and SWISABWES = '" & oldYSWIECH0.SWIECHWES & "' and SWISABWMTK = '" & oldYSWIECH0.SWIECHWMTK & "'" _
         & " and SWISABWSTA = 'V'"
         
Else
    xSql = "select *  from " & paramIBM_Library_SABSPE & ".YSWISAB0N " _
         & " where SWISABOPEN = " & oldYSWIECH0.SWIECHOPEN & " and SWISABOPEC = '" & oldYSWIECH0.SWIECHOPEC & "'" _
         & " and SWISABWES = '" & oldYSWIECH0.SWIECHWES & "' and SWISABWMTK = '" & oldYSWIECH0.SWIECHWMTK & "'" _
         & " and SWISABWDEV = '" & oldYSWIECH0.SWIECHWDEV & "' and SWISABWMTD = '" & cur_P(oldYSWIECH0.SWIECHWMTD) & "'"
End If

Set rsSabX = cnsab.Execute(xSql)
    
'Do While Not rsSabX.EOF
If Not rsSabX.EOF Then
    newYSWIECH0 = oldYSWIECH0
    newYSWIECH0.SWIECHSWIX = rsSabX("SWISABSWID")
    newYSWIECH0.SWIECHSTA = " "
    newYSWIECH0.SWIECHSTAK = "="
    newYSWIECH0.SWIECHYAMJ = DSys
    newYSWIECH0.SWIECHYHMS = time_Hms
    newYSWIECH0.SWIECHYUSR = usrName_UCase
    mYSWIECH0_Fct = "Update"
    Call YSWIECH0_Update
End If
'    rsSabX.MoveNext

'Loop

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & currentAction
Exit_sub:

End Sub


Public Sub YSWIECH0_Match_OPE_21()
Dim xSql As String

On Error GoTo Error_Handler

xSql = "select *  from " & paramIBM_Library_SABSPE & ".YSWISAB0N " _
     & " where SWISABSWID > " & oldYSWIECH0.SWIECHSWID & " and SWISABWN20 = '" & oldYSWIECH0.SWIECHW22C & "'" _
     & " and SWISABWES = '" & oldYSWIECH0.SWIECHWES & "' and SWISABWMTK = '" & oldYSWIECH0.SWIECHWMTK & "'" _
     & " and SWISABWDEV = '" & oldYSWIECH0.SWIECHWDEV & "' and SWISABWMTD = '" & cur_P(oldYSWIECH0.SWIECHWMTD) & "'"
Set rsSabX = cnsab.Execute(xSql)
    
'Do While Not rsSabX.EOF
If Not rsSabX.EOF Then
    newYSWIECH0 = oldYSWIECH0
    newYSWIECH0.SWIECHSWIX = rsSabX("SWISABSWID")
    newYSWIECH0.SWIECHSTA = " "
    newYSWIECH0.SWIECHSTAK = "="
    
    newYSWIECH0.SWIECHYAMJ = DSys
    newYSWIECH0.SWIECHYHMS = time_Hms
    newYSWIECH0.SWIECHYUSR = usrName_UCase
    mYSWIECH0_Fct = "Update"
    Call YSWIECH0_Update
End If

'    rsSabX.MoveNext

'Loop

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & currentAction
Exit_sub:

End Sub


Public Sub YSWIECH0_Match_950()
Dim xSql As String
Dim wSTAK As String, blnOk As Boolean

On Error GoTo Error_Handler
xSql = "select *  from " & paramIBM_Library_SABSPE & ".YSWI950" _
     & " where SWI950WBIC = '" & Trim(oldYSWIECH0.SWIECHWBIC) & "'" _
     & " and SWI950WES = '" & oldYSWIECH0.SWIECHWES & "'" _
     & " and SWI950WDEV = '" & oldYSWIECH0.SWIECHWDEV & "'" _
     & " and SWI950WMTD = " & cur_P(oldYSWIECH0.SWIECHWMTD) _
     & " and SWI950SENS = '" & oldYSWIECH0.SWIECHSENS & "'" _
     & " order by SWI950SWID"
Set rsSabX = cnsab.Execute(xSql)
    
blnOk = False
Do While Not rsSabX.EOF
    
    If rsSabX("SWI950WN20") = oldYSWIECH0.SWIECHWN20 Or rsSabX("SWI950WN20") = oldYSWIECH0.SWIECHW22C Then
        blnOk = True
        If rsSabX("SWI950WVAL") = oldYSWIECH0.SWIECHW30V Then
           wSTAK = "="
        Else
            wSTAK = "~"
        End If
    Else
        If Abs(DateDiff("d", Format(rsSabX("SWI950WVAL"), "0000/00/00"), Format(oldYSWIECH0.SWIECHW30V, "0000/00/00"))) <= 3 Then
           wSTAK = "%"
           blnOk = True
        End If
    End If
    
    If blnOk Then
        newYSWIECH0 = oldYSWIECH0
        newYSWIECH0.SWIECHSWIX = rsSabX("SWI950SWID")
        newYSWIECH0.SWIECHSWIL = rsSabX("SWI950SWIL")
        newYSWIECH0.SWIECHSTA = " "
        newYSWIECH0.SWIECHSTAK = wSTAK
        newYSWIECH0.SWIECHYAMJ = DSys
        newYSWIECH0.SWIECHYHMS = time_Hms
        newYSWIECH0.SWIECHYUSR = usrName_UCase

        mYSWIECH0_Fct = "Update"
        
        newYSWI950.SWI950SWID = rsSabX("SWI950SWID")
        newYSWI950.SWI950SWIL = rsSabX("SWI950SWIL")
        mYSWI950_SQL_Set = " set SWI950SWIX = " & oldYSWIECH0.SWIECHSWID
        mYSWI950_Fct = "SQL"
        
        Call YSWIECH0_Update
        Exit Do
    End If
    rsSabX.MoveNext

Loop

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & currentAction
Exit_sub:

End Sub



Public Sub YSWIECH0_Importation_MT(xSql As String)
Dim blnOk As Boolean
On Error GoTo Error_Handler

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
    blnOk = True
    mYSWIECH0_Fct = ""
    If blnOk Then
        Call YSWIRAM0_Fields(oldYSWISAB0.SWISABWID1, oldYSWISAB0.SWISABWIDL, oldYSWISAB0.SWISABWIDH)
        
        Call rsYSWIECH0_Init(xYSWIECH0)
        
        mMTK = oldYSWISAB0.SWISABWMTK

        Call YSWIECH0_Importation_fgDetail
        
        xYSWIECH0.SWIECHSWID = oldYSWISAB0.SWISABSWID
        xYSWIECH0.SWIECHSEQ0 = 0
        xYSWIECH0.SWIECHSER = oldYSWISAB0.SWISABSER
        xYSWIECH0.SWIECHSSE = oldYSWISAB0.SWISABSSE
        xYSWIECH0.SWIECHOPEC = oldYSWISAB0.SWISABOPEC
        xYSWIECH0.SWIECHOPEN = oldYSWISAB0.SWISABOPEN
        xYSWIECH0.SWIECHSTA = "#"
        xYSWIECH0.SWIECHWN20 = wMT_20
        ''''xYSWIECH0.SWIECHWL20 = wMT_21
        xYSWIECH0.SWIECHW22C = wMT_22C
        xYSWIECH0.SWIECHYAMJ = DSys
        xYSWIECH0.SWIECHYHMS = time_Hms
        xYSWIECH0.SWIECHDECH = wMT_30V
        xYSWIECH0.SWIECHW30V = wMT_30V
        wMT_30V_JS1 = dateElp("Ouvré", 1, wMT_30V)
       If blnReprise Then
            xYSWIECH0.SWIECHYUSR = "BIA_INFO"
        Else
            xYSWIECH0.SWIECHYUSR = usrName_UCase
        End If
        
         Select Case oldYSWISAB0.SWISABWMTK
            Case "103":
                Select Case oldYSWISAB0.SWISABOPEC
                    Case "TRF": Call YSWIECH0_Importation_MT103_TRF
                    Case "CPT": Call YSWIECH0_Importation_MT103_CPT
                    Case "CDE": Call YSWIECH0_Importation_MT103_CDE
                    Case Else: Call YSWIECH1_Exclus("?", "103 " & oldYSWISAB0.SWISABOPEC & " " & oldYSWISAB0.SWISABOPEN)
                End Select
            Case "202":
                Select Case oldYSWISAB0.SWISABOPEC
                    Case "TRF": Call YSWIECH0_Importation_MT202
                End Select
            Case "300":
                If wMT_22A <> "NEWT" Then
                    Call YSWIECH1_Exclus("X", "22A <> 'NEWT'")
                Else
                    Call YSWIECH0_Importation_MT300_B1
                    Call YSWIECH0_Importation_MT300_B2
                End If
            Case "320"
                 If wMT_22A <> "NEWT" Then
                    Call YSWIECH1_Exclus("X", "22A <> 'NEWT'")
                Else
                    If wMT_17R = "L" Then
                         Call YSWIECH0_Importation_MT320_PRE_MAD
                         Call YSWIECH0_Importation_MT320_PRE_RBT
                     Else
                         If wMT_22B = "ROLL" Then
                             Call YSWIECH1_Exclus("X", "22B = 'ROLL'")
                         Else
                             Call YSWIECH0_Importation_MT320_EMP_MAD
                             Call YSWIECH0_Importation_MT320_EMP_RBT
                         End If
                    End If
                
                End If
    
        End Select
    End If
    
    rsSab.MoveNext

Loop


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : YSWIECH0_Importation_MT"
Exit_sub:

    

End Sub


Public Sub YSWIECH0_Importation_fgDetail()
Dim K As Integer
On Error GoTo Error_Handler

mMTK_Seq = ""
wMT_82A = "": wMT_87A = ""
wMT_20 = "": wMT_21 = "": wMT_30V = ""
wMT_22A = "": wMT_22C = ""
wMT_53A_B1 = "": wMT_53A_B2 = ""
wMT_57A_B1 = "": wMT_57A_B2 = ""
wMT_32_DEV = "": wMT_33_DEV = ""
wMT_32_MTD = 0: wMT_33_MTD = 0
wMT_52A = "": wMT_57A = "": wMT_54A = ""

For K = 0 To fgDetail.Rows - 1
     fgDetail.Row = K
     fgDetail.Col = 0: xField_K = fgDetail.Text
     fgDetail.Col = 1: xField_V = fgDetail.Text
     Call YSWIRAM0_Field_V
     
     Select Case mMTK
          Case "103", "202":
            Select Case xField_K
             Case "20": wMT_20 = xField_V
             Case "21": wMT_21 = xField_V
             Case "32A"
                wMT_30V = Mid$(xField_V, 1, 6) + 20000000
                wMT_32_DEV = Mid$(xField_V, 7, 3)
                wMT_32_MTD = num_CDec(Mid$(xField_V, 10, Len(xField_V) - 9))
             Case "52A": wMT_52A = xField_V
             Case "57A": wMT_57A = xField_V
             Case "54A": wMT_54A = xField_V
            End Select
        Case "300":
            Select Case xField_K
             Case "20": wMT_20 = xField_V
             Case "21": wMT_21 = xField_V
             Case "22A": wMT_22A = xField_V
             Case "22C": wMT_22C = xField_V
             Case "30V": wMT_30V = xField_V
             Case "82A": wMT_82A = xField_V
             Case "87A": wMT_87A = xField_V
             Case "53A-B1": wMT_53A_B1 = xField_V
             Case "57A-B1": wMT_57A_B1 = xField_V
             Case "53A-B2": wMT_53A_B2 = xField_V
             Case "57A-B2": wMT_57A_B2 = xField_V
             Case "32B-B1"
                wMT_32_DEV = Mid$(xField_V, 1, 3)
                wMT_32_MTD = num_CDec(Mid$(xField_V, 4, Len(xField_V) - 3))
             Case "33B-B2"
                wMT_33_DEV = Mid$(xField_V, 1, 3)
                wMT_33_MTD = num_CDec(Mid$(xField_V, 4, Len(xField_V) - 3))
            End Select
          Case "320":
            Select Case xField_K
             Case "17R": wMT_17R = xField_V
             Case "20": wMT_20 = xField_V
             Case "21": wMT_21 = xField_V
             Case "22A": wMT_22A = xField_V
             Case "22B": wMT_22B = xField_V
             Case "22C": wMT_22C = xField_V
             Case "30V": wMT_30V = xField_V
             Case "30X": wMT_30X = xField_V
             Case "82A": wMT_82A = xField_V
             Case "87A": wMT_87A = xField_V
             Case "53A-C": wMT_53A_B1 = xField_V
             Case "57A-C": wMT_57A_B1 = xField_V
             Case "53A-D": wMT_53A_B2 = xField_V
             Case "57A-D": wMT_57A_B2 = xField_V
             Case "32B"
                wMT_32_DEV = Mid$(xField_V, 1, 3)
                wMT_32_MTD = num_CDec(Mid$(xField_V, 4, Len(xField_V) - 3))
             Case "34E"
                If IsNumeric(Mid$(xField_V, 4, 1)) Then
                    wMT_34_N = ""
                    wMT_34_DEV = Mid$(xField_V, 1, 3)
                    wMT_34_MTD = num_CDec(Mid$(xField_V, 4, Len(xField_V) - 3))
                Else
                    wMT_34_N = Mid$(xField_V, 1, 1)
                    wMT_34_DEV = Mid$(xField_V, 2, 3)
                    wMT_34_MTD = num_CDec(Mid$(xField_V, 5, Len(xField_V) - 4))
                End If
                
            End Select
                        
     End Select
     
Next K
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : YSWIECH0_Importation_fgDetail"
Exit_sub:

End Sub



Public Sub YSWIRAM0_Importation_MT(xSql As String)
Dim blnOk As Boolean
On Error GoTo Error_Handler

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
    blnOk = True
    If oldYSWISAB0.SWISABWMTK = "399" And oldYSWISAB0.SWISABWES = "S" Then blnOk = False
    If InStr(oldYSWISAB0.SWISABWBIC, "EMIDITMM") > 0 Then blnOk = False

    If blnOk Then
        Call YSWIRAM0_Fields(oldYSWISAB0.SWISABWID1, oldYSWISAB0.SWISABWIDL, oldYSWISAB0.SWISABWIDH)
        
        Call rsYSWIRAM0_Init(xYSWIRAM0)
        
        If oldYSWISAB0.SWISABOPEN <> 0 Then
            xYSWIRAM0.SWIRAMXOPE = Trim(oldYSWISAB0.SWISABSER & oldYSWISAB0.SWISABSSE & oldYSWISAB0.SWISABOPEC & Format(oldYSWISAB0.SWISABOPEN, "000000000"))
        End If
        xYSWIRAM0.SWIRAMXID = oldYSWISAB0.SWISABSWID
        xYSWIRAM0.SWIRAMXBIC = Mid$(oldYSWISAB0.SWISABWBIC, 1, 8)
        xYSWIRAM0.SWIRAMXMTK = oldYSWISAB0.SWISABWMTK
        xYSWIRAM0.SWIRAMXES = oldYSWISAB0.SWISABWES
        xYSWIRAM0.SWIRAMYAMJ = oldYSWISAB0.SWISABWAMJ
        xYSWIRAM0.SWIRAMYHMS = oldYSWISAB0.SWISABWHMS
        xYSWIRAM0.SWIRAMSTA = "#"
       If blnReprise Then
            xYSWIRAM0.SWIRAMYUSR = "BIA_INFO"
        Else
            xYSWIRAM0.SWIRAMYUSR = usrName_UCase
        End If
        mMTK = oldYSWISAB0.SWISABWMTK
        Call YSWIRAM0_Importation_fgDetail
        
        Select Case oldYSWISAB0.SWISABWMTK
            Case "300", "320":
                Call YSWIRAM0_Match_300
            Case "399"
                xYSWIRAM0.SWIRAMXREF = oldYSWISAB0.SWISABWL20
                mYSWIRAM0_Fct = "New": newYSWIRAM0 = xYSWIRAM0
                Call YSWIRAM0_Update
    
        End Select
    End If
    
    rsSab.MoveNext

Loop

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : YSWIRAM0_Importation_MT"
Exit_sub:

    

End Sub


Public Sub YSWIRAM0_Fields(Mesg_aid As Long, mesg_s_umidl As Long, mesg_s_umidh As Long)
Dim xSql As String

On Error GoTo Error_Handler
fgDetail_Reset

fgDetail.Rows = 1

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

        'Debug.Print rsSIDE_DB("field_code") & rsSIDE_DB("field_option") & ": " & rsSIDE_DB("value")
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
        fgDetail_DisplayLine_rText fgDetail.Row
        'Debug.Print xrText.text_data_block
    End If
End If

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    Dim V
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

End Sub

Public Sub YJPLSLD1_Exportation()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim mBackColor As Long, mForeColor As Long
Dim xSql As String, Nb As Long
'______________________________________________
Nb = 0
Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "Inspection"
    .Subject = "Liquidité"
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "Inspection"

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
    .Font.Size = 8
    .Font.Name = "Arial Unicode MS"
End With

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

wsExcel.Cells(1, 1) = "Opé": wsExcel.Columns(1).ColumnWidth = 7
wsExcel.Cells(1, 2) = "numéro": wsExcel.Columns(2).ColumnWidth = 9: wsExcel.Columns(2).NumberFormat = "### ### ### ###"
wsExcel.Cells(1, 3) = "Devise": wsExcel.Columns(3).ColumnWidth = 5
wsExcel.Cells(1, 4) = "PCI": wsExcel.Columns(4).ColumnWidth = 7
wsExcel.Cells(1, 5) = "Compte": wsExcel.Columns(5).ColumnWidth = 20
wsExcel.Cells(1, 6) = "solde dev": wsExcel.Columns(6).ColumnWidth = 15: wsExcel.Columns(6).NumberFormat = "### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(1, 7) = "solde €": wsExcel.Columns(7).ColumnWidth = 15: wsExcel.Columns(7).NumberFormat = "### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(1, 8) = "montant initial": wsExcel.Columns(8).ColumnWidth = 15: wsExcel.Columns(8).NumberFormat = "### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(1, 9) = "date début": wsExcel.Columns(9).ColumnWidth = 10: wsExcel.Columns(9).NumberFormat = "mm/dd/yyyy"
wsExcel.Cells(1, 10) = "date fin": wsExcel.Columns(10).ColumnWidth = 10: wsExcel.Columns(10).NumberFormat = "mm/dd/yyyy"
wsExcel.Cells(1, 11) = "Nb jours": wsExcel.Columns(11).ColumnWidth = 7: wsExcel.Columns(11).NumberFormat = "### ### ### ###"
wsExcel.Cells(1, 12) = "DRC": wsExcel.Columns(12).ColumnWidth = 7: wsExcel.Columns(12).NumberFormat = "### ### ### ###"
wsExcel.Cells(1, 13) = "Nature": wsExcel.Columns(13).ColumnWidth = 5
wsExcel.Cells(1, 14) = "Client": wsExcel.Columns(14).ColumnWidth = 10
wsExcel.Cells(1, 15) = "Type client": wsExcel.Columns(15).ColumnWidth = 5
wsExcel.Cells(1, 16) = "Intitulé": wsExcel.Columns(16).ColumnWidth = 40


xSql = "select * from " & paramIBM_Library_SABSPE & ".YJPLSLD1 "

Set rsSab = cnsab.Execute(xSql)
Nb = 1
Do While Not rsSab.EOF
    Nb = Nb + 1
        wsExcel.Cells(Nb, 1) = rsSab("DOSSLDOPE")
        wsExcel.Cells(Nb, 2) = rsSab("DOSSLDNUM")
        wsExcel.Cells(Nb, 3) = rsSab("DOSSLDDEV")
        wsExcel.Cells(Nb, 4) = rsSab("DOSSLDPCI")
        wsExcel.Cells(Nb, 5) = rsSab("DOSSLDCPT")
        wsExcel.Cells(Nb, 6) = rsSab("DOSSLDSLD")
        wsExcel.Cells(Nb, 7) = rsSab("DOSSLDSLDE")
        wsExcel.Cells(Nb, 8) = rsSab("DOSSLDMTD")
        wsExcel.Cells(Nb, 9) = dateImp10(rsSab("DOSSLDDDEB"))
        wsExcel.Cells(Nb, 10) = dateImp10(rsSab("DOSSLDDFIN"))
        wsExcel.Cells(Nb, 11) = rsSab("DOSSLDDUR0")
        wsExcel.Cells(Nb, 12) = rsSab("DOSSLDDURJ")
        wsExcel.Cells(Nb, 13) = rsSab("DOSSLDNAT")
        wsExcel.Cells(Nb, 14) = rsSab("DOSSLDCLI")
        wsExcel.Cells(Nb, 15) = rsSab("DOSSLDCLIK")
        wsExcel.Cells(Nb, 16) = rsSab("DOSSLDCLIL")

    rsSab.MoveNext
Loop

wbExcel.SaveAs "C:\Temp\Inspection_Liquidité.xlsx"

wbExcel.Close


Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
    Set rsSab = Nothing

End Sub

Public Sub fgEVE_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgEVE.Visible = False: fraDetail.Visible = False
mRow = fgEVE.Row

If lRow > 0 And lRow < fgEVE.Rows Then
    fgEVE.Row = lRow
    For I = 3 To fgEVE.FixedCols Step -1
        fgEVE.Col = I: fgEVE.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgEVE.Row = mRow
    If fgEVE.Row > 0 Then
        lRow = fgEVE.Row
        lColor_Old = fgEVE.CellBackColor
        For I = 3 To fgEVE.FixedCols Step -1
          fgEVE.Col = I: fgEVE.CellBackColor = lColor
        Next I
    End If
End If
fgEVE.LeftCol = fgEVE.FixedCols
fgEVE.Visible = True: fraDetail.Visible = True
End Sub
Private Sub fgEVE_Display()
Dim wColor As Long
Dim X As String, xWhere As String, xOPE As String
Dim xSql As String
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
fgEVE.Visible = False: fraDetail.Visible = False
fgEVE_Reset

fgEVE.Rows = 1
fgEVE.FormatString = fgEVE_FormatString
fgEVE.Row = 0

currentAction = "fgEVE_Display"

For I = 1 To arrYGOSEVE0_Nb
    xYGOSEVE0 = arrYGOSEVE0(I)
    fgEVE.Rows = fgEVE.Rows + 1
    fgEVE.Row = fgEVE.Rows - 1
    fgEVE_DisplayLine I


Next I

fgEVE.Visible = True: fraDetail.Visible = True
'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub fgEVE_DisplayLine(lIndex As Long)
Dim K As Integer, wColor As Long, wBackColor As Long
Dim HeightOfLine As Long, LinesOfText As Long
Dim xGOSEVENAT As String


On Error Resume Next
'On Error GoTo Error_Handler
xGOSEVENAT = Trim(xYGOSEVE0.GOSEVENAT)
wBackColor = 0
If xYGOSEVE0.GOSEVESTAE = " " Then
     Select Case xGOSEVENAT
        Case "Sus*":
            Select Case oldYGOSDOS0.GOSDOSSTAG
                Case "R": wColor = vbYellow: wBackColor = vbRed  'mColor_W1
                Case "V": wColor = RGB(0, 0, 128): wBackColor = mColor_G2
                Case Else: wColor = RGB(0, 0, 128): wBackColor = RGB(255, 255, 128)
            End Select
        Case "Note": wColor = RGB(0, 0, 128): wBackColor = RGB(255, 255, 192)
        Case "Res*", "AnnV", "AnnR", "AnnC": wColor = vbMagenta: wBackColor = mColor_Y1
        Case "Mail": wColor = RGB(128, 128, 128)
        Case "PJ**", "Swi+": wColor = RGB(0, 0, 96): wBackColor = RGB(220, 255, 255)
        Case "Swi>":
                    If xYGOSEVE0.GOSEVESWID > 0 Then
                        wColor = RGB(0, 64, 0)
                    Else
                        wColor = vbRed
                    End If
                    
        Case "Val": wColor = RGB(0, 0, 128): wBackColor = mColor_G2 'RGB(220, 255, 220)
        Case "Rej": wColor = vbYellow: wBackColor = vbRed 'RGB(255, 112, 220)
        Case "Clo": wColor = RGB(0, 0, 128): wBackColor = RGB(220, 220, 220)
       Case Else: wColor = vbMagenta 'RGB(0, 96, 96) '&HC0&       'RGB(0, 80, 0)
    End Select
Else
     Select Case xYGOSEVE0.GOSEVESTAE
        Case "A": wColor = RGB(164, 164, 164)
        Case Else: wColor = vbRed
     End Select

End If


fgEVE.Col = 0: fgEVE.Text = xYGOSEVE0.GOSEVEIDE & " - " & dateImp10_S(xYGOSEVE0.GOSEVEUAMJ)

fgEVE.Col = 1
K = Val(Mid$(xYGOSEVE0.GOSEVEUSRV, 2, 2)): fgEVE.Text = arrService_Lib(K)

If xYGOSEVE0.GOSEVEUSRV <> xYGOSEVE0.GOSEVEGSRV Then
    fgEVE.Col = 2
    K = Val(Mid$(xYGOSEVE0.GOSEVEGSRV, 2, 2)): fgEVE.Text = arrService_Lib(K)
    fgEVE.CellForeColor = wColor
    'fgEVE.CellFontBold = True
End If
fgEVE.Col = 3: fgEVE.Text = xYGOSEVE0.GOSEVENAT & " " & xYGOSEVE0.GOSEVESTAE

fgEVE.Col = 4

fgEVE.Text = Trim(xYGOSEVE0.GOSEVETXT)
txtFg = fgEVE.Text
'If xYGOSEVE0.GOSEVENAT <> "Swi>" Then
Select Case xGOSEVENAT
    Case "Sus*", "Note", "Val", "Rej":
        HeightOfLine = fgEVE.RowHeightMin / 3 + 100 '- 20 'Me.TextHeight(txteve.Text)
    
        LinesOfText = SendMessage(txtFg.hwnd, EM_GETLINECOUNT, 0&, 0&) + 1
        
        If fgEVE.RowHeight(fgEVE.Row) < (LinesOfText * HeightOfLine) Then
           fgEVE.RowHeight(fgEVE.Row) = LinesOfText * HeightOfLine
        End If
End Select

For K = 0 To fgEVE_arrIndex
    fgEVE.Col = K
    fgEVE.CellBackColor = wBackColor
    fgEVE.CellForeColor = wColor
        If xGOSEVENAT = "Sus*" Then fgEVE.CellFontBold = True
Next K

'Select Case xYGOSEVE0.GOSEVENAT
'   Case "Sus*", "Note": For K = 0 To fgEVE_arrIndex: fgEVE.Col = K: fgEVE.CellBackColor = mColor_B0: Next K
'   Case "PJ**", "Swi+": For K = 0 To fgEVE_arrIndex: fgEVE.Col = K: fgEVE.CellBackColor = mColor_G0: Next K
'   Case "Swi>": For K = 0 To fgEVE_arrIndex: fgEVE.Col = K: fgEVE.CellBackColor = mColor_Y1: Next K
'End Select
        
If xYGOSEVE0.GOSEVESTAE = "?" Then
    For K = 0 To fgEVE_arrIndex
        fgEVE.Col = K
        fgEVE.CellForeColor = vbYellow
        fgEVE.CellBackColor = vbRed
    Next K
End If

fgEVE.Col = fgEVE_arrIndex: fgEVE.Text = lIndex
'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Public Sub fgEVE_Reset()
fgEVE.Clear
fgEVE_Sort1 = 0: fgEVE_Sort2 = 0
fgEVE_Sort1_Old = -1
fgEVE_RowDisplay = 0: fgEVE_RowClick = 0
fgEVE_arrIndex = fgEVE.Cols - 1
blnfgEVE_DisplayLine = False
fgEVE_SortAD = 6
fgEVE.LeftCol = fgEVE.FixedCols

End Sub




Public Sub fgEVE_Sort()
If fgEVE.Rows > 1 Then
    fgEVE.Row = 1
    fgEVE.RowSel = fgEVE.Rows - 1
    
    If fgEVE_Sort1_Old = fgEVE_Sort1 Then
        If fgEVE_SortAD = 5 Then
            fgEVE_SortAD = 6
        Else
            fgEVE_SortAD = 5
        End If
    Else
        fgEVE_SortAD = 5
    End If
    fgEVE_Sort1_Old = fgEVE_Sort1
    
    fgEVE.Col = fgEVE_Sort1
    fgEVE.ColSel = fgEVE_Sort2
    fgEVE.Sort = fgEVE_SortAD
End If

End Sub



'______________________________________________________________________
Private Sub fgSelect_Display()
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim m_Aid As Integer, m_mesg_s_umidl  As Long, m_mesg_s_umidh  As Long
    
 
 
On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = Replace(fgSelect_FormatString, "Information", "Date valeur")
fgSelect.Row = 0

fgSelect.Col = 3: fgSelect.CellAlignment = 1

currentAction = "fgSelect_Display"
    
Do While Not rsSIDE_DB.EOF
    'V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
    If m_Aid = rsSIDE_DB("aid") _
    And m_mesg_s_umidl = rsSIDE_DB("mesg_s_umidl") _
    And m_mesg_s_umidh = rsSIDE_DB("mesg_s_umidh") Then
    
    Else

        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I

            m_Aid = rsSIDE_DB("aid")
            m_mesg_s_umidl = rsSIDE_DB("mesg_s_umidl")
            m_mesg_s_umidh = rsSIDE_DB("mesg_s_umidh")
    End If
    rsSIDE_DB.MoveNext

Loop
         
If blnSelect_SQL_1_rText Then
    Set rsSIDE_DB = cnSIDE_DB.Execute(cmdSelect_SQL_1_rText)
    Do While Not rsSIDE_DB.EOF
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
        rsSIDE_DB.MoveNext
    Loop
    
End If
         
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

Private Sub fgSelect_Display_YSWISAB0()
Dim wColor As Long, X As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
Select Case cmdSelect_SQL_K
    Case "1trf":
        X = Replace(fgSelect_FormatString, "Information", "EBA / TGT")
        X = Replace(X, "Date de réception SAA", "Donneur d'ordre")
        
        fgSelect.FormatString = Replace(X, "statut      ", "Bénéficiaire")
    Case "6", "6E", "6#":
        fgSelect.FormatString = Replace(fgSelect_FormatString, "Date de réception SAA       ", "date dernier événement RAM")
        mYSWIRAM0_Match_XOPE = ""
End Select


fgSelect.Row = 0
fgSelect.Col = 3: fgSelect.CellAlignment = 1

currentAction = "fgSelect_Display_YSWISAB0"

mYSWILNK0_Display = ""

Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine_YSWISAB0 I
    
    If blnYSWILNK0_Display Then
        If rsSab("SWISABXGOS") = "G" Or rsSab("SWISABXEVE") = "G" Then
            mYSWILNK0_Display = mYSWILNK0_Display & " " & rsSab("SWISABSWID") & " ,"
        End If
    End If

    rsSab.MoveNext

Loop
         
If blnYSWILNK0_Display And mYSWILNK0_Display <> "" Then
    fgSelect_Display_YSWISAB0_YSWILNK0
End If

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgSelect_Display_YGOSDOS0()
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<N°- Type    |<BIC émis/reçu          |<référence                           |" _
                      & ">Montant               |<Dev    |<Etat |<Echéance     |<Srv Initiateur       |<Srv en charge|" _
                      & "<RCOM            |<Client        |<Motif du suspens         |"

fgSelect.Row = 0

currentAction = "fgSelect_Display"
    
For I = 1 To arrYGOSDOS0_Nb
    xYGOSDOS0 = arrYGOSDOS0(I)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine_YGOSDOS0 I


Next I
    

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYGOSDOS0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgSelect_Display_YSWIECH0()
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Opération            |<BIC                     |<Echéance      |<Msg        " _
                      & "|>Montant               |<Dev    |<Réf commune          |<52A                 |<57A               " _
                      & "|<Etat |<MàJ                                                           |<L/Réf                          |<N/Réf                            " _
                      & "|>swid|>swix|>swil|>S0"
fgSelect.Row = 0
fgSelect.Col = 4: fgSelect.CellAlignment = 1
currentAction = "fgSelect_Display_YSWIECH0"
    
Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    Call rsYSWIECH0_GetBuffer(rsSab, xYSWIECH0)
    fgSelect_DisplayLine_YSWIECH0 I
    
    rsSab.MoveNext

Loop
    

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgSelect_Display_YSWIECH1()
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Opération            |<BIC                     |<Date              |<Msg     " _
                      & "|>Montant               |<Dev    |<Fonction |<Motif                                      |<L/Réf                          |<N/Réf                            " _
                      & "|<MàJ                                                           |>swid            "
fgSelect.Row = 0
fgSelect.Col = 4: fgSelect.CellAlignment = 1
currentAction = "fgSelect_Display_YSWIECH1"
    
Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    Call rsYSWIECH1_GetBuffer(rsSab, xYSWIECH1)
    Call rsYSWISAB0_GetBuffer(rsSab, xYSWISAB0)
    If Not IsNull(rsSab("SWIECHSWID")) Then
        Call rsYSWIECH0_GetBuffer(rsSab, xYSWIECH0)
    Else
        xYSWIECH0.SWIECHSWID = 0
    End If
    fgSelect_DisplayLine_YSWIECH1 I
    
    rsSab.MoveNext

Loop
    

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub fgSelect_DisplayLine_YSWIECH0(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wBackColor As Long
On Error Resume Next

If xYSWIECH0.SWIECHWES = "S" Then
    wColor = RGB(16, 96, 16)
Else
    wColor = vbBlue
End If
Select Case xYSWIECH0.SWIECHSTA
    Case " ": wBackColor = vbWhite 'RGB(245, 255, 245)
    Case "#":
        If xYSWIECH0.SWIECHDECH < DSys Then
            wBackColor = mColor_Y2
        Else
            wBackColor = mColor_Y1
        End If
        
    Case "I": wBackColor = RGB(230, 230, 230)
    Case "A": wBackColor = RGB(210, 210, 210)
    Case "E": wBackColor = mColor_W1
    Case Else: wBackColor = mColor_B0
End Select
    
fgSelect.Col = 0: fgSelect.Text = xYSWIECH0.SWIECHSER & "-" & xYSWIECH0.SWIECHSSE & " " & xYSWIECH0.SWIECHOPEC & " " & xYSWIECH0.SWIECHOPEN
If xYSWIECH0.SWIECHWMTK = "950" Then
    fgSelect.Col = 3: fgSelect.Text = xYSWIECH0.SWIECHWMTK & " " & xYSWIECH0.SWIECHWES & " " & xYSWIECH0.SWIECHSENS
Else
    fgSelect.Col = 3: fgSelect.Text = xYSWIECH0.SWIECHWMTK & " " & xYSWIECH0.SWIECHWES
End If

fgSelect.Col = 1: fgSelect.Text = xYSWIECH0.SWIECHWBIC
fgSelect.Col = 2: fgSelect.Text = dateImp10_S(xYSWIECH0.SWIECHDECH)
fgSelect.Col = 4: fgSelect.Text = Format$(xYSWIECH0.SWIECHWMTD, "### ### ### ##0.00")
fgSelect.Col = 5: fgSelect.Text = xYSWIECH0.SWIECHWDEV


fgSelect.Col = 12: fgSelect.Text = xYSWIECH0.SWIECHWN20
fgSelect.Col = 11: fgSelect.Text = xYSWIECH0.SWIECHWL20
fgSelect.Col = 6: fgSelect.Text = xYSWIECH0.SWIECHW22C
fgSelect.Col = 7: fgSelect.Text = xYSWIECH0.SWIECHW52A
fgSelect.Col = 8: fgSelect.Text = xYSWIECH0.SWIECHW57A
fgSelect.Col = 13: fgSelect.Text = xYSWIECH0.SWIECHSWID
fgSelect.Col = 14: fgSelect.Text = xYSWIECH0.SWIECHSWIX
fgSelect.Col = 15: fgSelect.Text = xYSWIECH0.SWIECHSWIL
fgSelect.Col = 16: fgSelect.Text = xYSWIECH0.SWIECHSEQ0

For K = 0 To 9
    fgSelect.Col = K
    fgSelect.CellForeColor = wColor
    fgSelect.CellBackColor = wBackColor
Next K
For K = 10 To 16
    fgSelect.Col = K
    fgSelect.CellForeColor = RGB(128, 128, 128)
    fgSelect.CellBackColor = wBackColor
Next K

fgSelect.Col = 9: fgSelect.Text = xYSWIECH0.SWIECHSTA & xYSWIECH0.SWIECHSTAK
'fgSelect.CellForeColor = vbBlack
fgSelect.Col = 10: fgSelect.Text = dateImp10_S(xYSWIECH0.SWIECHYAMJ) & "  " & timeImp8(xYSWIECH0.SWIECHYHMS) & xYSWIECH0.SWIECHYUSR & " v" & xYSWIECH0.SWIECHYVER
'fgSelect.CellForeColor = vbBlack

'fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
End Sub
Public Sub fgSelect_DisplayLine_YSWIECH1(lIndex As Long)
Dim K As Integer, K2 As Integer, X As String
Dim wColor As Long, wBackColor As Long
On Error Resume Next

If xYSWISAB0.SWISABWES = "S" Then
    wColor = RGB(16, 96, 16)
Else
    wColor = vbBlue
End If
wBackColor = vbWhite
    
fgSelect.Col = 0: fgSelect.Text = xYSWISAB0.SWISABSER & "-" & xYSWISAB0.SWISABSSE & " " & xYSWISAB0.SWISABOPEC & " " & xYSWISAB0.SWISABOPEN

K = InStr(xYSWIECH1.SWIEC1INFO, "<FCT:") + 5
If K > 5 Then
    K2 = InStr(K, xYSWIECH1.SWIEC1INFO, ">")
    X = Mid$(xYSWIECH1.SWIEC1INFO, K, K2 - K)
    Select Case X
        Case "X": X = "Exclus": wBackColor = mColor_Y1
        Case "I": X = "Ignoré": wBackColor = RGB(220, 220, 220)
        Case "A": X = "Annulé": wBackColor = mColor_W1
        Case "#": X = "Restauré": wBackColor = mColor_G1
        Case "?": X = "non géré": wBackColor = mColor_W0
    End Select
    fgSelect.Col = 6: fgSelect.Text = X
    fgSelect.CellBackColor = wBackColor
End If
K = InStr(K2, xYSWIECH1.SWIEC1INFO, "<X:") + 3
If K > 3 Then
    K2 = InStr(K, xYSWIECH1.SWIEC1INFO, ">")
    fgSelect.Col = 7: fgSelect.Text = Mid$(xYSWIECH1.SWIEC1INFO, K, K2 - K)
    fgSelect.CellBackColor = wBackColor
End If

If xYSWIECH0.SWIECHSWID = 0 Then
    fgSelect.Col = 1: fgSelect.Text = xYSWISAB0.SWISABWBIC
    fgSelect.Col = 2: fgSelect.Text = dateImp10_S(xYSWISAB0.SWISABWAMJ)
    fgSelect.Col = 3: fgSelect.Text = xYSWISAB0.SWISABWMTK
    fgSelect.CellBackColor = wBackColor
    fgSelect.Col = 4: fgSelect.Text = Format$(xYSWISAB0.SWISABWMTD, "### ### ### ##0.00")
    fgSelect.Col = 5: fgSelect.Text = xYSWISAB0.SWISABWDEV
    fgSelect.Col = 9: fgSelect.Text = xYSWISAB0.SWISABWN20
    fgSelect.Col = 8: fgSelect.Text = xYSWISAB0.SWISABWL20
Else
    fgSelect.Col = 1: fgSelect.Text = xYSWIECH0.SWIECHWBIC
    fgSelect.Col = 2: fgSelect.Text = dateImp10_S(xYSWIECH0.SWIECHDECH)
    fgSelect.Col = 3: fgSelect.Text = xYSWIECH0.SWIECHWMTK
    fgSelect.CellBackColor = wBackColor
    fgSelect.Col = 4: fgSelect.Text = Format$(xYSWIECH0.SWIECHWMTD, "### ### ### ##0.00")
    fgSelect.Col = 5: fgSelect.Text = xYSWIECH0.SWIECHWDEV
    fgSelect.Col = 9: fgSelect.Text = xYSWIECH0.SWIECHWN20
    fgSelect.Col = 8: fgSelect.Text = xYSWIECH0.SWIECHWL20
End If



For K = 0 To 9
    fgSelect.Col = K
    fgSelect.CellForeColor = wColor
Next K

fgSelect.Col = 10: fgSelect.Text = dateImp10_S(xYSWIECH1.SWIEC1YAMJ) & "  " & timeImp8(xYSWIECH1.SWIEC1YHMS) & xYSWIECH1.SWIEC1YUSR & " v" & xYSWIECH1.SWIEC1YVER
fgSelect.CellForeColor = RGB(128, 128, 128)
fgSelect.Col = 11: fgSelect.Text = xYSWIECH1.SWIEC1SWID
fgSelect.CellForeColor = RGB(128, 128, 128)

'fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
End Sub


Private Sub fgSelect_Display_Echéancier()
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<N°- Type    |<BIC émis/reçu       |<référence |" _
                      & ">Montant               |<Dev    |<Etat |<Echéance     |<Srv Initiateur       |<Srv en charge|" _
                      & "<Action |<par                |<date /heure                           |"

fgSelect.Row = 0

currentAction = "fgSelect_Display"
    
Call rsYGOSDOS0_Init(xYGOSDOS0)
ReDim arrYGOSDOS0(201)
arrYGOSDOS0_Max = 200: arrYGOSDOS0_Nb = 0


Do While Not rsSab.EOF
    If xYGOSDOS0.GOSDOSIDD = rsSab("GOSDOSIDD") Then
        V = rsYGOSEVE0_GetBuffer(rsSab, xYGOSEVE0)
    Else
        If xYGOSDOS0.GOSDOSIDD > 0 Then
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect_DisplayLine_Echéancier arrYGOSDOS0_Nb
        End If
        V = rsYGOSDOS0_GetBuffer(rsSab, xYGOSDOS0)
         arrYGOSDOS0_Nb = arrYGOSDOS0_Nb + 1
         If arrYGOSDOS0_Nb > arrYGOSDOS0_Max Then
             arrYGOSDOS0_Max = arrYGOSDOS0_Max + 100
             ReDim Preserve arrYGOSDOS0(arrYGOSDOS0_Max)
         End If
         
         arrYGOSDOS0(arrYGOSDOS0_Nb) = xYGOSDOS0

        V = rsYGOSEVE0_GetBuffer(rsSab, xYGOSEVE0)
    End If
    
    rsSab.MoveNext
Loop
If xYGOSDOS0.GOSDOSIDD > 0 Then
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine_Echéancier arrYGOSDOS0_Nb
End If

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYGOSDOS0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgSelect_Display_4Journal()
Dim wColor As Long
Dim mGOSDOSISRV As String, mGOSEVEUSRV As String
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<N°- Type    |Service       |<BIC émis/reçu     |" _
                      & ">Montant Dev               |<Etat       |" _
                      & "<Information                                                                                                                                    |||"

fgSelect.Row = 0

currentAction = "fgSelect_Display_4Journal"
Call rsYGOSDOS0_Init(oldYGOSDOS0)

mGOSDOSISRV = Trim(Mid$(cboSelect_4_GOSDOSISRV, 1, 3))
mGOSEVEUSRV = Trim(Mid$(cboSelect_4_GOSDOSGSRV, 1, 3))

Do While Not rsSab.EOF
    blnOk = False
    If mGOSDOSISRV = "" Or mGOSDOSISRV = rsSab("GOSDOSISRV") Then blnOk = True
    If mGOSEVEUSRV = "" Or mGOSEVEUSRV = rsSab("GOSEVEUSRV") Then blnOk = True
    
    If blnOk Then
        V = rsYGOSDOS0_GetBuffer(rsSab, xYGOSDOS0)
        V = rsYGOSEVE0_GetBuffer(rsSab, xYGOSEVE0)
        
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine_4Journal
    End If
    rsSab.MoveNext
Loop
    

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYGOSDOS0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgSelect_Display_YGOSEVE0()
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<MàJ le                                      |<par le service                            |<pour le service     |<Nature   |<Evénement                                                                                                           |SWISABSWID       |"
      
fgSelect.Row = 0

currentAction = "fgSelect_Display"
    
Do While Not rsSab.EOF
    Call rsYGOSEVE0_GetBuffer(rsSab, xYGOSEVE0)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine_YGOSEVE0 I

    rsSab.MoveNext

Loop
    

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_DisplayLine_YGOSEVE0(lIndex As Long)
Dim K As Integer, wColor As Long
Dim HeightOfLine As Long, LinesOfText As Long


On Error Resume Next
Select Case xYGOSEVE0.GOSEVESTAE
    Case " ":
        If cmdSelect_SQL_K = "4 Swi>" Then
            wColor = RGB(128, 0, 0)
        Else
            wColor = vbBlue
        End If
    Case "A": wColor = &H606060
    Case Else: wColor = vbMagenta
End Select
fgSelect.Col = 0: fgSelect.Text = xYGOSEVE0.GOSEVEIDD & " - " & xYGOSEVE0.GOSEVEIDE & " : " & dateImp10_S(xYGOSEVE0.GOSEVEUAMJ) & "  " & timeImp8(xYGOSEVE0.GOSEVEUHMS)
fgSelect.CellForeColor = wColor
fgSelect.Col = 1
K = Val(Mid$(xYGOSEVE0.GOSEVEUSRV, 2, 2)): fgSelect.Text = arrService_Lib(K) & "/" & xYGOSEVE0.GOSEVEUUSR
fgSelect.CellForeColor = wColor
fgSelect.Col = 2
K = Val(Mid$(xYGOSEVE0.GOSEVEGSRV, 2, 2)): fgSelect.Text = arrService_Lib(K)
fgSelect.CellForeColor = wColor
fgSelect.Col = 3: fgSelect.Text = xYGOSEVE0.GOSEVENAT
fgSelect.CellForeColor = wColor
fgSelect.Col = 4
fgSelect.Text = Trim(xYGOSEVE0.GOSEVETXT)
txtFg = fgSelect.Text
 HeightOfLine = fgSelect.RowHeightMin / 3 + 50 '- 20 'Me.TextHeight(txteve.Text)

 LinesOfText = SendMessage(txtFg.hwnd, EM_GETLINECOUNT, 0&, 0&) + 1
 
 If fgSelect.RowHeight(fgSelect.Row) < (LinesOfText * HeightOfLine) Then
    fgSelect.RowHeight(fgSelect.Row) = LinesOfText * HeightOfLine
 End If
fgSelect.CellForeColor = wColor
 fgSelect.Col = 5
fgSelect.Text = xYGOSEVE0.GOSEVESWID

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

End Sub

Private Sub fglist_Display_YGOSDOS0()
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgList.Visible = False
fgList_Reset

fgList.Rows = 1
fgList.FormatString = "> Dossier|<Service   |<BIC émis/reçu          |<référence                          |<Type    |>Montant               |<Dev    |<Etat |<Echéance     |<RCOM            |<Client        |<Motif du suspens         |"
fgList.BackColorFixed = fgSelect.BackColorFixed

fgList.Row = 0
mList_Row = 0
currentAction = "fglist_Display"
    
For I = 1 To arrYGOSDOS0_Nb
    xYGOSDOS0 = arrYGOSDOS0(I)
    fgList.Rows = fgList.Rows + 1
    fgList.Row = fgList.Rows - 1
    fglist_DisplayLine_YGOSDOS0 I

Next I
    

fgList.Visible = True
If mList_Row > 0 Then fgList.TopRow = mList_Row

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYGOSDOS0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fglist_Display_YSWISAB0(lWhere As String)
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgList.Visible = False
fgList_Reset

fgList.Rows = 1
fgList.FormatString = "<Type    |<L/référence             |<SAB              |>Montant               |<Dev    |date réception                  |||"
fgList.BackColorFixed = fgSelect.BackColorFixed

fgList.Row = 0
mList_Row = 0
currentAction = "fglist_Display"

 
Set rsSab = cnsab.Execute("select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " & lWhere)

Do While Not rsSab.EOF
    Call rsYSWISAB0_GetBuffer(rsSab, xYSWISAB0)

    fgList.Rows = fgList.Rows + 1
    fgList.Row = fgList.Rows - 1
    fglist_DisplayLine_YSWISAB0

    rsSab.MoveNext

Loop
         
    

fgList.Visible = True
If mList_Row > 0 Then fgList.TopRow = mList_Row

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYGOSDOS0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgList_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgList.Visible = False: fraDetail.Visible = False
mRow = fgList.Row

If lRow > 0 And lRow < fgList.Rows Then
    fgList.Row = lRow
    For I = fglist_arrIndex To fgList.FixedCols Step -1
        fgList.Col = I: fgList.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgList.Row = mRow
    If fgList.Row > 0 Then
        lRow = fgList.Row
        lColor_Old = fgList.CellBackColor
        For I = fglist_arrIndex To fgList.FixedCols Step -1
          fgList.Col = I: fgList.CellBackColor = lColor
        Next I
    End If
End If
fgList.LeftCol = fgList.FixedCols
fgList.Visible = True: fraDetail.Visible = True
End Sub
Public Sub fgList_Reset()
fgList.Clear
fglist_Sort1 = 0: fglist_Sort2 = 0
fglist_Sort1_Old = -1
fglist_RowDisplay = 0: fglist_RowClick = 0
fglist_arrIndex = fgList.Cols - 1
blnfglist_DisplayLine = False
fglist_SortAD = 6
fgList.LeftCol = fgList.FixedCols

End Sub




Public Sub fgList_Sort()
If fgList.Rows > 1 Then
    fgList.Row = 1
    fgList.RowSel = fgList.Rows - 1
    
    If fglist_Sort1_Old = fglist_Sort1 Then
        If fglist_SortAD = 5 Then
            fglist_SortAD = 6
        Else
            fglist_SortAD = 5
        End If
    Else
        fglist_SortAD = 5
    End If
    fglist_Sort1_Old = fglist_Sort1
    
    fgList.Col = fglist_Sort1
    fgList.ColSel = fglist_Sort2
    fgList.Sort = fglist_SortAD
End If

End Sub




Private Sub arrYGOSDOS0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYGOSDOS0(101)
arrYGOSDOS0_Max = 100: arrYGOSDOS0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSDOS0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYGOSDOS0_GetBuffer(rsSab, xYGOSDOS0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYGOSDOS0.fgselect_Display"
        '' Exit Sub
     Else
         arrYGOSDOS0_Nb = arrYGOSDOS0_Nb + 1
         If arrYGOSDOS0_Nb > arrYGOSDOS0_Max Then
             arrYGOSDOS0_Max = arrYGOSDOS0_Max + 100
             ReDim Preserve arrYGOSDOS0(arrYGOSDOS0_Max)
         End If
         
         arrYGOSDOS0(arrYGOSDOS0_Nb) = xYGOSDOS0
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

Private Sub arrYGOSEVE0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYGOSEVE0(101)
arrYGOSEVE0_Max = 100: arrYGOSEVE0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYGOSEVE0_GetBuffer(rsSab, xYGOSEVE0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYGOSEVE0.fgselect_Display"
        '' Exit Sub
     Else
         arrYGOSEVE0_Nb = arrYGOSEVE0_Nb + 1
         If arrYGOSEVE0_Nb > arrYGOSEVE0_Max Then
             arrYGOSEVE0_Max = arrYGOSEVE0_Max + 100
             ReDim Preserve arrYGOSEVE0(arrYGOSEVE0_Max)
         End If
         
         arrYGOSEVE0(arrYGOSEVE0_Nb) = xYGOSEVE0
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
    If frmYGOSDOS0_Param.Visible Then frmYGOSDOS0_Param.Hide

    K = InStr(cboSelect_SQL, "-")
    If K > 1 Then
        cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, K - 1))
    Else
        cmdSelect_SQL_K = "???"
    End If
    
    fraSelect_Options.Visible = True: fraSelect_Options_1.Visible = False: fraSelect_Options_3.Visible = False
    fraSelect_Options_9.Visible = False
    fraSelect_Options_4.Visible = False
    fraSelect_Options_7.Visible = False
    fraSelect_Options_Stat.Visible = False
    Select Case cmdSelect_SQL_K
        Case "1", "6", "6#", "7", "7J", "7#":
            chkSelect_GOSDOSIAMJ.Enabled = False: chkSelect_GOSDOSIAMJ = "1"
            fraSelect_Options_1.Visible = True
            fgSelect.BackColorFixed = &HA0A000
            cboSelect_GOSDOSKSRV.Visible = False
            chkSelect_GOSDOSKSRV.Visible = False
            fraSelect_Options_1b.Visible = False
            'txtSelect_rTextField.Visible = True: lblSelect_rTextField.Visible = True
            'cmdSelect_Ok_Click
            cmdSelect_Ok.Visible = True
        Case "1b", "1?", "1?*", "1trf":
            chkSelect_GOSDOSIAMJ.Enabled = False: chkSelect_GOSDOSIAMJ = "1"
            fraSelect_Options_1.Visible = True
            fgSelect.BackColorFixed = &HA0A000
            cboSelect_GOSDOSKSRV.Visible = True
            chkSelect_GOSDOSKSRV.Visible = True
            cboSelect_SWISABWSTA.Visible = True
            fraSelect_Options_1b.Visible = True
            'txtSelect_rTextField.Visible = False: lblSelect_rTextField.Visible = False
            cmdSelect_Ok.Visible = True
        Case "2":
            chkSelect_GOSDOSIAMJ.Enabled = True: chkSelect_GOSDOSIAMJ = "1"
            fraSelect_Options_1.Visible = True
            fgSelect.BackColorFixed = &HFF80FF    '&HFF00FF
            cboSelect_GOSDOSWMTK.ListIndex = 3
            fraSelect_Options_1b.Visible = True
            'txtSelect_rTextField.Visible = False: lblSelect_rTextField.Visible = False
            cboSelect_GOSDOSKSRV.Visible = False
            chkSelect_GOSDOSKSRV.Visible = False
            cboSelect_SWISABWSTA.Visible = False
            cboSelect_GOSDOSWES.ListIndex = 1
            cmdSelect_Ok_Click
            cmdSelect_Ok.Visible = True
         Case "3", "3x":
            chkSelect_GOSDOSIAMJ.Enabled = True: chkSelect_GOSDOSIAMJ = "0"
            fraSelect_Options_3.Visible = True
            fgSelect.BackColorFixed = &H808000
            fraDetail_C.Visible = True
            cmdSelect_Ok_Click
            cmdSelect_Ok.Visible = True
         Case "4":
            'chkSelect_GOSDOSIAMJ.Enabled = False
            Call cmdSelect_Reset_4("4")
            cboSelect_4_GOSDOSISRV.ListIndex = 0
            fraSelect_Options_4.Visible = True
            fgSelect.BackColorFixed = &H808000
            fraDetail_C.Visible = True
            cmdSelect_Ok_Click
            cmdSelect_Ok.Visible = True
         Case "4 Journal":
            'chkSelect_GOSDOSIAMJ.Enabled = False
            Call cmdSelect_Reset_4("4 Journal")
            Call cbo_Scan(currentSSIWINUNIT, cboSelect_4_GOSDOSISRV)
            fraSelect_Options_4.Visible = True
            fgSelect.BackColorFixed = &H808000
            fraDetail_C.Visible = True
            cmdSelect_Ok_Click
            cmdSelect_Ok.Visible = True
         Case "4 Swi>":
            'chkSelect_GOSDOSIAMJ.Enabled = False
            Call cmdSelect_Reset_4("4 Swi>")
            Call cbo_Scan(currentSSIWINUNIT, cboSelect_4_GOSDOSGSRV)
            fraSelect_Options_4.Visible = True
            fgSelect.BackColorFixed = &H808000
            fraDetail_C.Visible = True
            cmdSelect_Ok_Click
            cmdSelect_Ok.Visible = True
          Case "5":
            blnSwift_Display = True
            chkSelect_GOSDOSIAMJ.Enabled = True: chkSelect_GOSDOSIAMJ = "0"
            fraSelect_Options_9.Visible = True
            fgSelect.BackColorFixed = &HFF9090
            cboSelect_GOSDOSWMTK.ListIndex = 1
            cmdSelect_Ok_Click
            cmdSelect_Ok.Visible = True
           Case "5h":
            blnSwift_Display = True
            chkSelect_GOSDOSIAMJ.Enabled = True: chkSelect_GOSDOSIAMJ = "1"
            fraSelect_Options_9.Visible = False
            fgSelect.BackColorFixed = &HFF9090
            cboSelect_GOSDOSWMTK.ListIndex = 1
            cmdSelect_Ok_Click
            cmdSelect_Ok.Visible = True
         Case "7E", "7E*":
            
            fraSelect_Options_7.Visible = True
            fgSelect.BackColorFixed = &H808000
            fraDetail_C.Visible = True
            cmdSelect_Ok_Click
            cmdSelect_Ok.Visible = True
       Case "9":
            chkSelect_GOSDOSIAMJ.Enabled = True: chkSelect_GOSDOSIAMJ = "0"
            fraSelect_Options_9.Visible = True
            fgSelect.BackColorFixed = &H80FF&
            cboSelect_GOSDOSWMTK.ListIndex = 1
            cmdSelect_Ok_Click
            cmdSelect_Ok.Visible = True
         Case "9+":
            chkSelect_GOSDOSIAMJ.Enabled = True: chkSelect_GOSDOSIAMJ = "0"
            fraSelect_Options_1.Visible = True
            fgSelect.BackColorFixed = &H80FF&
            cboSelect_GOSDOSWMTK.ListIndex = 1
            'cmdSelect_Ok_Click
            cmdSelect_Ok.Visible = True
       Case "Stat", "Stat BIC":
            fraSelect_Options_Stat.Visible = True
      Case Else
           chkSelect_GOSDOSIAMJ.Enabled = False: chkSelect_GOSDOSIAMJ = "0"
           fraSelect_Options.Visible = False
           cmdSelect_Ok.Visible = True
    End Select

End If
End Sub

Public Sub cmdDetail_Reset()
If blnControl Then
    lstErr.Clear
    If fgDetail.Visible Then
        fgDetail.Visible = False: fraDetail.Visible = False
        fraMail_MT.Visible = False: fraEVE.Visible = False
        fgDetail_Display
    End If
End If

End Sub


Private Sub cmdSelect_SQL_3()
Dim V
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_3("
blnOk = False

If Val(txtSelect_3_GOSDOSIDD) > 0 Then
    xWhere = "where GOSDOSIDD = " & Val(txtSelect_3_GOSDOSIDD)
Else
    Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Min, wAmjMin)
    Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Max, wAmjMax)
    
    xWhere = ""
    X = Trim(Mid$(cboSelect_3_GOSDOSGSRV, 1, 3))
    If X <> "" Then xWhere = xWhere & " and ( GOSDOSISRV = '" & X & "' or GOSDOSGSRV = '" & X & "')"
    
    X = Trim(Mid$(cboSelect_3_GOSDOSRCOM, 1, 3))
    If X <> "" Then xWhere = xWhere & " and GOSDOSRCOM = '" & X & "'"
    
    X = Trim(Mid$(cboSelect_3_GOSDOSWBIC, 1, 11))
    If X <> "" Then xWhere = xWhere & " and GOSDOSWBIC like '" & X & "%'"
    
    X = Trim(cboSelect_3_GOSDOSCLI)
    If X <> "" Then xWhere = xWhere & " and GOSDOSCLI = '" & Format$(Val(X), "0000000") & "'"
    
    X = Trim(Mid$(cboSelect_3_GOSDOSSTAD, 1, 1))
    If X <> "*" Then xWhere = xWhere & " and GOSDOSSTAD = '" & X & "'"
    
    X = Trim(Mid$(cboSelect_3_GOSDOSSTAG, 1, 1))
    If X <> "" Then xWhere = xWhere & " and GOSDOSSTAG = '" & X & "'"
    
    If chkSelect_GOSDOSIAMJ = "1" Then xWhere = xWhere & " and GOSDOSIAMJ >= " & wAmjMin & " and GOSDOSIAMJ <= " & wAmjMax
    
    xWhere = Replace(xWhere, " and", " where", 1, 1)
End If

Call arrYGOSDOS0_SQL(xWhere)
  

fgSelect_Display_YGOSDOS0

If fgSelect.Rows = 2 Then fgSelect.Row = 1: Call fgSelect_MouseDown(0, 0, 100, fgSelect.RowHeightMin + 100)
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub cmdSelect_SQL_4()
Dim V
Dim xSql As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_4"
blnOk = False


Call DTPicker_Control(txtSelect_4_GOSDOSECHD, wAmjMax)

xWhere = " where  GOSDOSSTAD = ' ' and GOSDOSECHD <= " & wAmjMax
X = Trim(Mid$(cboSelect_4_GOSDOSISRV, 1, 3))
If X <> "" Then xWhere = xWhere & " and GOSDOSISRV = '" & X & "'"
X = Trim(Mid$(cboSelect_4_GOSDOSGSRV, 1, 3))
If X <> "" Then xWhere = xWhere & " and GOSDOSGSRV = '" & X & "'"

' Call arrYGOSDOS0_SQL(xWhere & " order by GOSDOSECHD")
  
'fgSelect_Display_YGOSDOS0

xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0 ," & paramIBM_Library_SABSPE & ".YGOSDOS0" _
     & xWhere _
     & " and  GOSEVEIDD > 0 and GOSEVEIDD = GOSDOSIDD" _
     & " order by GOSEVEIDD , GOSEVEIDE"

Set rsSab = cnsab.Execute(xSql)

  
fgSelect_Display_Echéancier


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub Auto_BIA_GOS()
Dim V, X As String
Dim xSql As String
Dim rsADO As New ADODB.Recordset
On Error GoTo Error_Handler
0
currentAction = "Auto_BIA_GOS | Echéancier"
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents


Call DTPicker_Set(txtSelect_4_GOSDOSECHD, DSys) '
Call cmdSelect_Reset_4("4")

'=================================================================================================
xSql = "select distinct GOSDOSGSRV from " & paramIBM_Library_SABSPE & ".YGOSDOS0" _
      & " where  GOSDOSSTAD = ' ' and GOSDOSECHD <= " & DSys _
      & " order by GOSDOSGSRV"

Set rsADO = cnsab.Execute(xSql)
cboSelect_4_GOSDOSISRV.ListIndex = 0
Do While Not rsADO.EOF
    X = rsADO(0)
    Call cbo_Scan(X, cboSelect_4_GOSDOSGSRV)
    xSql = " where  GOSDOSSTAD = ' ' and GOSDOSECHD <= " & DSys & " and GOSDOSGSRV = '" & X & "' order by GOSDOSECHD"
    Call arrYGOSDOS0_SQL(xSql)
    If arrYGOSDOS0_Nb > 0 Then
        cmdSelect_SQL_K = "4"
        fgSelect_Display_YGOSDOS0
        X = mailAdresse_Recipient_Unit(X, 2)
        Call cmdSendMail_fgSelect(X)
    End If
    rsADO.MoveNext
Loop

'=================================================================================================
xSql = "select distinct GOSDOSISRV from " & paramIBM_Library_SABSPE & ".YGOSDOS0" _
      & " where  GOSDOSSTAD = ' ' and GOSDOSECHD <= " & DSys _
      & " order by GOSDOSISRV"

Set rsADO = cnsab.Execute(xSql)
cboSelect_4_GOSDOSGSRV.ListIndex = 0

Do While Not rsADO.EOF
    X = rsADO(0)
    Call cbo_Scan(X, cboSelect_4_GOSDOSISRV)
    xSql = " where  GOSDOSSTAD = ' ' and GOSDOSECHD <= " & DSys & " and GOSDOSISRV = '" & X & "' order by GOSDOSECHD"
    Call arrYGOSDOS0_SQL(xSql)
    If arrYGOSDOS0_Nb > 0 Then
        cmdSelect_SQL_K = "4"
        fgSelect_Display_YGOSDOS0
        X = mailAdresse_Recipient_Unit(X, 2)
        Call cmdSendMail_fgSelect(X)
    End If
    rsADO.MoveNext
Loop


'=================================================================================================
Journal:

Call cmdSelect_Reset_4("4 Journal")

xSql = "select distinct GOSEVEUSRV from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
      & " where GOSEVEIDD > 0 order by GOSEVEUSRV"

Set rsADO = cnsab.Execute(xSql)
Call DTPicker_Set(txtSelect_4_GOSDOSECHD, YBIATAB0_DATE_CPT_J)

Do While Not rsADO.EOF
    X = rsADO(0)
    Call cbo_Scan(X, cboSelect_4_GOSDOSGSRV)
    Call cbo_Scan(X, cboSelect_4_GOSDOSISRV)
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0 ," & paramIBM_Library_SABSPE & ".YGOSDOS0" _
         & " where  GOSEVEIDD > 0 and GOSEVEIDD = GOSDOSIDD and  GOSEVEUAMJ = " & YBIATAB0_DATE_CPT_J _
         & " order by GOSEVEIDD , GOSEVEIDE"

    Set rsSab = cnsab.Execute(xSql)
    cmdSelect_SQL_K = "4 Journal"
    fgSelect_Display_4Journal
    If fgSelect.Rows > 1 Then
        X = mailAdresse_Recipient_Unit(X, 2)
        Call cmdSendMail_fgSelect(X)
    End If
    rsADO.MoveNext
Loop


'=================================================================================================
Swi:

Call cmdSelect_Reset_4("4 Swi>")

xSql = "select distinct GOSEVEUSRV from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where  GOSEVEIDD > 0 and GOSEVENAT = 'Swi>' and  GOSEVESWID =  0 and GOSEVESTAE = ''"

Set rsADO = cnsab.Execute(xSql)

Do While Not rsADO.EOF
    X = rsADO(0)
    Call cbo_Scan(X, cboSelect_4_GOSDOSGSRV)
    cmdSelect_SQL_K = "4 Swi>"
    cmdSelect_SQL_4Swi
    If fgSelect.Rows > 1 Then
        X = mailAdresse_Recipient_Unit(X, 2)
        Call cmdSendMail_fgSelect(X)
    End If
    rsADO.MoveNext
Loop


'=================================================================================================


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_4Journal()
Dim V
Dim xSql As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_4j"


Call DTPicker_Control(txtSelect_4_GOSDOSECHD, wAmjMax)

xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0 ," & paramIBM_Library_SABSPE & ".YGOSDOS0" _
     & " where  GOSEVEIDD > 0 and GOSEVEIDD = GOSDOSIDD and  GOSEVEUAMJ = " & wAmjMax _
     & " order by GOSEVEIDD , GOSEVEIDE"

Set rsSab = cnsab.Execute(xSql)

  
fgSelect_Display_4Journal

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_4Swi()
Dim V, X As String, xWhere As String
Dim xSql As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_4Swi"

xWhere = ""
X = Trim(Mid$(cboSelect_4_GOSDOSGSRV, 1, 3))
If X <> "" Then xWhere = " and GOSEVEUSRV = '" & X & "'"

xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where  GOSEVEIDD > 0 and GOSEVENAT = 'Swi>' and  GOSEVESWID =  0 and GOSEVESTAE = ''" _
     & xWhere _
     & " order by GOSEVEIDD , GOSEVEIDE"

Set rsSab = cnsab.Execute(xSql)

  
fgSelect_Display_YGOSEVE0

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_9M()
Dim V, X As String
Dim xSql As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_9"

Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Min, wAmjMin)
Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Max, wAmjMax)

xWhere = " where SWISABWAMJ >= " & wAmjMin & " and SWISABWAMJ <= " & wAmjMax


X = Trim(cboSelect_GOSDOSWMTK)
If X <> "" Then
    If InStr(X, "%") Then
        xWhere = xWhere & " and SWISABWMTK like '" & X & "'"
    Else
        If InStr(X, ",") Then
            xWhere = xWhere & " and SWISABWMTK in ('" & Replace(X, ",", "','") & "')"
        Else
            xWhere = xWhere & " and SWISABWMTK = '" & X & "'"
        End If
    End If
End If

X = Trim(cboSelect_GOSDOSWES)
Select Case X
    Case "Entrant": xWhere = xWhere & " and SWISABWEs = 'E'"
    Case "Sortant": xWhere = xWhere & " and SWISABWEs = 'S'"
End Select

X = Trim(cboSelect_GOSDOSWDEV)
If X <> "" Then xWhere = xWhere & " and SWISABWDEV = '" & X & "'"

X = Trim(cboSelect_GOSDOSWBIC)
If X <> "" Then xWhere = xWhere & " and SWISABWBIC like '" & X & "%'"

X = Trim(txtSelect_GOSDOSWTRN)
If X <> "" Then xWhere = xWhere & " and ( SWISABWL20 like '%" & X & "%' or SWISABWN20 like '%" & X & "%')"

X = Trim(Mid$(cboSelect_GOSDOSKSRV, 1, 3))
If X <> "" Then
    If X = "S00" Then
        xWhere = xWhere & " and SWISABKSRV = '" & X & "'"
    Else
        If chkSelect_GOSDOSKSRV = "1" Then
            xWhere = xWhere & " and SWISABKSRV in ('S00' , '" & X & "')"
        Else
            xWhere = xWhere & " and SWISABKSRV = '" & X & "'"
       End If
        
    End If
End If

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " & xWhere _
     & " order by SWISABWBIC,SWISABWAMJ,SWISABWHMS,SWISABWMTK"
Set rsSab = cnsab.Execute(xSql)

fgSelect_Display_YSWISAB0

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_9()
Dim V, X As String
Dim xSql As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_9"

If Mid$(cboSelect_9_SWISABKSTA, 1, 1) = "!" Then
    xWhere = "where SWISABK999 = '!'"
Else
    xWhere = "where SWISABK999 <> ' '"
    If chkSelect_GOSDOSIAMJ <> "1" Then
        Call MsgBox("Préciser la période de recherche de l'historique", vbExclamation, "Gestion des messages *99")
        Exit Sub
    End If
End If

Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Min, wAmjMin)
Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Max, wAmjMax)
If chkSelect_GOSDOSIAMJ = "1" Then xWhere = xWhere & " and SWISABWAMJ >= " & wAmjMin & " and SWISABWAMJ <= " & wAmjMax

X = Trim(Mid$(cboSelect_9_SWISABKSRV, 1, 3))
If X <> "" Then
    If X = "S00" Then
        xWhere = xWhere & " and SWISABKSRV = '" & X & "'"
    Else
        If chkSelect_9_SWISABKSRV = "1" Then
            xWhere = xWhere & " and SWISABKSRV in ('S00' , '" & X & "')"
        Else
            xWhere = xWhere & " and SWISABKSRV = '" & X & "'"
       End If
        
    End If
End If


xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " & xWhere _
     & " order by SWISABWBIC,SWISABWAMJ,SWISABWHMS,SWISABWMTK"
Set rsSab = cnsab.Execute(xSql)

fgSelect_Display_YSWISAB0

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_6()
Dim V, X As String
Dim xSql As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_6"
mYSWIRAM0_Col = 0

If cmdSelect_SQL_K = "6" Or cmdSelect_SQL_K = "6#" Then
    X = Trim(txtSelect_SWISABOPEN)
    If X <> "" Then
        xWhere = " and SWIRAMXOPE like '%" & X & "'"
    Else
        Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Min, wAmjMin)
        Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Max, wAmjMax)
        
        xWhere = ""
        If chkSelect_GOSDOSIAMJ = "1" Then xWhere = xWhere & " and SWIRAMYAMJ >= " & wAmjMin & " and SWIRAMYAMJ <= " & wAmjMax
        X = Trim(cboSelect_GOSDOSWBIC)
        If X <> "" Then xWhere = xWhere & " and SWIRAMXBIC like '" & X & "%'"
        
         X = Trim(cboSelect_GOSDOSWMTK)
         If X <> "" Then
            If InStr(X, "%") > 0 Then
                xWhere = xWhere & " and SWIRAMXMTK like '" & X & "'"
            Else
                If InStr(X, ",") Then
                    xWhere = xWhere & " and SWIRAMXMTK in ('" & Replace(X, ",", "','") & "')"
                Else
                    xWhere = xWhere & " and SWIRAMXMTK = '" & X & "'"
                End If
            End If
        End If
        
        X = Trim(cboSelect_GOSDOSWES)
        Select Case X
            Case "Sortant": xWhere = xWhere & " and SWIRAMXES = 'S'"
            Case "Entrant": xWhere = xWhere & " and SWIRAMXES = 'E'"
        End Select
    End If
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIRAM0, " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
         & " where swisabswid = SWIRAMXID " & xWhere _
         & " order by SWIRAMXOPE,SWIRAMXBIC,SWIRAMXMTK,SWIRAMXREF,SWIRAMXID"
Else
    xWhere = " where SWIRAMSTA in ('#' , '?') and swisabswid = SWIRAMXID"
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIRAM0S , " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
         & xWhere _
         & " order by SWIRAMXBIC,SWIRAMXID,SWIRAMXREF,SWISABWMTD,SWIRAMXES"
End If

Set rsSab = cnsab.Execute(xSql)

fgSelect_Display_YSWISAB0

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub cmdSelect_SQL_7()
Dim V, X As String
Dim xSql As String, xWhere As String, xOPEC As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_7"

X = Trim(cboSelect_SWISABOPEC)
If X <> "" Then
    xOPEC = " and SWIECHOPEC = '" & X & "'"
Else
    xOPEC = ""
End If


Select Case cmdSelect_SQL_K
    Case "7", "7#"
        X = Trim(txtSelect_SWISABOPEN)
        If X <> "" Then
            xWhere = " and SWIECHOPEN = " & X
            'X = Trim(cboSelect_SWISABOPEC)
            'If X <> "" Then xWhere = xWhere & " and SWIECHOPEC = '" & X & "'"
        Else
            Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Min, wAmjMin)
            Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Max, wAmjMax)
            
            xWhere = ""
             If chkSelect_GOSDOSIAMJ = "1" Then xWhere = xWhere & " and SWIECHDECH >= " & wAmjMin & " and SWIECHDECH <= " & wAmjMax
            X = Trim(cboSelect_GOSDOSWBIC)
            If X <> "" Then xWhere = xWhere & " and SWIECHWBIC like '" & X & "%'"
            
             X = Trim(cboSelect_GOSDOSWMTK)
             If X <> "" Then
                If InStr(X, "%") > 0 Then
                    xWhere = xWhere & " and SWIECHWMTK like '" & X & "'"
                Else
                    If InStr(X, ",") Then
                        xWhere = xWhere & " and SWIECHWMTK in ('" & Replace(X, ",", "','") & "')"
                    Else
                        xWhere = xWhere & " and SWIECHWMTK = '" & X & "'"
                    End If
                End If
            End If
            
            X = Trim(cboSelect_GOSDOSWES)
            Select Case X
                Case "Sortant": xWhere = xWhere & " and SWIECHWES = 'S'"
                Case "Entrant": xWhere = xWhere & " and SWIECHWES = 'E'"
            End Select
        End If
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIECH0" _
             & " where SWIECHSWID > 0 " & xOPEC & xWhere _
             & " order by SWIECHOPEC,SWIECHOPEN,SWIECHWBIC,SWIECHSWID,SWIECHDECH,SWIECHWMTK,SWIECHSWIX"

'    Case "7E"
'        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIECH0E " _
'             & " where SWIECHDECH <= " & DSys & xOPEC _
'             & " order by SWIECHDECH,SWIECHWES desc,SWIECHOPEC,SWIECHOPEN,SWIECHSWID"
'    Case "7E*"
'        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIECH0E " _
'             & " where SWIECHDECH > 0 " & xOPEC _
'             & " order by SWIECHDECH,SWIECHWES desc,SWIECHOPEC,SWIECHOPEN,SWIECHSWID"
End Select

Set rsSab = cnsab.Execute(xSql)

fgSelect_Display_YSWIECH0

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_7E()
Dim V, X As String
Dim xSql As String, xWhere As String, xOPEC As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_7E"

X = Trim(cboSelect_7_SRV)
If X = "" Then
    xOPEC = ""
Else
    Select Case Mid$(X, 1, 3)
        Case "S01"
            xOPEC = " and SWIECHSER = '00' and SWIECHOPEC not in ( 'CDE' , 'CDI')"
         Case "S10"
            xOPEC = " and SWIECHSER = '00' and SWIECHOPEC in ( 'CDE' , 'CDI')"
        Case "S32"
            xOPEC = " and SWIECHSER = 'TC'"
           
    End Select
End If

Select Case cmdSelect_SQL_K
    Case "7E"
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIECH0E " _
             & " where SWIECHDECH <= " & DSys & xOPEC _
             & " order by SWIECHDECH,SWIECHWES desc,SWIECHOPEC,SWIECHOPEN,SWIECHSWID"
    Case "7E*"
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIECH0E " _
             & " where SWIECHDECH > 0 " & xOPEC _
             & " order by SWIECHDECH,SWIECHWES desc,SWIECHOPEC,SWIECHOPEN,SWIECHSWID"
End Select

Set rsSab = cnsab.Execute(xSql)

fgSelect_Display_YSWIECH0

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_7J()
Dim V, X As String
Dim xSql As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_7J"

    X = Trim(txtSelect_SWISABOPEN)
    If X <> "" Then
        xWhere = " and SWISABOPEN = " & X
        X = Trim(cboSelect_SWISABOPEC)
        If X <> "" Then xWhere = xWhere & " and SWISABOPEC = '" & X & "'"
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIECH1 A " _
             & " left outer join " & paramIBM_Library_SABSPE & ".YSWIECH0 B on  ( A.SWIEC1SWID = B.SWIECHSWID and A.SWIEC1SEQ0 = B.SWIECHSEQ0 ) ," _
             & paramIBM_Library_SABSPE & ".YSWISAB0 " & " where A.SWIEC1SWID = SWISABSWID " _
             & xWhere & " order by SWISABOPEC,SWISABOPEN,SWISABWBIC,SWISABSWID"
    Else
        Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Min, wAmjMin)
        Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Max, wAmjMax)
        
        xWhere = ""
        If chkSelect_GOSDOSIAMJ = "1" Then xWhere = xWhere & " and A.SWIEC1YAMJ >= " & wAmjMin & " and A.SWIEC1YAMJ <= " & wAmjMax
        X = Trim(cboSelect_GOSDOSWBIC)
        If X <> "" Then xWhere = xWhere & " and SWISABWBIC like '" & X & "%'"
        
         X = Trim(cboSelect_GOSDOSWMTK)
         If X <> "" Then
            If InStr(X, "%") > 0 Then
                xWhere = xWhere & " and SWISABWMTK like '" & X & "'"
            Else
                If InStr(X, ",") Then
                    xWhere = xWhere & " and SWISABWMTK in ('" & Replace(X, ",", "','") & "')"
                Else
                    xWhere = xWhere & " and SWISABWMTK = '" & X & "'"
                End If
            End If
        End If
        
        X = Trim(cboSelect_GOSDOSWES)
        Select Case X
            Case "Sortant": xWhere = xWhere & " and SWISABWES = 'S'"
            Case "Entrant": xWhere = xWhere & " and SWISABWES = 'E'"
        End Select
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIECH1 A " _
         & " left outer join " & paramIBM_Library_SABSPE & ".YSWIECH0 B on  ( A.SWIEC1SWID = B.SWIECHSWID and A.SWIEC1SEQ0 = B.SWIECHSEQ0 ) ," _
         & paramIBM_Library_SABSPE & ".YSWISAB0 " & " where A.SWIEC1SWID = SWISABSWID " _
         & xWhere & " order by SWISABOPEC,SWISABOPEN,SWISABWBIC,SWISABSWID"
End If
'xSQL = "select * from " & paramIBM_Library_SABSPE & ".YTVAFAC0 inner join " & paramIBM_Library_SAB & ".ZCLIENA0 on  CLIENACLI = TVAFACCLI" & xWhere & "order by CLIENARES,CLIENACLI"

Set rsSab = cnsab.Execute(xSql)

fgSelect_Display_YSWIECH1

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub cmdSelect_SQL_5()
Dim V, X As String
Dim xSql As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_5"

Select Case Mid$(cboSelect_9_SWISABKSTA, 1, 1)
    Case "!"
        xWhere = "where (SWISABKPDE in ('!','?') or SWISABK20 = '!')"
     Case " "
        xWhere = "where SWISABWES = 'E' "
   Case Else
        xWhere = "where (SWISABKPDE <> ' ' or SWISABK20 <> ' ')"
End Select

Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Min, wAmjMin)
Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Max, wAmjMax)
If chkSelect_GOSDOSIAMJ = "1" Then xWhere = xWhere & " and SWISABWAMJ >= " & wAmjMin & " and SWISABWAMJ <= " & wAmjMax

X = Trim(Mid$(cboSelect_9_SWISABKSRV, 1, 3))
If X <> "" Then
    If X = "S00" Then
        xWhere = xWhere & " and SWISABKSRV = '" & X & "'"
    Else
        If chkSelect_9_SWISABKSRV = "1" Then
            xWhere = xWhere & " and SWISABKSRV in ('S00' , '" & X & "')"
        Else
            xWhere = xWhere & " and SWISABKSRV = '" & X & "'"
       End If
        
    End If
End If

If Mid$(cboSelect_9_SWISABKSTA, 1, 1) = " " And chkSelect_GOSDOSIAMJ = "0" Then
    Call MsgBox("Précisez la période SVP", vbInformation, "option 5 - modifiée 11-12-2012")
Else
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " & xWhere _
         & " order by SWISABWMTK,SWISABWBIC,SWISABWAMJ,SWISABWHMS"
    Set rsSab = cnsab.Execute(xSql)
    
    fgSelect_Display_YSWISAB0
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub Importation_SAB_SWISABKPDE()
Dim V, X As String
Dim xSql As String, xWhere As String
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass

currentAction = "Importation_SAB_SWISABKPDE"
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

xWhere = "where SWISABOPEN <> 0 and (SWISABKPDE = '!' or SWISABK20 = '!')"

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " & xWhere _
     & " order by SWISABSWID"
Set rsSab = cnsab.Execute(xSql)


Do While Not rsSab.EOF
    V = rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)

    newYSWISAB0 = oldYSWISAB0
    If oldYSWISAB0.SWISABK20 <> " " Then newYSWISAB0.SWISABK20 = "@"
    If oldYSWISAB0.SWISABKPDE <> " " Then newYSWISAB0.SWISABKPDE = "@"
    '¤JPL 20120705 If oldYSWISAB0.SWISABK999 <> " " Then newYSWISAB0.SWISABK999 = "@"
    
    V = sqlYSWISAB0_Update(newYSWISAB0, oldYSWISAB0)
    If Not IsNull(V) Then GoTo Error_MsgBox
    
    rsSab.MoveNext

Loop
'=======================================================================================
xWhere = "where (SWISABKPDE = '!' or SWISABK999 = '!') and SWISABWAMJ < " & YBIATAB0_DATE_CPT_JP1


xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " & xWhere _
     & " order by SWISABSWID"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)

    newYSWISAB0 = oldYSWISAB0
    If oldYSWISAB0.SWISABKPDE <> " " Then newYSWISAB0.SWISABKPDE = "@"
    '¤JPL 20120705 If oldYSWISAB0.SWISABK999 <> " " Then newYSWISAB0.SWISABK999 = "@"
    
    V = sqlYSWISAB0_Update(newYSWISAB0, oldYSWISAB0)
    If Not IsNull(V) Then GoTo Error_MsgBox
    
    rsSab.MoveNext

Loop

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
Exit_sub:
    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If


Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub Importation_SAB_SWISABSWID_MT700()
Dim V, X As String, xMesg As String, Nb As Long
Dim xSql As String, xWhere As String
Dim rsSab As New ADODB.Recordset, rsSabX As New ADODB.Recordset
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass

currentAction = "Importation_SAB_SWISABSWID_MT700"

xWhere = "where SWISABSWID > " & mSWISABSWID_MT700 & " and SWISABWES = 'E' and SWISABWMTK = '700'"

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " & xWhere _
     & " order by SWISABSWID"
Set rsSab = cnsab.Execute(xSql)


Do While Not rsSab.EOF
    
    Nb = DateDiff("s", Format(rsSab("SWISABWAMJ"), "@@@@/@@/@@") & " " & Format(rsSab("SWISABWHMS"), "00:00:00") _
                      , Format(rsSab("SWISABXAMJ"), "@@@@/@@/@@") & " " & Format(rsSab("SWISABXHMS"), "00:00:00"))
                      
    xSql = "select CDODOSDOS , CDODOSMON , CDODOSDEV from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
         & " where CDODOSCOP = 'CDE' and CDODOSEVE <> '90' and CDODOSEXT = '" & rsSab("SWISABWL20") & "'"
    Set rsSabX = cnsab.Execute(xSql)
    
    Do While Not rsSabX.EOF
    
        X = "Message SWIFT MT700 entrant ayant la même référence que le dossier CDE " & rsSabX("CDODOSDOS")
        If Nb < 300 Then
            xMesg = X
        Else
            
            xMesg = htmlFontColor_Red & "<div align = CENTER><B> Attention</B><BR>Cette alerte peut être due à une désynchronisation des automates : <BR>" _
                  & "   - Swift Alliance vers SAB <BR>" _
                  & "   - Swift Alliance vers BIA_GOS (<= SIDE Reporting) <BR></div>" & htmlFontColor_Blue _
                  & "<BR><BR>" & X
                  
        End If
        
        xSql = "select *  from rMesg  " _
        & " where Aid = " & rsSab("SWISABWID1") _
        & " and mesg_s_umidl = " & rsSab("SWISABWIDL") _
        & " and mesg_s_umidh  =  " & rsSab("SWISABWIDH")
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
        
        If Not rsSIDE_DB.EOF Then

            Call srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
            Call cmdSendMail_SAA_Alerte_rMesg("SAB_Doublon", X, xMesg & " - " & rsSabX("CDODOSDEV") & " " & Format(rsSabX("CDODOSMON"), "### ### #### ##0.00"), "S10", "")
        End If
        rsSabX.MoveNext
    
    Loop
    rsSab.MoveNext

Loop
'=======================================================================================



GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
Exit_sub:
    


Me.Enabled = True: Me.MousePointer = 0


End Sub



Private Sub cmdSelect_SQL_5h()
Dim V, X As String
Dim xSql As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_5h"

xWhere = " where GOSEVEIDD = 0 "

Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Min, wAmjMin)
Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Max, wAmjMax)
If chkSelect_GOSDOSIAMJ = "1" Then xWhere = xWhere & " and GOSEVEUAMJ >= " & wAmjMin & " and GOSEVEUAMJ <= " & wAmjMax

'X = Trim(Mid$(cboSelect_9_SWISABKSRV, 1, 3))
'If X <> "" Then xWhere = xWhere & " and SWISABKSRV = '" & X & "'"



xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0 " & xWhere _
     & " order by GOSEVEIDE"
Set rsSab = cnsab.Execute(xSql)

fgSelect_Display_YGOSEVE0

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub cmdSelect_SQL_1()
Dim V, curX As Currency
Dim xSql As String, X As String
Dim xWhere As String, xAnd As String, xAnd_rText As String
Dim wCli As Long
Dim blnOk As Boolean
Dim wSwift_address_K As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"
blnOk = False: blnSelect_SQL_1_rText = False
xWhere = ""
X = Trim(txtSelect_SWISABOPEN)
If X = "" Then

    Call DTPicker_Amj8_tiret(txtSelect_GOSDOSIAMJ_Min, wAmj8_tiret)
    xAmj8_from_crea_date_time = wAmj8_tiret
    Call DTPicker_Amj8_tiret(txtSelect_GOSDOSIAMJ_Max, wAmj8_tiret)
    xAmj8_to_crea_date_time = wAmj8_tiret

    X = Trim(cboSelect_GOSDOSWMTK)
    If X <> "" Then
        If InStr(X, "%") Then
            xWhere = xWhere & " and mesg_type like '" & X & "'"
        Else
            If InStr(X, ",") Then
                xWhere = xWhere & " and mesg_type in ('" & Replace(X, ",", "','") & "')"
            Else
                xWhere = xWhere & " and mesg_type = '" & X & "'"
            End If
        End If
    End If
    
    X = Trim(cboSelect_GOSDOSWES)
    wSwift_address_K = "mesg_sender_swift_address"
    Select Case X
        Case "Sortant": xWhere = xWhere & " and mesg_sub_format = 'INPUT'"
                     wSwift_address_K = "mesg_receiver_swift_address"
        Case "Entrant": xWhere = xWhere & " and mesg_sub_format = 'OUTPUT'"
    End Select
    
    X = Trim(cboSelect_GOSDOSWDEV)
    If X <> "" Then xWhere = xWhere & " and x_fin_ccy = '" & X & "'"
    
    X = Trim(cboSelect_GOSDOSWBIC)
    'If X <> "" Then xWhere = xWhere & " and " & wSwift_address_K & " like '" & X & "%'"
    If X <> "" Then xWhere = xWhere & " and substring(mesg_uumid,2,11) like '" & X & "%'"
    
    X = Trim(txtSelect_GOSDOSWTRN)
    If X <> "" Then xWhere = xWhere & " and  ( mesg_trn_ref like '%" & X & "%' or mesg_rel_trn_ref like '%" & X & "%')"
    
    curX = Val(txtSelect_GOSDOSWMTD)
    If curX <> 0 Then xWhere = xWhere & " and ( x_fin_amount >=  " & curX & " and x_fin_amount < " & curX + 1 & ")"
    
    X = Trim(txtSelect_rTextField1)
    If X = "" Then
        xSql = "select * from rMesg , rInst " _
                  & "where Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
                  & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
                  & xWhere _
                  & " and  rInst.Aid =  rMesg.Aid" _
                  & " and inst_s_umidl = mesg_s_umidl" _
                  & " and inst_s_umidh  =  mesg_s_umidh and inst_num = 0"
 
    Else
        xAnd = ""
        xAnd_rText = ""
       If Trim(txtSelect_rTextField2) <> "" Then
            If optSelect_rTextField_AND Then
                xAnd = " and ( rtextField.value like '%" & Trim(txtSelect_rTextField2) & "%' or rtextField.value_memo like '%" & Trim(txtSelect_rTextField2) & "%'))"
                xAnd_rText = " and rtext.text_data_block like '%" & Trim(txtSelect_rTextField2) & "%')"
            Else
                xAnd = " or rtextField.value like '%" & Trim(txtSelect_rTextField2) & "%')"
                xAnd_rText = " or rtext.text_data_block like '%" & Trim(txtSelect_rTextField2) & "%')"
            End If
            
        Else
            xAnd = ")"
            xAnd_rText = ")"
        End If
       If Trim(txtSelect_rTextField_Code) <> "" Then
            xAnd = xAnd & " and rtextField.field_code = " & Trim(txtSelect_rTextField_Code)
        End If
        
        xSql = "select * from rMesg , rInst , rtextField " _
                  & "where ((rtextField.value like '%" & X & "%' or rtextField.value_memo like '%" & X & "%')" & xAnd _
                  & " and Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
                  & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
                  & xWhere _
                  & " and  rtextField.Aid =  rMesg.Aid" _
                  & " and rtextField.text_s_umidl = mesg_s_umidl" _
                  & " and rtextField.text_s_umidh  =  mesg_s_umidh " _
                  & " and  rInst.Aid =  rMesg.Aid" _
                  & " and inst_s_umidl = mesg_s_umidl" _
                  & " and inst_s_umidh  =  mesg_s_umidh and inst_num = 0" _
                  & " order by rMesg.Aid , mesg_s_umidl , mesg_s_umidh"
                  
        blnSelect_SQL_1_rText = True
        cmdSelect_SQL_1_rText = "select * from rMesg , rInst , rtext  " _
                  & "where (rtext.text_data_block like '%" & X & "%'" & xAnd_rText _
                  & " and Mesg_crea_date_time >= {ts '" & xAmj8_from_crea_date_time & " 00:00:00.000'} " _
                  & " and Mesg_crea_date_time < {ts '" & xAmj8_to_crea_date_time & " 23:59:59.000'} " _
                  & xWhere _
                  & " and  rtext.Aid =  rMesg.Aid" _
                 & " and rtext.text_s_umidl = mesg_s_umidl" _
                  & " and rtext.text_s_umidh  =  mesg_s_umidh " _
                  & " and  rInst.Aid =  rMesg.Aid" _
                  & " and inst_s_umidl = mesg_s_umidl" _
                  & " and inst_s_umidh  =  mesg_s_umidh and inst_num = 0"
    End If
    
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
  

    fgSelect_Display
    
Else
    blnYSWILNK0_Display = True
    xWhere = " where SWISABOPEN = " & X
    If Trim(cboSelect_SWISABOPEC) <> "" Then xWhere = xWhere & " and SWISABOPEC = '" & Trim(cboSelect_SWISABOPEC) & "'"
    

    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " & xWhere & " order by SWISABSWID"
    Set rsSab = cnsab.Execute(xSql)
      
    fgSelect_Display_YSWISAB0
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_1b()
Dim V, curX As Currency
Dim xSql As String, X As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean
Dim wSwift_address_K As String

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1b"
blnOk = False
xWhere = ""
X = Trim(txtSelect_SWISABOPEN)
If X = "" Then

    Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Min, wAmjMin)
    Call DTPicker_Control(txtSelect_GOSDOSIAMJ_Max, wAmjMax)

    X = Trim(cboSelect_GOSDOSWMTK)
    If cmdSelect_SQL_K = "1trf" Then
        If X = "" Then
            xWhere = xWhere & " and SWISABWMTK in ('103' , '202')"
        Else
            xWhere = xWhere & " and SWISABWMTK = '" & X & "'"
        End If
    Else
        If X <> "" Then
            If InStr(X, "%") Then
                xWhere = xWhere & " and SWISABWMTK like '" & X & "'"
            Else
                If InStr(X, ",") Then
                    xWhere = xWhere & " and SWISABWMTK in ('" & Replace(X, ",", "','") & "')"
                Else
                    xWhere = xWhere & " and SWISABWMTK = '" & X & "'"
                End If
            End If
        End If
    End If
    
    X = Trim(cboSelect_GOSDOSWES)
    If X <> "" Then xWhere = xWhere & " and SWISABWES = '" & Mid$(X, 1, 1) & "'"
    
    X = Trim(cboSelect_GOSDOSWDEV)
    If X <> "" Then xWhere = xWhere & " and SWISABWDEV = '" & X & "'"
    
    X = Trim(cboSelect_GOSDOSWBIC)
    If X <> "" Then xWhere = xWhere & " and SWISABWBIC like '" & X & "%'"
    
    X = Trim(txtSelect_GOSDOSWTRN)
    If X <> "" Then xWhere = xWhere & " and  ( SWISABWL20 like '%" & X & "%' or SWISABWN20 like '%" & X & "%')"
    
    If cboSelect_GOSDOSKSRV.Visible Then
        X = Trim(Mid$(cboSelect_GOSDOSKSRV, 1, 3))
        If X <> "" Then
            If X = "S00" Then
                xWhere = xWhere & " and SWISABKSRV = '" & X & "'"
            Else
                If chkSelect_GOSDOSKSRV = "1" Then
                    xWhere = xWhere & " and SWISABKSRV in ('S00' , '" & X & "')"
                Else
                    xWhere = xWhere & " and SWISABKSRV = '" & X & "'"
               End If
                
            End If
        End If
            X = Trim(cboSelect_SWISABWSTA)
        If X <> "" Then
            Select Case Mid$(X, 1, 1)
                Case "L": xWhere = xWhere & " and SWISABWSTA in (' ', '#')"
                          wAmjMin = "00000000": wAmjMax = "99999999"
                Case "N": xWhere = xWhere & " and SWISABWSTA in ('E', '#')"
                Case "A": xWhere = xWhere & " and SWISABWSTA = 'V'"
            End Select
        End If
    End If
    
    If Trim(cboSelect_SWISABOPEC) <> "" Then xWhere = xWhere & " and SWISABOPEC = '" & Trim(cboSelect_SWISABOPEC) & "'"
    
    If cmdSelect_SQL_K = "1?" Then xWhere = xWhere & " and SWISABOPEN = 0"
    
    curX = Val(txtSelect_GOSDOSWMTD)
    If curX <> 0 Then xWhere = xWhere & " and ( SWISABWMTD >=  " & curX & " and SWISABWMTD < " & curX + 1 & ")"
    
'__________________________________________________________________________________________________________
    If cmdSelect_SQL_K = "1trf" Then
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0, " _
                  & paramIBM_Library_SABSPE & ".YSWISAB1 " _
                  & " where SWISABWAMJ >= '" & wAmjMin & "' and SWISABWAMJ <= '" & wAmjMax & "'" _
                  & xWhere & " and SWISAB1ID = SWISABSWID order by SWISABSWID"
    Else
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
                  & " where SWISABWAMJ >= '" & wAmjMin & "' and SWISABWAMJ <= '" & wAmjMax & "'" _
                  & xWhere & " order by SWISABSWID"
    End If
    
    Set rsSab = cnsab.Execute(xSql)
        
    fgSelect_Display_YSWISAB0
    
Else
    blnYSWILNK0_Display = True
    xWhere = " where SWISABOPEN = " & X
    If Trim(cboSelect_SWISABOPEC) <> "" Then xWhere = xWhere & " and SWISABOPEC = '" & Trim(cboSelect_SWISABOPEC) & "'"
    
    If cmdSelect_SQL_K = "1trf" Then
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
            & " left outer join " & paramIBM_Library_SABSPE & ".YSWISAB1 on SWISAB1ID = SWISABSWID " _
            & xWhere & " order by SWISABSWID"
    Else
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " & xWhere & " order by SWISABSWID"
    End If
    Set rsSab = cnsab.Execute(xSql)
      
    fgSelect_Display_YSWISAB0
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_1L()
Dim V
Dim xSql As String, X As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean
Dim wSwift_address_K As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1L"
blnOk = False: blnSelect_SQL_1_rText = False
chkSIDE_DB_Show.value = "1"
xWhere = ""
'==========================================================================================================
If xAmj8_1Live = "" Then
    xAmj8_1Live = YBIATAB0_DATE_CPT_AP1
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
         & " where SWISABWSTA in(' ', '#') order by SWISABWAMJ"
    
    Set rsSab = cnsab.Execute(xSql)
    
   If Not rsSab.EOF Then xAmj8_1Live = rsSab("SWISABWAMJ")
            
    xAmj8_1Live = "{ts '" & Mid$(xAmj8_1Live, 1, 4) & "-" & Mid$(xAmj8_1Live, 5, 2) & "-" & Mid$(xAmj8_1Live, 7, 2) & " 00:00:00.000'}"

'==========================================================================================================


End If

'    xSQL = "select * from rMesg , rInst " _
'              & "where mesg_status <> 'COMPLETED'" _
'              & " and Mesg_crea_date_time >= {ts '2008-01-01 00:00:00.000'} " _
'              & " and  rInst.Aid =  rMesg.Aid" _
'              & " and inst_s_umidl = mesg_s_umidl" _
'              & " and inst_s_umidh  =  mesg_s_umidh and inst_num = 0"

If cmdSelect_SQL_K = "1Live_Sortan" Then
    xSql = "select * from rMesg , rInst " _
              & "where Mesg_crea_date_time >= " & xAmj8_1Live & " and mesg_status <> 'COMPLETED'" _
              & " and substring(mesg_uumid, 1, 1) = 'I' and mesg_type > '099'" _
              & " and  rInst.Aid =  rMesg.Aid" _
              & " and inst_s_umidl = mesg_s_umidl" _
              & " and inst_s_umidh  =  mesg_s_umidh and inst_num = 0" _
              & " and inst_status <> 'COMPLETED'"
Else
    xSql = "select * from rMesg , rInst " _
              & "where Mesg_crea_date_time >= " & xAmj8_1Live & " and mesg_status <> 'COMPLETED'" _
              & " and substring(mesg_uumid, 1, 1) = 'O'" _
              & " and  rInst.Aid =  rMesg.Aid" _
              & " and inst_s_umidl = mesg_s_umidl" _
              & " and inst_s_umidh  =  mesg_s_umidh and inst_num = 1" _
              & " and inst_status <> 'COMPLETED'"
End If

    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
  

    fgSelect_Display
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_1Lmail()
Dim V
Dim xSql As String, X As String, wDateTime As String
Dim xinst_crea_date_time As String, xinst_rp_name As String
Dim blnUpdate As Boolean, blnMail As Boolean
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1Lmail"
wDateTime = Date & " " & Time
blnUpdate = False
If SAA_Alerte_Live_Entrant.Umidh = 0 Then
        New_YBIATAB0.BIATABID = "SAA_Alerte"
        New_YBIATAB0.BIATABK1 = "Live"
        New_YBIATAB0.BIATABK2 = ""
        New_YBIATAB0.BIATABTXT = ""
        
        Call lstErr_AddItem(lstErr, cmdContext, "SAA_1Lmail (init):" & wDateTime): DoEvents
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
             & " where BIATABID = '" & New_YBIATAB0.BIATABID & "' and BIATABK1 = '" & New_YBIATAB0.BIATABK1 & "'"
        Set rsSab = cnsab.Execute(xSql)
        If rsSab.EOF Then
            New_YBIATAB0.BIATABTXT = String(60, "9")
            Parametrage_New

        Else
            X = rsSab("BIATABTXT")
            SAA_Alerte_Live_Entrant.Aid = Val(Mid$(X, 1, 9))
            SAA_Alerte_Live_Entrant.Umidl = Val(Mid$(X, 11, 9))
            SAA_Alerte_Live_Entrant.Umidh = Val(Mid$(X, 21, 9))
            
            SAA_Alerte_Live_Sortant.Aid = Val(Mid$(X, 31, 9))
            SAA_Alerte_Live_Sortant.Umidl = Val(Mid$(X, 41, 9))
            SAA_Alerte_Live_Sortant.Umidh = Val(Mid$(X, 51, 9))
        End If
End If

'SAA_Alerte_Live_Entrant
X = Mid$(YBIATAB0_DATE_CPT_JP1, 1, 4) & "-" & Mid$(YBIATAB0_DATE_CPT_JP1, 5, 2) & "-" & Mid$(YBIATAB0_DATE_CPT_JP1, 7, 2) & " 00:00:00.000"

xSql = "select * from rMesg , rInst " _
          & "where mesg_status <> 'COMPLETED'" _
          & " and Mesg_crea_date_time >= {ts '" & X & "'} " _
          & " and  rInst.Aid =  rMesg.Aid" _
          & " and inst_s_umidl = mesg_s_umidl" _
          & " and inst_s_umidh  =  mesg_s_umidh and inst_num = 0" _
          & " order by rMesg.Aid , mesg_s_umidl desc ,mesg_s_umidh desc"

Set rsSIDE_Loop = cnSIDE_DB.Execute(xSql)
  
Do While Not rsSIDE_Loop.EOF
    If Not IsNull(rsSIDE_Loop("inst_rp_name")) Then
        xinst_rp_name = rsSIDE_Loop("inst_rp_name")
        blnMail = False
        Select Case xinst_rp_name
            Case "_MP_mod_emi_secu"
                If rsSIDE_Loop("mesg_s_umidl") < SAA_Alerte_Live_Sortant.Umidl Then
                    blnMail = True
                Else
                    If rsSIDE_Loop("mesg_s_umidh") < SAA_Alerte_Live_Sortant.Umidh Then blnMail = True
                End If
                
                If blnMail Then
                    X = "Message SWIFT sortant " & rsSIDE_Loop("mesg_type") & " NON EMIS, en attente dans la queue " & xinst_rp_name
                    Call srvrMesg_GetBuffer_ODBC(rsSIDE_Loop, xrMesg)
                    Call cmdSendMail_SAA_Alerte_rMesg("SAA_Live", X, X, "S62", rsSIDE_Loop("inst_unit_name"))
                    blnUpdate = True
                    SAA_Alerte_Live_Sortant.Aid = rsSIDE_Loop("Aid")
                    SAA_Alerte_Live_Sortant.Umidl = rsSIDE_Loop("mesg_s_umidl")
                    SAA_Alerte_Live_Sortant.Umidh = rsSIDE_Loop("mesg_s_umidh")
                End If
    
            Case "_MP_mod_reception"
                    xinst_crea_date_time = rsSIDE_Loop("inst_crea_date_time")
                    V = DateDiff("n", xinst_crea_date_time, wDateTime)
                    If V > 30 Then
                        If rsSIDE_Loop("mesg_s_umidl") < SAA_Alerte_Live_Entrant.Umidl Then
                            blnMail = True
                        Else
                            If rsSIDE_Loop("mesg_s_umidh") < SAA_Alerte_Live_Entrant.Umidh Then blnMail = True
                        End If
                        
                        If blnMail Then
                            X = "Message SWIFT entrant " & rsSIDE_Loop("mesg_type") & " NON Routé, en attente dans la queue " & xinst_rp_name
                            Call srvrMesg_GetBuffer_ODBC(rsSIDE_Loop, xrMesg)
                            Call cmdSendMail_SAA_Alerte_rMesg("SAA_Live", X, X, "S62", "")
                            blnUpdate = True
                            SAA_Alerte_Live_Entrant.Aid = rsSIDE_Loop("Aid")
                            SAA_Alerte_Live_Entrant.Umidl = rsSIDE_Loop("mesg_s_umidl")
                            SAA_Alerte_Live_Entrant.Umidh = rsSIDE_Loop("mesg_s_umidh")
                        End If
                    End If
        End Select
    End If
    
    rsSIDE_Loop.MoveNext

Loop

If blnUpdate Then
        Old_YBIATAB0.BIATABID = "SAA_Alerte"
        Old_YBIATAB0.BIATABK1 = "Live"
        Old_YBIATAB0.BIATABK2 = ""
        Old_YBIATAB0.BIATABTXT = ""
        New_YBIATAB0 = Old_YBIATAB0
        New_YBIATAB0.BIATABTXT = Format(SAA_Alerte_Live_Entrant.Aid, "000000000") & " " _
                               & Format(SAA_Alerte_Live_Entrant.Umidl, "000000000") & " " _
                               & Format(SAA_Alerte_Live_Entrant.Umidh, "00000000") & " " _
                               & Format(SAA_Alerte_Live_Sortant.Aid, "000000000") & " " _
                               & Format(SAA_Alerte_Live_Sortant.Umidl, "000000000") & " " _
                               & Format(SAA_Alerte_Live_Sortant.Umidh, "00000000")
        Parametrage_Update

End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub Importation_Jrnl()
Dim V, xSql As String, X As String, K As Long, xDbl As Double
Dim xSubject As String, wDateTime As String
Dim blnUpdate As Boolean, blnTransaction As Boolean, blnTOPK As Boolean
Dim wMail_To As String
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "Importation_Jrnl"
wDateTime = Date & " " & Time
blnUpdate = False
If SAA_Alerte_Jrnl.Umidh = 0 Then

        Importation_Jrnl_Init
        
        New_YBIATAB0.BIATABID = "SAA_Alerte"
        New_YBIATAB0.BIATABK1 = "Jrnl"
        New_YBIATAB0.BIATABK2 = ""
        New_YBIATAB0.BIATABTXT = ""
        
        Call lstErr_AddItem(lstErr, cmdContext, "Importation_Jrnl (init):" & wDateTime): DoEvents
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
             & " where BIATABID = '" & New_YBIATAB0.BIATABID & "' and BIATABK1 = '" & New_YBIATAB0.BIATABK1 & "'"
        Set rsSab = cnsab.Execute(xSql)
        If rsSab.EOF Then
            SAA_Alerte_Jrnl.Aid = 0
            SAA_Alerte_Jrnl.Umidl = 1200802192 - DateDiff("s", "01/01/2000 00:00:00", "26/01/2012 00:00:00")
            SAA_Alerte_Jrnl.Umidh = 9999999999#
            New_YBIATAB0.BIATABTXT = Format(SAA_Alerte_Jrnl.Aid, "000000000") & " " _
                                   & Format(SAA_Alerte_Jrnl.Umidl, "000000000") & " " _
                                   & Format(SAA_Alerte_Jrnl.Umidh, "0000000000") & " "
            Parametrage_New

        Else
            X = rsSab("BIATABTXT")
            SAA_Alerte_Jrnl.Aid = Val(Mid$(X, 1, 9))
            SAA_Alerte_Jrnl.Umidl = Val(Mid$(X, 11, 9))
            SAA_Alerte_Jrnl.Umidh = Val(Mid$(X, 21, 10))
            
 '_______________________________________________________________________________________________________
            xSql = "select jrnl_date_time  from rJrnl " _
                      & " where Aid = 0" _
                      & " and jrnl_rev_date_time <= " & SAA_Alerte_Jrnl.Umidl
            
            Set rsSIDE_Loop = cnSIDE_DB.Execute(xSql)
              
            If Not rsSIDE_Loop.EOF Then last_Jrnl_date_time_EVE = rsSIDE_Loop("jrnl_date_time")
  '_______________________________________________________________________________________________________
       End If
End If

xDbl = 1200802192 - DateDiff("s", "01/01/2000 00:00:00", wDateTime)

If SAA_Alerte_Jrnl.Umidl - xDbl > 300 Then
    blnUpdate = True
    If SAA_Alerte_Jrnl.Umidl - xDbl > 900 Then Importation_Jrnl_Init
End If

xSql = "select Aid , jrnl_rev_date_time , jrnl_seq_nbr , jrnl_comp_name , jrnl_event_num , jrnl_oper_nickname , jrnl_event_name" _
     & " , jrnl_date_time , cast (Jrnl_merged_text as varchar(200)) as Jrnl_merged_text from rJrnl " _
          & " where Aid = 0" _
          & " and jrnl_rev_date_time <= " & SAA_Alerte_Jrnl.Umidl _
          & " and  jrnl_rev_date_time >= " & xDbl _
          & " order by jrnl_rev_date_time desc , jrnl_seq_nbr desc"

Set rsSIDE_Loop = cnSIDE_DB.Execute(xSql)
  
Do While Not rsSIDE_Loop.EOF
    blnOk = True
    
    If rsSIDE_Loop("jrnl_rev_date_time") < SAA_Alerte_Jrnl.Umidl Then
    Else
        If rsSIDE_Loop("jrnl_rev_date_time") > SAA_Alerte_Jrnl.Umidl Then
            blnOk = False
        Else
            If rsSIDE_Loop("jrnl_seq_nbr") >= SAA_Alerte_Jrnl.Umidh Then blnOk = False
        End If
    End If
    
    
    If blnOk Then
        xrJrnl.jrnl_event_num = rsSIDE_Loop("jrnl_event_num")
        
        SAA_Alerte_Jrnl.Aid = rsSIDE_Loop("Aid")
        SAA_Alerte_Jrnl.Umidl = rsSIDE_Loop("jrnl_rev_date_time")
        SAA_Alerte_Jrnl.Umidh = rsSIDE_Loop("jrnl_seq_nbr")
        
        newYSAAJRN0.SAAJRNAID = SAA_Alerte_Jrnl.Aid
        newYSAAJRN0.SAAJRNAMJH = SAA_Alerte_Jrnl.Umidl
        newYSAAJRN0.SAAJRNSEQ = SAA_Alerte_Jrnl.Umidh
        
        last_Jrnl_date_time_EVE = rsSIDE_Loop("jrnl_date_time")
        If rsSIDE_Loop("jrnl_comp_name") = "SIS" Then
            If rsSIDE_Loop("jrnl_event_num") = 8066 Then   'rsSIDE_Loop("jrnl_event_num") = 8064 Or
                last_Jrnl_date_time_ES = rsSIDE_Loop("jrnl_date_time")
            End If
        End If
        
        For K = 1 To arrJrnl_Nb
            If xrJrnl.jrnl_event_num = arrJrnl_Event_Num(K) Then
                If arrJrnl_Comp_Name(K) = rsSIDE_Loop("jrnl_comp_name") Then
                
                        xrJrnl.jrnl_merged_text = rsSIDE_Loop("Jrnl_merged_text")
    
                    
                    If arrJrnl_Top(K) <> "" Then
                        If Not blnTransaction Then
                            V = cnSAB_Transaction("BeginTrans")
                            If Not IsNull(V) Then GoTo Error_MsgBox
                        End If
                        
                        blnTransaction = True
                        xrJrnl.jrnl_oper_nickname = rsSIDE_Loop("jrnl_oper_nickname")
                        xrJrnl.jrnl_comp_name = rsSIDE_Loop("jrnl_comp_name")
                        
                        newYSAAJRN0.SAAJRNEVEC = xrJrnl.jrnl_comp_name
                        newYSAAJRN0.SAAJRNEVEN = xrJrnl.jrnl_event_num
    
                        Importation_Jrnl_Top True
                        blnTOPK = True
                    Else
                        blnTOPK = False
                    End If
    
                    If arrJrnl_Alerte(K) <> "" Then
                        wMail_To = Trim(arrJrnl_Alerte(K))
                        
                        newYSAAJRN0.SAAJRNEVEN = xrJrnl.jrnl_event_num
                        If Not blnTOPK Then Importation_Jrnl_Top False
                        xSubject = Trim(rsSIDE_Loop("jrnl_event_name")) & " : " & Trim(newYSAAJRN0.SAAJRNTOPX)
                        If IsNull(cmdSelect_SQL_YSAAJRN0_rMesg(newYSAAJRN0.SAAJRNTOPX, newYSAAJRN0.SAAJRNSUFX)) Then
                        
                        If newYSAAJRN0.SAAJRNEVEN = 10007 Then wMail_To = Importation_Jrnl_Top_10007
    
                            xSql = "select * from rMesg , rInst " _
                                & " where rMesg.Aid = " & Mesg_aid _
                                & " and mesg_s_umidl = " & mesg_s_umidl _
                                & " and mesg_s_umidh = " & mesg_s_umidh _
                                & " and  rInst.Aid =  rMesg.Aid" _
                                & " and inst_s_umidl = mesg_s_umidl" _
                                & " and inst_s_umidh  =  mesg_s_umidh and inst_num = 0"
                            Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
                            If rsSIDE_DB.EOF Then
                                Call srvrMesg_Init(xrMesg)
                                If wMail_To = "S=" Then wMail_To = "S01"
                                Call cmdSendMail_SAA_Alerte_rMesg("SAA_Event", xSubject, xrJrnl.jrnl_merged_text, wMail_To, "")
                                blnUpdate = True
                            Else
                                Call srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
                                If wMail_To = "S=" Then wMail_To = rsSIDE_DB("inst_unit_name")
                                Call cmdSendMail_SAA_Alerte_rMesg("SAA_Event", xSubject, xrJrnl.jrnl_merged_text, wMail_To, "")
                                blnUpdate = True
                            End If
                        Else
                            xSql = "select * from rJrnl " _
                                & " where Aid = " & rsSIDE_Loop("Aid") _
                                & " and jrnl_rev_date_time = " & rsSIDE_Loop("jrnl_rev_date_time") _
                                & " and jrnl_seq_nbr = " & rsSIDE_Loop("jrnl_seq_nbr")
                                
                            Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
                            If wMail_To = "S=" Then wMail_To = "S01"
                            If rsSIDE_DB.EOF Then
                                Call srvrJrnl_GetBuffer_ODBC(rsSIDE_Loop, xrJrnl)
                            Else
                                Call srvrJrnl_GetBuffer_ODBC(rsSIDE_DB, xrJrnl)
                            End If
                            blnUpdate = True
                            Call cmdSendMail_SAA_Alerte_rJrnl("SAA_Event", xSubject, xrJrnl.jrnl_merged_text, wMail_To, "")
                        End If
                        
                        
                    End If
                    Exit For
                End If
            Else
                'If xrJrnl.jrnl_event_num < arrJrnl_Event_Num(K) Then Exit For
            End If
        Next K
    End If
    
    rsSIDE_Loop.MoveNext
Loop

If blnUpdate Then
        Old_YBIATAB0.BIATABID = "SAA_Alerte"
        Old_YBIATAB0.BIATABK1 = "Jrnl"
        Old_YBIATAB0.BIATABK2 = ""
        Old_YBIATAB0.BIATABTXT = ""
        New_YBIATAB0 = Old_YBIATAB0
        New_YBIATAB0.BIATABTXT = Format(SAA_Alerte_Jrnl.Aid, "000000000") & " " _
                               & Format(SAA_Alerte_Jrnl.Umidl, "000000000") & " " _
                               & Format(SAA_Alerte_Jrnl.Umidh, "0000000000") & " "
        
        If Not blnTransaction Then
            Parametrage_Update
        Else
            V = sqlYBIATAB0_Update(New_YBIATAB0, Old_YBIATAB0)
        End If

End If

GoTo Exit_sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If
End Sub



Private Sub Importation_SAA()
Dim V, X As String, K As Long, K2 As Long, Nb As Integer
Dim xSql As String
Dim xAMJ As String, xHMS As String, xUUMID As String
Dim xYSWISAB0 As typeYSWISAB0, xYSWILNK0 As typeYSWILNK0
Dim arrYSWISAB0(5000) As typeYSWISAB0
Dim mAMJ_R As Long
Dim xrMesg As typerMesg
Dim wK115 As String
On Error GoTo Error_Handler

'________________________________________________________________________
Call lstErr_Clear(lstErr, cmdContext, "Importation SAA : 1"): DoEvents
'________________________________________________________________________

currentAction = "Importation_SAA => YSWISAB0 "
mAMJ_R = Val(dateElp("MoisAdd", -3, DSys))

Nb = 0
mSWISABSWID_MT700 = 0

xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YSWISAB0 "
Set rsSab = cnsab.Execute(xSql)

xSql = "select SWISABSWID,SWISABWAMJ,SWISABWHMS from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABSWID >= " & rsSab(0) & " order by SWISABSWID desc "
Set rsSab = cnsab.Execute(xSql)

If rsSab.EOF Then
    ' Exit Sub
    xAMJ = "20040301"
    X = "2004-03-01" & " " & "00:00:00.000"
Else
    importSWISABSWID = rsSab("SWISABSWID")
    xAMJ = rsSab("SWISABWAMJ")
    xHMS = Format$(rsSab("SWISABWHMS"), "000000")
    X = Mid$(xAMJ, 1, 4) & "-" & Mid$(xAMJ, 5, 2) & "-" & Mid$(xAMJ, 7, 2) & " " & Mid$(xHMS, 1, 2) & ":" & Mid$(xHMS, 3, 2) & ":00.000"
End If

'$JPL 2016-01-13
blnK115 = False
mK115_SWISABSWID = importSWISABSWID

'==================================================================

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation SAA : 2"): DoEvents
'________________________________________________________________________

 xSql = "select * from rMesg " _
           & " where Mesg_crea_date_time >= {ts '" & X & "'} order by Mesg_crea_date_time"
   
   
 '?????????????????????????????????????????????????????????????????
 Dim xMax As String, xWhere As String, X2 As String
xMax = dateElp("Jour", 10, xAMJ)
X2 = Mid$(xMax, 1, 4) & "-" & Mid$(xMax, 5, 2) & "-" & Mid$(xMax, 7, 2) & " " & "00:00:00.000"
xWhere = " and Mesg_crea_date_time <= {ts '" & X2 & "'} "

 xSql = "select * from rMesg " _
           & " where Mesg_crea_date_time >= {ts '" & X & "'}" & xWhere & " order by Mesg_crea_date_time"

'??????????????????????????????????????????????????????????????????

   
   
Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
  
Do While Not rsSIDE_DB.EOF
    If Not IsNull(rsSIDE_DB("mesg_type")) Then
            
        last_mesg_crea_date_time = rsSIDE_DB("mesg_crea_date_time")
            
        X = rsSIDE_DB("mesg_type")
'        If X < "100" Or X = "198" Or X = "298" Or X = "960" Or X = "961" Or X = "962" Or X = "963" Or X = "964" Then
        If X < "100" Or X = "960" Or X = "961" Or X = "962" Or X = "963" Or X = "964" Then
        Else

            Call srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
            
            last_mesg_crea_date_time = xrMesg.mesg_crea_date_time
            
            rsYSWISAB0_Init xYSWISAB0
            xYSWISAB0.SWISABWID1 = xrMesg.Aid
            xYSWISAB0.SWISABWIDL = xrMesg.mesg_s_umidl
            xYSWISAB0.SWISABWIDH = xrMesg.mesg_s_umidh
            
            
            xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
                 & " where SWISABWID1 = " & xYSWISAB0.SWISABWID1 _
                 & " and   SWISABWIDL = " & xYSWISAB0.SWISABWIDL _
                 & " and   SWISABWIDH = " & xYSWISAB0.SWISABWIDH _
                 
            Set rsSab = cnsab.Execute(xSql)
            
            If rsSab.EOF Then

                xYSWISAB0.SWISABWMTK = xrMesg.mesg_type
                xUUMID = xrMesg.mesg_uumid
                If Mid$(xUUMID, 1, 1) = "I" Then
                    xYSWISAB0.SWISABWES = "S"
                    xYSWISAB0.SWISABWN20 = Replace(xrMesg.mesg_trn_ref, "'", " ")
                    xYSWISAB0.SWISABWL20 = Replace(xrMesg.mesg_rel_trn_ref, "'", " ")
                    If xYSWISAB0.SWISABWMTK = "950" Then xYSWISAB0.SWISABZSWI = -1
                    If xrMesg.mesg_crea_rp_name = "_MP_creation" Then xYSWISAB0.SWISABZSWI = -2
                Else
                    xYSWISAB0.SWISABWES = "E"
                    xYSWISAB0.SWISABWL20 = Replace(xrMesg.mesg_trn_ref, "'", " ")
                    xYSWISAB0.SWISABWN20 = Replace(xrMesg.mesg_rel_trn_ref, "'", " ")
                End If
                
                xYSWISAB0.SWISABWBIC = Mid$(xUUMID, 2, 11)
                If Not IsNull(xrMesg.x_fin_ccy) Then
                    xYSWISAB0.SWISABWDEV = xrMesg.x_fin_ccy
                    xYSWISAB0.SWISABWMTD = CCur(xrMesg.x_fin_amount)
                End If
                xYSWISAB0.SWISABWSRV = xrMesg.x_inst0_unit_name
                Call dateJma10_Amj(Mid$(xrMesg.mesg_crea_date_time, 1, 10), X)
                xYSWISAB0.SWISABWAMJ = Val(X)
                X = Mid$(xrMesg.mesg_crea_date_time, 12, 8)
                xYSWISAB0.SWISABWHMS = Val(Mid$(X, 1, 2) & Mid$(X, 4, 2) & Mid$(X, 7, 2))
                
                Call dateJma10_Amj(Mid$(xrMesg.last_update, 1, 10), X)
                xYSWISAB0.SWISABXAMJ = Val(X)
                X = Mid$(xrMesg.last_update, 12, 8)
                xYSWISAB0.SWISABXHMS = Val(Mid$(X, 1, 2) & Mid$(X, 4, 2) & Mid$(X, 7, 2))
                
                If Not IsNull(xrMesg.mesg_possible_dup_creation) Then
                    If Trim(xrMesg.mesg_possible_dup_creation) <> "" Then
                        xYSWISAB0.SWISABKPDE = "!"
                    End If
                End If
'$JPL 20110720 ________________________________________________________________

                'xYSWISAB0.SWISABKSRV = SAA_2_BIA_Unit(xYSWISAB0.SWISABWSRV)
                '$JPL 20120611  ________________________________________________________________
                Select Case xYSWISAB0.SWISABWSRV
                    Case "SOBF", "ORPA", "GDMP": xYSWISAB0.SWISABKSRV = "S01"
                    Case "SOBI": xYSWISAB0.SWISABKSRV = "S10"
                    Case "DAFI", "BOTC": xYSWISAB0.SWISABKSRV = "S32"
                    Case "DCOM": xYSWISAB0.SWISABKSRV = "S41"
                    Case Else: xYSWISAB0.SWISABKSRV = "S00"
                End Select
                
                If xYSWISAB0.SWISABWES = "S" Then
                
                    ''If Trim(xrMesg.mesg_status) = "COMPLETED" Then xYSWISAB0.SWISABWSTA = "V"
               Else
                    'If Mid$(xYSWISAB0.SWISABWMTK, 2, 2) = "99" Then
                    If Mid$(xYSWISAB0.SWISABWMTK, 2, 1) = "9" Then   '$JPL 2012-10-18
                        xYSWISAB0.SWISABK999 = "!"
                    Else
                        If xYSWISAB0.SWISABKSRV = "S00" Then
                            Select Case Mid$(xYSWISAB0.SWISABWMTK, 1, 1)
                                Case "7": xYSWISAB0.SWISABKSRV = "S10"
                                Case "3", "4", "5": xYSWISAB0.SWISABKSRV = "S32"
                            End Select
                        End If
                    End If
    
                End If
'$JPL 20110720 ________________________________________________________________
                If xYSWISAB0.SWISABWMTK = "950" Then
                    xYSWISAB0.SWISABKSRV = "S11"
                    xYSWISAB0.SWISABSER = "00"
                    xYSWISAB0.SWISABSSE = "00"
                    xYSWISAB0.SWISABOPEC = "XXX"
                    xYSWISAB0.SWISABOPEN = -1
                    ''xYSWISAB0.SWISABWSTA = "V"
                End If
                
                Nb = Nb + 1
                xYSWISAB0.SWISABSWID = Nb + importSWISABSWID
                arrYSWISAB0(Nb) = xYSWISAB0
                If Nb = 5000 Then Exit Do
            End If
        End If
    End If
    rsSIDE_DB.MoveNext
    DoEvents

Loop

Call lstErr_AddItem(lstErr, cmdContext, "SAA : " & Nb & " > " & xYSWISAB0.SWISABSWID): DoEvents
         
'________________________________________________________________________

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation SAA : 3"): DoEvents
'________________________________________________________________________


For K = 1 To Nb
    If arrYSWISAB0(K).SWISABWES = "E" Then

        xYSWISAB0 = arrYSWISAB0(K)
'¤JPL 2015-12-07
        If xYSWISAB0.SWISABWMTK = "198" Or xYSWISAB0.SWISABWMTK = "298" Then
            X = Importation_SAA_198(xYSWISAB0.SWISABWID1, xYSWISAB0.SWISABWIDL, xYSWISAB0.SWISABWIDH, wK115)
            arrYSWISAB0(K).SWISABWN20 = Replace(Trim(Replace(X, Asc13, "")), "'", " ")
            arrYSWISAB0(K).SWISABK999 = " "
            If wK115 <> "N" Then
                blnK115 = True
                arrYSWISAB0(K).SWISABWSTA = wK115
            End If
            
        Else
            If xYSWISAB0.SWISABWMTK = "700" Then mSWISABSWID_MT700 = importSWISABSWID
        End If
        
        If xYSWISAB0.SWISABWMTK = "103" Or xYSWISAB0.SWISABWMTK = "202" Or xYSWISAB0.SWISABWMTK = "200" _
        Or xYSWISAB0.SWISABWMTK = "700" Then
        
            xSql = "select count(*)  as Tally from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
             & " where SWISABWES = 'E'" _
             & " and SWISABWBIC = '" & xYSWISAB0.SWISABWBIC & "'" _
             & " and SWISABWL20 = '" & Replace(xYSWISAB0.SWISABWL20, "'", "''") & "'" _
             & " and SWISABWMTK = '" & xYSWISAB0.SWISABWMTK & "'"
            Set rsSab = cnsab.Execute(xSql)
            If rsSab("Tally") <> 0 Then
                arrYSWISAB0(K).SWISABK20 = "!"
                If arrYSWISAB0(K).SWISABKPDE = " " Then
                    xSql = "select count(*)  as Tally from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
                    & " where SWISABWES = 'E'" _
                    & " and SWISABWBIC = '" & xYSWISAB0.SWISABWBIC & "'" _
                    & " and SWISABWL20 = '" & xYSWISAB0.SWISABWL20 & "'" _
                    & " and SWISABWMTK = '" & xYSWISAB0.SWISABWMTK & "'" _
                    & " and SWISABWDEV = '" & xYSWISAB0.SWISABWDEV & "'" _
                    & " and SWISABWMTD = " & cur_P(xYSWISAB0.SWISABWMTD) _
                    & " and SWISABWAMJ >= " & mAMJ_R
                    Set rsSab = cnsab.Execute(xSql)
                    If rsSab("Tally") <> 0 Then arrYSWISAB0(K).SWISABKPDE = "?"
                End If
'===========================================================================================
           Else
                For K2 = 1 To K - 1
                    If xYSWISAB0.SWISABWBIC = arrYSWISAB0(K2).SWISABWBIC _
                   And xYSWISAB0.SWISABWL20 = arrYSWISAB0(K2).SWISABWL20 _
                   And xYSWISAB0.SWISABWMTK = arrYSWISAB0(K2).SWISABWMTK Then
                        arrYSWISAB0(K).SWISABK20 = "!"
                        
                        If arrYSWISAB0(K).SWISABKPDE = " " Then
                             If xYSWISAB0.SWISABWDEV = arrYSWISAB0(K2).SWISABWDEV _
                             And xYSWISAB0.SWISABWMTD = arrYSWISAB0(K2).SWISABWMTD Then arrYSWISAB0(K2).SWISABKPDE = "?"
                        End If
                        
                        Exit For
                    End If
                Next K2
                
            End If
        End If
    End If
Next K
Call lstErr_AddItem(lstErr, cmdContext, "Importation SAA : " & Nb & " > " & xYSWISAB0.SWISABSWID): DoEvents

'________________________________________________________________________

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation SAA : 4 TRANSACTION début"): DoEvents
'________________________________________________________________________

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

For K = 1 To Nb
    If arrYSWISAB0(K).SWISABWSRV = "None" Then arrYSWISAB0(K).SWISABWSRV = ""
'$jpl==============================================
    'Select Case Trim(arrYSWISAB0(K).SWISABWN20)
    '    Case "None", "NONREF": arrYSWISAB0(K).SWISABWN20 = ""
    'End Select
    'Select Case Trim(arrYSWISAB0(K).SWISABWL20)
    '    Case "None", "NONREF": arrYSWISAB0(K).SWISABWL20 = ""
    'End Select
'$jpl==============================================
    V = sqlYSWISAB0_Insert(arrYSWISAB0(K))
    If Not IsNull(V) Then GoTo Exit_sub
    
        
Next K

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation SAA : 5 TRANSACTION fin"): DoEvents
'________________________________________________________________________

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation SAA : 6 exit"): DoEvents
'________________________________________________________________________

End Sub

Public Function Importation_SAA_198(lSWISABWID1 As Long, lSWISABWIDL As Long, lSWISABWIDH As Long, lK115 As String) As String
Dim xSql As String, X As String, K As Integer, K2 As Integer, K3 As Integer
Dim mField As String
Dim wText_Data_Block As String

On Error GoTo Error_Handler
Importation_SAA_198 = ""
lK115 = " "
'==================================================================
xSql = "select *  from rtextField  " _
& "where Aid = " & lSWISABWID1 _
& " and text_s_umidl = " & lSWISABWIDL _
& " and text_s_umidh  =  " & lSWISABWIDH _
& " order by field_cnt"
Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then
    Do While Not rsSIDE_DB.EOF
        mField = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
        Select Case mField
            Case "108": Importation_SAA_198 = rsSIDE_DB("value") ': Exit Do
            Case "115": lK115 = Mid$(rsSIDE_DB("value"), 1, 1)

        End Select
        rsSIDE_DB.MoveNext
    
    Loop
Else
    xSql = "select * from rtext " _
        & "where Aid = " & lSWISABWID1 _
        & " and text_s_umidl = " & lSWISABWIDL _
        & " and text_s_umidh  =  " & lSWISABWIDH
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    If Not rsSIDE_DB.EOF Then
        V = rsSIDE_DB("text_data_block"): wText_Data_Block = IIf(IsNull(V), "", V & vbCrLf & ":")
   '_____________________________________________________
         K = InStr(wText_Data_Block, ":108:") + 5
        If K > 5 Then
            K3 = InStr(K, wText_Data_Block, Asc10 & ":")
            If K3 > 0 Then Importation_SAA_198 = Mid$(wText_Data_Block, K, K3 - K)
        End If
   '_____________________________________________________
         K = InStr(wText_Data_Block, ":115:") + 5
        If K > 5 Then
            lK115 = Mid$(wText_Data_Block, K, 1)
        End If
    '_____________________________________________________
    End If
End If


GoTo Exit_sub

'==================================================================
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation_SAA_198"
Exit_sub:

End Function

Private Sub Importation_SAA_SWISABWSTA()
Dim V, X As String, K As Long, K2 As Long, Nb As Integer
Dim xSql As String
Dim xAMJ As String
Dim blnOk As Boolean
Dim wSWISABWSTA As String

On Error GoTo Error_Handler
'SWISABWSTA_appe_date_time
'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABWSTA: 1"): DoEvents
'==========================================================================================================
currentAction = "Importation_SAA_SWISABWSTA => YSWISAB0 "

xAMJ = DSys


'________________________________________________________________________
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
         
'==========================================================================================================

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABWSTA in(' ', '#') order by SWISABSWID"

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    wSWISABWSTA = rsSab("SWISABWSTA")
    xSql = "select * from rMesg , rAppe " _
        & " where rMesg.aid = " & rsSab("SWISABWID1") _
        & " and mesg_s_umidl = " & rsSab("SWISABWIDL") & " and mesg_s_umidh = " & rsSab("SWISABWIDH") _
    & " and rMesg.aid = rAppe.Aid and  mesg_s_umidl = appe_s_umidl and  mesg_s_umidh = appe_s_umidh" _
    & " order by appe_date_time , appe_seq_nbr"

    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    
    Do While Not rsSIDE_DB.EOF
       ' If rsSab("SWISABSWID") = 1301582 Or rsSab("SWISABSWID") = 1314431 Then
       '     Debug.Print
       ' End If
        
        
        Call Importation_SAA_SWISABWSTA_Control(wSWISABWSTA, "rAppe")
        rsSIDE_DB.MoveNext
    Loop
        
    rsSab.MoveNext

Loop
'==========================================================================================================

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABWSTA  in(' ', '#') order by SWISABSWID"

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    wSWISABWSTA = rsSab("SWISABWSTA")
    xSql = "select * from rMesg" _
        & " where rMesg.aid = " & rsSab("SWISABWID1") _
        & " and mesg_s_umidl = " & rsSab("SWISABWIDL") & " and mesg_s_umidh = " & rsSab("SWISABWIDH")

    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    
    Do While Not rsSIDE_DB.EOF
        Call Importation_SAA_SWISABWSTA_Control(wSWISABWSTA, "rMesg")
        rsSIDE_DB.MoveNext
    Loop
        
    rsSab.MoveNext

Loop

'________________________________________________________________________
        Call lstErr_AddItem(lstErr, cmdContext, "SWISABWSTA_appe_date_time :" & SWISABWSTA_appe_date_time): DoEvents
'________________________________________________________________________

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABWSTA : 5 TRANSACTION fin"): DoEvents
'________________________________________________________________________

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
'==========================================================================================================
        Call Importation_SAA_SWISABWSTA_rAppe
'==========================================================================================================
    End If


'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABWSTA : 6 exit"): DoEvents
'________________________________________________________________________

End Sub
Private Sub Importation_SAA_SWISABWSTA_rAppe()
Dim V, X As String, K As Long, K2 As Long, Nb As Integer
Dim xSql As String
Dim xAMJ As String
Dim blnOk As Boolean
Dim wSWISABWSTA As String

On Error GoTo Error_Handler
'SWISABWSTA_appe_date_time
'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABWSTA: 1"): DoEvents
'==========================================================================================================
currentAction = "Importation_SAA_SWISABWSTA => YSWISAB0 "

xAMJ = DSys

If SWISABWSTA_appe_date_time = "" Then
    
    Do
        Call lstErr_AddItem(lstErr, cmdContext, "SWISABWSTA_appe_date_time (init):" & xAMJ): DoEvents
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
             & " where SWISABWAMJ = " & xAMJ & " and SWISABWSTA <> ' ' and SWISABWES = 'S' order by SWISABSWID desc "
        Set rsSab = cnsab.Execute(xSql)
        If Not rsSab.EOF Then
            xAMJ = rsSab("SWISABWAMJ")
            SWISABWSTA_appe_date_time = Mid$(xAMJ, 7, 2) & "/" & Mid$(xAMJ, 5, 2) & "/" & Mid$(xAMJ, 1, 4) & " 00:00:00"
            blnOk = True
        Else
            xAMJ = dateElp("Jour", -1, xAMJ)
        End If
        
    Loop Until blnOk


    xSql = "select * from rAppe " _
        & " where aid = " & rsSab("SWISABWID1") _
        & " and appe_s_umidl = " & rsSab("SWISABWIDL") & " and appe_s_umidH = " & rsSab("SWISABWIDH") _
        & " and appe_inst_num = 0 order by appe_seq_nbr"
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    
    Do While Not rsSIDE_DB.EOF
        Select Case rsSIDE_DB("appe_network_delivery_status")
            Case "DLV_NACKED": SWISABWSTA_appe_date_time = rsSIDE_DB("appe_date_time"): Exit Do
            Case "DLV_ACKED": SWISABWSTA_appe_date_time = rsSIDE_DB("appe_date_time"): Exit Do
        End Select
        rsSIDE_DB.MoveNext
    Loop

End If

'________________________________________________________________________ ' and appe_inst_num = 0
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox


xSql = "select * from rMesg , rAppe " _
    & "where appe_date_time >= '" & SWISABWSTA_appe_date_time & "'" _
    & " and rMesg.aid = rAppe.Aid and  mesg_s_umidl = appe_s_umidl and  mesg_s_umidh = appe_s_umidh" _
    & " and substring(mesg_uumid,2,4) <> 'XXXX'" _
    & " order by appe_date_time , appe_seq_nbr"
    
Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
If Not rsSIDE_DB.EOF Then

    Do While Not rsSIDE_DB.EOF
        SWISABWSTA_appe_date_time = rsSIDE_DB("appe_date_time")
        If Not IsNull(rsSIDE_DB("mesg_type")) Then
            X = rsSIDE_DB("mesg_type")
        Else
            Call Importation_SAA_SWISABWSTA_Control("", "")
        End If
        If X < "100" Or X = "198" Or X = "298" Or X = "960" Or X = "961" Or X = "962" Or X = "963" Or X = "964" Then
        Else
            Call Importation_SAA_SWISABWSTA_Control("", "rAppe")
        End If
        
        rsSIDE_DB.MoveNext
    Loop
End If

         
'________________________________________________________________________
        Call lstErr_AddItem(lstErr, cmdContext, "SWISABWSTA_appe_date_time :" & SWISABWSTA_appe_date_time): DoEvents
'________________________________________________________________________

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABWSTA : 5 TRANSACTION fin"): DoEvents
'________________________________________________________________________

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABWSTA : 6 exit"): DoEvents
'________________________________________________________________________


End Sub


Private Sub Importation_SAA_Alerte()
Dim V, X As String, X2 As String, K As Long, K2 As Long, Nb As Integer
Dim xSql As String
Dim xAMJ As String
Dim blnOk As Boolean, blnAlerte As Boolean
Dim intv_oper_nickname As String, curX As Currency
'Dim maxSAA_Alerte_Amount As String, maxSAA_Alerte_Approval As String, maxSAA_Alerte_Routage As String
Dim wSAA_Date As String
On Error GoTo Error_Handler
'SAA_Alerte_Amount
'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_Alerte: 1"): DoEvents
'==========================================================================================================
currentAction = "Importation_SAA_Alerte => YSWISAB0 "

xAMJ = YBIATAB0_DATE_CPT_J '"20111001" 'DSys
If SAA_Alerte_Amount = "" Then

    Importation_SAA_Alerte_Init

    newSAA_YBIATAB0.BIATABID = "SAA_Alerte"
    newSAA_YBIATAB0.BIATABK1 = "Sécurité"
    newSAA_YBIATAB0.BIATABK2 = ""
    newSAA_YBIATAB0.BIATABTXT = ""
    
    Call lstErr_AddItem(lstErr, cmdContext, "SAA_Alerte_Amount (init):" & xAMJ): DoEvents
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
         & " where BIATABID = '" & newSAA_YBIATAB0.BIATABID & "' and BIATABK1 = '" & newSAA_YBIATAB0.BIATABK1 & "'"
    Set rsSab = cnsab.Execute(xSql)
    If rsSab.EOF Then
        
        SAA_Alerte_Amount = Mid$(xAMJ, 7, 2) & "/" & Mid$(xAMJ, 5, 2) & "/" & Mid$(xAMJ, 1, 4) & " 00:00:00"
        blnOk = True
        newSAA_YBIATAB0.BIATABTXT = SAA_Alerte_Amount
        New_YBIATAB0 = newSAA_YBIATAB0
        Parametrage_New

    Else
        SAA_Alerte_Amount = Mid$(rsSab("BIATABTXT"), 1, 19)
        SAA_Alerte_Approval = Mid$(rsSab("BIATABTXT"), 21, 19)
        If Trim(SAA_Alerte_Approval) = "" Then SAA_Alerte_Approval = SAA_Alerte_Amount
        SAA_Alerte_Routage = Mid$(rsSab("BIATABTXT"), 41, 19)
        If Trim(SAA_Alerte_Routage) = "" Then SAA_Alerte_Routage = SAA_Alerte_Amount
    End If
    
    X = dateElp("Jour", -31, DSys)
    minSAA_Alerte_Approval = Mid$(X, 1, 4) & "-" & Mid$(X, 5, 2) & "-" & Mid$(X, 7, 2) & " 00:00:00.000"

    Call sqlYBIATAB0_Read("PDC", "USD", YBIATAB0_DATE_CPT_J, X)
    If IsNumeric(Mid$(X, 9, 15)) Then
        SAA_Alerte_cours_USD = CDbl(Mid$(X, 9, 15) / 1000000000)
    Else
        SAA_Alerte_cours_USD = 1
    End If
End If
'===================================

blnAlerte = False


If DSys_S <> Mid$(SAA_Alerte_Amount, 1, 10) Then blnAlerte = True
'=====================================================================================================
xSql = "select * from rMesg" _
        & " where mesg_crea_date_time > '" & SAA_Alerte_Amount & "'" _
        & " order by mesg_crea_date_time desc"

Set rsSIDE_Loop = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_Loop.EOF Then
    wSAA_Date = rsSIDE_Loop("mesg_crea_date_time")
Else
    wSAA_Date = SAA_Alerte_Amount
End If

xSql = "select * from rMesg , rInst  " _
          & " where x_last_emi_appe_date_time >  '" & SAA_Alerte_Routage & "' " _
          & " and mesg_sub_format = 'INPUT'" _
          & " and mesg_crea_rp_name = '_MP_creation'" _
        & " and rInst.Aid = rMesg.aid" _
        & " and Inst_s_umidl = mesg_s_umidl" _
        & " and Inst_s_umidh  = mesg_s_umidh" _
        & " and Inst_num  =  0" _
        & " and inst_mpfn_name = '_SI_to_SWIFT'" _
        & " and inst_auth_oper_nickname Is Null order by x_last_emi_appe_date_time"
Set rsSIDE_Loop = cnSIDE_DB.Execute(xSql)

If rsSIDE_Loop.EOF Then
    SAA_Alerte_Routage = wSAA_Date
Else
    Do While Not rsSIDE_Loop.EOF
            blnAlerte = True
            Call srvrMesg_GetBuffer_ODBC(rsSIDE_Loop, xrMesg)
            X = "Message SWIFT sortant : émis sans autorisation SAA"
            X2 = "Message émis SANS AUTORISATION SAA - créé par " & xrMesg.mesg_crea_oper_nickname & " le " & xrMesg.mesg_crea_date_time
            Call cmdSendMail_SAA_Alerte_rMesg("SAA_Routage", X, X2, "S12" & ";" & "S42", rsSIDE_Loop("x_inst0_unit_name"))
        SAA_Alerte_Routage = rsSIDE_Loop("x_last_emi_appe_date_time")
        rsSIDE_Loop.MoveNext
    Loop
End If
'________________________________________________________________________

'=====================================================================================================



xSql = "select * from rMesg , rIntv  " _
          & " where Mesg_crea_date_time >= {ts '" & minSAA_Alerte_Approval & "'} " _
          & " and intv_Date_Time >  '" & SAA_Alerte_Approval & "' " _
          & " and mesg_sub_format = 'INPUT'" _
        & " and rIntv.Aid = rMesg.aid" _
        & " and Intv_s_umidl = mesg_s_umidl" _
        & " and Intv_s_umidh  = mesg_s_umidh" _
        & " and Intv_inst_num  =  0" _
        & " and substring(Intv_merged_text,1,55) =  'Routed from rp [_MP_authorisation] to rp [_SI_to_SWIFT]' order by intv_date_time"
Set rsSIDE_Loop = cnSIDE_DB.Execute(xSql)

If rsSIDE_Loop.EOF Then
    SAA_Alerte_Approval = wSAA_Date
Else
    Do While Not rsSIDE_Loop.EOF
        intv_oper_nickname = rsSIDE_Loop("intv_oper_nickname")
        blnOk = False
        For K = 1 To arrSAA_Usr_Nb
            If intv_oper_nickname = arrSAA_Usr_Id(K) Then
                blnOk = True
                
                If IsNull(rsSIDE_Loop("x_fin_amount")) Then
                    curX = 0
                Else
                    curX = CCur(rsSIDE_Loop("x_fin_amount"))
                    Select Case rsSIDE_Loop("x_fin_ccy")
                        Case "EUR"
                        Case "USD":
                            curX = curX / SAA_Alerte_cours_USD
                        Case Else
                            Call sqlYBIATAB0_Read("PDC", rsSIDE_Loop("x_fin_ccy"), xAMJ, X)
                            If IsNumeric(Mid$(X, 9, 15)) Then curX = curX / CDbl(Mid$(X, 9, 15) / 1000000000)
    
                   End Select
                End If
                If curX > arrSAA_Usr_MTD(K) Then blnAlerte = True: cmdSendMail_SAA_Alerte_Approval "MTD"
                Exit For
            End If
        Next K
        If Not blnOk Then blnAlerte = True: cmdSendMail_SAA_Alerte_Approval "USR"
        
        SAA_Alerte_Approval = rsSIDE_Loop("intv_date_time")
        rsSIDE_Loop.MoveNext
    Loop
End If
'________________________________________________________________________

'=====================================================================================================
SAA_Alerte_Amount:

xSql = "select * from rMesg" _
        & " where mesg_crea_date_time > '" & SAA_Alerte_Amount & "'" _
        & " and mesg_sub_format = 'INPUT'" _
        & " and mesg_type in ('103','202') order by mesg_crea_date_time"

Set rsSIDE_Loop = cnSIDE_DB.Execute(xSql)
If rsSIDE_Loop.EOF Then
    SAA_Alerte_Amount = wSAA_Date
Else

    Do While Not rsSIDE_Loop.EOF
        blnOk = False
        If IsNull(rsSIDE_Loop("x_fin_amount")) Then
            curX = 0
        Else
            curX = CCur(rsSIDE_Loop("x_fin_amount"))
            Select Case rsSIDE_Loop("x_fin_ccy")
                Case "EUR"
                Case "USD":
                    curX = curX / SAA_Alerte_cours_USD
                Case Else
                    Call sqlYBIATAB0_Read("PDC", rsSIDE_Loop("x_fin_ccy"), xAMJ, X)
                    If IsNumeric(Mid$(X, 9, 15)) Then curX = curX / CDbl(Mid$(X, 9, 15) / 1000000000)
    
           End Select
        End If
        Select Case rsSIDE_Loop("mesg_type")
            Case "103":
               If curX > curSAA_103_EUR Then blnAlerte = True: cmdSendMail_SAA_Alerte_Amount "103"
            Case "202":
                If Trim(rsSIDE_Loop("x_inst0_unit_name")) = "BOTC" Then
                    If curX > curSAA_202_BOTC_EUR Then blnAlerte = True: cmdSendMail_SAA_Alerte_Amount "202_BOTC"
                Else
                    If curX > curSAA_202_EUR Then blnAlerte = True: cmdSendMail_SAA_Alerte_Amount "202"
                End If
                
          End Select
          
        SAA_Alerte_Amount = rsSIDE_Loop("mesg_crea_date_time")
        
        rsSIDE_Loop.MoveNext
    Loop
End If

'==========================================================================================================

If blnAlerte Then
    Old_YBIATAB0 = newSAA_YBIATAB0
    Mid$(newSAA_YBIATAB0.BIATABTXT, 1, 19) = SAA_Alerte_Amount
    Mid$(newSAA_YBIATAB0.BIATABTXT, 21, 19) = SAA_Alerte_Approval
    Mid$(newSAA_YBIATAB0.BIATABTXT, 41, 19) = SAA_Alerte_Routage
    New_YBIATAB0 = newSAA_YBIATAB0
    Parametrage_Update
    Importation_SAA_Alerte_Init
End If
'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "SAA_Alerte_Amount :" & SAA_Alerte_Amount): DoEvents
'________________________________________________________________________

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_Alerte : 5 TRANSACTION fin"): DoEvents
'________________________________________________________________________

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

    
'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_Alerte : 6 exit"): DoEvents
'________________________________________________________________________

End Sub

Private Sub Importation_SAA_Reprise(lAMJ As String)
Dim V, X As String, K As Long, K2 As Long, Nb As Integer, X2 As String, xWhere As String
Dim xSql As String
Dim xAMJ As String, xHMS As String, xUUMID As String
Dim xYSWISAB0 As typeYSWISAB0, xYSWILNK0 As typeYSWILNK0
Dim arrYSWISAB0(5000) As typeYSWISAB0
Dim mAMJ_R As Long
Dim xrMesg As typerMesg

On Error GoTo Error_Handler


currentAction = "Importation_SAA => YSWISAB0 "

Call lstErr_AddItem(lstErr, cmdContext, "Importation SAA : " & lAMJ & " > " & importSWISABSWID): DoEvents
X = Mid$(lAMJ, 1, 4) & "-" & Mid$(lAMJ, 5, 2) & "-" & Mid$(lAMJ, 7, 2) & " " & "00:00:00.000"
X2 = Mid$(lAMJ, 1, 4) & "-" & Mid$(lAMJ, 5, 2) & "-" & Mid$(lAMJ, 7, 2) & " " & "23:59:59.000"
xWhere = " and Mesg_crea_date_time <= {ts '" & X2 & "'} "


 xSql = "select * from rMesg " _
           & " where Mesg_crea_date_time >= {ts '" & X & "'}" & xWhere & " order by Mesg_crea_date_time"

'??????????????????????????????????????????????????????????????????

   
   
Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
  
Do While Not rsSIDE_DB.EOF
    If Not IsNull(rsSIDE_DB("mesg_type")) Then
        X = rsSIDE_DB("mesg_type")
        If X < "100" Or X = "198" Or X = "298" Or X = "960" Or X = "961" Or X = "962" Or X = "963" Or X = "964" Then
        Else
            Call srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
            rsYSWISAB0_Init xYSWISAB0
            xYSWISAB0.SWISABWID1 = xrMesg.Aid
            xYSWISAB0.SWISABWIDL = xrMesg.mesg_s_umidl
            xYSWISAB0.SWISABWIDH = xrMesg.mesg_s_umidh
            
                xYSWISAB0.SWISABWMTK = xrMesg.mesg_type
                xUUMID = xrMesg.mesg_uumid
                If Mid$(xUUMID, 1, 1) = "I" Then
                    xYSWISAB0.SWISABWES = "S"
                    xYSWISAB0.SWISABWN20 = Replace(xrMesg.mesg_trn_ref, "'", " ")
                    xYSWISAB0.SWISABWL20 = Replace(xrMesg.mesg_rel_trn_ref, "'", " ")
                Else
                    xYSWISAB0.SWISABWES = "E"
                    xYSWISAB0.SWISABWL20 = Replace(xrMesg.mesg_trn_ref, "'", " ")
                    xYSWISAB0.SWISABWN20 = Replace(xrMesg.mesg_rel_trn_ref, "'", " ")
                End If
                
                xYSWISAB0.SWISABWBIC = Mid$(xUUMID, 2, 11)
                If Not IsNull(xrMesg.x_fin_ccy) Then
                    xYSWISAB0.SWISABWDEV = xrMesg.x_fin_ccy
                    xYSWISAB0.SWISABWMTD = CCur(xrMesg.x_fin_amount)
                End If
                xYSWISAB0.SWISABWSRV = xrMesg.x_inst0_unit_name
                Call dateJma10_Amj(Mid$(xrMesg.mesg_crea_date_time, 1, 10), X)
                xYSWISAB0.SWISABWAMJ = Val(X)
                X = Mid$(xrMesg.mesg_crea_date_time, 12, 8)
                xYSWISAB0.SWISABWHMS = Val(Mid$(X, 1, 2) & Mid$(X, 4, 2) & Mid$(X, 7, 2))
                
                Call dateJma10_Amj(Mid$(xrMesg.last_update, 1, 10), X)
                xYSWISAB0.SWISABXAMJ = Val(X)
                X = Mid$(xrMesg.last_update, 12, 8)
                xYSWISAB0.SWISABXHMS = Val(Mid$(X, 1, 2) & Mid$(X, 4, 2) & Mid$(X, 7, 2))
                
                If Not IsNull(xrMesg.mesg_possible_dup_creation) Then
                    If Trim(xrMesg.mesg_possible_dup_creation) <> "" Then
                        xYSWISAB0.SWISABKPDE = "X"
                    End If
                End If
                
'$JPL 20110720 ________________________________________________________________
                ' xYSWISAB0.SWISABKSRV = SAA_2_BIA_Unit(xYSWISAB0.SWISABWSRV)
                '$JPL 20120611  ________________________________________________________________
               Select Case xYSWISAB0.SWISABWSRV
                    Case "SOBF", "ORPA": xYSWISAB0.SWISABKSRV = "S01"
                    Case "SOBI": xYSWISAB0.SWISABKSRV = "S10"
                    Case "DAFI", "BOTC": xYSWISAB0.SWISABKSRV = "S32"
                    Case "DCOM": xYSWISAB0.SWISABKSRV = "S41"
                    Case Else: xYSWISAB0.SWISABKSRV = "S00"
                End Select
                
                If xYSWISAB0.SWISABWES = "S" Then
                
                    ''If Trim(xrMesg.mesg_status) = "COMPLETED" Then xYSWISAB0.SWISABWSTA = "V"
               Else
                    'If Mid$(xYSWISAB0.SWISABWMTK, 2, 2) = "99" Then
                    If Mid$(xYSWISAB0.SWISABWMTK, 2, 1) = "9" Then      '$JPL 20121018
                        xYSWISAB0.SWISABK999 = "!"
                    Else
                        If xYSWISAB0.SWISABKSRV = "S00" Then
                            Select Case Mid$(xYSWISAB0.SWISABWMTK, 1, 1)
                                Case "7": xYSWISAB0.SWISABKSRV = "S10"
                                Case "3", "4", "5": xYSWISAB0.SWISABKSRV = "S32"
                            End Select
                        End If
                    End If
    
                End If
'$JPL 20110720 ________________________________________________________________
                If xYSWISAB0.SWISABWMTK = "950" Then
                    xYSWISAB0.SWISABKSRV = "S11"
                    xYSWISAB0.SWISABSER = "00"
                    xYSWISAB0.SWISABSSE = "00"
                    xYSWISAB0.SWISABOPEC = "XXX"
                    xYSWISAB0.SWISABOPEN = -1
                    ''xYSWISAB0.SWISABWSTA = "V"
                End If
                
                Nb = Nb + 1
                'xYSWISAB0.SWISABSWID = Nb + importSWISABSWID
                importSWISABSWID = importSWISABSWID + 1
                xYSWISAB0.SWISABSWID = importSWISABSWID
                arrYSWISAB0(Nb) = xYSWISAB0
                If Nb = 5000 Then Exit Do
'JPL 20110404
            'End If
        End If
    End If
    rsSIDE_DB.MoveNext
    DoEvents

Loop

'Call lstErr_AddItem(lstErr, cmdContext, "SAA : " & Nb & " > " & xYSWISAB0.SWISABSWID): DoEvents
         

'________________________________________________________________________

'Call lstErr_AddItem(lstErr, cmdContext, "Importation SAA : " & Nb & " > " & xYSWISAB0.SWISABSWID): DoEvents

'________________________________________________________________________
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

For K = 1 To Nb
    If arrYSWISAB0(K).SWISABWSRV = "None" Then arrYSWISAB0(K).SWISABWSRV = ""
'$jpl==============================================
    'Select Case Trim(arrYSWISAB0(K).SWISABWN20)
    '    Case "None", "NONREF": arrYSWISAB0(K).SWISABWN20 = ""
    'End Select
    'Select Case Trim(arrYSWISAB0(K).SWISABWL20)
    '    Case "None", "NONREF": arrYSWISAB0(K).SWISABWL20 = ""
    'End Select
'$jpl==============================================
    V = sqlYSWISAB0_Insert(arrYSWISAB0(K))
    If Not IsNull(V) Then GoTo Exit_sub
    
        
Next K

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If

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
fgDetail_32B = "": fgDetail_33B = "": fgDetail_36 = ""
fgDetail_30V = "": fgDetail_37G = "": fgDetail_34E = ""
fgDetail_30T = "": fgDetail_30P = "": fgDetail_82A = "": fgDetail_87A = ""
fgDetail_22C = ""
fgDetail_70 = "": fgDetail_72 = ""

Call rsYSWISAB0_Init(oldYSWISAB0) '$JPL 2014-11-25

If cmdSelect_SQL_K = "1trf" Then
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0," & paramIBM_Library_SABSPE & ".YSWISAB1 where SWISABWID1 = " & Mesg_aid _
    & " and SWISABWIDL = " & mesg_s_umidl _
    & " and SWISABWIDH = " & mesg_s_umidh _
    & " and SWISAB1ID = SWISABSWID"

Else
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABWID1 = " & Mesg_aid _
    & " and SWISABWIDL = " & mesg_s_umidl _
    & " and SWISABWIDH = " & mesg_s_umidh
    
'$JPL 2013-11-06 ____________________________________________________________________________________________________
    If oldYGOSDOS0.GOSDOSWMTK = "BIA" Then
        X = "select SWILNKSWID from " & paramIBM_Library_SABSPE & ".YSWILNK0 where SWILNKAPPC = 'GOS'" _
        & " and SWILNKAPPN = " & oldYGOSDOS0.GOSDOSIDD _
        & " order by SWILNKSWID FETCH FIRST 1 ROWS ONLY"
        Set rsSab = cnsab.Execute(X)
        If Not rsSab.EOF Then
            X = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABSWID = " & rsSab("SWILNKSWID")
            Set rsSab = cnsab.Execute(X)
            If Not rsSab.EOF Then
                Mesg_aid = rsSab("SWISABWID1")
                mesg_s_umidl = rsSab("SWISABWIDL")
                mesg_s_umidh = rsSab("SWISABWIDH")
                xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABWID1 = " & Mesg_aid _
                & " and SWISABWIDL = " & mesg_s_umidl _
                & " and SWISABWIDH = " & mesg_s_umidh
            End If
            
        End If
    End If
'$JPL 2013-11-06 ____________________________________________________________________________________________________
End If

Set rsSab = cnsab.Execute(xSql)
If rsSab.EOF Then

    If Mid$(cmdSelect_SQL_K, 1, 1) = "1" Or Mid$(cmdSelect_SQL_K, 1, 1) = "J" Or Mid$(cmdSelect_SQL_K, 1, 1) = "6" Or Mid$(cmdSelect_SQL_K, 1, 1) = "7" Then
        Call fgSwift_Display(0)
    Else
        tabDetail.Tab = 0
        If oldYGOSDOS0.GOSDOSIDD = 0 Then
            tabDetail.Caption = "Création d'un dossier"
        Else
            tabDetail.Caption = "Dossier  n° " & oldYGOSDOS0.GOSDOSIDD & " : " & "créé par le service  " & Trim(arrService_Lib(Mid$(oldYGOSDOS0.GOSDOSISRV, 2, 2)))
        End If
    End If
Else
    Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
    If Mid$(cmdSelect_SQL_K, 1, 1) = "1" Or Mid$(cmdSelect_SQL_K, 1, 1) = "J" Or Mid$(cmdSelect_SQL_K, 1, 1) = "6" Or Mid$(cmdSelect_SQL_K, 1, 1) = "7" Then
        fgSwift_Display rsSab("SWISABSWID")
    Else
        If oldYSWISAB0.SWISABWES = "E" Then
            X = "reçu de "
            wColor = RGB(190, 240, 255)
            wColorFixed = vbBlue
        Else
            X = "sortant vers "
            wColor = RGB(220, 255, 220)
            wColorFixed = RGB(0, 64, 0)
        End If
        libDetail_SWISABSWID = "SAB : " & Trim(oldYSWISAB0.SWISABOPEC) & " " & Format(oldYSWISAB0.SWISABOPEN, "### ###")
        
 
   
        libDetail_SWISABSWID.BackColor = wColor
        lblGOSDOSSTAD.BackColor = wColor
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
        If oldYGOSDOS0.GOSDOSIDD = 0 Then
            tabDetail.Caption = "Création d'un dossier"
        Else
            tabDetail.Caption = "Dossier  n° " & oldYGOSDOS0.GOSDOSIDD & " : " & "créé par le service  " & Trim(arrService_Lib(Mid$(oldYGOSDOS0.GOSDOSISRV, 2, 2)))
        End If
        fgDetail.Visible = True
    End If
End If
If blnYGOSDOS0_New Then
    cmdSelect_SQL_Kbis = "3"
Else
    cmdSelect_SQL_Kbis = cmdSelect_SQL_K
End If

Select Case cmdSelect_SQL_Kbis
    Case "1", "1Live_Entran", "1Live_Sortan", "1b": 'fraDetail.Visible = True
    Case "1?", "1?*": fraSWISABKSRV_Display

    Case "2", "2-RAM": fraDetail_LAB_Init
    Case "3", "4", "4 Journal":
        lblGOSDOSIAMJ = "créé le " & dateImp10_S(oldYGOSDOS0.GOSDOSIAMJ) & "   par " & arrService_Lib(Mid$(oldYGOSDOS0.GOSDOSISRV, 2, 2))
        lblGOSDOSUAMJ = "màj  le " & dateImp10_S(oldYGOSDOS0.GOSDOSUAMJ) & " à " & timeImp8(oldYGOSDOS0.GOSDOSUHMS) & "   par " & Trim(oldYGOSDOS0.GOSDOSUUSR) & " - " & arrService_Lib(Mid$(oldYGOSDOS0.GOSDOSUSRV, 2, 2))
        fraDetail_LAB_Display
        fraDetail_LAB.Visible = True
        For I = 1 To 10
            fgDetail.Row = Val(Mid$(oldYGOSDOS0.GOSDOSITOP, I * 2 - 1, 2))
            If fgDetail.Row > 0 And fgDetail.Row < fgDetail.Rows Then fgDetail.Col = 1: fgDetail.CellForeColor = vbRed: fgDetail.CellFontBold = True
        Next I
        'X = "GOS n° " & oldYGOSDOS0.GOSDOSIDD & " : "
        X = ""
         Select Case oldYGOSDOS0.GOSDOSSTAG
            Case "V":  X = "Validé, ": fraDetail_Y.BackColor = &H8000&     '&H40FF40
            Case "R":  X = "Rejeté, ": fraDetail_Y.BackColor = &HFF&       '&HFFA0FF
        End Select
        Select Case oldYGOSDOS0.GOSDOSSTAD
            Case " ": lblGOSDOSSTAD = X & " en gestion au service  " & Trim(arrService_Lib(Mid$(oldYGOSDOS0.GOSDOSGSRV, 2, 2)))
                                    
                        fraDetail_Y.BackColor = RGB(255, 255, 190) '&H808000    '&HC0FFC0
            Case "A": lblGOSDOSSTAD = X & " Annulé": fraDetail_Y.BackColor = &H404040    '&H808080
            Case "C": lblGOSDOSSTAD = X & " Clôturé": fraDetail_Y.BackColor = &H404040    '&H808080
            Case "x": lblGOSDOSSTAD = X & " à clôturer": fraDetail_Y.BackColor = &H404040    '&H808080
            Case Else: lblGOSDOSSTAD = X & " ????": fraDetail_Y.BackColor = vbMagenta
        End Select
        tabDetail.Caption = lblGOSDOSSTAD & ", Echéance : " & dateImp10(oldYGOSDOS0.GOSDOSECHD)
    
    Case "5": tabDetail.Caption = " Messages PDE & références en double": fraDetail.Visible = True: fraList_Display
    
    Case "9", "9+", "5h":
        tabDetail.Caption = " Messages *99": fraDetail.Visible = True: fraList_Display
        m999_YSWISAB0 = oldYSWISAB0
    
    Case "6", "6E": Call YSWIRAM0_Match_6E
End Select



'___________________________________________________________________________
cmdMail_MT.Visible = True

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
Dim wSWIL As Integer
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
        mMOUVEMSER = oldYSWISAB0.SWISABSER
        mMOUVEMSSE = oldYSWISAB0.SWISABSSE
        mMOUVEMOPE = oldYSWISAB0.SWISABOPEC
        mMOUVEMNUM = oldYSWISAB0.SWISABOPEN

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
         oldYSWISAB0.SWISABWBIC = rsSIDE_DB("mesg_Sender_X1")
         X = rsSIDE_DB("mesg_crea_date_time")
         oldYSWISAB0.SWISABWAMJ = Mid$(X, 7, 4) & Mid$(X, 4, 2) & Mid$(X, 1, 2)
         oldYSWISAB0.SWISABWHMS = Mid$(X, 12, 2) & Mid$(X, 15, 2) & Mid$(X, 18, 2)
    End If
End If
'--------------------------------------------------------------
    If oldYSWISAB0.SWISABWES = "E" Then
        X = "reçu de "
        wColor = RGB(190, 240, 255)
        wColorFixed = vbBlue
    Else
        X = "sortant vers "
        wColor = RGB(220, 255, 220)
        wColorFixed = RGB(0, 64, 0)
    End If
    
    libSWIFT_SWISABSWID = "SAB : " & Trim(oldYSWISAB0.SWISABOPEC) & " " & Format(oldYSWISAB0.SWISABOPEN, "### ###")
    If oldYSWISAB0.SWISABXGOS = "G" Or oldYSWISAB0.SWISABXEVE = "G" Then
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWILNK0 where SWILNKSWID = " & oldYSWISAB0.SWISABSWID & " and SWILNKAPPC = 'GOS' and SWILNKSTA = ''"
        Set rsSab = cnsab.Execute(xSql)
        If Not rsSab.EOF Then libSWIFT_SWISABSWID = libSWIFT_SWISABSWID & "            ===>   GOS : " & rsSab("SWILNKAPPN")
        libSWIFT_SWISABSWID.BackColor = vbYellow
        libSWIFT_SWISABSWID.ForeColor = vbRed
    Else
        libSWIFT_SWISABSWID.BackColor = mColor_GB
        libSWIFT_SWISABSWID.ForeColor = vbWhite
    End If
    
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
            If mSWIECHSWIL = 0 Then
                  fgSwift_DisplayLine fgSwift.Row, wColor, wColorFixed
            Else
                wSWIL = wSWIL + 1
                If mSWIECHSWIL = wSWIL Then
                    fgSwift_DisplayLine fgSwift.Row, wColor, vbRed
                Else
                    fgSwift_DisplayLine fgSwift.Row, wColor, wColorFixed
                End If
            End If
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

If Mid$(cmdSelect_SQL_K, 1, 1) = "1" Or Mid$(cmdSelect_SQL_K, 1, 1) = "J" Or Mid$(cmdSelect_SQL_K, 1, 1) = "6" Or Mid$(cmdSelect_SQL_K, 1, 1) = "7" Then
    Set fraSwift.Container = fraTab0
    fraSwift.Top = fgSelect.Top + fgSelect.RowHeightMin 'fraTab0.Top + 1600
    fraSwift.Left = fraTab0.Width - fraSwift.Width - 100

    If chkSIDE_DB_Show Then Call frmSIDE_DB.fgSwift_Display(lSWISABSWID, Mesg_aid, mesg_s_umidl, mesg_s_umidh)
    
    If mMOUVEMNUM = 0 Then
        chkSAB_Dossier_DB_Show.Enabled = False
    Else
        chkSAB_Dossier_DB_Show.Enabled = True
        If chkSAB_Dossier_DB_Show Then Call frmSAB_Dossier_DB.Form_Init("", "", "", "", mMOUVEMSER, mMOUVEMSSE, mMOUVEMOPE, mMOUVEMNUM)
    End If
Else
    Set fraSwift.Container = fraDetail
    fraSwift.Top = 600
    fraSwift.Left = fraDetail.Width - fraSwift.Width - 500     'fraDetail.Width - 6000
End If


'___________________________________________________________________________

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
    fgSelect.CellForeColor = wColor


fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
fgSelect.Col = 8: fgSelect.Text = rsSIDE_DB("aid")
fgSelect.Col = 9: fgSelect.Text = rsSIDE_DB("mesg_s_umidl")
fgSelect.Col = 10: fgSelect.Text = rsSIDE_DB("mesg_s_umidh")

If Trim(rsSIDE_DB("mesg_status")) = "COMPLETED" Then
    fgSelect.Col = 7: fgSelect.Text = rsSIDE_DB("x_inst0_unit_name") & " - " & Trim(rsSIDE_DB("inst_mpfn_name")) 'xUUMID
    fgSelect.CellForeColor = wColor
Else
    fgSelect.Col = 7: fgSelect.Text = rsSIDE_DB("x_inst0_unit_name") & " : " & Trim(rsSIDE_DB("inst_mpfn_name")) & "  =>  " & Trim(rsSIDE_DB("inst_rp_name"))
    fgSelect.CellForeColor = vbBlue
    fgSelect.CellFontBold = True
    For K = 0 To 10
        fgSelect.Col = K: fgSelect.CellBackColor = wBackColor
    Next K
End If

End Sub
Public Sub fgSelect_DisplayLine_YSWISAB0(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wBackColor As Long
Dim X As String, X2 As String, X3 As String, txtSWISABWSTA As String
On Error Resume Next


If Mid$(cmdSelect_SQL_K, 1, 1) = "6" Then
    Dim blnUnderline As Boolean
    If mYSWIRAM0_Match_XOPE <> rsSab("SWIRAMXOPE") Then
        mYSWIRAM0_Match_XOPE = rsSab("SWIRAMXOPE")
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        If cmdSelect_SQL_K = "6" And rsSab("SWIRAMSTA") = " " And rsSab("SWIRAMX22") <> "2" Then blnUnderline = True
    End If
End If

If rsSab("SWISABWES") = "S" Then
    X = rsSab("SWISABWMTK") & " S"
    wColor = RGB(16, 96, 16)
    wBackColor = mColor_W0
    X3 = rsSab("SWISABWN20")
Else
    X = rsSab("SWISABWMTK") & " E"
    wColor = vbBlue
    wBackColor = mColor_B0
    X3 = rsSab("SWISABWL20")
End If

txtSWISABWSTA = ""
Select Case rsSab("SWISABWSTA")
    Case "V", "X"
    Case " ": For K = 0 To 11: fgSelect.Col = K: fgSelect.CellBackColor = wBackColor: Next K: txtSWISABWSTA = " Live"
    Case "E": For K = 0 To 11: fgSelect.Col = K: fgSelect.CellBackColor = mColor_W0: Next K: txtSWISABWSTA = " Nacked"
    Case "#": For K = 0 To 11: fgSelect.Col = K: fgSelect.CellBackColor = mColor_W1: Next K: txtSWISABWSTA = " Nacked => Live"
    Case "B": For K = 0 To 11: fgSelect.Col = K: fgSelect.CellBackColor = mColor_Y2: Next K: txtSWISABWSTA = " Blocked"
    Case "H": For K = 0 To 11: fgSelect.Col = K: fgSelect.CellBackColor = mColor_Y2: Next K: txtSWISABWSTA = " Hold"
    Case Else: For K = 0 To 11: fgSelect.Col = K: fgSelect.CellBackColor = mColor_Y1: Next K: txtSWISABWSTA = " " & rsSab("SWISABWSTA")
End Select

Select Case cmdSelect_SQL_K
    Case "5": X2 = rsSab("SWISABKPDE")
            If X2 = "!" Or X = "?" Then wColor = vbMagenta
            X2 = rsSab("SWISABK20")
            If X2 = "!" Then wColor = vbMagenta
    Case "9", "9+": X2 = rsSab("SWISABK999")
            If X2 = "!" Then wColor = vbMagenta
End Select

'If rsSab("SWISABXEVE") <> " " Then wColor = vbMagenta
    

fgSelect.Col = 0: fgSelect.Text = X
fgSelect.CellForeColor = wColor

fgSelect.Col = 1: fgSelect.Text = rsSab("SWISABWBIC")
fgSelect.CellForeColor = wColor
fgSelect.Col = 2: fgSelect.Text = X3
fgSelect.CellForeColor = wColor
Select Case rsSab("SWISABK20")
    Case "!": fgSelect.CellBackColor = RGB(220, 220, 255)
    Case Is <> " ": fgSelect.CellBackColor = RGB(220, 255, 220)
End Select

fgSelect.Col = 3: fgSelect.Text = Format$(CCur(rsSab("SWISABWMTD")), "### ### ### ##0.00")
fgSelect.CellForeColor = vbRed
fgSelect.CellFontBold = True
fgSelect.Col = 4: fgSelect.Text = rsSab("SWISABWDEV")
fgSelect.CellForeColor = wColor

'____________________________________________________________________________________________
If cmdSelect_SQL_K = "1trf" Then
    fgSelect.Col = 5
     Select Case rsSab("SWISABW71A")
        Case "O": X = "OUR - "
        Case "S": X = "SHA - "
        Case "B": X = "BEN - "
        Case Else: X = rsSab("SWISABW71A") & " - "
    End Select
   Select Case rsSab("SWISABWEBA")
        Case "E": fgSelect.Text = X & "EBA": fgSelect.CellForeColor = wColor
        Case "T": fgSelect.Text = X & "TGT": fgSelect.CellForeColor = vbMagenta
        Case Else: fgSelect.Text = X: fgSelect.CellForeColor = vbMagenta
    End Select
    
    fgSelect.Col = 6: fgSelect.Text = rsSab("SWISABW50P") & "   " & rsSab("SWISABW50Z") & "   " & rsSab("SWISABW52A")
    fgSelect.CellForeColor = wColor
    fgSelect.Col = 7: fgSelect.Text = rsSab("SWISABW59P") & "   " & rsSab("SWISABW59Z") & "   " & rsSab("SWISABW57A")
    fgSelect.CellForeColor = wColor
Else

    fgSelect.Col = 5
    
    If rsSab("SWISABK20") <> " " Then
        fgSelect.Text = rsSab("SWISABK20") & " réf ="
    Else
        If rsSab("SWISABKPDE") <> " " Then
            fgSelect.Text = rsSab("SWISABKPDE") & " PDE"
            Select Case rsSab("SWISABK20")
                Case "!", "?": fgSelect.CellBackColor = RGB(255, 220, 220)
                Case Else: fgSelect.CellBackColor = RGB(220, 255, 220)
            End Select
        Else
            If rsSab("SWISABK999") <> " " Then fgSelect.Text = rsSab("SWISABK999") & "-9-" '$JPL 20121018 " %99"
        End If
    End If
    
    X = rsSab("SWISABXGOS") & rsSab("SWISABXEVE")
    If X <> "  " Then
        fgSelect.Text = fgSelect.Text & "  " & X
        fgSelect.CellBackColor = mColor_Y1
    End If
    fgSelect.CellForeColor = wColor
    fgSelect.Col = 6: fgSelect.Text = dateImp_Amj(rsSab("SWISABWAMJ")) & "   " & timeImp8(rsSab("SWISABWHMS")) & txtSWISABWSTA
    fgSelect.CellForeColor = RGB(80, 80, 80)
    
    fgSelect.Col = 7
    X = ""
    If rsSab("SWISABXGOS") <> " " Then
        X = " # GOS"
        fgSelect.CellBackColor = vbYellow

    Else
        Select Case rsSab("SWISABXEVE")
            Case " ", "="
            Case "G": X = " # EVE": fgSelect.CellBackColor = RGB(245, 222, 131)
            Case "*":
                
                If rsSab("SWISABK999") = "I" Then
                    fgSelect.CellBackColor = RGB(220, 220, 220)
                'Else
                     'fgSelect.CellBackColor = mColor_B0
               End If
                
            Case Else: X = " ???": fgSelect.CellBackColor = mColor_W1
        End Select
    End If
    
    K = Val(Mid$(rsSab("SWISABKSRV"), 2, 2))
    If rsSab("SWISABOPEN") = 0 Then
        fgSelect.Text = arrService_Lib(K) & X
    Else
        fgSelect.Text = arrService_Lib(K) & " - " & rsSab("SWISABSER") & " " & rsSab("SWISABSSE") & " " & rsSab("SWISABOPEC") & " " & rsSab("SWISABOPEN") & X
    End If
    'fgSelect.Text = rsSab("SWISABWN20")
    fgSelect.CellForeColor = wColor
End If

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
fgSelect.Col = 8: fgSelect.Text = rsSab("SWISABWID")
fgSelect.Col = 9: fgSelect.Text = rsSab("SWISABWIDL")
fgSelect.Col = 10: fgSelect.Text = rsSab("SWISABWIDH")
fgSelect.Col = 11: fgSelect.Text = rsSab("SWISABSWID")

If Mid$(cmdSelect_SQL_K, 1, 1) = "6" Then
   ' Dim blnUnderline As Boolean
   ' If mYSWIRAM0_Match_XOPE <> rsSab("SWIRAMXOPE") Then mYSWIRAM0_Match_XOPE = rsSab("SWIRAMXOPE"): blnUnderline = True

    fgSelect.Col = 6: fgSelect.Text = dateImp_Amj(rsSab("SWIRAMYAMJ")) & "   " & timeImp8(rsSab("SWIRAMYHMS")) & txtSWISABWSTA
    fgSelect.Col = 2: fgSelect.Text = rsSab("SWIRAMXREF")
    Select Case rsSab("SWIRAMSTA")
        Case " ":
            
             If rsSab("SWIRAMXES") = "E" Then
                wBackColor = mColor_G2
            Else
                wBackColor = mColor_G1 'RGB(255, 230, 255)
            End If
            
        Case "#":
                'If cmdSelect_SQL_K = "6" Then
                    If rsSab("SWIRAMXES") = "E" Then
                        wBackColor = mColor_Y2
                    Else
                        wBackColor = mColor_Y1 'RGB(255, 230, 255)
                    End If
                'End If
        Case "I":
             If rsSab("SWIRAMYUPD") = " " Then
                wBackColor = RGB(220, 220, 220) 'vbWhite
            Else
                wBackColor = RGB(200, 200, 200)
            End If
       Case "A": wBackColor = RGB(180, 180, 180)
       Case Else: wBackColor = mColor_B0
    End Select
    Select Case rsSab("SWIRAMX22")
        Case "1": X = "CONF"
        Case "2": X = "MATU": If rsSab("SWIRAMSTA") = " " Then wBackColor = mColor_G0
        Case "3": X = "ROLL"
        Case "4": X = "AMND"
        Case "5": X = "CANC"
        Case Else: X = "???"
    End Select
    fgSelect.Col = 5: fgSelect.Text = X & " - " & rsSab("SWIRAMSTA") & " " & rsSab("SWIRAMYUPD")

    For K = 0 To fgSelect_arrIndex
        fgSelect.Col = K
        fgSelect.CellBackColor = wBackColor
        fgSelect.CellFontBold = blnUnderline
    Next K
    If rsSab("SWIRAMYUPD") = "M" Then fgSelect.Col = 5: fgSelect.CellBackColor = mColor_Y2
End If

End Sub

Public Sub fgSelect_DisplayLine_YGOSDOS0(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wBackColor As Long, wColor_Row As Long
Dim xSql As String, xUUMID As String
On Error Resume Next

Call fgSelect_DisplayLine_YGOSDOS0_Color(wColor, wBackColor)

fgSelect.Col = 0: fgSelect.Text = xYGOSDOS0.GOSDOSIDD & "-" & xYGOSDOS0.GOSDOSWMTK & " " & xYGOSDOS0.GOSDOSWES
fgSelect.CellForeColor = wColor
fgSelect.Col = 1: fgSelect.Text = xYGOSDOS0.GOSDOSWBIC
fgSelect.CellForeColor = wColor
fgSelect.Col = 2: fgSelect.Text = xYGOSDOS0.GOSDOSWTRN
fgSelect.CellForeColor = wColor
fgSelect.Col = 3: fgSelect.Text = Format$(xYGOSDOS0.GOSDOSWMTD, "### ### ### ##0.00")
If xYGOSDOS0.GOSDOSSTAD = " " Then
    fgSelect.CellForeColor = vbRed
    fgSelect.CellFontBold = True
Else
    fgSelect.CellForeColor = wColor
End If
fgSelect.Col = 4: fgSelect.Text = xYGOSDOS0.GOSDOSWDEV
fgSelect.CellForeColor = wColor

fgSelect.Col = 5: fgSelect.Text = xYGOSDOS0.GOSDOSSTAG & "-" & xYGOSDOS0.GOSDOSSTAD
fgSelect.CellForeColor = wColor

fgSelect.Col = 6: fgSelect.Text = dateImp10_S(xYGOSDOS0.GOSDOSECHD)
fgSelect.CellForeColor = wColor

fgSelect.Col = 7
K = Val(Mid$(xYGOSDOS0.GOSDOSISRV, 2, 2)): fgSelect.Text = arrService_Lib(K)
fgSelect.CellForeColor = mColor_GB
fgSelect.CellFontBold = True
fgSelect.Col = 8
K = Val(Mid$(xYGOSDOS0.GOSDOSGSRV, 2, 2)): fgSelect.Text = arrService_Lib(K)
fgSelect.CellForeColor = vbRed
fgSelect.CellFontBold = True
fgSelect.Col = 9
K = Val(Mid$(xYGOSDOS0.GOSDOSRCOM, 2, 2)): fgSelect.Text = arrRCOM_Lib(K)
fgSelect.CellForeColor = wColor
fgSelect.Col = 10: fgSelect.Text = xYGOSDOS0.GOSDOSCLI
fgSelect.CellForeColor = wColor

fgSelect.Col = 11: fgSelect.Text = xYGOSDOS0.GOSDOSLABK
fgSelect.CellForeColor = wColor

For K = 0 To fgSelect_arrIndex
    fgSelect.Col = K
    fgSelect.CellBackColor = wBackColor

Next K
'If cmdSelect_SQL_K = "4" Then
fgSelect.Col = 6
    If xYGOSDOS0.GOSDOSSTAD = " " Or xYGOSDOS0.GOSDOSSTAD = "x" Then
        fgSelect.CellBackColor = RGB(255, 208, 128)
        If xYGOSDOS0.GOSDOSECHD <= DSys Then fgSelect.CellBackColor = mColor_W0
    End If
    
'End If

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
End Sub
Public Sub fgSelect_DisplayLine_Echéancier(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wBackColor As Long, wColor_Row As Long
Dim xSql As String, xUUMID As String
On Error Resume Next

Call fgSelect_DisplayLine_YGOSDOS0_Color(wColor, wBackColor)

fgSelect.Col = 0: fgSelect.Text = xYGOSDOS0.GOSDOSIDD & "-" & xYGOSDOS0.GOSDOSWMTK & " " & xYGOSDOS0.GOSDOSWES
fgSelect.CellForeColor = wColor
fgSelect.Col = 1: fgSelect.Text = xYGOSDOS0.GOSDOSWBIC
fgSelect.CellForeColor = wColor
fgSelect.Col = 2: fgSelect.Text = xYGOSDOS0.GOSDOSWTRN
fgSelect.CellForeColor = wColor
fgSelect.Col = 3: fgSelect.Text = Format$(xYGOSDOS0.GOSDOSWMTD, "### ### ### ##0.00")
If xYGOSDOS0.GOSDOSSTAD = " " Then
    fgSelect.CellForeColor = vbRed
    fgSelect.CellFontBold = True
Else
    fgSelect.CellForeColor = wColor
End If
fgSelect.Col = 4: fgSelect.Text = xYGOSDOS0.GOSDOSWDEV
fgSelect.CellForeColor = wColor

fgSelect.Col = 5: fgSelect.Text = xYGOSDOS0.GOSDOSSTAG & "-" & xYGOSDOS0.GOSDOSSTAD
fgSelect.CellForeColor = wColor

fgSelect.Col = 6: fgSelect.Text = dateImp10_S(xYGOSDOS0.GOSDOSECHD)
fgSelect.CellForeColor = wColor

fgSelect.Col = 7
K = Val(Mid$(xYGOSDOS0.GOSDOSISRV, 2, 2)): fgSelect.Text = arrService_Lib(K)
fgSelect.CellForeColor = mColor_GB
fgSelect.CellFontBold = True
fgSelect.Col = 8
K = Val(Mid$(xYGOSDOS0.GOSDOSGSRV, 2, 2)): fgSelect.Text = arrService_Lib(K)
fgSelect.CellForeColor = vbRed
fgSelect.CellFontBold = True

For K = 0 To fgSelect_arrIndex
    fgSelect.Col = K
    fgSelect.CellBackColor = wBackColor

Next K
'If cmdSelect_SQL_K = "4" Then
fgSelect.Col = 6
    If xYGOSDOS0.GOSDOSSTAD = " " Or xYGOSDOS0.GOSDOSSTAD = "x" Then
        fgSelect.CellBackColor = RGB(255, 208, 128)
        If xYGOSDOS0.GOSDOSECHD <= DSys Then fgSelect.CellBackColor = mColor_W0
    End If
    
'End If
fgSelect.Col = 9
fgSelect.Text = xYGOSEVE0.GOSEVENAT
fgSelect.CellForeColor = wColor
fgSelect.Col = 10: fgSelect.Text = xYGOSEVE0.GOSEVEUUSR
fgSelect.CellForeColor = wColor

fgSelect.Col = 11: fgSelect.Text = dateImp10_S(xYGOSEVE0.GOSEVEUAMJ) & "  " & timeImp8(xYGOSEVE0.GOSEVEUHMS)
fgSelect.CellForeColor = wColor
If xYGOSEVE0.GOSEVEUSRV <> xYGOSDOS0.GOSDOSGSRV Then
    For K = 9 To 11
        fgSelect.Col = K
        fgSelect.CellBackColor = RGB(255, 208, 128)
    
    Next K
End If
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

End Sub

Public Sub fgSelect_DisplayLine_4Journal()
Dim K As Integer, X As String
Dim wColor As Long, wBackColor As Long
On Error Resume Next

If oldYGOSDOS0.GOSDOSIDD <> xYGOSDOS0.GOSDOSIDD Then
    oldYGOSDOS0 = xYGOSDOS0
    wBackColor = 0
    Call fgSelect_DisplayLine_YGOSDOS0_Color(wColor, wBackColor)
    If wBackColor = 0 Then wBackColor = mColor_B0
    fgSelect.Col = 0: fgSelect.Text = xYGOSDOS0.GOSDOSIDD & "-" & xYGOSDOS0.GOSDOSWMTK & " " & xYGOSDOS0.GOSDOSWES
    fgSelect.CellForeColor = wColor
    fgSelect.Col = 1
    K = Val(Mid$(xYGOSDOS0.GOSDOSISRV, 2, 2)): fgSelect.Text = arrService_Lib(K)
    fgSelect.CellForeColor = wColor
    fgSelect.Col = 2: fgSelect.Text = xYGOSDOS0.GOSDOSWBIC
    fgSelect.CellForeColor = wColor
    fgSelect.Col = 3: fgSelect.Text = Format$(xYGOSDOS0.GOSDOSWMTD, "### ### ### ##0.00") & " " & xYGOSDOS0.GOSDOSWDEV
    If xYGOSDOS0.GOSDOSSTAD = " " Then
        fgSelect.CellForeColor = vbRed
        fgSelect.CellFontBold = True
    Else
        fgSelect.CellForeColor = wColor
    End If
    
    fgSelect.Col = 4: fgSelect.Text = "  " & xYGOSDOS0.GOSDOSSTAG & " - " & xYGOSDOS0.GOSDOSSTAD
    fgSelect.CellForeColor = wColor
    
    
    For K = 0 To 7
        fgSelect.Col = K
        fgSelect.CellBackColor = wBackColor
        fgSelect.CellFontBold = True
    Next K
    fgSelect.Col = 6: fgSelect.Text = xYGOSDOS0.GOSDOSIDD

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
End If


Select Case xYGOSEVE0.GOSEVESTAE
    Case " ": wColor = RGB(60, 60, 60)
    Case "A": wColor = &H606060
    Case Else: wColor = vbMagenta
End Select

fgSelect.Col = 2
fgSelect.Text = xYGOSEVE0.GOSEVEUUSR
fgSelect.CellFontSize = 8
fgSelect.CellForeColor = wColor
fgSelect.Col = 3
fgSelect.Text = dateImp10_S(xYGOSEVE0.GOSEVEUAMJ) & " à " & timeImp8(xYGOSEVE0.GOSEVEUHMS)
fgSelect.CellForeColor = wColor
fgSelect.CellFontSize = 8
X = Trim(xYGOSEVE0.GOSEVETXT)
     Select Case Trim(xYGOSEVE0.GOSEVENAT)
        Case "Sus*": wColor = RGB(0, 0, 128): wBackColor = RGB(255, 255, 128)
        Case "Note": wColor = RGB(0, 0, 128): wBackColor = mColor_Y0
        Case "Res*", "AnnV", "AnnR", "AnnC": wColor = vbMagenta: wBackColor = mColor_Y0
        Case "Mail": wColor = vbBlack: K = InStr(X, vbCr): If K > 0 Then X = Mid$(X, 1, K - 1)
        Case "PJ**", "Swi+": wColor = RGB(0, 0, 96): wBackColor = RGB(240, 255, 255): K = InStr(X, vbCr): If K > 0 Then X = Mid$(X, 1, K - 1)
        Case "Swi>": wColor = RGB(0, 96, 0): wBackColor = mColor_Y0: If Len(X) > 48 Then X = Mid$(X, 1, 48)
        Case "Val": wColor = RGB(0, 0, 128): wBackColor = mColor_G2
        Case "Rej": wColor = RGB(0, 0, 128): wBackColor = mColor_W0
        Case "Clo": wColor = RGB(0, 0, 128): wBackColor = RGB(224, 224, 224)
       Case Else: wColor = vbMagenta
    End Select
    
fgSelect.Col = 2: fgSelect.CellBackColor = wBackColor
fgSelect.Col = 3: fgSelect.CellBackColor = wBackColor

fgSelect.Col = 4: fgSelect.Text = "  " & xYGOSEVE0.GOSEVENAT
fgSelect.CellFontBold = True
fgSelect.CellForeColor = wColor
fgSelect.CellBackColor = wBackColor
fgSelect.Col = 5
fgSelect.Text = X
fgSelect.CellForeColor = wColor
fgSelect.CellBackColor = wBackColor

fgSelect.Col = 6: fgSelect.Text = xYGOSEVE0.GOSEVEIDD
fgSelect.Col = 7: fgSelect.Text = xYGOSEVE0.GOSEVEIDE

End Sub


Public Sub fglist_DisplayLine_YGOSDOS0(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long, wBackColor As Long
Dim xSql As String, xUUMID As String
On Error Resume Next

Call fgSelect_DisplayLine_YGOSDOS0_Color(wColor, wBackColor)


fgList.Col = 0: fgList.Text = xYGOSDOS0.GOSDOSIDD
fgList.CellForeColor = wColor
fgList.Col = 1
K = Val(Mid$(xYGOSDOS0.GOSDOSGSRV, 2, 2)): fgList.Text = arrService_Lib(K)
fgList.CellForeColor = vbRed
fgList.CellFontBold = True

fgList.Col = 2: fgList.Text = xYGOSDOS0.GOSDOSWBIC
fgList.CellForeColor = wColor
If mList_YGOSDOS0.GOSDOSWBIC = Trim(xYGOSDOS0.GOSDOSWBIC) Then fgList.CellBackColor = vbGreen
fgList.Col = 3: fgList.Text = xYGOSDOS0.GOSDOSWTRN
fgList.CellForeColor = wColor
If mList_YGOSDOS0.GOSDOSWTRN = Trim(xYGOSDOS0.GOSDOSWTRN) Then fgList.CellBackColor = vbGreen

fgList.Col = 4: fgList.Text = xYGOSDOS0.GOSDOSWMTK & " " & xYGOSDOS0.GOSDOSWES
fgList.CellForeColor = wColor

fgList.Col = 5: fgList.Text = Format$(xYGOSDOS0.GOSDOSWMTD, "### ### ### ##0.00")
If xYGOSDOS0.GOSDOSSTAD = " " Then
    fgList.CellForeColor = vbRed
    fgList.CellFontBold = True
Else
    fgList.CellForeColor = wColor
End If
fgList.Col = 6: fgList.Text = xYGOSDOS0.GOSDOSWDEV
fgList.CellForeColor = wColor

fgList.Col = 7: fgList.Text = xYGOSDOS0.GOSDOSSTAD
fgList.CellForeColor = wColor

fgList.Col = 8: fgList.Text = dateImp10_S(xYGOSDOS0.GOSDOSECHD)
fgList.CellForeColor = wColor
fgList.Col = 9
K = Val(Mid$(xYGOSDOS0.GOSDOSRCOM, 2, 2)): fgList.Text = arrRCOM_Lib(K)
fgList.CellForeColor = wColor
fgList.Col = 10: fgList.Text = xYGOSDOS0.GOSDOSCLI
fgList.CellForeColor = wColor

fgList.Col = 11: fgList.Text = xYGOSDOS0.GOSDOSLABK
fgList.CellForeColor = wColor

For K = 0 To fglist_arrIndex
    fgList.Col = K
    fgList.CellBackColor = wBackColor
Next K

fgList.Col = fglist_arrIndex: fgList.Text = lIndex
End Sub

Public Sub fglist_DisplayLine_YSWISAB0()
Dim K As Integer
Dim wColor As Long, wColor_WL20 As Long
Dim xSql As String, xSWISABWL20 As String, xSWIENACET As Integer
On Error Resume Next

wColor_WL20 = RGB(64, 64, 64)
xSWISABWL20 = xYSWISAB0.SWISABWL20
xSWIENACET = 0

xSql = "select SWIENAREF, SWIENACET from " & paramIBM_Library_SAB & ".ZSWIENA0 " _
         & " where SWIENAETA = " & currentZMNURUT0.MNURUTETB _
         & " and SWIENAINT = " & xYSWISAB0.SWISABZSWI
Set rsSabX = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    xSWISABWL20 = rsSabX("SWIENAREF")
    xSWIENACET = rsSabX("SWIENACET")
    If xYSWISAB0.SWISABWL20 <> xSWISABWL20 Then wColor_WL20 = RGB(32, 128, 32)
End If




If xYSWISAB0.SWISABSWID = oldYSWISAB0.SWISABSWID Then
    wColor = vbMagenta
Else
    wColor = vbBlue
End If

fgList.Col = 0: fgList.Text = xYSWISAB0.SWISABWMTK & " " & xYSWISAB0.SWISABWES
fgList.CellForeColor = wColor
fgList.Col = 1: fgList.Text = xSWISABWL20
fgList.CellForeColor = wColor_WL20
If xYSWISAB0.SWISABSWID = oldYSWISAB0.SWISABSWID Then fgList.CellBackColor = vbGreen

fgList.Col = 2: fgList.Text = xYSWISAB0.SWISABOPEC & "  " & xYSWISAB0.SWISABOPEN
fgList.CellForeColor = wColor
If xSWIENACET = "1" Then fgList.Text = "annulé": fgList.CellBackColor = RGB(220, 220, 220)

fgList.Col = 3: fgList.Text = Format$(xYSWISAB0.SWISABWMTD, "### ### ### ##0.00")
fgList.CellForeColor = wColor
fgList.Col = 4: fgList.Text = xYSWISAB0.SWISABWDEV
fgList.CellForeColor = wColor

fgList.Col = 5: fgList.Text = dateImp10_S(xYSWISAB0.SWISABWAMJ) & "  " & timeImp8(xYSWISAB0.SWISABWHMS)
fgList.CellForeColor = wColor

fgList.Col = fglist_arrIndex: fgList.Text = xYSWISAB0.SWISABSWID

End Sub



Public Sub fgDetail_DisplayLine(lIndex As Long)
Dim K As Integer, iAsc13 As Integer, iLen As Integer
Dim blnAsc13 As Boolean
Dim xValue As String

On Error Resume Next
fgDetail.Col = 0: fgDetail.Text = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
fgDetail.Col = 1: xValue = Trim(rsSIDE_DB("value"))
Select Case rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
    Case "50", "50K", "50F": fgDetail_50 = xValue
    Case "57", "57A": fgDetail_57 = xValue
    Case "59", "59A", "59F": fgDetail_59 = xValue
    Case "32B": fgDetail_32B = xValue
    Case "33B": fgDetail_33B = xValue
    Case "36": fgDetail_36 = xValue
    Case "30V": fgDetail_30V = xValue
    Case "37G": fgDetail_37G = xValue
    Case "34E": fgDetail_34E = xValue
    Case "30T": fgDetail_30T = xValue
    Case "30P": fgDetail_30P = xValue
    Case "82A": fgDetail_82A = xValue
    Case "87A": fgDetail_87A = xValue
    Case "22C": fgDetail_22C = xValue
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
Dim blnField_79 As Boolean

blnField_79 = False

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
            'K2 = InStr(2, x, ":")
            If blnField_79 Then
                K2 = 0
            Else
                K2 = InStr(2, X, ":")
            End If
            If K2 > 0 Then
                fgDetail.Text = Trim(Mid$(X, K2 + 1, Len(X) - K2))
                fgDetail.Col = 0: fgDetail.Text = Trim(Mid$(X, 2, K2 - 2))
                If Trim(Mid$(X, 2, K2 - 2)) = "79" Then blnField_79 = True
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


K = InStr(xValue, ":32B") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_32B = Mid$(xValue, K, K2 - K - 3)
    End If
End If

K = InStr(xValue, ":33B") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_33B = Mid$(xValue, K, K2 - K - 3)
    End If
End If

K = InStr(xValue, ":36") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_36 = Mid$(xValue, K, K2 - K - 3)
    End If
End If

K = InStr(xValue, ":30V") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_30V = Mid$(xValue, K, K2 - K - 3)
    End If
End If

K = InStr(xValue, ":37G") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_37G = Mid$(xValue, K, K2 - K - 3)
    End If
End If

K = InStr(xValue, ":34E") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_34E = Mid$(xValue, K, K2 - K - 3)
    End If
End If

K = InStr(xValue, ":30T") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_30T = Mid$(xValue, K, K2 - K - 3)
    End If
End If

K = InStr(xValue, ":30P") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_30P = Mid$(xValue, K, K2 - K - 3)
    End If
End If

K = InStr(xValue, ":82A") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_82A = Mid$(xValue, K, K2 - K - 3)
    End If
End If

K = InStr(xValue, ":87A") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_87A = Mid$(xValue, K, K2 - K - 3)
    End If
End If

K = InStr(xValue, ":22C") + 2
If K > 2 Then
    K = InStr(K, xValue, ":") + 1
    If K > 1 Then
        K2 = InStr(K, xValue, ":") + 1
        fgDetail_22C = Mid$(xValue, K, K2 - K - 3)
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
Dim blnField_79 As Boolean
Dim wSWIL As Integer

On Error Resume Next
blnField_79 = False

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
            If blnField_79 Then
                K2 = 0
            Else
                K2 = InStr(2, X, ":")
            End If
            If K2 > 0 Then
                fgSwift.Text = Trim(Mid$(X, K2 + 1, Len(X) - K2))
                fgSwift.Col = 0: fgSwift.Text = Trim(Mid$(X, 2, K2 - 2))
                If Trim(Mid$(X, 2, K2 - 2)) = "79" Then blnField_79 = True
                
                fgSwift.CellBackColor = lCellBackColor
                fgSwift.CellForeColor = lColorFixed
                
            If mSWIECHSWIL > 0 Then
                wSWIL = wSWIL + 1
                If mSWIECHSWIL = wSWIL Then
                    fgSwift.CellForeColor = vbRed
                    fgSwift.Col = 1: fgSwift.CellForeColor = vbRed
                End If
            End If

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


Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long
fgSelect.Visible = False

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = lK
    Select Case lK
        Case 0: fgSelect.Col = 0: X = Format$(Val(fgSelect.Text), "0000000000")
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
    Case "@YSWISAB0":
        If paramEnvironnement = constProduction Then
            blnAuto = True: Form_Init
            Me.Enabled = False: Me.MousePointer = vbHourglass
            Importation_SAA
            DoEvents: Wait_SS 2
            
            Importation_SAA_SWISABZSWI
            DoEvents: Wait_SS 2
            
            Importation_SAA_SWISABZSWI_Reprise
            DoEvents: Wait_SS 2
            
            Importation_SAB_Dossier
            DoEvents: Wait_SS 2
            Importation_SAA_SWISABWSTA
            DoEvents: Wait_SS 2
            Importation_SAB_YSWIMON0
            DoEvents: Wait_SS 2
            Importation_SAB_YSWISAB1
            DoEvents: Wait_SS 2
            Importation_SAA_Alerte
            DoEvents: Wait_SS 2
            cmdSelect_SQL_1Lmail
            DoEvents: Wait_SS 2
            Importation_Jrnl
            
            DoEvents: Wait_SS 2
            Importation_SAB_ZSWIHIA0
            
            DoEvents: Wait_SS 2
            Importation_SAB_YSWIMON0_Synchro1
            Importation_SAB_YSWIMON0_Synchro2
            Importation_SAB_YSWIMON0_Synchro3
            
            DoEvents: Wait_SS 2
            Importation_SAB_YGOSDOS0_Synchro1
            Importation_SAB_YGOSDOS0_Synchro2
            
            Importation_SIDE_Reporting_Control
            Importation_SAA_Origine_MT
            Importation_SAA_Modification_MT
            
 '$JPL 2016-01-13
            DoEvents: Wait_SS 2
            Importation_SAA_198_Alerte
            
 '$JPL 2015-01-05
            YSWIRAM0_Importation
 '$JPL 2015-12-04
            DoEvents: Wait_SS 2
            YSWIECH0_Auto
            Unload Me
        Else
            Call MsgBox("Interdit en TEST", vbCritical, "@YSWISAB0")
        End If
        
    Case "@BIA_GOS": blnAuto = True: Form_Init
        Me.Enabled = False: Me.MousePointer = vbHourglass
        
        Call Auto_BIA_GOS
        
        cboSelect_SQL.AddItem "6 - Journal des messages 300/320 du " & dateImp10_S(YBIATAB0_DATE_CPT_JP0)
        Call cbo_Scan("6", cboSelect_SQL)
        chkSelect_GOSDOSIAMJ = "1"
        Call DTPicker_Set(txtSelect_GOSDOSIAMJ_Max, YBIATAB0_DATE_CPT_J) '
        Call DTPicker_Set(txtSelect_GOSDOSIAMJ_Min, YBIATAB0_DATE_CPT_JP0) '
        cmdSelect_SQL_6
        mnuPrint2_Mail_Click
        
        cboSelect_SQL.AddItem "7J - Journal des messages Swift,évènements du " & dateImp10_S(YBIATAB0_DATE_CPT_JP0)
        Call cbo_Scan("7J", cboSelect_SQL)
        chkSelect_GOSDOSIAMJ = "1"
        Call DTPicker_Set(txtSelect_GOSDOSIAMJ_Max, YBIATAB0_DATE_CPT_J) '
        Call DTPicker_Set(txtSelect_GOSDOSIAMJ_Min, YBIATAB0_DATE_CPT_JP0) '
        cboSelect_GOSDOSWMTK.ListIndex = 0
        cboSelect_GOSDOSWES.ListIndex = 0
        cmdSelect_SQL_7J
        mnuPrint2_Mail_Click

        Me.Enabled = True: Me.MousePointer = 0
        Unload Me
    Case "@YSWIRAM0": blnAuto = True: Form_Init
        Me.Enabled = False: Me.MousePointer = vbHourglass
        cboSelect_SQL.AddItem "6E - Echéancier des messages 300/320 en attente"
        Call cbo_Scan("6E", cboSelect_SQL)
        cmdSelect_SQL_6
        mnuPrint2_Mail_Click
        
        cboSelect_SQL.AddItem "7E - Echéancier des messages Swift en attente"
        Call cbo_Scan("7E", cboSelect_SQL)
        cmdSelect_SQL_7
        mnuPrint2_Mail_Click
        
        Me.Enabled = True: Me.MousePointer = 0
        Unload Me
    Case "JPL": blnAuto = True: Form_Init
            DSys = 20151130
 '           Me.Enabled = False: Me.MousePointer = vbHourglass
            Importation_SAA
            
            Importation_SAA_SWISABZSWI
            
            Importation_SAA_SWISABZSWI_Reprise
            
            Importation_SAB_Dossier
            Importation_SAA_SWISABWSTA
            Importation_SAB_YSWIMON0
            Importation_SAB_YSWISAB1
            'Importation_SAA_Alerte
           ' cmdSelect_SQL_1Lmail
           ' Importation_Jrnl
            
            Importation_SAB_ZSWIHIA0
            
            Importation_SAB_YSWIMON0_Synchro1
            Importation_SAB_YSWIMON0_Synchro2
            Importation_SAB_YSWIMON0_Synchro3
            
            Importation_SAB_YGOSDOS0_Synchro1
            Importation_SAB_YGOSDOS0_Synchro2
            
            Importation_SIDE_Reporting_Control
            Importation_SAA_Origine_MT
            Importation_SAA_Modification_MT

        Me.Enabled = True: Me.MousePointer = 0
        Unload Me
    Case Else: blnAuto = False: Form_Init
        If Mid$(Msg, 13, 6) = "SQL_3:" Then
            Call cbo_Scan("3", cboSelect_SQL)
            txtSelect_3_GOSDOSIDD = Val(Mid$(Msg, 19, Len(Msg) - 18))
            cmdSelect_SQL_3
        End If

End Select
End Sub


Public Sub Form_Init()
Dim V, xSql As String, X As String, X2 As String
Dim K As Long, wListIndex As Integer

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True

If arrMT_Nb = 0 Then arrMT_Load

cmdReset
blnControl = False


cboParam_GOSDOSLABK_GSrv.ForeColor = vbMagenta
txtParam_GOSDOSLABK_J.ForeColor = vbMagenta
lblParam_GOSDOSLABK_GSrv.ForeColor = vbRed
lblParam_GOSDOSLABK_J.ForeColor = vbRed
tabDetail.ForeColor = vbBlue                  'vbMagenta
'chkSelect_SWISABOPEN.ForeColor = vbMagenta

txtList_Add.BackColor = &HFFFF80
cmdList_Display.Top = cmdList_Add.Top
cmdList_Display.Left = cmdList_Add.Left

fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False
cboSelect_SWISABOPEC.BackColor = lblSelect_SWISABOPEN.BackColor
txtSelect_SWISABOPEN.BackColor = lblSelect_SWISABOPEN.BackColor

fraSelect_Options_1a.Visible = True
Set fraSelect_Options_1a.Container = fraSelect_Options
fraSelect_Options_1a.Top = fraSelect_Options_1.Top
fraSelect_Options_1a.Left = 5800

fraSelect_Options_1b.Visible = False
Set fraSelect_Options_1b.Container = fraSelect_Options
fraSelect_Options_1b.Top = fraSelect_Options_1.Top
fraSelect_Options_1b.Left = 5800

fraSelect_Options_3.Visible = False
Set fraSelect_Options_3.Container = fraSelect_Options
fraSelect_Options_3.Top = fraSelect_Options_1.Top
fraSelect_Options_3.Left = fraSelect_Options_1.Left

fraSelect_Options_4.Visible = False
Set fraSelect_Options_4.Container = fraSelect_Options
fraSelect_Options_4.Top = fraSelect_Options_1.Top
fraSelect_Options_4.Left = fraSelect_Options_1.Left

fraSelect_Options_9.Visible = False
Set fraSelect_Options_9.Container = fraSelect_Options
fraSelect_Options_9.Top = fraSelect_Options_1.Top
fraSelect_Options_9.Left = fraSelect_Options_1.Left

fraSelect_Options_7.Visible = False
Set fraSelect_Options_7.Container = fraSelect_Options
fraSelect_Options_7.Top = fraSelect_Options_1.Top
fraSelect_Options_7.Left = fraSelect_Options_1.Left



fraSelect_Options_Stat.Visible = False
Set fraSelect_Options_Stat.Container = fraSelect_Options
fraSelect_Options_Stat.Top = fraSelect_Options_1.Top
fraSelect_Options_Stat.Left = fraSelect_Options_1.Left

X = dateElp("FinDeMoisP", -1, DSys)
Call DTPicker_Set(txtSelect_Stat_AMJMax, X) '
X = dateElp("M-FM", -3, X)
X = dateElp("Jour", 1, X)
Call DTPicker_Set(txtSelect_Stat_AMJMin, X) '

fraSelect_Options_1.BorderStyle = 0
fraDetail_EVE.ForeColor = vbMagenta

fraDetail.Visible = False
Set fraDetail.Container = fraTab0
fraDetail.Top = fgSelect.Top
fraDetail.Left = fgSelect.Left + fgSelect.Width - fraDetail.Width - 50 '- 300


tabDetail.Tab = 0
Set fraDetail_C.Container = tabDetail
fraDetail_C.Top = fraList.Top ' libDetail_SWISABSWID.Top
fraDetail_C.Left = fraList.Left ' libDetail_SWISABSWID.Left + libDetail_SWISABSWID.Width + 200

tabDetail.Tab = 1
Set fraEVE.Container = tabDetail
fraEVE.Top = fraEVE_C.Top + 200

fraEVE.Left = fraEVE_C.Left + fraEVE_C.Width - fraEVE.Width - 200
fraEVE.Visible = False
fraEVE.Height = 7000
'fraEVE_S.BackColor = fraEVE.BackColor

Set fgModèle.Container = fraEVE_S
fgModèle.Top = txtGOSEVETXT.Top
fgModèle.Height = txtGOSEVETXT.Height
fgModèle.Left = txtGOSEVETXT.Left
'fgModèle.Width = txtGOSEVETXT.Width

Set fraEVE_Swift.Container = fraEVE_S
fraEVE_Swift.Top = txtGOSEVETXT.Top
fraEVE_Swift.Left = txtGOSEVETXT.Left

cmdEVE_Quit.Left = cmdEVE_Ok.Left + 160
cmdEVE_Dupliquer.Left = cmdEVE_Quit.Left
cmdEVE_Dupliquer.Top = cmdEVE_Quit.Top - 2000


cmdEVE_Invalidation.Left = cmdEVE_Validation.Left
cmdEVE_Invalidation.Top = cmdEVE_Validation.Top
cmdEVE_Ignore.Left = cmdEVE_Ok.Left
cmdEVE_Ignore.Top = cmdEVE_Ok.Top

Call DTPicker_Set(txtGOSEVEECHD, DSys) '

fraDetail_LAB.Enabled = False
fraDetail_LAB.Visible = False



Set fraPJ.Container = fraDetail
fraPJ.Visible = False
fraPJ.Left = tabDetail.Left
fraPJ.Top = tabDetail.Top

'Set fraList.Container = fraDetail_G
fraList.Visible = False
fraList.Top = fraDetail_C.Top ' 0  '450
'fraList.Left = fraDetail_C.Left ' 0 '6450

Set fraSwift.Container = fraDetail
fraSwift.Visible = False
fraSwift.Top = 1430
fraSwift.Height = fraDetail.Height - 300 '7200
fraSwift.Left = fraDetail.Width - fraSwift.Width - 500
libSWIFT_SWISABSWID.BackColor = mColor_Y0
libSWIFT_SWISABSWID.ForeColor = vbMagenta 'RGB(128, 64, 0)
fgSwift_FormatString = fgSwift.FormatString

Set fraSWISABKSRV.Container = fraTab0
fraSWISABKSRV.Visible = False
fraSWISABKSRV.Top = 2000
'fraSWISABKSRV.Height = 7200
fraSWISABKSRV.Left = 2800

ProgressBar1.Top = 345
ProgressBar1.Left = 6135


'Set fraMail_MT.Container = fraDetail
'fraMail_MT.Top = 100
'fraMail_MT.Left = 100
'fraMail_MT.Visible = False
'libMail_Subject.ForeColor = vbWhite


Set fraMail_MT.Container = fraTab0
fraMail_MT.Visible = False
fraMail_MT.Top = fraDetail.Top
fraMail_MT.Left = fraDetail.Left
lstMail_MT_CC.BackColor = cmdMail_MT_CC.BackColor
lstMail_MT_To.BackColor = cmdMail_MT_To.BackColor
lstMail_MT_Message.BackColor = cmdMail_MT_Message.BackColor


fgDetail.Visible = False: fraDetail.Visible = False
fgDetail_FormatString = fgDetail.FormatString
Call DTPicker_Set(txtSelect_GOSDOSIAMJ_Max, YBIATAB0_DATE_CPT_JS1) '
Call DTPicker_Set(txtSelect_GOSDOSIAMJ_Min, YBIATAB0_DATE_CPT_JS1) '
chkSelect_GOSDOSIAMJ.Enabled = True
chkSelect_GOSDOSIAMJ = "1"


lstW.Visible = False
lstW.Clear

cboSelect_3_GOSDOSSTAD.Clear
cboSelect_3_GOSDOSSTAD.AddItem " en cours"
cboSelect_3_GOSDOSSTAD.AddItem "*tous"
cboSelect_3_GOSDOSSTAD.AddItem "x à clôturer"
cboSelect_3_GOSDOSSTAD.AddItem "Annulés"
cboSelect_3_GOSDOSSTAD.AddItem "Clôturés"
cboSelect_3_GOSDOSSTAD.ListIndex = 0

cboSelect_3_GOSDOSSTAG.Clear
cboSelect_3_GOSDOSSTAG.AddItem " "
cboSelect_3_GOSDOSSTAG.AddItem "Validés"
cboSelect_3_GOSDOSSTAG.AddItem "Rejetés"
cboSelect_3_GOSDOSSTAG.ListIndex = 0

txtSelect_3_GOSDOSIDD.BackColor = lblSelect_3_GOSDOSIDD.BackColor
'Initialisation WES_______________________________________________________________________________
cboSelect_GOSDOSWES.Clear
cboSelect_GOSDOSWES.AddItem ""
cboSelect_GOSDOSWES.AddItem "Entrant"
cboSelect_GOSDOSWES.AddItem "Sortant"
cboSelect_GOSDOSWES.ListIndex = 0


'Initialisation devise________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdPrint, "Initialisation devise,MT,LABK ")

cboSelect_GOSDOSWDEV.Clear
cboSelect_GOSDOSWDEV.AddItem ""
cboSelect_GOSDOSWDEV.AddItem "EUR"
cboSelect_GOSDOSWDEV.AddItem "USD"
cboSelect_GOSDOSWDEV.ListIndex = 0

'Initialisation MTK_______________________________________________________________________________
cboSelect_GOSDOSWMTK.Clear
cboSelect_GOSDOSWMTK.AddItem ""
cboSelect_GOSDOSWMTK.AddItem "103"
cboSelect_GOSDOSWMTK.AddItem "103,202"
cboSelect_GOSDOSWMTK.AddItem "202"
cboSelect_GOSDOSWMTK.AddItem "700,701"
cboSelect_GOSDOSWMTK.AddItem "%99"
cboSelect_GOSDOSWMTK.ListIndex = 0
'Initialisation LABK______________________________________________________________________________

libGOSDOSCLI = ""

libGOSDOSLABK.Visible = False
libGOSDOSLABK.Top = cboGOSDOSLABK.Top
libGOSDOSLABK.Left = cboGOSDOSLABK.Left
lstParam.AddItem "Motif"
lstParam.AddItem "Note"
lstParam.AddItem "SwFR"
lstParam.AddItem "SwGB"
lstParam.AddItem "Mail"
lstParam.AddItem "Stat Code"
lstParam.AddItem "Stat Pays"
lstParam.AddItem "Motif => Swift"

lstParam.Visible = True
lstParam_GOSDOSLABK.Visible = False
fraParam_GOSDOSLABK.Visible = False
Call lstParam_GOSDOSLABK_Load("DOS", "cboGOSDOSLABK")
'If cboGOSDOSLABK.ListCount > 0 Then cboGOSDOSLABK.ListIndex = 0


'Initialisation SER SSE_______________________________________________________________________________
cboSWISABSER.Clear
cboSWISABSER.AddItem "00-00"
cboSWISABSER.AddItem "00-CR"
cboSWISABSER.AddItem "00-TR"
cboSWISABSER.AddItem "CP-CP"
cboSWISABSER.AddItem "TC-TC"

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
cboGOSDOSGSRV.Clear
cboGOSEVEGSRV.Clear



'lstParam_MAIL.Clear
'lstParam_MAIL.AddItem "1 - adresses mail GOS_DOS"
'lstParam_MAIL.AddItem "2 - adresses mail SAA_Alerte"
'lstParam_MAIL.AddItem "3 - complément adresses mail RCOM"
'lstParam_MAIL.ListIndex = 0
'lstParam_GOSDOSMAIL.Clear

'txtSelect_rTextField.Visible = False
'txtSelect_rTextField.Top = cboSelect_GOSDOSKSRV.Top
'txtSelect_rTextField.Left = cboSelect_GOSDOSKSRV.Left
'txtSelect_rTextField.Height = cboSelect_GOSDOSKSRV.Height
'txtSelect_rTextField.Width = cboSelect_GOSDOSKSRV.Width

'lblSelect_rTextField.Visible = False
'lblSelect_rTextField.Top = chkSelect_GOSDOSKSRV.Top
'lblSelect_rTextField.Left = chkSelect_GOSDOSKSRV.Left
'lblSelect_rTextField.Height = chkSelect_GOSDOSKSRV.Height
'lblSelect_rTextField.Width = chkSelect_GOSDOSKSRV.Width

cboSelect_GOSDOSKSRV.Clear
cboSelect_GOSDOSKSRV.AddItem "  "
cboSelect_GOSDOSKSRV.AddItem "S00 - none"

cboSelect_3_GOSDOSGSRV.Clear
cboSelect_3_GOSDOSGSRV.AddItem "  "

cboSelect_7_SRV.Clear
cboSelect_7_SRV.AddItem "  "

cboSelect_9_SWISABKSRV.Clear
cboSelect_9_SWISABKSRV.AddItem "  "

cboList_SWISABKSRV.Clear
cboSWISABKSRV.Clear
cboParam_GOSDOSLABK_GSrv.Clear

cboParam_SAA_Alerte.Clear
cboParam_SAA_Alerte.AddItem ""
cboParam_SAA_Alerte.AddItem "S=  - service émetteur"

xSql = "select *from " & paramIBM_Library_SABSPE & ".YSSIUSR0 where SSIUSRNAT= 'S' and SSIUSRSTAK = ' ' order by SSIUSRUNIT"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    X = rsSab("SSIUSRUNIT") & " - " & rsSab("SSIUSRPRFX")
    cboGOSDOSGSRV.AddItem X
    cboGOSEVEGSRV.AddItem X
    cboSelect_3_GOSDOSGSRV.AddItem X
    'lstParam_GOSDOSMAIL.AddItem X
    cboSelect_9_SWISABKSRV.AddItem X
    cboList_SWISABKSRV.AddItem X
    cboSelect_GOSDOSKSRV.AddItem X
    cboSWISABKSRV.AddItem X
    cboParam_GOSDOSLABK_GSrv.AddItem X
    cboParam_SAA_Alerte.AddItem X
    X2 = rsSab("SSIUSRUNIT")
    If X2 = "S01" Or X2 = "S10" Or X2 = "S32" Then cboSelect_7_SRV.AddItem X
    K = Val(Mid$(rsSab("SSIUSRUNIT"), 2, 2))
    arrService_Code(K) = Mid$(rsSab("SSIUSRUNIT"), 1, 3)
    arrService_Lib(K) = Trim(Mid$(rsSab("SSIUSRPRFX"), 1, 12))
    rsSab.MoveNext
Loop
cboGOSDOSGSRV.ListIndex = 0
cboSelect_GOSDOSKSRV.ListIndex = 0
Call cbo_Scan(currentSSIWINUNIT, cboSelect_9_SWISABKSRV)
Call cbo_Scan(currentSSIWINUNIT, cboSelect_GOSDOSKSRV)
Call cbo_Scan(currentSSIWINUNIT, cboSelect_3_GOSDOSGSRV)
Call cbo_Scan(currentSSIWINUNIT, cboSelect_7_SRV)

'

lstParam_YSSIMEL0_Load
lstParam_YSSIMEL0_USR_Load

xSql = "select count(*) as Tally from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'SWISABWSRV' "
Set rsSab = cnsab.Execute(xSql)
If rsSab("Tally") = 0 Then Param_SWISABWSRV_Init

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'SWISABWSRV' order by BIATABK1"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF

    K = Val(Mid$(rsSab("BIATABK1"), 2, 2))
    arrService_Code_SAA(K) = Trim(Mid$(rsSab("BIATABTXT"), 1, 4))
    rsSab.MoveNext
Loop


'Initialisation PAYS ______________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdPrint, "Initialisation pays ")

cboGOSDOSPAYS.Clear
cboGOSDOSPAYS.AddItem "  "

X = "select * from  " & paramIBM_Library_SAB & ".ZBASTAB0 " _
    & " where BASTABETA = 1 and BASTABNUM = 11 order by BASTABARG"
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    cboGOSDOSPAYS.AddItem Mid$(rsSab("BASTABARG"), 4, 2) & " - " & Mid$(rsSab("BASTABLO1"), 4, 9) & Mid$(rsSab("BASTABLO2"), 1, 3)
    rsSab.MoveNext
Loop

'Initialisation RCOM ______________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdPrint, "Initialisation Responsables commerciaux ")

wListIndex = -1

For K = 1 To 99
     arrRCOM_Code(K) = "R" & Format$(K, "00"): arrRCOM_Lib(K) = arrRCOM_Code(K)
     arrRCOM_Mail(K) = ""
Next K

cboGOSDOSRCOM.Clear
cboGOSDOSRCOM.AddItem "  "
cboSelect_3_GOSDOSRCOM.Clear
cboSelect_3_GOSDOSRCOM.AddItem "  "

X = "select * from  " & paramIBM_Library_SAB & ".ZBASTAB0 " _
    & " where  BASTABETA = 1 and BASTABNUM = 6 and BASTABARG like 'CLIR%' order by BASTABARG"
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    X = Mid$(rsSab("BASTABARG"), 4, 3) & " - " & Mid$(rsSab("BASTABDON"), 24, 10)
    cboGOSDOSRCOM.AddItem X
    cboSelect_3_GOSDOSRCOM.AddItem X
    K = Val(Mid$(rsSab("BASTABARG"), 5, 2))
    arrRCOM_Code(K) = Mid$(rsSab("BASTABARG"), 4, 3)
    arrRCOM_Lib(K) = Trim(Mid$(rsSab("BASTABDON"), 24, 10))
    'If arrRCOM_Lib(K) = "" Then arrRCOM_Lib(K) = arrRCOM_Code(K)
    '''If K = 32 Then arrRCOM_Lib(K) = arrRCOM_Lib(K) & ";PERRET"
    ''''If K = 80 Then arrRCOM_Lib(K) = "GIACOMONI;CHEYRON"
    If arrRCOM_Lib(K) <> "" Then
        If InStr(usrName_UCase, arrRCOM_Lib(K)) > 0 Then
            cboSelect_3_GOSDOSGSRV.ListIndex = -1
            wListIndex = cboSelect_3_GOSDOSRCOM.ListCount - 1
        End If
    End If
    rsSab.MoveNext
Loop

xSql = "select *from " & paramIBM_Library_SABSPE & ".YSSIMEL0 where SSIMELNAT = '@' and SSIMELUIDX like 'RCOM.%'"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    X = Trim(rsSab("SSIMELUIDX"))
    K = InStr(X, ".R")
    K = Val(Mid$(X, K + 2, 2))
    If arrRCOM_Lib(K) = "" Then
        arrRCOM_Lib(K) = Trim(rsSab("SSIMELINFO"))
    Else
        arrRCOM_Lib(K) = arrRCOM_Lib(K) & ";" & Trim(rsSab("SSIMELINFO"))
    End If
    rsSab.MoveNext
Loop

'xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'GOSDOSMAIL_R' order by BIATABK1"
'Set rsSab = cnsab.Execute(xSql)
'Do While Not rsSab.EOF
'    K = Val(Mid$(rsSab("BIATABK1"), 2, 2))
'    If arrRCOM_Lib(K) = "" Then
'        arrRCOM_Lib(K) = Trim(rsSab("BIATABTXT"))
'    Else
'        arrRCOM_Lib(K) = arrRCOM_Lib(K) & ";" & Trim(rsSab("BIATABTXT"))
'    End If
'    rsSab.MoveNext
'Loop

cboSelect_3_GOSDOSRCOM.ListIndex = wListIndex

'Initialisation WBIC ______________________________________________________________________________

cboSelect_GOSDOSWBIC.Clear
cboSelect_GOSDOSWBIC.AddItem ""
cboSelect_GOSDOSWBIC.ListIndex = 0


cboSelect_3_GOSDOSWBIC.Clear
cboSelect_3_GOSDOSWBIC.AddItem ""
xSql = "select distinct GOSDOSWBIC from " & paramIBM_Library_SABSPE & ".YGOSDOS0 order by GOSDOSWBIC"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_3_GOSDOSWBIC.AddItem Trim(rsSab("GOSDOSWBIC"))
    rsSab.MoveNext
Loop
'Initialisation CLI ______________________________________________________________________________
cboSelect_3_GOSDOSCLI.Clear
cboSelect_3_GOSDOSCLI.AddItem ""
xSql = "select distinct GOSDOSCLI from " & paramIBM_Library_SABSPE & ".YGOSDOS0 order by GOSDOSCLI"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_3_GOSDOSCLI.AddItem Trim(rsSab("GOSDOSCLI"))
    rsSab.MoveNext
Loop

'Initialisation cboSelect_SWISABOPEC ______________________________________________________________________________
cboSelect_SWISABOPEC.Clear
cboSelect_SWISABOPEC.AddItem ""
cboSWISABOPEC.Clear
xSql = "select distinct SWISABOPEC from " & paramIBM_Library_SABSPE & ".YSWISAB0 order by SWISABOPEC"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    cboSelect_SWISABOPEC.AddItem Trim(rsSab("SWISABOPEC"))
    cboSWISABOPEC.AddItem Trim(rsSab("SWISABOPEC"))
    rsSab.MoveNext
Loop

'Initialisation mGOSDOSIDD_Last ______________________________________________________________________________
xSql = "select count(*) as Tally    from " & paramIBM_Library_SABSPE & ".YGOSDOS0 "
Set rsSab = cnsab.Execute(xSql)
mGOSDOSIDD_Last = rsSab("Tally")
Call lstErr_AddItem(lstErr, cmdPrint, "nb dossiers :  " & mGOSDOSIDD_Last)


'Initialisation SWISABWSTA_______________________________________________________________________________
cboSelect_SWISABWSTA.Clear
cboSelect_SWISABWSTA.AddItem ""
cboSelect_SWISABWSTA.AddItem "Live"
cboSelect_SWISABWSTA.AddItem "Nacked"
cboSelect_SWISABWSTA.AddItem "Acked"

'__________________________________________________________________________________________________________

tabParam.ForeColor = vbMagenta
tabParam.Tab = 0
tabParam.Caption = "paramétrage du service " & arrService_Lib(Mid$(currentSSIWINUNIT, 2, 2))

'__________________________________________________________________________________________________________
paramGOSDOS_Path = paramServer("\\ROPDOS\") & "GOSDOS\" & paramEnvironnement & "\"
paramGOSDOS_Path_DROPI = paramServer("\\ROPDOS_DROPI\" & "GOSDOS\" & paramEnvironnement & "\")

dirListBox.path = "C:\Temp"
blnfilDoc_Path = False
On Error Resume Next
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'GOSDOSPJ**' and BIATABK1 = '" & usrName_UCase & "'"
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    dirListBox.path = Trim(rsSab("BIATABTXT"))
    blnfilDoc_Path = True
End If

mfilDoc_Path = dirListBox.path
cmdGOSDOSPJ.Visible = False
'__________________________________________________________________________________________________________
fgEVE_FormatString = fgEVE.FormatString
txtFg.Width = fgEVE.ColWidth(4)
txtFg.FontName = fgEVE.Font.Name
txtFg.FontSize = fgEVE.Font.Size

fraEVE_Swift.Visible = False
cboEVE_Swift_MT.Clear
cboEVE_Swift_MT.AddItem "199"
cboEVE_Swift_MT.AddItem "292"
cboEVE_Swift_MT.AddItem "299"
cboEVE_Swift_MT.AddItem "392"
cboEVE_Swift_MT.AddItem "399"
cboEVE_Swift_MT.AddItem "499"
cboEVE_Swift_MT.AddItem "799"
cboEVE_Swift_MT.AddItem "999"
cboEVE_Swift_MT.BackColor = mColor_W1
'________________________________________________________________________________________________________

cboSelect_9_SWISABKSTA.Clear
cboSelect_9_SWISABKSTA.AddItem "! messages en attente"
cboSelect_9_SWISABKSTA.AddItem "* historique"
If arrHab(11) Then cboSelect_9_SWISABKSTA.AddItem "  tous les messages entrants"
cboSelect_9_SWISABKSTA.ListIndex = 0

'Initialisation Echéancier ______________________________________________________________________________

Call DTPicker_Set(txtSelect_4_GOSDOSECHD, DSys) '

cboSelect_4_GOSDOSISRV.Clear
cboSelect_4_GOSDOSISRV.AddItem ""
cboSelect_4_GOSDOSISRV.AddItem currentSSIWINUNIT & " - " & arrService_Lib(Mid$(currentSSIWINUNIT, 2, 2))
xSql = "select distinct GOSDOSISRV from " & paramIBM_Library_SABSPE & ".YGOSDOS0 order by GOSDOSISRV"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    X = rsSab("GOSDOSISRV")
    If X <> currentSSIWINUNIT Then
        K = Val(Mid$(X, 2, 2))
        cboSelect_4_GOSDOSISRV.AddItem X & " - " & arrService_Lib(K)
    End If
    rsSab.MoveNext
Loop

cboSelect_4_GOSDOSGSRV.Clear
cboSelect_4_GOSDOSGSRV.AddItem " "
cboSelect_4_GOSDOSGSRV.AddItem currentSSIWINUNIT & " - " & arrService_Lib(Mid$(currentSSIWINUNIT, 2, 2))
xSql = "select distinct GOSDOSGSRV from " & paramIBM_Library_SABSPE & ".YGOSDOS0 order by GOSDOSGSRV"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    X = rsSab("GOSDOSGSRV")
    If X <> currentSSIWINUNIT Then
        K = Val(Mid$(X, 2, 2))
        cboSelect_4_GOSDOSGSRV.AddItem X & " - " & arrService_Lib(K)
    End If
    rsSab.MoveNext
Loop

Call cbo_Scan(currentSSIWINUNIT, cboSelect_4_GOSDOSGSRV)


'Initialisation param SAA Alertes______________________________________________________________________________
If arrHab(16) Then
    xSql = "select distinct BIATABK1 from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
         & " where BIATABID = 'SAA' order by BIATABK1"
    Set rsSab = cnsab.Execute(xSql)
    
    Do While Not rsSab.EOF
        lstParam_SAA_Id.AddItem rsSab("BIATABK1")
        rsSab.MoveNext
    Loop

    lstParam_SAA_Id.Visible = True
    lstParam_SAA_K1.Visible = True
End If
'________________________________________________________________________________________________________
fraSelect_Options.Visible = True

If arrHab(15) Then Form_Init_Options_J


cboParam_SAA_TopK.Clear
cboParam_SAA_TopK.AddItem ""
cboParam_SAA_TopK.AddItem "* - tous les évènements"
cboParam_SAA_TopK.AddItem "# - cas particuliers"



'If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0
'If currentSSIWINUNIT = "S42" Then
If cboSelect_SQL.ListCount > 0 Then Call cbo_Scan("3", cboSelect_SQL): mSelect_SQL_Listindex_3 = cboSelect_SQL.ListIndex
'Else
'End If

SSTab1.Tab = 2: SSTab1.Tab = 1: SSTab1.Tab = 0
tabDetail.Tab = 1: tabDetail.Tab = 0
tabParam.Tab = 3: tabParam.Tab = 2: tabParam.Tab = 1: tabParam.Tab = 0

If last_Jrnl_date_time_ES = "00:00:00" Then
    last_Jrnl_date_time_ES = dateImp10_S(DSys) & " 00:00:02"
    last_mesg_crea_date_time = last_Jrnl_date_time_ES
    last_Jrnl_date_time_EVE = last_Jrnl_date_time_ES
    
    last_Alerte_date_time_ES = dateImp10_S(DSys) & " 00:00:01"
    last_Alerte_date_time_EVE = last_Alerte_date_time_ES
End If

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
    For I = 1 To fgSelect.FixedCols Step -1
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
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






Private Sub cboEVE_Swift_BIC_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub cboEVE_Swift_MT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub cboGOSDOSLABK_Click()
Dim X As String
X = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where GOSEVEIDD = -1 and GOSEVEGSRV = '" & currentSSIWINUNIT & "'" _
     & " and GOSEVENAT = 'DOS' and  SUBSTRING(GOSEVETXT , 1 , 10) = '" & Trim(cboGOSDOSLABK.Text) & "'"
Set rsSab = cnsab.Execute(X)
If Not rsSab.EOF Then
    If cmdSelect_SQL_K = "2" Or cmdSelect_SQL_K = "2d" Or cmdSelect_SQL_K = "2-RAM" Or cmdSelect_SQL_K = "9" Or cmdSelect_SQL_K = "9+" Then
        X = Trim(rsSab("GOSEVETXT"))
        
        txtGOSDOSTXT = Mid$(X, 31, Len(X) - 30)
        Call cbo_Scan(Mid$(X, 16, 3), cboGOSDOSGSRV)
        Call DTPicker_Set(txtGOSDOSECHD, dateElp("Ouvré", Mid$(X, 12, 3), DSys)) '
    End If
End If


End Sub

Private Sub cboGOSEVENAT_Change()
cboGOSEVENAT_Display
End Sub

Private Sub cboGOSEVENAT_Click()
cboGOSEVENAT_Display
End Sub

Private Sub cboList_SWISABKSRV_Change()
'fraList_Display_Habilitation
cmdList_SWISABKSRV.BackColor = vbGreen
End Sub

Private Sub cboList_SWISABKSRV_Click()
'fraList_Display_Habilitation
cmdList_SWISABKSRV.BackColor = vbGreen
End Sub


Private Sub cboSelect_3_GOSDOSCLI_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cboSelect_3_GOSDOSCLI_Click()
If blnControl Then cmdSelect_Clear


End Sub


Private Sub cboSelect_3_GOSDOSCLI_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub

Private Sub cboSelect_3_GOSDOSGSRV_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cboSelect_3_GOSDOSGSRV_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_3_GOSDOSRCOM_Change()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_3_GOSDOSRCOM_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_3_GOSDOSSTAD_Change()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_3_GOSDOSSTAD_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_3_GOSDOSSTAG_Change()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_3_GOSDOSSTAG_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_3_GOSDOSWBIC_Change()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_3_GOSDOSWBIC_Click()
If blnControl Then cmdSelect_Clear

End Sub


Private Sub cboSelect_3_GOSDOSWBIC_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub cboSelect_9_SWISABKSRV_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cboSelect_9_SWISABKSRV_Click()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cboSelect_9_SWISABKSTA_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cboSelect_9_SWISABKSTA_Click()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cboSelect_GOSDOSKSRV_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cboSelect_GOSDOSWDEV_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub cboSelect_GOSDOSWES_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cboSelect_GOSDOSWES_Click()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cboSelect_GOSDOSWES_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub cboSelect_GOSDOSWMTK_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cboSelect_GOSDOSWMTK_KeyPress(KeyAscii As Integer)
'num_KeyAscii KeyAscii

End Sub

Private Sub cboSelect_GOSDOSWBIC_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub cboSelect_SQL_Click()
cmdSelect_Reset
End Sub


Private Sub cboSelect_GOSDOSWBIC_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cboSelect_GOSDOSWBIC_Click()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cboSelect_GOSDOSWDEV_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cboSelect_GOSDOSWDEV_Click()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub cboSelect_GOSDOSWMTK_Click()
If blnControl Then cmdSelect_Clear
End Sub


Private Sub chkGOSEVEGSRV_Click()
If chkGOSEVEGSRV.value = "1" Then
    cboGOSEVEGSRV.Enabled = True
Else
    cboGOSEVEGSRV.Enabled = False
End If

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

Private Sub chkSelect_9_SWISABKSRV_Click()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub chkSelect_GOSDOSIAMJ_Click()
If blnControl Then cmdSelect_Clear
End Sub

Private Sub chkSelect_GOSDOSKSRV_Click()
If blnControl Then cmdSelect_Clear

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

Private Sub cmdDetail_Lab_Link_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

mYSWILNK0_Fct = "New"
rsYSWILNK0_Init oldYSWILNK0
oldYSWILNK0.SWILNKAPPC = "GOS"
oldYSWILNK0.SWILNKSWID = oldYSWISAB0.SWISABSWID


m999_YSWISAB0 = oldYSWISAB0

cmdList_New.Visible = False
cmdList_Add.Visible = False: cmdList_Display.Visible = True
cmdList_SAB_Modification.Visible = False
cmdList_Ignore.Visible = False
cmdList_SAB_Annulation.Visible = False
cmdList_Ignore.Visible = False
cmdList_SWISABKSRV.Visible = False
cboList_SWISABKSRV.Enabled = False
fraList_Options.Visible = True 'arrHab(13)

Call arrYGOSDOS0_SQL("where GOSDOSSTAD = ' ' order by GOSDOSIDD")
fglist_Display_YGOSDOS0
fraList.Visible = True
fraDetail_C.Visible = False
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDetail_LAB_Ok_Click()
Dim xSql As String
Me.Enabled = False: Me.MousePointer = vbHourglass

If IsNull(fraDetail_LAB_Control) Then
    If newYGOSDOS0.GOSDOSIDD <> 0 Then
        mYGOSDOS0_Fct = "Update": mYGOSEVE0_Fct = "Update"
    Else
        mYGOSDOS0_Fct = "New": mYGOSEVE0_Fct = "New"
        xSql = "select GOSDOSIDD from " & paramIBM_Library_SABSPE & ".YGOSDOS0 " _
             & "  Where GOSDOSIDD >= " & mGOSDOSIDD_Last & " order by GOSDOSIDD desc"
        Set rsSab = cnsab.Execute(xSql)
        
        If rsSab.EOF Then
            If mGOSDOSIDD_Last = 0 Then
                newYGOSDOS0.GOSDOSIDD = 1
                newYGOSEVE0.GOSEVEIDD = 1
                newYGOSEVE0.GOSEVEIDE = 1
            Else
                'cmdYGOSDOS0_Update = "cmdYGOSDOS0_Update : erreur initialisation GOSDOSIDD"
                MsgBox "cmdYGOSDOS0_Update : erreur initialisation GOSDOSIDD", vbCritical, Me.Name & " : cmdYGOSDOS0_Update"
                Exit Sub
            End If
        Else
            'mGOSDOSIDD_Last =
            newYGOSDOS0.GOSDOSIDD = rsSab("GOSDOSIDD") + 1
            newYGOSEVE0.GOSEVEIDD = newYGOSDOS0.GOSDOSIDD
            newYGOSEVE0.GOSEVEIDE = 1
    
        End If
        
        xSql = "select SWISABSWID from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABWID1 = " & newYGOSDOS0.GOSDOSWID1 _
        & " and SWISABWIDL = " & newYGOSDOS0.GOSDOSWIDL _
        & " and SWISABWIDH = " & newYGOSDOS0.GOSDOSWIDH
        
        Set rsSab = cnsab.Execute(xSql)
        If Not rsSab.EOF Then newYGOSEVE0.GOSEVESWID = rsSab("SWISABSWID")
        
        newYGOSEVE0.GOSEVEUAMJ = DSys
        arrYGOSEVE0_Nb = 1
        ReDim arrYGOSEVE0(1)
        arrYGOSEVE0(1) = newYGOSEVE0

        mYSWILNK0_Fct = "New"
        rsYSWILNK0_Init oldYSWILNK0
        oldYSWILNK0.SWILNKAPPC = "GOS"
        oldYSWILNK0.SWILNKSWID = newYGOSEVE0.GOSEVESWID

    End If
    blnYGOSDOS0_Update = True
    
    Select Case Mid$(cmdSelect_SQL_K, 1, 1)
        Case "9":
                    mYSWISAB0_Fct = "Update"
                    oldYSWISAB0 = m999_YSWISAB0
                    newYSWISAB0 = oldYSWISAB0
                    newYSWISAB0.SWISABXGOS = "G": newYSWISAB0.SWISABK999 = "G"
                    cmdMail_MT_NOK_Click
        Case "2"
                    If cmdSelect_SQL_K = "2-RAM" Then
                        cmdMail_MT_NOK_Click
                    Else
                        If currentSSIWINUNIT = Trim(newYGOSDOS0.GOSDOSGSRV) Then
                            'cmdMail_MT_NOK_Click
                            fraMail_Confirm  '$JPL 2012-05-09
                        Else
                            fraMail_Confirm
                        End If
                    End If
        Case Else: fraMail_Confirm
    End Select
       
    
    If blnYGOSDOS0_Display Then
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSDOS0 where GOSDOSIDD = " & newYGOSDOS0.GOSDOSIDD
        Set rsSab = cnsab.Execute(xSql)
        
        If Not rsSab.EOF Then
            V = rsYGOSDOS0_GetBuffer(rsSab, oldYGOSDOS0)
            newYGOSDOS0 = oldYGOSDOS0
            m999_YGOSDOS0 = oldYGOSDOS0
            blnYGOSDOS0_New = True
            Mesg_aid = oldYGOSDOS0.GOSDOSWID1
            mesg_s_umidl = oldYGOSDOS0.GOSDOSWIDL
             mesg_s_umidh = oldYGOSDOS0.GOSDOSWIDH
            
            fgDetail_Display
            cmdEVE_New_Click
        End If

    End If
    
        'cmdYGOSDOS0_Update "New"
        'fraDetail.Visible = False
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Function cmdYGOSDOS0_Update(lYGOSDOS0_Fct As String, lYGOSEVE0_Fct As String, lYSWILNK0_Fct As String, lYSWISAB0_Fct As String, lZSWIENA0_Fct As String)
Dim K As Integer, xSql As String, X As String
Dim V
Dim blnYBIADTAQ As Boolean

'On Error GoTo Error_Handler

If cmdSelect_SQL_K = "2d" Then
    lYSWILNK0_Fct = ""
Else
    If lYGOSDOS0_Fct = "New" And newYGOSDOS0.GOSDOSWIDL = 0 Then
        Call MsgBox("Erreur sporadique non résolue, prévenir JPL", vbCritical, "BIA_GOS : cmdYGOSDOS0_Update")
        Exit Function
    End If
End If
If lYGOSEVE0_Fct = "New" Then
    xSql = "select GOSEVEIDE from " & paramIBM_Library_SABSPE & ".YGOSEVE0 " _
         & "  Where GOSEVEIDD = " & newYGOSEVE0.GOSEVEIDD & " order by GOSEVEIDE desc"
    Set rsSab = cnsab.Execute(xSql)
    
    If rsSab.EOF Then
        newYGOSEVE0.GOSEVEIDE = 1
    Else
        newYGOSEVE0.GOSEVEIDE = rsSab("GOSEVEIDE") + 1
    End If
    If newYGOSEVE0.GOSEVENAT = "PJ**" Then
        If Not IsNull(cmdYGOSEVE0_Update_PJ) Then Exit Function
    End If
End If
If lYGOSEVE0_Fct = "New" And newYGOSEVE0.GOSEVENAT = "Swi>" Then
    blnYBIADTAQ = True
    newYGOSEVE0.GOSEVESTAE = "?"
    Call sqlYBIADTAQ_BIADTAID(newYBIADTAQ)
Else
    blnYBIADTAQ = False
End If
'________________________________________________________________________________
 If mYSWILNK0_Fct = "New" Then
    newYSWILNK0 = oldYSWILNK0
    newYSWILNK0.SWILNKAPPN = newYGOSDOS0.GOSDOSIDD
    newYGOSEVE0.GOSEVESWID = newYSWILNK0.SWILNKSWID
End If
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case lYGOSDOS0_Fct
    Case "Update": V = sqlYGOSDOS0_Update(newYGOSDOS0, oldYGOSDOS0, True)
    Case "New": V = sqlYGOSDOS0_Insert(newYGOSDOS0)
    Case "Delete": V = sqlYGOSDOS0_Delete(oldYGOSDOS0)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case lYGOSEVE0_Fct
    Case "Update":     V = sqlYGOSEVE0_Update(newYGOSEVE0, oldYGOSEVE0, True)
    Case "New": V = sqlYGOSEVE0_Insert(newYGOSEVE0)
    Case "Delete": V = sqlYGOSEVE0_Delete(oldYGOSEVE0)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
If blnYBIADTAQ Then
    newYBIADTAQ.BIADTAFCT = "GOS-SWI>"
    newYBIADTAQ.BIADTATXTE = Format$(newYGOSEVE0.GOSEVEIDD, "000000000") _
                          & Format$(newYGOSEVE0.GOSEVEIDE, "000000000") _
                          & arrService_Code_SAA(Mid$(newYGOSEVE0.GOSEVEUSRV, 2, 2))
    V = sqlYBIADTAQ_Insert(newYBIADTAQ)
    If Not IsNull(V) Then GoTo Error_MsgBox
End If
'________________________________________________________________________________
If lYSWILNK0_Fct = "New" Then
    If lYGOSDOS0_Fct = "New" Then
        X = "Set SWISABXGOS = 'G'"
    Else
        X = "Set SWISABXEVE = 'G'"
    End If
    V = sqlYSWISAB0_Update_Field(newYSWILNK0.SWILNKSWID, X)
    If Not IsNull(V) Then GoTo Error_MsgBox
End If
Select Case lYSWILNK0_Fct
    Case "Update":     V = sqlYSWILNK0_Update(newYSWILNK0, oldYSWILNK0, True)
    Case "New": V = sqlYSWILNK0_Insert(newYSWILNK0)
    Case "Delete": V = sqlYSWILNK0_Delete(oldYSWILNK0)
End Select
mYSWILNK0_Fct = ""
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case lYSWISAB0_Fct
    Case "Update":     V = sqlYSWISAB0_Update(newYSWISAB0, oldYSWISAB0)
    Case "New": V = sqlYSWISAB0_Insert(newYSWISAB0)
    Case "Delete": V = sqlYSWISAB0_Delete(oldYSWISAB0)
End Select
mYSWISAB0_Fct = ""
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case lZSWIENA0_Fct
    Case "Update":     V = sqlZSWIENA0_Update(newZSWIENA0, oldZSWIENA0)
End Select
mZSWIENA0_Fct = ""
If Not IsNull(V) Then GoTo Error_MsgBox

'________________________________________________________________________________


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYGOSDOS0_Update"
Exit_sub:

    cmdYGOSDOS0_Update = V
    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        
        If blnYBIADTAQ Then
        
            fraMail_MT.Visible = False

            ProgressBar1.Visible = True
            ProgressBar1.Min = 0: ProgressBar1.Max = 31
            ProgressBar1.value = 0
            V = sqlYBIADTAQ_BIADTASTA(newYBIADTAQ, ProgressBar1)
            ProgressBar1.Visible = False

            If Not IsNull(V) Then Call MsgBox(V, vbCritical, "Génération de message SWIFT dans SAB")
        End If
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    

End Function

Private Sub cmdDetail_LAB_Quit_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdContext_Quit
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdEVE_Annulation_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
cboGOSEVENAT.Clear
cboGOSEVENAT.AddItem "Ann  (Annulation)"
cboGOSEVENAT.ListIndex = 0
Call cbo_Scan(oldYGOSEVE0.GOSEVEGSRV, cboGOSEVEGSRV)

cboGOSEVEGSRV.Visible = False
txtGOSEVETXT.Text = ""

Me.Enabled = True: Me.MousePointer = 0
fraEVE_S.Enabled = True
fraEVE.Visible = True

cmdEVE_Set_Ok 3

fraEVE.ZOrder 0
txtGOSEVETXT.SetFocus
End Sub

Private Sub cmdEVE_Clôture_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cboGOSEVENAT.Clear
cboGOSEVENAT.AddItem "Clo  (Clôture)"
cboGOSEVENAT.ListIndex = 0
'Call cbo_Scan(oldYGOSEVE0.GOSEVEGSRV, cboGOSEVEGSRV)
chkGOSEVEGSRV = "0"
cboGOSEVEGSRV.Enabled = False
Call cbo_Scan(oldYGOSDOS0.GOSDOSGSRV, cboGOSEVEGSRV)

txtGOSEVETXT.Text = ""
Me.Enabled = True: Me.MousePointer = 0

fraEVE_S.Enabled = True
cmdEVE_Set_Ok 3

fraEVE.Visible = True: fraEVE.ZOrder 0

txtGOSEVETXT.SetFocus

End Sub

Private Sub cmdEVE_Dupliquer_Click()

On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
cboGOSEVENAT.Clear
cboGOSEVENAT.AddItem "Swi=  (duplication Swift)"
cboGOSEVENAT.ListIndex = 0
Call cbo_Scan(oldYGOSEVE0.GOSEVEGSRV, cboGOSEVEGSRV)

duplic_YGOSEVE0 = oldYGOSEVE0
'cboGOSEVEGSRV.Visible = False
txtGOSEVETXT.Text = Mid$(oldYGOSEVE0.GOSEVETXT, 60, Len(oldYGOSEVE0.GOSEVETXT) - 59)
lblGOSEVEUAMJ = Mid$(oldYGOSEVE0.GOSEVETXT, 1, 56)

fraEVE_S.Enabled = True
Me.Enabled = True: Me.MousePointer = 0
cmdEVE_Set_Ok 2

fraEVE.Visible = True: fraEVE.ZOrder 0

txtGOSEVETXT.SetFocus

End Sub

Private Sub cmdEVE_Ignore_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

X = MsgBox("Confirmez-vous l'annulation de cet événement ?", vbQuestion + vbYesNo, Me.Caption)

If X = vbYes Then
    mYGOSDOS0_Fct = ""
    newYGOSDOS0 = oldYGOSDOS0
    mYGOSEVE0_Fct = "Update"
    newYGOSEVE0 = oldYGOSEVE0
    newYGOSEVE0.GOSEVESTAE = "A"
    blnYGOSDOS0_Update = True
    fraEVE.Visible = False
    
    If oldYGOSEVE0.GOSEVESWID > 0 Then
    
        X = "select * from " & paramIBM_Library_SABSPE & ".YSWILNK0 " _
            & " where SWILNKSWID = " & oldYGOSEVE0.GOSEVESWID & " and SWILNKAPPC = 'GOS'"
        Set rsSab = cnsab.Execute(X)
        
        If Not rsSab.EOF Then
            Call rsYSWILNK0_GetBuffer(rsSab, oldYSWILNK0)
            mYSWILNK0_Fct = "Delete"
            'mYSWILNK0_Fct = "Update"
            'newYSWILNK0 = oldYSWILNK0
            'newYSWILNK0.SWILNKSTA = "A"
        End If
        
    End If
    
    fraMail_Confirm
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdEVE_Invalidation_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
cboGOSEVENAT.Clear
If oldYGOSDOS0.GOSDOSSTAD = "C" Then
    cboGOSEVENAT.AddItem "AnnC  (annul Clôture)"
Else
    If oldYGOSDOS0.GOSDOSSTAG = "V" Then
        cboGOSEVENAT.AddItem "AnnV  (annul Validation)"
    Else
        cboGOSEVENAT.AddItem "AnnR  (annul Rejet)"
    End If
End If
cboGOSEVENAT.ListIndex = 0
Call cbo_Scan(oldYGOSEVE0.GOSEVEGSRV, cboGOSEVEGSRV)

'cboGOSEVEGSRV.Visible = False
txtGOSEVETXT.Text = ""

fraEVE_S.Enabled = True
Me.Enabled = True: Me.MousePointer = 0
cmdEVE_Set_Ok 3

fraEVE.Visible = True: fraEVE.ZOrder 0

txtGOSEVETXT.SetFocus

End Sub

Private Sub cmdEVE_New_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass

fgEVE.Enabled = False
fraEVE_S.Enabled = True
fraPJ.Visible = False

txtGOSEVEECHD.Enabled = blnHab_YGOSEVE0_New
cboGOSEVENAT.Clear

chkGOSEVEGSRV = "0"
chkGOSEVEGSRV.Enabled = blnHab_YGOSEVE0_New
cboGOSEVEGSRV.Enabled = False
cboGOSEVEGSRV.Visible = True

cboGOSEVENAT.AddItem "Note (complément d'informations)"
cboGOSEVENAT.AddItem "PJ** (pièce jointe)"
cboGOSEVENAT.AddItem "Mail (envoi par courriel)"

If currentSSIWINUNIT = oldYGOSDOS0.GOSDOSISRV Then
    Call cbo_Scan(oldYGOSDOS0.GOSDOSGSRV, cboGOSEVEGSRV)
    'chkGOSEVEGSRV.Enabled = blnHab_YGOSEVE0_New
Else
    Call cbo_Scan(oldYGOSDOS0.GOSDOSISRV, cboGOSEVEGSRV)
    'If currentSSIWINUNIT = oldYGOSDOS0.GOSDOSGSRV Then chkGOSEVEGSRV.Enabled = blnHab_YGOSEVE0_New
End If

If blnHab_YGOSEVE0_New Then
    cboGOSEVENAT.AddItem "SwFR (Emission d'un MT*99)"
    cboGOSEVENAT.AddItem "SwGB (Emission d'un MT*99)"
    If arrHab(3) Then lblGOSEVEECHD.BackColor = vbGreen
End If

'Call cbo_Scan(oldYGOSDOS0.GOSDOSGSRV, cboGOSEVEGSRV)

cboGOSEVENAT.ListIndex = -1

X = oldYGOSDOS0.GOSDOSECHD
If X >= DSys Then
    Call DTPicker_Set(txtGOSEVEECHD, X) '
    lblGOSEVEECHD.BackColor = mColor_G1
Else
    Call DTPicker_Set(txtGOSEVEECHD, DSys) '
    lblGOSEVEECHD.BackColor = mColor_W0
End If



txtGOSEVETXT.Text = ""
Me.Enabled = True: Me.MousePointer = 0

fraEVE.BackColor = fraDetail.BackColor
fraEVE.Visible = True: fraEVE.ZOrder 0

'===========
cmdEVE_Set
'===========

txtGOSEVETXT.SetFocus
End Sub


Private Sub cmdEVE_Ok_àClôturer_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

blnGOSDOSSTAD_C = False
blnGOSDOSSTAD_X = True
cmdEVE_Ok_Control

Me.Enabled = True: Me.MousePointer = 0
fgEVE.Enabled = True

End Sub

Private Sub cmdEVE_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

blnGOSDOSSTAD_C = False
blnGOSDOSSTAD_X = False
cmdEVE_Ok_Control

Me.Enabled = True: Me.MousePointer = 0
fgEVE.Enabled = True

End Sub

Private Sub cmdEVE_Ok_Clôture_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

blnGOSDOSSTAD_C = True
blnGOSDOSSTAD_X = False
cmdEVE_Ok_Control

Me.Enabled = True: Me.MousePointer = 0
fgEVE.Enabled = True

End Sub

Private Sub cmdEVE_Quit_Click()
cmdContext_Quit
End Sub

Private Sub cmdEVE_Rejet_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
cboGOSEVENAT.Clear
cboGOSEVENAT.AddItem "Rej  (Rejet)"
cboGOSEVENAT.ListIndex = 0
Call cbo_Scan(oldYGOSEVE0.GOSEVEGSRV, cboGOSEVEGSRV)

'cboGOSEVEGSRV.Visible = False
txtGOSEVETXT.Text = ""

fraEVE_S.Enabled = True
Me.Enabled = True: Me.MousePointer = 0
cmdEVE_Set_Ok 3

fraEVE.Visible = True: fraEVE.ZOrder 0

txtGOSEVETXT.SetFocus

End Sub

Private Sub cmdEVE_Restauration_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass

cboGOSEVENAT.Clear
cboGOSEVENAT.AddItem "Res*  (Restauration)"
cboGOSEVENAT.ListIndex = 0
Call cbo_Scan(oldYGOSEVE0.GOSEVEGSRV, cboGOSEVEGSRV)

fraEVE_S.Enabled = True
fraEVE.Visible = True: fraEVE.ZOrder 0

txtGOSEVETXT.Text = ""
Me.Enabled = True: Me.MousePointer = 0
cmdEVE_Reset
cmdEVE_Ok.Visible = arrHab(3)

txtGOSEVETXT.SetFocus
End Sub

Private Sub cmdEVE_Swift_Ok_Click()
Dim xSql As String, X As String, blnOk As Boolean
If Trim(cboEVE_Swift_MT) = "" Then
    Call MsgBox("Préciser le TYPE de message", vbQuestion, "GOS_DOS : émission d'un message SWIFT")
Else
    If Trim(cboEVE_Swift_20) = "" Then
        Call MsgBox("Préciser NOTRE référence", vbQuestion, "GOS_DOS : émission d'un message SWIFT")
    Else

        If Trim(txtEVE_Swift_21) = "" Then
            Call MsgBox("Préciser LEUR référence", vbQuestion, "GOS_DOS : émission d'un message SWIFT")
        Else
            X = Trim(cboEVE_Swift_BIC)
            If X = "" Then
                Call MsgBox("Préciser le BIC du destinataire", vbQuestion, "GOS_DOS : émission d'un message SWIFT")
            Else
                blnOk = True
                
                If X = oldYGOSDOS0.GOSDOSWBIC Then
                    xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIBKU0 " _
                         & " where SWIBKUBIC = '" & X & "'"
                    Set rsSab = cnsab.Execute(xSql)
                    If rsSab.EOF Then
                        ''blnOk = False
                        Call MsgBox("Contrôle / SAB(ZSWIBKU0)  : pas en RMA avec ce BIC ", vbQuestion, "GOS_DOS : émission d'un message SWIFT")
                  End If
                Else
                    xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIBIC0 " _
                         & " left outer join " & paramIBM_Library_SAB & ".ZSWIBKU0 on SWIBICBIC = SWIBKUBIC " _
                         & " where SWIBICBIC = '" & X & "'"
                    Set rsSab = cnsab.Execute(xSql)
    
                    If rsSab.EOF Then
                        Call MsgBox("BIC inconnu", vbCritical, "GOS_DOS : émission d'un message SWIFT")
                        blnOk = False
                    End If
                End If
                
                If blnOk Then
                    If Trim(cboEVE_Swift_MT) <> "999" And IsNull(rsSab("SWIBKUBIC")) Then
                        Call MsgBox("Ne ne sommes pas en clé avec cette banque, faire un MT 999", vbInformation, "GOS_DOS : émission d'un message SWIFT")
                    Else
                        fraEVE_Swift.Visible = False
                        Call lstParam_GOSDOSLABK_Load(Mid$(cboGOSEVENAT, 1, 4), "fgModèle")
                        fgModèle.ZOrder 0
                        fgModèle.Visible = True
                        
                    End If
                End If
            End If
        End If
    End If
End If


End Sub

Private Sub cmdEVE_Validation_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass
cboGOSEVENAT.Clear
cboGOSEVENAT.AddItem "Val  (Validation)"
cboGOSEVENAT.ListIndex = 0
'Call cbo_Scan(oldYGOSEVE0.GOSEVEGSRV, cboGOSEVEGSRV)
chkGOSEVEGSRV = "0"
cboGOSEVEGSRV.Enabled = False
Call cbo_Scan(oldYGOSDOS0.GOSDOSGSRV, cboGOSEVEGSRV)

txtGOSEVETXT.Text = ""
Me.Enabled = True: Me.MousePointer = 0

fraEVE_S.Enabled = True
cmdEVE_Set_Ok 3

fraEVE.Visible = True: fraEVE.ZOrder 0

txtGOSEVETXT.SetFocus
End Sub

Private Sub cmdGOSDOSPJ_Click()
Dim xSql As String
Me.Enabled = False: Me.MousePointer = vbHourglass

mfilDoc_Path = filDoc.path

New_YBIATAB0.BIATABID = "GOSDOSPJ**"
New_YBIATAB0.BIATABK1 = usrName_UCase 'currentSSIWINUNIT
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = mfilDoc_Path
If Not blnfilDoc_Path Then
    blnfilDoc_Path = True
    Parametrage_New
Else
    
    Old_YBIATAB0 = New_YBIATAB0
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
        & "where BIATABID = '" & New_YBIATAB0.BIATABID & "' and BIATABK1 = '" & New_YBIATAB0.BIATABK1 & "' and BIATABK2 = '" & New_YBIATAB0.BIATABK2 & "'"
    
    Set rsSab = cnsab.Execute(xSql)
    If Not rsSab.EOF Then
        Old_YBIATAB0.BIATABTXT = rsSab("BIATABTXT")
        Parametrage_Update
    End If
            
End If
cmdGOSDOSPJ.Visible = False
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdList_Add_Click()
On Error Resume Next
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass
X = "select * from " & paramIBM_Library_SABSPE & ".YSWILNK0 " _
     & " where SWILNKSWID = " & oldYSWILNK0.SWILNKSWID & " and SWILNKAPPC = 'GOS' and SWILNKSTA = ''"
Set rsSab = cnsab.Execute(X)

If Not rsSab.EOF Then
    X = MsgBox("Ce message swift est déjà affecté au dossier n° " & rsSab("SWILNKAPPN") & vbCrLf & "Voulez-vous l'affecter à un autre dossier ?", vbQuestion + vbYesNo + vbDefaultButton2, "BIA_GOS : contrôle")
    If X = vbNo Then GoTo Exit_sub
End If


'fraList.Visible = False
fgEVE.Enabled = False
fraEVE_S.Enabled = True
fraPJ.Visible = False

Me.Enabled = True: Me.MousePointer = 0

fraEVE.Visible = True: fraEVE.ZOrder 0

fraEVE_S.Visible = True
cboGOSEVENAT.Clear
cboGOSEVENAT.AddItem "Swi+"  'm999_YGOSDOS0.GOSDOSWMTK & m999_YGOSDOS0.GOSDOSWES
cboGOSEVENAT.ListIndex = 0

txtGOSEVETXT.Text = ""
txtGOSEVETXT.SetFocus
If oldYGOSDOS0.GOSDOSSTAD <> " " Then
    X = MsgBox("Ce dossier est clos, voulez_vous continuer ?", vbYesNo, "BIA_GOS : cmdList_Add")
    If X <> vbYes Then GoTo Exit_sub
End If

cmdEVE_Ok.Visible = arrHab(2)

mYSWILNK0_Fct = "New"



tabDetail.Tab = 1
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdList_Display_Click()
Dim xSql As String

xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSDOS0 where GOSDOSIDD = " & Val(Trim(txtList_Add))
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then
    V = rsYGOSDOS0_GetBuffer(rsSab, oldYGOSDOS0)
    fraDetail_LAB_Display_Dossier
Else
    Call MsgBox("Dossier inconnu : " & txtList_Add, vbCritical, "BIA_GOS : cmdList_Display")
End If

End Sub

Private Sub cmdList_Ignore_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

'If newYSWISAB0.SWISABKSRV <> Mid$(cboList_SWISABKSRV, 1, 3) Then cmdList_SWISABKSRV_Click

mYSWISAB0_Fct = "Update"
newYSWISAB0 = oldYSWISAB0
newYSWISAB0.SWISABXEVE = "*"
If oldYSWISAB0.SWISABK20 <> " " Then newYSWISAB0.SWISABK20 = "I"
If oldYSWISAB0.SWISABKPDE <> " " Then newYSWISAB0.SWISABKPDE = "I"
If oldYSWISAB0.SWISABK999 <> " " Then newYSWISAB0.SWISABK999 = "I"


rsYGOSEVE0_Init newYGOSEVE0
newYGOSEVE0.GOSEVESWID = newYSWISAB0.SWISABSWID
newYGOSEVE0.GOSEVENAT = "*I"
newYGOSEVE0.GOSEVETXT = "Vu"
mYGOSEVE0_Fct = "New"

blnYGOSDOS0_Display = True
cmdYGOSDOS0_Update "", mYGOSEVE0_Fct, "", mYSWISAB0_Fct, ""

Call cmdSelect_Reset_Post_Update

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdList_New_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass
X = "select * from " & paramIBM_Library_SABSPE & ".YSWILNK0 " _
     & " where SWILNKSWID = " & oldYSWISAB0.SWISABSWID & " and SWILNKAPPC = 'GOS' and SWILNKSTA = ''"
Set rsSab = cnsab.Execute(X)

If Not rsSab.EOF Then
    X = MsgBox("Ce message swift est déjà affecté au dossier n° " & rsSab("SWILNKAPPN") & vbCrLf & "Voulez-vous créer un autre dossier ?", vbQuestion + vbYesNo + vbDefaultButton2, "BIA_GOS : contrôle")
    If X = vbNo Then GoTo Exit_sub
End If

'If newYSWISAB0.SWISABKSRV <> Mid$(cboList_SWISABKSRV, 1, 3) Then cmdList_SWISABKSRV_Click

oldYGOSDOS0 = m999_YGOSDOS0
tabDetail.Tab = 0
fraList.Visible = False
fgEVE.Visible = False
fraDetail_LAB_Init

mYSWILNK0_Fct = "New"

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdList_Quit_Click()
cmdContext_Quit

End Sub

Private Sub cmdList_SAB_Annulation_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass

V = ZSWIENA0_Read

If Not IsNull(V) Then
    Call MsgBox(V, vbInformation, "BIA_GOS : cmdList_SAB_Annulation_Click")
    GoTo Exit_sub
End If

'If newYSWISAB0.SWISABKSRV <> Mid$(cboList_SWISABKSRV, 1, 3) Then cmdList_SWISABKSRV_Click

mZSWIENA0_Fct = "Update"
newZSWIENA0 = oldZSWIENA0
newZSWIENA0.SWIENACET = "1"



mYSWISAB0_Fct = "Update"
newYSWISAB0 = oldYSWISAB0
newYSWISAB0.SWISABZSWI = oldZSWIENA0.SWIENAINT
newYSWISAB0.SWISABXEVE = "*"
If oldYSWISAB0.SWISABK20 <> " " Then newYSWISAB0.SWISABK20 = "A"
If oldYSWISAB0.SWISABKPDE <> " " Then newYSWISAB0.SWISABKPDE = "A"


rsYGOSEVE0_Init newYGOSEVE0
newYGOSEVE0.GOSEVESWID = newYSWISAB0.SWISABSWID
newYGOSEVE0.GOSEVENAT = "*A"
newYGOSEVE0.GOSEVETXT = "annulation de l'enregistrement dans SAB073/ZSWIENA0 " & oldZSWIENA0.SWIENAINT
mYGOSEVE0_Fct = "New"

blnYGOSDOS0_Display = True
cmdYGOSDOS0_Update "", mYGOSEVE0_Fct, "", mYSWISAB0_Fct, mZSWIENA0_Fct

Call cmdSelect_Reset_Post_Update

'oldYGOSEVE0.GOSEVESWID

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdList_SAB_Modification_Click()
Dim V, X As String, K As Long, xSql As String
Me.Enabled = False: Me.MousePointer = vbHourglass
V = ZSWIENA0_Read

If Not IsNull(V) Then
    Call MsgBox(V, vbInformation, "BIA_GOS : cmdList_SAB_Modification_Click")
    GoTo Exit_sub
End If


X = InputBox("Saisir une nouvelle référence ( maximum 16 caractères)" & vbCrLf & oldZSWIENA0.SWIENAREF & vbCrLf & "(effacer la référence pour abandonner la mise à jour)", "Modification de la référence d'un message SWIFT reçu", Trim(oldZSWIENA0.SWIENAREF))
If Trim(X) = "" Then
    Call MsgBox("Modification abandonnée", vbInformation, "Modification de la référence d'un message SWIFT reçu")
    GoTo Exit_sub
End If
X = UCase$(Trim(X))
If Len(X) > 16 Then
    Call MsgBox("Modification abandonnée : la référence a plus de 16 caractères ", vbInformation, "Modification de la référence d'un message SWIFT reçu")
    GoTo Exit_sub
End If




xSql = " select count(*) as Tally from " & paramIBM_Library_SAB & ".ZSWIENA0 where SWIENAREF = '" & X & "'" _
     & " and SWIENAEME= '" & oldZSWIENA0.SWIENAEME & "' And SWIENAMES = " & oldZSWIENA0.SWIENAMES
Set rsSab = cnsab.Execute(xSql)

K = rsSab("Tally")
If K > 0 Then
    Call MsgBox("Cette référence est déjà utilisée (en cours ZSWIENA0)", vbInformation, "Modification de la référence d'un message SWIFT reçu")
    GoTo Exit_sub
End If

xSql = " select count(*) as Tally from " & paramIBM_Library_SAB & ".ZSWIMEA0 where SWIMEAREF = '" & X & "'" _
     & " and SWIMEAEME= '" & oldZSWIENA0.SWIENAEME & "' And SWIMEAMES = " & oldZSWIENA0.SWIENAMES
Set rsSab = cnsab.Execute(xSql)

K = rsSab("Tally")
If K > 0 Then
    Call MsgBox("Cette référence est déjà utilisée et traitée (ZSWIMEA0)", vbInformation, "Modification de la référence d'un message SWIFT reçu")
    GoTo Exit_sub
End If

'================================================================================================
'If newYSWISAB0.SWISABKSRV <> Mid$(cboList_SWISABKSRV, 1, 3) Then cmdList_SWISABKSRV_Click


mZSWIENA0_Fct = "Update"
newZSWIENA0 = oldZSWIENA0
newZSWIENA0.SWIENAREF = X



mYSWISAB0_Fct = "Update"
newYSWISAB0 = oldYSWISAB0
newYSWISAB0.SWISABZSWI = oldZSWIENA0.SWIENAINT
newYSWISAB0.SWISABXEVE = "*"
If oldYSWISAB0.SWISABK20 <> " " Then newYSWISAB0.SWISABK20 = "M"
If oldYSWISAB0.SWISABKPDE <> " " Then newYSWISAB0.SWISABKPDE = "M"


rsYGOSEVE0_Init newYGOSEVE0
newYGOSEVE0.GOSEVESWID = newYSWISAB0.SWISABSWID
newYGOSEVE0.GOSEVENAT = "*M"
newYGOSEVE0.GOSEVETXT = oldZSWIENA0.SWIENAREF & " => " & newZSWIENA0.SWIENAREF & " : modification de la référence de l'enregistrement dans SAB073/ZSWIENA0 " & oldZSWIENA0.SWIENAINT
mYGOSEVE0_Fct = "New"

blnYGOSDOS0_Display = True
cmdYGOSDOS0_Update "", mYGOSEVE0_Fct, "", mYSWISAB0_Fct, mZSWIENA0_Fct

Call cmdSelect_Reset_Post_Update

'oldYGOSEVE0.GOSEVESWID

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdList_SWISABKSRV_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass


mYSWISAB0_Fct = "Update"
newYSWISAB0 = oldYSWISAB0
newYSWISAB0.SWISABXEVE = "="
newYSWISAB0.SWISABKSRV = Mid$(cboList_SWISABKSRV, 1, 3)



rsYGOSEVE0_Init newYGOSEVE0
newYGOSEVE0.GOSEVESWID = newYSWISAB0.SWISABSWID
newYGOSEVE0.GOSEVENAT = "*>"
newYGOSEVE0.GOSEVEGSRV = newYSWISAB0.SWISABKSRV
newYGOSEVE0.GOSEVETXT = "affectation au service " & cboList_SWISABKSRV
mYGOSEVE0_Fct = "New"

blnYGOSDOS0_Display = True
cmdYGOSDOS0_Update "", mYGOSEVE0_Fct, "", mYSWISAB0_Fct, ""

Call cmdSelect_Reset_Post_Update
'oldYGOSEVE0.GOSEVESWID

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdMail_MT_CC_Click()
lstMail_MT_Message.Visible = False
lstMail_MT_To.Visible = False
lstMail_MT_CC.Visible = True: lstMail_MT_CC.ZOrder 0


End Sub

Private Sub cmdMail_MT_Click()
If blnGOSEVE_Mail Then
    Call MsgBox("Evènement 'envoi en cours'", vbQuestion, "BIA_GOS : mail")
Else
    cmdMail_MT_NOK.Visible = False
    fraMail_MT.Visible = True
    fraMail_MT.ZOrder 0
    
    If Trim(txtMail_MT_To) = "" Then lstMail_MT_To.Visible = True
End If
End Sub

Private Sub cmdMail_MT_Message_Click()
Dim X As String
lstMail_MT_To.Visible = False
lstMail_MT_CC.Visible = False

lstMail_MT_Message.Clear

X = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where GOSEVEIDD = -1 and GOSEVEGSRV = '" & currentSSIWINUNIT & "'" _
     & " and GOSEVENAT = 'Mail' order by SUBSTRING(GOSEVETXT , 1 , 10)"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    X = Trim(rsSab("GOSEVETXT"))
    lstMail_MT_Message.AddItem X
    
    rsSab.MoveNext
Loop
lstMail_MT_Message.Visible = True

End Sub

Private Sub cmdMail_MT_NOK_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
If blnYGOSDOS0_Update Then
    blnYGOSDOS0_Display = True
    cmdYGOSDOS0_Update mYGOSDOS0_Fct, mYGOSEVE0_Fct, mYSWILNK0_Fct, mYSWISAB0_Fct, ""
End If
fraMail_MT.Visible = False
blnGOSEVE_Mail = False

Call cmdSelect_Reset_Post_Update
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdMail_MT_Ok_Click()
Dim X As String, Xto As String, xCC As String
Dim V, blnOk As Boolean


Me.Enabled = False: Me.MousePointer = vbHourglass
xCC = "": Xto = ""

X = Trim(txtMail_MT_To)
If X = "" Then
    Call MsgBox("préciser le destinataire du mail", vbQuestion, "BIA_GOS : contrôle destinataire mail")
    GoTo Exit_sub
End If

V = mailAdresse_Production_Control(X, Xto)
If Not IsNull(V) Then
    Call MsgBox("destinataire inconnu : " & V, vbQuestion, "Contrôle destinataire mail  ")
    GoTo Exit_sub
End If

X = Trim(txtMail_MT_CC)
If X <> "" Then

    V = mailAdresse_Production_Control(X, xCC)
    If Not IsNull(V) Then
        Call MsgBox("destinataire en copie inconnu : " & V, vbQuestion, "Contrôle destinataire mail")
    GoTo Exit_sub
    End If
End If
'_________________________________________________________________________________________
If Not blnGOSEVE_Mail Then
    blnOk = False
    If fraDetail.Visible Then
        If currentSSIWINUNIT = oldYGOSDOS0.GOSDOSISRV _
        Or currentSSIWINUNIT = oldYGOSDOS0.GOSDOSGSRV Then
            Call cmdSendMail_YGOSDOS0(Xto, xCC)
            blnOk = True
        End If
    End If
    
    If Not blnOk Then Call cmdSendMail_MT(Xto, xCC, Trim(txtMail_MT_Message))
Else
    If blnYGOSDOS0_Update Then
    
        If IsNull(cmdYGOSDOS0_Update(mYGOSDOS0_Fct, mYGOSEVE0_Fct, mYSWILNK0_Fct, mYSWISAB0_Fct, "")) Then
        
            oldYGOSDOS0 = newYGOSDOS0
            X = " where GOSEVEIDD = " & oldYGOSDOS0.GOSDOSIDD & " order by GOSEVEIDE"
            arrYGOSEVE0_SQL X
    
            If IsNull(cmdSendMail_YGOSDOS0(Xto, xCC)) Then
                newYGOSEVE0.GOSEVEIDE = 0
                newYGOSEVE0.GOSEVESWID = 0
                newYGOSEVE0.GOSEVESTAE = ""
                newYGOSEVE0.GOSEVENAT = "Mail"
                newYGOSEVE0.GOSEVETXT = "=> " & Trim(txtMail_MT_To) & " + " & Trim(txtMail_MT_CC) & vbCrLf & vbCrLf & Trim(txtMail_MT_Message)
                If mYGOSDOS0_Fct <> "New" Then blnYGOSDOS0_Display = True
                If IsNull(cmdYGOSDOS0_Update("", "New", "", "", "")) Then
                    'framail_mt.Visible = False
                End If
            End If
        End If
    Else
        Call cmdSendMail_YGOSDOS0(Xto, xCC)
    End If
    blnGOSEVE_Mail = False
End If

'_________________________________________________________________________________________
lstMail_MT_To.Visible = False
lstMail_MT_CC.Visible = False
lstMail_MT_Message.Visible = False

fraMail_MT.Visible = False
blnGOSEVE_Mail = False

Call cmdSelect_Reset_Post_Update

Exit_sub:

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdMail_MT_Quit_Click()
lstMail_MT_To.Visible = False
lstMail_MT_CC.Visible = False
fraMail_MT.Visible = False
cmdContext_Quit
blnGOSEVE_Mail = False

End Sub

Private Sub cmdMail_MT_To_Click()
lstMail_MT_Message.Visible = False
lstMail_MT_CC.Visible = False
lstMail_MT_To.Visible = True

End Sub




Private Sub cmdMail_Ok_Click()
Dim X As String, Xto As String, xCC As String
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
xCC = "": Xto = ""

X = Trim(txtMail_MT_To)
If X = "" Then
    Call MsgBox("préciser le destinataire du mail", vbQuestion, "BIA_GOS : contrôle destinataire mail")
    GoTo Exit_sub
End If

V = mailAdresse_Production_Control(X, Xto)
If Not IsNull(V) Then
    Call MsgBox("destinataire inconnu : " & V, vbQuestion, "Contrôle destinataire mail  (paramétrage SAB)")
    GoTo Exit_sub
End If

X = Trim(txtMail_MT_CC)
If X <> "" Then

    V = mailAdresse_Production_Control(X, xCC)
    If Not IsNull(V) Then
        Call MsgBox("destinataire en copie inconnu : " & V, vbQuestion, "Contrôle destinataire mail  (paramétrage SAB)")
    GoTo Exit_sub
    End If
End If


If blnYGOSDOS0_Update Then

    If IsNull(cmdYGOSDOS0_Update(mYGOSDOS0_Fct, mYGOSEVE0_Fct, mYSWILNK0_Fct, mYSWISAB0_Fct, "")) Then
    
        oldYGOSDOS0 = newYGOSDOS0
        X = " where GOSEVEIDD = " & oldYGOSDOS0.GOSDOSIDD & " order by GOSEVEIDE"
        arrYGOSEVE0_SQL X

        If IsNull(cmdSendMail_YGOSDOS0(Xto, xCC)) Then
            newYGOSEVE0.GOSEVEIDE = 0
            newYGOSEVE0.GOSEVESWID = 0
            newYGOSEVE0.GOSEVESTAE = ""
            newYGOSEVE0.GOSEVENAT = "Mail"
            newYGOSEVE0.GOSEVETXT = "=> " & Trim(txtMail_MT_To) & " + " & Trim(txtMail_MT_CC) & vbCrLf & vbCrLf & Trim(txtMail_MT_Message)
            If mYGOSDOS0_Fct <> "New" Then blnYGOSDOS0_Display = True
            If IsNull(cmdYGOSDOS0_Update("", "New", "", "", "")) Then
                'framail_mt.Visible = False
            End If
        End If
    End If
Else
    Call cmdSendMail_YGOSDOS0(Xto, xCC)
End If

Exit_sub:

fraMail_MT.Visible = False
blnGOSEVE_Mail = False

Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Function cmdSendMail_YGOSDOS0(lMail_To As String, lMail_CC As String)
Dim wSendMail As typeSendMail
Dim xDétail_D As String, xHeader_D As String
Dim xDétail_W As String, xHeader_W As String, mbgColor As String, mbgColor2 As String
Dim K As Long, htmlFontColor_K As String, K2 As Long
Dim X As String, xGOSEVETXT As String
Dim X0 As String, X1 As String
Dim mForeColor1 As String, mForeColor2 As String, mForeColor3 As String
Dim xSql As String
On Error GoTo Error_Handler:

cmdSendMail_YGOSDOS0 = Null
'-----------------------------------------------------------------------------------
X0 = "GOS n° : " & newYGOSDOS0.GOSDOSIDD
X1 = dateImp10_S(newYGOSDOS0.GOSDOSIAMJ)
xHeader_D = "<TR>" _
         & "<TD bgcolor=#0090A0 width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'><Font color=#FFFFFF><B>" & X0 & "</B/TD>" _
         & "<TD bgcolor=#0090A0 width=800 height=7><span style='font-size:10.0pt;font-family:Calibri'><Font color=#FFFFFF>" & X1 & "</TD>" _
        & "</TR>"

xDétail_D = ""
mbgColor = "bgcolor = #00B0C0"
mbgColor2 = "bgcolor = #FFFFFF"

xDétail_D = xDétail_D _
     & "<TD " & mbgColor & " width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_White & "Service initiateur" & "</TD>" _
     & "<TD " & mbgColor2 & " width=800 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Blue & Trim(arrService_Lib(Val(Mid$(newYGOSDOS0.GOSDOSISRV, 2, 2)))) & "</TD>" _
     & "</TR>" _
     & "<TD " & mbgColor & " width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_White & "Motif" & "</TD>" _
     & "<TD " & mbgColor2 & " width=800 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Blue & newYGOSDOS0.GOSDOSLABK & "</TD>" _
     & "<TR>"
If Trim(newYGOSDOS0.GOSDOSCLI) <> "" Then
    xDétail_D = xDétail_D _
     & "<TD " & mbgColor & " width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_White & "Client" & "</TD>" _
     & "<TD " & mbgColor2 & " width=800 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Blue & newYGOSDOS0.GOSDOSCLI & " - " & libGOSDOSCLI & "</TD>" _
     & "</TR>" _
     & "<TD " & mbgColor & " width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_White & "Responsable commercial" & "</TD>" _
     & "<TD " & mbgColor2 & " width=800 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Blue & Trim(arrRCOM_Lib(Val(Mid$(newYGOSDOS0.GOSDOSRCOM, 2, 2)))) & "</TD>" _
     & "</TR>"
End If
     
    ' & "<TD " & mbgColor & " width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_white & "Service destinataire" & "</TD>" _
    ' & "<TD " & mbgColor & " width=800 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Blue & Trim(arrService_Lib(Mid$(newYGOSDOS0.GOSDOSGSRV, 2, 2))) & "</TD>" _
    ' & "</TR>"
'-----------------------------------------------------------------------------------
If arrYGOSEVE0_Nb > 0 Then

    xDétail_D = xDétail_D _
         & "<TD  bgcolor=#0090A0 width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'><Font color=#FFFFFF>" & "date - évènement" & "</TD>" _
         & "<TD  bgcolor=#0090A0  width=800 height=7><span style='font-size:10.0pt;font-family:Calibri'><Font color=#FFFFFF>" & "Informations" & "</TD>" _
         & "</TR>"

End If


For K = 1 To arrYGOSEVE0_Nb

    xYGOSEVE0 = arrYGOSEVE0(K)
    
    mbgColor = "bgcolor = #FFFFF0" '#FAFAD2"
   
    'If Mid$(xYGOSEVE0.GOSEVENAT, 4, 1) = "*" Then
    If xYGOSEVE0.GOSEVEGSRV <> xYGOSEVE0.GOSEVEUSRV Then
        K2 = Val(Mid$(xYGOSEVE0.GOSEVEGSRV, 2, 2))
        X0 = htmlFontColor_Red & "<BR>(Service gestionnaire : " & Trim(arrService_Lib(K2)) & ")"
    Else
        X0 = ""
    End If
    
    If xYGOSEVE0.GOSEVENAT = "PJ**" Then
        If xYGOSEVE0.GOSEVESTAE <> "A" Then
            X1 = "<a href=" & Asc34 _
            & paramGOSDOS_Path_DROPI & xYGOSEVE0.GOSEVEIDD _
            & "\" & xYGOSEVE0.GOSEVEIDD & "_" & xYGOSEVE0.GOSEVEIDE _
            & "." & Trim(fileName_Extension(xYGOSEVE0.GOSEVETXT)) _
            & Asc34 & ">" & Trim(xYGOSEVE0.GOSEVETXT)
        Else
            X1 = Trim(xYGOSEVE0.GOSEVETXT)
        End If
        
    Else
        If xYGOSEVE0.GOSEVENAT = "Swi>" And xYGOSEVE0.GOSEVESWID > 0 Then
            X1 = cmdSendMail_rText_Line(Mid$(xYGOSEVE0.GOSEVETXT, 1, 56))
        Else
            X1 = cmdSendMail_rText_Line(Trim(xYGOSEVE0.GOSEVETXT))
        End If
    End If
    
    If xYGOSEVE0.GOSEVESTAE = "A" Then
        X0 = " **** Annulé"
        mForeColor1 = htmlFontColor_Gray
        mForeColor2 = htmlFontColor_Gray
        mForeColor3 = htmlFontColor_Gray
    Else
        mForeColor1 = htmlFontColor_Blue
        mForeColor2 = htmlFontColor_Magenta
        Select Case Trim(xYGOSEVE0.GOSEVENAT)
            Case "Sus*":
                Select Case newYGOSDOS0.GOSDOSSTAG
                    Case "R": mForeColor3 = "<Font color = #FFFF00><B>": mbgColor = "bgcolor = #FF0000"
                    Case "V": mForeColor3 = "<Font color = #000080><B>": mbgColor = "bgcolor = #80FF80"
                
                    Case Else: mForeColor3 = "<Font color = #000080><B>": mbgColor = "bgcolor = #FFFF00"
                End Select
                
            Case "Note": mForeColor3 = "<Font color = #000080><B>": mbgColor = "bgcolor = #FFFFA0"
            Case "Res*", "AnnV", "AnnR", "AnnC": mForeColor3 = htmlFontColor_Magenta & "<B>": mbgColor = "bgcolor = #FFFFA0"
            Case "Mail": mForeColor3 = "<Font color = #4040FF>"
            Case "PJ**", "Swi+":
                        If xYGOSEVE0.GOSEVESTAE <> "A" Then
                            mForeColor3 = "<Font color = #000080>": mbgColor = "bgcolor = #F6FCFF" 'AFEEEE"
                            htmlFontColor_rText = htmlFontColor_Blue
                        End If
            Case "Swi>":
                        htmlFontColor_rText = htmlFontColor_Green
                        If xYGOSEVE0.GOSEVESWID > 0 Then
                            mForeColor3 = htmlFontColor_Green: mbgColor = "bgcolor = #F4FFF4"
                        Else
                            mForeColor3 = htmlFontColor_Red: mbgColor = "bgcolor = #E0E0E0"
                        End If
                        
            Case "Val": mForeColor3 = "<Font color = #000080><B>": mbgColor = "bgcolor = #80FF80"
            Case "Rej": mForeColor3 = "<Font color = #FFFF00><B>": mbgColor = "bgcolor = #FF0000"
                        '''mForeColor3 = "<Font color = #000080><B>": mbgColor = "bgcolor = #FFC0CB"
            Case "Clo": mForeColor3 = "<Font color = #000080><B>": mbgColor = "bgcolor = #E0E0E0"
           Case Else: mForeColor3 = htmlFontColor_Magenta
        End Select
   End If
   

'___________________________________________________________


    X = "<span style='font-size:8.0pt;font-family:Calibri'>" & mForeColor1 & " " & dateImp10_S(xYGOSEVE0.GOSEVEUAMJ) _
        & " " & timeImp8(xYGOSEVE0.GOSEVEUHMS) & " - " & xYGOSEVE0.GOSEVEUUSR & " : " _
        & "<span style='font-size:10.0pt;font-family:Calibri'>" & mForeColor2 & xYGOSEVE0.GOSEVENAT & X0
    xDétail_D = xDétail_D _
         & "<TD " & mbgColor & " width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & X & "</TD>" _
         & "<TD " & mbgColor & " width=800 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & mForeColor3 & X1 & "</B></TD>" _
         & "</TR>"
         
         
    If K = 1 Then
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABSWID = " & xYGOSEVE0.GOSEVESWID
        Set rsSab = cnsab.Execute(xSql)
        
        If Not rsSab.EOF Then
            X = "le " & dateImp10(rsSab("SWISABWAMJ")) & " " & timeImp8(rsSab("SWISABWHMS")) & " :<BR>"
        Else
            X = ""
        End If

        If newYGOSDOS0.GOSDOSWES = "S" Then
            X0 = htmlFontColor_Blue & X & htmlFontColor_Magenta & newYGOSDOS0.GOSDOSWMTK & " Sortant vers "
        Else
            X0 = htmlFontColor_Blue & X & htmlFontColor_Magenta & newYGOSDOS0.GOSDOSWMTK & " reçu de "
        End If
        
        Call arrMT_Fields_Load(newYGOSDOS0.GOSDOSWMTK)

        X = cmdSendMail_GOSDOSITOP(newYGOSDOS0.GOSDOSWID1, newYGOSDOS0.GOSDOSWIDL, newYGOSDOS0.GOSDOSWIDH, newYGOSDOS0.GOSDOSITOP)
        xDétail_D = xDétail_D _
             & "<TD bgcolor = #FFFFFF width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & X0 & newYGOSDOS0.GOSDOSWBIC & "</TD>" _
             & "<TD bgcolor = #FFFFFF width=800 height=7><span style='font-size:10.0pt;font-family:Courier New'>" & X & "</TD>" _
             & "</TR>"
    
    Else
        If xYGOSEVE0.GOSEVESWID <> 0 Then
            If xYGOSEVE0.GOSEVESTAE <> "A" Then
                X = cmdSendMail_SWISABSWID(xYGOSEVE0.GOSEVESWID, X0)
                xDétail_D = xDétail_D _
                     & "<TD bgcolor = #FFFFFF width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & X0 & "</TD>" _
                     & "<TD " & mbgColor & " width=800 height=7><span style='font-size:10.0pt;font-family:Courier New'>" & htmlFontColor_Blue & X & "</TD>" _
                    & "</TR>"
            End If
        End If
    End If
Next K

'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = "GOS_" & usrName_ULCase
wSendMail.Recipient = Trim(lMail_To)
wSendMail.CcRecipient = Trim(lMail_CC)

'wSendMail.Subject = "Gestion des opérations en suspens, dossier n° : " & newYGOSDOS0.GOSDOSIDD & " en date du " & Date & " " & Time
wSendMail.Subject = "GOS n° " & newYGOSDOS0.GOSDOSIDD & " : MT" & newYGOSDOS0.GOSDOSWMTK & " " & newYGOSDOS0.GOSDOSWBIC _
                  & "  " & Format$(newYGOSDOS0.GOSDOSWMTD, "### ### ### ###.##") & " " & newYGOSDOS0.GOSDOSWDEV & " L/réf : " & newYGOSDOS0.GOSDOSWTRN


X = Replace(Trim(txtMail_MT_Message), vbCrLf, "<BR> ") & "<BR>"

If mYGOSDOS0_Fct = "New" Then
    xGOSEVETXT = htmlFontColor_Blue & "<span style='font-size:12.0pt'>" & Replace(Trim(arrYGOSEVE0(1).GOSEVETXT), vbCrLf, "<BR> ") & "<BR><BR>"
Else
    xGOSEVETXT = "<BR>"
End If

wSendMail.Attachment = ""

mbgColor = "<body bgcolor = #FFFFFF>"

'Select Case newYGOSDOS0.GOSDOSSTAG
'    Case "V": mbgColor = "<body bgcolor = #E0FFE0>"
'    Case "R": mbgColor = "<body bgcolor = #FFE0FF>"
'    Case Else: mbgColor = "<body bgcolor = #FFFFFF>"
'End Select

'Select Case newYGOSDOS0.GOSDOSSTAD
'    Case "A", "C": mbgColor = "<body bgcolor = #E0E0E0>"
'End Select


wSendMail.Message = mbgColor _
                    & htmlFontColor_Gray & "<span style='font-size:11.0pt;font-family:Calibri'>" & X & xGOSEVETXT _
                    & "<TABLE   width=1000 border=1 cellpadding=4 ></B>" _
                    & xHeader_D _
                    & xDétail_D _
                    & "</TABLE>"

wSendMail.AsHTML = True
srvSendMail.Monitor wSendMail
Exit Function

Error_Handler:
    cmdSendMail_YGOSDOS0 = Error
    
End Function
Public Function cmdSendMail_MT(lMail_To As String, lMail_CC As String, lMail_Message As String)
Dim wSendMail As typeSendMail
Dim xDétail_D As String, xHeader_D As String
Dim xDétail_W As String, xHeader_W As String, mbgColor As String, mbgColor2 As String
Dim K As Long, htmlFontColor_K As String
Dim X As String, xGOSEVETXT As String
Dim X0 As String, X1 As String
Dim mForeColor1 As String, mForeColor2 As String, mForeColor3 As String
Dim xSql As String
On Error GoTo Error_Handler:

cmdSendMail_MT = Null
If oldYSWISAB0.SWISABSWID = 0 Then
    Call MsgBox("message inconnu dans YSWISAB0 (pb de synchronisation ?)", vbInformation, "BIA_GOS :cmdSendMail_MT")
    Exit Function
End If


'-----------------------------------------------------------------------------------
X0 = "SAB : " & oldYSWISAB0.SWISABSER & " " & oldYSWISAB0.SWISABSER & " " & oldYSWISAB0.SWISABOPEC & " " & oldYSWISAB0.SWISABOPEN

If oldYSWISAB0.SWISABWES = "S" Then
    X = " Sortant vers "
Else
    X1 = " reçu de "
End If
X1 = "Message Swift : " & oldYSWISAB0.SWISABWMTK & X & oldYSWISAB0.SWISABWBIC _
    & "  le " & dateImp10_S(oldYSWISAB0.SWISABWAMJ) & " à " & timeImp8(oldYSWISAB0.SWISABWHMS) '& "      " & oldYSWISAB0.SWISABWSTA


xHeader_D = "<TR>" _
         & "<TD bgcolor=#0090A0 width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'><Font color=#FFFFFF>" & X0 & "</TD>" _
         & "<TD bgcolor=#0090A0 width=800 height=7><span style='font-size:10.0pt;font-family:Calibri'><Font color=#FFFFFF>" & X1 & "</TD>" _
        & "</TR>"

xDétail_D = ""
mbgColor = "bgcolor = #00B0C0"
mbgColor2 = "bgcolor = #FFFFFF"

'-----------------------------------------------------------------------------------
    X = cmdSendMail_SWISABSWID(oldYSWISAB0.SWISABSWID, X0)
    xDétail_D = xDétail_D _
         & "<TD bgcolor = #FFFFFF width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & X0 & "</TD>" _
         & "<TD bgcolor = #FFFFFF width=800 height=7><span style='font-size:10.0pt;font-family:Courier New'>" & htmlFontColor_Blue & X & "</TD>" _
         & "</TR>"


'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = "BIA_GOS"
wSendMail.Recipient = Trim(lMail_To)
wSendMail.CcRecipient = Trim(lMail_CC)

wSendMail.Subject = X1
X = Replace(lMail_Message, vbCrLf, "<BR>") & "<BR>"

If mYGOSDOS0_Fct = "New" Then
    xGOSEVETXT = htmlFontColor_Blue & "<span style='font-size:12.0pt'>" & Replace(Trim(arrYGOSEVE0(1).GOSEVETXT), vbCrLf, "<BR> ") & "<BR><BR>"
Else
    xGOSEVETXT = "<BR>"
End If

wSendMail.Attachment = ""


wSendMail.Message = "<" & mbgColor & ">" _
                    & htmlFontColor_Gray & X & xGOSEVETXT _
                    & "<TABLE   width=1000 border=1 cellpadding=4 ></B>" _
                    & xHeader_D _
                    & xDétail_D _
                    & "</TABLE>"

wSendMail.AsHTML = True
srvSendMail.Monitor wSendMail
Exit Function

Error_Handler:
    cmdSendMail_MT = Error
    MsgBox Error, vbCritical, "BIA_GOS : cmdSendMail_MT"
End Function

Public Function cmdSendMail_rText_Line(lTxt As String) As String
Dim lenX As Long, K As Integer, K1 As Integer, K2 As Integer
Dim I As Integer, I1 As Integer, iSpace As Integer
Dim htmlTxt As String, blnEnd As Boolean
Dim wNb As Integer, wReturn As String, wReturnT As String
Dim blnNext As Boolean

wNb = 90
wReturn = "<BR/>"
wReturnT = "<BR/>&#160;&#160;&#160;&#160;&#160;"
htmlTxt = "" '"<pre>" 'vbCrLf

K = 1
lenX = Len(lTxt)
blnEnd = False
Do
    K1 = InStr(K, lTxt, vbCrLf)
    If K1 > 0 Then
        blnNext = False
        Do
            If K1 - K < wNb Then
                blnNext = True
                htmlTxt = htmlTxt & Mid$(lTxt, K, K1 - K) & wReturnT
            Else
                K2 = K + wNb
                For I = K2 To K2 - 15 Step -1
                    If Mid$(lTxt, I, 1) = " " Then K2 = I: Exit For
                Next I
                htmlTxt = htmlTxt & Mid$(lTxt, K, K2 - K) & wReturnT
                K = K2 + 1
            End If
        Loop Until blnNext
        K = K1 + 2
    Else
        blnEnd = True
        K1 = lenX
        blnNext = False
        Do
            If K1 - K < wNb Then
                blnNext = True
                htmlTxt = htmlTxt & Mid$(lTxt, K, K1 - K + 1) & wReturn
            Else
                K2 = K + wNb
                For I = K2 To K2 - 15 Step -1
                    If Mid$(lTxt, I, 1) = " " Then K2 = I: Exit For
                Next I
                htmlTxt = htmlTxt & Mid$(lTxt, K, K2 - K) & wReturnT
                K = K2 + 1
            End If
        Loop Until blnNext
    End If
Loop Until blnEnd

cmdSendMail_rText_Line = htmlTxt
End Function

Private Sub cmdParam_GOSDOSLABK_Add_Click()
Dim X As String, XX As String, xSql As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

blnOk = True
X = Mid$(Trim(txtParam_GOSDOSLABK) & Space(11), 1, 10)
If Trim(X) = "" Then
    blnOk = False
    Call MsgBox("Préciser le code du motif", vbCritical, "BIA_GOS : paramétrage")
Else
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
         & " where GOSEVEIDD = -1 and GOSEVEGSRV = '" & currentSSIWINUNIT & "'" _
         & " and GOSEVENAT = '" & oldParam.GOSEVENAT & "' and  SUBSTRING(GOSEVETXT , 1 , 10) = '" & X & "'"
    Set rsSab = cnsab.Execute(xSql)
    If Not rsSab.EOF Then
        blnOk = False
        Call MsgBox("Ce code existe déjà", vbCritical, "BIA_GOS : paramétrage")
End If

V = cmdParam_GOSDOSLABK_Control

If Not IsNull(V) Then
    blnOk = False
    Call MsgBox(V, vbCritical, "BIA_GOS : paramétrage")
End If

    
 If blnOk Then
    If cboParam_GOSDOSLABK_GSrv.Visible Then
        If currentSSIWINUNIT = Mid$(cboParam_GOSDOSLABK_GSrv, 1, 3) Then
            If vbYes <> MsgBox("Confirmez-vous être le service gestionnaire de ce dossier ?", vbQuestion & vbYesNo, "BIA_GOS : paramétrage") Then blnOk = False
        End If
    End If
End If
 If blnOk Then
        X = Mid$(Trim(txtParam_GOSDOSLABK) & Space(11), 1, 10)
        newYGOSEVE0 = oldParam
        XX = Space$(12)
        If oldParam.GOSEVENAT = "DOS " Then Mid$(XX, 2, 3) = Mid$(cboParam_StaC, 1, 3)
        newYGOSEVE0.GOSEVETXT = X & " " & Format(txtParam_GOSDOSLABK_J, "000") & " " & Mid$(cboParam_GOSDOSLABK_GSrv, 1, 3) & XX & Trim(txtParam_GOSDOSLABK_Lib)
        cmdYGOSDOS0_Update "", "New", "", "", ""
        
        lstParam_GOSDOSLABK_Load newYGOSEVE0.GOSEVENAT, ""
    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Function Parametrage_Delete()
Dim xSql As String
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
Private Sub cmdParam_SAA_Update_Click()
Dim X As String, XX As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

blnOk = True
X = Trim(txtParam_SAA_K2)
If X <> Trim(Old_YBIATAB0.BIATABK2) Then
    Call MsgBox("Le code du motif a été modifié," & vbCrLf & " la mise à jour n'est pas possible", vbCritical, "BIA_GOS : paramétrage SAA")

Else
    
    'New_YBIATAB0 = Old_YBIATAB0
    'New_YBIATAB0.BIATABTXT = Format(Val(txtParam_SAA_MTD), "000000000000")
    cmdParam_SAA_Control
    If IsNull(Parametrage_Update) Then lstParam_SAA_Load
End If



Me.Enabled = True: Me.MousePointer = 0

End Sub
Private Function Parametrage_New()
Dim xSql As String
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


Private Sub cmdParam_GOSDOSLABK_Delete_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

X = Trim(txtParam_GOSDOSLABK)
If X <> Trim(Mid$(oldParam.GOSEVETXT, 1, 10)) Then
    Call MsgBox("Le code du motif a été modifié," & vbCrLf & " la suppression n'est pas possible", vbCritical, "BIA_GOS : paramétrage")
Else
     oldYGOSEVE0 = oldParam
     cmdYGOSDOS0_Update "", "Delete", "", "", ""
    
     lstParam_GOSDOSLABK_Load oldYGOSEVE0.GOSEVENAT, ""
End If


Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_GOSDOSLABK_Quit_Click()
fraParam_GOSDOSLABK.Visible = False
End Sub

Private Sub cmdParam_GOSDOSLABK_Update_Click()
Dim V, X As String, XX As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

blnOk = True
X = Trim(txtParam_GOSDOSLABK)
If X <> Trim(Mid$(oldParam.GOSEVETXT, 1, 10)) Then
    blnOk = False
    Call MsgBox("Le code du motif a été modifié," & vbCrLf & " la mise à jour n'est pas possible", vbCritical, "BIA_GOS : paramétrage")
End If

V = cmdParam_GOSDOSLABK_Control

If Not IsNull(V) Then
    blnOk = False
    Call MsgBox(V, vbCritical, "BIA_GOS : paramétrage")
End If

    
 If blnOk Then
     oldYGOSEVE0 = oldParam
     newYGOSEVE0 = oldParam
     XX = Space$(12)
     If oldParam.GOSEVENAT = "DOS " Then Mid$(XX, 2, 3) = Mid$(cboParam_StaC, 1, 3)
     newYGOSEVE0.GOSEVETXT = Mid$(oldParam.GOSEVETXT, 1, 10) & " " & Format(txtParam_GOSDOSLABK_J, "000") & " " & Mid$(cboParam_GOSDOSLABK_GSrv, 1, 3) & XX & Trim(txtParam_GOSDOSLABK_Lib)
     cmdYGOSDOS0_Update "", "Update", "", "", ""
    
     lstParam_GOSDOSLABK_Load newYGOSEVE0.GOSEVENAT, ""
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Function cmdParam_GOSDOSLABK_Control()
Dim X As String, XX As String

cmdParam_GOSDOSLABK_Control = Null
X = Trim(txtParam_GOSDOSLABK_Lib)
If txtParam_GOSDOSLABK_Lib.Visible And X = "" Then
    cmdParam_GOSDOSLABK_Control = "Préciser le libellé"
    
Else
    Select Case oldParam.GOSEVENAT
        Case "SwFR", "SwGB"
                XX = fraEVE_Control_Swift_Txt(X)
            
                If XX <> "" Then cmdParam_GOSDOSLABK_Control = " - caractères interdits : " & XX
        Case "StaC"
            X = Trim(txtParam_GOSDOSLABK)

            If Len(X) > 3 Then
                cmdParam_GOSDOSLABK_Control = "le code ne doit pas dépasser 3 caractères numériques"
            Else
                If Not IsNumeric(X) Then
                    cmdParam_GOSDOSLABK_Control = "le code doit être numérique"
                End If
            End If
        Case "StaP"
            X = UCase$(Trim(txtParam_GOSDOSLABK))

            If Len(X) <> 2 Then
                cmdParam_GOSDOSLABK_Control = "Code PAYS = 2 caractères"
            Else
                 XX = "select * from  " & paramIBM_Library_SAB & ".ZBASTAB0 " _
                    & " where BASTABETA = 1 and BASTABNUM = 11 and BASTABARG = 'CLI" & X & "'"
                Set rsSab = cnsab.Execute(XX)
                If Not rsSab.EOF Then
                    txtParam_GOSDOSLABK = X
                    txtParam_GOSDOSLABK_Lib = Mid$(rsSab("BASTABLO1"), 4, 9) & Mid$(rsSab("BASTABLO2"), 1, 3)
                Else
                    cmdParam_GOSDOSLABK_Control = "PAYS inconnu"
                End If
            End If
        Case "DOS "
            X = Trim(Mid$(cboParam_StaC, 1, 3))
            If cboParam_StaC.ListCount > 0 And Len(X) = 0 Then
                cmdParam_GOSDOSLABK_Control = "Préciser le code statistique"
            Else
                If Len(X) = 1 Then
                   cmdParam_GOSDOSLABK_Control = "Ce code statistique n'est pas autorisé"
                End If
            End If
           
    End Select
End If
End Function

Private Sub cmdPrint_Click()
Dim X As String, I As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, ">cmdPrint : Initialisation ")

Select Case SSTab1.Tab
    Case 0:

        Select Case cmdSelect_SQL_K
            Case "3"
                If fraDetail.Visible Then
                    fraMail_Print
                Else
                    If fgSelect.Rows > 1 Then
                        Call cmdSendMail_fgSelect(currentSSIWINMAIL)

                    End If
                End If
            Case "4": Call cmdSendMail_fgSelect(currentSSIWINMAIL) 'cmdSendMail_Echéancier
            Case "4 Journal": Call cmdSendMail_fgSelect(currentSSIWINMAIL)
            Case Else: 'cmdPrint_Excel
                Me.PopupMenu mnuPrint2, vbPopupMenuLeftButton
        End Select
    Case 1
        Select Case tabParam.Tab
            Case 1: cmdPrint_Excel
            Case 2: cmdPrint_Excel
        End Select
        
        'Me.PopupMenu mnuPrint, vbPopupMenuLeftButton
    End Select
Call lstErr_AddItem(lstErr, cmdPrint, "<cmdPrint : terminé ")

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdParam_SAA_Add_Click()
Dim xSql As String
Me.Enabled = False: Me.MousePointer = vbHourglass

Old_YBIATAB0.BIATABK2 = Trim(txtParam_SAA_K2)
If Trim(Old_YBIATAB0.BIATABK2) = "" Then
    Call MsgBox("Préciser le code ", vbCritical, "BIA_GOS : paramétrage SAA")
Else
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = '" & Old_YBIATAB0.BIATABID & "' and BIATABK1 = '" & Old_YBIATAB0.BIATABK1 & "'" _
     & " and BIATABK2 = '" & Old_YBIATAB0.BIATABK2 & "'"
    Set rsSab = cnsab.Execute(xSql)
    If Not rsSab.EOF Then
        Call MsgBox("Ce code existe déjà", vbCritical, "BIA_GOS : paramétrage SAA")
    Else
    
        'New_YBIATAB0 = Old_YBIATAB0
        'New_YBIATAB0.BIATABTXT = Format(Val(txtParam_SAA_MTD), "000000000000")
        cmdParam_SAA_Control
    
        If IsNull(Parametrage_New) Then lstParam_SAA_Load
    End If
End If
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
           cmdPrint_Excel_rMesg
    Case 1
        Select Case tabParam.Tab
            Case 1: cmdPrint_Excel_YBIATAB0_Mail
            Case 2: cmdPrint_Excel_YBIATAB0_SAA
        End Select
        

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
Public Sub cmdSelect_SQL_Stat()
On Error GoTo Error_Handler
Dim xSql As String, K As Long, K2 As Long
Dim X As String, wFilex As String
Dim blnCALCS As Boolean
Dim xAlerte As String

On Error GoTo Error_Handler
'===================================================================================
Call DTPicker_Control(txtSelect_Stat_AMJMin, wAmjMin)
Call DTPicker_Control(txtSelect_Stat_AMJMax, wAmjMax)

'______________________________________________'

X = "C:\Temp\"

mXls1_File = mXls1_File + 1

wFile = X & Trim("GOS_Stat " & DSYS_Time & mXls1_File & ".xlsx")
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
Call lstErr_AddItem(lstErr, cmdContext, "BIA_GOS : initialisation "): DoEvents

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "GOS"
    .Subject = "Statistiques"
End With

'__________________________________________________________________________________

xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where GOSEVEIDD = -1 and GOSEVEGSRV = '" & currentSSIWINUNIT & "'" _
     & " and GOSEVENAT = 'StaC' "
Set rsSab = cnsab.Execute(xSql)

arrStaC_Nb = rsSab(0)
ReDim arrStaC(arrStaC_Nb + 1)
K = 0

xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where GOSEVEIDD = -1 and GOSEVEGSRV = '" & currentSSIWINUNIT & "'" _
     & " and GOSEVENAT = 'StaC' order by SUBSTRING(GOSEVETXT , 1 , 10)"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    K = K + 1
    X = rsSab("GOSEVETXT")
    arrStaC(K).Code = Trim(Mid$(X, 1, 10))
    arrStaC(K).Lib = Trim(Mid$(X, 31, 32))
    rsSab.MoveNext
Loop
If arrStaP_Nb >= (arrStaC_Nb + 2) Then
    arrStaP(arrStaC_Nb + 2).Lib = "non affectés"
End If
'__________________________________________________________________________________

xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where GOSEVEIDD = -1 and GOSEVEGSRV = '" & currentSSIWINUNIT & "'" _
     & " and GOSEVENAT = 'StaP' "
Set rsSab = cnsab.Execute(xSql)

arrStaP_Nb = rsSab(0)
ReDim arrStaP(arrStaP_Nb + 2)
K = 0

xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where GOSEVEIDD = -1 and GOSEVEGSRV = '" & currentSSIWINUNIT & "'" _
     & " and GOSEVENAT = 'StaP' order by SUBSTRING(GOSEVETXT , 1 , 10)"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    K = K + 1
    X = rsSab("GOSEVETXT")
    arrStaP(K).Code = Trim(Mid$(X, 1, 10))
    arrStaP(K).Lib = Trim(Mid$(X, 31, 20))
    rsSab.MoveNext
Loop
'arrStaP(arrStaP_Nb + 2).Lib = "autres pays"

mXls1_Cols = 7 + arrStaP_Nb
mXls1_Cols_WMTK = mXls1_Cols + 1
mWMTK_Col = Mid$("ABCDEFGHIJKLMNOPQRSTUVW", mXls1_Cols_WMTK, 1)


If arrStaP_Nb = 0 Then
    mXls1_Col_1 = 0: mXls1_Col_2 = 0
Else
    mXls1_Col_1 = 5: mXls1_Col_2 = 6 + arrStaP_Nb
End If
'__________________________________________________________________________________
'__________________________________________________________________________________

xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where GOSEVEIDD = -1 and GOSEVEGSRV = '" & currentSSIWINUNIT & "'" _
     & " and GOSEVENAT = 'DOS ' "
Set rsSab = cnsab.Execute(xSql)

arrDOS_Nb = rsSab(0)
ReDim arrDOS(arrDOS_Nb + 1)
K = 0

xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where GOSEVEIDD = -1 and GOSEVEGSRV = '" & currentSSIWINUNIT & "'" _
     & " and GOSEVENAT = 'DOS ' order by SUBSTRING(GOSEVETXT , 1 , 10)"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    K = K + 1
    X = rsSab("GOSEVETXT")
    arrDOS(K).Code = Trim(Mid$(X, 1, 10))
    arrDOS(K).Lib = Trim(Mid$(X, 20, 3))
    For K2 = 1 To arrStaC_Nb
        If arrDOS(K).Lib = arrStaC(K2).Code Then
            arrDOS(K).Row1 = K2
            Exit For
        End If
    Next K2
    If arrDOS(K).Row1 = 0 Then xAlerte = xAlerte & vbCrLf & "- " & arrDOS(K).Code
        
    rsSab.MoveNext
Loop

If xAlerte <> "" Then
    Call MsgBox("code motif orphelin : " & xAlerte, vbInformation, "BIA_GOS : statistiques")

End If
'__________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "BIA_GOS : Détail "): DoEvents
If appExcel.Worksheets.Count < 2 Then
    appExcel.Worksheets.Add
End If

Call cmdSelect_SQL_Stat_Init(1, "Récapitulatif " & currentSSIWINUNIT)
Call cmdSelect_SQL_Stat_Init(2, "Détail")

mXls1_Row = 1
cmdSelect_SQL_Stat_Detail

Call lstErr_AddItem(lstErr, cmdContext, "BIA_GOS : Récapitulatif "): DoEvents

Set wsExcel = wbExcel.Sheets(1)
mXls1_Row = 1

Call cmdSelect_SQL_Stat_Recapitulatif("")

mXls1_Row = mXls1_Row + 5
For K = 1 To mXls1_Cols
    wsExcel.Cells(mXls1_Row, K) = wsExcel.Cells(1, K)
    wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_Y2
Next K
wsExcel.Cells(mXls1_Row, 2) = "Motif MT 700"
wsExcel.Cells(mXls1_Row, 2).Font.Color = vbBlue
Call cmdSelect_SQL_Stat_Recapitulatif("=MT 700")

mXls1_Row = mXls1_Row + 5
For K = 1 To mXls1_Cols
    wsExcel.Cells(mXls1_Row, K) = wsExcel.Cells(1, K)
    wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_Y2
Next K
wsExcel.Cells(mXls1_Row, 2) = "Motif autres MT"
wsExcel.Cells(mXls1_Row, 2).Font.Color = vbBlue
Call cmdSelect_SQL_Stat_Recapitulatif("<>MT 700")

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
Call MsgBox("Exportation terminée.")

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

Public Sub cmdSelect_SQL_Stat_BIC()
On Error GoTo Error_Handler
Dim X As String, K As Long
Dim wFile_Orig As String, wFilex As String, xSql As String

On Error GoTo Error_Handler
Call DTPicker_Control(txtSelect_Stat_AMJMin, wAmjMin)
Call DTPicker_Control(txtSelect_Stat_AMJMax, wAmjMax)
mAMJ_SQL = " where GOSDOSSTAD <> 'A' and GOSDOSIAMJ >= " & wAmjMin & " and GOSDOSIAMJ <= " & wAmjMax
'______________________________________________'
If blnAuto Then
     wFile_Orig = Trim("C:\Temp\BIA_GOS " & DSys & "_" & time_Hms)
Else

    wFile_Orig = Trim("C:\Temp\BIA_GOS " & DSys & "_" & time_Hms)

    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile_Orig _
        & vbCrLf & "     =========================", "GSOP reporting : nom du fichier d'exportation", wFile_Orig)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile_Orig <> wFilex Then
        wFile_Orig = wFilex
    End If
    
    
End If
If Trim(Dir(wFile_Orig & ".xlsx")) <> "" Then
    Kill wFile_Orig & ".xlsx"
End If
Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "BIA_GOS" & X
    .Subject = "BIA_GOS : Stat BIC"
End With
'_________________________________________

Call cmdSelect_SQL_Stat_BIC_Detail
Call cmdSelect_SQL_Stat_BIC_Recap
'____________________________________________________________________________________

Set rsSab = Nothing


wbExcel.SaveAs wFile_Orig

wbExcel.Close

appExcel.Quit

'_________________________________________

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents
appExcel.Quit

End Sub
Public Sub cmdSelect_SQL_Stat_BIC_Detail()
Dim xSql As String, X As String, K As Long

'____________________________________________________________________________________

Call cmdSelect_SQL_Stat_BIC_Init_2("")
'______________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSDOS0" _
    & mAMJ_SQL _
    & " order by GOSDOSWBIC, GOSDOSCLI, GOSDOSISRV, GOSDOSLABK, GOSDOSIAMJ"
    
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    If xYGOSDOS0.GOSDOSWBIC <> rsSab("GOSDOSWBIC") Then Call lstErr_ChangeLastItem(lstErr, cmdContext, xYGOSDOS0.GOSDOSWBIC): DoEvents
    Call rsYGOSDOS0_GetBuffer(rsSab, xYGOSDOS0)
    Call lstErr_ChangeLastItem(lstErr, cmdContext, xYGOSDOS0.GOSDOSWBIC): DoEvents

    mXls2_Row = mXls2_Row + 1
    wsExcel.Cells(mXls2_Row, 1) = xYGOSDOS0.GOSDOSIDD
    wsExcel.Cells(mXls2_Row, 2) = xYGOSDOS0.GOSDOSWBIC
    wsExcel.Cells(mXls2_Row, 3) = xYGOSDOS0.GOSDOSCLI
    Select Case xYGOSDOS0.GOSDOSISRV
        Case "S01": wsExcel.Cells(mXls2_Row, 4) = arrService_Lib(1)
        Case "S42": wsExcel.Cells(mXls2_Row, 4) = arrService_Lib(42)
        Case Else
            wsExcel.Cells(mXls2_Row, 4) = arrService_Lib(Mid$(xYGOSDOS0.GOSDOSISRV, 2, 2))
    End Select
    wsExcel.Cells(mXls2_Row, 5) = xYGOSDOS0.GOSDOSLABK
    wsExcel.Cells(mXls2_Row, 6) = xYGOSDOS0.GOSDOSIAMJ
    wsExcel.Cells(mXls2_Row, 7) = xYGOSDOS0.GOSDOSWES
    wsExcel.Cells(mXls2_Row, 8) = xYGOSDOS0.GOSDOSWMTK
    wsExcel.Cells(mXls2_Row, 9) = xYGOSDOS0.GOSDOSPAYS
    wsExcel.Cells(mXls2_Row, 10) = xYGOSDOS0.GOSDOSRCOM
    wsExcel.Cells(mXls2_Row, 11) = xYGOSDOS0.GOSDOSSTAG
    wsExcel.Cells(mXls2_Row, 12) = xYGOSDOS0.GOSDOSSTAD
    
    Select Case xYGOSDOS0.GOSDOSSTAD
        Case "A": wsExcel.Cells(mXls2_Row, 12) = "Ann"
            For K = 1 To mXls2_Col
                wsExcel.Cells(mXls2_Row, K).Interior.Color = RGB(220, 220, 220)
                wsExcel.Cells(mXls2_Row, K).Font.Color = RGB(64, 64, 64)
            Next K
        Case "C", "x": wsExcel.Cells(mXls2_Row, 12) = "Clos"
            wsExcel.Cells(mXls2_Row, 12).Interior.Color = RGB(220, 220, 220)
    End Select
    Select Case xYGOSDOS0.GOSDOSSTAG
        Case "V": wsExcel.Cells(mXls2_Row, 11).Interior.Color = mColor_G1: wsExcel.Cells(mXls2_Row, 11) = "Val"
        Case "R": wsExcel.Cells(mXls2_Row, 11).Interior.Color = mColor_W1: wsExcel.Cells(mXls2_Row, 11) = "Rejet"
    End Select

    rsSab.MoveNext
Loop
End Sub

Public Sub cmdSelect_SQL_Stat_BIC_Init_2(lX As String)
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim X As String, K As Long
Dim xSql As String, wFile As String

On Error GoTo Error_Handler
'__________________________________________________________________________________________________

'Set wsExcel = wbExcel.ActiveSheet
Set wsExcel = wbExcel.Sheets(2)

wsExcel.Name = "Detail"


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
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 100
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14BIA_GOS : statistiques BIC  " & lX _
                                & "&B&U&10     ( édité le " & dateImp10(DSys) & " " & Time & ")"
wsExcel.PageSetup.PrintTitleRows = "$A1:$G1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


wsExcel.Columns(1).ColumnWidth = 7: wsExcel.Cells(1, 1) = "Dossier"
wsExcel.Columns(2).ColumnWidth = 14: wsExcel.Cells(1, 2) = "BIC"
wsExcel.Columns(3).ColumnWidth = 8: wsExcel.Cells(1, 3) = "Racine"
wsExcel.Columns(4).ColumnWidth = 7: wsExcel.Cells(1, 4) = "Service"
wsExcel.Columns(5).ColumnWidth = 15: wsExcel.Cells(1, 54) = "Motif"
wsExcel.Columns(6).ColumnWidth = 9: wsExcel.Cells(1, 6) = "Date Création"
wsExcel.Columns(7).ColumnWidth = 4: wsExcel.Cells(1, 7) = "Sens"
wsExcel.Columns(8).ColumnWidth = 5: wsExcel.Cells(1, 8) = "type MT"
wsExcel.Columns(9).ColumnWidth = 5: wsExcel.Cells(1, 9) = "Pays"
wsExcel.Columns(10).ColumnWidth = 5: wsExcel.Cells(1, 10) = "Rcom"
wsExcel.Columns(11).ColumnWidth = 7: wsExcel.Cells(1, 10) = "Val/Rej"
wsExcel.Columns(12).ColumnWidth = 7: wsExcel.Cells(1, 10) = "Clos"

wsExcel.Rows(1).RowHeight = 34
wsExcel.Rows(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Rows(1).VerticalAlignment = Excel.xlVAlignCenter

mXls2_Col = 12: mXls2_Row = 1

For K = 1 To mXls2_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next


'=======================================================================================
'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée "): DoEvents

End Sub

Public Sub cmdSelect_SQL_Stat_BIC_Init_1(lX As String)
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long, xCol_Address As String, blnOk As Boolean
Dim X As String, K As Long
Dim xSql As String, wFile As String
Dim arrGOSDOSISRV(500) As String, arrGOSDOSLABK(500) As String
Dim wGOSDOSISRV As String, wGOSDOSLABK As String, wGOSDOSWBIC As String, wGOSDOSCLI As String
On Error GoTo Error_Handler
'__________________________________________________________________________________________________

'Set wsExcel = wbExcel.ActiveSheet
Set wsExcel = wbExcel.Sheets(1)

wsExcel.Name = "Recap"


With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .HorizontalAlignment = Excel.xlHAlignRight
    .WrapText = True
    .Font.Size = 9
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 6
wsExcel.PageSetup.Zoom = 70
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14BIA_GOS : statistiques BIC  " & lX _
                                & "&B&U&10     ( édité le " & dateImp10(DSys) & " " & Time & ")"
wsExcel.PageSetup.PrintTitleRows = "$A1:$G1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


mXls1_Col = 2
wsExcel.Columns(1).ColumnWidth = 14: wsExcel.Cells(1, 1) = "BIC"
wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 10: wsExcel.Cells(1, 2) = "Racine"
wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Rows(1).RowHeight = 72
wsExcel.Rows(1).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Rows(1).VerticalAlignment = Excel.xlVAlignCenter

xSql = "select gosdosisrv , gosdoslabk from " & paramIBM_Library_SABSPE & ".YGOSDOS0" _
    & mAMJ_SQL _
    & " group by gosdosisrv, gosdoslabk" _
    & " order by gosdosisrv, gosdoslabk"

Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    mXls1_Col = mXls1_Col + 1
    Select Case rsSab("GOSDOSISRV")
        Case "S01": X = arrService_Lib(1)
        Case "S42": X = arrService_Lib(42)
        Case Else
            X = arrService_Lib(Mid$(rsSab("GOSDOSISRV"), 2, 2))
    End Select
    wsExcel.Columns(mXls1_Col).ColumnWidth = 6
    wsExcel.Cells(1, mXls1_Col) = X & vbCr & Trim(rsSab("GOSDOSLABK"))
    
    arrGOSDOSISRV(mXls1_Col) = rsSab("GOSDOSISRV")
    arrGOSDOSLABK(mXls1_Col) = Trim(rsSab("GOSDOSLABK"))
    
    rsSab.MoveNext
Loop

xCol_Address = wsExcel.Cells(1, mXls1_Col).Address
Mid$(xCol_Address, 1, 1) = ":"
K = InStr(xCol_Address, "$")
xCol_Address = Mid$(xCol_Address, 1, K - 1)

wsExcel.Rows(1).RowHeight = 34
wsExcel.Rows(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Rows(1).VerticalAlignment = Excel.xlVAlignCenter

mXls1_Row = 1

For K = 1 To mXls1_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next
'=======================================================================================
xSql = "select gosdoswbic , gosdoscli ,gosdosisrv , gosdoslabk , count(*)  from " & paramIBM_Library_SABSPE & ".YGOSDOS0" _
    & mAMJ_SQL _
    & " group by gosdoswbic , gosdoscli , gosdosisrv, gosdoslabk" _
    & " order by gosdoswbic , gosdoscli , gosdosisrv, gosdoslabk"

Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    If wGOSDOSWBIC <> rsSab("GOSDOSWBIC") Then Call lstErr_ChangeLastItem(lstErr, cmdContext, "Recap : " & rsSab("GOSDOSWBIC")): DoEvents
    If wGOSDOSWBIC <> rsSab("GOSDOSWBIC") Or wGOSDOSCLI <> rsSab("GOSDOSCLI") Then
        mXls1_Row = mXls1_Row + 1
        wGOSDOSWBIC = rsSab("GOSDOSWBIC")
        wGOSDOSCLI = rsSab("GOSDOSCLI")
        wsExcel.Cells(mXls1_Row, 1) = " " & rsSab("GOSDOSWBIC")
        wsExcel.Cells(mXls1_Row, 2) = " " & rsSab("GOSDOSCLI")
    End If
    wGOSDOSISRV = rsSab("GOSDOSISRV")
    wGOSDOSLABK = Trim(rsSab("GOSDOSLABK"))
    blnOk = False
    For K = 3 To mXls1_Col
        If arrGOSDOSISRV(K) = wGOSDOSISRV And arrGOSDOSLABK(K) = wGOSDOSLABK Then
            wsExcel.Cells(mXls1_Row, K) = rsSab(4)
            wsExcel.Cells(mXls1_Row, K).Font.Bold = True
            wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G0
            'Debug.Print wGOSDOSWBIC & " " & wGOSDOSCLI & " " & wGOSDOSISRV & " " & wGOSDOSLABK, rsSab(4)
            blnOk = True
            Exit For
        End If
        
    Next K
    If Not blnOk Then Call MsgBox(wGOSDOSWBIC & " " & wGOSDOSCLI & " " & wGOSDOSISRV & " " & wGOSDOSLABK, vbCritical, "BIA_GOS : stat BIC")
    
    rsSab.MoveNext
Loop

'=======================================================================================
wCol = mXls1_Col + 1
wsExcel.Columns(wCol).ColumnWidth = 6: wsExcel.Cells(1, wCol) = "Total": wsExcel.Cells(1, wCol).Interior.Color = mColor_G1
wsExcel.Cells(2, wCol).FormulaLocal = "=SOMME(C2" & xCol_Address & 2 & ")": wsExcel.Cells(2, wCol).Interior.Color = mColor_G1
wsExcel.Cells(2, wCol).Select: wsExcel.Cells(2, wCol).Copy

For K = 3 To mXls1_Row
    'wsExcel.Cells(K, wCol).Select: ActiveSheet.Paste

    wsExcel.Cells(K, wCol).FormulaLocal = "=SOMME(C" & K & xCol_Address & K & ")"
    wsExcel.Cells(K, wCol).Interior.Color = mColor_G1
Next K

'=======================================================================================

wRow = mXls1_Row + 1
wsExcel.Cells(wRow, 1) = "Total": wsExcel.Cells(wRow, 1).Interior.Color = mColor_G1
wsExcel.Cells(wRow, 2).Interior.Color = mColor_G1
wsExcel.Cells(wRow, 3).FormulaLocal = "=SOMME(C2:C" & mXls1_Row & ")": wsExcel.Cells(wRow, 3).Interior.Color = mColor_G1
wsExcel.Cells(wRow, 3).Select: wsExcel.Cells(wRow, 3).Copy

For K = 4 To wCol
    wsExcel.Cells(wRow, K).Select: ActiveSheet.Paste
    'wsExcel.Cells(wRow, K).Interior.Color = mColor_G1
Next K
'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée "): DoEvents

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
wsExcel.Columns(5).ColumnWidth = 15: wsExcel.Cells(mXls1_Row, 5) = "Montant":
wsExcel.Columns(5).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
'wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignRight
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
    fgSelect.Col = 3: wsExcel.Cells(mXls1_Row, 5) = Val(fgSelect.Text): wsExcel.Cells(mXls1_Row, 5).Font.Color = wForecolor
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
Public Sub cmdSelect_SQL_Stat_Init(lSheet As Integer, lName As String)
Dim xSql As String, X As String, K As Long
On Error GoTo Error_Handler


'On Error GoTo Error_Handler
'===================================================================================

Set wsExcel = wbExcel.Sheets(lSheet): wsExcel.Name = lName
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
    .Font.Size = 10
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 70
K = Val(Mid$(currentSSIWINUNIT, 2, 2))

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14" & Trim(arrService_Lib(K)) & " : dossiers GOS créés du " & dateImp10_S(wAmjMin) & " au " & dateImp10_S(wAmjMax) _
                               & vbCr & "(édité le " & dateImp10(wAmjMin) & ")" _
                                 & vbCr

wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"


wsExcel.Columns(1).ColumnWidth = 5: wsExcel.Cells(1, 1) = "Code "
wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(2).ColumnWidth = 15: wsExcel.Cells(1, 2) = "Motif"
wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignLeft

If lSheet = 1 Then
    wsExcel.Columns(1).Font.Bold = True
    wsExcel.Columns(2).Font.Bold = True
    wsExcel.Columns(3).Font.Bold = True
    wsExcel.Columns(4).Font.Bold = True
    wsExcel.Columns(3).ColumnWidth = 12: wsExcel.Cells(1, 3) = "Bloqués"
    wsExcel.Columns(4).ColumnWidth = 12: wsExcel.Cells(1, 4) = "dont rejetés"
    wsExcel.Columns(mXls1_Cols - 1).ColumnWidth = 12: wsExcel.Cells(1, mXls1_Cols - 1) = "Embargo pays"
    wsExcel.Columns(mXls1_Cols - 1).NumberFormat = "[Blue]### ### ###"
Else
    wsExcel.Cells.Font.Name = "Courier New"
    wsExcel.Cells.Font.Size = 9
    wsExcel.Columns(3).ColumnWidth = 12: wsExcel.Cells(1, 3) = "N° dossier"
    wsExcel.Columns(4).ColumnWidth = 12: wsExcel.Cells(1, 4) = "Rejetés"
    'wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignCenter
    wsExcel.Columns(mXls1_Cols - 1).ColumnWidth = 22: wsExcel.Cells(1, mXls1_Cols - 1) = "CLI  BIC  D.O  BEN"
    wsExcel.Columns(mXls1_Cols - 1).HorizontalAlignment = Excel.xlHAlignCenter
    wsExcel.Columns(mXls1_Cols).ColumnWidth = 12: wsExcel.Cells(1, mXls1_Cols) = "Client"
    wsExcel.Columns(mXls1_Cols).HorizontalAlignment = Excel.xlHAlignCenter
    wsExcel.Columns(mXls1_Cols_WMTK).ColumnWidth = 8: wsExcel.Cells(1, mXls1_Cols_WMTK) = "MT"
End If

For K = 1 To arrStaP_Nb
    wsExcel.Columns(K + 4).ColumnWidth = 12: wsExcel.Cells(1, K + 4) = arrStaP(K).Lib
    wsExcel.Columns(K + 4).NumberFormat = "### ### ###"
Next K
wsExcel.Columns(mXls1_Cols - 2).ColumnWidth = 12: wsExcel.Cells(1, mXls1_Cols - 2) = "autres pays"
wsExcel.Columns(mXls1_Cols - 2).NumberFormat = "[Blue]### ### ###"

wsExcel.Columns(3).NumberFormat = "[Blue]### ### ###"
wsExcel.Columns(4).NumberFormat = "[red]### ### ###"

For K = 1 To mXls1_Cols_WMTK 'mXls1_Cols
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = vbWhite
Next




'======================================================================================================

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub


Public Sub cmdPrint_Excel_YBIATAB0_Mail()
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
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14SAA_Alerte : paramétrage Mail" _
                                & "  (édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$E1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


wsExcel.Columns(1).ColumnWidth = 10: wsExcel.Cells(1, 1) = "Service "
wsExcel.Columns(2).ColumnWidth = 15: wsExcel.Cells(1, 2) = "Libellé"
wsExcel.Columns(3).ColumnWidth = 45: wsExcel.Cells(1, 3) = "Destinataires GOS_DOS"
wsExcel.Columns(4).ColumnWidth = 45: wsExcel.Cells(1, 4) = "Destinataires SAA_Alerte"

For K = 1 To 4
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = vbWhite
Next

For K = 1 To 99

    mXls1_Row = mXls1_Row + 1
    wsExcel.Cells(mXls1_Row, 1) = arrService_Code(K)
    wsExcel.Cells(mXls1_Row, 2) = arrService_Lib(K)
    wsExcel.Cells(mXls1_Row, 3) = arrService_Mail(K, 1)
    wsExcel.Cells(mXls1_Row, 4) = arrService_Mail(K, 2)
    
Next K


'======================================================================================================

Exit_sub:
'__________________________________________________________________________________


'_____________________________
Exit Sub

Error_Handler:

End Sub


Private Sub cmdPrint_YGOSDOS0(blnDetail As Boolean)
Dim X As String, xSql As String, I As Integer, K As Integer
Dim wAmj As String, xWhere As String
Dim soldeD As typeYGOSDOS0, soldeF As typeYGOSDOS0, Total As typeYGOSDOS0
Dim blnXprt_Line As Boolean



Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> BIA_GOS_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Clear

Select Case cmdSelect_SQL_K
    Case "1": cmdSelect_SQL_1
    Case "1Live_Entran", "1Live_Sortan": cmdSelect_SQL_1L
    Case "1Lmail": cmdSelect_SQL_1Lmail
    Case "1b", "1?", "1?*", "1trf": cmdSelect_SQL_1b
    Case "2": cmdSelect_SQL_1b
    Case "2d": fraDetail_LAB_Init_Vierge
    Case "3": cmdSelect_SQL_3
    Case "3x": cmdSelect_SQL_3x
    Case "4": cmdSelect_SQL_4
    Case "4 Journal": cmdSelect_SQL_4Journal
    Case "4 Swi>": cmdSelect_SQL_4Swi
    Case "5": cmdSelect_SQL_5
    Case "5h": cmdSelect_SQL_5h
    Case "6", "6#": cmdSelect_SQL_6
    Case "6E": cmdSelect_SQL_6 '''6E
    Case "2-RAM": cmdSelect_SQL_K = "6E": cmdSelect_SQL_6 '''6E
    Case "7", "7#": cmdSelect_SQL_7
    Case "7E", "7E*": cmdSelect_SQL_7E
    Case "7J": cmdSelect_SQL_7J
    Case "9": cmdSelect_SQL_9
    Case "9+": cmdSelect_SQL_9M
    Case "Stat": cmdSelect_SQL_Stat
    Case "Stat BIC": cmdSelect_SQL_Stat_BIC
    Case "JPL":
            'cmdSelect_JPL
            Importation_SAA_198_Alerte
            'Importation_SAA_SWISABWSTA
            'YSWIECH0_Auto
            'YSWIRAM0_Importation
            'cmdSelect_JPL
             'Importation_SAA
             'Importation_Jrnl
            ' Importation_SIDE_Reporting_Control
            'Importation_SAA_Modification_MT
            'Importation_SAA_Origine_MT
           ' Importation_SAB_ZSWIHIA0
            'Importation_SAB_ZCDODOS0
            'Importation_SAA_Alerte'
            'Importation_SAB_SWISABSWID_MT700
            'Importation_Jrnl
            ' Form_Migration_20120523
           ' Importation_SAB_YGOSDOS0_Synchro2
            'Importation_SAB_YGOSDOS0_Synchro1
            'Form_Migration_20120511
            'Auto_BIA_GOS
           ' Importation_SAB_ZCDODOS0
           ' Importation_SAB_SWISABWN20
            'param_Init_MT_Fields
             '   cmdSelect_JPL
             'Importation_SAB_YSWIMON0_Synchro1
             'Importation_SAB_YSWIMON0_Synchro2
             'Importation_SAB_YSWIMON0_Synchro3
             
             'Importation_SAB_YSWIMON0
            'Importation_SAB_YSWISAB1
            'Importation_SAA_SWISABWSTA
            'Importation_Jrnl
            'cmdSelect_SQL_rCorr
            'cmdSelect_SQL_JPL
                'Importation_SAA_Alerte
                'mSWISABSWID_Xd = 0
                'Importation_SAB_YSWISAB1
                
                'Importation_SAB_YSWIMON0
                'Importation_SAA_SWISABWSTA
                'Importation_Reprise_SWISABWSTA
                'cmdSelect_SQL_JPL
                'Importation_Reprise_20110720
                'Importation_SAB_SWISABKPDE
                'Importation_Reprise_SWISABN20
    
    ' REPRISE Case "$Xi":
    '   Dim K As Integer, xAMJ As String, blnOk As Boolean
    '   blnOk = True
    '   xAMJ = "20040301"
    '   Do
    '        Call Importation_SAA_Reprise(xAMJ)
    '        xAMJ = dateElp("Jour", 1, xAMJ)
    '        If xAMJ > DSys Then blnOk = False
    '   Loop Until Not blnOk
    
    'Case "Xr": Importation_SAA_SWISABZSWI: Importation_SAA_SWISABZSWI_Reprise
    'Case "SH": YJPLSLD1_Exportation
    'Case "XX":  mSWISABSWID_Xd = 0
                '$$$$$$$'''''''''''''''''Importation_SAB_ZSWICLA0
                '$$$$$$''''''''''''''''''Importation_SAB_SWISABWN20 'avec SWISABWES = 'S'
                'Importation_SAB_ZCDODOS0
                'Importation_SAB_ZCHGOPE0
                'Importation_SAB_SWISABWN20
                
    'Case "Xd": Importation_SAB_Dossier
    'Case "#L20": Call Importation_Reprise_SWISABWL20_SWISABOPEN(700000, 799999)
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_GOS_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
cmdSelect_Ok.BackColor = fgSelect.BackColorFixed
End Sub


Private Sub cmdSwift_Print_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

Call prtYSWISAB0_Monitor(oldYSWISAB0.SWISABSWID, Mesg_aid, mesg_s_umidl, mesg_s_umidh)

Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSWISABKSRV__Update_Click()

Dim V
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
If IsNull(fraSWISABKSRV_Control) Then

    mYSWISAB0_Fct = "Update"
    
    
    rsYGOSEVE0_Init newYGOSEVE0
    newYGOSEVE0.GOSEVESWID = newYSWISAB0.SWISABSWID
    newYGOSEVE0.GOSEVENAT = "*D"
    newYGOSEVE0.GOSEVEGSRV = newYSWISAB0.SWISABKSRV
    newYGOSEVE0.GOSEVETXT = "réaffectation SAB : " & oldYSWISAB0.SWISABSER & "-" & oldYSWISAB0.SWISABSSE & "-" & oldYSWISAB0.SWISABOPEC & "-" & oldYSWISAB0.SWISABOPEN _
                        & "=> " & newYSWISAB0.SWISABSER & "-" & newYSWISAB0.SWISABSSE & "-" & newYSWISAB0.SWISABOPEC & "-" & newYSWISAB0.SWISABOPEN
    mYGOSEVE0_Fct = "New"
    V = cmdYGOSDOS0_Update("", mYGOSEVE0_Fct, "", mYSWISAB0_Fct, "")
    
    
    If IsNull(V) Then
        fraSWISABKSRV.Visible = False
        cmdSelect_Clear
        cmdSelect_SQL_1b
        Me.Enabled = True: Me.MousePointer = 0
        Exit Sub
    End If
End If
Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
    Call MsgBox(V, vbCritical, "Mise à jour YSWISAB0")
    Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdSWISABKSRV_Quit_Click()
fraSWISABKSRV.Visible = False
End Sub

Private Sub dirListBox_Change()
filDoc.path = dirListBox.path
filDoc.Pattern = "*.*"
If mfilDoc_Path <> filDoc.path Then cmdGOSDOSPJ.Visible = True
End Sub

Private Sub dirListBox_Click()
'filDoc.Path = dirListBox.Path
'filDoc.Pattern = "*.*"

End Sub


Private Sub DriveListBox_Change()
On Error Resume Next
dirListBox.path = DriveListBox.Drive ' .PATH

End Sub

Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next


If y <= fgDetail.RowHeightMin Then
Else
    If fgDetail.Rows > 1 Then
        If fraDetail_LAB.Enabled Then
            fgDetail.Col = 1
            If fgDetail.CellForeColor = vbRed Then
                fgDetail.CellForeColor = fgDetail.ForeColor
                fgDetail.CellFontBold = False
            Else
                fgDetail.CellForeColor = vbRed
                fgDetail.CellFontBold = True
            End If
        End If
        Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
        'fgDetail.Col = fgDetail_arrIndex:  arrYGOSDOS0_Index = CLng(fgDetail.Text)
        'oldYGOSDOS0 = arrYGOSDOS0(arrYGOSDOS0_Index)
        'xYGOSDOS0 = oldYGOSDOS0

   End If
End If
Wait_SS 0
fgDetail.LeftCol = 0

End Sub


Private Sub fgEVE_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, xUUMID As String
On Error Resume Next


If y <= fgEVE.RowHeightMin Then
Else
    If fgEVE.Rows > 1 Then
        'Call fgEVE_Color(fgEVE_RowClick, MouseMoveUsr.BackColor, fgEVE_ColorClick)
        fgEVE.Col = fgEVE_arrIndex:  arrYGOSEVE0_Index = CLng(fgEVE.Text)
        oldYGOSEVE0 = arrYGOSEVE0(arrYGOSEVE0_Index)
        fraEVE_S.Enabled = False
        fraEVE_Display
 
   End If
End If
Wait_SS 0
fgEVE.LeftCol = 0

End Sub

Private Sub fgList_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String
Dim K As Integer
On Error Resume Next


If y <= fgList.RowHeightMin Then
    fglist_Sort1 = fgList.Col: fglist_Sort2 = fgList.Col + 1: fgList_Sort
    'Select Case fgList.Col
    'End Select
Else
    If fgList.Rows > 1 Then
        If cmdSelect_SQL_K = "5" Then
            fgList.Col = fglist_arrIndex
            Call fgSwift_Display(CLng(fgList.Text))
        Else
    
            fgList.Col = fglist_arrIndex:  arrYGOSDOS0_Index = CLng(fgList.Text)
            Call fgList_Color(fglist_RowClick, MouseMoveUsr.BackColor, fglist_ColorClick)
            oldYGOSDOS0 = arrYGOSDOS0(arrYGOSDOS0_Index)
            txtList_Add = oldYGOSDOS0.GOSDOSIDD
            
            fraDetail_LAB_Display_Dossier
            
        End If
        
        Wait_SS 0
       fgList.LeftCol = 0
    End If
End If
Wait_SS 0
fgList.LeftCol = 0

End Sub


Private Sub fgModèle_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, xUUMID As String, Nb As Integer
On Error Resume Next


If y <= fgModèle.RowHeightMin Then
Else
    If fgModèle.Rows > 1 Then
        fgModèle.Visible = False
        fgModèle.Col = 2
        wX = fgModèle.Text
        wX = Replace(wX, "#MTK", "MT" & oldYGOSDOS0.GOSDOSWMTK)
        wX = Replace(wX, "#AMJ", dateImp10_S(oldYSWISAB0.SWISABWAMJ))
        wX = Replace(wX, "#DEV", oldYGOSDOS0.GOSDOSWDEV)
        wX = Replace(wX, "#MTD", Format(oldYGOSDOS0.GOSDOSWMTD, "### ### ### ##0.00"))
        wX = Replace(wX, "#L20", oldYSWISAB0.SWISABWL20)
        wX = Replace(wX, "#N20", oldYSWISAB0.SWISABWN20)
        wX = Replace(wX, "#GOS", cboEVE_Swift_20)
        
        If InStr(wX, "#50") > 0 And fgDetail_50 <> "" Then
            'fgDetail_50 = Replace(fgDetail_50, Asc13 & Asc10, " - ")
            wX = Replace(wX, "#50", fgDetail_50)
        End If
        
        If InStr(wX, "#57") > 0 And fgDetail_57 <> "" Then
            wX = Replace(wX, "#57", fgDetail_57)
        End If
        If InStr(wX, "#59") > 0 And fgDetail_59 <> "" Then
            wX = Replace(wX, "#59", fgDetail_59)
        End If
        
        If InStr(wX, "#32B") > 0 And fgDetail_32B <> "" Then
            wX = Replace(wX, "#32B", fgDetail_32B)
        End If
        
        If InStr(wX, "#33B") > 0 And fgDetail_33B <> "" Then
            wX = Replace(wX, "#33B", fgDetail_33B)
        End If
        
        If InStr(wX, "#36") > 0 And fgDetail_36 <> "" Then
            wX = Replace(wX, "#36", fgDetail_36)
        End If
        
        If InStr(wX, "#30V") > 0 And fgDetail_30V <> "" Then
            wX = Replace(wX, "#30V", fgDetail_30V)
        End If
        
        If InStr(wX, "#37G") > 0 And fgDetail_37G <> "" Then
            wX = Replace(wX, "#37G", fgDetail_37G)
        End If
        
        If InStr(wX, "#34E") > 0 And fgDetail_34E <> "" Then
            wX = Replace(wX, "#34E", fgDetail_34E)
        End If
        
        If InStr(wX, "#30T") > 0 And fgDetail_30T <> "" Then
            wX = Replace(wX, "#30T", fgDetail_30T)
        End If
        
        If InStr(wX, "#30P") > 0 And fgDetail_30P <> "" Then
            wX = Replace(wX, "#30P", fgDetail_30P)
        End If
        
        If InStr(wX, "#82A") > 0 And fgDetail_82A <> "" Then
            wX = Replace(wX, "#82A", fgDetail_82A)
        End If
        
        If InStr(wX, "#87A") > 0 And fgDetail_87A <> "" Then
            wX = Replace(wX, "#87A", fgDetail_87A)
        End If
        
        If InStr(wX, "#22C") > 0 And fgDetail_22C <> "" Then
            wX = Replace(wX, "#22C", fgDetail_22C)
        End If
        
        If InStr(wX, "#70") > 0 And fgDetail_70 <> "" Then
            wX = Replace(wX, "#70", fgDetail_70)
        End If
        
        If InStr(wX, "#72") > 0 And fgDetail_72 <> "" Then
            wX = Replace(wX, "#72", fgDetail_72)
        End If
        
        txtGOSEVETXT = wX
'_____________________________________________
        fgModèle.Col = 1

        Nb = Val(Mid$(fgModèle.Text, 1, 3))
        If Nb > 0 Then
            wX = dateElp("Ouvré", Nb, DSys)
            If wX > oldYGOSDOS0.GOSDOSECHD Then Call DTPicker_Set(txtGOSEVEECHD, wX)
        End If
        If Mid$(fgModèle.Text, 5, 3) <> oldYGOSDOS0.GOSDOSGSRV Then
            Call cbo_Scan(Mid$(fgModèle.Text, 5, 3), cboGOSEVEGSRV)
            If blnHab_YGOSEVE0_New Then chkGOSEVEGSRV.value = "1"
        
        End If
'_____________________________________________
        cmdEVE_Ok.Visible = True
    End If
End If
fgModèle.LeftCol = 0

End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, xUUMID As String
Dim verifNewId As Long
On Error Resume Next


If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0:
            If cmdSelect_SQL_K = "3" Or cmdSelect_SQL_K = "4" Then
                fgSelect_SortX 0
            Else
                fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
            End If
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
        Case 11: fgSelect_Sort1 = 11: fgSelect_Sort2 = 11: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fraDetail.Visible = False
        fraDetail_C.Visible = False
        fraMail_MT.Visible = False
        fgEVE.Visible = False
        fraSwift.Visible = False
        fraMail_MT.Visible = False
        fgFree.Visible = False
        mSWIECHSWIL = 0
        blnYGOSDOS0_New = False
        
        Select Case cmdSelect_SQL_K
            Case "1", "1Live_Entran", "1Live_Sortan", "1b", "1?", "1?*", "1trf", "2":
                rsYGOSDOS0_Init oldYGOSDOS0
                arrYGOSDOS0_Index = 1: ReDim arrYGOSDOS0(2)
                rsYGOSEVE0_Init oldYGOSEVE0
                fgSelect.Col = 8: Mesg_aid = Val(fgSelect.Text)
                fgSelect.Col = 9: mesg_s_umidl = Val(fgSelect.Text)
                fgSelect.Col = 10: mesg_s_umidh = Val(fgSelect.Text)
                'fgSelect.Col = 11: oldYGOSEVE0.GOSEVESWID = Val(fgSelect.Text)
                fgDetail_Display
            Case "3", "4":
                fgSelect.Col = fgSelect_arrIndex:  arrYGOSDOS0_Index = CLng(fgSelect.Text)
                oldYGOSDOS0 = arrYGOSDOS0(arrYGOSDOS0_Index)
                newYGOSDOS0 = oldYGOSDOS0
                Mesg_aid = oldYGOSDOS0.GOSDOSWID1
                mesg_s_umidl = oldYGOSDOS0.GOSDOSWIDL
                mesg_s_umidh = oldYGOSDOS0.GOSDOSWIDH
                fraDetail_C.Visible = True
                fgDetail_Display
             Case "4 Journal":
                fgSelect.Col = 6
                wX = "select * from " & paramIBM_Library_SABSPE & ".YGOSDOS0 where GOSDOSIDD = " & Val(fgSelect.Text)
                Set rsSab = cnsab.Execute(wX)
                If Not rsSab.EOF Then
                    V = rsYGOSDOS0_GetBuffer(rsSab, oldYGOSDOS0)

                    newYGOSDOS0 = oldYGOSDOS0
                    Mesg_aid = oldYGOSDOS0.GOSDOSWID1
                    mesg_s_umidl = oldYGOSDOS0.GOSDOSWIDL
                    mesg_s_umidh = oldYGOSDOS0.GOSDOSWIDH
                    fraDetail_C.Visible = True
                    fgDetail_Display
                End If
            Case "9", "9+", "5":
                rsYGOSDOS0_Init oldYGOSDOS0
                arrYGOSDOS0_Index = 1: ReDim arrYGOSDOS0(2)
                rsYGOSEVE0_Init oldYGOSEVE0
                fgSelect.Col = 1: oldYGOSDOS0.GOSDOSWBIC = Trim(fgSelect.Text)
                fgSelect.Col = 0: oldYGOSDOS0.GOSDOSWMTK = Mid$(Trim(fgSelect.Text), 1, 3)
                oldYGOSDOS0.GOSDOSWES = Mid$(fgSelect.Text, 5, 1)
                fgSelect.Col = 2: oldYGOSDOS0.GOSDOSWTRN = Trim(fgSelect.Text)
                fgSelect.Col = 4: oldYGOSDOS0.GOSDOSWDEV = Trim(fgSelect.Text)
                fgSelect.Col = 3: oldYGOSDOS0.GOSDOSWMTD = CCur(fgSelect.Text)
                'fgSelect.Col = 5: Call dateJMA_AMJ(Trim(fgSelect.Text), wX)
                'oldYGOSDOS0.GOSDOSWDVA = Val(wX)
                fgSelect.Col = 8: Mesg_aid = Val(fgSelect.Text)
                fgSelect.Col = 9: mesg_s_umidl = Val(fgSelect.Text)
                fgSelect.Col = 10: mesg_s_umidh = Val(fgSelect.Text)
                
                rsYSWILNK0_Init oldYSWILNK0
                oldYSWILNK0.SWILNKAPPC = "GOS"
                fgSelect.Col = 11: oldYSWILNK0.SWILNKSWID = Val(fgSelect.Text)
                oldYGOSEVE0.GOSEVESWID = oldYSWILNK0.SWILNKSWID
                m999_YGOSDOS0 = oldYGOSDOS0
                fgDetail_Display

            Case "5h":
                 fgSelect.Col = 5: oldYSWISAB0.SWISABSWID = Val(fgSelect.Text)
                wX = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABSWID = " & oldYSWISAB0.SWISABSWID
                Set rsSab = cnsab.Execute(wX)
                If Not rsSab.EOF Then
                    Mesg_aid = rsSab("SWISABWID1")
                    mesg_s_umidl = rsSab("SWISABWIDL")
                    mesg_s_umidh = rsSab("SWISABWIDH")
                    oldYGOSDOS0.GOSDOSWES = rsSab("SWISABWES")
                    oldYGOSDOS0.GOSDOSWBIC = rsSab("SWISABWBIC")
                    oldYGOSDOS0.GOSDOSWMTK = rsSab("SWISABWMTK")
                    oldYGOSDOS0.GOSDOSWTRN = rsSab("SWISABWTRN")
                    oldYGOSDOS0.GOSDOSWDEV = rsSab("SWISABWDEV")
                    oldYGOSDOS0.GOSDOSWMTD = rsSab("SWISABWMTD")
                    oldYGOSEVE0.GOSEVESWID = oldYSWISAB0.SWISABSWID
                    fgDetail_Display
                End If
            Case "6", "6E", "2-RAM":
                If cmdSelect_SQL_K = "2-RAM" Then cmdSelect_SQL_K = "6E"
                Dim blnSuite As Boolean
                rsYGOSDOS0_Init oldYGOSDOS0
                fgSelect.Col = 11
                If fgSelect.CellFontBold = True Then mYSWIRAM0_Col = 0
                If mYSWIRAM0_Col = 0 Then blnSuite = True
                If Val(fgSelect.Text) > 0 Then
                    fgSelect.Col = 8: Mesg_aid = Val(fgSelect.Text)
                    fgSelect.Col = 9: mesg_s_umidl = Val(fgSelect.Text)
                    fgSelect.Col = 10: mesg_s_umidh = Val(fgSelect.Text)
                    fgDetail_Display
                    If blnSuite And fgSelect.Row < fgSelect.Rows - 1 Then
                        fgSelect.Row = fgSelect.Row + 1
                        fgSelect.Col = 11
                        If Val(fgSelect.Text) > 0 Then
                            fgSelect.Col = 8: Mesg_aid = Val(fgSelect.Text)
                            fgSelect.Col = 9: mesg_s_umidl = Val(fgSelect.Text)
                            fgSelect.Col = 10: mesg_s_umidh = Val(fgSelect.Text)
                            fgDetail_Display
                        End If
                    End If
                 End If
            Case "6#":
                 fgSelect.Col = 11
                 oldYSWIRAM0.SWIRAMXID = Val(fgSelect.Text)
                 If oldYSWIRAM0.SWIRAMXID > 0 Then YSWIRAM0_STA_Reset
             Case "7", "7E", "7E*", "7J":
                 If cmdSelect_SQL_K = "7J" Then
                    fgSelect.Col = 11: oldYSWISAB0.SWISABSWID = Val(fgSelect.Text)
                 Else
                    If X < 1450 Then
                        fgSelect.Col = 13
                    Else
                        fgSelect.Col = 14
                    End If
                    
                    oldYSWISAB0.SWISABSWID = Val(fgSelect.Text)
                    'MAJ 2017/02/23 : Suite Ã  la demande de Christian Reol sur les message 198 une rÃ©affectation des id est faite pour rÃ©cupÃ©rer le nouveau id avec la date d'&chÃ©ance
                    If fgSelect.Col = 14 Then
                        
                        verifNewId = reaff_id_echeance
                        If verifNewId > -1 Then
                            oldYSWISAB0.SWISABSWID = verifNewId
                        End If
                    End If
                    If oldYSWISAB0.SWISABSWID = 0 Then
                           fgSelect.Col = 13: oldYSWISAB0.SWISABSWID = Val(fgSelect.Text)
                    End If
                    fgSelect.Col = 15: mSWIECHSWIL = Val(fgSelect.Text)
                 End If
                wX = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABSWID = " & oldYSWISAB0.SWISABSWID
                Set rsSab = cnsab.Execute(wX)
                If Not rsSab.EOF Then
                    Mesg_aid = rsSab("SWISABWID1")
                    mesg_s_umidl = rsSab("SWISABWIDL")
                    mesg_s_umidh = rsSab("SWISABWIDH")
                    oldYGOSDOS0.GOSDOSWES = rsSab("SWISABWES")
                    oldYGOSDOS0.GOSDOSWBIC = rsSab("SWISABWBIC")
                    oldYGOSDOS0.GOSDOSWMTK = rsSab("SWISABWMTK")
                    oldYGOSDOS0.GOSDOSWTRN = rsSab("SWISABWTRN")
                    oldYGOSDOS0.GOSDOSWDEV = rsSab("SWISABWDEV")
                    oldYGOSDOS0.GOSDOSWMTD = rsSab("SWISABWMTD")
                    oldYGOSEVE0.GOSEVESWID = oldYSWISAB0.SWISABSWID
                    fgDetail_Display
                End If
             Case "7#":
                  fgSelect.Col = 13: oldYSWIECH0.SWIECHSWID = Val(fgSelect.Text)
                  fgSelect.Col = 16: oldYSWIECH0.SWIECHSEQ0 = Val(fgSelect.Text)
                 wX = "select * from " & paramIBM_Library_SABSPE & ".YSWIECH0" _
                    & " where SWIECHSWID = " & oldYSWIECH0.SWIECHSWID & " and  SWIECHSEQ0 = " & oldYSWIECH0.SWIECHSEQ0
                Set rsSab = cnsab.Execute(wX)
                If Not rsSab.EOF Then
                    Call rsYSWIECH0_GetBuffer(rsSab, oldYSWIECH0)
                    Select Case oldYSWIECH0.SWIECHSTA
                        Case "#": mnuYSWIECH0_Res.Enabled = False: mnuYSWIECH0_Ann.Enabled = True
                        Case "A": mnuYSWIECH0_Ann.Enabled = False: mnuYSWIECH0_Res.Enabled = True
                        Case Else: mnuYSWIECH0_Ann.Enabled = True: mnuYSWIECH0_Res.Enabled = True
                    End Select
                    Me.PopupMenu mnuYSWIECH0, vbPopupMenuLeftButton
                End If

    End Select
        
   End If
End If
Wait_SS 0
fgSelect.LeftCol = 0

End Sub
Private Function reaff_id_echeance() As Long
Dim newid As Long
Dim newdate As Long
Dim xSql As String
Dim monRecordset As New ADODB.Recordset


    reaff_id_echeance = -1
    fgSelect.Col = 3
    If Val(Trim(fgSelect.Text)) <> 198 Then
        Exit Function
    End If
    fgSelect.Col = 2
    newdate = Val(Mid(fgSelect.Text, 7, 4) & Mid(fgSelect.Text, 4, 2) & Mid(fgSelect.Text, 1, 2))
    fgSelect.Col = 12
    newid = Val(Trim(Mid(fgSelect.Text, 8)))
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABOPEN = " & newid & " and SWISABWAMJ = " & newdate _
    & " and SWISABWMTK = 198"
    Set monRecordset = cnsab.Execute(xSql)
    Do While Not monRecordset.EOF
        reaff_id_echeance = CLng(monRecordset("SWISABSWID"))
        Exit Do
    Loop
    Set monRecordset = Nothing
    
End Function


Private Sub fgSwift_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xField As String, xSql As String
If fgSwift.Rows > 1 Then
    If X <= 500 Then
    fgSwift.Col = 0
        xField = Trim(fgSwift.Text)
        Call arrMT_Type_Scan(oldYSWISAB0.SWISABWMTK)
        mnuSWIBICBIC.Caption = arrMT_Fields_Scan(xField)
        Me.PopupMenu mnuZSWIBIC0, vbPopupMenuLeftButton
    Else
        fgSwift.Col = 1
        If ZSWIBIC0_Select(fgSwift.Text) <> "" Then
            mnuSWIBICBIC.Caption = Trim(rsSabX("SWIBICIN1")) & "  " & Trim(rsSabX("SWIBICVIL")) & "  " & Trim(rsSabX("SWIBICCOM"))
            Me.PopupMenu mnuZSWIBIC0, vbPopupMenuLeftButton
        End If
    End If
End If
fgSwift.Col = 0
End Sub


Private Sub filDoc_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

If mfilDoc_Path <> filDoc.path Then cmdGOSDOSPJ.Visible = True

oldFileName = filDoc.path & "\" & filDoc.FileName
newDirPath = paramGOSDOS_Path & oldYGOSDOS0.GOSDOSIDD
newFileName = filDoc.FileName
newFileExtension = fileName_Extension(filDoc.FileName)
txtGOSEVETXT = filDoc.FileName

    If Dir(oldFileName) <> "" Then
        If Right(UCase(oldFileName), 4) = ".LNK" Then
            'faire apparaitre les fichiers contenus dans le lien, dans la fenêtre filDoc
            Dim Obj As Object
            Dim ShortCut As Object
            Set Obj = CreateObject("WScript.Shell")
            Set ShortCut = Obj.CreateShortcut(oldFileName)
            dirListBox.path = ShortCut.TargetPath
            filDoc.Enabled = True
            Set ShortCut = Nothing
            Set Obj = Nothing
        Else
            Select Case newFileExtension
             Case "DOCX": Call frmElpPrt.WinWord(oldFileName)
             Case "DOC": Call frmElpPrt.WinWord(oldFileName)
             Case "XLS": Call frmElpPrt.Excel(oldFileName)
             Case "XLSX": Call frmElpPrt.Excel(oldFileName)
             Case "PDF": Call frmElpPrt.Acrord32(oldFileName)
             Case "TXT": Call frmElpPrt.WordPad(oldFileName) 'NotePad(X)
             Case "RTF": Call frmElpPrt.WordPad(oldFileName)
             Case Else: Call frmElpPrt.IExplore(oldFileName)
            End Select
            fraPJ.Visible = False
        End If
    End If

'cmdEVE_Ok_Click
Me.Enabled = True: Me.MousePointer = 0
On Error Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If chkSIDE_DB_Show = "1" Then frmSIDE_DB.Hide
If chkSAB_Dossier_DB_Show = "1" Then frmSAB_Dossier_DB.Hide
    cnSIDE_DB.Close
    Set cnSIDE_DB = Nothing

End Sub

Private Sub lstMail_MT_CC_Click()
If lstMail_MT_CC.Visible Then lstMail_MT_CC_TXT

End Sub

Private Sub lstMail_MT_Message_Click()
Dim X As String
X = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where GOSEVEIDD = -1 and GOSEVEGSRV = '" & currentSSIWINUNIT & "'" _
     & " and GOSEVENAT = 'Mail' and  SUBSTRING(GOSEVETXT , 1 , 10) = '" & Trim(Mid$(lstMail_MT_Message, 1, 10)) & "'"
Set rsSab = cnsab.Execute(X)


If Not rsSab.EOF Then
    X = Trim(rsSab("GOSEVETXT"))
    txtMail_MT_Message = Mid$(X, 31, Len(X) - 30)
Else
    txtMail_MT_Message = ""
End If
lstMail_MT_Message.Visible = False

End Sub

Private Sub lstMail_MT_To_Click()
If lstMail_MT_To.Visible Then lstMail_MT_To_TXT

End Sub

Private Sub lstParam_Click()
Dim X As String
Select Case lstParam.Text
    Case "Motif":
        Call lstParam_GOSDOSLABK_Load("StaC", "cboParam_StaC")
        Call lstParam_GOSDOSLABK_Load("DOS", "")
    Case "Note": Call lstParam_GOSDOSLABK_Load("NOTE", "")
    Case "SwFR": Call lstParam_GOSDOSLABK_Load("SwFR", "")
    Case "SwGB": Call lstParam_GOSDOSLABK_Load("SwGB", "")
    Case "Mail": Call lstParam_GOSDOSLABK_Load("Mail", "")
    Case "Stat Code": Call lstParam_GOSDOSLABK_Load("StaC", "")
    Case "Stat Pays": Call lstParam_GOSDOSLABK_Load("StaP", "")
    Case "Motif => Swift": Call lstParam_GOSDOSLABK_Load("DoSw", "")
End Select
If lstParam_GOSDOSLABK.ListCount = 1 Then lstParam_GOSDOSLABK_Click
End Sub

Private Sub lstParam_GOSDOSLABK_Click()
Dim xSql As String, X As String

'_______________________________________________
fraParam_GOSDOSLABK.Visible = False
'_______________________________________________
cmdParam_GOSDOSLABK_Delete.Visible = False
cmdParam_GOSDOSLABK_Update.Visible = False
cmdParam_GOSDOSLABK_Add.Visible = False


xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where GOSEVEIDD = -1 and GOSEVEGSRV = '" & currentSSIWINUNIT & "'" _
     & " and GOSEVENAT = '" & oldParam.GOSEVENAT & "' and  SUBSTRING(GOSEVETXT , 1 , 10) = '" & Trim(Mid$(lstParam_GOSDOSLABK, 1, 10)) & "'"
Set rsSab = cnsab.Execute(xSql)

txtParam_GOSDOSLABK.Enabled = arrHab(18)
cmdParam_GOSDOSLABK_Add.Visible = arrHab(18)

If Not rsSab.EOF Then
    V = rsYGOSEVE0_GetBuffer(rsSab, oldParam)
    X = Trim(oldParam.GOSEVETXT)
    txtParam_GOSDOSLABK = Trim(Mid$(X, 1, 10))
    txtParam_GOSDOSLABK_J = Mid$(X, 12, 3)
    Call cbo_Scan(Mid$(X, 16, 3), cboParam_GOSDOSLABK_GSrv)
    If Len(X) > 19 Then
        txtParam_GOSDOSLABK_Lib = Mid$(X, 31, Len(X) - 30)
    Else
        txtParam_GOSDOSLABK_Lib = ""
    End If
    txtParam_GOSDOSLABK_Lib.Enabled = arrHab(18)
    cmdParam_GOSDOSLABK_Delete.Visible = arrHab(18)
    cmdParam_GOSDOSLABK_Update.Visible = arrHab(18)
    If oldParam.GOSEVENAT = "DOS " Then Call cbo_Scan(Mid$(X, 20, 3), cboParam_StaC)
Else
    txtParam_GOSDOSLABK = ""
    txtParam_GOSDOSLABK_J = 0
    Call cbo_Scan(currentSSIWINUNIT, cboParam_GOSDOSLABK_GSrv)
    txtParam_GOSDOSLABK_Lib = ""
    If cboParam_StaC.ListCount > 0 Then cboParam_StaC.ListIndex = -1
End If

cboParam_StaC.Visible = False: lblParam_StaC.Visible = False
txtParam_GOSDOSLABK_Lib.Visible = True

Select Case oldParam.GOSEVENAT
    Case "Mail", "StaC", "StaP":
        txtParam_GOSDOSLABK_J.Visible = False: lblParam_GOSDOSLABK_J.Visible = False
        cboParam_GOSDOSLABK_GSrv.Visible = False: lblParam_GOSDOSLABK_GSrv.Visible = False
        
        If oldParam.GOSEVENAT = "StaP" Then txtParam_GOSDOSLABK_Lib.Visible = False
       
    Case Else
        txtParam_GOSDOSLABK_J.Visible = True: lblParam_GOSDOSLABK_J.Visible = True
        cboParam_GOSDOSLABK_GSrv.Visible = True: lblParam_GOSDOSLABK_GSrv.Visible = True
        If oldParam.GOSEVENAT = "DOS " Then
            cboParam_StaC.Visible = True: lblParam_StaC.Visible = True
        End If
End Select
fraParam_GOSDOSLABK.Visible = True
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

cmdMail_MT.Visible = False

If SSTab1.Tab = 1 Then
    Select Case tabParam.Tab
        Case 0: fraParam_GOSDOSLABK.Visible = False: lstParam_GOSDOSLABK.Visible = False
        Case 1: 'fraParam_GOSDOSMAIL.Visible = False
        Case Else: SSTab1.Tab = 0
    End Select
    Exit Sub
End If


If fraSwift.Visible Then:    fraSwift.Visible = False: Exit Sub

If fraEVE.Visible Then
    fraEVE.Visible = False: fgEVE.Enabled = True
    fgModèle.Visible = False: fraEVE_Swift.Visible = False: fraPJ.Visible = False
    Exit Sub
End If
'_________________________________________________________________
If fgFree.Visible Then fgFree.Visible = False: Exit Sub
If lstMail_MT_To.Visible Then lstMail_MT_To.Visible = False: Exit Sub
If lstMail_MT_CC.Visible Then lstMail_MT_CC.Visible = False: Exit Sub
If lstMail_MT_Message.Visible Then lstMail_MT_Message.Visible = False: Exit Sub
If fraMail_MT.Visible Then fraMail_MT.Visible = False: blnGOSEVE_Mail = False:    Exit Sub


If fgModèle.Visible Then fgModèle.Visible = False: Exit Sub

If fraPJ.Visible Then fraPJ.Visible = False: Exit Sub

If fraSWISABKSRV.Visible Then fraSWISABKSRV.Visible = False: Exit Sub

If fraList.Visible Then fraList.Visible = False: fraDetail.Visible = False:   Exit Sub
    
If fraMail_MT.Visible Then fraMail_MT.Visible = False:    Exit Sub

If fraDetail.Visible Then fraDetail.Visible = False: Exit Sub

If fgSelect.Visible Then fgSelect.Visible = False: Exit Sub


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






Public Sub fgSelect_ForeColor(lColor As Long)
For I = 0 To fgSelect_arrIndex
  fgSelect.Col = I: fgSelect.CellForeColor = lColor
Next I

End Sub







Private Sub mnuPrint_Detail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdPrint_YGOSDOS0 True

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint_Recap_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdPrint_YGOSDOS0 False

Me.Enabled = True: Me.MousePointer = 0
End Sub




















Private Sub mnuPrint2_Excel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim xTitle As String
Select Case cmdSelect_SQL_K
    Case "7", "7E":
            xTitle = cmdSelect_SQL_K & " : Messages SWIFT à émettre/recevoir "
            Call MSflexGrid_Excel("", "BIA_GOS", xTitle, fgSelect, fgSelect.Cols - 1)
    Case "7J":
            xTitle = cmdSelect_SQL_K & " : Journal des messages SWIFT MT300 et MT320 "
            Call MSflexGrid_Excel("", "BIA_GOS", xTitle, fgSelect, fgSelect.Cols - 1)
    Case Else:
        Call cmdPrint_Excel
End Select

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint2_Mail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim xSubject As String, xAPP As String, xTitle As String, xDest As String

xAPP = "BIA_GOS : " & cboSelect_SQL.Text
xSubject = xAPP & " (" & dateImp10_S(DSys) & " " & Time & ")"
xTitle = xSubject
xDest = currentSSIWINMAIL

Select Case cmdSelect_SQL_K
    Case "6": xTitle = "Liste de rapprochement des messages 300 et 320 du " _
                     & dateImp10_S(wAmjMin) & " au " & dateImp10_S(wAmjMax)
            If blnAuto Then xDest = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire("S32")
            Call MSFlexGrid_SendMail(xDest, xAPP, xSubject, xTitle, fgSelect, 7)

    Case "6E": xTitle = "Messages SWIFT MT300 et MT320 en attente de rapprochement E/S"
            If blnAuto Then xDest = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire("S32_RAM")
            Call MSFlexGrid_SendMail(xDest, xAPP, xSubject, xTitle, fgSelect, 7)
    
        Case "7", "7E": xTitle = cmdSelect_SQL_K & " : Messages SWIFT à émettre/recevoir "
            If blnAuto Then xDest = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire("S32_RAM")
            Call MSFlexGrid_SendMail(xDest, xAPP, xSubject, xTitle, fgSelect, 15)
        Case "7J": xTitle = cmdSelect_SQL_K & " : Journal des évènements du suivi des messages SWIFT à émettre/recevoir "
            If blnAuto Then xDest = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire("S32_RAM")
            Call MSFlexGrid_SendMail(xDest, xAPP, xSubject, xTitle, fgSelect, 15)
End Select




Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuYSWIECH0_Ann_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call YSWIECH0_Update_STA("A")
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuYSWIECH0_Res_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call YSWIECH0_Update_STA("#")
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub rtfPJ_Click()
rtfPJ.Top = 120
rtfPJ.Left = 120
rtfPJ.Width = fraPJ.Width - 240
rtfPJ.Height = fraPJ.Height - 240


End Sub


Private Sub cboEVE_Swift_20_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtEVE_Swift_21_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtGOSDOSCLI_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub

Private Sub txtGOSDOSCLI_Validate(Cancel As Boolean)
Dim X As String
Call fraDetail_LAB_Control_Client(X)

End Sub


Private Sub txtGOSEVETXT_KeyPress(KeyAscii As Integer)
If Mid$(cboGOSEVENAT, 1, 2) = "Sw" Then KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtList_Add_Change()
If Trim(txtList_Add) <> "" Then cmdList_Display.Visible = True
End Sub

Private Sub txtList_Add_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtMail_MT_CC_Click()
lstMail_MT_Message.Visible = False
lstMail_MT_CC.Visible = False
lstMail_MT_To.Visible = False

End Sub


Private Sub txtMail_MT_Message_Click()
lstMail_MT_Message.Visible = False
lstMail_MT_CC.Visible = False
lstMail_MT_To.Visible = False

End Sub


Private Sub txtMail_MT_To_Click()
lstMail_MT_Message.Visible = False
lstMail_MT_CC.Visible = False
lstMail_MT_To.Visible = False

End Sub


Private Sub txtParam_GOSDOSLABK_J_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtParam_GOSDOSLABK_Lib_KeyPress(KeyAscii As Integer)
If oldParam.GOSEVENAT = "SwFR" Or oldParam.GOSEVENAT = "SwGB" Then
    KeyAscii = convUCase(KeyAscii)
End If
End Sub

Private Sub txtParam_SAA_K2_KeyPress(KeyAscii As Integer)
    KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtParam_SAA_MTD_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub

Private Sub cmdParam_SAA_Quit_Click()
fraParam_SAA.Visible = False
End Sub

Private Sub txtSelect_3_GOSDOSIDD_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub txtSelect_3_GOSDOSIDD_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtSelect_GOSDOSIAMJ_Max_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub txtSelect_GOSDOSIAMJ_Max_Click()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub txtSelect_GOSDOSIAMJ_Min_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub txtSelect_GOSDOSIAMJ_Min_Click()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub txtSelect_GOSDOSWMTD_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub txtSelect_GOSDOSWMTD_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub






Private Sub txtSelect_rTextField_Code_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub


Private Sub txtSelect_rTextField1_KeyPress(KeyAscii As Integer)
If chkSelect_rTextField = "1" Then KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSelect_rTextField2_KeyPress(KeyAscii As Integer)
If chkSelect_rTextField = "1" Then KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelect_SWISABOPEN_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub txtSelect_SWISABOPEN_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelect_GOSDOSWTRN_Change()
If blnControl Then cmdSelect_Clear

End Sub

Private Sub txtSelect_GOSDOSWTRN_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Public Sub fraDetail_LAB_Init()
Dim X As String, wBIC As String
arrYGOSEVE0_Nb = 0
cboGOSDOSLABK.ListIndex = -1
cboGOSDOSPAYS.ListIndex = 0
cboGOSDOSRCOM.ListIndex = 0
cboGOSDOSGSRV.ListIndex = 0
lblGOSDOSIAMJ = ""
lblGOSDOSUAMJ = ""
txtGOSDOSCLI = "": libGOSDOSCLI = ""
txtGOSDOSTXT = ""
lblGOSDOSSTAD = "Nouveau dossier"
libGOSDOSLABK.Visible = False
cboGOSDOSLABK.Visible = True
fraDetail_Y.BackColor = vbMagenta
Call DTPicker_Set(txtGOSDOSECHD, DSys) '
fraDetail_LAB.Enabled = arrHab(5)
cmdDetail_LAB_Ok.Caption = constAjouter
cmdDetail_LAB_Ok.Visible = arrHab(5) 'fraDetail_LAB.Enabled
cmdDetail_Lab_Link.Visible = arrHab(5)
If oldYSWISAB0.SWISABWIDL = 0 Then
    Call MsgBox("Erreur sporadique non résolue, prévenir JPL", vbCritical, "BIA_GOS : fraDetail_LAB_Init")
    Exit Sub
End If
oldYGOSDOS0.GOSDOSWBIC = oldYSWISAB0.SWISABWBIC
oldYGOSDOS0.GOSDOSWES = oldYSWISAB0.SWISABWES
oldYGOSDOS0.GOSDOSWMTK = oldYSWISAB0.SWISABWMTK
oldYGOSDOS0.GOSDOSWTRN = oldYSWISAB0.SWISABWL20
oldYGOSDOS0.GOSDOSWMTD = oldYSWISAB0.SWISABWMTD
oldYGOSDOS0.GOSDOSWDEV = oldYSWISAB0.SWISABWDEV

oldYGOSDOS0.GOSDOSWID1 = oldYSWISAB0.SWISABWID1
oldYGOSDOS0.GOSDOSWIDH = oldYSWISAB0.SWISABWIDH
oldYGOSDOS0.GOSDOSWIDL = oldYSWISAB0.SWISABWIDL

oldYGOSEVE0.GOSEVESWID = oldYSWISAB0.SWISABSWID
oldYGOSDOS0.GOSDOSISRV = currentSSIWINUNIT

oldYGOSDOS0.GOSDOSIAMJ = DSys
oldYGOSEVE0.GOSEVENAT = "Sus*"

X = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 ," & paramIBM_Library_SABSPE & ".YSWILNK0" _
     & " where SWISABWID1 = " & Mesg_aid _
     & " and  SWISABWIDL = " & mesg_s_umidl _
     & " and  SWISABWIDH = " & mesg_s_umidh _
     & " and  SWISABSWID = SWILNKSWID"
Set rsSab = cnsab.Execute(X)
'____________________________________________________________________________________
X = ""
Do While Not rsSab.EOF
    X = X & rsSab("SWILNKAPPC") & " " & rsSab("SWILNKAPPN") & vbCrLf
    rsSab.MoveNext
Loop


If X <> "" Then
    X = MsgBox("Ce message swift est relié à " & vbCrLf & X _
      & "Voulez-vous continuer ?", vbQuestion + vbYesNo + vbDefaultButton2, "BIA_GOS : contrôle")
    If X = vbNo Then Exit Sub

End If
'____________________________________________________________________________________

X = "select * from " & paramIBM_Library_SABSPE & ".YGOSDOS0 " _
     & " where GOSDOSWID1 = " & Mesg_aid _
     & " and  GOSDOSWIDL = " & mesg_s_umidl _
     & " and  GOSDOSWIDH = " & mesg_s_umidh
Set rsSab = cnsab.Execute(X)

X = ""
Do While Not rsSab.EOF
    X = X & rsSab("GOSDOSIDD") & vbCrLf
    rsSab.MoveNext
Loop


If X <> "" Then
    X = MsgBox("Dossier concernant ce message swift :" & vbCrLf & X & "Voulez-vous créer un autre dossier ?", vbQuestion + vbYesNo + vbDefaultButton2, "BIA_GOS : contrôle")
    If X = vbNo Then
        fraDetail.Visible = False
    Else
        fraDetail.Visible = True
    End If
Else
    fraDetail.Visible = True

End If
'____________________________________________________________________________________

Call lstParam_GOSDOSLABK_Load("DOS", "cboGOSDOSLABK")

mYGOSDOS0_Fct = "": mYGOSEVE0_Fct = "": mYSWILNK0_Fct = "": mYSWISAB0_Fct = "": mZSWIENA0_Fct = ""
wBIC = Mid$(oldYSWISAB0.SWISABWBIC, 1, 8)

If wBIC = "SOGEFRPP" Then Call fraDetail_LAB_Init_BIC(wBIC)

If Not IsNull(fraDetail_LAB_Init_GOSDOSCLI(wBIC)) Then
    Call fraDetail_LAB_Init_BIC(wBIC)
    Call fraDetail_LAB_Init_GOSDOSCLI(wBIC)
End If

fraDetail_LAB.Visible = True
fraDetail_C.Visible = True
End Sub
Public Sub fraDetail_LAB_Init_Vierge()
Dim X As String, wBIC As String

Call rsYGOSDOS0_Init(oldYGOSDOS0) '$JPL 2014-11-25

arrYGOSEVE0_Nb = 0
cboGOSDOSLABK.ListIndex = -1
cboGOSDOSPAYS.ListIndex = 0
cboGOSDOSRCOM.ListIndex = 0
cboGOSDOSGSRV.ListIndex = 0
lblGOSDOSIAMJ = ""
lblGOSDOSUAMJ = ""
txtGOSDOSCLI = "": libGOSDOSCLI = ""
txtGOSDOSTXT = ""
lblGOSDOSSTAD = "Nouveau dossier"
libGOSDOSLABK.Visible = False
cboGOSDOSLABK.Visible = True
fraDetail_Y.BackColor = vbMagenta
Call DTPicker_Set(txtGOSDOSECHD, DSys) '
fraDetail_LAB.Enabled = arrHab(5)
cmdDetail_LAB_Ok.Caption = constAjouter
cmdDetail_LAB_Ok.Visible = arrHab(5) 'fraDetail_LAB.Enabled
cmdDetail_Lab_Link.Visible = arrHab(5)
oldYGOSDOS0.GOSDOSIDD = 0
oldYGOSDOS0.GOSDOSWBIC = ""
oldYGOSDOS0.GOSDOSWES = ""
oldYGOSDOS0.GOSDOSWMTK = ""
oldYGOSDOS0.GOSDOSWTRN = ""
oldYGOSDOS0.GOSDOSWMTD = 0
oldYGOSDOS0.GOSDOSWDEV = ""

oldYGOSDOS0.GOSDOSWID1 = 0
oldYGOSDOS0.GOSDOSWIDH = 0
oldYGOSDOS0.GOSDOSWIDL = 0
Mesg_aid = 0
mesg_s_umidl = 0
mesg_s_umidh = 0

oldYGOSEVE0.GOSEVESWID = 0
oldYGOSDOS0.GOSDOSISRV = currentSSIWINUNIT

oldYGOSDOS0.GOSDOSIAMJ = DSys
oldYGOSEVE0.GOSEVENAT = "Sus*"
Call rsYGOSDOS0_Init(newYGOSDOS0)
'____________________________________________________________________________________

'____________________________________________________________________________________

Call lstParam_GOSDOSLABK_Load("DOS", "cboGOSDOSLABK")

mYGOSDOS0_Fct = "New": mYGOSEVE0_Fct = "": mYSWILNK0_Fct = "": mYSWISAB0_Fct = "": mZSWIENA0_Fct = ""
wBIC = Mid$(oldYSWISAB0.SWISABWBIC, 1, 8)
tabDetail.Tab = 1: tabDetail.Caption = "Création d'un dossier"
tabDetail.Tab = 0: tabDetail.Caption = ""
libDetail_SWISABSWID = ""

fgDetail.Visible = False

fraDetail.Visible = True
fraDetail_LAB.Visible = True
fraDetail_C.Visible = True
End Sub

Public Sub fraDetail_LAB_Init_BIC(lBIC As String)
Dim X As String

X = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB1 " _
     & " where SWISAB1ID  = " & oldYSWISAB0.SWISABSWID
Set rsSab = cnsab.Execute(X)

If Not rsSab.EOF Then
    If Trim(rsSab("SWISABW57A")) <> "" Then lBIC = Mid$(rsSab("SWISABW57A"), 1, 8)

    If lBIC = "BIARFRPP" Then
        If oldYSWISAB0.SWISABWMTK = "202" Then
            If Trim(rsSab("SWISABW52A")) <> "" Then lBIC = Mid$(rsSab("SWISABW52A"), 1, 8)
        End If
    End If
End If
    
End Sub

Public Function fraDetail_LAB_Init_GOSDOSCLI(lBIC As String)
Dim X As String, wBIC As String

If lBIC = "BIARFRPP" Then
    fraDetail_LAB_Init_GOSDOSCLI = fraDetail_LAB_Init_GOSDOSCLI_BIARFRPP(fgDetail_59)
Else

    X = "select CLIENACLI , CLIENARA1 , CLIENARA2 , CLIENARES , CLIENARSD from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
         & " where CLIENASIG  = '" & lBIC & "'"
    Set rsSab = cnsab.Execute(X)
    
    If Not rsSab.EOF Then
        txtGOSDOSCLI = rsSab("CLIENACLI"): libGOSDOSCLI = Trim(rsSab("CLIENARA1")) & Trim(rsSab("CLIENARA2"))
        Call cbo_Scan(rsSab("CLIENARES"), cboGOSDOSRCOM)
        Call cbo_Scan(Trim(rsSab("CLIENARSD")), cboGOSDOSPAYS)
        fraDetail_LAB_Init_GOSDOSCLI = Null
    Else
        fraDetail_LAB_Init_GOSDOSCLI = lBIC
    End If
End If
End Function

Public Function fraDetail_LAB_Init_GOSDOSCLI_BIARFRPP(lTxt As String)
Dim X As String, K As Integer
fraDetail_LAB_Init_GOSDOSCLI_BIARFRPP = "?"
If Mid$(lTxt, 1, 1) = "/" Then

    K = InStr(lTxt, Asc13)
    If K < 1 Then K = Len(lTxt) + 1
    X = Mid$(lTxt, 2, K - 2) & "                         "
    X = Replace(X, " ", "")
    X = Replace(X, "IBAN:", "")
    X = Replace(X, "IBAN", "")
    K = InStr(X, "1217900001")
    If K > 0 Then
        X = Mid$(X, K + 10, 5)
    Else
        X = Mid$(X, 1, 5)
    End If
    

    X = "select CLIENACLI , CLIENARA1 , CLIENARA2 , CLIENARES , CLIENARSD from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
         & " where CLIENACLI   = '00" & X & "'"
    Set rsSab = cnsab.Execute(X)
    
    If Not rsSab.EOF Then
        txtGOSDOSCLI = rsSab("CLIENACLI"): libGOSDOSCLI = Trim(rsSab("CLIENARA1")) & Trim(rsSab("CLIENARA2"))
        Call cbo_Scan(rsSab("CLIENARES"), cboGOSDOSRCOM)
        Call cbo_Scan(Trim(rsSab("CLIENARSD")), cboGOSDOSPAYS)
        fraDetail_LAB_Init_GOSDOSCLI_BIARFRPP = Null
    End If
End If
End Function

Public Function fraDetail_LAB_Control()
Dim V, X As String, blnOk As Boolean, K As Integer, wMsgBox As String
Dim Nb As Long
On Error GoTo Exit_sub
'fraDetail_Update.Enabled = False

blnYGOSDOS0_Display = False
newYGOSDOS0 = oldYGOSDOS0
wMsgBox = ""
blnOk = False
fraDetail_LAB_Control = "?"

newYGOSDOS0.GOSDOSPAYS = Mid$(cboGOSDOSPAYS, 1, 2)
newYGOSDOS0.GOSDOSLABK = Mid$(cboGOSDOSLABK, 1, 10)
If Trim(newYGOSDOS0.GOSDOSLABK) = "" Then
    wMsgBox = wMsgBox & " - préciser le motif " & vbCrLf
End If

If Not IsNull(fraDetail_LAB_Control_Client(newYGOSDOS0.GOSDOSCLI)) Then
    wMsgBox = wMsgBox & " - client inconnu " & vbCrLf
End If

If Trim(newYGOSDOS0.GOSDOSWBIC) = "" Then
    newYGOSDOS0.GOSDOSWBIC = fraDetail_LAB_Control_Client_BIC(newYGOSDOS0.GOSDOSCLI)
End If
If Trim(newYGOSDOS0.GOSDOSWMTK) = "" Then newYGOSDOS0.GOSDOSWMTK = "BIA"


newYGOSDOS0.GOSDOSRCOM = Mid$(cboGOSDOSRCOM, 1, 3)
newYGOSDOS0.GOSDOSGSRV = Mid$(cboGOSDOSGSRV, 1, 4)



X = ""
fgDetail.Col = 1
For K = 1 To fgDetail.Rows - 1
    fgDetail.Row = K
    If fgDetail.CellForeColor = vbRed Then
        X = X & Format$(K, "00")
    End If
Next K
If Len(X) > 20 Then wMsgBox = wMsgBox & " - maximum 10 lignes topées " & vbCrLf
newYGOSDOS0.GOSDOSITOP = Mid$(X, 1, 20)

fgDetail.Col = 0

''newYGOSDOS0.GOSDOSISRV = currentSSIWINUNIT

''newYGOSDOS0.GOSDOSIAMJ = DSys

Call DTPicker_Control(txtGOSDOSECHD, X)
newYGOSDOS0.GOSDOSECHD = X
If newYGOSDOS0.GOSDOSECHD < DSys Then wMsgBox = wMsgBox & " - date échéance  < aujourd'hui" & vbCrLf

newYGOSEVE0 = oldYGOSEVE0
If Trim(txtGOSDOSTXT) = "" Then wMsgBox = wMsgBox & " - préciser l'action à faire " & vbCrLf
newYGOSEVE0.GOSEVETXT = Trim(txtGOSDOSTXT)
'newYGOSEVE0.GOSEVENAT = "Sus*"
newYGOSEVE0.GOSEVEGSRV = newYGOSDOS0.GOSDOSGSRV

If wMsgBox <> "" Then
    fraDetail_LAB_Control = "?"
    Call MsgBox(wMsgBox, vbCritical, "BIA_GOS : contrôle détail")
Else
    fraDetail_LAB_Control = Null
End If

Exit_sub:
End Function
Public Function fraEVE_Control()
Dim V, X As String, blnOk As Boolean, K As Integer, wMsgBox As String
Dim Nb As Long, X1 As String, wAmj As Long
On Error GoTo Exit_sub

newYGOSDOS0 = oldYGOSDOS0
'==========================
oldYGOSEVE0 = arrYGOSEVE0(arrYGOSEVE0_Nb)
oldYGOSEVE0.GOSEVESWID = 0
oldYGOSEVE0.GOSEVESTAE = ""

newYGOSEVE0 = oldYGOSEVE0

wMsgBox = ""
blnOk = False
fraEVE_Control = "?"

Call DTPicker_Control(txtGOSEVEECHD, X)
wAmj = CLng(X)
If wAmj <> oldYGOSDOS0.GOSDOSECHD Then
    If wAmj < DSys Then
        wMsgBox = wMsgBox & " - date échéance  < aujourd'hui" & vbCrLf
    Else
        newYGOSDOS0.GOSDOSECHD = wAmj
    End If
End If

newYGOSEVE0.GOSEVENAT = Mid$(cboGOSEVENAT, 1, 4)
If Trim(newYGOSEVE0.GOSEVENAT) = "" Then
    If arrHab(3) And newYGOSDOS0.GOSDOSECHD <> oldYGOSDOS0.GOSDOSECHD Then
        newYGOSEVE0.GOSEVENAT = "Ech"
    Else
        wMsgBox = wMsgBox & " - préciser la nature de l'événement " & vbCrLf
    End If
End If
If chkGOSEVEGSRV = 1 Then
    newYGOSEVE0.GOSEVEGSRV = Mid$(cboGOSEVEGSRV, 1, 3)
    newYGOSDOS0.GOSDOSGSRV = newYGOSEVE0.GOSEVEGSRV
End If


'If Mid$(newYGOSEVE0.GOSEVENAT, 2, 3) = "99>" Then
Select Case newYGOSEVE0.GOSEVENAT
    Case "SwFR", "SwGB":
        newYGOSEVE0.GOSEVENAT = "Swi>"
        X = Trim(txtGOSEVETXT)
        X = fraEVE_Control_Swift_Txt(X)
    
        If X <> "" Then wMsgBox = wMsgBox & " - caractères interdits : " & X & vbCrLf
        
        X = "   _:_           _:_                _:_                "
        Mid$(X, 1, 3) = Trim(cboEVE_Swift_MT)
        Mid$(X, 7, 11) = Trim(cboEVE_Swift_BIC)
        Mid$(X, 21, 16) = Trim(cboEVE_Swift_20)
        Mid$(X, 40, 16) = Trim(txtEVE_Swift_21)
        
        If Trim(txtGOSEVETXT) = "" Then
            newYGOSEVE0.GOSEVETXT = X
        Else
             newYGOSEVE0.GOSEVETXT = X & vbCrLf & vbCrLf & fraEVE_Control_Swift_Space(txtGOSEVETXT)
       End If
    Case "Swi=":
        newYGOSEVE0.GOSEVENAT = "Swi>"
        X = Trim(txtGOSEVETXT)
        X = fraEVE_Control_Swift_Txt(X)
    
        If X <> "" Then wMsgBox = wMsgBox & " - caractères interdits : " & X & vbCrLf
        
       newYGOSEVE0.GOSEVETXT = Mid$(duplic_YGOSEVE0.GOSEVETXT, 1, 56) & vbCrLf & vbCrLf & fraEVE_Control_Swift_Space(txtGOSEVETXT)
    Case "Swi+":
        X = "   _:_           _:_                _:_                "
        Mid$(X, 1, 3) = oldYSWISAB0.SWISABWMTK
        Mid$(X, 7, 11) = oldYSWISAB0.SWISABWBIC
        Mid$(X, 21, 16) = oldYSWISAB0.SWISABWL20
        Mid$(X, 40, 16) = oldYSWISAB0.SWISABWN20
        If Trim(txtGOSEVETXT) = "" Then
            newYGOSEVE0.GOSEVETXT = X
        Else
             newYGOSEVE0.GOSEVETXT = X & vbCrLf & vbCrLf & Trim(txtGOSEVETXT)
       End If
       
        mYSWISAB0_Fct = "Update"
        oldYSWISAB0 = m999_YSWISAB0
        newYSWISAB0 = oldYSWISAB0
        newYSWISAB0.SWISABXEVE = "G": newYSWISAB0.SWISABK999 = "G"
        If newYSWISAB0.SWISABKSRV <> Mid$(cboList_SWISABKSRV, 1, 3) Then newYSWISAB0.SWISABKSRV = Mid$(cboList_SWISABKSRV, 1, 3)
        If newYSWISAB0.SWISABOPEN = 0 Then
            newYSWISAB0.SWISABOPEC = "GOS"
            newYSWISAB0.SWISABOPEN = oldYGOSDOS0.GOSDOSIDD
            mYSWISAB0_Fct = "Update"
        End If
    Case "Ech"
    
    Case Else
        If Trim(txtGOSEVETXT) = "" Then
            If newYGOSEVE0.GOSEVENAT <> "Mail" Then wMsgBox = wMsgBox & " - préciser la description de l'événement " & vbCrLf
        End If
        newYGOSEVE0.GOSEVETXT = Trim(txtGOSEVETXT)

End Select

If wMsgBox <> "" Then
    fraEVE_Control = "?"
    Call MsgBox(wMsgBox, vbCritical, "BIA_GOS : contrôle événement")
Else
    fraEVE_Control = Null
    Select Case Trim(newYGOSEVE0.GOSEVENAT)
        Case "Val": newYGOSDOS0.GOSDOSSTAG = "V": mYGOSDOS0_Fct = "Update": mYGOSEVE0_Fct = "New"
        Case "Rej": newYGOSDOS0.GOSDOSSTAG = "R": mYGOSDOS0_Fct = "Update": mYGOSEVE0_Fct = "New"
        Case "AnnV", "AnnR": newYGOSDOS0.GOSDOSSTAG = " ": mYGOSDOS0_Fct = "Update": mYGOSEVE0_Fct = "New"
        
        Case "Ann": newYGOSDOS0.GOSDOSSTAD = "A": mYGOSDOS0_Fct = "Update": mYGOSEVE0_Fct = "New"
        Case "Clo": newYGOSDOS0.GOSDOSSTAD = "C": mYGOSDOS0_Fct = "Update": mYGOSEVE0_Fct = "New"
        Case "AnnC": newYGOSDOS0.GOSDOSSTAD = " ": mYGOSDOS0_Fct = "Update": mYGOSEVE0_Fct = "New"
        Case "Mail": mYGOSDOS0_Fct = "": mYGOSEVE0_Fct = ""
        Case "Ech": newYGOSDOS0.GOSDOSSTAD = " ": mYGOSDOS0_Fct = "Update": mYGOSEVE0_Fct = ""
        Case "Res*": newYGOSDOS0.GOSDOSSTAD = " ": newYGOSDOS0.GOSDOSSTAG = " ": mYGOSDOS0_Fct = "Update": mYGOSEVE0_Fct = "New"
        Case Else:
            If blnHab_YGOSEVE0_New Then
                newYGOSDOS0.GOSDOSSTAD = " ": mYGOSDOS0_Fct = "Update": mYGOSEVE0_Fct = "New"
            Else
                mYGOSDOS0_Fct = "Update": mYGOSEVE0_Fct = "New"
            End If

    End Select
    If blnGOSDOSSTAD_C Then newYGOSDOS0.GOSDOSSTAD = "C": mYGOSDOS0_Fct = "Update"
    If blnGOSDOSSTAD_X Then newYGOSDOS0.GOSDOSSTAD = "x": mYGOSDOS0_Fct = "Update"


End If

Exit_sub:
End Function

Public Sub cmdSelect_Clear()

lstErr.Clear
fgSelect.Visible = False
fraDetail.Visible = False
fraDetail_C.Visible = False
fraDetail_LAB.Enabled = False
fraEVE.Visible = False
fgEVE.Visible = False
lstW.Visible = False
fraMail_MT.Visible = False
fraList.Visible = False
fraSwift.Visible = False
fraSWISABKSRV.Visible = False
fgFree.Visible = False

'cmdSelect_Ok.Visible = False 'True
fraPJ.Visible = False
cmdEVE_Annulation.Visible = False
cmdEVE_Rejet.Visible = False
cmdEVE_Validation.Visible = False
cmdEVE_Restauration.Visible = False
cmdEVE_New.Visible = False
cmdEVE_Clôture.Visible = False

cmdMail_MT.Visible = False
fraMail_MT.Visible = False

If chkSelect_GOSDOSIAMJ = "1" Then
    txtSelect_GOSDOSIAMJ_Min.Visible = True
    txtSelect_GOSDOSIAMJ_Max.Visible = True
Else
    txtSelect_GOSDOSIAMJ_Min.Visible = False
    txtSelect_GOSDOSIAMJ_Max.Visible = False
End If

mYGOSDOS0_Fct = "": mYGOSEVE0_Fct = "": mYSWILNK0_Fct = "": mYSWISAB0_Fct = "": mZSWIENA0_Fct = ""
blnSwift_Display = False
blnYSWILNK0_Display = False
cmdSelect_Ok.BackColor = vbGreen
End Sub

Public Sub fraDetail_LAB_Display()
Dim X As String, xWhere As String

On Error Resume Next

If blnYGOSDOS0_New Then
    cmdSelect_SQL_Kbis = "3"
Else
    cmdSelect_SQL_Kbis = cmdSelect_SQL_K
End If


tabDetail.Tab = 1
tabDetail.Caption = "Détail du dossier n° " & oldYGOSDOS0.GOSDOSIDD

If oldYGOSDOS0.GOSDOSISRV <> currentSSIWINUNIT Then
    cboGOSDOSLABK.Visible = False
    libGOSDOSLABK.Visible = True
    libGOSDOSLABK.Caption = Trim(oldYGOSDOS0.GOSDOSLABK)
Else
    libGOSDOSLABK.Visible = False
    cboGOSDOSLABK.Visible = True
    Call cbo_Scan(Trim(oldYGOSDOS0.GOSDOSLABK), cboGOSDOSLABK)
End If


Call cbo_Scan(oldYGOSDOS0.GOSDOSPAYS, cboGOSDOSPAYS)
Call cbo_Scan(oldYGOSDOS0.GOSDOSRCOM, cboGOSDOSRCOM)
 

txtGOSDOSCLI = oldYGOSDOS0.GOSDOSCLI
X = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0  where CLIENACLI = '" & oldYGOSDOS0.GOSDOSCLI & "'"
Set rsSab = cnsab.Execute(X)

If Not rsSab.EOF Then
    libGOSDOSCLI = rsSab("CLIENARA1")
Else
    libGOSDOSCLI = "???????????????"
End If

X = oldYGOSDOS0.GOSDOSECHD
Call DTPicker_Set(txtGOSDOSECHD, X) '

'-----------------------------------------------------------------------
xWhere = " where GOSEVEIDD = " & oldYGOSDOS0.GOSDOSIDD & " order by GOSEVEIDE"
arrYGOSEVE0_SQL xWhere
If arrYGOSEVE0_Nb > 0 Then
    oldYGOSEVE0 = arrYGOSEVE0(1)
    txtGOSDOSTXT = Trim(oldYGOSEVE0.GOSEVETXT)
    Call cbo_Scan(oldYGOSEVE0.GOSEVEGSRV, cboGOSDOSGSRV)

Else
    txtGOSDOSTXT = "? ERREUR LECTURE YGOSEVE0"
    Call cbo_Scan(oldYGOSDOS0.GOSDOSGSRV, cboGOSDOSGSRV)
End If
fgEVE_Display

'_______________________________________________________


'-----------------------------------------------------------------------
cmdDetail_LAB_Ok.Caption = constModifier
cmdDetail_LAB_Ok.Visible = False
cmdDetail_Lab_Link.Visible = False
fraDetail_LAB.Enabled = False

cmdEVE_Validation.Visible = False
cmdEVE_Rejet.Visible = False
cmdEVE_Annulation.Visible = False
cmdEVE_New.Visible = False
cmdEVE_Ignore.Visible = False
cmdEVE_Restauration.Visible = False
cmdEVE_Clôture.Visible = False
cmdEVE_Invalidation.Visible = False

fraDetail_LAB.Enabled = False

fraEVE.Visible = False
fraSwift.Visible = False
fraDetail.BackColor = &HD8DFD8

blnHab_YGOSEVE0_New = False

'If cmdSelect_SQL_Kbis = "2" Or cmdSelect_SQL_Kbis = "3" Or cmdSelect_SQL_Kbis = "4" Then
If cmdSelect_SQL_Kbis = "3" Or cmdSelect_SQL_Kbis = "4" Then
    fraDetail.BackColor = &HC0E0FF
    If oldYGOSDOS0.GOSDOSSTAG = "V" Then
        cmdEVE_Invalidation.Caption = "Annulation de la Validation"
    Else
        cmdEVE_Invalidation.Caption = "Annulation du Rejet"
        End If


    Select Case oldYGOSDOS0.GOSDOSSTAD
        Case " "
            If currentSSIWINUNIT = oldYGOSDOS0.GOSDOSISRV Then
                
                fraDetail_LAB.Enabled = arrHab(3)
                cmdDetail_LAB_Ok.Visible = arrHab(3)
                'cmdDetail_Lab_Link.Visible = arrHab(3)
                cmdEVE_Annulation.Visible = arrHab(3)
                cmdEVE_New.Visible = arrHab(2)
                cmdEVE_Clôture.Visible = arrHab(3)
                blnHab_YGOSEVE0_New = arrHab(2)
                
                If oldYGOSDOS0.GOSDOSSTAG = " " Then
                    cmdEVE_Validation.Visible = arrHab(3)
                    cmdEVE_Rejet.Visible = arrHab(3)
                Else
                    cmdEVE_Invalidation.Visible = arrHab(3)
                End If
                
            Else
                If currentSSIWINUNIT = oldYGOSDOS0.GOSDOSGSRV Then
                    blnHab_YGOSEVE0_New = arrHab(2)
                    cmdEVE_New.Visible = arrHab(2)
                    fraDetail.BackColor = &HFFE0FF
                Else
                    If currentSSIWINUNIT = "S41" Then
                        cmdEVE_New.Visible = arrHab(2)
                        fraDetail.BackColor = &HFFE0FF
                    End If
                End If
            End If
            
        Case "x"
    
            Select Case currentSSIWINUNIT
                Case oldYGOSDOS0.GOSDOSISRV
                        cmdEVE_Annulation.Visible = arrHab(3)
                        cmdEVE_New.Visible = arrHab(2)
                        blnHab_YGOSEVE0_New = arrHab(2)
                        cmdEVE_Clôture.Visible = arrHab(3)
                        If oldYGOSDOS0.GOSDOSSTAG = " " Then
                            cmdEVE_Validation.Visible = arrHab(3)
                            cmdEVE_Rejet.Visible = arrHab(3)
                        Else
                            cmdEVE_Invalidation.Visible = arrHab(3)
                        End If
                Case oldYGOSDOS0.GOSDOSGSRV
                        cmdEVE_New.Visible = arrHab(2)
                        blnHab_YGOSEVE0_New = arrHab(2)
                        cmdEVE_Clôture.Visible = arrHab(3)
            End Select
            
        Case "C", "A"
            If currentSSIWINUNIT = oldYGOSDOS0.GOSDOSISRV _
            Or currentSSIWINUNIT = oldYGOSDOS0.GOSDOSUSRV Then
                cmdEVE_New.Visible = arrHab(2)
                cmdEVE_Restauration.Visible = arrHab(3)
                If oldYGOSDOS0.GOSDOSSTAD = "C" Then
                    cmdEVE_Invalidation.Visible = arrHab(3)
                    cmdEVE_Invalidation.Caption = "Annulation de la clôture"
                End If

            End If
    End Select
End If

End Sub


Public Function fraDetail_LAB_Control_Client(lGOSDOSCLI As String)

fraDetail_LAB_Control_Client = "?"

If Val(txtGOSDOSCLI) = 0 Then
    fraDetail_LAB_Control_Client = Null
    lGOSDOSCLI = ""
    libGOSDOSCLI = ""
Else
    lGOSDOSCLI = Format(Val(txtGOSDOSCLI), "0000000")

    X = "select CLIENARA1 , CLIENARES ,CLIENARSD from " & paramIBM_Library_SAB & ".ZCLIENA0  where CLIENACLI = '" & lGOSDOSCLI & "'"
    Set rsSab = cnsab.Execute(X)
    
    If Not rsSab.EOF Then
        fraDetail_LAB_Control_Client = Null
        libGOSDOSCLI = rsSab("CLIENARA1")
        If lGOSDOSCLI <> newYGOSDOS0.GOSDOSCLI Then
            Call cbo_Scan(rsSab("CLIENARES"), cboGOSDOSRCOM)
            Call cbo_Scan(Trim(rsSab("CLIENARSD")), cboGOSDOSPAYS)
        End If
    Else
        libGOSDOSCLI = "? client inconnu"
    End If
End If



End Function

Public Function fraDetail_LAB_Control_Client_BIC(lGOSDOSCLI As String)

fraDetail_LAB_Control_Client_BIC = ""

If Val(txtGOSDOSCLI) = 0 Then
Else

    X = "select ADRESSRA12 from " & paramIBM_Library_SAB & ".ZADRESS0  where ADRESSTYP = 4 and ADRESSNUM = '" & " " & Format(Val(txtGOSDOSCLI), "0000000") & "'"
    Set rsSab = cnsab.Execute(X)
    
    If Not rsSab.EOF Then fraDetail_LAB_Control_Client_BIC = Trim(rsSab("ADRESSRA12"))
End If



End Function

Public Sub fraMail_Confirm()
Dim X As String, K As Integer
'libMail_Subject = "Objet : Gestion des opérations en suspens " & Date & " " & Time

blnGOSEVE_Mail = True
cmdMail_MT_NOK.Visible = True

txtMail_MT_To = arrService_Mail(Val(Mid$(newYGOSDOS0.GOSDOSGSRV, 2, 2)), 1)

lstMail_MT_To.Visible = False
X = UCase$(txtMail_MT_To)
For K = 0 To lstMail_MT_To.ListCount - 1
    If InStr(X, lstMail_MT_To.List(K)) > 0 Then
        lstMail_MT_To.Selected(K) = True
    Else
        lstMail_MT_To.Selected(K) = False
    End If
Next K
Call lstMail_MT_To_TXT

If Trim(newYGOSDOS0.GOSDOSRCOM) <> "" Then
    X = Trim(arrRCOM_Lib(Val(Mid$(newYGOSDOS0.GOSDOSRCOM, 2, 2))))
    If X = "" Or X = arrService_Mail(41, 1) Then
        txtMail_MT_CC = arrService_Mail(41, 1)
    Else
        'txtMail_MT_CC = arrService_Mail(41, 1) & ";" & Trim(arrRCOM_Lib(Val(Mid$(newYGOSDOS0.GOSDOSRCOM, 2, 2))))
        '$JPL 2014-11-12
        txtMail_MT_CC = Trim(arrRCOM_Lib(Val(Mid$(newYGOSDOS0.GOSDOSRCOM, 2, 2))))
    End If
End If

If Mid$(cmdSelect_SQL_K, 1, 1) = "2" Then
 If currentSSIWINUNIT <> Trim(newYGOSDOS0.GOSDOSGSRV) Then txtMail_MT_CC = txtMail_MT_CC + ";" & arrService_Mail(Val(Mid$(newYGOSDOS0.GOSDOSISRV, 2, 2)), 1)

End If



lstMail_MT_CC.Visible = False
X = UCase$(txtMail_MT_CC)

' ________________ AJOUT KOKOU 10/02/2025  SUR DEMANDE DE MR BENMALEK __________________
If Trim(X) = "BENMALEK;" Then
    X = "DCOM;"
End If

'___________________ FIN AJOUT KOKOU ____________________________________________________


For K = 0 To lstMail_MT_CC.ListCount - 1
    If InStr(X, lstMail_MT_CC.List(K)) > 0 Then
        lstMail_MT_CC.Selected(K) = True
    Else
        lstMail_MT_CC.Selected(K) = False
    End If
Next K
Call lstMail_MT_CC_TXT

If mYGOSDOS0_Fct = "New" Then
    txtMail_MT_Message = "Nous vous remercions de bien vouloir nous faire parvenir dans les meilleurs délais les renseignements demandés ci-après :" & vbCrLf  ' _
                    '& Trim(newYGOSEVE0.GOSEVETXT)
Else
    txtMail_MT_Message = "Veuillez trouver ci-après la situation du dossier :" & vbCrLf
End If

fraMail_MT.Visible = True
End Sub

Public Sub fraEVE_Display()
fraPJ.Visible = False
txtGOSEVETXT = oldYGOSEVE0.GOSEVETXT
Call cbo_Scan(oldYGOSEVE0.GOSEVEGSRV, cboGOSEVEGSRV)
cboGOSEVEGSRV.Visible = True
cboGOSEVENAT.Clear
cboGOSEVENAT.AddItem oldYGOSEVE0.GOSEVENAT
cboGOSEVENAT.ListIndex = 0
Call DTPicker_Set(txtGOSEVEECHD, CStr(oldYGOSDOS0.GOSDOSECHD))
lblGOSEVEECHD.BackColor = mColor_G1

lblGOSEVEUAMJ = "màj  le " & dateImp10_S(oldYGOSEVE0.GOSEVEUAMJ) & " à " & timeImp8(oldYGOSEVE0.GOSEVEUHMS) & "   par " & Trim(oldYGOSEVE0.GOSEVEUUSR) & " - " & arrService_Lib(Mid$(oldYGOSEVE0.GOSEVEUSRV, 2, 2))
fraEVE.Visible = True
fraEVE.ZOrder 0

Select Case oldYGOSEVE0.GOSEVENAT
    Case "PJ**":    Call fraEVE_Display_PJ_FileName(oldYGOSEVE0.GOSEVETXT, True)
    Case "Swi+": Call fgSwift_Display(oldYGOSEVE0.GOSEVESWID)
    Case "Swi>": If oldYGOSEVE0.GOSEVESWID > 0 Then Call fgSwift_Display(oldYGOSEVE0.GOSEVESWID)
End Select

fraEVE_S.Enabled = False

cmdEVE_Reset

If oldYGOSDOS0.GOSDOSSTAD <> "C" And oldYGOSDOS0.GOSDOSSTAD <> "A" Then
    If currentSSIWINUNIT = oldYGOSDOS0.GOSDOSISRV Or currentSSIWINUNIT = oldYGOSDOS0.GOSDOSGSRV Then
        If oldYGOSEVE0.GOSEVENAT = "Swi>" Then cmdEVE_Dupliquer.Visible = arrHab(2)
    End If
End If

If oldYGOSEVE0.GOSEVESTAE = " " _
And arrHab(3) _
And oldYGOSEVE0.GOSEVEUSRV = currentSSIWINUNIT Then

    Select Case oldYGOSEVE0.GOSEVENAT
        Case "Swi>", "Swi+", "Note", "PJ**"
            fraEVE_S.Enabled = True
            cmdEVE_Ignore.Visible = True
            
   End Select
End If
If blnSwift_Display And oldYGOSEVE0.GOSEVESWID <> 0 Then Call fgSwift_Display(oldYGOSEVE0.GOSEVESWID)
fraEVE.BackColor = fraDetail.BackColor
End Sub

Public Function fraEVE_Display_PJ_FileName(lTxt As String, blnDisplay As Boolean) As String
Dim wExtension As String, X As String

wExtension = UCase$(fileName_Extension(Trim(lTxt)))
X = paramGOSDOS_Path_DROPI & oldYGOSEVE0.GOSEVEIDD _
    & "\" & oldYGOSEVE0.GOSEVEIDD & "_" & oldYGOSEVE0.GOSEVEIDE _
    & "." & wExtension
If blnDisplay Then Call frmElpPrt.Windows_Display_File(X)
    'If Dir(X) <> "" Then
    '    Select Case wExtension
    '     Case "DOC": Call frmElpPrt.WinWord(X)
    '     Case "XLS": Call frmElpPrt.Excel(X)
    '     Case "PDF": Call frmElpPrt.Acrord32(X)
    '     Case "TXT": Call frmElpPrt.WordPad(X) 'NotePad(X)
    '     Case "RTF": Call frmElpPrt.WordPad(X)
    '     Case Else: Call frmElpPrt.IExplore(X)
    '    End Select
    'End If
'End If
fraEVE_Display_PJ_FileName = X
End Function
Public Sub fraMail_Print()
blnYGOSDOS0_Update = False
'libMail_Subject = "Objet : Gestion des opérations en suspens " & dateImp10_S(DSys) & " " & time_Hms
txtMail_MT_To = currentSSIWINMAIL
txtMail_MT_CC = ""
txtMail_MT_Message = ""

fraMail_MT.Visible = True

End Sub




Public Sub fraPJ_Init()

txtGOSEVETXT = ""
txtGOSEVETXT.Locked = True 'False
rtfPJ.Top = filDoc.Top + filDoc.Height + 200 ' 3400
rtfPJ.Left = filDoc.Left
rtfPJ.Width = filDoc.Width
rtfPJ.Height = filDoc.Height '2025
fraPJ.Visible = True
filDoc.Pattern = "_.*"
filDoc.Pattern = "*.*"

oldFileName = "": newFileName = ""
rtfPJ.Text = ""

End Sub

Public Sub cboGOSEVENAT_Display()
Dim X As String
fraPJ.Visible = False
txtGOSEVETXT.Locked = False

If fraEVE_S.Enabled Then
    Select Case Mid$(cboGOSEVENAT, 1, 4)
        Case "Note":
                cmdEVE_Ok.Visible = False
                Call lstParam_GOSDOSLABK_Load("NOTE", "fgModèle")
                fgModèle.ZOrder 0
                If fgModèle.Rows > 2 Then
                    fgModèle.Visible = True
                Else
                    cmdEVE_Ok.Visible = arrHab(2)
                End If
         Case "SwFR", "SwGB":
                If IsNumeric(Mid$(oldYGOSDOS0.GOSDOSWMTK, 1, 1)) Then
                    cboEVE_Swift_MT = Mid$(oldYGOSDOS0.GOSDOSWMTK, 1, 1) & "99"
                Else
                    cboEVE_Swift_MT = "999"
                End If
                
                cmdEVE_Ok.Visible = False
                cboEVE_Swift_BIC.Clear
                cboEVE_Swift_BIC.AddItem oldYGOSDOS0.GOSDOSWBIC
                cboEVE_Swift_BIC.ListIndex = 0
                
                cboEVE_Swift_20.Clear
                
                cboEVE_Swift_20.AddItem arrService_Code_SAA(Mid$(currentSSIWINUNIT, 2, 2)) & "GOS" & Format$(oldYGOSDOS0.GOSDOSIDD, "00000")
                cboEVE_Swift_20.ListIndex = 0
                If oldYSWISAB0.SWISABOPEN <> 0 Then
                    X = oldYSWISAB0.SWISABSER & oldYSWISAB0.SWISABOPEC & Format(oldYSWISAB0.SWISABOPEN, "000000000")
                    X = SAA_from_SAB_TRN(X, "")
                    cboEVE_Swift_20.AddItem X
                    cboEVE_Swift_20.BackColor = mColor_W1
                    '$JPL 2014-10-10 If currentSSIWINUNIT <> "S01" Then cboEVE_Swift_20.ListIndex = 0
                Else
                    cboEVE_Swift_20.BackColor = mColor_G0
                End If
                
                txtEVE_Swift_21 = Trim(oldYGOSDOS0.GOSDOSWTRN)
                
                If txtEVE_Swift_21 = "" And Mid$(oldYGOSDOS0.GOSDOSWMTK, 1, 1) = "3" Then
                
                    X = "select SWIRAMXREF from  " & paramIBM_Library_SABSPE & ".YSWIRAM0 " _
                         & " WHERE SWIRAMXID = " & oldYSWISAB0.SWISABSWID
     
                    Set rsSabX = cnsab.Execute(X)

                    If Not rsSabX.EOF Then txtEVE_Swift_21 = rsSabX("SWIRAMXREF")

                End If
                fraEVE_Swift.Visible = True
                fraEVE_Swift.ZOrder 0
         Case "Mail": cmdEVE_Ok_Click: Exit Sub
         Case "PJ**": fraPJ_Init: cmdEVE_Ok.Visible = arrHab(2)
    End Select
Else
'    fraPJ.Visible = False
'    txtGOSEVETXT.Locked = False
End If

End Sub


Public Function cmdYGOSEVE0_Update_PJ()
Dim Archive_Folder As String, Archive_File As String
Dim App_Event As String
On Error GoTo Error_Handler

cmdYGOSEVE0_Update_PJ = Null

newFileName = newDirPath & "\" & newYGOSEVE0.GOSEVEIDD & "_" & newYGOSEVE0.GOSEVEIDE _
            & "." & newFileExtension
    
App_Event = "MkDir " & newDirPath

If Not msFileSystem.FolderExists(newDirPath) Then MkDir newDirPath
App_Event = "CopyFile " & oldFileName & vbCrLf & newFileName
If Dir(newFileName) <> "" Then Kill newFileName

msFileSystem.CopyFile oldFileName, newFileName

If Mid$(oldFileName, 1, 8) = "C:\Temp\" Then
    Archive_Folder = "C:\Temp\Archive"
    App_Event = "MkDir " & Archive_Folder
   If Not msFileSystem.FolderExists(Archive_Folder) Then MkDir Archive_Folder
    X = Mid$(oldFileName, 8, Len(oldFileName) - 7)
    X = Replace(X, "\", "_")
    Archive_File = Archive_Folder & "\" & newYGOSEVE0.GOSEVEIDD & "_" & DSys & "_" & time_Hms & "_" & X
    App_Event = "MoveFile " & oldFileName & vbCrLf & Archive_File
   msFileSystem.MoveFile oldFileName, Archive_File
End If

Exit Function

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V & vbCrLf & App_Event, vbCritical, frmElp_Caption & "cmdYGOSEVE0_Update_PJ"
    cmdYGOSEVE0_Update_PJ = V
Exit_sub:

End Function

Public Sub cmdSendMail_fgSelect(lRecipient As String)
Dim wSendMail As typeSendMail
Dim xDétail As String, xHeader As String, mbgColor As String
Dim K As Long, htmlFontColor_K As String
Dim iRow As Integer, iCol As Integer, X As String, xTD As String
Dim wForecolor As String, wBackColor As String, xColor As String
Dim wCols As Integer, xDetail As String
Dim xWidth As String
On Error Resume Next


wCols = fgSelect.Cols - 2
xHeader = ""
Select Case cmdSelect_SQL_K
    Case "3", "3x"
        wCols = 11
        wSendMail.Subject = "BIA_GOS : Liste des dossiers en suspens (édité le " & Date & " " & Time & ")"
        xHeader = xHeader & htmlFontColor_Gray & "<BR> &#149;&#160;&#160; Code état &#160;&#160;&#160;&#160; : " & htmlFontColor_Blue & cboSelect_3_GOSDOSSTAG & cboSelect_3_GOSDOSSTAD
        If Trim(cboSelect_3_GOSDOSGSRV) <> "" Then xHeader = xHeader & htmlFontColor_Gray & "<BR> &#149;&#160;&#160; Service &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; : " & htmlFontColor_Blue & cboSelect_3_GOSDOSGSRV
        If Trim(cboSelect_3_GOSDOSRCOM) <> "" Then xHeader = xHeader & htmlFontColor_Gray & "<BR> &#149;&#160;&#160; Responsable : " & htmlFontColor_Blue & cboSelect_3_GOSDOSRCOM
        If Trim(cboSelect_3_GOSDOSCLI) <> "" Then xHeader = xHeader & htmlFontColor_Gray & "<BR> &#149;&#160;&#160; Client &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; : " & htmlFontColor_Blue & cboSelect_3_GOSDOSCLI
        If Trim(cboSelect_3_GOSDOSWBIC) <> "" Then xHeader = xHeader & htmlFontColor_Gray & "<BR> &#149;&#160;&#160; BIC &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; : " & htmlFontColor_Blue & cboSelect_3_GOSDOSWBIC
        If chkSelect_GOSDOSIAMJ = "1" Then xHeader = xHeader & htmlFontColor_Gray & "<BR> &#149;&#160;&#160; Créés du &#160;&#160;&#160;&#160;&#160;&#160; : " & htmlFontColor_Blue & Mid$(txtSelect_GOSDOSIAMJ_Min, 1, 10) & htmlFontColor_Black & " au " & htmlFontColor_Blue & Mid$(txtSelect_GOSDOSIAMJ_Max, 1, 10)
    
    Case "4"
        wCols = 11
        wSendMail.Subject = "BIA_GOS : Echéancier des dossiers en gestion - (édité le " & Date & " " & Time & ")"
        If Trim(cboSelect_4_GOSDOSISRV) <> "" Then
            xHeader = xHeader & htmlFontColor_Red & "<BR> &#149;&#160;&#160;" & Trim(lblSelect_4_GOSDOSISRV) & "&#160;&#160;&#160;&#160;&#160;&#160; : " & htmlFontColor_Blue & cboSelect_4_GOSDOSISRV
            wSendMail.Subject = "BIA_GOS : Echéancier des dossiers initiés par le service " & cboSelect_4_GOSDOSISRV & " - (édité le " & Date & " " & Time & ")"
        End If
        If Trim(cboSelect_4_GOSDOSGSRV) <> "" Then xHeader = xHeader & htmlFontColor_Gray & "<BR> &#149;&#160;&#160;" & Trim(lblSelect_4_GOSDOSGSRV) & "&#160; : " & htmlFontColor_Blue & cboSelect_4_GOSDOSGSRV
        xHeader = xHeader & htmlFontColor_Gray & "<BR> &#149;&#160;&#160;" & Trim(lblSelect_4_GOSDOSECHD) & " : " & htmlFontColor_Blue & Mid$(txtSelect_4_GOSDOSECHD, 1, 10)
    Case "4 Journal"
        wCols = 5
        wSendMail.Subject = "BIA_GOS : Journal des évènements du " & Mid$(txtSelect_4_GOSDOSECHD, 1, 10) & " - (édité le " & Date & " " & Time & ")"
        If Trim(cboSelect_4_GOSDOSISRV) <> "" Then xHeader = xHeader & htmlFontColor_Gray & "<BR> &#149;&#160;&#160;" & Trim(lblSelect_4_GOSDOSISRV) & "&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; : " & htmlFontColor_Blue & cboSelect_4_GOSDOSISRV
        If Trim(cboSelect_4_GOSDOSGSRV) <> "" Then xHeader = xHeader & htmlFontColor_Red & "<BR> &#149;&#160;&#160;" & Trim(lblSelect_4_GOSDOSGSRV) & " : " & htmlFontColor_Blue & cboSelect_4_GOSDOSGSRV
        xHeader = xHeader & htmlFontColor_Gray & "<BR>  &#149;&#160;&#160;" & Trim(lblSelect_4_GOSDOSECHD) & "&#160;&#160;&#160;&#160;&#160;&#160; : " & htmlFontColor_Blue & Mid$(txtSelect_4_GOSDOSECHD, 1, 10)
    Case "4 Swi>"
        wCols = 5
        wSendMail.Subject = "BIA_GOS : Swift créé via GOS non émis SAA " & " - (édité le " & Date & " " & Time & ")"
        If Trim(cboSelect_4_GOSDOSGSRV) <> "" Then xHeader = xHeader & htmlFontColor_Red & "<BR> &#149;&#160;&#160;" & Trim(lblSelect_4_GOSDOSGSRV) & " : " & htmlFontColor_Blue & cboSelect_4_GOSDOSGSRV
        xHeader = xHeader & htmlFontColor_Gray
    Case Else
         wSendMail.Subject = "BIA_GOS : Liste des dossiers en suspens (édité le " & Date & " " & Time & ")"
   
End Select


xDetail = ""
For iRow = 0 To fgSelect.Rows - 1
    fgSelect.Row = iRow
    xTD = ""
    For iCol = 0 To wCols
            fgSelect.Col = iCol
            X = Trim(fgSelect.Text)
            If iCol = 3 Then X = "<div align=" & Asc34 & "right" & Asc34 & ">" & X
            If iCol <> wCols Then
                xWidth = " width=270px NOWRAP"
            Else
                xWidth = " width=270px"
           End If
            
            If iRow = 0 Then
                wForecolor = RGB_Html_Color(fgSelect.ForeColorFixed)
                wBackColor = RGB_Html_Color(fgSelect.BackColorFixed)
                xTD = xTD _
                     & "<TD bgcolor=" & wBackColor & xWidth & "><span style='font-size:10.0pt;font-family:Calibri'><Font color=" & wForecolor & "><B>" _
                     & X & "</B/TD>"
            Else
                wForecolor = RGB_Html_Color(fgSelect.CellForeColor)
                If fgSelect.CellBackColor = 0 Then
                    wBackColor = RGB_Html_Color(fgSelect.BackColor)
                Else
                     wBackColor = RGB_Html_Color(fgSelect.CellBackColor)
               End If
                If fgSelect.CellFontBold Then X = "<B>" & X & "</B>"
                xTD = xTD _
                     & "<TD bgcolor=" & wBackColor & xWidth & "><span style='font-size:10.0pt;font-family:Calibri'><Font color=" & wForecolor & ">" _
                     & X & "</TD>"
            End If
    Next iCol
    xDetail = xDetail & "<TR>" & xTD & "</TR>"

Next iRow

mbgColor = "bgcolor = #E0E0E0"

wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = "BIA_GOS"
wSendMail.Recipient = lRecipient
wSendMail.Attachment = ""
  
wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                    & "<span style='font-size:12.0pt;font-family:Calibri'>" & "<Font color = #404040>" _
                    & "<B><U>" & htmlFontColor_Gray & wSendMail.Subject & "</B></U><BR>" _
                    & xHeader & "<BR><BR>" _
                    & "<TABLE   width=2430px border=1 cellpadding=5 ></B>" _
                    & xDetail _
                    & "</div></TABLE>"

'                    & "<div align=" & Asc34 & "left" & Asc34 _

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail


End Sub
Public Sub cmdSendMail_Echéancier()
Dim xDétail As String
Dim K As Long, htmlFontColor_K As String
Dim X As String, xTD As String
Dim wForecolor As String, wBackColor As String, xColor As String
Dim blnHeader As Boolean, iCol As Integer
On Error Resume Next

If arrYGOSDOS0_Nb > 1 Then
    arrYGOSDOS0(0) = arrYGOSDOS0(1)
    blnHeader = False
    xDétail = ""
    For K = 1 To arrYGOSDOS0_Nb
    
        If arrYGOSDOS0(0).GOSDOSGSRV <> arrYGOSDOS0(K).GOSDOSGSRV Then
            Call cmdSendMail_Echéancier_Ok(arrYGOSDOS0(0).GOSDOSGSRV, xDétail)
            arrYGOSDOS0(0) = arrYGOSDOS0(K)
            blnHeader = False
            xDétail = ""
        End If

        If Not blnHeader Then
            blnHeader = True
            fgSelect.Row = 0
            xTD = ""
            For iCol = 0 To fgSelect.Cols - 2
                fgSelect.Col = iCol
                X = Trim(fgSelect.Text)
                wForecolor = RGB_Html_Color(fgSelect.ForeColorFixed)
                wBackColor = RGB_Html_Color(fgSelect.BackColorFixed)
                xTD = xTD _
                     & "<TD bgcolor=" & wBackColor & " width=270px><span style='font-size:9.0pt;font-family:Calibri'><Font color=" & wForecolor & "><B>" _
                     & X & "</B/TD>"
            Next iCol
            
            xDétail = xDétail & "<TR>" & xTD & "</TR>"
       End If
       
       xYGOSDOS0 = arrYGOSDOS0(K)
        xDétail = xDétail & "<TR>" _
                 & "<TD width=270px><span style='font-size:9.0pt;font-family:Calibri'>" & htmlFontColor_Blue & xYGOSDOS0.GOSDOSWMTK & " " & xYGOSDOS0.GOSDOSWES & "</TD>" _
                 & "<TD width=270px><span style='font-size:9.0pt;font-family:Calibri'>" & htmlFontColor_Blue & xYGOSDOS0.GOSDOSWBIC & "</TD>" _
                 & "<TD width=270px><span style='font-size:9.0pt;font-family:Calibri'>" & htmlFontColor_Blue & xYGOSDOS0.GOSDOSWTRN & "</TD>" _
                 & "<TD width=270px><span style='font-size:9.0pt;font-family:Calibri'>" & htmlFontColor_Red & Format$(xYGOSDOS0.GOSDOSWMTD, "### ### ### ##0.00") & "</TD>" _
                 & "<TD width=270px><span style='font-size:9.0pt;font-family:Calibri'>" & htmlFontColor_Blue & xYGOSDOS0.GOSDOSWDEV & "</TD>" _
                 & "<TD width=270px><span style='font-size:9.0pt;font-family:Calibri'>" & htmlFontColor_Blue & xYGOSDOS0.GOSDOSSTAD & "</TD>" _
                 & "<TD width=270px><span style='font-size:9.0pt;font-family:Calibri'>" & htmlFontColor_Blue & dateImp10_S(xYGOSDOS0.GOSDOSECHD) & "</TD>" _
                 & "<TD width=270px><span style='font-size:9.0pt;font-family:Calibri'>" & htmlFontColor_Red & arrService_Lib(Val(Mid$(xYGOSDOS0.GOSDOSGSRV, 2, 2))) & "</TD>" _
                 & "<TD width=270px><span style='font-size:9.0pt;font-family:Calibri'>" & htmlFontColor_Blue & arrRCOM_Lib(Val(Mid$(xYGOSDOS0.GOSDOSRCOM, 2, 2))) & "</TD>" _
                 & "<TD width=270px><span style='font-size:9.0pt;font-family:Calibri'>" & htmlFontColor_Blue & xYGOSDOS0.GOSDOSCLI & "</TD>" _
                 & "<TD width=270px><span style='font-size:9.0pt;font-family:Calibri'>" & htmlFontColor_Blue & xYGOSDOS0.GOSDOSLABK & "</TD>" _
                & "</TR>"
    
    
    Next K
    
    Call cmdSendMail_Echéancier_Ok(arrYGOSDOS0(0).GOSDOSGSRV, xDétail)

End If

End Sub

Public Sub cmdSendMail_Echéancier_Ok(lGOSDOSGSRV As String, lMessage As String)
Dim wSendMail As typeSendMail
Dim X As String, Xto As String, mbgColor As String
Dim xError As String
On Error Resume Next

wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = "BIA_GOS"
'wSendMail.RecipientDisplayName = "BIA_GOS"
X = arrService_Mail(Val(Mid$(lGOSDOSGSRV, 2, 2)), 1)

V = mailAdresse_Production_Control(X, Xto)
If IsNull(V) Then
    wSendMail.Recipient = Xto
    mbgColor = "<body bgcolor = #FFFFFF>"
    xError = ""
Else
    wSendMail.Recipient = "informatique@bia-paris.fr"
    mbgColor = "<body bgcolor = #FFB0FF>"
    xError = "<B>ERREUR DE DECODAGE ADRESSE MAIL : " & htmlFontColor_Red & X & "</B><BR><BR>"
End If


wSendMail.Subject = "BIA_GOS : Echéancier des dossiers en suspens au " & Date & " " & Time
wSendMail.Attachment = ""
  
wSendMail.Message = mbgColor _
                    & "<span style='font-size:10.0pt;font-family:Calibri'>" & "<Font color = #404040>" _
                    & "<U>" & wSendMail.Subject & "</U><BR><BR>" & xError _
                    & "<TABLE   width=2430px border=1 cellpadding=10 ></B>" _
                    & "<div align=" & Asc34 & "right" & Asc34 _
                    & lMessage _
                    & "</div></TABLE>"

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail


End Sub

Public Sub lstParam_GOSDOSLABK_Load(lNAT As String, lControl As String)
Dim xSql As String, X As String

lstParam_GOSDOSLABK.Visible = False
fraParam_GOSDOSLABK.Visible = False

Call rsYGOSEVE0_Init_Param(oldParam)
oldParam.GOSEVENAT = lNAT
Select Case lControl
    Case "cboGOSDOSLABK": cboGOSDOSLABK.Clear
    Case "cboParam_StaC": cboParam_StaC.Clear
    Case "fgModèle": fgModèle.Rows = 2
End Select

lstParam_GOSDOSLABK.Clear
lstParam_GOSDOSLABK.AddItem "Ajouter un enregistrement"


xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where GOSEVEIDD = -1 and GOSEVEGSRV = '" & currentSSIWINUNIT & "'" _
     & " and GOSEVENAT = '" & lNAT & "' order by SUBSTRING(GOSEVETXT , 1 , 10)"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    X = Trim(rsSab("GOSEVETXT"))
    Select Case lNAT
        Case "Mail", "StaC", "StaP", "DoSw"
            lstParam_GOSDOSLABK.AddItem Mid$(X, 1, 10) & "      " & Mid$(X, 31, 30)
        Case Else
            lstParam_GOSDOSLABK.AddItem Mid$(X, 1, 10) & "      " & Mid$(X, 11, 20)
    End Select
    
    Select Case lControl
        Case "cboGOSDOSLABK": cboGOSDOSLABK.AddItem Trim(Mid$(X, 1, 10))
        Case "cboParam_StaC": cboParam_StaC.AddItem Mid$(X, 1, 3) & "- " & Trim(Mid$(X, 31, 30))
        Case "fgModèle":
            fgModèle.Rows = fgModèle.Rows + 1
            fgModèle.Row = fgModèle.Rows - 1
            fgModèle.Col = 0: fgModèle = Trim(Mid$(X, 1, 10))
            fgModèle.Col = 1: fgModèle = Mid$(X, 12, 7)
            fgModèle.Col = 2: fgModèle = Trim(Mid$(X, 31, Len(X) - 30))
    End Select
    rsSab.MoveNext
Loop
lstParam_GOSDOSLABK.Visible = True

End Sub

Private Sub lstParam_SAA_Id_Click()
Dim X As String
Dim xSql As String, K As Integer


Old_YBIATAB0.BIATABID = "SAA"
Old_YBIATAB0.BIATABK1 = Trim(lstParam_SAA_Id.Text)

lstParam_SAA_Load

If lstParam_SAA_K1.ListCount = 1 Then lstParam_SAA_K1_Click
End Sub
Private Sub lstParam_SAA_K1_Click()
Dim xSql As String, X As String

'_______________________________________________
fraParam_SAA.Visible = False
'_______________________________________________
txtParam_SAA_K2.Enabled = False
cmdParam_SAA_Add.Visible = False
cmdParam_SAA_Delete.Visible = False
cmdParam_SAA_Update.Visible = False
cmdParam_SAA_Add.Visible = False
txtParam_SAA_MTD.Visible = False: lblParam_SAA_MTD.Visible = False
fraParam_SAA_Jrnl_Event.Visible = False

Old_YBIATAB0.BIATABK2 = Trim(lstParam_SAA_K1.Text)

xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'SAA'" _
     & " and BIATABK1 = '" & Old_YBIATAB0.BIATABK1 & "' and BIATABK2 = '" & Old_YBIATAB0.BIATABK2 & "'"

Set rsSab = cnsab.Execute(xSql)


If Not rsSab.EOF Then
    Old_YBIATAB0.BIATABTXT = rsSab("BIATABTXT")
    Old_YBIATAB0.BIATABK2 = rsSab("BIATABK2")
Else
    Old_YBIATAB0.BIATABTXT = ""
    Old_YBIATAB0.BIATABK2 = ""
End If

optParam_SAA_TXT_RMA_X = True
Select Case Trim(Old_YBIATAB0.BIATABK1)
    Case "Amount", "Approval"
        txtParam_SAA_K2 = Trim(Old_YBIATAB0.BIATABK2)
        txtParam_SAA_MTD = Val(Old_YBIATAB0.BIATABTXT)
        txtParam_SAA_MTD.Visible = True: lblParam_SAA_MTD.Visible = True
        txtParam_SAA_K2.Enabled = arrHab(16)
        cmdParam_SAA_Add.Visible = arrHab(16)
        cmdParam_SAA_Delete.Visible = arrHab(16)
        cmdParam_SAA_Update.Visible = arrHab(16)

    Case "Jrnl_Event"
        txtParam_SAA_K2 = Old_YBIATAB0.BIATABK2
        txtParam_SAA_TXT = Mid$(Old_YBIATAB0.BIATABTXT, 1, 99)
        txtParam_SAA_K2.Enabled = arrHab(19)
        cmdParam_SAA_Add.Visible = arrHab(19)
        cmdParam_SAA_Delete.Visible = arrHab(19)
        fraParam_SAA_Jrnl_Event.Visible = True
        cboParam_SAA_TopK.Visible = True: lblParam_SAA_TopK.Visible = True: lblParam_SAA_TopK_2.Visible = True
        cboParam_SAA_TopK = Mid$(Old_YBIATAB0.BIATABTXT, 103, 1)
        cboParam_SAA_Alerte.Visible = True: lblParam_SAA_Alerte.Visible = True
        cboParam_SAA_Alerte = Mid$(Old_YBIATAB0.BIATABTXT, 100, 3)
        Select Case Mid$(Old_YBIATAB0.BIATABTXT, 104, 1)
            Case "A": optParam_SAA_TXT_RMA_A = True
            Case "R": optParam_SAA_TXT_RMA_R = True
        End Select
        cmdParam_SAA_Update.Visible = arrHab(16)
    Case "Mesg_Type", "Mesg_Fields"
        txtParam_SAA_K2 = Trim(Old_YBIATAB0.BIATABK2)
        txtParam_SAA_TXT = Trim(Old_YBIATAB0.BIATABTXT)
        fraParam_SAA_Jrnl_Event.Visible = True
        cboParam_SAA_TopK.Visible = False: lblParam_SAA_TopK.Visible = False: lblParam_SAA_TopK_2.Visible = False
        cboParam_SAA_Alerte.Visible = False: lblParam_SAA_Alerte.Visible = False
        txtParam_SAA_K2.Enabled = arrHab(19)
        cmdParam_SAA_Add.Visible = arrHab(19)
        cmdParam_SAA_Delete.Visible = arrHab(19)
        cmdParam_SAA_Update.Visible = arrHab(19)
    Case Else
        txtParam_SAA_K2 = Old_YBIATAB0.BIATABTXT
        txtParam_SAA_K2.Enabled = arrHab(19)
        cmdParam_SAA_Add.Visible = arrHab(19)
        cmdParam_SAA_Delete.Visible = arrHab(19)
        cmdParam_SAA_Update.Visible = arrHab(19)
End Select
    
fraParam_SAA.Visible = True
End Sub


Private Sub cmdParam_SAA_Delete_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

X = Trim(txtParam_SAA_K2)
If X <> Trim(Old_YBIATAB0.BIATABK2) Then
    Call MsgBox("Le code  a été modifié," & vbCrLf & " la suppression n'est pas possible", vbCritical, "BIA_GOS : paramétrage SAA")
Else
    
    If IsNull(Parametrage_Delete) Then lstParam_SAA_Load
End If


Me.Enabled = True: Me.MousePointer = 0

End Sub
Public Sub lstParam_YSSIMEL0_Load()
Dim xSql As String, K As Integer, X As String

'fraParam_GOSDOSMAIL.Visible = False
xSql = "select *from " & paramIBM_Library_SABSPE & ".YSSIMEL0 where SSIMELNAT = '@' and SSIMELUIDX like 'BIA_GOS.%'"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    X = Trim(rsSab("SSIMELUIDX"))
    K = InStr(X, ".S")
    K = Val(Mid$(X, K + 2, 2))
    X = Trim(rsSab("SSIMELINFO"))
    '''X = Replace(X, "@BIA-PARIS.FR", "")
    arrService_Mail(K, 1) = X
    rsSab.MoveNext
Loop

xSql = "select *from " & paramIBM_Library_SABSPE & ".YSSIMEL0 where SSIMELNAT = '@' and SSIMELUIDX like 'SAA_Alerte.%'"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    X = Trim(rsSab("SSIMELUIDX"))
    K = InStr(X, ".S")
    K = Val(Mid$(X, K + 2, 2))
    X = Trim(rsSab("SSIMELINFO"))
    '''X = Replace(X, "@BIA-PARIS.FR", "")
    arrService_Mail(K, 2) = X
    rsSab.MoveNext
Loop

'xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'GOSDOSMAIL'"
'Set rsSab = cnsab.Execute(xSql)
'Do While Not rsSab.EOF
'    K = Val(Mid$(rsSab("BIATABK1"), 2, 2))
'    arrService_Mail(K, 1) = Trim(rsSab("BIATABTXT"))
'    rsSab.MoveNext
'Loop


'xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'SAA_Alerte' and BIATABK1 = 'Mail'"
'Set rsSab = cnsab.Execute(xSql)
'Do While Not rsSab.EOF
'    K = Val(Mid$(rsSab("BIATABK2"), 2, 2))
'    arrService_Mail(K, 2) = Trim(rsSab("BIATABTXT"))
'    rsSab.MoveNext
'Loop

End Sub

Public Function mailAdresse_Recipient_Unit(lUnit As String, lMail_K As Integer) As String
Dim V, K As Integer, X As String, Xto As String
Dim wUnit_List As String, kInstr As Integer, xUnit As String, xDest As String
Xto = ""
wUnit_List = Trim(lUnit) & ";"
'If Mid$(wUnit, Len(wUnit), 1) <> ";" Then wUnit = wUnit & ";"
kInstr = 1
'_________________________________________________________________________________

Do
K = InStr(kInstr, wUnit_List, ";")
If K > 0 Then
    xUnit = Mid$(wUnit_List, kInstr, K - kInstr)
    kInstr = K + 1
    xDest = ""
    
    If xUnit = "ORPA" Or xUnit = "GDMP" Or xUnit = "SOBF" Then xUnit = "S01"

    If IsNumeric(Mid$(xUnit, 2, 2)) Then
        V = mailAdresse_Production_Control(arrService_Mail(Mid$(xUnit, 2, 2), lMail_K), xDest)
    Else
        For K = 1 To 99
            If arrService_Code_SAA(K) = xUnit Then
                X = arrService_Mail(K, lMail_K)
                V = mailAdresse_Production_Control(X, xDest)
                Exit For
            End If
        Next K
        If xDest = "" Then Call mailAdresse_Production_Control_UTI(xUnit, xDest)
    End If
    
    If xDest <> "" Then
        If Xto = "" Then
            Xto = xDest
        Else
            Xto = Xto & ";" & xDest
        End If
        
    End If
End If

Loop Until K = 0
'_________________________________________________________________________________
mailAdresse_Recipient_Unit = Xto

End Function



Public Sub lstParam_YSSIMEL0_USR_Load()
Dim X As String

lstMail_MT_To.Clear
lstMail_MT_CC.Clear

X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
     & " where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN' and SSIDOMUNIT <> '' and SSIDOMPRFK <> 'X'" _
     & " order by SSIDOMUIDX"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    X = Trim(rsSab("SSIDOMUIDX"))
    'lstParam_GOSDOSMAIL_Usr.AddItem X
    lstMail_MT_To.AddItem X
    lstMail_MT_CC.AddItem X
    rsSab.MoveNext
Loop
lstMail_MT_To.AddItem ".S01 GDMP"
lstMail_MT_To.AddItem ".S10 CREDOC"
lstMail_MT_To.AddItem ".S32 GDC-BOTC"
lstMail_MT_To.AddItem ".S42 SecFin"

' _________________AJOUT KOKOU SUR 10/02/2025 SUR DEMANDE DE MR BENMALEK ____________________________

lstMail_MT_To.AddItem "DCOM"
lstMail_MT_CC.AddItem "DCOM"

' _______________ FIN AJOUT KOKOU ___________________________________________________________________
End Sub

Public Sub fraList_Display()
Dim xSql As String

fgEVE.Visible = False
fraEVE.Visible = False

cmdList_SAB_Annulation.Visible = False
cmdList_SAB_Modification.Visible = False
cmdList_Ignore.Visible = False
cmdList_Add.Visible = False
cmdList_Display.Visible = True: txtList_Add = ""
cmdList_New.Visible = False
cmdList_SWISABKSRV.Visible = False
cboList_SWISABKSRV.Enabled = True

mList_YGOSDOS0 = oldYGOSDOS0

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABSWID = " & oldYGOSEVE0.GOSEVESWID
Set rsSab = cnsab.Execute(xSql)
If rsSab.EOF Then
    Call MsgBox("Erreur de lecture YSWISAB0.SWISABSWID : " & oldYGOSEVE0.GOSEVESWID, vbCritical, "BIA_GOS :cmdList_Ignore_Click")
    GoTo Exit_sub
End If

Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)

If oldYSWISAB0.SWISABXGOS = "G" Or oldYSWISAB0.SWISABXEVE = "G" Then
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWILNK0 where SWILNKSWID = " & oldYGOSEVE0.GOSEVESWID & " and SWILNKAPPC = 'GOS' and SWILNKSTA = ''"
    Set rsSab = cnsab.Execute(xSql)
    If Not rsSab.EOF Then txtList_Add = rsSab("SWILNKAPPN")

End If

tabDetail.Tab = 0
tabDetail.Caption = "SWIFT id = " & oldYSWISAB0.SWISABSWID & "    du " & dateImp10(oldYSWISAB0.SWISABWAMJ) & "  " & timeImp(oldYSWISAB0.SWISABWHMS)

Call arrYGOSEVE0_SQL(" where GOSEVESWID = " & oldYGOSEVE0.GOSEVESWID & " and GOSEVEIDD = 0")
If arrYGOSEVE0_Nb > 0 Then

    fgEVE_Display
    tabDetail.Tab = 1
    tabDetail.Caption = "Evénements dossier n° " & oldYGOSEVE0.GOSEVESWID
Else
    tabDetail.Tab = 1
    tabDetail.Caption = ""

End If
tabDetail.Tab = 0

 
Call cbo_Scan(oldYSWISAB0.SWISABKSRV, cboList_SWISABKSRV)

If oldYSWISAB0.SWISABKSRV = "S00" Then
    cmdList_SWISABKSRV.BackColor = vbYellow
Else
    cmdList_SWISABKSRV.BackColor = RGB(200, 200, 200)
End If

cmdList_Quit.Visible = True

If oldYSWISAB0.SWISABKSRV = "S00" Then
    cmdList_SWISABKSRV.Visible = arrHab(13)
    cboList_SWISABKSRV.Enabled = arrHab(13)
End If
'Else

    Select Case cmdSelect_SQL_K
        Case "5"
            'If oldYSWISAB0.SWISABK20 = "!" Or oldYSWISAB0.SWISABKPDE = "!" Or oldYSWISAB0.SWISABKPDE = "!" Then
                xSql = " where SWISABWL20 = '" & oldYSWISAB0.SWISABWL20 & "'" _
                     & " and SWISABWES = 'E' and SWISABWBIC = '" & oldYSWISAB0.SWISABWBIC & "' And SWISABWMTK = " & oldYSWISAB0.SWISABWMTK
                fglist_Display_YSWISAB0 xSql
                cmdList_SAB_Annulation.Visible = arrHab(11)
                cmdList_SAB_Modification.Visible = arrHab(11)
                cmdList_SWISABKSRV.Visible = arrHab(11)
                fraList_Options.Visible = arrHab(11)
                cmdList_Ignore.Visible = arrHab(13)
            'End If
        Case "5h"
                xSql = " where SWISABWL20 = '" & oldYSWISAB0.SWISABWL20 & "'" _
                     & " and SWISABWES = 'E' and SWISABWBIC = '" & oldYSWISAB0.SWISABWBIC & "' And SWISABWMTK = " & oldYSWISAB0.SWISABWMTK
                fglist_Display_YSWISAB0 xSql
                fraList_Options.Visible = False
    
        Case Else
            fraList_Display_Habilitation
            fraList_Options.Visible = True 'arrHab(13)
    
            Call arrYGOSDOS0_SQL("where GOSDOSSTAD = ' ' order by GOSDOSIDD")
            fglist_Display_YGOSDOS0
            
        'End If
    End Select
 'End If
 
 
 
If Not arrHab(19) Then
    If currentSSIWINUNIT = oldYSWISAB0.SWISABKSRV Or oldYSWISAB0.SWISABKSRV = "S00" Then
        
    Else
        cmdList_SAB_Annulation.Visible = False
        cmdList_SAB_Modification.Visible = False
        cmdList_Ignore.Visible = False
        cmdList_Add.Visible = False
        cmdList_New.Visible = False
        cmdList_SWISABKSRV.Visible = arrHab(11) 'False
        cboList_SWISABKSRV.Enabled = arrHab(11) 'False
    End If

End If
cmdList_Display.Visible = True

fraDetail_LAB.Visible = True

Exit_sub:
fraList.Visible = True

End Sub


Public Function cmdSendMail_SWISABSWID(lSWISABSWID As Long, lX0 As String) As String
Dim xSql As String
Dim wSWISABWID1 As Long, wSWISABWIDL As Long, wSWISABWIDH As Long

lX0 = ""

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABSWID = " & lSWISABSWID
Set rsSab = cnsab.Execute(xSql)

If rsSab("SWISABWES") = "S" Then
    lX0 = htmlFontColor_Blue & rsSab("SWISABWMTK") & " Sortant vers " & htmlFontColor_Magenta & rsSab("SWISABWBIC")
    htmlFontColor_rText = htmlFontColor_Green
Else
    lX0 = htmlFontColor_Blue & rsSab("SWISABWMTK") & " reçu de " & htmlFontColor_Magenta & rsSab("SWISABWBIC")
    htmlFontColor_rText = htmlFontColor_Blue
End If

Call arrMT_Type_Scan(rsSab("SWISABWMTK"))

wSWISABWID1 = rsSab("SWISABWID1")
wSWISABWIDL = rsSab("SWISABWIDL")
wSWISABWIDH = rsSab("SWISABWIDH")
cmdSendMail_SWISABSWID = cmdSendMail_rText(wSWISABWID1, wSWISABWIDL, wSWISABWIDH)

End Function

Public Function cmdSendMail_SAA_Alerte_Approval(lFct As String) As String
Dim wSendMail As typeSendMail
Dim xSql As String
Dim xDétail_D As String, xHeader_D As String
Dim mbgColor As String
Dim X0 As String, X1 As String

On Error GoTo Error_Handler

Call srvrMesg_GetBuffer_ODBC(rsSIDE_Loop, xrMesg)
Call srvrIntv_GetBuffer_ODBC(rsSIDE_Loop, xrIntv)

Call arrMT_Fields_Load(xrMesg.mesg_type)

If xrMesg.mesg_sub_format = "INPUT" Then
    X0 = htmlFontColor_Blue & xrMesg.mesg_type & " Sortant vers " & Mid$(xrMesg.mesg_uumid, 2, 11)
    htmlFontColor_rText = htmlFontColor_Green
Else
    X0 = htmlFontColor_Blue & xrMesg.mesg_type & " reçu de " & Mid$(xrMesg.mesg_uumid, 2, 11)
    htmlFontColor_rText = htmlFontColor_Blue
End If

X1 = "Approuvé par " & xrIntv.intv_oper_nickname & "<BR>le " & xrIntv.intv_date_time _
    & "<BR>Créé par " & xrMesg.mesg_crea_oper_nickname & "<BR>le " & xrMesg.mesg_crea_date_time
If lFct = "USR" Then
    wSendMail.Subject = "Message Swift " & X1
    X = "Cet utilisateur n'est pas habilité à approuver des messages Swift<BR> sur la plateforme SAA," _
      & " (BIA_GOS - paramétrage SAA Alertes)<BR><BR>"
Else
    wSendMail.Subject = "Message Swift " & X1
    X = "Le montant de ce message Swift saisi sur la plateforme SAA est supérieur<BR> aux habilitations de cet utilisateur," _
      & " (BIA_GOS - paramétrage SAA Alertes)<BR><BR><BR>"
End If




xDétail_D = ""
mbgColor = "bgcolor = #FF00FF"
xHeader_D = "<TR>" _
         & "<TD bgcolor=#FFD0FF width=200 height=10><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Blue & X1 & "</TD>" _
         & "<TD bgcolor=#FFD0FF width=500 height=10><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Blue & X & "</TD>" _
        & "</TR>"



'-----------------------------------------------------------------------------------
    X = cmdSendMail_rText(CLng(xrMesg.Aid), xrMesg.mesg_s_umidl, xrMesg.mesg_s_umidh)
    xDétail_D = xDétail_D _
         & "<TD bgcolor = #FFFFFF width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & X0 & "</TD>" _
         & "<TD bgcolor = #FFFFFF width=500 height=7><span style='font-size:10.0pt;font-family:Courier New'>" & htmlFontColor_Blue & X & "</TD>" _
         & "</TR>"


'-----------------------------------------------------------------------------------
'wSendMail.FromDisplayName = "SAA_Sécurité"
'wSendMail.RecipientDisplayName = "SAA_Alerte"

wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = "SAA_Sécurité_Approval"
wSendMail.Recipient = mailAdresse_Recipient_Unit("S12", 2)
'-----------------------------------------------------------------------------------
'wSendMail.From = currentSSIWINMAIL

wSendMail.CcRecipient = ""
wSendMail.Attachment = ""


wSendMail.Message = "<" & mbgColor & "><BR>" _
                    & "<TABLE   width=700 border=1 cellpadding=4 ></B>" _
                    & xHeader_D _
                    & xDétail_D _
                    & "</TABLE>"
                    
'                    & htmlFontColor_Red & X _

wSendMail.AsHTML = True
srvSendMail.Monitor wSendMail
Exit Function

'------------------------------------------
Error_Handler:
    
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"

End Function
Public Function xxxxx_cmdSendMail_SAA_Alerte_rInst(lFct As String) As String
Dim wSendMail As typeSendMail
Dim xSql As String
Dim xDétail_D As String, xHeader_D As String
Dim mbgColor As String
Dim X0 As String, X1 As String

On Error GoTo Error_Handler

Call srvrMesg_GetBuffer_ODBC(rsSIDE_Loop, xrMesg)
'Call srvrInts_GetBuffer_ODBC(rsSIDE_Loop, xrInst)

Call arrMT_Fields_Load(xrMesg.mesg_type)

If xrMesg.mesg_sub_format = "INPUT" Then
    X0 = htmlFontColor_Blue & xrMesg.mesg_type & " Sortant vers " & Mid$(xrMesg.mesg_uumid, 2, 11)
    htmlFontColor_rText = htmlFontColor_Green
Else
    X0 = htmlFontColor_Blue & xrMesg.mesg_type & " reçu de " & Mid$(xrMesg.mesg_uumid, 2, 11)
    htmlFontColor_rText = htmlFontColor_Blue
End If

X1 = "le " & xrMesg.last_update
wSendMail.Subject = "Message Swift " & X1
X = "Ce message créée par " & Trim(xrMesg.mesg_crea_oper_nickname) & "<BR>, a été émis sans être approuvé sur la plateforme SAA."




xDétail_D = ""
mbgColor = "bgcolor = #FF00FF"
xHeader_D = "<TR>" _
         & "<TD bgcolor=#FFD0FF width=200 height=10><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Blue & X1 & "</TD>" _
         & "<TD bgcolor=#FFD0FF width=500 height=10><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Blue & X & "</TD>" _
        & "</TR>"



'-----------------------------------------------------------------------------------
    X = cmdSendMail_rText(CLng(xrMesg.Aid), xrMesg.mesg_s_umidl, xrMesg.mesg_s_umidh)
    xDétail_D = xDétail_D _
         & "<TD bgcolor = #FFFFFF width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & X0 & "</TD>" _
         & "<TD bgcolor = #FFFFFF width=500 height=7><span style='font-size:10.0pt;font-family:Courier New'>" & htmlFontColor_Blue & X & "</TD>" _
         & "</TR>"


'-----------------------------------------------------------------------------------
'wSendMail.FromDisplayName = "SAA_Sécurité"
'wSendMail.RecipientDisplayName = "SAA_Alerte"

wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = "SAA_Sécurité_Approval"
wSendMail.Recipient = mailAdresse_Recipient_Unit("S12", 2)
'-----------------------------------------------------------------------------------
'wSendMail.From = currentSSIWINMAIL

wSendMail.CcRecipient = ""
wSendMail.Attachment = ""


wSendMail.Message = "<" & mbgColor & "><BR>" _
                    & "<TABLE   width=700 border=1 cellpadding=4 ></B>" _
                    & xHeader_D _
                    & xDétail_D _
                    & "</TABLE>"
                    
'                    & htmlFontColor_Red & X _

wSendMail.AsHTML = True
srvSendMail.Monitor wSendMail
Exit Function

'------------------------------------------
Error_Handler:
    
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"

End Function


Public Function cmdSendMail_SAA_Alerte_Amount(lFct As String) As String
Dim wSendMail As typeSendMail
Dim xSql As String
Dim xDétail_D As String, xHeader_D As String
Dim mbgColor As String
Dim X0 As String, X1 As String, curX As Currency

On Error GoTo Error_Handler

Call srvrMesg_GetBuffer_ODBC(rsSIDE_Loop, xrMesg)

Call arrMT_Fields_Load(xrMesg.mesg_type)

If xrMesg.mesg_sub_format = "INPUT" Then
    X0 = htmlFontColor_Blue & xrMesg.mesg_type & " Sortant vers " & Mid$(xrMesg.mesg_uumid, 2, 11)
    htmlFontColor_rText = htmlFontColor_Green
Else
    X0 = htmlFontColor_Blue & xrMesg.mesg_type & " reçu de " & Mid$(xrMesg.mesg_uumid, 2, 11)
    htmlFontColor_rText = htmlFontColor_Blue
End If

X1 = "Créé par " & xrMesg.mesg_crea_oper_nickname & "<BR>le " & xrMesg.last_update

wSendMail.Subject = "Message Swift " & xrMesg.mesg_type & " : " & xrMesg.x_fin_ccy & " " & Format$(xrMesg.x_fin_amount, "### ### ### ### ##0.00")

Select Case lFct
    Case "103": curX = curSAA_103_EUR
    Case "202": curX = curSAA_202_EUR
    Case "202_BOTC": curX = curSAA_202_BOTC_EUR
End Select

X = "Le montant de ce message Swift est supérieur au seuil d'alerte : EUR " & Format$(curX, "### ### ### ### ##0") _
    & "<BR>(BIA_GOS - paramétrage SAA Alertes)<BR><BR>"

xDétail_D = ""
mbgColor = "bgcolor = #FF00FF"
xHeader_D = "<TR>" _
         & "<TD bgcolor=#FFD0FF width=200 height=10><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Blue & X1 & "</TD>" _
         & "<TD bgcolor=#FFD0FF width=500 height=10><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Blue & X & "</TD>" _
        & "</TR>"



'-----------------------------------------------------------------------------------
    X = cmdSendMail_rText(CLng(xrMesg.Aid), xrMesg.mesg_s_umidl, xrMesg.mesg_s_umidh)
    xDétail_D = xDétail_D _
         & "<TD bgcolor = #FFFFFF width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & X0 & "</TD>" _
         & "<TD bgcolor = #FFFFFF width=500 height=7><span style='font-size:10.0pt;font-family:Courier New'>" & htmlFontColor_Blue & X & "</TD>" _
         & "</TR>"


'-----------------------------------------------------------------------------------
'wSendMail.FromDisplayName = "SAA_Sécurité"
'wSendMail.RecipientDisplayName = "SAA_Alerte"
wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = "SAA_Sécurité_Amount"
wSendMail.Recipient = mailAdresse_Recipient_Unit("S12", 2)
wSendMail.CcRecipient = ""
'-----------------------------------------------------------------------------------
'wSendMail.From = currentSSIWINMAIL


wSendMail.Attachment = ""


wSendMail.Message = "<" & mbgColor & "><BR>" _
                    & "<TABLE   width=700 border=1 cellpadding=4 ></B>" _
                    & xHeader_D _
                    & xDétail_D _
                    & "</TABLE>"

wSendMail.AsHTML = True
srvSendMail.Monitor wSendMail
Exit Function

'------------------------------------------
Error_Handler:
    
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"

End Function
Public Function cmdSendMail_SAA_Alerte_rMesg(lFct As String, lSubject As String, lMsg As String, lMail_To As String, lMail_CC As String) As String
Dim wSendMail As typeSendMail
Dim xSql As String
Dim xDétail_D As String, xHeader_D As String
Dim mbgColor As String
Dim X0 As String, X1 As String, curX As Currency

On Error GoTo Error_Handler

Call arrMT_Fields_Load(xrMesg.mesg_type)

If xrMesg.mesg_sub_format = "INPUT" Then
    X0 = htmlFontColor_Blue & xrMesg.mesg_type & " Sortant vers " & Mid$(xrMesg.mesg_uumid, 2, 11)
    htmlFontColor_rText = htmlFontColor_Green
Else
    X0 = htmlFontColor_Blue & xrMesg.mesg_type & " reçu de " & Mid$(xrMesg.mesg_uumid, 2, 11)
    htmlFontColor_rText = htmlFontColor_Blue
End If

X1 = "le " & xrMesg.last_update

wSendMail.Subject = lSubject

xDétail_D = ""
mbgColor = "bgcolor = #FFFFFF"
xHeader_D = "<TR>" _
         & "<TD bgcolor=#FFFFA0 width=200 height=10><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Red & X1 & "</TD>" _
         & "<TD bgcolor=#FFFFA0 width=500 height=10><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Red & lMsg & "</TD>" _
        & "</TR>"



'-----------------------------------------------------------------------------------
    X = cmdSendMail_rText(CLng(xrMesg.Aid), xrMesg.mesg_s_umidl, xrMesg.mesg_s_umidh)
    xDétail_D = xDétail_D _
         & "<TD bgcolor = #FFFFFF width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & X0 & "</TD>" _
         & "<TD bgcolor = #FFFFFF width=500 height=7><span style='font-size:10.0pt;font-family:Courier New'>" & htmlFontColor_Blue & X & "</TD>" _
         & "</TR>"


'-----------------------------------------------------------------------------------
wSendMail.FromDisplayName = lFct
wSendMail.From = currentSSIWINMAIL

Select Case lFct
    Case "SAA_Event"
        wSendMail.Recipient = mailAdresse_Recipient_Unit(lMail_To, 2)
        If lMail_CC <> "" Then wSendMail.CcRecipient = mailAdresse_Recipient_Unit(lMail_CC, 2)
    Case "SAA_Live"
        wSendMail.Recipient = mailAdresse_Recipient_Unit(lMail_To, 2)
        If lMail_CC <> "" Then wSendMail.CcRecipient = mailAdresse_Recipient_Unit(lMail_CC, 2)
    Case "SAA_Routage"
        mbgColor = "bgcolor = #FFA0FF"
        wSendMail.Recipient = mailAdresse_Recipient_Unit(lMail_To, 2)
        If lMail_CC <> "" Then wSendMail.CcRecipient = mailAdresse_Recipient_Unit(lMail_CC, 2)
    Case "SAA_Origine_MT"
        mbgColor = "bgcolor = #FF80FF"
        wSendMail.Recipient = mailAdresse_Recipient_Unit(lMail_To, 2)
        If lMail_CC <> "" Then wSendMail.CcRecipient = mailAdresse_Recipient_Unit(lMail_CC, 2)
    Case Else
        wSendMail.Recipient = mailAdresse_Recipient_Unit(lMail_To, 2)
        If lMail_CC <> "" Then wSendMail.CcRecipient = mailAdresse_Recipient_Unit(lMail_CC, 2)
End Select
'-----------------------------------------------------------------------------------
If wSendMail.CcRecipient = "" Then
    wSendMail.CcRecipient = ""
Else
    wSendMail.CcRecipient = ";" & wSendMail.CcRecipient
End If

wSendMail.Attachment = ""


wSendMail.Message = "<body " & mbgColor & "><BR><span style='font-size:11.0pt;font-family:Calibri'>" & htmlFontColor_Blue & lMsg & "<BR><BR>" _
                    & "<TABLE   width=700 border=1 cellpadding=4 ></B>" _
                    & xHeader_D _
                    & xDétail_D _
                    & "</TABLE>"

wSendMail.AsHTML = True
srvSendMail.Monitor wSendMail
Exit Function

'------------------------------------------
Error_Handler:
    
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"

End Function

Public Function cmdSendMail_SAA_Alerte_rJrnl(lFct As String, lSubject As String, lMsg As String, lMail_To As String, lMail_CC As String) As String
Dim wSendMail As typeSendMail
Dim X As String
Dim xDétail_D As String, xHeader_D As String
Dim mbgColor As String
Dim X0 As String, X1 As String

On Error GoTo Error_Handler

X0 = htmlFontColor_Blue & xrJrnl.jrnl_oper_nickname

X1 = "le " & xrJrnl.jrnl_date_time

wSendMail.Subject = lSubject

xDétail_D = ""
mbgColor = "bgcolor = #FF00FF"
xHeader_D = "<TR>" _
         & "<TD bgcolor=#FFFFA0 width=200 height=10><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Red & X1 & "</TD>" _
         & "<TD bgcolor=#FFFFA0 width=500 height=10><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Red & lMsg & "</TD>" _
        & "</TR>"

'-----------------------------------------------------------------------------------
'Dim K As Integer
'For K = 0 To 25
'    X = X & rsSIDE_DB(K) & "<BR>"
'Next K
X = xrJrnl.jrnl_comp_name & " " & xrJrnl.jrnl_event_num & " " & xrJrnl.jrnl_event_name & "<BR>" _
  & xrJrnl.jrnl_event_severity & "<BR>" _
  & xrJrnl.jrnl_appl_serv_name & "<BR>" _
  & xrJrnl.jrnl_func_name & "<BR><BR>" _
  & xrJrnl.jrnl_alarm_status & "<BR>" _
  & xrJrnl.jrnl_alarm_date_time & "<BR>" _
  & xrJrnl.jrnl_alarm_oper_nickname & "<BR>"
  
    xDétail_D = xDétail_D _
         & "<TD bgcolor = #FFFFFF width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'>" & X0 & "</TD>" _
         & "<TD bgcolor = #FFFFFF width=500 height=7><span style='font-size:10.0pt;font-family:Courier New'>" & htmlFontColor_Blue & X & "</TD>" _
         & "</TR>"


'-----------------------------------------------------------------------------------
wSendMail.FromDisplayName = lFct
wSendMail.From = currentSSIWINMAIL

If lFct = "SAA_Event" Then
    wSendMail.Recipient = mailAdresse_Recipient_Unit(lMail_To, 2)
    If wSendMail.Recipient = "" Then
        xHeader_D = htmlFontColor_Red & "DESTINATAIRE : " & lMail_To & " inconnu. A REVOIR <BR><BR>" & xHeader_D
        wSendMail.Recipient = mailAdresse_Recipient_Unit("S12", 2)
    End If
Else
    wSendMail.Recipient = Trim(lMail_To)
    wSendMail.CcRecipient = Trim(lMail_CC)
End If
'-----------------------------------------------------------------------------------

If wSendMail.CcRecipient = "" Then
    wSendMail.CcRecipient = ""
Else
    wSendMail.CcRecipient = ";" & wSendMail.CcRecipient
End If


wSendMail.Message = "<" & mbgColor & "><BR>" _
                    & "<TABLE   width=700 border=1 cellpadding=4 ></B>" _
                    & xHeader_D _
                    & xDétail_D _
                    & "</TABLE>"

wSendMail.AsHTML = True
srvSendMail.Monitor wSendMail
Exit Function

'------------------------------------------
Error_Handler:
    
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"

End Function

Public Function cmdSendMail_SAA_Alerte(lFct As String, lSubject As String, lMsg As String, lMail_To As String, lMail_CC As String) As String
Dim wSendMail As typeSendMail
Dim X As String
Dim mbgColor As String
Dim X0 As String, X1 As String

On Error GoTo Error_Handler

X0 = htmlFontColor_Blue & xrJrnl.jrnl_oper_nickname

X1 = "le " & xrJrnl.jrnl_date_time

wSendMail.Subject = lSubject


'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
wSendMail.FromDisplayName = lFct
wSendMail.From = currentSSIWINMAIL

wSendMail.Recipient = mailAdresse_Recipient_Unit(lMail_To, 2)
If wSendMail.Recipient = "" Then
    wSendMail.Recipient = mailAdresse_Recipient_Unit("S12", 2)
End If
'-----------------------------------------------------------------------------------

wSendMail.CcRecipient = mailAdresse_Recipient_Unit(lMail_CC, 2)
If wSendMail.CcRecipient = "" Then
    wSendMail.CcRecipient = ""
Else
    wSendMail.CcRecipient = ";" & wSendMail.CcRecipient
End If


wSendMail.Message = lMsg

wSendMail.AsHTML = True
srvSendMail.Monitor wSendMail
Exit Function

'------------------------------------------
Error_Handler:
    
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"

End Function

Public Function cmdSendMail_ZSWIHIA0(lSWIHIANUM As Long) As String
Dim wSendMail As typeSendMail
Dim X As String, xSql As String
Dim xDétail_D As String, xHeader_D As String
Dim mbgColor As String
Dim X0 As String, X1 As String
Dim xSWIHIAUTI As String

On Error GoTo Error_Handler

xSql = "select * from  " & paramIBM_Library_SAB & ".zswihia0 " _
     & " WHERE swihianum = " & lSWIHIANUM
     
Set rsSabX = cnsab.Execute(xSql)

If rsSabX.EOF Then Exit Function

xSWIHIAUTI = rsSabX("SWIHIAUTI")
X1 = "A VERIFIER (N° " & lSWIHIANUM & ")"
X0 = "Le message a été envoyé de SAB le : " & dateImp10_S(19000000 + rsSabX("SWIHIADEN")) & " " & timeImp8(rsSabX("SWIHIAHEN")) _
   & " ( " & xSWIHIAUTI & ")<BR>Il n'a pas été identifié sur la plate-forme SWIFT ALLIANCE "

wSendMail.Subject = "Alerte concernant le message SWIFT : " & rsSabX("SWIHIAMES") & " " & rsSabX("SWIHIAREF") & " vers " & rsSabX("SWIHIADES")

xDétail_D = ""
mbgColor = "bgcolor = #FF00FF"
xHeader_D = "<TR>" _
         & "<TD bgcolor=#FFFFA0 width=200 height=10><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Red & X1 & "</TD>" _
         & "<TD bgcolor=#FFFFA0 width=800 height=10><span style='font-size:10.0pt;font-family:Calibri'>" & htmlFontColor_Red & X0 & "</TD>" _
        & "</TR>"
        
xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIHIB0 " _
     & " where SWIHIBETA = " & currentZMNURUT0.MNURUTETB _
     & " and SWIHIBNUM = " & lSWIHIANUM _
     & " order by SWIHIBNEN , SWIHIBNLI"
Set rsSabX = cnsab.Execute(xSql)

Do Until rsSabX.EOF

'-----------------------------------------------------------------------------------
    xDétail_D = xDétail_D _
         & "<TD bgcolor = #FFFFFF width=200 height=7><span style='font-size:10.0pt;font-family:Calibri'></TD>" _
         & "<TD bgcolor = #FFFFFF width=500 height=7><span style='font-size:10.0pt;font-family:Courier New'>" & htmlFontColor_Blue & Trim(rsSabX("SWIHIBDET")) & "</TD>" _
         & "</TR>"

    rsSabX.MoveNext
Loop



'-----------------------------------------------------------------------------------
wSendMail.FromDisplayName = "SAB_SAA anomalie"
wSendMail.From = currentSSIWINMAIL

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN'" _
     & " and SSIDOMUIDX = '" & xSWIHIAUTI & "'"
Set rsSabX = cnsab.Execute(xSql)

If rsSabX.EOF Then
    X = "S01"
Else
    X = Mid$(rsSabX("BIATABTXT"), 26, 3)
End If


wSendMail.Recipient = mailAdresse_Recipient_Unit(X, 2)
'-----------------------------------------------------------------------------------

If wSendMail.CcRecipient = "" Then
    wSendMail.CcRecipient = ""
Else
    wSendMail.CcRecipient = ";" & wSendMail.CcRecipient
End If


wSendMail.Message = "<" & mbgColor & "><BR>" _
                    & "<TABLE   width=700 border=1 cellpadding=4 ></B>" _
                    & xHeader_D _
                    & xDétail_D _
                    & "</TABLE>"

wSendMail.AsHTML = True
srvSendMail.Monitor wSendMail
Exit Function

'------------------------------------------
Error_Handler:
    
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"

End Function


Public Function cmdSendMail_rText(lSWISABWID1 As Long, lSWISABWIDL As Long, lSWISABWIDH As Long) As String
Dim xSql As String, xDetail As String, X As String
Dim xValue As String
Dim xText_Data_Block As String, xField_Code As String
Dim K As Integer, K2 As Integer, iAsc13 As Integer
Dim blnField_79 As Boolean
Dim xField_Lib As String, xField_CodeX As String
Dim V

blnField_79 = False

On Error GoTo Error_Handler

xDetail = ""

xSql = "select * from rtextField " _
    & "where Aid = " & lSWISABWID1 _
    & " and text_s_umidl = " & lSWISABWIDL _
    & " and text_s_umidh  =  " & lSWISABWIDH _
    & " order by field_cnt"

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
If Not rsSIDE_DB.EOF Then
    Do While Not rsSIDE_DB.EOF
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

        xField_Code = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
        xField_Lib = htmlFontColor_Gray & arrMT_Fields_Scan(xField_Code) & "<BR/>&#160;&#160;&#160;&#160;&#160;"

        If Len(xField_Code) = 2 Then
             xField_CodeX = xField_Code & "&#160;" & ": "
        Else
             xField_CodeX = xField_Code & ": "
        End If
        xDetail = xDetail & htmlFontColor_Magenta & xField_CodeX & xField_Lib & htmlFontColor_rText & cmdSendMail_rText_Line(Trim(xValue))  '& "<BR>"
        Select Case xField_Code
            Case "52A", "56A", "57A", "58A", "51A", "42A", "53A"
                xDetail = xDetail & ZSWIBIC0_Select_Html(xValue)
        End Select
        
        rsSIDE_DB.MoveNext
    Loop
Else
    xSql = "select * from rtext " _
        & "where Aid = " & lSWISABWID1 _
        & " and text_s_umidl = " & lSWISABWIDL _
        & " and text_s_umidh  =  " & lSWISABWIDH
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    If Not rsSIDE_DB.EOF Then
        Call srvrText_GetBuffer_ODBC(rsSIDE_DB, xrText)
            xText_Data_Block = xrText.text_data_block & Asc13
            'iLen = Len(xText_Data_Block)
            If Mid$(xText_Data_Block, 1, 3) = Asc13 & Asc10 & ":" Then
                K = 3
            Else
                K = 1
            End If
            Do
                xField_CodeX = "": xField_Lib = ""
                iAsc13 = InStr(K, xText_Data_Block, Asc13)
                If iAsc13 > 0 Then
                    X = Trim(Mid$(xText_Data_Block, K, iAsc13 - K))
                    If Mid$(X, 1, 1) <> ":" Then
                        xValue = Trim(Mid$(xText_Data_Block, K, iAsc13 - K))
                    Else
                        If blnField_79 Then
                            K2 = 0
                        Else
                            K2 = InStr(2, X, ":")
                        End If
                        'K2 = InStr(2, x, ":")
                        If K2 > 0 Then
                            xValue = Trim(Mid$(X, K2 + 1, Len(X) - K2))
                            xField_Code = Trim(Mid$(X, 2, K2 - 2))
                            xField_Lib = htmlFontColor_Gray & arrMT_Fields_Scan(xField_Code) & "<BR>"
                            If xField_Code = "79" Then blnField_79 = True
                           If Len(xField_Code) = 2 Then
                                xField_CodeX = xField_Code & "&#160;" & ": "
                           Else
                                xField_CodeX = xField_Code & ": "
                           End If

                        Else
                            xValue = Trim(Mid$(xText_Data_Block, K, iAsc13 - K))
                        End If
                    End If
                    
                    xDetail = xDetail & htmlFontColor_Magenta & xField_CodeX & xField_Lib & htmlFontColor_rText & "&#160;&#160;&#160;&#160;&#160;" & xValue & "<BR>"
                    Select Case xField_Code
                        Case "52A", "56A", "57A", "58A", "51A", "42A", "53A"
                            xDetail = xDetail & ZSWIBIC0_Select_Html(xValue)
                    End Select
                    K = iAsc13 + 2
                End If
             Loop Until iAsc13 = 0

    End If
End If

cmdSendMail_rText = xDetail

Exit Function
'------------------------------------------
Error_Handler:
    
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"


End Function

Public Function cmdSendMail_GOSDOSITOP(lSWISABWID1 As Long, lSWISABWIDL As Long, lSWISABWIDH As Long, lGOSDOSITOP As String) As String
Dim xSql As String, xDetail As String
Dim arrX(500) As String, Nb As Integer
Dim K As Integer, K2 As Integer, iAsc13 As Integer, iLen As Integer, I As Integer
Dim blnAsc13 As Boolean
Dim xValue As String
Dim xText_Data_Block As String, xField_Code As String
Dim xField_Lib As String, xField_CodeX As String
Dim blnSuite As Boolean
Dim blnField_79 As Boolean
Dim V
On Error GoTo Error_Handler
blnField_79 = False

xSql = "select * from rtextField " _
    & "where Aid = " & lSWISABWID1 _
    & " and text_s_umidl = " & lSWISABWIDL _
    & " and text_s_umidh  =  " & lSWISABWIDH _
    & " order by field_cnt"

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
Nb = 0
If Not rsSIDE_DB.EOF Then
    Do While Not rsSIDE_DB.EOF
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
        xField_Code = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
        xField_Lib = htmlFontColor_Gray & arrMT_Fields_Scan(xField_Code) & "<BR/>&#160;&#160;&#160;&#160;&#160;"

        If Len(xField_Code) = 2 Then
             xField_CodeX = xField_Code & "&#160;" & ": "
        Else
             xField_CodeX = xField_Code & ": "
        End If
        xValue = htmlFontColor_Magenta & xField_CodeX & xField_Lib & htmlFontColor_Blue & xValue
        'X = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
        ' If Len(X) = 2 Then X = X & "&#160;"
         'xValue = htmlFontColor_Blue & X & " : " & htmlFontColor_gray & rsSIDE_DB("value")
         blnSuite = False
         iLen = Len(xValue)
         K = 1
         Do
            iAsc13 = InStr(K, xValue, Asc13)
            If iAsc13 > 0 Then
                Nb = Nb + 1
                If Not blnSuite Then
                    arrX(Nb) = Trim(Mid$(xValue, K, iAsc13 - K))
                Else
                    arrX(Nb) = "&#160;&#160;&#160;&#160;&#160;" & htmlFontColor_Blue & Trim(Mid$(xValue, K, iAsc13 - K))
                End If
                
                K = iAsc13 + 2
                blnSuite = True
            End If
         Loop Until iAsc13 = 0
        
        Nb = Nb + 1
        If Not blnSuite Then
            arrX(Nb) = Trim(Mid$(xValue, K, iLen - K + 1))
        Else
            arrX(Nb) = "&#160;&#160;&#160;&#160;&#160;" & htmlFontColor_Blue & Trim(Mid$(xValue, K, iLen - K + 1))
        End If
        
        
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
            
            xText_Data_Block = xrText.text_data_block & Asc13
            iLen = Len(xText_Data_Block)
            If Mid$(xText_Data_Block, 1, 3) = Asc13 & Asc10 & ":" Then
                K = 3
            Else
                K = 1
            End If
            Do
                xField_CodeX = "": xField_Lib = ""
                iAsc13 = InStr(K, xText_Data_Block, Asc13)
                If iAsc13 > 0 Then
                    X = Trim(Mid$(xText_Data_Block, K, iAsc13 - K))
                    If Mid$(X, 1, 1) <> ":" Then
                        xValue = Trim(Mid$(xText_Data_Block, K, iAsc13 - K))
                    Else
                        'K2 = InStr(2, x, ":")
                        If blnField_79 Then
                            K2 = 0
                        Else
                            K2 = InStr(2, X, ":")
                        End If
                        If K2 > 0 Then
                            xValue = Trim(Mid$(X, K2 + 1, Len(X) - K2))
                            xField_Code = Trim(Mid$(X, 2, K2 - 2))
                            xField_Lib = htmlFontColor_Gray & arrMT_Fields_Scan(xField_Code) & "<BR>"
                            If xField_Code = "79" Then blnField_79 = True
                           If Len(xField_Code) = 2 Then
                                xField_CodeX = xField_Code & "&#160;" & ": "
                           Else
                                xField_CodeX = xField_Code & ": "
                           End If

                           ' If Len(xField_Code) = 3 Then xField_Code = xField_Code & "&#160;"
                        Else
                            xValue = Trim(Mid$(xText_Data_Block, K, iAsc13 - K))
                        End If
                    End If
                    
                    Nb = Nb + 1
                   ' arrX(Nb) = htmlFontColor_Blue & xField_Code & " : " & htmlFontColor_gray & xValue
                    arrX(Nb) = htmlFontColor_Magenta & xField_CodeX & xField_Lib & htmlFontColor_Blue & "&#160;&#160;&#160;&#160;&#160;" & xValue
                    K = iAsc13 + 2
                End If
             Loop Until iAsc13 = 0
 End If
End If
'-----------------------------------------
For I = 1 To 10
    K = Val(Mid$(oldYGOSDOS0.GOSDOSITOP, I * 2 - 1, 2))
    If K > 0 Then arrX(K) = Replace(arrX(K), htmlFontColor_Blue, htmlFontColor_Red)
    'arrX(K) = htmlFontColor_Red & arrX(K) & htmlFontColor_gray
Next I

For I = 1 To Nb
    xDetail = xDetail & arrX(I) & "<BR>"
Next I
cmdSendMail_GOSDOSITOP = xDetail
Exit Function
'------------------------------------------
Error_Handler:
    
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdSendMail_GOSDOSITOP"
    
cmdSendMail_GOSDOSITOP = xDetail

End Function




Public Function ZSWIENA0_Read()
Dim xSql As String, Nb As Integer
'Dim paramIBM_Library_SABJPL As String
ZSWIENA0_Read = Null


If oldYSWISAB0.SWISABZSWI <> 0 Then
    xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIENA0 " _
         & " where SWIENAETA = " & currentZMNURUT0.MNURUTETB _
         & " and SWIENAINT = " & oldYSWISAB0.SWISABZSWI

Else
    xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIENA0 " _
         & " where SWIENAETA = " & currentZMNURUT0.MNURUTETB _
         & " and SWIENAREF = '" & oldYSWISAB0.SWISABWL20 & "'" _
         & " and SWIENAMES = '" & oldYSWISAB0.SWISABWMTK & "'" _
         & " and SWIENADE1 = '" & oldYSWISAB0.SWISABWDEV & "'" _
         & " and SWIENAMON = " & cur_P(oldYSWISAB0.SWISABWMTD) _
         & " and SWIENAEME like '" & Mid$(oldYSWISAB0.SWISABWBIC, 1, 8) & "%'" _
         & " and SWIENADRE >= " & oldYSWISAB0.SWISABWAMJ - 19000000
End If
Set rsSab = cnsab.Execute(xSql)

Nb = 0
Do While Not rsSab.EOF
    V = srvYSWIENA0_GetBuffer_ODBC(rsSab, oldZSWIENA0)
    Nb = Nb + 1
    rsSab.MoveNext
Loop
Select Case Nb
    Case 0:
            ZSWIENA0_Read = "Le message n'a pas été trouvé dans le fichier des swifts en cours (ZSWIENA0)"
    Case Is > 1:
            ZSWIENA0_Read = "Il y a plus d'UN message doublon dans le fichier des swifts en cours (ZSWIENA0)"
    Case Else:
            If oldZSWIENA0.SWIENACET <> " " Then
                ZSWIENA0_Read = "Ce message est déjà annulé dans le fichier des swifts en cours (ZSWIENA0)"
            End If
End Select

End Function

Public Sub Importation_SAA_SWISABZSWI()
Dim V, X As String, K As Long, K2 As Long, Nb As Long
Dim xSql As String
Dim arrYSWISAB0(5000) As typeYSWISAB0, arrYSWISAB0_Nb As Integer
Dim arrZSWIENA0(5000) As typeZSWIENA0, arrZSWIENA0_Nb As Integer, xZSWIENA0 As typeZSWIENA0
Dim blnSWISABZSWI As Boolean
Dim blnTransaction As Boolean, blnMatch As Boolean, blnMatch_OK As Boolean
Dim wAmj As String, wSWIENBNUM As Long
Dim arrSWIENA_NREF(5000) As String, wSWIENAREF As String, wSWIENADRE As Long
On Error GoTo Error_Handler


'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABZSWI : 1"): DoEvents
'________________________________________________________________________


currentAction = "Importation_SAA_SWISABZSWI"
'==================================================================
'recherche dernier n° SWIENAINT

'If autoSWISABZSWI = 0 Then
'    K = importSWISABSWID - 5000
'Else
'    K = autoSWISABSWID
'End If
blnSWISABZSWI = False
wAmj = DSys

Do
    xSql = "select SWISABSWID,SWISABZSWI from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
         & " where SWISABWAMJ >= " & Val(wAmj) & " and SWISABWES = 'E' and SWISABZSWI <> 0 order by SWISABZSWI desc "
    Set rsSab = cnsab.Execute(xSql)
    
    If Not rsSab.EOF Then
        blnSWISABZSWI = True
        autoSWISABSWID = rsSab("SWISABSWID")
        autoSWISABZSWI = rsSab("SWISABZSWI")
    End If
    wAmj = dateElp("Ouvré", -1, wAmj)
Loop Until blnSWISABZSWI = True

'==================================================================

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABZSWI : 2"): DoEvents
'________________________________________________________________________


arrYSWISAB0_Nb = 0
xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABSWID > " & autoSWISABSWID & " and SWISABWES = 'E' and SWISABZSWI = 0 order by SWISABSWID desc"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    arrYSWISAB0_Nb = arrYSWISAB0_Nb + 1
    V = rsYSWISAB0_GetBuffer(rsSab, arrYSWISAB0(arrYSWISAB0_Nb))
    If arrYSWISAB0_Nb > 5000 Then Exit Do
    rsSab.MoveNext
Loop
'==================================================================

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABZSWI : 3"): DoEvents
'________________________________________________________________________


arrZSWIENA0_Nb = 0
xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIENA0 " _
     & " where SWIENAINT > " & autoSWISABZSWI & " order by SWIENAINT desc"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    arrZSWIENA0_Nb = arrZSWIENA0_Nb + 1
    V = srvYSWIENA0_GetBuffer_ODBC(rsSab, arrZSWIENA0(arrZSWIENA0_Nb))
    arrSWIENA_NREF(arrZSWIENA0_Nb) = ""
    If arrZSWIENA0_Nb > 5000 Then Exit Do
    rsSab.MoveNext
Loop
'==================================================================

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABZSWI : 4"): DoEvents
'________________________________________________________________________


Nb = 0
xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIENB0 " _
     & " where SWIENBNUM > " & autoSWISABZSWI & " and SWIENBCHA in (20 , 21) and SWIENBIND = ' '"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    wSWIENBNUM = rsSab("SWIENBNUM")
    For K = 1 To arrZSWIENA0_Nb
        If arrZSWIENA0(K).SWIENAINT = wSWIENBNUM Then
            Select Case rsSab("SWIENBCHA")
                Case 20: arrZSWIENA0(K).SWIENAREF = Trim(rsSab("SWIENBVAL"))
                Case 21: arrSWIENA_NREF(K) = Trim(rsSab("SWIENBVAL"))
            End Select
            Exit For
        End If
    Next K
    rsSab.MoveNext
Loop

'==================================================================

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABZSWI : 5"): DoEvents
'________________________________________________________________________

Nb = 0
blnMatch_OK = True
For K = 1 To arrZSWIENA0_Nb

    xZSWIENA0 = arrZSWIENA0(K)
    wSWIENAREF = Trim(xZSWIENA0.SWIENAREF)
    wSWIENADRE = xZSWIENA0.SWIENADRE + 19000000
    blnMatch = False
    For K2 = 1 To arrYSWISAB0_Nb
        If wSWIENAREF = arrYSWISAB0(K2).SWISABWL20 _
        And arrSWIENA_NREF(K) = arrYSWISAB0(K2).SWISABWN20 _
        And xZSWIENA0.SWIENAMES = arrYSWISAB0(K2).SWISABWMTK _
        And Mid$(xZSWIENA0.SWIENAEME, 1, 8) = Mid$(arrYSWISAB0(K2).SWISABWBIC, 1, 8) _
        And wSWIENADRE >= arrYSWISAB0(K2).SWISABWAMJ _
        And arrYSWISAB0(K2).SWISABZSWI = 0 Then
       
            If xZSWIENA0.SWIENAMES <> "103" _
            And xZSWIENA0.SWIENAMES <> "200" _
            And xZSWIENA0.SWIENAMES <> "202" _
            And xZSWIENA0.SWIENAMES <> "210" Then
                blnMatch = True
            Else
                 If xZSWIENA0.SWIENADE1 = arrYSWISAB0(K2).SWISABWDEV _
                 And xZSWIENA0.SWIENAMON = arrYSWISAB0(K2).SWISABWMTD Then
                     blnMatch = True
                 Else
                    If xZSWIENA0.SWIENADE1 = arrYSWISAB0(K2).SWISABWDEV _
                    And Fix(xZSWIENA0.SWIENAMON) = Fix(arrYSWISAB0(K2).SWISABWMTD) Then
                       blnMatch = True
                    End If
                 End If
           End If
           If blnMatch Then
              arrYSWISAB0(K2).SWISABZSWI = xZSWIENA0.SWIENAINT
              arrZSWIENA0(K).SWIENAINT = 0
              Nb = Nb + 1
              Exit For
           End If
        End If
    Next K2
    If Not blnMatch Then
        blnMatch_OK = False ': Debug.Print xZSWIENA0.SWIENAINT
    End If
Next K
'==================================================================

'==================================================================

Call lstErr_AddItem(lstErr, cmdContext, "rapprochement SWIENAINT : " & Nb): DoEvents

If Nb = 0 Then GoTo Exit_sub

blnTransaction = True
'For K = 1 To arrZSWIENA0_Nb
'    blnOk = False
'    xZSWIENA0 = arrZSWIENA0(K)
'    For K2 = 1 To arrYSWISAB0_Nb
'        If arrYSWISAB0(K2).SWISABZSWI = xZSWIENA0.SWIENAINT Then blnOk = True: Exit For
'    Next K2
'    If Not blnOk Then Debug.Print xZSWIENA0.SWIENAINT

'Next K
'blnTransaction = True

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABZSWI : 6"): DoEvents
'________________________________________________________________________


V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

For K = 1 To arrYSWISAB0_Nb
    If arrYSWISAB0(K).SWISABZSWI <> 0 Then
        oldYSWISAB0 = arrYSWISAB0(K)
        oldYSWISAB0.SWISABZSWI = 0
        V = sqlYSWISAB0_Update(arrYSWISAB0(K), oldYSWISAB0)
        If Not IsNull(V) Then GoTo Exit_sub
    End If
        
Next K

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABZSWI : 7"): DoEvents
'________________________________________________________________________


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABZSWI : 8"): DoEvents
'________________________________________________________________________


End Sub
Public Sub Importation_SAA_SWISABZSWI_Reprise()
Dim V, X As String, K As Long, K2 As Long, Nb As Long
Dim xSql As String
Dim arrN() As Long, arrN_Nb As Long, arrN_Ok As Long
Dim blnSWISABZSWI As Boolean
Dim blnTransaction As Boolean
Dim mSWISABZSWI As Long, wSWISABZSWI As Long
Dim ibmAMJ_Min As Long

On Error GoTo Error_Handler
'==================================================================

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABZSWI_Reprise : 1"): DoEvents
'________________________________________________________________________


xSql = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABZSWI = 0 and SWISABWES = 'E' and SWISABWAMJ > " & YBIATAB0_DATE_CPT_MP2 & " and SWISABWMTK in (103,202,700)"
Set rsSab = cnsab.Execute(xSql)

K = rsSab("Tally")

If K = 0 Then Exit Sub
'==================================================================
xSql = "select SWISABWAMJ  from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABZSWI = 0 and SWISABWES = 'E' and SWISABWAMJ > " & YBIATAB0_DATE_CPT_MP2 & " and SWISABWMTK in (103,202,700) order by SWISABWAMJ"
Set rsSab = cnsab.Execute(xSql)

If rsSab.EOF Then Exit Sub

ibmAMJ_Min = dateIBM(rsSab("SWISABWAMJ"))


currentAction = "Importation_SAA_SWISABZSWI_Reprise"
'==================================================================
ZSWIENA0_Reprise:
'==================================================================
arrN_Nb = 0: arrN_Ok = 0
xSql = "select count(*) as Tally from " & paramIBM_Library_SAB & ".ZSWIENA0 " _
     & " where SWIENADRE >= " & ibmAMJ_Min
Set rsSab = cnsab.Execute(xSql)

K = rsSab("Tally") + 1
ReDim arrN(K)

xSql = "select SWIENAINT from " & paramIBM_Library_SAB & ".ZSWIENA0 " _
     & " where SWIENADRE >= " & ibmAMJ_Min & "  order by SWIENAINT"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    arrN_Nb = arrN_Nb + 1
    arrN(arrN_Nb) = rsSab("SWIENAINT")
    rsSab.MoveNext
Loop

If arrN_Nb = 0 Then GoTo ZSWIMEA0_Reprise

xSql = "select SWISABZSWI from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABZSWI >= " & arrN(1) & " and SWISABWES = 'E' order by SWISABZSWI"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    wSWISABZSWI = rsSab("SWISABZSWI")
    For K = 1 To arrN_Nb
        If wSWISABZSWI = arrN(K) Then arrN(K) = 0: arrN_Ok = arrN_Ok + 1
    Next K
    rsSab.MoveNext
Loop

If arrN_Nb = arrN_Ok Then GoTo ZSWIMEA0_Reprise

For K = 1 To arrN_Nb
    If arrN(K) > 0 Then
        xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIENA0 " _
            & " where SWIENAINT = " & arrN(K)
        Set rsSab = cnsab.Execute(xSql)
        If Not rsSab.EOF Then
            matchYSWISAB0.SWISABWBIC = Mid$(rsSab("SWIENAEME"), 1, 8)
            matchYSWISAB0.SWISABWMTK = rsSab("SWIENAMES")
            matchYSWISAB0.SWISABWL20 = Trim(rsSab("SWIENAREF"))
            matchYSWISAB0.SWISABWN20 = ""
            matchYSWISAB0.SWISABWDEV = rsSab("SWIENADE1")
            matchYSWISAB0.SWISABWMTD = rsSab("SWIENAMON")
            matchYSWISAB0.SWISABZSWI = rsSab("SWIENAINT")
            matchYSWISAB0.SWISABWAMJ = rsSab("SWIENADRE") + 19000000
            xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIENB0 " _
                 & " where SWIENBETA = 1 and  SWIENBNUM = " & matchYSWISAB0.SWISABZSWI & " and SWIENBCHA in (20 , 21) and SWIENBIND = ' '"
            Set rsSab = cnsab.Execute(xSql)
            
            Do While Not rsSab.EOF
                Select Case rsSab("SWIENBCHA")
                    Case 20: matchYSWISAB0.SWISABWL20 = Trim(rsSab("SWIENBVAL"))
                    Case 21: matchYSWISAB0.SWISABWN20 = Trim(rsSab("SWIENBVAL"))
                End Select
                rsSab.MoveNext
            Loop

            Importation_SAA_SWISABZSWI_Reprise_Match
        End If
    End If
            
Next K



'==================================================================
ZSWIMEA0_Reprise:
'==================================================================

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABZSWI_Reprise : 2"): DoEvents
'________________________________________________________________________


arrN_Nb = 0: arrN_Ok = 0
xSql = "select count(*) as Tally from " & paramIBM_Library_SAB & ".ZSWIMEA0 " _
     & " where SWIMEADRE >= " & ibmAMJ_Min
Set rsSab = cnsab.Execute(xSql)

K = rsSab("Tally") + 1
ReDim arrN(K)

xSql = "select SWIMEANUM from " & paramIBM_Library_SAB & ".ZSWIMEA0 " _
     & " where SWIMEADRE >= " & ibmAMJ_Min & "  order by SWIMEANUM"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    arrN_Nb = arrN_Nb + 1
    arrN(arrN_Nb) = rsSab("SWIMEANUM")
    rsSab.MoveNext
Loop

If arrN_Nb = 0 Then GoTo BSWIENA0_Reprise

xSql = "select SWISABZSWI from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABZSWI >= " & arrN(1) & " and SWISABWES = 'E' order by SWISABZSWI"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    wSWISABZSWI = rsSab("SWISABZSWI")
    For K = 1 To arrN_Nb
        If wSWISABZSWI = arrN(K) Then arrN(K) = 0: arrN_Ok = arrN_Ok + 1
    Next K
    rsSab.MoveNext
Loop

If arrN_Nb = arrN_Ok Then GoTo BSWIENA0_Reprise
For K = 1 To arrN_Nb
    If arrN(K) > 0 Then
        xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIMEA0 " _
            & " where SWIMEANUM = " & arrN(K)
        Set rsSab = cnsab.Execute(xSql)
        If Not rsSab.EOF Then
            matchYSWISAB0.SWISABWBIC = Mid$(rsSab("SWIMEAEME"), 1, 8)
            matchYSWISAB0.SWISABWMTK = rsSab("SWIMEAMES")
            matchYSWISAB0.SWISABWL20 = Trim(rsSab("SWIMEAREF"))
            matchYSWISAB0.SWISABWN20 = ""
            matchYSWISAB0.SWISABWDEV = rsSab("SWIMEADEV")
            matchYSWISAB0.SWISABWMTD = rsSab("SWIMEAMON")
            matchYSWISAB0.SWISABZSWI = rsSab("SWIMEANUM")
            matchYSWISAB0.SWISABWAMJ = rsSab("SWIMEADRE") + 19000000
            xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIMEB0 " _
                 & " where  SWIMEBETA = 1 and  SWIMEBNUM = " & matchYSWISAB0.SWISABZSWI & " and SWIMEBCHA in (20 , 21) and SWIMEBIND = ' '"
            Set rsSab = cnsab.Execute(xSql)
            
            Do While Not rsSab.EOF
                Select Case rsSab("SWIMEBCHA")
                    Case 20: matchYSWISAB0.SWISABWL20 = Trim(rsSab("SWIMEBVAL"))
                    Case 21: matchYSWISAB0.SWISABWN20 = Trim(rsSab("SWIMEBVAL"))
                End Select
                rsSab.MoveNext
            Loop

            
            Importation_SAA_SWISABZSWI_Reprise_Match
        End If
    End If
            
Next K


'==================================================================
BSWIENA0_Reprise:
'==================================================================

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABZSWI_Reprise : 3"): DoEvents
'________________________________________________________________________


arrN_Nb = 0: arrN_Ok = 0
xSql = "select count(*) as Tally from " & paramIBM_Library_SABSPE & ".BSWIENA0 " _
     & " where SWIENADRE >= " & ibmAMJ_Min
Set rsSab = cnsab.Execute(xSql)

K = rsSab("Tally") + 1
ReDim arrN(K)

xSql = "select SWIENAINT from " & paramIBM_Library_SABSPE & ".BSWIENA0 " _
     & " where SWIENADRE >= " & ibmAMJ_Min & "  order by SWIENAINT"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    arrN_Nb = arrN_Nb + 1
    arrN(arrN_Nb) = rsSab("SWIENAINT")
    rsSab.MoveNext
Loop

If arrN_Nb = 0 Then Exit Sub

xSql = "select SWISABZSWI from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABZSWI >= " & arrN(1) & " and SWISABWES = 'E' order by SWISABZSWI"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    wSWISABZSWI = rsSab("SWISABZSWI")
    For K = 1 To arrN_Nb
        If wSWISABZSWI = arrN(K) Then arrN(K) = 0: arrN_Ok = arrN_Ok + 1
    Next K
    rsSab.MoveNext
Loop

If arrN_Nb = arrN_Ok Then Exit Sub
For K = 1 To arrN_Nb
    If arrN(K) > 0 Then
        xSql = "select * from " & paramIBM_Library_SABSPE & ".BSWIENA0 " _
            & " where SWIENAINT = " & arrN(K)
        Set rsSab = cnsab.Execute(xSql)
        If Not rsSab.EOF Then
            matchYSWISAB0.SWISABWBIC = Mid$(rsSab("SWIENAEME"), 1, 8)
            matchYSWISAB0.SWISABWMTK = rsSab("SWIENAMES")
            matchYSWISAB0.SWISABWL20 = Trim(rsSab("SWIENAREF"))
            matchYSWISAB0.SWISABWN20 = ""
            matchYSWISAB0.SWISABWDEV = rsSab("SWIENADE1")
            matchYSWISAB0.SWISABWMTD = rsSab("SWIENAMON")
            matchYSWISAB0.SWISABZSWI = rsSab("SWIENAINT")
            matchYSWISAB0.SWISABWAMJ = rsSab("SWIENADRE") + 19000000
            xSql = "select * from " & paramIBM_Library_SABSPE & ".BSWIENB0 " _
                 & " where SWIENBETA = 1 and  SWIENBNUM = " & matchYSWISAB0.SWISABZSWI & " and SWIENBCHA in (20 , 21) and SWIENBIND = ' '"
            Set rsSab = cnsab.Execute(xSql)
            
            Do While Not rsSab.EOF
                Select Case rsSab("SWIENBCHA")
                    Case 20: matchYSWISAB0.SWISABWL20 = Trim(rsSab("SWIENBVAL"))
                    Case 21: matchYSWISAB0.SWISABWN20 = Trim(rsSab("SWIENBVAL"))
                End Select
                rsSab.MoveNext
            Loop

            Importation_SAA_SWISABZSWI_Reprise_Match
        End If
    End If
            
Next K



'==================================================================

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABZSWI_Reprise : 4"): DoEvents
'________________________________________________________________________


End Sub



Public Sub Importation_SAA_SWISABZSWI_Reprise_Match()
Dim V, xSql As String
Dim blnMatch As Boolean, blnTransaction As Boolean


blnMatch = False

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABZSWI =0 and SWISABWES = 'E' and SWISABWMTK = '" & matchYSWISAB0.SWISABWMTK & "'" _
     & " and SWISABWBIC like '" & Trim(matchYSWISAB0.SWISABWBIC) & "%'"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
    
     If matchYSWISAB0.SWISABWL20 = oldYSWISAB0.SWISABWL20 _
     And matchYSWISAB0.SWISABWN20 = oldYSWISAB0.SWISABWN20 _
     And matchYSWISAB0.SWISABWMTK = oldYSWISAB0.SWISABWMTK _
     And Mid$(matchYSWISAB0.SWISABWBIC, 1, 8) = Mid$(oldYSWISAB0.SWISABWBIC, 1, 8) _
     And matchYSWISAB0.SWISABWAMJ >= oldYSWISAB0.SWISABWAMJ _
     And oldYSWISAB0.SWISABZSWI = 0 Then
    
         If matchYSWISAB0.SWISABWMTK <> "103" _
         And matchYSWISAB0.SWISABWMTK <> "200" _
         And matchYSWISAB0.SWISABWMTK <> "202" _
         And matchYSWISAB0.SWISABWMTK <> "210" Then
             blnMatch = True
         Else
              If matchYSWISAB0.SWISABWDEV = oldYSWISAB0.SWISABWDEV _
              And matchYSWISAB0.SWISABWMTD = oldYSWISAB0.SWISABWMTD Then
                  blnMatch = True
              Else
                 If matchYSWISAB0.SWISABWDEV = oldYSWISAB0.SWISABWDEV _
                 And Fix(matchYSWISAB0.SWISABWMTD) = Fix(oldYSWISAB0.SWISABWMTD) Then
                    blnMatch = True
                 End If
              End If
        End If
    
        If blnMatch Then Exit Do
        
     End If
    rsSab.MoveNext
Loop
'================================================================================
'If blnMatch Then GoTo Match_True
'================================================================================



If Not blnMatch Then Exit Sub


'================================================================================
Match_True:
'================================================================================

blnTransaction = True
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

newYSWISAB0 = oldYSWISAB0
newYSWISAB0.SWISABZSWI = matchYSWISAB0.SWISABZSWI

V = sqlYSWISAB0_Update(newYSWISAB0, oldYSWISAB0)
If Not IsNull(V) Then GoTo Exit_sub
        
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation_SAA_SWISABZSWI_Reprise_Match"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If

End Sub

Public Sub SWISABWN20_Reprise()
Dim xSql As String
Dim xrText As typerText
Dim xrMesg As typerMesg

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABSWID = 379869 " _
     & " and SWISABWES = 'E' "

Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF

'===================================================================

    xSql = "select * from rMesg " _
        & "where Aid = " & rsSab("SWISABWID1") _
        & " and Mesg_s_umidl = " & rsSab("SWISABWIDL") _
        & " and Mesg_s_umidh  =  " & rsSab("SWISABWIDH")
   Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    If Not rsSIDE_DB.EOF Then
    V = srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)

      '  Debug.Print rsSab("SWISABSWID"), rsSIDE_DB("value")
    End If
'====================================================

    xSql = "select * from rtext " _
        & "where Aid = " & rsSab("SWISABWID1") _
        & " and text_s_umidl = " & rsSab("SWISABWIDL") _
        & " and text_s_umidh  =  " & rsSab("SWISABWIDH")
   Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    If Not rsSIDE_DB.EOF Then
    V = srvrText_GetBuffer_ODBC(rsSIDE_DB, xrText)

      '  Debug.Print rsSab("SWISABSWID"), rsSIDE_DB("value")
    End If
'====================================================
    rsSab.MoveNext
Loop

End Sub

Public Sub Param_SWISABWSRV_Init()
New_YBIATAB0.BIATABID = "SWISABWSRV"

New_YBIATAB0.BIATABK2 = ""

New_YBIATAB0.BIATABK1 = "S01": New_YBIATAB0.BIATABTXT = "SOBF": Parametrage_New
New_YBIATAB0.BIATABK1 = "S10": New_YBIATAB0.BIATABTXT = "SOBI": Parametrage_New
New_YBIATAB0.BIATABK1 = "S11": New_YBIATAB0.BIATABTXT = "COBK": Parametrage_New
New_YBIATAB0.BIATABK1 = "S32": New_YBIATAB0.BIATABTXT = "BOTC": Parametrage_New
New_YBIATAB0.BIATABK1 = "S41": New_YBIATAB0.BIATABTXT = "DCOM": Parametrage_New
New_YBIATAB0.BIATABK1 = "S62": New_YBIATAB0.BIATABTXT = "STLX": Parametrage_New

End Sub

Public Sub Importation_SAB_ZCDODOS0()
Dim V, X As String, K As Long, K2 As Long, Nb As Long
Dim xSql As String, blnTransaction As Boolean
Dim blnUpdate As Boolean, blnOk As Boolean
Dim mCDODOSOUV As Long

On Error GoTo Error_Handler
'==================================================================
blnTransaction = False
Call lstErr_AddItem(lstErr, cmdContext, "Importation ZCDODOS0"): DoEvents
Call rsYSWISAB0_Init(oldYSWISAB0)
newYSWISAB0 = oldYSWISAB0

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
blnTransaction = True

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 , " & paramIBM_Library_SAB & ".ZCDODOS0" _
     & " where  SWISABSWID > " & mSWISABSWID_Xd & " and SWISABOPEN = 0 and substring(SWISABWMTK , 1 , 1) = '7' " _
     & " and SWISABWL20 <> '' and SWISABWL20 = CDODOSEXT " _
     & " order by SWISABSWID,CDODOSOUV"
     
Set rsSab = cnsab.Execute(xSql)
blnUpdate = False

Do While Not rsSab.EOF
    
    If rsSab("SWISABSWID") <> oldYSWISAB0.SWISABSWID Then
    
        If blnUpdate Or blnOk Then V = sqlYSWISAB0_Update(newYSWISAB0, oldYSWISAB0)
        blnUpdate = False
        blnOk = False
    End If
    
    If Not blnUpdate Then
        Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
        newYSWISAB0 = oldYSWISAB0
        If rsSab("CDODOSOUV") > rsSab("CDODOSEMI") Then
            mCDODOSOUV = rsSab("CDODOSOUV") + 19000000
        Else
            mCDODOSOUV = rsSab("CDODOSEMI") + 19000000
        End If
        blnOk = False
        If oldYSWISAB0.SWISABWMTK = "700" Then
            If oldYSWISAB0.SWISABWMTD = rsSab("CDODOSMON") And oldYSWISAB0.SWISABWDEV = rsSab("CDODOSDEV") _
            And oldYSWISAB0.SWISABWAMJ <= mCDODOSOUV Then
                blnOk = True: blnUpdate = True
            Else
                If oldYSWISAB0.SWISABWAMJ <= mCDODOSOUV Then blnOk = True
            End If
        Else
            If oldYSWISAB0.SWISABWAMJ >= mCDODOSOUV Then blnOk = True
        End If
        
        If blnOk Then
            If Not blnUpdate And Mid$(oldYSWISAB0.SWISABWL20, 1, 2) = "NO" Then
                blnOk = False
            End If
            newYSWISAB0.SWISABSER = rsSab("CDODOSSER")
            newYSWISAB0.SWISABSSE = rsSab("CDODOSSSE")
            newYSWISAB0.SWISABOPEC = rsSab("CDODOSCOP")
            newYSWISAB0.SWISABOPEN = rsSab("CDODOSDOS")
            newYSWISAB0.SWISABKSRV = "S10"
        End If
    End If
    Nb = Nb + 1
    If Nb = 100 Then
        Nb = 0
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "ZCDODOS0 : " & oldYSWISAB0.SWISABSWID & " " & newYSWISAB0.SWISABOPEN): DoEvents
    End If
    rsSab.MoveNext
Loop

If blnUpdate Or blnOk Then V = sqlYSWISAB0_Update(newYSWISAB0, oldYSWISAB0)

'==================================================================

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If


End Sub
Public Sub Importation_SAB_SWISABWN20()
Dim V, X As String, K As Long, K1 As Long, K2 As Long, Nb As Long, Nb_Update As Long
Dim xSql As String, blnTransaction As Boolean
Dim xOPEN As String, wOPEN As Double, blnOpen As Boolean, X1 As String, blnOPEC As Boolean
Dim blnOk As Boolean

On Error GoTo Error_Handler
'==================================================================
blnTransaction = False
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABN20"): DoEvents
rsYSWISAB0_Init oldYSWISAB0
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

Nb_Update = 0

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where  SWISABSWID > " & mSWISABSWID_Xd & " and SWISABOPEN = 0 and SWISABWN20 <> '' and SWISABWMTK <> '950' order by SWISABSWID"
If cmdSelect_SQL_K = "XX" Then
    'xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
    '     & " where  SWISABOPEN = 0 and SWISABWES = 'S' and SWISABWN20 <> '' and SWISABWMTK <> '950' order by SWISABSWID"
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
         & " where  SWISABOPEN = 0  and SWISABWN20 <> '' and SWISABWMTK <> '950' order by SWISABSWID"
End If

Set rsSab = cnsab.Execute(xSql)
blnTransaction = True

Do While Not rsSab.EOF
    blnOk = False
    xYSWISAB0.SWISABOPEC = "": xYSWISAB0.SWISABOPEN = 0
    X = UCase$(Trim(rsSab("SWISABWN20")))
    If Len(X) >= 8 Then
        If IsNumeric(Mid$(X, 8, Len(X) - 7)) Then
            xYSWISAB0.SWISABOPEC = Mid$(X, 5, 3)
            Select Case UCase$(Mid$(X, 1, 7))
                Case "SOBICDE", "SOBICDI", "SOBIGOS", "SOBIRDE", "SOBIRDI":
                            blnOk = True: xYSWISAB0.SWISABSER = "00": xYSWISAB0.SWISABSSE = "00": xYSWISAB0.SWISABKSRV = "S10"

                Case "ORPATRF", "ORPACPT", "SOBFGOS", "SOBFRV0", "ORPARDE", "ORPARDI":
                            blnOk = True: xYSWISAB0.SWISABSER = "00": xYSWISAB0.SWISABSSE = "TR": xYSWISAB0.SWISABKSRV = "S01"
                            
                Case "SOBITRF":
                            blnOk = True: xYSWISAB0.SWISABSER = "00": xYSWISAB0.SWISABSSE = "CD": xYSWISAB0.SWISABKSRV = "S10" '$DR 2014-10-22
                
                Case "BOTCCPT", "BOTCTRF", "BOTCTER", "BOTCPRE", "BOTCEMP", "BOTCGOS":
                            blnOk = True: xYSWISAB0.SWISABSER = "TC": xYSWISAB0.SWISABSSE = "TC": xYSWISAB0.SWISABKSRV = "S32"
            End Select
            If blnOk Then
                xYSWISAB0.SWISABOPEN = Val(Mid$(X, 8, Len(X) - 7))
                If xYSWISAB0.SWISABOPEC = "RDE" Or xYSWISAB0.SWISABOPEC = "RDI" Then
                    If xYSWISAB0.SWISABOPEN > 10000000 Then xYSWISAB0.SWISABOPEN = Fix(xYSWISAB0.SWISABOPEN / 100) ' $JPL 2014-09-09
                End If
            End If
        End If
'-------------------------------------------------------------------------------------------
        If Not blnOk Then
            xOPEN = "": wOPEN = 0: blnOpen = False: blnOPEC = False
            For K2 = 1 To Len(X)
                X1 = Mid$(X, K2, 1)
                If X1 <> " " And X1 <> "." Then
                    If IsNumeric(X1) Then
                        If blnOPEC Then blnOpen = True: xOPEN = xOPEN & X1
                    Else
                        blnOPEC = True
                        If blnOpen Then Exit For
                    End If
                End If
            Next K2
            wOPEN = Val(xOPEN)
            
            K1 = InStr(X, "CDE")
            If K1 > 0 Then
                If wOPEN > 50000 And wOPEN < 999999 Then
                    blnOk = True
                    xYSWISAB0.SWISABSER = "00": xYSWISAB0.SWISABSSE = "00": xYSWISAB0.SWISABKSRV = "S10"
                    xYSWISAB0.SWISABOPEC = "CDE"
                    xYSWISAB0.SWISABOPEN = wOPEN
                End If
            Else
                K1 = InStr(X, "CDI")
                If K1 > 0 Then
                    If wOPEN > 99999 And wOPEN < 999999 Then
                        blnOk = True
                        xYSWISAB0.SWISABSER = "00": xYSWISAB0.SWISABSSE = "00": xYSWISAB0.SWISABKSRV = "S10"
                        xYSWISAB0.SWISABOPEC = "CDI"
                        xYSWISAB0.SWISABOPEN = wOPEN
                    End If
                Else
                
                    K1 = InStr(X, "RDE")
                    If K1 > 0 Then
                        If wOPEN > 10000000 Then wOPEN = Fix(wOPEN / 100) ' $JPL 2014-05-26
                        If wOPEN > 99999 And wOPEN < 999999 Then
                            blnOk = True
                            xYSWISAB0.SWISABSER = "00": xYSWISAB0.SWISABSSE = "00": xYSWISAB0.SWISABKSRV = "S10"
                            xYSWISAB0.SWISABOPEC = "RDE"
                            xYSWISAB0.SWISABOPEN = wOPEN
                        End If
                    Else
            
                        K1 = InStr(X, "RDI")
                       If wOPEN > 10000000 Then wOPEN = Fix(wOPEN / 100) ' $JPL 2014-05-26
                       If K1 > 0 Then
                            If wOPEN > 99999 And wOPEN < 999999 Then
                                blnOk = True
                                xYSWISAB0.SWISABSER = "00": xYSWISAB0.SWISABSSE = "00": xYSWISAB0.SWISABKSRV = "S10"
                                xYSWISAB0.SWISABOPEC = "RDI"
                                xYSWISAB0.SWISABOPEN = wOPEN
                            End If
                        Else
                
                            K1 = InStr(X, "RV0")
                            If K1 > 0 Then
                                If wOPEN > 9999 And wOPEN < 999999 Then
                                    blnOk = True
                                    xYSWISAB0.SWISABSER = "00": xYSWISAB0.SWISABSSE = "00": xYSWISAB0.SWISABKSRV = "S01"
                                    xYSWISAB0.SWISABOPEC = "RV0"
                                    xYSWISAB0.SWISABOPEN = wOPEN
                                End If
                            Else
                
                                K1 = InStr(X, "RVO")
                                If K1 > 0 Then
                                    If wOPEN > 9999 And wOPEN < 999999 Then
                                        blnOk = True
                                        xYSWISAB0.SWISABSER = "00": xYSWISAB0.SWISABSSE = "00": xYSWISAB0.SWISABKSRV = "S01"
                                        xYSWISAB0.SWISABOPEC = "RV0"
                                        xYSWISAB0.SWISABOPEN = wOPEN
                                    End If
                                Else
            
                                    K1 = InStr(X, "ORPA")
                                    If K1 > 0 Then
                                        If wOPEN > 999 And wOPEN < 9999999 Then
                                            blnOk = True
                                            xYSWISAB0.SWISABSER = "00": xYSWISAB0.SWISABSSE = "TR": xYSWISAB0.SWISABKSRV = "S01"
                                            xYSWISAB0.SWISABOPEC = "TRF"
                                            xYSWISAB0.SWISABOPEN = wOPEN
                                        End If
                                    Else
                            
                                        K1 = InStr(X, "SOBF")
                                        If K1 > 0 Then
                                            If wOPEN > 99999 And wOPEN < 999999 Then
                                                blnOk = True
                                                xYSWISAB0.SWISABSER = "00": xYSWISAB0.SWISABSSE = "TR": xYSWISAB0.SWISABKSRV = "S01"
                                                xYSWISAB0.SWISABOPEC = "TRF"
                                                xYSWISAB0.SWISABOPEN = wOPEN
                                            End If
                                        Else
                                
                                            K1 = InStr(X, "DAFI/")
                                            If K1 > 0 Then
                                                K1 = InStr(K1 + 5, X, "/")
                                                If K1 > 0 Then
                                                    xOPEN = "": wOPEN = 0: blnOpen = False: blnOPEC = False
                                                    For K2 = K1 To Len(X)
                                                        X1 = Mid$(X, K2, 1)
                                                        If X1 <> " " And X1 <> "." Then
                                                            If IsNumeric(X1) Then
                                                                If blnOPEC Then blnOpen = True: xOPEN = xOPEN & X1
                                                            Else
                                                                blnOPEC = True
                                                                If blnOpen Then Exit For
                                                            End If
                                                        End If
                                                    Next K2
                                                    wOPEN = Val(xOPEN)

                                                    If wOPEN > 0 Then
                                                        blnOk = True
                                                        xYSWISAB0.SWISABSER = "00": xYSWISAB0.SWISABSSE = "00": xYSWISAB0.SWISABKSRV = "S32"
                                                        xYSWISAB0.SWISABOPEC = "GDC"
                                                        xYSWISAB0.SWISABOPEN = wOPEN
                                                    End If
                                                End If
                                        End If
                                    End If
'===================================================
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
'-------------------------------------------------------------------------------------------
    End If
    
        
    If blnOk Then
        Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
        newYSWISAB0 = oldYSWISAB0
        newYSWISAB0.SWISABSER = xYSWISAB0.SWISABSER
        newYSWISAB0.SWISABSSE = xYSWISAB0.SWISABSSE
        newYSWISAB0.SWISABOPEC = xYSWISAB0.SWISABOPEC
        newYSWISAB0.SWISABOPEN = xYSWISAB0.SWISABOPEN
        newYSWISAB0.SWISABKSRV = xYSWISAB0.SWISABKSRV
        Nb_Update = Nb_Update + 1
        Nb = Nb + 1
        If Nb = 100 Then
            Nb = 0
            Call lstErr_ChangeLastItem(lstErr, cmdContext, "SWISABWN20 : " & Nb_Update & " " & oldYSWISAB0.SWISABSWID & " " & newYSWISAB0.SWISABOPEN): DoEvents
        End If
        V = sqlYSWISAB0_Update(newYSWISAB0, oldYSWISAB0)
    End If
    
    rsSab.MoveNext
    'If Nb_Update > 20000 Then Exit Do
Loop



'==================================================================

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation_SAA_SWISABWN20"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If


End Sub

Public Sub Importation_SAB_YSWISAB1()
Dim V, X As String, K As Long, K1 As Long, K2 As Long, K3 As Long, Nb As Long, Nb_Update As Long
Dim xSql As String, blnTransaction As Boolean
Dim mField As String
Dim wPays_K As Integer
Dim wText_Data_Block As String
Dim bln202 As Boolean

On Error GoTo Error_Handler
'==================================================================

 'Call MsgBox("à faire 103 rtextfield")

blnTransaction = False
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_SWISABN20"): DoEvents
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

Nb_Update = 0


xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YSWISAB1 "
Set rsSab = cnsab.Execute(xSql)

xSql = "select SWISAB1ID from " & paramIBM_Library_SABSPE & ".YSWISAB1 " _
     & " where SWISAB1ID >= " & rsSab(0) & " order by SWISAB1ID desc "
Set rsSab = cnsab.Execute(xSql)

If rsSab.EOF Then
    mSWISABSWID_Xd = 0
Else
    mSWISABSWID_Xd = rsSab("SWISAB1ID")
End If

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where  SWISABSWID > " & mSWISABSWID_Xd & " and SWISABWMTK in ('103','202') order by SWISABSWID"

Set rsSab = cnsab.Execute(xSql)




blnTransaction = True

Do While Not rsSab.EOF
    wMT_BIC_E = "": wMT_BIC_S = ""
    wMT_50A = "": wMT_50P = ""
    wMT_52A = "": wMT_52P = ""
    wMT_57A = "": wMT_57P = ""
    wMT_58A = "": wMT_58P = ""
    wMT_59A = "": wMT_59P = ""
    rsYSWISAB1_Init newYSWISAB1
    newYSWISAB1.SWISAB1ID = rsSab("SWISABSWID")
    
    
    If rsSab("SWISABWES") = "E" Then
        wMT_BIC_E = rsSab("SWISABWBIC")
        wMT_BIC_S = "BIARFRPPXXX"
    Else
        wMT_BIC_E = "BIARFRPPXXX"
        wMT_BIC_S = rsSab("SWISABWBIC")
    End If
    
    If rsSab("SWISABWMTK") = "202" Then
        bln202 = True
        Importation_SAB_YSWISAB1_202
    Else
        bln202 = False
        Importation_SAB_YSWISAB1_103
   End If
    
  '  If  bln202 Then
    

        
    '_____________________________________________________________________________________________________
        If newYSWISAB1.SWISABW50P <> "" Then
            Call Importation_SAB_YSWISAB1_Pays(newYSWISAB1.SWISABW50P, wPays_K, newYSWISAB1.SWISABW50Z)
            If wPays_K = 0 Then newYSWISAB1.SWISABW50P = "": newYSWISAB1.SWISABW50Z = "**"
        End If
        'If newYSWISAB1.SWISABW50P = "" Then
        '    If Len(newYSWISAB1.SWISABW52A) >= 6 Then
        '        newYSWISAB1.SWISABW50P = Mid$(newYSWISAB1.SWISABW52A, 5, 2)
        '        Call Importation_SAB_YSWISAB1_Pays(newYSWISAB1.SWISABW50P, wPays_K, newYSWISAB1.SWISABW50Z)
        '        If wPays_K = 0 Then newYSWISAB1.SWISABW50P = "": newYSWISAB1.SWISABW50Z = "**"
        '    End If
        'End If
            
        If newYSWISAB1.SWISABW59P <> "" Then
            Call Importation_SAB_YSWISAB1_Pays(newYSWISAB1.SWISABW59P, wPays_K, newYSWISAB1.SWISABW59Z)
            If wPays_K = 0 Then newYSWISAB1.SWISABW59P = "": newYSWISAB1.SWISABW59Z = "**"
        End If
        'If newYSWISAB1.SWISABW59P = "" Then
        '    If Len(newYSWISAB1.SWISABW57A) >= 6 Then
        '        newYSWISAB1.SWISABW59P = Mid$(newYSWISAB1.SWISABW57A, 5, 2)
        '        Call Importation_SAB_YSWISAB1_Pays(newYSWISAB1.SWISABW59P, wPays_K, newYSWISAB1.SWISABW59Z)
        '        If wPays_K = 0 Then newYSWISAB1.SWISABW59P = "": newYSWISAB1.SWISABW59Z = "**"
        '    End If
        'End If
            
            
        Nb = Nb + 1
        If Nb = 100 Then
            Nb = 0
            Call lstErr_ChangeLastItem(lstErr, cmdContext, "Importation_SAB_YSWISAB1 : " & Nb_Update & " Id :" & newYSWISAB1.SWISAB1ID): DoEvents
        End If
        V = sqlYSWISAB1_Insert(newYSWISAB1)
        If IsNull(V) Then Nb_Update = Nb_Update + 1
   ' End If
'=========================================
'    If Nb_Update = 10000 Then Exit Do
'=========================================
    rsSab.MoveNext
Loop



'==================================================================

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation_SAA_SWISABWN20"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If


End Sub


Public Sub Importation_SAB_ZSWICLA0()
Dim V, X As String, K As Long, K2 As Long, Nb As Long
Dim xSql As String, blnTransaction As Boolean

On Error GoTo Error_Handler
'==================================================================
blnTransaction = False
Call lstErr_AddItem(lstErr, cmdContext, "Importation ZSWICLA0"): DoEvents



V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox


xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 , " & paramIBM_Library_SABSPE & ".YSWIMON0 , " & paramIBM_Library_SAB & ".ZSWICLA0" _
     & " where SWISABSWID > " & mSWISABSWID_Xd & " and SWISABOPEN = 0" _
     & " and  SWISABWID1 = SAAAID And SWISABWIDL = SAAUMIDL and swisabwidh = saaumidh and swisabnum = swiclaint order by SWISABSWID desc"
Set rsSab = cnsab.Execute(xSql)
blnTransaction = True

Do While Not rsSab.EOF
    Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
    newYSWISAB0 = oldYSWISAB0
    newYSWISAB0.SWISABZSWI = rsSab("SWICLAINT")
    newYSWISAB0.SWISABSER = rsSab("SWICLASER")
    newYSWISAB0.SWISABSSE = rsSab("SWICLASES")
    newYSWISAB0.SWISABOPEC = rsSab("SWICLAOPR")
    Select Case newYSWISAB0.SWISABOPEC
        Case "CDE", "CDI":   newYSWISAB0.SWISABOPEN = Fix(rsSab("SWICLANUM") / 10000)
        Case "RDE", "RDI":   newYSWISAB0.SWISABOPEN = Fix(rsSab("SWICLANUM") / 1000)
        Case Else:   newYSWISAB0.SWISABOPEN = rsSab("SWICLANUM")
    End Select
    
    Select Case newYSWISAB0.SWISABSSE
        Case "TC", "CR": newYSWISAB0.SWISABKSRV = "S32"
        Case "TR": newYSWISAB0.SWISABKSRV = "S01"
    End Select

    Nb = Nb + 1
    If Nb = 100 Then
        Nb = 0
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "ZSWICLA0 : " & oldYSWISAB0.SWISABSWID & " " & newYSWISAB0.SWISABOPEN): DoEvents
    End If
    V = sqlYSWISAB0_Update(newYSWISAB0, oldYSWISAB0)
    rsSab.MoveNext
Loop

'==================================================================

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If


End Sub

Public Sub Importation_SAB_YSWIMON0()
Dim V, X As String, K As Long, K2 As Long, Nb As Long
Dim xSql As String, blnTransaction As Boolean

On Error GoTo Error_Handler
'==================================================================
blnTransaction = False
Call lstErr_AddItem(lstErr, cmdContext, "Importation YSWIMON0"): DoEvents



V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox


xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 , " & paramIBM_Library_SABSPE & ".YSWIMON0 " _
     & " where SWISABSWID > " & mSWISABSWID_Xd & " and SWISABZSWI = 0 and SWISABWES = 'S' and SWISABWMTK <> '950'" _
     & " and  SWISABWID1 = SAAAID And SWISABWIDL = SAAUMIDL  And SWISABWIDH = SAAUMIDH order by SWISABSWID"
Set rsSab = cnsab.Execute(xSql)
blnTransaction = True

Do While Not rsSab.EOF
    Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
    newYSWISAB0 = oldYSWISAB0
    newYSWISAB0.SWISABZSWI = rsSab("SWISABNUM")

    Nb = Nb + 1
    If Nb = 100 Then
        Nb = 0
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "YSWIMON0 : " & oldYSWISAB0.SWISABSWID & " " & newYSWISAB0.SWISABZSWI): DoEvents
    End If
    V = sqlYSWISAB0_Update(newYSWISAB0, oldYSWISAB0)
    rsSab.MoveNext
Loop

'==================================================================

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If


End Sub
Public Sub Importation_SAB_YSWIMON0_Synchro1()
Dim V, X As String, K As Long, K2 As Long, Nb As Long
Dim xSql As String, blnTransaction As Boolean

On Error GoTo Error_Handler
'==================================================================
blnTransaction = False
Call lstErr_AddItem(lstErr, cmdContext, "Importation YSWIMON0_Synchro1"): DoEvents

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIMON0 , " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SAAUMIDL = 0  and SWISABNUM = SWISABZSWI and SWISABWES = 'S'" _
     & " order by SWIMONID"
Set rsSab = cnsab.Execute(xSql)
blnTransaction = True

Do While Not rsSab.EOF
    Call srvYSWIMON0_GetBuffer_ODBC(rsSab, oldYSWIMON0)
    newYSWIMON0 = oldYSWIMON0
    newYSWIMON0.SAAAID = rsSab("SWISABWID1")
    newYSWIMON0.SAAUMIDL = rsSab("SWISABWIDL")
    newYSWIMON0.SAAUMIDH = rsSab("SWISABWIDH")
    Nb = Nb + 1
    If Nb Mod 10 = 0 Then
        'Nb = 0
            Call lstErr_ChangeLastItem(lstErr, cmdContext, "YSWIMON0_Synchro1 : " & Nb): DoEvents
    End If
    V = sqlYSWIMON0_Update(newYSWIMON0, oldYSWIMON0, cnSab_Update)
    rsSab.MoveNext
Loop

'==================================================================
Call lstErr_ChangeLastItem(lstErr, cmdContext, "YSWIMON0_Synchro1 : " & Nb): DoEvents

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation YSWIMON0_Synchro1"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If


End Sub

Public Sub Importation_SAB_YGOSDOS0_Synchro1()
Dim V, X As String, K As Long, K2 As Long, Nb As Long
Dim xSql As String, blnTransaction As Boolean

On Error GoTo Error_Handler
'==================================================================
blnTransaction = False
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAB_YGOSDOS0_Synchro1"): DoEvents

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

'==================================================================
Call lstErr_ChangeLastItem(lstErr, cmdContext, "Importation_SAB_YGOSDOS0_Synchro1 : " & Nb): DoEvents
'==================================================================
Nb = 0
xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSDOS0 , " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where GOSDOSIDD > 0  and GOSDOSWID1 = SWISABWID1 and GOSDOSWIDL = SWISABWIDL and GOSDOSWIDH = SWISABWIDH" _
     & " and SWISABXGOS = '' order by GOSDOSIDD"
Set rsSab = cnsab.Execute(xSql)
blnTransaction = True

Do While Not rsSab.EOF
    
    Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
    newYSWISAB0 = oldYSWISAB0
    newYSWISAB0.SWISABXGOS = "G"
    Nb = Nb + 1
    If Nb Mod 10 = 0 Then
        'Nb = 0
            Call lstErr_ChangeLastItem(lstErr, cmdContext, "Importation_SAB_YGOSDOS0_Synchro1-2 : " & Nb): DoEvents
    End If
    V = sqlYSWISAB0_Update(newYSWISAB0, oldYSWISAB0)
    rsSab.MoveNext
Loop

'==================================================================
Call lstErr_ChangeLastItem(lstErr, cmdContext, "Importation_SAB_YGOSDOS0_Synchro1 : " & Nb): DoEvents
'==================================================================
Nb = 0
xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0 , " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where GOSEVEIDD > 0 and GOSEVEIDE > 1 and GOSEVESWID >  0 and GOSEVESWID = SWISABSWID" _
     & " and SWISABXEVE = '' order by GOSEVEIDD"
Set rsSab = cnsab.Execute(xSql)
blnTransaction = True

Do While Not rsSab.EOF
    Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
    newYSWISAB0 = oldYSWISAB0
    newYSWISAB0.SWISABXEVE = "G": newYSWISAB0.SWISABK999 = "G"
    Nb = Nb + 1
    If Nb Mod 10 = 0 Then
        'Nb = 0
            Call lstErr_ChangeLastItem(lstErr, cmdContext, "Importation_SAB_YGOSDOS0_Synchro1 : " & Nb): DoEvents
    End If
    V = sqlYSWISAB0_Update(newYSWISAB0, oldYSWISAB0)
    rsSab.MoveNext
Loop
'==================================================================
Call lstErr_ChangeLastItem(lstErr, cmdContext, "Importation_SAB_YGOSDOS0_Synchro1-3 : " & Nb): DoEvents
'==================================================================
Nb = 0

xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " left outer join " & paramIBM_Library_SABSPE & ".YSWILNK0 on SWILNKSWID = GOSEVESWID and  SWILNKAPPN = GOSEVEIDD and SWILNKAPPC = 'GOS'" _
     & " where GOSEVEIDD > 0  and GOSEVESWID > 0 and SWILNKSWID is Null " _
     & "  order by GOSEVEIDD"
Set rsSab = cnsab.Execute(xSql)
blnTransaction = True

Do While Not rsSab.EOF
    If IsNull(rsSab("SWILNKSWID")) Then
        newYSWILNK0.SWILNKSWID = rsSab("GOSEVESWID")
        newYSWILNK0.SWILNKAPPN = rsSab("GOSEVEIDD")
        newYSWILNK0.SWILNKAPPC = "GOS"
        newYSWILNK0.SWILNKSTA = ""


        V = sqlYSWILNK0_Insert(newYSWILNK0)
    End If
    rsSab.MoveNext
Loop
'==================================================================

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation_SAB_YGOSDOS0_Synchro1"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If


End Sub

Public Sub Importation_SAB_YGOSDOS0_Synchro2()
Dim V, X As String, K As Long, K2 As Long, Nb As Long
Dim xSql As String, blnTransaction As Boolean

On Error GoTo Error_Handler
'==================================================================
blnTransaction = False
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAB_YGOSDOS0_Synchro2"): DoEvents

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

'==================================================================
Nb = 0
xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where GOSEVENAT = 'Swi>' and GOSEVESWID = 0 and GOSEVESTAE = ' '" _
     & " order by GOSEVEIDD , GOSEVEIDE"

' $JPL 2015-09-13  and GOSEVESTAE = ' '

Set rsSab = cnsab.Execute(xSql)
blnTransaction = True

Do While Not rsSab.EOF
    
    X = Format(rsSab("GOSEVEIDD"), "000000000") & Format(rsSab("GOSEVEIDE"), "000000000")
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIADTAQ" _
        & " where substring(BIADTATXTE , 1 , 18) = '" & X & "'"
    Set rsSabX = cnsab.Execute(xSql)
    
    If Not rsSabX.EOF Then
        Nb = Val(Mid$(rsSabX("BIADTATXTS"), 11, 12))
        If Nb > 0 Then
            xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0" _
                & " where SWISABWES = 'S' and SWISABZSWI = " & Nb
            Set rsSabX = cnsab.Execute(xSql)
            
            If Not rsSabX.EOF Then
                X = "Set SWISABXEVE = 'G'"
                V = sqlYSWISAB0_Update_Field(rsSabX("SWISABSWID"), X)
                If Not IsNull(V) Then GoTo Error_MsgBox
                
                Call rsYGOSEVE0_GetBuffer(rsSab, oldYGOSEVE0)
                newYGOSEVE0 = oldYGOSEVE0
                newYGOSEVE0.GOSEVESWID = rsSabX("SWISABSWID")
                V = sqlYGOSEVE0_Update(newYGOSEVE0, oldYGOSEVE0, False)
                
                If Not IsNull(V) Then GoTo Error_MsgBox
                
                Nb = Nb + 1
                If Nb Mod 10 = 0 Then
                    'Nb = 0
                        Call lstErr_ChangeLastItem(lstErr, cmdContext, "Importation_SAB_YGOSDOS0_Synchro2-2 : " & Nb): DoEvents
                End If
            End If
        End If
    End If
    rsSab.MoveNext
Loop

'==================================================================
Call lstErr_ChangeLastItem(lstErr, cmdContext, "Importation_SAB_YGOSDOS0_Synchro2 : " & Nb): DoEvents
'==================================================================

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation_SAB_YGOSDOS0_Synchro2"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If


End Sub



Public Sub Importation_SAB_YSWIMON0_Synchro2()
Dim V, X As String, K As Long, K2 As Long, Nb As Long
Dim xSql As String, blnTransaction As Boolean


On Error GoTo Error_Handler
'==================================================================
blnTransaction = False
Call lstErr_AddItem(lstErr, cmdContext, "Importation SAB_YSWIMON0_Synchro2"): DoEvents

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIMON0 , " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where substring(SWIMONSTA , 1 , 2 ) <> 'S9'" _
      & " and SAAAID = SWISABWID1 and SAAUMIDL = SWISABWIDL and SAAUMIDH = SWISABWIDH" _
     & " order by SWIMONID"
Set rsSab = cnsab.Execute(xSql)
blnTransaction = True

Do While Not rsSab.EOF
    X = ""
    Select Case rsSab("SWISABWSTA")
        Case " ": X = "S300"
        Case "V": X = "S901"
        Case Else: X = "S903"
    End Select
    If X <> "" Then
        If X <> rsSab("SWIMONSTA") Then
            Call srvYSWIMON0_GetBuffer_ODBC(rsSab, oldYSWIMON0)
            newYSWIMON0 = oldYSWIMON0
            newYSWIMON0.SWIMONSTA = X
        
            Nb = Nb + 1
            If Nb Mod 10 = 0 Then
                'Nb = 0
                Call lstErr_ChangeLastItem(lstErr, cmdContext, "YSWIMON0_Synchro2 : " & Nb): DoEvents
            End If
            V = sqlYSWIMON0_Update(newYSWIMON0, oldYSWIMON0, cnSab_Update)
        End If
    End If
    
    rsSab.MoveNext
Loop

'==================================================================
Call lstErr_ChangeLastItem(lstErr, cmdContext, "YSWIMON0_Synchro2 : " & Nb): DoEvents

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation YSWIMON0_Synchro2"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If


End Sub
Public Sub Importation_SAB_YSWIMON0_Synchro3()
Dim V, X As String, K As Long, K2 As Long, Nb As Long
Dim xSql As String, blnTransaction As Boolean
Dim wSWIMONFLUXD_SSS As Double

On Error GoTo Error_Handler
'==================================================================
blnTransaction = False
Call lstErr_AddItem(lstErr, cmdContext, "Importation SAB_YSWIMON0_Synchro3"): DoEvents

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIMON0" _
     & " where SAAUMIDL = 0 and SWIMONSTA = 'S200'" _
     & " order by SWIMONID"
Set rsSab = cnsab.Execute(xSql)
blnTransaction = True

Do While Not rsSab.EOF
    wSWIMONFLUXD_SSS = rsSab("SWIMONFLUD") * 100000 + Time_Hms_Sss(Format(rsSab("SWIMONFLUH"), "000000")) + 300

    
    If SAA_Alerte_SWIHIADEN_SSS > wSWIMONFLUXD_SSS Then
        Call srvYSWIMON0_GetBuffer_ODBC(rsSab, oldYSWIMON0)
        newYSWIMON0 = oldYSWIMON0
        newYSWIMON0.SWIMONSTA = "S998"
    
        Nb = Nb + 1
        If Nb Mod 10 = 0 Then
            'Nb = 0
            Call lstErr_ChangeLastItem(lstErr, cmdContext, "YSWIMON0_Synchro3 : " & Nb): DoEvents
        End If
        V = sqlYSWIMON0_Update(newYSWIMON0, oldYSWIMON0, cnSab_Update)
    End If
    
    rsSab.MoveNext
Loop

'==================================================================
Call lstErr_ChangeLastItem(lstErr, cmdContext, "YSWIMON0_Synchro3 : " & Nb): DoEvents

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation YSWIMON0_Synchro3"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If


End Sub


Public Sub Importation_SAB_ZSWIHIA0()
Dim V, X As String, K As Long, K2 As Long, Nb_Lu As Long, Nb_Equal As Long, Nb_ok As Long
Dim xSql As String, blnTransaction As Boolean
Dim wAmj As Long, blnN20 As Boolean, arrSWISABWAMJ_SSS() As Double, Nb As Long
Dim xZSWIHIA0 As typeZSWIHIA0, wSWIHIAHEN As Long
Dim newSAA_Alerte_SWIHIADEN As Long, newSAA_Alerte_SWIHIADEN_SSS As Double
Dim wSWIHIADEN_SSS As Double, xRef As String
Dim blnSAA_Alerte_New As Boolean
Dim arrSWIHIANUM() As Long, arrSWIHIADEN_SSS() As Double, arrSWIHIANUM_Nb As Integer, arrSWIHIANUM_Max As Integer
Dim mK_Match As Long, blnK_Match As Boolean
On Error GoTo Error_Handler

blnTransaction = False
currentAction = "Importation_SAB_ZSWIHIA0"

If lastSWIHIADEN = 0 Then
    xSql = "select * from  " & paramIBM_Library_SAB & ".zswihia0 " _
         & " WHERE swihianum  in ( select swisabzswi from " & paramIBM_Library_SABSPE & ".YSWISAB0" _
         & " where SWISABWES = 'S')" _
         & " and swihiaval = 'O'" _
         & " ORDER BY swihiaden desc "
    Set rsSab = cnsab.Execute(xSql)
    If Not rsSab.EOF Then lastSWIHIADEN = rsSab("SWIHIADEN")
End If

If SAA_Alerte_SWIHIADEN = 0 Then
        New_YBIATAB0.BIATABID = "SAA_Alerte"
        New_YBIATAB0.BIATABK1 = "ZSWIHIA0"
        New_YBIATAB0.BIATABK2 = ""
        New_YBIATAB0.BIATABTXT = ""
        
        Call lstErr_AddItem(lstErr, cmdContext, currentAction): DoEvents
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
             & " where BIATABID = '" & New_YBIATAB0.BIATABID & "' and BIATABK1 = '" & New_YBIATAB0.BIATABK1 & "'"
        Set rsSab = cnsab.Execute(xSql)
        If rsSab.EOF Then
            New_YBIATAB0.BIATABTXT = "20120301 2012030100000"
            Parametrage_New

        Else
            X = rsSab("BIATABTXT")
            SAA_Alerte_SWIHIADEN = Val(Mid$(X, 1, 8))
            SAA_Alerte_SWIHIADEN_SSS = Val(Mid$(X, 10, 13))
        End If
End If



newSAA_Alerte_SWIHIADEN = SAA_Alerte_SWIHIADEN
newSAA_Alerte_SWIHIADEN_SSS = SAA_Alerte_SWIHIADEN_SSS
blnSAA_Alerte_New = False

'==================================================================

Call lstErr_AddItem(lstErr, cmdContext, "Importation YSWISAB0"): DoEvents

'$JPL 2012-08-27_________________________________________________________________________
'xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
'     & " where SWISABZSWI = 0 and SWISABWES = 'S'" _
'     & " and SWISABWAMJ >= " & SAA_Alerte_SWIHIADEN
xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABZSWI = 0 and SWISABWES = 'S'"
'$JPL 2012-08-27_________________________________________________________________________
     
     
Set rsSab = cnsab.Execute(xSql)
ReDim arrYSWISAB0(rsSab(0) + 100), arrSWISABWAMJ_SSS(rsSab(0) + 100)

arrYSWISAB0_Nb = 0
'$JPL 2012-08-27_________________________________________________________________________
'xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
'     & " where SWISABZSWI = 0 and SWISABWES = 'S'" _
'     & " and SWISABWAMJ >= " & SAA_Alerte_SWIHIADEN & " order by SWISABSWID"
xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABZSWI = 0 and SWISABWES = 'S'" _
     & " order by SWISABSWID"
'$JPL 2012-08-27_________________________________________________________________________
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    arrYSWISAB0_Nb = arrYSWISAB0_Nb + 1
    Call rsYSWISAB0_GetBuffer(rsSab, arrYSWISAB0(arrYSWISAB0_Nb))
    X = Format(arrYSWISAB0(arrYSWISAB0_Nb).SWISABWHMS, "000000")
     arrSWISABWAMJ_SSS(arrYSWISAB0_Nb) = Val(arrYSWISAB0(arrYSWISAB0_Nb).SWISABWAMJ) * 100000 + Time_Hms_Sss(X)
    rsSab.MoveNext
Loop

'==================================================================


V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
blnTransaction = True




'xSQL = "select * from  JPLTST.zswihia0 " _
'$JPL 20120524
'$JPL 2012-08-27_________________________________________________________________________
'xSql = "select * from  " & paramIBM_Library_SAB & ".zswihia0 " _
'     & " WHERE swihianum not in ( select swisabzswi from " & paramIBM_Library_SABSPE & ".YSWISAB0" _
'     & " where SWISABWES = 'S' and SWISABWAMJ >= " & SAA_Alerte_SWIHIADEN & ")" _
'     & " and swihiaval = 'O' and swihiaden >= " & SAA_Alerte_SWIHIADEN - 19000000 _
'     & " ORDER BY swihiaden , SWIHIAHEN "
xSql = "select * from  " & paramIBM_Library_SAB & ".zswihia0 " _
     & " WHERE swihianum not in ( select swisabzswi from " & paramIBM_Library_SABSPE & ".YSWISAB0" _
     & " where SWISABWES = 'S')" _
     & " and swihiaval = 'O' and SWIHIADEN >= " & lastSWIHIADEN _
     & " ORDER BY swihiaden , SWIHIAHEN "
     
     
'$JPL 2012-08-27_________________________________________________________________________
     
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    Nb_Lu = Nb_Lu + 1
   ' If rsSab("SWIHIANUM") = 477231 Then
   '     Debug.Print
   ' End If
    xZSWIHIA0.SWIHIAREF = Trim(rsSab("SWIHIAREF"))
    xZSWIHIA0.SWIHIAMES = Trim(rsSab("SWIHIAMES"))
    xZSWIHIA0.SWIHIADES = Trim(rsSab("SWIHIADES"))
    xZSWIHIA0.SWIHIADEN = rsSab("SWIHIADEN")
    xZSWIHIA0.SWIHIAHEN = rsSab("SWIHIAhEN")
    
    blnN20 = False
    X = Trim(rsSab("SWIHIAREF"))
    Select Case Mid$(X, 1, 5)
        Case "00TRF":
                    Select Case rsSab("SWIHIASSE")
                        Case "CD": X = Replace(X, "00TRF", "SOBITRF"): blnN20 = True
                        Case Else: X = Replace(X, "00TRF", "ORPATRF"): blnN20 = True
                    End Select '$DR 2014-10-22
        Case "00CPT": X = Replace(X, "00CPT", "ORPACPT"): blnN20 = True
        Case "TCTRF": X = Replace(X, "TCTRF", "BOTCTRF"): blnN20 = True
        Case "TCCPT": X = Replace(X, "TCCPT", "BOTCCPT"): blnN20 = True
        Case "TCSWP": X = Replace(X, "TCSWP", "BOTCSWP"): blnN20 = True
        Case "00CDE": X = Replace(X, "00CDE", "SOBICDE"): blnN20 = True
        Case "00CDI": X = Replace(X, "00CDI", "SOBICDI"): blnN20 = True
        Case "00RDE": X = Replace(X, "00RDE", "SOBIRDE"): blnN20 = True
        Case "00RDI": X = Replace(X, "00RDI", "SOBIRDI"): blnN20 = True
    End Select
    If Not blnN20 Then
           Select Case Mid$(X, 1, 3)
               Case "EMP": X = Replace(X, "EMP", "BOTCEMP"): blnN20 = True
               Case "PRE": X = Replace(X, "PRE", "BOTCPRE"): blnN20 = True
               Case "CDE": X = Replace(X, "CDE", "SOBICDE"): blnN20 = True
               Case "CDI": X = Replace(X, "CDI", "SOBICDI"): blnN20 = True
               Case "CPT":
                    Select Case rsSab("SWIHIASSE")
                        Case "TR": X = Replace(X, "CPT", "ORPACPT"): blnN20 = True
                        Case "TC": X = Replace(X, "CPT", "BOTCCPT"): blnN20 = True
                    End Select
           End Select
    End If
    
   xZSWIHIA0.SWIHIAREF = X
    
   wSWIHIADEN_SSS = Val((19000000 + xZSWIHIA0.SWIHIADEN)) * 100000 + Time_Hms_Sss(Format(xZSWIHIA0.SWIHIAHEN, "000000"))
    
    Nb_Equal = 0: mK_Match = 0: blnK_Match = False
    For K = 1 To arrYSWISAB0_Nb
    

        If arrYSWISAB0(K).SWISABZSWI = 0 Then
            If xZSWIHIA0.SWIHIAMES = arrYSWISAB0(K).SWISABWMTK _
            And Mid$(xZSWIHIA0.SWIHIADES, 1, 8) = Mid$(arrYSWISAB0(K).SWISABWBIC, 1, 8) _
            And Trim(xZSWIHIA0.SWIHIAREF) = Trim(arrYSWISAB0(K).SWISABWN20) Then
            
            '$JPL 20120524
                If arrYSWISAB0(K).SWISABWMTD = 0 Then
                    'If Nb_Equal = 0 Then
                        mK_Match = K
                        Nb_Equal = Nb_Equal + 1
                        Exit For
                    'End If
                Else
                
                'If Nb_Equal = 0 Then
                        If rsSab("SWIHIADE1") = arrYSWISAB0(K).SWISABWDEV _
                        And Abs(rsSab("SWIHIAMON") - arrYSWISAB0(K).SWISABWMTD) < 1 Then
                            blnK_Match = True: mK_Match = K: Nb_Equal = Nb_Equal + 1
                            Exit For
                        Else
                            'If rsSab("SWIHIADE1") = arrYSWISAB0(K).SWISABWDEV Then
                                If xZSWIHIA0.SWIHIAMES > "190" Then
                                    blnK_Match = True: mK_Match = K: Nb_Equal = Nb_Equal + 1
                                    Exit For
                                End If
                            'End If
                        End If
                        
                End If
                
            End If
        End If
    Next K

    If Nb_Equal > 0 Then
        arrYSWISAB0(mK_Match).SWISABZSWI = rsSab("SWIHIANUM")

                'If Not blnSAA_Alerte_New Then
                '    blnSAA_Alerte_New = True
                    newSAA_Alerte_SWIHIADEN = 19000000 + xZSWIHIA0.SWIHIADEN
                    newSAA_Alerte_SWIHIADEN_SSS = Val(newSAA_Alerte_SWIHIADEN) * 100000 + Time_Hms_Sss(Format(xZSWIHIA0.SWIHIAHEN, "000000"))
                'End If

        Nb_ok = Nb_ok + 1
        If Nb_ok Mod 10 Then
            'Nb = 0
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "ZSWIHIA0 : " & oldYSWISAB0.SWISABSWID & " " & newYSWISAB0.SWISABZSWI): DoEvents
        End If
    Else
        ''If wSWIHIADEN_SSS < newSAA_Alerte_SWIHIADEN_SSS - 600 Then
            If arrSWIHIANUM_Nb >= arrSWIHIANUM_Max Then
                ReDim Preserve arrSWIHIANUM(arrSWIHIANUM_Max + 10), arrSWIHIADEN_SSS(arrSWIHIANUM_Max + 10)
                arrSWIHIANUM_Max = arrSWIHIANUM_Max + 10
            End If
            arrSWIHIANUM_Nb = arrSWIHIANUM_Nb + 1
            arrSWIHIANUM(arrSWIHIANUM_Nb) = rsSab("SWIHIANUM")
            arrSWIHIADEN_SSS(arrSWIHIANUM_Nb) = wSWIHIADEN_SSS
            '
           ' Debug.Print rsSab("SWIHIANUM"), Nb_Equal, xZSWIHIA0.SWIHIAREF, xZSWIHIA0.SWIHIAMES, Trim(rsSab("SWIHIADES")), Trim(rsSab("SWIHIADEN")), Trim(rsSab("SWIHIAHEN"))
        ''End If
    End If
    
    rsSab.MoveNext
Loop

'==================================================================
'Debug.Print Nb_ok; " / "; Nb_Lu

'GoTo Exit_sub

For K = 1 To arrYSWISAB0_Nb

    If arrYSWISAB0(K).SWISABZSWI > 0 Then
         xSql = "update " & paramIBM_Library_SABSPE & ".YSWISAB0" _
              & " set SWISABZSWI = " & arrYSWISAB0(K).SWISABZSWI _
              & " where SWISABSWID = " & arrYSWISAB0(K).SWISABSWID
        Call FEU_ROUGE
        Set rsSab = cnSab_Update.Execute(xSql, Nb)
        Call FEU_VERT
        ' Tester si la mise à jour a été effectuée
        '===================================================================================
        
        If Nb = 0 Then
            If Not blnAuto Then MsgBox "Erreur màj : " & arrYSWISAB0(K).SWISABSWID, vbCritical, Me.Name & " : Importation_SAB_ZSWIHIA0"
        End If
    Else
    
    
    End If
Next K


'==================================================================
        Old_YBIATAB0.BIATABID = "SAA_Alerte"
        Old_YBIATAB0.BIATABK1 = "ZSWIHIA0"
        Old_YBIATAB0.BIATABK2 = ""
        Old_YBIATAB0.BIATABTXT = SAA_Alerte_SWIHIADEN & " " & SAA_Alerte_SWIHIADEN_SSS

        New_YBIATAB0 = Old_YBIATAB0
        New_YBIATAB0.BIATABTXT = newSAA_Alerte_SWIHIADEN & " " & newSAA_Alerte_SWIHIADEN_SSS
        If Old_YBIATAB0.BIATABTXT <> New_YBIATAB0.BIATABTXT Then
            V = sqlYBIATAB0_Update(New_YBIATAB0, Old_YBIATAB0)
        End If

Call lstErr_ChangeLastItem(lstErr, cmdContext, "Importation_SAB_ZSWIHIA0 : " & Nb): DoEvents

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        If arrSWIHIANUM_Nb > 0 Then
            For K = 1 To arrSWIHIANUM_Nb
                '$JPL 20120524
                If arrSWIHIADEN_SSS(K) >= SAA_Alerte_SWIHIADEN_SSS And arrSWIHIADEN_SSS(K) < newSAA_Alerte_SWIHIADEN_SSS Then
                    cmdSendMail_ZSWIHIA0 (arrSWIHIANUM(K))
                End If
            Next K
        End If
        SAA_Alerte_SWIHIADEN = newSAA_Alerte_SWIHIADEN
        SAA_Alerte_SWIHIADEN_SSS = newSAA_Alerte_SWIHIADEN_SSS

    End If
End If


End Sub
Public Sub Importation_SAA_Origine_MT()
Dim V, X As String, Nb As Long, Nb_Read As Long, Nb_Update As Long
Dim xSql As String, blnTransaction As Boolean
Dim blnSAA_Alerte_New As Boolean, newSAA_Alerte_Origine_MT As Long
On Error GoTo Error_Handler

blnTransaction = False
currentAction = "Importation_SAA_Origine_MT"

'==================================================================

Call lstErr_AddItem(lstErr, cmdContext, currentAction & " - 1"): DoEvents

'==================================================================
If SAA_Alerte_Origine_MT = 0 Then
        New_YBIATAB0.BIATABID = "SAA_Alerte"
        New_YBIATAB0.BIATABK1 = "Origine_MT"
        New_YBIATAB0.BIATABK2 = ""
        New_YBIATAB0.BIATABTXT = ""
        
        Call lstErr_AddItem(lstErr, cmdContext, currentAction): DoEvents
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
             & " where BIATABID = '" & New_YBIATAB0.BIATABID & "' and BIATABK1 = '" & New_YBIATAB0.BIATABK1 & "'"
        Set rsSab = cnsab.Execute(xSql)
        If rsSab.EOF Then
            New_YBIATAB0.BIATABTXT = "000915700"
            Parametrage_New

        Else
            X = rsSab("BIATABTXT")
            SAA_Alerte_Origine_MT = Val(Mid$(X, 1, 9))
        End If
End If
newSAA_Alerte_Origine_MT = SAA_Alerte_Origine_MT
blnSAA_Alerte_New = False
'==================================================================

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
blnTransaction = True

xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABZSWI = 0 and SWISABWES = 'S'" _
     & " and SWISABWMTK = '950'"
    
Set rsSab = cnsab.Execute(xSql)
If rsSab(0) > 0 Then

    xSql = "update " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
         & "set SWISABZSWI = -1 " _
         & " where SWISABZSWI = 0 and SWISABWES = 'S'" _
         & " and SWISABWMTK = '950'"
     Call FEU_ROUGE
    Set rsSab = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    If Nb = 0 Then
        If Not blnAuto Then MsgBox "Erreur màj : ", vbCritical, Me.Name & currentAction & " - 2"
    End If
End If


xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABZSWI = 0 and SWISABWES = 'S'" _
     & " order by SWISABSWID"
     
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    Nb_Read = Nb_Read + 1
    xSql = "select * from rMesg " _
        & "where Aid = " & rsSab("SWISABWID1") _
        & " and Mesg_s_umidl = " & rsSab("SWISABWIDL") _
        & " and Mesg_s_umidh  =  " & rsSab("SWISABWIDH")
   Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
   
    If Not rsSIDE_DB.EOF Then
         If Not IsNull(rsSIDE_DB("mesg_crea_rp_name")) Then
            If Trim(rsSIDE_DB("mesg_crea_rp_name")) = "_MP_creation" Then
                Nb_Update = Nb_Update + 1
                Call sqlYSWISAB0_Update_Field(rsSab("SWISABSWID"), "set SWISABZSWI = -2 ")
            End If
         End If
    End If
    If Nb_Read Mod 1000 = 0 Then Call lstErr_ChangeLastItem(lstErr, cmdContext, currentAction & " Update " & Nb_Update & "/ Read : " & Nb_Read): DoEvents

    rsSab.MoveNext
Loop

cnSAB_Transaction ("Commit")
blnTransaction = False

Call lstErr_ChangeLastItem(lstErr, cmdContext, currentAction & " - 9"): DoEvents
'==================================================================

Call lstErr_AddItem(lstErr, cmdContext, currentAction & " - 11"): DoEvents
'_____________________________________________________________________________________
xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABSWID > " & SAA_Alerte_Origine_MT _
     & " and SWISABZSWI = 0 and SWISABWES = 'S'" _
     & " order by SWISABSWID"
     
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    xSql = "select * from rMesg " _
        & "where Aid = " & rsSab("SWISABWID1") _
        & " and Mesg_s_umidl = " & rsSab("SWISABWIDL") _
        & " and Mesg_s_umidh  =  " & rsSab("SWISABWIDH")
   Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
   
    If Not rsSIDE_DB.EOF Then
    
        newSAA_Alerte_Origine_MT = rsSab("SWISABSWID")
        
        Call srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
        
        X = "Vérifier l'origine de ce message SWIFT sortant dans SAA : <BR>"
'DR 11/12/2018. Suppression (cf CDE 101683 MT730 du 24/07/2012)
'        xSql = X & "<BR> - le message a peut-être été importé 2 fois (origine SAB)" _
                 & "<BR> - le programme de rapprochement SAB => SAA n'a pas pu identifier le message" _
                 & "<BR> - (mauvaise gestion par SAB : validation / modification concomitantes (cf CDE 101683 MT730 du 24/07/2012) " _
                 & "<BR> - le message provient d'une source inconnue" _
                 & "<BR>"
        xSql = X & "<BR> - le message a peut-être été importé 2 fois (origine SAB)" _
                 & "<BR> - le programme de rapprochement SAB => SAA n'a pas pu identifier le message" _
                 & "<BR> - (mauvaise gestion par SAB : validation / modification concomitantes " _
                 & "<BR> - le message provient d'une source inconnue" _
                 & "<BR>"
        Call cmdSendMail_SAA_Alerte_rMesg("SAA_Origine_MT", X, xSql, xrMesg.x_inst0_unit_name, "")

    End If

    rsSab.MoveNext
Loop

'____________________________________________________________________________________
If newSAA_Alerte_Origine_MT <> SAA_Alerte_Origine_MT Then
        Old_YBIATAB0.BIATABID = "SAA_Alerte"
        Old_YBIATAB0.BIATABK1 = "Origine_MT"
        Old_YBIATAB0.BIATABK2 = ""
        Old_YBIATAB0.BIATABTXT = Format$(SAA_Alerte_Origine_MT, "000000000")

        New_YBIATAB0 = Old_YBIATAB0
        New_YBIATAB0.BIATABTXT = Format$(newSAA_Alerte_Origine_MT, "000000000")
        If Old_YBIATAB0.BIATABTXT <> New_YBIATAB0.BIATABTXT Then
            Parametrage_Update
            'V = sqlYBIATAB0_Update(New_YBIATAB0, Old_YBIATAB0)
        End If
        SAA_Alerte_Origine_MT = newSAA_Alerte_Origine_MT
End If
'==================================================================

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & currentAction
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If


End Sub

Public Sub Importation_SAA_Modification_MT()
Dim V, X As String, Nb As Long, Nb_Read As Long, Nb_Update As Long
Dim xSql As String, blnTransaction As Boolean
Dim newSAA_Alerte_Modification_MT As String
Dim xUUMID As String
On Error GoTo Error_Handler

blnTransaction = False
currentAction = "Importation_SAA_Modification_MT"

'==================================================================

Call lstErr_AddItem(lstErr, cmdContext, currentAction & " - 1"): DoEvents

'==================================================================
If SAA_Alerte_Modification_MT = "" Then
        New_YBIATAB0.BIATABID = "SAA_Alerte"
        New_YBIATAB0.BIATABK1 = "Modification_MT"
        New_YBIATAB0.BIATABK2 = ""
        New_YBIATAB0.BIATABTXT = ""
        
        Call lstErr_AddItem(lstErr, cmdContext, currentAction): DoEvents
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
             & " where BIATABID = '" & New_YBIATAB0.BIATABID & "' and BIATABK1 = '" & New_YBIATAB0.BIATABK1 & "'"
        Set rsSab = cnsab.Execute(xSql)
        If rsSab.EOF Then
            SAA_Alerte_Modification_MT = "01/04/2011 00:00:00.000"
            New_YBIATAB0.BIATABTXT = SAA_Alerte_Modification_MT
            Parametrage_New

        Else
            SAA_Alerte_Modification_MT = Trim(rsSab("BIATABTXT"))
        End If
End If
newSAA_Alerte_Modification_MT = SAA_Alerte_Modification_MT
'==================================================================

 xSql = "select * from rMesg " _
           & " where Mesg_crea_date_time <> Mesg_mod_date_time and mesg_sub_format = 'INPUT'" _
           & " and Mesg_mod_date_time >= '" & SAA_Alerte_Modification_MT & "'" _
           & " order by Mesg_mod_date_time"
           
'           & " and Mesg_mod_date_time >= {ts '" & SAA_Alerte_Modification_MT & "'}" _

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
  
Do While Not rsSIDE_DB.EOF

    Call srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
    newSAA_Alerte_Modification_MT = xrMesg.mesg_mod_date_time
    
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
         & " where SWISABWID1 = " & xrMesg.Aid _
         & " and   SWISABWIDL = " & xrMesg.mesg_s_umidl _
         & " and   SWISABWIDH = " & xrMesg.mesg_s_umidh _
         
    Set rsSab = cnsab.Execute(xSql)
    
    If Not rsSab.EOF Then
        Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
        newYSWISAB0 = oldYSWISAB0
        
        newYSWISAB0.SWISABWMTK = xrMesg.mesg_type
        xUUMID = xrMesg.mesg_uumid
        If Mid$(xUUMID, 1, 1) = "I" Then
            newYSWISAB0.SWISABWES = "S"
            newYSWISAB0.SWISABWN20 = Trim(Replace(xrMesg.mesg_trn_ref, "'", " "))
            newYSWISAB0.SWISABWL20 = Trim(Replace(xrMesg.mesg_rel_trn_ref, "'", " "))
            If newYSWISAB0.SWISABWMTK = "950" Then newYSWISAB0.SWISABZSWI = -1
            If xrMesg.mesg_crea_rp_name = "_MP_creation" Then newYSWISAB0.SWISABZSWI = -2
        Else
            newYSWISAB0.SWISABWES = "E"
            newYSWISAB0.SWISABWL20 = Trim(Replace(xrMesg.mesg_trn_ref, "'", " "))
            newYSWISAB0.SWISABWN20 = Trim(Replace(xrMesg.mesg_rel_trn_ref, "'", " "))
        End If
        
        newYSWISAB0.SWISABWBIC = Mid$(xUUMID, 2, 11)
        If Not IsNull(xrMesg.x_fin_ccy) Then
            newYSWISAB0.SWISABWDEV = xrMesg.x_fin_ccy
            newYSWISAB0.SWISABWMTD = CCur(xrMesg.x_fin_amount)
        End If
        newYSWISAB0.SWISABWSRV = xrMesg.x_inst0_unit_name
        If newYSWISAB0.SWISABWSRV = "None" Then newYSWISAB0.SWISABWSRV = ""
        
        If newYSWISAB0.SWISABWES <> oldYSWISAB0.SWISABWES _
        Or newYSWISAB0.SWISABWMTK <> oldYSWISAB0.SWISABWMTK _
        Or newYSWISAB0.SWISABWBIC <> oldYSWISAB0.SWISABWBIC _
        Or newYSWISAB0.SWISABWDEV <> Trim(oldYSWISAB0.SWISABWDEV) _
        Or newYSWISAB0.SWISABWMTD <> oldYSWISAB0.SWISABWMTD _
        Or newYSWISAB0.SWISABWN20 <> oldYSWISAB0.SWISABWN20 _
        Or newYSWISAB0.SWISABWL20 <> oldYSWISAB0.SWISABWL20 _
        Or newYSWISAB0.SWISABWSRV <> Trim(oldYSWISAB0.SWISABWSRV) Then
        
            If Not blnTransaction Then
                V = cnSAB_Transaction("BeginTrans")
                If Not IsNull(V) Then GoTo Error_MsgBox
                blnTransaction = True
            End If
            If newYSWISAB0.SWISABWN20 <> oldYSWISAB0.SWISABWN20 Then
                If Trim(xrMesg.mesg_crea_rp_name) = "_MP_creation" Then
                    'Debug.Print oldYSWISAB0.SWISABSWID & " N20 : " & newYSWISAB0.SWISABWN20 & " | " & newYSWISAB0.SWISABOPEN
                    newYSWISAB0.SWISABOPEC = ""
                    newYSWISAB0.SWISABOPEN = 0
                End If
            End If
            
            If newYSWISAB0.SWISABWSRV <> Trim(oldYSWISAB0.SWISABWSRV) Then
                Select Case newYSWISAB0.SWISABWSRV
                    Case "SOBF", "ORPA", "GDMP": newYSWISAB0.SWISABKSRV = "S01"
                    Case "SOBI": newYSWISAB0.SWISABKSRV = "S10"
                    Case "DAFI", "BOTC": newYSWISAB0.SWISABKSRV = "S32"
                    Case "DCOM": newYSWISAB0.SWISABKSRV = "S41"
                    '''''Case Else: xYSWISAB0.SWISABKSRV = "S00"
                End Select
            End If
            V = sqlYSWISAB0_Update(newYSWISAB0, oldYSWISAB0)
            If Not IsNull(V) Then GoTo Error_MsgBox
        End If
        
    End If

    
    rsSIDE_DB.MoveNext
Loop


If blnTransaction Then
    cnSAB_Transaction ("Commit")
    blnTransaction = False
End If

Call lstErr_ChangeLastItem(lstErr, cmdContext, currentAction & " - 9"): DoEvents
'==================================================================

'_____________________________________________________________________________________
If newSAA_Alerte_Modification_MT <> SAA_Alerte_Modification_MT Then
        Old_YBIATAB0.BIATABID = "SAA_Alerte"
        Old_YBIATAB0.BIATABK1 = "Modification_MT"
        Old_YBIATAB0.BIATABK2 = ""
        Old_YBIATAB0.BIATABTXT = SAA_Alerte_Modification_MT

        New_YBIATAB0 = Old_YBIATAB0
        New_YBIATAB0.BIATABTXT = newSAA_Alerte_Modification_MT
        If Old_YBIATAB0.BIATABTXT <> New_YBIATAB0.BIATABTXT Then
            Parametrage_Update
            'V = sqlYBIATAB0_Update(New_YBIATAB0, Old_YBIATAB0)
        End If
        
        SAA_Alerte_Modification_MT = newSAA_Alerte_Modification_MT

End If
'==================================================================

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & currentAction
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If


End Sub


Public Sub cmdSelect_JPL()
Dim V, X As String, K As Long, K2 As Long, Nb_Lu As Long, Nb_Equal As Long, Nb_ok As Long, X2 As String
Dim xSql As String, blnTransaction As Boolean

On Error GoTo Error_Handler
Dim xOPEN As String, wOPEN As Double, blnOpen As Boolean, blnOPEC As Boolean
Dim blnOk As Boolean

Dim wSWISABWSTA As String, newSWISABWSTA As String

'===============================================================================================
Exit Sub
GoTo SWISABWSTA_Update
'===============================================================================================

Dim wK115 As String

Open "c:/temp/SWISABWSTA.txt" For Output As #3

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABSWID > 1000000  and SWISABWMTk in (198 , 298) order by SWISABSWID"

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    X2 = Importation_SAA_198(rsSab("SWISABWID1"), rsSab("SWISABWIDL"), rsSab("SWISABWIDH"), wK115)
    If wK115 <> "N" Then
        newSWISABWSTA = wK115
        Print #3, rsSab("SWISABSWID") & " " & newSWISABWSTA & "-" & rsSab("SWISABWSTA") & " " & wK115; rsSab("SWISABOPEC"); rsSab("SWISABOPEN"); rsSab("SWISABWMTK"); rsSab("SWISABWAMJ")
        Debug.Print rsSab("SWISABSWID") & " " & wK115 & "-" & wSWISABWSTA; rsSab("SWISABOPEC"); rsSab("SWISABOPEN"); rsSab("SWISABWMTK"); rsSab("SWISABWAMJ")
       ' If  wK115 <> "J" Then
      '  xSql = "select * from rMesg " _
       '      & "where Aid = " & rsSab("SWISABWID1") _
       '      & " and Mesg_s_umidl = " & rsSab("SWISABWIDL") _
       '      & " and Mesg_s_umidh  =  " & rsSab("SWISABWIDH")
       ' Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
        
       '  If Not rsSIDE_DB.EOF Then
                
       '      Call srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
             
       '      X = "Message SWIFT " & xrMesg.mesg_type & " _ champ 115 : " & wK115 & " (" & X2 & ")"
       '      Call cmdSendMail_SAA_Alerte_rMesg("SAA_" & xrMesg.mesg_type, X, X, xrMesg.x_inst0_unit_name, "")
            
       ' End If
    End If
    
    rsSab.MoveNext

Loop
Close #3
Exit Sub
'===============================================================================================


Open "c:/temp/SWISABWSTA.txt" For Output As #3


xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABSWID > 0  and SWISABSWID < 700000  and SWISABWES = 'S' order by SWISABSWID"

Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    wSWISABWSTA = rsSab("SWISABWSTA")
    'xSql = "select * from rMesg , rAppe " _
    '    & " where rMesg.aid = " & rsSab("SWISABWID1") _
    '    & " and mesg_s_umidl = " & rsSab("SWISABWIDL") & " and mesg_s_umidh = " & rsSab("SWISABWIDH") _
    '& " and rMesg.aid = rAppe.Aid and  mesg_s_umidl = appe_s_umidl and  mesg_s_umidh = appe_s_umidh" _
    '& " order by appe_date_time , appe_seq_nbr"
    xSql = "select * from rAppe " _
        & " where rAppe.aid = " & rsSab("SWISABWID1") _
        & " and Appe_s_umidl = " & rsSab("SWISABWIDL") & " and Appe_s_umidh = " & rsSab("SWISABWIDH") _
    & " and appe_inst_num = 0" _
    & " order by appe_date_time , appe_seq_nbr"

    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    
    newSWISABWSTA = ""
    Do While Not rsSIDE_DB.EOF
        'Call Importation_SAA_SWISABWSTA_Control(wSWISABWSTA, "rAppe")
                Select Case Trim(rsSIDE_DB("appe_network_delivery_status"))
                    Case "DLV_ACKED":
                        If Trim(rsSIDE_DB("appe_crea_rp_name")) = "_SI_to_SWIFT" Then
                            newSWISABWSTA = "V"
                        Else
                            newSWISABWSTA = "E"
                        End If
                        
                    Case Else: newSWISABWSTA = "E"
                End Select
        rsSIDE_DB.MoveNext
    Loop
        
    If wSWISABWSTA <> newSWISABWSTA And newSWISABWSTA <> "" Then
        
        Debug.Print rsSab("SWISABSWID") & " " & newSWISABWSTA & "-" & wSWISABWSTA; rsSab("SWISABOPEC"); rsSab("SWISABOPEN"); rsSab("SWISABWMTK"); rsSab("SWISABWAMJ")
        Print #3, rsSab("SWISABSWID") & " " & newSWISABWSTA & "-" & wSWISABWSTA; rsSab("SWISABOPEC"); rsSab("SWISABOPEN"); rsSab("SWISABWMTK"); rsSab("SWISABWAMJ")
    End If
    rsSab.MoveNext

Loop

Close #3


Exit Sub
'===============================================================================================

SWISABWSTA_Update:
'________________________________________________________________________
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'------------------------------------------
blnTransaction = True

Open "c:/temp/SWISABWSTA.txt" For Input As #3

Do Until EOF(3)
    Line Input #3, X
    K2 = InStr(X, " ")
    K = Val(Mid$(X, 1, K2))
    X2 = Mid$(X, K2 + 1, 1)
    If X2 = "V" Then
        V = sqlYSWISAB0_Update_Field(K, "set SWISABWSTA = 'V'")
    Else
        V = sqlYSWISAB0_Update_Field(K, "set SWISABWSTA = '" & X2 & "'")
   End If
    
    If Not IsNull(V) Then
        Debug.Print X; V
    End If
Loop
    
Close #3


GoTo Exit_sub

'===============================================================================================
Dim cnJPL As New ADODB.Connection, rsJPL As New ADODB.Recordset
paramODBC_DSN_SIDE_DB = "DSN=SQL2010_BIA" & ";UID=SIDE_READ" & "; PWD=" ' & xMemo

    xSql = "SELECT    syscolumns.name From sysobjects, syscolumns " _
          & " WHERE  ( sysobjects.id = syscolumns.id) And  (sysobjects.xtype = 'U') " _
          & " AND sysobjects.name LIKE 'rCorr' ORDER BY syscolumns.colorder"
   ' xSql = "select * from sys.tables"
    'xSql = "select * from rCorr where corr_X1 = 'SOGEFRPPXXX' and corr_mod_date_time > '06/01/2015 00:00:00'"
     'xSql = "select * from rCorr where corr_BIC_can_be_updated = 0 "
    xSql = "select * from rCorr where corr_X1 like 'BYLAGB2L%' "
    
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

    Do While Not rsSIDE_DB.EOF
        For K = 0 To 28
            Debug.Print rsSIDE_DB(K) & " | "; '; rsSIDE_DB(1); rsSIDE_DB(2); rsSIDE_DB(3); rsSIDE_DB(4); rsSIDE_DB(5)
        Next K
        Debug.Print
        rsSIDE_DB.MoveNext
    Loop

Exit Sub
'===============================================================================================
'corr_type
'corr_X1
'corr_X2
'corr_X3
'corr_X4
'corr_nature
'corr_BIC_can_be_updated
'corr_inheritance
'corr_language
'corr_information
'corr_institution_name
'corr_branch_info
'corr_location
'corr_city_name
'corr_physical_address
'corr_ctry_code
'corr_ctry_name
'corr_subtype
'corr_pob_number
'corr_pob_location
'corr_pob_ctry_code
'corr_pob_ctry_name
'corr_status
'corr_crea_oper_nickname
'corr_crea_date_time
'corr_mod_oper_nickname
'corr_mod_date_time
'corr_token
'corr_data_last
'===============================================================================================
'Dim cnJPL As New ADODB.Connection, rsJPL As New ADODB.Recordset
'paramODBC_DSN_SIDE_DB = "DSN=SQL2010_BIA" & ";UID=SIDE_READ" & "; PWD=" & xMemo
cnJPL.Open "DSN=JPL"

    'xSql = "SELECT    syscolumns.name From sysobjects, syscolumns " _
    '      & " WHERE  ( sysobjects.id = syscolumns.id) And  (sysobjects.xtype = 'U') " _
    '      & " AND sysobjects.name LIKE 'spt_values' ORDER BY syscolumns.colorder"
    xSql = "select * from backupfile" 'sys.tables"
    Set rsJPL = cnJPL.Execute(xSql)
    Do While Not rsJPL.EOF
        'For K = 1 To 100
            Debug.Print rsJPL(0); rsJPL(1); rsJPL(2) '; rsJPL(3); rsJPL(4); rsJPL(5)
        'Next K
        rsJPL.MoveNext
    Loop




Exit Sub

'===============================================================================================

If wOPEN > 10000000 Then wOPEN = Fix(wOPEN / 100) ' $JPL 2014-05-26
If wOPEN > 10000000 Then wOPEN = Fix(wOPEN / 100) ' $JPL 2014-05-26

GoTo Exit_sub


xSql = "select *  from rtextfield  where Aid = 0 and text_s_umidl = 792264946 and text_s_umidh  =  -105477 "


Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
  
Do While Not rsSIDE_DB.EOF
    For K = 0 To 10
        Debug.Print "rTextField"; K; rsSIDE_DB(K)
    Next K
    rsSIDE_DB.MoveNext
Loop

xSql = "select *  from rtext  where Aid = 0 and text_s_umidl = 792264946 and text_s_umidh  =  -105477 "


Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
  
Do While Not rsSIDE_DB.EOF
    For K = 0 To 10
        Debug.Print "rText"; K; rsSIDE_DB(K)
    Next K
    rsSIDE_DB.MoveNext
Loop
    
Exit Sub
'===============================================================================================

 xSql = "select * from rMesg " _
           & " where Mesg_crea_date_time <> Mesg_mod_date_time and mesg_sub_format = 'INPUT'"

'??????????????????????????????????????????????????????????????????

   
   
Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
  
Do While Not rsSIDE_DB.EOF
    Debug.Print rsSIDE_DB("mesg_uumid") & "  " & rsSIDE_DB("mesg_crea_date_time")
    rsSIDE_DB.MoveNext
Loop
    

Exit Sub


'==================================================================
            xOPEN = "": wOPEN = 0: blnOpen = False: blnOPEC = False
            For K2 = 1 To Len(X)
                X1 = Mid$(X, K2, 1)
                If X1 <> " " And X1 <> "." Then
                    If IsNumeric(X1) Then
                        If blnOPEC Then blnOpen = True: xOPEN = xOPEN & X1
                    Else
                        blnOPEC = True
                        If blnOpen Then Exit For
                    End If
                End If
            Next K2
            wOPEN = Val(xOPEN)

Exit Sub


'==================================================================
blnTransaction = False
Call lstErr_AddItem(lstErr, cmdContext, "Importation YSWIMON0"): DoEvents



V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

    xSql = "Update " & paramIBM_Library_SABSPE & ".YSWIMON0 " _
         & " set SWIMONSTA = 'S999'" _
         & " where SWIMONID > 0 and SWIMONFLUX < '20120329' and SWIMONSTA like 'S2%'"
Call FEU_ROUGE
Set rsSab = cnsab.Execute(xSql)
Call FEU_VERT
blnTransaction = True


'==================================================================

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If


End Sub


Public Sub Importation_SAB_ZCHGOPE0()
Dim V, X As String, K As Long, K2 As Long, Nb As Long
Dim xSql As String, blnTransaction As Boolean
Dim blnUpdate As Boolean, blnOk As Boolean, blnMTD As Boolean
Dim mCHGOPECRE As Long

On Error GoTo Error_Handler
'==================================================================
blnTransaction = False
Call lstErr_AddItem(lstErr, cmdContext, "Importation ZCHGOPE0"): DoEvents
Call rsYSWISAB0_Init(oldYSWISAB0)
newYSWISAB0 = oldYSWISAB0

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
blnTransaction = True

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 , " & paramIBM_Library_SAB & ".ZCHGMES0 , " & paramIBM_Library_SAB & ".ZCHGOPE0 " _
     & " where SWISABSWID > " & mSWISABSWID_Xd & " and SWISABOPEN = 0 and substring(SWISABWMTK , 1 , 1) <> '7' " _
     & " and SWISABWL20 <> '' and SWISABWL20 = CHGMESVOS" _
     & " and CHGOPEETA = CHGMESETA  and CHGOPEAGE = CHGMESAGE and CHGOPESER = CHGMESSER and CHGOPESSE = CHGMESSSE" _
     & " and CHGOPEOPE = CHGMESOPE  and CHGOPEDOS = CHGMESDOS" _
     & " order by SWISABSWID,CHGOPECRE"
     
Set rsSab = cnsab.Execute(xSql)

blnUpdate = False

Do While Not rsSab.EOF
    
    If rsSab("SWISABSWID") <> oldYSWISAB0.SWISABSWID Then
    
        If blnUpdate Or blnOk Then V = sqlYSWISAB0_Update(newYSWISAB0, oldYSWISAB0)
        blnUpdate = False
        blnOk = False
    End If
    
    If Not blnUpdate Then
        Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
        newYSWISAB0 = oldYSWISAB0
        mCHGOPECRE = rsSab("CHGOPECRE") + 19000000
        blnOk = False
        
        Select Case oldYSWISAB0.SWISABWMTK
            Case "103", "110", "111": blnMTD = True
            Case "200", "201", "202", "205": blnMTD = True
            Case Else: blnMTD = False
        End Select
        If blnMTD Then
            If oldYSWISAB0.SWISABWMTD = rsSab("CHGOPEMO1") And oldYSWISAB0.SWISABWDEV = rsSab("CHGOPEDE1") _
            And oldYSWISAB0.SWISABWAMJ <= mCHGOPECRE Then
                blnOk = True: blnUpdate = True
           ' Else
           '     If oldYSWISAB0.SWISABWAMJ <= mCHGOPECRE Then blnOk = True
            End If
        Else
            If oldYSWISAB0.SWISABWAMJ >= mCHGOPECRE Then blnOk = True
        End If
        
        If blnOk Then
            If Not blnUpdate And Mid$(oldYSWISAB0.SWISABWL20, 1, 2) = "NO" Then
                blnOk = False
            End If
            newYSWISAB0.SWISABSER = rsSab("CHGMESSER")
            newYSWISAB0.SWISABSSE = rsSab("CHGMESSSE")
            newYSWISAB0.SWISABOPEC = rsSab("CHGMESOPE")
            newYSWISAB0.SWISABOPEN = rsSab("CHGMESDOS")
            Select Case newYSWISAB0.SWISABSSE
                Case "TC", "CR": newYSWISAB0.SWISABKSRV = "S32"
                Case "CD": newYSWISAB0.SWISABKSRV = "S10"
                Case Else: newYSWISAB0.SWISABKSRV = "S01"
            End Select

        End If
    End If
    Nb = Nb + 1
    If Nb = 100 Then
        Nb = 0
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "ZCHGOPE0 : " & oldYSWISAB0.SWISABSWID & " " & newYSWISAB0.SWISABOPEN): DoEvents
    End If
    rsSab.MoveNext
Loop

If blnUpdate Or blnOk Then V = sqlYSWISAB0_Update(newYSWISAB0, oldYSWISAB0)

'==================================================================


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

If blnTransaction Then
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")

    End If
End If


End Sub




Public Sub Importation_SAB_Dossier()
Dim xSql As String


'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAB_Dossier : 1"): DoEvents
'________________________________________________________________________


xSql = "select SWISABSWID from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
     & " where SWISABWAMJ >= " & dateElp("MoisAdd", -3, DSys) & " and SWISABOPEN = 0 order by SWISABSWID "
Set rsSab = cnsab.Execute(xSql)

If rsSab.EOF Then
    mSWISABSWID_Xd = 0
Else
    mSWISABSWID_Xd = rsSab("SWISABSWID")
End If

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAB_Dossier : 2-ZSWICLA0"): DoEvents
'________________________________________________________________________
Importation_SAB_ZSWICLA0

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAB_Dossier : 3-ZCDODOS0"): DoEvents
'________________________________________________________________________
Importation_SAB_ZCDODOS0

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAB_Dossier : 4-ZCHGOPE0"): DoEvents
'________________________________________________________________________
Importation_SAB_ZCHGOPE0

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAB_Dossier : 5-SWISABWN20"): DoEvents
'________________________________________________________________________
Importation_SAB_SWISABWN20

'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAB_Dossier : 6-SWISABKPDE"): DoEvents
'________________________________________________________________________
Importation_SAB_SWISABKPDE
'________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAB_Dossier : 7-SWISABSWID_MT700"): DoEvents
'________________________________________________________________________
If mSWISABSWID_MT700 > 0 Then Importation_SAB_SWISABSWID_MT700

End Sub

Public Sub fraSWISABKSRV_Display()
Dim blnUpdate As Boolean

blnUpdate = False
cmdSWISABKSRV__Update.Visible = False
cboSWISABOPEC.Locked = True
txtSWISABOPEN.Locked = True
cboSWISABSER.Locked = True
libSWISABKSRV = " Consultation du dossier : " & oldYSWISAB0.SWISABSWID

If (oldYSWISAB0.SWISABKSRV = "S00" Or currentSSIWINUNIT = oldYSWISAB0.SWISABKSRV) And arrHab(12) Then blnUpdate = True
If arrHab(19) Then blnUpdate = True

If blnUpdate Then
    cmdSWISABKSRV__Update.Visible = True
    cboSWISABOPEC.Locked = False
    txtSWISABOPEN.Locked = False
    cboSWISABSER.Locked = False
    libSWISABKSRV = " mise à jour du dossier : " & oldYSWISAB0.SWISABSWID
End If



Call cbo_Scan(oldYSWISAB0.SWISABKSRV, cboSWISABKSRV)

Call cbo_Scan(oldYSWISAB0.SWISABOPEC, cboSWISABOPEC)

Call cbo_Scan(oldYSWISAB0.SWISABSER & "-" & oldYSWISAB0.SWISABSSE, cboSWISABSER)

lblSWISABKSRV = "Service BIA : " & oldYSWISAB0.SWISABKSRV
lblSWISABOPEC = "Service SAB : " & oldYSWISAB0.SWISABSER & "-" & oldYSWISAB0.SWISABSSE
lblSWISABOPEC = "Code opération : " & oldYSWISAB0.SWISABOPEC
lblSWISABOPEN = "Numéro opération : " & oldYSWISAB0.SWISABOPEN
If oldYSWISAB0.SWISABOPEN = 0 Then
    txtSWISABOPEN = ""
Else
    txtSWISABOPEN = oldYSWISAB0.SWISABOPEN
End If
fraSWISABKSRV.Visible = True
End Sub

Public Function fraSWISABKSRV_Control()
Dim xMsg As String

fraSWISABKSRV_Control = Null
xMsg = ""
newYSWISAB0 = oldYSWISAB0

newYSWISAB0.SWISABKSRV = Mid$(cboSWISABKSRV, 1, 3)
newYSWISAB0.SWISABOPEC = cboSWISABOPEC
newYSWISAB0.SWISABOPEN = Val(txtSWISABOPEN)
newYSWISAB0.SWISABSER = Mid$(cboSWISABSER, 1, 2)
newYSWISAB0.SWISABSSE = Mid$(cboSWISABSER, 4, 2)

If Not arrHab(19) Then

    If newYSWISAB0.SWISABKSRV <> oldYSWISAB0.SWISABKSRV Then
        If oldYSWISAB0.SWISABKSRV <> "S00" Then
             If newYSWISAB0.SWISABSER = oldYSWISAB0.SWISABSER _
            And newYSWISAB0.SWISABSSE = oldYSWISAB0.SWISABSSE _
            And newYSWISAB0.SWISABOPEC = oldYSWISAB0.SWISABOPEC _
            And newYSWISAB0.SWISABOPEN = oldYSWISAB0.SWISABOPEN Then
             Else
                 xMsg = "Vous ne pouvez pas modifier l'affectation du dossier si vous changer de service"
             End If
            End If
    Else
        If newYSWISAB0.SWISABOPEN <> 0 Then
            If Trim(newYSWISAB0.SWISABOPEC) = "" Then xMsg = "Précisez le code opération"
        End If
    End If
End If

If xMsg <> "" Then
    fraSWISABKSRV_Control = "?"
    MsgBox xMsg, vbCritical, "Affectation d'un message SWIFT"
End If

End Function



Public Function ZSWIBIC0_Select(lMsg As String) As String
Dim xSql As String
xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIBIC0 where SWIBICBIC like '" & Trim(lMsg) & "%' order by SWIBICBIC"
Set rsSabX = cnsab.Execute(xSql)

If Not rsSabX.EOF Then
    ZSWIBIC0_Select = Trim(rsSabX("SWIBICIN1")) & "  " & Trim(rsSabX("SWIBICVIL")) & "  " & Trim(rsSabX("SWIBICCOM"))
Else
    ZSWIBIC0_Select = ""
End If

End Function
Public Function ZSWIBIC0_Select_Html(lMsg As String) As String
Dim xSql As String
xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIBIC0 where SWIBICBIC like '" & Trim(lMsg) & "%' order by SWIBICBIC"
Set rsSabX = cnsab.Execute(xSql)

If Not rsSabX.EOF Then
    ZSWIBIC0_Select_Html = "<Font color = #800000>" _
                         & "&#160;&#160;&#160;&#160;&#160;" & Trim(rsSabX("SWIBICIN1")) _
                         & "<BR/>&#160;&#160;&#160;&#160;&#160;" & Trim(rsSabX("SWIBICVIL")) _
                         & "<BR/>&#160;&#160;&#160;&#160;&#160;" & Trim(rsSabX("SWIBICCOM")) & "<BR/>"
Else
    ZSWIBIC0_Select_Html = ""
End If

End Function

Public Sub Importation_SAA_SWISABWSTA_Control(lSWISABWSTA As String, lFct As String)
Dim xSql As String, Nb As Long, oldSWISABWSTA As String, newSWISABWSTA As String

On Error GoTo Error_Handler

oldSWISABWSTA = lSWISABWSTA
newSWISABWSTA = ""


If lFct = "rAppe" Then

    If Mid$(rsSIDE_DB("mesg_uumid"), 1, 1) = "O" Then
    
        If Trim(rsSIDE_DB("mesg_status")) = "COMPLETED" Then
            If rsSIDE_DB("appe_session_holder") = "FileSabOutput" Or rsSIDE_DB("appe_session_holder") = "FileMT950Input" Then newSWISABWSTA = "V"
        End If
    Else
        'If Trim(rsSIDE_DB("mesg_status")) = "COMPLETED" Then newSWISABWSTA = "E"
        If rsSIDE_DB("appe_inst_num") = 0 Then
            If Trim(rsSIDE_DB("mesg_status")) = "COMPLETED" Then
        
                Select Case Trim(rsSIDE_DB("appe_network_delivery_status"))
                    'Case "DLV_NACKED": newSWISABWSTA = "E"
                    Case "DLV_ACKED":
                        If Trim(rsSIDE_DB("appe_crea_rp_name")) = "_SI_to_SWIFT" Then
                            newSWISABWSTA = "V"
                        Else
                            newSWISABWSTA = "E"
                        End If
                        
                    Case Else: newSWISABWSTA = "E"
                End Select
            Else
                newSWISABWSTA = " "
                Select Case Trim(rsSIDE_DB("appe_network_delivery_status"))
                    Case "DLV_NACKED": newSWISABWSTA = "#"
                    'Case "DLV_ACKED": newSWISABWSTA = "V"
                End Select
            End If
        End If
    End If
Else

    If Trim(rsSIDE_DB("mesg_status")) = "COMPLETED" Then newSWISABWSTA = "E"


End If

If newSWISABWSTA <> "" Then

    If oldSWISABWSTA = "" Then
        xSql = "select SWISABWSTA from " & paramIBM_Library_SABSPE_XXX & ".YSWISAB0 " _
            & " where SWISABWID1 = " & rsSIDE_DB("Aid") _
        & " and SWISABWIDl = " & rsSIDE_DB("mesg_s_umidl") & " and SWISABWIDh = " & rsSIDE_DB("mesg_s_umidh")
        Set rsSabX = cnsab.Execute(xSql)
        If Not rsSabX.EOF Then oldSWISABWSTA = rsSabX("SWISABWSTA")
    End If
    
    If newSWISABWSTA <> oldSWISABWSTA Then

        xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWISAB0 " _
            & "set SWISABWSTA = '" & newSWISABWSTA & "' where SWISABWID1 = " & rsSIDE_DB("Aid") _
        & " and SWISABWIDl = " & rsSIDE_DB("mesg_s_umidl") & " and SWISABWIDh = " & rsSIDE_DB("mesg_s_umidh")
        Call FEU_ROUGE
        Set rsSab_Update = cnSab_Update.Execute(xSql, Nb)
        Call FEU_VERT
        If Nb = 0 Then
           'Debug.Print "err maj "; xSql
        End If
    End If
End If

Exit Sub

Error_Handler:

End Sub


Public Function fraEVE_Control_Swift_Txt(lTxt As String) As String
Dim X As String, K As Integer, wTXT As String
Dim X1 As String
Dim blnOk As Boolean

wTXT = Replace(lTxt, "#MTK", "")
wTXT = Replace(wTXT, "#AMJ", "")
wTXT = Replace(wTXT, "#DEV", "")
wTXT = Replace(wTXT, "#MTD", "")
wTXT = Replace(wTXT, "#L20", "")
wTXT = Replace(wTXT, "#N20", "")
wTXT = Replace(wTXT, "#GOS", "")
wTXT = Replace(wTXT, "#50", "")
wTXT = Replace(wTXT, "#59", "")
wTXT = Replace(wTXT, "#57", "")
wTXT = Replace(wTXT, "#32B", "")
wTXT = Replace(wTXT, "#33B", "")
wTXT = Replace(wTXT, "#36", "")
wTXT = Replace(wTXT, "#30V", "")
wTXT = Replace(wTXT, "#37G", "")
wTXT = Replace(wTXT, "#34E", "")
wTXT = Replace(wTXT, "#30T", "")
wTXT = Replace(wTXT, "#30P", "")
wTXT = Replace(wTXT, "#82A", "")
wTXT = Replace(wTXT, "#87A", "")
wTXT = Replace(wTXT, "#22C", "")
wTXT = Replace(wTXT, "#70", "")
wTXT = Replace(wTXT, "#72", "")

X = ""
For K = 1 To Len(wTXT)
    X1 = Mid$(wTXT, K, 1)
    If (X1 >= "A" And X1 <= "Z") Or X1 = " " Or X1 = vbCr Or X1 = vbLf Then
    Else
        If (X1 >= "0" And X1 <= "9") Then
        Else
            If (X1 >= "a" And X1 <= "z") Then
            Else
                If X1 = "/" Or X1 = "-" Or X1 = "?" Or X1 = ":" Or X1 = "(" Or X1 = ")" Or X1 = "." Or X1 = "," Or X1 = "'" Or X1 = "+" Then
                Else
                    X = X & " " & X1
                End If
            End If
        End If
    End If
Next K

'==============================
fraEVE_Control_Swift_Txt = X
'==============================




End Function

Public Function fraEVE_Control_Swift_Space(lTxt As String) As String
Dim X As String, K As Integer, wTXT As String
Dim X1 As String, K1 As Integer, K2 As Integer
Dim blnOk As Boolean

X = Trim(Replace(lTxt, vbCrLf & vbCrLf, vbCrLf & "." & vbCrLf))
K2 = 1
blnOk = False
Do
    K1 = InStr(K2, X, " " & vbCr)
    If K1 = 0 Then
        blnOk = True
    Else
        X = Replace(X, " " & vbCr, vbCr)
    End If

Loop Until blnOk

K2 = 1
blnOk = False
Do
    K1 = InStr(K2, X, vbCr & vbCr)
    If K1 = 0 Then
        blnOk = True
    Else
        X = Replace(X, vbCr & vbCr, vbCr)
    End If

Loop Until blnOk


K2 = 1
blnOk = False
Do
    K1 = InStr(K2, X, vbLf & vbLf)
    If K1 = 0 Then
        blnOk = True
    Else
        X = Replace(X, vbLf & vbLf, vbLf)
    End If

Loop Until blnOk

K2 = 1
blnOk = False
Do
        
    K1 = InStr(K2, X, vbCrLf)
    If K1 = 0 Then
        blnOk = True
    Else
        K2 = InStr(K1 + 2, X, vbCrLf)
        If K2 = 0 Then
            blnOk = True
        Else
            If Trim(Mid$(X, K1 + 2, K2 - K1 - 2)) = "" Then Mid$(X, K1 + 2, 1) = "."
        End If
    End If
    
Loop Until blnOk

fraEVE_Control_Swift_Space = X

End Function


Public Function Importation_SAB_YSWISAB1_IBAN(lTxt As String) As String
Static sabPays_K As Integer
Dim X As String, X1 As String, K As Integer
On Error GoTo Error_Handler
'==================================================================
Importation_SAB_YSWISAB1_IBAN = ""
If Mid$(lTxt, 1, 1) = "/" Then
    K = InStr(lTxt, Asc13)
    If K < 1 Then K = Len(lTxt) + 1
    X = Mid$(lTxt, 2, K - 2)
    X = Replace(X, " ", "")
    X = Replace(X, "IBAN:", "")
    X = Replace(X, "IBAN", "")
    If Mid$(X, 1, 5) = "12179" Then X = "FR76" & X
    'Debug.Print Mid$(lTxt, 2, K - 2), X
    If IsNull(Iban_Check(X)) Then
        Importation_SAB_YSWISAB1_IBAN = Mid$(X, 1, 2)
        X1 = ""
        If Mid$(X, 5, 10) = "1217900001" Then X1 = Mid$(X, 15, 5)
        
        If X1 <> "" Then
            X = "select CLIENARSD from " & paramIBM_Library_SAB & ".ZCLIENA0" _
                 & " where CLIENACLI = '" & "00" & X1 & "'"
            Set rsSabX = cnsab.Execute(X)
            If Not rsSabX.EOF Then Importation_SAB_YSWISAB1_IBAN = Trim(rsSabX("CLIENARSD"))
        End If
    End If
    
End If
GoTo Exit_sub

'==================================================================
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation_SAB_YSWISAB1"
Exit_sub:
End Function
Public Sub Importation_SAB_YSWISAB1_Pays(lPays As String, lPays_K As Integer, lPays_Zone As String)
Static sabPays_K As Integer
Dim X As String, K As Integer
On Error GoTo Error_Handler
'==================================================================
'___________________________________________________________________________
Select Case lPays
    Case "FR": lPays_K = sabPays_FR
    Case "DZ": lPays_K = sabPays_DZ
    Case "LY": lPays_K = sabPays_LY
    Case "US": lPays_K = sabPays_US
    Case Else
        If lPays = sabPays(sabPays_K).Id Then
            lPays_K = sabPays_K
        Else
            lPays_K = 0
            For sabPays_K = 1 To sabPays_NB
                If lPays = sabPays(sabPays_K).Id Then lPays_K = sabPays_K: Exit For
            Next sabPays_K
            If sabPays_K > sabPays_NB Then sabPays_K = 0
        End If

End Select
Select Case sabPays(lPays_K).Fiscal
    Case "1", "2", "3": lPays_Zone = "FR"
    Case "4", "5": lPays_Zone = "UE"
    Case Else
        Select Case sabPays(lPays_K).Id
            Case "CH", "IS", "LI": lPays_Zone = "UE"
            Case Else: lPays_Zone = "**"
        End Select
    End Select
GoTo Exit_sub

'==================================================================
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation_SAB_YSWISAB1"
Exit_sub:

End Sub


Public Sub lstParam_SAA_Load()
Dim xSql As String

fraParam_SAA.Visible = False
lstParam_SAA_K1.Clear
lstParam_SAA_K1.AddItem "Ajouter un enregistrement"

xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'SAA'" _
     & " and BIATABK1 = '" & Old_YBIATAB0.BIATABK1 & "' order by BIATABK2"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
        
    If Trim(rsSab("BIATABTXT")) = "" Then
        lstParam_SAA_K1.AddItem rsSab("BIATABK2")
    Else
        Select Case Trim(rsSab("BIATABK1"))
            Case "Amount", "Approval"
                lstParam_SAA_K1.AddItem rsSab("BIATABK2") & "      " & Format(CCur(Trim(rsSab("BIATABTXT"))), "### ### ### ### ##0")
            Case "Jrnl_Event"
                lstParam_SAA_K1.AddItem rsSab("BIATABK2") & "      " & Mid$(rsSab("BIATABTXT"), 100, 4) & " " & Trim(rsSab("BIATABTXT"))
            Case Else
                lstParam_SAA_K1.AddItem rsSab("BIATABK2") & "      " & Trim(rsSab("BIATABTXT"))
        End Select
    End If
    rsSab.MoveNext
Loop

End Sub

Public Sub Importation_SAA_Alerte_Init()
Dim xSql As String
Call lstErr_AddItem(lstErr, cmdContext, "Importation_SAA_Alerte_Init"): DoEvents

xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'SAA'"
Set rsSab = cnsab.Execute(xSql)
ReDim arrSAA_Usr_Id(rsSab(0) + 1), arrSAA_Usr_MTD(rsSab(0) + 1)

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'SAA' order by BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSql)

arrSAA_Usr_Nb = 0
Do While Not rsSab.EOF
    Select Case Trim(rsSab("BIATABK1"))
        Case "Approval"
            arrSAA_Usr_Nb = arrSAA_Usr_Nb + 1
            arrSAA_Usr_Id(arrSAA_Usr_Nb) = Trim(rsSab("BIATABK2"))
            arrSAA_Usr_MTD(arrSAA_Usr_Nb) = Val(Trim(rsSab("BIATABTXT")))
         Case "Amount"
            Select Case Trim(rsSab("BIATABK2"))
            Case "103": curSAA_103_EUR = Val(Trim(rsSab("BIATABTXT")))
            Case "202": curSAA_202_EUR = Val(Trim(rsSab("BIATABTXT")))
            Case "202_BOTC": curSAA_202_BOTC_EUR = Val(Trim(rsSab("BIATABTXT")))
            End Select
            
   End Select
    rsSab.MoveNext
Loop

End Sub

Public Sub Form_Init_Options_J()
Dim xSql As String



'______________________________________________________________________
xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'SAA' and BIATABK1 = 'Jrnl_Event'"
Set rsSab = cnsab.Execute(xSql)
arrJrnl_Event_Nb = rsSab(0) + 1
ReDim arrJrnl_Event_Id(arrJrnl_Event_Nb), arrJrnl_Event_Lib(arrJrnl_Event_Nb)
arrJrnl_Event_Nb = 0

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'SAA' and BIATABK1 = 'Jrnl_Event'" _
     & "  order by BIATABK2"
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    arrJrnl_Event_Nb = arrJrnl_Event_Nb + 1
    arrJrnl_Event_Id(arrJrnl_Event_Nb) = Trim(rsSab("BIATABK2"))
    arrJrnl_Event_Lib(arrJrnl_Event_Nb) = Trim(Mid$(rsSab("BIATABTXT"), 1, 99))
    
    rsSab.MoveNext
Loop

End Sub

Public Sub cmdParam_SAA_Control()

New_YBIATAB0 = Old_YBIATAB0
Select Case Trim(New_YBIATAB0.BIATABK1)
    Case "Amount", "Approval"
        New_YBIATAB0.BIATABTXT = Format(Val(txtParam_SAA_MTD), "000000000000")
    Case "Jrnl_Event"

        New_YBIATAB0.BIATABK2 = txtParam_SAA_K2
        New_YBIATAB0.BIATABTXT = txtParam_SAA_TXT
        Mid$(New_YBIATAB0.BIATABTXT, 103, 1) = Mid$(cboParam_SAA_TopK, 1, 1)
        Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = Mid$(cboParam_SAA_Alerte, 1, 3)
        If optParam_SAA_TXT_RMA_A = True Then Mid$(New_YBIATAB0.BIATABTXT, 104, 1) = "A"
        If optParam_SAA_TXT_RMA_R = True Then Mid$(New_YBIATAB0.BIATABTXT, 104, 1) = "R"
    Case "Mesg_Type", "Mesg_Fields"
        New_YBIATAB0.BIATABK2 = Trim(txtParam_SAA_K2)
        New_YBIATAB0.BIATABTXT = Trim(txtParam_SAA_TXT)
   Case Else
        New_YBIATAB0.BIATABTXT = txtParam_SAA_K2
        New_YBIATAB0.BIATABK2 = txtParam_SAA_K2
End Select

End Sub

Public Sub Importation_Jrnl_Init()
Dim X As String, K As Long

arrJrnl_Nb = 0
X = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'SAA' and BIATABK1 = 'Jrnl_Event' and substring(BIATABTXT,100,4) <> ''"
Set rsSab = cnsab.Execute(X)
K = rsSab(0) + 1
ReDim arrJrnl_Comp_Name(K), arrJrnl_Event_Num(K), arrJrnl_Alerte(K), arrJrnl_Top(K)

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'SAA' and BIATABK1 = 'Jrnl_Event' and substring(BIATABTXT,100,4) <> ''" _
     & " order by substring(BIATABK2,5,7)"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    arrJrnl_Nb = arrJrnl_Nb + 1
    arrJrnl_Comp_Name(arrJrnl_Nb) = Trim(Mid$(rsSab("BIATABK2"), 1, 4))

    arrJrnl_Event_Num(arrJrnl_Nb) = Val(Trim(Mid$(rsSab("BIATABK2"), 5, 7)))
    
    X = rsSab("BIATABTXT")
    arrJrnl_Alerte(arrJrnl_Nb) = Trim(Mid$(X, 100, 3))
    arrJrnl_Top(arrJrnl_Nb) = Trim(Mid$(X, 103, 1))
    rsSab.MoveNext
Loop

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
    
    'Call cmdSelect_SQL_rJrnl_Date(X & " 00:00:00", X & " 23:59:59", wDateFrom, wDateTo)
    
    wDateFrom = 1200802192 - DateDiff("s", "01/01/2000 00:00:00", X & " 00:00:00")
    wDateTo = 1200802192 - DateDiff("s", "01/01/2000 00:00:00", X & " 23:59:59")

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

cmdMail_MT.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Public Function Importation_Jrnl_Top_10007() As String
Dim X As String, K1 As Integer, K2 As Integer

Importation_Jrnl_Top_10007 = "S10"

Call cmdSelect_SQL_YSAAJRN0_rMesg(newYSAAJRN0.SAAJRNTOPX, newYSAAJRN0.SAAJRNSUFX)
X = "select text_data_block from rtext " _
    & "where Aid = " & Mesg_aid _
    & " and text_s_umidl = " & mesg_s_umidl _
    & " and text_s_umidh  =  " & mesg_s_umidh
Set rsSIDE_X = cnSIDE_DB.Execute(X)
If Not rsSIDE_X.EOF Then
    X = rsSIDE_X("text_data_block")
    K1 = InStr(1, X, ":20:")
    If K1 > 0 And Len(X) > K1 + 8 Then
        Importation_Jrnl_Top_10007 = Mid$(X, K1 + 4, 4)
    End If
End If

End Function

Public Sub Importation_SAB_YSWISAB1_103()
Dim xSql As String, X As String, K As Integer, K2 As Integer, K3 As Integer
Dim mField As String
Dim wText_Data_Block As String

On Error GoTo Error_Handler
'==================================================================
xSql = "select *  from rtextField  " _
& "where Aid = " & rsSab("SWISABWID1") _
& " and text_s_umidl = " & rsSab("SWISABWIDL") _
& " and text_s_umidh  =  " & rsSab("SWISABWIDH") _
& " order by field_cnt"
Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then
    Do While Not rsSIDE_DB.EOF
        mField = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
        X = rsSIDE_DB("value") & Asc13
        Select Case mField
            Case "50F":
                        K3 = InStr(X, vbLf & "3/") + 3
                        If K3 > 3 Then wMT_50P = Mid$(X, K3, 2)
            Case "50K":
                        If Mid$(X, 1, 1) = "/" Then
                            Call Importation_SAB_YSWISAB1_103_Pays(X, wMT_50P)
                        End If
            Case "52A": Call Importation_SAB_YSWISAB1_103_BIC(X, wMT_52A, wMT_52P)
            Case "57A": Call Importation_SAB_YSWISAB1_103_BIC(X, wMT_57A, wMT_57P)
            Case "59A", "59F", "59": Call Importation_SAB_YSWISAB1_103_Pays(X, wMT_59P)
            Case "71A": newYSWISAB1.SWISABW71A = Mid$(X, 1, 1)
            Case "72":
                        K2 = InStr(X, "/REC/EBA")
                         If K2 > 0 Then
                             newYSWISAB1.SWISABWEBA = "E"
                         Else
                             K2 = InStr(X, "/PDT/TGT")
                             If K2 > 0 Then newYSWISAB1.SWISABWEBA = "T"
                        End If

        End Select
        rsSIDE_DB.MoveNext
    
    Loop
Else
    xSql = "select * from rtext " _
        & "where Aid = " & rsSab("SWISABWID1") _
        & " and text_s_umidl = " & rsSab("SWISABWIDL") _
        & " and text_s_umidh  =  " & rsSab("SWISABWIDH")
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    If Not rsSIDE_DB.EOF Then
        V = rsSIDE_DB("text_data_block"): wText_Data_Block = IIf(IsNull(V), "", V & vbCrLf & ":")
    '_____________________________________________________
         K = InStr(wText_Data_Block, ":50F:") + 5
        If K > 5 Then
            K3 = InStr(K, wText_Data_Block, Asc10 & ":")
            If K3 > 0 Then
                X = Mid$(wText_Data_Block, K, K3 - K)
                K3 = InStr(X, vbLf & "3/") + 3
                If K3 > 3 Then wMT_50P = Mid$(X, K3, 2)
                
            End If
        End If
     '_____________________________________________________
         K = InStr(wText_Data_Block, ":50K:") + 5
        If K > 5 Then
            If Mid$(wText_Data_Block, 1, 1) = "/" Then
                K3 = InStr(K, wText_Data_Block, Asc10 & ":")
                If K3 > 0 Then
                    X = Mid$(wText_Data_Block, K, K3 - K)
                    Call Importation_SAB_YSWISAB1_103_Pays(X, wMT_50P)
                End If
            End If
        End If
   '_____________________________________________________
         K = InStr(wText_Data_Block, ":52A:") + 5
        If K > 5 Then
            K3 = InStr(K, wText_Data_Block, Asc10 & ":")
            If K3 > 0 Then
                X = Mid$(wText_Data_Block, K, K3 - K)
                Call Importation_SAB_YSWISAB1_103_BIC(X, wMT_52A, wMT_52P)
                
            End If
        End If
    '_____________________________________________________
        K = InStr(wText_Data_Block, ":57A:") + 5
        If K > 5 Then
            K3 = InStr(K, wText_Data_Block, Asc10 & ":")
            If K3 > 0 Then
                X = Mid$(wText_Data_Block, K, K3 - K)
                Call Importation_SAB_YSWISAB1_103_BIC(X, wMT_57A, wMT_57P)
                
            End If
        End If
     '_____________________________________________________
         K = InStr(wText_Data_Block, ":59A:") + 5
        If K <= 5 Then
            K = InStr(wText_Data_Block, ":59:") + 4
        End If
        If K > 4 Then
            K3 = InStr(K, wText_Data_Block, Asc10 & ":")
            If K3 > 0 Then
                X = Mid$(wText_Data_Block, K, K3 - K)
                Call Importation_SAB_YSWISAB1_103_Pays(X, wMT_59P)
                
            End If
        End If
     '_____________________________________________________
    
        K = InStr(wText_Data_Block, ":71A:") + 5
        If K > 5 Then newYSWISAB1.SWISABW71A = Mid$(wText_Data_Block, K, 1)
   '_____________________________________________________
    
        K = InStr(wText_Data_Block, ":72:") + 4
        If K > 4 Then
            K2 = InStr(K, wText_Data_Block, "/REC/EBA")
             If K2 > 0 Then
                 newYSWISAB1.SWISABWEBA = "E"
             Else
                 K2 = InStr(K, wText_Data_Block, "/PDT/TGT")
                 If K2 > 0 Then newYSWISAB1.SWISABWEBA = "T"
            End If
        End If
    End If
End If
'_____________________________________________________
'Debug.Print wText_Data_Block


newYSWISAB1.SWISABW52A = wMT_52A
If newYSWISAB1.SWISABW52A = "" Then newYSWISAB1.SWISABW52A = wMT_BIC_E

newYSWISAB1.SWISABW50P = wMT_50P
If newYSWISAB1.SWISABW50P = "" Then
    newYSWISAB1.SWISABW50P = Mid$(newYSWISAB1.SWISABW52A, 5, 2)
End If
newYSWISAB1.SWISABW57A = wMT_57A
If newYSWISAB1.SWISABW57A = "" Then newYSWISAB1.SWISABW57A = wMT_BIC_S
newYSWISAB1.SWISABW59P = wMT_59P
If newYSWISAB1.SWISABW59P = "" Then
    newYSWISAB1.SWISABW59P = Mid$(newYSWISAB1.SWISABW57A, 5, 2)
End If
GoTo Exit_sub

'==================================================================
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation_SAB_YSWISAB1"
Exit_sub:

End Sub
Public Sub Importation_SAB_YSWISAB1_202()
Dim xSql As String, X As String, K As Integer, K2 As Integer, K3 As Integer
Dim mField As String
Dim wText_Data_Block As String
On Error GoTo Error_Handler
'==================================================================
xSql = "select *  from rtextField  " _
& "where Aid = " & rsSab("SWISABWID1") _
& " and text_s_umidl = " & rsSab("SWISABWIDL") _
& " and text_s_umidh  =  " & rsSab("SWISABWIDH") _
& " order by field_cnt"
Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)

If Not rsSIDE_DB.EOF Then
    Do While Not rsSIDE_DB.EOF
        mField = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
        X = rsSIDE_DB("value") & Asc13
        Select Case mField
            Case "52A": Call Importation_SAB_YSWISAB1_202_BIC(X, wMT_52A, wMT_52P)
            Case "57A": Call Importation_SAB_YSWISAB1_202_BIC(X, wMT_57A, wMT_57P)
            Case "58A": Call Importation_SAB_YSWISAB1_202_BIC(X, wMT_58A, wMT_58P)
            Case "71A": newYSWISAB1.SWISABW71A = Mid$(X, K, 1)
            Case "72":
                        K2 = InStr(X, "/REC/EBA")
                         If K2 > 0 Then
                             newYSWISAB1.SWISABWEBA = "E"
                         Else
                             K2 = InStr(X, "/PDT/TGT")
                             If K2 > 0 Then newYSWISAB1.SWISABWEBA = "T"
                        End If

        End Select
        rsSIDE_DB.MoveNext
    
    Loop
Else
    xSql = "select * from rtext " _
        & "where Aid = " & rsSab("SWISABWID1") _
        & " and text_s_umidl = " & rsSab("SWISABWIDL") _
        & " and text_s_umidh  =  " & rsSab("SWISABWIDH")
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    If Not rsSIDE_DB.EOF Then
        V = rsSIDE_DB("text_data_block"): wText_Data_Block = IIf(IsNull(V), "", V & vbCrLf & ":")
    '_____________________________________________________
         K = InStr(wText_Data_Block, ":52A:") + 5
        If K > 5 Then
            K3 = InStr(K, wText_Data_Block, Asc10 & ":")
            If K3 > 0 Then
                X = Mid$(wText_Data_Block, K, K3 - K)
                Call Importation_SAB_YSWISAB1_202_BIC(X, wMT_52A, wMT_52P)
                
            End If
        End If
    '_____________________________________________________
        K = InStr(wText_Data_Block, ":57A:") + 5
        If K > 5 Then
            K3 = InStr(K, wText_Data_Block, Asc10 & ":")
            If K3 > 0 Then
                X = Mid$(wText_Data_Block, K, K3 - K)
                Call Importation_SAB_YSWISAB1_202_BIC(X, wMT_57A, wMT_57P)
                
            End If
        End If
     '_____________________________________________________
         K = InStr(wText_Data_Block, ":58A:") + 5
        If K > 5 Then
            K3 = InStr(K, wText_Data_Block, Asc10 & ":")
            If K3 > 0 Then
                X = Mid$(wText_Data_Block, K, K3 - K)
                Call Importation_SAB_YSWISAB1_202_BIC(X, wMT_58A, wMT_58P)
                
            End If
        End If
     '_____________________________________________________
    
        K = InStr(wText_Data_Block, ":71A:") + 5
        If K > 5 Then newYSWISAB1.SWISABW71A = Mid$(wText_Data_Block, K, 1)
   '_____________________________________________________
    
        K = InStr(wText_Data_Block, ":72:") + 4
        If K > 4 Then
            K2 = InStr(K, wText_Data_Block, "/REC/EBA")
             If K2 > 0 Then
                 newYSWISAB1.SWISABWEBA = "E"
             Else
                 K2 = InStr(K, wText_Data_Block, "/PDT/TGT")
                 If K2 > 0 Then newYSWISAB1.SWISABWEBA = "T"
            End If
        End If
    End If
End If
'_____________________________________________________
'Debug.Print wText_Data_Block

newYSWISAB1.SWISABW52A = wMT_52A
newYSWISAB1.SWISABW50P = wMT_52P
If newYSWISAB1.SWISABW52A = "" Then
    newYSWISAB1.SWISABW52A = wMT_BIC_E
    If newYSWISAB1.SWISABW50P = "" Then newYSWISAB1.SWISABW50P = Mid$(newYSWISAB1.SWISABW52A, 5, 2)
End If
newYSWISAB1.SWISABW57A = wMT_58A
If newYSWISAB1.SWISABW57A = "" Then newYSWISAB1.SWISABW57A = wMT_57A
newYSWISAB1.SWISABW59P = wMT_58P
If newYSWISAB1.SWISABW59P = "" Then newYSWISAB1.SWISABW59P = wMT_57P
GoTo Exit_sub

'==================================================================
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation_SAB_YSWISAB1"
Exit_sub:

End Sub



Public Sub Importation_SAB_YSWISAB1_202_BIC(lX As String, wMT_BIC As String, wMT_Pays As String)
Dim K2 As Integer, K3 As Integer, X1 As String, X2 As String
On Error GoTo Error_Handler
'==================================================================
wMT_BIC = "": wMT_Pays = ""

K2 = InStr(1, lX, Asc13)
X = Mid$(lX, 1, K2 - 1)
If Mid$(X, 1, 1) = "/" Then
    X1 = Importation_SAB_YSWISAB1_IBAN(X)
    If X1 <> "" Then wMT_Pays = X1
Else
    wMT_BIC = X
End If

K2 = InStr(1, lX, vbLf)
If K2 > 0 Then
    K3 = InStr(K2 + 1, lX, Asc13)
    If K3 > 0 Then wMT_BIC = Mid$(lX, K2 + 1, K3 - K2)

End If

If wMT_Pays = "" And Len(wMT_BIC) > 6 Then wMT_Pays = Mid$(wMT_BIC, 5, 2)
GoTo Exit_sub

'==================================================================
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation_SAB_YSWISAB1"
Exit_sub:

End Sub
Public Sub Importation_SAB_YSWISAB1_103_Pays(lX As String, wMT_Pays As String)
Dim K2 As Integer, K3 As Integer, X1 As String, X2 As String
On Error GoTo Error_Handler
'==================================================================
wMT_Pays = ""

K2 = InStr(1, lX, Asc13)
X = Mid$(lX, 1, K2 - 1)
If Mid$(X, 1, 1) = "/" Then
    X1 = Importation_SAB_YSWISAB1_IBAN(X)
    If X1 <> "" Then wMT_Pays = X1
End If

GoTo Exit_sub

'==================================================================
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation_SAB_YSWISAB1"
Exit_sub:

End Sub

Public Sub Importation_SAB_YSWISAB1_103_BIC(lX As String, wMT_BIC As String, wMT_Pays As String)
Dim K2 As Integer, K3 As Integer, X1 As String, X2 As String
On Error GoTo Error_Handler
'==================================================================
wMT_BIC = "": wMT_Pays = ""

K2 = InStr(1, lX, Asc13)
X = Mid$(lX, 1, K2 - 1)
If Mid$(X, 1, 1) = "/" Then
    X1 = Importation_SAB_YSWISAB1_IBAN(X)
    If X1 <> "" Then wMT_Pays = X1
Else
    wMT_BIC = X
End If

K2 = InStr(1, lX, vbLf)
If K2 > 0 Then
    K3 = InStr(K2 + 1, lX, Asc13)
    If K3 > 0 Then wMT_BIC = Mid$(lX, K2 + 1, K3 - K2)

End If

If wMT_Pays = "" And Len(wMT_BIC) > 6 Then wMT_Pays = Mid$(wMT_BIC, 5, 2)
GoTo Exit_sub

'==================================================================
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : Importation_SAB_YSWISAB1"
Exit_sub:

End Sub




Public Sub cmdEVE_Reset()
cmdEVE_Ok.Visible = False
cmdEVE_Ok_Clôture.Visible = False
cmdEVE_Ok_àClôturer.Visible = False
cmdEVE_Ignore.Visible = False
cmdEVE_Invalidation.Visible = False
cmdEVE_Dupliquer.Visible = False

blnGOSDOSSTAD_C = False
blnGOSDOSSTAD_X = False
End Sub
Public Sub cmdEVE_Set()

cmdEVE_Reset

Select Case currentSSIWINUNIT
    Case oldYGOSDOS0.GOSDOSISRV
        cmdEVE_Ok.Visible = arrHab(2)
        If blnHab_YGOSEVE0_New Then
            cmdEVE_Ok_àClôturer.Visible = arrHab(2)
            cmdEVE_Ok_Clôture.Visible = arrHab(3)
        End If

    Case oldYGOSDOS0.GOSDOSGSRV
        cmdEVE_Ok.Visible = arrHab(2)
        If oldYGOSDOS0.GOSDOSSTAG <> " " Then
            If blnHab_YGOSEVE0_New Then
                cmdEVE_Ok_àClôturer.Visible = arrHab(2)
                cmdEVE_Ok_Clôture.Visible = arrHab(3)
            End If
        End If
End Select


End Sub


Public Sub cmdEVE_Set_Ok(lK As Integer)

cmdEVE_Reset

cmdEVE_Ok.Visible = arrHab(lK)

End Sub

Public Sub cmdEVE_Ok_Control()

If fraPJ.Visible Then
    fraPJ.Visible = False
    If Len(rtfPJ.Text) > 0 Then
        oldFileName = "C:\temp\DROPI.rtf"
        If Dir(oldFileName) <> "" Then Kill oldFileName
        newDirPath = paramGOSDOS_Path & oldYGOSDOS0.GOSDOSIDD
        X = InputBox("Préciser le nom du document", "BIA_GOS : Pièce jointe")
        If X = "" Then X = DSYS_Time
        newFileName = X & ".rtf"
        newFileExtension = "rtf"
        txtGOSEVETXT = newFileName
        txtGOSEVETXT.Locked = False
        rtfPJ.SaveFile oldFileName
    End If

Else
    If IsNull(fraEVE_Control) Then
        fraEVE.Visible = False
        blnYGOSDOS0_Update = True
        fraMail_Confirm
        
    End If
End If

End Sub


Public Sub cmdSelect_SQL_3x()
cmdSelect_SQL_K = "3"
Call cbo_Scan("x", cboSelect_3_GOSDOSSTAD)
cboSelect_3_GOSDOSSTAG.ListIndex = 0
cmdSelect_SQL_3
End Sub

Public Sub fgSelect_DisplayLine_YGOSDOS0_Color(lColor As Long, lBackColor As Long)
Select Case xYGOSDOS0.GOSDOSSTAD
    Case " "
        Select Case xYGOSDOS0.GOSDOSSTAG
            Case "V": lBackColor = mColor_G1
            Case "R": lBackColor = mColor_W0
            Case Else: lBackColor = 0
        End Select

    Case "C": lBackColor = &HE0E0E0
    Case "x": lBackColor = mColor_Y1
    Case "A": lBackColor = mColor_W1
End Select

Select Case xYGOSDOS0.GOSDOSSTAG
    Case "V": lColor = &H4000&
    Case "R": lColor = vbMagenta
    Case Else: lColor = vbBlue
End Select

End Sub

Public Sub param_Init_MT_Fields()
Dim X As String, X2 As String, K As Integer

Dim V

On Error GoTo Error_Handler
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

New_YBIATAB0.BIATABID = "SAA"
New_YBIATAB0.BIATABK1 = "Mesg_"

Open "c:\Temp\MTxxx.txt" For Input As #1
Do Until EOF(1)
    Line Input #1, X
    X = Trim(X)
    If X <> " " Then
        Select Case Mid$(X, 1, 2)
            Case "MT":
                X = Trim(Replace(X, "MT", ""))
                New_YBIATAB0.BIATABK1 = "MT_" & Mid$(X, 1, 3)
            Case "M ", "O "
                K = InStr(3, X, " ")
                If K > 0 Then
                    X2 = Replace(Mid$(X, 3, K - 3), "a", "")
                    New_YBIATAB0.BIATABK2 = X2
                    New_YBIATAB0.BIATABTXT = Mid$(X, K + 1, Len(X) - K)
                    V = sqlYBIATAB0_Insert(New_YBIATAB0)
                Else
                    MsgBox X, vbQuestion
                End If
        End Select
    End If
               
Loop
Close


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

End Sub


Public Sub cmdSelect_Reset_4(lFct As String)
Select Case lFct
    Case "4"
        lblSelect_4_GOSDOSISRV = "Service initiateur"
        lblSelect_4_GOSDOSGSRV = "Service gestionnaire"
        lblSelect_4_GOSDOSECHD = "Date échéance <= au"
        
        lblSelect_4_GOSDOSISRV.Visible = True
        lblSelect_4_GOSDOSGSRV.Visible = True
        lblSelect_4_GOSDOSECHD.Visible = True
        cboSelect_4_GOSDOSISRV.Visible = True
        cboSelect_4_GOSDOSGSRV.Visible = True
        txtSelect_4_GOSDOSECHD.Visible = True
    Case "4 Journal"
        lblSelect_4_GOSDOSISRV = "Dossiers créés par"
        lblSelect_4_GOSDOSGSRV = "+ autres interventions de"
        lblSelect_4_GOSDOSECHD = "Date des évènements"
        
        lblSelect_4_GOSDOSISRV.Visible = True
        lblSelect_4_GOSDOSGSRV.Visible = True
        lblSelect_4_GOSDOSECHD.Visible = True
        cboSelect_4_GOSDOSISRV.Visible = True
        cboSelect_4_GOSDOSGSRV.Visible = True
        txtSelect_4_GOSDOSECHD.Visible = True
    Case "4 Swi>"
        lblSelect_4_GOSDOSGSRV = "Service origine de l'événement Swi>"
        
        lblSelect_4_GOSDOSGSRV.Visible = True
        lblSelect_4_GOSDOSISRV.Visible = False
        lblSelect_4_GOSDOSECHD.Visible = False
        cboSelect_4_GOSDOSISRV.Visible = False
        cboSelect_4_GOSDOSGSRV.Visible = True
        txtSelect_4_GOSDOSECHD.Visible = False
End Select
End Sub

Public Sub Form_Migration_20120511()
Dim xSql As String, X As String
xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0" _
     & " where GOSEVEIDD = -1" _
     & " order by SUBSTRING(GOSEVETXT , 1 , 10)"
     
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    Call rsYGOSEVE0_GetBuffer(rsSab, oldYGOSEVE0)
    newYGOSEVE0 = oldYGOSEVE0
    X = Mid$(oldYGOSEVE0.GOSEVETXT, 20, Len(oldYGOSEVE0.GOSEVETXT) - 19)
     newYGOSEVE0.GOSEVETXT = Mid$(oldYGOSEVE0.GOSEVETXT, 1, 19) & Space$(11) & Trim(X)
     cmdYGOSDOS0_Update "", "Update", "", "", ""
    rsSab.MoveNext
Loop

End Sub

Public Sub Form_Migration_20120523()
Dim xSql As String, X As String, Nb As Long

X = " set SWISABK999 = ' ' where SWISABK999 = 'D'"
xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWISAB0" & X
Call FEU_ROUGE
Set rsSab = cnsab.Execute(xSql, Nb)

MsgBox X & vbCrLf & "Nb maj : " & Nb, vbInformation, "Form_Migration_20120523"
'___________________________________________________________________________________

X = " set SWISABKSRV = 'S01 ' where SWISABXGOS = 'G' and SWISABKSRV = 'S00'"
xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWISAB0" & X

Set rsSab = cnsab.Execute(xSql, Nb)
    
MsgBox X & vbCrLf & "Nb maj : " & Nb, vbInformation, "Form_Migration_20120523"
'___________________________________________________________________________________

X = " set SWISABKSRV = 'S01 ' where SWISABXEVE = 'G' and SWISABKSRV = 'S00'"
xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWISAB0" & X

Set rsSab = cnsab.Execute(xSql, Nb)
    
MsgBox X & vbCrLf & "Nb maj : " & Nb, vbInformation, "Form_Migration_20120523"
'___________________________________________________________________________________


X = " set SWISABKSRV = 'S01 ' where SWISABXEVE = '*' and SWISABKSRV = 'S00'"
xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWISAB0" & X

Set rsSab = cnsab.Execute(xSql, Nb)
    
MsgBox X & vbCrLf & "Nb maj : " & Nb, vbInformation, "Form_Migration_20120523"
'___________________________________________________________________________________

X = " set SWISABXEVE = '=' where SWISABXEVE = '*' and SWISABKSRV not in ( 'S00' , 'S01' ) and SWISABK999 in ('!' , '@')"
xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWISAB0" & X

Set rsSab = cnsab.Execute(xSql, Nb)
    
MsgBox X & vbCrLf & "Nb maj : " & Nb, vbInformation, "Form_Migration_20120523"
'___________________________________________________________________________________

X = " set SWISABK999 = 'G' where SWISABXEVE = 'G' and  SWISABK999 <> 'G'"
xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWISAB0" & X

Set rsSab = cnsab.Execute(xSql, Nb)
Call FEU_VERT
MsgBox X & vbCrLf & "Nb maj : " & Nb, vbInformation, "Form_Migration_20120523"
'___________________________________________________________________________________
End Sub

Public Sub cmdSelect_SQL_Stat_Detail()
Dim xSql As String, K As Long, X As String
Dim arrIDD_Nb As Integer
Dim arrIDD() As Long, arrPays() As String
Dim wGOSDOSIDD As Long
On Error GoTo Error_Handler

'===================================================================================

xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YGOSDOS0" _
     & " where GOSDOSIDD > 0 and GOSDOSISRV ='" & currentSSIWINUNIT & "' and GOSDOSSTAD <> 'A'" _
     & " and GOSDOSIAMJ >= " & wAmjMin & " and GOSDOSIAMJ <= " & wAmjMax
Set rsSab = cnsab.Execute(xSql)
arrIDD_Nb = rsSab(0)
ReDim arrIDD(arrIDD_Nb + 1) As Long, arrPays(arrIDD_Nb + 1) As String

xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSDOS0, " & paramIBM_Library_SABSPE & ".YSWISAB0, " & paramIBM_Library_SABSPE & ".YSWISAB1 " _
     & " where GOSDOSIDD > 0 and GOSDOSISRV ='" & currentSSIWINUNIT & "' and GOSDOSSTAD <> 'A'" _
     & " and GOSDOSIAMJ >= " & wAmjMin & " and GOSDOSIAMJ <= " & wAmjMax _
     & " and GOSDOSWID1 = SWISABWID1 and GOSDOSWIDL = SWISABWIDL and GOSDOSWIDH = SWISABWIDH " _
     & " and SWISABSWID = SWISAB1ID" _
     & " order by  GOSDOSIDD"
Set rsSab = cnsab.Execute(xSql)

K = 0
Do While Not rsSab.EOF
    K = K + 1
    arrIDD(K) = rsSab("GOSDOSIDD")
    arrPays(K) = rsSab("GOSDOSPAYS") & " - " & Mid$(rsSab("SWISABWBIC"), 5, 2) & " - " & rsSab("SWISABW50P") & " - " & rsSab("SWISABW59P")
    
    rsSab.MoveNext
Loop

'===================================================================================

Call lstErr_AddItem(lstErr, cmdContext, "BIA_GOS : Détail ventilation"): DoEvents

xSql = "select * from " & paramIBM_Library_SABSPE & ".YGOSDOS0, " & paramIBM_Library_SABSPE & ".YGOSEVE0 " _
     & " where GOSDOSIDD > 0 and GOSDOSISRV ='" & currentSSIWINUNIT & "' and GOSDOSSTAD <> 'A'" _
     & " and GOSDOSIAMJ >= " & wAmjMin & " and GOSDOSIAMJ <= " & wAmjMax _
     & " and GOSEVEIDD = -1 and GOSEVEGSRV = '" & currentSSIWINUNIT & "' and GOSDOSLABK = substring(GOSEVETXT , 1 , 10) " _
     & " order by substring(GOSEVETXT , 20 , 3) , GOSDOSLABK"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    mXls1_Row = mXls1_Row + 1
    wGOSDOSIDD = rsSab("GOSDOSIDD")
    X = Trim(Mid$(rsSab("GOSEVETXT"), 20, 3))
    For K = 1 To arrStaC_Nb
        If X = arrStaC(K).Code Then
            arrStaC(K).Row2 = mXls1_Row
            If arrStaC(K).Row1 = 0 Then arrStaC(K).Row1 = mXls1_Row
        End If
    Next K
    
    wsExcel.Cells(mXls1_Row, 1) = X
    wsExcel.Cells(mXls1_Row, 2) = rsSab("GOSDOSLABK")
    wsExcel.Cells(mXls1_Row, 3) = wGOSDOSIDD
    If rsSab("GOSDOSSTAG") = "R" Then wsExcel.Cells(mXls1_Row, 4) = 1
    
    X = rsSab("GOSDOSPAYS")
    For K = 1 To arrIDD_Nb
        If wGOSDOSIDD = arrIDD(K) Then X = arrPays(K)
    Next K
    
    wsExcel.Cells(mXls1_Row, mXls1_Cols - 1) = X
    For K = 1 To arrStaP_Nb
        If InStr(X, arrStaP(K).Code) Then Exit For
    Next K
    wsExcel.Cells(mXls1_Row, K + 4) = 1
    wsExcel.Cells(mXls1_Row, mXls1_Cols) = rsSab("GOSDOSCLI")
    wsExcel.Cells(mXls1_Row, mXls1_Cols_WMTK) = "MT " & rsSab("GOSDOSWMTK")
    rsSab.MoveNext
Loop
'======================================================================================================

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
End Sub
Public Sub cmdSelect_SQL_Stat_Recapitulatif(lFct As String)
Dim K As Long, K2 As Long, X As String
Dim mRow1 As Long
On Error GoTo Error_Handler

'===================================================================================

For K = 1 To arrStaC_Nb
    If Len(arrStaC(K).Code) = 1 Then
        If mRow1 > 0 Then
            wsExcel.Cells(mRow1, 1).Interior.Color = mColor_G2
            wsExcel.Cells(mRow1, 2).Interior.Color = mColor_G2
            For K2 = 3 To mXls1_Cols - 1
                X = Mid$("ABCDEFGHIJKLMNOPQRSTUVW", K2, 1)
                wsExcel.Cells(mRow1, K2).Interior.Color = mColor_G2
                wsExcel.Cells(mRow1, K2).FormulaLocal = "=SOMME(" & X & mRow1 + 1 & ":" & X & mXls1_Row & ")"
                wsExcel.Cells(mRow1, K2).Font.Bold = True
                wsExcel.Cells(mRow1, K2).Font.Color = vbBlue
            Next K2
        End If
        mRow1 = mXls1_Row + 1
    End If
    
    mXls1_Row = mXls1_Row + 1
    wsExcel.Cells(mXls1_Row, 1) = arrStaC(K).Code
    
    wsExcel.Cells(mXls1_Row, 2) = arrStaC(K).Lib
    
    For K2 = 4 To mXls1_Cols - 2
        If arrStaC(K).Row1 > 0 Then
            X = Mid$("ABCDEFGHIJKLMNOPQRSTUVW", K2, 1)
            Select Case lFct
                Case "":
                        wsExcel.Cells(mXls1_Row, K2).FormulaLocal = "=SOMME(Détail!" & X & arrStaC(K).Row1 & ":" & "Détail!" & X & arrStaC(K).Row2 & ")"
                Case Else:
                        wsExcel.Cells(mXls1_Row, K2).FormulaLocal = "=SOMME.SI.ENS(Détail!" & X & arrStaC(K).Row1 & ":Détail!" & X & arrStaC(K).Row2 _
                                            & ";Détail!" & mWMTK_Col & arrStaC(K).Row1 & ":Détail!" & mWMTK_Col & arrStaC(K).Row2 & ";" & Asc34 & lFct & Asc34 & ")"
            End Select
        End If
    Next K2
    wsExcel.Cells(mXls1_Row, 3).FormulaLocal = "=SOMME(E" & mXls1_Row & ":" & Mid$("ABCDEFGHIJKLMNOPQRSTUVW", mXls1_Cols - 2, 1) & mXls1_Row & ")"
    If arrStaP_Nb > 0 Then wsExcel.Cells(mXls1_Row, mXls1_Cols - 1).FormulaLocal = "=SOMME(E" & mXls1_Row & ":" & Mid$("ABCDEFGHIJKLMNOPQRSTUVW", mXls1_Cols - 3, 1) & mXls1_Row & ")"
    wsExcel.Cells(mXls1_Row, 3).Interior.Color = mColor_G0
    wsExcel.Cells(mXls1_Row, mXls1_Cols - 2).Interior.Color = mColor_G0
    wsExcel.Cells(mXls1_Row, mXls1_Cols - 1).Interior.Color = mColor_G0
Next K
    
'___________________________________________________________________________
If mRow1 > 0 Then
    wsExcel.Cells(mRow1, 1).Interior.Color = mColor_G2
    wsExcel.Cells(mRow1, 2).Interior.Color = mColor_G2
    For K2 = 3 To mXls1_Cols - 1
        X = Mid$("ABCDEFGHIJKLMNOPQRSTUVW", K2, 1)
        wsExcel.Cells(mRow1, K2).Interior.Color = mColor_G2
        wsExcel.Cells(mRow1, K2).FormulaLocal = "=SOMME(" & X & mRow1 + 1 & ":" & X & mXls1_Row & ")"
        wsExcel.Cells(mRow1, K2).Font.Bold = True
        wsExcel.Cells(mRow1, K2).Font.Color = vbBlue
    Next K2
End If

'===================================================================================

Call lstErr_AddItem(lstErr, cmdContext, "BIA_GOS : Récapitulatif REJETS"): DoEvents

mXls1_Row = mXls1_Row + 2
wsExcel.Cells(mXls1_Row, 2) = "Motif du rejet " & lFct
For K = 1 To 4
    wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_GB
    wsExcel.Cells(mXls1_Row, K).Font.Color = vbWhite
Next K

Select Case lFct
    Case "": X = ""
    Case "=MT 700": X = " and GOSDOSWMTK = '700'"
    Case "<>MT 700": X = " and GOSDOSWMTK <> '700'"
End Select

X = "select GOSDOSLABK , count(*)  from " & paramIBM_Library_SABSPE & ".YGOSDOS0 " _
     & " where GOSDOSIDD > 0 and GOSDOSISRV ='" & currentSSIWINUNIT & "' and GOSDOSSTAD <> 'A'" _
     & " and GOSDOSIAMJ >= " & wAmjMin & " and GOSDOSIAMJ <= " & wAmjMax _
     & " and GOSDOSSTAG = 'R'" & X _
     & " group by  GOSDOSLABK order by  GOSDOSLABK"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    mXls1_Row = mXls1_Row + 1
    wsExcel.Cells(mXls1_Row, 2) = rsSab("GOSDOSLABK")
    wsExcel.Cells(mXls1_Row, 2).Interior.Color = mColor_W0
    wsExcel.Cells(mXls1_Row, 4) = rsSab(1)
    wsExcel.Cells(mXls1_Row, 4).Font.Color = vbRed
    wsExcel.Cells(mXls1_Row, 4).Font.Bold = True
    
    rsSab.MoveNext
Loop
'======================================================================================================

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub



Public Sub fraDetail_LAB_Display_Dossier()
            
newYGOSDOS0 = oldYGOSDOS0

fraDetail_LAB_Display

If currentSSIWINUNIT = oldYGOSDOS0.GOSDOSISRV _
Or currentSSIWINUNIT = oldYGOSDOS0.GOSDOSGSRV Then
    cmdList_Add.Visible = arrHab(2)
    cmdList_Display.Visible = Not cmdList_Add.Visible
Else
    cmdList_Add.Visible = False
    cmdList_Display.Visible = True
End If

End Sub

Public Sub fraList_Display_Habilitation()

If currentSSIWINUNIT <> Mid$(cboList_SWISABKSRV, 1, 3) Then
    cmdList_SAB_Annulation.Visible = False
    cmdList_SAB_Modification.Visible = False
    cmdList_Ignore.Visible = False
    cmdList_Add.Visible = False
    cmdList_Display.Visible = False
    cmdList_New.Visible = False
    cmdList_SWISABKSRV.Visible = True
Else
    cmdList_New.Visible = arrHab(5)
    cmdList_Ignore.Visible = arrHab(13)
    cmdList_SWISABKSRV.Visible = arrHab(13)
    cboList_SWISABKSRV.Enabled = arrHab(13)

End If

End Sub

Public Sub fgSelect_Display_YSWISAB0_YSWILNK0()
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
    
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine_YSWISAB0 0
        
        rsSab.MoveNext
    
    Loop

    
    rsSabX.MoveNext
Loop

If fgSelect.Rows > 2 Then fgSelect_Sort1 = 11: fgSelect_Sort2 = 11: fgSelect_Sort

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub cmdSelect_Reset_Post_Update()
If SSTab1.Tab = 0 Then
    If cmdSelect_SQL_K = "3" Then
       
        oldYGOSDOS0 = newYGOSDOS0
        arrYGOSDOS0(arrYGOSDOS0_Index) = newYGOSDOS0
        xYGOSDOS0 = newYGOSDOS0
        
        Mesg_aid = newYGOSDOS0.GOSDOSWID1
        mesg_s_umidl = newYGOSDOS0.GOSDOSWIDL
        mesg_s_umidh = newYGOSDOS0.GOSDOSWIDH

        fgSelect_DisplayLine_YGOSDOS0 arrYGOSDOS0_Index
        fgDetail_Display
    Else
        fraDetail.Visible = False
        If cmdSelect_SQL_K = "2" Then cmdSelect_SQL_1b
        If cmdSelect_SQL_K = "5" Then cmdSelect_SQL_5
        If cmdSelect_SQL_K = "9" Then cmdSelect_SQL_9
        If cmdSelect_SQL_K = "9+" Then cmdSelect_SQL_9M
        If cmdSelect_SQL_K = "2d" Then txtSelect_3_GOSDOSIDD = newYGOSDOS0.GOSDOSIDD: cboSelect_SQL.ListIndex = mSelect_SQL_Listindex_3
    End If
End If

End Sub

Public Sub Importation_SIDE_Reporting_Control()
'last_Alerte_Loop_EVE
Dim xSubject As String, xMsg As String, T As String, wSS_Max As Long

T = Time
If T > "08:30:00" And T < "19:00:00" Then
    If T < "17:00:00" Then
        wSS_Max = 2500
    Else
        wSS_Max = 3600
    End If
    
    If DateDiff("s", last_Jrnl_date_time_EVE, CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now) & " " & T)) > wSS_Max Then
        If last_Alerte_date_time_EVE < last_Jrnl_date_time_EVE Then
            last_Alerte_Loop_EVE = last_Alerte_Loop_EVE + 1
            If last_Alerte_Loop_EVE > 5 Then
                blnAlerte_date_time_EVE = True
                last_Alerte_Loop_EVE = 0
                last_Alerte_date_time_EVE = last_Jrnl_date_time_EVE
                xSubject = "SIDE_Reporting : Aucun événement dans le journal depuis " & Fix(wSS_Max / 60) & " minutes "
                xMsg = "<body bgcolor= #FFA07A><CENTER>" & "<Font color = #0000FF><B>" _
                     & xSubject _
                     & " <BR><BR><Font color = #FF0000>La base de données SIDE_Reporting n'est peut_être plus synchronisée avec SAA" _
                     & "<BR><BR><Font color = #0000FF>" & "Dernier événement (rJrnl) reçu de la plateforme SAA dans SIDE_Reporting: " & last_Jrnl_date_time_EVE
                Call cmdSendMail_SAA_Alerte("SAA_Synchronisation", xSubject, xMsg, "S40", "S97")
            End If
        End If
    Else
        last_Alerte_Loop_EVE = 0
        If blnAlerte_date_time_EVE Then
            blnAlerte_date_time_EVE = False
            xSubject = "La base de données SIDE_Reporting est à nouveau synchronisée avec SAA"
            xMsg = "<body bgcolor= #C0FFC0><CENTER>" & "<Font color = #000000><B>" _
                     & xSubject & "<BR><BR>" & "Dernier événement (rJrnl) reçu de la plateforme SAA dans SIDE_Reporting: " & last_Jrnl_date_time_EVE
            Call cmdSendMail_SAA_Alerte("SAA_Synchronisation", xSubject, xMsg, "S40", "S97")
        End If
    End If
End If


If DateDiff("s", last_mesg_crea_date_time, last_Jrnl_date_time_ES) > 600 Then
    If last_Alerte_date_time_ES < last_mesg_crea_date_time Then
        last_Alerte_Loop_ES = last_Alerte_Loop_ES + 1
        If last_Alerte_Loop_ES > 5 Then
            blnAlerte_date_time_ES = True
            last_Alerte_Loop_ES = 0
            last_Alerte_date_time_ES = last_mesg_crea_date_time
            xSubject = "La table rMesg de SIDE_Reporting n'est plus synchronisée avec SAA"
            xMsg = "<body bgcolor= #FF0000><CENTER>" & "<Font color = #FFFF00><B>" _
                 & xSubject & "<BR><BR>" & "Dernier message reçu sur la plateforme SAA : " & last_Jrnl_date_time_ES _
                 & "<BR><BR>" & "Dernier message intégré dans SIDE_Reporting : " & last_mesg_crea_date_time
            Call cmdSendMail_SAA_Alerte("SAA_Synchronisation", xSubject, xMsg, "S40", "S97")
        End If
    End If
Else
    last_Alerte_Loop_ES = 0
    If blnAlerte_date_time_ES Then
        blnAlerte_date_time_ES = False
        xSubject = "La table rMesg de SIDE_Reporting est à nouveau synchronisée avec SAA"
        xMsg = "<body bgcolor= #C0FFC0><CENTER>" & "<Font color = #0000FF><B>" _
             & xSubject & "<BR><BR>" & "Dernier message reçu sur la plateforme SAA : " & last_Jrnl_date_time_ES _
             & "<BR><BR>" & "Dernier message intégré dans SIDE_Reporting : " & last_mesg_crea_date_time
        Call cmdSendMail_SAA_Alerte("SAA_Synchronisation", xSubject, xMsg, "S40", "S97")
    End If
End If
End Sub



Public Sub cmdSelect_SQL_Stat_BIC_Recap()
Call cmdSelect_SQL_Stat_BIC_Init_1("")
End Sub

Public Sub lstMail_MT_To_TXT()
Dim K As Integer, X As String
X = ""
For K = 0 To lstMail_MT_To.ListCount - 1


    If lstMail_MT_To.Selected(K) Then
        If Mid$(lstMail_MT_To.List(K), 1, 2) = ".S" Then
            lstMail_MT_To.Selected(K) = False
            Call lstMail_MT_To_TXT_S(Mid$(lstMail_MT_To.List(K), 1, 4))
        Else
            X = X & lstMail_MT_To.List(K) & ";"
        End If
        
    End If
    
Next K
If X <> "" Then Mid$(X, Len(X), 1) = " "

txtMail_MT_To = X

End Sub

Public Sub lstMail_MT_To_TXT_S(lSRV As String)
Dim K As Integer, X As String
Dim K1 As Integer, K2 As Integer, kMax As Integer, blnOk As Boolean, xSSIMELINFO As String
X = "select  SSIMELINFO from " & paramIBM_Library_SABSPE & ".YSSIMEL0 where SSIMELNAT = '@'" _
     & " and SSIMELUIDX = 'BIA_GOS" & lSRV & "'"
Set rsSabX = cnsab.Execute(X)
If rsSabX.EOF Then
    xSSIMELINFO = ""
Else
    xSSIMELINFO = rsSabX("SSIMELINFO")
End If
kMax = Len(X)
K1 = 1
blnOk = True

Do
    K2 = InStr(K1, xSSIMELINFO, ".")
    If K2 > 0 Then
        X = UCase$(Mid$(xSSIMELINFO, K1, K2 - K1))
        K1 = InStr(K2, xSSIMELINFO, ";")
        If K1 > 0 Then
            K1 = K1 + 1
        Else
            blnOk = False
        End If
            
        For K = 0 To lstMail_MT_To.ListCount - 1
            If lstMail_MT_To.List(K) = X Then
                    If Not lstMail_MT_To.Selected(K) Then lstMail_MT_To.Selected(K) = True
            End If
        Next K
    Else
        blnOk = False
    End If
    
Loop Until blnOk = False
End Sub

Public Sub lstMail_MT_CC_TXT()
Dim K As Integer, X As String
X = ""
For K = 0 To lstMail_MT_CC.ListCount - 1

    If lstMail_MT_CC.Selected(K) Then X = X & lstMail_MT_CC.List(K) & ";"
Next K
If X <> "" Then Mid$(X, Len(X), 1) = " "

txtMail_MT_CC = X

End Sub

Public Function YSWIRAM0_Update()
Dim K As Integer, xSql As String, X As String
Dim V
On Error GoTo Error_Handler

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

'________________________________________________________________________________

Select Case mYSWIRAM0_Fct
    Case "Update": V = sqlYSWIRAM0_Update(newYSWIRAM0, oldYSWIRAM0)
    Case "New": V = sqlYSWIRAM0_Insert(newYSWIRAM0)
    Case "Delete": V = sqlYSWIRAM0_Delete(oldYSWIRAM0)
    Case "SQL": V = sqlYSWIRAM0_Update_Field(newYSWIRAM0, mYSWIRAM0_SQL_Set)
End Select

If Not IsNull(V) Then GoTo Error_MsgBox

If mYSWIRAM0_Fct = "STA_Reset" Then
    Do While Not rsSabX.EOF
    
        Call rsYSWIRAM0_GetBuffer(rsSabX, oldYSWIRAM0)
        Call YSWIRAM0_STA_Reset_New
        V = sqlYSWIRAM0_Update(newYSWIRAM0, oldYSWIRAM0)
        If Not IsNull(V) Then GoTo Error_MsgBox
        
        If oldYSWIRAM0.SWIRAMXES = "E" Then
            X = " set SWISABSER = ''" _
              & " , SWISABSSE = ''" _
              & " , SWISABOPEC = ''" _
              & " , SWISABOPEN = 0"
        
            V = sqlYSWISAB0_Update_Field(newYSWIRAM0.SWIRAMXID, X)
            If Not IsNull(V) Then GoTo Error_MsgBox
        End If
        rsSabX.MoveNext
    
    Loop
    
End If

If mYSWIRAM0_Match_XOPE <> "" Then
    X = " set SWISABSER = '" & Mid$(mYSWIRAM0_Match_XOPE, 1, 2) & "'" _
      & " , SWISABSSE = '" & Mid$(mYSWIRAM0_Match_XOPE, 3, 2) & "'" _
      & " , SWISABOPEC = '" & Mid$(mYSWIRAM0_Match_XOPE, 5, 3) & "'" _
      & " , SWISABOPEN =  " & Mid$(mYSWIRAM0_Match_XOPE, 8, 9)

    V = sqlYSWISAB0_Update_Field(newYSWIRAM0.SWIRAMXID, X)
    If Not IsNull(V) Then GoTo Error_MsgBox
End If

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : YSWIRAM0_Update"
Exit_sub:

    YSWIRAM0_Update = V
    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
           
    End If
    
    mYSWIRAM0_Fct = ""
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    

End Function

Public Sub YSWIRAM0_Match_300()
Dim V, X As String, K As Long, K2 As Long, Nb As Integer
Dim xSql As String
On Error GoTo Error_Handler


blnYSWIRAM0_Match_CONF = False
If oldYSWISAB0.SWISABWES = "E" Then
    X = " and SWIRAMXES = 'S'"
Else
    X = " and SWIRAMXES = 'E'"
End If

'If InStr(oldYSWISAB0.SWISABWBIC, "EMIDITMM") > 0 Then
'    xYSWIRAM0.SWIRAMSTA = "I"
'Else
    'If blnReprise Then
    '    If oldYSWISAB0.SWISABWMTK = "300" And InStr(oldYSWISAB0.SWISABWBIC, "BSAH") > 0 Then xYSWIRAM0.SWIRAMSTA = "I"
    'End If
'End If


If xYSWIRAM0.SWIRAMSTA <> "I" Then
'____________________________________________________________________________________________________________

    If xYSWIRAM0.SWIRAMSTA = "?" Then
       xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIRAM0S , " & paramIBM_Library_SABSPE & ".YSWISAB0" _
             & " where SWIRAMSTA = ' ' and SWIRAMXBIC = '" & xYSWIRAM0.SWIRAMXBIC & "' and SWIRAMXREF = '" & xYSWIRAM0.SWIRAMXREF & "'" _
             & " and SWIRAMXMTK = '" & xYSWIRAM0.SWIRAMXMTK & "'" _
             & " and SWISABSWID = SWIRAMXID and SWIRAMXES = 'S'"
    Else
       xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIRAM0S , " & paramIBM_Library_SABSPE & ".YSWISAB0" _
             & " where SWIRAMSTA = '#' and SWIRAMXBIC = '" & xYSWIRAM0.SWIRAMXBIC & "' and SWIRAMXREF = '" & xYSWIRAM0.SWIRAMXREF & "'" _
             & " and SWIRAMXMTK = '" & xYSWIRAM0.SWIRAMXMTK & "'" _
             & " and SWISABSWID = SWIRAMXID" & X
    End If
    
    Set rsSabX = cnsab.Execute(xSql)
    
    Do While Not rsSabX.EOF
    
        Call YSWIRAM0_Fields(rsSabX("SWISABWID1"), rsSabX("SWISABWIDL"), rsSabX("SWISABWIDH"))
        
        Call YSWIRAM0_Match_fgDetail

        If blnYSWIRAM0_Match_CONF Then Exit Do
        
        rsSabX.MoveNext
    
    Loop
'____________________________________________________________________________________________________________
End If

mYSWIRAM0_Match_XOPE = ""

If blnYSWIRAM0_Match_Retry Then
    If Not blnYSWIRAM0_Match_CONF Then GoTo Exit_sub
    Call rsYSWIRAM0_GetBuffer(rsSab, oldYSWIRAM0)
    newYSWIRAM0 = oldYSWIRAM0
    newYSWIRAM0.SWIRAMXOPE = rsSabX("SWIRAMXOPE")
    newYSWIRAM0.SWIRAMSTA = " "
    If Not blnReprise Then
        newYSWIRAM0.SWIRAMYUSR = usrName_UCase
        newYSWIRAM0.SWIRAMYAMJ = DSys
        newYSWIRAM0.SWIRAMYHMS = time_Hms
    End If
    mYSWIRAM0_Fct = "Update"
    mYSWIRAM0_Match_XOPE = newYSWIRAM0.SWIRAMXOPE
    Call YSWIRAM0_Update
    GoTo Exit_sub
End If


If Not blnYSWIRAM0_Match_CONF Then
    mYSWIRAM0_Fct = "New": newYSWIRAM0 = xYSWIRAM0
    Call YSWIRAM0_Update
Else
    If xYSWIRAM0.SWIRAMSTA = "#" Then
        
        Call rsYSWIRAM0_GetBuffer(rsSabX, oldYSWIRAM0)
        newYSWIRAM0 = oldYSWIRAM0
        If Trim(newYSWIRAM0.SWIRAMXOPE) = "" Then newYSWIRAM0.SWIRAMXOPE = xYSWIRAM0.SWIRAMXOPE: mYSWIRAM0_Match_XOPE = newYSWIRAM0.SWIRAMXOPE
        newYSWIRAM0.SWIRAMSTA = " "
        If Not blnReprise Then
            newYSWIRAM0.SWIRAMYUSR = usrName_UCase
            newYSWIRAM0.SWIRAMYAMJ = DSys
            newYSWIRAM0.SWIRAMYHMS = time_Hms
        End If
        mYSWIRAM0_Fct = "Update"
        Call YSWIRAM0_Update
        
        xYSWIRAM0.SWIRAMSTA = " "
        If Trim(xYSWIRAM0.SWIRAMXOPE) = "" Then xYSWIRAM0.SWIRAMXOPE = oldYSWIRAM0.SWIRAMXOPE: mYSWIRAM0_Match_XOPE = newYSWIRAM0.SWIRAMXOPE
        mYSWIRAM0_Fct = "New": newYSWIRAM0 = xYSWIRAM0
        If Not blnReprise Then
            newYSWIRAM0.SWIRAMYUSR = usrName_UCase
            newYSWIRAM0.SWIRAMYAMJ = DSys
            newYSWIRAM0.SWIRAMYHMS = time_Hms
        End If
        Call YSWIRAM0_Update
    Else
        Call rsYSWIRAM0_GetBuffer(rsSabX, newYSWIRAM0)
        newYSWIRAM0.SWIRAMXOPE = rsSabX("SWIRAMXOPE")
        newYSWIRAM0.SWIRAMSTA = " "
        If Not blnReprise Then
            newYSWIRAM0.SWIRAMYUSR = usrName_UCase
            newYSWIRAM0.SWIRAMYAMJ = DSys
            newYSWIRAM0.SWIRAMYHMS = time_Hms
        End If
        mYSWIRAM0_Fct = "New": newYSWIRAM0 = xYSWIRAM0
        mYSWIRAM0_Match_XOPE = newYSWIRAM0.SWIRAMXOPE
        Call YSWIRAM0_Update
   End If
    
End If

GoTo Exit_sub

'____________________________________________________________________________________________________________
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

End Sub

Public Sub YSWIRAM0_Match_fgDetail()
Dim K As Long, K2 As Long, bln57A_Err As Boolean
On Error GoTo Error_Handler

For K2 = 0 To 100
    arrYSWIRAM0_Fields_X1(K2) = ""
    arrYSWIRAM0_Fields_X2(K2) = ""
Next K2

'If oldYSWISAB0.SWISABSWID = 1167621 Or oldYSWISAB0.SWISABSWID = 1167674 Then
'    Debug.Print "YSWIRAM0_Match_fgDetail"
'End If

mMTK = oldYSWISAB0.SWISABWMTK

mMTK_Seq = "": bln57A_Err = False
If oldYSWISAB0.SWISABWES = "S" Then wField_57A = ""

blnYSWIRAM0_Match_CONF = True
For K = 0 To fgDetail.Rows - 1
    fgDetail.Row = K
    fgDetail.Col = 0: xField_K = fgDetail.Text
    If xField_K <> "" Then
        fgDetail.Col = 1: xField_V = fgDetail.Text
        Call YSWIRAM0_Field_V
        Select Case oldYSWISAB0.SWISABWMTK
        'ignorer le champ 30T (incorrect pour BIA
            Case "300":
                Select Case xField_K
                    Case "82A": xField_K2 = "87A"
                    Case "87A": xField_K2 = "82A"
                    Case "32B-B1": xField_K2 = "33B-B2"
                    Case "33B-B2": xField_K2 = "32B-B1"
                    'Case "53A-B1": xField_K2 = "53A-B2"
                    Case "57A-B1": xField_K2 = "57A-B2"
                    'Case "53A-B2": xField_K2 = "53A-B1"
                    Case "57A-B2": xField_K2 = "57A-B1"
                                  If oldYSWISAB0.SWISABWES = "S" Then wField_57A = xField_V
                  Case "22A", "22C", "30V", "36": xField_K2 = xField_K
                    Case Else: xField_K2 = ""
                 End Select
           Case "320":
                Select Case xField_K
                    Case "82A": xField_K2 = "87A"
                    Case "87A": xField_K2 = "82A"
                    Case "53A-C": xField_K2 = "57A-C"
                    Case "57A-C": xField_K2 = "53A-C"
                                If oldYSWISAB0.SWISABWES = "S" Then wField_57A = xField_V
    
                    Case "22B": xField_K2 = xField_K
                                If xField_V = "ROLL" Then xField_V = "CONF"
                    Case "14D":  xField_K2 = IIf(blnReprise, "", xField_K)
                    Case "17R", "22A", "22C", "30V", "30P", "32B", "30X", "34E", "37G": xField_K2 = xField_K
                    Case Else: xField_K2 = ""
                End Select
       End Select
       If xField_K2 <> "" Then
            arrYSWIRAM0_Fields_X2(K) = "#"
            For K2 = 0 To arrYSWIRAM0_Fields_Nb1
                If xField_K2 = arrYSWIRAM0_Fields_K1(K2) Then
                    If xField_V = arrYSWIRAM0_Fields_V1(K2) Then
                        arrYSWIRAM0_Fields_X1(K2) = "="
                        arrYSWIRAM0_Fields_X2(K) = "="
                        Exit For
                    Else
                        If xField_K2 = "53A-C" Or xField_K2 = "57A-C" Or xField_K2 = "57A-B1" Or xField_K2 = "57A-B2" Then
                            If Not blnReprise Then bln57A_Err = True
                            arrYSWIRAM0_Fields_X1(K2) = "*"
                            arrYSWIRAM0_Fields_X2(K) = "*"
                        Else
                            blnYSWIRAM0_Match_CONF = False
                            arrYSWIRAM0_Fields_X1(K2) = "#"
                            arrYSWIRAM0_Fields_X2(K) = "#"
                        End If
                        Exit For
                    End If
                    
                End If
            Next K2
       End If
    End If
   'If Not blnYSWIRAM0_Match_CONF Then Exit For
Next K
        
If bln57A_Err And wField_57A <> "" Then
    If InStr(wField_57A, "BIARFRPP") = 0 Then blnYSWIRAM0_Match_CONF = False
End If
'____________________________________________________________________________________________________________

GoTo Exit_sub

'____________________________________________________________________________________________________________
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : YSWIRAM0_Match_fgDetail"
Exit_sub:

End Sub


Public Sub YSWIRAM0_Field_V()

On Error GoTo Error_Handler

If xField_K = "82A" Or xField_K = "87A" _
Or xField_K = "53A" Or xField_K = "57A" Then
    If Mid$(fgDetail.Text, 1, 1) = "/" Then fgDetail.Row = fgDetail.Row + 1
    xField_V = Mid$(fgDetail.Text, 1, 8)
Else
    xField_V = fgDetail.Text
End If

If xField_K = "32A" Or xField_K = "32B" Or xField_K = "33A" Or xField_K = "33B" Or xField_K = "34E" Then
    xField_V = Replace(xField_V, ",00", ",")
End If

If xField_K = "34E" Or xField_K = "34E" Then
    If Mid$(xField_V, 1, 1) = "N" Then xField_V = Mid$(xField_V, 2, Len(xField_V) - 1)
End If
If xField_K = "36" Then xField_V = CCur(xField_V)

If xField_K = "37G" Then
    If Mid$(xField_V, 1, 1) = "N" Then xField_V = Mid$(xField_V, 2, Len(xField_V) - 1)
    xField_V = CCur(xField_V)
End If

Select Case mMTK
    Case "300"
        Select Case xField_K
            Case "32B": mMTK_Seq = "-B1"
            Case "33B": mMTK_Seq = "-B2"
        End Select
        xField_K = xField_K & mMTK_Seq
    Case "320"
        Select Case xField_K
            Case "15C": mMTK_Seq = "-C"
            Case "15D": mMTK_Seq = "-D"
        End Select
        xField_K = xField_K & mMTK_Seq
End Select
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    Dim V
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
Exit_sub:

End Sub

Public Sub YSWIRAM0_Importation_fgDetail()
Dim K As Long

On Error GoTo Error_Handler

For K = 0 To 100
    arrYSWIRAM0_Fields_K1(K) = ""
    arrYSWIRAM0_Fields_V1(K) = ""
    arrYSWIRAM0_Fields_X1(K) = ""
    arrYSWIRAM0_Fields_X2(K) = ""
Next K

mMTK_Seq = "": wField_57A = ""
arrYSWIRAM0_Fields_Nb1 = 0

 For K = 0 To fgDetail.Rows - 1
     fgDetail.Row = K
     fgDetail.Col = 0: xField_K = fgDetail.Text
     fgDetail.Col = 1: xField_V = fgDetail.Text
     Call YSWIRAM0_Field_V
     
     Select Case oldYSWISAB0.SWISABWMTK
         Case "300":
             If xField_K = "22C" Then xYSWIRAM0.SWIRAMXREF = xField_V
             If xField_K = "22A" Then
                 Select Case xField_V
                     Case "NEWT": xYSWIRAM0.SWIRAMX22 = 1
                     Case "AMND": xYSWIRAM0.SWIRAMX22 = 4: xYSWIRAM0.SWIRAMSTA = "?": xField_V = "NEWT"
                     Case "CANC": xYSWIRAM0.SWIRAMX22 = 5: xYSWIRAM0.SWIRAMSTA = "?": xField_V = "NEWT"
                     Case Else: xYSWIRAM0.SWIRAMX22 = 6: xYSWIRAM0.SWIRAMSTA = "?": xField_V = "NEWT"
                 End Select
             End If
             If xField_K = "57A-B2" And oldYSWISAB0.SWISABWES = "S" Then wField_57A = xField_V
             
         Case "320":
             If xField_K = "22C" Then xYSWIRAM0.SWIRAMXREF = xField_V
             If xField_K = "22B" Then
                 Select Case xField_V
                     Case "CONF": xYSWIRAM0.SWIRAMX22 = 1
                     Case "ROLL": xYSWIRAM0.SWIRAMX22 = 3: xField_V = "CONF"
                     '"MATU"
                     Case Else: xYSWIRAM0.SWIRAMX22 = 2: xYSWIRAM0.SWIRAMSTA = "?": xField_V = "CONF"
                 End Select
             End If
             If xField_K = "22A" Then
                 Select Case xField_V
                     Case "NEWT"
                     Case "AMND": xYSWIRAM0.SWIRAMX22 = 4: xYSWIRAM0.SWIRAMSTA = "?": xField_V = "NEWT"
                     Case "CANC": xYSWIRAM0.SWIRAMX22 = 5: xYSWIRAM0.SWIRAMSTA = "?": xField_V = "NEWT"
                     Case Else: xYSWIRAM0.SWIRAMX22 = 6: xYSWIRAM0.SWIRAMSTA = "?": xField_V = "NEWT"
                 End Select
             End If
             If xField_K = "17R" Then xField_V = IIf(xField_V = "B", "L", "B")
             If xField_K = "57A-C" And oldYSWISAB0.SWISABWES = "S" Then wField_57A = xField_V
             
         Case "202", "210":
             If xField_K = "20" Then xYSWIRAM0.SWIRAMXREF = xField_V
         Case "900", "910":
             If xField_K = "21" Then xYSWIRAM0.SWIRAMXREF = xField_V
     End Select
     
     arrYSWIRAM0_Fields_Nb1 = arrYSWIRAM0_Fields_Nb1 + 1
     arrYSWIRAM0_Fields_K1(arrYSWIRAM0_Fields_Nb1) = xField_K
     arrYSWIRAM0_Fields_V1(arrYSWIRAM0_Fields_Nb1) = xField_V
     arrYSWIRAM0_Fields_X1(arrYSWIRAM0_Fields_Nb1) = ""
Next K
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : YSWIRAM0_Importation_fgDetail"
Exit_sub:

End Sub


Public Sub YSWIRAM0_Match_6E_Fct(lFct As String)
Dim xSet As String
On Error GoTo Error_Handler

xSet = " , SWIRAMYUPD = 'M' , SWIRAMYAMJ = " & DSys & " , SWIRAMYHMS = " & time_Hms & " , SWIRAMYUSR = '" & usrName_UCase & "'"
Select Case lFct
    Case "I1"
        X = MsgBox("Voulez-vous ignorer ce message ?", vbYesNo, "YSWIRAM0: Echéancier")
        If X = vbYes Then
            mYSWIRAM0_Match_XOPE = ""
            fraSwift.Visible = False
            mYSWIRAM0_Fct = "SQL"
            mYSWIRAM0_SQL_Set = " set SWIRAMSTA = 'I'" & xSet
            newYSWIRAM0.SWIRAMXID = oldYSWISAB0_1.SWISABSWID
            Call YSWIRAM0_Update
            If oldYSWISAB0_1.SWISABXGOS = "G" Then
                oldYSWISAB0 = oldYSWISAB0_1
                Call YSWIRAM0_YGOSDOS0_EVE(lFct)
            End If
            Call cmdSelect_SQL_6
        End If
    Case "I2"
        X = MsgBox("Voulez-vous ignorer ce message ?", vbYesNo, "YSWIRAM0: Echéancier")
        If X = vbYes Then
            mYSWIRAM0_Match_XOPE = ""
            fraSwift.Visible = False
            mYSWIRAM0_Fct = "SQL"
            mYSWIRAM0_SQL_Set = " set SWIRAMSTA = 'I'" & xSet
            newYSWIRAM0.SWIRAMXID = oldYSWISAB0_2.SWISABSWID
            Call YSWIRAM0_Update
            If oldYSWISAB0_2.SWISABXGOS = "G" Then
                oldYSWISAB0 = oldYSWISAB0_2
                Call YSWIRAM0_YGOSDOS0_EVE(lFct)
            End If
            Call cmdSelect_SQL_6
        End If
    Case "M"
        X = MsgBox("Voulez-vous rapprocher ces deux messages ?", vbYesNo, "YSWIRAM0: Echéancier")
        If X = vbYes Then
            If oldYSWISAB0_1.SWISABOPEN <> 0 Then
                xYSWIRAM0.SWIRAMXOPE = Trim(oldYSWISAB0_1.SWISABSER & oldYSWISAB0_1.SWISABSSE & oldYSWISAB0_1.SWISABOPEC & Format(oldYSWISAB0_1.SWISABOPEN, "000000000"))
            Else
                xYSWIRAM0.SWIRAMXOPE = Trim(oldYSWISAB0_2.SWISABSER & oldYSWISAB0_2.SWISABSSE & oldYSWISAB0_2.SWISABOPEC & Format(oldYSWISAB0_2.SWISABOPEN, "000000000"))
            End If
            mYSWIRAM0_Match_XOPE = xYSWIRAM0.SWIRAMXOPE
            
            fraSwift.Visible = False
            
            mYSWIRAM0_Fct = "SQL"
            mYSWIRAM0_SQL_Set = " set SWIRAMSTA = ' ' , SWIRAMXOPE = '" & xYSWIRAM0.SWIRAMXOPE & "'" & xSet
            newYSWIRAM0.SWIRAMXID = oldYSWISAB0_1.SWISABSWID
            Call YSWIRAM0_Update
            
            mYSWIRAM0_Fct = "SQL"
            mYSWIRAM0_SQL_Set = " set SWIRAMSTA = ' ' , SWIRAMXOPE = '" & xYSWIRAM0.SWIRAMXOPE & "'" & xSet
            newYSWIRAM0.SWIRAMXID = oldYSWISAB0_2.SWISABSWID
            Call YSWIRAM0_Update
            
            Call YSWIRAM0_YGOSDOS0_EVE(lFct)
            
               
            Call cmdSelect_SQL_6
        End If
    Case "L"
        X = MsgBox("Voulez-vous associer le message entrant avec notre message sortant ?", vbYesNo, "YSWIRAM0: Echéancier")
        If X = vbYes Then
            If oldYSWISAB0_1.SWISABOPEN <> 0 Then
                xYSWIRAM0.SWIRAMXOPE = Trim(oldYSWISAB0_1.SWISABSER & oldYSWISAB0_1.SWISABSSE & oldYSWISAB0_1.SWISABOPEC & Format(oldYSWISAB0_1.SWISABOPEN, "000000000"))
            Else
                xYSWIRAM0.SWIRAMXOPE = Trim(oldYSWISAB0_2.SWISABSER & oldYSWISAB0_2.SWISABSSE & oldYSWISAB0_2.SWISABOPEC & Format(oldYSWISAB0_2.SWISABOPEN, "000000000"))
            End If
            mYSWIRAM0_Match_XOPE = xYSWIRAM0.SWIRAMXOPE
            
            fraSwift.Visible = False
            
            If oldYSWISAB0_1.SWISABWES = "E" Then
               mYSWIRAM0_Fct = "SQL"
                mYSWIRAM0_SQL_Set = " set SWIRAMSTA = '%' , SWIRAMXOPE = '" & xYSWIRAM0.SWIRAMXOPE & "'" & xSet
                newYSWIRAM0.SWIRAMXID = oldYSWISAB0_1.SWISABSWID
                Call YSWIRAM0_Update
            Else
                mYSWIRAM0_Fct = "SQL"
                mYSWIRAM0_SQL_Set = " set SWIRAMSTA = '%' , SWIRAMXOPE = '" & xYSWIRAM0.SWIRAMXOPE & "'" & xSet
                newYSWIRAM0.SWIRAMXID = oldYSWISAB0_2.SWISABSWID
                Call YSWIRAM0_Update
            End If
            
            Call YSWIRAM0_YGOSDOS0_EVE(lFct)
            
            Call cmdSelect_SQL_6
        End If
    Case "GOS_New"
        
        rsYGOSDOS0_Init oldYGOSDOS0
        arrYGOSDOS0_Index = 1: ReDim arrYGOSDOS0(2)
        rsYGOSEVE0_Init oldYGOSEVE0
        Mesg_aid = oldYSWISAB0_1.SWISABWID1
        mesg_s_umidl = oldYSWISAB0_1.SWISABWIDL
        mesg_s_umidh = oldYSWISAB0_1.SWISABWIDH
        cmdSelect_SQL_K = "2-RAM"
        fraSwift.Visible = False
        fgDetail_Display
End Select

GoTo Exit_sub

'____________________________________________________________________________________________________________
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : YSWIRAM0_Match_6E_Fct"
Exit_sub:

End Sub
Public Sub YSWIRAM0_Match_6E()
Dim X As String, blnVisible As Boolean
On Error GoTo Error_Handler

Call YSWIRAM0_Fields(oldYSWISAB0.SWISABWID1, oldYSWISAB0.SWISABWIDL, oldYSWISAB0.SWISABWIDH)
If oldYSWISAB0_1.SWISABSWID = oldYSWISAB0.SWISABSWID Then mYSWIRAM0_Col = 0
If mYSWIRAM0_Col = 0 Then
    mMTK = oldYSWISAB0.SWISABWMTK
    oldYSWISAB0_1 = oldYSWISAB0
    Call rsYSWISAB0_Init(oldYSWISAB0_2)
    
    Call YSWIRAM0_Importation_fgDetail
Else
    oldYSWISAB0_2 = oldYSWISAB0
    YSWIRAM0_Match_fgDetail
End If

X = "select * from " & paramIBM_Library_SABSPE & ".YSWIRAM0  where SWIRAMXID = " & oldYSWISAB0.SWISABSWID
Set rsSab = cnsab.Execute(X)

If Not rsSab.EOF Then

    fgSwift.Rows = fgSwift.Rows + 5
    fgSwift.Row = fgSwift.Rows - 4
    fgSwift.Col = 0: fgSwift.Text = "Statut": fgSwift.Col = 1
    Select Case rsSab("SWIRAMSTA")
        Case " ": X = "Rapproché"
        Case "#": X = "CONF en attente"
        Case "?": X = "??? en attente"
        Case "I": X = "Ignoré"
        Case "A": X = "Annulé"
        Case Else: X = rsSab("SWIRAMsta")
    End Select
    Select Case rsSab("SWIRAMYUPD")
        Case " ":
                If rsSab("SWIRAMSTA") <> "#" And rsSab("SWIRAMSTA") <> "?" Then X = X & " automatiquement"
        Case "M": X = X & " manuellement"
        Case Else: X = X & rsSab("SWIRAMYUPD")
    End Select
    fgSwift.Text = X: fgSwift.CellForeColor = vbMagenta: fgSwift.CellBackColor = mColor_Y2
    fgSwift.Row = fgSwift.Rows - 3
    fgSwift.Col = 0: fgSwift.Text = "màj par ": fgSwift.Col = 1: fgSwift.Text = rsSab("SWIRAMYUSR")
    fgSwift.Row = fgSwift.Rows - 2
    fgSwift.Col = 0: fgSwift.Text = "le": fgSwift.Col = 1: fgSwift.Text = dateImp10_S(rsSab("SWIRAMYAMJ")) & "  " & timeImp8(rsSab("SWIRAMYHMS"))
End If
If arrHab(14) Then
    blnVisible = IIf(cmdSelect_SQL_K = "6E", True, False)
Else
    blnVisible = False
End If
Call frmYGOSDOS0_Param.Form_Init(fgSwift, blnVisible)
mYSWIRAM0_Col = 2

GoTo Exit_sub

'____________________________________________________________________________________________________________
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : YSWIRAM0_Match_6E"
Exit_sub:


End Sub


Public Sub YSWIRAM0_YGOSDOS0_EVE(lFct As String)
Dim V, X As String

On Error GoTo Error_Handler

If lFct = "M" Or lFct = "L" Then
    If oldYSWISAB0_1.SWISABXGOS = "G" Then
        oldYSWISAB0 = oldYSWISAB0_1
    Else
        If oldYSWISAB0_2.SWISABXGOS = "G" Then
           oldYSWISAB0 = oldYSWISAB0_2
           oldYSWISAB0_2 = oldYSWISAB0_1
        Else
            oldYSWISAB0.SWISABXGOS = ""
        End If
    End If
End If

If oldYSWISAB0.SWISABXGOS = "G" Then
    X = "select * from " & paramIBM_Library_SABSPE & ".YGOSDOS0 " _
         & " where GOSDOSWID1 = " & oldYSWISAB0.SWISABWID1 _
         & " and  GOSDOSWIDL = " & oldYSWISAB0.SWISABWIDL _
         & " and  GOSDOSWIDH = " & oldYSWISAB0.SWISABWIDH
    Set rsSab = cnsab.Execute(X)
    
    If rsSab.EOF Then
        Call MsgBox("Erreur de lecture : dossier non trouvé", vbError, "6E : YSWIRAM0_YGOSDOS0_EVE")
    Else
        Call rsYGOSDOS0_GetBuffer(rsSab, oldYGOSDOS0)
        If oldYGOSDOS0.GOSDOSSTAD = "C" Then
            Call MsgBox("Dossier " & oldYGOSDOS0.GOSDOSIDD & " déjà clôturé ", vbError, "6E : YSWIRAM0_YGOSDOS0_EVE")
            GoTo Exit_sub
        Else
            If lFct = "L" Then
                X = MsgBox("Voulez-vous mettre à jour le dossier " & oldYGOSDOS0.GOSDOSIDD, vbQuestion + vbYesNo, "6E - RAM : Dossier GOS en cours")
            Else
                X = MsgBox("Voulez-vous CLOTURER le dossier " & oldYGOSDOS0.GOSDOSIDD, vbQuestion + vbYesNo, "6E - RAM : Dossier GOS en cours")
            End If
            If X = vbYes Then 'Call MsgBox(" à faire")
                Call rsYGOSEVE0_Init(newYGOSEVE0)
                newYGOSEVE0.GOSEVEIDD = oldYGOSDOS0.GOSDOSIDD
                newYGOSEVE0.GOSEVEGSRV = oldYGOSDOS0.GOSDOSISRV
                If lFct = "M" Or lFct = "L" Then
                    newYGOSEVE0.GOSEVESWID = oldYSWISAB0_2.SWISABSWID
                    newYGOSEVE0.GOSEVENAT = "Swi+"
                    X = "   _:_           _:_                _:_                "
                    Mid$(X, 1, 3) = oldYSWISAB0_2.SWISABWMTK
                    Mid$(X, 7, 11) = oldYSWISAB0_2.SWISABWBIC
                    Mid$(X, 21, 16) = oldYSWISAB0_2.SWISABWL20
                    Mid$(X, 40, 16) = oldYSWISAB0_2.SWISABWN20
                    newYGOSEVE0.GOSEVETXT = X & " (6E-RAM)"
                    
                    V = cmdYGOSDOS0_Update("", "New", "", "", "")
                    If Not IsNull(V) Then GoTo Error_MsgBox
                End If
                
                newYGOSEVE0.GOSEVESWID = 0
                
                If lFct <> "L" Then
                    newYGOSDOS0 = oldYGOSDOS0
                    newYGOSDOS0.GOSDOSSTAD = "C"
                    newYGOSEVE0.GOSEVEIDE = 0
                    newYGOSEVE0.GOSEVESWID = 0
                    newYGOSEVE0.GOSEVENAT = "Clo"
                    Select Case lFct
                        Case "M": newYGOSEVE0.GOSEVETXT = "rapprochement validé"
                        Case "L": newYGOSEVE0.GOSEVETXT = "rapprochement partiel"
                        Case Else: newYGOSEVE0.GOSEVETXT = "message ignoré"
                    End Select
                    V = cmdYGOSDOS0_Update("Update", "New", "", "", "")
                    If Not IsNull(V) Then GoTo Error_MsgBox
                End If
                Call YSWIRAM0_YGOSDOS0_EVE_Mail(lFct)
                
            End If
        End If
    End If
End If

GoTo Exit_sub

'____________________________________________________________________________________________________________
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : YSWIRAM0_YGOSDOS0_EVE"
Exit_sub:

End Sub

Public Sub YSWIRAM0_YGOSDOS0_EVE_Mail(lFct As String)
Dim V, X As String, wTXT As String, Xto As String, xCC As String
Dim K1 As Integer, K2 As Integer

On Error GoTo Error_Handler

Xto = currentSSIWINMAIL
xCC = ""
If lFct = "L" Then
    txtMail_MT_Message = "RAM : rapprochement partiel"
Else
    txtMail_MT_Message = "RAM : Clôture du dossier"
End If
newYGOSEVE0.GOSEVETXT = "=> " & usrName & " + " & vbCrLf & vbCrLf & Trim(txtMail_MT_Message)

X = "select * from " & paramIBM_Library_SABSPE & ".YGOSEVE0 " _
     & " where GOSEVEIDD = " & oldYGOSDOS0.GOSDOSIDD _
     & " and  GOSEVENAT = 'Mail' order by GOSEVEIDE FETCH FIRST 1 ROWS ONLY"
Set rsSab = cnsab.Execute(X)

If Not rsSab.EOF Then
    'Call rsYGOSEVE0_GetBuffer(rsSab, oldYGOSEVE0)
    wTXT = rsSab("GOSEVETXT")
    K1 = InStr(wTXT, "+")
    If K1 > 0 Then
        X = Mid$(wTXT, 4, K1 - 5)
        V = mailAdresse_Production_Control(X, Xto)
        If Not IsNull(V) Then Xto = currentSSIWINMAIL

        K2 = InStr(K1, wTXT, vbCrLf)
        If K2 > 0 Then
            X = Mid$(wTXT, K1 + 1, K2 - K1)
            V = mailAdresse_Production_Control(X, xCC)
            If Not IsNull(V) Then xCC = ""
            newYGOSEVE0.GOSEVETXT = Mid$(wTXT, 1, K2 + 3) & txtMail_MT_Message
        End If
    End If
End If
newYGOSEVE0.GOSEVEIDE = 0
newYGOSEVE0.GOSEVENAT = "Mail"
V = cmdYGOSDOS0_Update("", "New", "", "", "")

If Not IsNull(V) Then GoTo Error_MsgBox

X = " where GOSEVEIDD = " & oldYGOSDOS0.GOSDOSIDD & " order by GOSEVEIDE"
arrYGOSEVE0_SQL X
newYGOSDOS0 = oldYGOSDOS0
Call cmdSendMail_YGOSDOS0(Xto, xCC)

GoTo Exit_sub

'____________________________________________________________________________________________________________
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : YSWIRAM0_YGOSDOS0_EVE"
Exit_sub:

End Sub

Public Sub YSWIRAM0_STA_Reset()
Dim xSql As String, Nb As Integer, K As Integer

On Error GoTo Error_Handler


xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIRAM0 " _
     & " where SWIRAMXID = " & oldYSWIRAM0.SWIRAMXID
Set rsSabX = cnsab.Execute(xSql)
If rsSabX.EOF Then
    V = "Erreur de lecture SWIRAMXID = " & oldYSWIRAM0.SWIRAMXID
    GoTo Error_MsgBox
End If

oldYSWIRAM0.SWIRAMXOPE = Trim(rsSabX("SWIRAMXOPE"))
If oldYSWIRAM0.SWIRAMXOPE = "" Then
    Nb = 1
Else
    xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YSWIRAM0 where SWIRAMXOPE = '" & oldYSWIRAM0.SWIRAMXOPE & "'"
    Set rsSabX = cnsab.Execute(xSql)
    Nb = rsSabX(0)
End If

If Not arrHab(14) Then GoTo Exit_sub

mYSWIRAM0_Match_XOPE = ""
Select Case Nb
    Case 1: X = MsgBox("Voulez-vous restaurer le statut de ce message ?", vbYesNo + vbQuestion, Me.Name & " : YSWIRAM0_STA_Reset")
            If X = vbYes Then
                Call rsYSWIRAM0_GetBuffer(rsSabX, oldYSWIRAM0)
                Call YSWIRAM0_STA_Reset_New
                mYSWIRAM0_Fct = "Update"
                V = YSWIRAM0_Update

            End If
    
    Case Is > 1: X = MsgBox("Voulez-vous restaurer le statut de ces " & Nb & " messages ?", vbYesNo + vbQuestion, Me.Name & " : YSWIRAM0_STA_Reset")
                If X = vbYes Then
                    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWIRAM0 where SWIRAMXOPE = '" & oldYSWIRAM0.SWIRAMXOPE & "'"
                    Set rsSabX = cnsab.Execute(xSql)
                    mYSWIRAM0_Fct = "STA_Reset"
                    V = YSWIRAM0_Update
               End If
    Case Else: X = vbNo
End Select
  
Call cmdSelect_SQL_6
'____________________________________________________________________________________________________________

GoTo Exit_sub

'____________________________________________________________________________________________________________
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : YSWIRAM0_STA_Reset"
Exit_sub:

End Sub

Public Sub YSWIRAM0_STA_Reset_New()
newYSWIRAM0 = oldYSWIRAM0
newYSWIRAM0.SWIRAMSTA = "#"
newYSWIRAM0.SWIRAMYUPD = "M"
If newYSWIRAM0.SWIRAMXES = "E" Then newYSWIRAM0.SWIRAMXOPE = ""
newYSWIRAM0.SWIRAMYUSR = usrName_UCase
newYSWIRAM0.SWIRAMYAMJ = DSys
newYSWIRAM0.SWIRAMYHMS = time_Hms


End Sub

Public Sub YSWIECH0_Importation_MT300_B1()
Dim blnBIARFRPP As Boolean, blnSOGEFRPP As Boolean

If wMT_53A_B1 <> "" Then
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_53A_B1, 1, 8)
Else
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_57A_B1, 1, 8)
End If
If xYSWIECH0.SWIECHWBIC = "BIARFRPP" Then blnBIARFRPP = True
If xYSWIECH0.SWIECHWBIC = wMT_87A Then blnBIARFRPP = True
If Mid$(xYSWIECH0.SWIECHWBIC, 1, 4) = "SOGE" Then blnSOGEFRPP = True

xYSWIECH0.SWIECHWDEV = wMT_32_DEV
xYSWIECH0.SWIECHWMTD = wMT_32_MTD

If blnBIARFRPP Then
'________________________________________________________________
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    xYSWIECH0.SWIECHSENS = "D"
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "900"
    newYSWIECH0.SWIECHWES = "S"
    newYSWIECH0.SWIECHWBIC = wMT_87A
    newYSWIECH0.SWIECHW52A = wMT_87A
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
    
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "950"
    newYSWIECH0.SWIECHWES = "S"
    newYSWIECH0.SWIECHWBIC = wMT_87A
    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
Else
    If blnSOGEFRPP Then
'________________________________________________________________
        xYSWIECH0.SWIECHSENS = "C"
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "202"
        newYSWIECH0.SWIECHWES = "E"
        newYSWIECH0.SWIECHWBIC = "SOGEFRPP"
        newYSWIECH0.SWIECHW52A = wMT_87A
        newYSWIECH0.SWIECHW57A = "SOGEFRPP"
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "950"
        newYSWIECH0.SWIECHWES = "E"
        newYSWIECH0.SWIECHWBIC = "SOGEFRPP"
        newYSWIECH0.SWIECHDECH = wMT_30V_JS1
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    Else
'________________________________________________________________
        xYSWIECH0.SWIECHSENS = "C"
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "210"
        newYSWIECH0.SWIECHWES = "S"
        newYSWIECH0.SWIECHW52A = wMT_87A
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "910"
        newYSWIECH0.SWIECHWES = "E"
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "950"
        newYSWIECH0.SWIECHWES = "E"
        newYSWIECH0.SWIECHDECH = wMT_30V_JS1
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    End If
End If
'________________________________________________________________

End Sub
Public Sub YSWIECH0_Importation_MT103_CDE()
Dim blnBIARFRPP As Boolean, blnSOGEFRPP As Boolean, xSql As String

xYSWIECH0.SWIECHWDEV = wMT_32_DEV
xYSWIECH0.SWIECHWMTD = wMT_32_MTD
xYSWIECH0.SWIECHDECH = wMT_30V
xYSWIECH0.SWIECHW30V = wMT_30V
wMT_30V_JS1 = dateElp("Ouvré", 1, wMT_30V)

xYSWIECH0.SWIECHWBIC = Mid$(oldYSWISAB0.SWISABWBIC, 1, 8)
'________________________________________________________________
If Mid$(xYSWIECH0.SWIECHWBIC, 1, 4) = "SOGE" Then
    xYSWIECH0.SWIECHSENS = "D"
        
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "198"
    newYSWIECH0.SWIECHWES = "E"
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
    
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "950"
    newYSWIECH0.SWIECHWES = "E"
    newYSWIECH0.SWIECHWBIC = "SOGEFRPP"
    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
Else
'________________________________________________________________
    If YSWI950_BIC_Nostro(xYSWIECH0.SWIECHWBIC) Then
        xYSWIECH0.SWIECHSENS = "D"
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "900"
        newYSWIECH0.SWIECHWES = "E"
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "950"
        newYSWIECH0.SWIECHWES = "E"
        newYSWIECH0.SWIECHDECH = wMT_30V_JS1
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    End If



    If wMT_57A <> "" Then
        xYSWIECH0.SWIECHWBIC = Mid$(wMT_57A, 1, 8)
    '________________________________________________________________
        If YSWI950_BIC_Loro(xYSWIECH0.SWIECHWBIC) Then
            xYSWIECH0.SWIECHSENS = "C"
            
            xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
            newYSWIECH0 = xYSWIECH0
            newYSWIECH0.SWIECHWMTK = "910"
            newYSWIECH0.SWIECHWES = "S"
            mYSWIECH0_Fct = "New"
            Call YSWIECH0_Update
            
            xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
            newYSWIECH0 = xYSWIECH0
            newYSWIECH0.SWIECHWMTK = "950"
            newYSWIECH0.SWIECHWES = "S"
            newYSWIECH0.SWIECHDECH = wMT_30V_JS1
            mYSWIECH0_Fct = "New"
            Call YSWIECH0_Update
        End If
    End If
End If


End Sub

Public Sub YSWIECH0_Importation_MT202()

xYSWIECH0.SWIECHWDEV = wMT_32_DEV
xYSWIECH0.SWIECHWMTD = wMT_32_MTD
xYSWIECH0.SWIECHDECH = wMT_30V
xYSWIECH0.SWIECHW30V = wMT_30V
wMT_30V_JS1 = dateElp("Ouvré", 1, wMT_30V)

xYSWIECH0.SWIECHWBIC = Mid$(oldYSWISAB0.SWISABWBIC, 1, 8)
'________________________________________________________________
If Mid$(xYSWIECH0.SWIECHWBIC, 1, 4) = "SOGE" Then
    xYSWIECH0.SWIECHSENS = "D"
        
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "298"
    newYSWIECH0.SWIECHWES = "E"
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
    
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "950"
    newYSWIECH0.SWIECHWES = "E"
    newYSWIECH0.SWIECHWBIC = "SOGEFRPP"
    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
Else
'________________________________________________________________
    If YSWI950_BIC_Nostro(xYSWIECH0.SWIECHWBIC) Then
        xYSWIECH0.SWIECHSENS = "D"
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "900"
        newYSWIECH0.SWIECHWES = "E"
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "950"
        newYSWIECH0.SWIECHWES = "E"
        newYSWIECH0.SWIECHDECH = wMT_30V_JS1
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    End If
End If


If wMT_52A <> "" Then
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_52A, 1, 8)
'________________________________________________________________
    If YSWI950_BIC_Loro(xYSWIECH0.SWIECHWBIC) Then Call YSWIECH0_Importation_MT103_52A
    
End If



End Sub


Public Sub YSWIECH0_Importation_MT320_PRE_RBT()
Dim blnBIARFRPP As Boolean, blnSOGEFRPP As Boolean

If wMT_53A_B2 <> "" Then
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_53A_B2, 1, 8)
Else
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_57A_B2, 1, 8)
End If
If xYSWIECH0.SWIECHWBIC = "BIARFRPP" Then blnBIARFRPP = True
If xYSWIECH0.SWIECHWBIC = wMT_87A Then blnBIARFRPP = True
If Mid$(xYSWIECH0.SWIECHWBIC, 1, 4) = "SOGE" Then blnSOGEFRPP = True

xYSWIECH0.SWIECHWDEV = wMT_32_DEV
xYSWIECH0.SWIECHWMTD = wMT_32_MTD + wMT_34_MTD
xYSWIECH0.SWIECHDECH = wMT_30X
xYSWIECH0.SWIECHW30V = wMT_30X
wMT_30V_JS1 = dateElp("Ouvré", 1, wMT_30X)

If blnBIARFRPP Then
'________________________________________________________________
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    xYSWIECH0.SWIECHSENS = "D"
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "900"
    newYSWIECH0.SWIECHWES = "S"
    newYSWIECH0.SWIECHWBIC = wMT_87A
    newYSWIECH0.SWIECHW52A = wMT_87A
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
    
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "950"
    newYSWIECH0.SWIECHWES = "S"
    newYSWIECH0.SWIECHWBIC = wMT_87A
    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
Else
    If blnSOGEFRPP Then
'________________________________________________________________
        xYSWIECH0.SWIECHSENS = "C"
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "202"
        newYSWIECH0.SWIECHWES = "E"
        newYSWIECH0.SWIECHWBIC = "SOGEFRPP"
        newYSWIECH0.SWIECHW52A = wMT_87A
        newYSWIECH0.SWIECHW57A = "SOGEFRPP"
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "950"
        newYSWIECH0.SWIECHWES = "E"
        newYSWIECH0.SWIECHWBIC = "SOGEFRPP"
        newYSWIECH0.SWIECHDECH = wMT_30V_JS1
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    Else
'________________________________________________________________
        xYSWIECH0.SWIECHSENS = "C"
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "210"
        newYSWIECH0.SWIECHWES = "S"
        newYSWIECH0.SWIECHW52A = wMT_87A
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "910"
        newYSWIECH0.SWIECHWES = "E"
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "950"
        newYSWIECH0.SWIECHWES = "E"
        newYSWIECH0.SWIECHDECH = wMT_30V_JS1
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    End If
End If
'________________________________________________________________

End Sub

Public Sub YSWIECH0_Importation_MT320_EMP_MAD()
Dim blnBIARFRPP As Boolean, blnSOGEFRPP As Boolean

If wMT_53A_B2 <> "" Then
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_53A_B2, 1, 8)
Else
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_57A_B2, 1, 8)
End If
If xYSWIECH0.SWIECHWBIC = "BIARFRPP" Then blnBIARFRPP = True
If xYSWIECH0.SWIECHWBIC = wMT_87A Then blnBIARFRPP = True
If Mid$(xYSWIECH0.SWIECHWBIC, 1, 4) = "SOGE" Then blnSOGEFRPP = True

xYSWIECH0.SWIECHWDEV = wMT_32_DEV
xYSWIECH0.SWIECHWMTD = wMT_32_MTD

If blnBIARFRPP Then
'________________________________________________________________
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    xYSWIECH0.SWIECHSENS = "D"
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "900"
    newYSWIECH0.SWIECHWES = "S"
    newYSWIECH0.SWIECHWBIC = wMT_87A
    newYSWIECH0.SWIECHW52A = wMT_87A
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
    
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "950"
    newYSWIECH0.SWIECHWES = "S"
    newYSWIECH0.SWIECHWBIC = wMT_87A
    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
Else
    If blnSOGEFRPP Then
'________________________________________________________________
        xYSWIECH0.SWIECHSENS = "C"
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "202"
        newYSWIECH0.SWIECHWES = "E"
        newYSWIECH0.SWIECHWBIC = "SOGEFRPP"
        newYSWIECH0.SWIECHW52A = wMT_87A
        newYSWIECH0.SWIECHW57A = "SOGEFRPP"
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "950"
        newYSWIECH0.SWIECHWES = "E"
        newYSWIECH0.SWIECHWBIC = "SOGEFRPP"
        newYSWIECH0.SWIECHDECH = wMT_30V_JS1
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    Else
'________________________________________________________________
        xYSWIECH0.SWIECHSENS = "C"
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "210"
        newYSWIECH0.SWIECHWES = "S"
        newYSWIECH0.SWIECHW52A = wMT_87A
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "910"
        newYSWIECH0.SWIECHWES = "E"
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "950"
        newYSWIECH0.SWIECHWES = "E"
        newYSWIECH0.SWIECHDECH = wMT_30V_JS1
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    End If
End If
'________________________________________________________________

End Sub


Public Sub YSWIECH0_Importation_MT300_B2()
Dim blnBIARFRPP As Boolean

'If xYSWIECH0.SWIECHOPEN = 5505 Then
'    Debug.Print "YSWIECH0_Importation_MT300_B2"
'End If

If wMT_57A_B2 = "BIARFRPP" Then blnBIARFRPP = True

If wMT_53A_B2 <> "" Then
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_53A_B2, 1, 8)
Else
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_57A_B2, 1, 8)
End If

xYSWIECH0.SWIECHWDEV = wMT_33_DEV
xYSWIECH0.SWIECHWMTD = wMT_33_MTD

If blnBIARFRPP Then
'________________________________________________________________
     xYSWIECH0.SWIECHSENS = "C"
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "910"
    newYSWIECH0.SWIECHWES = "S"
    newYSWIECH0.SWIECHWBIC = wMT_87A
    newYSWIECH0.SWIECHW52A = wMT_87A
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
    
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "950"
    newYSWIECH0.SWIECHWES = "S"
    newYSWIECH0.SWIECHWBIC = wMT_87A
    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
Else
'________________________________________________________________
    xYSWIECH0.SWIECHSENS = "D"
    If xYSWIECH0.SWIECHWBIC <> wMT_87A Then
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "202"
        newYSWIECH0.SWIECHWES = "S"
        newYSWIECH0.SWIECHW52A = wMT_87A
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    End If
    
    If xYSWIECH0.SWIECHWBIC <> "SOGEFRPP" Then
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "900"
        newYSWIECH0.SWIECHWES = "E"
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    End If
    
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "950"
    newYSWIECH0.SWIECHWES = "E"
    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
End If
'________________________________________________________________

End Sub

Public Sub YSWIECH0_Importation_MT320_PRE_MAD()
Dim blnBIARFRPP As Boolean

'If xYSWIECH0.SWIECHOPEN = 5505 Then
'    Debug.Print "YSWIECH0_Importation_MT300_B1"
'End If

If wMT_57A_B1 = "BIARFRPP" Then blnBIARFRPP = True

If wMT_53A_B1 <> "" Then
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_53A_B1, 1, 8)
Else
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_57A_B1, 1, 8)
End If

xYSWIECH0.SWIECHWDEV = wMT_32_DEV
xYSWIECH0.SWIECHWMTD = wMT_32_MTD

If blnBIARFRPP Then
'________________________________________________________________
    xYSWIECH0.SWIECHSENS = "C"
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "910"
    newYSWIECH0.SWIECHWES = "S"
    newYSWIECH0.SWIECHWBIC = wMT_87A
    newYSWIECH0.SWIECHW52A = wMT_87A
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
    
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "950"
    newYSWIECH0.SWIECHWES = "S"
    newYSWIECH0.SWIECHWBIC = wMT_87A
    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
Else
'________________________________________________________________
    xYSWIECH0.SWIECHSENS = "D"
    If xYSWIECH0.SWIECHWBIC <> wMT_87A Then
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "202"
        newYSWIECH0.SWIECHWES = "S"
        newYSWIECH0.SWIECHW52A = wMT_87A
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    End If
    
    If xYSWIECH0.SWIECHWBIC <> "SOGEFRPP" Then
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "900"
        newYSWIECH0.SWIECHWES = "E"
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    End If
    
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "950"
    newYSWIECH0.SWIECHWES = "E"
    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
End If
'________________________________________________________________

End Sub


Public Sub YSWIECH0_Importation_MT320_EMP_RBT()
Dim blnBIARFRPP As Boolean

If wMT_57A_B1 = "BIARFRPP" Then blnBIARFRPP = True

If wMT_53A_B1 <> "" Then
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_53A_B1, 1, 8)
Else
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_57A_B1, 1, 8)
End If

xYSWIECH0.SWIECHWDEV = wMT_32_DEV
xYSWIECH0.SWIECHWMTD = wMT_32_MTD + wMT_34_MTD
xYSWIECH0.SWIECHDECH = wMT_30X
xYSWIECH0.SWIECHW30V = wMT_30X
wMT_30V_JS1 = dateElp("Ouvré", 1, wMT_30X)

If blnBIARFRPP Then
'________________________________________________________________
    xYSWIECH0.SWIECHSENS = "C"
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "910"
    newYSWIECH0.SWIECHWES = "S"
    newYSWIECH0.SWIECHWBIC = wMT_87A
    newYSWIECH0.SWIECHW52A = wMT_87A
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
    
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "950"
    newYSWIECH0.SWIECHWES = "S"
    newYSWIECH0.SWIECHWBIC = wMT_87A
    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
Else
'________________________________________________________________
    xYSWIECH0.SWIECHSENS = "D"
    If xYSWIECH0.SWIECHWBIC <> wMT_87A Then
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "202"
        newYSWIECH0.SWIECHWES = "S"
        newYSWIECH0.SWIECHW52A = wMT_87A
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    End If
    
    If xYSWIECH0.SWIECHWBIC <> "SOGEFRPP" Then
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "900"
        newYSWIECH0.SWIECHWES = "E"
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    End If
    
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "950"
    newYSWIECH0.SWIECHWES = "E"
    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
End If
'________________________________________________________________

End Sub



Public Function YSWIECH0_Update()
Dim K As Integer, xSql As String, X As String
Dim V
On Error GoTo Error_Handler

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

'________________________________________________________________________________

Select Case mYSWIECH0_Fct
    Case "Update": V = sqlYSWIECH0_Update(newYSWIECH0, oldYSWIECH0)
    Case "New": V = sqlYSWIECH0_Insert(newYSWIECH0)
    Case "Delete": V = sqlYSWIECH0_Delete(oldYSWIECH0)
    'Case "SQL": V = sqlYSWIECH0_Update_Field(newYSWIECH0, mYSWIECH0_SQL_Set)
End Select

If Not IsNull(V) Then GoTo Error_MsgBox

Select Case mYSWIECH1_Fct
    Case "Update": V = sqlYSWIECH1_Update(newYSWIECH1, oldYSWIECH1)
    Case "New": V = sqlYSWIECH1_Insert(newYSWIECH1)
    Case "Delete": V = sqlYSWIECH1_Delete(oldYSWIECH1)
'    Case "SQL": V = sqlYSWIECH1_Update_Field(newYSWIECH1, mYSWIECH1_SQL_Set)
End Select

If Not IsNull(V) Then GoTo Error_MsgBox

Select Case mYSWI950_Fct
    Case "Update": V = sqlYSWI950_Update(newYSWI950, oldYSWI950)
    Case "New": V = sqlYSWI950_Insert(newYSWI950)
    Case "Delete": V = sqlYSWI950_Delete(oldYSWI950)
    Case "SQL": V = sqlYSWI950_Update_Field(newYSWI950, mYSWI950_SQL_Set)
    Case "SQL_Table": V = sqlYSWI950_Update_Table(mYSWI950_SQL_Set)
End Select

If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : YSWIECH0_Update"
Exit_sub:

    YSWIECH0_Update = V
    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
           
    End If
    
    mYSWIECH0_Fct = "": mYSWIECH1_Fct = "": mYSWI950_Fct = ""
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    

End Function




Public Sub YSWIECH1_Exclus(lFct As String, lInfo As String)

mYSWIECH1_Fct = "New"

newYSWIECH1.SWIEC1SWID = oldYSWISAB0.SWISABSWID
newYSWIECH1.SWIEC1SEQ0 = 0
newYSWIECH1.SWIEC1SEQ1 = 0
newYSWIECH1.SWIEC1YUSR = usrName_UCase
newYSWIECH1.SWIEC1YAMJ = DSys
newYSWIECH1.SWIEC1YHMS = time_Hms
newYSWIECH1.SWIEC1YVER = 0
newYSWIECH1.SWIEC1INFO = "<FCT:" & lFct & "><X:" & lInfo & ">"
Call YSWIECH0_Update
mYSWIECH1_Fct = ""


End Sub
Public Sub YSWIECH1_Info(lFct As String, lInfo As String)
Dim xSql As String

xSql = "select SWIEC1SEQ1  from " & paramIBM_Library_SABSPE & ".YSWIECH1 " _
     & " where SWIEC1SWID = " & oldYSWIECH0.SWIECHSWID _
     & " order by SWIEC1SEQ1 desc FETCH FIRST 1 ROWS ONLY"
     
 '& " and  SWIEC1SEQ0 = " & oldYSWIECH0.SWIECHSEQ0
     
Set rsSabX = cnsab.Execute(xSql)

If rsSabX.EOF Then
    newYSWIECH1.SWIEC1SEQ1 = 1
Else
    newYSWIECH1.SWIEC1SEQ1 = rsSabX(0) + 1
End If

mYSWIECH1_Fct = "New"

newYSWIECH1.SWIEC1SWID = oldYSWIECH0.SWIECHSWID
newYSWIECH1.SWIEC1SEQ0 = oldYSWIECH0.SWIECHSEQ0

newYSWIECH1.SWIEC1YUSR = usrName_UCase
newYSWIECH1.SWIEC1YAMJ = DSys
newYSWIECH1.SWIEC1YHMS = time_Hms
newYSWIECH1.SWIEC1YVER = 0
newYSWIECH1.SWIEC1INFO = "<FCT:" & lFct & "><X:" & lInfo & ">"


End Sub


Public Sub YSWIECH0_Update_STA(lSTA As String)
Dim X As String

X = Trim(InputBox("Préciser le motif :"))
If X = "" Then
    Call MsgBox("Abandon de la transaction", vbCritical, "YSWIECH0_Update_STA")
Else

    mYSWIECH0_Fct = "Update"
    newYSWIECH0 = oldYSWIECH0
    newYSWIECH0.SWIECHSWIX = 0
    newYSWIECH0.SWIECHSWIL = 0
    newYSWIECH0.SWIECHSTA = lSTA
    newYSWIECH0.SWIECHSTAK = ""
    newYSWIECH0.SWIECHYAMJ = DSys
    newYSWIECH0.SWIECHYHMS = time_Hms
    newYSWIECH0.SWIECHYUSR = usrName_UCase
    Call YSWIECH1_Info(lSTA, X)
        
    If oldYSWIECH0.SWIECHSWIX > 0 Then
        newYSWI950.SWI950SWID = oldYSWIECH0.SWIECHSWIX
        newYSWI950.SWI950SWIL = oldYSWIECH0.SWIECHSWIL
        mYSWI950_SQL_Set = " set SWI950SWIX = 0"
        mYSWI950_Fct = "SQL"
    End If
    Call YSWIECH0_Update
    Call cmdSelect_SQL_7
End If
mYSWIECH0_Fct = "": mYSWIECH1_Fct = ""

End Sub

Public Sub YSWIECH0_Importation_MT103_CPT()
Dim blnBIARFRPP As Boolean, blnSOGEFRPP As Boolean, xSql As String

xYSWIECH0.SWIECHWDEV = wMT_32_DEV
xYSWIECH0.SWIECHWMTD = wMT_32_MTD
xYSWIECH0.SWIECHDECH = wMT_30V
xYSWIECH0.SWIECHW30V = wMT_30V
wMT_30V_JS1 = dateElp("Ouvré", 1, wMT_30V)

xYSWIECH0.SWIECHWBIC = Mid$(oldYSWISAB0.SWISABWBIC, 1, 8)
'________________________________________________________________
If Mid$(xYSWIECH0.SWIECHWBIC, 1, 4) = "SOGE" Then
    xYSWIECH0.SWIECHSENS = "D"
        
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "198"
    newYSWIECH0.SWIECHWES = "E"
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
    
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "950"
    newYSWIECH0.SWIECHWES = "E"
    newYSWIECH0.SWIECHWBIC = "SOGEFRPP"
    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
Else
'________________________________________________________________
    If YSWI950_BIC_Nostro(xYSWIECH0.SWIECHWBIC) Then
        xYSWIECH0.SWIECHSENS = "D"
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "900"
        newYSWIECH0.SWIECHWES = "E"
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "950"
        newYSWIECH0.SWIECHWES = "E"
        newYSWIECH0.SWIECHDECH = wMT_30V_JS1
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    End If


    If wMT_57A <> "" Then
        xYSWIECH0.SWIECHWBIC = Mid$(wMT_57A, 1, 8)
    '________________________________________________________________
        If YSWI950_BIC_Loro(xYSWIECH0.SWIECHWBIC) Then
            xYSWIECH0.SWIECHSENS = "C"
            
            xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
            newYSWIECH0 = xYSWIECH0
            newYSWIECH0.SWIECHWMTK = "910"
            newYSWIECH0.SWIECHWES = "S"
            mYSWIECH0_Fct = "New"
            Call YSWIECH0_Update
            
            xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
            newYSWIECH0 = xYSWIECH0
            newYSWIECH0.SWIECHWMTK = "950"
            newYSWIECH0.SWIECHWES = "S"
            newYSWIECH0.SWIECHDECH = wMT_30V_JS1
            mYSWIECH0_Fct = "New"
            Call YSWIECH0_Update
        Else
    '________________________________________________________________
            If YSWI950_BIC_Nostro(xYSWIECH0.SWIECHWBIC) Then
                If xYSWIECH0.SWIECHWBIC <> Mid$(oldYSWISAB0.SWISABWBIC, 1, 8) Then
                    xYSWIECH0.SWIECHSENS = "C"
                    
                    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
                    newYSWIECH0 = xYSWIECH0
                    newYSWIECH0.SWIECHWMTK = "950"
                    newYSWIECH0.SWIECHWES = "E"
                    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
                    mYSWIECH0_Fct = "New"
                    Call YSWIECH0_Update
                End If
            End If
        End If
    End If
End If


If wMT_52A <> "" Then
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_52A, 1, 8)
    If YSWI950_BIC_Loro(xYSWIECH0.SWIECHWBIC) Then Call YSWIECH0_Importation_MT103_52A
End If




End Sub

Public Sub YSWIECH0_Importation_MT103_TRF()

xYSWIECH0.SWIECHWDEV = wMT_32_DEV
xYSWIECH0.SWIECHWMTD = wMT_32_MTD
xYSWIECH0.SWIECHDECH = wMT_30V
xYSWIECH0.SWIECHW30V = wMT_30V
wMT_30V_JS1 = dateElp("Ouvré", 1, wMT_30V)

xYSWIECH0.SWIECHWBIC = Mid$(oldYSWISAB0.SWISABWBIC, 1, 8)
'________________________________________________________________
If Mid$(xYSWIECH0.SWIECHWBIC, 1, 4) = "SOGE" Then
    xYSWIECH0.SWIECHSENS = "D"
        
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "198"
    newYSWIECH0.SWIECHWES = "E"
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
    
    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
    newYSWIECH0 = xYSWIECH0
    newYSWIECH0.SWIECHWMTK = "950"
    newYSWIECH0.SWIECHWES = "E"
    newYSWIECH0.SWIECHWBIC = "SOGEFRPP"
    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
    mYSWIECH0_Fct = "New"
    Call YSWIECH0_Update
Else
'________________________________________________________________
    If YSWI950_BIC_Nostro(xYSWIECH0.SWIECHWBIC) Then
        xYSWIECH0.SWIECHSENS = "D"
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "900"
        newYSWIECH0.SWIECHWES = "E"
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
        
        xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
        newYSWIECH0 = xYSWIECH0
        newYSWIECH0.SWIECHWMTK = "950"
        newYSWIECH0.SWIECHWES = "E"
        newYSWIECH0.SWIECHDECH = wMT_30V_JS1
        mYSWIECH0_Fct = "New"
        Call YSWIECH0_Update
    End If


    If wMT_57A <> "" Then
        xYSWIECH0.SWIECHWBIC = Mid$(wMT_57A, 1, 8)
    '________________________________________________________________
        If YSWI950_BIC_Loro(xYSWIECH0.SWIECHWBIC) Then
            xYSWIECH0.SWIECHSENS = "C"
            
            xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
            newYSWIECH0 = xYSWIECH0
            newYSWIECH0.SWIECHWMTK = "910"
            newYSWIECH0.SWIECHWES = "S"
            mYSWIECH0_Fct = "New"
            Call YSWIECH0_Update
            
            xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
            newYSWIECH0 = xYSWIECH0
            newYSWIECH0.SWIECHWMTK = "950"
            newYSWIECH0.SWIECHWES = "S"
            newYSWIECH0.SWIECHDECH = wMT_30V_JS1
            mYSWIECH0_Fct = "New"
            Call YSWIECH0_Update
        Else
    '________________________________________________________________
            If YSWI950_BIC_Nostro(xYSWIECH0.SWIECHWBIC) Then
                If xYSWIECH0.SWIECHWBIC <> Mid$(oldYSWISAB0.SWISABWBIC, 1, 8) Then
                    xYSWIECH0.SWIECHSENS = "C"
                    
                    xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
                    newYSWIECH0 = xYSWIECH0
                    newYSWIECH0.SWIECHWMTK = "950"
                    newYSWIECH0.SWIECHWES = "E"
                    newYSWIECH0.SWIECHDECH = wMT_30V_JS1
                    mYSWIECH0_Fct = "New"
                    Call YSWIECH0_Update
                End If
            End If
        End If
    End If
End If

If wMT_52A <> "" Then
    xYSWIECH0.SWIECHWBIC = Mid$(wMT_52A, 1, 8)
'________________________________________________________________
    If YSWI950_BIC_Loro(xYSWIECH0.SWIECHWBIC) Then Call YSWIECH0_Importation_MT103_52A
    
End If

End Sub

Public Sub YSWI950_BIC_Load()
Dim xSql As String

xSql = "select count(distinct SWISABWBIC)  from " & paramIBM_Library_SABSPE & ".YSWISAB0" _
     & " where SWISABWES = 'S' and SWISABWMTK = '950' and SWISABWAMJ > " & DSys - 10000
Set rsSab = cnsab.Execute(xSql)

arrBIC_Loro_Nb = rsSab(0) + 1

ReDim arrBIC_Loro(arrBIC_Loro_Nb)

xSql = "select distinct SWISABWBIC  from " & paramIBM_Library_SABSPE & ".YSWISAB0" _
     & " where SWISABWES = 'S' and SWISABWMTK = '950' and SWISABWAMJ > " & DSys - 10000 _
     & " group by SWISABWBIC order by SWISABWBIC"
Set rsSab = cnsab.Execute(xSql)

arrBIC_Loro_Nb = 0
Do While Not rsSab.EOF
    arrBIC_Loro_Nb = arrBIC_Loro_Nb + 1
    arrBIC_Loro(arrBIC_Loro_Nb) = Mid$(rsSab("SWISABWBIC"), 1, 8)
    rsSab.MoveNext

Loop
'___________________________________________________________________________________
xSql = "select count(distinct SWISABWBIC)  from " & paramIBM_Library_SABSPE & ".YSWISAB0" _
     & " where SWISABWES = 'E' and SWISABWMTK = '950' and SWISABWAMJ > " & DSys - 10000
Set rsSab = cnsab.Execute(xSql)

arrBIC_Nostro_Nb = rsSab(0) + 1

ReDim arrBIC_Nostro(arrBIC_Nostro_Nb)

xSql = "select distinct SWISABWBIC  from " & paramIBM_Library_SABSPE & ".YSWISAB0" _
     & " where SWISABWES = 'E' and SWISABWMTK = '950' and SWISABWAMJ > " & DSys - 10000 _
     & " group by SWISABWBIC order by SWISABWBIC"
Set rsSab = cnsab.Execute(xSql)

arrBIC_Nostro_Nb = 0
Do While Not rsSab.EOF
    arrBIC_Nostro_Nb = arrBIC_Nostro_Nb + 1
    arrBIC_Nostro(arrBIC_Nostro_Nb) = Mid$(rsSab("SWISABWBIC"), 1, 8)
    rsSab.MoveNext

Loop

End Sub
Public Function YSWI950_BIC_Loro(lBIC As String) As Boolean
Dim K As Integer

YSWI950_BIC_Loro = False
For K = 1 To arrBIC_Loro_Nb
    If lBIC = arrBIC_Loro(K) Then
        YSWI950_BIC_Loro = True
        Exit For
    End If
Next K

End Function


Public Function YSWI950_BIC_Nostro(lBIC As String) As Boolean
Dim K As Integer

YSWI950_BIC_Nostro = False
For K = 1 To arrBIC_Nostro_Nb
    If lBIC = arrBIC_Nostro(K) Then
        YSWI950_BIC_Nostro = True
        Exit For
    End If
Next K

End Function


Public Sub YSWIECH0_Auto()

'Debug.Print "YSWIECH0_Auto"
'blnReprise = True

If arrBIC_Loro_Nb = 0 Then YSWI950_BIC_Load
Call YSWI950_Importation

Call YSWIECH0_Importation

Call YSWIECH0_Match

Call YSWIECH0_Importation_2
Call YSWIECH0_Match

Call YSWIECH0_Match_Ignore

End Sub

Public Sub YSWIECH0_Importation_MT103_52A()
Dim xSql As String, curP As Currency, curD As Currency

 xSql = "select *  from " & paramIBM_Library_SAB & ".ZCHGDET0 " _
      & " where CHGDETETA =1 and CHGDETAGE = 1 and CHGDETSER = '" & oldYSWISAB0.SWISABSER & "'and CHGDETSSE = '" & oldYSWISAB0.SWISABSSE & "'" _
      & " and CHGDETOPE = '" & oldYSWISAB0.SWISABOPEC & "' and CHGDETDOS = " & oldYSWISAB0.SWISABOPEN _
      & " and substring(CHGDETORD ,1 , 1) = 'R' and CHGDETMON > 0"
 Set rsSabX = cnsab.Execute(xSql)
 
 curP = 0: curD = 0
 Do Until rsSabX.EOF
     If rsSabX("CHGDETTYP") = "P" Then
         curP = curP + rsSabX("CHGDETMON")
         xYSWIECH0.SWIECHWDEV = rsSabX("CHGDETDE1")
     Else
         curD = curD + rsSabX("CHGDETMON")
     End If

     rsSabX.MoveNext
 Loop
 
 
 xYSWIECH0.SWIECHSENS = "D"
 If curD <> 0 Then
     xYSWIECH0.SWIECHWMTD = curD
     
     xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
     newYSWIECH0 = xYSWIECH0
     newYSWIECH0.SWIECHWMTK = "950"
     newYSWIECH0.SWIECHWES = "S"
     newYSWIECH0.SWIECHDECH = wMT_30V_JS1
     mYSWIECH0_Fct = "New"
     Call YSWIECH0_Update
 End If
 If curP <> 0 Then
 
     xYSWIECH0.SWIECHWMTD = curP
     
     xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
     newYSWIECH0 = xYSWIECH0
     newYSWIECH0.SWIECHWMTK = "950"
     newYSWIECH0.SWIECHWES = "S"
     newYSWIECH0.SWIECHDECH = wMT_30V_JS1
     mYSWIECH0_Fct = "New"
     Call YSWIECH0_Update

     xYSWIECH0.SWIECHWMTD = curP + curD
              
     xYSWIECH0.SWIECHSEQ0 = xYSWIECH0.SWIECHSEQ0 + 1
     newYSWIECH0 = xYSWIECH0
     newYSWIECH0.SWIECHWMTK = "900"
     newYSWIECH0.SWIECHWES = "S"
     mYSWIECH0_Fct = "New"
     Call YSWIECH0_Update
 End If


End Sub

Public Sub Importation_SAA_198_Alerte()

Dim xSql As String, X As String, X2 As String, wK115 As String

On Error Resume Next

If blnK115 Then
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 " _
         & " where SWISABSWID > " & mK115_SWISABSWID & " and SWISABWMTK in (198 , 298) and SWISABWSTA not in ('V',' ') order by SWISABSWID"
    
    Set rsSab = cnsab.Execute(xSql)
    
    Do While Not rsSab.EOF
        X2 = Importation_SAA_198(rsSab("SWISABWID1"), rsSab("SWISABWIDL"), rsSab("SWISABWIDH"), wK115)
        xSql = "select * from rMesg " _
             & "where Aid = " & rsSab("SWISABWID1") _
             & " and Mesg_s_umidl = " & rsSab("SWISABWIDL") _
             & " and Mesg_s_umidh  =  " & rsSab("SWISABWIDH")
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
        
         If Not rsSIDE_DB.EOF Then
                
             Call srvrMesg_GetBuffer_ODBC(rsSIDE_DB, xrMesg)
             
             X = "Message SWIFT " & xrMesg.mesg_type & " _ champ 115 : " & wK115 & " (" & X2 & ")"
             Call cmdSendMail_SAA_Alerte_rMesg("SAA_" & xrMesg.mesg_type, X, X, xrMesg.x_inst0_unit_name, "")
            
        End If
        
        rsSab.MoveNext
    
    Loop
End If
End Sub
