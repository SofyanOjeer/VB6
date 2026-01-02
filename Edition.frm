VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmEdition 
   AutoRedraw      =   -1  'True
   Caption         =   "Edition"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "Edition.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9150
   ScaleWidth      =   13560
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   585
      Left            =   5745
      TabIndex        =   43
      Top             =   2085
      Visible         =   0   'False
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   1032
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1260
      Top             =   45
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8550
      Left            =   60
      TabIndex        =   3
      Top             =   780
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   15081
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "NoPaper"
      TabPicture(0)   =   "Edition.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraNoPaper"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Gestion des impressions < NoPaper"
      TabPicture(1)   =   "Edition.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraSPLF"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "X"
      TabPicture(2)   =   "Edition.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   7650
         Left            =   -74850
         TabIndex        =   44
         Top             =   660
         Width           =   13140
         _ExtentX        =   23178
         _ExtentY        =   13494
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "Edition.frx":035E
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "dirW"
         Tab(0).Control(1)=   "filW"
         Tab(0).Control(2)=   "lstW"
         Tab(0).Control(3)=   "txtModèle_RTF"
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Mail"
         TabPicture(1)   =   "Edition.frx":037A
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "fraMail_MT"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Tab 2"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         Begin VB.DirListBox dirW 
            Height          =   1440
            Left            =   -64815
            TabIndex        =   58
            Top             =   1545
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.FileListBox filW 
            Height          =   2625
            Left            =   -65310
            TabIndex        =   57
            Top             =   4350
            Visible         =   0   'False
            Width           =   2745
         End
         Begin VB.ListBox lstW 
            BackColor       =   &H00FFFFFF&
            Height          =   1620
            Left            =   -74385
            Sorted          =   -1  'True
            TabIndex        =   56
            Top             =   5655
            Visible         =   0   'False
            Width           =   8460
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
            Left            =   315
            TabIndex        =   45
            Top             =   285
            Visible         =   0   'False
            Width           =   12600
            Begin VB.CheckBox chkMail_Mt_Detail 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0FFF0&
               Caption         =   "inclure le détail du document dans le corps du courriel"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   480
               TabIndex        =   63
               Top             =   3765
               Value           =   1  'Checked
               Width           =   4920
            End
            Begin VB.CheckBox chkMail_Mt_Lnk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00F0FFF0&
               Caption         =   "inclure le lien hypertexte dans le corps du courriel"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   480
               TabIndex        =   60
               Top             =   3195
               Value           =   1  'Checked
               Width           =   4920
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
               Left            =   6390
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   54
               Top             =   2280
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
               Left            =   2955
               Sorted          =   -1  'True
               Style           =   1  'Checkbox
               TabIndex        =   53
               Top             =   1320
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
               Left            =   2970
               MultiLine       =   -1  'True
               TabIndex        =   52
               Top             =   585
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
               Left            =   2985
               MultiLine       =   -1  'True
               TabIndex        =   51
               Top             =   1575
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
               Height          =   1755
               Left            =   2925
               MultiLine       =   -1  'True
               TabIndex        =   50
               Top             =   4560
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
               TabIndex        =   49
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
               TabIndex        =   48
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
               Left            =   465
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   465
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
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   1545
               Width           =   1770
            End
            Begin VB.TextBox txtMail_Mt_Subject 
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
               Left            =   3000
               MultiLine       =   -1  'True
               TabIndex        =   62
               Top             =   2640
               Width           =   9000
            End
            Begin VB.Label libMail_Mt_Subject 
               BackColor       =   &H00F0FFF0&
               Caption         =   "Objet"
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
               Left            =   480
               TabIndex        =   61
               Top             =   2685
               Width           =   2130
            End
            Begin VB.Label txtMail_Mt_Comment 
               BackColor       =   &H00F0FFF0&
               Caption         =   "Commentaire"
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
               Left            =   495
               TabIndex        =   59
               Top             =   4530
               Width           =   2130
            End
         End
         Begin RichTextLib.RichTextBox txtModèle_RTF 
            Height          =   4515
            Left            =   -74310
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   465
            Width           =   8505
            _ExtentX        =   15002
            _ExtentY        =   7964
            _Version        =   393217
            BackColor       =   14737632
            HideSelection   =   0   'False
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"Edition.frx":0396
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame fraNoPaper 
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
         Height          =   7875
         Left            =   135
         TabIndex        =   27
         Top             =   405
         Width           =   13260
         Begin VB.Frame fraNoPaper_Select 
            BackColor       =   &H00F0FFFF&
            Height          =   1425
            Left            =   165
            TabIndex        =   28
            Top             =   150
            Width           =   13035
            Begin VB.ComboBox cboNoPaper_Unit 
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
               Left            =   9870
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   240
               Width           =   2895
            End
            Begin VB.TextBox txtNoPaper_Doc 
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
               Left            =   5400
               TabIndex        =   37
               Top             =   810
               Width           =   3075
            End
            Begin VB.ComboBox cboNoPaper_User 
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
               ItemData        =   "Edition.frx":040D
               Left            =   9855
               List            =   "Edition.frx":040F
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   855
               Width           =   2880
            End
            Begin VB.ListBox lstNoPaper 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   960
               Left            =   1635
               Sorted          =   -1  'True
               TabIndex        =   33
               Top             =   225
               Width           =   2760
            End
            Begin VB.OptionButton optNoPaper_Test 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Test"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   150
               TabIndex        =   32
               Top             =   975
               Width           =   1410
            End
            Begin VB.OptionButton optNoPaper_Prod 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Production"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   150
               TabIndex        =   31
               Top             =   660
               Value           =   -1  'True
               Width           =   1410
            End
            Begin VB.OptionButton optNoPaper_Archive 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Archive"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   150
               TabIndex        =   30
               Top             =   315
               Width           =   1410
            End
            Begin MSComCtl2.DTPicker txtNoPaper_AMJ_Min 
               Height          =   300
               Left            =   5415
               TabIndex        =   40
               Top             =   300
               Width           =   1605
               _ExtentX        =   2831
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
               CheckBox        =   -1  'True
               CustomFormat    =   "dd  MM yyy"
               Format          =   99614723
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtNoPaper_AMJ_Max 
               Height          =   300
               Left            =   7185
               TabIndex        =   41
               Top             =   315
               Width           =   1350
               _ExtentX        =   2381
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
               Format          =   99614723
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblNoPaper_AMJ 
               BackColor       =   &H00F0FFFF&
               Caption         =   "période"
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
               Left            =   4620
               TabIndex        =   42
               Top             =   300
               Width           =   1110
            End
            Begin VB.Label lblNoPaper_Unit 
               BackColor       =   &H00F0FFFF&
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
               Height          =   270
               Left            =   8745
               TabIndex        =   39
               Top             =   300
               Width           =   690
            End
            Begin VB.Label lblNoPaper_Doc 
               BackColor       =   &H00F0FFFF&
               Caption         =   "document"
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
               Left            =   4560
               TabIndex        =   36
               Top             =   825
               Width           =   1110
            End
            Begin VB.Label lblNoPaper_User 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Utilisateur"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   8805
               TabIndex        =   34
               Top             =   930
               Width           =   1020
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgNoPaper 
            Height          =   6180
            Left            =   60
            TabIndex        =   29
            Top             =   1590
            Width           =   13125
            _ExtentX        =   23151
            _ExtentY        =   10901
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            FixedCols       =   0
            RowHeightMin    =   200
            BackColor       =   15794175
            ForeColor       =   12582912
            BackColorFixed  =   12648447
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   15794175
            AllowBigSelection=   0   'False
            TextStyle       =   4
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"Edition.frx":0411
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
      Begin VB.Frame fraSPLF 
         Height          =   7905
         Left            =   -74900
         TabIndex        =   5
         Top             =   350
         Width           =   13350
         Begin VB.FileListBox filDoc 
            Height          =   1455
            Left            =   135
            TabIndex        =   18
            Top             =   1470
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.ComboBox cboSPLF_Filigrane 
            Height          =   315
            Left            =   11520
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   165
            Width           =   1770
         End
         Begin VB.ComboBox cboSPLF_Folder 
            Height          =   315
            Left            =   840
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   240
            Width           =   3180
         End
         Begin VB.TextBox txtSPLF_PageStart 
            Height          =   285
            Left            =   11520
            TabIndex        =   15
            Text            =   "1"
            Top             =   600
            Width           =   720
         End
         Begin VB.TextBox txtSPLF_PageEnd 
            Height          =   285
            Left            =   12480
            TabIndex        =   14
            Text            =   "10"
            Top             =   615
            Width           =   690
         End
         Begin VB.ComboBox cboSPLF_User 
            Height          =   315
            ItemData        =   "Edition.frx":0498
            Left            =   4680
            List            =   "Edition.frx":049A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   720
            Width           =   2730
         End
         Begin VB.Frame fraContextOptions 
            Caption         =   "Options"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3360
            Left            =   9240
            TabIndex        =   9
            Top             =   1440
            Width           =   3855
            Begin VB.OptionButton optSplf_RefreshNo 
               Caption         =   "ne jamais rafraîchir"
               Height          =   240
               Left            =   240
               TabIndex        =   12
               Top             =   1350
               Width           =   3195
            End
            Begin VB.OptionButton optSplf_RefreshTimer 
               Caption         =   "Rafraîchir automatiquement (10sec)"
               Height          =   240
               Left            =   210
               TabIndex        =   11
               Top             =   825
               Width           =   3270
            End
            Begin VB.OptionButton optSplf_RefreshActivate 
               Caption         =   "rafraîchir quand la fenêtre devient active"
               Height          =   240
               Left            =   240
               TabIndex        =   10
               Top             =   405
               Value           =   -1  'True
               Width           =   3300
            End
         End
         Begin VB.ComboBox cboSPLF_Unit 
            Height          =   315
            Left            =   840
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   630
            Width           =   2100
         End
         Begin VB.ComboBox cboSPLF_Amj 
            Height          =   315
            Left            =   4680
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   1755
         End
         Begin VB.TextBox txtSPLF_Name 
            Height          =   285
            Left            =   8520
            TabIndex        =   6
            Top             =   720
            Width           =   1695
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6735
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   13125
            _ExtentX        =   23151
            _ExtentY        =   11880
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            FixedCols       =   0
            RowHeightMin    =   200
            BackColor       =   14737632
            ForeColor       =   12582912
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyle       =   4
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"Edition.frx":049C
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
         Begin VB.Label lblSPLF_Filigrane 
            Caption         =   "Filigrane"
            Height          =   255
            Left            =   10680
            TabIndex        =   26
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblSPLF_Folder 
            Caption         =   "Origine"
            Height          =   270
            Left            =   135
            TabIndex        =   25
            Top             =   360
            Width           =   555
         End
         Begin VB.Label lblSPLF_Page 
            Caption         =   "pages de..à"
            Height          =   240
            Left            =   10320
            TabIndex        =   24
            Top             =   720
            Width           =   885
         End
         Begin VB.Label lblSPLF_User 
            Caption         =   "Utilisateur"
            Height          =   225
            Left            =   3840
            TabIndex        =   23
            Top             =   720
            Width           =   750
         End
         Begin VB.Label lblSPLF_Unit 
            Caption         =   "Unité"
            Height          =   270
            Left            =   150
            TabIndex        =   22
            Top             =   720
            Width           =   555
         End
         Begin VB.Label lblSPLF_Amj 
            Caption         =   "Date"
            Height          =   270
            Left            =   4080
            TabIndex        =   21
            Top             =   360
            Width           =   555
         End
         Begin VB.Label lblSPLF_Name 
            Caption         =   "doc = *xxx*"
            Height          =   255
            Left            =   7560
            TabIndex        =   20
            Top             =   720
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8520
      TabIndex        =   1
      Top             =   30
      Width           =   4320
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   12960
      Picture         =   "Edition.frx":0565
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   15
      Width           =   500
   End
   Begin VB.Label libSelect 
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
      Left            =   1245
      TabIndex        =   4
      Top             =   0
      Width           =   5490
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextReset 
         Caption         =   "Restaurer les sélections"
      End
      Begin VB.Menu mnuContextRefresh 
         Caption         =   "Rafraîchir l'affichage"
      End
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContextX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContext_NoPaper 
         Caption         =   "Lancer Auto_NoPaper"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuContext_NoPaper_Recap_Dispatch 
         Caption         =   "Lancer Auto_NoPaper_Recap => Dispatch services"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuContextX1b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuSPLF 
      Caption         =   "mnuSplf"
      Visible         =   0   'False
      Begin VB.Menu mnuSplf_Afficher 
         Caption         =   "Afficher"
      End
      Begin VB.Menu mnuSplf_Afficher_Partiel 
         Caption         =   "Afficher (pages de ...à ...)"
      End
      Begin VB.Menu mnuSplf_X1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSplf_Imprimer 
         Caption         =   "Imprimer"
      End
      Begin VB.Menu mnuSplf_Imprimer_Partiel 
         Caption         =   "Imprimer (pages de ...à ...)"
      End
      Begin VB.Menu mnuSplf_X2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSplf_Archiver 
         Caption         =   "Archiver"
      End
      Begin VB.Menu mnuSplf_Supprimer 
         Caption         =   "Supprimer"
      End
      Begin VB.Menu mnuSplf_X3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSPLF_Save 
         Caption         =   "Enregistrer sous C:\Temp\       .rtf"
      End
      Begin VB.Menu mnuSPLF_Email 
         Caption         =   "Envoyer le document > Email"
      End
      Begin VB.Menu mnuSplf_X4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSPLF_CACLS 
         Caption         =   "Affecter les droits - CACLS"
      End
   End
   Begin VB.Menu mnuNoPaper 
      Caption         =   "mnuNoPaper"
      Visible         =   0   'False
      Begin VB.Menu mnuNoPaper_Display 
         Caption         =   "Afficher ce document"
      End
      Begin VB.Menu mnuNoPaper_Mail 
         Caption         =   "Envoyer ce document par mail"
      End
      Begin VB.Menu mnuNoPaper_Archiver 
         Caption         =   "Archiver ce document"
      End
      Begin VB.Menu mnuNoPaper_Supprimer 
         Caption         =   "Supprimer ce document"
      End
      Begin VB.Menu mnuNoPaper_x 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNoPaper_Recap 
         Caption         =   "Lancer NoPaper_Recapitulatif => Mail (moi uniquement)"
      End
      Begin VB.Menu mnuNoPaper_xx 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecapTOUS 
         Caption         =   "Lancer NoPaper_Recapitulatif => Mail (pour tous les services)"
      End
   End
End
Attribute VB_Name = "frmEdition"
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
Dim EditionAut As typeAuthorization
Dim blnAuto As Boolean, blnError As Boolean, blnForm_Init As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim xRtfEdition As typeEdition

Dim xLine As String * 256, xLine_SelBold As String * 256, xLine_SelUnderline As String * 256
Dim blnSelBold As Boolean, blnSelUnderline As Boolean
Dim arrBold_SelStart() As Long, arrBold_SelLength() As Long, arrBold_Nb As Long
Dim arrUnderline_SelStart() As Long, arrUnderline_SelLength() As Long, arrUnderline_Nb As Long
Dim lenRTF As Long, lenLine As Integer
Dim blnTimer_Enabled As Boolean
Dim fsoFolder As Folder

Dim mfilDoc_List As String
Dim meSplfJob As typeSplfJob

Dim fgSelect_FileName As String, wFileName_Save As String

Dim xElpTable As typeElpTable, xEdition_Form As typeEdition_Form
Dim xUser As typeUser, xUnit As typeUnit
Dim selectUnit As String, selectUser As String, selectAmj As String, selectName As String

Dim mPageStart As Long, mPageEnd As Long

Dim k_Filigrane_Test As Integer

'_____________________________________________________________________________
Dim msFolder As Scripting.Folder, msSubFolder As Scripting.Folder, msFile As Scripting.File
Dim mNoPaper_Folder As String
Dim fgNoPaper_FormatString As String, fgNoPaper_K As Integer
Dim fgNoPaper_RowDisplay As Integer, fgNoPaper_RowClick As Integer, fgNoPaper_ColClick As Integer
Dim fgNoPaper_ColorClick As Long, fgNoPaper_ColorDisplay As Long
Dim fgNoPaper_Sort1 As Integer, fgNoPaper_Sort2 As Integer
Dim fgNoPaper_SortAD As Integer, fgNoPaper_Sort1_Old As Integer
Dim fgNoPaper_arrIndex As Integer
Dim blnfgNoPaper_DisplayLine As Boolean

Dim mNoPaper_Opt As String, mNoPaper_Unit As String, mNoPaper_User As String, mNoPaper_Doc As String

Dim mMail_Txt As String, mMail_Lnk As String

Public Function AImprimer(fichier As String) As Boolean
Dim fic As Long
Dim ligIn As String

    AImprimer = True
    If xUser.QSYSOPR = "1" Or Trim(xUser.Id) = "BIA_INFO" Then
        fic = FreeFile
        Open fichier For Input As #fic
        Do Until EOF(fic)
            Line Input #fic, ligIn
            If InStr(ligIn, "Rien à imprimer") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "RIEN A IMPRIMER") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "Aucun avis à éditer") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "AUCUN AVIS A EDITER") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "Rien à traiter") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "RIEN A TRAITER") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "PAS D" & Chr(34) & "ENREGISTREMENTS SELECTIONNES") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "PAS D'ENREGISTREMENTS SELECTIONNES") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "PAS D'ANOMALIES DETECTÉES") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "PAS D'ANOMALIES DETECTEES") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "PAS DE CHEQUES A REGULARISER") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "AUCUN ENREGISTREMENT A TRAITER") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "RIEN A LISTER") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "AUCUN ENREGISTREMENT DETAIL N'A ETE DECLARE") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "Rien à éditer") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "RIEN A EDITER") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "Il n'y a pas d'instructions à traiter pour l'opération") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "Aucun effet n'a été envoyé en encaissement") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "PAS D'OPERATIONS SIT SELECTIONNEES") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "PAS DE MENTION") > 0 Then
                AImprimer = False
                Exit Do
            End If
            If InStr(ligIn, "Pas d'anomalies détectés") > 0 Then
                AImprimer = False
                Exit Do
            End If
        Loop
        Close #fic
    End If
    
End Function

Public Sub ecrit_erreur(z As String)
Dim FicSortie As Long
Dim ficName As String
Dim ligOut As String

On Error Resume Next
    ficName = "d:\logs\splf.log"
    ligOut = z & " " & Mid(CStr(100 + Day(Now)), 2) & "/" & Mid(CStr(100 + Month(Now)), 2) & "/" & CStr(Year(Now) & " " & Mid(CStr(100 + Hour(Now)), 2) & ":" & Mid(CStr(100 + Minute(Now)), 2) & ":" & Mid(CStr(100 + Second(Now)), 2))
    FicSortie = FreeFile
    Open ficName For Append As #FicSortie
    Print #FicSortie, ligOut
    Close #FicSortie

End Sub

Public Function special_DAFI(fl As String) As String
Dim dfic As Long
Dim ligIn As String

    special_DAFI = "_DAFI"
    dfic = FreeFile
    Open fl For Input As #dfic
    Do Until EOF(dfic)
        Line Input #dfic, ligIn
        If InStr(UCase(ligIn), "CAUTE007P1") > 0 And InStr(UCase(ligIn), "00CD") > 0 Then
            special_DAFI = "_SOBI"
            Exit Do
        End If
        If InStr(UCase(ligIn), "BIA/SCA601") > 0 And InStr(UCase(ligIn), "/00CD") > 0 Then
            special_DAFI = "_SOBI"
            Exit Do
        End If
    Loop
    Close #dfic
    
End Function

Public Sub fgSelect_DisplayLine()
 Dim X As String, K As Integer, K1 As Integer, K2 As Integer
Dim lenX As Integer, blnOk As Boolean
Dim wEdition_Form_Id As String * 12
Dim wAmj As String, wJob As String
On Error Resume Next

If selectUnit = "" Or Mid$(selectUser, 1, 1) = "_" Then
    blnOk = True
Else
    blnOk = False
End If

X = filDoc.FileName

If selectName <> "" Then
    K = InStr(1, X, selectName)
    If K = 0 Then blnOk = False
End If

K = InStr(1, X, ".")

If K > 0 Then
    xUser.Id = Mid$(X, 1, K - 1)
    Call Table_User(xUser)
    If Trim(xUser.Unit) = selectUnit Then blnOk = True
Else
    xUser.Id = ""
    xUser.Unit = ""
End If

If selectUser <> "" And selectUser <> Trim(xUser.Id) Then blnOk = False



K1 = K + 1
K = InStr(K1, X, "_")
If K > 0 Then
    K2 = InStr(K + 1, X, "_")
    If K2 > 0 Then
        wAmj = Mid$(X, K1, K2 - K1)
        K1 = K2 + 1
        K2 = InStr(K1, X, "_")
        If K2 > 0 Then
            wEdition_Form_Id = Mid$(X, K1, K2 - K1)
            K1 = K2 + 1
            lenX = Len(X)
            If K2 < lenX Then wJob = Mid$(X, K1, lenX - K1)
        End If
    End If
End If
If selectAmj <> "" And selectAmj <> Mid$(wAmj, 1, 8) Then blnOk = False


If blnOk Then
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = X
    fgSelect.Col = 0: fgSelect.Text = xUser.Unit
    fgSelect.Col = 1: fgSelect.Text = xUser.Id              ''''''mId$(X, 1, K - 1)
    fgSelect.Col = 3: fgSelect.Text = wAmj                 '''''''mId$(X, K1, K2 - K1)
    fgSelect.Col = 4: fgSelect.Text = wJob                  '''''''mId$(X, K1, lenX - K1)
 
    
    fgSelect.Col = 2
    xEdition_Form.K1 = "SAB"
    xEdition_Form.K2 = wEdition_Form_Id
    fgSelect.Text = wEdition_Form_Id & rsEdition_Form(xEdition_Form)
End If
End Sub



Private Sub fgSelect_Display()
Dim K As Long
SSTab1.Tab = 1

fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Visible = False

selectUnit = Trim(cboSPLF_Unit)
selectUser = Trim(cboSPLF_User)
selectAmj = Trim(cboSPLF_Amj)
selectName = Trim(txtSPLF_Name)

For K = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = K
    fgSelect_DisplayLine

Next K

    
    



fgSelect_Sort
fgSelect.Visible = True
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

Public Sub fgNoPaper_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgNoPaper.Row

If lRow > 0 And lRow < fgNoPaper.Rows Then
    fgNoPaper.Row = lRow
    For I = 0 To fgNoPaper_arrIndex
        fgNoPaper.Col = I: fgNoPaper.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgNoPaper.Row = mRow
    If fgNoPaper.Row > 0 Then
        lRow = fgNoPaper.Row
        lColor_Old = fgNoPaper.CellBackColor
        For I = 0 To fgNoPaper_arrIndex
          fgNoPaper.Col = I: fgNoPaper.CellBackColor = lColor
        Next I
        fgNoPaper.Col = 0
    End If
End If

End Sub



Public Sub fgSelect_ForeColor(lColor As Long)
For I = 0 To fgSelect_arrIndex
  fgSelect.Col = I: fgSelect.CellForeColor = lColor
Next I

End Sub



Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 3: fgSelect_Sort2 = 3
fgSelect_Sort1_Old = 3
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 5
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 5
fgSelect.LeftCol = 0

End Sub


Public Sub fgNoPaper_Reset()
fgNoPaper.Clear
fgNoPaper_Sort1 = 3: fgNoPaper_Sort2 = 3
fgNoPaper_Sort1_Old = 3
fgNoPaper_RowDisplay = 0: fgNoPaper_RowClick = 0
fgNoPaper_arrIndex = 5
blnfgNoPaper_DisplayLine = False
fgNoPaper_SortAD = 5
fgNoPaper.LeftCol = 0
fgNoPaper.BackColor = RGB(255, 255, 255)
End Sub
Public Sub fgSelect_Sort()
fgSelect.Visible = False
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
fgSelect.Visible = True
End Sub

Public Sub fgNoPaper_Sort()
fgNoPaper.Visible = False
If fgNoPaper.Rows > 1 Then
    fgNoPaper.Row = 1
    fgNoPaper.RowSel = fgNoPaper.Rows - 1
    
    If fgNoPaper_Sort1_Old = fgNoPaper_Sort1 Then
        If fgNoPaper_SortAD = 5 Then
            fgNoPaper_SortAD = 6
        Else
            fgNoPaper_SortAD = 5
        End If
    Else
        fgNoPaper_SortAD = 5
    End If
    fgNoPaper_Sort1_Old = fgNoPaper_Sort1
    
    fgNoPaper.Col = fgNoPaper_Sort1
    fgNoPaper.ColSel = fgNoPaper_Sort2
    fgNoPaper.Sort = fgNoPaper_SortAD
End If
fgNoPaper.Visible = True
End Sub


Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
Dim mK As Integer
mK = lK
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    Select Case lK
        Case 5:
            fgSelect.Col = 0: X = fgSelect.Text
            fgSelect.Col = 2: X = X & fgSelect.Text
         Case 6
    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
fgSelect_Sort1 = mK: fgSelect_Sort2 = mK
End Sub


Public Sub fgNoPaper_SortX(lK As Integer)
Dim I As Integer, X As String
Dim mK As Integer
mK = lK
For I = 1 To fgNoPaper.Rows - 1
    fgNoPaper.Row = I
    fgNoPaper.Col = fgNoPaper_arrIndex
    Select Case lK
        Case 5:
            fgNoPaper.Col = 0: X = fgNoPaper.Text
            fgNoPaper.Col = 2: X = X & fgNoPaper.Text
         Case 6
    End Select
    fgNoPaper.Col = fgNoPaper_arrIndex - 1
    fgNoPaper.Text = X
Next I


fgNoPaper_Sort1 = fgNoPaper_arrIndex - 1: fgNoPaper_Sort2 = fgNoPaper_arrIndex - 1
fgNoPaper_Sort
fgNoPaper_Sort1 = mK: fgNoPaper_Sort2 = mK
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

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), EditionAut)
fraSPLF.Enabled = False

wFct = UCase$(Trim(Mid$(Msg, 1, 12)))

Form_Init

Select Case wFct
                    
    Case "@PRINT_TEST":   blnAuto = True: Auto_SendMail   '''Auto_Edition Msg
    Case "@PRINT_PROD":   blnAuto = True: Auto_Edition Msg
End Select

End Sub


Public Sub Form_Init()
Dim xSQL As String, K As Integer

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

blnControl = False: blnForm_Init = True
Me.Enabled = False

If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistant", vbCritical, "frmEdition.param_init"
    Unload Me
End If

fgSelect.Enabled = True

Call DTPicker_Set(txtNoPaper_AMJ_Min, DSys) ' YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtNoPaper_AMJ_Max, DSys) 'YBIATAB0_DATE_CPT_J)

cboNoPaper_Unit.Clear
cboNoPaper_Unit.AddItem ""
xSQL = "select SSIUSRUIDX , SSIUSRUNIT from " & paramIBM_Library_SABSPE & ".YSSIUSR0 " _
     & " where SSIUSRNAT = 'S' and SSIUSRSTAK = ' ' order by SSIUSRUNIT "

Set rsSab = cnsab.Execute(xSQL)
Do Until rsSab.EOF
    cboNoPaper_Unit.AddItem rsSab("SSIUSRUNIT") & " : " & rsSab("SSIUSRUIDX")
    If currentSSIWINUNIT = rsSab("SSIUSRUNIT") Then K = cboNoPaper_Unit.ListCount - 1
    rsSab.MoveNext
Loop

cboNoPaper_Unit.ListIndex = K

cmdReset
blnControl = True: blnForm_Init = False
optNoPaper_Prod_Click
cboSPLF_Folder_Click

SSTab1.Tab = 0
ProgressBar1.Top = 4000
ProgressBar1.Left = 4000

Set fraMail_MT.Container = fraNoPaper
fraMail_MT.Visible = False
fraMail_MT.Top = fraNoPaper.Top
fraMail_MT.Left = fraNoPaper.Left
lstMail_MT_CC.BackColor = cmdMail_MT_CC.BackColor
lstMail_MT_To.BackColor = cmdMail_MT_To.BackColor
lstParam_YSSIMEL0_USR_Load

mnuNoPaper_Supprimer.Visible = EditionAut.Xspécial

Me.Enabled = True

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

Public Sub lstMail_MT_To_TXT()
Dim K As Integer, X As String
X = ""
For K = 0 To lstMail_MT_To.ListCount - 1

    If lstMail_MT_To.Selected(K) Then X = X & lstMail_MT_To.List(K) & ";"
Next K
If X <> "" Then Mid$(X, Len(X), 1) = " "

txtMail_MT_To = X

End Sub
Private Sub cboNoPaper_Unit_Change()
fgNoPaper.Visible = False: DoEvents
mNoPaper_Unit = Mid$(cboNoPaper_Unit.Text, 1, 3)
cboNoPaper_User_Load (mNoPaper_Unit)
End Sub

Private Sub cboNoPaper_Unit_Click()
fgNoPaper.Visible = False: DoEvents
mNoPaper_Unit = Mid$(cboNoPaper_Unit.Text, 1, 3)
cboNoPaper_User_Load (mNoPaper_Unit)
'lstNoPaper_Click
End Sub

Private Sub cboNoPaper_User_Change()
fgNoPaper.Visible = False: DoEvents
mNoPaper_User = Trim(cboNoPaper_User.Text)
lstNoPaper_Click
End Sub

Private Sub cboNoPaper_User_Click()
fgNoPaper.Visible = False: DoEvents
mNoPaper_User = Trim(cboNoPaper_User.Text)
lstNoPaper_Click

End Sub


Private Sub cboSPLF_Amj_Click()
If blnControl Then fgSelect_Display

End Sub

Private Sub cboSPLF_User_Click()
If blnControl Then lstCourrier_Load

End Sub

Private Sub cmdMail_MT_CC_Click()
lstMail_MT_To.Visible = False
lstMail_MT_CC.Visible = True

End Sub

Private Sub cmdMail_MT_Ok_Click()
Dim X As String, Xto As String, xCC As String
Dim V, blnOk As Boolean, xMessage As String
Dim wSendMail As typeSendMail


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

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'envoi par mail

If chkMail_Mt_Lnk = "0" Then mMail_Lnk = ""

If chkMail_Mt_Detail = "0" Then
    xMessage = ""
Else
    'X = "<p class=MsoPlainText><span style='font-size:8.0pt;font-family:font-family:Courier New'>" _
    '    & "<TD bgcolor = #FFFFFF width=20% height=7><span style='font-size:8.0pt;font-family:Courier New'>" _
    '    & "<TD bgcolor = #FFFFFF width=80% height=7><span style='font-size:8.0pt;font-family:Courier New'>" _
    '    & Replace(mMail_Txt, " ", Chr(160)) & "<BR></TD></TR>" _
    '   & "<o:p></o:p></span></p>"
      X = "<span style='font-size:10;font-family:Courier New'>" & Replace(mMail_Txt, " ", Chr(160))
    xMessage = Replace(X, vbCrLf, "<o:p></o:p></span></p>" & vbCrLf & "<p class=MsoPlainText><span style='font-size:10;font-family:Courier New'>")
End If

wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = "NoPaper"
wSendMail.Recipient = Xto
wSendMail.CcRecipient = xCC

wSendMail.Subject = Trim(txtMail_Mt_Subject)
wSendMail.Attachment = ""


wSendMail.Message = mHtml_Head & "<span style='font-size:10.0pt;font-family:Calibri'>" _
                    & Replace(Trim(txtMail_MT_Message), vbCrLf, "<BR> ") & "<BR><BR>" _
                    & mMail_Lnk & "<BR><BR>" _
                  & htmlFontColor_Black & xMessage & "</div></body></html>"
wSendMail.AsHTML = True
srvSendMail.Monitor wSendMail '$JPL 2014-10-10
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'_________________________________________________________________________________________
lstMail_MT_To.Visible = False
lstMail_MT_CC.Visible = False

fraMail_MT.Visible = False

Exit_sub:

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdMail_MT_Quit_Click()
fraMail_MT.Visible = False

End Sub


Private Sub cmdMail_MT_To_Click()
lstMail_MT_CC.Visible = False
lstMail_MT_To.Visible = True
End Sub

Private Sub fgNoPaper_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String
Dim V
Dim xFolder_In  As String, xName  As String, xExtension As String
On Error Resume Next
If y <= fgNoPaper.RowHeightMin Then
    Select Case fgNoPaper.Col
        Case 0: fgNoPaper_Sort1 = 0: fgNoPaper_Sort2 = 3: fgNoPaper_Sort
        Case 1:  fgNoPaper_Sort1 = 1: fgNoPaper_Sort2 = 3: fgNoPaper_Sort
        Case 2: fgNoPaper_Sort1 = 2: fgNoPaper_Sort2 = 3: fgNoPaper_Sort
        Case 3: fgNoPaper_Sort1 = 3: fgNoPaper_Sort2 = 3: fgNoPaper_Sort
        Case fgNoPaper_arrIndex:  fgNoPaper_SortX fgNoPaper_arrIndex
    End Select
Else
    If fgNoPaper.Rows > 1 Then
        Call fgNoPaper_Color(fgNoPaper_RowClick, MouseMoveUsr.BackColor, fgNoPaper_ColorClick)
        fgNoPaper.Col = fgNoPaper_arrIndex
        'If Button = vbRightButton Then
            'If InStr(mNoPaper_Folder, "\Prod_") > 0 Then
            '    mnuNoPaper_Archiver.Visible = True
            'Else
            '    mnuNoPaper_Archiver.Visible = False
            'End If
            
            Me.PopupMenu mnuNoPaper, vbPopupMenuLeftButton
        'Else
        '    Call frmElpPrt.Windows_Display_File(fgNoPaper.Text)
        'End If
    End If
End If

End Sub


Private Sub fgSelect_Click()
fgSelect.LeftCol = 0

End Sub


Private Sub fgSelect_LeaveCell()
On Error Resume Next
'fgSelect.CellBackColor = &HE0E0E0
End Sub


Private Sub fgSelect_LostFocus()
'Timer1_Monitor

End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String
Dim V

On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex
        fgSelect_FileName = fgSelect.Text
        

        If Me.Enabled Then
            Call lstErr_Clear(lstErr, cmdContext, "> Choisir une option ... "): DoEvents

           mfilDoc_List = filDoc
            Timer1.Enabled = False
            If Button = vbLeftButton Then Me.PopupMenu mnuSPLF, vbPopupMenuLeftButton
        End If
    End If
End If

End Sub


Public Function param_Init()
Dim K As Integer, K1 As Integer, X As String
On Error Resume Next
Dim V
param_Init = Null
'$20060626 jpl$ If Not IsNull(paramEdition_Init(frmEdition.lstErr, Me.cmdContext)) Then Exit Function

dirW.PATH = paramEditionSplf_Folder
cboSPLF_Folder.Clear
    
If EditionAut.Xspécial Then
    K1 = Len(Trim(dirW.PATH))
    For K = 0 To dirW.ListCount - 1
        X = dirW.List(K)
        cboSPLF_Folder.AddItem Mid$(X, K1 + 2, Len(X) - K1)
    Next K
    mnuContext_NoPaper.Visible = True
    mnuNoPaper_Recap.Visible = True
Else
    cboSPLF_Folder.AddItem constArchive
    cboSPLF_Folder.AddItem constCorbeille
    cboSPLF_Folder.AddItem constProduction
    cboSPLF_Folder.AddItem constTest
End If

If UCase(currentUser.Unit) = "INFO" Then
    mnuRecapTOUS.Visible = True
Else
    mnuRecapTOUS.Visible = False
End If
If cboSPLF_Folder.ListCount > 0 Then Call cbo_Scan(constProduction, cboSPLF_Folder)
k_Filigrane_Test = 0
filW.PATH = paramEditionFiligrane_Folder
filW.Pattern = "*.jpg"
cboSPLF_Filigrane.Clear
cboSPLF_Filigrane.AddItem "(aucun)"
cboSPLF_Filigrane.AddItem " Automatique"
For K = 0 To filW.ListCount - 1
    filW.ListIndex = K
    K1 = InStr(1, filW.FileName, ".")
    If K1 > 0 Then
        X = Mid$(filW.FileName, 1, K1 - 1)
        cboSPLF_Filigrane.AddItem X
        If UCase$(Trim(X)) = "TEST" Then k_Filigrane_Test = cboSPLF_Filigrane.ListCount
    End If
Next K
cboSPLF_Filigrane.ListIndex = 0
cboSPLF_Unit.Clear
If currentUser.Edition_Aut = "1" Or EditionAut.Xspécial Then
    Call cbo_LoadId_K2("Unit", "", cboSPLF_Unit)
Else
    cboSPLF_Unit.AddItem currentUser.Unit
End If
cboSPLF_Unit.AddItem " "
cboSPLF_User.Clear
If currentUser.Edition_Aut = "1" Or EditionAut.Xspécial Then
    Call cbo_LoadId("User", cboSPLF_User)
    cboSPLF_User.AddItem " "
    cboSPLF_User.AddItem "V_SALLE"   '$JPL 2014-10-09
Else
    Call Table_User_Load(currentUser.Unit, cboSPLF_User, True)
End If
If cboSPLF_User.ListCount = 0 Then cboSPLF_User.AddItem currentUser.Id
cboSPLF_User.AddItem "_" & currentUser.Unit
cboSPLF_User.AddItem "_T_" & currentUser.Unit

If cboSPLF_User.ListCount > 0 Then Call cbo_Scan(Trim(currentUser.Id), cboSPLF_User)

cboSPLF_Amj.Clear
cboSPLF_Amj.AddItem " "
cboSPLF_Amj.AddItem DSys
For I = 1 To 7
    cboSPLF_Amj.AddItem dateElp("Ouvré", -I, DSys)
Next I
cboSPLF_Amj.ListIndex = cboSPLF_Amj.ListCount - 1

End Function




'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
currentAction = ""
SSTab1.Tab = 0

If Not IsNull(txtNoPaper_AMJ_Min.Value) Then
    txtNoPaper_AMJ_Max.Visible = True
Else
    txtNoPaper_AMJ_Max.Visible = False
End If

mfilDoc_List = ""
filDoc.Visible = False
filW.Visible = False
dirW.Visible = False
blnTimer_Enabled = False
blnControl = True
cboSPLF_Filigrane.Enabled = EditionAut.Xspécial
mnuSPLF.Enabled = EditionAut.Xspécial
mnuSPLF_CACLS.Visible = EditionAut.Xspécial
fraContextOptions.Visible = False
fraSPLF.Enabled = True

fraMail_MT.Visible = False

Timer1_Monitor

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

Private Sub cboSPLF_Filigrane_Click()
Dim X As String
X = cboSPLF_Filigrane
If X = " Automatique" Then X = cboSPLF_Folder

prtFiligrane_Name = paramEditionFiligrane_Folder & X & ".jpg" ' ".bmp"
If X = "(aucun)" Then
    blnFiligrane = False
Else
    blnFiligrane = True
End If

End Sub


Private Sub cboSPLF_Folder_Click()
If blnControl Then
    lstCourrier_Load
    If cboSPLF_Folder = "Production" Then
        cboSPLF_Filigrane = "(aucun)"
    Else
        cboSPLF_Filigrane = "(aucun)"
        '20050627JL  If cboSPLF_Filigrane.ListCount > 0 Then cboSPLF_Filigrane.ListIndex = 0: cboSPLF_Filigrane_Click
    End If
End If

End Sub

Private Sub cboSPLF_Unit_Click()
If blnControl Then fgSelect_Display

End Sub


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
mnuContextRefresh.Enabled = Not optSplf_RefreshTimer
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub


Private Sub filDoc_LostFocus()
'Timer1_Monitor
End Sub

Private Sub Form_GotFocus()
If optSplf_RefreshActivate Then lstCourrier_Load

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If optSplf_RefreshActivate Then lstCourrier_Load

End Sub

Private Sub libSelect_Click()
Select Case SSTab1.Tab
    Case 0: lstNoPaper_Click
    Case 1: lstCourrier_Load
End Select
End Sub

Private Sub lstMail_MT_CC_Click()
If lstMail_MT_CC.Visible Then lstMail_MT_CC_TXT
End Sub

Private Sub lstMail_MT_To_Click()
If lstMail_MT_To.Visible Then lstMail_MT_To_TXT

End Sub



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

End Sub


Private Sub lstNoPaper_Click()

If blnControl Then
    Me.Enabled = False
    
    Call fgNoPaper_Files_Load(lstNoPaper.Text, "")
    
    Me.Enabled = True
End If
End Sub

Private Sub mnuContext_NoPaper_Click()
'YBIATAB0_DATE_CPT_J
X = Trim(InputBox("date <aaaammjj>", "Indiquer la date de la comptabilité ", YBIATAB0_DATE_CPT_J))
If X <> "" Then
    YBIATAB0_DATE_CPT_J = X
    Me.Enabled = False
    Auto_NoPaper
    Me.Enabled = True
End If

End Sub

Private Sub mnuContext_NoPaper_Recap_Dispatch_Click()

'X = Trim(InputBox("date <aaaammjj>", "Indiquer la date de l'archive ", YBIATAB0_DATE_CPT_J))
'If X <> "" Then

    Me.Enabled = False
    Call Auto_NoPaper_Recap(paramEditionNoPaper_Folder & "PDF\" & lstNoPaper, "")
    Me.Enabled = True
'End If

End Sub


Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub


Private Sub mnuContextOptions_Click()
fraContextOptions.Visible = True
End Sub

Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
'If optSplf_RefreshActivate Then lstCourrier_Load: Debug.Print "Activate"
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
If fraContextOptions.Visible Then fraContextOptions.Visible = False: Exit Sub
If lstMail_MT_To.Visible Then lstMail_MT_To.Visible = False: Exit Sub
If lstMail_MT_CC.Visible Then lstMail_MT_CC.Visible = False: Exit Sub
If fraMail_MT.Visible Then fraMail_MT.Visible = False: Exit Sub


If SSTab1.Tab = 0 Then Unload Me: Exit Sub

blnControl = False
lstErr.Clear: lstErr.Height = 200
Timer1_Monitor
If currentAction = "" Then
Else
    X = MsgBox("Voulez-vous réellement abandonner la mise à jour?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)
    If X = vbYes Then
        currentAction = ""
    Else
        Exit Sub
    End If
End If

End Sub

Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
    'If Not fgSelect.Visible Then cmdSelect_Ok_Click
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
fgSelect_FormatString = fgSelect.FormatString
fgNoPaper_FormatString = fgNoPaper.FormatString

End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If UnloadMode = 0 Then Cancel = Form_QueryUnload_Msgbox
End Sub

Private Sub Form_Resize()
Timer1_Monitor
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



Private Sub mnuContextRefresh_Click()
lstCourrier_Load
End Sub

Private Sub mnuContextReset_Click()
On Error Resume Next
Call cbo_Scan(constProduction, cboSPLF_Folder)
cboSPLF_Amj.ListIndex = 0
cboSPLF_User.ListIndex = 0
cboSPLF_Unit.ListIndex = 0
txtSPLF_PageStart = 1
txtSPLF_PageEnd = 10

End Sub

Private Sub mnuNoPaper_Archiver_Click()
On Error GoTo Error_Handle
Me.Enabled = False
Call FEU_ROUGE

msFileSystem.MoveFile fgNoPaper.Text, Replace(fgNoPaper.Text, "\Prod_", "\Archive_")
lstNoPaper_Click

Me.Enabled = True
Me.Show
Call FEU_VERT
Exit Sub

Error_Handle:
MsgBox fgNoPaper.Text & ":" & Error, vbCritical, "mnuNoPaper_Archiver_Click"
Me.Enabled = True

End Sub

Private Sub mnuNoPaper_Display_Click()
Me.Enabled = False
Call frmElpPrt.Windows_Display_File(fgNoPaper.Text)
Me.Enabled = True
Me.Show
End Sub

Private Sub mnuNoPaper_Mail_Click()
Dim X As String, K As Integer, xMessage As String, xDestinataire_Mail As String
Dim msFile_doc As Scripting.TextStream
Dim xDestinataire As String, xFile As String, xLib As String, xAMJ_HMS  As String, xJob As String
On Error GoTo Error_Handle
Me.Enabled = False

txtMail_MT_To = usrName_UCase

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


fraMail_MT.Visible = True


'_____________________________________________________________________________________________

fgNoPaper.Col = fgNoPaper_arrIndex: xFile = Trim(fgNoPaper.Text)
fgNoPaper.Col = 1: xDestinataire = Trim(fgNoPaper.Text)
fgNoPaper.Col = 2: xLib = Trim(fgNoPaper.Text)
fgNoPaper.Col = 3: xAMJ_HMS = Trim(fgNoPaper.Text)
fgNoPaper.Col = 4: xJob = Trim(fgNoPaper.Text)

txtMail_Mt_Subject = xLib & "  " & xAMJ_HMS

mMail_Lnk = "<B>" & htmlFontColor_Blue & xLib & "..... </B>" _
      & "<A HREF=" & Asc34 & "\\" & xFile & Asc34 & ">" _
       & "(" & xDestinataire _
     & " - " & xAMJ_HMS & " / " & xJob _
     & ")</A>"
'_____________________________________________________________________________________________

mMail_Txt = ""
chkMail_Mt_Detail.Enabled = False

If InStr(xFile, ".pdf") > 0 Then
   X = Replace(xFile, paramEditionNoPaper_Partage, paramEditionNoPaper_Folder & "DOC\")
   X = Replace(X, ".pdf", ".doc")
   
   If Dir(X) <> "" Then
        chkMail_Mt_Detail.Enabled = True
        chkMail_Mt_Detail = "1"
       Set msFile = msFileSystem.GetFile(X)
       Set msFile_doc = msFile.OpenAsTextStream(ForReading)
       
       
       X = msFile_doc.ReadLine
       X = msFile_doc.ReadLine:
       K = InStr(X, "\pard")
       K = InStr(K, X, " ")
       If K > 0 Then mMail_Txt = mMail_Txt & Mid$(X, K, Len(X) - K) & "<BR>"
       For K = 1 To 1000
       
           X = msFile_doc.ReadLine
           If Len(X) < 2 Then Exit For
           mMail_Txt = mMail_Txt & Replace(X, "\par ", "") & "<BR>"
       Next K
       
       mMail_Txt = Replace(mMail_Txt, "\b0 ", "<\b>")
       mMail_Txt = Replace(mMail_Txt, "\b ", "<b>")
       mMail_Txt = Replace(mMail_Txt, "\'e0", "à")
       mMail_Txt = Replace(mMail_Txt, "\'e9", "é")
       mMail_Txt = Replace(mMail_Txt, "\'e8", "è")
       mMail_Txt = Replace(mMail_Txt, "\'ea", "ê")
       mMail_Txt = Replace(mMail_Txt, "\'e7", "ç")
       mMail_Txt = Replace(mMail_Txt, "\'b0", Chr$(186))
       
       msFile_doc.Close
   End If
End If
txtMail_MT_Message = ""

lstMail_MT_To.Visible = False
lstMail_MT_CC.Visible = False
fraMail_MT.Visible = True


Me.Enabled = True
Exit Sub

Error_Handle:
MsgBox fgNoPaper.Text & ":" & Error, vbCritical, "mnuNoPaper_Archiver_Click"

Exit_sub:

Me.Enabled = True

End Sub

Private Sub mnuNoPaper_Recap_Click()
Dim X As String

Me.Enabled = False
'    Call frmEdition.Auto_NoPaper_Recap(paramEditionNoPaper_Folder & "PDF\Archive_" & YBIATAB0_DATE_CPT_JP0, "")
Call frmEdition.Auto_NoPaper_Recap(paramEditionNoPaper_Folder & "PDF\" & lstNoPaper, "*") 'liste des documents affichés => utilisateur
Me.Enabled = True
Me.Show
End Sub

Private Sub mnuNoPaper_Supprimer_Click()
On Error GoTo Error_Handle

X = MsgBox(fgNoPaper.Text, vbYesNo, "Confirmez-vous la suppression définitive de ce fichier ?")
If X = vbYes Then

    Me.Enabled = False
    
    'msFileSystem.DeleteFile fgNoPaper.Text
    If Dir(fgNoPaper.Text) <> "" Then Kill fgNoPaper.Text
    
    lstNoPaper_Click
    
    Me.Enabled = True
End If
Me.Show
Exit Sub

Error_Handle:
MsgBox fgNoPaper.Text & ":" & Error, vbCritical, "mnuNoPaper_Supprimer_Click"
Me.Enabled = True

End Sub

Private Sub mnuRecapTOUS_Click()
Dim reponse As String
Dim newdate As String

    newdate = dateAAAAMMJJTOJJ_MM_AAAA(YBIATAB0_DATE_CPT_J)
    reponse = InputBox("Veuillez préciser la date de l'archive NoPaper, svp", "Édition du récapitulatif NoPaper", newdate)
    If reponse <> "" Then
        If IsDate(reponse) Then
            newdate = Replace(reponse, "/", "")
            newdate = Right(newdate, 4) & Mid(newdate, 3, 2) & Left(newdate, 2)
            Call frmEdition.Auto_NoPaper_Recap(paramEditionNoPaper_Folder & "PDF\Archive_" & newdate, "")
            Call MsgBox("Fin du Recap_NoPaper...")
        Else
            Call MsgBox("La date est incorrecte !")
        End If
    End If
    
End Sub

Private Sub mnuSplf_Afficher_Click()

Me.Enabled = False
mPageStart = 1: mPageEnd = 10000

mnuSplf_Afficher_Ok

End Sub

Private Sub mnuSplf_Afficher_Ok()

On Error GoTo Error_Handle
Me.Enabled = False
Me.MousePointer = vbHourglass

mnuSplf_Read filDoc.PATH & "\" & fgSelect_FileName

frmRTF.WindowState = vbMaximized
frmRTF.Show vbModal
Unload frmRTF
Me.Enabled = True
Me.MousePointer = 0
Exit Sub

Error_Handle:
Close
MsgBox filDoc.PATH & "\" & fgSelect_FileName & ":" & Error, vbCritical, "mnuSplf_Afficher_Click"
Me.Enabled = True
Me.MousePointer = 0


End Sub

Private Sub mnuSplf_Save_Ok(lFilDoc As String, lFileName As String, lDir_Save As String, lFileName_Save As String)
Dim wFileName As String
'On Error GoTo Error_Handle
Me.Enabled = False
'Me.MousePointer = vbHourglass

mnuSplf_Read lFilDoc & "\" & lFileName
frmRTF_Caller = "frmEdition  SAVE   "
frmRTF.Msg_Rcv ("frmEdition     ")

If Not msFileSystem.FolderExists(lDir_Save) Then MkDir lDir_Save

wFileName = "C:\Temp\" & paramIMP_PDFCreator_Name & ".rtf"
If Dir(wFileName) <> "" Then Kill wFileName
frmRTF.txtRTF.SaveFile wFileName
'---------------------------------------------------------------------------
Dim msFile As Scripting.File
Dim msFile_rtf As Scripting.TextStream, msFile_doc As Scripting.TextStream

Set msFile = msFileSystem.GetFile(wFileName)
Set msFile_rtf = msFile.OpenAsTextStream(ForReading)

X = msFile_rtf.ReadAll
X = Replace(X, "\'87", "\page")

If frmRTF_prtOrientation = vbPRORPortrait Then
    X = Replace(X, "\viewkind4", "\viewkind4\paperw12240\paperh15840\margl720\margr720\margt720\margb720\gutter0\ltrsect")
Else
    X = Replace(X, "\viewkind4", "\viewkind4\paperw16838\paperh11906\margl720\margr720\margt720\margb720\gutter0\ltrsect\sectd \ltrsect\lndscpsxn\linex0\headery708\footery708\colsx708\endnhere\sectlinegrid360\sectdefaultcl\sectrsid12462243\sftnbj")
End If

wFileName_Save = lDir_Save & lFileName_Save

Set msFile_doc = msFileSystem.CreateTextFile(wFileName_Save, True)

msFile_doc.Write X

msFile_rtf.Close
msFile_doc.Close

'---------------------------------------------------------------------------


'---------------------------------------------------------------------------------
'Unload frmRTF
'Me.Enabled = True
'Me.MousePointer = 0
Exit Sub

Error_Handle:
Close
If Not blnTimer_Enabled Then MsgBox filDoc.PATH & "\" & fgSelect_FileName & ":" & Error, vbCritical, "mnuSplf_Save_Ok"
Me.Enabled = True
Me.MousePointer = 0


End Sub

Private Sub mnuSplf_Afficher_Partiel_Click()
mPageStart = Val(Trim(txtSPLF_PageStart)): mPageEnd = Val(Trim(txtSPLF_PageEnd))
mnuSplf_Afficher_Ok
End Sub

Private Sub mnuSplf_Archiver_Click()
On Error GoTo Error_Handle
Me.Enabled = False
Call FEU_ROUGE
msFileSystem.MoveFile filDoc.PATH & "\" & fgSelect_FileName, paramEditionSplf_Folder & "Archive\" & fgSelect_FileName
filDoc.Pattern = "X.X"
cboSPLF_Folder_Click
Me.Enabled = True
Call FEU_VERT
Exit Sub

Error_Handle:
MsgBox filDoc.PATH & "\" & fgSelect_FileName & ":" & Error, vbCritical, "mnuSplf_Supprimer_Click"
Me.Enabled = True

End Sub

Private Sub mnuSPLF_CACLS_Click()
On Error GoTo Error_Handle



Me.Enabled = False
Me.MousePointer = vbHourglass

mnuSPLF_CACLS_Exe

GoTo Exit_sub

Error_Handle:
MsgBox filDoc.PATH & "\" & fgSelect_FileName & ":" & Error, vbCritical, "mnuSPLF_Email_Click"

Exit_sub:
Me.Enabled = True
Me.MousePointer = 0
End Sub

Private Sub mnuSPLF_Email_Click()
Dim X As String, xSubject As String

On Error GoTo Error_Handle
Me.Enabled = False
Me.MousePointer = vbHourglass
X = Trim(InputBox("<nom>.<initiales du prénom>  ", "Indiquer l'adresse de messagerie chez @bia-paris.fr ", currentSSIWINMAIL))
If X = "" Then GoTo Exit_sub

mPageStart = 1: mPageEnd = 100
'mnuSplf_Read filDoc.PATH & "\" & fgSelect_FileName
Call mnuSplf_Save_Ok(filDoc.PATH, fgSelect_FileName, "C:\Temp\SAB_", fgSelect_FileName)

If InStr(X, "@") > 0 Then
    If InStr(X, paramSendMail_BIA_URL) = 0 Then
        Call MsgBox("adresse de messagerie  obligatoire : " & paramSendMail_BIA_URL, vbCritical, "frmEdition > mnuSPLF_Email_Click")
        GoTo Exit_sub
    End If

Else
    X = X & paramSendMail_BIA_URL
End If

fgSelect.Col = 2: xSubject = "SAB : " & Trim(fgSelect.Text)
fgSelect.Col = 1: xSubject = xSubject & "   (" & LCase$(Trim(fgSelect.Text))
fgSelect.Col = 3: xSubject = xSubject & "   " & Trim(fgSelect.Text)
fgSelect.Col = 4: xSubject = xSubject & "  " & Trim(fgSelect.Text) & ")"
'"SAB : " & fgSelect_FileName
Call Email_Standard(X, xSubject, frmRTF.txtRTF.Text, False, wFileName_Save)   '"")
Unload frmRTF
Call lstErr_Clear(lstErr, cmdContext, "> Mail : " & X): DoEvents
GoTo Exit_sub

Error_Handle:
MsgBox filDoc.PATH & "\" & fgSelect_FileName & ":" & Error, vbCritical, "mnuSPLF_Email_Click"

Exit_sub:
Me.Enabled = True
Me.MousePointer = 0


End Sub

Private Sub Auto_NoPaper()
Dim xFileName As String, xFileName_Unit_Doc As String, xFileName_Unit_PDF As String, K As Integer, K1 As Integer, K2 As Integer, lenX As Integer
Dim X As String, xDestinataire As String, xDestinataire_Mail As String
Dim wAmj As String, wAMJ_HMS As String, wJob As String, wEdition_Form_Id As String
Dim blnDestinataire_Ok As Boolean, blnOk As Boolean, xSubject As String, xMessage As String
Dim xDir_Save_Doc As String, xDir_Save_PDF As String, xDir_Partage_PDF As String
Dim blnService As Boolean
Dim xText As String, xLnk As String
Dim wSendMail As typeSendMail
Dim mSSIUSRUNIT As String
Dim wUser_CACLS As String, wUnit_CACLS As String
Dim xFile_PDF As String
Dim fsoFile As Scripting.File, blnFile2Big As Boolean, lenFile2Big As Long
Dim blnDoc2PDF_Ok As Boolean
Dim SssSys As Long

On Error GoTo Error_Handle

'Automate : après 3 erreurs consécutives => temporisation
'------------------------------------------------------------------
If mMakePDF_Error > 2 Then
    mMakePDF_Error_Loop = mMakePDF_Error_Loop + 1
    If mMakePDF_Error_Loop < 30 Then Exit Sub
End If
mMakePDF_Error = 0: mMakePDF_Error_Loop = 0

blnMakePDF_Actif = True
'SssSys = Time_Sys_Sss + 1 * 60 * 3 'Exit si boucle > 3 minutes
SssSys = Time_Sys_Sss + 1 * 60 * 5 'Exit si boucle > 5 minutes

txtSPLF_PageStart = 1: txtSPLF_PageEnd = 100000
mPageStart = 1: mPageEnd = 10000

If paramEnvironnement = constTest Then
    filDoc.PATH = paramEditionNoPaper_Folder & "JPL\"
Else
    filDoc.PATH = paramEditionNoPaper_Folder & "TXT\"
End If
filDoc.Pattern = "xx*.XXX"
filDoc.Pattern = "*.txt"

For K = 0 To filDoc.ListCount - 1

    If Time_Sys_Sss > SssSys Then Exit Sub 'Ne pas boucler trop longtemps (nb et taille des spoules)
    
    filDoc.ListIndex = K
    blnOk = False
    blnDestinataire_Ok = False
    mSSIUSRUNIT = ""
    wUser_CACLS = "": wUnit_CACLS = ""
    
    xFileName = filDoc.FileName
    blnFile2Big = False
    If Mid$(xFileName, 1, 1) = "_" Then
        lenFile2Big = 2000000
    Else
        lenFile2Big = 1000000   ' fil de l'eau
    End If
    
    Set fsoFile = msFileSystem.GetFile(filDoc.PATH & "\" & xFileName)
    If fsoFile.Size > lenFile2Big Then  '$JPL 2014-11-19 !!! voir Auto_NoPaper et mnuSplf_Read !!!
                                    '==============================================================
        blnFile2Big = True
        If Dir(filDoc.PATH & "\" & xFileName) <> "" Then Kill filDoc.PATH & "\" & xFileName
    End If
   
   If Not blnFile2Big Then
        K1 = InStr(1, xFileName, ".")
        If Mid$(xFileName, 1, 1) = "_" Then
            blnService = True
           wUnit_CACLS = Mid$(xFileName, 2, K1 - 2)
     '____________________________________________________________________________________________________
            If InStr(1, xFileName, "AUT329") Then
                Dim msFile As Scripting.File
                Dim msFile_rtf As Scripting.TextStream
                
                Set msFile = msFileSystem.GetFile(filDoc.PATH & "\" & xFileName)
                Set msFile_rtf = msFile.OpenAsTextStream(ForReading)
                X = msFile_rtf.ReadAll
                If InStr(1, X, "Responsable :R50") > 0 Then wUnit_CACLS = "DER"
                If InStr(1, X, "Responsable :R60") > 0 Then wUnit_CACLS = "DRH"
                If InStr(1, X, "Responsable :R80") > 0 Then wUnit_CACLS = "DTX"
                If InStr(1, X, "Responsable :R85") > 0 Then wUnit_CACLS = "DEON"
                If Len(X) < 1500 Then wUnit_CACLS = "S99"
                msFile_rtf.Close
            End If
    '____________________________________________________________________________________________________
           mSSIUSRUNIT = Table_Unit_SSI("", wUnit_CACLS)
        Else
            blnService = False
            If K1 > 0 Then
                xDestinataire = Mid$(xFileName, 1, K1 - 1)
                 X = "select SSIWINUIDX , SSIWINMAIL , SSIUSRUNIT from " & paramIBM_Library_SABSPE & ".YSSIWIN0 ," & paramIBM_Library_SABSPE & ".YSSIDOM0 ," _
                    & paramIBM_Library_SABSPE & ".YSSIUSR0" _
                    & " where ssiwinuidd = ssidomuidd and ssidomnat = ' ' and ssidomdidx = 'WIN' and ssidomstak = ' '" _
                    & " and   ssidomuidn = (select ssidomuidn from " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
                    & " where ssidomnat = ' ' and ssidomdidx = 'IBM' and ssidomuidx = '" & xDestinataire & "')" _
                    & " and ssiusruidn = ssidomuidn"
                Set rsSab = cnsab.Execute(X)
                If Not rsSab.EOF Then
                    blnDestinataire_Ok = True
                    xDestinataire = Trim(rsSab("SSIWINUIDX"))
                    xDestinataire_Mail = Trim(rsSab("SSIWINMAIL"))
                    mSSIUSRUNIT = Trim(rsSab("SSIUSRUNIT"))
                    wUnit_CACLS = Table_Unit_SSI("S", mSSIUSRUNIT)
                    wUser_CACLS = xDestinataire
                End If
            End If
        End If
        If mSSIUSRUNIT = "" Then mSSIUSRUNIT = "S00"
        xFileName_Unit_Doc = Replace(xFileName, ".txt", " (" & mSSIUSRUNIT & ").doc")
        xFileName_Unit_PDF = Replace(xFileName, ".txt", " (" & mSSIUSRUNIT & ").pdf")
        Call Auto_NoPaper_lnk(filDoc.PATH, xFileName_Unit_PDF, xText, xLnk, wAmj)
        If blnService Then
            If xEdition_Form.NoPaper_Prod = "1" Then    '$JPL 2014-12-15 Ne pas archiver les .pdf
                xDir_Save_Doc = paramEditionNoPaper_Folder & "DOC\" & "Prod_" & wAmj & "\"
                xDir_Save_PDF = paramEditionNoPaper_Folder & "PDF\" & "Prod_" & wAmj & "\"
                xDir_Partage_PDF = paramEditionNoPaper_Partage & "Prod_" & wAmj & "\"
            Else
                xDir_Save_Doc = paramEditionNoPaper_Folder & "DOC\" & "Archive_" & wAmj & "\"
                xDir_Save_PDF = paramEditionNoPaper_Folder & "PDF\" & "Archive_" & wAmj & "\"
                xDir_Partage_PDF = paramEditionNoPaper_Partage & "Archive_" & wAmj & "\"
            End If
            Call mnuSplf_Save_Ok(filDoc.PATH, xFileName, xDir_Save_Doc, xFileName_Unit_Doc)
            'DR 11/09/2019 On ne test plus si makePDF actif sur BIA2008
'            If nomDuServeur = paramServerSplf Then
                blnDoc2PDF_Ok = Auto_NoPaper_MakePDF(xFileName_Unit_Doc, xDir_Save_Doc, xDir_Save_PDF)
'            Else
'                If blnMakePDF_Actif Then
'                    blnDoc2PDF_Ok = Auto_NoPaper_MakePDF(xFileName_Unit_Doc, xDir_Save_Doc, xDir_Save_PDF)
'                Else
'                   If Not blnTimer_Enabled Then
'                        blnDoc2PDF_Ok = Auto_NoPaper_PDFCreator(xFileName_Unit_Doc, xDir_Save_Doc, xDir_Save_PDF)
'                    End If
'                End If
'            End If
            If blnDoc2PDF_Ok Then
                 If Dir(filDoc.PATH & "\" & xFileName) <> "" Then Kill filDoc.PATH & "\" & xFileName
           End If
       Else
            If blnDestinataire_Ok Then
                If Mid$(xDestinataire, 1, 2) = "T_" Then
                    xDir_Save_Doc = paramEditionNoPaper_Folder & "DOC\" & "Test_" & wAmj & "\"
                    xDir_Save_PDF = paramEditionNoPaper_Folder & "PDF\" & "Test_" & wAmj & "\"
                    xDir_Partage_PDF = paramEditionNoPaper_Partage & "Test_" & wAmj & "\"
                Else
                    If xEdition_Form.Save = "1" Then
                        xDir_Save_Doc = paramEditionNoPaper_Folder & "DOC\" & "Archive_" & wAmj & "\"
                        xDir_Save_PDF = paramEditionNoPaper_Folder & "PDF\" & "Archive_" & wAmj & "\"
                        xDir_Partage_PDF = paramEditionNoPaper_Partage & "Archive_" & wAmj & "\"
                    Else
                        xDir_Save_Doc = paramEditionNoPaper_Folder & "DOC\" & "Prod_" & wAmj & "\"
                        xDir_Save_PDF = paramEditionNoPaper_Folder & "PDF\" & "Prod_" & wAmj & "\"
                        xDir_Partage_PDF = paramEditionNoPaper_Partage & "Prod_" & wAmj & "\"
                    End If
                End If
                Call mnuSplf_Save_Ok(filDoc.PATH, xFileName, xDir_Save_Doc, xFileName_Unit_Doc)
                'DR 11/09/2019 On ne test plus si makePDF actif sur BIA2008
'                If nomDuServeur = paramServerSplf Then
                    blnDoc2PDF_Ok = Auto_NoPaper_MakePDF(xFileName_Unit_Doc, xDir_Save_Doc, xDir_Save_PDF)
'                Else
'                    If blnMakePDF_Actif Then
'                       blnDoc2PDF_Ok = Auto_NoPaper_MakePDF(xFileName_Unit_Doc, xDir_Save_Doc, xDir_Save_PDF)
'                    Else
'                       If Not blnTimer_Enabled Then
'                            blnDoc2PDF_Ok = Auto_NoPaper_PDFCreator(xFileName_Unit_Doc, xDir_Save_Doc, xDir_Save_PDF)
'                        End If
'                    End If
'                End If
                xLnk = Replace(xLnk, filDoc.PATH & "\", xDir_Partage_PDF)
                If blnDoc2PDF_Ok Then
                    If Dir(filDoc.PATH & "\" & xFileName) <> "" Then Kill filDoc.PATH & "\" & xFileName
                End If
           End If
        End If
'___________________________________________________________________________________________________
        If wUnit_CACLS = "S00" Or wUnit_CACLS = "S99" Then wUser_CACLS = "": wUnit_CACLS = ""
        Call File_ICACLS(xDir_Save_Doc & xFileName_Unit_Doc, wUser_CACLS, wUnit_CACLS)    'gestion des droits  ACL
        xFile_PDF = xDir_Save_PDF & Replace(xFileName_Unit_Doc, ".doc", ".pdf")
        Call File_ICACLS(xFile_PDF, wUser_CACLS, wUnit_CACLS)  'gestion des droits  ACL
    '___________________________________________________________________________________________________
    End If 'blnFile2Big
'Automate : Tester si MakePDF est actif
'------------------------------------------------------------------
    'DR 12/08/2019 On ne test plus si makePDF actif sur BIA2008
    'If blnTimer_Enabled And Not blnMakePDF_Actif Then Exit Sub
    If mMakePDF_Error > 2 Then mMakePDF_Error_Loop = 1: Exit Sub
Next K
GoTo Exit_sub

Error_Handle:
If Not blnTimer_Enabled Then
    Call ecrit_erreur(Err.Number & " " & xFileName & ": " & Err.Description & " " & Printer.Devicename)
    'MsgBox Err.Number & " " & xFileName & ": (2) " & Error & " " & Err.Description & " " & Printer.Devicename, vbCritical, "frmEdition > Auto_NoPaper"
    Resume
End If
Exit_sub:


End Sub



Private Sub mnuSplf_Imprimer_Click()
Dim K As Integer, X As String

Me.Enabled = False: Screen.MousePointer = vbHourglass
mPageStart = 1: mPageEnd = 10000

mnuSplf_Imprimer_Ok

End Sub

Private Sub mnuSplf_Imprimer_Ok()
Dim K As Integer, X As String
Dim wFileName_Print As String

On Error GoTo Error_Handle
Me.Enabled = False: Screen.MousePointer = vbHourglass

            blnFiligrane = False 'True
            prtFiligrane_Name = paramEditionFiligrane_Folder & "Test" & ".jpg" ' ".bmp"
If UCase$(Trim(cboSPLF_Folder)) = "TEST" Then
    blnFiligrane = True
    'cboSPLF_Filigrane.ListIndex = k_Filigrane_Test
    prtFiligrane_Name = paramEditionFiligrane_Folder & "Test" & ".jpg"

End If

wFileName_Print = filDoc.PATH & "\" & fgSelect_FileName

X = vbYes
K = InStr(1, fgSelect_FileName, "ECHAVI02P1")
If K > 0 Then
   ' X = MsgBox("Version Test : mémorisation ", vbYesNo + vbQuestion + vbDefaultButton2, "Impression des AVIS ECHELLES")
   ' If X = vbYes Then Call prtSAB_Echelles_ECHAVI02P1(filDoc.PATH & "\" & fgSelect_FileName): GoTo Exit_Sub
    X = MsgBox("Vider les bacs 2 & 3 de l'imprimante, alimenter le bac 1 avec du papier au format A5", vbYesNo + vbQuestion + vbDefaultButton2, "Impression des AVIS ECHELLES")
    If X <> vbYes Then GoTo Exit_sub
    
       
End If

K = InStr(1, fgSelect_FileName, "ECHEDI01P2")
If K > 0 Then
    'X = MsgBox("Version Test : Impression ", vbYesNo + vbQuestion + vbDefaultButton2, "Impression des AVIS ECHELLES")
    'If X = vbYes Then Call prtSAB_Echelles_ECHEDI01P2(filDoc.PATH & "\" & fgSelect_FileName): GoTo Exit_Sub
    X = MsgBox("Impression BIA :" & Asc10_13 & " AS400 ==> Préparation,Archivage" & Asc10_13 & "NT  ==>  FTP , Impression", vbYesNo + vbQuestion + vbDefaultButton2, "Impression des RELEVES ECHELLES")
    If X = vbYes Then Call prtSAB_Echelles_FTP(filDoc.PATH & "\" & fgSelect_FileName, Me)
    GoTo Exit_sub
End If

K = InStr(1, fgSelect_FileName, "SITTE003P1")
If K > 0 Then
    Call prtSAB_SITTE003P1(filDoc.PATH & "\" & fgSelect_FileName)    ', Me)
    GoTo Exit_sub
End If

K = InStr(1, fgSelect_FileName, "ECHEDI04P1")
If K > 0 Then
'2015-10-01 JPL     X = MsgBox(" OUI : sélection des clients dont la racine est > 50000" & Asc10_13 & "NON : spoule natif SAB", vbYesNoCancel + vbQuestion + vbDefaultButton1, "ECHEDI04P1 : Impression Etat de contrôle des échelles :")
     X = MsgBox(" OUI : exclure les comptes techniques (N..., R....) " & Asc10_13 & "NON : spoule natif SAB", vbYesNoCancel + vbQuestion + vbDefaultButton1, "ECHEDI04P1 : Impression Etat de contrôle des échelles :")
    If X = vbCancel Then GoTo Exit_sub
    
    If X = vbYes Then
    'extraction dans un fichier local
        Call prtSAB_ECHEDI04P1(wFileName_Print, paramFolder_Local & "\" & fgSelect_FileName)
        wFileName_Print = paramFolder_Local & "\" & fgSelect_FileName
    Else
        X = vbYes
    End If
End If

If X = vbYes Then
    
    mnuSplf_Read wFileName_Print
    'tester si le document ontien "Rien à imprime, Aucun avis à éditer...
    
    '                                       '
    frmRTF_Caller = "frmEdition  Print  "
    
    frmRTF.Tag = fgSelect_FileName 'DRDR
    
    frmRTF.Msg_Rcv ("frmEdition     ")
    Unload frmRTF
End If

GoTo Exit_sub

Error_Handle:
MsgBox filDoc.PATH & "\" & fgSelect_FileName & ":" & Error, vbCritical, "mnuSplf_Imprimer_Click"

Exit_sub:
Me.Enabled = True: Screen.MousePointer = 0
End Sub


Private Sub mnuSplf_Imprimer_Partiel_Click()
mPageStart = Val(Trim(txtSPLF_PageStart)): mPageEnd = Val(Trim(txtSPLF_PageEnd))
mnuSplf_Imprimer_Ok
End Sub

Private Sub mnuSPLF_Save_Click()
Me.Enabled = False ': Screen.MousePointer = vbHourglass
mPageStart = 1: mPageEnd = 10000
Call lstErr_Clear(lstErr, cmdContext, "> enregistrer sous ... "): DoEvents

Call mnuSplf_Save_Ok(filDoc.PATH, fgSelect_FileName, "C:\Temp\SAB_", fgSelect_FileName)

Call lstErr_Clear(lstErr, cmdContext, "> " & wFileName_Save): DoEvents

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuSplf_Supprimer_Click()
On Error GoTo Error_Handle
Me.Enabled = False

Kill filDoc.PATH & "\" & fgSelect_FileName
filDoc.Pattern = "X.X"
cboSPLF_Folder_Click
Me.Enabled = True

Exit Sub

Error_Handle:
MsgBox filDoc.PATH & "\" & fgSelect_FileName & ":" & Error, vbCritical, "mnuSplf_Supprimer_Click"
Me.Enabled = True

End Sub


Private Sub optNoPaper_Archive_Click()
Me.Enabled = False

mNoPaper_Opt = "Archive_"
lstNoPaper_Folders_Load

Me.Enabled = True
End Sub

Private Sub optNoPaper_Prod_Click()
Me.Enabled = False

mNoPaper_Opt = "Prod_"
lstNoPaper_Folders_Load

Me.Enabled = True

End Sub


Private Sub optNoPaper_Test_Click()
Me.Enabled = False

mNoPaper_Opt = "Test_"
lstNoPaper_Folders_Load

Me.Enabled = True

End Sub

Private Sub optSplf_RefreshActivate_Click()
fraContextOptions.Visible = False
Timer1_Monitor
End Sub

Private Sub optSplf_RefreshNo_Click()
fraContextOptions.Visible = False
Timer1_Monitor
End Sub

Private Sub optSplf_RefreshTimer_Click()
fraContextOptions.Visible = False
Timer1_Monitor
End Sub


Private Sub RecapTOUS_Click()

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Timer1_Monitor
End Sub

Private Sub Timer1_Timer()
Me.Enabled = False
lstCourrier_Load
Me.Enabled = True
End Sub









Public Sub lstCourrier_Load_filDoc(lFolder As String, lUser As String)
On Error Resume Next
Me.Enabled = False
Screen.MousePointer = vbHourglass
libSelect.Caption = "Rachaîchissement : " & Time
filDoc.PATH = paramEditionSplf_Folder & lFolder & "\"

fgSelect.Visible = False
''filDoc.Pattern = lUser & "*.XXX"
''filDoc.Pattern = lUser & "*.*"
''If mfilDoc_List = "" Then
''    filDoc.ListIndex = filDoc.ListCount - 1
''Else
''    Call fileListBox_Scan(mfilDoc_List, filDoc)
''End If

filDoc.Pattern = "xx*.XXX"
filDoc.Pattern = lUser & "*.*"

If lFolder = "Archive" Then
    mnuSplf_Archiver.Enabled = False
Else
    mnuSplf_Archiver.Enabled = True
End If
fgSelect.Visible = True

Screen.MousePointer = vbDefault
Me.Enabled = True

End Sub

Public Sub lstNoPaper_Load_filDoc(lFolder As String, lUser As String)
On Error Resume Next
'Me.Enabled = False
'Screen.MousePointer = vbHourglass
libSelect.Caption = "Rachaîchissement : " & Time
filDoc.PATH = paramEditionNoPaper_Folder & lFolder & "\"

'fgSelect.Visible = False

filDoc.Pattern = "xx*.XXX"
filDoc.Pattern = lUser & "*.*"

If lFolder = "Archive" Then
    mnuSplf_Archiver.Enabled = False
Else
    mnuSplf_Archiver.Enabled = True
End If
'fgSelect.Visible = True

'Screen.MousePointer = vbDefault
'Me.Enabled = True

End Sub


Public Sub fgNoPaper_Files_Load(lFolder As String, lUser As String)
Dim kFilDoc As Integer
Dim X As String, K As Integer, K1 As Integer, K2 As Integer
Dim lenX As Integer, blnOk As Boolean
Dim wEdition_Form_Id As String
Dim wAmj As String, wJob As String, wUnit As String, wFileName As String
On Error Resume Next

If lFolder = "" Then
    Call MsgBox("Sélectionner un répertoire", vbInformation, "Edition NoPaper")
    GoTo Error_MsgBox
End If
'fgNoPaper.Clear
ProgressBar1.Visible = True

ProgressBar1.Min = 0: ProgressBar1.Max = 1000
ProgressBar1.Value = 1
fgNoPaper.Visible = False: DoEvents

fgNoPaper_Reset

fgNoPaper.Rows = 1
fgNoPaper.FormatString = fgNoPaper_FormatString
mNoPaper_Doc = Trim(txtNoPaper_Doc)




mNoPaper_Folder = paramEditionNoPaper_Partage & lFolder & "\"
''Set msFolder = msFileSystem.GetFolder(mNoPaper_Folder)
''For Each msFile In msFolder.Files
filDoc.PATH = mNoPaper_Folder
filDoc.Pattern = "xx.XXX"
If mNoPaper_User <> "" Then
    filDoc.Pattern = "*" & mNoPaper_User & "*.*"
Else
    If mNoPaper_Unit <> "" Then
        filDoc.Pattern = "*(" & mNoPaper_Unit & ")*.*"
    Else
        filDoc.Pattern = "*.*"
    End If
End If



For kFilDoc = 0 To filDoc.ListCount - 1

    filDoc.ListIndex = kFilDoc
    wFileName = filDoc.FileName
    blnOk = True
    If InStr(1, wFileName, ".db") > 0 Then blnOk = False
    
'    If mNoPaper_User <> "" Then
'        If InStr(1, wFileName, mNoPaper_User) = 0 Then blnOk = False
'    End If
   
    
    If mNoPaper_Unit <> "" Then
        If InStr(1, wFileName, "(" & mNoPaper_Unit & ")") = 0 Then blnOk = False
    End If
  
'    If blnOk Then
        K = InStr(1, wFileName, ".")
        If K > 0 Then
            xUser.Id = Mid$(wFileName, 1, K - 1)
        Else
            xUser.Id = ""
            xUser.Unit = ""
        End If
        
        K1 = K + 1
        K = InStr(K1, wFileName, "_")
        If K > 0 Then
            K2 = InStr(K + 1, wFileName, "_")
            If K2 > 0 Then
                wAmj = Mid$(wFileName, K1, K2 - K1)
                K1 = K2 + 1
                K2 = InStr(K1, wFileName, "_")
                If K2 > 0 Then
                    wEdition_Form_Id = Mid$(wFileName, K1, K2 - K1)
                    K1 = K2 + 1
                    K2 = InStr(K1, wFileName, "(")
                    If K2 > 0 Then
                        wUnit = Mid$(wFileName, K2 + 1, 3)
                        wJob = Mid$(wFileName, K1, K2 - K1)
                    End If
                    
                End If
            End If
        End If
        
        If mNoPaper_Doc <> "" Then
            If InStr(1, UCase(wEdition_Form_Id), mNoPaper_Doc) = 0 Then blnOk = False
        End If
 '  End If
    If blnOk Then
        fgNoPaper.Rows = fgNoPaper.Rows + 1
        fgNoPaper.Row = fgNoPaper.Rows - 1
        fgNoPaper.Col = fgNoPaper_arrIndex: fgNoPaper.Text = mNoPaper_Folder & wFileName
        fgNoPaper.Col = 0: fgNoPaper.Text = wUnit
        fgNoPaper.Col = 1: fgNoPaper.Text = xUser.Id
        fgNoPaper.Col = 3: fgNoPaper.Text = Format(wAmj, "@@@@-@@-@@ @@@:@@:@@")
        fgNoPaper.Col = 4: fgNoPaper.Text = wJob
     
        
        fgNoPaper.Col = 2
        xEdition_Form.K1 = "SAB"
        xEdition_Form.K2 = wEdition_Form_Id
        fgNoPaper.Text = wEdition_Form_Id & " " & rsEdition_Form(xEdition_Form)
    End If
    ProgressBar1.Value = ProgressBar1.Value + 1

''Next
Next kFilDoc

If fgNoPaper.Rows > 1 Then fgNoPaper_Sort1 = 3: fgNoPaper_Sort2 = 3: fgNoPaper_Sort
If Mid$(lFolder, 1, 7) = "Archive" Then
    mnuNoPaper_Archiver.Enabled = False
Else
    mnuNoPaper_Archiver.Enabled = True
End If
Error_MsgBox:

fgNoPaper.Visible = True
ProgressBar1.Visible = False
End Sub


Public Sub lstCourrier_Load()
Dim xName As String
Me.Enabled = False: Me.MousePointer = vbHourglass
If Not blnForm_Init Then
    xName = Trim(cboSPLF_User)
    If xName = "" Then xName = "*" & Trim(txtSPLF_Name)
    lstCourrier_Load_filDoc Trim(cboSPLF_Folder), xName
    fgSelect_Display
End If
Me.Enabled = True: Me.MousePointer = 0
End Sub



Private Sub txtMail_MT_Message_Click()
lstMail_MT_CC.Visible = False
lstMail_MT_To.Visible = False

End Sub


Private Sub txtNoPaper_AMJ_Max_Change()
lstNoPaper_Folders_Load
End Sub


Private Sub txtNoPaper_AMJ_Min_Change()
lstNoPaper_Folders_Load
End Sub


Private Sub txtNoPaper_Doc_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
fgNoPaper.BackColor = RGB(192, 192, 192)
End Sub


Private Sub txtNoPaper_Doc_LostFocus()
fgNoPaper.Visible = False: DoEvents
lstNoPaper_Click

End Sub

Private Sub txtSPLF_Name_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtSPLF_PageEnd_GotFocus()
Call txt_GotFocus(txtSPLF_PageEnd)

End Sub


Private Sub txtSPLF_PageEnd_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtSPLF_PageEnd_LostFocus()
Call txt_LostFocus(txtSPLF_PageEnd)

End Sub


Private Sub txtSPLF_PageStart_Click()
Call txt_GotFocus(txtSPLF_PageStart)

End Sub


Private Sub txtSPLF_PageStart_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtSPLF_PageStart_LostFocus()
Call txt_LostFocus(txtSPLF_PageStart)

End Sub



Public Sub fctSelBold_Ok()
Dim I As Integer, X As String, lenX As Integer
Dim blnOn As Boolean

lenLine = Len(RTrim(xLine))

If blnSelBold Then
    X = RTrim(xLine_SelBold)
    lenX = Len(X)
    If lenX > lenLine Then lenLine = lenX
    
    blnOn = False
   
   For I = 1 To lenX
         If Mid$(X, I, 1) = Mid$(xLine, I, 1) Then
           If Not blnOn Then
                If Mid$(X, I, 1) <> " " Then
                    blnOn = True
                    arrBold_Nb = arrBold_Nb + 1
                    If UBound(arrBold_SelStart) <= arrBold_Nb Then
                        ReDim Preserve arrBold_SelStart(arrBold_Nb + 500)
                        ReDim Preserve arrBold_SelLength(arrBold_Nb + 500)
                    End If
                    
                    arrBold_SelStart(arrBold_Nb) = lenRTF + I - 1
                    arrBold_SelLength(arrBold_Nb) = 1
                End If
            Else
                 arrBold_SelLength(arrBold_Nb) = arrBold_SelLength(arrBold_Nb) + 1
           End If
            
        Else
            blnOn = False
        End If
    Next I
End If


If blnSelBold Then
    X = RTrim(xLine_SelUnderline)
    lenX = Len(X)
    If lenX > lenLine Then lenLine = lenX
    
    blnOn = False
    
    For I = 1 To lenX
        If Mid$(X, I, 1) = "_" Then
            If Not blnOn Then
                blnOn = True
                arrUnderline_Nb = arrUnderline_Nb + 1
                    If UBound(arrUnderline_SelStart) <= arrUnderline_Nb Then
                        ReDim Preserve arrUnderline_SelStart(arrUnderline_Nb + 500)
                        ReDim Preserve arrUnderline_SelLength(arrUnderline_Nb + 500)
                    End If
                arrUnderline_SelStart(arrUnderline_Nb) = lenRTF + I - 1
                arrUnderline_SelLength(arrUnderline_Nb) = 1
            Else
                 arrUnderline_SelLength(arrUnderline_Nb) = arrUnderline_SelLength(arrUnderline_Nb) + 1
           End If
            
        Else
            blnOn = False
        End If
    Next I
End If
End Sub
Public Sub fctSelBold_Scan(lTxt As String)
Dim I As Integer, X1 As String * 1
For I = 1 To Len(lTxt)
    X1 = Mid$(lTxt, I, 1)
    Select Case X1
        Case " "
        Case "_":
            Mid$(xLine_SelUnderline, I, 1) = X1: blnSelUnderline = True
        Case Else:
            If X1 = Mid$(xLine, I, 1) Then
                Mid$(xLine_SelBold, I, 1) = X1
            Else
                Mid$(xLine, I, 1) = X1
            End If
  End Select
Next I

End Sub


Public Sub fctSelBold_Reset()
xLine = ""
blnSelBold = False: xLine_SelBold = ""
blnSelUnderline = False: xLine_SelUnderline = ""

End Sub

Public Sub Timer1_Monitor()
If Me.WindowState = 1 Then
    Timer1.Enabled = False
Else
    If SSTab1.Tab = 1 And optSplf_RefreshTimer Then
        Timer1.Enabled = True
        lstCourrier_Load
    Else
        Timer1.Enabled = False
    End If
End If

End Sub

Public Sub mnuSplf_Read(lFileName As String)

'========================================================================================
' !! imprimante "IMP_AVIS"
' gère le papier au format A5
' Test en dur dans prtEdition.prtCourrier_Open
'     If prtPaperSize = vbPRPSA5 And Trim(XPrt.Devicename) = "\\Printsrv\IMP_GUICHET" Then
'========================================================================================


Dim seq As Long, xIn As String, xInRTF As String, xInTxt As String
Dim iSaut As Integer, K As Integer
Dim absSaut As Integer, mSaut As Integer
Dim xRTF As String
Dim wPage As Long
Dim xSJQFILE As String
On Error GoTo Error_Handle
Dim blnGHGVE053P1 As Boolean, blnAVI002P1 As Boolean, blnGUIVE As Boolean
Dim blnDATAVI As Boolean, blnDATAVIP1 As Boolean, blnDATAVIP6 As Boolean
Dim blnECHAVI02P1 As Boolean
Dim blnAUT329P3 As Boolean, blnPrt_Responsable As Boolean
Dim blnSITTE019P1 As Boolean
Dim blnCRIGS014P1 As Boolean

Dim blnModifier As Boolean
Dim blnIgnore As Boolean
Dim xIn_EnTete As String, blnEnTete As Boolean, blnEnTete_Init As Boolean
Dim nbLine As Integer
Dim blnTest As Boolean
'__________________________________________________________________________________________

Dim blnBanqueIslamique As Boolean, iBanqueIslamique As Integer
Dim stringRemplacement As String

ReDim arrBold_SelStart(500)
ReDim arrUnderline_SelStart(500)
ReDim arrUnderline_SelLength(500)
ReDim arrBold_SelLength(500)

seq = 0: nbLine = 0
'Denis ROSILLETTE version Bia_SabFREEFILE du 18/09/2013
Dim fic As Long
fic = FreeFile
Open lFileName For Input As #fic
    
Line Input #fic, xIn

txtModèle_RTF.TextRTF = ""
txtModèle_RTF.Font.Name = prtFontName_CourierNew
txtModèle_RTF.Font.Size = 7.4
frmRTF_prtOrientation = vbPRORLandscape
frmRTF_blnCourrier = False
xSJQFILE = Mid$(xIn, 4 + 20, 10)
xEdition_Form.K1 = "SAB"
xEdition_Form.K2 = xSJQFILE

Call rsEdition_Form(xEdition_Form)
xUser.Id = Mid$(xIn, 4 + 30, 10)
Call Table_User(xUser)

frmRTF_prtPaperSize = vbPRPSA4

If xEdition_Form.Courrier = "0" Then
   frmRTF_blnCourrier = False
Else
   frmRTF_blnCourrier = True
End If
If xEdition_Form.Orientation = "0" Then
   frmRTF_prtOrientation = vbPRORPortrait
Else
   frmRTF_prtOrientation = vbPRORLandscape
End If
txtModèle_RTF.Font.Size = xEdition_Form.FontSize
txtModèle_RTF.Font.Name = Trim(xEdition_Form.FontName)

frmRTF_blnA5 = False
blnGHGVE053P1 = False
blnECHAVI02P1 = False
blnGUIVE = False
blnAVI002P1 = False
blnDATAVI = False: blnDATAVIP1 = False: blnDATAVIP6 = False: blnBanqueIslamique = False
blnModifier = False
blnAUT329P3 = False
blnSITTE019P1 = False
blnCRIGS014P1 = False

blnEnTete = False: xIn_EnTete = "": blnEnTete_Init = False

Select Case Trim(xEdition_Form.K2)
    Case "AVI002P1":
        ' IMP_AVIS (imprimante locale) sauf traitement de nuit
        If xUser.QSYSOPR <> 1 Then
            frmRTF_blnA5 = True: blnAVI002P1 = True: frmRTF_prtPaperSize = vbPRPSA5
        End If
    Case "ECHAVI02P1": frmRTF_blnA5 = True
    'Case "CHGVE053P1": frmRTF_blnA5 = True: blnGHGVE053P1 = True
    Case "GUIVE300P2", "GUIVE303P1", "GUIVE306P1":  blnGUIVE = True
    Case "DATAVIP1": blnDATAVI = True: blnDATAVIP1 = True: blnBanqueIslamique = False
    Case "DATAVIP6": blnDATAVI = True: blnDATAVIP1 = True: blnDATAVIP6 = True: blnBanqueIslamique = False
    Case "DATAVIP3", "DATAVIP4", "DATAVIP5", "DATAVIP8": blnDATAVI = True: blnBanqueIslamique = False
    Case "AUT329P3": If xUser.QSYSOPR = "1" Then blnAUT329P3 = True
    Case "SITTE019P1": blnSITTE019P1 = True
    Case "CRIGS014P1": blnCRIGS014P1 = True
End Select

absSaut = 9999
wPage = 0
arrBold_Nb = 0: arrUnderline_Nb = 0
fctSelBold_Reset

xRTF = ""
Do Until EOF(fic)
    If Len(xRTF) > 2000000 Then  '$JPL 2014-11-19 !!! voir Auto_NoPaper et mnuSplf_Read !!!
                                 '===========================================================
        xRTF = "!! Document trop volumineux => Impression tronquée !!" & vbCrLf & vbCrLf & vbCrLf & xRTF
        Exit Do
    End If
    seq = seq + 1
    Line Input #fic, xIn
    If xIn <> "" Then
    
       blnIgnore = False
        
        If Mid$(xIn, 1, 1) <> "$" Then
        
            mSaut = absSaut
            absSaut = absSaut + Val(Mid$(xIn, 4, 1))
            iSaut = Val(Mid$(xIn, 1, 3))
 '======================================================================================================
           
            If blnDATAVI Then
            
                If iSaut = 64 Then iSaut = 0: absSaut = absSaut + 5: blnBanqueIslamique = False
                If iSaut = 65 Then iSaut = 0: absSaut = absSaut + 5: blnBanqueIslamique = False
                
                'If blnDATAVIP6 Then
                    If blnBanqueIslamique Then
                        If InStr(xIn, "Taux nominal") Then xIn = Replace(xIn, "Taux nominal  ", "Taux de profit")
                        If InStr(xIn, "Intérêts bruts") Then xIn = Replace(xIn, "Intérêts bruts", "Marge brute   ")
                        If InStr(xIn, "les intérêts non") Then xIn = Replace(xIn, "les intérêts non", "la marge brute")
                        If InStr(xIn, "nantis seront portés") Then xIn = Replace(xIn, "nantis seront portés", "non nantie sera portée")
                        If InStr(xIn, "Les intérêts nantis") Then xIn = Replace(xIn, "Les intérêts nantis", "La marge brute nantie")
                        If InStr(xIn, "capital majoré des intérêts sera porté") Then xIn = Replace(xIn, "capital majoré des intérêts sera porté", "capital sera porté")
                    
                    Else
                         If InStr(xIn, "50451") Then blnBanqueIslamique = True
                         If blnBanqueIslamique_Loop Then
                            For iBanqueIslamique = 1 To arrBanqueIslamique_Nb
                                If InStr(xIn, arrBanqueIslamique(arrBanqueIslamique_Nb)) Then blnBanqueIslamique = True: Exit For
                            Next iBanqueIslamique
                        End If
                            
                   End If
                'End If
                        
                        
                If blnDATAVIP1 Then
                    '!!!!!!!!!!!!!!!!!!!!! le test est fait à la lecture de chaque ligne !!!!!!!!!!!!!!!!!!!!!!!!!
                    If Mid$(xIn, 5, 34) = " A l'échéance, selon vos instructi" Then
                        blnModifier = True
                        xIn = Mid$(xIn, 1, 4) & " Sauf instructions contraires de votre part 3 jours avant l'échéance,"
                    End If
                    If blnModifier And Mid$(xIn, 5, 34) = " au crédit de votre compte courant" Then
                        xIn = Mid$(xIn, 1, 4) & " ce dépôt à terme sera renouvelé dans les mêmes conditions."
                    End If
                    
                    If blnModifier And Mid$(xIn, 5, 34) = " Valeur                           " Then
                        absSaut = absSaut + 3
                        xIn = Mid$(xIn, 1, 4) & Space$(40) & paramSOC_RS
                    End If
                    
                    'DR Le 07/02/2020
'                    If (InStr(xIn, "Sauf instructions contraires de votre part 3 jours avant l'échéance,")) > 0 Then
'                        stringRemplacement = "Conformément à la convention signée avec notre établissement, le renouvellement de votre" & vbCr
'                        stringRemplacement = stringRemplacement & "compte à terme devra intervenir de manière expresse au plus tard 3 jours ouvrés avant" & vbCr
'                        stringRemplacement = stringRemplacement & "l'échéance de celui-ci." & vbCr
'                        stringRemplacement = stringRemplacement & "Sans instruction expresse de votre part dans le délai requis, nous vous informons que ce" & vbCr
'                        stringRemplacement = stringRemplacement & "dépôt à terme sera clôturé et le montant en capital et intérêts sera transféré sur votre" & vbCr
'                        stringRemplacement = stringRemplacement & "compte courant."
'                        xIn = Mid$(xIn, 1, 4) & stringRemplacement
'                    End If
                    
                End If
                
            End If
'======================================================================================================
            If blnGHGVE053P1 Then
                If iSaut = 38 Then iSaut = 0: absSaut = absSaut + 2
                If Mid$(xIn, 4, 1) = "3" Then absSaut = absSaut - 2
            End If
'======================================================================================================
            If blnGUIVE Then
                 If Mid$(xIn, 4, 1) = "2" Then absSaut = absSaut - 1
            End If
'======================================================================================================
            If blnSITTE019P1 Then
                If iSaut <> 0 Then xIn = "001 Code Banque : 12179  code Guichet : 00001"
                If Mid$(xIn, 5, 3) = "SAB" Then: absSaut = absSaut + 2: Mid$(xIn, 5, 10) = Space$(10)
           End If
'======================================================================================================
            If blnCRIGS014P1 And iSaut > 0 Then
            
                If Mid$(xIn, 1, 14) <> "001 CRIGS014P1" Then
                    If Mid$(xIn, 5, 26) = "Numéro client............:" Then
                        absSaut = absSaut + 3: iSaut = 0
                    Else
                        absSaut = absSaut + 1: iSaut = 0
                    End If
                End If
            
            End If
'======================================================================================================
            If blnAUT329P3 Then
                If Mid$(xIn, 1, 4) = "001 " Then
                    If blnEnTete Then
                        absSaut = mSaut + 1: iSaut = 0
                        blnTest = False
                        Do
                            Line Input #fic, xIn
''''2003.06.11 JPL à REVOIR   If Not blnEnTete_Init Then xIn_EnTete = xIn_EnTete & xIn & Asc13 & Chr$(10)
                            K = InStr(1, xIn, "D.Compte/Code Autorisation/Dossier")
                            If K > 0 Then blnTest = True
                        Loop Until blnTest
                        blnEnTete_Init = True
                        Line Input #fic, xIn
                        Line Input #fic, xIn
                    Else
                        blnEnTete = True
                    End If
                        
                End If
              
              Select Case Mid$(xIn, 14, 10)
                Case "       CDE": blnIgnore = True
                Case "       CRE": blnIgnore = True
                Case "       ENG": blnIgnore = True
                Case "       EMP": blnIgnore = True
                Case "       GAR": blnIgnore = True
                Case "       CAM": blnIgnore = True
                Case "       PRE": blnIgnore = True
                Case "       RDE": blnIgnore = True
               End Select
               'Debug.Print xIn
              If Mid$(xIn, 74, 129) = "|                 |                 |                    |           |          |          |                 |             |    |" Then
                    X = UCase$(Mid$(xIn, 5, 60))
                    K = InStr(1, X, "GROUPE")
                    If K = 0 Then blnIgnore = True
              End If
              
              If Not blnIgnore Then
                    If Mid$(xIn, 4, 1) <> "0" Then
                        If nbLine > 64 Then
                            iSaut = 1: nbLine = 1
''''2003.06.11 JPL à REVOIR   xIn = xIn_EnTete & xIn
                        End If
                        nbLine = nbLine + 1
                    End If
                End If
                
            End If
'======================================================================================================
            If blnIgnore Then
                 absSaut = mSaut
            Else
                xInRTF = ""
               If iSaut > 0 Then
                    If iSaut < absSaut Then
                        blnModifier = False
                        absSaut = iSaut
                       ''2003.12.24  If wPage = mPageEnd Then Call MsgBox("Pages  " & mPageStart & "  à  " & mPageEnd, vbExclamation, "Lecture partielle du document"): Exit Do
                        If wPage = mPageEnd Then Exit Do
                       wPage = wPage + 1
                       ' If wPage >= 126 Then
                       '     Debug.Print wPage, Len(xRTF)
                       ' End If
                       If wPage > mPageStart Then xInRTF = Asc13 & Chr$(10) & Chr$(135) & Asc13 & Chr$(10)
                    Else
                        absSaut = iSaut
                    End If
                End If
                 For K = mSaut + 1 To absSaut
                     xInRTF = xInRTF & Asc13 & Chr$(10)
                 Next K
            
                If wPage >= mPageStart Then
                    If Len(xIn) > 4 Then
                        xInTxt = Mid$(xIn, 5, Len(xIn) - 4)
                    Else
                        xInTxt = ""
                    End If
                    
                    If Mid$(xIn, 4, 1) = "0" Then
                        blnSelBold = True
                        fctSelBold_Scan xInTxt
                    Else
                       fctSelBold_Ok
                        xRTF = xRTF & txtModèle_RTF.Text & Mid$(xLine, 1, lenLine) & xInRTF
                        lenRTF = Len(xRTF)
                        fctSelBold_Reset
                        xLine = xInTxt
                        
                    End If
                ''Debug.Print mId$(xRTF, 1, 40); xLine
                End If
            End If
        End If
    End If
Loop
Close #fic

fctSelBold_Ok
xRTF = xRTF & txtModèle_RTF.Text & Mid$(xLine, 1, lenLine)

Call lstErr_Clear(frmElpKM.lstErr, frmElpKM.cmdContext, "cmdSPLF_Click fin : " & seq)

txtModèle_RTF.Text = xRTF

For K = 1 To arrUnderline_Nb
    txtModèle_RTF.SelStart = arrUnderline_SelStart(K)
    txtModèle_RTF.SelLength = arrUnderline_SelLength(K)
    txtModèle_RTF.SelUnderline = True
Next K

For K = 1 To arrBold_Nb
    txtModèle_RTF.SelStart = arrBold_SelStart(K)
    txtModèle_RTF.SelLength = arrBold_SelLength(K)
    txtModèle_RTF.SelBold = True
Next K

''xRtvEdition_Nb = 0
recEdition_Init xRtfEdition

frmRTF_FileName = lFileName
K = InStr(1, lFileName, "\Splf")
If K < 1 Then
    K = 1
Else
    K = K + 5
End If

frmRTF_Form_K2 = Trim(xEdition_Form.K2)
frmRTF_Référence = Trim(xEdition_Form.Name)
frmRTF_UsrId_Origine = Mid$(lFileName, K, Len(lFileName) - K - 1)

frmRTF_Caller = "frmEdition  Display"
xRtfEdition.Memo2 = txtModèle_RTF.TextRTF
frmRTF_recEdition = xRtfEdition
frmRTF_Buffer_Name = ""
frmRTF_blnOK = False
''DoEvents
frmRTF.cboRTF_Police = txtModèle_RTF.Font.Name
frmRTF.txtRTF_Size = txtModèle_RTF.Font.Size
frmRTF.txtRTF.TextRTF = frmRTF_recEdition.Memo2

Error_Handle:
Close

End Sub

Public Sub Auto_Edition(lMsg As String)
Dim K As Integer, iCopies As Integer
Dim X As String, xFileName As String
Dim blnHold As Boolean
Dim wPut_Folder As String
Dim wPrinter_Name As String

On Error Resume Next

Call FEU_ROUGE

txtSPLF_PageStart = 1: txtSPLF_PageEnd = 100000
mPageStart = 1: mPageEnd = 10000

lstCourrier_Load_filDoc "Print", ""
'jpl MsgBox "frmEDITION Auto_EDition TEST"
'jpl  lstCourrier_Load_filDoc "Print_Test_Auto", ""

'________________________________________________________________________
'$JPL 2014-11-12 états cautions DAFI ou SOBI dans frmSPLFJOB
'________________________________________________________________________

For K = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = K
    mnuSplf_Read filDoc.PATH & "\" & filDoc.FileName
    wPrinter_Name = Auto_Edition_Printer(filDoc.FileName)
    If wPrinter_Name <> "" Then
        If xUser.ProdTest = "T" Then
            blnFiligrane = False 'True
            prtFiligrane_Name = paramEditionFiligrane_Folder & "Test\" & ".jpg" ' ".bmp"
        Else
            blnFiligrane = False
            prtFiligrane_Name = paramEditionFiligrane_Folder & "Production\" & ".jpg" ' ".bmp"
        End If
        ''''''''''''''''''''wPut_Folder = paramEditionCorbeille_Folder
        'DR 21/04/2021
        If AImprimer(filDoc.PATH & "\" & filDoc.FileName) Then frmRTF_Caller = "frmEdition  Print  "
        frmRTF.Msg_Rcv ("frmEdition     ")
        If xEdition_Form.Copies > 1 Then
            If Trim(xEdition_Form.K2) = "AVI002P1" Then
                blnFiligrane = True
                prtFiligrane_Color = vbMagenta
                prtFiligrane_Name = "exemplaire Banque"
            End If
           For iCopies = 2 To xEdition_Form.Copies
               frmRTF_Caller = "frmEdition  Print  "
                frmRTF.Msg_Rcv ("frmEdition     ")
            Next iCopies
        End If
        If xUser.QSYSOPR = "1" And Trim(xEdition_Form.K2) = "CPT096P1" Then
            Call Printer_Set_Unit("BOTC")
            frmRTF_Caller = "frmEdition  Print  "
            frmRTF.Msg_Rcv ("frmEdition     ")
             Call Printer_Set_Unit("FOTC")
            frmRTF_Caller = "frmEdition  Print  "
            frmRTF.Msg_Rcv ("frmEdition     ")
       End If
    End If
    Select Case xUser.ProdTest
        Case "P": wPut_Folder = paramEditionSplf_Folder & "Production\"
        Case "T": wPut_Folder = paramEditionSplf_Folder & "Test\"
        Case "I": wPut_Folder = paramEditionSplf_Folder & "System\"
        Case Else: wPut_Folder = paramEditionSplf_Folder & "Production\" '''"System\"
    End Select
    xFileName = wPut_Folder & "\" & filDoc.FileName
    If Trim(Dir(xFileName)) <> "" Then Kill xFileName
    msFileSystem.MoveFile filDoc.PATH & "\" & filDoc.FileName, xFileName
Next K
Auto_SendMail
Auto_NoPaper  '$JPL 2014-10-01
Call FEU_VERT
End Sub
Public Sub Auto_SendMail()
Dim K As Integer, iCopies As Integer
Dim X As String, xFileName As String
Dim blnSendMail As Boolean
Dim wPut_Folder As String
Dim wFileName As String
Dim kSearch As Long
On Error Resume Next


txtSPLF_PageStart = 1: txtSPLF_PageEnd = 100000
mPageStart = 1: mPageEnd = 10000

lstCourrier_Load_filDoc "SendMail", ""

For K = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = K
    
    xFileName = filDoc.PATH & "\" & filDoc.FileName
    mnuSplf_Read xFileName
    frmRTF_Caller = "frmEdition  SAVE   "
    frmRTF.Msg_Rcv ("frmEdition     ")
    blnSendMail = True
    If InStr(filDoc.FileName, "FCIGS018P1") > 0 And xUser.ProdTest = "P" Then
        kSearch = InStr(frmRTF.txtRTF.Text, "PAS DE COMPTES DECLARES")
        If kSearch > 0 Then blnSendMail = False
    End If
    If InStr(filDoc.FileName, "FCIGS018P3") > 0 And xUser.ProdTest = "P" Then
        kSearch = InStr(frmRTF.txtRTF.Text, "PAS D'ANOMALIES DETECTEES")
        If kSearch > 0 Then blnSendMail = False
    End If
   
    If blnSendMail Then
        wFileName = "C:\Temp\" & filDoc.FileName & ".rtf"
        If Dir(wFileName) <> "" Then Kill wFileName
        frmRTF.txtRTF.SaveFile wFileName
        
        If InStr(filDoc.FileName, "SCHGE005P1") > 0 Then
            X = "<body bgcolor=" & Asc34 & "MAGENTA" & Asc34 & ">" _
                        & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
                        & htmlFontColor("BLUE") & "<BR><BR>" & "Détection d'un document 'SCHGE005P1' d'anomalie de comptabilisation" _
                        & "<BR><BR>" & "Vérifier le contenu de la pièce jointe : " & wFileName
    
            Call Email_Alerte("ALERTE", "CPT", "Document : " & filDoc.FileName, X, True, wFileName)
            
            '$JPL 2013-06-11
            msFileSystem.CopyFile xFileName, paramZSCHCRO0_SPLF & filDoc.FileName

        End If

        If InStr(filDoc.FileName, "FCIGS018P1") > 0 Then
            X = "<body bgcolor=" & Asc34 & "YELLOW" & Asc34 & ">" _
                        & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
                        & htmlFontColor("BLUE") & "<BR><BR>" & "Détection d'un document 'FCIGS018P1' de déclaration de chèques impayés" _
                        & "<BR><BR>" & "Veuillez consulter le contenu de la pièce jointe : " & wFileName
    
            Call Email_Alerte("FCI=>BDF", "FCI", "Document : " & filDoc.FileName, X, True, wFileName)
        End If
 
        If InStr(filDoc.FileName, "FCIGS018P3") > 0 Then
            X = "<body bgcolor=" & Asc34 & "MAGENTA" & Asc34 & ">" _
                        & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
                        & htmlFontColor("BLUE") & "<BR><BR>" & "Détection d'un document 'FCIGS018P3' Anomalies de la déclaration des comptes (chèques impayés)" _
                        & "<BR><BR>" & "Veuillez consulter le contenu de la pièce jointe : " & wFileName
    
            Call Email_Alerte("FCI_ALERTE", "FCI", "Document : " & filDoc.FileName, X, True, wFileName)
        End If
       
        If Dir(wFileName) <> "" Then Kill wFileName
   End If
     If Dir(xFileName) <> "" Then Kill xFileName
    
        
Next K



End Sub


Public Function Auto_Edition_Printer(lFileName As String)
Dim wPrinter_Name As String
Dim blnOk As Boolean
Dim K As Integer
Dim X As String

On Error GoTo Error_Handler

Auto_Edition_Printer = ""
wPrinter_Name = ""
If xUser.QSYSOPR = "1" Then
    xUnit.Id = Trim(xEdition_Form.Unit)
' pb AVIS CHGVE053P1 à imprimer  : SOBF / ORPA
' frmSPLFJOB affecte le service dans le nom de fichier en fonction du code nature : num => SOBF sinon ORPA
' il faut diriger sur l'imprimante du service
    If Mid$(lFileName, 1, 1) = "_" Then
        K = InStr(2, lFileName, ".")
        If K > 0 Then xUnit.Id = Mid$(lFileName, 2, K - 2)
    End If
    Call Table_Unit(xUnit)
    wPrinter_Name = Trim(xUnit.Printer)
Else
    'JPL 2014-09-22 NoPaper__________________________________________________________
    'If xEdition_Form.PrinterUnit <> "1" Then wPrinter_Name = Trim(xUser.Printer)
    wPrinter_Name = Trim(xUser.Printer)
    'JPL 2014-09-22 NoPaper__________________________________________________________
End If
If wPrinter_Name = "" Then
    xUnit.Id = xUser.Unit
    Call Table_Unit(xUnit)
    wPrinter_Name = Trim(xUnit.Printer)
End If
Auto_Edition_Printer = Printer_Set(wPrinter_Name)
Exit Function

Error_Handler:
'$JPL-20040526 Shell_MsgBox wPrinter_Name & ": " & Error, vbInformation, "# Auto_Edition_Printer # ", True

If blnOff_Line Then Auto_Edition_Printer = "jpl"

End Function

Public Sub mnuSPLF_CACLS_Exe()
Dim X As String
Dim K As Integer, K2 As Integer
Dim wUser_CACLS As typeUser

For K = 0 To filDoc.ListCount - 1
    DoEvents
    filDoc.ListIndex = K
    X = filDoc.FileName
    K2 = InStr(1, X, ".")
    
    If K2 > 0 Then
        wUser_CACLS.Id = Mid$(X, 1, K2 - 1)
        Call Table_User_CACLS(wUser_CACLS)
        Call File_CACLS(filDoc.PATH & "\" & filDoc.FileName, wUser_CACLS.Id, wUser_CACLS.Unit)    'gestion des droits  ACL
        DoEvents
    Else
        MsgBox "manque <user.>" & X, vbInformation, "affectation des droits CACLS"
    End If
Next K


End Sub

Public Sub Auto_NoPaper_Recap(lFolder As String, lUnit As String)

Dim K As Integer, X As String, K1 As Integer, K2 As Integer, K3 As Integer, wAmj As String
Dim xSubject As String, xFileName As String, mUnit As String, mUnit_Nom As String
Dim msFile As Scripting.File
Dim msFile_doc As Scripting.TextStream
Dim xText As String, xLnk As String, xHTML_Space As String
Dim blnFile_Ok As Boolean

Dim wSendMail As typeSendMail
Dim xDétail_D As String, xHeader_D As String
'=============================
On Error GoTo Err_msFIle
'=============================
xHTML_Space = String(30, Chr(160))
lstW.Clear

If lUnit = "*" Then
    For K = 1 To fgNoPaper.Rows - 1
        fgNoPaper.Row = K
        fgNoPaper.Col = fgNoPaper_arrIndex: X = fgNoPaper.Text
        lstW.AddItem "S00.*:" & Replace(X, mNoPaper_Folder, "") ' mNoPaper_Folder & msFile.Name
    Next K
Else
    filW.PATH = lFolder
    filW.Pattern = "xx*.xxx"
    filW.Pattern = "*.*" 'pdf"
    For K = 0 To filW.ListCount - 1
        filW.ListIndex = K
        K1 = InStr(filW.FileName, " (")
        X = "": mUnit = ""
        If K1 > 0 Then
            mUnit = Mid$(filW.FileName, K1 + 2, 3)
            
            K1 = InStr(filW.FileName, ".")
            K2 = InStr(K1 + 15, filW.FileName, "_")
            If K2 > 0 Then
                K3 = InStr(K2 + 1, filW.FileName, "_")
                If K3 > 0 Then X = mUnit & "." & Mid$(filW.FileName, K2 + 1, K3 - K2 - 1) & Mid$(filW.FileName, K1, K2 - K1)
            End If
        End If
        
        If lUnit = "" Then
            If mUnit <> "S99" Then lstW.AddItem X & ":" & filW.FileName
        Else
            If lUnit = mUnit Then lstW.AddItem X & ":" & filW.FileName
        End If
    Next K
End If

mUnit = ""

K = InStr(lFolder, "\NoPaper\")
If K > 0 Then
    xSubject = Mid$(lFolder, K, Len(lFolder) - K + 1) & " "
Else
    xSubject = "NoPaper\"
End If

xHeader_D = "<TR>" _
         & "<TH bgcolor = #0090A0 width=10% height=7><span style='font-size:10.0pt;font-family:Calibri'><Font color=#FFFFFF>" _
         & "Code" & "</TH>" _
         & "<TH bgcolor=#0090A0 width=90% height=7><span style='font-size:10.0pt;font-family:Calibri'><Font color=#FFFFFF>" _
         & "Document" & "</TH></TR>"

xDétail_D = ""

For K = 0 To lstW.ListCount - 1
    lstW.ListIndex = K
'__________________________________________________________________________________________
    K1 = InStr(lstW.Text, ".")
    If K1 > 0 Then
        If mUnit <> Mid$(lstW, 1, K1 - 1) Then
            If xDétail_D <> "" Then
                wSendMail.Recipient = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire(mUnit)
                wSendMail.From = currentSSIWINMAIL
                wSendMail.FromDisplayName = "NoPaper Archive"
                wSendMail.CcRecipient = ""
                wSendMail.AsHTML = True
                wSendMail.Subject = xSubject & mUnit & mUnit_Nom
                wSendMail.Attachment = ""
                wSendMail.Message = "<bgcolor = #00B0C0>" _
                                    & "<span style='font-size:12.0pt'>" _
                                    & "<TABLE   width=100% border=1 cellpadding=4 ></B>" _
                                    & xHeader_D _
                                    & xDétail_D _
                                    & "</TABLE>"
                
                srvSendMail.Monitor wSendMail
            End If
            '============================
            xDétail_D = ""
            mUnit = Mid$(lstW, 1, K1 - 1)
            X = "select SSIUSRUIDX from " & paramIBM_Library_SABSPE & ".YSSIUSR0" _
                & " where ssiusrnat = 'S' and ssiusrunit = '" & mUnit & "'"
            Set rsSab = cnsab.Execute(X)
                   
            If Not rsSab.EOF Then
                mUnit_Nom = " : " & Trim(rsSab("SSIUSRUIDX"))
            Else
                mUnit_Nom = ""
            End If
            
            '============================
        End If
    End If
'__________________________________________________________________________________________
    K1 = InStr(lstW.Text, ":")
    If K1 > 0 Then xFileName = Mid$(lstW, K1 + 1, Len(lstW.Text) - K1)
    Call Auto_NoPaper_lnk(lFolder, xFileName, xText, xLnk, wAmj)
    
    xLnk = Replace(xLnk, paramEditionNoPaper_Folder & "PDF\", paramEditionNoPaper_Partage)
    xDétail_D = xDétail_D _
         & "<TD COLSPAN=2 bgcolor = #C0FFC0  height=7><span style='font-size:11.0pt;font-family:Calibri'>" _
         & xLnk & "</TD></TR>"
 
    xText = ""
    If InStr(xFileName, ".pdf") > 0 Then
        X = Replace(lFolder, "\PDF\", "\DOC\") & "\" & Replace(xFileName, ".pdf", ".doc")
        
        If Dir(X) <> "" Then
            blnFile_Ok = True
            '=======================
            Set msFile = msFileSystem.GetFile(X)
            Set msFile_doc = msFile.OpenAsTextStream(ForReading)
            
            If blnFile_Ok Then

                X = msFile_doc.ReadLine
                X = msFile_doc.ReadLine:
                K1 = InStr(X, "\pard")
                K1 = InStr(K1, X, " ")
                If K1 > 0 Then xText = xText & Mid$(X, K1, Len(X) - K1) & "<BR>"
                For K1 = 1 To 15
                
                    X = msFile_doc.ReadLine
                    If Len(X) < 2 Then Exit For
                    xText = xText & Replace(X, "\par ", "") & "<BR>"
                Next K1
                
                xText = Replace(xText, "\b0 ", "<\b>")
                xText = Replace(xText, "\b ", "<b>")
                xText = Replace(xText, "\'e0", "à")
                xText = Replace(xText, "\'e9", "é")
                xText = Replace(xText, "\'e8", "è")
                xText = Replace(xText, "\'ea", "ê")
                xText = Replace(xText, "\'e7", "ç")
                xText = Replace(xText, "\'b0", Chr$(186))
            End If
            msFile_doc.Close
        End If
    End If
    xDétail_D = xDétail_D _
     & "<TD bgcolor = #FFFFFF width=20% height=7><span style='font-size:8.0pt;font-family:Courier New'>" & xHTML_Space _
     & "<TD bgcolor = #FFFFFF width=80% height=7><span style='font-size:8.0pt;font-family:Courier New'>" _
     & Replace(xText, " ", Chr(160)) & "<BR></TD></TR>"

Next K

'-----------------------------------------------------------------------------------

If xDétail_D <> "" Then
    If lUnit = "*" Then
        wSendMail.Recipient = currentSSIWINMAIL
        wSendMail.Subject = xSubject & lstNoPaper.Text
    ElseIf lUnit = "" Then
        wSendMail.Recipient = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire(mUnit)
        wSendMail.Subject = xSubject & lstNoPaper.Text
    Else
        wSendMail.Recipient = currentSSIWINMAIL
        wSendMail.Subject = xSubject & mUnit & mUnit_Nom
   End If
    wSendMail.From = currentSSIWINMAIL
    wSendMail.FromDisplayName = "NoPaper Archive"
    wSendMail.CcRecipient = ""
    wSendMail.AsHTML = True
    wSendMail.Attachment = ""
    wSendMail.Message = "<bgcolor = #00B0C0>" _
                        & "<span style='font-size:12.0pt'>" _
                        & "<TABLE   width=100% border=1 cellpadding=4 ></B>" _
                        & xHeader_D _
                        & xDétail_D _
                        & "</TABLE>"
    
    srvSendMail.Monitor wSendMail
End If

Exit Sub

Err_msFIle:
    blnFile_Ok = False
    Resume Next

End Sub

Public Sub Auto_NoPaper_lnk(lFolder As String, lFile As String, lDisplay As String, lLnk As String, lAMJ As String)
Dim K1 As Integer, K2 As Integer, lenX As Integer
Dim wAMJ_HMS As String, wJob As String, wEdition_Form_Id As String
Dim xLib As String
Dim xDestinataire As String ', xFile_PDF As String
    
'xFile_PDF = Replace(lFile, ".txt", ".pdf")

K1 = InStr(1, lFile, ".")
If K1 > 0 Then xDestinataire = Mid$(lFile, 1, K1 - 1)
    
K1 = K1 + 1
K2 = InStr(K1, lFile, "_")
If K2 > 0 Then
    lAMJ = Mid$(lFile, K1, K2 - K1)
    K2 = InStr(K2 + 1, lFile, "_")
    If K2 > 0 Then
        wAMJ_HMS = Mid$(lFile, K1, K2 - K1)
        K1 = K2 + 1
        K2 = InStr(K1, lFile, "_")
        If K2 > 0 Then
            wEdition_Form_Id = Mid$(lFile, K1, K2 - K1)
            K1 = K2 + 1
            lenX = Len(lFile)
            If K2 < lenX Then wJob = Mid$(lFile, K1, lenX - K1 + 1)
            
            xEdition_Form.K1 = "SAB"
            xEdition_Form.K2 = wEdition_Form_Id
        End If
    End If
End If
xLib = Trim(rsEdition_Form(xEdition_Form))
If xLib = "" Then xLib = LCase(wEdition_Form_Id)
lDisplay = wEdition_Form_Id & " --  " & xLib _
                       & "  -- (" & LCase$(xDestinataire) _
                        & "   " & Format(wAMJ_HMS, "@@@@-@@-@@  @@@:@@:@@") _
                        & "  " & wJob & ")"

lLnk = "<B>" & htmlFontColor_Blue & wEdition_Form_Id & "</B>" _
     & htmlFontColor_Magenta & " --  " & xLib & " --  " & htmlFontColor_Blue _
      & "<A HREF=" & Asc34 & "\\" & lFolder & "\" & lFile & Asc34 & ">" _
       & "(" & xDestinataire _
     & " - " & Format(wAMJ_HMS, "@@@@-@@-@@ @@@:@@:@@") & " / " & wJob _
     & ")</A>"
'____________________________________________________________________________________________
' document généré automatiquement à DSys concernant la journée comptable précédente
If Mid$(xDestinataire, 1, 1) = "_" Then
    If wAMJ_HMS < DSys & "_120000" And wAMJ_HMS > YBIATAB0_DATE_CPT_J & "_000000" Then lAMJ = YBIATAB0_DATE_CPT_J
End If

'____________________________________________________________________________________________
End Sub


Public Sub lstNoPaper_Folders_Load()
Dim wAMJMin As String, WAMJMax As String
On Error Resume Next

Dim K As Integer

lstNoPaper.Clear

fgNoPaper.Clear
fgNoPaper.Visible = False: DoEvents

If Not IsNull(txtNoPaper_AMJ_Min.Value) Then
    txtNoPaper_AMJ_Max.Visible = True
    Call DTPicker_Control(txtNoPaper_AMJ_Min, wAMJMin)
    Call DTPicker_Control(txtNoPaper_AMJ_Max, WAMJMax)
    If WAMJMax < wAMJMin Then WAMJMax = wAMJMin
    wAMJMin = mNoPaper_Opt & wAMJMin
    WAMJMax = mNoPaper_Opt & WAMJMax
Else
    txtNoPaper_AMJ_Max.Visible = False
End If


Set msFolder = msFileSystem.GetFolder(paramEditionNoPaper_Partage)
For Each msSubFolder In msFolder.SubFolders
    If InStr(msSubFolder.Name, mNoPaper_Opt) > 0 Then
        If Not txtNoPaper_AMJ_Max.Visible Then
            lstNoPaper.AddItem msSubFolder.Name
        Else
            If msSubFolder.Name >= wAMJMin And msSubFolder.Name <= WAMJMax Then lstNoPaper.AddItem msSubFolder.Name
        End If
        
    End If
    
Next

If lstNoPaper.ListCount = 1 Then lstNoPaper.ListIndex = 0

End Sub


Public Function Auto_NoPaper_PDFCreator(lFileName As String, lFolder_In As String, lFolder_Out As String) As Boolean

On Error GoTo Error_Handle

    Auto_NoPaper_PDFCreator = False
    If nomDuServeur <> paramServerSplf Then
        oPDF.cOption("UseAutosave") = 1
        oPDF.cOption("UseAutosaveDirectory") = 1
        oPDF.cOption("AutosaveDirectory") = lFolder_Out
        oPDF.cOption("AutosaveFilename") = lFileName & ".pdf"
        oPDF.cOption("AutosaveFormat") = 0
        oPDF.cDefaultPrinter = "oPDF"
        oPDF.cClearCache
        oPDF.cPrintFile (lFolder_In & "\" & lFileName)   ''''(pathPdf & FileName & ".rtf")
        oPDF.cOption("AutosaveStartStandardProgram") = 0 '1
        Auto_NoPaper_PDFCreator = True
    End If
    Exit Function
    
Error_Handle:
    Auto_NoPaper_PDFCreator = False
End Function

Public Function Auto_NoPaper_MakePDF(lFileName As String, lFolder_In As String, lFolder_Out As String) As Boolean
Dim xFile_PDF As String, xFile_Out As String, K1 As Long
On Error GoTo Error_Handle

Auto_NoPaper_MakePDF = False

If InStr(UCase(lFileName), "ECHEDI01P2") > 0 Then
    Auto_NoPaper_MakePDF = True
    mMakePDF_Error = 0
    GoTo Exit_sub
End If

If Not blnMakePDF_Actif Then GoTo Error_MakePDF

xFile_Out = lFolder_Out & Replace(lFileName, ".doc", ".pdf")

If Not msFileSystem.FolderExists(lFolder_Out) Then MkDir lFolder_Out

xFile_PDF = paramEditionNoPaper_Folder_MakePDF & Replace(lFileName, ".doc", ".pdf")
If Dir(xFile_PDF) <> "" Then Kill xFile_PDF
msFileSystem.CopyFile lFolder_In & lFileName, paramEditionNoPaper_Folder_MakePDF & lFileName

For K1 = 1 To 1800 '3600 '400  2015-01-10 JPL
   DoEvents
    If Dir(paramEditionNoPaper_Folder_MakePDF & lFileName) = "" Then
        If Dir(xFile_PDF) <> "" Then Exit For
    End If
    Sleep 200 '500
 Next K1
 
If K1 < 400 Then
    DoEvents: Sleep 200 '500
    If Dir(xFile_Out) <> "" Then Kill xFile_Out
    'DR 04/10/2019
    'msFileSystem.MoveFile xFile_PDF, xFile_Out
    Call FileCopy(xFile_PDF, xFile_Out)
    Kill xFile_PDF 'ajout si copy et non move
    '                                       '
    Auto_NoPaper_MakePDF = True
    mMakePDF_Error = 0
Else
    mMakePDF_Error = mMakePDF_Error + 1
End If
GoTo Exit_sub

Error_Handle:
If Not blnTimer_Enabled Then MsgBox paramEditionNoPaper_Folder_MakePDF & lFileName & Error, vbCritical, "frmEdition > Auto_NoPaper_MakePDF"

On Error Resume Next
'___________________________________________________________________
If Dir(xFile_Out) <> "" Then Kill xFile_Out
'msFileSystem.MoveFile xFile_PDF, xFile_Out
Call FileCopy(xFile_PDF, xFile_Out)
Kill xFile_PDF 'ajout si copy et non move
'                                       '
Error_MakePDF:
blnMakePDF_Actif = True

Exit_sub:

End Function
Public Sub ecrit_LogPrintings(nFileName As String, pName As String)
Dim FicSortie As Long
Dim ficName As String
Dim newDirectory As String
Dim xMemo As String
Dim z As String

On Error Resume Next
    xMemo = paramServer("\\LOGPRINTINGS\")
    If IsNull(xMemo) Then
        Exit Sub
    End If
    If Dir(Trim(xMemo), vbDirectory) = "" Then
        MkDir Trim(xMemo)
    End If
    ficName = Trim(xMemo) & "BIA_SAB_Printing_"
    ficName = ficName & CStr(Year(Now)) & Mid(CStr(100 + Month(Now)), 2) & Mid(CStr(100 + Day(Now)), 2) & ".log"
    FicSortie = FreeFile
    If Dir(ficName, vbNormal) = "" Then
        Open ficName For Output As #FicSortie
        Close #FicSortie
    End If
    Open ficName For Append As #FicSortie
    z = "Heure = " & Format(Time, "hh:nn:ss") & " Fichier = " & nFileName
    z = z & " Imprimante = " & pName
    Print #FicSortie, z
    Close #FicSortie

End Sub

Public Sub cboNoPaper_User_Load(lUnit As String)
Dim xSQL As String, K As Integer, X As String
cboNoPaper_User.Clear
cboNoPaper_User.AddItem ""
If lUnit <> "" Then
    xSQL = "select SSIDOMUIDX  from " & paramIBM_Library_SABSPE & ".YSSIUSR0 ," & paramIBM_Library_SABSPE & ".YSSIDOM0" _
         & " where SSIUSRNAT = ' ' and SSIUSRSTAK = ' ' and SSIUSRUNIT = '" & lUnit & "'" _
         & " and SSIDOMNAT = ' ' and SSIDOMUIDN = SSIUSRUIDN and SSIDOMDIDX = 'IBM' and SSIDOMSTAK = ' ' order by SSIDOMUIDX "
Else
    xSQL = "select SSIDOMUIDX  from " & paramIBM_Library_SABSPE & ".YSSIUSR0 ," & paramIBM_Library_SABSPE & ".YSSIDOM0" _
         & " where SSIUSRNAT = ' ' and SSIUSRSTAK = ' '" _
         & " and SSIDOMNAT = ' ' and SSIDOMUIDN = SSIUSRUIDN and SSIDOMDIDX = 'IBM' and SSIDOMSTAK = ' ' order by SSIDOMUIDX "
End If
Set rsSab = cnsab.Execute(xSQL)
Do Until rsSab.EOF
    X = Trim(rsSab("SSIDOMUIDX"))
    cboNoPaper_User.AddItem X
    If currentSSIWINUIDX_U = X Then K = cboNoPaper_User.ListCount - 1
    rsSab.MoveNext
Loop

If lUnit <> "" Then cboNoPaper_User.ListIndex = K

End Sub
