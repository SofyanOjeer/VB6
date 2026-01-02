VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmYCLISCO0 
   AutoRedraw      =   -1  'True
   Caption         =   "Scoring client-prospect"
   ClientHeight    =   12165
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   16335
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "YCLISCO0.frx":0000
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
      Left            =   8685
      TabIndex        =   2
      Top             =   15
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   11640
      Left            =   45
      TabIndex        =   3
      Top             =   375
      Width           =   16290
      _ExtentX        =   28734
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
      TabPicture(0)   =   "YCLISCO0.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "YCLISCO0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "YCLISCO0.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtRTF"
      Tab(2).Control(1)=   "lstPays"
      Tab(2).Control(2)=   "fraPays"
      Tab(2).Control(3)=   "fraRTF"
      Tab(2).ControlCount=   4
      Begin VB.Frame fraRTF 
         BackColor       =   &H80000005&
         Height          =   6990
         Left            =   -70515
         TabIndex        =   69
         Top             =   1110
         Width           =   4230
         Begin VB.FileListBox filRTF 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3510
            Left            =   375
            TabIndex        =   72
            Top             =   1080
            Width           =   3510
         End
         Begin VB.CommandButton cmdRTF_Update 
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
            Height          =   600
            Left            =   2550
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   5985
            Width           =   1200
         End
         Begin VB.CommandButton cmdRTF_Quit 
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
            Height          =   600
            Left            =   315
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   5910
            Width           =   1200
         End
      End
      Begin VB.Frame fraPays 
         BackColor       =   &H00F0F0F0&
         Height          =   9000
         Left            =   -75195
         TabIndex        =   59
         Top             =   975
         Width           =   4575
         Begin VB.CheckBox chkBIA_2 
            BackColor       =   &H0080C0FF&
            Caption         =   "liste BIA - 2"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   285
            TabIndex        =   80
            Top             =   3990
            Width           =   4020
         End
         Begin VB.CheckBox chkBIA_1 
            BackColor       =   &H0000FFFF&
            Caption         =   "liste BIA - 1"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   285
            TabIndex        =   79
            Top             =   3400
            Width           =   4020
         End
         Begin VB.CheckBox chkBIA_CA 
            BackColor       =   &H000000FF&
            Caption         =   "liste BIA - CA"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   285
            TabIndex        =   73
            Top             =   2800
            Width           =   4020
         End
         Begin VB.CommandButton cmdPays_Update 
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
            Height          =   600
            Left            =   2775
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   8025
            Width           =   1200
         End
         Begin VB.CommandButton cmdPays_Quit 
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
            Height          =   600
            Left            =   540
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   8010
            Width           =   1200
         End
         Begin VB.TextBox txtScore 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3090
            MaxLength       =   3
            TabIndex        =   66
            Top             =   6465
            Width           =   1050
         End
         Begin VB.CheckBox chkCRS 
            BackColor       =   &H00C0FFC0&
            Caption         =   "CRS : pays signataire de l'accord d'échange automatique d'information"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Left            =   285
            TabIndex        =   63
            Top             =   4590
            Width           =   4020
         End
         Begin VB.CheckBox chkEmbargo 
            BackColor       =   &H008080FF&
            Caption         =   "pays sous embargo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   285
            TabIndex        =   62
            Top             =   2200
            Width           =   4020
         End
         Begin VB.CheckBox chkGAFI_G 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Liste grise GAFI"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   285
            TabIndex        =   61
            Top             =   1600
            Width           =   4020
         End
         Begin VB.CheckBox chkGAFI_N 
            BackColor       =   &H00000000&
            Caption         =   "Liste noire GAFI"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   285
            TabIndex        =   60
            Top             =   1000
            Width           =   4020
         End
         Begin VB.Label lblScore 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  Score"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   285
            TabIndex        =   65
            Top             =   6435
            Width           =   3930
         End
         Begin VB.Label libPays 
            BackColor       =   &H00C0E0FF&
            Caption         =   "pays"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   285
            TabIndex        =   64
            Top             =   210
            Width           =   4020
         End
      End
      Begin VB.ListBox lstPays 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9510
         Left            =   -65715
         TabIndex        =   57
         Top             =   1485
         Width           =   6675
      End
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
         Height          =   11055
         Left            =   45
         TabIndex        =   4
         Top             =   525
         Width           =   16155
         Begin VB.CommandButton cmdDetail_Init 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Saisir un dossier"
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
            Left            =   11655
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   615
            Width           =   1335
         End
         Begin VB.CommandButton cmdDetail_Quit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Abandonner (= Echap )"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   7680
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   10065
            Width           =   1725
         End
         Begin VB.Frame fraDetail 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Id"
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
            Left            =   7050
            TabIndex        =   14
            Top             =   1275
            Visible         =   0   'False
            Width           =   8865
            Begin VB.TextBox txtInfo 
               BackColor       =   &H00C0FFFF&
               Height          =   1230
               Left            =   225
               MultiLine       =   -1  'True
               TabIndex        =   58
               Top             =   7455
               Width           =   8385
            End
            Begin VB.CommandButton cmdCRS_Non 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Non"
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
               Left            =   5600
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   6200
               Width           =   1000
            End
            Begin VB.CommandButton cmdFAT_Non 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Non"
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
               Left            =   5600
               Style           =   1  'Graphical
               TabIndex        =   52
               Top             =   5600
               Width           =   1000
            End
            Begin VB.CommandButton cmdCOB_Non 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Non"
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
               Left            =   5600
               Style           =   1  'Graphical
               TabIndex        =   51
               Top             =   5000
               Width           =   1000
            End
            Begin VB.CommandButton cmdACT_Non 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Non"
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
               Left            =   5600
               Style           =   1  'Graphical
               TabIndex        =   50
               Top             =   4300
               Width           =   1000
            End
            Begin VB.CommandButton cmdCRS_Oui 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Oui"
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
               Left            =   4600
               Style           =   1  'Graphical
               TabIndex        =   49
               Top             =   6200
               Width           =   1000
            End
            Begin VB.CommandButton cmdFAT_Oui 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Oui"
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
               Left            =   4600
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   5600
               Width           =   1000
            End
            Begin VB.CommandButton cmdCOB_Oui 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Oui"
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
               Left            =   4600
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   5000
               Width           =   1000
            End
            Begin VB.CommandButton cmdACT_Oui 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Oui"
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
               Left            =   4590
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   4300
               Width           =   1000
            End
            Begin VB.TextBox txtCRSV 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   8000
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   45
               Top             =   6200
               Width           =   600
            End
            Begin VB.TextBox txtFATV 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   8000
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   44
               Top             =   5600
               Width           =   600
            End
            Begin VB.TextBox txtCOBV 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   8000
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   43
               Top             =   5000
               Width           =   600
            End
            Begin VB.TextBox txtACTV 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   8000
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   42
               Top             =   4300
               Width           =   600
            End
            Begin VB.TextBox txtNATV 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   8000
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   41
               Top             =   3500
               Width           =   600
            End
            Begin VB.TextBox txtRESV 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   8000
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   40
               Top             =   2610
               Width           =   600
            End
            Begin VB.TextBox txtDEBV 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   8000
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   33
               Top             =   2000
               Width           =   600
            End
            Begin VB.CommandButton cmdDEB_Technique 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Technique"
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
               Left            =   6600
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   2000
               Width           =   1000
            End
            Begin VB.CommandButton cmdDEB_Non 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Non"
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
               Left            =   5600
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   2000
               Width           =   1000
            End
            Begin VB.CommandButton cmdDEB_Oui 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Oui"
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
               Left            =   4600
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   2000
               Width           =   1000
            End
            Begin VB.TextBox txtCLIX 
               Enabled         =   0   'False
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
               Left            =   3525
               TabIndex        =   25
               Top             =   825
               Width           =   5145
            End
            Begin VB.TextBox txtCLIR 
               Enabled         =   0   'False
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
               Left            =   2400
               MaxLength       =   12
               TabIndex        =   28
               Top             =   360
               Width           =   3180
            End
            Begin VB.TextBox txtCLID 
               Enabled         =   0   'False
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
               Left            =   2400
               MaxLength       =   7
               TabIndex        =   24
               Top             =   825
               Width           =   1110
            End
            Begin VB.TextBox txtPPEV 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   8000
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   23
               Top             =   1500
               Width           =   600
            End
            Begin VB.CommandButton cmdPPE_Non 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Non"
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
               Left            =   5600
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   1500
               Width           =   1000
            End
            Begin VB.CommandButton cmdPPE_Oui 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Oui"
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
               Left            =   4600
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   1500
               Width           =   1000
            End
            Begin VB.TextBox txtSCOV 
               Alignment       =   1  'Right Justify
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
               Left            =   7350
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   18
               Top             =   6900
               Width           =   1290
            End
            Begin VB.CommandButton cmdDetail_Delete 
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
               Height          =   600
               Left            =   7230
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   8800
               Width           =   1635
            End
            Begin VB.CommandButton cmdDetail_Add 
               BackColor       =   &H0080FF80&
               Caption         =   "Ajouter un nouveau dossier"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   2745
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   8800
               Width           =   1635
            End
            Begin VB.CommandButton cmdDetail_Update 
               BackColor       =   &H0000FFFF&
               Caption         =   "Enregistrer les modifications"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   4965
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   8800
               Width           =   1635
            End
            Begin VB.Label libNATV 
               BackColor       =   &H00E0FFFF&
               Caption         =   "libNATV"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   195
               TabIndex        =   76
               Top             =   3885
               Width           =   4395
            End
            Begin VB.Label libRESV 
               BackColor       =   &H00E0FFFF&
               Caption         =   "libRESV"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   195
               TabIndex        =   75
               Top             =   3105
               Width           =   4500
            End
            Begin VB.Label libNoRelation 
               Alignment       =   2  'Center
               BackColor       =   &H00E0FFFF&
               Caption         =   "No Relation"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   3525
               TabIndex        =   74
               Top             =   6930
               Width           =   3600
            End
            Begin VB.Label libNAT 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   4605
               TabIndex        =   56
               Top             =   3500
               Width           =   2895
            End
            Begin VB.Label libRES 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   4600
               TabIndex        =   55
               Top             =   2670
               Width           =   2880
            End
            Begin VB.Label lblCRS 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Présence d'un des critères CRS ?"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   195
               TabIndex        =   39
               Top             =   6200
               Width           =   4185
            End
            Begin VB.Label lblFAT 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Présence d'un des critères FATCA ?"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   195
               TabIndex        =   38
               Top             =   5600
               Width           =   4320
            End
            Begin VB.Label lblCOB 
               BackColor       =   &H00E0FFFF&
               Caption         =   "S'agit-il d'une relation 'correspondent banking'"
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
               Left            =   195
               TabIndex        =   37
               Top             =   5000
               Width           =   4410
            End
            Begin VB.Label lblACT 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Le domaine d'activité du prospect est-il dans la ""liste activité"" ?"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Left            =   195
               TabIndex        =   36
               Top             =   4300
               Width           =   4245
            End
            Begin VB.Label lblNAT 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Quel est le pays de nationalité du propect ?"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   195
               TabIndex        =   35
               Top             =   3450
               Width           =   4410
            End
            Begin VB.Label lblRES 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Quel est le pays de résidence du propect ?"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   195
               TabIndex        =   34
               Top             =   2730
               Width           =   4410
            End
            Begin VB.Label lblDEB 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Le client/prospect est-il présent lors de l'entrée en relation ?"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   195
               TabIndex        =   29
               Top             =   2010
               Width           =   4410
            End
            Begin VB.Label lblCLIR 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Chargé de clientèle"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   195
               TabIndex        =   27
               Top             =   400
               Width           =   1890
            End
            Begin VB.Label libSCO 
               Alignment       =   2  'Center
               BackColor       =   &H00E0FFFF&
               Caption         =   "scoring"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   195
               TabIndex        =   26
               Top             =   6900
               Width           =   2835
            End
            Begin VB.Label lblPPE 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Le client/prospect est-il un PPE ?"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   195
               TabIndex        =   20
               Top             =   1590
               Width           =   4410
            End
            Begin VB.Label lblCLID 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Racine / Intitulé"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   195
               TabIndex        =   19
               Top             =   900
               Width           =   1995
            End
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
            Left            =   14355
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   585
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
            Left            =   11610
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   180
            Width           =   4155
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
            Height          =   1140
            Left            =   105
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
            Width           =   11205
            Begin VB.TextBox txtSelect_CLISCOID 
               Height          =   285
               Left            =   1905
               TabIndex        =   13
               Top             =   360
               Width           =   1575
            End
            Begin MSComCtl2.DTPicker txtSelect_AMJ_Min 
               Height          =   300
               Left            =   9315
               TabIndex        =   10
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
               Format          =   100204547
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_AMJ_Max 
               Height          =   300
               Left            =   9555
               TabIndex        =   11
               Top             =   705
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
               Format          =   100204547
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_CLISCOID 
               BackColor       =   &H00F0FFFF&
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
               Height          =   270
               Left            =   675
               TabIndex        =   12
               Top             =   390
               Width           =   1110
            End
            Begin VB.Label lblSelect_AMJ 
               BackColor       =   &H00F0FFFF&
               Caption         =   "mise à jour le"
               Height          =   270
               Left            =   7650
               TabIndex        =   9
               Top             =   315
               Width           =   1335
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   9750
            Left            =   90
            TabIndex        =   8
            Top             =   1140
            Width           =   15825
            _ExtentX        =   27914
            _ExtentY        =   17198
            _Version        =   393216
            Cols            =   14
            FixedCols       =   0
            RowHeightMin    =   400
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   8421376
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            AllowUserResizing=   3
            FormatString    =   $"YCLISCO0.frx":035E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin RichTextLib.RichTextBox txtRTF 
         Height          =   9570
         Left            =   -66180
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   810
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   16880
         _Version        =   393217
         BackColor       =   15790320
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"YCLISCO0.frx":0469
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
      Left            =   15690
      Picture         =   "YCLISCO0.frx":04E9
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   -30
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
      Begin VB.Menu mnuPrint_Excel 
         Caption         =   "Excel"
      End
      Begin VB.Menu mnuPrint_Mail 
         Caption         =   "Mail"
      End
   End
End
Attribute VB_Name = "frmYCLISCO0"
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
Dim rsSAB_X As New ADODB.Recordset
Dim mMail_Destinataires As String

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

'______________________________________________________________________

Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long

Dim xYCLISCO0 As typeYCLISCO0, oldYCLISCO0 As typeYCLISCO0, newYCLISCO0 As typeYCLISCO0, mYCLISCO0_Update As String
Dim HeightOfLine As Long, LinesOfText As Long


Dim mCLISCOZSWI As Long

Dim mCLIENACLI As String, wCLIENACLI As Long

Dim paramScoring_Path As String, txtRTF_LoadFile As String, mField As String

Dim mSCOPAY As typeSCOPAY, arrSCOPAY() As typeSCOPAY, arrSCOPAY_Nb As Integer

Dim newParam As typeYBIATAB0, oldParam As typeYBIATAB0, mYBIATAB0_Update As String

Dim blnPays_Scan As Boolean

Dim blnNoRelation_RESK As Boolean, blnNoRelation_NATK As Boolean
Public Sub Form_Init()
Dim V, xSQL As String, X As String
Dim K As Long

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True


cmdReset
blnControl = False
Call DTPicker_Set(txtSelect_AMJ_Min, YBIATAB0_DATE_CPT_JS1)
Call DTPicker_Set(txtSelect_AMJ_Max, YBIATAB0_DATE_CPT_JS1)
txtSelect_AMJ_Min.Value = Null

fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False
fraSelect_Options.Visible = True

txtRTF.Visible = False
Set txtRTF.Container = fraSelect
txtRTF.Top = fraDetail.Top
txtRTF.Left = fraSelect.Left + 100

lstPays.Visible = False
Set lstPays.Container = fraSelect
lstPays.Top = fraDetail.Top
lstPays.Left = fraSelect.Left + 100

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'SAB' and BIATABK1 = 'CLIENAPAY' order by BIATABK2"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    lstPays.AddItem Mid$(rsSab("BIATABk2"), 4, 2) & " - " & Mid$(rsSab("BIATABTXT"), 16, 30)
    rsSab.MoveNext
Loop

arrSCOPAY_Load


fraPays.Visible = False
Set fraPays.Container = fraSelect
fraPays.Top = fgSelect.Top + fgSelect.RowHeightMin
fraPays.Left = fgSelect.Left + fgSelect.Width - fraPays.Width - 100
chkGAFI_N.ForeColor = RGB(255, 255, 255)
chkEmbargo.BackColor = mColor_W1

paramScoring_Path = paramGSOP_Dossier_Path & "_Scoring\"

fraRTF.Visible = False
Set fraRTF.Container = fraSelect
fraRTF.Top = fraDetail.Top
fraRTF.Left = fgSelect.Left + fgSelect.Width - fraRTF.Width - 100
filRTF.PATH = paramScoring_Path

libNoRelation.Visible = False
libNoRelation.ForeColor = vbYellow
libNoRelation.BackColor = vbRed
libNoRelation.Caption = "NE PAS ENTRER EN RELATION"
blnControl = True
cboSelect_SQL.ListIndex = 0

'cmdSelect_Reset

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
        Case 2: fgSelect.Col = 2: X = Format$(Val(fgSelect.Text), "0000000000")

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

Dim K As Long

On Error GoTo Error_Handler
currentAction = "fgSelect_Display_1"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = fgSelect_FormatString

fgSelect.Rows = 1
                 
fgSelect.Row = 0
Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_Display_1_Line
    rsSab.MoveNext

Loop

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents
fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgSelect_Display_2()


On Error GoTo Error_Handler
currentAction = "fgSelect_Display_2"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = "<Pays       |<Libellé                                                                                |>Score         " _
                      & "|<GAFI noir|<GAFI gris|<embargo|<liste CRS|<BIA-CA|<BIA-1|<BIA-2|<Mise à jour par ..... le                         "

fgSelect.Rows = 1
                 
fgSelect.Row = 0
fgSelect.Col = 2: fgSelect.CellAlignment = 1

Do While Not rsSab.EOF

    fgSelect_Display_2_Line
    rsSab.MoveNext

Loop

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents
fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_Display_1_Line()
Dim K As Integer, wColor As Long

On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = rsSab("CLISCOID")
fgSelect.Col = 1: fgSelect.Text = Trim(rsSab("CLISCOCLID")) & " " & Trim(rsSab("CLISCOCLIX"))
fgSelect.Col = 2: fgSelect.Text = Val(rsSab("CLISCOSCOV"))

fgSelect.Col = 3
Select Case rsSab("CLISCOSCOK")
    Case "1": fgSelect.Text = "standard": wColor = mColor_G2
    Case "2": fgSelect.Text = "complémentaire": wColor = mColor_Y1
    Case "3": fgSelect.Text = "renforcée": wColor = mColor_W1
    Case Else: fgSelect.Text = "?  " & rsSab("CLISCOSCOK")
End Select

fgSelect.Col = 4: fgSelect.Text = " " & rsSab("CLISCORESK")
fgSelect.Col = 5: fgSelect.Text = " " & rsSab("CLISCONATK")

fgSelect.Col = 6: fgSelect.Text = Trim(rsSab("CLISCOYUSR")) & "  " & dateImp10_S(rsSab("CLISCOYAMJ") + 19000000) & " " & timeImp8(rsSab("CLISCOYHMS")) & "-" & rsSab("CLISCOYVER")

For K = 2 To 3
    fgSelect.Col = K
    fgSelect.CellBackColor = wColor
Next K


End Sub

Public Sub fgSelect_Display_2_Line()
Dim K As Integer, wColor As Long, X2 As String
Dim blnOk As Boolean

On Error Resume Next

X2 = Mid$(rsSab("BIATABk2"), 4, 2)
For K = 1 To arrSCOPAY_Nb
    If X2 = arrSCOPAY(K).Id Then
        mSCOPAY = arrSCOPAY(K)
        blnOk = True
        Exit For
    End If
Next K

If cmdSelect_SQL_K = "3" And Not blnOk Then
Else
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect.Col = 0: fgSelect.Text = X2
    fgSelect.Col = 1: fgSelect.Text = Mid$(rsSab("BIATABTXT"), 16, 30)
    If blnOk Then
        fgSelect.Col = 2: fgSelect.Text = mSCOPAY.V
        If mSCOPAY.GAFI_N = "Y" Then fgSelect.Col = 3: fgSelect.CellBackColor = RGB(8, 8, 8)
        If mSCOPAY.GAFI_G = "Y" Then fgSelect.Col = 4: fgSelect.CellBackColor = chkGAFI_G.BackColor
        If mSCOPAY.Embargo = "Y" Then fgSelect.Col = 5: fgSelect.CellBackColor = chkEmbargo.BackColor
        If mSCOPAY.CRS = "Y" Then fgSelect.Col = 6: fgSelect.CellBackColor = chkCRS.BackColor
        If mSCOPAY.BIA_CA = "Y" Then fgSelect.Col = 7: fgSelect.CellBackColor = chkBIA_CA.BackColor
        If mSCOPAY.BIA_1 = "Y" Then fgSelect.Col = 8: fgSelect.CellBackColor = chkBIA_1.BackColor
        If mSCOPAY.BIA_2 = "Y" Then fgSelect.Col = 9: fgSelect.CellBackColor = chkBIA_2.BackColor
        If mSCOPAY.YAMJ <> 0 Then
            fgSelect.Col = 10: fgSelect.Text = mSCOPAY.YUSR & " " & dateImp10_S(mSCOPAY.YAMJ) & " " & timeImp8(mSCOPAY.YHMS)
        End If
    End If
End If


End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim wFct As String

mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

wFct = UCase$(Trim(Mid$(Msg, 1, 12)))
Call BIA_VB_HAB(wFct, arrHab(), cboSelect_SQL)

Form_Init

mMail_Destinataires = currentSSIWINMAIL

Select Case wFct
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


Private Sub chkEmbargo_Click()
CLISCOSCOV_Pays
End Sub

Private Sub chkGAFI_G_Click()
CLISCOSCOV_Pays
End Sub

Private Sub chkGAFI_N_Click()
CLISCOSCOV_Pays
End Sub


Private Sub cmdACT_Non_Click()

cmdACT_Non.BackColor = mColor_G2
cmdACT_Oui.BackColor = RGB(230, 230, 230)

newYCLISCO0.CLISCOACTK = "N"
newYCLISCO0.CLISCOACTV = 0
txtACTV = newYCLISCO0.CLISCOACTV

CLISCOSCOV_Total


End Sub

Private Sub cmdACT_Oui_Click()
cmdACT_Oui.BackColor = mColor_W1
cmdACT_Non.BackColor = RGB(230, 230, 230)

newYCLISCO0.CLISCOACTK = "Y"
newYCLISCO0.CLISCOACTV = 50
txtACTV = newYCLISCO0.CLISCOACTV

CLISCOSCOV_Total

End Sub

Private Sub cmdCOB_Non_Click()

cmdCOB_Non.BackColor = mColor_G2
cmdCOB_Oui.BackColor = RGB(230, 230, 230)

newYCLISCO0.CLISCOCOBK = "N"
newYCLISCO0.CLISCOCOBV = 0
txtCOBV = newYCLISCO0.CLISCOCOBV

CLISCOSCOV_Total


End Sub

Private Sub cmdCOB_Oui_Click()
cmdCOB_Oui.BackColor = mColor_W1
cmdCOB_Non.BackColor = RGB(230, 230, 230)

newYCLISCO0.CLISCOCOBK = "Y"
newYCLISCO0.CLISCOCOBV = 50
txtCOBV = newYCLISCO0.CLISCOCOBV

CLISCOSCOV_Total

End Sub


Private Sub cmdCRS_Non_Click()

cmdCRS_Non.BackColor = mColor_G2
cmdCRS_Oui.BackColor = RGB(230, 230, 230)

newYCLISCO0.CLISCOCRSK = "N"
newYCLISCO0.CLISCOCRSV = 0
txtCRSV = newYCLISCO0.CLISCOCRSV

CLISCOSCOV_Total


End Sub

Private Sub cmdCRS_Oui_Click()
cmdCRS_Oui.BackColor = mColor_W1
cmdCRS_Non.BackColor = RGB(230, 230, 230)

newYCLISCO0.CLISCOCRSK = "Y"
newYCLISCO0.CLISCOCRSV = 10
txtCRSV = newYCLISCO0.CLISCOCRSV

CLISCOSCOV_Total

End Sub

Private Sub cmdDEB_Non_Click()
cmdDEB_Oui.BackColor = RGB(230, 230, 230)
cmdDEB_Non.BackColor = mColor_W1
cmdDEB_Technique.BackColor = RGB(230, 230, 230)

newYCLISCO0.CLISCODEBK = "N"
newYCLISCO0.CLISCODEBV = 50
txtDEBV = newYCLISCO0.CLISCODEBV

CLISCOSCOV_Total
End Sub

Private Sub cmdDEB_Oui_Click()
cmdDEB_Oui.BackColor = mColor_G1
cmdDEB_Non.BackColor = RGB(230, 230, 230)
cmdDEB_Technique.BackColor = RGB(230, 230, 230)

newYCLISCO0.CLISCODEBK = "Y"
newYCLISCO0.CLISCODEBV = 0
txtDEBV = newYCLISCO0.CLISCODEBV

CLISCOSCOV_Total

End Sub

Private Sub cmdDEB_Technique_Click()
cmdDEB_Oui.BackColor = RGB(230, 230, 230)
cmdDEB_Non.BackColor = RGB(230, 230, 230)
cmdDEB_Technique.BackColor = mColor_Y1

newYCLISCO0.CLISCODEBK = "T"
newYCLISCO0.CLISCODEBV = 0
txtDEBV = newYCLISCO0.CLISCODEBV

CLISCOSCOV_Total

End Sub


Private Sub cmdDEtail_Init_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Saisir un dossier ........"): DoEvents
fraDetail_Init
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdFAT_Non_Click()

'libNoRelation.Visible = False
cmdFAT_Non.BackColor = mColor_G2
cmdFAT_Oui.BackColor = RGB(230, 230, 230)

newYCLISCO0.CLISCOFATK = "N"
newYCLISCO0.CLISCOFATV = 0
txtFATV = newYCLISCO0.CLISCOFATV

CLISCOSCOV_Total


End Sub

Private Sub cmdFAT_Oui_Click()
'libNoRelation.Visible = True
cmdFAT_Oui.BackColor = mColor_W1
cmdFAT_Non.BackColor = RGB(230, 230, 230)

newYCLISCO0.CLISCOFATK = "Y"
newYCLISCO0.CLISCOFATV = 100
txtFATV = newYCLISCO0.CLISCOFATV

CLISCOSCOV_Total

End Sub

Private Sub cmdPays_Quit_Click()
fraPays.Visible = False
End Sub

Private Sub cmdPays_Update_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

newParam = oldParam
newParam.BIATABTXT = Format(txtScore, "000") & "NNNNNNN"
If chkGAFI_N = "1" Then Mid$(newParam.BIATABTXT, 4, 1) = "Y"
If chkGAFI_G = "1" Then Mid$(newParam.BIATABTXT, 5, 1) = "Y"
If chkEmbargo = "1" Then Mid$(newParam.BIATABTXT, 6, 1) = "Y"
If chkCRS = "1" Then Mid$(newParam.BIATABTXT, 7, 1) = "Y"
If chkBIA_CA = "1" Then Mid$(newParam.BIATABTXT, 8, 1) = "Y"
If chkBIA_1 = "1" Then Mid$(newParam.BIATABTXT, 9, 1) = "Y"
If chkBIA_2 = "1" Then Mid$(newParam.BIATABTXT, 10, 1) = "Y"
Mid$(newParam.BIATABTXT, 105, 10) = usrName_UCase & "          "
Mid$(newParam.BIATABTXT, 115, 8) = DSys
Mid$(newParam.BIATABTXT, 123, 6) = time_Hms

V = cmdUpdate

If IsNull(V) Then
    fraPays.Visible = False
    arrSCOPAY_Load
    cmdSelect_SQL_2
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdPPE_Non_Click()

cmdPPE_Non.BackColor = mColor_G2
cmdPPE_Oui.BackColor = RGB(230, 230, 230)

newYCLISCO0.CLISCOPPEK = "N"
newYCLISCO0.CLISCOPPEV = 0
txtPPEV = newYCLISCO0.CLISCOPPEV

CLISCOSCOV_Total

End Sub

Private Sub cmdPPE_Oui_Click()

cmdPPE_Oui.BackColor = mColor_W1
cmdPPE_Non.BackColor = RGB(230, 230, 230)

newYCLISCO0.CLISCOPPEK = "Y"
newYCLISCO0.CLISCOPPEV = 100
txtPPEV = newYCLISCO0.CLISCOPPEV

CLISCOSCOV_Total

End Sub

Private Sub cmdPrint_Click()
Dim X As String, I As Integer

If fraDetail.Visible Then
    Call MsgBox("pour imprimer cette fenêtre :" & vbCrLf _
              & "- maintenir la touche 'Alt GR' appuyée" & vbCrLf _
              & "- appuyer sur la touche 'Impr écran'" & vbCrLf _
              , vbInformation, "Scoring client / prospect")
Else

    Me.Enabled = False: Me.MousePointer = vbHourglass
    Select Case SSTab1.Tab
        Case 0:
                Me.PopupMenu mnuPrint, vbPopupMenuLeftButton
        End Select
    
    Me.Enabled = True: Me.MousePointer = 0
End If




End Sub

Private Sub cmdDetail_Add_Click()
Dim V, xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass
If IsNull(fraDetail_Control) Then

    xSQL = "select CLISCOID  from " & paramIBM_Library_SABSPE & ".YCLISCO0 " _
         & "order by CLISCOID desc FETCH FIRST 1 ROWS ONLY"
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        newYCLISCO0.CLISCOID = rsSab(0) + 1
    Else
        newYCLISCO0.CLISCOID = 1
    End If
    

    mYCLISCO0_Update = "New"
    V = cmdUpdate
    
    If IsNull(V) Then cmdSelect_SQL_1_Reset
    
End If

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdDetail_Delete_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
If IsNull(fraDetail_Control) Then


    mYCLISCO0_Update = "Delete"
    V = cmdUpdate
    
    If IsNull(V) Then cmdSelect_SQL_1_Reset
    
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDetail_Quit_Click()
fraDetail.Visible = False: cmdDetail_Quit.Visible = False
lstPays.Visible = False
txtRTF.Visible = False
End Sub

Private Sub cmdDetail_Update_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
If IsNull(fraDetail_Control) Then


    mYCLISCO0_Update = "Update"
    V = cmdUpdate
    
    If IsNull(V) Then cmdSelect_SQL_1_Reset
    
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdRTF_Quit_Click()
fraRTF.Visible = False
txtRTF.Visible = False
End Sub

Private Sub cmdRTF_Update_Click()
txtRTF.SaveFile paramScoring_Path & txtRTF_LoadFile
cmdRTF_Update.Visible = False
txtRTF.Locked = False
txtRTF.TextRTF = ""
End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, K As Integer, xSQL As String
On Error Resume Next


If y <= fgSelect.RowHeightMin Then
    fgSelect.Visible = False
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2:
            Select Case cmdSelect_SQL_K
                Case "2", "3": fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_SortX 2
                Case Else: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
            End Select
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_Sort
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
    End Select
    fgSelect.Visible = True
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        Select Case cmdSelect_SQL_K
            Case "1": fgSelect.Col = 0: Call fraDetail_Display(Val(fgSelect.Text))
            Case "2", "3": fraPays_Display
           End Select
        
   End If
End If
fgSelect.LeftCol = 0


End Sub

Private Sub filRTF_Click()

txtRTF_LoadFile = filRTF.FileName
txtRTF.LoadFile paramScoring_Path & txtRTF_LoadFile
cmdRTF_Update.Visible = arrHab(5)
txtRTF.Locked = Not arrHab(5)

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

mYCLISCO0_Update = ""
mYBIATAB0_Update = ""


lstErr.Clear
fgSelect.Visible = False
fraDetail.Visible = False: cmdDetail_Quit.Visible = False
lstPays.Visible = False
txtRTF.Visible = False: txtRTF.Locked = True
fraPays.Visible = False
fraRTF.Visible = False
cmdDetail_Init.Visible = False
mField = ""
txtRTF_LoadFile = ""

cmdSelect_Ok.BackColor = vbGreen
If Not IsNull(txtSelect_AMJ_Min.Value) Then
    txtSelect_AMJ_Max.Visible = True
Else
    txtSelect_AMJ_Max.Visible = False
End If
End Sub

Private Sub lblACT_Click()
mField = "ACT"
lblACT.BackColor = vbYellow
txtRTF_LoadFile = "Scoring_ACT.rtf"
txtRTF.LoadFile paramScoring_Path & txtRTF_LoadFile

lstPays.Visible = False
txtRTF.Visible = True


End Sub

Private Sub lblCLID_Click()
lstPays.Visible = False
txtRTF.Visible = False

End Sub

Private Sub lblCLIR_Click()
lstPays.Visible = False
txtRTF.Visible = False

End Sub

Private Sub lblCOB_Click()
lstPays.Visible = False
txtRTF.Visible = False

End Sub

Private Sub LBLCRS_Click()
mField = "CRS"
lblCRS.BackColor = vbYellow
txtRTF_LoadFile = "Scoring_CRS.rtf"
txtRTF.LoadFile paramScoring_Path & txtRTF_LoadFile

lstPays.Visible = False
txtRTF.Visible = True



End Sub


Private Sub lblDEB_Click()
lstPays.Visible = False
txtRTF.Visible = False

End Sub

Private Sub lblFAT_Click()
mField = "FAT"
lblFAT.BackColor = vbYellow
txtRTF_LoadFile = "Scoring_FAT.rtf"
txtRTF.LoadFile paramScoring_Path & txtRTF_LoadFile

lstPays.Visible = False
txtRTF.Visible = True


End Sub

Private Sub lblNAT_Click()
mField = "NAT"
txtRTF.Visible = False: lstPays.Visible = True
End Sub

Private Sub lblPPE_Click()
lstPays.Visible = False
txtRTF.Visible = False

End Sub

Private Sub lblRES_Click()
mField = "RES"
txtRTF.Visible = False: lstPays.Visible = True
End Sub


Private Sub lblScore_Click()
lstPays.Visible = False
txtRTF.Visible = False

End Sub

Private Sub libSCO_Click()

Select Case newYCLISCO0.CLISCOSCOK

    Case "1": txtRTF_LoadFile = "Scoring_V1.rtf"
    Case "2": txtRTF_LoadFile = "Scoring_V2.rtf"
    Case "3": txtRTF_LoadFile = "Scoring_V3.rtf"
    Case Else: txtRTF.TextRTF = "": txtRTF_LoadFile = ""
End Select


If txtRTF_LoadFile <> "" Then txtRTF.LoadFile paramScoring_Path & txtRTF_LoadFile

txtRTF.Visible = True


txtRTF.Visible = True

End Sub


Private Sub lstPays_Click()


If Not blnPays_Scan Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    
    Call lstPays_Display(lstPays.Text)
    
    CLISCOSCOV_Total
    Me.Enabled = True: Me.MousePointer = 0
End If
End Sub

Private Sub lstPays_Display(lPays As String)
Dim X As String, X2 As String, K As Integer, wColor As Long, libX As String, xPlus As String

wColor = fraDetail.BackColor ' RGB(255, 255, 255)
mSCOPAY.V = 0
mSCOPAY.BIA_CA = "N"

X = lstPays.Text
X2 = Mid$(X, 1, 2)


For K = 1 To arrSCOPAY_Nb
    If X2 = arrSCOPAY(K).Id Then
        mSCOPAY = arrSCOPAY(K)
        If mSCOPAY.GAFI_N = "Y" Then
            wColor = vbRed
            
        Else
            If mSCOPAY.GAFI_G = "Y" Then
               wColor = RGB(192, 192, 192)
               
            Else
                If mSCOPAY.Embargo = "Y" Then
                   wColor = mColor_W1
                   
                Else
                    If mSCOPAY.CRS = "Y" Then
                        wColor = mColor_G1
                        
                    Else
                        If mSCOPAY.BIA_CA = "Y" Then
                            wColor = vbRed 'mColor_Y2
                        Else
                            If mSCOPAY.BIA_1 = "Y" Then
                                wColor = chkBIA_1.BackColor
                            Else
                                If mSCOPAY.BIA_2 = "Y" Then
                                    wColor = chkBIA_2.BackColor
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        If mSCOPAY.GAFI_N = "Y" Then libX = "GAFI NOIR": xPlus = " + "
        If mSCOPAY.GAFI_G = "Y" Then libX = libX & xPlus & "GAFI GRIS": xPlus = " + "
        If mSCOPAY.Embargo = "Y" Then libX = libX & xPlus & "Embargo": xPlus = " + "
        If mSCOPAY.CRS = "Y" Then libX = libX & xPlus & "CRS": xPlus = " + "
        If mSCOPAY.BIA_CA = "Y" Then libX = libX & xPlus & "BIA-CA": xPlus = " + "
        If mSCOPAY.BIA_1 = "Y" Then libX = libX & xPlus & "BIA-1": xPlus = " + "
        If mSCOPAY.BIA_2 = "Y" Then libX = libX & xPlus & "BIA-2": xPlus = " + "

        Exit For
    End If
Next K
   
Select Case mField
    Case "RES": libRES = X: libRES.BackColor = wColor
                newYCLISCO0.CLISCORESK = X2
                newYCLISCO0.CLISCORESV = mSCOPAY.V
                txtRESV = newYCLISCO0.CLISCORESV
                libRESV = libX: libRESV.BackColor = wColor
                blnNoRelation_RESK = IIf(mSCOPAY.BIA_CA = "Y", True, False)
                
    Case "NAT": libNAT = X: libNAT.BackColor = wColor
                newYCLISCO0.CLISCONATK = X2
                newYCLISCO0.CLISCONATV = mSCOPAY.V
                txtNATV = newYCLISCO0.CLISCONATV
                libNATV = libX: libNATV.BackColor = wColor
                blnNoRelation_NATK = IIf(mSCOPAY.BIA_CA = "Y", True, False)
End Select

CLISCOSCOV_Total

End Sub


Private Sub txtCLID_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtCLID_LostFocus()
If Not cmdDetail_Update.Visible Then
    If Trim(txtCLID) <> "" Then Call ZCLIENA0_Read(Trim(txtCLID))
End If
End Sub


Private Sub txtFATV_Change()
lblFAT.FontBold = False
End Sub


Private Sub txtScore_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub

Private Sub txtSelect_AMJ_Max_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_AMJ_Min_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_CLISCOID_Change()
cmdSelect_Clear

End Sub


Private Sub txtSelect_CLISCOID_GotFocus()
Call txt_GotFocus(txtSelect_CLISCOID)

End Sub


Private Sub txtSelect_CLISCOID_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtSelect_CLISCOID_LostFocus()
Call txt_LostFocus(txtSelect_CLISCOID)
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
    Select Case cmdSelect_SQL_K
        Case "1": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
                 cmdDetail_Init.Visible = arrHab(2)
    End Select

End If

End Sub


Private Sub cmdSelect_SQL_1()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"
xWhere = ""
If Trim(txtSelect_CLISCOID) <> "" Then
    xWhere = " Where CLISCOID = " & Val(Trim(txtSelect_CLISCOID))
End If

If Not IsNull(txtSelect_AMJ_Min.Value) Then
    Call DTPicker_Control(txtSelect_AMJ_Min, wAMJMin)
    Call DTPicker_Control(txtSelect_AMJ_Max, WAMJMax)
    If xWhere = "" Then
        xWhere = " where"
    Else
        xWhere = " and"
    End If
    xWhere = xWhere & " CLISCOYAMJ >= " & wAMJMin & " And CLISCOYAMJ <= " & WAMJMax
End If

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCLISCO0 " & xWhere & " order by CLISCOID"
Set rsSab = cnsab.Execute(xSQL)
  
Call fgSelect_Display_1

Set rsSab = Nothing

If fgSelect.Rows = 2 Then fgSelect.Row = 1: fgSelect.Col = 0: Call fraDetail_Display(Val(fgSelect.Text))
'fraDetail_Init

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_2()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'SAB' and BIATABK1 = 'CLIENAPAY' order by BIATABK2"
Set rsSab = cnsab.Execute(xSQL)
  
Call fgSelect_Display_2

Set rsSab = Nothing


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_9()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_9"

fraRTF.Visible = True
txtRTF.Visible = True
txtRTF.TextRTF = ""
cmdRTF_Update.Visible = False
txtRTF.Locked = Not arrHab(5)

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


Public Sub cmdContext_Return()
    If SSTab1.Tab = 0 Then
        If cmdPays_Update.Visible Then
            cmdPays_Update_Click
        Else
            'cmdSelect_Ok_Click
        End If
    Else
        SendKeys "{TAB}"
    End If
End Sub


Public Sub cmdContext_Quit()
lstErr.Clear: lstErr.Height = 200

If fraPays.Visible Then
    fraPays.Visible = False
    Exit Sub
End If

If fraRTF.Visible Then
    fraRTF.Visible = False
    txtRTF.Visible = False
    Exit Sub
End If

If fraDetail.Visible Then
    If lstPays.Visible Then
        lstPays.Visible = False
        Exit Sub
    Else
        If txtRTF.Visible Then
            txtRTF.Visible = False
            mField = ""
            txtRTF_LoadFile = ""
            Exit Sub
        Else
            fraDetail.Visible = False: cmdDetail_Quit.Visible = False
            Exit Sub
        End If
    End If
    
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
Call lstErr_Clear(lstErr, cmdContext, "> YCLISCO0_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Clear

Select Case cmdSelect_SQL_K
    Case "1": cmdSelect_SQL_1
    Case "2": cmdSelect_SQL_2
    Case "3": cmdSelect_SQL_2
    Case "9": cmdSelect_SQL_9
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< YCLISCO0_cmdSelect_Ok"): DoEvents
lstErr.Height = 480
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
cmdSelect_Ok.BackColor = fgSelect.BackColorFixed
End Sub




Private Sub mnuPrint_Excel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String
Call lstErr_AddItem(lstErr, cmdContext, "YCLISCO0 : export Excel ...."): DoEvents
    Select Case cmdSelect_SQL_K
        Case "1":
            X = "Scoring : Liste prospect"
            Call MSflexGrid_Excel("", "Scoring prospect", X, fgSelect, 7)
        Case "2", "3":
            X = "Scoring : Liste Pays"
            Call MSflexGrid_Excel("", "Scoring pays", X, fgSelect, 11)
    End Select

Call lstErr_AddItem(lstErr, cmdContext, "export Excel terminé"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint_Mail_Click()
Dim X As String

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_AddItem(lstErr, cmdContext, "> YCLISCO0 :  : export mail ...."): DoEvents
    Select Case cmdSelect_SQL_K
        Case "1":
            X = "Scoring : Liste prospect"
            Call MSFlexGrid_SendMail(mMail_Destinataires, "Scoring", X, X, fgSelect, 7)
         Case "2", "3":
            X = "Scoring : Liste Pays"
            Call MSFlexGrid_SendMail(mMail_Destinataires, "Scoring pays", X, X, fgSelect, 11)
   End Select

Call lstErr_AddItem(lstErr, cmdContext, "export mail terminé"): DoEvents


Me.Enabled = True: Me.MousePointer = 0

End Sub







Public Sub fraDetail_Display(lCLISCOID As Long)
Dim V, X As String, K As Long
Dim xSQL As String
On Error GoTo Error_Handler

currentAction = "fraDetail_Display"
fraDetail_Init

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCLISCO0 where CLISCOID = " & lCLISCOID
Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    V = "Erreur de lecture " & xSQL
    GoTo Error_MsgBox:
End If

Call rsYCLISCO0_GetBuffer(rsSab, oldYCLISCO0)
newYCLISCO0 = oldYCLISCO0

fraDetail.Caption = oldYCLISCO0.CLISCOID

txtCLIR = Trim(oldYCLISCO0.CLISCOCLIR)
txtCLID = Trim(oldYCLISCO0.CLISCOCLID)
txtCLIX = Trim(oldYCLISCO0.CLISCOCLIX)
txtInfo = Trim(oldYCLISCO0.CLISCOINFO)

txtPPEV = oldYCLISCO0.CLISCOPPEV
If oldYCLISCO0.CLISCOPPEK = "Y" Then
    cmdPPE_Oui.BackColor = mColor_W1
    cmdPPE_Non.BackColor = RGB(230, 230, 230)
Else
    cmdPPE_Non.BackColor = mColor_G2
    cmdPPE_Oui.BackColor = RGB(230, 230, 230)
End If



txtDEBV = oldYCLISCO0.CLISCODEBV
Select Case oldYCLISCO0.CLISCODEBK
    Case "Y"
        cmdDEB_Oui.BackColor = mColor_G2
        cmdDEB_Non.BackColor = RGB(230, 230, 230)
        cmdDEB_Technique.BackColor = RGB(230, 230, 230)
    Case "N"
        cmdDEB_Non.BackColor = mColor_W1
        cmdDEB_Oui.BackColor = RGB(230, 230, 230)
        cmdDEB_Technique.BackColor = RGB(230, 230, 230)
    Case Else
        cmdDEB_Non.BackColor = RGB(230, 230, 230)
        cmdDEB_Oui.BackColor = RGB(230, 230, 230)
        cmdDEB_Technique.BackColor = mColor_Y1
End Select

blnPays_Scan = True

mField = "RES"
For K = 0 To lstPays.ListCount - 1
    lstPays.ListIndex = K
    If oldYCLISCO0.CLISCORESK = Mid$(lstPays, 1, 2) Then Exit For
Next K
Call lstPays_Display(lstPays.Text)

mField = "NAT"
For K = 0 To lstPays.ListCount - 1
    lstPays.ListIndex = K
    If oldYCLISCO0.CLISCONATK = Mid$(lstPays, 1, 2) Then Exit For
Next K
Call lstPays_Display(lstPays.Text)

blnPays_Scan = False

txtACTV = oldYCLISCO0.CLISCOACTV
If oldYCLISCO0.CLISCOACTK = "Y" Then
    cmdACT_Oui.BackColor = mColor_W1
    cmdACT_Non.BackColor = RGB(230, 230, 230)
Else
    cmdACT_Non.BackColor = mColor_G2
    cmdACT_Oui.BackColor = RGB(230, 230, 230)
End If

txtCOBV = oldYCLISCO0.CLISCOCOBV
If oldYCLISCO0.CLISCOCOBK = "Y" Then
    cmdCOB_Oui.BackColor = mColor_W1
    cmdCOB_Non.BackColor = RGB(230, 230, 230)
Else
    cmdCOB_Non.BackColor = mColor_G2
    cmdCOB_Oui.BackColor = RGB(230, 230, 230)
End If

txtCRSV = oldYCLISCO0.CLISCOCRSV
If oldYCLISCO0.CLISCOCRSK = "Y" Then
    cmdCRS_Oui.BackColor = mColor_W1
    cmdCRS_Non.BackColor = RGB(230, 230, 230)
Else
    cmdCRS_Non.BackColor = mColor_G2
    cmdCRS_Oui.BackColor = RGB(230, 230, 230)
End If

txtFATV = oldYCLISCO0.CLISCOFATV
If oldYCLISCO0.CLISCOFATK = "Y" Then
    cmdFAT_Oui.BackColor = mColor_W1
    cmdFAT_Non.BackColor = RGB(230, 230, 230)
Else
    cmdFAT_Non.BackColor = mColor_G2
    cmdFAT_Oui.BackColor = RGB(230, 230, 230)
End If

If arrHab(3) Then
    cmdDetail_Update.Visible = True
    cmdDetail_Delete.Visible = True
Else
    If Trim(oldYCLISCO0.CLISCOYUSR) = usrName_UCase Then
        cmdDetail_Update.Visible = True
        cmdDetail_Delete.Visible = True
    End If
End If

fraDetail.Enabled = cmdDetail_Update.Visible

currentAction = "fraDetail_Display_Fin"
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fraDetail_Init()

Call rsYCLISCO0_Init(newYCLISCO0)

cmdDetail_Update.Visible = False
cmdDetail_Delete.Visible = False

txtCLIR = usrName_UCase
txtCLID = ""
txtCLIX = ""
txtInfo = ""
txtRTF.TextRTF = ""
txtRTF.Visible = False
lstPays.Visible = False

cmdPPE_Oui.BackColor = RGB(230, 230, 230)
cmdPPE_Non.BackColor = RGB(230, 230, 230)
txtPPEV = ""
cmdDEB_Oui.BackColor = RGB(230, 230, 230)
cmdDEB_Non.BackColor = RGB(230, 230, 230)
cmdDEB_Technique.BackColor = RGB(230, 230, 230)
txtDEBV = ""
cmdACT_Oui.BackColor = RGB(230, 230, 230)
cmdACT_Non.BackColor = RGB(230, 230, 230)
txtACTV = ""
cmdCOB_Oui.BackColor = RGB(230, 230, 230)
cmdCOB_Non.BackColor = RGB(230, 230, 230)
txtCOBV = ""
cmdFAT_Oui.BackColor = RGB(230, 230, 230)
cmdFAT_Non.BackColor = RGB(230, 230, 230)
txtFATV = ""
cmdCRS_Oui.BackColor = RGB(230, 230, 230)
cmdCRS_Non.BackColor = RGB(230, 230, 230)
txtCRSV = ""

libRES.BackColor = RGB(230, 230, 230)
libRES = ""

libNAT.BackColor = RGB(230, 230, 230)
libNAT = ""

libSCO.BackColor = RGB(230, 230, 230)
libSCO = ""
txtSCOV = ""

libNoRelation.Visible = False
libRESV = "": libRESV.BackColor = fraDetail.BackColor: blnNoRelation_RESK = False
libNATV = "": libNATV.BackColor = fraDetail.BackColor: blnNoRelation_NATK = False

lblPPE.BackColor = vbYellow
fraDetail.Visible = True: fraDetail.Enabled = arrHab(2)
cmdDetail_Quit.Visible = True: cmdDetail_Add.Visible = arrHab(2)
End Sub

Public Sub CLISCOSCOV_Total()
Dim wColor As Long

lstPays.Visible = False
txtRTF.Visible = False

newYCLISCO0.CLISCOSCOV = newYCLISCO0.CLISCOPPEV + newYCLISCO0.CLISCODEBV _
                        + newYCLISCO0.CLISCORESV + newYCLISCO0.CLISCONATV _
                        + newYCLISCO0.CLISCOACTV + newYCLISCO0.CLISCOCOBV _
                        + newYCLISCO0.CLISCOFATV + newYCLISCO0.CLISCOCRSV
txtSCOV = newYCLISCO0.CLISCOSCOV

Select Case newYCLISCO0.CLISCOSCOV
    Case Is < 50: newYCLISCO0.CLISCOSCOK = 1: wColor = mColor_G2
    Case Is < 100: newYCLISCO0.CLISCOSCOK = 2: wColor = mColor_Y1
    Case Else: newYCLISCO0.CLISCOSCOK = 3: wColor = mColor_W1
End Select

libSCO = "Vigilance " & newYCLISCO0.CLISCOSCOK
libSCO.BackColor = wColor

lblPPE.BackColor = fraDetail.BackColor
lblDEB.BackColor = fraDetail.BackColor
lblRES.BackColor = fraDetail.BackColor
lblNAT.BackColor = fraDetail.BackColor
lblACT.BackColor = fraDetail.BackColor
lblCOB.BackColor = fraDetail.BackColor
lblFAT.BackColor = fraDetail.BackColor
lblCRS.BackColor = fraDetail.BackColor

If InStr(txtRTF_LoadFile, "_V") > 0 Then libSCO_Click

If newYCLISCO0.CLISCOPPEK = "" Then
    lblPPE.BackColor = vbYellow: mField = "PPE"
Else
    If newYCLISCO0.CLISCODEBK = "" Then
        lblDEB.BackColor = vbYellow: mField = "DEB"
    Else
        If newYCLISCO0.CLISCORESK = "" Then
            lblRES.BackColor = vbYellow: lblRES_Click
        Else
            If newYCLISCO0.CLISCONATK = "" Then
                lblNAT.BackColor = vbYellow: lblNAT_Click
            Else
                If newYCLISCO0.CLISCOACTK = "" Then
                    lblACT.BackColor = vbYellow: lblACT_Click
                Else
                    If newYCLISCO0.CLISCOCOBK = "" Then
                        lblCOB.BackColor = vbYellow: mField = "COB"
                    Else
                        If newYCLISCO0.CLISCOFATK = "" Then
                            lblFAT.BackColor = vbYellow: lblFAT_Click
                        Else
                            If newYCLISCO0.CLISCOCRSK = "" Then
                                lblCRS.BackColor = vbYellow: LBLCRS_Click
                            Else
                                libSCO_Click
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

If newYCLISCO0.CLISCOFATK = "Y" Or blnNoRelation_RESK = True Or blnNoRelation_NATK = True Then
    libNoRelation.Visible = True
Else
    libNoRelation.Visible = False
End If

End Sub

Public Sub arrSCOPAY_Load()
Dim X As String

X = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID = 'Scoring' and BIATABK1 = 'Pays'"
Set rsSab = cnsab.Execute(X)
ReDim arrSCOPAY(rsSab(0) + 1)

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID = 'Scoring' and BIATABK1 = 'Pays' order by BIATABK2"
    
Set rsSab = cnsab.Execute(X)
arrSCOPAY_Nb = 0

Do While Not rsSab.EOF
        arrSCOPAY_Nb = arrSCOPAY_Nb + 1
        arrSCOPAY(arrSCOPAY_Nb).Id = Trim(rsSab("BIATABK2"))
        X = Trim(rsSab("BIATABTXT"))
        arrSCOPAY(arrSCOPAY_Nb).V = Val(Mid$(X, 1, 3))
        arrSCOPAY(arrSCOPAY_Nb).GAFI_N = Mid$(X, 4, 1)
        arrSCOPAY(arrSCOPAY_Nb).GAFI_G = Mid$(X, 5, 1)
        arrSCOPAY(arrSCOPAY_Nb).Embargo = Mid$(X, 6, 1)
        arrSCOPAY(arrSCOPAY_Nb).CRS = Mid$(X, 7, 1)
        arrSCOPAY(arrSCOPAY_Nb).BIA_CA = Mid$(X, 8, 1)
        arrSCOPAY(arrSCOPAY_Nb).BIA_1 = Mid$(X, 9, 1)
        arrSCOPAY(arrSCOPAY_Nb).BIA_2 = Mid$(X, 10, 1)
        arrSCOPAY(arrSCOPAY_Nb).YUSR = Mid$(X, 105, 10)
        arrSCOPAY(arrSCOPAY_Nb).YAMJ = Val(Mid$(X, 115, 8))
        arrSCOPAY(arrSCOPAY_Nb).YHMS = Val(Mid$(X, 123, 6))
    rsSab.MoveNext
Loop


End Sub

Public Sub fraPays_Display()
Dim X2 As String, K As Integer, xSQL As String

mYBIATAB0_Update = "New"
txtScore = ""
chkGAFI_N.Value = 0
chkGAFI_G.Value = 0
chkEmbargo.Value = 0
chkCRS.Value = 0
chkBIA_CA.Value = 0
chkBIA_1.Value = 0
chkBIA_2.Value = 0
fgSelect.Col = 0: X2 = fgSelect.Text
fgSelect.Col = 1: libPays = X2 & " - " & Trim(fgSelect.Text)

oldParam.BIATABID = "Scoring"
oldParam.BIATABK1 = "Pays"
oldParam.BIATABK2 = X2
oldParam.BIATABTXT = "000NNNN"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'Scoring' and BIATABK1 = 'Pays'  and BIATABK2 = '" & X2 & "'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    mYBIATAB0_Update = "Update"
    oldParam.BIATABTXT = rsSab("BIATABTXT")

    txtScore = Val(Mid$(oldParam.BIATABTXT, 1, 3))
    If Mid$(oldParam.BIATABTXT, 4, 1) = "Y" Then chkGAFI_N.Value = 1
    If Mid$(oldParam.BIATABTXT, 5, 1) = "Y" Then chkGAFI_G.Value = 1
    If Mid$(oldParam.BIATABTXT, 6, 1) = "Y" Then chkEmbargo.Value = 1
    If Mid$(oldParam.BIATABTXT, 7, 1) = "Y" Then chkCRS = 1
    If Mid$(oldParam.BIATABTXT, 8, 1) = "Y" Then chkBIA_CA = 1
    If Mid$(oldParam.BIATABTXT, 9, 1) = "Y" Then chkBIA_1 = 1
    If Mid$(oldParam.BIATABTXT, 10, 1) = "Y" Then chkBIA_2 = 1

End If
cmdPays_Update.Visible = arrHab(4)

fraPays.Visible = True
End Sub

Public Function cmdUpdate()
Dim V
On Error GoTo Error_Handler

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYCLISCO0_Update
    Case "Update": V = sqlYCLISCO0_Update(newYCLISCO0, oldYCLISCO0)
    Case "New": V = sqlYCLISCO0_Insert(newYCLISCO0)
    Case "Delete": V = sqlYCLISCO0_Delete(oldYCLISCO0)
End Select
'________________________________________________________________________________

If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case mYBIATAB0_Update
    Case "Update": V = sqlYBIATAB0_Update(newParam, oldParam)
    Case "New": V = sqlYBIATAB0_Insert(newParam)
End Select
'________________________________________________________________________________
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
        mYCLISCO0_Update = ""
        mYBIATAB0_Update = ""
    End If
    
    cmdUpdate = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function

Public Sub CLISCOSCOV_Pays()
txtScore = 0
If chkGAFI_N = 1 Then
    If chkEmbargo = 1 Then
        txtScore = 200
    Else
        txtScore = 100
    End If
End If
If chkGAFI_G = 1 Then txtScore = txtScore + 50
If chkEmbargo = 1 Then txtScore = txtScore + 100
    
End Sub

Public Function fraDetail_Control()
Dim X As String, wMsg As String, xSQL As String
Dim K As Integer
Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données "): DoEvents

wMsg = ""
newYCLISCO0.CLISCOCLIR = Trim(txtCLIR)
If newYCLISCO0.CLISCOCLIR = "" Then wMsg = wMsg & "- préciser le chargé de clientèle" & vbCrLf
newYCLISCO0.CLISCOCLIX = Trim(txtCLIX)
If newYCLISCO0.CLISCOCLIX = "" Then wMsg = wMsg & "- préciser l'intitulé du prospect" & vbCrLf
newYCLISCO0.CLISCOCLID = Trim(txtCLID)
If newYCLISCO0.CLISCOPPEK = "" Then wMsg = wMsg & "- préciser l'information 'PPE'" & vbCrLf
If newYCLISCO0.CLISCODEBK = "" Then wMsg = wMsg & "- préciser l'information 'entrée en relation'" & vbCrLf
If newYCLISCO0.CLISCORESK = "" Then wMsg = wMsg & "- préciser l'information 'pays de résidence'" & vbCrLf
If newYCLISCO0.CLISCONATK = "" Then wMsg = wMsg & "- préciser l'information 'pays de nationalité'" & vbCrLf
If newYCLISCO0.CLISCOACTK = "" Then wMsg = wMsg & "- préciser l'information 'Activité'" & vbCrLf
If newYCLISCO0.CLISCOCOBK = "" Then wMsg = wMsg & "- préciser l'information 'Correspondant Banking'" & vbCrLf
If newYCLISCO0.CLISCOFATK = "" Then wMsg = wMsg & "- préciser l'information 'FATCA'" & vbCrLf
If newYCLISCO0.CLISCOCRSK = "" Then wMsg = wMsg & "- préciser l'information 'CRS'" & vbCrLf

newYCLISCO0.CLISCOINFO = Trim(txtInfo)
        

'__________________________________________________________________________

If wMsg = "" Then
    fraDetail_Control = Null
    
    newYCLISCO0.CLISCOYUSR = usrName_UCase
    newYCLISCO0.CLISCOYAMJ = DSys
    newYCLISCO0.CLISCOYHMS = time_Hms

Else
    Call MsgBox(wMsg, vbCritical, "Scoring : Contrôle des informations")
    fraDetail_Control = "?_________fraDetail_Control"
End If

End Function

Public Sub cmdSelect_SQL_1_Reset()
fraDetail.Visible = False: cmdDetail_Quit.Visible = False
txtRTF.Visible = False
lstPays.Visible = False

cmdSelect_SQL_1

End Sub

Public Sub ZCLIENA0_Read(lCLI As String)
Dim xSQL As String, K As Long

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = " & Format(lCLI, "00000000")
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    txtCLIX = Trim(rsSab("CLIENARA1"))
    
    blnPays_Scan = True
    
    newYCLISCO0.CLISCORESK = Trim(rsSab("CLIENARSD"))
    mField = "RES"
    For K = 0 To lstPays.ListCount - 1
        lstPays.ListIndex = K
        If newYCLISCO0.CLISCORESK = Mid$(lstPays, 1, 2) Then Exit For
    Next K
    Call lstPays_Display(lstPays.Text)
    
    newYCLISCO0.CLISCONATK = Trim(rsSab("CLIENANAT"))
    mField = "NAT"
    For K = 0 To lstPays.ListCount - 1
        lstPays.ListIndex = K
        If newYCLISCO0.CLISCONATK = Mid$(lstPays, 1, 2) Then Exit For
    Next K
    Call lstPays_Display(lstPays.Text)
    
    blnPays_Scan = False
End If

End Sub
