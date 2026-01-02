VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmYCOMRCD0 
   AutoRedraw      =   -1  'True
   Caption         =   "Commissions : rétrocession"
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
   Icon            =   "YCOMRCD0.frx":0000
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
      Left            =   30
      TabIndex        =   3
      Top             =   435
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
      TabPicture(0)   =   "YCOMRCD0.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "YCOMRCD0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFg"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "YCOMRCD0.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
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
         Text            =   "YCOMRCD0.frx":035E
         Top             =   1155
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Frame fraSelect 
         BackColor       =   &H00E0E0E0&
         Height          =   11055
         Left            =   75
         TabIndex        =   4
         Top             =   525
         Width           =   16155
         Begin VB.Frame fraCOMRCDCLI 
            BackColor       =   &H00E0E0E0&
            Caption         =   "fraCOMRCDCLI"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   7755
            Left            =   6345
            TabIndex        =   21
            Top             =   2310
            Visible         =   0   'False
            Width           =   7500
            Begin VB.CommandButton cmdCOMRCDCLI_New 
               BackColor       =   &H000000FF&
               Caption         =   "Ajouter un client"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   5625
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   6435
               Width           =   1650
            End
            Begin VB.Frame fraCOMRCDCLI_1 
               BackColor       =   &H00E0FFFF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3015
               Left            =   105
               TabIndex        =   35
               Top             =   2940
               Width           =   7305
               Begin VB.TextBox txtCOMRCDZCOM 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   350
                  Left            =   2130
                  MaxLength       =   6
                  TabIndex        =   36
                  Top             =   660
                  Width           =   1710
               End
               Begin VB.TextBox txtCOMRCDMTD 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   350
                  Left            =   2145
                  MaxLength       =   12
                  TabIndex        =   37
                  Top             =   1305
                  Width           =   1680
               End
               Begin VB.TextBox txtCOMRCDMTR 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   350
                  Left            =   2160
                  MaxLength       =   12
                  TabIndex        =   38
                  Top             =   1875
                  Width           =   1680
               End
               Begin MSComCtl2.DTPicker txtCOMRCDDTR 
                  Height          =   300
                  Left            =   2190
                  TabIndex        =   39
                  Top             =   2520
                  Width           =   1665
                  _ExtentX        =   2937
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
                  Format          =   3735555
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin VB.Label lblCOMRCDDTR 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Date de début"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   375
                  Left            =   255
                  TabIndex        =   43
                  Top             =   2490
                  Width           =   1575
               End
               Begin VB.Label lblCOMRCDZCOM 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Code commission"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   540
                  Left            =   240
                  TabIndex        =   42
                  Top             =   690
                  Width           =   1920
               End
               Begin VB.Label lblCOMRCDMTD 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Montant minimum"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   375
                  Left            =   195
                  TabIndex        =   41
                  Top             =   1335
                  Width           =   1935
               End
               Begin VB.Label lblCOMRCDMTR 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Mt rétrocédé"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   375
                  Left            =   195
                  TabIndex        =   40
                  Top             =   1935
                  Width           =   1575
               End
            End
            Begin VB.Frame fraCOMRCDCLI_0 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2235
               Left            =   75
               TabIndex        =   27
               Top             =   435
               Width           =   7305
               Begin VB.TextBox txtCOMRCDOPE 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   350
                  Left            =   2250
                  MaxLength       =   3
                  TabIndex        =   34
                  Top             =   1485
                  Width           =   855
               End
               Begin VB.TextBox txtCOMRCDSSE 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   350
                  Left            =   3315
                  MaxLength       =   2
                  TabIndex        =   33
                  Top             =   930
                  Width           =   585
               End
               Begin VB.TextBox txtCOMRCDSER 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   350
                  Left            =   2250
                  MaxLength       =   2
                  TabIndex        =   32
                  Top             =   960
                  Width           =   585
               End
               Begin VB.TextBox txtCOMRCDCLI 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   350
                  Left            =   2250
                  MaxLength       =   7
                  TabIndex        =   29
                  Top             =   360
                  Width           =   1710
               End
               Begin VB.Label lblCOMRCDOPE 
                  Caption         =   "Code opération"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   375
                  Left            =   225
                  TabIndex        =   31
                  Top             =   1440
                  Width           =   1575
               End
               Begin VB.Label lblCOMRCDSER 
                  Caption         =   "Service"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   375
                  Left            =   225
                  TabIndex        =   30
                  Top             =   975
                  Width           =   1395
               End
               Begin VB.Label lblCOMRCDCLI 
                  Caption         =   "Racine"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   375
                  Left            =   225
                  TabIndex        =   28
                  Top             =   375
                  Width           =   1395
               End
            End
            Begin VB.CommandButton cmdCOMRCDCLI_Quit 
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
               Height          =   630
               Left            =   285
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   6435
               Width           =   1260
            End
            Begin VB.CommandButton cmdCOMRCDCLI_Update 
               BackColor       =   &H0080FF80&
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
               Height          =   630
               Left            =   5070
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   6465
               Width           =   1650
            End
            Begin VB.CommandButton cmdCOMRCDCLI_Add 
               BackColor       =   &H000080FF&
               Caption         =   "Ajouter un code commission"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   3630
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   6435
               Width           =   1455
            End
            Begin VB.CommandButton cmdCOMRCDCLI_Delete 
               BackColor       =   &H00FF80FF&
               Caption         =   "Suspendre / restaurer un code commission"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   1665
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   6420
               Width           =   1725
            End
            Begin VB.Label lblParam_Mnu_Quid 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblParam_quid"
               Height          =   330
               Left            =   255
               TabIndex        =   26
               Top             =   7170
               Width           =   6825
            End
         End
         Begin VB.CommandButton cmdCOMRCDRLV_Ok 
            BackColor       =   &H0000FF00&
            Caption         =   "Validation de ce relevé"
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
            Left            =   11865
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   630
            Visible         =   0   'False
            Width           =   1335
         End
         Begin RichTextLib.RichTextBox txtRTF 
            Height          =   3735
            Left            =   960
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   7185
            Visible         =   0   'False
            Width           =   14775
            _ExtentX        =   26061
            _ExtentY        =   6588
            _Version        =   393217
            BackColor       =   14745599
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"YCOMRCD0.frx":0366
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
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   9675
            Left            =   10140
            TabIndex        =   12
            Top             =   1320
            Visible         =   0   'False
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   17066
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
            FormatString    =   "<? |<Opération                              |< Date                 |> N° CRE          |<Doc |<Comment|>nb    |"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
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
            Left            =   13905
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   645
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
            Left            =   135
            TabIndex        =   5
            Top             =   105
            Visible         =   0   'False
            Width           =   11205
            Begin VB.TextBox txtSelect_COMRCDRLV 
               Height          =   285
               Left            =   5505
               TabIndex        =   19
               Top             =   360
               Width           =   1575
            End
            Begin VB.ComboBox cboSelect_COMRCDCLI 
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
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   315
               Width           =   2430
            End
            Begin MSComCtl2.DTPicker txtSelect_COMRCDDTR_Min 
               Height          =   300
               Left            =   9315
               TabIndex        =   16
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
               Format          =   3735555
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_COMRCDDTR_Max 
               Height          =   300
               Left            =   9555
               TabIndex        =   17
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
               Format          =   3735555
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_COMRCDRLV 
               BackColor       =   &H00F0FFFF&
               Caption         =   "N° relevé"
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
               Left            =   3945
               TabIndex        =   18
               Top             =   405
               Width           =   1110
            End
            Begin VB.Label lblSelect_COMRCDDTR 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Période"
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
               Left            =   8115
               TabIndex        =   15
               Top             =   480
               Width           =   855
            End
            Begin VB.Label libCOMRCDCLI 
               BackColor       =   &H00F0FFFF&
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   270
               Left            =   1155
               TabIndex        =   14
               Top             =   795
               Width           =   4575
            End
            Begin VB.Label lblSelect_COMRCDCLI 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Racine"
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
               Left            =   210
               TabIndex        =   10
               Top             =   300
               Width           =   1155
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   9750
            Left            =   120
            TabIndex        =   9
            Top             =   1260
            Width           =   15825
            _ExtentX        =   27914
            _ExtentY        =   17198
            _Version        =   393216
            Cols            =   11
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
            AllowUserResizing=   3
            FormatString    =   $"YCOMRCD0.frx":03E6
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
      Picture         =   "YCOMRCD0.frx":04EA
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
   Begin VB.Menu mnuUpdate 
      Caption         =   "mnuUpdate"
      Visible         =   0   'False
      Begin VB.Menu mnuUpdate_Display 
         Caption         =   "Afficher la pièce comptable"
      End
      Begin VB.Menu mnuUpdate_Ignorer 
         Caption         =   "Ignorer cette écriture"
      End
      Begin VB.Menu mnuUpdate_MTR 
         Caption         =   "Modifier le mt rétrocédé"
      End
      Begin VB.Menu mnuUpdate_Restaurer 
         Caption         =   "Restaurer cette ligne"
      End
   End
End
Attribute VB_Name = "frmYCOMRCD0"
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

'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long

Dim xYCOMRCD0 As typeYCOMRCD0, oldYCOMRCD0 As typeYCOMRCD0, newYCOMRCD0 As typeYCOMRCD0
Dim HeightOfLine As Long, LinesOfText As Long

Dim txtRTF_prtForeColor_Header As Long


Dim VB_RTF_Modèle As String
Dim mCOMRCDMTR As Currency, mCOMRCDCLI As String, mCOMRCDOPE As String, mCOMRCDRLV_Where As String, mCOMRCDRLV As Long, mCOMRCDNAT As String
Dim mCOMRCDSER As String, mCOMRCDSSE As String
Dim mCOMRCDMTD As Currency
Dim mExportation As String

Dim xYBIAMVTH As typeYBIAMVT0
Dim blnCOMRCDZCOM_Add As Boolean, blnCOMRCDCLI_Add As Boolean

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

currentAction = "fgDetail_Display"
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.Row = 0
'___________________________________________________________________________

  
Do While Not rsSab.EOF

    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    'Call rsYCOMRCD0_GetBuffer(rsSab, xYCOMRCD0)
    fgDetail_Display_Line
    
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



Public Sub fgDetail_Display_Line()
On Error Resume Next

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
Call DTPicker_Set(txtSelect_COMRCDDTR_Min, YBIATAB0_DATE_CPT_MP1)
Call DTPicker_Set(txtSelect_COMRCDDTR_Max, YBIATAB0_DATE_CPT_J)
txtSelect_COMRCDDTR_Min.Value = Null

fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False

fgDetail_FormatString = fgDetail.FormatString
fgDetail.Enabled = True
fgDetail.Visible = False
fgDetail.Top = fgSelect.Top
fgDetail.Left = fgSelect.Left + fgSelect.Width - fgDetail.Width - 200

fraSelect_Options.Visible = True

Set fraCOMRCDCLI.Container = fgSelect.Container
fraCOMRCDCLI.Left = fgSelect.Left + fgSelect.Width - fraCOMRCDCLI.Width - 200
fraCOMRCDCLI.Top = fgSelect.Top
fraCOMRCDCLI.Visible = False
fraCOMRCDCLI.ForeColor = vbMagenta
fraCOMRCDCLI_1.ForeColor = vbMagenta

If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0
cboSelect_COMRCDCLI.Clear
cboSelect_COMRCDCLI.AddItem " "

xSQL = "select distinct COMRCDNAT , COMRCDCLI , COMRCDOPE  , COMRCDSER  , COMRCDSSE from " & paramIBM_Library_SABSPE & ".YCOMRCD0 " _
     & " order by COMRCDNAT desc , COMRCDCLI  , COMRCDOPE  , COMRCDSER  , COMRCDSSE"

Set rsSab_X = cnsab.Execute(xSQL)
Do Until rsSab_X.EOF
    If rsSab_X("COMRCDNAT") <> "#" Then
        cboSelect_COMRCDCLI.AddItem rsSab_X("COMRCDNAT") & " " & rsSab_X("COMRCDCLI") & " " & rsSab_X("COMRCDOPE") & " " & rsSab_X("COMRCDSER") & " " & rsSab_X("COMRCDSSE")
    End If
    rsSab_X.MoveNext
Loop
'If cboSelect_COMRCDCLI.ListCount > 1 Then
'    cboSelect_COMRCDCLI.ListIndex = 1
'Else
    cboSelect_COMRCDCLI.ListIndex = 0
'End If
blnControl = True

    '
txtRTF.LoadFile paramServer("\\BiaDoc\Filigrane\VB_RTF_Modèle.rtf")
VB_RTF_Modèle = txtRTF.TextRTF

txtRTF.Visible = False


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



Private Sub fgSelect_Display_1()

Dim K As Long

On Error GoTo Error_Handler
currentAction = "fgSelect_Display_1"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = fgSelect_FormatString

fgSelect.Rows = 1
fgSelect.Col = 2: fgSelect.CellAlignment = 1
fgSelect.Col = 4: fgSelect.CellAlignment = 1
                 
fgSelect.Row = 0
 mCOMRCDMTR = 0: mCOMRCDMTD = 0
Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_Display_1_Line
    mCOMRCDMTR = mCOMRCDMTR + rsSab("COMRCDMTR")
    If rsSab("COMRCDSTA") = " " Or rsSab("COMRCDSTA") = "M" Then mCOMRCDMTD = mCOMRCDMTD + rsSab("COMRCDMTD")
    rsSab.MoveNext

Loop
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
fgSelect.Col = 0: fgSelect.Text = "Total"

fgSelect.Col = 4: fgSelect.Text = Format(mCOMRCDMTR, "### ### ##0.00") & "  "
If mCOMRCDMTR < 0 Then
    fgSelect.CellForeColor = vbRed
Else
    fgSelect.CellForeColor = vbBlue
End If
fgSelect.Col = 2: fgSelect.Text = Format(-mCOMRCDMTD, "### ### ##0.00") & "  "
If mCOMRCDMTD > 0 Then
    fgSelect.CellForeColor = vbRed
Else
    fgSelect.CellForeColor = vbBlue
End If
    For K = 0 To 8
        fgSelect.Col = K
        fgSelect.CellBackColor = mColor_G1
    Next K
fgSelect.Visible = True

'If fgSelect.Rows = 2 Then
'    fgSelect.Col = 0
'    Call fgDetail_Display(Trim(fgSelect.Text))
'End If

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgSelect_Display_3()

Dim K As Long, I As Long, mDEV As String

On Error GoTo Error_Handler
currentAction = "fgSelect_Display_3"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = "<Racine           |<Intitulé                                                                   |<Service Opération   |<Code COM       |> Montant min |>Montant rétrocédé|< date de début|<Statut                                                                                       |<Id               |"
fgSelect.Rows = 1
                 
fgSelect.Row = 0
fgSelect.Col = 6: fgSelect.CellAlignment = 1
fgSelect.Col = 4: fgSelect.CellAlignment = 1
fgSelect.Col = 5: fgSelect.CellAlignment = 1
mCOMRCDMTR = 0
Do While Not rsSab.EOF
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_Display_3_Line
    rsSab.MoveNext

Loop


fgSelect.Visible = True
If fgSelect.Rows = 1 Then
    xYCOMRCD0.COMRCDPIE = 0: xYCOMRCD0.COMRCDECR = 0
    Call fraCOMRCDCLI_Display
End If


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Private Sub fgSelect_Display_6RDC()

Dim K As Long, I As Long, mDEV As String

On Error GoTo Error_Handler
currentAction = "fgSelect_Display_6"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = "<Racine           |<Devise   |<Code commission |<Libellé                                                                       |> Nb      |> Montant com (Devise) |>Montant rétrocédé EUR||"
fgSelect.Rows = 1
                 
fgSelect.Row = 0
fgSelect.Col = 6: fgSelect.CellAlignment = 1
fgSelect.Col = 4: fgSelect.CellAlignment = 1
fgSelect.Col = 5: fgSelect.CellAlignment = 1
mCOMRCDMTR = 0
Do While Not rsSab.EOF
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_Display_6RDC_Line
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


Private Sub fgSelect_Display_6DC()

Dim K As Long, I As Long, mDEV As String

On Error GoTo Error_Handler
currentAction = "fgSelect_Display_6DC"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = "<Devise   |<Code commission   |<Libellé                                                                       |> Nb      |> Montant com (Devise) |>Montant rétrocédé EUR||"
fgSelect.Rows = 1
                 
fgSelect.Row = 0
fgSelect.Col = 5: fgSelect.CellAlignment = 1
fgSelect.Col = 3: fgSelect.CellAlignment = 1
fgSelect.Col = 4: fgSelect.CellAlignment = 1
mCOMRCDMTR = 0
Do While Not rsSab.EOF
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_Display_6DC_Line
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



Private Sub fgSelect_Display_2()

Dim K As Long, I As Long, mDEV As String

On Error GoTo Error_Handler
currentAction = "fgSelect_Display_2"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.FormatString = "<Date traitement       |<Opération                  |>Mt rétrocession              |> Relevé N°|<Vos références                                          |<Autres références                                        ||"
fgSelect.Rows = 1
                 
fgSelect.Row = 0
fgSelect.Col = 2: fgSelect.CellAlignment = 1
fgSelect.Col = 3: fgSelect.CellAlignment = 1
mCOMRCDMTR = 0
Do While Not rsSab.EOF
    If rsSab("COMRCDSTA") = " " Or rsSab("COMRCDSTA") = "M" Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_Display_2_Line
        mCOMRCDMTR = mCOMRCDMTR + rsSab("COMRCDMTR")
    End If
    rsSab.MoveNext

Loop

fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
fgSelect.Col = 0: fgSelect.Text = "Total"

fgSelect.Col = 2: fgSelect.Text = Format(mCOMRCDMTR, "### ### ##0.00") & "  "
If mCOMRCDMTR < 0 Then
    fgSelect.CellForeColor = vbRed
Else
    fgSelect.CellForeColor = vbBlue
End If
    For K = 0 To 8
        fgSelect.Col = K
        fgSelect.CellBackColor = mColor_G1
    Next K

fgSelect.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub fgSelect_Display_1_Line()
Dim K As Integer, wColor As Long

On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = dateImp10_S(rsSab("COMRCDDTR") + 19000000)
fgSelect.Col = 1: fgSelect.Text = rsSab("COMRCDSER") & " " & rsSab("COMRCDSSE") & " " & rsSab("COMRCDOPE") & " " & rsSab("COMRCDNUM")
If rsSab("COMRCDMTD") <> 0 Then
    fgSelect.Col = 2: fgSelect.Text = Format(-rsSab("COMRCDMTD"), "### ### ##0.00") & "  "
    If rsSab("COMRCDMTD") > 0 Then
        fgSelect.CellForeColor = vbRed
    Else
        fgSelect.CellForeColor = vbBlue
    End If
End If

fgSelect.Col = 3: fgSelect.Text = " " & rsSab("COMRCDDEV"): fgSelect.CellFontBold = True
If rsSab("COMRCDMTR") <> 0 Then
    fgSelect.Col = 4: fgSelect.Text = Format(rsSab("COMRCDMTR"), "### ### ##0.00") & "  "
    If rsSab("COMRCDMTR") < 0 Then
        fgSelect.CellForeColor = vbRed
    Else
        fgSelect.CellForeColor = vbBlue
    End If
End If
fgSelect.Col = 6: fgSelect.Text = rsSab("COMRCDZCOM")
fgSelect.Col = 7: fgSelect.Text = rsSab("COMRCDPIE") & " - " & rsSab("COMRCDECR")

fgSelect.Col = 8: fgSelect.Text = rsSab("COMRCDSTA") & " " & rsSab("COMRCDYUSR") & " - " & dateImp10_S(rsSab("COMRCDYAMJ")) & " " & timeImp8(rsSab("COMRCDYHMS")) & " - " & rsSab("COMRCDYVER")


If rsSab("COMRCDSTA") <> " " Then
    Select Case rsSab("COMRCDSTA")
        Case "M": wColor = mColor_Y2
        Case "I": wColor = RGB(230, 230, 230)
        Case "A": wColor = mColor_W1
        Case "Z": wColor = RGB(230, 230, 230)
    End Select
    For K = 0 To 8
        fgSelect.Col = K
        fgSelect.CellBackColor = wColor
    Next K
End If
fgSelect.Col = 5
If rsSab("COMRCDRLV") <> 0 Then
    fgSelect.Text = rsSab("COMRCDRLV")
    fgSelect.CellBackColor = mColor_Y1
Else
    If rsSab("COMRCDSTA") = " " Then fgSelect.CellBackColor = mColor_G1
End If


End Sub

Public Sub fgSelect_Display_2_Line()
Dim K As Integer, wColor As Long, xSQL As String

On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = dateImp10_S(rsSab("COMRCDDTR") + 19000000)
fgSelect.Col = 1: fgSelect.Text = rsSab("COMRCDSER") & " " & rsSab("COMRCDSSE") & " " & rsSab("COMRCDOPE") & " " & rsSab("COMRCDNUM")
'If rsSab("COMRCDMTD") <> 0 Then
'    fgSelect.Col = 2: fgSelect.Text = Format(-rsSab("COMRCDMTD"), "### ### ##0.00") & "  "
'    If rsSab("COMRCDMTD") > 0 Then
'        fgSelect.CellForeColor = vbRed
'    Else
'        fgSelect.CellForeColor = vbBlue
'    End If
    
'End If

'fgSelect.Col = 3: fgSelect.Text = " " & rsSab("COMRCDDEV"): fgSelect.CellFontBold = True
If rsSab("COMRCDMTR") <> 0 Then
    fgSelect.Col = 2: fgSelect.Text = Format(rsSab("COMRCDMTR"), "### ### ##0.00") & "  "
    If rsSab("COMRCDMTR") < 0 Then
        fgSelect.CellForeColor = vbRed
    Else
        fgSelect.CellForeColor = vbBlue
    End If
End If
fgSelect.Col = 3
If rsSab("COMRCDRLV") <> 0 Then
    fgSelect.Text = rsSab("COMRCDRLV")
    fgSelect.CellBackColor = mColor_Y1
Else
    fgSelect.CellBackColor = mColor_G1
End If

xSQL = "select CHGMESVOS , CHGMESNOS from " & paramIBM_Library_SAB & ".ZCHGMES0 " _
     & " where CHGMESETA = 1 and CHGMESAGE = 1 and CHGMESSER = '" & rsSab("COMRCDSER") & "' and CHGMESSSE = '" & rsSab("COMRCDSSE") & "'" _
     & " and CHGMESOPE = '" & rsSab("COMRCDOPE") & "' and CHGMESDOS = " & rsSab("COMRCDNUM") & " and CHGMESSEQ = '  '"

Set rsSab_X = cnsab.Execute(xSQL)

If Not rsSab_X.EOF Then
    fgSelect.Col = 4: fgSelect.Text = rsSab_X("CHGMESVOS")
    fgSelect.Col = 5: fgSelect.Text = rsSab_X("CHGMESNOS")
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
    Case Else: blnAuto = False:

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


Private Sub cboSelect_COMRCDCLI_Click()
cmdSelect_Clear
End Sub

Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

End Sub


Private Sub cmdCOMRCDCLI_Add_Click()
Dim V, xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Saisie d'un code commission"): DoEvents
blnCOMRCDZCOM_Add = True


If Not cmdCOMRCDCLI_Add.BackColor = vbGreen Then
    fraCOMRCDCLI_1.Caption = ""
    fraCOMRCDCLI_1.Visible = True
    cmdCOMRCDCLI_New.Visible = False
    cmdCOMRCDCLI_Add.Caption = "Enregistrer un code commission"
    cmdCOMRCDCLI_Add.BackColor = vbGreen
Else

    If IsNull(fraCOMRCDCLI_Control) Then
        currentAction = "cmdCOMRCDCLI_Add_Click"
        xSQL = "select COMRCDECR from " & paramIBM_Library_SABSPE & ".YCOMRCD0" _
             & " where COMRCDnat = '$' and COMRCDPIE = " & newYCOMRCD0.COMRCDPIE _
             & "  order by COMRCDECR desc"
        Set rsSab = cnsab.Execute(xSQL)
        
        If rsSab.EOF Then
            newYCOMRCD0.COMRCDECR = 1
        Else
            newYCOMRCD0.COMRCDECR = rsSab("COMRCDECR") + 1
        End If
        
        newYCOMRCD0.COMRCDSTA = ""
        newYCOMRCD0.COMRCDRLV = 0
        
        'V = cnSAB_Transaction("BeginTrans")
        cnSab_Update.Open paramODBC_DSN_SAB
        V = sqlYCOMRCD0_Insert(newYCOMRCD0)
        
        cnSab_Update.Close
        'V = cnSAB_Transaction("Commit")
        If Not IsNull(V) Then
            Call MsgBox(V, vbCritical, Me.Name & " : " & currentAction)
        Else
            fraCOMRCDCLI.Visible = False
            Call cmdSelect_SQL_3
        End If
    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdCOMRCDCLI_Delete_Click()
Dim xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> COM_Retro cmdCOMRCDCLI_Delete ........"): DoEvents
    newYCOMRCD0 = oldYCOMRCD0
If Trim(newYCOMRCD0.COMRCDSTA) = "" Then
    newYCOMRCD0.COMRCDSTA = "I"
Else
    newYCOMRCD0.COMRCDSTA = " "
End If

'V = cnSAB_Transaction("BeginTrans")
cnSab_Update.Open paramODBC_DSN_SAB
V = sqlYCOMRCD0_Update(newYCOMRCD0, oldYCOMRCD0)
cnSab_Update.Close
'V = cnSAB_Transaction("Commit")
If Not IsNull(V) Then
    Call MsgBox(V, vbCritical, Me.Name & " : " & currentAction)
Else
    fraCOMRCDCLI.Visible = False
    Call cmdSelect_SQL_3
End If

Call lstErr_AddItem(lstErr, cmdContext, "< COM_Retro cmdCOMRCDCLI_Delete terminé"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdCOMRCDCLI_New_Click()
Dim V, xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Saisie d'un client"): DoEvents

If Not cmdCOMRCDCLI_New.BackColor = vbGreen Then
    fraCOMRCDCLI_0.Enabled = True
    fraCOMRCDCLI_1.Visible = False
    fraCOMRCDCLI.Caption = ""
    cmdCOMRCDCLI_Add.Visible = False
    cmdCOMRCDCLI_New.Caption = "Enregistrer un nouveau client"
    cmdCOMRCDCLI_New.BackColor = vbGreen
Else
    If IsNull(fraCOMRCDCLI_0_Control) Then
        currentAction = "cmdCOMRCDCLI_new_Click"
        xSQL = "select COMRCDPIE from " & paramIBM_Library_SABSPE & ".YCOMRCD0" _
             & " where COMRCDnat = '$' and COMRCDECR = 0" _
             & "  order by COMRCDPIE desc"
        Set rsSab = cnsab.Execute(xSQL)
        
        If rsSab.EOF Then
            newYCOMRCD0.COMRCDPIE = 1
        Else
            newYCOMRCD0.COMRCDPIE = rsSab("COMRCDPIE") + 1
        End If
        
        newYCOMRCD0.COMRCDSTA = ""
        newYCOMRCD0.COMRCDRLV = 0
        
        'V = cnSAB_Transaction("BeginTrans")
        cnSab_Update.Open paramODBC_DSN_SAB
        V = sqlYCOMRCD0_Insert(newYCOMRCD0)
        
        cnSab_Update.Close
        'V = cnSAB_Transaction("Commit")
        If Not IsNull(V) Then
            Call MsgBox(V, vbCritical, Me.Name & " : " & currentAction)
        Else
            fraCOMRCDCLI.Visible = False
            Call cmdSelect_SQL_3
        End If
        
    End If
End If

Me.Enabled = True: Me.MousePointer = 0



End Sub

Private Sub cmdCOMRCDCLI_Quit_Click()
fraCOMRCDCLI.Visible = False
End Sub

Private Sub cmdCOMRCDCLI_Update_Click()
Dim xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> COM_Retro cmdCOMRCDCLI_Update ........"): DoEvents
If IsNull(fraCOMRCDCLI_Control) Then
    currentAction = "cmdCOMRCDCLI_Add_Click"

    'V = cnSAB_Transaction("BeginTrans")
    cnSab_Update.Open paramODBC_DSN_SAB
    V = sqlYCOMRCD0_Update(newYCOMRCD0, oldYCOMRCD0)
    cnSab_Update.Close
    'V = cnSAB_Transaction("Commit")
    If Not IsNull(V) Then
        Call MsgBox(V, vbCritical, Me.Name & " : " & currentAction)
    Else
        fraCOMRCDCLI.Visible = False
        Call cmdSelect_SQL_3
    End If

    Call lstErr_AddItem(lstErr, cmdContext, "< COM_Retro cmdCOMRCDCLI_Update terminé"): DoEvents
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdCOMRCDRLV_Ok_Click()
Dim xWhere As String, xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> cmdCOMRCDRLV_Ok........"): DoEvents

xWhere = " where COMRCDNAT = '$' and COMRCDCLI = '" & mCOMRCDCLI & "'" _
     & " and COMRCDOPE = '" & mCOMRCDOPE & "' and COMRCDSER = '" & mCOMRCDSER & "'  and COMRCDSSE = '" & mCOMRCDSSE & "'" _
     & " and COMRCDECR = 0"

xSQL = "select COMRCDRLV from " & paramIBM_Library_SABSPE & ".YCOMRCD0" & xWhere
Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    Call MsgBox(xSQL, vbCritical, "cmdCOMRCDRLV_Ok_Click :  erreur de lecture")
Else
    mCOMRCDRLV = rsSab("COMRCDRLV") + 1
    'V = cnSAB_Transaction("BeginTrans")
    cnSab_Update.Open paramODBC_DSN_SAB
    xSQL = "Update " & paramIBM_Library_SABSPE & ".YCOMRCD0 set COMRCDRLV = " & mCOMRCDRLV & xWhere
    Call sqlYCOMRCD0_Update_CMD(xSQL)
    xSQL = "Update " & paramIBM_Library_SABSPE & ".YCOMRCD0 set COMRCDRLV = " & mCOMRCDRLV & mCOMRCDRLV_Where
    Call sqlYCOMRCD0_Update_CMD(xSQL)
    cnSab_Update.Close
    'V = cnSAB_Transaction("Commit")
    Call cmdSelect_Clear
    txtSelect_COMRCDRLV = mCOMRCDRLV
    Call cmdSelect_SQL_2
    
    Call mnuPrint_Excel_Click
End If
cmdCOMRCDRLV_Ok.Visible = False
    
Call lstErr_AddItem(lstErr, cmdContext, "< cmdCOMRCDRLV_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

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

Private Sub cmdUpdate_Ok_Click()
Dim xSQL As String, xSet As String, X As String
On Error GoTo Error_Handler

GoTo Exit_sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, K As Integer
On Error Resume Next
txtRTF.Visible = False


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
        Select Case cmdSelect_SQL_K
            Case "1"
                
            Case "2"
 '               fgDetail.Col = 1: wX = Trim(fgDetail.Text)
        End Select
    End If
End If
fgDetail.LeftCol = 0



End Sub




Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, K As Integer, xSQL As String
On Error Resume Next

txtRTF.Visible = False
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
                If arrHab(2) And mCOMRCDNAT = "$" Then
                   fgSelect.Col = 7: wX = Trim(fgSelect.Text)
                   K = InStr(wX, "-")
                   If K > 0 Then
                        xYCOMRCD0.COMRCDPIE = Val(Mid$(wX, 1, K - 1))
                        xYCOMRCD0.COMRCDECR = Val(Mid$(wX, K + 1, Len(wX) - K))
                   Else
                        Call MsgBox("erreur n° de pièce", vbCritical, "COM_Retro : fgSelect_MouseDown")
                        Exit Sub
                    End If
                    
                   fgSelect.Col = 5
                    If Trim(fgSelect.Text) <> "" Then
                        fgSelect.Col = 1: wX = Trim(fgSelect.Text)
                        Call frmSAB_Dossier_DB.Form_Init("MOUVEMDTR", "", "", "", Mid$(wX, 1, 2), Mid$(wX, 4, 2), Mid$(wX, 7, 3), Val(Mid$(wX, 11, 9)))
                    Else
                        fgSelect.Col = 8
                         If Mid$(fgSelect.Text, 1, 1) = " " Then
                             mnuUpdate_Restaurer.Visible = False
                             mnuUpdate_Ignorer.Visible = True
                             mnuUpdate_MTR.Visible = True
                         Else
                             mnuUpdate_Restaurer.Visible = True
                             mnuUpdate_Ignorer.Visible = False
                             mnuUpdate_MTR.Visible = False
                         End If
                         Me.PopupMenu mnuUpdate, vbPopupMenuLeftButton
                    End If
                Else
                    fgSelect.Col = 1: wX = Trim(fgSelect.Text)
                    Call frmSAB_Dossier_DB.Form_Init("MOUVEMDTR", "", "", "", Mid$(wX, 1, 2), Mid$(wX, 4, 2), Mid$(wX, 7, 3), Val(Mid$(wX, 11, 9)))
                End If
            Case "2"
                fgSelect.Col = 1: wX = Trim(fgSelect.Text)
                Call frmSAB_Dossier_DB.Form_Init("MOUVEMDTR", "", "", "", Mid$(wX, 1, 2), Mid$(wX, 4, 2), Mid$(wX, 7, 3), Val(Mid$(wX, 11, 9)))
             Case "3"
                If arrHab(18) Then
                   fgSelect.Col = 8: wX = Trim(fgSelect.Text)
                   K = InStr(wX, "-")
                   If K > 0 Then
                        xYCOMRCD0.COMRCDPIE = Val(Mid$(wX, 1, K - 1))
                        xYCOMRCD0.COMRCDECR = Val(Mid$(wX, K + 1, Len(wX) - K))
                        Call fraCOMRCDCLI_Display
                   Else
                        Call MsgBox("erreur Id", vbCritical, "COM_Retro : fgSelect_MouseDown")
                        Exit Sub
                    End If
                    
                End If
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
fgDetail.Visible = False
txtRTF.Visible = False
cmdSelect_Ok.BackColor = vbGreen
libCOMRCDCLI = ""
    cmdCOMRCDRLV_Ok.Visible = False
If Not IsNull(txtSelect_COMRCDDTR_Min.Value) Then
    txtSelect_COMRCDDTR_Max.Visible = True
Else
    txtSelect_COMRCDDTR_Max.Visible = False
End If
fraCOMRCDCLI.Visible = False
End Sub

Private Sub mnuUpdate_Display_Click()
Dim wX As String
fgSelect.Col = 1: wX = Trim(fgSelect.Text)
Call frmSAB_Dossier_DB.Form_Init("MOUVEMDTR", "", "", "", Mid$(wX, 1, 2), Mid$(wX, 4, 2), Mid$(wX, 7, 3), Val(Mid$(wX, 11, 9)))

End Sub

Private Sub mnuUpdate_Ignorer_Click()
Dim xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> COM_Retro Ignorer cette écriture ........"): DoEvents
    xSQL = "Select * from " & paramIBM_Library_SABSPE & ".YCOMRCD0" _
         & " where COMRCDNAT = ' ' and COMRCDPIE = " & xYCOMRCD0.COMRCDPIE & "  and COMRCDECR = " & xYCOMRCD0.COMRCDECR
    Set rsSab = cnsab.Execute(xSQL)
    If rsSab.EOF Then
        Call MsgBox("Erreur lecture YCOMRCD0", vbCritical, "COM_RETRO : mnuUpdate_Ignorer")
    Else
        Call rsYCOMRCD0_GetBuffer(rsSab, oldYCOMRCD0)
        newYCOMRCD0 = oldYCOMRCD0
        newYCOMRCD0.COMRCDSTA = "I"
        'V = cnSAB_Transaction("BeginTrans")
        cnSab_Update.Open paramODBC_DSN_SAB
        Call sqlYCOMRCD0_Update(newYCOMRCD0, oldYCOMRCD0)
        cnSab_Update.Close
        'V = cnSAB_Transaction("Commit")
    End If
Call lstErr_AddItem(lstErr, cmdContext, "< COM_Retro Ignorer cette écriture terminé"): DoEvents
Call cmdSelect_SQL_1
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuUpdate_MTR_Click()
Dim X As String, xCur As Currency
fgSelect.Col = 4:
If Trim(fgSelect.Text) = "" Then
    xCur = 0
Else
    xCur = CCur(fgSelect.Text)
End If
X = InputBox("Montant rétrocédé pr ex 12,34 (virgule)" & vbCr & "tapez 0 pour effacer le montant", " modification du montant rétrocédé", xCur)
If Trim(X) = "" Then
    Exit Sub
Else
    If Not IsNumeric(X) Then
        Call MsgBox("saisir un montant sous la forme 12,34 (virgule)", vbCritical, "COM_RETRO montant rétrocédé")
    Else
        xCur = CCur(X)
        Me.Enabled = False: Me.MousePointer = vbHourglass
        Call lstErr_Clear(lstErr, cmdContext, "> COM_Retro modification ........"): DoEvents
            X = "Select * from " & paramIBM_Library_SABSPE & ".YCOMRCD0" _
                 & " where COMRCDNAT = ' ' and COMRCDPIE = " & xYCOMRCD0.COMRCDPIE & "  and COMRCDECR = " & xYCOMRCD0.COMRCDECR
            Set rsSab = cnsab.Execute(X)
            If rsSab.EOF Then
                Call MsgBox("Erreur lecture YCOMRCD0", vbCritical, "COM_RETRO : mnuUpdate_modification ")
            Else
                Call rsYCOMRCD0_GetBuffer(rsSab, oldYCOMRCD0)
                newYCOMRCD0 = oldYCOMRCD0
                newYCOMRCD0.COMRCDMTR = xCur
                newYCOMRCD0.COMRCDSTA = "M"
                'V = cnSAB_Transaction("BeginTrans")
                cnSab_Update.Open paramODBC_DSN_SAB
                Call sqlYCOMRCD0_Update(newYCOMRCD0, oldYCOMRCD0)
                cnSab_Update.Close
                'V = cnSAB_Transaction("Commit")
            End If
        Call lstErr_AddItem(lstErr, cmdContext, "< COM_Retro modification  terminée"): DoEvents
        Call cmdSelect_SQL_1
        Me.Enabled = True: Me.MousePointer = 0

    End If
End If
End Sub

Private Sub mnuUpdate_Restaurer_Click()
Dim xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> COM_Retro Restaurer cette écriture ........"): DoEvents
    xSQL = "Select * from " & paramIBM_Library_SABSPE & ".YCOMRCD0" _
         & " where COMRCDNAT = ' ' and COMRCDPIE = " & xYCOMRCD0.COMRCDPIE & "  and COMRCDECR = " & xYCOMRCD0.COMRCDECR
    Set rsSab = cnsab.Execute(xSQL)
    If rsSab.EOF Then
        Call MsgBox("Erreur lecture YCOMRCD0", vbCritical, "COM_RETRO : mnuUpdate_Ignorer")
    Else
        Call rsYCOMRCD0_GetBuffer(rsSab, oldYCOMRCD0)
        newYCOMRCD0 = oldYCOMRCD0
        newYCOMRCD0.COMRCDSTA = " "
        'V = cnSAB_Transaction("BeginTrans")
        cnSab_Update.Open paramODBC_DSN_SAB
        Call sqlYCOMRCD0_Update(newYCOMRCD0, oldYCOMRCD0)
        cnSab_Update.Close
        'V = cnSAB_Transaction("Commit")
    End If
Call lstErr_AddItem(lstErr, cmdContext, "< COM_Retro Restaurer cette écriture terminé"): DoEvents
Call cmdSelect_SQL_1
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub txtCOMRCDCLI_GotFocus()
txtCOMRCDCLI.BackColor = focusUsr.BackColor

End Sub

Private Sub txtCOMRCDCLI_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub


Private Sub txtCOMRCDCLI_LostFocus()
txtCOMRCDCLI.BackColor = txtUsr.BackColor

End Sub


Private Sub txtCOMRCDDTR_GotFocus()
'txtCOMRCDDTR.BackColor = focusUsr.BackColor

End Sub


Private Sub txtCOMRCDDTR_LostFocus()
'txtCOMRCDDTR.BackColor = txtUsr.BackColor

End Sub


Private Sub txtCOMRCDMTD_GotFocus()
txtCOMRCDMTD.BackColor = focusUsr.BackColor
End Sub


Private Sub txtCOMRCDMTD_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtCOMRCDMTD)
End Sub


Private Sub txtCOMRCDMTD_LostFocus()
txtCOMRCDMTD.BackColor = txtUsr.BackColor

End Sub


Private Sub txtCOMRCDMTR_GotFocus()
txtCOMRCDMTR.BackColor = focusUsr.BackColor

End Sub


Private Sub txtCOMRCDMTR_KeyPress(KeyAscii As Integer)
Call num_Montant(KeyAscii, txtCOMRCDMTR)

End Sub


Private Sub txtCOMRCDMTR_LostFocus()
txtCOMRCDMTR.BackColor = txtUsr.BackColor

End Sub


Private Sub txtCOMRCDOPE_GotFocus()
txtCOMRCDOPE.BackColor = focusUsr.BackColor

End Sub


Private Sub txtCOMRCDOPE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtCOMRCDOPE_LostFocus()
txtCOMRCDOPE.BackColor = txtUsr.BackColor

End Sub


Private Sub txtCOMRCDSER_GotFocus()
txtCOMRCDSER.BackColor = focusUsr.BackColor

End Sub


Private Sub txtCOMRCDSER_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtCOMRCDSER_LostFocus()
txtCOMRCDSER.BackColor = txtUsr.BackColor

End Sub


Private Sub txtCOMRCDSSE_GotFocus()
txtCOMRCDSSE.BackColor = focusUsr.BackColor

End Sub


Private Sub txtCOMRCDSSE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtCOMRCDSSE_LostFocus()
txtCOMRCDSSE.BackColor = txtUsr.BackColor

End Sub


Private Sub txtCOMRCDZCOM_GotFocus()
txtCOMRCDZCOM.BackColor = focusUsr.BackColor

End Sub


Private Sub txtCOMRCDZCOM_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtCOMRCDZCOM_LostFocus()
txtCOMRCDZCOM.BackColor = txtUsr.BackColor

End Sub


Private Sub txtSelect_COMRCDDTR_Max_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_COMRCDDTR_Min_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_COMRCDRLV_Change()
cmdSelect_Clear

End Sub


Private Sub txtSelect_COMRCDRLV_GotFocus()
Call txt_GotFocus(txtSelect_COMRCDRLV)

End Sub


Private Sub txtSelect_COMRCDRLV_KeyPress(KeyAscii As Integer)
If KeyAscii <> 45 Then KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtSelect_COMRCDRLV_LostFocus()
Call txt_LostFocus(txtSelect_COMRCDRLV)
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
    cmdCOMRCDRLV_Ok.Visible = False
  '   txtSelect_COMRCDRLV.Visible = False
    Select Case cmdSelect_SQL_K
        Case "1": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
        Case "2": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True ': txtSelect_COMRCDRLV.Visible = True
        Case "3": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
        Case "6 DC", "6 RDC": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
    End Select

End If
End Sub


Private Sub cmdSelect_SQL_3()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_3"
xWhere = ""
mCOMRCDNAT = Trim(Mid$(cboSelect_COMRCDCLI, 1, 1))
mCOMRCDCLI = Trim(Mid$(cboSelect_COMRCDCLI, 3, 7))
mCOMRCDOPE = Trim(Mid$(cboSelect_COMRCDCLI, 11, 3))
mCOMRCDSER = Trim(Mid$(cboSelect_COMRCDCLI, 15, 2))
mCOMRCDSSE = Trim(Mid$(cboSelect_COMRCDCLI, 18, 2))
If mCOMRCDCLI <> "" Then
    xSQL = "select CLIENARA1 , CLIENARA2 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & mCOMRCDCLI & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        libCOMRCDCLI = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
    Else
        libCOMRCDCLI = "???"
    End If
    
    xWhere = " where COMRCDnat = '$' and COMRCDCLI = '" & mCOMRCDCLI & "'"
    mExportation = "Client  " & mCOMRCDCLI & " -" & libCOMRCDCLI & " - Paramétrage "
Else
    xWhere = " where COMRCDnat = '$' "
    mExportation = " - Paramétrage "
End If

    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCOMRCD0, " & paramIBM_Library_SAB & ".ZCLIENA0 " _
         & xWhere & "and CLIENAETB = 1 and CLIENACLI = COMRCDCLI order by COMRCDCLI , COMRCDPIE , COMRCDECR"
    Set rsSab = cnsab.Execute(xSQL)
      
    Call fgSelect_Display_3


Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Private Sub cmdSelect_SQL_6()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_6"
xWhere = ""
mCOMRCDNAT = Trim(Mid$(cboSelect_COMRCDCLI, 1, 1))
mCOMRCDCLI = Trim(Mid$(cboSelect_COMRCDCLI, 3, 7))
mCOMRCDOPE = Trim(Mid$(cboSelect_COMRCDCLI, 11, 3))
mCOMRCDSER = Trim(Mid$(cboSelect_COMRCDCLI, 15, 2))
mCOMRCDSSE = Trim(Mid$(cboSelect_COMRCDCLI, 18, 2))
If mCOMRCDCLI <> "" Then
    xSQL = "select CLIENARA1 , CLIENARA2 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & mCOMRCDCLI & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        libCOMRCDCLI = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
    Else
        libCOMRCDCLI = "???"
    End If
    
    xWhere = " where COMRCDnat = ' ' and COMRCDCLI = '" & mCOMRCDCLI & "'"
    mExportation = "Client  " & mCOMRCDCLI & " -" & libCOMRCDCLI & " - Statistique "
Else
    xWhere = " where COMRCDnat = ' ' "
    mExportation = " - Statistique "
End If

If Not IsNull(txtSelect_COMRCDDTR_Min.Value) Then
    Call DTPicker_Control(txtSelect_COMRCDDTR_Min, wAmjMin)
    Call DTPicker_Control(txtSelect_COMRCDDTR_Max, wAmjMax)
    xWhere = xWhere & " and COMRCDDTR >= " & wAmjMin - 19000000 & " And COMRCDDTR <= " & wAmjMax - 19000000
     mExportation = mExportation & "du " & dateImp10_S(wAmjMin) & "au " & dateImp10_S(wAmjMax)
Else
    V = "Préciser la période"
    GoTo Error_MsgBox
End If

Select Case cmdSelect_SQL_K
    Case "6 RDC"
          xSQL = "select comrcdcli , comrcddev , comrcdzcom  , count(*) , sum(comrcdmtd) , sum(comrcdmtr) from " & paramIBM_Library_SABSPE & ".YCOMRCD0" _
               & xWhere _
               & " group by comrcdcli , comrcddev , comrcdzcom" _
               & " order by comrcdcli , comrcddev , comrcdzcom"
        Set rsSab = cnsab.Execute(xSQL)
          
        Call fgSelect_Display_6RDC
         xSQL = "select  comrcdzcom  , count(*)  , sum(comrcdmtr) from " & paramIBM_Library_SABSPE & ".YCOMRCD0" _
               & xWhere _
               & " group by  comrcdzcom" _
               & " order by comrcdzcom"
        Set rsSab = cnsab.Execute(xSQL)
          
        Call fgSelect_Display_6RDC_Total
    Case "6 DC"
          xSQL = "select   comrcddev , comrcdzcom  , count(*) , sum(comrcdmtd) , sum(comrcdmtr)  from " & paramIBM_Library_SABSPE & ".YCOMRCD0 " _
               & xWhere _
               & " group by  comrcddev , comrcdzcom" _
               & " order by  comrcddev , comrcdzcom"
        Set rsSab = cnsab.Execute(xSQL)
          
        Call fgSelect_Display_6DC
End Select


Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub


Public Sub fgSelect_Display_3_Line()
Dim K As Integer
Dim wColor As Long
Dim X As String
On Error Resume Next

fgSelect.Col = 8: fgSelect.Text = rsSab("COMRCDPIE") & " - " & rsSab("COMRCDECR")

fgSelect.Col = 0: fgSelect.Text = rsSab("COMRCDCLI"): fgSelect.CellFontBold = True
fgSelect.Col = 1: fgSelect.Text = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))

fgSelect.Col = 3: fgSelect.Text = rsSab("COMRCDZCOM"): fgSelect.CellFontBold = True: fgSelect.CellForeColor = vbBlue

fgSelect.Col = 2: fgSelect.Text = rsSab("COMRCDSER") & " " & rsSab("COMRCDSSE") & " " & rsSab("COMRCDOPE")
If rsSab("COMRCDMTD") <> 0 Then
    fgSelect.CellForeColor = vbBlue: fgSelect.CellFontBold = True
    fgSelect.Col = 4
    fgSelect.Text = Format(rsSab("COMRCDMTD"), "### ### ##0.00") & "  "
    fgSelect.CellForeColor = vbBlue
End If
If rsSab("COMRCDMTR") <> 0 Then
    fgSelect.Col = 5
    fgSelect.Text = Format(rsSab("COMRCDMTR"), "### ### ##0.00") & "  "
    fgSelect.CellForeColor = vbRed
End If
If rsSab("COMRCDDTR") <> 0 Then fgSelect.Col = 6: fgSelect.Text = dateImp10_S(rsSab("COMRCDDTR") + 19000000)

fgSelect.Col = 7: fgSelect.Text = rsSab("COMRCDSTA") & " " & rsSab("COMRCDYUSR") & " - " & dateImp10_S(rsSab("COMRCDYAMJ")) & " " & timeImp8(rsSab("COMRCDYHMS")) & " - " & rsSab("COMRCDYVER")


If rsSab("COMRCDSTA") <> " " Then
    Select Case rsSab("COMRCDSTA")
        Case "M": wColor = mColor_Y2
        Case "I": wColor = RGB(230, 230, 230)
        Case "A": wColor = mColor_W1
        Case "Z": wColor = RGB(230, 230, 230)
    End Select
    For K = 0 To 8
        fgSelect.Col = K
        fgSelect.CellBackColor = wColor
    Next K
End If

End Sub
Public Sub fgSelect_Display_6RDC_Line()
Dim K As Integer
Dim wColor As Long
Dim X As String
On Error Resume Next


fgSelect.Col = 0: fgSelect.Text = rsSab(0): fgSelect.CellFontBold = True
fgSelect.Col = 1: fgSelect.Text = rsSab(1)
fgSelect.Col = 2: fgSelect.Text = rsSab(2)

fgSelect.Col = 4: fgSelect.Text = Format(rsSab(3), "### ### ##0") & "  "

If rsSab(4) <> 0 Then
    fgSelect.Col = 5
    fgSelect.Text = Format(-rsSab(4), "### ### ##0.00") & "  "
    If rsSab(4) > 0 Then
        fgSelect.CellForeColor = vbRed
    Else
        fgSelect.CellForeColor = vbBlue
    End If
End If
If rsSab(5) <> 0 Then
    fgSelect.Col = 6
    fgSelect.Text = Format(rsSab(5), "### ### ##0.00") & "  "
    fgSelect.CellForeColor = vbRed
End If


X = "select BASTABDON from " & paramIBM_Library_SAB & ".ZBASTAB0" _
     & " where BASTABETA = 1  and BASTABNUM = 44 and BASTABARG = '" & rsSab(2) & "'"
Set rsSab_X = cnsab.Execute(X)

If Not rsSab_X.EOF Then
    fgSelect.Col = 3: fgSelect.Text = Mid$(rsSab_X("BASTABDON"), 1, 30)
End If

If Trim(rsSab(2)) = "Z" Then
    For K = 0 To 6
        fgSelect.Col = K
        fgSelect.CellBackColor = RGB(230, 230, 230)
    Next K
End If

End Sub
Public Sub fgSelect_Display_6DC_Line()
Dim K As Integer
Dim wColor As Long
Dim X As String
On Error Resume Next


fgSelect.Col = 0: fgSelect.Text = rsSab(0)
fgSelect.Col = 1: fgSelect.Text = rsSab(1)

fgSelect.Col = 3: fgSelect.Text = Format(rsSab(2), "### ### ##0") & "  "

If rsSab(3) <> 0 Then
    fgSelect.Col = 4
    fgSelect.Text = Format(-rsSab(3), "### ### ##0.00") & "  "
    If rsSab(3) > 0 Then
        fgSelect.CellForeColor = vbRed
    Else
        fgSelect.CellForeColor = vbBlue
    End If
End If
If rsSab(4) <> 0 Then
    fgSelect.Col = 5
    fgSelect.Text = Format(rsSab(4), "### ### ##0.00") & "  "
    fgSelect.CellForeColor = vbRed
End If

X = "select BASTABDON from " & paramIBM_Library_SAB & ".ZBASTAB0" _
     & " where BASTABETA = 1  and BASTABNUM = 44 and BASTABARG = '" & rsSab(1) & "'"
Set rsSab_X = cnsab.Execute(X)

If Not rsSab_X.EOF Then
    fgSelect.Col = 2: fgSelect.Text = Mid$(rsSab_X("BASTABDON"), 1, 30)
End If


End Sub

Public Sub fgSelect_Display_6RDC_Total()
Dim K As Integer, nbT As Long, curT As Currency
Dim wColor As Long
Dim X As String
On Error Resume Next


fgSelect.Visible = False

Do While Not rsSab.EOF
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1

    fgSelect.Col = 2: fgSelect.Text = rsSab(0)
    
    fgSelect.Col = 4: fgSelect.Text = Format(rsSab(1), "### ### ##0") & "  "
    fgSelect.CellForeColor = vbBlue
    nbT = nbT + rsSab(1)
    If rsSab(2) <> 0 Then
        fgSelect.Col = 6
        fgSelect.Text = Format(rsSab(2), "### ### ##0.00") & "  "
        fgSelect.CellForeColor = vbRed
        curT = curT + rsSab(2)
    End If
    
    X = "select BASTABDON from " & paramIBM_Library_SAB & ".ZBASTAB0" _
         & " where BASTABETA = 1  and BASTABNUM = 44 and BASTABARG = '" & rsSab(0) & "'"
    Set rsSab_X = cnsab.Execute(X)
    
    If Not rsSab_X.EOF Then
        fgSelect.Col = 3: fgSelect.Text = Mid$(rsSab_X("BASTABDON"), 1, 30)
    End If
    
    For K = 0 To 6
        fgSelect.Col = K
        fgSelect.CellBackColor = mColor_Y1
    Next K

    rsSab.MoveNext

Loop

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1

    fgSelect.Col = 2: fgSelect.Text = "Total"
    
    fgSelect.Col = 4: fgSelect.Text = Format(nbT, "### ### ##0") & "  "
    fgSelect.Col = 6: fgSelect.Text = Format(curT, "### ### ##0.00") & "  "
    fgSelect.CellForeColor = vbRed
    For K = 0 To 6
        fgSelect.Col = K
        fgSelect.CellBackColor = mColor_Y2
        fgSelect.CellFontBold = True
    Next K

fgSelect.Visible = True

End Sub



Private Sub cmdSelect_SQL_1()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"
xWhere = ""
mCOMRCDNAT = Trim(Mid$(cboSelect_COMRCDCLI, 1, 1))
mCOMRCDCLI = Trim(Mid$(cboSelect_COMRCDCLI, 3, 7))
mCOMRCDOPE = Trim(Mid$(cboSelect_COMRCDCLI, 11, 3))
mCOMRCDSER = Trim(Mid$(cboSelect_COMRCDCLI, 15, 2))
mCOMRCDSSE = Trim(Mid$(cboSelect_COMRCDCLI, 18, 2))
If mCOMRCDCLI <> "" Then
    xSQL = "select CLIENARA1 , CLIENARA2 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & mCOMRCDCLI & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        libCOMRCDCLI = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
    Else
        libCOMRCDCLI = "???"
    End If
    
    xWhere = " where COMRCDnat = ' ' and COMRCDCLI = '" & mCOMRCDCLI & "'  and COMRCDOPE = '" & mCOMRCDOPE & "' and COMRCDSER = '" & mCOMRCDSER & "'  and COMRCDSSE = '" & mCOMRCDSSE & "'"
    mExportation = "Client  " & mCOMRCDCLI & " -" & libCOMRCDCLI & " - Liste des commissions "
    If Not IsNull(txtSelect_COMRCDDTR_Min.Value) Then
        Call DTPicker_Control(txtSelect_COMRCDDTR_Min, wAmjMin)
        Call DTPicker_Control(txtSelect_COMRCDDTR_Max, wAmjMax)
        xWhere = xWhere & " and COMRCDDTR >= " & wAmjMin - 19000000 & " And COMRCDDTR <= " & wAmjMax - 19000000
         mExportation = mExportation & "du " & dateImp10_S(wAmjMin) & "au " & dateImp10_S(wAmjMax)
    Else
        xWhere = xWhere & " and COMRCDRLV = " & Val(txtSelect_COMRCDRLV)
         mExportation = mExportation & " - Relevé N°" & Val(txtSelect_COMRCDRLV)
    End If
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCOMRCD0 " & xWhere & " order by COMRCDDTR , COMRCDPIE , COMRCDECR"
    Set rsSab = cnsab.Execute(xSQL)
      
    Call fgSelect_Display_1

End If

Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_2()
Dim V, X As String
Dim xSQL As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2"
mCOMRCDRLV_Where = ""
If Mid$(cboSelect_COMRCDCLI, 1, 1) <> "$" Then
    Call MsgBox("Ce client n'est pas paramétré pour le calcul de rétrocession de commissions", vbExclamation)
Else
    mCOMRCDNAT = Trim(Mid$(cboSelect_COMRCDCLI, 1, 1))
    mCOMRCDCLI = Trim(Mid$(cboSelect_COMRCDCLI, 3, 7))
    mCOMRCDOPE = Trim(Mid$(cboSelect_COMRCDCLI, 11, 3))
    mCOMRCDSER = Trim(Mid$(cboSelect_COMRCDCLI, 15, 2))
    mCOMRCDSSE = Trim(Mid$(cboSelect_COMRCDCLI, 18, 2))
    xSQL = "select CLIENARA1 , CLIENARA2 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & mCOMRCDCLI & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        libCOMRCDCLI = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
    Else
        libCOMRCDCLI = "???"
    End If
    
    mCOMRCDRLV_Where = " where COMRCDnat = ' ' and COMRCDCLI = '" & mCOMRCDCLI & "'" _
                     & " and COMRCDOPE = '" & mCOMRCDOPE & "' and COMRCDSER = '" & mCOMRCDSER & "'  and COMRCDSSE = '" & mCOMRCDSSE & "'" _
                     & " and COMRCDSTA in (' ' , 'M') and COMRCDRLV = " & Val(txtSelect_COMRCDRLV)
    mExportation = "Client  " & mCOMRCDCLI & " -" & libCOMRCDCLI & " - Relevé des rétrocessions de commissions "
    If Not IsNull(txtSelect_COMRCDDTR_Min.Value) Then
        Call DTPicker_Control(txtSelect_COMRCDDTR_Min, wAmjMin)
        Call DTPicker_Control(txtSelect_COMRCDDTR_Max, wAmjMax)
        mCOMRCDRLV_Where = mCOMRCDRLV_Where & " and COMRCDDTR >= " & wAmjMin - 19000000 & " And COMRCDDTR <= " & wAmjMax - 19000000
        mExportation = mExportation & " du " & dateImp10_S(wAmjMin) & "au " & dateImp10_S(wAmjMax)
    End If
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCOMRCD0 " & mCOMRCDRLV_Where & " order by COMRCDDTR , COMRCDPIE , COMRCDECR"
    Set rsSab = cnsab.Execute(xSQL)
      
    Call fgSelect_Display_2

End If

Set rsSab = Nothing
If Val(txtSelect_COMRCDRLV) = 0 And fgSelect.Rows > 2 Then cmdCOMRCDRLV_Ok.Visible = arrHab(2)
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
        cmdSelect_Ok_Click
    Else
        SendKeys "{TAB}"
    End If
End Sub


Public Sub cmdContext_Quit()
lstErr.Clear: lstErr.Height = 200

If fraCOMRCDCLI.Visible Then
    fraCOMRCDCLI.Visible = False
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

If fgDetail.Visible Then
    fgDetail.Visible = False
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
Call lstErr_Clear(lstErr, cmdContext, "> COM_Retro_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Clear

Select Case cmdSelect_SQL_K
    Case "1": cmdSelect_SQL_1
    Case "2": cmdSelect_SQL_2
    Case "3": cmdSelect_SQL_3
    Case "6 DC", "6 RDC": cmdSelect_SQL_6
'    Case "SPLF": cmdSelect_SQL_SPLF
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< COM_Retro_cmdSelect_Ok"): DoEvents
lstErr.Height = 480
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
cmdSelect_Ok.BackColor = fgSelect.BackColorFixed
End Sub




Private Sub mnuPrint_Excel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String
Call lstErr_AddItem(lstErr, cmdContext, "Rétrocession de commissions  : export Excel ...."): DoEvents
    Select Case cmdSelect_SQL_K
        Case "1":
            X = mExportation
            Call MSflexGrid_Excel("", "COM_Rétrocession " & mCOMRCDCLI, X, fgSelect, 7)
        Case "2":
            X = mExportation
            Call MSflexGrid_Excel("", "COM_Rétrocession " & mCOMRCDCLI, X, fgSelect, 7)
        Case "3":
            X = mExportation
            Call MSflexGrid_Excel("", "COM_Rétrocession " & mCOMRCDCLI, X, fgSelect, 7)
        Case "6 RDC", "6 DC":
            X = mExportation
            Call MSflexGrid_Excel("", "COM_Rétrocession " & mCOMRCDCLI, X, fgSelect, 6)
    End Select

Call lstErr_AddItem(lstErr, cmdContext, "Rétrocession de commissions  : export Excel terminé"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint_Mail_Click()
Dim X As String

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_AddItem(lstErr, cmdContext, "> Rétrocession de commissions  : export mail ...."): DoEvents
    Select Case cmdSelect_SQL_K
        Case "1":
            X = Replace(mExportation, vbCrLf, "<BR>")
            Call MSFlexGrid_SendMail(mMail_Destinataires, "COM_Rétrocession", X, X, fgSelect, 7)
        Case "2":
            X = Replace(mExportation, vbCrLf, "<BR>")
            Call MSFlexGrid_SendMail(mMail_Destinataires, "COM_Rétrocession", X, X, fgSelect, 7)
        Case "3":
            X = Replace(mExportation, vbCrLf, "<BR>")
            Call MSFlexGrid_SendMail(mMail_Destinataires, "COM_Rétrocession", X, X, fgSelect, 7)
        Case "6 RDC", "6 DC":
            X = Replace(mExportation, vbCrLf, "<BR>")
            Call MSFlexGrid_SendMail(mMail_Destinataires, "COM_Rétrocession", X, X, fgSelect, 6)
    End Select

Call lstErr_AddItem(lstErr, cmdContext, "Rétrocession de commissions  : export mail terminé"): DoEvents


Me.Enabled = True: Me.MousePointer = 0

End Sub






Public Sub fraDetail_Display()
Dim xSQL As String

fraDetail_Display_txtRTF
End Sub
Public Sub fraDetail_Display_txtRTF()
Dim xRTF As String, intFile As Integer, xIn As String

End Sub


Public Sub txtRTF_Visible()
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "{\f0\fswiss\fprq2\fcharset0 Calibri;}", "{\f0\fmodern\fprq1\fcharset0 Courier New;}")
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "\cf1\f0\fs20 a\cf2 b\cf3 c\cf4 d\cf5 e\cf6 f\cf7 g\cf8 h\cf9 i\cf10 j\cf11 k\cf12 l\cf13 m\cf14 n\cf15 o\cf16 p", "")
txtRTF.TextRTF = Replace(txtRTF.TextRTF, "[#]", "")
txtRTF.Visible = True

End Sub













Public Sub fraCOMRCDCLI_Display()
Dim xSQL As String
On Error GoTo Error_Handler
currentAction = "fraCOMRCDCLI_Display"
blnCOMRCDZCOM_Add = False
blnCOMRCDCLI_Add = False

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCOMRCD0, " & paramIBM_Library_SAB & ".ZCLIENA0 " _
     & " where COMRCDnat = '$' and COMRCDPIE = " & xYCOMRCD0.COMRCDPIE & " and COMRCDECR = " & xYCOMRCD0.COMRCDECR _
     & " and CLIENAETB = 1 and CLIENACLI = COMRCDCLI order by COMRCDCLI , COMRCDPIE , COMRCDECR"
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    If xYCOMRCD0.COMRCDPIE = 0 Then
        fraCOMRCDCLI.Caption = "Nouveau client"
        Call rsYCOMRCD0_Init(oldYCOMRCD0)
        oldYCOMRCD0.COMRCDCLI = mCOMRCDCLI
        oldYCOMRCD0.COMRCDSER = mCOMRCDSER
        oldYCOMRCD0.COMRCDSSE = mCOMRCDSSE
        oldYCOMRCD0.COMRCDOPE = mCOMRCDOPE
        oldYCOMRCD0.COMRCDNAT = "$"
    Else
        V = "erreur lecture " & xSQL
        GoTo Error_MsgBox
    End If
Else
    Call rsYCOMRCD0_GetBuffer(rsSab, oldYCOMRCD0)
    fraCOMRCDCLI.Caption = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
End If

txtCOMRCDCLI = oldYCOMRCD0.COMRCDCLI
txtCOMRCDSER = oldYCOMRCD0.COMRCDSER
txtCOMRCDSSE = oldYCOMRCD0.COMRCDSSE
txtCOMRCDOPE = oldYCOMRCD0.COMRCDOPE
fraCOMRCDCLI_0.Enabled = False

If oldYCOMRCD0.COMRCDECR = 0 Then
    fraCOMRCDCLI_1.Visible = False
    cmdCOMRCDCLI_Delete.Visible = False
    cmdCOMRCDCLI_Update.Visible = False
    cmdCOMRCDCLI_Add.Visible = True
    cmdCOMRCDCLI_Add.Caption = "Ajouter un code commission"
    cmdCOMRCDCLI_Add.BackColor = vbYellow
    cmdCOMRCDCLI_New.Visible = True
    cmdCOMRCDCLI_New.Caption = "Ajouter un client"
    cmdCOMRCDCLI_New.BackColor = vbRed
    txtCOMRCDZCOM = ""
    txtCOMRCDMTD = ""
    txtCOMRCDMTR = ""
    Call DTPicker_Set(txtCOMRCDDTR, YBIATAB0_DATE_CPT_J)
Else
    fraCOMRCDCLI_1.Visible = True
    cmdCOMRCDCLI_Delete.Visible = True
    cmdCOMRCDCLI_Update.Visible = True
    cmdCOMRCDCLI_Add.Visible = False
    cmdCOMRCDCLI_New.Visible = False
    txtCOMRCDZCOM = Trim(oldYCOMRCD0.COMRCDZCOM)
    txtCOMRCDMTD = IIf(oldYCOMRCD0.COMRCDMTD = 0, "", Format$(Abs(oldYCOMRCD0.COMRCDMTD), "### ### ### ##0.00"))
    txtCOMRCDMTR = IIf(oldYCOMRCD0.COMRCDMTR = 0, "", Format$(Abs(oldYCOMRCD0.COMRCDMTR), "### ### ### ##0.00"))
    
    If oldYCOMRCD0.COMRCDDTR > 0 Then wAmjMin = oldYCOMRCD0.COMRCDDTR + 19000000: Call DTPicker_Set(txtCOMRCDDTR, wAmjMin)
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
         & " where BASTABETA = 1  and BASTABNUM = 44 and BASTABARG = '" & Trim(oldYCOMRCD0.COMRCDZCOM) & "'"
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        fraCOMRCDCLI_1.Caption = Mid$(rsSab("BASTABDON"), 1, 30)
    Else
        fraCOMRCDCLI_1.Caption = ""
    End If
    
End If
fraCOMRCDCLI.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Public Function fraCOMRCDCLI_Control()
Dim X As String, wMsg As String

Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données "): DoEvents
newYCOMRCD0 = oldYCOMRCD0

wMsg = ""
newYCOMRCD0.COMRCDZCOM = Trim(txtCOMRCDZCOM)
If newYCOMRCD0.COMRCDZCOM = "" Then
    wMsg = wMsg & "- préciser le code commission" & vbCrLf
End If
newYCOMRCD0.COMRCDMTD = CCur(num_CDec(txtCOMRCDMTD))
If newYCOMRCD0.COMRCDMTD <= 0 Then
    wMsg = wMsg & "- préciser le montant minimum" & vbCrLf
End If
newYCOMRCD0.COMRCDMTR = CCur(num_CDec(txtCOMRCDMTR))
If newYCOMRCD0.COMRCDMTR <= 0 Then
    wMsg = wMsg & "- préciser le montant rétrocédé" & vbCrLf
End If
If newYCOMRCD0.COMRCDMTR > newYCOMRCD0.COMRCDMTD Then
    wMsg = wMsg & "- le montant rétrocédé > montant de la commission min" & vbCrLf
End If

Call DTPicker_Control(txtCOMRCDDTR, wAmjMin)
newYCOMRCD0.COMRCDDTR = CLng(wAmjMin) - 19000000

If oldYCOMRCD0.COMRCDECR = 0 Then

    If wAmjMin < YBIATAB0_DATE_CPT_J Then
        wMsg = wMsg & "- date de début < date du jour" & vbCrLf
    End If
    
    X = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
         & " where BASTABETA = 1  and BASTABNUM = 44 and BASTABARG = '" & newYCOMRCD0.COMRCDZCOM & "'"
    Set rsSab = cnsab.Execute(X)
    
    If Not rsSab.EOF Then
        fraCOMRCDCLI_1.Caption = Mid$(rsSab("BASTABDON"), 1, 30)
    Else
        wMsg = wMsg & "- code commission inconnu" & vbCrLf
    End If

Else
    If newYCOMRCD0.COMRCDZCOM <> Trim(oldYCOMRCD0.COMRCDZCOM) Then
        wMsg = wMsg & "- le code commission a été modifié !" & vbCrLf
    End If
End If



'__________________________________________________________________________

If wMsg = "" Then
    fraCOMRCDCLI_Control = Null
Else
    Call MsgBox(wMsg, vbCritical, "COM_RETRO : fraCOMRCDCLI_Control")
    fraCOMRCDCLI_Control = "?_________fraCOMRCDCLI_Control"
End If

End Function


Public Function fraCOMRCDCLI_0_Control()
Dim xSQL As String, wMsg As String

Call lstErr_AddItem(lstErr, cmdContext, ">_________Contrôle des données "): DoEvents
newYCOMRCD0 = oldYCOMRCD0

wMsg = ""
newYCOMRCD0.COMRCDCLI = Format(Trim(txtCOMRCDCLI), "0000000")
If newYCOMRCD0.COMRCDCLI = "" Then
    wMsg = wMsg & "- préciser la racine client" & vbCrLf
End If
newYCOMRCD0.COMRCDSER = Trim(txtCOMRCDSER)
If newYCOMRCD0.COMRCDSER = "" Then
    wMsg = wMsg & "- préciser le service" & vbCrLf
End If
newYCOMRCD0.COMRCDSSE = Trim(txtCOMRCDSSE)
If newYCOMRCD0.COMRCDSSE = "" Then
    wMsg = wMsg & "- préciser le sous-service" & vbCrLf
End If

newYCOMRCD0.COMRCDOPE = Trim(txtCOMRCDOPE)
If newYCOMRCD0.COMRCDSSE = "" Then
    wMsg = wMsg & "- préciser le code opération" & vbCrLf
End If

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
     & " where CLIENAETB = 1 and CLIENACLI = '" & newYCOMRCD0.COMRCDCLI & "'"
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    wMsg = wMsg & "- client inconnu " & vbCrLf
Else
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCOMRCD0" _
         & " where COMRCDnat = '$'  and COMRCDECR = 0 and COMRCDCLI = '" & newYCOMRCD0.COMRCDCLI & "'" _
         & " and COMRCDSER = '" & newYCOMRCD0.COMRCDSER & "' and COMRCDSse = '" & newYCOMRCD0.COMRCDSSE & "' and COMRCDope = '" & newYCOMRCD0.COMRCDOPE & "'"
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        wMsg = wMsg & "- client déjà enregistré " & vbCrLf
    End If
End If

'__________________________________________________________________________

If wMsg = "" Then
    fraCOMRCDCLI_0_Control = Null
Else
    Call MsgBox(wMsg, vbCritical, "COM_RETRO : fraCOMRCDCLI_Control")
    fraCOMRCDCLI_0_Control = "?_________fraCOMRCDCLI_Control"
End If

End Function








