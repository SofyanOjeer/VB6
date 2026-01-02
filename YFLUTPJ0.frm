VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmYFLUTPJ0 
   AutoRedraw      =   -1  'True
   Caption         =   "Tresorerie prévisionnelle"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   405
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
   Icon            =   "YFLUTPJ0.frx":0000
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
      Height          =   9852
      Left            =   0
      TabIndex        =   2
      Top             =   520
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   17383
      _Version        =   393216
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
      TabCaption(0)   =   "CB : tableau de suivi de trésorerie prévisionnelle"
      TabPicture(0)   =   "YFLUTPJ0.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Gestion des flux HORS SAB"
      TabPicture(1)   =   "YFLUTPJ0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraFlux"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Paramétrage"
      TabPicture(2)   =   "YFLUTPJ0.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraParam"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraFlux 
         Height          =   9372
         Left            =   -74640
         TabIndex        =   18
         Top             =   480
         Width           =   12852
         Begin VB.Frame fraFlux_Detail 
            BackColor       =   &H00B0E0FF&
            Caption         =   "DETAIL"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   4296
            Left            =   6720
            TabIndex        =   31
            Top             =   1440
            Visible         =   0   'False
            Width           =   6012
            Begin MSFlexGridLib.MSFlexGrid fgFlux_Detail 
               Height          =   3060
               Left            =   240
               TabIndex        =   32
               Top             =   1080
               Width           =   5580
               _ExtentX        =   9843
               _ExtentY        =   5398
               _Version        =   393216
               Cols            =   7
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   15794175
               ForeColor       =   12582912
               BackColorFixed  =   12632064
               ForeColorFixed  =   16777215
               ForeColorSel    =   8388608
               BackColorBkg    =   15794175
               AllowUserResizing=   3
               FormatString    =   "<Seq|>Echéance      |>Montant                    |<Dev   |<Origine|> Id                  |"
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
            Begin VB.Label libFlux_FLUTPJTXT 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   492
               Left            =   240
               TabIndex        =   66
               Top             =   360
               Width           =   5532
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame fraFlux_Update 
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3492
            Left            =   6720
            TabIndex        =   28
            Top             =   5760
            Visible         =   0   'False
            Width           =   6012
            Begin VB.Frame fraFlux_UPDATE_C 
               BackColor       =   &H00E0FFFF&
               Height          =   1600
               Left            =   240
               TabIndex        =   64
               Top             =   1800
               Width           =   5412
               Begin VB.ComboBox cboFlux_FLUTPJCLI 
                  Height          =   312
                  Left            =   1200
                  Style           =   2  'Dropdown List
                  TabIndex        =   69
                  Top             =   360
                  Width           =   3936
               End
               Begin VB.TextBox txtFLUX_FLUTPJTXT 
                  Height          =   684
                  Left            =   120
                  MultiLine       =   -1  'True
                  TabIndex        =   65
                  Top             =   840
                  Width           =   5172
               End
               Begin VB.Label lblFlux_FLUTPJCLI 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Racine"
                  Height          =   252
                  Left            =   120
                  TabIndex        =   68
                  Top             =   360
                  Width           =   972
               End
               Begin VB.Label lblFlux_FLUTPJTXT 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Informations complémentaires :"
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Left            =   120
                  TabIndex        =   67
                  Top             =   0
                  Width           =   2772
               End
            End
            Begin VB.Frame fraFlux_Update_B 
               BackColor       =   &H00E0FFFF&
               Height          =   852
               Left            =   240
               TabIndex        =   49
               Top             =   960
               Width           =   4212
               Begin VB.TextBox txtFlux_FLUTPJMTD 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   1800
                  TabIndex        =   54
                  Top             =   360
                  Width           =   2052
               End
               Begin MSComCtl2.DTPicker txtFlux_FLUTPJECH 
                  Height          =   300
                  Left            =   240
                  TabIndex        =   53
                  Top             =   360
                  Width           =   1212
                  _ExtentX        =   2143
                  _ExtentY        =   529
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial Unicode MS"
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
                  Format          =   103153667
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin VB.Label lblFlux_FLUTPJMTD 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Montant"
                  Height          =   252
                  Left            =   3000
                  TabIndex        =   56
                  Top             =   120
                  Width           =   732
               End
               Begin VB.Label lblFlux_FLUTPJECH 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Echéance"
                  Height          =   252
                  Left            =   360
                  TabIndex        =   55
                  Top             =   120
                  Width           =   732
               End
            End
            Begin VB.Frame fraFlux_Update_A 
               BackColor       =   &H00E0FFFF&
               Height          =   852
               Left            =   240
               TabIndex        =   48
               Top             =   240
               Width           =   4212
               Begin VB.ComboBox cboFlux_Frequence 
                  Height          =   312
                  Left            =   1800
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   52
                  Top             =   240
                  Width           =   2256
               End
               Begin VB.OptionButton optFlux_Encaissement 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Encaissement"
                  Height          =   372
                  Left            =   240
                  TabIndex        =   51
                  Top             =   420
                  Width           =   1572
               End
               Begin VB.OptionButton optFlux_Decaissement 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Décaissement"
                  Height          =   372
                  Left            =   240
                  TabIndex        =   50
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   1572
               End
            End
            Begin VB.CommandButton cmdFlux_Update_Ok 
               BackColor       =   &H0080FF80&
               Caption         =   "Enregistrer"
               Height          =   600
               Left            =   4680
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   1080
               Width           =   1020
            End
            Begin VB.CommandButton cmdFlux_Update_Quit 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Abandonner"
               Height          =   600
               Left            =   4680
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   240
               Width           =   1020
            End
         End
         Begin VB.CommandButton cmdFlux_Ok 
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
            Left            =   10200
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   600
            Width           =   1335
         End
         Begin VB.Frame fraFlux_Options 
            BackColor       =   &H00F0FFFF&
            Height          =   852
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   8712
            Begin VB.ComboBox cboFlux_FLUTPJDEV 
               Height          =   312
               Left            =   6840
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   450
               Width           =   1176
            End
            Begin VB.ComboBox cboFlux_FLUTPJOD 
               Height          =   312
               Left            =   2520
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   450
               Width           =   3576
            End
            Begin MSComCtl2.DTPicker txtFlux_FLUTPJECH_Min 
               Height          =   300
               Left            =   240
               TabIndex        =   21
               Top             =   450
               Width           =   1212
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Unicode MS"
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
               Format          =   103153667
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblFlux_FLUTPJDEV 
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
               Height          =   252
               Left            =   6960
               TabIndex        =   25
               Top             =   120
               Width           =   612
            End
            Begin VB.Label lblFlux_TP7OPHDTR 
               BackColor       =   &H00F0FFFF&
               Caption         =   "date du flux"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   360
               TabIndex        =   23
               Top             =   180
               Width           =   852
            End
            Begin VB.Label lblFlux_FLUTPJOD 
               BackColor       =   &H00F0FFFF&
               Caption         =   "code interne"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3000
               TabIndex        =   22
               Top             =   240
               Width           =   972
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgFlux 
            Height          =   7860
            Left            =   120
            TabIndex        =   27
            Top             =   1440
            Visible         =   0   'False
            Width           =   12540
            _ExtentX        =   22119
            _ExtentY        =   13864
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   15794175
            ForeColor       =   12582912
            BackColorFixed  =   12632064
            ForeColorFixed  =   16777215
            ForeColorSel    =   8388608
            BackColorBkg    =   15794175
            AllowUserResizing=   3
            FormatString    =   $"YFLUTPJ0.frx":035E
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
      Begin VB.Frame fraParam 
         Height          =   9252
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   13212
         Begin VB.ListBox lstW_Sorted 
            Height          =   696
            Left            =   7800
            Sorted          =   -1  'True
            TabIndex        =   71
            Top             =   2400
            Visible         =   0   'False
            Width           =   4692
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
            Height          =   5460
            Left            =   7680
            TabIndex        =   17
            Top             =   3480
            Visible         =   0   'False
            Width           =   5292
         End
         Begin VB.Frame fraParam_K 
            BackColor       =   &H00E0FFFF&
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4932
            Left            =   7440
            TabIndex        =   33
            Top             =   3840
            Width           =   5652
            Begin VB.CommandButton cmdParam_Ok 
               BackColor       =   &H0080FF80&
               Caption         =   "Enregistrer"
               Height          =   840
               Left            =   3240
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   3840
               Width           =   1260
            End
            Begin VB.CommandButton cmdParam_Quit 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Abandonner"
               Height          =   840
               Left            =   1320
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   3840
               Width           =   1260
            End
            Begin VB.Frame fraParam_K2 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   0  'None
               Height          =   3252
               Left            =   120
               TabIndex        =   34
               Top             =   360
               Width           =   5412
               Begin VB.TextBox txtParam_K2 
                  Height          =   288
                  Left            =   2880
                  TabIndex        =   36
                  Top             =   120
                  Width           =   1332
               End
               Begin VB.ComboBox cboParam_Frequence 
                  Height          =   312
                  Left            =   1200
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   40
                  Top             =   2760
                  Width           =   2256
               End
               Begin VB.ComboBox cboParam_CCB_CR 
                  Height          =   312
                  Left            =   1200
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   39
                  Top             =   2160
                  Width           =   4056
               End
               Begin VB.ComboBox cboParam_CCB_DB 
                  Height          =   312
                  Left            =   1200
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   38
                  Top             =   1560
                  Width           =   4056
               End
               Begin VB.TextBox txtParam_X 
                  Height          =   732
                  Left            =   1200
                  MultiLine       =   -1  'True
                  TabIndex        =   37
                  Text            =   "YFLUTPJ0.frx":043D
                  Top             =   600
                  Width           =   4092
               End
               Begin VB.TextBox txtParam_K 
                  Height          =   288
                  Left            =   1200
                  TabIndex        =   35
                  Top             =   120
                  Width           =   1332
               End
               Begin VB.Label lblParam_Frequence 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "fréquence"
                  Height          =   252
                  Left            =   240
                  TabIndex        =   45
                  Top             =   2880
                  Width           =   732
               End
               Begin VB.Label lblParam_CCB_CR 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Code CB encaissement"
                  Height          =   492
                  Left            =   120
                  TabIndex        =   44
                  Top             =   2160
                  Width           =   972
               End
               Begin VB.Label lblParam_CCB_DB 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Code CB décaissement"
                  Height          =   492
                  Left            =   120
                  TabIndex        =   43
                  Top             =   1560
                  Width           =   972
               End
               Begin VB.Label lblParam_K 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Code"
                  Height          =   252
                  Left            =   120
                  TabIndex        =   42
                  Top             =   120
                  Width           =   732
               End
               Begin VB.Label lblParam_X 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "libellé"
                  Height          =   372
                  Left            =   120
                  TabIndex        =   41
                  Top             =   720
                  Width           =   852
               End
            End
         End
         Begin VB.ListBox lstParam_Action 
            BackColor       =   &H00F0FFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Left            =   360
            TabIndex        =   16
            Top             =   240
            Width           =   6660
         End
         Begin VB.ListBox lstParam_K 
            BackColor       =   &H00F0FFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6900
            Left            =   360
            TabIndex        =   15
            Top             =   2160
            Width           =   6780
         End
         Begin MSFlexGridLib.MSFlexGrid fgFLUTPJSTAT 
            Height          =   3180
            Left            =   7440
            TabIndex        =   70
            Top             =   240
            Visible         =   0   'False
            Width           =   5580
            _ExtentX        =   9843
            _ExtentY        =   5609
            _Version        =   393216
            Cols            =   7
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   15794175
            ForeColor       =   12582912
            BackColorFixed  =   12632064
            ForeColorFixed  =   16777215
            ForeColorSel    =   8388608
            BackColorBkg    =   15794175
            AllowUserResizing=   3
            FormatString    =   "Code          | Dev |>NB            |> Moyenne         |> Ecart-Type           |> M + 2*S          |"
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
      Begin VB.Frame fraSelect 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9420
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   13296
         Begin VB.Frame fraSelect_Option_2 
            BackColor       =   &H00F0FFFF&
            BorderStyle     =   0  'None
            Height          =   612
            Left            =   360
            TabIndex        =   59
            Top             =   0
            Visible         =   0   'False
            Width           =   5292
            Begin VB.ComboBox cboSelect_FLUTPJOPE 
               Height          =   312
               Left            =   2760
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   61
               Top             =   240
               Width           =   1896
            End
            Begin VB.ComboBox cboSelect_FLUTPJORIG 
               Height          =   330
               Left            =   120
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label lblSelect_FLUTPJOPE 
               BackColor       =   &H00F0FFFF&
               Caption         =   "code opération"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3120
               TabIndex        =   63
               Top             =   0
               Width           =   1332
            End
            Begin VB.Label lblSelect_FLUTPJORIG 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Origine du flux"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   360
               TabIndex        =   62
               Top             =   0
               Width           =   1332
            End
         End
         Begin VB.Frame fraDetail 
            BackColor       =   &H00B0E0FF&
            Caption         =   "DETAIL"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   7416
            Left            =   3480
            TabIndex        =   12
            Top             =   1680
            Visible         =   0   'False
            Width           =   9492
            Begin MSFlexGridLib.MSFlexGrid fgDetail 
               Height          =   6780
               Left            =   360
               TabIndex        =   13
               Top             =   360
               Width           =   8820
               _ExtentX        =   15558
               _ExtentY        =   11959
               _Version        =   393216
               Cols            =   11
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   15794175
               ForeColor       =   12582912
               BackColorFixed  =   12632064
               ForeColorFixed  =   16777215
               ForeColorSel    =   8388608
               BackColorBkg    =   15794175
               AllowUserResizing=   3
               FormatString    =   $"YFLUTPJ0.frx":0443
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
            Height          =   324
            Left            =   9360
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   360
            Width           =   3732
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
            Left            =   11760
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   720
            Width           =   1335
         End
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00F0FFFF&
            Height          =   972
            Left            =   360
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   8712
            Begin VB.CheckBox chkSelect_FLUTPJDEV 
               BackColor       =   &H00F0FFFF&
               Caption         =   "flux exprimés en contre-valeur €"
               Height          =   252
               Left            =   600
               TabIndex        =   57
               Top             =   480
               Width           =   2652
            End
            Begin MSComCtl2.DTPicker txtSelect_FLUTPJECH_Max 
               Height          =   300
               Left            =   6960
               TabIndex        =   9
               Top             =   480
               Width           =   1212
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Unicode MS"
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
               Format          =   103153667
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_FLUTPJECH_Min 
               Height          =   300
               Left            =   5400
               TabIndex        =   11
               Top             =   480
               Width           =   1212
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Unicode MS"
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
               Format          =   103153667
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label libSelect_FLUTPJDEV 
               BackColor       =   &H00F0FFFF&
               Caption         =   "<date de cours>"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3240
               TabIndex        =   58
               Top             =   480
               Visible         =   0   'False
               Width           =   2052
            End
            Begin VB.Label lblSelect_TP7OPHDTR 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Flux de la période"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   6000
               TabIndex        =   10
               Top             =   120
               Width           =   1452
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   8028
            Left            =   360
            TabIndex        =   5
            Top             =   1320
            Visible         =   0   'False
            Width           =   12672
            _ExtentX        =   22357
            _ExtentY        =   14155
            _Version        =   393216
            Rows            =   1
            Cols            =   12
            FixedCols       =   2
            RowHeightMin    =   300
            BackColor       =   15794175
            ForeColor       =   0
            BackColorFixed  =   8421376
            ForeColorFixed  =   16777215
            BackColorSel    =   12648384
            BackColorBkg    =   15794175
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"YFLUTPJ0.frx":04DA
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
      Picture         =   "YFLUTPJ0.frx":0607
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
      Begin VB.Menu mnuCB_Export 
         Caption         =   "CB Export .xlsx"
      End
      Begin VB.Menu mnuECH_Export 
         Caption         =   "ECH Export .xlsx"
      End
      Begin VB.Menu mnuECH_Mail 
         Caption         =   "ECH Mail"
      End
      Begin VB.Menu mnuParam_Export 
         Caption         =   "Param Export .xlsx"
      End
   End
   Begin VB.Menu mnuParam 
      Caption         =   "mnuParam"
      Visible         =   0   'False
      Begin VB.Menu mnuParam_Add 
         Caption         =   "ajouter un enregistrement"
      End
      Begin VB.Menu mnuParam_Update 
         Caption         =   "modifier cet enregistrement"
      End
      Begin VB.Menu mnuParam_Delete 
         Caption         =   "supprimer cet enregistrement"
      End
   End
   Begin VB.Menu mnuFlux_Detail 
      Caption         =   "mnuFlux_Detail"
      Visible         =   0   'False
      Begin VB.Menu mnuFlux_Update 
         Caption         =   "modifier ce flux"
      End
      Begin VB.Menu mnuFlux_Annulation 
         Caption         =   "annuler ce flux"
      End
      Begin VB.Menu mnuFlux_Delete 
         Caption         =   "supprimer ce flux"
      End
   End
   Begin VB.Menu mnuFlux 
      Caption         =   "mnuFlux"
      Visible         =   0   'False
      Begin VB.Menu mnuFlux_Add 
         Caption         =   "ajouter un dossier"
      End
      Begin VB.Menu mnuFlux_Close 
         Caption         =   "clôturer ce dossier"
      End
      Begin VB.Menu mnuFlux_FLUTPJTXT 
         Caption         =   "modifier le commentaire"
      End
   End
End
Attribute VB_Name = "frmYFLUTPJ0"
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
Dim intReturn As Integer, currentError As String
Dim YFLUTPJ0_Aut As typeAuthorization
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

Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long
Dim xYBIACPT0 As typeYBIACPT0, newYBIACPT0 As typeYBIACPT0, oldYBIACPT0 As typeYBIACPT0
Dim arrYBIACPT0() As typeYBIACPT0, arrYBIACPT0_Nb As Long, arrYBIACPT0_Max As Long, arrYBIACPT0_Index As Long

Dim xYFLUTPJ0 As typeYFLUTPJ0, newYFLUTPJ0 As typeYFLUTPJ0, oldYFLUTPJ0 As typeYFLUTPJ0
Dim arrYFLUTPJ0() As typeYFLUTPJ0, arrYFLUTPJ0_Nb As Long, arrYFLUTPJ0_Max As Long, arrYFLUTPJ0_Index As Long

Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean



Dim fgFlux_Detail_FormatString As String, fgFlux_Detail_K As Integer
Dim fgFlux_Detail_RowDisplay As Integer, fgFlux_Detail_RowClick As Integer, fgFlux_Detail_ColClick As Integer
Dim fgFlux_Detail_ColorClick As Long, fgFlux_Detail_ColorDisplay As Long
Dim fgFlux_Detail_Sort1 As Integer, fgFlux_Detail_Sort2 As Integer
Dim fgFlux_Detail_SortAD As Integer, fgFlux_Detail_Sort1_Old As Integer
Dim fgFlux_Detail_arrIndex As Integer
Dim blnfgFlux_Detail_DisplayLine As Boolean

'''''Dim xFlux_Detail As typeYFLUTPJ0, newFlux_Detail As typeYFLUTPJ0, oldFlux_Detail As typeYFLUTPJ0
Dim arrfgFlux_Detail() As typeYFLUTPJ0, arrfgFlux_Detail_Nb As Long, arrfgFlux_Detail_Max As Long, arrfgFlux_Detail_Index As Long

Dim fgFlux_FormatString As String, fgFlux_K As Integer
Dim fgFlux_RowDisplay As Integer, fgFlux_RowClick As Integer, fgFlux_ColClick As Integer
Dim fgFlux_ColorClick As Long, fgFlux_ColorDisplay As Long
Dim fgFlux_Sort1 As Integer, fgFlux_Sort2 As Integer
Dim fgFlux_SortAD As Integer, fgFlux_Sort1_Old As Integer
Dim fgFlux_arrIndex As Integer
Dim blnfgFlux_DisplayLine As Boolean

Dim xFlux As typeYFLUTPJ0, newFlux As typeYFLUTPJ0, oldFlux As typeYFLUTPJ0
Dim arrfgFlux() As typeYFLUTPJ0, arrfgFlux_Nb As Long, arrfgFlux_Max As Long, arrfgFlux_Index As Long
Dim mnuFlux_Option As String


Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim arrDev() As String, arrDev_Nb As Integer, arrDev_DB() As Currency, arrDev_CR() As Currency
Dim arrCCB_K() As String, arrCCB_Lib() As String, arrCCB_Nb As Integer
Dim mAMJDEV As String, arrDev_Cours() As Double

Dim lstParam_Action_K As String
Dim xYBIATAB0 As typeYBIATAB0, Old_YBIATAB0 As typeYBIATAB0, New_YBIATAB0 As typeYBIATAB0
Dim mnuParam_Option As String

Dim mCCB_DB_Row As Integer, mCCB_CR_Row As Integer, mCCB_SD_Row As Integer
Dim mSelect_Where As String
Dim mDate_J7 As String, mDate_A2 As String

Dim paramOD As typeYBIATAB0
Dim arrOD_K() As String, arrOD_Lib() As String, arrOD_Nb As Integer
Dim arrOPE_K() As String, arrOPE_Lib() As String, arrOPE_Nb As Integer
Dim cmdPrint_Option As String

Dim xYFLUTPJ1 As typeYFLUTPJ1, oldYFLUTPJ1 As typeYFLUTPJ1, newYFLUTPJ1 As typeYFLUTPJ1
Dim xYFLUTPJ1_Detail As typeYFLUTPJ1, oldYFLUTPJ1_Detail As typeYFLUTPJ1, newYFLUTPJ1_Detail As typeYFLUTPJ1


Dim arrParam() As typeYBIATAB0
Dim rsSabX As New ADODB.Recordset

Dim auto_YFLUTPJ0_mFile As String

Dim xYFLUTP20 As typeYFLUTP20, oldYFLUTP20 As typeYFLUTP20
Dim arrYFLUTP20() As typeYFLUTP20, arrYFLUTP20_Nb As Long, arrYFLUTP20_Max As Long, arrYFLUTP20_Index As Long
Dim arrSINFO_LIQU(4) As typeSINFO_LIQU
Dim arrAMJMIN(3) As String, arrAMJMAX(3) As String
Public Sub YFLUTPJ0_CB_Export(blnInputBox As Boolean, wFile As String)
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFilex As String, wFile2 As String, xSQL As String
Dim wAMJMin As String, WAMJMax As String
Dim X As String, K As Long
'______________________________________________
Call DTPicker_Control(txtSelect_FLUTPJECH_Min, wAMJMin)

wFile = Trim("C:\Temp\Tréserorie Prévisionnelle CB " & dateImp_Amj(DSys) & ".xlsx")
'______________________________________________
If blnInputBox Then

    X = InputBox("par défaut : " & wFile _
        & vbCrLf & vbCrLf & "     =========================" _
        & vbCrLf & "     =========================", "Trésorerie prévisionnelle : nom du fichier d'exportation CB", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
 End If

'_________________________________________


If Dir(wFile) <> "" Then Kill wFile

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "FLUTPJ"
    .Subject = ""
End With

appExcel.Worksheets.Add , , 2

'If chkSelect_FLUTPJDEV = "1" Then
'    Set wsExcel = wbExcel.Sheets(2)
'    wsExcel.Name = "TP € " & dateImp10(DSys)
'    YFLUTPJ0_CB_Export_Recap
    
'    chkSelect_FLUTPJDEV = "0"
'    cmdSelect_Ok_Click

'    Set wsExcel = wbExcel.Sheets(1) 'wbExcel.ActiveSheet
'    wsExcel.Name = "TP dev " & dateImp10(DSys)
'    YFLUTPJ0_CB_Export_Recap
'Else
    Call cbo_Scan("1", cboSelect_SQL)
    chkSelect_FLUTPJDEV = "0"
    cmdSelect_Ok_Click
    Set wsExcel = wbExcel.Sheets(1) 'wbExcel.ActiveSheet
    wsExcel.Name = "TP dev " & dateImp10(DSys)
    YFLUTPJ0_CB_Export_Recap
    
    chkSelect_FLUTPJDEV = "1"
    cmdSelect_Ok_Click
    
    Set wsExcel = wbExcel.Sheets(2)
    wsExcel.Name = "TP € " & dateImp10(DSys)
    YFLUTPJ0_CB_Export_Recap
'End If

Set wsExcel = wbExcel.Sheets(3)
wsExcel.Name = "Detail " & dateImp10(DSys)

arrYFLUTPJ0_SQL mSelect_Where & " order by FLUTPJECH, FLUTPJCCB , FLUTPJDEV"
YFLUTPJ0_CB_Export_Detail

'__________________________________________________________________________________
Call cbo_Scan("3", cboSelect_SQL)
cmdSelect_Ok_Click

Set wsExcel = wbExcel.Sheets(4) 'wbExcel.ActiveSheet
wsExcel.Name = "Refinancement " & dateImp10(DSys)
YFLUTPJ0_CB_Export_Refinancement

Set wsExcel = wbExcel.Sheets(5)
wsExcel.Name = "Detail R" & dateImp10(DSys)
YFLUTPJ0_CB_Export_Refinancement_Detail

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
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents

End Sub

Public Sub YFLUTPJ0_Param_Export()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSQL As String
Dim wAMJMin As String, WAMJMax As String
Dim X As String, K As Long, K1 As Long, K2 As Long
Dim XDB As String, xCR As String
Dim wColor As Long

'______________________________________________
Call DTPicker_Control(txtSelect_FLUTPJECH_Min, wAMJMin)

wFile = Trim("C:\Temp\Tréserorie Prévisionnelle Param " & dateImp_Amj(wAMJMin) & ".xlsx")
'______________________________________________

X = InputBox("par défaut : " & wFile _
    & vbCrLf & vbCrLf & "     =========================" _
    & vbCrLf & "     =========================", "Trésorerie prévisionnelle : nom du fichier d'exportation CB", wFile)
If Trim(X) = "" Then Exit Sub

wFilex = Trim(X)
'______________________________________________
If wFile <> wFilex Then
    wFile = wFilex
End If
'_________________________________________

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

lstW_Sorted.Clear

If Dir(wFile) <> "" Then Kill wFile

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "FLUTPJ Param"
    .Subject = ""
End With

appExcel.Worksheets.Add , , 3
'__________________________________________________________________________________

Set wsExcel = wbExcel.Sheets(1) 'wbExcel.ActiveSheet
wsExcel.Name = "Code CB"

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


wsExcel.Columns(1).ColumnWidth = 8
wsExcel.Columns(2).ColumnWidth = 50

wRow = 1
wsExcel.Cells(1, 1) = "Code": wsExcel.Cells(1, 1).Interior.Color = RGB(255, 170, 80)
wsExcel.Cells(1, 2) = wsExcel.Name & " (" & dateImp10(DSys) & ")": wsExcel.Cells(1, 2).Interior.Color = RGB(255, 170, 80)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJCCB' order by BIATABK1"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    Call rsYBIATAB0_GetBuffer(rsSab, xYBIATAB0)
    wRow = wRow + 1
    wsExcel.Cells(wRow, 1) = xYBIATAB0.BIATABK1
    wsExcel.Cells(wRow, 2) = xYBIATAB0.BIATABTXT
    lstW_Sorted.AddItem xYBIATAB0.BIATABK1 & "  : :" & xYBIATAB0.BIATABTXT
    rsSab.MoveNext
Loop

'__________________________________________________________________________________
Set wsExcel = wbExcel.Sheets(2)
wsExcel.Name = "Flux échéancés"

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

wsExcel.Columns(1).ColumnWidth = 8
wsExcel.Columns(2).ColumnWidth = 8
wsExcel.Columns(3).ColumnWidth = 8
wsExcel.Columns(4).ColumnWidth = 50


wRow = 1
wsExcel.Cells(1, 1) = "Code": wsExcel.Cells(1, 1).Interior.Color = RGB(255, 170, 80)
wsExcel.Cells(1, 2) = "Nature": wsExcel.Cells(1, 2).Interior.Color = RGB(255, 170, 80)
wsExcel.Cells(1, 3) = "DB  CR": wsExcel.Cells(1, 3).Interior.Color = RGB(255, 170, 80)
wsExcel.Cells(1, 4) = wsExcel.Name & " (" & dateImp10(DSys) & ")": wsExcel.Cells(1, 4).Interior.Color = RGB(255, 170, 80)


xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJOPE' order by BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    Call rsYBIATAB0_GetBuffer(rsSab, xYBIATAB0)
    wRow = wRow + 1
    wsExcel.Cells(wRow, 1) = xYBIATAB0.BIATABK1
    wsExcel.Cells(wRow, 2) = xYBIATAB0.BIATABK2
    wsExcel.Cells(wRow, 3) = xYBIATAB0.BIATABTXT
    X = mId$(xYBIATAB0.BIATABK1, 1, 3) & mId$(xYBIATAB0.BIATABK2, 1, 3)
    For K = 1 To arrOPE_Nb
        If X = arrOPE_K(K) Then Exit For
    Next K
    wsExcel.Cells(wRow, 4) = arrOPE_Lib(K)
    
    X = ":" & Trim(xYBIATAB0.BIATABK1) & " " & Trim(xYBIATAB0.BIATABK2) & ":" & arrOPE_Lib(K)
    If mId$(xYBIATAB0.BIATABTXT, 1, 3) <> "000" Then
        lstW_Sorted.AddItem mId$(xYBIATAB0.BIATABTXT, 1, 3) & "ED" & X
    End If
    If mId$(xYBIATAB0.BIATABTXT, 5, 3) <> "000" Then
        lstW_Sorted.AddItem mId$(xYBIATAB0.BIATABTXT, 5, 3) & "EC" & X
    End If
    rsSab.MoveNext
Loop
'__________________________________________________________________________________
Set wsExcel = wbExcel.Sheets(3)
wsExcel.Name = "Flux complémentaires"

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


wsExcel.Columns(1).ColumnWidth = 8
wsExcel.Columns(2).ColumnWidth = 8
wsExcel.Columns(3).ColumnWidth = 50
wRow = 1
wsExcel.Cells(1, 1) = "Code": wsExcel.Cells(1, 1).Interior.Color = RGB(255, 170, 80)
wsExcel.Cells(1, 2) = "": wsExcel.Cells(1, 2).Interior.Color = RGB(255, 170, 80)
wsExcel.Cells(1, 3) = wsExcel.Name & " (" & dateImp10(DSys) & ")": wsExcel.Cells(1, 3).Interior.Color = RGB(255, 170, 80)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJOD' order by BIATABK1"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    Call rsYBIATAB0_GetBuffer(rsSab, xYBIATAB0)
    wRow = wRow + 1
    wsExcel.Cells(wRow, 1) = xYBIATAB0.BIATABK1
    wsExcel.Cells(wRow, 2) = xYBIATAB0.BIATABK2
    wsExcel.Cells(wRow, 3) = xYBIATAB0.BIATABTXT
    
    X = ":" & Trim(xYBIATAB0.BIATABK1) & " " & Trim(xYBIATAB0.BIATABK2) & ":" & Trim(mId$(xYBIATAB0.BIATABTXT, 11, 32))
    If mId$(xYBIATAB0.BIATABTXT, 1, 3) <> "000" Then
        lstW_Sorted.AddItem mId$(xYBIATAB0.BIATABTXT, 1, 3) & "MD" & X
    End If
    If mId$(xYBIATAB0.BIATABTXT, 5, 3) <> "000" Then
        lstW_Sorted.AddItem mId$(xYBIATAB0.BIATABTXT, 5, 3) & "MC" & X
    End If

    rsSab.MoveNext

Loop

'__________________________________________________________________________________
Set wsExcel = wbExcel.Sheets(4)
wsExcel.Name = "Flux Nostro"

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

wsExcel.Columns(1).ColumnWidth = 8
wsExcel.Columns(2).ColumnWidth = 8
wsExcel.Columns(3).ColumnWidth = 8
wsExcel.Columns(4).ColumnWidth = 50
wRow = 1
wsExcel.Cells(1, 1) = "Code": wsExcel.Cells(1, 1).Interior.Color = RGB(255, 170, 80)
wsExcel.Cells(1, 2) = ">> code": wsExcel.Cells(1, 2).Interior.Color = RGB(255, 170, 80)
wsExcel.Cells(1, 3) = "DB  CR": wsExcel.Cells(1, 3).Interior.Color = RGB(255, 170, 80)
wsExcel.Cells(1, 4) = wsExcel.Name & " (" & dateImp10(DSys) & ")": wsExcel.Cells(1, 4).Interior.Color = RGB(255, 170, 80)


xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJCPT' order by BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    Call rsYBIATAB0_GetBuffer(rsSab, xYBIATAB0)
    wRow = wRow + 1
    wsExcel.Cells(wRow, 1) = xYBIATAB0.BIATABK1
    If Trim(xYBIATAB0.BIATABK2) = "=" Then
        wsExcel.Cells(wRow, 2) = xYBIATAB0.BIATABTXT
    Else
        wsExcel.Cells(wRow, 3) = xYBIATAB0.BIATABTXT
    End If
    X = Trim(mId$(xYBIATAB0.BIATABK1, 1, 3))
    For K = 1 To arrOPE_Nb
        If X = Trim(arrOPE_K(K)) Then Exit For
    Next K
    wsExcel.Cells(wRow, 4) = arrOPE_Lib(K)
    
    If Trim(xYBIATAB0.BIATABK2) <> "=" Then

        X = ":" & Trim(xYBIATAB0.BIATABK1) & ":" & arrOPE_Lib(K)
        If mId$(xYBIATAB0.BIATABTXT, 1, 3) <> "000" Then
            lstW_Sorted.AddItem mId$(xYBIATAB0.BIATABTXT, 1, 3) & "#D" & X
        End If
        If mId$(xYBIATAB0.BIATABTXT, 5, 3) <> "000" Then
            lstW_Sorted.AddItem mId$(xYBIATAB0.BIATABTXT, 5, 3) & "#C" & X
        End If
    End If
    
    rsSab.MoveNext
Loop

'__________________________________________________________________________________
Set wsExcel = wbExcel.Sheets(5)
wsExcel.Name = "Flux = racine"

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

wsExcel.Columns(1).ColumnWidth = 8
wsExcel.Columns(2).ColumnWidth = 8: wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(3).ColumnWidth = 50
wRow = 1
wsExcel.Cells(1, 1) = "Code": wsExcel.Cells(1, 1).Interior.Color = RGB(255, 170, 80)
wsExcel.Cells(1, 2) = "Racine": wsExcel.Cells(1, 2).Interior.Color = RGB(255, 170, 80)

wsExcel.Cells(1, 3) = wsExcel.Name & " (" & dateImp10(DSys) & ")": wsExcel.Cells(1, 3).Interior.Color = RGB(255, 170, 80)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJCLI' order by BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call rsYBIATAB0_GetBuffer(rsSab, xYBIATAB0)
    wRow = wRow + 1
    wsExcel.Cells(wRow, 1) = xYBIATAB0.BIATABK1
    wsExcel.Cells(wRow, 2) = xYBIATAB0.BIATABK2
    xSQL = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & xYBIATAB0.BIATABK2 & "'"
    Set rsSabX = cnsab.Execute(xSQL)
    If Not rsSabX.EOF Then
        wsExcel.Cells(wRow, 3) = Trim(rsSabX("CLIENARA1"))
    Else
        wsExcel.Cells(wRow, 3) = "???"
    End If

    
    rsSab.MoveNext
Loop

'__________________________________________________________________________________
Set wsExcel = wbExcel.Sheets(6)
wsExcel.Name = "Récapitulatif"

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


wsExcel.Columns(1).ColumnWidth = 8: wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 8
wsExcel.Columns(3).ColumnWidth = 50
wRow = 1
wsExcel.Cells(1, 1) = "CB": wsExcel.Cells(1, 1).Interior.Color = RGB(255, 170, 80)
wsExcel.Cells(1, 2) = "Code": wsExcel.Cells(1, 2).Interior.Color = RGB(255, 170, 80)
wsExcel.Cells(1, 3) = wsExcel.Name & " (" & dateImp10(DSys) & ")": wsExcel.Cells(1, 3).Interior.Color = RGB(255, 170, 80)

For K = 0 To lstW_Sorted.ListCount - 1
    wRow = wRow + 1
    lstW_Sorted.ListIndex = K
    K1 = InStr(lstW_Sorted.Text, ":")
    Select Case mId$(lstW_Sorted.Text, K1 - 1, 1)
        Case "D": wColor = vbRed
        Case "C": wColor = vbBlue
        Case Else: wColor = vbBlack
    End Select
     Select Case mId$(lstW_Sorted.Text, K1 - 2, 1)
        Case " ":  wsExcel.Cells(wRow, 1) = mId$(lstW_Sorted.Text, 1, K1 - 3)
                   wsExcel.Cells(wRow, 1).Interior.Color = RGB(255, 255, 192): wsExcel.Cells(wRow, 1).Font.Bold = True
                   wsExcel.Cells(wRow, 2).Interior.Color = RGB(255, 255, 192): wsExcel.Cells(wRow, 1).Font.Bold = True
                   wsExcel.Cells(wRow, 3).Interior.Color = RGB(255, 255, 192): wsExcel.Cells(wRow, 1).Font.Bold = True
        Case Else: wsExcel.Cells(wRow, 1) = mId$(lstW_Sorted.Text, K1 - 2, 1)
    End Select
    wsExcel.Cells(wRow, 1).Font.Color = wColor
    
    K2 = InStr(K1 + 1, lstW_Sorted.Text, ":")
    wsExcel.Cells(wRow, 2) = mId$(lstW_Sorted.Text, K1 + 1, K2 - K1 - 1)
    wsExcel.Cells(wRow, 2).Font.Color = wColor
    K1 = Len(lstW_Sorted.Text)
    wsExcel.Cells(wRow, 3) = mId$(lstW_Sorted.Text, K2 + 1, K1 - K2)
    wsExcel.Cells(wRow, 3).Font.Color = wColor
Next K


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
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents

End Sub


Public Sub YFLUTPJ0_ECH_Export(blnInputBox As Boolean, wFile As String)
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFilex As String, wFile2 As String, xSQL As String
Dim wAMJMin As String, WAMJMax As String
Dim X As String, K As Long
'______________________________________________
Call DTPicker_Control(txtSelect_FLUTPJECH_Min, wAMJMin)

wFile = Trim("C:\Temp\Tréserorie Prévisionnelle Echéancier " & dateImp_Amj(DSys) & ".xlsx")
'______________________________________________
If blnInputBox Then
    X = InputBox("par défaut : " & wFile _
        & vbCrLf & vbCrLf & "     =========================" _
        & vbCrLf & "     =========================", "Trésorerie prévisionnelle Echéancier: nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If
'_________________________________________


If Dir(wFile) <> "" Then Kill wFile

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "Echéancier"
    .Subject = ""
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "Ech " & dateImp10(DSys)

YFLUTPJ0_Ech_Export_Recap

Set wsExcel = wbExcel.Sheets(2)
wsExcel.Name = "Detail " & dateImp10(DSys)

arrYFLUTPJ0_SQL mSelect_Where & " order by FLUTPJECH, FLUTPJCCB , FLUTPJDEV"
YFLUTPJ0_CB_Export_Detail

'__________________________________________________________________________________
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
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents

End Sub


Public Sub cmdSendMail_Ech()
Dim wSendMail As typeSendMail
Dim xDétail As String, xHeader As String, mbgColor As String
Dim K As Long, htmlFontColor_K As String
Dim iRow As Integer, iCol As Integer, X As String, xTD As String
Dim wForecolor As String, wBackColor As String, xColor As String
Dim xFLUTPJOPE As String, kLib As Integer
Dim xLnk As String

On Error Resume Next


xHeader = ""
For iRow = 0 To fgSelect.Rows - 1
    fgSelect.Row = iRow
    xTD = ""
    For iCol = 0 To fgSelect.Cols - 2
        If iCol <> 1 Then
            fgSelect.Col = iCol
            X = Trim(fgSelect.Text)
            If iRow = 0 Then
                wForecolor = cmdSendMail_Cell_Color(fgSelect.ForeColorFixed)
                wBackColor = cmdSendMail_Cell_Color(fgSelect.BackColorFixed)
                xTD = xTD _
                     & "<TD bgcolor=" & wBackColor & " width=100><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=" & wForecolor & "><B>" _
                     & X & "</B/TD>"
            Else
                wForecolor = cmdSendMail_Cell_Color(fgSelect.CellForeColor)
                wBackColor = cmdSendMail_Cell_Color(fgSelect.BackColor)
                xTD = xTD _
                     & "<TD bgcolor=" & wBackColor & "><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=" & wForecolor & ">" _
                     & X & "</TD>"
            End If
        End If
    Next iCol
    xHeader = xHeader & "<TR>" & xTD & "</TR>"

Next iRow

xDétail = ""
mbgColor = "bgcolor = #E0E0E0"

arrYFLUTPJ0_SQL mSelect_Where & " and FLUTPJORIG = '*' and FLUTPJSTA = ' ' order by  FLUTPJCCB, FLUTPJOPE, FLUTPJDEV, FLUTPJECH "
rsYFLUTPJ0_Init oldYFLUTPJ0

For K = 1 To arrYFLUTPJ0_Nb
    xYFLUTPJ0 = arrYFLUTPJ0(K)
    If xYFLUTPJ0.FLUTPJOPE = oldYFLUTPJ0.FLUTPJOPE Then
        xFLUTPJOPE = ""
    Else
        For kLib = 1 To arrOD_Nb
            If Trim(xYFLUTPJ0.FLUTPJOPE) = arrOD_K(kLib) Then
                xFLUTPJOPE = xYFLUTPJ0.FLUTPJOPE & " - " & arrOD_Lib(kLib)
                Exit For
            End If
        Next kLib
    End If
    If xYFLUTPJ0.FLUTPJCCB = oldYFLUTPJ0.FLUTPJCCB Then
        xYFLUTPJ0.FLUTPJCCB = 0
        If xYFLUTPJ0.FLUTPJSER = oldYFLUTPJ0.FLUTPJSER Then
            xYFLUTPJ0.FLUTPJSER = ""
            If xYFLUTPJ0.FLUTPJSSE = oldYFLUTPJ0.FLUTPJSSE Then
                xYFLUTPJ0.FLUTPJSSE = ""

            End If
        End If
    End If
    
    wBackColor = "#FFFFF0" 'htmlFontColor_Blue
    wForecolor = IIf(xYFLUTPJ0.FLUTPJMTD < 0, "#FF0000", "#0000FF")
    xDétail = xDétail & "<TR>" _
         & "<TD bgcolor=" & wBackColor & " width=50><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=#000000>" _
         & Format(xYFLUTPJ0.FLUTPJCCB, "###") & "</TD>" _
         & "<TD bgcolor=" & wBackColor & " width=100><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=#000000>" _
         & xYFLUTPJ0.FLUTPJSER & " " & xYFLUTPJ0.FLUTPJSSE & "</TD>" _
         & "<TD bgcolor=" & wBackColor & " width=300><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=#0000FF>" _
         & xFLUTPJOPE & "</TD>" _
         & "<TD bgcolor=" & wBackColor & " width=100><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=#000000>" _
         & dateImp10_S(xYFLUTPJ0.FLUTPJECH) & "</TD>" _
         & "<TD bgcolor=" & wBackColor & " width=200><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=" & wForecolor & ">" _
         & "<div align=" & Asc34 & "right" & Asc34 & ">" _
         & Format(xYFLUTPJ0.FLUTPJMTD, "### ### ### ##0.00") & "</div></TD>" _
         & "<TD bgcolor=" & wBackColor & " width=50><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=#000000>" _
         & xYFLUTPJ0.FLUTPJDEV & "</TD>" _
          & "<TD bgcolor=" & wBackColor & " width=50><span style='font-size:7.0pt;font-family:Arial Unicode MS'><Font color=#000000>" _
         & xYFLUTPJ0.FLUTPJORIG & " " & xYFLUTPJ0.FLUTPJEVE & " " & xYFLUTPJ0.FLUTPJSTA & "</TD>" _
        & "</TR>"
        
    oldYFLUTPJ0 = arrYFLUTPJ0(K)
    
    
Next K

wSendMail.From = currentSSIWINMAIL
wSendMail.AsHTML = True
wSendMail.Subject = "Flux de trésorerie prévisionnelle au " & dateImp10(YBIATAB0_DATE_CPT_JS1)

If blnAuto Then
    wSendMail.FromDisplayName = "@FLUX_TREPRE"
    wSendMail.RecipientDisplayName = "FLUX_TREPREV"
    wSendMail.Attachment = ""
    paramEditionNoPaper_Auto_PgmName = "BIA_PCI_COMPTE"
    Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S54", auto_YFLUTPJ0_mFile, "Prod", "BIA-FLUX-TREPREV")
    xLnk = htmlFontColor_Black & "<BR>" & paramEditionNoPaper_Auto_Lnk & "<BR>"
Else
    wSendMail.Recipient = currentSSIWINMAIL
    wSendMail.Attachment = auto_YFLUTPJ0_mFile
    xLnk = ""
End If


wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                    & "<span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & "<Font color = #404040>" _
                    & xLnk _
                    & "<BR><Font color = #303030>  <U>1 - Echéancier :</U><BR>" _
                    & "<BR>" & htmlFontColor_Blue & X _
                    & "<TABLE   width=1000 border=1 cellpadding=5 ></B>" _
                    & "<div align=" & Asc34 & "right" & Asc34 _
                    & xHeader _
                    & "</div></TABLE>" _
                    & "<BR>" _
                    & "<span style='font-size:8.0pt;font-family:Arial Unicode MS'>" & htmlFontColor_Gray _
                    & "<BR><Font color = #303030>  <U>2 - Opérations hors SAB :</U>" _
                    & "<BR><BR>" _
                    & "<TABLE   width=1000 border=1 cellpadding=5 ></B>" _
                    & "<div align=" & Asc34 & "left" & Asc34 _
                    & xDétail _
                    & "</div></TABLE>"



srvSendMail.Monitor wSendMail

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

Public Sub YFLUTPJ0_CB_Export_Recap()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSQL As String
'______________________________________________

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

wsExcel.Columns(1).ColumnWidth = 5
wsExcel.Columns(2).ColumnWidth = 50
For wCol = 3 To fgSelect.Cols - 1
    wsExcel.Columns(wCol).ColumnWidth = 12: wsExcel.Columns(wCol).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
Next wCol

For wRow = 0 To fgSelect.Rows - 1
    fgSelect.Row = wRow
    For wCol = 0 To fgSelect.Cols - 1
        fgSelect.Col = wCol
        If wCol < fgSelect.FixedCols Or wRow < fgSelect.FixedRows Then
            wsExcel.Cells(wRow + 1, wCol + 1).Interior.Color = colorHex_RGB(fgSelect.BackColorFixed)
            wsExcel.Cells(wRow + 1, wCol + 1).Font.Color = colorHex_RGB(fgSelect.ForeColorFixed)
            wsExcel.Cells(wRow + 1, wCol + 1) = fgSelect.Text
            If wRow = 0 Then
                If wCol = 1 Then wsExcel.Cells(wRow + 1, wCol + 1).Interior.Color = colorHex_RGB(fgSelect.CellBackColor)
                If wCol > 1 Then wsExcel.Cells(wRow + 1, wCol + 1).HorizontalAlignment = Excel.xlHAlignRight
            End If
       Else
            If Trim(fgSelect.Text) = "" Then
                wsExcel.Cells(wRow + 1, wCol + 1).Interior.Color = colorHex_RGB(fgSelect.BackColor)
            Else
                wsExcel.Cells(wRow + 1, wCol + 1).Interior.Color = colorHex_RGB(fgSelect.CellBackColor)
                wsExcel.Cells(wRow + 1, wCol + 1).Font.Color = colorHex_RGB(fgSelect.CellForeColor)
                wsExcel.Cells(wRow + 1, wCol + 1) = Val(fgSelect.Text)
            End If
        End If
    Next wCol
Next wRow

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub YFLUTPJ0_CB_Export_Refinancement()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSQL As String
'______________________________________________

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

wsExcel.Columns(1).ColumnWidth = 5
wsExcel.Columns(2).ColumnWidth = 40
For wCol = 3 To 6
    wsExcel.Columns(wCol).ColumnWidth = 25  ': wsExcel.Columns(wCol).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
Next wCol

For wRow = 0 To fgSelect.Rows - 1
    fgSelect.Row = wRow
    For wCol = 0 To 5
        fgSelect.Col = wCol
        If wCol < fgSelect.FixedCols Or wRow < fgSelect.FixedRows Then
            wsExcel.Cells(wRow + 1, wCol + 1).Interior.Color = colorHex_RGB(fgSelect.BackColorFixed)
            wsExcel.Cells(wRow + 1, wCol + 1).Font.Color = colorHex_RGB(fgSelect.ForeColorFixed)
            wsExcel.Cells(wRow + 1, wCol + 1) = fgSelect.Text
            If wRow = 0 Then
                If wCol = 1 Then wsExcel.Cells(wRow + 1, wCol + 1).Interior.Color = colorHex_RGB(fgSelect.CellBackColor)
                If wCol > 1 Then wsExcel.Cells(wRow + 1, wCol + 1).HorizontalAlignment = Excel.xlHAlignRight
            End If
       Else
            If Trim(fgSelect.Text) = "" Then
                wsExcel.Cells(wRow + 1, wCol + 1).Interior.Color = colorHex_RGB(fgSelect.BackColor)
            Else
                wsExcel.Cells(wRow + 1, wCol + 1).Interior.Color = colorHex_RGB(fgSelect.CellBackColor)
                wsExcel.Cells(wRow + 1, wCol + 1).Font.Color = colorHex_RGB(fgSelect.CellForeColor)
                wsExcel.Cells(wRow + 1, wCol + 1).HorizontalAlignment = Excel.xlHAlignRight
                wsExcel.Cells(wRow + 1, wCol + 1) = fgSelect.Text
                'Select Case wRow
                '    Case 1: wsExcel.Cells(wRow + 1, wCol + 1) = fgSelect.Text 'Format(Val(fgSelect.Text), "### ### ### ### ##0")
                '    Case 2: wsExcel.Cells(wRow + 1, wCol + 1) = fgSelect.Text 'Format(Val(fgSelect.Text), "### ### ##0")
                '    Case 3: wsExcel.Cells(wRow + 1, wCol + 1) = fgSelect.Text 'Format(Val(fgSelect.Text), "### ### ### ### ###.###")
                '    Case 4: wsExcel.Cells(wRow + 1, wCol + 1) = fgSelect.Text 'Format(Val(fgSelect.Text), "### ##0.00000")
                'End Select
            End If
        End If
    Next wCol
Next wRow

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub YFLUTPJ0_Ech_Export_Recap()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSQL As String
'______________________________________________

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

wsExcel.Columns(1).ColumnWidth = 10
wsExcel.Columns(2).ColumnWidth = 1
For wCol = 3 To fgSelect.Cols - 1
    wsExcel.Columns(wCol).ColumnWidth = 12: wsExcel.Columns(wCol).NumberFormat = "### ### ### ##0;[Red]-### ### ### ##0"
Next wCol

For wRow = 0 To fgSelect.Rows - 1
    fgSelect.Row = wRow
    For wCol = 0 To fgSelect.Cols - 1
        fgSelect.Col = wCol
        If wCol < fgSelect.FixedCols Or wRow < fgSelect.FixedRows Then
            wsExcel.Cells(wRow + 1, wCol + 1).Interior.Color = colorHex_RGB(fgSelect.BackColorFixed)
            wsExcel.Cells(wRow + 1, wCol + 1).Font.Color = colorHex_RGB(fgSelect.ForeColorFixed)
            wsExcel.Cells(wRow + 1, wCol + 1) = fgSelect.Text
            'If wCol = 1 And wRow = 0 Then wsExcel.Cells(wRow + 1, wCol + 1).Interior.Color = colorHex_RGB(fgSelect.CellBackColor)
            If wRow = 0 Then
                If wCol = 1 Then wsExcel.Cells(wRow + 1, wCol + 1) = "" 'wsExcel.Cells(wRow + 1, wCol + 1).Interior.Color = colorHex_RGB(fgSelect.CellBackColor)
                If wCol > 1 Then wsExcel.Cells(wRow + 1, wCol + 1).HorizontalAlignment = Excel.xlHAlignRight
            End If

       Else
            wsExcel.Cells(wRow + 1, wCol + 1).Interior.Color = colorHex_RGB(fgSelect.BackColor)
            wsExcel.Cells(wRow + 1, wCol + 1).Font.Color = colorHex_RGB(fgSelect.CellForeColor)
            wsExcel.Cells(wRow + 1, wCol + 1) = Val(fgSelect.Text)
        End If
    Next wCol
Next wRow

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub


Public Sub YFLUTPJ0_CB_Export_Detail()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long, K As Long, I As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSQL As String
Dim kColor As Integer
Dim xFLUTPJOPE As String, kLib As Integer
'______________________________________________

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
wsExcel.Cells(1, 1) = "CB": wsExcel.Columns(1).ColumnWidth = 5
wsExcel.Cells(1, 1).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Cells(1, 2) = "Service": wsExcel.Columns(2).ColumnWidth = 8
wsExcel.Cells(1, 3) = "Opération": wsExcel.Columns(3).ColumnWidth = 30
wsExcel.Cells(1, 4) = "Numéro": wsExcel.Columns(8).ColumnWidth = 10: wsExcel.Columns(4).NumberFormat = "### ### ### ##0"
wsExcel.Cells(1, 4).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Cells(1, 5) = "Echéance": wsExcel.Columns(5).ColumnWidth = 10 ': wsExcel.Columns(5).NumberFormat = "mm/dd/yyyy"
wsExcel.Cells(1, 5).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Cells(1, 6) = "Montant": wsExcel.Columns(6).ColumnWidth = 22: wsExcel.Columns(6).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(1, 6).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Cells(1, 7) = "Dev": wsExcel.Columns(7).ColumnWidth = 6
wsExcel.Cells(1, 8) = "Origine": wsExcel.Columns(8).ColumnWidth = 6
wsExcel.Cells(1, 9) = "Id": wsExcel.Columns(9).ColumnWidth = 15: wsExcel.Columns(9).NumberFormat = "### ### ### ##0"
wsExcel.Cells(1, 9).HorizontalAlignment = Excel.xlHAlignRight

For K = 1 To 9
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 170, 80)
Next K
kColor = 1
rsYFLUTPJ0_Init xYFLUTPJ0

For K = 1 To arrYFLUTPJ0_Nb
    wRow = K + 1
    If xYFLUTPJ0.FLUTPJECH <> arrYFLUTPJ0(K).FLUTPJECH Then
        kColor = IIf(kColor = 1, 2, 1)
    End If
    For I = 1 To 9
        If kColor = 1 Then
            wsExcel.Cells(wRow, I).Interior.Color = RGB(255, 255, 190)
        Else
            wsExcel.Cells(wRow, I).Interior.Color = RGB(255, 255, 230)
        End If
    Next I
    
    xYFLUTPJ0 = arrYFLUTPJ0(K)
'____________________________________________________________________________________
    If xYFLUTPJ0.FLUTPJORIG <> "*" Then
        xFLUTPJOPE = xYFLUTPJ0.FLUTPJOPE & " " & xYFLUTPJ0.FLUTPJNAT
    Else
        For kLib = 1 To arrOD_Nb
            If Trim(xYFLUTPJ0.FLUTPJOPE) = arrOD_K(kLib) Then
                xFLUTPJOPE = xYFLUTPJ0.FLUTPJOPE & " - " & arrOD_Lib(kLib)
                Exit For
            End If
        Next kLib
    End If

    
    wsExcel.Cells(wRow, 1) = xYFLUTPJ0.FLUTPJCCB
    wsExcel.Cells(wRow, 2) = xYFLUTPJ0.FLUTPJSER & " " & xYFLUTPJ0.FLUTPJSSE
    wsExcel.Cells(wRow, 3) = xFLUTPJOPE
    wsExcel.Cells(wRow, 4) = xYFLUTPJ0.FLUTPJDOS
    wsExcel.Cells(wRow, 5) = dateImp10(xYFLUTPJ0.FLUTPJECH) & "  "
    wsExcel.Cells(wRow, 6) = xYFLUTPJ0.FLUTPJMTD
    wsExcel.Cells(wRow, 7) = xYFLUTPJ0.FLUTPJDEV
    If xYFLUTPJ0.FLUTPJDEV <> "EUR" Then wsExcel.Cells(wRow, 7).Font.Color = vbMagenta: wsExcel.Cells(wRow, 3).Font.Color = vbMagenta
    wsExcel.Cells(wRow, 8) = xYFLUTPJ0.FLUTPJORIG & " " & xYFLUTPJ0.FLUTPJEVE & " " & xYFLUTPJ0.FLUTPJSTA
    wsExcel.Cells(wRow, 9) = xYFLUTPJ0.FLUTPJID
Next K

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub


Public Sub YFLUTPJ0_CB_Export_Refinancement_Detail()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long, K As Long, I As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSQL As String
Dim kColor As Integer
Dim xFLUTP2OPE As String, kLib As Integer
'______________________________________________

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
wsExcel.Cells(1, 1) = "CB": wsExcel.Columns(1).ColumnWidth = 5
wsExcel.Cells(1, 1).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Cells(1, 2) = "Service": wsExcel.Columns(2).ColumnWidth = 8
wsExcel.Cells(1, 3) = "Opération": wsExcel.Columns(3).ColumnWidth = 10
wsExcel.Cells(1, 4) = "Numéro": wsExcel.Columns(8).ColumnWidth = 5: wsExcel.Columns(4).NumberFormat = "### ### ### ##0"
wsExcel.Cells(1, 4).HorizontalAlignment = Excel.xlHAlignRight


wsExcel.Cells(1, 5) = "Date négo": wsExcel.Columns(5).ColumnWidth = 10 ': wsExcel.Columns(5).NumberFormat = "mm/dd/yyyy"
wsExcel.Cells(1, 5).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Cells(1, 6) = "Date mad": wsExcel.Columns(6).ColumnWidth = 10 ': wsExcel.Columns(6).NumberFormat = "mm/dd/yyyy"
wsExcel.Cells(1, 6).HorizontalAlignment = Excel.xlHAlignCenter

wsExcel.Cells(1, 7) = "Echéance": wsExcel.Columns(7).ColumnWidth = 10 ' : wsExcel.Columns(7).NumberFormat = "mm/dd/yyyy"
wsExcel.Cells(1, 7).HorizontalAlignment = Excel.xlHAlignCenter

wsExcel.Cells(1, 8) = "Nb J": wsExcel.Columns(8).ColumnWidth = 5: wsExcel.Columns(8).NumberFormat = "### ### ##0"
wsExcel.Cells(1, 8).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Cells(1, 9) = "MM": wsExcel.Columns(9).ColumnWidth = 4: wsExcel.Columns(9).NumberFormat = "### ### ##0"
wsExcel.Cells(1, 9).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Cells(1, 10) = "Taux BIA": wsExcel.Columns(10).ColumnWidth = 8: wsExcel.Columns(10).NumberFormat = "### ##0.00000"
wsExcel.Cells(1, 10).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Cells(1, 11) = "Taux EUR": wsExcel.Columns(11).ColumnWidth = 8: wsExcel.Columns(11).NumberFormat = "### ##0.00000"
wsExcel.Cells(1, 11).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Cells(1, 12) = "code taux": wsExcel.Columns(12).ColumnWidth = 8
wsExcel.Cells(1, 12).HorizontalAlignment = Excel.xlHAlignCenter

wsExcel.Cells(1, 13) = "Montant": wsExcel.Columns(13).ColumnWidth = 16: wsExcel.Columns(13).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(1, 13).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Cells(1, 14) = "Dev": wsExcel.Columns(14).ColumnWidth = 6
wsExcel.Cells(1, 15) = "Origine": wsExcel.Columns(15).ColumnWidth = 6
wsExcel.Cells(1, 16) = "Id": wsExcel.Columns(16).ColumnWidth = 10: wsExcel.Columns(16).NumberFormat = "### ### ### ##0"
wsExcel.Cells(1, 16).HorizontalAlignment = Excel.xlHAlignRight

For K = 1 To 16
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 170, 80)
Next K
kColor = 1
rsYFLUTP20_Init xYFLUTP20

For K = 1 To arrYFLUTP20_Nb
    wRow = K + 1
    If xYFLUTP20.FLUTP2ECH <> arrYFLUTP20(K).FLUTP2ECH Then
        kColor = IIf(kColor = 1, 2, 1)
    End If
    For I = 1 To 16
        If kColor = 1 Then
            'wsExcel.Cells(wRow, I).Interior.Color = RGB(255, 255, 190)
        Else
            wsExcel.Cells(wRow, I).Interior.Color = RGB(255, 255, 230)
        End If
    Next I
    
    xYFLUTP20 = arrYFLUTP20(K)
'____________________________________________________________________________________
        xFLUTP2OPE = xYFLUTP20.FLUTP2OPE & " " & xYFLUTP20.FLUTP2NAT

    
    wsExcel.Cells(wRow, 1) = xYFLUTP20.FLUTP2CCB
    wsExcel.Cells(wRow, 2) = xYFLUTP20.FLUTP2SER & " " & xYFLUTP20.FLUTP2SSE
    wsExcel.Cells(wRow, 3) = xFLUTP2OPE
    wsExcel.Cells(wRow, 4) = xYFLUTP20.FLUTP2DOS
    wsExcel.Cells(wRow, 5) = dateImp10(xYFLUTP20.FLUTP2NEG) & "  "
    wsExcel.Cells(wRow, 6) = dateImp10(xYFLUTP20.FLUTP2MAD) & "  "
    wsExcel.Cells(wRow, 7) = dateImp10(xYFLUTP20.FLUTP2ECH) & "  "
    wsExcel.Cells(wRow, 8) = xYFLUTP20.FLUTP2NBJ
    wsExcel.Cells(wRow, 9) = xYFLUTP20.FLUTP2ECHK
    wsExcel.Cells(wRow, 10) = xYFLUTP20.FLUTP2TX
    If xYFLUTP20.FLUTP2TXCB = 0 Then
        wsExcel.Cells(wRow, 11).Interior.Color = vbMagenta
    Else
        wsExcel.Cells(wRow, 11) = xYFLUTP20.FLUTP2TXCB
    End If
    wsExcel.Cells(wRow, 12) = xYFLUTP20.FLUTP2TXK
    
    wsExcel.Cells(wRow, 13) = xYFLUTP20.FLUTP2MTD
    wsExcel.Cells(wRow, 14) = xYFLUTP20.FLUTP2DEV
    wsExcel.Cells(wRow, 15) = xYFLUTP20.FLUTP2ORIG & " " & xYFLUTP20.FLUTP2EVE & " " & xYFLUTP20.FLUTP2STA
    wsExcel.Cells(wRow, 16) = xYFLUTP20.FLUTP2ID
Next K

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub



Private Function cmdFlux_Add()
Dim V, X As String, xSQL As String

'_________________________________________________________________________________
    xSQL = "select FLUTPJID from " & paramIBM_Library_SABSPE_XXX & ".YFLUTPJ0 " _
         & "  where FLUTPJORIG = '*' order by FLUTPJID desc"
    Set rsSab = cnsab.Execute(xSQL)
    
    If rsSab.EOF Then
        'Call MsgBox("Problème de lecture YFLUTPJ0 : FLUTPJID", vbCritical, "FLUX-TREPREV : ajout OD")
        newFlux.FLUTPJID = 1 ':GoTo Exit_Sub
    Else
        newFlux.FLUTPJID = rsSab("FLUTPJID") + 1
    End If
    
    
    xSQL = "select FLUTPJDOS from " & paramIBM_Library_SABSPE_XXX & ".YFLUTPJ0 " _
         & "  where FLUTPJORIG = '*' order by FLUTPJDOS desc"
    Set rsSab = cnsab.Execute(xSQL)
    
    If rsSab.EOF Then
        newFlux.FLUTPJDOS = 1
    
        'Call MsgBox("Problème de lecture YFLUTPJ0 : FLUTPJDOS", vbCritical, "FLUX-TREPREV : ajout OD")
        'GoTo Exit_Sub
    Else
        newFlux.FLUTPJDOS = rsSab("FLUTPJDOS") + 1
    End If
    newFlux.FLUTPJDOSQ = 1
    newYFLUTPJ1.FLUTPJDOS = newFlux.FLUTPJDOS
    newYFLUTPJ1.FLUTPJDOSQ = 0
    If newFlux.FLUTPJEVE = "U" Or newFlux.FLUTPJECH > mDate_J7 Then
        V = cmdFLUX_Transaction("Insert")
    Else
        V = cmdFLUX_Transaction("Insert+")
    End If

Exit_sub:
cmdFlux_Add = V

End Function



Private Sub cboSelect_FLUTPJOPE_Click()
cmdSelect_Reset

End Sub

Private Sub cboSelect_FLUTPJORIG_Click()
cmdSelect_Reset

End Sub

Private Sub chkSelect_FLUTPJDEV_Click()
cmdSelect_Reset

End Sub

Private Sub cmdFlux_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> cmdFlux_Ok ........"): DoEvents

If fgFlux.Visible Then cmdFlux_Reset
fgFlux.Visible = False

cmdFlux_SQL_1
    
Call lstErr_AddItem(lstErr, cmdContext, "< cmdflux_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
If cmdFlux_Ok.Visible Then cmdFlux_Ok.SetFocus

End Sub

Private Sub cmdFlux_Update_Quit_Click()
fraFlux_Update.Visible = False

End Sub







Private Sub cmdFlux_Update_Ok_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case mnuFlux_Option
    Case "mnuFlux_Add"
        If IsNull(fraflux_Control) Then newFlux.FLUTPJSTA = " ": cmdFlux_Add
        
    Case "mnuFlux_Delete":
        X = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YFLUTPJ0 " _
            & " where FLUTPJORIG = '*' and FLUTPJDOS = " & oldFlux.FLUTPJDOS
        Set rsSab = cnsab.Execute(X)
        If rsSab("Tally") > 1 Then
            cmdFLUX_Transaction ("Delete")
        Else
            cmdFLUX_Transaction ("Delete_YFLUTPJ1")
        End If
    
    Case "mnuFlux_Update"
        If IsNull(fraflux_Control) Then cmdFLUX_Transaction ("Update")
        
    Case "mnuFlux_Annulation"
        newFlux = oldFlux
        Select Case oldFlux.FLUTPJSTA
            Case "A": newFlux.FLUTPJSTA = " ": cmdFLUX_Transaction ("Update")
            Case " ": newFlux.FLUTPJSTA = "A": cmdFLUX_Transaction ("Update")
        End Select

    Case "mnuFlux_Close":
                        newFlux = oldFlux
                        newFlux.FLUTPJSTA = "X"
                        cmdFLUX_Transaction ("Close")
    Case "mnuFlux_FLUTPJTXT"
        If IsNull(fraflux_Control) Then cmdFLUX_Transaction ("Update_FLUTPJTXT")
End Select
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdParam_Ok_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

Select Case mnuParam_Option
    Case "mnuParam_Add"
        If IsNull(parametrage_Control) Then
            If IsNull(Parametrage_New) Then lstParam_Action_Click
        End If
    Case "mnuParam_Delete"
        If IsNull(Parametrage_Delete) Then lstParam_Action_Click
    Case "mnuParam_Update"
        If IsNull(parametrage_Control) Then
            If IsNull(Parametrage_Update) Then lstParam_Action_Click
        End If
End Select
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdParam_Quit_Click()
fraParam_K.Visible = False

End Sub



'______________________________________________________________________
Private Sub fgSelect_Display()
Dim wColor As Long

Dim I As Long, K As Long
Dim wDev As String, wCCB As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wMTD As Currency

On Error GoTo Error_Handler
currentAction = "fgSelect_Display"

SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset


fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Row = 0
fgSelect.Col = 1: fgSelect.Text = "Période du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax)
fgSelect.CellBackColor = vbMagenta

If chkSelect_FLUTPJDEV.Value = "1" Then
    arrDev_Cours_Load (wAMJMin)
    libSelect_FLUTPJDEV = "cours au " & dateImp10(wAMJMin)
    libSelect_FLUTPJDEV.Visible = True
    For K = 1 To arrDev_Nb
        fgSelect.Col = K + 1
        fgSelect.Text = arrDev_Cours(K) & "  " & arrDev(K)
    Next K

End If


For K = 1 To arrCCB_Nb
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect.Col = 0: fgSelect.Text = arrCCB_K(K)
    fgSelect.Col = 1: fgSelect.Text = arrCCB_Lib(K)
    fgSelect.CellForeColor = RGB(0, 0, 64)
    Select Case arrCCB_K(K)
        Case "100": mCCB_DB_Row = fgSelect.Row: fgSelect.CellBackColor = RGB(255, 224, 176)
        Case "200": mCCB_CR_Row = fgSelect.Row: fgSelect.CellBackColor = RGB(255, 224, 176)
        Case "300": mCCB_SD_Row = fgSelect.Row: fgSelect.CellBackColor = RGB(255, 192, 144)
        Case Else: fgSelect.CellBackColor = RGB(255, 255, 230)
    End Select
    'fgSelect.CellBackColor = RGB(0, 164, 164)
Next K
For K = 1 To arrDev_Nb
    arrDev_DB(K) = 0: arrDev_CR(K) = 0
Next K

    
Do While Not rsSab.EOF
    wCCB = rsSab(0)
    wDev = rsSab(1)
    wMTD = rsSab(2)
    For K = 1 To arrDev_Nb
        If wDev = arrDev(K) Then
            fgSelect.Col = K + 1
            If chkSelect_FLUTPJDEV.Value = "1" Then wMTD = Round(wMTD * arrDev_Cours(K), 2)
            Exit For
        End If
    Next K
    If wCCB < "200" Then
        arrDev_DB(K) = arrDev_DB(K) + wMTD
    Else
        arrDev_CR(K) = arrDev_CR(K) + wMTD
    End If
    For K = 1 To arrCCB_Nb
        If wCCB = arrCCB_K(K) Then fgSelect.Row = K: Exit For
    Next K
     fgSelect.Text = Format$(wMTD, "### ### ### ###")
    fgSelect.CellForeColor = IIf(wMTD < 0, vbRed, vbBlue)
    fgSelect.CellBackColor = RGB(255, 255, 235)
    rsSab.MoveNext
Loop

For K = 1 To arrDev_Nb
    fgSelect.Col = K + 1
    fgSelect.Row = mCCB_DB_Row
    fgSelect.CellBackColor = RGB(255, 224, 176)
    fgSelect.Text = Format$(arrDev_DB(K), "### ### ### ##0")
    fgSelect.CellForeColor = IIf(arrDev_DB(K) < 0, vbRed, vbBlue)
    fgSelect.CellFontBold = True
    
    fgSelect.Row = mCCB_CR_Row
    fgSelect.CellBackColor = RGB(255, 224, 176)
    fgSelect.Text = Format$(arrDev_CR(K), "### ### ### ##0")
    fgSelect.CellForeColor = IIf(arrDev_CR(K) < 0, vbRed, vbBlue)
    fgSelect.CellFontBold = True

    wMTD = arrDev_DB(K) + arrDev_CR(K)
    fgSelect.Row = mCCB_SD_Row
    fgSelect.CellBackColor = RGB(255, 192, 144)
    fgSelect.Text = Format$(wMTD, "### ### ### ##0")
    fgSelect.CellForeColor = IIf(wMTD < 0, vbRed, vbBlue)
    fgSelect.CellFontBold = True
Next K

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : "): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgSelect_Display_3()
Dim wColor As Long

Dim I As Long, K As Long
Dim wDev As String, wCCB As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim sMTD As Currency, sNb As Long, sDurée As Long, sEcart As Long

On Error GoTo Error_Handler
currentAction = "fgSelect_Display"

SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset


fgSelect.Rows = 1
fgSelect.FormatString = "<CB  |<                                                       |> col2|> col3|> col4|> col5||||||||"
fgSelect.FormatString = Replace(fgSelect.FormatString, "col2", "Financements jour le jour         ")
fgSelect.FormatString = Replace(fgSelect.FormatString, "col3", "Financements > 1 jour et -/= 3 mois")
fgSelect.FormatString = Replace(fgSelect.FormatString, "col4", "Financements >  3 mois et -/= 1 an ")
fgSelect.FormatString = Replace(fgSelect.FormatString, "col5", "Financements > 1 an   ")

'fgSelect.FormatString = "<CB  |                                            |> Financements jour le jour|> Financements > 1 jour et = 3 mois|> Financements <= 1 an|> Financement > 1 an"
fgSelect.Col = 1: fgSelect.Text = "Période du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax)
fgSelect.CellBackColor = vbMagenta

For I = 1 To 5
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    Select Case I
        Case 1
            fgSelect.Col = 0: fgSelect.Text = "1.1"
            fgSelect.Col = 1: fgSelect.Text = "Montant total des financements"
            For K = 1 To 4
                fgSelect.Col = K + 1: fgSelect.Text = Format$(arrSINFO_LIQU(K).MTD, "### ### ### ### ##0")
                fgSelect.CellForeColor = vbBlue: fgSelect.CellBackColor = RGB(255, 255, 230)
            Next K
        Case 2
            fgSelect.Col = 0: fgSelect.Text = "1.2"
            fgSelect.Col = 1: fgSelect.Text = "Nombre de financement"
            For K = 1 To 4
                fgSelect.Col = K + 1: fgSelect.Text = Format$(arrSINFO_LIQU(K).Nb, "### ### ### ### ##0")
                 fgSelect.CellForeColor = vbBlue: fgSelect.CellBackColor = RGB(255, 255, 230)
           Next K
        Case 3
            fgSelect.Col = 0: fgSelect.Text = "1.3"
            fgSelect.Col = 1: fgSelect.Text = "Durée moyenne pondérée"
            For K = 1 To 4
                fgSelect.Col = K + 1: fgSelect.Text = Format$(arrSINFO_LIQU(K).Durée, "### ### ###.###")
                fgSelect.CellForeColor = vbBlue: fgSelect.CellBackColor = RGB(255, 255, 230)
            Next K
        Case 4
            fgSelect.Col = 0: fgSelect.Text = "1.4"
            fgSelect.Col = 1: fgSelect.Text = "Ecart moyen"
            For K = 1 To 4
                fgSelect.Col = K + 1: fgSelect.Text = Format$(arrSINFO_LIQU(K).Ecart, "### ### ##0.00000")
                 fgSelect.CellForeColor = vbBlue: fgSelect.CellBackColor = RGB(255, 255, 230)
           Next K
        Case 5
            fgSelect.Col = 0: fgSelect.Text = "xxx"
            fgSelect.Col = 1: fgSelect.Text = "nb jours moyen pondéré"
            For K = 1 To 4
                fgSelect.Col = K + 1: fgSelect.Text = Format$(arrSINFO_LIQU(K).NBJ, "### ### ### ### ##0")
                fgSelect.CellForeColor = vbBlue: fgSelect.CellBackColor = RGB(255, 255, 230)
            Next K
        End Select
Next I

    



fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : "): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgSelect_Display_Ech()
Dim wColor As Long

Dim I As Long, K As Long
Dim wDev As String, wCCB As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wMTD As Currency
Dim wECH As Long
On Error GoTo Error_Handler
currentAction = "fgSelect_Display"

SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = Replace(fgSelect_FormatString, "<CB  |<                                                       ", "<Echéance      |")

fgSelect.Row = 0
For K = 1 To arrDev_Nb
    arrDev_DB(K) = 0: arrDev_CR(K) = 0
Next K

wECH = 0
Do While Not rsSab.EOF
    If wECH <> rsSab(0) Then
        fgSelect.Rows = fgSelect.Rows + 1
         fgSelect.Row = fgSelect.Rows - 1
       wECH = rsSab(0)
       fgSelect.Col = 0
       fgSelect.Text = dateImp10(wECH)
       
    End If
    
    wDev = rsSab(1)
    wMTD = rsSab(2)
    For K = 1 To arrDev_Nb
        If wDev = arrDev(K) Then fgSelect.Col = K + 1: Exit For
    Next K
     fgSelect.Text = Format$(wMTD, "### ### ### ##0")
    fgSelect.CellForeColor = IIf(wMTD < 0, vbRed, vbBlue)
   
    rsSab.MoveNext
Loop



fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : "): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub arrYFLUTPJ0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYFLUTPJ0(101)
arrYFLUTPJ0_Max = 100: arrYFLUTPJ0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YFLUTPJ0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYFLUTPJ0_GetBuffer(rsSab, xYFLUTPJ0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYFLUTPJ0.fgselect_Display"
        '' Exit Sub
     Else
         arrYFLUTPJ0_Nb = arrYFLUTPJ0_Nb + 1
         If arrYFLUTPJ0_Nb > arrYFLUTPJ0_Max Then
             arrYFLUTPJ0_Max = arrYFLUTPJ0_Max + 100
             ReDim Preserve arrYFLUTPJ0(arrYFLUTPJ0_Max)
         End If
         
         arrYFLUTPJ0(arrYFLUTPJ0_Nb) = xYFLUTPJ0
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
Private Sub arrYFLUTP20_SQL(xWhere As String)
Dim V
Dim nbErr As Long
Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYFLUTP20(101)

SINFO_LIQU_Init arrSINFO_LIQU(1)
arrSINFO_LIQU(2) = arrSINFO_LIQU(1)
arrSINFO_LIQU(3) = arrSINFO_LIQU(1)
arrSINFO_LIQU(4) = arrSINFO_LIQU(1)
rsYFLUTP20_Init oldYFLUTP20
nbErr = 0
arrYFLUTP20_Max = 100: arrYFLUTP20_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YFLUTP20 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYFLUTP20_GetBuffer(rsSab, xYFLUTP20)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYFLUTP20.fgselect_Display"
        '' Exit Sub
     Else
         arrYFLUTP20_Nb = arrYFLUTP20_Nb + 1
         If arrYFLUTP20_Nb > arrYFLUTP20_Max Then
             arrYFLUTP20_Max = arrYFLUTP20_Max + 100
             ReDim Preserve arrYFLUTP20(arrYFLUTP20_Max)
         End If
         
         xYFLUTP20.FLUTP2MTD = -xYFLUTP20.FLUTP2MTD
         
         arrYFLUTP20(arrYFLUTP20_Nb) = xYFLUTP20
         
         Select Case xYFLUTP20.FLUTP2ECHK
            Case 0: K = 1
                If oldYFLUTP20.FLUTP2NEG <> xYFLUTP20.FLUTP2NEG Then arrSINFO_LIQU(K).Nb = arrSINFO_LIQU(K).Nb + 1
                arrSINFO_LIQU(K).Ecart_S = arrSINFO_LIQU(K).Ecart_S + xYFLUTP20.FLUTP2MTD * (xYFLUTP20.FLUTP2TX - xYFLUTP20.FLUTP2TXCB)
           Case Is <= 3: K = 2
                arrSINFO_LIQU(K).Nb = arrSINFO_LIQU(K).Nb + 1
                arrSINFO_LIQU(K).Ecart_S = arrSINFO_LIQU(K).Ecart_S + xYFLUTP20.FLUTP2MTD * (xYFLUTP20.FLUTP2TX - xYFLUTP20.FLUTP2TXCB)
           Case Is <= 12: K = 3
                arrSINFO_LIQU(K).Nb = arrSINFO_LIQU(K).Nb + 1
                arrSINFO_LIQU(K).Ecart_S = arrSINFO_LIQU(K).Ecart_S + xYFLUTP20.FLUTP2MTD * (xYFLUTP20.FLUTP2TX - xYFLUTP20.FLUTP2TXCB)
           Case Else: K = 4
                arrSINFO_LIQU(K).Nb = arrSINFO_LIQU(K).Nb + 1
                arrSINFO_LIQU(K).Durée_S = arrSINFO_LIQU(K).Durée_S + xYFLUTP20.FLUTP2MTD * xYFLUTP20.FLUTP2NBJ
                arrSINFO_LIQU(K).Ecart_S = arrSINFO_LIQU(K).Ecart_S + xYFLUTP20.FLUTP2MTD * xYFLUTP20.FLUTP2NBJ * (xYFLUTP20.FLUTP2TX - xYFLUTP20.FLUTP2TXCB)
      End Select
        
        arrSINFO_LIQU(K).MTD = arrSINFO_LIQU(K).MTD + xYFLUTP20.FLUTP2MTD
        arrSINFO_LIQU(K).NBJ = arrSINFO_LIQU(K).NBJ + xYFLUTP20.FLUTP2MTD * xYFLUTP20.FLUTP2NBJ
        If xYFLUTP20.FLUTP2TXCB = 0 Then nbErr = nbErr + 1
        oldYFLUTP20 = xYFLUTP20
    End If
    rsSab.MoveNext
Loop
For K = 1 To 4
    If arrSINFO_LIQU(K).MTD <> 0 Then
        arrSINFO_LIQU(K).NBJ = arrSINFO_LIQU(K).NBJ / arrSINFO_LIQU(K).MTD
        If K = 4 Then
            arrSINFO_LIQU(K).Ecart = arrSINFO_LIQU(K).Ecart_S / arrSINFO_LIQU(K).Durée_S
            arrSINFO_LIQU(K).Durée = arrSINFO_LIQU(K).Durée_S / arrSINFO_LIQU(K).MTD / 365
       Else
            arrSINFO_LIQU(K).Ecart = arrSINFO_LIQU(K).Ecart_S / arrSINFO_LIQU(K).MTD
        End If
    End If
Next K
    If Not blnAuto And nbErr > 0 Then MsgBox "Il y a " & nbErr & " opérations sans valeur de taux EURJ**", vbCritical, Me.Name & " :  "

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub arrfgFlux_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrfgFlux(101)
arrfgFlux_Max = 100: arrfgFlux_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YFLUTPJ0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYFLUTPJ0_GetBuffer(rsSab, xFlux)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmfgFlux.fgselect_Display"
        '' Exit Sub
     Else
         arrfgFlux_Nb = arrfgFlux_Nb + 1
         If arrfgFlux_Nb > arrfgFlux_Max Then
             arrfgFlux_Max = arrfgFlux_Max + 100
             ReDim Preserve arrfgFlux(arrfgFlux_Max)
         End If
         
         arrfgFlux(arrfgFlux_Nb) = xFlux
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


Private Sub arrfgFlux_Detail_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrfgFlux_Detail(101)
arrfgFlux_Detail_Max = 100: arrfgFlux_Detail_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YFLUTPJ0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYFLUTPJ0_GetBuffer(rsSab, xFlux)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmfgFlux_Detail.fgselect_Display"
        '' Exit Sub
     Else
         arrfgFlux_Detail_Nb = arrfgFlux_Detail_Nb + 1
         If arrfgFlux_Detail_Nb > arrfgFlux_Detail_Max Then
             arrfgFlux_Detail_Max = arrfgFlux_Detail_Max + 100
             ReDim Preserve arrfgFlux_Detail(arrfgFlux_Detail_Max)
         End If
         
         arrfgFlux_Detail(arrfgFlux_Detail_Nb) = xFlux
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
    lstErr.Clear
    fgSelect.Visible = False
    fraDetail.Visible = False
    chkSelect_FLUTPJDEV.Visible = False
    libSelect_FLUTPJDEV.Visible = False
    lstW.Visible = False
    cmdSelect_Ok.Visible = False 'True
    fraSelect_Option_2.Visible = False

    cmdSelect_SQL_K = Trim(mId$(cboSelect_SQL, 1, 2))
    Select Case cmdSelect_SQL_K
        Case "1":
            Call DTPicker_Set(txtSelect_FLUTPJECH_Max, arrAMJMAX(1)) '
            Call DTPicker_Set(txtSelect_FLUTPJECH_Min, arrAMJMIN(1)) '

            fraSelect_Options.Visible = True
            chkSelect_FLUTPJDEV.Visible = True
            cmdSelect_Ok.Visible = True
            'cmdSelect_Ok_Click
        Case "2":
            Call DTPicker_Set(txtSelect_FLUTPJECH_Max, arrAMJMAX(2)) '
            Call DTPicker_Set(txtSelect_FLUTPJECH_Min, arrAMJMIN(2)) '
            fraSelect_Options.Visible = True
            fraSelect_Option_2.Visible = True
            cmdSelect_Ok.Visible = True
        Case "3":
            Call DTPicker_Set(txtSelect_FLUTPJECH_Max, arrAMJMAX(3)) '
            Call DTPicker_Set(txtSelect_FLUTPJECH_Min, arrAMJMIN(3)) '
            cmdSelect_Ok.Visible = True
        Case Else
            cmdSelect_Ok.Visible = True
    End Select

End If

End Sub

Public Sub cmdFlux_Reset()
If blnControl Then
    lstErr.Clear
    fgFlux.Visible = False
    fraFlux_Detail.Visible = False
    fraFlux_Update.Visible = False

End If

End Sub
Public Sub cmdDetail_Reset()
If blnControl Then
    lstErr.Clear
    If fraDetail.Visible Then
        fraDetail.Visible = False
        fraFlux_Detail.Visible = False
        fgFlux.Visible = False
        fgDetail_Display
    End If
End If

End Sub


Private Sub cmdSelect_SQL_1()
Dim V
Dim xSQL As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdYFLUTPJ0_SQL"
blnOk = False
Call DTPicker_Control(txtSelect_FLUTPJECH_Min, wAMJMin)
Call DTPicker_Control(txtSelect_FLUTPJECH_Max, WAMJMax)
'If wAmjMin < 20100000 Then
'    V = "la date de début doit être supérieure au 01-01-2010"
'    GoTo Error_MsgBox
'End If
If wAMJMin > WAMJMax Then
    V = "la date de début doit être inférieure à la date de fin"
    GoTo Error_MsgBox
End If

mSelect_Where = " Where FLUTPJECH >= " & wAMJMin & " and FLUTPJECH <= " & WAMJMax
        xWhere = " and FLUTPJSTA = ' '"
xSQL = "select FLUTPJCCB , FLUTPJDEV , SUM(FLUTPJMTD) from " & paramIBM_Library_SABSPE & ".YFLUTPJ0" _
     & mSelect_Where & xWhere _
     & " group by  FLUTPJCCB , FLUTPJDEV" _
     & " order by FLUTPJCCB , FLUTPJDEV"
Set rsSab = cnsab.Execute(xSQL)

fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_Ech()
Dim V
Dim xSQL As String
Dim xWhere As String, xAnd As String, xAnd_FLUTPJOPE As String
Dim wCli As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdYFLUTPJ0_SQL"
blnOk = False
Call DTPicker_Control(txtSelect_FLUTPJECH_Min, wAMJMin)
Call DTPicker_Control(txtSelect_FLUTPJECH_Max, WAMJMax)

X = Trim(mId$(cboSelect_FLUTPJORIG, 1, 2))
Select Case X
    Case "*": xAnd = " and FLUTPJORIG = '*'"
    Case "1": xAnd = " and FLUTPJID >= 1000000 and FLUTPJID <= 1999999"
    Case "1*": xAnd = " and FLUTPJID >= 1000000 and FLUTPJID <= 1999999 and FLUTPJNAT not in ('EJJ','EJD','PJJ','PJD')"
    Case "2": xAnd = " and FLUTPJID >= 2000000 and FLUTPJID <= 2999999"
    Case "3": xAnd = " and FLUTPJID >= 3000000 and FLUTPJID <= 3999999"
    Case "4": xAnd = " and FLUTPJID >= 4000000 and FLUTPJID <= 4999999"
    Case "5": xAnd = " and FLUTPJID >= 5000000 and FLUTPJID <= 5999999"
    Case "6": xAnd = " and FLUTPJID >= 6000000 and FLUTPJID <= 6999999"
    Case "7": xAnd = " and FLUTPJID >= 7000000 and FLUTPJID <= 7999999"
    Case Else: xAnd = ""
End Select
X = Trim(cboSelect_FLUTPJOPE)
Select Case X
    Case "": xAnd_FLUTPJOPE = ""
    Case Else: xAnd_FLUTPJOPE = " and FLUTPJOPE = '" & X & "'"
End Select


mSelect_Where = " Where FLUTPJECH >= " & wAMJMin & " and FLUTPJECH <= " & WAMJMax & xAnd & xAnd_FLUTPJOPE
xSQL = "select FLUTPJECH , FLUTPJDEV , SUM(FLUTPJMTD) from " & paramIBM_Library_SABSPE & ".YFLUTPJ0" _
     & mSelect_Where _
     & " group by  FLUTPJECH , FLUTPJDEV" _
     & " order by FLUTPJECH , FLUTPJDEV"
Set rsSab = cnsab.Execute(xSQL)


fgSelect_Display_Ech

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdFlux_SQL_1()
Dim V, X As String
Dim xSQL As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdYFLUTPJ0_SQL"
blnOk = False
Call DTPicker_Control(txtFlux_FLUTPJECH_Min, X)
Call rsYFLUTPJ0_Init(oldFlux)
oldFlux.FLUTPJORIG = "*"
oldFlux.FLUTPJECH = YBIATAB0_DATE_CPT_JS1

xWhere = " Where FLUTPJORIG = '*' and FLUTPJECH >= " & X
X = Trim(cboFlux_FLUTPJDEV)
If X <> "" Then oldFlux.FLUTPJDEV = X: xWhere = xWhere & "   and FLUTPJDEV = '" & X & "'"
X = mId$(Trim(cboFlux_FLUTPJOD), 1, 3)
If X <> "" Then oldFlux.FLUTPJOPE = X: xWhere = xWhere & "   and FLUTPJOPE = '" & X & "'"

arrfgFlux_SQL xWhere & " order by FLUTPJCCB ,FLUTPJOPE , FLUTPJDOS , FLUTPJDOSQ"

fgFlux_Display

If arrfgFlux_Nb = 0 Then
    If Trim(oldFlux.FLUTPJDEV) = "" Or Trim(oldFlux.FLUTPJOPE) = "" Then
        Call MsgBox("Pour créer un  nouveau flux, préciser :" & vbCrLf & " - la devise" & vbCrLf & " - le code opération", vbExclamation, "FLUX_TREPREV")
    Else
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJOD' and BIATABK1 = '" & Trim(oldFlux.FLUTPJOPE) & "'"
        Set rsSab = cnsab.Execute(xSQL)
        
        If rsSab.EOF Then
            Call MsgBox("erreur lecture : " & vbCrLf & xSQL, vbCritical, currentAction)
        Else
            Call rsYBIATAB0_GetBuffer(rsSab, paramOD)
            oldFlux.FLUTPJEVE = mId$(paramOD.BIATABTXT, 9, 1)
            If mId$(paramOD.BIATABTXT, 1, 3) = "000" Then
                optFlux_Encaissement.Value = True
            Else
                optFlux_Decaissement.Value = True
            End If
            
            fraFlux_Display
            fraFlux_Display_FLUTPJCLI
            mnuFlux_Add_Click
            'mnuFlux_Option = "mnuFlux_Add"

            'fraFlux_Update.Caption = "Nouveau dossier"
            ''cmdFlux_Update_Ok.Visible = True
            'fraFlux_Update_A.Enabled = True
            'fraFlux_Update_B.Enabled = True
        End If
    End If
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_3()
Dim V, X As String
Dim xSQL As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdFlux_SQL_3"
blnOk = False
Call DTPicker_Control(txtSelect_FLUTPJECH_Min, wAMJMin)
Call DTPicker_Control(txtSelect_FLUTPJECH_Max, WAMJMax)

If wAMJMin > WAMJMax Then
    V = "la date de début doit être inférieure à la date de fin"
    GoTo Error_MsgBox
End If
arrAMJMAX(3) = WAMJMax
arrAMJMIN(3) = wAMJMin

mSelect_Where = " Where FLUTP2NEG >= " & wAMJMin & " and FLUTP2NEG <= " & WAMJMax

arrYFLUTP20_SQL mSelect_Where & " order by FLUTP2ECHK ,FLUTP2NEG "   ' ne pas modifier ORDER

fgSelect_Display_3

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
fraDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.Row = 0

currentAction = "fgDetail_Display"

For I = 1 To arrYFLUTPJ0_Nb
         
    xYFLUTPJ0 = arrYFLUTPJ0(I)
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_DisplayLine I
    
Next I

fraDetail.Visible = True:


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgDetail_Display_3()
Dim wColor As Long
Dim X As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
fraDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = "<Opération            |>Montant               |>Négociation   |>Mise à dispo  |>Echéance      |>NBJ |> MM|> Taux        |> Taux réf CB|<Code Tx   "
fgDetail.Row = 0

currentAction = "fgDetail_Display_3"

For I = 1 To arrYFLUTP20_Nb
         
    xYFLUTP20 = arrYFLUTP20(I)
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_DisplayLine_3 I
    
Next I

fraDetail.Visible = True:


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgFlux_Detail_Display()
Dim wColor As Long
Dim X As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
fraFlux_Detail.Visible = False
fgFlux_Detail_Reset

fgFlux_Detail.Rows = 1
fgFlux_Detail.FormatString = fgFlux_Detail_FormatString
fgFlux_Detail.Row = 0

currentAction = "fgFlux_Detail_Display"



For I = 1 To arrfgFlux_Detail_Nb
         
    xFlux = arrfgFlux_Detail(I)
    fgFlux_Detail.Rows = fgFlux_Detail.Rows + 1
    fgFlux_Detail.Row = fgFlux_Detail.Rows - 1
    fgFlux_Detail_DisplayLine I
    
   
Next I

fraFlux_Detail.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgFlux_Display()
Dim wColor As Long, K As Long

On Error GoTo Error_Handler
fgFlux.Visible = False
fgFlux_Reset

fgFlux.Rows = 1
fgFlux.FormatString = fgFlux_FormatString
fgFlux.Row = 0

currentAction = "fgFlux_Display"

rsYFLUTPJ0_Init xFlux

For K = 1 To arrfgFlux_Nb
    If xFlux.FLUTPJOPE <> arrfgFlux(K).FLUTPJOPE Or xFlux.FLUTPJDOS <> arrfgFlux(K).FLUTPJDOS Then
        xFlux = arrfgFlux(K)
             
        fgFlux.Rows = fgFlux.Rows + 1
        fgFlux.Row = fgFlux.Rows - 1
        fgFlux_DisplayLine K
        

    End If
    
Next K

fgFlux.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub



Public Sub fgDetail_DisplayLine(lIndex As Long)
Dim K As Integer, blnOk As Boolean
Dim wColor As Long, wColor_Row As Long
Dim blnSolde As Boolean

On Error Resume Next

If xYFLUTPJ0.FLUTPJSTA = " " Then
    blnOk = True
    wColor = vbBlue
Else
    blnOk = False
    wColor = RGB(128, 128, 128)
End If

fgDetail.Col = 0: fgDetail.Text = xYFLUTPJ0.FLUTPJCCB
fgDetail.CellForeColor = wColor
fgDetail.Col = 1: fgDetail.Text = xYFLUTPJ0.FLUTPJSER & " " & xYFLUTPJ0.FLUTPJSSE
fgDetail.CellForeColor = wColor
fgDetail.Col = 2: fgDetail.Text = xYFLUTPJ0.FLUTPJOPE & " " & xYFLUTPJ0.FLUTPJNAT & " " & Format$(xYFLUTPJ0.FLUTPJDOS, "@@@ @@@ @@@")
fgDetail.CellForeColor = wColor
fgDetail.Col = 3: fgDetail.Text = dateImp10(xYFLUTPJ0.FLUTPJECH) & "  "
fgDetail.CellForeColor = wColor
fgDetail.Col = 4: fgDetail.Text = Format$(xYFLUTPJ0.FLUTPJMTD, "### ### ### ###.00")
If blnOk Then
    fgDetail.CellForeColor = IIf(xYFLUTPJ0.FLUTPJMTD < 0, vbRed, vbBlue)
Else
    fgDetail.CellForeColor = wColor
End If
fgDetail.Col = 5: fgDetail.Text = xYFLUTPJ0.FLUTPJDEV
If blnOk Then
    fgDetail.CellForeColor = IIf(xYFLUTPJ0.FLUTPJDEV = "EUR", vbBlue, vbMagenta)
Else
    fgDetail.CellForeColor = wColor
End If
fgDetail.CellForeColor = wColor
fgDetail.Col = 6: fgDetail.Text = xYFLUTPJ0.FLUTPJORIG & " " & xYFLUTPJ0.FLUTPJEVE & " " & xYFLUTPJ0.FLUTPJSTA
fgDetail.CellForeColor = wColor
fgDetail.Col = 7: fgDetail.Text = xYFLUTPJ0.FLUTPJID
fgDetail.CellForeColor = wColor
fgDetail.Col = 8: fgDetail.Text = Format$(xYFLUTPJ0.FLUTPJDOSQ, "### ### ##0")
fgDetail.CellForeColor = wColor

fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex
End Sub


Public Sub fgDetail_DisplayLine_3(lIndex As Long)
Dim K As Integer, blnOk As Boolean
Dim wColor As Long, wColor_Row As Long
Dim blnSolde As Boolean

On Error Resume Next

If xYFLUTP20.FLUTP2STA = " " Then
    blnOk = True
    wColor = vbBlue
Else
    blnOk = False
    wColor = RGB(128, 128, 128)
End If
If xYFLUTP20.FLUTP2TXCB = 0 Then wColor = vbMagenta

fgDetail.Col = 0: fgDetail.Text = xYFLUTP20.FLUTP2OPE & " " & xYFLUTP20.FLUTP2NAT & " " & Format$(xYFLUTP20.FLUTP2DOS, "@@@ @@@")
fgDetail.CellForeColor = wColor
fgDetail.Col = 1: fgDetail.Text = Format$(xYFLUTP20.FLUTP2MTD, "### ### ### ###.00")
fgDetail.CellForeColor = wColor
fgDetail.Col = 2: fgDetail.Text = dateImp10(xYFLUTP20.FLUTP2NEG) & "  "
fgDetail.CellForeColor = wColor
fgDetail.Col = 3: fgDetail.Text = dateImp10(xYFLUTP20.FLUTP2MAD) & "  "
fgDetail.CellForeColor = wColor
fgDetail.Col = 4: fgDetail.Text = dateImp10(xYFLUTP20.FLUTP2ECH) & "  "
fgDetail.CellForeColor = wColor


fgDetail.Col = 5: fgDetail.Text = Format$(xYFLUTP20.FLUTP2NBJ, "### ### ### ##0")
fgDetail.CellForeColor = wColor
fgDetail.Col = 6: fgDetail.Text = Format$(xYFLUTP20.FLUTP2ECHK, "### ### ### ##0")
fgDetail.CellForeColor = wColor

fgDetail.Col = 7: fgDetail.Text = Format$(xYFLUTP20.FLUTP2TX, "### ##0.00000")
fgDetail.CellForeColor = wColor
If xYFLUTP20.FLUTP2TXCB > 0 Then fgDetail.Col = 8: fgDetail.Text = Format$(xYFLUTP20.FLUTP2TXCB, "### ##0.00000")
fgDetail.CellForeColor = wColor
fgDetail.Col = 9: fgDetail.Text = xYFLUTP20.FLUTP2TXK
fgDetail.CellForeColor = wColor


fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex
End Sub



Public Sub fgFlux_Detail_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim blnOk As Boolean

If xFlux.FLUTPJSTA = " " Then
    blnOk = True
    wColor = vbBlue
Else
    blnOk = False
    wColor = RGB(128, 128, 128)
End If

On Error Resume Next
fgFlux_Detail.Col = 0: fgFlux_Detail.Text = Format$(xFlux.FLUTPJDOSQ, "### ### ##0")
fgFlux_Detail.CellForeColor = wColor
fgFlux_Detail.Col = 1: fgFlux_Detail.Text = dateImp10(xFlux.FLUTPJECH) & "  "
fgFlux_Detail.CellForeColor = wColor
fgFlux_Detail.Col = 2: fgFlux_Detail.Text = Format$(xFlux.FLUTPJMTD, "### ### ### ###.00")
If blnOk Then
    fgFlux_Detail.CellForeColor = IIf(xFlux.FLUTPJMTD < 0, vbRed, vbBlue)
Else
    fgFlux_Detail.CellForeColor = wColor
End If

fgFlux_Detail.Col = 3: fgFlux_Detail.Text = xFlux.FLUTPJDEV
fgFlux_Detail.CellForeColor = wColor
fgFlux_Detail.Col = 4: fgFlux_Detail.Text = xFlux.FLUTPJORIG & " " & xFlux.FLUTPJEVE & " " & xFlux.FLUTPJSTA
fgFlux_Detail.CellForeColor = wColor

fgFlux_Detail.Col = 5: fgFlux_Detail.Text = xFlux.FLUTPJID
fgFlux_Detail.CellForeColor = wColor

fgFlux_Detail.Col = fgFlux_Detail_arrIndex: fgFlux_Detail.Text = lIndex
End Sub


Public Sub fgFlux_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim blnOk As Boolean

On Error Resume Next
If xFlux.FLUTPJSTA = " " Then
    blnOk = True
    wColor = vbBlue
Else
    blnOk = False
    wColor = RGB(128, 128, 128)
End If

fgFlux.Col = 0: fgFlux.Text = xFlux.FLUTPJCCB
    fgFlux.CellForeColor = wColor
fgFlux.Col = 1: fgFlux.Text = xFlux.FLUTPJOPE
For K = 1 To arrOD_Nb
    If Trim(xFlux.FLUTPJOPE) = arrOD_K(K) Then
        fgFlux.Text = xFlux.FLUTPJOPE & " - " & arrOD_Lib(K)
        Exit For
    End If
Next K
    fgFlux.CellForeColor = wColor

fgFlux.Col = 2: fgFlux.Text = Format$(xFlux.FLUTPJDOS, "@@@ @@@ @@@")
    fgFlux.CellForeColor = wColor
fgFlux.Col = 3: fgFlux.Text = dateImp10(xFlux.FLUTPJECH) & "  " & xFlux.FLUTPJEVE
    fgFlux.CellForeColor = wColor
fgFlux.Col = 4: fgFlux.Text = Format$(xFlux.FLUTPJMTD, "### ### ### ###.00")
If blnOk Then
    fgFlux.CellForeColor = IIf(xFlux.FLUTPJMTD < 0, vbRed, vbBlue)
Else
    fgFlux.CellForeColor = wColor
End If

fgFlux.Col = 5: fgFlux.Text = xFlux.FLUTPJDEV
fgFlux.CellForeColor = IIf(xFlux.FLUTPJDEV = "EUR", wColor, vbMagenta)

fgFlux_Display_FLUTPJTXT
fgFlux.Col = 6: fgFlux.Text = Trim(xYFLUTPJ1.FLUTPJCLI) & " - " & xYFLUTPJ1.FLUTPJTXT
fgFlux.CellForeColor = vbMagenta

fgFlux.Col = fgFlux_arrIndex: fgFlux.Text = lIndex
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


Public Sub fgFlux_Detail_Sort()
If fgFlux_Detail.Rows > 1 Then
    fgFlux_Detail.Row = 1
    fgFlux_Detail.RowSel = fgFlux_Detail.Rows - 1
    
    If fgFlux_Detail_Sort1_Old = fgFlux_Detail_Sort1 Then
        If fgFlux_Detail_SortAD = 5 Then
            fgFlux_Detail_SortAD = 6
        Else
            fgFlux_Detail_SortAD = 5
        End If
    Else
        fgFlux_Detail_SortAD = 5
    End If
    fgFlux_Detail_Sort1_Old = fgFlux_Detail_Sort1
    
    fgFlux_Detail.Col = fgFlux_Detail_Sort1
    fgFlux_Detail.ColSel = fgFlux_Detail_Sort2
    fgFlux_Detail.Sort = fgFlux_Detail_SortAD
End If

End Sub


Public Sub fgFlux_Sort()
If fgFlux.Rows > 1 Then
    fgFlux.Row = 1
    fgFlux.RowSel = fgFlux.Rows - 1
    
    If fgFlux_Sort1_Old = fgFlux_Sort1 Then
        If fgFlux_SortAD = 5 Then
            fgFlux_SortAD = 6
        Else
            fgFlux_SortAD = 5
        End If
    Else
        fgFlux_SortAD = 5
    End If
    fgFlux_Sort1_Old = fgFlux_Sort1
    
    fgFlux.Col = fgFlux_Sort1
    fgFlux.ColSel = fgFlux_Sort2
    fgFlux.Sort = fgFlux_SortAD
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



Public Sub fgDetail_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgDetail.Rows - 1
    fgDetail.Row = I
    fgDetail.Col = fgDetail_arrIndex
    wIndex = Val(fgDetail.Text)
    Select Case lK
        Case 3: X = arrYFLUTPJ0(wIndex).FLUTPJECH
        Case 4: X = Format$(Abs(arrYFLUTPJ0(wIndex).FLUTPJMTD), "00000000000000.00")
        Case 5: X = arrYFLUTPJ0(wIndex).FLUTPJDEV & Format$(Abs(arrYFLUTPJ0(wIndex).FLUTPJMTD), "00000000000000.00")
        Case 6: X = arrYFLUTPJ0(wIndex).FLUTPJORIG & arrYFLUTPJ0(wIndex).FLUTPJEVE & Format$(arrYFLUTPJ0(wIndex).FLUTPJID, "00000000000")
        Case 7: X = Format$(arrYFLUTPJ0(wIndex).FLUTPJID, "00000000000")
        
   End Select
    fgDetail.Col = fgDetail_arrIndex - 1
    fgDetail.Text = X
Next I

fgDetail_Sort1 = fgDetail_arrIndex - 1: fgDetail_Sort2 = fgDetail_arrIndex - 1
fgdetail_Sort
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
Call BiaPgmAut_Init(wFct, YFLUTPJ0_Aut)

'blnSetfocus = True
Form_Init


Select Case wFct
    Case "@FLUX_TREPRE": blnAuto = True
                         auto_YFLUTPJ0
                         'cmdSelect_SQL_K = 1
                         'cmdSelect_Ok_Click
                         Call YFLUTPJ0_CB_Export(False, auto_YFLUTPJ0_mFile)
                         blnControl = False
                         cmdSelect_SQL_K = 2
                         Call DTPicker_Set(txtSelect_FLUTPJECH_Max, arrAMJMAX(2)) '
                         Call DTPicker_Set(txtSelect_FLUTPJECH_Min, arrAMJMIN(2)) '

                         cmdSelect_Ok_Click
                         cmdSendMail_Ech
                         Unload Me
    Case Else: blnAuto = False
End Select

End Sub


Public Sub Form_Init()
Dim V, xSQL As String, X As String
Dim K As Long, blnOk As Boolean

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True


cmdReset
blnControl = False

fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False



lstW.Visible = False
lstW.Clear



fraDetail.Visible = False
fraDetail.ForeColor = vbMagenta
libFlux_FLUTPJTXT.ForeColor = vbMagenta
fgDetail.Visible = True
fgDetail_FormatString = fgDetail.FormatString
mDate_A2 = dateElp("AnAdd", 2, YBIATAB0_DATE_CPT_JS1)
mDate_J7 = dateElp("Jour", 7, YBIATAB0_DATE_CPT_J)
Call DTPicker_Set(txtSelect_FLUTPJECH_Max, mDate_J7) '
Call DTPicker_Set(txtSelect_FLUTPJECH_Min, YBIATAB0_DATE_CPT_JS1) '

arrAMJMIN(1) = YBIATAB0_DATE_CPT_JS1
arrAMJMAX(1) = mDate_J7
arrAMJMIN(2) = YBIATAB0_DATE_CPT_JS1
arrAMJMAX(2) = mDate_J7
arrAMJMAX(3) = dateFinDeMois(YBIATAB0_DATE_CPT_J)
arrAMJMIN(3) = dateElp("MoisAdd", -2, mId$(YBIATAB0_DATE_CPT_J, 1, 6) & "01")

fraFlux_Detail.Visible = False
fraFlux_Detail.ForeColor = vbBlue
fgFlux_Detail_FormatString = fgFlux_Detail.FormatString

fgFlux.Visible = False
fgFlux_FormatString = fgFlux.FormatString

cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1  - CB : tableau de trésorerie prévisionnel"
cboSelect_SQL.AddItem "2  - Echéancier des flux"
cboSelect_SQL.AddItem "3  - INFO-LIQU coût de refinancement"
If YFLUTPJ0_Aut.Xspécial Then cboSelect_SQL.AddItem "Pr Parametrage_Reprise"
cboSelect_SQL.ListIndex = 0: cmdSelect_SQL_K = "1"



'Echéancier : filtre _________________________________________
fraSelect_Option_2.Visible = False
Set fraSelect_Option_2.Container = fraSelect_Options
fraSelect_Option_2.Left = 120
fraSelect_Option_2.Top = 200
cboSelect_FLUTPJORIG.Clear
cboSelect_FLUTPJORIG.AddItem "  - tous les flux"
cboSelect_FLUTPJORIG.AddItem "* - Flux internes"
cboSelect_FLUTPJORIG.AddItem "1 - TRE"
cboSelect_FLUTPJORIG.AddItem "1*- TRE hors *JJ & *JD"
cboSelect_FLUTPJORIG.AddItem "2 - CDO"
cboSelect_FLUTPJORIG.AddItem "3 - DAT"
cboSelect_FLUTPJORIG.AddItem "4 - CHG"
cboSelect_FLUTPJORIG.AddItem "5 - CRE"
cboSelect_FLUTPJORIG.AddItem "6 - REM"
cboSelect_FLUTPJORIG.AddItem "7 - NOS"


cboSelect_FLUTPJOPE.Clear
cboSelect_FLUTPJOPE.AddItem ""
xSQL = "select distinct FLUTPJOPE from " & paramIBM_Library_SABSPE & ".YFLUTPJ0 order by FLUTPJOPE"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    cboSelect_FLUTPJOPE.AddItem rsSab("FLUTPJOPE")
    rsSab.MoveNext
Loop


'parametrage ________________________________________________________________________________
lstParam_K.Visible = False
fraParam_K.Visible = False
fraParam_K.ForeColor = vbMagenta

lstParam_Action.AddItem "1- codes CB => libellé"
lstParam_Action.AddItem "2- codes flux échéancés => codes CB"
lstParam_Action.AddItem "3- codes flux complémentaires =>  codes CB"
lstParam_Action.AddItem "4- codes flux NOSTRO => codes CB"
lstParam_Action.AddItem "5- flux complémentaires / racine client"
lstParam_Action.AddItem "6- habilitations utilisateurs / codes internes"

cboParam_Frequence.AddItem "U - unitaire"
cboParam_Frequence.AddItem "J - journalière"
cboParam_Frequence.AddItem "M - mensuelle"
cboParam_Frequence.AddItem "T - trimestrielle"
cboParam_Frequence.AddItem "S - semestrielle"
cboParam_Frequence.AddItem "A - annuelle"

parametrage_FLUTPJCCB_Load

'Initialisation devise________________________________________________________________________________
arrDev_Nb = 0
ReDim Preserve arrDev(1000)

arrDev(1) = "EUR"
arrDev(2) = "USD"
arrDev(3) = "GBP"
arrDev(4) = "CAD"
arrDev(5) = "JPY"
arrDev(6) = "CHF"
arrDev(7) = "DKK"
arrDev(8) = "SEK"
arrDev(9) = "AED"
arrDev_Nb = 9

cboFlux_FLUTPJDEV.Clear
cboFlux_FLUTPJDEV.AddItem ""
xSQL = "select distinct FLUTPJDEV from " & paramIBM_Library_SABSPE & ".YFLUTPJ0 order by FLUTPJDEV"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    X = Trim(rsSab("FLUTPJDEV"))
    Select Case X
        Case "EUR", "USD", "GBP", "CAD", "JPY", "CHF", "DKK", "SEK", "AED"
        Case Else
            arrDev_Nb = arrDev_Nb + 1
            arrDev(arrDev_Nb) = X
            fgSelect_FormatString = fgSelect_FormatString & ">" & X & "          |"
    End Select
    rsSab.MoveNext
Loop
ReDim Preserve arrDev(arrDev_Nb + 1)
ReDim arrDev_DB(arrDev_Nb + 1), arrDev_CR(arrDev_Nb + 1), arrDev_Cours(arrDev_Nb + 1)
For K = 1 To arrDev_Nb
    cboFlux_FLUTPJDEV.AddItem arrDev(K)
    arrDev_Cours(K) = 0
Next K

libSelect_FLUTPJDEV.ForeColor = vbMagenta
'____________________________________________________________________________________________

fraFlux_Update.Visible = False
fraFlux_Update.ForeColor = vbMagenta
Call DTPicker_Set(txtFlux_FLUTPJECH_Min, YBIATAB0_DATE_CPT_JS1) '
fraFlux_Update_A.BorderStyle = 0
fraFlux_Update_B.BorderStyle = 0
fraFlux_UPDATE_C.BorderStyle = 0

fgFlux.Visible = False
fgFlux_FormatString = fgFlux.FormatString

parametrage_FLUTPJOD_Load

cboFlux_Frequence.AddItem "U - unitaire"
cboFlux_Frequence.AddItem "J - journalière"
cboFlux_Frequence.AddItem "M - mensuelle"
cboFlux_Frequence.AddItem "T - trimestrielle"
cboFlux_Frequence.AddItem "S - semestrielle"
cboFlux_Frequence.AddItem "A - annuelle"

'Initialisation libellé OPERATION________________________________________________________________________________
ReDim arrOPE_K(1000), arrOPE_Lib(1000)

arrOPE_Nb = 7
arrOPE_K(1) = "TRE   ": arrOPE_Lib(1) = "Trésorerie"
arrOPE_K(2) = "CDO   ": arrOPE_Lib(2) = "Crédit documentaire"
arrOPE_K(3) = "DAT   ": arrOPE_Lib(3) = "Dépôt à terme"
arrOPE_K(4) = "CHG   ": arrOPE_Lib(4) = "Change à terme,swap"
arrOPE_K(5) = "CRE   ": arrOPE_Lib(5) = "Crédit"
arrOPE_K(6) = "RDO   ": arrOPE_Lib(6) = "Remise documentaire"
arrOPE_K(7) = "KS    ": arrOPE_Lib(7) = "opérations de caisse"


xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 where BASTABNUM = 23 order by BASTABARG"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    arrOPE_Nb = arrOPE_Nb + 1
    X = Trim(rsSab("BASTABDON"))
    arrOPE_K(arrOPE_Nb) = Trim(rsSab("BASTABARG"))
    arrOPE_Lib(arrOPE_Nb) = Trim(mId$(rsSab("BASTABDON"), 1, 30))
    
    rsSab.MoveNext
Loop

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCRETAB0 where CRETABNUM = 9 order by CRETABARG"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    arrOPE_Nb = arrOPE_Nb + 1
    X = Trim(rsSab("CRETABDON"))
    arrOPE_K(arrOPE_Nb) = "CRE" & Trim(rsSab("CRETABARG"))
    arrOPE_Lib(arrOPE_Nb) = Trim(mId$(rsSab("CRETABDON"), 1, 30))
    
    rsSab.MoveNext
Loop

arrOPE_K(arrOPE_Nb + 1) = "      ": arrOPE_Lib(arrOPE_Nb + 1) = "???"

ReDim Preserve arrOPE_K(arrOPE_Nb + 1), arrOPE_Lib(arrOPE_Nb + 1)

'_________________________________________________________________________________________
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

Public Sub fgFlux_Detail_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgFlux_Detail.Visible = False
mRow = fgFlux_Detail.Row

If lRow > 0 And lRow < fgFlux_Detail.Rows Then
    fgFlux_Detail.Row = lRow
    For I = fgFlux_Detail_arrIndex To fgFlux_Detail.FixedCols Step -1
        fgFlux_Detail.Col = I: fgFlux_Detail.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgFlux_Detail.Row = mRow
    If fgFlux_Detail.Row > 0 Then
        lRow = fgFlux_Detail.Row
        lColor_Old = fgFlux_Detail.CellBackColor
        For I = fgFlux_Detail_arrIndex To fgFlux_Detail.FixedCols Step -1
          fgFlux_Detail.Col = I: fgFlux_Detail.CellBackColor = lColor
        Next I
    End If
End If
fgFlux_Detail.LeftCol = fgFlux_Detail.FixedCols
fgFlux_Detail.Visible = True
End Sub

Public Sub fgFlux_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgFlux.Visible = False
mRow = fgFlux.Row

If lRow > 0 And lRow < fgFlux.Rows Then
    fgFlux.Row = lRow
    For I = fgFlux_arrIndex To fgFlux.FixedCols Step -1
        fgFlux.Col = I: fgFlux.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgFlux.Row = mRow
    If fgFlux.Row > 0 Then
        lRow = fgFlux.Row
        lColor_Old = fgFlux.CellBackColor
        For I = fgFlux_arrIndex To fgFlux.FixedCols Step -1
          fgFlux.Col = I: fgFlux.CellBackColor = lColor
        Next I
    End If
End If
fgFlux.LeftCol = fgFlux.FixedCols
fgFlux.Visible = True
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
Private Function Parametrage_Delete()
Dim xSQL As String
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

Private Function Parametrage_New()
Dim xSQL As String
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

Public Sub Parametrage_Reprise()

Dim K As Integer, X As String

Call MsgBox("Reprise paramétrage INTERDIT", vbCritical, "FLUTPJ")
Exit Sub


New_YBIATAB0.BIATABID = "FLUTPJCCB"
New_YBIATAB0.BIATABK1 = "100"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "DECAISSEMENTS PREVISIONNELS"
Call Parametrage_New
New_YBIATAB0.BIATABK1 = "200"
New_YBIATAB0.BIATABTXT = "ENCAISSEMENTS PREVISIONNELS"
Call Parametrage_New

New_YBIATAB0.BIATABK1 = "101"
New_YBIATAB0.BIATABTXT = "Opérations avec la banque centrale (Eurosystème)"
Call Parametrage_New
New_YBIATAB0.BIATABK1 = "201"
Call Parametrage_New

New_YBIATAB0.BIATABK1 = "102"
New_YBIATAB0.BIATABTXT = "Prêts / emprunts interbancaires (dont intragroupe)"
Call Parametrage_New
New_YBIATAB0.BIATABK1 = "202"
Call Parametrage_New

New_YBIATAB0.BIATABK1 = "103"
New_YBIATAB0.BIATABTXT = "Achat / Prise en pension de titres financiers"
Call Parametrage_New
New_YBIATAB0.BIATABK1 = "203"
New_YBIATAB0.BIATABTXT = "vente / Mise en pension de titres financiers"
Call Parametrage_New

New_YBIATAB0.BIATABK1 = "104"
New_YBIATAB0.BIATABTXT = "Retraits/ dépôts à vue de la clientèle (nets)"
Call Parametrage_New
New_YBIATAB0.BIATABK1 = "204"
New_YBIATAB0.BIATABTXT = "Titres financiers émis"
Call Parametrage_New

New_YBIATAB0.BIATABK1 = "105"
New_YBIATAB0.BIATABTXT = "Retraits/ dépôts à terme de la clientèle (nets)"
Call Parametrage_New
New_YBIATAB0.BIATABK1 = "205"
New_YBIATAB0.BIATABTXT = "Remboursement clientèle"
Call Parametrage_New

New_YBIATAB0.BIATABK1 = "106"
New_YBIATAB0.BIATABTXT = "Titres financiers émis"
Call Parametrage_New
New_YBIATAB0.BIATABK1 = "206"
New_YBIATAB0.BIATABTXT = "Instruments financiers à terme"
Call Parametrage_New

New_YBIATAB0.BIATABK1 = "107"
New_YBIATAB0.BIATABTXT = "Prêts clientèle et engagements mis en force"
Call Parametrage_New
New_YBIATAB0.BIATABK1 = "207"
New_YBIATAB0.BIATABTXT = "Titrisations"
Call Parametrage_New

New_YBIATAB0.BIATABK1 = "108"
New_YBIATAB0.BIATABTXT = "Instruments financiers à terme"
Call Parametrage_New
New_YBIATAB0.BIATABK1 = "208"
New_YBIATAB0.BIATABTXT = "Engagements de financement reçus"
Call Parametrage_New

New_YBIATAB0.BIATABK1 = "109"
New_YBIATAB0.BIATABTXT = "Titrisations"
Call Parametrage_New
New_YBIATAB0.BIATABK1 = "209"
New_YBIATAB0.BIATABTXT = "opérations de change (swaps de devises)"
Call Parametrage_New

New_YBIATAB0.BIATABK1 = "110"
New_YBIATAB0.BIATABTXT = "Autres opérations de marché dont opérations de change (swaps de devises)"
Call Parametrage_New
New_YBIATAB0.BIATABK1 = "210"
New_YBIATAB0.BIATABTXT = "Autres encaissements (à préciser)"
Call Parametrage_New

New_YBIATAB0.BIATABK1 = "111"
New_YBIATAB0.BIATABTXT = "Autres décaissements (à préciser)"
Call Parametrage_New

New_YBIATAB0.BIATABK1 = "300"
New_YBIATAB0.BIATABTXT = "SOLDE NET PREVISIONNEL"
Call Parametrage_New

'_______________________________________________________________________

New_YBIATAB0.BIATABID = "FLUTPJOPE"
New_YBIATAB0.BIATABK1 = "TRE"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "102 202"
Call Parametrage_New
New_YBIATAB0.BIATABID = "FLUTPJOPE"
New_YBIATAB0.BIATABK1 = "CDO"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "107 000" '"107 208"
Call Parametrage_New
New_YBIATAB0.BIATABID = "FLUTPJOPE"
New_YBIATAB0.BIATABK1 = "DAT"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "105 210"
Call Parametrage_New
New_YBIATAB0.BIATABID = "FLUTPJOPE"
New_YBIATAB0.BIATABK1 = "CHG"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "110 209"
Call Parametrage_New
New_YBIATAB0.BIATABID = "FLUTPJOPE"
New_YBIATAB0.BIATABK1 = "CRE"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "107 205"
Call Parametrage_New
New_YBIATAB0.BIATABID = "FLUTPJOPE"
New_YBIATAB0.BIATABK1 = "RDO"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "107 205"
Call Parametrage_New
'_______________________________________________________________________
New_YBIATAB0.BIATABID = "FLUTPJOPE"
New_YBIATAB0.BIATABK1 = "CRE"
New_YBIATAB0.BIATABK2 = "DBI": New_YBIATAB0.BIATABTXT = "102 202": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "CBD": New_YBIATAB0.BIATABTXT = "102 202": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "CAB": New_YBIATAB0.BIATABTXT = "102 202": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "CFI": New_YBIATAB0.BIATABTXT = "102 202": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "CGB": New_YBIATAB0.BIATABTXT = "102 202": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "SPT": New_YBIATAB0.BIATABTXT = "102 202": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "CDB": New_YBIATAB0.BIATABTXT = "102 202": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "ECB": New_YBIATAB0.BIATABTXT = "102 202": Call Parametrage_New

New_YBIATAB0.BIATABK2 = "DTI": New_YBIATAB0.BIATABTXT = "106 204": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "CDT": New_YBIATAB0.BIATABTXT = "106 204": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "PAR": New_YBIATAB0.BIATABTXT = "106 204": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "TDC": New_YBIATAB0.BIATABTXT = "106 204": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "TIN": New_YBIATAB0.BIATABTXT = "106 204": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "TBB": New_YBIATAB0.BIATABTXT = "106 204": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "TID": New_YBIATAB0.BIATABTXT = "106 204": Call Parametrage_New

'_______________________________________________________________________

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "RET"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "104 000 J Retraits espèces"
Call Parametrage_New

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "SAL"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "111 000 M Salaires et appointements"
Call Parametrage_New

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "FGX"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "111 000 U Autres frais généraux/immos"
Call Parametrage_New

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "IS"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "111 000 T Impôts sur les sociétés"
Call Parametrage_New

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "CET"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "111 000 S Contribution économique territoriale"
Call Parametrage_New

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "TVA"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "111 000 M TVA"
Call Parametrage_New

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "IMP"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "111 000 A Autres impôts"
Call Parametrage_New

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "TRF"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "111 000 J Transferts"
Call Parametrage_New

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "DIV"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "111 000 A Dividendes versés BIA"
Call Parametrage_New

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "PAR"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "111 000 A Intéressement et participations"
Call Parametrage_New

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "RES"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "111 000 A Prime de résultat"
Call Parametrage_New

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "COM"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "111 210 U Commissions"
Call Parametrage_New


New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "DIR"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "000 210 A Dividendes reçus"
Call Parametrage_New

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "RAP"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "000 210 U Rapatriement"
Call Parametrage_New

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "VER"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "000 210 U Versements espèces"
Call Parametrage_New

New_YBIATAB0.BIATABID = "FLUTPJOD"
New_YBIATAB0.BIATABK1 = "CRI"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "000 210 U Crédit d'impôt"
Call Parametrage_New
'_______________________________________________________________________
New_YBIATAB0.BIATABID = "FLUTPJCPT"
New_YBIATAB0.BIATABK1 = "AV0 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "104 104": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "REM 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "104 104": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "RI0 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "104 104": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "RP0 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "104 104": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "RV0 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "104 104": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "CPT 00 TR": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "104 104": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "TRF 00 TR": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "104 210": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "KS  00 GU": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "104 104": Call Parametrage_New

New_YBIATAB0.BIATABK1 = "*B1 CP CP": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*B1 TC TC": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*B1 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*C5 TC TC": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*C6 00 CR": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*C6 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*G1 00 GU": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*G1 00 MP": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*G1 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*L7 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*L8 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*T1 TC TC": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*T1 00 TR": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*T1 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*T4 TC TC": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*T5 00 CR": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*T5 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*Z1 CP CP": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*Z1 CP JC": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*Z1 RH RH": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "*Z1 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "-SL RH JH": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "-TR CP CP": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "-TR CP JC": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "CDE 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "CDI 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "CPT TC TC": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "CRE 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "EMP TC TC": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "ENG 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "FRS 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "PRE TC TC": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "RA0 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "RDE 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "RDI 00 00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "SWP TC TC": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "TRF TC TC": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "TRF 00 CR": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "000 000": Call Parametrage_New


New_YBIATAB0.BIATABK2 = "="
New_YBIATAB0.BIATABK1 = "AI1 00 00": New_YBIATAB0.BIATABTXT = "AV0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "AL0 00 00": New_YBIATAB0.BIATABTXT = "AV0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "AL1 00 00": New_YBIATAB0.BIATABTXT = "AV0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "AP1 00 00": New_YBIATAB0.BIATABTXT = "AV0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "AV0 00 TR": New_YBIATAB0.BIATABTXT = "AV0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "AV1 00 00": New_YBIATAB0.BIATABTXT = "AV0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "A0V 00 TR": New_YBIATAB0.BIATABTXT = "AV0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "A0V 00 00": New_YBIATAB0.BIATABTXT = "AV0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "REM 00 GU": New_YBIATAB0.BIATABTXT = "REM 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "RI0 00 TR": New_YBIATAB0.BIATABTXT = "RI0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "RL0 00 TR": New_YBIATAB0.BIATABTXT = "RI0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "RP0 00 TR": New_YBIATAB0.BIATABTXT = "RP0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "RT0 00 00": New_YBIATAB0.BIATABTXT = "RP0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "RV0 00 TR": New_YBIATAB0.BIATABTXT = "RV0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "RV1 00 00": New_YBIATAB0.BIATABTXT = "RV0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "R0V 00 00": New_YBIATAB0.BIATABTXT = "RV0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "R2V 00 00": New_YBIATAB0.BIATABTXT = "RV0 00 00": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "TRF 00 00": New_YBIATAB0.BIATABTXT = "TRF 00 TR": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "TRF 00 MP": New_YBIATAB0.BIATABTXT = "TRF 00 TR": Call Parametrage_New

New_YBIATAB0.BIATABK1 = "RE  00 GU": New_YBIATAB0.BIATABTXT = "KS  00 GU": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "VE  00 GU": New_YBIATAB0.BIATABTXT = "KS  00 GU": Call Parametrage_New
New_YBIATAB0.BIATABK1 = "CPT 00 GU": New_YBIATAB0.BIATABTXT = "KS  00 GU": Call Parametrage_New
'_______________________________________________________________________
New_YBIATAB0.BIATABID = "FLUTPJCLI"
New_YBIATAB0.BIATABK1 = "DIR"
New_YBIATAB0.BIATABK2 = "0011015": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0050651": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0011018": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0050487": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0050697": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0050695": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0050694": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0050696": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0050698": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0050655": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0011105": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0050531": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0011040": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0011454": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0011457": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0011459": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New
New_YBIATAB0.BIATABK2 = "0011466": New_YBIATAB0.BIATABTXT = "": Call Parametrage_New


'_______________________________________________________________________

End Sub

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








Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

End Sub





Private Sub cboFlux_FLUTPJDEV_Change()
cmdFlux_Reset

End Sub

Private Sub cboFlux_FLUTPJDEV_Click()
cmdFlux_Reset

End Sub

Private Sub cboFlux_FLUTPJOD_Click()
cmdDetail_Reset
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

cmdPrint_Option = ""
mnuCB_Export.Enabled = False
mnuECH_Export.Enabled = False
mnuECH_Mail.Enabled = False
mnuParam_Export.Enabled = False
Select Case SSTab1.Tab
    Case 0:
        Select Case cmdSelect_SQL_K
            Case "1", "3": mnuCB_Export.Enabled = True
            Case "2": mnuECH_Export.Enabled = True: mnuECH_Mail.Enabled = True
        End Select

        Me.PopupMenu mnuPrint, vbPopupMenuLeftButton
    Case 2:
        mnuParam_Export.Enabled = True
        Me.PopupMenu mnuPrint, vbPopupMenuLeftButton
End Select

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdPrint_YFLUTPJ0(blnDetail As Boolean)
Dim X As String, xSQL As String, I As Integer, K As Integer
Dim wAmj As String, xWhere As String
Dim soldeD As typeYFLUTPJ0, soldeF As typeYFLUTPJ0, total As typeYFLUTPJ0
Dim blnXprt_Line As Boolean



Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> FLUX_TREPRE_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Reset
fgSelect.Visible = False
'fraSelect_Options.Visible = False

Select Case cmdSelect_SQL_K
    Case "1": fraSelect_Options.Visible = True: cmdSelect_SQL_1
    Case "2": fraSelect_Options.Visible = True: cmdSelect_SQL_Ech
    Case "3": fraSelect_Options.Visible = True: cmdSelect_SQL_3
    Case "Pr": Parametrage_Reprise
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< FLUX_TREPRE_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
End Sub


Private Sub fgFLUTPJSTAT_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next

If y <= fgFLUTPJSTAT.RowHeightMin Then
Else
    If fgFLUTPJSTAT.Rows > 1 Then
        'Call fgFLUTPJSTAT_Color(fgFLUTPJSTAT_RowClick, MouseMoveUsr.BackColor, fgFLUTPJSTAT_ColorClick)
        'fgFLUTPJSTAT.Col = fgFLUTPJSTAT_arrIndex:  arrfgFLUTPJSTAT_Index = CLng(fgFLUTPJSTAT.Text)
        'oldFlux = arrfgFLUTPJSTAT(arrfgFLUTPJSTAT_Index)
        Dim xSQL As String, xFLUCPTOPE As String, xFLUCPTDEV As String, xFLUCPTMTD, xFLUCPTMTD_Sens As String
        lstW.Clear
        lstW.AddItem "Liste des opérations exclues de la moyenne journalière"
        lstW.AddItem "------------------------------------------------------"
        fgFLUTPJSTAT.Col = 0: xFLUCPTOPE = mId$(fgFLUTPJSTAT.Text, 1, 3)
        fgFLUTPJSTAT.Col = 1: xFLUCPTDEV = Trim(fgFLUTPJSTAT.Text)
        fgFLUTPJSTAT.Col = 5: xFLUCPTMTD = Val(fgFLUTPJSTAT.Text)
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YFLUCPT0 " _
             & " where FLUCPTDEV  = '" & xFLUCPTDEV & "' " _
             & " and FLUCPTOPE = '" & xFLUCPTOPE & "' and FLUCPTMTD > " & xFLUCPTMTD & " order by FLUCPTECH"
        If xFLUCPTMTD < 0 Then xSQL = Replace(xSQL, "FLUCPTMTD >", "FLUCPTMTD <")
        
        Set rsSab = cnsab.Execute(xSQL)
        
        Do While Not rsSab.EOF
            lstW.AddItem dateImp_Amj(rsSab("FLUCPTECH")) & " " & Format$(rsSab("FLUCPTMTD"), "### ### ### ##0.00")
            rsSab.MoveNext
        Loop

        lstW.Visible = True
   End If
End If
End Sub


Private Sub fgFlux_Detail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next

If y <= fgFlux_Detail.RowHeightMin Then
Else
    If fgFlux_Detail.Rows > 1 Then
        Call fgFlux_Detail_Color(fgFlux_Detail_RowClick, MouseMoveUsr.BackColor, fgFlux_Detail_ColorClick)
        fgFlux_Detail.Col = fgFlux_Detail_arrIndex:  arrfgFlux_Detail_Index = CLng(fgFlux_Detail.Text)
        oldFlux = arrfgFlux_Detail(arrfgFlux_Detail_Index)
        xFlux = oldFlux
        fraFlux_Display
        
        cboFlux_FLUTPJCLI.Locked = True

        mnuFlux_Option = ""
        fraFlux_Update.Caption = ""
        cmdFlux_Update_Ok.Visible = False
        fraFlux_Update_A.Enabled = False
        fraFlux_Update_B.Enabled = False
        fraFlux_UPDATE_C.Enabled = False
        If YFLUTPJ0_Aut.Valider Then Me.PopupMenu mnuFlux_Detail, , 11800, 8000

   End If
End If


End Sub


Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next


If y <= fgDetail.RowHeightMin Then
    Select Case fgDetail.Col
        Case 0: fgDetail_Sort1 = 0: fgDetail_Sort2 = 2: fgdetail_Sort
        Case 1:  fgDetail_Sort1 = 1: fgDetail_Sort2 = 2: fgdetail_Sort
        Case 2:  fgDetail_Sort1 = 2: fgDetail_Sort2 = 2: fgdetail_Sort
        Case 3:  fgDetail_Sort1 = 3: fgDetail_Sort2 = 3: fgDetail_SortX 3
        Case 4:  fgDetail_Sort1 = 4: fgDetail_Sort2 = 4: fgDetail_SortX 4
        Case 5:  fgDetail_Sort1 = 5: fgDetail_Sort2 = 5: fgDetail_SortX 5
        Case 6:  fgDetail_Sort1 = 6: fgDetail_Sort2 = 6: fgDetail_SortX 6
        Case 7:  fgDetail_Sort1 = 7: fgDetail_Sort2 = 7: fgDetail_SortX 7
        'Case fgDetail_arrIndex:  fgdetail_SortX fgDetail_arrIndex
    End Select
Else
    If fgDetail.Rows > 1 Then
        Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
        fgDetail.Col = fgDetail_arrIndex:  arrYFLUTPJ0_Index = CLng(fgDetail.Text)
        oldYFLUTPJ0 = arrYFLUTPJ0(arrYFLUTPJ0_Index)
        xYFLUTPJ0 = oldYFLUTPJ0
        'fgFlux_Detail_Display

   End If
End If

End Sub


Private Sub fgFlux_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Dim xWhere As String
If y <= fgFlux.RowHeightMin Then
Else
    If fgFlux.Rows > 1 Then
        Call fgFlux_Color(fgFlux_RowClick, MouseMoveUsr.BackColor, fgFlux_ColorClick)
        fgFlux.Col = fgFlux_arrIndex:  arrfgFlux_Index = CLng(fgFlux.Text)
        oldFlux = arrfgFlux(arrfgFlux_Index)
        xFlux = oldFlux
         xWhere = " Where FLUTPJORIG = '*' and FLUTPJDOS = " & oldFlux.FLUTPJDOS _
                & " and FLUTPJDOSQ >= " & oldFlux.FLUTPJDOSQ & " order by FLUTPJDOSQ" _
    
        arrfgFlux_Detail_SQL xWhere
        fgFlux.Col = 1
        fraFlux_Detail.Caption = "dossier : " & oldFlux.FLUTPJDOS & " - " & fgFlux.Text
        fgFlux_Detail_Display
        fgFlux_Display_FLUTPJTXT
        
        fraFlux_Display
        fgFlux_Display_FLUTPJTXT
        txtFLUX_FLUTPJTXT = Trim(xYFLUTPJ1.FLUTPJTXT)
        cboFlux_FLUTPJCLI.Locked = False
        fraFlux_Display_FLUTPJCLI
        If Trim(xYFLUTPJ1.FLUTPJCLI) <> "" Then
            libFlux_FLUTPJTXT = cboFlux_FLUTPJCLI.Text & vbCrLf & xYFLUTPJ1.FLUTPJTXT
        Else
            libFlux_FLUTPJTXT = xYFLUTPJ1.FLUTPJTXT
        End If
        
        oldYFLUTPJ1 = xYFLUTPJ1
        mnuFlux_Option = ""
        fraFlux_Update.Caption = ""
        cmdFlux_Update_Ok.Visible = False
        fraFlux_Update_A.Enabled = False
        fraFlux_Update_B.Enabled = False
        fraFlux_UPDATE_C.Enabled = False

        If YFLUTPJ0_Aut.Valider Then Me.PopupMenu mnuFlux, , 11800, 8000
   End If
End If

End Sub

Private Sub lstParam_Action_Click()

lstParam_Action_K = mId$(lstParam_Action, 1, 1)
lstParam_K.Visible = False
fgFLUTPJSTAT.Visible = False
lstW.Visible = False

lstParam_K.Clear
fraParam_K.Visible = False
cmdParam_Quit.Visible = True
mnuParam_Add.Visible = YFLUTPJ0_Aut.Comptabiliser
mnuParam_Update.Visible = YFLUTPJ0_Aut.Comptabiliser
mnuParam_Delete.Visible = YFLUTPJ0_Aut.Comptabiliser
Select Case lstParam_Action_K
    Case "1": Parametrage_FLUTPJCCB_Init
    Case "2": Parametrage_FLUTPJOPE_Init
    Case "3": Parametrage_FLUTPJOD_Init
    Case "4": Parametrage_FLUTPJCPT_Init
    Case "5": Parametrage_FLUTPJCLI_Init
End Select
End Sub

Private Sub xxx_lstParam_K_Click()
Dim xSQL As String, K As Integer
Old_YBIATAB0.BIATABID = lstParam_Action_K
Old_YBIATAB0.BIATABK1 = mId$(lstParam_K, 1, 3)
txtParam_K = Old_YBIATAB0.BIATABK1
Select Case lstParam_Action_K
    Case "FLUTPJOPE":
        Old_YBIATAB0.BIATABK2 = mId$(lstParam_K, 7, 3)
        txtParam_K2 = Old_YBIATAB0.BIATABK2
        txtParam_K2.Visible = True
    Case "FLUTPJCPT":
        K = InStr(lstParam_K, "=")
        Old_YBIATAB0.BIATABK2 = mId$(lstParam_K, 7, 3)
        txtParam_K2 = Old_YBIATAB0.BIATABK2
        txtParam_K2.Visible = True
    Case Else:
        Old_YBIATAB0.BIATABK2 = ""
        txtParam_K2 = Old_YBIATAB0.BIATABK2
        txtParam_K2.Visible = False
End Select

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = '" & Old_YBIATAB0.BIATABID & "' and BIATABK1 = '" & Trim(Old_YBIATAB0.BIATABK1) & "'  and BIATABK2 = '" & Trim(Old_YBIATAB0.BIATABK2) & "'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    Old_YBIATAB0.BIATABTXT = rsSab("BIATABTXT")
    txtParam_X = Trim(Old_YBIATAB0.BIATABTXT)
    Select Case lstParam_Action_K
        Case "FLUTPJOPE"
            txtParam_X = ""
            Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 1, 3), cboParam_CCB_DB)
            Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 5, 3), cboParam_CCB_CR)
        Case "FLUTPJOD"
            Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 1, 3), cboParam_CCB_DB)
            Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 5, 3), cboParam_CCB_CR)
            Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 9, 1), cboParam_Frequence)
            txtParam_X = Trim(mId$(Old_YBIATAB0.BIATABTXT, 11, 64))
    End Select
    
Else
    txtParam_X = ""
End If
mnuParam_Option = ""
cmdParam_Ok.Visible = False
fraParam_K2.Enabled = False
fraParam_K.Visible = True
fraParam_K.Caption = ""
If YFLUTPJ0_Aut.Comptabiliser Then Me.PopupMenu mnuParam, , 10000, 8300

End Sub

Private Sub lstParam_K_Click()
Dim xSQL As String, K As Integer, X As String

lstW.Visible = False
Old_YBIATAB0 = arrParam(lstParam_K.ListIndex + 1)
txtParam_K = Trim(Old_YBIATAB0.BIATABK1)
txtParam_K2 = Trim(Old_YBIATAB0.BIATABK2)
txtParam_X = Trim(Old_YBIATAB0.BIATABTXT)
fgFLUTPJSTAT.Visible = False
Select Case lstParam_Action_K
    Case "FLUTPJOPE"
        txtParam_K2.Visible = True
        txtParam_X = ""
        Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 1, 3), cboParam_CCB_DB)
        Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 5, 3), cboParam_CCB_CR)
    Case "FLUTPJOD"
        txtParam_K2.Visible = False
        Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 1, 3), cboParam_CCB_DB)
        Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 5, 3), cboParam_CCB_CR)
        Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 9, 1), cboParam_Frequence)
        txtParam_X = Trim(mId$(Old_YBIATAB0.BIATABTXT, 11, 64))
    Case "FLUTPJCPT"
        If Trim(Old_YBIATAB0.BIATABK2) = "=" Then
            txtParam_K2.Visible = True
            txtParam_K2 = Old_YBIATAB0.BIATABTXT
            cboParam_CCB_DB.Visible = False
            cboParam_CCB_CR.Visible = False
            txtParam_X = "Indiquer le code opé service sous-service de regroupement *** ** **"
        Else
            txtParam_K2.Visible = False
            cboParam_CCB_DB.Visible = True
            cboParam_CCB_CR.Visible = True
            txtParam_X = ""
            Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 1, 3), cboParam_CCB_DB)
            Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 5, 3), cboParam_CCB_CR)
            xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJSTAT'" _
                 & " and BIATABK1 = '" & Old_YBIATAB0.BIATABK1 & "'  order by BIATABK2"
            Set rsSab = cnsab.Execute(xSQL)
            fgFLUTPJSTAT.Rows = 1
            Do While Not rsSab.EOF
                fgFLUTPJSTAT.Rows = fgFLUTPJSTAT.Rows + 1
                fgFLUTPJSTAT.Row = fgFLUTPJSTAT.Rows - 1
                fgFLUTPJSTAT.Col = 0: fgFLUTPJSTAT.Text = Trim(rsSab("BIATABK1"))
                fgFLUTPJSTAT.Col = 1: fgFLUTPJSTAT.Text = Trim(rsSab("BIATABK2"))
                X = rsSab("BIATABTXT")
                fgFLUTPJSTAT.Col = 2: fgFLUTPJSTAT.Text = Format(Val(mId$(X, 1, 9)), "### ### ###")
                fgFLUTPJSTAT.CellForeColor = vbRed
                fgFLUTPJSTAT.Col = 3: fgFLUTPJSTAT.Text = Format(-Val(mId$(X, 11, 17)), "### ### ### ###")
                fgFLUTPJSTAT.CellForeColor = vbRed
                fgFLUTPJSTAT.Col = 4: fgFLUTPJSTAT.Text = Format(-Val(mId$(X, 29, 17)), "### ### ### ###")
                fgFLUTPJSTAT.CellForeColor = vbRed
                fgFLUTPJSTAT.Col = 5: fgFLUTPJSTAT.Text = Format(-Val(mId$(X, 47, 17)), "### ### ### ###")
                 fgFLUTPJSTAT.CellForeColor = vbRed
                fgFLUTPJSTAT.Rows = fgFLUTPJSTAT.Rows + 1
                fgFLUTPJSTAT.Row = fgFLUTPJSTAT.Rows - 1
                fgFLUTPJSTAT.Col = 0: fgFLUTPJSTAT.Text = Trim(rsSab("BIATABK1"))
                fgFLUTPJSTAT.Col = 1: fgFLUTPJSTAT.Text = Trim(rsSab("BIATABK2"))
                fgFLUTPJSTAT.Col = 2: fgFLUTPJSTAT.Text = Format(Val(mId$(X, 66, 9)), "### ### ###")
                fgFLUTPJSTAT.Col = 3: fgFLUTPJSTAT.Text = Format(Val(mId$(X, 76, 17)), "### ### ### ###")
                fgFLUTPJSTAT.Col = 4: fgFLUTPJSTAT.Text = Format(Val(mId$(X, 94, 17)), "### ### ### ###")
                fgFLUTPJSTAT.Col = 5: fgFLUTPJSTAT.Text = Format(Val(mId$(X, 112, 17)), "### ### ### ###")
              rsSab.MoveNext
            Loop
            fgFLUTPJSTAT.Visible = True

        End If
    Case "FLUTPJCLI"
            txtParam_K2.Visible = True
            txtParam_K2 = Old_YBIATAB0.BIATABK2
            cboParam_CCB_DB.Visible = False
            cboParam_CCB_CR.Visible = False
            txtParam_X = "Indiquer le code opé  et la racine client"
    Case Else:
        txtParam_K2.Visible = False
End Select

mnuParam_Option = ""
cmdParam_Ok.Visible = False
fraParam_K2.Enabled = False
fraParam_K.Visible = True
fraParam_K.Caption = ""
mnuParam_Delete.Enabled = True
If lstParam_Action_K = "FLUTPJOPE" And txtParam_K2 = "" Then
    Select Case txtParam_K
        Case "TRE", "CDO", "DAT", "HG", "CRE", "REM", "NOS": mnuParam_Delete.Enabled = False
    End Select
End If
If YFLUTPJ0_Aut.Comptabiliser Then Me.PopupMenu mnuParam, , 10000, 8300

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

Select Case SSTab1.Tab
    Case 0
        
        If fraDetail.Visible Then
            fraDetail.Visible = False
            Exit Sub
        End If
        
        If fgSelect.Visible Then
            fgSelect.Visible = False
            Exit Sub
        End If
        
        Unload Me
    Case 1
        If fraFlux_Update.Visible Then
            fraFlux_Update.Visible = False
            Exit Sub
        End If
        If fraFlux_Detail.Visible Then
            fraFlux_Detail.Visible = False
            Exit Sub
        End If
        If fgFlux.Visible Then
            fgFlux.Visible = False
            Exit Sub
        End If
        SSTab1.Tab = 0
    Case 2
        If fraParam_K.Visible Then
            fraParam_K.Visible = False
            Exit Sub
        End If
        If lstParam_K.Visible Then
            lstParam_K.Visible = False
            Exit Sub
        End If
        SSTab1.Tab = 0
End Select

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
Dim xDetail As String, xSQL As String, xWhere As String
Dim xAMJ As String
On Error Resume Next

Select Case cmdSelect_SQL_K
    Case "1"
        xDetail = "Flux "
        If y <= fgSelect.RowHeightMin Then
            If X < 3100 Then
                xWhere = " "
                xDetail = "Tous les flux"
            Else
                xWhere = " and FLUTPJDEV = '" & arrDev(fgSelect.Col - 1) & "'"
                xDetail = xDetail & " - devise = " & arrDev(fgSelect.Col - 1)
            End If
        Else
            If X >= 3100 Then
                xWhere = " and FLUTPJDEV = '" & arrDev(fgSelect.Col - 1) & "'"
                xDetail = xDetail & " - devise = " & arrDev(fgSelect.Col - 1)
            End If
            Select Case fgSelect.Row
                Case mCCB_DB_Row: xWhere = xWhere & " and FLUTPJCCB like '1%'"
                                  xDetail = Replace(xDetail, "Flux", "Décaissements")
                Case mCCB_CR_Row: xWhere = xWhere & " and FLUTPJCCB like '2%'"
                                  xDetail = Replace(xDetail, "Flux", "Encaissements")
                Case mCCB_SD_Row:
                                   xDetail = Replace(xDetail, "Flux", "Tous les flux")
               Case Else: xWhere = xWhere & " and FLUTPJCCB = " & arrCCB_K(fgSelect.Row)
                                  xDetail = Replace(xDetail, "Flux", arrCCB_Lib(fgSelect.Row))
           End Select
            'Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        
        End If
        fraDetail.Caption = xDetail
        arrYFLUTPJ0_SQL mSelect_Where & xWhere & " order by FLUTPJDEV , FLUTPJECH"
        fgDetail_Display
'________________________________________________________________________
    Case "2"
        xDetail = "Flux "
        If y <= fgSelect.RowHeightMin Then
            If X < 1150 Then
                xWhere = " "
                xDetail = "Tous les flux"
            Else
                xWhere = " and FLUTPJDEV = '" & arrDev(fgSelect.Col - 1) & "'"
                xDetail = xDetail & " - devise = " & arrDev(fgSelect.Col - 1)
            End If
        Else
            If X >= 1150 Then
                xWhere = " and FLUTPJDEV = '" & arrDev(fgSelect.Col - 1) & "'"
                xDetail = xDetail & " - devise = " & arrDev(fgSelect.Col - 1)
            End If
            fgSelect.Col = 0:  Call dateJMA_AMJ(Trim(fgSelect.Text), xAMJ)
            xWhere = xWhere & " and FLUTPJECH = " & xAMJ
            xDetail = Replace(xDetail, "Flux", "Echéance : " & fgSelect.Text)
           ' Call fgselect_Color(fgselect_RowClick, MouseMoveUsr.BackColor, fgselect_ColorClick)
        
        End If
        fraDetail.Caption = xDetail
        arrYFLUTPJ0_SQL mSelect_Where & xWhere & " order by FLUTPJDEV , FLUTPJECH"
        fgDetail_Display
        
    Case "3"
        fraDetail.Caption = "EMP € "
        fgDetail_Display_3

End Select
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




Public Sub fgFlux_Detail_Reset()
fgFlux_Detail.Clear
fgFlux_Detail_Sort1 = 0: fgFlux_Detail_Sort2 = 0
fgFlux_Detail_Sort1_Old = -1
fgFlux_Detail_RowDisplay = 0: fgFlux_Detail_RowClick = 0
fgFlux_Detail_arrIndex = fgFlux_Detail.Cols - 1
blnfgFlux_Detail_DisplayLine = False
fgFlux_Detail_SortAD = 6
fgFlux_Detail.LeftCol = fgFlux_Detail.FixedCols

End Sub

Public Sub fgFlux_Reset()
fgFlux.Clear
fgFlux_Sort1 = 0: fgFlux_Sort2 = 0
fgFlux_Sort1_Old = -1
fgFlux_RowDisplay = 0: fgFlux_RowClick = 0
fgFlux_arrIndex = fgFlux.Cols - 1
blnfgFlux_DisplayLine = False
fgFlux_SortAD = 6
fgFlux.LeftCol = fgFlux.FixedCols

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







Private Sub mnuCB_Export_Click()
Dim wFile As String
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Option = "mnuCB_Export"
Call YFLUTPJ0_CB_Export(True, wFile)

Me.Enabled = True: Me.MousePointer = 0
End Sub

Public Function colorHex_RGB(lColor As Long) As Long
Dim xColor As String, X As String
Dim lRed As Integer, lGreen As Integer, LBlue As Integer
lRed = lColor Mod 256
lGreen = Int(lColor / 256) Mod 256
LBlue = Int(lColor / 65536) Mod 256
colorHex_RGB = RGB(lRed, lGreen, LBlue)
End Function

Private Sub mnuECH_export_Click()
Dim wFile As String
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Option = "mnuECH_Export"
Call YFLUTPJ0_ECH_Export(True, wFile)
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuECH_Mail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Option = "mnuECH_Mail"
cmdSendMail_Ech
Me.Enabled = True: Me.MousePointer = 0
End Sub


Private Sub mnuFlux_Add_Click()
fraFlux_Update = "Nouveau flux"
mnuFlux_Option = "mnuFlux_Add"
cmdFlux_Update_Ok.Visible = YFLUTPJ0_Aut.Valider
fraFlux_Update_A.Enabled = True
fraFlux_Update_B.Enabled = True
fraFlux_UPDATE_C.Enabled = True
End Sub

Private Sub mnuFlux_Annulation_Click()
fraFlux_Update = "Annulation du flux"
mnuFlux_Option = "mnuFlux_Annulation"
cmdFlux_Update_Ok.Visible = YFLUTPJ0_Aut.Valider
fraFlux_Update_A.Enabled = False
fraFlux_Update_B.Enabled = False
fraFlux_UPDATE_C.Enabled = False

End Sub

Private Sub mnuFlux_Close_Click()
fraFlux_Update = "Clôture du dossier"
mnuFlux_Option = "mnuFlux_Close"
cmdFlux_Update_Ok.Visible = YFLUTPJ0_Aut.Valider
fraFlux_Update_A.Enabled = False
fraFlux_Update_B.Enabled = False
fraFlux_UPDATE_C.Enabled = False

End Sub

Private Sub mnuFlux_Delete_Click()
fraFlux_Update = "suppression du flux"
mnuFlux_Option = "mnuFlux_Delete"
cmdFlux_Update_Ok.Visible = YFLUTPJ0_Aut.Valider
fraFlux_Update_A.Enabled = False
fraFlux_Update_B.Enabled = False
fraFlux_UPDATE_C.Enabled = False

End Sub

Private Sub mnuFlux_FLUTPJTXT_Click()
fraFlux_Update = "Mise à jour du commentaire"
mnuFlux_Option = "mnuFlux_FLUTPJTXT"
cmdFlux_Update_Ok.Visible = YFLUTPJ0_Aut.Valider
fraFlux_Update_A.Enabled = False
fraFlux_Update_B.Enabled = False
fraFlux_UPDATE_C.Enabled = True

End Sub

Private Sub mnuFlux_Update_Click()
fraFlux_Update = "modification du flux"
mnuFlux_Option = "mnuFlux_Update"
cmdFlux_Update_Ok.Visible = YFLUTPJ0_Aut.Valider
fraFlux_Update_A.Enabled = False
fraFlux_Update_B.Enabled = True
fraFlux_UPDATE_C.Enabled = True

End Sub


Private Sub mnuParam_Add_Click()
fraParam_K = "Nouvel enregistrement"
'txtParam_K = ""
'txtParam_K2 = ""
mnuParam_Option = "mnuParam_Add"
cmdParam_Ok.Visible = YFLUTPJ0_Aut.Comptabiliser
fraParam_K2.Enabled = True
txtParam_K.Enabled = True
txtParam_K2.Enabled = True
End Sub

Private Sub mnuParam_Delete_Click()
Dim Nb As Long
mnuParam_Option = "mnuParam_Delete"
fraParam_K = "Suppresion de l'enregistrement"
fraParam_K2.Enabled = False

Select Case lstParam_Action_K
    Case "FLUTPJCCB"
        X = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YFLUTPJ0 " _
            & " where FLUTPJCCB = '" & Trim(Old_YBIATAB0.BIATABK1) & "'"
        Set rsSab = cnsab.Execute(X)
        Nb = rsSab("Tally")
    Case "FLUTPJOPE", "FLUTPJOD", "FLUTPJCPT"
        X = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YFLUTPJ0 " _
            & " where FLUTPJOPE = '" & Trim(Old_YBIATAB0.BIATABK1) & "'"
        Set rsSab = cnsab.Execute(X)
        Nb = rsSab("Tally")
    Case Else
    Nb = 0
End Select

If Nb = 0 Then
    cmdParam_Ok.Visible = YFLUTPJ0_Aut.Comptabiliser
Else
    Call MsgBox("Suppression impossible : il y a " & Nb & " flux rattachés à ce code", vbCritical, "FLUX : paramétrage")
End If

End Sub

Private Sub mnuParam_Export_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Option = "mnuParam_Export"
YFLUTPJ0_Param_Export
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuParam_Update_Click()
txtParam_K.Enabled = False
txtParam_K2.Enabled = False
fraParam_K = "Mise à jour de l'enregistrement"
mnuParam_Option = "mnuParam_Update"
cmdParam_Ok.Visible = YFLUTPJ0_Aut.Comptabiliser
fraParam_K2.Enabled = True

End Sub


Private Sub mnuPrint_Detail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdPrint_YFLUTPJ0 True

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint_Recap_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdPrint_YFLUTPJ0 False

Me.Enabled = True: Me.MousePointer = 0
End Sub




















Private Sub txtParam_K_KeyPress(KeyAscii As Integer)
If lstParam_Action_K = "FLUTPJCCB" Then
    KeyAscii = ctlNum(KeyAscii)
Else
    KeyAscii = convUCase(KeyAscii)
End If
End Sub


Private Sub txtFlux_FLUTPJMTD_GotFocus()
txtFlux_FLUTPJMTD.BackColor = focusUsr.BackColor
End Sub


Private Sub txtFlux_FLUTPJMTD_KeyPress(KeyAscii As Integer)
    Call num_Montant(KeyAscii, txtFlux_FLUTPJMTD)

End Sub


Private Sub txtFlux_FLUTPJMTD_LostFocus()
txtFlux_FLUTPJMTD.BackColor = txtUsr.BackColor

End Sub


Private Sub txtParam_K2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtSelect_FLUTPJECH_Max_Change()
Call DTPicker_Control(txtSelect_FLUTPJECH_Max, arrAMJMAX(cmdSelect_SQL_K))
cmdSelect_Reset


End Sub

Private Sub txtSelect_FLUTPJECH_Max_Click()
cmdSelect_Reset


End Sub

Private Sub txtSelect_FLUTPJECH_Min_Change()
Call DTPicker_Control(txtSelect_FLUTPJECH_Min, arrAMJMIN(cmdSelect_SQL_K))
cmdSelect_Reset


End Sub

Private Sub txtSelect_FLUTPJECH_Min_Click()
cmdSelect_Reset


End Sub





Public Sub Parametrage_FLUTPJCCB_Init()
Dim xSQL As String
ReDim arrParam(1000)
lstParam_Action_K = "FLUTPJCCB"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJCCB' order by BIATABK1"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call rsYBIATAB0_GetBuffer(rsSab, xYBIATAB0)
    lstParam_K.AddItem Trim(xYBIATAB0.BIATABK1) & " : " & Trim(xYBIATAB0.BIATABTXT)
    arrParam(lstParam_K.ListCount) = xYBIATAB0
    rsSab.MoveNext
Loop
ReDim Preserve arrParam(lstParam_K.ListCount)

txtParam_K.Enabled = True
txtParam_X.Enabled = True
cboParam_CCB_DB.Visible = False: lblParam_CCB_DB.Visible = False
cboParam_CCB_CR.Visible = False: lblParam_CCB_CR.Visible = False
cboParam_Frequence.Visible = False: lblParam_Frequence.Visible = False


lstParam_K.Visible = True
End Sub

Public Sub Parametrage_FLUTPJOPE_Init()
Dim xSQL As String, K As Integer, X As String, X1 As String, X2 As String
ReDim arrParam(1000)

lstParam_Action_K = "FLUTPJOPE"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJOPE' order by BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call rsYBIATAB0_GetBuffer(rsSab, xYBIATAB0)
    X1 = mId$(xYBIATAB0.BIATABK1, 1, 3)
    X2 = mId$(xYBIATAB0.BIATABK2, 1, 3)
    X = X1 & X2
    For K = 1 To arrOPE_Nb
        If X = arrOPE_K(K) Then Exit For
    Next K
    lstParam_K.AddItem X1 & " - " & X2 & " : " & Trim(xYBIATAB0.BIATABTXT) & " " & arrOPE_Lib(K)
    
    arrParam(lstParam_K.ListCount) = xYBIATAB0
    rsSab.MoveNext
Loop

ReDim Preserve arrParam(lstParam_K.ListCount)
txtParam_K.Enabled = True
txtParam_X.Enabled = True
cboParam_CCB_DB.Visible = True: lblParam_CCB_DB.Visible = True
cboParam_CCB_CR.Visible = True: lblParam_CCB_CR.Visible = True
cboParam_Frequence.Visible = False: lblParam_Frequence.Visible = False

lstParam_K.Visible = True

End Sub

Public Sub Parametrage_FLUTPJCPT_Init()
Dim xSQL As String, K As Integer, X As String, X1 As String, X2 As String

ReDim arrParam(1000)

    

lstParam_Action_K = "FLUTPJCPT"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJCPT' order by BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call rsYBIATAB0_GetBuffer(rsSab, xYBIATAB0)
    X1 = mId$(xYBIATAB0.BIATABK1, 1, 10)
    X2 = mId$(xYBIATAB0.BIATABK2, 1, 5)
    X = Trim(mId$(X1, 1, 3))
    For K = 1 To arrOPE_Nb
        If X = Trim(arrOPE_K(K)) Then Exit For
    Next K
    lstParam_K.AddItem X1 & X2 & " : " & mId$(xYBIATAB0.BIATABTXT, 1, 10) & " " & arrOPE_Lib(K)
    
    arrParam(lstParam_K.ListCount) = xYBIATAB0
    rsSab.MoveNext
Loop
ReDim Preserve arrParam(lstParam_K.ListCount)

txtParam_K.Enabled = True
txtParam_X.Enabled = True
cboParam_CCB_DB.Visible = True: lblParam_CCB_DB.Visible = True
cboParam_CCB_CR.Visible = True: lblParam_CCB_CR.Visible = True
cboParam_Frequence.Visible = False: lblParam_Frequence.Visible = False

lstParam_K.Visible = True

End Sub

Public Sub Parametrage_FLUTPJCLI_Init()
Dim xSQL As String, K As Integer, X As String, X1 As String, X2 As String

ReDim arrParam(1000)

lstParam_Action_K = "FLUTPJCLI"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJCLI' order by BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call rsYBIATAB0_GetBuffer(rsSab, xYBIATAB0)
    X1 = mId$(xYBIATAB0.BIATABK1, 1, 3)
    X2 = mId$(xYBIATAB0.BIATABK2, 1, 7)
    xSQL = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & X2 & "'"
    Set rsSabX = cnsab.Execute(xSQL)
    If Not rsSabX.EOF Then
        X = Trim(rsSabX("CLIENARA1"))
    Else
        X = "???"
    End If

    lstParam_K.AddItem X1 & " - " & X2 & " : " & X
     arrParam(lstParam_K.ListCount) = xYBIATAB0
   
    rsSab.MoveNext
Loop
ReDim Preserve arrParam(lstParam_K.ListCount)

txtParam_K.Enabled = True
txtParam_X.Enabled = True
cboParam_CCB_DB.Visible = False: lblParam_CCB_DB.Visible = False
cboParam_CCB_CR.Visible = False: lblParam_CCB_CR.Visible = False
cboParam_Frequence.Visible = False: lblParam_Frequence.Visible = False

lstParam_K.Visible = True

End Sub


Public Sub Parametrage_FLUTPJOD_Init()
Dim xSQL As String
ReDim arrParam(1000)




lstParam_Action_K = "FLUTPJOD"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJOD' order by BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call rsYBIATAB0_GetBuffer(rsSab, xYBIATAB0)
    lstParam_K.AddItem mId$(xYBIATAB0.BIATABK1, 1, 3) & " : " & Trim(xYBIATAB0.BIATABTXT)
    
    arrParam(lstParam_K.ListCount) = xYBIATAB0
    rsSab.MoveNext
Loop
ReDim Preserve arrParam(lstParam_K.ListCount)

txtParam_K.Enabled = True
txtParam_X.Enabled = True
cboParam_CCB_DB.Visible = True: lblParam_CCB_DB.Visible = True
cboParam_CCB_CR.Visible = True: lblParam_CCB_CR.Visible = True
cboParam_Frequence.Visible = True: lblParam_Frequence.Visible = True

lstParam_K.Visible = True

End Sub

Public Sub parametrage_FLUTPJCCB_Load()
Dim xSQL As String
'Initialisation opération________________________________________________________________________________
arrCCB_Nb = 0
ReDim Preserve arrCCB_K(1000), arrCCB_Lib(1000)
cboParam_CCB_DB.Clear: cboParam_CCB_DB.AddItem "000 -"
cboParam_CCB_CR.Clear: cboParam_CCB_CR.AddItem "000 -"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJCCB' order by BIATABK1"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    arrCCB_Nb = arrCCB_Nb + 1
    arrCCB_K(arrCCB_Nb) = Trim(rsSab("BIATABK1"))
    arrCCB_Lib(arrCCB_Nb) = Trim(rsSab("BIATABTXT"))
    If arrCCB_K(arrCCB_Nb) < 200 Then
        cboParam_CCB_DB.AddItem arrCCB_K(arrCCB_Nb) & " - " & arrCCB_Lib(arrCCB_Nb)
    Else
        cboParam_CCB_CR.AddItem arrCCB_K(arrCCB_Nb) & " - " & arrCCB_Lib(arrCCB_Nb)
    End If
    If arrCCB_K(arrCCB_Nb) = 104 Or arrCCB_K(arrCCB_Nb) = 105 Then
         cboParam_CCB_CR.AddItem arrCCB_K(arrCCB_Nb) & " - " & arrCCB_Lib(arrCCB_Nb)
    End If
    rsSab.MoveNext
Loop
ReDim Preserve arrCCB_K(arrCCB_Nb + 1), arrCCB_Lib(arrCCB_Nb + 1)

End Sub

Public Function cmdFLUX_Transaction(lFct As String)
Dim V, V2, X As String, xSQL As String
Dim Nb As Long
Dim mMsgBox As String
Dim blnInsert As Boolean
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdFLUX_Transaction"
'-------------------------------------------------------
cmdFLUX_Transaction = Null
fgSelect.Visible = False: fraDetail.Visible = False
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case lFct
    Case "Insert": V = sqlYFLUTPJ0_Insert(newFlux)
                   If IsNull(V) Then
                        If Trim(newYFLUTPJ1.FLUTPJTXT) <> "" Then V = sqlYFLUTPJ1_Insert(newYFLUTPJ1)
                    End If
    Case "Insert+":
                    V = sqlYFLUTPJ0_Insert(newFlux)
                    If IsNull(V) Then
                        V = cmdFLUX_Transaction_Sequence
                        If IsNull(V) Then
                             If Trim(newYFLUTPJ1.FLUTPJTXT) <> "" Then V = sqlYFLUTPJ1_Insert(newYFLUTPJ1)
                         End If
                    End If
    Case "Update": V = sqlYFLUTPJ0_Update(newFlux, oldFlux)
                   If IsNull(V) Then
                        If Trim(newYFLUTPJ1_Detail.FLUTPJTXT) = "" Then
                            If Trim(oldYFLUTPJ1_Detail.FLUTPJTXT) <> "" Then V = sqlYFLUTPJ1_Delete(oldYFLUTPJ1_Detail)
                        Else
                            If Trim(oldYFLUTPJ1_Detail.FLUTPJTXT) = "" Then
                                V = sqlYFLUTPJ1_Insert(newYFLUTPJ1_Detail)
                            Else
                                V = sqlYFLUTPJ1_Update(newYFLUTPJ1_Detail, oldYFLUTPJ1_Detail)
                            End If
                            
                        End If
                   End If
    Case "Delete": V = sqlYFLUTPJ0_Delete(oldFlux)
                   If IsNull(V) Then
                        If Trim(oldYFLUTPJ1_Detail.FLUTPJTXT) <> "" Then V = sqlYFLUTPJ1_Delete(oldYFLUTPJ1_Detail)
                   End If
    Case "Delete_YFLUTPJ1": V = sqlYFLUTPJ0_Delete(oldFlux)
                   If IsNull(V) Then
                        V = sqlYFLUTPJ1_Delete_FLUTPJDOS(oldYFLUTPJ1_Detail)
                   End If
                   
    Case "Close":
                    V = sqlYFLUTPJ0_Update(newFlux, oldFlux)
                    If IsNull(V) Then V = cmdFlux_Transaction_Close
                    
     Case "Update_FLUTPJTXT":
                If Trim(newYFLUTPJ1.FLUTPJTXT) <> "" Then V = sqlYFLUTPJ1_Update(newYFLUTPJ1, oldYFLUTPJ1)

End Select

If Not IsNull(V) Then GoTo Error_MsgBox
GoTo Exit_sub
'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V2 = cnSAB_Transaction("Rollback")
    Else
        V2 = cnSAB_Transaction("Commit")
        cmdFlux_Reset
        cmdFlux_SQL_1
    End If
    
    cmdFLUX_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function
Public Sub Error_Route(V)

currentError = CStr(V) & "             ( " & Me.Name & " ~ " & App_Debug & " )"
If blnAuto Then
  '  Call cmdSendMail_Alerte(Me.Name & " ~ " & App_Debug, CStr(V))
Else
    MsgBox V, vbCritical, Me.Name & " ~ " & App_Debug
End If

End Sub

Public Sub parametrage_FLUTPJOD_Load()
Dim xSQL As String
'Initialisation opération________________________________________________________________________________
cboFlux_FLUTPJOD.Clear
cboFlux_FLUTPJOD.AddItem ""
ReDim Preserve arrOD_K(1000), arrOD_Lib(1000)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJOD' order by BIATABK1"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    arrOD_Nb = arrOD_Nb + 1
    arrOD_K(arrOD_Nb) = Trim(rsSab("BIATABK1"))
    arrOD_Lib(arrOD_Nb) = Trim(mId$(rsSab("BIATABTXT"), 11, 64))
    cboFlux_FLUTPJOD.AddItem Trim(rsSab("BIATABK1")) & " - " & Trim(mId$(rsSab("BIATABTXT"), 11, 64))
    
    rsSab.MoveNext
Loop
ReDim Preserve arrOD_K(arrOD_Nb + 1), arrOD_Lib(arrOD_Nb + 1)

End Sub


Public Function parametrage_Control()
Dim wMsg As String, xLib As String, X As String

parametrage_Control = "?"
wMsg = ""
New_YBIATAB0 = Old_YBIATAB0
If Trim(txtParam_K) = "" Then wMsg = wMsg & vbCrLf & "- préciser le code"
New_YBIATAB0.BIATABK1 = Trim(txtParam_K)
xLib = Trim(txtParam_X)

Select Case lstParam_Action_K
    Case "FLUTPJCCB"
        If New_YBIATAB0.BIATABK1 < "100" Then wMsg = wMsg & vbCrLf & "- le code CB doit être supérieur à 100"
        If New_YBIATAB0.BIATABK1 > "300" Then wMsg = wMsg & vbCrLf & "- le code CB doit être inférieur à 300"
        If xLib = "" Then wMsg = wMsg & vbCrLf & "- préciser le libellé"
        New_YBIATAB0.BIATABTXT = xLib
    Case "FLUTPJOPE"
        New_YBIATAB0.BIATABK2 = Trim(txtParam_K2)
        If New_YBIATAB0.BIATABK1 = "" Then wMsg = wMsg & vbCrLf & "- préciser le code opération SAB"
        If Len(Trim(New_YBIATAB0.BIATABK1)) > 3 Then wMsg = wMsg & vbCrLf & "- le code opération SAB est trop long (<= 3 caractères)"
        New_YBIATAB0.BIATABTXT = mId$(cboParam_CCB_DB, 1, 3) & " " & mId$(cboParam_CCB_CR, 1, 3)
    Case "FLUTPJOD"
        If New_YBIATAB0.BIATABK1 = "" Then wMsg = wMsg & vbCrLf & "- préciser le code opération SAB"
        If Len(Trim(New_YBIATAB0.BIATABK1)) > 3 Then wMsg = wMsg & vbCrLf & "- le code opération SAB est trop long (<= 3 caractères)"
        If xLib = "" Then wMsg = wMsg & vbCrLf & "- préciser le libellé"
       
        New_YBIATAB0.BIATABTXT = mId$(cboParam_CCB_DB, 1, 3) & " " & mId$(cboParam_CCB_CR, 1, 3) _
                               & " " & mId$(cboParam_Frequence, 1, 1) & " " & xLib
    Case "FLUTPJCPT"
       If New_YBIATAB0.BIATABK1 = "" Then wMsg = wMsg & vbCrLf & "- préciser le code opération SAB"
       If Len(Trim(New_YBIATAB0.BIATABK1)) <> 9 Then wMsg = wMsg & vbCrLf & "- '*** SS ss' = code opération & Service & sous-service (9 caractères) "
       Mid$(New_YBIATAB0.BIATABK1, 4, 1) = " ": Mid$(New_YBIATAB0.BIATABK1, 7, 1) = " "
       
       If Trim(Old_YBIATAB0.BIATABK2) = "=" Then
            X = Trim(txtParam_K2)
            If X = "" Then wMsg = wMsg & vbCrLf & "- préciser le code opération équivalent"
            If Len(Trim(X)) <> 9 Then wMsg = wMsg & vbCrLf & "- '*** SS ss' = code opération & Service & sous-service (9 caractères) "
            New_YBIATAB0.BIATABTXT = X
            Mid$(New_YBIATAB0.BIATABTXT, 4, 1) = " ": Mid$(New_YBIATAB0.BIATABTXT, 7, 1) = " "
            New_YBIATAB0.BIATABK2 = "="
        Else
            New_YBIATAB0.BIATABTXT = Trim(txtParam_K2)
            New_YBIATAB0.BIATABTXT = mId$(cboParam_CCB_DB, 1, 3) & " " & mId$(cboParam_CCB_CR, 1, 3)
        End If
    Case "FLUTPJCLI"
        If New_YBIATAB0.BIATABK1 = "" Then wMsg = wMsg & vbCrLf & "- préciser le code flux complémentaire"
        If Len(Trim(New_YBIATAB0.BIATABK1)) > 3 Then wMsg = wMsg & vbCrLf & "- le code flux complémentaire est trop long (<= 3 caractères)"
        New_YBIATAB0.BIATABK2 = Format$(Trim(txtParam_K2), "0000000")
        If Trim(txtParam_K2) = "" Then wMsg = wMsg & vbCrLf & "- préciser la racine client"
        New_YBIATAB0.BIATABK2 = Format$(Trim(txtParam_K2), "0000000")
        
        X = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & New_YBIATAB0.BIATABK2 & "'"
        Set rsSabX = cnsab.Execute(X)
        If Not rsSabX.EOF Then
            New_YBIATAB0.BIATABTXT = Trim(rsSabX("CLIENARA1"))
        Else
            wMsg = wMsg & vbCrLf & "- racine client inconnue"
        End If

    
End Select
If wMsg = "" Then
    parametrage_Control = Null
Else
    Call MsgBox(wMsg, vbExclamation, "Flux Tresorerie prévisionnelle : paramétrage")
End If

End Function

Public Sub fraFlux_Display()
Dim xSQL As String, X As String

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJOD' and BIATABK1 = '" & Trim(oldFlux.FLUTPJOPE) & "'"
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    Call MsgBox("erreur lecture : " & vbCrLf & xSQL, vbCritical, currentAction)
Else
    Call rsYBIATAB0_GetBuffer(rsSab, paramOD)
End If
optFlux_Decaissement.Value = IIf(mId$(paramOD.BIATABTXT, 1, 3) = "000", 0, 1)
optFlux_Encaissement.Value = IIf(mId$(paramOD.BIATABTXT, 5, 3) = "000", 0, 1)

fraFlux_Update.Caption = oldFlux.FLUTPJDEV & " - " & oldFlux.FLUTPJOPE & Trim(mId$(paramOD.BIATABTXT, 11, 64)) & " - dossier : " & oldFlux.FLUTPJDOS & " - " & oldFlux.FLUTPJDOSQ
If oldFlux.FLUTPJMTD < 0 Then optFlux_Decaissement.Value = True
If oldFlux.FLUTPJMTD <> 0 Then
    txtFlux_FLUTPJMTD = Format$(Abs(oldFlux.FLUTPJMTD), "### ### ### ##0.00")
Else
    txtFlux_FLUTPJMTD = ""
End If

X = oldFlux.FLUTPJECH
Call DTPicker_Set(txtFlux_FLUTPJECH, X) '
cbo_Scan oldFlux.FLUTPJEVE, cboFlux_Frequence

fraFlux_Display_FLUTPJTXT

txtFLUX_FLUTPJTXT = Trim(xYFLUTPJ1_Detail.FLUTPJTXT)
oldYFLUTPJ1_Detail = xYFLUTPJ1_Detail

fraFlux_Update.Visible = True 'YFLUTPJ0_Aut.Comptabiliser
End Sub

Public Function fraflux_Control()
Dim V, X As String, blnOk As Boolean, K As Integer, wMsgBox As String

wMsgBox = ""
blnOk = False
fraflux_Control = Null

newFlux = oldFlux
Call DTPicker_Control(txtFlux_FLUTPJECH, X)
If X < YBIATAB0_DATE_CPT_JS1 Then
    wMsgBox = "- la date d'échéance ne doit pas être inférieure au " & dateImp10(YBIATAB0_DATE_CPT_JS1)
Else
    newFlux.FLUTPJECH = X
End If
newFlux.FLUTPJSER = "CP"
newFlux.FLUTPJSSE = "CP"
newFlux.FLUTPJEVE = mId$(cboFlux_Frequence, 1, 1)
If Trim(txtFlux_FLUTPJMTD) = "" Then
    wMsgBox = "- préciser le montant"
Else
    If optFlux_Decaissement = True Then
        newFlux.FLUTPJCCB = mId$(paramOD.BIATABTXT, 1, 3)
        Call cbo_Scan(mId$(paramOD.BIATABTXT, 5, 3), cboParam_CCB_CR)
        newFlux.FLUTPJMTD = -CCur(txtFlux_FLUTPJMTD)
    Else
        newFlux.FLUTPJCCB = mId$(paramOD.BIATABTXT, 5, 3)
        newFlux.FLUTPJMTD = CCur(txtFlux_FLUTPJMTD)
    End If
End If
'_________________________________________________________________________________

X = Trim(txtFLUX_FLUTPJTXT)
If mnuFlux_Option = "mnuFlux_FLUTPJTXT" Or mnuFlux_Option = "mnuFlux_Add" Then
    If Trim(txtFLUX_FLUTPJTXT) = "" Then
        wMsgBox = "- préciser le commentaire"
    Else
        newYFLUTPJ1 = oldYFLUTPJ1
        newYFLUTPJ1.FLUTPJTXT = X
    End If
Else
    newYFLUTPJ1_Detail = oldYFLUTPJ1_Detail
    newYFLUTPJ1_Detail.FLUTPJTXT = X
End If
X = Trim(cboFlux_FLUTPJCLI)
If X = "" Then
    newYFLUTPJ1.FLUTPJCLI = ""
Else
    K = InStr(X & " ", " ")
    newYFLUTPJ1.FLUTPJCLI = Format$(Val(mId$(X, 1, K)), "0000000")
End If
'_________________________________________________________________________________

If wMsgBox <> "" Then
    fraflux_Control = "?"
    Call MsgBox(wMsgBox, vbCritical, "Flux : contrôle détail")
End If

Exit_sub:

End Function

Public Function cmdFLUX_Transaction_Sequence()
Dim blnExit As Boolean
Dim V, xAMJ As String
V = Null
If newFlux.FLUTPJEVE <> "U" Then
    Do
        If newFlux.FLUTPJECH > mDate_J7 Then
            blnExit = True
        Else
            newFlux.FLUTPJID = newFlux.FLUTPJID + 1
            newFlux.FLUTPJDOSQ = newFlux.FLUTPJDOSQ + 1
            Select Case newFlux.FLUTPJEVE
                Case "J": newFlux.FLUTPJECH = dateElp("Ouvré", 1, newFlux.FLUTPJECH)
                Case "M": newFlux.FLUTPJECH = dateElp("MoisAdd", 1, newFlux.FLUTPJECH)
                Case "T": newFlux.FLUTPJECH = dateElp("TrimestreAdd", 1, newFlux.FLUTPJECH)
                Case "S": newFlux.FLUTPJECH = dateElp("SemestreAdd", 1, newFlux.FLUTPJECH)
                Case "A": newFlux.FLUTPJECH = dateElp("AnAdd", 1, newFlux.FLUTPJECH)
                Case Else: blnExit = True
            End Select
            
            V = sqlYFLUTPJ0_Insert(newFlux)
            If Not IsNull(V) Then blnExit = True
        End If
    Loop Until blnExit
End If
cmdFLUX_Transaction_Sequence = V
End Function

Public Function cmdFlux_Transaction_Close()
Dim K As Integer, mFLUTPJDOSQ As Long
cmdFlux_Transaction_Close = "?"
mFLUTPJDOSQ = oldFlux.FLUTPJDOSQ
For K = 1 To arrfgFlux_Detail_Nb
    If arrfgFlux_Detail(K).FLUTPJDOSQ > mFLUTPJDOSQ Then
        oldFlux = arrfgFlux_Detail(K)
        V = sqlYFLUTPJ0_Delete(oldFlux)
        If Not IsNull(V) Then cmdFlux_Transaction_Close = V: Exit Function

    End If
Next K
cmdFlux_Transaction_Close = Null
End Function

Public Sub arrDev_Cours_Load(lAMJ As String)
Dim xMemo As String, K As Integer
Static currentAMJ As String
If lAMJ <> currentAMJ Then
    If lAMJ > YBIATAB0_DATE_CPT_J Then
        mAMJDEV = YBIATAB0_DATE_CPT_J
    Else
        mAMJDEV = lAMJ
    End If
    For K = 1 To arrDev_Nb
        If arrDev(K) = "EUR" Then
            arrDev_Cours(K) = 1
        Else
            arrDev_Cours(K) = 0
        End If
        Call sqlYBIATAB0_Read("PDC", arrDev(K), mAMJDEV, xMemo)
        If Not IsNumeric(mId$(xMemo, 9, 15)) Then Call sqlYBIATAB0_Read("FIXING", arrDev(K), "J", xMemo)

        If IsNumeric(mId$(xMemo, 9, 15)) Then arrDev_Cours(K) = CDbl(mId$(xMemo, 9, 15) / 1000000000)
    Next K
    currentAMJ = lAMJ

End If
End Sub

Public Sub auto_YFLUTPJ0()
Dim V2
Dim xSQL As String
Dim mFLUTPJID As Long
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "auto_YFLUTPJ0"
'_________________________________________________________________________________
xSQL = "select FLUTPJID from " & paramIBM_Library_SABSPE_XXX & ".YFLUTPJ0 " _
     & "  where FLUTPJORIG = '*' order by FLUTPJID desc"
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    mFLUTPJID = 1 '
Else
    mFLUTPJID = rsSab("FLUTPJID")
End If
'_________________________________________________________________________________________
'-------------------------------------------------------
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

rsYFLUTPJ0_Init newFlux: newFlux.FLUTPJECH = 99999999
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YFLUTPJ0" _
     & " where FLUTPJORIG = '*' and FLUTPJECH > " & YBIATAB0_DATE_CPT_JP1 _
     & " order by FLUTPJDOS , FLUTPJDOSQ"
     
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYFLUTPJ0_GetBuffer(rsSab, xYFLUTPJ0)
    If newFlux.FLUTPJDOS <> xYFLUTPJ0.FLUTPJDOS Then
        If newFlux.FLUTPJECH <= mDate_J7 And newFlux.FLUTPJSTA <> "X" Then
            newFlux.FLUTPJID = mFLUTPJID
            cmdFLUX_Transaction_Sequence
            mFLUTPJID = newFlux.FLUTPJID
        End If
    End If
    newFlux = xYFLUTPJ0
    
    rsSab.MoveNext
Loop

If newFlux.FLUTPJECH <= mDate_J7 And newFlux.FLUTPJSTA <> "X" Then
    newFlux.FLUTPJID = mFLUTPJID
    cmdFLUX_Transaction_Sequence
    mFLUTPJID = newFlux.FLUTPJID
End If
GoTo Exit_sub
'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V2 = cnSAB_Transaction("Rollback")
    Else
        V2 = cnSAB_Transaction("Commit")
        cmdFlux_Reset
        'cmdFlux_SQL_1
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Sub

Public Sub fgFlux_Display_FLUTPJTXT()
Dim xSQL As String

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YFLUTPJ1 where FLUTPJDOS = " & xFlux.FLUTPJDOS & " and FLUTPJDOSQ = 0"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    V = rsYFLUTPJ1_GetBuffer(rsSab, xYFLUTPJ1)
Else
    xYFLUTPJ1.FLUTPJDOS = xFlux.FLUTPJDOS
    xYFLUTPJ1.FLUTPJDOSQ = xFlux.FLUTPJDOSQ
    xYFLUTPJ1.FLUTPJCLI = ""
    xYFLUTPJ1.FLUTPJTXT = ""
End If
End Sub
Public Sub fraFlux_Display_FLUTPJTXT()
Dim xSQL As String

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YFLUTPJ1 where FLUTPJDOS = " & xFlux.FLUTPJDOS & " and FLUTPJDOSQ = " & xFlux.FLUTPJDOSQ
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    V = rsYFLUTPJ1_GetBuffer(rsSab, xYFLUTPJ1_Detail)
Else
    xYFLUTPJ1_Detail.FLUTPJDOS = xFlux.FLUTPJDOS
    xYFLUTPJ1_Detail.FLUTPJDOSQ = xFlux.FLUTPJDOSQ
    xYFLUTPJ1_Detail.FLUTPJCLI = ""
    xYFLUTPJ1_Detail.FLUTPJTXT = ""
End If

End Sub
Public Sub fraFlux_Display_FLUTPJCLI()
Dim xSQL As String, xCLI As String, X As String

cboFlux_FLUTPJCLI.Clear
xSQL = "select BIATABK2 from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'FLUTPJCLI' and BIATABK1 = '" & Trim(oldFlux.FLUTPJOPE) & "'"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    xCLI = rsSab("BIATABK2")
    xSQL = "select CLIENARA1 from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & xCLI & "'"
    Set rsSabX = cnsab.Execute(xSQL)
    If Not rsSabX.EOF Then
        X = Trim(rsSabX("CLIENARA1"))
    Else
        X = "???"
    End If
    cboFlux_FLUTPJCLI.AddItem Trim(xCLI) & " : " & X
    rsSab.MoveNext
Loop

cbo_Scan xYFLUTPJ1.FLUTPJCLI, cboFlux_FLUTPJCLI
    
End Sub

