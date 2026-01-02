VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSAB_Client 
   AutoRedraw      =   -1  'True
   Caption         =   "SAb : base CLIENT"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13560
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SAB_Client.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9270
   ScaleWidth      =   13560
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
      Left            =   7800
      TabIndex        =   3
      Top             =   0
      Width           =   5175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   -30
      TabIndex        =   2
      Top             =   435
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Sélection des clients"
      TabPicture(0)   =   "SAB_Client.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mise à jour des mandataires"
      TabPicture(1)   =   "SAB_Client.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraUpdate"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Piste d'audit"
      TabPicture(2)   =   "SAB_Client.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblUpdLog_CLIRGPCLI"
      Tab(2).Control(1)=   "lblUpdLog_CLIRGPREG"
      Tab(2).Control(2)=   "txtUpdLog_AmjMin"
      Tab(2).Control(3)=   "fgSelect"
      Tab(2).Control(4)=   "cmdUpdLog_Ok"
      Tab(2).Control(5)=   "txtUpdLog_CLIRGPCLI"
      Tab(2).Control(6)=   "txtUpdLog_CLIRGPREG"
      Tab(2).Control(7)=   "fraSelect_Options_No"
      Tab(2).Control(8)=   "chkUpdLog_AmjMin"
      Tab(2).Control(9)=   "optUpdLog_YKYCDOS0"
      Tab(2).Control(10)=   "optUpdLog_YUPDLOG0"
      Tab(2).Control(11)=   "fraSelect_Options_Xgsop"
      Tab(2).Control(12)=   "fraSelect_Options_KYCgsop"
      Tab(2).Control(13)=   "fraSelect_Options_4"
      Tab(2).Control(14)=   "fraSelect_Options_3"
      Tab(2).Control(15)=   "fraSelect_Options_KYCech"
      Tab(2).ControlCount=   16
      TabCaption(3)   =   "Paramétrage"
      TabPicture(3)   =   "SAB_Client.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ssTab_Param"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame fraSelect_Options_KYCech 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   -69360
         TabIndex        =   208
         Top             =   2745
         Visible         =   0   'False
         Width           =   6000
         Begin VB.ComboBox cboSelect_Options_KYCech_Doc 
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
            Left            =   1470
            Sorted          =   -1  'True
            TabIndex        =   215
            Top             =   750
            Width           =   4230
         End
         Begin VB.ComboBox cboSelect_Options_KYCech_CLIENARES 
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
            Left            =   1455
            Sorted          =   -1  'True
            TabIndex        =   209
            Text            =   "CLIENARES"
            Top             =   270
            Width           =   1680
         End
         Begin MSComCtl2.DTPicker txtSelect_Options_KYCech_KYCDOSDECH 
            Height          =   300
            Left            =   4500
            TabIndex        =   211
            Top             =   270
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   139526147
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.Label lblSelect_Options_KYCech_Doc 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Document"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   255
            TabIndex        =   214
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label lblSelect_Options_KYCech_KYCDOSDECH 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Echéance <="
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3345
            TabIndex        =   212
            Top             =   300
            Width           =   1020
         End
         Begin VB.Label lblSelect_Options_KYCech_CLIENARES 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Responsable"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   270
            TabIndex        =   210
            Top             =   300
            Width           =   1020
         End
      End
      Begin VB.Frame fraSelect_Options_3 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   -74115
         TabIndex        =   85
         Top             =   2025
         Visible         =   0   'False
         Width           =   6000
         Begin VB.OptionButton optSelect_Options_3C 
            BackColor       =   &H00E0FFFF&
            Caption         =   "crédits confirmés (91120)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   88
            Top             =   210
            Value           =   -1  'True
            Width           =   2235
         End
         Begin VB.OptionButton optSelect_Options_3N 
            BackColor       =   &H00E0FFFF&
            Caption         =   "crédits notifiés (98050)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   270
            TabIndex        =   87
            Top             =   555
            Width           =   2190
         End
         Begin VB.CheckBox chkSelect_Options_3 
            BackColor       =   &H00E0FFFF&
            Caption         =   "inclure les comptes d'engagement soldés"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2505
            TabIndex        =   86
            Top             =   570
            Width           =   3435
         End
         Begin VB.Label libSelect_Options_3 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Les comptes clos sont ignorés (12120)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2580
            TabIndex        =   89
            Top             =   240
            Width           =   3060
         End
      End
      Begin VB.Frame fraSelect_Options_4 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   -73965
         TabIndex        =   91
         Top             =   3555
         Visible         =   0   'False
         Width           =   6000
         Begin VB.ComboBox cboSelect_Options_4_ECHISBFIN 
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
            Left            =   3650
            TabIndex        =   104
            Top             =   180
            Width           =   1680
         End
         Begin VB.ComboBox cboSelect_Options_4_CLIENARES 
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
            Left            =   3650
            Sorted          =   -1  'True
            TabIndex        =   101
            Text            =   "CLIENARES"
            Top             =   630
            Width           =   1680
         End
         Begin VB.ComboBox cboSelect_Options_4_Code 
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
            Left            =   1300
            Sorted          =   -1  'True
            TabIndex        =   96
            Text            =   "IDE"
            Top             =   555
            Width           =   945
         End
         Begin VB.CheckBox chkSelect_Options_4_ECHTABDON_S 
            BackColor       =   &H00E0FFFF&
            Caption         =   "inclure les conditions annulées"
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
            Left            =   120
            TabIndex        =   95
            Top             =   1005
            Width           =   2565
         End
         Begin VB.TextBox txtSelect_Options_4_CLIENACLI 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1300
            TabIndex        =   93
            Top             =   210
            Width           =   930
         End
         Begin VB.CheckBox chkSelect_Options_4_AUTSICMON 
            BackColor       =   &H00E0FFFF&
            Caption         =   "inclure les autorisations de déc = 0"
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
            Left            =   2895
            TabIndex        =   92
            Top             =   1005
            Width           =   2865
         End
         Begin VB.Label lblSelect_Options_4_ECHISBFIN 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Date d'arrêté"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2500
            TabIndex        =   103
            Top             =   255
            Width           =   1035
         End
         Begin VB.Label lblSelect_Options_4_CLIENARES 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Responsable"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2500
            TabIndex        =   102
            Top             =   660
            Width           =   1020
         End
         Begin VB.Label lblSelect_Options_4_Code 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Code"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   97
            Top             =   570
            Width           =   870
         End
         Begin VB.Label lblSelect_Options_4_CLIENACLI 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Racine / 99999"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   94
            Top             =   270
            Width           =   1275
         End
      End
      Begin VB.Frame fraSelect_Options_KYCgsop 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   -68715
         TabIndex        =   197
         Top             =   5895
         Visible         =   0   'False
         Width           =   6000
         Begin VB.Frame farSelect_Options_KYCgsop_Detail_ 
            BackColor       =   &H00C0E0FF&
            Caption         =   "détail"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   3900
            TabIndex        =   203
            Top             =   75
            Width           =   1830
            Begin VB.OptionButton optSelect_Options_KYCgsop_Detail_Missing 
               BackColor       =   &H00C0E0FF&
               Caption         =   "doc manquants"
               Height          =   210
               Left            =   60
               TabIndex        =   206
               Top             =   780
               Width           =   1600
            End
            Begin VB.OptionButton optSelect_Options_KYCgsop_Detail_OK 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Oui"
               Height          =   210
               Left            =   60
               TabIndex        =   205
               Top             =   525
               Width           =   1600
            End
            Begin VB.OptionButton optSelect_Options_KYCgsop_Detail_NOK 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Non"
               Height          =   210
               Left            =   60
               TabIndex        =   204
               Top             =   240
               Value           =   -1  'True
               Width           =   1600
            End
         End
         Begin VB.OptionButton optSelect_Options_KYCgsop_All 
            BackColor       =   &H00E0FFFF&
            Caption         =   "tous les dossiers"
            Height          =   210
            Left            =   1500
            TabIndex        =   202
            Top             =   825
            Width           =   2010
         End
         Begin VB.OptionButton optSelect_Options_KYCgsop_OK 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Dossiers complets"
            Height          =   210
            Left            =   1500
            TabIndex        =   201
            Top             =   525
            Width           =   2010
         End
         Begin VB.OptionButton optSelect_Options_KYCgsop_NOK 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Dossiers incomplets"
            Height          =   210
            Left            =   1500
            TabIndex        =   200
            Top             =   210
            Value           =   -1  'True
            Width           =   2010
         End
         Begin VB.ComboBox cboSelect_Options_KYCgsop_CLIENARES 
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
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   198
            Text            =   "CLIENARES"
            Top             =   600
            Width           =   1080
         End
         Begin VB.Label lblSelect_Options_KYCgsop_CLIENARES 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Responsable"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   199
            Top             =   210
            Width           =   1020
         End
      End
      Begin VB.Frame fraSelect_Options_Xgsop 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   -74085
         TabIndex        =   105
         Top             =   6600
         Visible         =   0   'False
         Width           =   6000
         Begin VB.ComboBox cboSelect_Options_Xgsop_Archive 
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
            Left            =   3690
            TabIndex        =   213
            Top             =   420
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.ComboBox cboSelect_Options_Xgsop_CLIENARES 
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
            Left            =   1695
            Sorted          =   -1  'True
            TabIndex        =   106
            Text            =   "CLIENARES"
            Top             =   465
            Width           =   1680
         End
         Begin VB.Label lblSelect_Options_Xgsop_CLIENARES 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Responsable"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   420
            TabIndex        =   107
            Top             =   480
            Width           =   1020
         End
      End
      Begin VB.OptionButton optUpdLog_YUPDLOG0 
         BackColor       =   &H0080C0FF&
         Caption         =   "Mandataires"
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
         Left            =   -74835
         TabIndex        =   183
         Top             =   750
         Width           =   1500
      End
      Begin VB.OptionButton optUpdLog_YKYCDOS0 
         BackColor       =   &H0080C0FF&
         Caption         =   "KYC"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74835
         TabIndex        =   182
         Top             =   390
         Value           =   -1  'True
         Width           =   1515
      End
      Begin VB.CheckBox chkUpdLog_AmjMin 
         BackColor       =   &H00C0E0FF&
         Caption         =   "date màj"
         Height          =   330
         Left            =   -70905
         TabIndex        =   181
         Top             =   400
         Width           =   1200
      End
      Begin TabDlg.SSTab ssTab_Param 
         Height          =   8145
         Left            =   -74910
         TabIndex        =   108
         Top             =   330
         Width           =   13350
         _ExtentX        =   23548
         _ExtentY        =   14367
         _Version        =   393216
         Tabs            =   5
         Tab             =   3
         TabsPerRow      =   5
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
         TabCaption(0)   =   "GSOP reporting"
         TabPicture(0)   =   "SAB_Client.frx":04B2
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fraParam"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Documents justificatifs "
         TabPicture(1)   =   "SAB_Client.frx":04CE
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraParam_YKYCDOS0"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "YKYCDOS0"
         TabPicture(2)   =   "SAB_Client.frx":04EA
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraYKYCDOS0"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "PJ"
         TabPicture(3)   =   "SAB_Client.frx":0506
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "fraPJ"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Tab 4"
         TabPicture(4)   =   "SAB_Client.frx":0522
         Tab(4).ControlEnabled=   0   'False
         Tab(4).ControlCount=   0
         Begin VB.Frame fraPJ 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ajouter une Pièce Jointe"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6100
            Left            =   1875
            TabIndex        =   167
            Top             =   825
            Visible         =   0   'False
            Width           =   11280
            Begin VB.CommandButton cmdPJ_réseau 
               BackColor       =   &H0080C0FF&
               Caption         =   "accèder à \\DOCSRV2013\_SCAN"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   255
               Style           =   1  'Graphical
               TabIndex        =   217
               Top             =   735
               Width           =   4005
            End
            Begin VB.CommandButton cmdPJ_Quit 
               BackColor       =   &H00808080&
               Caption         =   "Abandonner"
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
               Left            =   150
               Style           =   1  'Graphical
               TabIndex        =   176
               Top             =   5400
               Width           =   1035
            End
            Begin VB.CommandButton cmdPJ_OK 
               BackColor       =   &H0000FF00&
               Caption         =   "Enregistrer la pièce jointe"
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
               Left            =   1545
               Style           =   1  'Graphical
               TabIndex        =   175
               Top             =   5400
               Visible         =   0   'False
               Width           =   2800
            End
            Begin VB.FileListBox filDoc 
               ForeColor       =   &H00008000&
               Height          =   2820
               Left            =   4455
               Pattern         =   "*.doc;*.pdf;*.rtf;*.xls;*.txt"
               TabIndex        =   172
               Top             =   360
               Width           =   6600
            End
            Begin VB.DriveListBox DriveListBox 
               Height          =   330
               Left            =   240
               TabIndex        =   171
               Top             =   345
               Width           =   4000
            End
            Begin VB.DirListBox dirListBox 
               Height          =   2970
               Left            =   255
               TabIndex        =   170
               Top             =   1215
               Width           =   4000
            End
            Begin VB.CommandButton cmdPJ_Path 
               BackColor       =   &H0080C0FF&
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
               TabIndex        =   169
               Top             =   4290
               Width           =   2800
            End
            Begin RichTextLib.RichTextBox rtfPJ 
               Height          =   2460
               Left            =   4500
               TabIndex        =   168
               TabStop         =   0   'False
               Top             =   3390
               Width           =   6510
               _ExtentX        =   11483
               _ExtentY        =   4339
               _Version        =   393217
               BackColor       =   12648447
               Enabled         =   -1  'True
               HideSelection   =   0   'False
               ScrollBars      =   3
               AutoVerbMenu    =   -1  'True
               TextRTF         =   $"SAB_Client.frx":053E
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
            Begin VB.Label librtfPJ 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Caption         =   "Click droit pour copier/coller ==>"
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
               Left            =   900
               TabIndex        =   173
               Top             =   4920
               Width           =   3600
            End
         End
         Begin VB.Frame fraYKYCDOS0 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8175
            Left            =   -74535
            TabIndex        =   137
            Top             =   960
            Visible         =   0   'False
            Width           =   13335
            Begin MSFlexGridLib.MSFlexGrid fgYKYCDOS0 
               Height          =   6105
               Left            =   -8100
               TabIndex        =   153
               Top             =   4650
               Visible         =   0   'False
               Width           =   13200
               _ExtentX        =   23283
               _ExtentY        =   10769
               _Version        =   393216
               Cols            =   7
               FixedCols       =   0
               RowHeightMin    =   275
               BackColor       =   16777215
               ForeColor       =   0
               BackColorFixed  =   8421376
               ForeColorFixed  =   16777215
               BackColorBkg    =   -2147483633
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               FormatString    =   $"SAB_Client.frx":05BA
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
            Begin VB.Frame fraYKYCDOS0_Update 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Code"
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
               Height          =   8000
               Left            =   2055
               TabIndex        =   154
               Top             =   -180
               Visible         =   0   'False
               Width           =   9450
               Begin VB.TextBox txtYKYCDOS0_ZCLIENA0 
                  BackColor       =   &H0000FFFF&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   315
                  Left            =   255
                  MultiLine       =   -1  'True
                  TabIndex        =   216
                  Text            =   "SAB_Client.frx":0721
                  Top             =   1935
                  Visible         =   0   'False
                  Width           =   8715
               End
               Begin VB.CommandButton cmdYKYCDOS0_Ignore 
                  BackColor       =   &H00FFFF00&
                  Caption         =   "Ignorer ce document"
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
                  TabIndex        =   188
                  Top             =   4560
                  Visible         =   0   'False
                  Width           =   1035
               End
               Begin VB.CommandButton cmdYKYCDOS0_Missing 
                  BackColor       =   &H000000FF&
                  Caption         =   "Document manquant"
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
                  Left            =   8295
                  MaskColor       =   &H00FFFFFF&
                  Style           =   1  'Graphical
                  TabIndex        =   187
                  Top             =   5550
                  Visible         =   0   'False
                  Width           =   1035
               End
               Begin VB.CommandButton cmdPJ_Delete 
                  BackColor       =   &H00FF80FF&
                  Caption         =   "Supprimer la pièce jointe"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   540
                  Left            =   2190
                  Style           =   1  'Graphical
                  TabIndex        =   178
                  Top             =   7020
                  Visible         =   0   'False
                  Width           =   4700
               End
               Begin VB.CommandButton cmdYKYCDOS0_PJ 
                  BackColor       =   &H0000FFFF&
                  Caption         =   "Ajouter une pièce Jointe"
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
                  Left            =   8400
                  Style           =   1  'Graphical
                  TabIndex        =   174
                  Top             =   3495
                  Visible         =   0   'False
                  Width           =   1100
               End
               Begin VB.CommandButton cmdYKYCDOS0_Add 
                  BackColor       =   &H0000FF00&
                  Caption         =   "Ajouter"
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
                  Left            =   7800
                  Style           =   1  'Graphical
                  TabIndex        =   165
                  Top             =   3500
                  Visible         =   0   'False
                  Width           =   1035
               End
               Begin VB.CommandButton cmdYKYCDOS0_Update 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Modifier"
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
                  Left            =   7800
                  Style           =   1  'Graphical
                  TabIndex        =   164
                  Top             =   4500
                  Visible         =   0   'False
                  Width           =   1035
               End
               Begin VB.CommandButton cmdYKYCDOS0_Delete 
                  BackColor       =   &H00FF80FF&
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
                  Left            =   7800
                  Style           =   1  'Graphical
                  TabIndex        =   163
                  Top             =   5500
                  Visible         =   0   'False
                  Width           =   1035
               End
               Begin VB.CommandButton cmdYKYCDOS0_Quit 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Abandonner"
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
                  Left            =   7800
                  Style           =   1  'Graphical
                  TabIndex        =   162
                  Top             =   6500
                  Width           =   1035
               End
               Begin VB.TextBox txtYKYCDOS0_KYCDOSDLIB 
                  Height          =   840
                  Left            =   2130
                  MultiLine       =   -1  'True
                  TabIndex        =   159
                  Top             =   2370
                  Width           =   7020
               End
               Begin MSComCtl2.DTPicker txtYKYCDOS0_KYCDOSDAMJ 
                  Height          =   300
                  Left            =   2265
                  TabIndex        =   160
                  Top             =   1545
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   16777215
                  CalendarForeColor=   0
                  CalendarTitleBackColor=   8421504
                  CalendarTitleForeColor=   16777215
                  CalendarTrailingForeColor=   12632256
                  CustomFormat    =   "dd  MM yyy"
                  Format          =   139526145
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   2.00001157407407
               End
               Begin MSComCtl2.DTPicker txtYKYCDOS0_KYCDOSDECH 
                  Height          =   300
                  Left            =   5295
                  TabIndex        =   161
                  Top             =   1530
                  Width           =   1365
                  _ExtentX        =   2408
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   16777215
                  CalendarForeColor=   0
                  CalendarTitleBackColor=   8421504
                  CalendarTitleForeColor=   16777215
                  CalendarTrailingForeColor=   12632256
                  CustomFormat    =   "dd  MM yyy"
                  Format          =   139526145
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   2.00001157407407
               End
               Begin MSFlexGridLib.MSFlexGrid fgPJ 
                  Height          =   3420
                  Left            =   2190
                  TabIndex        =   177
                  Top             =   3510
                  Width           =   4980
                  _ExtentX        =   8784
                  _ExtentY        =   6033
                  _Version        =   393216
                  Cols            =   1
                  FixedCols       =   0
                  RowHeightMin    =   350
                  BackColor       =   16777215
                  ForeColor       =   16384
                  BackColorFixed  =   12640511
                  ForeColorFixed  =   0
                  BackColorBkg    =   15794175
                  WordWrap        =   -1  'True
                  AllowUserResizing=   3
                  FormatString    =   "<Pièces jointes                                                                                                                  "
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
               Begin VB.Label libYKYCDOS0_Comment 
                  BackColor       =   &H00D0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Commentaire"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   420
                  Left            =   195
                  TabIndex        =   196
                  Top             =   1005
                  Width           =   8925
               End
               Begin VB.Label lblYKYCDOS0_KYCDOSUUSR 
                  BackColor       =   &H00C0E0FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "UMAJ"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   270
                  Left            =   135
                  TabIndex        =   166
                  Top             =   7650
                  Width           =   9150
               End
               Begin VB.Label lblYKYCDOS0_KYCDOSDLIB 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Commentaire"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   240
                  TabIndex        =   158
                  Top             =   2595
                  Width           =   1545
               End
               Begin VB.Label lblYKYCDOS0_KYCDOSDECH 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Echéance"
                  ForeColor       =   &H00000000&
                  Height          =   225
                  Left            =   4215
                  TabIndex        =   157
                  Top             =   1545
                  Width           =   1005
               End
               Begin VB.Label lblYKYCDOS0_KYCDOSDAMJ 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Date du document"
                  ForeColor       =   &H00000000&
                  Height          =   390
                  Left            =   225
                  TabIndex        =   156
                  Top             =   1590
                  Width           =   1740
               End
               Begin VB.Label libYKYCDOS0_Document 
                  BackColor       =   &H00D0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Document"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   465
                  Left            =   165
                  TabIndex        =   155
                  Top             =   345
                  Width           =   9000
               End
            End
            Begin VB.ListBox lstYKYCDOS0_CLIENACAT 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5910
               Left            =   60
               TabIndex        =   151
               Top             =   2000
               Visible         =   0   'False
               Width           =   8190
            End
            Begin VB.Frame fraYKYCDOS_ZCLIENA0 
               BackColor       =   &H00F0FFFF&
               Caption         =   "type de clientèle"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF00FF&
               Height          =   1665
               Left            =   195
               TabIndex        =   139
               Top             =   90
               Width           =   8205
               Begin VB.CommandButton cmdYKYCDOS0_Delete_All 
                  BackColor       =   &H00FFC0FF&
                  Caption         =   "Effacer le dossier"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Left            =   6225
                  Style           =   1  'Graphical
                  TabIndex        =   184
                  Top             =   1290
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin VB.Label lblYKYCDOS0_CLIENARSD 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Résidence :"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   180
                  Left            =   150
                  TabIndex        =   149
                  Top             =   1350
                  Width           =   900
               End
               Begin VB.Label lblYKYCDOS0_CLIENANAT 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Nationalité :"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   180
                  Left            =   150
                  TabIndex        =   148
                  Top             =   1050
                  Width           =   900
               End
               Begin VB.Label libYKYCDOS0_CLIENANAT 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "CLIENANAT"
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
                  Left            =   1290
                  TabIndex        =   147
                  Top             =   1050
                  Width           =   2715
               End
               Begin VB.Label libYKYCDOS0_CLIENARSD 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "CLIENARSD"
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
                  Left            =   1305
                  TabIndex        =   146
                  Top             =   1350
                  Width           =   2610
               End
               Begin VB.Label libYKYCDOS0_CLIENAETA 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "CLIENAETA"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   1300
                  TabIndex        =   145
                  Top             =   400
                  Width           =   2910
               End
               Begin VB.Label libYKYCDOS0_CLIENACAT 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "CLIENACAT"
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
                  Left            =   4320
                  TabIndex        =   144
                  Top             =   390
                  Width           =   3630
               End
               Begin VB.Label libYKYCDOS0_CLIENARES 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "CLIENARES"
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
                  Left            =   4320
                  TabIndex        =   143
                  Top             =   1050
                  Width           =   3660
               End
               Begin VB.Label libYKYCDOS0_CLIENACOL 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Collectif"
                  Height          =   285
                  Left            =   180
                  TabIndex        =   142
                  Top             =   400
                  Width           =   975
               End
               Begin VB.Label libYKYCDOS0_CLIENARA1 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Intitulé"
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
                  Left            =   1300
                  TabIndex        =   141
                  Top             =   750
                  Width           =   6500
               End
               Begin VB.Label libYKYCDOS0_CLIENACLI 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "racine"
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
                  Left            =   150
                  TabIndex        =   140
                  Top             =   750
                  Width           =   975
               End
            End
            Begin MSFlexGridLib.MSFlexGrid fgYKYCDOS0_ZADRESS0 
               Height          =   8000
               Left            =   8445
               TabIndex        =   150
               Top             =   200
               Visible         =   0   'False
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   14129
               _Version        =   393216
               Cols            =   4
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16777215
               ForeColor       =   16384
               BackColorFixed  =   12640511
               ForeColorFixed  =   0
               BackColorBkg    =   -2147483633
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               FormatString    =   $"SAB_Client.frx":0738
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
         End
         Begin VB.Frame fraParam_YKYCDOS0 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8325
            Left            =   -74880
            TabIndex        =   118
            Top             =   360
            Visible         =   0   'False
            Width           =   13200
            Begin VB.CommandButton cmdParam_YKYCDOS0_4c_Actualisation 
               BackColor       =   &H0000FFFF&
               Caption         =   $"SAB_Client.frx":07E2
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1815
               Left            =   6780
               Style           =   1  'Graphical
               TabIndex        =   207
               Top             =   165
               Visible         =   0   'False
               Width           =   3210
            End
            Begin MSFlexGridLib.MSFlexGrid fgParam_YKYCDOS0_4c 
               Height          =   5460
               Left            =   5490
               TabIndex        =   190
               Top             =   4260
               Visible         =   0   'False
               Width           =   13200
               _ExtentX        =   23283
               _ExtentY        =   9631
               _Version        =   393216
               Cols            =   7
               FixedCols       =   0
               RowHeightMin    =   275
               BackColor       =   16773375
               ForeColor       =   0
               BackColorFixed  =   8388736
               ForeColorFixed  =   16777215
               BackColorBkg    =   -2147483633
               WordWrap        =   -1  'True
               GridLines       =   2
               AllowUserResizing=   3
               FormatString    =   $"SAB_Client.frx":086C
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
            Begin VB.Frame fraParam_YKYCDOS0_4c 
               BackColor       =   &H00FFF0FF&
               Caption         =   "fraParam_YKYCDOS0_4c"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   1995
               Left            =   6585
               TabIndex        =   189
               Top             =   2400
               Visible         =   0   'False
               Width           =   9495
               Begin VB.TextBox txtParam_YKYCDOS0_4c 
                  Height          =   870
                  Left            =   390
                  MaxLength       =   128
                  MultiLine       =   -1  'True
                  TabIndex        =   195
                  Top             =   510
                  Width           =   8715
               End
               Begin VB.CheckBox chkParam_YKYCDOS0_4c 
                  BackColor       =   &H0000FF00&
                  Caption         =   "document obligatoire"
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
                  Left            =   390
                  TabIndex        =   193
                  Top             =   1560
                  Width           =   2625
               End
               Begin VB.CommandButton cmdParam_YKYCDOS0_4c_Update 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Modifier"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   500
                  Left            =   7620
                  Style           =   1  'Graphical
                  TabIndex        =   192
                  Top             =   1440
                  Width           =   1500
               End
               Begin VB.CommandButton cmdParam_YKYCDOS0_4c_Quit 
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
                  Height          =   500
                  Left            =   5520
                  Style           =   1  'Graphical
                  TabIndex        =   191
                  Top             =   1455
                  Width           =   1500
               End
               Begin VB.Label lblParam_YKYCDOS0_4c 
                  BackColor       =   &H00FFF0FF&
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
                  Height          =   390
                  Left            =   405
                  TabIndex        =   194
                  Top             =   270
                  Width           =   1080
               End
            End
            Begin VB.Frame fraParam_YKYCDOS0_JD 
               BackColor       =   &H00808000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7605
               Left            =   3855
               TabIndex        =   131
               Top             =   4155
               Visible         =   0   'False
               Width           =   12015
               Begin VB.CommandButton cmdParam_YKYCDOS0_JD_Quit 
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
                  Height          =   500
                  Left            =   7800
                  Style           =   1  'Graphical
                  TabIndex        =   136
                  Top             =   7035
                  Width           =   1500
               End
               Begin VB.CommandButton cmdParam_YKYCDOS0_JD_Update 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Enregistrer les modifications"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   500
                  Left            =   10140
                  Style           =   1  'Graphical
                  TabIndex        =   135
                  Top             =   7050
                  Width           =   1500
               End
               Begin VB.ListBox lstParam_YKYCDOS0_JD 
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   5685
                  Left            =   7380
                  Sorted          =   -1  'True
                  Style           =   1  'Checkbox
                  TabIndex        =   134
                  Top             =   900
                  Width           =   4455
               End
               Begin VB.ListBox lstParam_YKYCDOS0_D 
                  BackColor       =   &H00E0FFE0&
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   4545
                  ItemData        =   "SAB_Client.frx":09B2
                  Left            =   270
                  List            =   "SAB_Client.frx":09B9
                  TabIndex        =   133
                  Top             =   2685
                  Width           =   6855
               End
               Begin VB.ListBox lstParam_YKYCDOS0_J 
                  BackColor       =   &H0080FF80&
                  Height          =   1530
                  ItemData        =   "SAB_Client.frx":09D2
                  Left            =   250
                  List            =   "SAB_Client.frx":09D9
                  TabIndex        =   132
                  Top             =   945
                  Width           =   6855
               End
               Begin VB.Label libParam_YKYCDOS0_D 
                  BackColor       =   &H00E0FFE0&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "D"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   465
                  Left            =   3180
                  TabIndex        =   180
                  Top             =   180
                  Width           =   8640
                  WordWrap        =   -1  'True
               End
               Begin VB.Label libParam_YKYCDOS0_J 
                  BackColor       =   &H0080FF80&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "J"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   465
                  Left            =   250
                  TabIndex        =   179
                  Top             =   180
                  Width           =   2790
                  WordWrap        =   -1  'True
               End
            End
            Begin VB.ListBox lstParam_KYCDOSNAT 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1860
               Left            =   150
               TabIndex        =   130
               Top             =   150
               Width           =   3375
            End
            Begin VB.ListBox lstParam_YKYCDOS0 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5460
               Left            =   165
               TabIndex        =   123
               Top             =   2145
               Width           =   12915
            End
            Begin VB.Frame fraParam_YKYCDOS0_Update 
               BackColor       =   &H00E0FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1920
               Left            =   3840
               TabIndex        =   119
               Top             =   810
               Visible         =   0   'False
               Width           =   9495
               Begin VB.TextBox txtParam_KYCDOSDECH 
                  Height          =   270
                  Left            =   8580
                  MaxLength       =   2
                  TabIndex        =   186
                  Top             =   195
                  Width           =   465
               End
               Begin VB.CommandButton cmdParam_YKYCDOS0_Update 
                  BackColor       =   &H0080FF80&
                  Caption         =   "Modifier"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   500
                  Left            =   4650
                  Style           =   1  'Graphical
                  TabIndex        =   129
                  Top             =   1300
                  Width           =   1500
               End
               Begin VB.CheckBox chkParam_KYCDOSSTAK 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "obligatoire"
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
                  Left            =   2850
                  TabIndex        =   128
                  Top             =   225
                  Width           =   2625
               End
               Begin VB.TextBox txtParam_KYCDOSDLIB 
                  Height          =   585
                  Left            =   1050
                  MaxLength       =   128
                  MultiLine       =   -1  'True
                  TabIndex        =   127
                  Top             =   645
                  Width           =   8115
               End
               Begin VB.TextBox txtParam_KYCDOSSEQ 
                  Height          =   270
                  Left            =   1080
                  MaxLength       =   7
                  TabIndex        =   125
                  Top             =   210
                  Width           =   1425
               End
               Begin VB.CommandButton cmdParam_YKYCDOS0_Quit 
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
                  Height          =   500
                  Left            =   120
                  Style           =   1  'Graphical
                  TabIndex        =   122
                  Top             =   1300
                  Width           =   1500
               End
               Begin VB.CommandButton cmdParam_YKYCDOS0_Add 
                  BackColor       =   &H000080FF&
                  Caption         =   "Ajouter"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   500
                  Left            =   7335
                  Style           =   1  'Graphical
                  TabIndex        =   121
                  Top             =   1300
                  Width           =   1500
               End
               Begin VB.CommandButton cmdParam_YKYCDOS0_Delete 
                  BackColor       =   &H00FF80FF&
                  Caption         =   "Supprimer"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   500
                  Left            =   2190
                  Style           =   1  'Graphical
                  TabIndex        =   120
                  Top             =   1300
                  Width           =   1500
               End
               Begin VB.Label lblParam_KYCDOSDECH 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Durée de validité du document "
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
                  Left            =   5730
                  TabIndex        =   185
                  Top             =   225
                  Width           =   2670
               End
               Begin VB.Label lblParam_KYCDOSDLIB 
                  BackColor       =   &H00E0FFFF&
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
                  Left            =   180
                  TabIndex        =   126
                  Top             =   735
                  Width           =   765
               End
               Begin VB.Label lblParam_KYCDOSSEQ 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Identifiant"
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
                  Left            =   135
                  TabIndex        =   124
                  Top             =   240
                  Width           =   915
               End
            End
         End
         Begin VB.Frame fraParam 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8325
            Left            =   -74940
            TabIndex        =   109
            Top             =   350
            Visible         =   0   'False
            Width           =   13200
            Begin VB.Frame fraParam_Update 
               BackColor       =   &H00E0FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7215
               Left            =   3600
               TabIndex        =   111
               Top             =   345
               Width           =   9495
               Begin VB.TextBox txtParam_Id 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Left            =   4170
                  MaxLength       =   10
                  TabIndex        =   115
                  Top             =   6390
                  Width           =   2040
               End
               Begin VB.CommandButton cmdParam_Delete 
                  BackColor       =   &H00FF80FF&
                  Caption         =   "Supprimer"
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
                  Left            =   2085
                  Style           =   1  'Graphical
                  TabIndex        =   114
                  Top             =   6195
                  Width           =   1500
               End
               Begin VB.CommandButton cmdParam_Add 
                  BackColor       =   &H000080FF&
                  Caption         =   "Ajouter"
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
                  Left            =   7200
                  Style           =   1  'Graphical
                  TabIndex        =   113
                  Top             =   6210
                  Width           =   1500
               End
               Begin VB.CommandButton cmdParam_Quit 
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
                  Left            =   255
                  Style           =   1  'Graphical
                  TabIndex        =   112
                  Top             =   6180
                  Width           =   1500
               End
               Begin MSFlexGridLib.MSFlexGrid fgParam 
                  Height          =   5715
                  Left            =   165
                  TabIndex        =   116
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   9045
                  _ExtentX        =   15954
                  _ExtentY        =   10081
                  _Version        =   393216
                  Cols            =   4
                  FixedCols       =   0
                  RowHeightMin    =   350
                  BackColor       =   15794175
                  ForeColor       =   8192
                  BackColorFixed  =   8421376
                  ForeColorFixed  =   16777215
                  BackColorBkg    =   15794175
                  AllowUserResizing=   3
                  FormatString    =   $"SAB_Client.frx":09F2
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
               Begin VB.Label lblParam_Id 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Identifiant"
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
                  TabIndex        =   117
                  Top             =   6120
                  Width           =   1620
               End
            End
            Begin VB.ListBox lstParam 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3885
               Left            =   225
               TabIndex        =   110
               Top             =   1245
               Width           =   3105
            End
         End
      End
      Begin VB.Frame fraSelect_Options_No 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   -73965
         TabIndex        =   100
         Top             =   5130
         Width           =   6000
      End
      Begin VB.TextBox txtUpdLog_CLIRGPREG 
         Height          =   315
         Left            =   -69465
         TabIndex        =   79
         Top             =   750
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtUpdLog_CLIRGPCLI 
         Height          =   315
         Left            =   -72780
         TabIndex        =   78
         Top             =   750
         Width           =   1350
      End
      Begin VB.CommandButton cmdUpdLog_Ok 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Exécuter la requête"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -63600
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   480
         Width           =   1575
      End
      Begin VB.Frame fraUpdate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8175
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   13335
         Begin VB.Frame fraUpdate_Add 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Ajouter une entité"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1215
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   6015
            Begin VB.TextBox txtUpdate_Select 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4200
               TabIndex        =   34
               Top             =   240
               Width           =   1695
            End
            Begin VB.OptionButton optUpdate_Add_Old 
               BackColor       =   &H00E0FFFF&
               Caption         =   "existant"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2520
               TabIndex        =   33
               Top             =   240
               Width           =   1695
            End
            Begin VB.OptionButton optUpdate_Add_PP 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Personne Physique"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2520
               TabIndex        =   27
               Top             =   600
               Value           =   -1  'True
               Width           =   1695
            End
            Begin VB.OptionButton optUpdate_Add_Sté 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Société"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2520
               TabIndex        =   26
               Top             =   960
               Width           =   1695
            End
            Begin VB.ComboBox cboUpdate_Add 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   360
               Width           =   2055
            End
            Begin VB.CommandButton cmdUpdate_Ok 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Mise à jour"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   525
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   600
               Width           =   1695
            End
         End
         Begin VB.Frame fraUpdate_Détail 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7815
            Left            =   6360
            TabIndex        =   21
            Top             =   240
            Width           =   6615
            Begin VB.CheckBox chkUpdate_ADRESSAD1 
               Caption         =   "Néant"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   84
               Top             =   2280
               Width           =   975
            End
            Begin VB.Frame fraUpdate_Détail_Sté 
               Caption         =   "Société"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2055
               Left            =   135
               TabIndex        =   69
               Top             =   5325
               Width           =   6375
               Begin VB.ComboBox cboUpdate_CLIENBTER 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   71
                  Top             =   1560
                  Width           =   495
               End
               Begin VB.TextBox txtUpdate_CLIENASRN 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   73
                  Text            =   "NEANT"
                  Top             =   360
                  Width           =   2415
               End
               Begin VB.TextBox txtUpdate_CLIENAREG 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   72
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.Label lblUpdate_CLIENBTER 
                  Caption         =   "Territorialité"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   76
                  Top             =   1560
                  Width           =   1215
               End
               Begin VB.Label Label14 
                  Caption         =   "Siren"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   75
                  Top             =   360
                  Width           =   735
               End
               Begin VB.Label Label15 
                  Caption         =   "code APE"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   74
                  Top             =   960
                  Width           =   1455
               End
            End
            Begin VB.TextBox txtUpdate_CLIENARA2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1080
               TabIndex        =   54
               Top             =   1440
               Width           =   5415
            End
            Begin VB.Frame fraUpdate_Détail_PP 
               Caption         =   "Naissance"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2655
               Left            =   120
               TabIndex        =   36
               Top             =   5040
               Width           =   6375
               Begin VB.ComboBox cboUpdate_CLIENBLIE 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   5760
                  Style           =   2  'Dropdown List
                  TabIndex        =   67
                  Top             =   1680
                  Width           =   495
               End
               Begin VB.TextBox txtUpdate_CLIENBCIN 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   63
                  Top             =   1200
                  Width           =   855
               End
               Begin VB.TextBox txtUpdate_CLIENBCOM 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   62
                  Top             =   720
                  Width           =   4335
               End
               Begin VB.TextBox txtUpdate_CLIENAFIL 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   66
                  Top             =   2160
                  Width           =   4335
               End
               Begin VB.ComboBox cboUpdate_CLIENBNAS 
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   1920
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   65
                  Top             =   1680
                  Width           =   2415
               End
               Begin VB.TextBox txtUpdate_CLIENBINS 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   3240
                  TabIndex        =   64
                  Top             =   1200
                  Width           =   3015
               End
               Begin MSComCtl2.DTPicker txtUpdate_CLIENADNA 
                  Height          =   300
                  Left            =   1920
                  TabIndex        =   61
                  Top             =   240
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   529
                  _Version        =   393216
                  CalendarBackColor=   16777215
                  CalendarForeColor=   0
                  CalendarTitleBackColor=   8421504
                  CalendarTitleForeColor=   16777215
                  CalendarTrailingForeColor=   12632256
                  CustomFormat    =   "dd  MM yyy"
                  Format          =   139067395
                  CurrentDate     =   36299
                  MaxDate         =   401768
                  MinDate         =   -328351
               End
               Begin VB.Label Label13 
                  Caption         =   "lieu FICOBA"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   4440
                  TabIndex        =   50
                  Top             =   1800
                  Width           =   1215
               End
               Begin VB.Label Label12 
                  Caption         =   "Nom de jeune fille"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   49
                  Top             =   2280
                  Width           =   1455
               End
               Begin VB.Label Label11 
                  Caption         =   "Pays"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   48
                  Top             =   1800
                  Width           =   1575
               End
               Begin VB.Label Label10 
                  Caption         =   "Commune"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   47
                  Top             =   840
                  Width           =   1095
               End
               Begin VB.Label Label9 
                  Caption         =   "Département       code (insee,libellé)"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   240
                  TabIndex        =   46
                  Top             =   1200
                  Width           =   1575
               End
               Begin VB.Label Label1 
                  Caption         =   "date"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   45
                  Top             =   360
                  Width           =   975
               End
            End
            Begin VB.ComboBox cboUpdate_CLIENANAT 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2040
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   4560
               Width           =   4335
            End
            Begin VB.ComboBox cboUpdate_CLIENARSD 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2040
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   70
               Top             =   4080
               Width           =   4335
            End
            Begin VB.ComboBox cboUpdate_ADRESSPAY 
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2040
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Top             =   3480
               Width           =   4335
            End
            Begin VB.TextBox txtUpdate_CLIENASIG 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3720
               TabIndex        =   52
               Top             =   480
               Width           =   2295
            End
            Begin VB.ComboBox cboUpdate_CLIENAETA 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   360
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   51
               Top             =   480
               Width           =   2175
            End
            Begin VB.TextBox txtUpdate_ADRESSCOP 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1080
               TabIndex        =   58
               Top             =   3000
               Width           =   855
            End
            Begin VB.TextBox txtUpdate_ADRESSVIL 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2040
               TabIndex        =   59
               Top             =   3000
               Width           =   4455
            End
            Begin VB.TextBox txtUpdate_ADRESSAD3 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1080
               TabIndex        =   57
               Top             =   2640
               Width           =   5415
            End
            Begin VB.TextBox txtUpdate_ADRESSAD2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1080
               TabIndex        =   56
               Top             =   2280
               Width           =   5415
            End
            Begin VB.TextBox txtUpdate_ADRESSAD1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1080
               TabIndex        =   55
               Top             =   1920
               Width           =   5415
            End
            Begin VB.TextBox txtUpdate_CLIENARA1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1080
               TabIndex        =   53
               Top             =   960
               Width           =   5415
            End
            Begin VB.Label Label8 
               Caption         =   "pays nationalité"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   44
               Top             =   4560
               Width           =   1575
            End
            Begin VB.Label Label7 
               Caption         =   "pays résidence"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   43
               Top             =   4200
               Width           =   1575
            End
            Begin VB.Label Label6 
               Caption         =   "pays"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   42
               Top             =   3480
               Width           =   615
            End
            Begin VB.Label Label5 
               Caption         =   "Cp Ville"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   41
               Top             =   3000
               Width           =   735
            End
            Begin VB.Label Label4 
               Caption         =   "Adresse"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   40
               Top             =   1920
               Width           =   735
            End
            Begin VB.Label Label3 
               Caption         =   "Prénom"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   39
               Top             =   1440
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "Nom"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   38
               Top             =   960
               Width           =   735
            End
            Begin VB.Label qqq 
               Caption         =   "Sigle"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2760
               TabIndex        =   37
               Top             =   480
               Width           =   855
            End
         End
         Begin ComctlLib.TreeView tvwUpdate 
            Height          =   6495
            Left            =   120
            TabIndex        =   22
            Top             =   1560
            Width           =   6000
            _ExtentX        =   10583
            _ExtentY        =   11456
            _Version        =   327682
            Indentation     =   706
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   6
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame fraTab0 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8205
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13290
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
            Height          =   6855
            Left            =   165
            TabIndex        =   9
            Top             =   1185
            Width           =   13095
            Begin MSFlexGridLib.MSFlexGrid fgDetail 
               Height          =   6555
               Left            =   6000
               TabIndex        =   90
               Top             =   4545
               Visible         =   0   'False
               Width           =   6825
               _ExtentX        =   12039
               _ExtentY        =   11562
               _Version        =   393216
               Cols            =   7
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16777215
               ForeColor       =   16384
               BackColorFixed  =   8421376
               ForeColorFixed  =   -2147483633
               BackColorBkg    =   -2147483633
               AllowUserResizing=   3
               FormatString    =   $"SAB_Client.frx":0ABC
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
            Begin VB.Frame fraSelect_Update 
               Caption         =   "Mise à jour"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2655
               Left            =   11040
               TabIndex        =   28
               Top             =   105
               Width           =   1935
               Begin VB.OptionButton optSelect_YKYCDOS0 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Dossier client"
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
                  Left            =   150
                  TabIndex        =   138
                  Top             =   2310
                  Value           =   -1  'True
                  Width           =   1575
               End
               Begin VB.OptionButton optSelect_CLIENARES 
                  Caption         =   "LAB actif / inactif"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   105
                  TabIndex        =   83
                  Top             =   1440
                  Width           =   1575
               End
               Begin VB.CommandButton cmdSelect_Update 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Ok"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   525
                  Left            =   345
                  Style           =   1  'Graphical
                  TabIndex        =   32
                  Top             =   1710
                  Width           =   1095
               End
               Begin VB.OptionButton optSelect_Add 
                  Caption         =   "Ajouter un lien"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   31
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.OptionButton optSelect_Modification 
                  Caption         =   "Modifier les caractèristiques"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   105
                  TabIndex        =   30
                  Top             =   630
                  Width           =   1695
               End
               Begin VB.OptionButton optSelect_Suppress 
                  Caption         =   "Supprimer ce lien"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   29
                  Top             =   1080
                  Width           =   1575
               End
            End
            Begin ComctlLib.TreeView tvwSelect 
               Height          =   6495
               Left            =   90
               TabIndex        =   10
               Top             =   210
               Width           =   6000
               _ExtentX        =   10583
               _ExtentY        =   11456
               _Version        =   327682
               Indentation     =   706
               LabelEdit       =   1
               LineStyle       =   1
               Sorted          =   -1  'True
               Style           =   6
               BorderStyle     =   1
               Appearance      =   1
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
            Begin ComctlLib.TreeView tvwInverse 
               Height          =   3855
               Left            =   6240
               TabIndex        =   16
               Top             =   2880
               Width           =   6720
               _ExtentX        =   11853
               _ExtentY        =   6800
               _Version        =   327682
               Indentation     =   706
               LabelEdit       =   1
               LineStyle       =   1
               Sorted          =   -1  'True
               Style           =   6
               BorderStyle     =   1
               Appearance      =   1
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
            Begin VB.Label lblInverse 
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   6240
               TabIndex        =   19
               Top             =   1560
               Width           =   4695
            End
            Begin VB.Label lblSelect 
               BackColor       =   &H00E0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   6240
               TabIndex        =   18
               Top             =   240
               Width           =   4695
            End
         End
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00E0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1200
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   11625
            Begin VB.CheckBox chkSelect_Groupes 
               BackColor       =   &H00E0FFFF&
               Caption         =   "sélectionner uniquement les groupes "
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6135
               TabIndex        =   99
               Top             =   885
               Width           =   3630
            End
            Begin VB.CheckBox chkSelect_Racine 
               BackColor       =   &H00E0FFFF&
               Caption         =   "sélectionner toutes les racines actives"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6135
               TabIndex        =   98
               Top             =   690
               Width           =   3000
            End
            Begin VB.ComboBox cboSelect_CLIENACAT 
               Height          =   330
               Left            =   1665
               Sorted          =   -1  'True
               TabIndex        =   35
               Text            =   "CAT"
               Top             =   270
               Width           =   1995
            End
            Begin VB.TextBox txtSelect_PLANCOPRO 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1680
               TabIndex        =   15
               Text            =   "CAV"
               Top             =   780
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.ComboBox cboSelect_SQL 
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
               Left            =   6105
               Sorted          =   -1  'True
               TabIndex        =   13
               Text            =   "cboSelect_SQL"
               Top             =   225
               Width           =   5250
            End
            Begin VB.TextBox txtSelect_CLIENARA1 
               Height          =   285
               Left            =   4440
               TabIndex        =   12
               Top             =   270
               Width           =   1455
            End
            Begin VB.Label lblSelect_PLANCOPRO 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Type de compte"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   330
               TabIndex        =   14
               Top             =   810
               Visible         =   0   'False
               Width           =   1320
            End
            Begin VB.Label lblSelect_CLIENARA1 
               BackColor       =   &H00E0FFFF&
               Caption         =   "nom"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3735
               TabIndex        =   11
               Top             =   345
               Width           =   495
            End
            Begin VB.Label lblSelect_CLIENACAT 
               BackColor       =   &H00E0FFFF&
               Caption         =   "Catégorie client"
               Height          =   255
               Left            =   240
               TabIndex        =   8
               Top             =   375
               Width           =   1320
            End
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Exécuter la requête"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   11880
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   7425
         Left            =   -74790
         TabIndex        =   17
         Top             =   1200
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   13097
         _Version        =   393216
         Rows            =   1
         Cols            =   13
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
         FormatString    =   $"SAB_Client.frx":0B45
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
      Begin MSComCtl2.DTPicker txtUpdLog_AmjMin 
         Height          =   300
         Left            =   -70935
         TabIndex        =   82
         Top             =   750
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   12632256
         CustomFormat    =   "dd  MM yyy"
         Format          =   139132931
         CurrentDate     =   38699.44875
         MaxDate         =   401768
         MinDate         =   36526.4425347222
      End
      Begin VB.Label lblUpdLog_CLIRGPREG 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N° ADM-DIR-MAN"
         Height          =   255
         Left            =   -69465
         TabIndex        =   81
         Top             =   400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblUpdLog_CLIRGPCLI 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N° Client"
         Height          =   255
         Left            =   -72780
         TabIndex        =   80
         Top             =   405
         Width           =   1290
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
      Picture         =   "SAB_Client.frx":0BE5
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Label Label16 
      BackColor       =   &H00F0FFFF&
      Caption         =   "CLIENARES"
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
      Left            =   0
      TabIndex        =   152
      Top             =   0
      Width           =   3660
   End
   Begin VB.Label libSelect 
      BackColor       =   &H00FFFED9&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15
      TabIndex        =   4
      Top             =   0
      Width           =   4905
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnux1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
      Begin VB.Menu mnux2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Print_Liste 
         Caption         =   "Imprimer listes "
      End
   End
   Begin VB.Menu mnuPrint1 
      Caption         =   "mnuPrint1"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint1_Liste 
         Caption         =   "Imprimer liste"
      End
   End
   Begin VB.Menu mnuExcel 
      Caption         =   "mnuExcel"
      Visible         =   0   'False
      Begin VB.Menu mnuExcel_Exportation 
         Caption         =   "Exportation "
      End
   End
   Begin VB.Menu mnuPrint2 
      Caption         =   "mnuPrint2"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint2_Excel 
         Caption         =   "Excel"
      End
      Begin VB.Menu mnuPrint2_Mail 
         Caption         =   "envoi Mail"
      End
   End
End
Attribute VB_Name = "frmSAB_Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'tvwSelect arborescence
' Node.key  :
'   client      : CLI******
'               : CLI*******ADM         niveau lien
'               : CLI*******ADM-------  niveau client lié
'---------------------------------------------------------

Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim arrHab(19) As Boolean
'Dim SAB_CLIENT_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency
Dim blnAuto As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim rsSabX As New ADODB.Recordset

Dim cnAdo As New ADODB.Connection, rsAdo As New ADODB.Recordset, errADO As ADODB.Error
Dim rsAdo_ZCLIGRP0 As New ADODB.Recordset
Dim rsADO_YBIACPT0 As New ADODB.Recordset
Dim blnTransaction As Boolean

'______________________________________________________________________

Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long
Dim xZCLIENA0 As typeZCLIENA0, meZCLIENA0 As typeZCLIENA0
Dim selZCLIENA0 As typeZCLIENA0, oldZCLIENA0 As typeZCLIENA0, newZCLIENA0 As typeZCLIENA0
Dim arrZCLIENA0() As typeZCLIENA0, arrZCLIENA0_NB As Long, arrZCLIENA0_Max As Long, arrZCLIENA0_Index As Long
Dim xZCLIENB0 As typeZCLIENB0, oldZCLIENB0 As typeZCLIENB0, newZCLIENB0 As typeZCLIENB0

Dim arrZCLIGRP0() As typeZCLIGRP0, arrZCLIGRP0_Nb As Long, arrZCLIGRP0_Max As Long, arrZCLIGRP0_Index As Long
Dim selZCLIGRP0 As typeZCLIGRP0, newZCLIGRP0 As typeZCLIGRP0, oldZCLIGRP0 As typeZCLIGRP0, xZCLIGRP0 As typeZCLIGRP0
Dim mSelect_SQL As String
Dim xZADRESS0 As typeZADRESS0, newZADRESS0 As typeZADRESS0, oldZADRESS0 As typeZADRESS0

Dim meYBIACPT0 As typeYBIACPT0

Dim mSelect_Node_Key As String, mUpdate_Node_Key As String
Dim blnUpdate_Ok As Boolean
Dim meYUPDLOG0 As typeYUPDLOG0


Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim wsExcel_Row As Long
Dim wFile_Orig  As String, wFile As String
Dim cmdSelect_SQL_K As String
Dim xYBIACPT0 As typeYBIACPT0, oldYBIACPT0 As typeYBIACPT0

Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean

Dim mXls1_Row As Long, mXls1_Col As Long, mXls2_Row As Long, mXls2_Col As Long
Dim arrYBIACPT0() As typeYBIACPT0, arrYBIACPT0_Nb As Long
Dim wIBM_AmjMin As Long

Dim arrZAUTSYC0() As typeZAUTSYC0, arrZAUTSYC0_Nb As Long, xZAUTSYC0 As typeZAUTSYC0
Dim arrCLIENARA1() As String

Dim arrCLIRGPREG() As String, arrCLIRGPREG_Hierarchie() As String, arrCLIRGPREG_Niveau() As Integer, arrCLIRGPREG_Nb As Long, arrCLIRGPREG_K As Long

Dim arrZBAST12_Arg() As String, arrZBAST12_lib() As String, arrZBAST12_Nb As Long, arrZBAST12_K As Long

Dim blnSelect_SQL_XzR As Boolean
Dim lstPLANCOPRO As String, kPLANCOPRO As Integer

Dim arrCLIENARES(100) As String, arrCLIENARES_Nb As Integer, currentCLIENARES As String
Dim arrWECHISB0() As typeWECHISB0, arrWECHISB0_nb As Long, oldWECHISB0 As typeWECHISB0
Dim mXgsop As typeXgsop
Dim paramXgsop_PLANCOPRO As String, paramXgsop_NonRéclamés As String, paramXgsop_HorsGsop As String


Dim fgParam_FormatString As String, fgParam_K As Integer
Dim fgParam_RowDisplay As Integer, fgParam_RowClick As Integer, fgParam_ColClick As Integer
Dim fgParam_ColorClick As Long, fgParam_ColorDisplay As Long
Dim fgParam_Sort1 As Integer, fgParam_Sort2 As Integer
Dim fgParam_SortAD As Integer, fgParam_Sort1_Old As Integer
Dim fgParam_arrIndex As Integer
Dim blnfgParam_DisplayLine As Boolean
Dim lstParam_K As String, blnParam_KYCDOSNAT_4c As Boolean
Dim rsParam As New ADODB.Recordset

Dim Old_YBIATAB0 As typeYBIATAB0, New_YBIATAB0 As typeYBIATAB0
Dim oldYKYCDOS0 As typeYKYCDOS0, newYKYCDOS0 As typeYKYCDOS0, xYKYCDOS0 As typeYKYCDOS0
Dim mParam_KYCDOSNAT As String
Dim oldParam_J() As Long, newParam_J() As Long, mParam_J As Long
Dim oldParam_D() As Long, newParam_D() As Long, mParam_D As Long
Dim arrKYCDOSDLIB_J() As String, arrKYCDOSDLIB_D() As String
Dim arrKYCDOSSTAK_J() As String, arrKYCDOSSTAK_D() As String
Dim arrKYCDOSDECH_D() As Long
Dim blnYKYCDOS0_JD As Boolean, blnSelect_YKYCDOS0 As Boolean
Dim fgYKYCDOS0_ZADRESS0_FormatString As String
Dim currentYKYCDOS0 As typeYKYCDOS0
Dim arrYKYCDOS0_JD() As typeYKYCDOS0, arrYKYCDOS0_JD_Nb As Integer
Dim arrYKYCDOS0() As typeYKYCDOS0


Dim fgYKYCDOS0_FormatString As String, fgYKYCDOS0_K As Integer
Dim fgYKYCDOS0_RowDisplay As Integer, fgYKYCDOS0_RowClick As Integer, fgYKYCDOS0_ColClick As Integer
Dim fgYKYCDOS0_ColorClick As Long, fgYKYCDOS0_ColorDisplay As Long
Dim fgYKYCDOS0_Sort1 As Integer, fgYKYCDOS0_Sort2 As Integer
Dim fgYKYCDOS0_SortAD As Integer, fgYKYCDOS0_Sort1_Old As Integer
Dim fgYKYCDOS0_arrIndex As Integer
Dim blnfgYKYCDOS0_DisplayLine As Boolean

Dim mfilDoc_Path As String, blnfilDoc_Path As Boolean
Dim oldFileName As String, newFileName As String, newDirPath As String, newFileExtension As String
Dim mExe_Sequence As Long
Dim currentPJ_Path_FileName As String, currentPJ_FileName As String
Dim mKYCDOSDECH_Warn As Long
Dim mParam_YKYCDOS0_4c_Actualisation_Nb As Long

Dim oldYKYCSTA0 As typeYKYCSTA0, newYKYCSTA0 As typeYKYCSTA0, xYKYCSTA0 As typeYKYCSTA0
Dim blnYKYCSTA0_Update As Boolean, mKYCSTASTAK As String, mKYCSTASTAX As String, mKYCSTASTAY As String, mKYCSTADCLO As Long
Dim rsSab_YKYCSTA0 As New ADODB.Recordset
Dim blnXgsop_Archive As Boolean
Dim arrK3_Old(4, 6) As Integer, arrK3_New(4, 6) As Integer
Dim arrX_Lib(4) As String, arrY_Lib(6) As String

Dim wKYCCTL() As typeWKYCCTL, wKYCCTL_Nb As Long

Private Sub cmdSelect_SQL_ZRELEVE0()
Dim blnOk As Boolean
Dim blnCptOrdinaire As Boolean, blnRIB As Boolean, blnTest As Boolean, blnIban As Boolean
Dim xSQL As String, wRELEVEREL As String, mRELEVEREL As String
Dim rsSabX As New ADODB.Recordset

If Not blnAuto Then
    mRELEVEREL = InputBox("Code relevé à sélectionner : D M A W ou ? (autre)", , "?")
        
    If Trim(mRELEVEREL) = "" Then Exit Sub
Else
    mRELEVEREL = "?"
End If


fgDetail.Visible = False
fgDetail_Reset

fgDetail.FormatString = "<Compte                              |<Code Relevé" _
                       & "|<Devise   |<Date dernier mvt  |<PCI                 " _
                       & "|<Intitulé                                                                                                         "
         
fgDetail.Left = 100
fgDetail.Top = 150
fgDetail.Width = 13000

fgDetail.Rows = 1
fgDetail.Row = 0

Select Case mRELEVEREL
    Case "M", "A", "W", "D"
        wRELEVEREL = mRELEVEREL
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZRELEVE0 R , " _
                                & paramIBM_Library_SABSPE & ".YBIACPT0 C" _
             & " where RELEVEREL = '" & mRELEVEREL & "'" _
             & " and   RELEVECOM = C.COMPTECOM" _
             & " and COMPTEFON <> '4'" _
             & " order by C.COMPTECOM"
    
    Case Else
        mRELEVEREL = ""
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 C" _
             & " where COMPTEFON <> '4'" _
             & " order by C.COMPTECOM"
End Select
     
Set rsSab = cnsab.Execute(xSQL)


Do Until rsSab.EOF
    V = rsYBIACPT0_GetBuffer(rsSab, xYBIACPT0)
    blnOk = True
    
        Call fctPCEC_Atribut(xYBIACPT0.COMPTEOBL, xYBIACPT0.COMPTEDEV, blnCptOrdinaire, blnRIB, blnTest, blnIban)
        If Not blnCptOrdinaire Then
            blnOk = False
        Else
            If mRELEVEREL = "" Then
                wRELEVEREL = ""
                xSQL = "select RELEVEREL from " & paramIBM_Library_SAB & ".ZRELEVE0" _
                     & " where RELEVECOM = '" & xYBIACPT0.COMPTECOM & "'"
                     
                Set rsSabX = cnsab.Execute(xSQL)
            
                Do Until rsSabX.EOF
                    wRELEVEREL = rsSabX("RELEVEREL")
                    If wRELEVEREL = "M" Or wRELEVEREL = "W" Or wRELEVEREL = "D" Then
                    'Or rsSabX("RELEVEREL") = "A" Then
                        blnOk = False
                        Exit Do
                    End If
                    rsSabX.MoveNext
                Loop
            End If
        End If
                

    If blnOk Then
        fgDetail.Rows = fgDetail.Rows + 1
        fgDetail.Row = fgDetail.Rows - 1
        
        fgDetail.Col = 0: fgDetail.Text = xYBIACPT0.COMPTECOM
        fgDetail.Col = 1: fgDetail.Text = wRELEVEREL
        fgDetail.Col = 2: fgDetail.Text = xYBIACPT0.COMPTEDEV
        fgDetail.Col = 3: fgDetail.Text = dateImp_Amj(xYBIACPT0.SOLDEDMO + 19000000)
        fgDetail.Col = 4: fgDetail.Text = xYBIACPT0.COMPTEOBL
        fgDetail.Col = 5: fgDetail.Text = xYBIACPT0.COMPTEINT
        
    End If
    rsSab.MoveNext
Loop

fgDetail.Visible = True
fraSelect.Visible = True
Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "cmdSelect_SQL_ZRELEVE0 : " & fgDetail.Rows)

End Sub



Public Sub fgDetail_4_Exportation_Detail(lSheet As Integer, lCode As String, lLib As String, lLib2 As String, lCLIENARES As String)
On Error GoTo Error_Handler
Dim X As String
Dim K As Integer, iRow As Integer
Dim blnOk As Boolean

Dim wColor As Long

'==========================================================================================================

Set wsExcel = wbExcel.Sheets(lSheet)
wsExcel.Name = lLib2

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
    .Font.Name = "Calibri" '"Arial Unicode MS"
    .Font.Color = RGB(0, 64, 128)
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14" & lLib _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr _
                                & lstPLANCOPRO
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True
wsExcel.PageSetup.PrintTitleRows = "$A1:$R1"

If mSelect_SQL = "5" Then
    wsExcel.PageSetup.Zoom = 85
    mXls2_Col = 14
Else
    wsExcel.PageSetup.Zoom = 56
    mXls2_Col = 23
    wsExcel.Columns(19).HorizontalAlignment = Excel.xlHAlignRight
    wsExcel.Columns(19).NumberFormat = "### ##0.00"
    wsExcel.Columns(20).HorizontalAlignment = Excel.xlHAlignRight
    wsExcel.Columns(20).NumberFormat = "### ##0.00"
    wsExcel.Columns(21).HorizontalAlignment = Excel.xlHAlignRight
    wsExcel.Columns(21).NumberFormat = "###0.00"
    wsExcel.Columns(22).HorizontalAlignment = Excel.xlHAlignRight
    wsExcel.Columns(22).NumberFormat = "### ##0.00"
    wsExcel.Columns(23).HorizontalAlignment = Excel.xlHAlignRight
    wsExcel.Columns(23).NumberFormat = "### ##0.00"

End If

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

mXls2_Row = 1

fgDetail.Row = 0
For K = 1 To mXls2_Col
    fgDetail.Col = K - 1
    wsExcel.Columns(K).ColumnWidth = fgDetail.CellWidth / 100
    wsExcel.Cells(1, K) = fgDetail.Text

Next K
wsExcel.Columns(7).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(7).NumberFormat = "#### ### ##0.00"
wsExcel.Columns(10).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(10).NumberFormat = "#### ### ##0.00"
wsExcel.Columns(12).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(12).NumberFormat = "###0.00"

'wsExcel.Cells.EntireRow.AutoFit

For K = 1 To mXls2_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next
wsExcel.Columns(2).Font.Name = prtFontName_CourierNew
wsExcel.Columns(2).Font.Size = 6
wsExcel.Columns(18).Font.Name = prtFontName_CourierNew
wsExcel.Columns(18).Font.Size = 6

'==========================================================================================================
For iRow = 1 To fgDetail.Rows - 1
    fgDetail.Row = iRow
    blnOk = True
    
    If lCode <> "" Then
        fgDetail.Col = 3: X = Trim(fgDetail.Text)
        If lCode <> X Then blnOk = False
    End If
    
    If lCLIENARES <> "" Then
        fgDetail.Col = 15
        If lCLIENARES <> Trim(fgDetail.Text) Then blnOk = False
    End If
    If blnOk Then
    
        mXls2_Row = mXls2_Row + 1
        For K = 1 To mXls2_Col
            fgDetail.Col = K - 1
            Select Case K
                Case 7, 10, 12, 19, 20, 21, 22, 23: If Val(fgDetail.Text) <> 0 Then wsExcel.Cells(mXls2_Row, K) = Val(fgDetail.Text)
                Case Else: wsExcel.Cells(mXls2_Row, K) = fgDetail.Text
            End Select
            
            If fgDetail.CellBackColor <> 0 Then wsExcel.Cells(mXls2_Row, K).Interior.Color = fgDetail.CellBackColor
            wsExcel.Cells(mXls2_Row, K).Font.Color = fgDetail.CellForeColor
            
'            If K = 2 Then
'                wsExcel.Cells(mXls2_Row, K).Font.Name = prtFontName_CourierNew
'                wsExcel.Cells(mXls2_Row, K).Font.Size = 7
'            End If
            If (iRow Mod 10) = 0 Then Call lstErr_ChangeLastItem(lstErr, cmdContext, "> NB : " & iRow & " / " & fgDetail.Rows - 1): DoEvents
    
        Next K
    End If
Next iRow
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name


End Sub

Public Sub fgDetail_1r_Exportation_Detail(lLib As String, lLib2 As String)
On Error GoTo Error_Handler
Dim X As String
Dim K As Integer, iRow As Integer
Dim blnOk As Boolean

Dim wColor As Long

'==========================================================================================================

Set wsExcel = wbExcel.Sheets(1)
wsExcel.Name = lLib2

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
    .Font.Name = "Calibri" '"Arial Unicode MS"
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

wsExcel.PageSetup.Zoom = 85

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

mXls2_Col = 7
mXls2_Row = 1

fgDetail.Row = 0
For K = 1 To mXls2_Col
    fgDetail.Col = K - 1
    wsExcel.Columns(K).ColumnWidth = fgDetail.CellWidth / 100
    wsExcel.Cells(1, K) = fgDetail.Text

Next K

wsExcel.Columns(1).ColumnWidth = 50
wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(6).HorizontalAlignment = Excel.xlHAlignCenter

'wsExcel.Cells.EntireRow.AutoFit

For K = 1 To mXls2_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next
'==========================================================================================================
For iRow = 1 To fgDetail.Rows - 1
    fgDetail.Row = iRow
    mXls2_Row = mXls2_Row + 1
    For K = 1 To mXls2_Col
        fgDetail.Col = K - 1
        wsExcel.Cells(mXls2_Row, K) = fgDetail.Text
        If fgDetail.CellBackColor <> 0 Then wsExcel.Cells(mXls2_Row, K).Interior.Color = fgDetail.CellBackColor
        wsExcel.Cells(mXls2_Row, K).Font.Color = fgDetail.CellForeColor
        If (iRow Mod 10) = 0 Then Call lstErr_ChangeLastItem(lstErr, cmdContext, "> NB : " & iRow & " / " & fgDetail.Rows - 1): DoEvents

    Next K
Next iRow
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name


End Sub


Public Sub fgDetail_4_Exportation()
On Error GoTo Error_Handler
Dim X As String, K As Long, xWhere As String
Dim wFile As String, wFilex As String
Dim blnCALCS As Boolean
Dim xLib As String, xLib2 As String, xCLIENARES As String
Dim mLib As String
On Error GoTo Error_Handler
'===================================================================================
If blnAuto Then
    X = paramServer("\\CPT_Archive\")
Else
    X = ""
End If
If X = "" Then X = "C:\Temp\"
If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"

blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

xLib2 = Trim(txtSelect_Options_4_CLIENACLI)
If xLib2 = "" Then xLib2 = "global"
xLib = "SAB_Client " & xLib2

xLib = xLib & " - CAV conditions échelles"
xCLIENARES = Mid$(cboSelect_Options_4_CLIENARES, 1, 3)

wFile = X & xLib & " " & dateImp_Amj(DSys) & ".xlsx"

If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "SAB_Client : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If

Call lstErr_AddItem(lstErr, cmdContext, "> SAB_Client exportation"): DoEvents


If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile
'_________________________________________


Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "SAB_Client"
    .Subject = ""
End With

Call fgDetail_4_Exportation_Detail(1, "3", xLib, "Conditions", xCLIENARES)
Call fgDetail_4_Exportation_Detail(2, "1", xLib, "Autorisations", xCLIENARES)
Call fgDetail_4_Exportation_Detail(3, "0", xLib, "Standard", xCLIENARES)

wbExcel.SaveAs wFile
wbExcel.Close
appExcel.Quit

'===================================================================================================
If mSelect_SQL = "4" And xCLIENARES = "" Then
    X = MsgBox("Voulez-vous extraire les conditions par RESPONSABLE ?", vbQuestion + vbYesNo, "Exportation des conditions échelles")
    If X = vbYes Then
        wFilex = wFile
        mLib = xLib
        For K = 1 To arrCLIENARES_Nb

            xLib2 = arrCLIENARES(K)
            wFile = Replace(wFilex, "global", xLib2)
            xLib = Replace(mLib, "global", xLib2)
            If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile
            '_________________________________________
                            
            Call lstErr_AddItem(lstErr, cmdContext, "> SAB_Client exportation : " & xLib2): DoEvents
            
            Set appExcel = CreateObject("Excel.Application")
            appExcel.Workbooks.Add
            Set wbExcel = appExcel.ActiveWorkbook
            With wbExcel
                .Title = "SAB_Client"
                .Subject = ""
            End With
            
            Call fgDetail_4_Exportation_Detail(1, "3", xLib, "Conditions", arrCLIENARES(K))
            Call fgDetail_4_Exportation_Detail(2, "1", xLib, "Autorisations", arrCLIENARES(K))
            
            wbExcel.SaveAs wFile
            wbExcel.Close
            appExcel.Quit
    
        
        Next K
    End If

End If

'===================================================================================================

Exit_sub:
'__________________________________________________________________________________

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

End Sub

Public Sub fgDetail_5_Exportation()
On Error GoTo Error_Handler
Dim X As String, K As Long, xWhere As String
Dim wFile As String, wFilex As String
Dim blnCALCS As Boolean
Dim xLib As String, xLib2 As String, xCLIENARES As String
Dim mLib As String
On Error GoTo Error_Handler
'===================================================================================
If blnAuto Then
    X = paramServer("\\CPT_Archive\")
Else
    X = ""
End If
If X = "" Then X = "C:\Temp\"
If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"

blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

xLib2 = Trim(txtSelect_Options_4_CLIENACLI)
If xLib2 = "" Then xLib2 = "global"
xLib = "SAB_Client " & xLib2
xCLIENARES = ""
Select Case mSelect_SQL
    Case "4!e": xLib = xLib & " - CAV surveillance autorisation - échelles"
    Case "5": xLib = xLib & " - DEC surveillance autorisation - code blocage"
End Select

wFile = X & xLib & " " & dateImp_Amj(DSys) & ".xlsx"

If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "SAB_Client : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If

Call lstErr_AddItem(lstErr, cmdContext, "> SAB_Client exportation"): DoEvents


If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile
'_________________________________________


Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "SAB_Client"
    .Subject = ""
End With

Call fgDetail_4_Exportation_Detail(1, "", xLib, xLib2, xCLIENARES)

wbExcel.SaveAs wFile
wbExcel.Close
appExcel.Quit

'===================================================================================================

Exit_sub:
'__________________________________________________________________________________

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

End Sub



Public Sub fgDetail_1r_Exportation()
On Error GoTo Error_Handler
Dim X As String, K As Long, xWhere As String
Dim wFile As String, wFilex As String
Dim blnCALCS As Boolean
Dim xLib As String, xLib2 As String
On Error GoTo Error_Handler
'===================================================================================
If blnAuto Then
    X = paramServer("\\CPT_Archive\")
Else
    X = ""
End If
If X = "" Then X = "C:\Temp\"
If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"

blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

xLib2 = Trim(txtSelect_CLIENARA1)
If chkSelect_Groupes = "1" Then
    xLib2 = "Groupes"
    xLib = "SAB_Client-Synthèse des relations hierarchiques " & xLib2
Else
    If xLib2 = "" Then xLib2 = "clientèle"
    xLib = "SAB_Client-Synthèse des relations hierarchiques " & xLib2
End If

wFile = X & xLib & " " & dateImp_Amj(DSys) & ".xlsx"

If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "SAB_Client : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If

If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile
'_________________________________________


Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "SAB_Client"
    .Subject = ""
End With

Call fgDetail_1r_Exportation_Detail(xLib, xLib2)

wbExcel.SaveAs wFile
wbExcel.Close
appExcel.Quit
'===================================================================================================
Exit_sub:
'__________________________________________________________________________________

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

End Sub




Private Sub fgDetail_Display(lCLIENACLI As String)
Dim wColor As Long
Dim X As String, xWhere As String, xOPE As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler

fgDetail.Left = 6210
fgDetail.Top = 150
fgDetail.Width = 6825
fgDetail.FormatString = fgDetail_FormatString
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.Row = 0

currentAction = "fgDetail_Display"
If optSelect_Options_3C Then
    X = "('91120','12120')"
Else
    X = "('98050','12120')"
End If


X = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
    & " where CLIENACLI = '" & lCLIENACLI & "' and substr(COMPTEOBL,1,5) in " & X _
    & " order by CLIENACLI,COMPTEOBL"


Set rsAdo = cnAdo.Execute(X)

Do While Not rsAdo.EOF
    Call rsYBIACPT0_GetBuffer(rsAdo, xYBIACPT0)
    
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_DisplayLine I
    
    rsAdo.MoveNext

Loop
         


fgDetail.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgDetail_1r_Display()
Dim I As Long, blnEnd As Boolean

On Error GoTo Error_Handler

fraSelect.Visible = False
fgDetail.Visible = False

fgDetail.Left = 100
fgDetail.Top = 150
fgDetail.Width = 13000
fgDetail_Reset
fgDetail.FormatString = "<Hierarchie                                                                                " _
                      & "|>Niveau|<Relation                  |>             %     |<Racine                    |<Nationalité|<Intitulé                                                                                                                                             "

fgDetail.Rows = 1
fgDetail.Row = 0

currentAction = "fgDetail_1r_Display"

ReDim arrCLIRGPREG(1000) As String, arrCLIRGPREG_Niveau(1000), arrCLIRGPREG_Hierarchie(1000)

For I = 1 To arrZCLIENA0_NB
    xZCLIENA0 = arrZCLIENA0(I)
    arrCLIRGPREG_Nb = 1: arrCLIRGPREG(1) = xZCLIENA0.CLIENACLI: arrCLIRGPREG_Niveau(1) = 0
    arrCLIRGPREG_Hierarchie(1) = xZCLIENA0.CLIENACLI
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_1r_DisplayLine I
    
    blnEnd = False
    arrCLIRGPREG_K = 0
    Do
        arrCLIRGPREG_K = arrCLIRGPREG_K + 1
        fgDetail_1r_Display_ZCLIGRP0 arrCLIRGPREG_K
    Loop Until arrCLIRGPREG_K = arrCLIRGPREG_Nb
Next I
         
fgDetail_Sort1 = 0: fgDetail_Sort2 = 1: fgDetail_Sort
fgDetail.Visible = True
         
fraSelect.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub fgDetail_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim blnSolde As Boolean

On Error Resume Next
fgDetail.Col = 0: fgDetail.Text = xYBIACPT0.PLANCOPRO

fgDetail.Col = 1: fgDetail.Text = xYBIACPT0.COMPTECOM
fgDetail.Col = 2: fgDetail.Text = xYBIACPT0.COMPTEDEV
fgDetail.Col = 3: fgDetail.Text = xYBIACPT0.COMPTEFON
fgDetail.Col = 4: fgDetail.Text = xYBIACPT0.COMPTEINT

End Sub


Public Sub fgDetail_1r_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, xRA1 As String
Dim blnSolde As Boolean

On Error Resume Next
xRA1 = Trim(xZCLIENA0.CLIENARA1) & " " & Trim(xZCLIENA0.CLIENARA2)
If Mid$(xZCLIENA0.CLIENARES, 1, 1) = "X" Then
    xRA1 = "## " & LCase$(xRA1) & " ##"
    wColor = RGB(90, 90, 90)
Else
    wColor = vbBlue
End If


fgDetail.Col = 0: fgDetail.Text = xZCLIENA0.CLIENACLI
fgDetail.CellForeColor = wColor
fgDetail.Col = 1: fgDetail.Text = 0
fgDetail.CellForeColor = wColor
fgDetail.Col = 2: fgDetail.Text = ""
fgDetail.CellForeColor = wColor
fgDetail.Col = 4: fgDetail.Text = xZCLIENA0.CLIENACLI
fgDetail.CellForeColor = wColor
fgDetail.Col = 5: fgDetail.Text = xZCLIENA0.CLIENANAT
fgDetail.CellForeColor = wColor
fgDetail.Col = 6: fgDetail.Text = xRA1
fgDetail.CellForeColor = wColor

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
    fgDetail.ColSel = 0
End If

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
Private Sub fgSelect_Display()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 2
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgselect_Display"
    
Do While Not rsAdo.EOF

    Call srvYUPDLOG0_GetBuffer_ODBC(rsAdo, meYUPDLOG0)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine I
    rsAdo.MoveNext

Loop
fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort

fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Nb mises à jour : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Private Sub fgSelect_YKYCDOSH_Display()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 2
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<Mise à jour            |<Utilisateur|<Fonction           |<Nature            " _
                      & "|<Identifiant|>Seq  |>Seq 2|<O/E|<PJ|<Echéance  |<Date document" _
                      & "|<Commentaire                                            "
currentAction = "fgselect_Display"
    
Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_YKYCDOSH_DisplayLine I
    rsSab.MoveNext

Loop

fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Nb mises à jour : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


Private Sub arrZCLIENA0_sql(xWhere As String)
Dim V
Dim X As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrZCLIENA0(501)
arrZCLIENA0_Max = 500: arrZCLIENA0_NB = 0

Set rsAdo = Nothing

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 " & xWhere

Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
    V = rsZCLIENA0_GetBuffer(rsAdo, xZCLIENA0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, Me.Name & ".arrZCLIENA0_sql"
        '' Exit Sub
     Else
         arrZCLIENA0_NB = arrZCLIENA0_NB + 1
         If arrZCLIENA0_NB > arrZCLIENA0_Max Then
             arrZCLIENA0_Max = arrZCLIENA0_Max + 100
             ReDim Preserve arrZCLIENA0(arrZCLIENA0_Max)
         End If
         
         arrZCLIENA0(arrZCLIENA0_NB) = xZCLIENA0
    End If
    rsAdo.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub arrZCLIGRP0_sql(xWhere As String, blnSelect As Boolean)
Dim V, I As Integer, wCLIENACLI As String
Dim X As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrZCLIGRP0(501)
arrZCLIGRP0_Max = 500: arrZCLIGRP0_Nb = 0

rsZCLIGRP0_Init xZCLIGRP0
Set rsAdo_ZCLIGRP0 = Nothing

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIGRP0  where " & xWhere

Set rsAdo_ZCLIGRP0 = cnAdo.Execute(xSQL)

Do While Not rsAdo_ZCLIGRP0.EOF
    xZCLIGRP0.CLIGRPCLI = rsAdo_ZCLIGRP0("CLIGRPCLI")
    xZCLIGRP0.CLIGRPREG = rsAdo_ZCLIGRP0("CLIGRPREG")
    If Mid$(xZCLIGRP0.CLIGRPREG, 1, 2) = "9X" Then
        Mid$(xZCLIGRP0.CLIGRPREG, 1, 2) = "99"
        xZCLIGRP0.CLIGRPREG_9X = True
    End If
    xZCLIGRP0.CLIGRPREL = rsAdo_ZCLIGRP0("CLIGRPREL")
    xZCLIGRP0.CLIGRPTAU = rsAdo_ZCLIGRP0("CLIGRPTAU")
    xZCLIGRP0.CLIGRPCLI_RA1 = "??"
         arrZCLIGRP0_Nb = arrZCLIGRP0_Nb + 1
         If arrZCLIGRP0_Nb > arrZCLIGRP0_Max Then
             arrZCLIGRP0_Max = arrZCLIGRP0_Max + 100
             ReDim Preserve arrZCLIGRP0(arrZCLIGRP0_Max)
         End If
         
         arrZCLIGRP0(arrZCLIGRP0_Nb) = xZCLIGRP0
    rsAdo_ZCLIGRP0.MoveNext

Loop


For I = 1 To arrZCLIGRP0_Nb

    If blnSelect Then
        wCLIENACLI = arrZCLIGRP0(I).CLIGRPREG
    Else
        wCLIENACLI = arrZCLIGRP0(I).CLIGRPCLI
   End If
    xSQL = "select CLIENARA1,CLIENARA2,CLIENARES,CLIENANAT from " & paramIBM_Library_SAB & ".ZCLIENA0  where CLIENACLI = '" & wCLIENACLI & "'"
    Set rsAdo_ZCLIGRP0 = cnAdo.Execute(xSQL)
    If Not rsAdo_ZCLIGRP0.EOF Then
        arrZCLIGRP0(I).CLIGRPCLI_RA1 = rsAdo_ZCLIGRP0("CLIENARA1")
        arrZCLIGRP0(I).CLIGRPCLI_RA2 = rsAdo_ZCLIGRP0("CLIENARA2")
        arrZCLIGRP0(I).CLIGRPCLI_RES = rsAdo_ZCLIGRP0("CLIENARES")
        arrZCLIGRP0(I).CLIGRPCLI_NAT = rsAdo_ZCLIGRP0("CLIENANAT")
    End If
Next I
Exit Sub
Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
Dim X As String, lenX As Integer
Dim xSQL As String
On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = meYUPDLOG0.UPDLOGID
fgSelect.Col = 1: fgSelect.Text = dateImp10(meYUPDLOG0.UPDLOGAMJ)
fgSelect.Col = 2: fgSelect.Text = timeImp(meYUPDLOG0.UPDLOGHMS)
fgSelect.Col = 3: fgSelect.Text = meYUPDLOG0.UPDLOGUSR
fgSelect.Col = 4: fgSelect.Text = meYUPDLOG0.UPDLOGAPP
fgSelect.Col = 5: fgSelect.Text = meYUPDLOG0.UPDLOGFCT
fgSelect.Col = 6: fgSelect.Text = meYUPDLOG0.UPDLOGTXT
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

End Sub


Public Sub fgSelect_YKYCDOSH_DisplayLine(lIndex As Long)
Dim X As String, lenX As Integer
Dim xSQL As String
On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = dateImp10_S(rsSab("KYCDOSUAMJ")) & "-" & timeImp8(rsSab("KYCDOSUHMS")) & "-" & rsSab("KYCDOSUVER")
fgSelect.Col = 1: fgSelect.Text = rsSab("KYCDOSUUSR")
fgSelect.Col = 2: fgSelect.CellBackColor = mColor_G0
Select Case rsSab("KYCDOSUFCT")
    Case "A": fgSelect.Text = "Ajout"
    Case "U": fgSelect.Text = "Mofification"
    Case "D": fgSelect.Text = "Supression"
    Case "J": fgSelect.Text = "màj PJ"
    Case "C": fgSelect.Text = "Type Clientèle"
    Case "+": fgSelect.Text = "PJ Ajout"
    Case "-": fgSelect.Text = "PJ Supression"
    Case "Z": fgSelect.Text = "Effacement du dossier"
    Case "#": fgSelect.Text = "Statut du dossier"
    Case "?": fgSelect.Text = "Document manquant"
    Case "I": fgSelect.Text = "Document à ignorer"
    Case Else: fgSelect.Text = rsSab("KYCDOSUFCT")
End Select


fgSelect.Col = 3
Select Case rsSab("KYCDOSNAT")
    Case " ": fgSelect.Text = "Client"
    Case "D": fgSelect.Text = "Document"
    Case "J": fgSelect.Text = "Justificatif"
    Case "*": fgSelect.Text = "Type clientèle"
    Case "=": fgSelect.Text = "KYC référence"
    Case Else: fgSelect.Text = rsSab("KYCDOSNAT")
End Select
fgSelect.Col = 4: fgSelect.Text = rsSab("KYCDOSID"): fgSelect.CellBackColor = mColor_G0
fgSelect.Col = 5: fgSelect.Text = rsSab("KYCDOSSEQ")
fgSelect.Col = 6: fgSelect.Text = rsSab("KYCDOSSEQ2")
fgSelect.Col = 7: fgSelect.Text = rsSab("KYCDOSSTAK")
fgSelect.Col = 8: fgSelect.Text = rsSab("KYCDOSPJ")
If rsSab("KYCDOSDECH") > 0 Then fgSelect.Col = 9: fgSelect.Text = dateImp10_S(rsSab("KYCDOSDECH"))
If rsSab("KYCDOSDAMJ") > 0 Then fgSelect.Col = 10: fgSelect.Text = dateImp10_S(rsSab("KYCDOSDAMJ"))
fgSelect.Col = 11: fgSelect.Text = rsSab("KYCDOSDLIB")

End Sub

Public Sub fgSelect_Reset()

fgSelect.Top = 1200

fgSelect.Clear
fgSelect.FormatString = fgSelect_FormatString
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = fgSelect.Cols - 1
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

End Sub

Public Sub fgYKYCDOS0_Reset()
fgYKYCDOS0.Clear
fgYKYCDOS0.FormatString = fgYKYCDOS0_FormatString
fgYKYCDOS0_Sort1 = 0: fgYKYCDOS0_Sort2 = 0
fgYKYCDOS0_Sort1_Old = -1
fgYKYCDOS0_RowDisplay = 0: fgYKYCDOS0_RowClick = 0
fgYKYCDOS0_arrIndex = fgYKYCDOS0.Cols - 1
blnfgYKYCDOS0_DisplayLine = False
fgYKYCDOS0_SortAD = 6
fgYKYCDOS0.LeftCol = 0

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

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

'Call BiaPgmAut_Init(Mid$(Msg, 1, 12), SAB_CLIENT_Aut)
Call BIA_VB_HAB(Mid$(Msg, 1, 12), arrHab(), cboSelect_SQL)

Form_Init


Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@SAB_CLIENT": blnAuto = True
    
' 2013-10-01 remplacé par AUTO_SAB_CLIENT
'-----------------------------------------
        'If Mid$(YBIATAB0_DATE_CPT_J, 1, 6) <> Mid$(YBIATAB0_DATE_CPT_JS1, 1, 6) Then
            cmdSelect_SQL_Xgsop_Init
            fraSelect_Options_Xgsop.Visible = True
            blnYKYCSTA0_Update = True
            Call cmdSelect_SQL_Xgsop_Auto
        'End If
        
' 2015-01-19 demande CCGA (F. Legouard)
            mSelect_SQL = "KYC ech"
            Call cmdSelect_SQL_YKYCDOS0_Ech
            Call mnuPrint2_Mail_Click
            
' 2016-01-04
            mSelect_SQL = "KYC Releve"
            Call cmdSelect_SQL_ZRELEVE0
            Call mnuPrint2_Mail_Click
        
        
        Unload Me

    Case Else: blnAuto = False
End Select


End Sub


Public Sub Form_Init()
Dim xSQL As String
Me.Enabled = False
Me.MousePointer = vbHourglass
blnControl = False

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents


Set fraSelect_Options_3.Container = fraSelect_Options
fraSelect_Options_3.Left = 60 '4380
fraSelect_Options_3.Top = 0

Set fraSelect_Options_4.Container = fraSelect_Options
fraSelect_Options_4.Left = 60 '4380
fraSelect_Options_4.Top = 0

Set fraSelect_Options_No.Container = fraSelect_Options
fraSelect_Options_No.Left = 60 '4380
fraSelect_Options_No.Top = 0

Set fraSelect_Options_Xgsop.Container = fraSelect_Options
fraSelect_Options_Xgsop.Left = 60 '4380
fraSelect_Options_Xgsop.Top = 0

Set fraSelect_Options_KYCgsop.Container = fraSelect_Options
fraSelect_Options_KYCgsop.Left = 60 '4380
fraSelect_Options_KYCgsop.Top = 0

Set fraSelect_Options_KYCech.Container = fraSelect_Options
fraSelect_Options_KYCech.Left = 60 '4380
fraSelect_Options_KYCech.Top = 0
Call DTPicker_Set(txtSelect_Options_KYCech_KYCDOSDECH, DSys)
'

fgDetail.Visible = False
fgDetail_FormatString = fgDetail.FormatString
fgDetail.BackColorFixed = mColor_GB
fgDetail.ForeColorFixed = vbWhite

fraParam_YKYCDOS0_JD.Left = lstParam_KYCDOSNAT.Left + 600

fraParam_YKYCDOS0_JD.Top = lstParam_KYCDOSNAT.Top

Call DTPicker_Set(txtUpdLog_AmjMin, DSys)
txtUpdLog_AmjMin.Visible = False

SSTab1.Tab = 1

fraYKYCDOS_ZCLIENA0.Top = 120
fraYKYCDOS_ZCLIENA0.Left = 120
Set fraYKYCDOS0.Container = SSTab1
fraYKYCDOS0.Top = fraUpdate.Top
fraYKYCDOS0.Left = fraUpdate.Left
fgYKYCDOS0_ZADRESS0_FormatString = fgYKYCDOS0_ZADRESS0.FormatString
fgYKYCDOS0_FormatString = fgYKYCDOS0.FormatString
fgYKYCDOS0.Top = lstYKYCDOS0_CLIENACAT.Top
fgYKYCDOS0.Left = lstYKYCDOS0_CLIENACAT.Left

fraYKYCDOS0_Update.Top = fgYKYCDOS0_ZADRESS0.Top
fraYKYCDOS0_Update.Height = fgYKYCDOS0_ZADRESS0.Height
fraYKYCDOS0_Update.Left = fraYKYCDOS0.Left + fraYKYCDOS0.Width - fraYKYCDOS0_Update.Width - 120
cmdYKYCDOS0_PJ.Top = cmdYKYCDOS0_Add.Top
cmdYKYCDOS0_PJ.Left = cmdYKYCDOS0_Add.Left
cmdYKYCDOS0_Missing.Top = cmdYKYCDOS0_Delete.Top
cmdYKYCDOS0_Missing.Left = cmdYKYCDOS0_Delete.Left
cmdYKYCDOS0_Ignore.Top = cmdYKYCDOS0_Update.Top
cmdYKYCDOS0_Ignore.Left = cmdYKYCDOS0_Update.Left

Set fraPJ.Container = fraYKYCDOS0
fraPJ.Visible = False
fraPJ.Left = fraYKYCDOS0.Width - fraPJ.Width - 200
fraPJ.Top = fgYKYCDOS0.Top
fraYKYCDOS0_Update.Visible = False

fraParam_YKYCDOS0_Update.Top = lstParam_KYCDOSNAT.Top
fraParam_YKYCDOS0_Update.Left = lstParam_KYCDOSNAT.Left + lstParam_KYCDOSNAT.Width + 200
cmdParam_YKYCDOS0_4c_Actualisation.Left = lstParam_KYCDOSNAT.Left + lstParam_KYCDOSNAT.Width + 3000
cmdParam_YKYCDOS0_4c_Actualisation.Visible = False

fraParam_YKYCDOS0_4c.Left = fraParam_YKYCDOS0_Update.Left
fraParam_YKYCDOS0_4c.Top = fraParam_YKYCDOS0_Update.Top
fraParam_YKYCDOS0_4c.Visible = False
fgParam_YKYCDOS0_4c.Visible = False
fgParam_YKYCDOS0_4c.Left = lstParam_YKYCDOS0.Left + 2000
fgParam_YKYCDOS0_4c.Top = lstParam_YKYCDOS0.Top
fraParam_YKYCDOS0_4c.ForeColor = &H800080

Call lstParam_YKYCDOS0_J_Load
Call lstParam_YKYCDOS0_D_Load
mKYCDOSDECH_Warn = dateElp("Jour", -100, DSys)

dirListBox.PATH = "\\DOCSRV2013\_SCAN" '"C:\Temp"
blnfilDoc_Path = False
On Error Resume Next
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'GSOP_PJ**' and BIATABK1 = '" & usrName_UCase & "'"
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    dirListBox.PATH = Trim(rsSab("BIATABTXT"))
    blnfilDoc_Path = True
End If

mfilDoc_Path = dirListBox.PATH
cmdPJ_Path.Visible = False


SSTab1.Tab = 0

If Not IsNull(param_Init) Then
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmSAB_CLIENT.paramSAA_Init"
    Unload Me
Else
    lstErr.Clear
End If



fgSelect_FormatString = fgSelect.FormatString

fgSelect.Enabled = True
blnControl = True
If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0

Me.Enabled = True
Me.MousePointer = 0
End Sub


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
fraSelect_Options.Enabled = True
fraUpdate.Visible = False
fraUpdate_Add.ForeColor = vbRed
fraUpdate_Détail.ForeColor = vbRed
fraSelect_Update.Visible = False
optUpdate_Add_Old = True
optSelect_YKYCDOS0.BackColor = &HC0FFC0

'cmdSelect_Ok_Click


blnControl = True



End Sub


'---------------------------------------------------------
Public Sub cmdSelect_Clear()
'---------------------------------------------------------
Dim K As Integer, X As String
blnControl = True
currentAction = ""
SSTab1.Tab = 0
fraSelect_Options.Enabled = True
fraUpdate.Visible = False
fraSelect_Update.Visible = False
optUpdate_Add_Old = True
fgSelect.Visible = False
fraUpdate.Visible = False
fraSelect_Update.Visible = False
fraSelect.Visible = False
fgDetail.Clear
fgDetail.Visible = False
chkSelect_Racine.Visible = False
chkSelect_Groupes.Visible = False
cmdParam_YKYCDOS0_4c_Actualisation.Visible = False
fraParam_YKYCDOS0_4c.Visible = False
fgParam_YKYCDOS0_4c.Visible = False
fraYKYCDOS0.Visible = False
fraPJ.Visible = False
blnYKYCSTA0_Update = False

'Select Case cboSelect_SQL.ListIndex
'K = InStr(cboSelect_SQL, "-")
'If K > 0 Then
'    X = Trim(Mid$(cboSelect_SQL, 1, K - 1))
'Else
'    Trim (Mid$(cboSelect_SQL, 1, 3))
'End If

'Select Case X
'    Case Is = "1": lblSelect_CLIENACAT = "Catégorie client": cboSelect_CLIENACAT = ""
'    Case Is = "1r": lblSelect_CLIENACAT = "Catégorie client": cboSelect_CLIENACAT = "": chkSelect_Racine.Visible = True: chkSelect_Groupes.Visible = True
'    Case Is = "2": lblSelect_CLIENACAT = "Type de lien": cboSelect_CLIENACAT = ""
'    Case Is = "3": fraSelect_Options_3.Visible = True
'    Case Is = "4": fraSelect_Options_4.Visible = True
'    Case Is = "5": fraSelect_Options_4.Visible = True
'    'Case Is = "Xgs", "KYC gsop", "Xgsop@": fraSelect_Options_Xgsop.Visible = True
'    Case "4!e": fraSelect_Options_No.Visible = True

'End Select




End Sub


Public Function param_Init()
Dim xSQL As String

param_Init = Null
Call lstErr_Clear(lstErr, cmdContext, "SAB_CLIENT : param_init"): DoEvents

fraSelect.Visible = False

lstPLANCOPRO = "'CAV','DTT','DTX','IDH','IMP','LDX','LIE','LOR','NOB','NOS','TDT'"

'cboSelect_SQL.Clear
'cboSelect_SQL.AddItem "1  - Arborescence des relations  'clientèle'"
'cboSelect_SQL.AddItem "1r - Synthèse des relations hierarchiques 'clientèle'"
'cboSelect_SQL.AddItem "2  - Arborescence des mandataires"
'cboSelect_SQL.AddItem "3  - Bq émettrice de credoc sans compte LOR"
'cboSelect_SQL.AddItem "4  - CAV : conditions autorisation / échelles"
'cboSelect_SQL.AddItem "4!e- CAV : surveillance autorisation / échelles"
'cboSelect_SQL.AddItem "5  - DEC : surveillance autorisation / code blocage"
'cboSelect_SQL.AddItem "Xf - Exportation fiche 'Clients actifs'"
'cboSelect_SQL.AddItem "Xg - Exportation Groupes-Clients-Comptes"
'cboSelect_SQL.AddItem "Xp - Publipostage ( Clients actifs + adresse courrier)"
'cboSelect_SQL.AddItem "Xa - Exportation Clients actifs + comptes + adresses"
'cboSelect_SQL.AddItem "Xa*- Exportation TOUS les Clients + comptes + adresses"
'cboSelect_SQL.AddItem "Xz - Exportation Comptes 'client' à clôturer(sans solde depuis ? mois)"
'cboSelect_SQL.AddItem "XzR- Exportation par RES des Comptes 'client' à clôturer(sans solde depuis ? mois)"
'If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0

cboSelect_CLIENACAT.Clear
rsYBIATAB0_cboK2 "SAB", "CLIENACAT", cboSelect_CLIENACAT
cboSelect_CLIENACAT.AddItem " "
cboSelect_CLIENACAT.ListIndex = 0

cboUpdate_CLIENAETA.Clear
rsYBIATAB0_cboK2 "SAB", "CLIENAETA", cboUpdate_CLIENAETA

cboUpdate_ADRESSPAY.Clear
rsYBIATAB0_cboK2 "SAB", "CLIENAPAY", cboUpdate_ADRESSPAY

cboUpdate_CLIENARSD.Clear
rsYBIATAB0_cboK2 "SAB", "CLIENAPAY", cboUpdate_CLIENARSD

cboUpdate_CLIENANAT.Clear
rsYBIATAB0_cboK2 "SAB", "CLIENAPAY", cboUpdate_CLIENANAT

cboUpdate_CLIENBNAS.Clear
rsYBIATAB0_cboK2 "SAB", "CLIENAPAY", cboUpdate_CLIENBNAS


xSQL = "select count(*) from " & paramIBM_Library_SAB & ".ZBASTAB0" _
     & " where BASTABETA = 1 and BASTABNUM = 12"
Set rsSab = cnsab.Execute(xSQL)
arrZBAST12_Nb = rsSab(0)
ReDim arrZBAST12_Arg(arrZBAST12_Nb + 1), arrZBAST12_lib(arrZBAST12_Nb + 1)

xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
     & " where BASTABETA = 1 and BASTABNUM = 12 order by BASTABARG"
Set rsSab = cnsab.Execute(xSQL)
arrZBAST12_Nb = 0
Do While Not rsSab.EOF
    arrZBAST12_Nb = arrZBAST12_Nb + 1
    arrZBAST12_Arg(arrZBAST12_Nb) = Mid$(rsSab("BASTABARG"), 4, 3)
    arrZBAST12_lib(arrZBAST12_Nb) = rsSab("BASTABLO1")
    
    rsSab.MoveNext
Loop


xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
     & " where BASTABETA = 1 and BASTABNUM = 6 order by BASTABARG"
Set rsSab = cnsab.Execute(xSQL)
cboSelect_Options_4_CLIENARES.Clear
cboSelect_Options_4_CLIENARES.AddItem ""
Do While Not rsSab.EOF
    cboSelect_Options_4_CLIENARES.AddItem Mid$(rsSab("BASTABARG"), 4, 3) & " - " & Trim(rsSab("BASTABLO1"))
    rsSab.MoveNext
Loop

cboUpdate_CLIENBLIE.Clear
cboUpdate_CLIENBLIE.AddItem "1"
cboUpdate_CLIENBLIE.AddItem "2"
cboUpdate_CLIENBLIE.AddItem "3"
cboUpdate_CLIENBLIE.AddItem "4"
cboUpdate_CLIENBLIE.ListIndex = 0

cboUpdate_CLIENBTER.Clear
cboUpdate_CLIENBTER.AddItem "1"
cboUpdate_CLIENBTER.AddItem "2"
cboUpdate_CLIENBTER.AddItem "3"
cboUpdate_CLIENBTER.AddItem "4"
cboUpdate_CLIENBTER.ListIndex = 0


cboSelect_Options_4_Code.Clear
cboSelect_Options_4_Code.AddItem ""
cboSelect_Options_4_Code.AddItem "CDM"
cboSelect_Options_4_Code.AddItem "ICR"
cboSelect_Options_4_Code.AddItem "IDE"
cboSelect_Options_4_Code.AddItem "PFD"
cboSelect_Options_4_Code.AddItem "TDC"
cboSelect_Options_4_Code.ListIndex = 0

'_____________________________________________________________________________________
If arrHab(18) Then
    fgParam_FormatString = fgParam.FormatString
    fraParam.Visible = True
    fraParam_Update.Visible = False
    fgParam.Clear
    txtParam_Id = ""
    
    lstParam.AddItem "1 - Clients non réclamés"
    lstParam.AddItem "2 - Clients gérés hors GSOP"
    lstParam.AddItem "3 - Code produit à analyser"
End If

If arrHab(16) Then

    fraParam_YKYCDOS0.Visible = True
    fraParam_YKYCDOS0_Update.Visible = False
    lstParam_YKYCDOS0.Visible = False
    lstParam_YKYCDOS0_J.BackColor = libParam_YKYCDOS0_J.BackColor
    'lstParam_YKYCDOS0_JD.BackColor = cmdParam_YKYCDOS0_JD_Update.BackColor
    lstParam_YKYCDOS0_D.BackColor = libParam_YKYCDOS0_D.BackColor
    lstParam_KYCDOSNAT.AddItem "1 - Documents"
    lstParam_KYCDOSNAT.AddItem "2 - Justificatifs"
    lstParam_KYCDOSNAT.AddItem "3 - Type de clientèle"
    lstParam_KYCDOSNAT.AddItem "4 - KYC"
    lstParam_KYCDOSNAT.AddItem "4c- KYC commentaires"
End If

Me.Enabled = True: Me.MousePointer = 0

End Function


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

Private Function cmdParam_YKYCDOS0_Transaction(lFct As String)
Dim xSQL As String
On Error GoTo Error_Handler

Dim V
App_Debug = "Parametrage_Delete"

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case lFct
    Case "New": V = sqlYKYCDOS0_Insert(newYKYCDOS0)
    Case "Update", "PJ_New": V = sqlYKYCDOS0_Update(newYKYCDOS0, oldYKYCDOS0, True)
    Case "Update": V = sqlYKYCDOS0_Update(newYKYCDOS0, oldYKYCDOS0, True)
    Case "Delete": V = sqlYKYCDOS0_Delete(oldYKYCDOS0, True)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox

If lFct = "PJ_New" Then
    newYKYCDOS0.KYCDOSUFCT = "+"
    newYKYCDOS0.KYCDOSDLIB = currentPJ_FileName
    newYKYCDOS0.KYCDOSUVER = -newYKYCDOS0.KYCDOSUVER
    V = sqlYKYCDOSH_Insert(newYKYCDOS0)
    If Not IsNull(V) Then GoTo Error_MsgBox
End If
If lFct = "PJ_Delete" Then
    newYKYCDOS0 = oldYKYCDOS0
    newYKYCDOS0.KYCDOSUFCT = "-"
    newYKYCDOS0.KYCDOSDLIB = currentPJ_FileName
    newYKYCDOS0.KYCDOSUVER = -newYKYCDOS0.KYCDOSUVER
    newYKYCDOS0.KYCDOSUUSR = usrName_UCase
    newYKYCDOS0.KYCDOSUAMJ = DSys
    newYKYCDOS0.KYCDOSUHMS = time_Hms
    V = sqlYKYCDOSH_Insert(newYKYCDOS0)
    If Not IsNull(V) Then GoTo Error_MsgBox
End If
If lFct = "Delete_All" Then

    xSQL = " where KYCDOSNAT = ' ' and KYCDOSID = '" & currentYKYCDOS0.KYCDOSID & "'"
    V = sqlYKYCDOS0_Delete_Where(xSQL)
    If Not IsNull(V) Then GoTo Error_MsgBox
    newYKYCDOS0 = oldYKYCDOS0
    newYKYCDOS0.KYCDOSUFCT = "Z"
    newYKYCDOS0.KYCDOSUVER = -newYKYCDOS0.KYCDOSUVER
    newYKYCDOS0.KYCDOSUUSR = usrName_UCase
    newYKYCDOS0.KYCDOSUAMJ = DSys
    newYKYCDOS0.KYCDOSUHMS = time_Hms
    V = sqlYKYCDOSH_Insert(newYKYCDOS0)
    If Not IsNull(V) Then GoTo Error_MsgBox
End If
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    cmdParam_YKYCDOS0_Transaction = V
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function

Private Function cmdSelect_JPL()
Dim xSQL As String, Nb As Long
On Error GoTo Error_Handler

Dim V
App_Debug = "Parametrage_Delete"

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

xSQL = "Update " & paramIBM_Library_SABSPE & ".YKYCDOS0" _
     & " set KYCDOSSTAK = 'O' where KYCDOSNAT = '=' and KYCDOSSEQ in (1 , 2 , 3 , 4 , 5 , 6 , 9) and KYCDOSSEQ2 =0"
Call FEU_ROUGE
Set rsSab = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
xSQL = "Update " & paramIBM_Library_SABSPE & ".YKYCDOS0" _
     & " set KYCDOSSTAK = 'O' where KYCDOSNAT = '=' and KYCDOSSEQ2 in (66 , 68 , 69 , 73 , 74 , 75 , 77 , 78 , 80 , 81 , 83 , 86 , 87 , 88 ) "

'     & " set KYCDOSSTAK = 'O' where KYCDOSNAT = '=' and KYCDOSSEQ2 in (11 , 12 , 13 , 14 , 15 , 20 , 50 , 51 , 52 , 53 , 54 , 57 , 58 , 59 , 60 , 62 , 64 , 65 ) "
Call FEU_ROUGE
Set rsSab = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
'________________________________________________________________________________
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
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function

Private Function cmdSelect_JPL_YKYCSTA0()
Dim xSQL As String, Nb As Long
Dim rsSab_Read As New ADODB.Recordset

On Error GoTo Error_Handler

Dim V
App_Debug = "cmdSelect_JPL_YKYCSTA0e"
'______________________________________________________________________________________

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YKYCSTA0" _
    & " where KYCSTADCLO <> 0 and KYCSTASTAK <> 9" _
    & " order by KYCSTACLI , KYCSTADSIT"
    
Set rsSab_Read = cnsab.Execute(xSQL)
Do While Not rsSab_Read.EOF
    Call rsYKYCSTA0_GetBuffer(rsSab_Read, xYKYCSTA0)
    If xYKYCSTA0.KYCSTASTAK = 0 Then
        xSQL = "Update " & paramIBM_Library_SABSPE & ".YKYCSTA0" _
             & " set KYCSTASTAK = '9' where KYCSTACLI = '" & xYKYCSTA0.KYCSTACLI & "' and KYCSTADSIT = " & xYKYCSTA0.KYCSTADSIT
        Set rsSab = cnSab_Update.Execute(xSQL, Nb)
    Else
        If Mid$(xYKYCSTA0.KYCSTADSIT, 1, 6) = Mid$(xYKYCSTA0.KYCSTADCLO, 1, 6) Then
            xSQL = "Update " & paramIBM_Library_SABSPE & ".YKYCSTA0" _
                 & " set KYCSTASTAK = '9' where KYCSTACLI = '" & xYKYCSTA0.KYCSTACLI & "' and KYCSTADSIT = " & xYKYCSTA0.KYCSTADSIT
            Set rsSab = cnSab_Update.Execute(xSQL, Nb)
        Else
            If Mid$(xYKYCSTA0.KYCSTADSIT, 1, 6) > Mid$(xYKYCSTA0.KYCSTADCLO, 1, 6) Then
               ' xSQL = "Update " & paramIBM_Library_SABSPE & ".YKYCSTA0" _
               '      & " set KYCSTASTAK = 'X' where KYCSTACLI = '" & xYKYCSTA0.KYCSTACLI & "' and KYCSTADSIT = " & xYKYCSTA0.KYCSTADSIT
                xSQL = "Delete from " & paramIBM_Library_SABSPE & ".YKYCSTA0" _
                     & " where KYCSTACLI = '" & xYKYCSTA0.KYCSTACLI & "' and KYCSTADSIT = " & xYKYCSTA0.KYCSTADSIT
                Set rsSab = cnSab_Update.Execute(xSQL, Nb)
            End If
        End If
        
    End If
    rsSab_Read.MoveNext
Loop
'________________________________________________________________________________
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


Public Sub ZCLIENA0_Export()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFilex As String, wFile As String, xSQL As String
Dim wAMJMin As String, WAMJMax As String
Dim X As String, K As Long
'______________________________________________

wFile = Trim("C:\Temp\Publipostage  " & dateImp_Amj(DSys) & ".xlsx")
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut : " & wFile _
        & vbCrLf & vbCrLf & "     =========================" _
        & vbCrLf & "     =========================", "Publipostage des adresses courrier : nom du fichier d'exportation", wFile)
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
    .Title = "CLIENT"
    .Subject = ""
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "Courrier " & dateImp10(DSys)


ZCLIENA0_Export_Detail

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

Public Sub ZADRESS0_Exportation(blnActif As Boolean)
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFilex As String, wFile As String, xSQL As String
Dim wAMJMin As String, WAMJMax As String
Dim X As String, K As Long
'______________________________________________

wFile = Trim("C:\Temp\SAB adresses " & dateImp_Amj(DSys) & ".xlsx")
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut : " & wFile _
        & vbCrLf & vbCrLf & "     =========================" _
        & vbCrLf & "     =========================", "SAB exportation des adresses : nom du fichier d'exportation", wFile)
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
    .Title = "CLIENT"
    .Subject = ""
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "Adresses " & dateImp10(DSys)


ZADRESS0_Exportation_Detail blnActif

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


Public Sub ZCLIENA0_Export_Detail()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSQL As String
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, kLen As Integer

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

wsExcel.Columns(1).ColumnWidth = 7: wsExcel.Cells(1, 1) = "Client ": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 9: wsExcel.Cells(1, 2) = "Tiers de référence"
wsExcel.Columns(3).ColumnWidth = 9: wsExcel.Cells(1, 3) = "PP | PM": wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(4).ColumnWidth = 9: wsExcel.Cells(1, 4) = "Catégorie": wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(5).ColumnWidth = 60: wsExcel.Cells(1, 5) = "Nom | RS"
wsExcel.Columns(6).ColumnWidth = 32: wsExcel.Cells(1, 6) = "Adresse"
wsExcel.Columns(7).ColumnWidth = 32: wsExcel.Cells(1, 7) = "Adresse"
wsExcel.Columns(8).ColumnWidth = 32: wsExcel.Cells(1, 8) = "Adresse"
wsExcel.Columns(9).ColumnWidth = 32: wsExcel.Cells(1, 9) = "Code postal Ville"
wsExcel.Columns(10).ColumnWidth = 32: wsExcel.Cells(1, 10) = "Pays"

For K = 1 To 10
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 170, 80)
Next K

arrZCLIENA0_Max = 1000: arrZCLIENA0_NB = 0
ReDim arrZCLIENA0(arrZCLIENA0_Max)

xSQL = "select DISTINCT(CLIENACLI) from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where comptefon <> '4' and plancopro in ('CAV','LOR') order by CLIENACLI"

Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
         arrZCLIENA0_NB = arrZCLIENA0_NB + 1
         If arrZCLIENA0_NB > arrZCLIENA0_Max Then
             arrZCLIENA0_Max = arrZCLIENA0_Max + 500
             ReDim Preserve arrZCLIENA0(arrZCLIENA0_Max)
         End If
         
         arrZCLIENA0(arrZCLIENA0_NB).CLIENACLI = rsAdo(0)
    rsAdo.MoveNext

Loop
Set rsAdo = Nothing


For wRow = 1 To arrZCLIENA0_NB
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
     & " where CLIENACLI = '" & arrZCLIENA0(wRow).CLIENACLI & "'"
    Set rsAdo = cnAdo.Execute(xSQL)
    V = rsZCLIENA0_GetBuffer(rsAdo, xZCLIENA0)
    If Trim(xZCLIENA0.CLIENATIE) = "" Then
        newZCLIENA0 = xZCLIENA0
    Else
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
         & " where CLIENACLI = '" & xZCLIENA0.CLIENATIE & "'"
        Set rsAdo = cnAdo.Execute(xSQL)
        V = rsZCLIENA0_GetBuffer(rsAdo, newZCLIENA0)
    End If
    
    wsExcel.Cells(wRow + 1, 1) = xZCLIENA0.CLIENACLI
    wsExcel.Cells(wRow + 1, 2) = xZCLIENA0.CLIENATIE
    wsExcel.Cells(wRow + 1, 4) = newZCLIENA0.CLIENACAT
    
    If Trim(xZCLIENA0.CLIENARA2) = "." Then xZCLIENA0.CLIENARA2 = ""
    
    
   ' If newZCLIENA0.CLIENACAT = "PAR" Or newZCLIENA0.CLIENACAT = "PER" Then
   '    wsExcel.Cells(wRow + 1, 5) = Trim(xZCLIENA0.CLIENAETA) & " " & Trim(xZCLIENA0.CLIENARA1) & " " & Trim(xZCLIENA0.CLIENARA2)
   ' Else
   '     wsExcel.Cells(wRow + 1, 5) = Trim(xZCLIENA0.CLIENARA1) & " " & Trim(xZCLIENA0.CLIENARA2)
   ' End If
    xZADRESS0.ADRESSCOA = "CO"
    xZADRESS0.ADRESSNUM = xZCLIENA0.CLIENACLI
    Call rsZADRESS0_Client(xZADRESS0)
    wsExcel.Cells(wRow + 1, 3) = xZADRESS0.ADRESSCOA

    wsExcel.Cells(wRow + 1, 5) = Trim(xZADRESS0.ADRESSRA1) & " " & Trim(xZADRESS0.ADRESSRA2)
    wsExcel.Cells(wRow + 1, 6) = Trim(xZADRESS0.ADRESSAD1)
    wsExcel.Cells(wRow + 1, 7) = Trim(xZADRESS0.ADRESSAD2)
    wsExcel.Cells(wRow + 1, 8) = Trim(xZADRESS0.ADRESSAD3)
    If Trim(xZADRESS0.ADRESSCOP) = "." Or Trim(xZADRESS0.ADRESSCOP) = "" Then
        wsExcel.Cells(wRow + 1, 9) = Trim(xZADRESS0.ADRESSVIL)
    Else
        wsExcel.Cells(wRow + 1, 9) = Trim(xZADRESS0.ADRESSCOP) & " " & Trim(xZADRESS0.ADRESSVIL)
    End If
    If Trim(xZADRESS0.ADRESSPAY) <> "FRANCE" Then wsExcel.Cells(wRow + 1, 10) = Trim(xZADRESS0.ADRESSPAY)


Next wRow


Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub




Public Sub ZADRESS0_Exportation_Detail(blnActif As Boolean)
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSQL As String
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, kLen As Integer
Dim xWhere As String, wCellBackColor As Long, wCellBackColor_Rupture As Long, wCellBackColor_CO As Long
Dim wCLIEANCLI As String, mCLIEANCLI As String, wActif As String, mCLIEANRES As String
Dim wTITULACOM As String, mTITULACOM As String, wTITULATPR As String, wBIC As String
Dim X As String
Dim blnSelect As Boolean, blnRow_Compte As Boolean, blnRow_Adresse As Boolean, blnClient_End As Boolean
Dim blnClient_Select As Boolean
'______________________________________________

'__________________________________________________________________________________

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .WrapText = True
    .Font.Size = 8
    .Font.Name = "Arial Unicode MS"
End With

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

wsExcel.Columns(1).ColumnWidth = 7: wsExcel.Cells(1, 1) = "Client ": wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 3: wsExcel.Cells(1, 2) = "A": wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(3).ColumnWidth = 6: wsExcel.Cells(1, 3) = "Ges": wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(4).ColumnWidth = 15: wsExcel.Cells(1, 4) = "Sigle"
wsExcel.Columns(5).ColumnWidth = 34: wsExcel.Cells(1, 5) = "Nom 1"
wsExcel.Columns(6).ColumnWidth = 34: wsExcel.Cells(1, 6) = "Nom 2"
wsExcel.Columns(7).ColumnWidth = 3: wsExcel.Cells(1, 7) = "P": wsExcel.Columns(7).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(8).ColumnWidth = 14: wsExcel.Cells(1, 8) = "Compte"

wsExcel.Columns(9).ColumnWidth = 5: wsExcel.Cells(1, 9) = "COA": wsExcel.Columns(9).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(10).ColumnWidth = 34: wsExcel.Cells(1, 10) = "Nom 1"
wsExcel.Columns(11).ColumnWidth = 34: wsExcel.Cells(1, 11) = "Nom 2"
wsExcel.Columns(12).ColumnWidth = 34: wsExcel.Cells(1, 12) = "Adresse 1"
wsExcel.Columns(13).ColumnWidth = 34: wsExcel.Cells(1, 13) = "Adresse 2"
wsExcel.Columns(14).ColumnWidth = 34: wsExcel.Cells(1, 14) = "Adresse 3"
wsExcel.Columns(15).ColumnWidth = 10: wsExcel.Cells(1, 15) = "Code postal"
wsExcel.Columns(16).ColumnWidth = 25: wsExcel.Cells(1, 16) = "Ville"
wsExcel.Columns(17).ColumnWidth = 34: wsExcel.Cells(1, 17) = "Pays"

For K = 1 To 17
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 170, 80)
Next K

'____________________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Table CLIENT : "): DoEvents
arrZCLIENA0_Max = 1000: arrZCLIENA0_NB = 0
ReDim arrZCLIENA0(arrZCLIENA0_Max)

xSQL = "select DISTINCT(CLIENACLI) from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where comptefon <> '4' and plancopro in ('CAV','LOR','NOS') order by CLIENACLI"

Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
         arrZCLIENA0_NB = arrZCLIENA0_NB + 1
         If arrZCLIENA0_NB > arrZCLIENA0_Max Then
             arrZCLIENA0_Max = arrZCLIENA0_Max + 500
             ReDim Preserve arrZCLIENA0(arrZCLIENA0_Max)
         End If
         
         arrZCLIENA0(arrZCLIENA0_NB).CLIENACLI = rsAdo(0)
    rsAdo.MoveNext

Loop
Set rsAdo = Nothing
'_____________________________________________________________________________
mCLIEANCLI = ""
blnSelect = False
xWhere = ""
X = Trim(Mid$(cboSelect_CLIENACAT, 1, 3))
If X <> "" Then xWhere = xWhere & " and C.CLIENACAT = '" & X & "'"

X = Trim(txtSelect_CLIENARA1)
If X <> "" Then
    If IsNumeric(X) Then
        xWhere = xWhere & " and C.CLIENACLI like '%" & X & "%'"
    Else
        xWhere = xWhere & " and C.CLIENARA1 like '%" & X & "%'"
    End If
End If

'xSql = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 C, " _
     & paramIBM_Library_SAB & ".ZTITULA0 T, " _
     & paramIBM_Library_SABSPE & ".YBIACPT0 B, " _
     & paramIBM_Library_SAB & ".ZADRESS0 A" _
     & " where C.CLIENACLI > '0010000' and C.CLIENACLI < '0099999'" _
     & xWhere _
     & " and   C.CLIENACLI = T.TITULACLI" _
     & " and   B.COMPTECOM = T.TITULACOM" _
     & " and   B.plancopro in ('CAV','LOR')" _
     & " and   A.ADRESSTYP = 1 and   substring(A.ADRESSNUM , 2 , 7) = C.CLIENACLI" _
     & " order by C.CLIENACLI , T.TITULACOM , A.ADRESSCOA"
     
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 C, " _
     & paramIBM_Library_SAB & ".ZADRESS0 A" _
     & " where C.CLIENACLI > '0010000' and C.CLIENACLI < '0099999'" _
     & xWhere _
     & " and   A.ADRESSTYP = 1 and   substring(A.ADRESSNUM , 2 , 7) = C.CLIENACLI" _
     & " order by C.CLIENACLI , A.ADRESSCOA"

Set rsSab = cnsab.Execute(xSQL)
wRow = 0
Do While Not blnClient_End 'rsSab.EOF
    wRow = wRow + 1
    wCellBackColor_Rupture = RGB(255, 255, 255)
    
    If rsSab.EOF Then
        blnClient_End = True
        wCLIEANCLI = ""
    Else
        wCLIEANCLI = rsSab(1)
    End If
    If mCLIEANCLI <> wCLIEANCLI Then
'------------------------------------------------------------------------------
        Call lstErr_AddItem(lstErr, cmdContext, "  " & mCLIEANCLI): DoEvents

        If blnSelect Then
             blnRow_Compte = False
             xSQL = "select * from " & paramIBM_Library_SAB & ".ZTITULA0 T, " _
                  & paramIBM_Library_SABSPE & ".YBIACPT0 B " _
                  & " where T.TITULACLI = '" & mCLIEANCLI & "'" _
                  & " and   B.COMPTECOM = T.TITULACOM" _
                  & " and   B.plancopro in ('CAV','LOR','NOS')" _
                  & " order by  T.TITULACOM "
         
             Set rsAdo = cnsab.Execute(xSQL)
             Do While Not rsAdo.EOF
                If blnRow_Compte Then
                    wRow = wRow + 1
                Else
                    blnRow_Compte = True
                End If
                wsExcel.Cells(wRow + 1, 1) = mCLIEANCLI
                wsExcel.Cells(wRow + 1, 1).Interior.Color = wCellBackColor
                wsExcel.Cells(wRow + 1, 2) = wActif  'rsSab("COMPTEFON")
                wsExcel.Cells(wRow + 1, 2).Interior.Color = wCellBackColor
                wsExcel.Cells(wRow + 1, 3) = mCLIEANRES ' CLIENARES
                wsExcel.Cells(wRow + 1, 3).Interior.Color = wCellBackColor

                wTITULACOM = rsAdo(2) '"TITULACOM"
                
                If rsAdo(5) = 0 Then
                    wTITULATPR = "P"
                Else
                    wTITULATPR = ""
                End If
                wsExcel.Cells(wRow + 1, 7) = wTITULATPR  ' "TITULATPR"
                If rsAdo("COMPTEFON") = 4 Then
                   wsExcel.Cells(wRow + 1, 7).Interior.Color = RGB(220, 220, 220)
                Else
                    wsExcel.Cells(wRow + 1, 7).Interior.Color = RGB(220, 255, 220)
                End If
                wsExcel.Cells(wRow + 1, 8) = wTITULACOM
                    wsExcel.Cells(wRow + 1, 9).Interior.Color = wCellBackColor_CO
                wsExcel.Cells(wRow + 1, 7).Interior.Color = wsExcel.Cells(wRow + 1, 9).Interior.Color
                wsExcel.Cells(wRow + 1, 8).Interior.Color = wsExcel.Cells(wRow + 1, 9).Interior.Color
'------------------------------------------------------------------------------
                xSQL = "select * from " & paramIBM_Library_SAB & ".ZADRESS0" _
                     & " where ADRESSTYP = 2 and   ADRESSNUM  = '" & wTITULACOM & "'" _
                     & " order by  ADRESSCOA "
                blnRow_Adresse = False
                Set rsSabX = cnsab.Execute(xSQL)
                    Do While Not rsSabX.EOF
                    If blnRow_Adresse Then
                        wRow = wRow + 1
                    Else
                        blnRow_Adresse = True
                    End If
                    wsExcel.Cells(wRow + 1, 1) = mCLIEANCLI
                    wsExcel.Cells(wRow + 1, 1).Interior.Color = wCellBackColor
                    wsExcel.Cells(wRow + 1, 2) = wActif  'rsSab("COMPTEFON")
                    wsExcel.Cells(wRow + 1, 2).Interior.Color = wCellBackColor
                    wsExcel.Cells(wRow + 1, 3) = mCLIEANRES ' CLIENARES
                    wsExcel.Cells(wRow + 1, 3).Interior.Color = wCellBackColor
                    wsExcel.Cells(wRow + 1, 8) = wTITULACOM
                    wsExcel.Cells(wRow + 1, 7) = wTITULATPR
                    X = Trim(rsSabX("ADRESSCOA"))
                    wsExcel.Cells(wRow + 1, 9) = X
                    Select Case X
                        Case "CO": wsExcel.Cells(wRow + 1, 9).Interior.Color = RGB(255, 170, 170) 'RGB(255, 160, 64)
                        Case "CH": wsExcel.Cells(wRow + 1, 9).Interior.Color = vbMagenta
                        Case "": wsExcel.Cells(wRow + 1, 9).Interior.Color = vbRed
                    End Select
                    wsExcel.Cells(wRow + 1, 8).Interior.Color = wsExcel.Cells(wRow + 1, 9).Interior.Color
                    wsExcel.Cells(wRow + 1, 7).Interior.Color = wsExcel.Cells(wRow + 1, 9).Interior.Color

                    X = Trim(rsSabX("ADRESSRA11") & rsSabX("ADRESSRA12") & rsSabX("ADRESSRA13"))
                    If X <> "" Then
                        wsExcel.Cells(wRow + 1, 10) = X
                        wsExcel.Cells(wRow + 1, 10).Interior.Color = RGB(255, 220, 220)
                    End If
                    wsExcel.Cells(wRow + 1, 11) = Trim(rsSabX("ADRESSRA2"))
                    wsExcel.Cells(wRow + 1, 12) = Trim(rsSabX("ADRESSAD1"))
                    wsExcel.Cells(wRow + 1, 13) = Trim(rsSabX("ADRESSAD2"))
                    wsExcel.Cells(wRow + 1, 14) = Trim(rsSabX("ADRESSAD3"))
                    wsExcel.Cells(wRow + 1, 15) = Trim(rsSabX("ADRESSCOP"))
                    wsExcel.Cells(wRow + 1, 16) = Trim(rsSabX("ADRESSVIL"))
                    wsExcel.Cells(wRow + 1, 17) = Trim(rsSabX("ADRESSPAY"))
                    rsSabX.MoveNext
                Loop
                
'------------------------------------------------------------------------------
                 rsAdo.MoveNext
             Loop
             If blnRow_Compte Then wRow = wRow + 1

        End If
        If blnClient_End Then Exit Sub
        
'------------------------------------------------------------------------------
        mCLIEANCLI = wCLIEANCLI
        wCellBackColor = RGB(230, 230, 230)
        wCellBackColor_Rupture = wCellBackColor
        wCellBackColor_CO = wCellBackColor
        wActif = "X"
        blnSelect = blnActif
        For K1 = 1 To arrZCLIENA0_NB
            If wCLIEANCLI = arrZCLIENA0(K1).CLIENACLI Then
                wCellBackColor = RGB(180, 255, 180)
                wCellBackColor_CO = wCellBackColor
                wActif = "": blnSelect = True
                Exit For
            End If
        Next K1
        
        mCLIEANRES = rsSab(13)
        If wActif = "X" Then
            If Mid$(mCLIEANRES, 1, 1) = "R" Then
                wCellBackColor = RGB(160, 192, 255)
                wCellBackColor_CO = wCellBackColor
                wActif = "-": blnSelect = True
            End If

        End If
        If blnSelect Then
            wsExcel.Cells(wRow + 1, 4) = rsSab(6) ' CLIENASIG
            wsExcel.Cells(wRow + 1, 5) = rsSab(4) '"CLIENARA1"
            wsExcel.Cells(wRow + 1, 6) = rsSab(5) '"CLIENARA2"
            wsExcel.Cells(wRow + 1, 7) = rsSab(12) ' CLIENARSD
            wBIC = ""
            xSQL = "select ADRESSRA12 from " & paramIBM_Library_SAB & ".ZADRESS0" _
                 & " where ADRESSTYP = 4 and   ADRESSNUM  = ' " & mCLIEANCLI & "'"
            Set rsSabX = cnsab.Execute(xSQL)
            If Not rsSabX.EOF Then wBIC = rsSabX("ADRESSRA12")
            

        End If

    End If
    If Not blnSelect Then
        wRow = wRow - 1
    Else
        wsExcel.Cells(wRow + 1, 1) = wCLIEANCLI
        wsExcel.Cells(wRow + 1, 1).Interior.Color = wCellBackColor
        wsExcel.Cells(wRow + 1, 2) = wActif  'rsSab("COMPTEFON")
        wsExcel.Cells(wRow + 1, 2).Interior.Color = wCellBackColor
        wsExcel.Cells(wRow + 1, 3) = mCLIEANRES ' CLIENARES
        wsExcel.Cells(wRow + 1, 3).Interior.Color = wCellBackColor
        
        X = Trim(rsSab("ADRESSCOA"))
        wsExcel.Cells(wRow + 1, 9) = X
        Select Case X
            Case "": wsExcel.Cells(wRow + 1, 9).Interior.Color = wCellBackColor: wCellBackColor_CO = RGB(220, 255, 220)
            Case "CO": wsExcel.Cells(wRow + 1, 9).Interior.Color = RGB(255, 220, 160): wCellBackColor_CO = wsExcel.Cells(wRow + 1, 9).Interior.Color
            Case Else: wsExcel.Cells(wRow + 1, 9).Interior.Color = vbRed
        End Select
        wsExcel.Cells(wRow + 1, 4).Interior.Color = wsExcel.Cells(wRow + 1, 9).Interior.Color
        wsExcel.Cells(wRow + 1, 5).Interior.Color = wsExcel.Cells(wRow + 1, 9).Interior.Color
        wsExcel.Cells(wRow + 1, 6).Interior.Color = wsExcel.Cells(wRow + 1, 9).Interior.Color
        wsExcel.Cells(wRow + 1, 8).Interior.Color = wsExcel.Cells(wRow + 1, 9).Interior.Color 'wCellBackColor
        wsExcel.Cells(wRow + 1, 7).Interior.Color = wsExcel.Cells(wRow + 1, 9).Interior.Color ' wCellBackColor
        
        If X = "" And wBIC <> "" Then
            wsExcel.Cells(wRow + 1, 8) = wBIC
            wsExcel.Cells(wRow + 1, 8).Interior.Color = RGB(96, 255, 96)
        End If
        
        X = Trim(rsSab("ADRESSRA11") & rsSab("ADRESSRA12") & rsSab("ADRESSRA13"))
        If X <> "" Then
            wsExcel.Cells(wRow + 1, 10) = X
            wsExcel.Cells(wRow + 1, 10).Interior.Color = RGB(255, 200, 200)
        End If
        wsExcel.Cells(wRow + 1, 11) = Trim(rsSab("ADRESSRA2"))
        wsExcel.Cells(wRow + 1, 12) = Trim(rsSab("ADRESSAD1"))
        wsExcel.Cells(wRow + 1, 13) = Trim(rsSab("ADRESSAD2"))
        wsExcel.Cells(wRow + 1, 14) = Trim(rsSab("ADRESSAD3"))
        wsExcel.Cells(wRow + 1, 15) = Trim(rsSab("ADRESSCOP"))
        wsExcel.Cells(wRow + 1, 16) = Trim(rsSab("ADRESSVIL"))
        wsExcel.Cells(wRow + 1, 17) = Trim(rsSab("ADRESSPAY"))
    End If
    
    rsSab.MoveNext
Loop
Set rsSab = Nothing



Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

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



Private Sub cboSelect_CLIENACAT_GotFocus()
cmdSelect_Clear
txt_GotFocus cboSelect_CLIENACAT

End Sub

Private Sub cboSelect_CLIENACAT_LostFocus()
txt_LostFocus cboSelect_CLIENACAT

End Sub

Private Sub cboSelect_Options_4_CLIENARES_Click()
cmdSelect_Clear

End Sub


Private Sub cboSelect_Options_4_Code_Click()
cmdSelect_Clear

End Sub

Private Sub cboSelect_Options_KYCgsop_CLIENARES_Click()
cmdSelect_Clear

End Sub


Private Sub cboSelect_Options_Xgsop_CLIENARES_Click()
cmdSelect_Clear
If cboSelect_Options_Xgsop_CLIENARES.Text = "Archives" Then
    cboSelect_Options_Xgsop_Archive.Visible = True
    cboSelect_Options_Xgsop_Archive.ListIndex = 0
Else
    cboSelect_Options_Xgsop_Archive.Visible = False
End If

End Sub


Private Sub cboSelect_SQL_Click()

cmdSelect_Reset
End Sub


Private Sub cboSelect_SQL_GotFocus()
cboSelect_SQL.ForeColor = vbBlue
'txt_GotFocus cboSelect_SQL
End Sub


Private Sub cboSelect_SQL_LostFocus()
'txt_LostFocus cboSelect_SQL
cboSelect_SQL.ForeColor = vbBlack

End Sub


Private Sub cboUpdate_Add_GotFocus()
txt_GotFocus cboUpdate_Add

End Sub


Private Sub cboUpdate_Add_LostFocus()
txt_LostFocus cboUpdate_Add

End Sub


Private Sub cboUpdate_ADRESSPAY_GotFocus()
txt_GotFocus cboUpdate_ADRESSPAY

End Sub


Private Sub cboUpdate_ADRESSPAY_LostFocus()
txt_LostFocus cboUpdate_ADRESSPAY

End Sub


Private Sub cboUpdate_CLIENAETA_Click()
If Mid$(cboUpdate_CLIENAETA, 1, 3) = "MME" Then
    txtUpdate_CLIENAFIL.Visible = True
Else
    txtUpdate_CLIENAFIL.Visible = False
End If

End Sub

Private Sub cboUpdate_CLIENAETA_GotFocus()
txt_GotFocus cboUpdate_CLIENAETA

End Sub


Private Sub cboUpdate_CLIENAETA_LostFocus()
txt_LostFocus cboUpdate_CLIENAETA

End Sub


Private Sub cboUpdate_CLIENANAT_GotFocus()
txt_GotFocus cboUpdate_CLIENANAT

End Sub


Private Sub cboUpdate_CLIENANAT_LostFocus()
txt_LostFocus cboUpdate_CLIENANAT

End Sub


Private Sub cboUpdate_CLIENARSD_GotFocus()
txt_GotFocus cboUpdate_CLIENARSD

End Sub


Private Sub cboUpdate_CLIENARSD_LostFocus()
txt_LostFocus cboUpdate_CLIENARSD

End Sub


Private Sub cboUpdate_CLIENBLIE_GotFocus()
txt_GotFocus cboUpdate_CLIENBLIE

End Sub


Private Sub cboUpdate_CLIENBLIE_LostFocus()
txt_LostFocus cboUpdate_CLIENBLIE

End Sub


Private Sub cboUpdate_CLIENBNAS_GotFocus()
txt_GotFocus cboUpdate_CLIENBNAS

End Sub


Private Sub cboUpdate_CLIENBNAS_LostFocus()
txt_LostFocus cboUpdate_CLIENBNAS

End Sub


Private Sub cboUpdate_CLIENBTER_GotFocus()
txt_GotFocus cboUpdate_CLIENBTER

End Sub


Private Sub cboUpdate_CLIENBTER_LostFocus()
txt_LostFocus cboUpdate_CLIENBTER

End Sub


Private Sub chkSelect_Options_3_Click()
cmdSelect_Clear

End Sub

Private Sub chkSelect_Options_4_AUTSICMON_Click()
cmdSelect_Clear

End Sub


Private Sub chkSelect_Options_4_ECHTABDON_S_Click()
cmdSelect_Clear

End Sub


Private Sub chkUpdLog_AmjMin_Click()
fgSelect.Visible = False
If chkUpdLog_AmjMin.Value = "1" Then
    txtUpdLog_AmjMin.Visible = True
Else
    txtUpdLog_AmjMin.Visible = False
End If

End Sub

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdParam_YKYCDOS0_4c_Actualisation_Click()
Dim V, X As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

Dim rsSab_Local As New ADODB.Recordset

X = "select * from " & paramIBM_Library_SABSPE & ".YKYCDOS0 where KYCDOSNAT = ''" _
        & " and KYCDOSDLIB = '" & oldYKYCDOS0.KYCDOSID & "' and KYCDOSSEQ  = 0 and  KYCDOSSEQ2  = 0 "

Set rsSab_Local = cnsab.Execute(X)

Do While Not rsSab_Local.EOF
    X = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & rsSab_Local("KYCDOSID") & "'"
    
    Set rsSab = cnsab.Execute(X)
    Call rsZCLIENA0_GetBuffer(rsSab, selZCLIENA0)
    cmdSelect_SQL_YKYCDOS0_Init

    rsSab_Local.MoveNext
Loop
cmdParam_YKYCDOS0_4c_Actualisation.Visible = False
fraYKYCDOS0.Visible = False
lstParam_YKYCDOS0.Visible = True
SSTab1.Tab = 3
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_YKYCDOS0_4c_Quit_Click()
fraParam_YKYCDOS0_4c.Visible = False
End Sub

Private Sub cmdParam_YKYCDOS0_4c_Update_Click()
Dim V, X As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

newYKYCDOS0 = oldYKYCDOS0
blnOk = True

newYKYCDOS0.KYCDOSDLIB = Trim(txtParam_YKYCDOS0_4c)


If chkParam_YKYCDOS0_4c.Value <> "1" Then
    newYKYCDOS0.KYCDOSSTAK = " "
Else
    newYKYCDOS0.KYCDOSSTAK = "O"
End If

If blnOk Then
    newYKYCDOS0.KYCDOSUFCT = "U"
    V = cmdParam_YKYCDOS0_Transaction("Update")
    If IsNull(V) Then
        fgParam_YKYCDOS0_4c.Visible = False
        fraParam_YKYCDOS0_4c.Visible = False
        cmdParam_YKYCDOS0_4c_Actualisation.Visible = False
        lstParam_YKYCDOS0_4c_Load
        fgParam_YKYCDOS0_4c.Visible = True
        If oldYKYCDOS0.KYCDOSSTAK <> newYKYCDOS0.KYCDOSSTAK And mParam_YKYCDOS0_4c_Actualisation_Nb > 0 Then
            cmdParam_YKYCDOS0_4c_Actualisation.Visible = True
        End If
    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_YKYCDOS0_Add_Click()
Dim V, X As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

newYKYCDOS0 = oldYKYCDOS0
blnOk = True
X = Trim(txtParam_KYCDOSSEQ)
If X = "" Then
    blnOk = False
    Call MsgBox("Préciser l'identifiant", vbCritical, "SAB_Client: paramétrage")
Else
    Select Case mParam_KYCDOSNAT
        Case "D", "J":
                newYKYCDOS0.KYCDOSSEQ = Val(X)
                If newYKYCDOS0.KYCDOSSEQ < 1 Or newYKYCDOS0.KYCDOSSEQ > 9999 Then
                    blnOk = False
                    Call MsgBox("Préciser l'identifiant entre 1 et 999", vbCritical, "SAB_Client: paramétrage")
                End If
        Case Else: newYKYCDOS0.KYCDOSID = X
    End Select
End If

X = Trim(txtParam_KYCDOSDLIB)
If X = "" Then
    blnOk = False
    Call MsgBox("Préciser le libellé", vbCritical, "SAB_Client: paramétrage")
Else
    newYKYCDOS0.KYCDOSDLIB = X
End If


If chkParam_KYCDOSSTAK.Value <> "1" Then
    newYKYCDOS0.KYCDOSSTAK = " "
Else
    Select Case mParam_KYCDOSNAT
        Case "D": newYKYCDOS0.KYCDOSSTAK = "O"
        Case "J": newYKYCDOS0.KYCDOSSTAK = "O"
        Case Else: newYKYCDOS0.KYCDOSSTAK = " "
    End Select
End If
If mParam_KYCDOSNAT = "D" Then newYKYCDOS0.KYCDOSDECH = Val(txtParam_KYCDOSDECH)

If blnOk Then
    newYKYCDOS0.KYCDOSUFCT = "A"
    V = cmdParam_YKYCDOS0_Transaction("New")
    If IsNull(V) Then
        fraParam_YKYCDOS0_JD.Visible = False
        fraParam_YKYCDOS0_Update.Visible = False
        lstParam_YKYCDOS0_Load
    End If
End If

Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdParam_YKYCDOS0_Delete_Click()
Dim V, X As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

newYKYCDOS0 = oldYKYCDOS0
X = Trim(txtParam_KYCDOSSEQ)
Select Case mParam_KYCDOSNAT
    Case "D", "J":
            newYKYCDOS0.KYCDOSSEQ = Val(X)
    Case Else: newYKYCDOS0.KYCDOSID = X
End Select
If Trim(newYKYCDOS0.KYCDOSID) = Trim(oldYKYCDOS0.KYCDOSID) _
And newYKYCDOS0.KYCDOSSEQ = oldYKYCDOS0.KYCDOSSEQ _
And newYKYCDOS0.KYCDOSSEQ2 = oldYKYCDOS0.KYCDOSSEQ2 Then
    blnOk = True
Else
    blnOk = False
    Call MsgBox("l'identifiant a été modifié : suppression impossible", vbCritical, "SAB_Client: paramétrage")
End If
Call MsgBox("Contrôles à faire", vbInformation, "cmdParam_YKYCDOS0_Delete")

If blnOk Then
    oldYKYCDOS0.KYCDOSUFCT = "D"
    V = cmdParam_YKYCDOS0_Transaction("Delete")
    If IsNull(V) Then
        fraParam_YKYCDOS0_JD.Visible = False
        fraParam_YKYCDOS0_Update.Visible = False
        lstParam_YKYCDOS0_Load
    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdParam_YKYCDOS0_JD_Quit_Click()
If blnYKYCDOS0_JD Then
    If MsgBox("Les modifications ne sont pas sauvegardées, confirmez-vous l'abandon ?", vbYesNo, "Paramétrage KYC") = vbYes Then
        fraParam_YKYCDOS0_JD.Visible = False
    End If
Else
    fraParam_YKYCDOS0_JD.Visible = False
End If
End Sub

Private Sub cmdParam_YKYCDOS0_JD_Update_Click()
Dim V, K As Long
Me.Enabled = False: Me.MousePointer = vbHourglass

    

On Error GoTo Error_Handler

App_Debug = "Parametrage_Delete"

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

Call rsYKYCDOS0_Init(newYKYCDOS0)
newYKYCDOS0.KYCDOSNAT = "="

newYKYCDOS0.KYCDOSID = oldYKYCDOS0.KYCDOSID

For K = 1 To 999
    If oldParam_J(K) <> newParam_J(K) Then
        newYKYCDOS0.KYCDOSSEQ = K: newYKYCDOS0.KYCDOSSEQ2 = 0

        If oldParam_J(K) = 0 Then
            newYKYCDOS0.KYCDOSSTAK = arrKYCDOSSTAK_J(K)
            newYKYCDOS0.KYCDOSUFCT = "A": newYKYCDOS0.KYCDOSUVER = 0
            V = sqlYKYCDOS0_Insert(newYKYCDOS0)
        Else
            If newParam_J(K) = 0 Then
                newYKYCDOS0.KYCDOSUFCT = "D": newYKYCDOS0.KYCDOSUVER = 999
                V = sqlYKYCDOS0_Delete(newYKYCDOS0, False)
            End If
        End If
        If Not IsNull(V) Then GoTo Error_MsgBox
    End If
Next K
For K = 1 To 999
    If oldParam_D(K) <> newParam_D(K) Then
        If oldParam_D(K) = 0 Then
            newYKYCDOS0.KYCDOSSEQ = newParam_D(K): newYKYCDOS0.KYCDOSSEQ2 = K
            newYKYCDOS0.KYCDOSSTAK = arrKYCDOSSTAK_D(K)
            newYKYCDOS0.KYCDOSUFCT = "A": newYKYCDOS0.KYCDOSUVER = 0
            V = sqlYKYCDOS0_Insert(newYKYCDOS0)
        Else
            If newParam_J(K) = 0 Then
                newYKYCDOS0.KYCDOSSEQ = oldParam_D(K): newYKYCDOS0.KYCDOSSEQ2 = K
                newYKYCDOS0.KYCDOSUFCT = "D": newYKYCDOS0.KYCDOSUVER = 999
                V = sqlYKYCDOS0_Delete(newYKYCDOS0, False)
            Else
                newYKYCDOS0.KYCDOSSEQ = oldParam_D(K): newYKYCDOS0.KYCDOSSEQ2 = K
                newYKYCDOS0.KYCDOSUFCT = "D" ' newYKYCDOS0.KYCDOSUVER = 999
                V = sqlYKYCDOS0_Delete(newYKYCDOS0, False)
                If Not IsNull(V) Then GoTo Error_MsgBox
                newYKYCDOS0.KYCDOSSEQ = newParam_D(K): newYKYCDOS0.KYCDOSSEQ2 = K
                newYKYCDOS0.KYCDOSSTAK = arrKYCDOSSTAK_D(K)
                newYKYCDOS0.KYCDOSUFCT = "A": newYKYCDOS0.KYCDOSUVER = 0
                V = sqlYKYCDOS0_Insert(newYKYCDOS0)
           End If
        End If
        If Not IsNull(V) Then GoTo Error_MsgBox
    End If
Next K

'________________________________________________________________________________

fraParam_YKYCDOS0_JD.Visible = False

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
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_YKYCDOS0_Quit_Click()
fraParam_YKYCDOS0_Update.Visible = False
fraParam_YKYCDOS0_JD.Visible = False

End Sub

Private Sub cmdParam_YKYCDOS0_Update_Click()
Dim V, X As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

newYKYCDOS0 = oldYKYCDOS0
X = Trim(txtParam_KYCDOSSEQ)
Select Case mParam_KYCDOSNAT
    Case "D", "J":
            newYKYCDOS0.KYCDOSSEQ = Val(X)
    Case Else: newYKYCDOS0.KYCDOSID = X
End Select
If newYKYCDOS0.KYCDOSID = oldYKYCDOS0.KYCDOSID _
And newYKYCDOS0.KYCDOSSEQ = oldYKYCDOS0.KYCDOSSEQ _
And newYKYCDOS0.KYCDOSSEQ2 = oldYKYCDOS0.KYCDOSSEQ2 Then
    blnOk = True
Else
    blnOk = False
    Call MsgBox("l'identifiant a été modifié : modification impossible", vbCritical, "SAB_Client: paramétrage")
End If

X = Trim(txtParam_KYCDOSDLIB)
If X = "" Then
    blnOk = False
    Call MsgBox("Préciser le libellé", vbCritical, "SAB_Client: paramétrage")
Else
    newYKYCDOS0.KYCDOSDLIB = X
End If
If chkParam_KYCDOSSTAK.Value <> "1" Then
    newYKYCDOS0.KYCDOSSTAK = ""
Else
    Select Case mParam_KYCDOSNAT
        Case "D": newYKYCDOS0.KYCDOSSTAK = "O"
        Case "J": newYKYCDOS0.KYCDOSSTAK = "O"
        Case Else: newYKYCDOS0.KYCDOSSTAK = ""
    End Select
End If

If mParam_KYCDOSNAT = "D" Then newYKYCDOS0.KYCDOSDECH = Val(txtParam_KYCDOSDECH)

If blnOk Then
    newYKYCDOS0.KYCDOSUFCT = "U"
    V = cmdParam_YKYCDOS0_Transaction("Update")
    If IsNull(V) Then
        fraParam_YKYCDOS0_JD.Visible = False
        fraParam_YKYCDOS0_Update.Visible = False
        lstParam_YKYCDOS0_Load
    End If
End If

Me.Enabled = True: Me.MousePointer = 0



End Sub

Private Sub cmdPJ_Delete_Click()
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass

If Dir(currentPJ_Path_FileName) <> "" Then
    msFileSystem.DeleteFile currentPJ_Path_FileName
    V = cmdParam_YKYCDOS0_Transaction("PJ_Delete")

End If

cmdPJ_Delete.Visible = False
fgPJ_Display

GoTo Exit_sub

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, "cmdPJ_Delete_Click :" & currentPJ_Path_FileName
Exit_sub:

Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdPJ_OK_Click()
Dim Archive_Folder As String, Archive_File As String
Dim App_Event As String, wFile_Id As String
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass

If Len(rtfPJ.Text) > 0 Then
    oldFileName = "C:\temp\GSOP_" & DSYS_Time & ".rtf"
    If Dir(oldFileName) <> "" Then Kill oldFileName
    newDirPath = paramGSOP_Dossier_Path & oldYKYCDOS0.KYCDOSID
    'newFileName = "GSOP_" & DSYS_Time & ".rtf"
    newFileExtension = "rtf"
    rtfPJ.SaveFile oldFileName
End If

mExe_Sequence = mExe_Sequence + 1
wFile_Id = oldYKYCDOS0.KYCDOSID & "_" & oldYKYCDOS0.KYCDOSSEQ2 & "_" & DSYS_Time & mExe_Sequence

newFileName = newDirPath & "\" & wFile_Id & "." & newFileExtension
    
App_Event = "MkDir " & newDirPath

If Not msFileSystem.FolderExists(newDirPath) Then MkDir newDirPath
App_Event = "CopyFile " & oldFileName & vbCrLf & newFileName
If Dir(newFileName) <> "" Then Kill newFileName

msFileSystem.CopyFile oldFileName, newFileName

currentPJ_Path_FileName = newFileName
currentPJ_FileName = wFile_Id & "." & newFileExtension

fraPJ.Visible = False

    newYKYCDOS0.KYCDOSPJ = "*"
    newYKYCDOS0.KYCDOSUFCT = "J"
    V = cmdParam_YKYCDOS0_Transaction("PJ_New")
    If IsNull(V) Then
        fraYKYCDOS0_Update.Visible = False
        cmdSelect_SQL_YKYCDOS0_Init
    End If

GoTo Exit_sub

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V & vbCrLf & App_Event, vbCritical, frmElp_Caption & "cmdYGOSEVE0_Update_PJ"
Exit_sub:

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdPJ_Path_Click()
Dim xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass

mfilDoc_Path = filDoc.PATH

New_YBIATAB0.BIATABID = "GSOP_PJ**"
New_YBIATAB0.BIATABK1 = usrName_UCase 'currentService_Code
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = mfilDoc_Path
If Not blnfilDoc_Path Then
    blnfilDoc_Path = True
    Parametrage_New
Else
    
    Old_YBIATAB0 = New_YBIATAB0
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
        & "where BIATABID = '" & New_YBIATAB0.BIATABID & "' and BIATABK1 = '" & New_YBIATAB0.BIATABK1 & "' and BIATABK2 = '" & New_YBIATAB0.BIATABK2 & "'"
    
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        Old_YBIATAB0.BIATABTXT = rsSab("BIATABTXT")
        Parametrage_Update
    End If
            
End If
cmdPJ_Path.Visible = False
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdPJ_Quit_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
fraPJ.Visible = False
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdPJ_réseau_Click()
dirListBox.PATH = "\\docsrv2013\_scan\"
End Sub

Private Sub cmdPrint_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 0:
        If mSelect_SQL = "KYC gsop" Or mSelect_SQL = "KYC ech" Or mSelect_SQL = "KYC Releve" Then
            Me.PopupMenu mnuPrint2, vbPopupMenuLeftButton
        Else

            If mSelect_SQL = "4" Then
                fgDetail_4_Exportation
            Else
                If mSelect_SQL = "4!e" Or mSelect_SQL = "5" Then
                    fgDetail_5_Exportation
                Else
                    If Mid$(mSelect_SQL, 1, 2) = "1r" Then fgDetail_1r_Exportation
                 'If fgSelect.Rows > 1 Then
                    ' Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
                End If
            End If
        End If
                    
    Case 1:
        If blnSelect_YKYCDOS0 Then
            Me.PopupMenu mnuPrint2, vbPopupMenuLeftButton
        Else
            Me.PopupMenu mnuPrint1, vbPopupMenuLeftButton
        End If
    Case 2: Me.PopupMenu mnuPrint2, vbPopupMenuLeftButton
    Case 3:
            Select Case ssTab_Param.Tab
                Case 0
                Case 1: cmdParam_YKYCDOS0_Print
            End Select
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL_1()
Dim I As Integer, X As String
Dim wCli As String
Dim wCLIENACLI As String
Dim V
Dim xWhere As String, xAnd As String
Dim xCLIGRPREG As String
On Error GoTo Error_Handler

arrZCLIGRP0_Index = 1
'====================================================================================
xCLIGRPREG = Trim(txtSelect_CLIENARA1)
If IsNumeric(xCLIGRPREG) Then xCLIGRPREG = Format$(xCLIGRPREG, "0000000")
If xCLIGRPREG >= "0006000" And xCLIGRPREG <= "0009999" Then
    xWhere = ""
     X = "select CLIGRPCLI from " & paramIBM_Library_SAB & ".ZCLIGRP0" _
    & " where CLIGRPREG = '" & xCLIGRPREG & "'" _
    & "  order by CLIGRPCLI"
     Set rsSab = cnsab.Execute(X)
     
     Do While Not rsSab.EOF
         xWhere = xWhere & ",'" & rsSab("CLIGRPCLI") & "'"
         rsSab.MoveNext
     Loop
     If xWhere = "" Then
         Call MsgBox("Il n'y a pas de racines rattachées à ce groupe ", vbCritical, "cmdSelect_SQL_Exportation_Liste")
         Exit Sub
     Else
         Mid$(xWhere, 1, 1) = " "
         xWhere = "where CLIENACLI in (" & xWhere & ")"
     End If
Else
    xWhere = " where CLIENAAGE = 1 and CLIENACLI > '0010000' and CLIENACLI < '0090000' "
    xAnd = " and "
    X = Trim(Mid$(cboSelect_CLIENACAT, 1, 3))
    If X <> "" Then xWhere = xWhere & xAnd & "CLIENACAT = '" & X & "'": xAnd = " and "
    
    X = Trim(txtSelect_CLIENARA1)
    If X <> "" Then
        If IsNumeric(X) Then
            xWhere = xWhere & xAnd & "CLIENACLI like '%" & X & "%'": xAnd = " and "
        Else
            xWhere = xWhere & xAnd & "CLIENARA1 like '%" & X & "%'": xAnd = " and "
        End If
    End If
End If
'====================================================================================


tvwSelect.Nodes.Clear
tvwInverse.Nodes.Clear
lblSelect = ""
lblInverse = ""
Set rsAdo = Nothing

X = "select CLIENACLI,CLIENARA1,CLIENARA2 from " & paramIBM_Library_SAB & ".ZCLIENA0 " & xWhere & " order by CLIENACLI"


Set rsAdo = cnAdo.Execute(X)

Do While Not rsAdo.EOF
    wCLIENACLI = rsAdo("CLIENACLI")
    wCli = "CLI" & wCLIENACLI
    tvwSelect.Nodes.Add , , wCli, wCLIENACLI & "   " & Trim(rsAdo("CLIENARA1")) & "   " & Trim(rsAdo("CLIENARA2"))
    tvwSelect.Nodes(wCli).Sorted = True
    'tvwSelect_Display_ZCLIGRP0 wCLIENACLI

    rsAdo.MoveNext

Loop

fraSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub cmdSelect_SQL_1r()
Dim I As Integer, xCAT As String, xRA1 As String
Dim wCli As String
Dim wCLIENACLI As String
Dim V
Dim xWhere As String, xAnd As String
Dim xCLIGRPREG As String
On Error GoTo Error_Handler

arrZCLIGRP0_Index = 1
'====================================================================================
xWhere = " where CLIENAAGE = 1 and substring(CLIENARES, 1 , 1) <> 'X' "
xAnd = " and "
xRA1 = Trim(txtSelect_CLIENARA1)

If chkSelect_Groupes = "1" Then
'_________________________________________________________________________________________________________
    xWhere = xWhere & xAnd & "CLIENACLI < '0010000'": xAnd = " and "
    If xRA1 = "" Then
    Else
        If IsNumeric(xRA1) Then
            xWhere = xWhere & xAnd & "CLIENACLI like '%" & xRA1 & "%'": xAnd = " and "
        Else
            xWhere = xWhere & xAnd & "CLIENARA1 like '%" & xRA1 & "%' ": xAnd = " and "
        End If
    End If
Else
'_________________________________________________________________________________________________________
    xCAT = Trim(Mid$(cboSelect_CLIENACAT, 1, 3))
    If xCAT <> "" Then xWhere = xWhere & xAnd & "CLIENACAT = '" & xCAT & "' and CLIENACLI > '0010000' and CLIENACLI < '0090000' ": xAnd = " and "

    If xRA1 = "" Then
        If xCAT = "" Then xWhere = xWhere & " and CLIENACLI > '0010000'"
    Else
        If IsNumeric(xRA1) Then
            xWhere = xWhere & xAnd & "CLIENACLI like '%" & xRA1 & "%'": xAnd = " and "   ' and CLIENACLI < '0090000'
        Else
            xWhere = xWhere & xAnd & "CLIENARA1 like '%" & xRA1 & "%' and CLIENACLI > '0010000' ": xAnd = " and "   ' and CLIENACLI < '0090000'
        End If
    End If
End If

'====================================================================================
arrZCLIENA0_sql xWhere

fgDetail_1r_Display

fraSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_3()
Dim I As Integer, X As String
Dim wCli As String
Dim wCLIENACLI As String
Dim V
Dim blnLoro_Actif As Boolean, bln91120_Actif As Boolean
Dim mCOMPTEOBL As String
On Error GoTo Error_Handler

arrZCLIGRP0_Index = 1


tvwSelect.Nodes.Clear
tvwInverse.Nodes.Clear
lblSelect = ""
lblInverse = ""
Set rsAdo = Nothing
If optSelect_Options_3C Then
    mCOMPTEOBL = "91120"
    X = "('91120','12120')"
Else
    mCOMPTEOBL = "98050"
    X = "('98050','12120')"
End If

'X = "select CLIENACLI,CLIENARA1,CLIENARA2 from " & paramIBM_Library_SAB & ".ZCLIENA0 " & xWhere & " order by CLIENACLI"
X = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
    & " where COMPTEFON <> '4' and substr(COMPTEOBL,1,5) in " & X _
    & " order by CLIENACLI,COMPTEOBL"

Call rsYBIACPT0_Init(oldYBIACPT0)
blnLoro_Actif = False
bln91120_Actif = False

Set rsAdo = cnAdo.Execute(X)

Do While Not rsAdo.EOF
    Call rsYBIACPT0_GetBuffer(rsAdo, xYBIACPT0)
    
    If xYBIACPT0.CLIENACLI <> oldYBIACPT0.CLIENACLI Then
        If bln91120_Actif And Not blnLoro_Actif Then
            wCLIENACLI = oldYBIACPT0.CLIENACLI
            wCli = "CLI" & wCLIENACLI
            tvwSelect.Nodes.Add , , wCli, wCLIENACLI & "   " & Trim(oldYBIACPT0.CLIENARA1) & "   " & Trim(oldYBIACPT0.CLIENARA2)
            tvwSelect.Nodes(wCli).Sorted = True
        End If
        oldYBIACPT0 = xYBIACPT0
        blnLoro_Actif = False
        bln91120_Actif = False
    End If
    
    
    If Mid$(xYBIACPT0.COMPTEOBL, 1, 5) = mCOMPTEOBL Then
        If xYBIACPT0.SOLDECEN <> 0 Then
            bln91120_Actif = True
        Else
            If chkSelect_Options_3 = "1" Then bln91120_Actif = True
        End If
    End If
    If Mid$(xYBIACPT0.COMPTEOBL, 1, 5) = "12120" Then blnLoro_Actif = True
    rsAdo.MoveNext

Loop

If bln91120_Actif And Not blnLoro_Actif Then
    wCLIENACLI = oldYBIACPT0.CLIENACLI
    wCli = "CLI" & wCLIENACLI
    tvwSelect.Nodes.Add , , wCli, wCLIENACLI & "   " & Trim(oldYBIACPT0.CLIENARA1) & "   " & Trim(oldYBIACPT0.CLIENARA2)
    tvwSelect.Nodes(wCli).Sorted = True
End If

fraSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_4()
Dim I As Integer, X As String
Dim wCli As String
Dim mECHISBCOM As String
Dim V
On Error GoTo Error_Handler


tvwSelect.Nodes.Clear
tvwInverse.Nodes.Clear
lblSelect = ""
lblInverse = ""
Set rsAdo = Nothing

X = "select count(DISTINCT  ECHISBCOM) from " & paramIBM_Library_SAB & ".ZECHISB0 " _
    & " where ECHISBFIN =  " & Val(cboSelect_Options_4_ECHISBFIN) - 19000000 & " and ECHISBCAL = 'F' and ECHISBPER = 'N'"

Set rsAdo = cnAdo.Execute(X)
ReDim arrWECHISB0(rsAdo(0) + 1)
arrWECHISB0_nb = 0

X = "select * from " & paramIBM_Library_SAB & ".ZECHISB0 " _
    & " where ECHISBFIN = " & Val(cboSelect_Options_4_ECHISBFIN) - 19000000 & " and ECHISBCAL = 'F' and ECHISBPER = 'N'" _
    & " order by ECHISBCOM"

Set rsAdo = cnAdo.Execute(X)

Do While Not rsAdo.EOF
    If mECHISBCOM <> rsAdo("ECHISBCOM") Then
        arrWECHISB0_nb = arrWECHISB0_nb + 1
        arrWECHISB0(arrWECHISB0_nb).ECHISBCOM = rsAdo("ECHISBCOM")
        mECHISBCOM = rsAdo("ECHISBCOM")
    End If
    
    Select Case Trim(rsAdo("ECHISBCMI"))
        Case "CDM": arrWECHISB0(arrWECHISB0_nb).ECHISBCDM = arrWECHISB0(arrWECHISB0_nb).ECHISBCDM - CCur(rsAdo("ECHISBMTC"))
                    If rsAdo("ECHISBAUT") <> "O" Then arrWECHISB0(arrWECHISB0_nb).ECHISBCDM_AUT = rsAdo("ECHISBAUT")
        Case "ICR": arrWECHISB0(arrWECHISB0_nb).ECHISBICR = arrWECHISB0(arrWECHISB0_nb).ECHISBICR - CCur(rsAdo("ECHISBMTC"))
                    If rsAdo("ECHISBAUT") <> "O" Then arrWECHISB0(arrWECHISB0_nb).ECHISBICR_AUT = rsAdo("ECHISBAUT")
        Case "IDE": arrWECHISB0(arrWECHISB0_nb).ECHISBIDE = arrWECHISB0(arrWECHISB0_nb).ECHISBIDE - CCur(rsAdo("ECHISBMTC"))
                    If rsAdo("ECHISBAUT") <> "O" Then arrWECHISB0(arrWECHISB0_nb).ECHISBIDE_AUT = rsAdo("ECHISBAUT")
        Case "PRE": arrWECHISB0(arrWECHISB0_nb).ECHISBPRE = arrWECHISB0(arrWECHISB0_nb).ECHISBPRE - CCur(rsAdo("ECHISBMTC"))
                    If rsAdo("ECHISBAUT") <> "O" Then arrWECHISB0(arrWECHISB0_nb).ECHISBPRE_AUT = rsAdo("ECHISBAUT")
        Case "PFD": arrWECHISB0(arrWECHISB0_nb).ECHISBPFD = arrWECHISB0(arrWECHISB0_nb).ECHISBPFD - CCur(rsAdo("ECHISBMTC"))
                    If rsAdo("ECHISBAUT") <> "O" Then arrWECHISB0(arrWECHISB0_nb).ECHISBPFD_AUT = rsAdo("ECHISBAUT")
        Case "TDC": arrWECHISB0(arrWECHISB0_nb).ECHISBTDC = arrWECHISB0(arrWECHISB0_nb).ECHISBTDC - CCur(rsAdo("ECHISBMTC"))
                    If rsAdo("ECHISBAUT") <> "O" Then arrWECHISB0(arrWECHISB0_nb).ECHISBTDC_AUT = rsAdo("ECHISBAUT")
    End Select
    rsAdo.MoveNext
Loop



fgDetail_4_Display

fraSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_4_Surveillance_Echelles()
Dim I As Integer, X As String
Dim wCli As String
Dim wCLIENACLI As String
Dim V
On Error GoTo Error_Handler


tvwSelect.Nodes.Clear
tvwInverse.Nodes.Clear
lblSelect = ""
lblInverse = ""
Set rsAdo = Nothing

fgDetail_4_Surveillance_Echelles_Display

fraSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_5_Surveillance_Blocage()
Dim I As Integer, X As String
Dim wCli As String
Dim wCLIENACLI As String
Dim V
On Error GoTo Error_Handler


tvwSelect.Nodes.Clear
tvwInverse.Nodes.Clear
lblSelect = ""
lblInverse = ""
Set rsAdo = Nothing

fgDetail_5_Surveillance_Blocage_Display

fraSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgDetail_4_Display()
Dim V, xWhere As String, xWhere2 As String, xWhere2S As String, xWhere3 As String, wAmj As String, K As Integer, K2 As Integer
Dim X As String, wColor As Long
Dim xECHTABARG As String, xECHTABDON As String
Dim mCOMPTECOM As String, mCLIENACLI As String, mCOMPTEINT As String, mCLIENARES As String
Dim blnConditionsParticulières As Boolean, blnConditions_Ok As Boolean
Dim xFiscal As String
Dim xSTD_Id As String, arrSTD_Id(), arrSTD_Row(), arrSTD_Nb As Long, xECH_Code As String
Dim arrCOMPTECOM(), arrCOMPTECOM_Nb As Long, xCOMPTECOM As String
Dim mSTD_Id As String, mSTD_K As Long
Dim blnOk As Boolean, blnSTD_Row As Boolean
Dim currentRow As Long, blnDupliquer As Boolean
Dim X3 As String, X6 As String, X7 As String, X8 As String, X9 As String
Dim curTDC As Currency

Dim xTaux1 As String, xTaux2 As String, curSeuil1 As String, curSeuil2 As String
Dim kCol As Integer
Dim blnAttribut_EQ As Boolean
On Error GoTo Error_Handler

fgDetail.Visible = False
fgDetail_Reset
'fgDetail.FormatString = "<Racine           |<Compte                                                                                 |<Code   |<Nature|<Intitulé                                                                                  " _
'                      & "|<Echéance     |> Montant              |<Devise /Taux          |> Montant             |<Devise /Taux             " _
'                      & "|Resp |Pays |Etat|<Date validation   "

fgDetail.FormatString = "<Racine    |<Atributs                   |<Compte           |<K|<Intitulé                                                                     " _
                      & "|<IDE taux                 |>IDE seuil              |<IDE taux                      |<ICR taux               |>ICR seuil              |<ICR taux                    " _
                      & "|>TDC          |<CDM           |<PFD           |<PRE           " _
                      & "|<Resp |<Pays |<Attibuts standard        |>IDE               |>ICR                |>TDC               |>CDM                |>PFD            |>PRE           "

fgDetail.Left = 100
fgDetail.Top = 150
fgDetail.Width = 13000

fgDetail.Rows = 1
fgDetail.Row = 0
currentCLIENARES = ""
arrCLIENARES_Nb = 0
mSTD_Id = ""

xWhere2 = ""
xWhere3 = ""
xWhere = " and AUTSYCMON <> 0 "
If chkSelect_Options_4_AUTSICMON = "1" Then xWhere = ""
mCLIENACLI = Trim(txtSelect_Options_4_CLIENACLI)
If mCLIENACLI <> "" Then
    X = Format(Val(mCLIENACLI), "0000000")
    xWhere = xWhere & " and AUTSYCCLI = '" & X & "'"
    xWhere2 = " and CLIENACLI = '" & X & "'"
End If
mCLIENARES = Mid$(cboSelect_Options_4_CLIENARES, 1, 3)
If mCLIENARES <> "" Then
    xWhere = xWhere & " and CLIENARES = '" & mCLIENARES & "'"
    xWhere2 = " and CLIENARES = '" & mCLIENARES & "'"
End If
xWhere2S = xWhere2
'============================================================================================================
' conditions standard
arrSTD_Nb = 0
ReDim arrSTD_Id(1), arrSTD_Row(1)

If mCLIENACLI = "" Or mCLIENACLI = "99999" Then
Else
    X = "select DISTINCT CLIENACAT from " & paramIBM_Library_SABSPE & ".YBIACPT0" _
         & " where CLIENACLI = '" & Format(mCLIENACLI, "0000000") & "'"
    Set rsAdo = cnAdo.Execute(X)
    If Not rsAdo.EOF Then xWhere3 = "  and substring(ECHTABARG,31,3) = '" & rsAdo("CLIENACAT") & "'"
End If


'If mCLIENACLI = "" Or mCLIENACLI = "99999" Then
    X = Trim(cboSelect_Options_4_Code)
    If X <> "" Then
        xWhere3 = xWhere3 & " and substring(ECHTABARG,64,3) = '" & X & "'"
    End If
    

    X = "select count(*) from " & paramIBM_Library_SAB & ".ZECHTAB0" _
         & " where ECHTABETA = " & currentZMNURUT0.MNURUTETB _
         & " and   ECHTABNUM = 8" & xWhere3
    Set rsAdo = cnAdo.Execute(X)
   ReDim arrSTD_Id(rsAdo(0) + 1), arrSTD_Row(rsAdo(0) + 1)
   
   
    X = "select * from " & paramIBM_Library_SAB & ".ZECHTAB0" _
         & " where ECHTABETA = " & currentZMNURUT0.MNURUTETB _
         & " and   ECHTABNUM = 8" & xWhere3 _
         & " order by ECHTABARG"
    Set rsAdo = cnAdo.Execute(X)
    
    Do While Not rsAdo.EOF
        xECHTABDON = rsAdo("ECHTABDON")
        If chkSelect_Options_4_ECHTABDON_S = "0" And Mid$(xECHTABDON, 219, 1) = "S" Then
        
        Else
        
            xECHTABARG = rsAdo("ECHTABARG")
            
            xSTD_Id = RTrim(Mid$(xECHTABARG, 1, 3) & "  " & Mid$(xECHTABARG, 21, 3) & "  " & Mid$(xECHTABARG, 31, 3) & "  " & Mid$(xECHTABARG, 41, 3))
            If xSTD_Id <> arrSTD_Id(arrSTD_Nb) Then
                fgDetail.Rows = fgDetail.Rows + 1
                fgDetail.Row = fgDetail.Rows - 1
                
                fgDetail.Col = 1: fgDetail.Text = xSTD_Id
                fgDetail.Col = 3: fgDetail.Text = "0"
                
                fgDetail.Col = 5: fgDetail.CellBackColor = mColor_Y1
                fgDetail.Col = 6: fgDetail.CellBackColor = mColor_Y1
                fgDetail.Col = 7: fgDetail.CellBackColor = mColor_Y1
                fgDetail.Col = 8: fgDetail.CellBackColor = mColor_G1
                fgDetail.Col = 9: fgDetail.CellBackColor = mColor_G1
                fgDetail.Col = 10: fgDetail.CellBackColor = mColor_G1
                fgDetail.Col = 15: fgDetail.CellBackColor = RGB(240, 240, 240)
                fgDetail.Col = 16: fgDetail.CellBackColor = RGB(240, 240, 240)
                fgDetail.Col = 17: fgDetail.CellBackColor = RGB(240, 240, 240)
               
                arrSTD_Nb = arrSTD_Nb + 1
                arrSTD_Id(arrSTD_Nb) = xSTD_Id
                arrSTD_Row(arrSTD_Nb) = fgDetail.Row
            End If
            xECH_Code = Mid$(xECHTABARG, 64, 3)
        
            If xECH_Code = "TDC" Then
                curTDC = CCur(convX2P(Mid$(xECHTABDON, 21, 4))) / 100
                fgDetail.Col = 11: fgDetail.Text = Format(curTDC, "### ### ### ##0.00")
            Else
                xTaux1 = Mid$(xECHTABDON, 25, 6) & " " & num_Taux_Display(CDbl(convX2P(Mid$(xECHTABDON, 35, 5))) / 1000000)
                curSeuil1 = Format(CCur(convX2P(Mid$(xECHTABDON, 41, 7))) / 100, "### ### ### ##0.00")
                If Trim(Mid$(xECHTABDON, 56, 6)) <> "" Then
                   xTaux2 = Mid$(xECHTABDON, 56, 6) & " " & num_Taux_Display(CDbl(convX2P(Mid$(xECHTABDON, 66, 5))) / 1000000)
                   curSeuil2 = Format(CCur(convX2P(Mid$(xECHTABDON, 72, 7))), "### ### ### ##0.00")
                Else
                    xTaux2 = ""
                    curSeuil2 = ""
                End If
                If CCur(convX2P(Mid$(xECHTABDON, 41, 7))) <> 0 Then
                    Call MsgBox("Seuil 1 <> 0.00", vbCritical, "Conditions échelles")
                End If
                Select Case xECH_Code
                    Case "IDE":
                        fgDetail.Col = 5: fgDetail.Text = xTaux1
                        fgDetail.Col = 6: fgDetail.Text = curSeuil2
                        fgDetail.Col = 7: fgDetail.Text = xTaux2
                    Case "ICR":
                        fgDetail.Col = 8: fgDetail.Text = xTaux1
                        fgDetail.Col = 9: fgDetail.Text = curSeuil2
                        fgDetail.Col = 10: fgDetail.Text = xTaux2
                    Case "CDM":
                        fgDetail.Col = 12: fgDetail.Text = xTaux1
                    Case "PFD":
                        fgDetail.Col = 13: fgDetail.Text = xTaux1
                    Case "PRE":
                        fgDetail.Col = 14: fgDetail.Text = xTaux1
                    Case Else
                        Call MsgBox("code inconnu : " & xECH_Code, vbCritical, "Conditions échelles")
                End Select
            End If
            
            If Mid$(xECHTABDON, 219, 1) = "S" Then
                wAmj = convX2P(Mid$(xECHTABDON, 215, 4)) + 19000000
                fgDetail.Col = 17: fgDetail.Text = "Ech : " & dateImp10(wAmj)
        
                For K = 0 To 17: fgDetail.Col = K: fgDetail.CellForeColor = vbRed: Next K
            Else
                For K = 0 To 17: fgDetail.Col = K: fgDetail.CellForeColor = vbBlue: Next K
            
            End If
            
        End If

        rsAdo.MoveNext
    Loop


    If mCLIENACLI = "99999" Then GoTo Select_End
'End If

'============================================================================================================

'______________________________________________________________________________________________________
X = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0 , " & paramIBM_Library_SAB & ".ZCLIENA0" _
    & " where AUTSYCAUT = 'DEC' and AUTSYCCLI = CLIENACLI" & xWhere _
    & " order by AUTSYCCLI"

Set rsAdo = cnAdo.Execute(X)

Do While Not rsAdo.EOF
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    
    wColor = vbBlue
    If rsAdo("AUTSYCFIN") > 0 Then
        wAmj = rsAdo("AUTSYCFIN") + 19000000
        fgDetail.Col = 8: fgDetail.Text = "E:" & dateImp10(wAmj)
        If wAmj <= YBIATAB0_DATE_CPT_J Then wColor = vbMagenta
        fgDetail.CellForeColor = wColor
    End If

    fgDetail.Col = 0
    If rsAdo("AUTSYCGPE") = "O" Then
        fgDetail.Text = Val(rsAdo("AUTSYCCLI")) & "-G"
    Else
        fgDetail.Text = Val(rsAdo("AUTSYCCLI"))
    End If
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 3: fgDetail.Text = "1"
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 5: fgDetail.Text = rsAdo("AUTSYCAUT")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 4: fgDetail.Text = rsAdo("CLIENARA1")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 6: fgDetail.Text = Format(rsAdo("AUTSYCMON"), "### ### ### ##0.00")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 7: fgDetail.Text = rsAdo("AUTSYCDEV")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 10: fgDetail.Text = "Blocage : " & rsAdo("AUTSYCBLO")
    fgDetail.CellForeColor = wColor
    
    fgDetail.Col = 15: fgDetail.Text = rsAdo("CLIENARES")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 16: fgDetail.Text = rsAdo("CLIENARSD")
    fgDetail.CellForeColor = wColor
        
    If rsAdo("AUTSYCDVL") > 0 Then
        wAmj = rsAdo("AUTSYCDVL") + 19000000
        fgDetail.Col = 17: fgDetail.Text = "màj : " & rsAdo("AUTSYCCET") & " - " & dateImp10(wAmj)
        fgDetail.CellForeColor = wColor
    Else
        fgDetail.Col = 17: fgDetail.Text = "màj : " & rsAdo("AUTSYCCET")
        fgDetail.CellForeColor = wColor
    End If

    For I = 0 To 14: fgDetail.Col = I: fgDetail.CellBackColor = mColor_B0: Next I
    fgDetail.Col = 15: fgDetail.CellBackColor = RGB(240, 240, 240)
    fgDetail.Col = 16: fgDetail.CellBackColor = RGB(240, 240, 240)
    fgDetail.Col = 17: fgDetail.CellBackColor = RGB(240, 240, 240)
    rsAdo.MoveNext
Loop

'______________________________________________________________________________________________________
X = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0 , " & paramIBM_Library_SABSPE & ".YBIACPT0" _
    & " where AUTSYCAUT = COMPTECOM and COMPTEFON <> '4'" & xWhere _
    & " order by AUTSYCCLI, COMPTECOM"

Set rsAdo = cnAdo.Execute(X)

Do While Not rsAdo.EOF
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    
    wColor = vbBlue
    If rsAdo("AUTSYCFIN") > 0 Then
        wAmj = rsAdo("AUTSYCFIN") + 19000000
        fgDetail.Col = 8: fgDetail.Text = "E:" & dateImp10(wAmj)
        If wAmj <= YBIATAB0_DATE_CPT_J Then wColor = vbMagenta
        fgDetail.CellForeColor = wColor
    End If
    xFiscal = prtYBIAMVT0_A4_Pauget_Constans_Fiscal(Trim(rsAdo("CLIENARSD")))

    fgDetail.Col = 0: fgDetail.Text = Val(rsAdo("AUTSYCCLI"))
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 1: fgDetail.Text = rsAdo("PLANCOPRO") & "  " & rsAdo("COMPTEDEV") & "  " & rsAdo("CLIENACAT") & "  " & xFiscal
    fgDetail.Col = 2: fgDetail.Text = rsAdo("COMPTECOM")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 3: fgDetail.Text = 2
    fgDetail.CellForeColor = wColor
    'fgDetail.Col = 3: fgDetail.Text = rsAdo("AUTSYCAUT")
    'fgDetail.CellForeColor = wColor
    fgDetail.Col = 4: fgDetail.Text = rsAdo("COMPTEINT")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 6: fgDetail.Text = Format(rsAdo("AUTSYCMON"), "### ### ### ##0.00")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 7: fgDetail.Text = rsAdo("AUTSYCDEV")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 10: fgDetail.Text = "Blocage : " & rsAdo("AUTSYCBLO")
    fgDetail.CellForeColor = wColor
    
    fgDetail.Col = 15: fgDetail.Text = rsAdo("CLIENARES")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 16: fgDetail.Text = rsAdo("CLIENARSD")
    fgDetail.CellForeColor = wColor
        
    If rsAdo("AUTSYCDVL") > 0 Then
        wAmj = rsAdo("AUTSYCDVL") + 19000000
        fgDetail.Col = 17: fgDetail.Text = "màj : " & rsAdo("AUTSYCCET") & " - " & dateImp10(wAmj)
        fgDetail.CellForeColor = wColor
    Else
        fgDetail.Col = 17: fgDetail.Text = "màj : " & rsAdo("AUTSYCCET")
        fgDetail.CellForeColor = wColor
    End If

    For I = 0 To 14: fgDetail.Col = I: fgDetail.CellBackColor = mColor_B0: Next I
    fgDetail.Col = 15: fgDetail.CellBackColor = RGB(240, 240, 240)
    fgDetail.Col = 16: fgDetail.CellBackColor = RGB(240, 240, 240)
    fgDetail.Col = 17: fgDetail.CellBackColor = RGB(240, 240, 240)
    Call fgDetail_4_Display_WECHISB0(rsAdo("COMPTECOM"))
    rsAdo.MoveNext
Loop
'______________________________________________________________________________________________________

blnConditions_Ok = False
X = Trim(cboSelect_Options_4_Code)
If X <> "" Then
    xWhere2 = xWhere2 & " and ECHTABARG like '%" & X & "%'"
End If
'_________________________________________________________________________________________________________

X = "select count(*) from " & paramIBM_Library_SAB & ".ZECHTAB0 , " & paramIBM_Library_SABSPE & ".YBIACPT0" _
     & " where ECHTABETA = " & currentZMNURUT0.MNURUTETB _
     & " and   ECHTABNUM = 9 And substring(ECHTABARG, 3, 20) = COMPTECOM and COMPTEFON <> '4'" & xWhere2
Set rsAdo = cnAdo.Execute(X)
ReDim arrCOMPTECOM(rsAdo(0) + 1)
arrCOMPTECOM_Nb = 0
'_________________________________________________________________________________________________________

X = "select * from " & paramIBM_Library_SAB & ".ZECHTAB0 , " & paramIBM_Library_SABSPE & ".YBIACPT0" _
     & " where ECHTABETA = " & currentZMNURUT0.MNURUTETB _
     & " and   ECHTABNUM = 9 And substring(ECHTABARG, 3, 20) = COMPTECOM and COMPTEFON <> '4'" & xWhere2 _
     & " order by ECHTABARG" ' CLIENACLI"
Set rsAdo = cnAdo.Execute(X)

'______________________________________________________________________________________________________
'If Not blnConditionsParticulières Then

    Do While Not rsAdo.EOF
        xECHTABDON = rsAdo("ECHTABDON")
        xECHTABARG = rsAdo("ECHTABARG")
        If chkSelect_Options_4_ECHTABDON_S = "0" And Mid$(xECHTABDON, 219, 1) = "S" Then
        
        Else
            If arrCOMPTECOM(arrCOMPTECOM_Nb) <> rsAdo("COMPTECOM") Then
            
                fgDetail.Rows = fgDetail.Rows + 1
                fgDetail.Row = fgDetail.Rows - 1
                
                fgDetail.Col = 5: fgDetail.CellBackColor = mColor_Y1
                fgDetail.Col = 6: fgDetail.CellBackColor = mColor_Y1
                fgDetail.Col = 7: fgDetail.CellBackColor = mColor_Y1
                fgDetail.Col = 8: fgDetail.CellBackColor = mColor_G1
                fgDetail.Col = 9: fgDetail.CellBackColor = mColor_G1
                fgDetail.Col = 10: fgDetail.CellBackColor = mColor_G1
                fgDetail.Col = 15: fgDetail.CellBackColor = RGB(240, 240, 240)
                fgDetail.Col = 16: fgDetail.CellBackColor = RGB(240, 240, 240)
                fgDetail.Col = 17: fgDetail.CellBackColor = RGB(240, 240, 240)
   
                fgDetail.Col = 0: fgDetail.Text = Val(rsAdo("CLIENACLI"))
                fgDetail.Col = 1: fgDetail.Text = rsAdo("PLANCOPRO") & "  " & rsAdo("COMPTEDEV") & "  " & rsAdo("CLIENACAT") & "  " & xFiscal
                fgDetail.Col = 2: fgDetail.Text = Trim(Mid$(xECHTABARG, 3, 20))
                fgDetail.Col = 3: fgDetail.Text = 3
                fgDetail.Col = 4: fgDetail.Text = rsAdo("COMPTEINT")
                fgDetail.Col = 15: fgDetail.Text = rsAdo("CLIENARES")
                fgDetail.Col = 16: fgDetail.Text = rsAdo("CLIENARSD")
                
                arrCOMPTECOM_Nb = arrCOMPTECOM_Nb + 1
                arrCOMPTECOM(arrCOMPTECOM_Nb) = rsAdo("COMPTECOM")
                
                For K = 0 To 17: fgDetail.Col = K: fgDetail.CellForeColor = RGB(160, 0, 96): Next K
               
                Call fgDetail_4_Display_WECHISB0(rsAdo("COMPTECOM"))

            End If
            
            xECHTABARG = rsAdo("ECHTABARG")
            xFiscal = prtYBIAMVT0_A4_Pauget_Constans_Fiscal(Trim(rsAdo("CLIENARSD")))

            xECH_Code = Mid$(xECHTABARG, 26, 3)

            If xECH_Code = "TDC" Then
                curTDC = CCur(convX2P(Mid$(xECHTABDON, 21, 4))) / 100
                fgDetail.Col = 11: fgDetail.Text = Format(curTDC, "### ### ### ##0.00")
                If -oldWECHISB0.ECHISBTDC <> curTDC Then fgDetail.CellBackColor = mColor_W1
            Else
                xTaux1 = Mid$(xECHTABDON, 25, 6) & " " & num_Taux_Display(CDbl(convX2P(Mid$(xECHTABDON, 35, 5))) / 1000000)
                curSeuil1 = Format(CCur(convX2P(Mid$(xECHTABDON, 41, 7))) / 100, "### ### ### ##0.00")
                If Trim(Mid$(xECHTABDON, 56, 6)) <> "" Then
                   xTaux2 = Mid$(xECHTABDON, 56, 6) & " " & num_Taux_Display(CDbl(convX2P(Mid$(xECHTABDON, 66, 5))) / 1000000)
                   curSeuil2 = Format(CCur(convX2P(Mid$(xECHTABDON, 72, 7))), "### ### ### ##0.00")
                Else
                    xTaux2 = ""
                    curSeuil2 = ""
                End If
                If CCur(convX2P(Mid$(xECHTABDON, 41, 7))) <> 0 Then
                    Call MsgBox("Seuil 1 <> 0.00", vbCritical, "Conditions échelles")
                End If
                Select Case xECH_Code
                    Case "IDE":
                        fgDetail.Col = 5: fgDetail.Text = xTaux1
                        fgDetail.Col = 6: fgDetail.Text = curSeuil2
                        fgDetail.Col = 7: fgDetail.Text = xTaux2
                    Case "ICR":
                        fgDetail.Col = 8: fgDetail.Text = xTaux1
                        fgDetail.Col = 9: fgDetail.Text = curSeuil2
                        fgDetail.Col = 10: fgDetail.Text = xTaux2
                    Case "CDM":
                        fgDetail.Col = 12: fgDetail.Text = xTaux1
                    Case "PFD":
                        fgDetail.Col = 13: fgDetail.Text = xTaux1
                    Case "PRE":
                        fgDetail.Col = 14: fgDetail.Text = xTaux1
                    Case Else
                        Call MsgBox("code inconnu : " & xECH_Code, vbCritical, "Conditions échelles")
                End Select
            End If
            
            If Mid$(xECHTABDON, 219, 1) = "S" Then
                wAmj = convX2P(Mid$(xECHTABDON, 215, 4)) + 19000000
                fgDetail.Col = 17: fgDetail.Text = "Ech : " & dateImp10(wAmj)
        
                For K = 0 To 17: fgDetail.Col = K: fgDetail.CellForeColor = vbRed: Next K
            End If
            
        End If
        rsAdo.MoveNext
    Loop
'Else
'______________________________________________________________________________________________________
'    Do While Not rsAdo.EOF
'        xECHTABARG = rsAdo("ECHTABARG")
'        If mCOMPTECOM <> rsAdo("COMPTECOM") Then
'            If Not blnConditions_Ok Then
'                fgDetail.Rows = fgDetail.Rows + 1
'                fgDetail.Row = fgDetail.Rows - 1
'
'                fgDetail.Col = 0: fgDetail.Text = mCLIENACLI
'                fgDetail.Col = 1: fgDetail.Text = mCOMPTECOM
'                fgDetail.Col = 2: fgDetail.Text = "3?"
'                fgDetail.Col = 3: fgDetail.Text = "IDE"
'                fgDetail.Col = 4: fgDetail.Text = mCOMPTEINT
'
'                fgDetail.Col = 15: fgDetail.Text = rsAdo("CLIENARES")
'                fgDetail.Col = 16: fgDetail.Text = rsAdo("CLIENARSD")
'
'                For K = 0 To 17: fgDetail.Col = K: fgDetail.CellForeColor = vbRed: Next K
'
'            End If
'
'            blnConditions_Ok = False
'            mCOMPTECOM = rsAdo("COMPTECOM")
'            mCLIENACLI = rsAdo("CLIENACLI")
'            mCOMPTEINT = rsAdo("COMPTEINT")
'        End If
'
'        If Mid$(rsAdo("ECHTABDON"), 219, 1) <> "S" Then blnConditions_Ok = True
'
'        rsAdo.MoveNext
'    Loop

'End If


'============================================================================================================

mSTD_Id = "": mSTD_K = 0
X = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0" _
     & " where PLANCOPRO in (" & lstPLANCOPRO & ") and COMPTEFON <> '4'" & xWhere2S _
     & " order by CLIENACLI"

Set rsAdo = cnAdo.Execute(X)


'______________________________________________________________________________________________________

    Do While Not rsAdo.EOF
        blnOk = True
        xCOMPTECOM = rsAdo("COMPTECOM")
        For K = 1 To arrCOMPTECOM_Nb
            If xCOMPTECOM = arrCOMPTECOM(K) Then
                blnOk = False
                Exit For
            End If
        Next K
        If blnOk Then
            fgDetail.Rows = fgDetail.Rows + 1
            fgDetail.Row = fgDetail.Rows - 1
            
            fgDetail.Col = 5: fgDetail.CellBackColor = mColor_Y1
            fgDetail.Col = 6: fgDetail.CellBackColor = mColor_Y1
            fgDetail.Col = 7: fgDetail.CellBackColor = mColor_Y1
            fgDetail.Col = 8: fgDetail.CellBackColor = mColor_G1
            fgDetail.Col = 9: fgDetail.CellBackColor = mColor_G1
            fgDetail.Col = 10: fgDetail.CellBackColor = mColor_G1
            fgDetail.Col = 15: fgDetail.CellBackColor = RGB(240, 240, 240)
            fgDetail.Col = 16: fgDetail.CellBackColor = RGB(240, 240, 240)
            fgDetail.Col = 17: fgDetail.CellBackColor = RGB(240, 240, 240)
                
            xFiscal = prtYBIAMVT0_A4_Pauget_Constans_Fiscal(Trim(rsAdo("CLIENARSD")))

            fgDetail.Col = 0: fgDetail.Text = Val(rsAdo("CLIENACLI"))
            xSTD_Id = rsAdo("PLANCOPRO") & "  " & rsAdo("COMPTEDEV") & "  " & rsAdo("CLIENACAT") & "  " & xFiscal
            fgDetail.Col = 1: fgDetail.Text = xSTD_Id
            fgDetail.Col = 2: fgDetail.Text = rsAdo("COMPTECOM")
            fgDetail.Col = 3: fgDetail.Text = 3
            fgDetail.Col = 4: fgDetail.Text = rsAdo("COMPTEINT")
            fgDetail.Col = 15: fgDetail.Text = rsAdo("CLIENARES")
            fgDetail.Col = 16: fgDetail.Text = rsAdo("CLIENARSD")
            
            Call fgDetail_4_Display_WECHISB0(rsAdo("COMPTECOM"))

            If mSTD_Id = xSTD_Id Then
            Else
                blnSTD_Row = False
                blnAttribut_EQ = False
                For mSTD_K = 1 To arrSTD_Nb
                    If xSTD_Id = arrSTD_Id(mSTD_K) Then
                        blnSTD_Row = True
                        mSTD_Id = xSTD_Id
                        blnAttribut_EQ = True
                        Exit For
                    End If
                Next mSTD_K
                
                If Not blnSTD_Row Then
                    X = Trim(Mid$(xSTD_Id, 1, 13))
                    For mSTD_K = 1 To arrSTD_Nb
                        If X = arrSTD_Id(mSTD_K) Then
                            blnSTD_Row = True
                            mSTD_Id = xSTD_Id
                            Exit For
                        End If
                    Next mSTD_K
                End If
                
                If Not blnSTD_Row Then
                    X = Trim(Mid$(xSTD_Id, 1, 9))
                    For mSTD_K = 1 To arrSTD_Nb
                        If X = arrSTD_Id(mSTD_K) Then
                            blnSTD_Row = True
                            mSTD_Id = xSTD_Id
                            Exit For
                        End If
                    Next mSTD_K
                End If
                
                 If Not blnSTD_Row Then
                    X = Trim(Mid$(xSTD_Id, 1, 3))
                    For mSTD_K = 1 To arrSTD_Nb
                        If X = arrSTD_Id(mSTD_K) Then
                            blnSTD_Row = True
                            mSTD_Id = xSTD_Id
                            Exit For
                        End If
                    Next mSTD_K
                End If
                
               
            End If
            
            fgDetail.Col = 17: fgDetail.Text = arrSTD_Id(mSTD_K)
            
            currentRow = fgDetail.Row
            If blnSTD_Row Then
                If Not blnAttribut_EQ Then fgDetail.Col = 1: fgDetail.CellBackColor = RGB(240, 240, 240)
                For K = 5 To 14
                
                    fgDetail.Col = K: fgDetail.Row = arrSTD_Row(mSTD_K): X = fgDetail.Text: fgDetail.Row = fgDetail.Rows - 1: fgDetail.Text = X
                Next K
                fgDetail.Col = 11: curTDC = Val(fgDetail.Text)
                If -oldWECHISB0.ECHISBTDC <> curTDC Then fgDetail.CellBackColor = mColor_W1
            Else
                fgDetail.Col = 17: fgDetail.Text = "???????????????????"
                fgDetail.Col = 1: fgDetail.CellBackColor = mColor_W1
            End If
            
        
        End If

        rsAdo.MoveNext
    Loop
'============================================================================================================

For K = 2 To fgDetail.Rows - 1
    fgDetail.Row = K
    fgDetail.Col = 15
    If currentCLIENARES <> Trim(fgDetail.Text) Then
        currentCLIENARES = Trim(fgDetail.Text)
        blnOk = False
        For K2 = 1 To arrCLIENARES_Nb
            If currentCLIENARES = arrCLIENARES(K2) Then blnOk = True: Exit For
        Next K2
        If Not blnOk Then
            arrCLIENARES_Nb = arrCLIENARES_Nb + 1
            arrCLIENARES(arrCLIENARES_Nb) = currentCLIENARES
        End If
    End If
    
Next K


Select_End:

fgDetail_Sort1 = 1: fgDetail_Sort2 = 3: fgDetail_Sort
fgDetail.Visible = True
Set rsAdo = Nothing
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgDetail_4_Surveillance_Echelles_Display()
Dim V, xWhere As String, xWhere2 As String, wAmj As String, K As Integer, I As Integer
Dim X As String, wColor As Long
Dim xECHTABARG As String, xECHTABDON As String, xECHTABARG_Compte As String
Dim mCOMPTECOM As String, mCLIENACLI As String, mCOMPTEINT As String
Dim blnConditionsParticulières As Boolean, blnConditions_Ok As Boolean

On Error GoTo Error_Handler

fgDetail.Visible = False
fgDetail_Reset
fgDetail.FormatString = "<Racine           |<Compte                                                |<Code   |<Nature|<Intitulé                                                                " _
                      & "|<Echéance     |> Montant                  |<Devise /Taux             |> Montant                  |<Devise /Taux             " _
                      & "|Resp |Pays |Etat|<Date validation   "
fgDetail.Left = 100
fgDetail.Top = 150
fgDetail.Width = 13000

fgDetail.Rows = 1
fgDetail.Row = 0
xWhere2 = ""
xWhere = " and AUTSYCMON <> 0 "
If chkSelect_Options_4_AUTSICMON = "1" Then xWhere = ""
X = Trim(txtSelect_Options_4_CLIENACLI)
If X <> "" Then
    X = Format(Val(X), "0000000")
    xWhere = xWhere & " and AUTSYCCLI = '" & X & "'"
    xWhere2 = " and CLIENACLI = '" & X & "'"
End If
'______________________________________________________________________________________________________
X = "select count(*) from " & paramIBM_Library_SAB & ".ZAUTSYC0 , " & paramIBM_Library_SAB & ".ZCLIENA0" _
    & " where AUTSYCAUT = 'DEC' and AUTSYCCLI = CLIENACLI" & xWhere
Set rsAdo = cnAdo.Execute(X)

ReDim arrZAUTSYC0(rsAdo(0) + 1), arrCLIENARA1(rsAdo(0) + 1)
arrZAUTSYC0_Nb = 0
X = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0 , " & paramIBM_Library_SAB & ".ZCLIENA0" _
    & " where AUTSYCAUT = 'DEC' and AUTSYCCLI = CLIENACLI" & xWhere _
    & " order by AUTSYCCLI"

Set rsAdo = cnAdo.Execute(X)

Do While Not rsAdo.EOF
    arrZAUTSYC0_Nb = arrZAUTSYC0_Nb + 1
    Call rsZAUTSYC0_GetBuffer(rsAdo, arrZAUTSYC0(arrZAUTSYC0_Nb))
    arrCLIENARA1(arrZAUTSYC0_Nb) = rsAdo("CLIENARA1")
    rsAdo.MoveNext
Loop



'______________________________________________________________________________________________________
X = "select count(*) from " & paramIBM_Library_SAB & ".ZAUTSYC0 , " & paramIBM_Library_SABSPE & ".YBIACPT0" _
    & " where AUTSYCAUT = COMPTECOM and COMPTEFON <> '4'" & xWhere

Set rsAdo = cnAdo.Execute(X)

ReDim Preserve arrZAUTSYC0(arrZAUTSYC0_Nb + rsAdo(0) + 1), arrCLIENARA1(arrZAUTSYC0_Nb + rsAdo(0) + 1)

X = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0 , " & paramIBM_Library_SABSPE & ".YBIACPT0" _
    & " where AUTSYCAUT = COMPTECOM and COMPTEFON <> '4'" & xWhere _
    & " order by AUTSYCCLI, COMPTECOM"

Set rsAdo = cnAdo.Execute(X)

Do While Not rsAdo.EOF
    kPLANCOPRO = InStr(lstPLANCOPRO, rsAdo("PLANCOPRO"))
    If kPLANCOPRO > 0 Then
        arrZAUTSYC0_Nb = arrZAUTSYC0_Nb + 1
        Call rsZAUTSYC0_GetBuffer(rsAdo, arrZAUTSYC0(arrZAUTSYC0_Nb))
        arrCLIENARA1(arrZAUTSYC0_Nb) = rsAdo("COMPTEINT")
    End If
    rsAdo.MoveNext
Loop

'================================================================================================================

X = "select * from " & paramIBM_Library_SAB & ".ZECHTAB0 , " & paramIBM_Library_SABSPE & ".YBIACPT0" _
     & " where ECHTABETA = " & currentZMNURUT0.MNURUTETB _
     & " and   ECHTABNUM = 9 And ECHTABARG like '%IDE%'  And substring(ECHTABARG, 3, 20) = COMPTECOM  and COMPTEFON <> '4'" _
     & " order by CLIENACLI"
Set rsAdo = cnAdo.Execute(X)

Do While Not rsAdo.EOF
    xECHTABARG = rsAdo("ECHTABARG"): xECHTABARG_Compte = Mid$(xECHTABARG, 3, 20)
    xECHTABDON = rsAdo("ECHTABDON")
    
    If Mid$(xECHTABDON, 219, 1) <> "S" Then
    
        For K = 1 To arrZAUTSYC0_Nb
            If xECHTABARG_Compte = arrZAUTSYC0(K).AUTSYCAUT _
            Or rsAdo("CLIENACLI") = arrZAUTSYC0(K).AUTSYCCLI Then
                If arrZAUTSYC0(K).AUTSYCETA > 0 Then
                    arrZAUTSYC0(K).AUTSYCETA = -1
                    xZAUTSYC0 = arrZAUTSYC0(K)
            
                    fgDetail.Rows = fgDetail.Rows + 1
                    fgDetail.Row = fgDetail.Rows - 1
                    
                    wColor = vbBlue
                    If xZAUTSYC0.AUTSYCFIN > 0 Then
                        wAmj = xZAUTSYC0.AUTSYCFIN + 19000000
                        fgDetail.Col = 5: fgDetail.Text = dateImp10(wAmj)
                        If wAmj <= YBIATAB0_DATE_CPT_J Then wColor = vbMagenta
                        fgDetail.CellForeColor = wColor
                    End If
                    fgDetail.Col = 0
                    If xZAUTSYC0.AUTSYCGPE = "O" Then
                        fgDetail.Text = xZAUTSYC0.AUTSYCCLI & "-G"
                    Else
                        fgDetail.Text = xZAUTSYC0.AUTSYCCLI
                    End If
                    fgDetail.CellForeColor = wColor
                    'fgDetail.Col = 1: fgDetail.Text = rsAdo("CLIENACLI")
                    'fgDetail.CellForeColor = wColor
                    'fgDetail.Col = 2: fgDetail.Text = 2
                    'fgDetail.CellForeColor = wColor
                    'fgDetail.Col = 3: fgDetail.Text = xZAUTSYC0.AUTSYCAUT
                    'fgDetail.CellForeColor = wColor
                    If Len(Trim(xZAUTSYC0.AUTSYCAUT)) > 3 Then
                        fgDetail.Col = 1: fgDetail.Text = xZAUTSYC0.AUTSYCAUT
                        fgDetail.CellForeColor = wColor
                        fgDetail.Col = 2: fgDetail.Text = "2"
                        fgDetail.CellForeColor = wColor
                    Else
                        fgDetail.Col = 2
                            fgDetail.Text = "1"
                        fgDetail.CellForeColor = wColor
                        fgDetail.Col = 3: fgDetail.Text = xZAUTSYC0.AUTSYCAUT
                        fgDetail.CellForeColor = wColor
                    End If

                    fgDetail.Col = 4: fgDetail.Text = rsAdo("CLIENARA1")
                    fgDetail.CellForeColor = wColor
                    fgDetail.Col = 6: fgDetail.Text = Format(xZAUTSYC0.AUTSYCMON, "### ### ### ##0.00")
                    fgDetail.CellForeColor = wColor
                    fgDetail.Col = 7: fgDetail.Text = xZAUTSYC0.AUTSYCDEV
                    fgDetail.CellForeColor = wColor
                    fgDetail.Col = 8: fgDetail.Text = "Blocage : " & xZAUTSYC0.AUTSYCBLO
                    fgDetail.CellForeColor = wColor
    
                fgDetail.Col = 10: fgDetail.Text = rsAdo("CLIENARES")
                fgDetail.CellForeColor = wColor
                fgDetail.Col = 11: fgDetail.Text = rsAdo("CLIENARSD")
                fgDetail.CellForeColor = wColor
        
                    fgDetail.Col = 12: fgDetail.Text = xZAUTSYC0.AUTSYCCET
                    fgDetail.CellForeColor = wColor
                    If xZAUTSYC0.AUTSYCDVL > 0 Then
                        wAmj = xZAUTSYC0.AUTSYCDVL + 19000000
                        fgDetail.Col = 13: fgDetail.Text = dateImp10(wAmj)
                        fgDetail.CellForeColor = wColor
                    End If

                    For I = 0 To 13: fgDetail.Col = I: fgDetail.CellBackColor = mColor_B0: Next I

                End If
                fgDetail.Rows = fgDetail.Rows + 1
                fgDetail.Row = fgDetail.Rows - 1
                fgDetail.Col = 0: fgDetail.Text = rsAdo("CLIENACLI")
                fgDetail.Col = 1: fgDetail.Text = Mid$(xECHTABARG, 3, 20)
                fgDetail.Col = 2: fgDetail.Text = 3
                fgDetail.Col = 3: fgDetail.Text = Mid$(xECHTABARG, 26, 3)
                fgDetail.Col = 4: fgDetail.Text = rsAdo("COMPTEINT")
            
            
                fgDetail.Col = 6: fgDetail.Text = Format(CCur(convX2P(Mid$(xECHTABDON, 41, 7))) / 100, "### ### ### ##0.00")
                fgDetail.Col = 7: fgDetail.Text = Mid$(xECHTABDON, 25, 6) & " " & num_Taux_Display(CDbl(convX2P(Mid$(xECHTABDON, 35, 5))) / 1000000)
    
                fgDetail.Col = 10: fgDetail.Text = rsAdo("CLIENARES")
                fgDetail.CellForeColor = wColor
                fgDetail.Col = 11: fgDetail.Text = rsAdo("CLIENARSD")
                fgDetail.CellForeColor = wColor
        
                If Trim(Mid$(xECHTABDON, 56, 6)) <> "" Then
                    fgDetail.Col = 8: fgDetail.Text = Format(CCur(convX2P(Mid$(xECHTABDON, 72, 7))), "### ### ### ##0.00")
                    fgDetail.Col = 9: fgDetail.Text = Mid$(xECHTABDON, 56, 6) & " " & num_Taux_Display(CDbl(convX2P(Mid$(xECHTABDON, 66, 5))) / 1000000)
                End If
                
                Exit For
                    
            End If
        Next K
    End If
    rsAdo.MoveNext
Loop


For K = 1 To arrZAUTSYC0_Nb
        If arrZAUTSYC0(K).AUTSYCETA > 0 Then
            arrZAUTSYC0(K).AUTSYCETA = -1
            xZAUTSYC0 = arrZAUTSYC0(K)
    
            fgDetail.Rows = fgDetail.Rows + 1
            fgDetail.Row = fgDetail.Rows - 1
            
            wColor = vbRed
            If xZAUTSYC0.AUTSYCFIN > 0 Then
                wAmj = xZAUTSYC0.AUTSYCFIN + 19000000
                fgDetail.Col = 5: fgDetail.Text = dateImp10(wAmj)
                If wAmj <= YBIATAB0_DATE_CPT_J Then wColor = vbMagenta
                fgDetail.CellForeColor = wColor
            End If
                fgDetail.Col = 0
                If xZAUTSYC0.AUTSYCGPE = "O" Then
                    fgDetail.Text = xZAUTSYC0.AUTSYCCLI & "-G"
                Else
                    fgDetail.Text = xZAUTSYC0.AUTSYCCLI
                End If
                fgDetail.CellForeColor = wColor
            If Len(Trim(xZAUTSYC0.AUTSYCAUT)) > 3 Then
                fgDetail.Col = 1: fgDetail.Text = xZAUTSYC0.AUTSYCAUT
                fgDetail.CellForeColor = wColor
                fgDetail.Col = 2: fgDetail.Text = "2"
                fgDetail.CellForeColor = wColor
            Else
                fgDetail.Col = 2
                    fgDetail.Text = "1"
                fgDetail.CellForeColor = wColor
                fgDetail.Col = 3: fgDetail.Text = xZAUTSYC0.AUTSYCAUT
                fgDetail.CellForeColor = wColor
            End If
            fgDetail.Col = 4: fgDetail.Text = arrCLIENARA1(K)
            fgDetail.CellForeColor = wColor
            fgDetail.Col = 6: fgDetail.Text = Format(xZAUTSYC0.AUTSYCMON, "### ### ### ##0.00")
            fgDetail.CellForeColor = wColor
            fgDetail.Col = 7: fgDetail.Text = xZAUTSYC0.AUTSYCDEV
            fgDetail.CellForeColor = wColor
            fgDetail.Col = 8: fgDetail.Text = "Blocage : " & xZAUTSYC0.AUTSYCBLO
            fgDetail.CellForeColor = wColor
            fgDetail.Col = 12: fgDetail.Text = xZAUTSYC0.AUTSYCCET
            fgDetail.CellForeColor = wColor
            If xZAUTSYC0.AUTSYCDVL > 0 Then
                wAmj = xZAUTSYC0.AUTSYCDVL + 19000000
                fgDetail.Col = 13: fgDetail.Text = dateImp10(wAmj)
                fgDetail.CellForeColor = wColor
            End If

           For I = 0 To 13: fgDetail.Col = I: fgDetail.CellBackColor = mColor_W0: Next I
        End If
            
Next K

'============================================================================================================
fgDetail_Sort1 = 0: fgDetail_Sort2 = 2: fgDetail_Sort
fgDetail.Visible = True
Set rsAdo = Nothing
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgDetail_5_Surveillance_Blocage_Display()
Dim V, xWhere As String, xWhere2 As String, wAmj As String, K As Integer, I As Integer
Dim X As String, wColor As Long, wBackColor As Long
Dim blnBackColor As Boolean, wCOMPTEFON As String
Dim blnOk As Boolean
On Error GoTo Error_Handler

fgDetail.Visible = False
fgDetail_Reset
fgDetail.FormatString = "<Racine          |<Compte                                                         |<Code   |<Nature|<Intitulé                                                               " _
                      & "|<Echéance     |> Montant                  |<Devise  |> Montant                      |<     " _
                      & "|Resp |Pays |Etat|<Date validation   "
fgDetail.Left = 100
fgDetail.Top = 150
fgDetail.Width = 13000

fgDetail.Rows = 1
fgDetail.Row = 0
xWhere2 = ""
xWhere = "" '" and AUTSYCMON <> 0 "
If chkSelect_Options_4_AUTSICMON = "1" Then xWhere = ""
X = Trim(txtSelect_Options_4_CLIENACLI)
If X <> "" Then
    X = Format(Val(X), "0000000")
    xWhere = xWhere & " and AUTSYCCLI = '" & X & "'"
    xWhere2 = " and CLIENACLI = '" & X & "'"
End If

'================================================================================================================
X = "select count(*) from " & paramIBM_Library_SAB & ".ZAUTSYC0 " _
    & " where AUTSYCTYP = '1' " & xWhere
Set rsAdo = cnAdo.Execute(X)

ReDim arrZAUTSYC0(rsAdo(0) + 1), arrCLIENARA1(rsAdo(0) + 1)
arrZAUTSYC0_Nb = 0
'______________________________________________________________________________________________________
X = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0 , " & paramIBM_Library_SAB & ".ZCLIENA0" _
    & " where AUTSYCTYP = '1' and AUTSYCAUT = 'DEC' and AUTSYCCLI = CLIENACLI" & xWhere _
    & " order by AUTSYCCLI"

Set rsAdo = cnAdo.Execute(X)

Do While Not rsAdo.EOF
    arrZAUTSYC0_Nb = arrZAUTSYC0_Nb + 1
    Call rsZAUTSYC0_GetBuffer(rsAdo, xZAUTSYC0)
    arrZAUTSYC0(arrZAUTSYC0_Nb) = xZAUTSYC0
    arrCLIENARA1(arrZAUTSYC0_Nb) = rsAdo("CLIENARA1")

    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    
    wColor = vbBlue
    If xZAUTSYC0.AUTSYCFIN > 0 Then
        wAmj = xZAUTSYC0.AUTSYCFIN + 19000000
        fgDetail.Col = 5: fgDetail.Text = dateImp10(wAmj)
        If wAmj <= YBIATAB0_DATE_CPT_J Then wColor = vbMagenta
        fgDetail.CellForeColor = wColor
    End If
    fgDetail.Col = 0
    If xZAUTSYC0.AUTSYCGPE = "O" Then
        fgDetail.Text = xZAUTSYC0.AUTSYCCLI & "-G"
    Else
        fgDetail.Text = xZAUTSYC0.AUTSYCCLI
    End If
    fgDetail.CellForeColor = wColor
    'fgDetail.Col = 1: fgDetail.Text = rsAdo("CLIENACLI")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 2
        fgDetail.Text = "1"
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 3: fgDetail.Text = xZAUTSYC0.AUTSYCAUT
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 4: fgDetail.Text = rsAdo("CLIENARA1")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 6: fgDetail.Text = Format(xZAUTSYC0.AUTSYCMON, "### ### ### ##0.00")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 7: fgDetail.Text = xZAUTSYC0.AUTSYCDEV
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 8: fgDetail.Text = "Blocage : " & xZAUTSYC0.AUTSYCBLO
    fgDetail.CellForeColor = wColor
    
    fgDetail.Col = 10: fgDetail.Text = rsAdo("CLIENARES")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 11: fgDetail.Text = rsAdo("CLIENARSD")
    fgDetail.CellForeColor = wColor
        
    fgDetail.Col = 12: fgDetail.Text = xZAUTSYC0.AUTSYCCET
    fgDetail.CellForeColor = wColor
    If xZAUTSYC0.AUTSYCDVL > 0 Then
        wAmj = xZAUTSYC0.AUTSYCDVL + 19000000
        fgDetail.Col = 13: fgDetail.Text = dateImp10(wAmj)
        fgDetail.CellForeColor = wColor
    End If

    If xZAUTSYC0.AUTSYCBLO = 2 Then
        wBackColor = mColor_B0
    Else
        wBackColor = mColor_W0
    End If
    
    For I = 0 To 13: fgDetail.Col = I: fgDetail.CellBackColor = wBackColor: Next I

    rsAdo.MoveNext
Loop

'______________________________________________________________________________________________________
X = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0 , " & paramIBM_Library_SAB & ".ZCLIENA0" _
    & " where AUTSYCTYP = '1' and AUTSYCAUT <> 'DEC' and AUTSYCCLI = CLIENACLI" & xWhere _
    & " order by AUTSYCCLI"

Set rsAdo = cnAdo.Execute(X)

Do While Not rsAdo.EOF
    Call rsZAUTSYC0_GetBuffer(rsAdo, xZAUTSYC0)
    blnOk = False
    For K = 1 To arrZAUTSYC0_Nb
        If xZAUTSYC0.AUTSYCGPE = arrZAUTSYC0(K).AUTSYCGPE _
        And xZAUTSYC0.AUTSYCCLI = arrZAUTSYC0(K).AUTSYCCLI _
        And xZAUTSYC0.AUTSYCPER = arrZAUTSYC0(K).AUTSYCADR Then
            blnOk = True
            Exit For
        End If
    Next K
    
    If blnOk Then
        arrZAUTSYC0_Nb = arrZAUTSYC0_Nb + 1
        arrZAUTSYC0(arrZAUTSYC0_Nb) = xZAUTSYC0
        arrCLIENARA1(arrZAUTSYC0_Nb) = rsAdo("CLIENARA1")
    
        fgDetail.Rows = fgDetail.Rows + 1
        fgDetail.Row = fgDetail.Rows - 1
        
        wColor = vbBlue
        If xZAUTSYC0.AUTSYCFIN > 0 Then
            wAmj = xZAUTSYC0.AUTSYCFIN + 19000000
            fgDetail.Col = 5: fgDetail.Text = dateImp10(wAmj)
            If wAmj <= YBIATAB0_DATE_CPT_J Then wColor = vbMagenta
            fgDetail.CellForeColor = wColor
        End If
        fgDetail.Col = 0
    
        If xZAUTSYC0.AUTSYCGPE = "O" Then
            fgDetail.Text = xZAUTSYC0.AUTSYCCLI & "-G"
        Else
            fgDetail.Text = xZAUTSYC0.AUTSYCCLI
        End If
        fgDetail.CellForeColor = wColor
        'fgDetail.Col = 1: fgDetail.Text = rsAdo("CLIENACLI")
        fgDetail.CellForeColor = wColor
        fgDetail.Col = 2
            fgDetail.Text = "1"
        fgDetail.CellForeColor = wColor
        fgDetail.Col = 3: fgDetail.Text = xZAUTSYC0.AUTSYCAUT
        fgDetail.CellForeColor = wColor
        fgDetail.Col = 4: fgDetail.Text = rsAdo("CLIENARA1")
        fgDetail.CellForeColor = wColor
        fgDetail.Col = 6: fgDetail.Text = Format(xZAUTSYC0.AUTSYCMON, "### ### ### ##0.00")
        fgDetail.CellForeColor = wColor
        fgDetail.Col = 7: fgDetail.Text = xZAUTSYC0.AUTSYCDEV
        fgDetail.CellForeColor = wColor
        fgDetail.Col = 8: fgDetail.Text = "Blocage : " & xZAUTSYC0.AUTSYCBLO
        fgDetail.CellForeColor = wColor
        fgDetail.Col = 10: fgDetail.Text = rsAdo("CLIENARES")
        fgDetail.CellForeColor = wColor
        fgDetail.Col = 11: fgDetail.Text = rsAdo("CLIENARSD")
        fgDetail.CellForeColor = wColor
        fgDetail.Col = 12: fgDetail.Text = xZAUTSYC0.AUTSYCCET
        fgDetail.CellForeColor = wColor
        If xZAUTSYC0.AUTSYCDVL > 0 Then
            wAmj = xZAUTSYC0.AUTSYCDVL + 19000000
            fgDetail.Col = 13: fgDetail.Text = dateImp10(wAmj)
            fgDetail.CellForeColor = wColor
        End If
    
        If xZAUTSYC0.AUTSYCBLO = 2 Then
            wBackColor = mColor_B0
        Else
            wBackColor = mColor_W0
        End If
        
        For I = 0 To 13: fgDetail.Col = I: fgDetail.CellBackColor = wBackColor: Next I
    End If
    rsAdo.MoveNext
Loop


'______________________________________________________________________________________________________
'X = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0 , " & paramIBM_Library_SAB & ".ZCLIENA0 , " & paramIBM_Library_SAB & ".ZCOMPTE0 , " _
'    & paramIBM_Library_SAB & ".ZPLAN0" _
'    & " where AUTSYCTYP = '2'  and AUTSYCCLI = CLIENACLI and AUTSYCAUT = COMPTECOM and COMPTEFON <> '4' and COMPTEOBL = PLANCOOBL " & xWhere _
'    & " order by AUTSYCCLI, COMPTECOM"
X = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0 , " & paramIBM_Library_SAB & ".ZCLIENA0 , " & paramIBM_Library_SAB & ".ZCOMPTE0 " _
    & " where AUTSYCTYP = '2'  and AUTSYCCLI = CLIENACLI and AUTSYCAUT = COMPTECOM and COMPTEFON <> '4' " & xWhere _
    & " order by AUTSYCCLI, COMPTECOM"

Set rsAdo = cnAdo.Execute(X)

Do While Not rsAdo.EOF
    Call rsZAUTSYC0_GetBuffer(rsAdo, xZAUTSYC0)
    blnOk = False
    For K = 1 To arrZAUTSYC0_Nb
        If xZAUTSYC0.AUTSYCGPE = arrZAUTSYC0(K).AUTSYCGPE _
        And xZAUTSYC0.AUTSYCCLI = arrZAUTSYC0(K).AUTSYCCLI _
        And xZAUTSYC0.AUTSYCPER = arrZAUTSYC0(K).AUTSYCADR Then
            blnOk = True
            Exit For
        End If
    Next K
    
    If blnOk Then
            wCOMPTEFON = rsAdo("COMPTEFON")

            fgDetail.Rows = fgDetail.Rows + 1
            fgDetail.Row = fgDetail.Rows - 1
            
            wColor = vbBlue
            If xZAUTSYC0.AUTSYCFIN > 0 Then
                wAmj = xZAUTSYC0.AUTSYCFIN + 19000000
                fgDetail.Col = 5: fgDetail.Text = dateImp10(wAmj)
                If wAmj <= YBIATAB0_DATE_CPT_J Then wColor = vbMagenta
                fgDetail.CellForeColor = wColor
            End If
            fgDetail.Col = 0

            If rsAdo("AUTSYCGPE") = "O" Then
                fgDetail.Text = rsAdo("AUTSYCCLI") & "-G"
            Else
                fgDetail.Text = rsAdo("AUTSYCCLI")
            End If
            fgDetail.CellForeColor = wColor
            fgDetail.Col = 1: fgDetail.Text = rsAdo("COMPTECOM")
            fgDetail.CellForeColor = wColor
            fgDetail.Col = 2: fgDetail.Text = "2"
            fgDetail.CellForeColor = wColor
            'fgDetail.Col = 3: fgDetail.Text = rsAdo("PLANCOPRO")
            fgDetail.CellForeColor = wColor
            fgDetail.Col = 4: fgDetail.Text = rsAdo("COMPTEINT")
            fgDetail.CellForeColor = wColor
            fgDetail.Col = 6: fgDetail.Text = Format(xZAUTSYC0.AUTSYCMON, "### ### ### ##0.00")
            fgDetail.CellForeColor = wColor
            fgDetail.Col = 7: fgDetail.Text = xZAUTSYC0.AUTSYCDEV
            fgDetail.CellForeColor = wColor
            fgDetail.Col = 8
            fgDetail.Text = "cpt :" & wCOMPTEFON & " / " & "Blocage : " & xZAUTSYC0.AUTSYCBLO
            fgDetail.CellForeColor = wColor

            fgDetail.Col = 10: fgDetail.Text = rsAdo("CLIENARES")
            fgDetail.CellForeColor = wColor
            fgDetail.Col = 11: fgDetail.Text = rsAdo("CLIENARSD")
            fgDetail.CellForeColor = wColor
            
            fgDetail.Col = 12: fgDetail.Text = xZAUTSYC0.AUTSYCCET
            fgDetail.CellForeColor = wColor
            If xZAUTSYC0.AUTSYCDVL > 0 Then
                wAmj = xZAUTSYC0.AUTSYCDVL + 19000000
                fgDetail.Col = 13: fgDetail.Text = dateImp10(wAmj)
                fgDetail.CellForeColor = wColor
            End If
            
            blnBackColor = False
            wBackColor = mColor_W0

            Select Case wCOMPTEFON
                Case "0":
                        If xZAUTSYC0.AUTSYCBLO <> 0 Then blnBackColor = True
                Case "3"
                Case Else
                    If xZAUTSYC0.AUTSYCBLO = 0 Then blnBackColor = True
            End Select
            
            If blnBackColor Then
                'For I = 0 To 13: fgDetail.Col = I: fgDetail.CellBackColor = wBackColor: Next I
                fgDetail.Col = 8: fgDetail.CellBackColor = wBackColor
            End If
        End If

    
    rsAdo.MoveNext
Loop



'============================================================================================================
fgDetail_Sort1 = 0: fgDetail_Sort2 = 3: fgDetail_Sort
fgDetail.Visible = True
Set rsAdo = Nothing

Call lstErr_AddItem(lstErr, cmdContext, "Nb  : " & fgDetail.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub



Private Sub cmdSelect_SQL_YUPDLOG0()
Dim I As Integer, X As String
Dim wCli As String
Dim wCLIENACLI As String
Dim V
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

SSTab1.Tab = 2
xWhere = ""
X = Trim(txtUpdLog_CLIRGPCLI)
If X <> "" Then xWhere = " where UPDLOGTXT like '%" & X & "%'"
X = Trim(txtUpdLog_CLIRGPREG)
If X <> "" Then
    If xWhere = "" Then
        xWhere = " where UPDLOGTXT like '%" & X & "%'"
    Else
        xWhere = xWhere & " and UPDLOGTXT like '%" & X & "%'"
    End If
End If

If chkUpdLog_AmjMin.Value = "1" Then
    Call DTPicker_Control(txtUpdLog_AmjMin, wAMJMin)
    If wAMJMin <> "00000000" Then
        If xWhere = "" Then
            xWhere = " where UPDLOGAMJ = " & wAMJMin
        Else
            xWhere = xWhere & " and UPDLOGAMJ = " & wAMJMin
        End If
    End If
End If
If xWhere = "" Then
    MsgBox "Préciser au moins un critère de séléction", vbQuestion, "SAB_Client : Piste d'audit"
    Exit Sub
End If
Set rsAdo = Nothing

X = "select * from " & paramIBM_Library_SABSPE & ".YUPDLOG0 " & xWhere & " order by UPDLOGID"


Set rsAdo = cnAdo.Execute(X)
fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_YKYCDOSH()
Dim I As Integer, X As String
Dim V
Dim xWhere As String, xOrder As String
On Error GoTo Error_Handler

SSTab1.Tab = 2
xWhere = "": xOrder = ""
X = Trim(txtUpdLog_CLIRGPCLI)
If X <> "" Then xWhere = " where KYCDOSID like '%" & X & "%'": xOrder = " KYCDOSID ,"

If chkUpdLog_AmjMin.Value = "1" Then
    Call DTPicker_Control(txtUpdLog_AmjMin, wAMJMin)
    If wAMJMin <> "00000000" Then
        If xWhere = "" Then
            xWhere = " where KYCDOSUAMJ = " & wAMJMin
        Else
            xWhere = xWhere & " and KYCDOSUAMJ = " & wAMJMin
        End If
    End If
End If
If xWhere = "" Then
    MsgBox "Préciser au moins un critère de séléction", vbQuestion, "SAB_Client : Piste d'audit"
    Exit Sub
End If
Set rsAdo = Nothing

X = "select * from " & paramIBM_Library_SABSPE & ".YKYCDOSH " & xWhere _
  & " order by " & xOrder & " KYCDOSUAMJ , KYCDOSUHMS , KYCDOSUVER"


Set rsSab = cnsab.Execute(X)
fgSelect_YKYCDOSH_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub cmdSelect_SQL_2()
Dim I As Integer, X As String
Dim wCli As String
Dim wCLIENACLI As String
Dim V
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

arrZCLIGRP0_Index = 1

Select Case mSelect_SQL
    Case "2*": xWhere = " where CLIENAAGE = 1 and CLIENACLI > '9900000'"
    Case "2i": xWhere = " where CLIENAAGE = 1 and CLIENACLI > '9900000'and substring(CLIENARES , 1 , 1) = 'X'"
    Case Else: xWhere = " where CLIENAAGE = 1 and CLIENACLI > '9900000'and substring(CLIENARES , 1 , 1) <> 'X'"
End Select

xAnd = " and "
X = Trim(Mid$(cboSelect_CLIENACAT, 1, 3))
If X <> "" Then xWhere = xWhere & xAnd & "CLIENACAT = '" & X & "'": xAnd = " and "

X = Trim(txtSelect_CLIENARA1)
If X <> "" Then
    If IsNumeric(X) Then
        xWhere = xWhere & xAnd & "CLIENACLI like '%" & X & "%'": xAnd = " and "
    Else
        xWhere = xWhere & xAnd & "CLIENARA1 like '%" & X & "%'": xAnd = " and "
    End If
End If

tvwSelect.Nodes.Clear
tvwInverse.Nodes.Clear
lblSelect = ""
lblInverse = ""
Set rsAdo = Nothing

X = "select CLIENACLI,CLIENARA1,CLIENARA2 from " & paramIBM_Library_SAB & ".ZCLIENA0 " & xWhere & " order by CLIENACLI"


Set rsAdo = cnAdo.Execute(X)

Do While Not rsAdo.EOF
    wCLIENACLI = rsAdo("CLIENACLI")
    wCli = "CLI" & wCLIENACLI
    tvwSelect.Nodes.Add , , wCli, wCLIENACLI & "   " & Trim(rsAdo("CLIENARA1")) & "   " & Trim(rsAdo("CLIENARA2"))
    tvwSelect.Nodes(wCli).Sorted = True
    'tvwSelect_Display_ZCLIGRP0 wCLIENACLI

    rsAdo.MoveNext

Loop

fraSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_99()
Dim I As Integer, X As String
Dim wCli As String
Dim wCLIENACLI As String
Dim V
Dim xWhere As String, xAnd As String
Dim xRA1 As String

On Error GoTo Error_Handler

arrZCLIGRP0_Index = 1

xWhere = " where CLIENAAGE = 1 and CLIENACLI > '9900000' "
xAnd = " and "

X = Trim(txtUpdate_Select)
If X <> "" Then
    If IsNumeric(X) Then
        xWhere = xWhere & xAnd & "CLIENACLI like '%" & X & "%'": xAnd = " and "
    Else
        xWhere = xWhere & xAnd & "CLIENARA1 like '%" & X & "%'": xAnd = " and "
    End If
End If

tvwUpdate.Nodes.Clear
Set rsAdo = Nothing

X = "select CLIENACLI,CLIENARA1,CLIENARA2,CLIENARES from " & paramIBM_Library_SAB & ".ZCLIENA0 " & xWhere & " order by CLIENACLI"

Call FEU_ROUGE
Set rsAdo = cnAdo.Execute(X)
Call FEU_VERT
If Not rsAdo.EOF Then
    optUpdate_Add_Old.Enabled = True
Else
    optUpdate_Add_Old.Enabled = False
End If


Do While Not rsAdo.EOF
    wCLIENACLI = rsAdo("CLIENACLI")
    wCli = "CLI" & wCLIENACLI
    xRA1 = Trim(rsAdo("CLIENARA1")) & "  " & Trim(rsAdo("CLIENARA2"))
    If Mid$(rsAdo("CLIENARES"), 1, 1) = "X" Then xRA1 = "## " & LCase$(xRA1) & " ##"
    tvwUpdate.Nodes.Add , , wCli, wCLIENACLI & "   " & xRA1
    tvwUpdate.Nodes(wCli).Sorted = True

    rsAdo.MoveNext

Loop

fraSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_YKYCCTL()
Dim I As Integer, X As String
Dim wCli As String
Dim wCLIENACLI As String
Dim V
Dim xWhere As String, xAnd As String
Dim xRA1 As String, Nb1 As Long, K As Long, blnOk As Boolean, blnID_Match As Boolean

On Error GoTo Error_Handler
ReDim wKYCCTL(2000)
wKYCCTL_Nb = 0
X = "select CLIENACLI,CLIENARA1,CLIENARA2,CLIENARES,CLIENCPIE,CLILIBDA1 from " & paramIBM_Library_SAB & ".ZCLIENA0, " _
     & paramIBM_Library_SAB & ".ZCLIENC0, " & paramIBM_Library_SAB & ".ZCLILIB0 " _
     & " where CLIENcCLI = clienacli and clilibcli = clienacli and substring(CLIENARES , 1, 1) = 'R'" _
     & " and cliencpie <> ' ' and clienacli > '0010000' and clienacli < '0100000'" _
     & " and CLIENACOL <> 1 "
If mSelect_SQL = "KYC Ctl PP" Then
    X = X & " and substring(CLIENACAT , 1, 1) = 'P' order by CLIENACLI"
Else
    X = X & " order by CLIENACLI"
End If

Set rsAdo = cnAdo.Execute(X)


Do While Not rsAdo.EOF
    wKYCCTL_Nb = wKYCCTL_Nb + 1
    wKYCCTL(wKYCCTL_Nb).STA = "? GSOP"
    
    
    wKYCCTL(wKYCCTL_Nb).CLIENACLI = rsAdo("CLIENACLI")
    wKYCCTL(wKYCCTL_Nb).CLIENARA = Trim(rsAdo("CLIENARA1")) & " " & Trim(rsAdo("CLIENARA2"))
    wKYCCTL(wKYCCTL_Nb).CLIENARES = rsAdo("CLIENARES")
    wKYCCTL(wKYCCTL_Nb).CLIENCPIE = rsAdo("CLIENCPIE")
    wKYCCTL(wKYCCTL_Nb).CLILIBDA1 = rsAdo("CLILIBDA1")
    If wKYCCTL(wKYCCTL_Nb).CLILIBDA1 > 0 Then wKYCCTL(wKYCCTL_Nb).CLILIBDA1 = wKYCCTL(wKYCCTL_Nb).CLILIBDA1 + 19000000

    rsAdo.MoveNext

Loop
Nb1 = wKYCCTL_Nb


X = "select KYCDOSID,KYCDOSSEQ2,KYCDOSDAMJ,KYCDOSDECH,KYCDOSDLIB,CLIENARA1,CLIENARA2,CLIENARES from " & paramIBM_Library_SABSPE & ".YKYCDOS0, " & paramIBM_Library_SAB & ".ZCLIENA0 " _
  & " where kycdosnat = ' ' and kycdosseq2 in (20, 21 , 22 , 23 , 220) and clienacli = kycdosid and substring (CLIENARES , 1 , 1 ) = 'R'" _
  & " and CLIENACOL <> 1 "

If mSelect_SQL = "KYC Ctl PP" Then
    X = X & " and substring(CLIENACAT , 1, 1) = 'P' order by KYCDOSID , KYCDOSSEQ2"
Else
    X = X & " order by KYCDOSID , KYCDOSSEQ2"
End If

Set rsAdo = cnAdo.Execute(X)


Do While Not rsAdo.EOF

    wCLIENACLI = rsAdo("KYCDOSID")
    blnOk = False: blnID_Match = False
    For K = 1 To Nb1
        If wKYCCTL(K).CLIENACLI = wCLIENACLI Then
            blnID_Match = True
            'If wKYCCTL(K).KYCDOSSEQ2 = 0 Then
                Select Case rsAdo("KYCDOSSEQ2")
                    Case 22, 220: If wKYCCTL(K).CLIENCPIE = "PA" Then blnOk = True
                    Case 21: If wKYCCTL(K).CLIENCPIE = "CI" Then blnOk = True
                    Case 23: If wKYCCTL(K).CLIENCPIE = "CS" Or wKYCCTL(K).CLIENCPIE = "RT" Then blnOk = True
                End Select
                If blnOk Then
                    If wKYCCTL(K).CLILIBDA1 = rsAdo("KYCDOSDAMJ") Then
                        wKYCCTL(K).STA = ""
                    Else
                        wKYCCTL(K).STA = "#"
                    End If
                    Exit For
                End If
            'End If
        End If
        
    Next K
    
    If Not blnOk Then
        wKYCCTL_Nb = wKYCCTL_Nb + 1
        K = wKYCCTL_Nb
        If blnID_Match Then
            wKYCCTL(wKYCCTL_Nb).STA = "! doc #"
        Else
            wKYCCTL(wKYCCTL_Nb).STA = "? SAB KYC"
        End If
        wKYCCTL(wKYCCTL_Nb).CLIENACLI = rsAdo("KYCDOSID")
        wKYCCTL(wKYCCTL_Nb).CLIENARA = Trim(rsAdo("CLIENARA1")) & " " & Trim(rsAdo("CLIENARA2"))
        wKYCCTL(wKYCCTL_Nb).CLIENARES = rsAdo("CLIENARES")
   End If
    
    wKYCCTL(K).KYCDOSSEQ2 = rsAdo("KYCDOSSEQ2")
    wKYCCTL(K).KYCDOSDAMJ = rsAdo("KYCDOSDAMJ")
    wKYCCTL(K).KYCDOSDECH = rsAdo("KYCDOSDECH")
    wKYCCTL(K).KYCDOSDLIB = rsAdo("KYCDOSDLIB")
    
   rsAdo.MoveNext

Loop

 
Set rsAdo = Nothing

fgYKYCCTL_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub cmdSelect_Ok_Click()
Dim V
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAB_Client_cmdSelect_Ok ........"): DoEvents

'cmdSelect_Reset
fgSelect.Visible = False
currentAction = "cmdSelect_Ok"
Select Case mSelect_SQL
    Case "1": cmdSelect_SQL_1
    Case "1r": cmdSelect_SQL_1r
    Case "2", "2i", "2*": cmdSelect_SQL_2
    Case "3": cmdSelect_SQL_3
    Case "4": cmdSelect_SQL_4
    Case "4!e": cmdSelect_SQL_4_Surveillance_Echelles
    Case "5": cmdSelect_SQL_5_Surveillance_Blocage
    Case "Xf": wFile_Orig = "fiche Clients actifs": cmdSelect_SQL_Xf
    Case "Xman": wFile_Orig = "fiche Mandataires": cmdSelect_SQL_Xf
    Case "Xman#": wFile_Orig = "liste Mandataires à clôturer": cmdSelect_SQL_XmanZ
    Case "Xg": cmdSelect_SQL_XG
    Case "Xp": ZCLIENA0_Export
    Case "Xa": ZADRESS0_Exportation False
    Case "Xa*": ZADRESS0_Exportation True
    Case "Xz": cmdSelect_SQL_Xz
    Case "XzR": cmdSelect_SQL_XzR
    Case "Xgsop@":
        Call MsgBox("Xgsop@ interdit : lancer @SAB_CLIENT en fin de mois", vbInformation, "SAB_Client")
        'cmdSelect_SQL_Xgsop_Auto 'cmdSelect_SQL_Xgsop_Auto_Reprise
    Case "Xgsop": cmdSelect_SQL_Xgsop
    Case "KYC gsop": cmdSelect_SQL_YKYCDOS0
    Case "KYC ech": cmdSelect_SQL_YKYCDOS0_Ech
    Case "KYC Ctl", "KYC Ctl PP": cmdSelect_SQL_YKYCCTL
    Case "KYC Releve": cmdSelect_SQL_ZRELEVE0
    Case "JPL":
        'cmdSelect_SQL_ZRELEVE0
        'cmdSelect_SQL_YKYCCTL
                'paramKYCDOSECH_Reprise
                'cmdSelect_JPL_YKYCSTA0
                'cmdSelect_JPL
                'cmdSelect_SQL_Xgsop_Auto_Reprise
End Select
    
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_Client_cmdSelect_Ok "): DoEvents
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
  

End Sub

Public Sub cmdSelect_SQL_Xf()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim X As String, K As Long
Dim wFilex As String, wFile As String, xSQL As String

On Error GoTo Error_Handler
'______________________________________________'

wFile = Trim("C:\Temp\CLI " & wFile_Orig & " " & dateImp_Amj(YBIATAB0_DATE_CPT_J) & ".xlsx")

If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "CLI " & wFile_Orig & " : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If

'_________________________________________

If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile


Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "ZCLIENA0"
    .Subject = ""
End With

appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add
'Set wbExcel = wbExcel.Sheets(1)

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "CLI- " & dateImp10(YBIATAB0_DATE_CPT_J)


With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .HorizontalAlignment = Excel.xlHAlignCenter
    .WrapText = True
    .Font.Size = 9
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 85
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14CLI : " & wFile_Orig & " en date du " & dateImp10_S(YBIATAB0_DATE_CPT_J) _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$L1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


wsExcel.Columns(1).ColumnWidth = 3: wsExcel.Cells(1, 1) = "Bq"
wsExcel.Columns(2).ColumnWidth = 3: wsExcel.Cells(1, 2) = "Ag"
wsExcel.Columns(3).ColumnWidth = 8: wsExcel.Cells(1, 3) = "Client"
wsExcel.Columns(4).ColumnWidth = 6: wsExcel.Cells(1, 4) = "R.Com"
wsExcel.Columns(5).ColumnWidth = 6: wsExcel.Cells(1, 5) = "Collectif"
wsExcel.Columns(6).ColumnWidth = 6: wsExcel.Cells(1, 6) = "Cotation"
wsExcel.Columns(7).ColumnWidth = 6: wsExcel.Cells(1, 7) = "Int.Chq"
wsExcel.Columns(8).ColumnWidth = 6: wsExcel.Cells(1, 8) = "Douteux"
wsExcel.Columns(9).ColumnWidth = 6: wsExcel.Cells(1, 9) = "sélection"
wsExcel.Columns(10).ColumnWidth = 6: wsExcel.Cells(1, 10) = "Lien"
wsExcel.Columns(11).ColumnWidth = 6: wsExcel.Cells(1, 11) = "Etat"
wsExcel.Columns(12).ColumnWidth = 6: wsExcel.Cells(1, 12) = "Eco"
wsExcel.Columns(13).ColumnWidth = 6: wsExcel.Cells(1, 13) = "Cat"
wsExcel.Columns(14).ColumnWidth = 6: wsExcel.Cells(1, 14) = "P.Nat"
wsExcel.Columns(15).ColumnWidth = 6: wsExcel.Cells(1, 15) = "P.Rés"

wsExcel.Columns(16).ColumnWidth = 12: wsExcel.Cells(1, 16) = "Sigle": wsExcel.Columns(16).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(17).ColumnWidth = 30: wsExcel.Cells(1, 17) = "Intitulé 1/2": wsExcel.Columns(17).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(18).ColumnWidth = 30: wsExcel.Cells(1, 18) = "Intitulé 2/2": wsExcel.Columns(18).HorizontalAlignment = Excel.xlHAlignLeft


Select Case mSelect_SQL
    Case "Xf":
        wsExcel.Columns(19).ColumnWidth = 8: wsExcel.Cells(1, 19) = "APE": wsExcel.Columns(19).HorizontalAlignment = Excel.xlHAlignLeft
        wsExcel.Columns(20).ColumnWidth = 15: wsExcel.Cells(1, 20) = "SIREN": wsExcel.Columns(20).HorizontalAlignment = Excel.xlHAlignLeft
        wsExcel.Columns(21).ColumnWidth = 50: wsExcel.Cells(1, 21) = "libellé APE": wsExcel.Columns(21).HorizontalAlignment = Excel.xlHAlignLeft
        For K = 1 To 21
            wsExcel.Cells(1, K).Interior.Color = mColor_GB
            wsExcel.Cells(1, K).Font.Color = mColor_Z0
        Next
        cmdSelect_SQL_Xf_Detail
        Set wsExcel = wbExcel.Sheets(2): cmdSelect_SQL_Xf_Table "Eco", 1
        Set wsExcel = wbExcel.Sheets(3): cmdSelect_SQL_Xf_Table "Eta", 5
        Set wsExcel = wbExcel.Sheets(4): cmdSelect_SQL_Xf_Table "Cat", 8
        Set wsExcel = wbExcel.Sheets(5): cmdSelect_SQL_Xf_Table "APE", 2
        Set wsExcel = wbExcel.Sheets(6): cmdSelect_SQL_Xf_Table "CSP", 121
    Case "Xman":
        wsExcel.Columns(8).ColumnWidth = 11: wsExcel.Cells(1, 8) = "D. Naissance"
        wsExcel.Columns(9).ColumnWidth = 11: wsExcel.Cells(1, 9) = "Siren"

        wsExcel.Columns(19).ColumnWidth = 30: wsExcel.Cells(1, 19) = "Adresse 1": wsExcel.Columns(19).HorizontalAlignment = Excel.xlHAlignLeft
        wsExcel.Columns(20).ColumnWidth = 30: wsExcel.Cells(1, 20) = "Adresse 2": wsExcel.Columns(20).HorizontalAlignment = Excel.xlHAlignLeft
        wsExcel.Columns(21).ColumnWidth = 30: wsExcel.Cells(1, 21) = "Adresse 3": wsExcel.Columns(21).HorizontalAlignment = Excel.xlHAlignLeft
        wsExcel.Columns(22).ColumnWidth = 10: wsExcel.Cells(1, 22) = "Code postal": wsExcel.Columns(22).HorizontalAlignment = Excel.xlHAlignLeft
        wsExcel.Columns(23).ColumnWidth = 25: wsExcel.Cells(1, 23) = "Ville": wsExcel.Columns(23).HorizontalAlignment = Excel.xlHAlignLeft
        wsExcel.Columns(24).ColumnWidth = 30: wsExcel.Cells(1, 24) = "Pays": wsExcel.Columns(24).HorizontalAlignment = Excel.xlHAlignLeft
        For K = 1 To 24
            wsExcel.Cells(1, K).Interior.Color = mColor_GB
            wsExcel.Cells(1, K).Font.Color = mColor_Z0
        Next
        wsExcel.PageSetup.Zoom = 75
        cmdSelect_SQL_Xf_Mandataires
End Select
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
Public Sub cmdSelect_SQL_XmanZ()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim X As String, K As Long
Dim wFilex As String, wFile As String, xSQL As String

Dim manCLIENACLI As String, manRow As Long, mCLIENACLI_Actif As String, blnActif As Boolean
On Error GoTo Error_Handler
'______________________________________________'

wFile = Trim("C:\Temp\CLI " & wFile_Orig & " " & dateImp_Amj(YBIATAB0_DATE_CPT_J) & ".xlsx")

If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "CLI " & wFile_Orig & " : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If

'_________________________________________

If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile


Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "ZCLIENA0"
    .Subject = ""
End With

appExcel.Worksheets.Add
appExcel.Worksheets.Add
appExcel.Worksheets.Add
'Set wbExcel = wbExcel.Sheets(1)

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "CLI- " & dateImp10(YBIATAB0_DATE_CPT_J)


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
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14CLI : " & wFile_Orig & " en date du " & dateImp10_S(YBIATAB0_DATE_CPT_J) _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$L1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


wsExcel.Columns(1).ColumnWidth = 10: wsExcel.Cells(1, 1) = "Mandataires"
wsExcel.Columns(2).ColumnWidth = 45: wsExcel.Cells(1, 2) = "Intitulé"
wsExcel.Columns(3).ColumnWidth = 10: wsExcel.Cells(1, 3) = "Client"
wsExcel.Columns(4).ColumnWidth = 45: wsExcel.Cells(1, 4) = "Intitulé"
wsExcel.Columns(5).ColumnWidth = 13: wsExcel.Cells(1, 5) = "!! Client actif"
For K = 1 To 5
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next


'__________________________________________________________________________________
X = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 left outer join " & paramIBM_Library_SAB & ".ZCLIGRP0 " _
  & " on clienacli = cligrpreg" _
  & " where clienacli like '99%'" _
  & " and substring(CLIENARES , 1 , 1) <> 'X' " _
  & " order by clienacli"

Set rsSab = cnsab.Execute(X)

wRow = 1
Do While Not rsSab.EOF
    xZCLIENA0.CLIENACLI = rsSab("CLIENACLI")
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "> " & xZCLIENA0.CLIENACLI): DoEvents
    
    If manCLIENACLI <> xZCLIENA0.CLIENACLI Then
        If mCLIENACLI_Actif <> "" Then
            If manRow > 0 Then
                wsExcel.Cells(wRow, 5) = mCLIENACLI_Actif
                wsExcel.Cells(wRow, 5).Interior.Color = mColor_W1
            End If
        End If

        manCLIENACLI = xZCLIENA0.CLIENACLI
        manRow = 0
        blnActif = False
        mCLIENACLI_Actif = ""
    End If
    
    If IsNull(rsSab("CLIGRPCLI")) Then
        wRow = wRow + 1
        wsExcel.Cells(wRow, 1) = xZCLIENA0.CLIENACLI
        wsExcel.Cells(wRow, 1).Font.Color = vbMagenta
        wsExcel.Cells(wRow, 2) = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
        wsExcel.Cells(wRow, 2).Font.Color = vbMagenta
        wsExcel.Cells(wRow, 4) = "sans lien"
         wsExcel.Cells(wRow, 4).Font.Color = vbMagenta
   Else
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
         & " where clienacli = '" & rsSab("CLIGRPCLI") & "'" _
         & " and comptefon <> '4' "
    
        Set rsAdo = cnAdo.Execute(xSQL)
        
        If Not rsAdo.EOF Then
            blnActif = True
            mCLIENACLI_Actif = rsSab("CLIGRPCLI")
        Else
           wRow = wRow + 1
           If manRow = 0 Then manRow = wRow
           
            wsExcel.Cells(wRow, 1) = xZCLIENA0.CLIENACLI
            wsExcel.Cells(wRow, 2) = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
            wsExcel.Cells(wRow, 3) = rsSab("CLIGRPCLI")
            xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
             & " where clienacli = '" & rsSab("CLIGRPCLI") & "'"
            Set rsAdo = cnAdo.Execute(xSQL)
            If Not rsAdo.EOF Then
                wsExcel.Cells(wRow, 4) = Trim(rsAdo("CLIENARA1")) & " " & Trim(rsAdo("CLIENARA2"))
            End If
            If blnActif Then
                wsExcel.Cells(wRow, 5) = mCLIENACLI_Actif
                wsExcel.Cells(wRow, 5).Interior.Color = mColor_W1
            End If
            
        End If
    End If
    
    rsSab.MoveNext
Loop



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

Public Sub cmdSelect_SQL_Xz()
On Error GoTo Error_Handler
Dim X As String, K As Long
Dim wFilex As String, wFile As String, xSQL As String

On Error GoTo Error_Handler
'______________________________________________'

K = 18
If Not blnAuto Then

    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & "Sélectionner les comptes n'ayant pas fonctionné depuis 18 mois" _
        & vbCrLf & "     =========================", "CLI Comptes client à clôturer", K)
    If Trim(X) <> "" Then K = Val(X)
End If

wFile = Trim("C:\Temp\CLI Comptes client à clôturer " & dateImp_Amj(YBIATAB0_DATE_CPT_J) & "(" & K & " mois).xlsx")

If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "CLI Comptes client à clôturer : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
    
    
End If

'_________________________________________
cmdSelect_SQL_Xz_Init "CLI", wFile

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents

End Sub
Public Sub cmdSelect_SQL_Xgsop()
On Error GoTo Error_Handler
Dim X As String, K As Long
Dim wFilex As String, xSQL As String
Dim rsSab_RES As New ADODB.Recordset

On Error GoTo Error_Handler
'______________________________________________'
If blnAuto Then
     wFile_Orig = Trim("C:\Temp\GSOP reporting " & YBIATAB0_DATE_CPT_J)
Else

    If cboSelect_Options_Xgsop_CLIENARES.Text = "Archives" Then
        blnXgsop_Archive = True
        wFile_Orig = Trim("C:\Temp\GSOP Archive du  " & cboSelect_Options_Xgsop_Archive & " " & DSys & "_" & time_Hms)
    Else
        blnXgsop_Archive = False
        wFile_Orig = Trim("C:\Temp\GSOP reporting " & DSys & "_" & time_Hms)
    End If


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
'02/09/2013 DENIS ROSILLETTE
'remplacement de la ligne suivante, mise en commentaire
'If Trim(Dir(wFile_Orig)) <> "" Then Kill wFile_Orig
If Trim(Dir(wFile_Orig & ".xlsx")) <> "" Then
    Kill wFile_Orig & ".xlsx"
End If
'_________________________________________

If blnXgsop_Archive Then
    Call lstErr_AddItem(lstErr, cmdContext, "GSOP Archive"): DoEvents
    Call cmdSelect_SQL_Xgsop_Archive

Else
    paramXgsop_Init
    
    
    Select Case cboSelect_Options_Xgsop_CLIENARES
        Case "*":
            Call lstErr_AddItem(lstErr, cmdContext, "GSOP"): DoEvents
            Call cmdSelect_SQL_Xgsop_RES("")
        Case "* + R"
    
            Call lstErr_AddItem(lstErr, cmdContext, "GSOP"): DoEvents
            Call cmdSelect_SQL_Xgsop_RES("")
            
            X = "select distinct CLIENARES from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
              & " order by CLIENARES"
            
            Set rsSab_RES = cnsab.Execute(X)
            
            Do While Not rsSab_RES.EOF
                Call lstErr_AddItem(lstErr, cmdContext, "Responsable : " & rsSab_RES("CLIENARES")): DoEvents
                Call cmdSelect_SQL_Xgsop_RES(rsSab_RES("CLIENARES"))
                rsSab_RES.MoveNext
            Loop
        Case Else
                Call cmdSelect_SQL_Xgsop_RES(cboSelect_Options_Xgsop_CLIENARES.Text)
    End Select
End If


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
appExcel.Quit

End Sub

Private Sub cmdParam_Add_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

X = Trim(txtParam_Id)
If X = "" Then
    Call MsgBox("Préciser l'identifiant", vbCritical, "BIA_GAFI: paramétrage")
Else
    New_YBIATAB0 = Old_YBIATAB0
    Select Case lstParam_K
        Case "3": New_YBIATAB0.BIATABK2 = X
        Case Else:    New_YBIATAB0.BIATABK2 = Format$(Val(X), "0000000")
    End Select
    New_YBIATAB0.BIATABTXT = usrName_UCase & " " & dateImp10(DSys) & " " & Time
    If fgParam_Display_Lib(New_YBIATAB0.BIATABK2) = "?" Then
        Call MsgBox("identifiant inconnu : " & New_YBIATAB0.BIATABK2, vbCritical, "BIA_GAFI: paramétrage")
    Else
        If IsNull(Parametrage_New) Then txtParam_Id = "": fgParam_Display
    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Delete_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass
X = Trim(txtParam_Id)
If X = "" Then
    Call MsgBox("Préciser l'identifiant à supprimer", vbCritical, "BIA_GAFI : paramétrage")
Else
    New_YBIATAB0 = Old_YBIATAB0
    New_YBIATAB0.BIATABK2 = X
    Old_YBIATAB0.BIATABK2 = X
    If IsNull(Parametrage_Delete) Then txtParam_Id = "": fgParam_Display
End If


Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdParam_Quit_Click()
fraParam_Update.Visible = False

End Sub

Public Sub cmdSelect_SQL_Xgsop_Detail(xWhere As String)
On Error GoTo Error_Handler
Dim X As String, K As Long
Dim xSQL As String

On Error GoTo Error_Handler
'______________________________________________'

Call typeXgsop_Init(mXgsop)


'__________________________________________________________________________________
X = "select * from " & paramIBM_Library_SAB & ".ZTITULA0 , " & paramIBM_Library_SAB & ".ZCLIENA0 , " _
    & paramIBM_Library_SAB & ".ZCOMPTE0 , " & paramIBM_Library_SAB & ".ZPLAN0 " _
   & " where clienacli = titulacli " _
  & " and titulacom = comptecom " _
  & " and compteobl = PLANCOOBL " _
  & xWhere _
  & " order by clienacli"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    If mXgsop.CLIENACLI <> rsSab("CLIENACLI") Then
        Call cmdSelect_SQL_Xgsop_Cumul

        mXgsop.CLIENACLI = rsSab("CLIENACLI")
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "> " & mXgsop.CLIENACLI): DoEvents
        mXgsop.CLIENAETA = rsSab("CLIENAETA")
        mXgsop.CLIENACAT = rsSab("CLIENACAT")
        mXgsop.CLIENACOL = rsSab("CLIENACOL")
        mXgsop.CLIENARA1 = rsSab("CLIENARA1")
        mXgsop.CLIENARA2 = rsSab("CLIENARA2")
        mXgsop.CLIENARES = rsSab("CLIENARES")
        mXgsop.CLIENARSD = rsSab("CLIENARSD")
        mXgsop.CLIENANAT = rsSab("CLIENANAT")
    End If
    
    X = rsSab("PLANCOPRO")
    If rsSab("COMPTEFON") <> 4 And InStr(mXgsop.PLANCOPRO, X) = 0 Then mXgsop.PLANCOPRO = mXgsop.PLANCOPRO & X & " "
    
    If InStr(paramXgsop_PLANCOPRO, X) > 0 Then
        If rsSab("COMPTEFON") = 4 Then
            mXgsop.CAV_Clos = mXgsop.CAV_Clos + 1
            If rsSab("COMPTECLO") > mXgsop.COMPTECLO Then mXgsop.COMPTECLO = rsSab("COMPTECLO")
        Else
            If rsSab("TITULATPR") = 0 Then
                mXgsop.CAV_Client = mXgsop.CAV_Client + 1
            Else
                mXgsop.CAV_Tiers = mXgsop.CAV_Tiers + 1
            End If
            
        End If
    Else
        If rsSab("COMPTEFON") = 4 Then
            mXgsop.Tech_Clos = mXgsop.Tech_Clos + 1
            If rsSab("COMPTECLO") > mXgsop.COMPTECLO Then mXgsop.COMPTECLO = rsSab("COMPTECLO")
        Else
            If rsSab("TITULATPR") = 0 Then
                mXgsop.Tech_Client = mXgsop.Tech_Client + 1
            Else
                mXgsop.Tech_Tiers = mXgsop.Tech_Tiers + 1
            End If
            
        End If
    End If
    
    rsSab.MoveNext
Loop

Call cmdSelect_SQL_Xgsop_Cumul


'__________________________________________________________________________________
Exit_sub:
'__________________________________________________________________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents

End Sub

Public Sub cmdSelect_SQL_Xgsop_Detail_6(xWhere As String)
On Error GoTo Error_Handler
Dim X As String, K As Long
Dim xSQL As String

On Error GoTo Error_Handler
'______________________________________________'

Call typeXgsop_Init(mXgsop)


'__________________________________________________________________________________
X = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0" _
   & " where clienacli > '0010000' and  clienacli < '0070000' " _
   & " and clienacli not in (Select TITULACLI from " & paramIBM_Library_SAB & ".ZTITULA0)" _
   & xWhere _
   & " order by clienacli"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    
    If mXgsop.CLIENACLI <> rsSab("CLIENACLI") Then
        Call cmdSelect_SQL_Xgsop_Cumul

        mXgsop.CLIENACLI = rsSab("CLIENACLI")
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "> " & mXgsop.CLIENACLI): DoEvents
        mXgsop.CLIENAETA = rsSab("CLIENAETA")
        mXgsop.CLIENACAT = rsSab("CLIENACAT")
        mXgsop.CLIENACOL = rsSab("CLIENACOL")
        mXgsop.CLIENARA1 = rsSab("CLIENARA1")
        mXgsop.CLIENARA2 = rsSab("CLIENARA2")
        mXgsop.CLIENARES = rsSab("CLIENARES")
        mXgsop.CLIENARSD = rsSab("CLIENARSD")
        mXgsop.CLIENANAT = rsSab("CLIENANAT")
    End If
    
    
    rsSab.MoveNext
Loop

Call cmdSelect_SQL_Xgsop_Cumul


'__________________________________________________________________________________
Exit_sub:
'__________________________________________________________________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents

End Sub



Public Sub cmdSelect_SQL_XzR()
On Error GoTo Error_Handler
Dim X As String, K As Long
Dim wFilex As String, wFile As String, xSQL As String
On Error GoTo Error_Handler
'______________________________________________'

K = 18
If Not blnAuto Then

    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & "Sélectionner les comptes n'ayant pas fonctionné depuis 18 mois" _
        & vbCrLf & "     =========================", "Comptes client à clôturer", K)
    If Trim(X) <> "" Then K = Val(X)
End If

wFile = Trim("C:\Temp\CLI Comptes client à clôturer " & dateImp_Amj(YBIATAB0_DATE_CPT_J) & "(" & K & " mois).xlsx")

'If Not blnAuto Then
'    x = InputBox("par défaut : " _
'        & vbCrLf & "     =========================" & vbCrLf & wFile _
'        & vbCrLf & "     =========================", "Comptes client à clôturer : nom du fichier d'exportation", wFile)
'    If Trim(x) = "" Then Exit Sub
'    wFilex = Trim(x)
    '______________________________________________
'    If wFile <> wFilex Then
'        wFile = wFilex
'    End If
    
    
'End If

wFile = Trim("C:\Temp\CLI Comptes client à clôturer " & dateImp_Amj(YBIATAB0_DATE_CPT_J) & "(" & K & " mois).xlsx")

'_________________________________________

X = "select count(DISTINCT CLIENARES)  from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
  & " where clienacli <> ''" _
  & " and COMPTEFON <> '4' "

Set rsSab = cnsab.Execute(X)
K = rsSab(0) + 1

'ReDim arrCLIENARES(K) As String
arrCLIENARES_Nb = 0


X = "select DISTINCT CLIENARES  from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
  & " where clienacli <> ''" _
  & " and COMPTEFON <> '4' " _
  & " order by CLIENARES"

Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    arrCLIENARES_Nb = arrCLIENARES_Nb + 1
    arrCLIENARES(arrCLIENARES_Nb) = rsSab("CLIENARES")
    rsSab.MoveNext
Loop

For K = 1 To arrCLIENARES_Nb

    cmdSelect_SQL_Xz_Init arrCLIENARES(K), wFile
Next K
'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents

End Sub


Public Sub cmdSelect_SQL_Xz_Init(lCLIENARES As String, lFile As String)
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim X As String, K As Long
Dim xSQL As String, wFile As String

On Error GoTo Error_Handler
'______________________________________________'

K = 18
'_________________________________________
wFile = Replace(lFile, "\CLI", "\" & lCLIENARES)

If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile
wAMJMin = dateElp("MoisAdd", -K, YBIATAB0_DATE_CPT_J)
wIBM_AmjMin = wAMJMin - 19000000

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "ZCLIENA0"
    .Subject = ""
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = lCLIENARES & " - " & dateImp10(YBIATAB0_DATE_CPT_J)


With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .HorizontalAlignment = Excel.xlHAlignCenter
    .WrapText = True
    .Font.Size = 8
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 100
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14CLI : Comptes 'client' à clôturer" _
                                & vbCr & "&B&U&10(tous les comptes du client sont soldés depuis " & K & " mois, édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$G1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


wsExcel.Columns(1).ColumnWidth = 7: wsExcel.Cells(1, 1) = "Racine"
wsExcel.Columns(2).ColumnWidth = 6: wsExcel.Cells(1, 2) = "Nature"
wsExcel.Columns(3).ColumnWidth = 20: wsExcel.Cells(1, 3) = "Compte": wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(4).ColumnWidth = 6: wsExcel.Cells(1, 4) = "Devise"
wsExcel.Columns(5).ColumnWidth = 12: wsExcel.Cells(1, 5) = "Date dernier mouvement"
wsExcel.Columns(6).ColumnWidth = 12: wsExcel.Cells(1, 6) = "Date d'ouverture"
wsExcel.Columns(7).ColumnWidth = 64: wsExcel.Cells(1, 7) = "Intitulé": wsExcel.Columns(7).HorizontalAlignment = Excel.xlHAlignLeft

mXls1_Col = 7
For K = 1 To mXls1_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next

cmdSelect_SQL_Xz_Detail lCLIENARES

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

'=======================================================================================
If lCLIENARES <> "CLI" Then
    If mXls1_Row = 1 Then
        If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile
    End If
End If

'=======================================================================================
'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée " & lCLIENARES): DoEvents

End Sub

Public Sub cmdSelect_SQL_Xgsop_Init_2(lCLIENARES As String)
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim X As String, K As Long
Dim xSQL As String, wFile As String

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
wsExcel.PageSetup.Zoom = 70
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14GSOP : reporting clientèle  " & lCLIENARES _
                                & "&B&U&10     ( édité le " & dateImp10(DSys) & " " & Time & ")" _
                                & vbCr & "(X : 1-PP, 2-PM, 3-BQ, 4-Autres)" _
                                & "    (Y : 1-Clients, 2-Techniques, 3-Tiers, 4-Non réclamés, 5-hors GSOP, 6-Racines sans compte)"
wsExcel.PageSetup.PrintTitleRows = "$A1:$G1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


wsExcel.Columns(1).ColumnWidth = 3: wsExcel.Cells(1, 1) = "X"
wsExcel.Columns(2).ColumnWidth = 3: wsExcel.Cells(1, 2) = "Y"
wsExcel.Columns(3).ColumnWidth = 6: wsExcel.Cells(1, 3) = "Racine"
wsExcel.Columns(4).ColumnWidth = 7: wsExcel.Cells(1, 4) = "Collectif"
wsExcel.Columns(5).ColumnWidth = 8: wsExcel.Cells(1, 5) = "Code Etat"
wsExcel.Columns(6).ColumnWidth = 8: wsExcel.Cells(1, 6) = "Catégorie"
wsExcel.Columns(7).ColumnWidth = 5: wsExcel.Cells(1, 7) = "CAV Client": wsExcel.Columns(7).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(8).ColumnWidth = 5: wsExcel.Cells(1, 8) = "CAV Tiers": wsExcel.Columns(8).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(9).ColumnWidth = 5: wsExcel.Cells(1, 9) = "CAV clos": wsExcel.Columns(9).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(10).ColumnWidth = 5: wsExcel.Cells(1, 10) = "Tech Client": wsExcel.Columns(10).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(11).ColumnWidth = 5: wsExcel.Cells(1, 11) = "Tech Tiers": wsExcel.Columns(11).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(12).ColumnWidth = 5: wsExcel.Cells(1, 12) = "Tech clos": wsExcel.Columns(12).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(13).ColumnWidth = 10: wsExcel.Cells(1, 13) = "D. dernière clôture"
wsExcel.Columns(14).ColumnWidth = 6: wsExcel.Cells(1, 14) = "R.Com"
wsExcel.Columns(15).ColumnWidth = 30: wsExcel.Cells(1, 15) = "Intitulé"
wsExcel.Columns(16).ColumnWidth = 30: wsExcel.Cells(1, 16) = "Type de compte"
wsExcel.Columns(17).ColumnWidth = 6: wsExcel.Cells(1, 17) = "Dossier"
wsExcel.Columns(18).ColumnWidth = 8: wsExcel.Cells(1, 18) = "Nationalité": wsExcel.Cells(1, 18).Font.Size = 8
wsExcel.Columns(19).ColumnWidth = 8: wsExcel.Cells(1, 19) = "Résidence": wsExcel.Cells(1, 19).Font.Size = 8
wsExcel.Columns(20).ColumnWidth = 8: wsExcel.Cells(1, 20) = "Ouv/Clos": wsExcel.Cells(1, 20).Font.Size = 8

wsExcel.Rows(1).RowHeight = 34
wsExcel.Rows(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Rows(1).VerticalAlignment = Excel.xlVAlignCenter
wsExcel.Columns(14).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(18).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(19).HorizontalAlignment = Excel.xlHAlignCenter

mXls2_Col = 20: mXls2_Row = 1

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

Public Sub cmdSelect_SQL_Xgsop_Init_1(lCLIENARES As String)
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim X As String, K As Long
Dim xSQL As String, wFile As String, xRange_A As String, xRange_B As String, xRange_T As String

On Error GoTo Error_Handler
'__________________________________________________________________________________________________

'Set wsExcel = wbExcel.ActiveSheet
Set wsExcel = wbExcel.Sheets(1)

wsExcel.Name = "Reporting " '& lCLIENARES


With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .HorizontalAlignment = Excel.xlHAlignCenter
    .WrapText = True
    .Font.Size = 11
    .Font.Bold = True
    .Font.Name = "Calibri"
    .RowHeight = 15
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 100
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14GSOP : reporting clientèle  " & lCLIENARES _
                                & "&B&U&10   ( édité le " & dateImp10(DSys) & " " & Time & ")"
wsExcel.PageSetup.PrintTitleRows = "$A1:$F1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True
'wsExcel.PageSetup.PrintArea = "$A$1:$F$7"

wsExcel.Columns(1).ColumnWidth = 20: wsExcel.Cells(1, 1) = "  Clientèle": 'wsExcel.Cells(1, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(2).ColumnWidth = 15: wsExcel.Cells(1, 2) = "Particuliers": ' wsExcel.Cells(1, 2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(3).ColumnWidth = 15: wsExcel.Cells(1, 3) = "Pers. morales": ' wsExcel.Cells(1, 3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(4).ColumnWidth = 15: wsExcel.Cells(1, 4) = "Banques": ' wsExcel.Cells(1, 4).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(5).ColumnWidth = 15: wsExcel.Cells(1, 5) = "Autres": 'wsExcel.Cells(1, 5).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(6).ColumnWidth = 15: wsExcel.Cells(1, 6) = "Total": 'wsExcel.Cells(1, 6).HorizontalAlignment = Excel.xlHAlignCenter

xRange_A = "Detail!A1:Detail!A" & mXls2_Row
xRange_B = "Detail!B1:Detail!B" & mXls2_Row
xRange_T = "Detail!T1:Detail!T" & mXls2_Row


wsExcel.Cells(2, 1) = "  Clients": ' wsExcel.Cells(2, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Cells(2, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";1)" _
                               & " - NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";1;" & xRange_T & ";""Clos"")"
wsExcel.Cells(2, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";1)" _
                               & " - NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";1;" & xRange_T & ";""Clos"")"
wsExcel.Cells(2, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";1)" _
                               & " - NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";1;" & xRange_T & ";""Clos"")"
wsExcel.Cells(2, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";1)" _
                               & " - NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";1;" & xRange_T & ";""Clos"")"
wsExcel.Cells(2, 6).FormulaLocal = "=SOMME(B2:E2)": wsExcel.Cells(2, 6).Interior.Color = mColor_G0

wsExcel.Cells(3, 1) = "  Techniques": 'wsExcel.Cells(3, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Cells(3, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";2)" _
                               & " - NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";2;" & xRange_T & ";""Clos"")"
wsExcel.Cells(3, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";2)" _
                               & " - NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";2;" & xRange_T & ";""Clos"")"
wsExcel.Cells(3, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";2)" _
                               & " - NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";2;" & xRange_T & ";""Clos"")"
wsExcel.Cells(3, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";2)" _
                               & " - NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";2;" & xRange_T & ";""Clos"")"
wsExcel.Cells(3, 6).FormulaLocal = "=SOMME(B3:E3)": wsExcel.Cells(3, 6).Interior.Color = mColor_G0

wsExcel.Cells(4, 1) = "  Tiers": 'wsExcel.Cells(4, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Cells(4, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";3)" _
                               & " - NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";3;" & xRange_T & ";""Clos"")"
wsExcel.Cells(4, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";3)" _
                               & " - NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";3;" & xRange_T & ";""Clos"")"
wsExcel.Cells(4, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";3)" _
                               & " - NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";3;" & xRange_T & ";""Clos"")"
wsExcel.Cells(4, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";3)" _
                               & " - NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";3;" & xRange_T & ";""Clos"")"
wsExcel.Cells(4, 6).FormulaLocal = "=SOMME(B4:E4)": wsExcel.Cells(4, 6).Interior.Color = mColor_G0

wsExcel.Cells(5, 1) = "  BIA non réclamés": ' wsExcel.Cells(5, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Cells(5, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";4)" _
                               & " - NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";4;" & xRange_T & ";""Clos"")"
wsExcel.Cells(5, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";4)" _
                               & " - NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";4;" & xRange_T & ";""Clos"")"
wsExcel.Cells(5, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";4)" _
                               & " - NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";4;" & xRange_T & ";""Clos"")"
wsExcel.Cells(5, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";4)" _
                               & " - NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";4;" & xRange_T & ";""Clos"")"
wsExcel.Cells(5, 6).FormulaLocal = "=SOMME(B5:E5)": wsExcel.Cells(5, 6).Interior.Color = mColor_G0

wsExcel.Cells(6, 1) = "  hors GSOP": 'wsExcel.Cells(6, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Cells(6, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";5)" _
                               & " - NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";5;" & xRange_T & ";""Clos"")"
wsExcel.Cells(6, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";5)" _
                               & " - NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";5;" & xRange_T & ";""Clos"")"
wsExcel.Cells(6, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";5)" _
                               & " - NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";5;" & xRange_T & ";""Clos"")"
wsExcel.Cells(6, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";5)" _
                               & " - NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";5;" & xRange_T & ";""Clos"")"
wsExcel.Cells(6, 6).FormulaLocal = "=SOMME(B6:E6)": wsExcel.Cells(6, 6).Interior.Color = mColor_G0

wsExcel.Cells(7, 1) = "  Racines sans compte": 'wsExcel.Cells(6, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Cells(7, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";6)" _
                               & " - NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";6;" & xRange_T & ";""Clos"")"
wsExcel.Cells(7, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";6)" _
                               & " - NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";6;" & xRange_T & ";""Clos"")"
wsExcel.Cells(7, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";6)" _
                               & " - NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";6;" & xRange_T & ";""Clos"")"
wsExcel.Cells(7, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";6)" _
                               & " - NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";6;" & xRange_T & ";""Clos"")"
wsExcel.Cells(7, 6).FormulaLocal = "=SOMME(B7:E7)": wsExcel.Cells(7, 6).Interior.Color = mColor_G0

wsExcel.Cells(8, 1) = "  Total": ' wsExcel.Cells(1, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Cells(8, 2).FormulaLocal = "=SOMME(B2:B7)"
wsExcel.Cells(8, 3).FormulaLocal = "=SOMME(C2:C7)"
wsExcel.Cells(8, 4).FormulaLocal = "=SOMME(D2:D7)"
wsExcel.Cells(8, 5).FormulaLocal = "=SOMME(E2:E7)"
wsExcel.Cells(8, 6).FormulaLocal = "=SOMME(F2:F7)"

mXls1_Col = 6: mXls1_Row = 1

For K = 1 To mXls1_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next
For K = 1 To 8
    wsExcel.Cells(K, 1).Interior.Color = mColor_GB
    wsExcel.Cells(K, 1).Font.Color = mColor_Z0
Next
For K = 2 To mXls1_Col
    wsExcel.Cells(8, K).Interior.Color = mColor_G0
Next
wsExcel.Cells(8, 6).Interior.Color = mColor_G2


'__________________________________________________________________________________
'=======================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée "): DoEvents

End Sub

Public Sub cmdSelect_SQL_Xgsop_Init_3(lCLIENARES As String)
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim X As String, K As Long
Dim xSQL As String, wFile As String, xRange_A As String, xRange_B As String, xRange_Q As String

On Error GoTo Error_Handler
'__________________________________________________________________________________________________

'Set wsExcel = wbExcel.ActiveSheet
Set wsExcel = wbExcel.Sheets(3)

wsExcel.Name = "Dossier " & lCLIENARES


With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .VerticalAlignment = Excel.xlVAlignCenter
    .HorizontalAlignment = Excel.xlHAlignCenter
    .WrapText = True
    .Font.Size = 11
    .Font.Bold = True
    .Font.Name = "Calibri"
    .RowHeight = 30
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 100
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14GSOP : reporting Dossier  " & lCLIENARES _
                                & "&B&U&10   ( édité le " & dateImp10(DSys) & " " & Time & ")"
wsExcel.PageSetup.PrintTitleRows = "$A1:$F1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True
wsExcel.PageSetup.PrintArea = "$A$1:$F$7"

wsExcel.Columns(1).ColumnWidth = 20: wsExcel.Cells(1, 1) = "  Clientèle": 'wsExcel.Cells(1, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(2).ColumnWidth = 15: wsExcel.Cells(1, 2) = "Particuliers": ' wsExcel.Cells(1, 2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(3).ColumnWidth = 15: wsExcel.Cells(1, 3) = "Personnes morales": ' wsExcel.Cells(1, 3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(4).ColumnWidth = 15: wsExcel.Cells(1, 4) = "Banques": ' wsExcel.Cells(1, 4).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(5).ColumnWidth = 15: wsExcel.Cells(1, 5) = "Autres": 'wsExcel.Cells(1, 5).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(6).ColumnWidth = 15: wsExcel.Cells(1, 6) = "Total": 'wsExcel.Cells(1, 6).HorizontalAlignment = Excel.xlHAlignCenter

xRange_A = "Detail!A1:Detail!A" & mXls2_Row
xRange_B = "Detail!B1:Detail!B" & mXls2_Row
xRange_Q = "Detail!Q1:Detail!Q" & mXls2_Row

X = Asc34 & " " & Asc34
wsExcel.Cells(2, 1) = "  dossiers complets": ' wsExcel.Cells(2, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Cells(2, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";1;" & xRange_Q & ";" & X & ")"
wsExcel.Cells(2, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";1;" & xRange_Q & ";" & X & ")"
wsExcel.Cells(2, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";1;" & xRange_Q & ";" & X & ")"
wsExcel.Cells(2, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";1;" & xRange_Q & ";" & X & ")"
wsExcel.Cells(2, 6).FormulaLocal = "=SOMME(B2:E2)": wsExcel.Cells(2, 6).Interior.Color = mColor_G0

X = Asc34 & "X" & Asc34
wsExcel.Cells(3, 1) = "  dossiers incomplets": ' wsExcel.Cells(2, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Cells(3, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";1;" & xRange_Q & ";" & X & ")"
wsExcel.Cells(3, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";1;" & xRange_Q & ";" & X & ")"
wsExcel.Cells(3, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";1;" & xRange_Q & ";" & X & ")"
wsExcel.Cells(3, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";1;" & xRange_Q & ";" & X & ")"
wsExcel.Cells(3, 6).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";5;" & xRange_B & ";1;" & xRange_Q & ";" & X & ")"
wsExcel.Cells(3, 6).FormulaLocal = "=SOMME(B3:E3)": wsExcel.Cells(3, 6).Interior.Color = mColor_G0

X = Asc34 & "néant" & Asc34
wsExcel.Cells(4, 1) = "  dossiers à faire": ' wsExcel.Cells(2, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Cells(4, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";1;" & xRange_Q & ";" & X & ")"
wsExcel.Cells(4, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";1;" & xRange_Q & ";" & X & ")"
wsExcel.Cells(4, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";1;" & xRange_Q & ";" & X & ")"
wsExcel.Cells(4, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";1;" & xRange_Q & ";" & X & ")"
wsExcel.Cells(4, 6).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";5;" & xRange_B & ";1;" & xRange_Q & ";" & X & ")"
wsExcel.Cells(4, 6).FormulaLocal = "=SOMME(B4:E4)": wsExcel.Cells(4, 6).Interior.Color = mColor_G0

wsExcel.Cells(5, 1) = "  Total": ' wsExcel.Cells(1, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Cells(5, 2).FormulaLocal = "=SOMME(B2:B4)"
wsExcel.Cells(5, 3).FormulaLocal = "=SOMME(C2:C4)"
wsExcel.Cells(5, 4).FormulaLocal = "=SOMME(D2:D4)"
wsExcel.Cells(5, 5).FormulaLocal = "=SOMME(E2:E4)"
wsExcel.Cells(5, 6).FormulaLocal = "=SOMME(F2:F4)"

mXls1_Col = 6: mXls1_Row = 1

For K = 1 To mXls1_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next
For K = 1 To 5
    wsExcel.Cells(K, 1).Interior.Color = mColor_GB
    wsExcel.Cells(K, 1).Font.Color = mColor_Z0
Next
For K = 2 To mXls1_Col
    wsExcel.Cells(5, K).Interior.Color = mColor_G0
Next

'__________________________________________________________________________________
Exit_sub:
'=======================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée "): DoEvents

End Sub


Public Sub cmdSelect_SQL_Xf_Detail()
On Error GoTo Error_Handler
Dim X As String, XX As String
Dim wRow As Long, wCol As Long
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, kLen As Integer
Dim K5 As Integer, K6 As Integer, K7 As Integer, K8 As Integer
Dim rsSabX As ADODB.Recordset
'__________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents


X = "select *  from " & paramIBM_Library_SAB & ".zcliena0 , " & paramIBM_Library_SAB & ".zclienb0, " & paramIBM_Library_SAB & ".zclienc0" _
  & " where clienacli in (select titulacli from " & paramIBM_Library_SAB & ".ztitula0" _
  & " where titulacom in (select comptecom from  " & paramIBM_Library_SAB & ".zcompte0" _
  & " where comptefon <> '4'))" _
  & " and clienacli = clienbcli and clienacli = clienccli" _
  & " order by clienacli"

Set rsSab = cnsab.Execute(X)

wRow = 1
Do While Not rsSab.EOF
    V = rsZCLIENA0_GetBuffer(rsSab, xZCLIENA0)
    wRow = wRow + 1
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "> " & xZCLIENA0.CLIENACLI): DoEvents
    
    wsExcel.Cells(wRow, 1) = xZCLIENA0.CLIENAETB
    wsExcel.Cells(wRow, 2) = xZCLIENA0.CLIENAAGE
    wsExcel.Cells(wRow, 3) = xZCLIENA0.CLIENACLI
    wsExcel.Cells(wRow, 4) = xZCLIENA0.CLIENARES
    Select Case xZCLIENA0.CLIENACOL
        Case "0": wsExcel.Cells(wRow, 5) = ""
        Case "1": wsExcel.Cells(wRow, 5) = "C": wsExcel.Cells(wRow, 5).Interior.Color = mColor_Y1
        Case "2": wsExcel.Cells(wRow, 5) = "-"
        Case Else: wsExcel.Cells(wRow, 5) = xZCLIENA0.CLIENACOL: wsExcel.Cells(wRow, 5).Interior.Color = mColor_W1
    End Select
    wsExcel.Cells(wRow, 6) = xZCLIENA0.CLIENACOT
    
    Select Case xZCLIENA0.CLIENACHQ
        Case "N": wsExcel.Cells(wRow, 7) = ""
        Case Else: wsExcel.Cells(wRow, 7) = "Interdit": wsExcel.Cells(wRow, 7).Interior.Color = mColor_Y1
    End Select
    
    Select Case xZCLIENA0.CLIENADOU
        Case "N": wsExcel.Cells(wRow, 8) = ""
        Case Else: wsExcel.Cells(wRow, 8) = "Dtx": wsExcel.Cells(wRow, 8).Interior.Color = mColor_Y1
    End Select
    
    wsExcel.Cells(wRow, 9) = xZCLIENA0.CLIENASEL
    
    Select Case xZCLIENA0.CLIENAENT
        Case "000": wsExcel.Cells(wRow, 10) = ""
        Case Else: wsExcel.Cells(wRow, 10) = xZCLIENA0.CLIENAENT: wsExcel.Cells(wRow, 10).Interior.Color = mColor_Y1
    End Select
    
    wsExcel.Cells(wRow, 11) = xZCLIENA0.CLIENAETA
    wsExcel.Cells(wRow, 12) = xZCLIENA0.CLIENAECO
    wsExcel.Cells(wRow, 13) = xZCLIENA0.CLIENACAT
    
    wsExcel.Cells(wRow, 14) = xZCLIENA0.CLIENANAT
    wsExcel.Cells(wRow, 15) = xZCLIENA0.CLIENARSD
    If xZCLIENA0.CLIENANAT <> xZCLIENA0.CLIENARSD Then wsExcel.Cells(wRow, 15).Interior.Color = mColor_Y1
    wsExcel.Cells(wRow, 16) = xZCLIENA0.CLIENASIG
    wsExcel.Cells(wRow, 17) = xZCLIENA0.CLIENARA1
    wsExcel.Cells(wRow, 18) = xZCLIENA0.CLIENARA2
    wsExcel.Cells(wRow, 20) = rsSab("CLIENASRN"): wsExcel.Cells(wRow, 19).Interior.Color = mColor_Y1
    If Trim(xZCLIENA0.CLIENAREG) <> "" Then
        wsExcel.Cells(wRow, 19) = xZCLIENA0.CLIENAREG
        X = "select *  from " & paramIBM_Library_SAB & ".ZBASTAB0" _
          & " where BASTABETA = 1 and BASTABNUM = 2 and BASTABARG = 'CLI" & Trim(xZCLIENA0.CLIENAREG) & "'"
        Set rsSabX = cnsab.Execute(X)
        If Not rsSabX.EOF Then wsExcel.Cells(wRow, 21) = Mid$(rsSabX("BASTABDON"), 19, 80)
    Else
        wsExcel.Cells(wRow, 19) = rsSab("CLIENCPRF"): wsExcel.Cells(wRow, 19).Interior.Color = mColor_Y1
        X = "select *  from " & paramIBM_Library_SAB & ".ZBASTAB0" _
          & " where BASTABETA = 1 and BASTABNUM = 121 and BASTABARG = '" & Trim(rsSab("CLIENCPRF")) & "'"
        Set rsSabX = cnsab.Execute(X)
        If Not rsSabX.EOF Then wsExcel.Cells(wRow, 21) = Trim(Mid$(rsSabX("BASTABDON"), 1, 24)): wsExcel.Cells(wRow, 21).Interior.Color = mColor_Y1
        
    End If
    
    

    rsSab.MoveNext
Loop

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub
Public Sub cmdSelect_SQL_Xf_Mandataires()
On Error GoTo Error_Handler
Dim X As String, XX As String
Dim wRow As Long, wCol As Long
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, kLen As Integer
Dim K5 As Integer, K6 As Integer, K7 As Integer, K8 As Integer
Dim rsSabX As ADODB.Recordset
Dim Nb_Total As Long
'__________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents


X = "select count(*)  from " & paramIBM_Library_SAB & ".zcliena0 where clienacli like '99%'"
Set rsSab = cnsab.Execute(X)
Nb_Total = rsSab(0)

X = "select *  from " & paramIBM_Library_SAB & ".zcliena0 , " & paramIBM_Library_SAB & ".zclienb0 , " _
  & paramIBM_Library_SAB & ".ZADRESS0" _
  & " where clienacli like '99%'" _
  & " and clienacli = clienbcli" _
  & " and substring(ADRESSNUM , 2 , 19) = clienacli " _
  & " order by clienacli"

Set rsSab = cnsab.Execute(X)

wRow = 1
Do While Not rsSab.EOF
    V = rsZCLIENA0_GetBuffer(rsSab, xZCLIENA0)
    wRow = wRow + 1
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "> " & xZCLIENA0.CLIENACLI): DoEvents
    
    wsExcel.Cells(wRow, 1) = xZCLIENA0.CLIENAETB
    wsExcel.Cells(wRow, 2) = xZCLIENA0.CLIENAAGE
    wsExcel.Cells(wRow, 3) = xZCLIENA0.CLIENACLI
    wsExcel.Cells(wRow, 4) = xZCLIENA0.CLIENARES
    Select Case xZCLIENA0.CLIENACOL
        Case "0": wsExcel.Cells(wRow, 5) = ""
        Case "1": wsExcel.Cells(wRow, 5) = "C": wsExcel.Cells(wRow, 5).Interior.Color = mColor_Y1
        Case "2": wsExcel.Cells(wRow, 5) = "-"
        Case Else: wsExcel.Cells(wRow, 5) = xZCLIENA0.CLIENACOL: wsExcel.Cells(wRow, 5).Interior.Color = mColor_W1
    End Select
    wsExcel.Cells(wRow, 6) = xZCLIENA0.CLIENACOT
    
    'Select Case xZCLIENA0.CLIENACHQ
    '    Case "N": wsExcel.Cells(wRow, 7) = ""
    '    Case Else: wsExcel.Cells(wRow, 7) = "Interdit": wsExcel.Cells(wRow, 7).Interior.Color = mColor_Y1
    'End Select
    If Trim(xZCLIENA0.CLIENASRN) = "" Then
        wsExcel.Cells(wRow, 8) = dateImp10_S(xZCLIENA0.CLIENADNA + 19000000)
    Else
        wsExcel.Cells(wRow, 9) = xZCLIENA0.CLIENASRN
    End If
    
    
    Select Case xZCLIENA0.CLIENAENT
        Case "000": wsExcel.Cells(wRow, 10) = ""
        Case Else: wsExcel.Cells(wRow, 10) = xZCLIENA0.CLIENAENT: wsExcel.Cells(wRow, 10).Interior.Color = mColor_Y1
    End Select
    
    wsExcel.Cells(wRow, 11) = xZCLIENA0.CLIENAETA
    wsExcel.Cells(wRow, 12) = xZCLIENA0.CLIENAECO
    wsExcel.Cells(wRow, 13) = xZCLIENA0.CLIENACAT
    
    wsExcel.Cells(wRow, 14) = xZCLIENA0.CLIENANAT
    wsExcel.Cells(wRow, 15) = xZCLIENA0.CLIENARSD
    If xZCLIENA0.CLIENANAT <> xZCLIENA0.CLIENARSD Then wsExcel.Cells(wRow, 15).Interior.Color = mColor_Y1
    wsExcel.Cells(wRow, 16) = xZCLIENA0.CLIENASIG
    wsExcel.Cells(wRow, 17) = xZCLIENA0.CLIENARA1
    wsExcel.Cells(wRow, 18) = xZCLIENA0.CLIENARA2
    
                    wsExcel.Cells(wRow, 19) = Trim(rsSab("ADRESSAD1"))
                    wsExcel.Cells(wRow, 20) = Trim(rsSab("ADRESSAD2"))
                    wsExcel.Cells(wRow, 21) = Trim(rsSab("ADRESSAD3"))
                    wsExcel.Cells(wRow, 22) = Trim(rsSab("ADRESSCOP"))
                    wsExcel.Cells(wRow, 23) = Trim(rsSab("ADRESSVIL"))
                    wsExcel.Cells(wRow, 24) = Trim(rsSab("ADRESSPAY"))
                    
    If Mid$(xZCLIENA0.CLIENARES, 1, 1) = "X" Then
        wsExcel.Cells(wRow, 7) = "annulé"
        For K = 1 To 24: wsExcel.Cells(wRow, K).Font.Color = RGB(128, 128, 128): Next K
    Else
        For K = 1 To 24: wsExcel.Cells(wRow, K).Font.Color = RGB(0, 0, 224): Next K
    
    End If

    rsSab.MoveNext
Loop

If Nb_Total <> wRow - 1 Then Call MsgBox("Nb_Total : " & Nb_Total & " # Nb_export : " & wRow - 1, vbCritical, "Exportation des Mandataires")
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub cmdSelect_SQL_Xf_Table(lTxt As String, lBASTABNUM As Integer)
On Error GoTo Error_Handler
Dim X As String, XX As String
Dim wRow As Long, wCol As Long
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, kLen As Integer
Dim K5 As Integer, K6 As Integer, K7 As Integer, K8 As Integer
Dim rsSabX As ADODB.Recordset
'__________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours :Table " & lTxt): DoEvents
wsExcel.Name = lTxt


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
    .Font.Name = "Courier New"
    .RowHeight = 17
End With
wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 100
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14CLI : Tables " & lTxt & " en date du " & dateImp10_S(YBIATAB0_DATE_CPT_J) _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$L1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


wsExcel.Columns(1).ColumnWidth = 5: wsExcel.Cells(1, 1) = "Table"
wsExcel.Columns(2).ColumnWidth = 8: wsExcel.Cells(1, 2) = "Code"
wsExcel.Columns(3).ColumnWidth = 80: wsExcel.Cells(1, 3) = "Libellé"
For K = 1 To 3
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next


wsExcel_Row = 1


X = "select *  from " & paramIBM_Library_SAB & ".ZBASTAB0 " _
  & " where BASTABETA = 1 and BASTABNUM =" & lBASTABNUM _
  & " order by BASTABARG"

Set rsSab = cnsab.Execute(X)


Do While Not rsSab.EOF
    wsExcel_Row = wsExcel_Row + 1
    
    wsExcel.Cells(wsExcel_Row, 1) = lTxt
    Select Case lBASTABNUM
        Case 2
            wsExcel.Cells(wsExcel_Row, 2) = Trim(Mid$(rsSab("BASTABARG"), 4, 13))
            wsExcel.Cells(wsExcel_Row, 3) = RTrim(Mid$(rsSab("BASTABDON"), 19, 80))
        Case 121
            wsExcel.Cells(wsExcel_Row, 2) = Trim(rsSab("BASTABARG"))
            wsExcel.Cells(wsExcel_Row, 3) = RTrim(rsSab("BASTABDON"))
        Case Else
            wsExcel.Cells(wsExcel_Row, 2) = Trim(Mid$(rsSab("BASTABARG"), 4, 13))
            wsExcel.Cells(wsExcel_Row, 3) = rsSab("BASTABLO2") & RTrim(rsSab("BASTABDON"))
        End Select

    rsSab.MoveNext
Loop

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub cmdSelect_SQL_Xz_Detail(lCLIENARES As String)
On Error GoTo Error_Handler
Dim X As String, XX As String
Dim wRow As Long, wCol As Long
Dim K As Long, K1 As Long, K2 As Long
Dim blnActif As Boolean
'__________________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

If lCLIENARES = "CLI" Then
    XX = ""
Else
    XX = " and CLIENARES = '" & lCLIENARES & "'"
End If

X = "select count(*)  from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
  & " where clienacli <> ''" _
  & " and COMPTEFON <> '4' " & XX

Set rsSab = cnsab.Execute(X)

ReDim arrYBIACPT0(rsSab(0) + 1)
arrYBIACPT0_Nb = 0

X = "select *  from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
  & " where clienacli <> ''" _
  & " and COMPTEFON <> '4' " & XX _
  & " order by clienacli,COMPTECOM"

Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    arrYBIACPT0_Nb = arrYBIACPT0_Nb + 1
    V = rsYBIACPT0_GetBuffer(rsSab, arrYBIACPT0(arrYBIACPT0_Nb))
    rsSab.MoveNext
Loop

rsYBIACPT0_Init oldYBIACPT0
K1 = 0: K2 = 0
blnActif = True
mXls1_Row = 1

For K = 1 To arrYBIACPT0_Nb
    If oldYBIACPT0.CLIENACLI <> arrYBIACPT0(K).CLIENACLI Then
        If Not blnActif Then Call cmdSelect_SQL_Xz_Detail_Compte(K1, K - 1)
        K1 = K
        blnActif = False
        oldYBIACPT0 = arrYBIACPT0(K)
    End If
    If arrYBIACPT0(K).SOLDECEN <> 0 Then
        blnActif = True
    Else
        If arrYBIACPT0(K).SOLDEDMO > wIBM_AmjMin Then blnActif = True
        If arrYBIACPT0(K).COMPTEOUV > wIBM_AmjMin Then blnActif = True
    End If
Next K
If Not blnActif Then Call cmdSelect_SQL_Xz_Detail_Compte(K1, K - 1)


Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub cmdSelect_SQL_Xz_Detail_Compte(lK1 As Long, lK2 As Long)
On Error GoTo Error_Handler
Dim K As Integer
'___________________________

Call lstErr_ChangeLastItem(lstErr, cmdContext, "> " & oldYBIACPT0.CLIENACLI): DoEvents

mXls1_Row = mXls1_Row + 1

wsExcel.Cells(mXls1_Row, 1) = oldYBIACPT0.CLIENACLI: wsExcel.Cells(mXls1_Row, 1).Font.Bold = True
wsExcel.Cells(mXls1_Row, 5) = "(" & oldYBIACPT0.CLIENARES & ")"
wsExcel.Cells(mXls1_Row, 6) = "(" & oldYBIACPT0.CLIENARSD & ")"
wsExcel.Cells(mXls1_Row, 7) = Trim(oldYBIACPT0.CLIENARA1) & " " & Trim(oldYBIACPT0.CLIENARA2)
For K = 1 To mXls1_Col: wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_Y0: Next K

For K = lK1 To lK2

    mXls1_Row = mXls1_Row + 1
    xYBIACPT0 = arrYBIACPT0(K)
    'wsExcel.Cells(mXls1_Row, 1) = xYBIACPT0.CLIENACLI
    wsExcel.Cells(mXls1_Row, 2) = xYBIACPT0.PLANCOPRO
    wsExcel.Cells(mXls1_Row, 3) = xYBIACPT0.COMPTECOM
    wsExcel.Cells(mXls1_Row, 4) = xYBIACPT0.COMPTEDEV
    wsExcel.Cells(mXls1_Row, 5) = dateImp10(xYBIACPT0.SOLDEDMO + 19000000)
    wsExcel.Cells(mXls1_Row, 6) = dateImp10(xYBIACPT0.COMPTEOUV + 19000000)
    wsExcel.Cells(mXls1_Row, 7) = xYBIACPT0.COMPTEINT
Next K

Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub cmdSelect_Reset()
If blnControl Then
    cmdSelect_Clear
    fraSelect_Options_3.Visible = False
    fraSelect_Options_4.Visible = False
    fraSelect_Options_No.Visible = False

    fraSelect_Options_3.Visible = False
    fraSelect_Options_Xgsop.Visible = False
    fraSelect_Options_KYCgsop.Visible = False
    fraSelect_Options_KYCech.Visible = False
    
    cmdSelect_Ok.Visible = True
    Dim K As Integer
    K = InStr(cboSelect_SQL, "-")
    mSelect_SQL = Trim(Mid$(cboSelect_SQL, 1, K - 1))
    
    Select Case mSelect_SQL
        Case Is = "1": lblSelect_CLIENACAT = "Catégorie client": cboSelect_CLIENACAT = ""
        Case Is = "1r": lblSelect_CLIENACAT = "Catégorie client": cboSelect_CLIENACAT = "": chkSelect_Racine.Visible = True: chkSelect_Groupes.Visible = True
        Case Is = "2": lblSelect_CLIENACAT = "Type de lien": cboSelect_CLIENACAT = ""
        Case Is = "3": fraSelect_Options_3.Visible = True
        Case Is = "4": fraSelect_Options_4.Visible = True
        Case Is = "5": fraSelect_Options_4.Visible = True
        'Case Is = "Xgs", "KYC gsop", "Xgsop@": fraSelect_Options_Xgsop.Visible = True
        Case "4!e": fraSelect_Options_No.Visible = True

       Case "3":       fraSelect_Options_3.Visible = True
       Case "4":
                If cboSelect_Options_4_ECHISBFIN.ListCount = 0 Then
                    Dim K2 As Integer, wAAAA As String, X As String
                    cboSelect_Options_4_ECHISBFIN.Clear
                    wAAAA = Mid$(DSys, 1, 4)
                    For K2 = wAAAA To wAAAA - 2 Step -1
                        For K = 12 To 1 Step -1
                            X = K2 & Format(K, "00") & "01"
                            If X < DSys Then cboSelect_Options_4_ECHISBFIN.AddItem X
                        Next K
                    Next K2
                    
                    cboSelect_Options_4_ECHISBFIN.ListIndex = 0
                End If
                 cboSelect_Options_4_Code.Visible = True
                 chkSelect_Options_4_AUTSICMON.Visible = True
                 chkSelect_Options_4_ECHTABDON_S.Visible = True
       Case "5": cboSelect_Options_4_Code.Visible = False
                 chkSelect_Options_4_AUTSICMON.Visible = False
                 chkSelect_Options_4_ECHTABDON_S.Visible = False
       Case "Xgsop", "Xgsop@":
                    cmdSelect_SQL_Xgsop_Init
                    fraSelect_Options_Xgsop.Visible = True
       Case "KYC gsop":
                    cmdSelect_SQL_Xgsop_Init
                    fraSelect_Options_KYCgsop.Visible = True
       Case "KYC ech":
                    If cboSelect_Options_KYCech_Doc.ListCount = 0 Then cboSelect_Options_KYCech_Doc_Load
                    cmdSelect_SQL_Xgsop_Init
                    fraSelect_Options_KYCech.Visible = True
    End Select

End If
End Sub

Private Sub cmdSelect_Update_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

fraUpdate.Visible = False
fraYKYCDOS0.Visible = False
If optSelect_CLIENARES Then
    newZCLIENA0 = oldZCLIENA0
    If Mid$(oldZCLIENA0.CLIENARES, 1, 1) = "X" Then
        newZCLIENA0.CLIENARES = "T99"
    Else
        newZCLIENA0.CLIENARES = "X99"
    End If
    newZCLIENB0 = oldZCLIENB0
    newZADRESS0 = oldZADRESS0
    cmdUpdate_Modification
    tvwSelect_Display_10 Mid$(mSelect_Node_Key, 4, 7)
Else
    If optSelect_Suppress Then
        cmdUpdate_Delete_Old
        tvwSelect_Display_10 Mid$(mSelect_Node_Key, 4, 7)
    Else
        SSTab1.Tab = 1
        mUpdate_Node_Key = ""
        fraUpdate.Enabled = arrHab(2)
        SSTab1.Caption = "Mise à jour des mandataires"
        If optSelect_Modification Then
             cboUpdate_Add.Enabled = False
             optUpdate_Add_Old.Enabled = False
             optUpdate_Add_PP.Enabled = False
             optUpdate_Add_Sté.Enabled = False
             tvwUpdate.Enabled = False
             fraUpdate_Détail.Enabled = True
             cmdUpdate_Ok.Caption = "Modification"
             fraUpdate_Add.Caption = "Modification d'un tiers"
    
            fraUpdate.Visible = True
        Else
            If optSelect_Add Then
                cboUpdate_Add.Clear
                cboUpdate_Add.AddItem "MAN Mandataire"
                If Trim(selZCLIENA0.CLIENASRN) = "" Then
                    optUpdate_Add_Sté.Enabled = False
                Else
                    cboUpdate_Add.AddItem "ADM Administrateur"
                    cboUpdate_Add.AddItem "DIR Direction"
                    optUpdate_Add_Sté.Enabled = True
                End If
                fraUpdate_Init
                 cboUpdate_Add.Enabled = True
                 cboUpdate_Add.ListIndex = 0
                 optUpdate_Add_Old.Enabled = True
                 optUpdate_Add_PP.Enabled = True
                 optUpdate_Add_PP.Value = True
                 tvwUpdate.Enabled = True
                cmdSelect_SQL_99
                fraUpdate.Visible = True
                fraUpdate.Enabled = True
            End If
        End If
    End If
End If
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdUpdate_Ok_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CLIENT mise à jour ........"): DoEvents
V = "cmdUpdate_Ok_Click"
If optSelect_Modification Then
        newZCLIENA0 = oldZCLIENA0
        newZCLIENB0 = oldZCLIENB0
        newZADRESS0 = oldZADRESS0
        cmdUpdate_Add_Control
        If blnUpdate_Ok Then V = cmdUpdate_Modification

Else
    If optUpdate_Add_Old Then
        If blnUpdate_Ok Then
            V = cmdUpdate_Add_Old
        Else
            Call lstErr_AddItem(lstErr, cmdContext, "? choisir un TIERS"): DoEvents
        End If
    Else
    'If optUpdate_Add_PP Then
        rsZCLIENA0_Init newZCLIENA0
        rsZCLIENB0_Init newZCLIENB0
        rsZADRESS0_Init newZADRESS0
        cmdUpdate_Add_Control
        If blnUpdate_Ok Then V = cmdUpdate_Add_New
    End If
End If
If IsNull(V) Then
fraUpdate_Init:         SSTab1.Tab = 0
    tvwSelect_Display_10 Mid$(mSelect_Node_Key, 4, 7)
    Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CLIENT terminé"): DoEvents
    fraUpdate.Enabled = False
    fraSelect_Update.Visible = False
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Function cmdUpdate_Transaction()
Dim V
V = cnSAB_Transaction("BeginTrans")
cmdUpdate_Transaction = V
End Function


Private Sub cmdUpdLog_Ok_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CLIENT_Piste d'audit ........"): DoEvents

If optUpdLog_YUPDLOG0 Then
    cmdSelect_SQL_YUPDLOG0
Else
    cmdSelect_SQL_YKYCDOSH
End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CLIENT_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdYKYCDOS0_Add_Click()
Dim V, X As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

If fraYKYCDOS0_Update_Control Then
    newYKYCDOS0.KYCDOSUFCT = "A"
    V = cmdParam_YKYCDOS0_Transaction("New")
    If IsNull(V) Then
        fraYKYCDOS0_Update.Visible = False
        cmdSelect_SQL_YKYCDOS0_Init
    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdYKYCDOS0_Delete_All_Click()
Dim V, X As String, blnOk As Boolean, objFolder, objFiles, fsoFile As File
Dim wPJ_Nb As Integer, wDocument_Nb As Integer

On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass

If vbNo = MsgBox("Confirmez-vous l'effacement de ce dossier (incluant les pièces jointes) ?", vbYesNo, "cmdYKYCDOS0_Delete_All_Click") Then GoTo Exit_sub

X = paramGSOP_Dossier_Path & oldYKYCDOS0.KYCDOSID & "\"
If Dir(X) <> "" Then
    Set objFolder = msFileSystem.GetFolder(X)
    Set objFiles = objFolder.Files
    For Each fsoFile In objFiles
        wPJ_Nb = wPJ_Nb + 1
        msFileSystem.DeleteFile X & fsoFile.Name
    Next
End If


X = "select count(*) from  " & paramIBM_Library_SABSPE & ".YKYCDOS0 " _
    & " where  KYCDOSNAT = ' ' and KYCDOSID = '" & selZCLIENA0.CLIENACLI & "'"
Set rsSab = cnsab.Execute(X)
If Not rsSab.EOF Then wDocument_Nb = rsSab(0)

oldYKYCDOS0.KYCDOSDLIB = "Suppression de " & wDocument_Nb & " documents et de " & wPJ_Nb & " pièces jointes"

oldYKYCDOS0.KYCDOSUFCT = "Z"

V = cmdParam_YKYCDOS0_Transaction("Delete_All")
If IsNull(V) Then
    fraYKYCDOS0_Update.Visible = False
    cmdSelect_SQL_YKYCDOS0_Init
End If

GoTo Exit_sub

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, "cmdPJ_Delete_Click :" & currentPJ_Path_FileName
Exit_sub:

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdYKYCDOS0_Delete_Click()
Dim V, X As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

oldYKYCDOS0.KYCDOSUFCT = "D"

    V = cmdParam_YKYCDOS0_Transaction("Delete")
    If IsNull(V) Then
        fraYKYCDOS0_Update.Visible = False
        cmdSelect_SQL_YKYCDOS0_Init
    End If

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdYKYCDOS0_Ignore_Click()
Dim V, X As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

If fraYKYCDOS0_Update_Control Then
    newYKYCDOS0.KYCDOSUFCT = "I"
    newYKYCDOS0.KYCDOSSTAK = "I"
    newYKYCDOS0.KYCDOSDECH = 0: newYKYCDOS0.KYCDOSDAMJ = DSys
    newYKYCDOS0.KYCDOSDLIB = Trim(txtYKYCDOS0_KYCDOSDLIB)
    If newYKYCDOS0.KYCDOSDLIB = "" Then
        newYKYCDOS0.KYCDOSDLIB = "Document sans objet (" & usrName_UCase & ")"
    End If
    
    V = cmdParam_YKYCDOS0_Transaction("New")
    If IsNull(V) Then
        fraYKYCDOS0_Update.Visible = False
        cmdSelect_SQL_YKYCDOS0_Init
    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdYKYCDOS0_Missing_Click()
Dim V, X As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

If fraYKYCDOS0_Update_Control Then
    newYKYCDOS0.KYCDOSUFCT = "?"
    newYKYCDOS0.KYCDOSSTAK = "?"
    newYKYCDOS0.KYCDOSDECH = 0: newYKYCDOS0.KYCDOSDAMJ = 0
    V = cmdParam_YKYCDOS0_Transaction("New")
    If IsNull(V) Then
        fraYKYCDOS0_Update.Visible = False
        cmdSelect_SQL_YKYCDOS0_Init
    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdYKYCDOS0_PJ_Click()
If fraYKYCDOS0_Update_Control Then

    cmdPJ_OK.Visible = False
    rtfPJ.Top = filDoc.Top + filDoc.Height + 200 ' 3400
    rtfPJ.Left = filDoc.Left
    rtfPJ.Width = filDoc.Width
    rtfPJ.Height = filDoc.Height '2025
    fraPJ.Visible = True
    filDoc.Pattern = "_.*"
    filDoc.Pattern = "*.*"
    
    oldFileName = "": newFileName = ""
    rtfPJ.Text = ""
    
    fraPJ.ZOrder 0
    fraPJ.Visible = True
End If
End Sub

Private Sub cmdYKYCDOS0_Quit_Click()
fraYKYCDOS0_Update.Visible = False
End Sub

Private Sub cmdYKYCDOS0_Update_Click()
Dim V, X As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

If fraYKYCDOS0_Update_Control Then
    newYKYCDOS0.KYCDOSUFCT = "U"
    V = cmdParam_YKYCDOS0_Transaction("Update")
    If IsNull(V) Then
        fraYKYCDOS0_Update.Visible = False
        cmdSelect_SQL_YKYCDOS0_Init
    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub dirListBox_Change()
filDoc.PATH = dirListBox.PATH
filDoc.Pattern = "*.*"
If mfilDoc_Path <> filDoc.PATH Then cmdPJ_Path.Visible = True

End Sub

Private Sub DriveListBox_Change()
On Error Resume Next
dirListBox.PATH = DriveListBox.Drive ' .PATH

End Sub


Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim XX As String
On Error Resume Next
fraYKYCDOS0_Update.Visible = False
If y <= fgDetail.RowHeightMin Then

    Select Case fgDetail.Col
        Case 0: fgDetail_Sort1 = 0: fgDetail_Sort2 = 0: fgDetail_Sort
        Case 1: fgDetail_Sort1 = 1: fgDetail_Sort2 = 1: fgDetail_Sort
        Case 2: fgDetail_Sort1 = 2: fgDetail_Sort2 = 2: fgDetail_Sort
        Case 3: fgDetail_Sort1 = 3: fgDetail_Sort2 = 3: fgDetail_Sort
        Case 4: fgDetail_Sort1 = 4: fgDetail_Sort2 = 4: fgDetail_Sort
        Case 5: fgDetail_Sort1 = 5: fgDetail_Sort2 = 5: fgDetail_Sort
        Case 6: fgDetail_Sort1 = 6: fgDetail_Sort2 = 6: fgDetail_Sort
        Case 7: fgDetail_Sort1 = 7: fgDetail_Sort2 = 7: fgDetail_Sort
        Case 8: fgDetail_Sort1 = 8: fgDetail_Sort2 = 8: fgDetail_Sort
       'Case fgDetail_arrIndex:  fgDetail_SortX fgDetail_arrIndex
    End Select


Else
    If fgDetail.Rows > 1 Then
        If mSelect_SQL = "KYC gsop" Or mSelect_SQL = "KYC ech" Then
        
            fgDetail.Col = 0: XX = Trim(fgDetail.Text)
            

            XX = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
                    & " where clienacli = '" & XX & "'"

            Set rsSab = cnsab.Execute(XX)
            
            If Not rsSab.EOF Then
                Call rsZCLIENA0_GetBuffer(rsSab, selZCLIENA0)
                cmdSelect_SQL_YKYCDOS0_Init
                
            End If

        End If
   End If
End If
fgDetail.LeftCol = 0
fgDetail.Col = 0

End Sub

Private Sub fgParam_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If fgParam.Rows > 1 Then
    fgParam.Col = 0: txtParam_Id = Trim(fgParam.Text)
    cmdParam_Delete.Visible = True
End If

End Sub


Private Sub fgParam_Display()
Dim xSQL As String, V
Dim X As String

On Error GoTo Error_Handler
fgParam.Visible = False
cmdParam_Delete.Visible = False

fgParam.Rows = 1
fgParam.FormatString = fgParam_FormatString
fgParam.Row = 0

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = '" & Old_YBIATAB0.BIATABID & "' and BIATABK1 = '" & Old_YBIATAB0.BIATABK1 & "' order by BIATABK2 "
Set rsParam = cnsab.Execute(xSQL)

Do While Not rsParam.EOF
    fgParam.Rows = fgParam.Rows + 1
    fgParam.Row = fgParam.Rows - 1
    fgParam.Col = 2: fgParam.Text = rsParam("BIATABTXT")
    X = Trim(rsParam("BIATABK2"))
    fgParam.Col = 0: fgParam.Text = X
    fgParam.Col = 1: fgParam.Text = fgParam_Display_Lib(X)
    rsParam.MoveNext

Loop



fgParam.Visible = True
'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : fgParam_Display"


End Sub

Public Function fgParam_Display_Lib(lBIATABK2 As String) As String
Dim xSQL As String

fgParam_Display_Lib = "?"

Select Case lstParam_K
    Case "3"
    
        xSQL = "select * from  " & paramIBM_Library_SAB & ".ZBASTAB0 " _
            & " where BASTABETA = 1 and BASTABNUM = 14 and BASTABARG = '" & lBIATABK2 & "'"
        Set rsSabX = cnsab.Execute(xSQL)
        If Not rsSabX.EOF Then
            fgParam_Display_Lib = rsSabX("BASTABLO2") & rsSabX("BASTABDON")
        End If
    Case Else
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
             & " where CLIENACLI = '" & lBIATABK2 & "'"
            Set rsSabX = cnsab.Execute(xSQL)
    
            If Not rsSabX.EOF Then
                fgParam_Display_Lib = Trim(rsSabX("CLIENARA1")) & " " & Trim(rsSabX("CLIENARA2"))
            End If
End Select
End Function


Private Sub fgParam_YKYCDOS0_4c_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K1 As Long, K2 As Long, xSQL As String
On Error Resume Next
fraParam_YKYCDOS0_4c.Visible = False
If y <= fgParam_YKYCDOS0_4c.RowHeightMin Then
Else
    If fgParam_YKYCDOS0_4c.Rows > 1 Then
        fgParam_YKYCDOS0_4c.Col = 0:  K1 = CLng(fgParam_YKYCDOS0_4c.Text)
        fgParam_YKYCDOS0_4c.Col = 1: K2 = CLng(fgParam_YKYCDOS0_4c.Text)
        fgParam_YKYCDOS0_4c.Col = 2: fraParam_YKYCDOS0_4c = K1 & "-" & K2 & " :  " & fgParam_YKYCDOS0_4c.Text
        xSQL = "select *from " & paramIBM_Library_SABSPE & ".YKYCDOS0 where KYCDOSNAT = '='" _
             & " and KYCDOSID = '" & oldYKYCDOS0.KYCDOSID & "' and KYCDOSSEQ  = " & K1 & " and KYCDOSSEQ2  = " & K2
        
        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            Call rsYKYCDOS0_GetBuffer(rsSab, oldYKYCDOS0)
            If oldYKYCDOS0.KYCDOSSTAK = "O" Then
                chkParam_YKYCDOS0_4c.Value = "1"
            Else
                chkParam_YKYCDOS0_4c.Value = "0"
            End If
            txtParam_YKYCDOS0_4c = oldYKYCDOS0.KYCDOSDLIB
            fraParam_YKYCDOS0_4c.Visible = True
        End If
        
   End If
End If
fgParam_YKYCDOS0_4c.Col = 0
fgParam_YKYCDOS0_4c.LeftCol = 0

End Sub


Private Sub fgPJ_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
If y <= fgPJ.RowHeightMin Then
Else
    If fgPJ.Rows > 1 Then
   End If
End If
fgPJ.Col = 0

currentPJ_FileName = fgPJ.Text
currentPJ_Path_FileName = paramGSOP_Dossier_Path & oldYKYCDOS0.KYCDOSID & "\" & fgPJ.Text
Call frmElpPrt.Windows_Display_File(currentPJ_Path_FileName)
If arrHab(16) Then
    cmdPJ_Delete.Caption = "Supprimer le document : " & fgPJ.Text
    cmdPJ_Delete.Visible = True
End If


End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long, XX As String
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_Sort
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  arrZCLIENA0_Index = CLng(fgSelect.Text)
        fgSelect.LeftCol = 0
            oldZCLIENA0 = arrZCLIENA0(arrZCLIENA0_Index)
            ''fraSelect_Update_Display
   End If
End If
fgSelect.LeftCol = 0
End Sub

Private Sub fgYKYCDOS0_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next
fraYKYCDOS0_Update.Visible = False
If y <= fgYKYCDOS0.RowHeightMin Then
Else
    If fgYKYCDOS0.Rows > 1 Then
        'Call fgYKYCDOS0_Color(fgYKYCDOS0_RowClick, MouseMoveUsr.BackColor, fgYKYCDOS0_ColorClick)
        fgYKYCDOS0.Col = 6:  K = CLng(fgYKYCDOS0.Text)
        'If K > 0 Then
           fgYKYCDOS0.Col = 0
            libYKYCDOS0_Document = fgYKYCDOS0.Text
            oldYKYCDOS0 = arrYKYCDOS0(fgYKYCDOS0.Row)
            fraYKYCDOS0_Update.Caption = K
            
            'If arrHab(15) Then
                fraYKYCDOS0_Update_Display
            'Else
            '    If oldYKYCDOS0.KYCDOSSTAK <> "X" Then fraYKYCDOS0_Update_Display
            'End If
        'End If
   End If
End If
fgYKYCDOS0.LeftCol = 0
fgYKYCDOS0.Col = 0
End Sub

Private Sub fgYKYCDOS0_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If fgYKYCDOS0.Width = fraYKYCDOS_ZCLIENA0.Width Then fgYKYCDOS0.Width = fraYKYCDOS0.Width - 120
End Sub

Private Sub fgYKYCDOS0_ZADRESS0_Click()
If fgYKYCDOS0.Width <> fraYKYCDOS_ZCLIENA0.Width Then
    fgYKYCDOS0.Width = fraYKYCDOS_ZCLIENA0.Width
    fgYKYCDOS0_ZADRESS0.LeftCol = 0
    fgYKYCDOS0_ZADRESS0.Col = 0
Else
    fgYKYCDOS0.Width = fraYKYCDOS0.Width - 120
End If
End Sub

Private Sub fgYKYCDOS0_ZADRESS0_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'If fgYKYCDOS0.Width <> fraYKYCDOS_ZCLIENA0.Width Then fgYKYCDOS0.Width = fraYKYCDOS_ZCLIENA0.Width
End Sub

Private Sub filDoc_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

If mfilDoc_Path <> filDoc.PATH Then cmdPJ_Path.Visible = True

oldFileName = filDoc.PATH & "\" & filDoc.FileName
newDirPath = paramGSOP_Dossier_Path & oldYKYCDOS0.KYCDOSID
newFileName = filDoc.FileName
newFileExtension = fileName_Extension(filDoc.FileName)
Call frmElpPrt.Windows_Display_File(oldFileName)

cmdPJ_OK.Visible = True
Me.Enabled = True: Me.MousePointer = 0
On Error Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)

cnAdo_Close
End Sub

Private Sub lstParam_Click()
Dim xSQL As String, X As String

fraParam_Update.Visible = False
cmdParam_Delete.Visible = False
cmdParam_Add.Visible = False
txtParam_Id = ""

Old_YBIATAB0.BIATABID = "SAB_CLIENT"
lstParam_K = Mid$(lstParam, 1, 1)
Old_YBIATAB0.BIATABK1 = lstParam_K
Old_YBIATAB0.BIATABK2 = ""

If lstParam_K <> "" Then
    fgParam_Display
    fraParam_Update.Visible = True
    txtParam_Id.Enabled = True
    cmdParam_Add.Visible = True
    
End If

End Sub


Private Sub lstParam_KYCDOSNAT_Click()
Dim xSQL As String, X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

fraYKYCDOS0.Visible = False

fraParam_YKYCDOS0_Update.Visible = False
lstParam_YKYCDOS0.Visible = False
fraParam_YKYCDOS0_JD.Visible = False
cmdParam_YKYCDOS0_Delete.Visible = False
cmdParam_YKYCDOS0_Add.Visible = False
cmdParam_YKYCDOS0_Update.Visible = False
txtParam_KYCDOSSEQ = ""
txtParam_KYCDOSDLIB = ""
fgParam_YKYCDOS0_4c.Visible = False
fraParam_YKYCDOS0_4c.Visible = False
cmdParam_YKYCDOS0_4c_Actualisation.Visible = False

Call rsYKYCDOS0_Init(oldYKYCDOS0)

Select Case Mid$(lstParam_KYCDOSNAT.Text, 1, 1)
    Case "1": mParam_KYCDOSNAT = "D": Call lstParam_YKYCDOS0_Load
    Case "2": mParam_KYCDOSNAT = "J": Call lstParam_YKYCDOS0_Load
    Case "3": mParam_KYCDOSNAT = "*": Call lstParam_YKYCDOS0_Load
    Case "4": mParam_KYCDOSNAT = "=": Call lstParam_YKYCDOS0_Load
                                      Call lstParam_YKYCDOS0_J_Load
                                      Call lstParam_YKYCDOS0_D_Load
                                      If Mid$(lstParam_KYCDOSNAT.Text, 2, 1) = " " Then
                                            blnParam_KYCDOSNAT_4c = False
                                      Else
                                            blnParam_KYCDOSNAT_4c = True
                                      End If

                                    
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub lstParam_YKYCDOS0_Click()
Dim xSQL As String, X As String, K As Long
Me.Enabled = False: Me.MousePointer = vbHourglass

'_______________________________________________
fraParam_YKYCDOS0_Update.Visible = False
fraYKYCDOS0.Visible = False

'_______________________________________________
txtParam_KYCDOSSEQ.Enabled = True

cmdParam_YKYCDOS0_Add.Visible = arrHab(16)
cmdParam_YKYCDOS0_Delete.Visible = arrHab(16)
cmdParam_YKYCDOS0_Update.Visible = arrHab(16)

Call rsYKYCDOS0_Init(oldYKYCDOS0)
oldYKYCDOS0.KYCDOSNAT = mParam_KYCDOSNAT

K = InStr(lstParam_YKYCDOS0.Text, ":")
If K > 0 Then
    Select Case mParam_KYCDOSNAT
        Case "D", "J"
            oldYKYCDOS0.KYCDOSSEQ = Val(Mid$(lstParam_YKYCDOS0.Text, 1, K - 1))
        Case Else
            oldYKYCDOS0.KYCDOSID = Trim(Mid$(lstParam_YKYCDOS0.Text, 1, K - 1))
    End Select
        

    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YKYCDOS0 where KYCDOSNAT = '" & oldYKYCDOS0.KYCDOSNAT & "'" _
         & " and KYCDOSID = '" & oldYKYCDOS0.KYCDOSID & "' and KYCDOSSEQ  = " & oldYKYCDOS0.KYCDOSSEQ
    
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then Call rsYKYCDOS0_GetBuffer(rsSab, oldYKYCDOS0)

End If

txtParam_KYCDOSDECH.Visible = False: lblParam_KYCDOSDECH.Visible = False
Select Case mParam_KYCDOSNAT
    Case "D": chkParam_KYCDOSSTAK.Caption = "Document obligatoire": chkParam_KYCDOSSTAK.Visible = True
              txtParam_KYCDOSSEQ = oldYKYCDOS0.KYCDOSSEQ
              txtParam_KYCDOSDECH = oldYKYCDOS0.KYCDOSDECH
              txtParam_KYCDOSDECH.Visible = True: lblParam_KYCDOSDECH.Visible = True
    Case "J": chkParam_KYCDOSSTAK.Caption = "justificatif obligatoire": chkParam_KYCDOSSTAK.Visible = True
              txtParam_KYCDOSSEQ = oldYKYCDOS0.KYCDOSSEQ
    Case "*": chkParam_KYCDOSSTAK.Visible = False
              txtParam_KYCDOSSEQ = Trim(oldYKYCDOS0.KYCDOSID)
    Case "=":
              If Not blnParam_KYCDOSNAT_4c Then
                lstParam_YKYCDOS0_JD_Load
                fraParam_YKYCDOS0_JD.Visible = True
                mParam_J = 0: mParam_D = 0
                blnYKYCDOS0_JD = False
                libParam_YKYCDOS0_D = "": libParam_YKYCDOS0_J = ""
            Else
                lstParam_YKYCDOS0_4c_Load
                fgParam_YKYCDOS0_4c.Visible = True
                fraParam_YKYCDOS0_4c.Visible = False
                cmdParam_YKYCDOS0_4c_Actualisation.Visible = False
                cmdParam_YKYCDOS0_4c_Actualisation.Caption = "Les caractéristiques de cette classe de clientèle ont été modifiées, il est nécessaire d'actualiser le statut des dossiers concernés :"
                xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YKYCDOS0 where KYCDOSNAT = ''" _
                        & " and KYCDOSDLIB = '" & oldYKYCDOS0.KYCDOSID & "' and KYCDOSSEQ  = 0 and  KYCDOSSEQ2  = 0 "
    
                Set rsSab = cnsab.Execute(xSQL)
                mParam_YKYCDOS0_4c_Actualisation_Nb = rsSab(0)
                cmdParam_YKYCDOS0_4c_Actualisation.Caption = "Les caractéristiques de cette classe de clientèle ont été modifiées, il est nécessaire d'actualiser le statut des " _
                                                           & mParam_YKYCDOS0_4c_Actualisation_Nb & " dossiers concernés."
            End If
End Select
If oldYKYCDOS0.KYCDOSSTAK = " " Then
    chkParam_KYCDOSSTAK.Value = "0"
Else
    chkParam_KYCDOSSTAK.Value = "1"
End If


txtParam_KYCDOSDLIB = Trim(oldYKYCDOS0.KYCDOSDLIB)

If mParam_KYCDOSNAT <> "=" Then fraParam_YKYCDOS0_Update.Visible = True
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub lstParam_YKYCDOS0_D_Click()
Dim K As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass

libParam_YKYCDOS0_D = lstParam_YKYCDOS0_D.Text
K = InStr(lstParam_YKYCDOS0_D.Text, ":")
If K > 0 Then
    mParam_D = Val(Mid$(lstParam_YKYCDOS0_D.Text, 1, K - 1))
End If
If mParam_J > 0 Then
    If newParam_D(mParam_D) = 0 Then
        newParam_D(mParam_D) = mParam_J
        lstParam_YKYCDOS0_JD_Display
    End If
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub lstParam_YKYCDOS0_J_Click()
Dim K As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass

libParam_YKYCDOS0_J = lstParam_YKYCDOS0_J.Text

K = InStr(lstParam_YKYCDOS0_J.Text, ":")
If K > 0 Then
    mParam_J = Val(Mid$(lstParam_YKYCDOS0_J.Text, 1, K - 1))
End If
If newParam_J(mParam_J) = 0 Then
    newParam_J(mParam_J) = mParam_J
    lstParam_YKYCDOS0_JD_Display
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub lstParam_YKYCDOS0_JD_Click()

If lstParam_YKYCDOS0_JD.Visible Then
    Dim K As Integer, K1 As Integer, K2 As Integer
    Me.Enabled = False: Me.MousePointer = vbHourglass
    
    libParam_YKYCDOS0_J = "": libParam_YKYCDOS0_D = ""

    K = InStr(lstParam_YKYCDOS0_JD.Text, ".")
    If K > 0 Then
        K1 = Val(Mid$(lstParam_YKYCDOS0_JD.Text, 1, K - 1))
        K2 = InStr(lstParam_YKYCDOS0_JD.Text, ":")
        If K2 > 0 Then
            K2 = Val(Mid$(lstParam_YKYCDOS0_JD.Text, K + 1, K2 - K - 1))
        End If
        If K2 > 0 Then
            newParam_D(K2) = 0
        Else
            If K1 = mParam_J Then mParam_J = 0: lstParam_YKYCDOS0_J.ListIndex = -1
            newParam_J(K1) = 0
            For K2 = 1 To 999
                If newParam_D(K2) = K1 Then newParam_D(K2) = 0
            Next K2
        End If
            
        lstParam_YKYCDOS0_JD_Display
    End If
    Me.Enabled = True: Me.MousePointer = 0
End If
End Sub


Private Sub lstYKYCDOS0_CLIENACAT_Click()
Dim V, K As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass

K = InStr(lstYKYCDOS0_CLIENACAT.Text, ":")
If MsgBox("Confirmez-vous le type de client : " & vbCrLf & lstYKYCDOS0_CLIENACAT.Text, vbYesNo, "Nouveau dossier Client") = vbYes Then
    Call rsYKYCDOS0_Init(newYKYCDOS0)
    newYKYCDOS0.KYCDOSID = selZCLIENA0.CLIENACLI
    newYKYCDOS0.KYCDOSSTAK = "N"
    newYKYCDOS0.KYCDOSUFCT = "C"
    newYKYCDOS0.KYCDOSDLIB = Trim(Mid$(lstYKYCDOS0_CLIENACAT.Text, 1, K - 1))
    V = cmdParam_YKYCDOS0_Transaction("New")
    If IsNull(V) Then lstYKYCDOS0_CLIENACAT.Visible = False: cmdSelect_SQL_YKYCDOS0_Init
End If

Me.Enabled = True: Me.MousePointer = 0

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
    Case Is = 13:
                If SSTab1.Tab = 0 Then KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select


End Sub

Public Sub cmdContext_Quit()
blnControl = False
lstErr.Clear: lstErr.Height = 200

If fraParam_YKYCDOS0_4c.Visible Then fraParam_YKYCDOS0_4c.Visible = False:  Exit Sub
If fgParam_YKYCDOS0_4c.Visible Then
    fraParam_YKYCDOS0_4c.Visible = False
    fgParam_YKYCDOS0_4c.Visible = False
    cmdParam_YKYCDOS0_4c_Actualisation.Visible = False
    Exit Sub
End If

If fraPJ.Visible Then fraPJ.Visible = False:  Exit Sub
If fraYKYCDOS0_Update.Visible Then fraYKYCDOS0_Update.Visible = False: Exit Sub

If fraYKYCDOS0.Visible Then fraYKYCDOS0.Visible = False: SSTab1.Tab = 0: Exit Sub
If fgDetail.Visible Then fgDetail.Visible = False: Exit Sub
If fraUpdate.Enabled Then
    X = MsgBox("Voulez-vous abandonner la saisie en cours ?", vbQuestion + vbYesNo + vbDefaultButton2, "Mise à jour Mandataires")
    If X = vbYes Then fraUpdate_Init: SSTab1.Tab = 0: Exit Sub
End If

If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
Else
    SSTab1.Tab = SSTab1.Tab - 1
End If

End Sub

Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
    cmdSelect_Ok_Click
Else
    SendKeys "{TAB}"
End If

End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
On Error GoTo Error_Handler

mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False

cnAdo_Open
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

End Sub


Private Sub mnuPrint1_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint1_Liste
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdPrint1_Liste()
Dim iRow As Integer, K As Integer, I As Integer
Dim iRowMax As Integer
Dim blnOk As Boolean
Dim arrX(7) As String
Dim mX0 As String

fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression cmdPrint1_Liste: " & fgSelect.Rows - 1)


prtSAB_Client_Liste_Open '
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
mX0 = ""
For iRow = 1 To fgSelect.Rows - 1
    prtSAB_Client_Liste_NewLine

        fgSelect.Row = iRow
        For I = 0 To 7
            fgSelect.Col = I
            arrX(I) = fgSelect.Text
        Next I
        If mX0 <> arrX(0) Then
            XPrt.DrawWidth = 1
            XPrt.CurrentY = XPrt.CurrentY - 50
            If mX0 <> "" Then XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
            XPrt.CurrentY = XPrt.CurrentY + 50
            mX0 = arrX(0)
        End If
        prtSAB_Client_Liste_Line arrX()
    
Next iRow

XPrt.DrawWidth = 5
prtSAB_Client_Liste_NewLine
XPrt.CurrentY = XPrt.CurrentY - 50
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
prtSAB_Client_Liste_Close
fgSelect.Visible = True
Me.Show

End Sub

Private Sub mnuPrint2_Excel_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String

Call lstErr_AddItem(lstErr, cmdContext, "> Exportation EXCEL ... " & fgDetail.Rows & " lignes"): DoEvents

Select Case SSTab1.Tab
    Case 0:
        Select Case mSelect_SQL
            Case "KYC gsop"
                X = "Liste Client"
                Call MSflexGrid_Excel("", "GSOP", X, fgDetail, fgDetail.Cols - 1)
            Case "KYC ech"
                X = "Echéancier des documents justificatifs"
                Call MSflexGrid_Excel("", "GSOP", X, fgDetail, fgDetail.Cols - 1)
            Case "KYC Releve"
                X = "GSOP - Relevé de compte"
                Call MSflexGrid_Excel("", "GSOP", X, fgDetail, fgDetail.Cols - 1)
        End Select
    Case 2:
        If mSelect_SQL = "KYC Ctl" Then
            X = "Contrôle KYC : SAB / GSOP"
            Call MSflexGrid_Excel("", "GSOP", X, fgSelect, 9)
        Else
            If optUpdLog_YKYCDOS0 Then
                X = "Historique des mises à jour des dossiers KYC"
            Else
                X = "Historique des mises à jour des mandataires"
            End If
            
            Call MSflexGrid_Excel("", "GSOP", X, fgSelect, fgSelect.Cols - 1)
        End If
    Case 1
        If blnSelect_YKYCDOS0 Then
            X = "GSOP - dossier du client : " & libYKYCDOS0_CLIENACLI & " - " & libYKYCDOS0_CLIENARA1 _
                   & vbLf & "Type de clientèle : " & fraYKYCDOS_ZCLIENA0 _
                   & "  Gestionnaire : " & libYKYCDOS0_CLIENARES
            
            Call MSflexGrid_Excel("", "GSOP", X, fgYKYCDOS0, fgYKYCDOS0.Cols - 1)
            fgYKYCDOS0.LeftCol = 0
        End If
End Select
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation EXCEL terminée "): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuPrint2_Mail_Click()
Dim xObjet As String, xMesg As String, xDest As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_AddItem(lstErr, cmdContext, "> Envoi par Mail ... " & fgDetail.Rows & " lignes"): DoEvents

Select Case SSTab1.Tab
    Case 0:
        Select Case mSelect_SQL
            Case "KYC gsop"
                xObjet = "GSOP - Liste Client"
                xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
                 & xObjet
        
                Call MSFlexGrid_SendMail(currentSSIWINMAIL, "GSOP", xObjet, xMesg, fgDetail, fgDetail.Cols - 1)
            Case "KYC ech"
                xObjet = "GSOP - Echéancier des documents justificatifs"
                xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
                 & xObjet
                If blnAuto Then
                    xDest = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire("S11")
                Else
                    xDest = currentSSIWINMAIL
                End If
                Call MSFlexGrid_SendMail(xDest, "GSOP", xObjet, xMesg, fgDetail, fgDetail.Cols - 1)
            Case "KYC Releve"
                xObjet = "GSOP - Relevé de compte"
                xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
                 & xObjet
                If blnAuto Then
                    xDest = frmElpPrt.prtIMP_PDF_NoPaper_Destinaire("S11")
                Else
                    xDest = currentSSIWINMAIL
                End If
                Call MSFlexGrid_SendMail(xDest, "GSOP", xObjet, xMesg, fgDetail, fgDetail.Cols - 1)
        End Select

    Case 2:
         If mSelect_SQL = "KYC Ctl" Then
            xObjet = "GSOP : Contrôle KYC : SAB / GSOP"
            xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
                 & xObjet
            Call MSFlexGrid_SendMail(currentSSIWINMAIL, "GSOP", xObjet, xMesg, fgSelect, 9)
        Else
           If optUpdLog_YKYCDOS0 Then
                xObjet = "GSOP - Historique des mises à jour des dossiers KYC"
            Else
                xObjet = "GSOP - Historique des mises à jour des mandataires"
            End If
            xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
                 & xObjet
        
            Call MSFlexGrid_SendMail(currentSSIWINMAIL, "GSOP", xObjet, xMesg, fgSelect, fgSelect.Cols - 1)
        End If
    Case 1
        If blnSelect_YKYCDOS0 Then
            xObjet = "GSOP - dossier du client : " & libYKYCDOS0_CLIENACLI & " - " & libYKYCDOS0_CLIENARA1
            
            xMesg = "<span style='font-size:8.0pt;font-family:Calibri'>" _
                   & htmlFontColor_Black & "<pre>GSOP - dossier du client : " & htmlFontColor_Blue & "<B>" & libYKYCDOS0_CLIENACLI & " - " & libYKYCDOS0_CLIENARA1 & "</B><BR>" _
                   & htmlFontColor_Black & "<pre>Type de clientèle        : " & htmlFontColor_Blue & fraYKYCDOS_ZCLIENA0 & "<BR>" _
                   & htmlFontColor_Black & "<pre>Gestionnaire             : " & htmlFontColor_Blue & libYKYCDOS0_CLIENARES & "<BR>" _
                   & htmlFontColor_Black & "<pre>Nationalité              : " & htmlFontColor_Blue & libYKYCDOS0_CLIENANAT & "<BR>" _
                   & htmlFontColor_Black & "<pre>Résidence                : " & htmlFontColor_Blue & libYKYCDOS0_CLIENARSD & "<BR>"
                   
                 Dim X As String, objFolder, objFiles, fsoFile As File, K1 As Integer, K2 As Integer, blnOk As Boolean
                
                X = currentYKYCDOS0.KYCDOSID & "\"
                If Dir(paramGSOP_Dossier_Path_DROPI & X) <> "" Then
                    Set objFolder = msFileSystem.GetFolder(paramGSOP_Dossier_Path_DROPI & X)
                    Set objFiles = objFolder.Files
                    
                   
                   
                    For Each fsoFile In objFiles
                        X = "pièce jointe"
                        K1 = InStr(1, fsoFile.Name, "_") + 1
                        If K1 > 0 Then
                            K2 = InStr(K1, fsoFile.Name, "_")
                            If K2 > 0 Then
                                K1 = Mid$(fsoFile.Name, K1, K2 - K1)
                                If K1 > 0 And K1 < 999 Then X = Trim(arrKYCDOSDLIB_D(K1))
                            End If
                        End If
                        
                        X = "<a href=" & Asc34 _
                           & paramGSOP_Dossier_Path_DROPI & currentYKYCDOS0.KYCDOSID & "\" & fsoFile.Name _
                           & Asc34 & "><pre>" & "<span style='font-size:8.0pt;font-family:Calibri'>" & htmlFontColor_Green & X & "   " & htmlFontColor_Gray & fsoFile.DateCreated & ""

                        xMesg = xMesg & X
                    Next
                End If
           
            Call MSFlexGrid_SendMail(currentSSIWINMAIL, "GSOP", xObjet, xMesg, fgYKYCDOS0, fgYKYCDOS0.Cols - 1)
            fgYKYCDOS0.LeftCol = 0
        End If
End Select
Call lstErr_AddItem(lstErr, cmdContext, "< Envoi MAIL terminé "): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuSelect_Print_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
'cmdPrint_List1_Ok
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub optSelect_Add_Click()
 blnSelect_YKYCDOS0 = False

End Sub

Private Sub optSelect_Options_3C_Click()
cmdSelect_Clear

End Sub

Private Sub optSelect_Options_3N_Click()
cmdSelect_Clear

End Sub


Private Sub optSelect_Options_KYCgsop_All_Click()
cmdSelect_Clear

End Sub

Private Sub optSelect_Options_KYCgsop_Detail_Missing_Click()
cmdSelect_Clear

End Sub


Private Sub optSelect_Options_KYCgsop_Detail_NOK_Click()
cmdSelect_Clear

End Sub


Private Sub optSelect_Options_KYCgsop_Detail_OK_Click()
cmdSelect_Clear

End Sub


Private Sub optSelect_Options_KYCgsop_NOK_Click()
cmdSelect_Clear

End Sub


Private Sub optSelect_Options_KYCgsop_OK_Click()
cmdSelect_Clear

End Sub


Private Sub optSelect_YKYCDOS0_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdSelect_SQL_YKYCDOS0_Init

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdSelect_SQL_YKYCDOS0_Init()
Dim X As String
blnSelect_YKYCDOS0 = optSelect_YKYCDOS0.Value

fraYKYCDOS0_Update.Visible = False
fraYKYCDOS0.Visible = False
fgYKYCDOS0.Visible = False
fraPJ.Visible = False
cmdYKYCDOS0_Delete_All.Visible = False
lstParam_YKYCDOS0.Visible = False

SSTab1.Tab = 1
SSTab1.Caption = "Mise à jour des dossiers CLIENT"

'_______________________________________________________________________________________________________

libYKYCDOS0_CLIENACLI.Caption = selZCLIENA0.CLIENACLI
libYKYCDOS0_CLIENARA1.Caption = Trim(selZCLIENA0.CLIENARA1) & " " & Trim(selZCLIENA0.CLIENARA2)
Select Case selZCLIENA0.CLIENACOL
    Case 0: libYKYCDOS0_CLIENACOL.Caption = "(client)"
    Case 1: libYKYCDOS0_CLIENACOL.Caption = "(collectif)"
    Case 2: libYKYCDOS0_CLIENACOL.Caption = "(autre)"
    Case Else: libYKYCDOS0_CLIENACOL.Caption = "??? " & selZCLIENA0.CLIENACOL
End Select

X = "select * from  " & paramIBM_Library_SAB & ".ZBASTAB0 " _
    & " where  BASTABETA = 1 and BASTABNUM = 5 and BASTABARG = 'CLI" & selZCLIENA0.CLIENAETA & "' "
Set rsSab = cnsab.Execute(X)
If Not rsSab.EOF Then
    libYKYCDOS0_CLIENAETA.Caption = selZCLIENA0.CLIENAETA & " - " & rsSab("BASTABLO2") & Trim(Mid$(rsSab("BASTABDON"), 1, 16))
Else
    libYKYCDOS0_CLIENAETA.Caption = selZCLIENA0.CLIENAETA
End If
libYKYCDOS0_CLIENAETA.ForeColor = vbMagenta


X = "select * from  " & paramIBM_Library_SAB & ".ZBASTAB0 " _
    & " where  BASTABETA = 1 and BASTABNUM = 8 and BASTABARG = 'CLI" & selZCLIENA0.CLIENACAT & "' "
Set rsSab = cnsab.Execute(X)
If Not rsSab.EOF Then
    libYKYCDOS0_CLIENACAT.Caption = selZCLIENA0.CLIENACAT & " - " & rsSab("BASTABLO2") & Trim(Mid$(rsSab("BASTABDON"), 1, 16))
Else
    libYKYCDOS0_CLIENACAT.Caption = selZCLIENA0.CLIENACAT
End If
libYKYCDOS0_CLIENACAT.ForeColor = vbMagenta

X = "select * from  " & paramIBM_Library_SAB & ".ZBASTAB0 " _
    & " where  BASTABETA = 1 and BASTABNUM = 6 and BASTABARG = 'CLI" & selZCLIENA0.CLIENARES & "' "
Set rsSab = cnsab.Execute(X)
If Not rsSab.EOF Then
    libYKYCDOS0_CLIENARES.Caption = selZCLIENA0.CLIENARES & " - " & Trim(Mid$(rsSab("BASTABDON"), 24, 10))
Else
    libYKYCDOS0_CLIENARES.Caption = selZCLIENA0.CLIENARES
End If


X = "select * from  " & paramIBM_Library_SAB & ".ZBASTAB0 " _
    & " where  BASTABETA = 1 and BASTABNUM = 11 and BASTABARG = 'CLI" & selZCLIENA0.CLIENANAT & "' "
Set rsSab = cnsab.Execute(X)
If Not rsSab.EOF Then
    libYKYCDOS0_CLIENANAT.Caption = selZCLIENA0.CLIENANAT & " - " & Trim(Mid$(rsSab("BASTABLO2"), 4, 16) & Mid$(rsSab("BASTABDON"), 1, 16))
Else
    libYKYCDOS0_CLIENANAT.Caption = selZCLIENA0.CLIENANAT
End If

X = "select * from  " & paramIBM_Library_SAB & ".ZBASTAB0 " _
    & " where  BASTABETA = 1 and BASTABNUM = 11 and BASTABARG = 'CLI" & selZCLIENA0.CLIENARSD & "' "
Set rsSab = cnsab.Execute(X)
If Not rsSab.EOF Then
    libYKYCDOS0_CLIENARSD.Caption = selZCLIENA0.CLIENARSD & " - " & Trim(Mid$(rsSab("BASTABLO2"), 4, 16) & Mid$(rsSab("BASTABDON"), 1, 16))
Else
    libYKYCDOS0_CLIENARSD.Caption = selZCLIENA0.CLIENARSD
End If

'_______________________________________________________________________________________________________


fgYKYCDOS0_ZADRESS0.Clear
fgYKYCDOS0_ZADRESS0.FormatString = fgYKYCDOS0_ZADRESS0_FormatString
fgYKYCDOS0_ZADRESS0.Visible = False

fgYKYCDOS0_ZADRESS0.Rows = 1
X = "select * from  " & paramIBM_Library_SAB & ".ZADRESS0 " _
    & " where  ADRESSETA = 1 and ADRESSNUM  like '%" & Val(selZCLIENA0.CLIENACLI) & "%' order by ADRESSTYP , ADRESSCOA , ADRESSNUM"
Set rsSab = cnsab.Execute(X)
    
Do While Not rsSab.EOF

    fgYKYCDOS0_ZADRESS0.Rows = fgYKYCDOS0_ZADRESS0.Rows + 1
    fgYKYCDOS0_ZADRESS0.Row = fgYKYCDOS0_ZADRESS0.Rows - 1
    fgYKYCDOS0_ZADRESS0.RowHeight(fgYKYCDOS0_ZADRESS0.Row) = 1300
    fgYKYCDOS0_ZADRESS0.Col = 0: fgYKYCDOS0_ZADRESS0.Text = rsSab("ADRESSCOA")
    fgYKYCDOS0_ZADRESS0.Col = 1: fgYKYCDOS0_ZADRESS0.Text = rsSab("ADRESSNUM")
    X = Trim(rsSab("ADRESSRA1")) & Trim(rsSab("ADRESSRA2"))_
      & vbCrLf & Trim(rsSab("ADRESSAD1")) _
      & vbCrLf & Trim(rsSab("ADRESSAD2")) _
      & vbCrLf & Trim(rsSab("ADRESSAD3")) _
      & vbCrLf & Trim(rsSab("ADRESSCOP")) & " " & Trim(rsSab("ADRESSVIL")) _
      & vbCrLf & Trim(rsSab("ADRESSPAY"))
      
    fgYKYCDOS0_ZADRESS0.Col = 2: fgYKYCDOS0_ZADRESS0.Text = X
    rsSab.MoveNext

Loop
fgYKYCDOS0_ZADRESS0.Visible = True

If fgYKYCDOS0_ZADRESS0.Rows > 2 Then
    fgYKYCDOS0_ZADRESS0.BackColorFixed = mColor_Y2
    fgYKYCDOS0_ZADRESS0.FormatString = Replace(fgYKYCDOS0_ZADRESS0_FormatString, "Intitulé", fgYKYCDOS0_ZADRESS0.Rows - 1 & " adresses")

Else
    fgYKYCDOS0_ZADRESS0.BackColorFixed = mColor_G1
End If

'_______________________________________________________________________________________________________


X = "select * from  " & paramIBM_Library_SABSPE & ".YKYCDOS0 " _
    & " where  KYCDOSNAT = ' ' and KYCDOSID = '" & selZCLIENA0.CLIENACLI & "'  and KYCDOSSEQ = 0  and KYCDOSSEQ2 = 0 "
Set rsSab = cnsab.Execute(X)
If Not rsSab.EOF Then

    Call rsYKYCDOS0_GetBuffer(rsSab, currentYKYCDOS0)
    
    X = "select * from  " & paramIBM_Library_SABSPE & ".YKYCDOS0 " _
        & " where  KYCDOSNAT = '*'and KYCDOSID = '" & currentYKYCDOS0.KYCDOSDLIB & "'  and KYCDOSSEQ = 0 "
    Set rsSab = cnsab.Execute(X)
    If Not rsSab.EOF Then
        fraYKYCDOS_ZCLIENA0.Caption = Trim(rsSab("KYCDOSID")) & " : " & Trim(rsSab("KYCDOSDLIB"))
    Else
        fraYKYCDOS_ZCLIENA0.Caption = Trim(rsSab("KYCDOSID")) & " ??????????????????"
    End If
    If arrHab(15) Then
        fraYKYCDOS_ZCLIENA0.ForeColor = vbMagenta
    Else
        fraYKYCDOS_ZCLIENA0.ForeColor = vbBlue
    End If
    fgYKYCDOS0_Display
    If arrHab(16) Then cmdYKYCDOS0_Delete_All.Visible = True
Else
    If arrHab(15) Then
        fraYKYCDOS_ZCLIENA0.Caption = "Préciser le type de clientèle"
        fraYKYCDOS_ZCLIENA0.ForeColor = vbRed
        lstYKYCDOS0_CLIENACAT_Load
        lstYKYCDOS0_CLIENACAT.Visible = True
    End If
End If

'_______________________________________________________________________________________________________




fraYKYCDOS0.Visible = True
End Sub

Private Sub optUpdate_Add_Old_Click()
fraUpdate_Détail.Enabled = False
If mUpdate_Node_Key <> "" Then
    blnUpdate_Ok = True
Else
    blnUpdate_Ok = False
End If
cmdUpdate_Ok.Caption = "Ajouter un lien"
End Sub

Private Sub optUpdate_Add_PP_Click()
fraUpdate_Détail_PP.Visible = True
fraUpdate_Détail_Sté.Visible = False

fraUpdate_Détail.Enabled = True
blnUpdate_Ok = False
cmdUpdate_Ok.Caption = "Créér un TIERS + lien"

End Sub


Private Sub optUpdate_Add_Sté_Click()
fraUpdate_Détail_PP.Visible = False
fraUpdate_Détail_Sté.Visible = True
fraUpdate_Détail.Enabled = True
blnUpdate_Ok = False
cmdUpdate_Ok.Caption = "Créér un TIERS + lien"

End Sub


Private Sub optUpdLog_YKYCDOS0_Click()
lblUpdLog_CLIRGPREG.Visible = optUpdLog_YUPDLOG0.Value
txtUpdLog_CLIRGPREG.Visible = optUpdLog_YUPDLOG0.Value
fgSelect.Visible = False
End Sub

Private Sub optUpdLog_YUPDLOG0_Click()
lblUpdLog_CLIRGPREG.Visible = optUpdLog_YUPDLOG0.Value
txtUpdLog_CLIRGPREG.Visible = optUpdLog_YUPDLOG0.Value
fgSelect.Visible = False
End Sub

Private Sub rtfPJ_Change()
If Len(rtfPJ.Text) > 0 Then
    cmdPJ_OK.Visible = True
Else
    cmdPJ_OK.Visible = False
End If

End Sub

Private Sub rtfPJ_Click()
rtfPJ.Top = 120  '!!!! cf cmdPJ_Ok_Click
rtfPJ.Left = 120
rtfPJ.Width = fraPJ.Width - 240
rtfPJ.Height = fraPJ.Height - 720

rtfPJ.ZOrder 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then cmdSelect_Ok.SetFocus

End Sub














Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
   ' Case 1: fgSAA.LeftCol = 0
End Select
End Sub


Public Sub blnTransaction_Set()
If Not blnTransaction Then
    blnTransaction = True
   ' Set rsADO_Update = cnado.Execute("SET TRANSACTION ISOLATION LEVEL READ COMMITTED")

End If

End Sub


Public Sub lblSelect_Display(lCLIENACLI As String, lblX As Label)
Dim V

Set rsAdo = Nothing
X = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0  where CLIENACLI = '" & lCLIENACLI & "'"
'Lecture CLIENT
'---------------
Set rsAdo = cnAdo.Execute(X)
If Not rsAdo.EOF Then
    V = rsZCLIENA0_GetBuffer(rsAdo, xZCLIENA0)
Else
    MsgBox lCLIENACLI & " : client inconnu", vbCritical, Me.Name & " lblSelect_Display"
End If

X = "select * from " & paramIBM_Library_SAB & ".ZCLIENB0  where CLIENBCLI = '" & lCLIENACLI & "'"
'Lecture CLIENT
'---------------
Set rsAdo = cnAdo.Execute(X)
If Not rsAdo.EOF Then
    V = rsZCLIENB0_GetBuffer(rsAdo, xZCLIENB0)
Else
    MsgBox lCLIENACLI & " : clientB inconnu", vbCritical, Me.Name & " lblSelect_Display"
End If
'Lecture ADRESSE
'---------------
X = "select * from " & paramIBM_Library_SAB & ".ZADRESS0" _
    & " where ADRESSNUM = ' " & lCLIENACLI & "'" _
    & " and ADRESSCOA = '  ' and ADRESSPLA = 0 and ADRESSETA = 1"

Set rsAdo = cnAdo.Execute(X)
If Not rsAdo.EOF Then
    V = rsZADRESS0_GetBuffer(rsAdo, xZADRESS0)
Else
    MsgBox lCLIENACLI & " : adresse inconnue", vbCritical, Me.Name & " lblSelect_Display"
End If

'Call lstErr_Clear(lstErr, cmdContext, "> Affichage : " & keyX): DoEvents
'Affichage
'---------------------------------------------------
lblX = xZCLIENA0.CLIENACLI & vbCrLf _
            & Trim(xZCLIENA0.CLIENARA1) & " " & Trim(xZCLIENA0.CLIENARA2) & vbCrLf _
            & xZADRESS0.ADRESSAD1 & vbCrLf _
            & xZADRESS0.ADRESSAD2 & vbCrLf _
            & xZADRESS0.ADRESSAD3 & vbCrLf _
            & xZADRESS0.ADRESSVIL & "  " & xZADRESS0.ADRESSPAY & vbCrLf _

End Sub

Public Sub cnAdo_Close()
On Error Resume Next

cnAdo.Close
Set cnAdo = Nothing


End Sub

Public Sub cnAdo_Open()
On Error GoTo Error_Handler
Dim X As String

cnAdo.Open paramODBC_DSN_SAB


Exit Sub

Error_Handler:
cnAdo.Open "SAB073_MDB"
blnControl = False
If Not blnAuto Then MsgBox Error

End Sub


Private Sub tvwInverse_NodeClick(ByVal Node As ComctlLib.Node)
Dim lenX As Integer

Me.Enabled = False

lenX = Len(Node.key)
Select Case lenX
    Case 10:
            lblSelect_Display Mid$(Node.key, 4, 7), lblInverse
    Case 20:
            lblSelect_Display Mid$(Node.key, 14, 7), lblInverse
End Select
Me.Enabled = True

End Sub

Private Sub tvwSelect_NodeClick(ByVal Node As ComctlLib.Node)
Dim lenX As Integer
Dim X As String
Me.Enabled = False
mSelect_Node_Key = Node.key
fraSelect_Update.Visible = False
fraUpdate.Visible = False
tvwInverse.Nodes.Clear
tvwInverse.Visible = True 'False
lblSelect = ""
lblInverse = ""
fgDetail.Visible = False

lenX = Len(Node.key)

Select Case mSelect_SQL
Case 1, 3
    Select Case lenX
        Case 10:
                X = Mid$(Node.key, 4, 7)
                tvwSelect_Display_10 X
                
                If mSelect_SQL = 3 Then fgDetail_Display X
              'fraSelect_Update_Display Mid$(Node.key, 1, 10)
        Case 20:
                X = Mid$(Node.key, 14, 7)
                tvwInverse_Display_ZCLIGRP0_Reset X
                lblSelect_Display X, lblSelect
                oldZCLIENA0 = xZCLIENA0
                oldZCLIENB0 = xZCLIENB0
                oldZADRESS0 = xZADRESS0
                optUpdate_Add_Old.Value = True
                fraUpdate_Détail_Display
    
                tvwInverse.Visible = True
                
               If Mid$(Node.key, 14, 2) = "99" Then
                    fraSelect_Update.Visible = True
                    optSelect_Add.Enabled = False
                    optSelect_Modification.Enabled = True
                    optSelect_Suppress.Enabled = arrHab(3)
                    optSelect_Modification.Value = True
                    optSelect_YKYCDOS0.Enabled = False
                    
                    optSelect_CLIENARES_Init
                    
                End If
    
    End Select
Case 2
    Select Case lenX
        Case 10:
                X = Mid$(Node.key, 4, 7)
                tvwSelect2_Display_10 X
                oldZCLIENA0 = xZCLIENA0
                oldZCLIENB0 = xZCLIENB0
                oldZADRESS0 = xZADRESS0
                optSelect_CLIENARES_Init
                fraUpdate_Détail_Display
    End Select
End Select
Me.Enabled = True




End Sub

Private Sub tvwUpdate_NodeClick(ByVal Node As ComctlLib.Node)
Dim lenX As Integer

Me.Enabled = False
mUpdate_Node_Key = Node.key
lenX = Len(Node.key)
Select Case lenX
    Case 10:
            lblSelect_Display Mid$(Node.key, 4, 7), lblInverse
            
            oldZCLIENA0 = xZCLIENA0
            oldZCLIENB0 = xZCLIENB0
            oldZADRESS0 = xZADRESS0
            fraUpdate_Détail_Display
            optUpdate_Add_Old.Value = True
   Case 20:
'            lblSelect_Display Mid$(Node.key, 14, 7), lblUpdate
End Select
Me.Enabled = True

End Sub


Private Sub txtParam_Id_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtParam_KYCDOSDECH_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtParam_KYCDOSSEQ_KeyPress(KeyAscii As Integer)
Select Case mParam_KYCDOSNAT
    Case "D", "J": Call num_KeyAscii(KeyAscii)
        
End Select

End Sub


Private Sub txtSelect_CLIENARA1_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_CLIENARA1_GotFocus()
txt_GotFocus txtSelect_CLIENARA1
End Sub

Private Sub txtSelect_CLIENARA1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub



Public Sub tvwInverse_Display_ZCLIGRP0(lCLIGRPREG As String)
Dim wCLIGRPREL As String, keyCLIGRPREL As String
Dim wCLIGRPCLI As String
Dim wCli As String
Dim blnZCLIGRP0 As Boolean
Dim X As String
Dim xRA1 As String
blnZCLIGRP0 = False
wCLIGRPREL = ""
wCli = "CLI" & lCLIGRPREG

For I = 1 To arrZCLIGRP0_Nb
        If arrZCLIGRP0(I).CLIGRPREG = lCLIGRPREG Then
    
            blnZCLIGRP0 = True
            wCLIGRPCLI = arrZCLIGRP0(I).CLIGRPCLI
            X = arrZCLIGRP0(I).CLIGRPREL
            If wCLIGRPREL <> X Then
                wCLIGRPREL = X
                X = libCLIGRPREL(wCLIGRPREL)
                keyCLIGRPREL = "CLI" & lCLIGRPREG & wCLIGRPREL
                tvwInverse.Nodes.Add wCli, tvwChild, keyCLIGRPREL, X
                tvwInverse.Nodes(keyCLIGRPREL).Expanded = True
            End If
            xRA1 = Trim(arrZCLIGRP0(I).CLIGRPCLI_RA1) & " " & Trim(arrZCLIGRP0(I).CLIGRPCLI_RA2)
            If Mid$(arrZCLIGRP0(I).CLIGRPCLI_RES, 1, 1) = "X" Then xRA1 = "## " & LCase$(xRA1) & " ##"
            xRA1 = xRA1 & "  [" & arrZCLIGRP0(I).CLIGRPCLI_NAT & "]"
            If arrZCLIGRP0(I).CLIGRPTAU <> 0 Then xRA1 = xRA1 & "  % " & arrZCLIGRP0(I).CLIGRPTAU
            tvwInverse.Nodes.Add keyCLIGRPREL, tvwChild, keyCLIGRPREL & wCLIGRPCLI, wCLIGRPCLI & "  " & xRA1
        End If
'    End If
Next I
If blnZCLIGRP0 Then
'    tvwInverse.Nodes(wCLI).Expanded = True
    tvwInverse.Nodes(wCli).Sorted = True
End If

End Sub

Public Sub tvwSelect_Display_ZCLIGRP0(lCLIGRPCLI As String)
Dim wCLIGRPREL As String, keyCLIGRPREL As String
Dim wCLIGRPREG As String
Dim wCli As String
Dim blnZCLIGRP0 As Boolean
Dim X As String
Dim xRA1 As String
blnZCLIGRP0 = False
wCLIGRPREL = ""
wCli = "CLI" & lCLIGRPCLI

'For I = arrZCLIGRP0_Index To arrZCLIGRP0_Nb
'    If arrZCLIGRP0(I).CLIGRPCLI > lCLIGRPCLI Then
'        arrZCLIGRP0_Index = I
'        Exit For
'    Else
For I = 1 To arrZCLIGRP0_Nb
        If arrZCLIGRP0(I).CLIGRPCLI = lCLIGRPCLI Then
    
            blnZCLIGRP0 = True
            wCLIGRPREG = arrZCLIGRP0(I).CLIGRPREG
            X = arrZCLIGRP0(I).CLIGRPREL
            If wCLIGRPREL <> X Then
                wCLIGRPREL = X
                X = libCLIGRPREL(wCLIGRPREL)
                keyCLIGRPREL = "CLI" & lCLIGRPCLI & wCLIGRPREL
                tvwSelect.Nodes.Add wCli, tvwChild, keyCLIGRPREL, X
                tvwSelect.Nodes(keyCLIGRPREL).Expanded = True
            End If
            xRA1 = Trim(arrZCLIGRP0(I).CLIGRPCLI_RA1) & " " & Trim(arrZCLIGRP0(I).CLIGRPCLI_RA2)
            If Mid$(arrZCLIGRP0(I).CLIGRPCLI_RES, 1, 1) = "X" Then xRA1 = "## " & LCase$(xRA1) & " ##"
            xRA1 = xRA1 & "  [" & arrZCLIGRP0(I).CLIGRPCLI_NAT & "]"
            If arrZCLIGRP0(I).CLIGRPTAU <> 0 Then xRA1 = xRA1 & "  % " & arrZCLIGRP0(I).CLIGRPTAU
            tvwSelect.Nodes.Add keyCLIGRPREL, tvwChild, keyCLIGRPREL & wCLIGRPREG, wCLIGRPREG & "  " & xRA1
        End If
'    End If
Next I
If blnZCLIGRP0 Then
'    tvwSelect.Nodes(wCLI).Expanded = True
    tvwSelect.Nodes(wCli).Sorted = True
End If

End Sub


Public Sub tvwSelect_Display_ZCLIGRP0_Reset(lCLIENACLI As String)
Dim wCLIGRPREL As String, keyCLIGRPREL As String
Dim wCLIGRPREG As String
Dim wCli As String
Dim X As String
Dim xNode As Node
Dim xRA1 As String

wCli = "CLI" & lCLIENACLI
tvwSelect.Nodes.Remove (wCli)
'Set xNode = tvwSelect.Node(tvwSelect.SelectedItem.Index).Child
'tvwSelect.Nodes.Remove xNode.Index
'Exit Sub
Set rsAdo = Nothing
X = "select CLIENARA1,CLIENARA2,CLIENARES,CLIENANAT from " & paramIBM_Library_SAB & ".ZCLIENA0  where CLIENACLI = '" & lCLIENACLI & "'"

Set rsAdo = cnAdo.Execute(X)
If Not rsAdo.EOF Then
    xRA1 = Trim(rsAdo("CLIENARA1")) & " " & Trim(rsAdo("CLIENARA2"))
Else
    xRA1 = "?????????????????????"
End If
If Mid$(rsAdo("CLIENARES"), 1, 1) = "X" Then xRA1 = "## " & LCase$(xRA1) & " ##"
tvwSelect.Nodes.Add , , wCli, lCLIENACLI & "   " & xRA1 & " [" & rsAdo("CLIENANAT") & "]"
tvwSelect.Nodes(wCli).Sorted = True

''arrZCLIGRP0_sql " G.CLIGRPREG = C.CLIENACLI and CLIGRPCLI = '" & lCLIENACLI & "' order by CLIGRPCLI, CLIGRPREL, CLIGRPREG"
arrZCLIGRP0_sql " CLIGRPCLI = '" & lCLIENACLI & "' order by CLIGRPCLI, CLIGRPREL, CLIGRPREG", True

tvwSelect_Display_ZCLIGRP0 lCLIENACLI
tvwSelect.Nodes(wCli).Expanded = True
tvwSelect.Nodes(wCli).Selected = True
End Sub
Public Sub tvwInverse_Display_ZCLIGRP0_Reset(lCLIENACLI As String)
Dim wCLIGRPREL As String, keyCLIGRPREL As String
Dim wCLIGRPREG As String
Dim wCli As String
Dim X As String, xWhere As String
Dim xNode As Node
Dim xRA1 As String
wCli = "CLI" & lCLIENACLI
tvwInverse.Nodes.Clear
Set rsAdo = Nothing
X = "select CLIENARA1, CLIENARES from " & paramIBM_Library_SAB & ".ZCLIENA0  where CLIENACLI = '" & lCLIENACLI & "'"

Set rsAdo = cnAdo.Execute(X)
If Not rsAdo.EOF Then
    xRA1 = Trim(rsAdo("CLIENARA1"))
Else
    xRA1 = "?????????????????????"
End If
xRA1 = Trim(arrZCLIGRP0(I).CLIGRPCLI_RA1) & " " & Trim(arrZCLIGRP0(I).CLIGRPCLI_RA2)
If Mid$(arrZCLIGRP0(I).CLIGRPCLI_RES, 1, 1) = "X" Then xRA1 = "## " & LCase$(xRA1) & " ##"

tvwInverse.Nodes.Add , , wCli, lCLIENACLI & "   " & xRA1
tvwInverse.Nodes(wCli).Sorted = True

'arrZCLIGRP0_sql " G.CLIGRPCLI = C.CLIENACLI and CLIGRPREG = '" & lCLIENACLI & "' order by CLIGRPREG, CLIGRPREL, CLIGRPCLI"

If Mid$(lCLIENACLI, 1, 1) = "9" Then
    xWhere = " CLIGRPREG like '9%" & Mid$(lCLIENACLI, 3, 5)
Else
    xWhere = " CLIGRPREG = '" & lCLIENACLI
End If

'arrZCLIGRP0_sql xWhere & "' order by CLIGRPREG, CLIGRPREL, CLIGRPCLI", False
arrZCLIGRP0_sql xWhere & "' order by  CLIGRPREL, CLIGRPCLI", False

tvwInverse_Display_ZCLIGRP0 lCLIENACLI
tvwInverse.Nodes(wCli).Expanded = True
tvwInverse.Nodes(wCli).Selected = True
End Sub


Public Sub fraUpdate_Init()
On Error Resume Next
fraUpdate.Enabled = False
fraUpdate.Visible = False
tvwUpdate.Nodes.Clear
fraUpdate_Détail.Caption = ""
cboUpdate_Add.ListIndex = 1
cbo_Scan "MR ", cboUpdate_CLIENAETA
txtUpdate_CLIENAFIL.Visible = False

txtUpdate_CLIENASIG = ""
txtUpdate_CLIENARA1 = ""
txtUpdate_CLIENARA2 = ""
txtUpdate_ADRESSAD1 = ""
txtUpdate_ADRESSAD2 = ""
txtUpdate_ADRESSAD3 = ""
txtUpdate_ADRESSVIL = ""
txtUpdate_ADRESSCOP = ""
cbo_Scan "FR", cboUpdate_ADRESSPAY
cbo_Scan "FR", cboUpdate_CLIENANAT
cbo_Scan "FR", cboUpdate_CLIENARSD
cbo_Scan "FR", cboUpdate_CLIENBNAS

Call DTPicker_Set(txtUpdate_CLIENADNA, "19000101")
txtUpdate_CLIENBCOM = ""
txtUpdate_CLIENBCIN = ""
txtUpdate_CLIENBINS = ""
txtUpdate_CLIENAFIL = ""
cboUpdate_CLIENBLIE.ListIndex = 0

txtUpdate_CLIENASRN = "NEANT"
txtUpdate_CLIENAREG = ""
cboUpdate_CLIENBTER.ListIndex = 0

txtUpdate_CLIENARA1.SetFocus
End Sub

Public Sub fraUpdate_Détail_Display()
On Error Resume Next
If Trim(oldZCLIENA0.CLIENASRN) = "" Then
    optUpdate_Add_PP = True
Else
    optUpdate_Add_Sté = True
End If


fraUpdate_Détail.Caption = oldZCLIENA0.CLIENACLI & " " & Trim(oldZCLIENA0.CLIENARA1)
cbo_Scan oldZCLIENA0.CLIENAETA, cboUpdate_CLIENAETA
txtUpdate_CLIENASIG = Trim(oldZCLIENA0.CLIENASIG)
txtUpdate_CLIENARA1 = Trim(oldZCLIENA0.CLIENARA1)
txtUpdate_CLIENARA2 = Trim(oldZCLIENA0.CLIENARA2)
txtUpdate_ADRESSAD1 = Trim(oldZADRESS0.ADRESSAD1)
txtUpdate_ADRESSAD2 = Trim(oldZADRESS0.ADRESSAD2)
txtUpdate_ADRESSAD3 = Trim(oldZADRESS0.ADRESSAD3)
txtUpdate_ADRESSVIL = Trim(oldZADRESS0.ADRESSVIL)
txtUpdate_ADRESSCOP = Trim(oldZADRESS0.ADRESSCOP)

cboUpdate_ADRESSPAY.Clear
rsYBIATAB0_cboK2 "SAB", "CLIENAPAY", cboUpdate_ADRESSPAY
cboUpdate_ADRESSPAY.AddItem "__ " & oldZADRESS0.ADRESSPAY
cboUpdate_ADRESSPAY.ListIndex = 0

'cboUpdate_ADRESSPAY.AddItem "   " & cboUpdate_ADRESSPAY
cbo_Scan oldZCLIENA0.CLIENANAT, cboUpdate_CLIENANAT
cbo_Scan oldZCLIENA0.CLIENARSD, cboUpdate_CLIENARSD

If optUpdate_Add_PP Then
    DTPicker_Set txtUpdate_CLIENADNA, oldZCLIENA0.CLIENADNA + 19000000
    txtUpdate_CLIENBCOM = Trim(oldZCLIENB0.CLIENBCOM)
    txtUpdate_CLIENBCIN = Trim(oldZCLIENB0.CLIENBCIN)
    txtUpdate_CLIENBINS = Trim(oldZCLIENB0.CLIENBINS)
    cbo_Scan oldZCLIENB0.CLIENBLIE, cboUpdate_CLIENBLIE
    txtUpdate_CLIENAFIL = Trim(oldZCLIENA0.CLIENAFIL)
    
    cbo_Scan oldZCLIENB0.CLIENBNAS, cboUpdate_CLIENBNAS
Else
    txtUpdate_CLIENASRN = Trim(oldZCLIENA0.CLIENASRN)
    txtUpdate_CLIENAREG = Trim(oldZCLIENA0.CLIENAREG)
    cbo_Scan oldZCLIENB0.CLIENBTER, cboUpdate_CLIENBTER
End If

End Sub

Public Sub cmdUpdate_Add_Control()
Dim X As String

blnUpdate_Ok = True
newZCLIENA0.CLIENAETB = currentZMNUUTI0.MNUUTIETB
newZCLIENA0.CLIENAAGE = currentZMNUUTI0.MNUUTIAGE
newZCLIENA0.CLIENAETA = Mid$(cboUpdate_CLIENAETA, 1, 4)
X = Trim(txtUpdate_CLIENARA1)
If X = "" Then blnUpdate_Ok = False: Call lstErr_AddItem(lstErr, cmdContext, "? préciser le nom")
newZCLIENA0.CLIENARA1 = X
newZCLIENA0.CLIENARA2 = txtUpdate_CLIENARA2
X = Trim(txtUpdate_CLIENASIG)
If X = "" Then blnUpdate_Ok = False: Call lstErr_AddItem(lstErr, cmdContext, "? préciser le sigle")
newZCLIENA0.CLIENASIG = X

newZCLIENA0.CLIENANAT = Mid$(cboUpdate_CLIENANAT, 1, 3)
newZCLIENA0.CLIENARSD = Mid$(cboUpdate_CLIENARSD, 1, 3)
newZCLIENA0.CLIENARES = "T99" '"R99"
newZCLIENA0.CLIENACAT = "GAR" '"TIE"
newZCLIENA0.CLIENACOT = "NA"
newZCLIENA0.CLIENACHQ = "N"
newZCLIENA0.CLIENAENT = "000"
newZCLIENA0.CLIENAMES = "1"
If txtUpdate_CLIENAFIL.Visible Then
    newZCLIENA0.CLIENAFIL = txtUpdate_CLIENAFIL
Else
    newZCLIENA0.CLIENAFIL = ""
End If

newZCLIENA0.CLIENADOU = "N"
newZCLIENA0.CLIENACOL = "2"
newZCLIENA0.CLIENACRE = DSys - 19000000
'___________________________________________________________________________________________________
newZCLIENB0.CLIENBETB = currentZMNUUTI0.MNUUTIETB

'___________________________________________________________________________________________________
If optUpdate_Add_PP Then
    Call DTPicker_Amj7(txtUpdate_CLIENADNA, newZCLIENA0.CLIENADNA)
    If newZCLIENA0.CLIENADNA > DSys - 19000000 Then blnUpdate_Ok = False: Call lstErr_AddItem(lstErr, cmdContext, "? Date naissance > " & DSys)
    If newZCLIENA0.CLIENADNA < 0 Then blnUpdate_Ok = False: Call lstErr_AddItem(lstErr, cmdContext, "? Date naissance < 1900.01.01 ")
    newZCLIENA0.CLIENAECO = "C03"
    newZCLIENB0.CLIENBNAS = Mid$(cboUpdate_CLIENBNAS, 1, 3)
    
    newZCLIENB0.CLIENBINS = txtUpdate_CLIENBINS
    newZCLIENB0.CLIENBCOM = txtUpdate_CLIENBCOM
    newZCLIENB0.CLIENBCIN = txtUpdate_CLIENBCIN
    newZCLIENB0.CLIENBLIE = Mid$(cboUpdate_CLIENBLIE, 1, 1)
    newZCLIENB0.CLIENBTER = newZCLIENB0.CLIENBLIE
Else
    newZCLIENA0.CLIENAECO = "C01"
    X = Trim(txtUpdate_CLIENASRN)
    If X = "" Then blnUpdate_Ok = False: Call lstErr_AddItem(lstErr, cmdContext, "? préciser le code SIREN")
    newZCLIENA0.CLIENASRN = X
    newZCLIENA0.CLIENAREG = Trim(txtUpdate_CLIENAREG)
    newZCLIENB0.CLIENBTER = Mid$(cboUpdate_CLIENBTER, 1, 1)
End If

'___________________________________________________________________________________________________

newZADRESS0.ADRESSETA = currentZMNUUTI0.MNUUTIETB
newZADRESS0.ADRESSTYP = "1"
If chkUpdate_ADRESSAD1 = "1" Then
    newZADRESS0.ADRESSAD1 = txtUpdate_ADRESSAD1
    newZADRESS0.ADRESSVIL = txtUpdate_ADRESSVIL
Else
    X = Trim(txtUpdate_ADRESSAD1)
    If X = "" Then blnUpdate_Ok = False: Call lstErr_AddItem(lstErr, cmdContext, "? préciser la première ligen adresse")
    newZADRESS0.ADRESSAD1 = X
    X = Trim(txtUpdate_ADRESSVIL)
    If X = "" Then blnUpdate_Ok = False: Call lstErr_AddItem(lstErr, cmdContext, "? préciser la ville")
    newZADRESS0.ADRESSVIL = X
End If

newZADRESS0.ADRESSAD2 = txtUpdate_ADRESSAD2
newZADRESS0.ADRESSAD3 = txtUpdate_ADRESSAD3
newZADRESS0.ADRESSCOP = txtUpdate_ADRESSCOP
newZADRESS0.ADRESSPAY = Mid$(cboUpdate_ADRESSPAY, 4)

End Sub

Public Function cmdUpdate_Add_Old()
Dim V, X As String, xSQL As String
Dim Nb As Long
Dim mMsgBox As String
Dim mCLIGRPREG_99 As String, mCLIGRPREG_9X As String
Dim wCLIGRPREL As String

On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdUpdate_Add_Old"
'-------------------------------------------------------
If Trim(selZCLIENA0.CLIENASRN) = "" Then
    If Trim(oldZCLIENA0.CLIENASRN) <> "" Then
        MsgBox "Une société ne peut être mandataire d'une personne physique", vbCritical, mMsgBox
        Exit Function
    End If
End If

mCLIGRPREG_99 = Mid$(mUpdate_Node_Key, 4, 7) 'oldZCLIENA0.CLIENACLI
mCLIGRPREG_9X = "9X" & Mid$(mUpdate_Node_Key, 6, 5) 'oldZCLIENA0.CLIENACLI
rsZCLIGRP0_Init newZCLIGRP0
newZCLIGRP0.CLIGRPETB = currentZMNUUTI0.MNUUTIETB
newZCLIGRP0.CLIGRPCLI = Mid$(mSelect_Node_Key, 4, 7)   'selZCLIENA0.CLIENACLI
newZCLIGRP0.CLIGRPREL = Mid$(cboUpdate_Add, 1, 3)
newZCLIGRP0.CLIGRPREG = mCLIGRPREG_99
mMsgBox = newZCLIGRP0.CLIGRPCLI & " - " & newZCLIGRP0.CLIGRPREL & " - " & newZCLIGRP0.CLIGRPREG

Set rsAdo = Nothing

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIGRP0 " _
    & " where CLIGRPCLI = '" & newZCLIGRP0.CLIGRPCLI & "'" _
    & " and CLIGRPREG = '" & mCLIGRPREG_99 & "'" _
    & " and CLIGRPETB = " & newZCLIGRP0.CLIGRPETB
Set rsAdo = cnAdo.Execute(xSQL)

If Not rsAdo.EOF Then
    newZCLIGRP0.CLIGRPREG = mCLIGRPREG_9X
    wCLIGRPREL = rsAdo("CLIGRPREL")
    If newZCLIGRP0.CLIGRPREL = wCLIGRPREL Then
        MsgBox "Il existe déjà ce lien entre ces deux entités ", vbCritical, mMsgBox
        Exit Function
    End If
    If newZCLIGRP0.CLIGRPREL = "MAN" And wCLIGRPREL = "DIR" Then
        MsgBox "Lien MAN interdit s'il existe déjà DIR", vbCritical, mMsgBox
        Exit Function
    End If
    If newZCLIGRP0.CLIGRPREL = "DIR" And wCLIGRPREL = "MAN" Then
        MsgBox "Lien DIR interdit s'il existe déjà MAN", vbCritical, mMsgBox
        Exit Function
    End If
End If

Set rsAdo = Nothing

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIGRP0 " _
    & " where CLIGRPCLI = '" & newZCLIGRP0.CLIGRPCLI & "'" _
    & " and CLIGRPREG = '" & mCLIGRPREG_9X & "'" _
    & " and CLIGRPETB = " & newZCLIGRP0.CLIGRPETB
Set rsAdo = cnAdo.Execute(xSQL)

If Not rsAdo.EOF Then
        wCLIGRPREL = rsAdo("CLIGRPREL")
    If newZCLIGRP0.CLIGRPREL = wCLIGRPREL Then
        MsgBox "Il existe déjà ce lien entre ces deux entités ", vbCritical, mMsgBox & " " & mCLIGRPREG_9X
        Exit Function
    End If
    If newZCLIGRP0.CLIGRPREL = "MAN" And wCLIGRPREL = "DIR" Then
        MsgBox "Lien MAN interdit s'il existe déjà DIR", vbCritical, mMsgBox
        Exit Function
    End If
    If newZCLIGRP0.CLIGRPREL = "DIR" And wCLIGRPREL = "MAN" Then
        MsgBox "Lien DIR interdit s'il existe déjà MAN", vbCritical, mMsgBox
        Exit Function
    End If
End If

X = MsgBox("Voulez-vous créér un nouveau lien ?", vbQuestion + vbYesNo + vbDefaultButton2, mMsgBox)
If X = vbNo Then Exit Function


'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cmdUpdate_Transaction
If Not IsNull(V) Then GoTo Error_MsgBox

Call lstErr_AddItem(lstErr, cmdContext, "Ajout lien : " & mMsgBox): DoEvents
'________________________________________________________________________________
V = sqlYUPDLOG0_Init(meYUPDLOG0)
If Not IsNull(V) Then GoTo Error_MsgBox
meYUPDLOG0.UPDLOGUSR = usrName_UCase
meYUPDLOG0.UPDLOGAPP = "SAB_CLIENT"
meYUPDLOG0.UPDLOGFCT = "Add_Old"
meYUPDLOG0.UPDLOGTXT = newZCLIGRP0.CLIGRPCLI & newZCLIGRP0.CLIGRPREL & newZCLIGRP0.CLIGRPREG
V = sqlYUPDLOG0_Insert(meYUPDLOG0)
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

V = sqlZCLIGRP0_Insert(newZCLIGRP0)
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
    
    cmdUpdate_Add_Old = V
    '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function

Public Function cmdUpdate_Delete_Old()
Dim V, X As String, xSQL As String
Dim Nb As Long
Dim mMsgBox As String

On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdUpdate_Add_Old"
'-------------------------------------------------------

rsZCLIGRP0_Init newZCLIGRP0
newZCLIGRP0.CLIGRPETB = currentZMNUUTI0.MNUUTIETB
newZCLIGRP0.CLIGRPCLI = Mid$(mSelect_Node_Key, 4, 7)  ''selZCLIENA0.CLIENACLI
newZCLIGRP0.CLIGRPREL = Mid$(mSelect_Node_Key, 11, 3)
newZCLIGRP0.CLIGRPREG = Mid$(mSelect_Node_Key, 14, 7)
mMsgBox = newZCLIGRP0.CLIGRPCLI & " " & newZCLIGRP0.CLIGRPREL & " " & newZCLIGRP0.CLIGRPREG

Set rsAdo = Nothing

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIGRP0 " _
    & " where CLIGRPCLI = '" & newZCLIGRP0.CLIGRPCLI & "'" _
    & " and CLIGRPREL = '" & newZCLIGRP0.CLIGRPREL & "'" _
    & " and CLIGRPREG like '9%" & Mid$(newZCLIGRP0.CLIGRPREG, 3, 5) & "'" _
    & " and CLIGRPETB = " & newZCLIGRP0.CLIGRPETB
Set rsAdo = cnAdo.Execute(xSQL)

If rsAdo.EOF Then
    MsgBox "Il n'existe pas de lien entre ces deux entités ", vbCritical, mMsgBox
    Exit Function
Else
    newZCLIGRP0.CLIGRPREG = rsAdo("CLIGRPREG")
    X = MsgBox("Voulez-vous supprimer ce lien ?", vbQuestion + vbYesNo + vbDefaultButton2, mMsgBox)
    If X = vbNo Then Exit Function
End If


'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cmdUpdate_Transaction
If Not IsNull(V) Then GoTo Error_MsgBox

Call lstErr_AddItem(lstErr, cmdContext, "Suppression du lien : " & mMsgBox): DoEvents
'________________________________________________________________________________
V = sqlYUPDLOG0_Init(meYUPDLOG0)
If Not IsNull(V) Then GoTo Error_MsgBox
meYUPDLOG0.UPDLOGUSR = usrName_UCase
meYUPDLOG0.UPDLOGAPP = "SAB_CLIENT"
meYUPDLOG0.UPDLOGFCT = "Delete_Old"
meYUPDLOG0.UPDLOGTXT = newZCLIGRP0.CLIGRPCLI & newZCLIGRP0.CLIGRPREL & newZCLIGRP0.CLIGRPREG
V = sqlYUPDLOG0_Insert(meYUPDLOG0)
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

V = sqlZCLIGRP0_Delete(newZCLIGRP0)
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
    
    cmdUpdate_Delete_Old = V
    '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function


Public Function cmdUpdate_Add_New()
Dim V, X As String, xSQL As String
Dim Nb As Long
Dim mMsgBox As String
Dim mCLIENACLI As String

On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdUpdate_Add_New"
'-------------------------------------------------------

Set rsAdo = Nothing
mCLIENACLI = "9900000"
xSQL = "select CLIENACLI from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
    & " where CLIENACLI > '9900000' order by CLIENACLI"
Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
    mCLIENACLI = rsAdo("CLIENACLI")
    rsAdo.MoveNext
Loop


rsZCLIGRP0_Init newZCLIGRP0
newZCLIGRP0.CLIGRPETB = currentZMNUUTI0.MNUUTIETB
newZCLIGRP0.CLIGRPCLI = Mid$(mSelect_Node_Key, 4, 7)   ''selZCLIENA0.CLIENACLI
newZCLIGRP0.CLIGRPREL = Mid$(cboUpdate_Add, 1, 3)
newZCLIGRP0.CLIGRPREG = Format$(Val(mCLIENACLI) + 1, "0000000")
mMsgBox = newZCLIGRP0.CLIGRPCLI & " " & newZCLIGRP0.CLIGRPREL & " " & newZCLIGRP0.CLIGRPREG

X = MsgBox("Voulez-vous créér un nouveau TIERS + lien ?", vbQuestion + vbYesNo + vbDefaultButton2, mMsgBox)
If X = vbNo Then Exit Function


'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cmdUpdate_Transaction
If Not IsNull(V) Then GoTo Error_MsgBox

Call lstErr_AddItem(lstErr, cmdContext, "Ajout lien : " & mMsgBox): DoEvents
'________________________________________________________________________________
V = sqlYUPDLOG0_Init(meYUPDLOG0)
If Not IsNull(V) Then GoTo Error_MsgBox
meYUPDLOG0.UPDLOGUSR = usrName_UCase
meYUPDLOG0.UPDLOGAPP = "SAB_CLIENT"
meYUPDLOG0.UPDLOGFCT = "Add_New"
meYUPDLOG0.UPDLOGTXT = newZCLIGRP0.CLIGRPCLI & newZCLIGRP0.CLIGRPREL & newZCLIGRP0.CLIGRPREG
V = sqlYUPDLOG0_Insert(meYUPDLOG0)
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

V = sqlZCLIGRP0_Insert(newZCLIGRP0)
If Not IsNull(V) Then GoTo Error_MsgBox

newZCLIENA0.CLIENACLI = newZCLIGRP0.CLIGRPREG
V = sqlZCLIENA0_Insert(newZCLIENA0)
If Not IsNull(V) Then GoTo Error_MsgBox


newZCLIENB0.CLIENBCLI = newZCLIGRP0.CLIGRPREG
V = sqlZCLIENB0_Insert(newZCLIENB0)
If Not IsNull(V) Then GoTo Error_MsgBox

newZADRESS0.ADRESSNUM = " " & newZCLIGRP0.CLIGRPREG
V = sqlZADRESS0_Insert(newZADRESS0)
''V = "TEST => Rollback"
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
    
    cmdUpdate_Add_New = V
    '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function

Public Sub cmdSelect_SQL_XG()
On Error GoTo Error_Handler
Dim Nb As Long, NbGroupes As Long, NbClients As Long, NbComptes As Long
Dim wFile As String, xSQL As String
Dim xWhere As String, I As Long
Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim X As String
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CLIENT_Export ........"): DoEvents

wFile = "C:\temp\" & DSys & "_GroupesClientsComptes.xlsx"
X = MsgBox("export du fichier : " & wFile & " ?", vbYesNo, "Export des Groupes > Clients > Comptes")
If X <> vbYes Then Exit Sub
'Kill wFile
Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "GrpCliCpt"
'__________________________________________________________________________________

Nb = 1
wsExcel.Cells(Nb, 1) = "Code": wsExcel.Columns(1).ColumnWidth = 6
wsExcel.Cells(Nb, 2) = "Groupe": wsExcel.Columns(2).ColumnWidth = 8
wsExcel.Cells(Nb, 3) = "Libellé Groupe": wsExcel.Columns(3).ColumnWidth = 41
wsExcel.Cells(Nb, 4) = "Relation": wsExcel.Columns(4).ColumnWidth = 6
wsExcel.Cells(Nb, 5) = "% part": wsExcel.Columns(5).ColumnWidth = 7
wsExcel.Cells(Nb, 6) = "Gest": wsExcel.Columns(6).ColumnWidth = 4
wsExcel.Cells(Nb, 7) = "Client": wsExcel.Columns(7).ColumnWidth = 8
wsExcel.Cells(Nb, 8) = "Libellé Client": wsExcel.Columns(8).ColumnWidth = 50
wsExcel.Cells(Nb, 9) = "Compte": wsExcel.Columns(9).ColumnWidth = 20
wsExcel.Cells(Nb, 10) = "Libellé Compte": wsExcel.Columns(10).ColumnWidth = 41
wsExcel.Cells(Nb, 11) = "Solde": wsExcel.Columns(11).ColumnWidth = 17: wsExcel.Columns(11).NumberFormat = "### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(Nb, 12) = "Devise": wsExcel.Columns(12).ColumnWidth = 6
wsExcel.Cells(Nb, 13) = "Statut": wsExcel.Columns(13).ColumnWidth = 5
wsExcel.Cells(Nb, 14) = "Produit": wsExcel.Columns(14).ColumnWidth = 7
wsExcel.Cells(Nb, 15) = "Dernier mvt": wsExcel.Columns(15).ColumnWidth = 11: wsExcel.Columns(15).NumberFormat = "dd/mm/yyyy"
wsExcel.Columns(15).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Cells(Nb, 16) = "Titulaire P": wsExcel.Columns(16).ColumnWidth = 8: wsExcel.Columns(16).NumberFormat = "#######"


xSQL = "select count(*) as Tally from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI < '0010000'"
Set rsSab = cnsab.Execute(xSQL)
NbGroupes = rsSab("Tally")
ReDim arrZCLIENA0(NbGroupes + 10)

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI < '0010000' order by CLIENACLI"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Nb = Nb + 1
    wsExcel.Cells(Nb, 1) = 1
    arrZCLIENA0(Nb).CLIENACLI = rsSab("CLIENACLI")
    wsExcel.Cells(Nb, 2) = arrZCLIENA0(Nb).CLIENACLI
    arrZCLIENA0(Nb).CLIENARA1 = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
    wsExcel.Cells(Nb, 3) = arrZCLIENA0(Nb).CLIENARA1
    rsSab.MoveNext
Loop
Set rsSab = Nothing
Call lstErr_AddItem(lstErr, cmdContext, "- Groupes : " & Nb): DoEvents

'__________________________________________________________________________

xSQL = "select count(*) as Tally from " & paramIBM_Library_SAB & ".ZCLIGRP0" _
     & " where CLIGRPREG < '0010000'"
Set rsSab = cnsab.Execute(xSQL)
NbClients = rsSab("Tally")
ReDim arrZCLIGRP0(NbClients + 10)
NbClients = 0

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIGRP0 ,  " & paramIBM_Library_SAB & ".ZCLIENA0" _
     & " where CLIGRPREG < '0010000' and CLIGRPCLI = CLIENACLI order by CLIGRPREG"

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    NbClients = NbClients + 1
    V = rsZCLIGRP0_GetBuffer(rsSab, xZCLIGRP0)
    arrZCLIGRP0(NbClients) = xZCLIGRP0
    arrZCLIGRP0(NbClients).CLIGRPCLI_RA1 = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
    
    arrZCLIGRP0(NbClients).CLIGRPCLI_RES = rsSab("CLIENARES")
    For I = 1 To NbGroupes
        If arrZCLIGRP0(NbClients).CLIGRPREG = arrZCLIENA0(I).CLIENACLI Then
            arrZCLIGRP0(NbClients).CLIGRPREG_RA1 = arrZCLIENA0(I).CLIENARA1
            Exit For
        End If
    Next I
    Nb = Nb + 1
    wsExcel.Cells(Nb, 1) = 2
    wsExcel.Cells(Nb, 2) = arrZCLIGRP0(NbClients).CLIGRPREG
    wsExcel.Cells(Nb, 3) = arrZCLIGRP0(NbClients).CLIGRPREG_RA1
    wsExcel.Cells(Nb, 4) = arrZCLIGRP0(NbClients).CLIGRPREL
    wsExcel.Cells(Nb, 5) = arrZCLIGRP0(NbClients).CLIGRPTAU
    wsExcel.Cells(Nb, 6) = arrZCLIGRP0(NbClients).CLIGRPCLI_RES
    wsExcel.Cells(Nb, 7) = arrZCLIGRP0(NbClients).CLIGRPCLI
    wsExcel.Cells(Nb, 8) = arrZCLIGRP0(NbClients).CLIGRPCLI_RA1
    rsSab.MoveNext
Loop
Set rsSab = Nothing
Call lstErr_AddItem(lstErr, cmdContext, "- Clients: " & NbClients): DoEvents
'____________________________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "- SAB_CLIENT_Export  Comptes: "): DoEvents
NbComptes = 0
For I = 1 To NbClients
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "- Groupe-Client : " & arrZCLIGRP0(I).CLIGRPREG & " - " & arrZCLIGRP0(I).CLIGRPCLI & "(" & NbComptes & ")"): DoEvents
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0" _
         & " where CLIENACLI = '" & arrZCLIGRP0(I).CLIGRPCLI & "' and COMPTEFON <> '4' order by COMPTECOM"
    
    Set rsSab = cnsab.Execute(xSQL)
    
    Do While Not rsSab.EOF
        Nb = Nb + 1
        NbComptes = NbComptes + 1
        wsExcel.Cells(Nb, 1) = 3
        wsExcel.Cells(Nb, 2) = arrZCLIGRP0(I).CLIGRPREG
        wsExcel.Cells(Nb, 3) = arrZCLIGRP0(I).CLIGRPREG_RA1
        wsExcel.Cells(Nb, 4) = arrZCLIGRP0(I).CLIGRPREL
        wsExcel.Cells(Nb, 5) = arrZCLIGRP0(I).CLIGRPTAU
        wsExcel.Cells(Nb, 6) = arrZCLIGRP0(I).CLIGRPCLI_RES
        wsExcel.Cells(Nb, 7) = arrZCLIGRP0(I).CLIGRPCLI
        wsExcel.Cells(Nb, 8) = arrZCLIGRP0(I).CLIGRPCLI_RA1
        wsExcel.Cells(Nb, 9) = rsSab("COMPTECOM")
        wsExcel.Cells(Nb, 10) = rsSab("COMPTEINT")
        wsExcel.Cells(Nb, 11) = -CCur(rsSab("SOLDECEN")) / 1000
        wsExcel.Cells(Nb, 12) = rsSab("COMPTEDEV")
        wsExcel.Cells(Nb, 13) = Val(rsSab("COMPTEFON"))
        wsExcel.Cells(Nb, 14) = rsSab("PLANCOPRO")
        X = rsSab("SOLDEDMO") + 19000000
        If X <> 19000000 Then wsExcel.Cells(Nb, 15) = dateAMJ10(X) 'Mid$(X, 7, 2) & "/" & Mid$(X, 5, 2) & "/" & Mid$(X, 1, 4)
        wsExcel.Cells(Nb, 16) = Val(rsSab("TITULACLI"))

        rsSab.MoveNext
    Loop
Next I
Set rsSab = Nothing
Call lstErr_AddItem(lstErr, cmdContext, "- SAB_CLIENT_Export  Comptes : " & NbComptes): DoEvents
'____________________________________________________________________________________

wbExcel.SaveAs wFile

wbExcel.Close
appExcel.Quit

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing

Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CLIENT_Export terminé"): DoEvents
'_____________________________
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CLIENT_Export terminé"): DoEvents

End Sub

Public Function cmdUpdate_Modification()
Dim V, X As String, xSQL As String
Dim Nb As Long
Dim mMsgBox As String
Dim mCLIENACLI As String

On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdUpdate_Modification"
cmdUpdate_Modification = "cmdUpdate_Modification"
'-------------------------------------------------------

Set rsAdo = Nothing
mMsgBox = oldZCLIENA0.CLIENACLI

X = MsgBox("Voulez-vous modifier ce TIERS  ?", vbQuestion + vbYesNo + vbDefaultButton2, mMsgBox)
If X = vbNo Then Exit Function


'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cmdUpdate_Transaction
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
V = sqlYUPDLOG0_Init(meYUPDLOG0)
If Not IsNull(V) Then GoTo Error_MsgBox
meYUPDLOG0.UPDLOGUSR = usrName_UCase
meYUPDLOG0.UPDLOGAPP = "SAB_CLIENT"
meYUPDLOG0.UPDLOGFCT = "Modification"
meYUPDLOG0.UPDLOGTXT = oldZCLIENA0.CLIENACLI
V = sqlYUPDLOG0_Insert(meYUPDLOG0)
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

V = sqlZCLIENA0_Update(newZCLIENA0, oldZCLIENA0)
If Not IsNull(V) Then GoTo Error_MsgBox


V = sqlZCLIENB0_Update(newZCLIENB0, oldZCLIENB0)
If Not IsNull(V) Then GoTo Error_MsgBox

V = sqlZADRESS0_Update(newZADRESS0, oldZADRESS0)
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
    
    cmdUpdate_Modification = V
    '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function



Public Sub tvwSelect_Display_10(X As String)
tvwSelect_Display_ZCLIGRP0_Reset X
lblSelect_Display X, lblSelect
selZCLIENA0 = xZCLIENA0
fraUpdate_Add.Caption = xZCLIENA0.CLIENACLI & " " & Trim(xZCLIENA0.CLIENARA1)
tvwInverse_Display_ZCLIGRP0_Reset X

fraSelect_Update.Visible = True
optSelect_Add.Enabled = True
optSelect_Modification.Enabled = False
optSelect_Suppress.Enabled = False
optSelect_YKYCDOS0.Enabled = True: optSelect_YKYCDOS0.Value = True: blnSelect_YKYCDOS0 = True
If blnSelect_YKYCDOS0 Then
    optSelect_YKYCDOS0.Value = blnSelect_YKYCDOS0
    optSelect_YKYCDOS0_Click
Else
    optSelect_Add.Value = True
End If
End Sub

Public Sub tvwSelect2_Display_10(X As String)
tvwSelect_Display_ZCLIGRP0_Reset X
lblSelect_Display X, lblSelect
selZCLIENA0 = xZCLIENA0
fraUpdate_Add.Caption = xZCLIENA0.CLIENACLI & " " & Trim(xZCLIENA0.CLIENARA1)
tvwInverse_Display_ZCLIGRP0_Reset X

fraSelect_Update.Visible = True
optSelect_Add.Enabled = False
optSelect_Modification.Value = True
optSelect_Modification.Enabled = True
optSelect_Suppress.Enabled = False
optSelect_YKYCDOS0.Enabled = False
End Sub


Private Sub txtSelect_CLIENARA1_LostFocus()
txt_LostFocus txtSelect_CLIENARA1

End Sub

Private Sub txtSelect_Options_4_CLIENACLI_GotFocus()
txt_GotFocus txtSelect_Options_4_CLIENACLI
'cmdSelect_Clear

End Sub

Private Sub txtSelect_Options_4_CLIENACLI_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtSelect_Options_4_CLIENACLI_LostFocus()
txt_LostFocus txtSelect_Options_4_CLIENACLI

End Sub

Private Sub txtSelect_PLANCOPRO_Change()
cmdSelect_Clear

End Sub

Private Sub txtSelect_PLANCOPRO_GotFocus()
txt_GotFocus txtSelect_PLANCOPRO
End Sub

Private Sub txtSelect_PLANCOPRO_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSelect_PLANCOPRO_LostFocus()
txt_LostFocus txtSelect_PLANCOPRO

End Sub

Private Sub txtUpdate_ADRESSAD1_GotFocus()
txt_GotFocus txtUpdate_ADRESSAD1
End Sub

Private Sub txtUpdate_ADRESSAD1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUpdate_ADRESSAD1_LostFocus()
txt_LostFocus txtUpdate_ADRESSAD1

End Sub

Private Sub txtUpdate_ADRESSAD2_GotFocus()
txt_GotFocus txtUpdate_ADRESSAD2
End Sub

Private Sub txtUpdate_ADRESSAD2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUpdate_ADRESSAD2_LostFocus()
txt_LostFocus txtUpdate_ADRESSAD2

End Sub

Private Sub txtUpdate_ADRESSAD3_GotFocus()
txt_GotFocus txtUpdate_ADRESSAD3

End Sub

Private Sub txtUpdate_ADRESSAD3_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUpdate_ADRESSAD3_LostFocus()
txt_LostFocus txtUpdate_ADRESSAD3
End Sub

Private Sub txtUpdate_ADRESSCOP_GotFocus()
txt_GotFocus txtUpdate_ADRESSCOP

End Sub

Private Sub txtUpdate_ADRESSCOP_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUpdate_ADRESSCOP_LostFocus()
txt_LostFocus txtUpdate_ADRESSCOP
End Sub

Private Sub txtUpdate_ADRESSVIL_GotFocus()
txt_GotFocus txtUpdate_ADRESSVIL
End Sub

Private Sub txtUpdate_ADRESSVIL_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUpdate_ADRESSVIL_LostFocus()
txt_LostFocus txtUpdate_ADRESSVIL
End Sub

Private Sub txtUpdate_CLIENADNA_GotFocus()
DTPicker_GotFocus txtUpdate_CLIENADNA

End Sub

Private Sub txtUpdate_CLIENADNA_LostFocus()
DTPicker_LostFocus txtUpdate_CLIENADNA
End Sub


Private Sub txtUpdate_CLIENAFIL_GotFocus()
txt_GotFocus txtUpdate_CLIENAFIL

End Sub

Private Sub txtUpdate_CLIENAFIL_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUpdate_CLIENAFIL_LostFocus()
txt_LostFocus txtUpdate_CLIENAFIL
End Sub

Private Sub txtUpdate_CLIENARA1_GotFocus()
'cmdSelect_Clear
txt_GotFocus txtUpdate_CLIENARA1
End Sub

Private Sub txtUpdate_CLIENARA1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUpdate_CLIENARA1_LostFocus()
txt_LostFocus txtUpdate_CLIENARA1
If Trim(txtUpdate_CLIENASIG) = "" Then txtUpdate_CLIENASIG = txtUpdate_CLIENARA1
End Sub


Private Sub txtUpdate_CLIENARA2_GotFocus()
txt_GotFocus txtUpdate_CLIENARA2
End Sub

Private Sub txtUpdate_CLIENARA2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUpdate_CLIENARA2_LostFocus()
txt_LostFocus txtUpdate_CLIENARA2
End Sub

Private Sub txtUpdate_CLIENAREG_GotFocus()
txt_GotFocus txtUpdate_CLIENAREG
End Sub

Private Sub txtUpdate_CLIENAREG_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtUpdate_CLIENAREG_LostFocus()
txt_LostFocus txtUpdate_CLIENAREG
End Sub

Private Sub txtUpdate_CLIENASIG_GotFocus()
txt_GotFocus txtUpdate_CLIENASIG
End Sub

Private Sub txtUpdate_CLIENASIG_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUpdate_CLIENASIG_LostFocus()
txt_LostFocus txtUpdate_CLIENASIG
End Sub

Private Sub txtUpdate_CLIENASRN_GotFocus()
txt_GotFocus txtUpdate_CLIENASRN
End Sub

Private Sub txtUpdate_CLIENASRN_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtUpdate_CLIENASRN_LostFocus()
txt_LostFocus txtUpdate_CLIENASRN
End Sub

Private Sub txtUpdate_CLIENBCIN_GotFocus()
txt_GotFocus txtUpdate_CLIENBCIN
End Sub

Private Sub txtUpdate_CLIENBCIN_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUpdate_CLIENBCIN_LostFocus()
txt_LostFocus txtUpdate_CLIENBCIN
End Sub

Private Sub txtUpdate_CLIENBCOM_GotFocus()
txt_GotFocus txtUpdate_CLIENBCOM
End Sub

Private Sub txtUpdate_CLIENBCOM_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUpdate_CLIENBCOM_LostFocus()
txt_LostFocus txtUpdate_CLIENBCOM
End Sub

Private Sub txtUpdate_CLIENBINS_GotFocus()
txt_GotFocus txtUpdate_CLIENBINS
End Sub

Private Sub txtUpdate_CLIENBINS_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUpdate_CLIENBINS_LostFocus()
txt_LostFocus txtUpdate_CLIENBINS
End Sub

Private Sub txtUpdate_Select_GotFocus()
txt_GotFocus txtUpdate_Select
End Sub

Private Sub txtUpdate_Select_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUpdate_Select_LostFocus()
txt_LostFocus txtUpdate_Select
End Sub



Public Sub optSelect_CLIENARES_Init()
If Mid$(oldZCLIENA0.CLIENARES, 1, 1) = "X" Then
    optSelect_CLIENARES.Caption = "activer LAB"
Else
    optSelect_CLIENARES.Caption = "désactiver LAB"
End If
optSelect_CLIENARES.Enabled = arrHab(3)

End Sub


Public Sub fgDetail_1r_Display_ZCLIGRP0(lK As Long)

Dim X As String, xSQL As String, wColor As Long, xRA1 As String
Dim wNiveau As Integer, wCLIGRPREG As String

wCLIGRPREG = arrCLIRGPREG(arrCLIRGPREG_K)
wNiveau = arrCLIRGPREG_Niveau(arrCLIRGPREG_K) + 1

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIGRP0 ," & paramIBM_Library_SAB & ".ZCLIENA0" _
      & " where CLIGRPREG = '" & wCLIGRPREG & "'and CLIGRPCLI > '0010000' and CLIENACLI = CLIGRPCLI"
    
Set rsSab = cnsab.Execute(xSQL)
    
If arrCLIRGPREG_Niveau(arrCLIRGPREG_K) = 0 Then
    If chkSelect_Racine <> "1" Then
        If rsSab.EOF Then
            fgDetail.Rows = fgDetail.Rows - 1
            'fgDetail.Row = fgDetail.Rows - 1
            Exit Sub
        End If
    End If
End If
    
Do While Not rsSab.EOF
    X = LCase$(rsSab("CLIGRPREL")) & " " & rsSab("CLIENACLI")
    If InStr(arrCLIRGPREG_Hierarchie(arrCLIRGPREG_K), X) > 0 Then
        wColor = vbRed
    Else
        arrCLIRGPREG_Nb = arrCLIRGPREG_Nb + 1
        arrCLIRGPREG(arrCLIRGPREG_Nb) = rsSab("CLIENACLI")
        arrCLIRGPREG_Niveau(arrCLIRGPREG_Nb) = wNiveau
        arrCLIRGPREG_Hierarchie(arrCLIRGPREG_Nb) = arrCLIRGPREG_Hierarchie(arrCLIRGPREG_K) & " > " & X
        
        xRA1 = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
        If Mid$(rsSab("CLIENARES"), 1, 1) = "X" Then
            xRA1 = "## " & LCase$(xRA1) & " ##"
            wColor = RGB(90, 90, 90)
        Else
            Select Case wNiveau
                Case 1: wColor = RGB(0, 96, 0)
                Case 2: wColor = RGB(160, 64, 0)
                Case Else: wColor = RGB(128, 0, 255)
            End Select

        End If
        fgDetail.Rows = fgDetail.Rows + 1
        fgDetail.Row = fgDetail.Rows - 1
        fgDetail.Col = 0: fgDetail.Text = arrCLIRGPREG_Hierarchie(arrCLIRGPREG_Nb)
        fgDetail.CellForeColor = wColor
        fgDetail.Col = 1: fgDetail.Text = wNiveau
        fgDetail.CellForeColor = wColor
        
        X = rsSab("CLIGRPREL")
        If X <> arrZBAST12_Arg(arrZBAST12_K) Then
            For arrZBAST12_K = 1 To arrZBAST12_Nb
                If X = arrZBAST12_Arg(arrZBAST12_K) Then Exit For
            Next arrZBAST12_K
        End If
        
        fgDetail.Col = 2: fgDetail.Text = arrZBAST12_lib(arrZBAST12_K)
        fgDetail.CellForeColor = wColor
        If rsSab("CLIGRPTAU") <> 0 Then
            fgDetail.Col = 3: fgDetail.Text = Format(rsSab("CLIGRPTAU"), "##0.00")
            fgDetail.CellForeColor = wColor
        End If
        fgDetail.Col = 4: fgDetail.Text = rsSab("CLIENACLI")
        fgDetail.CellForeColor = wColor
        fgDetail.Col = 5: fgDetail.Text = rsSab("CLIENANAT")
        fgDetail.CellForeColor = wColor
        

        fgDetail.Col = 6: fgDetail.Text = xRA1
        fgDetail.CellForeColor = wColor
    End If
    
    rsSab.MoveNext

Loop

End Sub

Public Sub fgDetail_4_Display_WECHISB0(lCOMPTECOM As String)
Dim K As Integer, curX As Currency
On Error GoTo Error_Handler

oldWECHISB0.ECHISBCOM = ""
oldWECHISB0.ECHISBIDE = 0: oldWECHISB0.ECHISBICR = 0: oldWECHISB0.ECHISBTDC = 0: oldWECHISB0.ECHISBCDM = 0: oldWECHISB0.ECHISBPFD = 0: oldWECHISB0.ECHISBPRE = 0:

For K = 1 To arrWECHISB0_nb
    If arrWECHISB0(K).ECHISBCOM = lCOMPTECOM Then
        oldWECHISB0 = arrWECHISB0(K)
        
        curX = oldWECHISB0.ECHISBIDE
        fgDetail.Col = 18
        If oldWECHISB0.ECHISBIDE_AUT <> "" Then fgDetail.CellBackColor = RGB(255, 140, 0)
        If curX <> 0 Then
            fgDetail.Text = Format$(curX, "### ##0.00")
            fgDetail.CellForeColor = IIf(curX < 0, vbRed, vbBlue)
        End If
   
        curX = oldWECHISB0.ECHISBICR
        fgDetail.Col = 19
        If oldWECHISB0.ECHISBICR_AUT <> "" Then fgDetail.CellBackColor = RGB(255, 140, 0)
        If curX <> 0 Then
            fgDetail.Text = Format$(curX, "### ##0.00")
            fgDetail.CellForeColor = IIf(curX < 0, vbRed, vbBlue)
        End If
    
        curX = oldWECHISB0.ECHISBTDC
        fgDetail.Col = 20
        If oldWECHISB0.ECHISBTDC_AUT <> "" Then fgDetail.CellBackColor = RGB(255, 140, 0)
        If curX <> 0 Then
            fgDetail.Text = Format$(curX, "### ##0.00")
            fgDetail.CellForeColor = IIf(curX < 0, vbRed, vbBlue)
       End If
    
        curX = oldWECHISB0.ECHISBCDM
        fgDetail.Col = 21
        If oldWECHISB0.ECHISBCDM_AUT <> "" Then fgDetail.CellBackColor = RGB(255, 140, 0)
        If curX <> 0 Then
            fgDetail.Text = Format$(curX, "### ##0.00")
            fgDetail.CellForeColor = IIf(curX < 0, vbRed, vbBlue)
        End If
    
        curX = oldWECHISB0.ECHISBPFD
        fgDetail.Col = 22
        If oldWECHISB0.ECHISBPFD_AUT <> "" Then fgDetail.CellBackColor = RGB(255, 140, 0)
        If curX <> 0 Then
            fgDetail.Text = Format$(curX, "### ##0.00")
            fgDetail.CellForeColor = IIf(curX < 0, vbRed, vbBlue)
        End If
    
        curX = oldWECHISB0.ECHISBPRE
        fgDetail.Col = 23
        If oldWECHISB0.ECHISBPRE_AUT <> "" Then fgDetail.CellBackColor = RGB(255, 140, 0)
        If curX <> 0 Then
            fgDetail.Text = Format$(curX, "### ##0.00")
            fgDetail.CellForeColor = IIf(curX < 0, vbRed, vbBlue)
        End If
        Exit For
    End If

Next K

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub cmdSelect_SQL_Xgsop_Cumul()
Dim K As Integer, wColor As Long, Nb_T As Integer, blnY6 As Boolean
Dim rsAdo As New ADODB.Recordset
Dim xSQL As String
Dim blnClos As Boolean
If mXgsop.CLIENACLI = "" Then Exit Sub

'DR 07/04/2014
Dim debutMois As String
Dim finMois As String

mXls2_Row = mXls2_Row + 1
wsExcel.Cells(mXls2_Row, 3) = mXgsop.CLIENACLI
wsExcel.Cells(mXls2_Row, 14) = mXgsop.CLIENARES
wsExcel.Cells(mXls2_Row, 15) = mXgsop.CLIENARA1
wsExcel.Cells(mXls2_Row, 16) = mXgsop.PLANCOPRO
wsExcel.Cells(mXls2_Row, 18) = mXgsop.CLIENANAT
wsExcel.Cells(mXls2_Row, 19) = mXgsop.CLIENARSD

Select Case mXgsop.CLIENACOL
    Case 0:
    Case 1: wsExcel.Cells(mXls2_Row, 4) = "Collectif"
    Case 2: wsExcel.Cells(mXls2_Row, 4) = "X"
    Case Else: wsExcel.Cells(mXls2_Row, 4) = mXgsop.CLIENACOL
End Select

wsExcel.Cells(mXls2_Row, 5) = mXgsop.CLIENAETA
wsExcel.Cells(mXls2_Row, 6) = mXgsop.CLIENACAT
        'wsExcel.Cells(mXls2_Row, 5) = xZCLIENA0.CLIENACLI
Select Case Trim(mXgsop.CLIENACAT)
    Case "PAR", "PER", "PRR": mXgsop.X = 1
    Case "STE", "EI": mXgsop.X = 2
    Case "BPR", "BQE", "BQG": mXgsop.X = 3
    Case "GAR":
            Select Case Mid$(mXgsop.CLIENAETA, 1, 1)
                Case "M": mXgsop.X = 1
                Case "S": mXgsop.X = 2
                Case "B": mXgsop.X = 3
                Case Else: mXgsop.X = 4
            End Select
    Case Else: mXgsop.X = 4
End Select
wsExcel.Cells(mXls2_Row, 1) = mXgsop.X

Select Case mXgsop.X
    Case 1: wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_G2
    Case 2: wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_G1
    Case 3: wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_B0
    Case 4: wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_G0
    Case Else: wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_W0
End Select

If mXgsop.CAV_Client <> 0 Then
    mXgsop.y = 1
Else
    If mXgsop.CAV_Tiers = 0 And mXgsop.Tech_Client <> 0 Then
        mXgsop.y = 2
    Else
        If mXgsop.CAV_Tiers + mXgsop.Tech_Tiers <> 0 Then mXgsop.y = 3
    End If
End If

'$JPL 20121005 If mXgsop.CLIENACOL = 2 And mXgsop.y = 2 Then mXgsop.y = 3

If InStr(paramXgsop_NonRéclamés, mXgsop.CLIENACLI) > 0 And mXgsop.CAV_Client <> 0 Then mXgsop.y = 4
If InStr(paramXgsop_HorsGsop, mXgsop.CLIENACLI) > 0 Then mXgsop.y = 5

wsExcel.Cells(mXls2_Row, 2) = mXgsop.y

Select Case mXgsop.y
    Case 1: wsExcel.Cells(mXls2_Row, 2).Interior.Color = mColor_G2
    Case 2: wsExcel.Cells(mXls2_Row, 2).Interior.Color = mColor_Y2
    Case 3: wsExcel.Cells(mXls2_Row, 2).Interior.Color = mColor_Y0
    Case 4: wsExcel.Cells(mXls2_Row, 2).Interior.Color = vbRed 'RGB(255, 128, 0)
    Case 5: wsExcel.Cells(mXls2_Row, 2).Interior.Color = vbMagenta 'RGB(255, 255, 0)
    Case 6: wsExcel.Cells(mXls2_Row, 2).Interior.Color = vbMagenta 'RGB(255, 255, 0)
    'Case Else: wsExcel.Cells(mXls2_Row, 2).Interior.Color = mColor_w0
End Select

If mXgsop.CAV_Client <> 0 Then wsExcel.Cells(mXls2_Row, 7) = mXgsop.CAV_Client
If mXgsop.CAV_Tiers <> 0 Then wsExcel.Cells(mXls2_Row, 8) = mXgsop.CAV_Tiers
If mXgsop.CAV_Clos <> 0 Then wsExcel.Cells(mXls2_Row, 9) = mXgsop.CAV_Clos

If mXgsop.Tech_Client <> 0 Then wsExcel.Cells(mXls2_Row, 10) = mXgsop.Tech_Client
If mXgsop.Tech_Tiers <> 0 Then wsExcel.Cells(mXls2_Row, 11) = mXgsop.Tech_Tiers
If mXgsop.Tech_Clos <> 0 Then wsExcel.Cells(mXls2_Row, 12) = mXgsop.Tech_Clos

blnY6 = False
Nb_T = mXgsop.CAV_Client + mXgsop.CAV_Tiers + mXgsop.CAV_Clos + mXgsop.Tech_Client + mXgsop.Tech_Tiers + mXgsop.Tech_Clos
If Nb_T = 0 Then
    If mXgsop.CLIENACOL = 0 Or mXgsop.CLIENACOL = 1 Then
        If Not retourne_Client_CLOS(mXgsop.CLIENACLI, mXgsop.CLIENARA1 & mXgsop.CLIENARA2) Then
        'If InStr(mXgsop.CLIENARA1 & mXgsop.CLIENARA2, "CLOS") = 0 Then
            'If InStr(mXgsop.CLIENARA1 & mXgsop.CLIENARA2, "CLOTURE") = 0 Then
                blnY6 = True
                wsExcel.Cells(mXls2_Row, 2) = 6: mXgsop.y = 6
                wsExcel.Cells(mXls2_Row, 2).Interior.Color = vbMagenta
            'End If
        End If
    End If
End If
        
    
    
If mXgsop.CAV_Client + mXgsop.CAV_Tiers + mXgsop.Tech_Client + mXgsop.Tech_Tiers = 0 Then
    If Not blnY6 Then
        wColor = RGB(220, 220, 220)
        If mXgsop.COMPTECLO > 0 Then wsExcel.Cells(mXls2_Row, 13) = mXgsop.COMPTECLO + 19000000
        For K = 1 To mXls2_Col: wsExcel.Cells(mXls2_Row, K).Interior.Color = wColor: Next K
        
        If mXgsop.CLIENACOL = 0 Then wsExcel.Cells(mXls2_Row, 4).Interior.Color = mColor_W0
        If Mid$(mXgsop.CLIENARES, 1, 1) <> "X" Then wsExcel.Cells(mXls2_Row, 14).Interior.Color = mColor_W0
        If Not retourne_Client_CLOS(mXgsop.CLIENACLI, mXgsop.CLIENARA1 & mXgsop.CLIENARA2) Then
        'If InStr(mXgsop.CLIENARA1 & mXgsop.CLIENARA2, "CLOS") = 0 Then
            'If InStr(mXgsop.CLIENARA1 & mXgsop.CLIENARA2, "CLOTURE") = 0 Then
                wsExcel.Cells(mXls2_Row, 15).Interior.Color = mColor_W0
            'End If
        End If
    End If
End If
If mXgsop.CAV_Client > 0 And mXgsop.CLIENACOL = 2 Then wsExcel.Cells(mXls2_Row, 4).Interior.Color = vbMagenta

wsExcel.Cells(mXls2_Row, 7).Interior.Color = mColor_G2
wsExcel.Cells(mXls2_Row, 8).Interior.Color = mColor_G0
wsExcel.Cells(mXls2_Row, 10).Interior.Color = mColor_Y2
wsExcel.Cells(mXls2_Row, 11).Interior.Color = mColor_Y0
wsExcel.Cells(mXls2_Row, 9).Interior.Color = RGB(220, 220, 220)
wsExcel.Cells(mXls2_Row, 12).Interior.Color = RGB(220, 220, 220)

'______________________________________________________________________________________

X = "select KYCDOSSTAK from " & paramIBM_Library_SABSPE & ".YKYCDOS0" _
  & " where KYCDOSID = " & mXgsop.CLIENACLI & " and KYCDOSNAT = ' ' and KYCDOSSEQ = 0 and KYCDOSSEQ2 = 0"
Set rsAdo = cnsab.Execute(X)

If rsAdo.EOF Then
    wsExcel.Cells(mXls2_Row, 17) = "néant"
    wsExcel.Cells(mXls2_Row, 17).Interior.Color = mColor_Y2
Else
    If rsAdo("KYCDOSSTAK") = " " Then
        wsExcel.Cells(mXls2_Row, 17) = " "
        wsExcel.Cells(mXls2_Row, 17).Interior.Color = mColor_G0
    Else
        wsExcel.Cells(mXls2_Row, 17) = "X"
        wsExcel.Cells(mXls2_Row, 17).Interior.Color = mColor_W1
    End If
End If
    
'______________________________________________________________________________________
xSQL = "select KYCSTASTAK , KYCSTASTAX , KYCSTASTAY from " & paramIBM_Library_SABSPE & ".YKYCSTA0" _
    & " where KYCSTACLI = '" & mXgsop.CLIENACLI & "'" _
    & " order by KYCSTADSIT desc"
    
Set rsSab_YKYCSTA0 = cnsab.Execute(xSQL)
If rsSab_YKYCSTA0.EOF Then
     mKYCSTASTAK = ""
     mKYCSTASTAX = ""
     mKYCSTASTAY = ""
Else
    mKYCSTASTAK = rsSab_YKYCSTA0("KYCSTASTAK")
    mKYCSTASTAX = rsSab_YKYCSTA0("KYCSTASTAX")
    mKYCSTASTAY = rsSab_YKYCSTA0("KYCSTASTAY")
End If
    

blnClos = False
If mXgsop.y = 0 Then blnClos = True
If mXgsop.y = 5 And wsExcel.Cells(mXls2_Row, 13) <> "" Then blnClos = True: wsExcel.Cells(mXls2_Row, 2) = 0

     
If blnClos Then 'jpl 2014-05-05

    If mKYCSTASTAK <> "9" Then
    'DR 09/04/2014
        'jpl 2014-05-05 If CLng(wsExcel.Cells(mXls2_Row, 13)) > CLng(debutMois) - 1 And CLng(wsExcel.Cells(mXls2_Row, 13)) < CLng(finMois) + 1 Then
            wsExcel.Cells(mXls2_Row, 20) = "Clos"
            wsExcel.Cells(mXls2_Row, 20).Interior.Color = mColor_W1
            wsExcel.Cells(mXls2_Row, 2) = mKYCSTASTAY  'jpl 2013-12-10
        'jpl 2014-05-05 End If
    End If
Else
    Select Case mKYCSTASTAK
        Case "":
                wsExcel.Cells(mXls2_Row, 20) = "Ouv"
                wsExcel.Cells(mXls2_Row, 20).Interior.Color = mColor_G0
       Case "9": newYKYCSTA0.KYCSTASTAK = "2"
                wsExcel.Cells(mXls2_Row, 20) = "Réouv"
                wsExcel.Cells(mXls2_Row, 20).Interior.Color = mColor_Y1
       Case Else
            If mKYCSTASTAX = Trim(wsExcel.Cells(mXls2_Row, 1)) And mKYCSTASTAY = Trim(wsExcel.Cells(mXls2_Row, 2)) Then
            Else
                newYKYCSTA0.KYCSTASTAK = "3"
                wsExcel.Cells(mXls2_Row, 20) = "!! " & mKYCSTASTAX & "-" & mKYCSTASTAY
                wsExcel.Cells(mXls2_Row, 20).Interior.Color = mColor_Y2
                arrK3_Old(mKYCSTASTAX, mKYCSTASTAY) = arrK3_Old(mKYCSTASTAX, mKYCSTASTAY) + 1
                arrK3_New(mXgsop.X, mXgsop.y) = arrK3_New(mXgsop.X, mXgsop.y) + 1
            End If
            
    End Select
End If

'!!!!! maintenance : cmdSelect_SQL_Xgsop_Auto_YKYCSTA0 et cmdSelect_SQL_Xgsop_Cumul
'=========================================================================================

If blnYKYCSTA0_Update Then Call cmdSelect_SQL_Xgsop_Auto_YKYCSTA0

'______________________________________________________________________________________

Call typeXgsop_Init(mXgsop)

End Sub

Public Function DateFinMois(dat As Date) As Date

    '********************************************
    '*   Convertit une date format windows en   *
    '*    une autre date représentant la fin    *
    '*        (28 30 31) du mois courant        *
    '********************************************
    
    Dim MoisCour As Integer         ' Le mois courant
    Dim PremMoisSuivant As Date  ' Le premier du mois suivant

    MoisCour = Month(dat)
    If MoisCour < 12 Then
        PremMoisSuivant = DateSerial(Year(dat), MoisCour + 1, 1)
    Else
        PremMoisSuivant = DateSerial(Year(dat) + 1, 1, 1)
    End If

    DateFinMois = PremMoisSuivant - 1

End Function

Public Sub cmdSelect_SQL_Xgsop_RES(lCLIENARES As String)
Dim xWhere As String, X As String, wFile_RES As String
If lCLIENARES = "" Then
    X = ""
    xWhere = ""
    wFile_RES = wFile_Orig
Else
    X = "_" & lCLIENARES
    xWhere = " and CLIENARES = '" & lCLIENARES & "'"
    wFile_RES = wFile_Orig & "_" & lCLIENARES
End If

If Dir(wFile_RES) <> "" Then msFileSystem.DeleteFile wFile_RES

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "GSOP" & X
    .Subject = "GSOP reporting"
End With
'____________________________________________________________________________________

Call cmdSelect_SQL_Xgsop_Init_2(lCLIENARES)
Call cmdSelect_SQL_Xgsop_Detail(xWhere)
Call cmdSelect_SQL_Xgsop_Detail_6(xWhere)

Call cmdSelect_SQL_Xgsop_Init_1(lCLIENARES)
Call cmdSelect_SQL_Xgsop_Total_Ouv
Call cmdSelect_SQL_Xgsop_Total_Clos
Call cmdSelect_SQL_Xgsop_Total_Réouv
Call cmdSelect_SQL_Xgsop_Total_K3
Call cmdSelect_SQL_Xgsop_Init_3(lCLIENARES)
'____________________________________________________________________________________

Set rsSab = Nothing


wbExcel.SaveAs wFile_RES

wbExcel.Close

appExcel.Quit

End Sub

Public Sub cmdSelect_SQL_Xgsop_Archive()
Dim xSQL As String, X As String, K As Long

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "GSOP" & X
    .Subject = "GSOP Archive"
End With
'____________________________________________________________________________________

Call cmdSelect_SQL_Xgsop_Init_2("")
'______________________________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YKYCSTA0" _
    & " where KYCSTADSIT = " & cboSelect_Options_Xgsop_Archive _
    & " order by KYCSTACLI"
    
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    Call rsYKYCSTA0_GetBuffer(rsSab, xYKYCSTA0)
    Call lstErr_ChangeLastItem(lstErr, cmdContext, xYKYCSTA0.KYCSTACLI): DoEvents
    

    mXls2_Row = mXls2_Row + 1
    wsExcel.Cells(mXls2_Row, 1) = xYKYCSTA0.KYCSTASTAX
    wsExcel.Cells(mXls2_Row, 2) = xYKYCSTA0.KYCSTASTAY
    wsExcel.Cells(mXls2_Row, 3) = xYKYCSTA0.KYCSTACLI
    wsExcel.Cells(mXls2_Row, 4) = xYKYCSTA0.KYCSTAZCOL
    wsExcel.Cells(mXls2_Row, 5) = xYKYCSTA0.KYCSTAZETA
    wsExcel.Cells(mXls2_Row, 6) = xYKYCSTA0.KYCSTAZCAT
    wsExcel.Cells(mXls2_Row, 7) = xYKYCSTA0.KYCSTACAVC
    wsExcel.Cells(mXls2_Row, 8) = xYKYCSTA0.KYCSTACAVT
    wsExcel.Cells(mXls2_Row, 9) = xYKYCSTA0.KYCSTACAVC
    wsExcel.Cells(mXls2_Row, 10) = xYKYCSTA0.KYCSTATECC
    wsExcel.Cells(mXls2_Row, 11) = xYKYCSTA0.KYCSTATECT
    wsExcel.Cells(mXls2_Row, 12) = xYKYCSTA0.KYCSTATECC
    wsExcel.Cells(mXls2_Row, 13) = xYKYCSTA0.KYCSTADCLO
    wsExcel.Cells(mXls2_Row, 14) = xYKYCSTA0.KYCSTAZRES
    wsExcel.Cells(mXls2_Row, 15) = xYKYCSTA0.KYCSTAZRA1
    wsExcel.Cells(mXls2_Row, 16) = xYKYCSTA0.KYCSTAZPCI
    wsExcel.Cells(mXls2_Row, 17) = xYKYCSTA0.KYCSTAYKYC
    wsExcel.Cells(mXls2_Row, 18) = xYKYCSTA0.KYCSTAZNAT
    wsExcel.Cells(mXls2_Row, 19) = xYKYCSTA0.KYCSTAZRSD
    
    Select Case xYKYCSTA0.KYCSTASTAX
        Case 1: wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_G2
        Case 2: wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_G1
        Case 3: wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_B0
        Case 4: wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_G0
        Case Else: wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_W0
    End Select

    wsExcel.Cells(mXls2_Row, 7).Interior.Color = mColor_G2
    wsExcel.Cells(mXls2_Row, 8).Interior.Color = mColor_G0
    wsExcel.Cells(mXls2_Row, 10).Interior.Color = mColor_Y2
    wsExcel.Cells(mXls2_Row, 11).Interior.Color = mColor_Y0
    wsExcel.Cells(mXls2_Row, 9).Interior.Color = RGB(220, 220, 220)
    wsExcel.Cells(mXls2_Row, 12).Interior.Color = RGB(220, 220, 220)
    Select Case xYKYCSTA0.KYCSTASTAY
        Case 0:
            For K = 1 To 20
                wsExcel.Cells(mXls2_Row, K).Interior.Color = RGB(220, 220, 220)
            Next K
        Case 1: wsExcel.Cells(mXls2_Row, 2).Interior.Color = mColor_G2
        Case 2: wsExcel.Cells(mXls2_Row, 2).Interior.Color = mColor_Y2
        Case 3: wsExcel.Cells(mXls2_Row, 2).Interior.Color = mColor_Y0
        Case 4: wsExcel.Cells(mXls2_Row, 2).Interior.Color = vbRed 'RGB(255, 128, 0)
        Case 5: wsExcel.Cells(mXls2_Row, 2).Interior.Color = vbMagenta 'RGB(255, 255, 0)
        Case 6: wsExcel.Cells(mXls2_Row, 2).Interior.Color = vbMagenta 'RGB(255, 255, 0)
    End Select
                
    If xYKYCSTA0.KYCSTAYKYC = "N" Then
        wsExcel.Cells(mXls2_Row, 17) = "néant"
        wsExcel.Cells(mXls2_Row, 17).Interior.Color = mColor_Y2
    Else
        If xYKYCSTA0.KYCSTAYKYC = " " Then
            wsExcel.Cells(mXls2_Row, 17) = " "
            wsExcel.Cells(mXls2_Row, 17).Interior.Color = mColor_G0
        Else
            wsExcel.Cells(mXls2_Row, 17) = "X"
            wsExcel.Cells(mXls2_Row, 17).Interior.Color = mColor_W1
        End If
    End If
    
    Select Case xYKYCSTA0.KYCSTASTAK
        Case "1": wsExcel.Cells(mXls2_Row, 20) = ""
        Case "0": wsExcel.Cells(mXls2_Row, 20) = "Ouv": wsExcel.Cells(mXls2_Row, 20).Interior.Color = mColor_G0
        Case "2": wsExcel.Cells(mXls2_Row, 20) = "Réouv": wsExcel.Cells(mXls2_Row, 20).Interior.Color = mColor_Y1
        Case "3": wsExcel.Cells(mXls2_Row, 20) = "!!": wsExcel.Cells(mXls2_Row, 20).Interior.Color = mColor_Y2
        
        Case "9": wsExcel.Cells(mXls2_Row, 20) = "Clos": wsExcel.Cells(mXls2_Row, 20).Interior.Color = mColor_W1
        
                xSQL = "select * from " & paramIBM_Library_SABSPE & ".YKYCSTA0" _
                    & " where KYCSTACLI = '" & xYKYCSTA0.KYCSTACLI & "'" _
                    & " order by KYCSTADSIT desc "
                Set rsSabX = cnsab.Execute(xSQL)
                Do While Not rsSabX.EOF
                    If rsSabX("KYCSTASTAY") <> 0 Then
                        wsExcel.Cells(mXls2_Row, 2) = rsSabX("KYCSTASTAY")
                        Exit Do
                    End If
                    rsSabX.MoveNext
                Loop
    End Select
    rsSab.MoveNext
Loop
Call cmdSelect_SQL_Xgsop_Init_1("au " & dateImp10(cboSelect_Options_Xgsop_Archive))
Call cmdSelect_SQL_Xgsop_Total_Ouv
Call cmdSelect_SQL_Xgsop_Total_Clos
Call cmdSelect_SQL_Xgsop_Total_Réouv
Call cmdSelect_SQL_Xgsop_Total_K3
'____________________________________________________________________________________

Set rsSab = Nothing


wbExcel.SaveAs wFile_Orig

wbExcel.Close

appExcel.Quit

End Sub


Public Sub cmdSelect_SQL_Xgsop_Init()

cboSelect_Options_Xgsop_CLIENARES.Clear
cboSelect_Options_Xgsop_CLIENARES.AddItem "*"
cboSelect_Options_Xgsop_CLIENARES.AddItem "Archives"
cboSelect_Options_Xgsop_CLIENARES.AddItem "* + R"
cboSelect_Options_KYCgsop_CLIENARES.Clear
cboSelect_Options_KYCgsop_CLIENARES.AddItem "*"
cboSelect_Options_KYCech_CLIENARES.Clear
cboSelect_Options_KYCech_CLIENARES.AddItem ""
cboSelect_Options_KYCech_CLIENARES.AddItem "*"

X = "select distinct CLIENARES from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
  & " order by CLIENARES"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    cboSelect_Options_Xgsop_CLIENARES.AddItem rsSab("CLIENARES")
    cboSelect_Options_KYCgsop_CLIENARES.AddItem rsSab("CLIENARES")
    cboSelect_Options_KYCech_CLIENARES.AddItem rsSab("CLIENARES")
    rsSab.MoveNext
Loop
cboSelect_Options_Xgsop_CLIENARES.ListIndex = 0
cboSelect_Options_KYCgsop_CLIENARES.ListIndex = 0
cboSelect_Options_KYCech_CLIENARES.ListIndex = 0

cboSelect_Options_Xgsop_Archive.Clear
X = "select distinct KYCSTADSIT from " & paramIBM_Library_SABSPE & ".YKYCSTA0 " _
  & " order by KYCSTADSIT desc"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    cboSelect_Options_Xgsop_Archive.AddItem rsSab("KYCSTADSIT")
    rsSab.MoveNext
Loop

Dim K As Integer, K2 As Integer
For K = 1 To 4
    For K2 = 1 To 6
        arrK3_Old(K, K2) = 0
        arrK3_New(K, K2) = 0
        
    Next K2
Next K

arrX_Lib(1) = "1-Particuliers"
arrX_Lib(2) = "2-Pers. morales"
arrX_Lib(3) = "3-Banques"
arrX_Lib(4) = "4-Autres"

arrY_Lib(1) = "1-Clients"
arrY_Lib(2) = "2-Techniques"
arrY_Lib(3) = "3-Tiers"
arrY_Lib(4) = "4-BIA non réclamés"
arrY_Lib(5) = "5-hors GSOP"
arrY_Lib(6) = "6-Racines sans compte"

End Sub

Public Sub lstParam_YKYCDOS0_Load()
Dim X As String, xSQL As String
On Error GoTo Error_Handler

Call lstParam_YKYCDOS0_J_Load
Call lstParam_YKYCDOS0_D_Load

lstParam_YKYCDOS0.Clear
X = mParam_KYCDOSNAT
If mParam_KYCDOSNAT = "=" Then
    X = "*"
Else
    lstParam_YKYCDOS0.AddItem "Ajouter un enregistrement"
End If

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YKYCDOS0" _
     & " where KYCDOSNAT = '" & X & "'" _
     & " order by KYCDOSID , KYCDOSSEQ"

Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
    lstParam_YKYCDOS0.AddItem Trim(rsAdo("KYCDOSID")) & Format((rsAdo("KYCDOSSEQ")), "####") & Format((rsAdo("KYCDOSSEQ2")), "####") & " : " & Trim(rsAdo("KYCDOSDLIB"))
    rsAdo.MoveNext

Loop
lstParam_YKYCDOS0.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : lstParam_YKYCDOS0_Load"


End Sub
Public Sub lstYKYCDOS0_CLIENACAT_Load()
Dim X As String, xSQL As String
On Error GoTo Error_Handler

lstYKYCDOS0_CLIENACAT.Clear
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YKYCDOS0" _
     & " where KYCDOSNAT = '*'" _
     & " order by KYCDOSID , KYCDOSSEQ"

Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
    lstYKYCDOS0_CLIENACAT.AddItem Trim(rsAdo("KYCDOSID")) & Format((rsAdo("KYCDOSSEQ")), "####") & " : " & Trim(rsAdo("KYCDOSDLIB"))
    rsAdo.MoveNext

Loop
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : lstYKYCDOS0_CLIENACAT_Load"


End Sub

Public Sub lstParam_YKYCDOS0_J_Load()
Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler

ReDim arrKYCDOSDLIB_J(999), arrKYCDOSSTAK_J(999)
lstParam_YKYCDOS0_J.Clear
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YKYCDOS0" _
     & " where KYCDOSNAT = 'J'" _
     & " order by KYCDOSSEQ"

Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
    K = rsAdo("KYCDOSSEQ")
    X = Trim(rsAdo("KYCDOSDLIB"))
    lstParam_YKYCDOS0_J.AddItem Format(K, "####") & " : " & X
    arrKYCDOSDLIB_J(K) = X
    arrKYCDOSSTAK_J(K) = rsAdo("KYCDOSSTAK")
    rsAdo.MoveNext

Loop
lstParam_YKYCDOS0_J.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : lstParam_YKYCDOS0_Load"


End Sub



Public Sub lstParam_YKYCDOS0_D_Load()
Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler

ReDim arrKYCDOSDLIB_D(999), arrKYCDOSSTAK_D(999), arrKYCDOSDECH_D(999)

lstParam_YKYCDOS0_D.Clear
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YKYCDOS0" _
     & " where KYCDOSNAT = 'D'" _
     & " order by KYCDOSSEQ"

Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
    K = rsAdo("KYCDOSSEQ")
    X = Trim(rsAdo("KYCDOSDLIB"))
    lstParam_YKYCDOS0_D.AddItem Format(K, "####") & " : " & X
    arrKYCDOSDLIB_D(K) = X
    arrKYCDOSSTAK_D(K) = rsAdo("KYCDOSSTAK")
    arrKYCDOSDECH_D(K) = rsAdo("KYCDOSDECH")
    rsAdo.MoveNext

Loop
lstParam_YKYCDOS0_D.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : lstParam_YKYCDOS0_D_Load"


End Sub


Public Sub lstParam_YKYCDOS0_JD_Load()
Dim X As String, xSQL As String, K1 As Long, K2 As Long
On Error GoTo Error_Handler

ReDim oldParam_J(999), newParam_J(999)
ReDim oldParam_D(999), newParam_D(999)

lstParam_YKYCDOS0_JD.Clear
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YKYCDOS0" _
     & " where KYCDOSNAT = '=' and KYCDOSID = '" & oldYKYCDOS0.KYCDOSID & "'" _
     & " order by KYCDOSSEQ , KYCDOSSEQ2"

Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
    K1 = rsAdo("KYCDOSSEQ") '/ 10000
    K2 = rsAdo("KYCDOSSEQ2") ' - K1 * 10000
    If K2 = 0 Then
        oldParam_J(K1) = K1: newParam_J(K1) = K1
    Else
        oldParam_D(K2) = K1: newParam_D(K2) = K1
    End If
    rsAdo.MoveNext

Loop
lstParam_YKYCDOS0_JD_Display
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : lstParam_YKYCDOS0_D_Load"


End Sub



Public Sub lstParam_YKYCDOS0_4c_Load()
Dim xSQL As String, K As Long, wColor As Long
On Error GoTo Error_Handler

fgParam_YKYCDOS0_4c.Visible = False
fgParam_YKYCDOS0_4c.Rows = 1

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YKYCDOS0" _
     & " where KYCDOSNAT = '=' and KYCDOSID = '" & oldYKYCDOS0.KYCDOSID & "'" _
     & " order by KYCDOSSEQ , KYCDOSSEQ2"

Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
    fgParam_YKYCDOS0_4c.Rows = fgParam_YKYCDOS0_4c.Rows + 1
    fgParam_YKYCDOS0_4c.Row = fgParam_YKYCDOS0_4c.Rows - 1
    fgParam_YKYCDOS0_4c.Col = 0: fgParam_YKYCDOS0_4c.Text = rsAdo("KYCDOSSEQ")
    fgParam_YKYCDOS0_4c.Col = 1: fgParam_YKYCDOS0_4c.Text = rsAdo("KYCDOSSEQ2")
    fgParam_YKYCDOS0_4c.Col = 2:
        If rsAdo("KYCDOSSEQ2") = 0 Then
            wColor = mColor_G2
            fgParam_YKYCDOS0_4c.Text = arrKYCDOSDLIB_J(rsAdo("KYCDOSSEQ"))
        Else
            wColor = mColor_G1
            fgParam_YKYCDOS0_4c.Text = arrKYCDOSDLIB_D(rsAdo("KYCDOSSEQ2"))
        End If
    fgParam_YKYCDOS0_4c.Col = 3: fgParam_YKYCDOS0_4c.Text = rsAdo("KYCDOSSTAK")
    fgParam_YKYCDOS0_4c.Col = 4: fgParam_YKYCDOS0_4c.Text = rsAdo("KYCDOSDLIB")
    If rsAdo("KYCDOSSTAK") = "O" Then
        For K = 0 To 4: fgParam_YKYCDOS0_4c.Col = K: fgParam_YKYCDOS0_4c.CellBackColor = wColor: Next K
    Else
        If rsAdo("KYCDOSSEQ2") = 0 Then
            For K = 0 To 4: fgParam_YKYCDOS0_4c.Col = K: fgParam_YKYCDOS0_4c.CellBackColor = mColor_Y2: Next K
        End If
    End If
    rsAdo.MoveNext

Loop
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : lstParam_YKYCDOS0_D_Load"


End Sub

Public Sub lstParam_YKYCDOS0_JD_Display()

Dim K As Integer
lstParam_YKYCDOS0_JD.Visible = False
lstParam_YKYCDOS0_JD.Clear
For K = 1 To 999
    If newParam_J(K) <> 0 Then
        lstParam_YKYCDOS0_JD.AddItem K & "." & " : " & arrKYCDOSDLIB_J(K)
    End If
Next K
For K = 1 To 999
    If newParam_D(K) <> 0 Then
        lstParam_YKYCDOS0_JD.AddItem newParam_D(K) & "." & K & vbTab & " : " & arrKYCDOSDLIB_D(K)
    End If
Next K

For K = 0 To lstParam_YKYCDOS0_JD.ListCount - 1
    lstParam_YKYCDOS0_JD.Selected(K) = True
Next K

blnYKYCDOS0_JD = True
lstParam_YKYCDOS0_JD.ListIndex = -1
lstParam_YKYCDOS0_JD.Visible = True

End Sub

Public Sub fgYKYCDOS0_Display()

Dim K As Long, K1 As Long, K2 As Long, wKYCDOSSEQ As Long, wKYCDOSSEQ2 As Long, blnOk As Boolean
Dim wColor As Long
Dim I As Integer, X As String
On Error GoTo Error_Handler
fgYKYCDOS0.Visible = False
fgYKYCDOS0_Reset

fgYKYCDOS0.Rows = 1
fgYKYCDOS0.FormatString = fgYKYCDOS0_FormatString
currentAction = "fgYKYCDOS0_Display"

cmdSelect_SQL_YKYCDOS0_Load

'_______________________________________________________________________________________________________


For K = 1 To arrYKYCDOS0_JD_Nb
    fgYKYCDOS0.Rows = fgYKYCDOS0.Rows + 1
    fgYKYCDOS0.Row = fgYKYCDOS0.Rows - 1
    
    K1 = arrYKYCDOS0_JD(K).KYCDOSSEQ
    K2 = arrYKYCDOS0_JD(K).KYCDOSSEQ2
    
    fgYKYCDOS0.Col = 5: fgYKYCDOS0.Text = K1
    fgYKYCDOS0.Col = 6: fgYKYCDOS0.Text = K2
    If K2 = 0 Then
        fgYKYCDOS0.CellFontBold = True: fgYKYCDOS0.CellFontSize = 11
        
        Select Case arrYKYCDOS0(K).KYCDOSSTAK
            Case "=": wColor = &H80FF80 ' vbgreen
            Case " ": wColor = &H80FFFF    'vbYellow
            Case Else: wColor = &HFF80FF: arrYKYCDOS0(0).KYCDOSSTAK = "N"
                    fgYKYCDOS0.Col = 2: fgYKYCDOS0.Text = "obligatoire": fgYKYCDOS0.CellForeColor = vbRed

        End Select

        'fgYKYCDOS0.CellBackColor = wColor
        fgYKYCDOS0.Col = 0: fgYKYCDOS0.Text = arrKYCDOSDLIB_J(K1)
        fgYKYCDOS0.CellFontBold = True: fgYKYCDOS0.CellFontSize = 11

        fgYKYCDOS0.CellBackColor = wColor
    Else
      Select Case arrYKYCDOS0(K).KYCDOSSTAK
        Case "=":
                For I = 0 To 6: fgYKYCDOS0.Col = I: fgYKYCDOS0.CellBackColor = mColor_G1: Next I
        Case "?": arrYKYCDOS0(0).KYCDOSSTAK = "N"
                For I = 0 To 6: fgYKYCDOS0.Col = I: fgYKYCDOS0.CellBackColor = mColor_W0: Next I
                fgYKYCDOS0.Col = 2: fgYKYCDOS0.Text = "manquant": fgYKYCDOS0.CellForeColor = vbRed
        Case "O": arrYKYCDOS0(0).KYCDOSSTAK = "N"
                For I = 0 To 6: fgYKYCDOS0.Col = I: fgYKYCDOS0.CellBackColor = mColor_Y1: Next I
                    fgYKYCDOS0.Col = 2: fgYKYCDOS0.Text = "obligatoire": fgYKYCDOS0.CellForeColor = vbRed
        Case "I":
                For I = 0 To 6: fgYKYCDOS0.Col = I: fgYKYCDOS0.CellBackColor = RGB(220, 220, 220): Next I
                fgYKYCDOS0.Col = 2: fgYKYCDOS0.Text = "doc ignoré": fgYKYCDOS0.CellForeColor = vbRed

        End Select
        
      fgYKYCDOS0.Col = 0: fgYKYCDOS0.Text = arrKYCDOSDLIB_D(K2)
      fgYKYCDOS0.CellFontSize = 8

      fgYKYCDOS0.Col = 1
      If arrYKYCDOS0(K).KYCDOSDAMJ <> 0 Then fgYKYCDOS0.Text = dateImp10_S(arrYKYCDOS0(K).KYCDOSDAMJ)
      
      fgYKYCDOS0.Col = 2
      If arrKYCDOSDECH_D(K2) <> 0 And arrYKYCDOS0(K).KYCDOSSTAK = "=" Then
        If arrYKYCDOS0(K).KYCDOSDECH <> 0 Then fgYKYCDOS0.Text = dateImp10_S(arrYKYCDOS0(K).KYCDOSDECH)
        If arrYKYCDOS0(K).KYCDOSDECH < mKYCDOSDECH_Warn Then
            fgYKYCDOS0.CellForeColor = vbRed
            fgYKYCDOS0.CellBackColor = mColor_Y2
        Else
            fgYKYCDOS0.CellForeColor = vbBlue
        End If
      End If
        fgYKYCDOS0.Col = 3: fgYKYCDOS0.Text = arrYKYCDOS0(K).KYCDOSPJ
        If arrYKYCDOS0(K).KYCDOSPJ <> " " Then fgYKYCDOS0.CellBackColor = vbGreen
        
   End If
   
   fgYKYCDOS0.Col = 4
   If Trim(arrYKYCDOS0(K).KYCDOSDLIB) <> "" Then
        fgYKYCDOS0.Text = arrYKYCDOS0(K).KYCDOSDLIB
   Else
        fgYKYCDOS0.Text = arrYKYCDOS0_JD(K).KYCDOSDLIB
        fgYKYCDOS0.CellFontItalic = True
        fgYKYCDOS0.CellForeColor = &H800080
   End If
   
Next K

 If arrYKYCDOS0(0).KYCDOSSTAK <> currentYKYCDOS0.KYCDOSSTAK Then
    oldYKYCDOS0 = currentYKYCDOS0
    newYKYCDOS0 = arrYKYCDOS0(0)
    newYKYCDOS0.KYCDOSUFCT = "#"
    V = cmdParam_YKYCDOS0_Transaction("Update")
   
 End If
 
fgYKYCDOS0.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Nb : " & fgYKYCDOS0.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction




End Sub
Public Sub fgYKYCCTL_Display()

Dim K As Long, K1 As Long, K2 As Long, wKYCDOSSEQ As Long, wKYCDOSSEQ2 As Long, blnOk As Boolean
Dim wColor As Long
Dim I As Integer, X As String
On Error GoTo Error_Handler
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Top = 480
fgSelect.ZOrder 0
fgSelect.Rows = 1
fgSelect.FormatString = "<Ctl       |<Resp |<Racine  |<Intitulé                               " _
                       & "|<SAB|<Date de délivrance" _
                       & "|<BIA     |<Date de délivrance|<Echéance    |< commentaire                   |<|<|<|<"
currentAction = "fgselect_Display"


'_______________________________________________________________________________________________________


For K = 1 To wKYCCTL_Nb
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
     fgSelect.Col = 0: fgSelect.Text = wKYCCTL(K).STA
     fgSelect.Col = 1: fgSelect.Text = wKYCCTL(K).CLIENARES
     fgSelect.Col = 2: fgSelect.Text = wKYCCTL(K).CLIENACLI
     fgSelect.Col = 3: fgSelect.Text = wKYCCTL(K).CLIENARA
     fgSelect.Col = 4: fgSelect.Text = wKYCCTL(K).CLIENCPIE
     If wKYCCTL(K).CLILIBDA1 > 0 Then fgSelect.Col = 5: fgSelect.Text = dateImp10_S(wKYCCTL(K).CLILIBDA1)
     fgSelect.Col = 6
     Select Case wKYCCTL(K).KYCDOSSEQ2
            Case 0
            Case 20, 21: fgSelect.Text = wKYCCTL(K).KYCDOSSEQ2 & "-CNI"
            Case 22, 220: fgSelect.Text = wKYCCTL(K).KYCDOSSEQ2 & "-PA"
            Case 23: fgSelect.Text = wKYCCTL(K).KYCDOSSEQ2 & "-CR"
            Case Else: fgSelect.Text = wKYCCTL(K).KYCDOSSEQ2
    End Select
     If wKYCCTL(K).KYCDOSDAMJ > 0 Then fgSelect.Col = 7: fgSelect.Text = dateImp10_S(wKYCCTL(K).KYCDOSDAMJ)
     If wKYCCTL(K).KYCDOSDECH > 0 Then fgSelect.Col = 8: fgSelect.Text = dateImp10_S(wKYCCTL(K).KYCDOSDECH)
     fgSelect.Col = 9: fgSelect.Text = " " & Trim(wKYCCTL(K).KYCDOSDLIB)
   
        
        Select Case wKYCCTL(K).STA
            Case "": wColor = vbWhite
            Case "#": wColor = mColor_W0
            Case "? GSOP": wColor = mColor_W1
            Case "? SAB KYC": wColor = mColor_Y2
            Case "! doc #": wColor = mColor_B0
            Case Else: wColor = mColor_W1
        End Select
        If wKYCCTL(K).CLIENARES = "R80" Then
            If wKYCCTL(K).STA <> "" Then wColor = RGB(230, 230, 230)
        End If
        
        For I = 0 To 9: fgSelect.Col = I: fgSelect.CellBackColor = wColor: Next I

   
Next K

fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
 
fgSelect.Visible = True
SSTab1.Tab = 2
Call lstErr_AddItem(lstErr, cmdContext, "Nb : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction




End Sub


Public Sub cmdSelect_SQL_YKYCDOS0_Load()

Dim K As Long, K1 As Long, K2 As Long, wKYCDOSSEQ As Long, wKYCDOSSEQ2 As Long, blnOk As Boolean
Dim I As Integer, X As String
On Error GoTo Error_Handler

Call rsYKYCDOS0_Init(xYKYCDOS0)
xYKYCDOS0.KYCDOSNAT = ""
xYKYCDOS0.KYCDOSID = currentYKYCDOS0.KYCDOSID
'_______________________________________________________________________________________________________

X = "select count(*) from  " & paramIBM_Library_SABSPE & ".YKYCDOS0 " _
    & " where  KYCDOSNAT = '='and KYCDOSID = '" & currentYKYCDOS0.KYCDOSDLIB & "'"
Set rsSabX = cnsab.Execute(X)
If Not rsSabX.EOF Then ReDim arrYKYCDOS0_JD(rsSabX(0) + 1): ReDim arrYKYCDOS0(rsSabX(0) + 1)
Dim arrOK_Nb(999), arrNOK_Nb(999)

arrYKYCDOS0_JD_Nb = 0

X = "select * from  " & paramIBM_Library_SABSPE & ".YKYCDOS0 " _
    & " where  KYCDOSNAT = '='and KYCDOSID = '" & currentYKYCDOS0.KYCDOSDLIB & "'  order by KYCDOSSEQ , KYCDOSSEQ2 "
Set rsSabX = cnsab.Execute(X)
        
Do While Not rsSabX.EOF
    arrYKYCDOS0_JD_Nb = arrYKYCDOS0_JD_Nb + 1
    Call rsYKYCDOS0_GetBuffer(rsSabX, arrYKYCDOS0_JD(arrYKYCDOS0_JD_Nb))
    arrYKYCDOS0(arrYKYCDOS0_JD_Nb) = xYKYCDOS0
    arrYKYCDOS0(arrYKYCDOS0_JD_Nb).KYCDOSSEQ = arrYKYCDOS0_JD(arrYKYCDOS0_JD_Nb).KYCDOSSEQ
    arrYKYCDOS0(arrYKYCDOS0_JD_Nb).KYCDOSSEQ2 = arrYKYCDOS0_JD(arrYKYCDOS0_JD_Nb).KYCDOSSEQ2
    arrYKYCDOS0(arrYKYCDOS0_JD_Nb).KYCDOSSTAK = arrYKYCDOS0_JD(arrYKYCDOS0_JD_Nb).KYCDOSSTAK
    rsSabX.MoveNext
Loop
'_______________________________________________________________________________________________________


'_______________________________________________________________________________________________________


X = "select * from  " & paramIBM_Library_SABSPE & ".YKYCDOS0 " _
    & " where  KYCDOSNAT = ' 'and KYCDOSID = '" & currentYKYCDOS0.KYCDOSID & "'  order by KYCDOSSEQ , KYCDOSSEQ2  "
Set rsSabX = cnsab.Execute(X)
        
Do While Not rsSabX.EOF
    wKYCDOSSEQ = rsSabX("KYCDOSSEQ")
    wKYCDOSSEQ2 = rsSabX("KYCDOSSEQ2")
    blnOk = False
    For K = 0 To arrYKYCDOS0_JD_Nb
        If wKYCDOSSEQ = arrYKYCDOS0_JD(K).KYCDOSSEQ _
        And wKYCDOSSEQ2 = arrYKYCDOS0_JD(K).KYCDOSSEQ2 Then
            Call rsYKYCDOS0_GetBuffer(rsSabX, arrYKYCDOS0(K))
            blnOk = True
            If arrYKYCDOS0(K).KYCDOSSTAK = " " Then   'Or arrYKYCDOS0(K).KYCDOSSTAK = "I" Then
                arrYKYCDOS0(K).KYCDOSSTAK = "="
                arrOK_Nb(wKYCDOSSEQ) = arrOK_Nb(wKYCDOSSEQ) + 1
            Else
                If arrYKYCDOS0(K).KYCDOSSTAK <> "I" Then arrNOK_Nb(wKYCDOSSEQ) = arrNOK_Nb(wKYCDOSSEQ) + 1
            End If
            
            'If arrYKYCDOS0(K).KYCDOSSTAK = " " Then arrYKYCDOS0(K).KYCDOSSTAK = "="
            
            Exit For
        End If
    Next K
    If Not blnOk Then
        Call MsgBox("Client : " & currentYKYCDOS0.KYCDOSID & " - Type  : " & currentYKYCDOS0.KYCDOSDLIB & vbCrLf _
        & "Justificatif = " & wKYCDOSSEQ & " , Document = " & wKYCDOSSEQ2, vbCritical, "Paramétrage erroné")
    End If
    rsSabX.MoveNext
Loop
'_______________________________________________________________________________________________________

For I = 1 To arrYKYCDOS0_JD_Nb
    If arrYKYCDOS0(I).KYCDOSSEQ2 <> 0 Then
        If arrYKYCDOS0(I).KYCDOSSTAK = "O" Then
            arrNOK_Nb(arrYKYCDOS0(I).KYCDOSSEQ) = arrNOK_Nb(arrYKYCDOS0(I).KYCDOSSEQ) + 1
        End If
    Else
    
    End If
Next I
'_______________________________________________________________________________________________________

For I = 1 To arrYKYCDOS0_JD_Nb
    If arrYKYCDOS0(I).KYCDOSSEQ2 = 0 Then
        If arrNOK_Nb(arrYKYCDOS0(I).KYCDOSSEQ) <> 0 Then
            arrYKYCDOS0(I).KYCDOSSTAK = "O"
            
        Else
            If arrOK_Nb(arrYKYCDOS0(I).KYCDOSSEQ) > 0 Then
                arrYKYCDOS0(I).KYCDOSSTAK = "="
            Else
                arrYKYCDOS0(I).KYCDOSSTAK = arrYKYCDOS0_JD(I).KYCDOSSTAK
            End If
        End If
    Else
    
    End If
Next I

'_______________________________________________________________________________________________________
arrYKYCDOS0(0).KYCDOSSTAK = " "


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction




End Sub

Public Sub cmdParam_YKYCDOS0_Print()
On Error GoTo Error_Handler
Dim X As String, K As Long, K1 As Long, K2 As Long
Dim wFilex As String, xSQL As String
Dim rsSab_RES As New ADODB.Recordset

On Error GoTo Error_Handler
'______________________________________________'


 wFile_Orig = Trim("C:\Temp\GSOP KYC param  " & DSys & "_" & time_Hms)

If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile_Orig _
        & vbCrLf & "     =========================", "GSOP KYC param : nom du fichier d'exportation", wFile_Orig)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile_Orig <> wFilex Then
        wFile_Orig = wFilex
    End If
    
    
End If

'_________________________________________
Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "GSOP KYC"
    .Subject = "GSOP KYC param"
End With
'____________________________________________________________________________________


Set wsExcel = wbExcel.Sheets(1)

wsExcel.Name = "KYC param"


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
wsExcel.PageSetup.Zoom = 72
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14GSOP KYC param  " _
                                & "&B&U&10     ( édité le " & dateImp10(DSys) & " " & Time & ")"
wsExcel.PageSetup.PrintTitleRows = "$A1:$G1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


wsExcel.Columns(1).ColumnWidth = 15: wsExcel.Cells(1, 1) = "Nature"
wsExcel.Columns(2).ColumnWidth = 7: wsExcel.Cells(1, 2) = "Code"
wsExcel.Columns(2).Font.Bold = True
wsExcel.Columns(3).ColumnWidth = 6: wsExcel.Cells(1, 3) = "N°"
wsExcel.Columns(3).Font.Bold = True

wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(4).ColumnWidth = 6: wsExcel.Cells(1, 4) = "Doc"
wsExcel.Columns(4).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(4).Font.Bold = True
wsExcel.Columns(5).ColumnWidth = 11: wsExcel.Cells(1, 5) = "Etat"
wsExcel.Columns(6).ColumnWidth = 100: wsExcel.Cells(1, 6) = "Libellé"
wsExcel.Columns(6).Font.Size = 8
wsExcel.Columns(7).ColumnWidth = 27: wsExcel.Cells(1, 7) = "Mise à jour"


mXls2_Col = 7: mXls2_Row = 1

For K = 1 To mXls2_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next

Call lstParam_YKYCDOS0_J_Load
Call lstParam_YKYCDOS0_D_Load

X = "select * from " & paramIBM_Library_SABSPE & ".YKYCDOS0" _
   & " where KYCDOSNAT <> ' '" _
  & " order by KYCDOSId ,KYCDOSNAT , KYCDOSSEQ , KYCDOSSEQ2"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    mXls2_Row = mXls2_Row + 1
    Select Case rsSab("KYCDOSNAT")
        Case "*": wsExcel.Cells(mXls2_Row, 1) = "Catégorie client": wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_G2
                                                                    wsExcel.Cells(mXls2_Row, 2).Interior.Color = mColor_G2
                                                                    wsExcel.Cells(mXls2_Row, 3).Interior.Color = mColor_G2
                                                                    wsExcel.Cells(mXls2_Row, 4).Interior.Color = mColor_G2
                                                                    wsExcel.Cells(mXls2_Row, 5).Interior.Color = mColor_G2
                                                                    wsExcel.Cells(mXls2_Row, 6).Interior.Color = mColor_G2
                                                                    wsExcel.Cells(mXls2_Row, 7).Interior.Color = mColor_G2
                                                                    wsExcel.Cells(mXls2_Row, 6).Font.Bold = True
        Case "=": 'wsExcel.Cells(mXls2_Row, 1) = "Doc / Catégorie"
        Case "D": wsExcel.Cells(mXls2_Row, 1) = "Document": wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_Y0
        Case "J": wsExcel.Cells(mXls2_Row, 1) = "justificatif":: wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_Y1
        Case Else: wsExcel.Cells(mXls2_Row, 1) = rsSab("KYCDOSNAT"): wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_W0
    End Select
   
   Select Case rsSab("KYCDOSSTAK")
        'Case "E": wsExcel.Cells(mXls2_Row, 5) = "Echéance"
        Case "O": wsExcel.Cells(mXls2_Row, 5) = "Obligatoire"
    End Select
    
        K1 = rsSab("KYCDOSSEQ")
        K2 = rsSab("KYCDOSSEQ2")
        If K1 = 0 Then
                        wsExcel.Cells(mXls2_Row, 2) = rsSab("KYCDOSID")
                        wsExcel.Cells(mXls2_Row, 6) = rsSab("KYCDOSDLIB")
        Else
            If K2 = 0 Then
                    wsExcel.Cells(mXls2_Row, 3) = K1
                    wsExcel.Cells(mXls2_Row, 6) = arrKYCDOSDLIB_J(K1)
                Select Case rsSab("KYCDOSNAT")
                   Case "J":  wsExcel.Cells(mXls2_Row, 6) = arrKYCDOSDLIB_J(K1)
                                wsExcel.Cells(mXls2_Row, 2) = rsSab("KYCDOSID")
                   Case "D":  wsExcel.Cells(mXls2_Row, 6) = arrKYCDOSDLIB_D(K1)
                                  wsExcel.Cells(mXls2_Row, 2) = rsSab("KYCDOSID")
                 'Case "*":
                  Case "="
                        wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_G1
                        wsExcel.Cells(mXls2_Row, 2).Interior.Color = mColor_G1
                        wsExcel.Cells(mXls2_Row, 3).Interior.Color = mColor_G1
                        wsExcel.Cells(mXls2_Row, 4).Interior.Color = mColor_G1
                        wsExcel.Cells(mXls2_Row, 5).Interior.Color = mColor_G1
                        wsExcel.Cells(mXls2_Row, 6).Interior.Color = mColor_G1
                        wsExcel.Cells(mXls2_Row, 7).Interior.Color = mColor_G1
                        wsExcel.Cells(mXls2_Row, 6).Font.Bold = True
                        If Trim(rsSab("KYCDOSDLIB")) <> "" Then
                            mXls2_Row = mXls2_Row + 1
                            wsExcel.Cells(mXls2_Row, 6) = Trim(rsSab("KYCDOSDLIB"))
                            wsExcel.Cells(mXls2_Row, 6).Font.Color = RGB(128, 0, 0)
                            wsExcel.Cells(mXls2_Row, 6).Font.Italic = True
                        End If
                End Select
            Else
                wsExcel.Cells(mXls2_Row, 4) = K2
                wsExcel.Cells(mXls2_Row, 6) = arrKYCDOSDLIB_D(K2)
           End If
        End If
    'End If
    wsExcel.Cells(mXls2_Row, 7) = rsSab("KYCDOSUUSR") & " " & dateImp10_S(rsSab("KYCDOSUAMJ")) & " " & timeImp8(Format(rsSab("KYCDOSUHMS"), "000000"))
    rsSab.MoveNext
Loop

'____________________________________________________________________________________



wbExcel.SaveAs wFile_Orig

wbExcel.Close

appExcel.Quit
Set rsSab = Nothing

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

Public Sub fraYKYCDOS0_Update_Display()
Dim X As String, blnUpdate As Boolean
fraPJ.Visible = False

If oldYKYCDOS0.KYCDOSSEQ2 = 0 Then
    blnUpdate = False
    txtYKYCDOS0_KYCDOSDECH.Visible = False: lblYKYCDOS0_KYCDOSDECH.Visible = False
    txtYKYCDOS0_KYCDOSDLIB.Visible = False: lblYKYCDOS0_KYCDOSDLIB.Visible = False
    txtYKYCDOS0_KYCDOSDAMJ.Visible = False: lblYKYCDOS0_KYCDOSDAMJ.Visible = False
    fgPJ.Visible = False
Else
    blnUpdate = arrHab(15)
    txtYKYCDOS0_KYCDOSDECH.Visible = True: lblYKYCDOS0_KYCDOSDECH.Visible = True
    txtYKYCDOS0_KYCDOSDECH.Enabled = arrHab(15)
    txtYKYCDOS0_KYCDOSDLIB.Visible = True: lblYKYCDOS0_KYCDOSDLIB.Visible = True
    txtYKYCDOS0_KYCDOSDLIB.Enabled = arrHab(15)
    txtYKYCDOS0_KYCDOSDAMJ.Visible = True: lblYKYCDOS0_KYCDOSDAMJ.Visible = True
    txtYKYCDOS0_KYCDOSDAMJ.Enabled = arrHab(15)
    fgPJ.Visible = True
End If

cmdPJ_Delete.Visible = False
fgPJ.Rows = 1
'libYKYCDOS0_Comment
If oldYKYCDOS0.KYCDOSSTAK = " " Or oldYKYCDOS0.KYCDOSSTAK = "O" Then
    oldYKYCDOS0.KYCDOSDAMJ = DSys
    'oldYKYCDOS0.KYCDOSDECH = DSys
    txtYKYCDOS0_KYCDOSDLIB = ""
    cmdYKYCDOS0_Delete.Visible = False
    cmdYKYCDOS0_Update.Visible = False
    cmdYKYCDOS0_PJ.Visible = False
    cmdYKYCDOS0_Add.Visible = blnUpdate
    cmdYKYCDOS0_Missing.Visible = blnUpdate
    cmdYKYCDOS0_Ignore.Visible = blnUpdate
    'libYKYCDOS0_Document.ForeColor = &H800080
    fraYKYCDOS0_Update.ForeColor = vbBlue '&H800080
    lblYKYCDOS0_KYCDOSUUSR = ""
Else
    txtYKYCDOS0_KYCDOSDLIB = oldYKYCDOS0.KYCDOSDLIB
    cmdYKYCDOS0_Delete.Visible = blnUpdate
    cmdYKYCDOS0_Update.Visible = blnUpdate
    cmdYKYCDOS0_PJ.Visible = blnUpdate
    cmdYKYCDOS0_Add.Visible = False
    cmdYKYCDOS0_Missing.Visible = False
    cmdYKYCDOS0_Ignore.Visible = False
    'libYKYCDOS0_Document.ForeColor = &H800080
    fraYKYCDOS0_Update.ForeColor = vbBlue ' &H800080
    lblYKYCDOS0_KYCDOSUUSR = oldYKYCDOS0.KYCDOSUUSR & "  " & dateImp10_S(oldYKYCDOS0.KYCDOSUAMJ) & "  " & timeImp8(oldYKYCDOS0.KYCDOSUHMS) & "  - " & oldYKYCDOS0.KYCDOSUFCT & oldYKYCDOS0.KYCDOSUVER
End If

Call DTPicker_Set(txtYKYCDOS0_KYCDOSDAMJ, CStr(oldYKYCDOS0.KYCDOSDAMJ))
If arrKYCDOSDECH_D(oldYKYCDOS0.KYCDOSSEQ2) <> 0 Then
    If oldYKYCDOS0.KYCDOSDECH = 0 Then
        X = dateElp("AnAdd", arrKYCDOSDECH_D(oldYKYCDOS0.KYCDOSSEQ2), CStr(oldYKYCDOS0.KYCDOSDAMJ))
        oldYKYCDOS0.KYCDOSDECH = Val(dateElp("Jour", -1, X))
    End If

    Call DTPicker_Set(txtYKYCDOS0_KYCDOSDECH, CStr(oldYKYCDOS0.KYCDOSDECH))
    txtYKYCDOS0_KYCDOSDECH.Visible = True
    lblYKYCDOS0_KYCDOSDECH.Visible = True
Else
    txtYKYCDOS0_KYCDOSDECH.Visible = False
    lblYKYCDOS0_KYCDOSDECH.Visible = False
End If

If Trim(oldYKYCDOS0.KYCDOSPJ) <> "" Then fgPJ_Display

libYKYCDOS0_Comment = arrYKYCDOS0_JD(fgYKYCDOS0.Row).KYCDOSDLIB
libYKYCDOS0_Comment.ForeColor = &H800080

If oldYKYCDOS0.KYCDOSSEQ2 >= 20 And oldYKYCDOS0.KYCDOSSEQ2 <= 23 Then
    Dim xCLIENCPIE As String, wCLILIBDA1 As Long, xCLILIBDA1 As String, wColor As Long
    txtYKYCDOS0_ZCLIENA0.Visible = True
    
    X = "select CLIENCPIE from " & paramIBM_Library_SAB & ".ZCLIENC0 where CLIENCCLI = '" & oldYKYCDOS0.KYCDOSID & "'"
    Set rsAdo = cnAdo.Execute(X)
    
    If rsAdo.EOF Then
        xCLIENCPIE = " erreur de lecture ZCLIENC0"
        wColor = vbRed
    Else
        xCLIENCPIE = rsAdo("CLIENCPIE")
         X = "select CLILIBDA1 from " & paramIBM_Library_SAB & ".ZCLILIB0 where CLILIBCLI = '" & oldYKYCDOS0.KYCDOSID & "'"
        Set rsAdo = cnAdo.Execute(X)
        
        If rsAdo.EOF Then
            xCLILIBDA1 = " pas de fiche KYC SAB"
             wColor = vbMagenta
        Else
            wCLILIBDA1 = rsAdo("CLILIBDA1")
            If wCLILIBDA1 = 0 Then
                xCLILIBDA1 = " date du document absente"
                 wColor = mColor_Y2
            Else
                xCLILIBDA1 = dateImp10_S(wCLILIBDA1 + 19000000)
                If oldYKYCDOS0.KYCDOSDAMJ = wCLILIBDA1 + 19000000 Then
                    wColor = mColor_G1
                Else
                    wColor = mColor_W0
                End If
                
            End If
            
        End If
        
    End If
    txtYKYCDOS0_ZCLIENA0 = "Fiche client SAB :      " & xCLIENCPIE & "     " & xCLILIBDA1
    txtYKYCDOS0_ZCLIENA0.BackColor = wColor

Else
    txtYKYCDOS0_ZCLIENA0.Visible = False
End If

fraYKYCDOS0_Update.Visible = True
fraYKYCDOS0_Update.ZOrder 0
End Sub


Public Function fraYKYCDOS0_Update_Control()
Dim V, X As String, blnOk As Boolean

newYKYCDOS0 = oldYKYCDOS0
blnOk = True
newYKYCDOS0.KYCDOSSTAK = " "
newYKYCDOS0.KYCDOSDLIB = Trim(txtYKYCDOS0_KYCDOSDLIB)
Call DTPicker_Control(txtYKYCDOS0_KYCDOSDAMJ, X)
newYKYCDOS0.KYCDOSDAMJ = X
If newYKYCDOS0.KYCDOSDAMJ > valDSys Then
    Call MsgBox("Date du document > aujourd'hui", vbCritical, "cmdYKYCDOS0_Add_Click")
    blnOk = False
End If

If arrKYCDOSDECH_D(newYKYCDOS0.KYCDOSSEQ2) <> 0 Then
    Call DTPicker_Control(txtYKYCDOS0_KYCDOSDECH, X)
    newYKYCDOS0.KYCDOSDECH = X
    If newYKYCDOS0.KYCDOSDECH < newYKYCDOS0.KYCDOSDAMJ Then
        Call MsgBox("Echéance < date du document", vbCritical, "cmdYKYCDOS0_Add_Click")
        blnOk = False
    End If
Else
    newYKYCDOS0.KYCDOSDECH = 0
End If
fraYKYCDOS0_Update_Control = blnOk
End Function

Public Sub fgPJ_Display()
Dim X As String, objFolder, objFiles, fsoFile As File

fgPJ.Rows = 1
X = oldYKYCDOS0.KYCDOSID & "\"
If Dir(paramGSOP_Dossier_Path_DROPI & X) <> "" Then
    Set objFolder = msFileSystem.GetFolder(paramGSOP_Dossier_Path_DROPI & X)
    Set objFiles = objFolder.Files
    X = oldYKYCDOS0.KYCDOSID & "_" & oldYKYCDOS0.KYCDOSSEQ2
    For Each fsoFile In objFiles
        'If InStr(fsoFile.Type, "Document") > 0 Then
        If InStr(fsoFile.Name, X) Then
            fgPJ.Rows = fgPJ.Rows + 1
            fgPJ.Row = fgPJ.Rows - 1
            'fgPJ.Col = 0: fgPJ.Text = fsoFile.DateCreated
            fgPJ.Col = 0: fgPJ.Text = fsoFile.Name
        'End If
        End If
    Next
End If

End Sub

Private Sub txtUpdLog_AmjMin_Change()
fgSelect.Visible = False

End Sub

Private Sub txtUpdLog_CLIRGPCLI_Change()
fgSelect.Visible = False

End Sub

Private Sub txtUpdLog_CLIRGPREG_Change()
fgSelect.Visible = False

End Sub


Public Sub cmdSelect_SQL_YKYCDOS0()
Dim X As String, K As Long, wColor As Long, K1 As Long, K2 As Long
Dim xSQL As String, mCLIENACLI As String, blnClient As Boolean, wCLIENARA As String, blnClos As Boolean
Dim xAnd As String
On Error GoTo Error_Handler

paramXgsop_Init

Select Case cboSelect_Options_KYCgsop_CLIENARES
    Case "*": xAnd = ""
    Case "* + R": V = "Interdit pour cette fonction": GoTo Error_MsgBox
            
    Case Else
            xAnd = " and CLIENARes = '" & cboSelect_Options_KYCgsop_CLIENARES.Text & "'"
End Select

If optSelect_Options_KYCgsop_NOK Then xAnd = xAnd & " and KYCDOSSTAK <> ' ' "
If optSelect_Options_KYCgsop_OK Then xAnd = xAnd & " and KYCDOSSTAK = ' ' "




fgDetail.Visible = False
fgDetail_Reset

fgDetail.FormatString = "<Racine        |<Client                                                                                                " _
                       & "|<Ges|<Document                                                                                                                   " _
                       & "|<Date Doc         |<Echéance         |<PJ|<Commentaires                                                                                                                                                |>J   |>D   "
         
fgDetail.Left = 100
fgDetail.Top = 150
fgDetail.Width = 13000

fgDetail.Rows = 1
fgDetail.Row = 0


X = "select * from " & paramIBM_Library_SABSPE & ".YKYCDOS0 , " & paramIBM_Library_SAB & ".ZCLIENA0  " _
  & " where KYCDOSID = clienacli and KYCDOSNAT = ' ' and KYCDOSSEQ = 0 and KYCDOSSEQ2 = 0" & xAnd _
  & " order by KYCDOSID"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    blnClos = True
    wCLIENARA = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
    If Not retourne_Client_CLOS(rsSab("CLIENACLI"), wCLIENARA) Then
    'If InStr(wCLIENARA, "CLOS") = 0 Then
        'If InStr(wCLIENARA, "CLOTURE") = 0 Then
            blnClos = False
        'End If
    End If
    If Not blnClos Then
        fgDetail.Rows = fgDetail.Rows + 1
        fgDetail.Row = fgDetail.Rows - 1
        If rsSab("KYCDOSSTAK") = " " Then
            wColor = mColor_G1
        Else
            wColor = mColor_W0
        End If
        fgDetail.Col = 0: fgDetail.Text = rsSab("CLIENACLI")
        fgDetail.Col = 1: fgDetail.Text = wCLIENARA
        fgDetail.Col = 2: fgDetail.Text = rsSab("CLIENARES")
        For I = 0 To 9: fgDetail.Col = I: fgDetail.CellBackColor = wColor: Next I
    '__________________________________________________________________________________
        If optSelect_Options_KYCgsop_Detail_NOK Then
        
        Else
        
            Call rsYKYCDOS0_GetBuffer(rsSab, currentYKYCDOS0)
            Call cmdSelect_SQL_YKYCDOS0_Load
            If optSelect_Options_KYCgsop_Detail_Missing Then
                cmdSelect_SQL_YKYCDOS0_Detail_Missing
            Else
                cmdSelect_SQL_YKYCDOS0_Detail_Ok
            End If
    
        End If
    End If
    rsSab.MoveNext
Loop

'__________________________________________________________________________________
If optSelect_Options_KYCgsop_All Then

    X = "select * from " & paramIBM_Library_SAB & ".ZTITULA0 , " & paramIBM_Library_SAB & ".ZCLIENA0 , " _
        & paramIBM_Library_SAB & ".ZCOMPTE0 , " & paramIBM_Library_SAB & ".ZPLAN0  " _
       & " where clienacli = titulacli " & xAnd _
      & " and titulacom = comptecom " _
      & " and compteobl = PLANCOOBL " _
      & " and clienacli not in (select KYCDOSID from " & paramIBM_Library_SABSPE & ".YKYCDOS0 where KYCDOSNAT = ' ' and KYCDOSSEQ = 0 and KYCDOSSEQ2 = 0)" _
      & " order by clienacli"
    
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        
        If oldZCLIENA0.CLIENACLI <> rsSab("CLIENACLI") Then
            If blnClient Then
                fgDetail.Rows = fgDetail.Rows + 1
                fgDetail.Row = fgDetail.Rows - 1
                fgDetail.Col = 0: fgDetail.Text = oldZCLIENA0.CLIENACLI
                fgDetail.Col = 1: fgDetail.Text = Trim(oldZCLIENA0.CLIENARA1) & " " & Trim(oldZCLIENA0.CLIENARA2)
                fgDetail.Col = 2: fgDetail.Text = oldZCLIENA0.CLIENARES
            End If
            Call rsZCLIENA0_GetBuffer(rsSab, oldZCLIENA0)
            blnClient = False
        End If
        
        X = rsSab("PLANCOPRO")
        
        
        If InStr(paramXgsop_PLANCOPRO, X) > 0 Then
            If rsSab("COMPTEFON") <> 4 Then blnClient = True
        End If
        
        rsSab.MoveNext
    Loop
End If

'__________________________________________________________________________________

fgDetail.Visible = True
fraSelect.Visible = True
'__________________________________________________________________________________
Exit_sub:
'__________________________________________________________________________________
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, Me.Name
    Call lstErr_AddItem(lstErr, cmdContext, "< cmdSelect_SQL_YKYCDOS0"): DoEvents

End Sub
Public Sub cmdSelect_SQL_YKYCDOS0_Ech()
Dim X As String, K As Long, wColor As Long, K1 As Long, K2 As Long
Dim xSQL As String, mCLIENACLI As String, blnClient As Boolean, wCLIENARA As String, blnClos As Boolean
Dim xAnd As String
On Error GoTo Error_Handler

lstParam_YKYCDOS0_Load
paramXgsop_Init

Call DTPicker_Control(txtSelect_Options_KYCech_KYCDOSDECH, wAMJMin)
xAnd = " and KYCDOSDECH <= " & wAMJMin

If Not blnAuto Then
    Select Case cboSelect_Options_KYCech_CLIENARES
        Case "*"
        Case "": xAnd = xAnd & " and substring( CLIENARES , 1 , 1) <> 'X'"
        Case "* + R": V = "Interdit pour cette fonction": GoTo Error_MsgBox
                
        Case Else
                xAnd = xAnd & " and CLIENARES = '" & cboSelect_Options_KYCech_CLIENARES.Text & "'"
    End Select
    
    
    X = Trim(cboSelect_Options_KYCech_Doc)
    If X <> "" Then
        K = InStr(X, ":")
        If K > 0 Then
            xAnd = xAnd & " and KYCDOSSEQ2 = " & Val(Mid$(X, 1, K - 1))
        End If
    End If
End If

fgDetail.Visible = False
fgDetail_Reset

fgDetail.FormatString = "<Racine        |<Client                                                                                                " _
                       & "|<Ges|<Document                                                                                                                   " _
                       & "|<Date Doc         |<Echéance         |<PJ|<Commentaires                                                                                                                                                |>J   |>D   "
         
fgDetail.Left = 100
fgDetail.Top = 150
fgDetail.Width = 13000

fgDetail.Rows = 1
fgDetail.Row = 0


X = "select * from " & paramIBM_Library_SABSPE & ".YKYCDOS0 , " & paramIBM_Library_SAB & ".ZCLIENA0  " _
  & " where KYCDOSID = clienacli and KYCDOSNAT = ' ' and KYCDOSDECH > 0  " & xAnd _
  & " order by KYCDOSID"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    blnClos = True
    wCLIENARA = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
    If Not retourne_Client_CLOS(rsSab("CLIENACLI"), wCLIENARA) Then
    'If InStr(wCLIENARA, "CLOS") = 0 Then
        'If InStr(wCLIENARA, "CLOTURE") = 0 Then
            blnClos = False
        'End If
    End If
    If Not blnClos Then
        fgDetail.Rows = fgDetail.Rows + 1
        fgDetail.Row = fgDetail.Rows - 1
        If rsSab("KYCDOSSTAK") = " " Then
            wColor = mColor_G1
        Else
            wColor = mColor_W0
        End If
        
        fgDetail.Col = 0: fgDetail.Text = rsSab("CLIENACLI")
        fgDetail.Col = 1: fgDetail.Text = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
        fgDetail.Col = 2: fgDetail.Text = rsSab("CLIENARES")
     '   For I = 0 To 9: fgDetail.Col = I: fgDetail.CellBackColor = wColor: Next I
    '__________________________________________________________________________________
        arrYKYCDOS0_JD_Nb = 1: K = 1
        Call rsYKYCDOS0_GetBuffer(rsSab, xYKYCDOS0)
        fgDetail.Col = 3: fgDetail.Text = arrKYCDOSDLIB_D(xYKYCDOS0.KYCDOSSEQ2)
        fgDetail.Col = 4
        If xYKYCDOS0.KYCDOSDAMJ <> 0 Then fgDetail.Text = dateImp10(xYKYCDOS0.KYCDOSDAMJ)
        fgDetail.Col = 5
        If xYKYCDOS0.KYCDOSDECH <> 0 Then fgDetail.Text = dateImp10(xYKYCDOS0.KYCDOSDECH)
        If xYKYCDOS0.KYCDOSDECH < mKYCDOSDECH_Warn Then
            fgDetail.CellForeColor = vbRed
            fgDetail.CellBackColor = mColor_Y2
        Else
            fgDetail.CellForeColor = vbBlue
        End If
        fgDetail.Col = 6: fgDetail.Text = xYKYCDOS0.KYCDOSPJ
        If xYKYCDOS0.KYCDOSPJ <> " " Then fgDetail.CellBackColor = vbGreen
        If Trim(xYKYCDOS0.KYCDOSDLIB) <> "" Then fgDetail.Col = 7: fgDetail.Text = xYKYCDOS0.KYCDOSDLIB
    End If

    rsSab.MoveNext
Loop

'__________________________________________________________________________________

'__________________________________________________________________________________

fgDetail.Visible = True
fraSelect.Visible = True
'__________________________________________________________________________________
Exit_sub:
'__________________________________________________________________________________
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, Me.Name
    Call lstErr_AddItem(lstErr, cmdContext, "< cmdSelect_SQL_YKYCDOS0"): DoEvents

End Sub


Public Sub paramXgsop_Init()

paramXgsop_NonRéclamés = ""
paramXgsop_HorsGsop = ""
paramXgsop_PLANCOPRO = ""

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
   & " where BIATABID = 'SAB_CLIENT'" _
  & " order by BIATABK1 , BIATABK2"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    X = Trim(rsSab("BIATABK2")) & ";"
    Select Case Trim(rsSab("BIATABK1"))
        Case "1": paramXgsop_NonRéclamés = paramXgsop_NonRéclamés & X
        Case "2": paramXgsop_HorsGsop = paramXgsop_HorsGsop & X
        Case "3": paramXgsop_PLANCOPRO = paramXgsop_PLANCOPRO & X
    End Select
    
    rsSab.MoveNext
Loop

End Sub

Private Sub txtYKYCDOS0_KYCDOSDAMJ_Change()
If arrKYCDOSDECH_D(oldYKYCDOS0.KYCDOSSEQ2) <> 0 Then
    Dim X As String, X1 As String
    Call DTPicker_Control(txtYKYCDOS0_KYCDOSDAMJ, X)
    X1 = dateElp("AnAdd", arrKYCDOSDECH_D(oldYKYCDOS0.KYCDOSSEQ2), X)
    X = dateElp("Jour", -1, X1)
    Call DTPicker_Set(txtYKYCDOS0_KYCDOSDECH, X)
End If

End Sub



Public Sub cmdSelect_SQL_YKYCDOS0_Detail_Missing()
Dim X As String, K As Long, wColor As Long, K1 As Long, K2 As Long
Dim blnDisplay As Boolean
For K = 1 To arrYKYCDOS0_JD_Nb
    
    K1 = arrYKYCDOS0_JD(K).KYCDOSSEQ
    K2 = arrYKYCDOS0_JD(K).KYCDOSSEQ2
    blnDisplay = False
    If K2 = 0 Then
        
        Select Case arrYKYCDOS0(K).KYCDOSSTAK
            Case "=":
            Case " ":
            Case Else: blnDisplay = True
                    fgDetail.Rows = fgDetail.Rows + 1
                    fgDetail.Row = fgDetail.Rows - 1
                    fgDetail.CellFontBold = True: fgDetail.CellFontSize = 9

                    wColor = mColor_W1: arrYKYCDOS0(0).KYCDOSSTAK = "N"
                    fgDetail.Col = 5: fgDetail.Text = "obligatoire": fgDetail.CellForeColor = vbRed

        End Select
        If blnDisplay Then
            fgDetail.Col = 3: fgDetail.Text = arrKYCDOSDLIB_J(K1)
            fgDetail.CellFontBold = True: fgDetail.CellFontSize = 9
    
            fgDetail.CellBackColor = wColor
        End If
        
    Else
    
      Select Case arrYKYCDOS0(K).KYCDOSSTAK
        Case "=":
                If arrKYCDOSDECH_D(K2) <> 0 And arrYKYCDOS0(K).KYCDOSDECH < mKYCDOSDECH_Warn Then
                    blnDisplay = True
                    fgDetail.Rows = fgDetail.Rows + 1
                    fgDetail.Row = fgDetail.Rows - 1
                End If

        Case "?": blnDisplay = True
                    fgDetail.Rows = fgDetail.Rows + 1
                    fgDetail.Row = fgDetail.Rows - 1
                For I = 3 To 9: fgDetail.Col = I: fgDetail.CellBackColor = mColor_W0: Next I
                 fgDetail.Col = 5: fgDetail.Text = "manquant": fgDetail.CellForeColor = vbRed
        Case "O": blnDisplay = True
                    fgDetail.Rows = fgDetail.Rows + 1
                    fgDetail.Row = fgDetail.Rows - 1
                For I = 3 To 9: fgDetail.Col = I: fgDetail.CellBackColor = mColor_Y2: Next I
                 fgDetail.Col = 5: fgDetail.Text = "obligatoire": fgDetail.CellForeColor = vbRed
        Case "I":
      End Select
       
    If blnDisplay Then
        fgDetail.Col = 8: fgDetail.Text = K1
        fgDetail.Col = 9: fgDetail.Text = K2
            
          fgDetail.Col = 3: fgDetail.Text = arrKYCDOSDLIB_D(K2)
          fgDetail.CellFontSize = 8
    
          fgDetail.Col = 4
          If arrYKYCDOS0(K).KYCDOSDAMJ <> 0 Then fgDetail.Text = dateImp10(arrYKYCDOS0(K).KYCDOSDAMJ)
          
          fgDetail.Col = 5
          If arrKYCDOSDECH_D(K2) <> 0 And arrYKYCDOS0(K).KYCDOSSTAK = "=" Then
            If arrYKYCDOS0(K).KYCDOSDECH <> 0 Then fgDetail.Text = dateImp10(arrYKYCDOS0(K).KYCDOSDECH)
            If arrYKYCDOS0(K).KYCDOSDECH < mKYCDOSDECH_Warn Then
                fgDetail.CellForeColor = vbRed
                fgDetail.CellBackColor = mColor_Y2
            Else
                fgDetail.CellForeColor = vbBlue
            End If
          End If
            fgDetail.Col = 6: fgDetail.Text = arrYKYCDOS0(K).KYCDOSPJ
            If arrYKYCDOS0(K).KYCDOSPJ <> " " Then fgDetail.CellBackColor = vbGreen
            
       End If
       
       fgDetail.Col = 7
       If Trim(arrYKYCDOS0(K).KYCDOSDLIB) <> "" Then
            fgDetail.Text = arrYKYCDOS0(K).KYCDOSDLIB
       Else
            fgDetail.Text = arrYKYCDOS0_JD(K).KYCDOSDLIB
            fgDetail.CellFontItalic = True
            fgDetail.CellForeColor = &H800080
       End If
   End If
Next K

End Sub

Public Sub cmdSelect_SQL_YKYCDOS0_Detail_Ok()
Dim X As String, K As Long, wColor As Long, K1 As Long, K2 As Long

For K = 1 To arrYKYCDOS0_JD_Nb
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    
    K1 = arrYKYCDOS0_JD(K).KYCDOSSEQ
    K2 = arrYKYCDOS0_JD(K).KYCDOSSEQ2
    
    fgDetail.Col = 8: fgDetail.Text = K1
    fgDetail.Col = 9: fgDetail.Text = K2
    If K2 = 0 Then
        fgDetail.CellFontBold = True: fgDetail.CellFontSize = 9
        
        Select Case arrYKYCDOS0(K).KYCDOSSTAK
            Case "=": wColor = mColor_G2
            Case " ": wColor = vbYellow
            Case Else: wColor = mColor_W1: arrYKYCDOS0(0).KYCDOSSTAK = "N"
                             fgDetail.Col = 5: fgDetail.Text = "obligatoire": fgDetail.CellForeColor = vbRed

        End Select

        fgDetail.Col = 3: fgDetail.Text = arrKYCDOSDLIB_J(K1)
        fgDetail.CellFontBold = True: fgDetail.CellFontSize = 9

        fgDetail.CellBackColor = wColor
    Else
      Select Case arrYKYCDOS0(K).KYCDOSSTAK
        Case "=":
                For I = 3 To 9: fgDetail.Col = I: fgDetail.CellBackColor = mColor_G1: Next I
        Case "?": arrYKYCDOS0(0).KYCDOSSTAK = "N"
                For I = 3 To 9: fgDetail.Col = I: fgDetail.CellBackColor = mColor_W0: Next I
                 fgDetail.Col = 5: fgDetail.Text = "manquant": fgDetail.CellForeColor = vbRed
        Case "O": arrYKYCDOS0(0).KYCDOSSTAK = "N"
                For I = 3 To 9: fgDetail.Col = I: fgDetail.CellBackColor = mColor_Y2: Next I
                 fgDetail.Col = 5: fgDetail.Text = "obligatoire": fgDetail.CellForeColor = vbRed
        Case "I":
                For I = 3 To 9: fgDetail.Col = I: fgDetail.CellBackColor = RGB(230, 230, 230): Next I
                  fgDetail.Col = 5: fgDetail.Text = "doc ignoré": fgDetail.CellForeColor = vbRed
      End Select
        
      fgDetail.Col = 3: fgDetail.Text = arrKYCDOSDLIB_D(K2)
      fgDetail.CellFontSize = 8

      fgDetail.Col = 4
      If arrYKYCDOS0(K).KYCDOSDAMJ <> 0 Then fgDetail.Text = dateImp10(arrYKYCDOS0(K).KYCDOSDAMJ)
      
      fgDetail.Col = 5
      If arrKYCDOSDECH_D(K2) <> 0 And arrYKYCDOS0(K).KYCDOSSTAK = "=" Then
        If arrYKYCDOS0(K).KYCDOSDECH <> 0 Then fgDetail.Text = dateImp10(arrYKYCDOS0(K).KYCDOSDECH)
        If arrYKYCDOS0(K).KYCDOSDECH < mKYCDOSDECH_Warn Then
            fgDetail.CellForeColor = vbRed
            fgDetail.CellBackColor = mColor_Y2
        Else
            fgDetail.CellForeColor = vbBlue
        End If
      End If
        fgDetail.Col = 6: fgDetail.Text = arrYKYCDOS0(K).KYCDOSPJ
        If arrYKYCDOS0(K).KYCDOSPJ <> " " Then fgDetail.CellBackColor = vbGreen
        
   End If
   
   fgDetail.Col = 7
   If Trim(arrYKYCDOS0(K).KYCDOSDLIB) <> "" Then
        fgDetail.Text = arrYKYCDOS0(K).KYCDOSDLIB
   Else
        fgDetail.Text = arrYKYCDOS0_JD(K).KYCDOSDLIB
        fgDetail.CellFontItalic = True
        fgDetail.CellForeColor = &H800080
   End If
   
Next K

End Sub

Public Sub cmdSelect_SQL_Xgsop_Auto()
Dim wSendMail As typeSendMail
blnAuto = True

blnYKYCSTA0_Update = True
Call rsYKYCSTA0_Init(newYKYCSTA0)

Call sqlYKYCSTA0_Delete_Where(" where KYCSTADSIT = " & YBIATAB0_DATE_CPT_J)

Call cmdSelect_SQL_Xgsop

Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S11", wFile_Orig & ".xlsx", "Archive", "XGSOP")

wSendMail.From = currentSSIWINMAIL
If blnAuto Then
    wSendMail.FromDisplayName = "@Xgsop"
    wSendMail.RecipientDisplayName = "SAB_CLIENT"
Else
    wSendMail.Recipient = currentSSIWINMAIL
End If

wSendMail.Subject = "GSOP Reporting " & dateImp10_S(YBIATAB0_DATE_CPT_J)

wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                    & "<span style='font-size:10.0pt;font-family:Arial Unicode MS'>" & "<Font color = #404040>"


wSendMail.AsHTML = True
wSendMail.Attachment = "" ' wFile_Orig & ".xlsx"

wSendMail.Message = mHtml_Head & "<span style='font-size:10.0pt;font-family:Calibri'>" & paramEditionNoPaper_Auto_Lnk & "<BR><BR>" _
                 & htmlFontColor_Black & "</div></body></html>"

srvSendMail.Monitor wSendMail


End Sub

Public Sub cmdSelect_SQL_Xgsop_Auto_YKYCSTA0()
Dim xSQL As String, blnUpdate As Boolean
On Error GoTo Error_Handler
    
    
    newYKYCSTA0.KYCSTADSIT = YBIATAB0_DATE_CPT_J
    newYKYCSTA0.KYCSTAYVER = 0
    
    newYKYCSTA0.KYCSTASTAX = Trim(wsExcel.Cells(mXls2_Row, 1))
    newYKYCSTA0.KYCSTASTAY = Trim(wsExcel.Cells(mXls2_Row, 2))
    newYKYCSTA0.KYCSTACLI = "00" & Trim(wsExcel.Cells(mXls2_Row, 3))
    newYKYCSTA0.KYCSTAZCOL = UCase(Mid$(Trim(wsExcel.Cells(mXls2_Row, 4)) & " ", 1, 1))
    newYKYCSTA0.KYCSTAZETA = Mid$(Trim(wsExcel.Cells(mXls2_Row, 5)) & "     ", 1, 4)
    newYKYCSTA0.KYCSTAZCAT = Mid$(Trim(wsExcel.Cells(mXls2_Row, 6)) & "     ", 1, 3)

    newYKYCSTA0.KYCSTACAVC = Val(wsExcel.Cells(mXls2_Row, 7))
    newYKYCSTA0.KYCSTACAVT = Val(wsExcel.Cells(mXls2_Row, 8))
    newYKYCSTA0.KYCSTACAVX = Val(wsExcel.Cells(mXls2_Row, 9))

    newYKYCSTA0.KYCSTATECC = Val(wsExcel.Cells(mXls2_Row, 10))
    newYKYCSTA0.KYCSTATECT = Val(wsExcel.Cells(mXls2_Row, 11))
    newYKYCSTA0.KYCSTATECX = Val(wsExcel.Cells(mXls2_Row, 12))
    
    newYKYCSTA0.KYCSTADCLO = Val(wsExcel.Cells(mXls2_Row, 13))
    
    newYKYCSTA0.KYCSTAZRES = Mid$(Trim(wsExcel.Cells(mXls2_Row, 14)) & "     ", 1, 3)
    newYKYCSTA0.KYCSTAZRA1 = Trim(wsExcel.Cells(mXls2_Row, 15))
    If Len(newYKYCSTA0.KYCSTAZRA1) > 64 Then newYKYCSTA0.KYCSTAZRA1 = Mid$(newYKYCSTA0.KYCSTAZRA1, 1, 64)
    newYKYCSTA0.KYCSTAZPCI = Trim(wsExcel.Cells(mXls2_Row, 16))
    If Len(newYKYCSTA0.KYCSTAZPCI) > 64 Then newYKYCSTA0.KYCSTAZPCI = Mid$(newYKYCSTA0.KYCSTAZPCI, 1, 64)
    newYKYCSTA0.KYCSTAYKYC = UCase(Mid$(Trim(wsExcel.Cells(mXls2_Row, 17)) & " ", 1, 1))
    newYKYCSTA0.KYCSTAZNAT = Mid$(Trim(wsExcel.Cells(mXls2_Row, 18)) & "     ", 1, 2)
    newYKYCSTA0.KYCSTAZRSD = Mid$(Trim(wsExcel.Cells(mXls2_Row, 19)) & "     ", 1, 2)

' maintenance : cmdSelect_SQL_Xgsop_Auto_YKYCSTA0 et
'=========================================================================================
blnUpdate = False
If newYKYCSTA0.KYCSTASTAY <> "0" Then
    If wsExcel.Cells(mXls2_Row, 20) = "Clos" Then
        If mKYCSTASTAK <> "9" Then
            newYKYCSTA0.KYCSTASTAK = "9"
            blnUpdate = True
            Call sqlYKYCSTA0_Insert(newYKYCSTA0)
        End If
    Else
        Select Case mKYCSTASTAK
            Case "": newYKYCSTA0.KYCSTASTAK = "0"
            Case "9": newYKYCSTA0.KYCSTASTAK = "2"
            Case Else: newYKYCSTA0.KYCSTASTAK = "1"
        End Select
        blnUpdate = True
        Call sqlYKYCSTA0_Insert(newYKYCSTA0)
    End If
End If

If blnUpdate Then
    If mKYCSTASTAX = newYKYCSTA0.KYCSTASTAX And mKYCSTASTAY = newYKYCSTA0.KYCSTASTAY Then
    Else
        If mKYCSTASTAX <> "" Then
            Debug.Print newYKYCSTA0.KYCSTACLI & " : x " & mKYCSTASTAX & ">" & newYKYCSTA0.KYCSTASTAX & " : Y " & mKYCSTASTAY & ">" & newYKYCSTA0.KYCSTASTAY
        End If
    End If
End If

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, wsExcel.Cells(mXls2_Row, 3) & " : cmdSelect_SQL_Xgsop_Auto_YKYCSTA0"): DoEvents

End Sub

Public Sub cmdSelect_SQL_Xgsop_Auto_Reprise()
'______________________________________________
Dim blnExit As Boolean, xSQL As String
On Error GoTo Error_Handler

Set appExcel = CreateObject("Excel.Application")


'YBIATAB0_DATE_CPT_J = "20140131"
'Set wbExcel = appExcel.Workbooks.Open("C:\TEMP\GSOP\GSOP reporting 20140131.xlsx")

'YBIATAB0_DATE_CPT_J = "20140228"
'Set wbExcel = appExcel.Workbooks.Open("C:\TEMP\GSOP\GSOP reporting 20140228.xlsx")

YBIATAB0_DATE_CPT_J = "20140331"
Set wbExcel = appExcel.Workbooks.Open("C:\TEMP\GSOP\GSOP reporting 20140331.xlsx")

Set wsExcel = wbExcel.Worksheets("Detail")
'__________________________________________________________________________________
Call sqlYKYCSTA0_Delete_Where(" where KYCSTADSIT = " & YBIATAB0_DATE_CPT_J)

mXls2_Row = 1
Do While Not blnExit
    mXls2_Row = mXls2_Row + 1
    If Trim(wsExcel.Cells(mXls2_Row, 1)) = "" Then
        blnExit = True
    Else
    
            xSQL = "select * from " & paramIBM_Library_SABSPE & ".YKYCSTA0" _
            & " where KYCSTACLI = '00" & Trim(wsExcel.Cells(mXls2_Row, 3)) & "'" _
            & " order by KYCSTADSIT desc"
            
        Set rsSab_YKYCSTA0 = cnsab.Execute(xSQL)
        If rsSab_YKYCSTA0.EOF Then
             mKYCSTASTAK = ""
             mKYCSTASTAX = ""
             mKYCSTASTAY = ""
             mKYCSTADCLO = 0
        Else
            mKYCSTASTAK = rsSab_YKYCSTA0("KYCSTASTAK")
            mKYCSTASTAX = rsSab_YKYCSTA0("KYCSTASTAX")
            mKYCSTASTAY = rsSab_YKYCSTA0("KYCSTASTAY")
            mKYCSTADCLO = rsSab_YKYCSTA0("KYCSTADCLO")
            If mKYCSTASTAX = Trim(wsExcel.Cells(mXls2_Row, 1)) And mKYCSTASTAY = Trim(wsExcel.Cells(mXls2_Row, 2)) Then
            Else
                newYKYCSTA0.KYCSTASTAK = "3"
                arrK3_Old(mKYCSTASTAX, mKYCSTASTAY) = arrK3_Old(mKYCSTASTAX, mKYCSTASTAY) + 1
                arrK3_New(mXgsop.X, mXgsop.y) = arrK3_New(mXgsop.X, mXgsop.y) + 1
            End If
        End If
        If mKYCSTASTAY = "5" And mKYCSTADCLO > 0 And mKYCSTADCLO < 20140000 Then
        Else
            Call cmdSelect_SQL_Xgsop_Auto_YKYCSTA0
        End If
    End If

Loop


'

'____________________________________________________________________________________
'wbExcel.Close
appExcel.Quit

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, wsExcel.Cells(mXls2_Row, 3) & " : cmdSelect_SQL_Xgsop_Auto_YKYCSTA0"): DoEvents

End Sub

Public Sub cmdSelect_SQL_Xgsop_Total_Ouv()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim X As String, K As Long
Dim xSQL As String, wFile As String, xRange_A As String, xRange_B As String, xRange_T As String

On Error GoTo Error_Handler

wsExcel.Columns(11).ColumnWidth = 20: wsExcel.Cells(11, 1) = " dont Ouvertures": 'wsExcel.Cells(11, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(12).ColumnWidth = 15: wsExcel.Cells(11, 2) = "Particuliers": ' wsExcel.Cells(11, 2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(13).ColumnWidth = 15: wsExcel.Cells(11, 3) = "Pers. morales": ' wsExcel.Cells(11, 3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(14).ColumnWidth = 15: wsExcel.Cells(11, 4) = "Banques": ' wsExcel.Cells(11, 4).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(15).ColumnWidth = 15: wsExcel.Cells(11, 5) = "Autres": 'wsExcel.Cells(11, 5).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(16).ColumnWidth = 15: wsExcel.Cells(11, 6) = "Total": 'wsExcel.Cells(11, 6).HorizontalAlignment = Excel.xlHAlignCenter

xRange_A = "Detail!A1:Detail!A" & mXls2_Row
xRange_B = "Detail!B1:Detail!B" & mXls2_Row
xRange_T = "Detail!T1:Detail!T" & mXls2_Row

wsExcel.Cells(12, 1) = "  Clients": wsExcel.Cells(12, 1).Font.Color = vbBlue
wsExcel.Cells(12, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";1;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";1;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(12, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";1;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";1;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(12, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";1;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";1;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(12, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";1;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";1;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(12, 6).FormulaLocal = "=SOMME(B12:E12)": wsExcel.Cells(12, 6).Interior.Color = mColor_G0
For K = 2 To 6
    If Val(wsExcel.Cells(12, K)) <> 0 Then
        wsExcel.Cells(12, K).Font.Color = vbBlue
    Else
        wsExcel.Cells(12, K) = ""
    End If
Next K

wsExcel.Cells(13, 1) = "  Techniques": wsExcel.Cells(13, 1).Font.Color = vbBlue
wsExcel.Cells(13, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";2;" & xRange_T & ";""Ouv"")" '_
                                  '& " + NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";2;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(13, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";2;" & xRange_T & ";""Ouv"")" ' _
                                  '& " + NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";2;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(13, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";2;" & xRange_T & ";""Ouv"")" '_
                                  '& " + NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";2;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(13, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";2;" & xRange_T & ";""Ouv"")" '_
                                  '& " + NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";2;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(13, 6).FormulaLocal = "=SOMME(B13:E13)": wsExcel.Cells(13, 6).Interior.Color = mColor_G0
For K = 2 To 6
    If Val(wsExcel.Cells(13, K)) <> 0 Then
        wsExcel.Cells(13, K).Font.Color = vbBlue
    Else
        wsExcel.Cells(13, K) = ""
    End If
Next K
wsExcel.Cells(14, 1) = "  Tiers": wsExcel.Cells(14, 1).Font.Color = vbBlue
wsExcel.Cells(14, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";3;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";3;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(14, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";3;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";3;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(14, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";3;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";3;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(14, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";3;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";3;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(14, 6).FormulaLocal = "=SOMME(B14:E14)": wsExcel.Cells(14, 6).Interior.Color = mColor_G0
For K = 2 To 6
    If Val(wsExcel.Cells(14, K)) <> 0 Then
        wsExcel.Cells(14, K).Font.Color = vbBlue
    Else
        wsExcel.Cells(14, K) = ""
    End If
Next K

wsExcel.Cells(15, 1) = "  BIA non réclamés": wsExcel.Cells(15, 1).Font.Color = vbBlue
wsExcel.Cells(15, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";4;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";4;" & xRange_T & ";""Réouv" ' ")"
wsExcel.Cells(15, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";4;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";4;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(15, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";4;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";4;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(15, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";4;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";4;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(15, 6).FormulaLocal = "=SOMME(B15:E15)": wsExcel.Cells(15, 6).Interior.Color = mColor_G0
For K = 2 To 6
    If Val(wsExcel.Cells(15, K)) <> 0 Then
        wsExcel.Cells(15, K).Font.Color = vbBlue
    Else
        wsExcel.Cells(15, K) = ""
    End If
Next K

wsExcel.Cells(16, 1) = "  hors GSOP": wsExcel.Cells(16, 1).Font.Color = vbBlue
wsExcel.Cells(16, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";5;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";5;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(16, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";5;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";5;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(16, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";5;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";5;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(16, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";5;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";5;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(16, 6).FormulaLocal = "=SOMME(B16:E16)": wsExcel.Cells(16, 6).Interior.Color = mColor_G0
For K = 2 To 6
    If Val(wsExcel.Cells(16, K)) <> 0 Then
        wsExcel.Cells(16, K).Font.Color = vbBlue
    Else
        wsExcel.Cells(16, K) = ""
    End If
Next K

wsExcel.Cells(17, 1) = "  Racines sans compte": wsExcel.Cells(17, 1).Font.Color = vbBlue
wsExcel.Cells(17, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";6;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";6;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(17, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";6;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";6;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(17, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";6;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";6;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(17, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";6;" & xRange_T & ";""Ouv"")" ' _
                                  & " + NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";6;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(17, 6).FormulaLocal = "=SOMME(B17:E17)": wsExcel.Cells(17, 6).Interior.Color = mColor_G0
For K = 2 To 6
    If Val(wsExcel.Cells(17, K)) <> 0 Then
        wsExcel.Cells(17, K).Font.Color = vbBlue
    Else
        wsExcel.Cells(17, K) = ""
    End If
Next K

wsExcel.Cells(18, 1) = "  Total": wsExcel.Cells(18, 1).Font.Color = vbBlue
wsExcel.Cells(18, 2).FormulaLocal = "=SOMME(B12:B17)"
wsExcel.Cells(18, 3).FormulaLocal = "=SOMME(C12:C17)"
wsExcel.Cells(18, 4).FormulaLocal = "=SOMME(D12:D17)"
wsExcel.Cells(18, 5).FormulaLocal = "=SOMME(E12:E17)"
wsExcel.Cells(18, 6).FormulaLocal = "=SOMME(F12:F17)"
For K = 2 To 6
    If Val(wsExcel.Cells(18, K)) <> 0 Then
        wsExcel.Cells(18, K).Font.Color = vbBlue
    Else
        wsExcel.Cells(18, K) = ""
    End If
Next K

mXls1_Col = 6: mXls1_Row = 1

For K = 2 To mXls1_Col
    wsExcel.Cells(11, K).Interior.Color = mColor_G1
    wsExcel.Cells(11, K).Font.Color = vbBlue
Next

For K = 12 To 17
    wsExcel.Cells(K, 1).Interior.Color = mColor_G1
    wsExcel.Cells(K, 1).Font.Color = vbBlue
Next
For K = 2 To mXls1_Col
    wsExcel.Cells(18, K).Interior.Color = mColor_G0
Next

wsExcel.Cells(11, 1).Interior.Color = mColor_G9
wsExcel.Cells(11, 1).Font.Color = mColor_Z0
wsExcel.Cells(11, 6).Interior.Color = mColor_G9
wsExcel.Cells(11, 6).Font.Color = mColor_Z0
wsExcel.Cells(18, 1).Interior.Color = mColor_G9
wsExcel.Cells(18, 1).Font.Color = mColor_Z0
wsExcel.Cells(18, 6).Interior.Color = mColor_G2

'__________________________________________________________________________________
Exit_sub:
'=======================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée "): DoEvents

End Sub
Public Sub cmdSelect_SQL_Xgsop_Total_Clos()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim X As String, K As Long
Dim xSQL As String, wFile As String, xRange_A As String, xRange_B As String, xRange_T As String

On Error GoTo Error_Handler

wsExcel.Columns(21).ColumnWidth = 20: wsExcel.Cells(21, 1) = " dont Clôtures": 'wsExcel.Cells(21, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Columns(22).ColumnWidth = 15: wsExcel.Cells(21, 2) = "Particuliers": ' wsExcel.Cells(21, 2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(23).ColumnWidth = 15: wsExcel.Cells(21, 3) = "Pers. morales": ' wsExcel.Cells(21, 3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(24).ColumnWidth = 15: wsExcel.Cells(21, 4) = "Banques": ' wsExcel.Cells(21, 4).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(25).ColumnWidth = 15: wsExcel.Cells(21, 5) = "Autres": 'wsExcel.Cells(21, 5).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(26).ColumnWidth = 15: wsExcel.Cells(21, 6) = "Total": 'wsExcel.Cells(21, 6).HorizontalAlignment = Excel.xlHAlignCenter

xRange_A = "Detail!A1:Detail!A" & mXls2_Row
xRange_B = "Detail!B1:Detail!B" & mXls2_Row
xRange_T = "Detail!T1:Detail!T" & mXls2_Row

wsExcel.Cells(22, 1) = "  Clients": wsExcel.Cells(22, 1).Font.Color = vbRed
wsExcel.Cells(22, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";1;" & xRange_T & ";""Clos"")"
wsExcel.Cells(22, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";1;" & xRange_T & ";""Clos"")"
wsExcel.Cells(22, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";1;" & xRange_T & ";""Clos"")"
wsExcel.Cells(22, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";1;" & xRange_T & ";""Clos"")"
wsExcel.Cells(22, 6).FormulaLocal = "=SOMME(B22:E22)": wsExcel.Cells(22, 6).Interior.Color = mColor_W0
For K = 2 To 6
    If Val(wsExcel.Cells(22, K)) <> 0 Then
        wsExcel.Cells(22, K).Font.Color = vbRed
    Else
        wsExcel.Cells(22, K) = ""
    End If
Next K


wsExcel.Cells(23, 1) = "  Techniques": wsExcel.Cells(23, 1).Font.Color = vbRed
wsExcel.Cells(23, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";2;" & xRange_T & ";""Clos"")"
wsExcel.Cells(23, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";2;" & xRange_T & ";""Clos"")"
wsExcel.Cells(23, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";2;" & xRange_T & ";""Clos"")"
wsExcel.Cells(23, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";2;" & xRange_T & ";""Clos"")"
wsExcel.Cells(23, 6).FormulaLocal = "=SOMME(B23:E23)": wsExcel.Cells(23, 6).Interior.Color = mColor_W0
For K = 2 To 6
    If Val(wsExcel.Cells(23, K)) <> 0 Then
        wsExcel.Cells(23, K).Font.Color = vbRed
    Else
        wsExcel.Cells(23, K) = ""
    End If
Next K

wsExcel.Cells(24, 1) = "  Tiers": wsExcel.Cells(24, 1).Font.Color = vbRed
wsExcel.Cells(24, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";3;" & xRange_T & ";""Clos"")"
wsExcel.Cells(24, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";3;" & xRange_T & ";""Clos"")"
wsExcel.Cells(24, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";3;" & xRange_T & ";""Clos"")"
wsExcel.Cells(24, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";3;" & xRange_T & ";""Clos"")"
wsExcel.Cells(24, 6).FormulaLocal = "=SOMME(B24:E24)": wsExcel.Cells(24, 6).Interior.Color = mColor_W0
For K = 2 To 6
    If Val(wsExcel.Cells(24, K)) <> 0 Then
        wsExcel.Cells(24, K).Font.Color = vbRed
    Else
        wsExcel.Cells(24, K) = ""
    End If
Next K


wsExcel.Cells(25, 1) = "  BIA non réclamés": wsExcel.Cells(25, 1).Font.Color = vbRed
wsExcel.Cells(25, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";4;" & xRange_T & ";""Clos"")"
wsExcel.Cells(25, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";4;" & xRange_T & ";""Clos"")"
wsExcel.Cells(25, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";4;" & xRange_T & ";""Clos"")"
wsExcel.Cells(25, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";4;" & xRange_T & ";""Clos"")"
wsExcel.Cells(25, 6).FormulaLocal = "=SOMME(B25:E25)": wsExcel.Cells(25, 6).Interior.Color = mColor_W0
For K = 2 To 6
    If Val(wsExcel.Cells(25, K)) <> 0 Then
        wsExcel.Cells(25, K).Font.Color = vbRed
    Else
        wsExcel.Cells(25, K) = ""
    End If
Next K

wsExcel.Cells(26, 1) = "  hors GSOP": wsExcel.Cells(26, 1).Font.Color = vbRed
wsExcel.Cells(26, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";5;" & xRange_T & ";""Clos"")"
wsExcel.Cells(26, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";5;" & xRange_T & ";""Clos"")"
wsExcel.Cells(26, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";5;" & xRange_T & ";""Clos"")"
wsExcel.Cells(26, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";5;" & xRange_T & ";""Clos"")"
wsExcel.Cells(26, 6).FormulaLocal = "=SOMME(B26:E26)": wsExcel.Cells(26, 6).Interior.Color = mColor_W0
For K = 2 To 6
    If Val(wsExcel.Cells(26, K)) <> 0 Then
        wsExcel.Cells(26, K).Font.Color = vbRed
    Else
        wsExcel.Cells(26, K) = ""
    End If
Next K

wsExcel.Cells(27, 1) = "  Racines sans compte": wsExcel.Cells(27, 1).Font.Color = vbRed
wsExcel.Cells(27, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";6;" & xRange_T & ";""Clos"")" _
                                  & " + NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";0;" & xRange_T & ";""Clos"")"
wsExcel.Cells(27, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";6;" & xRange_T & ";""Clos"")" _
                                  & " + NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";0;" & xRange_T & ";""Clos"")"
wsExcel.Cells(27, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";6;" & xRange_T & ";""Clos"")" _
                                  & " + NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";0;" & xRange_T & ";""Clos"")"
wsExcel.Cells(27, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";6;" & xRange_T & ";""Clos"")" _
                                  & " + NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";0;" & xRange_T & ";""Clos"")"
wsExcel.Cells(27, 6).FormulaLocal = "=SOMME(B27:E27)": wsExcel.Cells(27, 6).Interior.Color = mColor_W0
For K = 2 To 6
    If Val(wsExcel.Cells(27, K)) <> 0 Then
        wsExcel.Cells(27, K).Font.Color = vbRed
    Else
        wsExcel.Cells(27, K) = ""
    End If
Next K

wsExcel.Cells(28, 1) = "  Total": wsExcel.Cells(28, 1).Font.Color = vbRed
wsExcel.Cells(28, 2).FormulaLocal = "=SOMME(B22:B27)": wsExcel.Cells(28, 2).Font.Color = vbRed
wsExcel.Cells(28, 3).FormulaLocal = "=SOMME(C22:C27)": wsExcel.Cells(28, 3).Font.Color = vbRed
wsExcel.Cells(28, 4).FormulaLocal = "=SOMME(D22:D27)": wsExcel.Cells(28, 4).Font.Color = vbRed
wsExcel.Cells(28, 5).FormulaLocal = "=SOMME(E22:E27)": wsExcel.Cells(28, 5).Font.Color = vbRed
wsExcel.Cells(28, 6).FormulaLocal = "=SOMME(F22:F27)": wsExcel.Cells(28, 6).Font.Color = vbRed

mXls1_Col = 6: mXls1_Row = 1

For K = 2 To mXls1_Col
    wsExcel.Cells(21, K).Interior.Color = mColor_W0
    wsExcel.Cells(21, K).Font.Color = vbRed
Next

For K = 22 To 27
    wsExcel.Cells(K, 1).Interior.Color = mColor_W0
    wsExcel.Cells(K, 1).Font.Color = vbRed
Next
For K = 2 To mXls1_Col
    wsExcel.Cells(28, K).Interior.Color = mColor_W0
Next

wsExcel.Cells(21, 1).Interior.Color = mColor_W1
wsExcel.Cells(21, 1).Font.Color = vbBlack
wsExcel.Cells(21, 6).Interior.Color = mColor_W1
wsExcel.Cells(21, 6).Font.Color = vbBlack
wsExcel.Cells(28, 1).Interior.Color = mColor_W1
wsExcel.Cells(28, 1).Font.Color = vbBlack
wsExcel.Cells(28, 6).Interior.Color = mColor_W1
wsExcel.Cells(28, 6).Font.Color = vbBlack

'__________________________________________________________________________________
Exit_sub:
'=======================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée "): DoEvents

End Sub

Public Sub cmdSelect_SQL_Xgsop_Total_Réouv()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim X As String, K As Long
Dim xSQL As String, wFile As String, xRange_A As String, xRange_B As String, xRange_T As String

On Error GoTo Error_Handler


wsExcel.Cells(35, 1) = " dont Réouvertures": 'wsExcel.Cells(35, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Cells(35, 2) = "Particuliers": ' wsExcel.Cells(35, 2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Cells(35, 3) = "Pers. morales": ' wsExcel.Cells(35, 3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Cells(35, 4) = "Banques": ' wsExcel.Cells(35, 4).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Cells(35, 5) = "Autres": 'wsExcel.Cells(35, 5).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Cells(35, 6) = "Total": 'wsExcel.Cells(35, 6).HorizontalAlignment = Excel.xlHAlignCenter

xRange_A = "Detail!A1:Detail!A" & mXls2_Row
xRange_B = "Detail!B1:Detail!B" & mXls2_Row
xRange_T = "Detail!T1:Detail!T" & mXls2_Row

wsExcel.Cells(36, 1) = "  Clients": wsExcel.Cells(36, 1).Font.Color = vbBlue
wsExcel.Cells(36, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";1;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(36, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";1;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(36, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";1;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(36, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";1;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(36, 6).FormulaLocal = "=SOMME(B36:E36)": wsExcel.Cells(36, 6).Interior.Color = mColor_B0
For K = 2 To 6
    If Val(wsExcel.Cells(36, K)) <> 0 Then
        wsExcel.Cells(36, K).Font.Color = vbBlue
    Else
        wsExcel.Cells(36, K) = ""
    End If
Next K


wsExcel.Cells(37, 1) = "  Techniques": wsExcel.Cells(37, 1).Font.Color = vbBlue
wsExcel.Cells(37, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";2;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(37, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";2;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(37, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";2;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(37, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";2;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(37, 6).FormulaLocal = "=SOMME(B37:E37)": wsExcel.Cells(37, 6).Interior.Color = mColor_B0
For K = 2 To 6
    If Val(wsExcel.Cells(37, K)) <> 0 Then
        wsExcel.Cells(37, K).Font.Color = vbBlue
    Else
        wsExcel.Cells(37, K) = ""
    End If
Next K

wsExcel.Cells(38, 1) = "  Tiers": wsExcel.Cells(38, 1).Font.Color = vbBlue
wsExcel.Cells(38, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";3;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(38, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";3;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(38, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";3;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(38, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";3;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(38, 6).FormulaLocal = "=SOMME(B38:E38)": wsExcel.Cells(38, 6).Interior.Color = mColor_B0
For K = 2 To 6
    If Val(wsExcel.Cells(38, K)) <> 0 Then
        wsExcel.Cells(38, K).Font.Color = vbBlue
    Else
        wsExcel.Cells(38, K) = ""
    End If
Next K


wsExcel.Cells(39, 1) = "  BIA non réclamés": wsExcel.Cells(39, 1).Font.Color = vbBlue
wsExcel.Cells(39, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";4;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(39, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";4;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(39, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";4;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(39, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";4;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(39, 6).FormulaLocal = "=SOMME(B39:E39)": wsExcel.Cells(39, 6).Interior.Color = mColor_B0
For K = 2 To 6
    If Val(wsExcel.Cells(39, K)) <> 0 Then
        wsExcel.Cells(39, K).Font.Color = vbBlue
    Else
        wsExcel.Cells(39, K) = ""
    End If
Next K

wsExcel.Cells(40, 1) = "  hors GSOP": wsExcel.Cells(40, 1).Font.Color = vbBlue
wsExcel.Cells(40, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";5;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(40, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";5;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(40, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";5;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(40, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";5;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(40, 6).FormulaLocal = "=SOMME(B40:E40)": wsExcel.Cells(40, 6).Interior.Color = mColor_B0
For K = 2 To 6
    If Val(wsExcel.Cells(40, K)) <> 0 Then
        wsExcel.Cells(40, K).Font.Color = vbBlue
    Else
        wsExcel.Cells(40, K) = ""
    End If
Next K

wsExcel.Cells(41, 1) = "  Racines sans compte": wsExcel.Cells(41, 1).Font.Color = vbBlue
wsExcel.Cells(41, 2).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";6;" & xRange_T & ";""Réouv"")" _
                                  & " + NB.SI.ENS(" & xRange_A & ";1;" & xRange_B & ";0;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(41, 3).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";6;" & xRange_T & ";""Réouv"")" _
                                  & " + NB.SI.ENS(" & xRange_A & ";2;" & xRange_B & ";0;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(41, 4).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";6;" & xRange_T & ";""Réouv"")" _
                                  & " + NB.SI.ENS(" & xRange_A & ";3;" & xRange_B & ";0;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(41, 5).FormulaLocal = "=NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";6;" & xRange_T & ";""Réouv"")" _
                                  & " + NB.SI.ENS(" & xRange_A & ";4;" & xRange_B & ";0;" & xRange_T & ";""Réouv"")"
wsExcel.Cells(41, 6).FormulaLocal = "=SOMME(B41:E41)": wsExcel.Cells(41, 6).Interior.Color = mColor_B0
For K = 2 To 6
    If Val(wsExcel.Cells(41, K)) <> 0 Then
        wsExcel.Cells(41, K).Font.Color = vbBlue
    Else
        wsExcel.Cells(41, K) = ""
    End If
Next K

wsExcel.Cells(42, 1) = "  Total": wsExcel.Cells(42, 1).Font.Color = vbBlue
wsExcel.Cells(42, 2).FormulaLocal = "=SOMME(B36:B41)": wsExcel.Cells(42, 2).Font.Color = vbBlue
wsExcel.Cells(42, 3).FormulaLocal = "=SOMME(C36:C41)": wsExcel.Cells(42, 3).Font.Color = vbBlue
wsExcel.Cells(42, 4).FormulaLocal = "=SOMME(D36:D41)": wsExcel.Cells(42, 4).Font.Color = vbBlue
wsExcel.Cells(42, 5).FormulaLocal = "=SOMME(E36:E41)": wsExcel.Cells(42, 5).Font.Color = vbBlue
wsExcel.Cells(42, 6).FormulaLocal = "=SOMME(F36:F41)": wsExcel.Cells(42, 6).Font.Color = vbBlue

mXls1_Col = 6: mXls1_Row = 1

For K = 2 To mXls1_Col
    wsExcel.Cells(35, K).Interior.Color = mColor_B0
    wsExcel.Cells(35, K).Font.Color = vbBlue
Next

For K = 36 To 41
    wsExcel.Cells(K, 1).Interior.Color = mColor_B0
    wsExcel.Cells(K, 1).Font.Color = vbBlue
Next
For K = 2 To mXls1_Col
    wsExcel.Cells(42, K).Interior.Color = mColor_B0
Next

wsExcel.Cells(35, 1).Interior.Color = mColor_B1
wsExcel.Cells(35, 1).Font.Color = vbBlack
wsExcel.Cells(35, 6).Interior.Color = mColor_B1
wsExcel.Cells(35, 6).Font.Color = vbBlack
wsExcel.Cells(42, 1).Interior.Color = mColor_B1
wsExcel.Cells(42, 1).Font.Color = vbBlack
wsExcel.Cells(42, 6).Interior.Color = mColor_B1
wsExcel.Cells(42, 6).Font.Color = vbBlack

'__________________________________________________________________________________
Exit_sub:
'=======================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée "): DoEvents

End Sub


Public Sub cmdSelect_SQL_Xgsop_Total_K3()
On Error GoTo Error_Handler
Dim wRow As Long, wCol As Long
Dim X As String, Kx As Long, Ky As Long
Dim xSQL As String, wFile As String, xRange_A As String, xRange_B As String, xRange_T As String

On Error GoTo Error_Handler

If Not blnAuto Then Exit Sub

wRow = 45

wsExcel.Cells(wRow, 1) = " dont état ancien": 'wsExcel.Cells(wrow, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Cells(wRow, 2) = "1-Particuliers": ' wsExcel.Cells(wrow, 2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Cells(wRow, 3) = "2-Pers. morales": ' wsExcel.Cells(wrow, 3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Cells(wRow, 4) = "3-Banques": ' wsExcel.Cells(wrow, 4).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Cells(wRow, 5) = "4-Autres": 'wsExcel.Cells(wrow, 5).HorizontalAlignment = Excel.xlHAlignCenter
'wsExcel.Cells(wRow, 6) = "Total": 'wsExcel.Cells(wrow, 6).HorizontalAlignment = Excel.xlHAlignCenter
For Kx = 1 To 5
    wsExcel.Cells(wRow, Kx).Interior.Color = mColor_Y2
Next Kx

For Ky = 1 To 6
    wRow = wRow + 1
    wsExcel.Cells(wRow, 1) = arrY_Lib(Ky)
    wsExcel.Cells(wRow, 1).Interior.Color = mColor_Y2
    For Kx = 1 To 4
        wsExcel.Cells(wRow, Kx + 1) = arrK3_Old(Kx, Ky)
        wsExcel.Cells(wRow, Kx + 1).Font.Color = vbRed
    Next Kx
Next Ky

wRow = 53

wsExcel.Cells(wRow, 1) = " dont état nouveau": 'wsExcel.Cells(wrow, 1).HorizontalAlignment = Excel.xlHAlignLeft
wsExcel.Cells(wRow, 2) = "1-Particuliers": ' wsExcel.Cells(wrow, 2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Cells(wRow, 3) = "2-Pers. morales": ' wsExcel.Cells(wrow, 3).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Cells(wRow, 4) = "3-Banques": ' wsExcel.Cells(wrow, 4).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Cells(wRow, 5) = "4-Autres": 'wsExcel.Cells(wrow, 5).HorizontalAlignment = Excel.xlHAlignCenter
'wsExcel.Cells(wRow, 6) = "Total": 'wsExcel.Cells(wrow, 6).HorizontalAlignment = Excel.xlHAlignCenter
For Kx = 1 To 5
    wsExcel.Cells(wRow, Kx).Interior.Color = mColor_Y2
Next Kx

For Ky = 1 To 6
    wRow = wRow + 1
    wsExcel.Cells(wRow, 1) = arrY_Lib(Ky)
    wsExcel.Cells(wRow, 1).Interior.Color = mColor_Y2
    For Kx = 1 To 4
        wsExcel.Cells(wRow, 1) = arrY_Lib(Ky)
        wsExcel.Cells(wRow, Kx + 1) = arrK3_New(Kx, Ky)
        wsExcel.Cells(wRow, Kx + 1).Font.Color = vbBlue
    Next Kx
Next Ky


'__________________________________________________________________________________
Exit_sub:
'=======================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée "): DoEvents

End Sub





Public Sub paramKYCDOSECH_Reprise()
    
Dim V, X As String, X1 As String

App_Debug = "paramKYCDOSECH_Reprise"
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox

X = "select * from " & paramIBM_Library_SABSPE & ".YKYCDOS0 " _
  & " where KYCDOSNAT = ' ' and KYCDOSDAMJ > 0 and KYCDOSSEQ2 = 90 " _
  & " order by KYCDOSID"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    Call rsYKYCDOS0_GetBuffer(rsSab, oldYKYCDOS0)
    newYKYCDOS0 = oldYKYCDOS0
    X = oldYKYCDOS0.KYCDOSDAMJ
    
    X1 = dateElp("AnAdd", 4, X)
    X = dateElp("Jour", -1, X1)
    newYKYCDOS0.KYCDOSDECH = Val(X)
    
    V = sqlYKYCDOS0_Update(newYKYCDOS0, oldYKYCDOS0, False)
    
    If Not IsNull(V) Then GoTo Error_MsgBox

    rsSab.MoveNext
Loop

'________________________________________________________________________________
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
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    

End Sub

Public Sub cboSelect_Options_KYCech_Doc_Load()

Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler

cboSelect_Options_KYCech_Doc.Clear
cboSelect_Options_KYCech_Doc.AddItem ""
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YKYCDOS0" _
     & " where KYCDOSNAT = 'D' and KYCDOSDECH > 0" _
     & " order by KYCDOSSEQ"

Set rsAdo = cnAdo.Execute(xSQL)

Do While Not rsAdo.EOF
    K = rsAdo("KYCDOSSEQ")
    cboSelect_Options_KYCech_Doc.AddItem Format(K, "####") & " : " & Trim(rsAdo("KYCDOSDLIB"))
    rsAdo.MoveNext

Loop
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cboSelect_Options_KYCech_Doc_Load"


End Sub
