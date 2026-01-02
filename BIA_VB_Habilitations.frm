VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBIA_VB_Habilitations 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA_VB_Habilitations"
   ClientHeight    =   10305
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BIA_VB_Habilitations.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10305
   ScaleWidth      =   13530
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   6120
      TabIndex        =   2
      Top             =   0
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9720
      Left            =   0
      TabIndex        =   3
      Top             =   495
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   17145
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
      TabCaption(0)   =   "Gestion des habilitations"
      TabPicture(0)   =   "BIA_VB_Habilitations.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Affectation utilisateurs / services"
      TabPicture(1)   =   "BIA_VB_Habilitations.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraUsr_Srv"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "."
      TabPicture(2)   =   "BIA_VB_Habilitations.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtRTF"
      Tab(2).Control(1)=   "fraSelect_4_Options"
      Tab(2).Control(2)=   "txtFg"
      Tab(2).Control(3)=   "fraParam_Mnu"
      Tab(2).Control(4)=   "lstW"
      Tab(2).ControlCount=   5
      Begin VB.ListBox lstW 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   -64695
         Sorted          =   -1  'True
         TabIndex        =   70
         Top             =   555
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Frame fraParam_Mnu 
         BackColor       =   &H00E0E0E0&
         Caption         =   "fraParam_Mnu"
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
         Left            =   -69840
         TabIndex        =   60
         Top             =   2100
         Visible         =   0   'False
         Width           =   7500
         Begin VB.TextBox txtParam_Mnu_Lib 
            Height          =   660
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   67
            Text            =   "BIA_VB_Habilitations.frx":035E
            Top             =   1170
            Width           =   6900
         End
         Begin VB.TextBox txtParam_Mnu_Code 
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
            Left            =   1830
            MaxLength       =   12
            TabIndex        =   66
            Top             =   510
            Width           =   1710
         End
         Begin VB.CommandButton cmdParam_Mnu_Delete 
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
            Left            =   1995
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   6435
            Width           =   900
         End
         Begin VB.CommandButton cmdParam_Mnu_Add 
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
            TabIndex        =   64
            Top             =   6435
            Width           =   900
         End
         Begin VB.CommandButton cmdParam_Mnu_Update 
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
            Left            =   5865
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   6435
            Width           =   900
         End
         Begin VB.CommandButton cmdParam_Mnu_Quit 
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
            Left            =   285
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   6435
            Width           =   990
         End
         Begin VB.ListBox lstParam_Mnu_Droit 
            BackColor       =   &H00D0FFD0&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3885
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   61
            Top             =   2145
            Width           =   6900
         End
         Begin VB.Label lblParam_Mnu_Quid 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblParam_quid"
            Height          =   330
            Left            =   255
            TabIndex        =   69
            Top             =   7170
            Width           =   6825
         End
         Begin VB.Label libParam_Mnu_Code 
            BackColor       =   &H00404040&
            Caption         =   "Code (12 car)"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   60
            TabIndex        =   68
            Top             =   510
            Width           =   1725
         End
      End
      Begin VB.Frame fraUsr_Srv 
         Height          =   9300
         Left            =   -74955
         TabIndex        =   45
         Top             =   360
         Visible         =   0   'False
         Width           =   13170
         Begin VB.ListBox lstUsr_VB_No 
            BackColor       =   &H00FFE0FF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2580
            Left            =   3825
            Sorted          =   -1  'True
            TabIndex        =   78
            Top             =   3870
            Width           =   3045
         End
         Begin VB.ListBox lstUsr_VB 
            BackColor       =   &H00E0EFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2580
            Left            =   3810
            Sorted          =   -1  'True
            TabIndex        =   73
            Top             =   510
            Width           =   3045
         End
         Begin VB.Frame fraUsr 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Affectation Utilisateur => Service"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2200
            Left            =   7140
            TabIndex        =   49
            Top             =   6900
            Visible         =   0   'False
            Width           =   5850
            Begin VB.TextBox txtUsr_Code 
               Height          =   315
               Left            =   240
               MaxLength       =   12
               TabIndex        =   75
               Top             =   450
               Width           =   1560
            End
            Begin VB.CommandButton cmdUsr_Quit 
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
               Left            =   4635
               Style           =   1  'Graphical
               TabIndex        =   59
               Top             =   390
               Width           =   990
            End
            Begin VB.ComboBox cboUsr_Srv 
               Height          =   330
               Left            =   240
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   58
               Top             =   1080
               Width           =   2295
            End
            Begin VB.Label libUsr_Code 
               BackColor       =   &H00F0FFFF&
               Caption         =   "code"
               Height          =   300
               Left            =   240
               TabIndex        =   57
               Top             =   1680
               Width           =   5385
            End
         End
         Begin VB.Frame fraSrv 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Gestion des services"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2200
            Left            =   105
            TabIndex        =   48
            Top             =   6900
            Visible         =   0   'False
            Width           =   5775
            Begin VB.CommandButton cmdSrv_Quit 
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
               Left            =   4620
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   555
               Width           =   990
            End
            Begin VB.TextBox txtSrv_Lib2 
               Height          =   690
               Left            =   1125
               MaxLength       =   99
               TabIndex        =   55
               Top             =   1470
               Width           =   3990
            End
            Begin VB.TextBox txtSrv_Lib1 
               Height          =   315
               Left            =   1125
               MaxLength       =   12
               TabIndex        =   54
               Top             =   1035
               Width           =   1875
            End
            Begin VB.TextBox txtSrv_Code 
               Height          =   315
               Left            =   1140
               MaxLength       =   2
               TabIndex        =   53
               Top             =   510
               Width           =   600
            End
            Begin VB.Label lblSrv_Lib2 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Libellé"
               Height          =   255
               Left            =   120
               TabIndex        =   52
               Top             =   1485
               Width           =   915
            End
            Begin VB.Label lblSrv_Lib1 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Abrégé"
               Height          =   255
               Left            =   120
               TabIndex        =   51
               Top             =   1035
               Width           =   915
            End
            Begin VB.Label lblSrv_Code 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Code S**"
               Height          =   255
               Left            =   120
               TabIndex        =   50
               Top             =   615
               Width           =   915
            End
         End
         Begin VB.ListBox lstSrv 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5940
            Left            =   105
            Sorted          =   -1  'True
            TabIndex        =   47
            Top             =   525
            Width           =   3450
         End
         Begin VB.ListBox lstUsr 
            BackColor       =   &H00F0FFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5940
            Left            =   7170
            Sorted          =   -1  'True
            TabIndex        =   46
            Top             =   555
            Width           =   5805
         End
         Begin VB.Label lblUsr_VB_No 
            Caption         =   "Comptes NON WIN, actifs dans SAB"
            Height          =   270
            Left            =   3780
            TabIndex        =   77
            Top             =   3465
            Width           =   3300
         End
         Begin VB.Label lblUsr_VB 
            Caption         =   "Comptes WIN, NON actifs dans SAB"
            Height          =   270
            Left            =   3705
            TabIndex        =   74
            Top             =   210
            Width           =   3315
         End
         Begin VB.Label lblSrv 
            Caption         =   "Service (ROPDOSISRV)"
            Height          =   270
            Left            =   390
            TabIndex        =   72
            Top             =   210
            Width           =   2130
         End
         Begin VB.Label lblUsr 
            Caption         =   "Comptes WIN, actifs dans SAB"
            Height          =   270
            Left            =   9000
            TabIndex        =   71
            Top             =   225
            Width           =   3240
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
         Left            =   -74400
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   43
         Text            =   "BIA_VB_Habilitations.frx":0364
         Top             =   3135
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Frame fraSelect_4_Options 
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
         Height          =   945
         Left            =   -74400
         TabIndex        =   38
         Top             =   1005
         Visible         =   0   'False
         Width           =   8300
         Begin VB.ComboBox cboSelect_4_APP 
            Height          =   330
            Left            =   495
            Sorted          =   -1  'True
            TabIndex        =   76
            Text            =   "cboSelect_App"
            Top             =   495
            Width           =   1860
         End
         Begin VB.ComboBox cboSelect_4_Usr1 
            Height          =   330
            Left            =   3405
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   540
            Width           =   1860
         End
         Begin VB.ComboBox cboSelect_4_Usr2 
            Height          =   330
            Left            =   6045
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   495
            Width           =   1860
         End
         Begin VB.Label lblSelect_4_Usr1 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Duplication des habilitations de "
            Height          =   270
            Left            =   1215
            TabIndex        =   42
            Top             =   165
            Width           =   3600
         End
         Begin VB.Label lblSelect_4_Usr2 
            BackColor       =   &H00F0FFFF&
            Caption         =   "vers"
            Height          =   270
            Left            =   6600
            TabIndex        =   41
            Top             =   135
            Width           =   705
         End
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
         Height          =   9630
         Left            =   0
         TabIndex        =   4
         Top             =   495
         Width           =   13425
         Begin VB.Frame fraParam_Hab 
            BackColor       =   &H00F0FFFF&
            Caption         =   "APP / USR"
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
            Left            =   8385
            TabIndex        =   26
            Top             =   1845
            Visible         =   0   'False
            Width           =   8970
            Begin VB.CommandButton cmdParam_Hab_Quit 
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
               Height          =   645
               Left            =   780
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   6510
               Width           =   1380
            End
            Begin VB.CommandButton cmdParam_Hab_Update 
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
               Height          =   585
               Left            =   6585
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   6525
               Width           =   1335
            End
            Begin VB.ListBox lstParam_Hab 
               BackColor       =   &H00D0FFD0&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5160
               Left            =   390
               Style           =   1  'Checkbox
               TabIndex        =   27
               Top             =   420
               Width           =   8250
            End
            Begin VB.Label lblParam_Hab_Quid 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblParam_Hab_Quid"
               Height          =   330
               Left            =   345
               TabIndex        =   37
               Top             =   7290
               Width           =   8160
            End
            Begin VB.Label libParam_Hab_Srv 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               Caption         =   "Habilitation accordée pour l'utilisateur uniquement pour son service actuel."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   450
               Left            =   420
               TabIndex        =   35
               Top             =   5865
               Visible         =   0   'False
               Width           =   8115
            End
         End
         Begin VB.Frame fraParam_App 
            BackColor       =   &H00E0E0E0&
            Caption         =   "fraParam_App"
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
            Height          =   4230
            Left            =   195
            TabIndex        =   12
            Top             =   2955
            Visible         =   0   'False
            Width           =   7320
            Begin VB.CommandButton cmdParam_App_Add_19 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Ajouter 19 = Admin"
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
               Left            =   6015
               Style           =   1  'Graphical
               TabIndex        =   83
               Top             =   1850
               Width           =   950
            End
            Begin VB.CommandButton cmdParam_App_Add_18 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Ajouter 18 = Param"
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
               Left            =   4935
               Style           =   1  'Graphical
               TabIndex        =   82
               Top             =   1850
               Width           =   900
            End
            Begin VB.CommandButton cmdParam_App_Add_17 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Ajouter 17 =  Doc"
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
               Left            =   3720
               Style           =   1  'Graphical
               TabIndex        =   81
               Top             =   1850
               Width           =   900
            End
            Begin VB.TextBox txtParam_APP_Doc 
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
               Left            =   1830
               MaxLength       =   10
               TabIndex        =   80
               Top             =   1950
               Width           =   1710
            End
            Begin VB.OptionButton optParam_App_Z 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "sans restriction"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   3700
               TabIndex        =   34
               Top             =   270
               Width           =   3200
            End
            Begin VB.OptionButton optParam_App_S 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "ou aux utilisateurs d'un service"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   3700
               TabIndex        =   33
               Top             =   1125
               Width           =   3200
            End
            Begin VB.OptionButton optParam_App_X 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Caption         =   "réservé aux administrateurs"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   3700
               TabIndex        =   32
               Top             =   690
               Width           =   3200
            End
            Begin VB.TextBox txtParam_APP_Seq 
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
               Left            =   1830
               MaxLength       =   2
               TabIndex        =   23
               Top             =   840
               Width           =   600
            End
            Begin VB.TextBox txtParam_App_Lib 
               Height          =   525
               Left            =   180
               MultiLine       =   -1  'True
               TabIndex        =   19
               Text            =   "BIA_VB_Habilitations.frx":036C
               Top             =   2370
               Width           =   6840
            End
            Begin VB.TextBox txtParam_APP_Code 
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
               Left            =   1830
               MaxLength       =   12
               TabIndex        =   18
               Top             =   270
               Width           =   1710
            End
            Begin VB.CommandButton cmdParam_APP_Delete 
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
               Left            =   2130
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   3015
               Width           =   900
            End
            Begin VB.CommandButton cmdParam_App_Add 
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
               Left            =   4125
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   3000
               Width           =   900
            End
            Begin VB.CommandButton cmdParam_App_Update 
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
               Left            =   6060
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   3030
               Width           =   900
            End
            Begin VB.CommandButton cmdParam_App_Quit 
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
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   3000
               Width           =   990
            End
            Begin VB.ComboBox cboParam_APP_VBP 
               Height          =   330
               Left            =   1830
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   1455
               Width           =   2295
            End
            Begin VB.Label lblParam_APP_Doc 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Id documentation"
               Height          =   255
               Left            =   180
               TabIndex        =   79
               Top             =   2000
               Width           =   1560
            End
            Begin VB.Label lblParam_quid 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblParam_quid"
               Height          =   330
               Left            =   180
               TabIndex        =   36
               Top             =   3750
               Width           =   6810
            End
            Begin VB.Label lblParam_APP_Seq 
               BackColor       =   &H00E0E0E0&
               Caption         =   "séquence"
               Height          =   255
               Left            =   180
               TabIndex        =   22
               Top             =   960
               Width           =   1245
            End
            Begin VB.Label lblParam_APP_Code 
               BackColor       =   &H00404040&
               Caption         =   "Code (12 car)"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   75
               TabIndex        =   21
               Top             =   315
               Width           =   1725
            End
            Begin VB.Label lblParam_APP_VBP 
               BackColor       =   &H00E0E0E0&
               Caption         =   "VBP"
               Height          =   255
               Left            =   180
               TabIndex        =   20
               Top             =   1500
               Width           =   1245
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   4000
            Left            =   105
            TabIndex        =   11
            Top             =   5100
            Visible         =   0   'False
            Width           =   13200
            _ExtentX        =   23283
            _ExtentY        =   7064
            _Version        =   393216
            Cols            =   4
            RowHeightMin    =   300
            BackColor       =   15790320
            ForeColor       =   4210752
            BackColorFixed  =   8438015
            ForeColorFixed  =   0
            BackColorBkg    =   15790320
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   $"BIA_VB_Habilitations.frx":0372
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
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   555
            Left            =   11895
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox cboSelect_SQL 
            Height          =   330
            ItemData        =   "BIA_VB_Habilitations.frx":046C
            Left            =   8625
            List            =   "BIA_VB_Habilitations.frx":046E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   210
            Width           =   4620
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
            Height          =   945
            Left            =   120
            TabIndex        =   5
            Top             =   105
            Visible         =   0   'False
            Width           =   8300
            Begin VB.ComboBox cboSelect_Usr 
               Height          =   330
               Left            =   5775
               Sorted          =   -1  'True
               TabIndex        =   31
               Text            =   "cboSelect_Usr"
               Top             =   570
               Width           =   1860
            End
            Begin VB.ComboBox cboSelect_Srv 
               Height          =   330
               Left            =   2865
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   525
               Width           =   1860
            End
            Begin VB.ComboBox cboSelect_App 
               Height          =   330
               Left            =   180
               Sorted          =   -1  'True
               TabIndex        =   10
               Text            =   "cboSelect_App"
               Top             =   540
               Width           =   1860
            End
            Begin VB.Label lblSelect_Usr 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Utilisateur%"
               Height          =   270
               Left            =   5910
               TabIndex        =   30
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label lblSelect_Srv 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Service"
               Height          =   270
               Left            =   3195
               TabIndex        =   24
               Top             =   210
               Width           =   705
            End
            Begin VB.Label lblSelect_App 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Application%"
               Height          =   270
               Left            =   345
               TabIndex        =   9
               Top             =   210
               Width           =   1155
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   3900
            Left            =   105
            TabIndex        =   8
            Top             =   1100
            Visible         =   0   'False
            Width           =   13200
            _ExtentX        =   23283
            _ExtentY        =   6879
            _Version        =   393216
            Cols            =   4
            RowHeightMin    =   300
            BackColor       =   16777215
            ForeColor       =   16711680
            BackColorFixed  =   12632064
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   $"BIA_VB_Habilitations.frx":0470
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
      Begin RichTextLib.RichTextBox txtRTF 
         Height          =   3360
         Left            =   -73935
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   5775
         Visible         =   0   'False
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   5927
         _Version        =   393217
         BackColor       =   15790320
         HideSelection   =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"BIA_VB_Habilitations.frx":0567
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
      Picture         =   "BIA_VB_Habilitations.frx":05E7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
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
   End
End
Attribute VB_Name = "frmBIA_VB_Habilitations"
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
'''Dim BIA_VB_Habilitations_Aut As typeAuthorization
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

'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long

Dim HeightOfLine As Long, LinesOfText As Long

Dim txtRTF_prtForeColor_Header As Long


Dim Old_YBIATAB0 As typeYBIATAB0, New_YBIATAB0 As typeYBIATAB0, X_YBIATAB0 As typeYBIATAB0
Dim App_YBIATAB0 As typeYBIATAB0, Fct_YBIATAB0 As typeYBIATAB0
Dim arrBIA_VB_HAB() As typeYBIATAB0, arrBIA_VB_HAB_Nb As Integer
Dim arrBIA_VB_MNU() As typeYBIATAB0, arrBIA_VB_Mnu_Nb As Integer, arrBIA_VB_Mnu_K As Integer
Dim mSelect_App As String, mSelect_Droit As String, mSelect_Droit_X As String, mSelect_Droit_Seq As Integer
Dim mSelect_Usr As String, mSelect_Usr_Srv As String, arrFct_X(100) As String, arrFct_ListIndex(100) As Integer
Dim blnAdmin As Boolean

'______________________________________________________________________

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim mXls2_Cols As Integer, mXls2_Row As Integer
Dim rsSabX As New ADODB.Recordset
Dim arrFCT(100) As Integer, arrUsr() As String, arrUsr_Nb As Integer, arrUsr_K As Integer
Dim arrMNU_Row(100) As Integer, arrMNU_Hab(100) As String, arrMNU_Nb As Integer

Dim blnParam_Hab_Change As Boolean
Public Sub cmdSelect_SQL_Exportation()

On Error GoTo Error_Handler
Dim X As String, K As Long, K2 As Long, xSQL As String, xWhere As String, wNum As Long
Dim wFile As String, wFilex As String
Dim blnCALCS As Boolean
Dim xLib As String
On Error GoTo Error_Handler
'===================================================================================
'If blnAuto Then
'    X = paramServer("\\CPT_Archive\")
'Else
    X = ""
'End If
If X = "" Then X = "C:\Temp\"
If mId$(X, Len(X), 1) <> "\" Then X = X & "\"

blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

xLib = "Habililations VB "

wFile = X & xLib & " " & DSYS_Time & ".xlsx"

'If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "Habililations VB : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
'End If

If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile
'_________________________________________


Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "BIA_VB_HAB"
    .Subject = ""
End With

xWhere = ""
X = Trim(mId$(cboSelect_Srv, 1, 3))
If X <> "" Then xWhere = " and substring(BIATABTXT,100,3) = '" & X & "'"
X = Trim(cboSelect_Usr)
If X <> "" Then xWhere = xWhere & " and BIATABK2 like '" & X & "%'"
X = Trim(cboSelect_App)
If X <> "" Then xWhere = xWhere & " and BIATABK1 like '" & X & "%'"
'===================================================================================================
Call cmdSelect_SQL_Exportation_Page1(1, "Habilitations")
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB'" & xWhere & " order by BIATABK1, BIATABK2"
Set rsSab = cnsab.Execute(xSQL)
Call cmdSelect_SQL_Exportation_Habilitations


'===================================================================================================
xSQL = "select count(distinct  BIATABK2) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB'" & xWhere
Set rsSab = cnsab.Execute(xSQL)

ReDim arrUsr(rsSab(0) + 1)
ReDim arrBIA_VB_MNU(500)

arrUsr_Nb = 0
xSQL = "select distinct  BIATABK2 from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB'" & xWhere & " order by BIATABK2"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrUsr_Nb = arrUsr_Nb + 1
    arrUsr(arrUsr_Nb) = Trim(rsSab("BIATABK2"))
    rsSab.MoveNext
Loop

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB'" & xWhere & " order by BIATABK1"
Set rsSab = cnsab.Execute(xSQL)

Call cmdSelect_SQL_Exportation_Page3(3, "Utilisateurs")
Call cmdSelect_SQL_Exportation_Detail

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB'" & xWhere & " order by BIATABK1"
Set rsSab = cnsab.Execute(xSQL)
Call cmdSelect_SQL_Exportation_Page2(2, "Droits-Menus")
Call cmdSelect_SQL_Exportation_Mnu


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
Call lstErr_AddItem(lstErr, frmElp.cmdContext, "< Exportation terminée"): DoEvents
'_____________________________
Exit Sub

Error_Handler:

If Not blnCALCS Then
    X = "C:\Temp\"
    Resume Next
End If
MsgBox Error, vbCritical, "cmdSelect_SQL_Exportation"
Call lstErr_AddItem(lstErr, frmElp.cmdContext, "< Exportation terminée"): DoEvents

End Sub




Public Sub cmdSelect_SQL_Exportation_Detail()
Dim X As String, K As Integer, K2 As Integer, mCol As Integer, xSQL As String, blnMnu_Ok As Boolean
Dim APP_Row As Integer
On Error GoTo Error_Handler
'==========================================================================================================
Call rsYBIATAB0_Init(X_YBIATAB0)
arrUsr_K = 0: mCol = 4
'_________________________________________________________________________________________________


Do While Not rsSab.EOF

    Call rsYBIATAB0.rsYBIATAB0_GetBuffer(rsSab, Old_YBIATAB0)
    
    If Old_YBIATAB0.BIATABK1 <> X_YBIATAB0.BIATABK1 Then
'==========================================================================================================
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
             & " where BIATABID = 'BIA_VB_APP' and BIATABK1 = '" & Old_YBIATAB0.BIATABK1 & "'"
        Set rsSabX = cnsab.Execute(xSQL)
        Call rsYBIATAB0.rsYBIATAB0_GetBuffer(rsSabX, X_YBIATAB0)
        mXls2_Row = mXls2_Row + 1
        APP_Row = mXls2_Row
        If (mXls2_Row Mod 10) Then Call lstErr_ChangeLastItem(lstErr, frmElp.cmdContext, "Utilisateurs :" & Old_YBIATAB0.BIATABK1): DoEvents

        If mId$(X_YBIATAB0.BIATABK1, 1, 1) = "=" Then
            wsExcel.Cells(mXls2_Row, 1) = "* " & X_YBIATAB0.BIATABK1
        Else
            wsExcel.Cells(mXls2_Row, 1) = X_YBIATAB0.BIATABK1
        End If
        wsExcel.Cells(mXls2_Row, 2) = mId$(X_YBIATAB0.BIATABTXT, 100, 1)
        wsExcel.Cells(mXls2_Row, 4) = Trim(mId$(X_YBIATAB0.BIATABTXT, 1, 69))
        For K = 1 To mXls2_Cols
            wsExcel.Cells(mXls2_Row, K).Interior.Color = RGB(255, 255, 96)
        Next K
'_________________________________________________________________________________________________
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
             & " where BIATABID = 'BIA_VB_DROIT' and BIATABK1 = '" & Old_YBIATAB0.BIATABK1 & "' order by substring(BIATABTXT,101,2)"
        
        For K = 0 To 19: arrFCT(K) = 0: Next K
        
        Set rsSabX = cnsab.Execute(xSQL)
        Do While Not rsSabX.EOF
            mXls2_Row = mXls2_Row + 1
            wsExcel.Cells(mXls2_Row, 3) = rsSabX("BIATABK2")
            X = rsSabX("BIATABTXT")
            wsExcel.Cells(mXls2_Row, 4) = Trim(mId$(X, 1, 69))
            wsExcel.Cells(mXls2_Row, 2) = mId$(X, 100, 3)
            If mId$(X, 100, 1) <> "*" Then
                wsExcel.Cells(mXls2_Row, 2).Interior.Color = RGB(255, 255, 96)
            Else
                wsExcel.Cells(mXls2_Row, 2).Interior.Color = mColor_Y1
            End If
            arrFCT(mId$(X, 101, 2)) = mXls2_Row
            rsSabX.MoveNext
        Loop
'_________________________________________________________________________________________________
        arrMNU_Nb = 0
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
             & " where BIATABID = 'BIA_VB_MNU' and BIATABK1 = '" & Old_YBIATAB0.BIATABK1 & "' order by BIATABK2"
                
        Set rsSabX = cnsab.Execute(xSQL)
        
        arrBIA_VB_Mnu_Nb = 0
        lstW.Clear
        Do While Not rsSabX.EOF
            arrBIA_VB_Mnu_Nb = arrBIA_VB_Mnu_Nb + 1
            Call rsYBIATAB0_GetBuffer(rsSabX, arrBIA_VB_MNU(arrBIA_VB_Mnu_Nb))
            lstW.AddItem arrBIA_VB_MNU(arrBIA_VB_Mnu_Nb).BIATABK2 & "|" & arrBIA_VB_Mnu_Nb
            rsSabX.MoveNext
        Loop
        
        arrMNU_Nb = 0
        For K = 0 To arrBIA_VB_Mnu_Nb - 1
            lstW.ListIndex = K
            K2 = InStr(lstW.Text, "|")
            If K2 > 0 Then arrBIA_VB_Mnu_K = Val(mId$(lstW.Text, K2 + 1, Len(lstW.Text) - K2))
            mXls2_Row = mXls2_Row + 1
            wsExcel.Cells(mXls2_Row, 3) = Trim(arrBIA_VB_MNU(arrBIA_VB_Mnu_K).BIATABK2)
            X = arrBIA_VB_MNU(arrBIA_VB_Mnu_K).BIATABTXT
            wsExcel.Cells(mXls2_Row, 4) = Trim(mId$(X, 20, 79))
            wsExcel.Cells(mXls2_Row, 2).Interior.Color = mColor_G1
            
            arrMNU_Nb = arrMNU_Nb + 1
            arrMNU_Hab(arrMNU_Nb) = mId$(X, 1, 19)
            arrMNU_Row(arrMNU_Nb) = mXls2_Row

        Next K

'==========================================================================================================
    End If
    
    
    If Old_YBIATAB0.BIATABK2 <> arrUsr(arrUsr_K) Then
        For arrUsr_K = 1 To arrUsr_Nb
                If Trim(Old_YBIATAB0.BIATABK2) = arrUsr(arrUsr_K) Then
                    mCol = arrUsr_K + 4
                    mSelect_Usr = arrUsr(arrUsr_K)
                    lstParam_Usr_Srv
                    Exit For
                End If
        Next arrUsr_K
        
    End If
    
    For K = 1 To 19
        X = mId$(Old_YBIATAB0.BIATABTXT, K, 1)
        If X <> " " Then
            If arrFCT(K) > 0 Then
                 wsExcel.Cells(arrFCT(K), mCol) = X
                 wsExcel.Cells(arrFCT(K), mCol).Interior.Color = RGB(255, 255, 96)
                 wsExcel.Cells(arrFCT(K), mCol).Font.Bold = True
                 If mSelect_Usr_Srv <> mId$(Old_YBIATAB0.BIATABTXT, 100, 3) Then
                     wsExcel.Cells(arrFCT(K), mCol).Interior.Color = mColor_W1
                 Else
                  If mId$(Old_YBIATAB0.BIATABTXT, 100, 3) = "S99" Then wsExcel.Cells(arrFCT(K), mCol).Interior.Color = mColor_W0
                End If
            Else
                 wsExcel.Cells(APP_Row, mCol).Interior.Color = vbRed
                 wsExcel.Cells(APP_Row, 3) = "Droits # Hab"
                 wsExcel.Cells(APP_Row, 3).Interior.Color = vbRed
                 wsExcel.Cells(APP_Row, 3).Font.Color = vbYellow
            End If
            
        End If
    Next K
'_________________________________________________________________________________________________
    For K = 1 To arrMNU_Nb
        blnMnu_Ok = True
        X = arrMNU_Hab(K)
        For K2 = 1 To 19
            If mId$(X, K2, 1) <> " " Then
                If mId$(Old_YBIATAB0.BIATABTXT, K2, 1) = " " Then blnMnu_Ok = False
            End If
        Next K2
        If blnMnu_Ok Then
            wsExcel.Cells(arrMNU_Row(K), mCol) = "@"
            wsExcel.Cells(arrMNU_Row(K), mCol).Font.Bold = True
            wsExcel.Cells(arrMNU_Row(K), mCol).Font.Color = RGB(0, 96, 0)
            'wsExcel.Cells(arrMNU_Row(K), mCol).Interior.Color = mColor_G1
        End If
        
    Next K

'_________________________________________________________________________________________________
    rsSab.MoveNext
Loop

'==========================================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdSelect_SQL_Exportation_Detail"


End Sub
Public Sub cmdSelect_SQL_Exportation_Habilitations()
Dim X As String, K As Integer, K2 As Integer, mCol As Integer, xSQL As String, blnOk As Boolean
Dim APP_Row As Integer, arrDroits(20) As String
Dim arrUsr_X(999) As String, arrUsr_Color(999) As Long, arrUsr_Nb As Integer, arrUsr_K As Integer
Dim wColor As Long
On Error GoTo Error_Handler
'==========================================================================================================
For K = 1 To 994 Step 5
    arrUsr_Color(K) = mColor_G1
    arrUsr_Color(K + 1) = mColor_Y1
    arrUsr_Color(K + 2) = mColor_B0
    arrUsr_Color(K + 3) = mColor_G0
    arrUsr_Color(K + 4) = mColor_Y0
    
Next K
Do While Not rsSab.EOF

    Call rsYBIATAB0.rsYBIATAB0_GetBuffer(rsSab, Old_YBIATAB0)
    blnOk = False
    For arrUsr_K = 1 To arrUsr_Nb
        If Old_YBIATAB0.BIATABK2 = arrUsr_X(arrUsr_K) Then blnOk = True: Exit For
    Next arrUsr_K
    If Not blnOk Then
        arrUsr_Nb = arrUsr_Nb + 1
        arrUsr_X(arrUsr_Nb) = Old_YBIATAB0.BIATABK2
        arrUsr_K = arrUsr_Nb
    End If
    wColor = arrUsr_Color(arrUsr_K)
    mXls2_Row = mXls2_Row + 1
    
    wsExcel.Cells(mXls2_Row, 3).Interior.Color = wColor
    wsExcel.Cells(mXls2_Row, 4).Interior.Color = wColor

    If Old_YBIATAB0.BIATABK1 <> X_YBIATAB0.BIATABK1 Then
'==========================================================================================================
        
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
             & " where BIATABID = 'BIA_VB_APP' and BIATABK1 = '" & Old_YBIATAB0.BIATABK1 & "'"
        Set rsSabX = cnsab.Execute(xSQL)
        Call rsYBIATAB0.rsYBIATAB0_GetBuffer(rsSabX, X_YBIATAB0)
        APP_Row = mXls2_Row
        If (mXls2_Row Mod 10) Then Call lstErr_ChangeLastItem(lstErr, frmElp.cmdContext, "Utilisateurs :" & Old_YBIATAB0.BIATABK1): DoEvents

        If mId$(X_YBIATAB0.BIATABK1, 1, 1) = "=" Then
            wsExcel.Cells(mXls2_Row, 1) = "* " & X_YBIATAB0.BIATABK1
        Else
            wsExcel.Cells(mXls2_Row, 1) = X_YBIATAB0.BIATABK1
        End If
        wsExcel.Cells(mXls2_Row, 2) = Trim(mId$(X_YBIATAB0.BIATABTXT, 1, 69))
        'For K = 1 To mXls2_Cols
        '    wsExcel.Cells(mXls2_Row, K).Interior.Color = RGB(255, 255, 96)
        'Next K
'_________________________________________________________________________________________________
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
             & " where BIATABID = 'BIA_VB_DROIT' and BIATABK1 = '" & Old_YBIATAB0.BIATABK1 & "'" ' order by substring(BIATABTXT,101,2)"
        
        For K = 1 To 19: arrDroits(K) = "": Next K
        
        Set rsSabX = cnsab.Execute(xSQL)
        Do While Not rsSabX.EOF
            K = Val(mId$(rsSabX("BIATABTXT"), 101, 2))
            arrDroits(K) = rsSabX("BIATABK2")
            rsSabX.MoveNext
        Loop
        arrDroits(1) = Old_YBIATAB0.BIATABK1
'_________________________________________________________________________________________________

'==========================================================================================================
    End If
    
    wsExcel.Cells(mXls2_Row, 3) = Old_YBIATAB0.BIATABK2
    wsExcel.Cells(mXls2_Row, 4) = mId$(Old_YBIATAB0.BIATABTXT, 100, 3)
    For K = 1 To 19
        If mId$(Old_YBIATAB0.BIATABTXT, K, 1) <> " " Then
            wsExcel.Cells(mXls2_Row, 4 + K) = arrDroits(K)
            wsExcel.Cells(mXls2_Row, 4 + K).Interior.Color = wColor
        End If
    Next K
'_________________________________________________________________________________________________
    rsSab.MoveNext
Loop

'==========================================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdSelect_SQL_Exportation_Detail"


End Sub




Public Sub cmdSelect_SQL_Exportation_Mnu()
Dim X As String, K As Integer, K2 As Integer, mCol As Integer, xSQL As String, blnMnu_Ok As Boolean
Dim X1 As String
On Error GoTo Error_Handler
'==========================================================================================================
Call rsYBIATAB0_Init(X_YBIATAB0)
arrUsr_K = 0: mCol = 4

'_________________________________________________________________________________________________


Do While Not rsSab.EOF

    Call rsYBIATAB0.rsYBIATAB0_GetBuffer(rsSab, Old_YBIATAB0)
    
    If Old_YBIATAB0.BIATABK1 <> X_YBIATAB0.BIATABK1 Then
'==========================================================================================================
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
             & " where BIATABID = 'BIA_VB_APP' and BIATABK1 = '" & Old_YBIATAB0.BIATABK1 & "'"
        Set rsSabX = cnsab.Execute(xSQL)
        Call rsYBIATAB0.rsYBIATAB0_GetBuffer(rsSabX, X_YBIATAB0)
        mXls2_Row = mXls2_Row + 1
        If (mXls2_Row Mod 10) Then Call lstErr_ChangeLastItem(lstErr, frmElp.cmdContext, "Menu : " & Old_YBIATAB0.BIATABK1): DoEvents
        If mId$(X_YBIATAB0.BIATABK1, 1, 1) = "=" Then
            wsExcel.Cells(mXls2_Row, 1) = "* " & X_YBIATAB0.BIATABK1
        Else
            wsExcel.Cells(mXls2_Row, 1) = X_YBIATAB0.BIATABK1
        End If
        wsExcel.Cells(mXls2_Row, 2) = mId$(X_YBIATAB0.BIATABTXT, 100, 1)
        wsExcel.Cells(mXls2_Row, 4) = Trim(mId$(X_YBIATAB0.BIATABTXT, 1, 69))
        For K = 1 To mXls2_Cols
            wsExcel.Cells(mXls2_Row, K).Interior.Color = RGB(255, 255, 96)
        Next K
'_________________________________________________________________________________________________
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
             & " where BIATABID = 'BIA_VB_DROIT' and BIATABK1 = '" & Old_YBIATAB0.BIATABK1 & "' order by substring(BIATABTXT,101,2)"
        
        Set rsSabX = cnsab.Execute(xSQL)
        Do While Not rsSabX.EOF
            mXls2_Row = mXls2_Row + 1
            wsExcel.Cells(mXls2_Row, 3) = rsSabX("BIATABK2")
            X = rsSabX("BIATABTXT")
            wsExcel.Cells(mXls2_Row, 4) = Trim(mId$(X, 1, 69))
            wsExcel.Cells(mXls2_Row, 2) = mId$(X, 100, 3)
            
            mCol = mId$(X, 101, 2) + 4
            wsExcel.Cells(mXls2_Row, mCol) = mId$(X, 100, 1)
            wsExcel.Cells(mXls2_Row, mCol).Font.Bold = True

            If mId$(X, 100, 1) <> "*" Then
                wsExcel.Cells(mXls2_Row, 2).Interior.Color = RGB(255, 255, 96)
            Else
                wsExcel.Cells(mXls2_Row, 2).Interior.Color = mColor_Y1
            End If
            
            rsSabX.MoveNext
        Loop

'_________________________________________________________________________________________________

        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
             & " where BIATABID = 'BIA_VB_MNU' and BIATABK1 = '" & Old_YBIATAB0.BIATABK1 & "' order by BIATABK2"
                
        Set rsSabX = cnsab.Execute(xSQL)
        arrBIA_VB_Mnu_Nb = 0
        lstW.Clear
        Do While Not rsSabX.EOF
            arrBIA_VB_Mnu_Nb = arrBIA_VB_Mnu_Nb + 1
            Call rsYBIATAB0_GetBuffer(rsSabX, arrBIA_VB_MNU(arrBIA_VB_Mnu_Nb))
            lstW.AddItem arrBIA_VB_MNU(arrBIA_VB_Mnu_Nb).BIATABK2 & "|" & arrBIA_VB_Mnu_Nb
            rsSabX.MoveNext
        Loop
        For K = 0 To arrBIA_VB_Mnu_Nb - 1
            lstW.ListIndex = K
            K2 = InStr(lstW.Text, "|")
            If K2 > 0 Then arrBIA_VB_Mnu_K = Val(mId$(lstW.Text, K2 + 1, Len(lstW.Text) - K2))
            mXls2_Row = mXls2_Row + 1
            wsExcel.Cells(mXls2_Row, 3) = Trim(arrBIA_VB_MNU(arrBIA_VB_Mnu_K).BIATABK2)
            X = arrBIA_VB_MNU(arrBIA_VB_Mnu_K).BIATABTXT
            wsExcel.Cells(mXls2_Row, 4) = Trim(mId$(X, 20, 79))
            wsExcel.Cells(mXls2_Row, 2).Interior.Color = mColor_G1
            
            For K2 = 1 To 19
                X1 = mId$(X, K2, 1)
                If X1 <> " " Then
                    mCol = K2 + 4
                    wsExcel.Cells(mXls2_Row, mCol) = X1
                    wsExcel.Cells(mXls2_Row, mCol).Font.Bold = True
                    wsExcel.Cells(mXls2_Row, mCol).Font.Color = RGB(0, 96, 0)
                End If
            Next K2
        Next K
'==========================================================================================================
    End If
    
        
'_________________________________________________________________________________________________
    rsSab.MoveNext
Loop

'==========================================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdSelect_SQL_Exportation_Detail"


End Sub



Public Sub cmdSelect_SQL_Exportation_Page3(lSheet As Integer, lLib As String)

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

mXls2_Cols = 4 + arrUsr_Nb
mXls2_Row = 1

wsExcel.Rows(1).Orientation = xlVertical
wsExcel.Rows(1).RowHeight = 130
'wsExcel.Rows(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Rows(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Rows(1).VerticalAlignment = Excel.xlVAlignTop

For K = 1 To arrUsr_Nb
    wsExcel.Cells(1, 4 + K) = arrUsr(K): wsExcel.Columns(4 + K).ColumnWidth = 2
    wsExcel.Columns(4 + K).HorizontalAlignment = Excel.xlHAlignCenter
    'wsExcel.Columns(4 + K).VerticalAlignment = Excel.xlVAlignTop
Next K

wsExcel.Cells(1, 1) = "Application": wsExcel.Columns(1).ColumnWidth = 12
wsExcel.Cells(1, 1).Orientation = xlHorizontal
wsExcel.Cells(1, 1).VerticalAlignment = Excel.xlVAlignCenter
wsExcel.Cells(1, 2) = "X": wsExcel.Columns(2).ColumnWidth = 4
wsExcel.Cells(1, 2).Orientation = xlHorizontal
wsExcel.Cells(1, 2).VerticalAlignment = Excel.xlVAlignCenter
wsExcel.Cells(1, 3) = "Droit": wsExcel.Columns(3).ColumnWidth = 10
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
    MsgBox Error, vbCritical, "cmdSelect_SQL_Exportation_Detail"


End Sub
Public Sub cmdSelect_SQL_Exportation_Page2(lSheet As Integer, lLib As String)

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

mXls2_Cols = 4 + 19
mXls2_Row = 1

'wsExcel.Rows(1).Orientation = xlVertical
'wsExcel.Rows(1).RowHeight = 115
'wsExcel.Rows(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Rows(1).HorizontalAlignment = Excel.xlHAlignCenter
'wsExcel.Rows(1).VerticalAlignment = Excel.xlVAlignTop

For K = 1 To 19
    wsExcel.Cells(1, 4 + K) = Format$(K, "00"): wsExcel.Columns(4 + K).ColumnWidth = 2
    'wsExcel.Columns(4 + K).HorizontalAlignment = Excel.xlHAlignCenter
    'wsExcel.Columns(4 + K).VerticalAlignment = Excel.xlVAlignTop
Next K

wsExcel.Cells(1, 1) = "Application": wsExcel.Columns(1).ColumnWidth = 12
wsExcel.Cells(1, 1).Orientation = xlHorizontal
wsExcel.Cells(1, 1).VerticalAlignment = Excel.xlVAlignCenter
wsExcel.Cells(1, 2) = "X": wsExcel.Columns(2).ColumnWidth = 4
wsExcel.Cells(1, 2).Orientation = xlHorizontal
wsExcel.Cells(1, 2).VerticalAlignment = Excel.xlVAlignCenter
wsExcel.Cells(1, 3) = "Menu": wsExcel.Columns(3).ColumnWidth = 10
wsExcel.Cells(1, 3).Orientation = xlHorizontal
wsExcel.Cells(1, 3).VerticalAlignment = Excel.xlVAlignCenter
wsExcel.Cells(1, 4) = "Libellé": wsExcel.Columns(4).ColumnWidth = 40
wsExcel.Cells(1, 4).Orientation = xlHorizontal
wsExcel.Cells(1, 4).VerticalAlignment = Excel.xlVAlignCenter


For K = 1 To mXls2_Cols
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next

'==========================================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdSelect_SQL_Exportation_Detail"


End Sub

Public Sub cmdSelect_SQL_Exportation_Page1(lSheet As Integer, lLib As String)

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

wsExcel.PageSetup.Zoom = 47

Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : " & lLib): DoEvents

mXls2_Cols = 4 + 19
mXls2_Row = 1

wsExcel.Rows(1).HorizontalAlignment = Excel.xlHAlignCenter

For K = 1 To 19
    wsExcel.Cells(1, 4 + K) = Format$(K, "00"): wsExcel.Columns(4 + K).ColumnWidth = 10
    'wsExcel.Columns(4 + K).HorizontalAlignment = Excel.xlHAlignCenter
    'wsExcel.Columns(4 + K).VerticalAlignment = Excel.xlVAlignTop
Next K

wsExcel.Cells(1, 1) = "Application": wsExcel.Columns(1).ColumnWidth = 12: wsExcel.Columns(1).Font.Bold = True
wsExcel.Cells(1, 1).Orientation = xlHorizontal
wsExcel.Cells(1, 1).VerticalAlignment = Excel.xlVAlignCenter
wsExcel.Cells(1, 2) = "Libellé": wsExcel.Columns(2).ColumnWidth = 40
wsExcel.Cells(1, 2).Orientation = xlHorizontal
wsExcel.Cells(1, 2).VerticalAlignment = Excel.xlVAlignCenter
wsExcel.Cells(1, 3) = "Utilisateur": wsExcel.Columns(3).ColumnWidth = 10: wsExcel.Columns(3).Font.Bold = True
wsExcel.Cells(1, 3).Orientation = xlHorizontal
wsExcel.Cells(1, 3).VerticalAlignment = Excel.xlVAlignCenter
wsExcel.Cells(1, 4) = "Service": wsExcel.Columns(4).ColumnWidth = 6
wsExcel.Cells(1, 4).Orientation = xlHorizontal
wsExcel.Cells(1, 4).VerticalAlignment = Excel.xlVAlignCenter


For K = 1 To mXls2_Cols
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next

'==========================================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdSelect_SQL_Exportation_Detail"


End Sub







Public Sub lstParam_Hab_Load(lSrv As String, lUsr As String)
Dim xWhere As String, xSQL As String
Dim K As Integer, X As String

fraParam_Hab.Visible = False

lstParam_Hab.Clear
xWhere = ""
If lSrv <> "" Then
    xWhere = " and SSIDOMUNIT = '" & lSrv & "'"
Else
    xWhere = " and SSIDOMUNIT <> ''"
End If

If lUsr <> "" Then xWhere = xWhere & " and SSIDOMUIDX like '" & lUsr & "%'"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN'" _
     & " and SSIDOMPRFX <> 'X'" & xWhere & "order by SSIDOMUIDX"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    lstParam_Hab.AddItem Trim(mId$(rsSab("SSIDOMUIDX"), 1, 12))
    rsSab.MoveNext
Loop

'_______________________________________________________________________________
ReDim arrBIA_VB_HAB(lstParam_Hab.ListCount + 1)
arrBIA_VB_HAB_Nb = 0

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB' and BIATABK1 = '" & mSelect_App & "'" _
  & " and BIATABK2 in (select SSIDOMUIDX from " & paramIBM_Library_SABSPE & ".YSSIDOM0 where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN'  and SSIDOMPRFX <> 'X'" & xWhere & ")" & "order by BIATABK2"

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    arrBIA_VB_HAB_Nb = arrBIA_VB_HAB_Nb + 1
    Call rsYBIATAB0_GetBuffer(rsSab, arrBIA_VB_HAB(arrBIA_VB_HAB_Nb))
    If mId$(arrBIA_VB_HAB(arrBIA_VB_HAB_Nb).BIATABTXT, mSelect_Droit_Seq, 1) <> " " Then
        X = Trim(arrBIA_VB_HAB(arrBIA_VB_HAB_Nb).BIATABK2)
        For K = 0 To lstParam_Hab.ListCount - 1
            lstParam_Hab.ListIndex = K
            If X = lstParam_Hab.Text Then lstParam_Hab.Selected(K) = True
        Next K
    End If
    rsSab.MoveNext
Loop
'_______________________________________________________________________________
If cmdSelect_SQL_K = "1L" Then
    For K = lstParam_Hab.ListCount - 1 To 0 Step -1
        If Not lstParam_Hab.Selected(K) Then
            lstParam_Hab.RemoveItem (K)
        End If
    Next K
End If
'_______________________________________________________________________________
lstParam_Hab.Enabled = True
cmdParam_Hab_Update.Visible = arrHab(2)
If mSelect_Droit_X = "X" Then
    If Not blnAdmin Then
        cmdParam_Hab_Update.Visible = False
        lstParam_Hab.Enabled = False
    End If
End If
 
If mSelect_Droit_X = "S" Then
    libParam_Hab_Srv = "Habilitation accordée pour l'utilisateur uniquement pour son service actuel."
    libParam_Hab_Srv.Visible = True
Else
    libParam_Hab_Srv.Visible = False
End If

blnParam_Hab_Change = False
lblParam_Hab_Quid = ""
fraParam_Hab.Caption = mSelect_App & " / " & mSelect_Droit
fraParam_Hab.Visible = True

End Sub

Public Sub lstParam_Droit_Load()
Dim xWhere As String, xSQL As String
Dim K As Integer, K2 As Integer, X As String, blnOk As Boolean, mTxt As String

fraParam_Hab.Visible = False
libParam_Hab_Srv.Visible = False
'_________________________________________________________________________________________________
lstParam_Usr_Srv
'_________________________________________________________________________________________________________
For K = 1 To 19: arrFct_X(K) = "": arrFct_ListIndex(K) = -1: Next K

lstParam_Hab.Clear
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
    & " where BIATABID = 'BIA_VB_DROIT' and BIATABK1 = '" & mSelect_App & "' order by substring(BIATABTXT,101,2)"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    blnOk = True
    X = rsSab("BIATABTXT")
    If mId$(X, 100, 1) = "X" Then
        If Not blnAdmin Then blnOk = False
    End If
    If blnOk Then
        lstParam_Hab.AddItem mId$(X, 100, 3) & " " & rsSab("BIATABK2") & " : " & X
        arrFct_X(mId$(X, 101, 2)) = mId$(X, 100, 1)
        arrFct_ListIndex(mId$(X, 101, 2)) = lstParam_Hab.ListCount - 1
    End If
    rsSab.MoveNext
Loop
'_______________________________________________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB' and BIATABK1 = '" & mSelect_App & "'and BIATABK2 = '" & mSelect_Usr & "'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    mTxt = rsSab("BIATABTXT")
    For K = 1 To 19
        If mId$(mTxt, K, 1) <> " " Then
            If arrFct_ListIndex(K) >= 0 And arrFct_ListIndex(K) < 19 Then lstParam_Hab.Selected(arrFct_ListIndex(K)) = True
        End If
    Next K
End If
'_______________________________________________________________________________
If cmdSelect_SQL_K = "2L" Then
    For K = lstParam_Hab.ListCount - 1 To 0 Step -1
        If Not lstParam_Hab.Selected(K) Then
            lstParam_Hab.RemoveItem (K)
        End If
    Next K
    
    For K = 1 To 19: arrFct_X(K) = "": arrFct_ListIndex(K) = -1: Next K
 
    For K = 0 To lstParam_Hab.ListCount - 1
        lstParam_Hab.ListIndex = K
        K2 = Val(mId$(lstParam_Hab.Text, 2, 2))
        arrFct_X(K2) = mId$(lstParam_Hab.Text, 1, 1)
        arrFct_ListIndex(K2) = K
    Next K
End If
'_______________________________________________________________________________
fraParam_Hab.Caption = mSelect_App & " / " & mSelect_Usr
If mId$(mTxt, 100, 3) <> "" And mSelect_Usr_Srv <> mId$(mTxt, 100, 3) Then
    libParam_Hab_Srv = "Les habilitations ont été accordées pour le service : " & mId$(mTxt, 100, 3) _
                    & ", l'utilisateur est affecté actuellement au service : " & mSelect_Usr_Srv
     libParam_Hab_Srv.Visible = True
End If

blnParam_Hab_Change = False
lblParam_Hab_Quid = " MAJ : " & mId$(mTxt, 105, 10) & " " & dateImp10_S(mId$(mTxt, 115, 8)) & " " & timeImp8(mId$(mTxt, 123, 6))
cmdParam_Hab_Update.Visible = arrHab(2)
fraParam_Hab.Visible = True
'_______________________________________________________________________________
End Sub


Public Sub lstParam_Mnu_Droit_Load()
Dim xWhere As String, xSQL As String
Dim K As Integer, K2 As Integer, X As String, blnOk As Boolean, mTxt As String

'_________________________________________________________________________________________________________
For K = 1 To 19: arrFct_X(K) = "": arrFct_ListIndex(K) = -1: Next K

lstParam_Mnu_Droit.Clear
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
    & " where BIATABID = 'BIA_VB_DROIT' and BIATABK1 = '" & mSelect_App & "' order by substring(BIATABTXT,101,2)"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    blnOk = True
    X = rsSab("BIATABTXT")
    lstParam_Mnu_Droit.AddItem mId$(X, 100, 3) & " " & rsSab("BIATABK2") & " : " & X
    arrFct_X(mId$(X, 101, 2)) = mId$(X, 100, 1)
    arrFct_ListIndex(mId$(X, 101, 2)) = lstParam_Mnu_Droit.ListCount - 1
    rsSab.MoveNext
Loop
'_______________________________________________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB' and BIATABK1 = '" & mSelect_App & "'and BIATABK2 = '" & mSelect_Usr & "'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    mTxt = rsSab("BIATABTXT")
    For K = 1 To 19
        If mId$(mTxt, K, 1) <> " " Then
            If arrFct_ListIndex(K) >= 0 And arrFct_ListIndex(K) < 19 Then lstParam_Mnu_Droit.Selected(arrFct_ListIndex(K)) = True
        End If
    Next K
End If
'_______________________________________________________________________________
If cmdSelect_SQL_K = "2L" Then
    For K = lstParam_Mnu_Droit.ListCount - 1 To 0 Step -1
        If Not lstParam_Mnu_Droit.Selected(K) Then
            lstParam_Mnu_Droit.RemoveItem (K)
        End If
    Next K
    
    For K = 1 To 19: arrFct_X(K) = "": arrFct_ListIndex(K) = -1: Next K
 
    For K = 0 To lstParam_Mnu_Droit.ListCount - 1
        lstParam_Mnu_Droit.ListIndex = K
        K2 = Val(mId$(lstParam_Mnu_Droit.Text, 2, 2))
        arrFct_X(K2) = mId$(lstParam_Mnu_Droit.Text, 1, 1)
        arrFct_ListIndex(K2) = K
    Next K
End If
'_______________________________________________________________________________
fraParam_Hab.Caption = mSelect_App & " / " & mSelect_Usr
If mId$(mTxt, 100, 3) <> "" And mSelect_Usr_Srv <> mId$(mTxt, 100, 3) Then
    libParam_Hab_Srv = "Les habilitations ont été accordées pour le service : " & mId$(mTxt, 100, 3) _
                    & ", l'utilisateur est affecté actuellement au service : " & mSelect_Usr_Srv
     libParam_Hab_Srv.Visible = True
End If

lblParam_Hab_Quid = " MAJ : " & mId$(mTxt, 105, 10) & " " & dateImp10_S(mId$(mTxt, 115, 8)) & " " & timeImp8(mId$(mTxt, 123, 6))
cmdParam_Hab_Update.Visible = arrHab(2)
fraParam_Hab.Visible = True
'_______________________________________________________________________________
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
Dim X As String
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "fgDetail_Display"
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString '&H00C0FFC0&
Select Case cmdSelect_SQL_K
    Case "9Mnu": fgDetail.FormatString = Replace(fgDetail_FormatString, "Droits         ", "Options de menu")
                fgDetail.BackColorFixed = &HC0FFC0
    
End Select
fgDetail.Row = 0
'___________________________________________________________________________

 
Do While Not rsSab.EOF
    X_YBIATAB0.BIATABK2 = rsSab("BIATABK2")
    Call rsYBIATAB0_GetBuffer(rsSab, X_YBIATAB0)
    blnOk = True
    If mId$(X_YBIATAB0.BIATABTXT, 100, 1) = "X" And Not blnAdmin Then blnOk = False
    If blnOk Then

        fgDetail.Rows = fgDetail.Rows + 1
        fgDetail.Row = fgDetail.Rows - 1
        fgDetail_Display_Line
    End If
    rsSab.MoveNext

Loop

If cmdSelect_SQL_K = "2" Or cmdSelect_SQL_K = "2L" Then
    If fgDetail.Rows > 2 Then fgDetail_Sort1 = 0: fgDetail_Sort2 = 0: fgdetail_Sort
    If fgDetail.Rows = 2 Then fgDetail.Col = 0: mSelect_Usr = Trim(fgDetail.Text)
End If

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgDetail.Rows - 1): DoEvents
fgDetail.Visible = True

'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub fgDetail_USR_Display()
Dim X As String
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "fgDetail_USR_Display"
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString '&H00C0FFC0&
fgDetail.Row = 0
'___________________________________________________________________________

 
Do While Not rsSab.EOF
        fgDetail.Rows = fgDetail.Rows + 1
        fgDetail.Row = fgDetail.Rows - 1
        fgDetail.Col = 0: fgDetail.Text = rsSab("SSIDOMUIDX")
        fgDetail.CellFontBold = True
        fgDetail.Col = 2: fgDetail.Text = rsSab("SSIDOMUNIT")

    rsSab.MoveNext

Loop

If cmdSelect_SQL_K = "2" Or cmdSelect_SQL_K = "2L" Then
    If fgDetail.Rows > 2 Then fgDetail_Sort1 = 0: fgDetail_Sort2 = 0: fgdetail_Sort
    If fgDetail.Rows = 2 Then fgDetail.Col = 0: mSelect_Usr = Trim(fgDetail.Text)
End If

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgDetail.Rows - 1): DoEvents
fgDetail.Visible = True

'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgDetail_Display_SQL(lId As String, lK1 As String, lK2 As String)
Dim X As String, xWhere As String
Dim xSQL As String

On Error GoTo Error_Handler

'___________________________________________________________________________

xWhere = "where BIATABID = '" & lId & "' and BIATABK1 = '" & lK1 & "'"
If lK2 <> "" Then xWhere = xWhere & " and BIATABK2 ='" & lK2 & "'"

If lId = "BIA_VB_DROIT" Then
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " & xWhere & " order by substring(BIATABTXT,101,2)"
Else

    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " & xWhere & " order by BIATABK1 , BIATABK2"
End If

Set rsSab = cnsab.Execute(xSQL)
 
fgDetail_Display
'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgDetail_Display_1DU()
Dim X As String
Dim xSQL As String

On Error GoTo Error_Handler

'___________________________________________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_DROIT'" _
       & " and BIATABK1 = '" & mSelect_App & "'" _
       & " and substring(BIATABTXT,101,2) = '" & mSelect_Droit_Seq & "'"

Set rsSab = cnsab.Execute(xSQL)
 
fgDetail_Display

If fgDetail.Rows > 1 Then
    Call lstParam_Hab_Load(Trim(mId$(cboSelect_Srv, 1, 3)), Trim(cboSelect_Usr))
End If
'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgDetail_Display_Line()
Dim X As String, wColor As Long

On Error Resume Next
Select Case cmdSelect_SQL_K
    Case "2":
        fgDetail.Col = 0: fgDetail.Text = X_YBIATAB0.BIATABK1
        fgDetail.CellFontBold = True
        fgDetail.Col = 2: fgDetail.Text = mId$(X_YBIATAB0.BIATABTXT, 26, 3)
    Case "2L":
        fgDetail.Col = 0: fgDetail.Text = X_YBIATAB0.BIATABK2
        fgDetail.CellFontBold = True
        If fgSelect.Rows = 2 Then
            fgDetail.Col = 2: fgDetail.Text = mId$(X_YBIATAB0.BIATABTXT, 100, 3)
            fgDetail.Col = 3: fgDetail.Text = X_YBIATAB0.BIATABK1 & " : " & mId$(X_YBIATAB0.BIATABTXT, 1, 19)
        End If
    Case "9Mnu":
     
        fgDetail.Col = 0: fgDetail.Text = X_YBIATAB0.BIATABK2
        fgDetail.CellFontBold = True
        fgDetail.Col = 3: fgDetail.Text = mId$(X_YBIATAB0.BIATABTXT, 20, 79)

    Case Else:
     
        fgDetail.Col = 0: fgDetail.Text = X_YBIATAB0.BIATABK2
        fgDetail.CellFontBold = True
        fgDetail.Col = 1: fgDetail.Text = mId$(X_YBIATAB0.BIATABTXT, 100, 1)
        fgDetail.Col = 2: fgDetail.Text = mId$(X_YBIATAB0.BIATABTXT, 101, 2)
        fgDetail.Col = 3: fgDetail.Text = mId$(X_YBIATAB0.BIATABTXT, 1, 99)
End Select
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

fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False

fgDetail_FormatString = fgDetail.FormatString
fgDetail.Enabled = True
fgDetail.Visible = False

fraParam_App.Top = fgSelect.Top
fraParam_App.Left = 5475
fraParam_App.ForeColor = &HC000C0

lblParam_APP_Code.ForeColor = vbWhite
libParam_Hab_Srv.ForeColor = vbYellow

fraParam_Hab.Top = fgSelect.Top
fraParam_Hab.Left = fgSelect.Left + fgSelect.Width - fraParam_Hab.Width - 200
fraParam_Hab.ForeColor = &HC000C0

Set fraSelect_4_Options.Container = fraSelect_Options.Container
fraSelect_4_Options.Left = fraSelect_Options.Left
fraSelect_4_Options.Top = fraSelect_Options.Top
fraSelect_4_Options.Visible = False

fraUsr_Srv.Visible = False

Set fraParam_Mnu.Container = fgSelect.Container
fraParam_Mnu.Left = fgSelect.Left + fgSelect.Width - fraParam_Mnu.Width - 200
fraParam_Mnu.Top = fgSelect.Top
fraParam_Mnu.Visible = False

'=================
'param_Init
'param_Init_SAB_DOSSIER
'param_Init_BIA_GOS

blnAdmin = False
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & "where BIATABID = 'BIA_VB_HAB' and BIATABK1 = 'BIA_VB_HAB' and BIATABK2 = '" & usrName_UCase & "'"
    
Set rsSab = cnsab.Execute(X)

If Not rsSab.EOF Then
    If mId$(rsSab("BIATABTXT"), 19, 1) <> " " Then blnAdmin = True
End If

'02 = "habilitation MAJ"
'09 = "consultation sans MAJ, incompatible 02 et 19"

If Not blnAdmin Then
    If arrHab(2) And arrHab(9) Then
        X = MsgBox("Voulez-vous procéder à des mises à jour ?", vbQuestion + vbYesNo, "Habilitations : choix MAJ ou Inspection")
        If X <> vbYes Then
            For K = 2 To 19: arrHab(K) = False: Next K
            blnAdmin = True
        End If
    End If
End If

'=================

fraSelect_Options.Visible = True
'___________________________________________________________________________
'cboSelect_SQL.Clear
'cboSelect_SQL.AddItem "1 - sélection droit => Utilisateurs"
'cboSelect_SQL.AddItem "1L -liste des utilisateurs habilités à ...."
'cboSelect_SQL.AddItem "2 - sélection Utilisateur => droit"
'cboSelect_SQL.AddItem "2L -liste des habilitations d'un utilisateur"
'cboSelect_SQL.AddItem "Xh - exportation des habilitations"

'If arrHab(2) Then
'    cboSelect_SQL.AddItem "3 - suppression de TOUTES les habilitations d'un utilisateur"
'    cboSelect_SQL.AddItem "4 - Duplication de TOUTES les habilitations d'un utilisateur"
'End If


'If arrHab(4) Then
'    cboSelect_SQL.AddItem "3# - habilitations des utilisateurs à supprimer  ?"
'    cboSelect_SQL.AddItem "9App - gestion des applications et droits"
'    cboSelect_SQL.AddItem "9Usr - gestion des utilisateurs et des services"
'    cboSelect_SQL.AddItem "9Mnu - gestion des options de menu"
'End If

If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0
'___________________________________________________________________________

cboParam_APP_VBP.Clear
cboParam_APP_VBP.AddItem "BIA_SAB"
cboParam_APP_VBP.AddItem "BIA_SWIFT"
cboParam_APP_VBP.AddItem "BIA_AUDIT"
cboParam_APP_VBP.AddItem "BIA_SYSTEM"
cboParam_APP_VBP.AddItem "BIA_DWH"
cboParam_APP_VBP.AddItem "BIA_JRN"
cboParam_APP_VBP.AddItem "BIA_SIDESRV"
cboParam_APP_VBP.AddItem "BIAS820I"
cboParam_APP_VBP.AddItem "X"

'___________________________________________________________________________
Parametrage_Load_App

Form_Init_cboSelect_Srv
Form_Init_cboSelect_Usr

'___________________________________________________________________________

blnParam_Hab_Change = False
blnControl = True


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



Private Sub fgSelect_Display()

Dim K As Long, blnOk As Boolean, wX As String

On Error GoTo Error_Handler
currentAction = "fgSelect_Display"
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Height = 3900

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
'Droits
fgSelect.Row = 0

Do While Not rsSab.EOF
    Call rsYBIATAB0_GetBuffer(rsSab, X_YBIATAB0)
    blnOk = True
    If mId$(X_YBIATAB0.BIATABTXT, 100, 1) = "X" And Not blnAdmin Then blnOk = False
    If blnOk Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_Display_Line
    End If
    rsSab.MoveNext

Loop

fgSelect.Visible = True

If fgSelect.Rows = 2 Then
    fgSelect.Col = 0: wX = Trim(fgSelect.Text)
    
    Select Case cmdSelect_SQL_K
        Case "1", "1L"
            mSelect_App = wX
            Call fgDetail_Display_SQL("BIA_VB_DROIT", wX, "")
        Case "2", "2L"
            mSelect_App = wX
            If mSelect_Usr <> "" Then lstParam_Droit_Load
    End Select
End If

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgSelect_3_Display(lFct As String)

Dim X As String

On Error GoTo Error_Handler
currentAction = "fgSelect_3_Display"
fgSelect.Visible = False
If lFct = "G_MIN" Then
    fgSelect_Reset
    fgSelect.Height = 8000
    fgSelect.Rows = 1
    X = Replace(fgSelect_FormatString, "Application", "Utilisateur")
    X = Replace(fgSelect_FormatString, "libellé de l'application", Space(20))
    fgSelect.FormatString = X
                     
    fgSelect.Row = 0
    
    Do While Not rsSab.EOF
        
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect.Col = 0: fgSelect.Text = rsSab("BIATABK2")
        fgSelect.Col = 3: fgSelect.Text = "utilisateur ayant des habilitations VB et n'ayant plus d'habilitations SAB"
        rsSab.MoveNext
    
    Loop
Else
    Do While Not rsSab.EOF
        
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect.Col = 0: fgSelect.Text = rsSab("BIATABK2")
        fgSelect.CellForeColor = vbYellow 'vbRed
        fgSelect.Col = 3: fgSelect.Text = lFct
        fgSelect.CellForeColor = vbRed
        rsSab.MoveNext
    
    Loop
    If fgSelect.Rows > 2 Then
        fgSelect_SortAD = 6
        fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
        fgSelect_Sort
    End If
    fgSelect.Visible = True
End If

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_Display_Line()
Dim X As String

On Error Resume Next
'Select Case cmdSelect_SQL_K
'    Case "1", "9App", "2", ":"
                fgSelect.Col = 0: fgSelect.Text = X_YBIATAB0.BIATABK1
                fgSelect.CellFontBold = True
                fgSelect.Col = 3: fgSelect.Text = Trim(mId$(X_YBIATAB0.BIATABTXT, 1, 69))
                fgSelect.Col = 1: fgSelect.Text = mId$(X_YBIATAB0.BIATABTXT, 100, 1)
                fgSelect.Col = 2: fgSelect.Text = Trim(mId$(X_YBIATAB0.BIATABTXT, 80, 12))
                
'End Select
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim wFct As String

mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

wFct = UCase$(Trim(mId$(Msg, 1, 12)))
'Call BiaPgmAut_Init(wFct, BIA_VB_Habilitations_Aut)

Call BIA_VB_HAB(wFct, arrHab(), cboSelect_SQL)

Select Case wFct
    'Case "@?????":
    Case Else: blnAuto = False: Form_Init

End Select
End Sub



Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgSelect.Visible = False
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = 1 To 1 Step -1
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 1 To 1 Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
    End If
End If
fgSelect.LeftCol = fgSelect.FixedCols
fgSelect.Visible = True
End Sub


Private Sub cboSelect_4_APP_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub cboSelect_4_Usr1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub cboSelect_4_Usr2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub cboSelect_App_Click()
If blnControl Then cmdSelect_Reset

End Sub

Private Sub cboSelect_App_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

End Sub


Private Sub cboSelect_Srv_Click()
If blnControl Then cmdSelect_Reset

End Sub

Private Sub cboSelect_Usr_Click()
If blnControl Then cmdSelect_Reset

End Sub

Private Sub cboSelect_Usr_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub cmdParam_App_Add_17_Click()
txtParam_APP_Code = "Documentatio"
txtParam_APP_Seq = "17"
txtParam_App_Lib = "accès à la documentation (DocuShare)"
optParam_App_S = True
cmdParam_App_Add_Click
End Sub

Private Sub cmdParam_App_Add_18_Click()
txtParam_APP_Code = "Paramétrage"
txtParam_APP_Seq = "18"
txtParam_App_Lib = "Paramétrage"
optParam_App_S = True
cmdParam_App_Add_Click

End Sub


Private Sub cmdParam_App_Add_19_Click()
txtParam_APP_Code = "Admin"
txtParam_APP_Seq = "19"
txtParam_App_Lib = "réservé administrateur"
optParam_App_X = True
cmdParam_App_Add_Click

End Sub


Private Sub cmdParam_Hab_Quit_Click()
fraParam_Hab.Visible = False
End Sub

Private Sub cmdParam_Hab_Update_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

Select Case cmdSelect_SQL_K
    Case "1", "1L", "1D*U": cmdParam_Hab_Update_1
    Case "2", "2L": cmdParam_Hab_Update_2
    
End Select
blnParam_Hab_Change = False

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Hab_Update_1()
Dim K As Integer, K2 As Integer, blnOk As Boolean
Dim xHab As String, blnUpdate As Boolean


For K = 0 To lstParam_Hab.ListCount - 1
    lstParam_Hab.ListIndex = K
    mSelect_Usr = Trim(lstParam_Hab)
    If lstParam_Hab.Selected(K) Then
        xHab = IIf(mSelect_Droit_X = "", "*", mSelect_Droit_X)
    Else
        xHab = " "
    End If
    blnOk = False
    For K2 = 1 To arrBIA_VB_HAB_Nb
        If Trim(arrBIA_VB_HAB(K2).BIATABK2) = mSelect_Usr Then
            blnOk = True
            Old_YBIATAB0 = arrBIA_VB_HAB(K2)
            Exit For
        End If
    Next K2
    If Not blnOk Then
        If xHab <> " " Then
            New_YBIATAB0.BIATABID = "BIA_VB_HAB"
            New_YBIATAB0.BIATABK1 = mSelect_App
            New_YBIATAB0.BIATABK2 = mSelect_Usr
            New_YBIATAB0.BIATABTXT = ""
            
            lstParam_Usr_Srv
            Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = mSelect_Usr_Srv
            
            If mSelect_Droit_Seq > 0 And mSelect_Droit_Seq <= 19 Then
                Mid$(New_YBIATAB0.BIATABTXT, mSelect_Droit_Seq, 1) = xHab
                If Not IsNull(Parametrage_New) Then
                    Call MsgBox(Error, vbCritical, "cmdParam_Hab_Update.New " & mSelect_App & mSelect_Usr)
                End If
            End If
        End If
    Else
        If mSelect_Droit_Seq > 0 And mSelect_Droit_Seq <= 19 Then
            
            If xHab <> mId$(Old_YBIATAB0.BIATABTXT, mSelect_Droit_Seq, 1) Then
                New_YBIATAB0 = Old_YBIATAB0
                Mid$(New_YBIATAB0.BIATABTXT, mSelect_Droit_Seq, 1) = xHab
                
                lstParam_Usr_Srv
                Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = mSelect_Usr_Srv
                If Trim(mId$(New_YBIATAB0.BIATABTXT, 1, 19)) = "" Then
                    V = Parametrage_Delete
                Else
                    V = Parametrage_Update
                End If
                If Not IsNull(V) Then
                    Call MsgBox(Error, vbCritical, "cmdParam_Hab_Update.upadte " & mSelect_App & mSelect_Usr)
                End If
            End If
        End If
    End If
    
Next K
fraParam_Hab.Visible = False

End Sub

Private Sub cmdParam_Hab_Update_2()
Dim K As Integer, K2 As Integer, xSQL As String, Kseq As Integer
Dim xHab As String, blnNew As Boolean
Dim X As String


xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB'" _
     & " and BIATABK1 = '" & mSelect_App & "' and BIATABK2 = '" & mSelect_Usr & "'"

Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    blnNew = False
    Call rsYBIATAB0_GetBuffer(rsSab, Old_YBIATAB0)
    New_YBIATAB0 = Old_YBIATAB0
Else
    blnNew = True
    New_YBIATAB0.BIATABID = "BIA_VB_HAB"
    New_YBIATAB0.BIATABK1 = mSelect_App
    New_YBIATAB0.BIATABK2 = mSelect_Usr
    New_YBIATAB0.BIATABTXT = ""
    Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = mSelect_Usr_Srv
End If

If mId$(New_YBIATAB0.BIATABTXT, 100, 3) <> mSelect_Usr_Srv Then
    X = MsgBox("Cet utilisateur a changé de service : " & mId$(New_YBIATAB0.BIATABTXT, 100, 3) & " => " & mSelect_Usr_Srv _
     & vbCrLf & "confirmez-vous ses habilitations ?", vbQuestion + vbYesNo, "Mise à jour des habilitations : " & mSelect_Usr)
     If X <> vbYes Then Exit Sub
End If

Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = mSelect_Usr_Srv

For K = 0 To lstParam_Hab.ListCount - 1
    lstParam_Hab.ListIndex = K
    Kseq = 0
    For K2 = 1 To 19
        If arrFct_ListIndex(K2) = K Then
            Kseq = K2
            Exit For
        End If
    Next K2
    If lstParam_Hab.Selected(K) Then
        xHab = mId$(lstParam_Hab.Text, 1, 1)
        If xHab = " " Then xHab = "*"
    Else
        xHab = " "
    End If
    If Kseq > 0 Then Mid$(New_YBIATAB0.BIATABTXT, Kseq, 1) = xHab
Next K

If blnNew Then
    If Not IsNull(Parametrage_New) Then
       Call MsgBox(Error, vbCritical, "cmdParam_Hab_Update.New " & mSelect_App & mSelect_Usr)
    End If
Else
    If Trim(mId$(New_YBIATAB0.BIATABTXT, 1, 19)) = "" Then
        V = Parametrage_Delete
    Else
        V = Parametrage_Update
    End If
    If Not IsNull(V) Then
       Call MsgBox(Error, vbCritical, "cmdParam_Hab_Update.upadte " & mSelect_App & mSelect_Usr)
    End If
End If


fraParam_Hab.Visible = False

End Sub



Private Sub cmdParam_Mnu_Add_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

If fraParam_Mnu_Control("New") Then
    If IsNull(Parametrage_New) Then
        fraParam_Mnu.Visible = False
        Call fgDetail_Display_SQL("BIA_VB_MNU", mSelect_App, "")
    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Mnu_Delete_Click()
Dim X As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

blnOk = True
X = Trim(txtParam_Mnu_Code)
If X <> Trim(Old_YBIATAB0.BIATABK2) Then
    blnOk = False
    Call MsgBox("Le code identifiant a été modifié," & vbCrLf & " la suppression n'est pas possible", vbCritical, "BIA_VB_Habilitations : paramétrage")
End If


If blnOk Then
    If IsNull(Parametrage_Delete) Then
        fraParam_Mnu.Visible = False
        Call fgDetail_Display_SQL("BIA_VB_MNU", mSelect_App, "")
    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Mnu_Quit_Click()
fraParam_Mnu.Visible = False
End Sub

Private Sub cmdParam_Mnu_Update_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

If fraParam_Mnu_Control("Update") Then

    If New_YBIATAB0.BIATABK1 <> Old_YBIATAB0.BIATABK1 _
    Or New_YBIATAB0.BIATABK2 <> Old_YBIATAB0.BIATABK2 Then
        Call MsgBox("Le code identifiant a été modifié," & vbCrLf & " la modification n'est pas possible", vbCritical, "BIA_VB_Habilitations : paramétrage")
    Else

        If IsNull(Parametrage_Update) Then
            fraParam_Mnu.Visible = False
            Call fgDetail_Display_SQL("BIA_VB_MNU", mSelect_App, "")
        End If
    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdPrint_Click()
If Not txtRTF.Visible Then
    'cmdPrint_Display
Else
End If



End Sub

Private Sub cmdSrv_Quit_Click()
fraSrv.Visible = False

End Sub

Private Sub cmdUsr_Quit_Click()
fraUsr.Visible = False

End Sub

Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, xUUMID As String
On Error Resume Next

fraParam_Hab.Visible = False

If y <= fgDetail.RowHeightMin Then
    fgDetail.Visible = False
    Select Case fgDetail.Col
        Case 0: fgDetail_Sort1 = 0: fgDetail_Sort2 = 0: fgdetail_Sort
        Case 1:  fgDetail_Sort1 = 1: fgDetail_Sort2 = 2: fgdetail_Sort
        Case 2:  fgDetail_Sort1 = 2: fgDetail_Sort2 = 2: fgdetail_Sort
        Case 3:  fgDetail_Sort1 = 3: fgDetail_Sort2 = 3: fgdetail_Sort
    End Select
    fgDetail.Visible = True
Else
    If fgDetail.Rows > 1 Then
        lstParam_Hab_Change
        
        Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
        fgDetail.Col = 0: wX = Trim(fgDetail.Text)
        Select Case cmdSelect_SQL_K
            Case "1", "1L"
                mSelect_Droit = wX
                fgDetail.Col = 2: mSelect_Droit_Seq = Val(fgDetail.Text)
                fgDetail.Col = 1: mSelect_Droit_X = Trim(fgDetail.Text)
                Call lstParam_Hab_Load(Trim(mId$(cboSelect_Srv, 1, 3)), Trim(cboSelect_Usr))
            Case "2", "2L"
                mSelect_Usr = wX
                If mSelect_App <> "" Then lstParam_Droit_Load
            Case "9App"
                mSelect_Droit = wX
                Call fraParam_App_Display("BIA_VB_DROIT", App_YBIATAB0.BIATABK1, wX)
            Case "9Mnu"
                mSelect_Droit = wX
                Call fraParam_Mnu_Display("BIA_VB_MNU", mSelect_App, mSelect_Droit)
        End Select
   End If
End If
fgDetail.LeftCol = 0

End Sub




Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, xUUMID As String
On Error Resume Next

fraParam_Hab.Visible = False

If y <= fgSelect.RowHeightMin Then
    fgSelect.Visible = False
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 2:  fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3:  fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
    End Select
    fgSelect.Visible = True
Else
    If fgSelect.Rows > 1 Then
        
        lstParam_Hab_Change
        
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = 0: wX = Trim(fgSelect.Text)
        
        Select Case cmdSelect_SQL_K
            Case "1", "1L"
                mSelect_App = wX
                Call fgDetail_Display_SQL("BIA_VB_DROIT", wX, "")
            Case "1D*U"
                mSelect_App = wX
                Call fgDetail_Display_1DU
            Case "2", "2L"
                mSelect_App = wX
                If mSelect_Usr <> "" Then lstParam_Droit_Load
            Case "3#"
                    cmdSelect_SQL_3_Delete wX, ""
            Case "9App"
                
                mSelect_App = wX
                Call fraParam_App_Display("BIA_VB_APP", mSelect_App, "")
                Call fgDetail_Display_SQL("BIA_VB_DROIT", mSelect_App, "")
                If fgDetail.Rows = 1 Then
                    fgDetail.Rows = 2: fgDetail.Row = 1
                    fgDetail.Col = 3: fgDetail.Text = "Ajouter un droit"
                End If
            Case "9Mnu"
                
                mSelect_App = wX
                Call fgDetail_Display_SQL("BIA_VB_MNU", mSelect_App, "")
                If fgDetail.Rows = 1 Then
                    fgDetail.Rows = 2: fgDetail.Row = 1
                    fgDetail.Col = 3: fgDetail.Text = "Ajouter une option de menu"
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

lstParam_Hab_Change

lstErr.Clear
fgSelect.Visible = False
fgDetail.Visible = False
fraParam_App.Visible = False
fraParam_Hab.Visible = False

cmdSelect_Ok.BackColor = vbGreen

End Sub

Public Sub lstParam_Hab_Change()
Dim X As String
If blnParam_Hab_Change Then
    X = MsgBox("Voulez-vous enregistrer les modifications ?", vbQuestion & vbYesNo, "Mise à jour des habilitations")
    If X = vbYes Then cmdParam_Hab_Update_Click
End If
blnParam_Hab_Change = False


End Sub

Public Sub cmdSelect_Reset()
Dim K As Integer
If blnControl Then
    cmdSelect_Clear
    K = InStr(cboSelect_SQL, "-")
    If K > 1 Then
        cmdSelect_SQL_K = Trim(mId$(cboSelect_SQL, 1, K - 1))
    Else
        cmdSelect_SQL_K = "???"
    End If
    
    fraSelect_4_Options.Visible = False
    fraSelect_Options.Visible = True
    cboSelect_App.Visible = True
    cboSelect_Srv.Visible = True
    cboSelect_Usr.Visible = True
    cmdSelect_Ok.Visible = True
    
    Select Case cmdSelect_SQL_K
        Case "1", "1L":
        Case "1D*U": 'cboSelect_Srv.Visible = False: cboSelect_Usr.Visible = False
        Case "2", "2L":
        Case "9App", "9Mnu", "XHab"
        Case "3":  cboSelect_Srv.Visible = False
        Case "3#": fraSelect_Options.Visible = False
        Case "4": fraSelect_Options.Visible = False: fraSelect_4_Options.Visible = True
                  cmdSelect_SQL_4_Init

        Case "9Usr": fraSelect_Options.Visible = False ': fraUsr_Srv_Init
        Case "SI_Doc": fraSelect_Options.Visible = False ': fraUsr_Srv_Init

    End Select

End If
End Sub


Private Sub cmdSelect_SQL_2()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2"

mSelect_Usr = "": mSelect_Usr_Srv = ""

'____________________________________________________________________________________________
If cmdSelect_SQL_K <> "2L" Then
    xWhere = ""
    X = Trim(mId$(cboSelect_Srv, 1, 3))
    If X <> "" Then
        xWhere = " and SSIDOMUNIT = '" & X & "'"
    Else
        xWhere = " and SSIDOMUNIT <> ''"
    End If
    
    X = Trim(cboSelect_Usr)
    If X <> "" Then xWhere = xWhere & " and SSIDOMUIDX like '" & X & "%'"
    
    cmdSelect_SQL_1
    
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN'" _
         & " and SSIDOMUNIT <> 'S99' " & xWhere & "order by SSIDOMUIDX"
    Set rsSab = cnsab.Execute(xSQL)
    
    '''fgDetail_Display
    fgDetail_USR_Display
Else

    xWhere = ""
    X = Trim(cboSelect_App)
    If X <> "" Then
        xWhere = xWhere & " and BIATABK1 like '" & X & "%'"
    End If
    X = Trim(mId$(cboSelect_Srv, 1, 3))
    If X <> "" Then xWhere = xWhere & " and substring(BIATABTXT,100,3) = '" & X & "'"
    X = Trim(cboSelect_Usr)
    If X <> "" Then xWhere = xWhere & " and BIATABK2 like '" & X & "%'"
    
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_APP' " _
      & " and BIATABK1 in (select distinct(BIATABK1) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB'" & xWhere & ") " & "order by BIATABK1"
   
    Set rsSab = cnsab.Execute(xSQL)
    fgSelect_Display
    
    If fgSelect.Rows > 2 Then
        xSQL = "select distinct(BIATABK2) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB'" & xWhere & " order by BIATABK2 "
    Else
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB'" & xWhere & " order by BIATABK2 "
    End If
    Set rsSab = cnsab.Execute(xSQL)
    
    fgDetail_Display
End If
'____________________________________________________________________________________________________


Set rsSab = Nothing

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
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_3"
cmdSelect_SQL_3_Delete Trim(cboSelect_Usr), Trim(cboSelect_App)


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub cmdSelect_SQL_3_Delete(lUsr As String, lAPP As String)
Dim V, X As String, xWhere As String
Dim xSQL As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_3"

xWhere = ""
If lAPP <> "" Then xWhere = " and BIATABK1 like '" & lAPP & "%'"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB' and BIATABK2 = '" & Trim(lUsr) & "'" & xWhere
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    Call MsgBox("Il n'y a pas d'habilitations pour l'utilisateur : " & lUsr, vbInformation, "BIA_VB_HAB : suppression")
Else
    If lAPP = "" Then
        X = MsgBox("Confirmez-vous la suppression de TOUTES les habilitations pour l'utilisateur : " & lUsr, vbYesNo, "BIA_VB_HAB : suppression")
    Else
        X = MsgBox("Confirmez-vous la suppression de l'habilitation " & lAPP & " pour l'utilisateur : " & lUsr, vbYesNo, "BIA_VB_HAB : suppression")
    End If
    
    If X = vbYes Then
        xSQL = "delete from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB' and BIATABK2 = '" & Trim(lUsr) & "'" & xWhere
        Call FEU_ROUGE
        Call Parametrage_SQL(xSQL)
        Call FEU_VERT
        If cmdSelect_SQL_K = "3#" Then cmdSelect_SQL_3_VB_SAB
    End If
End If

Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_4()
Dim V, X As String, xUsr1 As String, xUsr2 As String
Dim xSQL As String, xWhere As String, xSrv As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_4"
xUsr1 = Trim(cboSelect_4_Usr1)
xUsr2 = Trim(cboSelect_4_Usr2)

If xUsr1 = "" Then
    Call MsgBox("Préciser l'utilisateur d'origine : ", vbInformation, "BIA_VB_HAB : suppression")
    Exit Sub

End If
If xUsr2 = "" Then
    Call MsgBox("Préciser l'utilisateur de destination : ", vbInformation, "BIA_VB_HAB : suppression")
    Exit Sub

End If
If xUsr2 = xUsr1 Then
    Call MsgBox("l'utilisateur d'origine =  l'utilisateur de destination : ", vbInformation, "BIA_VB_HAB : suppression")
    Exit Sub

End If
'___________________________________________________________________________________________
     
xSQL = "select *from " & paramIBM_Library_SABSPE & ".YSSIDOM0 where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN'" _
     & "  and SSIDOMUIDX = '" & xUsr1 & "'"

Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    Call MsgBox("Erreur de lecture : " & xSQL, vbCritical, "cmdSelect_SQL_4")
    Exit Sub
End If
X = rsSab("SSIDOMUNIT")
xSQL = "select *from " & paramIBM_Library_SABSPE & ".YSSIDOM0 where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN'" _
     & "  and SSIDOMUIDX = '" & xUsr2 & "'"

Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    Call MsgBox("Erreur de lecture : " & xSQL, vbCritical, "cmdSelect_SQL_4")
    Exit Sub
End If
xSrv = rsSab("SSIDOMUNIT")
If X <> xSrv Then

    X = MsgBox("Ces utilisateurs ne sont pas dans le même service : " & X & " <=> " & xSrv _
     & vbCrLf & "confirmez-vous la duplication ?", vbQuestion + vbYesNo, "Duplication des habilitations ")
     If X <> vbYes Then Exit Sub
End If

xWhere = ""
If Trim(cboSelect_4_APP) <> "" Then
    xWhere = "and BIATABK1 ='" & Trim(cboSelect_4_APP) & "'"
End If


'___________________________________________________________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB' and BIATABK2 = '" & xUsr1 & "'" & xWhere
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    Call MsgBox("Il n'y a pas d'habilitations pour l'utilisateur : " & cboSelect_4_Usr1, vbInformation, "BIA_VB_HAB : suppression")
    Exit Sub
End If

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB' and BIATABK2 = '" & xUsr2 & "'" & xWhere
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    If xWhere = "" Then
        X = MsgBox("Confirmez-vous la suppression de TOUTES les habilitations pour l'utilisateur : " & cboSelect_4_Usr2, vbYesNo, "BIA_VB_HAB : Duplication des habilitations")
    Else
        X = MsgBox("Confirmez-vous la suppression de l'habilitation " & Trim(cboSelect_4_APP) & " pour l'utilisateur : " & cboSelect_4_Usr2, vbYesNo, "BIA_VB_HAB : Duplication des habilitations")
    End If
    
    If X = vbYes Then
        xSQL = "delete from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB' and BIATABK2 = '" & xUsr2 & "'" & xWhere
        Call FEU_ROUGE
        Call Parametrage_SQL(xSQL)
        Call FEU_VERT
    Else
        Exit Sub
    End If
End If

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB' and BIATABK2 = '" & xUsr1 & "'" & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call rsYBIATAB0_GetBuffer(rsSab, New_YBIATAB0)
    New_YBIATAB0.BIATABK2 = xUsr2
    Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = xSrv
    Parametrage_New
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

Private Function Parametrage_SQL(lSQL As String)
Dim Nb As Long
On Error GoTo Error_Handler

Dim V
App_Debug = "Parametrage_SQL"

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Set rsSab_Update = cnSab_Update.Execute(lSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    V = "Erreur MAJ : " & lSQL
    GoTo Error_MsgBox
End If

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

Mid$(New_YBIATAB0.BIATABTXT, 105, 10) = usrName_UCase & "          "
Mid$(New_YBIATAB0.BIATABTXT, 115, 8) = DSys
Mid$(New_YBIATAB0.BIATABTXT, 123, 6) = time_Hms

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

Mid$(New_YBIATAB0.BIATABTXT, 105, 10) = usrName_UCase & "          "
Mid$(New_YBIATAB0.BIATABTXT, 115, 8) = DSys
Mid$(New_YBIATAB0.BIATABTXT, 123, 6) = time_Hms

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



Private Sub cmdSelect_SQL_1()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"
   
mSelect_App = "": mSelect_Usr = "": mSelect_Usr_Srv = ""

xWhere = ""
X = Trim(cboSelect_App)
If X <> "" Then xWhere = " and BIATABK1 like '" & X & "%'"


If blnAdmin Then

    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_APP'" & xWhere & " order by BIATABK1"
Else

    X = Trim(mId$(cboSelect_Srv, 1, 3))
    If X <> "" Then xWhere = xWhere & " and substring(BIATABTXT,100,3) = '" & X & "'"
    X = Trim(cboSelect_Usr)
    If X <> "" Then xWhere = xWhere & " and BIATABK2 like '" & X & "%'"
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_APP' " _
      & " and BIATABK1 in (select distinct(BIATABK1) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB'" & xWhere & ")" & "order by BIATABK1"
    
End If
Set rsSab = cnsab.Execute(xSQL)
  
fgSelect_Display



Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_1DU()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1DU"
   
mSelect_App = "": mSelect_Usr = "": mSelect_Usr_Srv = ""

X = InputBox("préciser le numéro du droit 01-19 " _
    & vbCrLf & "     =========================" & vbCrLf & 17 _
    & vbCrLf & "     =========================", "Habililations VB : numéro du droit ", 17)
If Not IsNumeric(X) Then Exit Sub
If Val(X) < 1 Or Val(X) > 19 Then Exit Sub

mSelect_Droit_Seq = Format(X, "00")

xWhere = ""
X = Trim(cboSelect_App)
If X <> "" Then xWhere = " and BIATABK1 like '" & X & "%'"



xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_APP'" & xWhere _
      & " and BIATABK1 in (select distinct(BIATABK1) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_DROIT'" _
      & " and substring(BIATABTXT,101,2) = '" & mSelect_Droit_Seq & "')" _
      & " order by BIATABK1"
      
Set rsSab = cnsab.Execute(xSQL)
  
fgSelect_Display



Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_SI_Doc()
Dim X As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_SI_Doc"
X = InputBox("Préciser la valeur recherchée :", "DocuShare : documentation informatique")
If Trim(X) <> "" Then
    DS_Server_Open
    Call DS_Document_Load(X, paramDocuShare_Collection_SI_Doc)
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub cmdParam_App_Add_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

If fraParam_App_Control("New") Then
    If IsNull(Parametrage_New) Then
        If Trim(New_YBIATAB0.BIATABID) = "BIA_VB_APP" Then
            New_YBIATAB0.BIATABID = "BIA_VB_DROIT"
            New_YBIATAB0.BIATABTXT = "accès à l'application " & New_YBIATAB0.BIATABK1
            Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "#01"
            Call Parametrage_New
            New_YBIATAB0.BIATABID = "BIA_VB_DROIT"
            New_YBIATAB0.BIATABK2 = "Admin"
            New_YBIATAB0.BIATABTXT = "réservé administrateur"
            Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "X19"
            Call Parametrage_New
            New_YBIATAB0.BIATABID = "BIA_VB_DROIT"
            New_YBIATAB0.BIATABK2 = "Paramétrage"
            New_YBIATAB0.BIATABTXT = "paramétrage"
            Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "X18"
            Call Parametrage_New
        End If
        Select Case Trim(Old_YBIATAB0.BIATABID)
            Case "BIA_VB_APP":
                cmdSelect_Reset
                cmdSelect_Ok_Click
            Case Else
                Call fgDetail_Display_SQL("BIA_VB_DROIT", mSelect_App, "")
        End Select
         fraParam_App.Visible = False

    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub
Private Function fraParam_App_Control(lFct As String)
Dim xId As String, xText As String, xSeq As String, xSQL As String
Dim blnOk As Boolean, blnSeq_Ok As Boolean

blnOk = True
xId = mId$(Trim(txtParam_APP_Code) & Space(11), 1, 12)
If Trim(xId) = "" Then
    blnOk = False
    Call MsgBox("Préciser le code ", vbCritical, "BIA_VB_Habilitations : paramétrage")
Else
    xText = mId$(Trim(txtParam_App_Lib) & Space(100), 1, 99)
    If Trim(xId) = "" Then
        blnOk = False
        Call MsgBox("Préciser le libellé ", vbCritical, "BIA_VB_Habilitations : paramétrage")
    End If
End If

'________________________________________________________________
If blnOk Then
    New_YBIATAB0 = Old_YBIATAB0
    Select Case Trim(New_YBIATAB0.BIATABID)
        Case "BIA_VB_APP":
            New_YBIATAB0.BIATABK1 = mId$(xId, 1, 12)
            New_YBIATAB0.BIATABTXT = mId$(xText, 1, 69)
            Mid$(New_YBIATAB0.BIATABTXT, 100, 1) = IIf(optParam_App_X = "1", "X", "*")
            Mid$(New_YBIATAB0.BIATABTXT, 80, 12) = cboParam_APP_VBP
            Mid$(New_YBIATAB0.BIATABTXT, 70, 10) = Trim(txtParam_APP_Doc)
       Case "BIA_VB_DROIT"
            New_YBIATAB0.BIATABK2 = mId$(xId, 1, 12)
            New_YBIATAB0.BIATABTXT = mId$(xText, 1, 99)
            If optParam_App_X Then
                Mid$(New_YBIATAB0.BIATABTXT, 100, 1) = "X"
            Else
                If optParam_App_S Then
                    Mid$(New_YBIATAB0.BIATABTXT, 100, 1) = "S"
                Else
                    Mid$(New_YBIATAB0.BIATABTXT, 100, 1) = "*"
                End If
            End If
            
            xSeq = Format(Val(txtParam_APP_Seq), "00")
            Mid$(New_YBIATAB0.BIATABTXT, 101, 2) = xSeq
            blnSeq_Ok = False
            If lFct = "Update" Then
                If mId$(New_YBIATAB0.BIATABTXT, 101, 2) = mId$(Old_YBIATAB0.BIATABTXT, 101, 2) Then blnSeq_Ok = True
            End If
            
            If xSeq < "01" Or xSeq > "19" Then
                blnOk = False
                Call MsgBox("séquence invalide (01-19) ", vbCritical, "BIA_VB_Habilitations : paramétrage")
            Else
                If Not blnSeq_Ok Then
                    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
                        & " where BIATABID = '" & New_YBIATAB0.BIATABID & "' and BIATABK1 = '" & New_YBIATAB0.BIATABK1 & "'" _
                        & " and substring(BIATABTXT,101,2) = '" & xSeq & "'"
                    Call FEU_ROUGE
                    Set rsSab = cnsab.Execute(xSQL)
                    Call FEU_VERT
                    If Not rsSab.EOF Then
                        blnOk = False
                        Call MsgBox("Ce numéro est déjà utilisé ", vbCritical, "BIA_VB_Habilitations : paramétrage")
                    End If
                End If
            End If
        Case Else: blnOk = False

    End Select
End If
If blnOk Then
    If mId$(New_YBIATAB0.BIATABTXT, 100, 1) <> mId$(Old_YBIATAB0.BIATABTXT, 100, 1) Then
        Call MsgBox("ATTENTION : le changement de code (X,S) n'a pas d'impact automatique sur les habilitations" _
          & vbCrLf & " faire la révision par utilisateur.", vbInformation, "BIA_VB_Habilitations : paramétrage")
    End If
End If

fraParam_App_Control = blnOk
End Function

Private Function fraParam_Mnu_Control(lFct As String)
Dim xId As String, xText As String, xHab As String
Dim blnOk As Boolean, K As Integer, Kseq As Integer

blnOk = True
xId = mId$(Trim(txtParam_Mnu_Code) & Space(11), 1, 12)
If Trim(xId) = "" Then
    blnOk = False
    Call MsgBox("Préciser le code ", vbCritical, "BIA_VB_Habilitations : paramétrage")
Else
    xText = mId$(Trim(txtParam_Mnu_Lib) & Space(100), 1, 99)
    If Trim(xId) = "" Then
        blnOk = False
        Call MsgBox("Préciser le libellé ", vbCritical, "BIA_VB_Habilitations : paramétrage")
    End If
End If

'________________________________________________________________
If blnOk Then
    New_YBIATAB0 = Old_YBIATAB0
    New_YBIATAB0.BIATABK2 = mId$(xId, 1, 12)
    New_YBIATAB0.BIATABTXT = ""
    Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = mId$(xText, 1, 69)
    
    For K = 0 To lstParam_Mnu_Droit.ListCount - 1
        If lstParam_Mnu_Droit.Selected(K) Then
            lstParam_Mnu_Droit.ListIndex = K
            xHab = mId$(lstParam_Mnu_Droit.Text, 1, 1)
            If xHab = " " Then xHab = "*"
            Kseq = Val(mId$(lstParam_Mnu_Droit.Text, 2, 2))
            If Kseq > 0 And Kseq < 20 Then Mid$(New_YBIATAB0.BIATABTXT, Kseq, 1) = xHab
        End If
    Next K

End If

fraParam_Mnu_Control = blnOk
End Function



Private Function fraSrv_Control(lFct As String)
Dim xId As String, xText As String, xSeq As String, xSQL As String
Dim blnOk As Boolean, blnSeq_Ok As Boolean

blnOk = True
xId = "S" & Format(txtSrv_Code, "00")
If Trim(txtSrv_Lib1) = "" Then
    blnOk = False
    Call MsgBox("Préciser l'abrégé ", vbCritical, "BIA_VB_Habilitations : paramétrage")
Else
    If Trim(txtSrv_Lib2) = "" Then
        blnOk = False
        Call MsgBox("Préciser le libellé ", vbCritical, "BIA_VB_Habilitations : paramétrage")
    End If
End If

'________________________________________________________________
If blnOk Then
    New_YBIATAB0 = Old_YBIATAB0
    New_YBIATAB0.BIATABK1 = xId
    New_YBIATAB0.BIATABK2 = ""
    Mid$(New_YBIATAB0.BIATABTXT, 1, 12) = Trim(txtSrv_Lib1) & Space(12)
    Mid$(New_YBIATAB0.BIATABTXT, 13, 64) = Trim(txtSrv_Lib2) & Space(64)
End If

fraSrv_Control = blnOk
End Function



Private Sub cmdParam_App_Update_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

If fraParam_App_Control("Update") Then

    If New_YBIATAB0.BIATABK1 <> Old_YBIATAB0.BIATABK1 _
    Or New_YBIATAB0.BIATABK2 <> Old_YBIATAB0.BIATABK2 Then
        Call MsgBox("Le code identifiant a été modifié," & vbCrLf & " la modification n'est pas possible", vbCritical, "BIA_VB_Habilitations : paramétrage")
    Else

        If IsNull(Parametrage_Update) Then
            fraParam_App.Visible = False
            Select Case Trim(Old_YBIATAB0.BIATABID)
                Case "BIA_VB_APP":
                    cmdSelect_Reset
                    cmdSelect_Ok_Click
                Case Else
                    Call fgDetail_Display_SQL("BIA_VB_DROIT", mSelect_App, "")
            End Select
        End If
    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdParam_App_Quit_Click()
fraParam_App.Visible = False
End Sub

Private Sub cmdParam_App_Delete_Click()
Dim X As String, xSQL As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

blnOk = True
X = Trim(txtParam_APP_Code)
Select Case Trim(Old_YBIATAB0.BIATABID)
    Case "BIA_VB_APP":
        If X <> Trim(Old_YBIATAB0.BIATABK1) Then
            blnOk = False
            Call MsgBox("Le code APPLICATION a été modifié," & vbCrLf & " la suppression n'est pas possible", vbCritical, "BIA_VB_Habilitations : paramétrage")
        End If
        xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB' and BIATABK1 = '" & X & "'"
        Set rsSab = cnsab.Execute(xSQL)
        If rsSab(0) > 0 Then
            blnOk = False
            Call MsgBox("Il y a " & rsSab(0) & " utilisateurs habilités ," & vbCrLf & " la suppression n'est pas possible : " & X, vbCritical, "BIA_VB_Habilitations : suppression")
        End If
        xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_DROIT' and BIATABK1 = '" & X & "'"
        Set rsSab = cnsab.Execute(xSQL)
        If rsSab(0) > 0 Then
            blnOk = False
            Call MsgBox("Il y a " & rsSab(0) & " droits associés ," & vbCrLf & " la suppression n'est pas possible : " & X, vbCritical, "BIA_VB_Habilitations : suppression")
        End If
    Case "BIA_VB_DROIT"
        If X <> Trim(Old_YBIATAB0.BIATABK2) Then
            blnOk = False
            Call MsgBox("Le code identifiant a été modifié," & vbCrLf & " la suppression n'est pas possible", vbCritical, "BIA_VB_Habilitations : paramétrage")
        End If
    Case Else: blnOk = False
End Select


If blnOk Then
    If IsNull(Parametrage_Delete) Then
        Select Case Trim(Old_YBIATAB0.BIATABID)
            Case "BIA_VB_APP":
                cmdSelect_Reset
                cmdSelect_Ok_Click
            Case Else
                Call fgDetail_Display_SQL("BIA_VB_DROIT", mSelect_App, "")
        End Select
        fraParam_App.Visible = False
    End If
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Return()
    If SSTab1.Tab = 0 Then
        If fgSelect.Visible = False Then cmdSelect_Ok_Click
    Else
        SendKeys "{TAB}"
    End If
End Sub


Public Sub cmdContext_Quit()
lstErr.Clear: lstErr.Height = 200


If txtRTF.Visible Then
    txtRTF.Visible = False
    Exit Sub
End If

If txtFg.Visible Then
    txtFg.Visible = False
    Exit Sub
End If
If fraParam_Mnu.Visible Then fraParam_Mnu.Visible = False: Exit Sub

If fraSrv.Visible Then
    fraSrv.Visible = False
    Exit Sub
End If

If fraUsr.Visible Then
    fraUsr.Visible = False
    Exit Sub
End If


If fraParam_Hab.Visible Then
    fraParam_Hab.Visible = False
    Exit Sub
End If
If fraParam_App.Visible Then
    fraParam_App.Visible = False
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
Call lstErr_Clear(lstErr, cmdContext, "> BIA_VB_Habilitations_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Clear

Select Case cmdSelect_SQL_K
    Case "1", "1L": cmdSelect_SQL_1
    Case "1D*U": cmdSelect_SQL_1DU
    Case "2", "2L": cmdSelect_SQL_2
    Case "3": cmdSelect_SQL_3
    Case "3#": cmdSelect_SQL_3_VB_SAB
    Case "4": cmdSelect_SQL_4
    Case "9App", "9Mnu": cmdSelect_SQL_1
    Case "XHab": cmdSelect_SQL_Exportation
    Case "9Usr":  fraUsr_Srv_Init
    Case "SI_Doc": cmdSelect_SQL_SI_Doc
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_VB_Habilitations_cmdSelect_Ok"): DoEvents
lstErr.Height = 480
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
cmdSelect_Ok.BackColor = fgSelect.BackColorFixed
End Sub




Public Sub param_Init()
Dim X As String, K As Integer, blnOk As Boolean

usrName_UCase = "Reprise"

Exit Sub

'____________________________________________________________________________________
Dim recBiaPgm As typeElpTable, recBiaPgmAut As typeElpTable

rsElpTable_Init recBiaPgm

X = "select * from ElpTable where SNN = 0" _
    & " and id = 'BiaPgm' order by K1"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    Call rsElpTable_GetBuffer(rsMDB, recBiaPgm)
    New_YBIATAB0.BIATABID = "BIA_VB_APP"
    New_YBIATAB0.BIATABK1 = UCase$(recBiaPgm.K1)
    New_YBIATAB0.BIATABK2 = ""
    New_YBIATAB0.BIATABTXT = UCase$(recBiaPgm.Name)
    Mid$(New_YBIATAB0.BIATABTXT, 80, 12) = mId$(recBiaPgm.Memo, 21, 20)
    If Trim(New_YBIATAB0.BIATABK1) = "SAB_BALANCE" Then
        New_YBIATAB0.BIATABTXT = "balance, extraits de compte, RIB ..."
        Mid$(New_YBIATAB0.BIATABTXT, 80, 12) = "BIA_SAB"
    End If
    
    If Trim(New_YBIATAB0.BIATABK1) = "XUSRID" Then
        New_YBIATAB0.BIATABTXT = "identification user"
        Mid$(New_YBIATAB0.BIATABTXT, 80, 12) = ""
        Mid$(New_YBIATAB0.BIATABTXT, 100, 1) = "X"
    End If
    If mId$(New_YBIATAB0.BIATABK1, 1, 1) = "@" Then Mid$(New_YBIATAB0.BIATABTXT, 100, 1) = "X"
    If Not IsNull(Parametrage_New) Then
        Call MsgBox(Error, vbCritical, "param_init : " & New_YBIATAB0.BIATABID & New_YBIATAB0.BIATABK1)
    End If

    New_YBIATAB0.BIATABID = "BIA_VB_DROIT"
    New_YBIATAB0.BIATABTXT = "accès à l'application " & New_YBIATAB0.BIATABK1
    Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "#01"
    If Not IsNull(Parametrage_New) Then
        Call MsgBox(Error, vbCritical, "param_init : " & New_YBIATAB0.BIATABID & New_YBIATAB0.BIATABK1 & New_YBIATAB0.BIATABK2)
    End If

    For K = 2 To 9
        If mId$(recBiaPgm.Memo, K, 1) = "X" Then
            Select Case K
                Case "2":
                    New_YBIATAB0.BIATABK2 = "MAJ"
                    New_YBIATAB0.BIATABTXT = "Mise à jour"
                    Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "S02"
                Case "3":
                    New_YBIATAB0.BIATABK2 = "Validation"
                    New_YBIATAB0.BIATABTXT = "Validation"
                    Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "S03"
                Case "4":
                    New_YBIATAB0.BIATABK2 = "4C"
                    New_YBIATAB0.BIATABTXT = "4C"
                    Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "*04"
                Case "5":
                    New_YBIATAB0.BIATABK2 = "5R"
                    New_YBIATAB0.BIATABTXT = "Rapprochement"
                    Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "*05"
                Case "6":
                    New_YBIATAB0.BIATABK2 = "6S"
                    New_YBIATAB0.BIATABTXT = "6S"
                    Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "*06"
                Case "7":
                    New_YBIATAB0.BIATABK2 = "7V"
                    New_YBIATAB0.BIATABTXT = "7V"
                    Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "*07"
                Case "8":
                    New_YBIATAB0.BIATABK2 = "8A"
                    New_YBIATAB0.BIATABTXT = "8A"
                    Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "*08"
                Case "9":
                    New_YBIATAB0.BIATABK2 = "Admin"
                    New_YBIATAB0.BIATABTXT = "réservé administrateur"
                    Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "X19"
                'Case "10":
                '    New_YBIATAB0.BIATABK2 = "10M"
                '    New_YBIATAB0.BIATABTXT = "10M"
                '    Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "X10"
            
            End Select
            'If blnOk Then
                If Not IsNull(Parametrage_New) Then
                    Call MsgBox(Error, vbCritical, "param_init : " & New_YBIATAB0.BIATABID & New_YBIATAB0.BIATABK1 & New_YBIATAB0.BIATABK2)
                End If
            'End If
            
        End If
    Next K
    
    rsMDB.MoveNext
Loop

'____________________________________________________________________________________


X = "select * from ElpTable where SNN = 0" _
    & " and id = 'BiaPgm_Aut'" _
    & " order by K1 , K2"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    Call rsElpTable_GetBuffer(rsMDB, recBiaPgmAut)
    
    X = Trim(recBiaPgmAut.K2)
    If X = "$usr_Forçage" Or X = "$usr_Service" Or X = "X_I5A7" Then
    
    Else
        New_YBIATAB0.BIATABID = "BIA_VB_HAB"
        New_YBIATAB0.BIATABK1 = UCase$(recBiaPgmAut.K2)
        New_YBIATAB0.BIATABK2 = UCase$(recBiaPgmAut.K1)
        New_YBIATAB0.BIATABTXT = "#"
        
        If Trim(New_YBIATAB0.BIATABK2) <> mSelect_Usr Then
            mSelect_Usr = Trim(New_YBIATAB0.BIATABK2)
            lstParam_Usr_Srv
        End If
        Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = mSelect_Usr_Srv
    
        For K = 2 To 8
            If mId$(recBiaPgmAut.Memo, K, 1) = "X" Then Mid$(New_YBIATAB0.BIATABTXT, K, 1) = "*"
        Next K
        If mId$(recBiaPgmAut.Memo, 9, 1) = "X" Then Mid$(New_YBIATAB0.BIATABTXT, 19, 1) = "X"
        
        
        If Trim(New_YBIATAB0.BIATABK1) = "DROPI" Then Mid$(New_YBIATAB0.BIATABTXT, 3, 4) = "    "
        If Trim(New_YBIATAB0.BIATABK1) = "SAB_Compta" Then Mid$(New_YBIATAB0.BIATABTXT, 2, 7) = "       "
        If Trim(New_YBIATAB0.BIATABK1) = "SAB_Ordonnan" Then Mid$(New_YBIATAB0.BIATABTXT, 2, 7) = "       "
        
        If Not IsNull(Parametrage_New) Then
            Call MsgBox(Error, vbCritical, "param_init : " & New_YBIATAB0.BIATABID & New_YBIATAB0.BIATABK1 & New_YBIATAB0.BIATABK2)
        End If
     End If
     
            
    rsMDB.MoveNext
Loop

X = "delete from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID like 'BIA_VB_%' and BIATABK1 = 'BIA_VB_HAB'"
Call FEU_ROUGE
Call Parametrage_SQL(X)
X = "delete from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID like 'BIA_VB_%' and BIATABK1 = 'YSWISAB0'"
Call Parametrage_SQL(X)
Call FEU_VERT
'=====================================================================================================================
New_YBIATAB0.BIATABID = "BIA_VB_APP"
New_YBIATAB0.BIATABK1 = "BIA_VB_HAB"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "gestion des habilitations des applications BIA.vbp"
Mid$(New_YBIATAB0.BIATABTXT, 80, 12) = "BIA_SYSTEM"
Mid$(New_YBIATAB0.BIATABTXT, 100, 1) = "X"
Parametrage_New

New_YBIATAB0.BIATABID = "BIA_VB_DROIT"
New_YBIATAB0.BIATABK1 = "BIA_VB_HAB"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "accès à l'application BIA_VB_HAB"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "#01"
Parametrage_New

New_YBIATAB0.BIATABID = "BIA_VB_DROIT"
New_YBIATAB0.BIATABK1 = "BIA_VB_HAB"
New_YBIATAB0.BIATABK2 = "Admin"
New_YBIATAB0.BIATABTXT = "habilitation administrateur"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "X19"
Parametrage_New

New_YBIATAB0.BIATABID = "BIA_VB_DROIT"
New_YBIATAB0.BIATABK1 = "BIA_VB_HAB"
New_YBIATAB0.BIATABK2 = "Inspection"
New_YBIATAB0.BIATABTXT = "habilitation Inspection"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "X09"
Parametrage_New

New_YBIATAB0.BIATABID = "BIA_VB_DROIT"
New_YBIATAB0.BIATABK1 = "BIA_VB_HAB"
New_YBIATAB0.BIATABK2 = "Paramétrage"
New_YBIATAB0.BIATABTXT = "Paramétrage des Applications, droits, menus"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "X18"
Parametrage_New

New_YBIATAB0.BIATABID = "BIA_VB_DROIT"
New_YBIATAB0.BIATABK1 = "BIA_VB_HAB"
New_YBIATAB0.BIATABK2 = "MAJ"
New_YBIATAB0.BIATABTXT = "MAJ des habilitations (Utilisateur / Application)"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "*02"
Parametrage_New


'New_YBIATAB0.BIATABID = "BIA_VB_DROIT"
'New_YBIATAB0.BIATABK1 = "BIA_VB_HAB"
'New_YBIATAB0.BIATABK2 = "XHab"
'New_YBIATAB0.BIATABTXT = "exportation des habilitations"
'Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = " 03"
'Parametrage_New


New_YBIATAB0.BIATABID = "BIA_VB_HAB"
New_YBIATAB0.BIATABK1 = "BIA_VB_HAB"
New_YBIATAB0.BIATABK2 = "LOULERGUE"
New_YBIATAB0.BIATABTXT = "#*"
Mid$(New_YBIATAB0.BIATABTXT, 18, 2) = "XX"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "S40"
Parametrage_New
'_______________________________________________________________________________________
New_YBIATAB0.BIATABID = "BIA_VB_MNU"
New_YBIATAB0.BIATABK1 = "BIA_VB_HAB"
New_YBIATAB0.BIATABK2 = "1"
New_YBIATAB0.BIATABTXT = "#"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "sélection droit => Utilisateurs"
Parametrage_New

New_YBIATAB0.BIATABK2 = "1L"
New_YBIATAB0.BIATABTXT = "#"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "liste des utilisateurs habilités à ...."
Parametrage_New

New_YBIATAB0.BIATABK2 = "2"
New_YBIATAB0.BIATABTXT = "#"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "sélection Utilisateur => droit"
Parametrage_New

New_YBIATAB0.BIATABK2 = "2L"
New_YBIATAB0.BIATABTXT = "#"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "liste des habilitations d'un utilisateur"
Parametrage_New

New_YBIATAB0.BIATABK2 = "XHab"
New_YBIATAB0.BIATABTXT = "#"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "exportation des habilitations"
Parametrage_New

New_YBIATAB0.BIATABK2 = "3"
New_YBIATAB0.BIATABTXT = "#*"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "suppression de TOUTES les habilitations d'un utilisateur"
Parametrage_New

New_YBIATAB0.BIATABK2 = "4"
New_YBIATAB0.BIATABTXT = "#*"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "Duplication de TOUTES les habilitations d'un utilisateur"
Parametrage_New


New_YBIATAB0.BIATABK2 = "3#"
New_YBIATAB0.BIATABTXT = "#*"
Mid$(New_YBIATAB0.BIATABTXT, 18, 1) = "X"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "habilitations des utilisateurs à supprimer ?"
Parametrage_New

New_YBIATAB0.BIATABK2 = "9App"
New_YBIATAB0.BIATABTXT = "#*"
Mid$(New_YBIATAB0.BIATABTXT, 18, 1) = "X"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "gestion des applications et droits"
Parametrage_New

New_YBIATAB0.BIATABK2 = "9Usr"
New_YBIATAB0.BIATABTXT = "#*"
Mid$(New_YBIATAB0.BIATABTXT, 18, 1) = "X"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "gestion des utilisateurs et des services"
Parametrage_New

New_YBIATAB0.BIATABK2 = "9Mnu"
New_YBIATAB0.BIATABTXT = "#*"
Mid$(New_YBIATAB0.BIATABTXT, 18, 1) = "X"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "gestion des options de menu"
Parametrage_New
'=================================================================================


usrName_UCase = "LOULERGUE"
End Sub

Public Sub fraParam_App_Display(lId As String, lK1 As String, lK2 As String)
Dim V, X As String, blnOk As Boolean
fraParam_App.Visible = False

fraParam_App.BackColor = fgDetail.BackColorFixed

blnOk = True
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & "where BIATABID = '" & lId & "' and BIATABK1 = '" & lK1 & "' and BIATABK2 = '" & lK2 & "'"
    
Set rsSab = cnsab.Execute(X)

If rsSab.EOF Then
    If lId = "BIA_VB_DROIT" And lK2 = "" Then
        Old_YBIATAB0.BIATABID = lId
        Old_YBIATAB0.BIATABK1 = lK1
        Old_YBIATAB0.BIATABK2 = ""
        Old_YBIATAB0.BIATABTXT = ""
        
    Else
        blnOk = False
        Call MsgBox("Erreur lecture : " & X, vbCritical, "fraParam_App_Display")
    End If
End If

If blnOk Then
    Call rsYBIATAB0_GetBuffer(rsSab, Old_YBIATAB0)
    If mId$(Old_YBIATAB0.BIATABTXT, 100, 1) = "#" Then
        If fgDetail.Rows = 2 Then
            cmdParam_APP_Delete.Visible = True
        Else
            cmdParam_APP_Delete.Visible = False
        End If
        cmdParam_App_Update.Visible = False
    Else
        cmdParam_APP_Delete.Visible = True
        cmdParam_App_Update.Visible = True
    End If
    
    Select Case lId
        Case "BIA_VB_APP":
            App_YBIATAB0 = Old_YBIATAB0
            txtParam_APP_Seq.Visible = False: lblParam_APP_Seq.Visible = False
            cboParam_APP_VBP.Visible = True: lblParam_APP_VBP.Visible = True
            txtParam_APP_Doc.Visible = True: lblParam_APP_Doc.Visible = True
            fraParam_App.Caption = "Gestion des applications"
            lblParam_APP_Code = "Application"
            lblParam_APP_Code.BackColor = fgSelect.BackColorFixed
            txtParam_APP_Code = lK1
            txtParam_App_Lib = Trim(mId$(Old_YBIATAB0.BIATABTXT, 1, 69))
            Call cbo_Scan(Trim(mId$(Old_YBIATAB0.BIATABTXT, 80, 12)), cboParam_APP_VBP)
            txtParam_APP_Doc = Trim(mId$(Old_YBIATAB0.BIATABTXT, 70, 10))
            optParam_App_Z.Value = True
            optParam_App_X.Value = IIf(mId$(Old_YBIATAB0.BIATABTXT, 100, 1) = "X", True, False)
            optParam_App_S.Visible = False
            fraParam_App.Top = fgSelect.Top
            fraParam_App.BackColor = fgSelect.BackColorFixed
            cmdParam_App_Add_17.Visible = False
            cmdParam_App_Add_18.Visible = False
            cmdParam_App_Add_19.Visible = False

        Case "BIA_VB_DROIT"
            Fct_YBIATAB0 = Old_YBIATAB0
            txtParam_APP_Seq.Visible = True: lblParam_APP_Seq.Visible = True
            cboParam_APP_VBP.Visible = False: lblParam_APP_VBP.Visible = True
            txtParam_APP_Doc.Visible = False: lblParam_APP_Doc.Visible = True
            fraParam_App.Caption = "Gestion des droits : " & lK1
            lblParam_APP_Code = "Droit"
            lblParam_APP_Code.BackColor = fgDetail.BackColorFixed
            txtParam_APP_Code = lK2
            txtParam_App_Lib = Trim(mId$(Old_YBIATAB0.BIATABTXT, 1, 69))
            txtParam_APP_Seq = Trim(mId$(Old_YBIATAB0.BIATABTXT, 101, 2))
            optParam_App_Z.Value = True
            optParam_App_X.Value = IIf(mId$(Old_YBIATAB0.BIATABTXT, 100, 1) = "X", True, False)
            optParam_App_S.Value = IIf(mId$(Old_YBIATAB0.BIATABTXT, 100, 1) = "S", True, False)
            optParam_App_S.Visible = True
           fraParam_App.Top = fgDetail.Top
           fraParam_App.BackColor = fgDetail.BackColorFixed
            cmdParam_App_Add_17.Visible = True
            cmdParam_App_Add_18.Visible = True
            cmdParam_App_Add_19.Visible = True


    End Select
    lblParam_quid = " MAJ : " & mId$(Old_YBIATAB0.BIATABTXT, 105, 10) & " " & dateImp10_S(mId$(Old_YBIATAB0.BIATABTXT, 115, 8)) & " " & timeImp8(mId$(Old_YBIATAB0.BIATABTXT, 123, 6))
    fraParam_App.Visible = True
End If

End Sub
Public Sub fraParam_Mnu_Display(lId As String, lK1 As String, lK2 As String)
Dim V, X As String, blnOk As Boolean, K As Integer
fraParam_Mnu.Visible = False

fraParam_Mnu.BackColor = fgDetail.BackColorFixed
cmdParam_Mnu_Delete.Visible = False
cmdParam_Mnu_Update.Visible = False

blnOk = True
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & "where BIATABID = '" & lId & "' and BIATABK1 = '" & lK1 & "' and BIATABK2 = '" & lK2 & "'"
    
Set rsSab = cnsab.Execute(X)

If rsSab.EOF Then
    If lId = "BIA_VB_MNU" And lK2 = "" Then
        Old_YBIATAB0.BIATABID = lId
        Old_YBIATAB0.BIATABK1 = lK1
        Old_YBIATAB0.BIATABK2 = ""
        Old_YBIATAB0.BIATABTXT = "#"
        
    Else
        blnOk = False
        Call MsgBox("Erreur lecture : " & X, vbCritical, "fraParam_Mnu_Display")
    End If
Else
    Call rsYBIATAB0_GetBuffer(rsSab, Old_YBIATAB0)
    cmdParam_Mnu_Delete.Visible = True
    cmdParam_Mnu_Update.Visible = True
End If

If blnOk Then
    Fct_YBIATAB0 = Old_YBIATAB0
    fraParam_Mnu.Caption = "Gestion des menus : " & lK1
    libParam_Mnu_Code = "Code option de menu"
    libParam_Mnu_Code.BackColor = fgDetail.BackColorFixed
    txtParam_Mnu_Code = lK2
    txtParam_Mnu_Lib = Trim(mId$(Old_YBIATAB0.BIATABTXT, 20, 79))
    fraParam_Mnu.BackColor = fgDetail.BackColorFixed

    lblParam_Mnu_Quid = " MAJ : " & mId$(Old_YBIATAB0.BIATABTXT, 105, 10) & " " & dateImp10_S(mId$(Old_YBIATAB0.BIATABTXT, 115, 8)) & " " & timeImp8(mId$(Old_YBIATAB0.BIATABTXT, 123, 6))
'_________________________________________________________________________________________________________
    For K = 1 To 19: arrFct_X(K) = "": arrFct_ListIndex(K) = -1: Next K
    
    lstParam_Mnu_Droit.Clear
    X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
        & " where BIATABID = 'BIA_VB_DROIT' and BIATABK1 = '" & mSelect_App & "' order by substring(BIATABTXT,101,2)"
    Set rsSab = cnsab.Execute(X)
    
    Do While Not rsSab.EOF
        blnOk = True
        X = rsSab("BIATABTXT")
        lstParam_Mnu_Droit.AddItem mId$(X, 100, 3) & " " & rsSab("BIATABK2") & " : " & X
        arrFct_X(mId$(X, 101, 2)) = mId$(X, 100, 1)
        arrFct_ListIndex(mId$(X, 101, 2)) = lstParam_Mnu_Droit.ListCount - 1
        rsSab.MoveNext
    Loop
'_______________________________________________________________________________

        X = Old_YBIATAB0.BIATABTXT
        For K = 1 To 19
            If mId$(X, K, 1) <> " " Then
                If arrFct_ListIndex(K) >= 0 And arrFct_ListIndex(K) < 19 Then lstParam_Mnu_Droit.Selected(arrFct_ListIndex(K)) = True
            End If
        Next K
    
'_______________________________________________________________________________
    fraParam_Mnu.Visible = True
End If

End Sub


Private Sub lstParam_Hab_Click()
blnParam_Hab_Change = True
End Sub

Private Sub lstParam_Hab_KeyPress(KeyAscii As Integer)
blnParam_Hab_Change = True
End Sub


Private Sub lstSrv_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Old_YBIATAB0.BIATABID = "ROPDOSISRV"
Old_YBIATAB0.BIATABK1 = mId$(lstSrv.Text, 1, 3)
Old_YBIATAB0.BIATABK2 = ""
If IsNull(sqlYBIATAB0_Read(Old_YBIATAB0.BIATABID, Old_YBIATAB0.BIATABK1, Old_YBIATAB0.BIATABK2, Old_YBIATAB0.BIATABTXT)) Then
    txtSrv_Code = mId$(Old_YBIATAB0.BIATABK1, 2, 2)
    txtSrv_Lib1 = Trim(mId$(Old_YBIATAB0.BIATABTXT, 1, 12))
    txtSrv_Lib2 = Trim(mId$(Old_YBIATAB0.BIATABTXT, 13, 64))
    fraSrv.Visible = True
End If

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub lstUsr_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
'Old_YBIATAB0.BIATABID = "ROPDOSGUSR"
'Old_YBIATAB0.BIATABK1 = mId$(lstUsr.Text, 1, 10)
'Old_YBIATAB0.BIATABK2 = ""
'txtUsr_Code.Enabled = False
'txtUsr_Code = Trim(Old_YBIATAB0.BIATABK1)

'If IsNull(sqlYBIATAB0_Read(Old_YBIATAB0.BIATABID, Old_YBIATAB0.BIATABK1, Old_YBIATAB0.BIATABK2, Old_YBIATAB0.BIATABTXT)) Then
'    Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 26, 3), cboUsr_Srv)
'Else
'   Call cbo_Scan("S99", cboUsr_Srv)
'End If
Dim K As Integer
K = InStr(lstUsr.Text, "-")
If K > 0 Then
    txtUsr_Code = mId$(lstUsr.Text, 1, K - 1)
    Call cbo_Scan(mId$(lstUsr.Text, K + 2, 3), cboUsr_Srv)

End If
libUsr_Code = lstUsr.Text
fraUsr.Visible = True


Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub lstUsr_VB_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass
Old_YBIATAB0.BIATABID = "ROPDOSGUSR"
Old_YBIATAB0.BIATABK1 = mId$(lstUsr_VB.Text, 1, 12)
Old_YBIATAB0.BIATABK2 = ""
txtUsr_Code.Enabled = True
txtUsr_Code = Trim(Old_YBIATAB0.BIATABK1)
'K = InStr(lstUsr_VB.Text, "-")

If IsNull(sqlYBIATAB0_Read(Old_YBIATAB0.BIATABID, Old_YBIATAB0.BIATABK1, Old_YBIATAB0.BIATABK2, Old_YBIATAB0.BIATABTXT)) Then
    Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 26, 3), cboUsr_Srv)
Else

    Call cbo_Scan("S99", cboUsr_Srv)
End If
libUsr_Code = lstUsr_VB.Text
fraUsr.Visible = True


Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub lstUsr_VB_No_Click()
Dim K As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass
Old_YBIATAB0.BIATABID = "ROPDOSGUSR"
K = InStr(lstUsr_VB_No.Text, ":")
Old_YBIATAB0.BIATABK1 = mId$(lstUsr_VB_No.Text, 1, K - 1)
Old_YBIATAB0.BIATABK2 = ""
txtUsr_Code.Enabled = False
txtUsr_Code = Trim(Old_YBIATAB0.BIATABK1)
'K = InStr(lstUsr_VB.Text, "-")

If IsNull(sqlYBIATAB0_Read(Old_YBIATAB0.BIATABID, Old_YBIATAB0.BIATABK1, Old_YBIATAB0.BIATABK2, Old_YBIATAB0.BIATABTXT)) Then
    Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 26, 3), cboUsr_Srv)
Else

    Call cbo_Scan("S99", cboUsr_Srv)
End If
libUsr_Code = lstUsr_VB_No.Text
fraUsr.Visible = True


Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub txtParam_APP_Code_GotFocus()
txtParam_APP_Code.BackColor = focusUsr.BackColor

End Sub

Private Sub txtParam_APP_Code_KeyPress(KeyAscii As Integer)
    If Trim(Old_YBIATAB0.BIATABID) = "BIA_VB_APP" Then KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtParam_APP_Code_LostFocus()
txtParam_APP_Code.BackColor = txtUsr.BackColor
End Sub

Private Sub txtParam_App_Lib_GotFocus()
txtParam_App_Lib.BackColor = focusUsr.BackColor

End Sub

Private Sub txtParam_App_Lib_LostFocus()
txtParam_App_Lib.BackColor = txtUsr.BackColor
End Sub

Private Sub txtParam_APP_Seq_GotFocus()
txtParam_APP_Seq.BackColor = focusUsr.BackColor

End Sub

Private Sub txtParam_APP_Seq_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub

Private Sub txtParam_APP_Seq_LostFocus()
txtParam_APP_Seq.BackColor = txtUsr.BackColor
End Sub


Public Sub Parametrage_Load_App()
Dim xSQL As String


cboSelect_App.Clear
If blnAdmin Then
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_APP' order by BIATABK1"
Else
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_APP' " _
      & " and BIATABK1 in (select distinct(BIATABK1) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB'" _
      & " and substring(BIATABTXT,100,3) = '" & currentSSIWINUNIT & "'" & ")" & "order by BIATABK1"
End If

Set rsSab = cnsab.Execute(xSQL)
 
Do While Not rsSab.EOF
    cboSelect_App.AddItem Trim(rsSab("BIATABK1"))
    rsSab.MoveNext

Loop

End Sub

Public Sub cmdSelect_SQL_4_Init()
Dim K As Integer

cboSelect_4_Usr1.Clear
cboSelect_4_Usr2.Clear
blnControl = False

For K = 0 To cboSelect_Usr.ListCount - 1
    cboSelect_Usr.ListIndex = K
    cboSelect_4_Usr1.AddItem cboSelect_Usr.Text
    cboSelect_4_Usr2.AddItem cboSelect_Usr.Text
Next K

cboSelect_4_APP.Clear
For K = 0 To cboSelect_App.ListCount - 1
    cboSelect_App.ListIndex = K
    cboSelect_4_APP.AddItem cboSelect_App.Text
Next K

blnControl = True
End Sub

Public Sub lstParam_Usr_Srv()
Dim xSQL As String
'_________________________________________________________________________________________________
     
xSQL = "select *from " & paramIBM_Library_SABSPE & ".YSSIDOM0 where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN'" _
     & "  and SSIDOMUIDX = '" & mSelect_Usr & "'"


Set rsSabX = cnsab.Execute(xSQL)
If rsSabX.EOF Then
    'Call MsgBox("Erreur de lecture : " & xsql, vbCritical, "lstParam_Usr_Srv")
    Select Case mSelect_Usr
        Case "BIA_INFO", "BIA_AUTO", "BIA_SWIFT": mSelect_Usr_Srv = "S40"
        Case Else: mSelect_Usr_Srv = "S99"
    End Select
Else
    mSelect_Usr_Srv = rsSabX("SSIDOMUNIT")
End If

End Sub




Public Sub fraUsr_Srv_Init()
Dim X As String, K As Integer, blnOk As Boolean

lstW.Clear
lstUsr.Clear: lstUsr.BackColor = &HF0FFFF
lstUsr_VB.Clear: lstUsr_VB.BackColor = &HE0EFFF
lstUsr_VB_No.Clear: lstUsr_VB_No.BackColor = &HFFE0FF

'Call lstZMNURUT0_Load_Actif_Production(YSSIUSR0_Actif_Load
Call YSSIUSR0_Actif_Load(lstW)

'load Usr_VB ___________________________________________________________________________
ReDim arrUsr(lstW.ListCount + 1), arrUsr_code(lstW.ListCount + 1), arrUsr_lib(lstW.ListCount + 1)
arrUsr_Nb = 0

For K = 0 To lstW.ListCount - 1
    lstW.ListIndex = K
    arrUsr_Nb = arrUsr_Nb + 1
    arrUsr_code(K) = mId$(lstW.Text, 1, 10)
    arrUsr(K) = UCase$(Trim(arrUsr_code(K)))
    arrUsr_lib(K) = mId$(lstW.Text, 13, Len(lstW.Text) - 13)
Next K

lstUsr_VB.Clear
X = "select *from " & paramIBM_Library_SABSPE & ".YSSIDOM0 where SSIDOMNAT = ' ' and SSIDOMDIDX = 'SAB'" _
  & " and SSIDOMPRFK <> 'X' order by SSIDOMUIDX"
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    X = UCase$(Trim(rsSab("SSIDOMUIDX")))
    blnOk = False
    For K = 0 To lstW.ListCount - 1
        If arrUsr(K) = X Then
            lstUsr.AddItem arrUsr_code(K) & vbTab & "- " & rsSab("SSIDOMUNIT") & " : " & arrUsr_lib(K)
            arrUsr(K) = ""
            blnOk = True: Exit For
        End If
    Next K

    If Not blnOk Then lstUsr_VB_No.AddItem X & " : " & rsSab("SSIDOMUNIT")
    rsSab.MoveNext
Loop

For K = 0 To arrUsr_Nb
    If arrUsr(K) <> "" Then
        lstUsr_VB.AddItem arrUsr_code(K) & " : " & arrUsr_lib(K)
    End If
Next K


'load Service ___________________________________________________________________________
lstSrv.Clear
lstSrv.BackColor = &HC0E0FF
cboUsr_Srv.Clear

X = "select *from " & paramIBM_Library_SABSPE & ".YSSIUSR0 where SSIUSRNAT= 'S' order by SSIUSRUNIT"
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    X = rsSab("SSIUSRUNIT") & " - " & rsSab("SSIUSRUIDX")
    lstSrv.AddItem X
    cboUsr_Srv.AddItem X
    rsSab.MoveNext
Loop

'______________________________________________________________________

fraUsr.Visible = False
fraSrv.Visible = False
fraUsr_Srv.Visible = True
SSTab1.Tab = 1
End Sub

Private Sub txtSrv_Code_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub



Public Sub cmdSelect_SQL_3_VB_SAB()
    
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_3#"
   
mSelect_App = "": mSelect_Usr = "": mSelect_Usr_Srv = ""

xWhere = ""


xSQL = "select distinct(BIATABK2) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB' " _
  & " and BIATABK2 in (select distinct(MNURUTUTI) from " & paramIBM_Library_SAB_P & ".ZMNURUT0 , " & paramIBM_Library_SAB_P & ".ZMNUUTI0" _
  & " where MNUUTICUT = MNURUTCUT and MNUUTIGR2 = 'G_MIN') order by BIATABK2"
    
Set rsSab = cnsab.Execute(xSQL)
  
fgSelect_3_Display "G_MIN"

'____________________________________________________________________________________________________
xSQL = "select distinct(BIATABK2) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB' " _
  & " and BIATABK2 not in (select distinct(MNURUTUTI) from " & paramIBM_Library_SAB_P & ".ZMNURUT0 , " & paramIBM_Library_SAB_P & ".ZMNUUTI0" _
  & " where MNUUTICUT = MNURUTCUT) order by BIATABK2"
    
Set rsSab = cnsab.Execute(xSQL)
  
fgSelect_3_Display "utilisateur ayant des habilitations VB et inconnu dans SAB"

'____________________________________________________________________________________________________
xSQL = "select distinct(BIATABK2) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_HAB' " _
  & " and substring(BIATABTXT,100,3) = 'S99'"
    
Set rsSab = cnsab.Execute(xSQL)
  
fgSelect_3_Display "utilisateur ayant des habilitations VB affectées au service S99"

Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub param_Init_SAB_DOSSIER()
Dim X As String
'=====================================================================================================================
Call param_Init_Delete("SAB_DOSSIER")

'New_YBIATAB0.BIATABK2 = "MAJ"
'New_YBIATAB0.BIATABTXT = "Mise à jour"
'Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "S02"
'Parametrage_New

New_YBIATAB0.BIATABK2 = "Administration reprise CREDOC"
New_YBIATAB0.BIATABTXT = "Validation"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "S03"
Parametrage_New

New_YBIATAB0.BIATABK2 = "DER"
New_YBIATAB0.BIATABTXT = "Etats des engagements"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "S04"
Parametrage_New

New_YBIATAB0.BIATABK2 = "CREDOC"
New_YBIATAB0.BIATABTXT = "Etats CREDOC"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "S05"
Parametrage_New
'=====================================================================================================================

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'BIA_VB_HAB' and BIATABK1 = 'SAB_DOSSIER' order by BIATABK2"
        
Set rsSabX = cnsab.Execute(X)
Do While Not rsSabX.EOF
    Call rsYBIATAB0_GetBuffer(rsSabX, Old_YBIATAB0)
    New_YBIATAB0 = Old_YBIATAB0
    Mid$(New_YBIATAB0.BIATABTXT, 2, 18) = Space$(18)
    Parametrage_Update

    rsSabX.MoveNext
Loop
'_______________________________________________________________________________________
New_YBIATAB0.BIATABID = "BIA_VB_MNU"
New_YBIATAB0.BIATABK1 = "SAB_DOSSIER"

New_YBIATAB0.BIATABK2 = "1"
New_YBIATAB0.BIATABTXT = "#"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "Dossier"
Parametrage_New

New_YBIATAB0.BIATABK2 = "2"
New_YBIATAB0.BIATABTXT = "#"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "Client"
Parametrage_New

New_YBIATAB0.BIATABK2 = "3"
New_YBIATAB0.BIATABTXT = "#   S"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "Evénements en attente de validation"
Parametrage_New

New_YBIATAB0.BIATABK2 = "5"
New_YBIATAB0.BIATABTXT = "#   S"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "Code état"
Parametrage_New

New_YBIATAB0.BIATABK2 = "6"
New_YBIATAB0.BIATABTXT = "#   S"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "Etat des provisions (PCI-Client-Dossier)"
Parametrage_New

New_YBIATAB0.BIATABK2 = "X#"
New_YBIATAB0.BIATABTXT = "#   S"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "Etat de surveillance CREDOC (.xls)"
Parametrage_New

New_YBIATAB0.BIATABK2 = "Xc"
New_YBIATAB0.BIATABTXT = "#   S"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "Commissions CREDOC à recevoir (.xls)"
Parametrage_New

New_YBIATAB0.BIATABK2 = "Xi"
New_YBIATAB0.BIATABTXT = "#  S"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "Engagements CREDOC Intragroupe (.xls)"
Parametrage_New

New_YBIATAB0.BIATABK2 = "XE1an"
New_YBIATAB0.BIATABTXT = "#  S"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "Engagements CREDOC +/- 1 an (.xls)"
Parametrage_New

New_YBIATAB0.BIATABK2 = "2#"
New_YBIATAB0.BIATABTXT = "# S"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "Affectation des OD"
Parametrage_New

New_YBIATAB0.BIATABK2 = "zOD"
New_YBIATAB0.BIATABTXT = "# S"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "Liste des OD"
Parametrage_New

New_YBIATAB0.BIATABK2 = "zSD"
New_YBIATAB0.BIATABTXT = "# S"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "Liste des ajustements de soldes CPT / GES"
Parametrage_New
'=====================================================================================================================

End Sub
Public Sub param_Init_BIA_GOS()
Dim X As String
'=====================================================================================================================
Call param_Init_Delete("BIA_GOS")
'=====================================================================================================================
New_YBIATAB0.BIATABK2 = "GOS_MAJ"
New_YBIATAB0.BIATABTXT = "Mise à jour des dossiers YGOSDOS0"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "S02"
Parametrage_New

New_YBIATAB0.BIATABK2 = "GOS_ADM"
New_YBIATAB0.BIATABTXT = "Administration des dossiers YGOSDOS0"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "S03"
Parametrage_New

New_YBIATAB0.BIATABK2 = "Paramétrage"
New_YBIATAB0.BIATABTXT = "Paramétrage"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "S18"
Parametrage_New

New_YBIATAB0.BIATABK2 = "GOS"
New_YBIATAB0.BIATABTXT = "Menu de gestion des opérations en suspens"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "S04"
Parametrage_New

New_YBIATAB0.BIATABK2 = "SWISAB_SQL"
New_YBIATAB0.BIATABTXT = "Menu de consultation de la base BIA-SWIFT)"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "*10"
Parametrage_New

New_YBIATAB0.BIATABK2 = "SWISAB_PDE"
New_YBIATAB0.BIATABTXT = "Menu de gestion PDE, msg orhelins, ..."
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "S11"
Parametrage_New

New_YBIATAB0.BIATABK2 = "SWISAB_MAJ"
New_YBIATAB0.BIATABTXT = "Mise à jour YSWISAB0"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "S12"
Parametrage_New
'=====================================================================================================================


X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" _
     & " where BIATABID = 'BIA_VB_HAB' and BIATABK1 = 'BIA_GOS' order by BIATABK2"
        
Set rsSabX = cnsab.Execute(X)
Do While Not rsSabX.EOF
    Call rsYBIATAB0_GetBuffer(rsSabX, Old_YBIATAB0)
    New_YBIATAB0 = Old_YBIATAB0
    Mid$(New_YBIATAB0.BIATABTXT, 2, 18) = Space$(18)
    Parametrage_Update

    rsSabX.MoveNext
Loop

'_______________________________________________________________________________________
New_YBIATAB0.BIATABID = "BIA_VB_MNU"
New_YBIATAB0.BIATABK1 = "BIA_GOS"

New_YBIATAB0.BIATABK2 = "JPL"
New_YBIATAB0.BIATABTXT = "#                 X"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "TEST JPL"
Parametrage_New

New_YBIATAB0.BIATABK2 = "1"
New_YBIATAB0.BIATABTXT = "#"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "SAA : consultation des messages (base SAA)"
Parametrage_New

New_YBIATAB0.BIATABK2 = "9"
New_YBIATAB0.BIATABTXT = "#"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "SAA : gestion des messages *99"
Parametrage_New

New_YBIATAB0.BIATABK2 = "1L"
New_YBIATAB0.BIATABTXT = "#        *"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "SAA : messages 'LIVE' (base SAA)"
Parametrage_New

New_YBIATAB0.BIATABK2 = "1b"
New_YBIATAB0.BIATABTXT = "#        *"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "SAB : consultation des messages (base SAB)"
Parametrage_New

New_YBIATAB0.BIATABK2 = "5"
New_YBIATAB0.BIATABTXT = "#         *"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "SAB : gestion des PDE et références en double"
Parametrage_New

New_YBIATAB0.BIATABK2 = "5h"
New_YBIATAB0.BIATABTXT = "#         *"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "SAB : historique des actions (PDE et références en double)"
Parametrage_New


New_YBIATAB0.BIATABK2 = "3"
New_YBIATAB0.BIATABTXT = "#  S"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "GOS : gestion des opérations en suspens"
Parametrage_New

New_YBIATAB0.BIATABK2 = "4"
New_YBIATAB0.BIATABTXT = "#  S"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "GOS : Echéancier / service"
Parametrage_New

New_YBIATAB0.BIATABK2 = "2"
New_YBIATAB0.BIATABTXT = "# S"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "SAA : création d'un dossier GOS"
Parametrage_New

New_YBIATAB0.BIATABK2 = "1?"
New_YBIATAB0.BIATABTXT = "#        *"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "SAB : messages orphelins(base SAB)"
Parametrage_New

New_YBIATAB0.BIATABK2 = "1?*"
New_YBIATAB0.BIATABTXT = "#        *"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "SAB : réaffectation des messages(base SAB)"
Parametrage_New

New_YBIATAB0.BIATABK2 = "9+"
New_YBIATAB0.BIATABTXT = "#         *"
Mid$(New_YBIATAB0.BIATABTXT, 20, 79) = "SAA : gestion des messages *99 (filtre)"
Parametrage_New
'=====================================================================================================================
' -
End Sub

Public Sub param_Init_Delete(lK1 As String)
Dim X As String

X = "delete from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_DROIT' and BIATABK1 = '" & lK1 & "'"
Call FEU_ROUGE
Call Parametrage_SQL(X)
X = "delete from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_VB_MNU' and BIATABK1 = '" & lK1 & "'"
Call Parametrage_SQL(X)
Call FEU_VERT
'=====================================================================================================================
New_YBIATAB0.BIATABID = "BIA_VB_DROIT"
New_YBIATAB0.BIATABK1 = lK1
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "accès à l'application " & lK1
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "#01"
Parametrage_New

New_YBIATAB0.BIATABK2 = "Admin"
New_YBIATAB0.BIATABTXT = "réservé administrateur"
Mid$(New_YBIATAB0.BIATABTXT, 100, 3) = "X19"
Parametrage_New


End Sub

Private Sub txtUsr_Code_GotFocus()
txtUsr_Code.BackColor = focusUsr.BackColor

End Sub


Private Sub txtUsr_Code_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtUsr_Code_LostFocus()
txtUsr_Code.BackColor = txtUsr.BackColor
End Sub



Public Sub Form_Init_cboSelect_Srv()
Dim X As String, xSQL As String

'load Service ___________________________________________________________________________
cboSelect_Srv.Clear
If blnAdmin Or arrHab(9) Then
    cboSelect_Srv.AddItem ""
    X = ""
Else
    X = " and SSIUSRUNIT = '" & currentSSIWINUNIT & "'"
End If
xSQL = "select *from " & paramIBM_Library_SABSPE & ".YSSIUSR0 where SSIUSRNAT= 'S'" & X & " order by SSIUSRUNIT"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    X = rsSab("SSIUSRUNIT") & " - " & rsSab("SSIUSRUIDX")
    cboSelect_Srv.AddItem X
    rsSab.MoveNext
Loop

End Sub

Public Sub Form_Init_cboSelect_Usr()
Dim X As String, xSQL As String

'load Usr ___________________________________________________________________________
cboSelect_Usr.Clear
If blnAdmin Or arrHab(9) Then
     X = " and SSIDOMUNIT <> ''"
    cboSelect_Usr.AddItem ""
Else
   cboSelect_Srv.ListIndex = 0
    X = " and SSIDOMUNIT = '" & currentSSIWINUNIT & "'"
End If

xSQL = "select *from " & paramIBM_Library_SABSPE & ".YSSIDOM0 where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN' " & X & " order by SSIDOMUIDX"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    X = Trim(rsSab("SSIDOMUIDX"))
    cboSelect_Usr.AddItem X
    rsSab.MoveNext
Loop


End Sub
