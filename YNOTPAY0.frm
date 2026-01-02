VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmYNOTPAY0 
   AutoRedraw      =   -1  'True
   Caption         =   "Surveillance Notation Pays"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13575
   Icon            =   "YNOTPAY0.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9315
   ScaleWidth      =   13575
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   0
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8688
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   15319
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "Notation pays"
      TabPicture(0)   =   "YNOTPAY0.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Paramétrage liste des pays"
      TabPicture(1)   =   "YNOTPAY0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgPays"
      Tab(1).Control(1)=   "fraPays"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "......."
      TabPicture(2)   =   "YNOTPAY0.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblSelect_NOTPAYXAMJ"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fgSAB_Client_Detail"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtSelect_NOTPAYXAMJ"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lstPays"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "fraDetail"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "fraSelect_Import"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "fraJRNENT0"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "fgSAB_Client"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cboNOTPAYLOGK"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "fraParam"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      Begin VB.Frame fraParam 
         BackColor       =   &H00E0FFFF&
         Caption         =   "Nature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3012
         Left            =   360
         TabIndex        =   60
         Top             =   3480
         Visible         =   0   'False
         Width           =   3852
         Begin VB.CommandButton cmdParam_New 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Ajouter"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1900
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   2400
            Width           =   900
         End
         Begin VB.CommandButton cmdParam_Delete 
            BackColor       =   &H000000FF&
            Caption         =   "Supprimer"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1000
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   2400
            Width           =   900
         End
         Begin VB.CommandButton cmdParam_Update 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Annuler / remplacer"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2800
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   2400
            Width           =   900
         End
         Begin VB.CommandButton cmdParam_Quit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Abandonner"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   100
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   2400
            Width           =   900
         End
         Begin VB.TextBox txtParam_Taux 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1680
            TabIndex        =   63
            Top             =   1560
            Width           =   852
         End
         Begin VB.TextBox txtParam_Code 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1680
            TabIndex        =   62
            Top             =   840
            Width           =   852
         End
         Begin VB.Label lblParam_Taux 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Valeur"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   480
            TabIndex        =   64
            Top             =   1560
            Width           =   972
         End
         Begin VB.Label lblParam_Code 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   480
            TabIndex        =   61
            Top             =   840
            Width           =   1212
         End
      End
      Begin VB.ComboBox cboNOTPAYLOGK 
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
         Left            =   720
         Sorted          =   -1  'True
         TabIndex        =   57
         Text            =   "cboNOTPAYLOGK"
         Top             =   3000
         Visible         =   0   'False
         Width           =   3612
      End
      Begin VB.Frame fraPays 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Paramétrage import PAYS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6012
         Left            =   -67080
         TabIndex        =   42
         Top             =   1560
         Visible         =   0   'False
         Width           =   4692
         Begin VB.CommandButton cmdPays_New 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Ajouter"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   4800
            Width           =   900
         End
         Begin VB.CommandButton cmdPays_Delete 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Supprimer"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   4800
            Width           =   900
         End
         Begin VB.CommandButton cmdPays_Update 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Annuler / remplacer"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   4800
            Width           =   900
         End
         Begin VB.CommandButton cmdPays_Quit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Abandonner"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   4800
            Width           =   900
         End
         Begin VB.TextBox txtPAYS_SP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1440
            MaxLength       =   32
            TabIndex        =   54
            Top             =   3840
            Width           =   3012
         End
         Begin VB.TextBox txtPAYS_OCDE_lib 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1440
            MaxLength       =   32
            TabIndex        =   53
            Top             =   3120
            Width           =   3012
         End
         Begin VB.TextBox txtPAYS_OCDE_Code 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   52
            Top             =   2520
            Width           =   612
         End
         Begin VB.TextBox txtPAYS_Coface 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1440
            TabIndex        =   51
            Top             =   1920
            Width           =   3012
         End
         Begin VB.TextBox txtPAYS_SAB 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1440
            MaxLength       =   32
            TabIndex        =   50
            Top             =   1320
            Width           =   3012
         End
         Begin VB.TextBox txtPAYS_ISO 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   49
            Top             =   720
            Width           =   372
         End
         Begin VB.Label lblPays_SP 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Libellé S P"
            Height          =   252
            Left            =   120
            TabIndex        =   48
            Top             =   3840
            Width           =   1212
         End
         Begin VB.Label lblPays_OCD_Lib 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Libellé OCDE"
            Height          =   252
            Left            =   120
            TabIndex        =   47
            Top             =   3120
            Width           =   1332
         End
         Begin VB.Label lblPays_OCDE_Code 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Code OCDE"
            Height          =   252
            Left            =   120
            TabIndex        =   46
            Top             =   2520
            Width           =   1332
         End
         Begin VB.Label lblPays_Coface 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Libellé Coface"
            Height          =   252
            Left            =   120
            TabIndex        =   45
            Top             =   1920
            Width           =   1332
         End
         Begin VB.Label lblPays_SAB 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Libellé SAB"
            Height          =   252
            Left            =   120
            TabIndex        =   44
            Top             =   1320
            Width           =   1332
         End
         Begin VB.Label lblPays_ISo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Code ISO"
            Height          =   252
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   1332
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSAB_Client 
         Height          =   612
         Left            =   720
         TabIndex        =   40
         Top             =   2280
         Visible         =   0   'False
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   1085
         _Version        =   393216
         Cols            =   6
         FixedCols       =   3
         RowHeightMin    =   300
         BackColorFixed  =   15794175
         FormatString    =   "Pays                                                |>note BIA |> Taux  |>Nb client SAB |<CATEG*    |"
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
      Begin VB.Frame fraJRNENT0 
         BackColor       =   &H00B0F0FF&
         Caption         =   "Journalisation"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   6900
         Left            =   11880
         TabIndex        =   35
         Top             =   3000
         Visible         =   0   'False
         Width           =   3012
         Begin MSFlexGridLib.MSFlexGrid fgJRNENT0 
            Height          =   5000
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   2650
            _ExtentX        =   4683
            _ExtentY        =   8811
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            RowHeightMin    =   285
            BackColor       =   16316664
            ForeColor       =   8388608
            BackColorFixed  =   10543359
            ForeColorFixed  =   0
            BackColorSel    =   12648384
            BackColorBkg    =   16777215
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            ScrollBars      =   2
            AllowUserResizing=   3
            FormatString    =   "< Champ          |< Valeur                    "
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
      Begin VB.Frame fraSelect_Import 
         BackColor       =   &H00E0FFE0&
         Height          =   1005
         Left            =   720
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   9192
         Begin VB.CommandButton cmdSelect_Internet_SP 
            BackColor       =   &H00A0E0FF&
            Caption         =   "Internet  S and P"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   240
            Width           =   800
         End
         Begin VB.CommandButton cmdSelect_Internet_OCDE 
            BackColor       =   &H00A0E0FF&
            Caption         =   "Internet OCDE"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   240
            Width           =   800
         End
         Begin VB.CommandButton cmdSelect_Internet_Coface 
            BackColor       =   &H00A0E0FF&
            Caption         =   "Internet Coface"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   240
            Width           =   800
         End
         Begin VB.CommandButton cmdSelect_Import_SP 
            BackColor       =   &H00C0FFFF&
            Caption         =   "  Importer   S and P"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   240
            Width           =   800
         End
         Begin VB.CommandButton cmdSelect_Import_OCDE 
            BackColor       =   &H00C0FFFF&
            Caption         =   "  Importer OCDE"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   240
            Width           =   800
         End
         Begin VB.CommandButton cmdSelect_Import_Coface 
            BackColor       =   &H00C0FFFF&
            Caption         =   " Importer Coface.txt"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   240
            Width           =   800
         End
         Begin VB.CommandButton cmdSelect_Import_Validation 
            BackColor       =   &H0080FF80&
            Caption         =   "Validation"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   240
            Width           =   800
         End
         Begin VB.CommandButton cmdSelect_Import_Delete 
            BackColor       =   &H000000FF&
            Caption         =   "Effacer"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   5760
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   240
            Width           =   800
         End
         Begin MSComCtl2.DTPicker txtSelect_Import_Amj 
            Height          =   300
            Left            =   7800
            TabIndex        =   28
            Top             =   480
            Width           =   1332
            _ExtentX        =   2355
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
            Format          =   96862211
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.Label lblSelect_Import_Amj 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0FFE0&
            Caption         =   "Date de la notation"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   7680
            TabIndex        =   29
            Top             =   240
            Width           =   1452
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraDetail 
         BackColor       =   &H00A0E0FF&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6900
         Left            =   8280
         TabIndex        =   18
         Top             =   3240
         Visible         =   0   'False
         Width           =   4740
         Begin VB.Frame fraDetail_Update 
            BackColor       =   &H00B0F0FF&
            Height          =   5292
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   4452
            Begin VB.TextBox txtDetail_NOTPAYTXT 
               Height          =   288
               Left            =   1560
               TabIndex        =   85
               Top             =   4800
               Width           =   2652
            End
            Begin VB.TextBox txtDetail_NOTPAYFISC 
               Height          =   240
               Left            =   3720
               MaxLength       =   2
               TabIndex        =   83
               Top             =   4320
               Width           =   492
            End
            Begin VB.CheckBox chkDetail_NOTPAYPROV 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00B0F0FF&
               Caption         =   "Provision"
               ForeColor       =   &H00FF0000&
               Height          =   252
               Left            =   1680
               TabIndex        =   82
               Top             =   4320
               Width           =   1044
            End
            Begin VB.TextBox txtDetail_NOTPAYCEG 
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
               Height          =   276
               Left            =   2040
               MaxLength       =   2
               TabIndex        =   81
               Top             =   3000
               Width           =   372
            End
            Begin VB.ComboBox cboDetail_NOTPAYBIAN 
               BackColor       =   &H00E0FFFF&
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   312
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   79
               Top             =   3480
               Width           =   852
            End
            Begin VB.ComboBox cboDetail_NOTPAYSP 
               BackColor       =   &H00E0FFFF&
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   312
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   77
               Top             =   2280
               Width           =   852
            End
            Begin VB.ComboBox cboDetail_NOTPAYOCDE 
               BackColor       =   &H00E0FFFF&
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
               Height          =   312
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   75
               Top             =   1440
               Width           =   852
            End
            Begin VB.ComboBox cboDetail_NOTPAYCOF2 
               BackColor       =   &H00E0FFFF&
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
               Height          =   312
               Left            =   3000
               Style           =   2  'Dropdown List
               TabIndex        =   72
               Top             =   960
               Width           =   852
            End
            Begin VB.ComboBox cboDetail_NOTPAYCOFA 
               BackColor       =   &H00E0FFFF&
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
               Height          =   312
               Left            =   1560
               Style           =   2  'Dropdown List
               TabIndex        =   71
               Top             =   360
               Width           =   852
            End
            Begin MSComCtl2.DTPicker txtDetail_NOTPAYCOFD 
               Height          =   300
               Left            =   3000
               TabIndex        =   70
               Top             =   360
               Width           =   1332
               _ExtentX        =   2355
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
               Format          =   96862211
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtDetail_NOTPAYOCDD 
               Height          =   300
               Left            =   3000
               TabIndex        =   76
               Top             =   1440
               Width           =   1332
               _ExtentX        =   2355
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
               Format          =   96862211
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtDetail_NOTPAYSPD 
               Height          =   300
               Left            =   3000
               TabIndex        =   78
               Top             =   2280
               Width           =   1332
               _ExtentX        =   2355
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
               Format          =   96862211
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtDetail_NOTPAYBIAD 
               Height          =   300
               Left            =   3000
               TabIndex        =   80
               Top             =   3480
               Width           =   1332
               _ExtentX        =   2355
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
               Format          =   96862211
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblDetail_NOTPAYCEG 
               BackColor       =   &H00B0F0FF&
               Caption         =   "critère événement grave"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   252
               Left            =   240
               TabIndex        =   95
               Top             =   3000
               Width           =   1812
            End
            Begin VB.Label libDetail_NOTPAYTAUX 
               Alignment       =   2  'Center
               BackColor       =   &H00F0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "000.00%"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   468
               Left            =   240
               TabIndex        =   94
               Top             =   4200
               Width           =   1332
            End
            Begin VB.Label libDetail_BIA 
               BackColor       =   &H00B0F0FF&
               Caption         =   "BIAK"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   204
               Left            =   3050
               TabIndex        =   93
               Top             =   3840
               Width           =   1300
            End
            Begin VB.Label libDetail_SP 
               BackColor       =   &H00B0F0FF&
               Caption         =   "SPK"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   204
               Left            =   3050
               TabIndex        =   92
               Top             =   2640
               Width           =   1300
            End
            Begin VB.Label libDetail_OCDE 
               BackColor       =   &H00B0F0FF&
               Caption         =   "OCDK"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   204
               Left            =   3050
               TabIndex        =   91
               Top             =   1800
               Width           =   1300
            End
            Begin VB.Label libDetail_Coface 
               BackColor       =   &H00B0F0FF&
               Caption         =   "COFK"
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
               Left            =   3050
               TabIndex        =   90
               Top             =   720
               Width           =   1300
            End
            Begin VB.Label lblDetail_BIA 
               BackColor       =   &H00B0F0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "BIA"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   312
               Left            =   240
               TabIndex        =   89
               Top             =   3480
               Width           =   1332
            End
            Begin VB.Label lblDetail_SP 
               BackColor       =   &H00B0F0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "S P"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   312
               Left            =   240
               TabIndex        =   88
               Top             =   2280
               Width           =   1332
            End
            Begin VB.Label lblDetail_OCDE 
               BackColor       =   &H00B0F0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "OCDE"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   312
               Left            =   240
               TabIndex        =   87
               Top             =   1440
               Width           =   1332
            End
            Begin VB.Label lblDetail_NOTPAYTXT 
               BackColor       =   &H00B0F0FF&
               Caption         =   "Commentaire"
               Height          =   252
               Left            =   480
               TabIndex        =   86
               Top             =   4920
               Width           =   972
            End
            Begin VB.Label lblDetail_NOTPAYFISC 
               BackColor       =   &H00B0F0FF&
               Caption         =   "Fisc %"
               Height          =   252
               Left            =   3000
               TabIndex        =   84
               Top             =   4320
               Width           =   612
            End
            Begin VB.Label lblDetail_Coface 
               BackColor       =   &H00B0F0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Coface"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   312
               Left            =   240
               TabIndex        =   74
               Top             =   360
               Width           =   1332
            End
            Begin VB.Label lblDetail_NOTPAYCOF2 
               BackColor       =   &H00B0F0FF&
               Caption         =   "affaires"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   312
               Left            =   1680
               TabIndex        =   73
               Top             =   960
               Width           =   1332
            End
         End
         Begin VB.CommandButton cmdDetail_New 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Archive / Remplace"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   6000
            Width           =   900
         End
         Begin VB.CommandButton cmdDetail_Delete 
            BackColor       =   &H000000FF&
            Caption         =   "Supprimer"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   6000
            Width           =   900
         End
         Begin VB.CommandButton cmdDetail_Copy 
            BackColor       =   &H00FF80FF&
            Caption         =   "Copier"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   6000
            Width           =   800
         End
         Begin VB.CommandButton cmdDetail_Quit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Abandonner"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   6000
            Width           =   900
         End
         Begin VB.CommandButton cmdDetail_Update 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Annule / remplace"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   6000
            Width           =   900
         End
      End
      Begin VB.ListBox lstPays 
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
         Left            =   6360
         TabIndex        =   17
         Top             =   3240
         Visible         =   0   'False
         Width           =   7000
      End
      Begin VB.Frame fraTab0 
         Height          =   8232
         Left            =   -74880
         TabIndex        =   3
         Top             =   600
         Width           =   13296
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
            Left            =   9240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   240
            Width           =   3732
         End
         Begin VB.TextBox txtDetail 
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   11640
            TabIndex        =   10
            Top             =   2040
            Visible         =   0   'False
            Width           =   1452
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   555
            Left            =   10560
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   720
            Width           =   1335
         End
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00F0FFFF&
            Height          =   1005
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   8832
            Begin VB.CheckBox chkSelect_NOTPAYSEQ 
               BackColor       =   &H0080C0FF&
               Caption         =   "afficher l'historique des notations"
               Height          =   312
               Left            =   1320
               TabIndex        =   16
               Top             =   240
               Visible         =   0   'False
               Width           =   2652
            End
            Begin VB.TextBox txtSelect_Pays 
               Height          =   288
               Left            =   4200
               TabIndex        =   12
               Top             =   600
               Width           =   2172
            End
            Begin VB.ComboBox txtSelect_NOTPAYISO 
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
               Left            =   360
               Sorted          =   -1  'True
               TabIndex        =   8
               Text            =   "txtSelect_NOTPAYISO"
               Top             =   550
               Width           =   3612
            End
            Begin MSComCtl2.DTPicker txtSelect_AMJMIN 
               Height          =   300
               Left            =   6840
               TabIndex        =   14
               Top             =   600
               Width           =   1332
               _ExtentX        =   2355
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
               Format          =   96862211
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_AMJMIN 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00F0FFFF&
               Caption         =   "affichage des dates màj >="
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Left            =   6360
               TabIndex        =   15
               Top             =   240
               Width           =   1992
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblSelect_Pays 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Pays (alphabétique)"
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
               Left            =   4440
               TabIndex        =   13
               Top             =   240
               Width           =   1572
            End
            Begin VB.Label lblSelect_NOTPAYISO 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Pays (ISO)"
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
               TabIndex        =   9
               Top             =   240
               Width           =   972
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6825
            Left            =   240
            TabIndex        =   5
            Top             =   1320
            Width           =   12915
            _ExtentX        =   22781
            _ExtentY        =   12039
            _Version        =   393216
            Rows            =   1
            Cols            =   22
            RowHeightMin    =   350
            BackColor       =   -2147483633
            ForeColor       =   12582912
            BackColorFixed  =   15794175
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   -2147483633
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"YNOTPAY0.frx":035E
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
      Begin MSComCtl2.DTPicker txtSelect_NOTPAYXAMJ 
         Height          =   300
         Left            =   1800
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   1332
         _ExtentX        =   2355
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
         Format          =   96862211
         CurrentDate     =   38699.44875
         MaxDate         =   401768
         MinDate         =   36526.4425347222
      End
      Begin MSFlexGridLib.MSFlexGrid fgPays 
         Height          =   6828
         Left            =   -74880
         TabIndex        =   26
         Top             =   1080
         Width           =   12912
         _ExtentX        =   22781
         _ExtentY        =   12039
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   -2147483633
         ForeColor       =   12582912
         BackColorFixed  =   15794175
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   3
         FormatString    =   $"YNOTPAY0.frx":0503
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
      Begin MSFlexGridLib.MSFlexGrid fgSAB_Client_Detail 
         Height          =   612
         Left            =   720
         TabIndex        =   41
         Top             =   1680
         Visible         =   0   'False
         Width           =   5592
         _ExtentX        =   9869
         _ExtentY        =   1085
         _Version        =   393216
         Cols            =   3
         FixedCols       =   2
         RowHeightMin    =   300
         BackColorFixed  =   15794175
         BackColorBkg    =   15794175
         FormatString    =   "Pays               |Racine       |Intitulé                                                                      "
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
      Begin VB.Label lblSelect_NOTPAYXAMJ 
         BackColor       =   &H00F0FFFF&
         Caption         =   "en date du"
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
         Left            =   840
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   852
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
      Height          =   500
      Left            =   13080
      Picture         =   "YNOTPAY0.frx":05F3
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
   Begin VB.Menu mnufgSelect 
      Caption         =   "mnufgSelect"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmYNOTPAY0"
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
Dim YNOTPAY0_Aut As typeAuthorization
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

Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long
Dim xYNOTPAY0 As typeYNOTPAY0, newYNOTPAY0 As typeYNOTPAY0, oldYNOTPAY0 As typeYNOTPAY0
Dim arrYNOTPAY0() As typeYNOTPAY0, arrYNOTPAY0_Nb As Long, arrYNOTPAY0_Max As Long, arrYNOTPAY0_Index As Long
Dim lastYNOTPAY0 As typeYNOTPAY0
Dim selYNOTPAY0() As typeYNOTPAY0, selYNOTPAY0_Nb As Integer, selYNOTPAY0_Max As Integer

Dim txtDetail_Type As String, txtDetail_Update As String
Dim txtDetail_Field As String
Dim txtDetail_blnUpdate As Boolean, txtDetail_ColorUpdate As Long
Dim txtDetail_Row As Integer

Dim arrPAYS_ISO() As String, arrPays_NB As Integer
Dim arrPays_Lib() As String
Dim arrCoface() As typeYNOTPAY0, arrCoface_Nb As Integer
Dim arrOCDE() As typeYNOTPAY0, arrOCDE_Nb As Integer
Dim arrSP() As typeYNOTPAY0, arrSP_Nb As Integer
Dim arrBIAN() As typeYNOTPAY0, arrBIAN_Nb As Integer

Dim blnNOTPAYBIAN_à_Calculer As Boolean, blnNOTPAYBIAN_Manuel As Boolean

Dim blnDetail_Copy As Boolean

'____________________________________________________________________________________
Dim arrPays_Import() As typeYBIATAB0, arrPays_Import_Nb As Integer, arrPays_Import_Max As Integer
Dim xPays_Import As typeYBIATAB0

Dim fgPays_FormatString As String, fgPays_K As Integer
Dim fgPays_RowDisplay As Integer, fgPays_RowClick As Integer, fgPays_ColClick As Integer
Dim fgPays_ColorClick As Long, fgPays_ColorDisplay As Long
Dim fgPays_Sort1 As Integer, fgPays_Sort2 As Integer
Dim fgPays_SortAD As Integer, fgPays_Sort1_Old As Integer
Dim fgPays_arrIndex As Integer
Dim blnfgPays_DisplayLine As Boolean

Dim xJRNENT0 As typeJRNENT0

Dim Coface_Internet As typeYBIATAB0, OCDE_Internet As typeYBIATAB0, SP_Internet As typeYBIATAB0
Dim Old_YBIATAB0 As typeYBIATAB0, New_YBIATAB0 As typeYBIATAB0

Dim Coface_Notepad As typeYBIATAB0, OCDE_Notepad As typeYBIATAB0, SP_Notepad As typeYBIATAB0
Dim Export_Lien As typeYBIATAB0
Dim Compta_Lien As typeYBIATAB0

Dim Coface_DateNotation As typeYBIATAB0, OCDE_DateNotation As typeYBIATAB0, SP_DateNotation As typeYBIATAB0
Dim BIA_DateNotation As typeYBIATAB0
Dim Coface_DateNotation_Info As String, OCDE_DateNotation_Info As String, SP_DateNotation_Info As String
Dim BIA_DateNotation_Info As String

Dim newYNOTPAYLOG As typeYNOTPAYLOG, xYNOTPAYLOG As typeYNOTPAYLOG
Dim arrYNOTPAYLOG() As typeYNOTPAYLOG, arrYNOTPAYLOG_Nb As Long, arrYNOTPAYLOG_Max As Long, arrYNOTPAYLOG_Index As Long

Dim blnDetail_NOTPAYBIAK As Boolean


Dim arrSAB_CATEG_ISO() As String, arrSAB_CATEG_Code() As String, arrSAB_CATEG_Nb As Integer

Private Sub fraDetail_Display()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = "fgDetail_Display"
SSTab1.Tab = 0
fraDetail.Visible = False
'_____________________________________________________________
If cmdSelect_SQL_K = "Pn" Then
    fraParam_Display
    Exit Sub
End If

'_____________________________________________________________

fraDetail.Caption = oldYNOTPAY0.NOTPAYLIB
cmdDetail_Delete.Visible = False
cmdDetail_Copy.Visible = False
fraDetail_Update.Enabled = False


lblDetail_Coface = "Coface : " & oldYNOTPAY0.NOTPAYCOFA
libDetail_Coface = dateImp10(oldYNOTPAY0.NOTPAYCOFD) & "  " & oldYNOTPAY0.NOTPAYCOFK
libDetail_Coface.ForeColor = IIf(oldYNOTPAY0.NOTPAYCOFK = "M", vbMagenta, vbBlue)

cboDetail_NOTPAYCOFA = oldYNOTPAY0.NOTPAYCOFA
cboDetail_NOTPAYCOFA.BackColor = RGB(255, 255, 230)
cboDetail_NOTPAYCOFA.ForeColor = vbBlue

lblDetail_NOTPAYCOF2 = "affaires : " & oldYNOTPAY0.NOTPAYCOF2
cboDetail_NOTPAYCOF2 = oldYNOTPAY0.NOTPAYCOF2
cboDetail_NOTPAYCOF2.BackColor = RGB(255, 255, 230)
cboDetail_NOTPAYCOF2.ForeColor = vbBlue
If oldYNOTPAY0.NOTPAYCOFD = 0 Then
    X = DSys
Else
    X = oldYNOTPAY0.NOTPAYCOFD
End If
Call DTPicker_Set(txtDetail_NOTPAYCOFD, X)

lblDetail_OCDE = "OCDE : " & oldYNOTPAY0.NOTPAYOCDE
libDetail_OCDE = dateImp10(oldYNOTPAY0.NOTPAYOCDD) & "  " & oldYNOTPAY0.NOTPAYOCDK
libDetail_OCDE.ForeColor = IIf(oldYNOTPAY0.NOTPAYOCDK = "M", vbMagenta, vbBlue)
cboDetail_NOTPAYOCDE = oldYNOTPAY0.NOTPAYOCDE
cboDetail_NOTPAYOCDE.BackColor = RGB(255, 255, 230)
cboDetail_NOTPAYOCDE.ForeColor = vbBlue
If oldYNOTPAY0.NOTPAYOCDD = 0 Then
    X = DSys
Else
    X = oldYNOTPAY0.NOTPAYOCDD
End If
Call DTPicker_Set(txtDetail_NOTPAYOCDD, X)

lblDetail_SP = "S/P  : " & oldYNOTPAY0.NOTPAYSP
libDetail_SP = dateImp10(oldYNOTPAY0.NOTPAYSPD) & "  " & oldYNOTPAY0.NOTPAYSPK
libDetail_SP.ForeColor = IIf(oldYNOTPAY0.NOTPAYSPK = "M", vbMagenta, vbBlue)
cboDetail_NOTPAYSP = oldYNOTPAY0.NOTPAYSP
cboDetail_NOTPAYSP.BackColor = RGB(255, 255, 230)
cboDetail_NOTPAYSP.ForeColor = vbBlue
If oldYNOTPAY0.NOTPAYSPD = 0 Then
    X = DSys
Else
    X = oldYNOTPAY0.NOTPAYSPD
End If
Call DTPicker_Set(txtDetail_NOTPAYSPD, X)

lblDetail_BIA = "BIA  : " & oldYNOTPAY0.NOTPAYBIAN
libDetail_BIA = dateImp10(oldYNOTPAY0.NOTPAYBIAD) & "  " & oldYNOTPAY0.NOTPAYBIAK
libDetail_BIA.ForeColor = IIf(oldYNOTPAY0.NOTPAYBIAK = "M", vbMagenta, vbBlue)
cboDetail_NOTPAYBIAN = oldYNOTPAY0.NOTPAYBIAN
cboDetail_NOTPAYBIAN.BackColor = RGB(255, 255, 230)
cboDetail_NOTPAYBIAN.ForeColor = vbBlue
txtDetail_NOTPAYCEG = Trim(oldYNOTPAY0.NOTPAYCEG)
txtDetail_NOTPAYCEG.BackColor = RGB(255, 255, 230)
If oldYNOTPAY0.NOTPAYBIAD = 0 Then
    X = DSys
Else
    X = oldYNOTPAY0.NOTPAYBIAD
End If
Call DTPicker_Set(txtDetail_NOTPAYBIAD, X)

If oldYNOTPAY0.NOTPAYPROV = "P" Then
    chkDetail_NOTPAYPROV = "1"
Else
    chkDetail_NOTPAYPROV = "0"
End If


txtDetail_NOTPAYFISC = Trim(oldYNOTPAY0.NOTPAYFISC)
txtDetail_NOTPAYFISC.BackColor = RGB(255, 255, 230)
txtDetail_NOTPAYTXT = Trim(oldYNOTPAY0.NOTPAYTXT)
txtDetail_NOTPAYTXT.BackColor = RGB(255, 255, 230)
libDetail_NOTPAYTAUX = Format$(oldYNOTPAY0.NOTPAYTAUX, "##0.00") & "%"
libDetail_NOTPAYTAUX.BackColor = RGB(240, 240, 240)
libDetail_NOTPAYTAUX.ForeColor = RGB(0, 128, 0)

fraDetail.Visible = True
blnNOTPAYBIAN_à_Calculer = False
blnNOTPAYBIAN_Manuel = False
If cmdSelect_SQL_K = "1" And oldYNOTPAY0.NOTPAYSEQ = 0 Then
    cmdDetail_Delete.Visible = YNOTPAY0_Aut.Saisir
    cmdDetail_Copy.Visible = YNOTPAY0_Aut.Saisir
    fraDetail_Update.Enabled = YNOTPAY0_Aut.Saisir
Else
    If cmdSelect_SQL_K = "I" And oldYNOTPAY0.NOTPAYSEQ = -1 Then
        fraDetail_Update.Enabled = YNOTPAY0_Aut.Saisir
        cmdDetail_Delete.Visible = YNOTPAY0_Aut.Saisir
    End If
End If

cmdDetail_Update.Visible = False: cmdDetail_New.Visible = False
cmdDetail_Quit.Enabled = True
blnDetail_NOTPAYBIAK = False


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Public Sub cmdSelect_Reset()
If blnControl Then
    If Me.Enabled Then cmdContext.SetFocus
    lstErr.Clear
    lstPays.Visible = False
    fgSelect.Visible = False
    fgSAB_Client.Visible = False
    fgSAB_Client_Detail.Visible = False
    fraDetail.Visible = False
    fraJRNENT0.Visible = False
    txtDetail.Visible = False
    fraSelect_Options.Visible = False
    fraSelect_Import.Visible = False
    cmdSelect_Ok.Visible = True
    cboNOTPAYLOGK.Visible = False
    fraParam.Visible = False
    If Trim(Mid$(txtSelect_NOTPAYISO, 1, 2)) = "" Then chkSelect_NOTPAYSEQ = "0"
    cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, 2))
    
    Select Case cmdSelect_SQL_K
        Case "1", "1h", "J"
            txtSelect_NOTPAYISO.Visible = True
            txtSelect_Pays.Visible = True
            lblSelect_NOTPAYISO.Visible = True
            lblSelect_Pays.Visible = True
        Case Else
            txtSelect_NOTPAYISO.Visible = False
            txtSelect_Pays.Visible = False
            lblSelect_NOTPAYISO.Visible = False
            lblSelect_Pays.Visible = False
            chkSelect_NOTPAYSEQ.Visible = False
    End Select
    Select Case cmdSelect_SQL_K
        Case "1", "Em", "Ex", "Pn", "I", "Sc":
                  lblSelect_AMJMIN = "affichage des dates màj =>"
                  cmdSelect_Ok_Click
         Case "1h":
                  lblSelect_AMJMIN = "date d'arrêté ="
                  fraSelect_Options.Visible = True
         Case "L":
                  lblSelect_AMJMIN = "date de départ =>"
                  fraSelect_Options.Visible = True
                  cboNOTPAYLOGK.Visible = True
       Case "J": chkSelect_NOTPAYSEQ.Visible = False
                  lblSelect_AMJMIN = "à partir du"
                  fraSelect_Options.Visible = True
        Case Else: chkSelect_NOTPAYSEQ.Visible = False
                   txtSelect_NOTPAYISO.ListIndex = 0
                   txtSelect_Pays = ""
                   chkSelect_NOTPAYSEQ = "0"
    End Select

End If

End Sub


Private Sub cmdSelect_SQL_1()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdYNOTPAY0_SQL"
blnOk = False
   
Call DTPicker_Control(txtSelect_NOTPAYXAMJ, WAMJMax)
Call DTPicker_Control(txtSelect_AMJMIN, wAMJMin)

If chkSelect_NOTPAYSEQ = "1" Then
    xWhere = " where NOTPAYSEQ >=  0"
Else
    xWhere = " where NOTPAYSEQ =  0"
End If

X = Trim(Mid$(txtSelect_NOTPAYISO, 1, 2))
If X <> "" Then
    blnOk = True: xWhere = xWhere & " and NOTPAYISO = '" & X & "'"
    chkSelect_NOTPAYSEQ.Visible = True
Else
    chkSelect_NOTPAYSEQ.Visible = False
End If


arrYNOTPAY0_SQL xWhere & " order by NOTPAYISO , NOTPAYSEQ asc"

fgSelect_Display


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub cmdSelect_SQL_1h()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1h"
blnOk = False
   
Call DTPicker_Control(txtSelect_AMJMIN, wAMJMin)

xWhere = " where NOTPAYHAMJ = " & wAMJMin

X = Trim(Mid$(txtSelect_NOTPAYISO, 1, 2))
If X <> "" Then
    blnOk = True: xWhere = xWhere & " and NOTPAYISO = '" & X & "'"
    chkSelect_NOTPAYSEQ.Visible = True
Else
    chkSelect_NOTPAYSEQ.Visible = False
End If


arrYNOTPAY0_SQL xWhere & " order by NOTPAYISO , NOTPAYSEQ asc"

fgSelect_Display


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_Journalisation()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_Journalisation"
blnOk = False
   
'Call DTPicker_Control(txtSelect_NOTPAYXAMJ, wAmjMax)
Call DTPicker_Control(txtSelect_AMJMIN, wAMJMin)

X = Trim(Mid$(txtSelect_NOTPAYISO, 1, 2))
If X <> "" Then
    xWhere = "where NOTPAYISO = '" & X & "' and NOTPAYXAMJ >= " & wAMJMin
Else
    xWhere = "where NOTPAYXAMJ >= " & wAMJMin
End If


arrJNOTPAY0_SQL xWhere & " order by JORCV , JOSEQN"

fgSelect_Display


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_YNOTPAYLOG()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_YNOTPAYLOG"
blnOk = False
   
Call DTPicker_Control(txtSelect_AMJMIN, wAMJMin)

xWhere = "where  NOTPAYLOGD >= " & wAMJMin

X = Trim(cboNOTPAYLOGK)
If X <> "" Then xWhere = xWhere & " and NOTPAYLOGK like '%" & X & "%'"
arrYNOTPAYLOG_SQL xWhere & " order by NOTPAYLOGD,NOTPAYLOGH,NOTPAYLOGU,NOTPAYLOGS"

fgSelect_Display_YNOTPAYLOG


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub cmdSelect_SQL_P()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim wCli As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_P"
blnOk = False
   
'Call DTPicker_Control(txtSelect_NOTPAYXAMJ, wAmjMax)

xWhere = " where NOTPAYISO =  '$$'"


arrYNOTPAY0_SQL xWhere & " order by NOTPAYTXT"
fgSelect_Display


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrYNOTPAY0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYNOTPAY0(101)
arrYNOTPAY0_Max = 100: arrYNOTPAY0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YNOTPAY0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYNOTPAY0_GetBuffer(rsSab, xYNOTPAY0)
    If xYNOTPAY0.NOTPAYISO = "$$" Then
        xYNOTPAY0.NOTPAYLIB = xYNOTPAY0.NOTPAYTXT  'xYNOTPAY0.NOTPAYCOFA & " " & xYNOTPAY0.NOTPAYOCDE & " " & xYNOTPAY0.NOTPAYSP & " " & xYNOTPAY0.NOTPAYBIAN
    Else
        For K = 1 To arrPays_NB
            If xYNOTPAY0.NOTPAYISO = arrPAYS_ISO(K) Then xYNOTPAY0.NOTPAYLIB = arrPays_Lib(K)
        Next K
    End If
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYNOTPAY0.fgselect_Display"
        '' Exit Sub
     Else
         arrYNOTPAY0_Nb = arrYNOTPAY0_Nb + 1
         If arrYNOTPAY0_Nb > arrYNOTPAY0_Max Then
             arrYNOTPAY0_Max = arrYNOTPAY0_Max + 100
             ReDim Preserve arrYNOTPAY0(arrYNOTPAY0_Max)
         End If
         
         arrYNOTPAY0(arrYNOTPAY0_Nb) = xYNOTPAY0
    End If
    rsSab.MoveNext
Loop

If chkSelect_NOTPAYSEQ = "1" Then
    arrYNOTPAY0(0) = arrYNOTPAY0(1)
    For K = 1 To arrYNOTPAY0_Nb - 1        'To 1 Step -1
        arrYNOTPAY0(K) = arrYNOTPAY0(K + 1)
    Next K
    arrYNOTPAY0(arrYNOTPAY0_Nb) = arrYNOTPAY0(0)
End If


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrJNOTPAY0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYNOTPAY0(101)
arrYNOTPAY0_Max = 100: arrYNOTPAY0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABJRN & ".JNOTPAY0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsJNOTPAY0_GetBuffer(rsSab, xYNOTPAY0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYNOTPAY0.fgselect_Display"
        '' Exit Sub
     Else
         arrYNOTPAY0_Nb = arrYNOTPAY0_Nb + 1
         If arrYNOTPAY0_Nb > arrYNOTPAY0_Max Then
             arrYNOTPAY0_Max = arrYNOTPAY0_Max + 100
             ReDim Preserve arrYNOTPAY0(arrYNOTPAY0_Max)
         End If
         
         arrYNOTPAY0(arrYNOTPAY0_Nb) = xYNOTPAY0
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


Private Sub arrYNOTPAYLOG_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYNOTPAYLOG(101)
arrYNOTPAYLOG_Max = 100: arrYNOTPAYLOG_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YNOTPAYLOG " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYNOTPAYLOG_GetBuffer(rsSab, xYNOTPAYLOG)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYNOTPAYLOG.fgselect_Display"
        '' Exit Sub
     Else
         arrYNOTPAYLOG_Nb = arrYNOTPAYLOG_Nb + 1
         If arrYNOTPAYLOG_Nb > arrYNOTPAYLOG_Max Then
             arrYNOTPAYLOG_Max = arrYNOTPAYLOG_Max + 100
             ReDim Preserve arrYNOTPAYLOG(arrYNOTPAYLOG_Max)
         End If
         
         arrYNOTPAYLOG(arrYNOTPAYLOG_Nb) = xYNOTPAYLOG
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



Private Sub arrPays_Import_SQL()
Dim V
Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrPays_Import(221)
arrPays_Import_Max = 220: arrPays_Import_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YNOTPAY0' and BIATABK1 = 'Pays' order by BIATABK2"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYBIATAB0_GetBuffer(rsSab, xPays_Import)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYBIATAB0.fgselect_Display"
        '' Exit Sub
     Else
         arrPays_Import_Nb = arrPays_Import_Nb + 1
         If arrPays_Import_Nb > arrPays_Import_Max Then
             arrPays_Import_Max = arrPays_Import_Max + 100
             ReDim Preserve arrPays_Import(arrPays_Import_Max)
         End If
         
         arrPays_Import(arrPays_Import_Nb) = xPays_Import
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


'______________________________________________________________________

Private Sub fgSelect_Display()
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
cmdSelect_Ok.Visible = False
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Row = 0

wColor = &HC0FFFF
If cmdSelect_SQL_K = "I" Then wColor = RGB(180, 255, 200)


fgSelect.Col = 0: fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor
fgSelect.Col = 3: fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor: fgSelect.CellAlignment = 4
fgSelect.Col = 4: fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor: fgSelect.CellAlignment = 4
fgSelect.Col = 6: fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor: fgSelect.CellAlignment = 4
fgSelect.Col = 8: fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor: fgSelect.CellAlignment = 4
fgSelect.Col = 11: fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor: fgSelect.CellAlignment = 4
fgSelect.Col = 13: fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor
currentAction = "fgSelect_Display"
    
For I = 1 To arrYNOTPAY0_Nb
         
    xYNOTPAY0 = arrYNOTPAY0(I)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    If xYNOTPAY0.NOTPAYSEQ < 0 Then
        Call fgSelect_DisplayLine_Import(I)
    Else
        fgSelect_DisplayLine I
    End If
    
Next I

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYNOTPAY0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

'______________________________________________________________________

Private Sub fgSelect_Display_YNOTPAYLOG()
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "Date               |<Heure             |<Utilisateur          |>Seq     |" _
                      & "<Action             |<Libellé                                                                                                                                  |"
fgSelect.Row = 0

currentAction = "fgSelect_Display_YNOTPAYLOG"
    
For I = 1 To arrYNOTPAYLOG_Nb
         
    xYNOTPAYLOG = arrYNOTPAYLOG(I)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    
    fgSelect.Col = 0: fgSelect.Text = dateImp10(xYNOTPAYLOG.NOTPAYLOGD)
    fgSelect.Col = 1: fgSelect.Text = timeImp8(xYNOTPAYLOG.NOTPAYLOGH)
    fgSelect.Col = 2: fgSelect.Text = xYNOTPAYLOG.NOTPAYLOGU
    fgSelect.Col = 3: fgSelect.Text = xYNOTPAYLOG.NOTPAYLOGS
    fgSelect.Col = 4: fgSelect.Text = xYNOTPAYLOG.NOTPAYLOGK
    fgSelect.Col = 5: fgSelect.Text = xYNOTPAYLOG.NOTPAYLOGX

Next I

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYNOTPAYLOG_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


'______________________________________________________________________

Private Sub fgPays_Display()

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
cmdSelect_Ok.Visible = False
fgPays.Visible = False
fgPays_Reset

fgPays.Rows = 1
fgPays.FormatString = fgPays_FormatString
fgPays.Row = 0
currentAction = "fgPays_Display"
    
For I = 1 To arrPays_Import_Nb
         
    xPays_Import = arrPays_Import(I)
    fgPays.Rows = fgPays.Rows + 1
    fgPays.Row = fgPays.Rows - 1
    fgPays_DisplayLine I
Next I

fgPays.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrPays_Import_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub fgPays_DisplayLine(lIndex As Long)
Dim K As Integer, xPays As String
Dim wColor As Long
On Error Resume Next

If fgPays.Row Mod 2 = 0 Then
    wColor = &HFFFFFF
Else
    wColor = &HF0F0F0
End If

xPays = Trim(xPays_Import.BIATABK2)
For K = 1 To arrPays_NB
    If xPays = arrPAYS_ISO(K) Then Exit For
Next K
fgPays.Col = 0: fgPays.Text = xPays
fgPays.CellBackColor = wColor
fgPays.Col = 1: fgPays.Text = arrPays_Lib(K)
fgPays.CellBackColor = wColor
fgPays.Col = 2: fgPays.Text = Mid$(xPays_Import.BIATABTXT, 36, 32)
fgPays.CellBackColor = wColor
fgPays.Col = 3: fgPays.Text = Mid$(xPays_Import.BIATABTXT, 1, 3)
fgPays.CellBackColor = wColor
fgPays.Col = 4: fgPays.Text = Mid$(xPays_Import.BIATABTXT, 4, 32)
fgPays.CellBackColor = wColor
fgPays.Col = 5: fgPays.Text = Mid$(xPays_Import.BIATABTXT, 68, 32)
fgPays.CellBackColor = wColor

fgPays.Col = fgPays_arrIndex: fgPays.Text = lIndex
End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_NOTPAYTAUX As Long
On Error Resume Next

If fgSelect.Row Mod 2 = 0 Then
    wColor = RGB(255, 255, 200): wColor_NOTPAYTAUX = RGB(255, 255, 128)
Else
    wColor = RGB(255, 255, 230): wColor_NOTPAYTAUX = RGB(255, 255, 200)
End If
If chkSelect_NOTPAYSEQ = "1" And xYNOTPAY0.NOTPAYSEQ = 0 Then wColor = RGB(127, 255, 127)
fgSelect.Col = 0: fgSelect.Text = xYNOTPAY0.NOTPAYISO & " - " & xYNOTPAY0.NOTPAYLIB ' arrPays_Lib(K)
fgSelect.CellBackColor = wColor_NOTPAYTAUX ' wColor
If xYNOTPAY0.NOTPAYHAMJ <> 0 Then
    fgSelect.Col = 1: fgSelect.Text = dateImpS(xYNOTPAY0.NOTPAYHAMJ)
    fgSelect.CellBackColor = RGB(180, 255, 200)
End If
fgSelect.Col = 2: fgSelect.Text = xYNOTPAY0.NOTPAYPROV
fgSelect.CellAlignment = 4
fgSelect.Col = 3: fgSelect.Text = xYNOTPAY0.NOTPAYCOFA
fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor
fgSelect.CellAlignment = 4
fgSelect.Col = 4: fgSelect.Text = xYNOTPAY0.NOTPAYCOF2
fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor
fgSelect.CellAlignment = 4
fgSelect.Col = 5: fgSelect.Text = xYNOTPAY0.NOTPAYCOFK & " " & dateImpS(xYNOTPAY0.NOTPAYCOFD)
fgSelect.CellFontSize = 6
If xYNOTPAY0.NOTPAYCOFD >= wAMJMin Then fgSelect.CellBackColor = &HA0E0FF

fgSelect.Col = 6: fgSelect.Text = xYNOTPAY0.NOTPAYOCDE
fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor
fgSelect.CellAlignment = 4
fgSelect.Col = 7: fgSelect.Text = xYNOTPAY0.NOTPAYOCDK & " " & dateImpS(xYNOTPAY0.NOTPAYOCDD)
If xYNOTPAY0.NOTPAYOCDD >= wAMJMin Then fgSelect.CellBackColor = &HA0E0FF
fgSelect.CellFontSize = 6

fgSelect.Col = 8: fgSelect.Text = xYNOTPAY0.NOTPAYSP
fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor
fgSelect.CellAlignment = 4
fgSelect.Col = 9: fgSelect.Text = xYNOTPAY0.NOTPAYSPK & " " & dateImpS(xYNOTPAY0.NOTPAYSPD)
If xYNOTPAY0.NOTPAYSPD >= wAMJMin Then fgSelect.CellBackColor = &HA0E0FF
fgSelect.CellFontSize = 6

If xYNOTPAY0.NOTPAYCEG <> 0 Then
    fgSelect.Col = 10: fgSelect.Text = xYNOTPAY0.NOTPAYCEG
    fgSelect.CellFontBold = True: fgSelect.CellBackColor = &HA0E0FF
    fgSelect.CellAlignment = 4
End If

fgSelect.Col = 11: fgSelect.Text = xYNOTPAY0.NOTPAYBIAN
fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor
fgSelect.CellAlignment = 4
fgSelect.Col = 12: fgSelect.Text = xYNOTPAY0.NOTPAYBIAK & " " & dateImpS(xYNOTPAY0.NOTPAYBIAD)
fgSelect.CellFontSize = 6
If xYNOTPAY0.NOTPAYBIAD >= wAMJMin Then fgSelect.CellBackColor = &HA0E0FF

fgSelect.Col = 13: fgSelect.Text = Format$(xYNOTPAY0.NOTPAYTAUX, " ##0.00")
fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor_NOTPAYTAUX ' wColor

fgSelect.Col = 14: fgSelect.Text = xYNOTPAY0.NOTPAYFISC
fgSelect.Col = 15: fgSelect.Text = Format$(xYNOTPAY0.NOTPAYSEQ, " ##0") & " - " & xYNOTPAY0.NOTPAYXUSR & " " & dateImpS(xYNOTPAY0.NOTPAYXAMJ) & " " & timeImp8(xYNOTPAY0.NOTPAYXHMS)
fgSelect.Col = 16: fgSelect.Text = Trim(xYNOTPAY0.NOTPAYTXT)
fgSelect.Col = 17: fgSelect.Text = Trim(xYNOTPAY0.JORCV)
fgSelect.Col = 18: fgSelect.Text = Trim(xYNOTPAY0.JOSEQN)


fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
End Sub
Public Sub fgSelect_DisplayLine_Import(lIndex As Long)
Dim K As Integer, K2 As Integer
Dim wColor As Long
On Error Resume Next
If fgSelect.Row Mod 2 = 0 Then
    wColor = RGB(220, 255, 240)
Else
    wColor = RGB(180, 255, 200)
End If

For K = 1 To arrPays_NB
    If xYNOTPAY0.NOTPAYISO = arrPAYS_ISO(K) Then Exit For
Next K
For K2 = 1 To selYNOTPAY0_Nb
    If xYNOTPAY0.NOTPAYISO = selYNOTPAY0(K2).NOTPAYISO Then Exit For
Next K2

fgSelect.Col = 0: fgSelect.Text = xYNOTPAY0.NOTPAYISO & " - " & arrPays_Lib(K)
fgSelect.Col = 1: fgSelect.Text = dateImp(xYNOTPAY0.NOTPAYHAMJ)
fgSelect.CellBackColor = wColor
fgSelect.Col = 2: fgSelect.Text = xYNOTPAY0.NOTPAYPROV
fgSelect.CellAlignment = 4
fgSelect.Col = 3:
fgSelect.CellFontBold = True
If xYNOTPAY0.NOTPAYCOFA <> selYNOTPAY0(K2).NOTPAYCOFA Then
    fgSelect.Text = selYNOTPAY0(K2).NOTPAYCOFA & " > " & xYNOTPAY0.NOTPAYCOFA
    fgSelect.CellBackColor = wColor  'RGB(128, 255, 128)
    fgSelect.CellForeColor = vbRed
Else
    fgSelect.Text = xYNOTPAY0.NOTPAYCOFA
    fgSelect.CellBackColor = wColor
End If
fgSelect.CellAlignment = 4
fgSelect.Col = 4: fgSelect.Text = xYNOTPAY0.NOTPAYCOF2
fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor
fgSelect.CellAlignment = 4

fgSelect.Col = 5: fgSelect.Text = xYNOTPAY0.NOTPAYCOFK & " " & dateImpS(xYNOTPAY0.NOTPAYCOFD)
fgSelect.CellFontSize = 6

fgSelect.Col = 6
fgSelect.CellFontBold = True
If xYNOTPAY0.NOTPAYOCDE <> selYNOTPAY0(K2).NOTPAYOCDE Then
    fgSelect.Text = selYNOTPAY0(K2).NOTPAYOCDE & " > " & xYNOTPAY0.NOTPAYOCDE
    fgSelect.CellBackColor = wColor  'RGB(128, 255, 128)
    fgSelect.CellForeColor = vbRed
Else
    fgSelect.Text = xYNOTPAY0.NOTPAYOCDE
    fgSelect.CellBackColor = wColor
End If

fgSelect.CellAlignment = 4
fgSelect.Col = 7: fgSelect.Text = xYNOTPAY0.NOTPAYOCDK & " " & dateImpS(xYNOTPAY0.NOTPAYOCDD)
fgSelect.CellFontSize = 6

fgSelect.Col = 8
fgSelect.CellFontBold = True
If xYNOTPAY0.NOTPAYSP <> selYNOTPAY0(K2).NOTPAYSP Then
    fgSelect.Text = selYNOTPAY0(K2).NOTPAYSP & " > " & xYNOTPAY0.NOTPAYSP
    fgSelect.CellBackColor = wColor  'RGB(128, 255, 128)
    fgSelect.CellForeColor = vbRed
Else
    fgSelect.Text = xYNOTPAY0.NOTPAYSP
    fgSelect.CellBackColor = wColor
End If

fgSelect.CellAlignment = 4
fgSelect.Col = 9: fgSelect.Text = xYNOTPAY0.NOTPAYSPK & " " & dateImpS(xYNOTPAY0.NOTPAYSPD)
fgSelect.CellFontSize = 6

If xYNOTPAY0.NOTPAYCEG <> 0 Then
    fgSelect.Col = 10: fgSelect.Text = xYNOTPAY0.NOTPAYCEG
    fgSelect.CellFontBold = True: fgSelect.CellBackColor = RGB(128, 255, 128)
    fgSelect.CellAlignment = 4
End If

fgSelect.Col = 11
fgSelect.CellFontBold = True
If xYNOTPAY0.NOTPAYBIAN <> selYNOTPAY0(K2).NOTPAYBIAN Then
    fgSelect.Text = selYNOTPAY0(K2).NOTPAYBIAN & " > " & xYNOTPAY0.NOTPAYBIAN
    fgSelect.CellBackColor = wColor  'RGB(128, 255, 128)
    fgSelect.CellForeColor = vbRed
Else
    fgSelect.Text = xYNOTPAY0.NOTPAYBIAN
    fgSelect.CellBackColor = wColor
End If

fgSelect.CellAlignment = 4
fgSelect.Col = 12: fgSelect.Text = xYNOTPAY0.NOTPAYBIAK & " " & dateImpS(xYNOTPAY0.NOTPAYBIAD)
fgSelect.CellFontSize = 6

fgSelect.Col = 13: fgSelect.Text = Format$(xYNOTPAY0.NOTPAYTAUX, " ##0.00")
fgSelect.CellFontBold = True: fgSelect.CellBackColor = wColor

fgSelect.Col = 14: fgSelect.Text = xYNOTPAY0.NOTPAYFISC
fgSelect.Col = 15: fgSelect.Text = Format$(xYNOTPAY0.NOTPAYSEQ, " ##0") & " - " & xYNOTPAY0.NOTPAYXUSR & " " & dateImpS(xYNOTPAY0.NOTPAYXAMJ) & " " & timeImp8(xYNOTPAY0.NOTPAYXHMS)
fgSelect.Col = 16: fgSelect.Text = Trim(xYNOTPAY0.NOTPAYTXT)

fgSelect.Col = 17: fgSelect.Text = Trim(xYNOTPAY0.JORCV)
fgSelect.Col = 18: fgSelect.Text = Trim(xYNOTPAY0.JOSEQN)

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
End Sub




Public Sub fgPays_Sort()
If fgPays.Rows > 1 Then
    fgPays.Row = 1
    fgPays.RowSel = fgPays.Rows - 1
    
    If fgPays_Sort1_Old = fgPays_Sort1 Then
        If fgPays_SortAD = 5 Then
            fgPays_SortAD = 6
        Else
            fgPays_SortAD = 5
        End If
    Else
        fgPays_SortAD = 5
    End If
    fgPays_Sort1_Old = fgPays_Sort1
    
    fgPays.Col = fgPays_Sort1
    fgPays.ColSel = fgPays_Sort2
    fgPays.Sort = fgPays_SortAD
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

Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = lK
    X = Format$(Val(fgSelect.Text), "0000000")
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
    'Select Case lK
    '    Case 1, 2: fgSelect.Text = X
    'End Select
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
Call BiaPgmAut_Init(wFct, YNOTPAY0_Aut)

'blnSetfocus = True
Form_Init

Select Case wFct
    Case Else: blnAuto = False
End Select

End Sub


Public Sub Form_Init()
Dim V, xSQL As String, X As String

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

lstErr.Visible = True

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False
chkSelect_NOTPAYSEQ.Visible = False

fraDetail.Visible = False
Set fraDetail.Container = fraTab0
fraDetail.Left = 8100
fraDetail.Top = fgSelect.Top

fraParam.Visible = False
Set fraParam.Container = fraTab0
fraParam.Left = fgSelect.Left + 1700
fraParam.Top = fgSelect.Top + 400


'fgSAb_Client_FormatString = fgSAB_Client.FormatString
fgSAB_Client.Visible = False
Set fgSAB_Client.Container = fraTab0
fgSAB_Client.Left = fgSelect.Left
fgSAB_Client.Top = fgSelect.Top
fgSAB_Client.Height = fgSelect.Height

fgSAB_Client_Detail.Visible = False
Set fgSAB_Client_Detail.Container = fraTab0
fgSAB_Client_Detail.Left = fgSAB_Client.Left + fgSAB_Client.Width + 100
fgSAB_Client_Detail.Top = fgSelect.Top
fgSAB_Client_Detail.Height = fgSelect.Height

fraJRNENT0.Visible = False
Set fraJRNENT0.Container = fraTab0
fraJRNENT0.Left = fraDetail.Left - fraJRNENT0.Width
fraJRNENT0.Top = fgSelect.Top

cboNOTPAYLOGK.Visible = False
Set cboNOTPAYLOGK.Container = fraSelect_Options
cboNOTPAYLOGK.Left = txtSelect_NOTPAYISO.Left
cboNOTPAYLOGK.Top = txtSelect_NOTPAYISO.Top

Set txtDetail.Container = fraDetail

fgPays_FormatString = fgPays.FormatString
fraPays.Visible = False
txtPAYS_ISO.Enabled = False
txtPAYS_SAB.Enabled = False

cmdPays_Update.Visible = YNOTPAY0_Aut.Valider

fraSelect_Import.Visible = False
Set fraSelect_Import.Container = fraTab0
fraSelect_Import.Left = fraSelect_Options.Left
fraSelect_Import.Top = fraSelect_Options.Top

cmdReset

lstPays.Visible = False
lstPays.Left = 5800 '6300
lstPays.Top = fgSelect.Top
Set lstPays.Container = fraTab0


txtDetail.BackColor = &HA0E0FF ' &HC0FFC0
txtDetail_ColorUpdate = &HD0FFD0    '&HA0E0FF
Call DTPicker_Set(txtSelect_NOTPAYXAMJ, YBIATAB0_DATE_CPT_J) '

ddsYNOTPAY0_Init

param_Init

cmdParam_Update.Visible = YNOTPAY0_Aut.Valider
cmdParam_New.Visible = YNOTPAY0_Aut.Valider
cmdParam_Delete.Visible = YNOTPAY0_Aut.Valider

cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1  - Notation pays"
cboSelect_SQL.AddItem "1h - Notation pays à une date d'arrêté"
cboSelect_SQL.AddItem "Ex - Exportation .xls "
cboSelect_SQL.AddItem "Em - Exportation  mail "
cboSelect_SQL.AddItem "Sc - SAB : contrôle pays résidence Client"
cboSelect_SQL.AddItem "Sp - SAB : contrôle CATEG* / S&P"
If YNOTPAY0_Aut.Valider Then
    cboSelect_SQL.AddItem "J  - Consultation de la journalisation"
    cboSelect_SQL.AddItem "L  - Consultation du suivi des actions"
    cboSelect_SQL.AddItem "I - Importation Coface, OCDE, S & P"
    cboSelect_SQL.AddItem "Pn - Paramétrage des pondérations BIA"
    '''cboSelect_SQL.AddItem "Pp - Paramétrage des pays par source(Coface,SP) "
End If
If YNOTPAY0_Aut.Comptabiliser Then
    cboSelect_SQL.AddItem "Ec - Exportation Base comptabilité"
End If


If YNOTPAY0_Aut.Xspécial Then
    cboSelect_SQL.AddItem "Rn - Reprise notation (BIA *.xls)"
    cboSelect_SQL.AddItem "Rp - Reprise pays Import(OCDE, Coface,SP)"
    cboSelect_SQL.AddItem "Rb - Reprise BIA pondération"
End If

cboSelect_SQL.ListIndex = 0

If paramIBM_Library_SAB = "SAB073U" Then
    paramIBM_Library_SABJRN = "SAB073JRN"
    MsgBox "Form_Init : paramIBM_Library_SABJRN=SAB073JRN"
End If
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
fraDetail.ForeColor = vbRed
fraParam.ForeColor = vbRed
fraJRNENT0.ForeColor = vbRed
lblDetail_Coface.ForeColor = RGB(0, 0, 255)
lblDetail_NOTPAYCOF2.ForeColor = RGB(0, 0, 255)
lblDetail_OCDE.ForeColor = RGB(0, 0, 255)
lblDetail_SP.ForeColor = RGB(0, 0, 255)
lblDetail_BIA.ForeColor = RGB(0, 0, 255)
libDetail_NOTPAYTAUX.ForeColor = RGB(0, 0, 255)
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

Public Sub fgPays_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgPays.Visible = False
mRow = fgPays.Row

If lRow > 0 And lRow < fgPays.Rows Then
    fgPays.Row = lRow
    For I = fgPays_arrIndex To 0 Step -1
        fgPays.Col = I: fgPays.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgPays.Row = mRow
    If fgPays.Row > 0 Then
        lRow = fgPays.Row
        lColor_Old = fgPays.CellBackColor
        For I = fgPays_arrIndex To 0 Step -1
          fgPays.Col = I: fgPays.CellBackColor = lColor
        Next I
    End If
End If
fgPays.LeftCol = 0
fgPays.Visible = True
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

Private Sub cboDetail_NOTPAYBIAN_Click()
If fraDetail_Update.Enabled Then blnDetail_NOTPAYBIAK = True: fraDetail_Control
End Sub

Private Sub cboDetail_NOTPAYCOF2_Click()
If fraDetail_Update.Enabled Then fraDetail_Control

End Sub

Private Sub cboDetail_NOTPAYCOFA_Click()
If fraDetail_Update.Enabled Then fraDetail_Control
End Sub

Private Sub cboDetail_NOTPAYOCDE_Click()
If fraDetail_Update.Enabled Then fraDetail_Control

End Sub

Private Sub cboDetail_NOTPAYSP_Click()
If fraDetail_Update.Enabled Then fraDetail_Control

End Sub

Private Sub cboNOTPAYLOGK_Click()
cmdSelect_Reset

End Sub

Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

End Sub


Private Sub chkDetail_NOTPAYPROV_Click()
If fraDetail_Update.Enabled Then fraDetail_Control

End Sub

Private Sub chkSelect_NOTPAYSEQ_Click()
cmdSelect_Reset

End Sub

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdDetail_Copy_Click()
Dim blnOk As Boolean
Dim K As Integer, K2 As Integer

lstPays.Clear
lstPays.AddItem "  " & " - ABANDON DE LA COPIE"
lstPays.AddItem "  " & " - ==================="
For K = 1 To arrPays_NB
    blnOk = False
    For K2 = 1 To arrYNOTPAY0_Nb
        If arrPAYS_ISO(K) = arrYNOTPAY0(K2).NOTPAYISO Then blnOk = True: Exit For
    Next K2
    If Not blnOk Then lstPays.AddItem arrPAYS_ISO(K) & " - " & arrPays_Lib(K)
Next K
lstPays.Visible = True: blnDetail_Copy = True


End Sub

Private Sub cmdDetail_Delete_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

X = MsgBox("Confirmez-vous la suppression de l'enregistrement : " & oldYNOTPAY0.NOTPAYISO, vbYesNo, "Suppresion de la notation")
If X = vbYes Then
    cmdYNOTPAY0_Update "Delete"
    fraDetail.Visible = False
'++++++++++++++++++++++++++++++++++++++++++
    newYNOTPAYLOG.NOTPAYLOGK = "Pays Delete"
    newYNOTPAYLOG.NOTPAYLOGX = Trim(oldYNOTPAY0.NOTPAYISO) & " : " & Trim(oldYNOTPAY0.NOTPAYLIB)
    cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

    cmdSelect_Ok_Click
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdDetail_New_Click()
Dim V, xSQL As String
Dim mSEQ As Long

Me.Enabled = False: Me.MousePointer = vbHourglass

xSQL = "select NOTPAYISO , NOTPAYSEQ from " & paramIBM_Library_SABSPE & ".YNOTPAY0 " _
     & " where NOTPAYISO = '" & oldYNOTPAY0.NOTPAYISO & "' order by NOTPAYSEQ  desc"
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    Call MsgBox(xSQL, vbCritical, "cmdDetail_Update_Click : requête 1 sans réponse")
    GoTo Exit_sub
End If
mSEQ = rsSab("NOTPAYSEQ")

lastYNOTPAY0 = oldYNOTPAY0

If IsNull(cmdYNOTPAY0_Control) Then
    lastYNOTPAY0.NOTPAYSEQ = mSEQ + 1

    cmdYNOTPAY0_Update "Update + New"
    fraDetail.Visible = False
    '++++++++++++++++++++++++++++++++++++++++++
    newYNOTPAYLOG.NOTPAYLOGK = "Pays Archive"
    newYNOTPAYLOG.NOTPAYLOGX = Trim(oldYNOTPAY0.NOTPAYISO) & " : " & Trim(oldYNOTPAY0.NOTPAYLIB)
    cmdYNOTPAYLOG_New
    '++++++++++++++++++++++++++++++++++++++++++

    cmdSelect_Ok_Click
End If

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdDetail_Quit_Click()
fraDetail.Visible = False
fraJRNENT0.Visible = False

End Sub

Private Sub cmdDetail_Update_Click()
Dim V, xSQL As String

Me.Enabled = False: Me.MousePointer = vbHourglass
Call fraDetail_Control
xSQL = "select NOTPAYISO , NOTPAYSEQ from " & paramIBM_Library_SABSPE & ".YNOTPAY0 " _
     & " where NOTPAYISO = '" & oldYNOTPAY0.NOTPAYISO & "' order by NOTPAYSEQ  desc"
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    Call MsgBox(xSQL, vbCritical, "cmdDetail_Update_Click : requête 1 sans réponse")
    GoTo Exit_sub
End If

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YNOTPAY0 " _
     & " where NOTPAYISO = '" & oldYNOTPAY0.NOTPAYISO _
     & "'and NOTPAYSEQ =" & rsSab("NOTPAYSEQ")
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    Call MsgBox(xSQL, vbCritical, "cmdDetail_Update_Click : requête 2 sans réponse")
    GoTo Exit_sub
End If
V = rsYNOTPAY0_GetBuffer(rsSab, lastYNOTPAY0)

If lastYNOTPAY0.NOTPAYSEQ = 0 Then
    lastYNOTPAY0.NOTPAYCOFD = 20090531
    lastYNOTPAY0.NOTPAYOCDD = 20090531
    lastYNOTPAY0.NOTPAYSPD = 20090531
    lastYNOTPAY0.NOTPAYBIAD = 20090531
    
End If

If Not IsNull(V) Then
    Call MsgBox(xSQL, vbCritical, "cmdDetail_Update_Click : erreur décodage")
    GoTo Exit_sub
End If

    If IsNull(cmdYNOTPAY0_Control) Then
        cmdYNOTPAY0_Update "Update"
        fraDetail.Visible = False
        '++++++++++++++++++++++++++++++++++++++++++
        newYNOTPAYLOG.NOTPAYLOGK = "Pays Update"
        newYNOTPAYLOG.NOTPAYLOGX = Trim(oldYNOTPAY0.NOTPAYISO) & " : " & Trim(oldYNOTPAY0.NOTPAYLIB)
        cmdYNOTPAYLOG_New
        '++++++++++++++++++++++++++++++++++++++++++

        cmdSelect_Ok_Click
    End If

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdYNOTPAY0_Update(lFct As String)
Dim K As Integer

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

If lFct = "Update + New" Or lFct = "Update" Then

    V = sqlYNOTPAY0_Update(newYNOTPAY0, oldYNOTPAY0)
    If Not IsNull(V) Then GoTo Error_MsgBox
    
    If lFct = "Update + New" Then V = sqlYNOTPAY0_Insert(lastYNOTPAY0)
End If
'________________________________________________________________________________
If lFct = "New" Then V = sqlYNOTPAY0_Insert(newYNOTPAY0)

If lFct = "Delete" Then V = sqlYNOTPAY0_Delete(oldYNOTPAY0)
If lFct = "Delete_Import" Then V = sqlYNOTPAY0_Delete_Where(" Where NOTPAYSEQ = -1")
'________________________________________________________________________________
If lFct = "Import" Then
    For K = 1 To arrYNOTPAY0_Nb
        If Trim(selYNOTPAY0(K).NOTPAYXUSR) = "#" Then
            V = sqlYNOTPAY0_Update(selYNOTPAY0(K), arrYNOTPAY0(K))
        End If
    Next K
End If

'________________________________________________________________________________

If Not IsNull(V) Then GoTo Error_MsgBox
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, Me.Name & " ~ " & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Sub

Private Sub cmdParam_Delete_Click()
Dim xSQL As String

Me.Enabled = False: Me.MousePointer = vbHourglass

X = MsgBox("Confirmez-vous la suppression de l'enregistrement : " & oldYNOTPAY0.NOTPAYTXT, vbYesNo, "Suppresion de la notation")
If X = vbYes Then
    cmdYNOTPAY0_Update "Delete"
    fraDetail.Visible = False
    cmdParam_Quit_Click

    cmdSelect_Ok_Click
'++++++++++++++++++++++++++++++++++++++++++
    newYNOTPAYLOG.NOTPAYLOGK = "Param Delete"
    newYNOTPAYLOG.NOTPAYLOGX = Trim(oldYNOTPAY0.NOTPAYTXT) & " : " & Format(oldYNOTPAY0.NOTPAYTAUX, "##0.00") & " %"
    cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++
    Param_Load_Pondération

End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_New_Click()
Dim V, xSQL As String
Dim mSEQ As Long, wCode As String
Dim blnOk As Boolean, K As Integer

If Not txtParam_Code.Enabled Then
    txtParam_Code.Enabled = True
    GoTo Exit_sub
End If

Me.Enabled = False: Me.MousePointer = vbHourglass

newYNOTPAY0 = oldYNOTPAY0

newYNOTPAY0.NOTPAYTAUX = Val(txtParam_Taux)
wCode = Trim(txtParam_Code)
newYNOTPAY0.NOTPAYTXT = lblParam_Code & " " & wCode

Select Case lblParam_Code
    Case "1-COFACE": newYNOTPAY0.NOTPAYCOFA = wCode
    Case "2-OCDE": newYNOTPAY0.NOTPAYOCDE = wCode
    Case "2-S&P": newYNOTPAY0.NOTPAYSP = wCode
    Case "4-BIA":
        Select Case Len(wCode)
            Case 0: newYNOTPAY0.NOTPAYBIAN = "   "
            Case 1: newYNOTPAY0.NOTPAYBIAN = "  " & wCode
            Case 2: newYNOTPAY0.NOTPAYBIAN = " " & wCode
            Case Else: newYNOTPAY0.NOTPAYBIAN = wCode
        End Select
    Case Else
            Call MsgBox(lblParam_Code & " type inconnu", vbCritical, "cmdParam_New_Click ")
            GoTo Exit_sub
End Select

blnOk = True
For K = 1 To arrYNOTPAY0_Nb
    If newYNOTPAY0.NOTPAYCOFA = arrYNOTPAY0(K).NOTPAYCOFA _
    And newYNOTPAY0.NOTPAYOCDE = arrYNOTPAY0(K).NOTPAYOCDE _
    And newYNOTPAY0.NOTPAYSP = arrYNOTPAY0(K).NOTPAYSP _
    And newYNOTPAY0.NOTPAYBIAN = arrYNOTPAY0(K).NOTPAYBIAN Then
        blnOk = False
        Call MsgBox("Cet enregistrement existe déjà", vbCritical, "Paramétrage : nouvel enregistrement")
        GoTo Exit_sub
    End If
    
Next K

xSQL = "select NOTPAYSEQ from " & paramIBM_Library_SABSPE & ".YNOTPAY0 " _
     & " where NOTPAYISO = '$$' order by NOTPAYSEQ  desc"
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    Call MsgBox(xSQL, vbCritical, "cmdParam_New_Click : requête 1 sans réponse")
    GoTo Exit_sub
End If
mSEQ = rsSab("NOTPAYSEQ")

    cmdParam_Quit_Click
    newYNOTPAY0.NOTPAYSEQ = mSEQ + 1

    cmdYNOTPAY0_Update "New"
    fraDetail.Visible = False
    '++++++++++++++++++++++++++++++++++++++++++
    newYNOTPAYLOG.NOTPAYLOGK = "Param New"
    newYNOTPAYLOG.NOTPAYLOGX = Trim(newYNOTPAY0.NOTPAYISO) & " : " & Trim(newYNOTPAY0.NOTPAYLIB) & " : " & Format(newYNOTPAY0.NOTPAYTAUX, "##0.00") & " %"
    cmdYNOTPAYLOG_New
    '++++++++++++++++++++++++++++++++++++++++++
    Param_Load_Pondération
    cmdSelect_Ok_Click

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Quit_Click()

fraParam.Visible = False
lstPays.Visible = False

End Sub

Private Sub cmdParam_Update_Click()
Dim V, xSQL As String, X As String
Dim mSEQ As Long, wCode As String
Dim blnOk As Boolean, K As Integer, J As Integer
Dim wFile_Log As String
Dim newParam As typeYNOTPAY0, oldParam As typeYNOTPAY0
Dim X40 As String, wNOTPAYBIAN As String

On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass

 '______________________________________________________________________
Call Param_Load_Pondération

oldParam = oldYNOTPAY0
newParam = oldYNOTPAY0
newParam.NOTPAYTAUX = Val(txtParam_Taux)
If Trim(newParam.NOTPAYCOFA) <> "" Then
    For K = 1 To arrCoface_Nb
        If newParam.NOTPAYCOFA = arrCoface(K).NOTPAYCOFA Then
            arrCoface(K).NOTPAYTAUX = newParam.NOTPAYTAUX
            Exit For
        End If
    Next K
End If

If Trim(newParam.NOTPAYOCDE) <> "" Then
    For K = 1 To arrOCDE_Nb
        If newParam.NOTPAYOCDE = arrOCDE(K).NOTPAYOCDE Then
            arrOCDE(K).NOTPAYTAUX = newParam.NOTPAYTAUX
            Exit For
        End If
    Next K
End If
If Trim(newParam.NOTPAYSP) <> "" Then
    For K = 1 To arrSP_Nb
        If newParam.NOTPAYSP = arrSP(K).NOTPAYSP Then
            arrSP(K).NOTPAYTAUX = newParam.NOTPAYTAUX
            Exit For
        End If
    Next K
End If
If Trim(newParam.NOTPAYBIAN) <> "" Then
    For K = 1 To arrBIAN_Nb
        If newParam.NOTPAYBIAN = arrBIAN(K).NOTPAYBIAN Then
            arrBIAN(K).NOTPAYTAUX = newParam.NOTPAYTAUX
            Exit For
        End If
    Next K
End If

 '______________________________________________________________________
    
arrYNOTPAY0_SQL " where NOTPAYSEQ =  0 order by NOTPAYISO "
    
Dim kNOTPAYBIAN As Integer
ReDim selYNOTPAY0(UBound(arrYNOTPAY0))
For J = 1 To arrYNOTPAY0_Nb

    selYNOTPAY0(J) = arrYNOTPAY0(J)
    kNOTPAYBIAN = selYNOTPAY0(J).NOTPAYCEG
    blnOk = False
    For K = 1 To arrCoface_Nb
        If selYNOTPAY0(J).NOTPAYCOFA = arrCoface(K).NOTPAYCOFA Then
            kNOTPAYBIAN = kNOTPAYBIAN + arrCoface(K).NOTPAYTAUX * 2
            Exit For
        End If
    Next K
    For K = 1 To arrOCDE_Nb
        If selYNOTPAY0(J).NOTPAYOCDE = arrOCDE(K).NOTPAYOCDE Then
            kNOTPAYBIAN = kNOTPAYBIAN + arrOCDE(K).NOTPAYTAUX
            Exit For
        End If
    Next K
    X = kNOTPAYBIAN
    Select Case Len(X)
        Case 0: wNOTPAYBIAN = "   "
        Case 1: wNOTPAYBIAN = "  " & X
        Case 2: wNOTPAYBIAN = " " & X
        Case Else: wNOTPAYBIAN = X
    End Select
    If selYNOTPAY0(J).NOTPAYBIAK = "M" Then wNOTPAYBIAN = selYNOTPAY0(J).NOTPAYBIAN
    For K = 1 To arrBIAN_Nb
        If wNOTPAYBIAN = arrBIAN(K).NOTPAYBIAN Then
            If selYNOTPAY0(J).NOTPAYBIAK <> "M" Then
                selYNOTPAY0(J).NOTPAYBIAN = wNOTPAYBIAN
                selYNOTPAY0(J).NOTPAYBIAK = "A"
                selYNOTPAY0(J).NOTPAYBIAD = DSys
            End If
            selYNOTPAY0(J).NOTPAYTAUX = arrBIAN(K).NOTPAYTAUX
            blnOk = True: Exit For
        End If
    Next K
Next J

'_______________________________________________________________________________
'______________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YNOTPAY0' and BIATABK1 = 'Log' and BIATABK2 = ''"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    V = rsYBIATAB0_GetBuffer(rsSab, Old_YBIATAB0)
Else
    Old_YBIATAB0.BIATABTXT = "C:\Temp\Notation Pays\"
End If


'_________________________________________
wFile_Log = Trim(Old_YBIATAB0.BIATABTXT) & "Modif pondération " & DSYS_Time & ".log"
Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile_Log): DoEvents
Call FEU_ROUGE
Open wFile_Log For Output As #4

Print #4, "Modification du paramétrage "
Print #4, "================================================================="
Print #4, "fichier log : " & wFile_Log
Print #4, "-----------------------------------------------------------------"
Print #4, ""
For J = 1 To arrYNOTPAY0_Nb

    If selYNOTPAY0(J).NOTPAYBIAN <> arrYNOTPAY0(J).NOTPAYBIAN _
    Or selYNOTPAY0(J).NOTPAYTAUX <> arrYNOTPAY0(J).NOTPAYTAUX Then
        X40 = arrYNOTPAY0(J).NOTPAYISO & " ; " & arrYNOTPAY0(J).NOTPAYLIB
        Print #4, X40 & Space$(40 - Len(X40)) & " ; " & arrYNOTPAY0(J).NOTPAYBIAN & " = " & Format$(arrYNOTPAY0(J).NOTPAYTAUX, "##0.00") & " %" & " ; " & selYNOTPAY0(J).NOTPAYBIAN & " = " & Format$(selYNOTPAY0(J).NOTPAYTAUX, "##0.00") & " %"

    End If
Next J


Print #4, ""
Print #4, "-----------------------------------------------------------------"

Print #4, ""
Print #4, "-----------------------------------------------------------------"
Print #4, "Nb BIA traités/gérés ; " & " / " & arrYNOTPAY0_Nb
Print #4, ""
Print #4, "================================================================="

Close #4
Call FEU_VERT

Call cmdSelect_Import_Log(wFile_Log)

X = MsgBox("Voulez-vous valider cette modification de pondération ?", vbQuestion & vbYesNo, "Mise à jour de la pondération")
If X <> vbYes Then
    cmdSelect_SQL_P
    GoTo Exit_sub
End If
'________________________________________________________________________________

 '______________________________________________________________________
oldYNOTPAY0 = oldParam
newYNOTPAY0 = newParam
 '______________________________________________________________________

    cmdYNOTPAY0_Update "Update"
    fraDetail.Visible = False
    cmdParam_Quit_Click
    '++++++++++++++++++++++++++++++++++++++++++
    newYNOTPAYLOG.NOTPAYLOGK = "Param Update"
    newYNOTPAYLOG.NOTPAYLOGX = Trim(newYNOTPAY0.NOTPAYISO) & " : " & Trim(newYNOTPAY0.NOTPAYLIB) & " : " & Format(newYNOTPAY0.NOTPAYTAUX, "##0.00") & " %"
    cmdYNOTPAYLOG_New
    '++++++++++++++++++++++++++++++++++++++++++

'___________________________________________________
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

For J = 1 To arrYNOTPAY0_Nb

    If selYNOTPAY0(J).NOTPAYBIAN <> arrYNOTPAY0(J).NOTPAYBIAN _
    Or selYNOTPAY0(J).NOTPAYTAUX <> arrYNOTPAY0(J).NOTPAYTAUX Then

        V = sqlYNOTPAY0_Update(selYNOTPAY0(J), arrYNOTPAY0(J))
        If Not IsNull(V) Then GoTo Error_MsgBox
    End If
Next J
    
GoTo Exit_Transaction

'------------------------------------------
Error_Handler:
   V = Error
    If V = "Chemin d'accès introuvable" Then
        wFile_Log = "C:\Temp\Modif pondération " & DSYS_Time & ".log"
        Open wFile_Log For Output As #4
        Resume Next
    End If

Error_MsgBox:
    MsgBox V, vbCritical, Me.Name & " ~ " & App_Debug
Exit_Transaction:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

cmdSelect_Ok_Click


Exit_sub:
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdPays_Delete_Click()
Dim xSQL As String
On Error GoTo Error_Handler

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> cmdPays_Delete en cours ......"): DoEvents
Dim V
App_Debug = "cmdPays_Delete"

Old_YBIATAB0.BIATABID = "YNOTPAY0"
Old_YBIATAB0.BIATABK1 = "Pays"
Old_YBIATAB0.BIATABK2 = Trim(txtPAYS_ISO)

Call Parametrage_Delete
fraPays.Visible = False
'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Plib Delete"
newYNOTPAYLOG.NOTPAYLOGX = Trim(Old_YBIATAB0.BIATABK2) & " : " & Trim(txtPAYS_SAB)
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++


'Pays import_____________________________________________________________________________

arrPays_Import_SQL
fgPays_Display

SSTab1.Tab = 1
GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdPays_Delete"
    
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

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


Private Sub cmdPays_New_Click()
Dim xSQL As String
On Error GoTo Error_Handler

If Not txtPAYS_ISO.Enabled Then
    txtPAYS_ISO.Enabled = True
    txtPAYS_SAB = ""
    Exit Sub
End If

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> cmdPays_New en cours ......"): DoEvents
Dim V
App_Debug = "cmdPays_New"



New_YBIATAB0.BIATABID = "YNOTPAY0"
New_YBIATAB0.BIATABK1 = "Pays"
New_YBIATAB0.BIATABK2 = Trim(txtPAYS_ISO)
New_YBIATAB0.BIATABTXT = ""
Mid$(New_YBIATAB0.BIATABTXT, 1, 3) = Trim(txtPAYS_OCDE_Code)
Mid$(New_YBIATAB0.BIATABTXT, 4, 32) = LCase$(Trim(txtPAYS_OCDE_lib))
Mid$(New_YBIATAB0.BIATABTXT, 36, 32) = LCase$(Trim(txtPAYS_Coface))
Mid$(New_YBIATAB0.BIATABTXT, 68, 32) = LCase$(Trim(txtPAYS_SP))
If Trim(New_YBIATAB0.BIATABTXT) = "" Then
    MsgBox "préciser au moins 1 libellé", vbCritical, "cmdPays_New"
    GoTo Exit_sub
End If

Call Parametrage_New
fraPays.Visible = False
'++++++++++++++++++++++++++++++++++++++++++
If Trim(Mid$(New_YBIATAB0.BIATABTXT, 1, 35)) <> "" Then
    newYNOTPAYLOG.NOTPAYLOGK = "Plib new OCDE"
    newYNOTPAYLOG.NOTPAYLOGX = Trim(New_YBIATAB0.BIATABK2) & " : " & Mid$(New_YBIATAB0.BIATABTXT, 1, 35)
    cmdYNOTPAYLOG_New
End If
If Trim(Mid$(New_YBIATAB0.BIATABTXT, 36, 32)) <> "" Then
    newYNOTPAYLOG.NOTPAYLOGK = "Plib new COFACE"
    newYNOTPAYLOG.NOTPAYLOGX = Trim(New_YBIATAB0.BIATABK2) & " : " & Mid$(New_YBIATAB0.BIATABTXT, 36, 32)
    cmdYNOTPAYLOG_New
End If
If Trim(Mid$(New_YBIATAB0.BIATABTXT, 68, 32)) <> "" Then
    newYNOTPAYLOG.NOTPAYLOGK = "Plib new S&P"
    newYNOTPAYLOG.NOTPAYLOGX = Trim(New_YBIATAB0.BIATABK2) & " : " & Mid$(New_YBIATAB0.BIATABTXT, 68, 32)
    cmdYNOTPAYLOG_New
End If

'++++++++++++++++++++++++++++++++++++++++++


'Pays import_____________________________________________________________________________

arrPays_Import_SQL
fgPays_Display

SSTab1.Tab = 1
GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdPays_New"
    
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdPays_Quit_Click()
fraPays.Visible = False

End Sub

Private Sub cmdPays_Update_Click()
Dim xSQL As String
On Error GoTo Error_Handler

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> cmdPays_Update en cours ......"): DoEvents

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YNOTPAY0' and BIATABK1 = 'Pays'  and BIATABK2 = '" & Trim(txtPAYS_ISO) & "'"
Set rsSab = cnsab.Execute(xSQL)

If rsSab.EOF Then
    MsgBox xSQL, vbCritical, "cmdPays_Update_Click"
    GoTo Exit_sub
End If
V = rsYBIATAB0_GetBuffer(rsSab, Old_YBIATAB0)
New_YBIATAB0 = Old_YBIATAB0
New_YBIATAB0.BIATABTXT = ""
Mid$(New_YBIATAB0.BIATABTXT, 1, 3) = Trim(txtPAYS_OCDE_Code)
Mid$(New_YBIATAB0.BIATABTXT, 4, 32) = LCase$(Trim(txtPAYS_OCDE_lib))
Mid$(New_YBIATAB0.BIATABTXT, 36, 32) = LCase$(Trim(txtPAYS_Coface))
Mid$(New_YBIATAB0.BIATABTXT, 68, 32) = LCase$(Trim(txtPAYS_SP))

Call Parametrage_Update
fraPays.Visible = False
'++++++++++++++++++++++++++++++++++++++++++
If Mid$(New_YBIATAB0.BIATABTXT, 1, 35) <> Mid$(Old_YBIATAB0.BIATABTXT, 1, 35) Then
    newYNOTPAYLOG.NOTPAYLOGK = "Plib OCDE"
    newYNOTPAYLOG.NOTPAYLOGX = Trim(New_YBIATAB0.BIATABK2) & " (lib avant) : " & Mid$(Old_YBIATAB0.BIATABTXT, 1, 35)
    cmdYNOTPAYLOG_New
    newYNOTPAYLOG.NOTPAYLOGX = Trim(New_YBIATAB0.BIATABK2) & " (lib après) : " & Mid$(New_YBIATAB0.BIATABTXT, 1, 35)
    cmdYNOTPAYLOG_New
End If
If Mid$(New_YBIATAB0.BIATABTXT, 36, 32) <> Mid$(Old_YBIATAB0.BIATABTXT, 36, 32) Then
    newYNOTPAYLOG.NOTPAYLOGK = "Plib COFACE"
    newYNOTPAYLOG.NOTPAYLOGX = Trim(New_YBIATAB0.BIATABK2) & " (lib avant) : " & Mid$(Old_YBIATAB0.BIATABTXT, 36, 32)
    cmdYNOTPAYLOG_New
    newYNOTPAYLOG.NOTPAYLOGX = Trim(New_YBIATAB0.BIATABK2) & " (lib après) : " & Mid$(New_YBIATAB0.BIATABTXT, 36, 32)
    cmdYNOTPAYLOG_New
End If
If Mid$(New_YBIATAB0.BIATABTXT, 68, 32) <> Mid$(Old_YBIATAB0.BIATABTXT, 68, 32) Then
    newYNOTPAYLOG.NOTPAYLOGK = "Plib S&P"
    newYNOTPAYLOG.NOTPAYLOGX = Trim(New_YBIATAB0.BIATABK2) & " (lib avant) : " & Mid$(Old_YBIATAB0.BIATABTXT, 68, 32)
    cmdYNOTPAYLOG_New
    newYNOTPAYLOG.NOTPAYLOGX = Trim(New_YBIATAB0.BIATABK2) & " (lib après) : " & Mid$(New_YBIATAB0.BIATABTXT, 68, 32)
    cmdYNOTPAYLOG_New
End If

'++++++++++++++++++++++++++++++++++++++++++


'Pays import_____________________________________________________________________________

arrPays_Import_SQL
fgPays_Display

SSTab1.Tab = 1
GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbCritical, "cmdPays_Update"
    
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdPrint_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass
X = "Notation pays"

Select Case SSTab1.Tab
    Case 0:
        Select Case cmdSelect_SQL_K
            Case "1":
                If chkSelect_NOTPAYSEQ = "1" Then X = "Notation pays : historique"
                prtYNOTPAY0_Open "Form", X
                prtYNOTPAY0_Line fgSelect
                prtYNOTPAY0_Close
             Case "J":
                X = "Notation pays : Journalisation"
                prtYNOTPAY0_Open "Form", X
                prtYNOTPAY0_Line fgSelect
                prtYNOTPAY0_Close
            Case "I":
                If lstPays.Visible Then
                    prtYNOTPAY0_Open "lstX", "Compte-rendu d'importation"
                    prtYNOTPAY0_lstX lstPays
                    prtYNOTPAY0_Close

                Else
                    X = "Notation pays : importation"
                    prtYNOTPAY0_Open "Form", X
                    prtYNOTPAY0_Line fgSelect
                    prtYNOTPAY0_Close
                End If
            Case "Ec":
                prtYNOTPAY0_Open "lstX", "Notation pays : Exportation Comptabilité"
                prtYNOTPAY0_lstX lstPays
                prtYNOTPAY0_Close
             Case "P": X = "Notation pays : paramétrage"
                prtYNOTPAY0_Open "Form", X
                prtYNOTPAY0_Line fgSelect
                prtYNOTPAY0_Close
           Case "L": X = "Notation pays : Suivi des actions"
                prtYNOTPAY0_Open "Log", X
                prtYNOTPAY0_Log fgSelect
                prtYNOTPAY0_Close
            Case "Sp":
                If lstPays.Visible Then
                    prtYNOTPAY0_Open "lstX", "Notation pays : Contrôle paramétrage SAB CATEG*"
                    prtYNOTPAY0_lstX lstPays
                    prtYNOTPAY0_Close
                End If
       End Select
    Case 1:
        prtYNOTPAY0_Open "Pays_Form", "liste des correspondances pays COFACE, OCDE, S & P"
        prtYNOTPAY0_Pays fgPays
        
        prtYNOTPAY0_Close
End Select

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdSelect_Import_Coface_Click()
Dim wFile As String, wFilex As String, wFile_Log As String
Dim xIn As String, xPays_ISO As String, xPays_Coface As String, xLib As String, xMsg As String
Dim X As String, Xcom As String, xNew As String
Dim Nb_Coface As Integer, Nb_Coface_X As Integer, Nb_Coface_Err As Integer
Dim Nb_BIA_X As Integer, Nb_BIA_E As Integer, Nb_BIA_D As Integer, Nb_BIA_Err As Integer
Dim K As Integer, blnCoface As Boolean, blnOk As Boolean, blnErr As Boolean
Dim K1 As Integer, K2 As Integer, K3 As Integer, kLen As Integer

Dim xSQL As String

Dim arrX(500) As String, arrX_K As Integer, arrX_Nb As Integer
Dim blnValue As Boolean

On Error GoTo Error_Handler

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Import_Coface en cours ......"): DoEvents

blnCoface = False
Nb_Coface = 0: Nb_Coface_X = 0: Nb_Coface_Err = 0
Nb_BIA_X = 0: Nb_BIA_E = 0: Nb_BIA_D = 0: Nb_BIA_Err = 0


'______________________________________________
If Not cmdSelect_Import_Control("Coface") Then GoTo Exit_sub

wFile = Trim(Coface_Notepad.BIATABTXT)
'______________________________________________

X = InputBox("par défaut : " & wFile & vbCrLf & vbCrLf _
    & Coface_DateNotation_Info & vbCrLf & vbCrLf _
    & "     ================================" _
   & vbCrLf & "     NOTATION en DATE du : " & dateImp10(wAMJMin) _
   & vbCrLf & "     ================================", "Notation Pays : nom du fichier d'import Coface", wFile)
If Trim(X) = "" Then GoTo Exit_sub

wFilex = Trim(X)
If Dir(wFilex) = "" Then
    Call MsgBox("Le fichier : " & wFile & "n'existe pas", vbCritical, "Importation Coface")
    GoTo Exit_sub
End If
'______________________________________________
If wFile <> wFilex Then
    wFile = wFilex
    Old_YBIATAB0 = Coface_Notepad
    New_YBIATAB0 = Coface_Notepad
    New_YBIATAB0.BIATABTXT = wFilex
    Parametrage_Update
    Coface_Notepad = New_YBIATAB0
End If
'_________________________________________

ReDim selYNOTPAY0(arrYNOTPAY0_Nb)
For K = 1 To arrYNOTPAY0_Nb
    selYNOTPAY0(K) = arrYNOTPAY0(K)
    selYNOTPAY0(K).NOTPAYXAMJ = 0
    selYNOTPAY0(K).NOTPAYXHMS = 0
    selYNOTPAY0(K).NOTPAYXUSR = ""
Next K

wFile_Log = wFile & "_" & DSYS_Time & ".log"
Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile): DoEvents
Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile_Log): DoEvents

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Import COF"
newYNOTPAYLOG.NOTPAYLOGX = "log COFACE " & wFile_Log
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

Open wFile For Input As #3
Call FEU_ROUGE
Open wFile_Log For Output As #4
Print #4, "Importation du fichier Coface en date du : " & dateImp10(wAMJMin) & "  " & wFile
Print #4, "================================================================="
Print #4, "fichier log : " & wFile_Log
Print #4, "-----------------------------------------------------------------"
Print #4, ""

blnCoface = True
blnOk = False
arrX_Nb = 0
Do Until EOF(3)
    Line Input #3, xIn
    xIn = Trim(xIn)
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "- " & xIn): DoEvents
'____________________________________________________________________________________________
    blnCoface = False
    
    ' __________________________________________________________________________ Le site COFACE a change de structure donc cette portion de code mise en commentaire n'est plus valable  KOKOU 15/10/2024 _____________________
    
'    If InStr(1, xIn, "<td class=" & Asc34 & "country" & Asc34 & ">") Then
'        K1 = InStr(22, xIn, ">")
'        If K1 > 0 Then
'            K2 = InStr(K1, xIn, "<")
'            If K2 > 0 Then
'                xPays_Coface = LCase$(Trim(Mid$(xIn, K1 + 1, K2 - K1 - 1)))
'                If Len(xPays_Coface) > 32 Then xPays_Coface = Mid$(xPays_Coface, 1, 32)
'                Xcom = "*$@": xNew = "*$@"
'                blnValue = False
'                Do Until EOF(3)
'                    Line Input #3, xIn
'                    xIn = Trim(xIn)
'                    If Not blnValue Then
'                        If InStr(1, xIn, "<span class=") Then blnValue = True
'                    Else
'                        blnValue = False
'                        xIn = Trim(xIn)
'                        If Len(xIn) > 2 Then
'                            K1 = InStr(1, xIn, "<")
'                            If K1 > 0 Then xIn = Mid$(xIn, 1, K1 - 1)
'                        End If
'                        If xNew = "*$@" Then
'                            xNew = xIn
'                        Else
'                            Xcom = xIn
'                            blnCoface = True
'                        End If
'
'                    End If
'
'                    If blnCoface Then Exit Do
'                Loop
'            End If
'        End If
'    End If
    
    
' _______________________________________________________________ Ajout de KOKOU suite à la modification de la structure du site officiel de COFACE __________________________________________________

   '___________________________________________________________________________________ TEST KOKOU J'ai jouté cette portion____________________________________________
    xNew = ""
    Xcom = ""
    xPays_Coface = ""
    Dim J As Integer
    J = 0
    Dim test As String
    
    If InStr(1, xIn, "class=" & Asc34 & "countryComparisonArray__cont__line__el" & Asc34 & ">") Then

        ' Déclaration des variables
        Dim htmlString, htmlDoc, spanElements, I
        
        ' Chaîne HTML à analyser
        htmlString = xIn
        
        ' Créer un objet HTMLFile pour analyser le HTML
        Set htmlDoc = CreateObject("htmlfile")
        htmlDoc.Write htmlString
        
        ' Rechercher tous les éléments <span> dans le HTML
        Set spanElements = htmlDoc.getElementsByTagName("span")
        
        ' Parcourir les nœuds pour trouver les valeurs recherchées

        xPays_Coface = LCase(spanElements.Item(0).innerText)
        If xPays_Coface <> "" Then
            blnCoface = True
        End If
        

        If spanElements.Item(1).ClassName = "style-h2" And spanElements.Item(2).ClassName = "style-h3" Then
            xNew = spanElements.Item(1).innerText & spanElements.Item(2).innerText
            J = 3
        Else
         xNew = spanElements.Item(1).innerText
         J = 2
        End If
        
        If J > 0 Then
            If J < spanElements.Length - 1 Then
                Xcom = spanElements.Item(J).innerText & spanElements.Item(J + 1).innerText
            Else
             Xcom = spanElements.Item(J).innerText
            End If
        
        End If
        
        
        
   
        
        ' Nettoyage
        Set htmlDoc = Nothing
        Set spanElements = Nothing


    End If
    
    
    '___________________________________________________________________________________ TEST KOKOU J'ai jouté cette portion


    
'____________________________________________________________________________________________

    If blnCoface Then
        Nb_Coface = Nb_Coface + 1
        
'_____________________________________________________________
        
'_______________________________________________________________
        

    
        xPays_ISO = ""
        xMsg = Space$(25) & " ; "
        
        blnOk = False
        For K = 1 To arrCoface_Nb
            If xNew = Trim(arrCoface(K).NOTPAYCOFA) Then blnOk = True: Exit For
        Next K
        If Not blnOk Then
            Nb_Coface_X = Nb_Coface_X + 1
            Mid$(xMsg, 1, 25) = "notation Coface inconnue"
        Else
            blnOk = False
            For K = 1 To arrPays_Import_Nb
                If xPays_Coface = Trim(Mid$(arrPays_Import(K).BIATABTXT, 36, 32)) Then
                    xPays_ISO = Mid$(arrPays_Import(K).BIATABK2, 1, 2)
                    blnOk = True
                    Exit For
                End If
            Next K
            If Not blnOk Then
                Nb_Coface_Err = Nb_Coface_Err + 1
                Mid$(xMsg, 1, 25) = "code pays Coface inconnu " & xPays_Coface
            Else
                blnOk = False
                For K = 1 To arrYNOTPAY0_Nb
                    If xPays_ISO = selYNOTPAY0(K).NOTPAYISO Then blnOk = True: Exit For
                Next K
                If Not blnOk Then
                    Nb_Coface_X = Nb_Coface_X + 1
                    Mid$(xMsg, 1, 25) = "code pays non géré BIA " ' & xPays_ISO
                Else
                    newYNOTPAY0 = selYNOTPAY0(K)
                    newYNOTPAY0.NOTPAYCOFA = xNew
                    newYNOTPAY0.NOTPAYCOF2 = Xcom
                    newYNOTPAY0.NOTPAYCOFK = "A"
                    newYNOTPAY0.NOTPAYCOFD = wAMJMin
                    
                    X = cmdSelect_Import_NOTPAYBIAN_Auto
                    If X <> "" Then
                        Nb_BIA_Err = Nb_BIA_Err + 1
                        Mid$(xMsg, 4, 21) = X
                    End If
                    If newYNOTPAY0.NOTPAYCOFA = selYNOTPAY0(K).NOTPAYCOFA _
                    And newYNOTPAY0.NOTPAYCOFK = selYNOTPAY0(K).NOTPAYCOFK _
                    And newYNOTPAY0.NOTPAYCOFD = selYNOTPAY0(K).NOTPAYCOFD _
                    And newYNOTPAY0.NOTPAYCOF2 = selYNOTPAY0(K).NOTPAYCOF2 _
                    And newYNOTPAY0.NOTPAYBIAN = selYNOTPAY0(K).NOTPAYBIAN Then
                        Nb_BIA_E = Nb_BIA_E + 1
                        Mid$(xMsg, 1, 1) = "="
                        newYNOTPAY0.NOTPAYXUSR = "="
                    Else
                        Mid$(xMsg, 1, 1) = "#"
                        Nb_BIA_D = Nb_BIA_D + 1
                        newYNOTPAY0.NOTPAYXUSR = "#"
                    End If
                    
                    selYNOTPAY0(K) = newYNOTPAY0
                End If
            End If
            Print #4, xMsg & xPays_ISO & " ; " & xNew & " ; (" & Xcom & ") ; " & xPays_Coface & " ; " & xLib
           End If
        'End If
    End If
    'Call lstErr_AddItem(lstErr, cmdContext, "- pays BIA non renseigné Coface"): DoEvents
Loop

Print #4, ""
Print #4, "-----------------------------------------------------------------"
For K = 1 To arrYNOTPAY0_Nb
    If Trim(selYNOTPAY0(K).NOTPAYXUSR) = "" Then
        For K2 = 1 To arrPays_NB
            If selYNOTPAY0(K).NOTPAYISO = arrPAYS_ISO(K2) Then Exit For
        Next K2

        Nb_BIA_X = Nb_BIA_X + 1
        Print #4, "pays BIA non renseigné Coface" & "; " & selYNOTPAY0(K).NOTPAYISO & " ; ; " & arrPays_Lib(K2)
        Call lstErr_AddItem(lstErr, cmdContext, "! " & selYNOTPAY0(K).NOTPAYISO): DoEvents
    End If
Next K

Print #4, ""
Print #4, "-----------------------------------------------------------------"
Print #4, "Nb Coface lus          ; " & Nb_Coface
Print #4, "Nb Coface inconnus     ; " & Nb_Coface_Err
Print #4, "Nb Coface ignorés      ; " & Nb_Coface_X
Print #4, ""
Print #4, "Nb BIA identiques    ; " & Nb_BIA_E
Print #4, "Nb BIA mis à jour    ; " & Nb_BIA_D
Print #4, "Nb BIA sans notation ;" & Nb_BIA_X
Print #4, "Nb BIA traités/gérés ; " & Nb_BIA_E + Nb_BIA_D + Nb_BIA_X & " / " & arrYNOTPAY0_Nb
Print #4, ""
Print #4, "Nb BIA en anomalie   ; " & Nb_BIA_Err
Print #4, "================================================================="

Close #3
Close #4
Call FEU_VERT

Call lstErr_AddItem(lstErr, cmdContext, "- mise à jour du fichier YNOTPAY0"): DoEvents

Call cmdYNOTPAY0_Update("Import")
Call cmdSelect_Ok_Click

Call cmdSelect_Import_Log(wFile_Log)

cmdSelect_Import_DateNotation_Update ("Coface")

GoTo Exit_sub

Error_Handler:
    MsgBox xIn & vbCrLf & Error, vbCritical, "Importation Coface"
    
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0
Close
End Sub

Private Sub cmdSelect_Import_Coface_2013()
Dim wFile As String, wFilex As String, wFile_Log As String
Dim xIn As String, xPays_ISO As String, xPays_Coface As String, xLib As String, xMsg As String
Dim X As String, Xcom As String, xNew As String
Dim Nb_Coface As Integer, Nb_Coface_X As Integer, Nb_Coface_Err As Integer
Dim Nb_BIA_X As Integer, Nb_BIA_E As Integer, Nb_BIA_D As Integer, Nb_BIA_Err As Integer
Dim K As Integer, blnCoface As Boolean, blnOk As Boolean, blnErr As Boolean
Dim K1 As Integer, K2 As Integer, K3 As Integer, kLen As Integer
Dim xSQL As String

Dim arrX(500) As String, arrX_K As Integer, arrX_Nb As Integer

On Error GoTo Error_Handler

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Import_Coface en cours ......"): DoEvents

blnCoface = False
Nb_Coface = 0: Nb_Coface_X = 0: Nb_Coface_Err = 0
Nb_BIA_X = 0: Nb_BIA_E = 0: Nb_BIA_D = 0: Nb_BIA_Err = 0


'______________________________________________
If Not cmdSelect_Import_Control("Coface") Then GoTo Exit_sub

wFile = Trim(Coface_Notepad.BIATABTXT)
'______________________________________________

X = InputBox("par défaut : " & wFile & vbCrLf & vbCrLf _
    & Coface_DateNotation_Info & vbCrLf & vbCrLf _
    & "     ================================" _
   & vbCrLf & "     NOTATION en DATE du : " & dateImp10(wAMJMin) _
   & vbCrLf & "     ================================", "Notation Pays : nom du fichier d'import Coface", wFile)
If Trim(X) = "" Then GoTo Exit_sub

wFilex = Trim(X)
If Dir(wFilex) = "" Then
    Call MsgBox("Le fichier : " & wFile & "n'existe pas", vbCritical, "Importation Coface")
    GoTo Exit_sub
End If
'______________________________________________
If wFile <> wFilex Then
    wFile = wFilex
    Old_YBIATAB0 = Coface_Notepad
    New_YBIATAB0 = Coface_Notepad
    New_YBIATAB0.BIATABTXT = wFilex
    Parametrage_Update
    Coface_Notepad = New_YBIATAB0
End If
'_________________________________________

ReDim selYNOTPAY0(arrYNOTPAY0_Nb)
For K = 1 To arrYNOTPAY0_Nb
    selYNOTPAY0(K) = arrYNOTPAY0(K)
    selYNOTPAY0(K).NOTPAYXAMJ = 0
    selYNOTPAY0(K).NOTPAYXHMS = 0
    selYNOTPAY0(K).NOTPAYXUSR = ""
Next K

wFile_Log = wFile & "_" & DSYS_Time & ".log"
Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile): DoEvents
Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile_Log): DoEvents

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Import COF"
newYNOTPAYLOG.NOTPAYLOGX = "log COFACE " & wFile_Log
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

Open wFile For Input As #3
Call FEU_ROUGE
Open wFile_Log For Output As #4
Print #4, "Importation du fichier Coface en date du : " & dateImp10(wAMJMin) & "  " & wFile
Print #4, "================================================================="
Print #4, "fichier log : " & wFile_Log
Print #4, "-----------------------------------------------------------------"
Print #4, ""

blnCoface = True
blnOk = False
arrX_Nb = 0
Do Until EOF(3)
    Line Input #3, xIn
    xIn = Trim(xIn)
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "- " & xIn): DoEvents
    If Not blnCoface Then
        If InStr(1, xIn, "@rating") Then blnCoface = True
    Else
        '=========================
        If xIn = "" Then blnCoface = False 'Exit Do
        '=========================
        If Len(xIn) > 2 Then
            arrX_Nb = arrX_Nb + 1
            arrX(arrX_Nb) = xIn
        Else
            If Len(xIn) > 0 Then arrX(arrX_Nb) = arrX(arrX_Nb) & " " & xIn
        End If
        
    End If
Loop

For arrX_K = 1 To arrX_Nb
        xIn = arrX(arrX_K)
        Xcom = "*$@": xNew = "*$@"
        Nb_Coface = Nb_Coface + 1
        kLen = Len(xIn)
        For K2 = kLen To 1 Step -1
            If Mid$(xIn, K2, 1) = " " Then Xcom = Trim(Mid$(xIn, K2 + 1, kLen - K2)): Exit For
        Next K2
        For K3 = K2 - 1 To 1 Step -1
            If Mid$(xIn, K3, 1) <> " " Then Exit For
        Next K3
        For K1 = K3 To 1 Step -1
            If Mid$(xIn, K1, 1) = " " Then xNew = Trim(Mid$(xIn, K1 + 1, K2 - K1)): Exit For
        Next K1
        
'_____________________________________________________________
        blnOk = False
        For K = 1 To arrCoface_Nb
            If xNew = Trim(arrCoface(K).NOTPAYCOFA) Then blnOk = True: Exit For
        Next K
        If Not blnOk Then
            xNew = Xcom
            Xcom = ""
            K1 = K3
        End If
        
'_______________________________________________________________
        
        If K1 > 32 Then K1 = 32
        xPays_Coface = LCase$(Trim(Mid$(xIn, 1, K1)))

    
        xPays_ISO = ""
        xMsg = Space$(25) & " ; "
        
        blnOk = False
        For K = 1 To arrCoface_Nb
            If xNew = Trim(arrCoface(K).NOTPAYCOFA) Then blnOk = True: Exit For
        Next K
        If Not blnOk Then
            Nb_Coface_X = Nb_Coface_X + 1
            Mid$(xMsg, 1, 25) = "notation Coface inconnue"
        Else
            blnOk = False
            For K = 1 To arrPays_Import_Nb
                If xPays_Coface = Trim(Mid$(arrPays_Import(K).BIATABTXT, 36, 32)) Then
                    xPays_ISO = Mid$(arrPays_Import(K).BIATABK2, 1, 2)
                    blnOk = True
                    Exit For
                End If
            Next K
            If Not blnOk Then
                Nb_Coface_Err = Nb_Coface_Err + 1
                Mid$(xMsg, 1, 25) = "code pays Coface inconnu " & xPays_Coface
            Else
                blnOk = False
                For K = 1 To arrYNOTPAY0_Nb
                    If xPays_ISO = selYNOTPAY0(K).NOTPAYISO Then blnOk = True: Exit For
                Next K
                If Not blnOk Then
                    Nb_Coface_X = Nb_Coface_X + 1
                    Mid$(xMsg, 1, 25) = "code pays non géré BIA " ' & xPays_ISO
                Else
                    newYNOTPAY0 = selYNOTPAY0(K)
                    newYNOTPAY0.NOTPAYCOFA = xNew
                    newYNOTPAY0.NOTPAYCOF2 = Xcom
                    newYNOTPAY0.NOTPAYCOFK = "A"
                    newYNOTPAY0.NOTPAYCOFD = wAMJMin
                    
                    X = cmdSelect_Import_NOTPAYBIAN_Auto
                    If X <> "" Then
                        Nb_BIA_Err = Nb_BIA_Err + 1
                        Mid$(xMsg, 4, 21) = X
                    End If
                    If newYNOTPAY0.NOTPAYCOFA = selYNOTPAY0(K).NOTPAYCOFA _
                    And newYNOTPAY0.NOTPAYCOFK = selYNOTPAY0(K).NOTPAYCOFK _
                    And newYNOTPAY0.NOTPAYCOFD = selYNOTPAY0(K).NOTPAYCOFD _
                    And newYNOTPAY0.NOTPAYCOF2 = selYNOTPAY0(K).NOTPAYCOF2 _
                    And newYNOTPAY0.NOTPAYBIAN = selYNOTPAY0(K).NOTPAYBIAN Then
                        Nb_BIA_E = Nb_BIA_E + 1
                        Mid$(xMsg, 1, 1) = "="
                        newYNOTPAY0.NOTPAYXUSR = "="
                    Else
                        Mid$(xMsg, 1, 1) = "#"
                        Nb_BIA_D = Nb_BIA_D + 1
                        newYNOTPAY0.NOTPAYXUSR = "#"
                    End If
                    
                    selYNOTPAY0(K) = newYNOTPAY0
                End If
            End If
            Print #4, xMsg & xPays_ISO & " ; " & xNew & " ; (" & Xcom & ") ; " & xPays_Coface & " ; " & xLib
           End If
        'End If
    'End If
Next arrX_K
Call lstErr_AddItem(lstErr, cmdContext, "- pays BIA non renseigné Coface"): DoEvents

Print #4, ""
Print #4, "-----------------------------------------------------------------"
For K = 1 To arrYNOTPAY0_Nb
    If Trim(selYNOTPAY0(K).NOTPAYXUSR) = "" Then
        For K2 = 1 To arrPays_NB
            If selYNOTPAY0(K).NOTPAYISO = arrPAYS_ISO(K2) Then Exit For
        Next K2

        Nb_BIA_X = Nb_BIA_X + 1
        Print #4, "pays BIA non renseigné Coface" & "; " & selYNOTPAY0(K).NOTPAYISO & " ; ; " & arrPays_Lib(K2)
        Call lstErr_AddItem(lstErr, cmdContext, "! " & selYNOTPAY0(K).NOTPAYISO): DoEvents
    End If
Next K

Print #4, ""
Print #4, "-----------------------------------------------------------------"
Print #4, "Nb Coface lus          ; " & Nb_Coface
Print #4, "Nb Coface inconnus     ; " & Nb_Coface_Err
Print #4, "Nb Coface ignorés      ; " & Nb_Coface_X
Print #4, ""
Print #4, "Nb BIA identiques    ; " & Nb_BIA_E
Print #4, "Nb BIA mis à jour    ; " & Nb_BIA_D
Print #4, "Nb BIA sans notation ;" & Nb_BIA_X
Print #4, "Nb BIA traités/gérés ; " & Nb_BIA_E + Nb_BIA_D + Nb_BIA_X & " / " & arrYNOTPAY0_Nb
Print #4, ""
Print #4, "Nb BIA en anomalie   ; " & Nb_BIA_Err
Print #4, "================================================================="

Close #3
Close #4
Call FEU_VERT

Call lstErr_AddItem(lstErr, cmdContext, "- mise à jour du fichier YNOTPAY0"): DoEvents

Call cmdYNOTPAY0_Update("Import")
Call cmdSelect_Ok_Click

Call cmdSelect_Import_Log(wFile_Log)

cmdSelect_Import_DateNotation_Update ("Coface")

GoTo Exit_sub

Error_Handler:
    MsgBox xIn & vbCrLf & Error, vbCritical, "Importation Coface"
    
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0
Close
End Sub


Private Sub cmdSelect_Import_Delete_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

X = MsgBox("Confirmez-vous la suppression des enregistrments d'import ?", vbQuestion + vbYesNo, "Importation des notations COFACE, OCDE, S & P")
If X = vbYes Then
    cmdYNOTPAY0_Update "Delete_Import"
    cmdSelect_Ok_Click
End If
'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Import ANN"
newYNOTPAYLOG.NOTPAYLOGX = "Import : ANNULATION de la préparation "
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Import_OCDE_Click()
Dim wFile As String, wFilex As String, wFile_Log As String
Dim xIn As String, xPays_ISO As String, xPays_OCDE As String, xLib As String, xMsg As String
Dim X As String, xOld As String, xNew As String
Dim Nb_OCDE As Integer, Nb_OCDE_X As Integer, Nb_OCDE_Err As Integer
Dim Nb_BIA_X As Integer, Nb_BIA_E As Integer, Nb_BIA_D As Integer, Nb_BIA_Err As Integer
Dim K As Integer, K2 As Integer, blnOk As Boolean, blnErr As Boolean
On Error GoTo Error_Handler

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Import_OCDE en cours ......"): DoEvents

Nb_OCDE = 0: Nb_OCDE_X = 0: Nb_OCDE_Err = 0
Nb_BIA_X = 0: Nb_BIA_E = 0: Nb_BIA_D = 0: Nb_BIA_Err = 0

Call DTPicker_Control(txtSelect_Import_Amj, wAMJMin)
'______________________________________________
'______________________________________________
If Not cmdSelect_Import_Control("OCDE") Then GoTo Exit_sub

wFile = Trim(OCDE_Notepad.BIATABTXT)
X = InputBox("par défaut : " & wFile & vbCrLf & vbCrLf _
    & Coface_DateNotation_Info & vbCrLf & vbCrLf _
    & "     ================================" _
   & vbCrLf & "     EN DATE du : " & dateImp10(wAMJMin) _
   & vbCrLf & "     ================================", "Notation Pays : nom du fichier d'import ocde", wFile)
If Trim(X) = "" Then GoTo Exit_sub

wFilex = Trim(X)
If Dir(wFilex) = "" Then
    Call MsgBox("Le fichier : " & wFile & "n'existe pas", vbCritical, "Importation OCDE")
    GoTo Exit_sub
End If
'______________________________________________
If wFile <> wFilex Then
    wFile = wFilex
    Old_YBIATAB0 = OCDE_Notepad
    New_YBIATAB0 = OCDE_Notepad
    New_YBIATAB0.BIATABTXT = wFilex
    Parametrage_Update
    OCDE_Notepad = New_YBIATAB0
End If
'_________________________________________

ReDim selYNOTPAY0(arrYNOTPAY0_Nb)
For K = 1 To arrYNOTPAY0_Nb
    selYNOTPAY0(K) = arrYNOTPAY0(K)
    selYNOTPAY0(K).NOTPAYXAMJ = 0
    selYNOTPAY0(K).NOTPAYXHMS = 0
    selYNOTPAY0(K).NOTPAYXUSR = ""
Next K

wFile_Log = wFile & "_" & DSYS_Time & ".log"
Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile): DoEvents
Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile_Log): DoEvents

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Import OCDE"
newYNOTPAYLOG.NOTPAYLOGX = "log OCDE " & wFile_Log
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

Open wFile For Input As #3
Call FEU_ROUGE
Open wFile_Log For Output As #4
Print #4, "Import du fichier OCDE en date du : " & dateImp10(wAMJMin) & "  " & wFile
Print #4, "================================================================="
Print #4, "fichier log : " & wFile_Log
Print #4, "-----------------------------------------------------------------"
Print #4, ""


blnOk = False

' Commentaire KOKOU j'ai mis ce bloc en commentaire parce que j'ai apporté des modifications
'Do Until EOF(3)
'    Line Input #3, xIn
'    Call lstErr_ChangeLastItem(lstErr, cmdContext, "- " & xIn): DoEvents
'
'    If IsNumeric(xIn) And Val(xIn) = Nb_OCDE + 1 Then
'        blnErr = False
'        Nb_OCDE = Nb_OCDE + 1
'        Line Input #3, xIn
'        xPays_OCDE = Trim(xIn)
'        'Line Input #3, xIn 'gb
'        Line Input #3, xLib 'fr
'        Line Input #3, xOld
'        Line Input #3, xNew
'
'        xPays_ISO = ""
'        xMsg = Space$(25) & " ; "
'        If xNew = "-" Then xNew = " "
'
'        blnOk = False
'        For K = 1 To arrOCDE_Nb
'            If xNew = arrOCDE(K).NOTPAYOCDE Then blnOk = True: Exit For
'        Next K
'        If Not blnOk Then
'            Nb_OCDE_X = Nb_OCDE_X + 1
'            Mid$(xMsg, 1, 25) = "notation OCDE inconnue"
'        Else
'            blnOk = False
'            For K = 1 To arrPays_Import_Nb
'                If xPays_OCDE = Mid$(arrPays_Import(K).BIATABTXT, 1, 3) Then
'                    xPays_ISO = Mid$(arrPays_Import(K).BIATABK2, 1, 2)
'                    blnOk = True
'                    Exit For
'                End If
'            Next K
'            If Not blnOk Then
'                Nb_OCDE_Err = Nb_OCDE_Err + 1
'                Mid$(xMsg, 1, 25) = "code pays OCDE inconnu " & xPays_OCDE
'            Else
'                blnOk = False
'                For K = 1 To arrYNOTPAY0_Nb
'                    If xPays_ISO = selYNOTPAY0(K).NOTPAYISO Then blnOk = True: Exit For
'                Next K
'                If Not blnOk Then
'                    Nb_OCDE_X = Nb_OCDE_X + 1
'                    Mid$(xMsg, 1, 25) = "code pays non géré BIA " ' & xPays_ISO
'                Else
'                    newYNOTPAY0 = selYNOTPAY0(K)
'                    newYNOTPAY0.NOTPAYOCDE = xNew
'                    newYNOTPAY0.NOTPAYOCDK = "A"
'                    newYNOTPAY0.NOTPAYOCDD = wAMJMin
'
'                    X = cmdSelect_Import_NOTPAYBIAN_Auto
'                    If X <> "" Then
'                        Nb_BIA_Err = Nb_BIA_Err + 1
'                        Mid$(xMsg, 4, 21) = X
'                    End If
'                    If newYNOTPAY0.NOTPAYOCDE = selYNOTPAY0(K).NOTPAYOCDE _
'                    And newYNOTPAY0.NOTPAYOCDK = selYNOTPAY0(K).NOTPAYOCDK _
'                    And newYNOTPAY0.NOTPAYOCDD = selYNOTPAY0(K).NOTPAYOCDD _
'                    And newYNOTPAY0.NOTPAYBIAN = selYNOTPAY0(K).NOTPAYBIAN Then
'                        Nb_BIA_E = Nb_BIA_E + 1
'                        Mid$(xMsg, 1, 1) = "="
'                        newYNOTPAY0.NOTPAYXUSR = "="
'                    Else
'                        Mid$(xMsg, 1, 1) = "#"
'                        Nb_BIA_D = Nb_BIA_D + 1
'                        newYNOTPAY0.NOTPAYXUSR = "#"
'                    End If
'
'                    selYNOTPAY0(K) = newYNOTPAY0
'                End If
'            End If
'        End If
'       Print #4, xMsg & xPays_ISO & " ; " & xNew & " ; " & xPays_OCDE & " ; " & xLib
'    End If
'Loop


Do Until EOF(3)
    ' Lire le numéro de ligne (si possible)
        If Not EOF(3) Then
            Line Input #3, xIn
        Else
            Exit Do
        End If
        
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "- " & xIn): DoEvents
    
    ' Ignorer les lignes vides et les lignes non pertinentes
    If IsNumeric(xIn) Then
    
        blnErr = False
        Nb_OCDE = Nb_OCDE + 1
        
        ' Lecture du numéro de ligne (nombre) pour chaque pays
        'ligneCompteur = CInt(xIn)
        
        ' Lire le code ISO Alpha-3
        If Not EOF(3) Then
            Line Input #3, xPays_OCDE
            xPays_OCDE = Trim(xPays_OCDE)
        Else
            Exit Do
        End If
        

        ' Lire le nom du pays
        If Not EOF(3) Then
            Line Input #3, xLib
            xLib = Trim(xLib)
        Else
            Exit Do
        End If

        ' Lire les notes (précédente et actuelle)
        If Not EOF(3) Then
            Line Input #3, xOld
            xOld = Trim(xOld)
        Else
            Exit Do
        End If
        
        If Not EOF(3) Then
            Line Input #3, xNew
            xNew = Trim(xNew)
        Else
            Exit Do
        End If

        
        xPays_ISO = ""
        xMsg = Space$(25) & " ; "
        If xNew = "-" Then xNew = " "
        
        blnOk = False
        For K = 1 To arrOCDE_Nb
            If xNew = arrOCDE(K).NOTPAYOCDE Then blnOk = True: Exit For
        Next K
        If Not blnOk Then
            Nb_OCDE_X = Nb_OCDE_X + 1
            Mid$(xMsg, 1, 25) = "notation OCDE inconnue"
        Else
            blnOk = False
            For K = 1 To arrPays_Import_Nb
                If xPays_OCDE = Mid$(arrPays_Import(K).BIATABTXT, 1, 3) Then
                    xPays_ISO = Mid$(arrPays_Import(K).BIATABK2, 1, 2)
                    blnOk = True
                    Exit For
                End If
            Next K
            If Not blnOk Then
                Nb_OCDE_Err = Nb_OCDE_Err + 1
                Mid$(xMsg, 1, 25) = "code pays OCDE inconnu " & xPays_OCDE
            Else
                blnOk = False
                For K = 1 To arrYNOTPAY0_Nb
                    If xPays_ISO = selYNOTPAY0(K).NOTPAYISO Then blnOk = True: Exit For
                Next K
                If Not blnOk Then
                    Nb_OCDE_X = Nb_OCDE_X + 1
                    Mid$(xMsg, 1, 25) = "code pays non géré BIA " ' & xPays_ISO
                Else
                    newYNOTPAY0 = selYNOTPAY0(K)
                    newYNOTPAY0.NOTPAYOCDE = xNew
                    newYNOTPAY0.NOTPAYOCDK = "A"
                    newYNOTPAY0.NOTPAYOCDD = wAMJMin
                    
                    X = cmdSelect_Import_NOTPAYBIAN_Auto
                    If X <> "" Then
                        Nb_BIA_Err = Nb_BIA_Err + 1
                        Mid$(xMsg, 4, 21) = X
                    End If
                    If newYNOTPAY0.NOTPAYOCDE = selYNOTPAY0(K).NOTPAYOCDE _
                    And newYNOTPAY0.NOTPAYOCDK = selYNOTPAY0(K).NOTPAYOCDK _
                    And newYNOTPAY0.NOTPAYOCDD = selYNOTPAY0(K).NOTPAYOCDD _
                    And newYNOTPAY0.NOTPAYBIAN = selYNOTPAY0(K).NOTPAYBIAN Then
                        Nb_BIA_E = Nb_BIA_E + 1
                        Mid$(xMsg, 1, 1) = "="
                        newYNOTPAY0.NOTPAYXUSR = "="
                    Else
                        Mid$(xMsg, 1, 1) = "#"
                        Nb_BIA_D = Nb_BIA_D + 1
                        newYNOTPAY0.NOTPAYXUSR = "#"
                    End If
                    
                    selYNOTPAY0(K) = newYNOTPAY0
                End If
            End If
        End If
       Print #4, xMsg & xPays_ISO & " ; " & xNew & " ; " & xPays_OCDE & " ; " & xLib
    End If
Loop



'________________________________________ FIN MODIFICATION KOKOU __________________________________________________



Call lstErr_AddItem(lstErr, cmdContext, "- pays BIA non renseigné OCDE"): DoEvents

Print #4, ""
Print #4, "-----------------------------------------------------------------"
For K = 1 To arrYNOTPAY0_Nb
    If Trim(selYNOTPAY0(K).NOTPAYXUSR) = "" Then
        For K2 = 1 To arrPays_NB
            If selYNOTPAY0(K).NOTPAYISO = arrPAYS_ISO(K2) Then Exit For
        Next K2

        Nb_BIA_X = Nb_BIA_X + 1
        Print #4, "pays BIA non renseigné OCDE" & "; " & selYNOTPAY0(K).NOTPAYISO & " ; ; ; " & arrPays_Lib(K2)
        Call lstErr_AddItem(lstErr, cmdContext, "! " & selYNOTPAY0(K).NOTPAYISO): DoEvents
    End If
Next K

Print #4, ""
Print #4, "-----------------------------------------------------------------"
Print #4, "Nb OCDE lus          ; " & Nb_OCDE
Print #4, "Nb OCDE inconnus     ; " & Nb_OCDE_Err
Print #4, "Nb OCDE ignorés      ; " & Nb_OCDE_X
Print #4, ""
Print #4, "Nb BIA identiques    ; " & Nb_BIA_E
Print #4, "Nb BIA mis à jour    ; " & Nb_BIA_D
Print #4, "Nb BIA sans notation ;" & Nb_BIA_X
Print #4, "Nb BIA traités/gérés ; " & Nb_BIA_E + Nb_BIA_D + Nb_BIA_X & " / " & arrYNOTPAY0_Nb
Print #4, ""
Print #4, "Nb BIA en anomalie   ; " & Nb_BIA_Err
Print #4, "================================================================="

Close #3
Close #4
Call FEU_VERT

Call lstErr_AddItem(lstErr, cmdContext, "- mise à jour du fichier YNOTPAY0"): DoEvents

Call cmdYNOTPAY0_Update("Import")
Call cmdSelect_Ok_Click

Call cmdSelect_Import_Log(wFile_Log)

cmdSelect_Import_DateNotation_Update ("OCDE")


GoTo Exit_sub

Error_Handler:
    MsgBox Error, vbCritical, "Import OCDE"
    
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0
Close

End Sub

Private Sub cmdSelect_Import_SP_Click()
Dim wFile As String, wFilex As String, wFile_Log As String
Dim xIn As String, xPays_ISO As String, xPays_SP As String, xLib As String, xMsg As String
Dim X As String, Xcom As String, xNew As String
Dim Nb_SP As Integer, Nb_SP_X As Integer, Nb_SP_Err As Integer
Dim Nb_BIA_X As Integer, Nb_BIA_E As Integer, Nb_BIA_D As Integer, Nb_BIA_Err As Integer
Dim K As Integer, blnSP As Boolean, blnOk As Boolean, blnErr As Boolean
Dim K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer
Dim arrAlphabet As String, kAlphabet As Integer, mAlphabet As String
Dim blnEntity As Boolean, xIn_U As String
On Error GoTo Error_Handler


'$JPL 20101206 nouvelles pages internet

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Import_SP en cours ......"): DoEvents
arrAlphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
kAlphabet = 1

blnSP = False
Nb_SP = 0: Nb_SP_X = 0: Nb_SP_Err = 0
Nb_BIA_X = 0: Nb_BIA_E = 0: Nb_BIA_D = 0: Nb_BIA_Err = 0

Call DTPicker_Control(txtSelect_Import_Amj, wAMJMin)
'______________________________________________
If Not cmdSelect_Import_Control("SP") Then GoTo Exit_sub

'______________________________________________
wFile = Trim(SP_Notepad.BIATABTXT)
X = InputBox("par défaut : " & wFile & vbCrLf & vbCrLf _
    & Coface_DateNotation_Info & vbCrLf & vbCrLf _
    & "     ================================" _
   & vbCrLf & "     EN DATE du : " & dateImp10(wAMJMin) _
   & vbCrLf & "     ================================", "Notation Pays : nom du fichier d'import S & P", wFile)
If Trim(X) = "" Then GoTo Exit_sub

wFilex = Trim(X)
If Dir(wFilex) = "" Then
    Call MsgBox("Le fichier : " & wFile & "n'existe pas", vbCritical, "import sp")
    GoTo Exit_sub
End If
'______________________________________________
If wFile <> wFilex Then
    wFile = wFilex
    Old_YBIATAB0 = SP_Notepad
    New_YBIATAB0 = SP_Notepad
    New_YBIATAB0.BIATABTXT = wFilex
    Parametrage_Update
    SP_Notepad = New_YBIATAB0
End If
'_________________________________________

ReDim selYNOTPAY0(arrYNOTPAY0_Nb)
For K = 1 To arrYNOTPAY0_Nb
    selYNOTPAY0(K) = arrYNOTPAY0(K)
    selYNOTPAY0(K).NOTPAYXAMJ = 0
    selYNOTPAY0(K).NOTPAYXHMS = 0
    selYNOTPAY0(K).NOTPAYXUSR = ""
Next K

wFile_Log = wFile & "_" & DSYS_Time & ".log"
Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile): DoEvents
Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile_Log): DoEvents

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Import S&P"
newYNOTPAYLOG.NOTPAYLOGX = "log S & P " & wFile_Log
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

Open wFile For Input As #3
Call FEU_ROUGE
Open wFile_Log For Output As #4
Print #4, "Import du fichier S & P en date du : " & dateImp10(wAMJMin) & "  " & wFile
Print #4, "================================================================="
Print #4, "fichier log : " & wFile_Log
Print #4, "-----------------------------------------------------------------"
Print #4, ""


blnOk = False
blnEntity = False

Do Until EOF(3)
    Line Input #3, xIn
    xIn = Trim(xIn)
    If xIn = "G" Or xIn = "Czech Republic" Or xIn = "Gabonese Republic" Then
        Debug.Print xIn
    End If
    
    xIn_U = UCase$(xIn)
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "- " & xIn): DoEvents
'=============================================================

    'If InStr(xIn, "PreviousStandard & Poor's Views") > 0 Then blnEntity = False
    K4 = Len(xIn)
    If Not blnEntity Or K4 = 1 Then
        If Len(xIn) = 1 Then
            For K = kAlphabet To 36
                If xIn_U = Mid$(arrAlphabet, K, 1) Then
                    kAlphabet = K
                    mAlphabet = Mid$(arrAlphabet, K, 1)
                    blnEntity = True
                    Exit For
                End If
            Next K
        End If
                    
    Else
    
        If K4 > 0 Then
            If Mid$(xIn_U, 1, 1) <> mAlphabet Then
                blnEntity = False
            Else
            '____________________________________________________________
                If K4 > 10 Then
                    Nb_SP = Nb_SP + 1
        
                
                    For K3 = K4 To 1 Step -1
                        If Mid$(xIn, K3, 1) = " " Then Exit For
                    Next K3
            
                    For K2 = K3 - 1 To 1 Step -1
                        If Mid$(xIn, K2, 1) = " " Then xNew = Trim(Mid$(xIn, K2 + 1, K3 - K2 - 1)): Exit For
                    Next K2
                    For K1 = K2 - 1 To 1 Step -1
                        If Mid$(xIn, K1, 1) = " " Then Exit For
                    Next K1
                    
                    K = InStr(1, xIn, "(")
                    If K > 0 And K < K1 Then K1 = K - 1
                    If K1 > 32 Then K1 = 32
                    '=========================
                    xPays_SP = LCase$(Trim(Mid$(xIn, 1, K1)))
'$JPL 20111006 _______________________________________________________________________
                    
                    xPays_SP = Replace(xPays_SP, "sultanate of", "")
                    xPays_SP = Replace(xPays_SP, "state of the", "")
                    xPays_SP = Replace(xPays_SP, "state of", "")
                    xPays_SP = Replace(xPays_SP, "the republic of", "")
                    xPays_SP = Replace(xPays_SP, "republic of", "")
                    xPays_SP = Replace(xPays_SP, "republic", "")
                    xPays_SP = Replace(xPays_SP, "grand duchy of", "")
                    xPays_SP = Replace(xPays_SP, "emirate of", "")
                    xPays_SP = Replace(xPays_SP, "government of", "")
                    xPays_SP = Replace(xPays_SP, "kingdom of", "")
                    xPays_SP = Replace(xPays_SP, "principality of", "")

                    xPays_SP = Replace(xPays_SP, "hashemite", "")
                    xPays_SP = Replace(xPays_SP, "oriental", "")
                    xPays_SP = Replace(xPays_SP, "plurinational", "")
                    'xPays_SP = Replace(xPays_SP, "federation", "")
                    xPays_SP = Replace(xPays_SP, "federative", "")
                    xPays_SP = Replace(xPays_SP, "federal", "")
                    
                    xPays_SP = Trim(xPays_SP)
                   
'$JPL 20111006 _______________________________________________________________________
                    xPays_ISO = ""
                    xMsg = Space$(25) & " ; "
                    
                    blnOk = False
                    For K = 1 To arrSP_Nb
                        If xNew = Trim(arrSP(K).NOTPAYSP) Then blnOk = True: Exit For
                    Next K
                    If Not blnOk Then
                        Nb_SP_X = Nb_SP_X + 1
                        Mid$(xMsg, 1, 25) = "notation S & P inconnue"
                    Else
                        blnOk = False
                        For K = 1 To arrPays_Import_Nb
                            If xPays_SP = Trim(Mid$(arrPays_Import(K).BIATABTXT, 68, 32)) Then
                                xPays_ISO = Mid$(arrPays_Import(K).BIATABK2, 1, 2)
                                blnOk = True
                                Exit For
                            End If
                        Next K
                        If Not blnOk Then
                            Nb_SP_Err = Nb_SP_Err + 1
                            Mid$(xMsg, 1, 25) = "code pays S & P inconnu " & xPays_SP
                        Else
                            blnOk = False
                            For K = 1 To arrYNOTPAY0_Nb
                                If xPays_ISO = selYNOTPAY0(K).NOTPAYISO Then blnOk = True: Exit For
                            Next K
                            If Not blnOk Then
                                Nb_SP_X = Nb_SP_X + 1
                                Mid$(xMsg, 1, 25) = "code pays non géré BIA " ' & xPays_ISO
                            Else
                                newYNOTPAY0 = selYNOTPAY0(K)
                                newYNOTPAY0.NOTPAYSP = xNew
                                newYNOTPAY0.NOTPAYSPK = "A"
                                newYNOTPAY0.NOTPAYSPD = wAMJMin
                                
                                X = cmdSelect_Import_NOTPAYBIAN_Auto
                                If X <> "" Then
                                    Nb_BIA_Err = Nb_BIA_Err + 1
                                    Mid$(xMsg, 4, 21) = X
                                End If
                                If newYNOTPAY0.NOTPAYSP = selYNOTPAY0(K).NOTPAYSP _
                                And newYNOTPAY0.NOTPAYSPK = selYNOTPAY0(K).NOTPAYSPK _
                                And newYNOTPAY0.NOTPAYSPD = selYNOTPAY0(K).NOTPAYSPD _
                                And newYNOTPAY0.NOTPAYBIAN = selYNOTPAY0(K).NOTPAYBIAN Then
                                    Nb_BIA_E = Nb_BIA_E + 1
                                    Mid$(xMsg, 1, 1) = "="
                                    newYNOTPAY0.NOTPAYXUSR = "="
                                Else
                                    Mid$(xMsg, 1, 1) = "#"
                                    Nb_BIA_D = Nb_BIA_D + 1
                                    newYNOTPAY0.NOTPAYXUSR = "#"
                                End If
                                
                                selYNOTPAY0(K) = newYNOTPAY0
                            End If
                        End If
                        Print #4, xMsg & xPays_ISO & " ; " & xNew & " ; " & xPays_SP & " ; " & xLib
                   End If
                End If
            '____________________________________________________________
            End If
        End If
    End If
    
'=============================================================
    'If InStr(xIn, "Entity Domestic Rating Foreign Rating") > 0 Then blnEntity = True
'=============================================================
Loop
Call lstErr_AddItem(lstErr, cmdContext, "- pays BIA non renseigné S & P"): DoEvents

Print #4, ""
Print #4, "-----------------------------------------------------------------"
For K = 1 To arrYNOTPAY0_Nb
    If Trim(selYNOTPAY0(K).NOTPAYXUSR) = "" Then
        For K2 = 1 To arrPays_NB
            If selYNOTPAY0(K).NOTPAYISO = arrPAYS_ISO(K2) Then Exit For
        Next K2

        Nb_BIA_X = Nb_BIA_X + 1
        Print #4, "pays BIA non renseigné S & P" & "; " & selYNOTPAY0(K).NOTPAYISO & " ; ; " & arrPays_Lib(K2)
        Call lstErr_AddItem(lstErr, cmdContext, "! " & selYNOTPAY0(K).NOTPAYISO): DoEvents
    End If
Next K

Print #4, ""
Print #4, "-----------------------------------------------------------------"
Print #4, "Nb S & P lus          ; " & Nb_SP
Print #4, "Nb S & P inconnus     ; " & Nb_SP_Err
Print #4, "Nb S & P ignorés      ; " & Nb_SP_X
Print #4, ""
Print #4, "Nb BIA identiques     ; " & Nb_BIA_E
Print #4, "Nb BIA mis à jour     ; " & Nb_BIA_D
Print #4, "Nb BIA sans notation  ;" & Nb_BIA_X
Print #4, "Nb BIA traités/gérés  ; " & Nb_BIA_E + Nb_BIA_D + Nb_BIA_X & " / " & arrYNOTPAY0_Nb
Print #4, ""
Print #4, "Nb BIA en anomalie    ; " & Nb_BIA_Err
Print #4, "================================================================="

Close #3
Close #4
Call FEU_VERT

Call lstErr_AddItem(lstErr, cmdContext, "- mise à jour du fichier YNOTPAY0"): DoEvents

Call cmdYNOTPAY0_Update("Import")
Call cmdSelect_Ok_Click

Call cmdSelect_Import_Log(wFile_Log)

cmdSelect_Import_DateNotation_Update ("SP")


GoTo Exit_sub

Error_Handler:
    MsgBox xIn & vbCrLf & Error, vbCritical, "Import S & P"
    
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0
Close

End Sub

Private Sub cmdSelect_Import_SP_Click_2009()
Dim wFile As String, wFilex As String, wFile_Log As String
Dim xIn As String, xPays_ISO As String, xPays_SP As String, xLib As String, xMsg As String
Dim X As String, Xcom As String, xNew As String
Dim Nb_SP As Integer, Nb_SP_X As Integer, Nb_SP_Err As Integer
Dim Nb_BIA_X As Integer, Nb_BIA_E As Integer, Nb_BIA_D As Integer, Nb_BIA_Err As Integer
Dim K As Integer, blnSP As Boolean, blnOk As Boolean, blnErr As Boolean
Dim K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer
On Error GoTo Error_Handler

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Import_SP en cours ......"): DoEvents

blnSP = False
Nb_SP = 0: Nb_SP_X = 0: Nb_SP_Err = 0
Nb_BIA_X = 0: Nb_BIA_E = 0: Nb_BIA_D = 0: Nb_BIA_Err = 0

Call DTPicker_Control(txtSelect_Import_Amj, wAMJMin)
'______________________________________________
If Not cmdSelect_Import_Control("SP") Then GoTo Exit_sub

'______________________________________________
wFile = Trim(SP_Notepad.BIATABTXT)
X = InputBox("par défaut : " & wFile & vbCrLf & vbCrLf _
    & Coface_DateNotation_Info & vbCrLf & vbCrLf _
    & "     ================================" _
   & vbCrLf & "     EN DATE du : " & dateImp10(wAMJMin) _
   & vbCrLf & "     ================================", "Notation Pays : nom du fichier d'import S & P", wFile)
If Trim(X) = "" Then GoTo Exit_sub

wFilex = Trim(X)
If Dir(wFilex) = "" Then
    Call MsgBox("Le fichier : " & wFile & "n'existe pas", vbCritical, "import sp")
    GoTo Exit_sub
End If
'______________________________________________
If wFile <> wFilex Then
    wFile = wFilex
    Old_YBIATAB0 = SP_Notepad
    New_YBIATAB0 = SP_Notepad
    New_YBIATAB0.BIATABTXT = wFilex
    Parametrage_Update
    SP_Notepad = New_YBIATAB0
End If
'_________________________________________

ReDim selYNOTPAY0(arrYNOTPAY0_Nb)
For K = 1 To arrYNOTPAY0_Nb
    selYNOTPAY0(K) = arrYNOTPAY0(K)
    selYNOTPAY0(K).NOTPAYXAMJ = 0
    selYNOTPAY0(K).NOTPAYXHMS = 0
    selYNOTPAY0(K).NOTPAYXUSR = ""
Next K

wFile_Log = wFile & "_" & DSYS_Time & ".log"
Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile): DoEvents
Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile_Log): DoEvents

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Import S&P"
newYNOTPAYLOG.NOTPAYLOGX = "log S & P " & wFile_Log
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

Open wFile For Input As #3
Call FEU_ROUGE
Open wFile_Log For Output As #4
Print #4, "Import du fichier S & P en date du : " & dateImp10(wAMJMin) & "  " & wFile
Print #4, "================================================================="
Print #4, "fichier log : " & wFile_Log
Print #4, "-----------------------------------------------------------------"
Print #4, ""


blnOk = False

Do Until EOF(3)
    Line Input #3, xIn
    xIn = Trim(xIn)
    Call lstErr_ChangeLastItem(lstErr, cmdContext, "- " & xIn): DoEvents
        
    K3 = InStr(1, xIn, "/")
    K4 = InStr(K3 + 1, xIn, "/")
    K4 = InStr(K4 + 1, xIn, "/")

    If K3 > 0 Then
        Nb_SP = Nb_SP + 1
        K1 = InStr(1, xIn, " ")
        K2 = InStr(1, xIn, "(")
        If K2 <= 0 Then
            K = InStr(K1 + 1, xIn, " ")
            If K > 0 And K < K3 Then
                K1 = K
                K = InStr(K1 + 1, xIn, " ")
                If K > 0 And K < K3 Then
                    K1 = K
                    K = InStr(K1 + 1, xIn, " ")
                    If K > 0 And K < K3 Then
                        K1 = K
                    End If
                End If
            End If
               
            X = LCase(Trim(Mid$(xIn, 1, K1 - 1)))
        Else
            X = LCase(Trim(Mid$(xIn, 1, K2 - 2)))
        End If
        
        '=========================
        For K2 = K4 To 1 Step -1
            If Mid$(xIn, K2, 1) = " " Then xNew = Trim(Mid$(xIn, K2 + 1, K4 - K2 - 1)): Exit For
        Next K2
        '=========================
        xPays_SP = LCase$(Trim(Mid$(xIn, 1, K1)))

        xPays_ISO = ""
        xMsg = Space$(25) & " ; "
        
        blnOk = False
        For K = 1 To arrSP_Nb
            If xNew = Trim(arrSP(K).NOTPAYSP) Then blnOk = True: Exit For
        Next K
        If Not blnOk Then
            Nb_SP_X = Nb_SP_X + 1
            Mid$(xMsg, 1, 25) = "notation S & P inconnue"
        Else
            blnOk = False
            For K = 1 To arrPays_Import_Nb
                If xPays_SP = Trim(Mid$(arrPays_Import(K).BIATABTXT, 68, 32)) Then
                    xPays_ISO = Mid$(arrPays_Import(K).BIATABK2, 1, 2)
                    blnOk = True
                    Exit For
                End If
            Next K
            If Not blnOk Then
                Nb_SP_Err = Nb_SP_Err + 1
                Mid$(xMsg, 1, 25) = "code pays S & P inconnu " & xPays_SP
            Else
                blnOk = False
                For K = 1 To arrYNOTPAY0_Nb
                    If xPays_ISO = selYNOTPAY0(K).NOTPAYISO Then blnOk = True: Exit For
                Next K
                If Not blnOk Then
                    Nb_SP_X = Nb_SP_X + 1
                    Mid$(xMsg, 1, 25) = "code pays non géré BIA " ' & xPays_ISO
                Else
                    newYNOTPAY0 = selYNOTPAY0(K)
                    newYNOTPAY0.NOTPAYSP = xNew
                    newYNOTPAY0.NOTPAYSPK = "A"
                    newYNOTPAY0.NOTPAYSPD = wAMJMin
                    
                    X = cmdSelect_Import_NOTPAYBIAN_Auto
                    If X <> "" Then
                        Nb_BIA_Err = Nb_BIA_Err + 1
                        Mid$(xMsg, 4, 21) = X
                    End If
                    If newYNOTPAY0.NOTPAYSP = selYNOTPAY0(K).NOTPAYSP _
                    And newYNOTPAY0.NOTPAYSPK = selYNOTPAY0(K).NOTPAYSPK _
                    And newYNOTPAY0.NOTPAYSPD = selYNOTPAY0(K).NOTPAYSPD _
                    And newYNOTPAY0.NOTPAYBIAN = selYNOTPAY0(K).NOTPAYBIAN Then
                        Nb_BIA_E = Nb_BIA_E + 1
                        Mid$(xMsg, 1, 1) = "="
                        newYNOTPAY0.NOTPAYXUSR = "="
                    Else
                        Mid$(xMsg, 1, 1) = "#"
                        Nb_BIA_D = Nb_BIA_D + 1
                        newYNOTPAY0.NOTPAYXUSR = "#"
                    End If
                    
                    selYNOTPAY0(K) = newYNOTPAY0
                End If
            End If
            Print #4, xMsg & xPays_ISO & " ; " & xNew & " ; " & xPays_SP & " ; " & xLib
           End If
        End If
    'End If
Loop
Call lstErr_AddItem(lstErr, cmdContext, "- pays BIA non renseigné S & P"): DoEvents

Print #4, ""
Print #4, "-----------------------------------------------------------------"
For K = 1 To arrYNOTPAY0_Nb
    If Trim(selYNOTPAY0(K).NOTPAYXUSR) = "" Then
        For K2 = 1 To arrPays_NB
            If selYNOTPAY0(K).NOTPAYISO = arrPAYS_ISO(K2) Then Exit For
        Next K2

        Nb_BIA_X = Nb_BIA_X + 1
        Print #4, "pays BIA non renseigné S & P" & "; " & selYNOTPAY0(K).NOTPAYISO & " ; ; " & arrPays_Lib(K2)
        Call lstErr_AddItem(lstErr, cmdContext, "! " & selYNOTPAY0(K).NOTPAYISO): DoEvents
    End If
Next K

Print #4, ""
Print #4, "-----------------------------------------------------------------"
Print #4, "Nb S & P lus          ; " & Nb_SP
Print #4, "Nb S & P inconnus     ; " & Nb_SP_Err
Print #4, "Nb S & P ignorés      ; " & Nb_SP_X
Print #4, ""
Print #4, "Nb BIA identiques     ; " & Nb_BIA_E
Print #4, "Nb BIA mis à jour     ; " & Nb_BIA_D
Print #4, "Nb BIA sans notation  ;" & Nb_BIA_X
Print #4, "Nb BIA traités/gérés  ; " & Nb_BIA_E + Nb_BIA_D + Nb_BIA_X & " / " & arrYNOTPAY0_Nb
Print #4, ""
Print #4, "Nb BIA en anomalie    ; " & Nb_BIA_Err
Print #4, "================================================================="

Close #3
Close #4
Call FEU_VERT

Call lstErr_AddItem(lstErr, cmdContext, "- mise à jour du fichier YNOTPAY0"): DoEvents

Call cmdYNOTPAY0_Update("Import")
Call cmdSelect_Ok_Click

Call cmdSelect_Import_Log(wFile_Log)

cmdSelect_Import_DateNotation_Update ("SP")


GoTo Exit_sub

Error_Handler:
    MsgBox xIn & vbCrLf & Error, vbCritical, "Import S & P"
    
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0
Close

End Sub


Private Sub cmdSelect_Import_Validation_Click()
Dim V, xSQL As String
Dim mSEQ As Long
Dim K As Integer
Dim blnCompare_0 As Boolean, blnCompare_HAMJ As Boolean
Dim mNOTPAYHAMJ As String
Dim wFile As String, wFile_Log As String, wMsg As String
Dim Nb_Update As Long, Nb_NOTPAYBIAN As Long
Dim kPays As Integer
Dim X40 As String

Call DTPicker_Control(txtSelect_Import_Amj, mNOTPAYHAMJ)
wAMJMin = mNOTPAYHAMJ
If Not cmdSelect_Import_Control("BIA") Then GoTo Exit_Sub2

X = MsgBox("Confirmez-vous la validation des notations" _
    & vbCrLf & vbCrLf & "     =========================" _
   & vbCrLf & "           EN DATE du : " & dateImp10(mNOTPAYHAMJ) _
   & vbCrLf & "     =========================", vbYesNo, "Notation Pays : VALIDATION")
If X <> vbYes Then GoTo Exit_Sub2

Me.Enabled = False: Me.MousePointer = vbHourglass
'==========================================================================
App_Debug = "Import_Validation Début"

wFile = Trim(Coface_Notepad.BIATABTXT)

For K = Len(wFile) To 1 Step -1
    If Mid$(wFile, K, 1) = "\" Then Exit For
Next K
wFile_Log = Mid$(wFile, 1, K) & "Import_Validation_" & DSYS_Time & ".log"

Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile): DoEvents
Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile_Log): DoEvents

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Import Val"
newYNOTPAYLOG.NOTPAYLOGX = "log VALIDATION : " & wFile_Log
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++
Call FEU_ROUGE

Open wFile_Log For Output As #4
Print #4, "Notation Pays : VALIDATION Import en date du : " & dateImp10(mNOTPAYHAMJ)
Print #4, "================================================================="
Print #4, "fichier log : " & wFile_Log
Print #4, "-----------------------------------------------------------------"
Print #4, ""


Nb_Update = 0: Nb_NOTPAYBIAN = 0

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Call FEU_ROUGE
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
For K = 1 To arrYNOTPAY0_Nb
    V = Null
    newYNOTPAY0 = arrYNOTPAY0(K): newYNOTPAY0.NOTPAYSEQ = 0
    '''oldYNOTPAY0 = newYNOTPAY0: oldYNOTPAY0.NOTPAYSEQ = 0
    
    App_Debug = "Import_Validation lecture SEQ = 0"
    
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YNOTPAY0 " _
     & " where NOTPAYISO = '" & newYNOTPAY0.NOTPAYISO & "' and NOTPAYSEQ = 0"
    Set rsSab = cnsab.Execute(xSQL)
    
    If rsSab.EOF Then
        V = sqlYNOTPAY0_Insert(newYNOTPAY0)
    Else
        V = rsYNOTPAY0_GetBuffer(rsSab, oldYNOTPAY0)
        blnCompare_0 = sqlYNOTPAY0_Compare(newYNOTPAY0, oldYNOTPAY0)
        If Not blnCompare_0 Then
            V = sqlYNOTPAY0_Update(newYNOTPAY0, oldYNOTPAY0)
            
            Nb_Update = Nb_Update + 1
            For kPays = 1 To arrPays_NB
                If newYNOTPAY0.NOTPAYISO = arrPAYS_ISO(kPays) Then Exit For
            Next kPays
            X40 = newYNOTPAY0.NOTPAYISO & "; " & arrPays_Lib(kPays)
            X40 = X40 & Space$(40 - Len(X40))
            If oldYNOTPAY0.NOTPAYCOFA <> newYNOTPAY0.NOTPAYCOFA Then Print #4, X40 & ";Coface : " & oldYNOTPAY0.NOTPAYCOFA & " => " & newYNOTPAY0.NOTPAYCOFA
            If oldYNOTPAY0.NOTPAYCOF2 <> newYNOTPAY0.NOTPAYCOF2 Then Print #4, X40 & ";Affaires : " & oldYNOTPAY0.NOTPAYCOF2 & " => " & newYNOTPAY0.NOTPAYCOF2
            If oldYNOTPAY0.NOTPAYOCDE <> newYNOTPAY0.NOTPAYOCDE Then Print #4, X40 & ";OCDE : " & oldYNOTPAY0.NOTPAYOCDE & " => " & newYNOTPAY0.NOTPAYOCDE
            If oldYNOTPAY0.NOTPAYSP <> newYNOTPAY0.NOTPAYSP Then Print #4, X40 & ";S & P : " & oldYNOTPAY0.NOTPAYSP & " => " & newYNOTPAY0.NOTPAYSP
            If oldYNOTPAY0.NOTPAYCEG <> newYNOTPAY0.NOTPAYCEG Then Print #4, X40 & ";Critère événement grave : " & oldYNOTPAY0.NOTPAYCEG & " => " & newYNOTPAY0.NOTPAYCEG
            '
            If oldYNOTPAY0.NOTPAYBIAN <> newYNOTPAY0.NOTPAYBIAN Then
                Nb_NOTPAYBIAN = Nb_NOTPAYBIAN + 1
'                Print #4, newYNOTPAY0.NOTPAYISO & "; " & arrPays_Lib(kPays) & ";Note BIA : " & oldYNOTPAY0.NOTPAYBIAN & " => " & newYNOTPAY0.NOTPAYBIAN & " ;Taux BIA : " & oldYNOTPAY0.NOTPAYTAUX & " => " & newYNOTPAY0.NOTPAYTAUX
                Print #4, X40 & ";BIA " & oldYNOTPAY0.NOTPAYBIAN & " = " & Format$(oldYNOTPAY0.NOTPAYTAUX, "##0.00") & " %" & " ; " & newYNOTPAY0.NOTPAYBIAN & " = " & Format$(newYNOTPAY0.NOTPAYTAUX, "##0.00") & " %"
            End If
        End If
    End If
    If Not IsNull(V) Then GoTo Error_MsgBox
'_______________________________________________________________________
    App_Debug = "Import_Validation recherche SEQ"

    xSQL = "select NOTPAYISO , NOTPAYSEQ from " & paramIBM_Library_SABSPE & ".YNOTPAY0 " _
         & " where NOTPAYISO = '" & newYNOTPAY0.NOTPAYISO & "' order by NOTPAYSEQ  desc"
    Set rsSab = cnsab.Execute(xSQL)
    
    If rsSab.EOF Then
        mSEQ = 0
    Else
        mSEQ = rsSab("NOTPAYSEQ")
    End If
'_______________________________________________________________________
    App_Debug = "Import_Validation Archivage en cours"
    If Not blnCompare_0 Then
        mSEQ = mSEQ + 1
        oldYNOTPAY0.NOTPAYSEQ = mSEQ
        V = sqlYNOTPAY0_Insert(oldYNOTPAY0)
    End If
'_______________________________________________________________________
    
    App_Debug = "Import_Validation création arrêté"

    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YNOTPAY0 " _
     & " where NOTPAYISO = '" & newYNOTPAY0.NOTPAYISO & "' and NOTPAYHAMJ = " & mNOTPAYHAMJ
    Set rsSab = cnsab.Execute(xSQL)
    
    If rsSab.EOF Then
        newYNOTPAY0.NOTPAYHAMJ = mNOTPAYHAMJ
        newYNOTPAY0.NOTPAYSEQ = mSEQ + 1
        V = sqlYNOTPAY0_Insert(newYNOTPAY0)
    Else
        V = rsYNOTPAY0_GetBuffer(rsSab, oldYNOTPAY0)
        newYNOTPAY0.NOTPAYSEQ = oldYNOTPAY0.NOTPAYSEQ
        newYNOTPAY0.NOTPAYHAMJ = mNOTPAYHAMJ
        ''blnCompare_HAMJ = sqlYNOTPAY0_Compare(newYNOTPAY0, oldYNOTPAY0)
        ''If Not blnCompare_HAMJ Then
        V = sqlYNOTPAY0_Update(newYNOTPAY0, oldYNOTPAY0)
    End If
    If Not IsNull(V) Then GoTo Error_MsgBox

    
'________________________________________________________________________________

    
Next K

V = sqlYNOTPAY0_Delete_Where(" Where NOTPAYSEQ = -1")
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, Me.Name & " ~ " & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
        Call FEU_VERT

    Else
        V = cnSAB_Transaction("Commit")
        'cboSelect_SQL.ListIndex = 0
        fgSelect.Visible = False
        fraSelect_Import.Visible = False
        Call FEU_VERT
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Import VAL"
newYNOTPAYLOG.NOTPAYLOGX = "au " & dateImp10(mNOTPAYHAMJ) & " Nb BIA mod / nb màj / nb lus : " & Nb_NOTPAYBIAN & " / " & Nb_Update & " / " & arrYNOTPAY0_Nb & " "
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++
Print #4, ""
Print #4, "-----------------------------------------------------------------"
Print #4, "Nb lus          ; " & arrYNOTPAY0_Nb
Print #4, "Nb mises à jour   ; " & Nb_Update
Print #4, "Nb Notes BIA modifiées   ; " & Nb_NOTPAYBIAN
Print #4, "================================================================="

Close #4
Call FEU_VERT
Call cmdSelect_Import_Log(wFile_Log)


Exit_Sub2:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Internet_Coface_Click()
Dim wFile As String, wFilex As String, X As String
Dim xSQL As String
'______________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YNOTPAY0' and BIATABK1 = 'Internet' and BIATABK2 = 'Coface'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    V = rsYBIATAB0_GetBuffer(rsSab, Coface_Internet)
Else
    MsgBox xSQL, vbCritical, "Internet Coface : paramétrage inconsistant"
    
    Exit Sub
End If
wFile = Trim(Coface_Internet.BIATABTXT)
'______________________________________________

X = InputBox("par défaut : " & wFile & vbCrLf & vbCrLf _
   & "-click droit => AFFICHER la SOURCE" & vbCrLf _
   & "-Ctrl + A (tout sélectionner)" & vbCrLf _
   & "-Ctrl + C (tout copier)" & vbCrLf _
   & "-Ouvrir NOTEPAD" & vbCrLf _
   & "-Ctrl + V (tout coller)" & vbCrLf & vbCrLf _
   & "-Modifier le lien : " & Trim(Coface_Notepad.BIATABTXT), "Internet_Coface : nom du lien", wFile)
If Trim(X) = "" Then Exit Sub

wFilex = Trim(X)
If wFile <> wFilex Then
    Old_YBIATAB0 = Coface_Internet
    New_YBIATAB0 = Coface_Internet
    New_YBIATAB0.BIATABTXT = wFilex
    Parametrage_Update
    Coface_Internet = New_YBIATAB0
End If
'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Internet C"
newYNOTPAYLOG.NOTPAYLOGX = "Lien Coface : " & wFilex
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

Call frmElpPrt.IExplore(wFilex)

End Sub

Private Sub cmdSelect_Internet_OCDE_Click()

Dim wFile As String, wFilex As String, X As String

Dim xSQL As String
'______________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YNOTPAY0' and BIATABK1 = 'Internet' and BIATABK2 = 'OCDE'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    V = rsYBIATAB0_GetBuffer(rsSab, OCDE_Internet)
Else
    MsgBox xSQL, vbCritical, "Internet OCDE : paramétrage inconsistant"
    
    Exit Sub
End If
wFile = Trim(OCDE_Internet.BIATABTXT)
'______________________________________________

X = InputBox("par défaut : " & wFile & vbCrLf & vbCrLf _
   & "-Ctrl + A (tout sélectionner)" & vbCrLf _
   & "-Ctrl + C (tout copier)" & vbCrLf _
   & "-Ouvrir NOTEPAD" & vbCrLf _
   & "-Ctrl + V (tout coller)" & vbCrLf & vbCrLf _
   & "-Modifier le lien : " & Trim(OCDE_Notepad.BIATABTXT) & vbCrLf, "Internet_OCDE: nom du lien", wFile)


If Trim(X) = "" Then Exit Sub

wFilex = Trim(X)
If wFile <> wFilex Then
    Old_YBIATAB0 = OCDE_Internet
    New_YBIATAB0 = OCDE_Internet
    New_YBIATAB0.BIATABTXT = wFilex
    Parametrage_Update
    OCDE_Internet = New_YBIATAB0
End If

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Internet O"
newYNOTPAYLOG.NOTPAYLOGX = "lien OCDE: " & wFilex
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

Call frmElpPrt.IExplore(wFilex)


End Sub

Private Sub cmdSelect_Internet_SP_Click()
Dim wFile As String, wFilex As String, X As String
Dim xSQL As String
'______________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YNOTPAY0' and BIATABK1 = 'Internet' and BIATABK2 = 'SP'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    V = rsYBIATAB0_GetBuffer(rsSab, SP_Internet)
Else
    MsgBox xSQL, vbCritical, "Internet SP : paramétrage inconsistant"
    
    Exit Sub
End If
wFile = Trim(SP_Internet.BIATABTXT)
'______________________________________________

X = InputBox("par défaut : " & wFile & vbCrLf & vbCrLf _
   & "-Ctrl + A (tout sélectionner)" & vbCrLf _
   & "-Ctrl + C (tout copier)" & vbCrLf _
   & "-Ouvrir NOTEPAD" & vbCrLf _
   & "-Ctrl + V (tout coller)" & vbCrLf & vbCrLf _
   & "============================" & vbCrLf _
   & "-ATTENTION : plusieurs pages" & vbCrLf _
   & "============================" & vbCrLf & vbCrLf _
   & "-Modifier le lien : " & Trim(SP_Notepad.BIATABTXT) & vbCrLf, "Internet_S&P : nom du lien", wFile)
If Trim(X) = "" Then Exit Sub

wFilex = Trim(X)
If wFile <> wFilex Then
    Old_YBIATAB0 = SP_Internet
    New_YBIATAB0 = SP_Internet
    New_YBIATAB0.BIATABTXT = wFilex
    Parametrage_Update
    SP_Internet = New_YBIATAB0
End If
'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Internet S"
newYNOTPAYLOG.NOTPAYLOGX = "lien S & P: " & wFilex
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

Call frmElpPrt.IExplore(wFilex)

End Sub

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Notation Pays_cmdSelect_Ok ........"): DoEvents
fgSelect.Visible = False
fraSelect_Options.Visible = False
fraSelect_Import.Visible = False

Select Case cmdSelect_SQL_K
    Case "1": fraSelect_Options.Visible = True: cmdSelect_SQL_1
    Case "1h":  fraSelect_Options.Visible = True: cmdSelect_SQL_1h
    Case "Ex": cmdSelect_Export_xlsx
    Case "Ec": cmdSelect_Export_Compta
    Case "Em": cmdSelect_Export_Mail
    Case "Sc": cmdSelect_SAB_Client
    Case "Sp": cmdSelect_SAB_ZBALTAB0
    Case "I": cmdSelect_Import
    Case "J": fraSelect_Options.Visible = True: cmdSelect_SQL_Journalisation
    Case "L": fraSelect_Options.Visible = True: cmdSelect_SQL_YNOTPAYLOG
    
    Case "Pn": cmdSelect_SQL_P 'Reprise_Param '
    Case "Rn": Reprise_Notation
    Case "Rp": Reprise_Pays_Import
    Case "Rb": Reprise_Param
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< Notation Pays_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub fgPays_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
 If fgPays.Rows > 1 Then
    fgPays.Col = 0: txtPAYS_ISO = Trim(fgPays.Text)
    fgPays.Col = 1: txtPAYS_SAB = Trim(fgPays.Text)
    fgPays.Col = 2: txtPAYS_Coface = Trim(fgPays.Text)
    fgPays.Col = 3: txtPAYS_OCDE_Code = Trim(fgPays.Text)
    fgPays.Col = 4: txtPAYS_OCDE_lib = Trim(fgPays.Text)
    fgPays.Col = 5: txtPAYS_SP = Trim(fgPays.Text)
    
    txtPAYS_ISO.Enabled = False
    fraPays.Visible = True
End If

End Sub


Private Sub fgSAB_Client_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

fgSAB_Client.Col = 0
Call fgSAB_Client_Detail_Display(fgSAB_Client.Text)

End Sub


Private Sub lstPays_Click()
Dim V
Dim K As Integer, blnOk As Boolean

Me.Enabled = False: Me.MousePointer = vbHourglass
newYNOTPAY0.NOTPAYISO = Mid$(lstPays.Text, 1, 2)
Select Case cmdSelect_SQL_K
    Case "1":
        If blnDetail_Copy Then
        '_____________________________________________________________
            lstPays.Visible = False: blnDetail_Copy = False
            fraDetail.Visible = False
            If newYNOTPAY0.NOTPAYISO <> "  " Then
                V = sqlYNOTPAY0_Insert(newYNOTPAY0)
                If Not IsNull(V) Then
                    Call MsgBox(V, vbCritical, "Copie Notation Pays")
                End If
                '++++++++++++++++++++++++++++++++++++++++++
                newYNOTPAYLOG.NOTPAYLOGK = "Pays Copy"
                newYNOTPAYLOG.NOTPAYLOGX = Trim(newYNOTPAY0.NOTPAYISO) & " : " & Trim(newYNOTPAY0.NOTPAYLIB)
                cmdYNOTPAYLOG_New
                '++++++++++++++++++++++++++++++++++++++++++

                cmdSelect_Ok_Click
            End If
        Else
        '_____________________________________________________________
        
            blnOk = False
            For K = 1 To arrYNOTPAY0_Nb
                If newYNOTPAY0.NOTPAYISO = arrYNOTPAY0(K).NOTPAYISO Then
                    fgSelect.TopRow = K
                    blnOk = True
                    Exit For
                End If
            Next K
            If blnOk Then
                lstPays.Visible = False
                fgSelect.Row = K
                fgSelect_Display_Init
            End If
        End If
    Case "J":
            lstPays.Visible = False
            txtSelect_NOTPAYISO = Mid$(lstPays.Text, 1, 2)
End Select


blnControl = False
txtSelect_Pays = ""
blnControl = True

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

If fraParam.Visible Then
    fraParam.Visible = False: lstPays.Visible = False
    Exit Sub
End If

If fraPays.Visible Then
    fraPays.Visible = False
    Exit Sub
End If

If fgSAB_Client_Detail.Visible Then
    fgSAB_Client_Detail.Visible = False
    Exit Sub
End If
If fgSAB_Client.Visible Then
    fgSAB_Client.Visible = False
    Exit Sub
End If

If lstPays.Visible Then
    lstPays.Visible = False: blnDetail_Copy = False
    Exit Sub
End If

If txtDetail.Visible Then
    txtDetail.Visible = False
    Exit Sub
End If
If fraDetail.Visible Then
    fraDetail.Visible = False
    fraJRNENT0.Visible = False
    Exit Sub
End If
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
If txtDetail.Visible Then
    fraDetail_Control
Else
    If SSTab1.Tab = 0 Then
        If Not fgSelect.Version Then cmdSelect_Ok_Click
    Else
        SendKeys "{TAB}"
    End If
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
Dim wOrigine As String
On Error Resume Next

txtDetail.Visible = False

If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        fgSelect_Display_Init
   End If
End If

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

Public Sub fgPays_Reset()
fgPays.Clear
fgPays_Sort1 = 0: fgPays_Sort2 = 0
fgPays_Sort1_Old = -1
fgPays_RowDisplay = 0: fgPays_RowClick = 0
fgPays_arrIndex = fgPays.Cols - 1
blnfgPays_DisplayLine = False
fgPays_SortAD = 6
fgPays.LeftCol = 0

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







Private Sub txtDetail_Change()
txtDetail_blnUpdate = True
End Sub

Private Sub txtDetail_KeyPress(KeyAscii As Integer)
Select Case txtDetail_Type
    Case "A": KeyAscii = convUCase(KeyAscii)
    Case "N": Call num_KeyAscii(KeyAscii)
    Case "S": Call num_KeyAsciiS(KeyAscii)
    Case "C": Call num_KeyAsciiD(KeyAscii, txtDetail)
End Select
End Sub







Private Sub txtDetail_NOTPAYBIAD_Change()
If fraDetail_Update.Enabled Then fraDetail_Control
End Sub

Private Sub txtDetail_NOTPAYCEG_Change()
If fraDetail_Update.Enabled Then fraDetail_Control: txtDetail_NOTPAYCEG.SetFocus

End Sub

Private Sub txtDetail_NOTPAYCEG_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNumS(KeyAscii)
End Sub


Private Sub txtDetail_NOTPAYCOFD_Change()
If fraDetail_Update.Enabled Then fraDetail_Control
End Sub

Private Sub txtDetail_NOTPAYFISC_Change()
If fraDetail_Update.Enabled Then fraDetail_Control: txtDetail_NOTPAYFISC.SetFocus

End Sub

Private Sub txtDetail_NOTPAYFISC_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNumS(KeyAscii)
End Sub


Private Sub txtDetail_NOTPAYOCDD_Change()
If fraDetail_Update.Enabled Then fraDetail_Control
End Sub

Private Sub txtDetail_NOTPAYSPD_Change()
If fraDetail_Update.Enabled Then fraDetail_Control
End Sub

Private Sub txtDetail_NOTPAYTXT_Change()
If fraDetail_Update.Enabled Then fraDetail_Control: txtDetail_NOTPAYTXT.SetFocus

End Sub

Private Sub txtParam_Code_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtParam_Taux_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtParam_Taux)
End Sub


Private Sub txtPAYS_ISO_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtPAYS_OCDE_Code_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelect_AMJMIN_Change()
cmdSelect_Reset

End Sub

Private Sub txtselect_amjmin_KeyPress(KeyAscii As Integer)
cmdSelect_Reset

End Sub

Private Sub txtSelect_NOTPAYXAMJ_Change()
cmdSelect_Reset

End Sub


Private Sub txtSelect_NOTPAYISO_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_NOTPAYISO_Click()
chkSelect_NOTPAYSEQ = "1"
cmdSelect_Reset

End Sub

Private Sub txtSelect_NOTPAYISO_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub













Public Sub fraDetail_Control()
Dim V, X As String, blnOk As Boolean, K As Integer
Dim xNOTPAYBIAN As String
Dim kNOTPAYBIAN As Integer
fraDetail_Update.Enabled = False
blnOk = False
newYNOTPAY0 = oldYNOTPAY0

newYNOTPAY0.NOTPAYCOFA = cboDetail_NOTPAYCOFA
If newYNOTPAY0.NOTPAYCOFA = oldYNOTPAY0.NOTPAYCOFA Then
    newYNOTPAY0.NOTPAYCOFK = oldYNOTPAY0.NOTPAYCOFK
    cboDetail_NOTPAYCOFA.BackColor = RGB(255, 255, 230)
    cboDetail_NOTPAYCOFA.ForeColor = vbBlue
Else
    newYNOTPAY0.NOTPAYCOFK = "M"
    cboDetail_NOTPAYCOFA.BackColor = RGB(0, 255, 0)
    cboDetail_NOTPAYCOFA.ForeColor = vbMagenta
End If
If Trim(newYNOTPAY0.NOTPAYCOFA) = "" Then
    newYNOTPAY0.NOTPAYCOFD = 0
Else
    Call DTPicker_Control(txtDetail_NOTPAYCOFD, X)
    newYNOTPAY0.NOTPAYCOFD = X
End If
newYNOTPAY0.NOTPAYCOF2 = cboDetail_NOTPAYCOF2
If newYNOTPAY0.NOTPAYCOF2 = oldYNOTPAY0.NOTPAYCOF2 Then
    cboDetail_NOTPAYCOF2.BackColor = RGB(255, 255, 230)
    cboDetail_NOTPAYCOF2.ForeColor = vbBlue
Else
    cboDetail_NOTPAYCOF2.BackColor = RGB(0, 255, 0)
    cboDetail_NOTPAYCOF2.ForeColor = vbMagenta
End If


newYNOTPAY0.NOTPAYOCDE = cboDetail_NOTPAYOCDE
If newYNOTPAY0.NOTPAYOCDE = oldYNOTPAY0.NOTPAYOCDE Then
    newYNOTPAY0.NOTPAYOCDK = oldYNOTPAY0.NOTPAYOCDK
    cboDetail_NOTPAYOCDE.BackColor = RGB(255, 255, 230)
    cboDetail_NOTPAYOCDE.ForeColor = vbBlue
Else
    newYNOTPAY0.NOTPAYOCDK = "M"
    cboDetail_NOTPAYOCDE.BackColor = RGB(0, 255, 0)
    cboDetail_NOTPAYOCDE.ForeColor = vbMagenta
End If
If Trim(newYNOTPAY0.NOTPAYOCDE) = "" Then
    newYNOTPAY0.NOTPAYOCDD = 0
Else
    Call DTPicker_Control(txtDetail_NOTPAYOCDD, X)
    newYNOTPAY0.NOTPAYOCDD = X
End If
newYNOTPAY0.NOTPAYSP = cboDetail_NOTPAYSP
If newYNOTPAY0.NOTPAYSP = oldYNOTPAY0.NOTPAYSP Then
    newYNOTPAY0.NOTPAYSPK = oldYNOTPAY0.NOTPAYSPK
    cboDetail_NOTPAYSP.BackColor = RGB(255, 255, 230)
    cboDetail_NOTPAYSP.ForeColor = vbBlue
Else
    newYNOTPAY0.NOTPAYSPK = "M"
    cboDetail_NOTPAYSP.BackColor = RGB(0, 255, 0)
    cboDetail_NOTPAYSP.ForeColor = vbMagenta
End If
If Trim(newYNOTPAY0.NOTPAYSP) = "" Then
    newYNOTPAY0.NOTPAYSPD = 0
Else
    Call DTPicker_Control(txtDetail_NOTPAYSPD, X)
    newYNOTPAY0.NOTPAYSPD = X
End If
newYNOTPAY0.NOTPAYCEG = Val(txtDetail_NOTPAYCEG)
If newYNOTPAY0.NOTPAYCEG = oldYNOTPAY0.NOTPAYCEG Then
    txtDetail_NOTPAYCEG.BackColor = RGB(255, 255, 230)
Else
    txtDetail_NOTPAYCEG.BackColor = RGB(0, 255, 0)
End If

K = Val(txtDetail_NOTPAYFISC)
If K = 0 Then
    newYNOTPAY0.NOTPAYFISC = "  "
Else
    newYNOTPAY0.NOTPAYFISC = K
End If
If newYNOTPAY0.NOTPAYFISC = oldYNOTPAY0.NOTPAYFISC Then
    txtDetail_NOTPAYFISC.BackColor = RGB(255, 255, 230)
Else
    txtDetail_NOTPAYFISC.BackColor = RGB(0, 255, 0)
End If
newYNOTPAY0.NOTPAYTXT = Trim(txtDetail_NOTPAYTXT)
If newYNOTPAY0.NOTPAYTXT = oldYNOTPAY0.NOTPAYTXT Then
    txtDetail_NOTPAYTXT.BackColor = RGB(255, 255, 230)
Else
    txtDetail_NOTPAYTXT.BackColor = RGB(0, 255, 0)
End If
If chkDetail_NOTPAYPROV = "1" Then
    newYNOTPAY0.NOTPAYPROV = "P"
Else
    newYNOTPAY0.NOTPAYPROV = " "
End If

'____________________________________________________________________
If blnDetail_NOTPAYBIAK Then
        blnOk = True
        newYNOTPAY0.NOTPAYBIAN = cboDetail_NOTPAYBIAN
        newYNOTPAY0.NOTPAYBIAK = "M"
    
        cboDetail_NOTPAYBIAN.BackColor = RGB(255, 128, 255)
Else
    kNOTPAYBIAN = newYNOTPAY0.NOTPAYCEG
    For K = 1 To arrCoface_Nb
        If newYNOTPAY0.NOTPAYCOFA = arrCoface(K).NOTPAYCOFA Then
            kNOTPAYBIAN = kNOTPAYBIAN + arrCoface(K).NOTPAYTAUX * 2
            Exit For
        End If
    Next K
    For K = 1 To arrOCDE_Nb
        If newYNOTPAY0.NOTPAYOCDE = arrOCDE(K).NOTPAYOCDE Then
            kNOTPAYBIAN = kNOTPAYBIAN + arrOCDE(K).NOTPAYTAUX
            Exit For
        End If
    Next K
    
    X = kNOTPAYBIAN
    Select Case Len(X)
        Case 1: newYNOTPAY0.NOTPAYBIAN = "  " & X
        Case 2: newYNOTPAY0.NOTPAYBIAN = " " & X
        Case Else: newYNOTPAY0.NOTPAYBIAN = X
    End Select
    newYNOTPAY0.NOTPAYBIAK = "A"
End If
    
For K = 1 To arrBIAN_Nb
    If newYNOTPAY0.NOTPAYBIAN = arrBIAN(K).NOTPAYBIAN Then
        blnOk = True
        newYNOTPAY0.NOTPAYTAUX = arrBIAN(K).NOTPAYTAUX
        libDetail_NOTPAYTAUX = Format$(newYNOTPAY0.NOTPAYTAUX, "##0.00") & "%"
        cboDetail_NOTPAYBIAN = newYNOTPAY0.NOTPAYBIAN
        Exit For
    End If
Next K

If Not blnOk Then
    Call MsgBox("Le code BIAN n'est pas connu : " & newYNOTPAY0.NOTPAYBIAN, vbExclamation, "NOTPAYBIAN contrôle")
Else
          

    If newYNOTPAY0.NOTPAYBIAN = oldYNOTPAY0.NOTPAYBIAN Then
        cboDetail_NOTPAYBIAN.BackColor = RGB(255, 255, 230)
    Else
        If Not blnDetail_NOTPAYBIAK Then cboDetail_NOTPAYBIAN.BackColor = RGB(0, 255, 0)
    End If
    If newYNOTPAY0.NOTPAYTAUX = oldYNOTPAY0.NOTPAYTAUX Then
        libDetail_NOTPAYTAUX.BackColor = RGB(240, 240, 240)
        libDetail_NOTPAYTAUX.ForeColor = RGB(0, 128, 0)
    Else
        libDetail_NOTPAYTAUX.BackColor = RGB(0, 255, 0)
        libDetail_NOTPAYTAUX.ForeColor = vbMagenta
    End If
End If


Call DTPicker_Control(txtDetail_NOTPAYBIAD, X)
newYNOTPAY0.NOTPAYBIAD = X


'If blnNOTPAYBIAN_à_Calculer Then fraDetail_Control_NOTPAYBIAN_Auto
On Error Resume Next
If blnOk Then
    cmdDetail_Update.Visible = YNOTPAY0_Aut.Saisir
    If newYNOTPAY0.NOTPAYSEQ = 0 Then cmdDetail_New.Visible = YNOTPAY0_Aut.Saisir
Else
    newYNOTPAY0.NOTPAYBIAK = "M"
    cmdDetail_Update.Visible = False: cmdDetail_New.Visible = False
End If
On Error Resume Next
If cmdDetail_New.Visible Then
    cmdDetail_New.SetFocus
Else
    cmdDetail_Quit.SetFocus
End If

Exit_sub:
fraDetail_Update.Enabled = True
End Sub




Private Sub txtSelect_NOTPAYXAMJ_KeyPress(KeyAscii As Integer)
cmdSelect_Reset

End Sub


Public Sub Reprise_Notation()
Dim Nb As Integer, NbOk As Integer
Dim X As String, xIn As String
Dim K As Integer, xin_K As Integer
Dim blnOk As Boolean

Call lstErr_Clear(lstErr, cmdContext, "C:\Temp\Notation Pays\YNOTPAY0_Reprise.csv "): DoEvents
 Nb = 0: NbOk = 0
Open Trim("C:\Temp\Notation Pays\YNOTPAY0_Reprise.csv") For Input As #3
'

Do Until EOF(3)
    Line Input #3, xIn
    Nb = Nb + 1
    xin_K = 0
    X = CSV_Scan(xIn, xin_K)
    X = CSV_Scan(xIn, xin_K)
    blnOk = False
    For K = 1 To arrPays_NB
        If X = arrPays_Lib(K) Then
            blnOk = True
            xYNOTPAY0.NOTPAYISO = arrPAYS_ISO(K)
            Exit For
        End If
    Next K
    If Not blnOk Then
        blnOk = True
        Select Case X
            Case "ARABIE SAOUD": xYNOTPAY0.NOTPAYISO = "SA"
            Case "BURKINA-FASO": xYNOTPAY0.NOTPAYISO = "BF"
            Case "CENTRAFRIQUE": xYNOTPAY0.NOTPAYISO = "CF"
            Case "CHINE": xYNOTPAY0.NOTPAYISO = "CN"
            Case "CONGO REP": xYNOTPAY0.NOTPAYISO = "CG"
            Case "COREE DU SUD": xYNOTPAY0.NOTPAYISO = "KR"
            Case "COTE D'IVOIR": xYNOTPAY0.NOTPAYISO = "CI"
            Case "EGYPTE": xYNOTPAY0.NOTPAYISO = "EG"
            Case "EMIR.ARA.UNI": xYNOTPAY0.NOTPAYISO = "AE"
            Case "EQUATEUR": xYNOTPAY0.NOTPAYISO = "EC"
            Case "GUINEE": xYNOTPAY0.NOTPAYISO = "GN"
            Case "GUINEE-BISS.": xYNOTPAY0.NOTPAYISO = "GW"
            Case "HONG-KONG": xYNOTPAY0.NOTPAYISO = "HK"
            Case "ILES CAYMANS": xYNOTPAY0.NOTPAYISO = "KY"
            Case "MALTE": xYNOTPAY0.NOTPAYISO = "MT"
            Case "OMAN": xYNOTPAY0.NOTPAYISO = "OM"
            Case "PHILLIPINES": xYNOTPAY0.NOTPAYISO = "PH"
            Case "SAINT VINCENT": xYNOTPAY0.NOTPAYISO = "VC"
            Case "SERBIE MONTENEGRO": xYNOTPAY0.NOTPAYISO = "YU"
            Case "TCHEQUE REP.": xYNOTPAY0.NOTPAYISO = "CZ"
            Case "YEMEN SUD": xYNOTPAY0.NOTPAYISO = "YE"
            Case Else: blnOk = False: NbOk = NbOk + 1: Call lstErr_AddItem(lstErr, cmdContext, xIn): DoEvents
        End Select
    End If
    
    xYNOTPAY0.NOTPAYSEQ = 0       ' N° séquence (info)
    xYNOTPAY0.NOTPAYHAMJ = 0
    
    xYNOTPAY0.NOTPAYPROV = "P"                 ' Provisionable = 'P'
    xYNOTPAY0.NOTPAYCOFA = Trim(CSV_Scan(xIn, xin_K))    ' notation coface
    xYNOTPAY0.NOTPAYCOF2 = ""
    xYNOTPAY0.NOTPAYCOFK = "R"                 ' notation coface Auto / Manuel
    xYNOTPAY0.NOTPAYCOFD = 20090531            ' DATE maj
    X = CSV_Scan(xIn, xin_K)
    If X = "-" Then X = " "
    xYNOTPAY0.NOTPAYOCDE = X                   ' notation OCDE
    
    xYNOTPAY0.NOTPAYOCDK = "R"                 ' notation OCDE Auto / Manuel
    xYNOTPAY0.NOTPAYOCDD = 20090531            ' DATE maj
    X = Trim(CSV_Scan(xIn, xin_K))
    If X = "nc" Then X = " "
    xYNOTPAY0.NOTPAYSP = X                     ' notation S & P
    xYNOTPAY0.NOTPAYSPK = "R"                  ' notation S & P Auto / Manuel
    xYNOTPAY0.NOTPAYSPD = 20090531             ' DATE maj
    
    'X = CSV_Scan(xIn, xIn_K)
    xYNOTPAY0.NOTPAYCEG = Val(CSV_Scan(xIn, xin_K))     ' critère événement grave
    
    
    X = Trim(CSV_Scan(xIn, xin_K))
    If X = "#N/A" Then
        xYNOTPAY0.NOTPAYBIAN = "   "                ' notation BIA
    Else
        Select Case Len(X)
            Case 1: xYNOTPAY0.NOTPAYBIAN = "  " & X
            Case 2: xYNOTPAY0.NOTPAYBIAN = " " & X
            Case Else: xYNOTPAY0.NOTPAYBIAN = X
        End Select
        
    End If
    
    xYNOTPAY0.NOTPAYBIAK = "R"                 ' notation BIA Auto / Manuel
    xYNOTPAY0.NOTPAYBIAD = 20090531            ' DATE maj
    X = Replace(Trim(CSV_Scan(xIn, xin_K)), ",", ".")
    xYNOTPAY0.NOTPAYTAUX = Val(X)   ' taux BIA
    X = CSV_Scan(xIn, xin_K)
    If X = "-" Or X = "NP" Then
        xYNOTPAY0.NOTPAYFISC = " "
    Else
        xYNOTPAY0.NOTPAYFISC = Format(Val(X), "#0")   ' taux fisc
    End If
    xYNOTPAY0.NOTPAYTXT = ""                   ' commentaire
    xYNOTPAY0.NOTPAYXAMJ = DSys                ' DATE maj
    xYNOTPAY0.NOTPAYXHMS = time_Hms            ' heure maj
    xYNOTPAY0.NOTPAYXUSR = usrName_UCase       ' utilisateur maj
    sqlYNOTPAY0_Insert xYNOTPAY0
Loop
MsgBox "Nb ERR  : " & NbOk & " / " & Nb, vbCritical, "Reprise YNOTPAY0"
Close #3

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "R_Notation"
newYNOTPAYLOG.NOTPAYLOGX = "Reprise Notation : " & "Nb ERR  : " & NbOk & " / " & Nb
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

End Sub

Public Sub Reprise_Param()
Dim Nb As Integer, NbOk As Integer
Dim X1 As String, X2 As String, X3 As String, xIn As String
Dim dblX As Double, xin_K As Integer
Dim blnOk As Boolean

Call lstErr_Clear(lstErr, cmdContext, "C:\Temp\Notation Pays\YNOTPAY0_Reprise_Param.csv "): DoEvents
 Nb = 0: NbOk = 0
Open Trim("C:\Temp\Notation Pays\YNOTPAY0_Reprise_param.csv") For Input As #3
'

Do Until EOF(3)
    Line Input #3, xIn
    Nb = Nb + 1
    xin_K = 0
    X1 = CSV_Scan(xIn, xin_K)
    X2 = CSV_Scan(xIn, xin_K)
    X3 = CSV_Scan(xIn, xin_K)
    dblX = Val(X3)
    blnOk = False
    
    xYNOTPAY0.NOTPAYISO = "$$"
    xYNOTPAY0.NOTPAYSEQ = Nb       ' N° séquence (info)
    
    xYNOTPAY0.NOTPAYPROV = ""                 ' Provisionable = 'P'
    xYNOTPAY0.NOTPAYCOFA = ""       ' notation coface
    xYNOTPAY0.NOTPAYCOFK = ""                 ' notation coface Auto / Manuel
    xYNOTPAY0.NOTPAYCOFD = 0            ' DATE maj
    xYNOTPAY0.NOTPAYOCDE = ""                  ' notation OCDE
    xYNOTPAY0.NOTPAYOCDK = ""                 ' notation OCDE Auto / Manuel
    xYNOTPAY0.NOTPAYOCDD = 0        ' DATE maj
    xYNOTPAY0.NOTPAYSP = ""                     ' notation S & P
    xYNOTPAY0.NOTPAYSPK = ""                  ' notation S & P Auto / Manuel
    xYNOTPAY0.NOTPAYSPD = 0             ' DATE maj
    xYNOTPAY0.NOTPAYBIAN = ""  ' notation BIA
    xYNOTPAY0.NOTPAYBIAK = ""                 ' notation BIA Auto / Manuel
    xYNOTPAY0.NOTPAYBIAD = 0            ' DATE maj
    
    Select Case Mid$(X1, 1, 1)
        Case "1":
                X3 = ""
                xYNOTPAY0.NOTPAYCOFA = X2       ' notation coface
                xYNOTPAY0.NOTPAYCOFK = "R"                 ' notation coface Auto / Manuel
                xYNOTPAY0.NOTPAYCOFD = 20090531            ' DATE maj

        Case "2":
                X3 = ""
                xYNOTPAY0.NOTPAYOCDE = X2                   ' notation OCDE
                xYNOTPAY0.NOTPAYOCDK = "R"                 ' notation OCDE Auto / Manuel
                xYNOTPAY0.NOTPAYOCDD = 20090531            ' DATE maj
        Case "3":
                xYNOTPAY0.NOTPAYSP = X2                     ' notation S & P
                xYNOTPAY0.NOTPAYSPK = "R"                  ' notation S & P Auto / Manuel
                xYNOTPAY0.NOTPAYSPD = 20090531             ' DATE maj
       Case "4":
                X3 = ""
                Select Case Len(X2)
                    Case 1: X2 = "  " & X2
                    Case 2: X2 = " " & X2
                    Case Else: 'X2 = X2
                End Select
                xYNOTPAY0.NOTPAYBIAN = X2   ' notation BIA
                xYNOTPAY0.NOTPAYBIAK = "R"                 ' notation BIA Auto / Manuel
                xYNOTPAY0.NOTPAYBIAD = 20090531            ' DATE maj
    End Select
    
    
    xYNOTPAY0.NOTPAYCEG = 0    ' critère événement grave
    xYNOTPAY0.NOTPAYTAUX = dblX  ' taux BIA
    
    xYNOTPAY0.NOTPAYFISC = " "
    xYNOTPAY0.NOTPAYTXT = X1                 ' commentaire
    Mid$(xYNOTPAY0.NOTPAYTXT, 10, 5) = X2
    Mid$(xYNOTPAY0.NOTPAYTXT, 15, 10) = X3
    xYNOTPAY0.NOTPAYXAMJ = DSys                ' DATE maj
    xYNOTPAY0.NOTPAYXHMS = time_Hms            ' heure maj
    xYNOTPAY0.NOTPAYXUSR = usrName_UCase       ' utilisateur maj
    
    sqlYNOTPAY0_Insert xYNOTPAY0
Loop
MsgBox "Nb ERR  : " & NbOk & " / " & Nb, vbCritical, "Reprise YNOTPAY0"
Close #3

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "R_Param"
newYNOTPAYLOG.NOTPAYLOGX = "Reprise paramétrage note => taux : " & "Nb ERR  : " & NbOk & " / " & Nb
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

End Sub
Public Sub Reprise_Pays_Import()
Dim Nb As Integer, NbOk As Integer, K As Integer
Dim K1 As Integer, K2 As Integer, K3 As Integer
Dim X1 As String, X2 As String, X3 As String, xIn As String, xISO As String
Dim dblX As Double, xin_K As Integer
Dim blnOk As Boolean, blnISO As Boolean
Dim arrFR(250) As String, arrGB(250) As String
Call lstErr_Clear(lstErr, cmdContext, "C:\Temp\Notation Pays\OCDE.txt "): DoEvents
 Nb = 0: NbOk = 0
Open Trim("C:\Temp\Notation Pays\OCDE.txt") For Input As #3

ReDim arrPays_Import(250)
blnOk = False

Do Until EOF(3)
    Line Input #3, xIn
    If IsNumeric(xIn) Then
        If Val(xIn) = Nb + 1 Then
            Nb = Nb + 1
            arrPays_Import(Nb).BIATABID = "YNOTPAY0"
            arrPays_Import(Nb).BIATABK1 = "Pays"
            arrPays_Import(Nb).BIATABK2 = Space$(12)
            arrPays_Import(Nb).BIATABTXT = "*"
            Line Input #3, xIn
            X1 = xIn
            Select Case X1
                Case "AND": X2 = "AD"
                Case "AGO": X2 = "AO"
                Case "ATG": X2 = "AG"
                Case "ARM": X2 = "AM"
                Case "ABW": X2 = "AW"
                Case "AUT": X2 = "AT"
                Case "BHS": X2 = "BS"
                Case "BHR": X2 = "BH"
                Case "BGD": X2 = "BD"
                Case "BRB": X2 = "BB"
                Case "BLR": X2 = "BY"
                Case "BLZ": X2 = "BZ"
                Case "BEN": X2 = "BJ"
                Case "BIH": X2 = "BA"
                Case "BRN": X2 = "BN"
                Case "BDI": X2 = "BI"
                Case "CPV": X2 = "CV"
                Case "CYM": X2 = "KY"
                Case "CAF": X2 = "CF"
                Case "TCD": X2 = "TD"
                Case "CHI": X2 = "JE"
                Case "CHL": X2 = "CL"
                Case "CHN": X2 = "CN"
                Case "COM": X2 = "KM"
                Case "COG": X2 = "CG"
                Case "COD": X2 = "CD"
                Case "DNK": X2 = "DK"
                Case "SLV": X2 = "SV"
                Case "EST": X2 = "EE"
                Case "FRO": X2 = "FO"
                Case "PYF": X2 = "PF"
                Case "GRL": X2 = "GL"
                Case "GRD": X2 = "GD"
                Case "GIN": X2 = "GN"
                Case "GNB": X2 = "GW"
                Case "GNQ": X2 = "GQ"
                Case "GUY": X2 = "GY"
                Case "IRQ": X2 = "IQ"
                Case "IRL": X2 = "IE"
                Case "ISR": X2 = "IL"
                Case "JAM": X2 = "JM"
                Case "KAZ": X2 = "KZ"
                Case "KOR": X2 = "KR"
                Case "PRK": X2 = "KP"
                Case "LBR": X2 = "LR"
                Case "LBY": X2 = "LY"
                Case "MAC": X2 = "MO"
                Case "MAR": X2 = "MA"
                Case "MDG": X2 = "MG"
                Case "MDV": X2 = "MV"
                Case "MLT": X2 = "MT"
                Case "MYT": X2 = "YT"
                Case "MEX": X2 = "MX"
                Case "FSM": X2 = "FM"
                Case "MDA": X2 = "MD"
                Case "MNE": X2 = "ME"
                Case "MOZ": X2 = "MZ"
                Case "MNP": X2 = "MP"
                Case "PAK": X2 = "PK"
                Case "PLW": X2 = "PW"
                Case "PNG": X2 = "PG"
                Case "PRY": X2 = "PY"
                Case "POL": X2 = "PL"
                Case "PRT": X2 = "PT"
                Case "SEN": X2 = "SN"
                Case "SRB": X2 = "RS"
                Case "SYC": X2 = "SC"
                Case "SVK": X2 = "SK"
                Case "SVN": X2 = "SI"
                Case "SLB": X2 = "SB"
                Case "SUR": X2 = "SR"
                Case "SWZ": X2 = "SZ"
                Case "SWE": X2 = "SE"
                Case "TUN": X2 = "TN"
                Case "TUR": X2 = "TR"
                Case "TKM": X2 = "TM"
                Case "UKR": X2 = "UA"
                Case "ARE": X2 = "AE"
                Case "URY": X2 = "UY"
               
                Case Else: X2 = Mid$(X1, 1, 2)
            End Select
            
            arrPays_Import(Nb).BIATABK2 = X2
            Line Input #3, xIn
            arrGB(Nb) = LCase$(xIn)
            Line Input #3, xIn
            Mid$(arrPays_Import(Nb).BIATABTXT, 1, 3) = X1
            arrFR(Nb) = LCase$(xIn)
            Mid$(arrPays_Import(Nb).BIATABTXT, 4, 32) = arrFR(Nb)
            Line Input #3, xIn
            Line Input #3, xIn
        End If
    End If
Loop
MsgBox "Nb ERR  : " & NbOk & " / " & Nb, vbCritical, "Reprise PAys Import"
Close #3

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "R_Pays"
newYNOTPAYLOG.NOTPAYLOGX = "Reprise OCDE : " & "Nb ERR  : " & NbOk & " / " & Nb
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

arrPays_Import_Nb = Nb

Call lstErr_Clear(lstErr, cmdContext, "C:\Temp\Notation Pays\coface.txt "): DoEvents
 Nb = 0: NbOk = 0
Open Trim("C:\Temp\Notation Pays\coface.txt") For Input As #3


Do Until EOF(3)
    Line Input #3, xIn
    xIn = Trim(xIn)
    If xIn <> "" Then
        If Not blnOk Then
            If InStr(1, xIn, "@Rating") Then blnOk = True
        Else
            Nb = Nb + 1
            K1 = InStr(1, xIn, " ")
            K2 = InStr(K1 + 1, xIn, " ")
            K = InStr(K2 + 1, xIn, " ")
            If K > 0 Then
                K1 = K2
                K2 = K
                K = InStr(K2 + 1, xIn, " ")
                If K > 0 Then
                    K1 = K2
                    K2 = K
                End If
            End If
            X = LCase(Mid$(xIn, 1, K1 - 1))
            blnISO = False
            For K = 1 To arrPays_Import_Nb
                If X = arrFR(K) Then blnISO = True: Mid$(arrPays_Import(K).BIATABTXT, 36, 32) = X: Exit For
            Next K
            If Not blnISO Then
                Select Case X
                    Case "bahrein": xISO = "BH"
                    Case "bielorussie": xISO = "BY"
                    Case "bosnie herzégovine": xISO = "BA"
                    Case "cap vert": xISO = "CV"
                    Case "corée du nord": xISO = "KP"
                    Case "corée du sud": xISO = "KR"
                    Case "costa-rica": xISO = "CR"
                    Case "egypte": xISO = "EG"
                    Case "emirats arabes unis": xISO = "AE"
                    Case "equateur": xISO = "EC"
                    Case "etats-unis": xISO = "US"
                    Case "ethiopie": xISO = "ET"
                    Case "guinée bissau": xISO = "GW"
                    Case "guinée equatoriale": xISO = "GQ"
                    Case "haiti": xISO = "HT"
                    Case "hong kong": xISO = "HK"
                    Case "ile maurice": xISO = "MU"
                    Case "irak": xISO = "IQ"
                    Case "kirghizstan": xISO = "KG"
                    Case "libera": xISO = "LR"
                    Case "macédoine": xISO = "MK"
                    Case "moldavie": xISO = "MD"
                    Case "montenegro": xISO = "ME"
                    Case "nouvelle zélande": xISO = "NZ"
                    Case "ouzbekistan": xISO = "UZ"
                    Case "papouasie nlle guinée": xISO = "PG"
                    Case "rdc (ex-zaïre)": xISO = "CD"
                    Case "rép dominicaine": xISO = "DO"
                    Case "rép. centrafricaine": xISO = "CF"
                    Case "reptchèque": xISO = "CZ"
                    Case "royaume uni": xISO = "GB"
                    Case "russie": xISO = "RU"
                    Case "sao tome": xISO = "ST"
                    Case "slovaquie": xISO = "SK"
                    Case "taïwan": xISO = "TW"
                    Case "trinidad": xISO = "TT"
                    Case "vietnam": xISO = "VN"

                    Case Else: xISO = "??"
                End Select
                For K = 1 To arrPays_Import_Nb
                    If xISO = Trim(arrPays_Import(K).BIATABK2) Then blnISO = True: Mid$(arrPays_Import(K).BIATABTXT, 36, 32) = X: Exit For
                Next K
                If Not blnISO Then Debug.Print "case " & Chr$(34) & X & Chr$(34) & ": xISO = " & Chr$(34) & ".." & Chr$(34)
            End If
        End If
    End If
If X = "zimbabwe" Then Exit Do
Loop
MsgBox "Nb ERR  : " & NbOk & " / " & Nb, vbCritical, "Reprise PAys Import"
Close #3

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "R_Pays"
newYNOTPAYLOG.NOTPAYLOGX = "Reprise Coface : " & "Nb ERR  : " & NbOk & " / " & Nb
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++


Call lstErr_Clear(lstErr, cmdContext, "C:\Temp\Notation Pays\sp.txt "): DoEvents
 Nb = 0: NbOk = 0
Open Trim("C:\Temp\Notation Pays\sp.txt") For Input As #3


Do Until EOF(3)
    Line Input #3, xIn
    xIn = Trim(xIn)
    K3 = InStr(1, xIn, "/")

    If K3 > 0 Then
        Nb = Nb + 1
        K1 = InStr(1, xIn, " ")
        K2 = InStr(1, xIn, "(")
        If K2 <= 0 Then
            K = InStr(K1 + 1, xIn, " ")
            If K > 0 And K < K3 Then
                K1 = K
                K = InStr(K1 + 1, xIn, " ")
                If K > 0 And K < K3 Then
                    K1 = K
                    K = InStr(K1 + 1, xIn, " ")
                    If K > 0 And K < K3 Then
                        K1 = K
                    End If
                End If
            End If
               
            X = LCase(Trim(Mid$(xIn, 1, K1 - 1)))
        Else
            X = LCase(Trim(Mid$(xIn, 1, K2 - 2)))
        End If
        blnISO = False
        For K = 1 To arrPays_Import_Nb
            If X = arrGB(K) Then blnISO = True: Mid$(arrPays_Import(K).BIATABTXT, 68, 32) = X: Exit For
        Next K
        
        If Not blnISO Then
            Select Case X
                Case "abu dhabi": xISO = "AE"
                Case "fiji islands": xISO = "FJ"
                Case "hong kong": xISO = "HK"
                Case "macedonia": xISO = "MK"
                Case "taiwan": xISO = "TW"
                Case "vietnam": xISO = "VN"
                Case "aruba": xISO = "AW"
                Case "barbados": xISO = "BB"
                Case "belize": xISO = "BZ"
                Case "bermuda": xISO = "BM"
                Case "burkina faso": xISO = "BF"
                Case "canada": xISO = "CA"
                Case "cook islands": xISO = "CK"
                Case "czech republic": xISO = "CZ"
                Case "dominican republic": xISO = "DO"
                Case "emirate of ras al": xISO = ".."
                Case "gabonese republic": xISO = "GA"
                Case "grenada": xISO = "GD"
                Case "hellenic republic": xISO = "GR"
                Case "isle of man": xISO = "IM"
                Case "jamaica": xISO = "JM"
                Case "japan": xISO = "JP"
                Case "malaysia": xISO = "MY"
                Case "mongolia": xISO = "MN"
                Case "montserrat": xISO = "MS"
                Case "new zealand": xISO = "NZ"
                Case "russian federation": xISO = "RU"
                Case "slovak republic": xISO = "SK"
                Case "socialist people's libyan arab": xISO = "LY"
                Case "swiss confederation": xISO = "CH"
                Case "ukraine": xISO = "UA"
                Case "united kingdom": xISO = "GB"
                Case "united mexican states": xISO = "MX"
                Case "united states of america": xISO = "US"
                Case Else: xISO = "??"
            End Select
            For K = 1 To arrPays_Import_Nb
                If xISO = Trim(arrPays_Import(K).BIATABK2) Then blnISO = True: Mid$(arrPays_Import(K).BIATABTXT, 68, 32) = X: Exit For
            Next K
            If Not blnISO Then Debug.Print "case " & Chr$(34) & X & Chr$(34) & ": xISO = " & Chr$(34) & ".." & Chr$(34)
        End If
    End If
Loop
MsgBox "Nb ERR  : " & NbOk & " / " & Nb, vbCritical, "Reprise PAys Import"
Close #3

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "R_Pays"
newYNOTPAYLOG.NOTPAYLOGX = "Reprise S & P : " & "Nb ERR  : " & NbOk & " / " & Nb
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++


'Open Trim("C:\Temp\Notation Pays\pays_controle.txt") For Output As #3
'For nb = 1 To arrPays_Import_Nb
'    X = Mid$(arrPays_Import(nb).BIATABK2, 1, 2)
'    For K = 1 To arrPAYS_Nb
'        If X = arrPAYS_ISO(nb) Then: Exit For
'    Next K
'        Print #3, arrPays_Import(nb).BIATABK2 & ";" & arrPays_Lib(K) & ";" & Mid$(arrPays_Import(nb).BIATABTXT, 33, 32) & ";" & Mid$(arrPays_Import(nb).BIATABTXT, 1, 32)
'Next nb
'Close #3
'-------------------------------------------------------
App_Debug = "Reprise_Pays_Import"
'-------------------------------------------------------

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "R_Paramétrage"
newYNOTPAYLOG.NOTPAYLOGX = "Paramétrage liens internet, notepad, date notation, exportation "
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++


'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
For K = 1 To arrPays_Import_Nb
   V = sqlYBIATAB0_Insert(arrPays_Import(K))
Next K

'Lien Internet .....   _________________________________________________________________________

Coface_Internet.BIATABID = "YNOTPAY0"
Coface_Internet.BIATABK1 = "Internet"
Coface_Internet.BIATABK2 = "Coface"
Coface_Internet.BIATABTXT = "www.cofacerating.fr"
V = sqlYBIATAB0_Insert(Coface_Internet)

Coface_Notepad.BIATABID = "YNOTPAY0"
Coface_Notepad.BIATABK1 = "Notepad"
Coface_Notepad.BIATABK2 = "Coface"
Coface_Notepad.BIATABTXT = "V:\Notation Pays\coface.txt"
V = sqlYBIATAB0_Insert(Coface_Notepad)

Coface_Notepad.BIATABID = "YNOTPAY0"
Coface_Notepad.BIATABK1 = "DateNotation"
Coface_Notepad.BIATABK2 = "Coface"
Coface_Notepad.BIATABTXT = "20090531 *          20090531 240000"
V = sqlYBIATAB0_Insert(Coface_Notepad)
'.............................................................
OCDE_Internet.BIATABID = "YNOTPAY0"
OCDE_Internet.BIATABK1 = "Internet"
OCDE_Internet.BIATABK2 = "OCDE"
OCDE_Internet.BIATABTXT = "http://www.oecd.org/document/49/0,2340,en_2649_201185_17627441_1_1_1_1,00.html"
V = sqlYBIATAB0_Insert(OCDE_Internet)

OCDE_Notepad.BIATABID = "YNOTPAY0"
OCDE_Notepad.BIATABK1 = "Notepad"
OCDE_Notepad.BIATABK2 = "OCDE"
OCDE_Notepad.BIATABTXT = "V:\Notation Pays\OCDE.txt"
V = sqlYBIATAB0_Insert(OCDE_Notepad)

Coface_Notepad.BIATABID = "YNOTPAY0"
Coface_Notepad.BIATABK1 = "DateNotation"
Coface_Notepad.BIATABK2 = "OCDE"
Coface_Notepad.BIATABTXT = "20090531 *          20090531 240000"
V = sqlYBIATAB0_Insert(Coface_Notepad)
'.............................................................

SP_Internet.BIATABID = "YNOTPAY0"
SP_Internet.BIATABK1 = "Internet"
SP_Internet.BIATABK2 = "SP"
SP_Internet.BIATABTXT = "http://www2.standardandpoors.com/portal/site/sp/en/eu/page.topic/ratings_sov/2,1,8,0,0,0,0,0,0,0,4,0,0,50,0,0.html"
V = sqlYBIATAB0_Insert(SP_Internet)

SP_Notepad.BIATABID = "YNOTPAY0"
SP_Notepad.BIATABK1 = "Notepad"
SP_Notepad.BIATABK2 = "SP"
SP_Notepad.BIATABTXT = "V:\Notation Pays\SP.txt"
V = sqlYBIATAB0_Insert(SP_Notepad)

Coface_Notepad.BIATABID = "YNOTPAY0"
Coface_Notepad.BIATABK1 = "DateNotation"
Coface_Notepad.BIATABK2 = "SP"
Coface_Notepad.BIATABTXT = "20090531 *          20090531 240000"
V = sqlYBIATAB0_Insert(Coface_Notepad)
'.............................................................

Coface_Notepad.BIATABID = "YNOTPAY0"
Coface_Notepad.BIATABK1 = "DateNotation"
Coface_Notepad.BIATABK2 = "BIA"
Coface_Notepad.BIATABTXT = "20090531 *          20090531 240000"
V = sqlYBIATAB0_Insert(Coface_Notepad)
'.............................................................


Export_Lien.BIATABID = "YNOTPAY0"
Export_Lien.BIATABK1 = "Export"
Export_Lien.BIATABK2 = ""
Export_Lien.BIATABTXT = "V:\Notation Pays\TAUX_RISQUES_PAYS.xlsx"
V = sqlYBIATAB0_Insert(Export_Lien)

'.............................................................

New_YBIATAB0.BIATABID = "YNOTPAY0"
New_YBIATAB0.BIATABK1 = "Compta"
New_YBIATAB0.BIATABK2 = "Base"
New_YBIATAB0.BIATABTXT = "W:\Notation Pays\Compta Test.mdb"
V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABID = "YNOTPAY0"
New_YBIATAB0.BIATABK1 = "Compta"
New_YBIATAB0.BIATABK2 = "Table"
New_YBIATAB0.BIATABTXT = "T_PAYS_et_Rating"
V = sqlYBIATAB0_Insert(New_YBIATAB0)


New_YBIATAB0.BIATABID = "YNOTPAY0"
New_YBIATAB0.BIATABK1 = "Log"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = "V:\Notation Pays\"
V = sqlYBIATAB0_Insert(New_YBIATAB0)
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


'fgPays_Display

End Sub


Public Sub param_Init()
Dim X As String, K As Integer
Dim xSQL As String

X = "select count(*) as Tally   from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID = 'SAB' and BIATABK1 = 'CLIENAPAY' "
Set rsSab = cnsab.Execute(X)

ReDim arrPAYS_ISO(rsSab("Tally") + 1)
ReDim arrPays_Lib(rsSab("Tally") + 1)
arrPays_NB = 0

txtSelect_NOTPAYISO.Clear
txtSelect_NOTPAYISO.AddItem "   "
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID = 'SAB' and BIATABK1 = 'CLIENAPAY'  order by BIATABK2"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    arrPays_NB = arrPays_NB + 1
    arrPAYS_ISO(arrPays_NB) = Mid$(rsSab("BIATABK2"), 4, 2)
    arrPays_Lib(arrPays_NB) = Trim(Mid$(rsSab("BIATABTXT"), 16, 30))
    
    txtSelect_NOTPAYISO.AddItem arrPAYS_ISO(arrPays_NB) & " - " & arrPays_Lib(arrPays_NB)
    rsSab.MoveNext
Loop
txtSelect_NOTPAYISO.ListIndex = 0

Param_Load_Pondération

wAMJMin = dateElp("MoisAdd", -3, DSys)
Call DTPicker_Set(txtSelect_AMJMIN, wAMJMin) '
Call DTPicker_Set(txtSelect_Import_Amj, YBIATAB0_DATE_CPT_MP1) '

'Pays import_____________________________________________________________________________

arrPays_Import_SQL
fgPays_Display

'Initialisation Log________________________________________________________________________________


newYNOTPAYLOG.NOTPAYLOGD = DSys                ' DATE maj
newYNOTPAYLOG.NOTPAYLOGH = time_Hms            ' heure maj
newYNOTPAYLOG.NOTPAYLOGU = usrName_UCase       ' utilisateur maj
newYNOTPAYLOG.NOTPAYLOGS = 0
newYNOTPAYLOG.NOTPAYLOGX = ""                   ' commentaire


cboNOTPAYLOGK.Clear
cboNOTPAYLOGK.AddItem ""
xSQL = "select distinct NOTPAYLOGK from " & paramIBM_Library_SABSPE & ".YNOTPAYLOG order by NOTPAYLOGK"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    cboNOTPAYLOGK.AddItem Trim(rsSab("NOTPAYLOGK"))
    rsSab.MoveNext
Loop


End Sub

Public Function cmdSelect_Import_NOTPAYBIAN_Auto()
Dim K As Integer, blnOk As Boolean, X As String
Dim xNOTPAYBIAN As String
Dim kNOTPAYBIAN As Integer
cmdSelect_Import_NOTPAYBIAN_Auto = ""


If newYNOTPAY0.NOTPAYBIAK = "M" Then
    Exit Function
End If
'________________________________________________________
kNOTPAYBIAN = newYNOTPAY0.NOTPAYCEG
blnOk = False
For K = 1 To arrCoface_Nb
    If newYNOTPAY0.NOTPAYCOFA = arrCoface(K).NOTPAYCOFA Then
        kNOTPAYBIAN = kNOTPAYBIAN + arrCoface(K).NOTPAYTAUX * 2
        Exit For
    End If
Next K
For K = 1 To arrOCDE_Nb
    If newYNOTPAY0.NOTPAYOCDE = arrOCDE(K).NOTPAYOCDE Then
        kNOTPAYBIAN = kNOTPAYBIAN + arrOCDE(K).NOTPAYTAUX
        Exit For
    End If
Next K

If Trim(newYNOTPAY0.NOTPAYCOFA) = "" And Trim(newYNOTPAY0.NOTPAYOCDE) = "" Then
    xNOTPAYBIAN = ""
Else
    xNOTPAYBIAN = kNOTPAYBIAN
End If
If kNOTPAYBIAN = 0 Then xNOTPAYBIAN = ""
Select Case Len(xNOTPAYBIAN)
    Case 0: newYNOTPAY0.NOTPAYBIAN = "   "
    Case 1: newYNOTPAY0.NOTPAYBIAN = "  " & xNOTPAYBIAN
    Case 2: newYNOTPAY0.NOTPAYBIAN = " " & xNOTPAYBIAN
    Case Else: newYNOTPAY0.NOTPAYBIAN = xNOTPAYBIAN
End Select

blnOk = False
For K = 1 To arrBIAN_Nb
    If newYNOTPAY0.NOTPAYBIAN = arrBIAN(K).NOTPAYBIAN Then
        newYNOTPAY0.NOTPAYBIAK = "A"
        newYNOTPAY0.NOTPAYBIAD = wAMJMin
        newYNOTPAY0.NOTPAYTAUX = arrBIAN(K).NOTPAYTAUX
        blnOk = True: Exit For
    End If
Next K
If Not blnOk Then
    newYNOTPAY0.NOTPAYTAUX = 0
        newYNOTPAY0.NOTPAYBIAK = " "
        newYNOTPAY0.NOTPAYBIAD = 0
    cmdSelect_Import_NOTPAYBIAN_Auto = "code BIAN  : " & newYNOTPAY0.NOTPAYBIAN
End If

End Function

Private Sub txtSelect_Pays_Change()
Dim K As Integer, X As String
If blnControl Then
    fraDetail.Visible = False
    X = Trim(txtSelect_Pays)
    If X = "" Then
        lstPays.Visible = False
    Else
        lstPays.Clear
        
        For K = 1 To arrPays_NB
            If InStr(1, arrPays_Lib(K), X) Then lstPays.AddItem arrPAYS_ISO(K) & " - " & arrPays_Lib(K)
        Next K
        lstPays.Visible = True
    End If
End If
End Sub


Private Sub txtSelect_Pays_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Public Sub fgSelect_Display_Init()
Dim xSQL As String
fraJRNENT0.Visible = False
Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
fgSelect.Col = fgSelect_arrIndex:  arrYNOTPAY0_Index = CLng(fgSelect.Text)
oldYNOTPAY0 = arrYNOTPAY0(arrYNOTPAY0_Index)
newYNOTPAY0 = oldYNOTPAY0
fraDetail_Display
If cmdSelect_SQL_K = "J" Then
    xSQL = "select * from " & paramIBM_Library_SABJRN & ".JRNENT0 " _
         & " where jorcv = " & oldYNOTPAY0.JORCV _
         & " and joSEQN = " & oldYNOTPAY0.JOSEQN
         
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        V = srvJRNENT0_GetBuffer_ODBC(rsSab, xJRNENT0)
        If IsNull(V) Then
            xJRNENT0.JOUSER = oldYNOTPAY0.NOTPAYXUSR
            Call srvJRNENT0_fgX(xJRNENT0, fgJRNENT0)
            fraJRNENT0.Caption = JOENTT_Lib(xJRNENT0.JOENTT)
            fraJRNENT0.ForeColor = vbRed
            fraJRNENT0.Visible = True
        End If
    End If
End If
End Sub

Public Function cmdYNOTPAY0_Control()
Dim wMsg As String
cmdYNOTPAY0_Control = Null
wMsg = ""
If Trim(newYNOTPAY0.NOTPAYCOFA) <> "" Then
    If newYNOTPAY0.NOTPAYCOFD < lastYNOTPAY0.NOTPAYCOFD Then
        wMsg = wMsg & "- la date màj COFACE " & dateImp(newYNOTPAY0.NOTPAYCOFD) _
                    & " est antérieure à la précédente " & dateImp(lastYNOTPAY0.NOTPAYCOFD) _
                    & " ( séq : " & lastYNOTPAY0.NOTPAYSEQ & ")" & vbCrLf
    End If
End If

If Trim(newYNOTPAY0.NOTPAYOCDE) <> "" Then
    If newYNOTPAY0.NOTPAYOCDD < lastYNOTPAY0.NOTPAYOCDD Then
        wMsg = wMsg & "- la date màj OCDE " & dateImp(newYNOTPAY0.NOTPAYOCDD) _
                    & " est antérieure à la précédente " & dateImp(lastYNOTPAY0.NOTPAYOCDD) _
                    & " ( séq : " & lastYNOTPAY0.NOTPAYSEQ & ")" & vbCrLf
    End If
End If

If Trim(newYNOTPAY0.NOTPAYSP) <> "" Then
    If newYNOTPAY0.NOTPAYSPD < lastYNOTPAY0.NOTPAYSPD Then
        wMsg = wMsg & "- la date màj S & P " & dateImp(newYNOTPAY0.NOTPAYSPD) _
                    & " est antérieure à la précédente " & dateImp(lastYNOTPAY0.NOTPAYSPD) _
                    & " ( séq : " & lastYNOTPAY0.NOTPAYSEQ & ")" & vbCrLf
    End If
End If

If newYNOTPAY0.NOTPAYBIAD < lastYNOTPAY0.NOTPAYBIAD Then
    wMsg = wMsg & "- la date màj BIA " & dateImp(newYNOTPAY0.NOTPAYBIAD) _
                & " est antérieure à la précédente " & dateImp(lastYNOTPAY0.NOTPAYBIAD) _
                & " ( séq : " & lastYNOTPAY0.NOTPAYSEQ & ")" & vbCrLf
End If

If wMsg <> "" Then
    Call MsgBox(wMsg, vbExclamation, "Notation pays : Contrôle des dates")
    cmdYNOTPAY0_Control = wMsg
End If
End Function

Public Sub cmdSelect_Export_xlsx()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Exportation en cours ......"): DoEvents

arrYNOTPAY0_SQL " where NOTPAYSEQ =  0 order by NOTPAYISO "

YNOTPAY0_Export


Me.Enabled = True: Me.MousePointer = 0
End Sub
Public Sub cmdSelect_Export_Compta()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Exportation Compta en cours ......"): DoEvents

arrYNOTPAY0_SQL " where NOTPAYSEQ =  0 order by NOTPAYISO "
YNOTPAY0_Compta


Me.Enabled = True: Me.MousePointer = 0
End Sub

Public Sub YNOTPAY0_Export()
On Error GoTo Error_Handler
Dim Nb As Long, wId As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSQL As String
Dim wAMJMin As String, WAMJMax As String
Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim X As String, K As Long, kMax As Long, K2 As Long, K3 As Long
Dim rsSabX As New ADODB.Recordset
'______________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YNOTPAY0' and BIATABK1 = 'Export'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    V = rsYBIATAB0_GetBuffer(rsSab, Export_Lien)
End If

wFile = Trim(Export_Lien.BIATABTXT)
'______________________________________________

X = InputBox("par défaut : " & wFile _
    & vbCrLf & vbCrLf & "     =========================" _
    & vbCrLf & "     =========================", "Notation Pays : nom du fichier d'exportation", wFile)
If Trim(X) = "" Then Exit Sub

wFilex = Trim(X)
'______________________________________________
If wFile <> wFilex Then
    wFile = wFilex
    Old_YBIATAB0 = Export_Lien
    New_YBIATAB0 = Export_Lien
    New_YBIATAB0.BIATABTXT = wFilex
    Parametrage_Update
    Export_Lien = New_YBIATAB0
End If
'_________________________________________


If Dir(wFile) <> "" Then Kill wFile
'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Export"
newYNOTPAYLOG.NOTPAYLOGX = "fichier Excel : " & wFile
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "Notation pays"
    .Subject = "Notation pays"
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "Notation pays"
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
End With

Nb = 1
Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : " & Nb & " enregistrements"): DoEvents

wsExcel.Cells(Nb, 1) = "ISO": wsExcel.Columns(1).ColumnWidth = 8
wsExcel.Cells(Nb, 2) = "Pays": wsExcel.Columns(2).ColumnWidth = 32
wsExcel.Cells(Nb, 3) = "COFACE": wsExcel.Columns(3).ColumnWidth = 8
wsExcel.Columns(3).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
wsExcel.Cells(Nb, 4) = "OCDE": wsExcel.Columns(4).ColumnWidth = 8
wsExcel.Columns(4).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
wsExcel.Cells(Nb, 5) = "S & P": wsExcel.Columns(5).ColumnWidth = 8
wsExcel.Columns(5).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
wsExcel.Cells(Nb, 6) = "C E G": wsExcel.Columns(6).ColumnWidth = 8
wsExcel.Columns(6).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Cells(Nb, 7) = "Note BIA": wsExcel.Columns(7).ColumnWidth = 8
wsExcel.Columns(7).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Cells(Nb, 8) = "Taux": wsExcel.Columns(8).ColumnWidth = 8
wsExcel.Columns(8).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Cells(Nb, 9) = "Date màj": wsExcel.Columns(9).NumberFormat = "mm/dd/yyyy"
wsExcel.Columns(9).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight: wsExcel.Columns(9).ColumnWidth = 16
wsExcel.Cells(Nb, 10) = "Fisc": wsExcel.Columns(10).ColumnWidth = 8
wsExcel.Cells(Nb, 11) = "Commentaire": wsExcel.Columns(11).ColumnWidth = 32

For K = 1 To 11
    wsExcel.Columns(K).VerticalAlignment = Excel.XlVAlign.xlVAlignTop
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 255, 153)
Next K
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YNOTPAY0 " _
     & " where NOTPAYSEQ = 0" _
     & " order by NOTPAYISO"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYNOTPAY0_GetBuffer(rsSab, xYNOTPAY0)
    For K = 1 To arrPays_NB
        If xYNOTPAY0.NOTPAYISO = arrPAYS_ISO(K) Then xYNOTPAY0.NOTPAYLIB = arrPays_Lib(K)
    Next K

    Nb = Nb + 1
        wsExcel.Cells(Nb, 1) = xYNOTPAY0.NOTPAYISO
        wsExcel.Cells(Nb, 2) = xYNOTPAY0.NOTPAYLIB
        wsExcel.Cells(Nb, 3) = xYNOTPAY0.NOTPAYCOFA
        wsExcel.Cells(Nb, 4) = xYNOTPAY0.NOTPAYOCDE
        wsExcel.Cells(Nb, 5) = xYNOTPAY0.NOTPAYSP
        If xYNOTPAY0.NOTPAYCEG <> 0 Then wsExcel.Cells(Nb, 6) = xYNOTPAY0.NOTPAYCEG
        wsExcel.Cells(Nb, 7) = xYNOTPAY0.NOTPAYBIAN
        wsExcel.Cells(Nb, 8) = xYNOTPAY0.NOTPAYTAUX
        wsExcel.Cells(Nb, 9) = dateImp10_S(rsSab("NOTPAYBIAD"))
        wsExcel.Cells(Nb, 10) = xYNOTPAY0.NOTPAYFISC
        wsExcel.Cells(Nb, 11) = xYNOTPAY0.NOTPAYTXT
    
    rsSab.MoveNext
Loop

Call lstErr_ChangeLastItem(lstErr, cmdContext, "Exportation en cours : " & Nb & " enregistrements"): DoEvents
Set rsSab = Nothing

For K = 2 To Nb
    wsExcel.Cells(K, 3).Interior.Color = RGB(200, 255, 200)
    wsExcel.Cells(K, 4).Interior.Color = RGB(200, 255, 200)
    wsExcel.Cells(K, 5).Interior.Color = RGB(200, 255, 200)
    wsExcel.Cells(K, 8).Interior.Color = RGB(255, 200, 0)
Next K

wbExcel.SaveAs wFile

wbExcel.Close

'____________________________________________________________________________________
appExcel.Quit

Set rsSab = Nothing
Set rsSabX = Nothing

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Export"
newYNOTPAYLOG.NOTPAYLOGX = "Exportation terminé : " & Nb & " enregistrements"
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

Call lstErr_AddItem(lstErr, cmdContext, newYNOTPAYLOG.NOTPAYLOGX): DoEvents

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Public Sub YNOTPAY0_Compta()
On Error GoTo Error_Handler
Dim Nb As Long, wId As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSQL As String, wFile_Log As String
Dim wAMJMin As String, WAMJMax As String
Dim X As String, K As Long, kMax As Long, K2 As Long, K3 As Long, X40 As String
Dim rsSabX As New ADODB.Recordset
Dim wTable As String, wTableX As String

Dim arrNOTPAYTAUX()  As Double

Dim cnMDB_Compta As New ADODB.Connection
Dim rsMDB_Compta As New ADODB.Recordset
On Error GoTo Exit_sub

'______________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YNOTPAY0' and BIATABK1 = 'Compta' and BIATABK2 = 'Base'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    V = rsYBIATAB0_GetBuffer(rsSab, Compta_Lien)
End If

wFile = Trim(Compta_Lien.BIATABTXT)
'______________________________________________

X = InputBox("par défaut : " & wFile _
    & vbCrLf & vbCrLf & "     =========================" _
    & vbCrLf & "     =========================", "Notation Pays : nom de la base COMPTA", wFile)
If Trim(X) = "" Then Exit Sub

wFilex = Trim(X)
If InStr(wFilex, ".mdb") = 0 Then wFilex = wFilex & ".mdb"
'______________________________________________
If wFile <> wFilex Then
    wFile = wFilex
    Old_YBIATAB0 = Compta_Lien
    New_YBIATAB0 = Compta_Lien
    New_YBIATAB0.BIATABTXT = wFilex
    Parametrage_Update
    Compta_Lien = New_YBIATAB0
End If
'_________________________________________

'______________________________________________

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YNOTPAY0' and BIATABK1 = 'Compta' and BIATABK2 = 'Table'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    V = rsYBIATAB0_GetBuffer(rsSab, Compta_Lien)
End If

wTable = Trim(Compta_Lien.BIATABTXT)
'______________________________________________

X = InputBox("par défaut : " & wTable & vbCrLf _
    & vbCrLf & "     ================================" _
    & vbCrLf & "ATTENTION :" & vbCrLf & "PAS D'ESPACE DANS LE NOM DE LA TABLE" _
    & vbCrLf & vbCrLf & " - NOM de la colonne 1 = [Code Pays]" _
     & vbCrLf & "     ================================" _
   & vbCrLf & " - colonne 1 = code ISO" _
    & vbCrLf & " - colonne 4 = libellé" _
    & vbCrLf & " - colonne 6 = taux de provision", "Notation Pays : nom de la TABLE", wTable)

If Trim(X) = "" Then Exit Sub
If InStr(Trim(X), " ") > 0 Then
    Call MsgBox("Le nom de la table ne doit contenir d'espace", vbCritical, "Notation pays")
    Exit Sub
End If
wTableX = Trim(X)
'______________________________________________
If wTable <> wTableX Then
    wTable = wTableX
    Old_YBIATAB0 = Compta_Lien
    New_YBIATAB0 = Compta_Lien
    New_YBIATAB0.BIATABTXT = wTableX
    Parametrage_Update
    Compta_Lien = New_YBIATAB0
End If
'_________________________________________
'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Compta"
newYNOTPAYLOG.NOTPAYLOGX = "fichier MDB_Compta : " & wFile & " " & wTable
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

Set cnMDB_Compta = New ADODB.Connection
cnMDB_Compta.Provider = "Microsoft.Jet.OLEDB.4.0"
cnMDB_Compta.Mode = adModeReadWrite

cnMDB_Compta.Open wFilex

Nb = 0
X = "select count(*) as Tally   from " & wTableX
Set rsMDB_Compta = cnMDB_Compta.Execute(X)

ReDim selYNOTPAY0(rsMDB_Compta(0) + 1)
ReDim arrNOTPAYTAUX(rsMDB_Compta(0) + 1)

X = "select Champ1 , Champ4 , Champ6 from " & wTableX & " order by Champ1"
X = "select * from " & wTableX '''& " order by 'Code Pays'"
 
Set rsMDB_Compta = cnMDB_Compta.Execute(X)
Do Until rsMDB_Compta.EOF
    Nb = Nb + 1
    selYNOTPAY0(Nb).NOTPAYISO = rsMDB_Compta(0)
    selYNOTPAY0(Nb).NOTPAYLIB = rsMDB_Compta(4)
    X = rsMDB_Compta(5)
    If IsNumeric(X) Then
        selYNOTPAY0(Nb).NOTPAYTAUX = Round(X * 100, 2)
    Else
        selYNOTPAY0(Nb).NOTPAYTAUX = 0
    End If
    
    selYNOTPAY0(Nb).NOTPAYHAMJ = 0
    arrNOTPAYTAUX(Nb) = 0
    For K = 1 To arrYNOTPAY0_Nb
        If selYNOTPAY0(Nb).NOTPAYISO = arrYNOTPAY0(K).NOTPAYISO Then
            If Trim(arrYNOTPAY0(K).NOTPAYBIAN) <> "" Then
                selYNOTPAY0(Nb).NOTPAYHAMJ = DSys
                arrNOTPAYTAUX(Nb) = arrYNOTPAY0(K).NOTPAYTAUX
            End If
            Exit For
        End If
    Next K
    rsMDB_Compta.MoveNext
Loop
'_______________________________________________________________________________
''cnMDB_Compta.Close
'_______________________________________________________________________________
wFile_Log = wFilex & " " & DSYS_Time & ".log"
Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile_Log): DoEvents
Call FEU_ROUGE
Open wFile_Log For Output As #4
Print #4, "Contrôle Notation Pays Comptabilité / YNOTPAY0 "
Print #4, "================================================================="
Print #4, "fichier log : " & wFile_Log
Print #4, "-----------------------------------------------------------------"
Print #4, ""
For K = 1 To Nb
    If selYNOTPAY0(K).NOTPAYHAMJ <> 0 Then
        If selYNOTPAY0(K).NOTPAYTAUX <> arrNOTPAYTAUX(K) Then
            X40 = selYNOTPAY0(K).NOTPAYISO & " ; " & selYNOTPAY0(K).NOTPAYLIB
            Print #4, X40 & Space$(40 - Len(X40)) & " ; # ;" & Format$(selYNOTPAY0(K).NOTPAYTAUX, "##0.00") & " %" & " ; > ;" & Format$(arrNOTPAYTAUX(K), "##0.00") & " %"
        End If
    End If
Next K

Print #4, "-----------------------------------------------------------------"
Print #4, ""
For K = 1 To Nb
    If selYNOTPAY0(K).NOTPAYHAMJ <> 0 Then
        If selYNOTPAY0(K).NOTPAYTAUX = arrNOTPAYTAUX(K) Then
            X40 = selYNOTPAY0(K).NOTPAYISO & " ; " & selYNOTPAY0(K).NOTPAYLIB
            Print #4, X40 & Space$(40 - Len(X40)) & " ; = ;" & Format$(selYNOTPAY0(K).NOTPAYTAUX, "##0.00") & " %"
        End If
    End If
Next K
Print #4, "-----------------------------------------------------------------"
Print #4, ""
For K = 1 To Nb
    If selYNOTPAY0(K).NOTPAYHAMJ = 0 Then
        X40 = selYNOTPAY0(K).NOTPAYISO & " ; " & selYNOTPAY0(K).NOTPAYLIB
        Print #4, X40 & Space$(40 - Len(X40)) & " ; ? ;" & Format$(selYNOTPAY0(K).NOTPAYTAUX, "##0.00") & " %"
    End If
Next K



Print #4, ""
Print #4, "-----------------------------------------------------------------"

Print #4, ""
Print #4, "-----------------------------------------------------------------"
Print #4, "Nb traités ; " & " / " & Nb
Print #4, ""
Print #4, "================================================================="

Close #4
Call FEU_VERT

Call cmdSelect_Import_Log(wFile_Log)

X = MsgBox("Voulez-vous valider ces modifications de la base ?" _
    & vbCrLf & "si OUI, par prudence, faîtes une copie de la base avant mise à jour", vbQuestion & vbYesNo, "Mise à jour " & wTableX)

If X <> vbYes Then
    GoTo Exit_sub
End If
'________________________________________________________________________________
On Error GoTo Error_Handler
'cnMDB_Compta.Mode = adLockOptimistic
'cnMDB_Compta.Open wFilex

For K = 1 To Nb
    If selYNOTPAY0(K).NOTPAYHAMJ <> 0 Then
        If selYNOTPAY0(K).NOTPAYTAUX <> arrNOTPAYTAUX(K) Then
            X = "select * from " & wTableX & " where [Code Pays] = '" & selYNOTPAY0(K).NOTPAYISO & "'"
            If rsMDB_Compta.State = adStateOpen Then rsMDB_Compta.Close
            
            rsMDB_Compta.Open X, cnMDB_Compta, , adLockOptimistic
            rsMDB_Compta(5) = arrNOTPAYTAUX(K) / 100
            rsMDB_Compta.Update

        End If
    End If
Next K

'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Compta"
newYNOTPAYLOG.NOTPAYLOGX = "Compta terminé : " & Nb & " enregistrements"
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

Call lstErr_AddItem(lstErr, cmdContext, newYNOTPAYLOG.NOTPAYLOGX): DoEvents

GoTo Exit_sub
'_____________________________

Error_Handler:
    MsgBox Error, vbCritical, Me.Name
'_____________________________
Exit_sub:
'========
On Error Resume Next
Close #4
Call FEU_VERT
cnMDB_Compta.Close
Set cnMDB_Compta = Nothing


End Sub


Public Sub cmdSelect_Import()
Dim X As String, xSQL As String

Call arrYNOTPAY0_SQL(" where NOTPAYSEQ =  -1 order by NOTPAYISO")
Call cmdSelect_Import_Control(" ")
If arrYNOTPAY0_Nb > 0 Then
    ReDim selYNOTPAY0(arrYNOTPAY0_Nb + 50)
    selYNOTPAY0_Nb = 0: selYNOTPAY0_Max = arrYNOTPAY0_Nb + 50
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YNOTPAY0 where NOTPAYSEQ = 0 order by NOTPAYISO"
    Set rsSab = cnsab.Execute(xSQL)
    
    Do While Not rsSab.EOF
        V = rsYNOTPAY0_GetBuffer(rsSab, xYNOTPAY0)
        selYNOTPAY0_Nb = selYNOTPAY0_Nb + 1
        If selYNOTPAY0_Nb > selYNOTPAY0_Max Then
            selYNOTPAY0_Max = selYNOTPAY0_Max + 100
            ReDim Preserve selYNOTPAY0(selYNOTPAY0_Max)
        End If
            
        selYNOTPAY0(selYNOTPAY0_Nb) = xYNOTPAY0
        rsSab.MoveNext
    Loop


    fgSelect_Display
    fraSelect_Import.Visible = True
Else
    X = MsgBox("Voulez-vous préparer le processus d'importation ?", vbQuestion + vbYesNo, "Importation des notations COFACE, OCDE, S & P")
    If X = vbYes Then
        cmdSelect_Import_init
        Call arrYNOTPAY0_SQL(" where NOTPAYSEQ =  -1 order by NOTPAYISO , NOTPAYSEQ desc")
        fgSelect_Display
        fraSelect_Import.Visible = True
    End If
End If
End Sub

Public Sub cmdSelect_Import_init()

Dim K As Integer

Call arrYNOTPAY0_SQL(" where NOTPAYSEQ =  0 order by NOTPAYISO , NOTPAYSEQ desc")

If Not cmdSelect_Import_Control("Init") Then Exit Sub

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________

For K = 1 To arrYNOTPAY0_Nb
    arrYNOTPAY0(K).NOTPAYSEQ = -1
    arrYNOTPAY0(K).NOTPAYHAMJ = 0
    arrYNOTPAY0(K).NOTPAYCOFA = ""
    arrYNOTPAY0(K).NOTPAYCOFK = ""
    arrYNOTPAY0(K).NOTPAYCOFD = 0
    arrYNOTPAY0(K).NOTPAYCOF2 = ""
    
    arrYNOTPAY0(K).NOTPAYOCDE = ""
    arrYNOTPAY0(K).NOTPAYOCDK = ""
    arrYNOTPAY0(K).NOTPAYOCDD = 0
    
    arrYNOTPAY0(K).NOTPAYSP = ""
    arrYNOTPAY0(K).NOTPAYSPK = ""
    arrYNOTPAY0(K).NOTPAYSPD = 0

    ''arrYNOTPAY0(K).NOTPAYBIAN = ""
    ''arrYNOTPAY0(K).NOTPAYBIAK = ""
    ''arrYNOTPAY0(K).NOTPAYBIAD = 0

     ''arrYNOTPAY0(K).NOTPAYTAUX = 0
     arrYNOTPAY0(K).NOTPAYTXT = ""
  
    
    
     V = sqlYNOTPAY0_Insert(arrYNOTPAY0(K))

    If Not IsNull(V) Then GoTo Error_MsgBox
Next K


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, Me.Name & " ~ " & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
        Call cmdSelect_Import_DateNotation_Update("BIA")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'++++++++++++++++++++++++++++++++++++++++++
newYNOTPAYLOG.NOTPAYLOGK = "Import INI"
newYNOTPAYLOG.NOTPAYLOGX = "Importation : préparation " & arrYNOTPAY0_Nb & " enregistrements"
cmdYNOTPAYLOG_New
'++++++++++++++++++++++++++++++++++++++++++

End Sub

Public Sub cmdSelect_Import_Log(wFile_Log As String)
Dim xIn As String
lstPays.Clear
lstPays.AddItem "                     ***************"
lstPays.AddItem "                     ESC pour fermer"
lstPays.AddItem "                     ***************"
Open wFile_Log For Input As #4

Do Until EOF(4)
    Line Input #4, xIn
    lstPays.AddItem xIn
Loop
Close #4
lstPays.Visible = True

End Sub

Public Sub Parametrage_Update()
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
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Sub

Public Sub cmdSelect_SAB_Client()
Dim K As Integer, X As String, Nb As Long, xSQL As String
Dim blnOk As Boolean

param_Init_SAB_CATEG

arrYNOTPAY0_SQL " where NOTPAYSEQ =  0 order by NOTPAYISO "

fgSAB_Client.Visible = False
fgSAB_Client_Detail.Visible = False

fgSAB_Client.Rows = 1

xSQL = "select count(*) , CLIENARSD  from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
     & " where CLIENACLI >= '0010000' and CLIENACLI <= '00999999' " _
     & " group by CLIENARSD order by CLIENARSD"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    X = Trim(rsSab("CLIENARSD"))
    Nb = rsSab(0)
    
    fgSAB_Client.Rows = fgSAB_Client.Rows + 1
    fgSAB_Client.Row = fgSAB_Client.Rows - 1
    
    fgSAB_Client.Col = 3: fgSAB_Client.Text = Nb
    
    For K = 1 To arrPays_NB
        If X = arrPAYS_ISO(K) Then Exit For
    Next K
    fgSAB_Client.Col = 0: fgSAB_Client.Text = X & " - " & arrPays_Lib(K)

    blnOk = False
    For K = 1 To arrYNOTPAY0_Nb
        If X = arrYNOTPAY0(K).NOTPAYISO Then blnOk = True: Exit For
    Next K
    If blnOk Then
        fgSAB_Client.CellBackColor = RGB(255, 255, 230)
        fgSAB_Client.Col = 1: fgSAB_Client.Text = arrYNOTPAY0(K).NOTPAYBIAN
        If arrYNOTPAY0(K).NOTPAYTAUX <> 0 Then fgSAB_Client.Col = 2: fgSAB_Client.Text = Format(arrYNOTPAY0(K).NOTPAYTAUX, "##0.00")
        fgSAB_Client.Col = 1: fgSAB_Client.CellBackColor = RGB(255, 255, 230)
        fgSAB_Client.Col = 2: fgSAB_Client.CellBackColor = RGB(255, 255, 230)
        fgSAB_Client.Col = 3: fgSAB_Client.CellBackColor = RGB(255, 255, 230)
    Else
        fgSAB_Client.CellBackColor = RGB(255, 230, 230)
    End If
    
    blnOk = False
    For K = 1 To arrSAB_CATEG_Nb
        If X = arrSAB_CATEG_ISO(K) Then blnOk = True: Exit For
    Next K
    fgSAB_Client.Col = 4
    If blnOk Then
        fgSAB_Client.CellBackColor = RGB(255, 255, 230)
        fgSAB_Client.Text = arrSAB_CATEG_Code(K)
    Else
        fgSAB_Client.CellBackColor = vbMagenta 'RGB(255, 210, 210)
    End If
    
    rsSab.MoveNext
Loop

Set rsSab = Nothing
fgSAB_Client.Visible = True

End Sub
Public Sub cmdSelect_SAB_ZBALTAB0()
Dim K As Integer, X As String, Nb As Long, xSQL As String
Dim blnOk As Boolean
Dim wFile_Log As String
Dim K2 As Integer, xSP As String

param_Init_SAB_CATEG

wFile_Log = "C:\Temp\Notation_Pays_" & DSYS_Time & ".log"
Call lstErr_AddItem(lstErr, cmdContext, "- " & wFile_Log): DoEvents
Call FEU_ROUGE

Open wFile_Log For Output As #4
Print #4, "================================================================="
Print #4, "fichier log : " & wFile_Log
Print #4, "-----------------------------------------------------------------"
Print #4, ""
Print #4, "Contrôle pays de résidence des clients / catégorie SAB (ZBALTAB0) : "
Print #4, "================================================================="


xSQL = "select count(*) , CLIENARSD  from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
     & " where CLIENACLI >= '0010000' and CLIENACLI <= '00999999' " _
     & " group by CLIENARSD order by CLIENARSD"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    X = Trim(rsSab("CLIENARSD"))
    
    blnOk = False
    For K = 1 To arrSAB_CATEG_Nb
        If X = arrSAB_CATEG_ISO(K) Then blnOk = True: Exit For
    Next K
    If Not blnOk Then
            
        For K = 1 To arrPays_NB
            If X = arrPAYS_ISO(K) Then Exit For
        Next K

        Print #4, "pays de résidence non paramétré " & X & " - " & arrPays_Lib(K)
    End If
    

    rsSab.MoveNext
Loop
Set rsSab = Nothing
Print #4, ""
Print #4, "Contrôle catégorie SAB (ZBALTAB0) / notation S & P : "
Print #4, "================================================================="


arrYNOTPAY0_SQL " where NOTPAYSEQ =  0 order by NOTPAYISO , NOTPAYSEQ asc"

For K = 1 To arrSAB_CATEG_Nb
    X = arrSAB_CATEG_ISO(K)
    blnOk = False
    For K2 = 1 To arrYNOTPAY0_Nb
        If X = arrYNOTPAY0(K2).NOTPAYISO Then blnOk = True: xSP = arrYNOTPAY0(K2).NOTPAYSP: Exit For
    Next K2
    
    If Not blnOk Then
        Print #4, X & " : code ISO SAB inconnu dans l'application Notation Pays"
    Else
        blnOk = False
        For K2 = 1 To arrYNOTPAY0_Nb
            If xSP = arrSP(K2).NOTPAYSP Then blnOk = True: Exit For
        Next K2
        If Not blnOk Then
            Print #4, X & " : pas de correspondance" & xSP
        Else
            If arrSAB_CATEG_Code(K) <> Mid$(arrSP(K2).NOTPAYTXT, 15, 6) Then
                Print #4, X & " : catégorie SAB  " & arrSAB_CATEG_Code(K) & " # " & Mid$(arrSP(K2).NOTPAYTXT, 3, 19)
            End If
        End If
    End If
    

    
Next K


Print #4, ""
Print #4, "-----------------------------------------------------------------"

Print #4, "================================================================="


Close #4
Call FEU_VERT

Call lstErr_AddItem(lstErr, cmdContext, "- fin"): DoEvents

Call cmdSelect_Import_Log(wFile_Log)


End Sub

Public Sub fgSAB_Client_Detail_Display(lISO As String)
Dim xSQL As String, X As String
fgSAB_Client_Detail.Visible = False

fgSAB_Client_Detail.Rows = 1

xSQL = "select CLIENACLI,CLIENARA1,CLIENARA2,CLIENARSD  from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
     & " where CLIENARSD = '" & Mid$(lISO, 1, 2) _
     & "' and CLIENACLI >= '0010000' and CLIENACLI <= '00999999' "
     
     Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    fgSAB_Client_Detail.Rows = fgSAB_Client_Detail.Rows + 1
    fgSAB_Client_Detail.Row = fgSAB_Client_Detail.Rows - 1
    
    fgSAB_Client_Detail.Col = 0: fgSAB_Client_Detail.Text = lISO
    fgSAB_Client_Detail.Col = 1: fgSAB_Client_Detail.Text = rsSab("CLIENACLI")
    fgSAB_Client_Detail.CellFontBold = True
    fgSAB_Client_Detail.CellAlignment = 4
    X = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
    fgSAB_Client_Detail.Col = 2: fgSAB_Client_Detail.Text = X
    'If InStr(X, "CLOS ") > 0 Then
    If retourne_Client_CLOS(rsSab("CLIENACLI"), X) Then
        fgSAB_Client_Detail.CellBackColor = RGB(255, 230, 230)
    End If

    rsSab.MoveNext
Loop
Set rsSab = Nothing
fgSAB_Client_Detail.Visible = True

End Sub

Public Sub cmdYNOTPAYLOG_New()
Dim V

App_Debug = "cmdYNOTPAYLOG_New"
newYNOTPAYLOG.NOTPAYLOGH = time_Hms            ' heure maj
newYNOTPAYLOG.NOTPAYLOGS = newYNOTPAYLOG.NOTPAYLOGS + 1

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
 V = sqlYNOTPAYLOG_Insert(newYNOTPAYLOG)

'________________________________________________________________________________

If Not IsNull(V) Then GoTo Error_MsgBox
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, Me.Name & " ~ " & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Sub

Public Function cmdSelect_Import_Control(lFct As String) As Boolean
Dim xSQL As String, V
Dim blnSelect_Import_Validation As Boolean

cmdSelect_Import_Control = False
'==============================
rsYBIATAB0_Init Old_YBIATAB0
Coface_Notepad = Old_YBIATAB0
OCDE_Notepad = Old_YBIATAB0
SP_Notepad = Old_YBIATAB0
Coface_DateNotation = Old_YBIATAB0
OCDE_DateNotation = Old_YBIATAB0
SP_DateNotation = Old_YBIATAB0

'_____________________________________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YNOTPAY0' and BIATABK1 = 'Notepad' order by BIATABK2"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYBIATAB0_GetBuffer(rsSab, Old_YBIATAB0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYNOTPAY0.param_init"
        '' Exit Sub
     Else
        Select Case Trim(Old_YBIATAB0.BIATABK2)
            Case "Coface": Coface_Notepad = Old_YBIATAB0
            Case "OCDE": OCDE_Notepad = Old_YBIATAB0
            Case "SP": SP_Notepad = Old_YBIATAB0

        End Select
    End If
    rsSab.MoveNext
Loop
'_____________________________________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'YNOTPAY0' and BIATABK1 = 'DateNotation' order by BIATABK2"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYBIATAB0_GetBuffer(rsSab, Old_YBIATAB0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYNOTPAY0.param_init"
        '' Exit Sub
     Else
        Select Case Trim(Old_YBIATAB0.BIATABK2)
            Case "Coface": Coface_DateNotation = Old_YBIATAB0
                           Coface_DateNotation_Info = "Dernière notation : " & dateImp10(Mid$(Old_YBIATAB0.BIATABTXT, 1, 8)) & vbCrLf _
                                                    & "par " & Mid$(Old_YBIATAB0.BIATABTXT, 10, 10) & " le " _
                                                    & dateImp10(Mid$(Old_YBIATAB0.BIATABTXT, 21, 8)) & "  " & timeImp8(Mid$(Old_YBIATAB0.BIATABTXT, 30, 6))
            Case "OCDE": OCDE_DateNotation = Old_YBIATAB0
                           OCDE_DateNotation_Info = "Dernière notation : " & dateImp10(Mid$(Old_YBIATAB0.BIATABTXT, 1, 8)) & vbCrLf _
                                                    & "par " & Mid$(Old_YBIATAB0.BIATABTXT, 9, 10) & " le " _
                                                    & dateImp10(Mid$(Old_YBIATAB0.BIATABTXT, 21, 8)) & "  " & timeImp8(Mid$(Old_YBIATAB0.BIATABTXT, 30, 6))
            Case "SP": SP_DateNotation = Old_YBIATAB0
                           SP_DateNotation_Info = "Dernière notation : " & dateImp10(Mid$(Old_YBIATAB0.BIATABTXT, 1, 8)) & vbCrLf _
                                                    & "par " & Mid$(Old_YBIATAB0.BIATABTXT, 9, 10) & " le " _
                                                    & dateImp10(Mid$(Old_YBIATAB0.BIATABTXT, 21, 8)) & "  " & timeImp8(Mid$(Old_YBIATAB0.BIATABTXT, 30, 6))

             Case "BIA": BIA_DateNotation = Old_YBIATAB0
                           BIA_DateNotation_Info = "Dernière notation : " & dateImp10(Mid$(Old_YBIATAB0.BIATABTXT, 1, 8)) & vbCrLf _
                                                    & "par " & Mid$(Old_YBIATAB0.BIATABTXT, 9, 10) & " le " _
                                                    & dateImp10(Mid$(Old_YBIATAB0.BIATABTXT, 21, 8)) & "  " & timeImp8(Mid$(Old_YBIATAB0.BIATABTXT, 30, 6))
       End Select
    End If
    rsSab.MoveNext
Loop

'__________________________________________________________________________
Call DTPicker_Control(txtSelect_Import_Amj, wAMJMin)
If wAMJMin > DSys Then
    X = MsgBox(dateImp10(wAMJMin) & " est supérieure à la date du jour " & dateImp10(DSys), vbCritical, "Importation " & lFct)
    Exit Function
End If


Select Case lFct
    Case "Coface":
        If wAMJMin < Mid$(Coface_DateNotation.BIATABTXT, 1, 8) Then
            X = MsgBox(dateImp10(wAMJMin) & " est inférieure à la date précédente" & vbCrLf & vbCrLf _
             & Coface_DateNotation_Info & vbCrLf _
             & "Validez-vous cette date ?", vbYesNo + vbQuestion, "Importation " & lFct)
            If X = vbYes Then cmdSelect_Import_Control = True
        Else
            cmdSelect_Import_Control = True
        End If
     Case "OCDE":
        If wAMJMin < Mid$(OCDE_DateNotation.BIATABTXT, 1, 8) Then
            X = MsgBox(dateImp10(wAMJMin) & " est inférieure à la date précédente" & vbCrLf & vbCrLf _
             & OCDE_DateNotation_Info & vbCrLf _
             & "Validez-vous cette date ?", vbYesNo + vbQuestion, "Importation " & lFct)
            If X = vbYes Then cmdSelect_Import_Control = True
        Else
            cmdSelect_Import_Control = True
        End If
    Case "SP":
        If wAMJMin < Mid$(SP_DateNotation.BIATABTXT, 1, 8) Then
            X = MsgBox(dateImp10(wAMJMin) & " est inférieure à la date précédente" & vbCrLf & vbCrLf _
             & SP_DateNotation_Info & vbCrLf _
             & "Validez-vous cette date ?", vbYesNo + vbQuestion, "Importation " & lFct)
            If X = vbYes Then cmdSelect_Import_Control = True
        Else
            cmdSelect_Import_Control = True
        End If
        
    Case "Init": cmdSelect_Import_Control = True
     Case "BIA":
        If wAMJMin < Mid$(BIA_DateNotation.BIATABTXT, 1, 8) Then
            X = MsgBox(dateImp10(wAMJMin) & " est inférieure à la date précédente" & vbCrLf & vbCrLf _
             & BIA_DateNotation_Info & vbCrLf _
             & "Validez-vous cette date ?", vbYesNo + vbQuestion, "Importation " & lFct)
            If X = vbYes Then cmdSelect_Import_Control = True
        Else
            cmdSelect_Import_Control = True
        End If
           
    
End Select

blnSelect_Import_Validation = True
If Mid$(Coface_DateNotation.BIATABTXT, 21, 15) > Mid$(BIA_DateNotation.BIATABTXT, 21, 15) Then
    Call lstErr_AddItem(lstErr, cmdContext, "- COFACE " & Coface_DateNotation_Info): DoEvents
Else
    blnSelect_Import_Validation = False
    Call lstErr_AddItem(lstErr, cmdContext, "! Importation COFACE à faire"): DoEvents
End If

If Mid$(OCDE_DateNotation.BIATABTXT, 21, 15) > Mid$(BIA_DateNotation.BIATABTXT, 21, 15) Then
    Call lstErr_AddItem(lstErr, cmdContext, "- OCDE " & OCDE_DateNotation_Info): DoEvents
Else
    blnSelect_Import_Validation = False
    Call lstErr_AddItem(lstErr, cmdContext, "! Importation OCDE à faire"): DoEvents
End If

If Mid$(SP_DateNotation.BIATABTXT, 21, 15) > Mid$(BIA_DateNotation.BIATABTXT, 21, 15) Then
    Call lstErr_AddItem(lstErr, cmdContext, "- SP " & SP_DateNotation_Info): DoEvents
Else
    blnSelect_Import_Validation = False
    Call lstErr_AddItem(lstErr, cmdContext, "! Importation SP à faire"): DoEvents
End If

cmdSelect_Import_Validation.Visible = blnSelect_Import_Validation

End Function

Public Sub cmdSelect_Import_DateNotation_Update(lFct As String)

Select Case lFct
    Case "Coface": Old_YBIATAB0 = Coface_DateNotation
    Case "OCDE": Old_YBIATAB0 = OCDE_DateNotation
    Case "SP": Old_YBIATAB0 = SP_DateNotation
    Case "BIA": Old_YBIATAB0 = BIA_DateNotation
End Select

New_YBIATAB0 = Old_YBIATAB0
Mid$(New_YBIATAB0.BIATABTXT, 1, 8) = wAMJMin
Mid$(New_YBIATAB0.BIATABTXT, 10, 10) = usrName_UCase       ' utilisateur maj
Mid$(New_YBIATAB0.BIATABTXT, 21, 8) = DSys              ' DATE maj
Mid$(New_YBIATAB0.BIATABTXT, 30, 6) = time_Hms            ' heure maj
Call Parametrage_Update
Call cmdSelect_Import_Control(" ")
End Sub

Public Sub fraParam_Display()
Dim K As Integer, X As String
Dim xWhere As String, xSQL As String
Dim Nb As Integer, wText As String
Dim X40 As String, wNOTPAYBIAK As String
txtParam_Code.Enabled = False
txtParam_Taux.Enabled = YNOTPAY0_Aut.Valider

If Trim(oldYNOTPAY0.NOTPAYCOFA) <> "" Then
    lblParam_Code = "1-COFACE"
    txtParam_Code = Trim(oldYNOTPAY0.NOTPAYCOFA)
    xWhere = " and NOTPAYCOFA = '" & oldYNOTPAY0.NOTPAYCOFA & "'"
    wText = "COFACE : " & oldYNOTPAY0.NOTPAYCOFA
End If
If Trim(oldYNOTPAY0.NOTPAYOCDE) <> "" Then
    lblParam_Code = "2-OCDE"
    txtParam_Code = Trim(oldYNOTPAY0.NOTPAYOCDE)
    xWhere = " and NOTPAYOCDE = '" & oldYNOTPAY0.NOTPAYOCDE & "'"
    wText = "OCDE : " & oldYNOTPAY0.NOTPAYOCDE
End If
If Trim(oldYNOTPAY0.NOTPAYSP) <> "" Then
    lblParam_Code = "3-S&P"
    txtParam_Code = Trim(oldYNOTPAY0.NOTPAYSP)
    xWhere = " and NOTPAYSP = '" & oldYNOTPAY0.NOTPAYSP & "'"
    wText = "SPCE : " & oldYNOTPAY0.NOTPAYSP
End If
If Trim(oldYNOTPAY0.NOTPAYBIAN) <> "" Then
    lblParam_Code = "4-BIA"
    txtParam_Code = Trim(oldYNOTPAY0.NOTPAYBIAN)
    xWhere = " and NOTPAYBIAN = '" & oldYNOTPAY0.NOTPAYBIAN & "'"
    wText = "BIA : " & oldYNOTPAY0.NOTPAYBIAN
End If

fraParam.Caption = Trim(oldYNOTPAY0.NOTPAYTXT)
txtParam_Taux = Format$(oldYNOTPAY0.NOTPAYTAUX, "##0.00")
fraParam.Visible = True

'_________________________________________________________
lstPays.Clear
Nb = 0
lstPays.AddItem "   - Liste des pays concernés par la notation " & wText
lstPays.AddItem "   - =========================================================="
lstPays.AddItem "   "

X = "select NOTPAYISO,NOTPAYBIAN, NOTPAYTAUX,NOTPAYBIAK from " & paramIBM_Library_SABSPE & ".YNOTPAY0" _
  & "  where NOTPAYSEQ = 0 " & xWhere
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    Nb = Nb + 1
    X = rsSab("NOTPAYISO")
    For K = 1 To arrPays_NB
        If X = arrPAYS_ISO(K) Then Exit For
    Next K
    If rsSab("NOTPAYBIAK") = "M" Then
        wNOTPAYBIAK = "M !!!"
    Else
        wNOTPAYBIAK = "     "
    End If
    X40 = X & " - " & arrPays_Lib(K)
    lstPays.AddItem X40 & Space$(40 - Len(X40)) & wNOTPAYBIAK & rsSab("NOTPAYBIAN") & " = " & Format$(rsSab("NOTPAYTAUX"), "##0.00") & " %"
    rsSab.MoveNext
Loop
lstPays.AddItem "   "
lstPays.AddItem "   - =========================================================="
lstPays.AddItem "   -  " & Nb & " pays concernés"

lstPays.Visible = True

If Nb = 0 Then
    cmdParam_Delete.Enabled = True
Else
    cmdParam_Delete.Enabled = False
End If

End Sub

Public Sub Param_Load_Pondération()
'_____________________________________________________________________________

Dim X As String, K As Integer
Dim xSQL As String

arrYNOTPAY0_SQL " where NOTPAYISO =  '$$' order by NOTPAYTXT"
arrCoface_Nb = 0
arrOCDE_Nb = 0
arrSP_Nb = 0
arrBIAN_Nb = 0
For K = 1 To arrYNOTPAY0_Nb
    Select Case Mid$(arrYNOTPAY0(K).NOTPAYTXT, 1, 1)
        Case "1": arrCoface_Nb = arrCoface_Nb + 1
        Case "2": arrOCDE_Nb = arrOCDE_Nb + 1
        Case "3": arrSP_Nb = arrSP_Nb + 1
        Case "4": arrBIAN_Nb = arrBIAN_Nb + 1
    End Select
Next K
ReDim arrCoface(arrCoface_Nb + 1)
ReDim arrOCDE(arrOCDE_Nb + 1)
ReDim arrSP(arrSP_Nb + 1)
ReDim arrBIAN(arrBIAN_Nb + 1)
arrCoface_Nb = 0
arrOCDE_Nb = 0
arrSP_Nb = 0
arrBIAN_Nb = 0
cboDetail_NOTPAYCOFA.Clear
cboDetail_NOTPAYOCDE.Clear
cboDetail_NOTPAYSP.Clear
cboDetail_NOTPAYBIAN.Clear
For K = 1 To arrYNOTPAY0_Nb
    Select Case Mid$(arrYNOTPAY0(K).NOTPAYTXT, 1, 1)
        Case "1": arrCoface_Nb = arrCoface_Nb + 1: arrCoface(arrCoface_Nb) = arrYNOTPAY0(K)
                    cboDetail_NOTPAYCOFA.AddItem arrCoface(arrCoface_Nb).NOTPAYCOFA
                    cboDetail_NOTPAYCOF2.AddItem arrCoface(arrCoface_Nb).NOTPAYCOFA
        Case "2": arrOCDE_Nb = arrOCDE_Nb + 1: arrOCDE(arrOCDE_Nb) = arrYNOTPAY0(K)
                    cboDetail_NOTPAYOCDE.AddItem arrOCDE(arrOCDE_Nb).NOTPAYOCDE
        Case "3": arrSP_Nb = arrSP_Nb + 1: arrSP(arrSP_Nb) = arrYNOTPAY0(K)
                    cboDetail_NOTPAYSP.AddItem arrSP(arrSP_Nb).NOTPAYSP
        Case "4": arrBIAN_Nb = arrBIAN_Nb + 1: arrBIAN(arrBIAN_Nb) = arrYNOTPAY0(K)
                    cboDetail_NOTPAYBIAN.AddItem arrBIAN(arrBIAN_Nb).NOTPAYBIAN
    End Select
Next K

End Sub

Public Sub param_Init_SAB_CATEG()
Dim X As String, K As Integer
Dim xSQL As String

X = "select count(*) as Tally   from " & paramIBM_Library_SAB & ".ZBALTAB0 " _
    & " where BALTABNUM = 12 "
Set rsSab = cnsab.Execute(X)

ReDim arrSAB_CATEG_ISO(rsSab("Tally") + 1)
ReDim arrSAB_CATEG_Code(rsSab("Tally") + 1)
arrSAB_CATEG_Nb = 0

txtSelect_NOTPAYISO.Clear
txtSelect_NOTPAYISO.AddItem "   "
X = "select * from " & paramIBM_Library_SAB & ".ZBALTAB0 " _
    & " where BALTABNUM = 12 order by BALTABARG "
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    arrSAB_CATEG_Nb = arrSAB_CATEG_Nb + 1
    arrSAB_CATEG_ISO(arrSAB_CATEG_Nb) = Trim(rsSab("BALTABARG"))
    arrSAB_CATEG_Code(arrSAB_CATEG_Nb) = Trim(rsSab("BALTABLO1"))
    
    rsSab.MoveNext
Loop

End Sub

Public Sub cmdSelect_Export_Mail()
Dim X As String, I As Integer
Dim wSendMail As typeSendMail
Dim xHeader As String, xDétail As String, mbgColor As String
Dim xNOTPAYCEG As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Exportation en cours ......"): DoEvents

wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = "Notation_Pays"
X = InputBox("Confirmer l'adresse 'mail'du destinataire, SINON" & vbCrLf _
            & "saisir son adresse (<nom> <point> <initiales du prénom>)", "CONFIRMATION", "DCOM")

If Trim(X) = "" Then
    GoTo Exit_sub
Else
    If InStr(X, "@") > 0 Then
        X = Trim(X)
    Else
        X = Trim(X) & "@bia-paris.fr"
    End If
End If

wSendMail.Recipient = X

wSendMail.Subject = "Notation Pays  "
wSendMail.Attachment = ""


xHeader = "<TR>" _
         & "<TD bgcolor=#0090A0  width=300><span style='font-size:8.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF><div align=" & Asc34 & "Left" & Asc34 & ">" _
         & "Pays" & "</TD>" _
         & "<TD bgcolor=#0090A0  width=100 ><span style='font-size:8.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF><div align=" & Asc34 & "Center" & Asc34 & ">" _
         & "Coface" _
         & "<TD bgcolor=#0090A0  width=100 ><span style='font-size:8.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF><div align=" & Asc34 & "Center" & Asc34 & ">" _
         & "OCDE" & "</TD>" _
         & "<TD bgcolor=#0090A0  width=100 ><span style='font-size:8.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF><div align=" & Asc34 & "Center" & Asc34 & ">" _
         & "S & P" _
         & "<TD bgcolor=#0090A0  width=100 ><span style='font-size:8.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF><div align=" & Asc34 & "Center" & Asc34 & ">" _
         & "Pondération" & "</TD>" _
         & "<TD bgcolor=#0090A0  width=100 ><span style='font-size:8.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF><div align=" & Asc34 & "Center" & Asc34 & ">" _
         & "BIA" _
         & "<TD bgcolor=#0090A0  width=100 ><span style='font-size:8.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF><div align=" & Asc34 & "Right" & Asc34 & ">" _
         & "taux" & "</TD>" _
         & "<TD bgcolor=#0090A0  width=50 ><span style='font-size:8.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF><div align=" & Asc34 & "Right" & Asc34 & ">" _
         & "Fisc" _
         & "<TD bgcolor=#0090A0  width=50 ><span style='font-size:8.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF><div align=" & Asc34 & "Center" & Asc34 & ">" _
         & "Prov" & "</TD>" _
         & "<TD bgcolor=#0090A0  width=100 ><span style='font-size:8.0pt;font-family:Arial Unicode MS'><Font color=#FFFFFF><div align=" & Asc34 & "Center" & Asc34 & ">" _
         & "Date" & "</TD>" _
        & "</TR>"
        
xDétail = ""
mbgColor = "bgcolor = #F0F0F0"

For I = 1 To arrYNOTPAY0_Nb
         
    xYNOTPAY0 = arrYNOTPAY0(I)
    If xYNOTPAY0.NOTPAYCEG = 0 Then
        xNOTPAYCEG = ""
    Else
        xNOTPAYCEG = xYNOTPAY0.NOTPAYCEG
    End If
    xDétail = xDétail _
         & "<TR>" _
         & "<TD bgcolor = #0090A0 width=300 ><span style='font-size:7.0pt;font-family:Arial Unicode MS'><div align=" & Asc34 & "Left" & Asc34 & "><Font color=#FFFFFF>" & xYNOTPAY0.NOTPAYISO & " - " & xYNOTPAY0.NOTPAYLIB & "</TD>" _
         & "<TD " & mbgColor & " width=100 ><span style='font-size:7.0pt;font-family:Arial Unicode MS'><div align=" & Asc34 & "Center" & Asc34 & ">" & xYNOTPAY0.NOTPAYCOFA & "</TD>" _
         & "<TD " & mbgColor & " width=100 ><span style='font-size:7.0pt;font-family:Arial Unicode MS'><div align=" & Asc34 & "Center" & Asc34 & ">" & xYNOTPAY0.NOTPAYOCDE & "</TD>" _
         & "<TD " & mbgColor & " width=100 ><span style='font-size:7.0pt;font-family:Arial Unicode MS'><div align=" & Asc34 & "Center" & Asc34 & ">" & xYNOTPAY0.NOTPAYSP & "</TD>" _
         & "<TD " & mbgColor & " width=100 ><span style='font-size:7.0pt;font-family:Arial Unicode MS'><div align=" & Asc34 & "Center" & Asc34 & ">" & xNOTPAYCEG & "</TD>" _
         & "<TD " & mbgColor & " width=100 ><span style='font-size:7.0pt;font-family:Arial Unicode MS'><div align=" & Asc34 & "Center" & Asc34 & ">" & xYNOTPAY0.NOTPAYBIAN & "</TD>" _
         & "<TD bgcolor = #0090A0 width=100 ><span style='font-size:7.0pt;font-family:Arial Unicode MS'><div align=" & Asc34 & "Right" & Asc34 & "><Font color=#FFFFFF>" & Format$(xYNOTPAY0.NOTPAYTAUX, " ##0.00") & "</TD>" _
         & "<TD " & mbgColor & " width=50 ><span style='font-size:7.0pt;font-family:Arial Unicode MS'><div align=" & Asc34 & "Right" & Asc34 & ">" & xYNOTPAY0.NOTPAYFISC & "</TD>" _
         & "<TD " & mbgColor & " width=50 ><span style='font-size:7.0pt;font-family:Arial Unicode MS'><div align=" & Asc34 & "Center" & Asc34 & ">" & xYNOTPAY0.NOTPAYPROV & "</TD>" _
         & "<TD " & mbgColor & " width=100 ><span style='font-size:7.0pt;font-family:Arial Unicode MS'><div align=" & Asc34 & "Center" & Asc34 & ">" & dateImp10(xYNOTPAY0.NOTPAYBIAD) & "</TD>" _
         & "</TR>"
Next I



wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                    & "<TABLE  width=2430px border=1 cellpadding=10 ></B>" _
                    & xHeader _
                    & xDétail _
                    & "</div></TABLE>"

'                    & "<div align=" & Asc34 & "Left" & Asc34 _

wSendMail.AsHTML = True
srvSendMail.Monitor wSendMail


Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub
