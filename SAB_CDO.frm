VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAB_CDO 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_CDO : Crédits documentaires"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13560
   Icon            =   "SAB_CDO.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9150
   ScaleWidth      =   13560
   Begin VB.CommandButton cmdDocushare 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   7080
      Picture         =   "SAB_CDO.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   0
      Width           =   735
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   4
      Top             =   0
      Width           =   5175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   2
      Top             =   500
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Rechercher"
      TabPicture(0)   =   "SAB_CDO.frx":0D84
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Courrier Ouv/Mod"
      TabPicture(1)   =   "SAB_CDO.frx":0DA0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDossier"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Courrier Utilisations"
      TabPicture(2)   =   "SAB_CDO.frx":0DBC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraYCDOUTI0"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Détail fichiers"
      TabPicture(3)   =   "SAB_CDO.frx":0DD8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fgDossier"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraYCDOUTI0 
         Height          =   8175
         Left            =   -74880
         TabIndex        =   41
         Top             =   360
         Width           =   13335
         Begin VB.Frame fraYCDOUTI0_Display 
            Height          =   6495
            Left            =   120
            TabIndex        =   42
            Top             =   1680
            Width           =   13095
            Begin VB.TextBox txtYCDOUTI0_ATT 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   8760
               TabIndex        =   69
               Top             =   2040
               Width           =   4215
            End
            Begin VB.TextBox txtYCDOUTI0_DELAI 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   120
               TabIndex        =   68
               Top             =   6000
               Width           =   12855
            End
            Begin VB.Frame fraYCDOUTI0_Exp 
               Caption         =   "Expédition"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2895
               Left            =   6840
               TabIndex        =   55
               Top             =   3000
               Width           =   6255
               Begin VB.TextBox txtYCDOUTI0_Exp_Par 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   600
                  TabIndex        =   49
                  Top             =   360
                  Width           =   2415
               End
               Begin VB.TextBox txtYCDOUTI0_Exp_Le 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   3600
                  TabIndex        =   50
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.TextBox txtYCDOUTI0_Exp_De 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   600
                  TabIndex        =   51
                  Top             =   960
                  Width           =   2415
               End
               Begin VB.TextBox txtYCDOUTI0_Exp_A 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Left            =   3600
                  TabIndex        =   52
                  Top             =   945
                  Width           =   2535
               End
               Begin VB.TextBox txtYCDOUTI0_YCDODES0 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1365
                  Left            =   120
                  MultiLine       =   -1  'True
                  TabIndex        =   53
                  Top             =   1440
                  Width           =   6015
               End
               Begin VB.Label lblYCDOUTI0_Exp_Par 
                  Caption         =   "par"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   59
                  Top             =   480
                  Width           =   375
               End
               Begin VB.Label lblYCDOUTI0_Exp_Le 
                  Caption         =   "le"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   3120
                  TabIndex        =   58
                  Top             =   480
                  Width           =   375
               End
               Begin VB.Label lblYCDOUTI0_Exp_de 
                  Caption         =   "de"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   120
                  TabIndex        =   57
                  Top             =   1080
                  Width           =   495
               End
               Begin VB.Label lblYCDOUTI0_Exp_A 
                  Caption         =   "à"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   3120
                  TabIndex        =   56
                  Top             =   1080
                  Width           =   255
               End
            End
            Begin VB.Frame fraYCDOUTI0_Document 
               Height          =   6495
               Left            =   0
               TabIndex        =   45
               Top             =   0
               Width           =   6735
               Begin VB.TextBox txtYCDOUTI0_CrrGB_BqRbt 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   3840
                  TabIndex        =   77
                  Top             =   5640
                  Width           =   2655
               End
               Begin VB.TextBox txtYCDOUTI0_CrrGB_Tx 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   3840
                  TabIndex        =   75
                  Top             =   5160
                  Width           =   2655
               End
               Begin VB.TextBox txtYCDOUTI0_CrrGB_Mnt 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   3840
                  TabIndex        =   73
                  Top             =   4680
                  Width           =   2655
               End
               Begin VB.Frame fraYCDOUTI0_Document_Select 
                  BackColor       =   &H00F0FFFF&
                  Height          =   1215
                  Left            =   120
                  TabIndex        =   60
                  Top             =   1800
                  Width           =   6375
                  Begin VB.CommandButton cmdYCDOUTI0_Document_Remove 
                     BackColor       =   &H008080FF&
                     Caption         =   "Supprimer"
                     Height          =   400
                     Left            =   120
                     Style           =   1  'Graphical
                     TabIndex        =   66
                     Top             =   240
                     Width           =   1800
                  End
                  Begin VB.CommandButton cmdYCDOUTI0_Document_Quit 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "Ignorer"
                     Height          =   400
                     Left            =   2265
                     Style           =   1  'Graphical
                     TabIndex        =   65
                     Top             =   240
                     Width           =   1800
                  End
                  Begin VB.CommandButton cmdYCDOUTI0_Document_Ok 
                     BackColor       =   &H00C0FFC0&
                     Caption         =   "Ajouter"
                     Height          =   400
                     Left            =   4320
                     Style           =   1  'Graphical
                     TabIndex        =   64
                     Top             =   240
                     Width           =   1800
                  End
                  Begin VB.TextBox txtYCDOUTI0_Document_Jeu2 
                     Height          =   350
                     Left            =   5280
                     TabIndex        =   63
                     Top             =   720
                     Width           =   855
                  End
                  Begin VB.TextBox txtYCDOUTI0_Document_Jeu1 
                     Height          =   350
                     Left            =   4305
                     TabIndex        =   62
                     Top             =   735
                     Width           =   855
                  End
                  Begin VB.TextBox txtYCDOUTI0_Document 
                     Height          =   350
                     Left            =   105
                     TabIndex        =   61
                     Top             =   735
                     Width           =   4095
                  End
               End
               Begin VB.ListBox lstYCDOUTI0_Document 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1260
                  Left            =   120
                  Sorted          =   -1  'True
                  TabIndex        =   47
                  Top             =   240
                  Width           =   5535
               End
               Begin MSFlexGridLib.MSFlexGrid fgYCDOUTI0_Document_Select 
                  Height          =   1515
                  Left            =   120
                  TabIndex        =   46
                  Top             =   3120
                  Width           =   6375
                  _ExtentX        =   11245
                  _ExtentY        =   2672
                  _Version        =   393216
                  Rows            =   1
                  Cols            =   3
                  FixedCols       =   0
                  RowHeightMin    =   200
                  BackColor       =   15794175
                  ForeColor       =   4210688
                  BackColorFixed  =   8454143
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
                  FormatString    =   "Document                                             | 1er Jeu        | 2 ème Jeu     "
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin VB.Label lblYCDOUTI0_CrrGB_BqRbt 
                  Caption         =   "Nom bq remboursement"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   76
                  Top             =   5640
                  Width           =   2295
               End
               Begin VB.Label lblYCDOUTI0_CrrGB_Tx 
                  Caption         =   "Taux d'utilisation"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   74
                  Top             =   5160
                  Width           =   1815
               End
               Begin VB.Label lblYCDOUTI0_CrrGB_Mnt 
                  Caption         =   "Mnt 100% Utilisation"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   72
                  Top             =   4680
                  Width           =   1815
               End
               Begin VB.Label lblYCDOUTI0_CrrGB 
                  Caption         =   "Crr ANGLAIS"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   71
                  Top             =   4680
                  Width           =   1815
               End
            End
            Begin VB.TextBox txtYCDOUTI0_IBAN 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   7800
               TabIndex        =   48
               Top             =   2520
               Width           =   5175
            End
            Begin VB.ListBox lstYCDOUTI0_Courrier 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1410
               ItemData        =   "SAB_CDO.frx":0DF4
               Left            =   7920
               List            =   "SAB_CDO.frx":0DFB
               Style           =   1  'Checkbox
               TabIndex        =   43
               Top             =   240
               Width           =   5085
            End
            Begin VB.Label lblYCDOUTI0_ATT 
               Caption         =   "A l'attention de (AR)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6840
               TabIndex        =   70
               Top             =   2040
               Width           =   1815
            End
            Begin VB.Label lblYCDOUTI0_IBAN 
               Caption         =   "IBAN"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6840
               TabIndex        =   54
               Top             =   2520
               Width           =   855
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgYCDOUTI0 
            Height          =   1395
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   2461
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            FixedCols       =   0
            RowHeightMin    =   200
            BackColor       =   14737632
            ForeColor       =   4210688
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
            FormatString    =   $"SAB_CDO.frx":0E0B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame fraDossier 
         Height          =   8145
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   13260
         Begin VB.TextBox txtDossier_Garantie 
            Height          =   1455
            Left            =   2040
            MultiLine       =   -1  'True
            TabIndex        =   79
            Text            =   "SAB_CDO.frx":0EBC
            Top             =   6480
            Width           =   7815
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   0
            TabIndex        =   78
            Top             =   0
            Width           =   12855
         End
         Begin VB.Frame fraDossier_Info 
            Enabled         =   0   'False
            Height          =   6015
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   7095
            Begin VB.TextBox txtDossier_CDODOSBEN 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Left            =   1320
               MultiLine       =   -1  'True
               TabIndex        =   28
               Text            =   "SAB_CDO.frx":0EC2
               Top             =   3240
               Width           =   5535
            End
            Begin VB.TextBox txtDossier_CDODOSDON 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   1320
               MultiLine       =   -1  'True
               TabIndex        =   26
               Text            =   "SAB_CDO.frx":0EC8
               Top             =   2520
               Width           =   5535
            End
            Begin VB.TextBox txtDossier_CDODOSCOR 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   1320
               MultiLine       =   -1  'True
               TabIndex        =   17
               Text            =   "SAB_CDO.frx":0ECE
               Top             =   1800
               Width           =   5535
            End
            Begin VB.Label lblDossier_CDODOSMDI 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "P.Dif"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   3000
               TabIndex        =   36
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lblDossier_CDODOSMOV 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "A vue"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   3000
               TabIndex        =   35
               Top             =   840
               Width           =   975
            End
            Begin VB.Label txtDossier_CDODOSMDI 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Mnt PDIF"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   4440
               TabIndex        =   34
               Top             =   1320
               Width           =   2295
            End
            Begin VB.Label txtDossier_CDODOSMOV 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Mnt AVUE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   4440
               TabIndex        =   33
               Top             =   840
               Width           =   2295
            End
            Begin VB.Label txtDossier_CDODOSIRR 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Irr"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1800
               TabIndex        =   32
               Top             =   960
               Width           =   735
            End
            Begin VB.Label lblDossier_CDODOSIRR 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Irrévocable"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               TabIndex        =   31
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label txtDossier_CDOCOMMTV 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Mnt TVA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   4560
               TabIndex        =   30
               Top             =   5400
               Width           =   2295
            End
            Begin VB.Label lblDossier_CDOCOMMTV 
               Caption         =   "Montant T.V.A."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               TabIndex        =   29
               Top             =   5400
               Width           =   1455
            End
            Begin VB.Label lblDossier_CDODOSBEN 
               Caption         =   "Bénéficiaire"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               TabIndex        =   27
               Top             =   3480
               Width           =   1215
            End
            Begin VB.Label lblDossier_CDODOSDON 
               Caption         =   "D.O."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               TabIndex        =   25
               Top             =   2640
               Width           =   855
            End
            Begin VB.Label txtDossier_NotCDOCO2TX1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Taux"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4560
               TabIndex        =   24
               Top             =   4320
               Width           =   2295
            End
            Begin VB.Label txtDossier_CnfCDOCO2TX1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Taux"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1800
               TabIndex        =   23
               Top             =   4320
               Width           =   2295
            End
            Begin VB.Label lblDossier_CDOCO2TX1 
               Caption         =   "Taux"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               TabIndex        =   22
               Top             =   4320
               Width           =   1455
            End
            Begin VB.Label txtDossier_CDODOSDOS 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "dossier"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1800
               TabIndex        =   21
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label txtDossier_notCDOCOMMON 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "montant"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   4560
               TabIndex        =   20
               Top             =   4920
               Width           =   2295
            End
            Begin VB.Label lblDossier_CDOCOMMON 
               Caption         =   "Commission"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               TabIndex        =   19
               Top             =   4920
               Width           =   1455
            End
            Begin VB.Label txtDossier_cnfCDOCOMMON 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "montant"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1800
               TabIndex        =   18
               Top             =   4920
               Width           =   2295
            End
            Begin VB.Label txtDossier_CDODOSMON 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "montant"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   4440
               TabIndex        =   16
               Top             =   240
               Width           =   2295
            End
            Begin VB.Label lblDossier_CDODOSCOR 
               Caption         =   "Bq émettrice"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Left            =   120
               TabIndex        =   15
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label lblDossier_CDODOSDOS 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dossier"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   350
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame fraDossier_Saisie 
            Height          =   2775
            Left            =   7320
            TabIndex        =   9
            Top             =   3480
            Width           =   5895
            Begin VB.CheckBox chkDossier_prtNb 
               Alignment       =   1  'Right Justify
               Caption         =   "Courrier en 2 exemplaires"
               Height          =   375
               Left            =   120
               TabIndex        =   67
               Top             =   1680
               Value           =   1  'Checked
               Width           =   2295
            End
            Begin VB.OptionButton optDossier_UTI 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "UTILISATION"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   120
               TabIndex        =   39
               Top             =   960
               Width           =   2295
            End
            Begin VB.OptionButton optDossier_MOD 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "MOD"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   120
               TabIndex        =   38
               Top             =   600
               Width           =   2295
            End
            Begin VB.OptionButton optDossier_OUV 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "OUV"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   2295
            End
            Begin VB.ListBox lstDossier_Contact 
               Height          =   2205
               Left            =   3480
               TabIndex        =   12
               Top             =   240
               Width           =   2265
            End
            Begin VB.TextBox txtDossier_Annexe_Nb 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1320
               TabIndex        =   11
               Text            =   "4"
               Top             =   2160
               Width           =   1080
            End
            Begin VB.Label lblDossier_Annexe_Nb 
               Caption         =   "Annexes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   10
               Top             =   2160
               Width           =   960
            End
         End
         Begin VB.ListBox lstOptions 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2760
            ItemData        =   "SAB_CDO.frx":0ED4
            Left            =   7320
            List            =   "SAB_CDO.frx":0EDB
            Style           =   1  'Checkbox
            TabIndex        =   6
            Top             =   360
            Width           =   5805
         End
         Begin VB.Label lblDossier_Garantie 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "OUVERTURES              GARANTIE  de   BONNE   EXECUTION"
            Height          =   735
            Left            =   120
            TabIndex        =   80
            Top             =   6840
            Width           =   1815
         End
      End
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   135
         TabIndex        =   3
         Top             =   330
         Width           =   13290
         Begin VB.CheckBox chkCDODOSEVE 
            Caption         =   "Inclure les dossiers clôturés"
            Height          =   255
            Left            =   10440
            TabIndex        =   81
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtSelect 
            Height          =   285
            Left            =   135
            TabIndex        =   7
            Top             =   240
            Width           =   1230
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7425
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   13080
            _ExtentX        =   23072
            _ExtentY        =   13097
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            FixedCols       =   0
            RowHeightMin    =   200
            BackColor       =   14737632
            ForeColor       =   4210688
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
            FormatString    =   $"SAB_CDO.frx":0EEB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgDossier 
         Height          =   7755
         Left            =   -74880
         TabIndex        =   40
         Top             =   600
         Width           =   13275
         _ExtentX        =   23416
         _ExtentY        =   13679
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   200
         BackColor       =   14737632
         ForeColor       =   4210688
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
         FormatString    =   $"SAB_CDO.frx":0FBE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
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
      Height          =   500
      Left            =   13080
      Picture         =   "SAB_CDO.frx":1078
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
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
End
Attribute VB_Name = "frmSAB_CDO"
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
Dim SAB_CDO_Aut As typeAuthorization

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim fgDossier_FormatString As String, fgDossier_K As Integer
Dim fgDossier_RowDisplay As Integer, fgDossier_RowClick As Integer, fgDossier_ColClick As Integer
Dim fgDossier_ColorClick As Long, fgDossier_ColorDisplay As Long
Dim fgDossier_Sort1 As Integer, fgDossier_Sort2 As Integer
Dim fgDossier_SortAD As Integer, fgDossier_Sort1_Old As Integer
Dim fgDossier_arrIndex As Integer
Dim blnfgDossier_DisplayLine As Boolean

Dim fgYCDOUTI0_Document_Select_FormatString As String

Dim meYCDODOS0 As typeZCDODOS0, xYCDODOS0 As typeZCDODOS0
Dim meYCDOMOD0 As typeZCDOMOD0, xYCDOMOD0 As typeZCDOMOD0
Dim meYCDOTIE0 As typeZCDOTIE0, xYCDOTIE0 As typeZCDOTIE0
'Dim meYCDOUTI0 As typeZCDOUTI0, xYCDOUTI0 As typeZCDOUTI0
Dim blnError As Boolean

Dim meYCDOCOM0 As typeZCDOCOM0, xYCDOCOM0 As typeZCDOCOM0
Dim meYCDOCO20 As typeZCDOCO20, xYCDOCO20 As typeZCDOCO20
Dim meYCDOTC20 As typeZCDOTC20, xYCDOTC20 As typeZCDOTC20

Dim cnfYBIACDOCOM0 As typeYBIACDOCOM0, notYBIACDOCOM0 As typeYBIACDOCOM0
Dim W_BOO_Autre As String

Dim fgYCDOUTI0_FormatString As String, fgYCDOUTI0_K As Integer
Dim fgYCDOUTI0_RowDisplay As Integer, fgYCDOUTI0_RowClick As Integer, fgYCDOUTI0_ColClick As Integer
Dim fgYCDOUTI0_ColorClick As Long, fgYCDOUTI0_ColorDisplay As Long
Dim fgYCDOUTI0_Sort1 As Integer, fgYCDOUTI0_Sort2 As Integer
Dim fgYCDOUTI0_SortAD As Integer, fgYCDOUTI0_Sort1_Old As Integer
Dim fgYCDOUTI0_arrIndex As Integer
Dim blnfgYCDOUTI0_DisplayLine As Boolean

Dim meYCDOUTI0 As typeZCDOUTI0, xYCDOUTI0 As typeZCDOUTI0
Dim meYCDOREG0 As typeZCDOREG0, xYCDOREG0 As typeZCDOREG0
Dim meYCDODES0 As typeZCDODES0, xYCDODES0 As typeZCDODES0
Dim meYCDOSWI0 As typeZCDOSWI0, xYCDOSWI0 As typeZCDOSWI0
Dim meYCDOIRR0 As typeZCDOIRR0, xYCDOIRR0 As typeZCDOIRR0

Dim wYCDOIRR0 As String, blnYCDOIRR0  As Boolean

Dim meCDO_Courrier As typeCDO_Courrier
Dim blnYCDOUTI0_Ok As Boolean
Dim xElpTable As typeElpTable

Dim meYBIACDO As typeYBIACDO


Public Sub fgYCDOUTI0_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgYCDOUTI0.Row

If lRow > 0 And lRow < fgYCDOUTI0.Rows Then
    fgYCDOUTI0.Row = lRow
    For I = 0 To fgYCDOUTI0_arrIndex
        fgYCDOUTI0.Col = I: fgYCDOUTI0.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgYCDOUTI0.Row = mRow
    If fgYCDOUTI0.Row > 0 Then
        lRow = fgYCDOUTI0.Row
        lColor_Old = fgYCDOUTI0.CellBackColor
        For I = 0 To fgYCDOUTI0_arrIndex
          fgYCDOUTI0.Col = I: fgYCDOUTI0.CellBackColor = lColor
        Next I
        fgYCDOUTI0.Col = 0
    End If
End If

End Sub

Private Sub fgYCDOUTI0_Display()
Dim I As Integer
SSTab1.Tab = 1
fraDossier.Enabled = True
fgYCDOUTI0_Reset

fgYCDOUTI0.Rows = 1
fgYCDOUTI0.FormatString = fgYCDOUTI0_FormatString

End Sub


Public Sub fgYCDOUTI0_DisplayLine(lOrigine As String, lId As String, lText As String)
On Error Resume Next
fgYCDOUTI0.Rows = fgYCDOUTI0.Rows + 1
fgYCDOUTI0.Row = fgYCDOUTI0.Rows - 1
fgYCDOUTI0.Col = 0: fgYCDOUTI0.Text = lOrigine
fgYCDOUTI0.Col = 1: fgYCDOUTI0.Text = lId
fgYCDOUTI0.Col = 2: fgYCDOUTI0.Text = lText
End Sub



Private Sub chkCDODOSEVE_Click()
Me.Enabled = False
Me.MousePointer = vbHourglass

fgSelect_Display

Me.Enabled = True
Me.MousePointer = 0

End Sub

Private Sub cmdDocushare_Click()
Dim Nb As Long
Nb = DS_Document_Load(CStr(meYCDODOS0.CDODOSDOS), paramDocuShare_Collection_SAB_CDO)
Call lstErr_AddItem(lstErr, cmdContext, "nb documents : " & Nb): DoEvents
End Sub

Private Sub cmdYCDOUTI0_Document_Ok_Click()
If cmdYCDOUTI0_Document_Ok.Caption <> constModifier Then

    fgYCDOUTI0_Document_Select.Rows = fgYCDOUTI0_Document_Select.Rows + 1
    fgYCDOUTI0_Document_Select.Row = fgYCDOUTI0_Document_Select.Rows - 1
End If

fgYCDOUTI0_Document_Select.Col = 0: fgYCDOUTI0_Document_Select.Text = Trim(txtYCDOUTI0_Document)
fgYCDOUTI0_Document_Select.CellForeColor = vbMagenta
fgYCDOUTI0_Document_Select.Col = 1: fgYCDOUTI0_Document_Select.Text = Trim(txtYCDOUTI0_Document_Jeu1)
fgYCDOUTI0_Document_Select.CellForeColor = vbMagenta
fgYCDOUTI0_Document_Select.Col = 2: fgYCDOUTI0_Document_Select.Text = Trim(txtYCDOUTI0_Document_Jeu2)
fgYCDOUTI0_Document_Select.CellForeColor = vbMagenta
cmdYCDOUTI0_Document_Quit_Click
End Sub

Private Sub cmdYCDOUTI0_Document_Quit_Click()
fraYCDOUTI0_Document_Select.Visible = False
cmdYCDOUTI0_Document_Quit.Visible = False
cmdYCDOUTI0_Document_Ok.Visible = False
cmdYCDOUTI0_Document_Remove.Visible = False
txtYCDOUTI0_Document = ""
txtYCDOUTI0_Document_Jeu1 = ""
txtYCDOUTI0_Document_Jeu2 = ""

End Sub

Private Sub cmdYCDOUTI0_Document_Remove_Click()
Dim I As Integer, wRow As Integer, wRows As Integer
Dim X0 As String, X1 As String, X2 As String
wRow = fgYCDOUTI0_Document_Select.Row + 1
wRows = fgYCDOUTI0_Document_Select.Rows - 1

For I = wRow To wRows

    fgYCDOUTI0_Document_Select.Row = I
    
    fgYCDOUTI0_Document_Select.Col = 0: X0 = fgYCDOUTI0_Document_Select.Text
    fgYCDOUTI0_Document_Select.Col = 1: X1 = fgYCDOUTI0_Document_Select.Text
    fgYCDOUTI0_Document_Select.Col = 2: X2 = fgYCDOUTI0_Document_Select.Text
    
    fgYCDOUTI0_Document_Select.Row = I - 1
    fgYCDOUTI0_Document_Select.Col = 0: fgYCDOUTI0_Document_Select.Text = X0
    fgYCDOUTI0_Document_Select.CellForeColor = vbMagenta
    fgYCDOUTI0_Document_Select.Col = 1: fgYCDOUTI0_Document_Select.Text = X1
    fgYCDOUTI0_Document_Select.CellForeColor = vbMagenta
    fgYCDOUTI0_Document_Select.Col = 2: fgYCDOUTI0_Document_Select.Text = X2
    fgYCDOUTI0_Document_Select.CellForeColor = vbMagenta
Next I
fgYCDOUTI0_Document_Select.Rows = fgYCDOUTI0_Document_Select.Rows - 1

cmdYCDOUTI0_Document_Quit_Click

End Sub


Private Sub fgDossier_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
fgDossier.LeftCol = 0

End Sub

Private Sub fgSelect_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
fgSelect.LeftCol = 0
End Sub

Private Sub fgYCDOUTI0_Document_Select_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim strX As String, I As Integer, blnOk As Boolean
On Error Resume Next
On Error Resume Next
If y <= fgYCDOUTI0_Document_Select.RowHeightMin Then
Else
    If fgYCDOUTI0_Document_Select.Rows > 1 Then
        fgYCDOUTI0_Document_Select.Col = 0: txtYCDOUTI0_Document = fgYCDOUTI0_Document_Select.Text
        fgYCDOUTI0_Document_Select.Col = 1: txtYCDOUTI0_Document_Jeu1 = fgYCDOUTI0_Document_Select.Text
        fgYCDOUTI0_Document_Select.Col = 2: txtYCDOUTI0_Document_Jeu2 = fgYCDOUTI0_Document_Select.Text

        cmdYCDOUTI0_Document_Ok.Caption = constModifier
        cmdYCDOUTI0_Document_Ok.Visible = True
        cmdYCDOUTI0_Document_Quit.Visible = True
        cmdYCDOUTI0_Document_Remove.Visible = True
        fraYCDOUTI0_Document_Select.Visible = True
        txtYCDOUTI0_Document_Jeu1.SetFocus
   End If
End If


End Sub


Private Sub fgYCDOUTI0_LeaveCell()
On Error Resume Next
fgYCDOUTI0.CellBackColor = &HE0E0E0
End Sub



Public Sub fgYCDOUTI0_Reset()
fgYCDOUTI0.Clear: fgYCDOUTI0.Rows = 1
fgYCDOUTI0_Sort1 = 0: fgYCDOUTI0_Sort2 = 0
fgYCDOUTI0_Sort1_Old = -1
fgYCDOUTI0_RowDisplay = 0: fgYCDOUTI0_RowClick = 0
fgYCDOUTI0_arrIndex = 3
blnfgYCDOUTI0_DisplayLine = False
fgYCDOUTI0_SortAD = 6
fgYCDOUTI0.LeftCol = 0
End Sub

Public Sub fgYCDOUTI0_Sort()
If fgYCDOUTI0.Rows > 1 Then
    fgYCDOUTI0.Row = 1
    fgYCDOUTI0.RowSel = fgYCDOUTI0.Rows - 1
    
    If fgYCDOUTI0_Sort1_Old = fgYCDOUTI0_Sort1 Then
        If fgYCDOUTI0_SortAD = 5 Then
            fgYCDOUTI0_SortAD = 6
        Else
            fgYCDOUTI0_SortAD = 5
        End If
    Else
        fgYCDOUTI0_SortAD = 5
    End If
    fgYCDOUTI0_Sort1_Old = fgYCDOUTI0_Sort1
    
    fgYCDOUTI0.Col = fgYCDOUTI0_Sort1
    fgYCDOUTI0.ColSel = fgYCDOUTI0_Sort2
    fgYCDOUTI0.Sort = fgYCDOUTI0_SortAD
End If

End Sub
Public Sub fgYCDOUTI0_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgYCDOUTI0.Rows - 1
    fgYCDOUTI0.Row = I
    fgYCDOUTI0.Col = lK
    X = Format$(Val(fgYCDOUTI0.Text), "0000000")
    fgYCDOUTI0.Col = fgYCDOUTI0_arrIndex - 1
    Select Case lK
        Case 1, 2: fgYCDOUTI0.Text = X
    End Select
Next I


fgYCDOUTI0_Sort1 = fgYCDOUTI0_arrIndex - 1: fgYCDOUTI0_Sort2 = fgYCDOUTI0_arrIndex - 1
fgYCDOUTI0_Sort
End Sub


Private Sub fgYCDOUTI0_Click()
fgYCDOUTI0.LeftCol = 0

End Sub


Private Sub fgYCDOUTI0_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
If y <= fgYCDOUTI0.RowHeightMin Then
Else
    If fgYCDOUTI0.Rows > 1 Then
        fgYCDOUTI0.Enabled = False
        Call fgYCDOUTI0_Color(fgYCDOUTI0_RowClick, MouseMoveUsr.BackColor, fgYCDOUTI0_ColorClick)
        
        fraYCDOUTI0_Display_Reset
        
   End If
End If
End Sub


Private Sub fgDossier_Display()
Dim I As Integer
On Error Resume Next
SSTab1.Tab = 1
fraDossier.Enabled = True
fgDossier_Reset

fgDossier.Rows = 1
fgDossier.FormatString = fgDossier_FormatString

txtDossier_Annexe_Nb = 1
For I = 0 To lstOptions.ListCount - 1
    lstOptions.Selected(I) = False
Next I


End Sub

Public Sub fgDossier_DisplayLine(lOrigine As String, lId As String, lText As String)
On Error Resume Next
fgDossier.Rows = fgDossier.Rows + 1
fgDossier.Row = fgDossier.Rows - 1
fgDossier.Col = 0: fgDossier.Text = lOrigine
fgDossier.Col = 1: fgDossier.Text = lId
fgDossier.Col = 2: fgDossier.Text = lText
End Sub

Public Sub fgDossier_Sort()
If fgDossier.Rows > 1 Then
    fgDossier.Row = 1
    fgDossier.RowSel = fgDossier.Rows - 1
    
    If fgDossier_Sort1_Old = fgDossier_Sort1 Then
        If fgDossier_SortAD = 5 Then
            fgDossier_SortAD = 6
        Else
            fgDossier_SortAD = 5
        End If
    Else
        fgDossier_SortAD = 5
    End If
    fgDossier_Sort1_Old = fgDossier_Sort1
    
    fgDossier.Col = fgDossier_Sort1
    fgDossier.ColSel = fgDossier_Sort2
    fgDossier.Sort = fgDossier_SortAD
End If

End Sub
Public Sub fgDossier_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgDossier.Rows - 1
    fgDossier.Row = I
    fgDossier.Col = lK
    X = Format$(Val(fgDossier.Text), "0000000")
    fgDossier.Col = fgDossier_arrIndex - 1
    Select Case lK
        Case 1, 2: fgDossier.Text = X
    End Select
Next I


fgDossier_Sort1 = fgDossier_arrIndex - 1: fgDossier_Sort2 = fgDossier_arrIndex - 1
fgDossier_Sort
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
Dim V
Dim Nb As Long
Dim xSQL As String
Dim xW As String
On Error GoTo Error_Handler
SSTab1.Tab = 0
Call lstErr_Clear(lstErr, cmdContext, "Chargement des dossiers....")
Me.MousePointer = vbHourglass
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgSelect_Display"

Set rsSab = Nothing

xSQL = "select CDODOSDOS , CDODOSMON,CDODOSCOR from " & paramIBM_Library_SAB & ".ZCDODOS0 "
If chkCDODOSEVE <> "1" Then
    xSQL = xSQL & " where CDODOSCOP = 'CDE' and CDODOSEVE < '80' order by CDODOSDOS"
Else
    xSQL = xSQL & " where CDODOSCOP = 'CDE'  order by CDODOSDOS"
End If

'xSQL = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0"
'$2003.11.04  rsSab.Open xSQL, paramODBC_DSN_SAB
Set rsSab = cnsab.Execute(xSQL)  '$2003.11.04

Do While Not rsSab.EOF
    'Call srvYCDODOS0_GetBuffer_ODBC(rsSab, meYCDODOS0)
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect.Col = 0: fgSelect.Text = rsSab("CDODOSDOS")
        fgSelect.CellForeColor = vbBlue
        fgSelect.Col = 1: fgSelect.Text = rsSab("CDODOSCOR")

        fgSelect.Col = 2: fgSelect.Text = Format$(rsSab("CDODOSMON"), "### ### ### ##0.00")
        fgSelect.CellForeColor = vbBlue

    rsSab.MoveNext
Loop

fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_SortX 0
fgSelect.Visible = True
fgSelect.TopRow = fgSelect.Rows - 12
Call lstErr_AddItem(lstErr, cmdContext, "Nombre de dossiers : " & fgSelect.Rows - 1)
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Public Sub fgSelect_DisplayLine()
On Error Resume Next
End Sub


Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 6
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

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
    X = Format$(Val(fgSelect.Text), "000000000000000.00")
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

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), SAB_CDO_Aut)

'blnSetfocus = True
Form_Init


End Sub


Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistant", vbCritical, "frmYCDODOS0.param_init"
    Unload Me
Else
    lstErr.Clear
End If

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgDossier_FormatString = fgDossier.FormatString
fgDossier.Enabled = True
fgYCDOUTI0_FormatString = fgYCDOUTI0.FormatString
fgYCDOUTI0.Enabled = True
fgYCDOUTI0_Document_Select_FormatString = fgYCDOUTI0_Document_Select.FormatString
cmdReset
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
rsZCDODOS0_Init meYCDODOS0
xYCDODOS0 = meYCDODOS0
ReDim meYBIACDO.YCDODOS0(1): meYBIACDO.YCDODOS0(1) = meYCDODOS0


fraDossier.Enabled = False
blnControl = True
txtDossier_Annexe_Nb.ForeColor = vbMagenta
lstDossier_Contact.ForeColor = vbMagenta

fgSelect_Display
fraDossier_Info.Enabled = False   ' La frame n'est que affichage d'informations
txtDossier_CDODOSMON.ForeColor = vbBlue
txtDossier_cnfCDOCOMMON.ForeColor = vbBlue
txtDossier_notCDOCOMMON.ForeColor = vbBlue
txtDossier_Garantie.ForeColor = vbBlue
cmdPrint.Enabled = SAB_CDO_Aut.Saisir

'fraYCDOUTI0_Display.Left = 120
'fraYCDOUTI0_Display.Top = 480
fraYCDOUTI0_Display.Visible = False
'fgYCDOUTI0_Document_Select.BackColor = &HF0FFFF
txtYCDOUTI0_ATT.ForeColor = vbMagenta
txtYCDOUTI0_CrrGB_Mnt.ForeColor = vbMagenta
txtYCDOUTI0_CrrGB_Tx.ForeColor = vbMagenta
txtYCDOUTI0_CrrGB_BqRbt.ForeColor = vbMagenta
txtYCDOUTI0_IBAN.ForeColor = vbMagenta
txtYCDOUTI0_YCDODES0.ForeColor = vbMagenta
txtYCDOUTI0_DELAI.ForeColor = vbMagenta
txtYCDOUTI0_Exp_Par.ForeColor = vbMagenta
txtYCDOUTI0_Exp_Le.ForeColor = vbMagenta
txtYCDOUTI0_Exp_De.ForeColor = vbMagenta
txtYCDOUTI0_Exp_A.ForeColor = vbMagenta

End Sub


Public Function param_Init()
Dim K As Integer, K1 As Integer, X As String
Dim xName As String, xMemo As String
Dim wText As String
Dim V

param_Init = Null


lstDossier_Contact.Clear
Call lst_LoadK2("SOBI", "Contact", lstDossier_Contact, False)
Call rsElpTable_Read("SOBI", "Contact", usrName_UCase, xName, xMemo)
Call lst_Scan(Trim(xName), lstDossier_Contact)

lstYCDOUTI0_Document.Clear
X = "select * from YBIATAB0 " _
    & " where BIATABID = 'SAB'" _
    & " and BIATABK1 = 'ZCDOTAB0_004'"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    lstYCDOUTI0_Document.AddItem Mid$(rsMDB("BIATABTXT"), 25, 30)
    rsMDB.MoveNext
Loop

cmdDocushare.Visible = False: DS_Server_Open
End Function






Public Sub fgDossier_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgDossier.Row

If lRow > 0 And lRow < fgDossier.Rows Then
    fgDossier.Row = lRow
    For I = 0 To fgDossier_arrIndex
        fgDossier.Col = I: fgDossier.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgDossier.Row = mRow
    If fgDossier.Row > 0 Then
        lRow = fgDossier.Row
        lColor_Old = fgDossier.CellBackColor
        For I = 0 To fgDossier_arrIndex
          fgDossier.Col = I: fgDossier.CellBackColor = lColor
        Next I
        fgDossier.Col = 0
    End If
End If

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

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Dim prtI As Integer
Dim I As Long, iLen As Long

meCDO_Courrier.prtNb = 2
If chkDossier_prtNb = "0" Then meCDO_Courrier.prtNb = 1

meCDO_Courrier.Contact = lstDossier_Contact
meCDO_Courrier.Annexe_Nb = Val(txtDossier_Annexe_Nb)

meCDO_Courrier.ATT = txtYCDOUTI0_ATT
meCDO_Courrier.CrrGB_Mnt = txtYCDOUTI0_CrrGB_Mnt
meCDO_Courrier.CrrGB_Tx = txtYCDOUTI0_CrrGB_Tx
meCDO_Courrier.CrrGB_BqRbt = txtYCDOUTI0_CrrGB_BqRbt
meCDO_Courrier.IBAN = txtYCDOUTI0_IBAN
meCDO_Courrier.Exp_Par = txtYCDOUTI0_Exp_Par
meCDO_Courrier.Exp_Le = txtYCDOUTI0_Exp_Le
meCDO_Courrier.Exp_De = txtYCDOUTI0_Exp_De
meCDO_Courrier.Exp_A = txtYCDOUTI0_Exp_A

' Rubrique Garantie de bonne exécution pour OUVERTURES -OU- Frais bq émettrice pour MODIFICATIONS
meCDO_Courrier.Garantie_Nb = 1
ReDim meCDO_Courrier.Garantie(1)
meCDO_Courrier.Garantie(meCDO_Courrier.Garantie_Nb) = ""
iLen = Len(txtDossier_Garantie)

For I = 1 To iLen
    X = Mid$(txtDossier_Garantie, I, 1)
    Select Case X
        Case Asc10:
        Case Asc13:
            meCDO_Courrier.Garantie_Nb = meCDO_Courrier.Garantie_Nb + 1
            ReDim Preserve meCDO_Courrier.Garantie(meCDO_Courrier.Garantie_Nb)
            meCDO_Courrier.Garantie(meCDO_Courrier.Garantie_Nb) = ""
        Case Else: meCDO_Courrier.Garantie(meCDO_Courrier.Garantie_Nb) = meCDO_Courrier.Garantie(meCDO_Courrier.Garantie_Nb) & X
    End Select
Next I

' ZONES CONCERNANT UTILISATIONS
meCDO_Courrier.Exp_Nb = 1
ReDim meCDO_Courrier.Exp(1)
meCDO_Courrier.Exp(meCDO_Courrier.Exp_Nb) = ""
iLen = Len(txtYCDOUTI0_YCDODES0)

For I = 1 To iLen
    X = Mid$(txtYCDOUTI0_YCDODES0, I, 1)
    Select Case X
        Case Asc10:
        Case Asc13:
            meCDO_Courrier.Exp_Nb = meCDO_Courrier.Exp_Nb + 1
            ReDim Preserve meCDO_Courrier.Exp(meCDO_Courrier.Exp_Nb)
            meCDO_Courrier.Exp(meCDO_Courrier.Exp_Nb) = ""
        Case Else: meCDO_Courrier.Exp(meCDO_Courrier.Exp_Nb) = meCDO_Courrier.Exp(meCDO_Courrier.Exp_Nb) & X
    End Select
Next I

meCDO_Courrier.Delai_Nb = 1
ReDim meCDO_Courrier.Delai(1)
meCDO_Courrier.Delai(meCDO_Courrier.Delai_Nb) = ""
iLen = Len(txtYCDOUTI0_DELAI)

For I = 1 To iLen
    X = Mid$(txtYCDOUTI0_DELAI, I, 1)
    Select Case X
        Case Asc10:
        Case Asc13:
            meCDO_Courrier.Delai_Nb = meCDO_Courrier.Delai_Nb + 1
            ReDim Preserve meCDO_Courrier.Delai(meCDO_Courrier.Delai_Nb)
            meCDO_Courrier.Delai(meCDO_Courrier.Delai_Nb) = ""
        Case Else: meCDO_Courrier.Delai(meCDO_Courrier.Delai_Nb) = meCDO_Courrier.Delai(meCDO_Courrier.Delai_Nb) & X
    End Select
Next I

meCDO_Courrier.Irregul_Nb = 1
ReDim meCDO_Courrier.Irregul(1)
meCDO_Courrier.Irregul(meCDO_Courrier.Irregul_Nb) = ""
iLen = Len(wYCDOIRR0)

For I = 1 To iLen
    X = Mid$(wYCDOIRR0, I, 1)
    Select Case X
        Case Asc10:
        Case Asc13:
            meCDO_Courrier.Irregul_Nb = meCDO_Courrier.Irregul_Nb + 1
            ReDim Preserve meCDO_Courrier.Irregul(meCDO_Courrier.Irregul_Nb)
            meCDO_Courrier.Irregul(meCDO_Courrier.Irregul_Nb) = ""
        Case Else: meCDO_Courrier.Irregul(meCDO_Courrier.Irregul_Nb) = meCDO_Courrier.Irregul(meCDO_Courrier.Irregul_Nb) & X
    End Select
Next I

meCDO_Courrier.Document_Nb = fgYCDOUTI0_Document_Select.Rows - 1
ReDim meCDO_Courrier.Document(meCDO_Courrier.Document_Nb + 1)
ReDim meCDO_Courrier.Document_Jeu1(meCDO_Courrier.Document_Nb + 1)
ReDim meCDO_Courrier.Document_Jeu2(meCDO_Courrier.Document_Nb + 1)
For I = 1 To meCDO_Courrier.Document_Nb
    fgYCDOUTI0_Document_Select.Row = I
    fgYCDOUTI0_Document_Select.Col = 0
    meCDO_Courrier.Document(I) = fgYCDOUTI0_Document_Select.Text
    fgYCDOUTI0_Document_Select.Col = 1
    meCDO_Courrier.Document_Jeu1(I) = fgYCDOUTI0_Document_Select.Text
    fgYCDOUTI0_Document_Select.Col = 2
    meCDO_Courrier.Document_Jeu2(I) = fgYCDOUTI0_Document_Select.Text
Next I

If fgDossier.Rows > 1 Then
    For prtI = 1 To meCDO_Courrier.prtNb
        If prtI > 1 Then frmElpPrt.prtColor_Check_1
        Select Case SSTab1.Tab
            Case 1: prtCDO_Courrier_Monitor lstOptions, meYCDODOS0, meYCDOMOD0, cnfYBIACDOCOM0, notYBIACDOCOM0, meCDO_Courrier
            Case 2:
                If blnYCDOUTI0_Ok Then
                    prtCDO_Courrier_Monitor lstYCDOUTI0_Courrier, meYCDODOS0, meYCDOMOD0, cnfYBIACDOCOM0, notYBIACDOCOM0, meCDO_Courrier
                Else
                    MsgBox "Impression non gérée pour ce dossier", vbCritical, "frmSAB_CDO.cmdPrint"
                End If
                
        End Select
    Next prtI
    frmElpPrt.prtColor_Check
End If
End Sub

Private Sub cmdSelect(lK1 As String)
Dim wId As String, wId2 As String
Dim V
Dim wCDODOS As String
Dim X As String

On Error GoTo Error_Handler
Call lstErr_Clear(lstErr, cmdContext, "dossier : " & lK1): DoEvents

currentAction = "CmdSelect"
fraYCDOUTI0_Display.Visible = False
cmdDocushare.Visible = True
fgYCDOUTI0.Enabled = True
fgYCDOUTI0_Display
fgDossier_Display
rsZCDODOS0_Init meYCDODOS0
rsZCDOMOD0_Init meYCDOMOD0

meYBIACDO.YCDODOS0_Nb = 0
ReDim meYBIACDO.YCDOMOD0(5): meYBIACDO.YCDOMOD0_Nb = 0
ReDim meYBIACDO.YCDOTIE0(1): meYBIACDO.YCDOTIE0_Nb = 0
ReDim meYBIACDO.YCDOREG0(5): meYBIACDO.YCDOREG0_Nb = 0
ReDim meYBIACDO.YCDOUTI0(5): meYBIACDO.YCDOUTI0_Nb = 0
ReDim meYBIACDO.YCDODES0(5): meYBIACDO.YCDODES0_Nb = 0
ReDim meYBIACDO.YCDOSWI0(5): meYBIACDO.YCDOSWI0_Nb = 0
ReDim meYBIACDO.YCDOIRR0(5): meYBIACDO.YCDOIRR0_Nb = 0
ReDim meYBIACDO.YCDOCOM0(5): meYBIACDO.YCDOCOM0_Nb = 0
ReDim meYBIACDO.YCDOCO20(5): meYBIACDO.YCDOCO20_Nb = 0
ReDim meYBIACDO.YCDOTC20(5): meYBIACDO.YCDOTC20_Nb = 0

    currentAction = "Recherche Dossier "  ''& paramODBC_DSN_SAB
    meYBIACDO.YCDODOS0(1).CDODOSDOS = CLng(Val(lK1))
    V = srvYBIACDO_ODBC(meYBIACDO)
    If Not IsNull(V) Then GoTo Error_MsgBox
    fgDossier_Display_YBIACDO
    meYCDODOS0 = meYBIACDO.YCDODOS0(1)
    If meYBIACDO.YCDOMOD0_Nb > 0 Then meYCDOMOD0 = meYBIACDO.YCDOMOD0(meYBIACDO.YCDOMOD0_Nb)

     
Call srvYBIACDOCOM0_Load(meYBIACDO, cnfYBIACDOCOM0, notYBIACDOCOM0)
     
'Lecture BQE agence - Donneur d'ordre - Bénéficiaire
Call rsZCDOTIE_Adresse(meYCDODOS0.CDODOSCOT, meYCDODOS0.CDODOSCOR, X, xYCDOTIE0, meCDO_Courrier.BQE_ZADRESS0, meCDO_Courrier.BQE_Concat, "CD")
Call rsZCDOTIE_Adresse(meYCDODOS0.CDODOSDOR, meYCDODOS0.CDODOSDON, meYCDODOS0.CDODOSDOE, xYCDOTIE0, meCDO_Courrier.DON_ZADRESS0, meCDO_Courrier.DON_Concat, "CD")
Call rsZCDOTIE_Adresse(meYCDODOS0.CDODOSBER, meYCDODOS0.CDODOSBEN, meYCDODOS0.CDODOSBEI, xYCDOTIE0, meCDO_Courrier.BEN_ZADRESS0, meCDO_Courrier.BEN_Concat, "  ")

'Lecture Adresse siège pour couurier Bordereau Envoi Doc si bq émettrice : 00110066 (BNA)
If Trim(meYCDODOS0.CDODOSNOT) = "0011066" Then
    Call rsZCDOTIE_Adresse(meYCDODOS0.CDODOSNOR, meYCDODOS0.CDODOSNOT, X, xYCDOTIE0, meCDO_Courrier.BED_ZADRESS0, meCDO_Courrier.BED_Concat, "  ")
Else
    meCDO_Courrier.BED_ZADRESS0 = meCDO_Courrier.BQE_ZADRESS0
    meCDO_Courrier.BED_Concat = meCDO_Courrier.BQE_Concat
    
End If

' >>>>>>  Partie Gauche de l'écran avec les informations affichées du dossier

txtDossier_CDODOSDOS = meYCDODOS0.CDODOSCOP & " " & meYCDODOS0.CDODOSDOS
txtDossier_CDODOSMON = Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00") & " " & meYCDODOS0.CDODOSDEV

Select Case meYCDODOS0.CDODOSCON
    Case "C": lblDossier_CDODOSDOS = "CONFIRME"
    Case "N": lblDossier_CDODOSDOS = "NOTIFIE"
    Case "P": lblDossier_CDODOSDOS = "PARTIEL"
    Case Else: lblDossier_CDODOSDOS = "??????????"
End Select

If meYCDODOS0.CDODOSIRR = "O" Then   ' IRROVACABLE = O = OPERATIVONNEL
    txtDossier_CDODOSIRR = "Oui"
Else
    txtDossier_CDODOSIRR = "Non"
End If
txtDossier_CDODOSMOV = Format$(meYCDODOS0.CDODOSMOV, "### ### ### ##0.00")
txtDossier_CDODOSMDI = Format$(meYCDODOS0.CDODOSMDI, "### ### ### ##0.00")

txtDossier_CDODOSCOR = meYCDODOS0.CDODOSCOT & meYCDODOS0.CDODOSCOR & " " & meCDO_Courrier.BQE_Concat
txtDossier_CDODOSDON = meYCDODOS0.CDODOSDOR & meYCDODOS0.CDODOSDON & " " & meCDO_Courrier.DON_Concat
txtDossier_CDODOSBEN = meYCDODOS0.CDODOSBER & meYCDODOS0.CDODOSBEN & " " & meCDO_Courrier.BEN_Concat

txtDossier_CnfCDOCO2TX1 = Format$(cnfYBIACDOCOM0.CDOCO2TX1, "### ### ### ##0.00")
txtDossier_NotCDOCO2TX1 = Format$(notYBIACDOCOM0.CDOCO2TX1 / 100, "### ### ### ##0.00")
txtDossier_cnfCDOCOMMON = Format$(cnfYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00") & " " & cnfYBIACDOCOM0.CDOCOMDEV
txtDossier_notCDOCOMMON = Format$(notYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00") & " " & notYBIACDOCOM0.CDOCOMDEV
txtDossier_CDOCOMMTV = Format$(notYBIACDOCOM0.CDOCOMMTV, "### ### ### ##0.00") & " " & notYBIACDOCOM0.CDOCOMDEV
txtDossier_Garantie = ""

If meYCDOMOD0.CDOMODDOS = 0 Then
    If optDossier_OUV = "1" Then
        optDossier_OUV_Click
    Else
        optDossier_OUV = "1"
    End If
    
Else
    If optDossier_MOD = "1" Then
        optDossier_MOD_Click
    Else
        optDossier_MOD = "1"
    End If
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction & " : " & lK1
End Sub

Private Sub fgDossier_LeaveCell()
On Error Resume Next
fgDossier.CellBackColor = &HE0E0E0
End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wK1 As String
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_SortX fgSelect_Sort1
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
       ' fgSelect.Col = fgSelect_arrIndex: wK1 = fgSelect.Text
        fgSelect.Col = 0: wK1 = fgSelect.Text
        cmdSelect wK1
        
   End If
End If
End Sub



Private Sub lstYCDOUTI0_Document_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim strX As String, I As Integer, blnOk As Boolean
On Error Resume Next

cmdYCDOUTI0_Document_Ok.Visible = False
strX = Trim(lstYCDOUTI0_Document)
blnOk = True

For I = 1 To fgYCDOUTI0_Document_Select.Rows - 1
    fgYCDOUTI0_Document_Select.Col = 0
    fgYCDOUTI0_Document_Select.Row = I
    If strX = Trim(fgYCDOUTI0_Document_Select.Text) Then blnOk = False
Next I
cmdYCDOUTI0_Document_Ok.Visible = blnOk
cmdYCDOUTI0_Document_Quit.Visible = blnOk
fraYCDOUTI0_Document_Select.Visible = blnOk
If blnOk Then
    cmdYCDOUTI0_Document_Ok.Caption = constAjouter
    txtYCDOUTI0_Document = strX
    txtYCDOUTI0_Document_Jeu1.SetFocus
End If
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
        If currentActiveControl_Name <> "txtDossier_Garantie" Then KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select


End Sub

Public Sub cmdContext_Quit()
blnControl = False
lstErr.Clear: lstErr.Height = 200
If SSTab1.Tab = 2 Then
    If fraYCDOUTI0_Document_Select.Visible Then
        fraYCDOUTI0_Document_Select.Visible = False
        SSTab1.Tab = 2
        Exit Sub
    End If
    
    If fraYCDOUTI0_Display.Visible Then
        fraYCDOUTI0_Display.Visible = False
        SSTab1.Tab = 2
        fgYCDOUTI0.LeftCol = 0
        fgYCDOUTI0.Enabled = True
    
        Exit Sub
    End If
End If

If fraDossier.Enabled Then
    fraDossier.Enabled = False
    SSTab1.Tab = 0
    fgSelect.LeftCol = 0

    Exit Sub
End If


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
    fgSelect.Row = fgSelect.TopRow
    fgSelect.Col = fgSelect_arrIndex: ' wK1 = fgSelect.Text
    cmdSelect txtSelect ''fgSelect.Text

'    cmdSelect_Click
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
fgDossier.Clear: fgDossier.Row = 0

Exit Sub

Error_Handler:
blnControl = False

End Sub





Private Sub fgDossier_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String
Dim I As Integer
On Error Resume Next
If y <= fgDossier.RowHeightMin Then
Else
    If fgDossier.Rows > 1 Then
        Call fgDossier_Color(fgDossier_RowClick, MouseMoveUsr.BackColor, fgDossier_ColorClick)
        fgDossier.Col = 0: wOrigine = Trim(fgDossier.Text)
                            
            fgDossier.Col = 1
            I = Val(fgDossier.Text)
            Select Case wOrigine
            '     Case "ZCDODOS0": xYCDODOS0 = meYBIACDO.YCDODOS0(I): srvYCDODOS0_ElpDisplay xYCDODOS0
            '     Case "ZCDOMOD0": xYCDOMOD0 = meYBIACDO.YCDOMOD0(I): srvYCDOMOD0_ElpDisplay xYCDOMOD0
            '     Case "ZCDOTIE0": xYCDOTIE0 = meYBIACDO.YCDOTIE0(I): srvYCDOTIE0_ElpDisplay xYCDOTIE0
            '     Case "ZCDOCOM0": xYCDOCOM0 = meYBIACDO.YCDOCOM0(I): srvYCDOCOM0_ElpDisplay xYCDOCOM0
            '     Case "ZCDOCO20": xYCDOCO20 = meYBIACDO.YCDOCO20(I): srvYCDOCO20_ElpDisplay xYCDOCO20
            '     Case "ZCDOTC20": xYCDOTC20 = meYBIACDO.YCDOTC20(I): srvYCDOTC20_ElpDisplay xYCDOTC20
            '     Case "ZCDOUTI0": xYCDOUTI0 = meYBIACDO.YCDOUTI0(I): srvYCDOUTI0_ElpDisplay xYCDOUTI0
            '     Case "ZCDOREG0": xYCDOREG0 = meYBIACDO.YCDOREG0(I): srvYCDOREG0_ElpDisplay xYCDOREG0
            '     Case "ZCDODES0": xYCDODES0 = meYBIACDO.YCDODES0(I): srvYCDODES0_ElpDisplay xYCDODES0
            '     Case "ZCDOSWI0": xYCDOSWI0 = meYBIACDO.YCDOSWI0(I): srvYCDOSWI0_ElpDisplay xYCDOSWI0
            '     Case "ZCDOIRR0": xYCDOIRR0 = meYBIACDO.YCDOIRR0(I): srvYCDOIRR0_ElpDisplay xYCDOIRR0
            End Select
   End If
End If
End Sub

Public Sub fgDossier_Reset()
fgDossier.Clear
fgDossier_Sort1 = 0: fgDossier_Sort2 = 0
fgDossier_Sort1_Old = -1
fgDossier_RowDisplay = 0: fgDossier_RowClick = 0
fgDossier_arrIndex = 3
blnfgDossier_DisplayLine = False
fgDossier_SortAD = 6
fgDossier.LeftCol = 0

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



Private Sub optDossier_MOD_Click()
cmdSelect_MOD
End Sub

Private Sub optDossier_OUV_Click()
cmdSelect_OUV
End Sub

Private Sub optDossier_UTI_Click()
'''cmdSelect_UTI
SSTab1.Tab = 2
If fgYCDOUTI0.Rows = 2 Then fgYCDOUTI0.Row = 1: fraYCDOUTI0_Display_Reset

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then txtSelect.SetFocus

End Sub

Private Sub txtDossier_Annexe_Nb_GotFocus()
Call txt_GotFocus(txtDossier_Annexe_Nb)

End Sub


Private Sub txtDossier_Annexe_Nb_LostFocus()
Call txt_LostFocus(txtDossier_Annexe_Nb)
txtDossier_Annexe_Nb.ForeColor = vbMagenta

End Sub


Private Sub txtSelect_Change()
Dim I As Long, X As String, lenX As Integer
On Error Resume Next
X = Trim(txtSelect)
lenX = Len(X)
fgSelect.Col = 0
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    
    If X <= Mid$(fgSelect.Text, 1, lenX) Then
        fgSelect.LeftCol = 0
        fgSelect.TopRow = I
        Exit Sub
    End If
Next I

End Sub

Private Sub txtSelect_GotFocus()
Call txt_GotFocus(txtSelect)

End Sub


Private Sub txtSelect_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub


Private Sub txtSelect_LostFocus()
Call txt_LostFocus(txtSelect)

End Sub

Public Sub lstOptions_Load_OUV()

lstOptions.Clear

lstOptions.AddItem "OUV_01_OP_Forfait"
lstOptions.AddItem "OUV_08_NOP_Forfait"

lstOptions.AddItem "OUV_CNF_Page1"
lstOptions.AddItem "OUV_C_NOP_11_AVueNonRecl"
lstOptions.AddItem "OUV_C_NOP_12_AVueRecl"
lstOptions.AddItem "OUV_C_NOP_13_PDifNonRecl"
lstOptions.AddItem "OUV_C_NOP_14_PDifRecl"
lstOptions.AddItem "OUV_C_OP_04_AVueNonRecl"
lstOptions.AddItem "OUV_C_OP_05_AVueRecl"
lstOptions.AddItem "OUV_C_OP_06_PDifNonRecl"
lstOptions.AddItem "OUV_C_OP_07_PDifRecl"

lstOptions.AddItem "OUV_NOT_Page1"
lstOptions.AddItem "OUV_N_NOP_25_NonRecl"
lstOptions.AddItem "OUV_N_NOP_26_Recl"
lstOptions.AddItem "OUV_N_OP_20_NonRecl"
lstOptions.AddItem "OUV_N_OP_21_Recl"

lstOptions.AddItem "OUV_PAR_Page1"
lstOptions.AddItem "OUV_P_NOP_39_AVueNonRecl"
lstOptions.AddItem "OUV_P_NOP_40_AVueRecl"
lstOptions.AddItem "OUV_P_NOP_41_PDifNonRecl"
lstOptions.AddItem "OUV_P_NOP_42_PDifRecl"
lstOptions.AddItem "OUV_P_OP_32_AVueNonRecl"
lstOptions.AddItem "OUV_P_OP_33_AVueRecl"
lstOptions.AddItem "OUV_P_OP_34_PDifNonRecl"
lstOptions.AddItem "OUV_P_OP_35_PDifRecl"

lstOptions.AddItem "OUV_Annexe_02"
lstOptions.AddItem "OUV_Annexe_03"
lstOptions.AddItem "OUV_Annexe_09"
lstOptions.AddItem "OUV_Annexe_10"

' Le 21 / 1 / 2005:
lblDossier_Garantie.Caption = "OUVERTURES : GARANTIE de BONNE EXECUTION"

End Sub

Public Sub lstOptions_Load_UTI()

lstYCDOUTI0_Courrier.Clear

lstYCDOUTI0_Courrier.AddItem "UTI_CONFORME_BED1_Pdif"
lstYCDOUTI0_Courrier.AddItem "UTI_CONFORME_BED1_Pdif_GB"
lstYCDOUTI0_Courrier.AddItem "UTI_CONFORME_BED2_Avue"
lstYCDOUTI0_Courrier.AddItem "UTI_CONFORME_BED2_Avue_GB"
lstYCDOUTI0_Courrier.AddItem "UTI_CONFORME_C_Avue_AR1"
lstYCDOUTI0_Courrier.AddItem "UTI_CONFORME_C_Pdif_AR2"
lstYCDOUTI0_Courrier.AddItem "UTI_CONFORME_N_Avue_AR5"
lstYCDOUTI0_Courrier.AddItem "UTI_CONFORME_N_Pdif_AR6"
lstYCDOUTI0_Courrier.AddItem "UTI_NCONFORME_AR"
lstYCDOUTI0_Courrier.AddItem "UTI_NCONFORME_BED_Avue"
lstYCDOUTI0_Courrier.AddItem "UTI_NCONFORME_BED_Pdif"
lstYCDOUTI0_Courrier.AddItem "UTI_ACCORDRECU_C_Avue_AR10"
lstYCDOUTI0_Courrier.AddItem "UTI_ACCORDRECU_C_Pdif_AR11"
lstYCDOUTI0_Courrier.AddItem "UTI_ACCORDRECU_N_Avue_AR12"
lstYCDOUTI0_Courrier.AddItem "UTI_ACCORDRECU_N_Pdif_AR13"

End Sub


Public Sub lstOptions_Load_MOD()

lstOptions.Clear

lstOptions.AddItem "MOD_19_Diminution"

lstOptions.AddItem "MOD_FCB_01_ValProrogée"
lstOptions.AddItem "MOD_FCB_02_ValProrogée_Emb"
lstOptions.AddItem "MOD_FCB_03_Annexe"
lstOptions.AddItem "MOD_FCB_09_ValRaccourcie"
lstOptions.AddItem "MOD_FCB_10_ValRaccourcie_Emb"
lstOptions.AddItem "MOD_FCB_11_CNF_En_NOT"
lstOptions.AddItem "MOD_FCB_13_NOT_En_CNF"
lstOptions.AddItem "MOD_FCB_17_Augmentation_CNF"
lstOptions.AddItem "MOD_FCB_20_Augmentation_NOT"

lstOptions.AddItem "MOD_FCO_04_ValProrogée"
lstOptions.AddItem "MOD_FCO_05_ValProrogée_Emb"
lstOptions.AddItem "MOD_FCO_06_Annexe"
lstOptions.AddItem "MOD_FCO_07_ValRaccourcie"
lstOptions.AddItem "MOD_FCO_08_ValRaccourcie_Emb"
lstOptions.AddItem "MOD_FCO_12_CNF_En_NOT"
lstOptions.AddItem "MOD_FCO_14_NOT_En_CNF"
lstOptions.AddItem "MOD_FCO_18_Augmentation"

' Le 21 / 1 / 2005:
lblDossier_Garantie.Caption = "MODIFICATIONS : FRAIS BANQUE EMETTRICE"

End Sub


Public Sub cmdSelect_OUV()

' Chargement de la fenêtre des lettres disponibles en OUVERTURE
lstOptions_Load_OUV

' >>>>>> OUVERTURE : Cocher les cases automatiquement
Select Case meYCDODOS0.CDODOSCON
    Case "C": lstOptions.Selected(2) = True
              If meYCDODOS0.CDODOSIRR = "O" Then       ' >>>>> OPERATIF
                If cnfYBIACDOCOM0.CDOTC2MTF <> 0 Then    '( Forfaitaire )
                     lstOptions.Selected(0) = True
                Else
                    ' Si BEA - BADR - BDL : Réclamation des commissions en annexe joint à la lettre
                    If meYCDODOS0.CDODOSNOT = "0011001" Or meYCDODOS0.CDODOSNOT = "0011076" Or meYCDODOS0.CDODOSNOT = "0011077" Then
                      If meYCDODOS0.CDODOSMOV <> 0 Then lstOptions.Selected(8) = True      'Annexe_05
                      If meYCDODOS0.CDODOSMDI <> 0 Then lstOptions.Selected(10) = True     'Annexe_07
                    Else
                      If meYCDODOS0.CDODOSMOV <> 0 Then lstOptions.Selected(7) = True      'Annexe_04
                      If meYCDODOS0.CDODOSMDI <> 0 Then lstOptions.Selected(9) = True      'Annexe_06
                    End If
                End If
              Else                                     ' >>>>> NON OPERATIF
                If cnfYBIACDOCOM0.CDOTC2MTF <> 0 Then    '( Forfaitaire )
                     lstOptions.Selected(1) = True
                Else
                    ' Si BEA - BADR - BDL : Réclamation des commissions en annexe joint à la lettre
                    If meYCDODOS0.CDODOSNOT = "0011001" Or meYCDODOS0.CDODOSNOT = "0011076" Or meYCDODOS0.CDODOSNOT = "0011077" Then
                      If meYCDODOS0.CDODOSMOV <> 0 Then lstOptions.Selected(4) = True      'Annexe_12
                      If meYCDODOS0.CDODOSMDI <> 0 Then lstOptions.Selected(6) = True      'Annexe_14
                    Else
                      If meYCDODOS0.CDODOSMOV <> 0 Then lstOptions.Selected(3) = True      'Annexe_11
                      If meYCDODOS0.CDODOSMDI <> 0 Then lstOptions.Selected(5) = True      'Annexe_13
                    End If
                End If
              End If
              
    Case "N": lstOptions.Selected(11) = True
              If meYCDODOS0.CDODOSIRR = "O" Then       ' >>>>> OPERATIF
                If notYBIACDOCOM0.CDOTC2MTF <> 0 Then    '( Forfaitaire )
                    lstOptions.Selected(0) = True
                Else
                    ' Si BEA - BADR - BDL : Réclamation des commissions en annexe jointe à la lettre
                    ' If meYCDODOS0.CDODOSNOT = "0011001" Or meYCDODOS0.CDODOSNOT = "0011076" Or meYCDODOS0.CDODOSNOT = "0011077" Then
                    '   lstOptions.Selected(15) = True     'Annexe_21
                    ' Else
                      lstOptions.Selected(14) = True     'Annexe_20
                    ' End If
                End If
              Else                                     ' >>>>> NON OPERATIF
                If notYBIACDOCOM0.CDOTC2MTF <> 0 Then    '( Forfaitaire )
                      lstOptions.Selected(1) = True
                Else
                    ' Si BEA - BADR - BDL : Réclamation des commissions en annexe jointe à la lettre
                    ' If meYCDODOS0.CDODOSNOT = "0011001" Or meYCDODOS0.CDODOSNOT = "0011076" Or meYCDODOS0.CDODOSNOT = "0011077" Then
                    '   lstOptions.Selected(13) = True     'Annexe_26
                    ' Else
                      lstOptions.Selected(12) = True     'Annexe_25
                    ' End If
                End If
              End If
                  
    Case "P": lstOptions.Selected(16) = True
              If meYCDODOS0.CDODOSIRR = "O" Then       ' >>>>> OPERATIF
                ' If meYCDOTC20.CDOTC2MTF <> 0 Then    ( Forfaitaire )
                '     lstOptions.Selected(3) = True
                ' Else
                ' Si BEA - BADR - BDL : Réclamation des commissions en annexe joint à la lettre
                If meYCDODOS0.CDODOSNOT = "0011001" Or meYCDODOS0.CDODOSNOT = "0011076" Or meYCDODOS0.CDODOSNOT = "0011077" Then
                  If meYCDODOS0.CDODOSMOV <> 0 And meYCDODOS0.CDODOSMDI = 0 Then lstOptions.Selected(22) = True    'Annexe_33
                  If meYCDODOS0.CDODOSMDI <> 0 Then lstOptions.Selected(24) = True    'Annexe_35
                Else
                  If meYCDODOS0.CDODOSMOV <> 0 And meYCDODOS0.CDODOSMDI = 0 Then lstOptions.Selected(21) = True    'Annexe_32
                  If meYCDODOS0.CDODOSMDI <> 0 Then lstOptions.Selected(23) = True    'Annexe_34
                End If
                ' End If
              Else                                     ' >>>>> NON OPERATIF
                ' If meYCDOTC20.CDOTC2MTF <> 0 Then    ( Forfaitaire )
                '     lstOptions.Selected(3) = True
                ' Else
                ' Si BEA - BADR - BDL : Réclamation des commissions en annexe joint à la lettre
                If meYCDODOS0.CDODOSNOT = "0011001" Or meYCDODOS0.CDODOSNOT = "0011076" Or meYCDODOS0.CDODOSNOT = "0011077" Then
                  If meYCDODOS0.CDODOSMOV <> 0 And meYCDODOS0.CDODOSMDI = 0 Then lstOptions.Selected(18) = True    'Annexe_40
                  If meYCDODOS0.CDODOSMDI <> 0 Then lstOptions.Selected(20) = True    'Annexe_42
                Else
                  If meYCDODOS0.CDODOSMOV <> 0 And meYCDODOS0.CDODOSMDI = 0 Then lstOptions.Selected(17) = True    'Annexe_39
                  If meYCDODOS0.CDODOSMDI <> 0 Then lstOptions.Selected(19) = True    'Annexe_41
                End If
                ' End If
              End If
                  
End Select

End Sub

Public Sub cmdSelect_UTI()

' Chargement de la fenêtre des lettres disponibles en UTILISATION
lstOptions_Load_UTI

' >>>>>> UTILISATION : Cocher les cases automatiquement
Select Case meYCDOUTI0.CDOUTIDCO
    Case "O":   ' >>>>> UTILISATION AVEC DOCUMENTS CONFORMES
                Select Case meYCDOUTI0.CDOUTITMO
                    Case "C":   ' --> Utilisation CONFIRMEE
                        If meYCDOUTI0.CDOUTIMVU <> 0 And meYCDOUTI0.CDOUTIMDI = 0 Then   ' (AVUE)
                             lstYCDOUTI0_Courrier.Selected(2) = True    ' CONFORME_BED2
                             lstYCDOUTI0_Courrier.Selected(4) = True    ' CONFORME_AR1
                        End If
                        If meYCDOUTI0.CDOUTIMVU = 0 And meYCDOUTI0.CDOUTIMDI <> 0 Then   ' (PDIF)
                             lstYCDOUTI0_Courrier.Selected(0) = True    ' CONFORME_BED1
                             lstYCDOUTI0_Courrier.Selected(5) = True    ' CONFORME_AR2
                        End If
                    Case "N":   ' --> Utilisation NOTIFIEE
                        If meYCDOUTI0.CDOUTIMVU <> 0 And meYCDOUTI0.CDOUTIMDI = 0 Then   ' (AVUE)
                             lstYCDOUTI0_Courrier.Selected(2) = True    ' CONFORME_BED2
                             lstYCDOUTI0_Courrier.Selected(6) = True    ' CONFORME_AR5
                        End If
                        If meYCDOUTI0.CDOUTIMVU = 0 And meYCDOUTI0.CDOUTIMDI <> 0 Then   ' (PDIF)
                             lstYCDOUTI0_Courrier.Selected(0) = True    ' CONFORME_BED1
                             lstYCDOUTI0_Courrier.Selected(7) = True    ' CONFORME_AR6
                        End If
                End Select
    Case "N":   ' >>>>> UTILISATION AVEC DOCUMENTS NON CONFORMES : ECHU OU VALIDE SONT TESTE DANS LETTRES
                If meYCDOUTI0.CDOUTIEVE = "04" Then
                    ' Accord APRES réception documents
                    If meYCDOUTI0.CDOUTITMO = "C" Then  ' UTIL de la partie CONFIRME
                        If meYCDOUTI0.CDOUTIMVU <> 0 And meYCDOUTI0.CDOUTIMDI = 0 Then   ' (AVUE)
                            lstYCDOUTI0_Courrier.Selected(11) = True              ' ACCORDRECU_C_Avue_AR10
                        End If
                        If meYCDOUTI0.CDOUTIMVU = 0 And meYCDOUTI0.CDOUTIMDI <> 0 Then   ' (PDIF)
                            lstYCDOUTI0_Courrier.Selected(12) = True              ' ACCORDRECU_C_Pdif_AR11
                        End If
                    End If
                    If meYCDOUTI0.CDOUTITMO = "N" Then  ' UTIL de la partie NOTIFIEE
                        If meYCDOUTI0.CDOUTIMVU <> 0 And meYCDOUTI0.CDOUTIMDI = 0 Then   ' (AVUE)
                            lstYCDOUTI0_Courrier.Selected(13) = True              ' ACCORDRECU_C_Avue_AR12
                        End If
                        If meYCDOUTI0.CDOUTIMVU = 0 And meYCDOUTI0.CDOUTIMDI <> 0 Then   ' (PDIF)
                            lstYCDOUTI0_Courrier.Selected(14) = True              ' ACCORDRECU_C_Pdif_AR13
                        End If
                    End If
                Else
                    ' Documents partis pour accord
                    lstYCDOUTI0_Courrier.Selected(8) = True                 ' NCONFORME_AR
                    If meYCDOUTI0.CDOUTIMVU <> 0 And meYCDOUTI0.CDOUTIMDI = 0 Then   ' (AVUE)
                       lstYCDOUTI0_Courrier.Selected(9) = True              ' NCONFORME_BED_Avue
                    End If
                    If meYCDOUTI0.CDOUTIMVU = 0 And meYCDOUTI0.CDOUTIMDI <> 0 Then   ' (PDIF)
                       lstYCDOUTI0_Courrier.Selected(10) = True              ' NCONFORME_BED_Pdif
                    End If
                End If
End Select


End Sub


Public Sub cmdSelect_MOD()

' Chargement de la fenêtre des lettres disponibles en MODIFICATION
lstOptions_Load_MOD

' >>>>>>  MODIFICATION : Cocher les cases automatiquement
If meYCDOMOD0.CDOMODDOS = 0 Then Exit Sub

W_BOO_Autre = "O"
Select Case meYCDODOS0.CDODOSBEC     ' O=FCB / N=FCO

'   Frais charge bénficiaire
    Case "O":
              If meYCDODOS0.CDODOSMON < meYCDOMOD0.CDOMODMON Then         ' Mnt diminution
                  lstOptions.Selected(0) = True    'MOD_19_Diminution
                  W_BOO_Autre = "N"
              End If
              If meYCDODOS0.CDODOSVAL > meYCDOMOD0.CDOMODVAL Then         ' Prolongement validité
                  If meYCDODOS0.CDODOSDLE > meYCDOMOD0.CDOMODDLE Then     ' Prolongement date embarquement
                      lstOptions.Selected(2) = True                       ' MOD_FCB_02_ValProrogée_Emb
                      W_BOO_Autre = "N"
                  Else
                      lstOptions.Selected(1) = True                       ' MOD_FCB_01_ValProrogée
                      W_BOO_Autre = "N"
                  End If
              End If
              If meYCDODOS0.CDODOSVAL < meYCDOMOD0.CDOMODVAL Then         ' Raccourcir date validité
                  If meYCDODOS0.CDODOSDLE < meYCDOMOD0.CDOMODDLE Then     ' Raccourcir date embarquement
                      lstOptions.Selected(5) = True                       ' MOD_FCB_10_ValRaccourcie_Emb
                      W_BOO_Autre = "N"
                  Else
                      lstOptions.Selected(4) = True                       ' MOD_FCB_09_ValRaccourcie
                      W_BOO_Autre = "N"
                  End If
              End If
              If meYCDODOS0.CDODOSCON = "N" And meYCDOMOD0.CDOMODCON = "C" Then    ' CNF en NOT
                  lstOptions.Selected(6) = True                           ' MOD_FCB_11_CNF_En_NOT
                  W_BOO_Autre = "N"
              End If
              If meYCDODOS0.CDODOSCON = "C" And meYCDOMOD0.CDOMODCON = "N" Then    ' NOT en CNF
                  lstOptions.Selected(7) = True                           ' MOD_FCB_13_NOT_En_CNF
                  W_BOO_Autre = "N"
              End If
              If meYCDODOS0.CDODOSMON > meYCDOMOD0.CDOMODMON Then         ' Mnt augmenté
                  If meYCDODOS0.CDODOSMOC > meYCDOMOD0.CDOMODMOC Then     ' Partie CNF
                      lstOptions.Selected(8) = True                       ' MOD_FCB_17_Augmentation_CNF
                      W_BOO_Autre = "N"
                  Else
                      lstOptions.Selected(9) = True                       ' MOD_FCB_20_Augmentation_NOT
                      W_BOO_Autre = "N"
                  End If
              End If
              If W_BOO_Autre = "O" Then lstOptions.Selected(3) = True     ' MOD_FCB_03_Annexe

'   Frais charge donneur d'ordre
    Case "N":
              If meYCDODOS0.CDODOSMON < meYCDOMOD0.CDOMODMON Then         ' Mnt diminution
                  lstOptions.Selected(0) = True                           ' MOD_19_Diminution
                  W_BOO_Autre = "N"
              End If
              If meYCDODOS0.CDODOSVAL > meYCDOMOD0.CDOMODVAL Then         ' Prolongement validité
                  If meYCDODOS0.CDODOSDLE > meYCDOMOD0.CDOMODDLE Then     ' Prolongement date embarquement
                      lstOptions.Selected(11) = True                      ' MOD_FCO_05_ValProrogée_Emb
                      W_BOO_Autre = "N"
                  Else
                      lstOptions.Selected(10) = True                      ' MOD_FCO_04_ValProrogée
                      W_BOO_Autre = "N"
                  End If
              End If
              If meYCDODOS0.CDODOSVAL < meYCDOMOD0.CDOMODVAL Then         ' Raccourcir date validité
                  If meYCDODOS0.CDODOSDLE < meYCDOMOD0.CDOMODDLE Then     ' Raccourcir date embarquement
                      lstOptions.Selected(14) = True                      ' MOD_FCO_08_ValRaccourcie_Emb
                      W_BOO_Autre = "N"
                  Else
                      lstOptions.Selected(13) = True                      ' MOD_FCO_07_ValRaccourcie
                      W_BOO_Autre = "N"
                  End If
              End If
              If meYCDODOS0.CDODOSCON = "N" And meYCDOMOD0.CDOMODCON = "C" Then    ' CNF en NOT
                  lstOptions.Selected(15) = True                          ' MOD_FCO_12_CNF_En_NOT
                  W_BOO_Autre = "N"
              End If
              If meYCDODOS0.CDODOSCON = "C" And meYCDOMOD0.CDOMODCON = "N" Then    ' NOT en CNF
                  lstOptions.Selected(16) = True                          ' MOD_FCO_14_NOT_En_CNF
                  W_BOO_Autre = "N"
              End If
              If meYCDODOS0.CDODOSMON > meYCDOMOD0.CDOMODMON Then         ' Mnt augmenté
                  lstOptions.Selected(17) = True                          ' MOD_FCO_18_Augmentation
                  W_BOO_Autre = "N"
              End If
              If W_BOO_Autre = "O" Then lstOptions.Selected(12) = True    ' MOD_FCO_06_Annexe

End Select

End Sub

Public Sub fraYCDOUTI0_Display_Reset()
Dim I As Integer, blnOk As Boolean
Dim wYCDODES0 As String, blnYCDODES0  As Boolean
Dim X As String
Dim wDC As String
Dim K As Integer

' Utilisation

    fgYCDOUTI0.Col = 1
    K = Val(fgYCDOUTI0.Text)
    meYCDOUTI0 = meYBIACDO.YCDOUTI0(K)

meCDO_Courrier.YCDOUTI0 = meYCDOUTI0
cmdSelect_UTI

blnYCDOUTI0_Ok = True

fraYCDOUTI0_Display.Visible = True
fgYCDOUTI0_Document_Select.Clear
fgYCDOUTI0_Document_Select.FormatString = fgYCDOUTI0_Document_Select_FormatString
fgYCDOUTI0_Document_Select.Rows = 6

fgYCDOUTI0_Document_Select.Col = 0
fgYCDOUTI0_Document_Select.Row = 1: fgYCDOUTI0_Document_Select.Text = "FACTURE"
fgYCDOUTI0_Document_Select.Row = 2: fgYCDOUTI0_Document_Select.Text = "CONNAISSEMENT"
fgYCDOUTI0_Document_Select.Row = 3: fgYCDOUTI0_Document_Select.Text = "LETTRE DE TRANSPORT AERIEN"
fgYCDOUTI0_Document_Select.Row = 4: fgYCDOUTI0_Document_Select.Text = "NOTE DE POIDS"
fgYCDOUTI0_Document_Select.Row = 5: fgYCDOUTI0_Document_Select.Text = "LISTE DE COLISAGE"


txtYCDOUTI0_ATT = ""
txtYCDOUTI0_CrrGB_Mnt = ""
txtYCDOUTI0_CrrGB_Tx = ""
txtYCDOUTI0_CrrGB_BqRbt = ""
txtYCDOUTI0_IBAN = ""
txtYCDOUTI0_YCDODES0 = ""
txtYCDOUTI0_DELAI = ""
wYCDODES0 = ""
wYCDOIRR0 = ""

meCDO_Courrier.YCDOREG0_R_Nb = 0
rsZCDOREG0_Init meCDO_Courrier.YCDOREG0_R
meCDO_Courrier.YCDOREG0_C_Nb = 0
meCDO_Courrier.YCDOREG0_C = meCDO_Courrier.YCDOREG0_R
meCDO_Courrier.YCDOREG0_D_Nb = 0
meCDO_Courrier.YCDOREG0_D = meCDO_Courrier.YCDOREG0_R

meCDO_Courrier.Com_Nb = 0
ReDim meCDO_Courrier.CDOCOMCOM(1)
ReDim meCDO_Courrier.CDOCOMDEV(1)
ReDim meCDO_Courrier.CDOREGCRD(1)
ReDim meCDO_Courrier.CDOCOMMON(1)
ReDim meCDO_Courrier.CDOCOMMTV(1)

cmdYCDOUTI0_Document_Quit_Click
lstYCDOUTI0_Document.ListIndex = 0
blnOk = False

' IBAN
For I = 1 To meYBIACDO.YCDOSWI0_Nb
    xYCDOSWI0 = meYBIACDO.YCDOSWI0(I)
                
    If meYCDOUTI0.CDOUTIUTI = xYCDOSWI0.CDOSWIUTI Then
        If Not blnOk Then
            X = Trim(xYCDOSWI0.CDOSWIIBE)
            If X <> "" Then
                blnOk = True
                txtYCDOUTI0_IBAN = Format$(X, "&&&& &&&& &&&& &&&& &&&& &&&& &&&&!")
            End If
        End If
    End If
 Next I
 
' Marchandises
blnYCDODES0 = False
For I = 1 To meYBIACDO.YCDODES0_Nb
    xYCDODES0 = meYBIACDO.YCDODES0(I)
                
    If blnYCDODES0 Then
          wYCDODES0 = wYCDODES0 & Asc13 & Asc10 & Trim(xYCDODES0.CDODESTEX)
    Else
          wYCDODES0 = Trim(xYCDODES0.CDODESTEX)
          blnYCDODES0 = True
    End If
Next I

' Irrégularités
blnYCDOIRR0 = False
For I = 1 To meYBIACDO.YCDOIRR0_Nb
     xYCDOIRR0 = meYBIACDO.YCDOIRR0(I)
     If meYCDOUTI0.CDOUTIUTI = xYCDOIRR0.CDOIRRUTI Then
        If blnYCDOIRR0 Then
            wYCDOIRR0 = wYCDOIRR0 & Asc13 & Asc10 & Trim(xYCDOIRR0.CDOIRRTEX)
         Else
            wYCDOIRR0 = Trim(xYCDOIRR0.CDOIRRTEX)
            blnYCDOIRR0 = True
         End If
     End If
Next I

' Réglements  Restitution garantie, D bordereau envoi à la bq émettrice, C accusé réception au bénéficiaire
For I = 1 To meYBIACDO.YCDOREG0_Nb
    xYCDOREG0 = meYBIACDO.YCDOREG0(I)
    If meYCDOUTI0.CDOUTIUTI = xYCDOREG0.CDOREGUTI Then
        Select Case xYCDOREG0.CDOREGCRD
            Case "C": meCDO_Courrier.YCDOREG0_C_Nb = meCDO_Courrier.YCDOREG0_C_Nb + 1: meCDO_Courrier.YCDOREG0_C = xYCDOREG0
            Case "D": meCDO_Courrier.YCDOREG0_D_Nb = meCDO_Courrier.YCDOREG0_D_Nb + 1: meCDO_Courrier.YCDOREG0_D = xYCDOREG0
            Case "R": meCDO_Courrier.YCDOREG0_R_Nb = meCDO_Courrier.YCDOREG0_R_Nb + 1: meCDO_Courrier.YCDOREG0_R = xYCDOREG0
            Case Else: blnYCDOUTI0_Ok = False
                MsgBox " Code CDOREGCRD non géré : " & xYCDOREG0.CDOREGUTI, vbExclamation, "fraYCDOUTI0_Display_Reset"
        End Select
    End If
Next I

txtYCDOUTI0_YCDODES0 = wYCDODES0

txtYCDOUTI0_DELAI = "Documents présentés dans les délais (Art.43 des RUU) et dans la validité du crédit."

If meCDO_Courrier.YCDOREG0_C_Nb > 1 Or meCDO_Courrier.YCDOREG0_D_Nb > 1 Or meCDO_Courrier.YCDOREG0_R_Nb > 1 Then
   MsgBox "Nb réglement R/C/D > 1 : ", vbExclamation, "fraYCDOUTI0_Display_Reset"
   blnYCDOUTI0_Ok = False
End If


'*********** Commissions : cumul par ( code / devise / sens ) du montant Ht et de la TVA
For I = 1 To meYBIACDO.YCDOCOM0_Nb
    xYCDOCOM0 = meYBIACDO.YCDOCOM0(I)
                
        If meYCDOUTI0.CDOUTIUTI = xYCDOCOM0.CDOCOMUTR Then
            If xYCDOCOM0.CDOCOMREG <> 0 And xYCDOCOM0.CDOCOMVAL <> 0 Then
                Select Case xYCDOCOM0.CDOCOMNRE
                    Case meCDO_Courrier.YCDOREG0_C.CDOREGREG: wDC = "C"
                    Case meCDO_Courrier.YCDOREG0_D.CDOREGREG: wDC = "D"
                    Case Else: blnYCDOUTI0_Ok = False
                                MsgBox " xYCDOCOM0.CDOCOMNRE non identifié : " & xYCDOCOM0.CDOCOMNRE, vbExclamation, "fraYCDOUTI0_Display_Reset"
                End Select
                
                blnOk = False
                For K = 1 To meCDO_Courrier.Com_Nb
                        If meCDO_Courrier.CDOCOMCOM(K) = xYCDOCOM0.CDOCOMCOM _
                        And meCDO_Courrier.CDOCOMDEV(K) = xYCDOCOM0.CDOCOMDEV _
                        And meCDO_Courrier.CDOREGCRD(K) = wDC Then
                            blnOk = True
                            meCDO_Courrier.CDOCOMMON(K) = meCDO_Courrier.CDOCOMMON(K) + xYCDOCOM0.CDOCOMMON
                            meCDO_Courrier.CDOCOMMTV(K) = meCDO_Courrier.CDOCOMMTV(K) + xYCDOCOM0.CDOCOMMTV
                        End If
                Next K
                
                If Not blnOk Then
                        meCDO_Courrier.Com_Nb = meCDO_Courrier.Com_Nb + 1
                        ReDim Preserve meCDO_Courrier.CDOCOMCOM(meCDO_Courrier.Com_Nb)
                        ReDim Preserve meCDO_Courrier.CDOCOMDEV(meCDO_Courrier.Com_Nb)
                        ReDim Preserve meCDO_Courrier.CDOREGCRD(meCDO_Courrier.Com_Nb)
                        ReDim Preserve meCDO_Courrier.CDOCOMMON(meCDO_Courrier.Com_Nb)
                        ReDim Preserve meCDO_Courrier.CDOCOMMTV(meCDO_Courrier.Com_Nb)
                        
                        meCDO_Courrier.CDOCOMCOM(meCDO_Courrier.Com_Nb) = xYCDOCOM0.CDOCOMCOM
                        meCDO_Courrier.CDOCOMDEV(meCDO_Courrier.Com_Nb) = xYCDOCOM0.CDOCOMDEV
                        meCDO_Courrier.CDOREGCRD(meCDO_Courrier.Com_Nb) = wDC
                        meCDO_Courrier.CDOCOMMON(meCDO_Courrier.Com_Nb) = xYCDOCOM0.CDOCOMMON
                        meCDO_Courrier.CDOCOMMTV(meCDO_Courrier.Com_Nb) = xYCDOCOM0.CDOCOMMTV
                End If
            End If
        End If
Next I

End Sub

Private Sub txtYCDOUTI0_Document_GotFocus()
Call txt_GotFocus(txtYCDOUTI0_Document)

End Sub


Private Sub txtYCDOUTI0_Document_Jeu1_GotFocus()
Call txt_GotFocus(txtYCDOUTI0_Document_Jeu1)

End Sub


Private Sub txtYCDOUTI0_Document_Jeu1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtYCDOUTI0_Document_Jeu1_LostFocus()
Call txt_LostFocus(txtYCDOUTI0_Document_Jeu1)

End Sub


Private Sub txtYCDOUTI0_Document_Jeu2_GotFocus()
Call txt_GotFocus(txtYCDOUTI0_Document_Jeu2)

End Sub


Private Sub txtYCDOUTI0_Document_Jeu2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtYCDOUTI0_Document_Jeu2_LostFocus()
Call txt_LostFocus(txtYCDOUTI0_Document_Jeu2)

End Sub


Private Sub txtYCDOUTI0_Document_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtYCDOUTI0_Document_LostFocus()
Call txt_LostFocus(txtYCDOUTI0_Document)

End Sub


Private Sub txtYCDOUTI0_Exp_A_GotFocus()
Call txt_GotFocus(txtYCDOUTI0_Exp_A)

End Sub


Private Sub txtYCDOUTI0_Exp_A_KeyPress(KeyAscii As Integer)
' KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtYCDOUTI0_Exp_A_LostFocus()
Call txt_LostFocus(txtYCDOUTI0_Exp_A)
txtYCDOUTI0_Exp_A.ForeColor = vbMagenta

End Sub


Private Sub txtYCDOUTI0_Exp_De_GotFocus()
Call txt_GotFocus(txtYCDOUTI0_Exp_De)

End Sub


Private Sub txtYCDOUTI0_Exp_De_KeyPress(KeyAscii As Integer)
'KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtYCDOUTI0_Exp_De_LostFocus()
Call txt_LostFocus(txtYCDOUTI0_Exp_De)
txtYCDOUTI0_Exp_De.ForeColor = vbMagenta

End Sub


Private Sub txtYCDOUTI0_Exp_Le_GotFocus()
Call txt_GotFocus(txtYCDOUTI0_Exp_Le)

End Sub


Private Sub txtYCDOUTI0_Exp_Le_KeyPress(KeyAscii As Integer)
' KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtYCDOUTI0_Exp_Le_LostFocus()
Call txt_LostFocus(txtYCDOUTI0_Exp_Le)
txtYCDOUTI0_Exp_Le.ForeColor = vbMagenta

End Sub


Private Sub txtYCDOUTI0_Exp_Par_GotFocus()
Call txt_GotFocus(txtYCDOUTI0_Exp_Par)

End Sub


Private Sub txtYCDOUTI0_Exp_Par_KeyPress(KeyAscii As Integer)
'KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtYCDOUTI0_Exp_Par_LostFocus()
Call txt_LostFocus(txtYCDOUTI0_Exp_Par)
txtYCDOUTI0_Exp_Par.ForeColor = vbMagenta

End Sub


Private Sub txtYCDOUTI0_IBAN_GotFocus()
Call txt_GotFocus(txtYCDOUTI0_IBAN)

End Sub

Private Sub txtYCDOUTI0_IBAN_KeyPress(KeyAscii As Integer)
'$20070829 $JPL KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtYCDOUTI0_IBAN_LostFocus()
Call txt_LostFocus(txtYCDOUTI0_IBAN)
txtYCDOUTI0_IBAN.ForeColor = vbMagenta

End Sub

Private Sub txtYCDOUTI0_YCDODES0_GotFocus()
Call txt_GotFocus(txtYCDOUTI0_YCDODES0)

End Sub


Private Sub txtYCDOUTI0_YCDODES0_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtYCDOUTI0_YCDODES0_LostFocus()
Call txt_LostFocus(txtYCDOUTI0_YCDODES0)
txtYCDOUTI0_YCDODES0.ForeColor = vbMagenta

End Sub

Private Sub txtDossier_Garantie_GotFocus()
Call txt_GotFocus(txtDossier_Garantie)

End Sub


Private Sub txtDossier_Garantie_LostFocus()
Call txt_LostFocus(txtDossier_Garantie)
txtYCDOUTI0_YCDODES0.ForeColor = vbMagenta

End Sub



Public Sub fgDossier_Display_YBIACDO()
Dim X As String
Dim I As Integer

X = CStr(meYBIACDO.YCDODOS0(1).CDODOSDOS) & " " & meYBIACDO.YCDODOS0(1).CDODOSCOR
Call fgDossier_DisplayLine("ZCDODOS0", "1", X)

For I = 1 To meYBIACDO.YCDOMOD0_Nb
    X = CStr(meYBIACDO.YCDOMOD0(I).CDOMODDOS) & "_" & meYBIACDO.YCDOMOD0(I).CDOMODNMO
    Call fgDossier_DisplayLine("ZCDOMOD0", CStr(I), X)
Next I

If meYBIACDO.YCDOTIE0_Nb > 0 Then
    X = meYBIACDO.YCDOTIE0(1).CDOTIETIE & " " & meYBIACDO.YCDOTIE0(1).CDOTIERA1
    Call fgDossier_DisplayLine("ZCDOTIE0", "1", X)
End If


For I = 1 To meYBIACDO.YCDOUTI0_Nb
    If meYBIACDO.YCDOUTI0(I).CDOUTIDCO = "N" Then
        X = meYBIACDO.YCDOUTI0(I).CDOUTIUTI & " NON CONFORME  "
    Else
        X = meYBIACDO.YCDOUTI0(I).CDOUTIUTI & " conforme "
    End If
    
    X = X _
        & "   Date : " & dateIBM10(meYBIACDO.YCDOUTI0(I).CDOUTIPRE, True) _
        & "   Montant : " & Format$(meYBIACDO.YCDOUTI0(I).CDOUTIMON, "### ### ### ##0.00") _
        & "   Réf : " & meYBIACDO.YCDOUTI0(I).CDOUTIRER
       
    Call fgDossier_DisplayLine("ZCDOUTI0", CStr(I), X)
    Call fgYCDOUTI0_DisplayLine("ZCDOUTI0", CStr(I), X)
Next I

For I = 1 To meYBIACDO.YCDOREG0_Nb
    X = CStr(meYBIACDO.YCDOREG0(I).CDOREGDOS) & "_" & meYBIACDO.YCDOREG0(I).CDOREGUTI & "_" & meYBIACDO.YCDOREG0(I).CDOREGPAI
    Call fgDossier_DisplayLine("ZCDOREG0", CStr(I), X)
Next I

For I = 1 To meYBIACDO.YCDODES0_Nb
    X = CStr(meYBIACDO.YCDODES0(I).CDODESDOS) & "_" & meYBIACDO.YCDODES0(I).CDODESUTI & "_" & meYBIACDO.YCDODES0(I).CDODESSEQ
    Call fgDossier_DisplayLine("ZCDODES0", CStr(I), X)
Next I

For I = 1 To meYBIACDO.YCDOSWI0_Nb
    X = CStr(meYBIACDO.YCDOSWI0(I).CDOSWIDOS) & "_" & meYBIACDO.YCDOSWI0(I).CDOSWIUTI & "_" & meYBIACDO.YCDOSWI0(I).CDOSWIPAI
    Call fgDossier_DisplayLine("ZCDOSWI0", CStr(I), X)
Next I

For I = 1 To meYBIACDO.YCDOIRR0_Nb
    X = CStr(meYBIACDO.YCDOIRR0(I).CDOIRRDOS) & "_" & meYBIACDO.YCDOIRR0(I).CDOIRRUTI & "_" & meYBIACDO.YCDOIRR0(I).CDOIRRSEQ
    Call fgDossier_DisplayLine("ZCDOIRR0", CStr(I), X)
Next I

For I = 1 To meYBIACDO.YCDOCOM0_Nb
    X = CStr(meYBIACDO.YCDOCOM0(I).CDOCOMDOS) & "_" & meYBIACDO.YCDOCOM0(I).CDOCOMSEQ
    Call fgDossier_DisplayLine("ZCDOCOM0", CStr(I), X)
Next I

For I = 1 To meYBIACDO.YCDOCO20_Nb
    X = CStr(meYBIACDO.YCDOCO20(I).CDOCO2DOS) & "_" & meYBIACDO.YCDOCO20(I).CDOCO2SEQ
    Call fgDossier_DisplayLine("ZCDOCO20", CStr(I), X)
Next I

For I = 1 To meYBIACDO.YCDOTC20_Nb
    X = CStr(meYBIACDO.YCDOTC20(I).CDOTC2DOS) & "_" & meYBIACDO.YCDOTC20(I).CDOTC2SEQ
    Call fgDossier_DisplayLine("ZCDOTC20", CStr(I), X)
Next I

End Sub
