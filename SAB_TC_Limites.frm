VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAB_TC_Limites 
   AutoRedraw      =   -1  'True
   Caption         =   "TC : Limites trésorerie PRE / EMP"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SAB_TC_Limites.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10680
   ScaleWidth      =   13530
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
      TabIndex        =   4
      Top             =   0
      Width           =   5175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10170
      Left            =   0
      TabIndex        =   2
      Top             =   510
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   17939
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Rechercher"
      TabPicture(0)   =   "SAB_TC_Limites.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "ajustement des cours Fixing / BID"
      TabPicture(1)   =   "SAB_TC_Limites.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraEUR"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Calendrier"
      TabPicture(2)   =   "SAB_TC_Limites.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblDevF_AMJ"
      Tab(2).Control(1)=   "txtDevF_AMJ"
      Tab(2).Control(2)=   "fgDevF"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Paramétrage"
      TabPicture(3)   =   "SAB_TC_Limites.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblParamX"
      Tab(3).Control(1)=   "lstParam"
      Tab(3).Control(2)=   "fraParam"
      Tab(3).ControlCount=   3
      Begin VB.Frame fraParam 
         BackColor       =   &H00C0E0FF&
         Height          =   2580
         Left            =   -68310
         TabIndex        =   49
         Top             =   1650
         Width           =   5172
         Begin VB.TextBox txtParam 
            Height          =   288
            Left            =   3150
            MaxLength       =   10
            TabIndex        =   53
            Top             =   600
            Width           =   765
         End
         Begin VB.CommandButton cmdParam_Delete 
            BackColor       =   &H00FF80FF&
            Caption         =   "Supprimer"
            Height          =   480
            Left            =   3765
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   1425
            Width           =   900
         End
         Begin VB.CommandButton cmdParam_Add 
            BackColor       =   &H000080FF&
            Caption         =   "Ajouter"
            Height          =   480
            Left            =   2250
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   1440
            Width           =   900
         End
         Begin VB.CommandButton cmdParam_Quit 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Abandonner"
            Height          =   480
            Left            =   765
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   1440
            Width           =   990
         End
         Begin VB.Label lblParam 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Code AUT (3 car ) à exclure"
            Height          =   255
            Left            =   315
            TabIndex        =   54
            Top             =   585
            Width           =   2535
         End
      End
      Begin VB.ListBox lstParam 
         Height          =   7485
         Left            =   -74175
         TabIndex        =   48
         Top             =   1395
         Width           =   3255
      End
      Begin VB.Frame fraEUR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7812
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   12972
         Begin VB.Frame fraEURBID 
            BackColor       =   &H00E8FFFE&
            Caption         =   "Modification du cours BID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3132
            Left            =   7440
            TabIndex        =   29
            Top             =   1560
            Visible         =   0   'False
            Width           =   4092
            Begin VB.CommandButton cmdEURBID_Quit 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Abandonner"
               Height          =   525
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   2400
               Width           =   1212
            End
            Begin VB.CommandButton cmdEURBID_Ok 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Enregistrer"
               Height          =   525
               Left            =   2400
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   2400
               Width           =   1332
            End
            Begin VB.TextBox txtEURBID 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   372
               Left            =   1920
               TabIndex        =   32
               Text            =   "BID"
               Top             =   1440
               Width           =   1692
            End
            Begin VB.Label lblEURFIX 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E8FFFE&
               Caption         =   "fixing"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   372
               Left            =   1920
               TabIndex        =   31
               Top             =   720
               Width           =   1692
            End
            Begin VB.Label lblEUR 
               BackColor       =   &H00E8FFFE&
               Caption         =   "Label1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   372
               Left            =   480
               TabIndex        =   30
               Top             =   840
               Width           =   972
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgEUR 
            Height          =   6735
            Left            =   1080
            TabIndex        =   28
            Top             =   915
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   11880
            _Version        =   393216
            Cols            =   3
            ScrollBars      =   2
            FormatString    =   "Devise        |>    Fixing           |>             BID         "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComCtl2.DTPicker txtSelect_Amj1 
            Height          =   300
            Left            =   8640
            TabIndex        =   36
            Top             =   480
            Visible         =   0   'False
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
            Format          =   32636931
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblDevF 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "date du cours"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   1800
            TabIndex        =   41
            Top             =   360
            Width           =   2892
         End
         Begin VB.Label lblSelect_Amj1 
            Caption         =   "Date arrêté J+1"
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
            Left            =   7200
            TabIndex        =   35
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
         End
      End
      Begin VB.Frame fraTab0 
         Height          =   9600
         Left            =   135
         TabIndex        =   3
         Top             =   450
         Visible         =   0   'False
         Width           =   13290
         Begin VB.Frame fraDétail 
            Appearance      =   0  'Flat
            BackColor       =   &H00E8FFFE&
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   7665
            Left            =   90
            TabIndex        =   18
            Top             =   2895
            Visible         =   0   'False
            Width           =   11250
            Begin VB.Frame fraParam_PCT 
               BackColor       =   &H00FFE0FF&
               Caption         =   "Durée maximale des PCT"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   1335
               Left            =   5730
               TabIndex        =   42
               Top             =   240
               Width           =   3900
               Begin VB.CommandButton cmdParam_PCT_Delete 
                  BackColor       =   &H000000FF&
                  Caption         =   "Supprimer"
                  Height          =   525
                  Left            =   2415
                  Style           =   1  'Graphical
                  TabIndex        =   47
                  Top             =   675
                  Width           =   1320
               End
               Begin VB.CommandButton cmdParam_PCT_Update 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Enregistrer"
                  Height          =   525
                  Left            =   2415
                  Style           =   1  'Graphical
                  TabIndex        =   46
                  Top             =   150
                  Width           =   1332
               End
               Begin VB.OptionButton optParam_PCT_J 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "Jour"
                  Height          =   255
                  Left            =   800
                  TabIndex        =   45
                  Top             =   840
                  Width           =   735
               End
               Begin VB.OptionButton optParam_PCT_M 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "Mois"
                  Height          =   225
                  Left            =   810
                  TabIndex        =   44
                  Top             =   450
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.TextBox txtParam_PCT 
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
                  Height          =   285
                  Left            =   180
                  MaxLength       =   2
                  TabIndex        =   43
                  Top             =   450
                  Width           =   495
               End
            End
            Begin VB.CommandButton cmdSelect_NOk 
               BackColor       =   &H00C0C0FF&
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   795
               Left            =   10005
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   570
               Width           =   975
            End
            Begin MSFlexGridLib.MSFlexGrid fgDossier 
               Height          =   3885
               Left            =   135
               TabIndex        =   19
               Top             =   3315
               Width           =   11040
               _ExtentX        =   19473
               _ExtentY        =   6853
               _Version        =   393216
               Rows            =   1
               Cols            =   10
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16449535
               ForeColor       =   8388608
               BackColorFixed  =   12632064
               ForeColorFixed  =   16777215
               BackColorSel    =   12648384
               BackColorBkg    =   16449535
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
               GridLinesFixed  =   1
               ScrollBars      =   2
               AllowUserResizing=   3
               Appearance      =   0
               FormatString    =   $"SAB_TC_Limites.frx":037A
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
            Begin MSFlexGridLib.MSFlexGrid fgTotal 
               Height          =   1605
               Left            =   120
               TabIndex        =   20
               Top             =   1680
               Width           =   11025
               _ExtentX        =   19447
               _ExtentY        =   2831
               _Version        =   393216
               Rows            =   1
               Cols            =   9
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16449535
               ForeColor       =   12582912
               BackColorFixed  =   12632064
               ForeColorFixed  =   16777215
               BackColorSel    =   12648384
               BackColorBkg    =   16449535
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
               GridLinesFixed  =   1
               ScrollBars      =   2
               AllowUserResizing=   3
               Appearance      =   0
               FormatString    =   $"SAB_TC_Limites.frx":0439
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
            Begin VB.Label lblSelect_Client 
               Alignment       =   2  'Center
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblSelect_Client"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Left            =   165
               TabIndex        =   25
               Top             =   570
               Width           =   5310
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame fraSelect_Options 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1485
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   11610
            Begin VB.TextBox txtSelect_EnCours 
               Height          =   330
               Left            =   6855
               TabIndex        =   58
               Text            =   "6"
               Top             =   930
               Width           =   420
            End
            Begin VB.CheckBox chkSelect_EnCours 
               Alignment       =   1  'Right Justify
               Caption         =   "afficher toutes les autorisations"
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
               Left            =   4530
               TabIndex        =   22
               Top             =   225
               Value           =   1  'Checked
               Width           =   2835
            End
            Begin VB.CheckBox chkSelect_AmjMin 
               Caption         =   "date d'arrêté de l'encours"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   8055
               TabIndex        =   21
               Top             =   315
               Width           =   2160
            End
            Begin VB.ComboBox cboSelect_TREOPEAUT 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2955
               Sorted          =   -1  'True
               TabIndex        =   17
               Text            =   "AUT"
               Top             =   915
               Width           =   1300
            End
            Begin VB.ComboBox cboSelect_TREOPEOPR 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2895
               Sorted          =   -1  'True
               TabIndex        =   15
               Text            =   "OPE"
               Top             =   270
               Width           =   1300
            End
            Begin VB.ComboBox cboSelect_TREOPEDEV 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   705
               Sorted          =   -1  'True
               TabIndex        =   12
               Text            =   "DEV"
               Top             =   885
               Width           =   1095
            End
            Begin VB.TextBox txtSelect_TREOPECLI 
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
               Left            =   675
               TabIndex        =   9
               Top             =   300
               Width           =   1125
            End
            Begin MSComCtl2.DTPicker txtSelect_AmjMin 
               Height          =   300
               Left            =   10320
               TabIndex        =   8
               Top             =   345
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   8.25
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
               Format          =   32636931
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtSelect_AmjMax 
               Height          =   300
               Left            =   10290
               TabIndex        =   23
               Top             =   885
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   8.25
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
               Format          =   32636931
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblSelect_EnCours_B 
               Caption         =   "mois"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   7335
               TabIndex        =   59
               Top             =   915
               Width           =   420
            End
            Begin VB.Label lblSelect_EnCours_A 
               Caption         =   "Exclure les racines sans encours dont l'autorisation est échue depuis :"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   4530
               TabIndex        =   57
               Top             =   735
               Width           =   2805
            End
            Begin VB.Label lblSelect_AmjMax 
               Caption         =   "Date arrêté J+2 "
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
               Left            =   8310
               TabIndex        =   24
               Top             =   795
               Width           =   1575
            End
            Begin VB.Label lblSelect_TREOPEAUT 
               Caption         =   "Autorisation"
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
               Left            =   1995
               TabIndex        =   16
               Top             =   945
               Width           =   915
            End
            Begin VB.Label lblSelect_TREOPEOPR 
               Caption         =   "Opération"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1995
               TabIndex        =   14
               Top             =   360
               Width           =   750
            End
            Begin VB.Label lblSelect_TREOPEDEV 
               Caption         =   "Devise"
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
               Left            =   120
               TabIndex        =   13
               Top             =   960
               Width           =   600
            End
            Begin VB.Label lblSelect_TREOPECLI 
               Caption         =   "Client"
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
               Left            =   120
               TabIndex        =   10
               Top             =   330
               Width           =   645
            End
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Options"
            BeginProperty Font 
               Name            =   "Calibri"
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
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7665
            Left            =   120
            TabIndex        =   5
            Top             =   1680
            Width           =   8880
            _ExtentX        =   15663
            _ExtentY        =   13520
            _Version        =   393216
            Rows            =   1
            Cols            =   9
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   16449535
            ForeColor       =   16711680
            BackColorFixed  =   8421376
            ForeColorFixed  =   16777215
            BackColorSel    =   12648384
            BackColorBkg    =   16449535
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            ScrollBars      =   2
            AllowUserResizing=   3
            Appearance      =   0
            FormatString    =   $"SAB_TC_Limites.frx":04E2
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
         Begin VB.Label lblECH_Warning 
            BackColor       =   &H00E8FFFE&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Opérations échus non comptabilisées"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3135
            Left            =   9225
            TabIndex        =   56
            Top             =   5295
            Width           =   3855
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblDevF_Warning 
            BackColor       =   &H00E8FFFE&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "jours fériés"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3135
            Left            =   9150
            TabIndex        =   40
            Top             =   1650
            Width           =   3855
            WordWrap        =   -1  'True
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgDevF 
         Height          =   7548
         Left            =   -74880
         TabIndex        =   37
         Top             =   840
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   13309
         _Version        =   393216
         Rows            =   1
         Cols            =   33
         RowHeightMin    =   300
         BackColor       =   16777215
         ForeColor       =   8388608
         BackColorFixed  =   16777152
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   16777210
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   3
         GridLinesFixed  =   1
         ScrollBars      =   2
         AllowUserResizing=   3
         Appearance      =   0
         FormatString    =   $"SAB_TC_Limites.frx":0573
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
      Begin MSComCtl2.DTPicker txtDevF_AMJ 
         Height          =   300
         Left            =   -73440
         TabIndex        =   38
         Top             =   480
         Width           =   1212
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   12632256
         CustomFormat    =   "dd  MM yyy"
         Format          =   32636931
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   -328351
      End
      Begin VB.Label lblParamX 
         Alignment       =   2  'Center
         BackColor       =   &H00FF80FF&
         Caption         =   "Code AUT exclus"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -74145
         TabIndex        =   55
         Top             =   780
         Width           =   3180
      End
      Begin VB.Label lblDevF_AMJ 
         Caption         =   "à partir du"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   -74520
         TabIndex        =   39
         Top             =   480
         Width           =   972
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
      Picture         =   "SAB_TC_Limites.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
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
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   4905
      WordWrap        =   -1  'True
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
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Print_Liste 
         Caption         =   "Imprimer liste"
      End
      Begin VB.Menu mnuSelect_Print_Détail 
         Caption         =   "Imprimer liste détaillée"
      End
   End
End
Attribute VB_Name = "frmSAB_TC_Limites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'!!!!!!! CODES AUTORISATIONS : "PIB" pour PRET / "EMB" pour EMP !!!!!
'---------------------------------------------------------

Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim SAB_TC_Limites_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency
Dim blnAuto As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim rsSAB_X As New ADODB.Recordset

Dim fgDossier_FormatString As String, fgDossier_K As Integer
Dim fgDossier_RowDisplay As Integer, fgDossier_RowClick As Integer, fgDossier_ColClick As Integer
Dim fgDossier_ColorClick As Long, fgDossier_ColorDisplay As Long
Dim fgDossier_Sort1 As Integer, fgDossier_Sort2 As Integer
Dim fgDossier_SortAD As Integer, fgDossier_Sort1_Old As Integer
Dim fgDossier_arrIndex As Integer
Dim blnfgDossier_DisplayLine As Boolean

Dim meZTREOPE0 As typeZTREOPE0, xZTREOPE0 As typeZTREOPE0
Dim arrZTREOPE0() As typeZTREOPE0
Dim selZTREOPE0() As typeZTREOPE0, selZTREOPE0_Nb As Long, selZTREOPE0_Max As Long
Dim meZAUTSYC0 As typeZAUTSYC0, xZAUTSYC0 As typeZAUTSYC0
Dim arrZAUTSYC0() As typeZAUTSYC0, arrZAUTSYC0_Nb As Long, arrZAUTSYC0_Max As Long
Dim selZAUTSYC0() As typeZAUTSYC0, selZAUTSYC0_Nb As Long, selZAUTSYC0_Max As Long
Dim selZAUTSYC0_Display() As Boolean, selZAUTSYC0_PIB() As Boolean
Dim mAUTSYCADR_PIB As Long

Dim Ope_MAD() As Currency, Ope_ENG() As Currency
Dim ope_M0() As Currency, ope_M1() As Currency, ope_M2() As Currency
Dim Ope_Limite As Currency, Ope_Limite_Aut As Currency, Ope_Limite_Max As Currency
Dim Ope_Limite_Max_Code As String, blnLimite_Max_Code As Boolean
Dim Ope_LimiteX As Currency
Dim Ope_PJJ As Currency, Ope_PCT As Currency, Ope_PTM As Currency


Dim arrZAUTHST0() As typeZAUTHST0, arrZAUTHST0_Nb As Long, arrZAUTHST0_Max As Long
Dim xZAUTHST0  As typeZAUTHST0
Dim blnZAUTSYC0_EnCours() As Boolean

Dim meZCLIENA0 As typeZCLIENA0, xZCLIENA0 As typeZCLIENA0
Dim arrZCLIENA0() As typeZCLIENA0, arrClient_Nb As Long, arrClient_Max As Long

Dim fgTotal_FormatString As String, fgTotal_K As Integer
Dim fgTotal_RowDisplay As Integer, fgTotal_RowClick As Integer, fgTotal_ColClick As Integer
Dim fgTotal_ColorClick As Long, fgTotal_ColorDisplay As Long
Dim fgTotal_Sort1 As Integer, fgTotal_Sort2 As Integer
Dim fgTotal_SortAD As Integer, fgTotal_Sort1_Old As Integer
Dim fgTotal_arrIndex As Integer
Dim blnfgTotal_DisplayLine As Boolean

Dim meCV1 As typeCV, meCV2 As typeCV

Dim blnId_Print As Boolean

Dim wAmj_Selection_8C As String * 8, wAmj_Selection_7C As String * 7
Dim xWhere_Selection As String
Dim wAmjMax_8C As String * 8, wAmjMax_7C As String * 7
Dim wAmj1_8C As String * 8, wAmj1_7C As String * 7

Dim wAmjM1_7C As String * 7

Dim mAMJ_Fixing As String
'_________________________________________________________________________________
Dim arrEUR() As String, arrEURFIX() As Double, arrEURBID() As Double, arrEUR_nb As Integer

Dim arrDevF_ISO() As String, arrDevF_AMJ() As Long, arrDevF_Nb As Integer, arrDevF_Max As Integer
Dim arrDevF_AMJx()

Dim fgDevF_EUR_Row As Integer, blnDevF_Init As Boolean
Dim arrDevF_Ouvré(31) As Long
Dim arrDevF_Warning(7) As String, blnDevF_Warning As Boolean

Dim arrParam_PCT_CLI() As String, arrParam_PCT_AUT() As String, arrParam_PCT_Nb As Integer
Dim blnParam_PCT_Exist As Boolean
Dim oldParam_PCT As typeYBIATAB0, newParam_PCT As typeYBIATAB0

Dim arrZCLIGRP0() As typeZCLIGRP0, arrZCLIGRP0_Nb As Long

Dim arrParam_AUT() As String, arrParam_AUT_Nb As Integer
Dim oldParam As typeYBIATAB0, newParam As typeYBIATAB0

'20110617 JPL
Dim arrJ_AMJ_7c(100) As Long, arrJ_MTE(100) As Currency, arrJ_K As Integer, arrJ_K_Max As Integer, arrJ_AUTSYCMON As Currency

Dim mAUTSYCFIN_Min As Long
Public Sub cmdPrint_Ok_xlsManual()
Dim iRow As Integer, K As Integer, I As Integer
Dim xAUTSYCAUT As String
Dim wText As String
Dim wbExcel As Excel.Workbook
Dim nbSheetRows As Long
Dim currentRow As Long
Dim maxRows As Long
Dim maxRowsPlus As Long
Dim comptageRows As Long
Dim wColor As Long

'                                               '
Call init_xlsManual
'On recopie le classeur modèle de c:\BIASRV vers c:\temp\imp_pdf
FileCopy paramFolder_Local & "\Modeles\modele_TC_LIMITES.xlsx", paramIMP_PDF_Path_Temp & "\modele_TC_LIMITES.xlsx"
'on charge CE classeur dans Excel
Call appExcelPublic.Workbooks.Open(paramIMP_PDF_Path_Temp & "\modele_TC_LIMITES.xlsx")
Set wbExcel = appExcelPublic.ActiveWorkbook
With wbExcel
    .Title = "TC_LIMITES"
    .Subject = "TC_LIMITES"
End With
'                                               '
wbExcel.Worksheets(1).Activate
currentRow = 7
comptageRows = currentRow
maxRows = 50
maxRowsPlus = 4
fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Etat : " & fgSelect.Rows - 1)
wText = " au " & dateImp10(wAmj_Selection_8C) & " / spot : " & dateImp10(wAmjMax_8C)
If chkSelect_AmjMin = "1" Then
    wText = " *** !!!! date d'arrêté au " & dateImp10(wAmj_Selection_8C) & " !!!! ***"
End If
Call prtSAB_TC_Lmites_Open_xlsManual(wText & "         ( édité le " & Now & " )", "Today", "mad " & dateImp10(wAmjMax_8C), wbExcel.Worksheets(1))
If blnDevF_Warning Then
    wColor = vbRed
    For K = 1 To 7
        If arrDevF_Warning(K) <> "" Then
            Call prtSAB_TC_Lmites_NewLine_xlsManual(currentRow, wbExcel.Worksheets(1), comptageRows, maxRows, maxRowsPlus)
            Range("A5:P5").Select
            Selection.Copy
            Range("A" & CStr(currentRow)).Select
            ActiveSheet.Paste
            wbExcel.Worksheets(1).Cells(currentRow, 10) = "férié " & arrDevF_Warning(K)
            wbExcel.Worksheets(1).Cells(currentRow, 10).Font.Color = wColor
        End If
    Next K
    Call prtSAB_TC_Lmites_NewLine_xlsManual(currentRow, wbExcel.Worksheets(1), comptageRows, maxRows, maxRowsPlus)
    Range("A5:P5").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste
End If
For iRow = 1 To fgSelect.Rows - 1
    fgSelect.Row = iRow
    fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
    xZTREOPE0 = arrZTREOPE0(K)
    xZCLIENA0 = arrZCLIENA0(K)
    selZTREOPE0_SQL arrZTREOPE0(K).TREOPECLI
    selZAUTSYC0_SQL arrZTREOPE0(K).TREOPECLI
    fgSelect.Col = 3:  xAUTSYCAUT = Mid$(Trim(fgSelect.Text), 1, 3) 'sans enCours sélectionner le code AUT
    selZAUTSYC0_Add xAUTSYCAUT
    blnId_Print = True
    For I = 0 To selZAUTSYC0_Nb
        If selZAUTSYC0_Display(I) Then
            xZAUTSYC0 = selZAUTSYC0(I)
            Call fgTotal_PrintLine_xlsManual(I, currentRow, wbExcel.Worksheets(1), comptageRows, maxRows, maxRowsPlus)
        End If
    Next I
    For I = 1 To selZTREOPE0_Nb
        xZTREOPE0 = selZTREOPE0(I)
       Call fgDossier_PrintLine_xlsManual(selZTREOPE0_Nb, currentRow, wbExcel.Worksheets(1), comptageRows, maxRows, maxRowsPlus)
    Next I
Next iRow
'on supprime les 3 lignes modèles
Rows("4:7").Select
Selection.Delete
currentRow = currentRow - 4
wbExcel.Worksheets(1).Cells(currentRow + 1, 1) = "END_OF_SHEET"
nbSheetRows = retourne_fin_de_sheet(wbExcel.Worksheets(1))
Call zoneImpression_xlsManual(wbExcel.Worksheets(1).Name, nbSheetRows, wbExcel.Worksheets(1))
Call wbExcel.Worksheets(1).ExportAsFixedFormat(xlTypePDF, paramIMP_PDF_Path & "\" & paramEditionNoPaper_Auto_PgmName & ".pdf")
'sauvegarde du fichier
Call impressions_xlsManual.prtIMP_PDF_Monitor_xlsManual
Call wbExcel.Close(True)
Set wbExcel = Nothing
Kill paramIMP_PDF_Path_Temp & "\modele_TC_LIMITES.xlsx"
fgSelect.Visible = True
Me.Show
End Sub

Public Sub fgDossier_PrintLine_xlsManual(lIndex As Long, ByRef currentRow As Long, ByRef wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim K As Integer, X As String
Dim wColor As Long
'On Error Resume Next

Call prtSAB_TC_Lmites_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
Range("A7:P7").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
wColor = vbBlue
If xZTREOPE0.TREOPEECH < wAmj_Selection_7C Then wColor = vbRed
If xZTREOPE0.TREOPEECH = wAmj_Selection_7C Then wColor = vbMagenta
If xZTREOPE0.TREOPEDIS > wAmj_Selection_7C Then wColor = RGB(0, 100, 0)
If xZCLIENA0.CLIENACLI < "0010000" Then
    wsExcel.Cells(currentRow, 2) = xZTREOPE0.TREOPECLI
    wsExcel.Cells(currentRow, 2).Font.Color = wColor
End If
wsExcel.Cells(currentRow, 3) = xZTREOPE0.TREOPEOPR & " " & xZTREOPE0.TREOPENAT & " " & xZTREOPE0.TREOPENUM
wsExcel.Cells(currentRow, 3).Font.Color = wColor
X = Format$(xZTREOPE0.TREOPEMNT, "### ### ### ###.00")
wsExcel.Cells(currentRow, 7) = X
wsExcel.Cells(currentRow, 7).Font.Color = wColor
wsExcel.Cells(currentRow, 8) = xZTREOPE0.TREOPEDEV
wsExcel.Cells(currentRow, 8).Font.Color = wColor
wsExcel.Cells(currentRow, 9) = xZTREOPE0.TREOPEAUT
wsExcel.Cells(currentRow, 9).Font.Color = wColor
wsExcel.Cells(currentRow, 4) = dateIBM10(xZTREOPE0.TREOPEDIS, True)
wsExcel.Cells(currentRow, 4).Font.Color = wColor
wsExcel.Cells(currentRow, 5) = dateIBM10(xZTREOPE0.TREOPEECH, True)
wsExcel.Cells(currentRow, 5).Font.Color = wColor
If xZTREOPE0.TREOPEDEV <> "EUR" Then
    For K = 1 To arrEUR_nb
        If xZTREOPE0.TREOPEDEV = arrEUR(K) Then
            If arrEURBID(K) <> arrEURFIX(K) Then
                wColor = vbMagenta
            End If
            X = Format$(arrEURBID(K), "###.00000")
            wsExcel.Cells(currentRow, 10) = X
            wsExcel.Cells(currentRow, 10).Font.Color = wColor
            wColor = vbBlue
            Exit For
        End If
    Next K
End If
End Sub

Public Sub fgTotal_PrintLine_xlsManual(lIndex As Integer, ByRef currentRow As Long, ByRef wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim wForecolor As Long, X As String, curX As Currency
Dim curX1 As Currency, curX2 As Currency
'On Error Resume Next
Dim I As Integer, wAUTSYCAUT As String
Dim blnDepassement As Boolean, wDepassement As String
Dim K As Integer
Dim wColor As Long
Dim ii As Long

blnDepassement = False
wColor = vbBlack
If blnId_Print Then
    Call prtSAB_TC_Lmites_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    Range("A5:P5").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste '2 lignes
    Call prtSAB_TC_Lmites_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    Range("A5:P5").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste
    Select Case xZCLIENA0.CLIENAAGE
        Case 0: prtFillColor = RGB(224, 255, 255)
        Case 1: prtFillColor = RGB(224, 255, 224)
        Case 2: prtFillColor = RGB(192, 255, 192)
    End Select
    For ii = 1 To 13
        wsExcel.Cells(currentRow, ii).Interior.Color = prtFillColor
    Next ii
    prtFillColor = RGB(255, 255, 128)
    wsExcel.Cells(currentRow, 14).Interior.Color = prtFillColor
    prtFillColor = RGB(255, 225, 114)
    wsExcel.Cells(currentRow, 15).Interior.Color = prtFillColor
    prtFillColor = RGB(255, 200, 100)
    wsExcel.Cells(currentRow, 16).Interior.Color = prtFillColor
    If xZAUTSYC0.AUTSYCFIN < wAmjM1_7C Then
        prtFillColor = RGB(255, 200, 100)
        wsExcel.Cells(currentRow, 5).Interior.Color = prtFillColor
    End If
     If xZAUTSYC0.AUTSYCDEV <> "EUR" Then
        prtFillColor = RGB(255, 200, 100)
        wsExcel.Cells(currentRow, 8).Interior.Color = prtFillColor
    End If
    wColor = &H800000
    wsExcel.Cells(currentRow, 1) = Mid$(xZCLIENA0.CLIENARA2, 1, 12)
    wsExcel.Cells(currentRow, 1).Font.Color = wColor
    wsExcel.Cells(currentRow, 2) = "'" & xZCLIENA0.CLIENACLI
    wsExcel.Cells(currentRow, 2).Font.Color = wColor
    wsExcel.Cells(currentRow, 3) = Mid$(xZCLIENA0.CLIENARA1, 1, 28)
    wsExcel.Cells(currentRow, 3).Font.Color = wColor
    If Mid$(xZAUTSYC0.AUTSYCAUT, 1, 1) = "P" Then
        For K = 1 To arrParam_PCT_Nb
            If xZTREOPE0.TREOPECLI = arrParam_PCT_CLI(K) Then
                prtFillColor = RGB(255, 160, 225)
                wColor = &HFFFFFF
                wsExcel.Cells(currentRow, 6) = arrParam_PCT_AUT(K)
                wsExcel.Cells(currentRow, 6).Font.Color = wColor
                wsExcel.Cells(currentRow, 6).Interior.Color = prtFillColor
                wColor = &H800000
                Exit For
            End If
        Next K
    End If
Else
    Call prtSAB_TC_Lmites_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    Range("A7:P7").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste
End If
If xZAUTSYC0.AUTSYCFIN < wAmj_Selection_7C Then
    wColor = vbMagenta
End If
If xZAUTSYC0.AUTSYCFIN = 0 Then
    If blnId_Print Then
        wsExcel.Cells(currentRow, 7) = "NEANT"
        wsExcel.Cells(currentRow, 7).Font.Color = wColor
        wsExcel.Cells(currentRow, 7).Font.Bold = True
    End If
Else
    wsExcel.Cells(currentRow, 5) = dateIBM10(xZAUTSYC0.AUTSYCFIN, True)
    wsExcel.Cells(currentRow, 5).Font.Color = wColor
End If
wColor = &H800000
meCV_xAUTSYC0
arrJ_AUTSYCMON = meCV2.Montant
'_______________________________________________________
fgTotal_Limites lIndex
'________________________________________
If Ope_Limite_Aut = 0 Then
    If blnId_Print Then
        wColor = vbMagenta
        wsExcel.Cells(currentRow, 7) = "NEANT"
        wsExcel.Cells(currentRow, 7).Font.Color = wColor
        wsExcel.Cells(currentRow, 7).Font.Bold = True
    End If
Else
    X = Format$(xZAUTSYC0.AUTSYCMON, "### ### ### ###.00")
    wsExcel.Cells(currentRow, 7) = X
    wsExcel.Cells(currentRow, 8) = xZAUTSYC0.AUTSYCDEV
    wColor = RGB(128, 64, 0)
    X = Format$(Ope_Limite_Max, "### ### ### ###.00")
    wsExcel.Cells(currentRow, 12) = X
    wsExcel.Cells(currentRow, 12).Font.Color = wColor
End If
If xZAUTSYC0.AUTSYCNIV = 1 Then
    wColor = &H800000
    wsExcel.Cells(currentRow, 9) = Trim(xZAUTSYC0.AUTSYCAUT)
    wsExcel.Cells(currentRow, 9).Font.Color = wColor
Else
    wColor = &H800000
    wsExcel.Cells(currentRow, 9) = "---(" & Trim(xZAUTSYC0.AUTSYCAUT) & ")"
    wsExcel.Cells(currentRow, 9).Font.Color = wColor
End If
curX = Abs(Ope_ENG(lIndex))
If curX <> 0 Then
    X = Format$(curX, "### ### ### ###.00")
    wsExcel.Cells(currentRow, 13) = X
    If blnId_Print And curX > Ope_Limite_Max Then
        blnDepassement = True
        wDepassement = Format$(curX - Ope_Limite_Max, "### ### ### ###.00")
    End If
End If
wColor = &H800000
curX = Abs(Ope_MAD(lIndex))
If curX > Ope_Limite_Aut Then
    If Ope_Limite_Aut = 0 And Not blnId_Print Then
        wsExcel.Cells(currentRow, 11) = " "
    Else
        prtFillColor = RGB(255, 200, 225)
        wColor = vbRed
        X = Format$(curX - Ope_Limite_Aut, "### ### ### ###.00")
        wsExcel.Cells(currentRow, 9) = X
        wsExcel.Cells(currentRow, 9).Font.Color = wColor
        wsExcel.Cells(currentRow, 9).Interior.Color = prtFillColor
    End If
End If
'======================================================================
If xZAUTSYC0.AUTSYCMON <> 0 Then
        wColor = RGB(128, 64, 0)
        wsExcel.Cells(currentRow, 11) = Ope_Limite_Max_Code
        wsExcel.Cells(currentRow, 11).Font.Color = wColor
End If
'======================================================================
If xZAUTSYC0.AUTSYCMON <> 0 Then
    If Ope_PJJ <= 0 Then
        wColor = vbRed
    Else
        wColor = &H800000
        X = Format$(Ope_PJJ, "### ### ### ###.00")
        wsExcel.Cells(currentRow, 14) = X
    End If
    wsExcel.Cells(currentRow, 14).Font.Color = wColor
End If
If xZAUTSYC0.AUTSYCMON <> 0 Then
    If Ope_PTM <= 0 Then
        wColor = vbRed
    Else
        wColor = &H800000
        X = Format$(Ope_PTM, "### ### ### ###.00")
        wsExcel.Cells(currentRow, 15) = X
    End If
    wsExcel.Cells(currentRow, 15).Font.Color = wColor
End If
If xZAUTSYC0.AUTSYCMON <> 0 Then
    If Ope_PCT <= 0 Then
        wColor = vbRed
    Else
        wColor = &H800000
        X = Format$(Ope_PCT, "### ### ### ###.00")
        wsExcel.Cells(currentRow, 16) = X
    End If
End If

If blnDepassement Then
    Call prtSAB_TC_Lmites_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    Range("A5:P5").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste
    prtFillColor = RGB(255, 200, 225)
    wColor = vbRed
    X = Format$(wDepassement, "### ### ### ###.00")
    wsExcel.Cells(currentRow, 12) = X
    wsExcel.Cells(currentRow, 12).Font.Color = wColor
    wsExcel.Cells(currentRow, 12).Interior.Color = prtFillColor
End If
'_____________________________________________________________________________________
If blnId_Print Then
    For arrJ_K = 0 To arrJ_K_Max
        If xZAUTSYC0.AUTSYCFIN < arrJ_AMJ_7c(arrJ_K) And arrJ_MTE(arrJ_K) <> 0 Then
            Call prtSAB_TC_Lmites_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
            Range("A5:P5").Select
            Selection.Copy
            Range("A" & CStr(currentRow)).Select
            ActiveSheet.Paste
            prtFillColor = mColor_Y1
            wColor = vbRed
            wsExcel.Cells(currentRow, 5) = Trim(xZAUTSYC0.AUTSYCAUT) & ": en-cours au-delà de la date limite d'autorisation  : EUR "
            wsExcel.Cells(currentRow, 5).Font.Color = wColor
            wsExcel.Cells(currentRow, 5).Font.Bold = True
            wsExcel.Cells(currentRow, 5).Interior.Color = prtFillColor
            X = Format$(arrJ_MTE(arrJ_K), "### ### ### ###.00")
            wsExcel.Cells(currentRow, 12) = X
            wsExcel.Cells(currentRow, 12).Font.Color = wColor
            wsExcel.Cells(currentRow, 12).Font.Bold = True
            wsExcel.Cells(currentRow, 12).Interior.Color = prtFillColor
            Exit For
        Else
            If arrJ_MTE(arrJ_K) > arrJ_AUTSYCMON Then
                Call prtSAB_TC_Lmites_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
                Range("A5:P5").Select
                Selection.Copy
                Range("A" & CStr(currentRow)).Select
                ActiveSheet.Paste
                prtFillColor = mColor_Y1
                X = Format$(arrJ_MTE(arrJ_K), "### ### ### ###.00")
                wsExcel.Cells(currentRow, 5) = Trim(xZAUTSYC0.AUTSYCAUT) & ": à partir du " & dateImp(arrJ_AMJ_7c(arrJ_K) + 19000000) _
                 & ", l'en-cours " & X & " est supérieur au montant autorisé : EUR " & Format$(arrJ_AUTSYCMON, "### ### ### ##0")
                wsExcel.Cells(currentRow, 5).Font.Color = wColor
                wsExcel.Cells(currentRow, 5).Font.Bold = True
                wsExcel.Cells(currentRow, 5).Interior.Color = prtFillColor
                Exit For
            End If
        End If
    Next arrJ_K
End If
'_____________________________________________________________________________________
blnId_Print = False
End Sub

Private Sub zoneImpression_xlsManual(lFct As String, nbRows As Long, wsheet As Excel.Worksheet)

    Call init_TypePagesetup
    If nbRows > 0 Then
        wsheet.Activate
        wsheet.Range("A1:P" & CStr(nbRows)).Select
        zoneImpressionPagesetup.PrintArea = "$A$1:$P$" & CStr(nbRows)
        zoneImpressionPagesetup.LeftFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "prtSAB_TC_Limites   &D &T  BIA_INFO"
        zoneImpressionPagesetup.RightFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "&P"
        zoneImpressionPagesetup.Orientation = xlLandscape
        zoneImpressionPagesetup.Zoom = 75
    End If
    Call SetTypePageSetup(wsheet)

End Sub
Public Sub DevF_Load()
Dim X As String, xSQL As String
ReDim arrDevF_ISO(1000), arrDevF_AMJx(1000), arrDevF_AMJ(1000)

arrDevF_Max = 1000
arrDevF_Nb = 0
xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 where BASTABNUM = 36 order by BASTABARG"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
        arrDevF_Nb = arrDevF_Nb + 1
        If arrDevF_Nb > arrDevF_Max Then
            arrDevF_Max = arrDevF_Max + 100
            ReDim Preserve arrDevF_ISO(arrDevF_Max), arrDevF_AMJ(arrDevF_Max), arrDevF_AMJx(arrDevF_Max)
        End If
    X = rsSab("BASTABARG")
    arrDevF_ISO(arrDevF_Nb) = Mid$(X, 3, 3)
    arrDevF_AMJ(arrDevF_Nb) = 19000000 + convX2P(Mid$(X, 6, 4))
    arrDevF_AMJx(arrDevF_Nb) = dateImp10_S(arrDevF_AMJ(arrDevF_Nb))
    rsSab.MoveNext
Loop
End Sub

Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = 0 To fgSelect_arrIndex - 3
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 0 To fgSelect_arrIndex - 3
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.Col = 0
    End If
End If

End Sub
Public Sub fgTotal_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgTotal.Row

If lRow > 0 And lRow < fgTotal.Rows Then
    fgTotal.Row = lRow
    For I = 0 To fgTotal_arrIndex
        fgTotal.Col = I: fgTotal.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgTotal.Row = mRow
    If fgTotal.Row > 0 Then
        lRow = fgTotal.Row
        lColor_Old = fgTotal.CellBackColor
        For I = 0 To fgTotal_arrIndex
          fgTotal.Col = I: fgTotal.CellBackColor = lColor
        Next I
        fgTotal.Col = 0
    End If
End If

End Sub

Private Sub fgSelect_Display()
Dim V, I As Integer
Dim X As String, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
cmdPrint.Enabled = False

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgSelect_Display"
Set rsSab = Nothing

For I = 1 To arrClient_Nb
    xZTREOPE0 = arrZTREOPE0(I)
    fgSelect_Display_ZCLIENA0 xZTREOPE0.TREOPECLI, I
    blnOk = False
    blnDisplay = True
    fgSelect_DisplayLine (I)
    
Next I

If chkSelect_EnCours = "1" Then fgSelect_Display_Suite ' client avec autorisation sans en cours

fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Comptes : " & arrClient_Nb): DoEvents
If fgSelect.Rows > 1 Then
    cmdPrint.Enabled = True
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub fgEUR_Display()
Dim V, I As Integer
Dim X As String, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
fgEUR.Visible = False

fgEUR.Rows = 1
currentAction = "fgEUR_Display"

For I = 1 To arrEUR_nb
    fgEUR.Rows = fgEUR.Rows + 1
    fgEUR.Row = fgEUR.Rows - 1
    fgEUR.Col = 0: fgEUR.Text = arrEUR(I)
    fgEUR.Col = 1: fgEUR.Text = Format(arrEURFIX(I), "###.######")
    fgEUR.CellForeColor = vbBlue
    fgEUR.Col = 2: fgEUR.Text = Format(arrEURBID(I), "###.######")
    If arrEURFIX(I) = arrEURBID(I) Then
        fgEUR.CellForeColor = vbBlue
    Else
        fgEUR.CellForeColor = vbMagenta
    End If
    
    
Next I

fgEUR.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub fgTotal_Display()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgTotal.Visible = False
fgTotal_Reset
cmdPrint.Enabled = False

fgTotal.Rows = 1
'X = fgTotal_FormatString
'X = Replace(X, "    0jj/mm/aaaa", "dispo :" & dateImp10(wAmj_Selection_8C))
'X = Replace(X, "    1jj/mm/aaaa", "        Tomorrow   .")
'X = Replace(X, "    2jj/mm/aaaa", "spot : " & dateImp10(wAmjMax_8C))

fgTotal.FormatString = fgTotal_FormatString 'X
currentAction = "fgTotal_Display"

For I = 0 To selZAUTSYC0_Nb
    If selZAUTSYC0_Display(I) Then
        xZAUTSYC0 = selZAUTSYC0(I)
        fgTotal_DisplayLine (I)
    End If
Next I

'fgTotal.Height = fgTotal.CellHeight * fgTotal.Rows + fgTotal.CellHeight
If fgTotal.Rows > 1 Then fgTotal.TopRow = 1
fgTotal.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & selZAUTSYC0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub selZAUTSYC0_SQL(lTREOPECLI As String)
Dim V
Dim X As String, xSQL As String, blnAUT_Ok As Boolean, K As Integer

On Error GoTo Error_Handler
ReDim selZAUTSYC0(101), selZAUTSYC0_PIB(101)
selZAUTSYC0_Max = 100: selZAUTSYC0_Nb = 0

mAUTSYCADR_PIB = -1

Set rsSab = Nothing
X = "AUTSYCCLI =  '" & lTREOPECLI & "'"
'$20070731 $JPL xSql = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0 where " & X & " order by AUTSYCCLI , AUTSYCADR"
xSQL = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0 where " & X & " order by AUTSYCCLI , AUTSYCPER,AUTSYCADR"
Set rsSab = cnsab.Execute(xSQL)

    
Do While Not rsSab.EOF
    V = rsZAUTSYC0_GetBuffer(rsSab, xZAUTSYC0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSAB_Tc_Limites.fgTotal_Display"
        '' Exit Sub
     Else
     
        blnAUT_Ok = True
        'For K = 1 To arrParam_AUT_Nb
        '    If Mid$(xZAUTSYC0.AUTSYCAUT, 1, 3) = arrParam_AUT(K) Then
        '        blnAUT_Ok = False
        '        Exit For
        '    End If
        'Next K
        If blnAUT_Ok Then
     
     
             selZAUTSYC0_Nb = selZAUTSYC0_Nb + 1
             If selZAUTSYC0_Nb > selZAUTSYC0_Max Then
                 selZAUTSYC0_Max = selZAUTSYC0_Max + 50
                 ReDim Preserve selZAUTSYC0(selZAUTSYC0_Max + 50)
                 ReDim Preserve selZAUTSYC0_PIB(selZAUTSYC0_Max + 50)
             End If
            If xZAUTSYC0.AUTSYCFIN < wAmj_Selection_7C Then xZAUTSYC0.AUTSYCMON = 0
            If Trim(xZAUTSYC0.AUTSYCAUT) = "TPL" Then xZAUTSYC0.AUTSYCAUT = "PRE-TPL"
             selZAUTSYC0(selZAUTSYC0_Nb) = xZAUTSYC0
             
            ' Chargement AUTORISATION ANTERIEURE PIB & EMP
            '=============================================
            If chkSelect_AmjMin = "1" Then  ' Date antérieure cochée
                AUTORISATIONS_ANTERIEURES
                selZAUTSYC0(selZAUTSYC0_Nb) = xZAUTSYC0
                
            End If
            
            If Mid$(xZAUTSYC0.AUTSYCAUT, 1, 3) = "PIB" Then
                mAUTSYCADR_PIB = xZAUTSYC0.AUTSYCADR
                selZAUTSYC0_PIB(selZAUTSYC0_Nb) = True
            Else
                If xZAUTSYC0.AUTSYCPER = mAUTSYCADR_PIB Then
                    selZAUTSYC0_PIB(selZAUTSYC0_Nb) = True
                End If
            End If
        End If
    
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


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

Private Sub fgDossier_Display()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fraDétail.Visible = False
fgDossier_Reset
cmdPrint.Enabled = False

fgDossier.Rows = 1
fgDossier.FormatString = fgDossier_FormatString
currentAction = "fgdossier_Display"
    
For I = 1 To selZTREOPE0_Nb
         
    xZTREOPE0 = selZTREOPE0(I)
    fgDossier_DisplayLine (selZTREOPE0_Nb)
Next I
fgDossier.Height = fgDossier.CellHeight * fgDossier.Rows + fgDossier.CellHeight

fraDétail.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & selZTREOPE0_Nb): DoEvents
If fgDossier.Rows > 1 Then
 '   fgDossier_Sort1 = 0: fgDossier_Sort2 = 1: fgDossier_Sort
    fgDossier.TopRow = 1
    cmdPrint.Enabled = True
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub selZTREOPE0_SQL(lTREOPECLI As String)
Dim V
Dim X As String, xSQL As String
Dim blnEnCours As Boolean
Dim xRange As String
On Error GoTo Error_Handler
ReDim selZTREOPE0(101)
selZTREOPE0_Max = 100: selZTREOPE0_Nb = 0

Set rsSab = Nothing

xRange = "TREOPECLI =  '" & lTREOPECLI & "'"
If lTREOPECLI < "0010000" Then
'_________________________________________________________________________________
    xRange = "TREOPECLI in ("
    X = " where CLIGRPREL = 'AUT' and CLIGRPREG ='" & lTREOPECLI & "'"
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIGRP0" & X & " order by  CLIGRPCLI"
    Set rsSab = cnsab.Execute(xSQL)
    Do While Not rsSab.EOF
        xRange = xRange & "'" & rsSab("CLIGRPCLI") & "',"
        rsSab.MoveNext
    Loop
    Mid$(xRange, Len(xRange), 1) = ")"
End If
'_________________________________________________________________________________


X = xWhere_Selection & " and " & xRange _
    & " AND TREOPEOPR =  '" & xZTREOPE0.TREOPEOPR & "'"
xSQL = "select * from " & paramIBM_Library_SAB & ".ZTREOPE0  " & X & " order by TREOPEAUT , TREOPEECH"
Set rsSab = cnsab.Execute(xSQL)

    
Do While Not rsSab.EOF
    V = rsZTREOPE0_GetBuffer(rsSab, xZTREOPE0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSAB_Tc_Limites.fgdossier_Display"
        '' Exit Sub
     Else
        blnEnCours = True
' ? échéance anticipée

        If xZTREOPE0.TREOPEETA = "6" _
        And xZTREOPE0.TREOPEREE > 0 _
        And xZTREOPE0.TREOPEREE < wAmj_Selection_7C Then blnEnCours = False
        
        If blnEnCours Then

            selZTREOPE0_Nb = selZTREOPE0_Nb + 1
            If selZTREOPE0_Nb > selZTREOPE0_Max Then
                selZTREOPE0_Max = selZTREOPE0_Max + 50
                ReDim Preserve arrZTREOPE0(selZTREOPE0_Max + 50)
            End If
            
            selZTREOPE0(selZTREOPE0_Nb) = xZTREOPE0
        End If
        
    End If
    rsSab.MoveNext
Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub fgDossier_DisplayLine(lIndex As Long)
Dim mCellforeColor As Long
On Error Resume Next
mCellforeColor = vbBlue
If xZTREOPE0.TREOPEECH < wAmj_Selection_7C Then mCellforeColor = vbRed
If xZTREOPE0.TREOPEECH = wAmj_Selection_7C Then mCellforeColor = vbMagenta
If xZTREOPE0.TREOPEDIS > wAmj_Selection_7C Then mCellforeColor = RGB(0, 100, 0)

fgDossier.Rows = fgDossier.Rows + 1
fgDossier.Row = fgDossier.Rows - 1
fgDossier.Col = 0: fgDossier.Text = xZTREOPE0.TREOPEAUT
fgDossier.CellForeColor = mCellforeColor
fgDossier.Col = 1: fgDossier.Text = xZTREOPE0.TREOPEOPR & "    " & xZTREOPE0.TREOPENAT & "    " & xZTREOPE0.TREOPENUM
fgDossier.CellForeColor = mCellforeColor
fgDossier.Col = 5: fgDossier.Text = Format$(xZTREOPE0.TREOPEMNT, "### ### ### ###.00")
fgDossier.CellForeColor = mCellforeColor
fgDossier.Col = 6: fgDossier.Text = xZTREOPE0.TREOPEDEV
fgDossier.CellForeColor = mCellforeColor
meCV_xTREOPE0
fgDossier.Col = 7: fgDossier.Text = Format$(meCV2.Montant, "### ### ### ###.00")
fgDossier.CellForeColor = mCellforeColor
fgDossier.Col = 2: fgDossier.Text = dateIBM10(xZTREOPE0.TREOPENEG, True)
fgDossier.CellForeColor = mCellforeColor
fgDossier.Col = 3: fgDossier.Text = dateIBM10(xZTREOPE0.TREOPEDIS, True)
fgDossier.CellForeColor = mCellforeColor
fgDossier.Col = 4: fgDossier.Text = dateIBM10(xZTREOPE0.TREOPEECH, True)
fgDossier.CellForeColor = mCellforeColor
fgDossier.Col = 8: fgDossier.Text = xZTREOPE0.TREOPEETA
fgDossier.CellForeColor = mCellforeColor
fgDossier.Col = fgDossier_arrIndex: fgDossier.Text = lIndex
End Sub


Public Sub fgDossier_PrintLine(lIndex As Long)
Dim K As Integer, X As String
On Error Resume Next
prtSAB_TC_Lmites_NewLine
XPrt.FontSize = 7

XPrt.ForeColor = vbBlue
If xZTREOPE0.TREOPEECH < wAmj_Selection_7C Then XPrt.ForeColor = vbRed
If xZTREOPE0.TREOPEECH = wAmj_Selection_7C Then XPrt.ForeColor = vbMagenta
If xZTREOPE0.TREOPEDIS > wAmj_Selection_7C Then XPrt.ForeColor = RGB(0, 100, 0)

If xZCLIENA0.CLIENACLI < "0010000" Then
    XPrt.CurrentX = prtMinX + 1300
    XPrt.Print xZTREOPE0.TREOPECLI;
End If
XPrt.CurrentX = prtMinX + 2000: XPrt.Print xZTREOPE0.TREOPEOPR;
XPrt.CurrentX = prtMinX + 2500: XPrt.Print xZTREOPE0.TREOPENAT;
XPrt.CurrentX = prtMinX + 3000: XPrt.Print xZTREOPE0.TREOPENUM;
X = Format$(xZTREOPE0.TREOPEMNT, "### ### ### ###.00")
XPrt.CurrentX = prtMinX + 7000 - XPrt.TextWidth(X)
XPrt.Print X & " " & xZTREOPE0.TREOPEDEV;
XPrt.CurrentX = prtMinX + 7500: XPrt.Print xZTREOPE0.TREOPEAUT;

'XPrt.CurrentX = prtMinX + 4000: XPrt.Print dateIBM10(xZTREOPE0.TREOPENEG, True);
XPrt.CurrentX = prtMinX + 4000: XPrt.Print dateIBM10(xZTREOPE0.TREOPEDIS, True);
XPrt.CurrentX = prtMinX + 4900: XPrt.Print dateIBM10(xZTREOPE0.TREOPEECH, True);


If xZTREOPE0.TREOPEDEV <> "EUR" Then
    XPrt.FontItalic = True
    'XPrt.ForeColor = &H800000
    
    For K = 1 To arrEUR_nb
        If xZTREOPE0.TREOPEDEV = arrEUR(K) Then
            If arrEURBID(K) <> arrEURFIX(K) Then XPrt.ForeColor = vbMagenta
            X = Format$(arrEURBID(K), "###.00000")
            XPrt.CurrentX = prtMinX + 8500 - XPrt.TextWidth(X)
            XPrt.Print X;
            Exit For
        End If
    Next K
    XPrt.FontItalic = False
End If
End Sub
Public Sub fgDossier_Reset()
fgDossier.Clear
fgDossier_Sort1 = 0: fgDossier_Sort2 = 0
fgDossier_Sort1_Old = -1
fgDossier_RowDisplay = 0: fgDossier_RowClick = 0
fgDossier_arrIndex = fgDossier.Cols - 1
blnfgDossier_DisplayLine = False
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
'cboDevise_Reset
End Sub

Public Sub fgSelect_DisplayLine(lIndex As Integer)
Dim wForecolor As Long
Dim curAut As Currency
Dim K As Integer
On Error Resume Next

Dim curX As Currency
Dim I As Integer, wAUTSYCAUT As String

meCV2.Montant = 0
wForecolor = vbRed

If xZTREOPE0.TREOPEOPR = "PRE" Then
    wAUTSYCAUT = "P" '"PIB"
Else
    If xZTREOPE0.TREOPEOPR = "EMP" Then
        wAUTSYCAUT = "E" '"EMB"
    End If
End If

For I = 1 To arrZAUTSYC0_Nb
    If arrZAUTSYC0(I).AUTSYCCLI = xZTREOPE0.TREOPECLI _
    And Mid$(arrZAUTSYC0(I).AUTSYCAUT, 1, 1) = wAUTSYCAUT Then
    'And Trim(arrZAUTSYC0(I).AUTSYCAUT) = wAUTSYCAUT Then
        blnZAUTSYC0_EnCours(I) = True
        wForecolor = vbBlue
        xZAUTSYC0 = arrZAUTSYC0(I)
        meCV_xAUTSYC0
        curAut = meCV2.Montant
        If xZAUTSYC0.AUTSYCFIN > 0 Then
            
            If xZAUTSYC0.AUTSYCFIN < wAmj_Selection_7C Then meCV2.Montant = 0
            If xZTREOPE0.TREOPEMNT > meCV2.Montant Then
                wForecolor = vbRed
            End If
        End If
        Exit For
    End If
    
Next I
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
If xZTREOPE0.TREOPECOU < "0010000" Then fgSelect.Col = 1: fgSelect.Text = xZTREOPE0.TREOPECOU
fgSelect.Col = 2: fgSelect.Text = xZTREOPE0.TREOPECLI
fgSelect.Col = 3: fgSelect.Text = xZTREOPE0.TREOPEOPR & " " & xZAUTSYC0.AUTSYCTAU & "%"

fgSelect.Col = 0: fgSelect.Text = Trim(xZCLIENA0.CLIENARA2)
fgSelect.CellForeColor = wForecolor

fgSelect.Col = 4: fgSelect.Text = Format$(xZTREOPE0.TREOPEMNT, "### ### ### ###.00")
fgSelect.CellForeColor = wForecolor
fgSelect.Col = 5: fgSelect.Text = Format$(curAut, "### ### ### ###.00")
If xZAUTSYC0.AUTSYCFIN < wAmj_Selection_7C Then fgSelect.CellForeColor = vbRed

For K = 1 To arrParam_PCT_Nb
    If xZTREOPE0.TREOPECLI = arrParam_PCT_CLI(K) Then
        fgSelect.Col = 6: fgSelect.Text = arrParam_PCT_AUT(K)
        fgSelect.CellBackColor = mColor_Y1 'RGB(255, 190, 255) 'vbMagenta
        Exit For

    End If
Next K

If Trim(xZTREOPE0.TREOPECMT) <> "" Then
End If

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex


End Sub


Public Sub fgTotal_DisplayLine(lIndex As Integer)
Dim wForecolor As Long
Dim X As String, curX As Currency
Dim curX1 As Currency, curX2 As Currency

On Error Resume Next

Dim I As Integer, wAUTSYCAUT As String

fgTotal.Rows = fgTotal.Rows + 1
fgTotal.Row = fgTotal.Rows - 1
fgTotal.Col = 0: fgTotal.Text = xZAUTSYC0.AUTSYCAUT

meCV_xAUTSYC0

fgTotal.Col = 2: fgTotal.Text = Format$(Ope_Limite_Aut, "### ### ### ###.00")
If xZAUTSYC0.AUTSYCFIN = 0 Then
    fgTotal.Col = 1: fgTotal.Text = ""
Else

    If xZAUTSYC0.AUTSYCFIN < wAmj_Selection_7C Then    ' autorisation échue
        meCV2.Montant = 0
         wForecolor = vbRed
     Else
         wForecolor = vbBlue
     End If
     fgTotal.CellForeColor = wForecolor
     fgTotal.Col = 1: fgTotal.Text = dateIBM10(xZAUTSYC0.AUTSYCFIN, True)
     fgTotal.CellForeColor = wForecolor
End If

fgTotal.Col = 3
If Ope_ENG(lIndex) <> 0 Then
    fgTotal.Text = Format$(Ope_ENG(lIndex), "### ### ### ###.00")
Else
    fgTotal.Text = ""
End If

If Ope_Limite_Aut > 0 Then
    fgTotal.Col = 4

    If Ope_MAD(lIndex) > Ope_Limite_Aut Then
        fgTotal.Text = Format$(Ope_MAD(lIndex) - Ope_Limite_Aut, "### ### ### ###.00")
         fgTotal.CellForeColor = vbRed
    Else
         fgTotal.Text = ""
         fgTotal.CellForeColor = vbBlue
     End If
    
'___________________________________________________________________
    
fgTotal_Limites lIndex
'___________________________________________________________________

    fgTotal.Col = 5
    fgTotal.CellBackColor = RGB(210, 255, 210)
    fgTotal.Text = Format$(Ope_PJJ, "### ### ### ###.00")
    If Ope_Limite < 0 Then
         fgTotal.CellForeColor = vbRed
    Else
         fgTotal.CellForeColor = vbBlue
    End If

    
    fgTotal.Col = 6
    fgTotal.CellBackColor = RGB(255, 255, 128)
    fgTotal.Text = Format$(Ope_PTM, "### ### ### ###.00")
    If Ope_Limite < 0 Then
         fgTotal.CellForeColor = vbRed
    Else
         fgTotal.CellForeColor = vbBlue
    End If
    
    fgTotal.Col = 7
    fgTotal.CellBackColor = RGB(255, 200, 100)
    fgTotal.Text = Format$(Ope_PCT, "### ### ### ###.00")
    If Ope_Limite < 0 Then
         fgTotal.CellForeColor = vbRed
    Else
         fgTotal.CellForeColor = vbBlue
    End If

End If

fgTotal.Col = fgTotal_arrIndex: fgTotal.Text = lIndex

End Sub


Public Sub fgTotal_PrintLine(lIndex As Integer)
Dim wForecolor As Long, X As String, curX As Currency
Dim curX1 As Currency, curX2 As Currency
On Error Resume Next
Dim I As Integer, wAUTSYCAUT As String
Dim blnDepassement As Boolean, wDepassement As String
Dim K As Integer

prtSAB_TC_Lmites_NewLine
'XPrt.FontBold = True
XPrt.FontSize = 7
blnDepassement = False

If blnId_Print Then
    prtSAB_TC_Lmites_NewLine
    XPrt.FontBold = True

    If XPrt.CurrentY + 600 > prtMaxY Then
        frmElpPrt.prtNewPage
        prtSAB_TC_Lmites_Form
    End If
    'If xZCLIENA0.CLIENACLI < "0010000" Then
    '    prtFillColor = RGB(128, 255, 255)
    'Else
    '    prtFillColor = RGB(220, 255, 255)
    'End If
    Select Case xZCLIENA0.CLIENAAGE
        Case 0: prtFillColor = RGB(224, 255, 255)
        Case 1: prtFillColor = RGB(224, 255, 224)
        Case 2: prtFillColor = RGB(192, 255, 192)
    End Select
    XPrt.FontSize = 7

    Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMinX + 11700, XPrt.CurrentY + prtHeaderHeight, " ")
    
      '  prtFillColor = RGB(210, 255, 210)
      '  Call frmElpPrt.prtTrame_Color(prtMinX + 11750, XPrt.CurrentY, prtMinX + 13150, XPrt.CurrentY + prtHeaderHeight, " ")
    
    prtFillColor = RGB(255, 255, 128)
    Call frmElpPrt.prtTrame_Color(prtMinX + 11800, XPrt.CurrentY, prtMinX + 13150, XPrt.CurrentY + prtHeaderHeight, " ")
    
    prtFillColor = RGB(255, 225, 114)
    Call frmElpPrt.prtTrame_Color(prtMinX + 13200, XPrt.CurrentY, prtMinX + 14500, XPrt.CurrentY + prtHeaderHeight, " ")
   
    prtFillColor = RGB(255, 200, 100)
    Call frmElpPrt.prtTrame_Color(prtMinX + 14550, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, " ")
 
    If xZAUTSYC0.AUTSYCFIN < wAmjM1_7C Then
        prtFillColor = RGB(255, 200, 100)
        Call frmElpPrt.prtTrame_Color(prtMinX + 4800, XPrt.CurrentY, prtMinX + 5650, XPrt.CurrentY + prtHeaderHeight, " ")
    End If
    
    
     If xZAUTSYC0.AUTSYCDEV <> "EUR" Then
        prtFillColor = RGB(255, 200, 100)
        Call frmElpPrt.prtTrame_Color(prtMinX + 7000, XPrt.CurrentY, prtMinX + 7350, XPrt.CurrentY + prtHeaderHeight, " ")
    End If
   
    
   '&H00C0E0FF&
    XPrt.CurrentY = XPrt.CurrentY + 80
    'Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 20, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ", 250)
  ''  XPrt.Line (prtMinX, XPrt.CurrentY - 20)-(prtMaxX, XPrt.CurrentY - 20), prtLineColor
    XPrt.ForeColor = &H800000    'vbBlack
    XPrt.CurrentX = prtMinX
    XPrt.Print Mid$(xZCLIENA0.CLIENARA2, 1, 12); 'Mid$(xZCLIENA0.CLIENASIG, 1, 8);
    XPrt.CurrentX = prtMinX + 1300
    XPrt.Print xZCLIENA0.CLIENACLI;
    XPrt.CurrentX = prtMinX + 2000
    XPrt.Print Mid$(xZCLIENA0.CLIENARA1, 1, 28);
    ''If xZAUTSYC0.AUTSYCFIN < wAmj_Selection_7C Then XPrt.CurrentX = prtMinX + 5800: XPrt.Print "????";

    If Mid$(xZAUTSYC0.AUTSYCAUT, 1, 1) = "P" Then
        For K = 1 To arrParam_PCT_Nb
            If xZTREOPE0.TREOPECLI = arrParam_PCT_CLI(K) Then
                prtFillColor = RGB(255, 160, 225) 'RGB(255, 0, 225)
                XPrt.ForeColor = &HFFFFFF
                Call frmElpPrt.prtTrame_Color(prtMinX + 5650, XPrt.CurrentY - 80, prtMinX + 6000, XPrt.CurrentY + prtHeaderHeight - 80, " ")
                XPrt.CurrentX = prtMinX + 5700
                XPrt.Print arrParam_PCT_AUT(K);
                XPrt.ForeColor = &H800000    'vbBlack
    
                Exit For
            End If
        Next K
    End If

Else
    'XPrt.FontItalic = True
    XPrt.ForeColor = vbBlue
    XPrt.FontSize = 6
End If


'If xZAUTSYC0.AUTSYCFIN > 0 Then

    If xZAUTSYC0.AUTSYCFIN < wAmj_Selection_7C Then XPrt.ForeColor = vbMagenta: XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 4900
    If xZAUTSYC0.AUTSYCFIN = 0 Then
        If blnId_Print Then XPrt.Print "NEANT";
    Else
        XPrt.Print dateIBM10(xZAUTSYC0.AUTSYCFIN, True);
    End If
    XPrt.ForeColor = &H800000    'vbBlack
    XPrt.FontBold = False
'End If

meCV_xAUTSYC0
arrJ_AUTSYCMON = meCV2.Montant
'_______________________________________________________
fgTotal_Limites lIndex
'________________________________________



If Ope_Limite_Aut = 0 Then
    If blnId_Print Then
        XPrt.FontBold = True
        XPrt.ForeColor = vbMagenta
        XPrt.CurrentX = prtMinX + 6500: XPrt.Print "NEANT";
        XPrt.ForeColor = &H800000    'vbBlack
        XPrt.FontBold = False
    End If
Else
    X = Format$(xZAUTSYC0.AUTSYCMON, "### ### ### ###.00")
    XPrt.CurrentX = prtMinX + 7000 - XPrt.TextWidth(X)
    XPrt.Print X & " " & xZAUTSYC0.AUTSYCDEV;
'    X = Format$(Ope_Limite_Aut, "### ### ### ###.00")
    XPrt.ForeColor = RGB(128, 64, 0) 'RGB(200, 100, 0)
    X = Format$(Ope_Limite_Max, "### ### ### ###.00")
    XPrt.CurrentX = prtMinX + 10300 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.ForeColor = &H800000
End If
XPrt.CurrentX = prtMinX + 7500
If xZAUTSYC0.AUTSYCNIV = 1 Then
    XPrt.Print Trim(xZAUTSYC0.AUTSYCAUT);
Else
    XPrt.Print "---(" & Trim(xZAUTSYC0.AUTSYCAUT) & ")";
End If


curX = Abs(Ope_ENG(lIndex))
If curX <> 0 Then
    X = Format$(curX, "### ### ### ###.00")
    XPrt.CurrentX = prtMinX + 11600 - XPrt.TextWidth(X)
    XPrt.Print X;
    
    If blnId_Print And curX > Ope_Limite_Max Then
        blnDepassement = True
        wDepassement = Format$(curX - Ope_Limite_Max, "### ### ### ###.00")
    End If
        
End If

XPrt.ForeColor = &H800000 'vbBlack
curX = Abs(Ope_MAD(lIndex))
If curX > Ope_Limite_Aut Then
    If Ope_Limite_Aut = 0 And Not blnId_Print Then
        XPrt.CurrentX = prtMinX + 9000: XPrt.Print " "; ' ">";
    Else
        prtFillColor = RGB(255, 200, 225)
        Call frmElpPrt.prtTrame_Color(prtMinX + 7800, XPrt.CurrentY - 80, prtMinX + 8800, XPrt.CurrentY + prtHeaderHeight - 80, " ")
        XPrt.ForeColor = vbRed
        X = Format$(curX - Ope_Limite_Aut, "### ### ### ###.00")
        XPrt.CurrentX = prtMinX + 8750 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
End If
'======================================================================

If xZAUTSYC0.AUTSYCMON <> 0 Then
    '$JPL 20091013 If blnLimite_Max_Code Then

        XPrt.ForeColor = RGB(128, 64, 0) 'vbMagenta
        '$JPL 20090929
        'prtFillColor = RGB(210, 255, 210)
        'Call frmElpPrt.prtTrame_Color(prtMinX + 8800, XPrt.CurrentY - 80, prtMinX + 9200, XPrt.CurrentY + prtHeaderHeight - 80, " ")

        XPrt.CurrentX = prtMinX + 9150 - XPrt.TextWidth(Ope_Limite_Max_Code)
        
        XPrt.Print Ope_Limite_Max_Code;
    '$JPL 20091013 End If
    '$JPL 20090929 If Ope_Limite <= 0 Then
    '$JPL 20090929     XPrt.ForeColor = vbRed
    '$JPL 20090929 Else
    '$JPL 20090929     XPrt.ForeColor = &H800000 'vbBlack
    '$JPL 20090929 End If
    '$JPL 20090929 X = Format$(Ope_Limite, "### ### ### ###.00")
    '$JPL 20090929 XPrt.CurrentX = prtMinX + 12850 - XPrt.TextWidth(X)
    '$JPL 20090929 XPrt.Print X;
    
End If
'======================================================================

If xZAUTSYC0.AUTSYCMON <> 0 Then
    If Ope_PJJ <= 0 Then
        XPrt.ForeColor = vbRed
    Else
        XPrt.ForeColor = &H800000 'vbBlack
    'JPL 20091013 End If
        X = Format$(Ope_PJJ, "### ### ### ###.00")
        XPrt.CurrentX = prtMinX + 13000 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If  'JPL 20091013
    
End If


If xZAUTSYC0.AUTSYCMON <> 0 Then
    If Ope_PTM <= 0 Then
        XPrt.ForeColor = vbRed
    Else
        XPrt.ForeColor = &H800000 'vbBlack
    'JPL 20091013 End If
        X = Format$(Ope_PTM, "### ### ### ###.00")
        XPrt.CurrentX = prtMinX + 14300 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If  'JPL 20091013
    
End If

If xZAUTSYC0.AUTSYCMON <> 0 Then
    If Ope_PCT <= 0 Then
        XPrt.ForeColor = vbRed
    Else
        XPrt.ForeColor = &H800000 'vbBlack
    'JPL 20091013 End If
        X = Format$(Ope_PCT, "### ### ### ###.00")
        XPrt.CurrentX = prtMinX + 15700 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If  'JPL 20091013
    
End If

If blnDepassement Then
    prtSAB_TC_Lmites_NewLine
    prtFillColor = RGB(255, 200, 225)
    Call frmElpPrt.prtTrame_Color(prtMinX + 10400, XPrt.CurrentY - 80, prtMinX + 11700, XPrt.CurrentY + prtHeaderHeight - 80, " ")
    XPrt.ForeColor = vbRed
    X = Format$(wDepassement, "### ### ### ###.00")
    XPrt.CurrentX = prtMinX + 11600 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

'_____________________________________________________________________________________
If blnId_Print Then
    For arrJ_K = 0 To arrJ_K_Max
        If xZAUTSYC0.AUTSYCFIN < arrJ_AMJ_7c(arrJ_K) And arrJ_MTE(arrJ_K) <> 0 Then
            prtSAB_TC_Lmites_NewLine
            prtFillColor = mColor_Y1
            Call frmElpPrt.prtTrame_Color(prtMinX + 4900, XPrt.CurrentY - 80, prtMinX + 11700, XPrt.CurrentY + prtHeaderHeight - 80, " ")
            XPrt.ForeColor = vbRed
            XPrt.CurrentX = prtMinX + 4900
            XPrt.FontBold = True
            XPrt.Print Trim(xZAUTSYC0.AUTSYCAUT) & ": en-cours au-delà de la date limite d'autorisation  : EUR ";
            X = Format$(arrJ_MTE(arrJ_K), "### ### ### ###.00")
            XPrt.CurrentX = prtMinX + 11600 - XPrt.TextWidth(X)
            XPrt.Print X;
            Exit For
        Else
            If arrJ_MTE(arrJ_K) > arrJ_AUTSYCMON Then
                XPrt.CurrentY = XPrt.CurrentY + 80
                prtSAB_TC_Lmites_NewLine
                prtFillColor = mColor_Y1
                Call frmElpPrt.prtTrame_Color(prtMinX + 4900, XPrt.CurrentY - 80, prtMinX + 11700, XPrt.CurrentY + prtHeaderHeight - 80, " ")
                XPrt.ForeColor = vbRed
                XPrt.FontBold = True
                X = Format$(arrJ_MTE(arrJ_K), "### ### ### ###.00")

                XPrt.CurrentX = prtMinX + 4900
                XPrt.Print Trim(xZAUTSYC0.AUTSYCAUT) & ": à partir du " & dateImp(arrJ_AMJ_7c(arrJ_K) + 19000000) _
                 & ", l'en-cours " & X & " est supérieur au montant autorisé : EUR " & Format$(arrJ_AUTSYCMON, "### ### ### ##0");
                'XPrt.CurrentX = prtMinX + 11600 - XPrt.TextWidth(X)
                'XPrt.Print X;
                Exit For
            End If
        End If
        
    Next arrJ_K
End If
'_____________________________________________________________________________________

XPrt.FontBold = False
blnId_Print = False
End Sub

Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = fgSelect.Cols - 1
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

End Sub

Public Sub fgTotal_Reset()
fgTotal.Clear
fgTotal_Sort1 = 0: fgTotal_Sort2 = 0
fgTotal_Sort1_Old = -1
fgTotal_RowDisplay = 0: fgTotal_RowClick = 0
fgTotal_arrIndex = fgTotal.Cols - 1
blnfgTotal_DisplayLine = False
fgTotal_SortAD = 6
fgTotal.LeftCol = 0

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
Public Sub fgTotal_Sort()
If fgTotal.Rows > 1 Then
    fgTotal.Row = 1
    fgTotal.RowSel = fgTotal.Rows - 1
    
    If fgTotal_Sort1_Old = fgTotal_Sort1 Then
        If fgTotal_SortAD = 5 Then
            fgTotal_SortAD = 6
        Else
            fgTotal_SortAD = 5
        End If
    Else
        fgTotal_SortAD = 5
    End If
    fgTotal_Sort1_Old = fgTotal_Sort1
    
    fgTotal.Col = fgTotal_Sort1
    fgTotal.ColSel = fgTotal_Sort2
    fgTotal.Sort = fgTotal_SortAD
End If

End Sub

Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
Dim wK As Integer
wK = lK
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    If lK = 2 Then
        fgSelect.Col = 2
        X = fgSelect.Text
        wK = 3
    Else
        X = ""
    End If
    
    fgSelect.Col = wK
    X = X & Format$(Val(fgSelect.Text), "000000000000000.00")
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub

Public Sub fgTotal_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgTotal.Rows - 1
    fgTotal.Row = I
    If lK = 2 Then
        fgTotal.Col = 2
        X = fgTotal.Text
    Else
        X = ""
    End If
    
    fgTotal.Col = 3
    X = X & Format$(Val(fgTotal.Text), "000000000000000.00")
    fgTotal.Col = fgTotal_arrIndex - 1
    fgTotal.Text = X
Next I


fgTotal_Sort1 = fgTotal_arrIndex - 1: fgTotal_Sort2 = fgTotal_arrIndex - 1
fgTotal_Sort
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

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), SAB_TC_Limites_Aut)
If UCase$(Trim(Mid$(Msg, 1, 12))) = "@TC_LIMITES" Then
    blnAuto = True
Else
    blnAuto = False
End If
Form_Init

Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@TC_LIMITES":
                
                'blnAuto = True
                cmdSelect_SQL
                Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S54", "BIA-TC-LIMITES", "Archive")
                If xlsManual Then
                    Call cmdPrint_Ok_xlsManual
                Else
                    cmdPrint_Ok
                End If
                Unload Me
                
End Select
End Sub


Public Sub Form_Init()
Dim K As Integer, K2 As Integer, xWhere As String, xSQL As String, X As String
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistant", vbCritical, "frmZTREOPE0.param_init"
    Unload Me
Else
    lstErr.Clear
End If

blnControl = False
cmdReset
fraTab0.Visible = False

fgSelect_FormatString = fgSelect.FormatString
fgDossier_FormatString = fgDossier.FormatString
fgTotal_FormatString = fgTotal.FormatString
fgSelect.Enabled = True

fraDétail.Top = 1700
fraDétail.Left = 1900

'______________________________________________________________________
fgDevF_EUR_Row = 0
blnDevF_Init = False
DevF_Load
Call DTPicker_Set(txtDevF_AMJ, YBIATAB0_DATE_CPT_JS1)
Devf_Display
blnDevF_Init = True

Call DTPicker_Set(txtSelect_AMJMIN, YBIATAB0_DATE_CPT_JS1)
wAmj1_8C = ""
wAmjMax_8C = ""
For K = 1 To 31
    If arrDevF_Ouvré(K) > 0 Then
        If Trim(wAmj1_8C) = "" Then
            wAmj1_8C = arrDevF_Ouvré(K)
        Else
            If Trim(wAmjMax_8C) = "" Then
                wAmjMax_8C = arrDevF_Ouvré(K)
                Exit For
            End If
        End If
    End If
Next K

'wAmj1_8C = dateElp("Ouvré", 1, YBIATAB0_DATE_CPT_JS1)
Call DTPicker_Set(txtSelect_Amj1, wAmj1_8C)
'wAmjMax_8C = dateElp("Ouvré", 2, YBIATAB0_DATE_CPT_JS1)
Call DTPicker_Set(txtSelect_AmjMax, wAmjMax_8C)

mAMJ_Fixing = ""
Call param_Init_Fixing(YBIATAB0_DATE_CPT_JS1, YBIATAB0_DATE_CPT_J)

'______________________________________________________________________

'SSTab1.Tab = 2
'______________________________________________________________________


lblDevF_Warning = "             prochains jours fériés ( J + 7 )" & vbCrLf & "_____________________________" & vbCrLf
If blnDevF_Warning Then
    lblDevF_Warning.ForeColor = vbRed
    For K = 1 To 7
        If arrDevF_Warning(K) <> "" Then
            lblDevF_Warning = lblDevF_Warning & arrDevF_Warning(K) & vbCrLf & vbCrLf
        End If
    Next K
Else
    lblDevF_Warning.ForeColor = vbBlue
    lblDevF_Warning = lblDevF_Warning & vbCrLf & vbCrLf & "                         NEANT"
End If

'______________________________________________________________________


lblECH_Warning = "Opérations échues non comptabilisées" & vbCrLf & "_____________________________" & vbCrLf

xSQL = "select * from " & paramIBM_Library_SAB & ".ZTREOPE0" _
     & " where TREOPEETA <= '5'  and TREOPEECH < " & YBIATAB0_DATE_CPT_JS1 - 19000000 _
     & " order by TREOPECLI, TREOPEOPR"
Set rsSab = cnsab.Execute(xSQL)
K = 0
Do While Not rsSab.EOF
    K = K + 1
    X = rsSab("TREOPECLI") & "  " & rsSab("TREOPEOPR") & " " & rsSab("TREOPENUM") & " ech : " & dateIBM10(rsSab("TREOPEECH"), True) _
      & vbCrLf & " " & Format$(rsSab("TREOPEMNT"), "### ### ### ###.00") & " " & rsSab("TREOPEDEV")
    lblECH_Warning = lblECH_Warning & X & vbCrLf & vbCrLf
    rsSab.MoveNext
Loop

If K > 0 Then
    lblECH_Warning.ForeColor = vbRed
    chkSelect_AmjMin.Value = "1"
    If K = 1 Then
        X = " opération échue non tombée"
    Else
        X = " opérations échues non tombées"
    End If
    If Not blnAuto Then
        Call MsgBox("Il y a " & K & X & vbCrLf & vbCrLf & " la case 'date d'arrêté de l'encours' est cochée.", vbExclamation, "TC_Limites : contrôle")
    End If
Else
    lblECH_Warning.ForeColor = vbBlue
    lblECH_Warning = lblECH_Warning & vbCrLf & vbCrLf & "                         NEANT"
End If



cmdSelect_SQL
fraTab0.Visible = True


Me.Enabled = True
Me.MousePointer = 0
End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
lblSelect_Client.ForeColor = vbBlue
'lblSelect_Client.BackColor = vbCyan
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
currentAction = ""
rsZTREOPE0_Init meZTREOPE0
xZTREOPE0 = meZTREOPE0
fraSelect_Options.Enabled = False
'cmdSelect_Ok_Click
chkSelect_EnCours = "1"
chkSelect_AmjMin = "0"
txtSelect_AMJMIN.Visible = False
'chkSelect_AmjMin.Enabled = False


cboSelect_TREOPEOPR = "PRE"
meCV1.DeviseN = 0
meCV1.Montant = 0

meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
meCV2.OpéAmj = YBIATAB0_DATE_CPT_J
blnControl = True



End Sub


Public Function param_Init()

param_Init = Null
Call lstErr_Clear(lstErr, cmdContext, ". SAb_sTOCK_Import cbo"): DoEvents

fgSelect.Visible = False

wAmjM1_7C = dateIBM(dateElp("MoisAdd", 1, DSys))

Call rsYBIATAB0_cboK2("DEVISE", "ISO", cboSelect_TREOPEDEV)

cboSelect_TREOPEAUT.AddItem "   "
cboSelect_TREOPEAUT.AddItem "PJJ"
cboSelect_TREOPEAUT.AddItem "PCT"
cboSelect_TREOPEAUT.AddItem "PLT"
cboSelect_TREOPEAUT.AddItem "SUB"
cboSelect_TREOPEAUT.ListIndex = 0

cboSelect_TREOPEOPR.AddItem "   "
cboSelect_TREOPEOPR.AddItem "EMP"
cboSelect_TREOPEOPR.AddItem "PRE"
cboSelect_TREOPEOPR.ListIndex = 0

'cboSelect_TREOPENAT.AddItem "   "
'cboSelect_TREOPENAT.ListIndex = 0

'cboSelect_TREOPECLI.AddItem "   "
'cboSelect_TREOPECLI.ListIndex = 0
lstParam_AUT_load

fgSelect.Visible = True

Call lstErr_ChangeLastItem(lstErr, cmdContext, "= SAb_  Stock_Import"): DoEvents

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
V = sqlYBIATAB0_Delete(oldParam)
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
V = sqlYBIATAB0_Insert(newParam)
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


'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
On Error Resume Next
currentActiveControl_Name = C.Name
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub


'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
On Error Resume Next
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub


Private Sub chkSelect_AmjMin_Click()
If chkSelect_AmjMin = "1" Then
    txtSelect_AMJMIN.Visible = True
    chkSelect_AmjMin.ForeColor = vbRed
Else
    txtSelect_AMJMIN.Visible = False
    chkSelect_AmjMin.ForeColor = lblUsr.ForeColor
End If

End Sub

Private Sub chkSelect_EnCours_Click()
If chkSelect_EnCours = "1" Then
    lblSelect_EnCours_A.Visible = True
    lblSelect_EnCours_B.Visible = True
    txtSelect_EnCours.Visible = True
Else
    lblSelect_EnCours_A.Visible = False
    lblSelect_EnCours_B.Visible = False
    txtSelect_EnCours.Visible = False
End If
End Sub

Private Sub cmdEURBID_Ok_Click()
fraEURBID.Visible = False
arrEURBID(fgEUR.Row) = num_CDec(txtEURBID)
fgEUR.Col = 2: fgEUR.Text = Format(arrEURBID(fgEUR.Row), "###.######")
fgEUR.CellForeColor = vbMagenta
cmdSelect_SQL
End Sub

Private Sub cmdEURBID_Quit_Click()
fraEURBID.Visible = False
End Sub

Private Sub cmdParam_Add_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

X = Trim(txtParam)
If X = "" Or Len(X) <> 3 Then
    Call MsgBox("Préciser le code autorisation ( 3 caractères)", vbExclamation, "Paramétrage AUT à exclure")
Else
    newParam.BIATABID = "ZTREOPE0"
    newParam.BIATABK1 = "AUT"
    newParam.BIATABK2 = X
    newParam.BIATABTXT = ""
    If IsNull(Parametrage_New) Then lstParam_AUT_load
End If
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Delete_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass
X = Trim(txtParam)
If X <> Trim(lstParam) Then
    Call MsgBox("Le code  a été modifié," & vbCrLf & " la suppression n'est pas possible", vbCritical, "BIA_GOS : paramétrage")
Else
    oldParam.BIATABID = "ZTREOPE0"
    oldParam.BIATABK1 = "AUT"
    oldParam.BIATABK2 = X
    If IsNull(Parametrage_Delete) Then lstParam_AUT_load
End If


Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdParam_PCT_Delete_Click()
Dim V
If blnParam_PCT_Exist Then
    V = sqlYBIATAB0_Delete(oldParam_PCT)
End If
If Not IsNull(V) Then
    Call MsgBox(V, vbCritical, "Mise à jour Durée PCT")
Else
    Call cmdSelect_SQL
End If

End Sub

Private Sub cmdParam_PCT_Update_Click()
Dim V, Nb As Integer
newParam_PCT = oldParam_PCT
Nb = Val(txtParam_PCT)
If Nb = 0 Then
    Call MsgBox("préciser la durée maximale", vbInformation, "PCT")
Else
    newParam_PCT.BIATABTXT = Format$(Nb, "#0")
    If optParam_PCT_M Then
        Mid$(newParam_PCT.BIATABTXT, 4, 1) = "M"
    Else
        Mid$(newParam_PCT.BIATABTXT, 4, 1) = "J"
    End If
    
    If blnParam_PCT_Exist Then
        V = sqlYBIATAB0_Update(newParam_PCT, oldParam_PCT)
    Else
        V = sqlYBIATAB0_Insert(newParam_PCT)
    End If
    If Not IsNull(V) Then
        Call MsgBox(V, vbCritical, "Mise à jour Durée PCT")
    Else
        Call cmdSelect_SQL
    End If
End If
End Sub

Private Sub cmdParam_Quit_Click()
fraParam.Visible = False
End Sub

Private Sub cmdSelect_NOk_Click()
fraDétail.Visible = False
End Sub

Private Sub fgEUR_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next

On Error Resume Next
    If fgEUR.Rows > 1 Then
        fgEUR.Col = 2: txtEURBID = fgEUR.Text
        fgEUR.Col = 1: lblEURFIX = fgEUR.Text
        fgEUR.Col = 0: lblEUR = fgEUR.Text
        fraEURBID.Visible = True
        txtEURBID.SetFocus
        
   End If

End Sub


Private Sub fgTotal_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgTotal.RowHeightMin Then
    Select Case fgTotal.Col
        Case 0: fgTotal_Sort1 = 0: fgTotal_Sort2 = 2: fgTotal_Sort
        Case 1:  fgTotal_Sort1 = 1: fgTotal_Sort2 = 2: fgTotal_Sort
        Case 2: fgTotal_Sort1 = 2: fgTotal_Sort2 = 2: fgTotal_SortX 2
        Case 3: fgTotal_Sort1 = 3: fgTotal_Sort2 = 3: fgTotal_SortX 3
       Case fgTotal_arrIndex:  fgTotal_SortX fgTotal_arrIndex
    End Select
Else
    If fgTotal.Rows > 1 Then
        Call fgTotal_Color(fgTotal_RowClick, MouseMoveUsr.BackColor, fgTotal_ColorClick)
        fgTotal.Col = fgTotal_arrIndex:  K = CLng(fgTotal.Text)
        
   End If
End If
fgTotal.LeftCol = 0

End Sub


Private Sub lstParam_Click()
Dim xSQL As String
cmdParam_Delete.Visible = False
cmdParam_Add.Visible = SAB_TC_Limites_Aut.Valider
oldParam.BIATABK2 = ""
txtParam = Trim(Mid$(lstParam, 1, 10))

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'ZTREOPE0' and BIATABK1= 'AUT'  and BIATABK2= '" & txtParam & "'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    oldParam.BIATABTXT = rsSab("BIATABTXT")

    txtParam.Enabled = SAB_TC_Limites_Aut.Valider
    cmdParam_Delete.Visible = SAB_TC_Limites_Aut.Valider
    
Else
    txtParam = ""
End If

fraParam.Visible = True

End Sub

Private Sub txtDevF_AMJ_Change()
Devf_Display
End Sub

Private Sub txtEURBID_GotFocus()
Call txt_GotFocus(txtEURBID)

End Sub

Private Sub txtEURBID_KeyPress(KeyAscii As Integer)
Call num_KeyAsciiD(KeyAscii, txtEURBID)

End Sub


Private Sub txtEURBID_LostFocus()
Call txt_LostFocus(txtEURBID)

End Sub


Private Sub txtParam_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtParam_PCT_GotFocus()
Call txt_GotFocus(txtParam_PCT)
End Sub

Private Sub txtParam_PCT_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub


Private Sub txtParam_PCT_LostFocus()
Call txt_LostFocus(txtParam_PCT)

End Sub

Private Sub txtSelect_AmjMax_GotFocus()
DTPicker_GotFocus txtSelect_AmjMax

End Sub


Private Sub txtSelect_AmjMax_LostFocus()
DTPicker_LostFocus txtSelect_AmjMax

End Sub


Private Sub txtselect_amjmin_GotFocus()
DTPicker_GotFocus txtSelect_AMJMIN

End Sub

Private Sub txtselect_amjmin_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)

End Sub


Private Sub txtselect_amjmin_LostFocus()
DTPicker_LostFocus txtSelect_AMJMIN

End Sub










Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 0:
            If fgSelect.Rows > 1 Then
                cmdPrint_Ok
                
               ' Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
           End If
    Case 1:
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL()
Dim xSQL As String, xWhere As String, xAnd As String
Dim blnOk As Boolean, blnEnCours As Boolean
Dim I As Integer, nbDossier As Long
Dim K As Integer, blnGrp As Boolean
Dim wCLIGRPCLI As String, wCLIGRPREG As String
Dim mCLIGRPREG As String
Dim NbGroupe As Integer, blnGroupe_New As Boolean
Dim blnAUT_Ok As Boolean
ReDim arrZTREOPE0(101)
arrClient_Max = 100: arrClient_Nb = 0
blnOk = False
nbDossier = 0

fraDétail.Visible = False
'fgTotal.Visible = False
lblSelect_Client = ""

'' 15-06-2009 gestion des groupes : détournement du champ TREOPECOU (courtier) pour gérer la notion de groupe

' Case cochée = indication d'une date antérieure / Case vide = sans date
'TREOPEETA <= '5' : en cours
'TREOPEETA  = '6' : échu
'TREOPEETA  = '7' : annulé
'TREOPEETA  = '8' : comptabilisé puis annulé

mAUTSYCFIN_Min = dateElp("MoisAdd", -Val(txtSelect_EnCours), DSys) - 19000000

If chkSelect_AmjMin = "1" Then
    Call DTPicker_Control(txtSelect_AMJMIN, wAmj_Selection_8C)
    wAmj_Selection_7C = dateIBM(wAmj_Selection_8C)
    ' !! Parfois TREOPEECH peut = 0 alors il faudrait tester TREOPEREE
    xWhere_Selection = " where TREOPEETA <= '6' and TREOPENEG <=" & wAmj_Selection_7C & " and TREOPEECH >= " & wAmj_Selection_7C
Else
    
    wAmj_Selection_8C = YBIATAB0_DATE_CPT_JS1
    wAmj_Selection_7C = YBIATAB0_DIBM_CPT_JS1
    xWhere_Selection = " where TREOPEETA <= '5'"
End If
xWhere = xWhere_Selection

Call param_Init_Fixing(wAmj_Selection_8C, wAmj_Selection_8C)

Call DTPicker_Control(txtSelect_AmjMax, wAmjMax_8C)
wAmjMax_7C = dateIBM(wAmjMax_8C)
Call DTPicker_Control(txtSelect_Amj1, wAmj1_8C)
wAmj1_7C = dateIBM(wAmj1_8C)

'If wAmj_Selection_8C > wAmj1_8C Or wAmj_Selection_8C > wAmjMax_8C Or wAmj1_8C > wAmjMax_8C Then
If wAmj_Selection_8C >= wAmjMax_8C Then
   Call MsgBox("Incohérence des dates : " & vbCrLf _
            & "Date situation : " & dateImp10(wAmj_Selection_8C) & vbCrLf _
            & "Date J + 2     : " & dateImp10(wAmjMax_8C) _
            , vbCritical, "SAB_TC_Limites")

    Exit Sub
End If
X = Trim(txtSelect_TREOPECLI)

If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "TREOPECLI = '00" & X & "'"
End If
X = Trim(cboSelect_TREOPEAUT)

If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "TREOPEAUT = '" & X & "'"
End If
X = Trim(cboSelect_TREOPEOPR)
If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "TREOPEOPR = '" & X & "'"
End If
'x = Trim(cboSelect_TREOPENAT)
'If x <> "" Then
'    If xWhere = "" Then
'        xAnd = " where "
'    Else
'        xAnd = " and "
'    End If
'    xWhere = xWhere & xAnd & "TREOPENAT = '" & x & "'"
'End If
X = Trim(cboSelect_TREOPEDEV)
If X <> "" Then
    If xWhere = "" Then
        xAnd = " where "
    Else
        xAnd = " and "
    End If
    xWhere = xWhere & xAnd & "TREOPEDEV = '" & X & "'"
End If


xSQL = "select * from " & paramIBM_Library_SAB & ".ZTREOPE0" & xWhere & " order by TREOPECLI, TREOPEOPR"
Set rsSab = cnsab.Execute(xSQL)
libSelect.Caption = "Positions à : " & Time

Do While Not rsSab.EOF
    nbDossier = nbDossier + 1
    V = rsZTREOPE0_GetBuffer(rsSab, xZTREOPE0)
   
   If Not IsNull(V) Then
        MsgBox V, vbCritical, "frmSAB_Stock.cmdSelect_SQL : ZTREOPE0"
        Exit Sub
    Else
        blnEnCours = True
' ? échéance anticipée
        xZTREOPE0.TREOPECOU = xZTREOPE0.TREOPECLI '(pour groupe)

        If xZTREOPE0.TREOPEETA = "6" _
        And xZTREOPE0.TREOPEREE > 0 _
        And xZTREOPE0.TREOPEREE < wAmj_Selection_7C Then blnEnCours = False
        
        If blnEnCours Then
            If Not blnOk Then
                blnOk = True
                arrClient_Nb = 1
                arrZTREOPE0(1) = xZTREOPE0
                arrZTREOPE0(1).TREOPEMNT = 0
            Else
                If xZTREOPE0.TREOPECLI = arrZTREOPE0(arrClient_Nb).TREOPECLI _
                And xZTREOPE0.TREOPEOPR = arrZTREOPE0(arrClient_Nb).TREOPEOPR Then
                    '''arrZTREOPE0(arrClient_Nb).TREOPEMNT = arrZTREOPE0(arrClient_Nb).TREOPEMNT + xZTREOPE0.TREOPEMNT
            
                Else
                    arrClient_Nb = arrClient_Nb + 1
                    If arrClient_Nb > arrClient_Max Then
                        arrClient_Max = arrClient_Max + 50
                        ReDim Preserve arrZTREOPE0(arrClient_Max)
                    End If
                    
                    
                    arrZTREOPE0(arrClient_Nb) = xZTREOPE0
                    arrZTREOPE0(arrClient_Nb).TREOPEMNT = 0
                    arrZTREOPE0(arrClient_Nb).TREOPECMT = ""
                End If
    
            End If
            meCV_xTREOPE0
            arrZTREOPE0(arrClient_Nb).TREOPEMNT = arrZTREOPE0(arrClient_Nb).TREOPEMNT + meCV2.Montant
        End If
        
    End If
    rsSab.MoveNext
Loop

'______________________________________________________________________________ Groupe TRE

xWhere = " where CLIGRPREL = 'AUT'"
xSQL = "select count(*) as Tally   from " & paramIBM_Library_SAB & ".ZCLIGRP0" & xWhere
Set rsSab = cnsab.Execute(xSQL)
arrClient_Max = arrClient_Max + rsSab("Tally") + 1
ReDim Preserve arrZTREOPE0(arrClient_Max)
K = rsSab("Tally") + 1
ReDim arrZCLIGRP0(K)
arrZCLIGRP0_Nb = 0

NbGroupe = arrClient_Nb
mCLIGRPREG = ""

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIGRP0" & xWhere & " order by CLIGRPREG, CLIGRPCLI"
Set rsSab = cnsab.Execute(xSQL)
blnGrp = False
Do While Not rsSab.EOF
    wCLIGRPCLI = rsSab("CLIGRPCLI")
    wCLIGRPREG = rsSab("CLIGRPREG")
    
    arrZCLIGRP0_Nb = arrZCLIGRP0_Nb + 1
    arrZCLIGRP0(arrZCLIGRP0_Nb).CLIGRPCLI = wCLIGRPCLI
    arrZCLIGRP0(arrZCLIGRP0_Nb).CLIGRPREG = wCLIGRPREG
    arrZCLIGRP0(arrZCLIGRP0_Nb).CLIGRPCLI_RA2 = "-"

    
    If mCLIGRPREG <> wCLIGRPREG Then
        mCLIGRPREG = wCLIGRPREG
        blnGroupe_New = True
    End If
    For K = 1 To arrClient_Nb
        If wCLIGRPCLI = arrZTREOPE0(K).TREOPECLI Then
            arrZTREOPE0(K).TREOPECOU = wCLIGRPREG
            If blnGroupe_New Then
                blnGroupe_New = False
                NbGroupe = NbGroupe + 1
               arrZTREOPE0(NbGroupe) = arrZTREOPE0(K)
                arrZTREOPE0(NbGroupe).TREOPECLI = wCLIGRPREG
            Else
                arrZTREOPE0(NbGroupe).TREOPEMNT = arrZTREOPE0(NbGroupe).TREOPEMNT + arrZTREOPE0(K).TREOPEMNT
            End If
        End If
    Next K
    rsSab.MoveNext
Loop


 arrClient_Nb = NbGroupe
'______________________________________________________________________________

For K = 1 To arrZCLIGRP0_Nb
    xSQL = "select CLIENASIG from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENACLI = '" & arrZCLIGRP0(K).CLIGRPREG & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then arrZCLIGRP0(K).CLIGRPCLI_RA2 = Trim(rsSab("CLIENASIG"))
Next K

'______________________________________________________________________________
Suite:

ReDim arrZCLIENA0(arrClient_Max)


' Chargement AUTORISATION PIB & EMP
'===================================
ReDim arrZAUTSYC0(101) As typeZAUTSYC0
arrZAUTSYC0_Nb = 0: arrZAUTSYC0_Max = 100

'Select Case Trim(cboSelect_TREOPEOPR)
'    Case "PRE": xWhere = " where AUTSYCAUT = 'PIB'"
'    Case "EMP": xWhere = " where AUTSYCAUT = 'EMB'"
'    Case Else: xWhere = " where (AUTSYCAUT = 'PIB' or AUTSYCAUT = 'EMB')"
'End Select
Select Case Trim(cboSelect_TREOPEOPR)
    Case "PRE": xWhere = " where AUTSYCNIV = 1 and  AUTSYCTYP = '1' and ( AUTSYCAUT like 'P%' or AUTSYCAUT = 'TPL')"
    Case "EMP": xWhere = " where  AUTSYCNIV = 1 and  AUTSYCTYP = '1' and AUTSYCAUT like 'E%'"
    Case Else: xWhere = " where  AUTSYCNIV = 1 and  AUTSYCTYP = '1' and (AUTSYCAUT like 'P%' or AUTSYCAUT like 'E%' or AUTSYCAUT = 'TPL')"
End Select

X = Trim(txtSelect_TREOPECLI)

If X <> "" Then xWhere = xWhere & " and AUTSYCCLI = '00" & X & "'"

'$20070731 $JPL xSql = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0" & xWhere & " order by AUTSYCCLI"
xSQL = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0" & xWhere & " order by AUTSYCCLI, AUTSYCPER, AUTSYCADR"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsZAUTSYC0_GetBuffer(rsSab, xZAUTSYC0)
   '$JPL 20111129 If Mid$(xZAUTSYC0.AUTSYCAUT, 1, 3) <> "EGS" And Mid$(xZAUTSYC0.AUTSYCAUT, 1, 3) <> "PDI" Then
   If Trim(xZAUTSYC0.AUTSYCAUT) = "TPL" Then
        xZAUTSYC0.AUTSYCAUT = "PRE-TPL"
    End If
    blnAUT_Ok = True
    For K = 1 To arrParam_AUT_Nb
        If Mid$(xZAUTSYC0.AUTSYCAUT, 1, 3) = arrParam_AUT(K) Then
            blnAUT_Ok = False
            Exit For
        End If
    Next K
    If blnAUT_Ok Then
       If Not IsNull(V) Then
            MsgBox V, vbCritical, "frmSAB_Stock.cmdSelect_SQL : ZAUTSYC0"
        Else
            arrZAUTSYC0_Nb = arrZAUTSYC0_Nb + 1
            If arrZAUTSYC0_Nb > arrZAUTSYC0_Max Then
                arrZAUTSYC0_Max = arrZAUTSYC0_Max + 50
                ReDim Preserve arrZAUTSYC0(arrZAUTSYC0_Max)
            End If
            
    
            arrZAUTSYC0(arrZAUTSYC0_Nb) = xZAUTSYC0
    
            ' Chargement AUTORISATION ANTERIEURE PIB & EMP
            '=============================================
            If chkSelect_AmjMin = "1" Then  ' Date antérieure cochée
                AUTORISATIONS_ANTERIEURES
                arrZAUTSYC0(arrZAUTSYC0_Nb) = xZAUTSYC0
            End If
            
        End If
    End If
    rsSab.MoveNext
Loop

ReDim blnZAUTSYC0_EnCours(arrZAUTSYC0_Max)
For I = 1 To arrZAUTSYC0_Nb
    blnZAUTSYC0_EnCours(I) = False
Next I

'______________________________________________________________________________ PCT max

xWhere = " where BIATABID = 'ZTREOPE0' and BIATABK1 = 'PCT'"
xSQL = "select count(*) as Tally   from " & paramIBM_Library_SABSPE & ".YBIATAB0" & xWhere
Set rsSab = cnsab.Execute(xSQL)
arrParam_PCT_Nb = rsSab("Tally") + 1
K = 0
ReDim arrParam_PCT_CLI(arrParam_PCT_Nb), arrParam_PCT_AUT(arrParam_PCT_Nb)

xSQL = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0" & xWhere & " order by BIATABK2"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    K = K + 1
    arrParam_PCT_CLI(K) = Trim(rsSab("BIATABK2"))
    arrParam_PCT_AUT(K) = Trim(rsSab("BIATABTXT"))
    rsSab.MoveNext
Loop

Call arrJ_Init("AMJ")
Call arrJ_Init("MTE")

Call lstErr_AddItem(lstErr, cmdContext, "Lignes d'encours : " & nbDossier): DoEvents

fgSelect_Display

End Sub
Public Sub AUTORISATIONS_ANTERIEURES()
Dim xSQL As String
Dim V
Dim I As Long
Dim bln_Boucle As Boolean
Dim wAUTSYCDCR As Long

wAUTSYCDCR = xZAUTSYC0.AUTSYCDCR

ReDim arrZAUTHST0(51)
arrZAUTHST0_Nb = 0: arrZAUTHST0_Max = 50
xSQL = "select * from " & paramIBM_Library_SAB & ".ZAUTHST0 " _
    & "where AUTHSTETA = " & xZAUTSYC0.AUTSYCETA _
    & " and AUTHSTGPE = '" & xZAUTSYC0.AUTSYCGPE & "'" _
    & " and AUTHSTCLI = '" & xZAUTSYC0.AUTSYCCLI & "'" _
    & " and AUTHSTTYP = '" & xZAUTSYC0.AUTSYCTYP & "'" _
    & " and AUTHSTAUT = '" & xZAUTSYC0.AUTSYCAUT & "'" _
    & " order by AUTHSTMOD, AUTHSTSEQ "
    
'    & " and AUTHSTMOD <= " & wAmj_Selection_7C _
'    & " and AUTHSTDMO <= " & wAmj_Selection_7C _
'    & " order by AUTHSTDMO,AUTHSTMOD, AUTHSTSEQ "
    
Set rsSAB_X = cnsab.Execute(xSQL)

Do While Not rsSAB_X.EOF
    V = rsZAUTHST0_GetBuffer(rsSAB_X, xZAUTHST0)
    
    If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSAB_Stock.cmdSelect_SQL : ZAUTHST0"
    Else
         arrZAUTHST0_Nb = arrZAUTHST0_Nb + 1
         If arrZAUTHST0_Nb > arrZAUTHST0_Max Then
             arrZAUTHST0_Max = arrZAUTHST0_Max + 50
             ReDim Preserve arrZAUTHST0(arrZAUTHST0_Max)
         End If
         
         arrZAUTHST0(arrZAUTHST0_Nb) = xZAUTHST0
    End If
    
    rsSAB_X.MoveNext
Loop

' Appliquer les dernières modifs
If arrZAUTHST0_Nb > 0 Then
    wAUTSYCDCR = arrZAUTHST0(1).AUTHSTDCR
    For I = arrZAUTHST0_Nb To 1 Step -1
        If arrZAUTHST0(I).AUTHSTMOD > wAmj_Selection_7C Then
            xZAUTSYC0.AUTSYCMON = arrZAUTHST0(I).AUTHSTMON
            xZAUTSYC0.AUTSYCFIN = arrZAUTHST0(I).AUTHSTFIN
            xZAUTSYC0.AUTSYCDEV = arrZAUTHST0(I).AUTHSTDEV
        End If
    Next I
End If

' Date de création : en cours (
If wAUTSYCDCR > wAmj_Selection_7C Then
    xZAUTSYC0.AUTSYCMON = 0
    xZAUTSYC0.AUTSYCFIN = 0
    xZAUTSYC0.AUTSYCDEV = ""
End If

End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAb_Stock_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
fgDossier.Clear
fgTotal.Clear
fraDétail.Visible = False
If blnOk Then
    cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    cmdSelect_SQL
Else
    cmdSelect_Ok.Caption = "Exécuter la requête"
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAb_Stock_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
Dim xAUTSYCAUT As String

On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 3: fgSelect_SortX 2
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_SortX 3
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_SortX 4
        Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_SortX 5
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_SortX 6
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
        xZTREOPE0 = arrZTREOPE0(K)
        xZCLIENA0 = arrZCLIENA0(K)
        lblSelect_Client = Trim(xZCLIENA0.CLIENACLI & "     " & xZCLIENA0.CLIENARA1)
        Call fraParam_PCT_Display(xZCLIENA0.CLIENACLI)
        
        selZTREOPE0_SQL xZTREOPE0.TREOPECLI
        fgDossier_Display
        selZAUTSYC0_SQL xZCLIENA0.CLIENACLI 'xZTREOPE0.TREOPECLI
         fgSelect.Col = 3:  xAUTSYCAUT = Trim(fgSelect.Text)  'sans enCours sélectionner le code AUT
        selZAUTSYC0_Add xAUTSYCAUT
        fgTotal_Display
        
   End If
End If
fgSelect.LeftCol = 0
End Sub

Private Sub fgDossier_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgDossier.RowHeightMin Then
    Select Case fgDossier.Col
        Case 0: fgDossier_Sort1 = 0: fgDossier_Sort2 = 1: fgDossier_Sort
        Case 1:  fgDossier_Sort1 = 1: fgDossier_Sort2 = 1: fgDossier_Sort
        Case 2: fgDossier_Sort1 = 2: fgDossier_Sort2 = 2: fgDossier_Sort
        Case 3: fgDossier_Sort1 = 3: fgDossier_Sort2 = 3: fgDossier_Sort
        Case 4: fgDossier_Sort1 = 4: fgDossier_Sort2 = 4: fgDossier_Sort
    End Select
Else
    If fgDossier.Rows > 1 Then
        Call fgDossier_Color(fgDossier_RowClick, MouseMoveUsr.BackColor, fgDossier_ColorClick)
        fgDossier.Col = fgDossier_arrIndex:  K = CLng(fgDossier.Text)
       ' xZTREOPE0 = selZTREOPE0(K)

   End If
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
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select


End Sub

Public Sub cmdContext_Quit()
blnControl = False
lstErr.Clear: lstErr.Height = 200

If SSTab1.Tab = 1 And fraEURBID.Visible Then fraEURBID.Visible = False: Exit Sub

If fraParam.Visible Then fraParam.Visible = False: Exit Sub


If fraDétail.Visible Then fraDétail.Visible = False: Exit Sub
If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
Else
    SSTab1.Tab = SSTab1.Tab - 1
End If

End Sub

Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
    fgSelect.Row = fgSelect.TopRow
    fgSelect.Col = fgSelect_arrIndex: ' wK1 = fgSelect.Text
    'cmdSelect txtSelect ''fgSelect.Text

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

Exit Sub

Error_Handler:
blnControl = False
MsgBox Error
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





Private Sub mnuSelect_Print_Détail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Ok '"D "
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Ok '"L "
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then cmdSelect_Ok.SetFocus

End Sub














Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
    Case 1: fgDossier.LeftCol = 0
End Select
End Sub


Public Sub cmdPrint_Ok()
Dim iRow As Integer, K As Integer, I As Integer
Dim xAUTSYCAUT As String
Dim wText As String

fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Etat : " & fgSelect.Rows - 1)

wText = " au " & dateImp10(wAmj_Selection_8C) & " / spot : " & dateImp10(wAmjMax_8C)

If chkSelect_AmjMin = "1" Then
    wText = " *** !!!! date d'arrêté au " & dateImp10(wAmj_Selection_8C) & " !!!! ***"
End If
prtSAB_TC_Lmites_Open wText & "         ( édité le " & Now & " )", "Today", "mad " & dateImp10(wAmjMax_8C)

XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

If blnDevF_Warning Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
    XPrt.ForeColor = vbRed
    For K = 1 To 7
        If arrDevF_Warning(K) <> "" Then
            XPrt.CurrentX = prtMinX + 6000
            XPrt.Print "férié " & arrDevF_Warning(K);
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        End If
    Next K
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
End If

For iRow = 1 To fgSelect.Rows - 1
    
    fgSelect.Row = iRow
    fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
    xZTREOPE0 = arrZTREOPE0(K)
    xZCLIENA0 = arrZCLIENA0(K)
    
    selZTREOPE0_SQL arrZTREOPE0(K).TREOPECLI
    
    selZAUTSYC0_SQL arrZTREOPE0(K).TREOPECLI
    fgSelect.Col = 3:  xAUTSYCAUT = Mid$(Trim(fgSelect.Text), 1, 3) 'sans enCours sélectionner le code AUT
    selZAUTSYC0_Add xAUTSYCAUT
    
    blnId_Print = True
    For I = 0 To selZAUTSYC0_Nb
       ' Debug.Print selZAUTSYC0(I).AUTSYCAUT
        If selZAUTSYC0_Display(I) Then
            xZAUTSYC0 = selZAUTSYC0(I)
            XPrt.ForeColor = vbBlack
            fgTotal_PrintLine I

        End If
    Next I

    For I = 1 To selZTREOPE0_Nb
             
        xZTREOPE0 = selZTREOPE0(I)
        XPrt.ForeColor = vbBlue
       fgDossier_PrintLine (selZTREOPE0_Nb)
    Next I

Next iRow
prtSAB_TC_Lmites_Close
''prtSAB_Stock_Monitor lFct, fgSelect, arrZTREOPE0(), arrZCLIENA0(), arrClient_Nb
fgSelect.Visible = True
Me.Show
End Sub
Public Sub meCV_xTREOPE0()
meCV1.DeviseIso = xZTREOPE0.TREOPEDEV
meCV1.Montant = xZTREOPE0.TREOPEMNT
If meCV1.DeviseIso <> "EUR" Then
    meCV_arrEURBID
Else
    meCV2.Montant = xZTREOPE0.TREOPEMNT
End If
End Sub
Public Sub meCV_xAUTSYC0()

meCV1.DeviseIso = xZAUTSYC0.AUTSYCDEV
meCV1.Montant = xZAUTSYC0.AUTSYCMON
If meCV1.DeviseIso <> "EUR" Then
    meCV_arrEURBID
Else
    meCV2.Montant = xZAUTSYC0.AUTSYCMON
End If

Ope_Limite_Aut = meCV2.Montant
'Select Case meCV2.Montant
'    Case Is <= 20000000: Ope_Limite_Max_Code = "100%": Ope_Limite_Max = meCV2.Montant * 2
'    Case Is <= 40000000:  Ope_Limite_Max_Code = "80%": Ope_Limite_Max = meCV2.Montant * 1.8
'    Case Is <= 70000000:  Ope_Limite_Max_Code = "70%": Ope_Limite_Max = meCV2.Montant * 1.7
'    Case Is <= 100000000: Ope_Limite_Max_Code = "60%": Ope_Limite_Max = meCV2.Montant * 1.6
'    Case Else: Ope_Limite_Max_Code = "50%": Ope_Limite_Max = meCV2.Montant * 1.5
'End Select

'$JPL 2010-12-27
'Select Case meCV2.Montant
'    Case Is <= 20000000: Ope_Limite_Max_Code = "100%": Ope_Limite_Max = meCV2.Montant * 2
'    Case Is <= 40000000:  Ope_Limite_Max_Code = "80%": Ope_Limite_Max = meCV2.Montant * 1.8
'    Case Is <= 70000000:  Ope_Limite_Max_Code = "70%": Ope_Limite_Max = meCV2.Montant * 1.7
'    Case Is <= 100000000: Ope_Limite_Max_Code = "50%": Ope_Limite_Max = meCV2.Montant * 1.5
'    Case Is <= 120000000: Ope_Limite_Max_Code = "25%": Ope_Limite_Max = meCV2.Montant * 1.25
'    Case Else: Ope_Limite_Max_Code = "0%": Ope_Limite_Max = meCV2.Montant
'End Select

'$JPL 2011-10-07
'If xZAUTSYC0.AUTSYCCLI = "0050487" And meCV2.Montant = 100000000 Then _
'                        Ope_Limite_Max_Code = "25%": Ope_Limite_Max = meCV2.Montant * 1.25
'$JPL 2011-10-07

'$JPL 2012-03-06
'Select Case meCV2.Montant
'    Case Is <= 20000000: Ope_Limite_Max_Code = "100%": Ope_Limite_Max = meCV2.Montant * 2
'    Case Is < 40000000:  Ope_Limite_Max_Code = "50%": Ope_Limite_Max = meCV2.Montant * 1.5
'    Case Is <= 55000000:  Ope_Limite_Max_Code = "40%": Ope_Limite_Max = meCV2.Montant * 1.4
'    Case Else: Ope_Limite_Max_Code = "0%": Ope_Limite_Max = meCV2.Montant
'End Select

'$JPL 2012-03-07
Ope_Limite_Max_Code = xZAUTSYC0.AUTSYCTAU & "%"
Ope_Limite_Max = meCV2.Montant * (xZAUTSYC0.AUTSYCTAU / 100 + 1)

End Sub


Public Sub selZAUTSYC0_Add(lAUTSYCAUT As String)
Dim I As Long, K As Long, X As String
Dim KI As Long, blnPIB_Display As Boolean

' INITIALISATION
blnPIB_Display = False

Call arrJ_Init("MTE")
arrJ_K_Max = 0

ReDim selZAUTSYC0_Display(selZAUTSYC0_Nb)

ReDim Ope_ENG(selZAUTSYC0_Nb)
ReDim Ope_MAD(selZAUTSYC0_Nb)
ReDim ope_M0(selZAUTSYC0_Nb)
ReDim ope_M1(selZAUTSYC0_Nb)
ReDim ope_M2(selZAUTSYC0_Nb)

selZAUTSYC0(0).AUTSYCAUT = "???"
selZAUTSYC0(0).AUTSYCPER = 0
selZAUTSYC0(0).AUTSYCDEV = "???"

For I = 0 To selZAUTSYC0_Nb
    selZAUTSYC0_Display(I) = False
    If selZAUTSYC0_PIB(I) And selZAUTSYC0(I).AUTSYCMON > 0 Then selZAUTSYC0_Display(I) = True
    Ope_ENG(I) = 0
    Ope_MAD(I) = 0
    ope_M0(I) = 0
    ope_M1(I) = 0
    ope_M2(I) = 0
    If selZTREOPE0_Nb = 0 Then
        X = Mid$(Trim(selZAUTSYC0(I).AUTSYCAUT), 1, 3)
        If X = lAUTSYCAUT Then
                selZAUTSYC0_Display(I) = True
        End If
    End If

Next I

' Cumul des CV Eur par code autorisation
For K = 1 To selZTREOPE0_Nb
    blnPIB_Display = True
    
    xZTREOPE0 = selZTREOPE0(K)
    KI = 0
    X = Trim(xZTREOPE0.TREOPEAUT)
    If X = "TPL" Then X = "PRE-TPL"
    For I = 1 To selZAUTSYC0_Nb
        If X = Trim(selZAUTSYC0(I).AUTSYCAUT) Then
            KI = I
            Exit For
        End If
    Next I
    
    meCV_xTREOPE0
    selZAUTSYC0_Display(KI) = True
    
    '___________________________________________________ pb échéance passée non comptabilisée( cf 14-07)
    
    If xZTREOPE0.TREOPEECH >= wAmj_Selection_7C Then
    
    If xZTREOPE0.TREOPENEG <= wAmj_Selection_7C And xZTREOPE0.TREOPEECH >= wAmj_Selection_7C Then Ope_ENG(KI) = Ope_ENG(KI) + meCV2.Montant
    If xZTREOPE0.TREOPEDIS <= wAmj_Selection_7C And xZTREOPE0.TREOPEECH > wAmj_Selection_7C Then Ope_MAD(KI) = Ope_MAD(KI) + meCV2.Montant

    Select Case xZTREOPE0.TREOPEDIS
    
        Case Is <= wAmj_Selection_7C
            Select Case xZTREOPE0.TREOPEECH
                Case Is = wAmj_Selection_7C:
                Case Is > wAmjMax_7C
                                    ope_M0(KI) = ope_M0(KI) + meCV2.Montant
                                    ope_M1(KI) = ope_M1(KI) + meCV2.Montant
                                    ope_M2(KI) = ope_M2(KI) + meCV2.Montant
                Case Is = wAmjMax_7C
                                    ope_M0(KI) = ope_M0(KI) + meCV2.Montant
                                    ope_M1(KI) = ope_M1(KI) + meCV2.Montant
                Case Else
                                    ope_M0(KI) = ope_M0(KI) + meCV2.Montant
          End Select
        
        Case Is >= wAmjMax_7C: ope_M2(KI) = ope_M2(KI) + meCV2.Montant
        
        Case Else
            Select Case xZTREOPE0.TREOPEECH

                 Case Is >= wAmjMax_7C:
                                    ope_M1(KI) = ope_M1(KI) + meCV2.Montant
                                    ope_M2(KI) = ope_M2(KI) + meCV2.Montant
                Case Else:
                                    ope_M1(KI) = ope_M1(KI) + meCV2.Montant
            End Select
        End Select
    'End Select
    End If
'_____________________________________________________________________________________
    If selZAUTSYC0_PIB(KI) Then
        For arrJ_K = 0 To 100
            If xZTREOPE0.TREOPEECH <= arrJ_AMJ_7c(arrJ_K) Then
                If arrJ_K > arrJ_K_Max Then arrJ_K_Max = arrJ_K
                Exit For
            End If
            If xZTREOPE0.TREOPEDIS <= arrJ_AMJ_7c(arrJ_K) Then arrJ_MTE(arrJ_K) = arrJ_MTE(arrJ_K) + meCV2.Montant
        Next arrJ_K
        End If
Next K

' Cumul des opérations par niveau (père => Adresse)
' Niveau supérieur "PIB" ou "EMP"

For KI = selZAUTSYC0_Nb To 1 Step -1
     If selZAUTSYC0_Display(KI) Then
       X = Mid$(Trim(selZAUTSYC0(KI).AUTSYCAUT), 1, 3)
        '''If X = "PIB" Or X = "EMB" Then
        '''Else
           ' For I = KI - 1 To 1 Step -1
        If selZAUTSYC0(KI).AUTSYCPER > 1 Then
            For I = selZAUTSYC0_Nb To 1 Step -1
                If selZAUTSYC0(I).AUTSYCADR = selZAUTSYC0(KI).AUTSYCPER Then
                    selZAUTSYC0_Display(I) = True
                    Ope_ENG(I) = Ope_ENG(I) + Ope_ENG(KI)
                    If selZAUTSYC0(I).AUTSYCFIN = 0 Then selZAUTSYC0(I).AUTSYCFIN = selZAUTSYC0(KI).AUTSYCFIN
                    If selZAUTSYC0(I).AUTSYCMON = 0 Then selZAUTSYC0(I).AUTSYCMON = selZAUTSYC0(KI).AUTSYCMON
                   'If selZAUTSYC0_Total(I) = 0 Then
                        Ope_MAD(I) = Ope_MAD(I) + Ope_MAD(KI)
                        ope_M0(I) = ope_M0(I) + ope_M0(KI)
                        ope_M1(I) = ope_M1(I) + ope_M1(KI)
                        ope_M2(I) = ope_M2(I) + ope_M2(KI)
                         Exit For
                    'End If
                End If
            Next I
        End If
    End If
Next KI
If blnPIB_Display Then
    For K = 1 To selZAUTSYC0_Nb
        If Trim(selZAUTSYC0(K).AUTSYCAUT) = "PIB" Then selZAUTSYC0_Display(K) = True
    Next K

End If

End Sub


Public Sub xxxx_selZAUTSYC0_Add(lAUTSYCAUT As String)
Dim I As Long, K As Long, X As String
Dim KI As Long

' INITIALISATION
ReDim selZAUTSYC0_Display(selZAUTSYC0_Nb)

ReDim Ope_ENG(selZAUTSYC0_Nb)
ReDim Ope_MAD(selZAUTSYC0_Nb)
ReDim ope_M0(selZAUTSYC0_Nb)
ReDim ope_M1(selZAUTSYC0_Nb)
ReDim ope_M2(selZAUTSYC0_Nb)

selZAUTSYC0(0).AUTSYCAUT = "???"
selZAUTSYC0(0).AUTSYCPER = 0
selZAUTSYC0(0).AUTSYCDEV = "???"

For I = 0 To selZAUTSYC0_Nb
    selZAUTSYC0_Display(I) = False
    Ope_ENG(I) = 0
    Ope_MAD(I) = 0
    ope_M0(I) = 0
    ope_M1(I) = 0
    ope_M2(I) = 0
    If selZTREOPE0_Nb = 0 Then
        X = Trim(selZAUTSYC0(I).AUTSYCAUT)
        If X = lAUTSYCAUT Then
                selZAUTSYC0_Display(I) = True
        End If
    End If

Next I

' Cumul des CV Eur par code autorisation
For K = 1 To selZTREOPE0_Nb
    xZTREOPE0 = selZTREOPE0(K)
    KI = 0
    X = Trim(xZTREOPE0.TREOPEAUT)
    For I = 1 To selZAUTSYC0_Nb
        If X = Trim(selZAUTSYC0(I).AUTSYCAUT) Then
            KI = I
            Exit For
        End If
    Next I
    
    meCV_xTREOPE0
    selZAUTSYC0_Display(KI) = True
    
    '___________________________________________________ pb échéance passée non comptabilisée( cf 14-07)
    
    If xZTREOPE0.TREOPEECH >= wAmj_Selection_7C Then
    
    If xZTREOPE0.TREOPENEG <= wAmj_Selection_7C And xZTREOPE0.TREOPEECH >= wAmj_Selection_7C Then Ope_ENG(KI) = Ope_ENG(KI) + meCV2.Montant
    If xZTREOPE0.TREOPEDIS <= wAmj_Selection_7C And xZTREOPE0.TREOPEECH > wAmj_Selection_7C Then Ope_MAD(KI) = Ope_MAD(KI) + meCV2.Montant

    Select Case xZTREOPE0.TREOPEDIS
    
        Case Is <= wAmj_Selection_7C
            Select Case xZTREOPE0.TREOPEECH
                Case Is = wAmj_Selection_7C:
                Case Is > wAmjMax_7C
                                    ope_M0(KI) = ope_M0(KI) + meCV2.Montant
                                    ope_M1(KI) = ope_M1(KI) + meCV2.Montant
                                    ope_M2(KI) = ope_M2(KI) + meCV2.Montant
                Case Else
                                    ope_M0(KI) = ope_M0(KI) + meCV2.Montant
                                    ope_M1(KI) = ope_M1(KI) + meCV2.Montant
          End Select
        
        Case Is >= wAmjMax_7C: ope_M2(KI) = ope_M2(KI) + meCV2.Montant
        
        Case Else
            Select Case xZTREOPE0.TREOPEECH

                 Case Is >= wAmjMax_7C:
                                    'ope_M1(KI) = ope_M1(KI) + meCV2.Montant
                                    ope_M2(KI) = ope_M2(KI) + meCV2.Montant
                Case Else:
                                    ope_M1(KI) = ope_M1(KI) + meCV2.Montant
            End Select
        End Select
    'End Select
    End If
     
Next K

' Cumul des opérations par niveau (père => Adresse)
' Niveau supérieur "PIB" ou "EMP"

For KI = selZAUTSYC0_Nb To 1 Step -1
     If selZAUTSYC0_Display(KI) Then
       X = Mid$(Trim(selZAUTSYC0(KI).AUTSYCAUT), 1, 3)
        '''If X = "PIB" Or X = "EMB" Then
        '''Else
           ' For I = KI - 1 To 1 Step -1
        If selZAUTSYC0(KI).AUTSYCPER > 1 Then
            For I = selZAUTSYC0_Nb To 1 Step -1
                If selZAUTSYC0(I).AUTSYCADR = selZAUTSYC0(KI).AUTSYCPER Then
                    selZAUTSYC0_Display(I) = True
                    Ope_ENG(I) = Ope_ENG(I) + Ope_ENG(KI)
                    If selZAUTSYC0(I).AUTSYCFIN = 0 Then selZAUTSYC0(I).AUTSYCFIN = selZAUTSYC0(KI).AUTSYCFIN
                    If selZAUTSYC0(I).AUTSYCMON = 0 Then selZAUTSYC0(I).AUTSYCMON = selZAUTSYC0(KI).AUTSYCMON
                   'If selZAUTSYC0_Total(I) = 0 Then
                        Ope_MAD(I) = Ope_MAD(I) + Ope_MAD(KI)
                        ope_M0(I) = ope_M0(I) + ope_M0(KI)
                        ope_M1(I) = ope_M1(I) + ope_M1(KI)
                        ope_M2(I) = ope_M2(I) + ope_M2(KI)
                         Exit For
                    'End If
                End If
            Next I
        End If
    End If
Next KI

End Sub



Public Sub fgSelect_Display_Suite()
Dim curAut As Currency, X As String, xSQL As String
Dim V, I As Integer, wIndex As Integer, K As Integer

rsZTREOPE0_Init arrZTREOPE0(0)
ReDim Preserve arrZCLIENA0(500)
ReDim Preserve arrZTREOPE0(500)

For I = 1 To arrZAUTSYC0_Nb
    'Debug.Print I, arrZAUTSYC0(I).AUTSYCCLI
    If Not blnZAUTSYC0_EnCours(I) And arrZAUTSYC0(I).AUTSYCMON > 0 And arrZAUTSYC0(I).AUTSYCFIN > mAUTSYCFIN_Min Then
        xZAUTSYC0 = arrZAUTSYC0(I)
        meCV_xAUTSYC0
        curAut = meCV2.Montant
        
        arrClient_Nb = arrClient_Nb + 1
        wIndex = arrClient_Nb
        fgSelect_Display_ZCLIENA0 xZAUTSYC0.AUTSYCCLI, wIndex
        arrZCLIENA0(arrClient_Nb) = xZCLIENA0
        arrZTREOPE0(arrClient_Nb) = arrZTREOPE0(0)
        arrZTREOPE0(arrClient_Nb).TREOPECLI = xZAUTSYC0.AUTSYCCLI
        
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1

        If xZAUTSYC0.AUTSYCCLI < "0010000" Then
            arrZTREOPE0(arrClient_Nb).TREOPECOU = xZAUTSYC0.AUTSYCCLI
            fgSelect.Col = 1: fgSelect.Text = xZAUTSYC0.AUTSYCCLI
        End If
                fgSelect.CellForeColor = RGB(128, 128, 128)
                fgSelect.Col = 2: fgSelect.Text = xZAUTSYC0.AUTSYCCLI
                fgSelect.CellForeColor = RGB(128, 128, 128)

                fgSelect.Col = 3: fgSelect.Text = Trim(xZAUTSYC0.AUTSYCAUT) & " " & xZAUTSYC0.AUTSYCTAU & "%"
                 fgSelect.CellForeColor = RGB(128, 128, 128)
               
                fgSelect.Col = 0: fgSelect.Text = xZCLIENA0.CLIENARA2
                fgSelect.CellForeColor = RGB(128, 128, 128)
                fgSelect.Col = 5: fgSelect.Text = Format$(curAut, "### ### ### ###.00")
                fgSelect.CellForeColor = RGB(128, 128, 128)
                If xZAUTSYC0.AUTSYCFIN < wAmj_Selection_7C Then fgSelect.CellForeColor = vbRed
                For K = 1 To arrParam_PCT_Nb
                    If xZAUTSYC0.AUTSYCCLI = arrParam_PCT_CLI(K) Then
                        fgSelect.Col = 6: fgSelect.Text = arrParam_PCT_AUT(K)
                        'fgSelect.CellBackColor = RGB(255, 220, 255) 'vbMagenta
                        Exit For
                    End If
                Next K
                fgSelect.CellForeColor = RGB(128, 128, 128)

              fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = arrClient_Nb
    End If
    
Next I

End Sub

Public Sub fgSelect_Display_ZCLIENA0(lCLIENACLI As String, lIndex As Integer)
Dim X As String, xSQL As String
Dim K As Long

X = "CLIENACLI =  '" & lCLIENACLI & "'"
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 where " & X

Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    V = rsZCLIENA0_GetBuffer(rsSab, xZCLIENA0)
    If Not IsNull(V) Then
        MsgBox V, vbCritical, "frmSAB_Stock.fgSelect_Display"
       '' Exit Sub
    Else
        If xZCLIENA0.CLIENACLI < "0010000" Then
            xZCLIENA0.CLIENAAGE = 2
            xZCLIENA0.CLIENARA2 = Trim(xZCLIENA0.CLIENASIG) & "****"
        Else
            xZCLIENA0.CLIENARA2 = xZCLIENA0.CLIENASIG
            xZCLIENA0.CLIENAAGE = 0
            For K = 1 To arrZCLIGRP0_Nb
                
                If arrZCLIGRP0(K).CLIGRPCLI = xZCLIENA0.CLIENACLI Then
                    xZCLIENA0.CLIENAAGE = 1
                    xZCLIENA0.CLIENARA2 = arrZCLIGRP0(K).CLIGRPCLI_RA2 & "-" & xZCLIENA0.CLIENASIG
                    Exit For
                End If
            Next K
        End If
        arrZCLIENA0(lIndex) = xZCLIENA0
    End If
End If

End Sub

Public Sub meCV_arrEURBID()
Dim K As Integer, blnOk As Boolean
Dim dblMontant As Double

If meCV1.Montant = 0 Then
    meCV1.Cours = 0
    meCV2.Montant = 0
    Exit Sub
End If
blnOk = False
For K = 1 To arrEUR_nb
    If meCV1.DeviseIso = arrEUR(K) Then
        blnOk = True
        dblMontant = Abs(meCV1.Montant) / arrEURBID(K)
        meCV2.Montant = Fix((dblMontant + 0.00500001) * 100) / 100
        If meCV1.Montant < 0 Then meCV2.Montant = -meCV2.Montant
        Exit Sub
    End If
Next K

Call CV_Calc("J", meCV1, meCV2)

End Sub


Public Sub fgTotal_Limites(lIndex As Integer)
Dim curX1 As Currency, curX2 As Currency

Ope_Limite = Ope_Limite_Max - Ope_ENG(lIndex)
Ope_LimiteX = Ope_Limite
curX1 = Ope_Limite_Aut - ope_M0(lIndex)
Ope_PJJ = IIf(curX1 < Ope_Limite, curX1, Ope_Limite)

curX1 = Ope_Limite_Aut - ope_M1(lIndex)
Ope_PTM = IIf(curX1 < Ope_Limite, curX1, Ope_Limite)

curX1 = Ope_Limite_Aut - ope_M2(lIndex)
Ope_PCT = IIf(curX1 < Ope_Limite, curX1, Ope_Limite)


curX2 = IIf(Ope_PCT > Ope_PJJ, Ope_PCT, Ope_PJJ)
Ope_Limite = IIf(Ope_Limite < curX2, Ope_Limite, curX2)

If Ope_LimiteX = Ope_Limite Then
    blnLimite_Max_Code = True
Else
    blnLimite_Max_Code = False
End If

End Sub

Public Sub Devf_Display()
Dim K As Integer, K2 As Integer, Nbj As Long
Dim V, V2, X As String, xFormat As String
Dim wDev As String, mColor As Long

If Not blnDevF_Init Then
    blnDevF_Warning = False
    For K = 1 To 7
        arrDevF_Warning(K) = ""
    Next K
End If

wDev = ""
fgDevF.Visible = False
fgDevF.Clear
fgDevF.Rows = 1
Call DTPicker_Control(txtDevF_AMJ, X)
mColor = RGB(200, 255, 200)
V = dateImp10_S(X)
V2 = V

For K = 1 To arrDevF_Nb
    If wDev <> arrDevF_ISO(K) Then
        fgDevF.Rows = fgDevF.Rows + 1
        wDev = arrDevF_ISO(K)
        fgDevF.Row = fgDevF.Rows - 1
        fgDevF.Col = 0
        fgDevF.Text = wDev
        If wDev = "EUR" Then
            For K2 = 1 To 31
                fgDevF.Col = K2
                fgDevF.CellBackColor = mColor
            Next K2
            If Not blnDevF_Init Then fgDevF_EUR_Row = fgDevF.Row
       End If
    End If
    
    Nbj = DateDiff("d", V, arrDevF_AMJx(K))
    If Nbj >= 0 And Nbj < 31 Then
    
        fgDevF.Col = Nbj + 1
        fgDevF.CellBackColor = RGB(255, 128, 128)
        fgDevF.Text = Mid$(arrDevF_AMJx(K), 1, 2)
        
        If Not blnDevF_Init Then
            If Nbj >= 0 And Nbj < 7 Then
                blnDevF_Warning = True
                If arrDevF_Warning(Nbj) = "" Then
                   arrDevF_Warning(Nbj) = arrDevF_AMJx(K) & " : " & wDev & " "
               Else
                    arrDevF_Warning(Nbj) = arrDevF_Warning(Nbj) & wDev & " "
                End If
            End If
        End If
    End If
    
    
Next K
xFormat = " Devise|"

For K = 1 To 31
    arrDevF_Ouvré(K - 1) = Mid$(V2, 7, 4) & Mid$(V2, 4, 2) & Mid$(V2, 1, 2) 'V2
    Nbj = Weekday(V2) * 2 - 1
    xFormat = xFormat & Mid$("DiLuMaMeJeVeSa", Nbj, 2) & vbCr & Mid$(V2, 1, 2) & "|"
    If Nbj = 1 Or Nbj = 13 Then
        fgDevF.Col = K
        For K2 = 0 To fgDevF.Rows - 1
            fgDevF.Row = K2
            If Trim(fgDevF.Text) = "" Then fgDevF.CellBackColor = RGB(255, 255, 200)
        Next K2
    End If
    V2 = DateAdd("d", 1, V2)
Next K
fgDevF.FormatString = xFormat

For K = 1 To 31
    fgDevF.ColWidth(K) = 400
Next K

fgDevF.Row = 0: fgDevF.RowHeight(0) = 400
If Not blnDevF_Init Then
    fgDevF.Row = fgDevF_EUR_Row
    For K = 1 To 31
        fgDevF.Col = K
        If fgDevF.CellBackColor <> mColor Then
            arrDevF_Ouvré(K - 1) = 0
        End If
    Next K
End If

fgDevF.Visible = True
End Sub

Public Sub param_Init_Fixing(lAMJ As String, lAMJ_1 As String)
Dim xWhere As String, xSQL As String
Dim K As Integer, K2 As Integer
'____________________________________________________________ fixing jour / veille
If mAMJ_Fixing = lAMJ Then Exit Sub

mAMJ_Fixing = lAMJ

xWhere = " where BIATABID = 'PDC' and BIATABK2 = '" & lAMJ_1 & "'"
xSQL = "select count(*) as Tally   from " & paramIBM_Library_SABSPE & ".YBIATAB0" & xWhere
Set rsSab = cnsab.Execute(xSQL)
arrEUR_nb = rsSab("Tally")
ReDim Preserve arrEUR(arrEUR_nb + 1)
ReDim Preserve arrEURFIX(arrEUR_nb + 1)
ReDim Preserve arrEURBID(arrEUR_nb + 1)

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0" & xWhere & " order by BIATABK1"
Set rsSab = cnsab.Execute(xSQL)
K2 = 9
Do While Not rsSab.EOF
    X = Trim(rsSab("BIATABK1"))
    Select Case X
        Case "USD": K = 1
        Case "CHF": K = 2
        Case "GBP": K = 3
        Case "JPY": K = 4
        Case "CAD": K = 5
        Case "DKK": K = 6
        Case "NOK": K = 7
        Case "SEK": K = 8
        Case "AED": K = 9
        Case Else: K2 = K2 + 1: K = K2
    End Select
    arrEUR(K) = X
    arrEURFIX(K) = CDbl(Mid$(rsSab("BIATABTXT"), 9, 15) / 1000000000)
    arrEURBID(K) = arrEURFIX(K)
    rsSab.MoveNext
Loop

xWhere = " where BIATABID = 'PDC' and BIATABK2 = '" & lAMJ & "'"
xSQL = "select *   from " & paramIBM_Library_SABSPE & ".YBIATAB0" & xWhere
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    X = Trim(rsSab("BIATABK1"))
    For K = 1 To arrEUR_nb
        If arrEUR(K) = X Then
            arrEURFIX(K) = CDbl(Mid$(rsSab("BIATABTXT"), 9, 15) / 1000000000)
            arrEURBID(K) = arrEURFIX(K)
            Exit For
        End If
            
    Next K
    rsSab.MoveNext
Loop


lblDevF = "Cours connus au : " & dateImp10(lAMJ) & "   " & Time
fgEUR_Display
End Sub

Public Sub fraParam_PCT_Display(lCLI As String)
Dim xWhere As String, X As String, Nb As Integer
fraParam_PCT.Enabled = SAB_TC_Limites_Aut.Valider
oldParam_PCT.BIATABID = "ZTREOPE0"
oldParam_PCT.BIATABK1 = "PCT"
oldParam_PCT.BIATABK2 = lCLI
xWhere = " where BIATABID = 'ZTREOPE0' and BIATABK1 = 'PCT' and BIATABK2 = '" & lCLI & "'"
X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0" & xWhere
Set rsSab = cnsab.Execute(X)
If Not rsSab.EOF Then
    blnParam_PCT_Exist = True
    oldParam_PCT.BIATABTXT = rsSab("BIATABTXT")
    If Mid$(oldParam_PCT.BIATABTXT, 4, 1) = "M" Then
        optParam_PCT_M.Value = True
    Else
        optParam_PCT_J.Value = True
    End If
    
    txtParam_PCT = Trim(Val(Mid$(oldParam_PCT.BIATABTXT, 1, 2)))
    cmdParam_PCT_Delete.Visible = SAB_TC_Limites_Aut.Valider
    cmdParam_PCT_Update.Visible = SAB_TC_Limites_Aut.Valider
Else
    blnParam_PCT_Exist = False
    optParam_PCT_M.Value = True: optParam_PCT_J.Value = False
    txtParam_PCT = ""
    cmdParam_PCT_Delete.Visible = False
    cmdParam_PCT_Update.Visible = SAB_TC_Limites_Aut.Valider
End If
End Sub

Public Sub lstParam_AUT_load()
Dim xWhere As String, X As String, Nb As Integer

fgSelect.Clear
fgDossier.Clear
fgTotal.Clear
fraDétail.Visible = False
fraParam.Visible = False

xWhere = " where BIATABID = 'ZTREOPE0' and BIATABK1 = 'AUT' "
X = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YBIATAB0" & xWhere
Set rsSab = cnsab.Execute(X)
Nb = rsSab("Tally") + 1
ReDim arrParam_AUT(Nb)
lstParam.Clear
arrParam_AUT_Nb = 0
X = "select *  from " & paramIBM_Library_SABSPE & ".YBIATAB0" & xWhere
Set rsSab = cnsab.Execute(X)
Do Until rsSab.EOF
    arrParam_AUT_Nb = arrParam_AUT_Nb + 1
    arrParam_AUT(arrParam_AUT_Nb) = Trim(rsSab("BIATABK2"))
    lstParam.AddItem arrParam_AUT(arrParam_AUT_Nb)
    rsSab.MoveNext
Loop
cmdParam_Delete.Visible = SAB_TC_Limites_Aut.Valider
cmdParam_Add.Visible = SAB_TC_Limites_Aut.Valider
If arrParam_AUT_Nb = 0 Then fraParam.Visible = True


End Sub



Public Sub arrJ_Init(lFct As String)
Dim K As Integer
Select Case lFct
    Case "AMJ":
        Dim X8 As String
        
        arrJ_AMJ_7c(0) = wAmj_Selection_7C
        X8 = wAmj_Selection_8C
        For K = 1 To 100
            X8 = DateAdd_AMJ("d", 1, X8)
            arrJ_AMJ_7c(K) = Val(X8) - 19000000
        Next K
    Case "MTE":
        For K = 0 To 100: arrJ_MTE(K) = 0: Next K
        
End Select

End Sub

Private Sub txtSelect_EnCours_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


