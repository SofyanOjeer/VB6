VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBDF_CRT 
   AutoRedraw      =   -1  'True
   Caption         =   "BDF_CRT"
   ClientHeight    =   10305
   ClientLeft      =   120
   ClientTop       =   420
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
   Icon            =   "BDF_CRT.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10305
   ScaleWidth      =   13530
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
      Left            =   6120
      TabIndex        =   2
      Top             =   0
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9720
      Left            =   0
      TabIndex        =   3
      Top             =   450
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
      TabCaption(0)   =   "Compte_rendu de transactions (DGS n° 09-01)"
      TabPicture(0)   =   "BDF_CRT.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Paramétrage"
      TabPicture(1)   =   "BDF_CRT.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstParam_Id"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstParam_K"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraParam_Display"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "."
      TabPicture(2)   =   "BDF_CRT.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame fraParam_Display 
         BackColor       =   &H00CDEBFF&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6060
         Left            =   -68370
         TabIndex        =   12
         Top             =   3000
         Visible         =   0   'False
         Width           =   6405
         Begin VB.Frame fraParam_CRT_Mvt_Ann 
            BackColor       =   &H00CDEBFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   3030
            Left            =   1905
            TabIndex        =   37
            Top             =   1440
            Visible         =   0   'False
            Width           =   6150
            Begin VB.CommandButton cmdParam_CRT_Mvt_Ann 
               BackColor       =   &H000000FF&
               Caption         =   "Appliquer => Annulation des mouvements A REVOIR"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   960
               Left            =   3105
               MaskColor       =   &H0080FFFF&
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   1605
               Width           =   2580
            End
            Begin VB.TextBox txtParam_CRTMVTCPT 
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
               Left            =   1700
               TabIndex        =   42
               Top             =   500
               Width           =   3060
            End
            Begin VB.TextBox txtParam_CRTMVTSSE 
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
               Left            =   1700
               TabIndex        =   41
               Top             =   1500
               Width           =   645
            End
            Begin VB.TextBox txtParam_CRTMVTOPE 
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
               Left            =   1700
               TabIndex        =   40
               Top             =   2000
               Width           =   645
            End
            Begin VB.TextBox txtParam_CRTMVTEVE 
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
               Left            =   1700
               TabIndex        =   39
               Top             =   2500
               Width           =   645
            End
            Begin VB.TextBox txtParam_CRTMVTSER 
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
               Left            =   1700
               TabIndex        =   38
               Top             =   1000
               Width           =   645
            End
            Begin VB.Label lblParam_CRTMVTCPT 
               BackColor       =   &H00CDEBFF&
               Caption         =   "Compte"
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
               Left            =   140
               TabIndex        =   47
               Top             =   500
               Width           =   1050
            End
            Begin VB.Label lblParam_CRTMVTSER 
               BackColor       =   &H00CDEBFF&
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
               Height          =   225
               Left            =   140
               TabIndex        =   46
               Top             =   1000
               Width           =   1050
            End
            Begin VB.Label lblParam_CRTMVTSSE 
               BackColor       =   &H00CDEBFF&
               Caption         =   "Sous-service"
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
               Left            =   140
               TabIndex        =   45
               Top             =   1500
               Width           =   1050
            End
            Begin VB.Label lblParam_CRTMVTOPE 
               BackColor       =   &H00CDEBFF&
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
               Height          =   225
               Left            =   140
               TabIndex        =   44
               Top             =   2000
               Width           =   1410
            End
            Begin VB.Label lblParam_CRTMVTEVE 
               BackColor       =   &H00CDEBFF&
               Caption         =   "Evenement"
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
               Left            =   140
               TabIndex        =   43
               Top             =   2500
               Width           =   1050
            End
         End
         Begin VB.ComboBox cboParam_Nomenclature 
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
            Left            =   465
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   2500
            Width           =   5475
         End
         Begin VB.TextBox txtParam_K2 
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
            Left            =   2520
            MaxLength       =   10
            TabIndex        =   21
            Top             =   1200
            Width           =   2040
         End
         Begin VB.TextBox txtParam_Txt 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1155
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   18
            Text            =   "BDF_CRT.frx":035E
            Top             =   3500
            Width           =   5350
         End
         Begin VB.TextBox txtParam_K1 
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
            Left            =   2500
            MaxLength       =   12
            TabIndex        =   17
            Top             =   500
            Width           =   2040
         End
         Begin VB.CommandButton cmdParam_Delete 
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
            Left            =   465
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   5175
            Width           =   1200
         End
         Begin VB.CommandButton cmdParam_Add 
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
            Height          =   600
            Left            =   1860
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   5205
            Width           =   1200
         End
         Begin VB.CommandButton cmdParam_Update 
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
            Left            =   3465
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   5220
            Width           =   1200
         End
         Begin VB.CommandButton cmdParam_Quit 
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
            Left            =   4980
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   5220
            Width           =   1200
         End
         Begin VB.Label lblParam_Txt 
            BackColor       =   &H00CDEBFF&
            Caption         =   "Libellé :"
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
            Left            =   345
            TabIndex        =   36
            Top             =   3100
            Width           =   5580
         End
         Begin VB.Label lblParam_Nomenclature 
            BackColor       =   &H00CDEBFF&
            Caption         =   "Rubrique appartement à la nomenclature :"
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
            Left            =   330
            TabIndex        =   34
            Top             =   2100
            Width           =   5580
         End
         Begin VB.Label lblParam_K1 
            BackColor       =   &H00CDEBFF&
            Caption         =   "K1"
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
            TabIndex        =   20
            Top             =   500
            Width           =   1290
         End
         Begin VB.Label lblParam_K2 
            BackColor       =   &H00CDEBFF&
            Caption         =   "K2"
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
            Left            =   330
            TabIndex        =   19
            Top             =   1200
            Width           =   2070
         End
      End
      Begin VB.ListBox lstParam_K 
         BackColor       =   &H00F0FFFF&
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
         Left            =   -74790
         TabIndex        =   11
         Top             =   3000
         Width           =   12885
      End
      Begin VB.ListBox lstParam_Id 
         BackColor       =   &H00CDEBFF&
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
         Left            =   -74790
         TabIndex        =   10
         Top             =   800
         Width           =   4785
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
         Top             =   400
         Width           =   13425
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   7635
            Left            =   2295
            TabIndex        =   9
            Top             =   1545
            Visible         =   0   'False
            Width           =   10545
            _ExtentX        =   18600
            _ExtentY        =   13467
            _Version        =   393216
            Cols            =   14
            FixedCols       =   0
            RowHeightMin    =   400
            BackColor       =   15794175
            ForeColor       =   4210752
            BackColorFixed  =   14213080
            ForeColorFixed  =   64
            BackColorBkg    =   15794175
            GridColor       =   4210816
            GridColorFixed  =   4210816
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   $"BDF_CRT.frx":0393
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
            Left            =   11820
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   705
            Width           =   1335
         End
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
            IntegralHeight  =   0   'False
            ItemData        =   "BDF_CRT.frx":0498
            Left            =   9255
            List            =   "BDF_CRT.frx":049A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   300
            Width           =   4110
         End
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00F0FFFF&
            BorderStyle     =   0  'None
            Height          =   1305
            Left            =   40
            TabIndex        =   5
            Top             =   135
            Visible         =   0   'False
            Width           =   9165
            Begin VB.Frame fraSelect_Options_AMJ 
               BackColor       =   &H00F0FFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1100
               Left            =   7425
               TabIndex        =   30
               Top             =   120
               Width           =   1770
               Begin MSComCtl2.DTPicker txtSelect_AmjMin 
                  Height          =   300
                  Left            =   435
                  TabIndex        =   31
                  Top             =   255
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
                  Format          =   97910787
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin MSComCtl2.DTPicker txtSelect_AmjMax 
                  Height          =   300
                  Left            =   435
                  TabIndex        =   33
                  Top             =   705
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
                  Format          =   97910787
                  CurrentDate     =   38699.44875
                  MaxDate         =   401768
                  MinDate         =   36526.4425347222
               End
               Begin VB.Label lblSelect_AMJ 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "Période"
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
                  Left            =   585
                  TabIndex        =   32
                  Top             =   0
                  Width           =   795
               End
            End
            Begin VB.ComboBox cboSelect_CRTCPTSTA 
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
               Left            =   5910
               Sorted          =   -1  'True
               TabIndex        =   29
               Text            =   "cboSelect_App"
               Top             =   225
               Width           =   1440
            End
            Begin VB.ComboBox cboSelect_CRTCPTRUB 
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
               Left            =   990
               Sorted          =   -1  'True
               TabIndex        =   28
               Text            =   "cboSelect_App"
               Top             =   780
               Width           =   6495
            End
            Begin VB.TextBox txtSelect_COMPTEOBL 
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
               Left            =   3675
               TabIndex        =   27
               Top             =   255
               Width           =   1155
            End
            Begin VB.TextBox txtSelect_CRTCPTCPT 
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
               Left            =   1020
               TabIndex        =   23
               Top             =   255
               Width           =   1650
            End
            Begin VB.Label lblSelect_CRTCPTSTA 
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
               Height          =   180
               Left            =   5085
               TabIndex        =   26
               Top             =   255
               Width           =   795
            End
            Begin VB.Label lblSelect_CRTCPTRUB 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Rubrique"
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
               TabIndex        =   25
               Top             =   780
               Width           =   795
            End
            Begin VB.Label lblSelect_COMPTEOBL 
               BackColor       =   &H00F0FFFF&
               Caption         =   "PCEC"
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
               Left            =   2895
               TabIndex        =   24
               Top             =   255
               Width           =   570
            End
            Begin VB.Label lblSelect_CRTCPTCPT 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Compte"
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
               Left            =   270
               TabIndex        =   22
               Top             =   300
               Width           =   795
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7710
            Left            =   120
            TabIndex        =   8
            Top             =   1500
            Visible         =   0   'False
            Width           =   13140
            _ExtentX        =   23178
            _ExtentY        =   13600
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   400
            BackColor       =   16777215
            ForeColor       =   16711680
            BackColorFixed  =   8421376
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   $"BDF_CRT.frx":049C
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
      Begin TabDlg.SSTab SSTab3 
         Height          =   9165
         Left            =   -74970
         TabIndex        =   49
         Top             =   495
         Visible         =   0   'False
         Width           =   13395
         _ExtentX        =   23627
         _ExtentY        =   16166
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "fraSelect"
         TabPicture(0)   =   "BDF_CRT.frx":0596
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtRTF"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtFg"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "fraSelect_Log"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "fraSelect_2"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "fraYCRTCPT0"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "fraYCRTMVT0"
         TabPicture(1)   =   "BDF_CRT.frx":05B2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraYCRTMVT0"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "fraSwift "
         TabPicture(2)   =   "BDF_CRT.frx":05CE
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraSwift"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "fraYTVACOM0"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "fgLog"
         TabPicture(3)   =   "BDF_CRT.frx":05EA
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "fraYCRTLOG0"
         Tab(3).Control(1)=   "fgLog"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "fraOD"
         TabPicture(4)   =   "BDF_CRT.frx":0606
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "fraOD"
         Tab(4).ControlCount=   1
         Begin VB.Frame fraOD 
            BackColor       =   &H00D8DFD8&
            Caption         =   "Mouvement extra-comptable"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7410
            Left            =   -71340
            TabIndex        =   130
            Top             =   1605
            Width           =   8370
            Begin VB.TextBox txtCRTMVTTAUX_OD 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
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
               Left            =   5595
               Locked          =   -1  'True
               TabIndex        =   156
               Top             =   2985
               Width           =   2055
            End
            Begin VB.TextBox txtCRTMVTMTE_OD 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   5640
               TabIndex        =   154
               Top             =   4200
               Width           =   2055
            End
            Begin VB.ComboBox cboCRTMVTDEV_OD 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1965
               Style           =   2  'Dropdown List
               TabIndex        =   152
               Top             =   2760
               Width           =   1215
            End
            Begin VB.OptionButton optCRTMVTMTD_OD_Cr 
               BackColor       =   &H00D8DFD8&
               Caption         =   "Crédit"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   270
               Left            =   4155
               TabIndex        =   151
               Top             =   4335
               Value           =   -1  'True
               Width           =   1005
            End
            Begin VB.OptionButton optCRTMVTMTD_OD_Db 
               BackColor       =   &H00D8DFD8&
               Caption         =   "Débit"
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
               Height          =   270
               Left            =   4170
               TabIndex        =   150
               Top             =   4020
               Width           =   1005
            End
            Begin VB.CommandButton cmdOD_Log 
               BackColor       =   &H00FFFFC0&
               Caption         =   "historique des modifications"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   5385
               Style           =   1  'Graphical
               TabIndex        =   149
               ToolTipText     =   "Cliquer ici pour afficher toutes les écritures comptables concernant ce dossier"
               Top             =   6500
               Width           =   1300
            End
            Begin VB.TextBox txtCRTMVTMTD_OD 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2000
               TabIndex        =   148
               Top             =   4095
               Width           =   2055
            End
            Begin VB.TextBox txtCRTMVTTXT_OD 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1365
               Left            =   2000
               MultiLine       =   -1  'True
               TabIndex        =   146
               Top             =   4740
               Width           =   5745
            End
            Begin VB.TextBox txtCRTMVTCPT_OD 
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
               Left            =   2000
               TabIndex        =   137
               Top             =   1800
               Width           =   1695
            End
            Begin VB.ComboBox cboCRTMVTCLIP_OD 
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
               Left            =   2000
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   136
               Top             =   1200
               Width           =   3255
            End
            Begin VB.ComboBox cboCRTMVTRUB_OD 
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
               Left            =   2000
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   135
               Top             =   600
               Width           =   5550
            End
            Begin VB.CommandButton cmdOD_Delete 
               BackColor       =   &H000000FF&
               Caption         =   "Annuler"
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
               Left            =   270
               Style           =   1  'Graphical
               TabIndex        =   134
               Top             =   6500
               Width           =   1200
            End
            Begin VB.CommandButton cmdOD_Add 
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
               Height          =   600
               Left            =   1800
               Style           =   1  'Graphical
               TabIndex        =   133
               Top             =   6500
               Width           =   1200
            End
            Begin VB.CommandButton cmdOD_Update 
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
               Left            =   3375
               Style           =   1  'Graphical
               TabIndex        =   132
               Top             =   6500
               Width           =   1200
            End
            Begin VB.CommandButton cmdOD_Quit 
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
               Left            =   6855
               Style           =   1  'Graphical
               TabIndex        =   131
               Top             =   6500
               Width           =   1200
            End
            Begin MSComCtl2.DTPicker txtCRTMVTDTR_OD 
               Height          =   300
               Left            =   1965
               TabIndex        =   138
               Top             =   3315
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
               Format          =   97910787
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblCRTMVTTAUX_OD 
               BackColor       =   &H0080FFFF&
               Caption         =   "Cours / €"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   4185
               TabIndex        =   155
               Top             =   3045
               Width           =   1000
            End
            Begin VB.Label Label1 
               BackColor       =   &H00D8DFD8&
               Caption         =   "CV EUR"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   6210
               TabIndex        =   153
               Top             =   3720
               Width           =   840
            End
            Begin VB.Label lblCRTMVTDEV_OD 
               BackColor       =   &H0080FFFF&
               Caption         =   "Devise"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   300
               TabIndex        =   147
               Top             =   2715
               Width           =   1000
            End
            Begin VB.Label lblCRTMVTTXT_OD 
               BackColor       =   &H00D8DFD8&
               Caption         =   "Commentaire"
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
               Left            =   375
               TabIndex        =   145
               Top             =   4785
               Width           =   1215
            End
            Begin VB.Label lblCRTMVTDTR_OD 
               BackColor       =   &H0080FFFF&
               Caption         =   "Date"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   255
               TabIndex        =   144
               Top             =   3345
               Width           =   1000
            End
            Begin VB.Label lblCRTMVTMTD_OD 
               BackColor       =   &H00D8DFD8&
               Caption         =   "Montant en devise"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   585
               Left            =   255
               TabIndex        =   143
               Top             =   3870
               Width           =   1125
            End
            Begin VB.Label libCRTMVTCPT_OD 
               BackColor       =   &H00D8DFD8&
               Caption         =   "compte"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   4200
               TabIndex        =   142
               Top             =   1860
               Width           =   4005
            End
            Begin VB.Label lblCRTMVTCPT_OD 
               BackColor       =   &H00D8DFD8&
               Caption         =   "Compte"
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
               Left            =   300
               TabIndex        =   141
               Top             =   1800
               Width           =   885
            End
            Begin VB.Label lblCRTMVTRUB_OD 
               BackColor       =   &H00D8DFD8&
               Caption         =   "Rubrique"
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
               Left            =   300
               TabIndex        =   140
               Top             =   600
               Width           =   870
            End
            Begin VB.Label lblCRTMVTCLIP_OD 
               BackColor       =   &H00D8DFD8&
               Caption         =   "Pays"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   300
               TabIndex        =   139
               Top             =   1200
               Width           =   825
            End
         End
         Begin VB.Frame fraYTVACOM0 
            BackColor       =   &H00C0E0FF&
            Caption         =   "fraYTVACOM0"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1185
            Left            =   -67170
            TabIndex        =   128
            Top             =   1125
            Visible         =   0   'False
            Width           =   3420
            Begin VB.TextBox txtTVACOMCOMC 
               Height          =   375
               Left            =   330
               Locked          =   -1  'True
               TabIndex        =   129
               Text            =   "txtTVACOMCOMC"
               Top             =   525
               Width           =   1695
            End
         End
         Begin VB.Frame fraYCRTCPT0 
            BackColor       =   &H00D0F0FF&
            Caption         =   "Compte"
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4335
            Left            =   5820
            TabIndex        =   106
            Top             =   1980
            Visible         =   0   'False
            Width           =   7275
            Begin VB.CommandButton cmdYCRTCPT0_Exclure 
               BackColor       =   &H000000FF&
               Caption         =   "Exclure"
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
               Left            =   345
               Style           =   1  'Graphical
               TabIndex        =   120
               Top             =   3540
               Width           =   1000
            End
            Begin VB.ComboBox cboYCRTCPT0_CRTCPTSTA 
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
               Left            =   1815
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   119
               Top             =   2715
               Width           =   2715
            End
            Begin VB.ComboBox cboYCRTCPT0_CRTCPTRUB 
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
               Left            =   1845
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   118
               Top             =   2265
               Width           =   5070
            End
            Begin VB.CommandButton cmdSAB_Dossier_DB 
               BackColor       =   &H0080C0FF&
               Caption         =   "Extrait de compte"
               Height          =   600
               Left            =   6060
               Style           =   1  'Graphical
               TabIndex        =   117
               ToolTipText     =   "Cliquer ici pour afficher toutes les écritures comptables concernant ce dossier"
               Top             =   2850
               Width           =   1000
            End
            Begin VB.CommandButton cmdYCRTCPT0_Quit 
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
               Left            =   6090
               Style           =   1  'Graphical
               TabIndex        =   116
               Top             =   3585
               Width           =   1000
            End
            Begin VB.CommandButton cmdYCRTCPT0_Update 
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
               Left            =   4560
               Style           =   1  'Graphical
               TabIndex        =   115
               Top             =   3600
               Width           =   1000
            End
            Begin VB.CommandButton cmdYCRTCPT0_Ignore 
               BackColor       =   &H00FF80FF&
               Caption         =   "Ignorer"
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
               Left            =   1815
               Style           =   1  'Graphical
               TabIndex        =   114
               Top             =   3540
               Width           =   1000
            End
            Begin VB.TextBox txtD_COMPTECOM 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1905
               Locked          =   -1  'True
               TabIndex        =   113
               Text            =   "COMPTECOM"
               Top             =   500
               Width           =   2310
            End
            Begin VB.TextBox txtD_COMPTEINT 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1905
               Locked          =   -1  'True
               TabIndex        =   112
               Text            =   "COMPTEINT"
               Top             =   1010
               Width           =   5010
            End
            Begin VB.TextBox txtD_COMPTEOBL 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4770
               Locked          =   -1  'True
               TabIndex        =   111
               Text            =   "COMPTEOBL"
               Top             =   500
               Width           =   960
            End
            Begin VB.TextBox txtD_COMPTEFON 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1905
               Locked          =   -1  'True
               TabIndex        =   110
               Text            =   "COMPTEFON"
               Top             =   1500
               Width           =   405
            End
            Begin VB.TextBox txtD_PLANCOPRO 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6300
               Locked          =   -1  'True
               TabIndex        =   109
               Text            =   "PLANCOPRO"
               Top             =   500
               Width           =   495
            End
            Begin VB.TextBox txtD_COMPTEOUV 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3585
               Locked          =   -1  'True
               TabIndex        =   108
               Text            =   "COMPTEOUV"
               Top             =   1485
               Width           =   1140
            End
            Begin VB.TextBox txtD_COMPTECLO 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   5640
               Locked          =   -1  'True
               TabIndex        =   107
               Text            =   "COMPTECLO"
               Top             =   1455
               Width           =   1215
            End
            Begin VB.Label lblYCRTCPT0_CRTCPTSTA 
               BackColor       =   &H00D0F0FF&
               Caption         =   "code état"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   300
               TabIndex        =   125
               Top             =   2775
               Width           =   1470
            End
            Begin VB.Label lblYCRTCPT0_CRTCPTRUB 
               BackColor       =   &H00D0F0FF&
               Caption         =   "Rubrique"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   285
               TabIndex        =   124
               Top             =   2310
               Width           =   1500
            End
            Begin VB.Label lblD_COMPTECOM 
               BackColor       =   &H00D0F0FF&
               Caption         =   "compte PCI produit"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   180
               TabIndex        =   123
               Top             =   550
               Width           =   1530
            End
            Begin VB.Label lblD_COMPTEINT 
               BackColor       =   &H00D0F0FF&
               Caption         =   "intitulé"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   165
               TabIndex        =   122
               Top             =   1065
               Width           =   1530
            End
            Begin VB.Label lblD_COMPTEFON 
               BackColor       =   &H00D0F0FF&
               Caption         =   "code fonct,Dcre,Dclo"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   180
               TabIndex        =   121
               Top             =   1550
               Width           =   1530
            End
         End
         Begin VB.Frame fraSelect_2 
            BackColor       =   &H00CDEBFF&
            Caption         =   "Filtre des mouvements : Code origine  /   Code état / Code Pays"
            Height          =   735
            Left            =   180
            TabIndex        =   103
            Top             =   2010
            Width           =   4650
            Begin VB.ComboBox cboSelect_CRTMVTCLIP 
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
               Left            =   2565
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   157
               Top             =   270
               Width           =   1905
            End
            Begin VB.ComboBox cboSelect_CRTMVTORIG 
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
               TabIndex        =   105
               Text            =   "cboSelect_App"
               Top             =   300
               Width           =   1125
            End
            Begin VB.ComboBox cboSelect_CRTMVTSTA 
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
               Left            =   1215
               Sorted          =   -1  'True
               TabIndex        =   104
               Text            =   "cboSelect_App"
               Top             =   285
               Width           =   1215
            End
         End
         Begin VB.Frame fraSelect_Log 
            BackColor       =   &H00F0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1305
            Left            =   180
            TabIndex        =   95
            Top             =   450
            Visible         =   0   'False
            Width           =   9270
            Begin VB.ComboBox cboSelect_CRTLOGNAT 
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
               Left            =   1380
               Sorted          =   -1  'True
               TabIndex        =   97
               Text            =   "cboSelect_CRTLOGNAT"
               Top             =   660
               Width           =   5460
            End
            Begin VB.TextBox txtSelect_CRTLOGCPT 
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
               Left            =   1365
               TabIndex        =   96
               Top             =   180
               Width           =   1650
            End
            Begin MSComCtl2.DTPicker txtSelect_Log_AmjMin 
               Height          =   300
               Left            =   7850
               TabIndex        =   98
               Top             =   420
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
               Format          =   97910787
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_Log_AmjMax 
               Height          =   300
               Left            =   7850
               TabIndex        =   99
               Top             =   795
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
               Format          =   97910787
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_Log_Amj 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Période"
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
               Left            =   8050
               TabIndex        =   102
               Top             =   75
               Width           =   795
            End
            Begin VB.Label lblSelect_CRTLOGNAT 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Nature"
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
               TabIndex        =   101
               Top             =   600
               Width           =   795
            End
            Begin VB.Label lblSelect_CRTLOGCPT 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Compte"
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
               TabIndex        =   100
               Top             =   210
               Width           =   795
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
            Left            =   390
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   94
            Text            =   "BDF_CRT.frx":0622
            Top             =   5595
            Visible         =   0   'False
            Width           =   5775
         End
         Begin VB.Frame fraYCRTMVT0 
            BackColor       =   &H00F0FFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7815
            Left            =   -74880
            TabIndex        =   67
            Top             =   525
            Width           =   13095
            Begin VB.CommandButton cmdYCRTMVT0_Mvt_Ann 
               BackColor       =   &H000000FF&
               Caption         =   "Annulation des mouvements identiques : mise à jour du paramétrage + traitement"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2280
               Left            =   11415
               MaskColor       =   &H0080FFFF&
               Style           =   1  'Graphical
               TabIndex        =   80
               Top             =   4860
               Visible         =   0   'False
               Width           =   1300
            End
            Begin VB.TextBox libYCRTMVT0_CRTMVTDOS 
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
               Left            =   7455
               Locked          =   -1  'True
               TabIndex        =   79
               Top             =   5115
               Width           =   3375
            End
            Begin VB.TextBox txtYCRTMVT0_CRTMVTCPT 
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
               Locked          =   -1  'True
               TabIndex        =   78
               Top             =   6315
               Width           =   1695
            End
            Begin VB.TextBox txtYCRTMVT0_CRTMVTCLIN 
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
               Left            =   9210
               Locked          =   -1  'True
               TabIndex        =   77
               Top             =   5685
               Width           =   1635
            End
            Begin VB.ComboBox cboYCRTMVT0_CRTMVTORIG 
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
               Left            =   3150
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   76
               Top             =   7290
               Width           =   1935
            End
            Begin VB.CommandButton cmdYCRTMVT0_Log 
               BackColor       =   &H00FFFFC0&
               Caption         =   "historique des modifications"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   500
               Left            =   11400
               Style           =   1  'Graphical
               TabIndex        =   75
               ToolTipText     =   "Cliquer ici pour afficher toutes les écritures comptables concernant ce dossier"
               Top             =   4830
               Width           =   1300
            End
            Begin VB.CommandButton cmdYCRTMVT0_Question 
               BackColor       =   &H0080C0FF&
               Caption         =   "à revoir"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   500
               Left            =   11400
               Style           =   1  'Graphical
               TabIndex        =   74
               ToolTipText     =   "Cliquer ici pour afficher toutes les écritures comptables concernant ce dossier"
               Top             =   5970
               Width           =   1300
            End
            Begin VB.ComboBox cboYCRTMVT0_CRTMVTSTA 
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
               Left            =   1000
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   73
               Top             =   7290
               Width           =   1935
            End
            Begin VB.CommandButton cmdYCRTMVT0_Ignore 
               BackColor       =   &H00FF80FF&
               Caption         =   "Ignorer"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   500
               Left            =   11400
               Style           =   1  'Graphical
               TabIndex        =   72
               Top             =   5385
               Width           =   1300
            End
            Begin VB.CommandButton cmdYCRTMVT0_Update 
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
               Height          =   500
               Left            =   11400
               Style           =   1  'Graphical
               TabIndex        =   71
               Top             =   6570
               Width           =   1300
            End
            Begin VB.CommandButton cmdYCRTMVT0_Quit 
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
               Height          =   500
               Left            =   11400
               Style           =   1  'Graphical
               TabIndex        =   70
               Top             =   7185
               Width           =   1300
            End
            Begin VB.ComboBox cboYCRTMVT0_CRTMVTCLIP 
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
               Left            =   1000
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   69
               Top             =   5730
               Width           =   3255
            End
            Begin VB.ComboBox cboYCRTMVT0_CRTMVTRUB 
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
               Left            =   1000
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   68
               Top             =   5115
               Width           =   5400
            End
            Begin TabDlg.SSTab SSTab2 
               Height          =   4650
               Left            =   120
               TabIndex        =   81
               Top             =   120
               Width           =   12975
               _ExtentX        =   22886
               _ExtentY        =   8202
               _Version        =   393216
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   520
               TabCaption(0)   =   "Ecriture comptable"
               TabPicture(0)   =   "BDF_CRT.frx":062A
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "fgBIAMVT"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Messages Swift"
               TabPicture(1)   =   "BDF_CRT.frx":0646
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "fgYSWISAB0"
               Tab(1).ControlCount=   1
               Begin MSFlexGridLib.MSFlexGrid fgBIAMVT 
                  Height          =   4095
                  Left            =   120
                  TabIndex        =   82
                  Top             =   360
                  Width           =   12735
                  _ExtentX        =   22463
                  _ExtentY        =   7223
                  _Version        =   393216
                  Cols            =   10
                  FixedCols       =   0
                  RowHeightMin    =   350
                  BackColor       =   14737632
                  ForeColor       =   16384
                  BackColorFixed  =   9470064
                  ForeColorFixed  =   -2147483633
                  BackColorBkg    =   -2147483633
                  AllowUserResizing=   3
                  FormatString    =   $"BDF_CRT.frx":0662
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
               Begin MSFlexGridLib.MSFlexGrid fgYSWISAB0 
                  Height          =   4095
                  Left            =   -74880
                  TabIndex        =   83
                  ToolTipText     =   "cliquer pour afficher le détail du message swift"
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   12735
                  _ExtentX        =   22463
                  _ExtentY        =   7223
                  _Version        =   393216
                  Rows            =   1
                  Cols            =   13
                  FixedCols       =   0
                  RowHeightMin    =   300
                  BackColor       =   16777215
                  ForeColor       =   12582912
                  BackColorFixed  =   10526720
                  ForeColorFixed  =   16777215
                  BackColorSel    =   12648384
                  BackColorBkg    =   15794160
                  AllowBigSelection=   0   'False
                  FocusRect       =   2
                  HighLight       =   0
                  GridLinesFixed  =   1
                  AllowUserResizing=   3
                  FormatString    =   $"BDF_CRT.frx":0757
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
            Begin VB.Label lblYCRTMVT0_CRTMVTDOS 
               BackColor       =   &H00F0FFFF&
               Caption         =   "dossier"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   6600
               TabIndex        =   93
               Top             =   5115
               Width           =   735
            End
            Begin VB.Label lblYCRTMVT0_CRTMVTMTE 
               BackColor       =   &H00F0FFFF&
               Caption         =   "CV EUR"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   120
               TabIndex        =   92
               Top             =   6765
               Width           =   840
            End
            Begin VB.Label libYCRTMVT0_CRTMVTCPT 
               BackColor       =   &H00F0FFFF&
               Caption         =   "compte"
               Height          =   270
               Left            =   2805
               TabIndex        =   91
               Top             =   6375
               Width           =   3660
            End
            Begin VB.Label libYCRTMVT0_CRTMVTCLIN 
               BackColor       =   &H00F0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "client"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1350
               Left            =   6600
               TabIndex        =   90
               Top             =   6225
               Width           =   4335
            End
            Begin VB.Label lblYCRTMVT0_CRTMVTCLIN 
               BackColor       =   &H00F0FFFF&
               Caption         =   "client"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   6600
               TabIndex        =   89
               Top             =   5730
               Width           =   2415
            End
            Begin VB.Label lblYCRTMVT0_CRTMVTCPT 
               BackColor       =   &H00F0FFFF&
               Caption         =   "compte"
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
               Left            =   120
               TabIndex        =   88
               Top             =   6345
               Width           =   885
            End
            Begin VB.Label libYCRTMVT0_CRTMVTMTE 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Montant"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1005
               TabIndex        =   87
               Top             =   6780
               Width           =   1650
            End
            Begin VB.Label lblYCRTMVT0_CRTMVTSTA 
               BackColor       =   &H00F0FFFF&
               Caption         =   "code état"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   86
               Top             =   7365
               Width           =   840
            End
            Begin VB.Label lblYCRTMVT0_CRTMVTCLIP 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Pays"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   120
               TabIndex        =   85
               Top             =   5790
               Width           =   825
            End
            Begin VB.Label lblYCRTMVT0_CRTMVTRUB 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Rubrique"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   120
               TabIndex        =   84
               Top             =   5115
               Width           =   975
            End
         End
         Begin VB.Frame fraSwift 
            BackColor       =   &H00C0E0FF&
            Height          =   7320
            Left            =   -74865
            TabIndex        =   63
            Top             =   435
            Visible         =   0   'False
            Width           =   7185
            Begin VB.CheckBox chkSIDE_DB_Show 
               BackColor       =   &H00C0FFFF&
               Caption         =   "afficher le message et l'historique du traitement SAA"
               Height          =   255
               Left            =   60
               TabIndex        =   64
               Top             =   600
               Width           =   6945
            End
            Begin MSFlexGridLib.MSFlexGrid fgSwift 
               Height          =   6285
               Left            =   60
               TabIndex        =   65
               Top             =   930
               Width           =   7000
               _ExtentX        =   12356
               _ExtentY        =   11086
               _Version        =   393216
               Cols            =   4
               FixedCols       =   0
               RowHeightMin    =   350
               BackColor       =   16777215
               ForeColor       =   12582912
               BackColorFixed  =   16777168
               ForeColorFixed  =   16711680
               BackColorBkg    =   16777215
               GridColor       =   12632064
               GridColorFixed  =   12632064
               AllowUserResizing=   3
               FormatString    =   $"BDF_CRT.frx":084C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label libSWIFT_SWISABSWID 
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   75
               TabIndex        =   66
               Top             =   210
               Width           =   6960
            End
         End
         Begin VB.Frame fraYCRTLOG0 
            BackColor       =   &H00D0F0FF&
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7000
            Left            =   -69690
            TabIndex        =   50
            Top             =   1965
            Width           =   7275
            Begin VB.TextBox txtCRTLOGNAT 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   55
               Top             =   510
               Width           =   5700
            End
            Begin VB.TextBox txtCRTLOGUUSR 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1065
               Locked          =   -1  'True
               TabIndex        =   54
               Top             =   6495
               Width           =   1700
            End
            Begin VB.TextBox txtCRTLOGUAMJ 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3195
               Locked          =   -1  'True
               TabIndex        =   53
               Top             =   6480
               Width           =   3765
            End
            Begin VB.TextBox txtCRTLOGCPT 
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1455
               Locked          =   -1  'True
               TabIndex        =   52
               Top             =   930
               Width           =   1700
            End
            Begin VB.TextBox xttCRTLOGPIE 
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
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   51
               Top             =   1365
               Width           =   1700
            End
            Begin MSFlexGridLib.MSFlexGrid fgYCRTLOG0 
               Height          =   3690
               Left            =   195
               TabIndex        =   56
               Top             =   2700
               Width           =   6780
               _ExtentX        =   11959
               _ExtentY        =   6509
               _Version        =   393216
               Cols            =   3
               BackColor       =   16449525
               BackColorFixed  =   11394815
               ForeColorFixed  =   16448
               BackColorBkg    =   16449525
               GridColor       =   10526720
               GridColorFixed  =   10526720
               FormatString    =   "<Champ                   |<Ancienne valeur                       |<Nouvelle valeur                             "
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
            Begin VB.Label lblCRTLOGUUSR 
               BackColor       =   &H00D0F0FF&
               Caption         =   "Màj"
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
               Left            =   210
               TabIndex        =   62
               Top             =   6555
               Width           =   675
            End
            Begin VB.Label lblCRTLOGID 
               BackColor       =   &H00D0F0FF&
               Caption         =   "intitulé"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   255
               TabIndex        =   61
               Top             =   495
               Width           =   825
            End
            Begin VB.Label lblCRTLOGCPT 
               BackColor       =   &H00D0F0FF&
               Caption         =   "Compte"
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
               Left            =   270
               TabIndex        =   60
               Top             =   1035
               Width           =   675
            End
            Begin VB.Label libCRTLOGCPT 
               BackColor       =   &H00D0F0FF&
               Caption         =   "Compte"
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
               Left            =   3300
               TabIndex        =   59
               Top             =   990
               Width           =   3795
            End
            Begin VB.Label lblCRTLOGTXT 
               BackColor       =   &H00F0FAFF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Left            =   210
               TabIndex        =   58
               Top             =   1755
               Visible         =   0   'False
               Width           =   6615
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblCRTLOGPIE 
               BackColor       =   &H00D0F0FF&
               Caption         =   "Mouvement"
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
               Left            =   255
               TabIndex        =   57
               Top             =   1410
               Width           =   1140
            End
         End
         Begin RichTextLib.RichTextBox txtRTF 
            Height          =   5610
            Left            =   180
            TabIndex        =   126
            TabStop         =   0   'False
            Top             =   3105
            Visible         =   0   'False
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   9895
            _Version        =   393217
            BackColor       =   14737632
            Enabled         =   -1  'True
            HideSelection   =   0   'False
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"BDF_CRT.frx":08DB
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
         Begin MSFlexGridLib.MSFlexGrid fgLog 
            Height          =   7710
            Left            =   -74715
            TabIndex        =   127
            Top             =   750
            Visible         =   0   'False
            Width           =   13140
            _ExtentX        =   23178
            _ExtentY        =   13600
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   400
            BackColor       =   16777215
            ForeColor       =   16711680
            BackColorFixed  =   11394815
            ForeColorFixed  =   16448
            BackColorBkg    =   16777215
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   $"BDF_CRT.frx":095B
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
      Picture         =   "BDF_CRT.frx":0AA3
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
Attribute VB_Name = "frmBDF_CRT"
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
Dim arrHab(19) As Boolean
Dim blnAuto As Boolean, blnError As Boolean
Dim cmdSelect_SQL_K As String, mSQL_Exe As String, mSQL_Where As String

Dim rsSabX As New ADODB.Recordset

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean
Dim mSelect_Id As String

Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean
Dim mDetail_Id As String
Dim fgDetail_Left As Long, fgDetail_Width As Long
Dim fgBIAMVT_FormatString As String, fgBIAMVT_K As Integer
Dim fgBIAMVT_RowDisplay As Integer, fgBIAMVT_RowClick As Integer, fgBIAMVT_ColClick As Integer
Dim fgBIAMVT_ColorClick As Long, fgBIAMVT_ColorDisplay As Long
Dim fgBIAMVT_Sort1 As Integer, fgBIAMVT_Sort2 As Integer
Dim fgBIAMVT_SortAD As Integer, fgBIAMVT_Sort1_Old As Integer
Dim fgBIAMVT_arrIndex As Integer
Dim blnfgBIAMVT_DisplayLine As Boolean

Dim fgYSWISAB0_FormatString As String, fgYSWISAB0_K As Integer
Dim fgYSWISAB0_RowDisplay As Integer, fgYSWISAB0_RowClick As Integer, fgYSWISAB0_ColClick As Integer
Dim fgYSWISAB0_ColorClick As Long, fgYSWISAB0_ColorDisplay As Long
Dim fgYSWISAB0_Sort1 As Integer, fgYSWISAB0_Sort2 As Integer
Dim fgYSWISAB0_SortAD As Integer, fgYSWISAB0_Sort1_Old As Integer
Dim fgYSWISAB0_arrIndex As Integer
Dim blnfgYSWISAB0_DisplayLine As Boolean

Dim xYSWISAB0 As typeYSWISAB0, oldYSWISAB0 As typeYSWISAB0
Dim cnSIDE_DB As New ADODB.Connection, rsSIDE_DB As New ADODB.Recordset
Dim blnSIDE_DB As Boolean
Dim fgSwift_FormatString As String
Dim xrText As typerText

'______________________________________________________________________

Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long

Dim Old_YBIATAB0 As typeYBIATAB0, New_YBIATAB0 As typeYBIATAB0

Dim HeightOfLine As Long, LinesOfText As Long

Dim txtRTF_prtForeColor_Header As Long

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim mXls1_Row As Long, mXls1_Cols As Integer, mXls1_File As Integer
Dim mXls2_Cols As Integer, mXls2_Row As Integer

Dim newYCRTCPT0 As typeYCRTCPT0, oldYCRTCPT0 As typeYCRTCPT0, xYCRTCPT0 As typeYCRTCPT0
Dim newYCRTMVT0 As typeYCRTMVT0, oldYCRTMVT0 As typeYCRTMVT0, xYCRTMVT0 As typeYCRTMVT0
Dim newYCRTLOG0 As typeYCRTLOG0, oldYCRTLOG0 As typeYCRTLOG0, xYCRTLOG0 As typeYCRTLOG0

Dim mCRTMVTCPT_Display As String

Dim blnfrmSAB_Dossier_DB As Boolean
Dim mAmjMin_Exercice As String, mAmjMax_Exercice As String, mExercice As String, mAmjMin_CRTMVTDTR As String

Dim xYBIAMVTH As typeYBIAMVT0, oldYBIAMVTH  As typeYBIAMVT0
Dim arrLogNat_code() As String, arrLogNat_Lib() As String, arrLogNat_Nb As Integer
Dim arrCRT_RUB_I_code() As String, arrCRT_RUB_I_Nb As Integer
Dim arrCRT_Mvt_Ann() As typeYCRTMVT0, arrCRT_Mvt_Ann_Nb As Integer
Dim fgYCRTLOG0_FormatString As String
Dim fgLog0_FormatString As String
Dim xZCLIENA0 As typeZCLIENA0, xZADRESS0 As typeZADRESS0
Dim mPays_Exclus As String, mClients_Exclus As String

Dim mCRTMVTTAUX_OD As Double, blnCRTMVTTAUX_OD As Boolean
Dim mCRT_File As String, mCRT_File_Local As String
Dim arrCRT_Devise() As String, arrCRT_Devise_Nb As Integer
Dim arrNomenclature(3) As String, xmlDomain As String
Dim arrPays() As typePays, arrPays_Nb As Integer

Dim mCRT_File_Id As String, kNomenclature_xml As Integer, kNomenclature_File_No As Integer
Dim mCRT_àRevoir As Long
Public Sub cmdPrint_Excel(lTxt As String)
On Error GoTo Error_Handler
Dim xSQL As String
Dim X As String, wFilex As String
Dim blnCALCS As Boolean

On Error GoTo Error_Handler
'===================================================================================
'______________________________________________'
'X = paramServer("\\CPT_Archive\")
wAMJMin = DSys

blnCALCS = False
'If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True

'If X = "" Then

X = "C:\Temp\"
If mId$(X, Len(X), 1) <> "\" Then X = X & "\"

If cmdSelect_SQL_K = "CRT.xml" Then
    wFile = X & mCRT_File_Id & ".xlsx"
Else
    mXls1_File = mXls1_File + 1
    
    wFile = X & Trim("BDF_CRT " & lTxt & "  - " & DSYS_Time & mXls1_File & ".xlsx")
End If
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "BDF_CRT : nom du fichier d'exportation", wFile)
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
    .Title = "BDF_CRT"
    .Subject = "BDF_CRT"
End With

'__________________________________________________________________________________

'appExcel.Worksheets.Add

Set wsExcel = wbExcel.Sheets(1): wsExcel.Name = "BDF_CRT " & lTxt

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

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14BDF_CRT, arrêté au " & dateImp10(wAMJMin) _
                                 & vbCr

wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

mXls1_Row = 1

Select Case SSTab1.Tab
    Case 0:
        Select Case cmdSelect_SQL_K
            Case "1":   appExcel.Worksheets.Add
            
                        Set wsExcel = wbExcel.Sheets(1)
                        Call cmdPrint_Excel_YCRTCPT0(" ", "à déclarer")
                        Set wsExcel = wbExcel.Sheets(2)
                        Call cmdPrint_Excel_YCRTCPT0("I", "à ignorer")
                        Set wsExcel = wbExcel.Sheets(3)
                        Call cmdPrint_Excel_YCRTCPT0("E", "à exclure")
                        Set wsExcel = wbExcel.Sheets(4)
                        Call cmdPrint_Excel_YCRTCPT0_Rubriques
            Case "2", "2c", "Annulation$":
                        Set wsExcel = wbExcel.Sheets(1)
                        Call cmdPrint_Excel_YCRTMVT0
            Case "CRT.xls":
                        kNomenclature_xml = 0: kNomenclature_File_No = 0
                        Set wsExcel = wbExcel.Sheets(1)
                        Call cmdPrint_Excel_Déclaration
            Case "CRT.xml":
                        kNomenclature_xml = 0: kNomenclature_File_No = 0
                        Set wsExcel = wbExcel.Sheets(1)
                        Call cmdPrint_Excel_Déclaration
                        Call sqlYCRTLOG0_Insert_Transaction(newYCRTLOG0)

        End Select
    Case 1
        cmdPrint_Excel_YBIATAB0
        

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

Public Sub cmdPrint_Excel_YBIATAB0()
Dim xSQL As String, X As String, K As Long, mBIATABID As String
On Error GoTo Error_Handler

'===================================================================================

wsExcel.Name = "Paramétrage"

With wsExcel.Cells
    .HorizontalAlignment = Excel.xlHAlignLeft
    .Font.Size = 10
    .Font.Name = "Courier New"
End With
wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 75
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14BDF_CRT : paramétrage" _
                                & "  (édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$E1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


wsExcel.Columns(1).ColumnWidth = 15: wsExcel.Cells(1, 1) = "Table "
wsExcel.Columns(2).ColumnWidth = 15: wsExcel.Cells(1, 2) = "Code"
wsExcel.Columns(3).ColumnWidth = 15: wsExcel.Cells(1, 3) = "Valeur"
wsExcel.Columns(4).ColumnWidth = 120: wsExcel.Cells(1, 4) = "Libellé"

For K = 1 To 4
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = vbWhite
Next

mXls1_Row = mXls1_Row + 1
lstParam_Load_Pays_Exclus ("Excel")

mXls1_Row = mXls1_Row + 1
lstParam_Load_Clients_Exclus ("Excel")


xSQL = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID like 'CRT%'" _
     & " order by BIATABID , BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    If mBIATABID <> rsSab("BIATABID") Then
        mBIATABID = rsSab("BIATABID")
        mXls1_Row = mXls1_Row + 1
    End If
    mXls1_Row = mXls1_Row + 1
    wsExcel.Cells(mXls1_Row, 1) = rsSab("BIATABID")
    wsExcel.Cells(mXls1_Row, 2) = rsSab("BIATABK1")
    wsExcel.Cells(mXls1_Row, 3) = rsSab("BIATABK2")
    wsExcel.Cells(mXls1_Row, 4) = Trim(mId$(rsSab("BIATABTXT"), 1, 99))
    Select Case Trim(rsSab("BIATABID"))
        Case "CRT_Mvt=Ann"
                        xSQL = "select COMPTEINT from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '" & mId$(rsSab("BIATABTXT"), 1, 20) & "'"
                        Set rsSabX = cnsab.Execute(xSQL)
                        If Not rsSabX.EOF Then wsExcel.Cells(mXls1_Row, 4) = Trim(mId$(rsSab("BIATABTXT"), 1, 34)) & "      = " & rsSabX("COMPTEINT")
                        For K = 1 To 1: wsExcel.Cells(mXls1_Row, K).Font.Color = vbRed: Next K
        Case "CRT_Rub_I"
                        xSQL = "select BIATABTXT from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_Rubrique'" _
                             & " and BIATABK1 = '" & rsSab("BIATABK1") & "'"
                        Set rsSabX = cnsab.Execute(xSQL)
                        If Not rsSabX.EOF Then wsExcel.Cells(mXls1_Row, 4) = Trim(mId$(rsSabX("BIATABTXT"), 1, 99))
                        For K = 1 To 1: wsExcel.Cells(mXls1_Row, K).Font.Color = vbMagenta: Next K
                        
        
        Case "CRT_AAAA": For K = 1 To 4: wsExcel.Cells(mXls1_Row, K).Font.Color = vbBlue: Next K

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
Public Sub cmdPrint_Excel_YCRTMVT0()
Dim xSQL As String, X As String, K As Long, xWhere As String
Dim blnCRTMVTSTA_Color As Boolean
On Error GoTo Error_Handler

'===================================================================================

wsExcel.Name = "Mouvements"


'________________________________________________________________________________

xWhere = ""
Call DTPicker_Control(txtSelect_AmjMin, wAMJMin)
Call DTPicker_Control(txtSelect_AmjMax, WAMJMax)
xWhere = " and CRTMVTDTR >= " & wAMJMin & " and CRTMVTDTR <= " & WAMJMax

blnCRTMVTSTA_Color = True


X = Trim(mId$(cboSelect_CRTMVTCLIP, 1, 2))
If X <> "" Then xWhere = xWhere & " and CRTMVTCLIP = '" & X & "'"

X = Trim(mId$(cboSelect_CRTMVTORIG, 1, 1))
If X <> "" Then xWhere = xWhere & " and CRTMVTORIG = '" & X & "'"

X = Trim(mId$(cboSelect_CRTCPTRUB, 1, 5))
If X <> "" Then xWhere = xWhere & " and CRTMVTRUB ='" & X & "'"

If cmdSelect_SQL_K = "2c" Then
    xWhere = xWhere & " and CRTMVTCPT = '" & mCRTMVTCPT_Display & "'"
Else
    X = Trim(txtSelect_CRTCPTCPT)
    If X <> "" Then xWhere = xWhere & " and CRTMVTCPT like '" & X & "%'"

End If

            
X = Trim(mId$(cboSelect_CRTMVTSTA, 1, 1))

Select Case X
    Case "":
                wsExcel.Name = "à déclarer"
                Call cmdPrint_Excel_YCRTMVT0_Detail(xWhere & " and CRTMVTSTA = ' '", blnCRTMVTSTA_Color)
    Case "#":
                wsExcel.Name = "à déclarer"
                Call cmdPrint_Excel_YCRTMVT0_Detail(xWhere & " and CRTMVTSTA = ' '", blnCRTMVTSTA_Color)
                
                Set wsExcel = wbExcel.Sheets(2)
                wsExcel.Name = "à revoir"
                Call cmdPrint_Excel_YCRTMVT0_Detail(xWhere & " and CRTMVTSTA = '?'", blnCRTMVTSTA_Color)
                
                Set wsExcel = wbExcel.Sheets(3)
                wsExcel.Name = "annulés"
                Call cmdPrint_Excel_YCRTMVT0_Detail(xWhere & " and CRTMVTSTA = 'A'", blnCRTMVTSTA_Color)
    Case "*"
                appExcel.Worksheets.Add
                
                Set wsExcel = wbExcel.Sheets(1)
                wsExcel.Name = "à déclarer"
                Call cmdPrint_Excel_YCRTMVT0_Detail(xWhere & " and CRTMVTSTA = ' '", blnCRTMVTSTA_Color)
                
                Set wsExcel = wbExcel.Sheets(2)
                wsExcel.Name = "à revoir"
                Call cmdPrint_Excel_YCRTMVT0_Detail(xWhere & " and CRTMVTSTA = '?'", blnCRTMVTSTA_Color)
                
                Set wsExcel = wbExcel.Sheets(3)
                wsExcel.Name = "annulés"
                Call cmdPrint_Excel_YCRTMVT0_Detail(xWhere & " and CRTMVTSTA = 'A'", blnCRTMVTSTA_Color)
                 
                Set wsExcel = wbExcel.Sheets(4)
                wsExcel.Name = "ignorés"
                Call cmdPrint_Excel_YCRTMVT0_Detail(xWhere & " and CRTMVTSTA = 'I'", blnCRTMVTSTA_Color)
   Case "I":
                blnCRTMVTSTA_Color = False
                wsExcel.Name = "à ignorer"
                Call cmdPrint_Excel_YCRTMVT0_Detail(xWhere & " and CRTMVTSTA = '" & X & "'", blnCRTMVTSTA_Color)
   Case "?":
                blnCRTMVTSTA_Color = False
                wsExcel.Name = "à revoir"
                Call cmdPrint_Excel_YCRTMVT0_Detail(xWhere & " and CRTMVTSTA = '" & X & "'", blnCRTMVTSTA_Color)
    Case Else:
                wsExcel.Name = X
               Call cmdPrint_Excel_YCRTMVT0_Detail(xWhere & " and CRTMVTSTA = '" & X & "'", blnCRTMVTSTA_Color)
End Select




'======================================================================================================

Exit_sub:
'__________________________________________________________________________________


'_____________________________
Exit Sub

Error_Handler:
    If Not blnAuto Then Call MsgBox(Error, vbCritical, Me.Name)

End Sub




Public Sub cmdPrint_Excel_YCRTCPT0(lFct As String, lName As String)
Dim xSQL As String, X As String, K As Long
Dim mXls1_Cols As Integer, wColor As Long
Dim arrRub_Col() As Integer

On Error GoTo Error_Handler

Call lstErr_AddItem(lstErr, cmdContext, "> cmdPrint_Excel_YCRTCPT0 ........"): DoEvents

'===================================================================================
ReDim arrRub_Col(arrCRT_PCEC_Rub_Nb)
For K = 1 To arrCRT_PCEC_Rub_Nb
    X = arrCRT_PCEC_Rub(K).Code
    If X <> arrCRT_Rub(arrCRT_Rub_K).Code Then
        For arrCRT_Rub_K = 0 To arrCRT_Rub_Nb
            If X = arrCRT_Rub(arrCRT_Rub_K).Code Then Exit For
        Next arrCRT_Rub_K
    End If
    arrRub_Col(K) = arrCRT_Rub_K
Next K
'===================================================================================

wsExcel.Name = lName

With wsExcel.Cells
    .HorizontalAlignment = Excel.xlHAlignLeft
    .Font.Size = 9
    .Font.Name = "Courier New"
End With
wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 65
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14BDF_CRT : paramétrage des comptes" _
                                & "  (édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$E1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True

mXls1_Cols = 3 + arrCRT_Rub_Nb
mXls1_Row = 1

wsExcel.Rows(1).Orientation = xlVertical
wsExcel.Rows(1).RowHeight = 70
wsExcel.Rows(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Rows(1).VerticalAlignment = Excel.xlVAlignTop

wsExcel.Columns(1).ColumnWidth = 3: wsExcel.Cells(1, 1) = "Sta"
wsExcel.Cells(1, 1).Orientation = xlHorizontal: wsExcel.Cells(1, 1).VerticalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 15: wsExcel.Cells(1, 2) = "Compte"
wsExcel.Cells(1, 2).Orientation = xlHorizontal: wsExcel.Cells(1, 2).VerticalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(3).ColumnWidth = 35: wsExcel.Cells(1, 3) = "Libellé"
wsExcel.Cells(1, 3).Orientation = xlHorizontal: wsExcel.Cells(1, 3).VerticalAlignment = Excel.xlHAlignCenter

For K = 1 To arrCRT_Rub_Nb
    wsExcel.Cells(1, 3 + K) = arrCRT_Rub(K).Code: wsExcel.Columns(3 + K).ColumnWidth = 2
    wsExcel.Columns(3 + K).HorizontalAlignment = Excel.xlHAlignCenter
Next K
For K = 1 To mXls1_Cols
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = vbWhite
Next

Call lstErr_AddItem(lstErr, cmdContext, "> cmdPrint_Excel_YCRTCPT0 SQL"): DoEvents
mSQL_Where = ""

X = Trim(txtSelect_CRTCPTCPT)
If X <> "" Then mSQL_Where = " and CRTCPTCPT like '" & X & "%'"
X = Trim(txtSelect_COMPTEOBL)
If X <> "" Then mSQL_Where = mSQL_Where & " and COMPTEOBL like '" & X & "%'"
X = Trim(mId$(cboSelect_CRTCPTRUB, 1, 5))
If X <> "" Then mSQL_Where = mSQL_Where & " and CRTCPTRUB ='" & X & "'"

 
mSQL_Exe = "select * from " & paramIBM_Library_SABSPE & ".YCRTCPT0 , " & paramIBM_Library_SABSPE & ".YBIACPT0" _
     & " where CRTCPTCPT = COMPTECOM" & mSQL_Where _
     & " order by CRTCPTCPT"

Select Case lFct
    Case "E": X = Replace(mSQL_Exe, "order by", " and CRTCPTSTA = 'E' order by")
    Case "I": X = Replace(mSQL_Exe, "order by", " and CRTCPTSTA = 'I' order by")
    Case Else: X = Replace(mSQL_Exe, "order by", " and CRTCPTSTA not in ( 'I' , 'E') order by")
End Select


Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF

    mXls1_Row = mXls1_Row + 1
    
    X = rsSab("CRTCPTSTA")
    Select Case X
        Case "*": wColor = RGB(16, 96, 16)
        Case "E": wColor = vbRed
        Case "M": wColor = vbBlue
        Case Else: wColor = RGB(128, 128, 128)
        
    End Select
    wsExcel.Cells(mXls1_Row, 1) = X: wsExcel.Cells(mXls1_Row, 1).Font.Color = wColor
    
    X = rsSab("CRTCPTCPT")
    If mXls1_Row Mod 10 Then Call lstErr_ChangeLastItem(lstErr, cmdContext, "> YCRTCPT0 : " & X): DoEvents

    wsExcel.Cells(mXls1_Row, 2) = X: wsExcel.Cells(mXls1_Row, 2).Font.Color = wColor
    wsExcel.Cells(mXls1_Row, 3) = rsSab("COMPTEINT"): wsExcel.Cells(mXls1_Row, 3).Font.Color = wColor
    'wsExcel.Cells(mXls1_Row, 4) = Trim(mId$(rsSab("BIATABTXT"), 1, 99))

    
    For K = arrCRT_PCEC_Rub_Nb0 To arrCRT_PCEC_Rub_Nb
        If mId$(X, 1, arrCRT_PCEC_Rub(K).PCEC_Len) = arrCRT_PCEC_Rub(K).PCEC Then
            wsExcel.Cells(mXls1_Row, 3 + arrRub_Col(K)).Interior.Color = RGB(255, 204, 0)
        End If
        
    Next K
    
    X = Trim(rsSab("CRTCPTRUB"))
    If X <> "" Then
        For arrCRT_Rub_K = 1 To arrCRT_Rub_Nb
            If X = arrCRT_Rub(arrCRT_Rub_K).Code Then
                wsExcel.Cells(mXls1_Row, 3 + arrCRT_Rub_K).Interior.Color = mColor_G9
                wsExcel.Cells(mXls1_Row, 3 + arrCRT_Rub_K) = rsSab("CRTCPTSTA")
                wsExcel.Cells(mXls1_Row, 3 + arrCRT_Rub_K).Font.Bold = True
            End If
            
        Next arrCRT_Rub_K
    End If
    
    rsSab.MoveNext
Loop





'__________________________________________________________________________________


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:

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

'param_Init_Rubrique

fraSelect_2.Visible = False
Set fraSelect_2.Container = fraSelect_Options
fraSelect_2.Top = 30
fraSelect_2.Left = 2820


fraSelect_Log.Visible = False
Set fraSelect_Log.Container = fraSelect_Options
fraSelect_Log.Top = fraSelect_Options.Top
fraSelect_Log.Left = fraSelect_Options.Left
Call DTPicker_Set(txtSelect_Log_AmjMax, DSys)
Call DTPicker_Set(txtSelect_Log_AmjMin, DSys)


fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False

fgDetail_FormatString = fgDetail.FormatString
fgDetail.Enabled = True
fgDetail.Visible = False
fgDetail.Top = fgSelect.Top
fgDetail.Left = fgSelect.Left + fgSelect.Width - fgDetail.Width - 200

fgDetail_Left = fgDetail.Left
fgDetail_Width = fgDetail.Width

fraSelect_Options.Visible = True

fraYCRTCPT0.Visible = False
Set fraYCRTCPT0.Container = fraSelect
fraYCRTCPT0.Top = fgSelect.Top + 400
fraYCRTCPT0.Left = fgSelect.Left + fgSelect.Width - fraYCRTCPT0.Width

fraYCRTMVT0.Visible = False
Set fraYCRTMVT0.Container = fraSelect
fraYCRTMVT0.Top = fgSelect.Top
fraYCRTMVT0.Left = fgSelect.Left

fraSwift.Visible = False
Set fraSwift.Container = fraYCRTMVT0
fraSwift.Top = fgYSWISAB0.Top
fraSwift.Left = fraYCRTMVT0.Left + fraYCRTMVT0.Width - fraSwift.Width

fraYTVACOM0.Visible = False
Set fraYTVACOM0.Container = fraYCRTMVT0
fraYTVACOM0.Top = SSTab2.Top + 400
fraYTVACOM0.Left = fraYCRTMVT0.Left + fraYCRTMVT0.Width - fraYTVACOM0.Width



fgBIAMVT_FormatString = fgBIAMVT.FormatString
fgYSWISAB0_FormatString = fgYSWISAB0.FormatString
fgSwift_FormatString = fgSwift.FormatString

fraYCRTLOG0.Visible = False
Set fraYCRTLOG0.Container = fraSelect
fraYCRTLOG0.Top = fgSelect.Top + 400
fraYCRTLOG0.Left = fgSelect.Left + fgSelect.Width - fraYCRTLOG0.Width
lblCRTLOGID.ForeColor = vbMagenta
fgYCRTLOG0_FormatString = fgYCRTLOG0.FormatString

fgLog0_FormatString = fgLog.FormatString
Set fgLog.Container = fraSelect
fgLog.Top = fgSelect.Top
fgLog.Left = fgSelect.Left


fraParam_CRT_Mvt_Ann.Visible = False
fraParam_CRT_Mvt_Ann.Top = 1900
fraParam_CRT_Mvt_Ann.Left = 120

cmdYCRTMVT0_Mvt_Ann.Top = 4845
cmdYCRTMVT0_Mvt_Ann.Left = cmdYCRTMVT0_Quit.Left


fraOD.Visible = False
Set fraOD.Container = fraSelect
fraOD.Top = fraSelect.Top + 1150
fraOD.Left = fraSelect.Left + fraSelect.Width - fraOD.Width
optCRTMVTMTD_OD_Db.ForeColor = vbRed
optCRTMVTMTD_OD_Cr.ForeColor = vbBlue

Call rsYCRTMVT0_Init(zYCRTMVT0_OD)
zYCRTMVT0_OD.CRTMVTETA = currentSAB_ETA
zYCRTMVT0_OD.CRTMVTPLA = currentSAB_PLA
zYCRTMVT0_OD.CRTMVTORIG = "+"
'___________________________________________________________________________
xSQL = "select BIATABTXT from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_AAAA'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    mExercice = mId$(rsSab("BIATABTXT"), 1, 4)
    mAmjMin_Exercice = mExercice & "0101"
    mAmjMax_Exercice = mExercice & "1231"
    Call DTPicker_Set(txtSelect_AmjMax, mAmjMax_Exercice)
    Call DTPicker_Set(txtSelect_AmjMin, mAmjMin_Exercice)
    Call DTPicker_Set(txtCRTMVTDTR_OD, mAmjMax_Exercice)
    zYCRTMVT0_OD.CRTMVTDTR = mAmjMax_Exercice
Else
    Call MsgBox("Enregistrement CRT_AAAA manquant", vbCritical, "BDF_CRT")
    Unload Me
End If

lstParam_Id.AddItem "CRT_AAAA"
lstParam_Id.AddItem "CRT_Devise"
lstParam_Id.AddItem "CRT_Rubrique"
lstParam_Id.AddItem "CRT_Rub_I"
lstParam_Id.AddItem "CRT_LogNat"
lstParam_Id.AddItem "CRT_Mvt=Ann"
lstParam_Id.AddItem "SAB_Pays_Exclus"
lstParam_Id.AddItem "SAB_Clients_Exclus"
'___________________________________________________________________________

Call lstParam_Load_Pays_Exclus("Init")
Call lstParam_Load_Clients_Exclus("Init")
'___________________________________________________________________________
xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_Rub_I'"
Set rsSab = cnsab.Execute(xSQL)
K = rsSab(0) + 1
ReDim arrCRT_RUB_I_code(K)
arrCRT_RUB_I_Nb = 0
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_Rub_I' order by BIATABK1"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrCRT_RUB_I_Nb = arrCRT_RUB_I_Nb + 1
    arrCRT_RUB_I_code(arrCRT_RUB_I_Nb) = Trim(rsSab("BIATABK1"))
    rsSab.MoveNext
Loop

'___________________________________________________________________________
xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_Devise'"
Set rsSab = cnsab.Execute(xSQL)
K = rsSab(0) + 1
ReDim arrCRT_Devise(K)
arrCRT_Devise_Nb = 2
arrCRT_Devise(1) = "EUR"
arrCRT_Devise(2) = "USD"
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_Devise' order by BIATABK1"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    X = Trim(rsSab("BIATABK1"))
    Select Case X
        Case "EUR", "USD"
        Case Else
            arrCRT_Devise_Nb = arrCRT_Devise_Nb + 1
            arrCRT_Devise(arrCRT_Devise_Nb) = X
    End Select
    rsSab.MoveNext
Loop

'___________________________________________________________________________
xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_Mvt=Ann'"
Set rsSab = cnsab.Execute(xSQL)
K = rsSab(0) + 1
ReDim arrCRT_Mvt_Ann(K)
arrCRT_Mvt_Ann_Nb = 0
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_Mvt=Ann' order by BIATABTXT"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrCRT_Mvt_Ann_Nb = arrCRT_Mvt_Ann_Nb + 1
    X = rsSab("BIATABTXT")
    arrCRT_Mvt_Ann(arrCRT_Mvt_Ann_Nb).CRTMVTCPT = mId$(X, 1, 20)
    arrCRT_Mvt_Ann(arrCRT_Mvt_Ann_Nb).CRTMVTSER = mId$(X, 22, 2)
    arrCRT_Mvt_Ann(arrCRT_Mvt_Ann_Nb).CRTMVTSSE = mId$(X, 25, 2)
    arrCRT_Mvt_Ann(arrCRT_Mvt_Ann_Nb).CRTMVTOPE = mId$(X, 28, 3)
    arrCRT_Mvt_Ann(arrCRT_Mvt_Ann_Nb).CRTMVTEVE = mId$(X, 32, 3)
    
    rsSab.MoveNext
Loop
'___________________________________________________________________________
xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_LogNat'"
Set rsSab = cnsab.Execute(xSQL)
K = rsSab(0) + 1
ReDim arrLogNat_code(K), arrLogNat_Lib(K)
arrLogNat_Nb = 0
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_LogNat' order by BIATABK1"
Set rsSab = cnsab.Execute(xSQL)
cboSelect_CRTLOGNAT.Clear
cboSelect_CRTLOGNAT.AddItem ""
Do While Not rsSab.EOF
    arrLogNat_Nb = arrLogNat_Nb + 1
    arrLogNat_code(arrLogNat_Nb) = Trim(rsSab("BIATABK1"))
    arrLogNat_Lib(arrLogNat_Nb) = Trim(rsSab("BIATABTXT"))
    cboSelect_CRTLOGNAT.AddItem arrLogNat_code(arrLogNat_Nb) & "-" & arrLogNat_Lib(arrLogNat_Nb)
    rsSab.MoveNext
Loop
'___________________________________________________________________________
cboCRTMVTDEV_OD.Clear
cboCRTMVTDEV_OD.AddItem ""
xSQL = "select distinct CRTMVTDEV from " & paramIBM_Library_SABSPE & ".YCRTMVT0 order by CRTMVTDEV"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    cboCRTMVTDEV_OD.AddItem Trim(rsSab("CRTMVTDEV"))
    rsSab.MoveNext
Loop

'___________________________________________________________________________
arrYCPTRUB0_Load

cboYCRTCPT0_CRTCPTRUB.Clear
cboYCRTCPT0_CRTCPTRUB.ForeColor = vbMagenta
cboYCRTCPT0_CRTCPTRUB.AddItem "      "

cboYCRTMVT0_CRTMVTRUB.Clear
cboYCRTMVT0_CRTMVTRUB.ForeColor = vbMagenta
cboYCRTMVT0_CRTMVTRUB.AddItem "      "

cboSelect_CRTCPTRUB.Clear
cboSelect_CRTCPTRUB.AddItem "      "

cboCRTMVTRUB_OD.Clear
cboCRTMVTRUB_OD.AddItem "      "

For K = 1 To arrCRT_Rub_Nb
    cboSelect_CRTCPTRUB.AddItem arrCRT_Rub(K).Code & " -" & arrCRT_Rub(K).Lib
    cboYCRTCPT0_CRTCPTRUB.AddItem arrCRT_Rub(K).Code & " -" & arrCRT_Rub(K).Lib
    cboYCRTMVT0_CRTMVTRUB.AddItem arrCRT_Rub(K).Code & " -" & arrCRT_Rub(K).Lib
    cboCRTMVTRUB_OD.AddItem arrCRT_Rub(K).Code & " -" & arrCRT_Rub(K).Lib
Next K
'___________________________________________________________________________

cboYCRTCPT0_CRTCPTSTA.ForeColor = vbMagenta
cboYCRTCPT0_CRTCPTSTA.Clear
cboYCRTCPT0_CRTCPTSTA.AddItem "E - exclure ce compte"
cboYCRTCPT0_CRTCPTSTA.AddItem "I - inclure les mvts (NON déclarés)"
cboYCRTCPT0_CRTCPTSTA.AddItem "* - inclure les mvts (déclarable)"
cboYCRTCPT0_CRTCPTSTA.AddItem "M - à affecter manuellement"
cboYCRTCPT0_CRTCPTSTA.AddItem "? - à définir"

cboYCRTMVT0_CRTMVTSTA.ForeColor = vbBlack
cboYCRTMVT0_CRTMVTSTA.Clear
cboYCRTMVT0_CRTMVTSTA.AddItem "I - à ignorer"
cboYCRTMVT0_CRTMVTSTA.AddItem "  - sélectionné"
cboYCRTMVT0_CRTMVTSTA.AddItem "A - Annulé"
cboYCRTMVT0_CRTMVTSTA.AddItem "? - à revoir"

cboYCRTMVT0_CRTMVTORIG.Clear
cboYCRTMVT0_CRTMVTORIG.AddItem "* - automatique"
cboYCRTMVT0_CRTMVTORIG.AddItem "M - modification"
cboYCRTMVT0_CRTMVTORIG.AddItem "+ - saisie"

cboSelect_CRTCPTSTA.Clear
cboSelect_CRTCPTSTA.AddItem "   "
cboSelect_CRTCPTSTA.AddItem "E - exclure ce compte"
cboSelect_CRTCPTSTA.AddItem "I - inclure les mvts (NON déclarés)"
cboSelect_CRTCPTSTA.AddItem "* - inclure les mvts (déclarable)"
cboSelect_CRTCPTSTA.AddItem "M - à affecter manuellement"
cboSelect_CRTCPTSTA.AddItem "? - à définir"

cboSelect_CRTMVTSTA.Clear
cboSelect_CRTMVTSTA.AddItem "  - déclarables"
cboSelect_CRTMVTSTA.AddItem "# - tous sauf I "
cboSelect_CRTMVTSTA.AddItem "* - tous les mouvements"
cboSelect_CRTMVTSTA.AddItem "I - à ignorer"
cboSelect_CRTMVTSTA.AddItem "A - annulés"
cboSelect_CRTMVTSTA.AddItem "? - à revoir"
cboSelect_CRTMVTSTA.ListIndex = 0

cboSelect_CRTMVTORIG.Clear
cboSelect_CRTMVTORIG.AddItem "   "
cboSelect_CRTMVTORIG.AddItem "* - automatique"
cboSelect_CRTMVTORIG.AddItem "M - modifié"
cboSelect_CRTMVTORIG.AddItem "+ - ajout"
'___________________________________________________________________________

cboYCRTMVT0_CRTMVTCLIP.Clear
cboYCRTMVT0_CRTMVTCLIP.ForeColor = vbMagenta
cboCRTMVTCLIP_OD.Clear
cboSelect_CRTMVTCLIP.Clear
cboSelect_CRTMVTCLIP.AddItem "  "
For K = 1 To sabPays_NB
    cboYCRTMVT0_CRTMVTCLIP.AddItem sabPays(K).Id & " -" & sabPays(K).Nom
    cboCRTMVTCLIP_OD.AddItem sabPays(K).Id & " -" & sabPays(K).Nom
    cboSelect_CRTMVTCLIP.AddItem sabPays(K).Id & " -" & sabPays(K).Nom
Next K
'___________________________________________________________________________
cboParam_Nomenclature.Clear
arrNomenclature(1) = "1-Hors PFD AYANT un lien avec le compte de résultat"
cboParam_Nomenclature.AddItem arrNomenclature(1)
arrNomenclature(2) = "2-Hors PFD N'AYANT PAS de lien avec le compte de résultat"
cboParam_Nomenclature.AddItem arrNomenclature(2)
arrNomenclature(3) = "3-PFD (produits financiers dérivés)"
cboParam_Nomenclature.AddItem arrNomenclature(3)
'___________________________________________________________________________

rsYCRTLOG0_Init zYCRTLOG0
'========================================================================

If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0

blnControl = True


cmdSelect_Reset
Me.Enabled = True

End Sub
Private Sub fgYSWISAB0_Display(lSWISABOPEC As String, LSWISABOPEN As Long)
Dim xSQL As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgYSWISAB0.Visible = False
fgYSWISAB0_Reset

fgYSWISAB0.Rows = 1
fgYSWISAB0.FormatString = fgYSWISAB0_FormatString
fgYSWISAB0.Row = 0

currentAction = "fgYSWISAB0_Display"
'X = "SELECT * FROM " & paramIBM_Library_SAB & ".ZMOUVEMG left outer join " & paramIBM_Library_SABSPE & ".YTVACOM0" _
'  & " on tvacometa = mouvemeta and tvacompla = mouvempla and tvacompie = mouvempie and tvacomecr = mouvemecr " _

'__________________________________________________________________________________________________
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0  left outer join " & paramIBM_Library_SABSPE & ".YSWISAB1" _
     & " on SWISABSWID = SWISAB1ID " _
     & " where SWISABOPEC = '" & lSWISABOPEC & "'" _
     & " and   SWISABOPEN = " & LSWISABOPEN _
     & " order by SWISABWAMJ , SWISABWHMS"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF

    fgYSWISAB0.Rows = fgYSWISAB0.Rows + 1
    fgYSWISAB0.Row = fgYSWISAB0.Rows - 1
    fgYSWISAB0_DisplayLine I


    rsSab.MoveNext

Loop
         
    

fgYSWISAB0.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgYSWISAB0.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub fgYSWISAB0_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long
Dim X As String, X2 As String

On Error Resume Next

If rsSab("SWISABWES") = "S" Then
    X = rsSab("SWISABWMTK") & " S"
    wColor = RGB(16, 96, 16)
Else
    X = rsSab("SWISABWMTK") & " E"
    wColor = vbBlue
End If

  

fgYSWISAB0.Col = 0: fgYSWISAB0.Text = X
fgYSWISAB0.CellForeColor = wColor

fgYSWISAB0.Col = 1: fgYSWISAB0.Text = rsSab("SWISABWBIC")
fgYSWISAB0.CellForeColor = wColor
fgYSWISAB0.Col = 2: fgYSWISAB0.Text = rsSab("SWISABWL20")
fgYSWISAB0.CellForeColor = wColor
Select Case rsSab("SWISABK20")
    Case "!": fgYSWISAB0.CellBackColor = RGB(220, 220, 255)
    Case Is <> " ": fgYSWISAB0.CellBackColor = RGB(220, 255, 220)
End Select

fgYSWISAB0.Col = 3: fgYSWISAB0.Text = Format$(CCur(rsSab("SWISABWMTD")), "### ### ### ##0.00")
fgYSWISAB0.CellForeColor = vbRed
fgYSWISAB0.CellFontBold = True
fgYSWISAB0.Col = 4: fgYSWISAB0.Text = rsSab("SWISABWDEV")
fgYSWISAB0.CellForeColor = wColor
fgYSWISAB0.Col = 5
If Not IsNull(rsSab("SWISABW50P")) Then
     Select Case rsSab("SWISABW71A")
        Case "O": X = "OUR - "
        Case "S": X = "SHA - "
        Case "B": X = "BEN - "
        Case Else: X = rsSab("SWISABW71A") & " - "
    End Select

    fgYSWISAB0.Text = X & rsSab("SWISABW50P") & " => " & rsSab("SWISABW59P")
    fgYSWISAB0.CellBackColor = mColor_Y1
End If

fgYSWISAB0.CellForeColor = wColor
fgYSWISAB0.Col = 6: fgYSWISAB0.Text = dateImp10_S(rsSab("SWISABWAMJ")) & " " & timeImp8(rsSab("SWISABWHMS"))
fgYSWISAB0.CellForeColor = RGB(80, 80, 80)

fgYSWISAB0.Col = 7
K = Val(mId$(rsSab("SWISABKSRV"), 2, 2)): fgYSWISAB0.Text = rsSab("SWISABWN20") 'arrService_Lib(K) & " - " & rsSab("SWISABWN20")
'fgYSWISAB0.Text = rsSab("SWISABWN20")
fgYSWISAB0.CellForeColor = wColor

fgYSWISAB0.Col = fgYSWISAB0_arrIndex: fgYSWISAB0.Text = lIndex
fgYSWISAB0.Col = 8: fgYSWISAB0.Text = rsSab("SWISABWID")
fgYSWISAB0.Col = 9: fgYSWISAB0.Text = rsSab("SWISABWIDL")
fgYSWISAB0.Col = 10: fgYSWISAB0.Text = rsSab("SWISABWIDH")
fgYSWISAB0.Col = 11: fgYSWISAB0.Text = rsSab("SWISABSWID")
If rsSab("SWISABWSTA") <> "V" Then
    For K = 0 To 11
        fgYSWISAB0.Col = K
        fgYSWISAB0.CellBackColor = mColor_W1
    Next K
End If
End Sub

Public Sub fgYSWISAB0_Reset()
fgYSWISAB0.Clear
fgYSWISAB0_Sort1 = 0: fgYSWISAB0_Sort2 = 0
fgYSWISAB0_Sort1_Old = -1
fgYSWISAB0_RowDisplay = 0: fgYSWISAB0_RowClick = 0
fgYSWISAB0_arrIndex = fgYSWISAB0.Cols - 1
blnfgYSWISAB0_DisplayLine = False
fgYSWISAB0_SortAD = 6
fgYSWISAB0.LeftCol = fgYSWISAB0.FixedCols

End Sub

Public Sub fgYSWISAB0_Sort()
If fgYSWISAB0.Rows > 1 Then
    fgYSWISAB0.Row = 1
    fgYSWISAB0.RowSel = fgYSWISAB0.Rows - 1
    
    If fgYSWISAB0_Sort1_Old = fgYSWISAB0_Sort1 Then
        If fgYSWISAB0_SortAD = 5 Then
            fgYSWISAB0_SortAD = 6
        Else
            fgYSWISAB0_SortAD = 5
        End If
    Else
        fgYSWISAB0_SortAD = 5
    End If
    fgYSWISAB0_Sort1_Old = fgYSWISAB0_Sort1
    
    fgYSWISAB0.Col = fgYSWISAB0_Sort1
    fgYSWISAB0.ColSel = fgYSWISAB0_Sort2
    fgYSWISAB0.Sort = fgYSWISAB0_SortAD
End If

End Sub

Public Sub fgYSWISAB0_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgYSWISAB0.Rows - 1
    fgYSWISAB0.Row = I
    fgYSWISAB0.Col = lK
    Select Case lK
        Case 3: fgYSWISAB0.Col = 3: X = Format$(Val(fgYSWISAB0.Text), "000000000000000.00")
        Case 4:
            fgYSWISAB0.Col = 4: X = Trim(fgYSWISAB0.Text)
            fgYSWISAB0.Col = 3: X = X & Format$(Val(fgYSWISAB0.Text), "000000000000000.00")
    End Select
    fgYSWISAB0.Col = fgYSWISAB0_arrIndex - 1
    fgYSWISAB0.Text = X
Next I

fgYSWISAB0_Sort1 = fgYSWISAB0_arrIndex - 1: fgYSWISAB0_Sort2 = fgYSWISAB0_arrIndex - 1
fgYSWISAB0_Sort
End Sub


Public Sub fgYSWISAB0_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgYSWISAB0.Visible = False
mRow = fgYSWISAB0.Row

If lRow > 0 And lRow < fgYSWISAB0.Rows Then
    fgYSWISAB0.Row = lRow
    For I = fgYSWISAB0_arrIndex To fgYSWISAB0.FixedCols Step -1
        fgYSWISAB0.Col = I: fgYSWISAB0.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgYSWISAB0.Row = mRow
    If fgYSWISAB0.Row > 0 Then
        lRow = fgYSWISAB0.Row
        lColor_Old = fgYSWISAB0.CellBackColor
        For I = fgYSWISAB0_arrIndex To fgYSWISAB0.FixedCols Step -1
          fgYSWISAB0.Col = I: fgYSWISAB0.CellBackColor = lColor
        Next I
    End If
End If
fgYSWISAB0.LeftCol = fgYSWISAB0.FixedCols
fgYSWISAB0.Visible = True
End Sub


Private Sub cmdSelect_SQL_1_Importation_YCRTCPT0()
Dim V, X As String, xSQL As String
Dim K As Integer, K2 As Integer
Dim blnTransaction As Boolean
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
blnTransaction = False
'________________________________________________________________________________

currentAction = "cmdSelect_SQL_1"
 
xSQL = "select COMPTECOM , COMPTEOBL from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTEFON <> '4'" _
     & " and CLIENACLI = '' and substring(COMPTEOBL,1,1) <> '9' " _
     & " and COMPTECOM not in (select CRTCPTCPT from " & paramIBM_Library_SABSPE & ".YCRTCPT0)" _
     & " order by COMPTECOM"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    newYCRTCPT0.CRTCPTCPT = rsSab("COMPTECOM")
    newYCRTCPT0.CRTCPTRUB = ""
    newYCRTCPT0.CRTCPTSTA = "E"
    X = rsSab("COMPTEOBL")
    For K = arrCRT_PCEC_Rub_Nb0 To arrCRT_PCEC_Rub_Nb
    
        If arrCRT_PCEC_Rub(K).PCEC > mId$(X, 1, arrCRT_PCEC_Rub(K).PCEC_Len) Then Exit For
        If arrCRT_PCEC_Rub(K).PCEC = mId$(X, 1, arrCRT_PCEC_Rub(K).PCEC_Len) Then
            If newYCRTCPT0.CRTCPTRUB = "" Then
                newYCRTCPT0.CRTCPTRUB = arrCRT_PCEC_Rub(K).Code
                newYCRTCPT0.CRTCPTSTA = "*"
                For K2 = 1 To arrCRT_RUB_I_Nb
                    If newYCRTCPT0.CRTCPTRUB = arrCRT_RUB_I_code(K2) Then
                        newYCRTCPT0.CRTCPTSTA = "I"
                        Exit For
                    End If
                Next K2
            Else
                newYCRTCPT0.CRTCPTRUB = ""
                newYCRTCPT0.CRTCPTSTA = "?"
                Exit For
            End If
            
        
        End If
    
    Next K
    If Not blnTransaction Then
        blnTransaction = True
        V = cnSAB_Transaction("BeginTrans")
        If Not IsNull(V) Then GoTo Error_MsgBox
    End If
    V = sqlYCRTCPT0_Insert(newYCRTCPT0)
    If Not IsNull(V) Then GoTo Error_MsgBox
    
    
    newYCRTLOG0 = zYCRTLOG0
    newYCRTLOG0.CRTLOGNAT = "C01"
    newYCRTLOG0.CRTLOGCPT = newYCRTCPT0.CRTCPTCPT
    newYCRTLOG0.CRTLOGTXT = "<CRTCPTRUB = " & newYCRTCPT0.CRTCPTRUB & ">" & "<CRTCPTSTA = " & newYCRTCPT0.CRTCPTSTA & ">"
    V = sqlYCRTLOG0_Insert(newYCRTLOG0)
    If Not IsNull(V) Then GoTo Error_MsgBox

    rsSab.MoveNext
Loop

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If blnTransaction Then
        If Not IsNull(V) Then
            V = cnSAB_Transaction("Rollback")
        Else
            V = cnSAB_Transaction("Commit")
        End If
    End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL_1_Importation_xls()
Dim V, X As String, xSQL As String
Dim K As Integer, K2 As Integer, wCol_Nb As Integer, wRow As Long
Dim blnTransaction As Boolean, blnOk As Boolean, blnCRTCPTRUB As Boolean
Dim arrRubrique(500) As String
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
blnTransaction = False
'________________________________________________________________________________

currentAction = "cmdSelect_SQL_1_Importation_xls"
 
'______________________________________________

Set appExcel = CreateObject("Excel.Application")
Set wbExcel = appExcel.Workbooks.Open("C:\Temp\BDF_CRT Import.xlsx")
Set wsExcel = wbExcel.Worksheets(1)
'__________________________________________________________________________________

Call lstErr_AddItem(lstErr, cmdContext, "Importation BDF_CRT.xls "): DoEvents

K = 3
Do
    K = K + 1
    If Trim(wsExcel.Cells(1, K)) = "" Then
        blnOk = True
    Else
        arrRubrique(K) = Trim(wsExcel.Cells(1, K))
    End If
    

Loop Until blnOk

wCol_Nb = K - 1
blnOk = False
wRow = 1
Do
    wRow = wRow + 1
    X = Trim(wsExcel.Cells(wRow, 2))
    Call lstErr_ChangeLastItem(lstErr, cmdContext, X): DoEvents

    If X = "" Then
        blnOk = True
    Else
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCRTCPT0" _
         & " where CRTCPTCPT = '" & X & "'"
        Set rsSab = cnsab.Execute(xSQL)
        blnCRTCPTRUB = False
        If rsSab.EOF Then
            Call MsgBox("Compte inconnu : " & X, vbCritical, "cmdSelect_SQL_1_Importation_xls")
        Else
            Call rsYCRTCPT0_GetBuffer(rsSab, oldYCRTCPT0)
            newYCRTCPT0 = oldYCRTCPT0
            If Trim(wsExcel.Cells(wRow, 1)) <> "" Then
                newYCRTCPT0.CRTCPTSTA = "E"
            Else
'_______________________________________________________________________________
                For K = 4 To wCol_Nb
                    If Trim(wsExcel.Cells(wRow, K)) <> "" Then
                        If blnCRTCPTRUB Then
                            Call MsgBox("Compte RUB multiples: " & X, vbCritical, "cmdSelect_SQL_1_Importation_xls")
                        Else
                            newYCRTCPT0.CRTCPTRUB = arrRubrique(K)
                            newYCRTCPT0.CRTCPTSTA = "*"
                            For K2 = 1 To arrCRT_RUB_I_Nb
                                If newYCRTCPT0.CRTCPTRUB = arrCRT_RUB_I_code(K2) Then
                                    newYCRTCPT0.CRTCPTSTA = "I"
                                    Exit For
                                End If
                            Next K2

                        End If
                   End If
                Next K
            
                If blnCRTCPTRUB Then
                    Call MsgBox("Compte sans RUB : " & X, vbCritical, "cmdSelect_SQL_1_Importation_xls")
                End If
            End If
'_______________________________________________________________________________
            If newYCRTCPT0.CRTCPTRUB = oldYCRTCPT0.CRTCPTRUB _
            And newYCRTCPT0.CRTCPTSTA = oldYCRTCPT0.CRTCPTSTA Then
            Else
                cmdYCRTCPT0_Update_Transaction
            End If
        End If

    End If


Loop Until blnOk

'

'wbExcel.Close
wbExcel.Saved = True
'____________________________________________________________________________________
appExcel.Quit

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If blnTransaction Then
        If Not IsNull(V) Then
            V = cnSAB_Transaction("Rollback")
        Else
            V = cnSAB_Transaction("Commit")
        End If
    End If
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdSelect_SQL_1()
Dim X As String
On Error GoTo Error_Handler


'________________________________________________________________________________

currentAction = "cmdSelect_SQL_1"
mSQL_Where = ""

X = Trim(txtSelect_CRTCPTCPT)
If X <> "" Then mSQL_Where = " and CRTCPTCPT like '" & X & "%'"
X = Trim(txtSelect_COMPTEOBL)
If X <> "" Then mSQL_Where = mSQL_Where & " and COMPTEOBL like '" & X & "%'"
X = Trim(mId$(cboSelect_CRTCPTSTA, 1, 1))
If X <> "" Then mSQL_Where = mSQL_Where & " and CRTCPTSTA ='" & X & "'"
X = Trim(mId$(cboSelect_CRTCPTRUB, 1, 5))
If X <> "" Then mSQL_Where = mSQL_Where & " and CRTCPTRUB ='" & X & "'"

 
mSQL_Exe = "select * from " & paramIBM_Library_SABSPE & ".YCRTCPT0 , " & paramIBM_Library_SABSPE & ".YBIACPT0" _
     & " where CRTCPTCPT = COMPTECOM" & mSQL_Where _
     & " order by CRTCPTCPT"
Set rsSab = cnsab.Execute(mSQL_Exe)

fgSelect_Display

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:


End Sub


Private Sub cmdSelect_SQL_Log()
Dim X As String, xWhere As String
On Error GoTo Error_Handler


'________________________________________________________________________________

currentAction = "cmdSelect_SQL_Log"
Call DTPicker_Control(txtSelect_Log_AmjMin, wAMJMin)
Call DTPicker_Control(txtSelect_Log_AmjMax, WAMJMax)

xWhere = " where CRTLOGUAMJ >= " & wAMJMin & " and CRTLOGUAMJ <= " & WAMJMax
'________________________________________________________________________________

X = Trim(txtSelect_CRTLOGCPT)
If X <> "" Then xWhere = xWhere & " and CRTLOGCPT like '" & X & "%'"

X = Trim(cboSelect_CRTLOGNAT)
If X <> "" Then xWhere = xWhere & " and CRTLOGNAT = '" & mId$(X, 1, 3) & "'"

 
mSQL_Exe = "select * from " & paramIBM_Library_SABSPE & ".YCRTLOG0 " _
     & xWhere _
     & " order by CRTLOGID"
Set rsSab = cnsab.Execute(mSQL_Exe)

fgLog_Display_YCRTLOG0

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:


End Sub


Private Sub cmdSelect_SQL_2c()
Dim V, X As String
On Error GoTo Error_Handler


'________________________________________________________________________________

currentAction = "cmdSelect_SQL_2"
mSQL_Where = " and CRTCPTSTA in ('*' , 'I' , 'M')"

X = Trim(txtSelect_CRTCPTCPT)
If X <> "" Then mSQL_Where = mSQL_Where & " and CRTCPTCPT like '" & X & "%'"
X = Trim(mId$(cboSelect_CRTCPTRUB, 1, 5))
If X <> "" Then mSQL_Where = mSQL_Where & " and CRTCPTRUB ='" & X & "'"

 
mSQL_Exe = "select * from " & paramIBM_Library_SABSPE & ".YCRTCPT0 , " & paramIBM_Library_SABSPE & ".YBIACPT0" _
     & " where CRTCPTCPT = COMPTECOM" & mSQL_Where _
     & " order by CRTCPTCPT"
Set rsSab = cnsab.Execute(mSQL_Exe)

fgSelect_Display

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:


End Sub


Private Sub cmdSelect_SQL_2_Importation_All()
Dim V, X As String
Dim rsSabX As New ADODB.Recordset

On Error GoTo Error_Handler

mAmjMin_CRTMVTDTR = mAmjMin_Exercice

X = "select distinct crtmvtdtr from " & paramIBM_Library_SABSPE & ".YCRTMVT0 " _
    & " where crtmvtdtr > " & mAmjMin_Exercice & " order by crtmvtdtr desc"
Set rsSabX = cnsab.Execute(X)

If Not rsSabX.EOF Then
    mAmjMin_CRTMVTDTR = rsSabX("CRTMVTDTR")
    mAmjMin_CRTMVTDTR = dateElp("Jour", 1, mAmjMin_CRTMVTDTR)
    
End If
'________________________________________________________________________________

currentAction = "cmdSelect_SQL_2_Importation_All"
    newYCRTLOG0 = zYCRTLOG0
    newYCRTLOG0.CRTLOGNAT = "M10"
    newYCRTLOG0.CRTLOGTXT = ""

    V = sqlYCRTLOG0_Insert_Transaction(newYCRTLOG0)

 
mSQL_Exe = "select CRTCPTCPT from " & paramIBM_Library_SABSPE & ".YCRTCPT0 " _
     & " where CRTCPTSTA in ('*' , 'I' , 'M')" _
     & " order by CRTCPTCPT"
Set rsSabX = cnsab.Execute(mSQL_Exe)

Do While Not rsSabX.EOF
    Call cmdSelect_SQL_2_Importation_YCRTMVT0(rsSabX("CRTCPTCPT"))
    rsSabX.MoveNext
Loop
    

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:


End Sub


Private Sub fgSelect_Display()

Dim K As Long, blnOk As Boolean, X As String
Dim wColor As Long

On Error GoTo Error_Handler
currentAction = "fgSelect_Display"
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
'Droits
fgSelect.Row = 0
K = 0
Do While Not rsSab.EOF
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    X = Trim(rsSab("CRTCPTSTA"))
    Select Case X
        Case "*": wColor = RGB(16, 96, 16)
        Case "I": wColor = RGB(96, 96, 96)
        Case "M": wColor = vbBlue
        Case "E": wColor = RGB(160, 160, 160)
        Case Else: wColor = vbRed
        
    End Select
    
    fgSelect.Col = 0: fgSelect.Text = X: fgSelect.CellForeColor = wColor
    fgSelect.Col = 1: fgSelect.Text = rsSab("CRTCPTCPT"): fgSelect.CellForeColor = wColor
    
    fgSelect.Col = 2: fgSelect.Text = rsSab("COMPTEINT"): fgSelect.CellForeColor = wColor
    X = Trim(rsSab("CRTCPTRUB"))
    fgSelect.Col = 3: fgSelect.Text = X: fgSelect.CellForeColor = wColor
    
    If X <> arrCRT_Rub(arrCRT_Rub_K).Code Then
        For arrCRT_Rub_K = 0 To arrCRT_Rub_Nb
            If X = arrCRT_Rub(arrCRT_Rub_K).Code Then Exit For
        Next arrCRT_Rub_K
    End If
    
    fgSelect.Col = 4: fgSelect.Text = arrCRT_Rub(arrCRT_Rub_K).Lib: fgSelect.CellForeColor = wColor
    rsSab.MoveNext

Loop

fgSelect.Visible = True
cmdPrint.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub fgLog_Display_YCRTLOG0()

Dim K As Long, blnOk As Boolean, X As String
Dim wColor As Long

On Error GoTo Error_Handler
currentAction = "fgLog_Display_YCRTLOG0"
fgLog.Visible = False
'fgLog_Reset

fgLog.Rows = 1
'fgLog.FormatString = fgLog_FormatString
fgLog.Row = 0
K = 0
Do While Not rsSab.EOF
    fgLog.Rows = fgLog.Rows + 1
    fgLog.Row = fgLog.Rows - 1
    
    X = rsSab("CRTLOGNAT")
    Select Case mId$(X, 2, 2)
        Case "01": wColor = vbBlue
        Case "02": wColor = mColor_GB
        Case "03": wColor = vbRed
        Case Else: wColor = RGB(16, 16, 16)
        
    End Select

    fgLog.Col = 0: fgLog.Text = rsSab("CRTLOGID"): fgLog.CellForeColor = wColor
    fgLog.Col = 1: fgLog.Text = rsSab("CRTLOGCPT"): fgLog.CellForeColor = wColor
    
    fgLog.Col = 2: fgLog.Text = X: fgLog.CellForeColor = wColor
    fgLog.Col = 3: fgLog.Text = rsSab("CRTLOGUUSR"): fgLog.CellForeColor = wColor
        
    fgLog.Col = 4: fgLog.Text = dateImp10_S(rsSab("CRTLOGUAMJ")) & "  " & timeImp8(rsSab("CRTLOGUHMS")): fgLog.CellForeColor = wColor
    fgLog.Col = 5: fgLog.Text = rsSab("CRTLOGTXT"): fgLog.CellForeColor = wColor
    rsSab.MoveNext

Loop
fgLog.ZOrder 0
fgLog.Visible = True
cmdPrint.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgLog.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub fgDetail_Display()

Dim curX As Currency, mCRTMVTSTA As String
Dim wColor As Long

On Error GoTo Error_Handler
currentAction = "fgDetail_Display"
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString


fgDetail.Row = 0

If cmdSelect_SQL_K = "2+" Then
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail.Col = 9: fgDetail.Text = "Ajouter une écriture"
    fgDetail.CellBackColor = mColor_G9
End If

Do While Not rsSab.EOF
    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    mCRTMVTSTA = rsSab("CRTMVTSTA")
    Select Case mCRTMVTSTA
        Case " ": wColor = RGB(16, 96, 16)
        Case "I": wColor = RGB(160, 160, 160)
        Case "A": wColor = vbRed
        Case "M": wColor = vbBlue
        Case Else: wColor = vbMagenta
        
    End Select
    
    fgDetail.Col = 0: fgDetail.Text = mCRTMVTSTA: fgDetail.CellForeColor = wColor
    fgDetail.Col = 1: fgDetail.Text = rsSab("CRTMVTORIG"): fgDetail.CellForeColor = wColor
    
    fgDetail.Col = 2: fgDetail.Text = dateImp10_S(rsSab("CRTMVTDTR")): fgDetail.CellForeColor = wColor

    fgDetail.Col = 3: fgDetail.Text = rsSab("CRTMVTRUB"): fgDetail.CellForeColor = wColor
    fgDetail.Col = 4: fgDetail.Text = rsSab("CRTMVTCLIC") & " " & rsSab("CRTMVTCLIN"): fgDetail.CellForeColor = wColor
    
    fgDetail.Col = 5: fgDetail.Text = rsSab("CRTMVTCLIP"): fgDetail.CellForeColor = wColor
    curX = rsSab("CRTMVTMTE")
    fgDetail.Col = 6: fgDetail.Text = Format$(curX, "### ### ### ##0.00")
    If mCRTMVTSTA = "I" Then
        fgDetail.CellForeColor = wColor
    Else
        If curX < 0 Then
            fgDetail.CellForeColor = vbRed
        Else
            fgDetail.CellForeColor = vbBlue
        End If
    End If
    
    fgDetail.Col = 7: fgDetail.Text = rsSab("CRTMVTDEV"): fgDetail.CellForeColor = wColor
    fgDetail.Col = 8
    fgDetail.Text = rsSab("CRTMVTSER") & " " & rsSab("CRTMVTSSE") & " " & rsSab("CRTMVTOPE") & " " & rsSab("CRTMVTEVE") & " " & rsSab("CRTMVTDOS")
    fgDetail.CellForeColor = wColor
    fgDetail.Col = 10: fgDetail.Text = rsSab("CRTMVTCPT"): fgDetail.CellForeColor = wColor
    fgDetail.Col = 11: fgDetail.Text = rsSab("CRTMVTPIE"): fgDetail.CellForeColor = wColor
    fgDetail.Col = 12: fgDetail.Text = rsSab("CRTMVTECR"): fgDetail.CellForeColor = wColor
    
     If Not IsNull(rsSab("LIBELLIB1")) Then
        fgDetail.Col = 9: fgDetail.Text = Trim(rsSab("LIBELLIB1")): fgDetail.CellForeColor = wColor
    End If
    rsSab.MoveNext

Loop

fgDetail.Visible = True
cmdPrint.Visible = True


Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgDetail.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


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



Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim wFct As String

mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

wFct = UCase$(Trim(mId$(Msg, 1, 12)))
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


Private Sub cboCRTMVTDEV_OD_Change()
fraOD_Display_Cours
End Sub

Private Sub cboCRTMVTDEV_OD_Click()
fraOD_Display_Cours

End Sub

Private Sub cboSelect_CRTCPTRUB_Change()
cmdSelect_Clear
End Sub

Private Sub cboSelect_CRTCPTRUB_Click()
cmdSelect_Clear
End Sub

Private Sub cboSelect_CRTCPTRUB_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub


Private Sub cboSelect_CRTCPTSTA_Change()
cmdSelect_Clear
End Sub


Private Sub cboSelect_CRTCPTSTA_Click()
cmdSelect_Clear

End Sub


Private Sub cboSelect_CRTLOGNAT_Click()
cmdSelect_Clear

End Sub


Private Sub cboSelect_CRTMVTORIG_Change()
cmdSelect_Clear

End Sub

Private Sub cboSelect_CRTMVTORIG_Click()
cmdSelect_Clear

End Sub

Private Sub cboSelect_CRTMVTSTA_Change()
cmdSelect_Clear

End Sub

Private Sub cboSelect_CRTMVTSTA_Click()
cmdSelect_Clear

End Sub

Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

End Sub

Private Sub cmdOD_Add_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass

newYCRTMVT0 = zYCRTMVT0_OD

If IsNull(fraOD_Control) Then
     Call cmdYCRTMVT0_OD_Transaction("Add")
End If


Exit_sub:

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdOD_Delete_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass


newYCRTMVT0 = oldYCRTMVT0
newYCRTMVT0.CRTMVTSTA = "A"

cmdYCRTMVT0_OD_Transaction "Delete"

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdOD_Log_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, ">cmdYCRTMVT0_Log")



mSQL_Exe = "select * from " & paramIBM_Library_SABSPE & ".YCRTLOG0 " _
     & " where CRTLOGETA = " & currentSAB_ETA & " and CRTLOGPLA = " & currentSAB_PLA _
     & " and CRTLOGPIE = " & oldYCRTMVT0.CRTMVTPIE & " and CRTLOGECR = " & oldYCRTMVT0.CRTMVTECR _
     & " order by CRTLOGID"
Set rsSab = cnsab.Execute(mSQL_Exe)

fgLog_Display_YCRTLOG0

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdOD_Quit_Click()
fraOD.Visible = False
End Sub

Private Sub cmdOD_Update_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
newYCRTMVT0 = oldYCRTMVT0

If IsNull(fraOD_Control) Then
    newYCRTMVT0.CRTMVTSTA = " "

     Call cmdYCRTMVT0_OD_Transaction("Update")
End If


Exit_sub:

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Add_Click()
Dim xSQL As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

blnOk = True
New_YBIATAB0.BIATABID = Old_YBIATAB0.BIATABID

New_YBIATAB0.BIATABK1 = Trim(txtParam_K1)
New_YBIATAB0.BIATABK2 = Trim(txtParam_K2)
If Trim(New_YBIATAB0.BIATABK1) = "" Then
    Call MsgBox("Préciser le code K1", vbCritical, "BDF_CRT : paramétrage " & New_YBIATAB0.BIATABID)
    blnOk = False
End If

If New_YBIATAB0.BIATABID = "CRT_Mvt=Ann" Then

    If Not IsNumeric(Trim(txtParam_K1)) Then
            Call MsgBox("le code K1 n'est pas numérique", vbCritical, "BDF_CRT : paramétrage " & New_YBIATAB0.BIATABID)
            blnOk = False
    Else
        New_YBIATAB0.BIATABK1 = Format(Val(txtParam_K1), "000000000000")
        New_YBIATAB0.BIATABK2 = ""
    
    End If
    If Trim(txtParam_CRTMVTCPT) = "" Then
        Call MsgBox("Préciser le COMPTE", vbCritical, "BDF_CRT : paramétrage " & New_YBIATAB0.BIATABID)
        blnOk = False
    Else
        If Trim(txtParam_CRTMVTSER) = "" Then
            Call MsgBox("Préciser le SERVICE", vbCritical, "BDF_CRT : paramétrage " & New_YBIATAB0.BIATABID)
            blnOk = False
        Else
            If Trim(txtParam_CRTMVTSSE) = "" Then
                Call MsgBox("Préciser le SOUS-SERVICE", vbCritical, "BDF_CRT : paramétrage " & New_YBIATAB0.BIATABID)
                blnOk = False
            Else
                If Trim(txtParam_CRTMVTOPE) = "" Then
                    Call MsgBox("Préciser le CODE OPERATION", vbCritical, "BDF_CRT : paramétrage " & New_YBIATAB0.BIATABID)
                    blnOk = False
                End If
            End If
        End If
    End If
End If

If blnOk Then
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = '" & New_YBIATAB0.BIATABID & "' and BIATABK1 = '" & New_YBIATAB0.BIATABK1 & "'" _
     & " and BIATABK2 = '" & New_YBIATAB0.BIATABK2 & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        Call MsgBox("Ce code existe déjà", vbCritical, "BDF_CRT : paramétrage")
    Else
    
        Select Case Trim(New_YBIATAB0.BIATABID)
            Case "CRT_Rubrique"
                New_YBIATAB0.BIATABTXT = mId$(cboParam_Nomenclature, 1, 1) & "-" & Trim(txtParam_Txt)
            Case "CRT_Mvt=Ann"
                New_YBIATAB0.BIATABTXT = Space(34)
                Mid$(New_YBIATAB0.BIATABTXT, 1, 20) = Trim(txtParam_CRTMVTCPT)
                Mid$(New_YBIATAB0.BIATABTXT, 22, 2) = Trim(txtParam_CRTMVTSER)
                Mid$(New_YBIATAB0.BIATABTXT, 25, 2) = Trim(txtParam_CRTMVTSSE)
                Mid$(New_YBIATAB0.BIATABTXT, 28, 3) = Trim(txtParam_CRTMVTOPE)
                Mid$(New_YBIATAB0.BIATABTXT, 32, 3) = Trim(txtParam_CRTMVTEVE)
            Case Else
                New_YBIATAB0.BIATABTXT = Trim(txtParam_Txt)
        End Select
        If IsNull(Parametrage_New) Then lstParam_Load New_YBIATAB0.BIATABID
    End If
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_CRT_Mvt_Ann_Click()
Dim V, X As String, xWhere As String, blnOk As Boolean

Me.Enabled = False: Me.MousePointer = vbHourglass
blnOk = True
X = Old_YBIATAB0.BIATABTXT & Space$(50)
'txtParam_CRTMVTCPT = Trim(mId$(X, 1, 20))
'txtParam_CRTMVTSER = Trim(mId$(X, 22, 2))
'txtParam_CRTMVTSSE = Trim(mId$(X, 25, 2))
'txtParam_CRTMVTOPE = Trim(mId$(X, 28, 3))
'txtParam_CRTMVTEVE = Trim(mId$(X, 32, 3))

xWhere = " where CRTMVTSTA = '?'"
If Trim(mId$(X, 1, 20)) = "" Then
    blnOk = False
Else
    If InStr(Trim(mId$(X, 1, 20)), "%") > 0 Then
        xWhere = xWhere & " and CRTMVTCPT like '" & Trim(mId$(X, 1, 20)) & "'"
    Else
        xWhere = xWhere & " and CRTMVTCPT = '" & Trim(mId$(X, 1, 20)) & "'"
    End If
End If

If Trim(mId$(X, 22, 2)) = "" Then
    blnOk = False
Else
    xWhere = xWhere & " and CRTMVTSER = '" & Trim(mId$(X, 22, 2)) & "'"
End If

If Trim(mId$(X, 25, 2)) = "" Then
    blnOk = False
Else
    xWhere = xWhere & " and CRTMVTSSE= '" & Trim(mId$(X, 25, 2)) & "'"
End If

If Trim(mId$(X, 28, 3)) = "" Then
    blnOk = False
Else
    xWhere = xWhere & " and CRTMVTOPE = '" & Trim(mId$(X, 28, 3)) & "'"
End If
 If Trim(mId$(X, 32, 3)) <> "" Then
    xWhere = xWhere & " and CRTMVTEVE = '" & Trim(mId$(X, 32, 3)) & "'"
End If
   
If blnOk Then

    newYCRTLOG0 = zYCRTLOG0
    newYCRTLOG0.CRTLOGNAT = "P10"
    newYCRTLOG0.CRTLOGTXT = Old_YBIATAB0.BIATABK1 & " - " & Old_YBIATAB0.BIATABTXT
    newYCRTLOG0.CRTLOGTXT = "<BIATABID = " & Trim(Old_YBIATAB0.BIATABID) & " | " & ">" _
                      & "<BIATABK1 = " & Trim(Old_YBIATAB0.BIATABK1) & " | " & ">" _
                      & "<BIATABK2 = " & Trim(Old_YBIATAB0.BIATABK2) & " | " & ">" _
                      & "<BIATABTXT = " & Trim(Old_YBIATAB0.BIATABTXT) & " | " & ">"

    V = sqlYCRTLOG0_Insert_Transaction(newYCRTLOG0)
    
    X = "select * from " & paramIBM_Library_SABSPE & ".YCRTMVT0 " _
         & xWhere _
         & " order by CRTMVTDTR , CRTMVTPIE , CRTMVTECR"
    Set rsSab = cnsab.Execute(X)
    Do While Not rsSab.EOF
        Call rsYCRTMVT0_GetBuffer(rsSab, oldYCRTMVT0)
        
        newYCRTMVT0 = oldYCRTMVT0
        newYCRTMVT0.CRTMVTSTA = "A"
        cmdYCRTMVT0_Update_Transaction

        rsSab.MoveNext
    
    Loop
    
End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Delete_Click()
Dim X As String
Me.Enabled = False: Me.MousePointer = vbHourglass

X = Trim(txtParam_K1)
If X <> Trim(Old_YBIATAB0.BIATABK1) Then
    Call MsgBox("Le code K1 a été modifié," & vbCrLf & " la suppression n'est pas possible", vbCritical, "BDF_CRT : paramétrage")
Else
    X = Trim(txtParam_K2)
    If X <> Trim(Old_YBIATAB0.BIATABK2) Then
        Call MsgBox("Le code K2 a été modifié," & vbCrLf & " la suppression n'est pas possible", vbCritical, "BDF_CRT : paramétrage")
    Else
    
        If IsNull(Parametrage_Delete) Then lstParam_Load Old_YBIATAB0.BIATABID
    End If
End If


Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdParam_Quit_Click()
fraParam_Display.Visible = False
End Sub

Private Sub cmdParam_Update_Click()
Dim X As String, XX As String, blnOk As Boolean
Me.Enabled = False: Me.MousePointer = vbHourglass

blnOk = True
New_YBIATAB0.BIATABID = Old_YBIATAB0.BIATABID

New_YBIATAB0.BIATABK1 = Trim(txtParam_K1)
New_YBIATAB0.BIATABK2 = Trim(txtParam_K2)

If New_YBIATAB0.BIATABK1 <> Old_YBIATAB0.BIATABK1 Then
    Call MsgBox("Le code K1 a été modifié," & vbCrLf & " la mise à jour n'est pas possible", vbCritical, "BDF_CRT : paramétrage")
    blnOk = False
Else
    If New_YBIATAB0.BIATABK2 <> Old_YBIATAB0.BIATABK2 Then
        Call MsgBox("Le code K2 a été modifié," & vbCrLf & " la mise à jour n'est pas possible", vbCritical, "BDF_CRT : paramétrage")
        blnOk = False
    End If
End If
If New_YBIATAB0.BIATABID = "CRT_Mvt=Ann" Then
    If Trim(txtParam_CRTMVTCPT) = "" Then
        Call MsgBox("Préciser le COMPTE", vbCritical, "BDF_CRT : paramétrage " & New_YBIATAB0.BIATABID)
        blnOk = False
    Else
        If Trim(txtParam_CRTMVTSER) = "" Then
            Call MsgBox("Préciser le SERVICE", vbCritical, "BDF_CRT : paramétrage " & New_YBIATAB0.BIATABID)
            blnOk = False
        Else
            If Trim(txtParam_CRTMVTSSE) = "" Then
                Call MsgBox("Préciser le SOUS-SERVICE", vbCritical, "BDF_CRT : paramétrage " & New_YBIATAB0.BIATABID)
                blnOk = False
            Else
                If Trim(txtParam_CRTMVTOPE) = "" Then
                    Call MsgBox("Préciser le CODE OPERATION", vbCritical, "BDF_CRT : paramétrage " & New_YBIATAB0.BIATABID)
                    blnOk = False
                End If
            End If
        End If
    End If
End If

If blnOk Then
        
    Select Case Trim(New_YBIATAB0.BIATABID)
        Case "CRT_Rubrique"
            New_YBIATAB0.BIATABTXT = mId$(cboParam_Nomenclature, 1, 1) & "-" & Trim(txtParam_Txt)
         Case "CRT_Mvt=Ann"
            New_YBIATAB0.BIATABTXT = Space(34)
            Mid$(New_YBIATAB0.BIATABTXT, 1, 20) = Trim(txtParam_CRTMVTCPT)
            Mid$(New_YBIATAB0.BIATABTXT, 22, 2) = Trim(txtParam_CRTMVTSER)
            Mid$(New_YBIATAB0.BIATABTXT, 25, 2) = Trim(txtParam_CRTMVTSSE)
            Mid$(New_YBIATAB0.BIATABTXT, 28, 3) = Trim(txtParam_CRTMVTOPE)
            Mid$(New_YBIATAB0.BIATABTXT, 32, 3) = Trim(txtParam_CRTMVTEVE)
       Case Else
            New_YBIATAB0.BIATABTXT = Trim(txtParam_Txt)
    End Select
    If IsNull(Parametrage_Update) Then lstParam_Load Old_YBIATAB0.BIATABID
End If




Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdPrint_Click()

Dim X As String, I As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, ">cmdPrint : Initialisation ")

Select Case SSTab1.Tab
    Case 0:
        Select Case cmdSelect_SQL_K
            Case "1": Call cmdPrint_Excel("Comptes")
            Case "2", "2c", "Annulation$": cmdPrint_Excel ("Mouvements")
        End Select
    Case 1
        cmdPrint_Excel ("Paramétrage")
        
    End Select
Call lstErr_AddItem(lstErr, cmdPrint, "<cmdPrint : terminé ")

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdSAB_Dossier_DB_Click()
blnfrmSAB_Dossier_DB = True
Call frmSAB_Dossier_DB.Form_Init("", txtD_COMPTECOM, mAmjMin_Exercice, mAmjMax_Exercice, "", "", "", 0)

End Sub

Private Sub cmdYCRTCPT0_Exclure_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

Call MsgBox("Les mvts sélectionnés dans YCRTMVT0 ne sont pas supprimés", vbInformation, "Attention : Compte à ignorer")

newYCRTCPT0 = oldYCRTCPT0
newYCRTCPT0.CRTCPTSTA = "E"

cmdYCRTCPT0_Update_Transaction

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdYCRTCPT0_Ignore_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

Call MsgBox("Les mvts sélectionnés dans YCRTMVT0 ne sont pas supprimés", vbInformation, "Attention : Compte à ignorer")

newYCRTCPT0 = oldYCRTCPT0
newYCRTCPT0.CRTCPTSTA = "I"

cmdYCRTCPT0_Update_Transaction

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdYCRTCPT0_Quit_Click()
fraYCRTCPT0.Visible = False
End Sub

Private Sub cmdYCRTCPT0_Update_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass

newYCRTCPT0 = oldYCRTCPT0
newYCRTCPT0.CRTCPTRUB = mId$(cboYCRTCPT0_CRTCPTRUB, 1, 5)
newYCRTCPT0.CRTCPTSTA = mId$(cboYCRTCPT0_CRTCPTSTA, 1, 1)

If Trim(newYCRTCPT0.CRTCPTRUB) = "" Then
    Call MsgBox("Préciser la rubrique CRT", vbInformation, "Attention : Compte modifié")
    GoTo Exit_sub
End If

If Trim(newYCRTCPT0.CRTCPTSTA) = "" Then
    Call MsgBox("Préciser le code état", vbInformation, "Attention : Compte modifié")
    GoTo Exit_sub
End If

cmdYCRTCPT0_Update_Transaction

Exit_sub:

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdYCRTMVT0_Ignore_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
newYCRTMVT0 = oldYCRTMVT0
newYCRTMVT0.CRTMVTSTA = "I"

cmdYCRTMVT0_Update_Transaction

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdYCRTMVT0_Log_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, ">cmdYCRTMVT0_Log")



mSQL_Exe = "select * from " & paramIBM_Library_SABSPE & ".YCRTLOG0 " _
     & " where CRTLOGETA = " & currentSAB_ETA & " and CRTLOGPLA = " & currentSAB_PLA _
     & " and CRTLOGPIE = " & oldYCRTMVT0.CRTMVTPIE & " and CRTLOGECR = " & oldYCRTMVT0.CRTMVTECR _
     & " order by CRTLOGID"
Set rsSab = cnsab.Execute(mSQL_Exe)

fgLog_Display_YCRTLOG0

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdYCRTMVT0_Mvt_Ann_Click()
Dim V, xSQL As String
'oldYCRTMVT0

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_Mvt=Ann'" _
     & "order by BIATABK1 desc"

Set rsSabX = cnsab.Execute(xSQL)
If rsSabX.EOF Then
    New_YBIATAB0.BIATABK1 = 1
Else
    New_YBIATAB0.BIATABK1 = Format(Val(rsSabX("BIATABK1")) + 1, "000000000000")
End If
   
New_YBIATAB0.BIATABID = "CRT_Mvt=Ann"
New_YBIATAB0.BIATABK2 = ""
New_YBIATAB0.BIATABTXT = Space(34)
Mid$(New_YBIATAB0.BIATABTXT, 1, 20) = oldYCRTMVT0.CRTMVTCPT
Mid$(New_YBIATAB0.BIATABTXT, 22, 2) = oldYCRTMVT0.CRTMVTSER
Mid$(New_YBIATAB0.BIATABTXT, 25, 2) = oldYCRTMVT0.CRTMVTSSE
Mid$(New_YBIATAB0.BIATABTXT, 28, 3) = oldYCRTMVT0.CRTMVTOPE
Mid$(New_YBIATAB0.BIATABTXT, 32, 3) = oldYCRTMVT0.CRTMVTEVE
V = Parametrage_New
If Not IsNull(V) Then GoTo Exit_sub

Old_YBIATAB0 = New_YBIATAB0

Call cmdParam_CRT_Mvt_Ann_Click

cmdSelect_SQL_2

'=================================================================
GoTo Exit_sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name
Exit_sub:
    Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdYCRTMVT0_Question_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
newYCRTMVT0 = oldYCRTMVT0
newYCRTMVT0.CRTMVTSTA = "?"

cmdYCRTMVT0_Update_Transaction

Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdYCRTMVT0_Quit_Click()
fraYCRTMVT0.Visible = False

End Sub

Private Sub cmdYCRTMVT0_Update_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

newYCRTMVT0 = oldYCRTMVT0
newYCRTMVT0.CRTMVTRUB = mId$(cboYCRTMVT0_CRTMVTRUB, 1, 5)
newYCRTMVT0.CRTMVTCLIP = mId$(cboYCRTMVT0_CRTMVTCLIP, 1, 2)

If Trim(newYCRTMVT0.CRTMVTRUB) = "" Then
    Call MsgBox("Préciser la rubrique CRT", vbInformation, "cmdYCRTMVT0_Update")
    GoTo Exit_sub
End If

If Trim(newYCRTMVT0.CRTMVTCLIP) = "" Then
    Call MsgBox("Préciser le pays", vbInformation, "cmdYCRTMVT0_Update")
    GoTo Exit_sub
End If

newYCRTMVT0.CRTMVTORIG = "M"
If InStr(mPays_Exclus, Trim(newYCRTMVT0.CRTMVTCLIP)) > 0 Then
    newYCRTMVT0.CRTMVTSTA = "I"
Else
    newYCRTMVT0.CRTMVTSTA = " "
End If

If InStr(mClients_Exclus, newYCRTMVT0.CRTMVTCLIN) > 0 Then newYCRTMVT0.CRTMVTSTA = "I"

cmdYCRTMVT0_Update_Transaction


Exit_sub:

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub fgBIAMVT_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

If y <= fgBIAMVT.RowHeightMin Then
Else
    If fgBIAMVT.Rows > 1 Then
        fgBIAMVT.Col = 8:  xYBIAMVTH.MOUVEMPIE = Val(fgBIAMVT.Text)
        fgBIAMVT.Col = 9: xYBIAMVTH.MOUVEMECR = Val(fgBIAMVT.Text)
        Call fraYTVACOM0_Display(xYBIAMVTH.MOUVEMPIE, xYBIAMVTH.MOUVEMECR)

        
   End If
End If
fgBIAMVT.LeftCol = 0
End Sub


Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wMOUVEMPIE As Long, wMOUVEMECR As Long
On Error Resume Next


If y <= fgDetail.RowHeightMin Then
    fgDetail.Visible = False
    Select Case fgDetail.Col
        Case Is <= 12: fgDetail_Sort1 = fgDetail.Col: fgDetail_Sort2 = fgDetail.Col: fgdetail_Sort
    End Select
    fgDetail.Visible = True
Else
    If fgDetail.Rows > 1 Then
        Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
        fraOD.Visible = False
        fgDetail.Col = 11:  wMOUVEMPIE = Val(fgDetail.Text)
        fgDetail.Col = 12:  wMOUVEMECR = Val(fgDetail.Text)
        If wMOUVEMPIE = 0 Then
            Call fraOD_Display(wMOUVEMPIE, wMOUVEMECR)
        Else
            Call fraYCRTMVT0_Display(wMOUVEMPIE, wMOUVEMECR)
        End If
   End If
       

   End If

fgDetail.LeftCol = 0


End Sub


Private Sub cmdSelect_SQL_9xml()
Dim V, X As String
Dim xSQL As String, K As Long
Dim xWhere As String, xAnd As String
Dim Nb As Long
On Error GoTo Error_Handler

Close

Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL_9"): DoEvents
currentAction = "déclaration CRT à la direction de la balance des paiements"

xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YCRTMVT0 " _
     & " where CRTMVTDTR >= " & mAmjMin_Exercice & " and CRTMVTDTR <= " & mAmjMax_Exercice _
     & " and CRTMVTSTA = '?'"

Set rsSabX = cnsab.Execute(xSQL)
mCRT_àRevoir = rsSabX(0)
If mCRT_àRevoir > 0 Then
    Call MsgBox("Il y a " & mCRT_àRevoir & " mouvements à revoir", vbCritical, currentAction)
End If

mXls1_File = mXls1_File + 1

mCRT_File_Id = "BDF_CRT_" & mExercice & " - " & DSYS_Time & mXls1_File

Close
Call cmdSelect_SQL_9xml_Init(1, "HPD", "HPD")
'____________________________________________________________________________________________________________
newYCRTLOG0 = zYCRTLOG0
newYCRTLOG0.CRTLOGNAT = "D01"
newYCRTLOG0.CRTLOGTXT = "<AAAA = " & mExercice & ">" & "<nb mvt à revoir = " & mCRT_àRevoir & ">" _
                      & "<Fichier = " & mCRT_File_Local & ">"
'____________________________________________________________________________________________________________
Call cmdSelect_SQL_9xml_Init(3, "PFD", "PFD")
'_________________________________________________________________________
cmdPrint_Excel ("Déclaration-xml " & mExercice)
'_________________________________________________________________________
Print #1, "        </Report>"
Print #1, "</DeclarationReport>"

'_________________________________________________________________________
Print #3, "        </Report>"
Print #3, "</DeclarationReport>"

Close

'_________________________________________________________________________
Set rsSab = Nothing
Call MsgBox("Traitement terminé  ", vbInformation, currentAction)

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
'_______________________________________________________________________________________

End Sub

Private Sub cmdSelect_SQL_9xml_Init(lFile_No As Integer, lDomain As String, lReport As String)
Dim V, X As String

On Error GoTo Error_Handler


X = mId$(DSys, 1, 4) & "-" & mId$(DSys, 5, 2) & "-" & mId$(DSys, 7, 2) & "T" & Time & ".000"
'mCRT_File = paramFacturation_Path & "CRT\CRT_" & wAMJMin & " en date du " & DSys & "_" & time_Hms & ".xml"
mCRT_File_Local = "C:\Temp\" & mCRT_File_Id & ".xml"
mCRT_File_Local = Replace(mCRT_File_Local, "BDF_CRT_", "BDF_CRT_" & lReport & "_")
Open mCRT_File_Local For Output As #lFile_No
'=============================================================
Print #lFile_No, "<?xml version=" & Asc34 & "1.0" & Asc34 & " encoding=" & Asc34 & "UTF-8" & Asc34 & " standalone= " & Asc34 & "yes" & Asc34 & "?>"
Print #lFile_No, "<DeclarationReport xmlns=" & Asc34 & "http://www.onegate.eu/2010-01-01" & Asc34 & ">"
Print #lFile_No, "        <Administration creationTime =" & Asc34 & X & Asc34 & ">"
Print #lFile_No, "               <From declarerType=" & Asc34 & "SIREN" & Asc34 & ">" & socSiren & "</From>"
'Print #lFile_No, "               <From declarerType=" & Asc34 & "CIB" & Asc34 & ">" & strSocBdfE & "</From>"
Print #lFile_No, "               <To>BDF</To>"
Print #lFile_No, "               <Domain>" & lDomain & "</Domain>"
Print #lFile_No, "               <Response>"
Print #lFile_No, "                      <Email>compta@bia-paris.fr</Email>"
Print #lFile_No, "                      <Language>FR</Language>"
Print #lFile_No, "               </Response>"
Print #lFile_No, "        </Administration>"
Print #lFile_No, "        <Report date=" & Asc34 & mExercice & Asc34 & " code=" & Asc34 & lReport & Asc34 & ">"
'_________________________________________________________________________

'_________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction & " - cmdSelect_SQL_9xml_Init"
'_______________________________________________________________________________________

End Sub

Private Sub fgLog_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next


If y <= fgLog.RowHeightMin Then
Else
                
    fgLog.Col = 1: oldYCRTLOG0.CRTLOGCPT = Trim(fgLog.Text)
    fgLog.Col = 0: oldYCRTLOG0.CRTLOGID = Val(fgLog.Text)
    Call fraYCRTLOG0_Display(oldYCRTLOG0.CRTLOGID, oldYCRTLOG0.CRTLOGCPT)
End If
fgLog.LeftCol = 0


End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next


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
            Case "1": fgSelect.Col = 1: fraYCRTCPT0_Display fgSelect.Text
            Case "2c": fgSelect.Col = 1: cmdSelect_SQL_2_YCRTMVT0 fgSelect.Text
            Case "2c$": fgSelect.Col = 1: cmdSelect_SQL_2_Importation_YCRTMVT0 fgSelect.Text
       End Select
        
   End If
End If
fgSelect.LeftCol = 0
Call lstErr_AddItem(lstErr, cmdContext, "< BDF_CRT_cmdSelect_Ok"): DoEvents

End Sub

Private Sub fgYCRTLOG0_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

If X < 3900 Then
    fgYCRTLOG0.Col = 1: lblCRTLOGTXT.ForeColor = fgYCRTLOG0.CellForeColor
Else
    fgYCRTLOG0.Col = 2: lblCRTLOGTXT.ForeColor = fgYCRTLOG0.CellForeColor
End If
lblCRTLOGTXT = fgYCRTLOG0.Text
lblCRTLOGTXT.Visible = True
End Sub

Private Sub fgYSWISAB0_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xMTK As String, xBIC As String
On Error Resume Next


If y <= fgYSWISAB0.RowHeightMin Then
    Select Case fgYSWISAB0.Col
        Case 0: fgYSWISAB0_Sort1 = 0: fgYSWISAB0_Sort2 = 2: fgYSWISAB0_Sort
        Case 1:  fgYSWISAB0_Sort1 = 1: fgYSWISAB0_Sort2 = 2: fgYSWISAB0_Sort
        Case 2:  fgYSWISAB0_Sort1 = 2: fgYSWISAB0_Sort2 = 2: fgYSWISAB0_Sort
        Case 3:  fgYSWISAB0_Sort1 = 3: fgYSWISAB0_Sort2 = 3: fgYSWISAB0_SortX 3
        Case 4:  fgYSWISAB0_Sort1 = 4: fgYSWISAB0_Sort2 = 4: fgYSWISAB0_SortX 4
        Case 5:  fgYSWISAB0_Sort1 = 5: fgYSWISAB0_Sort2 = 5: fgYSWISAB0_Sort
        Case 6:  fgYSWISAB0_Sort1 = 6: fgYSWISAB0_Sort2 = 6: fgYSWISAB0_Sort
        Case 7:  fgYSWISAB0_Sort1 = 7: fgYSWISAB0_Sort2 = 7: fgYSWISAB0_Sort
        Case 8:  fgYSWISAB0_Sort1 = 8: fgYSWISAB0_Sort2 = 8: fgYSWISAB0_Sort
        Case 9:  fgYSWISAB0_Sort1 = 9: fgYSWISAB0_Sort2 = 9: fgYSWISAB0_Sort
        Case 10: fgYSWISAB0_Sort1 = 10: fgYSWISAB0_Sort2 = 10: fgYSWISAB0_Sort
        Case fgYSWISAB0_arrIndex:  fgYSWISAB0_SortX fgYSWISAB0_arrIndex
    End Select
Else
    If fgYSWISAB0.Rows > 1 Then
        Call fgYSWISAB0_Color(fgYSWISAB0_RowClick, MouseMoveUsr.BackColor, fgYSWISAB0_ColorClick)
        
        fgYSWISAB0.Col = 11: xYSWISAB0.SWISABSWID = fgYSWISAB0.Text
        fgYSWISAB0.Col = 0: xMTK = Trim(fgYSWISAB0.Text)
        fgYSWISAB0.Col = 1: xBIC = Trim(fgYSWISAB0.Text)
        fgYSWISAB0.Col = 6: xBIC = xBIC & "   swift du " & Trim(fgYSWISAB0.Text)
        
        Call fgSwift_Display(xYSWISAB0.SWISABSWID, xMTK, xBIC)
        
   End If
End If
Wait_SS 0
fgYSWISAB0.LeftCol = 0

End Sub


Private Sub Form_Activate()
Set XForm = Me

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13:
        If Not fraOD.Visible Then KeyCode = 0: cmdContext_Return
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
SSTab1.Tab = 0
blnControl = True

End Sub

Public Sub cmdSelect_Clear()

lstErr.Clear
fgSelect.Visible = False
fgDetail.Visible = False
fgLog.Visible = False
cmdSelect_Ok.BackColor = RGB(255, 255, 0)
fraYCRTCPT0.Visible = False
fraYCRTMVT0.Visible = False
fraOD.Visible = False
fraYCRTLOG0.Visible = False
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
    
    cmdPrint.Visible = False
    fraSelect_Options.Visible = False
    fraSelect_Log.Visible = False
    txtSelect_COMPTEOBL.Visible = False: lblSelect_COMPTEOBL.Visible = False
    cboSelect_CRTCPTSTA.Enabled = True
    fraSelect_2.Visible = False

    fgDetail.Left = fgDetail_Left
    fgDetail.Width = fgDetail_Width
    fgDetail.Left = fgSelect.Left + fgSelect.Width - fgDetail.Width - 200

    Select Case cmdSelect_SQL_K
        Case "1": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
                  fraSelect_Options_AMJ.Visible = False
                  txtSelect_COMPTEOBL.Visible = True
                  lblSelect_COMPTEOBL.Visible = True
                  cmdPrint.Visible = True
        Case "2c": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
                  fraSelect_Options_AMJ.Visible = True
                  fraSelect_2.Visible = True
        Case "2", "Annulation$", "2+": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
                  fraSelect_Options_AMJ.Visible = True
                  fraSelect_2.Visible = True
                  fgDetail.Left = fgSelect.Left
                  fgDetail.Width = fgSelect.Width
                   If cmdSelect_SQL_K = "Annulation$" Then
                        Call cbo_Scan("?", cboSelect_CRTMVTSTA)
                    Else
                        cmdPrint.Visible = True
                    End If
        Case "2c$": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
                  fraSelect_Options_AMJ.Visible = False
        Case "Log": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True: fraSelect_Log.Visible = True
                  fraSelect_Options_AMJ.Visible = True
                  cmdPrint.Visible = True
    End Select

End If
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


If fraYTVACOM0.Visible Then fraYTVACOM0.Visible = False: Exit Sub
If fraSwift.Visible Then fraSwift.Visible = False: Exit Sub
If fraYCRTLOG0.Visible Then fraYCRTLOG0.Visible = False: Exit Sub
If fgLog.Visible Then fgLog.Visible = False: Exit Sub
If fraYCRTCPT0.Visible Then fraYCRTCPT0.Visible = False: Exit Sub
If fraOD.Visible Then fraOD.Visible = False: Exit Sub
If fraYCRTMVT0.Visible Then fraYCRTMVT0.Visible = False: Exit Sub

If fraParam_Display.Visible Then fraParam_Display.Visible = False: Exit Sub

If txtRTF.Visible Then txtRTF.Visible = False: Exit Sub

If txtFg.Visible Then txtFg.Visible = False: Exit Sub

If fgDetail.Visible Then fgDetail.Visible = False: cmdPrint.Visible = False: Exit Sub

If fgSelect.Visible Then fgSelect.Visible = False: Exit Sub

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

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If blnfrmSAB_Dossier_DB Then frmSAB_Dossier_DB.Hide

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
Call lstErr_Clear(lstErr, cmdContext, "> BDF_CRT_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Clear
If blnfrmSAB_Dossier_DB Then frmSAB_Dossier_DB.Hide: blnfrmSAB_Dossier_DB = False

Select Case cmdSelect_SQL_K
    Case "1": cmdSelect_SQL_1
    Case "Comptes$": cmdSelect_SQL_1_Importation_YCRTCPT0
    Case "2", "Annulation$": cmdSelect_SQL_2
    Case "2+": cmdSelect_SQL_2
    Case "Mouvements$": cmdSelect_SQL_2_Importation_All
    Case "2c": cmdSelect_SQL_2c
    Case "2c$":
        Call MsgBox("Attention, incompatible avec 2$*", vbCritical, "BDF_CRT")
        cmdSelect_SQL_2c
    Case "CRT.xml": cmdSelect_SQL_9xml
    Case "CRT.xls": cmdPrint_Excel ("Déclaration-xls " & mExercice)
    Case "Clôture": cmdSelect_SQL_Clôture
    Case "Log": cmdSelect_SQL_Log
    'Case "JPL": cmdSelect_SQL_1_Importation_xls
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< BDF_CRT_cmdSelect_Ok"): DoEvents
lstErr.Height = 480
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
cmdSelect_Ok.BackColor = RGB(0, 255, 0) 'fgSelect.BackColorFixed
End Sub




Public Sub param_Init_Rubrique()
Dim Nb As Long, xSQL As String
On Error GoTo Error_Handler


'xSQL = "Delete from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID like 'CRT_%'"
'Set rsSab = cnsab.Execute(xSQL, Nb)
'_______________________________________________________________________________________________________________________________

New_YBIATAB0.BIATABID = "CRT_Mvt=Ann"

New_YBIATAB0.BIATABK1 = "000000000001": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "639430EUR100001      CP JC -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000002": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "639430EUR100001      CP CP -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000003": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "639431EUR100001      CP JC -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000004": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "639431EUR100001      CP CP -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000005": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "639432EUR100001      CP JC -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000006": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "639432EUR100001      CP CP -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000007": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "620030EUR100000      CP JC -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000008": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "620010EUR100000      CP JC -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000009": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "620070EUR100001      CP JC -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000010": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "620020EUR100000      CP JC -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000011": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "620010EUR100000      CP CP -TR      ": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000012": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "620010EUR100001      CP CP -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000013": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "620010EUR100001      CP CP -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000014": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "620040EUR100000      CP CP -TR ": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000015": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "620050EUR100001      CP CP -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000016": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "620070EUR100001      CP CP -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000017": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "690000EUR100001      CP JC *Z1 REG": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000018": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "601910EUR100001      00 GU *G1 SOG": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000019": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "608220EUR100001      CP JC -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000020": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "608220EUR100001      CP CP -TR": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000021": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "601910EUR100001      00 00 *G1 SOG": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000022": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "601910EUR100001      00 MP *G1 SOG ": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000023": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "601910EUR100001      00 00 *G1 EXT": Parametrage_New
New_YBIATAB0.BIATABK1 = "000000000024": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "601910EUR100001      00 MP *G1 EXT": Parametrage_New
'Exit Sub
'==============================
'_______________________________________________________________________________________________________________________________

New_YBIATAB0.BIATABID = "CRT_Rub_I"
New_YBIATAB0.BIATABK1 = "CA010": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID053": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "": Parametrage_New
New_YBIATAB0.BIATABK1 = "RV010": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV010": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV020": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV040": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV071": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV072": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV073": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV082": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "": Parametrage_New


New_YBIATAB0.BIATABID = "CRT_AAAA"
New_YBIATAB0.BIATABK1 = "": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "2011": Parametrage_New
'_______________________________________________________________________________________________________________________________
New_YBIATAB0.BIATABID = "CRT_LogNat"
New_YBIATAB0.BIATABK1 = "C01": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "compte ajouté (YCRTCPT0)": Parametrage_New
New_YBIATAB0.BIATABK1 = "C02": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "compte modifié (YCRTCPT0)": Parametrage_New
New_YBIATAB0.BIATABK1 = "C03": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "compte supprimé (YCRTCPT0)": Parametrage_New

New_YBIATAB0.BIATABK1 = "M01": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "mouvement ajouté (YCRTMVT0)": Parametrage_New
New_YBIATAB0.BIATABK1 = "M02": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "mouvement modifié (YCRTMVT0)": Parametrage_New
New_YBIATAB0.BIATABK1 = "M03": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "mouvement supprimé (YCRTMVT0)": Parametrage_New
New_YBIATAB0.BIATABK1 = "M10": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "importation globale 2$*(YCRTMVT0)": Parametrage_New

New_YBIATAB0.BIATABK1 = "P01": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "paramètre ajouté (YBIATAB0)": Parametrage_New
New_YBIATAB0.BIATABK1 = "P02": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "paramètre modifié (YBIATAB0)": Parametrage_New
New_YBIATAB0.BIATABK1 = "P03": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "paramètre supprimé (YBIATAB0)": Parametrage_New
New_YBIATAB0.BIATABK1 = "P10": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "traitement CRT_Mvt=Ann": Parametrage_New
'
New_YBIATAB0.BIATABK1 = "J00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "traitement journalier début": Parametrage_New
New_YBIATAB0.BIATABK1 = "J99": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "traitement journalier terminé": Parametrage_New

New_YBIATAB0.BIATABK1 = "E00": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "déclaration début du traitement": Parametrage_New
New_YBIATAB0.BIATABK1 = "E99": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "déclaration fin du traitement": Parametrage_New
'_______________________________________________________________________________________________________________________________

New_YBIATAB0.BIATABID = "CRT_Devise"
New_YBIATAB0.BIATABK1 = "EUR": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "USD": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "CHF": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "GBP": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "JPY": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "DKK": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "SEK": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "BGN": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "CZK": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "EEK": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "HUF": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "LTL": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "LVL": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "PLN": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "RON": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
New_YBIATAB0.BIATABK1 = "ZDV": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = New_YBIATAB0.BIATABK1: Parametrage_New
'_______________________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________________
New_YBIATAB0.BIATABID = "CRT_Rubrique"
' New_YBIATAB0.BIATABK1 = "SV0": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "": Parametrage_New

New_YBIATAB0.BIATABK1 = "SV010": New_YBIATAB0.BIATABK2 = "63": New_YBIATAB0.BIATABTXT = "1-Réparations": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV020": New_YBIATAB0.BIATABK2 = "63": New_YBIATAB0.BIATABTXT = "1-Courrier": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV040": New_YBIATAB0.BIATABK2 = "63": New_YBIATAB0.BIATABTXT = "1-Assurances : prime": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV040": New_YBIATAB0.BIATABK2 = "7089": New_YBIATAB0.BIATABTXT = "1-Assurances : primes": Parametrage_New

New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "708": New_YBIATAB0.BIATABTXT = "1-Produits et charges sur prestations de services financiers": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "608": New_YBIATAB0.BIATABTXT = "1-Produits et charges sur prestations de services financiers": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "7071": New_YBIATAB0.BIATABTXT = "1-Produits et charges sur prestations de services financiers": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "7072": New_YBIATAB0.BIATABTXT = "1-Produits et charges sur prestations de services financiers": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "633": New_YBIATAB0.BIATABTXT = "1-Commissions reçues et versées sur opérations de trésorerie, opérations interbancaires et avec la clientèle": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "6029": New_YBIATAB0.BIATABTXT = "1-Commissions reçues et versées sur opérations de trésorerie, opérations interbancaires et avec la clientèle": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "6019": New_YBIATAB0.BIATABTXT = "1-Commissions reçues et versées sur opérations de trésorerie, opérations interbancaires et avec la clientèle": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "7029": New_YBIATAB0.BIATABTXT = "1-Commissions reçues et versées sur opérations de trésorerie, opérations interbancaires et avec la clientèle": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "7019": New_YBIATAB0.BIATABTXT = "1-Commissions reçues et versées sur opérations de trésorerie, opérations interbancaires et avec la clientèle": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "7039": New_YBIATAB0.BIATABTXT = "1-Commissions reçues et versées sur opérations sur titres": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "70739": New_YBIATAB0.BIATABTXT = "1-Commissions reçues et versées sur opérations sur titres": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "60739": New_YBIATAB0.BIATABTXT = "1-Commissions reçues et versées sur opérations sur titres": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "6039": New_YBIATAB0.BIATABTXT = "1-Commissions reçues et versées sur opérations sur titres": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "7069": New_YBIATAB0.BIATABTXT = "1-Commissions reçues et versées sur opérations de change et sur instruments financiers à terme": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "70749": New_YBIATAB0.BIATABTXT = "1-Commissions reçues et versées sur opérations de change et sur instruments financiers à terme": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "6069": New_YBIATAB0.BIATABTXT = "1-Commissions reçues et versées sur opérations de change et sur instruments financiers à terme": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV051": New_YBIATAB0.BIATABK2 = "60749": New_YBIATAB0.BIATABTXT = "1-Commissions reçues et versées sur opérations de change et sur instruments financiers à terme": Parametrage_New

New_YBIATAB0.BIATABK1 = "SV060": New_YBIATAB0.BIATABK2 = "63": New_YBIATAB0.BIATABTXT = "1-Redevances sur brevets, échanges de savoir-faire,...": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV060": New_YBIATAB0.BIATABK2 = "7479": New_YBIATAB0.BIATABTXT = "1-Redevances sur brevets, échanges de savoir-faire,...": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV070": New_YBIATAB0.BIATABK2 = "7083": New_YBIATAB0.BIATABTXT = "1-Etudes, recherches et assistance technique": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV070": New_YBIATAB0.BIATABK2 = "63": New_YBIATAB0.BIATABTXT = "1-Etudes, recherches et assistance technique": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV071": New_YBIATAB0.BIATABK2 = "63": New_YBIATAB0.BIATABTXT = "1-Télécommunications": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV072": New_YBIATAB0.BIATABK2 = "63": New_YBIATAB0.BIATABTXT = "1-Services informatiques": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV072": New_YBIATAB0.BIATABK2 = "7479": New_YBIATAB0.BIATABTXT = "1-Services informatiques": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV073": New_YBIATAB0.BIATABK2 = "63": New_YBIATAB0.BIATABTXT = "1-Abonnements": Parametrage_New
New_YBIATAB0.BIATABK1 = "SV082": New_YBIATAB0.BIATABK2 = "63": New_YBIATAB0.BIATABTXT = "1-Publicité": Parametrage_New

New_YBIATAB0.BIATABK1 = "RV010": New_YBIATAB0.BIATABK2 = "611": New_YBIATAB0.BIATABTXT = "1-Transferts de salaires": Parametrage_New
New_YBIATAB0.BIATABK1 = "RV010": New_YBIATAB0.BIATABK2 = "613": New_YBIATAB0.BIATABTXT = "1-Transferts de salaires": Parametrage_New
New_YBIATAB0.BIATABK1 = "RV021": New_YBIATAB0.BIATABK2 = "6051": New_YBIATAB0.BIATABTXT = "1-Revenus d'investissements directs": Parametrage_New
New_YBIATAB0.BIATABK1 = "RV021": New_YBIATAB0.BIATABK2 = "6052": New_YBIATAB0.BIATABTXT = "1-Revenus d'investissements directs": Parametrage_New
New_YBIATAB0.BIATABK1 = "RV021": New_YBIATAB0.BIATABK2 = "7051": New_YBIATAB0.BIATABTXT = "1-Revenus d'investissements directs": Parametrage_New
New_YBIATAB0.BIATABK1 = "RV021": New_YBIATAB0.BIATABK2 = "7052": New_YBIATAB0.BIATABTXT = "1-Revenus d'investissements directs": Parametrage_New
New_YBIATAB0.BIATABK1 = "RV021": New_YBIATAB0.BIATABK2 = "7053": New_YBIATAB0.BIATABTXT = "1-Revenus d'investissements directs": Parametrage_New
New_YBIATAB0.BIATABK1 = "RV031": New_YBIATAB0.BIATABK2 = "62": New_YBIATAB0.BIATABTXT = "1-Transferts unilatéraux au profit d'administration publiques non résidentes": Parametrage_New
New_YBIATAB0.BIATABK1 = "RV031": New_YBIATAB0.BIATABK2 = "69": New_YBIATAB0.BIATABTXT = "1-Transferts unilatéraux au profit d'administration publiques non résidentes": Parametrage_New
New_YBIATAB0.BIATABK1 = "RV032": New_YBIATAB0.BIATABK2 = "63": New_YBIATAB0.BIATABTXT = "1-Autres transferts unilatéraux, Assurances : indemnités": Parametrage_New
New_YBIATAB0.BIATABK1 = "RV032": New_YBIATAB0.BIATABK2 = "7479": New_YBIATAB0.BIATABTXT = "1-Autres transferts unilatéraux, Assurances : indemnités": Parametrage_New
New_YBIATAB0.BIATABK1 = "CA022": New_YBIATAB0.BIATABK2 = "675": New_YBIATAB0.BIATABTXT = "1-Pertes ou profits sur créances ou engagements des intermédiaires financiers": Parametrage_New
New_YBIATAB0.BIATABK1 = "CA022": New_YBIATAB0.BIATABK2 = "676": New_YBIATAB0.BIATABTXT = "1-Pertes ou profits sur créances ou engagements des intermédiaires financiers": Parametrage_New

New_YBIATAB0.BIATABK1 = "CA010": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "2-Achats et ventes de brevets": Parametrage_New

New_YBIATAB0.BIATABK1 = "ID051": New_YBIATAB0.BIATABK2 = "4111": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes non cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID051": New_YBIATAB0.BIATABK2 = "4112": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes non cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID051": New_YBIATAB0.BIATABK2 = "41131": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes non cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID051": New_YBIATAB0.BIATABK2 = "41139": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes non cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID051": New_YBIATAB0.BIATABK2 = "4121": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes non cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID051": New_YBIATAB0.BIATABK2 = "4122": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes non cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID051": New_YBIATAB0.BIATABK2 = "41231": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes non cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID051": New_YBIATAB0.BIATABK2 = "41239": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes non cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID051": New_YBIATAB0.BIATABK2 = "415": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes non cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID051": New_YBIATAB0.BIATABK2 = "42": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes non cotées": Parametrage_New

New_YBIATAB0.BIATABK1 = "ID052": New_YBIATAB0.BIATABK2 = "4111": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID052": New_YBIATAB0.BIATABK2 = "4112": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID052": New_YBIATAB0.BIATABK2 = "41131": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID052": New_YBIATAB0.BIATABK2 = "41139": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID052": New_YBIATAB0.BIATABK2 = "4121": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID052": New_YBIATAB0.BIATABK2 = "4122": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID052": New_YBIATAB0.BIATABK2 = "41231": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID052": New_YBIATAB0.BIATABK2 = "41239": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes cotées": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID052": New_YBIATAB0.BIATABK2 = "415": New_YBIATAB0.BIATABTXT = "2-Investissements directs en capital social dans les entreprises non résidentes cotées": Parametrage_New

New_YBIATAB0.BIATABK1 = "ID056": New_YBIATAB0.BIATABK2 = "5611": New_YBIATAB0.BIATABTXT = "2-Investissements directs des non-résidents dans le capital social des entités résidentes": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID056": New_YBIATAB0.BIATABK2 = "5619": New_YBIATAB0.BIATABTXT = "2-Investissements directs des non-résidents dans le capital social des entités résidentes": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID056": New_YBIATAB0.BIATABK2 = "571": New_YBIATAB0.BIATABTXT = "2-Investissements directs des non-résidents dans le capital social des entités résidentes": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID056": New_YBIATAB0.BIATABK2 = "572": New_YBIATAB0.BIATABTXT = "2-Investissements directs des non-résidents dans le capital social des entités résidentes": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID056": New_YBIATAB0.BIATABK2 = "573": New_YBIATAB0.BIATABTXT = "2-Investissements directs des non-résidents dans le capital social des entités résidentes": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID056": New_YBIATAB0.BIATABK2 = "574": New_YBIATAB0.BIATABTXT = "2-Investissements directs des non-résidents dans le capital social des entités résidentes": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID056": New_YBIATAB0.BIATABK2 = "578": New_YBIATAB0.BIATABTXT = "2-Investissements directs des non-résidents dans le capital social des entités résidentes": Parametrage_New

New_YBIATAB0.BIATABK1 = "ID053": New_YBIATAB0.BIATABK2 = "432": New_YBIATAB0.BIATABTXT = "2-Investissements immobiliers des résidents": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID053": New_YBIATAB0.BIATABK2 = "442": New_YBIATAB0.BIATABTXT = "2-Investissements immobiliers des résidents": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID053": New_YBIATAB0.BIATABK2 = "452": New_YBIATAB0.BIATABTXT = "2-Investissements immobiliers des résidents": Parametrage_New

New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "4011": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "4019": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "402": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "416": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20111": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20112": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20119": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20211": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20212": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20213": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20219": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20311": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20312": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20313": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20314": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20315": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20316": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20317": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20318": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20319": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20411": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20412": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20419": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20511": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20512": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20513": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20514": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20515": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20516": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20517": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20518": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "20519": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2052": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2061": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "221": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2311": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2312": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2411": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2412": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2511": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2611": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "291": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2017": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2027": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2037": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2047": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2057": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2067": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "227": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2317": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2417": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "25171": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID054": New_YBIATAB0.BIATABK2 = "2617": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New

New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "4011": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "4019": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "402": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "416": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20111": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20112": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20119": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20211": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20212": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20213": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20219": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20311": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20312": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20313": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20314": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20315": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20316": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20317": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20318": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20319": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20411": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20412": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20419": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20511": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20512": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20513": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20514": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20515": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20516": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20517": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20518": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "20519": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2052": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2061": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "221": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2311": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2312": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2411": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2412": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2511": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2611": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "291": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2017": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2027": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2037": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2047": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2057": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2067": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "227": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2317": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2417": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "25171": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID055": New_YBIATAB0.BIATABK2 = "2617": New_YBIATAB0.BIATABTXT = "2-Prêts à long terme des résidents à tout non rédident du même groupe": Parametrage_New

New_YBIATAB0.BIATABK1 = "ID058": New_YBIATAB0.BIATABK2 = "5411": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID058": New_YBIATAB0.BIATABK2 = "5412": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID058": New_YBIATAB0.BIATABK2 = "5419": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID058": New_YBIATAB0.BIATABK2 = "5421": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID058": New_YBIATAB0.BIATABK2 = "5422": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID058": New_YBIATAB0.BIATABK2 = "2321": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID058": New_YBIATAB0.BIATABK2 = "2322": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID058": New_YBIATAB0.BIATABK2 = "2431": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID058": New_YBIATAB0.BIATABK2 = "2432": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID058": New_YBIATAB0.BIATABK2 = "2621": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID058": New_YBIATAB0.BIATABK2 = "2327": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID058": New_YBIATAB0.BIATABK2 = "2437": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID058": New_YBIATAB0.BIATABK2 = "25172": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID058": New_YBIATAB0.BIATABK2 = "2627": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New

New_YBIATAB0.BIATABK1 = "ID059": New_YBIATAB0.BIATABK2 = "5411": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID059": New_YBIATAB0.BIATABK2 = "5412": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID059": New_YBIATAB0.BIATABK2 = "5419": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID059": New_YBIATAB0.BIATABK2 = "5421": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID059": New_YBIATAB0.BIATABK2 = "5422": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID059": New_YBIATAB0.BIATABK2 = "2321": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID059": New_YBIATAB0.BIATABK2 = "2322": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID059": New_YBIATAB0.BIATABK2 = "2431": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID059": New_YBIATAB0.BIATABK2 = "2432": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID059": New_YBIATAB0.BIATABK2 = "2621": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID059": New_YBIATAB0.BIATABK2 = "2327": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID059": New_YBIATAB0.BIATABK2 = "2437": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID059": New_YBIATAB0.BIATABK2 = "25172": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New
New_YBIATAB0.BIATABK1 = "ID059": New_YBIATAB0.BIATABK2 = "2627": New_YBIATAB0.BIATABTXT = "2-Prêts à court terme des résidents à tout non rédident du même groupe": Parametrage_New

New_YBIATAB0.BIATABK1 = "OA110": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Actif-Instruments conditionnels": Parametrage_New
New_YBIATAB0.BIATABK1 = "OA111": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Réevaluations-Instruments conditionnels à l'actif": Parametrage_New
New_YBIATAB0.BIATABK1 = "OA120": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Appel de marge sur instruments conditionnels à l'actif": Parametrage_New
New_YBIATAB0.BIATABK1 = "OP110": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Passif-Instruments conditionnels": Parametrage_New
New_YBIATAB0.BIATABK1 = "OP111": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Réevaluations-Instruments conditionnels au passif": Parametrage_New
New_YBIATAB0.BIATABK1 = "OP120": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Appel de marge sur instruments conditionnels au passif": Parametrage_New

New_YBIATAB0.BIATABK1 = "OA210": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Actif-Instruments conditionnels": Parametrage_New
New_YBIATAB0.BIATABK1 = "OA211": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Réevaluations-Instruments conditionnels à l'actif": Parametrage_New
New_YBIATAB0.BIATABK1 = "OP210": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Passif-Instruments conditionnels": Parametrage_New
New_YBIATAB0.BIATABK1 = "OP211": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Réevaluations-Instruments conditionnels au passif": Parametrage_New

New_YBIATAB0.BIATABK1 = "SA110": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Actif-SWAP": Parametrage_New
New_YBIATAB0.BIATABK1 = "SA111": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Réevaluations-SWAP à l'actif": Parametrage_New
New_YBIATAB0.BIATABK1 = "SA120": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Appel de marge sur SWAP à l'actif": Parametrage_New
New_YBIATAB0.BIATABK1 = "SP110": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Passif-SWAP": Parametrage_New
New_YBIATAB0.BIATABK1 = "SP111": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Réevaluations-SWAP au passif": Parametrage_New
New_YBIATAB0.BIATABK1 = "SP120": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Appel de marge sur SWAP au passif": Parametrage_New


New_YBIATAB0.BIATABK1 = "SA210": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Actif-SWAP": Parametrage_New
New_YBIATAB0.BIATABK1 = "SA211": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Réevaluations-SWAP à l'actif": Parametrage_New
New_YBIATAB0.BIATABK1 = "SP210": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Passif-SWAP": Parametrage_New
New_YBIATAB0.BIATABK1 = "SP211": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Réevaluations-SWAP au passif": Parametrage_New

New_YBIATAB0.BIATABK1 = "FA120": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Appel de marge sur Future à l'actif": Parametrage_New
New_YBIATAB0.BIATABK1 = "FA121": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Réevaluations-Future à l'actif": Parametrage_New
New_YBIATAB0.BIATABK1 = "FP121": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Réevaluations-Future au passif": Parametrage_New

New_YBIATAB0.BIATABK1 = "FA210": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Actif-Contrats à terme de gré à gré": Parametrage_New
New_YBIATAB0.BIATABK1 = "FA211": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Réevaluations-Contrats à terme de gré à gré à l'actif": Parametrage_New
New_YBIATAB0.BIATABK1 = "FP210": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Passif-Contrats à terme de gré à gré": Parametrage_New
New_YBIATAB0.BIATABK1 = "FP211": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Réevaluations-Contrats à terme de gré à gré au passif": Parametrage_New

New_YBIATAB0.BIATABK1 = "DA210": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Actif-Tous PFD": Parametrage_New
New_YBIATAB0.BIATABK1 = "DA211": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Réevaluations-Tous PFD à l'actif": Parametrage_New
New_YBIATAB0.BIATABK1 = "DP210": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Transactions-Passif-Tous PFD": Parametrage_New
New_YBIATAB0.BIATABK1 = "DP211": New_YBIATAB0.BIATABK2 = "": New_YBIATAB0.BIATABTXT = "3-Réevaluations-Tous PFD au passif": Parametrage_New

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:


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
    
newYCRTLOG0 = zYCRTLOG0
newYCRTLOG0.CRTLOGNAT = "P03"
newYCRTLOG0.CRTLOGTXT = "<BIATABID = " & Trim(Old_YBIATAB0.BIATABID) & ">" & "<BIATABK1 = " & Trim(Old_YBIATAB0.BIATABK1) & ">" _
                      & "<BIATABK2 = " & Trim(Old_YBIATAB0.BIATABK2) & ">" & "<BIATABTXT = " & Trim(Old_YBIATAB0.BIATABTXT) & ">"
V = sqlYCRTLOG0_Insert(newYCRTLOG0)
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

newYCRTLOG0 = zYCRTLOG0
newYCRTLOG0.CRTLOGNAT = "P01"
newYCRTLOG0.CRTLOGTXT = "<BIATABID = " & Trim(New_YBIATAB0.BIATABID) & ">" & "<BIATABK1 = " & Trim(New_YBIATAB0.BIATABK1) & ">" _
                      & "<BIATABK2 = " & Trim(New_YBIATAB0.BIATABK2) & ">" & "<BIATABTXT = " & Trim(New_YBIATAB0.BIATABTXT) & ">"
V = sqlYCRTLOG0_Insert(newYCRTLOG0)
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
If Not IsNull(V) Then GoTo Error_MsgBox


newYCRTLOG0 = zYCRTLOG0
newYCRTLOG0.CRTLOGNAT = "P02"
newYCRTLOG0.CRTLOGTXT = "<BIATABID = " & Trim(Old_YBIATAB0.BIATABID) & " | " & Trim(New_YBIATAB0.BIATABID) & ">" _
                      & "<BIATABK1 = " & Trim(Old_YBIATAB0.BIATABK1) & " | " & Trim(New_YBIATAB0.BIATABK1) & ">" _
                      & "<BIATABK2 = " & Trim(Old_YBIATAB0.BIATABK2) & " | " & Trim(New_YBIATAB0.BIATABK2) & ">" _
                      & "<BIATABTXT = " & Trim(Old_YBIATAB0.BIATABTXT) & " | " & Trim(New_YBIATAB0.BIATABTXT) & ">"
V = sqlYCRTLOG0_Insert(newYCRTLOG0)
If Not IsNull(V) Then GoTo Error_MsgBox

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




Public Sub cmdPrint_Excel_YCRTCPT0_Rubriques()
Dim xSQL As String, X As String, K As Long
Dim mBIATABID As String
On Error GoTo Error_Handler

'===================================================================================

wsExcel.Name = "Rubriques"

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignLeft
    .WrapText = False ' True
    .Font.Size = 9
    .Font.Name = "Courier New"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 85

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14BDF_CRT, arrêté au " & dateImp10(wAMJMin) _
                                 & vbCr

wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

mXls1_Row = 0

wsExcel.Columns(1).ColumnWidth = 12: wsExcel.Cells(1, 1) = "Table "
wsExcel.Columns(2).ColumnWidth = 12: wsExcel.Cells(1, 2) = "Code"
wsExcel.Columns(3).ColumnWidth = 12: wsExcel.Cells(1, 3) = "Valeur"
wsExcel.Columns(4).ColumnWidth = 110: wsExcel.Cells(1, 4) = "Libellé"

For K = 1 To 4
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = vbWhite
Next

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID like 'CRT%'" _
     & " order by BIATABID , BIATABK1 , BIATABK2 "
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    If mBIATABID <> rsSab("BIATABID") Then
        mBIATABID = rsSab("BIATABID")
        mXls1_Row = mXls1_Row + 1
    End If
    mXls1_Row = mXls1_Row + 1
    wsExcel.Cells(mXls1_Row, 1) = mBIATABID
    wsExcel.Cells(mXls1_Row, 2) = Trim(rsSab("BIATABK1"))
    wsExcel.Cells(mXls1_Row, 3) = Trim(rsSab("BIATABK2"))
    wsExcel.Cells(mXls1_Row, 4) = Trim(mId$(rsSab("BIATABTXT"), 1, 99))
    rsSab.MoveNext
Loop


Exit Sub
'======================================================================================================

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:


End Sub

Public Sub cmdPrint_Excel_Déclaration()
Dim xSQL As String, X As String, K As Long, K2 As Long
Dim kNomenclature As Integer, mCRTMVTRUB As String, mCRTMVTDEV As String, mCRTMVTCLIP As String
Dim curX As Currency, curDB As Currency, curCR As Currency, kCRTMVTDEV As Integer
On Error GoTo Error_Handler

'===================================================================================

'wsExcel.Name = "Rubriques"
Call rsZBASTAB0_Pays(arrPays(), arrPays_Nb)

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignRight
    .WrapText = False ' True
    .Font.Size = 9
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 80

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14BDF_CRT Déclaration " & mExercice & ", arrêté au " & dateImp10(wAMJMin) _
                                 & vbCr

wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

mXls1_Row = 1
mXls1_Cols = 3 + 2 * arrCRT_Devise_Nb
kNomenclature = 1

wsExcel.Columns(1).ColumnWidth = 12: wsExcel.Cells(1, 1) = "Rubrique "
wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(2).ColumnWidth = 12: wsExcel.Cells(1, 2) = "Pays"
wsExcel.Columns(2).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Columns(3).ColumnWidth = 60: wsExcel.Cells(1, 3) = arrNomenclature(1)
wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignLeft

For K = 1 To arrCRT_Devise_Nb
    K2 = 2 + K * 2
    wsExcel.Columns(K2).ColumnWidth = 15: wsExcel.Cells(1, K2) = "débit " & arrCRT_Devise(K)
    wsExcel.Columns(K2).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"
    wsExcel.Columns(K2 + 1).ColumnWidth = 15: wsExcel.Cells(1, K2 + 1) = "crédit " & arrCRT_Devise(K)
    wsExcel.Columns(K2 + 1).NumberFormat = "[Blue]### ### ### ###;[Red]-### ### ### ###"
Next K

For K = 1 To mXls1_Cols
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = vbWhite
Next



xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID like 'CRT_Rubrique'" _
     & " order by substring(BIATABTXT,1,1) , BIATABK1  "
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    X = Trim(rsSab("BIATABTXT"))
    Call lstErr_ChangeLastItem(lstErr, cmdContext, Trim(rsSab("BIATABK1"))): DoEvents

    If kNomenclature <> mId$(X, 1, 1) Then
        kNomenclature = mId$(X, 1, 1)
        mXls1_Row = mXls1_Row + 1
        For K = 1 To mXls1_Cols
            wsExcel.Cells(mXls1_Row, K) = wsExcel.Cells(1, K)
            wsExcel.Cells(mXls1_Row, K).Interior.Color = wsExcel.Cells(1, K).Interior.Color
            wsExcel.Cells(mXls1_Row, K).Font.Color = wsExcel.Cells(1, K).Font.Color
        Next

        wsExcel.Cells(mXls1_Row, 3) = arrNomenclature(kNomenclature)
    End If
    
    
     If mCRTMVTRUB <> Trim(rsSab("BIATABK1")) Then
        mCRTMVTRUB = Trim(rsSab("BIATABK1"))
        mXls1_Row = mXls1_Row + 1
   
        wsExcel.Cells(mXls1_Row, 1) = mCRTMVTRUB
        wsExcel.Cells(mXls1_Row, 2) = ""
        wsExcel.Cells(mXls1_Row, 3) = mId$(rsSab("BIATABTXT"), 1, 72)
        'wsExcel.Cells(mXls1_Row, 3).Font.Size = 8
        For K = 1 To mXls1_Cols
            wsExcel.Cells(mXls1_Row, K).Interior.Color = mColor_G1
        Next
'____________________________________________________________________________________

        mCRTMVTCLIP = ""
        mCRTMVTDEV = "": kCRTMVTDEV = 0
        curDB = 0: curCR = 0
        
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCRTMVT0 " _
             & " where CRTMVTDTR >= " & mAmjMin_Exercice & " and CRTMVTDTR <= " & mAmjMax_Exercice _
             & " and CRTMVTRUB = '" & mCRTMVTRUB & "' and CRTMVTSTA = ' '" _
             & " order by CRTMVTCLIP , CRTMVTDEV  "
        Set rsSabX = cnsab.Execute(xSQL)
        
        Do While Not rsSabX.EOF
        
             If mCRTMVTCLIP <> rsSabX("CRTMVTCLIP") Then
                Call cmdPrint_Excel_Déclaration_Detail("D", kNomenclature, kCRTMVTDEV, mCRTMVTRUB, mCRTMVTDEV, mCRTMVTCLIP, curDB, curCR)
                'If kCRTMVTDEV <> 0 Then
                    'wsExcel.Cells(mXls1_Row, 2 + 2 * kCRTMVTDEV) = Round(curDB, 0)
                    'wsExcel.Cells(mXls1_Row, 2 + 2 * kCRTMVTDEV + 1) = Round(curCR, 0)
                'End If
                mCRTMVTDEV = "": kCRTMVTDEV = 0
                curDB = 0: curCR = 0
                mCRTMVTCLIP = rsSabX("CRTMVTCLIP")
                mXls1_Row = mXls1_Row + 1
                wsExcel.Cells(mXls1_Row, 1) = mCRTMVTRUB
                wsExcel.Cells(mXls1_Row, 2) = mCRTMVTCLIP
                For K = 0 To arrPays_Nb
                    If mCRTMVTCLIP = arrPays(K).Id Then
                        wsExcel.Cells(mXls1_Row, 3) = arrPays(K).Nom
                        Exit For
                    End If
                Next K
           End If
            
             If mCRTMVTDEV <> rsSabX("CRTMVTDEV") Then
                Call cmdPrint_Excel_Déclaration_Detail("D", kNomenclature, kCRTMVTDEV, mCRTMVTRUB, mCRTMVTDEV, mCRTMVTCLIP, curDB, curCR)
                'If kCRTMVTDEV <> 0 Then
                '    wsExcel.Cells(mXls1_Row, 2 + 2 * kCRTMVTDEV) = Round(curDB, 0)
                '    wsExcel.Cells(mXls1_Row, 2 + 2 * kCRTMVTDEV + 1) = Round(curCR, 0)
                'End If
                mCRTMVTDEV = rsSabX("CRTMVTDEV")
                curDB = 0: curCR = 0
                kCRTMVTDEV = arrCRT_Devise_Nb
                For K = 1 To arrCRT_Devise_Nb
                    If mCRTMVTDEV = arrCRT_Devise(K) Then kCRTMVTDEV = K: Exit For
                Next K
                Call lstErr_ChangeLastItem(lstErr, cmdContext, mCRTMVTRUB & " " & mCRTMVTCLIP & " " & mCRTMVTDEV): DoEvents
            End If
            
            curX = rsSabX("CRTMVTMTE")
            If curX < 0 Then
                If mId$(rsSabX("CRTMVTCPT"), 1, 1) = "7" Then
                    curCR = curCR + curX
                Else
                    curDB = curDB + curX
                End If
            Else
                If mId$(rsSabX("CRTMVTCPT"), 1, 1) = "6" Then
                    curDB = curDB + curX
                Else
                    curCR = curCR + curX
                End If
            End If
            rsSabX.MoveNext
        Loop
        Call cmdPrint_Excel_Déclaration_Detail("D", kNomenclature, kCRTMVTDEV, mCRTMVTRUB, mCRTMVTDEV, mCRTMVTCLIP, curDB, curCR)
        
        'If kCRTMVTDEV <> 0 Then
        '    wsExcel.Cells(mXls1_Row, 2 + 2 * kCRTMVTDEV) = Round(curDB, 0)
        '    wsExcel.Cells(mXls1_Row, 2 + 2 * kCRTMVTDEV + 1) = Round(curCR, 0)
        'End If
        
'____________________________________________________________________________________

    End If
    rsSab.MoveNext
Loop

Call cmdPrint_Excel_Déclaration_Detail("F", kNomenclature, kCRTMVTDEV, mCRTMVTRUB, mCRTMVTDEV, mCRTMVTCLIP, curDB, curCR)



Exit Sub
'======================================================================================================

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:


End Sub


Private Sub lstParam_Id_Click()

Select Case Trim(lstParam_Id.Text)
    Case "SAB_Pays_Exclus": lstParam_Load_Pays_Exclus ("Display")
                            lstParam_K.Enabled = False
    Case "SAB_Clients_Exclus": lstParam_Load_Clients_Exclus ("Display")
                            lstParam_K.Enabled = False
    Case Else

        Old_YBIATAB0.BIATABID = Trim(lstParam_Id.Text)
        lstParam_K.Enabled = True
        lstParam_Load Old_YBIATAB0.BIATABID
End Select

End Sub

Public Sub lstParam_Load(lBIATABID As String)
Dim xSQL As String

fraParam_Display.Visible = False
lstParam_K.Clear
lstParam_K.AddItem "Ajouter un enregistrement"

xSQL = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = '" & lBIATABID & "'" _
     & " order by BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
        
    lstParam_K.AddItem rsSab("BIATABK1") & " | " & rsSab("BIATABK2") & " | " & Trim(rsSab("BIATABTXT"))
    rsSab.MoveNext
Loop

End Sub


Private Sub lstParam_K_Click()
Dim xSQL As String, X As String
Dim K1 As Integer, K2 As Integer
'_______________________________________________
fraParam_Display.Visible = False
'_______________________________________________
txtParam_K2.Enabled = False
cmdParam_Add.Visible = False
cmdParam_Delete.Visible = False
cmdParam_Update.Visible = False
fraParam_CRT_Mvt_Ann.Visible = False

cboParam_Nomenclature.Visible = False: lblParam_Nomenclature.Visible = False

Old_YBIATAB0.BIATABK1 = ""
Old_YBIATAB0.BIATABK2 = ""

X = Trim(lstParam_K.Text)
K1 = InStr(1, X, "|")
If K1 > 0 Then
    Old_YBIATAB0.BIATABK1 = Trim(mId$(X, 1, K1 - 1))
    K1 = K1 + 1
    K2 = InStr(K1, X, "|")
    If K2 > 0 Then
        Old_YBIATAB0.BIATABK2 = Trim(mId$(X, K1, K2 - K1 - 1))
    End If
End If

xSQL = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = '" & Old_YBIATAB0.BIATABID & "'" _
     & " and BIATABK1 = '" & Old_YBIATAB0.BIATABK1 & "' and BIATABK2 = '" & Old_YBIATAB0.BIATABK2 & "'"

Set rsSab = cnsab.Execute(xSQL)

txtParam_K1 = Trim(Old_YBIATAB0.BIATABK1)
txtParam_K2 = Trim(Old_YBIATAB0.BIATABK2)


If Not rsSab.EOF Then
    Old_YBIATAB0.BIATABTXT = rsSab("BIATABTXT")
Else
    Old_YBIATAB0.BIATABTXT = ""
    
End If

    
Select Case Trim(Old_YBIATAB0.BIATABID)
    Case "CRT_Devise"
        cmdParam_Add.Visible = arrHab(18)
        cmdParam_Delete.Visible = arrHab(18)
        cmdParam_Update.Visible = arrHab(18)
        txtParam_Txt = Trim(Old_YBIATAB0.BIATABTXT)
    Case "CRT_Rubrique"
        cboParam_Nomenclature.Visible = True
        
        Call cbo_Scan(mId$(Old_YBIATAB0.BIATABTXT, 1, 1), cboParam_Nomenclature)
        If cboParam_Nomenclature.ListIndex < 0 Then cboParam_Nomenclature.ListIndex = 0
        txtParam_Txt = Trim(mId$(Old_YBIATAB0.BIATABTXT, 3, 125))
        txtParam_K2.Enabled = arrHab(18)
        cmdParam_Add.Visible = arrHab(18)
        cmdParam_Delete.Visible = arrHab(18)
        cmdParam_Update.Visible = arrHab(16)
    Case "CRT_LogNat"
        cmdParam_Add.Visible = arrHab(19)
        cmdParam_Delete.Visible = arrHab(19)
        cmdParam_Update.Visible = arrHab(19)
        txtParam_Txt = Trim(Old_YBIATAB0.BIATABTXT)
    Case "CRT_Rub_I"
        cmdParam_Add.Visible = arrHab(19)
        cmdParam_Delete.Visible = arrHab(19)
        cmdParam_Update.Visible = arrHab(19)
        txtParam_Txt = Trim(Old_YBIATAB0.BIATABTXT)
    Case "CRT_Mvt=Ann"
        cmdParam_Add.Visible = arrHab(18)
        cmdParam_Delete.Visible = arrHab(18)
        cmdParam_Update.Visible = arrHab(18)
        
        X = Old_YBIATAB0.BIATABTXT & Space$(50)
        txtParam_CRTMVTCPT = Trim(mId$(X, 1, 20))
        txtParam_CRTMVTSER = Trim(mId$(X, 22, 2))
        txtParam_CRTMVTSSE = Trim(mId$(X, 25, 2))
        txtParam_CRTMVTOPE = Trim(mId$(X, 28, 3))
        txtParam_CRTMVTEVE = Trim(mId$(X, 32, 3))
        If txtParam_CRTMVTCPT <> "" Then
            cmdParam_CRT_Mvt_Ann.Visible = arrHab(15) And arrHab(18)
        Else
            cmdParam_CRT_Mvt_Ann.Visible = False
        End If
        fraParam_CRT_Mvt_Ann.Visible = True
End Select
    
fraParam_Display.Visible = True

End Sub


Private Sub optCRTMVTMTD_OD_Cr_Click()
txtCRTMVTMTE_OD.ForeColor = vbBlue
txtCRTMVTMTD_OD.ForeColor = vbBlue

End Sub

Private Sub optCRTMVTMTD_OD_Db_Click()
txtCRTMVTMTE_OD.ForeColor = vbRed
txtCRTMVTMTD_OD.ForeColor = vbRed

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
    cmdPrint.Visible = True
Else
    cmdPrint.Visible = False
End If


End Sub

Private Sub SSTab2_GotFocus()
If SSTab2.Tab = 1 Then
    If Not fgYSWISAB0.Visible Then Call fgYSWISAB0_Display(oldYBIAMVTH.MOUVEMOPE, oldYBIAMVTH.MOUVEMNUM)
End If
End Sub


Private Sub txtCRTMVTCPT_OD_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub



Private Sub txtCRTMVTDTR_OD_Change()
fraOD_Display_Cours
End Sub


Private Sub txtCRTMVTMTD_OD_Change()
Call fraOD_Display_CRTMVTMTE

End Sub

Public Sub fraOD_Display_CRTMVTMTE()
Dim xCur1 As Currency, xCur2 As Currency, xDbl As Double
xCur1 = num_CDec(txtCRTMVTMTD_OD)
If mCRTMVTTAUX_OD = 0 Then
    txtCRTMVTMTE_OD = ""
Else
    xCur2 = Round(xCur1 / mCRTMVTTAUX_OD, 2)
    txtCRTMVTMTE_OD = Format(xCur2, "### ### ### ###.###")
    txtCRTMVTMTE_OD.BackColor = vbYellow
End If
End Sub

Private Sub txtCRTMVTMTD_OD_KeyPress(KeyAscii As Integer)
Call num_Montant(KeyAscii, txtCRTMVTMTD_OD)

End Sub



Private Sub txtParam_CRTMVTCPT_Change()
cmdParam_CRT_Mvt_Ann.Visible = False
End Sub

Private Sub txtParam_CRTMVTCPT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtParam_CRTMVTEVE_Change()
cmdParam_CRT_Mvt_Ann.Visible = False
End Sub

Private Sub txtParam_CRTMVTEVE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtParam_CRTMVTOPE_Change()
cmdParam_CRT_Mvt_Ann.Visible = False
End Sub

Private Sub txtParam_CRTMVTOPE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtParam_CRTMVTSER_Change()
cmdParam_CRT_Mvt_Ann.Visible = False
End Sub

Private Sub txtParam_CRTMVTSER_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtParam_CRTMVTSSE_Change()
cmdParam_CRT_Mvt_Ann.Visible = False
End Sub

Private Sub txtParam_CRTMVTSSE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtParam_K1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtParam_K2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtSelect_AmjMax_Change()
cmdSelect_Clear

End Sub


Private Sub txtSelect_AmjMax_Click()
cmdSelect_Clear
End Sub

Private Sub txtSelect_AmjMin_Change()
cmdSelect_Clear

End Sub


Private Sub txtSelect_AmjMin_Click()
cmdSelect_Clear
End Sub

Private Sub txtSelect_COMPTEOBL_Change()
cmdSelect_Clear
End Sub

Private Sub txtSelect_COMPTEOBL_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub


Private Sub txtSelect_CRTCPTCPT_Change()
cmdSelect_Clear
End Sub


Private Sub txtSelect_CRTCPTCPT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub



Public Sub fraYCRTCPT0_Display(lCOMPTECOM As String)
Dim xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass

fraYCRTCPT0.Visible = False
cmdYCRTCPT0_Update.Visible = False
cmdYCRTCPT0_Ignore.Visible = False
cmdYCRTCPT0_Exclure.Visible = False

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 ," & paramIBM_Library_SABSPE & ".YCRTCPT0" _
     & " where COMPTECOM = '" & lCOMPTECOM & "' and COMPTECOM = CRTCPTCPT"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    txtD_COMPTECOM = rsSab("COMPTECOM")
    txtD_COMPTEINT = rsSab("COMPTEINT")
    txtD_COMPTEOBL = rsSab("COMPTEOBL")
    txtD_COMPTEFON = rsSab("COMPTEFON")
    txtD_PLANCOPRO = rsSab("PLANCOPRO")
    If rsSab("COMPTEOUV") = 0 Then
        txtD_COMPTEOUV = ""
    Else
        txtD_COMPTEOUV = dateIBM10(rsSab("COMPTEOUV"), True)
    End If
    If rsSab("COMPTECLO") = 0 Then
        txtD_COMPTECLO = ""
    Else
        txtD_COMPTECLO = dateIBM10(rsSab("COMPTECLO"), True)
    End If
    oldYCRTCPT0.CRTCPTCPT = rsSab("CRTCPTCPT")
    oldYCRTCPT0.CRTCPTRUB = rsSab("CRTCPTRUB")
    oldYCRTCPT0.CRTCPTSTA = rsSab("CRTCPTSTA")
    
    lblYCRTCPT0_CRTCPTRUB = "Rubrique : " & oldYCRTCPT0.CRTCPTRUB
    Call cbo_Scan(oldYCRTCPT0.CRTCPTRUB, cboYCRTCPT0_CRTCPTRUB)
    lblYCRTCPT0_CRTCPTSTA = "Code état : " & oldYCRTCPT0.CRTCPTSTA
    Call cbo_Scan(oldYCRTCPT0.CRTCPTSTA, cboYCRTCPT0_CRTCPTSTA)
    
    If arrHab(16) Then
        cmdYCRTCPT0_Update.Visible = True
        If rsSab("CRTCPTSTA") <> "I" Then cmdYCRTCPT0_Ignore.Visible = True
        If rsSab("CRTCPTSTA") <> "E" Then cmdYCRTCPT0_Exclure.Visible = True
        
    End If
    fraYCRTCPT0.Visible = True
End If
Me.Enabled = True: Me.MousePointer = 0
End Sub
Public Sub fraOD_Display(lMOUVEMPIE As Long, lMOUVEMECR As Long)
Dim X As String, K As Integer, K2 As Integer
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass

blnCRTMVTTAUX_OD = False

fraOD.Visible = False
cmdOD_Update.Visible = False
cmdOD_Add.Visible = False
cmdOD_Delete.Visible = False
cmdOD_Log.Visible = False

txtCRTMVTMTE_OD = ""
txtCRTMVTTXT_OD = ""
txtCRTMVTMTD_OD = ""
txtCRTMVTTAUX_OD = ""

If lMOUVEMECR = 0 Then
   oldYCRTMVT0 = zYCRTMVT0_OD
   cboCRTMVTDEV_OD.ListIndex = 0
   cboCRTMVTRUB_OD.ListIndex = 0
   cboCRTMVTCLIP_OD.ListIndex = 0
   fraOD.Caption = "Création d'un mouvement extra-comptable"
   fraOD.ForeColor = mColor_GB
   blnCRTMVTTAUX_OD = True

Else
    X = "select * from " & paramIBM_Library_SABSPE & ".YCRTMVT0 " _
         & " where CRTMVTETA = " & currentSAB_ETA & " and CRTMVTPLA = " & currentSAB_PLA _
         & " and CRTMVTPIE = " & lMOUVEMPIE & " and CRTMVTECR = " & lMOUVEMECR
    Set rsSab = cnsab.Execute(X)
    If rsSab.EOF Then
        V = " Mouvement non trouvé : " & lMOUVEMPIE
        GoTo Error_MsgBox
    End If
    V = rsYCRTMVT0_GetBuffer(rsSab, oldYCRTMVT0)
    
    cmdOD_Log.Visible = True
    cmdOD_Update.Visible = arrHab(15)
    If oldYCRTMVT0.CRTMVTSTA = " " Then cmdOD_Delete.Visible = arrHab(15)
'_____________________________________________________________________________________________________________
    X = "select * from " & paramIBM_Library_SABSPE & ".YCRTLOG0 " _
         & " where CRTLOGETA = " & currentSAB_ETA & " and CRTLOGPLA = " & currentSAB_PLA _
         & " and CRTLOGPIE = " & oldYCRTMVT0.CRTMVTPIE & " and CRTLOGECR = " & oldYCRTMVT0.CRTMVTECR _
         & " order by CRTLOGID desc"
    Set rsSab = cnsab.Execute(X)
    If Not rsSab.EOF Then
    
        X = rsSab("CRTLOGTXT")
        K = InStr(X, "<CRTMVMTD = ") + 11
        If K > 11 Then
            K2 = InStr(K, X, ">")
            If K2 > 0 Then txtCRTMVTMTD_OD = Trim(mId$(X, K, K2 - K))
        End If
        
        K = InStr(X, "<CRTMVTXT = ") + 11
        If K > 11 Then
            K2 = InStr(K, X, ">")
            If K2 > 0 Then txtCRTMVTTXT_OD = Trim(mId$(X, K, K2 - K))
        End If
    End If
'_____________________________________________________________________________________________________________
    Select Case oldYCRTMVT0.CRTMVTSTA
        Case "A"
                fraOD.Caption = "Mouvement extra-comptable annulé"
                fraOD.ForeColor = vbRed
        Case Else
                fraOD.Caption = "Mouvement extra-comptable à déclarer"
                fraOD.ForeColor = vbBlue

    End Select
    Call cbo_Scan(oldYCRTMVT0.CRTMVTRUB, cboCRTMVTRUB_OD)
    Call cbo_Scan(oldYCRTMVT0.CRTMVTCLIP, cboCRTMVTCLIP_OD)
    'Call cbo_Scan(oldYCRTMVT0.CRTMVTDEV, cboCRTMVTDEV_OD)
    
    X = Trim(oldYCRTMVT0.CRTMVTCPT)
    txtCRTMVTCPT_OD = X
    libCRTMVTCPT_OD = ""
    If X <> "" Then
        X = "select COMPTEINT from " & paramIBM_Library_SAB & ".ZCOMPTE0 " _
             & " where COMPTECOM = '" & X & "'"
        Set rsSab = cnsab.Execute(X)
        If Not rsSab.EOF Then libCRTMVTCPT_OD = rsSab("COMPTEINT")

    End If
    
    If oldYCRTMVT0.CRTMVTMTE = 0 Then
        txtCRTMVTMTE_OD = ""
    Else
        txtCRTMVTMTE_OD = Format(Abs(oldYCRTMVT0.CRTMVTMTE), "### ### ##0.00")
    End If
    
    If oldYCRTMVT0.CRTMVTMTE < 0 Then
        optCRTMVTMTD_OD_Db = True
    Else
        optCRTMVTMTD_OD_Cr = True
    End If
    
    X = oldYCRTMVT0.CRTMVTDTR
    Call DTPicker_Set(txtCRTMVTDTR_OD, X)
    
    blnCRTMVTTAUX_OD = True
    cboCRTMVTDEV_OD = oldYCRTMVT0.CRTMVTDEV

End If
    
cmdOD_Add.Visible = arrHab(15)

If mId$(oldYCRTMVT0.CRTMVTDTR, 1, 4) <> mExercice Then
    cmdOD_Update.Visible = False
    cmdOD_Add.Visible = False
    cmdOD_Delete.Visible = False
End If

fraOD.Visible = True

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0
End Sub

Public Sub fraYCRTLOG0_Display(lCRTLOGID As Long, lCRTLOGCPT As String)
Dim xSQL As String, X As String, K As Integer, K1 As Integer, K2 As Integer
Dim blnEnd As Boolean

Me.Enabled = False: Me.MousePointer = vbHourglass

fraYCRTLOG0.Visible = False
lblCRTLOGTXT.Visible = False
If lCRTLOGCPT = "" Then
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCRTLOG0" _
         & " where CRTLOGID = " & lCRTLOGID
Else
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCRTLOG0 , " & paramIBM_Library_SAB & ".ZCOMPTE0" _
         & " where CRTLOGID = " & lCRTLOGID & " and COMPTECOM = CRTLOGCPT"
End If

Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    lblCRTLOGID = rsSab("CRTLOGID")
    X = rsSab("CRTLOGNAT")
    txtCRTLOGNAT = X
    For K = 1 To arrLogNat_Nb
        If X = arrLogNat_code(K) Then
            txtCRTLOGNAT = X & " - " & arrLogNat_Lib(K)
            Exit For
        End If
    Next K
    
    txtCRTLOGUUSR = rsSab("CRTLOGUUSR")
    If lCRTLOGCPT <> "" Then
        txtCRTLOGCPT = rsSab("CRTLOGCPT")
        libCRTLOGCPT = rsSab("COMPTEINT")
    Else
        libCRTLOGCPT = ""
    End If
    
    xttCRTLOGPIE = rsSab("CRTLOGPIE") & " - " & rsSab("CRTLOGECR")
    txtCRTLOGUAMJ = dateImp10_S(rsSab("CRTLOGUAMJ")) & "  " & timeImp8(rsSab("CRTLOGUHMS"))
    
    fgYCRTLOG0.Rows = 1
    fgYCRTLOG0.FormatString = fgYCRTLOG0_FormatString
    fgYCRTLOG0.Row = 0
    blnEnd = False
    K = 1
    X = rsSab("CRTLOGTXT")
    Do While Not blnEnd
        K = InStr(K, X, "<")
        If K > 0 Then
        
            fgYCRTLOG0.Rows = fgYCRTLOG0.Rows + 1
            fgYCRTLOG0.Row = fgYCRTLOG0.Rows - 1
            K1 = InStr(K, X, "=")
            If K1 > 0 Then
                fgYCRTLOG0.Col = 0
                fgYCRTLOG0.Text = mId$(X, K + 1, K1 - K - 1)
            Else
                K1 = K
            End If
            K2 = InStr(K1, X, "|")
            If K2 > 0 Then
                fgYCRTLOG0.Col = 1
                fgYCRTLOG0.Text = mId$(X, K1 + 1, K2 - K1 - 1)
                fgYCRTLOG0.CellForeColor = vbMagenta
            Else
                K2 = K1
            End If
            K = InStr(K2, X, ">")
            If K > 0 Then
                fgYCRTLOG0.Col = 2
                fgYCRTLOG0.Text = mId$(X, K2 + 1, K - K2 - 1)
                fgYCRTLOG0.CellForeColor = vbBlue
            Else
                blnEnd = True
            End If
          
                
        Else
            blnEnd = True
        End If
    Loop
    
    fraYCRTLOG0.Visible = True
    fraYCRTLOG0.ZOrder 0
End If
Me.Enabled = True: Me.MousePointer = 0
End Sub

Public Sub fraYCRTMVT0_Display(lMOUVEMPIE As Long, lMOUVEMECR As Long)
Dim V, X As String
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass

fraYCRTMVT0.Visible = False
cmdYCRTMVT0_Update.Visible = False
cmdYCRTMVT0_Ignore.Visible = False
cmdYCRTMVT0_Mvt_Ann.Visible = False

X = "select * from " & paramIBM_Library_SABSPE & ".YCRTMVT0 , " & paramIBM_Library_SAB & ".ZCOMPTE0 " _
     & " where CRTMVTETA = " & currentSAB_ETA & " and CRTMVTPLA = " & currentSAB_PLA _
     & " and CRTMVTPIE = " & lMOUVEMPIE & " and CRTMVTECR = " & lMOUVEMECR _
     & " and COMPTEETA = " & currentSAB_ETA & " and COMPTEPLA = " & currentSAB_PLA _
     & " and COMPTECOM = CRTMVTCPT"
Set rsSab = cnsab.Execute(X)
If rsSab.EOF Then
    V = " Mouvement non trouvé : " & lMOUVEMPIE
    GoTo Error_MsgBox
Else
    V = rsYCRTMVT0_GetBuffer(rsSab, oldYCRTMVT0)
    
    lblYCRTMVT0_CRTMVTRUB = "Rubrique : " & oldYCRTMVT0.CRTMVTRUB
    Call cbo_Scan(oldYCRTMVT0.CRTMVTRUB, cboYCRTMVT0_CRTMVTRUB)
    'lblYCRTMVT0_CRTMVTCLIP = "Pays : " & oldYCRTMVT0.CRTMVTCLIP
    Call cbo_Scan(oldYCRTMVT0.CRTMVTCLIP, cboYCRTMVT0_CRTMVTCLIP)
    'lblYCRTMVT0_CRTMVTSTA = "Code état : " & oldYCRTMVT0.CRTMVTSTA
    Call cbo_Scan(oldYCRTMVT0.CRTMVTSTA, cboYCRTMVT0_CRTMVTSTA)
    Call cbo_Scan(oldYCRTMVT0.CRTMVTORIG, cboYCRTMVT0_CRTMVTORIG)
    
    txtYCRTMVT0_CRTMVTCPT = oldYCRTMVT0.CRTMVTCPT
    'X = "select COMPTEINT from " & paramIBM_Library_SAB & ".ZCOMPTE0 " _
    '     & " where COMPTEETA = " & currentSAB_ETA & " and COMPTEPLA = " & currentSAB_PLA _
    '     & " and COMPTECOM = '" & oldYCRTMVT0.CRTMVTCPT & "'"
    'Set rsSabX = cnsab.Execute(X)
    'If rsSabX.EOF Then
    '    libYCRTMVT0_CRTMVTCPT = "???"
    'Else
        libYCRTMVT0_CRTMVTCPT = Trim(rsSab("COMPTEINT"))
    'End If
    
    
    If arrHab(15) Then
        cmdYCRTMVT0_Update.Visible = True
        If rsSab("CRTMVTSTA") <> "I" Then cmdYCRTMVT0_Ignore.Visible = True
        
    End If
    libYCRTMVT0_CRTMVTMTE = Format$(oldYCRTMVT0.CRTMVTMTE, "### ### ##0.00")
    If oldYCRTMVT0.CRTMVTMTE < 0 Then
        libYCRTMVT0_CRTMVTMTE.ForeColor = vbRed
    Else
        libYCRTMVT0_CRTMVTMTE.ForeColor = vbBlue
    End If
    
    libYCRTMVT0_CRTMVTDOS = oldYCRTMVT0.CRTMVTSER & " " & oldYCRTMVT0.CRTMVTSSE & " " & oldYCRTMVT0.CRTMVTOPE & " " & oldYCRTMVT0.CRTMVTNAT & " " & oldYCRTMVT0.CRTMVTEVE & " " & oldYCRTMVT0.CRTMVTDOS
    txtYCRTMVT0_CRTMVTCLIN = oldYCRTMVT0.CRTMVTCLIC & " " & oldYCRTMVT0.CRTMVTCLIN
    V = sqlCRTMVTCLI(oldYCRTMVT0.CRTMVTCLIC, oldYCRTMVT0.CRTMVTCLIN, xZCLIENA0, xZADRESS0)
    libYCRTMVT0_CRTMVTCLIN = Trim(xZCLIENA0.CLIENARA1) & " " & Trim(xZCLIENA0.CLIENARA2) _
                          & vbCrLf & Trim(xZADRESS0.ADRESSAD1) & Trim(xZADRESS0.ADRESSAD2) _
                          & vbCrLf & Trim(xZADRESS0.ADRESSCOP) & Trim(xZADRESS0.ADRESSVIL) _
                          & vbCrLf & Trim(xZADRESS0.ADRESSPAY)
    X = Trim(xZCLIENA0.CLIENARSD)
    lblYCRTMVT0_CRTMVTCLIN = "Pays de résidence : " & X
    If X = oldYCRTMVT0.CRTMVTCLIP Then
        lblYCRTMVT0_CRTMVTCLIN.BackColor = fraYCRTMVT0.BackColor
        lblYCRTMVT0_CRTMVTCLIN.ForeColor = fraYCRTMVT0.ForeColor
    Else
        lblYCRTMVT0_CRTMVTCLIN.BackColor = vbRed
        lblYCRTMVT0_CRTMVTCLIN.ForeColor = vbYellow
    End If
    
    cboYCRTMVT0_CRTMVTSTA.Locked = True
    cboYCRTMVT0_CRTMVTORIG.Locked = True
    fraYCRTMVT0.Visible = True

End If


X = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
     & " where MOUVEMETA = '" & currentSAB_ETA & "' and MOUVEMPIE = " & lMOUVEMPIE & " order by MOUVEMECR"
Set rsSab = cnsab.Execute(X)

Call fgBIAMVT_Display(lMOUVEMPIE, lMOUVEMECR)
If oldYBIAMVTH.MOUVEMNUM = 0 Then
    fgYSWISAB0.Visible = False
Else
    fgYSWISAB0.Visible = False
   ' Call fgYSWISAB0_Display(oldYBIAMVTH.MOUVEMOPE, oldYBIAMVTH.MOUVEMNUM)
End If
SSTab2.Tab = 0

If cmdSelect_SQL_K = "Annulation$" Then cmdYCRTMVT0_Mvt_Ann.Visible = arrHab(15) And arrHab(18)

If mId$(oldYCRTMVT0.CRTMVTDTR, 1, 4) <> mExercice Then
    cmdYCRTMVT0_Update.Visible = False
    cmdYCRTMVT0_Ignore.Visible = False
    cmdYCRTMVT0_Mvt_Ann.Visible = False
End If

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub fgSwift_Display(lSWISABSWID As Long, lMTK As String, lBIC As String)
Dim wColor As Long, wColorFixed As Long
'Dim X As String, xWhere As String, xOPE As String
Dim xSQL As String
'Dim I As Long
'Dim blnOk As Boolean, blnDisplay As Boolean
'Dim wAmj As String

On Error GoTo Error_Handler
fraSwift.Visible = False
'fgswift_Reset
If Not blnSIDE_DB Then
    cnSIDE_DB.Open paramODBC_DSN_SIDE_DB
    blnSIDE_DB = True
End If
fgSwift.Rows = 1
'fgSwift.FormatString = fgSwift_FormatString
fgSwift.FormatString = "<" & lMTK & "    |<" & lBIC & "                                                       ||"
fgSwift.Row = 0
fgSwift.Col = 0: fgSwift.CellFontBold = True
fgSwift.Col = 1: fgSwift.CellFontBold = True
currentAction = "fgswift_Display"


xSQL = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABSWID = " & lSWISABSWID
Set rsSab = cnsab.Execute(xSQL)
'___________________________________________________________________
 If Not rsSab.EOF Then
    Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)

    If oldYSWISAB0.SWISABWES = "E" Then
        X = "reçu de "
        wColor = RGB(190, 240, 255)
        wColorFixed = vbBlue
    Else
        X = "émis vers "
        wColor = RGB(220, 255, 220)
        wColorFixed = RGB(0, 64, 0)
    End If
    libSWIFT_SWISABSWID = "Dossier : " & Trim(oldYSWISAB0.SWISABOPEC) & " " & Format(oldYSWISAB0.SWISABOPEN, "### ###")
    fgSwift.Col = 0: fgSwift.Text = oldYSWISAB0.SWISABWMTK
    fgSwift.CellFontBold = True: fgSwift.CellBackColor = wColor
    fgSwift.ForeColorFixed = wColorFixed
    fgSwift.Col = 1: fgSwift.Text = X & oldYSWISAB0.SWISABWBIC & " le " & dateImp10(oldYSWISAB0.SWISABWAMJ) & " " & timeImp8(oldYSWISAB0.SWISABWHMS)
    fgSwift.CellFontBold = True: fgSwift.CellBackColor = wColor
    fgSwift.ForeColorFixed = wColorFixed
    fraSwift.BackColor = wColor

'If Not rsSab.EOF Then
'    libSWIFT_SWISABSWID = lSWISABSWID & " - " & rsSab("SWISABOPEC") & " " & rsSab("SWISABOPEN")

'    If rsSab("SWISABWES") = "E" Then
'        fgSwift.Col = 0: fgSwift.CellBackColor = RGB(32, 160, 255)
'        fgSwift.Col = 1: fgSwift.CellBackColor = RGB(32, 160, 255)
'        wColor = RGB(190, 240, 255)
'    Else
'        fgSwift.Col = 0: fgSwift.CellBackColor = RGB(32, 230, 190)
'        fgSwift.Col = 1: fgSwift.CellBackColor = RGB(32, 230, 190)
'        wColor = mColor_G0
'    End If
    xSQL = "select * from rtextField " _
        & "where Aid = " & rsSab("SWISABWID1") _
        & " and text_s_umidl = " & rsSab("SWISABWIDL") _
        & " and text_s_umidh  =  " & rsSab("SWISABWIDH") _
        & " order by field_cnt"
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
    If Not rsSIDE_DB.EOF Then
        Do While Not rsSIDE_DB.EOF
        
            fgSwift.Rows = fgSwift.Rows + 1
            fgSwift.Row = fgSwift.Rows - 1
        
            fgSwift_DisplayLine fgSwift.Row, wColor, wColorFixed
        
            rsSIDE_DB.MoveNext
        
        Loop
    Else
        xSQL = "select * from rtext " _
            & "where Aid = " & rsSab("SWISABWID1") _
            & " and text_s_umidl = " & rsSab("SWISABWIDL") _
            & " and text_s_umidh  =  " & rsSab("SWISABWIDH")
        Set rsSIDE_DB = cnSIDE_DB.Execute(xSQL)
        If Not rsSIDE_DB.EOF Then
            Call srvrText_GetBuffer_ODBC(rsSIDE_DB, xrText)
            fgSwift_DisplayLine_rText fgSwift.Row, wColor, wColorFixed
        End If
    End If
    
    fraSwift.Visible = True
    'fraSwift.ZOrder 0
    'If chkSIDE_DB_Show Then frmSIDE_DB.fgSwift_Display lSWISABSWID, 0, 0, 0

End If

'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub fgSwift_DisplayLine(lIndex As Long, lCellBackColor As Long, lColorFixed As Long)
Dim K As Integer, iAsc13 As Integer, iLen As Integer
Dim blnAsc13 As Boolean
Dim xValue As String

On Error Resume Next
fgSwift.Col = 0: fgSwift.Text = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
fgSwift.CellBackColor = lCellBackColor
fgSwift.CellForeColor = lColorFixed
fgSwift.Col = 1
fgSwift.CellForeColor = lColorFixed

Select Case rsSIDE_DB("field_code")
    Case "45", "46", "47":   xValue = rsSIDE_DB("value_memo")
    Case Else:    xValue = rsSIDE_DB("value")
End Select

 iLen = Len(xValue)
 K = 1
 Do
    iAsc13 = InStr(K, xValue, Asc13)
    If iAsc13 > 0 Then
        fgSwift.Text = Trim(mId$(xValue, K, iAsc13 - K))
        fgSwift.CellForeColor = lColorFixed
        K = iAsc13 + 2
        fgSwift.Rows = fgSwift.Rows + 1
        fgSwift.Row = fgSwift.Rows - 1
    End If
 Loop Until iAsc13 = 0

fgSwift.Text = Trim(mId$(xValue, K, iLen - K + 1))
fgSwift.CellForeColor = lColorFixed
fgSwift.Col = fgSwift.Cols - 1: fgSwift.Text = rsSIDE_DB("field_cnt")




End Sub



Public Sub fgSwift_DisplayLine_rText(lIndex As Long, lCellBackColor As Long, lColorFixed As Long)
Dim K As Integer, iAsc13 As Integer, iLen As Integer
Dim blnAsc13 As Boolean
Dim xValue As String, X As String, K2 As Integer

On Error Resume Next

xValue = xrText.text_data_block & Asc13
iLen = Len(xValue)
If mId$(xValue, 1, 3) = Asc13 & Asc10 & ":" Then
    K = 3
Else
    K = 1
End If
Do
    iAsc13 = InStr(K, xValue, Asc13)
    If iAsc13 > 0 Then
        fgSwift.Rows = fgSwift.Rows + 1
        fgSwift.Row = fgSwift.Rows - 1
        X = Trim(mId$(xValue, K, iAsc13 - K))
        fgSwift.Col = 1
        fgSwift.CellForeColor = lColorFixed
        If mId$(X, 1, 1) <> ":" Then
            fgSwift.Text = Trim(mId$(xValue, K, iAsc13 - K))
        Else
            K2 = InStr(2, X, ":")
            If K2 > 0 Then
                fgSwift.Text = Trim(mId$(X, K2 + 1, Len(X) - K2))
                fgSwift.Col = 0: fgSwift.Text = Trim(mId$(X, 2, K2 - 2))
                fgSwift.CellBackColor = lCellBackColor
                fgSwift.CellForeColor = lColorFixed
            Else
                fgSwift.Text = Trim(mId$(xValue, K, iAsc13 - K))
            End If
        End If
        
        K = iAsc13 + 2
    End If
 Loop Until iAsc13 = 0


End Sub


Public Sub fgBIAMVT_DisplayLine()

On Error Resume Next
fgBIAMVT.Col = 0: fgBIAMVT.Text = xYBIAMVTH.MOUVEMSER & " " & xYBIAMVTH.MOUVEMSSE & " " & xYBIAMVTH.MOUVEMOPE & " " & xYBIAMVTH.MOUVEMEVE & " " & xYBIAMVTH.MOUVEMNUM

fgBIAMVT.Col = 1: fgBIAMVT.Text = xYBIAMVTH.MOUVEMCOM

fgBIAMVT.Col = IIf(xYBIAMVTH.MOUVEMMON < 0, 3, 2)

fgBIAMVT.Text = Format$(Abs(xYBIAMVTH.MOUVEMMON), "### ### ### ##0.00")

If xYBIAMVTH.MOUVEMMON > 0 Then
    fgBIAMVT.CellForeColor = vbRed
Else
    fgBIAMVT.CellForeColor = vbBlue
End If
fgBIAMVT.Col = 4: fgBIAMVT.Text = xYBIAMVTH.COMPTEDEV

fgBIAMVT.Col = 5: fgBIAMVT.Text = xYBIAMVTH.LIBELLIB1 & xYBIAMVTH.LIBELLIB2 & xYBIAMVTH.LIBELLIB3 & xYBIAMVTH.LIBELLIB4
fgBIAMVT.Col = 6: fgBIAMVT.Text = dateImp10_S(xYBIAMVTH.MOUVEMDTR + 19000000)
fgBIAMVT.Col = 8: fgBIAMVT.Text = xYBIAMVTH.MOUVEMPIE
fgBIAMVT.Col = 9: fgBIAMVT.Text = xYBIAMVTH.MOUVEMECR
'fgBIAMVT.Col = fgBIAMVT_arrIndex: fgBIAMVT.Text = lIndex
End Sub
Private Sub fgBIAMVT_Display(lMOUVEMPIE As Long, lMOUVEMECR As Long)
Dim V, K As Integer, wRow As Long

On Error GoTo Error_Handler
fgBIAMVT.Visible = False
fgBIAMVT_Reset

fgBIAMVT.Rows = 1
fgBIAMVT.FormatString = fgBIAMVT_FormatString
fgBIAMVT.Row = 0

currentAction = "fgBIAMVT_Display"

Do While Not rsSab.EOF
    fgBIAMVT.Rows = fgBIAMVT.Rows + 1
    fgBIAMVT.Row = fgBIAMVT.Rows - 1
    V = rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVTH)
    fgBIAMVT_DisplayLine
    
    If lMOUVEMECR = xYBIAMVTH.MOUVEMECR Then
        oldYBIAMVTH = xYBIAMVTH
        wRow = fgBIAMVT.Row
        
        fgBIAMVT.Col = 0: fgBIAMVT.CellForeColor = vbMagenta
        fgBIAMVT.Col = 1: fgBIAMVT.CellForeColor = vbMagenta
        fgBIAMVT.Col = 5: fgBIAMVT.CellForeColor = vbMagenta
        For K = 0 To 9
            fgBIAMVT.Col = K: fgBIAMVT.CellBackColor = &HD0FFFF
        Next K
    End If
    
    rsSab.MoveNext

Loop
'fgBIAMVT.TopRow = wRow
fgBIAMVT.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgBIAMVT_Reset()
fgBIAMVT.Clear
fgBIAMVT_Sort1 = 0: fgBIAMVT_Sort2 = 0
fgBIAMVT_Sort1_Old = -1
fgBIAMVT_RowDisplay = 0: fgBIAMVT_RowClick = 0
fgBIAMVT_arrIndex = fgBIAMVT.Cols - 1
blnfgBIAMVT_DisplayLine = False
fgBIAMVT_SortAD = 6
fgBIAMVT.LeftCol = fgBIAMVT.FixedCols

End Sub


Public Sub cmdSelect_SQL_2_Importation_YCRTMVT0(lCRTCPTCPT As String)
Dim X As String, mCRTCPTSTA As String, dblX As Double
Dim blnTransaction As Boolean, blnOk As Boolean
Dim xMOUVEMOPE As String, K As Long
Dim blnCRTMVTSTA_Ann As Boolean

On Error GoTo Error_Handler

Me.Enabled = False: Me.MousePointer = vbHourglass
App_Debug = "cmdSelect_SQL_Importation_YCRTMVT0"
blnTransaction = False
'________________________________________________________________________________
Call lstErr_ChangeLastItem(lstErr, cmdContext, lCRTCPTCPT): DoEvents


Call rsYCRTMVT0_Init(newYCRTMVT0)

X = "SELECT COMPTEDEV , CRTCPTRUB , CRTCPTSTA FROM " & paramIBM_Library_SAB & ".ZCOMPTE0 ," & paramIBM_Library_SABSPE & ".YCRTCPT0" _
  & " where COMPTECOM = '" & lCRTCPTCPT & "'" _
  & " and CRTCPTCPT = COMPTECOM"
  
Set rsSab = cnsab.Execute(X)
If rsSab.EOF Then
    V = " Compte non trouvé : " & lCRTCPTCPT
    GoTo Error_MsgBox
Else
    newYCRTMVT0.CRTMVTDEV = rsSab("COMPTEDEV")
    newYCRTMVT0.CRTMVTRUB = rsSab("CRTCPTRUB")
    newYCRTMVT0.CRTMVTORIG = "*"
    mCRTCPTSTA = rsSab("CRTCPTSTA")
End If


currentAction = "cmdSelect_SQL_Importation_YCRTMVT0"

X = "SELECT * FROM " & paramIBM_Library_SABSPE & ".YBIAMVTHP left outer join " & paramIBM_Library_SABSPE & ".YTVACOM0" _
  & " on tvacometa = mouvemeta and tvacompla = mouvempla and tvacompie = mouvempie and tvacomecr = mouvemecr " _
  & " where mouvemeta = " & currentSAB_ETA & " and mouvempla = " & currentSAB_PLA & " and mouvemcom = '" & lCRTCPTCPT & "'" _
  & " and mouvemdtr >= " & mAmjMin_CRTMVTDTR - 19000000 & " and mouvemdtr <= " & mAmjMax_Exercice - 19000000 _
  & " order by MOUVEMDTR , MOUVEMPIE , MOUVEMECR"
  
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    xMOUVEMOPE = rsSab("MOUVEMOPE")
    blnOk = True
    Select Case xMOUVEMOPE
        Case "PPD", "C01": blnOk = False
        Case "ECH":
                Select Case rsSab("MOUVEMEVE")
                    Case "PRO": blnOk = False
                    Case "ECH": If InStr(rsSab("LIBELLIB1"), "REP. PROV.") > 0 Then blnOk = False
                End Select
    End Select
        
    If blnOk Then
    '___________________________________________________________________________________
    
            If Not blnTransaction Then
                V = cnSAB_Transaction("BeginTrans")
                If Not IsNull(V) Then GoTo Error_MsgBox
                blnTransaction = True
            End If
        '___________________________________________________________________________________
        'If rsSab("MOUVEMNUM") = 174066 Then
        '    Debug.Print rsSab("MOUVEMNUM")
        'End If
            newYCRTMVT0.CRTMVTETA = rsSab("MOUVEMETA")
            newYCRTMVT0.CRTMVTPLA = rsSab("MOUVEMPLA")
            newYCRTMVT0.CRTMVTPIE = rsSab("MOUVEMPIE")
            newYCRTMVT0.CRTMVTECR = rsSab("MOUVEMECR")
            newYCRTMVT0.CRTMVTCPT = rsSab("MOUVEMCOM")
            newYCRTMVT0.CRTMVTDTR = rsSab("MOUVEMDTR") + 19000000
            newYCRTMVT0.CRTMVTMTE = -rsSab("MOUVEMMON")
            
            newYCRTMVT0.CRTMVTSER = rsSab("MOUVEMSER")
            newYCRTMVT0.CRTMVTSSE = rsSab("MOUVEMSSE")
            newYCRTMVT0.CRTMVTOPE = rsSab("MOUVEMOPE")
            '''newYCRTMVT0.CRTMVTnat = rsSab("MOUVEMNAT")
            newYCRTMVT0.CRTMVTEVE = rsSab("MOUVEMEVE")
            newYCRTMVT0.CRTMVTDOS = rsSab("MOUVEMnum")
            
            If newYCRTMVT0.CRTMVTDEV <> "EUR" Then
                Call sqlYBIATAB0_Read("PDC", newYCRTMVT0.CRTMVTDEV, CStr(newYCRTMVT0.CRTMVTDTR), X)
                If IsNumeric(mId$(X, 9, 15)) Then
                    dblX = CDbl(mId$(X, 9, 15) / 1000000000)
                    If dblX <> 0 Then newYCRTMVT0.CRTMVTMTE = Round(newYCRTMVT0.CRTMVTMTE / dblX, 2)
                Else
                    newYCRTMVT0.CRTMVTMTE = 0
                End If
            End If
        
            Select Case mCRTCPTSTA
                Case "*": newYCRTMVT0.CRTMVTSTA = " "
                Case "I": newYCRTMVT0.CRTMVTSTA = "I"
                Case Else: newYCRTMVT0.CRTMVTSTA = "?"
            End Select
                
            newYCRTMVT0.CRTMVTCOMK = ""
            newYCRTMVT0.CRTMVTCLIC = ""
            newYCRTMVT0.CRTMVTCLIN = ""
            newYCRTMVT0.CRTMVTCLIP = ""
            blnCRTMVTSTA_Ann = False
           If IsNull(rsSab("TVACOMCLIP")) Then
                '''newYCRTMVT0.CRTMVTSTA = "?"
            Else
                X = Trim(rsSab("TVACOMCLIP"))
                If X = "XX" Then X = ""
                newYCRTMVT0.CRTMVTCOMK = rsSab("TVACOMGTYP")
                newYCRTMVT0.CRTMVTCLIC = rsSab("TVACOMCLIC")
                If Val(rsSab("TVACOMCLI")) <> 0 Then
                    newYCRTMVT0.CRTMVTCLIN = Format(Val(rsSab("TVACOMCLI")), "0000000")
                End If
                newYCRTMVT0.CRTMVTCLIP = X
                If Trim(rsSab("TVACOMCOMC")) = "CCHG44" Then blnCRTMVTSTA_Ann = True
                
                newYCRTMVT0.CRTMVTSER = rsSab("TVACOMSER")
                newYCRTMVT0.CRTMVTSSE = rsSab("TVACOMSSE")
                newYCRTMVT0.CRTMVTOPE = rsSab("TVACOMOPE")
                newYCRTMVT0.CRTMVTNAT = rsSab("TVACOMNAT")
                newYCRTMVT0.CRTMVTEVE = rsSab("TVACOMEVE")
                newYCRTMVT0.CRTMVTDOS = rsSab("TVACOMDOS")

            End If
            
            If mId$(xMOUVEMOPE, 1, 1) = "*" Then Call cmdSelect_SQL_2_Importation_YCRTMVT0_YDOSXOD0
            
            If newYCRTMVT0.CRTMVTCLIP = "" Then
                Select Case newYCRTMVT0.CRTMVTOPE
                    Case "CDE": Call cmdSelect_SQL_2_Importation_YCRTMVT0_ZCDODOS0(newYCRTMVT0.CRTMVTOPE, newYCRTMVT0.CRTMVTDOS)
                    Case "CRE": Call cmdSelect_SQL_2_Importation_YCRTMVT0_ZCREEMP0
                    Case "ENG", "AP1": Call cmdSelect_SQL_2_Importation_YCRTMVT0_ZCAUDOS0
                    Case "PRE", "EMP": Call cmdSelect_SQL_2_Importation_YCRTMVT0_ZTREOPE0
                    Case "TRF", "CPT": Call cmdSelect_SQL_2_Importation_YCRTMVT0_YSWISAB1(newYCRTMVT0.CRTMVTOPE, newYCRTMVT0.CRTMVTDOS)
                    Case "ECH": Call cmdSelect_SQL_2_Importation_YCRTMVT0_ECH
                End Select
            End If
            
            If Trim(newYCRTMVT0.CRTMVTCLIP) = "" Then
                If Trim(newYCRTMVT0.CRTMVTSTA) = "" Then newYCRTMVT0.CRTMVTSTA = "?"
            Else
                If InStr(mPays_Exclus, Trim(newYCRTMVT0.CRTMVTCLIP)) > 0 Then newYCRTMVT0.CRTMVTSTA = "I"
            End If
            
            If Trim(newYCRTMVT0.CRTMVTCLIN) <> "" Then
                If InStr(mClients_Exclus, newYCRTMVT0.CRTMVTCLIN) > 0 Then newYCRTMVT0.CRTMVTSTA = "I"
            End If
            If blnCRTMVTSTA_Ann Then newYCRTMVT0.CRTMVTSTA = "A"
            
            If newYCRTMVT0.CRTMVTSTA = "?" Then
                For K = 1 To arrCRT_Mvt_Ann_Nb
                    If arrCRT_Mvt_Ann(K).CRTMVTCPT > newYCRTMVT0.CRTMVTCPT Then Exit For
                    If arrCRT_Mvt_Ann(K).CRTMVTSER = newYCRTMVT0.CRTMVTSER _
                    And arrCRT_Mvt_Ann(K).CRTMVTSSE = newYCRTMVT0.CRTMVTSSE _
                    And arrCRT_Mvt_Ann(K).CRTMVTOPE = newYCRTMVT0.CRTMVTOPE Then
                    
                        If arrCRT_Mvt_Ann(K).CRTMVTEVE = newYCRTMVT0.CRTMVTEVE _
                        Or Trim(arrCRT_Mvt_Ann(K).CRTMVTEVE) = "" Then
                            newYCRTMVT0.CRTMVTSTA = "A"
                            Exit For
                        End If
                    End If
                Next K
            End If
            'JPL newYCRTMVT0.CRTMVTCLIN = rsSab("MOUVEMOPE") & rsSab("MOUVEMEVE")
            V = sqlYCRTMVT0_Insert(newYCRTMVT0)
            If Not IsNull(V) Then GoTo Error_MsgBox
    End If
'___________________________________________________________________________________
    rsSab.MoveNext
Loop

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If blnTransaction Then
        If Not IsNull(V) Then
            V = cnSAB_Transaction("Rollback")
        Else
            V = cnSAB_Transaction("Commit")
        End If
    End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdSelect_SQL_2_YCRTMVT0(lCRTCPTCPT As String)
Dim X As String, xWhere As String
On Error GoTo Error_Handler

mCRTMVTCPT_Display = lCRTCPTCPT

Call DTPicker_Control(txtSelect_AmjMin, wAMJMin)
Call DTPicker_Control(txtSelect_AmjMax, WAMJMax)

'________________________________________________________________________________

currentAction = "cmdSelect_SQL_2_YCRTMVT0"
xWhere = " and CRTMVTDTR >= " & wAMJMin & " and CRTMVTDTR <= " & WAMJMax

X = Trim(mId$(cboSelect_CRTMVTSTA, 1, 1))
Select Case X
    Case "": xWhere = xWhere & " and CRTMVTSTA = ' '"
    Case "#": xWhere = xWhere & " and CRTMVTSTA <> 'I'"
    Case "*"
    Case Else: xWhere = xWhere & " and CRTMVTSTA = '" & X & "'"
End Select

X = Trim(mId$(cboSelect_CRTMVTORIG, 1, 1))
If X <> "" Then xWhere = xWhere & " and CRTMVTORIG = '" & X & "'"

X = Trim(mId$(cboSelect_CRTMVTCLIP, 1, 2))
If X <> "" Then xWhere = xWhere & " and CRTMVTCLIP = '" & X & "'"

 
X = "select * from " & paramIBM_Library_SABSPE & ".YCRTMVT0 , " & paramIBM_Library_SABSPE & ".YBIAMVTHP" _
     & " where CRTMVTCPT = '" & lCRTCPTCPT & "'" & xWhere _
     & " and MOUVEMETA = CRTMVTETA and MOUVEMPIE = CRTMVTPIE and MOUVEMECR = CRTMVTECR" _
     & " order by CRTMVTDTR , CRTMVTPIE , CRTMVTECR"
Set rsSab = cnsab.Execute(X)

fgDetail_Display


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:

End Sub
Public Sub cmdSelect_SQL_2()
Dim X As String, xWhere As String
On Error GoTo Error_Handler

Call DTPicker_Control(txtSelect_AmjMin, wAMJMin)
Call DTPicker_Control(txtSelect_AmjMax, WAMJMax)

'________________________________________________________________________________

currentAction = "cmdSelect_SQL_2"
xWhere = " where CRTMVTDTR >= " & wAMJMin & " and CRTMVTDTR <= " & WAMJMax

If cmdSelect_SQL_K = "2+" Then xWhere = xWhere & " and CRTMVTORIG = '+'"

Select Case cmdSelect_SQL_K
    Case "Annulation$": xWhere = xWhere & " and CRTMVTSTA = '?'"
        
    Case Else:
    
        X = Trim(mId$(cboSelect_CRTMVTSTA, 1, 1))
        Select Case X
            Case "": xWhere = xWhere & " and CRTMVTSTA = ' '"
            Case "#": xWhere = xWhere & " and CRTMVTSTA <> 'I'"
            Case "*"
            Case Else: xWhere = xWhere & " and CRTMVTSTA = '" & X & "'"
        End Select
End Select

X = Trim(mId$(cboSelect_CRTMVTORIG, 1, 1))
If X <> "" Then xWhere = xWhere & " and CRTMVTORIG = '" & X & "'"

X = Trim(mId$(cboSelect_CRTMVTCLIP, 1, 2))
If X <> "" Then xWhere = xWhere & " and CRTMVTCLIP = '" & X & "'"

X = Trim(txtSelect_CRTCPTCPT)
If X <> "" Then xWhere = xWhere & " and CRTMVTCPT like '" & X & "%'"
X = Trim(mId$(cboSelect_CRTCPTRUB, 1, 5))
If X <> "" Then xWhere = xWhere & " and CRTMVTRUB ='" & X & "'"

 
X = "select * from " & paramIBM_Library_SABSPE & ".YCRTMVT0 left outer join " & paramIBM_Library_SABSPE & ".YBIAMVTHP" _
     & " on MOUVEMETA = CRTMVTETA and MOUVEMPIE = CRTMVTPIE and MOUVEMECR = CRTMVTECR" _
     & xWhere _
     & " order by CRTMVTDTR , CRTMVTPIE , CRTMVTECR"
Set rsSab = cnsab.Execute(X)

fgDetail_Display

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:

End Sub



Public Sub cmdYCRTMVT0_Update_Transaction()
Dim V, X As String
Dim blnTransaction As Boolean
On Error GoTo Error_Handler


blnTransaction = False

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
blnTransaction = True

V = sqlYCRTMVT0_Update(newYCRTMVT0, oldYCRTMVT0)
If Not IsNull(V) Then GoTo Error_MsgBox
    
newYCRTLOG0 = zYCRTLOG0
newYCRTLOG0.CRTLOGNAT = "M02"
newYCRTLOG0.CRTLOGCPT = newYCRTMVT0.CRTMVTCPT
newYCRTLOG0.CRTLOGETA = newYCRTMVT0.CRTMVTETA
newYCRTLOG0.CRTLOGPLA = newYCRTMVT0.CRTMVTPLA
newYCRTLOG0.CRTLOGPIE = newYCRTMVT0.CRTMVTPIE
newYCRTLOG0.CRTLOGECR = newYCRTMVT0.CRTMVTECR
newYCRTLOG0.CRTLOGCPT = newYCRTMVT0.CRTMVTCPT
X = ""
If newYCRTMVT0.CRTMVTDEV <> oldYCRTMVT0.CRTMVTDEV Then
    X = X & "<CRTMVTDEV = " & oldYCRTMVT0.CRTMVTDEV & " | " & newYCRTMVT0.CRTMVTDEV & ">"
End If
If newYCRTMVT0.CRTMVTCLIC <> oldYCRTMVT0.CRTMVTCLIC Then
    X = X & "<CRTMVTCLIC = " & oldYCRTMVT0.CRTMVTCLIC & " | " & newYCRTMVT0.CRTMVTCLIC & ">"
End If
If newYCRTMVT0.CRTMVTCLIN <> oldYCRTMVT0.CRTMVTCLIN Then
    X = X & "<CRTMVTCLIN = " & oldYCRTMVT0.CRTMVTCLIN & " | " & newYCRTMVT0.CRTMVTCLIN & ">"
End If
If newYCRTMVT0.CRTMVTCLIP <> oldYCRTMVT0.CRTMVTCLIP Then
    X = X & "<CRTMVTCLIP = " & oldYCRTMVT0.CRTMVTCLIP & " | " & newYCRTMVT0.CRTMVTCLIP & ">"
End If
If newYCRTMVT0.CRTMVTRUB <> oldYCRTMVT0.CRTMVTRUB Then
    X = X & "<CRTMVTRUB = " & oldYCRTMVT0.CRTMVTRUB & " | " & newYCRTMVT0.CRTMVTRUB & ">"
End If
If newYCRTMVT0.CRTMVTMTE <> oldYCRTMVT0.CRTMVTMTE Then
    X = X & "<CRTMVTMTE = " & oldYCRTMVT0.CRTMVTMTE & " | " & newYCRTMVT0.CRTMVTMTE & ">"
End If
If newYCRTMVT0.CRTMVTSTA <> oldYCRTMVT0.CRTMVTSTA Then
    X = X & "<CRTMVTSTA = " & oldYCRTMVT0.CRTMVTSTA & " | " & newYCRTMVT0.CRTMVTSTA & ">"
End If
If newYCRTMVT0.CRTMVTORIG <> oldYCRTMVT0.CRTMVTORIG Then
    X = X & "<CRTMVTORIG = " & oldYCRTMVT0.CRTMVTORIG & " | " & newYCRTMVT0.CRTMVTORIG & ">"
End If
newYCRTLOG0.CRTLOGTXT = X
V = sqlYCRTLOG0_Insert(newYCRTLOG0)
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If blnTransaction Then
        If Not IsNull(V) Then
            V = cnSAB_Transaction("Rollback")
        Else
            V = cnSAB_Transaction("Commit")
            fraYCRTMVT0.Visible = False
            Select Case cmdSelect_SQL_K
                Case "2c": cmdSelect_SQL_2_YCRTMVT0 newYCRTMVT0.CRTMVTCPT
                Case "2": cmdSelect_SQL_2
            End Select
        End If
    End If

End Sub

Public Sub cmdYCRTCPT0_Update_Transaction()
Dim V, X As String
Dim blnTransaction As Boolean
On Error GoTo Error_Handler

blnTransaction = False

'Call MsgBox("Les mvts sélectionnés dans YCRTMVT0 ne sont pas supprimés", vbInformation, "Attention : Compte à ignorer")

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
blnTransaction = True


V = sqlYCRTCPT0_Update(newYCRTCPT0, oldYCRTCPT0)
If Not IsNull(V) Then GoTo Error_MsgBox

newYCRTLOG0 = zYCRTLOG0
newYCRTLOG0.CRTLOGNAT = "C02"
newYCRTLOG0.CRTLOGCPT = newYCRTCPT0.CRTCPTCPT
X = ""
If newYCRTCPT0.CRTCPTRUB <> oldYCRTCPT0.CRTCPTRUB Then X = "<CRTCPTRUB = " & oldYCRTCPT0.CRTCPTRUB & " | " & newYCRTCPT0.CRTCPTRUB & ">"
If newYCRTCPT0.CRTCPTSTA <> oldYCRTCPT0.CRTCPTSTA Then X = X & "<CRTCPTSTA = " & oldYCRTCPT0.CRTCPTSTA & " | " & newYCRTCPT0.CRTCPTSTA & ">"
newYCRTLOG0.CRTLOGTXT = X
V = sqlYCRTLOG0_Insert(newYCRTLOG0)
If Not IsNull(V) Then GoTo Error_MsgBox
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If blnTransaction Then
        If Not IsNull(V) Then
            V = cnSAB_Transaction("Rollback")
        Else
            V = cnSAB_Transaction("Commit")
            If cmdSelect_SQL_K <> "JPL" Then
                fraYCRTCPT0.Visible = False
                cmdSelect_SQL_1
            End If
        End If
    End If


End Sub

Private Sub txtSelect_CRTLOGCPT_Change()
cmdSelect_Clear

End Sub



Public Sub cmdSelect_SQL_2_Importation_YCRTMVT0_YSWISAB1(lSWISABOPEC As String, LSWISABOPEN As Long)
Dim X As String

X = "SELECT SWISABW50P , SWISABW59P , SWISABW71A FROM " & paramIBM_Library_SABSPE & ".YSWISAB0N ," & paramIBM_Library_SABSPE & ".YSWISAB1" _
  & " where SWISABOPEC = '" & lSWISABOPEC & "' And SWISABOPEN = " & LSWISABOPEN _
  & " and SWISABSWID = SWISAB1ID" _
  & " order by SWISABSWID desc"
  
Set rsSabX = cnsab.Execute(X)
If Not rsSabX.EOF Then
    Select Case rsSabX("SWISABW71A")
        Case "O": newYCRTMVT0.CRTMVTCOMK = "D"
                  newYCRTMVT0.CRTMVTCLIP = rsSabX("SWISABW50P")
       Case Else: newYCRTMVT0.CRTMVTCOMK = rsSabX("SWISABW71A")
                  newYCRTMVT0.CRTMVTCLIP = rsSabX("SWISABW59P")
    End Select
    
    Exit Sub
End If

End Sub
Public Sub cmdSelect_SQL_2_Importation_YCRTMVT0_ZCREEMP0()
Dim X As String, wDos As Long
On Error GoTo Error_Handler

Select Case rsSab("MOUVEMEVE")
    Case "COU":
        X = rsSab("LIBELLIB1")
        wDos = cmdSelect_SQL_2_Importation_YCRTMVT0_Libellé(X) / 100
    Case "ECH": wDos = rsSab("MOUVEMNUM") / 100
End Select

X = "SELECT CREEMPNCL , CLIENARSD FROM " & paramIBM_Library_SAB & ".ZCREEMP0 ," & paramIBM_Library_SAB & ".ZCLIENA0" _
  & " where CREEMPETA =" & rsSab("MOUVEMETA") & " and  CREEMPAGE =" & rsSab("MOUVEMAGE") _
  & " and CREEMPSER =" & rsSab("MOUVEMSER") & " and  CREEMPSSE =" & rsSab("MOUVEMSSE") _
  & " and CREEMPDOS =" & wDos & " and  CREEMPSEQ = 1" _
  & " and CLIENAETB = CREEMPETA and  CLIENACLI = CREEMPNCL"

Set rsSabX = cnsab.Execute(X)
If Not rsSabX.EOF Then
    newYCRTMVT0.CRTMVTCLIN = Trim(rsSabX("CREEMPNCL"))
    newYCRTMVT0.CRTMVTCLIP = Trim(rsSabX("CLIENARSD"))
End If
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:

End Sub


Public Sub cmdSelect_SQL_2_Importation_YCRTMVT0_ZTREOPE0()
Dim X As String, wDos As Long
On Error GoTo Error_Handler


wDos = rsSab("MOUVEMNUM")


X = "SELECT TREOPECLI , CLIENARSD FROM " & paramIBM_Library_SAB & ".ZTREOPE0 ," & paramIBM_Library_SAB & ".ZCLIENA0" _
  & " where TREOPEETB =" & rsSab("MOUVEMETA") & " and  TREOPEAGE =" & rsSab("MOUVEMAGE") _
  & " and TREOPESER = '" & rsSab("MOUVEMSER") & "' and  TREOPESES = '" & rsSab("MOUVEMSSE") & "'" _
  & "  and TREOPEOPR = '" & rsSab("MOUVEMOPE") & "' and TREOPENUM =" & wDos _
  & " and CLIENAETB = TREOPEETB and  CLIENACLI = TREOPECLI"

Set rsSabX = cnsab.Execute(X)
If Not rsSabX.EOF Then
    newYCRTMVT0.CRTMVTCLIN = Trim(rsSabX("TREOPECLI"))
    newYCRTMVT0.CRTMVTCLIP = Trim(rsSabX("CLIENARSD"))
End If
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:

End Sub

Public Sub cmdSelect_SQL_2_Importation_YCRTMVT0_ZCAUDOS0()
Dim X As String, wDos As Long, K As Integer, xMTD As String
On Error GoTo Error_Handler


wDos = 0 'rsSab("MOUVEMNUM")
If rsSab("MOUVEMOPE") = "AP1" Then
    wDos = rsSab("MOUVEMNUM")
Else
    Select Case rsSab("MOUVEMEVE")
        Case "PRO": wDos = rsSab("MOUVEMNUM")
        Case Else:
            X = Trim(rsSab("LIBELLIB1"))
            K = InStr(X, " 00")
            If K > 0 Then wDos = mId$(X, K + 1, Len(X) - K)
    End Select
End If

xMTD = cur_P(-rsSab("MOUVEMMON"))

X = "SELECT CAUCOHCLI , CLIENARSD FROM " & paramIBM_Library_SAB & ".ZCAUCOH0 ," & paramIBM_Library_SAB & ".ZCLIENA0" _
  & " where CAUCOHETA =" & rsSab("MOUVEMETA") & " and  CAUCOHAGE =" & rsSab("MOUVEMAGE") _
  & " and CAUCOHSER = '" & rsSab("MOUVEMSER") & "' and  CAUCOHSSE = '" & rsSab("MOUVEMSSE") & "'" _
  & "  and CAUCOHNDO =" & wDos & " and CAUCOHTRA = " & rsSab("MOUVEMDTR") & " and CAUCOHMNT =" & xMTD _
  & " and CLIENAETB = CAUCOHETA and  CLIENACLI = CAUCOHCLI"

Set rsSabX = cnsab.Execute(X)
If Not rsSabX.EOF Then
    newYCRTMVT0.CRTMVTCLIN = Trim(rsSabX("CAUCOHCLI"))
    newYCRTMVT0.CRTMVTCLIP = Trim(rsSabX("CLIENARSD"))
Else
    X = "SELECT CAUCOHCLI , CLIENARSD FROM " & paramIBM_Library_SAB & ".ZCAUCOH0 ," & paramIBM_Library_SAB & ".ZCLIENA0" _
      & " where CAUCOHETA =" & rsSab("MOUVEMETA") & " and  CAUCOHAGE =" & rsSab("MOUVEMAGE") _
      & " and CAUCOHSER = '" & rsSab("MOUVEMSER") & "' and  CAUCOHSSE = '" & rsSab("MOUVEMSSE") & "'" _
      & "  and CAUCOHNDO =" & wDos & " and CAUCOHMNT =" & xMTD _
      & " and CLIENAETB = CAUCOHETA and  CLIENACLI = CAUCOHCLI"
    
    Set rsSabX = cnsab.Execute(X)
    If Not rsSabX.EOF Then
        newYCRTMVT0.CRTMVTCLIN = Trim(rsSabX("CAUCOHCLI"))
        newYCRTMVT0.CRTMVTCLIP = Trim(rsSabX("CLIENARSD"))
    End If
End If
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:

End Sub
Public Sub cmdSelect_SQL_2_Importation_YCRTMVT0_ECH()
Dim X As String, wDos As Long, K As Integer
On Error GoTo Error_Handler


wDos = 0
X = Trim(rsSab("LIBELLIB1"))
K = InStr(X, "AGIOS DU ") + 9
If K > 9 Then
    wDos = mId$(X, K, 5)
Else
    K = InStr(X, "INT DEB: ") + 9
    If K > 9 Then
        wDos = mId$(X, K, 5)
    Else
        If X = "Frais tenue de compte du" Then wDos = mId$(Trim(rsSab("LIBELLIB2")), 1, 5)
    End If
End If
    
If wDos > 0 Then
        newYCRTMVT0.CRTMVTCLIC = " "
        newYCRTMVT0.CRTMVTCLIN = Format(wDos, "0000000")
        newYCRTMVT0.CRTMVTCLIP = sqlCRTMVTCLI_Pays(newYCRTMVT0.CRTMVTCLIC, newYCRTMVT0.CRTMVTCLIN)
End If
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:

End Sub


Public Sub cmdSelect_SQL_2_Importation_YCRTMVT0_ZCDODOS0(lMOUVEMOPE As String, lMOUVEMNUM As Long)
Dim X As String, wDos As Long
On Error GoTo Error_Handler

X = "SELECT CDODOSBER , CDODOSBEN FROM " & paramIBM_Library_SAB & ".ZCDODOS0" _
  & " where CDODOSETB = " & currentSAB_ETA & " and  CDODOSAGE = " & currentSAB_AGE _
  & " and CDODOSSER = '00' and  CDODOSSSE = '00'" _
  & " and CDODOSCOP = '" & lMOUVEMOPE & "' and CDODOSDOS = " & lMOUVEMNUM

Set rsSabX = cnsab.Execute(X)
If Not rsSabX.EOF Then
    newYCRTMVT0.CRTMVTCLIC = Trim(rsSabX("CDODOSBER"))
    If newYCRTMVT0.CRTMVTCLIC = "T" Then newYCRTMVT0.CRTMVTCLIC = "D"
    newYCRTMVT0.CRTMVTCLIN = Trim(rsSabX("CDODOSBEN"))
    newYCRTMVT0.CRTMVTCLIP = sqlCRTMVTCLI_Pays(newYCRTMVT0.CRTMVTCLIC, newYCRTMVT0.CRTMVTCLIN)
End If
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:

End Sub

Public Function cmdSelect_SQL_2_Importation_YCRTMVT0_Libellé(lX) As Long
Dim K As Integer, K2 As Integer

cmdSelect_SQL_2_Importation_YCRTMVT0_Libellé = 0
K = InStr(lX, "N°") + 3
If K > 3 Then
    For K2 = K To Len(lX)
        If Not IsNumeric(mId$(lX, K2, 1)) Then
            If mId$(lX, K2, 1) <> " " Then Exit For
        End If
    Next K2
    cmdSelect_SQL_2_Importation_YCRTMVT0_Libellé = Val(mId$(lX, K, K2 - K))
End If

End Function

Public Sub cmdSelect_SQL_2_Importation_YCRTMVT0_YDOSXOD0()
Dim X As String, K As Integer
Dim xDOSXODOPE As String, xDOSXODNUM As Long
    Dim blnNum As Boolean, wDos As Long
    blnNum = False: wDos = 0

On Error GoTo Error_Handler

X = "SELECT DOSXODOPE , DOSXODNUM FROM " & paramIBM_Library_SABSPE & ".YDOSXOD0 " _
  & " where DOSXODDTR =" & rsSab("MOUVEMDTR") + 19000000 & " and  DOSXODPIE =" & rsSab("MOUVEMPIE") & " and  DOSXODECR =" & rsSab("MOUVEMECR")
  
Set rsSabX = cnsab.Execute(X)
If Not rsSabX.EOF Then
    If Trim(rsSabX("DOSXODOPE")) <> "" Then
            newYCRTMVT0.CRTMVTOPE = rsSabX("DOSXODOPE")
            newYCRTMVT0.CRTMVTDOS = rsSabX("DOSXODNUM")
'=====================
            Exit Sub
'=====================
    End If
End If
'______________________________________________________________________________________
If rsSab("MOUVEMOPE") = "*G1" And rsSab("MOUVEMEVE") = "FCB" Then
    X = UCase$(Trim(rsSab("LIBELLIB2")))
    K = InStr(X, "COMPTE") + 6
    If K = 6 Then
        K = InStr(X, "CPTE") + 4
        If K = 4 Then
            K = InStr(X, " CPT") + 4
            If K = 4 Then
                K = InStr(X, " CB ") + 4
            End If
        End If
    End If
                
    
    
    If K > 4 Then
        For K = K To Len(X)
            If IsNumeric(mId$(X, K, 1)) Then
                blnNum = True
                wDos = wDos * 10 + Val(mId$(X, K, 1))
            Else
                If blnNum Then Exit For
            End If
        Next K
        newYCRTMVT0.CRTMVTCLIC = " "
        newYCRTMVT0.CRTMVTCLIN = Format(wDos, "0000000")
        newYCRTMVT0.CRTMVTCLIP = sqlCRTMVTCLI_Pays(newYCRTMVT0.CRTMVTCLIC, newYCRTMVT0.CRTMVTCLIN)
'=====================
            Exit Sub
'=====================
    End If
End If
'______________________________________________________________________________________

X = UCase$(Trim(rsSab("LIBELLIB1")) & Trim(rsSab("LIBELLIB2")))
K = InStr(X, "CDE")
If K > 0 Then
    For K = K + 3 To Len(X)
        If IsNumeric(mId$(X, K, 1)) Then
            blnNum = True
            wDos = wDos * 10 + Val(mId$(X, K, 1))
        Else
            If blnNum Then Exit For
        End If
    Next K
    newYCRTMVT0.CRTMVTOPE = "CDE"
    newYCRTMVT0.CRTMVTDOS = wDos
'=====================
            Exit Sub
'=====================
    
End If
'______________________________________________________________________________________
If rsSab("MOUVEMOPE") = "*B1" Then ' And rsSab("MOUVEMEVE") = "CHA" Then
    cmdSelect_SQL_2_Importation_YCRTMVT0_Contrepartie
    Exit Sub
End If
If rsSab("MOUVEMOPE") = "*T1" Then  ' And rsSab("MOUVEMEVE") = "PAY" Then
    cmdSelect_SQL_2_Importation_YCRTMVT0_Contrepartie
    Exit Sub
End If
If rsSab("MOUVEMOPE") = "*L1" Then '  And rsSab("MOUVEMEVE") = "001" Then
    cmdSelect_SQL_2_Importation_YCRTMVT0_Contrepartie
    Exit Sub
End If
If rsSab("MOUVEMOPE") = "*G1" And rsSab("MOUVEMEVE") = "EXT" Then
    cmdSelect_SQL_2_Importation_YCRTMVT0_Contrepartie
    Exit Sub
End If


GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:

End Sub

Public Sub lstParam_Load_Pays_Exclus(lFct As String)
Dim xSQL As String, K As Integer
fraParam_Display.Visible = False
lstParam_K.Clear

xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 where BASTABNUM = 11 " _
     & " and substring (BASTABDON , 22 , 1) in ( '1' , '2' ) order by BASTABARG"
Set rsSab = cnsab.Execute(xSQL)
mPays_Exclus = ""
Do While Not rsSab.EOF
    Select Case lFct
    
        Case "Display": lstParam_K.AddItem mId$(rsSab("BASTABARG"), 4, 2) & " | " & rsSab("BASTABLO1") & rsSab("BASTABLO2") & Trim(rsSab("BASTABDON"))

        Case "Init": mPays_Exclus = mPays_Exclus & mId$(rsSab("BASTABARG"), 4, 2) & "_"
        Case "Excel":
            mXls1_Row = mXls1_Row + 1
            wsExcel.Cells(mXls1_Row, 1) = "Pays_Exclus"
            wsExcel.Cells(mXls1_Row, 2) = mId$(rsSab("BASTABARG"), 4, 2)
            wsExcel.Cells(mXls1_Row, 3) = ""
            wsExcel.Cells(mXls1_Row, 4) = rsSab("BASTABLO1") & rsSab("BASTABLO2") & Trim(rsSab("BASTABDON"))
            For K = 1 To 1: wsExcel.Cells(mXls1_Row, K).Font.Color = vbMagenta: Next K
    End Select
    rsSab.MoveNext
Loop


End Sub
Public Sub lstParam_Load_Clients_Exclus(lFct As String)
Dim xSQL As String, K As Integer
fraParam_Display.Visible = False
lstParam_K.Clear

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 where CLIENAETA = 'AMBA' " _
     & " order by CLIENACLI"
Set rsSab = cnsab.Execute(xSQL)
mClients_Exclus = ""
Do While Not rsSab.EOF
    Select Case lFct
    
        Case "Display": lstParam_K.AddItem rsSab("CLIENACLI") & " | " & Trim(rsSab("CLIENARA1")) & Trim(rsSab("CLIENARA2"))

        Case "Init": mClients_Exclus = mClients_Exclus & rsSab("CLIENACLI") & "_"
        Case "Excel":
            mXls1_Row = mXls1_Row + 1
            wsExcel.Cells(mXls1_Row, 1) = "Clients_Exclus"
            wsExcel.Cells(mXls1_Row, 2) = rsSab("CLIENACLI")
            wsExcel.Cells(mXls1_Row, 3) = rsSab("CLIENAETA")
            wsExcel.Cells(mXls1_Row, 4) = Trim(rsSab("CLIENARA1")) & Trim(rsSab("CLIENARA2"))
            For K = 1 To 1: wsExcel.Cells(mXls1_Row, K).Font.Color = vbMagenta: Next K

    End Select
    rsSab.MoveNext
Loop


End Sub



Public Sub fraYTVACOM0_Display(lMOUVEMPIE As Long, lMOUVEMECR As Long)
Dim xSQL As String
Me.Enabled = False: Me.MousePointer = vbHourglass

fraYTVACOM0.Visible = False

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YTVACOM0 " _
     & " where TVACOMETA = " & currentSAB_ETA & " and TVACOMPLA = " & currentSAB_PLA _
     & " and TVACOMPIE = " & lMOUVEMPIE & " and TVACOMECR = " & lMOUVEMECR
Set rsSabX = cnsab.Execute(xSQL)

If Not rsSabX.EOF Then
    txtTVACOMCOMC = rsSabX("TVACOMCOMC")
    fraYTVACOM0.Visible = True
End If
Me.Enabled = True: Me.MousePointer = 0
End Sub

Public Sub cmdPrint_Excel_YCRTMVT0_Detail(lWhere As String, blnCRTMVTSTA_Color As Boolean)

Dim xSQL As String, X As String, K As Long
Dim curX As String, wColor As Long, mCRTMVTCPT As String

On Error GoTo Error_Handler

With wsExcel.Cells
    .HorizontalAlignment = Excel.xlHAlignLeft
    .Font.Size = 9
    .Font.Name = "Courier New"
End With
wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 75
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14BDF_CRT : Mouvements comptables" _
                                & "  (édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$E1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


wsExcel.Columns(1).ColumnWidth = 8: wsExcel.Cells(1, 1) = "Rubrique "
wsExcel.Columns(2).ColumnWidth = 18: wsExcel.Cells(1, 2) = "Compte"
wsExcel.Columns(3).ColumnWidth = 12: wsExcel.Cells(1, 3) = "D. TRT"
wsExcel.Columns(4).ColumnWidth = 25: wsExcel.Cells(1, 4) = "Référence"
wsExcel.Columns(5).ColumnWidth = 15: wsExcel.Cells(1, 5) = "CV Euro": wsExcel.Columns(5).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(6).ColumnWidth = 15: wsExcel.Cells(1, 6) = "Mt Devise": wsExcel.Columns(6).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(6).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(7).ColumnWidth = 4: wsExcel.Cells(1, 7) = "Dev"
wsExcel.Columns(8).ColumnWidth = 6: wsExcel.Cells(1, 8) = "Pays"
wsExcel.Columns(9).ColumnWidth = 12: wsExcel.Cells(1, 9) = "Client"
wsExcel.Columns(10).ColumnWidth = 30: wsExcel.Cells(1, 10) = "libellé"
wsExcel.Columns(11).ColumnWidth = 6: wsExcel.Cells(1, 11) = "Etat"
wsExcel.Columns(12).ColumnWidth = 10: wsExcel.Cells(1, 12) = "Pièce-Ecr"

mXls1_Cols = 12


For K = 1 To mXls1_Cols
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = vbWhite
Next
mXls1_Row = 1

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YCRTMVT0 , " _
     & paramIBM_Library_SABSPE & ".YBIAMVTHP" _
     & " where CRTMVTETA = MOUVEMETA and CRTMVTPIE = MOUVEMPIE and CRTMVTECR = MOUVEMECR " & lWhere _
     & " order by CRTMVTRUB , CRTMVTCPT , CRTMVTDTR , CRTMVTPIE , CRTMVTECR"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    If rsSab("CRTMVTCPT") <> mCRTMVTCPT Then
        mCRTMVTCPT = rsSab("CRTMVTCPT")
        Call lstErr_ChangeLastItem(lstErr, cmdContext, rsSab("CRTMVTRUB") & " " & mCRTMVTCPT): DoEvents
        If mXls1_Row > 1 Then mXls1_Row = mXls1_Row + 1
    End If
        
    mXls1_Row = mXls1_Row + 1
    wsExcel.Cells(mXls1_Row, 1) = rsSab("CRTMVTRUB")
    wsExcel.Cells(mXls1_Row, 2) = mCRTMVTCPT
    wsExcel.Cells(mXls1_Row, 3) = dateImp10(rsSab("CRTMVTDTR"))
    wsExcel.Cells(mXls1_Row, 4) = rsSab("MOUVEMSER") & " " & rsSab("MOUVEMSSE") & " " & rsSab("MOUVEMOPE") & " " & rsSab("MOUVEMEVE") & " " & rsSab("MOUVEMNUM")
    wsExcel.Cells(mXls1_Row, 5) = rsSab("CRTMVTMTE")
    wsExcel.Cells(mXls1_Row, 6) = -rsSab("MOUVEMMON")
    wsExcel.Cells(mXls1_Row, 7) = rsSab("CRTMVTDEV")
    wsExcel.Cells(mXls1_Row, 8) = rsSab("CRTMVTCLIP")
    If rsSab("CRTMVTCLIN") > 0 Then wsExcel.Cells(mXls1_Row, 9) = rsSab("CRTMVTCLIC") & rsSab("CRTMVTCLIN")
    wsExcel.Cells(mXls1_Row, 10) = Trim(rsSab("LIBELLIB1")) & Trim(rsSab("LIBELLIB2")) & Trim(rsSab("LIBELLIB3")) & Trim(rsSab("LIBELLIB4"))
    wsExcel.Cells(mXls1_Row, 11) = rsSab("CRTMVTORIG") & rsSab("CRTMVTSTA")
    wsExcel.Cells(mXls1_Row, 12) = rsSab("CRTMVTPIE") & "-" & rsSab("CRTMVTECR")
    X = rsSab("CRTMVTSTA")
    If X <> " " Then
        Select Case X
            Case "I": wColor = RGB(96, 96, 96)
                For K = 1 To 4: wsExcel.Cells(mXls1_Row, K).Font.Color = wColor: Next K
                For K = 7 To mXls1_Cols: wsExcel.Cells(mXls1_Row, K).Font.Color = wColor: Next K
            Case "?": wColor = vbMagenta 'mColor_W0
                If blnCRTMVTSTA_Color Then
                    For K = 1 To mXls1_Cols: wsExcel.Cells(mXls1_Row, K).Font.Color = wColor: Next K
                Else
                    wsExcel.Cells(mXls1_Row, 11).Interior.Color = mColor_W1
                End If
            Case Else: wColor = vbRed
                For K = 1 To 4: wsExcel.Cells(mXls1_Row, K).Font.Color = wColor: Next K
                For K = 7 To mXls1_Cols: wsExcel.Cells(mXls1_Row, K).Font.Color = wColor: Next K
        End Select
           
    End If
        
    rsSab.MoveNext
Loop
'======================================================================================================

Exit_sub:
'__________________________________________________________________________________


'_____________________________
Exit Sub

Error_Handler:
    If Not blnAuto Then Call MsgBox(Error, vbCritical, Me.Name)
End Sub

Public Sub cmdSelect_SQL_2_Importation_YCRTMVT0_Contrepartie()
Dim X As String

On Error GoTo Error_Handler


X = "SELECT CLIENACLI , CLIENARSD FROM " & paramIBM_Library_SABSPE & ".YBIAMVTHP , " & paramIBM_Library_SABSPE & ".YBIACPT0" _
  & " where MOUVEMETA =" & rsSab("MOUVEMETA") & " and  MOUVEMPIE =" & rsSab("MOUVEMPIE") _
  & " and COMPTECOM = MOUVEMCOM order by  MOUVEMECR"
  
Set rsSabX = cnsab.Execute(X)
Do While Not rsSabX.EOF
    If Trim(rsSabX("CLIENACLI")) <> "" Then
        newYCRTMVT0.CRTMVTCLIC = " "
        newYCRTMVT0.CRTMVTCLIN = rsSabX("CLIENACLI")
        newYCRTMVT0.CRTMVTCLIP = rsSabX("CLIENARSD")
        Exit Do
    End If
    rsSabX.MoveNext
Loop

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:

End Sub

Public Function fraOD_Control()
Dim X As String, curX As Currency

On Error GoTo Exit_sub
fraOD_Control = "?"


newYCRTMVT0.CRTMVTRUB = mId$(cboCRTMVTRUB_OD, 1, 5)
If Trim(newYCRTMVT0.CRTMVTRUB) = "" Then
    Call MsgBox("Préciser la rubrique CRT", vbInformation, "fraOD_Control")
    GoTo Exit_sub
End If

newYCRTMVT0.CRTMVTCLIP = mId$(cboCRTMVTCLIP_OD, 1, 2)
If Trim(newYCRTMVT0.CRTMVTCLIP) = "" Then
    Call MsgBox("Préciser le pays", vbInformation, "fraOD_Control")
    GoTo Exit_sub
End If

Call DTPicker_Control(txtCRTMVTDTR_OD, X)
newYCRTMVT0.CRTMVTDTR = X
If newYCRTMVT0.CRTMVTDTR > mAmjMax_Exercice Then
    Call MsgBox("Date > à la fin d'exercice", vbInformation, "fraOD_Control")
    GoTo Exit_sub
End If
If newYCRTMVT0.CRTMVTDTR < mAmjMin_Exercice Then
    Call MsgBox("Date < à la fin d'exercice", vbInformation, "fraOD_Control")
    GoTo Exit_sub
End If

If Trim(txtCRTMVTTXT_OD) = "" Then
    Call MsgBox("Préciser un commentaire", vbInformation, "fraOD_Control")
    GoTo Exit_sub
End If

X = Trim(txtCRTMVTCPT_OD)
newYCRTMVT0.CRTMVTCPT = X
libCRTMVTCPT_OD = ""
If X <> "" Then
    X = "select COMPTEINT from " & paramIBM_Library_SAB & ".ZCOMPTE0 " _
         & " where COMPTECOM = '" & X & "'"
    Set rsSab = cnsab.Execute(X)
    If Not rsSab.EOF Then
        libCRTMVTCPT_OD = rsSab("COMPTEINT")
    Else
        Call MsgBox("Compte inconnu", vbInformation, "fraOD_Control")
        GoTo Exit_sub
    End If

End If


newYCRTMVT0.CRTMVTMTE = Abs(num_CDec(txtCRTMVTMTE_OD))

If newYCRTMVT0.CRTMVTMTE = 0 Then
    Call MsgBox("Préciser le montant", vbInformation, "fraOD_Control")
    GoTo Exit_sub
End If


If optCRTMVTMTD_OD_Db Then newYCRTMVT0.CRTMVTMTE = -newYCRTMVT0.CRTMVTMTE
newYCRTMVT0.CRTMVTDEV = mId$(cboCRTMVTDEV_OD, 1, 3)
If Trim(newYCRTMVT0.CRTMVTDEV) = "" Then
    Call MsgBox("Préciser la devise", vbInformation, "fraOD_Control")
    GoTo Exit_sub
End If

fraOD_Control = Null

Exit_sub:

End Function

Public Sub cmdYCRTMVT0_OD_Transaction(lFct As String)
Dim V, X As String
Dim blnTransaction As Boolean
On Error GoTo Error_Handler


blnTransaction = False

V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
blnTransaction = True

    
newYCRTLOG0 = zYCRTLOG0
Select Case lFct
    Case "Add": newYCRTLOG0.CRTLOGNAT = "M01"
                V = sqlYCRTMVT0_Insert_OD(newYCRTMVT0)
    Case "Update": newYCRTLOG0.CRTLOGNAT = "M02"
            V = sqlYCRTMVT0_Update(newYCRTMVT0, oldYCRTMVT0)
    Case "Delete": newYCRTLOG0.CRTLOGNAT = "M03"
            V = sqlYCRTMVT0_Update(newYCRTMVT0, oldYCRTMVT0)
    Case Else
            V = lFct & " non programmé"
End Select

If Not IsNull(V) Then GoTo Error_MsgBox

newYCRTLOG0.CRTLOGCPT = newYCRTMVT0.CRTMVTCPT
newYCRTLOG0.CRTLOGETA = newYCRTMVT0.CRTMVTETA
newYCRTLOG0.CRTLOGPLA = newYCRTMVT0.CRTMVTPLA
newYCRTLOG0.CRTLOGPIE = newYCRTMVT0.CRTMVTPIE
newYCRTLOG0.CRTLOGECR = newYCRTMVT0.CRTMVTECR
newYCRTLOG0.CRTLOGCPT = newYCRTMVT0.CRTMVTCPT
X = ""
If newYCRTMVT0.CRTMVTDEV <> oldYCRTMVT0.CRTMVTDEV Then
    X = X & "<CRTMVTDEV = " & oldYCRTMVT0.CRTMVTDEV & " | " & newYCRTMVT0.CRTMVTDEV & ">"
End If
If newYCRTMVT0.CRTMVTCLIC <> oldYCRTMVT0.CRTMVTCLIC Then
    X = X & "<CRTMVTCLIC = " & oldYCRTMVT0.CRTMVTCLIC & " | " & newYCRTMVT0.CRTMVTCLIC & ">"
End If
If newYCRTMVT0.CRTMVTCLIN <> oldYCRTMVT0.CRTMVTCLIN Then
    X = X & "<CRTMVTCLIN = " & oldYCRTMVT0.CRTMVTCLIN & " | " & newYCRTMVT0.CRTMVTCLIN & ">"
End If
If newYCRTMVT0.CRTMVTCLIP <> oldYCRTMVT0.CRTMVTCLIP Then
    X = X & "<CRTMVTCLIP = " & oldYCRTMVT0.CRTMVTCLIP & " | " & newYCRTMVT0.CRTMVTCLIP & ">"
End If
If newYCRTMVT0.CRTMVTRUB <> oldYCRTMVT0.CRTMVTRUB Then
    X = X & "<CRTMVTRUB = " & oldYCRTMVT0.CRTMVTRUB & " | " & newYCRTMVT0.CRTMVTRUB & ">"
End If
If newYCRTMVT0.CRTMVTMTE <> oldYCRTMVT0.CRTMVTMTE Then
    X = X & "<CRTMVTMTE = " & oldYCRTMVT0.CRTMVTMTE & " | " & newYCRTMVT0.CRTMVTMTE & ">"
End If
If newYCRTMVT0.CRTMVTSTA <> oldYCRTMVT0.CRTMVTSTA Then
    X = X & "<CRTMVTSTA = " & oldYCRTMVT0.CRTMVTSTA & " | " & newYCRTMVT0.CRTMVTSTA & ">"
End If
If newYCRTMVT0.CRTMVTORIG <> oldYCRTMVT0.CRTMVTORIG Then
    X = X & "<CRTMVTORIG = " & oldYCRTMVT0.CRTMVTORIG & " | " & newYCRTMVT0.CRTMVTORIG & ">"
End If

X = X & "<CRTMVMTD = " & Trim(txtCRTMVTMTD_OD) & ">"
X = X & "<CRTMVTXT = " & Trim(txtCRTMVTTXT_OD) & ">"

newYCRTLOG0.CRTLOGTXT = X
V = sqlYCRTLOG0_Insert(newYCRTLOG0)
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If blnTransaction Then
        If Not IsNull(V) Then
            V = cnSAB_Transaction("Rollback")
        Else
            V = cnSAB_Transaction("Commit")
            If IsNull(V) Then
                fraOD.Visible = False
                cmdSelect_SQL_2
            End If

        End If
    End If

End Sub

Public Sub fraOD_Display_Cours()
If blnCRTMVTTAUX_OD Then
    Dim xDEV As String, xDTR As String, X As String
    xDEV = mId$(cboCRTMVTDEV_OD, 1, 3)
    Call DTPicker_Control(txtCRTMVTDTR_OD, xDTR)
    
    mCRTMVTTAUX_OD = 1
    If xDEV <> "EUR" Then
        Call sqlYBIATAB0_Read("PDC", xDEV, xDTR, X)
        If IsNumeric(mId$(X, 9, 15)) Then mCRTMVTTAUX_OD = CDbl(mId$(X, 9, 15) / 1000000000)
    End If
    txtCRTMVTTAUX_OD = Format(mCRTMVTTAUX_OD, "# ###.### ##")
    If Trim(txtCRTMVTMTD_OD) <> "" Then fraOD_Display_CRTMVTMTE
End If
End Sub

Public Sub cmdPrint_Excel_Déclaration_Detail(lFct As String, kNomenclature As Integer, kCRTMVTDEV As Integer, lCRTMVTRUB As String, lCRTMVTDEV As String, lCRTMVTCLIP As String, lcurDB As Currency, lcurCR As Currency)

'Print #lFile_No, "                      <Language>FR</Language>"
'Print #lFile_No, "               </Response>"

Dim curX As Currency, wFile_No As Integer

If kCRTMVTDEV = 0 Then GoTo Exit_sub
'==================================

wsExcel.Cells(mXls1_Row, 2 + 2 * kCRTMVTDEV) = Round(lcurDB, 0)
wsExcel.Cells(mXls1_Row, 2 + 2 * kCRTMVTDEV + 1) = Round(lcurCR, 0)

If cmdSelect_SQL_K = "CRT.xml" Then
    If kNomenclature_xml <> kNomenclature Then
        If kNomenclature_xml <> 0 Then Print #kNomenclature_File_No, "               </Data>"
        kNomenclature_xml = kNomenclature
        Select Case kNomenclature
            Case 1:
                    kNomenclature_File_No = 1
                    Print #kNomenclature_File_No, "               <Data form=" & Asc34 & "HPFDRES" & Asc34 & ">"
                    
            Case 2:
                    kNomenclature_File_No = 1
                    Print #kNomenclature_File_No, "               <Data form=" & Asc34 & "HPFD" & Asc34 & ">"
            Case 3:
                    kNomenclature_File_No = 3
                    Print #kNomenclature_File_No, "               <Data form=" & Asc34 & "PFD" & Asc34 & ">"
        End Select
    End If
    
    curX = Round(Abs(lcurDB) / 1000, 0)
    If curX <> 0 Then
        Print #kNomenclature_File_No, "                      <Item>"
        Print #kNomenclature_File_No, "                             <Dim prop=" & Asc34 & "MONNAIE" & Asc34 & ">" & lCRTMVTDEV & "</Dim>"
        Print #kNomenclature_File_No, "                             <Dim prop=" & Asc34 & "PAYS_CTPT" & Asc34 & ">" & lCRTMVTCLIP & "</Dim>"
        Print #kNomenclature_File_No, "                             <Dim prop=" & Asc34 & "CODE_ECO" & Asc34 & ">" & lCRTMVTRUB & "</Dim>"
        Print #kNomenclature_File_No, "                             <Dim prop=" & Asc34 & "MTT_TRSCT" & Asc34 & ">" & curX & "</Dim>"
        Print #kNomenclature_File_No, "                             <Dim prop=" & Asc34 & "SENS_TRSCT" & Asc34 & ">2</Dim>"
        Print #kNomenclature_File_No, "                             <Dim prop=" & Asc34 & "ANNEE_REF" & Asc34 & ">" & mExercice & "</Dim>"
        Print #kNomenclature_File_No, "                             <Dim prop=" & Asc34 & "TYPE_FLUX" & Asc34 & ">2</Dim>"
        Print #kNomenclature_File_No, "                      </Item>"
    End If
    
    curX = Round(Abs(lcurCR) / 1000, 0)
    If curX <> 0 Then
        Print #kNomenclature_File_No, "                      <Item>"
        Print #kNomenclature_File_No, "                             <Dim prop=" & Asc34 & "MONNAIE" & Asc34 & ">" & lCRTMVTDEV & "</Dim>"
        Print #kNomenclature_File_No, "                             <Dim prop=" & Asc34 & "PAYS_CTPT" & Asc34 & ">" & lCRTMVTCLIP & "</Dim>"
        Print #kNomenclature_File_No, "                             <Dim prop=" & Asc34 & "CODE_ECO" & Asc34 & ">" & lCRTMVTRUB & "</Dim>"
        Print #kNomenclature_File_No, "                             <Dim prop=" & Asc34 & "MTT_TRSCT" & Asc34 & ">" & curX & "</Dim>"
        Print #kNomenclature_File_No, "                             <Dim prop=" & Asc34 & "SENS_TRSCT" & Asc34 & ">1</Dim>"
        Print #kNomenclature_File_No, "                             <Dim prop=" & Asc34 & "ANNEE_REF" & Asc34 & ">" & mExercice & "</Dim>"
        Print #kNomenclature_File_No, "                             <Dim prop=" & Asc34 & "TYPE_FLUX" & Asc34 & ">2</Dim>"
        Print #kNomenclature_File_No, "                      </Item>"
    End If
    

End If

Exit_sub:
'=========

lcurDB = 0: lcurCR = 0
If lFct = "F" And cmdSelect_SQL_K = "CRT.xml" Then
    If kNomenclature_xml <> 0 Then Print #kNomenclature_File_No, "               </Data>"
End If

End Sub

Public Sub cmdSelect_SQL_Clôture()
Dim V, X As String


X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'CRT_AAAA'"
Set rsSab = cnsab.Execute(X)

If Not rsSab.EOF Then
    Call rsYBIATAB0_GetBuffer(rsSab, Old_YBIATAB0)
    X = MsgBox("Avez-vous archivé les fichiers .xls et .xml de l'exercice " & mExercice & " ?", vbQuestion & vbYesNo, "CRT : déclaration terminée")
    If X = vbYes Then
        X = MsgBox("Confirmez-vous la clôture définitive de l'exercice " & mExercice & " ?", vbQuestion & vbYesNo, "CRT : déclaration terminée")
        If X = vbYes Then
            New_YBIATAB0 = Old_YBIATAB0
            Mid$(New_YBIATAB0.BIATABTXT, 1, 4) = mExercice + 1
            V = Parametrage_Update
            
            If IsNull(V) Then
    '____________________________________________________________________________________________________________
                newYCRTLOG0 = zYCRTLOG0
                newYCRTLOG0.CRTLOGNAT = "D10"
                newYCRTLOG0.CRTLOGTXT = "<AAAA = " & mExercice & "|" & mExercice + 1 & ">"
                Call sqlYCRTLOG0_Insert_Transaction(newYCRTLOG0)
            End If
        End If
    End If
End If
End Sub
