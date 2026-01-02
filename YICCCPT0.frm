VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmYICCCPT0 
   AutoRedraw      =   -1  'True
   Caption         =   "Créances rattachées"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13575
   Icon            =   "YICCCPT0.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10530
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
      Height          =   9852
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   17383
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Créances rattachées"
      TabPicture(0)   =   "YICCCPT0.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Paramétrage"
      TabPicture(1)   =   "YICCCPT0.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstW"
      Tab(1).Control(1)=   "fgParam_GRP"
      Tab(1).Control(2)=   "fgParam_PCI"
      Tab(1).Control(3)=   "fgParam_CPT"
      Tab(1).ControlCount=   4
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
         Height          =   1860
         Left            =   -67320
         TabIndex        =   17
         Top             =   6360
         Visible         =   0   'False
         Width           =   4212
      End
      Begin VB.Frame fraTab0 
         Height          =   9420
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   13296
         Begin VB.Frame fraDetail 
            BackColor       =   &H00E0FFFF&
            Height          =   972
            Left            =   4200
            TabIndex        =   19
            Top             =   1680
            Visible         =   0   'False
            Width           =   7932
            Begin VB.Label libICCCPTGRP_lib 
               BackColor       =   &H00E0FFFF&
               Caption         =   "groupe libellé"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   1080
               TabIndex        =   25
               Top             =   240
               Width           =   3252
            End
            Begin VB.Label libICCCPTGRP 
               BackColor       =   &H00E0FFFF&
               Caption         =   "groupe"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Width           =   972
            End
            Begin VB.Label libCOMPTEOBL 
               BackColor       =   &H00E0FFFF&
               Caption         =   "compteobl"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   120
               TabIndex        =   22
               Top             =   600
               Width           =   972
            End
            Begin VB.Label libICCMVTCOM 
               BackColor       =   &H00E0FFFF&
               Caption         =   "libellé"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3360
               TabIndex        =   21
               Top             =   600
               Width           =   3972
            End
            Begin VB.Label lblICCMVTCOM 
               BackColor       =   &H00E0FFFF&
               Caption         =   "compte"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   1080
               TabIndex        =   20
               Top             =   600
               Width           =   2172
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
            Left            =   9240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   240
            Width           =   3732
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   555
            Left            =   10440
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   720
            Width           =   1335
         End
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00F0FFFF&
            Height          =   1212
            Left            =   360
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   8712
            Begin VB.Frame fraSelect_Options_1 
               BackColor       =   &H00F0FFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   7.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   972
               Left            =   120
               TabIndex        =   11
               Top             =   120
               Width           =   6732
               Begin VB.TextBox txtSelect_ICCMVTDOS 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Left            =   5400
                  TabIndex        =   29
                  Top             =   480
                  Width           =   1332
               End
               Begin VB.ComboBox txtSelect_ICCMVTOPE 
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
                  Left            =   5400
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   28
                  Top             =   120
                  Width           =   972
               End
               Begin VB.ComboBox cboSelect_ICCCPTGRP 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   312
                  Left            =   0
                  Sorted          =   -1  'True
                  TabIndex        =   15
                  Text            =   "grp"
                  Top             =   480
                  Width           =   1776
               End
               Begin VB.TextBox txtSelect_ICCCPTCOM 
                  BeginProperty Font 
                     Name            =   "@Arial Unicode MS"
                     Size            =   7.5
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Left            =   2040
                  TabIndex        =   12
                  Top             =   480
                  Width           =   1812
               End
               Begin VB.Label lblSelect_ICCMVTOPE 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "dossier"
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
                  Left            =   4680
                  TabIndex        =   30
                  Top             =   120
                  Width           =   612
               End
               Begin VB.Label lblSelect_ICCCPTGRP 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "groupe"
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
                  Left            =   480
                  TabIndex        =   14
                  Top             =   120
                  Width           =   732
               End
               Begin VB.Label lblSelect_ICCCPTCOM 
                  BackColor       =   &H00F0FFFF&
                  Caption         =   "compte"
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
                  Left            =   2520
                  TabIndex        =   13
                  Top             =   120
                  Width           =   612
               End
            End
            Begin MSComCtl2.DTPicker txtSelect_ICCMVTAMJ_Max 
               Height          =   300
               Left            =   7200
               TabIndex        =   9
               Top             =   840
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
               Format          =   32636931
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_ICCMVTAMJ_Min 
               Height          =   300
               Left            =   7200
               TabIndex        =   18
               Top             =   396
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
               Format          =   32636931
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_ICCCPTAMJ 
               BackColor       =   &H00F0FFFF&
               Caption         =   "date de situation"
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
               Left            =   7200
               TabIndex        =   10
               Top             =   120
               Width           =   1212
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7668
            Left            =   360
            TabIndex        =   5
            Top             =   1560
            Visible         =   0   'False
            Width           =   3312
            _ExtentX        =   5847
            _ExtentY        =   13520
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   -2147483633
            ForeColor       =   12582912
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483637
            BackColorSel    =   12648384
            BackColorBkg    =   -2147483633
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   "> Groupe|<Dev  |<Compte                          |<utilisateur     |<mise à jour              ||"
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
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   6420
            Left            =   3720
            TabIndex        =   16
            Top             =   2760
            Visible         =   0   'False
            Width           =   9372
            _ExtentX        =   16536
            _ExtentY        =   11324
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            BackColor       =   16777215
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483633
            BackColorBkg    =   -2147483633
            FormatString    =   $"YICCCPT0.frx":0342
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
      Begin MSFlexGridLib.MSFlexGrid fgParam_GRP 
         Height          =   6828
         Left            =   -74640
         TabIndex        =   23
         Top             =   1080
         Width           =   2952
         _ExtentX        =   5212
         _ExtentY        =   12039
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   16777215
         ForeColor       =   12582912
         BackColorFixed  =   16711935
         ForeColorFixed  =   -2147483637
         BackColorSel    =   12648384
         BackColorBkg    =   16777215
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   3
         FormatString    =   " GRP|<Libellé                                       "
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
      Begin MSFlexGridLib.MSFlexGrid fgParam_PCI 
         Height          =   6828
         Left            =   -70560
         TabIndex        =   26
         Top             =   1080
         Width           =   3432
         _ExtentX        =   6059
         _ExtentY        =   12039
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   16777215
         ForeColor       =   12582912
         BackColorFixed  =   16711935
         ForeColorFixed  =   -2147483637
         BackColorSel    =   12648384
         BackColorBkg    =   16777215
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   3
         FormatString    =   ">PCI   de       |>PCI   à       |***     |>GRP   "
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
      Begin MSFlexGridLib.MSFlexGrid fgParam_CPT 
         Height          =   6828
         Left            =   -65880
         TabIndex        =   27
         Top             =   1080
         Width           =   3672
         _ExtentX        =   6482
         _ExtentY        =   12039
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   16777215
         ForeColor       =   12582912
         BackColorFixed  =   16711935
         ForeColorFixed  =   -2147483637
         BackColorSel    =   12648384
         BackColorBkg    =   16777215
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
         GridLinesFixed  =   1
         AllowUserResizing=   3
         FormatString    =   "Etat|<Compte                        |>Dev    |>GRP    "
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
      Picture         =   "YICCCPT0.frx":03D3
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
      Begin VB.Menu mnuPrint_Recap 
         Caption         =   "Etat récapitulatif"
      End
      Begin VB.Menu mnuPrint_Detail 
         Caption         =   "Etat détaillé"
      End
      Begin VB.Menu mnuExportation 
         Caption         =   "Exportation .xlsx"
      End
   End
   Begin VB.Menu mnuParam_GRP 
      Caption         =   "mnuParam_GRP"
      Visible         =   0   'False
      Begin VB.Menu mnuParam_GRP_Update 
         Caption         =   "GRP : modifier le libellé"
      End
      Begin VB.Menu mnuParam_GRP_New 
         Caption         =   "GRP : ajouter un groupe"
      End
      Begin VB.Menu mnuParam_GRP_Delete 
         Caption         =   "GRP : supprimer un groupe"
      End
   End
   Begin VB.Menu mnuParam_PCI 
      Caption         =   "mnuParam_PCI"
      Visible         =   0   'False
      Begin VB.Menu mnuParam_PCI_Update 
         Caption         =   "PCI : modifier une plage"
      End
      Begin VB.Menu mnuParam_PCI_New 
         Caption         =   "PCI : ajouter une plage "
      End
      Begin VB.Menu mnuParam_PCI_Delete 
         Caption         =   "PCI : supprimer une plage"
      End
   End
   Begin VB.Menu mnuParam_CPT 
      Caption         =   "mnuParam_CPT"
      Visible         =   0   'False
      Begin VB.Menu mnuParam_CPT_Update 
         Caption         =   "CPT : modifier le groupe"
      End
      Begin VB.Menu mnuParam_CPT_New 
         Caption         =   "CPT : ajouter 1 compte"
      End
      Begin VB.Menu mnuParam_CPT_Ignore 
         Caption         =   "CPT : Ignorer ce compte"
      End
   End
End
Attribute VB_Name = "frmYICCCPT0"
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
Dim YICCCPT0_Aut As typeAuthorization
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
Dim xYICCCPT0 As typeYICCCPT0, newYICCCPT0 As typeYICCCPT0, oldYICCCPT0 As typeYICCCPT0
Dim arrYICCCPT0() As typeYICCCPT0, arrYICCCPT0_Nb As Long, arrYICCCPT0_Max As Long, arrYICCCPT0_Index As Long

Dim xYICCMVT0 As typeYICCMVT0, newYICCMVT0 As typeYICCMVT0, oldYICCMVT0 As typeYICCMVT0
Dim arrYICCMVT0() As typeYICCMVT0, arrYICCMVT0_Nb As Long, arrYICCMVT0_Max As Long, arrYICCMVT0_Index As Long

Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean
Dim xFF As String

Dim sSD_1 As Currency, sRBT As Currency
Dim sSD As Currency, sPRO As Currency

Dim arrICCCPTGRP_Lib(100) As String
Dim arrYICCCPT0_Param() As typeYBIATAB0, arrYICCCPT0_Param_NB As Integer
Dim New_YBIATAB0 As typeYBIATAB0, Old_YBIATAB0 As typeYBIATAB0
Dim blnAvance As Boolean


Dim fgParam_GRP_FormatString As String, fgParam_GRP_K As Integer
Dim fgParam_GRP_RowDisplay As Integer, fgParam_GRP_RowClick As Integer, fgParam_GRP_ColClick As Integer
Dim fgParam_GRP_ColorClick As Long, fgParam_GRP_ColorDisplay As Long
Dim fgParam_GRP_Sort1 As Integer, fgParam_GRP_Sort2 As Integer
Dim fgParam_GRP_SortAD As Integer, fgParam_GRP_Sort1_Old As Integer
Dim fgParam_GRP_arrIndex As Integer
Dim blnfgParam_GRP_DisplayLine As Boolean

Dim fgParam_CPT_FormatString As String, fgParam_CPT_K As Integer
Dim fgParam_CPT_RowDisplay As Integer, fgParam_CPT_RowClick As Integer, fgParam_CPT_ColClick As Integer
Dim fgParam_CPT_ColorClick As Long, fgParam_CPT_ColorDisplay As Long
Dim fgParam_CPT_Sort1 As Integer, fgParam_CPT_Sort2 As Integer
Dim fgParam_CPT_SortAD As Integer, fgParam_CPT_Sort1_Old As Integer
Dim fgParam_CPT_arrIndex As Integer
Dim blnfgParam_CPT_DisplayLine As Boolean

Dim fgParam_PCI_FormatString As String, fgParam_PCI_K As Integer
Dim fgParam_PCI_RowDisplay As Integer, fgParam_PCI_RowClick As Integer, fgParam_PCI_ColClick As Integer
Dim fgParam_PCI_ColorClick As Long, fgParam_PCI_ColorDisplay As Long
Dim fgParam_PCI_Sort1 As Integer, fgParam_PCI_Sort2 As Integer
Dim fgParam_PCI_SortAD As Integer, fgParam_PCI_Sort1_Old As Integer
Dim fgParam_PCI_arrIndex As Integer
Dim blnfgParam_PCI_DisplayLine As Boolean

Dim mWhere_ICCMVTDOS As String
Dim rsSabX As New ADODB.Recordset

Private Sub cmdPrint_YICCCPT0_xlsManual(blnDetail As Boolean, ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim X As String, xSQL As String, I As Integer, K As Integer
Dim wAmj As String, xWhere As String
Dim soldeD As typeYICCMVT0, soldeF As typeYICCMVT0, Total As typeYICCMVT0
Dim blnXprt_Line As Boolean
Dim Nb_Detail As Long
Dim wColor As Long
Dim premierTitreSAV As String

If arrYICCCPT0_Nb = 0 Then Exit Sub
Me.Enabled = False: Me.MousePointer = vbHourglass
Call DTPicker_Control(txtSelect_ICCMVTAMJ_Min, wAMJMin)
Call DTPicker_Control(txtSelect_ICCMVTAMJ_Max, WAMJMax)
wAmj = dateElp("Jour", -1, wAMJMin)
fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
If blnDetail Then
    prtTitleText = "Etat détaillé des créances rattachées du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax) & " - " & arrICCCPTGRP_Lib(arrYICCCPT0(1).ICCCPTGRP)
    premierTitreSAV = prtTitleText
Else
    prtTitleText = "Etat récapitulatif des créances rattachées  du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax) & " - " & arrICCCPTGRP_Lib(arrYICCCPT0(1).ICCCPTGRP)
    premierTitreSAV = prtTitleText
End If
arrYICCCPT0(0) = arrYICCCPT0(1)
wsExcel.Cells(1, 4) = prtTitleText
For K = 1 To arrYICCCPT0_Nb
    xYICCCPT0 = arrYICCCPT0(K)
    If xYICCCPT0.ICCCPTGRP <> arrYICCCPT0(K - 1).ICCCPTGRP Then
        If blnDetail Then
            prtTitleText = "Etat détaillé des créances rattachées du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax) & " - " & arrICCCPTGRP_Lib(xYICCCPT0.ICCCPTGRP)
        Else
            prtTitleText = "Etat récapitulatif des créances rattachées  du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax) & " - " & arrICCCPTGRP_Lib(xYICCCPT0.ICCCPTGRP)
        End If
        Call prtYICCCPT0_Close_xlsManual(False, currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    End If
    '________________________________________________________________
    If InStr(xYICCCPT0.ICCCPTCOM, "38820") Then
        blnAvance = True
    Else
        blnAvance = False
    End If
    xWhere = " where ICCMVTETA  = " & currentZMNURUT0.MNURUTETB _
         & " and   ICCMVTAGE = " & currentZMNUUTI0.MNUUTIAGE _
         & " and ICCMVTCOM ='" & Trim(xYICCCPT0.ICCCPTCOM) & "'" _
         & " and ICCMVTAMJ >= " & wAmj & " and ICCMVTAMJ <= " & WAMJMax _
         & " order by ICCMVTAMJ , ICCMVTSER , ICCMVTSSE , ICCMVTOPE , ICCMVTEVE , ICCMVTDOS"
    Call arrYICCMVT0_SQL(xWhere)
    sSD_1 = 0: sPRO = 0
    sSD = 0: sRBT = 0
    rsYICCMVT0_Init Total
    soldeD = Total: soldeF = Total
    blnXprt_Line = True
    Nb_Detail = 0
    For I = 1 To arrYICCMVT0_Nb
        xYICCMVT0 = arrYICCMVT0(I)
        If xYICCMVT0.ICCMVTAMJ < wAMJMin Then
            If xYICCMVT0.ICCMVTSER = xFF Then soldeD = xYICCMVT0
        Else
            If xYICCMVT0.ICCMVTSER = xFF Then
                soldeF = xYICCMVT0
            Else
                Nb_Detail = Nb_Detail + 1
                Total.ICCMVTRBT = Total.ICCMVTRBT + xYICCMVT0.ICCMVTRBT
                Total.ICCMVTTDB = Total.ICCMVTTDB + xYICCMVT0.ICCMVTTDB
                Total.ICCMVTTCR = Total.ICCMVTTCR + xYICCMVT0.ICCMVTTCR
                Total.ICCMVTPRO = Total.ICCMVTPRO + xYICCMVT0.ICCMVTPRO
                '_____________________________________________________________________
                Select Case xYICCMVT0.ICCMVTOPE
                    Case "EMP", "EM1":
                        If xYICCMVT0.ICCMVTRBT <> xYICCMVT0.ICCMVTTDB Then soldeD.ICCMVTRBT = 3
                        If xYICCMVT0.ICCMVTPRO <> xYICCMVT0.ICCMVTTCR Then soldeD.ICCMVTPRO = 4
                    Case Else:
                        If xYICCMVT0.ICCMVTRBT <> xYICCMVT0.ICCMVTTCR Then soldeD.ICCMVTRBT = 1
                        If xYICCMVT0.ICCMVTPRO <> xYICCMVT0.ICCMVTTDB Then soldeD.ICCMVTPRO = 2
                End Select
                '____________________________________________________________________
            End If
        End If
    Next I
    Call prtYICCCPT0_Line_xlsManual(xYICCCPT0, Total, soldeD, soldeF, blnDetail, Nb_Detail, blnAvance, currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    If blnDetail Then
        For I = 1 To arrYICCMVT0_Nb
            xYICCMVT0 = arrYICCMVT0(I)
            If xYICCMVT0.ICCMVTAMJ < wAMJMin Then
            Else
                If xYICCMVT0.ICCMVTSER <> xFF Then
                    Call prtYICCCPT0_Line_Detail_xlsManual(xYICCMVT0, currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
                End If
            End If
        Next I
    End If
Next K
Call prtYICCCPT0_Close_xlsManual(True, currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
'restauration premier titre sauvegardé
wsExcel.Cells(1, 4) = premierTitreSAV
Me.Show
Me.Enabled = True: Me.MousePointer = 0
End Sub

'______________________________________________________________________
Private Sub fgSelect_Display()
Dim wColor As Long

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Row = 0

currentAction = "fgSelect_Display"
    
For I = 1 To arrYICCCPT0_Nb
         
    xYICCCPT0 = arrYICCCPT0(I)
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine I, True
    
Next I

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYICCCPT0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub arrYICCCPT0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYICCCPT0(101)
arrYICCCPT0_Max = 100: arrYICCCPT0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YICCCPT0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYICCCPT0_GetBuffer(rsSab, xYICCCPT0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYICCCPT0.fgselect_Display"
        '' Exit Sub
     Else
         arrYICCCPT0_Nb = arrYICCCPT0_Nb + 1
         If arrYICCCPT0_Nb > arrYICCCPT0_Max Then
             arrYICCCPT0_Max = arrYICCCPT0_Max + 100
             ReDim Preserve arrYICCCPT0(arrYICCCPT0_Max)
         End If
         
         arrYICCCPT0(arrYICCCPT0_Nb) = xYICCCPT0
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


Private Sub arrYICCMVT0_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYICCMVT0(101)
arrYICCMVT0_Max = 100: arrYICCMVT0_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YICCMVT0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYICCMVT0_GetBuffer(rsSab, xYICCMVT0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYICCMVT0.fgselect_Display"
        '' Exit Sub
     Else
         arrYICCMVT0_Nb = arrYICCMVT0_Nb + 1
         If arrYICCMVT0_Nb > arrYICCMVT0_Max Then
             arrYICCMVT0_Max = arrYICCMVT0_Max + 100
             ReDim Preserve arrYICCMVT0(arrYICCMVT0_Max)
         End If
         
         arrYICCMVT0(arrYICCMVT0_Nb) = xYICCMVT0
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
    fgDetail.Visible = False: fraDetail.Visible = False
    lstW.Visible = False
    cmdSelect_Ok.Visible = True
    cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, 2))
    Select Case cmdSelect_SQL_K
        Case "1":
            fraSelect_Options.Visible = True: fraSelect_Options_1.Visible = True
    End Select

End If

End Sub



Private Sub cmdSelect_SQL_1()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String, xSQL As String
Dim wCli As Long
Dim blnOk As Boolean

On Error GoTo Error_Handler

currentAction = "cmdYICCCPT0_SQL"
blnOk = False
xWhere = ""
mWhere_ICCMVTDOS = ""

If Trim(txtSelect_ICCMVTDOS) = "" And Trim(txtSelect_ICCMVTOPE) = "" Then
    xWhere = " where ICCCPTETA  = " & currentZMNURUT0.MNURUTETB _
         & " and   ICCCPTAGE = " & currentZMNUUTI0.MNUUTIAGE
    X = Trim(txtSelect_ICCCPTCOM)
    If X <> "" Then
        xWhere = xWhere & " and   ICCCPTCOM like'" & X & "%'"
    End If
    X = Trim(cboSelect_ICCCPTGRP)
    If X <> "" Then
        xWhere = xWhere & " and   ICCCPTGRP =" & Val(Mid$(X, 1, 2))
    End If
    arrYICCCPT0_SQL xWhere & " order by ICCCPTGRP , ICCCPTDEV , ICCCPTCOM"
    
    fgSelect_Display
Else
    If Trim(txtSelect_ICCMVTOPE) <> "" Then
        mWhere_ICCMVTDOS = " and ICCMVTOPE  = '" & Trim(txtSelect_ICCMVTOPE) & "'"
        If Trim(txtSelect_ICCMVTDOS) <> "" Then mWhere_ICCMVTDOS = mWhere_ICCMVTDOS & " and ICCMVTDOS  = " & Val(txtSelect_ICCMVTDOS)
    Else
        mWhere_ICCMVTDOS = " and ICCMVTDOS  = " & Val(txtSelect_ICCMVTDOS)
    End If
    
    arrYICCCPT0_Nb = 0
    ReDim arrYICCCPT0(150)
    xWhere = Replace(mWhere_ICCMVTDOS, "and", "where", 1, 1)
    xSQL = "select distinct ICCMVTCOM from " & paramIBM_Library_SABSPE & ".YICCMVT0 " & xWhere & " order by ICCMVTCOM"
    Set rsSab = cnsab.Execute(xSQL)
    
    Do While Not rsSab.EOF
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YICCCPT0 where ICCCPTCOM ='" & rsSab("ICCMVTCOM") & "'"
        Set rsSabX = cnsab.Execute(xSQL)

        If Not rsSabX.EOF Then
            V = rsYICCCPT0_GetBuffer(rsSabX, xYICCCPT0)
            arrYICCCPT0_Nb = arrYICCCPT0_Nb + 1
            arrYICCCPT0(arrYICCCPT0_Nb) = xYICCCPT0
        End If
        rsSab.MoveNext
    Loop
    fgSelect_Display
End If

If arrYICCCPT0_Nb = 1 Then
    oldYICCCPT0 = arrYICCCPT0(1)
    xYICCCPT0 = oldYICCCPT0
    fgDetail_Display
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub fgDetail_Display()
Dim wColor As Long
Dim xWhere As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String

On Error GoTo Error_Handler
fgDetail.Visible = False: fraDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.Row = 0

currentAction = "fgDetail_Display"
Call DTPicker_Control(txtSelect_ICCMVTAMJ_Min, wAMJMin)
Call DTPicker_Control(txtSelect_ICCMVTAMJ_Max, WAMJMax)
wAmj = dateElp("Jour", -1, wAMJMin)

lblICCMVTCOM = oldYICCCPT0.ICCCPTCOM
xWhere = "select COMPTEINT , COMPTEOBL from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " Where COMPTECOM = '" & oldYICCCPT0.ICCCPTCOM & "'"
Set rsSab = cnsab.Execute(xWhere)
If Not rsSab.EOF Then
    libICCMVTCOM = rsSab("COMPTEINT")
    libCOMPTEOBL = rsSab("COMPTEOBL")
Else
    libICCMVTCOM = "??????"
End If

libICCCPTGRP = oldYICCCPT0.ICCCPTGRP
libICCCPTGRP_lib = arrICCCPTGRP_Lib(oldYICCCPT0.ICCCPTGRP)

If InStr(oldYICCCPT0.ICCCPTCOM, "38820") Then
    blnAvance = True
Else
    blnAvance = False
End If



xWhere = " where ICCMVTETA  = " & currentZMNURUT0.MNURUTETB _
     & " and   ICCMVTAGE = " & currentZMNUUTI0.MNUUTIAGE _
     & " and ICCMVTCOM ='" & Trim(oldYICCCPT0.ICCCPTCOM) & "'" _
     & " and ICCMVTAMJ >= " & wAmj & " and ICCMVTAMJ <= " & WAMJMax _
     & mWhere_ICCMVTDOS _
     & " order by ICCMVTAMJ , ICCMVTSER , ICCMVTSSE , ICCMVTOPE , ICCMVTEVE , ICCMVTDOS"

Call arrYICCMVT0_SQL(xWhere)

sSD_1 = 0: sPRO = 0
sSD = 0: sRBT = 0

For I = 1 To arrYICCMVT0_Nb
         
    xYICCMVT0 = arrYICCMVT0(I)
    If xYICCMVT0.ICCMVTAMJ < wAMJMin Then
        If xYICCMVT0.ICCMVTSER = xFF Then
            fgDetail.Rows = fgDetail.Rows + 1
            fgDetail.Row = fgDetail.Rows - 1
            fgDetail_DisplayLine_Solde I
        End If
    Else
        fgDetail.Rows = fgDetail.Rows + 1
        fgDetail.Row = fgDetail.Rows - 1
         Select Case xYICCMVT0.ICCMVTSER
            Case Is = xFF: fgDetail_DisplayLine_Solde I
            Case Else: fgDetail_DisplayLine I
        End Select

        
    End If
    
Next I

fgDetail.Visible = True: fraDetail.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long, blnYICCCPT0 As Boolean)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim xSQL As String
On Error Resume Next

 Select Case xYICCCPT0.ICCCPTSTA
    Case Is = "A", "I": wColor = RGB(128, 128, 128)
    Case Else: wColor = RGB(64, 64, 128)
End Select

fgSelect.Col = 1: fgSelect.Text = xYICCCPT0.ICCCPTDEV
fgSelect.CellForeColor = wColor
fgSelect.Col = 2: fgSelect.Text = xYICCCPT0.ICCCPTCOM
fgSelect.CellForeColor = wColor
fgSelect.Col = 0: fgSelect.Text = xYICCCPT0.ICCCPTGRP
fgSelect.CellForeColor = wColor
fgSelect.Col = 3: fgSelect.Text = xYICCCPT0.ICCCPTUUSR
fgSelect.CellForeColor = wColor
fgSelect.Col = 4: fgSelect.Text = dateImp10(xYICCCPT0.ICCCPTUAMJ) & " " & timeImp8(xYICCCPT0.ICCCPTUHMS)
fgSelect.CellForeColor = wColor


fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
End Sub
Public Sub fgDetail_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long

On Error Resume Next
wColor = vbBlue: wColor_Row = vbWhite


fgDetail.Col = 0: fgDetail.Text = dateImp10(xYICCMVT0.ICCMVTAMJ)
fgDetail.CellFontBold = False
fgDetail.CellForeColor = wColor
fgDetail.CellBackColor = wColor_Row

fgDetail.Col = 1
fgDetail.Text = xYICCMVT0.ICCMVTSER & " " & xYICCMVT0.ICCMVTSSE & " " & xYICCMVT0.ICCMVTOPE & " " & xYICCMVT0.ICCMVTNAT
fgDetail.CellForeColor = wColor
fgDetail.CellBackColor = wColor_Row

fgDetail.Col = 2
fgDetail.Text = xYICCMVT0.ICCMVTEVE & " " & Format(xYICCMVT0.ICCMVTDOS, "### ###")
fgDetail.CellForeColor = wColor
fgDetail.CellBackColor = wColor_Row

fgDetail.Col = 4
If xYICCMVT0.ICCMVTTDB <> 0 Then
    fgDetail.Text = Format(xYICCMVT0.ICCMVTTDB, "### ### ### ###.00")
    fgDetail.CellForeColor = vbRed
End If
fgDetail.CellBackColor = wColor_Row

fgDetail.Col = 5
If xYICCMVT0.ICCMVTTCR <> 0 Then
    fgDetail.Text = Format(xYICCMVT0.ICCMVTTCR, "### ### ### ###.00")
    fgDetail.CellForeColor = vbBlue
End If
fgDetail.CellBackColor = wColor_Row
'_______________________________________________________________________________
fgDetail.Col = 3
If xYICCMVT0.ICCMVTRBT <> 0 Then fgDetail.Text = Format(xYICCMVT0.ICCMVTRBT, "### ### ### ###.00")
fgDetail.CellForeColor = wColor
sRBT = sRBT + xYICCMVT0.ICCMVTRBT

Select Case xYICCMVT0.ICCMVTOPE

    Case "EMP", "EM1":
        If xYICCMVT0.ICCMVTRBT = xYICCMVT0.ICCMVTTDB Then
            fgDetail.CellBackColor = RGB(255, 255, 230)
            fgDetail.Col = 4
            fgDetail.CellBackColor = RGB(255, 255, 230)
        Else
            fgDetail.CellBackColor = RGB(255, 180, 255)
            fgDetail.Col = 4
            fgDetail.CellBackColor = RGB(255, 180, 255)
        End If
    Case Else:
        If xYICCMVT0.ICCMVTRBT = xYICCMVT0.ICCMVTTCR Then
            fgDetail.CellBackColor = RGB(255, 255, 230)
            fgDetail.Col = 5
            fgDetail.CellBackColor = RGB(255, 255, 230)
        Else
            fgDetail.CellBackColor = RGB(255, 180, 255)
            fgDetail.Col = 5
            fgDetail.CellBackColor = RGB(255, 180, 255)
        End If
    End Select

fgDetail.Col = 6
If xYICCMVT0.ICCMVTPRO <> 0 Then fgDetail.Text = Format(xYICCMVT0.ICCMVTPRO, "### ### ### ###.00")

sPRO = sPRO + xYICCMVT0.ICCMVTPRO
fgDetail.CellForeColor = wColor

Select Case xYICCMVT0.ICCMVTOPE

    Case "EMP", "EM1":
        If xYICCMVT0.ICCMVTPRO = xYICCMVT0.ICCMVTTCR Then
            fgDetail.CellBackColor = RGB(230, 255, 230)
            fgDetail.Col = 5
            fgDetail.CellBackColor = RGB(230, 255, 230)
        Else
            fgDetail.CellBackColor = RGB(255, 220, 255)
            fgDetail.Col = 5
            fgDetail.CellBackColor = RGB(255, 220, 255)
        End If
    Case Else:
        If xYICCMVT0.ICCMVTPRO = xYICCMVT0.ICCMVTTDB Then
            fgDetail.CellBackColor = RGB(230, 255, 230)
            fgDetail.Col = 4
            fgDetail.CellBackColor = RGB(230, 255, 230)
        Else
            fgDetail.CellBackColor = RGB(255, 220, 255)
            fgDetail.Col = 4
            fgDetail.CellBackColor = RGB(255, 220, 255)
        End If

    End Select
fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex
End Sub

Public Sub fgDetail_DisplayLine_Solde(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long

On Error Resume Next
wColor = vbWhite: wColor_Row = RGB(0, 144, 144)


fgDetail.Col = 0: fgDetail.Text = dateImp10(xYICCMVT0.ICCMVTAMJ)
fgDetail.CellFontBold = True
fgDetail.CellForeColor = wColor
fgDetail.CellBackColor = wColor_Row

fgDetail.Col = 1
fgDetail.CellBackColor = wColor_Row

fgDetail.Col = 2
fgDetail.CellBackColor = wColor_Row

fgDetail.Col = 4
If xYICCMVT0.ICCMVTTDB <> 0 Then
    fgDetail.Text = Format(xYICCMVT0.ICCMVTTDB, "### ### ### ###.00")
    fgDetail.CellForeColor = wColor 'vbRed
End If
fgDetail.CellBackColor = wColor_Row
fgDetail.CellFontBold = True

fgDetail.Col = 5
If xYICCMVT0.ICCMVTTCR <> 0 Then
    fgDetail.Text = Format(xYICCMVT0.ICCMVTTCR, "### ### ### ###.00")
    fgDetail.CellForeColor = wColor 'vbBlue
End If
fgDetail.CellBackColor = wColor_Row
fgDetail.CellFontBold = True
'_______________________________________________________________________________
fgDetail.Col = 3
If xYICCMVT0.ICCMVTRBT <> 0 Then fgDetail.Text = Format(xYICCMVT0.ICCMVTRBT, "### ### ### ###.00")
fgDetail.CellBackColor = wColor_Row
fgDetail.CellForeColor = wColor
sRBT = sRBT + xYICCMVT0.ICCMVTRBT


fgDetail.Col = 6
If xYICCMVT0.ICCMVTPRO <> 0 Then fgDetail.Text = Format(xYICCMVT0.ICCMVTPRO, "### ### ### ###.00")
fgDetail.CellBackColor = wColor_Row

sPRO = sPRO + xYICCMVT0.ICCMVTPRO
fgDetail.CellForeColor = wColor

    If fgDetail.Row > 1 Then
        fgDetail.Col = 3
        fgDetail.Text = Format(sRBT, "### ### ### ###.00")
        fgDetail.CellForeColor = wColor 'vbBlue
    
        If Not blnAvance Then
            If sRBT = -sSD_1 Then
                fgDetail.CellBackColor = wColor_Row 'RGB(255, 255, 230)
            Else
                fgDetail.CellBackColor = RGB(255, 0, 255)
            End If
        End If
        
        fgDetail.Col = 6
        fgDetail.Text = Format(sPRO, "### ### ### ###.00")
        fgDetail.CellForeColor = wColor 'vbBlue
        If Not blnAvance Then
            If sPRO = xYICCMVT0.ICCMVTTDB + xYICCMVT0.ICCMVTTCR Then
                fgDetail.CellBackColor = wColor_Row 'RGB(230,255,230)
            Else
                fgDetail.CellBackColor = RGB(255, 0, 255)
            End If
        End If
    End If
    sRBT = 0: sPRO = 0
    sSD_1 = xYICCMVT0.ICCMVTTDB + xYICCMVT0.ICCMVTTCR
   
fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex
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



'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(Msg As String)
Dim wbExcel As Excel.Workbook
Dim nbSheetRows As Long
Dim currentRow As Long
Dim maxRows As Long
Dim maxRowsPlus As Long
Dim comptageRows As Long
'---------------------------------------------------------
Dim wFct As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

wFct = UCase$(Trim(Mid$(Msg, 1, 12)))
Call BiaPgmAut_Init(wFct, YICCCPT0_Aut)

'blnSetfocus = True
Form_Init


Select Case wFct
    Case "@ICC_MVT":
                    If xlsManual Then
                        Call init_xlsManual
                        'On recopie le classeur modèle de c:\BIASRV vers c:\temp\imp_pdf
                        FileCopy paramFolder_Local & "\Modeles\modele_ICC_MVT.xlsx", paramIMP_PDF_Path_Temp & "\modele_ICC_MVT.xlsx"
                        'on charge CE classeur dans Excel
                        Call appExcelPublic.Workbooks.Open(paramIMP_PDF_Path_Temp & "\modele_ICC_MVT.xlsx")
                        Set wbExcel = appExcelPublic.ActiveWorkbook
                        With wbExcel
                            .Title = "ICC_MVT"
                            .Subject = "ICC_MVT"
                        End With
                        '                                               '
                        wbExcel.Worksheets(1).Activate
                        currentRow = 6
                        comptageRows = currentRow
                        maxRows = 33
                        maxRowsPlus = 4
                    End If
                    blnAuto = True
                    If Not IsEmpty(XPrt) Then Set XPrt_Previous = XPrt
                    '$jpl 2014-10-10 Printer_PDF
                    Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S60", "BIA-ICC-MVT-Recap", "Archive")
                    cmdSelect_SQL_1
                    If xlsManual Then
                        Call cmdPrint_YICCCPT0_xlsManual(False, currentRow, wbExcel.Sheets(1), comptageRows, maxRows, maxRowsPlus)
                        'on supprime les 4 lignes modèles
                        Rows("3:6").Select
                        Selection.Delete
                        currentRow = currentRow - 4
                        wbExcel.Worksheets(1).Cells(currentRow + 1, 1) = "END_OF_SHEET"
                        nbSheetRows = retourne_fin_de_sheet(wbExcel.Worksheets(1))
                        Call zoneImpression_xlsManual(wbExcel.Worksheets(1).Name, nbSheetRows, wbExcel.Worksheets(1))
                        Call wbExcel.Worksheets(1).ExportAsFixedFormat(xlTypePDF, paramIMP_PDF_Path & "\" & paramEditionNoPaper_Auto_PgmName & ".pdf")
                        'sauvegarde du fichier
                        Call impressions_xlsManual.prtIMP_PDF_Monitor_xlsManual
                    Else
                        cmdPrint_YICCCPT0 False
                    End If
                    Call lstErr_AddItem(lstErr, cmdContext, "Temporisation 5 secondes ...."): DoEvents
                    Wait_SS 5
                    'cmdSendMail_ICC_MVT False
                    '/////////////////////////////// détail ////////////////////////////////////////////
                    If xlsManual Then
                        Call init_xlsManual
                        wbExcel.Worksheets(2).Activate
                        currentRow = 6
                        comptageRows = currentRow
                        maxRows = 33
                        maxRowsPlus = 4
                    End If
                    Call frmElpPrt.prtIMP_PDF_NoPaper_Init("S60", "BIA-ICC-MVT-Detail", "Archive")
                    If xlsManual Then
                        Call cmdPrint_YICCCPT0_xlsManual(True, currentRow, wbExcel.Sheets(2), comptageRows, maxRows, maxRowsPlus)
                        'on supprime les 4 lignes modèles
                        Rows("3:6").Select
                        Selection.Delete
                        currentRow = currentRow - 4
                        wbExcel.Worksheets(2).Cells(currentRow + 1, 1) = "END_OF_SHEET"
                        nbSheetRows = retourne_fin_de_sheet(wbExcel.Worksheets(2))
                        Call zoneImpression_xlsManual(wbExcel.Worksheets(2).Name, nbSheetRows, wbExcel.Worksheets(2))
                        Call wbExcel.Worksheets(2).ExportAsFixedFormat(xlTypePDF, paramIMP_PDF_Path & "\" & paramEditionNoPaper_Auto_PgmName & ".pdf")
                        'sauvegarde du fichier
                        Call impressions_xlsManual.prtIMP_PDF_Monitor_xlsManual
                        Call wbExcel.Close(True)
                        Set wbExcel = Nothing
                        Kill paramIMP_PDF_Path_Temp & "\modele_ICC_MVT.xlsx"
                        If Not appExcelPublic Is Nothing Then
                            appExcelPublic.Quit
                            Set appExcelPublic = Nothing
                        End If
                        Dim fic As Long
                        fic = FreeFile
                        Open "c:\temp\imp_pdf\BIA_SAB2008.log" For Append As #fic
                        Print #fic, "Fin TIMER_BIA_SAB2008 --> " & CDate(Now)
                        Close #fic
                    Else
                        cmdPrint_YICCCPT0 True
                    End If
                    Call lstErr_AddItem(lstErr, cmdContext, "Temporisation 5 secondes ...."): DoEvents
                    Wait_SS 5
                    'cmdSendMail_ICC_MVT True
                    If Not IsEmpty(XPrt_Previous) Then Set XPrt = XPrt_Previous
                    Unload Me
    Case Else: blnAuto = False
End Select

End Sub
Private Sub zoneImpression_xlsManual(lFct As String, nbRows As Long, wsheet As Excel.Worksheet)

    Call init_TypePagesetup
    If nbRows > 0 Then
        wsheet.Activate
        wsheet.Range("A1:J" & CStr(nbRows)).Select
        zoneImpressionPagesetup.PrintArea = "$A$1:$J$" & CStr(nbRows)
        zoneImpressionPagesetup.LeftFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "prtYICCCPT0   &D &T  BIA_INFO"
        zoneImpressionPagesetup.RightFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "&P"
        zoneImpressionPagesetup.Orientation = xlLandscape
        zoneImpressionPagesetup.Zoom = 80
    End If
    Call SetTypePageSetup(wsheet)

End Sub


Public Sub cmdSendMail_ICC_MVT(blnDetail As Boolean)
Dim wSendMail As typeSendMail
Dim xDétail As String, xHeader As String, mbgColor As String
Dim K As Long, htmlFontColor_K As String
Dim xAlerte As String, xSQL As String

On Error Resume Next

'____________________________________________________________________________________________

wSendMail.FromDisplayName = "@ICC_MVT"
wSendMail.RecipientDisplayName = "COMPTA"

If blnDetail Then
    wSendMail.Subject = "Etat détaillé des créances rattachées du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax)
Else
    wSendMail.Subject = "Etat récapitulatif des créances rattachées  du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax)
End If

wSendMail.Attachment = "" 'JPL 2014-10-13 prtIMP_PDF_FileName
wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                    & paramEditionNoPaper_Auto_Lnk _
                    & "<span style='font-size:12.0pt;font-family:Arial Unicode MS'>" & "<Font color = #000080>" _
                    & "Bonjour," _
                    & "<BR> Veuillez trouver ci-joint l'état de contrôle des créances rattachées." _
                    & "<BR><BR><Font color = #000080> Bonne réception." _

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub



Public Sub Form_Init()
Dim V, xSQL As String, X As String
Dim K As Long

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True

blnControl = False

cmdReset

libICCCPTGRP.ForeColor = vbMagenta
lblICCMVTCOM.ForeColor = vbMagenta

xFF = Chr$(159) & Chr$(159)
fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False


fraSelect_Options_1.BorderStyle = 0


lstW.Visible = False

fgDetail.Visible = False: fraDetail.Visible = False
fgDetail_FormatString = fgDetail.FormatString
V = rsYBIATAB0_Read("DATE", "CAL", "M", WAMJMax)
If YBIATAB0_DATE_CPT_J < WAMJMax Then V = rsYBIATAB0_Read("DATE", "CAL", "MP1", WAMJMax)
wAMJMin = Mid$(WAMJMax, 1, 6) & "01"
Call DTPicker_Set(txtSelect_ICCMVTAMJ_Max, WAMJMax) '
Call DTPicker_Set(txtSelect_ICCMVTAMJ_Min, wAMJMin) '

cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1  - sélection (filtre)"
cboSelect_SQL.ListIndex = 0

fgParam_GRP_FormatString = fgParam_GRP.FormatString
fgParam_GRP.Enabled = YICCCPT0_Aut.Valider
fgParam_GRP.Visible = True

param_GRP

fgParam_CPT_FormatString = fgParam_CPT.FormatString
fgParam_CPT.Enabled = YICCCPT0_Aut.Valider
fgParam_CPT.Visible = True

fgParam_CPT.Rows = 1   '''''Param_CPT

fgParam_PCI_FormatString = fgParam_PCI.FormatString
fgParam_PCI.Enabled = YICCCPT0_Aut.Valider
fgParam_PCI.Visible = True

fgParam_PCI.Rows = 1
param_PCI

lstW.Clear

txtSelect_ICCMVTOPE.Clear
txtSelect_ICCMVTOPE.AddItem ""
xSQL = "select distinct ICCMVTOPE from " & paramIBM_Library_SABSPE & ".YICCMVT0 order by ICCMVTOPE"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    txtSelect_ICCMVTOPE.AddItem Trim(rsSab("ICCMVTOPE"))
    rsSab.MoveNext
Loop


Me.Enabled = True
End Sub

Private Sub fgParam_PCI_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String, xSQL As String
Dim X1 As String
On Error Resume Next

If fgParam_PCI.Rows = 1 Then param_PCI: Exit Sub


If y <= fgParam_PCI.RowHeightMin Then
    Select Case fgParam_PCI.Col
        Case 0: fgParam_PCI_Sort1 = 0: fgParam_PCI_Sort2 = 2: fgParam_PCI_Sort
        Case 1:  fgParam_PCI_Sort1 = 1: fgParam_PCI_Sort2 = 2: fgParam_PCI_Sort
        Case 2:  fgParam_PCI_Sort1 = 2: fgParam_PCI_Sort2 = 3: fgParam_PCI_Sort
        Case 3:  fgParam_PCI_Sort1 = 3: fgParam_PCI_Sort2 = 3: fgParam_PCI_Sort
    End Select
Else
    If fgParam_PCI.Rows > 1 Then
        Call fgParam_PCI_Color(fgParam_PCI_RowClick, MouseMoveUsr.BackColor, fgParam_PCI_ColorClick)
        
        Old_YBIATAB0.BIATABID = "YICCCPT0_PCI"
        Old_YBIATAB0.BIATABK1 = "PCI"
        fgParam_PCI.Col = 0:  X1 = Trim(fgParam_PCI.Text)
        fgParam_PCI.Col = 1:  Old_YBIATAB0.BIATABK1 = X1 & Trim(fgParam_PCI.Text)
        fgParam_PCI.Col = 2:  Old_YBIATAB0.BIATABK2 = fgParam_PCI.Text

        If YICCCPT0_Aut.Valider Then
            Me.PopupMenu mnuParam_PCI, vbPopupMenuLeftButton
        End If
        
   End If
End If

End Sub





Private Sub mnuExportation_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Exportation en cours ......"): DoEvents

cmdSelect_SQL_1

YICCMVT0_Export


Me.Enabled = True: Me.MousePointer = 0
End Sub

Public Sub YICCMVT0_Export()
On Error GoTo Error_Handler
Dim Nb As Long, wId As Long
Dim wFile As String, wFilex As String, wFile2 As String, xSQL As String
Dim wAMJMin As String, WAMJMax As String
Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim X As String, K As Long, kMax As Long, K2 As Long, K3 As Long
Dim xWhere As String, X2 As String

Dim s_Solde As Currency, s_Prov As Currency
Dim t_Solde As Currency, t_Prov As Currency, x_Prov As Currency
Dim mICCMVTDOS As Long, mICCMVTAMJ As String
'______________________________________________
Call DTPicker_Control(txtSelect_ICCMVTAMJ_Min, wAMJMin)
Call DTPicker_Control(txtSelect_ICCMVTAMJ_Max, WAMJMax)

wFile = "C:\Temp\YICCMVT0 " & DSys & " " & time_Hms & ".xlsx"
'______________________________________________

X = InputBox("par défaut : " & wFile _
    & vbCrLf & vbCrLf & "     =========================" _
    & vbCrLf & "     =========================", "Créances rattachées : nom du fichier d'exportation", wFile)
If Trim(X) = "" Then Exit Sub

wFilex = Trim(X)
'______________________________________________


If Dir(wFile) <> "" Then Kill wFile

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "Créances rattachées"
    .Subject = "Créances rattachées"
End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "Créances rattachées"
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

wsExcel.Cells(Nb, 1) = "Compte": wsExcel.Columns(1).ColumnWidth = 20
wsExcel.Cells(Nb, 2) = "C.Opé": wsExcel.Columns(2).ColumnWidth = 6
wsExcel.Cells(Nb, 3) = "Dossier": wsExcel.Columns(3).ColumnWidth = 9: wsExcel.Columns(3).NumberFormat = "#######"
wsExcel.Cells(Nb, 4) = "Date situ": wsExcel.Columns(4).ColumnWidth = 12: wsExcel.Columns(4).NumberFormat = "mm/dd/yyyy"
wsExcel.Cells(Nb, 5) = "Ser": wsExcel.Columns(5).ColumnWidth = 6: wsExcel.Columns(5).NumberFormat = "00"
wsExcel.Columns(5).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
wsExcel.Cells(Nb, 6) = "Sse": wsExcel.Columns(6).ColumnWidth = 6: wsExcel.Columns(6).NumberFormat = "00"
wsExcel.Columns(6).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
wsExcel.Cells(Nb, 7) = "C.éve": wsExcel.Columns(7).ColumnWidth = 6
wsExcel.Cells(Nb, 8) = "Nature": wsExcel.Columns(8).ColumnWidth = 6
wsExcel.Cells(Nb, 9) = "Prov M-1"
wsExcel.Columns(9).ColumnWidth = 12: wsExcel.Columns(9).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(Nb, 10) = "Cumul DB"
wsExcel.Columns(10).ColumnWidth = 12: wsExcel.Columns(10).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(Nb, 11) = "Cumul CR"
wsExcel.Columns(11).ColumnWidth = 12: wsExcel.Columns(11).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(Nb, 12) = "Prov M"
wsExcel.Columns(12).ColumnWidth = 12: wsExcel.Columns(12).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(Nb, 13) = "Dev": wsExcel.Columns(13).ColumnWidth = 6
wsExcel.Cells(Nb, 14) = "Solde Dossier"
wsExcel.Columns(14).ColumnWidth = 12: wsExcel.Columns(14).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(Nb, 15) = "Prov Dossier"
wsExcel.Columns(15).ColumnWidth = 12: wsExcel.Columns(15).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(Nb, 16) = "Solde Compte"
wsExcel.Columns(16).ColumnWidth = 12: wsExcel.Columns(16).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Cells(Nb, 17) = "Prov Compte"
wsExcel.Columns(17).ColumnWidth = 12: wsExcel.Columns(17).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"


For K = 1 To 17
    wsExcel.Columns(K).VerticalAlignment = Excel.XlVAlign.xlVAlignTop
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 255, 153)
Next K


For K = 1 To arrYICCCPT0_Nb

    s_Solde = 0: s_Prov = 0
    t_Solde = 0: t_Prov = 0: x_Prov = 0
    mICCMVTDOS = -1
    xWhere = " where ICCMVTETA  = " & currentZMNURUT0.MNURUTETB _
     & " and   ICCMVTAGE = " & currentZMNUUTI0.MNUUTIAGE _
     & " and ICCMVTCOM ='" & Trim(arrYICCCPT0(K).ICCCPTCOM) & "'" _
     & " and ICCMVTAMJ >= " & wAMJMin & " and ICCMVTAMJ <= " & WAMJMax _
     & mWhere_ICCMVTDOS _
     & " order by ICCMVTOPE , ICCMVTDOS , ICCMVTAMJ , ICCMVTSER , ICCMVTSSE , ICCMVTEVE "
     
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YICCMVT0 " & xWhere
    Set rsSab = cnsab.Execute(xSQL)
    Do While Not rsSab.EOF
        V = rsYICCMVT0_GetBuffer(rsSab, xYICCMVT0)
        If xYICCMVT0.ICCMVTSER <> xFF Then
            If mICCMVTDOS <> xYICCMVT0.ICCMVTDOS Then
                If s_Solde <> 0 Or s_Prov <> 0 Then
                    If s_Solde <> s_Prov Then
                        wsExcel.Cells(Nb, 14).Interior.Color = RGB(255, 168, 125)
                        wsExcel.Cells(Nb, 15).Interior.Color = RGB(255, 168, 125)
                    End If
                    wsExcel.Cells(Nb, 14) = s_Solde
                    wsExcel.Cells(Nb, 15) = s_Prov
                End If
                If s_Solde <> 0 Then
                    t_Solde = t_Solde + s_Solde
                    t_Prov = t_Prov + s_Prov
                    If mICCMVTAMJ < WAMJMax Then wsExcel.Cells(Nb, 14).Interior.Color = RGB(255, 128, 255)
                End If
                x_Prov = x_Prov + s_Prov
                s_Solde = 0: s_Prov = 0

           End If
            Nb = Nb + 1
            wsExcel.Cells(Nb, 1) = xYICCMVT0.ICCMVTCOM
            wsExcel.Cells(Nb, 2) = xYICCMVT0.ICCMVTOPE
            wsExcel.Cells(Nb, 3) = xYICCMVT0.ICCMVTDOS
            wsExcel.Cells(Nb, 4) = dateImp10_S(xYICCMVT0.ICCMVTAMJ)
            wsExcel.Cells(Nb, 5) = xYICCMVT0.ICCMVTSER
            wsExcel.Cells(Nb, 6) = xYICCMVT0.ICCMVTSSE
            wsExcel.Cells(Nb, 7) = xYICCMVT0.ICCMVTEVE
            wsExcel.Cells(Nb, 8) = xYICCMVT0.ICCMVTNAT
            If xYICCMVT0.ICCMVTRBT <> 0 Then wsExcel.Cells(Nb, 9) = xYICCMVT0.ICCMVTRBT
            If xYICCMVT0.ICCMVTPRO <> 0 Then wsExcel.Cells(Nb, 12) = xYICCMVT0.ICCMVTPRO
            If xYICCMVT0.ICCMVTTDB <> 0 Then wsExcel.Cells(Nb, 10) = xYICCMVT0.ICCMVTTDB
            If xYICCMVT0.ICCMVTTCR <> 0 Then wsExcel.Cells(Nb, 11) = xYICCMVT0.ICCMVTTCR
            wsExcel.Cells(Nb, 13) = arrYICCCPT0(K).ICCCPTDEV
            s_Solde = s_Solde + xYICCMVT0.ICCMVTTDB + xYICCMVT0.ICCMVTTCR
            s_Prov = s_Prov + xYICCMVT0.ICCMVTRBT + xYICCMVT0.ICCMVTPRO
            mICCMVTDOS = xYICCMVT0.ICCMVTDOS
            mICCMVTAMJ = xYICCMVT0.ICCMVTAMJ
        End If
        rsSab.MoveNext
    Loop
    
    If s_Solde <> 0 Or s_Prov <> 0 Then
        If s_Solde <> s_Prov Then
            wsExcel.Cells(Nb, 14).Interior.Color = RGB(255, 168, 125)
            wsExcel.Cells(Nb, 15).Interior.Color = RGB(255, 168, 125)
        End If
        wsExcel.Cells(Nb, 14) = s_Solde
        wsExcel.Cells(Nb, 15) = s_Prov
    End If
    If s_Solde <> 0 Then
        t_Solde = t_Solde + s_Solde
        t_Prov = t_Prov + s_Prov
        If mICCMVTAMJ < WAMJMax Then wsExcel.Cells(Nb, 14).Interior.Color = RGB(255, 128, 255)
    End If
    x_Prov = x_Prov + s_Prov

    If t_Solde <> t_Prov Then
        wsExcel.Cells(Nb, 16).Interior.Color = RGB(255, 153, 255)
        wsExcel.Cells(Nb, 17).Interior.Color = RGB(255, 153, 255)
    Else
        wsExcel.Cells(Nb, 16).Interior.Color = RGB(153, 255, 153)
        wsExcel.Cells(Nb, 17).Interior.Color = RGB(153, 255, 153)
    
    End If
    wsExcel.Cells(Nb, 16) = t_Solde
    wsExcel.Cells(Nb, 17) = t_Prov
    wsExcel.Cells(Nb, 18) = x_Prov
Next K

Call lstErr_ChangeLastItem(lstErr, cmdContext, "Exportation en cours : " & Nb & " enregistrements"): DoEvents
Set rsSab = Nothing


wbExcel.SaveAs wFile

wbExcel.Close

'____________________________________________________________________________________
appExcel.Quit

Set rsSab = Nothing
Set rsSabX = Nothing

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing

Call lstErr_AddItem(lstErr, cmdContext, "Exportation terminée"): DoEvents

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub

Private Sub mnuParam_CPT_Ignore_Click()
newYICCCPT0 = oldYICCCPT0
If oldYICCCPT0.ICCCPTSTA = "I" Then
    newYICCCPT0.ICCCPTSTA = ""
Else
    newYICCCPT0.ICCCPTSTA = "I"
End If

Param_CPT_Update
Param_CPT

End Sub

Private Sub mnuParam_CPT_New_Click()
Dim X As String, X1 As String, X2 As String
Dim wAmj As Long, curX As Currency, K As Integer
New_YBIATAB0 = Old_YBIATAB0

X1 = InputBox("Préciser le Compte : ", "Créances rattachées : AJOUT d'un compte")
If Trim(X1) = "" Then Exit Sub
X1 = UCase$(Trim(X1))
Call sqlYBIATAB0_Read("DATE", "CAL", "MP1", X)
wAmj = Val(X)
X = "select * from " & paramIBM_Library_SAB & ".ZCOMPTE0 " _
  & " where COMPTECOM = '" & X1 & "'"
  
Set rsSab = cnsab.Execute(X)

If rsSab.EOF Then
    Call MsgBox("Compte inconnu (ZCOMPTE0) : " & X1, vbCritical, "Créances rattachées : AJOUT d'un compte")
    Exit Sub
End If

X = InputBox("Préciser le  groupe du compte : ", "Créances rattachées : AJOUT d'un compte", 91)
If Trim(X) = "" Then Exit Sub
K = Val(X)
If K < 1 Or K > 99 Then
    Call MsgBox("valeur hors-limite (1-99) : " & K, vbCritical, "Créances rattachées : AJOUT d'un compte")
    Exit Sub
End If
If Trim(arrICCCPTGRP_Lib(K)) = "" Then
    Call MsgBox("groupe inconnu : " & K, vbCritical, "Créances rattachées : AJOUT d'un compte")
    Exit Sub
End If


newYICCCPT0.ICCCPTETA = 1
newYICCCPT0.ICCCPTAGE = 1
newYICCCPT0.ICCCPTCOM = Trim(X1)
newYICCCPT0.ICCCPTDEV = rsSab("COMPTEDEV")
newYICCCPT0.ICCCPTGRP = K
newYICCCPT0.ICCCPTSTA = ""
newYICCCPT0.ICCCPTUUSR = usrName
newYICCCPT0.ICCCPTUAMJ = DSys
newYICCCPT0.ICCCPTUHMS = time_Hms
newYICCCPT0.ICCCPTUSEQ = 0

X = "select * from " & paramIBM_Library_SAB & ".ZSOLDE0 " _
  & " where SOLDECOM = '" & X1 & "'"
  
Set rsSab = cnsab.Execute(X)

If rsSab.EOF Then
    Call MsgBox("Compte inconnu (ZSOLDE0) : " & X1, vbCritical, "Créances rattachées : AJOUT d'un compte")
    Exit Sub
End If

curX = -rsSab("SOLDEC12")
newYICCMVT0.ICCMVTETA = newYICCCPT0.ICCCPTETA
newYICCMVT0.ICCMVTAGE = newYICCCPT0.ICCCPTAGE
newYICCMVT0.ICCMVTCOM = newYICCCPT0.ICCCPTCOM
newYICCMVT0.ICCMVTSER = ""
newYICCMVT0.ICCMVTSSE = ""
newYICCMVT0.ICCMVTOPE = ""
newYICCMVT0.ICCMVTDOS = 0
newYICCMVT0.ICCMVTEVE = ""
newYICCMVT0.ICCMVTAMJ = wAmj
newYICCMVT0.ICCMVTNAT = ""
newYICCMVT0.ICCMVTEVEG = ""
newYICCMVT0.ICCMVTRBT = 0
newYICCMVT0.ICCMVTPRO = 0
newYICCMVT0.ICCMVTTDB = 0
newYICCMVT0.ICCMVTTCR = 0
If curX < 0 Then
    newYICCMVT0.ICCMVTTDB = curX
Else
    newYICCMVT0.ICCMVTTCR = curX
End If

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
V = sqlYICCCPT0_Insert(newYICCCPT0)
'________________________________________________________________________________
If Not IsNull(V) Then GoTo Error_MsgBox

V = sqlYICCMVT0_Insert(newYICCMVT0)
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
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Sub

Private Sub mnuParam_PCI_Delete_Click()
Dim X As String


X = MsgBox("confirmer la SUPPRESSION du PCI : " & Old_YBIATAB0.BIATABK1 & vbCrLf & Old_YBIATAB0.BIATABK2, vbYesNo + vbQuestion, "Créances rattachées : SUPPRESSION d'un PCI")
If X <> vbYes Then Exit Sub

Parametrage_Delete
param_PCI

End Sub

Private Sub mnuParam_PCI_New_Click()
Dim X As String, X1 As String, X2 As String
New_YBIATAB0 = Old_YBIATAB0

X1 = InputBox("Préciser le PCI minimum (6 chiffres) : ", "Créances rattachées : AJOUT d'un PCI")
If Trim(X1) = "" Then Exit Sub
X2 = InputBox("Préciser le PCI maximum (6 chiffres) : ", "Créances rattachées : AJOUT d'un PCI")
If Trim(X2) = "" Then Exit Sub

New_YBIATAB0.BIATABK1 = X1 & X2

X = InputBox("Préciser la nature du compte (3 lettres)ou ZZZ pour tous : ", "Créances rattachées : AJOUT d'un PCI")
If Trim(X) = "" Then Exit Sub
New_YBIATAB0.BIATABK2 = UCase$(Trim(X))

X = InputBox("Préciser le groupe (2 chiffres) : ", "Créances rattachées : AJOUT d'un groupe")
If Trim(X) = "" Then Exit Sub
New_YBIATAB0.BIATABTXT = Format$(Val(X), "00")

Parametrage_New
param_PCI

End Sub


Private Sub mnuParam_PCI_Update_Click()
Dim X As String, K As Integer
X = InputBox("Préciser le nouveau groupe du PCI : " & Old_YBIATAB0.BIATABK1, "Créances rattachées : MODIFICATION d'un PCI", Trim(Old_YBIATAB0.BIATABTXT))
If Trim(X) = "" Then Exit Sub

K = Val(Trim(X))
If K < 1 Or K > 99 Then
    Call MsgBox("valeur hors-limite (1-99) : " & K, vbCritical, "Créances rattachées : MODIFICATION d'un PCI")
    Exit Sub
End If
If Trim(arrICCCPTGRP_Lib(K)) = "" Then
    Call MsgBox("groupe inconnu : " & K, vbCritical, "Créances rattachées : MODIFICATION d'un PCI")
    Exit Sub
End If

Call MsgBox("ATTENTION : les comptes déjà créés ne sont réaffectés automatiquement.", vbInformation, "Créances rattachées : MODIFICATION d'un PCI")
New_YBIATAB0 = Old_YBIATAB0
New_YBIATAB0.BIATABTXT = Format$(K, "00")

Parametrage_Update
param_PCI
End Sub


Public Sub fgParam_PCI_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgParam_PCI.Visible = False: fraDetail.Visible = False
mRow = fgParam_PCI.Row

If lRow > 0 And lRow < fgParam_PCI.Rows Then
    fgParam_PCI.Row = lRow
    For I = fgParam_PCI_arrIndex To fgParam_PCI.FixedCols Step -1
        fgParam_PCI.Col = I: fgParam_PCI.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgParam_PCI.Row = mRow
    If fgParam_PCI.Row > 0 Then
        lRow = fgParam_PCI.Row
        lColor_Old = fgParam_PCI.CellBackColor
        For I = fgParam_PCI_arrIndex To fgParam_PCI.FixedCols Step -1
          fgParam_PCI.Col = I: fgParam_PCI.CellBackColor = lColor
        Next I
    End If
End If
fgParam_PCI.LeftCol = fgParam_PCI.FixedCols
fgParam_PCI.Visible = True
End Sub


Public Sub fgParam_PCI_Reset()
fgParam_PCI.Clear
fgParam_PCI_Sort1 = 0: fgParam_PCI_Sort2 = 0
fgParam_PCI_Sort1_Old = -1
fgParam_PCI_RowDisplay = 0: fgParam_PCI_RowClick = 0
fgParam_PCI_arrIndex = fgParam_PCI.Cols - 1
blnfgParam_PCI_DisplayLine = False
fgParam_PCI_SortAD = 6
fgParam_PCI.LeftCol = fgParam_PCI.FixedCols

End Sub

Public Sub fgParam_PCI_Sort()
If fgParam_PCI.Rows > 1 Then
    fgParam_PCI.Row = 1
    fgParam_PCI.RowSel = fgParam_PCI.Rows - 1
    
    If fgParam_PCI_Sort1_Old = fgParam_PCI_Sort1 Then
        If fgParam_PCI_SortAD = 5 Then
            fgParam_PCI_SortAD = 6
        Else
            fgParam_PCI_SortAD = 5
        End If
    Else
        fgParam_PCI_SortAD = 5
    End If
    fgParam_PCI_Sort1_Old = fgParam_PCI_Sort1
    
    fgParam_PCI.Col = fgParam_PCI_Sort1
    fgParam_PCI.ColSel = fgParam_PCI_Sort2
    fgParam_PCI.Sort = fgParam_PCI_SortAD
End If

End Sub


Private Sub fgParam_CPT_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String, xSQL As String
On Error Resume Next

If fgParam_CPT.Rows = 1 Then Param_CPT: Exit Sub


If y <= fgParam_CPT.RowHeightMin Then
    Select Case fgParam_CPT.Col
        Case 0: fgParam_CPT_Sort1 = 0: fgParam_CPT_Sort2 = 2: fgParam_CPT_Sort
        Case 1:  fgParam_CPT_Sort1 = 1: fgParam_CPT_Sort2 = 2: fgParam_CPT_Sort
        Case 2:  fgParam_CPT_Sort1 = 2: fgParam_CPT_Sort2 = 2: fgParam_CPT_Sort
        Case 3:  fgParam_CPT_Sort1 = 3: fgParam_CPT_Sort2 = 3: fgParam_CPT_Sort
    End Select
Else
    If fgParam_CPT.Rows > 1 Then
        Call fgParam_CPT_Color(fgParam_CPT_RowClick, MouseMoveUsr.BackColor, fgParam_CPT_ColorClick)
        
        fgParam_CPT.Col = 1
        If Trim(fgParam_CPT.Text) = "" Then
            Call MsgBox("Enregistrement technique non modifiable", vbExclamation, "Créances rattachées")
            Exit Sub
        End If
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YICCCPT0 " _
            & " where ICCCPTETA  = " & currentZMNURUT0.MNURUTETB _
            & " and   ICCCPTAGE = " & currentZMNUUTI0.MNUUTIAGE _
            & " and   ICCCPTCOM = '" & Trim(fgParam_CPT.Text) & "'"
            
        Set rsSab = cnsab.Execute(xSQL)
        
        If Not rsSab.EOF Then
            V = rsYICCCPT0_GetBuffer(rsSab, oldYICCCPT0)

            If YICCCPT0_Aut.Valider Then
                If oldYICCCPT0.ICCCPTSTA = "I" Then
                    mnuParam_CPT_Ignore.Caption = "CPT : ACTIVER ce compte"
                Else
                    mnuParam_CPT_Ignore.Caption = "CPT : IGNORER ce compte"
                End If
                
                Me.PopupMenu mnuParam_CPT, vbPopupMenuLeftButton
            End If
        End If
        
   End If
End If

End Sub










Private Sub mnuParam_CPT_Update_Click()
Dim X As String, K As Long

X = InputBox("Préciser le nouveau groupe du compte : " & vbCrLf & oldYICCCPT0.ICCCPTCOM, "Créances rattachées : MODIFICATION d'un compte", oldYICCCPT0.ICCCPTGRP)
If Trim(X) = "" Then Exit Sub
K = Val(X)
If K < 1 Or K > 99 Then
    Call MsgBox("valeur hors-limite (1-99) : " & K, vbCritical, "Créances rattachées : MODIFICATION d'un compte")
    Exit Sub
End If
If Trim(arrICCCPTGRP_Lib(K)) = "" Then
    Call MsgBox("groupe inconnu : " & K, vbCritical, "Créances rattachées : MODIFICATION d'un compte")
    Exit Sub
End If
newYICCCPT0 = oldYICCCPT0
newYICCCPT0.ICCCPTGRP = K

Param_CPT_Update
Param_CPT
End Sub


Public Sub fgParam_CPT_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgParam_CPT.Visible = False: fraDetail.Visible = False
mRow = fgParam_CPT.Row

If lRow > 0 And lRow < fgParam_CPT.Rows Then
    fgParam_CPT.Row = lRow
    For I = fgParam_CPT_arrIndex To fgParam_CPT.FixedCols Step -1
        fgParam_CPT.Col = I: fgParam_CPT.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgParam_CPT.Row = mRow
    If fgParam_CPT.Row > 0 Then
        lRow = fgParam_CPT.Row
        lColor_Old = fgParam_CPT.CellBackColor
        For I = fgParam_CPT_arrIndex To fgParam_CPT.FixedCols Step -1
          fgParam_CPT.Col = I: fgParam_CPT.CellBackColor = lColor
        Next I
    End If
End If
fgParam_CPT.LeftCol = fgParam_CPT.FixedCols
fgParam_CPT.Visible = True
End Sub


Public Sub fgParam_CPT_Reset()
fgParam_CPT.Clear
fgParam_CPT_Sort1 = 0: fgParam_CPT_Sort2 = 0
fgParam_CPT_Sort1_Old = -1
fgParam_CPT_RowDisplay = 0: fgParam_CPT_RowClick = 0
fgParam_CPT_arrIndex = fgParam_CPT.Cols - 1
blnfgParam_CPT_DisplayLine = False
fgParam_CPT_SortAD = 6
fgParam_CPT.LeftCol = fgParam_CPT.FixedCols

End Sub

Public Sub fgParam_CPT_Sort()
If fgParam_CPT.Rows > 1 Then
    fgParam_CPT.Row = 1
    fgParam_CPT.RowSel = fgParam_CPT.Rows - 1
    
    If fgParam_CPT_Sort1_Old = fgParam_CPT_Sort1 Then
        If fgParam_CPT_SortAD = 5 Then
            fgParam_CPT_SortAD = 6
        Else
            fgParam_CPT_SortAD = 5
        End If
    Else
        fgParam_CPT_SortAD = 5
    End If
    fgParam_CPT_Sort1_Old = fgParam_CPT_Sort1
    
    fgParam_CPT.Col = fgParam_CPT_Sort1
    fgParam_CPT.ColSel = fgParam_CPT_Sort2
    fgParam_CPT.Sort = fgParam_CPT_SortAD
End If

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
fgDetail.Visible = False: fraDetail.Visible = False
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
fgDetail.Visible = True: fraDetail.Visible = True
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



Private Sub cboSelect_ICCCPTOPE_LostFocus()
'txt_losttFocus cboSelect_ICCCPTOPE

End Sub


Private Sub cboSelect_ICCCPTGRP_Click()
cmdSelect_Reset
End Sub

Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

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

Select Case SSTab1.Tab
    Case 0:
        Me.PopupMenu mnuPrint, vbPopupMenuLeftButton
    End Select

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdPrint_YICCCPT0(blnDetail As Boolean)
Dim X As String, xSQL As String, I As Integer, K As Integer
Dim wAmj As String, xWhere As String
Dim soldeD As typeYICCMVT0, soldeF As typeYICCMVT0, Total As typeYICCMVT0
Dim blnXprt_Line As Boolean
Dim Nb_Detail As Long

If arrYICCCPT0_Nb = 0 Then Exit Sub

Me.Enabled = False: Me.MousePointer = vbHourglass
Call DTPicker_Control(txtSelect_ICCMVTAMJ_Min, wAMJMin)
Call DTPicker_Control(txtSelect_ICCMVTAMJ_Max, WAMJMax)
wAmj = dateElp("Jour", -1, wAMJMin)

fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort

If blnDetail Then
    prtTitleText = "Etat détaillé des créances rattachées du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax) & " - " & arrICCCPTGRP_Lib(arrYICCCPT0(1).ICCCPTGRP)
Else
    prtTitleText = "Etat récapitulatif des créances rattachées  du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax) & " - " & arrICCCPTGRP_Lib(arrYICCCPT0(1).ICCCPTGRP)
End If

arrYICCCPT0(0) = arrYICCCPT0(1)
prtYICCCPT0_Open
For K = 1 To arrYICCCPT0_Nb

    xYICCCPT0 = arrYICCCPT0(K)
    If xYICCCPT0.ICCCPTGRP <> arrYICCCPT0(K - 1).ICCCPTGRP Then
        If blnDetail Then
            prtTitleText = "Etat détaillé des créances rattachées du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax) & " - " & arrICCCPTGRP_Lib(xYICCCPT0.ICCCPTGRP)
        Else
            prtTitleText = "Etat récapitulatif des créances rattachées  du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax) & " - " & arrICCCPTGRP_Lib(xYICCCPT0.ICCCPTGRP)
        End If
        prtYICCCPT0_Close False
    End If
'________________________________________________________________
If InStr(xYICCCPT0.ICCCPTCOM, "38820") Then
    blnAvance = True
Else
    blnAvance = False
End If

    
    xWhere = " where ICCMVTETA  = " & currentZMNURUT0.MNURUTETB _
         & " and   ICCMVTAGE = " & currentZMNUUTI0.MNUUTIAGE _
         & " and ICCMVTCOM ='" & Trim(xYICCCPT0.ICCCPTCOM) & "'" _
         & " and ICCMVTAMJ >= " & wAmj & " and ICCMVTAMJ <= " & WAMJMax _
         & " order by ICCMVTAMJ , ICCMVTSER , ICCMVTSSE , ICCMVTOPE , ICCMVTEVE , ICCMVTDOS"
    
    Call arrYICCMVT0_SQL(xWhere)
    
    sSD_1 = 0: sPRO = 0
    sSD = 0: sRBT = 0
    rsYICCMVT0_Init Total
    soldeD = Total: soldeF = Total
    blnXprt_Line = True
    Nb_Detail = 0

    For I = 1 To arrYICCMVT0_Nb
             
        xYICCMVT0 = arrYICCMVT0(I)
        If xYICCMVT0.ICCMVTAMJ < wAMJMin Then
            If xYICCMVT0.ICCMVTSER = xFF Then soldeD = xYICCMVT0
        Else
            If xYICCMVT0.ICCMVTSER = xFF Then
                soldeF = xYICCMVT0
            Else
                Nb_Detail = Nb_Detail + 1
                Total.ICCMVTRBT = Total.ICCMVTRBT + xYICCMVT0.ICCMVTRBT
                Total.ICCMVTTDB = Total.ICCMVTTDB + xYICCMVT0.ICCMVTTDB
                Total.ICCMVTTCR = Total.ICCMVTTCR + xYICCMVT0.ICCMVTTCR
                Total.ICCMVTPRO = Total.ICCMVTPRO + xYICCMVT0.ICCMVTPRO
'_____________________________________________________________________
                Select Case xYICCMVT0.ICCMVTOPE
            
                    Case "EMP", "EM1":
                        If xYICCMVT0.ICCMVTRBT <> xYICCMVT0.ICCMVTTDB Then soldeD.ICCMVTRBT = 3
                        If xYICCMVT0.ICCMVTPRO <> xYICCMVT0.ICCMVTTCR Then soldeD.ICCMVTPRO = 4
                    Case Else:
                        If xYICCMVT0.ICCMVTRBT <> xYICCMVT0.ICCMVTTCR Then soldeD.ICCMVTRBT = 1
                        If xYICCMVT0.ICCMVTPRO <> xYICCMVT0.ICCMVTTDB Then soldeD.ICCMVTPRO = 2
                End Select

'____________________________________________________________________
                
                
            End If
        End If
        
    Next I

    prtYICCCPT0_Line xYICCCPT0, Total, soldeD, soldeF, blnDetail, Nb_Detail, blnAvance
    
    If blnDetail Then
        For I = 1 To arrYICCMVT0_Nb
                 
            xYICCMVT0 = arrYICCMVT0(I)
            If xYICCMVT0.ICCMVTAMJ < wAMJMin Then
            Else
                If xYICCMVT0.ICCMVTSER <> xFF Then prtYICCCPT0_Line_Detail xYICCMVT0
            End If
        Next I
    End If
Next K
prtYICCCPT0_Close True

Me.Show
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Chèques circulants_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Reset
fgSelect.Visible = False
fraSelect_Options.Visible = False

Select Case cmdSelect_SQL_K
    Case "1": fraSelect_Options.Visible = True: cmdSelect_SQL_1
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< Chèques circulants_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus

End Sub


Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String
On Error Resume Next


If y <= fgDetail.RowHeightMin Then
Else
    If fgDetail.Rows > 1 Then
       ' blnControl = False
        Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
        fgDetail.Col = fgDetail_arrIndex:  arrYICCMVT0_Index = CLng(fgDetail.Text)

   End If
End If

End Sub


Private Sub fgParam_GRP_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String, xSQL As String
On Error Resume Next


If y <= fgParam_GRP.RowHeightMin Then
    Select Case fgParam_GRP.Col
        Case 0: fgParam_GRP_Sort1 = 0: fgParam_GRP_Sort2 = 2: fgParam_GRP_Sort
        Case 1:  fgParam_GRP_Sort1 = 1: fgParam_GRP_Sort2 = 2: fgParam_GRP_Sort
    End Select
Else
    If fgParam_GRP.Rows > 1 Then
        Call fgParam_GRP_Color(fgParam_GRP_RowClick, MouseMoveUsr.BackColor, fgParam_GRP_ColorClick)
        
        Old_YBIATAB0.BIATABID = "YICCCPT0"
        Old_YBIATAB0.BIATABK1 = "GRP"
        fgParam_GRP.Col = 0:  Old_YBIATAB0.BIATABK2 = fgParam_GRP.Text
        fgParam_GRP.Col = 1:  Old_YBIATAB0.BIATABTXT = fgParam_GRP.Text

        If YICCCPT0_Aut.Valider Then
            Me.PopupMenu mnuParam_GRP, vbPopupMenuLeftButton
        End If
        
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
    Case Is = 27: cmdContext_Quit: KeyCode = 0
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select


End Sub

Public Sub cmdContext_Quit()
'blnControl = False
lstErr.Clear: lstErr.Height = 200

If SSTab1.Tab <> 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If

If fgDetail.Visible Then
    fgDetail.Visible = False: fraDetail.Visible = False
    Exit Sub
End If
If fgSelect.Visible Then
    fgSelect.Visible = False
    Exit Sub
End If

If SSTab1.Tab = 0 Then
    Unload Me
End If

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
Dim wOrigine As String, xSQL As String
On Error Resume Next


If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  arrYICCCPT0_Index = CLng(fgSelect.Text)
        
    oldYICCCPT0 = arrYICCCPT0(arrYICCCPT0_Index)
    xYICCCPT0 = oldYICCCPT0
    fgDetail_Display
        
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







Private Sub mnuParam_GRP_Delete_Click()
Dim X As String



X = "select count(*) as Tally   from " & paramIBM_Library_SABSPE & ".YICCCPT0 " _
    & " where ICCCPTGRP = " & Val(Mid$(Old_YBIATAB0.BIATABK2, 1, 2))
Set rsSab = cnsab.Execute(X)

If rsSab("Tally") <> 0 Then
    Call MsgBox("SUPPRESSION impossible, il y a " & rsSab("Tally") & " comptes dans ce groupe.", vbCritical, "Créances rattachées : SUPPRESSION d'un groupe")

    Exit Sub
End If

X = MsgBox("confirmer la SUPPRESSION du groupe : " & Old_YBIATAB0.BIATABK2 & vbCrLf & Old_YBIATAB0.BIATABTXT, vbYesNo + vbQuestion, "Créances rattachées : SUPPRESSION d'un groupe")
If X <> vbYes Then Exit Sub

Parametrage_Delete
param_GRP

End Sub

Private Sub mnuParam_GRP_New_Click()
Dim X As String
New_YBIATAB0 = Old_YBIATAB0

X = InputBox("Préciser le code(2 chiffres) : ", "Créances rattachées : AJOUT d'un groupe")
If Trim(X) = "" Then Exit Sub
New_YBIATAB0.BIATABK2 = Format$(Val(X), "00")

X = InputBox("Préciser le libellé du groupe : " & New_YBIATAB0.BIATABK2, "Créances rattachées : AJOUT d'un groupe")
If Trim(X) = "" Then Exit Sub
New_YBIATAB0.BIATABTXT = Trim(X)

Parametrage_New
param_GRP

End Sub

Private Sub mnuParam_GRP_Update_Click()
Dim X As String
X = InputBox("Préciser le nouveau libellé du groupe : " & Old_YBIATAB0.BIATABK2, "Créances rattachées : MODIFICATION d'un groupe", Trim(Old_YBIATAB0.BIATABTXT))
If Trim(X) = "" Then Exit Sub
New_YBIATAB0 = Old_YBIATAB0
New_YBIATAB0.BIATABTXT = Trim(X)

Parametrage_Update
param_GRP
End Sub

Private Sub mnuPrint_Detail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdPrint_YICCCPT0 True

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint_Recap_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdPrint_YICCCPT0 False

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub txtSelect_ICCCPTCOM_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtSelect_ICCMVTAMJ_Max_Change()
If fgSelect.Visible Then cmdSelect_Reset

End Sub

Private Sub txtSelect_ICCMVTAMJ_Max_KeyPress(KeyAscii As Integer)
cmdSelect_Reset

End Sub

Private Sub txtSelect_ICCMVTAMJ_Min_Change()
If fgSelect.Visible Then cmdSelect_Reset

End Sub


Private Sub txtSelect_ICCMVTAMJ_Min_KeyPress(KeyAscii As Integer)
cmdSelect_Reset

End Sub




Private Sub txtSelect_ICCCPTCOM_Change()
cmdSelect_Reset

End Sub

Private Sub txtSelect_ICCCPTCOM_GotFocus()
txt_GotFocus txtSelect_ICCCPTCOM
If fgSelect.Visible Then cmdSelect_Reset

End Sub


Private Sub txtSelect_ICCCPTCOM_LostFocus()
txt_LostFocus txtSelect_ICCCPTCOM

End Sub



Public Sub param_GRP()
Dim X As String, K As Integer
Dim xSQL As String

fgParam_GRP.Visible = False
For K = 0 To 99
    arrICCCPTGRP_Lib(K) = ""
Next K

X = "select count(*) as Tally   from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID = 'YICCCPT0' and BIATABK1 = 'GRP'"
Set rsSab = cnsab.Execute(X)

If rsSab("Tally") = 0 Then Param_GRP_Init

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID = 'YICCCPT0'  and BIATABK1 = 'GRP' order by BIATABK1 , BIATABK2"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    K = Val(Mid$(rsSab("BIATABK2"), 1, 2))
    arrICCCPTGRP_Lib(K) = Trim(rsSab("BIATABTXT"))
    rsSab.MoveNext
Loop

fgParam_GRP_Reset
fgParam_GRP.FormatString = fgParam_GRP_FormatString
fgParam_GRP.Rows = 1
fgParam_GRP.Row = 0

cboSelect_ICCCPTGRP.Clear
cboSelect_ICCCPTGRP.AddItem ""
For K = 0 To 99
    If arrICCCPTGRP_Lib(K) <> "" Then
        cboSelect_ICCCPTGRP.AddItem Format$(K, "00") & " - " & arrICCCPTGRP_Lib(K)
        fgParam_GRP.Rows = fgParam_GRP.Rows + 1
        fgParam_GRP.Row = fgParam_GRP.Rows - 1
        fgParam_GRP.Col = 0: fgParam_GRP.Text = Format$(K, "00")
        fgParam_GRP.Col = 1: fgParam_GRP.Text = arrICCCPTGRP_Lib(K)
    End If
    
Next K
cboSelect_ICCCPTGRP.ListIndex = 0
fgParam_GRP.Visible = True
End Sub

Public Sub param_PCI()
Dim X As String, K As Integer
Dim xSQL As String

fgParam_PCI.Visible = False
fgParam_PCI_Reset
fgParam_PCI.FormatString = fgParam_PCI_FormatString
fgParam_PCI.Rows = 1
fgParam_PCI.Row = 0

X = "select count(*) as Tally   from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID = 'YICCCPT0_PCI'"
Set rsSab = cnsab.Execute(X)

If rsSab("Tally") = 0 Then Param_PCI_Init

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID = 'YICCCPT0_PCI' order by BIATABK1 , BIATABK2"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    fgParam_PCI.Rows = fgParam_PCI.Rows + 1
    fgParam_PCI.Row = fgParam_PCI.Rows - 1
    fgParam_PCI.Col = 0: fgParam_PCI.Text = Mid$(rsSab("BIATABK1"), 1, 6)
    fgParam_PCI.Col = 1: fgParam_PCI.Text = Mid$(rsSab("BIATABK1"), 7, 6)
    fgParam_PCI.Col = 2: fgParam_PCI.Text = Mid$(rsSab("BIATABK2"), 1, 3)
    fgParam_PCI.Col = 3: fgParam_PCI.Text = Mid$(rsSab("BIATABTXT"), 1, 2)
    rsSab.MoveNext
Loop

fgParam_PCI.Visible = True

End Sub


Public Sub Param_GRP_Init()

New_YBIATAB0.BIATABID = "YICCCPT0"
New_YBIATAB0.BIATABK1 = "GRP"

New_YBIATAB0.BIATABK2 = "00"
New_YBIATAB0.BIATABTXT = "dossiers non affectés"
V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABK2 = "11"
New_YBIATAB0.BIATABTXT = "Trésorerie Prêt"
V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABK2 = "12"
New_YBIATAB0.BIATABTXT = "Trésorerie Emprunt"
V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABK2 = "21"
New_YBIATAB0.BIATABTXT = "Dépôt à terme"
V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABK2 = "31"
New_YBIATAB0.BIATABTXT = "Crédit"
V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABK2 = "41"
New_YBIATAB0.BIATABTXT = "PAR"
V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABK2 = "51"
New_YBIATAB0.BIATABTXT = "CR / INT"
V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABK2 = "61"
New_YBIATAB0.BIATABTXT = "Caution"
V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABK2 = "71"
New_YBIATAB0.BIATABTXT = "Report / Deport"
V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABK2 = "81"
New_YBIATAB0.BIATABTXT = "Portefeuille"
V = sqlYBIATAB0_Insert(New_YBIATAB0)

MsgBox "Param_GRP_Init terminé"

End Sub

Public Sub Param_PCI_Init()

New_YBIATAB0.BIATABID = "YICCCPT0_PCI"
New_YBIATAB0.BIATABK1 = "PCI"

New_YBIATAB0.BIATABTXT = "11"

New_YBIATAB0.BIATABK1 = "131701131709"
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PIS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PLT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRP": V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABK1 = "141701141709"
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PIS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PLT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRP": V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABK1 = "303701303709"
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PIS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PLT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
'_______________________________________________________________________________
New_YBIATAB0.BIATABTXT = "12"

New_YBIATAB0.BIATABK1 = "132701132709"
New_YBIATAB0.BIATABK2 = "ECD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "ECT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "EJD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "EJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "MIP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "SUB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "OAT": V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABK1 = "547000547709"
New_YBIATAB0.BIATABK2 = "ECD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "ECT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "EJD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "EJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "MIP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "SUB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "OAT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
'_______________________________________________________________________________

New_YBIATAB0.BIATABTXT = "21"

New_YBIATAB0.BIATABK1 = "132731132739"
New_YBIATAB0.BIATABK2 = "DBQ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "DAT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "GEN": V = sqlYBIATAB0_Insert(New_YBIATAB0)
'_______________________________________________________________________________

New_YBIATAB0.BIATABTXT = "41"

New_YBIATAB0.BIATABK1 = "131730131739"
New_YBIATAB0.BIATABK2 = "PAR": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK1 = "255700255729"
New_YBIATAB0.BIATABK2 = "PAR": V = sqlYBIATAB0_Insert(New_YBIATAB0)
'_______________________________________________________________________________

New_YBIATAB0.BIATABTXT = "21"

New_YBIATAB0.BIATABK1 = "131730131739"
New_YBIATAB0.BIATABK2 = "ZZZ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK1 = "255700255729"
New_YBIATAB0.BIATABK2 = "ZZZ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
'_______________________________________________________________________________

New_YBIATAB0.BIATABTXT = "31"

New_YBIATAB0.BIATABK1 = "197750197759"
New_YBIATAB0.BIATABK2 = "ZZZ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK1 = "297750297759"
New_YBIATAB0.BIATABK2 = "ZZZ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK1 = "397750397759"
New_YBIATAB0.BIATABK2 = "ZZZ": V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABK1 = "197760197769"
New_YBIATAB0.BIATABK2 = "ZZZ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK1 = "297760297769"
New_YBIATAB0.BIATABK2 = "ZZZ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK1 = "397760397769"
New_YBIATAB0.BIATABK2 = "ZZZ": V = sqlYBIATAB0_Insert(New_YBIATAB0)

New_YBIATAB0.BIATABK1 = "131701131709"
New_YBIATAB0.BIATABK2 = "PAG": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PGB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "RBC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PST": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "ICD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "IDT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCF": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCM": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHH": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PEQ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTR": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PAD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "TCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PSP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
 
New_YBIATAB0.BIATABK1 = "197701197709"
New_YBIATAB0.BIATABK2 = "PAG": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PGB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "RBC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PST": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "ICD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "IDT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCF": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCM": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHH": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PEQ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTR": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PAD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "TCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PSP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
 
New_YBIATAB0.BIATABK1 = "202701202709"
New_YBIATAB0.BIATABK2 = "PAG": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PGB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "RBC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PST": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "ICD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "IDT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCF": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCM": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHH": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PEQ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTR": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PAD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "TCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PSP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
 
New_YBIATAB0.BIATABK1 = "203701203709"
New_YBIATAB0.BIATABK2 = "PAG": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PGB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "RBC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PST": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "ICD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "IDT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCF": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCM": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHH": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PEQ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTR": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PAD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "TCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PSP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
 
New_YBIATAB0.BIATABK1 = "204701204709"
New_YBIATAB0.BIATABK2 = "PAG": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PGB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "RBC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PST": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "ICD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "IDT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCF": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCM": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHH": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PEQ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTR": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PAD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "TCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PSP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
 
New_YBIATAB0.BIATABK1 = "205701205709"
New_YBIATAB0.BIATABK2 = "PAG": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PGB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "RBC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PST": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "ICD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "IDT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCF": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCM": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHH": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PEQ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTR": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PAD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "TCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PSP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
 
New_YBIATAB0.BIATABK1 = "206701206709"
New_YBIATAB0.BIATABK2 = "PAG": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PGB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "RBC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PST": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "ICD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "IDT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCF": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCM": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHH": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PEQ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTR": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PAD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "TCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PSP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
 
New_YBIATAB0.BIATABK1 = "297701297709"
New_YBIATAB0.BIATABK2 = "PAG": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PGB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "RBC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PST": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "ICD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "IDT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCF": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCM": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHH": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PEQ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTR": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PAD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "TCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PSP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
 
New_YBIATAB0.BIATABK1 = "303701303729"
New_YBIATAB0.BIATABK2 = "PAG": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PGB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "RBC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PST": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "ICD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "IDT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCF": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCM": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHH": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PEQ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTR": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PAD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "TCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PSP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
 
New_YBIATAB0.BIATABK1 = "397701397709"
New_YBIATAB0.BIATABK2 = "PAG": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PGB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "RBC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PST": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "ICD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "IDT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCF": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCM": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHH": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PEQ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTR": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PAD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "TCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PSP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
 
New_YBIATAB0.BIATABK1 = "497701497709"
New_YBIATAB0.BIATABK2 = "PAG": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PGB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "RBC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PST": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "ICD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "IDT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCF": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCM": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHH": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PEQ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTR": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PAD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "TCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PSP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
 
New_YBIATAB0.BIATABK1 = "978000978099"
New_YBIATAB0.BIATABK2 = "PAG": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PGB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "RBC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PST": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "ICD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "IDT": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCF": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCM": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PJJ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHH": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PHS": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PDB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PEQ": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTR": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PAD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PPA": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "TCD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PRC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTI": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PTB": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PCC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PSP": V = sqlYBIATAB0_Insert(New_YBIATAB0)
 
'_______________________________________________________________________________

New_YBIATAB0.BIATABTXT = "51"

New_YBIATAB0.BIATABK1 = "388200388209"
New_YBIATAB0.BIATABK2 = "PRD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
New_YBIATAB0.BIATABK2 = "PAD": V = sqlYBIATAB0_Insert(New_YBIATAB0)
'_______________________________________________________________________________

New_YBIATAB0.BIATABTXT = "61"

New_YBIATAB0.BIATABK1 = "388200388209"
New_YBIATAB0.BIATABK2 = "CAU": V = sqlYBIATAB0_Insert(New_YBIATAB0)
'_______________________________________________________________________________

New_YBIATAB0.BIATABTXT = "81"

New_YBIATAB0.BIATABK1 = "388200388209"
New_YBIATAB0.BIATABK2 = "ESC": V = sqlYBIATAB0_Insert(New_YBIATAB0)
'_______________________________________________________________________________

New_YBIATAB0.BIATABTXT = "71"

New_YBIATAB0.BIATABK1 = "934100934209"
New_YBIATAB0.BIATABK2 = "ZZZ": V = sqlYBIATAB0_Insert(New_YBIATAB0)


MsgBox "Param_GRP_Init terminé"

End Sub


Public Sub fgParam_GRP_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgParam_GRP.Visible = False: fraDetail.Visible = False
mRow = fgParam_GRP.Row

If lRow > 0 And lRow < fgParam_GRP.Rows Then
    fgParam_GRP.Row = lRow
    For I = fgParam_GRP_arrIndex To fgParam_GRP.FixedCols Step -1
        fgParam_GRP.Col = I: fgParam_GRP.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgParam_GRP.Row = mRow
    If fgParam_GRP.Row > 0 Then
        lRow = fgParam_GRP.Row
        lColor_Old = fgParam_GRP.CellBackColor
        For I = fgParam_GRP_arrIndex To fgParam_GRP.FixedCols Step -1
          fgParam_GRP.Col = I: fgParam_GRP.CellBackColor = lColor
        Next I
    End If
End If
fgParam_GRP.LeftCol = fgParam_GRP.FixedCols
fgParam_GRP.Visible = True
End Sub


Public Sub fgParam_GRP_Reset()
fgParam_GRP.Clear
fgParam_GRP_Sort1 = 0: fgParam_GRP_Sort2 = 0
fgParam_GRP_Sort1_Old = -1
fgParam_GRP_RowDisplay = 0: fgParam_GRP_RowClick = 0
fgParam_GRP_arrIndex = fgParam_GRP.Cols - 1
blnfgParam_GRP_DisplayLine = False
fgParam_GRP_SortAD = 6
fgParam_GRP.LeftCol = fgParam_GRP.FixedCols

End Sub
Public Sub fgParam_GRP_Sort()
If fgParam_GRP.Rows > 1 Then
    fgParam_GRP.Row = 1
    fgParam_GRP.RowSel = fgParam_GRP.Rows - 1
    
    If fgParam_GRP_Sort1_Old = fgParam_GRP_Sort1 Then
        If fgParam_GRP_SortAD = 5 Then
            fgParam_GRP_SortAD = 6
        Else
            fgParam_GRP_SortAD = 5
        End If
    Else
        fgParam_GRP_SortAD = 5
    End If
    fgParam_GRP_Sort1_Old = fgParam_GRP_Sort1
    
    fgParam_GRP.Col = fgParam_GRP_Sort1
    fgParam_GRP.ColSel = fgParam_GRP_Sort2
    fgParam_GRP.Sort = fgParam_GRP_SortAD
End If

End Sub
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

Public Function Param_CPT_Update()
Dim V

cnSab_Update.Open paramODBC_DSN_SAB
V = sqlYICCCPT0_Update(newYICCCPT0, oldYICCCPT0)
cnSab_Update.Close

End Function



Public Sub Param_CPT()
X = "select * from " & paramIBM_Library_SABSPE & ".YICCCPT0 " _
    & " where ICCCPTETA  = " & currentZMNURUT0.MNURUTETB _
    & " and   ICCCPTAGE = " & currentZMNUUTI0.MNUUTIAGE _
    & " order by ICCCPTCOM , ICCCPTDEV"
    
Set rsSab = cnsab.Execute(X)

fgParam_CPT.Visible = False
fgParam_CPT_Reset
fgParam_CPT.FormatString = fgParam_CPT_FormatString
fgParam_CPT.Rows = 1
fgParam_CPT.Row = 0

Do While Not rsSab.EOF
    V = rsYICCCPT0_GetBuffer(rsSab, xYICCCPT0)
    fgParam_CPT.Rows = fgParam_CPT.Rows + 1
    fgParam_CPT.Row = fgParam_CPT.Rows - 1
    fgParam_CPT.Col = 0: fgParam_CPT.Text = xYICCCPT0.ICCCPTSTA
    fgParam_CPT.Col = 1: fgParam_CPT.Text = xYICCCPT0.ICCCPTCOM
     fgParam_CPT.Col = 2: fgParam_CPT.Text = xYICCCPT0.ICCCPTDEV
    fgParam_CPT.Col = 3: fgParam_CPT.Text = xYICCCPT0.ICCCPTGRP
   
    rsSab.MoveNext
Loop
fgParam_CPT.Visible = True

End Sub

Private Sub txtSelect_ICCMVTDOS_Change()
cmdSelect_Reset

End Sub


Private Sub txtSelect_ICCMVTDOS_KeyPress(KeyAscii As Integer)
KeyAscii = ctlNum(KeyAscii)
End Sub


Private Sub txtSelect_ICCMVTDOS_LostFocus()
txt_LostFocus txtSelect_ICCMVTDOS

End Sub

Private Sub txtSelect_ICCMVTOPE_Change()
cmdSelect_Reset

End Sub


Private Sub txtSelect_ICCMVTOPE_GotFocus()
txt_GotFocus txtSelect_ICCMVTOPE
If fgSelect.Visible Then cmdSelect_Reset

End Sub


Private Sub txtSelect_ICCMVTOPE_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtSelect_ICCMVTOPE_LostFocus()
txt_LostFocus txtSelect_ICCMVTOPE
End Sub


