VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBIA_Gafi 
   Caption         =   "Compte : surveillance"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   5
      Top             =   0
      Width           =   5985
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   105
      TabIndex        =   0
      Top             =   360
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   15901
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Choix d'un état"
      TabPicture(0)   =   "BIA_Gafi.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBalance"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Scan Adresse 'C/O BIA'"
      TabPicture(1)   =   "BIA_Gafi.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblSelect_Folder"
      Tab(1).Control(1)=   "lblSelect_Filename"
      Tab(1).Control(2)=   "fgSelect"
      Tab(1).Control(3)=   "cmdSelect_Ok"
      Tab(1).Control(4)=   "fraSelect_Options"
      Tab(1).Control(5)=   "cmdExport"
      Tab(1).Control(6)=   "txtSelect_Folder"
      Tab(1).Control(7)=   "txtSelect_FileName"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "paramétrage"
      TabPicture(2)   =   "BIA_Gafi.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraParam"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraParam 
         Height          =   8325
         Left            =   -74895
         TabIndex        =   53
         Top             =   450
         Visible         =   0   'False
         Width           =   13395
         Begin VB.CommandButton cmdParam_Print 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Imprimer le paramétrage"
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
            Left            =   855
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   6900
            Width           =   1500
         End
         Begin VB.Frame fraParam_Update 
            BackColor       =   &H00E0FFFF&
            Height          =   7965
            Left            =   3855
            TabIndex        =   55
            Top             =   270
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
               Left            =   4455
               MaxLength       =   10
               TabIndex        =   59
               Top             =   7305
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
               Left            =   2055
               Style           =   1  'Graphical
               TabIndex        =   58
               Top             =   7185
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
               Left            =   7410
               Style           =   1  'Graphical
               TabIndex        =   57
               Top             =   7155
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
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   7215
               Width           =   1500
            End
            Begin MSFlexGridLib.MSFlexGrid fgParam 
               Height          =   6240
               Left            =   165
               TabIndex        =   61
               Top             =   240
               Visible         =   0   'False
               Width           =   9045
               _ExtentX        =   15954
               _ExtentY        =   11007
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
               FormatString    =   $"BIA_Gafi.frx":0054
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
               Left            =   4590
               TabIndex        =   60
               Top             =   6930
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
            TabIndex        =   54
            Top             =   1275
            Width           =   3105
         End
      End
      Begin VB.TextBox txtSelect_FileName 
         Height          =   285
         Left            =   -66600
         TabIndex        =   38
         Top             =   2460
         Width           =   3255
      End
      Begin VB.TextBox txtSelect_Folder 
         Height          =   285
         Left            =   -72960
         TabIndex        =   35
         Top             =   2460
         Width           =   5415
      End
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Exporter"
         Height          =   600
         Left            =   -62640
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Frame fraSelect_Options 
         Height          =   1725
         Left            =   -74760
         TabIndex        =   24
         Top             =   660
         Width           =   11475
         Begin VB.CheckBox chkSelect_CLIENARES_99 
            Caption         =   "inclure les dirigeants et mandataires "
            Height          =   255
            Left            =   2400
            TabIndex        =   51
            Top             =   600
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.CheckBox chkSelect_CLIENARES 
            Caption         =   "inclure les responsables X**"
            Height          =   255
            Left            =   2400
            TabIndex        =   50
            Top             =   240
            Width           =   2895
         End
         Begin VB.CheckBox chkSelect_COMPTEFON 
            Caption         =   "Inclure les comptes annulés"
            Height          =   285
            Left            =   2400
            TabIndex        =   44
            Top             =   1320
            Width           =   2985
         End
         Begin VB.CheckBox chkSelect_COMPTEOUV 
            Caption         =   "Date création >="
            Height          =   285
            Left            =   2400
            TabIndex        =   32
            Top             =   960
            Width           =   1545
         End
         Begin VB.Frame fraSelect_Scan 
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   5640
            TabIndex        =   28
            Top             =   240
            Width           =   5655
            Begin VB.ComboBox cboSelect_ADRESSCOA 
               Height          =   315
               Left            =   1680
               Sorted          =   -1  'True
               TabIndex        =   48
               Top             =   120
               Width           =   1185
            End
            Begin VB.TextBox txtSelect_ADRESSAD1 
               Height          =   285
               Left            =   1680
               TabIndex        =   31
               Top             =   960
               Width           =   3765
            End
            Begin VB.OptionButton optSelect_Scan_Texte 
               Caption         =   "Texte =>"
               Height          =   195
               Left            =   240
               TabIndex        =   30
               Top             =   1080
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton optSelect_Scan_Auto 
               Caption         =   "Auto 'BIA - B I A-B.I.A - Roosevelt"""
               Height          =   255
               Left            =   240
               TabIndex        =   29
               Top             =   600
               Width           =   3975
            End
            Begin VB.Label lblSelect_ADRESSCOA 
               Caption         =   "Code adresse"
               Height          =   270
               Left            =   240
               TabIndex        =   49
               Top             =   120
               Width           =   1050
            End
         End
         Begin VB.Frame fraSelect_ADRESSTYP 
            BorderStyle     =   0  'None
            Height          =   1455
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Width           =   2175
            Begin VB.OptionButton optSelect_ADRESSTYP_2 
               Caption         =   "Compte"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   1080
               Width           =   1785
            End
            Begin VB.OptionButton optSelect_ADRESSTYP_1 
               Caption         =   "Client (> 10000)"
               Height          =   435
               Left            =   120
               TabIndex        =   26
               Top             =   120
               Value           =   -1  'True
               Width           =   1965
            End
         End
         Begin MSComCtl2.DTPicker txtSelect_COMPTEOUV 
            Height          =   300
            Left            =   4200
            TabIndex        =   33
            Top             =   960
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
            Format          =   100073475
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
      End
      Begin VB.CommandButton cmdSelect_Ok 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Exécuter la requête"
         Height          =   645
         Left            =   -62640
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   780
         Width           =   1095
      End
      Begin VB.Frame fraBalance 
         Height          =   8535
         Left            =   120
         TabIndex        =   1
         Top             =   660
         Width           =   13455
         Begin VB.Frame fraPériode 
            Height          =   3255
            Left            =   360
            TabIndex        =   8
            Top             =   4800
            Width           =   12375
            Begin VB.Frame fraSelect 
               Caption         =   "Sélection"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1695
               Left            =   240
               TabIndex        =   16
               Top             =   1200
               Width           =   6195
               Begin VB.TextBox txtCr 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3720
                  TabIndex        =   20
                  Top             =   1080
                  Width           =   1695
               End
               Begin VB.TextBox txtDb 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   3720
                  TabIndex        =   19
                  Top             =   360
                  Width           =   1695
               End
               Begin VB.CheckBox chkCr 
                  Caption         =   "Crédit  >  ******* EUR"
                  Height          =   495
                  Left            =   360
                  TabIndex        =   18
                  Top             =   960
                  Width           =   2280
               End
               Begin VB.CheckBox chkDb 
                  Caption         =   "Débit >  ******* EUR"
                  Height          =   465
                  Left            =   360
                  TabIndex        =   17
                  Top             =   360
                  Width           =   2580
               End
            End
            Begin VB.CommandButton cmdOk 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Ok"
               Height          =   1245
               Left            =   8760
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   1440
               Width           =   2055
            End
            Begin MSComCtl2.DTPicker txtAmjMin 
               Height          =   300
               Left            =   840
               TabIndex        =   9
               Top             =   600
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
               Format          =   100073475
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtAmjMax 
               Height          =   300
               Left            =   2880
               TabIndex        =   12
               Top             =   600
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
               Format          =   100073475
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   2
            End
            Begin VB.Label libInfo 
               BackColor       =   &H00C0C0FF&
               Caption         =   "     sont exclus : les mvts CP ,TC, EM1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   6840
               TabIndex        =   15
               Top             =   600
               Width           =   4890
            End
            Begin VB.Label lblAmjMax 
               Caption         =   "au"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2280
               TabIndex        =   11
               Top             =   600
               Width           =   315
            End
            Begin VB.Label lblAmjMin 
               Caption         =   "du"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   10
               Top             =   600
               Width           =   315
            End
         End
         Begin VB.Frame fraScript 
            Caption         =   "Script"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3975
            Left            =   360
            TabIndex        =   6
            Top             =   480
            Width           =   3135
            Begin VB.CommandButton cmkOk_Mensuel 
               BackColor       =   &H00FFC0FF&
               Caption         =   "Etats mensuels"
               Height          =   885
               Left            =   720
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   2760
               Width           =   1695
            End
            Begin VB.CommandButton cmdOk_Quotidien 
               BackColor       =   &H00FFC0FF&
               Caption         =   "Etats quotidiens"
               Height          =   885
               Left            =   720
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   1320
               Width           =   1695
            End
            Begin VB.OptionButton optEtatManuel 
               Caption         =   "Manuel"
               Height          =   255
               Left            =   960
               TabIndex        =   7
               Top             =   360
               Value           =   -1  'True
               Width           =   1095
            End
         End
         Begin VB.Frame fraEtat 
            Caption         =   "Etat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4032
            Left            =   4200
            TabIndex        =   2
            Top             =   480
            Width           =   8415
            Begin VB.OptionButton optEtat09 
               Caption         =   "9- Surveillance des comptes Contentieux"
               Height          =   255
               Left            =   800
               TabIndex        =   63
               Top             =   2730
               Width           =   4092
            End
            Begin VB.OptionButton optEtat07 
               Caption         =   "7- Etat de surveillance - OBNL"
               Height          =   255
               Left            =   800
               TabIndex        =   52
               Top             =   1950
               Width           =   3492
            End
            Begin VB.OptionButton optEtat08 
               Caption         =   "8- Etat de surveillance - PEP"
               Height          =   255
               Left            =   800
               TabIndex        =   47
               Top             =   2340
               Width           =   4092
            End
            Begin VB.ListBox lstW 
               Enabled         =   0   'False
               Height          =   2595
               Left            =   5760
               Sorted          =   -1  'True
               TabIndex        =   46
               Top             =   720
               Width           =   2295
            End
            Begin VB.OptionButton optEtat06 
               Caption         =   "6- Etat de surveillance - TEST en cours"
               Height          =   255
               Left            =   800
               TabIndex        =   45
               Top             =   1560
               Width           =   4455
            End
            Begin VB.OptionButton optEtat05 
               Caption         =   "5-Etat de surveillance des comptes PTNC"
               Height          =   255
               Left            =   800
               TabIndex        =   41
               Top             =   1170
               Width           =   4455
            End
            Begin VB.OptionButton optEtat04 
               Caption         =   "4- Etat de surveillance des comptes PARADIS FISCAUX (> 7000)"
               Height          =   612
               Left            =   800
               TabIndex        =   40
               Top             =   3360
               Width           =   3492
            End
            Begin VB.OptionButton optEtat03 
               Caption         =   "3- Etat de surveillance des comptes TRACFIN"
               Height          =   255
               Left            =   800
               TabIndex        =   39
               Top             =   780
               Width           =   4455
            End
            Begin VB.OptionButton optEtat02 
               Caption         =   "2- Etat des cumuls des mvts > 7 000 "
               Height          =   255
               Left            =   800
               TabIndex        =   14
               Top             =   3120
               Value           =   -1  'True
               Width           =   3720
            End
            Begin VB.OptionButton optEtat01 
               Caption         =   "1- Etat des mouvements > 150 000"
               Height          =   255
               Left            =   800
               TabIndex        =   3
               Top             =   360
               Width           =   3492
            End
            Begin VB.Label Label1 
               BackColor       =   &H00C0C0FF&
               Caption         =   "     par RESPONSABLE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5400
               TabIndex        =   21
               Top             =   360
               Width           =   2730
            End
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   6225
         Left            =   -74880
         TabIndex        =   22
         Top             =   2940
         Width           =   13560
         _ExtentX        =   23918
         _ExtentY        =   10980
         _Version        =   393216
         Rows            =   1
         Cols            =   7
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
         FormatString    =   $"BIA_Gafi.frx":011E
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
      Begin VB.Label lblSelect_Filename 
         Caption         =   "Fichier"
         Height          =   255
         Left            =   -67320
         TabIndex        =   37
         Top             =   2460
         Width           =   615
      End
      Begin VB.Label lblSelect_Folder 
         Caption         =   "Destination : Répertoire"
         Height          =   255
         Left            =   -74760
         TabIndex        =   36
         Top             =   2460
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmBIA_Gafi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean, blnSetfocus As Boolean, blnControl As Boolean
Dim BIA_Gafi_Aut As typeAuthorization
Dim X As String, X1 As String, I As Long
'Dim Msg As String, valX As String, V As Variant

Dim optEtat As String * 1, optSolde As String * 1, optAmj As String * 8, SrvCptP0_Amj As String * 8
Dim blnCompteMinMax As Boolean, selCompteMin As String * 11, selCompteMax As String * 11
Dim blnDevise As Boolean, selDeviseN As String * 3, blnDeviseIn As Boolean, selDeviseCV As String * 3
Dim blnService As Boolean, selService As String * 3
Dim blnBiaTyp As Boolean, selBiaTyp As String * 3
Dim optSortK As String * 1
Dim optEtatSortK As String * 2
Dim mDestinataire As String, mEnTete As String
Dim PrintRupture_Len As Integer

Dim blnExport As Boolean, X137 As String * 137
Dim X1000 As String * 1000
Dim cmdImport_Select_Nb As Long, cmdImport_Nb As Long

Dim blnService_Enabled As Boolean
Dim wL As Long, wPAys As String * 4, wX As String
Dim wAMJMin As String * 8, WAMJMax As String * 8, wAmj As String * 8
Dim xAmjMin_IBM As Long, xAmjMax_IBM As Long
Dim xAmjMin As String, xAmjMax As String, xAMJ As String
Dim vReturn As Variant
Dim mID14 As String * 14, wMt As Currency, wMtCV As Currency
Dim wDB1 As Currency, wCR1 As Currency, wDB2 As Currency, wCR2 As Currency, wVR4 As Currency
Dim sSD1 As Currency, sCR As Currency, sDB As Currency, sSD2 As Currency
Dim tSD1 As Currency, tCR As Currency, tDB As Currency, tSD2 As Currency
Dim sDev As String * 3, tCompte As String * 11, tIntitulé As String

Dim paramCompteGafi_Cpt_Export As String, paramCompteGafi_Mvt_Import As String



Dim fgParam_FormatString As String, fgParam_K As Integer
Dim fgParam_RowDisplay As Integer, fgParam_RowClick As Integer, fgParam_ColClick As Integer
Dim fgParam_ColorClick As Long, fgParam_ColorDisplay As Long
Dim fgParam_Sort1 As Integer, fgParam_Sort2 As Integer
Dim fgParam_SortAD As Integer, fgParam_Sort1_Old As Integer
Dim fgParam_arrIndex As Integer
Dim blnfgParam_DisplayLine As Boolean
Dim lstParam_K As String
Dim rsParam As New ADODB.Recordset

Dim curX As Currency, curDB As Currency, curCR As Currency
Dim blnDb As Boolean, blnCr As Boolean
Dim mAmj_COMPTEOUV As Long
Dim blnYBIAMVT0_Import As Boolean
Dim meCV1 As typeCV, meCV2 As typeCV

Dim arrCLIENARES(100) As String, arrCLIENARES_Nb As Integer


Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim selZADRESS0() As typeZADRESS0, selZADRESS0_Nb As Long, selZADRESS0_Max As Long
Dim xZADRESS0 As typeZADRESS0
Dim xZCLIENA0 As typeZCLIENA0

Dim xZCOMPTE0 As typeZCOMPTE0
Dim arrYBIACPT0()  As typeYBIACPT0, arrYBIACPT0_Nb As Integer, arrYBIACPT0_NbMax As Integer
Dim xYBIACPT0  As typeYBIACPT0
Dim arrYBIACPT0_Index As Integer
Dim mCLIENARES As String
Dim xYBIAMVT0  As typeYBIAMVT0
Dim selYBIAMVT0()  As typeYBIAMVT0, selYBIAMVT0_Nb As Integer, selYBIAMVT0_NbMax As Integer
Dim selYBIAMVT9()  As typeYBIAMVT9
Dim autoGAFI As Boolean

Dim Old_YBIATAB0 As typeYBIATAB0, New_YBIATAB0 As typeYBIATAB0


Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim mXls1_Row As Long, mXls1_File As Integer
Dim mXls2_Cols As Integer, mXls2_Row As Integer
Dim rsSabX As New ADODB.Recordset



Public Sub cmdFlux_xlsManual()
Dim K As Integer, K2 As Integer, xSQL As String, V
Dim xEtat As String
Dim kIndex As Long, selIndex As Long
Dim wMOUVEMCOM As String
Dim X As String

cmdImport_Init

Call lstErr_Clear(lstErr, cmdOK, "Flux : Début du traitement")


Me.MousePointer = vbHourglass
Me.Enabled = False
Call DTPicker_Control(txtAmjMin, wAMJMin)
xAmjMin_IBM = dateIBM(wAMJMin)
Call DTPicker_Control(txtAmjMax, WAMJMax)
xAmjMax_IBM = dateIBM(WAMJMax)

If optEtat03 Then cmdFlux_optEtat03: Exit Sub
If optEtat04 Then cmdFlux_optEtat04: Exit Sub

If optEtat05 Then cmdFlux_optEtat05: Exit Sub
If optEtat06 Then cmdFlux_optEtat06: Exit Sub
If optEtat07 Then cmdFlux_optEtat07: Exit Sub
If optEtat08 Then cmdFlux_optEtat08: Exit Sub
If optEtat09 Then
    Call cmdImport09_MvtP0_xlsManual
    Exit Sub
End If
cmdImport_MvtP0

If optEtat01 Then xEtat = "01"
If optEtat02 Then xEtat = "02"
lstW.Clear
xYBIACPT0.COMPTECOM = ""

For K = 1 To selYBIAMVT0_Nb
    If xYBIACPT0.COMPTECOM <> selYBIAMVT0(K).MOUVEMCOM Then
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
             & " where COMPTECOM = '" & selYBIAMVT0(K).MOUVEMCOM & "'"
             
        Set rsSab = cnsab.Execute(xSQL)
        
        V = rsYBIACPT0_GetBuffer(rsSab, xYBIACPT0)
        If Not IsNull(V) Then
            MsgBox Error, vbCritical, "COMPTECOM = '" & selYBIAMVT0(K).MOUVEMCOM & "'"
        Else
            lstW.AddItem xYBIACPT0.CLIENARES & " " & Trim(xYBIACPT0.COMPTECOM) & " : " & K
        End If
    End If
Next K
lstW.AddItem "XXX:0"


mCLIENARES = ""
prtBIA_Gafi.arrYBIAMVT0_Nb = 0
ReDim prtBIA_Gafi.arrYBIAMVT0(selYBIAMVT0_Nb)

For K = 0 To lstW.ListCount - 1
    lstW.ListIndex = K
    X = lstW.Text
    kIndex = InStr(X, ":")
    If kIndex > 0 Then
        If mCLIENARES <> Mid$(X, 1, 3) Then
            If prtBIA_Gafi.arrYBIAMVT0_Nb > 0 Then
                Call lstErr_AddItem(lstErr, cmdOK, "Flux : " & mCLIENARES)
                prtBIA_Gafi_Monitor xEtat, curDB, curCR, wAMJMin, WAMJMax, mCLIENARES
            End If
            prtBIA_Gafi.arrYBIAMVT0_Nb = 0
            mCLIENARES = Mid$(X, 1, 3)
        End If
        
        selIndex = Val(Mid$(X, kIndex + 1, Len(X) - kIndex))
        If selIndex > 0 Then
            wMOUVEMCOM = selYBIAMVT0(selIndex).MOUVEMCOM
            For K2 = selIndex To selYBIAMVT0_Nb
                If wMOUVEMCOM = selYBIAMVT0(K2).MOUVEMCOM Then
                    prtBIA_Gafi.arrYBIAMVT0_Nb = prtBIA_Gafi.arrYBIAMVT0_Nb + 1
                    prtBIA_Gafi.arrYBIAMVT0(prtBIA_Gafi.arrYBIAMVT0_Nb) = selYBIAMVT0(K2)
                Else
                    Exit For
                End If
                
            Next K2
        End If
    End If
Next K


Me.MousePointer = 0
Me.Enabled = True

End Sub

Public Sub cmdImport09_MvtP0_xlsManual()
Dim xSQL As String, X As String, from2 As String
Dim I As Long
Dim aClienacli As String
Dim aCpt As String
Dim blnPrintCpt As Boolean
Dim wFileName As String
Dim currentSheet As Long
Dim currentRow As Long
Dim ligneorigine As Long
Dim nbrMvts As Long

    On Error Resume Next
    Call lstErr_Clear(lstErr, cmdOK, "Chargement des mouvements ...")

    ReDim selYBIAMVT9(0)
    selYBIAMVT9(0).MOUVEMNUM = 0
    Set rsSab = Nothing
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_GAFI' and BIATABK1 = '9' order by BIATABK2"
    Set rsParam = cnsab.Execute(xSQL)
    X = ""
    Do While Not rsParam.EOF
        X = X & "'" & Trim(rsParam("BIATABK2")) & "',"
        rsParam.MoveNext
    Loop
    Set rsParam = Nothing
    Set rsSab = Nothing
    from2 = paramIBM_Library_SABSPE & ".YBIACPT0"
    If X <> "" Then
        If Right(X, 1) = "," Then
            X = Left(X, Len(X) - 1)
        End If
        xSQL = "SELECT MOUVEMCOM,MOUVEMDTR,MOUVEMDOP,MOUVEMPIE,MOUVEMECR,MOUVEMMON,MOUVEMDVA,MOUVEMSER,MOUVEMSSE,"
        xSQL = xSQL & "MOUVEMOPE,MOUVEMNUM,MOUVEMEVE,CLIENACLI,CLIENARSD,"
        xSQL = xSQL & from2 & ".COMPTEDEV," & from2 & ".COMPTEINT,LIBELLIB1,LIBELLIB2,LIBELLIB3"
        xSQL = xSQL & " FROM " & paramIBM_Library_SABSPE & ".YBIAMVTH," & from2
        xSQL = xSQL & " WHERE MOUVEMDTR = " & dateIBM(DSys_VeilleO)
        xSQL = xSQL & " AND CLIENACLI IN(" & X & ")"
        xSQL = xSQL & " AND RTRIM(COMPTECOM) = RTRIM(MOUVEMCOM)"
        xSQL = xSQL & " ORDER BY MOUVEMCOM,MOUVEMDTR,MOUVEMPIE,MOUVEMECR"
        Set rsSab = cnsab.Execute(xSQL)
        Do While Not rsSab.EOF
            selYBIAMVT9(0).MOUVEMNUM = selYBIAMVT9(0).MOUVEMNUM + 1
            ReDim Preserve selYBIAMVT9(selYBIAMVT9(0).MOUVEMNUM)
            Call rsYBIAMVT9_GetBuffer(rsSab, selYBIAMVT9(selYBIAMVT9(0).MOUVEMNUM))
        rsSab.MoveNext
        Loop
    End If
    Call lstErr_Clear(lstErr, cmdOK, "Nb mouvements sélectionnés : " & CStr(selYBIAMVT9(selYBIAMVT9(0).MOUVEMNUM).MOUVEMNUM))
    prtTitleText = "Clients suivis par le service Contentieux : Surveillance des comptes pour la date du : " & dateAAAAMMJJTOJJ_MM_AAAA(DSys_VeilleO)
    aClienacli = ""
    aCpt = ""
    blnPrintCpt = False
    wFileName = paramIMP_PDF_Path_Temp & "\suivi_contentieux.pdf"
    If selYBIAMVT9(0).MOUVEMNUM > 0 Then
        If (Dir(wFileName)) <> "" Then
            Kill wFileName
        End If
        'On recopie le classeur modèle de c:\BIASRV vers c:\temp\imp_pdf
        FileCopy paramFolder_Local & "\Modeles\modele_SuiviContentieux.xlsx", paramIMP_PDF_Path_Temp & "\modele_SuiviContentieux.xlsx"
        'on charge CE classeur dans Excel
        Call init_xlsManual
        Call appExcelPublic.Workbooks.Open(paramIMP_PDF_Path_Temp & "\modele_SuiviContentieux.xlsx")
        Set wbExcel = appExcelPublic.ActiveWorkbook
        With wbExcel
            .Title = "Contentieux"
            .Subject = "Contentieux"
        End With
        Range("A9").Select
        currentSheet = 1
        currentRow = 1
        nbrMvts = 0
        wbExcel.Sheets(currentSheet).Cells(currentRow, 6) = prtTitleText
        currentRow = 8
    End If
    
    For I = 1 To selYBIAMVT9(0).MOUVEMNUM
        blnPrintCpt = False
        ligneorigine = 6
        If I < selYBIAMVT9(0).MOUVEMNUM Then
            If Trim(selYBIAMVT9(I).MOUVEMCOM) <> Trim(selYBIAMVT9(I + 1).MOUVEMCOM) Then
                ligneorigine = 7
            End If
        Else
            ligneorigine = 7
        End If
        If Trim(selYBIAMVT9(I).MOUVEMCOM) <> Trim(aCpt) Then
            blnPrintCpt = True
        End If
        If selYBIAMVT9(I).CLIENACLI <> aClienacli And aClienacli <> "" Then
            'ligne de type 5
            ligneorigine = 5
            aClienacli = selYBIAMVT9(I).CLIENACLI
        End If
        If ligneorigine = 6 Or ligneorigine = 7 Then
            nbrMvts = nbrMvts + 1
        End If
        Call prtBIA_Gafi.prtBIA_Gafi_Line9_xlsManual(selYBIAMVT9(I), blnPrintCpt, ligneorigine, currentRow, wbExcel.Sheets(currentSheet))
        aCpt = selYBIAMVT9(I).MOUVEMCOM
    Next I
    If selYBIAMVT9(0).MOUVEMNUM > 0 Then
        currentRow = currentRow + 1
        wbExcel.Sheets(currentSheet).Select
        Range("A8:J8").Select
        Selection.Copy
        Range("A" & CStr(currentRow)).Select
        ActiveSheet.Paste
        wbExcel.Sheets(currentSheet).Cells(currentRow, 1) = CStr(nbrMvts) & " mouvement(s)"
        'on supprime les 5 lignes modèles
        Rows("4:8").Select
        Selection.Delete
        currentRow = currentRow - 5
        Call zoneImpression_xlsManual("Contentieux", currentRow, wbExcel.Sheets(currentSheet))
        Call ActiveSheet.ExportAsFixedFormat(xlTypePDF, wFileName)
        Call wbExcel.Close(True)
        Set wbExcel = Nothing
        Kill paramIMP_PDF_Path_Temp & "\modele_SuiviContentieux.xlsx"
    End If
    If selYBIAMVT9(0).MOUVEMNUM > 0 Then
        If (Dir(wFileName) <> "") Then
            'envoi du mail à SUIVI_CONTENTIEUX.ALERTE dans YSSIMEL0
            X = "<body bgcolor=" & Asc34 & "MAGENTA" & Asc34 & ">" _
                            & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
                            & htmlFontColor("BLUE") & "<BR><BR>" & "Détection de mouvements sur les comptes suivis par le Service Contentieux." _
                            & "<BR><BR>" & "Vérifier le contenu de la pièce jointe : suivi_contentieux.pdf"
        
            Call Email_Alerte("ALERTE", "SUIVI_CONTENTIEUX", "Document : " & wFileName, X, True, wFileName)
        End If
        If Not autoGAFI Then
            Call MsgBox("Fin du traitement 'Clients suivis par le service Contentieux' ...")
            GoTo sortie
        Else
            GoTo sortie
        End If
    Else
        If Not autoGAFI Then
            Call MsgBox("Aucun mouvement pour les Clients suivis par le service Contentieux ...")
            GoTo sortie
        End If
    End If
sortie:
    Unload frmBIA_Gafi

End Sub

Private Sub cmdOk_Click_xlsManual()
Call lstErr_Clear(lstErr, cmdOK, "Début du traitement")
If blnControl Then cmdControl
If lstErr.ListCount <> 0 Then Exit Sub

Call cmdFlux_xlsManual
Me.Show

End Sub

Public Function isProspect(numCli As String) As Boolean
Dim xSQL As String
Dim rs As ADODB.Recordset

    xSQL = "select clirefref from " & paramIBM_Library_SAB & ".zcliref0 where clirefref ='" & Trim(numCli) & "'"
    Set rs = cnsab.Execute(xSQL)
    If Not rs.EOF Then
        isProspect = False
    Else
        isProspect = True
    End If
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing
    
End Function

Public Sub cmdImport09_MvtP0()
Dim xSQL As String, X As String, from2 As String
Dim I As Long
Dim aClienacli As String
Dim aCpt As String
Dim blnPrintCpt As Boolean
Dim aPrinter As String
Dim wFileName As String
Dim aClef As String

    On Error Resume Next
    Call lstErr_Clear(lstErr, cmdOK, "Chargement des mouvements ...")

'    If nomDuServeur <> paramServerSplf Then
'        aClef = get_PDFCreator_AutosaveFilename
'        If aClef <> "" Then
'            Call killProcessDotNet("PDFCREATOR")
'            DoEvents
'            Call set_PDFCreator_AutosaveFilename("suivi_contentieux")
'        End If
'    End If
    ReDim selYBIAMVT9(0)
    selYBIAMVT9(0).MOUVEMNUM = 0
    Set rsSab = Nothing
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_GAFI' and BIATABK1 = '9' order by BIATABK2"
    Set rsParam = cnsab.Execute(xSQL)
    X = ""
    Do While Not rsParam.EOF
        X = X & "'" & Trim(rsParam("BIATABK2")) & "',"
        rsParam.MoveNext
    Loop
    Set rsParam = Nothing
    Set rsSab = Nothing
    from2 = paramIBM_Library_SABSPE & ".YBIACPT0"
    If X <> "" Then
        If Right(X, 1) = "," Then
            X = Left(X, Len(X) - 1)
        End If
        xSQL = "SELECT MOUVEMCOM,MOUVEMDTR,MOUVEMDOP,MOUVEMPIE,MOUVEMECR,MOUVEMMON,MOUVEMDVA,MOUVEMSER,MOUVEMSSE,"
        xSQL = xSQL & "MOUVEMOPE,MOUVEMNUM,MOUVEMEVE,CLIENACLI,CLIENARSD,"
        xSQL = xSQL & from2 & ".COMPTEDEV," & from2 & ".COMPTEINT,LIBELLIB1,LIBELLIB2,LIBELLIB3"
        xSQL = xSQL & " FROM " & paramIBM_Library_SABSPE & ".YBIAMVTH," & from2
        xSQL = xSQL & " WHERE MOUVEMDTR = " & dateIBM(DSys_VeilleO)
        xSQL = xSQL & " AND CLIENACLI IN(" & X & ")"
        xSQL = xSQL & " AND RTRIM(COMPTECOM) = RTRIM(MOUVEMCOM)"
        xSQL = xSQL & " ORDER BY MOUVEMCOM,MOUVEMDTR,MOUVEMPIE,MOUVEMECR"
        Set rsSab = cnsab.Execute(xSQL)
        Do While Not rsSab.EOF
            selYBIAMVT9(0).MOUVEMNUM = selYBIAMVT9(0).MOUVEMNUM + 1
            ReDim Preserve selYBIAMVT9(selYBIAMVT9(0).MOUVEMNUM)
            Call rsYBIAMVT9_GetBuffer(rsSab, selYBIAMVT9(selYBIAMVT9(0).MOUVEMNUM))
        rsSab.MoveNext
        Loop
    End If
    Call lstErr_Clear(lstErr, cmdOK, "Nb mouvements sélectionnés : " & CStr(selYBIAMVT9(selYBIAMVT9(0).MOUVEMNUM).MOUVEMNUM))
    prtTitleText = "Clients suivis par le service Contentieux : Surveillance des comptes pour la date du : " & dateAAAAMMJJTOJJ_MM_AAAA(DSys_VeilleO)
    aClienacli = ""
    aCpt = ""
    blnPrintCpt = False
    'wFileName = paramIMP_PDF_Path_Temp & "\suivi_contentieux.pdf"
    wFileName = paramIMP_PDF_Path_Temp & "\Releve_.pdf"
    If selYBIAMVT9(0).MOUVEMNUM > 0 Then
        If (Dir(wFileName)) Then
            Kill wFileName
        End If
        aPrinter = Printer.Devicename
        Call prtBIA_Gafi.prtBIA_Gafi_Open9
        DoEvents
    End If
    For I = 1 To selYBIAMVT9(0).MOUVEMNUM
        If Trim(selYBIAMVT9(I).MOUVEMCOM) <> Trim(aCpt) Then
            blnPrintCpt = True
        End If
        If selYBIAMVT9(I).CLIENACLI <> aClienacli And aClienacli <> "" Then
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
            aClienacli = selYBIAMVT9(I).CLIENACLI
        End If
        Call prtBIA_Gafi.prtBIA_Gafi_Line9(selYBIAMVT9(I), blnPrintCpt)
        aCpt = selYBIAMVT9(I).MOUVEMCOM
    Next I
    If selYBIAMVT9(0).MOUVEMNUM > 0 Then
        Call prtBIA_Gafi.prtBIA_Gafi_Close9(aPrinter)
        'petite tempo
        Call Sleep(5000)
        If (Dir(wFileName)) Then
            'On renomme le fichier
            Name wFileName As paramIMP_PDF_Path_Temp & "\Suivi_contentieux.pdf"
            'envoi du mail à SUIVI_CONTENTIEUX.ALERTE dans YSSIMEL0
            X = "<body bgcolor=" & Asc34 & "MAGENTA" & Asc34 & ">" _
                            & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
                            & htmlFontColor("BLUE") & "<BR><BR>" & "Détection de mouvements sur les comptes suivis par le Service Contentieux." _
                            & "<BR><BR>" & "Vérifier le contenu de la pièce jointe : Suivi_contentieux.pdf"
        
            Call Email_Alerte("ALERTE", "SUIVI_CONTENTIEUX", "Document : " & wFileName, X, True, wFileName)
        End If
        If Not autoGAFI Then
            Call MsgBox("Fin du traitement 'Clients suivis par le service Contentieux' ...")
            GoTo sortie
        Else
            GoTo sortie
        End If
    Else
        If Not autoGAFI Then
            Call MsgBox("Aucun mouvement pour les Clients suivis par le service Contentieux ...")
            GoTo sortie
        End If
    End If
sortie:
    'On part du principe que le nom de fichier de sortie par défaut est Releve_
'    If nomDuServeur <> paramServerSplf Then
'        If aClef <> "" Then
'            Call set_PDFCreator_AutosaveFilename(aClef)
'        End If
'    End If
    Unload frmBIA_Gafi

    
End Sub




Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
         If TypeOf xobj Is CommandButton Or TypeOf xobj Is ListBox Then
           xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            xobj.ForeColor = MouseMoveActiveControl.ForeColor
        End If
        Exit For
    End If
Next xobj

End Sub

Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Or TypeOf C Is ListBox Then
            
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
       Else
            MouseMoveActiveControl.ForeColor = C.ForeColor
            C.ForeColor = MouseMoveUsr.ForeColor
        End If
    End If
End If

End Sub


Private Sub selZADRESS0_SQL()
Dim V
Dim X As String, xSQL As String

On Error GoTo Error_Handler
ReDim selZADRESS0(1001)
selZADRESS0_Max = 1000: selZADRESS0_Nb = 0

Set rsSab = Nothing
If optSelect_ADRESSTYP_1 Then
    xSQL = "ADRESSTYP =  '1' and adressnum > ' 0010000'"
Else
    xSQL = "ADRESSTYP =  '2'"
End If

X = Trim(cboSelect_ADRESSCOA)

If X <> "" Then
    If X = "<  >" Then X = "  "
    xSQL = xSQL & " and ADRESSCOA = '" & X & "'"
End If

xSQL = "select * from " & paramIBM_Library_SAB & ".ZADRESS0 where " & xSQL & " order by ADRESSNUM"
Set rsSab = cnsab.Execute(xSQL)

    
Do While Not rsSab.EOF
    V = rsZADRESS0_GetBuffer(rsSab, xZADRESS0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSAB_Tc_Limites.fgdossier_Display"
        '' Exit Sub
     Else
         selZADRESS0_Nb = selZADRESS0_Nb + 1
         If selZADRESS0_Nb > selZADRESS0_Max Then
             selZADRESS0_Max = selZADRESS0_Max + 50
             ReDim Preserve selZADRESS0(selZADRESS0_Max)
         End If
         
         selZADRESS0(selZADRESS0_Nb) = xZADRESS0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : selZADRESS0_SQL"


End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim meUnit As typeUnit
blnSetfocus = False

SrvCptP0_Amj = "00000000"
blnService_Enabled = True
Call BiaPgmAut_Init("BIA_GAFI", BIA_Gafi_Aut)

If Not IsNull(param_Init) Then cmdOK.Visible = False
cmdReset
blnSetfocus = True
Rem Modification D. ROSILLETTE du 02/01/2013
Rem les états GAFI ne sont plus édités quotidiennement
Rem sauf l'état des Clients suivis par le contentieux 12/11/2015
Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@BIA_GAFI":
            autoGAFI = True
            optEtat09 = True
            optEtat09_Click
            'Test du 07/02/2020 pour utiliser systématiquement le modèle Excel et non plus l'imprimante PDF
            xlsManual = True
            If appExcelPublic Is Nothing Then
                Set appExcelPublic = CreateObject("Excel.Application")
                appExcelPublic.Visible = False
                appExcelPublic.ControlCharacters = False
                appExcelPublic.Interactive = False
            End If
            '--- Fin du test ----------------
            If xlsManual Then
                cmdOk_Click_xlsManual
            Else
                cmdOk_Click
            End If
            
            If xlsManual Then
                If Not appExcelPublic Is Nothing Then
                    appExcelPublic.Quit
                    Set appExcelPublic = Nothing
                End If
                xlsManual = False
            End If

Rem               If paramEnvironnement = constProduction Then
Rem                   meUnit.Id = "DEON"
Rem                   Table_Unit meUnit
Rem                   Printer_Set meUnit.Printer
Rem               End If
Rem               SSTab1.Tab = 0
Rem               cmdOk_Quotidien_Click
               Unload Me
   Case Else:
End Select


End Sub

Public Sub zoneImpression_xlsManual(lFct As String, nbRows As Long, wsheet As Excel.Worksheet)

    Call init_TypePagesetup
    If nbRows > 0 Then
        wsheet.Activate
        wsheet.Range("A1:J" & CStr(nbRows)).Select
        zoneImpressionPagesetup.PrintArea = "$A$1:$J$" & CStr(nbRows)
        zoneImpressionPagesetup.LeftFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "prtSAB_GAFI   &D &T  BIA_INFO"
        zoneImpressionPagesetup.RightFooter = "&""Arial,Normal""&6&K04-024" & Chr(10) & "&P"
        zoneImpressionPagesetup.Orientation = xlLandscape
        zoneImpressionPagesetup.Zoom = 80
    End If
    Call SetTypePageSetup(wsheet)

End Sub

Public Function param_Init()
Dim V, X As String
Dim wNb As Long
Dim constRES As String

'Call lstErr_Clear(lstErr, cmdContext, "> Initialisation Comptes .....")
param_Init = Null
'------------------------------------------------------------------------------
arrCLIENARES_Nb = 0
X = "select * from YBIATAB0 where" _
    & " BIATABID = 'RESPONSABLE'" _
    & " and BIATABK1 like 'R%'"
    
Set rsMDB = cnMDB.Execute(X)
Do Until rsMDB.EOF
    arrCLIENARES_Nb = arrCLIENARES_Nb + 1
    arrCLIENARES(arrCLIENARES_Nb) = rsMDB("BIATABK1")
    rsMDB.MoveNext
Loop

Call lstErr_Clear(lstErr, cmdContext, "= NB Responsables  : " & arrCLIENARES_Nb)
'------------------------------------------------------------------------------

cboSelect_ADRESSCOA.AddItem "   "
cboSelect_ADRESSCOA.AddItem "<  >"
cboSelect_ADRESSCOA.AddItem "CO"
cboSelect_ADRESSCOA.AddItem "CH"
cboSelect_ADRESSCOA.ListIndex = 1

'_____________________________________________________________________________________
If BIA_Gafi_Aut.Xspécial Then
    fgParam_FormatString = fgParam.FormatString
    fraParam.Visible = True
    fraParam_Update.Visible = False
    fgParam.Clear
    txtParam_Id = ""
    
    lstParam.AddItem "3 - TRACFIN"
    lstParam.AddItem "4 - paradis fiscaux"
    lstParam.AddItem "5 - PTNC"
    lstParam.AddItem "6 - SFIx"
    lstParam.AddItem "7 - OBNL"
    lstParam.AddItem "8 - PEP"
    lstParam.AddItem "9 - Contentieux"
End If
    

End Function



Public Sub cmdControl()
If Not Me.Enabled Then Exit Sub
Me.Enabled = False

blnControl = False
lstErr.Clear
lstErr.Height = 200

vReturn = DTPicker_Control(txtAmjMin, wAMJMin)
If Not IsNull(vReturn) Then Call lstErr_AddItem(lstErr, txtAmjMin, vReturn): Exit Sub
vReturn = DTPicker_Control(txtAmjMax, WAMJMax)
If Not IsNull(vReturn) Then Call lstErr_AddItem(lstErr, txtAmjMax, vReturn): Exit Sub



If chkDb = "1" Then
    curDB = Abs(CCur(Val(txtDb)))
    blnDb = True
Else
    curDB = 0
    blnDb = False
End If

If chkCr = "1" Then
    curCR = Abs(CCur(Val(txtCr)))
    blnCr = True
Else
    curCR = 0
    blnCr = False
End If


ExitSub:

Me.Enabled = True
    
blnControl = True

End Sub
Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdExport_Click()
Dim K As Integer
Dim xFile As String, X As String
On Error GoTo Error_Handler

xFile = Trim(txtSelect_Folder) & Trim(txtSelect_FileName)
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> cmdExport_ ........"): DoEvents
Call FEU_ROUGE
Open xFile For Output As #2

For K = 1 To fgSelect.Rows - 1
    fgSelect.Row = K
    fgSelect.Col = 1: X = fgSelect.Text ''' xZADRESS0.ADRESSNUM
    fgSelect.Col = 2: X = X & ";" & fgSelect.Text 'ra
    fgSelect.Col = 3: X = X & ";" & fgSelect.Text  ' adresse
    X = Replace(X, "À", "A")
    X = Replace(X, "Â", "A")
    X = Replace(X, "Ä", "A")
    X = Replace(X, "É", "E")
    X = Replace(X, "È", "E")
    X = Replace(X, "Ê", "E")
    X = Replace(X, "Ë", "E")
    X = Replace(X, "Î", "I")
    X = Replace(X, "Ï", "I")
    X = Replace(X, "Ö", "O")
    X = Replace(X, "Ô", "O")
    X = Replace(X, "Ù", "U")
    X = Replace(X, "Û", "U")
    X = Replace(X, "Ü", "U")
    X = Replace(X, "'", " ")
    Print #2, X '& ";"
Next K
Close #2
Call FEU_VERT

MsgBox ("Export terminé !")


Call lstErr_AddItem(lstErr, cmdContext, "< cmdExport : " & fgSelect.Rows - 1): DoEvents
GoTo Exit_sub

Error_Handler:
    Close
    MsgBox Error, vbCritical, Me.Caption & " :  " & xFile
Exit_sub:
    Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdOk_Click()
Call lstErr_Clear(lstErr, cmdOK, "Début du traitement")
If blnControl Then cmdControl
If lstErr.ListCount <> 0 Then Exit Sub

cmdFlux
Me.Show

End Sub

Private Sub cmdOk_Quotidien_Click()
optEtat01 = True: optEtat01_Click: cmdOk_Click
optEtat03 = True: optEtat03_Click: cmdOk_Click
optEtat05 = True: optEtat05_Click: cmdOk_Click
'optEtat08 = True: optEtat08_Click: cmdOk_Click
If BIA_Gafi_Aut.Xspécial Then
    optEtat06 = True
    optEtat06_Click
    cmdOk_Click
End If

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
        Case "4", "5": New_YBIATAB0.BIATABK2 = X
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

Private Sub cmdParam_Print_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdPrint_Excel

Me.Enabled = True: Me.MousePointer = 0

End Sub

Public Sub cmdPrint_Excel()
On Error GoTo Error_Handler
Dim xSQL As String
Dim X As String, wFilex As String, wFile As String
Dim blnCALCS As Boolean

On Error GoTo Error_Handler
'===================================================================================
'______________________________________________'
X = "C:\Temp\"

mXls1_File = mXls1_File + 1

wFile = X & Trim("BIA_GAFI - paramétrage " & DSYS_Time & mXls1_File & ".xlsx")
'______________________________________________
'_________________________________________


If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile

'=========================================================================================
Call lstErr_AddItem(lstErr, cmdContext, "Fichier excel.... : "): DoEvents

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "BIA_GAFI"
    .Subject = "BIA_GAFI"
End With

'__________________________________________________________________________________

'appExcel.Worksheets.Add

Set wsExcel = wbExcel.Sheets(1): wsExcel.Name = "BIA_GAFI"

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

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14BIA_GAFI paramétrage au " & dateImp10(wAMJMin) _
                                 & vbCr

wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

mXls1_Row = 1

cmdPrint_Excel_YBIATAB0
        

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
    MsgBox Error, vbCritical, Me.Name
    Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents
    
    wbExcel.SaveAs wFile
    wbExcel.Close
    appExcel.Quit

End Sub


Private Sub cmdParam_Quit_Click()
fraParam_Update.Visible = False

End Sub

Private Sub cmdSelect_Ok_Click()
'???selZADRESS0_SQL
Dim blnOk As Boolean

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAb_Stock_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
If blnOk Then
    cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    fraSelect_Options.Enabled = False
    selZADRESS0_SQL
    fgSelect_Display
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    fraSelect_Options.Enabled = True
End If

Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
fraSelect_ADRESSTYP.BackColor = fraSelect_Options.BackColor
Call usrColor_Container(fraSelect_ADRESSTYP, fraSelect_Options.BackColor)
fraSelect_Scan.BackColor = fraSelect_Options.BackColor
Call usrColor_Container(fraSelect_Scan, fraSelect_Options.BackColor)

Call lstErr_AddItem(lstErr, cmdContext, "< BIA_GAFI_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmkOk_Mensuel_Click()
optEtat02 = True: optEtat02_Click: cmdOk_Click
optEtat05 = True: optEtat05_Click: cmdOk_Click
optEtat08 = True: optEtat08_Click: cmdOk_Click
optEtat07 = True: optEtat07_Click: cmdOk_Click

End Sub

Private Sub fgParam_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

If fgParam.Rows > 1 Then
    fgParam.Col = 0: txtParam_Id = Trim(fgParam.Text)
    cmdParam_Delete.Visible = True
End If

End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
   End If
End If
fgSelect.LeftCol = 0

End Sub

Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
Form_Init
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
Dim V, I As Long
Dim X As String, xSQL As String
Dim blnOk As Boolean, blnDisplay As Boolean, blnSkip As Boolean
Dim Trim_txtSelect_ADRESSAD1 As String

On Error GoTo Error_Handler
'SSTab1.Tab = 1
fgSelect.Visible = False
fgSelect_Reset
'cmdPrint.Enabled = False

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
'currentAction = "fgSelect_Display"
Set rsSab = Nothing
Call DTPicker_Amj7(txtSelect_COMPTEOUV, mAmj_COMPTEOUV)

Trim_txtSelect_ADRESSAD1 = Trim(txtSelect_ADRESSAD1)
If optSelect_Scan_Texte And Trim_txtSelect_ADRESSAD1 = "" Then
    blnSkip = True
Else
    blnSkip = False
End If

For I = 1 To selZADRESS0_Nb
    blnOk = False
    xZADRESS0 = selZADRESS0(I)
    X = Trim(UCase$(xZADRESS0.ADRESSAD1)) & " " & Trim(UCase$(xZADRESS0.ADRESSAD2)) & " " & Trim(UCase$(xZADRESS0.ADRESSAD3)) & ";" & Trim(UCase$(xZADRESS0.ADRESSCOP)) & ";" & Trim(UCase$(xZADRESS0.ADRESSVIL)) & ";" & Trim(UCase$(xZADRESS0.ADRESSPAY))
       
    If blnSkip Then
        blnOk = True
    Else
        If optSelect_Scan_Auto Then
            blnOk = optSelect_Scan_Auto_X(X)
        Else
            If InStr(1, X, Trim_txtSelect_ADRESSAD1) > 0 Then blnOk = True
    
        End If
    End If
    
    If blnOk Then
        
        If optSelect_ADRESSTYP_1 Then
            If xZCLIENA0_SQL(xZADRESS0.ADRESSNUM) Then
                If chkSelect_COMPTEOUV = "1" And xZCLIENA0.CLIENACRE < mAmj_COMPTEOUV Then blnOk = False
                If chkSelect_CLIENARES = "0" And Mid$(xZCLIENA0.CLIENARES, 1, 1) = "X" Then blnOk = False
                If chkSelect_CLIENARES_99 = "0" And Mid$(xZCLIENA0.CLIENACLI, 1, 1) = "9" Then blnOk = False
                If Mid$(xZCLIENA0.CLIENACLI, 1, 2) = "88" Then
                    If Not isProspect(xZCLIENA0.CLIENACLI) Then blnOk = False
                End If
            Else
                blnOk = False
            End If
        Else
            Call xZCOMPTE0_SQL(xZADRESS0.ADRESSNUM)
            If chkSelect_COMPTEOUV = "1" And xZCOMPTE0.COMPTEOUV < mAmj_COMPTEOUV Then blnOk = False
            If chkSelect_COMPTEFON = "0" And xZCOMPTE0.COMPTEFON = "4" Then
                blnOk = False
            End If
        End If
        
        If blnOk Then fgSelect_DisplayLine I, X
    End If


Next I

fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Détectés : " & fgSelect.Rows - 1 & " / " & selZADRESS0_Nb): DoEvents
If fgSelect.Rows > 1 Then
'   cmdPrint.Enabled = True
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : fgSelect_Display"


End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long, lX As String)
Dim wForecolor As Long
On Error Resume Next

Dim curX As Currency
Dim I As Integer, wAUTSYCAUT As String

wForecolor = vbRed

fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
fgSelect.Col = 0: fgSelect.Text = xZADRESS0.ADRESSCOA
fgSelect.Col = 1: fgSelect.Text = Trim(xZADRESS0.ADRESSNUM)
fgSelect.Col = 2:
If optSelect_ADRESSTYP_1 Then
    fgSelect.Text = Trim(xZCLIENA0.CLIENARA1) & " " & Trim(xZCLIENA0.CLIENARA2)
    fgSelect.Col = 4: fgSelect.Text = dateIBM10(xZCLIENA0.CLIENACRE, True)
Else
    fgSelect.Text = xZCOMPTE0.COMPTEINT
    fgSelect.Col = 4: fgSelect.Text = dateIBM10(xZCOMPTE0.COMPTEOUV, True)
End If
fgSelect.Col = 3: fgSelect.Text = lX


fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

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
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset
End Sub



Public Sub cmdContext_Return()

End Sub

Public Sub cmdContext_Quit()
If SSTab1.Tab = 2 Then
    If fraParam_Update.Visible Then fraParam_Update.Visible = False
Else
    Unload Me
End If
End Sub

Private Function cmdImport_MvtP0_Select() As String
Dim wNuméro As String * 11, wDeviseIso As String * 3, wAmjOpération As String * 8
Dim X1 As String * 1
Dim blnCptOrdinaire As Boolean, blnRIB As Boolean, blnMédiateur  As Boolean, blnIban As Boolean
cmdImport_MvtP0_Select = ""

meCV1.DeviseIso = xYBIAMVT0.COMPTEDEV

'=====================================
optEtat_Test:
' ignorer mvts TC & DAT

If xYBIAMVT0.MOUVEMSER = "TC" Then Exit Function
 
If xYBIAMVT0.MOUVEMOPE = "EM1" Then Exit Function   ' DAT

If xYBIAMVT0.MOUVEMSER = "CP" Then
    If xYBIAMVT0.MOUVEMOPE <> "-TR" Then Exit Function
End If

' comptes ordinaires

Call fctPCEC_Atribut(xYBIAMVT0.COMPTEOBL, meCV1.DeviseIso, blnCptOrdinaire, blnRIB, blnMédiateur, blnIban)
If Not blnCptOrdinaire Then Exit Function



 
optEtat_Ok:

curX = xYBIAMVT0.MOUVEMMON
If meCV1.DeviseIso <> "EUR" Then
    meCV1.DeviseN = 0
    meCV1.Montant = curX
    meCV1.OpéAmj = xYBIAMVT0.MOUVEMDTR + 19000000
    meCV2.OpéAmj = meCV1.OpéAmj
       
    Call CV_Calc("J  ", meCV1, meCV2)
    curX = meCV2.Montant
Else
    meCV2.Montant = curX
End If

If optEtat01 Then
    If curX > 0 Then
        If Abs(curX) < curDB Then Exit Function
    Else
        If Abs(curX) < curCR Then Exit Function
    End If
End If

cmdImport_MvtP0_Select = "OK"

End Function

Private Sub fraEtat_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraScript_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

End Sub


'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
currentActiveControl_Name = C.Name
End Sub
'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
lstErr.Clear
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub


Private Sub lstParam_Click()
Dim xSQL As String, X As String

fraParam_Update.Visible = False
cmdParam_Delete.Visible = False
cmdParam_Add.Visible = False
txtParam_Id = ""

Old_YBIATAB0.BIATABID = "BIA_GAFI"
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

Public Sub cmdPrint_Excel_YBIATAB0()
Dim xSQL As String, X As String, K As Long, wColor As Long, mK As Integer
On Error GoTo Error_Handler

'===================================================================================

wsExcel.Name = "BIA_GAFI"

With wsExcel.Cells
    .HorizontalAlignment = Excel.xlHAlignLeft
    .Font.Size = 9
    .Font.Name = "Courier New"
End With
wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 100
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14BIA_GAFI : paramétrage" _
                                & "  (édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.PrintTitleRows = "$A1:$E1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True


wsExcel.Columns(1).ColumnWidth = 20: wsExcel.Cells(1, 1) = "Filtre"
wsExcel.Columns(2).ColumnWidth = 10: wsExcel.Cells(1, 2) = "Code"
wsExcel.Columns(3).ColumnWidth = 50: wsExcel.Cells(1, 3) = "Intitulé"
wsExcel.Columns(4).ColumnWidth = 45: wsExcel.Cells(1, 4) = "MàJ"

For K = 1 To 4
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = vbWhite
Next

xSQL = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'BIA_GAFI'" _
     & " order by BIATABID , BIATABK1 , BIATABK2"
Set rsSab = cnsab.Execute(xSQL)
mK = 3
Do While Not rsSab.EOF

    mXls1_Row = mXls1_Row + 1
    lstParam_K = Trim(rsSab("BIATABK1"))
    If mK <> lstParam_K Then mXls1_Row = mXls1_Row + 1: mK = lstParam_K
    
    Select Case lstParam_K
        Case "3": wsExcel.Cells(mXls1_Row, 1) = "TRACFIN"
        Case "4": wsExcel.Cells(mXls1_Row, 1) = "paradis fiscaux"
        Case "5": wsExcel.Cells(mXls1_Row, 1) = "PTNC"
        Case "6": wsExcel.Cells(mXls1_Row, 1) = "SFIx"
        Case "7": wsExcel.Cells(mXls1_Row, 1) = "OBNL"
        Case "8": wsExcel.Cells(mXls1_Row, 1) = "PEP"
        Case "9": wsExcel.Cells(mXls1_Row, 1) = "Contentieux"
    End Select

    wsExcel.Cells(mXls1_Row, 1).Font.Color = mColor_GB
    wsExcel.Cells(mXls1_Row, 2) = rsSab("BIATABK2"): wsExcel.Cells(mXls1_Row, 2).Font.Color = vbBlue
    
    wsExcel.Cells(mXls1_Row, 4) = Trim(Mid$(rsSab("BIATABTXT"), 1, 99))
    wsExcel.Cells(mXls1_Row, 4).Font.Color = RGB(128, 128, 128)
    wsExcel.Cells(mXls1_Row, 3) = fgParam_Display_Lib(Trim(rsSab("BIATABK2")))
    wsExcel.Cells(mXls1_Row, 3).Font.Color = mColor_GB
    rsSab.MoveNext
Loop



'======================================================================================================

Exit_sub:
'__________________________________________________________________________________


'_____________________________
Exit Sub

Error_Handler:

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

Private Sub optEtat01_Click()
cmdReset
'chkCr = "1": txtCr = "150000"
'chkDb = "1": txtDb = "150000"

End Sub

Private Sub optEtat01_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set optEtat01
End Sub


Private Sub optEtat02_Click()

chkCr = "1": txtCr = "7000": chkCr.Caption = "Crédit & |Débit|  >  ******* EUR"
chkDb = "1": txtDb = "2000":: chkDb.Caption = "impr MVT >  ******* EUR"
Call DTPicker_Control(txtAmjMax, wAMJMin)
Mid$(wAMJMin, 7, 2) = "01"
Call DTPicker_Set(txtAmjMin, wAMJMin)
End Sub

Private Sub optEtat03_Click()
txtCr = "0": txtDb = "0"

End Sub

Private Sub optEtat04_Click()
chkCr = "1": txtCr = "7000": chkCr.Caption = "Crédit & |Débit|  >  ******* EUR"
chkDb = "1": txtDb = "2000":: chkDb.Caption = "impr MVT >  ******* EUR"
Call DTPicker_Control(txtAmjMax, wAMJMin)
Mid$(wAMJMin, 7, 2) = "01"
Call DTPicker_Set(txtAmjMin, wAMJMin)

End Sub

Private Sub optEtat05_Click()
txtCr = "0": txtDb = "0"

End Sub

Private Sub optEtat06_Click()
txtCr = "0": txtDb = "0"

End Sub

Private Sub optEtat07_Click()
txtCr = "0": txtDb = "0"

End Sub

Private Sub optEtat08_Click()
txtCr = "0": txtDb = "0"

End Sub

Private Sub optEtat09_Click()
    
    txtCr = "0": txtDb = "0"
    
End Sub

Private Sub optEtatManuel_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set optEtatManuel
End Sub


Public Sub cmdReset()
blnControl = False

SSTab1.Enabled = True 'BIA_Gafi_Aut.Saisir
blnYBIAMVT0_Import = False
DTPicker_Set txtAmjMin, YBIATAB0_DATE_CPT_JP0
DTPicker_Set txtAmjMax, YBIATAB0_DATE_CPT_J
chkDb = "1": chkDb.Caption = "Débit >  ******* EUR"
chkCr = "1": chkCr.Caption = "Crédit  >  ******* EUR"

txtCr = "150000": txtDb = "150000"

X1000 = ""
optEtat01.Value = True
optEtat06.Enabled = BIA_Gafi_Aut.Xspécial

meCV2.DeviseIso = "EUR"
chkSelect_COMPTEOUV = "0"
chkSelect_COMPTEFON.Enabled = False
Call DTPicker_Set(txtSelect_COMPTEOUV, YBIATAB0_DATE_CPT_J)

'migration SIDE2010
'txtSelect_Folder = " C:\Temp\"  '"\\SWIFTPROD\SWIFT_APPLI\Production\SIDE_File_Scanner\"
'txtSelect_FileName = "BIA_Client.txt"
txtSelect_Folder = "d:\BIASRV.DAT\Dat\"
txtSelect_FileName = "BIA_Client.txt"

blnControl = True

End Sub

Public Sub Form_Init()
Me.Enabled = False
SSTab1.Tab = 0
fraEtat.Enabled = False
fraScript.Enabled = True 'False
fraEtat.Enabled = True 'False

wAmj = dateElp("FinDeMoisP", 0, DSys)
Call DTPicker_Set(txtAmjMax, wAmj)
Mid$(wAmj, 7, 2) = "01"
Call DTPicker_Set(txtAmjMin, wAmj)
fgSelect_FormatString = fgSelect.FormatString
Me.Enabled = True

End Sub

Public Sub cmdFlux()
Dim K As Integer, K2 As Integer, xSQL As String, V
Dim xEtat As String
Dim kIndex As Long, selIndex As Long
Dim wMOUVEMCOM As String
Dim X As String

cmdImport_Init

Call lstErr_Clear(lstErr, cmdOK, "Flux : Début du traitement")


Me.MousePointer = vbHourglass
Me.Enabled = False
Call DTPicker_Control(txtAmjMin, wAMJMin)
xAmjMin_IBM = dateIBM(wAMJMin)
Call DTPicker_Control(txtAmjMax, WAMJMax)
xAmjMax_IBM = dateIBM(WAMJMax)

If optEtat03 Then cmdFlux_optEtat03: Exit Sub
If optEtat04 Then cmdFlux_optEtat04: Exit Sub

If optEtat05 Then cmdFlux_optEtat05: Exit Sub
If optEtat06 Then cmdFlux_optEtat06: Exit Sub
If optEtat07 Then cmdFlux_optEtat07: Exit Sub
If optEtat08 Then cmdFlux_optEtat08: Exit Sub
If optEtat09 Then
    'On passe systématiquement par xlsManual 16/06/222
        xlsManual = True
        If appExcelPublic Is Nothing Then
            Set appExcelPublic = CreateObject("Excel.Application")
            appExcelPublic.Visible = False
            appExcelPublic.ControlCharacters = False
            appExcelPublic.Interactive = False
        End If
        Call cmdImport09_MvtP0_xlsManual
        If Not appExcelPublic Is Nothing Then
            appExcelPublic.Quit
            Set appExcelPublic = Nothing
        End If
        xlsManual = False
    Exit Sub
End If
cmdImport_MvtP0

If optEtat01 Then xEtat = "01"
If optEtat02 Then xEtat = "02"
lstW.Clear
xYBIACPT0.COMPTECOM = ""

For K = 1 To selYBIAMVT0_Nb
    If xYBIACPT0.COMPTECOM <> selYBIAMVT0(K).MOUVEMCOM Then
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
             & " where COMPTECOM = '" & selYBIAMVT0(K).MOUVEMCOM & "'"
             
        Set rsSab = cnsab.Execute(xSQL)
        
        V = rsYBIACPT0_GetBuffer(rsSab, xYBIACPT0)
        If Not IsNull(V) Then
            MsgBox Error, vbCritical, "COMPTECOM = '" & selYBIAMVT0(K).MOUVEMCOM & "'"
        Else
            lstW.AddItem xYBIACPT0.CLIENARES & " " & Trim(xYBIACPT0.COMPTECOM) & " : " & K
        End If
    End If
Next K
lstW.AddItem "XXX:0"


mCLIENARES = ""
prtBIA_Gafi.arrYBIAMVT0_Nb = 0
ReDim prtBIA_Gafi.arrYBIAMVT0(selYBIAMVT0_Nb)

For K = 0 To lstW.ListCount - 1
    lstW.ListIndex = K
    X = lstW.Text
    kIndex = InStr(X, ":")
    If kIndex > 0 Then
        If mCLIENARES <> Mid$(X, 1, 3) Then
            If prtBIA_Gafi.arrYBIAMVT0_Nb > 0 Then
                Call lstErr_AddItem(lstErr, cmdOK, "Flux : " & mCLIENARES)
                prtBIA_Gafi_Monitor xEtat, curDB, curCR, wAMJMin, WAMJMax, mCLIENARES
            End If
            prtBIA_Gafi.arrYBIAMVT0_Nb = 0
            mCLIENARES = Mid$(X, 1, 3)
        End If
        
        selIndex = Val(Mid$(X, kIndex + 1, Len(X) - kIndex))
        If selIndex > 0 Then
            wMOUVEMCOM = selYBIAMVT0(selIndex).MOUVEMCOM
            For K2 = selIndex To selYBIAMVT0_Nb
                If wMOUVEMCOM = selYBIAMVT0(K2).MOUVEMCOM Then
                    prtBIA_Gafi.arrYBIAMVT0_Nb = prtBIA_Gafi.arrYBIAMVT0_Nb + 1
                    prtBIA_Gafi.arrYBIAMVT0(prtBIA_Gafi.arrYBIAMVT0_Nb) = selYBIAMVT0(K2)
                Else
                    Exit For
                End If
                
            Next K2
        End If
    End If
Next K


Me.MousePointer = 0
Me.Enabled = True



End Sub
Public Sub cmdImport_MvtP0()
Dim xSQL As String, X As String
Dim paramYBIAMVT0_Import As String
Dim wseq As Long

On Error Resume Next

Call lstErr_Clear(lstErr, cmdOK, "Chargement des mouvements ...")
Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH" _
     & " where MOUVEMDTR >= " & xAmjMin_IBM _
     & " and MOUVEMDTR <= " & xAmjMax_IBM _
     & " order by MOUVEMCOM,MOUVEMDTR,MOUVEMPIE,MOUVEMECR"

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVT0)
    vReturn = cmdImport_MvtP0_Select()
    If vReturn <> "" Then

        selYBIAMVT0_Nb = selYBIAMVT0_Nb + 1
        If selYBIAMVT0_Nb >= selYBIAMVT0_NbMax Then
            selYBIAMVT0_NbMax = selYBIAMVT0_NbMax + 500
            ReDim Preserve selYBIAMVT0(selYBIAMVT0_NbMax)
        End If
        selYBIAMVT0(selYBIAMVT0_Nb) = xYBIAMVT0
    End If
    rsSab.MoveNext
Loop
Call lstErr_Clear(lstErr, cmdOK, "Nb mouvements sélectionnés : " & arrYBIAMVT0_Nb)

End Sub


Public Function xZCLIENA0_SQL(lADRESSNUM As String) As Boolean
Dim xSQL As String, V

    X = "CLIENACLI =  '" & Trim(lADRESSNUM) & "'"
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 where " & X
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        V = rsZCLIENA0_GetBuffer(rsSab, xZCLIENA0)
        xZCLIENA0_SQL = True
        'If Not IsNull(V) Then rsZCLIENA0_Init xZCLIENA0
    Else
        xZCLIENA0_SQL = False
    End If

End Function

Public Function xZCOMPTE0_SQL(lADRESSNUM As String) As String
Dim xSQL As String, V
X = "COMPTECOM like  '" & Trim(lADRESSNUM) & "%'"
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCOMPTE0 where " & X

Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    V = rsZCOMPTE0_GetBuffer(rsSab, xZCOMPTE0)
    If Not IsNull(V) Then
        xZCOMPTE0_SQL = "???" & lADRESSNUM
    Else
         xZCOMPTE0_SQL = xZCOMPTE0.COMPTEINT
   End If
End If

End Function


Public Sub arrYBIACPT0_CLIENARSD(lCLIENARSD As String)
Dim xSQL As String, X As String
Dim blnOk As Boolean

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where CLIENARSD = '" & lCLIENARSD & "'"

arrYBIACPT0_SQL xSQL

End Sub
Public Sub arrYBIACPT0_COMPTECOM(lCOMPTECOM As String)
Dim xSQL As String

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '" & lCOMPTECOM & "'"

arrYBIACPT0_SQL xSQL

End Sub

Public Sub arrYBIACPT0_SQL(xSQL As String)
Dim blnOk As Boolean
Set rsSab = Nothing
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    X = rsSab("PLANCOPRO")
    Select Case X
        Case "NOS", "CCR", "CHB": blnOk = False
        Case Else: blnOk = True
    End Select
    If blnOk Then
        arrYBIACPT0_Nb = arrYBIACPT0_Nb + 1
        If arrYBIACPT0_Nb >= arrYBIACPT0_NbMax Then
            ReDim Preserve arrYBIACPT0(arrYBIACPT0_Nb + 50)
            arrYBIACPT0_NbMax = arrYBIACPT0_NbMax + 50
        End If
        Call rsYBIACPT0_GetBuffer(rsSab, arrYBIACPT0(arrYBIACPT0_Nb))
    End If
    rsSab.MoveNext
Loop

End Sub

Public Function optSelect_Scan_Auto_X(X As String) As Boolean
optSelect_Scan_Auto_X = False
    If InStr(1, X, "BIA") > 0 Then
        If InStr(1, X, "DABIA") = 0 Then
            If InStr(1, X, "FABIA") = 0 Then
                If InStr(1, X, "JABIA") = 0 Then
                    If InStr(1, X, "RABIA") = 0 Then
                        If InStr(1, X, "EL BIAR") = 0 Then
                          If InStr(1, X, "EL-BIAR") = 0 Then
                            If InStr(1, X, "GAMBIA") = 0 Then

                                    optSelect_Scan_Auto_X = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If InStr(1, X, "B.I.A") > 0 Then optSelect_Scan_Auto_X = True
    If InStr(1, X, "B I A") > 0 Then optSelect_Scan_Auto_X = True
    If InStr(1, X, "OSEVELT") > 0 Then If InStr(1, X, "67") > 0 Then optSelect_Scan_Auto_X = True
    

End Function

Private Sub optSelect_ADRESSTYP_1_Click()
chkSelect_COMPTEFON.Enabled = False

End Sub

Private Sub optSelect_ADRESSTYP_2_Click()
chkSelect_COMPTEFON.Enabled = True

End Sub


Private Sub txtParam_Id_KeyPress(KeyAscii As Integer)
Select Case lstParam_K
    Case "4", "5": KeyAscii = convUCase(KeyAscii)
    Case Else: num_KeyAscii KeyAscii
End Select

End Sub


Private Sub txtSelect_ADRESSAD1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub chkSelect_COMPTEOUV_Click()
If Me.Enabled Then fgSelect.Clear
If chkSelect_COMPTEOUV = "0" Then
    txtSelect_COMPTEOUV.Visible = False
Else
    txtSelect_COMPTEOUV.Visible = True
End If

End Sub


Public Sub cmdFlux_optEtat04()
Dim xSQL As String
cmdImport_Init

'Actualisé le 09/03/2010 (JR-MF)
'arrYBIACPT0_CLIENARSD "AE"
'arrYBIACPT0_CLIENARSD "BS"
'arrYBIACPT0_CLIENARSD "CH"
'arrYBIACPT0_CLIENARSD "GI"
'arrYBIACPT0_CLIENARSD "HK"
'arrYBIACPT0_CLIENARSD "IE"
'arrYBIACPT0_CLIENARSD "KY"
'arrYBIACPT0_CLIENARSD "LI"
'arrYBIACPT0_CLIENARSD "LU"
'arrYBIACPT0_CLIENARSD "MT"
'arrYBIACPT0_CLIENARSD "PA"
'arrYBIACPT0_CLIENARSD "VG"
'...........................................................
'$JPL 2011-09-08

'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "AI"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "BZ"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "BN"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "CR"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "DM"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "GD"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "GT"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "CK"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "MH"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "LR"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "MS"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "NR"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "NU"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "PA"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "PH"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "KN"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "LC"
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "VC"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_GAFI' and BIATABK1 = '4' order by BIATABK2"
Set rsParam = cnsab.Execute(xSQL)

X = ""
Do While Not rsParam.EOF
    arrYBIACPT0_CLIENARSD Trim(rsParam("BIATABK2"))
    rsParam.MoveNext
Loop


'$JPL 2011-09-08
'...........................................................

For arrYBIACPT0_Index = 1 To arrYBIACPT0_Nb
    cmdImport_arrYBIAMVT0_SQL
Next arrYBIACPT0_Index

prtBIA_Gafi_Monitor "04", curDB, curCR, wAMJMin, WAMJMax, ""

Me.MousePointer = 0
Me.Enabled = True

End Sub
Public Sub cmdFlux_optEtat03()
Dim xSQL As String, X As String
cmdImport_Init

'...........................................................
'$JPL 2011-09-08
'xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where PLANCOPRO = 'CAV' " _
'     & "and ( COMPTECOM LIKE '12044%' or COMPTECOM LIKE '12389%' or COMPTECOM LIKE '11507%'" _
'     & "or  COMPTECOM LIKE '12423%' or COMPTECOM LIKE '50085%' or COMPTECOM LIKE '50089%'" _
'     & "or  COMPTECOM LIKE '50457%' or  COMPTECOM LIKE '12300%'" _
'     & "or  COMPTECOM LIKE '50513%' or  COMPTECOM LIKE '50159%' or CLIENACLI = '0050120')" _
'     & " order by COMPTECOM "
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_GAFI' and BIATABK1 = '3' order by BIATABK2"
Set rsParam = cnsab.Execute(xSQL)

X = ""
Do While Not rsParam.EOF
    X = X & "'" & Trim(rsParam("BIATABK2")) & "',"
    rsParam.MoveNext
Loop

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where PLANCOPRO = 'CAV'" _
     & " and CLIENACLI in (" & X & "'FIN')" _
     & " order by COMPTECOM "

'$JPL 2011-09-08
'...........................................................

arrYBIACPT0_SQL xSQL

'arrYBIACPT0_COMPTECOM "12044978001"
'arrYBIACPT0_COMPTECOM "12389978001"
'arrYBIACPT0_COMPTECOM "11507978001"
'arrYBIACPT0_COMPTECOM "12423978001"
'arrYBIACPT0_COMPTECOM "12423400001"
'arrYBIACPT0_COMPTECOM "50085978001"
'arrYBIACPT0_COMPTECOM "50085400001"
'arrYBIACPT0_COMPTECOM "50089978001"
'arrYBIACPT0_COMPTECOM "50089400001"
'arrYBIACPT0_COMPTECOM "50457978001"
'arrYBIACPT0_COMPTECOM "50457400001"

For arrYBIACPT0_Index = 1 To arrYBIACPT0_Nb
    cmdImport_arrYBIAMVT0_SQL
Next arrYBIACPT0_Index

prtBIA_Gafi_Monitor "03", curDB, curCR, wAMJMin, WAMJMax, ""

Me.MousePointer = 0
Me.Enabled = True


End Sub

Public Sub cmdFlux_optEtat05()
Dim K As Integer, xSQL As String
Dim X As String
Dim kIndex As Long
cmdImport_Init

'$20060628 JPL$ arrYBIACPT0_CLIENARSD "CK"   ' iles cook
'$20060628 JPL$ 'arrYBIACPT0_CLIENARSD "ID"   ' indonesie
'$JPL 2011-09-08 arrYBIACPT0_CLIENARSD "MM"   'myanmar
'$20060628 JPL$ 'arrYBIACPT0_CLIENARSD "NG"   'nigeria
'$20060628 JPL$ 'arrYBIACPT0_CLIENARSD "NR"   'nauru
'$20060628 JPL$ 'arrYBIACPT0_CLIENARSD "PH"   ' philippines

'...........................................................
'$JPL 2011-09-08
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_GAFI' and BIATABK1 = '5' order by BIATABK2"
Set rsParam = cnsab.Execute(xSQL)

X = ""
Do While Not rsParam.EOF
    arrYBIACPT0_CLIENARSD Trim(rsParam("BIATABK2"))
    rsParam.MoveNext
Loop


'$JPL 2011-09-08
'...........................................................

lstW.Clear
For K = 1 To arrYBIACPT0_Nb
    lstW.AddItem arrYBIACPT0(K).CLIENARES & " " & Trim(arrYBIACPT0(K).COMPTECOM) & " : " & K
Next K
lstW.AddItem "XXX:0"

mCLIENARES = ""

For K = 0 To lstW.ListCount - 1
    lstW.ListIndex = K
    X = lstW.Text
    kIndex = InStr(X, ":")
    If kIndex > 0 Then
        arrYBIACPT0_Index = Val(Mid$(X, kIndex + 1, Len(X) - kIndex))
        xYBIACPT0 = arrYBIACPT0(arrYBIACPT0_Index)
        If mCLIENARES <> xYBIACPT0.CLIENARES Then
            If prtBIA_Gafi.arrYBIAMVT0_Nb > 0 Then
                Call lstErr_AddItem(lstErr, cmdOK, "Flux : " & mCLIENARES)
                
                prtBIA_Gafi_Monitor "05", curDB, curCR, wAMJMin, WAMJMax, mCLIENARES
            End If
            prtBIA_Gafi.arrYBIAMVT0_Nb = 0
            mCLIENARES = xYBIACPT0.CLIENARES
        End If
        cmdImport_arrYBIAMVT0_SQL
    End If
Next K

Me.MousePointer = 0
Me.Enabled = True

End Sub

Public Sub cmdFlux_optEtat06()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim K As Integer, xSQL As String
Dim kIndex As Long

cmdImport_Init

'=======================================================================
'Recherche des comptes CAV / client '******'

' accorder le droit XSPECIAL au responsable du service déontologie
'=======================================================================

Set rsSab = Nothing
' Le 07.03.2005
' xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where PLANCOPRO = 'CAV' and COMPTEINT like '%BOULAAYOU%'"
'$2007-08-08 $jpl suppression or CLIENACLI = '0050227'

'...........................................................
'$JPL 2011-09-08

'xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where PLANCOPRO = 'CAV' " _
'     & "and ( CLIENACLI = '0011834'  or CLIENACLI = '0050050'" _
'     & "or CLIENACLI = '0012482' or CLIENACLI = '0050343' or CLIENACLI = '0050086' " _
'     & "or CLIENACLI = '0050198' or CLIENACLI = '0050056' or CLIENACLI = '0050064' " _
'     & "or CLIENACLI = '0050847' or CLIENACLI = '0050789' or CLIENACLI = '0050823' " _
'     & "or CLIENACLI = '0050137' or CLIENACLI = '0050737' or CLIENACLI = '0050881' " _
'     & "or CLIENACLI = '0011803' or CLIENACLI = '0050926' or CLIENACLI = '0012043' " _
'     & "or CLIENACLI = '0050929' or CLIENACLI = '0050813' or CLIENACLI = '0050271'  or CLIENACLI = '0050551' " _
'     & "or CLIENACLI = '0050218' or CLIENACLI = '0050309') "
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_GAFI' and BIATABK1 = '6' order by BIATABK2"
Set rsParam = cnsab.Execute(xSQL)

X = ""
Do While Not rsParam.EOF
    X = X & "'" & Trim(rsParam("BIATABK2")) & "',"
    rsParam.MoveNext
Loop

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where PLANCOPRO = 'CAV'" _
     & " and CLIENACLI in (" & X & "'FIN')" _
     & " order by COMPTECOM "

'$JPL 2011-09-08
'...........................................................
Set rsSab = cnsab.Execute(xSQL)


Do While Not rsSab.EOF
    arrYBIACPT0_Nb = arrYBIACPT0_Nb + 1
    If arrYBIACPT0_Nb >= arrYBIACPT0_NbMax Then
        ReDim Preserve arrYBIACPT0(arrYBIACPT0_Nb + 50)
        arrYBIACPT0_NbMax = arrYBIACPT0_NbMax + 50
    End If
    Call rsYBIACPT0_GetBuffer(rsSab, arrYBIACPT0(arrYBIACPT0_Nb))
    rsSab.MoveNext
Loop

lstW.Clear
For K = 1 To arrYBIACPT0_Nb
    lstW.AddItem arrYBIACPT0(K).CLIENARES & " " & Trim(arrYBIACPT0(K).COMPTECOM) & " : " & K
Next K
lstW.AddItem "XXX:0"

mCLIENARES = ""

For K = 0 To lstW.ListCount - 1
    lstW.ListIndex = K
    X = lstW.Text
    kIndex = InStr(X, ":")
    If kIndex > 0 Then
        arrYBIACPT0_Index = Val(Mid$(X, kIndex + 1, Len(X) - kIndex))
        xYBIACPT0 = arrYBIACPT0(arrYBIACPT0_Index)
        If mCLIENARES <> xYBIACPT0.CLIENARES Then
            If prtBIA_Gafi.arrYBIAMVT0_Nb > 0 Then
                Call lstErr_AddItem(lstErr, cmdOK, "Flux : " & mCLIENARES)
                
                prtBIA_Gafi_Monitor "06", curDB, curCR, wAMJMin, WAMJMax, mCLIENARES
            End If
            prtBIA_Gafi.arrYBIAMVT0_Nb = 0
            mCLIENARES = xYBIACPT0.CLIENARES
        End If
        cmdImport_arrYBIAMVT0_SQL
    End If
Next K
'__________________________________________________________________
' 2006.04.18 : éclatement en deux groupes X & OBNL
'2007-11-15 trt mensuel cmdFlux_optEtat07
'________________________________________________________________
Me.MousePointer = 0
Me.Enabled = True

End Sub

Public Sub cmdFlux_optEtat07()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim K As Integer, xSQL As String
Dim kIndex As Long

cmdImport_Init

'=======================================================================
'Recherche des comptes CAV / client '******'

' OBNL  (association etc ....)
'=======================================================================

Set rsSab = Nothing
'...........................................................
'$JPL 2011-09-08
'xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where PLANCOPRO = 'CAV' " _
'     & "and ( CLIENACLI = '0012313' or CLIENACLI = '0012316' or CLIENACLI = '0012319'" _
'     & "or  CLIENACLI = '0012444' or CLIENACLI = '0012473' " _
'     & "or CLIENACLI = '0012474') "

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_GAFI' and BIATABK1 = '7' order by BIATABK2"
Set rsParam = cnsab.Execute(xSQL)

X = ""
Do While Not rsParam.EOF
    X = X & "'" & Trim(rsParam("BIATABK2")) & "',"
    rsParam.MoveNext
Loop

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where PLANCOPRO = 'CAV'" _
     & " and CLIENACLI in (" & X & "'FIN')" _
     & " order by COMPTECOM "

'$JPL 2011-09-08
'...........................................................

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    arrYBIACPT0_Nb = arrYBIACPT0_Nb + 1
    If arrYBIACPT0_Nb >= arrYBIACPT0_NbMax Then
        ReDim Preserve arrYBIACPT0(arrYBIACPT0_Nb + 50)
        arrYBIACPT0_NbMax = arrYBIACPT0_NbMax + 50
    End If
    Call rsYBIACPT0_GetBuffer(rsSab, arrYBIACPT0(arrYBIACPT0_Nb))
    rsSab.MoveNext
Loop

lstW.Clear
For K = 1 To arrYBIACPT0_Nb
    lstW.AddItem arrYBIACPT0(K).CLIENARES & " " & Trim(arrYBIACPT0(K).COMPTECOM) & " : " & K
Next K
lstW.AddItem "XXX:0"

mCLIENARES = ""

For K = 0 To lstW.ListCount - 1
    lstW.ListIndex = K
    X = lstW.Text
    kIndex = InStr(X, ":")
    If kIndex > 0 Then
        arrYBIACPT0_Index = Val(Mid$(X, kIndex + 1, Len(X) - kIndex))
        xYBIACPT0 = arrYBIACPT0(arrYBIACPT0_Index)
        If mCLIENARES <> xYBIACPT0.CLIENARES Then
            If prtBIA_Gafi.arrYBIAMVT0_Nb > 0 Then
                Call lstErr_AddItem(lstErr, cmdOK, "Flux : " & mCLIENARES)
                
                prtBIA_Gafi_Monitor "07", curDB, curCR, wAMJMin, WAMJMax, mCLIENARES
            End If
            prtBIA_Gafi.arrYBIAMVT0_Nb = 0
            mCLIENARES = xYBIACPT0.CLIENARES
        End If
        cmdImport_arrYBIAMVT0_SQL
    End If
Next K

Me.MousePointer = 0
Me.Enabled = True

End Sub
Public Sub cmdFlux_optEtat08()
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim K As Integer, xSQL As String
Dim kIndex As Long

'modif denis 25/05/2010
Dim tbl As String

cmdImport_Init

'=======================================================================
'Recherche des comptes CAV / client '******'

'$20060622 JPL : PEP
'xxxxAjout 30/11/2009 Denis R.
' 24/02/2010 JR/MF
' 04/03/2010 JR (Voir Mail Mme DEZOTEUX du 03/03/2010 à 16:16:00)
'=======================================================================
tbl = "('0050072','0011803','0011843','0011863','0011870','0011892','0012109','0012110','0012157',"
tbl = tbl & "'0050026','0050138','0050218','0050258','0050337','0050404','0050412','0050427',"
tbl = tbl & "'0050458','0050538','0012427','0050363','0012043','0050780','0011956','0050846',"
tbl = tbl & "'0050881','0050891','0012089','0011962','0012017','0011755','0011943','0012130',"
tbl = tbl & "'0012309','0011781','0050632','0012027','0011773','0050408','0011935','0012179',"
tbl = tbl & "'0050261','0050309','0050909','0050384','0050834','0011713','0050926','0012309',"
tbl = tbl & "'0011954')"

Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where PLANCOPRO = 'CAV' and CLIENACLI IN " & tbl
     
'xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where PLANCOPRO = 'CAV' " _
     & "and ( CLIENACLI = '0050072'" _
     & " or CLIENACLI = '0011803'" & " or CLIENACLI = '0011843'" _
     & " or CLIENACLI = '0011863'" & " or CLIENACLI = '0011870'" _
     & " or CLIENACLI = '0011892'" & " or CLIENACLI = '0012109'" _
     & " or CLIENACLI = '0012110'" & " or CLIENACLI = '0012157'" _
     & " or CLIENACLI = '0050026'" & " or CLIENACLI = '0050138'" _
     & " or CLIENACLI = '0050218'" & " or CLIENACLI = '0050258'" _
     & " or CLIENACLI = '0050337'" & " or CLIENACLI = '0050404'" _
     & " or CLIENACLI = '0050412'" & " or CLIENACLI = '0050427'" _
     & " or CLIENACLI = '0050458'" & " or CLIENACLI = '0050538'" _
     & " or CLIENACLI = '0012427'" & " or CLIENACLI = '0050363'" _
     & " or CLIENACLI = '0012043'" & " or CLIENACLI = '0050780'" _
     & " or CLIENACLI = '0011956'" & " or CLIENACLI = '0050846'" _
     & " or CLIENACLI = '0050881'" & " or CLIENACLI = '0050891'" _
     & " or CLIENACLI = '0012089'" & " or CLIENACLI = '0011962'" _
     & " or CLIENACLI = '0012017'" & " or CLIENACLI = '0011755'" _
     & " or CLIENACLI = '0011943'" & " or CLIENACLI = '0012130'" _
     & " or CLIENACLI = '0012309'" & " or CLIENACLI = '0011781'" _
     & " or CLIENACLI = '0050632'" & " or CLIENACLI = '0012027'" _
     & " or CLIENACLI = '0011773'" & " or CLIENACLI = '0050408'" _
     & " or CLIENACLI = '0011935'" & " or CLIENACLI = '0012179'" _
     & " or CLIENACLI = '0050261'" & " or CLIENACLI = '0050309'" _
     & " or CLIENACLI = '0050909'" & " or CLIENACLI = '0050384'" _
     & " or CLIENACLI = '0050834'" & " or CLIENACLI = '0011713'" & ")"
'...........................................................
'$JPL 2011-09-08
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_GAFI' and BIATABK1 = '8' order by BIATABK2"
Set rsParam = cnsab.Execute(xSQL)

X = ""
Do While Not rsParam.EOF
    X = X & "'" & Trim(rsParam("BIATABK2")) & "',"
    rsParam.MoveNext
Loop

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 where PLANCOPRO = 'CAV'" _
     & " and CLIENACLI in (" & X & "'FIN')" _
     & " order by COMPTECOM "

'$JPL 2011-09-08
'...........................................................


Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    arrYBIACPT0_Nb = arrYBIACPT0_Nb + 1
    If arrYBIACPT0_Nb >= arrYBIACPT0_NbMax Then
        ReDim Preserve arrYBIACPT0(arrYBIACPT0_Nb + 50)
        arrYBIACPT0_NbMax = arrYBIACPT0_NbMax + 50
    End If
    Call rsYBIACPT0_GetBuffer(rsSab, arrYBIACPT0(arrYBIACPT0_Nb))
    rsSab.MoveNext
Loop

lstW.Clear
For K = 1 To arrYBIACPT0_Nb
    lstW.AddItem arrYBIACPT0(K).CLIENARES & " " & Trim(arrYBIACPT0(K).COMPTECOM) & " : " & K
Next K
lstW.AddItem "XXX:0"

mCLIENARES = ""

For K = 0 To lstW.ListCount - 1
    lstW.ListIndex = K
    X = lstW.Text
    kIndex = InStr(X, ":")
    If kIndex > 0 Then
        arrYBIACPT0_Index = Val(Mid$(X, kIndex + 1, Len(X) - kIndex))
        xYBIACPT0 = arrYBIACPT0(arrYBIACPT0_Index)
        If mCLIENARES <> xYBIACPT0.CLIENARES Then
            If prtBIA_Gafi.arrYBIAMVT0_Nb > 0 Then
                Call lstErr_AddItem(lstErr, cmdOK, "Flux : " & mCLIENARES)
                
                prtBIA_Gafi_Monitor "08", curDB, curCR, wAMJMin, WAMJMax, mCLIENARES
            End If
            prtBIA_Gafi.arrYBIAMVT0_Nb = 0
            mCLIENARES = xYBIACPT0.CLIENARES
        End If
        cmdImport_arrYBIAMVT0_SQL
    End If
Next K

Me.MousePointer = 0
Me.Enabled = True

End Sub


Public Sub cmdImport_Init()

ReDim prtBIA_Gafi.arrYBIAMVT0(1000)
prtBIA_Gafi.arrYBIAMVT0_Nb = 0
prtBIA_Gafi.arrYBIAMVT0_Max = 1000

ReDim selYBIAMVT0(1000)
selYBIAMVT0_Nb = 0
selYBIAMVT0_NbMax = 1000

arrYBIACPT0_Nb = 0
arrYBIACPT0_NbMax = 50
ReDim arrYBIACPT0(51)
rsYBIACPT0_Init arrYBIACPT0(0): arrYBIACPT0_Index = 0

End Sub

Public Sub cmdImport_arrYBIAMVT0_SQL()
Dim xSQL As String, X As String
Dim blnOk As Boolean

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH" _
     & " where MOUVEMCOM = '" & arrYBIACPT0(arrYBIACPT0_Index).COMPTECOM & "'" _
     & " and MOUVEMDTR >= " & xAmjMin_IBM _
     & " and MOUVEMDTR <= " & xAmjMax_IBM _
     & " order by MOUVEMCOM,MOUVEMDTR,MOUVEMPIE,MOUVEMECR"

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    prtBIA_Gafi.arrYBIAMVT0_Nb = prtBIA_Gafi.arrYBIAMVT0_Nb + 1
    If prtBIA_Gafi.arrYBIAMVT0_Nb >= prtBIA_Gafi.arrYBIAMVT0_Max Then
        ReDim Preserve prtBIA_Gafi.arrYBIAMVT0(arrYBIAMVT0_Nb + 500)
        prtBIA_Gafi.arrYBIAMVT0_Max = prtBIA_Gafi.arrYBIAMVT0_Max + 500
    End If
    Call rsYBIAMVT0_GetBuffer(rsSab, prtBIA_Gafi.arrYBIAMVT0(arrYBIAMVT0_Nb))
    rsSab.MoveNext
Loop

End Sub


Public Function fgParam_Display_Lib(lBIATABK2 As String) As String
Dim xSQL As String

fgParam_Display_Lib = "?"

Select Case lstParam_K
    Case "4", "5"
    
        xSQL = "select * from  " & paramIBM_Library_SAB & ".ZBASTAB0 " _
            & " where BASTABETA = 1 and BASTABNUM = 11 and BASTABARG = 'CLI" & lBIATABK2 & "'"
        Set rsSabX = cnsab.Execute(xSQL)
        If Not rsSabX.EOF Then
            fgParam_Display_Lib = Mid$(rsSabX("BASTABLO2"), 4, 16)
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
