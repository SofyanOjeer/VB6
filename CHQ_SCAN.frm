VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCHQ_SCAN 
   AutoRedraw      =   -1  'True
   Caption         =   "CHQ_SCAN"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13560
   Icon            =   "CHQ_SCAN.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9150
   ScaleWidth      =   13560
   Begin VB.Frame fraCHQ 
      Height          =   5235
      Left            =   405
      TabIndex        =   61
      Top             =   3465
      Width           =   9960
      Begin VB.CommandButton cmdimgCHQ 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Recto/verso"
         Height          =   405
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   240
         Width           =   1815
      End
      Begin VB.Image imgCHQ_Verso 
         Height          =   4215
         Left            =   240
         Stretch         =   -1  'True
         Top             =   840
         Width           =   9495
      End
      Begin VB.Image imgCHQ 
         Height          =   4365
         Left            =   120
         Stretch         =   -1  'True
         Top             =   720
         Width           =   9750
      End
   End
   Begin VB.Frame fraPériode 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Préciser la période d'interrogation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   720
      TabIndex        =   4
      Top             =   6600
      Width           =   5415
      Begin VB.CommandButton cmdPériode_Ok 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ok"
         Height          =   525
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker txtPériode_Min 
         Height          =   300
         Left            =   1920
         TabIndex        =   6
         Top             =   480
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
         Format          =   101646339
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   -328351
      End
      Begin MSComCtl2.DTPicker txtPériode_Max 
         Height          =   300
         Left            =   3840
         TabIndex        =   7
         Top             =   480
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
         Format          =   101646339
         CurrentDate     =   36299
         MaxDate         =   401768
         MinDate         =   -328351
      End
      Begin VB.Label lblCHQ_Stat 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Période"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   2
      Top             =   0
      Width           =   5175
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
      Picture         =   "CHQ_SCAN.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   15266
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Suivi des Opérations"
      TabPicture(0)   =   "CHQ_SCAN.frx":0544
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fgSave"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTab0"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Affichage des traitements en cours"
      TabPicture(1)   =   "CHQ_SCAN.frx":0560
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraStatut"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Rapprochement SAB / CHQ_SCAN "
      TabPicture(2)   =   "CHQ_SCAN.frx":057C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDétail"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   13290
         Begin MSFlexGridLib.MSFlexGrid fgDossier 
            Height          =   6885
            Left            =   5760
            TabIndex        =   34
            Top             =   1200
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   12144
            _Version        =   393216
            Rows            =   1
            Cols            =   10
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   15269886
            ForeColor       =   8388608
            BackColorFixed  =   12648447
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   15269886
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"CHQ_SCAN.frx":0598
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Caption         =   "Exécuter la requête"
            Height          =   645
            Left            =   11880
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   240
            Width           =   1095
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   11355
            Begin VB.CheckBox chkSelect_GCC 
               Caption         =   "inclure 'GCC'"
               Height          =   192
               Left            =   8400
               TabIndex        =   63
               Top             =   720
               Width           =   1572
            End
            Begin VB.TextBox txtSelect_RefInterne 
               Height          =   285
               Left            =   6120
               TabIndex        =   55
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton optSelect_Archive 
               Caption         =   "Archive (serveur)"
               Height          =   255
               Left            =   240
               TabIndex        =   54
               Top             =   240
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton optSelect_Local 
               Caption         =   "poste local"
               Height          =   195
               Left            =   240
               TabIndex        =   53
               Top             =   600
               Width           =   1815
            End
            Begin VB.CheckBox chkSelect_Date 
               Caption         =   "Date saisie"
               Height          =   255
               Left            =   8400
               TabIndex        =   52
               Top             =   240
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.TextBox txtSelect_COMPTE 
               Height          =   285
               Left            =   6120
               TabIndex        =   51
               Top             =   600
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker txtSelect_Date 
               Height          =   300
               Left            =   9840
               TabIndex        =   56
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
               Format          =   101646339
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblSelect_RefInterne 
               Caption         =   "RefInterne"
               Height          =   255
               Left            =   5040
               TabIndex        =   58
               Top             =   240
               Width           =   1080
            End
            Begin VB.Label lblSelect_COMPTE 
               Caption         =   "Compte"
               Height          =   255
               Left            =   5040
               TabIndex        =   57
               Top             =   600
               Width           =   855
            End
         End
         Begin VB.Frame fraSelect_Update 
            Caption         =   "Mise à jour"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6255
            Left            =   8160
            TabIndex        =   35
            Top             =   1680
            Width           =   4815
            Begin VB.TextBox txtSelect_Update_Compte 
               Height          =   285
               Left            =   1800
               TabIndex        =   44
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtSelect_Update_RefClient 
               Height          =   285
               Left            =   1800
               TabIndex        =   43
               Top             =   1200
               Width           =   2775
            End
            Begin VB.TextBox txtSelect_Update_RefInterne 
               Height          =   285
               Left            =   1800
               TabIndex        =   42
               Top             =   1920
               Width           =   2055
            End
            Begin VB.TextBox txtSelect_Update_Nature 
               Height          =   285
               Left            =   1800
               TabIndex        =   41
               Top             =   2520
               Width           =   735
            End
            Begin VB.TextBox txtSelect_Update_Devise 
               Height          =   285
               Left            =   1800
               TabIndex        =   40
               Top             =   3240
               Width           =   735
            End
            Begin VB.CommandButton cmdSelect_Update_Ok 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Enregistrer"
               Height          =   885
               Left            =   1560
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   4680
               Width           =   1575
            End
            Begin VB.CommandButton cmdSelect_Update_Quit 
               BackColor       =   &H008080FF&
               Caption         =   "Abandonner"
               Height          =   645
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   5520
               Width           =   1095
            End
            Begin VB.CommandButton cmdSelect_Delete 
               BackColor       =   &H000000FF&
               Caption         =   "Supprimer la remise"
               Height          =   645
               Left            =   3360
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   5400
               Width           =   1095
            End
            Begin VB.CheckBox chkSelect_Update_StatutRem 
               Alignment       =   1  'Right Justify
               Caption         =   "remise ajustée"
               Height          =   195
               Left            =   240
               TabIndex        =   36
               Top             =   3840
               Width           =   1800
            End
            Begin VB.Label lblSelect_Update_Compte 
               Caption         =   "Compte"
               Height          =   255
               Left            =   240
               TabIndex        =   49
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label lblSelect_Update_RefCclient 
               Caption         =   "Référence Client"
               Height          =   255
               Left            =   240
               TabIndex        =   48
               Top             =   1200
               Width           =   1335
            End
            Begin VB.Label lblSelect_Update_RefInterne 
               Caption         =   "Référence interne"
               Height          =   255
               Left            =   240
               TabIndex        =   47
               Top             =   1920
               Width           =   1335
            End
            Begin VB.Label Label3lblSelect_Update_Nature 
               Caption         =   "Nature"
               Height          =   255
               Left            =   240
               TabIndex        =   46
               Top             =   2640
               Width           =   1335
            End
            Begin VB.Label lblSelect_Update_Devise 
               Caption         =   "Devise"
               Height          =   255
               Left            =   240
               TabIndex        =   45
               Top             =   3360
               Width           =   1335
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6945
            Left            =   120
            TabIndex        =   60
            Top             =   1200
            Width           =   12840
            _ExtentX        =   22648
            _ExtentY        =   12250
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
            FormatString    =   $"CHQ_SCAN.frx":067D
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
      Begin VB.Frame fraStatut 
         Height          =   7695
         Left            =   -74760
         TabIndex        =   17
         Top             =   480
         Width           =   12495
         Begin VB.CommandButton cmdStatut_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Actualiser"
            Height          =   645
            Left            =   10560
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   360
            Width           =   1095
         End
         Begin VB.FileListBox filDoc_Archive 
            ForeColor       =   &H00008000&
            Height          =   2625
            Left            =   240
            TabIndex        =   21
            Top             =   3720
            Visible         =   0   'False
            Width           =   3045
         End
         Begin VB.DirListBox dirListBox 
            Height          =   2115
            Left            =   7320
            TabIndex        =   20
            Top             =   3720
            Width           =   3735
         End
         Begin VB.FileListBox filDoc_Tampon 
            ForeColor       =   &H00008000&
            Height          =   2625
            Left            =   3720
            TabIndex        =   19
            Top             =   3720
            Visible         =   0   'False
            Width           =   3045
         End
         Begin VB.FileListBox filDoc 
            ForeColor       =   &H00008000&
            Height          =   870
            Left            =   7320
            TabIndex        =   18
            Top             =   6000
            Visible         =   0   'False
            Width           =   3765
         End
         Begin VB.Label lblStatut_DreamFile_ini 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblStatut_DreamFile_ini"
            Height          =   375
            Left            =   240
            TabIndex        =   32
            Top             =   1800
            Width           =   6495
         End
         Begin VB.Label libStatut_DreamFile_ini 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            Height          =   375
            Left            =   7320
            TabIndex        =   31
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblStatut_Codeline_dbl 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblStatut_Codeline_dbl"
            Height          =   375
            Left            =   240
            TabIndex        =   30
            Top             =   2400
            Width           =   6495
         End
         Begin VB.Label libStatut_Codeline_dbl 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            Height          =   375
            Left            =   7320
            TabIndex        =   29
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label lblStatut_Remise 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblStatut_Remise"
            Height          =   375
            Left            =   240
            TabIndex        =   28
            Top             =   720
            Width           =   6495
         End
         Begin VB.Label lblStatut_Cheques 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "libStatut_Cheques"
            Height          =   375
            Left            =   240
            TabIndex        =   27
            Top             =   1320
            Width           =   6495
         End
         Begin VB.Label libStatut_Remise 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            Height          =   375
            Left            =   7320
            TabIndex        =   26
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label libStatut_Cheques 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            Height          =   375
            Left            =   7320
            TabIndex        =   25
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label lblStatut_MyVision 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblStatut_MyVision"
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   3240
            Width           =   6495
         End
         Begin VB.Label libStatut_MyVision 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            Height          =   375
            Left            =   7320
            TabIndex        =   23
            Top             =   3240
            Width           =   1335
         End
      End
      Begin VB.Frame fraDétail 
         Height          =   8265
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   13290
         Begin VB.CheckBox chkRapprochement_Action 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Afficher uniquement les remises non traitées "
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Value           =   1  'Checked
            Width           =   3615
         End
         Begin VB.ListBox lstRapprochement_Action 
            BackColor       =   &H00C0FFC0&
            Height          =   450
            Left            =   4800
            TabIndex        =   13
            Top             =   240
            Width           =   4575
         End
         Begin VB.CommandButton cmdRapprochement_Update 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Validation Rapprochement"
            Height          =   795
            Left            =   10920
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid fgCHQRC1 
            Height          =   6645
            Left            =   120
            TabIndex        =   15
            Top             =   1200
            Width           =   6480
            _ExtentX        =   11430
            _ExtentY        =   11721
            _Version        =   393216
            Rows            =   1
            Cols            =   11
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   15269886
            ForeColor       =   8388608
            BackColorFixed  =   12648447
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   15269886
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"CHQ_SCAN.frx":070C
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
         Begin MSFlexGridLib.MSFlexGrid fgCHQMON 
            Height          =   6645
            Left            =   6840
            TabIndex        =   16
            Top             =   1200
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   11721
            _Version        =   393216
            Rows            =   1
            Cols            =   10
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   15268863
            ForeColor       =   8388608
            BackColorFixed  =   12640511
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14153215
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   "S |>Montant      | Dev|<Compte          |Date       |>Nb |>Rem  |||                                              "
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
      Begin MSFlexGridLib.MSFlexGrid fgSave 
         Height          =   2565
         Left            =   360
         TabIndex        =   10
         Top             =   2580
         Visible         =   0   'False
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   4524
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         RowHeightMin    =   300
         BackColor       =   14737632
         ForeColor       =   12582912
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   0
         GridLinesFixed  =   1
         FormatString    =   ">fichiers         C:\Zip                                                              <|Date dernière modif                  "
      End
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
      TabIndex        =   3
      Top             =   0
      Width           =   4905
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuDéontologie_ZIB 
         Caption         =   "Contrôle ZIB"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Sauvegarde (local + archive)"
      End
      Begin VB.Menu mnuArchivage 
         Caption         =   "Archivage => Serveur"
      End
      Begin VB.Menu mnuX1a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDéontologie 
         Caption         =   "Déontologie Stat Email"
      End
      Begin VB.Menu mnuDéontologie_Image 
         Caption         =   "Déontologie Image Email"
      End
      Begin VB.Menu mnuX3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport_Stat 
         Caption         =   "STAT.xlsx  (nb rem, nb chq, montant)"
      End
      Begin VB.Menu mnuX3b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export mdb => csv"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete CHEQUE /Ajust.mdb"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import csv => mdb"
      End
      Begin VB.Menu mnuUpdate_Table 
         Caption         =   "Update * Table"
      End
      Begin VB.Menu mnuUpdate_DEON 
         Caption         =   "STAT : Màj du dernier N° traité"
      End
      Begin VB.Menu mnuX5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRapprochement_Automatique 
         Caption         =   "Rapprochement : automatique"
      End
      Begin VB.Menu mnuRapprochement_Manuel 
         Caption         =   "Rapprochement : Manuel"
      End
      Begin VB.Menu mnuRapprochement_Display 
         Caption         =   "Rapprochement : affichage des remise du ........"
      End
      Begin VB.Menu mnuX6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRapprochement_SAB 
         Caption         =   "Rapprochement : extraction SAB"
      End
      Begin VB.Menu mnuRapprochement_CHQ_SCAN 
         Caption         =   "Rapprochement : extraction CHQ_SCAN"
      End
      Begin VB.Menu mnuRapprochement_Semi_Automatique 
         Caption         =   "Rapprochement : semi-automatique"
      End
      Begin VB.Menu mnuX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
      Begin VB.Menu mnuX2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Print_Liste 
         Caption         =   "Imprimer listes BIA & SG"
      End
      Begin VB.Menu mnuSelect_Print_Liste_BIA 
         Caption         =   "Imprimer liste BIA"
      End
      Begin VB.Menu mnuSelect_Print_Liste_SG 
         Caption         =   "Imprimer liste SG"
      End
      Begin VB.Menu mnuX4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCHQ_Stat 
         Caption         =   "Imprimer Statistiques"
      End
      Begin VB.Menu mnuPrint_ZINZIB 
         Caption         =   "Imprimer ZIN ZIB ="
      End
   End
   Begin VB.Menu mnuPrint2 
      Caption         =   "mnuPrint2"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint2_Liste 
         Caption         =   "Imrimer la liste des remises "
      End
   End
   Begin VB.Menu mnuPrint_imgCHQ 
      Caption         =   "mnuPrint_imgCHQ"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint_imgCHQ_All 
         Caption         =   "Imprimer toutes les images de la remise"
      End
   End
End
Attribute VB_Name = "frmCHQ_SCAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit

Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim x As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim CHQ_SCAN_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency
Dim blnAuto As Boolean

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

Public rsSab_Local As New ADODB.Recordset
Dim cnAdo_CHQ_ARCHIVE As New ADODB.Connection
Dim cnAdo_CHQ_LOCAL As New ADODB.Connection
Dim blnTransaction As Boolean

'______________________________________________________________________

Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long
Dim xCHQ_SCAN As typeCHQ_SCAN, meCHQ_SCAN As typeCHQ_SCAN
Dim oldCHQ_SCAN As typeCHQ_SCAN, newCHQ_SCAN As typeCHQ_SCAN
Dim arrCHQ_SCAN() As typeCHQ_SCAN, arrCHQ_SCAN_Nb As Long, arrCHQ_SCAN_Max As Long, arrCHQ_SCAN_Index As Long
Dim meYBIACPT0 As typeYBIACPT0
Dim xRemise As typeCHQ_SCAN
Dim autoDéontologie As typeCHQ_SCAN

Dim arrCHQ_SCAN_Détail() As typeCHQ_SCAN, arrCHQ_SCAN_Détail_Nb As Long, arrCHQ_SCAN_Détail_Max As Long, arrCHQ_SCAN_Détail_Index As Long
Dim curCHQ_SCAN_Détail As Currency
'______________________________________________________________________

Dim wFile As String, intFile As Integer
Dim blnTop1 As Boolean, blnTop2 As Boolean
Dim xIn As String
Dim wDreamFile_Date As String
Dim arrCHQ_Stat() As typeCHQ_Stat, arrCHQ_Stat_Nb As Integer, arrCHQ_Stat_Max As Integer
Dim meCV1 As typeCV, meCV2 As typeCV
'______________________________________________________________________

 
Dim xYCHQMON0 As typeYCHQMON0, oldYCHQMON0 As typeYCHQMON0, newYCHQMON0 As typeYCHQMON0
Dim arrYCHQMON0() As typeYCHQMON0, arrYCHQMON0_Nb As Integer, arrYCHQMON0_Max As Integer
Dim arrYCHQMON0_Link() As Long
Dim mCHQRC1_Index As Long, mCHQMON_Index As Long
Dim xZGUIRC10 As typeZGUIRC10
Dim fgCHQ_ForeColor As Long
Dim fgCHQRC1_FormatString As String, fgCHQRC1_K As Integer
Dim fgCHQRC1_RowDisplay As Integer, fgCHQRC1_RowClick As Integer, fgCHQRC1_ColClick As Integer
Dim fgCHQRC1_ColorClick As Long, fgCHQRC1_ColorDisplay As Long
Dim fgCHQRC1_Sort1 As Integer, fgCHQRC1_Sort2 As Integer
Dim fgCHQRC1_SortAD As Integer, fgCHQRC1_Sort1_Old As Integer
Dim fgCHQRC1_arrIndex As Integer
Dim blnfgCHQRC1_DisplayLine As Boolean

Dim fgCHQMON_FormatString As String, fgCHQMON_K As Integer
Dim fgCHQMON_RowDisplay As Integer, fgCHQMON_RowClick As Integer, fgCHQMON_ColClick As Integer
Dim fgCHQMON_ColorClick As Long, fgCHQMON_ColorDisplay As Long
Dim fgCHQMON_Sort1 As Integer, fgCHQMON_Sort2 As Integer
Dim fgCHQMON_SortAD As Integer, fgCHQMON_Sort1_Old As Integer
Dim fgCHQMON_arrIndex As Integer
Dim blnfgCHQMON_DisplayLine As Boolean

Dim blnRapprochement_Action As Boolean
Dim curRemise As Currency, curTotal As Currency

Dim xPath_ImgCHQ As String, xPath_ImgCHQ_Verso As String, blnImgCHQ_Recto As Boolean, blnImgCHQ_Verso As Boolean
Dim xYCHQDEON0 As typeYCHQDEON0
Dim arrYCHQDEON0() As typeYCHQDEON0, arrYCHQDEON0_Nb As Long, arrYCHQDEON0_Max As Long
Dim cmdPrint_imgCHQ_Text As String
Dim blnDéontologie_ZIB As Boolean

Dim mPériode_K As String

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
        For I = fgDossier_arrIndex To 0 Step -1
          fgDossier.Col = I: fgDossier.CellBackColor = lColor
        Next I
        'fgDossier.Col = 0
    End If
End If

End Sub

Private Sub fgDossier_Display()
Dim K As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgDossier_Reset
fgDossier.Visible = True
cmdPrint.Enabled = False
fraCHQ.Visible = False

fgDossier.Rows = 1
fgDossier.FormatString = fgDossier_FormatString
currentAction = "fgdossier_Display"
 
curRemise = CCur(xRemise.Zone1) / 100
curTotal = 0
For K = 1 To arrCHQ_SCAN_Détail_Nb
    xCHQ_SCAN = arrCHQ_SCAN_Détail(K)
        fgDossier.Rows = fgDossier.Rows + 1
        fgDossier.Row = fgDossier.Rows - 1
        fgDossier_DisplayLine K
Next K

Call lstErr_AddItem(lstErr, cmdContext, "Nb chèques : " & arrCHQ_SCAN_Détail_Nb): DoEvents
Call lstErr_AddItem(lstErr, cmdContext, "Montant Remise: " & curRemise): DoEvents
'If fgDossier.Rows > 1 Then
'    fgDossier_Sort1 = 0: fgDossier_Sort2 = 1: fgDossier_Sort
    cmdPrint.Enabled = True
'End If
If curRemise <> curTotal Then
    MsgBox "Montant de la remise :" & Format$(curRemise, "### ### ### ###.00") & vbCrLf & "Somme des chèques : &" & Format$(curTotal, "### ### ### ###.00"), vbCritical, "CHQ-SCAN : Dossier"
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgDossier_DisplayLine(lIndex As Long)
Dim curX As Currency
On Error Resume Next
fgDossier.Col = 0: fgDossier.Text = xCHQ_SCAN.NumLot
fgDossier.Col = 1: fgDossier.Text = xCHQ_SCAN.DateHourScan
curX = CCur(xCHQ_SCAN.Zone1) / 100
curTotal = curTotal + curX
fgDossier.Col = 2: fgDossier.Text = Format$(curX, "### ### ### ###.00")
fgDossier.Col = 3: fgDossier.Text = xCHQ_SCAN.Adresse0
fgDossier.Col = fgDossier_arrIndex: fgDossier.Text = lIndex

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

Public Sub fgCHQRC1_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgCHQRC1.Row

If lRow > 0 And lRow < fgCHQRC1.Rows Then
    fgCHQRC1.Row = lRow
    For I = 0 To fgCHQRC1_arrIndex
        fgCHQRC1.Col = I: fgCHQRC1.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgCHQRC1.Row = mRow
    If fgCHQRC1.Row > 0 Then
        lRow = fgCHQRC1.Row
        lColor_Old = fgCHQRC1.CellBackColor
        For I = 0 To fgCHQRC1_arrIndex
          fgCHQRC1.Col = I: fgCHQRC1.CellBackColor = lColor
        Next I
        fgCHQRC1.Col = 0
    End If
End If

End Sub


Private Sub fgCHQRC1_Display()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 2
lstRapprochement_Action.Visible = False
cmdRapprochement_Update.Visible = False
fgCHQRC1.Visible = False
fgCHQMON.Visible = False
mCHQRC1_Index = 0
mCHQMON_Index = 0

fgCHQRC1_Reset
fgCHQRC1.Rows = 1
fgCHQRC1.FormatString = fgCHQRC1_FormatString

fgCHQMON_Reset
fgCHQMON.Rows = 1
fgCHQMON.FormatString = fgCHQMON_FormatString
currentAction = "fgCHQRC1_Display"
    
For I = 1 To arrYCHQMON0_Nb
    
    xYCHQMON0 = arrYCHQMON0(I)
    If chkRapprochement_Action = "0" Then
        blnOk = True
    Else
        If xYCHQMON0.CHQMONSTA = " " Then
            blnOk = True
        Else
            blnOk = False
        End If
        
    End If
    
    If blnOk Then
        If xYCHQMON0.CHQMONSTA = " " Then
            fgCHQ_ForeColor = vbBlue
        Else
            fgCHQ_ForeColor = vbRed
        End If
        If xYCHQMON0.CHQRC1ETA = 0 Then
            fgCHQMON_DisplayLine (I)
        Else
            fgCHQRC1_DisplayLine (I)
        End If
    End If
    
    If arrYCHQMON0_Link(I) <> 0 Then cmdRapprochement_Update.Visible = True

Next I

Call lstErr_Clear(lstErr, cmdContext, "Remises : " & arrYCHQMON0_Nb): DoEvents
If fgCHQRC1.Rows > 1 Then
    fgCHQRC1_Sort1 = 5: fgCHQRC1_Sort2 = 5: fgCHQRC1_SortX 5
    cmdPrint.Enabled = True
End If
If fgCHQMON.Rows > 1 Then
    fgCHQMON_Sort1 = 1: fgCHQMON_Sort2 = 1: fgCHQMON_SortX 1
    cmdPrint.Enabled = True
End If
fgCHQRC1.Visible = True
fgCHQMON.Visible = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub fgCHQRC1_DisplayLine(lIndex As Long)
On Error Resume Next
fgCHQRC1.Rows = fgCHQRC1.Rows + 1
fgCHQRC1.Row = fgCHQRC1.Rows - 1
fgCHQRC1.Col = 0: fgCHQRC1.Text = xYCHQMON0.CHQMONSTA
fgCHQRC1.Col = 1: fgCHQRC1.Text = xYCHQMON0.CHQCREM
fgCHQRC1.CellForeColor = vbMagenta
fgCHQRC1.Col = 2: fgCHQRC1.Text = Val(xYCHQMON0.CHQRC1DOS)
fgCHQRC1.Col = 3: fgCHQRC1.Text = dateIBM10(xYCHQMON0.CHQRC1DCR, True)
fgCHQRC1.Col = fgCHQRC1_arrIndex: fgCHQRC1.Text = lIndex


    fgCHQRC1.Col = 4: fgCHQRC1.Text = xYCHQMON0.CHQCOMPTE
    fgCHQRC1.Col = 6: fgCHQRC1.Text = xYCHQMON0.CHQDEVISE
    fgCHQRC1.CellForeColor = fgCHQ_ForeColor
    fgCHQRC1.Col = 5: fgCHQRC1.Text = Format(xYCHQMON0.CHQMONTANT, "### ### ### ###.00")
    fgCHQRC1.CellForeColor = fgCHQ_ForeColor


End Sub

Public Sub fgCHQRC1_Reset()
fgCHQRC1.Clear
fgCHQRC1_Sort1 = 0: fgCHQRC1_Sort2 = 0
fgCHQRC1_Sort1_Old = -1
fgCHQRC1_RowDisplay = 0: fgCHQRC1_RowClick = 0
fgCHQRC1_arrIndex = fgCHQRC1.Cols - 1
blnfgCHQRC1_DisplayLine = False
End Sub



Public Sub fgCHQRC1_Sort()
If fgCHQRC1.Rows > 1 Then
    fgCHQRC1.Row = 1
    fgCHQRC1.RowSel = fgCHQRC1.Rows - 1
    
    If fgCHQRC1_Sort1_Old = fgCHQRC1_Sort1 Then
        If fgCHQRC1_SortAD = 5 Then
            fgCHQRC1_SortAD = 6
        Else
            fgCHQRC1_SortAD = 5
        End If
    Else
        fgCHQRC1_SortAD = 5
    End If
    fgCHQRC1_Sort1_Old = fgCHQRC1_Sort1
    
    fgCHQRC1.Col = fgCHQRC1_Sort1
    fgCHQRC1.ColSel = fgCHQRC1_Sort2
    fgCHQRC1.Sort = fgCHQRC1_SortAD
End If
'cboDevise_Reset
End Sub


Public Sub fgCHQMON_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
On Error Resume Next
mRow = fgCHQMON.Row

If lRow > 0 And lRow < fgCHQMON.Rows Then
    fgCHQMON.Row = lRow
    For I = 0 To fgCHQMON_arrIndex
        fgCHQMON.Col = I: fgCHQMON.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgCHQMON.Row = mRow
    If fgCHQMON.Row > 0 Then
        lRow = fgCHQMON.Row
        lColor_Old = fgCHQMON.CellBackColor
        For I = 0 To fgCHQMON_arrIndex
          fgCHQMON.Col = I: fgCHQMON.CellBackColor = lColor
        Next I
        fgCHQMON.Col = 0
    End If
End If

End Sub

Public Sub fgCHQMON_DisplayLine(lIndex As Long)
On Error Resume Next
fgCHQMON.Rows = fgCHQMON.Rows + 1
fgCHQMON.Row = fgCHQMON.Rows - 1
fgCHQMON.Col = 0: fgCHQMON.Text = xYCHQMON0.CHQMONSTA
fgCHQMON.Col = 3: fgCHQMON.Text = xYCHQMON0.CHQCOMPTE
fgCHQMON.Col = 1: fgCHQMON.Text = Format(xYCHQMON0.CHQMONTANT, "### ### ### ###.00")
fgCHQMON.CellForeColor = fgCHQ_ForeColor
fgCHQMON.Col = 2: fgCHQMON.Text = xYCHQMON0.CHQDEVISE
fgCHQMON.CellForeColor = fgCHQ_ForeColor
fgCHQMON.Col = 4: fgCHQMON.Text = dateImp10(xYCHQMON0.CHQDATE)
fgCHQMON.Col = 5: fgCHQMON.Text = xYCHQMON0.CHQNB
fgCHQMON.CellForeColor = vbMagenta 'fgCHQ_ForeColor
fgCHQMON.Col = 6: fgCHQMON.Text = Val(xYCHQMON0.CHQCREM)

fgCHQMON.Col = fgCHQMON_arrIndex: fgCHQMON.Text = lIndex

End Sub
Public Sub fgCHQMON_Reset()
fgCHQMON.Clear
fgCHQMON_Sort1 = 0: fgCHQMON_Sort2 = 0
fgCHQMON_Sort1_Old = -1
fgCHQMON_RowDisplay = 0: fgCHQMON_RowClick = 0
fgCHQMON_arrIndex = fgCHQMON.Cols - 1
blnfgCHQMON_DisplayLine = False
End Sub





Public Sub fgCHQMON_Sort()
If fgCHQMON.Rows > 1 Then
    fgCHQMON.Row = 1
    fgCHQMON.RowSel = fgCHQMON.Rows - 1
    
    If fgCHQMON_Sort1_Old = fgCHQMON_Sort1 Then
        If fgCHQMON_SortAD = 5 Then
            fgCHQMON_SortAD = 6
        Else
            fgCHQMON_SortAD = 5
        End If
    Else
        fgCHQMON_SortAD = 5
    End If
    fgCHQMON_Sort1_Old = fgCHQMON_Sort1
    
    fgCHQMON.Col = fgCHQMON_Sort1
    fgCHQMON.ColSel = fgCHQMON_Sort2
    fgCHQMON.Sort = fgCHQMON_SortAD
End If
'cboDevise_Reset
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
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgselect_Display"
    
For I = 1 To arrCHQ_SCAN_Nb
         
    xCHQ_SCAN = arrCHQ_SCAN(I)
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
Next I

fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrCHQ_SCAN_Nb): DoEvents
If fgSelect.Rows > 1 Then
    If Trim(txtSelect_COMPTE) = "" Then
        fgSelect_Sort1 = 0: fgSelect_Sort2 = 3: fgSelect_Sort
    Else
        fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
    End If
    
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub arrCHQ_SCAN_Remise_sql(xWhere As String)
Dim V
Dim x As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrCHQ_SCAN(101)
arrCHQ_SCAN_Max = 100: arrCHQ_SCAN_Nb = 0

Set rsSab = Nothing

xSQL = "select * from CHEQUE " & xWhere
If optSelect_Archive Then
    Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
Else
    Set rsSab = cnAdo_CHQ_LOCAL.Execute(xSQL)
End If

Do While Not rsSab.EOF
    V = srvCHQ_SCAN_GetBuffer_ODBC(rsSab, xCHQ_SCAN)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmCHQ_SCAN.fgselect_Display"
        '' Exit Sub
     Else
         arrCHQ_SCAN_Nb = arrCHQ_SCAN_Nb + 1
         If arrCHQ_SCAN_Nb > arrCHQ_SCAN_Max Then
             arrCHQ_SCAN_Max = arrCHQ_SCAN_Max + 100
             ReDim Preserve arrCHQ_SCAN(arrCHQ_SCAN_Max)
         End If
         
         arrCHQ_SCAN(arrCHQ_SCAN_Nb) = xCHQ_SCAN
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
Private Sub arrCHQ_SCAN_Détail_Sql(lDate As String, lCRem As String)
Dim V
Dim x As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrCHQ_SCAN_Détail(101)
arrCHQ_SCAN_Détail_Max = 100: arrCHQ_SCAN_Détail_Nb = 0
curCHQ_SCAN_Détail = 0
Set rsSab = Nothing

xSQL = "select * from CHEQUE Where Date = '" & lDate & "' and CRem = '" & lCRem & "' and ID = 'C' order by IMAGE"
If optSelect_Archive Then
    Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
Else
    Set rsSab = cnAdo_CHQ_LOCAL.Execute(xSQL)
End If

Do While Not rsSab.EOF
    V = srvCHQ_SCAN_GetBuffer_ODBC(rsSab, xCHQ_SCAN)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmCHQ_SCAN.fgselect_Display"
        '' Exit Sub
     Else
         arrCHQ_SCAN_Détail_Nb = arrCHQ_SCAN_Détail_Nb + 1
         If arrCHQ_SCAN_Détail_Nb > arrCHQ_SCAN_Détail_Max Then
             arrCHQ_SCAN_Détail_Max = arrCHQ_SCAN_Détail_Max + 100
             ReDim Preserve arrCHQ_SCAN_Détail(arrCHQ_SCAN_Détail_Max)
         End If
         
         arrCHQ_SCAN_Détail(arrCHQ_SCAN_Détail_Nb) = xCHQ_SCAN
         curCHQ_SCAN_Détail = curCHQ_SCAN_Détail + CCur(Val(xCHQ_SCAN.Zone1)) / 100
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


Public Sub cmdExport_Stat()
On Error GoTo Error_Handler
Dim wFile As String, wFilex As String, wFile2 As String, xSQL As String
Dim x As String, K As Long, kMax As Long, K2 As Long, K3 As Long
Dim xWhere As String, X2 As String
Dim arrCREM(20000) As String, arrCREM_K(20000) As Long, arrCREM_Nb As Long
Dim xNature As String, blnOk As Boolean, xCREM As String
Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim wsRow As Long
Dim arrCHQ_SCAN_Stat_Nb As Integer
Dim arrCHQ_SCAN_Stat() As typeCHQ_SCAN_Stat

wFile = "C:\Temp\CHQ_SCAN " & DSys & " " & time_Hms & ".xlsx"
'______________________________________________

x = InputBox("par défaut : " & wFile _
    & vbCrLf & "     ==================================================" _
    & vbCrLf, "CHQ_SCAN statistiques : nom du fichier d'exportation", wFile)
If Trim(x) = "" Then Exit Sub

wFilex = Trim(x)
'______________________________________________


If Dir(wFilex) <> "" Then Kill wFilex

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
'With wbExcel
'    .Title = "CHQ_SCAN statistiques"
'    .Subject = "CHQ_SCAN statistiques"
'End With

Set wsExcel = wbExcel.ActiveSheet
wsExcel.Name = "CHQ_SCAN statistiques"

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

wsRow = 1
Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

wsExcel.Cells(wsRow, 1) = "Devise": wsExcel.Columns(1).ColumnWidth = 6
wsExcel.Cells(wsRow, 2) = "Nature": wsExcel.Columns(2).ColumnWidth = 8
wsExcel.Cells(wsRow, 3) = "Date": wsExcel.Columns(2).ColumnWidth = 12: wsExcel.Columns(3).NumberFormat = "mm/dd/yyyy"
    wsExcel.Columns(3).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
wsExcel.Cells(wsRow, 4) = "Nb remises"
    wsExcel.Columns(4).ColumnWidth = 12: wsExcel.Columns(4).NumberFormat = "#######"
    wsExcel.Columns(4).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Cells(wsRow, 5) = "Nb chèques"
    wsExcel.Columns(5).ColumnWidth = 12: wsExcel.Columns(5).NumberFormat = "#######"
    wsExcel.Columns(5).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Cells(wsRow, 6) = "Montant"
    wsExcel.Columns(6).ColumnWidth = 15: wsExcel.Columns(6).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
    wsExcel.Columns(6).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight


For K = 1 To 6
    wsExcel.Columns(K).VerticalAlignment = Excel.XlVAlign.xlVAlignTop
    wsExcel.Cells(1, K).Interior.Color = RGB(255, 168, 125)
Next K

srvCHQ_SCAN_Init oldCHQ_SCAN
arrCREM_Nb = 0
arrCHQ_SCAN_Stat_Nb = 0
ReDim arrCHQ_SCAN_Stat(2000) As typeCHQ_SCAN_Stat

'_____________________________________________________________________________________________
x = "select * from CHEQUE  where ID = 'R' and Date >= '" & wAMJMin & "' and Date <= '" & WAMJMax & "' and StatutRem = 'AJ' order by Devise  , Nature , Date"
If optSelect_Archive Then
    Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(x)
Else
    Set rsSab = cnAdo_CHQ_LOCAL.Execute(x)
End If

Do While Not rsSab.EOF
    V = srvCHQ_SCAN_GetBuffer_ODBC(rsSab, xCHQ_SCAN)
    arrCREM_Nb = arrCREM_Nb + 1
    arrCREM(arrCREM_Nb) = xCHQ_SCAN.CRem
    
    If xCHQ_SCAN.StatutRem = "AJ" Then
        xNature = xCHQ_SCAN.Nature
    Else
        xNature = "? " & xCHQ_SCAN.Nature
    End If
    blnOk = False
    For K2 = 1 To arrCHQ_SCAN_Stat_Nb
        If arrCHQ_SCAN_Stat(K2).Devise = xCHQ_SCAN.Devise _
        And arrCHQ_SCAN_Stat(K2).Date = xCHQ_SCAN.Date _
        And arrCHQ_SCAN_Stat(K2).Nature = xNature Then
            blnOk = True
            arrCREM_K(arrCREM_Nb) = K2
           Exit For
        End If
    Next K2
    
    If Not blnOk Then
        arrCHQ_SCAN_Stat_Nb = arrCHQ_SCAN_Stat_Nb + 1
        arrCREM_K(arrCREM_Nb) = arrCHQ_SCAN_Stat_Nb
        arrCHQ_SCAN_Stat(arrCHQ_SCAN_Stat_Nb).Devise = xCHQ_SCAN.Devise
        arrCHQ_SCAN_Stat(arrCHQ_SCAN_Stat_Nb).Date = xCHQ_SCAN.Date
        arrCHQ_SCAN_Stat(arrCHQ_SCAN_Stat_Nb).Nature = xNature
        arrCHQ_SCAN_Stat(arrCHQ_SCAN_Stat_Nb).REM_nb = 0
        arrCHQ_SCAN_Stat(arrCHQ_SCAN_Stat_Nb).CHQ_nb = 0
        arrCHQ_SCAN_Stat(arrCHQ_SCAN_Stat_Nb).CHQ_mt = 0
        
    End If
    rsSab.MoveNext
Loop

'_________________________________________________________________________________________

x = "select CREM , IMAGE , Zone1 from CHEQUE  where  ID = 'C' and Date >= '" & wAMJMin & "' and Date <= '" & WAMJMax & "' order by  CREM , IMAGE"
If optSelect_Archive Then
    Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(x)
Else
    Set rsSab = cnAdo_CHQ_LOCAL.Execute(x)
End If
K = 0
Do While Not rsSab.EOF
    'V = srvCHQ_SCAN_GetBuffer_ODBC(rsSab, xCHQ_SCAN)
    xCREM = rsSab("CREM")
    If xCREM <> arrCREM(K) Then
        For K = 1 To arrCREM_Nb
            If xCREM = arrCREM(K) Then
                K2 = arrCREM_K(K)
               arrCHQ_SCAN_Stat(K2).REM_nb = arrCHQ_SCAN_Stat(K2).REM_nb + 1
               Exit For
            End If
        Next K
    End If
    x = rsSab("Zone1")
    If IsNumeric(x) Then
        arrCHQ_SCAN_Stat(K2).CHQ_nb = arrCHQ_SCAN_Stat(K2).CHQ_nb + 1
        arrCHQ_SCAN_Stat(K2).CHQ_mt = arrCHQ_SCAN_Stat(K2).CHQ_mt + CCur(x) / 100
    End If

    
    
    rsSab.MoveNext
Loop
arrCHQ_SCAN_Stat(0) = arrCHQ_SCAN_Stat(1)
arrCHQ_SCAN_Stat(0).REM_nb = 0
arrCHQ_SCAN_Stat(0).CHQ_nb = 0
arrCHQ_SCAN_Stat(0).CHQ_mt = 0

For K2 = 1 To arrCHQ_SCAN_Stat_Nb
        
        If arrCHQ_SCAN_Stat(K2).Devise <> arrCHQ_SCAN_Stat(0).Devise _
        Or arrCHQ_SCAN_Stat(K2).Nature <> arrCHQ_SCAN_Stat(0).Nature Then
            wsRow = wsRow + 1
            For K = 1 To 6
                wsExcel.Cells(wsRow, K).Interior.Color = RGB(255, 255, 153)
            Next K
            wsExcel.Cells(wsRow, 1) = arrCHQ_SCAN_Stat(0).Devise
            wsExcel.Cells(wsRow, 3) = "Total"
            wsExcel.Cells(wsRow, 2) = arrCHQ_SCAN_Stat(0).Nature
            wsExcel.Cells(wsRow, 4) = arrCHQ_SCAN_Stat(0).REM_nb
            wsExcel.Cells(wsRow, 5) = arrCHQ_SCAN_Stat(0).CHQ_nb
            wsExcel.Cells(wsRow, 6) = arrCHQ_SCAN_Stat(0).CHQ_mt
            arrCHQ_SCAN_Stat(0) = arrCHQ_SCAN_Stat(K2)
            arrCHQ_SCAN_Stat(0).REM_nb = 0
            arrCHQ_SCAN_Stat(0).CHQ_nb = 0
            arrCHQ_SCAN_Stat(0).CHQ_mt = 0
        End If
        
        wsRow = wsRow + 1
        wsExcel.Cells(wsRow, 1) = arrCHQ_SCAN_Stat(K2).Devise
        wsExcel.Cells(wsRow, 3) = dateImp10_S(arrCHQ_SCAN_Stat(K2).Date)
        wsExcel.Cells(wsRow, 2) = arrCHQ_SCAN_Stat(K2).Nature
        wsExcel.Cells(wsRow, 4) = arrCHQ_SCAN_Stat(K2).REM_nb
        wsExcel.Cells(wsRow, 5) = arrCHQ_SCAN_Stat(K2).CHQ_nb
        wsExcel.Cells(wsRow, 6) = arrCHQ_SCAN_Stat(K2).CHQ_mt
        arrCHQ_SCAN_Stat(0).REM_nb = arrCHQ_SCAN_Stat(0).REM_nb + arrCHQ_SCAN_Stat(K2).REM_nb
        arrCHQ_SCAN_Stat(0).CHQ_nb = arrCHQ_SCAN_Stat(0).CHQ_nb + arrCHQ_SCAN_Stat(K2).CHQ_nb
        arrCHQ_SCAN_Stat(0).CHQ_mt = arrCHQ_SCAN_Stat(0).CHQ_mt + arrCHQ_SCAN_Stat(K2).CHQ_mt
Next K2

wsRow = wsRow + 1

For K = 1 To 6
    wsExcel.Cells(wsRow, K).Interior.Color = RGB(255, 255, 153)
Next K
wsExcel.Cells(wsRow, 1) = arrCHQ_SCAN_Stat(0).Devise
wsExcel.Cells(wsRow, 3) = "Total"
wsExcel.Cells(wsRow, 2) = arrCHQ_SCAN_Stat(0).Nature
wsExcel.Cells(wsRow, 4) = arrCHQ_SCAN_Stat(0).REM_nb
wsExcel.Cells(wsRow, 5) = arrCHQ_SCAN_Stat(0).CHQ_nb
wsExcel.Cells(wsRow, 6) = arrCHQ_SCAN_Stat(0).CHQ_mt


Call lstErr_ChangeLastItem(lstErr, cmdContext, "Exportation terminée : " & wsRow & " enregistrements"): DoEvents
Set rsSab = Nothing


wbExcel.SaveAs wFilex

wbExcel.Close

'____________________________________________________________________________________
appExcel.Quit

Set rsSab = Nothing

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing

Call lstErr_AddItem(lstErr, cmdContext, "Exportation terminée"): DoEvents

'_____________________________
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, Me.Name

End Sub


Private Sub arrYCHQDEON0_Sql()
Dim V
Dim x As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrYCHQDEON0(101)
arrYCHQDEON0_Max = 100: arrYCHQDEON0_Nb = 0
Set rsSab = Nothing

xSQL = "select *  from " & paramIBM_Library_SABSPE & ".YCHQDEON0 "
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYCHQDEON0_GetBuffer_ODBC(rsSab, xYCHQDEON0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmCHQ_SCAN.arrYCHQDEON0_Sql"
        '' Exit Sub
     Else
         arrYCHQDEON0_Nb = arrYCHQDEON0_Nb + 1
         If arrYCHQDEON0_Nb > arrYCHQDEON0_Max Then
             arrYCHQDEON0_Max = arrYCHQDEON0_Max + 100
             ReDim Preserve arrYCHQDEON0(arrYCHQDEON0_Max)
         End If
         
         arrYCHQDEON0(arrYCHQDEON0_Nb) = xYCHQDEON0
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

Private Sub arrYCHQMON0_Sql(lWhere As String)
Dim V
Dim I As Long
Dim x As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrYCHQMON0(101)
arrYCHQMON0_Max = 100: arrYCHQMON0_Nb = 0
Set rsSab = Nothing

xSQL = "select *  from " & paramIBM_Library_SABSPE & ".YCHQMON0 " & lWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYCHQMON0_GetBuffer(rsSab, xYCHQMON0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYCHQMON0.fgselect_Display"
        '' Exit Sub
     Else
         arrYCHQMON0_Nb = arrYCHQMON0_Nb + 1
         If arrYCHQMON0_Nb > arrYCHQMON0_Max Then
             arrYCHQMON0_Max = arrYCHQMON0_Max + 100
             ReDim Preserve arrYCHQMON0(arrYCHQMON0_Max)
         End If
         
         arrYCHQMON0(arrYCHQMON0_Nb) = xYCHQMON0
    End If
    rsSab.MoveNext

Loop

ReDim arrYCHQMON0_Link(arrYCHQMON0_Max)
For I = 1 To arrYCHQMON0_Nb
    arrYCHQMON0_Link(I) = 0
    xZGUIRC10.GUIRC1ETA = arrYCHQMON0(I).CHQRC1ETA
    xZGUIRC10.GUIRC1AGE = arrYCHQMON0(I).CHQRC1AGE
    xZGUIRC10.GUIRC1SER = arrYCHQMON0(I).CHQRC1SER
    xZGUIRC10.GUIRC1SSE = arrYCHQMON0(I).CHQRC1SSE
    xZGUIRC10.GUIRC1OPE = arrYCHQMON0(I).CHQRC1OPE
    xZGUIRC10.GUIRC1DOS = arrYCHQMON0(I).CHQRC1DOS

    If IsNull(rsZGUIRC10_Read(xZGUIRC10)) Then
    ' alimentation tableau pour tri montant / Compte / Date
    '==============================
        arrYCHQMON0(I).CHQCREM = xZGUIRC10.GUIRC1NAT
        arrYCHQMON0(I).CHQDEVISE = xZGUIRC10.GUIRC1DE1
        arrYCHQMON0(I).CHQMONTANT = xZGUIRC10.GUIRC1MO2
        arrYCHQMON0(I).CHQCOMPTE = xZGUIRC10.GUIRC1CP2
    '==========================================================
    
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
Dim x As String, lenX As Integer
Dim xSQL As String
On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = dateImp10(xCHQ_SCAN.Date)
fgSelect.Col = 1: fgSelect.Text = xCHQ_SCAN.RefInterne
fgSelect.Col = 2: fgSelect.Text = xCHQ_SCAN.CRem
fgSelect.Col = 3: fgSelect.Text = xCHQ_SCAN.StatutRem
curX = CCur(xCHQ_SCAN.Zone1) / 100
fgSelect.Col = 4: fgSelect.Text = Format$(curX, "### ### ### ###.00")
If xCHQ_SCAN.StatutRem = "AJ" Then
    fgSelect.CellForeColor = vbBlue
Else
    fgSelect.CellForeColor = vbRed
End If

fgSelect.Col = 5: fgSelect.Text = xCHQ_SCAN.Devise
fgSelect.Col = 6: fgSelect.Text = xCHQ_SCAN.COMPTE

If xCHQ_SCAN.Id = "R" Then
    x = sqlYBIACPT0_COMPTEINT(xCHQ_SCAN.COMPTE)
Else
    x = xCHQ_SCAN.Cmc7
End If

fgSelect.Col = 7: fgSelect.Text = x
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

End Sub


Public Sub fgSelect_Reset()
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
Dim I As Integer, x As String
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    If lK = 2 Then
        fgSelect.Col = 2
        x = fgSelect.Text
    Else
        x = ""
    End If
    
    fgSelect.Col = 3
    x = x & Format$(Val(fgSelect.Text), "000000000000000.00")
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = x
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub

Public Sub fgCHQMON_SortX(lK As Integer)
Dim I As Integer, x As String, K As Long
For I = 1 To fgCHQMON.Rows - 1
    fgCHQMON.Row = I
    fgCHQMON.Col = fgCHQMON_arrIndex:  K = CLng(fgCHQMON.Text)
    Select Case lK
        Case 1: x = Format$(arrYCHQMON0(K).CHQMONTANT, "000000000000000.00") & arrYCHQMON0(K).CHQDATE & arrYCHQMON0(K).CHQCOMPTE
        Case 2: x = arrYCHQMON0(K).CHQDEVISE & Format$(arrYCHQMON0(K).CHQMONTANT, "000000000000000.00")
        Case 3: x = arrYCHQMON0(K).CHQCOMPTE & arrYCHQMON0(K).CHQDATE & Format$(arrYCHQMON0(K).CHQMONTANT, "000000000000000.00")
        Case 4: x = arrYCHQMON0(K).CHQDATE & Format$(arrYCHQMON0(K).CHQMONTANT, "000000000000000.00")
        Case 5: x = Format$(arrYCHQMON0(K).CHQNB, "000000000") & Format$(arrYCHQMON0(K).CHQMONTANT, "000000000000000.00")
        Case 6: x = arrYCHQMON0(K).CHQCREM
    End Select
   fgCHQMON.Col = fgCHQMON_arrIndex - 1
    fgCHQMON.Text = x
Next I


fgCHQMON_Sort1 = fgCHQMON_arrIndex - 1: fgCHQMON_Sort2 = fgCHQMON_arrIndex - 1
fgCHQMON_Sort
End Sub
Public Sub fgCHQRC1_SortX(lK As Integer)
Dim I As Integer, x As String, K As Integer
For I = 1 To fgCHQRC1.Rows - 1

    fgCHQRC1.Row = I
    fgCHQRC1.Col = fgCHQRC1_arrIndex:  K = CLng(fgCHQRC1.Text)
    Select Case lK
        Case 1: 'CHQCREM : naure de ZGUIRC10
                x = arrYCHQMON0(K).CHQCREM & Format$(arrYCHQMON0(K).CHQMONTANT, "000000000000000.00")
        Case 2: x = arrYCHQMON0(K).CHQRC1DOS & Format$(arrYCHQMON0(K).CHQMONTANT, "000000000000000.00")
        Case 3: x = arrYCHQMON0(K).CHQRC1DCR & Format$(arrYCHQMON0(K).CHQMONTANT, "000000000000000.00")
        Case 4: x = arrYCHQMON0(K).CHQCOMPTE & Format$(arrYCHQMON0(K).CHQMONTANT, "000000000000000.00")
        Case 5: x = Format$(arrYCHQMON0(K).CHQMONTANT, "000000000000000.00") & arrYCHQMON0(K).CHQDATE & arrYCHQMON0(K).CHQCOMPTE
        Case 6: x = arrYCHQMON0(K).CHQDEVISE & Format$(arrYCHQMON0(K).CHQMONTANT, "000000000000000.00")
    End Select
    fgCHQRC1.Col = fgCHQRC1_arrIndex - 1
    fgCHQRC1.Text = x

Next I


fgCHQRC1_Sort1 = fgCHQRC1_arrIndex - 1: fgCHQRC1_Sort2 = fgCHQRC1_arrIndex - 1
fgCHQRC1_Sort
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

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), CHQ_SCAN_Aut)

Form_Init

Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@CHQ_DEON": blnAuto = True
                      cmdDéontologie_Select
                       
                      Call DTPicker_Set(txtSelect_Date, YBIATAB0_DATE_CPT_J)
                      optSelect_Archive = True
                      'Surveillance BIOCORPS : compte annulé
                      'txtSelect_COMPTE = "50089"
                      'cmdDéontologie_Image_SendMail
                      
                      cnAdo_Close
                      Unload Me

    Case Else: blnAuto = False
End Select


End Sub


Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmCHQ_SCAN.paramSAA_Init"
    Unload Me
Else
    lstErr.Clear
End If


blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgCHQRC1_FormatString = fgCHQRC1.FormatString
fgCHQMON_FormatString = fgCHQMON.FormatString
fgDossier_FormatString = fgDossier.FormatString
fgSave.Enabled = False
fgSave.Visible = False
fgSelect.Enabled = True
cmdReset

Me.Enabled = True
Me.MousePointer = 0
End Sub

Private Sub cmdimgCHQ_Click()
If blnImgCHQ_Recto Then
    blnImgCHQ_Recto = False: imgCHQ.Visible = False: imgCHQ_Verso.Visible = True
Else
    blnImgCHQ_Recto = True: imgCHQ.Visible = True: imgCHQ_Verso.Visible = False
End If

End Sub

Private Sub cmdRapprochement_Update_Click()
Dim K As Integer
Dim V

Me.Enabled = False: Me.MousePointer = vbHourglass

SSTab1.Tab = 2
V = Null

For K = 1 To arrYCHQMON0_Nb
    newYCHQMON0 = arrYCHQMON0(K)
    oldYCHQMON0 = newYCHQMON0
    oldYCHQMON0.CHQMONSTA = " "    ' statut avant rapprochement
    
    Select Case newYCHQMON0.CHQMONSTA
        Case " ":
        Case "S": V = sqlYCHQMON0_Delete(oldYCHQMON0)
        Case "I": V = sqlYCHQMON0_Update(newYCHQMON0, oldYCHQMON0)
        Case "M", "="
                    If newYCHQMON0.CHQRC1ETA = 1 Then
                        xYCHQMON0 = arrYCHQMON0(arrYCHQMON0_Link(K))
                        oldYCHQMON0.CHQDATE = 0
                        oldYCHQMON0.CHQCOMPTE = ""
                        oldYCHQMON0.CHQCREM = ""
                        oldYCHQMON0.CHQDEVISE = ""
                        oldYCHQMON0.CHQMONTANT = 0
                        oldYCHQMON0.CHQNB = 0

                        newYCHQMON0.CHQDATE = xYCHQMON0.CHQDATE
                        newYCHQMON0.CHQCOMPTE = xYCHQMON0.CHQCOMPTE
                        newYCHQMON0.CHQCREM = xYCHQMON0.CHQCREM
                        newYCHQMON0.CHQDEVISE = xYCHQMON0.CHQDEVISE
                        newYCHQMON0.CHQMONTANT = xYCHQMON0.CHQMONTANT
                        newYCHQMON0.CHQNB = xYCHQMON0.CHQNB
                        V = sqlYCHQMON0_Update(newYCHQMON0, oldYCHQMON0)
                        If IsNull(V) Then V = sqlYCHQMON0_Delete(xYCHQMON0)
                    End If
    End Select
    If Not IsNull(V) Then GoTo Error_Handler
Next K
arrYCHQMON0_Nb = 0

GoTo Exit_sub

Error_Handler:
MsgBox "cmdRapprochement_Update " & Error, vbCritical, Me.Caption
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

'réaffichage après maj
mnuRapprochement_Manuel_Click

End Sub

Private Sub lstRapprochement_Action_Click()

If lstRapprochement_Action.Visible Then
    Select Case Mid$(lstRapprochement_Action.Text, 1, 1)
            Case 1: 'Rapprochement
            Case 2: arrYCHQMON0_Link(mCHQRC1_Index) = 0: arrYCHQMON0_Link(mCHQMON_Index) = 0
                    arrYCHQMON0(mCHQRC1_Index).CHQMONSTA = " ": arrYCHQMON0(mCHQMON_Index).CHQMONSTA = " "
                    fgCHQRC1_Display
            Case 3: arrYCHQMON0_Link(mCHQRC1_Index) = -mCHQRC1_Index: arrYCHQMON0(mCHQRC1_Index).CHQMONSTA = "I"
                    fgCHQRC1_Display
            Case 4: arrYCHQMON0_Link(mCHQRC1_Index) = -mCHQRC1_Index: arrYCHQMON0(mCHQRC1_Index).CHQMONSTA = "S"
                    fgCHQRC1_Display
            Case 5: arrYCHQMON0_Link(mCHQRC1_Index) = 0: arrYCHQMON0(mCHQRC1_Index).CHQMONSTA = " "
                    fgCHQRC1_Display
            Case 6: arrYCHQMON0_Link(mCHQMON_Index) = -mCHQMON_Index: arrYCHQMON0(mCHQMON_Index).CHQMONSTA = "I"
                    fgCHQRC1_Display
            Case 7: arrYCHQMON0_Link(mCHQMON_Index) = -mCHQMON_Index: arrYCHQMON0(mCHQMON_Index).CHQMONSTA = "S"
                    fgCHQRC1_Display
            Case 8: arrYCHQMON0_Link(mCHQMON_Index) = 0: arrYCHQMON0(mCHQMON_Index).CHQMONSTA = " "
                    fgCHQRC1_Display
            Case 9: lstRapprochement_Action.Visible = False
                    mCHQRC1_Index = 0
                    mCHQMON_Index = 0
    
    End Select
End If


End Sub


Private Sub chkRapprochement_Action_Click()
fgCHQRC1_Display
End Sub


Private Sub fgCHQMON_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgCHQMON.RowHeightMin Then
    Select Case fgCHQMON.Col
        Case 0: fgCHQMON_Sort1 = 0: fgCHQMON_Sort2 = 1: fgCHQMON_Sort
        Case 1:  fgCHQMON_Sort1 = 1: fgCHQMON_Sort2 = 1: fgCHQMON_SortX 1
        Case 2: fgCHQMON_Sort1 = 2: fgCHQMON_Sort2 = 2: fgCHQMON_SortX 2
        Case 3: fgCHQMON_Sort1 = 3: fgCHQMON_Sort2 = 3: fgCHQMON_SortX 3
        Case 4: fgCHQMON_Sort1 = 4: fgCHQMON_Sort2 = 4: fgCHQMON_SortX 4
        Case 5: fgCHQMON_Sort1 = 5: fgCHQMON_Sort2 = 5: fgCHQMON_SortX 5
        Case 6: fgCHQMON_Sort1 = 6: fgCHQMON_Sort2 = 6: fgCHQMON_SortX 6
    End Select
Else
    If fgCHQMON.Rows > 1 Then
        Call fgCHQMON_Color(fgCHQMON_RowClick, MouseMoveUsr.BackColor, fgCHQMON_ColorClick)
        fgCHQMON.Col = fgCHQMON_arrIndex:  mCHQMON_Index = CLng(fgCHQMON.Text)
        xYCHQMON0 = arrYCHQMON0(mCHQMON_Index)
        fgCHQMON.Col = 0: fgCHQMON.LeftCol = 0
        If blnRapprochement_Action Then
            If mCHQRC1_Index = 0 Then
                lstRapprochement_Action_Init
            Else
                cmdRapprochement_Validation
                lstRapprochement_Action.Visible = False
                    mCHQRC1_Index = 0
                    mCHQMON_Index = 0
                
            End If
        End If
   End If
End If

End Sub


Private Sub fgCHQRC1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgCHQRC1.RowHeightMin Then
    Select Case fgCHQRC1.Col
        Case 0: fgCHQRC1_Sort1 = 0: fgCHQRC1_Sort2 = 1: fgCHQRC1_Sort
        Case 1:  fgCHQRC1_Sort1 = 1: fgCHQRC1_Sort2 = 1: fgCHQRC1_SortX 1
        Case 2: fgCHQRC1_Sort1 = 2: fgCHQRC1_Sort2 = 2: fgCHQRC1_SortX 2
        Case 3: fgCHQRC1_Sort1 = 3: fgCHQRC1_Sort2 = 3: fgCHQRC1_SortX 3
        Case 4: fgCHQRC1_Sort1 = 4: fgCHQRC1_Sort2 = 4: fgCHQRC1_SortX 4
        Case 5: fgCHQRC1_Sort1 = 5: fgCHQRC1_Sort2 = 5: fgCHQRC1_SortX 6
        Case 6: fgCHQRC1_Sort1 = 6: fgCHQRC1_Sort2 = 6: fgCHQRC1_SortX 5
    End Select
Else
    If fgCHQRC1.Rows > 1 Then
        Call fgCHQRC1_Color(fgCHQRC1_RowClick, MouseMoveUsr.BackColor, fgCHQRC1_ColorClick)
        fgCHQRC1.Col = fgCHQRC1_arrIndex:  mCHQRC1_Index = CLng(fgCHQRC1.Text)
        oldYCHQMON0 = arrYCHQMON0(mCHQRC1_Index)
        DoEvents: fgCHQRC1.Col = 0: fgCHQRC1.LeftCol = 0
        If blnRapprochement_Action Then lstRapprochement_Action_Init
   End If
End If

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
fraSelect_Update.Visible = False
filDoc_Archive.Pattern = "*.xxx"

fraPériode.Visible = False
fraPériode.Top = SSTab1.Top + 120
fraPériode.Left = SSTab1.Left + SSTab1.Width - fraPériode.Width - 300

'cmdSelect_Ok_Click

mnuImport.Enabled = CHQ_SCAN_Aut.Xspécial
mnuExport.Enabled = CHQ_SCAN_Aut.Xspécial
mnuDelete.Enabled = CHQ_SCAN_Aut.Xspécial
mnuUpdate_Table.Enabled = CHQ_SCAN_Aut.Xspécial
mnuUpdate_DEON.Enabled = CHQ_SCAN_Aut.Xspécial
mnuDéontologie.Enabled = CHQ_SCAN_Aut.Xspécial
mnuDéontologie_Image.Enabled = CHQ_SCAN_Aut.Xspécial
''mnuDéontologie_ZIB.Enabled = CHQ_SCAN_Aut.Xspécial

mnuRapprochement_SAB.Enabled = CHQ_SCAN_Aut.Xspécial
mnuRapprochement_CHQ_SCAN.Enabled = CHQ_SCAN_Aut.Xspécial
mnuRapprochement_Semi_Automatique.Enabled = CHQ_SCAN_Aut.Xspécial
mnuRapprochement_Manuel.Enabled = CHQ_SCAN_Aut.Rapprocher
mnuRapprochement_Automatique.Enabled = CHQ_SCAN_Aut.Rapprocher

lstRapprochement_Action.Visible = False
lstRapprochement_Action.BackColor = &HC0FFC0
blnfgCHQRC1_DisplayLine = False
cmdRapprochement_Update.Visible = False
meCV1.DeviseN = 0
meCV1.Montant = 0


meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
meCV2.OpéAmj = YBIATAB0_DATE_CPT_J

fgDossier.Visible = False
fraCHQ.Visible = False
blnControl = True



End Sub


Public Sub cmdImport_OK()
Dim xIn As String
Dim x As String
Dim V
Dim K As Integer
Dim wFileName_Source As String, wIdFile_Source As Integer
On Error GoTo Error_Handle

wFileName_Source = "C:\Temp\Cheque.txt"
Call lstErr_AddItem(lstErr, cmdContext, "Import : " & wFileName_Source): DoEvents

Call lstErr_Clear(lstErr, cmdContext, "Source : " & wFileName_Source)
V = File_Export_Monitor("Input", wIdFile_Source, wFileName_Source)
If Not IsNull(V) Then Exit Sub

Do Until EOF(wIdFile_Source)
    Line Input #wIdFile_Source, xIn
    K = 0
    xCHQ_SCAN.Id = CSV_Scan(xIn, K)
    xCHQ_SCAN.Cmc7 = CSV_Scan(xIn, K)
    xCHQ_SCAN.Zone4 = CSV_Scan(xIn, K)
    xCHQ_SCAN.Zone3 = CSV_Scan(xIn, K)
    xCHQ_SCAN.Zone2 = CSV_Scan(xIn, K)
    xCHQ_SCAN.Zone1 = CSV_Scan(xIn, K)
    xCHQ_SCAN.PATH = CSV_Scan(xIn, K)
    xCHQ_SCAN.IMAGE = CSV_Scan(xIn, K)
    xCHQ_SCAN.Date = CSV_Scan(xIn, K)
    xCHQ_SCAN.COMPTE = CSV_Scan(xIn, K)
    xCHQ_SCAN.NumLot = CSV_Scan(xIn, K)
    xCHQ_SCAN.CRem = CSV_Scan(xIn, K)
    xCHQ_SCAN.Zone24 = CSV_Scan(xIn, K)
    xCHQ_SCAN.DateHourScan = CSV_Scan(xIn, K)
    xCHQ_SCAN.DateHourSaisie = CSV_Scan(xIn, K)
    xCHQ_SCAN.StatutRem = CSV_Scan(xIn, K)
    xCHQ_SCAN.MotifNonAJ = CSV_Scan(xIn, K)
    xCHQ_SCAN.Saisie = CSV_Scan(xIn, K)
    xCHQ_SCAN.RefClient = CSV_Scan(xIn, K)
    xCHQ_SCAN.RefInterne = CSV_Scan(xIn, K)
    xCHQ_SCAN.Nature = CSV_Scan(xIn, K)
    xCHQ_SCAN.Devise = CSV_Scan(xIn, K)


    If optSelect_Archive Then
        Call sqlCHQ_SCAN_Insert(xCHQ_SCAN, cnAdo_CHQ_ARCHIVE)
    Else
        Call sqlCHQ_SCAN_Insert(xCHQ_SCAN, cnAdo_CHQ_LOCAL)
    End If
    
Loop

Close wIdFile_Source

Call lstErr_AddItem(lstErr, cmdContext, "Import : terminé"): DoEvents

Exit Sub

Error_Handle:
Call lstErr_AddItem(lstErr, cmdContext, "Erreur : " & Error)

MsgBox wFileName_Source & " : " & Error, vbCritical, Me.Caption
Close wIdFile_Source

End Sub


Public Sub cmdExport_Ok()
Dim xIn As String
Dim x As String
Dim V
Dim K As Integer
Dim wFileName_Destination As String, wIdFile_Destination As Integer
On Error GoTo Error_Handle

wFileName_Destination = "C:\Temp\Cheque.txt"
Call lstErr_Clear(lstErr, cmdContext, "Export : " & wFileName_Destination): DoEvents

V = File_Export_Monitor("Output", wIdFile_Destination, wFileName_Destination)
If Not IsNull(V) Then Exit Sub

    
x = "select * from CHEQUE  order by DateHourScan , IMAGE"
If optSelect_Archive Then
    Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(x)
Else
    Set rsSab = cnAdo_CHQ_LOCAL.Execute(x)
End If

Do While Not rsSab.EOF
    V = srvCHQ_SCAN_GetBuffer_ODBC(rsSab, xCHQ_SCAN)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmCHQ_SCAN.fgselect_Display"
        '' Exit Sub
     Else
        Print #wIdFile_Destination, xCHQ_SCAN.Id & ";" _
      & xCHQ_SCAN.Cmc7 & ";" & xCHQ_SCAN.Zone4 & ";" _
      & xCHQ_SCAN.Zone3 & ";" & xCHQ_SCAN.Zone2 & ";" _
      & xCHQ_SCAN.Zone1 & ";" & xCHQ_SCAN.PATH & ";" _
      & xCHQ_SCAN.IMAGE & ";" & xCHQ_SCAN.Date & ";" _
      & xCHQ_SCAN.COMPTE & ";" & xCHQ_SCAN.NumLot & ";" _
      & xCHQ_SCAN.CRem & ";" & xCHQ_SCAN.Zone24 & ";" _
      & xCHQ_SCAN.DateHourScan & ";" & xCHQ_SCAN.DateHourSaisie & ";" _
      & xCHQ_SCAN.StatutRem & ";" & xCHQ_SCAN.MotifNonAJ & ";" _
     & xCHQ_SCAN.Saisie & ";" & xCHQ_SCAN.RefClient & ";" _
      & xCHQ_SCAN.RefInterne & ";" & xCHQ_SCAN.Nature & ";" _
      & xCHQ_SCAN.Devise

    End If
    rsSab.MoveNext

Loop

Close wIdFile_Destination

Call lstErr_AddItem(lstErr, cmdContext, "Export : terminé"): DoEvents

Exit Sub

Error_Handle:
Call lstErr_AddItem(lstErr, cmdContext, "Erreur : " & Error)

MsgBox wFileName_Destination & " : " & Error, vbCritical, Me.Caption
Close wIdFile_Destination

End Sub



Public Function param_Init()

param_Init = Null
Call lstErr_Clear(lstErr, cmdContext, "CHQ_SCAN : param_init"): DoEvents

fgSelect.Visible = False

Call DTPicker_Set(txtSelect_Date, DSys)
Call DTPicker_Set(txtPériode_Min, Mid$(YBIATAB0_DATE_CPT_J, 1, 6) & "01")
Call DTPicker_Set(txtPériode_Max, YBIATAB0_DATE_CPT_J)

lstRapprochement_Action_Init

Me.Enabled = True: Me.MousePointer = 0

End Function





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



Private Sub chkSelect_Date_Click()
If chkSelect_Date = "1" Then
    txtSelect_Date.Enabled = True
Else
    txtSelect_Date.Enabled = False
End If

End Sub

Private Sub cmdPériode_Ok_Click()
Dim xWhere As String

Me.Enabled = False: Me.MousePointer = vbHourglass

Call DTPicker_Control(txtPériode_Min, wAMJMin)
Call DTPicker_Control(txtPériode_Max, WAMJMax)
If mPériode_K = "mnuExport_Stat" Then
    cmdExport_Stat
Else

    Select Case SSTab1.Tab
        Case 0
                xWhere = " where ID = 'R' and Date >= '" & wAMJMin & "' and  Date <= '" & WAMJMax & "'"
                
                x = Trim(txtSelect_RefInterne)
                If x <> "" Then xWhere = xWhere & " and RefInterne = '" & x & "'" & " order by RefInterne"
                xWhere = xWhere & " and Nature <> 'GCC' "
                arrCHQ_SCAN_Remise_sql xWhere
                
                fgSelect_Display
                
                cmdPrint_CHQ_Stat
                
                fgSelect_Reset
        Case 2
        
        
                x = " where ( CHQDATE >= " & wAMJMin & " and CHQDATE <= " & WAMJMax & ") " _
                    & " OR    (CHQRC1DCR >= " & wAMJMin - 19000000 & " and CHQRC1DCR <= " & wAMJMin - 19000000 & ") "
                arrYCHQMON0_Sql x
                
                fgCHQRC1_Display
    
    End Select
End If
cmdContext_Quit
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

Select Case SSTab1.Tab
    Case 0:
        If fraCHQ.Visible Then
            cmdPrint_imgCHQ_RectoVerso
        Else
            If fgDossier.Visible Then
                Me.PopupMenu mnuPrint_imgCHQ, vbPopupMenuLeftButton
            Else
                Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
            End If
        End If
    Case 2: Me.PopupMenu mnuPrint2, vbPopupMenuLeftButton
    
End Select
Me.Show
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL()
Dim V
Dim x As String
Dim xWhere As String, xAnd As String, xOrder As String
Dim wAmj7 As Long
On Error GoTo Error_Handler

currentAction = "cmdCHQ_SCAN_SQL"
xOrder = ""

If chkSelect_Date = "1" Then
    Call DTPicker_Control(txtSelect_Date, wAMJMin)
    xWhere = " where Date = '" & wAMJMin & "' and ID = 'R'"
Else
    xWhere = " where  ID = 'R'"
End If

x = Trim(txtSelect_COMPTE)
If x <> "" Then xWhere = xWhere & " and compte like '%" & x & "%'": xOrder = " order by Crem"

x = Trim(txtSelect_RefInterne)
If x <> "" Then xWhere = xWhere & " and RefInterne = '" & x & "'": xOrder = " order by RefInterne"

If chkSelect_GCC <> "1" Then xWhere = xWhere & " and nature <> 'GCC' "

arrCHQ_SCAN_Remise_sql xWhere & xOrder


fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_Delete_Click()
Dim x As String
Dim K As Integer
Dim wPath_Image As String
Me.Enabled = False: Me.MousePointer = vbHourglass

arrCHQ_SCAN_Détail_Sql oldCHQ_SCAN.Date, oldCHQ_SCAN.CRem
x = MsgBox("Confirmer la suppression de cette remise et des " & arrCHQ_SCAN_Détail_Nb & " chèques?", vbYesNo + vbQuestion + vbDefaultButton2, "Suppression de la remise : " & xCHQ_SCAN.CRem)
If x = vbNo Then Exit Sub

'Suppression des enregistrements de la table CHEQUE
'====================================================
If optSelect_Archive Then
    wPath_Image = paramCHQ_SCAN_Image_Archive
    V = sqlCHQ_SCAN_Delete(oldCHQ_SCAN, rsSab, cnAdo_CHQ_ARCHIVE)
Else
    wPath_Image = paramCHQ_SCAN_Image_Local
    V = sqlCHQ_SCAN_Delete(oldCHQ_SCAN, rsSab, cnAdo_CHQ_LOCAL)
End If

'Suppression des images (*.jpg, *.tif) du répertoire MyVision
'====================================================
If IsNull(V) Then
    wPath_Image = wPath_Image & "\" & oldCHQ_SCAN.Date & "\"
    On Error Resume Next
    For K = 1 To arrCHQ_SCAN_Détail_Nb
        x = wPath_Image & "Archive\*" & arrCHQ_SCAN_Détail(K).IMAGE & ".*"
        msFileSystem.DeleteFile x
         x = wPath_Image & "Tampon\*" & arrCHQ_SCAN_Détail(K).IMAGE & ".*"
        msFileSystem.DeleteFile x
   Next K
    fraSelect_Update.Visible = False
    fraSelect_Options.Enabled = True
    cmdSelect_Ok_Click

Else
    MsgBox V, vbCritical, Me.Name & " : cmdSelect_Delete"
    Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents

End If

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> CHQ_SCAN_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
fgDossier.Clear
If blnOk Then
    cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    cmdSelect_SQL
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
    fgDossier.Visible = False

End If
Call lstErr_AddItem(lstErr, cmdContext, "< CHQ_SCAN_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub fgDossier_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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
        fgDossier.Col = fgDossier_arrIndex:  K = CLng(fgDossier.Text)
        Call fgDossier_Color(fgDossier_RowClick, MouseMoveUsr.BackColor, fgDossier_ColorClick)
        xCHQ_SCAN = arrCHQ_SCAN_Détail(K)
        blnImgCHQ_Verso = True
        blnImgCHQ_Recto = True: imgCHQ.Visible = True: imgCHQ_Verso.Visible = False

        imgCHQ_Load
       'imgCHQ.Stretch = True
   End If
End If

End Sub

Private Sub cmdSelect_Update_Ok_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass

If IsNull(fraSelect_Update_Control) Then
    If optSelect_Archive Then
        V = sqlCHQ_SCAN_Update(newCHQ_SCAN, oldCHQ_SCAN, rsSab, cnAdo_CHQ_ARCHIVE)
    Else
        V = sqlCHQ_SCAN_Update(newCHQ_SCAN, oldCHQ_SCAN, rsSab, cnAdo_CHQ_LOCAL)
    End If
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        arrCHQ_SCAN(arrCHQ_SCAN_Index) = newCHQ_SCAN
        xCHQ_SCAN = newCHQ_SCAN
        fgSelect_DisplayLine arrCHQ_SCAN_Index
        fraSelect_Update.Visible = False
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdSelect_Update_Ok"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
End If
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Update_Quit_Click()
fraSelect_Update.Visible = False
End Sub

Private Sub cmdStatut_Ok_Click()
On Error Resume Next
Me.Enabled = False: Me.MousePointer = vbHourglass

SSTab1.Tab = 1

cmdStatut_Ok_Remise
cmdStatut_Ok_DreamFile
cmdStatut_Ok_Codeline
cmdStatut_Ok_MyVision
filDoc.Visible = True
filDoc_Archive.Visible = True
filDoc_Tampon.Visible = True

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdStatut_Ok_DreamFile()
On Error GoTo Exit_sub

wFile = paramCHQ_SCAN_Appli_Local & "\DreamFile\DreamFile.ini"
libStatut_DreamFile_ini = ""
lblStatut_DreamFile_ini = wFile
intFile = FreeFile(0)
Open wFile For Input As #intFile
blnTop1 = False: blnTop2 = False
Do Until EOF(intFile)
    DoEvents
    Line Input #intFile, xIn
    If Not blnTop1 Then
        If Trim(xIn) = "[DONNEES]" Then blnTop1 = True
    Else
        If Mid$(Trim(xIn), 1, 5) = "Date=" Then wDreamFile_Date = Mid$(Trim(xIn), 6, 8)
        If Mid$(Trim(xIn), 1, 8) = "Compteur" Then libStatut_DreamFile_ini = xIn: Exit Do

    End If
    
Loop
Exit_sub:
On Error Resume Next
Close intFile

End Sub

Private Sub cmdStatut_Ok_Codeline()
On Error GoTo Exit_sub
Dim nbLine As Long
nbLine = 0
wFile = paramCHQ_SCAN_Appli_Local & "\MyVision\Codeline.dbl"
libStatut_Codeline_dbl = ""
lblStatut_Codeline_dbl = wFile & " (doublons)"
intFile = FreeFile(0)
Open wFile For Input As #intFile
blnTop1 = False: blnTop2 = False
Do Until EOF(intFile)
    DoEvents
    Line Input #intFile, xIn
        If Trim(xIn) <> "" Then nbLine = nbLine + 1
    
Loop
libStatut_Codeline_dbl = nbLine & " lignes"
Exit_sub:
On Error Resume Next
Close intFile

End Sub


Private Sub cmdStatut_Ok_Remise()
On Error GoTo Exit_sub
Dim nbRemise As Long, curRemise As Currency
Dim nbCheques As Long, curCheques As Currency
nbRemise = 0: curRemise = 0
nbCheques = 0: curCheques = 0

libStatut_Remise = ""
lblStatut_Remise = "remise AJUST.mdb"
libStatut_Cheques = ""
lblStatut_Cheques = "chèques  AJUST.mdb"

Set rsSab = cnAdo_CHQ_LOCAL.Execute("select * from CHEQUE ")

Do While Not rsSab.EOF
        Select Case rsSab("ID")
            Case "R": curRemise = curRemise + rsSab("Zone1"): nbRemise = nbRemise + 1
            Case "C": curCheques = curCheques + rsSab("Zone1"): nbCheques = nbCheques + 1
            Case Else: MsgBox rsSab("GUID"), vbCritical, "cmdStatut_Ok_Remise"
        End Select
    rsSab.MoveNext

Loop

libStatut_Remise = nbRemise & " Remises, total = " & curRemise
libStatut_Cheques = nbCheques & " Chèques, total = " & curCheques
Exit_sub:

End Sub


Private Sub dirListBox_Change()
'filDoc.Pattern = "*.*"
filDoc.PATH = DirListBox.PATH


End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim K As Long
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
        fgSelect.Col = fgSelect_arrIndex:  arrCHQ_SCAN_Index = CLng(fgSelect.Text)
        fgSelect.LeftCol = 0
        oldCHQ_SCAN = arrCHQ_SCAN(arrCHQ_SCAN_Index)
        fraSelect_Update_Display
        
        xRemise = arrCHQ_SCAN(arrCHQ_SCAN_Index)
        arrCHQ_SCAN_Détail_Sql xRemise.Date, xRemise.CRem

        fgDossier_Display
   End If
End If
fgSelect.LeftCol = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)

cnAdo_Close
End Sub

Private Sub mnuArchivage_Click()
Dim IdShell
Dim x As String, Nb As Long
Me.Enabled = False: Me.MousePointer = vbHourglass

SSTab1.Tab = 0

x = MsgBox("Voulez lancer le traitement d'archivage LOCAL => SERVEUR ?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)

If x = vbYes Then
    x = "C:\DreamCheques\DreamMAJServeur\DreamMajServeur.exe"
    IdShell = Shell(x, 1)
    DoEvents
    If IdShell > 0 Then
        ''AppActivate IdShell, True
    End If
'2005.01.24 ARRET DU PROGRAMME
    End
'===============
    Nb = 0
    Do
    Nb = Nb + 1
    x = MsgBox("Le traitement d'archivage est-il terminé ?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption & " / " & Nb)
    Loop Until x = vbYes
End If
Me.Enabled = True: Me.MousePointer = 0
Exit Sub

Error_Handler:
 MsgBox "mnuArchivage " & Error, vbCritical, Me.Caption, False

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuCHQ_Stat_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
mPériode_K = "mnuCHQ_Stat"
SSTab1.Enabled = False
fraPériode.Visible = True
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
blnControl = False
lstErr.Clear: lstErr.Height = 200
If fraCHQ.Visible Then fraCHQ.Visible = False: Exit Sub
If fgDossier.Visible Then fgDossier.Visible = False: Exit Sub
If fraSelect_Update.Visible Then fraSelect_Update.Visible = False: Exit Sub
If fraPériode.Visible Then fraPériode.Visible = False: SSTab1.Enabled = True: Exit Sub
If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
Else
    SSTab1.Tab = SSTab1.Tab - 1
End If

End Sub

Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
  '  fgSelect.Row = fgSelect.TopRow
  '  fgSelect.Col = fgSelect_arrIndex: ' wK1 = fgSelect.Text
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
optSelect_Archive.Visible = False
optSelect_Local.Visible = False

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





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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


Private Sub mnuDelete_Click()
Dim x As String

Me.Enabled = False: Me.MousePointer = vbHourglass
x = MsgBox("Confirmer l'effacement total de la table CHEQUE de Ajust.mdb ?", vbYesNo + vbQuestion + vbDefaultButton2, "mnuDelete")
If x = vbNo Then Exit Sub

'Suppression des enregistrements de la table CHEQUE
'====================================================
x = "delete * from CHEQUE"
Set rsSab = cnAdo_CHQ_LOCAL.Execute(x)

'If optSelect_Archive Then
'    Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(X)
'Else
'    Set rsSab = cnAdo_CHQ_LOCAL.Execute(X)
'End If


Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuDéontologie_Click()
Dim xWhere As String
Dim K As Integer
Dim x As String, xSQL As String

Me.Enabled = False: Me.MousePointer = vbHourglass
cmdDéontologie_Select

Me.Enabled = True: Me.MousePointer = 0


End Sub


Private Sub mnuDéontologie_Image_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
If Trim(txtSelect_COMPTE) = "" Then
    Call MsgBox("Préciser un numéro client", vbQuestion, "CHQ_SCAN : envoi image par e-mail")
    Exit Sub
End If
blnDéontologie_ZIB = False
cmdPrint_imgCHQ_Text = "Image chèque"
cmdDéontologie_Image_SendMail

Me.Enabled = True: Me.MousePointer = 0

End Sub


Public Sub cmdDéontologie_Image_SendMail()
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim wPath As String
Dim wIntitulé As String
Dim xDate As String
Dim K As Long, K2 As Long

cmdSelect_SQL
For K = 1 To arrCHQ_SCAN_Nb
    xRemise = arrCHQ_SCAN(K)
    arrCHQ_SCAN_Détail_Sql xRemise.Date, xRemise.CRem
    xDate = dateImp(xRemise.Date)
    wSendMail.FromDisplayName = "CHQ_IMAGE"
    wSendMail.RecipientDisplayName = "DEON"
    bgColor = "YELLOW"
    wSendMail.Subject = "Détection d'une remise de chèques sur le compte : " & xRemise.COMPTE & " ,le : " & xDate
    wIntitulé = "Le " & xDate & ", remise n° " & xRemise.CRem & " : " & arrCHQ_SCAN_Détail_Nb & " chèques pour " & xRemise.Devise & " " & Format$(CCur(xRemise.Zone1) / 100, "### ### ### ###.00")
    wPath = paramCHQ_SCAN_Image_Archive & "\" & xRemise.Date & "\Archive\"
    wSendMail.Attachment = ""

    For K2 = 1 To arrCHQ_SCAN_Détail_Nb
        xCHQ_SCAN = arrCHQ_SCAN_Détail(K2)
        wSendMail.Attachment = wSendMail.Attachment & wPath & xCHQ_SCAN.IMAGE & ".jpg" & ";"

    Next K2
    wSendMail.Message = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
                    & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
                    & htmlFontColor("MAGENTA") & "<CENTER>" & wSendMail.Subject _
                    & htmlFontColor("BLUE") & "<BR><BR>" & wIntitulé
 
                        

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

Next K




End Sub

Private Sub mnuDéontologie_ZIB_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
If chkSelect_Date <> "1" Then
    Call MsgBox("Préciser une date", vbQuestion, "CHQ_SCAN : recherche Chèque à rejeter")
    Exit Sub
End If
blnDéontologie_ZIB = True
cmdPrint_imgCHQ_Text = "Déontologie Rejet :Image chèque"
cmdDéontologie_ZIB_Control
blnDéontologie_ZIB = False
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuExport_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdExport_Ok
Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub mnuExport_Stat_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
mPériode_K = "mnuExport_Stat"
SSTab1.Enabled = False
fraPériode.Visible = True
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuImport_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdImport_OK
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuPrint_imgCHQ_All_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

cmdPrint_imgCHQ_All
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuPrint2_Liste_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
    cmdPrint_Rapprochement
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuRapprochement_Automatique_Click()

blnRapprochement_Action = CHQ_SCAN_Aut.Rapprocher

mnuRapprochement_SAB_Click
mnuRapprochement_CHQ_SCAN_Click
mnuRapprochement_semi_Automatique_Click
cmdRapprochement_Update_Click

End Sub

Private Sub mnuRapprochement_Display_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
SSTab1.Tab = 2
arrYCHQMON0_Nb = 0
chkRapprochement_Action = "0"
blnRapprochement_Action = False
SSTab1.Enabled = False
fraPériode.Visible = True
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuRapprochement_semi_Automatique_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

SSTab1.Tab = 2
blnRapprochement_Action = CHQ_SCAN_Aut.Rapprocher

x = " where CHQMONSTA = ' '"
arrYCHQMON0_Sql x

For mCHQRC1_Index = 1 To arrYCHQMON0_Nb
    oldYCHQMON0 = arrYCHQMON0(mCHQRC1_Index)
    If oldYCHQMON0.CHQRC1ETA = 1 And oldYCHQMON0.CHQMONSTA = " " Then
        For mCHQMON_Index = 1 To arrYCHQMON0_Nb
            xYCHQMON0 = arrYCHQMON0(mCHQMON_Index)
            If xYCHQMON0.CHQRC1ETA = 0 And xYCHQMON0.CHQMONSTA = " " Then
                If oldYCHQMON0.CHQMONTANT = xYCHQMON0.CHQMONTANT _
                And oldYCHQMON0.CHQCOMPTE = xYCHQMON0.CHQCOMPTE _
                And oldYCHQMON0.CHQRC1DCR = dateIBM(xYCHQMON0.CHQDATE) Then
                    ''Debug.Print xYCHQMON0.CHQMONTANT, mCHQRC1_Index, mCHQMON_Index, arrYCHQMON0(11).CHQMONSTA
                    cmdRapprochement_Validation_Ok "="
                    Exit For
                End If
            End If
        Next mCHQMON_Index
    End If
Next mCHQRC1_Index


chkRapprochement_Action = "0"
fgCHQRC1_Display

GoTo Exit_sub

Error_Handler:
 MsgBox "mnuRapprochement_Auto " & Error, vbCritical, Me.Caption, False
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuRapprochement_CHQ_SCAN_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

SSTab1.Tab = 2

cmdRapprochement_CHQ_SCAN

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdRapprochement_CHQ_SCAN()
Dim xSQL As String
Dim kRem As Integer
Dim wAmj As String, wCRem As String
Dim blnOk As Boolean

optSelect_Archive = True

'Recherche du dernier lot traité : 1- date de dernière maj
'======================================================
wAmj = DSys

Do
    x = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YCHQMON0 where CHQDATE =" & wAmj
    Set rsSab = cnsab.Execute(x)
    If rsSab("Tally") > 0 Then
        blnOk = True
    Else
        wAmj = dateElp("Ouvré", -1, wAmj)
        If wAmj < "20050101" Then
            Call MsgBox("Date de recherche < 2005.01.01", vbCritical, Me.Caption & "cmdRapprochement_CHQ_SCAN")
            Exit Sub
        End If
    End If
Loop Until blnOk

'Recherche du dernier lot traité : 2 -lots du jour triés par N° remise
'===================================================================
wCRem = ""

x = "select *   from " & paramIBM_Library_SABSPE & ".YCHQMON0 where CHQDATE = " & wAmj & " order by CHQCREM"
Set rsSab = cnsab.Execute(x)

Do While Not rsSab.EOF
    wCRem = rsSab("CHQCREM")
    rsSab.MoveNext
Loop


'Recherche du dernier lot traité : 2 -lots du jour triés par N° remise
'===================================================================
xSQL = " where  ID = 'R' and CRem > '" & wCRem & "' and Nature <> 'GCC' order by CRem"

arrCHQ_SCAN_Remise_sql xSQL

For kRem = 1 To arrCHQ_SCAN_Nb
    
    xRemise = arrCHQ_SCAN(kRem)
    rsYCHQMON0_Init xYCHQMON0
    
    xYCHQMON0.CHQRC1DOS = Val(xRemise.CRem)
    xYCHQMON0.CHQDATE = Val(xRemise.Date)
    xYCHQMON0.CHQCOMPTE = xRemise.COMPTE
    If IsNumeric(xYCHQMON0.CHQCOMPTE) Then
        xYCHQMON0.CHQCOMPTE = Val(xRemise.COMPTE)
    Else
        xYCHQMON0.CHQCOMPTE = xRemise.COMPTE
    End If
    
    xYCHQMON0.CHQCREM = xRemise.CRem
    xYCHQMON0.CHQDEVISE = xRemise.Devise
    xYCHQMON0.CHQMONTANT = CCur(xRemise.Zone1) / 100
    'Comptage
    '----------------------------------------------------------------
    xSQL = "select count(*) as Tally from CHEQUE where  ID = 'C' and CRem = '" & xRemise.CRem & "'"
    Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
    xYCHQMON0.CHQNB = rsSab("Tally")
    sqlYCHQMON0_Insert xYCHQMON0
    
Next kRem

GoTo Exit_sub

Error_Handler:
 MsgBox "mnuArchivage " & Error, vbCritical, Me.Caption, False
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuRapprochement_Manuel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

SSTab1.Tab = 2
blnRapprochement_Action = CHQ_SCAN_Aut.Rapprocher

x = " where CHQMONSTA = ' '"
arrYCHQMON0_Sql x

fgCHQRC1_Display

GoTo Exit_sub

Error_Handler:
 MsgBox "mnuRapprochement_Auto " & Error, vbCritical, Me.Caption, False
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub mnuRapprochement_SAB_Click()
Dim x As String, Nb As Long
Dim V

Me.Enabled = False: Me.MousePointer = vbHourglass

SSTab1.Tab = 2

cmdRapprochement_SAB
GoTo Exit_sub

Error_Handler:
 MsgBox "mnuArchivage " & Error, vbCritical, Me.Caption, False
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdRapprochement_SAB()
Dim x As String, Nb As Long
Dim V
Dim wAmj As String, wCRem As String
Dim blnOk As Boolean

optSelect_Archive = True

'Recherche du dernier lot traité YCHQMON0 : 1- date de dernière maj
'======================================================
wAmj = DSys

Do
    x = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YCHQMON0 where CHQRC1DCR =" & (wAmj - 19000000)
    Set rsSab_Local = cnsab.Execute(x)
    If rsSab_Local("Tally") > 0 Then
        blnOk = True
    Else
        wAmj = dateElp("Ouvré", -1, wAmj)
        If wAmj < "20050101" Then
            Call MsgBox("Date de recherche < 2005.01.01", vbCritical, Me.Caption & "cmdRapprochement_SAB")
            Exit Sub
        End If
    End If
Loop Until blnOk

'
'========================================================================================================
x = "select *  from " & paramIBM_Library_SAB & ".ZGUIRC10 where GUIRC1CEF = '1' and GUIRC1DCR >= " & (wAmj - 19000000)

Set rsSab_Local = cnsab.Execute(x)

Do While Not rsSab_Local.EOF
    V = rsZGUIRC10_GetBuffer(rsSab_Local, xZGUIRC10)
    If IsNull(V) Then
        rsYCHQMON0_Init xYCHQMON0
        xYCHQMON0.CHQRC1ETA = xZGUIRC10.GUIRC1ETA
        xYCHQMON0.CHQRC1AGE = xZGUIRC10.GUIRC1AGE
        xYCHQMON0.CHQRC1SER = xZGUIRC10.GUIRC1SER
        xYCHQMON0.CHQRC1SSE = xZGUIRC10.GUIRC1SSE
        xYCHQMON0.CHQRC1OPE = xZGUIRC10.GUIRC1OPE
        xYCHQMON0.CHQRC1DOS = xZGUIRC10.GUIRC1DOS
        xYCHQMON0.CHQRC1DCR = xZGUIRC10.GUIRC1DCR
 'Tester si existe déjà
 '========================
 
        If Not IsNull(rsYCHQMON0_Read(xYCHQMON0)) Then sqlYCHQMON0_Insert xYCHQMON0
    End If
    rsSab_Local.MoveNext
Loop

Exit Sub

Error_Handler:
 MsgBox "mnuArchivage " & Error, vbCritical, Me.Caption, False

End Sub


Private Sub mnuSave_Click()
Dim xSource As String, xDestination As String, xDestination_enCours As String
Dim wAMJHMS As String, X8 As String, xZip As String
Dim blnDate_Ctl As Boolean
On Error GoTo Error_Handler
Me.Enabled = False: Me.MousePointer = vbHourglass
'------------------------------------
cnAdo_Close
'------------------------------------
wAMJHMS = DSys & "_" & time_Hms & "_"
xZip = "C:\Zip"
xDestination = paramCHQ_SCAN_Save & wAMJHMS & "Local_" & paramCHQ_SCAN_Local_Folder
fgSave.Visible = True
blnDate_Ctl = True
filDoc.PATH = xZip
filDoc.Pattern = "*.*"
fgSave.Clear
fgSave.Redraw = False
fgSave.Rows = 1
For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(filDoc.PATH & "\" & filDoc.FileName)
    fgSave.Rows = fgSave.Rows + 1
    fgSave.Row = fgSave.Rows - 1
    fgSave.Col = 0: fgSave.Text = Trim(filDoc.FileName)
    fgSave.Col = 1: fgSave.Text = msFile.DateLastModified
    Call dateJMA6_AMJ(msFile.DateLastModified, X8)
    If X8 < DSys Then blnDate_Ctl = False
Next I
fgSave.Redraw = True
If filDoc.ListCount < 2 Then
        V = "On doit trouver 2 fichiers ZIP (DreamCheques et MyVision)"
        GoTo Error_MsgBox
End If
If Not blnDate_Ctl Then
    x = MsgBox("fichiers antérieurs au " & dateImp(DSys), vbQuestion & vbYesNo, "CHQ_SCAN > mnuSave")
    If x = vbNo Then
        V = "Abandon"
        GoTo Error_MsgBox
    End If
End If
'xSource = paramCHQ_SCAN_Appli_Local
'xDestination = paramCHQ_SCAN_Save & wAMJHMS & "Local_" & paramCHQ_SCAN_Local_Folder
xDestination_enCours = xDestination & "_enCours"
Call lstErr_Clear(lstErr, cmdContext, "Copie : " & paramCHQ_SCAN_Local_Folder): DoEvents
Me.MousePointer = vbHourglass
msFileSystem.CopyFolder xZip, xDestination_enCours
Name xDestination_enCours As xDestination & "_Zip"

'xSource = paramCHQ_SCAN_Image_Local
'xDestination = paramCHQ_SCAN_Save & wAMJHMS & "Local_" & paramCHQ_SCAN_Image_Folder
'xDestination_enCours = xDestination & "_enCours"
'Call lstErr_AddItem(lstErr, cmdContext, "Copie : " & paramCHQ_SCAN_Image_Folder): DoEvents
'Me.MousePointer = vbHourglass
'msFileSystem.CopyFolder xSource, xDestination_enCours
'Name xDestination_enCours As xDestination

'2005.01.06 pas d'accès copie par la ressource xSource = paramCHQ_SCAN_Appli_Archive
'xSource = "\\BIADOCSRVE$\APPLI.BIA\Dreamsearch\CHEQUE.MDB"

xSource = paramCHQ_SCAN_Appli_Archive & "\CHEQUE.MDB"
xDestination = paramCHQ_SCAN_Save & wAMJHMS & "Archive_CHEQUE.MDB"

' 20050405 : Ne plus copier le fichier CHEQUE.MDB
 'Call lstErr_AddItem(lstErr, cmdContext, "Copie : Archive_CHEQUE.MDB"): DoEvents
' msFileSystem.CopyFile xSource, xDestination

 Call lstErr_AddItem(lstErr, cmdContext, "Copie : terminée"): DoEvents
Me.MousePointer = vbHourglass
'------------------------------------
cnAdo_Open
'------------------------------------
fgSave.Visible = False
For I = 1 To fgSave.Rows - 1
    fgSave.Row = I
    fgSave.Col = 0: x = xZip & "\" & Trim(fgSave.Text)
    Kill x
Next I
MsgBox "Traitement terminé", vbInformation, "Copie de : " & xZip & " vers " & paramCHQ_SCAN_Save & wAMJHMS
Me.Enabled = True: Me.MousePointer = 0

Exit Sub
Error_Handler:
V = Error
Error_MsgBox:
MsgBox V, vbCritical, "Copie de : " & xZip & " vers " & xDestination
fgSave.Visible = False
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_List1_Ok
cmdPrint_List2_Ok
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_Liste_SG_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_List2_Ok
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_Liste_BIA_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_List1_Ok
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuUpdate_DEON_Click()
Dim xRem As String
Dim V
Dim xSQL As String
Dim meYBIAMON0 As typeYBIAMON0, oldYBIAMON0 As typeYBIAMON0

Me.Enabled = False: Me.MousePointer = vbHourglass
App_Debug = "mnuUpdate_DEON_Click"
xRem = Trim(InputBox("Dernier N° remise analysée + RAB date de traitement", "Màj SAB073SPE/YBIAMON7.MONJOB"))
If Not IsNumeric(xRem) Then MsgBox "Le n° doit être numérique !", vbCritical, "Bye": GoTo Exit_sub

'--------------------------------------------------------------------------------------
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "CHQ_DEON"
meYBIAMON0.MONSTATUS = ""

V = fctExploitation_Transaction_Control(meYBIAMON0)
If IsNull(V) Then
    oldYBIAMON0 = meYBIAMON0
    meYBIAMON0.MONJOB = Format(xRem, "0000000000")
    meYBIAMON0.MONFILE = ""
    meYBIAMON0.MONSTATUS = ""
    V = fctExploitation_Transaction_End(meYBIAMON0, oldYBIAMON0)
End If

'20050927_JPL xSql = "Insert CHEQUE set CRem = '" & Format$(Val(xRem), "00000000") & "'"
'20050927_JPL X = MsgBox("Confirmer la mise à jour " & xSql, vbYesNo + vbQuestion + vbDefaultButton2, "mnuDelete")
'20050927_JPL If X = vbNo Then Exit Sub

'Update de l'enregistrements 'X' 'DEON' de la table CHEQUE
'====================================================
'20050927_JPL xSql = "Delete from CHEQUE where ID = 'X' and COMPTE = 'DEON'"
'20050927_JPL Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSql)

'20050927_JPL srvCHQ_SCAN_Init autoDéontologie
'20050927_JPL autoDéontologie.ID = "X"
'20050927_JPL autoDéontologie.COMPTE = "DEON"
'20050927_JPL autoDéontologie.CRem = Format$(Val(xRem), "00000000")
'20050927_JPL V = sqlCHQ_SCAN_Insert(autoDéontologie, cnAdo_CHQ_ARCHIVE)
'20050927_JPL If Not IsNull(V) Then MsgBox V, vbCritical, Me.Caption & " mnuUpdate_DEON"

Exit_sub:
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuUpdate_Table_Click()
Dim x As String, xWhere As String
Dim xSQL As String
Dim Nb As Long, Nb_C As Long, K As Long, J As Long
Dim rsUpdate As New ADODB.Recordset
Dim rsAdo As ADODB.Recordset
Dim cnADO_ATHIC As New ADODB.Connection
Dim cnADO_BIA As New ADODB.Connection

Dim arrBIA_Nb(25000) As Integer, arrATHIC_Nb(25000) As Integer
Dim arrBIA_R_Nb(25000) As Integer, arrATHIC_R_Nb(25000) As Integer
Dim arrBIA_C_MTD(25000) As Currency, arrATHIC_C_MTD(25000) As Currency

Dim arrCHQ_SCAN(25000, 24), arrX(24) As String

Dim mCREM As String, blnCompare_All As Boolean

Nb = 0
'GoTo X_Migration
'GoTo X_Suppression

'GoTo X_CsansR

blnCompare_All = True: GoTo X_Compare

'GoTo X_Compare_Champ_R
'GoTo X_Compare_Champ_C
'GoTo X_IMAGE

'goto X_Comptage
Exit Sub
'========================================================================
' Correction des codes nature et devise erronés CHEQUE.mdb
'========================================================================
X_Suppression:

xSQL = "delete  from CHEQUE where CRem in ('00000049','00000051','00000052','00000053')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

xSQL = "delete  from CHEQUE where CRem in ('00006551')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

' ???'00000059',
xSQL = "delete  from CHEQUE where CRem in ('00000165','00000610','00000896'" _
     & ",'00001085','00001273','00001294','00001719'" _
     & ",'00002272','00002296','00002308','00002321'" _
     & ",'00002322','00002416','00002522','00003298'" _
     & ",'00003400','00003535','00003538','00003669','00004410')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

xSQL = "delete  from CHEQUE where CRem in ('00006551')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

xSQL = "delete  from CHEQUE where CRem in ('00010462','00011653','00012905')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

xSQL = "delete  from CHEQUE where CRem in ('00015289','00016254','00016258')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

xSQL = "delete  from CHEQUE where CRem in ('00016406','00019397','00019398','00019390')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

xSQL = "delete  from CHEQUE where CRem in ('00020090','00020395','00020713','00022437')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

xSQL = "delete  from CHEQUE where CRem in ('00023019','00023885')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)


xSQL = "delete  from CHEQUE where Id = 'R' and StatutRem <> 'AJ'"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

Exit Sub
'========================================================================
' Correction des codes nature et devise erronés CHEQUE.mdb
'========================================================================
X_Migration:


Nb = 0
xSQL = "select * from CHEQUE where Id = 'R'  and Nature not in ('SG','BIA','GCC')"

Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

Do While Not rsSab.EOF
    Nb = Nb + 1
    If IsNumeric(rsSab("RefInterne")) Then
        xSQL = "Update CHEQUE set Nature = 'SG'  where Id = 'R' and CREM ='" & rsSab("CREM") & "'"
    Else
        xSQL = "Update CHEQUE set Nature = 'BIA'  where Id = 'R' and CREM ='" & rsSab("CREM") & "'"
   End If
    Set rsUpdate = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

    rsSab.MoveNext

Loop

Call MsgBox("NB : " & Nb, vbInformation, "mnuUpdate_Table_Click")

Nb = 0
xSQL = "select * from CHEQUE where Id = 'R'  and Devise <> 'EUR'"

Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

Do While Not rsSab.EOF
    Select Case rsSab("Devise")
        Case "978", "BIA", "E", "EIR", "EIU", "EJR", "ER", "ERU", "EU", "EUD", "EUE", "EUI", "EYU", "QEU", "SG", "GCC", "DUR", "ETR", "DEV"
            Nb = Nb + 1
            xSQL = "Update CHEQUE set Devise = 'EUR'  where Id = 'R' and CREM ='" & rsSab("CREM") & "'"
            Set rsUpdate = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
        Case "400"
            Nb = Nb + 1
            xSQL = "Update CHEQUE set Devise = 'USD'  where Id = 'R' and CREM ='" & rsSab("CREM") & "'"
            Set rsUpdate = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
        Case "CHG"
            Nb = Nb + 1
            xSQL = "Update CHEQUE set Devise = 'CHF'  where Id = 'R' and CREM ='" & rsSab("CREM") & "'"
            Set rsUpdate = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
   End Select

    rsSab.MoveNext

Loop

Call MsgBox("NB : " & Nb, vbInformation, "mnuUpdate_Table_Click")


Me.Enabled = True: Me.MousePointer = 0

Exit Sub

'========================================================================
' Contrôle CHEQUE_BIA.mdb # CHEQUE_ATHIC.mdb
'========================================================================
X_Compare:
'========================================================================

cnADO_ATHIC = New ADODB.Connection
cnADO_ATHIC.Provider = "Microsoft.Jet.OLEDB.4.0"
cnADO_ATHIC.Mode = adModeReadWrite

x = "C:\TEMP\CHQ_SCAN\Reprise\CHEQUE_ATHIC.mdb"

cnADO_ATHIC.Open x

x = "CRem "

xSQL = "select " & x & " , count(*) from CHEQUE where  DAte > '20040000' and date < '20119999'" _
     & " group by " & x _
     & " order by " & x

Set rsSab = cnADO_ATHIC.Execute(xSQL)

Do While Not rsSab.EOF
    K = Val(rsSab(0))
    arrATHIC_Nb(K) = rsSab(1)
    rsSab.MoveNext
Loop

xSQL = "select CREM from CHEQUE where Id = 'R' and DAte > '20040000' and date < '20119999'" _
     & " order by CREM"

Set rsSab = cnADO_ATHIC.Execute(xSQL)

Do While Not rsSab.EOF
    K = Val(rsSab(0))
    arrATHIC_R_Nb(K) = arrATHIC_R_Nb(K) + 1
    rsSab.MoveNext

Loop

xSQL = "select CREM,Zone1 from CHEQUE where Id = 'C' and DAte > '20040000' and date < '20119999'" _
     & " order by CREM"

Set rsSab = cnADO_ATHIC.Execute(xSQL)

Do While Not rsSab.EOF
    K = Val(rsSab(0))
    arrATHIC_C_MTD(K) = arrATHIC_C_MTD(K) + CCur(Val(rsSab("Zone1"))) / 100
    rsSab.MoveNext

Loop

'_______________________________________________________________________________________

cnADO_BIA = New ADODB.Connection
cnADO_BIA.Provider = "Microsoft.Jet.OLEDB.4.0"
cnADO_BIA.Mode = adModeReadWrite

x = "C:\TEMP\CHQ_SCAN\Reprise\CHEQUE_BIA.mdb"

cnADO_BIA.Open x

x = "CRem "

xSQL = "select " & x & " , count(*) from CHEQUE where  DAte > '20040000' and date < '20119999'" _
     & " group by " & x _
     & " order by " & x

Set rsSab = cnADO_BIA.Execute(xSQL)

Do While Not rsSab.EOF
    K = Val(rsSab(0))
    arrBIA_Nb(K) = rsSab(1)
    rsSab.MoveNext

Loop

xSQL = "select CREM from CHEQUE where Id = 'R' and DAte > '20040000' and date < '20119999'" _
     & " order by CREM"

Set rsSab = cnADO_BIA.Execute(xSQL)

Do While Not rsSab.EOF
    K = Val(rsSab(0))
    arrBIA_R_Nb(K) = arrBIA_R_Nb(K) + 1
    rsSab.MoveNext
Loop

xSQL = "select CREM,Zone1 from CHEQUE where Id = 'C' and DAte > '20040000' and date < '20119999'" _
     & " order by CREM"

Set rsSab = cnADO_BIA.Execute(xSQL)

Do While Not rsSab.EOF
    K = Val(rsSab(0))
    arrBIA_C_MTD(K) = arrBIA_C_MTD(K) + CCur(Val(rsSab("Zone1"))) / 100
    rsSab.MoveNext

Loop

'_______________________________________________________________________________________

For K = 1 To 25000
    If arrBIA_R_Nb(K) = 0 And arrBIA_Nb(K) > 0 Then
        Debug.Print "C sans R : "; K, arrBIA_Nb(K), arrATHIC_Nb(K)
        arrBIA_Nb(K) = 0
        arrBIA_C_MTD(K) = 0
    End If
    If arrBIA_R_Nb(K) > 0 And arrBIA_Nb(K) = 0 Then
        Debug.Print "R sans C : "; K, arrBIA_Nb(K), arrATHIC_Nb(K)
    End If
    If arrBIA_R_Nb(K) > 1 And arrBIA_Nb(K) > 0 Then
        arrBIA_Nb(K) = arrBIA_Nb(K) - arrBIA_R_Nb(K) + 1
    End If
    
    If arrBIA_Nb(K) <> arrATHIC_Nb(K) Then
        Debug.Print "BIA # ATHIC nb: "; K, arrBIA_Nb(K), arrATHIC_Nb(K)
    End If
    If arrBIA_C_MTD(K) <> arrATHIC_C_MTD(K) Then
        Debug.Print "BIA # ATHIC mt: "; K, arrBIA_C_MTD(K), arrATHIC_C_MTD(K)
    End If
Next K

cnADO_ATHIC.Close
cnADO_BIA.Close
If Not blnCompare_All Then Exit Sub


'========================================================================
' Contrôle CHEQUE_BIA.mdb # CHEQUE_ATHIC.mdb
'========================================================================
X_Compare_Champ_R:
'========================================================================

cnADO_ATHIC = New ADODB.Connection
cnADO_ATHIC.Provider = "Microsoft.Jet.OLEDB.4.0"
cnADO_ATHIC.Mode = adModeReadWrite

x = "C:\TEMP\CHQ_SCAN\Reprise\CHEQUE_ATHIC.mdb"

cnADO_ATHIC.Open x

xSQL = "select * from CHEQUE where Id = 'R' and DAte > '20040000' and date < '20119999'" _
     & " order by CREM,DateHourScan"
    
Set rsSab = cnADO_ATHIC.Execute(xSQL)
Open "C:\TEMP\CHQ_SCAN\Compare_ATHIC\Cheque_R.txt" For Output As #1


Do While Not rsSab.EOF
    For J = 0 To 24
        If IsNull(rsSab(J)) Then
            arrX(J) = ""
        Else
            arrX(J) = rsSab(J)
        End If
    Next J
    Print #1, arrX(14) & ";" & arrX(5) & ";" & arrX(8) & ";" _
            & arrX(0) & ";" & arrX(1) & ";" & arrX(3) & ";" & arrX(4) & ";" & arrX(5) & ";" & arrX(6) & ";" _
            & arrX(7) & ";" & arrX(8) & ";" & arrX(11) & ";" & arrX(12) & ";" & arrX(14) & ";" _
            & arrX(16) & ";" & arrX(18) & ";" & arrX(20) & ";" & arrX(21) & ";" & arrX(22) & ";" _
            & arrX(23) & ";" & arrX(24)
    rsSab.MoveNext
Loop
Close #1


cnADO_BIA = New ADODB.Connection
cnADO_BIA.Provider = "Microsoft.Jet.OLEDB.4.0"
cnADO_BIA.Mode = adModeReadWrite

x = "C:\TEMP\CHQ_SCAN\Reprise\CHEQUE_BIA.mdb"

cnADO_BIA.Open x

xSQL = "select * from CHEQUE where Id = 'R' and DAte > '20040000' and date < '20119999'" _
     & " order by CREM,DateHourScan"
    
Set rsSab = cnADO_BIA.Execute(xSQL)
Open "C:\TEMP\CHQ_SCAN\Compare_BIA\Cheque_R.txt" For Output As #1


Do While Not rsSab.EOF
    For J = 0 To 24
        If IsNull(rsSab(J)) Then
            arrX(J) = ""
        Else
            arrX(J) = rsSab(J)
        End If
    Next J
    Print #1, arrX(14) & ";" & arrX(5) & ";" & arrX(8) & ";" _
            & arrX(0) & ";" & arrX(1) & ";" & arrX(3) & ";" & arrX(4) & ";" & arrX(5) & ";" & arrX(6) & ";" _
            & arrX(7) & ";" & arrX(8) & ";" & arrX(11) & ";" & arrX(12) & ";" & arrX(14) & ";" _
            & arrX(16) & ";" & arrX(18) & ";" & arrX(20) & ";" & arrX(21) & ";" & arrX(22) & ";" _
            & arrX(23) & ";" & arrX(24)
    rsSab.MoveNext
Loop
Close #1

cnADO_ATHIC.Close
cnADO_BIA.Close


If Not blnCompare_All Then Exit Sub


'========================================================================
' Contrôle CHEQUE_BIA.mdb # CHEQUE_ATHIC.mdb
'========================================================================
X_Compare_Champ_C:
'========================================================================

cnADO_ATHIC = New ADODB.Connection
cnADO_ATHIC.Provider = "Microsoft.Jet.OLEDB.4.0"
cnADO_ATHIC.Mode = adModeReadWrite

x = "C:\TEMP\CHQ_SCAN\Reprise\CHEQUE_ATHIC.mdb"

cnADO_ATHIC.Open x

xSQL = "select * from CHEQUE where Id = 'C' and DAte > '20040000' and date < '20119999'" _
     & " order by CREM, Zone4,Zone1,DateHourScan"
    
Set rsSab = cnADO_ATHIC.Execute(xSQL)
Open "C:\TEMP\CHQ_SCAN\Compare_ATHIC\Cheque_C.txt" For Output As #1


Do While Not rsSab.EOF
    For J = 0 To 24
        If IsNull(rsSab(J)) Then
            arrX(J) = ""
        Else
            arrX(J) = rsSab(J)
        End If
    Next J
    Print #1, arrX(14) & ";" & arrX(5) & ";" & arrX(8) & ";" & arrX(16) & ";" _
            & arrX(0) & ";" & arrX(1) & ";" & arrX(3) & ";" & arrX(4) & ";" & arrX(5) & ";" & arrX(6) & ";" _
            & arrX(7) & ";" & arrX(8) & ";" & arrX(11) & ";" & arrX(14) & ";" _
             & arrX(20) & ";" & arrX(21) & ";" & arrX(22)
    rsSab.MoveNext
Loop
Close #1


cnADO_BIA = New ADODB.Connection
cnADO_BIA.Provider = "Microsoft.Jet.OLEDB.4.0"
cnADO_BIA.Mode = adModeReadWrite

x = "C:\TEMP\CHQ_SCAN\Reprise\CHEQUE_BIA.mdb"

cnADO_BIA.Open x

xSQL = "select * from CHEQUE where Id = 'C' and DAte > '20040000' and date < '20119999'" _
     & " order by CREM, Zone4,Zone1,DateHourScan"
    
Set rsSab = cnADO_BIA.Execute(xSQL)
Open "C:\TEMP\CHQ_SCAN\Compare_BIA\Cheque_C.txt" For Output As #1


Do While Not rsSab.EOF
    For J = 0 To 24
        If IsNull(rsSab(J)) Then
            arrX(J) = ""
        Else
            arrX(J) = rsSab(J)
        End If
    Next J
    Print #1, arrX(14) & ";" & arrX(5) & ";" & arrX(8) & ";" & arrX(16) & ";" _
            & arrX(0) & ";" & arrX(1) & ";" & arrX(3) & ";" & arrX(4) & ";" & arrX(5) & ";" & arrX(6) & ";" _
            & arrX(7) & ";" & arrX(8) & ";" & arrX(11) & ";" & arrX(14) & ";" _
            & arrX(20) & ";" & arrX(21) & ";" & arrX(22)
    rsSab.MoveNext
Loop
Close #1

cnADO_ATHIC.Close
cnADO_BIA.Close


If Not blnCompare_All Then Exit Sub

'========================================================================
' Contrôle CHEQUE_ATHIC.mdb # images
'========================================================================
X_IMAGE:
'========================================================================

cnADO_ATHIC = New ADODB.Connection
cnADO_ATHIC.Provider = "Microsoft.Jet.OLEDB.4.0"
cnADO_ATHIC.Mode = adModeReadWrite

x = "C:\TEMP\CHQ_SCAN\Reprise\CHEQUE_ATHIC.mdb"

cnADO_ATHIC.Open x

Open "C:\TEMP\CHQ_SCAN\Compare_ATHIC\Image_Err.txt" For Output As #1

xSQL = "select * from CHEQUE where Id = 'C' and  DAte > '20040000' and date < '20119999'" _
     & " order by CREM , Image"

Set rsSab = cnADO_ATHIC.Execute(xSQL)

Do While Not rsSab.EOF
    If mCREM <> rsSab("CREM") Then
        mCREM = rsSab("CREM")
        Call lstErr_Clear(lstErr, cmdContext, "> Image CREM :" & mCREM): DoEvents
    End If
    Nb = Nb + 1
    xWhere = "\" & Trim(rsSab("IMAGE")) & ".jpg"
    x = Replace(rsSab("PATH"), "I:\ATHIC\", "\\appsrv2011\")
    If Dir(x & xWhere) = "" Then
        Print #1, "? image : "; rsSab("CREM"), rsSab("date"), xWhere, CCur(Val(rsSab("Zone1"))) / 100
    End If
    xWhere = "\ba" & Trim(rsSab("IMAGE")) & ".jpg"
    If Dir(x & xWhere) = "" Then
        Print #1, "? image : "; rsSab("CREM"), rsSab("date"), xWhere, CCur(Val(rsSab("Zone1"))) / 100
    End If

    rsSab.MoveNext
Loop

Print #1, "Image nb : "; Nb
Close #1
Call lstErr_Clear(lstErr, cmdContext, "= Image CREM :" & mCREM): DoEvents

Exit Sub


'========================================================================
' Comptage CHEQUE.mdb
'========================================================================
X_CsansR:
'========================================================================

cnADO_BIA = New ADODB.Connection
cnADO_BIA.Provider = "Microsoft.Jet.OLEDB.4.0"
cnADO_BIA.Mode = adModeReadWrite

x = "C:\TEMP\CHQ_SCAN\Reprise\CHEQUE_BIA.mdb"

cnADO_BIA.Open x

x = "CRem "

xSQL = "select " & x & " , count(*) from CHEQUE " _
     & " group by " & x _
     & " order by " & x

Set rsSab = cnADO_BIA.Execute(xSQL)

Do While Not rsSab.EOF
    K = Val(rsSab(0))
    arrBIA_Nb(K) = rsSab(1)
    rsSab.MoveNext

Loop

xSQL = "select CREM from CHEQUE where Id = 'R' " _
     & " order by CREM"

Set rsSab = cnADO_BIA.Execute(xSQL)

Do While Not rsSab.EOF
    K = Val(rsSab(0))
    arrBIA_R_Nb(K) = arrBIA_R_Nb(K) + 1
    rsSab.MoveNext

Loop


'_______________________________________________________________________________________
Open "C:\TEMP\CHQ_SCAN\Compare_BIA\Cheque_CsansR.txt" For Output As #1


For K = 1 To 25000
    If arrBIA_R_Nb(K) = 0 And arrBIA_Nb(K) > 0 Then
        Print #1, "C sans R : "; K, arrBIA_Nb(K), arrATHIC_Nb(K)
    End If
    If arrBIA_R_Nb(K) > 0 And arrBIA_Nb(K) = 0 Then
         Print #1, "R sans C : "; K, arrBIA_Nb(K), arrATHIC_Nb(K)
    End If
    
Next K
Exit Sub

'========================================================================
' Comptage CHEQUE.mdb
'========================================================================
X_Comptage:

x = "StatutRem , Nature "

xSQL = "select " & x & " , count(*) from CHEQUE where Id = 'R'" _
     & " group by " & x _
     & " order by " & x

Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

Do While Not rsSab.EOF
    Debug.Print rsSab(0), rsSab(1), rsSab(2)
    rsSab.MoveNext

Loop

xSQL = "select count(*) from CHEQUE where Id = 'C'"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

If Not rsSab.EOF Then Debug.Print "C : ", rsSab(0)


Nb = 0: Nb_C = 0

xSQL = "select CREM from CHEQUE where Id = 'R' and StatutRem <> 'AJ'"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
Do While Not rsSab.EOF
    Nb = Nb + 1
    xSQL = "select count(*) from CHEQUE where  ID = 'C' and CRem = '" & rsSab("CRem") & "'"
    Set rsAdo = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
    If Not rsAdo.EOF Then Nb_C = Nb_C + rsAdo(0)

    rsSab.MoveNext

Loop

Debug.Print "<> 'AJ' R : "; Nb & " C : " & Nb_C


xSQL = "select count(*) from CHEQUE where DAte > '20040000' and date < '20119999'"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

If Not rsSab.EOF Then Debug.Print "'2010 : ", rsSab(0)

xSQL = "select count(*) from CHEQUE where DAte > '20040000' and date < '20119999'"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

If Not rsSab.EOF Then Debug.Print "'2010 : ", rsSab(0)
xSQL = "select count(*) from CHEQUE where DAte > '20040000' and date < '20119999'"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

If Not rsSab.EOF Then Debug.Print "'2010 : ", rsSab(0)
xSQL = "select count(*) from CHEQUE where DAte > '20040000' and date < '20119999'"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

If Not rsSab.EOF Then Debug.Print "'2010 : ", rsSab(0)
xSQL = "select count(*) from CHEQUE where DAte > '20040000' and date < '20119999'"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

If Not rsSab.EOF Then Debug.Print "'2010 : ", rsSab(0)
xSQL = "select count(*) from CHEQUE where DAte > '20040000' and date < '20119999'"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

If Not rsSab.EOF Then Debug.Print "'2010 : ", rsSab(0)
xSQL = "select count(*) from CHEQUE where DAte > '20040000' and date < '20119999'"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

If Not rsSab.EOF Then Debug.Print "2010 : ", rsSab(0)
xSQL = "select count(*) from CHEQUE where DAte > '20110000' and date < '20119999'"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

If Not rsSab.EOF Then Debug.Print "2011 : ", rsSab(0)


Exit Sub
'========================================================================
' Suppression de remises CsansR et RsansC  CHEQUE.mdb
'========================================================================
X_Suppression_Old:

Me.Enabled = False: Me.MousePointer = vbHourglass
'RsansC

xSQL = "delete  from CHEQUE where CRem in ('00000518','00000524','00000633','00000044','00001294')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
xSQL = "delete  from CHEQUE where CRem in ('00000049','00000051','00000052','00000053','00000056')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
xSQL = "delete  from CHEQUE where CRem in ('00016406','00019392','00019397','00019398','000'20100')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
xSQL = "delete  from CHEQUE where CRem in ('00021197','00022437','00023019','00023885')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)

'CsansR
xSQL = "delete  from CHEQUE where CRem in ('00000610','00001273','00003298','000004410','000006551')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
xSQL = "delete  from CHEQUE where CRem in ('00010462','00011653','00012905','00015289','000016254')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
xSQL = "delete  from CHEQUE where CRem in ('00016258','00020395','00020713')"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)


Exit Sub
'===================================================Me.Enabled = False: Me.MousePointer = vbHourglass

'===================================================
xWhere = "Where Date = '20100802' and CRem = '00021306' and IMAGE = '21400006'"

xSQL = "select * from CHEQUE " & xWhere

If optSelect_Archive Then
    Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
Else
    Set rsSab = cnAdo_CHQ_LOCAL.Execute(xSQL)
End If

If rsSab.EOF Then
    MsgBox "non trouvé", vbCritical, "JPL CHEQUE.msb"
Else
    Debug.Print rsSab("ZONE1"), rsSab("ZONE2")
    xSQL = "Update CHEQUE set ZONE2 = '012322978001' " & xWhere
If optSelect_Archive Then
    Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
Else
    Set rsSab = cnAdo_CHQ_LOCAL.Execute(xSQL)
End If

End If
Me.Enabled = True: Me.MousePointer = 0



Exit Sub

'========================================================================
' Contrôle CHEQUE_BIA.mdb # CHEQUE_ATHIC.mdb
'========================================================================
X_Compare_Champ_X:
'========================================================================

cnADO_ATHIC = New ADODB.Connection
cnADO_ATHIC.Provider = "Microsoft.Jet.OLEDB.4.0"
cnADO_ATHIC.Mode = adModeReadWrite

x = "C:\TEMP\CHQ_SCAN\Reprise\CHEQUE_ATHIC.mdb"

cnADO_ATHIC.Open x

xSQL = "select * from CHEQUE where Id = 'R' and DAte > '20040000' and date < '20119999'" _
     & " order by CREM"
    
Set rsSab = cnADO_ATHIC.Execute(xSQL)

Do While Not rsSab.EOF
    K = Val(rsSab("CREM"))
    For J = 0 To 24
        arrCHQ_SCAN(K, J) = rsSab(J)
    Next J
    rsSab.MoveNext
Loop

cnADO_BIA = New ADODB.Connection
cnADO_BIA.Provider = "Microsoft.Jet.OLEDB.4.0"
cnADO_BIA.Mode = adModeReadWrite

x = "C:\TEMP\CHQ_SCAN\Reprise\CHEQUE_BIA.mdb"

cnADO_BIA.Open x

xSQL = "select * from CHEQUE where Id = 'R' and DAte > '20040000' and date < '20119999'" _
     & " order by CREM"
     
Set rsSab = cnADO_BIA.Execute(xSQL)

Do While Not rsSab.EOF
    K = Val(rsSab("CREM"))
     For J = 0 To 24
        If arrCHQ_SCAN(K, J) <> rsSab(J) Then
            If J = 2 Or J = 9 Or J = 10 Or J = 13 Or J = 17 Then
            Else
                Debug.Print "#Champ : "; K, J, rsSab(J), arrCHQ_SCAN(K, J)
            End If
        End If
    Next J
   rsSab.MoveNext
Loop



Exit Sub

'===================================================Me.Enabled = False: Me.MousePointer = vbHourglass
xSQL = "Update CHEQUE set Nature = 'SG' , Devise = 'EUR' where Id = 'R' and Devise= '0'"


xSQL = "select count(*) as Tally from CHEQUE where  Crem > '00009999'"
If optSelect_Archive Then
    Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
Else
    Set rsSab = cnAdo_CHQ_LOCAL.Execute(xSQL)
End If
Nb = rsSab("Tally")
xSQL = "delete * from CHEQUE  where Crem > '00009999'"
x = MsgBox("Confirmer la modif " & xSQL, vbYesNo + vbQuestion + vbDefaultButton2, "Count = " & Nb)
If x = vbNo Then Exit Sub

'Update des enregistrements de la table CHEQUE
'====================================================
If optSelect_Archive Then
    Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
Else
    Set rsSab = cnAdo_CHQ_LOCAL.Execute(xSQL)
End If
Me.Enabled = True: Me.MousePointer = 0


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


Public Sub cmdPrint_List1_Ok()
Dim iRow As Integer, K As Integer, I As Integer
Dim blnOk As Boolean
Dim xCOMPTEINT As String

fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Etat : " & fgSelect.Rows - 1)

If fgSelect.Rows > 1 Then
    fgSelect_Sort1_Old = -1
    fgSelect_Sort1 = 0: fgSelect_Sort2 = 3: fgSelect_Sort
End If

prtCHQ_SCAN_List1_Open '
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

For iRow = 1 To fgSelect.Rows - 1
    
    fgSelect.Row = iRow
    fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
    meCHQ_SCAN = arrCHQ_SCAN(K)
    If meCHQ_SCAN.Id = "R" Then
        xCOMPTEINT = sqlYBIACPT0_COMPTEINT(meCHQ_SCAN.COMPTE)
    Else
        xCOMPTEINT = ""
    End If
    arrCHQ_SCAN_Détail_Sql meCHQ_SCAN.Date, meCHQ_SCAN.CRem
    prtCHQ_SCAN_List1_Line meCHQ_SCAN, xCOMPTEINT, arrCHQ_SCAN_Détail_Nb, curCHQ_SCAN_Détail

Next iRow
prtCHQ_SCAN_List1_Close
fgSelect.Visible = True
Me.Show
End Sub


Public Sub cmdPrint_List2_Ok()
Dim iRow As Integer, K As Integer, I As Integer
Dim blnOk As Boolean
Dim xCOMPTEINT As String

fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression liste SG : " & fgSelect.Rows - 1)

If fgSelect.Rows > 1 Then
    fgSelect_Sort1_Old = -1
    fgSelect_Sort1 = 0: fgSelect_Sort2 = 3: fgSelect_Sort
End If

prtCHQ_SCAN_List2_Open '
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

For iRow = 1 To fgSelect.Rows - 1
    
    fgSelect.Row = iRow
    fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
    meCHQ_SCAN = arrCHQ_SCAN(K)
    arrCHQ_SCAN_Détail_Sql meCHQ_SCAN.Date, meCHQ_SCAN.CRem
    For K = 1 To arrCHQ_SCAN_Détail_Nb
        xCHQ_SCAN = arrCHQ_SCAN_Détail(K)
        prtCHQ_SCAN_List2_Line xCHQ_SCAN, meCHQ_SCAN
    Next K
Next iRow
prtCHQ_SCAN_List2_Close
fgSelect.Visible = True
Me.Show
End Sub

Public Sub cmdPrint_CHQ_Stat()
Dim iRow As Integer, K As Integer, I As Integer
Dim blnOk As Boolean, blnEUR As Boolean
Dim xCOMPTEINT As String
Dim iDétail As Integer, curX As Currency

ReDim arrCHQ_Stat(101) As typeCHQ_Stat: arrCHQ_Stat_Nb = 0: arrCHQ_Stat_Max = 100

fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Statistiques : " & fgSelect.Rows - 1): DoEvents
Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "Sélection ....."): DoEvents

If fgSelect.Rows > 1 Then
    fgSelect_Sort1_Old = -1
    fgSelect_Sort1 = 0: fgSelect_Sort2 = 3: fgSelect_Sort
End If

prtCHQ_Stat_Open " Statistiques des remises en banque du " & dateImp10(wAMJMin) & " au " & dateImp10(WAMJMax)
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

For iRow = 1 To fgSelect.Rows - 1
    
    fgSelect.Row = iRow
    fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
    meCHQ_SCAN = arrCHQ_SCAN(K)
    Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Sélection : " & meCHQ_SCAN.Date & " " & meCHQ_SCAN.CRem): DoEvents

    arrCHQ_SCAN_Détail_Sql meCHQ_SCAN.Date, meCHQ_SCAN.CRem
    blnOk = False
    For K = 1 To arrCHQ_Stat_Nb
        If arrCHQ_Stat(K).Date = meCHQ_SCAN.Date Then blnOk = True: Exit For
    Next K
    If Not blnOk Then
        arrCHQ_Stat_Nb = arrCHQ_Stat_Nb + 1
        If arrCHQ_Stat_Nb > arrCHQ_Stat_Max Then
             arrCHQ_Stat_Max = arrCHQ_Stat_Max + 100
             ReDim Preserve arrCHQ_Stat(arrCHQ_Stat_Max)
         End If

        K = arrCHQ_Stat_Nb
        prtCHQ_Stat_Z arrCHQ_Stat(K)
        arrCHQ_Stat(K).Date = meCHQ_SCAN.Date
    End If
    
    Select Case Trim(meCHQ_SCAN.Devise)
        Case "EUR", "eur", "", "0": blnEUR = True
        Case Else: blnEUR = False
    End Select
    
  '------------------------------------------------------------------------
  If Not blnEUR Then
        arrCHQ_Stat(K).Remise_Devise = arrCHQ_Stat(K).Remise_Devise + 1
        arrCHQ_Stat(K).Chèque_Devise = arrCHQ_Stat(K).Chèque_Devise + arrCHQ_SCAN_Détail_Nb
    
    Else
    
         
         Select Case Trim(meCHQ_SCAN.Nature)
             Case "SG", "R", "REM", ""
                 arrCHQ_Stat(K).Remise_SG = arrCHQ_Stat(K).Remise_SG + 1
                 arrCHQ_Stat(K).Chèque_SG = arrCHQ_Stat(K).Chèque_SG + arrCHQ_SCAN_Détail_Nb
                 arrCHQ_Stat(K).Montant_SG = arrCHQ_Stat(K).Montant_SG + CCur(Val(meCHQ_SCAN.Zone1)) / 100
             Case "BIA"
                 arrCHQ_Stat(K).Remise_BIA = arrCHQ_Stat(K).Remise_BIA + 1
                 arrCHQ_Stat(K).Chèque_BIA = arrCHQ_Stat(K).Chèque_BIA + arrCHQ_SCAN_Détail_Nb
                 arrCHQ_Stat(K).Montant_BIA = arrCHQ_Stat(K).Montant_BIA + CCur(Val(meCHQ_SCAN.Zone1)) / 100
           Case Else
                 arrCHQ_Stat(K).Remise_Divers = arrCHQ_Stat(K).Remise_Divers + 1
                 arrCHQ_Stat(K).Chèque_Divers = arrCHQ_Stat(K).Chèque_Divers + arrCHQ_SCAN_Détail_Nb
                 arrCHQ_Stat(K).Montant_Divers = arrCHQ_Stat(K).Montant_Divers + CCur(Val(meCHQ_SCAN.Zone1)) / 100
            
        End Select
         
        For iDétail = 1 To arrCHQ_SCAN_Détail_Nb
            curX = CCur(Val(arrCHQ_SCAN_Détail(iDétail).Zone1)) / 100
            If curX < 5000 Then
                 arrCHQ_Stat(K).Remise_Nb1 = arrCHQ_Stat(K).Remise_Nb1 + 1
             Else
                 If curX < 150000 Then
                     arrCHQ_Stat(K).Remise_Nb2 = arrCHQ_Stat(K).Remise_Nb2 + 1
                 Else
                     arrCHQ_Stat(K).Remise_Nb3 = arrCHQ_Stat(K).Remise_Nb3 + 1
                 End If
             End If
        Next iDétail
    End If
        
Next iRow

prtCHQ_Stat_Z arrCHQ_Stat(0)

For K = 1 To arrCHQ_Stat_Nb
    prtCHQ_Stat_Line arrCHQ_Stat(K)
    
    arrCHQ_Stat(0).Remise_SG = arrCHQ_Stat(0).Remise_SG + arrCHQ_Stat(K).Remise_SG
    arrCHQ_Stat(0).Chèque_SG = arrCHQ_Stat(0).Chèque_SG + arrCHQ_Stat(K).Chèque_SG
    arrCHQ_Stat(0).Montant_SG = arrCHQ_Stat(0).Montant_SG + arrCHQ_Stat(K).Montant_SG

    arrCHQ_Stat(0).Remise_BIA = arrCHQ_Stat(0).Remise_BIA + arrCHQ_Stat(K).Remise_BIA
    arrCHQ_Stat(0).Chèque_BIA = arrCHQ_Stat(0).Chèque_BIA + arrCHQ_Stat(K).Chèque_BIA
    arrCHQ_Stat(0).Montant_BIA = arrCHQ_Stat(0).Montant_BIA + arrCHQ_Stat(K).Montant_BIA
    
    arrCHQ_Stat(0).Remise_Divers = arrCHQ_Stat(0).Remise_Divers + arrCHQ_Stat(K).Remise_Divers
    arrCHQ_Stat(0).Chèque_Divers = arrCHQ_Stat(0).Chèque_Divers + arrCHQ_Stat(K).Chèque_Divers
    arrCHQ_Stat(0).Montant_Divers = arrCHQ_Stat(0).Montant_Divers + arrCHQ_Stat(K).Montant_Divers
    
    arrCHQ_Stat(0).Remise_Nb1 = arrCHQ_Stat(0).Remise_Nb1 + arrCHQ_Stat(K).Remise_Nb1
    arrCHQ_Stat(0).Remise_Nb2 = arrCHQ_Stat(0).Remise_Nb2 + arrCHQ_Stat(K).Remise_Nb2
    arrCHQ_Stat(0).Remise_Nb3 = arrCHQ_Stat(0).Remise_Nb3 + arrCHQ_Stat(K).Remise_Nb3

    arrCHQ_Stat(0).Remise_Devise = arrCHQ_Stat(0).Remise_Devise + arrCHQ_Stat(K).Remise_Devise
    arrCHQ_Stat(0).Chèque_Devise = arrCHQ_Stat(0).Chèque_Devise + arrCHQ_Stat(K).Chèque_Devise
Next K

prtCHQ_Stat_Close arrCHQ_Stat(0)
fgSelect.Visible = True
Me.Show
End Sub


Public Sub blnTransaction_Set()
If Not blnTransaction Then
    blnTransaction = True
   ' Set rsSab_Update = cnAdo_CHQ_ARCHIVE.Execute("SET TRANSACTION ISOLATION LEVEL READ COMMITTED")

End If

End Sub


Public Function sqlYBIACPT0_COMPTEINT(lCompte As String) As String
Dim x As String, lenX As Integer
Dim xSQL As String
If blnOff_Line Then
    x = lCompte & " intitulé"
Else
    x = Trim(lCompte)
    lenX = Len(x)
    If lenX > 11 Then x = Mid$(x, lenX - 10, 11)
    xSQL = "select COMPTEINT from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '" & x & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then x = rsSab("COMPTEINT")
End If
sqlYBIACPT0_COMPTEINT = x
End Function

Public Sub fraSelect_Update_Display()
fraSelect_Update.Enabled = False
Call lstErr_Clear(lstErr, cmdContext, "> Modification de l'image : " & oldCHQ_SCAN.IMAGE): DoEvents
txtSelect_Update_Compte = oldCHQ_SCAN.COMPTE
txtSelect_Update_RefClient = oldCHQ_SCAN.RefClient
txtSelect_Update_RefInterne = oldCHQ_SCAN.RefInterne
txtSelect_Update_Nature = oldCHQ_SCAN.Nature
txtSelect_Update_Devise = oldCHQ_SCAN.Devise
If oldCHQ_SCAN.StatutRem = "AJ" Then
    chkSelect_Update_StatutRem = "1"
    chkSelect_Update_StatutRem.Visible = True
Else
    chkSelect_Update_StatutRem = "0"
    chkSelect_Update_StatutRem.Visible = False
End If

If optSelect_Local Then
    cmdSelect_Delete.Visible = True
    cmdSelect_Update_Ok.Visible = True
Else
    cmdSelect_Delete.Visible = CHQ_SCAN_Aut.Xspécial
    cmdSelect_Update_Ok.Visible = CHQ_SCAN_Aut.Xspécial
End If

fraSelect_Update.Visible = True
fraSelect_Update.Enabled = True
End Sub

Public Function fraSelect_Update_Control()
Dim blnSelect_Update_Control As Boolean

blnSelect_Update_Control = True
Call lstErr_Clear(lstErr, cmdContext, "> Contrôle de l'image : " & oldCHQ_SCAN.IMAGE): DoEvents
newCHQ_SCAN = oldCHQ_SCAN
newCHQ_SCAN.COMPTE = Trim(txtSelect_Update_Compte)
newCHQ_SCAN.RefClient = Trim(txtSelect_Update_RefClient)
newCHQ_SCAN.RefInterne = Trim(txtSelect_Update_RefInterne)
newCHQ_SCAN.Nature = Trim(txtSelect_Update_Nature)
newCHQ_SCAN.Devise = Trim(txtSelect_Update_Devise)
If newCHQ_SCAN.COMPTE = "" Then blnSelect_Update_Control = False: Call lstErr_AddItem(lstErr, cmdContext, "? préciser le COMPTE : ")
'If newCHQ_SCAN.RefClient = "" Then blnSelect_Update_Control = False: Call lstErr_AddItem(lstErr, cmdContext, "? préciser le RefClient : ")
If newCHQ_SCAN.RefInterne = "" Then blnSelect_Update_Control = False: Call lstErr_AddItem(lstErr, cmdContext, "? préciser le Refinterne : ")
'If newCHQ_SCAN.Nature = "" Then blnSelect_Update_Control = False: Call lstErr_AddItem(lstErr, cmdContext, "? préciser le Nature : ")
'If newCHQ_SCAN.Devise = "" Then blnSelect_Update_Control = False: Call lstErr_AddItem(lstErr, cmdContext, "? préciser le Devise : ")
If chkSelect_Update_StatutRem = "0" Then newCHQ_SCAN.StatutRem = "NA": newCHQ_SCAN.Zone1 = "000000000001"
If blnSelect_Update_Control Then
    fraSelect_Update_Control = Null
Else
    fraSelect_Update_Control = "? fraSelect_Update_Control"
End If
End Function

Public Sub cnAdo_Close()
On Error Resume Next

cnAdo_CHQ_ARCHIVE.Close
Set cnAdo_CHQ_ARCHIVE = Nothing

cnAdo_CHQ_LOCAL.Close
Set cnAdo_CHQ_LOCAL = Nothing

End Sub

Public Sub cnAdo_Open()
On Error GoTo Error_Handler
Dim x As String
cmdStatut_Ok.Enabled = False

srvCHQ_SCAN_param
cnAdo_CHQ_ARCHIVE.Open paramODBC_DSN_CHQ_SCAN_ARCHIVE

x = paramCHQ_SCAN_Appli_Archive & "\CHEQUE"
If UCase$(x) <> UCase$(cnAdo_CHQ_ARCHIVE.DefaultDatabase) Then
    MsgBox x, vbCritical, "DSN 'CHQ_ARCHIVE' non conforme "
    cnAdo_Info cnAdo_CHQ_ARCHIVE
    End
End If

optSelect_Archive.Visible = True

' si la base locale existe : C:\DreamSearch\DreamFile\Ajust.mdb
x = paramCHQ_SCAN_Appli_Local & "\DreamFile\Ajust.mdb"
If Dir(x) = "" Then
    optSelect_Local.Visible = False
Else
    cnAdo_CHQ_LOCAL.Open paramODBC_DSN_CHQ_SCAN_LOCAL
    x = paramCHQ_SCAN_Appli_Local & "\DreamFile\AJUST"
    If UCase$(x) <> UCase$(cnAdo_CHQ_LOCAL.DefaultDatabase) Then
        MsgBox x, vbCritical, "DSN 'CHQ_LOCAL' non conforme "
        cnAdo_Info cnAdo_CHQ_LOCAL

        End
    End If
    optSelect_Local.Visible = True
    optSelect_Local.Value = True
    cmdStatut_Ok.Enabled = True
End If

Exit Sub

Error_Handler:
blnControl = False
If Not blnAuto Then MsgBox Error

End Sub

Public Sub cmdStatut_Ok_MyVision()
Dim Nb As Long
On Error Resume Next
libStatut_MyVision = ""
DirListBox.PATH = paramCHQ_SCAN_Image_Local
wFile = paramCHQ_SCAN_Image_Local & "\" & wDreamFile_Date
filDoc_Archive.PATH = wFile & "\Archive"
filDoc_Archive.Pattern = "*.*"
filDoc_Tampon.PATH = wFile & "\Tampon"
filDoc_Tampon.Pattern = "*.*"
lblStatut_MyVision = wFile & "\Archive ...... Tampon"



Nb = 0
wFile = wFile & "\Cmc7\Cmc7.txt"

intFile = FreeFile(0)
Open wFile For Input As #intFile
Do Until EOF(intFile)
    DoEvents
    Line Input #intFile, xIn
    Nb = Nb + 1
Loop
libStatut_MyVision = Nb & " lignes Cmc7"
End Sub

Private Sub txtSelect_Update_Devise_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtSelect_Update_Nature_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Public Sub cmdDéontologie_Select()
Dim curX As Currency
Dim xSQL As String
Dim K As Long, kRem As Long
Dim blnOk As Boolean, blnCHQ As Boolean
Dim wRemise_Text As String
Dim wCompte_Intitulé As String
Dim nbMail As Long
Dim nbRem_Devise As Long, nbChèque As Long, curRem As Currency
Dim nbRem As Long, curRem_Devise As Currency
Dim meYBIAMON0 As typeYBIAMON0, oldYBIAMON0 As typeYBIAMON0
Dim lastCRem As Long, CRem_X8 As String

nbMail = 0
nbRem = 0: nbChèque = 0: curRem = 0
nbRem_Devise = 0: curRem_Devise = 0

' Rechercher le dernier 'CRem' traité : fichier ARCHIVE
'-------------------------------------------------------
optSelect_Archive = True



'20050927_JPL xSql = "select * from CHEQUE where ID = 'X' and COMPTE = 'DEON'"

'20050927_JPL Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSql)

'20050927_JPL srvCHQ_SCAN_Init autoDéontologie
'20050927_JPL If Not rsSab.EOF Then
'20050927_JPL     V = srvCHQ_SCAN_GetBuffer_ODBC(rsSab, autoDéontologie)
'20050927_JPL Else
'20050927_JPL     MsgBox "manque enregistrement ID = 'X' et COMPTE = 'DEON'", vbCritical, Me.Caption
'20050927_JPL     Exit Sub
'20050927_JPL End If
'--------------------------------------------------------------------------------------
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "CHQ_DEON"
meYBIAMON0.MONSTATUS = ""

V = fctExploitation_Auto_Control(meYBIAMON0)
If Not IsNull(V) Then Exit Sub

oldYBIAMON0 = meYBIAMON0
lastCRem = Val(meYBIAMON0.MONJOB)
CRem_X8 = Format(lastCRem, "00000000")

If Not IsNull(V) Then
        MsgBox "manque enregistrement SAB073SPE/YBIAMON7.CHQ_SCAN.CHQ_DEON", vbCritical, Me.Caption
    Exit Sub
End If
Call lstErr_Clear(lstErr, cmdContext, "cmdDéontologie_Select > " & lastCRem): DoEvents

xSQL = " where  ID = 'R' and CRem > '" & CRem_X8 & "' and Nature <> 'GCC' order by CRem"

arrCHQ_SCAN_Remise_sql xSQL


For kRem = 1 To arrCHQ_SCAN_Nb
    
    xRemise = arrCHQ_SCAN(kRem)
    blnOk = False
    blnCHQ = False
    curX = CCur(xRemise.Zone1) / 100
    
    If xRemise.Devise = "EUR" Then
        nbRem = nbRem + 1
        curRem = curRem + curX
    Else
        meCV1.DeviseIso = xRemise.Devise
        meCV1.Montant = curX
        Call CV_Calc("J  ", meCV1, meCV2)
        
        If meCV1.Cours = 0 Then meCV2.Montant = meCV1.Montant ' erreur code devise non controlé en saisie
        
        curX = meCV2.Montant
        curRem_Devise = curRem_Devise + curX
        nbRem_Devise = nbRem_Devise + 1

    End If
'Montant > 150 000€ - lecture remise => compte => intitulé ==>
'envoi Email pour tous les chèques > 150 000
' sinon envoi Email poir la remise globale
'----------------------------------------------------------------------------------
    
    If curX >= 150000 Then blnOk = True
    
    If blnOk Then
        arrCHQ_SCAN_Détail_Sql xRemise.Date, xRemise.CRem

        wCompte_Intitulé = "Compte : " & xRemise.COMPTE & " : " & sqlYBIACPT0_COMPTEINT(xRemise.COMPTE)
        wRemise_Text = "Remise n° " & xRemise.CRem & " du " & dateImp10(xRemise.Date) _
                        & " : " & arrCHQ_SCAN_Détail_Nb & " chèques, total  " & Format$(curX, "### ### ### ###.00") & " " & xRemise.Devise
        
        For K = 1 To arrCHQ_SCAN_Détail_Nb
            xCHQ_SCAN = arrCHQ_SCAN_Détail(K)
        'Contre-valeur
            curX = CCur(xCHQ_SCAN.Zone1) / 100
            
            If xRemise.Devise <> "EUR" Then
                meCV1.DeviseIso = xRemise.Devise
                meCV1.Montant = curX
                Call CV_Calc("J  ", meCV1, meCV2)
                curX = meCV2.Montant
            End If
            If curX >= 150000 Then
                nbMail = nbMail + 1
                blnCHQ = True
                srvCHQ_SCAN_SendMail xCHQ_SCAN, wRemise_Text, wCompte_Intitulé
            End If
        Next K
        
        If Not blnCHQ Then
            srvCHQ_SCAN_SendMail xRemise, wRemise_Text, wCompte_Intitulé
            nbMail = nbMail + 1
        End If
        
    End If
    
Next kRem



'Comptage
'--------------------------------------------------------------------------------------------------------------
If arrCHQ_SCAN_Nb > 0 Then meYBIAMON0.MONJOB = Format(Val(arrCHQ_SCAN(arrCHQ_SCAN_Nb).CRem), "0000000000")
'---------------------------------------------------------------------------------------------------------------

xSQL = "select count(*) as Tally from CHEQUE where  ID = 'C' and CRem > '" & CRem_X8 & "'"
Set rsSab = cnAdo_CHQ_ARCHIVE.Execute(xSQL)
nbChèque = rsSab("Tally")

wRemise_Text = nbRem & " Remises en EUR - Total : " & Format$(curRem, "### ### ### ###.00") _
            & "<BR><BR>" & nbRem_Devise & " Remises en DEV - CV : " & Format$(curRem_Devise, "### ### ### ###.00")
            
wCompte_Intitulé = nbChèque & " chèques traités le " & dateImp10(YBIATAB0_DATE_CPT_J) & "( remises : " & CRem_X8 & " - " & meYBIAMON0.MONJOB & " )"
srvCHQ_SCAN_SendMail_Stat wRemise_Text, wCompte_Intitulé

Call lstErr_AddItem(lstErr, cmdContext, wCompte_Intitulé): DoEvents
'Mise à jour du compteur CRem pour la prochaine exploitation
'----------------------------------------------------------------

'20050927_JPL     newCHQ_SCAN = autoDéontologie
'20050927_JPL     newCHQ_SCAN.CRem = arrCHQ_SCAN(arrCHQ_SCAN_Nb).CRem
    
'20050927_JPL     V = sqlCHQ_SCAN_Update(newCHQ_SCAN, autoDéontologie, rsSab, cnAdo_CHQ_ARCHIVE)
'--------------------------------------------------------------------------------------
    
meYBIAMON0.MONSTATUS = ""
meYBIAMON0.MONFILE = YBIATAB0_DATE_CPT_J
V = fctExploitation_Transaction_End(meYBIAMON0, oldYBIAMON0)
'--------------------------------------------------------------------------------------

Call lstErr_AddItem(lstErr, cmdContext, "cmdDéontologie_Select = " & meYBIAMON0.MONJOB): DoEvents

End Sub


Public Sub lstRapprochement_Action_Init()

' ATTENTION lstRapprochement_Action.ListIndex = 0 déchenche une action si lstRapprochement_Action.Visible = true
'======================================================================
lstRapprochement_Action.Visible = False
lstRapprochement_Action.Clear
If mCHQRC1_Index > 0 Then
    Select Case arrYCHQMON0(mCHQRC1_Index).CHQMONSTA
        Case Is = " "
            lstRapprochement_Action.AddItem "1 - Rapprocher avec chèques numérisés"
            lstRapprochement_Action.AddItem "3 - Ignorer cette remise"
            lstRapprochement_Action.AddItem "4 - Supprimer cette remise"
        Case Is = "="
            lstRapprochement_Action.AddItem "2 - Annuler le rapprochement"
        Case Is = "S", "I"
            lstRapprochement_Action.AddItem "5 - Restaurer cette remise"
    End Select
    
    mCHQMON_Index = 0
    fgCHQMON.Row = 0
    Call fgCHQMON_Color(fgCHQMON_RowClick, MouseMoveUsr.BackColor, fgCHQMON_ColorClick)

Else
    If mCHQMON_Index > 0 Then
        Select Case arrYCHQMON0(mCHQMON_Index).CHQMONSTA
            Case Is = " "
                lstRapprochement_Action.AddItem "6 - Ignorer ce lot numérisé"
                lstRapprochement_Action.AddItem "7 - Supprimer ce lot numérisé"
            Case Is = "S", "I"
                lstRapprochement_Action.AddItem "8 - Restaurer cette remise"
        End Select
    End If
End If

lstRapprochement_Action.AddItem "9 - Abandonner cette action"
lstRapprochement_Action.ListIndex = 0

lstRapprochement_Action.Visible = True

End Sub

Public Sub cmdRapprochement_Validation()
Dim xValidation As String, x As String
Dim blnCHQMONTANT As Boolean
Dim blnCHQCOMPTE As Boolean
Dim blnCHQDATE As Boolean

Dim curX As Currency

xValidation = ""
If arrYCHQMON0(mCHQMON_Index).CHQMONSTA <> " " Then
    Call MsgBox("déjà rapproché", vbCritical, "Rapprochement impossible")
    Exit Sub
End If
' Habilité à valider <> montant supèrieure à 1 € ?
'---------------------------------------------------
blnCHQMONTANT = True

curX = arrYCHQMON0(mCHQRC1_Index).CHQMONTANT - arrYCHQMON0(mCHQMON_Index).CHQMONTANT
If curX <> 0 Then
    xValidation = "Ecart montant :  " & curX & vbCrLf
    '$JPL 20060522 If Abs(curX) > 1 Then
    blnCHQMONTANT = False
End If

' Habilité à valider <> COMPTE ?
'---------------------------------------------------
blnCHQCOMPTE = True

If arrYCHQMON0(mCHQRC1_Index).CHQCOMPTE <> arrYCHQMON0(mCHQMON_Index).CHQCOMPTE Then
    xValidation = xValidation & "Comptes différents" & vbCrLf
     blnCHQMONTANT = False
End If

' Habilité à valider <> DATE ?
'---------------------------------------------------
blnCHQDATE = True

If arrYCHQMON0(mCHQRC1_Index).CHQRC1DCR <> dateIBM(arrYCHQMON0(mCHQMON_Index).CHQDATE) Then
    xValidation = xValidation & "Dates différentes" & vbCrLf
     blnCHQDATE = False
End If

If blnCHQMONTANT And blnCHQCOMPTE And blnCHQDATE Then
    cmdRapprochement_Validation_Ok "="
    'fgCHQRC1_Display

Else
'20050628 JPl : suppression du contrôle
'-----------------------------------------
'    If Not CHQ_SCAN_Aut.Xspécial Then
'        Call MsgBox(xValidation, vbCritical, "Rapprochement impossible")
'    Else
        x = MsgBox(xValidation, vbQuestion + vbYesNo, "Confirmez le rapprochement ?")
        If x = vbYes Then cmdRapprochement_Validation_Ok "M"      ': fgCHQRC1_Display

'    End If
End If

End Sub

Public Sub cmdRapprochement_Validation_Ok(lCHQMONSTA As String)
On Error Resume Next
arrYCHQMON0_Link(mCHQRC1_Index) = mCHQMON_Index
arrYCHQMON0_Link(mCHQMON_Index) = mCHQRC1_Index
arrYCHQMON0(mCHQRC1_Index).CHQMONSTA = lCHQMONSTA
arrYCHQMON0(mCHQMON_Index).CHQMONSTA = lCHQMONSTA

fgCHQRC1.Col = 0: fgCHQRC1.Text = lCHQMONSTA
fgCHQMON.Col = 0: fgCHQMON.Text = lCHQMONSTA

End Sub

Public Sub cmdPrint_Rapprochement()
Dim iRow As Integer, K As Integer, I As Integer
Dim iRowMax As Integer
Dim blnOk As Boolean
Dim xCOMPTEINT As String

fgCHQRC1.Visible = False
fgCHQMON.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Etat Rapprochement: " & fgSelect.Rows - 1)


prtCHQ_SCAN_Rapprochement_Open '
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
If fgCHQRC1.Rows > fgCHQMON.Rows Then
    iRowMax = fgCHQRC1.Rows - 1
Else
    iRowMax = fgCHQMON.Rows - 1
End If
For iRow = 1 To iRowMax
    prtCHQ_SCAN_Rapprochement_NewLine

    If iRow <= fgCHQRC1.Rows - 1 Then
        fgCHQRC1.Row = iRow
        fgCHQRC1.Col = fgCHQRC1_arrIndex:  K = CLng(fgCHQRC1.Text)
        xYCHQMON0 = arrYCHQMON0(K)
        prtCHQ_SCAN_Rapprochement_Line xYCHQMON0
    End If
    If iRow <= fgCHQMON.Rows - 1 Then
        fgCHQMON.Row = iRow
        fgCHQMON.Col = fgCHQMON_arrIndex:  K = CLng(fgCHQMON.Text)
        xYCHQMON0 = arrYCHQMON0(K)
        prtCHQ_SCAN_Rapprochement_Line xYCHQMON0
    End If
    
Next iRow
prtCHQ_SCAN_Rapprochement_Close
fgCHQRC1.Visible = True
fgCHQMON.Visible = True
Me.Show

End Sub

Public Sub cmdPrint_imgCHQ()

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX: XPrt.Print "Numérisé le : " & xCHQ_SCAN.DateHourScan;
XPrt.CurrentX = prtMinX + 2500: XPrt.Print "Remise : " & xCHQ_SCAN.CRem;
XPrt.CurrentX = prtMinX + 5000: XPrt.Print "Image : " & xCHQ_SCAN.Date & " - " & xCHQ_SCAN.IMAGE;
XPrt.CurrentX = prtMinX + 7000: XPrt.Print "Compte : " & xRemise.COMPTE;
XPrt.CurrentX = prtMinX + 9000: XPrt.Print "Montant : "; Format$(CCur(xCHQ_SCAN.Zone1) / 100, "### ### ### ###.00") & " " & xRemise.Devise;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.PaintPicture imgCHQ.Picture, prtMinX, XPrt.CurrentY, 9744, 4368
XPrt.CurrentY = XPrt.CurrentY + 4450
If optSelect_Archive Then
    xPath_ImgCHQ = paramCHQ_SCAN_Image_Archive & "\"
Else
    xPath_ImgCHQ = paramCHQ_SCAN_Image_Local & "\"
End If
           
If blnImgCHQ_Verso Then
    XPrt.PaintPicture imgCHQ_Verso.Picture, prtMinX, XPrt.CurrentY, 9744, 4368
    XPrt.CurrentY = XPrt.CurrentY + 4450
End If
XPrt.DrawWidth = 5
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor

End Sub

Public Sub imgCHQ_Load()
Dim x As String

If optSelect_Archive Then
    x = paramCHQ_SCAN_Image_Archive & "\"
Else
    x = paramCHQ_SCAN_Image_Local & "\"
End If
xPath_ImgCHQ = x & Trim(xCHQ_SCAN.Date) & "\Archive\" & Trim(xCHQ_SCAN.IMAGE) & ".jpg"
If Dir(xPath_ImgCHQ) <> "" Then
    imgCHQ.Picture = LoadPicture(xPath_ImgCHQ)
    cmdPrint.Enabled = True
    fraCHQ.Visible = True
Else
    imgCHQ.Picture = LoadPicture("")
End If
If blnImgCHQ_Verso Then
    xPath_ImgCHQ = x & Trim(xCHQ_SCAN.Date) & "\Archive\ba" & Trim(xCHQ_SCAN.IMAGE) & ".jpg"
    If Dir(xPath_ImgCHQ) <> "" Then
        imgCHQ_Verso.Picture = LoadPicture(xPath_ImgCHQ)
    Else
        imgCHQ_Verso.Picture = LoadPicture("")
    End If
End If
End Sub

Public Sub cmdPrint_imgCHQ_RectoVerso()
prtCHQ_imgCHQ_Open "Image chèque"
cmdPrint_imgCHQ
prtCHQ_imgCHQ_Close

End Sub
Public Sub cmdPrint_imgCHQ_All()
Dim kPage As Integer, K As Integer
kPage = 0
blnImgCHQ_Verso = False
prtCHQ_imgCHQ_Open cmdPrint_imgCHQ_Text
For K = 1 To arrCHQ_SCAN_Détail_Nb
    xCHQ_SCAN = arrCHQ_SCAN_Détail(K)
    imgCHQ_Load
    If kPage = 3 Then kPage = 0: frmElpPrt.prtNewPage

    cmdPrint_imgCHQ
    kPage = kPage + 1
Next K
prtCHQ_imgCHQ_Close

End Sub


Public Sub cmdDéontologie_ZIB_Control()
Dim wPath As String
Dim xDate As String
Dim K As Long, K2 As Long, K3 As Long

cmdSelect_SQL
arrYCHQDEON0_Sql

For K = 1 To arrCHQ_SCAN_Nb
    xRemise = arrCHQ_SCAN(K)
    arrCHQ_SCAN_Détail_Sql xRemise.Date, xRemise.CRem

    For K2 = 1 To arrCHQ_SCAN_Détail_Nb
        xCHQ_SCAN = arrCHQ_SCAN_Détail(K2)
        For K3 = 1 To arrYCHQDEON0_Nb
            If xCHQ_SCAN.Zone4 = arrYCHQDEON0(K3).CHQDEONNUM _
            And xCHQ_SCAN.Zone3 = arrYCHQDEON0(K3).CHQDEONZIB _
            And xCHQ_SCAN.Zone2 = Trim(arrYCHQDEON0(K3).CHQDEONZIN) Then
                blnImgCHQ_Verso = True
                blnImgCHQ_Recto = True: imgCHQ.Visible = True: imgCHQ_Verso.Visible = False
        
                imgCHQ_Load
                MsgBox "Suite ?", vbQuestion + vbOKOnly, "CHQ_SCAN : Détection_ZIB"
                fraCHQ.Visible = False
            End If
        Next K3
        
    Next K2

Next K

End Sub

