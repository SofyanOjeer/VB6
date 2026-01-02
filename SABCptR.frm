VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSABCptR 
   AutoRedraw      =   -1  'True
   Caption         =   "SABCptR: Reprise des comptes ==> SAB"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   9420
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   5400
      TabIndex        =   9
      Top             =   0
      Width           =   3500
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   6
      Top             =   465
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "SABCptR.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraOption"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Liste des comptes"
      TabPicture(1)   =   "SABCptR.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgSelect"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Caractéristiques d'un compte"
      TabPicture(2)   =   "SABCptR.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraSC"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraSCReprise"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraSCUpdate"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Import / Export"
      TabPicture(3)   =   "SABCptR.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraImport"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraImport 
         Caption         =   "Import / Export"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   -74760
         TabIndex        =   74
         Top             =   600
         Width           =   8655
         Begin VB.OptionButton optImport_Batch 
            Caption         =   "BATCH (SABCPTRP0)"
            Height          =   255
            Left            =   6000
            TabIndex        =   82
            Top             =   1560
            Width           =   2175
         End
         Begin VB.CommandButton cmdImport_OK 
            BackColor       =   &H00C0FFC0&
            Caption         =   "OK"
            Height          =   975
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   4200
            Width           =   2415
         End
         Begin VB.TextBox txtImport_SAB 
            Height          =   285
            Left            =   2520
            TabIndex        =   78
            Text            =   "D:\Temp\ .txt"
            Top             =   840
            Width           =   4815
         End
         Begin VB.OptionButton optImport_RCOMPTE0_Write 
            Caption         =   "sans objet  RCOMPTE0 -  BIA > SAB"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   2040
            Width           =   3255
         End
         Begin VB.OptionButton optImport_RPLAN0_Write 
            Caption         =   "RPLAN0 -      BIA > SAB"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   1560
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.TextBox txtImport_BIA 
            Height          =   285
            Left            =   2520
            TabIndex        =   75
            Text            =   "D:\Temp\ .txt"
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label lblImport_SAB 
            Caption         =   "fichier d'échange SAB"
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label lblImport_BIA 
            Caption         =   "fichier de travail BIA"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame fraSCUpdate 
         Height          =   1380
         Left            =   5640
         TabIndex        =   65
         Top             =   4800
         Width           =   3570
         Begin VB.CommandButton cmdSCDisplayX 
            Caption         =   "     détail modifications"
            Height          =   495
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdSCDisplay 
            Caption         =   "détail origine"
            Height          =   495
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdSCNext 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Suivant"
            Height          =   465
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdSCPrevious 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Précédent"
            Height          =   450
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdSCOk 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Valider"
            Height          =   990
            Left            =   2730
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame fraSCReprise 
         Caption         =   "Reprise "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   120
         TabIndex        =   56
         Top             =   4800
         Width           =   5460
         Begin VB.ComboBox cboSCSTATUS 
            Height          =   315
            Left            =   4080
            TabIndex        =   57
            Text            =   "Combo1"
            Top             =   510
            Width           =   1245
         End
         Begin VB.Label lblSCSTATUT 
            Caption         =   "Statut reprise"
            Height          =   240
            Left            =   4140
            TabIndex        =   64
            Top             =   195
            Width           =   1050
         End
         Begin VB.Label txtSCUSRAMJ 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            Height          =   240
            Left            =   1695
            TabIndex        =   63
            Top             =   1005
            Width           =   3030
         End
         Begin VB.Label txtSCMODAMJ 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            Height          =   240
            Left            =   1695
            TabIndex        =   62
            Top             =   660
            Width           =   2265
         End
         Begin VB.Label txtSCCREAMJ 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            Height          =   240
            Left            =   1695
            TabIndex        =   61
            Top             =   300
            Width           =   2265
         End
         Begin VB.Label lblSCUSRAMJ 
            Caption         =   "Correction "
            Height          =   180
            Left            =   195
            TabIndex        =   60
            Top             =   1050
            Width           =   1245
         End
         Begin VB.Label lblSCMODAMJ 
            Caption         =   "Modification "
            Height          =   255
            Left            =   180
            TabIndex        =   59
            Top             =   690
            Width           =   1125
         End
         Begin VB.Label lblSCCREAMJ 
            Caption         =   "Création"
            Height          =   240
            Left            =   180
            TabIndex        =   58
            Top             =   330
            Width           =   1125
         End
      End
      Begin VB.Frame fraSC 
         Height          =   4365
         Left            =   90
         TabIndex        =   10
         Top             =   375
         Width           =   9200
         Begin VB.Frame fraSCCompte 
            Caption         =   "Compte"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2070
            Left            =   60
            TabIndex        =   42
            Top             =   100
            Width           =   9100
            Begin VB.TextBox txtSCCPTGEN 
               Height          =   285
               Left            =   7320
               TabIndex        =   73
               Top             =   720
               Width           =   1560
            End
            Begin VB.CheckBox chkSCSUCCES 
               Alignment       =   1  'Right Justify
               Caption         =   "Succession"
               Height          =   195
               Left            =   7470
               TabIndex        =   21
               Top             =   1200
               Width           =   1395
            End
            Begin VB.TextBox txtSCCLOMOT 
               Height          =   285
               Left            =   6450
               TabIndex        =   24
               Top             =   1545
               Width           =   2445
            End
            Begin VB.TextBox txtSCSITUAT 
               Height          =   285
               Left            =   1665
               TabIndex        =   18
               Top             =   1110
               Width           =   660
            End
            Begin VB.TextBox txtSCSECUR 
               Height          =   285
               Left            =   4155
               TabIndex        =   19
               Top             =   1110
               Width           =   465
            End
            Begin VB.TextBox txtSCLORO 
               Height          =   285
               Left            =   5880
               TabIndex        =   20
               Top             =   1140
               Width           =   255
            End
            Begin VB.TextBox txtSCINTITU 
               Height          =   285
               Left            =   4110
               TabIndex        =   13
               Top             =   315
               Width           =   4815
            End
            Begin VB.TextBox txtSCDEVISO 
               Height          =   285
               Left            =   1050
               TabIndex        =   14
               Top             =   705
               Width           =   480
            End
            Begin VB.TextBox txtSCPCEC 
               Height          =   285
               Left            =   5760
               TabIndex        =   17
               Top             =   720
               Width           =   1080
            End
            Begin VB.TextBox txtSCTDC 
               Height          =   285
               Left            =   4140
               TabIndex        =   16
               Top             =   690
               Width           =   840
            End
            Begin VB.TextBox txtSCSABID 
               Height          =   285
               Left            =   1665
               TabIndex        =   15
               Top             =   720
               Width           =   1695
            End
            Begin VB.TextBox txtSCCOMPTE 
               Height          =   285
               Left            =   1665
               TabIndex        =   12
               Top             =   255
               Width           =   1695
            End
            Begin VB.TextBox txtSCDEVISE 
               Height          =   285
               Left            =   1035
               TabIndex        =   11
               Top             =   270
               Width           =   495
            End
            Begin MSComCtl2.DTPicker txtSCOUVAMJ 
               Height          =   300
               Left            =   1695
               TabIndex        =   22
               Top             =   1515
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
               Format          =   19529731
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin MSComCtl2.DTPicker txtSCCLOAMJ 
               Height          =   300
               Left            =   4140
               TabIndex        =   23
               Top             =   1500
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
               Format          =   19529731
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblSCCLOMOT 
               Caption         =   "motif clôture"
               Height          =   210
               Left            =   5475
               TabIndex        =   53
               Top             =   1575
               Width           =   960
            End
            Begin VB.Label lblSCOUVAMJ 
               Caption         =   "Date d'ouverture"
               Height          =   240
               Left            =   105
               TabIndex        =   52
               Top             =   1560
               Width           =   1380
            End
            Begin VB.Label lblSCCLOAMJ 
               Caption         =   "Date de clôture"
               Height          =   240
               Left            =   2970
               TabIndex        =   51
               Top             =   1560
               Width           =   1125
            End
            Begin VB.Label lblSCSITUAT 
               Caption         =   "Code fonctionnement"
               Height          =   270
               Left            =   75
               TabIndex        =   50
               Top             =   1170
               Width           =   1590
            End
            Begin VB.Label lblSCSECUR 
               Caption         =   "Classe de sécurité"
               Height          =   240
               Left            =   2715
               TabIndex        =   49
               Top             =   1140
               Width           =   1335
            End
            Begin VB.Label lblSCLORO 
               Caption         =   "Loro,Nostro,A"
               Height          =   270
               Left            =   4800
               TabIndex        =   48
               Top             =   1140
               Width           =   1080
            End
            Begin VB.Label lblSCPCEC 
               Caption         =   "PCEC"
               Height          =   270
               Left            =   5040
               TabIndex        =   47
               Top             =   720
               Width           =   615
            End
            Begin VB.Label lblSCTDC 
               Caption         =   "Type"
               Height          =   255
               Left            =   3465
               TabIndex        =   46
               Top             =   735
               Width           =   630
            End
            Begin VB.Label lblSCINTITU 
               Caption         =   "Intitulé"
               Height          =   255
               Left            =   3435
               TabIndex        =   45
               Top             =   315
               Width           =   615
            End
            Begin VB.Label lblSCSABID 
               Caption         =   "SAB"
               Height          =   300
               Left            =   90
               TabIndex        =   44
               Top             =   705
               Width           =   375
            End
            Begin VB.Label lblSCCOMPTE 
               Caption         =   "Cobanque"
               Height          =   300
               Left            =   120
               TabIndex        =   43
               Top             =   285
               Width           =   795
            End
         End
         Begin VB.Frame fraSCAlias 
            Caption         =   "Alias"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   60
            TabIndex        =   40
            Top             =   3660
            Width           =   9100
            Begin VB.TextBox txtSCALIASCPT 
               Height          =   285
               Left            =   1680
               TabIndex        =   33
               Top             =   240
               Width           =   2550
            End
            Begin VB.Label lblSCALIASCPT 
               Caption         =   "SIT"
               Height          =   240
               Left            =   120
               TabIndex        =   41
               Top             =   300
               Width           =   510
            End
         End
         Begin VB.Frame fraSCTitulaire 
            Caption         =   "Titulaire"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   60
            TabIndex        =   38
            Top             =   2160
            Width           =   9100
            Begin VB.CheckBox chkSCTITRSP 
               Alignment       =   1  'Right Justify
               Caption         =   "Titulaire responsable"
               Height          =   225
               Left            =   7080
               TabIndex        =   28
               Top             =   270
               Width           =   1785
            End
            Begin VB.CheckBox chkSCTITPRN 
               Alignment       =   1  'Right Justify
               Caption         =   "Titulaire principal"
               Height          =   285
               Left            =   5160
               TabIndex        =   27
               Top             =   240
               Width           =   1785
            End
            Begin VB.CheckBox chkSCTITCPT 
               Alignment       =   1  'Right Justify
               Caption         =   "Compte principal"
               Height          =   240
               Left            =   3480
               TabIndex        =   26
               Top             =   270
               Width           =   1500
            End
            Begin VB.TextBox txtSCTITID 
               Height          =   285
               Left            =   1665
               TabIndex        =   25
               Top             =   225
               Width           =   1425
            End
            Begin VB.Label lblSCTITID 
               Caption         =   "N° client"
               Height          =   225
               Left            =   120
               TabIndex        =   39
               Top             =   300
               Width           =   960
            End
         End
         Begin VB.Frame fraSCRelevé 
            Caption         =   "Relevé"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   60
            TabIndex        =   34
            Top             =   2910
            Width           =   9100
            Begin VB.TextBox txtSCRELNOR 
               Height          =   285
               Left            =   6105
               TabIndex        =   31
               Top             =   180
               Width           =   780
            End
            Begin VB.CheckBox chkSCRELGES 
               Alignment       =   1  'Right Justify
               Caption         =   "Relevé gestionnaire"
               Height          =   210
               Left            =   6975
               TabIndex        =   32
               Top             =   270
               Width           =   1875
            End
            Begin VB.TextBox txtSCRELADR 
               Height          =   240
               Left            =   4635
               TabIndex        =   30
               Top             =   225
               Width           =   345
            End
            Begin VB.TextBox txtSCRELCOD 
               Height          =   285
               Left            =   1665
               TabIndex        =   29
               Top             =   225
               Width           =   225
            End
            Begin VB.Label lblSCRELNOR 
               Caption         =   "N° relevé"
               Height          =   240
               Left            =   5265
               TabIndex        =   37
               Top             =   240
               Width           =   795
            End
            Begin VB.Label lblSCRELADR 
               Caption         =   "Code Adresse"
               Height          =   195
               Left            =   3480
               TabIndex        =   36
               Top             =   285
               Width           =   1005
            End
            Begin VB.Label lblSCRELCOD 
               Caption         =   "Code relevé"
               Height          =   270
               Left            =   75
               TabIndex        =   35
               Top             =   300
               Width           =   1350
            End
         End
      End
      Begin VB.Frame fraOption 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   -74745
         TabIndex        =   8
         Top             =   585
         Width           =   8895
         Begin VB.TextBox txtSelectSABIDMax 
            Height          =   285
            Left            =   4800
            TabIndex        =   70
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtSelectSABIDMin 
            Height          =   285
            Left            =   3000
            TabIndex        =   69
            Top             =   1200
            Width           =   1575
         End
         Begin VB.OptionButton optSelectSCSABID 
            Caption         =   "Compte SAB"
            Height          =   255
            Left            =   360
            TabIndex        =   55
            Top             =   1200
            Width           =   1695
         End
         Begin VB.OptionButton optSelectSCCOMPTE 
            Caption         =   "Compte Cobanque"
            Height          =   255
            Left            =   360
            TabIndex        =   54
            Top             =   480
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.TextBox txtSelectCompteMax 
            Height          =   285
            Left            =   4800
            MaxLength       =   11
            TabIndex        =   2
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtSelectCompteMin 
            Height          =   285
            Left            =   3000
            MaxLength       =   11
            TabIndex        =   1
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton cmdSelect 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Rechercher"
            Height          =   975
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1920
            Width           =   2415
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   5250
         Left            =   -74880
         TabIndex        =   7
         Top             =   600
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   9260
         _Version        =   393216
         Rows            =   1
         Cols            =   13
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   14737632
         ForeColor       =   12582912
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         TextStyleFixed  =   4
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
         AllowUserResizing=   3
         FormatString    =   $"SABCptR.frx":0070
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Picture         =   "SABCptR.frx":0131
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label libRéférenceInterne 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1200
      TabIndex        =   5
      Top             =   0
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOption 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuOpération 
      Caption         =   "Opération"
      Visible         =   0   'False
      Begin VB.Menu mnuOpérationDisplay 
         Caption         =   "Afficher l'enregistrement d'origine"
      End
      Begin VB.Menu mnuOpérationDisplayX 
         Caption         =   "Afficher les modifications"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Print"
      Visible         =   0   'False
      Begin VB.Menu mnuOpérationPrint 
         Caption         =   "Imprimer ce contrat"
      End
      Begin VB.Menu mnuListPrint 
         Caption         =   "Imprimer la liste"
      End
   End
End
Attribute VB_Name = "frmSABCptR"
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
Dim SABCPTRAut As typeAuthorization

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim recSABCPTR As typeSABCPTR, xSABCPTR As typeSABCPTR, mSABCPTR As typeSABCPTR
Dim mXSABCPTR As typeSABCPTR, zSABCPTR As typeSABCPTR

Dim meSABCPTR() As typeSABCPTR, meXSABCPTR() As typeSABCPTR
Dim meSABCPTR_Nb As Integer, meSABCPTR_Index As Integer, meSABCPTR_NbMax As Integer

Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnfgSelect_DisplayLine As Boolean, blnfgEchéance_DisplayLine As Boolean


Dim blnSetfocus As Boolean
Dim blnSelectCompte As Boolean, wSelectCompteMin As String * 11, wSelectCompteMax As String * 11
Dim wSelectSABIDMin As String, wSelectSABIDMax As String
Dim recCompte As typeCompte

Private Sub chkSCRELGES_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkSCRELGES
End Sub


Private Sub chkSCSUCCES_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkSCSUCCES
End Sub


Private Sub chkSCTITCPT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkSCTITCPT
End Sub


Private Sub chkSCTITPRN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkSCTITPRN
End Sub


Private Sub chkSCTITRSP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkSCTITRSP
End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
blnControl = False

lstErr.Clear
If SSTab1.Tab > 0 Then
    SSTab1.Tab = SSTab1.Tab - 1
Else
    If currentAction = "" Then
        If blnMsgBox_Quit Then
            X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
        Else
           X = vbYes
        End If
        If X = vbYes Then Unload Me
    Else
        cmdReset
    End If
End If
End Sub
Public Sub cmdSelect_Control()
Dim lMin As Double, lMax As Double
If Not Me.Enabled Then Exit Sub
Me.Enabled = False

'cmdOk.Visible = False
'cmdSave.Visible = False
blnControl = False
'blnSetfocus = False

lstErr.Clear
lstErr.Height = 200
If optSelectSCCOMPTE Then
    lMin = CLng(Val(Trim(txtSelectCompteMin)))
    lMax = CLng(Val(Trim(txtSelectCompteMax)))
    If lMax = 0 Then lMax = lMin
    If lMax < 100000 Then lMax = lMax * 1000000 + 999999
    If lMin < 100000 Then lMin = lMin * 1000000

    wSelectCompteMin = Format$(lMin, "00000000000")
    wSelectCompteMax = Format$(lMax, "00000000000")
Else
    wSelectCompteMin = "00000000000"
    wSelectCompteMax = "99999999999"
    wSelectSABIDMin = Trim(txtSelectSABIDMin)
    wSelectSABIDMax = Trim(txtSelectSABIDMax)
    If wSelectSABIDMax = "" Then wSelectSABIDMax = wSelectSABIDMin & "99999999"
End If

If lstErr.ListCount > 0 Then
    lstErr.Visible = True
Else
    'cmdOk.Visible = blncmdOk_Visible
    'blnSetfocus = True: currentActiveControl_Name = "cmdOk"
End If

ExitSub:

Me.Enabled = True
    
blnControl = True

End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdImport_OK_Click()
Dim blnOk As Boolean
blnOk = True
Call lstErr_Clear(lstErr, cmdContext, "cmdImport_OK : Initialisation ")

If Trim(txtImport_BIA) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le fichier BIA ")
If Trim(txtImport_SAB) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le fichier SAB ")

If blnOk Then
    If optImport_RPLAN0_Write Then cmdImport_RPLAN0_Write
    If optImport_Batch Then cmdImport_Batch
End If

End Sub

Private Sub cmdImport_Batch()
'****************************************************************
' 1 - FTP   BIAFIL / SABCPTRP0 => d:\Temp\Sab\SABCPTRP0
' 2 - cmdImportBatch
'     - lecture SABCPTRp0 => BIA.mdb (MVTP0)  fichier tampon clé :Origine, Compte,Devise
'     - lecture fichier d'import *.txt
'           - lecture MVTP0 => meSABCPTR () et meXSABCPTR ()
'           - champ à modifier ? => modif meXSABCPTR () ( Addnew ou Update)
'                                   DTAQ pour mise à jour AS400
'****************************************************************

Dim blnUpdate As Boolean
Dim kIn As Integer, Seq As Long, V
On Error GoTo Error_Handle
Dim xIn As String, wSCPCEC As String * 10, wSCSABID As String * 20
MsgBox "!!!!! PRODUCTION ou TEST"

MsgBox "à faire :Me.Enabled = False"
Open Trim(txtImport_BIA) For Input As #1

tableMvtP0_Close
MDB.Execute "delete * from MVTP0"
tableMvtP0_Open
recMvtP0_Init recMvtp0
recMvtp0.Method = "AddNew"

Do Until EOF(1)
    Seq = Seq + 1
    If Seq Mod 1000 = 0 Then Call lstErr_Clear(frmSABCptR.lstErr, frmSABCptR.cmdContext, "import Batch : " & Seq)
    DoEvents
    Line Input #1, xIn
    recMvtp0.Id = mId$(xIn, 1, 15)
    recMvtp0.Text = xIn
    dbMvtP0_Update recMvtp0
Loop


Open Trim(txtImport_SAB) For Input As #2
recSABCPTR_Init zSABCPTR
xSABCPTR = zSABCPTR

Seq = 0
Mid$(MsgTxt, 1, 35 + recSABCPTRLen) = Space$(35 + recSABCPTRLen)
Call lstErr_Clear(frmSABCptR.lstErr, frmSABCptR.cmdContext, "import Batch début: " & Seq)

Do Until EOF(2)
    Seq = Seq + 1
    DoEvents
    Line Input #2, xIn
    kIn = 0
    
    xSABCPTR.SCCOMPTE = CSV_Scan(xIn, kIn)
    xSABCPTR.SCDEVISE = CSV_Scan(xIn, kIn)
    X = CSV_Scan(xIn, kIn) 'devise ISO
    X = CSV_Scan(xIn, kIn) 'N° compte auto
    wSCSABID = CSV_Scan(xIn, kIn)
    X = CSV_Scan(xIn, kIn) 'type de compte G A
    wSCPCEC = CSV_Scan(xIn, kIn)
    
    recMvtp0.Method = "Seek="
    
    recMvtp0.Id = " " & Format$(xSABCPTR.SCCOMPTE, "00000000000") & Format$(xSABCPTR.SCDEVISE, "000")
      
    If tableMvtP0_Read(recMvtp0) = 0 Then
        Mid$(MsgTxt, 35, recSABCPTRLen) = mId$(recMvtp0.Text, 1, recSABCPTRLen)
        MsgTxtIndex = 0
        srvSABCPTR_GetBuffer mSABCPTR

        Mid$(recMvtp0.Id, 1, 1) = "M"
          
        If tableMvtP0_Read(recMvtp0) = 0 Then
            Mid$(MsgTxt, 35, recSABCPTRLen) = mId$(recMvtp0.Text, 1, recSABCPTRLen)
            MsgTxtIndex = 0
            srvSABCPTR_GetBuffer xSABCPTR
            xSABCPTR.Method = constUpdate
           Else
            xSABCPTR = zSABCPTR
            xSABCPTR.SCORIG = "M"
            xSABCPTR.SCCOMPTE = mSABCPTR.SCCOMPTE
            xSABCPTR.SCDEVISE = mSABCPTR.SCDEVISE
            xSABCPTR.Method = constAddNew
        End If
        
        blnUpdate = False
        If Trim(xSABCPTR.SCPCEC) = "" Then
            If mSABCPTR.SCPCEC <> wSCPCEC Then xSABCPTR.SCPCEC = wSCPCEC: blnUpdate = True
        Else
            If xSABCPTR.SCPCEC <> wSCPCEC Then xSABCPTR.SCPCEC = wSCPCEC: blnUpdate = True
        End If
        
         If Trim(xSABCPTR.SCSABID) = "" Then
            If mSABCPTR.SCSABID <> wSCSABID Then xSABCPTR.SCSABID = wSCSABID: blnUpdate = True
        Else
            If xSABCPTR.SCSABID <> wSCSABID Then xSABCPTR.SCSABID = wSCSABID: blnUpdate = True
        End If
       
        If blnUpdate Then
     '''   Exit Sub
            
            xSABCPTR.SCMODAMJ = DSys
            xSABCPTR.SCMODHMS = time_Hms
            If xSABCPTR.SCCREAMJ = "00000000" Then
                xSABCPTR.SCCREAMJ = xSABCPTR.SCMODAMJ
                xSABCPTR.SCCREHMS = xSABCPTR.SCMODHMS
            End If
            
            xSABCPTR.SCUSRNOM = usrId
     '''       Debug.Print recMvtp0.Id
            Call lstErr_Clear(frmSABCptR.lstErr, frmSABCptR.cmdContext, "import PCEC : " & mSABCPTR.SCCOMPTE & " " & mSABCPTR.SCDEVISE)

            xSABCPTR.obj = zSABCPTR.obj ' effacer par _GetBuffer !!!!
            V = srvSABCPTR_Update(xSABCPTR)
            If Not IsNull(V) Then
                MsgBox "erreur : SABCPTR_cmdImport_Batch" & xIn, vbCritical, Error
            End If
        End If
   End If
   ''''GoTo Error_Handle
Loop

tableMvtP0_Close

Close
Call lstErr_Clear(frmSABCptR.lstErr, frmSABCptR.cmdContext, "import Batch fin: " & Seq)
Me.Enabled = True
Exit Sub


Error_Handle:
 MsgBox "erreur : SABCPTR_cmdImport_Batch" & xIn, vbCritical, Error
Close
tableMvtP0_Close


Me.Enabled = True
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

Private Sub cmdImport_RPLAN0_Write()

Dim blnUpdate As Boolean
Dim kIn As Integer, Seq As Long
On Error GoTo Error_Handle
Dim xIn As String, X As String, wRPLAN0 As String * 102, wPCEC As String * 38


Call lstErr_AddItem(lstErr, cmdContext, "cmdSAB_RPLAN0_Write: début"): DoEvents

Open Trim(txtImport_SAB) For Output As #2
Open Trim(txtImport_BIA) For Input As #1
Open Trim("C:\Temp\SAB\SAB_PCEC.txt") For Output As #3

Seq = 0


Do Until EOF(1)
    Seq = Seq + 1
    If Seq Mod 1000 = 0 Then Call lstErr_Clear(frmSABCptR.lstErr, frmSABCptR.cmdContext, "cmdSAB_RPLAN0_Write : " & Seq)
    DoEvents
    Line Input #1, xIn
    xIn = Trim(xIn)
    If xIn <> "" Then
        If IsNumeric(mId$(xIn, 1, 6)) Then
            kIn = 0
            wRPLAN0 = ""
            wPCEC = ""
            
            Mid$(wRPLAN0, 1, 10) = CSV_Scan(xIn, kIn)
            Mid$(wRPLAN0, 11, 32) = CSV_Scan(xIn, kIn)
            Mid$(wRPLAN0, 43, 3) = CSV_Scan(xIn, kIn)
            Mid$(wRPLAN0, 46, 2) = Format(Val(CSV_Scan(xIn, kIn)), "00")
            Mid$(wRPLAN0, 48, 1) = CSV_Scan(xIn, kIn)
            Mid$(wRPLAN0, 49, 1) = CSV_Scan(xIn, kIn)
            Mid$(wRPLAN0, 50, 1) = CSV_Scan(xIn, kIn)
            Mid$(wRPLAN0, 51, 1) = CSV_Scan(xIn, kIn)
            Mid$(wRPLAN0, 52, 1) = CSV_Scan(xIn, kIn)
            Mid$(wRPLAN0, 53, 2) = Format(Val(CSV_Scan(xIn, kIn)), "00")
            
  ' Champs non alimentés
  
            Mid$(wRPLAN0, 55, 1) = "0"  'CSV_Scan(xIn, kIn)
            Mid$(wRPLAN0, 56, 2) = "00" 'Format(Val(CSV_Scan(xIn, kIn)), "00")
            Mid$(wRPLAN0, 58, 5) = "00000" 'Format(Val(CSV_Scan(xIn, kIn)), "00000")
            X = "" 'Trim(CSV_Scan(xIn, kIn))
            If X <> "" Then
                Mid$(wRPLAN0, 63, 32) = X
            Else
                 Mid$(wRPLAN0, 63, 32) = mId$(wRPLAN0, 11, 32)
           End If
            
            Mid$(wRPLAN0, 95, 8) = "        " 'CSV_Scan(xIn, kIn)
            '''Text_Accent wRPLAN0
           Print #2, wRPLAN0
           Print #3, mId$(wRPLAN0, 1, 6) & mId$(wRPLAN0, 11, 32)
        End If
        End If
        
Loop


Close
Call lstErr_Clear(frmSABCptR.lstErr, frmSABCptR.cmdContext, "cmdSAB_RPLAN0_Write fin: " & Seq)
Exit Sub

Error_Handle:
 MsgBox "erreur : cmdSAB_RPLAN0_Write" & xIn, vbCritical, Error
Close


End Sub



Private Sub cmdSCDisplay_Click()
srvSABCPTR_ElpDisplay mSABCPTR

End Sub

Private Sub cmdSCDisplayX_Click()
srvSABCPTR_ElpDisplay mXSABCPTR

End Sub


Private Sub cmdSCNext_Click()
lstErr.Clear
If fgSelect.Row < fgSelect.Rows - 1 Then
    fgSelect.Row = fgSelect.Row + 1
    fgSelect_Click_Ok

    fraSC_Display
Else
    Call lstErr_AddItem(lstErr, cmdContext, "! fin de liste")

End If

End Sub

Private Sub cmdSCOk_Click()
Me.Enabled = False

lstErr.Clear
xSABCPTR = mXSABCPTR: cmdUpdate_Control

If lstErr.ListCount <> 0 Then GoTo Exit_Sub
Me.Enabled = False
fraSC.Enabled = False
'''cmdSCOk.Visible = False

xSABCPTR.SCMODAMJ = DSys
xSABCPTR.SCMODHMS = time_Hms
If xSABCPTR.SCCREAMJ = "00000000" Then
    xSABCPTR.SCCREAMJ = xSABCPTR.SCMODAMJ
    xSABCPTR.SCCREHMS = xSABCPTR.SCMODHMS
End If

xSABCPTR.SCUSRNOM = usrId

cmdUpdate_Db

Exit_Sub:

currentAction = ""
Me.Enabled = True
AppActivate Me.Caption

End Sub

Public Sub cmdUpdate_Db()
If lstErr.ListCount = 0 Then
    V = srvSABCPTR_Update(xSABCPTR)
    If IsNull(V) Then
        meXSABCPTR(meSABCPTR_Index) = xSABCPTR
        mXSABCPTR = xSABCPTR
        fgSelect_DisplayLine
        fraSC_Display
    End If
End If

End Sub

Private Sub cmdSCPrevious_Click()
lstErr.Clear
If fgSelect.Row > 1 Then
    fgSelect.Row = fgSelect.Row - 1
    fgSelect_Click_Ok

    fraSC_Display
Else
    Call lstErr_AddItem(lstErr, cmdContext, "! début de liste")

End If

End Sub

Private Sub cmdSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdSelect

End Sub

Private Sub cmdTEST_Click()
'recSABCPTR_Init mSABCPTR

'fraSC_Display
End Sub

Private Sub fraOption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub mnuOpérationDisplayX_Click()
srvSABCPTR_ElpDisplay mXSABCPTR

End Sub

Private Sub optImport_Batch_Click()
txtImport_BIA = "C:\Temp\SAB\SABCPTRP0"
txtImport_SAB = "C:\Temp\SAB\sabcpt_270802.csv"

End Sub

Private Sub optImport_RCOMPTE0_Write_Click()
txtImport_BIA = "C:\Temp\SAB\RCOMPTE0.BIA"
txtImport_SAB = "\\FR11024427\S820I_In\RCOMPTE0.SAB"

End Sub

Private Sub optImport_RPLAN0_Write_Click()
txtImport_BIA = "C:\Temp\SAB\RPLAN0.csv"
txtImport_SAB = "\\FR11024427\S820I_In\RPLAN0"

End Sub

Private Sub txtSCCLOMOT_LostFocus()
txt_LostFocus txtSCCLOMOT
End Sub

Private Sub txtSCCOMPTE_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub

Private Sub txtSCCOMPTE_LostFocus()
txt_LostFocus txtSCCOMPTE
End Sub

Private Sub txtSCDEVISE_LostFocus()
txt_LostFocus txtSCDEVISE
End Sub

Private Sub txtSCDEVISO_LostFocus()
txt_LostFocus txtSCDEVISO
End Sub


Private Sub txtSCINTITU_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtSCINTITU_LostFocus()
txt_LostFocus txtSCINTITU
End Sub

Private Sub txtSCLORO_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub

Private Sub txtSCLORO_LostFocus()
txt_LostFocus txtSCLORO
End Sub

Private Sub txtSCOUVAMJ_GotFocus()
DTPicker_GotFocus txtSCOUVAMJ

End Sub


Private Sub txtSCOUVAMJ_LostFocus()
DTPicker_LostFocus txtSCOUVAMJ

End Sub


Private Sub txtSCCLOAMJ_GotFocus()
DTPicker_GotFocus txtSCCLOAMJ

End Sub


Private Sub txtSCCLOAMJ_LostFocus()
DTPicker_LostFocus txtSCCLOAMJ
End Sub


Private Sub txtSCCLOMOT_GotFocus()
txt_GotFocus txtSCCLOMOT
End Sub


Private Sub txtSCCOMPTE_GotFocus()
txt_GotFocus txtSCCOMPTE
End Sub


Private Sub txtSCDEVISE_GotFocus()
txt_GotFocus txtSCDEVISE
End Sub


Private Sub txtSCDEVISO_GotFocus()
txt_GotFocus txtSCDEVISO
End Sub


Private Sub txtSCINTITU_GotFocus()
txt_GotFocus txtSCINTITU
End Sub


Private Sub txtSCLORO_GotFocus()
txt_GotFocus txtSCLORO
End Sub


Private Sub txtSCPCEC_GotFocus()
txt_GotFocus txtSCPCEC
End Sub


Private Sub txtSCPCEC_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtSCPCEC_LostFocus()
txt_LostFocus txtSCPCEC
End Sub


Private Sub txtSCRELADR_GotFocus()
txt_GotFocus txtSCRELADR
End Sub


Private Sub txtSCRELADR_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSCRELADR_LostFocus()
txt_LostFocus txtSCRELADR
End Sub


Private Sub txtSCRELCOD_GotFocus()
txt_GotFocus txtSCRELCOD
End Sub


Private Sub txtSCRELCOD_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSCRELCOD_LostFocus()
txt_LostFocus txtSCRELCOD
End Sub


Private Sub txtSCRELNOR_GotFocus()
txt_GotFocus txtSCRELNOR
End Sub


Private Sub txtSCRELNOR_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtSCRELNOR_LostFocus()
txt_LostFocus txtSCRELNOR
End Sub


Private Sub txtSCSABID_GotFocus()
txt_GotFocus txtSCSABID
End Sub


Private Sub txtSCSABID_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSCSABID_LostFocus()
txt_LostFocus txtSCSABID
End Sub


Private Sub txtSCSECUR_GotFocus()
txt_GotFocus txtSCSECUR
End Sub


Private Sub txtSCSECUR_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtSCSECUR_LostFocus()
txt_LostFocus txtSCSECUR
End Sub


Private Sub txtSCALIASCPT_GotFocus()
txt_GotFocus txtSCALIASCPT
End Sub


Private Sub txtSCALIASCPT_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtSCALIASCPT_LostFocus()
txt_LostFocus txtSCALIASCPT
End Sub


Private Sub txtSCSITUAT_GotFocus()
txt_GotFocus txtSCSITUAT
End Sub


Private Sub txtSCSITUAT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSCSITUAT_LostFocus()
txt_LostFocus txtSCSITUAT
End Sub


Private Sub txtSCTDC_GotFocus()
txt_GotFocus txtSCTDC
End Sub


Private Sub txtSCTDC_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtSCTDC_LostFocus()
txt_LostFocus txtSCTDC
End Sub


Private Sub txtSCTITID_GotFocus()
txt_GotFocus txtSCTITID
End Sub


Private Sub txtSCTITID_LostFocus()
txt_LostFocus txtSCTITID
End Sub


Private Sub txtSelectCompteMax_GotFocus()

txt_GotFocus txtSelectCompteMax

End Sub


Private Sub txtSelectCompteMax_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtSelectCompteMax_LostFocus()
txt_LostFocus txtSelectCompteMax

End Sub

Private Sub txtSelectCompteMin_GotFocus()
txt_GotFocus txtSelectCompteMin

End Sub

Private Sub txtSelectCompteMin_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)


End Sub


Private Sub txtSelectCompteMin_LostFocus()
txt_LostFocus txtSelectCompteMin


End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
recSABCPTR_Init zSABCPTR

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False

optSelectSCSABID.Enabled = SABCPTRAut.Xspécial ' !! vue incomplète si changement SABID

fgSelect_Reset

cboSCSTATUS_Init cboSCSTATUS
fraImport.Enabled = SABCPTRAut.Xspécial
optImport_RPLAN0_Write_Click

blnControl = True
End Sub


Public Function Compte_Load(lDevise As String, lCompteNuméro As String)
Compte_Load = Null
recCompte.Devise = lDevise
recCompte.Numéro = lCompteNuméro
If blnRéplication_Load Then
    V = mdbCptP0_Find(recCompte)
Else
    V = "Compte_Load à revoir "  'srvCompteFind(recCompte)
End If

If Not IsNull(V) Then Call lstErr_AddItem(lstErr, lstErr, "? compte inconnu : " & lDevise & lCompteNuméro): Compte_Load = "?": Exit Function

If recCompte.Situation <> " " Then
    Select Case recCompte.Situation
        Case "B": Call lstErr_AddItem(lstErr, lstErr, " ? Compte bloqué : " & lCompteNuméro): Compte_Load = "?"
        Case "A": Call lstErr_AddItem(lstErr, lstErr, " ? Compte annulé : " & lCompteNuméro): Compte_Load = "?"
        Case Else: Call lstErr_AddItem(lstErr, lstErr, " ? Situation du compte : " & lCompteNuméro): Compte_Load = "?"
    End Select
End If

End Function
Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

If lRow > 0 Then
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
Dim K2 As Integer, I As Integer
Dim curDB As Currency, curCR As Currency, curX As Currency

SSTab1.Tab = 1

fgSelect.Visible = True
fgSelect.Clear: fgSelect.Rows = 1: fgSelect_RowDisplay = 0

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Enabled = True
For meSABCPTR_Index = 1 To meSABCPTR_Nb
    If meSABCPTR(meSABCPTR_Index).Method <> constIgnore And meSABCPTR(meSABCPTR_Index).Method <> constDelete Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine
    End If
Next meSABCPTR_Index

fgSelect_SortAD = 5
If fgSelect.Rows > 1 Then fgSelect_Sort

End Sub
Public Sub fgSelect_DisplayLine()
fgSelect.Col = 0: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCSTATUS

fgSelect.Col = 1: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCNATGA
fgSelect.Col = 3: fgSelect.Text = Compte_Imp(meSABCPTR(meSABCPTR_Index).SCCOMPTE)
fgSelect.Col = 2: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCDEVISE
fgSelect.Col = 4: fgSelect.Text = CompteSAB_Imp(meSABCPTR(meSABCPTR_Index).SCSABID)
fgSelect.Col = 5: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCTDC
fgSelect.Col = 6: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCPCEC
fgSelect.Col = 7: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCINTITU

fgSelect.Col = fgSelect_arrIndex - 1: fgSelect.Text = ""
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = meSABCPTR_Index

If meXSABCPTR(meSABCPTR_Index).SCORIG <> " " Then
    For I = 0 To fgSelect_arrIndex
      fgSelect.Col = I: fgSelect.CellForeColor = warnUsrColor
    Next I
End If
End Sub
Public Sub fgSelect_Load()
Dim X As String, mMethod As String

recSABCPTR_Init xSABCPTR
xSABCPTR.SCCOMPTE = wSelectCompteMin
xSABCPTR.SCDEVISE = "000"

meSABCPTR(0) = xSABCPTR
meSABCPTR(0).SCDEVISE = "999"
meSABCPTR(0).SCCOMPTE = wSelectCompteMax

If optSelectSCCOMPTE Then
    xSABCPTR.Method = "SnapP0"
Else
    xSABCPTR.Method = "SnapL2"
    xSABCPTR.SCSABID = wSelectSABIDMin
    meSABCPTR(0).SCSABID = wSelectSABIDMax
End If


Call srvSABCPTR_Load(xSABCPTR, meSABCPTR(0))

'meSABCPTR_Nb = srvSABCPTR.arrSABCPTR_NB
'meSABCPTR_NbMax = meSABCPTR_Nb + 1: ReDim meSABCPTR(meSABCPTR_NbMax)

meSABCPTR_NbMax = srvSABCPTR.arrSABCPTR_NB + 1: ReDim meSABCPTR(meSABCPTR_NbMax)
ReDim meXSABCPTR(meSABCPTR_NbMax)
meSABCPTR_Nb = 0

For I = 1 To srvSABCPTR.arrSABCPTR_NB
    If srvSABCPTR.arrSABCPTR(I).SCORIG = " " Then
        meSABCPTR_Nb = meSABCPTR_Nb + 1
        meSABCPTR(meSABCPTR_Nb) = srvSABCPTR.arrSABCPTR(I)
        meXSABCPTR(meSABCPTR_Nb) = zSABCPTR
    Else
         meXSABCPTR(meSABCPTR_Nb) = srvSABCPTR.arrSABCPTR(I)
   End If
    
Next I

fgSelect_Display
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
    fgSelect.Col = fgSelect_arrIndex
    meSABCPTR_Index = Val(fgSelect.Text)
    fgSelect.Col = fgSelect_arrIndex - 1
   X = meSABCPTR(meSABCPTR_Index).SCCOMPTE & meSABCPTR(meSABCPTR_Index).SCDEVISE
    Select Case lK
        Case 0: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCSTATUS & X
        Case 1: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCNATGA & X
        Case 2: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCDEVISE & X
        Case 3: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCCOMPTE & X
        Case 4: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCSABID & X
        Case 5: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCTDC & X
        Case 6: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCPCEC & X
        Case 7: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCINTITU & X
        Case fgSelect_arrIndex: fgSelect.Text = Format$(meSABCPTR_Index, "0000000000")
    End Select
Next I
fgSelect.Col = 0: fgSelect.Text = meSABCPTR(meSABCPTR_Index).SCSTATUS


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub


Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

SSTab1.Tab = 0
ReDim meSABCPTR(10)

blnControl = False
fgSelect_FormatString = fgSelect.FormatString

cmdReset

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


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Dim Msg As String
Dim I As Integer

Me.Enabled = False

Msg = Space$(50)
'prtSABCPTR_Open Msg

'For I = 1 To fgSelect.Rows - 1
'    fgSelect.Row = I
'    fgSelect.Col = fgSelect_arrIndex
'    meSABCPTR_Index = Val(fgSelect.Text)
'    fgSelect.Col = fgSelect_arrIndex - 1
'    recSABCPTR = meSABCPTR(meSABCPTR_Index)
'    prtSABCPTR_Line recSABCPTR
'Next I


'prtSABCPTR_Close

Me.Enabled = True

End Sub

Private Sub cmdSelect_Click()
cmdSelect_Control
If lstErr.ListCount = 0 Then fgSelect_Load
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xStatut As String

If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_SortX 0
        Case 1: fgSelect_SortX 1
        Case 2: fgSelect_SortX 2
        Case 3: fgSelect_SortX 3
        Case 4: fgSelect_SortX 4
        Case 5: fgSelect_SortX 5
        Case 6: fgSelect_SortX 6
        Case 7: fgSelect_SortX 7
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        fgSelect_Click_Ok
    
    
        If Button = vbRightButton Then
            Me.PopupMenu mnuOpération, vbPopupMenuLeftButton
        Else
            fraSC_Display
        End If
        
    End If
End If

End Sub
Private Sub txtXXX_GotFocus()

'KeyAscii = convUCase(KeyAscii)

'txt_GotFocus txtXXX

'txt_LostFocus txtXXX
'If blnControl Then cmdControl

'DTPicker_GotFocus txtXXX

'DTPicker_LostFocus txtXXX
'If blnControl Then cmdControl


' Change : txtAmjfin_control
'MouseMoveActiveControl_Set txtXXX

End Sub


Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub

Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

Private Sub mnuOpérationDisplay_Click()
srvSABCPTR_ElpDisplay mSABCPTR
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), SABCPTRAut)    '

blnSetfocus = True
Form_Init


End Sub


Public Sub cmdContext_Return()
If SSTab1.Tab > 0 Then
    SSTab1.Tab = 0
Else
    SendKeys "{TAB}"
    
End If

End Sub


Public Sub fgSelect_Reset()
fgSelect_Sort1 = 1: fgSelect_Sort2 = 2
fgSelect_Sort1_Old = 0
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 12
blnfgSelect_DisplayLine = False

End Sub

Public Sub fraSC_Display()
Dim lColor As Long
libRéférenceInterne = mSABCPTR.SCDEVISE & "   " & Compte_Imp(mSABCPTR.SCCOMPTE)
lColor = lblUsr.ForeColor
fraSC.Enabled = False
SSTab1.Tab = 2
txtSCDEVISE.Enabled = False
txtSCDEVISO.Enabled = False
txtSCCOMPTE.Enabled = False
txtSCSABID.Enabled = True '$jplc 20021120 False
'chkSCUpdate.Enabled = False
cmdSCOk.Enabled = False
txtSCOUVAMJ.Enabled = False
txtSCCLOAMJ.Enabled = False
cboSCSTATUS.Enabled = False
txtSCCPTGEN.Enabled = False

txtSCDEVISE = Trim(mSABCPTR.SCDEVISE)
txtSCCOMPTE = Compte_Display(Trim(mSABCPTR.SCCOMPTE))

txtSCDEVISO = Trim(mSABCPTR.SCDEVISO)

txtSCSABID = CompteSAB_Imp(Trim(mSABCPTR.SCSABID)): lblSCSABID.ForeColor = lColor

txtSCTDC = Trim(mSABCPTR.SCTDC): lblSCTDC.ForeColor = lColor

txtSCPCEC = Trim(mSABCPTR.SCPCEC): lblSCPCEC.ForeColor = lColor

txtSCINTITU = Trim(mSABCPTR.SCINTITU): lblSCINTITU.ForeColor = lColor

Call DTPicker_Set(txtSCOUVAMJ, mSABCPTR.SCOUVAMJ): lblSCOUVAMJ.ForeColor = lColor
If mSABCPTR.SCCLOAMJ = "00000000" Or mSABCPTR.SCCLOAMJ = "        " Then
    txtSCCLOAMJ.Visible = False
    txtSCCLOMOT.Enabled = False
Else
    txtSCCLOMOT.Enabled = True
    txtSCCLOAMJ.Visible = True
    Call DTPicker_Set(txtSCCLOAMJ, mSABCPTR.SCCLOAMJ)
End If

txtSCLORO = Trim(mSABCPTR.SCLORO): lblSCLORO.ForeColor = lColor

chkSCSUCCES = IIf(mSABCPTR.SCSUCCES = "1", "1", "0"): chkSCSUCCES.ForeColor = lColor

txtSCSECUR = Trim(mSABCPTR.SCSECUR): lblSCSECUR.ForeColor = lColor

txtSCSITUAT = Trim(mSABCPTR.SCSITUAT): lblSCSITUAT.ForeColor = lColor

txtSCCLOMOT = Trim(mSABCPTR.SCCLOMOT): lblSCCLOMOT.ForeColor = lColor

txtSCTITID = Trim(mSABCPTR.SCTITID): lblSCTITID.ForeColor = lColor

chkSCTITCPT = IIf(mSABCPTR.SCTITCPT = "1", "1", "0"): chkSCTITCPT.ForeColor = lColor

chkSCTITPRN = IIf(mSABCPTR.SCTITPRN = "1", "1", "0"): chkSCTITPRN.ForeColor = lColor

chkSCTITRSP = IIf(mSABCPTR.SCTITRSP = "1", "1", "0"): chkSCTITRSP.ForeColor = lColor

txtSCRELCOD = Trim(mSABCPTR.SCRELCOD): lblSCRELCOD.ForeColor = lColor

txtSCRELADR = Trim(mSABCPTR.SCRELADR): lblSCRELADR.ForeColor = lColor

chkSCRELGES = IIf(mSABCPTR.SCRELGES = "1", "1", "0"): chkSCRELGES.ForeColor = lColor

txtSCRELNOR = Trim(mSABCPTR.SCRELNOR): lblSCRELNOR.ForeColor = lColor

txtSCALIASCPT = Trim(mSABCPTR.SCALIASCPT): lblSCALIASCPT.ForeColor = lColor

txtSCCPTGEN = Trim(mSABCPTR.SCCPTGEN)

txtSCCREAMJ = dateImp10(mSABCPTR.SCCREAMJ) & "  " & timeImp8(mSABCPTR.SCCREHMS)
txtSCMODAMJ = dateImp10(mSABCPTR.SCMODAMJ) & "  " & timeImp8(mSABCPTR.SCMODHMS)
txtSCUSRAMJ = ""

Call cbo_Scan(mSABCPTR.SCSTATUS, cboSCSTATUS)

If mXSABCPTR.SCORIG = "M" Then fraSC_Display_X

If Trim(mSABCPTR.SCSTATUS) = "" Then
    cmdSCOk.Enabled = SABCPTRAut.Saisir
    fraSC.Enabled = SABCPTRAut.Saisir
End If


End Sub
Public Sub fraSC_Display_X()

If Trim(mXSABCPTR.SCSABID) <> "" Then
    txtSCSABID = Trim(mXSABCPTR.SCSABID)
    lblSCSABID.ForeColor = warnUsrColor
End If

If Trim(mXSABCPTR.SCTDC) <> "" Then
    txtSCTDC = Trim(mXSABCPTR.SCTDC)
    lblSCTDC.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCPCEC) <> "" Then
    txtSCPCEC = Trim(mXSABCPTR.SCPCEC)
    lblSCPCEC.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCINTITU) <> "" Then
    txtSCINTITU = Trim(mXSABCPTR.SCINTITU)
    lblSCINTITU.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCLORO) <> "" Then
    txtSCLORO = Trim(mXSABCPTR.SCLORO)
    lblSCLORO.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCSECUR) <> "" Then
    txtSCSECUR = Trim(mXSABCPTR.SCSECUR)
    lblSCSECUR.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCSITUAT) <> "" Then
    txtSCSITUAT = Trim(mXSABCPTR.SCSITUAT)
    lblSCSITUAT.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCCLOMOT) <> "" Then
    txtSCCLOMOT = Trim(mXSABCPTR.SCCLOMOT)
    lblSCCLOMOT.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCTITID) <> "" Then
    txtSCTITID = Trim(mXSABCPTR.SCTITID)
    lblSCTITID.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCRELCOD) <> "" Then
    txtSCRELCOD = Trim(mXSABCPTR.SCRELCOD)
    lblSCRELCOD.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCRELADR) <> "" Then
    txtSCRELADR = Trim(mXSABCPTR.SCRELADR)
    lblSCRELADR.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCRELNOR) <> "" Then
    txtSCRELNOR = Trim(mXSABCPTR.SCRELNOR)
    lblSCRELNOR.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCALIASCPT) <> "" Then
    txtSCALIASCPT = Trim(mXSABCPTR.SCALIASCPT)
    lblSCALIASCPT.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCSUCCES) <> "" Then
    chkSCSUCCES = IIf(mXSABCPTR.SCSUCCES = "1", "1", "0")
    chkSCSUCCES.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCTITCPT) <> "" Then
    chkSCTITCPT = IIf(mXSABCPTR.SCTITCPT = "1", "1", "0")
    chkSCTITCPT.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCTITPRN) <> "" Then
    chkSCTITPRN = IIf(mXSABCPTR.SCTITPRN = "1", "1", "0")
    chkSCTITPRN.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCTITRSP) <> "" Then
    chkSCTITRSP = IIf(mXSABCPTR.SCTITRSP = "1", "1", "0")
    chkSCTITRSP.ForeColor = warnUsrColor
End If
If Trim(mXSABCPTR.SCRELGES) <> "" Then
    chkSCRELGES = IIf(mXSABCPTR.SCRELGES = "1", "1", "0")
    chkSCRELGES.ForeColor = warnUsrColor
End If
txtSCUSRAMJ = dateImp10(mXSABCPTR.SCMODAMJ) & "  " & timeImp8(mXSABCPTR.SCMODHMS) & "  " & mXSABCPTR.SCUSRNOM

End Sub


Public Sub fgSelect_Click_Ok()
fgSelect.Col = fgSelect_arrIndex
meSABCPTR_Index = Val(fgSelect.Text)
mSABCPTR = meSABCPTR(meSABCPTR_Index)
mXSABCPTR = meXSABCPTR(meSABCPTR_Index)

Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
mnuOpérationDisplay = SABCPTRAut.Consulter

End Sub

Public Sub cmdUpdate_Control()
If xSABCPTR.SCORIG = " " Then
    xSABCPTR.SCORIG = "M"
    xSABCPTR.SCCOMPTE = mSABCPTR.SCCOMPTE
    xSABCPTR.SCDEVISE = mSABCPTR.SCDEVISE
    xSABCPTR.Method = constAddNew
Else
    xSABCPTR.Method = constUpdate
End If
xSABCPTR.SCSABID = Trim(txtSCSABID)
If xSABCPTR.SCSABID = mSABCPTR.SCSABID Then xSABCPTR.SCSABID = ""
xSABCPTR.SCTDC = Trim(txtSCTDC)
If xSABCPTR.SCTDC = mSABCPTR.SCTDC Then xSABCPTR.SCTDC = ""
xSABCPTR.SCPCEC = Trim(txtSCPCEC)
If xSABCPTR.SCPCEC = mSABCPTR.SCPCEC Then xSABCPTR.SCPCEC = ""
xSABCPTR.SCINTITU = Trim(txtSCINTITU)
If xSABCPTR.SCINTITU = mSABCPTR.SCINTITU Then xSABCPTR.SCINTITU = ""
xSABCPTR.SCDEVISO = Trim(txtSCDEVISO)
If xSABCPTR.SCDEVISO = mSABCPTR.SCDEVISO Then xSABCPTR.SCDEVISO = ""
xSABCPTR.SCLORO = Trim(txtSCLORO)
If xSABCPTR.SCLORO = mSABCPTR.SCLORO Then xSABCPTR.SCLORO = ""
xSABCPTR.SCSUCCES = IIf(chkSCSUCCES = "1", "1", "0")
If xSABCPTR.SCSUCCES = mSABCPTR.SCSUCCES Then xSABCPTR.SCSUCCES = ""
lX = CLng(Val(Trim(txtSCSECUR)))
xSABCPTR.SCSECUR = Format(lX, "00")
If xSABCPTR.SCSECUR = mSABCPTR.SCSECUR Then xSABCPTR.SCSECUR = ""
xSABCPTR.SCSITUAT = Trim(txtSCSITUAT)
If xSABCPTR.SCSITUAT = mSABCPTR.SCSITUAT Then xSABCPTR.SCSITUAT = ""
xSABCPTR.SCCLOMOT = Trim(txtSCCLOMOT)
If xSABCPTR.SCCLOMOT = mSABCPTR.SCCLOMOT Then xSABCPTR.SCCLOMOT = ""
xSABCPTR.SCTITID = Trim(txtSCTITID)
If xSABCPTR.SCTITID = mSABCPTR.SCTITID Then xSABCPTR.SCTITID = ""
xSABCPTR.SCTITCPT = IIf(chkSCTITCPT = "1", "1", "0")
If xSABCPTR.SCTITCPT = mSABCPTR.SCTITCPT Then xSABCPTR.SCTITCPT = ""
xSABCPTR.SCTITPRN = IIf(chkSCTITPRN = "1", "1", "0")
If xSABCPTR.SCTITPRN = mSABCPTR.SCTITPRN Then xSABCPTR.SCTITPRN = ""
xSABCPTR.SCTITRSP = IIf(chkSCTITRSP = "1", "1", "0")
If xSABCPTR.SCTITRSP = mSABCPTR.SCTITRSP Then xSABCPTR.SCTITRSP = ""
xSABCPTR.SCRELCOD = Trim(txtSCRELCOD)
If xSABCPTR.SCRELCOD = mSABCPTR.SCRELCOD Then xSABCPTR.SCRELCOD = ""
xSABCPTR.SCRELADR = Trim(txtSCRELADR)
If xSABCPTR.SCRELADR = mSABCPTR.SCRELADR Then xSABCPTR.SCRELADR = ""
xSABCPTR.SCRELGES = IIf(chkSCRELGES = "1", "1", "0")
If xSABCPTR.SCRELGES = mSABCPTR.SCRELGES Then xSABCPTR.SCRELGES = ""
lX = CLng(Val(Trim(txtSCRELNOR)))
If lX = 0 Then
    xSABCPTR.SCRELNOR = ""
Else
    xSABCPTR.SCRELNOR = Format(lX, "000000")
End If
If xSABCPTR.SCRELNOR = mSABCPTR.SCRELNOR Then xSABCPTR.SCRELNOR = ""
xSABCPTR.SCALIASCPT = Trim(txtSCALIASCPT)
If xSABCPTR.SCALIASCPT = mSABCPTR.SCALIASCPT Then xSABCPTR.SCALIASCPT = ""
xSABCPTR.SCSTATUS = Trim(cboSCSTATUS)

End Sub
